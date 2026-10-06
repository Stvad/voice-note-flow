#!/usr/bin/env python3
"""Tests for the phone-side video audio extractor.

Run: python3 -m unittest discover -s android -p 'test_*.py'

The planning tests are pure. The end-to-end tests drive real ffmpeg/ffprobe on
tiny generated clips, and are skipped when those are not on PATH.
"""
import json
import logging
import os
import shutil
import subprocess
import tempfile
import unittest
from datetime import datetime
from pathlib import Path

from extract_video_audio import (
    Config,
    State,
    Video,
    apply_since,
    is_video_name,
    parse_since,
    plan_run,
    prune_state,
    run,
)

# The failure-path tests log on purpose; keep that out of the test output.
logging.getLogger("video-audio").addHandler(logging.NullHandler())
logging.getLogger("video-audio").propagate = False

NOW = 1_790_000_000.0  # fixed clock so tests are deterministic
MIN = 60.0
OPTS = dict(settle_seconds=120, max_attempts=3, initial_limit=0)


def vid(name, age_minutes, size=1000):
    return Video(name=name, mtime=NOW - age_minutes * MIN, size=size)


def names(videos):
    return [v.name for v in videos]


class TestIsVideoName(unittest.TestCase):
    def test_camera_videos_count(self):
        self.assertTrue(is_video_name("20261004_143012.mp4"))
        self.assertTrue(is_video_name("CLIP.MOV"))

    def test_in_progress_and_trashed_files_do_not(self):
        # Android names a file that is still being written, or sitting in the
        # trash, with a leading dot. Neither is a finished video.
        self.assertFalse(is_video_name(".pending-1790000000-20261004_143012.mp4"))
        self.assertFalse(is_video_name(".trashed-1790000000-20261004_143012.mp4"))

    def test_photos_do_not(self):
        self.assertFalse(is_video_name("20261004_143012.jpg"))
        self.assertFalse(is_video_name("no_extension"))


class TestFirstRun(unittest.TestCase):
    def test_does_not_replay_the_camera_roll(self):
        plan = plan_run([vid("old", 600), vid("older", 6000)], State(), NOW, **OPTS)

        self.assertEqual(plan.to_process, [])
        self.assertEqual(plan.floor, NOW)

    def test_a_video_recorded_after_the_first_run_is_picked_up(self):
        first = plan_run([vid("old", 600)], State(), NOW, **OPTS)
        state = State(floor=first.floor)

        later = NOW + 30 * MIN
        new = Video("new", mtime=NOW + 10 * MIN, size=1000)
        plan = plan_run([vid("old", 600), new], state, later, **OPTS)

        self.assertEqual(names(plan.to_process), ["new"])

    def test_initial_limit_takes_the_newest_finished_videos(self):
        videos = [vid("a", 300), vid("b", 200), vid("c", 100), vid("recording", 0)]

        plan = plan_run(videos, State(), NOW, **dict(OPTS, initial_limit=2))

        self.assertEqual(names(plan.to_process), ["b", "c"], "oldest-first")
        self.assertLess(plan.floor, videos[1].mtime, "floor keeps 'b' in scope")
        self.assertGreaterEqual(plan.floor, videos[0].mtime, "floor drops 'a'")
        self.assertGreater(videos[3].mtime, plan.floor,
                           "the clip still recording is picked up once it settles")


class TestSteadyState(unittest.TestCase):
    def state(self, **kw):
        return State(floor=NOW - 7 * 24 * 60 * MIN, **kw)

    def test_oldest_first(self):
        plan = plan_run([vid("new", 10), vid("old", 50), vid("mid", 30)],
                        self.state(), NOW, **OPTS)

        self.assertEqual(names(plan.to_process), ["old", "mid", "new"])

    def test_done_videos_are_skipped(self):
        state = self.state(done={"a": NOW - 50 * MIN})

        plan = plan_run([vid("a", 50), vid("b", 40)], state, NOW, **OPTS)

        self.assertEqual(names(plan.to_process), ["b"])

    def test_videos_still_being_written_wait_until_they_settle(self):
        # The camera keeps touching the file while it records, so a recent
        # mtime means it may not be finished yet.
        plan = plan_run([vid("fresh", 1)], self.state(), NOW, **OPTS)

        self.assertEqual(plan.to_process, [])
        self.assertEqual(plan.deferred, 1)

    def test_videos_at_or_below_the_floor_are_out_of_scope(self):
        state = State(floor=NOW - 60 * MIN)

        plan = plan_run([vid("before", 60), vid("after", 59)], state, NOW, **OPTS)

        self.assertEqual(names(plan.to_process), ["after"])

    def test_exhausted_videos_are_handed_back_not_retried(self):
        state = self.state(attempts={"stuck": 3, "flaky": 2})

        plan = plan_run([vid("stuck", 50), vid("flaky", 40)], state, NOW, **OPTS)

        self.assertEqual(names(plan.to_process), ["flaky"])
        self.assertEqual(names(plan.exhausted), ["stuck"])

    def test_floor_never_moves_once_set(self):
        state = self.state()

        plan = plan_run([vid("a", 10)], state, NOW, **OPTS)

        self.assertEqual(plan.floor, state.floor)


class TestSince(unittest.TestCase):
    SINCE = NOW - 3 * 24 * 60 * MIN

    def test_first_run_starts_from_the_date_instead_of_now(self):
        state = apply_since(State(), self.SINCE)

        plan = plan_run([vid("before", 3 * 24 * 60 + 1), vid("after", 3 * 24 * 60 - 1)],
                        state, NOW, **OPTS)

        self.assertEqual(names(plan.to_process), ["after"])

    def test_lowers_an_existing_floor_without_redoing_handled_videos(self):
        state = State(floor=NOW - 60 * MIN, done={"handled": NOW - 30 * MIN})

        lowered = apply_since(state, self.SINCE)
        plan = plan_run([vid("missed", 600), vid("handled", 30)], lowered, NOW, **OPTS)

        self.assertEqual(lowered.floor, self.SINCE)
        self.assertEqual(names(plan.to_process), ["missed"])

    def test_never_raises_the_floor(self):
        # Raising it would silently skip videos that are still pending.
        state = State(floor=self.SINCE - 60 * MIN)

        self.assertEqual(apply_since(state, self.SINCE).floor, state.floor)

    def test_no_date_changes_nothing(self):
        state = State(floor=NOW)

        self.assertEqual(apply_since(state, None), state)
        self.assertIsNone(apply_since(State(), None).floor)

    def test_a_bare_date_means_local_midnight(self):
        self.assertEqual(parse_since("2026-10-03"), datetime(2026, 10, 3).timestamp())
        self.assertEqual(parse_since("2026-10-03T18:30"),
                         datetime(2026, 10, 3, 18, 30).timestamp())


class TestPrune(unittest.TestCase):
    def test_forgets_videos_that_were_deleted_from_the_phone(self):
        state = State(floor=0.0, done={"kept": 1.0, "deleted": 2.0},
                      attempts={"kept-retry": 1, "deleted-retry": 2})

        pruned = prune_state(state, {"kept", "kept-retry"})

        self.assertEqual(pruned.done, {"kept": 1.0})
        self.assertEqual(pruned.attempts, {"kept-retry": 1})
        self.assertEqual(pruned.floor, 0.0)


FFMPEG = shutil.which("ffmpeg") and shutil.which("ffprobe")


def make_clip(path, audio=True):
    """A one-second clip, made with the encoders every ffmpeg build ships."""
    cmd = ["ffmpeg", "-nostdin", "-loglevel", "error", "-y",
           "-f", "lavfi", "-i", "testsrc=duration=1:size=64x64:rate=5"]
    if audio:
        cmd += ["-f", "lavfi", "-i", "sine=frequency=440:duration=1:sample_rate=48000",
                "-ac", "2", "-c:a", "aac"]
    cmd += ["-c:v", "mpeg4", "-shortest", str(path)]
    subprocess.run(cmd, check=True)


@unittest.skipUnless(FFMPEG, "needs ffmpeg and ffprobe on PATH")
class EndToEndCase(unittest.TestCase):
    def setUp(self):
        self.tmp = Path(tempfile.mkdtemp())
        self.addCleanup(shutil.rmtree, self.tmp)
        self.cfg = Config(
            camera_dir=self.tmp / "DCIM" / "Camera",
            outbox=self.tmp / "VideoAudio",
            staging=self.tmp / ".video-audio-staging",
            state_dir=self.tmp / "state",
            settle_seconds=120,
            initial_limit=0,
            max_attempts=3,
        )
        self.cfg.camera_dir.mkdir(parents=True)
        self.first_run()

    def first_run(self):
        # Start the ledger here, so the clips below count as new.
        run(self.cfg, NOW)

    def record(self, name, minutes_after_setup, audio=True):
        path = self.cfg.camera_dir / name
        make_clip(path, audio=audio)
        mtime = NOW + minutes_after_setup * MIN
        os.utime(path, (mtime, mtime))
        return path

    def state(self):
        return json.loads((self.cfg.state_dir / "state.json").read_text())

    def outbox(self):
        return sorted(p.name for p in self.cfg.outbox.iterdir())


class TestEndToEnd(EndToEndCase):
    def test_new_video_becomes_mono_m4a_in_the_outbox(self):
        self.record("20261004_143012.mp4", 5)

        run(self.cfg, NOW + 10 * MIN)

        out = self.cfg.outbox / "20261004_143012.m4a"
        self.assertTrue(out.exists())
        probe = subprocess.run(
            ["ffprobe", "-v", "error", "-show_entries",
             "stream=codec_type,codec_name,channels,sample_rate", "-of", "json", str(out)],
            capture_output=True, text=True, check=True)
        streams = json.loads(probe.stdout)["streams"]
        self.assertEqual(
            [(s["codec_type"], s["codec_name"], s["channels"], s["sample_rate"]) for s in streams],
            [("audio", "aac", 1, "16000")], "audio only, mono, speech sample rate")
        self.assertEqual(list(self.cfg.staging.glob("*.m4a")), [], "nothing left in staging")
        self.assertIn("20261004_143012.mp4", self.state()["done"])

    def test_audio_keeps_the_videos_recording_time(self):
        # Autosync carries the mtime over to Drive, where Code.js reads it as the
        # note's timestamp. Without this a backfilled clip is dated today.
        self.record("a.mp4", 5)

        run(self.cfg, NOW + 10 * MIN)

        self.assertEqual((self.cfg.outbox / "a.m4a").stat().st_mtime, NOW + 5 * MIN)

    def test_audio_over_the_upload_limit_is_kept_but_flagged(self):
        self.cfg.upload_limit_mb = 0.001
        self.record("long.mp4", 5)

        with self.assertLogs("video-audio", level="WARNING") as logs:
            run(self.cfg, NOW + 10 * MIN)

        self.assertIn("long.m4a", self.outbox())
        self.assertIn("upload limit", "\n".join(logs.output))

    def test_outbox_is_hidden_from_media_apps(self):
        self.assertTrue((self.cfg.outbox / ".nomedia").exists())

    def test_a_processed_video_is_not_extracted_again(self):
        self.record("a.mp4", 5)
        run(self.cfg, NOW + 10 * MIN)
        (self.cfg.outbox / "a.m4a").unlink()  # e.g. Autosync's "upload then delete"

        run(self.cfg, NOW + 20 * MIN)

        self.assertNotIn("a.m4a", self.outbox())

    def test_video_without_audio_is_done_without_output(self):
        # Hyperlapse and similar modes record no sound at all.
        self.record("hyperlapse.mp4", 5, audio=False)

        run(self.cfg, NOW + 10 * MIN)

        self.assertNotIn("hyperlapse.m4a", self.outbox())
        self.assertIn("hyperlapse.mp4", self.state()["done"])

    def test_unreadable_video_is_retried_then_given_up(self):
        # Without its index (written last) an mp4 can't be read.
        path = self.record("broken.mp4", 5)
        data = path.read_bytes()
        path.write_bytes(data[: len(data) // 2])
        mtime = NOW + 5 * MIN
        os.utime(path, (mtime, mtime))

        run(self.cfg, NOW + 10 * MIN)
        self.assertEqual(self.state()["attempts"], {"broken.mp4": 1})
        self.assertNotIn("broken.mp4", self.state()["done"])

        run(self.cfg, NOW + 20 * MIN)
        run(self.cfg, NOW + 30 * MIN)

        self.assertIn("broken.mp4", self.state()["done"], "gave up after 3 attempts")
        self.assertEqual(self.state()["attempts"], {})
        self.assertNotIn("broken.m4a", self.outbox())

    def test_one_bad_video_does_not_block_the_next(self):
        bad = self.record("bad.mp4", 5)
        bad.write_bytes(b"not a video")
        os.utime(bad, (NOW + 5 * MIN, NOW + 5 * MIN))
        self.record("good.mp4", 6)

        run(self.cfg, NOW + 10 * MIN)

        self.assertIn("good.m4a", self.outbox())


class TestBackfill(EndToEndCase):
    def first_run(self):
        pass  # the test does it, with a start date

    def test_first_run_with_a_date_picks_up_videos_since_then(self):
        self.record("friday.mp4", -3 * 24 * 60)
        self.record("saturday.mp4", -2 * 24 * 60)
        self.cfg.since = NOW - (2 * 24 * 60 + 60) * MIN

        run(self.cfg, NOW)

        self.assertEqual([n for n in self.outbox() if n.endswith(".m4a")], ["saturday.m4a"])


if __name__ == "__main__":
    unittest.main()
