#!/data/data/com.termux/files/usr/bin/python3
"""Extract the audio from new camera videos for the voice-note pipeline.

Runs on the phone under Termux, every 15 minutes via termux-job-scheduler (see
android/SETUP.md). Each new video in the camera folder becomes a small mono
.m4a in the outbox. Autosync uploads the outbox to the Drive folder listed in
VIDEO_FOLDER_IDS, and Code.js takes it from there.

Audio rather than video is the point: Apps Script can't send more than 50MB in
one request, and a phone video passes that within a minute or so. Mono 64 kbps
AAC is about 0.5 MB a minute, which leaves room for roughly 100 minutes.

What counts as new mirrors Code.js. A ledger of handled file names decides;
the floor only marks where the ledger starts, so the first run does not
replay the whole camera roll.
"""
import argparse
import fcntl
import json
import logging
import logging.handlers
import os
import shutil
import subprocess
import sys
import time
from dataclasses import asdict, dataclass, field
from pathlib import Path

SHARED = Path("/storage/emulated/0")
VIDEO_EXTENSIONS = {".mp4", ".mov", ".3gp", ".mkv", ".webm"}

log = logging.getLogger("video-audio")


@dataclass(frozen=True)
class Video:
    name: str
    mtime: float
    size: int


@dataclass
class State:
    floor: float | None = None  # None means we have never run
    done: dict[str, float] = field(default_factory=dict)  # name -> video mtime
    # Attempts that have not succeeded (yet). Counted as each one starts, so a
    # run killed mid-file still uses one up.
    attempts: dict[str, int] = field(default_factory=dict)


@dataclass
class Plan:
    to_process: list[Video]  # oldest first
    exhausted: list[Video]  # out of attempts: give up on these
    floor: float
    deferred: int  # new, but possibly still being recorded


@dataclass
class Config:
    camera_dir: Path
    outbox: Path
    staging: Path  # same filesystem as the outbox, so the final move is atomic
    state_dir: Path
    settle_seconds: float
    initial_limit: int
    max_attempts: int


# ============================================================
# Selection logic: pure functions, unit-tested
# ============================================================


def is_video_name(name: str) -> bool:
    # Dot-files are Android's in-progress (.pending-*) and trashed (.trashed-*)
    # media, never a finished video.
    if name.startswith("."):
        return False
    return Path(name).suffix.lower() in VIDEO_EXTENSIONS


def _oldest_first(v: Video):
    return (v.mtime, v.name)


def plan_run(videos, state, now, *, settle_seconds, max_attempts, initial_limit) -> Plan:
    # The camera keeps writing to the file while it records, so a recent mtime
    # can mean it is not finished. ffprobe catches the rest (see process_one).
    def settled(v):
        return now - v.mtime >= settle_seconds

    if state.floor is None:
        newest = sorted((v for v in videos if settled(v)), key=_oldest_first, reverse=True)
        keep = newest[:initial_limit]
        # Everything at or below the floor is the back catalogue. A clip still
        # recording right now will move above it as soon as it is written to.
        floor = keep[-1].mtime - 0.001 if keep else now
        return Plan(sorted(keep, key=_oldest_first), [], floor, 0)

    to_process, exhausted, deferred = [], [], 0
    for v in videos:
        if v.mtime <= state.floor or v.name in state.done:
            continue
        if not settled(v):
            deferred += 1
        elif state.attempts.get(v.name, 0) >= max_attempts:
            exhausted.append(v)
        else:
            to_process.append(v)
    return Plan(sorted(to_process, key=_oldest_first), sorted(exhausted, key=_oldest_first),
                state.floor, deferred)


def prune_state(state: State, present: set) -> State:
    """Forget videos that are no longer on the phone. Their names are the only
    thing that could ever bring them back, and the floor already keeps
    anything older than the ledger out of scope."""
    return State(
        floor=state.floor,
        done={n: m for n, m in state.done.items() if n in present},
        attempts={n: c for n, c in state.attempts.items() if n in present},
    )


# ============================================================
# I/O
# ============================================================


class ProbeError(Exception):
    pass


def scan(camera_dir: Path) -> list[Video]:
    videos = []
    with os.scandir(camera_dir) as entries:
        for entry in entries:
            if entry.is_file() and is_video_name(entry.name):
                st = entry.stat()
                videos.append(Video(entry.name, st.st_mtime, st.st_size))
    return videos


def audio_stream_count(path: Path) -> int:
    result = subprocess.run(
        ["ffprobe", "-v", "error", "-select_streams", "a",
         "-show_entries", "stream=index", "-of", "csv=p=0", str(path)],
        capture_output=True, text=True, timeout=300,
    )
    if result.returncode != 0:
        # Typically "moov atom not found": the mp4 index is written last, so
        # the file was cut short or is still being finished.
        raise ProbeError(result.stderr.strip() or f"ffprobe exited {result.returncode}")
    return len(result.stdout.split())


def extract_audio(src: Path, dst: Path) -> None:
    # Deepgram gains nothing from stereo or a higher bitrate.
    subprocess.run(
        ["ffmpeg", "-nostdin", "-hide_banner", "-loglevel", "error", "-y",
         "-i", str(src), "-map", "0:a:0", "-ac", "1", "-c:a", "aac", "-b:a", "64k",
         "-movflags", "+faststart", str(dst)],
        check=True, capture_output=True, text=True, timeout=1800,
    )


def process_one(video: Video, cfg: Config) -> str | None:
    """Returns the outbox file name, or None for a video with no sound."""
    src = cfg.camera_dir / video.name
    if audio_stream_count(src) == 0:
        return None
    out_name = Path(video.name).stem + ".m4a"
    tmp = cfg.staging / out_name
    extract_audio(src, tmp)
    # Autosync must never see a half-written file, so it only appears in the
    # outbox once complete.
    os.replace(tmp, cfg.outbox / out_name)
    return out_name


def load_state(path: Path) -> State:
    if not path.exists():
        return State()
    return State(**json.loads(path.read_text()))


def save_state(path: Path, state: State) -> None:
    tmp = path.with_suffix(".tmp")
    tmp.write_text(json.dumps(asdict(state), indent=1, sort_keys=True))
    os.replace(tmp, path)


def notify(title: str, content: str, notification_id: str | None = None) -> None:
    """Best effort phone notification (Termux:API). Nothing else would tell you.
    Reusing an id replaces the earlier notification instead of stacking."""
    if not shutil.which("termux-notification"):
        return
    cmd = ["termux-notification", "--title", title, "--content", content]
    if notification_id:
        cmd += ["--id", notification_id]
    try:
        subprocess.run(cmd, capture_output=True, timeout=30)
    except (OSError, subprocess.SubprocessError) as e:
        log.warning("could not post a notification: %s", e)


FILE_ERRORS = (ProbeError, subprocess.CalledProcessError, subprocess.TimeoutExpired)


def describe(err: Exception) -> str:
    if isinstance(err, subprocess.CalledProcessError):
        return (err.stderr or "").strip() or f"ffmpeg exited {err.returncode}"
    return str(err)


def run(cfg: Config, now: float) -> None:
    for tool in ("ffmpeg", "ffprobe"):
        if not shutil.which(tool):
            # Without this check every video would fail and be given up on.
            raise RuntimeError(f"{tool} not found; run: pkg install ffmpeg")
    for d in (cfg.outbox, cfg.staging, cfg.state_dir):
        d.mkdir(parents=True, exist_ok=True)
    # Keeps audio players and the Voice Recorder app from listing these files.
    # Created once rather than touched: a fresh mtime would get it re-uploaded.
    for d in (cfg.outbox, cfg.staging):
        marker = d / ".nomedia"
        if not marker.exists():
            marker.touch()

    state_path = cfg.state_dir / "state.json"
    state = load_state(state_path)
    videos = scan(cfg.camera_dir)
    plan = plan_run(videos, state, now, settle_seconds=cfg.settle_seconds,
                    max_attempts=cfg.max_attempts, initial_limit=cfg.initial_limit)
    if state.floor is None:
        log.info("first run: %d video(s) in %s are treated as already handled; "
                 "processing the %d newest", len(videos), cfg.camera_dir, len(plan.to_process))
    state.floor = plan.floor
    save_state(state_path, state)
    log.info("plan: %d to process, %d still settling, %d out of attempts",
             len(plan.to_process), plan.deferred, len(plan.exhausted))

    def give_up(video: Video, reason: str) -> None:
        state.done[video.name] = video.mtime
        state.attempts.pop(video.name, None)
        save_state(state_path, state)
        log.error("giving up on %s: %s", video.name, reason)
        notify("Video audio: gave up on " + video.name, reason[:200])

    for video in plan.exhausted:
        give_up(video, "a run was cut off while processing it")

    for video in plan.to_process:
        attempt = state.attempts.get(video.name, 0) + 1
        state.attempts[video.name] = attempt
        save_state(state_path, state)
        started = time.monotonic()
        try:
            out = process_one(video, cfg)
        except FILE_ERRORS as e:
            log.warning("%s failed (attempt %d/%d): %s",
                        video.name, attempt, cfg.max_attempts, describe(e))
            if attempt >= cfg.max_attempts:
                give_up(video, describe(e))
            continue
        state.done[video.name] = video.mtime
        state.attempts.pop(video.name, None)
        save_state(state_path, state)
        if out is None:
            log.info("%s has no audio track; skipped", video.name)
        else:
            log.info("%s -> %s (%.0f MB video, %.1fs)", video.name, out,
                     video.size / 1e6, time.monotonic() - started)

    save_state(state_path, prune_state(state, {v.name for v in videos}))


# ============================================================
# Entry point
# ============================================================


def parse_args(argv) -> tuple[Config, bool]:
    p = argparse.ArgumentParser(description=__doc__.split("\n\n")[0])
    p.add_argument("--camera-dir", type=Path, default=SHARED / "DCIM" / "Camera")
    p.add_argument("--outbox", type=Path, default=SHARED / "VideoAudio",
                   help="the folder Autosync uploads")
    p.add_argument("--staging", type=Path, default=SHARED / ".video-audio-staging")
    p.add_argument("--state-dir", type=Path,
                   default=Path.home() / ".local" / "state" / "video-audio")
    p.add_argument("--settle-seconds", type=float, default=120)
    p.add_argument("--initial-limit", type=int, default=0,
                   help="on the first run, also process this many of the newest videos")
    p.add_argument("--max-attempts", type=int, default=3)
    p.add_argument("--dry-run", action="store_true",
                   help="print what would be processed and change nothing")
    a = p.parse_args(argv)
    cfg = Config(a.camera_dir, a.outbox, a.staging, a.state_dir,
                 a.settle_seconds, a.initial_limit, a.max_attempts)
    return cfg, a.dry_run


def setup_logging(state_dir: Path) -> None:
    state_dir.mkdir(parents=True, exist_ok=True)
    fmt = logging.Formatter("%(asctime)s %(levelname)s %(message)s")
    to_file = logging.handlers.RotatingFileHandler(
        state_dir / "log.txt", maxBytes=512 * 1024, backupCount=1)
    to_stderr = logging.StreamHandler()
    for h in (to_file, to_stderr):
        h.setFormatter(fmt)
        log.addHandler(h)
    log.setLevel(logging.INFO)


def dry_run(cfg: Config, now: float) -> None:
    state = load_state(cfg.state_dir / "state.json")
    plan = plan_run(scan(cfg.camera_dir), state, now, settle_seconds=cfg.settle_seconds,
                    max_attempts=cfg.max_attempts, initial_limit=cfg.initial_limit)
    if state.floor is None:
        print("First run: everything in the camera folder now counts as already handled.")
    for label, videos in (("would process", plan.to_process), ("would give up", plan.exhausted)):
        for v in videos:
            print(f"{label}: {v.name}  ({v.size / 1e6:.0f} MB)")
    print(f"{len(plan.to_process)} to process, {plan.deferred} still settling, "
          f"{len(state.done)} in the ledger")


def main(argv=None) -> int:
    cfg, dry = parse_args(sys.argv[1:] if argv is None else argv)
    if dry:
        dry_run(cfg, time.time())
        return 0
    setup_logging(cfg.state_dir)
    # termux-job-scheduler can start a run while a long one is still going.
    with open(cfg.state_dir / "lock", "w") as lock:
        try:
            fcntl.flock(lock, fcntl.LOCK_EX | fcntl.LOCK_NB)
        except BlockingIOError:
            log.info("previous run still going; skipping")
            return 0
        try:
            run(cfg, time.time())
        except Exception as e:
            log.exception("run failed")
            # Fixed id: a persistent problem (say, a revoked storage permission)
            # shows one notification, not a new one every 15 minutes.
            notify("Video audio: run failed", f"{type(e).__name__}: {e}"[:200],
                   notification_id="video-audio-run-failed")
            return 1
    return 0


if __name__ == "__main__":
    sys.exit(main())
