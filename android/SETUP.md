# Video notes: phone setup (Galaxy S26 Ultra)

```
DCIM/Camera/*.mp4
  └─ Termux job, every 15 min: extract_video_audio.py → 16 kHz mono .m4a
       └─ <synced voice-notes folder>/VideoAudio ──Autosync──> Drive …/VideoAudio
            └─ Code.js (VIDEO_FOLDER_IDS) → Deepgram → Claude → Matrix
```

Google Photos backup carries on unchanged; this reads the same camera folder
alongside it. Allow about 30 minutes. Do the steps in order, since later ones
depend on earlier ones.

This guide assumes the **free version of Autosync**, with one folder pair
syncing a parent folder that holds your voice-note subfolders. The video audio
goes into a new `VideoAudio` subfolder there, so the existing pair uploads it
without any new pair.

How it behaves:
- **Which videos:** videos recorded after the start date you give on the first
  run (step 5). Everything older is left alone.
- **Silent clips:** clips with no speech post nothing to Matrix. Clips with
  speech arrive tagged `source:: [[video note]]`.
- **Note dates:** each note is dated by when the video was recorded, even when
  it's processed days later.
- **Delay:** expect roughly 15–30 minutes from recording to the Matrix message.
- **Length limit:** the audio is about 0.24 MB a minute, so Autosync's 20 MB
  free-tier cap is reached at about 80 minutes. A longer video's audio stays
  on the phone, and you get a "too big to upload" notification.

## 1. Install Termux and Termux:API from F-Droid

Both apps must come from the **same source**: they share a signing key. The
Google Play build of Termux is still experimental and works differently, so if
you have Termux from Play or GitHub, uninstall every Termux app first.

1. Install F-Droid from https://f-droid.org, or download both APKs directly
   from their pages:
   - https://f-droid.org/packages/com.termux/ (0.118.3 or newer)
   - https://f-droid.org/packages/com.termux.api/ (**0.52.0 or newer**; 0.51
     broke the job scheduler)
2. If Android blocks the install, allow "Install unknown apps" for the browser
   or F-Droid. If it's still blocked, turn off
   *Settings → Security and privacy → Auto Blocker* for the install. You can
   turn it back on afterwards.
3. Open **Termux:API** once from the app drawer, then open **Termux**.

## 2. Install packages

In Termux:

```sh
pkg update && pkg upgrade -y
pkg install -y git python ffmpeg termux-api
```

Check them:

```sh
ffprobe -version | head -1
termux-battery-status
```

`termux-battery-status` should print JSON within a second or two. If it hangs,
press Ctrl+C. Then force-stop Termux:API (*Settings → Apps → Termux:API →
Force stop*), open it from the app drawer and try again. This is a known
Samsung quirk.

## 3. Give Termux storage access

```sh
termux-setup-storage
```

Allow the permission prompt. Then check:

```sh
ls /storage/emulated/0/DCIM/Camera | tail -3
```

You should see your latest photos and videos. If you get *Permission denied*
even though the permission is granted (a known Samsung bug), go to
*Settings → Apps → Termux → Permissions → Files and media*. Switch it to
Don't allow, then back to Allow.

## 4. Get the script and create the job wrapper

Find the **local** path of the folder your Autosync pair syncs: open the pair
in Autosync and look at its local folder. For example, a folder called
`Voice notes` in internal storage is `/storage/emulated/0/Voice notes`. Put
that path in the first line below, keeping `/VideoAudio` on the end, and paste
the whole block into Termux:

```sh
OUTBOX="/storage/emulated/0/YOUR SYNCED FOLDER/VideoAudio"

git clone https://github.com/Stvad/voice-note-flow.git ~/voice-note-flow
cat > ~/video-audio-job <<EOF
#!/data/data/com.termux/files/usr/bin/sh
exec python3 "\$HOME/voice-note-flow/android/extract_video_audio.py" --outbox "$OUTBOX" "\$@"
EOF
chmod +x ~/video-audio-job
cat ~/video-audio-job
```

The last command prints the wrapper. Check that the `--outbox` path is right.
The scheduled job runs this wrapper because the scheduler can't pass arguments
to a script, so any option you want goes in here.

To update the script later: `git -C ~/voice-note-flow pull`

## 5. First run, with backfill from Saturday

`--since` sets how far back the first run goes, here midnight at the start of
Saturday 3 October. Preview first, then run it for real:

```sh
~/video-audio-job --since 2026-10-03 --dry-run
~/video-audio-job --since 2026-10-03
```

The dry run lists every video it would process. The real run extracts them
all straight away, oldest first, and logs a line per video. It also creates
the `VideoAudio` folder.

From then on, plain `~/video-audio-job` carries on from where this left off.
`--since` is only for moving the start date back, and you can use it again
later. Videos already processed are never processed twice.

> **Backfilling further back than a week?** Code.js only looks at files
> modified in the last 7 days. Temporarily set the `LOOKBACK_DAYS` script
> property (step 7) to cover the backfill, and remove it again once the notes
> have arrived.

## 6. Autosync: upload the new subfolder

Your existing folder pair already covers `VideoAudio`, so no new pair is
needed. Do two things:

- **Exclude dot-files:** add the exclude pattern `.*` to the pair, if it isn't
  there already. This keeps the `.nomedia` marker off Drive.
- **Sync:** tap sync. A `VideoAudio` folder should appear in Drive inside your
  voice-notes folder, holding the backfilled `.m4a` files.

## 7. Point Apps Script at the Drive folder

1. Open that `VideoAudio` folder in Drive in a browser. Its ID is the last part
   of the URL: `drive.google.com/drive/folders/<ID>`.
2. In the Apps Script editor, go to *Project Settings* (the gear icon), then
   *Script Properties*. Add `VIDEO_FOLDER_IDS` and set it to that ID.
   - **Don't add it to `FOLDER_IDS`.** Files there would be handled as voice
     notes.

Code.js is already deployed. Within a minute or two the backfilled notes start
arriving in Matrix, oldest first, a few per minute.

## 8. Battery settings

Samsung kills background work aggressively, and the job never fires reliably
unless you change these. Do this for each of **Termux**, **Termux:API** and
**Autosync**:

- *Settings → Apps → [app] → Battery →* **Unrestricted**
- In Settings, search **Background usage limits**. Make sure the app isn't
  under *Sleeping apps* or *Deep sleeping apps*. Add it to *Never sleeping
  apps*, which is called *Never auto sleeping apps* on some One UI versions.

## 9. End-to-end test

1. Record a 15-second video of yourself saying something.
2. Wait **2 minutes**. The script leaves videos alone until they've gone 2
   minutes without changing, in case they're still recording.
3. Run it by hand and look for `<name>.mp4 -> <name>.m4a`:
   ```sh
   ~/video-audio-job
   ```
4. Tap sync in Autosync, or wait for its interval.
5. A minute or so after the file lands in Drive, a Matrix message tagged
   `source:: [[video note]]` arrives, dated by when you recorded the video.

## 10. Schedule it

```sh
termux-job-scheduler --job-id 4242 --script ~/video-audio-job \
  --period-ms 900000 --persisted true
termux-job-scheduler -p
```

- **Timing:** it runs once in every 15-minute window, at a point Android
  chooses.
- **Conditions:** it only runs while you're online and the battery isn't low.
  Both are the scheduler's defaults. It doesn't need the network itself, but
  nothing can upload offline anyway.
- **Reboots:** `--persisted true` keeps the job after a reboot.
- **To stop it:** `termux-job-scheduler --cancel --job-id 4242`

## Checking on it

```sh
tail -n 30 ~/.local/state/video-audio/log.txt   # what each run did
~/video-audio-job --dry-run                      # what the next run would do
```

The script also posts phone notifications:
- **"Video audio: gave up on X":** that video couldn't be read in 3 runs, so
  it was skipped.
- **"Video audio: too big to upload":** the audio is over 20 MB, so Autosync
  won't upload it. The file stays in `VideoAudio`; upload it to the Drive
  folder by hand.
- **"Video audio: run failed":** the whole run failed, typically because the
  storage permission was lost (see step 3). The notification replaces itself
  rather than stacking up.

**Troubleshooting:**
- **No new log lines for hours:** recheck step 8. On some Samsung phones
  periodic jobs fire erratically, with gaps of up to an hour now and then.
  `termux-job-scheduler -p` shows whether the job is still scheduled.
- **Runs start but never finish:** in *Developer options*, turn on *Disable
  child process restrictions*.
- **Audio uploaded but no note:** check the Apps Script execution log. The
  `video folder … files in window` line should count your files. If it shows
  0, check `VIDEO_FOLDER_IDS`, and check that the `Window start` it logs is
  before the video was recorded.
- **Process one particular video again:** delete its line from `done` in
  `~/.local/state/video-audio/state.json`.
- **Start over:** run `rm -rf ~/.local/state/video-audio`. The next run is a
  first run again; pass `--since` to set its start date.
