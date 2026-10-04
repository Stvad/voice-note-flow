# Video notes: phone setup (Galaxy S26 Ultra)

```
DCIM/Camera/*.mp4
  └─ Termux job, every 15 min: extract_video_audio.py → mono .m4a
       └─ /storage/emulated/0/VideoAudio ──Autosync──> Drive "Video audio"
            └─ Code.js (VIDEO_FOLDER_IDS) → Deepgram → Claude → Matrix
```

Google Photos backup carries on unchanged; this reads the same camera folder
alongside it. Allow about 30 minutes. Do the steps in order, since later ones
depend on earlier ones.

How it behaves:
- **Existing videos:** videos already on the phone when you first run it are
  skipped. Only videos recorded after that are picked up.
- **Silent clips:** clips with no speech post nothing to Matrix. Clips with
  speech arrive tagged `source:: [[video note]]`.
- **Delay:** expect roughly 15–30 minutes from recording to the Matrix message.

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

```sh
git clone https://github.com/Stvad/voice-note-flow.git ~/voice-note-flow

cat > ~/video-audio-job <<'EOF'
#!/data/data/com.termux/files/usr/bin/sh
exec python3 "$HOME/voice-note-flow/android/extract_video_audio.py" "$@"
EOF
chmod +x ~/video-audio-job
```

The scheduled job runs `~/video-audio-job`. It can't pass arguments to the
script, so any option you want goes into this wrapper.

> **Autosync free version only.** The free version allows one folder pair. If
> that pair is already used for voice notes, the outbox has to live inside its
> local folder. Edit the `exec` line so it reads:
> `exec python3 "$HOME/voice-note-flow/android/extract_video_audio.py" --outbox "/storage/emulated/0/<voice notes folder>/VideoAudio" "$@"`
> Use that path instead of `/storage/emulated/0/VideoAudio` everywhere below.
> Free also caps uploads at 10 MB, which is about 20 minutes of audio. Longer
> clips won't upload.

To update the script later: `git -C ~/voice-note-flow pull`

## 5. First run: set the starting point

```sh
~/video-audio-job --dry-run
~/video-audio-job
```

The first run logs `first run: N video(s) ... are treated as already handled`.
It also creates `/storage/emulated/0/VideoAudio`, the outbox Autosync uploads
from. Nothing in your existing camera roll is processed.

## 6. Autosync: upload the outbox

**Pro version (recommended):** add a second folder pair in the *Synced
Folders* tab.

- **Local folder:** `VideoAudio`, in internal storage
- **Remote folder:** create a new one, e.g. `Video audio`. Use the folder icon
  with a plus sign.
- **Sync method:** **Upload Then Delete**. Choose Upload Only if you'd rather
  keep the `.m4a` files on the phone.
- **Exclude pattern:** `.*`. This keeps the `.nomedia` marker off Drive.
- **Instant upload:** on, for this pair only. The files are small, and without
  it you wait for the autosync interval as well.

**Free version:** nothing to add. The existing pair picks up the new
`VideoAudio` subfolder. Add `.*` to its exclude patterns if it isn't there
already.

Tap sync once. The `Video audio` folder should appear in Drive.

## 7. Point Apps Script at the Drive folder

1. Open the Drive folder in a browser. Its ID is the last part of the URL:
   `drive.google.com/drive/folders/<ID>`. On the free version, open the
   `VideoAudio` subfolder inside your voice-notes folder, not the parent.
2. In the Apps Script editor, go to *Project Settings* (the gear icon), then
   *Script Properties*. Add `VIDEO_FOLDER_IDS` and set it to that ID.

(Code.js has to be deployed with `clasp push` from the Mac first.)

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
4. Autosync uploads the `.m4a`. With Upload Then Delete, it disappears from
   `VideoAudio` once uploaded.
5. A minute or so after it lands in Drive, a Matrix message tagged
   `source:: [[video note]]` arrives.

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
- **"Video audio: run failed":** the whole run failed, typically because the
  storage permission was lost (see step 3). The notification replaces itself
  rather than stacking up.

**Troubleshooting:**
- **No new log lines for hours:** recheck step 8. On some Samsung phones
  periodic jobs fire erratically, with gaps of up to an hour now and then.
  `termux-job-scheduler -p` shows whether the job is still scheduled.
- **Runs start but never finish:** in *Developer options*, turn on *Disable
  child process restrictions*.
- **Process one particular video again:** delete its line from `done` in
  `~/.local/state/video-audio/state.json`.
- **Start over:** run `rm -rf ~/.local/state/video-audio`. The next run is a
  first run again, so videos already on the phone are skipped.
