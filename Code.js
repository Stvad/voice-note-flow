// ============================================================
// Voice Note Processor — Google Apps Script
// Watches Google Drive folders for new audio files,
// transcribes with gpt-4o-transcribe, post-processes with
// Claude Sonnet, and sends the result to Matrix.
// ============================================================

// --- Config helpers ---

function getConfig() {
  const props = PropertiesService.getScriptProperties();
  const splitList = (s) => (s || "").split(",").map(x => x.trim()).filter(Boolean);
  return {
    deepgramKey: props.getProperty("DEEPGRAM_API_KEY"),
    anthropicKey: props.getProperty("ANTHROPIC_API_KEY"),
    matrixAccessToken: props.getProperty("MATRIX_ACCESS_TOKEN"),
    matrixRoomId: props.getProperty("MATRIX_ROOM_ID"),
    // Comma-separated folder IDs to watch
    folderIds: splitList(props.getProperty("FOLDER_IDS")),
    // KEYTERMS: canonical names/projects/topics used for Claude auto-linking
    // (can be large — only items in this list get [[bracketed]])
    keyterms: splitList(props.getProperty("KEYTERMS")),
    // ACOUSTIC_KEYTERMS: focused subset sent to Deepgram for transcription
    // boost. Optional — if empty, falls back to first 95 of KEYTERMS.
    // Deepgram Nova-3 caps at 100 keyterms per request.
    acousticKeyterms: splitList(props.getProperty("ACOUSTIC_KEYTERMS")),
  };
}

// --- Trigger setup (run once) ---

function setupTrigger() {
  // Remove existing triggers for this function
  ScriptApp.getProjectTriggers().forEach(t => {
    const fn = t.getHandlerFunction();
    if (fn === "onDriveChange" || fn === "checkNewVoiceNotes") {
      ScriptApp.deleteTrigger(t);
    }
  });
  ScriptApp.newTrigger("checkNewVoiceNotes")
    .timeBased()
    .everyMinutes(1)
    .create();
  Logger.log("Trigger installed: checkNewVoiceNotes (every 1 minute)");
}

// --- Main entry point ---

function checkNewVoiceNotes() {
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(0)) {
    Logger.log("Previous run still in progress, skipping.");
    return;
  }

  try {
    checkNewVoiceNotesLocked_();
  } finally {
    lock.releaseLock();
  }
}

function checkNewVoiceNotesLocked_() {
  // INITIAL_LIMIT: configurable via Script Properties, default 1
  // On first run (no cutoff stored), only process this many most recent files.
  // On subsequent runs, process all files newer than the cutoff.
  const config = getConfig();
  const props = PropertiesService.getScriptProperties();
  const cutoff = Number(props.getProperty("LAST_PROCESSED_TIME") || "0");
  const initialLimit = Number(props.getProperty("INITIAL_LIMIT") || "1");
  const recentIds = JSON.parse(props.getProperty("RECENT_FILE_IDS") || "[]");
  const recentSet = new Set(recentIds);

  const audioMimeTypes = [
    "audio/mpeg", "audio/mp4", "audio/m4a", "audio/ogg",
    "audio/wav", "audio/x-wav", "audio/webm", "audio/aac",
    "audio/amr", "audio/3gpp", "video/mp4",
  ];

  Logger.log("Cutoff: " + cutoff + " (" + (cutoff ? new Date(cutoff).toISOString() : "none") + ")");
  Logger.log("Folder IDs: " + JSON.stringify(config.folderIds));
  Logger.log("Recent IDs tracked: " + recentIds.length);

  // Collect new audio files across all folders using Drive search API
  const newFiles = [];
  const audioExtensions = [".mp3", ".m4a", ".ogg", ".wav", ".webm", ".aac", ".amr", ".3gp", ".mp4"];

  // Build date filter for Drive query (createdDate not supported by DriveApp, use modifiedDate)
  const cutoffDate = cutoff > 0
    ? " and modifiedDate > '" + new Date(cutoff).toISOString() + "'"
    : "";

  for (const folderId of config.folderIds) {
    const query = "'" + folderId + "' in parents and trashed = false" + cutoffDate;
    Logger.log("Searching: " + query);

    let files;
    try {
      files = DriveApp.searchFiles(query);
    } catch (e) {
      Logger.log("Could not search folder " + folderId + ": " + e.message);
      continue;
    }

    let scanned = 0;
    let matched = 0;

    while (files.hasNext()) {
      const file = files.next();
      scanned++;
      if (recentSet.has(file.getId())) continue;
      const mime = file.getMimeType();
      const name = file.getName().toLowerCase();
      const ext = name.substring(name.lastIndexOf("."));
      if (!audioMimeTypes.includes(mime) && !audioExtensions.includes(ext)) continue;
      matched++;
      newFiles.push(file);
    }

    Logger.log("  " + scanned + " files from Drive, " + matched + " new audio matches");
  }

  // Sort newest first by modification time (consistent with query filter)
  newFiles.sort((a, b) => b.getLastUpdated().getTime() - a.getLastUpdated().getTime());

  // If no cutoff yet (first run), limit to initialLimit most recent
  const filesToProcess = cutoff === 0 ? newFiles.slice(0, initialLimit) : newFiles;

  for (const file of filesToProcess) {
    Logger.log("Processing: " + file.getName() + " (" + file.getMimeType() + ")");

    try {
      const result = processVoiceNote(file, config);
      sendMatrixMessage(result, config);
      Logger.log("Sent to Matrix: " + file.getName() + "\n" + result);
    } catch (e) {
      Logger.log("Error processing " + file.getName() + ": " + e.message);
    }

    // Update cutoff using modifiedDate (consistent with Drive query)
    const fileTime = file.getLastUpdated().getTime();
    if (fileTime > cutoff) {
      props.setProperty("LAST_PROCESSED_TIME", String(fileTime));
    }

    // Track file ID to prevent reprocessing on re-upload
    recentSet.add(file.getId());
  }

  // If first run and we limited, set cutoff to the newest file across all results
  if (cutoff === 0 && newFiles.length > 0) {
    const newestTime = newFiles[0].getLastUpdated().getTime();
    props.setProperty("LAST_PROCESSED_TIME", String(newestTime));
  }

  // Save recent IDs (keep last 50 to avoid unbounded growth)
  if (filesToProcess.length > 0) {
    const allIds = [...recentSet];
    const trimmed = allIds.slice(Math.max(0, allIds.length - 50));
    props.setProperty("RECENT_FILE_IDS", JSON.stringify(trimmed));
    Logger.log("Processed " + filesToProcess.length + " files this run.");
  }
}

// --- Process a single voice note ---

const SUMMARY_THRESHOLD_CHARS = 1000;

function processVoiceNote(file, config) {
  const fileName = file.getName();
  const webViewLink = file.getUrl();
  const created = file.getDateCreated();
  const tz = PropertiesService.getScriptProperties().getProperty("TIMEZONE") || "UTC";
  const timestamp = Utilities.formatDate(created, tz, "dd/MM/yyyy HH:mm:ss z");
  const footer = "\n- audio-url::" + webViewLink +
    "\n- audio-file-name::" + fileName +
    "\n- timestamp::" + timestamp;

  const transcription = transcribeAudio(file, config);
  if (!transcription || transcription.trim() === "") {
    return "- [[no speech detected]]" + footer;
  }
  const processed = postProcess(transcription, config);

  let body;
  if (processed.length > SUMMARY_THRESHOLD_CHARS) {
    const summary = summarize(processed, config);
    const indented = processed.split("\n").map(l => l.length > 0 ? "  " + l : l).join("\n");
    body = "- " + summary + "\n" + indented;
  } else {
    body = processed;
  }
  return body + footer;
}

// --- Step 1: Transcribe with Deepgram Nova-3 ---

function transcribeAudio(file, config) {
  const blob = file.getBlob();

  const acoustic = config.acousticKeyterms.length > 0
    ? config.acousticKeyterms
    : config.keyterms.slice(0, 95);
  const params = ["model=nova-3", "smart_format=true"];
  for (const term of acoustic) {
    params.push("keyterm=" + encodeURIComponent(term));
  }
  const url = "https://api.deepgram.com/v1/listen?" + params.join("&");

  const response = UrlFetchApp.fetch(url, {
    method: "post",
    headers: {
      "Authorization": "Token " + config.deepgramKey,
    },
    contentType: blob.getContentType(),
    payload: blob.getBytes(),
    muteHttpExceptions: true,
  });

  const code = response.getResponseCode();
  if (code !== 200) {
    throw new Error("Deepgram transcription failed (" + code + "): " + response.getContentText());
  }

  const result = JSON.parse(response.getContentText());
  const transcript = result.results.channels[0].alternatives[0].transcript;
  Logger.log("Raw transcript:\n" + transcript);
  return transcript;
}

// --- Step 2: Post-process with Claude ---

function postProcess(transcription, config) {
  const keytermsList = config.keyterms.length > 0
    ? config.keyterms.map(t => `- ${t}`).join("\n")
    : "(none — do not auto-link any names, projects, or topics)";

  const systemPrompt = `You are a transcription post-processor. You clean up voice note transcriptions and format them as hierarchical bulleted lists.

Rules:
- Remove filler words and duplications
- Convert formatting words like "comma" into actual formatting
- Format output as a hierarchical bulleted list (Roam Research style)
- Limit to 5 levels of nesting
- Each list item begins with "-" and each indentation level is 2 spaces
- Do not write any text not in a list item
- The only markdown syntax allowed: "-"-style unordered lists + any inline formatting like **bold**, and [[double bracket tags]]
- Do not use headings, paragraphs, or other markdown
- Maintain first person voice if used
- Preserve original meaning faithfully — do not editorialize or add information
- Don't prefix the response (no "Sure, here is..." etc.)
- NEVER respond with meta-commentary about the transcription (e.g. "the transcript seems cut off", "could you provide more"). Always process whatever text you receive, no matter how short, fragmented, or incomplete. Your only job is to clean up and format what's there.
- Auto-link with [[double brackets]] ONLY for terms from the "Known terms" list below. Do not bracket any other names, projects, or topics — leave them as plain text.
- Fuzzy match: a partial or mistranscribed mention should still link to the canonical list entry. E.g. if the speaker says "Bobby" and the list contains "Bobby Smith", output [[Bobby Smith]]. If multiple list entries could plausibly match, pick the first one in list order.
- Hierarchical entries (slashes, e.g. "wcs/whip", "to/buy", "Roam/garden") may be matched either from the full form ("to buy", "wcs whip") OR from the base/last segment ("whip", "buy") WHEN context makes the domain unambiguous — e.g. "whip" → [[wcs/whip]] inside a passage about swing dancing; "to buy" → [[to/buy]] when expressing a buying intent. If context is ambiguous (e.g. "whip" in a cooking discussion), leave it as plain text.
- Only bracket the FIRST occurrence of each matched term in the note; subsequent mentions of the same term stay as plain text (use whatever the speaker actually said).
- Tag any dates mentioned with Roam date format: [[Month DDth, YYYY]] (e.g. [[February 27th, 2026]], [[March 1st, 2026]]). This rule applies regardless of the Known terms list, and applies to every date mention (not just the first).
- If there are multiple topics and some include action items, TODOs, or commitments, add a final top-level bullet "- **Action items:**" with each action as a nested bullet. But if the entire note is essentially one short action item, do NOT add a separate Action items section — that would just duplicate the content.
- You are Claudia. ONLY respond to instructions explicitly addressed to "Claudia" by name (e.g. "Claudia, look this up"). Do NOT interpret general statements as instructions — if the speaker is not addressing Claudia, it is transcript content. Keep Claudia-addressed phrases in the transcribed output as-is, and append your responses under a top-level bullet "- **Claudia:**" at the end.

Known terms (the user's canonical names/projects/topics). Two purposes: (1) if a transcript word is phonetically close to one of these and the context fits, correct it; (2) use them as the source of truth for [[double bracket]] auto-linking per the rules above:
${keytermsList}`;

  const payload = {
    model: "claude-sonnet-4-6",
    max_tokens: 4096,
    system: systemPrompt,
    messages: [
      { role: "user", content: transcription },
    ],
  };

  const response = UrlFetchApp.fetch("https://api.anthropic.com/v1/messages", {
    method: "post",
    headers: {
      "x-api-key": config.anthropicKey,
      "anthropic-version": "2023-06-01",
      "Content-Type": "application/json",
    },
    payload: JSON.stringify(payload),
    muteHttpExceptions: true,
  });

  const code = response.getResponseCode();
  if (code !== 200) {
    throw new Error("Anthropic API failed (" + code + "): " + response.getContentText());
  }

  const result = JSON.parse(response.getContentText());
  return result.content[0].text;
}

// --- Summarization (only for long notes) ---

function summarize(processed, config) {
  const systemPrompt = `You summarize voice note transcripts.

Rules:
- Output a single concise sentence capturing the gist of the note (under 25 words).
- Output ONLY the sentence text — no bullet prefix, no quotes, no meta-commentary.
- Do not editorialize or add information that is not in the note.
- Preserve [[double bracket tags]] from the original that are central to the gist.`;

  const payload = {
    model: "claude-sonnet-4-6",
    max_tokens: 300,
    system: systemPrompt,
    messages: [
      { role: "user", content: processed },
    ],
  };

  const response = UrlFetchApp.fetch("https://api.anthropic.com/v1/messages", {
    method: "post",
    headers: {
      "x-api-key": config.anthropicKey,
      "anthropic-version": "2023-06-01",
      "Content-Type": "application/json",
    },
    payload: JSON.stringify(payload),
    muteHttpExceptions: true,
  });

  const code = response.getResponseCode();
  if (code !== 200) {
    throw new Error("Anthropic summary API failed (" + code + "): " + response.getContentText());
  }

  const result = JSON.parse(response.getContentText());
  return result.content[0].text.trim();
}

// --- Step 3: Send to Matrix ---

function sendMatrixMessage(message, config) {
  const txnId = Utilities.getUuid();
  const url = "https://matrix.org/_matrix/client/v3/rooms/"
    + encodeURIComponent(config.matrixRoomId)
    + "/send/m.room.message/"
    + encodeURIComponent(txnId);

  const response = UrlFetchApp.fetch(url, {
    method: "put",
    headers: {
      "Authorization": "Bearer " + config.matrixAccessToken,
      "Content-Type": "application/json",
    },
    payload: JSON.stringify({
      body: message,
      msgtype: "m.text",
    }),
    muteHttpExceptions: true,
  });

  const code = response.getResponseCode();
  if (code !== 200) {
    throw new Error("Matrix send failed (" + code + "): " + response.getContentText());
  }
}

// --- Manual test helper ---

function testWithLatestFile() {
  const config = getConfig();
  Logger.log("Folder IDs configured: " + JSON.stringify(config.folderIds));
  const folderId = config.folderIds[0];
  if (!folderId) {
    Logger.log("ERROR: No FOLDER_IDS configured in Script Properties");
    return;
  }
  Logger.log("Using folder ID: " + folderId);
  const folder = DriveApp.getFolderById(folderId);
  Logger.log("Folder name: " + folder.getName());
  const files = folder.getFiles();
  if (!files.hasNext()) {
    Logger.log("No files found in folder");
    return;
  }
  const file = files.next();
  Logger.log("Testing with: " + file.getName() + " (" + file.getMimeType() + ")");
  const result = processVoiceNote(file, config);
  Logger.log("Result:\n" + result);
  // Uncomment to also send to Matrix:
  // sendMatrixMessage(result, config);
}

// --- Reset processed files (use to start fresh) ---

function resetProcessedFiles() {
  const props = PropertiesService.getScriptProperties();
  props.deleteProperty("PROCESSED_FILE_IDS");
  props.deleteProperty("LAST_PROCESSED_TIME");
  props.deleteProperty("RECENT_FILE_IDS");
  Logger.log("Cleared processed file history. Next run will process only the " +
    (props.getProperty("INITIAL_LIMIT") || "1") + " most recent file(s), then track from there.");
}
