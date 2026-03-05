// ============================================================
// Voice Note Processor — Google Apps Script
// Watches Google Drive folders for new audio files,
// transcribes with gpt-4o-transcribe, post-processes with
// Claude Sonnet, and sends the result to Matrix.
// ============================================================

// --- Config helpers ---

function getConfig() {
  const props = PropertiesService.getScriptProperties();
  return {
    openaiKey: props.getProperty("OPENAI_API_KEY"),
    anthropicKey: props.getProperty("ANTHROPIC_API_KEY"),
    matrixAccessToken: props.getProperty("MATRIX_ACCESS_TOKEN"),
    matrixRoomId: props.getProperty("MATRIX_ROOM_ID"),
    // Comma-separated folder IDs to watch
    folderIds: (props.getProperty("FOLDER_IDS") || "").split(",").map(s => s.trim()).filter(Boolean),
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

  const audioMimeTypes = [
    "audio/mpeg", "audio/mp4", "audio/m4a", "audio/ogg",
    "audio/wav", "audio/x-wav", "audio/webm", "audio/aac",
    "audio/amr", "audio/3gpp", "video/mp4",
  ];

  Logger.log("Cutoff: " + cutoff + " (" + (cutoff ? new Date(cutoff).toISOString() : "none") + ")");
  Logger.log("Folder IDs: " + JSON.stringify(config.folderIds));

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
      const mime = file.getMimeType();
      const name = file.getName().toLowerCase();
      const ext = name.substring(name.lastIndexOf("."));
      if (!audioMimeTypes.includes(mime) && !audioExtensions.includes(ext)) continue;
      matched++;
      newFiles.push(file);
    }

    Logger.log("  " + scanned + " files from Drive, " + matched + " audio matches");
  }

  // Sort newest first
  newFiles.sort((a, b) => b.getDateCreated().getTime() - a.getDateCreated().getTime());

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

    // Update cutoff after each file
    const fileTime = file.getDateCreated().getTime();
    if (fileTime > cutoff) {
      props.setProperty("LAST_PROCESSED_TIME", String(fileTime));
    }
  }

  // If first run and we limited, set cutoff to the newest file in the folder
  // so older files are never picked up
  if (cutoff === 0 && newFiles.length > 0) {
    const newestTime = newFiles[0].getDateCreated().getTime();
    props.setProperty("LAST_PROCESSED_TIME", String(newestTime));
  }

  if (filesToProcess.length > 0) {
    Logger.log("Processed " + filesToProcess.length + " files this run.");
  }
}

// --- Process a single voice note ---

function processVoiceNote(file, config) {
  const fileName = file.getName();
  const webViewLink = file.getUrl();
  const created = file.getDateCreated();
  const tz = PropertiesService.getScriptProperties().getProperty("TIMEZONE") || "UTC";
  const timestamp = Utilities.formatDate(created, tz, "dd/MM/yyyy HH:mm:ss z");
  const footer = "\n- [" + fileName + "](" + webViewLink + ")" +
    "\n  - timestamp::" + timestamp;

  const transcription = transcribeAudio(file, config);
  if (!transcription || transcription.trim() === "") {
    return "- [[no speech detected]]" + footer;
  }
  const processed = postProcess(transcription, config);
  return processed + footer;
}

// --- Step 1: Transcribe with gpt-4o-transcribe ---

function transcribeAudio(file, config) {
  const blob = file.getBlob();
  const fileName = file.getName();

  // Build multipart/form-data payload
  const boundary = "----VoiceNoteFormBoundary" + Utilities.getUuid();

  const parts = [];

  // model field
  parts.push(
    "--" + boundary + "\r\n" +
    'Content-Disposition: form-data; name="model"\r\n\r\n' +
    "gpt-4o-transcribe\r\n"
  );

  // file field
  parts.push(
    "--" + boundary + "\r\n" +
    'Content-Disposition: form-data; name="file"; filename="' + fileName + '"\r\n' +
    "Content-Type: " + blob.getContentType() + "\r\n\r\n"
  );

  const closing = "\r\n--" + boundary + "--\r\n";

  // Assemble as byte array
  const preBytes = Utilities.newBlob(parts.join("")).getBytes();
  const fileBytes = blob.getBytes();
  const closingBytes = Utilities.newBlob(closing).getBytes();

  const payload = [...preBytes, ...fileBytes, ...closingBytes];

  const response = UrlFetchApp.fetch("https://api.openai.com/v1/audio/transcriptions", {
    method: "post",
    headers: {
      "Authorization": "Bearer " + config.openaiKey,
    },
    contentType: "multipart/form-data; boundary=" + boundary,
    payload: payload,
    muteHttpExceptions: true,
  });

  const code = response.getResponseCode();
  if (code !== 200) {
    throw new Error("OpenAI transcription failed (" + code + "): " + response.getContentText());
  }

  const result = JSON.parse(response.getContentText());
  Logger.log('Raw transcript:\n' + result.text)
  return result.text;
}

// --- Step 2: Post-process with Claude ---

function postProcess(transcription, config) {
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
- Tag people, projects, and notable topics with [[double brackets]] inline (e.g. [[John]], [[Project Alpha]])
- Tag any dates mentioned with Roam date format: [[Month DDth, YYYY]] (e.g. [[February 27th, 2026]], [[March 1st, 2026]])
- If there are action items, TODOs, or commitments mentioned, add a final top-level bullet "- **Action items:**" with each action as a nested bullet
- You are Claudia. The transcript may contain phrases addressed to "Claudia" — these are instructions to you. Keep them in the transcribed output as-is, but also follow them. Append your responses under a top-level bullet "- **Claudia:**" at the end, after the full transcription.`;

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
  Logger.log("Cleared processed file history. Next run will process only the " +
    (props.getProperty("INITIAL_LIMIT") || "1") + " most recent file(s), then track from there.");
}
