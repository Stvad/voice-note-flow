// ============================================================
// Voice Note Processor — Google Apps Script
// Watches Google Drive folders for new audio files,
// transcribes with Deepgram Nova-3, post-processes with
// Claude Sonnet, and sends the result to Matrix.
//
// Which files are "done" is decided by a ledger of file IDs
// (SEEN_FILES), never by a high-water timestamp. The timestamp
// window only bounds how much of Drive we ask about, so a note
// that finishes syncing out of order is still picked up.
// See planRun_ / pruneState_ for the rules.
// ============================================================

// --- Config helpers ---

function getConfig() {
  const props = PropertiesService.getScriptProperties();
  const splitList = (s) => (s || "").split(",").map(x => x.trim()).filter(Boolean);
  const dedupe = (list) => {
    const seen = {};
    return list.filter(t => {
      const key = t.toLowerCase();
      if (seen[key]) return false;
      seen[key] = true;
      return true;
    });
  };
  const num = (key, fallback) => {
    const parsed = Number(props.getProperty(key));
    return isFinite(parsed) && parsed > 0 ? parsed : fallback;
  };
  // Shorthands the fuzzy matcher would otherwise resolve to the wrong person.
  const disambiguations = parseDisambiguations_(props.getProperty("KEYTERM_ALIASES"));
  const manual = splitList(props.getProperty("MANUAL_KEYTERMS"));
  // A disambiguation target has to be a Known term, or the "only bracket
  // terms from the list" rule forbids the very link we are asking for.
  const canonicals = disambiguations.map(d => d.canonical);
  return {
    deepgramKey: props.getProperty("DEEPGRAM_API_KEY"),
    anthropicKey: props.getProperty("ANTHROPIC_API_KEY"),
    matrixAccessToken: props.getProperty("MATRIX_ACCESS_TOKEN"),
    matrixRoomId: props.getProperty("MATRIX_ROOM_ID"),
    // Comma-separated folder IDs to watch
    folderIds: splitList(props.getProperty("FOLDER_IDS")),
    // KEYTERMS: canonical names/projects/topics used for Claude auto-linking
    // (can be large — only items in this list get [[bracketed]]).
    // MANUAL_KEYTERMS: hand-added terms, kept in a separate property so
    // regenerating KEYTERMS from a Roam dump never clobbers them. They go
    // first: the post-processing prompt breaks ambiguous matches by list
    // order, and a term you added by hand is the higher-signal one.
    keyterms: dedupe(
      canonicals
        .concat(manual)
        .concat(splitList(props.getProperty("KEYTERMS")))
    ),
    // KEYTERM_ALIASES: "shorthand => Canonical Term" pairs, comma-separated.
    // Ordering the list is not enough on its own — "Vlad" is an exact
    // first-token match for "Vlad Sterzhanov" but not a token of
    // "Vladyslav Sitalo" at all, so the matcher genuinely prefers the wrong
    // one. These are stated as rules instead.
    disambiguations: disambiguations,
    // ACOUSTIC_KEYTERMS: focused subset sent to Deepgram for transcription
    // boost. Optional — if empty, falls back to first 95 of KEYTERMS.
    // Deepgram Nova-3 caps at 100 keyterms per request. Manual terms lead
    // here too, since they survive the URL-length cap that way.
    acousticKeyterms: dedupe(
      manual.concat(splitList(props.getProperty("ACOUSTIC_KEYTERMS")))
    ),
    // How far back the Drive query looks. A note that finishes syncing within
    // this window of its own modified time is still processed, regardless of
    // what synced before it. Raise it before a long offline stretch.
    lookbackMs: num("LOOKBACK_DAYS", 7) * 24 * 60 * 60 * 1000,
    // On a virgin install, process only this many of the most recent files
    // rather than the whole back catalogue.
    initialLimit: num("INITIAL_LIMIT", 1),
    // Give up on a file after this many failed runs (then post a Matrix notice).
    maxAttempts: num("MAX_ATTEMPTS", 3),
    // Ledger capacity. Must comfortably exceed the number of notes you record
    // per lookback window — see pruneState_ for what happens if it doesn't.
    maxSeenEntries: num("MAX_SEEN_ENTRIES", 120),
    // Stop starting new files after this long, so we commit cleanly instead of
    // being killed mid-file by the Apps Script execution limit.
    runBudgetMs: num("RUN_BUDGET_SECONDS", 270) * 1000,
  };
}

// Script Properties cap each value at 9KB; stay clear of the edge.
const PROPERTY_VALUE_LIMIT = 8500;

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

// ============================================================
// Selection logic — pure functions, unit-tested in
// scripts/test_selection.js. No Apps Script services in here.
// ============================================================

const AUDIO_MIME_TYPES = [
  "audio/mpeg", "audio/mp4", "audio/m4a", "audio/ogg",
  "audio/wav", "audio/x-wav", "audio/webm", "audio/aac",
  "audio/amr", "audio/3gpp", "video/mp4",
];
const AUDIO_EXTENSIONS = [
  ".mp3", ".m4a", ".ogg", ".wav", ".webm", ".aac", ".amr", ".3gp", ".mp4",
];

// Parse KEYTERM_ALIASES: "Vlad => Vladyslav Sitalo, Ash => Ashley Qian".
// Malformed entries are skipped rather than throwing — a typo in a Script
// Property should not take the whole pipeline down.
function parseDisambiguations_(raw) {
  const out = [];
  const seen = {};
  for (const entry of String(raw || "").split(",")) {
    const idx = entry.indexOf("=>");
    if (idx === -1) continue;
    const shorthand = entry.slice(0, idx).trim();
    const canonical = entry.slice(idx + 2).trim();
    if (!shorthand || !canonical) continue;
    const key = shorthand.toLowerCase();
    if (seen[key]) continue; // first mapping wins
    seen[key] = true;
    out.push({ shorthand: shorthand, canonical: canonical });
  }
  return out;
}

function renderDisambiguations_(disambiguations) {
  if (!disambiguations || disambiguations.length === 0) return "";
  return "\n\nPreferred resolutions. When the speaker says one of these " +
    "shorthands, always link it to the canonical term paired with it here — " +
    "even if a different Known term looks like a closer or more frequent " +
    "match. These override the fuzzy-match and list-order rules above:\n" +
    disambiguations.map(d => `- "${d.shorthand}" -> [[${d.canonical}]]`).join("\n");
}

function isAudioFile_(name, mimeType) {
  if (AUDIO_MIME_TYPES.indexOf(mimeType) !== -1) return true;
  const lower = String(name || "").toLowerCase();
  const dot = lower.lastIndexOf(".");
  if (dot === -1) return false;
  return AUDIO_EXTENSIONS.indexOf(lower.substring(dot)) !== -1;
}

// The Drive query's lower bound. The floor is a hard limit the ledger imposes
// (see pruneState_); the rolling lookback takes over once it overtakes the
// floor, which is what lets a late-syncing note back into view.
function computeWindowStart_(floorMs, nowMs, lookbackMs) {
  return Math.max(floorMs || 0, nowMs - lookbackMs);
}

function byModifiedAsc_(a, b) {
  return (a.modifiedMs - b.modifiedMs) || (a.id < b.id ? -1 : a.id > b.id ? 1 : 0);
}

// Decide what to process this run.
//
// candidates: [{id, name, mimeType, modifiedMs, size}] from the Drive query
// state:      {floor, seen: {id: modifiedMs}, attempts: {id: count}}
// returns:    {windowStart, toProcess (oldest-first), floor, bootstrap, skipped}
//
// Oldest-first matters: the run may be cut short by the execution limit, and
// each file is committed to the ledger as it completes. Going oldest-first
// means an interrupted run leaves the *newest* files pending, and those are
// still inside the window next time. Newest-first would strand the older ones.
function planRun_(candidates, state, nowMs, opts) {
  const floor = state.floor || 0;
  const seen = state.seen || {};
  const attempts = state.attempts || {};
  const windowStart = computeWindowStart_(floor, nowMs, opts.lookbackMs);
  const skipped = { outsideWindow: 0, notAudio: 0, seen: 0, exhausted: 0, empty: 0 };

  const eligible = [];
  for (const file of candidates) {
    if (file.modifiedMs <= windowStart) { skipped.outsideWindow++; continue; }
    if (!isAudioFile_(file.name, file.mimeType)) { skipped.notAudio++; continue; }
    if (seen[file.id] !== undefined) { skipped.seen++; continue; }
    if ((attempts[file.id] || 0) >= opts.maxAttempts) { skipped.exhausted++; continue; }
    // A zero-byte file is still uploading. Leave it out of the ledger entirely
    // so the next run reconsiders it.
    if (file.size === 0) { skipped.empty++; continue; }
    eligible.push(file);
  }

  // floor === 0 means we have never run. Take only the newest few and raise the
  // floor past everything else, so a fresh install does not replay the archive.
  if (floor === 0) {
    if (eligible.length === 0) {
      return { windowStart: windowStart, toProcess: [], floor: 0, bootstrap: false, skipped: skipped };
    }
    const newestFirst = eligible.slice().sort((a, b) => byModifiedAsc_(b, a));
    const keep = newestFirst.slice(0, Math.max(1, opts.initialLimit));
    const dropped = newestFirst.slice(keep.length);
    const oldestKept = keep[keep.length - 1].modifiedMs;
    const newFloor = dropped.length > 0
      ? Math.min(dropped[0].modifiedMs, oldestKept - 1)
      : oldestKept - 1;
    return {
      windowStart: windowStart,
      toProcess: keep.sort(byModifiedAsc_),
      floor: newFloor,
      bootstrap: true,
      skipped: skipped,
    };
  }

  return {
    windowStart: windowStart,
    toProcess: eligible.sort(byModifiedAsc_),
    floor: floor,
    bootstrap: false,
    skipped: skipped,
  };
}

// Keep the ledger bounded without ever letting a processed file become
// eligible again.
//
// Entries below the window can be forgotten for free — the query will not
// return those files anyway. If the ledger is *still* over capacity after
// that, we forget the oldest entries and raise the floor to match, which
// pulls the window up so the forgotten files stay out of scope. That is the
// invariant the whole design rests on: the window never reaches further back
// than the ledger can remember.
function pruneState_(state, windowStart, maxSeenEntries, liveIds) {
  let floor = state.floor || 0;
  const seen = {};
  for (const id in state.seen) {
    if (state.seen[id] > windowStart) seen[id] = state.seen[id];
  }

  const entries = Object.keys(seen).map(id => ({ id: id, modifiedMs: seen[id] }));
  if (entries.length > maxSeenEntries) {
    entries.sort(byModifiedAsc_);
    const dropped = entries.slice(0, entries.length - maxSeenEntries);
    for (const entry of dropped) delete seen[entry.id];
    floor = Math.max(floor, dropped[dropped.length - 1].modifiedMs);
  }

  // Attempt counters are only meaningful for files still in view and not yet
  // done; everything else expires.
  const attempts = {};
  for (const id in state.attempts) {
    if (liveIds && !liveIds.has(id)) continue;
    if (seen[id] !== undefined) continue;
    attempts[id] = state.attempts[id];
  }

  return { floor: floor, seen: seen, attempts: attempts };
}

// ============================================================
// State persistence
// ============================================================

function loadState_(props, nowMs) {
  const rawSeen = props.getProperty("SEEN_FILES");
  if (rawSeen !== null) {
    return {
      floor: Number(props.getProperty("PROCESS_FLOOR_TIME") || "0"),
      seen: JSON.parse(rawSeen),
      attempts: JSON.parse(props.getProperty("FAILED_ATTEMPTS") || "{}"),
    };
  }

  // Migrate from the old LAST_PROCESSED_TIME high-water mark. The old cutoff
  // becomes the floor, so nothing already handled gets replayed; the rolling
  // lookback takes over from there. Migrated IDs are stamped with *now* rather
  // than their real modified time so they survive a full lookback window —
  // the timestamp in the ledger only decides when an entry may be forgotten.
  const legacyFloor = Number(props.getProperty("LAST_PROCESSED_TIME") || "0");
  const legacyIds = JSON.parse(props.getProperty("RECENT_FILE_IDS") || "[]");
  const seen = {};
  for (const id of legacyIds) seen[id] = nowMs;

  // Persist both halves now, not at the end of the run. commitSeen_ writes
  // SEEN_FILES as files complete, so if this run dies early we would otherwise
  // come back with a ledger but no floor — which planRun_ reads as a virgin
  // install and would silently skip everything pending.
  props.setProperty("PROCESS_FLOOR_TIME", String(legacyFloor));
  props.setProperty("SEEN_FILES", JSON.stringify(seen));
  Logger.log("Migrated legacy state: floor=" + new Date(legacyFloor).toISOString() +
    ", " + legacyIds.length + " known file IDs");
  return { floor: legacyFloor, seen: seen, attempts: {} };
}

// Commit a single file's outcome immediately. Doing this per file (rather than
// once at the end) is what makes an interrupted run safe.
function commitSeen_(props, state, file) {
  state.seen[file.id] = file.modifiedMs;
  delete state.attempts[file.id];
  props.setProperty("SEEN_FILES", JSON.stringify(state.seen));
  props.setProperty("FAILED_ATTEMPTS", JSON.stringify(state.attempts));
}

function commitAttempt_(props, state, file) {
  const count = (state.attempts[file.id] || 0) + 1;
  state.attempts[file.id] = count;
  props.setProperty("FAILED_ATTEMPTS", JSON.stringify(state.attempts));
  return count;
}

function saveState_(props, state, windowStart, liveIds, config) {
  let cap = config.maxSeenEntries;
  let pruned = pruneState_(state, windowStart, cap, liveIds);
  // Defensive: if the ledger somehow still does not fit in a Script Property,
  // shrink until it does. pruneState_ raises the floor accordingly.
  while (JSON.stringify(pruned.seen).length > PROPERTY_VALUE_LIMIT && cap > 1) {
    cap = Math.floor(cap / 2);
    pruned = pruneState_(state, windowStart, cap, liveIds);
    Logger.log("WARNING: ledger too large for a Script Property; capped at " + cap +
      " entries and raised the floor. Lower LOOKBACK_DAYS or expect less " +
      "tolerance for late-syncing notes.");
  }

  props.setProperty("PROCESS_FLOOR_TIME", String(pruned.floor));
  props.setProperty("SEEN_FILES", JSON.stringify(pruned.seen));
  props.setProperty("FAILED_ATTEMPTS", JSON.stringify(pruned.attempts));
  state.floor = pruned.floor;
  state.seen = pruned.seen;
  state.attempts = pruned.attempts;
}

// ============================================================
// Main loop
// ============================================================

// Read every file in the window across all watched folders, as plain objects
// (plus the DriveApp handle) so the selection logic stays pure.
function collectCandidates_(config, windowStart) {
  const candidates = [];
  const since = new Date(windowStart).toISOString();

  for (const folderId of config.folderIds) {
    // createdDate is not supported by DriveApp's query language, so this
    // filters on modifiedDate. That is exactly why the ledger — not this
    // timestamp — decides what has been processed.
    const query = "'" + folderId + "' in parents and trashed = false" +
      " and modifiedDate > '" + since + "'";

    let files;
    try {
      files = DriveApp.searchFiles(query);
    } catch (e) {
      Logger.log("Could not search folder " + folderId + ": " + e.message);
      continue;
    }

    let scanned = 0;
    while (files.hasNext()) {
      const file = files.next();
      scanned++;
      candidates.push({
        id: file.getId(),
        name: file.getName(),
        mimeType: file.getMimeType(),
        modifiedMs: file.getLastUpdated().getTime(),
        size: file.getSize(),
        file: file,
      });
    }
    Logger.log("  folder " + folderId + ": " + scanned + " files in window");
  }

  return candidates;
}

function checkNewVoiceNotesLocked_() {
  const config = getConfig();
  const props = PropertiesService.getScriptProperties();
  const runStart = Date.now();
  const state = loadState_(props, runStart);

  const windowStart = computeWindowStart_(state.floor, runStart, config.lookbackMs);
  Logger.log("Window start: " + new Date(windowStart).toISOString() +
    " (floor " + (state.floor ? new Date(state.floor).toISOString() : "none") +
    ", lookback " + (config.lookbackMs / 86400000) + "d)");
  Logger.log("Ledger: " + Object.keys(state.seen).length + " known, " +
    Object.keys(state.attempts).length + " with failed attempts");

  const candidates = collectCandidates_(config, windowStart);
  const plan = planRun_(candidates, state, runStart, config);
  if (plan.bootstrap) {
    state.floor = plan.floor;
    props.setProperty("PROCESS_FLOOR_TIME", String(plan.floor));
    Logger.log("First run: processing " + plan.toProcess.length + " most recent file(s); " +
      "floor set to " + new Date(plan.floor).toISOString());
  }
  Logger.log("Plan: " + plan.toProcess.length + " to process, skipped " +
    JSON.stringify(plan.skipped));

  let done = 0;
  for (const item of plan.toProcess) {
    if (Date.now() - runStart > config.runBudgetMs) {
      Logger.log("Run budget reached; deferring " + (plan.toProcess.length - done) +
        " file(s) to the next run.");
      break;
    }
    done++;
    Logger.log("Processing: " + item.name + " (" + item.mimeType + ")");

    try {
      const result = processVoiceNote(item.file, config);
      sendMatrixMessage(result, config);
      commitSeen_(props, state, item);
      Logger.log("Sent to Matrix: " + item.name + "\n" + result);
    } catch (e) {
      const attempt = commitAttempt_(props, state, item);
      Logger.log("Error processing " + item.name + " (attempt " + attempt + "/" +
        config.maxAttempts + "): " + e.message);
      if (attempt >= config.maxAttempts) {
        // Out of retries. Record it as handled so it stops blocking the queue,
        // but say so out loud rather than dropping it silently.
        commitSeen_(props, state, item);
        notifyGiveUp_(item, e, config);
      }
    }
  }

  const liveIds = new Set(candidates.map(c => c.id));
  saveState_(props, state, windowStart, liveIds, config);
}

function notifyGiveUp_(item, err, config) {
  try {
    // getUrl() is a Drive call and can itself fail, so it stays inside the try.
    sendMatrixMessage(
      "- [[voice note failed]] " + item.name +
      "\n  - Gave up after " + config.maxAttempts + " attempts" +
      "\n  - error:: " + err.message +
      "\n  - audio-url::" + item.file.getUrl(),
      config
    );
  } catch (e) {
    Logger.log("Could not send failure notice for " + item.name + ": " + e.message);
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

  // Apps Script UrlFetchApp caps URLs at 2KB. Build the URL incrementally and
  // stop appending keyterms once we get close to the limit.
  const URL_LIMIT = 2000;
  const base = "https://api.deepgram.com/v1/listen?model=nova-3&smart_format=true";
  const acoustic = config.acousticKeyterms.length > 0
    ? config.acousticKeyterms
    : config.keyterms.slice(0, 95);
  let url = base;
  let included = 0;
  for (const term of acoustic) {
    const next = url + "&keyterm=" + encodeURIComponent(term);
    if (next.length > URL_LIMIT) break;
    url = next;
    included++;
  }
  if (included < acoustic.length) {
    Logger.log("URL length cap: sent " + included + "/" + acoustic.length + " keyterms to Deepgram");
  }

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
- Preserve EVERY sentence the speaker said. Including: opening remarks ("okay, let's see..."), test scaffolding ("checking if this works"), asides, throwaway comments, self-corrections. "Filler" means filler WORDS (uh, um, like, you know, sort of) and verbal repetitions of the same phrase — NOT entire sentences. If you're unsure whether something is content or filler, KEEP IT. Never drop a sentence because it seems unimportant, meta, or like the speaker was just warming up.
- Remove filler words and duplications (per the above scope)
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
- Links are ALWAYS plain [[Canonical Term]] with nothing after the term. NEVER write alias/pipe syntax like [[Bobby Smith|Bobby]] or [[Bobby Smith|he]] — that is Obsidian syntax, and it renders as broken text here. When the speaker uses a shorthand, replace it outright with the canonical form: "Bobby" becomes [[Bobby Smith]], not [[Bobby Smith|Bobby]].
- Hierarchical entries (slashes, e.g. "wcs/whip", "to/buy", "Roam/garden") may be matched either from the full form ("to buy", "wcs whip") OR from the base/last segment ("whip", "buy") WHEN context makes the domain unambiguous — e.g. "whip" → [[wcs/whip]] inside a passage about swing dancing; "to buy" → [[to/buy]] when expressing a buying intent. If context is ambiguous (e.g. "whip" in a cooking discussion), leave it as plain text.
- Only bracket the FIRST occurrence of each matched term in the note; subsequent mentions of the same term stay as plain text (use whatever the speaker actually said).
- Tag any dates mentioned with Roam date format: [[Month DDth, YYYY]] (e.g. [[February 27th, 2026]], [[March 1st, 2026]]). This rule applies regardless of the Known terms list, and applies to every date mention (not just the first).

Known terms (the user's canonical names/projects/topics). Two purposes: (1) if a transcript word is phonetically close to one of these and the context fits, correct it; (2) use them as the source of truth for [[double bracket]] auto-linking per the rules above:
${keytermsList}${renderDisambiguations_(config.disambiguations)}`;

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

// --- Keyterm inspection ---
//
// Prints the effective lists and how close each stored property is to the
// Script Property 9KB per-value cap (oversized writes are silently dropped).
// Run this before regenerating KEYTERMS so hand-added terms can be moved into
// MANUAL_KEYTERMS first, where a regeneration will not touch them.

function dumpKeyterms() {
  const props = PropertiesService.getScriptProperties();
  const config = getConfig();
  const raw = (key) => props.getProperty(key) || "";

  for (const key of ["KEYTERMS", "MANUAL_KEYTERMS", "ACOUSTIC_KEYTERMS"]) {
    const value = raw(key);
    const count = value ? value.split(",").filter(s => s.trim()).length : 0;
    Logger.log(key + ": " + count + " terms, " + value.length + " bytes" +
      (value.length > 8500 ? "  <-- NEAR THE 9KB CAP" : ""));
  }
  Logger.log("\nEffective link list (" + config.keyterms.length + " terms):\n" +
    config.keyterms.join(","));
  Logger.log("\nEffective acoustic list (" + config.acousticKeyterms.length + " terms):\n" +
    config.acousticKeyterms.join(","));
}

// --- Backfill: recover notes stranded below the floor ---
//
// Notes missed by the old high-water-mark scheme sit below the current floor,
// so the normal loop will never see them. This lists them without touching
// anything; nothing is sent to Matrix.
//
// Run backfillDryRun() from the editor, or set BACKFILL_DAYS to change the
// range (default 14).

function backfillDryRun(daysBack) {
  const config = getConfig();
  const props = PropertiesService.getScriptProperties();
  const days = daysBack || Number(props.getProperty("BACKFILL_DAYS") || "14");
  const now = Date.now();
  const state = loadState_(props, now);
  const since = now - days * 24 * 60 * 60 * 1000;

  const candidates = collectCandidates_(config, since)
    .filter(c => isAudioFile_(c.name, c.mimeType) && state.seen[c.id] === undefined)
    .sort(byModifiedAsc_);

  Logger.log("Unprocessed audio in the last " + days + " days: " + candidates.length);
  for (const c of candidates) {
    Logger.log("  " + new Date(c.modifiedMs).toISOString() + "  " + c.name + "  " + c.id);
  }
  Logger.log("\nTo process these: set BACKFILL_FILE_IDS to a comma-separated list " +
    "of the IDs above, then run backfillProcess().");
  return candidates.map(c => c.id);
}

// Process an explicit list of file IDs and add them to the ledger. Reads
// BACKFILL_FILE_IDS unless passed a list directly.
function backfillProcess(fileIds) {
  const config = getConfig();
  const props = PropertiesService.getScriptProperties();
  const ids = fileIds || (props.getProperty("BACKFILL_FILE_IDS") || "")
    .split(",").map(s => s.trim()).filter(Boolean);
  if (ids.length === 0) {
    Logger.log("Nothing to do: pass file IDs or set BACKFILL_FILE_IDS.");
    return;
  }

  const state = loadState_(props, Date.now());
  for (const id of ids) {
    let file;
    try {
      file = DriveApp.getFileById(id);
    } catch (e) {
      Logger.log("Skipping " + id + ": " + e.message);
      continue;
    }
    if (state.seen[id] !== undefined) {
      Logger.log("Skipping " + file.getName() + ": already in the ledger.");
      continue;
    }
    const item = { id: id, name: file.getName(), modifiedMs: file.getLastUpdated().getTime(), file: file };
    Logger.log("Backfilling: " + item.name);
    try {
      const result = processVoiceNote(file, config);
      sendMatrixMessage(result, config);
      commitSeen_(props, state, item);
      Logger.log("Sent to Matrix: " + item.name);
    } catch (e) {
      Logger.log("Error backfilling " + item.name + ": " + e.message);
    }
  }
}

// --- Reset processed files (use to start fresh) ---

function resetProcessedFiles() {
  const props = PropertiesService.getScriptProperties();
  props.deleteProperty("PROCESSED_FILE_IDS");
  props.deleteProperty("LAST_PROCESSED_TIME");
  props.deleteProperty("RECENT_FILE_IDS");
  props.deleteProperty("PROCESS_FLOOR_TIME");
  props.deleteProperty("SEEN_FILES");
  props.deleteProperty("FAILED_ATTEMPTS");
  Logger.log("Cleared processed file history. Next run will process only the " +
    (props.getProperty("INITIAL_LIMIT") || "1") + " most recent file(s), then track from there.");
}

// --- Node interop (for scripts/test_selection.js) ---
// `module` is undefined under Apps Script, so this is a no-op there.

if (typeof module !== "undefined") {
  module.exports = {
    computeWindowStart_: computeWindowStart_,
    isAudioFile_: isAudioFile_,
    planRun_: planRun_,
    pruneState_: pruneState_,
    parseDisambiguations_: parseDisambiguations_,
    renderDisambiguations_: renderDisambiguations_,
  };
}
