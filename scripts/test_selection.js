// Unit tests for the pure file-selection logic in Code.js.
// Run: node --test scripts/test_selection.js
//
// These cover the cases that are painful to reproduce by hand in Apps Script:
// out-of-order sync arrival, mid-run timeouts, and ledger overflow.

const test = require("node:test");
const assert = require("node:assert");

const {
  computeWindowStart_,
  isAudioFile_,
  planRun_,
  pruneState_,
} = require("../Code.js");

const HOUR = 3600 * 1000;
const DAY = 24 * HOUR;
const NOW = 1755300000000; // fixed clock so tests are deterministic

const OPTS = {
  lookbackMs: 7 * DAY,
  initialLimit: 1,
  maxAttempts: 3,
  maxSeenEntries: 150,
};

// Candidate factory: `ageHours` is how far in the past the file was modified.
function cand(id, ageHours, extra) {
  return Object.assign({
    id: id,
    name: id + ".m4a",
    mimeType: "audio/m4a",
    modifiedMs: NOW - ageHours * HOUR,
    size: 1000,
  }, extra || {});
}

function state(extra) {
  return Object.assign({ floor: NOW - 30 * DAY, seen: {}, attempts: {} }, extra || {});
}

const ids = (plan) => plan.toProcess.map(f => f.id);

// --- computeWindowStart_ ---

test("window start: rolling lookback wins when the floor is older", () => {
  assert.strictEqual(
    computeWindowStart_(NOW - 30 * DAY, NOW, 7 * DAY),
    NOW - 7 * DAY
  );
});

test("window start: floor wins when it is newer than the rolling lookback", () => {
  assert.strictEqual(
    computeWindowStart_(NOW - 2 * DAY, NOW, 7 * DAY),
    NOW - 2 * DAY
  );
});

// --- isAudioFile_ ---

test("audio detection: by extension when the mime type is unhelpful", () => {
  assert.strictEqual(isAudioFile_("Recording 42.M4A", "application/octet-stream"), true);
});

test("audio detection: by mime type when the extension is missing", () => {
  assert.strictEqual(isAudioFile_("voicenote", "audio/mpeg"), true);
});

test("audio detection: rejects non-audio, and does not choke on dotless names", () => {
  assert.strictEqual(isAudioFile_("notes.txt", "text/plain"), false);
  assert.strictEqual(isAudioFile_("README", "text/plain"), false);
});

// --- planRun_: the bug this whole change exists to fix ---

test("late-arriving older note is still selected after a newer one was processed", () => {
  // The exact failure: a big note recorded at T-5h finishes syncing *after* a
  // small note recorded at T-2h has already been processed. Under a high-water
  // sentinel the older file is invisible forever; the ledger must save it.
  const older = cand("big-old-note", 5);
  const newer = cand("small-new-note", 2);
  const s = state({ seen: { "small-new-note": newer.modifiedMs } });

  const plan = planRun_([older, newer], s, NOW, OPTS);

  assert.deepStrictEqual(ids(plan), ["big-old-note"]);
});

test("already-processed files are skipped", () => {
  const a = cand("a", 3);
  const s = state({ seen: { a: a.modifiedMs } });

  const plan = planRun_([a], s, NOW, OPTS);

  assert.deepStrictEqual(ids(plan), []);
  assert.strictEqual(plan.skipped.seen, 1);
});

test("selection is oldest-first so a timeout strands only the newest", () => {
  const plan = planRun_(
    [cand("newest", 1), cand("oldest", 9), cand("middle", 5)],
    state(),
    NOW,
    OPTS
  );

  assert.deepStrictEqual(ids(plan), ["oldest", "middle", "newest"]);
});

test("a run that dies partway through leaves the rest selectable next time", () => {
  const batch = [cand("a", 9), cand("b", 6), cand("c", 3)];
  const s = state();

  const first = planRun_(batch, s, NOW, OPTS);
  assert.deepStrictEqual(ids(first), ["a", "b", "c"]);

  // Simulate: "a" completed and was committed, then the 6-minute limit hit.
  s.seen["a"] = batch[0].modifiedMs;

  const second = planRun_(batch, s, NOW + 60000, OPTS);
  assert.deepStrictEqual(ids(second), ["b", "c"]);
});

test("still-uploading (zero-byte) files are deferred, not consumed", () => {
  const plan = planRun_([cand("partial", 1, { size: 0 })], state(), NOW, OPTS);

  assert.deepStrictEqual(ids(plan), []);
  assert.strictEqual(plan.skipped.empty, 1);
  // Crucially it is NOT marked seen, so the next run picks it up once complete.
  assert.strictEqual(plan.seen === undefined || !("partial" in (plan.seen || {})), true);
});

test("non-audio files in the folder are ignored", () => {
  const plan = planRun_(
    [cand("doc", 1, { name: "doc.pdf", mimeType: "application/pdf" })],
    state(),
    NOW,
    OPTS
  );

  assert.deepStrictEqual(ids(plan), []);
  assert.strictEqual(plan.skipped.notAudio, 1);
});

test("files at or below the window start are out of scope", () => {
  const plan = planRun_(
    [cand("ancient", 24 * 10), cand("fresh", 1)],
    state(),
    NOW,
    OPTS
  );

  assert.deepStrictEqual(ids(plan), ["fresh"]);
  assert.strictEqual(plan.skipped.outsideWindow, 1);
});

test("a file that keeps failing is dropped after maxAttempts", () => {
  const s = state({ attempts: { flaky: 3 } });

  const plan = planRun_([cand("flaky", 2)], s, NOW, OPTS);

  assert.deepStrictEqual(ids(plan), []);
  assert.strictEqual(plan.skipped.exhausted, 1);
});

test("a file below maxAttempts is retried", () => {
  const s = state({ attempts: { flaky: 2 } });

  assert.deepStrictEqual(ids(planRun_([cand("flaky", 2)], s, NOW, OPTS)), ["flaky"]);
});

// --- planRun_: first-run bootstrap ---

test("virgin install processes only INITIAL_LIMIT newest and floors out the rest", () => {
  const batch = [cand("a", 40), cand("b", 30), cand("c", 20), cand("d", 10)];
  const s = state({ floor: 0 });

  const plan = planRun_(batch, s, NOW, Object.assign({}, OPTS, { initialLimit: 2 }));

  assert.deepStrictEqual(ids(plan), ["c", "d"], "oldest-first among the kept ones");
  assert.strictEqual(plan.bootstrap, true);
  // The floor must exclude a/b but keep c selectable-by-window.
  assert.ok(plan.floor < batch[2].modifiedMs, "floor keeps 'c' inside the window");
  assert.ok(plan.floor >= batch[1].modifiedMs, "floor pushes 'b' out of the window");
});

test("virgin install with an empty folder stays virgin", () => {
  const plan = planRun_([], state({ floor: 0 }), NOW, OPTS);

  assert.deepStrictEqual(ids(plan), []);
  assert.strictEqual(plan.floor, 0, "do not bootstrap off an empty folder");
});

test("a set floor with an empty ledger is NOT a virgin install", () => {
  // loadState_ persists the floor and the ledger together for this reason: if
  // the floor were ever missing while files were pending, this would bootstrap
  // and silently skip them.
  const plan = planRun_([cand("pending", 2)], state({ floor: NOW - DAY, seen: {} }), NOW, OPTS);

  assert.strictEqual(plan.bootstrap, false);
  assert.deepStrictEqual(ids(plan), ["pending"]);
});

test("after migrating off the high-water mark, late syncs are caught", () => {
  // Post-migration state: floor is the old LAST_PROCESSED_TIME, ledger seeded
  // from RECENT_FILE_IDS. A note recorded before the last processed one but
  // synced after it must now be picked up.
  const migrated = state({
    floor: NOW - 2 * DAY,
    seen: { "already-sent": NOW - HOUR },
  });

  const plan = planRun_(
    [cand("synced-late", 5), cand("already-sent", 1)],
    migrated,
    NOW,
    OPTS
  );

  assert.deepStrictEqual(ids(plan), ["synced-late"]);
  assert.strictEqual(plan.floor, migrated.floor, "the floor never advances on its own");
});

// --- pruneState_ ---

test("prune drops ledger entries that fell out of the window, floor untouched", () => {
  const s = state({
    floor: NOW - 30 * DAY,
    seen: { old: NOW - 9 * DAY, recent: NOW - 1 * DAY },
  });
  const windowStart = computeWindowStart_(s.floor, NOW, 7 * DAY);

  const pruned = pruneState_(s, windowStart, 150, new Set(["recent"]));

  assert.deepStrictEqual(Object.keys(pruned.seen), ["recent"]);
  assert.strictEqual(pruned.floor, s.floor, "window already excludes it; no floor bump needed");
});

test("ledger overflow raises the floor instead of silently forgetting", () => {
  const seen = {};
  for (let i = 0; i < 10; i++) seen["f" + i] = NOW - (10 - i) * HOUR; // f0 oldest
  const s = state({ floor: NOW - 30 * DAY, seen: seen });
  const windowStart = computeWindowStart_(s.floor, NOW, 7 * DAY);

  const pruned = pruneState_(s, windowStart, 4, new Set());

  assert.strictEqual(Object.keys(pruned.seen).length, 4);
  assert.deepStrictEqual(Object.keys(pruned.seen).sort(), ["f6", "f7", "f8", "f9"]);
  assert.strictEqual(pruned.floor, seen["f5"], "floor rises to the newest dropped entry");
});

test("INVARIANT: a file forgotten by ledger overflow is never reprocessed", () => {
  // The window must never reach further back than the ledger can remember.
  const batch = [];
  const seen = {};
  for (let i = 0; i < 10; i++) {
    const c = cand("f" + i, 10 - i);
    batch.push(c);
    seen[c.id] = c.modifiedMs;
  }
  const s = state({ floor: NOW - 30 * DAY, seen: seen });
  const windowStart = computeWindowStart_(s.floor, NOW, 7 * DAY);

  const pruned = pruneState_(s, windowStart, 4, new Set(batch.map(c => c.id)));
  const next = planRun_(batch, pruned, NOW + HOUR, OPTS);

  assert.deepStrictEqual(ids(next), [], "no duplicates after the ledger overflowed");
});

test("attempt counters expire once the file leaves the window", () => {
  const s = state({ attempts: { gone: 2, present: 1 } });

  const pruned = pruneState_(s, NOW - 7 * DAY, 150, new Set(["present"]));

  assert.deepStrictEqual(pruned.attempts, { present: 1 });
});

test("attempt counters are cleared once the file succeeds", () => {
  const s = state({ seen: { done: NOW - HOUR }, attempts: { done: 2 } });

  const pruned = pruneState_(s, NOW - 7 * DAY, 150, new Set(["done"]));

  assert.deepStrictEqual(pruned.attempts, {});
});
