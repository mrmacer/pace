/* ─────────────────────────────────────────────────────────────────────────
   LIVE VISIT NOTES — pre-start notes, same-row draft updates, recovery UX,
   and completion hand-off coverage.

   The data-layer cases run in demo mode, so they never contact Microsoft.
   app.js remains a browser-controller file in this repository; its async
   ordering guarantees are pinned with structural assertions alongside the
   provider round trips.
   ───────────────────────────────────────────────────────────────────────── */

const assert = require("node:assert/strict");
const fs = require("node:fs");
const path = require("node:path");
const vm = require("node:vm");

const ROOT = path.resolve(__dirname, "..");
const app = fs.readFileSync(path.join(ROOT, "app.js"), "utf8");
const html = fs.readFileSync(path.join(ROOT, "index.html"), "utf8");
const now = new Date();
const TODAY = `${now.getFullYear()}-${String(now.getMonth() + 1).padStart(2, "0")}-${String(now.getDate()).padStart(2, "0")}`;

const ROOM_CONFIG = {
  ROOMS: [
    { id: "pace-room-1", label: "PACE Room 1" },
    { id: "pace-room-2", label: "PACE Room 2" }
  ]
};

const HELPERS = `
  function paceRoomLabelForId(id) { const room = CONFIG.ROOMS.find(r => r.id === id); return room ? room.label : (id || ""); }
  function paceRoomIdForLabel(value) { const raw = String(value ?? "").trim(); if (!raw) return ""; const room = CONFIG.ROOMS.find(r => r.label === raw || r.id === raw); return room ? room.id : raw; }
  const PACE_ROOM_FIELD_CANDIDATES = ["Room", "PACE Room", "Pace Room"];
`;

function makeDemo() {
  const storage = new Map();
  const context = vm.createContext({
    console,
    APP_MODE: "demo",
    DEMO_CONFIG: { storageKey: "paceRoomTrackerDemoData" },
    localStorage: {
      getItem: key => storage.get(key) || null,
      setItem: (key, value) => storage.set(key, value),
      removeItem: key => storage.delete(key)
    },
    crypto: { randomUUID: () => "generated-id" },
    ROSTER: { loaded: false },
    CONFIG: { BEHAVIOR_SPECIALISTS: [], STORAGE_KEYS: { VISIT_CONTEXT: "unused" }, ...ROOM_CONFIG }
  });
  vm.runInContext(HELPERS, context, { filename: "pace-room-helpers" });
  for (const file of ["recent-activity.js", "visit-workflow.js", "demo-data.js", "pace-data.js"]) {
    vm.runInContext(fs.readFileSync(path.join(ROOT, file), "utf8"), context, { filename: file });
  }
  vm.runInContext("this.provider = PACE_DATA", context);
  return { provider: context.provider, raw: () => JSON.parse(storage.get("paceRoomTrackerDemoData")) };
}

function openEntry(id, notes = "") {
  return {
    id,
    paceRoom: "pace-room-1",
    studentId: id,
    studentName: id === "blank" ? "Alex R." : "Jordan M.",
    date: TODAY,
    timeIn: "09:15",
    timeOut: "",
    durationMinutes: null,
    behaviors: ["Needs a Break"],
    interventions: [],
    scmUsed: null,
    notes,
    staffMembers: ["Dana Fielding"],
    cameFromTeacher: "Mrs. Ashford"
  };
}

async function verifyPrestartAndActiveNotesStayOnOneRow() {
  const demo = makeDemo();
  await demo.provider.createVisit(openEntry("blank"));
  await demo.provider.createVisit(openEntry("partial", "Student requested a calm start."));
  await demo.provider.createVisit(openEntry("partial", "Student requested a calm start."));

  let rows = await demo.provider.getVisits({ date: TODAY });
  assert.equal(rows.length, 2, "retrying a start with the same stable id must not create a duplicate");
  assert.equal(rows.find(row => row.id === "blank").Notes, "");
  assert.equal(rows.find(row => row.id === "partial").Notes, "Student requested a calm start.");

  await demo.provider.updateVisit("partial", { date: TODAY, notes: "Added intervention and student response." });
  rows = await demo.provider.getVisits({ date: TODAY });
  assert.equal(rows.length, 2, "active note editing must patch, not create");
  assert.equal(rows.find(row => row.id === "partial").Notes, "Added intervention and student response.");
  assert.equal(rows.find(row => row.id === "blank").Notes, "");

  const finalNotes = "Student returned to class successfully after support.";
  await demo.provider.completeVisit("partial", {
    timeOut: "09:42",
    durationMinutes: 27,
    interventions: ["Calm space"],
    scmUsed: false,
    notes: finalNotes
  });
  rows = await demo.provider.getVisits({ date: TODAY });
  const completed = rows.find(row => row.id === "partial");
  assert.equal(rows.length, 2, "completion must still update the original row");
  assert.equal(completed.Notes, finalNotes, "final completion notes must be saved with the completed visit");
  assert.equal(completed["Time Out"], "09:42");
}

async function verifyProductionNotesPatchIsPartial() {
  const context = vm.createContext({
    console,
    APP_MODE: "production",
    CONFIG: { SITE: "example:/sites/PACE", LISTS: { paceVisits: "IEP_Pace_Visits" } },
    AUTH: { acquireGraphToken: async () => "unused" },
    fetch: async () => { throw new Error("Unexpected network request"); },
    paceRoomLabelForId: id => id
  });
  vm.runInContext(fs.readFileSync(path.join(ROOT, "graph.js"), "utf8"), context, { filename: "graph.js" });
  vm.runInContext("this.graph = GRAPH", context);
  let call;
  let callOptions;
  context.graph.updateMappedListItem = async (list, id, fields, requestOptions) => { call = { list, id, fields }; callOptions = requestOptions; return { id }; };
  const signal = { aborted: false };
  await context.graph.updatePaceVisit("42", { date: TODAY, notes: "Draft notes from the active visit." }, { signal });
  const defined = Object.fromEntries(Object.entries(call.fields).filter(([, value]) => value !== undefined));
  assert.deepEqual(JSON.parse(JSON.stringify(defined)), {
    Date: TODAY,
    Notes: "Draft notes from the active visit."
  }, "active draft persistence must patch the existing row without touching visit timing or completion fields");
  assert.equal(callOptions.signal, signal, "the existing Graph update path must accept cancellation for sign-out safety");
}

async function verifyAutosaveCompletionRace() {
  const scheduleMarker = "///////////////////////////////////////////////////////////////////////////////////////////////\n// Function Name: scheduleActiveNoteSave";
  const resumeMarker = "///////////////////////////////////////////////////////////////////////////////////////////////\n// Function Name: resumeActiveNoteDraft";
  const saveSource = app.slice(app.indexOf("function saveActiveVisitNotes"), app.indexOf(scheduleMarker));
  const flushSource = app.slice(app.indexOf("async function flushActiveNoteSave"), app.indexOf(resumeMarker));
  const pending = [];
  const context = vm.createContext({
    console,
    ACTIVE_NOTE_DRAFTS: new Map(),
    ACTIVE_NOTE_SESSION: { generation: 0 },
    PACE_DATA: {
      updateVisit(id, entry) {
        return new Promise((resolve, reject) => pending.push({ id, entry, resolve, reject }));
      }
    },
    dateOnly: value => String(value || "").slice(0, 10),
    persistActiveNoteDrafts() {},
    updateActiveNoteStatus() {},
    scheduleActiveNoteSave() {}
  });
  vm.runInContext(`${saveSource}\n${flushSource}\nthis.saveDraft = saveActiveVisitNotes; this.flushDraft = flushActiveNoteSave;`, context);
  const draft = {
    value: "first version",
    savedValue: "",
    version: 1,
    status: "unsaved",
    error: "",
    timer: null,
    inFlight: null,
    suspended: false
  };
  context.ACTIVE_NOTE_DRAFTS.set("42", draft);
  const visit = { Date: TODAY };

  const firstSave = context.saveDraft("42", visit);
  draft.value = "latest version before completion";
  draft.version = 2;
  const flush = context.flushDraft("42", visit);
  assert.equal(pending.length, 1, "completion must first wait for the in-flight older autosave");
  pending.shift().resolve();
  await new Promise(resolve => setTimeout(resolve, 0));
  assert.equal(pending.length, 1, "completion must flush the newer text before suspending autosave");
  assert.equal(pending[0].entry.notes, "latest version before completion");
  pending.shift().resolve();
  await firstSave;
  await flush;
  assert.equal(draft.savedValue, "latest version before completion");
  assert.equal(draft.suspended, true, "completion must suspend further active-note autosaves");

  // A later completion edit cannot cause the suspended active-note autosave
  // path to issue a stale PATCH after the final completion write.
  draft.value = "final completion text";
  draft.version = 3;
  await context.saveDraft("42", visit);
  assert.equal(pending.length, 0, "no active-note PATCH may start after completion has suspended the draft");

  const failedDraft = {
    value: "retryable note",
    savedValue: "",
    version: 1,
    status: "unsaved",
    error: "",
    timer: null,
    inFlight: null,
    suspended: false
  };
  context.ACTIVE_NOTE_DRAFTS.set("43", failedDraft);
  const failedSave = context.saveDraft("43", visit);
  assert.equal(pending.length, 1);
  pending.shift().reject(new Error("temporary network failure"));
  await assert.rejects(failedSave, /temporary network failure/);
  assert.equal(failedDraft.status, "failed", "a rejected provider update must never display Saved");
  assert.equal(failedDraft.savedValue, "");
  const retry = context.saveDraft("43", visit);
  assert.equal(pending.length, 1, "failed saves must be explicitly retryable");
  pending.shift().resolve();
  await retry;
  assert.equal(failedDraft.status, "saved");
  assert.equal(failedDraft.savedValue, "retryable note");
}

function verifyPrivacyAndRaceGuards() {
  const persistBody = app.slice(app.indexOf("function persistActiveNoteDrafts()"), app.indexOf("function removeActiveNoteDraft"));
  assert.doesNotMatch(persistBody, /studentName|studentId|Student|email|displayName/i, "the recovery buffer must not store student names or other identity fields");
  assert.match(persistBody, /stored\[id\] = draft\.value/, "recovery entries must be keyed by the existing visit id");

  const saveBody = app.slice(app.indexOf("function saveActiveVisitNotes"), app.indexOf("function scheduleActiveNoteSave"));
  assert.match(saveBody, /const versionAtSend = draft\.version/);
  assert.match(saveBody, /const sessionGeneration = ACTIVE_NOTE_SESSION\.generation/);
  assert.match(saveBody, /if \(sessionGeneration !== ACTIVE_NOTE_SESSION\.generation\) return/);
  assert.match(saveBody, /if \(sessionGeneration !== ACTIVE_NOTE_SESSION\.generation\) throw err/);
  assert.match(saveBody, /draft\.version === versionAtSend && draft\.value === noteAtSend/);
  assert.match(app, /draft\.suspended = true;/, "completion must suspend draft autosave before final completion");
  assert.match(app, /await flushActiveNoteSave\(item\.id, item\)/, "completion must wait for pending autosave work");
  assert.match(app, /function discardActiveNoteDrafts\(\)/);
  assert.match(app, /discardActiveNoteDrafts\(\);\s*AUTH\.logout\(\);/s, "sign-out must clear drafts before logout");

  const signOutIndex = app.indexOf("discardActiveNoteDrafts();\n  AUTH.logout();");
  assert.ok(signOutIndex > 0, "the sign-out handler must invalidate the active-note session");
}

// UI structural guarantees: optional pre-start field, truthful save states,
// version checks, explicit retry, and the existing same-row completion path.
assert.match(html, /id="liveVisitNotesCard"/);
assert.match(html, /id="liveVisitNotesInput"[^>]*placeholder="Begin documenting observations, interventions, and student responses…"/);
assert.match(html, /for="liveVisitNotesInput">Visit Notes<\/label>/);
assert.match(app, /notes: STATE\.notes/);
assert.match(app, /PACE_DATA\.updateVisit\(key, \{\s*date: dateOnly\(visit\?\.Date\),\s*notes: noteAtSend/s);
assert.match(app, /const versionAtSend = draft\.version/);
assert.match(app, /draft\.version === versionAtSend && draft\.value === noteAtSend/);
assert.match(app, /Saving…/);
assert.match(app, /Save failed — tap Retry/);
assert.match(app, /button\.textContent = draft\.status === "failed" \? "Retry"/);
assert.match(app, /await flushActiveNoteSave\(item\.id, item\)/);
assert.match(app, /PACE_DATA\.completeVisit\(STATE\.completionTarget\.id, entry\)/);
assert.match(app, /if \(STATE\.saving\) return; \/\/ never allow a second tap/);
verifyPrivacyAndRaceGuards();

Promise.all([
  verifyPrestartAndActiveNotesStayOnOneRow(),
  verifyProductionNotesPatchIsPartial(),
  verifyAutosaveCompletionRace()
])
  .then(() => console.log("Live visit notes tests passed."))
  .catch(error => {
    console.error(error);
    process.exitCode = 1;
  });
