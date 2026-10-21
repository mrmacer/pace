/* ─────────────────────────────────────────────────────────────────────────
   NOTES MINIMUM PATCH — completed PACE visits need non-blank Notes (at
   least 1 non-whitespace character; Issue #10 dropped the old 20-character
   minimum). One shared rule (PACE_VISIT_WORKFLOW.validateCompletionNotes)
   enforced in the UI, the save handler, and the data layer. Open live-start
   visits never need Notes; historical records are never rewritten.

   Run with: node tests/patch-notes-minimum.test.js
   ───────────────────────────────────────────────────────────────────────── */

const assert = require("node:assert/strict");
const fs = require("node:fs");
const path = require("node:path");
const vm = require("node:vm");

const ROOT = path.resolve(__dirname, "..");
const Workflow = require(path.join(ROOT, "visit-workflow.js"));
const Corrections = require(path.join(ROOT, "visit-corrections.js"));
const app = fs.readFileSync(path.join(ROOT, "app.js"), "utf8");
const html = fs.readFileSync(path.join(ROOT, "index.html"), "utf8");
const now = new Date();
const TODAY = `${now.getFullYear()}-${String(now.getMonth() + 1).padStart(2, "0")}-${String(now.getDate()).padStart(2, "0")}`;

const MESSAGE = "Add a brief note before submitting this PACE visit.";

/* ── 1. The validator itself ─────────────────────────────────────────── */

const cases = [
  ["empty", "", 0, false],
  ["undefined", undefined, 0, false],
  ["null", null, 0, false],
  ["whitespace-only", "   \n\t  ", 0, false],
  ["single period", ".", 1, true],
  ["single letter", "a", 1, true],
  ["single digit", "7", 1, true],
  ["padding spaces never count", "   ok   ", 2, true],
  ["punctuation counts", "!!!", 3, true],
  ["long note", "Calm space, then back to class.", 26, true]
];
for (const [name, input, count, valid] of cases) {
  const result = Workflow.validateCompletionNotes(input);
  assert.equal(result.count, count, `${name}: count`);
  assert.equal(result.valid, valid, `${name}: valid`);
  assert.equal(result.required, 1, `${name}: required`);
  assert.equal(result.message, valid ? "" : MESSAGE, `${name}: message`);
}
assert.equal(Workflow.NOTES_REQUIREMENT_MESSAGE, MESSAGE);
assert.equal(Workflow.NOTES_MIN_MEANINGFUL, 1);

// The validator measures only: it never mutates or returns rewritten text.
assert.deepEqual(Object.keys(Workflow.validateCompletionNotes("abc")).sort(), ["count", "message", "required", "valid"]);

/* ── 2. Edit/correction rule (Notes may not be blanked) ─────────────── */

{
  const base = { timeIn: "09:00", timeOut: "09:30" };
  assert.equal(Corrections.completedVisitProblem({ ...base, notes: "" }), MESSAGE);
  assert.equal(Corrections.completedVisitProblem({ ...base, notes: "   " }), MESSAGE);
  assert.equal(Corrections.completedVisitProblem({ ...base, notes: "ok" }), "");
  assert.equal(Corrections.completedVisitProblem({ ...base, notes: "a" }), "");
  // Existing time rule unchanged.
  assert.match(Corrections.completedVisitProblem({ timeIn: "09:00", timeOut: "", notes: "ok" }), /both Time In and Time Out/);
}

/* ── 3. Data layer (demo provider; zero Microsoft calls) ─────────────── */

function makeDemo(seed = []) {
  const storage = new Map();
  storage.set("paceRoomTrackerDemoData", JSON.stringify({ paceVisits: seed }));
  let microsoftCalls = 0;
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
    GRAPH: new Proxy({}, { get: () => () => { microsoftCalls += 1; throw new Error("Microsoft request attempted"); } }),
    ROSTER: { loaded: false },
    CONFIG: {
      BEHAVIOR_SPECIALISTS: [], STORAGE_KEYS: { VISIT_CONTEXT: "unused" },
      ROOMS: [{ id: "pace-room-1", label: "PACE Room 1" }, { id: "pace-room-2", label: "PACE Room 2" }]
    }
  });
  vm.runInContext(`
    function paceRoomLabelForId(id) { const r = CONFIG.ROOMS.find(x => x.id === id); return r ? r.label : (id || ""); }
    function paceRoomIdForLabel(v) { const raw = String(v ?? "").trim(); if (!raw) return ""; const r = CONFIG.ROOMS.find(x => x.label === raw || x.id === raw); return r ? r.id : raw; }
    const PACE_ROOM_FIELD_CANDIDATES = ["Room", "PACE Room", "Pace Room"];
  `, context);
  for (const file of ["recent-activity.js", "visit-workflow.js", "visit-corrections.js", "demo-data.js", "pace-data.js"]) {
    vm.runInContext(fs.readFileSync(path.join(ROOT, file), "utf8"), context, { filename: file });
  }
  vm.runInContext("this.provider = PACE_DATA", context);
  return {
    provider: context.provider,
    stored: () => JSON.parse(storage.get("paceRoomTrackerDemoData")).paceVisits,
    raw: () => storage.get("paceRoomTrackerDemoData"),
    microsoftCalls: () => microsoftCalls
  };
}

const entry = (overrides = {}) => ({
  id: "v1", paceRoom: "pace-room-1", studentName: "Alex R.", date: TODAY,
  timeIn: "09:00", timeOut: "", durationMinutes: null, behaviors: ["Needs a Break"],
  interventions: [], scmUsed: null, notes: "", staffMembers: ["Dana Fielding"], cameFromTeacher: "Mrs. Ashford",
  ...overrides
});

async function verifyDemoDataLayer() {
  // Open visit without Notes is fine (Start Visit path).
  const open = makeDemo();
  await open.provider.createVisit(entry());
  assert.equal(open.stored().length, 1, "open visit without Notes still saves");
  assert.equal(open.stored()[0].Notes, "");

  // Mark Complete: rejected for blank Notes, row untouched.
  const mc = makeDemo();
  await mc.provider.createVisit(entry());
  const before = mc.raw();
  for (const bad of ["", "   ", " \n\t "]) {
    await assert.rejects(
      () => mc.provider.completeVisit("v1", { timeOut: "09:30", durationMinutes: 30, interventions: ["Calm space"], scmUsed: true, notes: bad }),
      /Add a brief note/,
      `Mark Complete must reject ${JSON.stringify(bad)}`
    );
    assert.equal(mc.raw(), before, "a rejected completion must not touch the stored row");
  }
  const original = "  ok  ";
  await mc.provider.completeVisit("v1", { timeOut: "09:30", durationMinutes: 30, interventions: ["Calm space"], scmUsed: true, notes: original });
  assert.equal(mc.stored()[0]["Time Out"], "09:30");
  assert.equal(mc.stored()[0].Notes, original, "Notes text is saved exactly as the caller supplied it — never rewritten");

  // After-the-fact entry (already has a Time Out).
  const aft = makeDemo();
  for (const bad of ["", " \n ", "\t"]) {
    await assert.rejects(
      () => aft.provider.createVisit(entry({ id: "x", timeOut: "09:30", durationMinutes: 30, notes: bad })),
      /Add a brief note/,
      `after-the-fact entry must reject ${JSON.stringify(bad)}`
    );
  }
  assert.equal(aft.stored().length, 0, "no rejected after-the-fact entry was stored");
  await aft.provider.createVisit(entry({ id: "ok1", timeOut: "09:30", durationMinutes: 30, notes: "a" }));
  await aft.provider.createVisit(entry({ id: "ok45", timeOut: "10:30", durationMinutes: 30, notes: "a".repeat(45) }));
  assert.deepEqual(aft.stored().map(v => v.id), ["ok1", "ok45"]);

  // Editing a completed visit: short historical Notes are fine as-is; only
  // blanking them is refused.
  const historical = {
    id: "old", demo: true, Room: "PACE Room 1", Student: "Jordan M.", Date: TODAY,
    "Time In": "08:00", "Time Out": "08:20", Duration: 20, Reason: "Needs a Break",
    "Intervention Used": "Calm space", "SCM Used": false, Notes: "ok", "Staff Member": "Dana Fielding", "Teacher Came From": "Mrs. Ashford"
  };
  const edit = makeDemo([historical]);
  const seeded = edit.raw();
  const rows = await edit.provider.getVisits({ date: TODAY });
  assert.equal(rows[0].Notes, "ok", "historical short Notes remain visible");
  assert.equal(edit.raw(), seeded, "reading historical records never rewrites them");

  const snapshot = {
    paceRoom: "pace-room-1", studentName: "Jordan M.", date: TODAY, timeIn: "08:00", timeOut: "08:20",
    behaviors: ["Needs a Break"], interventions: ["Calm space"], scmUsed: false, notes: "ok",
    staffMembers: ["Dana Fielding"], cameFromTeacher: "Mrs. Ashford"
  };
  // Blanking the note: rejected, nothing written.
  await assert.rejects(
    () => edit.provider.updateVisit("old", { ...snapshot, original: snapshot, notes: "   " }),
    /Add a brief note/
  );
  assert.equal(edit.raw(), seeded, "rejected edit wrote nothing");
  // Edit something else and keep the short note: saved.
  await edit.provider.updateVisit("old", { ...snapshot, original: snapshot, scmUsed: true });
  assert.equal(edit.stored()[0]["SCM Used"], true);
  assert.equal(edit.stored()[0].Notes, "ok");
  assert.equal(edit.stored().length, 1, "edit updates in place");

  for (const instance of [open, mc, aft, edit]) assert.equal(instance.microsoftCalls(), 0, "demo mode never calls Microsoft");
}

/* ── 4. UI + save-handler structure ──────────────────────────────────── */

{
  // Issue #10: the character counter is gone (meaningless at a 1-char minimum).
  assert.doesNotMatch(html, /notesCounter/, "no character counter on the Notes screen");
  assert.doesNotMatch(app, /updateNotesCounter|meaningful characters/);

  // Requirement message is shown inline, user stays on Notes, draft preserved.
  const next = app.slice(app.indexOf('document.getElementById("notesNextBtn").addEventListener'));
  const nextBody = next.slice(0, next.indexOf("\n});"));
  assert.match(nextBody, /validateCompletionNotes\(raw\)\.valid/);
  assert.match(nextBody, /showNotesRequirementError\(\);[\s\S]*return;/, "blocked Next returns before navigating");
  assert.ok(nextBody.indexOf("return;") < nextBody.indexOf('nav("confirm"'), "no navigation to Confirm on failure");
  assert.doesNotMatch(nextBody.slice(0, nextBody.indexOf("return;")), /noteText"\)\.value\s*=|STATE\.notes\s*=/, "failure path never clears or rewrites the draft");

  // Save handler: defensive check precedes saving state and every write call.
  const handlerStart = app.indexOf('document.getElementById("saveEntryBtn").addEventListener("click"');
  const check = app.indexOf("validateCompletionNotes(STATE.notes).valid", handlerStart);
  const saving = app.indexOf("STATE.saving = true;", handlerStart);
  assert.ok(check > handlerStart && check < saving, "save handler re-checks Notes before entering the saving state");
  for (const call of ["PACE_DATA.createVisit(entry)", "PACE_DATA.completeVisit(", "PACE_DATA.updateVisit("]) {
    assert.ok(check < app.indexOf(call, handlerStart), `Notes check precedes ${call}`);
  }
  const block = app.slice(check, saving);
  assert.match(block, /nav\("notes", "back"\)/, "bounces back to Notes");
  assert.match(block, /return;/);
  assert.doesNotMatch(block, /resetTrip\(\)|STATE\.(reasons|supports|scmUsed|student|staffMembers|cameFromTeacher|date|timeIn|timeOut|notes)\s*=[^=]/, "bounce preserves everything entered");

  // Start Visit path: unreachable from Notes, no Notes validation there.
  const live = app.slice(app.indexOf("async function saveLiveVisit()"));
  assert.doesNotMatch(live.slice(0, live.search(/\r?\n\}\r?\n/)), /Notes|notes/);

  // Single definition of the rule: no second hand-rolled length check in app.js.
  assert.doesNotMatch(app, /notes.{0,40}length\s*[<>]=?\s*\d/i);
  assert.match(fs.readFileSync(path.join(ROOT, "visit-workflow.js"), "utf8"), /NOTES_MIN_MEANINGFUL: 1,/);
}

/* ── 5. Safety: nothing outside this app / its data rules was touched ── */

{
  const graph = fs.readFileSync(path.join(ROOT, "graph.js"), "utf8");
  assert.doesNotMatch(graph, /NOTES_MIN|validateCompletionNotes/, "graph.js (schema/permissions layer) is untouched by the rule");
}

verifyDemoDataLayer()
  .then(() => console.log("Notes-minimum patch tests passed."))
  .catch(error => { console.error(error); process.exitCode = 1; });
