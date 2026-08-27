/* ─────────────────────────────────────────────────────────────────────────
   PATCH 010 — multiple Behavior Specialists + expanded Notes

   Covers what recent-activity.test.js / live-visit.test.js don't already:
   structural UI assertions (multi-select markup, no Notes maxlength) and a
   full-length Notes round-trip through both the demo and production write
   paths. See README.md "PATCH 010" for the full writeup.
   ───────────────────────────────────────────────────────────────────────── */

const assert = require("node:assert/strict");
const fs = require("node:fs");
const path = require("node:path");
const vm = require("node:vm");

const ROOT = path.resolve(__dirname, "..");
const html = fs.readFileSync(path.join(ROOT, "index.html"), "utf8");
const app = fs.readFileSync(path.join(ROOT, "app.js"), "utf8");
const now = new Date();
const TODAY = `${now.getFullYear()}-${String(now.getMonth() + 1).padStart(2, "0")}-${String(now.getDate()).padStart(2, "0")}`;

// ROOM-FIELD-NAME PATCH: see the identical block in live-visit.test.js for
// why this duplicates config.js's room helpers instead of loading the
// real file.
const PACE_ROOM_TEST_CONFIG = {
  ROOMS: [
    { id: "pace-room-1", label: "PACE Room 1", hallway: "Yellow Hall", color: "yellow" },
    { id: "pace-room-2", label: "PACE Room 2", hallway: "Green Hall", color: "green" }
  ]
};
const PACE_ROOM_HELPERS_SRC = `
  function paceRoomLabelForId(id) {
    const room = CONFIG.ROOMS.find(r => r.id === id);
    return room ? room.label : (id || "");
  }
  function paceRoomIdForLabel(value) {
    const raw = String(value ?? "").trim();
    if (!raw) return "";
    const room = CONFIG.ROOMS.find(r => r.label === raw || r.id === raw);
    return room ? room.id : raw;
  }
  const PACE_ROOM_FIELD_CANDIDATES = ["Room", "PACE Room", "Pace Room"];
`;

/* ── Multi-select UX (structural) ─────────────────────────────────────── */

// Tapping a card no longer navigates immediately (that was the entire
// single-select-and-advance bug this patch removes) — selectSpecialist()
// must not exist any more, only the toggle version.
assert.equal(app.includes("function selectSpecialist("), false, "old single-select-and-advance handler must be removed");
assert.match(app, /function toggleSpecialist\(name\)/);
assert.match(app, /STATE\.staffMembers\.push\(name\)/);
assert.match(app, /STATE\.staffMembers\.splice\(idx, 1\)/);

// A dedicated Continue button, disabled at zero selections — the required
// "at least one" gate, enforced independently of the live-visit save
// validation covered in live-visit.test.js.
assert.match(html, /id="specialistContinueBtn"[^>]*disabled[^>]*>CONTINUE</);
assert.match(app, /document\.getElementById\("specialistContinueBtn"\)\.disabled = count === 0/);
assert.match(app, /if \(STATE\.staffMembers\.length === 0\) return;/);

// STATE.staffMember (singular) must be fully gone from actual CODE — one
// competing source of truth only (STATE.staffMembers), per instruction.
// Comment-only historical mentions (e.g. "was STATE.staffMember before
// this patch") are fine and expected; only non-comment lines are checked.
const codeLines = app.split("\n").filter(line => !/^\s*(\/\/|\*)/.test(line));
const staleRefs = codeLines.filter(line => /STATE\.staffMember\b(?!s)/.test(line));
assert.deepEqual(staleRefs, [], "STATE.staffMember (singular) must not remain in any non-comment line of app.js");

// Confirm/Exit screens render one name per line via <br>, not a raw array
// and not silently collapsed to one name.
assert.match(app, /STATE\.staffMembers\.map\(escHtml\)\.join\("<br>"\)/);

/* ── Notes: no arbitrary limit ───────────────────────────────────────── */

assert.equal(/id="noteText"[^>]*maxlength/.test(html), false, "Notes textarea must not carry a frontend maxlength");
assert.match(html, /id="noteText" class="note-textarea hidden" rows="6"/); // still a real, appropriately-sized textarea
assert.equal(/STATE\.notes[^;]*\.(slice|substring|substr)\(/.test(app), false, "Notes must never be truncated in JS");

/* ── Full-length, multi-paragraph Notes round-trips unchanged ──────────── */

const LONG_NOTE = [
  "Line one of the observation, describing in detail what led up to the visit, including the classroom activity in progress and the specific antecedent that was observed by staff immediately beforehand.",
  "",
  "Line two, a second paragraph covering the intervention attempted and the student's response to it over several sentences, so this fixture reliably exceeds five hundred characters in total length for this test.",
  "",
  "Third paragraph: the follow-up plan, any family contact made, and a closing note summarizing the outcome once the student returned to class."
].join("\n");
assert.ok(LONG_NOTE.length > 500, "fixture note must actually exceed 500 characters");

async function verifyLongNoteRoundTrip() {
  // Demo path: DemoStorage.createVisit -> stored Notes field.
  const storage = new Map();
  const demoContext = vm.createContext({
    console,
    APP_MODE: "demo",
    DEMO_CONFIG: { storageKey: "paceRoomTrackerDemoData" },
    localStorage: {
      getItem: key => storage.get(key) || null,
      setItem: (key, value) => storage.set(key, value)
    },
    crypto: { randomUUID: () => "note-test-1" },
    CONFIG: { ...PACE_ROOM_TEST_CONFIG }
  });
  vm.runInContext(PACE_ROOM_HELPERS_SRC, demoContext, { filename: "pace-room-helpers" });
  vm.runInContext(fs.readFileSync(path.join(ROOT, "demo-data.js"), "utf8"), demoContext, { filename: "demo-data.js" });
  vm.runInContext("this.demoStorage = DemoStorage", demoContext);
  const demoVisit = demoContext.demoStorage.createVisit({
    id: "note-test-1",
    paceRoom: "pace-room-1",
    studentName: "Alex R.",
    date: TODAY,
    timeIn: "09:15",
    timeOut: "09:42",
    behaviors: ["Needs a Break"],
    staffMembers: ["Dana Fielding"],
    notes: LONG_NOTE
  });
  assert.equal(demoVisit.Notes, LONG_NOTE, "DemoStorage must preserve the full note, unmodified");
  assert.equal(demoVisit.Notes.length, LONG_NOTE.length);

  // Production mapping path: GRAPH.savePaceVisit's payload builder — never
  // actually reaches the network here (createMappedListItem/_get are
  // stubbed), only proves the payload itself carries the full text.
  const graphContext = vm.createContext({
    console,
    APP_MODE: "production",
    CONFIG: { SITE: "example.sharepoint.test:/sites/PACE", LISTS: { paceVisits: "IEP_Pace_Visits" }, ...PACE_ROOM_TEST_CONFIG },
    AUTH: { acquireGraphToken: async () => "unused" },
    fetch: async () => { throw new Error("Unexpected network request"); }
  });
  vm.runInContext(PACE_ROOM_HELPERS_SRC, graphContext, { filename: "pace-room-helpers" });
  vm.runInContext(fs.readFileSync(path.join(ROOT, "graph.js"), "utf8"), graphContext, { filename: "graph.js" });
  vm.runInContext("this.testGraph = GRAPH", graphContext);
  let mappedFields = null;
  graphContext.testGraph.createMappedListItem = async (list, fields) => { mappedFields = fields; return { id: "42" }; };
  await graphContext.testGraph.savePaceVisit({
    id: "note-test-1",
    paceRoom: "pace-room-1",
    studentName: "Alex R.",
    date: TODAY,
    timeIn: "09:15",
    timeOut: "09:42",
    behaviors: ["Needs a Break"],
    staffMembers: ["Dana Fielding"],
    notes: LONG_NOTE
  });
  assert.equal(mappedFields.Notes, LONG_NOTE, "the production save payload must carry the full note, unmodified, into mapFields()");
}

verifyLongNoteRoundTrip()
  .then(() => console.log("PATCH 010 (specialists + notes) tests passed."))
  .catch(error => {
    console.error(error);
    process.exitCode = 1;
  });
