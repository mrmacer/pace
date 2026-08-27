/* ─────────────────────────────────────────────────────────────────────────
   ROOM-FIELD-NAME PATCH — focused regression coverage

   Live-visit.test.js and recent-activity.test.js incidentally exercise the
   room slug<->label conversion, but neither one specifically pins down the
   thing THIS patch changed: the confirmed live SharePoint display name is
   "Room", not "PACE Room", and the write/read boundary now tolerates all
   three of Room / PACE Room / Pace Room. This file is the dedicated
   coverage for that behavior — see README.md's "ROOM-FIELD-NAME PATCH"
   section for the full writeup.
   ───────────────────────────────────────────────────────────────────────── */

const assert = require("node:assert/strict");
const fs = require("node:fs");
const path = require("node:path");
const vm = require("node:vm");

const ROOT = path.resolve(__dirname, "..");
const now = new Date();
const TODAY = `${now.getFullYear()}-${String(now.getMonth() + 1).padStart(2, "0")}-${String(now.getDate()).padStart(2, "0")}`;

// Kept an exact copy of config.js's real implementation — see the
// identical block in the other test files for why these tests don't load
// the real config.js (it needs `window.location` and pulls in unrelated
// production config).
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

/* ── Write side: savePaceVisit() sends the label under all three aliases ── */

async function verifyWriteSideAliases() {
  const context = vm.createContext({
    console,
    APP_MODE: "production",
    CONFIG: { SITE: "example.sharepoint.test:/sites/PACE", LISTS: { paceVisits: "IEP_Pace_Visits" }, ...PACE_ROOM_TEST_CONFIG },
    AUTH: { acquireGraphToken: async () => "unused" },
    fetch: async () => { throw new Error("Unexpected network request"); }
  });
  vm.runInContext(PACE_ROOM_HELPERS_SRC, context, { filename: "pace-room-helpers" });
  vm.runInContext(fs.readFileSync(path.join(ROOT, "graph.js"), "utf8"), context, { filename: "graph.js" });
  vm.runInContext("this.testGraph = GRAPH", context);

  for (const [roomId, expectedLabel] of [["pace-room-1", "PACE Room 1"], ["pace-room-2", "PACE Room 2"]]) {
    let mappedFields = null;
    context.testGraph.createMappedListItem = async (list, fields) => { mappedFields = fields; return { id: "new-item" }; };
    await context.testGraph.savePaceVisit({
      id: `entry-${roomId}`,
      paceRoom: roomId,
      studentName: "Alex R.",
      date: TODAY,
      timeIn: "09:15",
      timeOut: "09:42",
      behaviors: ["Needs a Break"],
      staffMembers: ["Dana Fielding"]
    });
    assert.equal(mappedFields.Room, expectedLabel, `${roomId}: "Room" alias must carry the label`);
    assert.equal(mappedFields["PACE Room"], expectedLabel, `${roomId}: "PACE Room" alias must carry the same label`);
    assert.equal(mappedFields["Pace Room"], expectedLabel, `${roomId}: "Pace Room" alias must carry the same label`);
  }
}

/* ── mapFields() actually keeps only the real column, drops the others ─── */

async function verifyMapFieldsKeepsOnlyRealColumn() {
  const context = vm.createContext({
    console,
    APP_MODE: "production",
    CONFIG: { SITE: "example.sharepoint.test:/sites/PACE", LISTS: { paceVisits: "IEP_Pace_Visits" }, ...PACE_ROOM_TEST_CONFIG },
    AUTH: { acquireGraphToken: async () => "unused" },
    fetch: async () => { throw new Error("Unexpected network request"); }
  });
  vm.runInContext(PACE_ROOM_HELPERS_SRC, context, { filename: "pace-room-helpers" });
  vm.runInContext(fs.readFileSync(path.join(ROOT, "graph.js"), "utf8"), context, { filename: "graph.js" });
  vm.runInContext("this.testGraph = GRAPH", context);

  // Simulate the confirmed live reality: a schema where the column is
  // literally named "Room" (internal name "Room0"), and neither
  // "PACE Room" nor "Pace Room" exist at all.
  context.testGraph.getListSchema = async () => ({
    Room: "Room0", "Entry ID": "EntryID0", Student: "Title", Date: "Date0",
    "Time In": "TimeIn0", "Time Out": "TimeOut0"
  });
  let createdFields = null;
  context.testGraph.createListItem = async (list, fields) => { createdFields = fields; return { id: "new-item" }; };

  await context.testGraph.savePaceVisit({
    id: "entry-1", paceRoom: "pace-room-1", studentName: "Alex R.",
    date: TODAY, timeIn: "09:15", timeOut: "09:42", behaviors: ["Needs a Break"], staffMembers: ["Dana Fielding"]
  });

  assert.deepEqual(Object.keys(createdFields), ["EntryID0", "Room0", "Title", "Date0", "TimeIn0", "TimeOut0"]);
  assert.equal(createdFields.Room0, "PACE Room 1", "only the real \"Room\" column receives the write");
}

/* ── Read side: readPaceRoomValue()/normalizeRoomOnRead() via pace-data.js's
   getVisits(), in demo mode, across all three possible raw key spellings,
   plus an older record with none of them ────────────────────────────────── */

async function verifyReadSideAliasResolution() {
  const storage = new Map();
  const context = vm.createContext({
    console,
    APP_MODE: "demo",
    DEMO_CONFIG: { storageKey: "paceRoomTrackerDemoData" },
    localStorage: {
      getItem: key => storage.get(key) || null,
      setItem: (key, value) => storage.set(key, value)
    },
    ROSTER: { loaded: false },
    CONFIG: { BEHAVIOR_SPECIALISTS: [], ...PACE_ROOM_TEST_CONFIG }
  });
  vm.runInContext(PACE_ROOM_HELPERS_SRC, context, { filename: "pace-room-helpers" });
  for (const file of ["recent-activity.js", "demo-data.js", "pace-data.js"]) {
    vm.runInContext(fs.readFileSync(path.join(ROOT, file), "utf8"), context, { filename: file });
  }
  vm.runInContext("this.provider = PACE_DATA", context);

  // Seed localStorage directly (bypassing createVisit) so each row can
  // carry a different, deliberately-chosen raw key spelling.
  storage.set("paceRoomTrackerDemoData", JSON.stringify({
    paceVisits: [
      { id: "a", Date: TODAY, Student: "Row A", "Room": "PACE Room 1" },
      { id: "b", Date: TODAY, Student: "Row B", "PACE Room": "PACE Room 2" },
      { id: "c", Date: TODAY, Student: "Row C", "Pace Room": "PACE Room 1" },
      // Older record: no room value under any candidate key at all —
      // must come back blank, never inferred/guessed/backfilled.
      { id: "d", Date: TODAY, Student: "Row D" }
    ]
  }));

  const rows = await context.provider.getVisits({ date: TODAY });
  const byId = Object.fromEntries(rows.map(r => [r.id, r["PACE Room"]]));
  assert.equal(byId.a, "pace-room-1", '"Room" key must resolve to the internal slug');
  assert.equal(byId.b, "pace-room-2", '"PACE Room" key must still resolve correctly');
  assert.equal(byId.c, "pace-room-1", '"Pace Room" key must still resolve correctly');
  assert.equal(byId.d, "", "an older record with no room value under any alias must stay blank, not inferred");
}

/* ── Full round trip through PACE_DATA.createVisit() (demo) for both rooms ── */

async function verifyDemoCreateRoundTrip() {
  const storage = new Map();
  const context = vm.createContext({
    console,
    APP_MODE: "demo",
    DEMO_CONFIG: { storageKey: "paceRoomTrackerDemoData" },
    localStorage: {
      getItem: key => storage.get(key) || null,
      setItem: (key, value) => storage.set(key, value)
    },
    crypto: { randomUUID: () => "round-trip-1" },
    ROSTER: { loaded: false },
    CONFIG: { BEHAVIOR_SPECIALISTS: [], ...PACE_ROOM_TEST_CONFIG }
  });
  vm.runInContext(PACE_ROOM_HELPERS_SRC, context, { filename: "pace-room-helpers" });
  for (const file of ["recent-activity.js", "demo-data.js", "pace-data.js"]) {
    vm.runInContext(fs.readFileSync(path.join(ROOT, file), "utf8"), context, { filename: file });
  }
  vm.runInContext("this.provider = PACE_DATA", context);

  for (const [roomId, expectedLabel] of [["pace-room-1", "PACE Room 1"], ["pace-room-2", "PACE Room 2"]]) {
    context.crypto.randomUUID = () => `round-trip-${roomId}`;
    await context.provider.createVisit({
      id: `round-trip-${roomId}`, paceRoom: roomId, studentName: "Alex R.", date: TODAY,
      timeIn: "09:15", timeOut: "", behaviors: ["Needs a Break"], staffMembers: ["Dana Fielding"]
    });
  }
  const stored = JSON.parse(storage.get("paceRoomTrackerDemoData")).paceVisits;
  const raw1 = stored.find(v => v.id === "round-trip-pace-room-1");
  const raw2 = stored.find(v => v.id === "round-trip-pace-room-2");
  assert.equal(raw1.Room, "PACE Room 1", "DemoStorage's own key must be \"Room\", holding the label");
  assert.equal(raw2.Room, "PACE Room 2");
  assert.equal(raw1["PACE Room"], undefined, "DemoStorage must not also write the old \"PACE Room\" key");

  const rows = await context.provider.getVisits({ date: TODAY });
  assert.equal(rows.find(r => r.id === "round-trip-pace-room-1")["PACE Room"], "pace-room-1");
  assert.equal(rows.find(r => r.id === "round-trip-pace-room-2")["PACE Room"], "pace-room-2");
}

Promise.all([
  verifyWriteSideAliases(),
  verifyMapFieldsKeepsOnlyRealColumn(),
  verifyReadSideAliasResolution(),
  verifyDemoCreateRoundTrip()
])
  .then(() => console.log("Room field-name compatibility tests passed."))
  .catch(error => {
    console.error(error);
    process.exitCode = 1;
  });
