const assert = require("node:assert/strict");
const fs = require("node:fs");
const path = require("node:path");
const vm = require("node:vm");

const ROOT = path.resolve(__dirname, "..");
const RecentActivity = require(path.join(ROOT, "recent-activity.js"));
const now = new Date();
const localDate = date => `${date.getFullYear()}-${String(date.getMonth() + 1).padStart(2, "0")}-${String(date.getDate()).padStart(2, "0")}`;
const TODAY = localDate(now);
const yesterday = new Date(now.getFullYear(), now.getMonth(), now.getDate() - 1);
const tomorrow = new Date(now.getFullYear(), now.getMonth(), now.getDate() + 1);
const YESTERDAY = localDate(yesterday);
const TOMORROW = localDate(tomorrow);
const appSource = fs.readFileSync(path.join(ROOT, "app.js"), "utf8");
const htmlSource = fs.readFileSync(path.join(ROOT, "index.html"), "utf8");
const cssSource = fs.readFileSync(path.join(ROOT, "styles.css"), "utf8");

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

// Existing screen guidance/states stay intact, and Refresh is a real
// touch control rather than polling. The render block contains none of the
// prohibited visit fields.
assert.match(htmlSource, /id="recentRefreshBtn"[^>]*>Refresh<\/button>/);
assert.match(htmlSource, /Today's visits can be edited here\. For an older entry, contact an administrator\./);
assert.match(appSource, /No PACE activity today yet\./);
assert.match(appSource, /Recent activity could not be loaded\. Try again\./);
assert.match(appSource, /PACE_DATA\.updateVisit\(STATE\.editingVisitId, entry\)/);
const recentRenderSource = appSource.slice(
  appSource.indexOf("async function loadRecentActivity"),
  appSource.indexOf("/* ── DEMO: reset control")
);
for (const prohibited of ["Notes", "Entry ID", "Student ID", "Submitted At", "SCM Used", "Email"]) {
  assert.equal(recentRenderSource.includes(prohibited), false, `${prohibited} entered the Recent render path`);
}

// Responsive contract for both requested landscape iPad sizes: content
// is capped below the viewport width, cards exceed the 44px touch target,
// and the panel scrolls vertically when ten cards exceed available height.
assert.match(cssSource, /\.panel-inner\s*\{[^}]*overflow-y:\s*auto/s);
assert.match(cssSource, /\.recent-list\s*\{[^}]*width:\s*min\(100%,\s*760px\)/s);
assert.match(cssSource, /\.recent-card\s*\{[^}]*min-height:\s*112px/s);
assert.match(cssSource, /\.btn-back, \.btn-text\s*\{[^}]*min-height:\s*44px/s);
for (const [width, height] of [[1024, 768], [1180, 820]]) {
  assert.ok(width - 48 >= 760, `${width}x${height} cannot contain the capped Recent list`);
  assert.ok(height > 44 + 112, `${width}x${height} cannot show header plus a card`);
}

function visit(overrides = {}) {
  return {
    id: "private-id",
    Student: "Alex R.",
    Date: TODAY,
    "Time In": "09:15",
    "Time Out": "09:42",
    Duration: 27,
    Reason: "Needs a Break",
    "Staff Member": "Sharon Morgan",
    "PACE Room": "pace-room-1",
    Notes: "must never render",
    Email: "must-never-render@example.test",
    "Submitted At": `${TODAY}T09:42:00-04:00`,
    ...overrides
  };
}

// Empty state input.
assert.deepEqual(
  RecentActivity.select([], { date: TODAY, room: "pace-room-1" }),
  { roomScoped: false, visits: [] }
);
assert.equal(RecentActivity.isEditableToday(visit(), TODAY), true);
assert.equal(RecentActivity.isEditableToday(visit({ Date: YESTERDAY }), TODAY), false);

// One completed visit, required display values, Needs a Break, and exact
// privacy allow-list (no Notes, IDs, timestamps, or email can leak).
const one = RecentActivity.select([visit()], { date: TODAY, room: "pace-room-1" });
assert.equal(one.roomScoped, true);
assert.equal(one.visits.length, 1);
const oneCard = RecentActivity.cardModel(one.visits[0], { roomScoped: one.roomScoped });
assert.deepEqual(Object.keys(oneCard).sort(), ["displayTime", "duration", "reason", "room", "specialist", "student"]);
assert.deepEqual(oneCard, {
  student: "Alex R.",
  displayTime: "09:42",
  duration: 27,
  reason: "Needs a Break",
  specialist: "Sharon Morgan",
  room: ""
});
assert.deepEqual(RecentActivity.editModel(visit({
  "Intervention Used": "Calm space, Break / reset",
  "SCM Used": "Yes",
  Notes: "Private editor-only note",
  "Teacher Came From": "Mrs. Ashford"
}), { fallbackRoom: "pace-room-2", validRooms: ["pace-room-1", "pace-room-2"] }), {
  itemId: "private-id",
  submissionId: "private-id",
  date: TODAY,
  timeIn: "09:15",
  timeOut: "09:42",
  reasons: ["Needs a Break"],
  supports: ["Calm space", "Break / reset"],
  scmUsed: true,
  notes: "Private editor-only note",
  cameFromTeacher: "Mrs. Ashford",
  specialists: ["Sharon Morgan"],
  student: "Alex R.",
  room: "pace-room-1"
});

// PATCH 010: multiple Behavior Specialists on one visit — cardModel's
// compact display (used by Recent Activity / Currently in PACE) joins up
// to 2 in full, then switches to "First +N"; editModel's array (used to
// hydrate STATE.staffMembers for editing) always returns every name.
const twoSpecialists = visit({ "Staff Member": "Sharon Morgan, Robyn Seiler" });
assert.equal(RecentActivity.cardModel(twoSpecialists).specialist, "Sharon Morgan, Robyn Seiler");
assert.deepEqual(
  RecentActivity.editModel(twoSpecialists, { fallbackRoom: "pace-room-1", validRooms: ["pace-room-1"] }).specialists,
  ["Sharon Morgan", "Robyn Seiler"]
);
const threeSpecialists = visit({ "Staff Member": "Sharon Morgan, Robyn Seiler, Carl Stine" });
assert.equal(RecentActivity.cardModel(threeSpecialists).specialist, "Sharon Morgan +2");
assert.deepEqual(RecentActivity.specialistNames("Sharon Morgan; Robyn Seiler"), ["Sharon Morgan", "Robyn Seiler"], "semicolon-joined legacy values still split");
assert.deepEqual(RecentActivity.specialistNames(""), []);
assert.equal(RecentActivity.formatSpecialistsCompact(""), "");

// Multiple visits sort newest first and stop at 10.
const many = Array.from({ length: 12 }, (_, index) => visit({
  id: String(index),
  Student: `Student ${index}`,
  "Submitted At": `${TODAY}T${String(8 + index).padStart(2, "0")}:00:00-04:00`
}));
const newest = RecentActivity.select(many, { date: TODAY, room: "pace-room-1", limit: 10 });
assert.equal(newest.visits.length, 10);
assert.equal(newest.visits[0].Student, "Student 11");
assert.equal(newest.visits.at(-1).Student, "Student 2");

// When room data exists, Room 1 excludes Room 2. If all room values are
// blank (current live schema behavior), it falls back across today's rows.
const roomScoped = RecentActivity.select([
  visit({ Student: "Room One", "PACE Room": "pace-room-1" }),
  visit({ Student: "Room Two", "PACE Room": "pace-room-2" })
], { date: TODAY, room: "pace-room-1" });
assert.deepEqual(roomScoped.visits.map(item => item.Student), ["Room One"]);

const roomFallback = RecentActivity.select([
  visit({ Student: "First", "PACE Room": "", "Time Out": "09:30" }),
  visit({ Student: "Second", "PACE Room": "", "Time Out": "10:30" })
], { date: TODAY, room: "pace-room-1" });
assert.equal(roomFallback.roomScoped, false);
assert.deepEqual(roomFallback.visits.map(item => item.Student), ["Second", "First"]);

// Optional specialist may be absent. Legacy/open rows missing Time Out are
// excluded without invoking duration parsing or crashing.
const optional = RecentActivity.cardModel(visit({ "Staff Member": "" }));
assert.equal(optional.specialist, "");
const legacy = RecentActivity.select([
  visit({ Student: "Open Legacy", "Time Out": "", Duration: null }),
  visit({ Student: "Completed" })
], { date: TODAY, room: "pace-room-1" });
assert.deepEqual(legacy.visits.map(item => item.Student), ["Completed"]);

// Stored duration can be absent; a valid completed record computes it.
assert.equal(RecentActivity.cardModel(visit({ Duration: null })).duration, 27);

async function verifyDemoProviderUsesLocalStorageOnly() {
  const storage = new Map();
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
    crypto: { randomUUID: () => "demo-visit-1" },
    GRAPH: new Proxy({}, { get: () => () => { microsoftCalls += 1; throw new Error("Microsoft request attempted"); } }),
    ROSTER: { loaded: false },
    CONFIG: { BEHAVIOR_SPECIALISTS: [], ...PACE_ROOM_TEST_CONFIG }
  });

  vm.runInContext(PACE_ROOM_HELPERS_SRC, context, { filename: "pace-room-helpers" });
  for (const file of ["recent-activity.js", "demo-data.js", "pace-data.js"]) {
    vm.runInContext(fs.readFileSync(path.join(ROOT, file), "utf8"), context, { filename: file });
  }
  vm.runInContext("this.testProvider = PACE_DATA", context);

  const before = await context.testProvider.getRecentVisits({ date: TODAY, room: "pace-room-1", limit: 10 });
  assert.equal(before.visits.length, 0);

  await context.testProvider.createVisit({
    id: "demo-visit-1",
    paceRoom: "pace-room-1",
    studentName: "Alex R.",
    date: TODAY,
    timeIn: "09:15",
    timeOut: "09:42",
    durationMinutes: 27,
    behaviors: ["Needs a Break"],
    staffMembers: ["Dana Fielding"],
    timestamp: `${TODAY}T09:42:00-04:00`
  });

  // A fresh provider read models opening/refreshing the screen after save.
  const after = await context.testProvider.getRecentVisits({ date: TODAY, room: "pace-room-1", limit: 10 });
  assert.equal(after.visits.length, 1);
  assert.equal(after.visits[0].Student, "Alex R.");

  await context.testProvider.updateVisit("demo-visit-1", {
    paceRoom: "pace-room-2",
    studentName: "Jordan M.",
    date: TODAY,
    timeIn: "10:00",
    timeOut: "10:35",
    durationMinutes: 35,
    behaviors: ["Peer conflict"],
    interventions: ["Restorative conversation"],
    scmUsed: true,
    notes: "Updated same-day note",
    staffMembers: ["Marcus Webb", "Priya Anand"],
    cameFromTeacher: "Mr. Bellamy"
  });
  const updated = await context.testProvider.getRecentVisits({ date: TODAY, room: "pace-room-2", limit: 10 });
  assert.equal(updated.visits.length, 1);
  assert.deepEqual(JSON.parse(JSON.stringify(updated.visits[0])), {
    id: "demo-visit-1",
    demo: true,
    // ROOM-FIELD-NAME PATCH: "Room" (not "PACE Room") is DemoStorage's own
    // stored key now, matching the confirmed live display name, and it
    // holds the label — see demo-data.js. getRecentVisits() -> getVisits()
    // -> normalizeRoomOnRead() then ADDS "PACE Room" back on as the
    // normalized internal slug every other consumer in this app expects;
    // both keys legitimately coexist on the object this test reads.
    "Room": "PACE Room 2",
    "PACE Room": "pace-room-2",
    Student: "Jordan M.",
    Date: TODAY,
    "Time In": "10:00",
    "Time Out": "10:35",
    Duration: 35,
    Reason: "Peer conflict",
    "Intervention Used": "Restorative conversation",
    "SCM Used": true,
    Notes: "Updated same-day note",
    "Staff Member": "Marcus Webb, Priya Anand",
    "Teacher Came From": "Mr. Bellamy",
    "Submitted By": "Demo Staff",
    "Submitted At": `${TODAY}T09:42:00-04:00`,
    createdAt: updated.visits[0].createdAt,
    modifiedAt: updated.visits[0].modifiedAt
  });
  await assert.rejects(
    () => context.testProvider.updateVisit("demo-visit-1", { date: YESTERDAY }),
    /Only today's visits can be edited/
  );
  assert.ok(storage.has("paceRoomTrackerDemoData"));
  assert.equal(microsoftCalls, 0);
}

async function verifyProductionReadIsDateBounded() {
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

  const paths = [];
  context.testGraph.getSiteId = async () => "site-id";
  context.testGraph.getListId = async () => "list-id";
  context.testGraph.getListSchema = async () => ({ Date: "VisitDate", Student: "Title", "Time In": "TimeIn" });
  context.testGraph._get = async (requestPath, headers) => {
    paths.push({ requestPath, headers });
    return {
      value: [{
        id: "42",
        createdDateTime: `${TODAY}T14:00:00Z`,
        fields: { VisitDate: `${TODAY}T00:00:00Z`, Title: "Alex R.", TimeIn: "09:15" }
      }]
    };
  };

  const rows = await context.testGraph.getPaceVisitsForDateByDisplayName(TODAY);
  assert.equal(paths.length, 1);
  const decodedPath = decodeURIComponent(paths[0].requestPath);
  assert.ok(decodedPath.includes(`fields/VisitDate ge '${TODAY}T00:00:00Z'`));
  assert.ok(decodedPath.includes(`fields/VisitDate lt '${TOMORROW}T00:00:00Z'`));
  assert.match(decodedPath, /\$top=200/);
  assert.equal(paths[0].headers.Prefer, "HonorNonIndexedQueriesWarningMayFailRandomly");
  assert.deepEqual(JSON.parse(JSON.stringify(rows)), [{
    id: "42",
    Created: `${TODAY}T14:00:00Z`,
    Date: `${TODAY}T00:00:00Z`,
    Student: "Alex R.",
    "Time In": "09:15"
  }]);

  let updateCall = null;
  context.testGraph.updateMappedListItem = async (list, id, fields) => {
    updateCall = { list, id, fields };
    return { id };
  };
  await context.testGraph.updatePaceVisit("42", {
    paceRoom: "pace-room-2",
    studentName: "Jordan M.",
    date: TODAY,
    timeIn: "10:00",
    timeOut: "10:35",
    durationMinutes: 35,
    behaviors: ["Peer conflict"],
    interventions: ["Restorative conversation"],
    scmUsed: true,
    notes: "Updated same-day note",
    staffMembers: ["Marcus Webb", "Priya Anand"],
    cameFromTeacher: "Mr. Bellamy"
  });
  assert.deepEqual(JSON.parse(JSON.stringify(updateCall)), {
    list: "IEP_Pace_Visits",
    id: "42",
    fields: {
      // ROOM-FIELD-NAME PATCH: label (not slug) under all three tolerated
      // aliases — see graph.js's updatePaceVisit(). mapFields() itself
      // (which would drop whichever alias isn't a real column) is stubbed
      // out here by the updateMappedListItem override above, so this
      // asserts graph.js's own payload-builder output, before mapping.
      "Room": "PACE Room 2",
      "PACE Room": "PACE Room 2",
      "Pace Room": "PACE Room 2",
      Student: "Jordan M.",
      Date: TODAY,
      "Time In": "10:00",
      "Time Out": "10:35",
      Duration: 35,
      Reason: "Peer conflict",
      "Intervention Used": "Restorative conversation",
      "SCM Used": true,
      Notes: "Updated same-day note",
      "Behavior Specialist": "Marcus Webb, Priya Anand",
      "Teacher Came From": "Mr. Bellamy"
    }
  });
}

Promise.all([
  verifyDemoProviderUsesLocalStorageOnly(),
  verifyProductionReadIsDateBounded()
])
  .then(() => console.log("Recent Activity tests passed."))
  .catch(error => {
    console.error(error);
    process.exitCode = 1;
  });
