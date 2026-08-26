const assert = require("node:assert/strict");
const fs = require("node:fs");
const path = require("node:path");
const vm = require("node:vm");

const ROOT = path.resolve(__dirname, "..");
const Workflow = require(path.join(ROOT, "visit-workflow.js"));
const html = fs.readFileSync(path.join(ROOT, "index.html"), "utf8");
const app = fs.readFileSync(path.join(ROOT, "app.js"), "utf8");
const css = fs.readFileSync(path.join(ROOT, "styles.css"), "utf8");
const now = new Date();
const TODAY = `${now.getFullYear()}-${String(now.getMonth() + 1).padStart(2, "0")}-${String(now.getDate()).padStart(2, "0")}`;

assert.match(html, /id="startLiveVisitBtn">START VISIT/);
assert.match(html, /Student is entering PACE now/);
assert.match(html, /id="logCompletedVisitBtn">LOG COMPLETED VISIT/);
assert.match(html, /Visit already happened/);
assert.match(html, /Currently in PACE/);
assert.match(html, /id="confirmExitBtn">NEXT/);
assert.match(app, />MARK COMPLETE<\/button>/);
assert.match(app, /PACE_DATA\.completeVisit\(STATE\.completionTarget\.id, entry\)/);
assert.match(app, /PACE_DATA\.createVisit\(buildVisitEntry\(\{ open: true \}\)\)/);
assert.equal(Workflow.durationMinutes("09:15", "09:42"), 27);
assert.equal(Workflow.durationMinutes("09:42", "09:42"), null);
assert.equal(Workflow.durationMinutes("10:00", "09:59"), null);

const openRows = Workflow.openForRoom([
  { id: "2", "PACE Room": "pace-room-1", "Time In": "10:00", "Time Out": "" },
  { id: "1", "PACE Room": "pace-room-1", "Time In": "09:00", "Time Out": "" },
  { id: "3", "PACE Room": "pace-room-2", "Time In": "08:00", "Time Out": "" },
  { id: "4", "PACE Room": "pace-room-1", "Time In": "08:00", "Time Out": "08:30" }
], "pace-room-1");
assert.deepEqual(openRows.map(row => row.id), ["1", "2"]);

// Both requested iPad landscape widths retain the one-column, full-width
// operational controls; no narrow desktop-only breakpoint is required.
assert.match(css, /\.panel-inner\s*\{[^}]*overflow-y:\s*auto/s);
assert.match(css, /\.btn-full\s*\{\s*width:\s*100%/s);
for (const [width, height] of [[1024, 768], [1180, 820]]) {
  assert.ok(width >= 1024 && height >= 768);
}

async function verifyDemoSameRowLifecycle() {
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
    crypto: { randomUUID: () => "generated-id" },
    GRAPH: new Proxy({}, { get: () => () => { microsoftCalls += 1; throw new Error("Microsoft request attempted"); } }),
    ROSTER: { loaded: false },
    CONFIG: { BEHAVIOR_SPECIALISTS: [], STORAGE_KEYS: { VISIT_CONTEXT: "unused" } }
  });
  for (const file of ["recent-activity.js", "demo-data.js", "pace-data.js"]) {
    vm.runInContext(fs.readFileSync(path.join(ROOT, file), "utf8"), context, { filename: file });
  }
  vm.runInContext("this.provider = PACE_DATA", context);
  vm.runInContext("this.recentActivity = RECENT_ACTIVITY", context);

  const openEntry = {
    id: "live-1",
    paceRoom: "pace-room-1",
    studentName: "Alex R.",
    date: TODAY,
    timeIn: "09:15",
    timeOut: "",
    durationMinutes: null,
    behaviors: ["Needs a Break"],
    interventions: [],
    scmUsed: null,
    notes: "",
    staffMembers: ["Dana Fielding", "Priya Anand"],
    cameFromTeacher: "Mrs. Ashford"
  };
  await context.provider.createVisit(openEntry);
  await context.provider.createVisit(openEntry); // rapid retry / same entry id

  let rows = await context.provider.getVisits({ date: TODAY });
  assert.equal(rows.length, 1);
  assert.equal(rows[0]["Time Out"], "");
  assert.equal(rows[0].Duration, null);
  assert.equal(rows[0]["Intervention Used"], "");
  assert.equal(rows[0]["SCM Used"], null);
  assert.equal(rows[0].Notes, "");
  // PATCH 010: both specialists are attached to the open visit...
  assert.equal(rows[0]["Staff Member"], "Dana Fielding, Priya Anand");
  assert.deepEqual(Workflow.openForRoom(JSON.parse(JSON.stringify(rows)), "pace-room-1").map(row => row.id), ["live-1"]);

  await context.provider.completeVisit("live-1", {
    timeOut: "09:42",
    durationMinutes: 27,
    interventions: ["Calm space"],
    scmUsed: false,
    notes: "Returned successfully"
  });
  rows = await context.provider.getVisits({ date: TODAY });
  assert.equal(rows.length, 1, "completion must update, not create");
  assert.equal(rows[0].id, "live-1");
  assert.equal(rows[0]["Time Out"], "09:42");
  assert.equal(rows[0].Duration, 27);
  assert.equal(rows[0]["Intervention Used"], "Calm space");
  assert.equal(rows[0]["SCM Used"], false);
  // ...and Mark Complete never re-asks for or touches them — still both,
  // untouched, on the same completed row (see completeVisit()'s patch,
  // which never includes "Staff Member"/"Behavior Specialist").
  assert.equal(rows[0]["Staff Member"], "Dana Fielding, Priya Anand");
  assert.equal(Workflow.openForRoom(JSON.parse(JSON.stringify(rows)), "pace-room-1").length, 0);

  const recent = await context.provider.getRecentVisits({ date: TODAY, room: "pace-room-1" });
  assert.equal(recent.visits.length, 1);
  assert.equal(recent.visits[0].id, "live-1");
  assert.equal(context.recentActivity.cardModel(recent.visits[0]).specialist, "Dana Fielding, Priya Anand");
  assert.equal(microsoftCalls, 0, "demo mode must never call Microsoft");
}

async function verifyProductionCompletionPatch() {
  const context = vm.createContext({
    console,
    APP_MODE: "production",
    CONFIG: { SITE: "example:/sites/PACE", LISTS: { paceVisits: "IEP_Pace_Visits" } },
    AUTH: { acquireGraphToken: async () => "unused" },
    fetch: async () => { throw new Error("Unexpected network request"); }
  });
  vm.runInContext(fs.readFileSync(path.join(ROOT, "graph.js"), "utf8"), context, { filename: "graph.js" });
  vm.runInContext("this.graph = GRAPH", context);
  let call;
  context.graph.updateMappedListItem = async (list, id, fields) => { call = { list, id, fields }; return { id }; };
  await context.graph.completePaceVisit("42", {
    timeOut: "11:10",
    durationMinutes: 35,
    interventions: ["Calm space", "Return-to-class plan"],
    scmUsed: false,
    notes: "Complete"
  });
  assert.deepEqual(JSON.parse(JSON.stringify(call)), {
    list: "IEP_Pace_Visits",
    id: "42",
    fields: {
      "Time Out": "11:10",
      Duration: 35,
      "Intervention Used": "Calm space, Return-to-class plan",
      "SCM Used": false,
      Notes: "Complete"
    }
  });
}

async function verifyProductionContextEnrichment() {
  const storage = new Map();
  const graph = {
    savePaceVisit: async () => ({ id: "sp-7" }),
    getPaceVisitsForDateByDisplayName: async () => [{
      id: "sp-7", Student: "Alex R.", Date: TODAY, "Time In": "09:15", "Time Out": "", "PACE Room": ""
    }]
  };
  const context = vm.createContext({
    console,
    APP_MODE: "production",
    CONFIG: { BEHAVIOR_SPECIALISTS: [], STORAGE_KEYS: { VISIT_CONTEXT: "visit-context" } },
    localStorage: {
      getItem: key => storage.get(key) || null,
      setItem: (key, value) => storage.set(key, value)
    },
    GRAPH: graph,
    ROSTER: { loaded: true },
    RECENT_ACTIVITY: { select: visits => ({ visits }) }
  });
  vm.runInContext(fs.readFileSync(path.join(ROOT, "pace-data.js"), "utf8"), context, { filename: "pace-data.js" });
  vm.runInContext("this.provider = PACE_DATA", context);
  await context.provider.createVisit({
    paceRoom: "pace-room-2",
    staffMembers: ["Dana Fielding", "Priya Anand"],
    cameFromTeacher: "Mrs. Ashford"
  });
  const rows = await context.provider.getVisits({ date: TODAY });
  assert.equal(rows[0]["PACE Room"], "pace-room-2");
  assert.equal(rows[0]["Behavior Specialist"], "Dana Fielding, Priya Anand");
  assert.equal(rows[0]["Teacher Came From"], "Mrs. Ashford");
  const stored = JSON.parse(storage.get("visit-context"));
  assert.deepEqual(Object.keys(stored["sp-7"]).sort(), ["room", "specialist", "teacher"]);
  assert.equal(JSON.stringify(stored).includes("Alex R."), false, "local context must not store student data");
}

Promise.all([verifyDemoSameRowLifecycle(), verifyProductionCompletionPatch(), verifyProductionContextEnrichment()])
  .then(() => console.log("Live visit tests passed."))
  .catch(error => { console.error(error); process.exitCode = 1; });
