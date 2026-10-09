/* ─────────────────────────────────────────────────────────────────────────
   RETRY-IDEMPOTENCY PATCH — GRAPH.savePaceVisit() must not create a second
   SharePoint row when "Try Again" re-sends an entry whose first POST
   actually succeeded but whose response was lost.
   ───────────────────────────────────────────────────────────────────────── */

const assert = require("node:assert/strict");
const fs = require("node:fs");
const path = require("node:path");
const vm = require("node:vm");

const ROOT = path.resolve(__dirname, "..");
const TODAY = "2026-10-09";

function loadGraph() {
  const context = vm.createContext({
    console,
    APP_MODE: "production",
    CONFIG: { SITE: "example.sharepoint.test:/sites/PACE", LISTS: { paceVisits: "IEP_Pace_Visits" } },
    AUTH: { acquireGraphToken: async () => "unused" },
    fetch: async () => { throw new Error("Unexpected network request"); },
    paceRoomLabelForId: id => id
  });
  vm.runInContext(fs.readFileSync(path.join(ROOT, "graph.js"), "utf8"), context, { filename: "graph.js" });
  vm.runInContext("this.testGraph = GRAPH", context);
  return context.testGraph;
}

// A fake IEP_Pace_Visits list. `loseNextResponse` makes the next POST
// persist its row but throw, exactly like a dropped network response.
function fakeList(graph, { hasEntryIdColumn }) {
  const world = { rows: [], posts: 0, reads: 0, loseNextResponse: false };
  graph.createMappedListItem = async (list, fields) => {
    world.posts += 1;
    const row = { id: String(world.rows.length + 1) };
    for (const key of ["Student ID", "Date", "Time In", "Time Out"]) row[key] = fields[key];
    if (hasEntryIdColumn) row["Entry ID"] = fields["Entry ID"];
    world.rows.push(row);
    if (world.loseNextResponse) { world.loseNextResponse = false; throw new Error("Failed to fetch"); }
    return { id: row.id };
  };
  graph.getPaceVisitsForDateByDisplayName = async date => {
    world.reads += 1;
    return world.rows.filter(row => row.Date === date);
  };
  return world;
}

const entry = (overrides = {}) => ({
  id: "sub-1", paceRoom: "pace-room-1", studentId: "S100", studentName: "Alex R.",
  date: TODAY, timeIn: "09:15", timeOut: "09:42", behaviors: ["Needs a Break"], ...overrides
});

async function lostResponseThenRetry(hasEntryIdColumn) {
  const graph = loadGraph();
  const world = fakeList(graph, { hasEntryIdColumn });

  world.loseNextResponse = true;
  await assert.rejects(graph.savePaceVisit(entry()), /Failed to fetch/);
  assert.equal(world.rows.length, 1, "first POST landed on SharePoint");

  const retry = await graph.savePaceVisit(entry());
  assert.equal(world.posts, 1, "Try Again must not POST a second row");
  assert.equal(world.rows.length, 1);
  assert.equal(retry.duplicatePrevented, true);
  assert.equal(retry.id, "1", "the already-saved row's id is returned");
}

async function firstAttemptSkipsTheRead() {
  const graph = loadGraph();
  const world = fakeList(graph, { hasEntryIdColumn: false });
  await graph.savePaceVisit(entry());
  assert.equal(world.reads, 0, "a normal save costs no extra request");
  assert.equal(world.posts, 1);
}

async function retryAfterRealFailureStillSaves() {
  const graph = loadGraph();
  const world = fakeList(graph, { hasEntryIdColumn: false });
  const realCreate = graph.createMappedListItem;
  graph.createMappedListItem = async () => { world.posts += 1; throw new Error("Graph POST 503"); };
  await assert.rejects(graph.savePaceVisit(entry()), /503/);
  graph.createMappedListItem = realCreate;

  const retry = await graph.savePaceVisit(entry());
  assert.equal(world.rows.length, 1, "nothing was saved the first time, so the retry creates the row");
  assert.notEqual(retry.duplicatePrevented, true);
}

async function differentVisitIsNotSwallowed() {
  const graph = loadGraph();
  const world = fakeList(graph, { hasEntryIdColumn: false });
  world.loseNextResponse = true;
  await assert.rejects(graph.savePaceVisit(entry()), /Failed to fetch/);

  // Same submission id re-sent after the user corrected the times: no
  // existing row matches, so it must be written.
  await graph.savePaceVisit(entry({ timeIn: "10:00", timeOut: "10:20" }));
  assert.equal(world.rows.length, 2);
}

(async () => {
  await lostResponseThenRetry(false);
  await lostResponseThenRetry(true);
  await firstAttemptSkipsTheRead();
  await retryAfterRealFailureStillSaves();
  await differentVisitIsNotSwallowed();
  console.log("save-retry-idempotency tests passed");
})().catch(err => { console.error(err); process.exit(1); });
