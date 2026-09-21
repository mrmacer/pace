/* STAFF EDIT + DELETE PATCH tests.

   Plain Node script (repo convention: `node tests/x.test.js`). The real
   graph.js, pace-data.js, recent-activity.js, visit-workflow.js and
   visit-corrections.js run together inside one vm context against a fake
   Microsoft Graph `fetch`, so the whole path is exercised:

     Graph item.id -> Recent Activity row -> editModel().itemId
       -> PATCH|DELETE sites/{site}/lists/{IEP_Pace_Visits}/items/{id}

   app.js is never loaded into a DOM (repo convention); UI, authorization
   and regression items are asserted structurally against its source. */

const assert = require("node:assert/strict");
const fs = require("node:fs");
const path = require("node:path");
const vm = require("node:vm");

const ROOT = path.resolve(__dirname, "..");
const read = file => fs.readFileSync(path.join(ROOT, file), "utf8");
const app = read("app.js");
const html = read("index.html");
const graphSrc = read("graph.js");
const paceDataSrc = read("pace-data.js");
const correctionsSrc = read("visit-corrections.js");
const sw = read("sw.js");
const plain = value => JSON.parse(JSON.stringify(value));

const now = new Date();
const TODAY = `${now.getFullYear()}-${String(now.getMonth() + 1).padStart(2, "0")}-${String(now.getDate()).padStart(2, "0")}`;

function functionBody(source, startPattern) {
  const start = source.search(startPattern);
  assert.ok(start >= 0, `could not find ${startPattern}`);
  const open = source.indexOf("{", start);
  let depth = 0;
  for (let i = open; i < source.length; i += 1) {
    if (source[i] === "{") depth += 1;
    if (source[i] === "}") { depth -= 1; if (depth === 0) return source.slice(start, i + 1); }
  }
  throw new Error(`unbalanced braces after ${startPattern}`);
}

/* ── Fake Microsoft Graph ─────────────────────────────────────────────── */

const SCHEMA_COLUMNS = [
  ["Room", "Room"], ["Student", "Title"], ["Date", "Date"], ["Time In", "TimeIn"], ["Time Out", "TimeOut"],
  ["Duration", "Duration"], ["Reason", "Reason"], ["Intervention Used", "InterventionUsed"],
  ["SCM Used", "SCMUsed"], ["Notes", "Notes"], ["Behavior Specialist", "BehaviorSpecialist"],
  ["Teacher Came From", "TeacherCameFrom"], ["Entry ID", "EntryID"], ["Submitted By", "SubmittedBy"],
  ["Created", "Created"], ["Modified", "Modified"], ["Editor", "Editor"]
];

function makeWorld({ deleteStatus = 204 } = {}) {
  const items = [
    { id: "42", createdDateTime: `${TODAY}T14:00:00Z`, fields: {
      Room: "PACE Room 1", Title: "Alex R.", Date: `${TODAY}T00:00:00Z`, TimeIn: "09:15", TimeOut: "09:42", Duration: 27,
      Reason: "Needs a Break", InterventionUsed: "Check-in", SCMUsed: false, Notes: "Original note",
      BehaviorSpecialist: "Dana Fielding", TeacherCameFrom: "Mrs. Ashford", EntryID: "sub-42", SubmittedBy: "Staff A"
    } },
    { id: "43", createdDateTime: `${TODAY}T14:10:00Z`, fields: {   // historically blank Notes
      Room: "PACE Room 1", Title: "Jordan M.", Date: `${TODAY}T00:00:00Z`, TimeIn: "10:00", TimeOut: "10:20", Duration: 20,
      Reason: "Peer conflict", InterventionUsed: "", SCMUsed: false, Notes: "", EntryID: "sub-43"
    } },
    { id: "44", createdDateTime: `${TODAY}T14:20:00Z`, fields: {   // OPEN visit: no Time Out
      Room: "PACE Room 1", Title: "Casey L.", Date: `${TODAY}T00:00:00Z`, TimeIn: "11:00", EntryID: "sub-44"
    } }
  ];
  const world = { items, calls: [], deleteStatus };
  world.mutations = () => world.calls.filter(call => call.method !== "GET");
  world.fetch = async (url, options = {}) => {
    const method = options.method || "GET";
    const rel = String(url).replace("https://graph.microsoft.com/v1.0/", "");
    world.calls.push({ method, rel, body: options.body === undefined ? undefined : JSON.parse(options.body), headers: options.headers });
    const json = data => ({ ok: true, status: 200, json: async () => data, text: async () => JSON.stringify(data) });
    if (method === "GET" && rel.startsWith("sites/example.sharepoint.test")) return json({ id: "site-id" });
    if (method === "GET" && /^sites\/site-id\/lists\?/.test(rel)) {
      return json({ value: [
        { id: "list-decoy", name: "IEP_Users2", displayName: "IEP_Users2" },
        { id: "list-pace", name: "IEP_Pace_Visits", displayName: "IEP_Pace_Visits" }
      ] });
    }
    if (method === "GET" && /^sites\/site-id\/lists\/list-pace\/columns/.test(rel)) {
      return json({ value: SCHEMA_COLUMNS.map(([displayName, name]) => ({ displayName, name })) });
    }
    if (method === "GET" && /^sites\/site-id\/lists\/list-pace\/items\?/.test(rel)) return json({ value: world.items });
    const target = rel.match(/^sites\/site-id\/lists\/(list-pace)\/items\/(\d+)(\/fields)?$/);
    if (method === "PATCH" && target && target[3]) {
      const item = world.items.find(entry => entry.id === target[2]);
      Object.assign(item.fields, JSON.parse(options.body));
      return json(item.fields);
    }
    if (method === "DELETE" && target && !target[3]) {
      if (world.deleteStatus === 204) {
        world.items = world.items.filter(entry => entry.id !== target[2]);
        return { ok: true, status: 204, json: async () => { throw new Error("204 has no body"); }, text: async () => "" };
      }
      return { ok: false, status: world.deleteStatus, json: async () => ({}), text: async () => "graph error body with details" };
    }
    throw new Error(`Unexpected Graph request: ${method} ${rel}`);
  };
  return world;
}

function makeContext(world) {
  const storage = new Map();
  const context = vm.createContext({
    console: { ...console, error() {}, warn() {} },
    APP_MODE: "production",
    CONFIG: {
      SITE: "example.sharepoint.test:/sites/IEP_Skook:",
      LISTS: { paceVisits: "IEP_Pace_Visits" },
      STORAGE_KEYS: { VISIT_CONTEXT: "visitContext" },
      ROOMS: [
        { id: "pace-room-1", label: "PACE Room 1" },
        { id: "pace-room-2", label: "PACE Room 2" }
      ],
      BEHAVIOR_SPECIALISTS: []
    },
    AUTH: { acquireGraphToken: async () => "test-token" },
    fetch: world.fetch,
    localStorage: {
      getItem: key => storage.get(key) ?? null,
      setItem: (key, value) => storage.set(key, value),
      removeItem: key => storage.delete(key)
    },
    ROSTER: { loaded: false }
  });
  vm.runInContext(`
    function paceRoomLabelForId(id) { const r = CONFIG.ROOMS.find(x => x.id === id); return r ? r.label : (id || ""); }
    function paceRoomIdForLabel(v) { const raw = String(v ?? "").trim(); if (!raw) return ""; const r = CONFIG.ROOMS.find(x => x.label === raw || x.id === raw); return r ? r.id : raw; }
    const PACE_ROOM_FIELD_CANDIDATES = ["Room", "PACE Room", "Pace Room"];
  `, context);
  for (const file of ["recent-activity.js", "visit-workflow.js", "visit-corrections.js", "graph.js", "pace-data.js"]) {
    vm.runInContext(read(file), context, { filename: file });
  }
  vm.runInContext("this.T = { PACE_DATA, GRAPH, RECENT_ACTIVITY, PACE_VISIT_WORKFLOW, PACE_VISIT_CORRECTIONS }", context);
  return context.T;
}

// Reproduces what app.js does when Edit is tapped and saved: the Recent
// Activity row -> editModel() -> STATE.editOriginal snapshot -> edited entry.
function snapshotFor(T, visit) {
  const edit = T.RECENT_ACTIVITY.editModel(visit, { fallbackRoom: "pace-room-1", validRooms: ["pace-room-1", "pace-room-2"] });
  const original = {
    paceRoom: edit.room, studentName: edit.student, date: edit.date, timeIn: edit.timeIn, timeOut: edit.timeOut,
    behaviors: [...edit.reasons], interventions: [...edit.supports], scmUsed: edit.scmUsed, notes: edit.notes,
    staffMembers: [...edit.specialists], cameFromTeacher: edit.cameFromTeacher || ""
  };
  return { edit, original, entry: { ...original, behaviors: [...original.behaviors], interventions: [...original.interventions], staffMembers: [...original.staffMembers], original } };
}

async function recentVisit(T, id) {
  const result = await T.PACE_DATA.getRecentVisits({ date: TODAY, room: "pace-room-1", limit: 10 });
  return { result, visit: result.visits.find(v => v.id === id) };
}

/* ── Functional: item-id path, PATCH edits ────────────────────────────── */

async function verifyEditPath() {
  const world = makeWorld();
  const T = makeContext(world);
  const { result, visit } = await recentVisit(T, "42");
  assert.ok(visit, "the completed visit is in Recent Activity");
  assert.equal(visit.id, "42", "Graph item.id survives into the Recent Activity row");
  assert.equal(result.visits.some(v => v.id === "44"), false, "an OPEN visit (no Time Out) is not represented in Recent Activity");

  const { edit, original, entry } = snapshotFor(T, visit);
  assert.equal(edit.itemId, "42", "editModel().itemId is the SharePoint item id");
  world.calls.length = 0;

  // Notes-only edit.
  entry.notes = "Corrected note";
  await T.PACE_DATA.updateVisit(edit.itemId, entry);
  const writes = world.mutations();
  assert.equal(writes.length, 1, "exactly one write");
  assert.equal(writes[0].method, "PATCH", "edit uses PATCH, not POST");
  assert.equal(writes[0].rel, "sites/site-id/lists/list-pace/items/42/fields", "PATCH targets the original item id on IEP_Pace_Visits");
  assert.deepEqual(plain(writes[0].body), { Notes: "Corrected note" }, "notes-only edit sends only Notes");
  assert.equal(world.calls.some(c => c.method === "POST" || c.method === "DELETE"), false, "never POST or DELETE while editing");
  assert.deepEqual(world.items.map(i => i.id), ["42", "43", "44"], "no record was replaced: same ids, same count");
  const stored = world.items.find(i => i.id === "42").fields;
  assert.equal(stored.TimeIn, "09:15"); assert.equal(stored.TimeOut, "09:42"); assert.equal(stored.Duration, 27);
  assert.equal(stored.EntryID, "sub-42", "Submission/Entry ID untouched");

  // Whole-form edit: every payload key is an editable field, never a system field.
  const second = snapshotFor(T, (await recentVisit(T, "42")).visit);
  Object.assign(second.entry, {
    paceRoom: "pace-room-2", studentName: "Taylor S.", timeIn: "09:10", timeOut: "09:50",
    behaviors: ["Peer conflict"], interventions: ["Break"], scmUsed: true, notes: "All changed",
    staffMembers: ["Marcus Webb"], cameFromTeacher: "Mr. Bellamy"
  });
  world.calls.length = 0;
  await T.PACE_DATA.updateVisit("42", second.entry);
  const body = plain(world.mutations()[0].body);
  const allowed = new Set(["Room", "Title", "Date", "TimeIn", "TimeOut", "Duration", "Reason", "InterventionUsed", "SCMUsed", "Notes", "BehaviorSpecialist", "TeacherCameFrom"]);
  for (const key of Object.keys(body)) assert.ok(allowed.has(key), `${key} is an editable field`);
  for (const forbidden of ["id", "EntryID", "SubmittedBy", "Created", "Modified", "Editor", "SubmissionID"]) {
    assert.equal(forbidden in body, false, `${forbidden} is never written`);
  }
  assert.equal(body.Duration, T.PACE_VISIT_WORKFLOW.durationMinutes("09:10", "09:50"), "Duration recalculated with the canonical function");
  assert.equal(body.TimeIn, "09:10"); assert.equal(body.TimeOut, "09:50");
  assert.equal(world.items.find(i => i.id === "42").fields.EntryID, "sub-42");
}

async function verifyTimeAndDuration() {
  const world = makeWorld();
  const T = makeContext(world);
  const { visit } = await recentVisit(T, "42");
  const { entry } = snapshotFor(T, visit);
  entry.timeOut = "09:58";
  world.calls.length = 0;
  await T.PACE_DATA.updateVisit("42", entry);
  const body = plain(world.mutations()[0].body);
  assert.deepEqual(body, { TimeOut: "09:58", Duration: T.PACE_VISIT_WORKFLOW.durationMinutes("09:15", "09:58") });
  assert.equal("TimeIn" in body, false, "an unchanged Time In is not rewritten");

  // Notes-only: no Time Out fabricated, no Duration written.
  const again = snapshotFor(T, (await recentVisit(T, "42")).visit);
  again.entry.notes = "Just a note";
  world.calls.length = 0;
  await T.PACE_DATA.updateVisit("42", again.entry);
  const notesOnly = plain(world.mutations()[0].body);
  assert.deepEqual(Object.keys(notesOnly), ["Notes"]);

  // An edit can never turn a completed visit into an open one.
  const open = snapshotFor(T, (await recentVisit(T, "42")).visit);
  open.entry.timeOut = "";
  world.calls.length = 0;
  await assert.rejects(() => T.PACE_DATA.updateVisit("42", open.entry), /both Time In and Time Out/);
  assert.equal(world.mutations().length, 0);
}

async function verifyNotesRequired() {
  const world = makeWorld();
  const T = makeContext(world);
  const { visit } = await recentVisit(T, "43");
  assert.equal(visit.Notes, "", "fixture: historically blank Notes");

  for (const blank of ["", "   ", "\n\t "]) {
    const { entry } = snapshotFor(T, visit);
    entry.notes = blank;
    entry.timeOut = "10:25"; // a real change, so only Notes can be the blocker
    world.calls.length = 0;
    await assert.rejects(() => T.PACE_DATA.updateVisit("43", entry), /Add a brief note/, `blank Notes ${JSON.stringify(blank)} must fail`);
    assert.equal(world.mutations().length, 0, "nothing was written for blank/whitespace Notes");
  }

  const fixed = snapshotFor(T, visit);
  fixed.entry.notes = "Added the missing note";
  world.calls.length = 0;
  await T.PACE_DATA.updateVisit("43", fixed.entry);
  assert.deepEqual(plain(world.mutations()[0].body), { Notes: "Added the missing note" });

  // Guard against the "no changes" path becoming a silent empty PATCH.
  const unchanged = snapshotFor(T, (await recentVisit(T, "42")).visit);
  world.calls.length = 0;
  await assert.rejects(() => T.PACE_DATA.updateVisit("42", unchanged.entry), /No changes to save/);
  assert.equal(world.mutations().length, 0);
}

/* ── Functional: DELETE ───────────────────────────────────────────────── */

async function verifyDeleteRequest() {
  const world = makeWorld();
  const T = makeContext(world);
  const { visit } = await recentVisit(T, "42");
  world.calls.length = 0;

  const result = await T.PACE_DATA.deleteVisit(visit.id);
  assert.deepEqual(plain(result), { deleted: true, id: "42" }, "HTTP 204 is handled as success");
  const writes = world.mutations();
  assert.equal(writes.length, 1, "exactly one Graph DELETE");
  assert.equal(writes[0].method, "DELETE");
  assert.equal(writes[0].rel, "sites/site-id/lists/list-pace/items/42", "targets IEP_Pace_Visits (list-pace, not another list) and the exact item id");
  assert.equal(writes[0].body, undefined, "no request body");
  assert.deepEqual(world.items.map(i => i.id), ["43", "44"], "only item 42 is gone");

  // Recent Activity reloaded from SharePoint no longer shows it.
  const after = await T.PACE_DATA.getRecentVisits({ date: TODAY, room: "pace-room-1", limit: 10 });
  assert.equal(after.visits.some(v => v.id === "42"), false);
  assert.equal(after.visits.some(v => v.id === "43"), true);
}

async function verifyDeleteRefusesBadIds() {
  for (const bad of ["", "   ", null, undefined, "abc", "42/fields", "42?x=1", "../7", "4 2", "4\n2", "a1"]) {
    const world = makeWorld();
    const T = makeContext(world);
    await assert.rejects(() => T.GRAPH.deletePaceVisit(bad), /valid SharePoint visit item/);
    assert.equal(world.calls.length, 0, `no network at all for id ${JSON.stringify(bad)}`);
  }
}

async function verifyDeleteFailuresKeepRecord() {
  for (const status of [401, 403, 404, 500]) {
    const world = makeWorld({ deleteStatus: status });
    const T = makeContext(world);
    const { visit } = await recentVisit(T, "42");
    let reloaded = 0;
    const outcome = await T.PACE_VISIT_CORRECTIONS.runDelete({
      visit, confirmed: true, isAuthorized: () => true,
      deleteVisit: id => T.PACE_DATA.deleteVisit(id),
      onSuccess: () => { reloaded += 1; }
    });
    assert.equal(outcome.status, "failed", `status ${status}`);
    assert.equal(reloaded, 0, "list reload/success feedback never runs after a failed DELETE");
    assert.equal(world.items.some(i => i.id === "42"), true, "record still exists");
    assert.equal(world.mutations().length, 1, "no retry, no second write");
    const message = T.PACE_VISIT_CORRECTIONS.deleteFailureMessage(outcome.error);
    assert.ok(message.length > 10 && !/graph error body/.test(message), "useful message, no raw Graph body");
    const still = await T.PACE_DATA.getRecentVisits({ date: TODAY, room: "pace-room-1", limit: 10 });
    assert.equal(still.visits.some(v => v.id === "42"), true, "record remains visible after refresh");
  }
}

async function verifyRunDeleteGuards() {
  const T = makeContext(makeWorld());
  const visit = { id: "42" };
  const cases = [
    ["cancel (not confirmed)", { confirmed: false, isAuthorized: () => true }],
    ["confirmed undefined", { confirmed: undefined, isAuthorized: () => true }],
    ["confirmed truthy-but-not-true", { confirmed: "yes", isAuthorized: () => true }],
    ["unauthorized", { confirmed: true, isAuthorized: () => false }],
    ["no authorization function", { confirmed: true }]
  ];
  for (const [label, extra] of cases) {
    let deletes = 0;
    await T.PACE_VISIT_CORRECTIONS.runDelete({ visit, ...extra, deleteVisit: async () => { deletes += 1; }, onSuccess() {} });
    assert.equal(deletes, 0, `${label}: zero DELETE calls`);
  }
  let deletes = 0; const ids = []; let order = "";
  const ok = await T.PACE_VISIT_CORRECTIONS.runDelete({
    visit, confirmed: true, isAuthorized: () => true,
    deleteVisit: async id => { deletes += 1; ids.push(id); order += "D"; },
    onSuccess: async () => { order += "S"; }
  });
  assert.equal(ok.status, "deleted");
  assert.equal(deletes, 1); assert.deepEqual(ids, ["42"]); assert.equal(order, "DS", "success feedback only after the delete resolved");
  let called = 0;
  const invalid = await T.PACE_VISIT_CORRECTIONS.runDelete({ visit: { Student: "x" }, confirmed: true, isAuthorized: () => true, deleteVisit: async () => { called += 1; } });
  assert.equal(invalid.status, "invalid"); assert.equal(called, 0, "a record with no item id is never deletable");
}

/* ── Structural: UI, authorization, regressions ───────────────────────── */

function verifyUiStructure() {
  const card = functionBody(app, /async function loadRecentActivity\(\)/);
  assert.match(card, /class="recent-edit-btn"[^>]*data-recent-edit=/, "Edit is on every Recent Activity card");
  assert.match(card, /class="recent-delete-btn"[^>]*data-recent-delete=[^>]*aria-label="Delete this visit"/, "trash action beside Edit");
  assert.ok(card.indexOf("recent-edit-btn") < card.indexOf("recent-delete-btn"), "[Edit] then [trash]");
  assert.doesNotMatch(card, /isAdmin|Administrator|AUTH\.role|\.role\b/, "card actions are not role-gated");

  // The confirmation is opened in exactly one place: the trash click path.
  assert.equal((app.match(/getElementById\("deleteVisitOverlay"\)\.classList\.remove\("hidden"\)/g) || []).length, 1);
  const request = functionBody(app, /function requestDeleteVisit\(/);
  assert.match(request, /deleteVisitOverlay/);
  assert.doesNotMatch(request, /PACE_DATA\.deleteVisit|runDelete|GRAPH\./, "opening the dialog never deletes");
  assert.match(request, /summary\.student, summary\.date, summary\.time, summary\.room/, "dialog names student, date, times, room");

  // Only the dialog's Delete button passes confirmed:true, and it is the only caller.
  assert.equal((app.match(/confirmed:\s*true/g) || []).length, 1);
  assert.equal((app.match(/PACE_DATA\.deleteVisit\(/g) || []).length, 1);
  assert.equal((app.match(/PACE_VISIT_CORRECTIONS\.runDelete\(/g) || []).length, 1);
  const confirm = app.slice(app.indexOf('getElementById("confirmDeleteVisitBtn").addEventListener'));
  assert.match(confirm, /isAuthorized:\s*isPaceAuthorized/, "delete handler itself re-checks the canonical PACE authorization");
  assert.match(confirm, /await loadRecentActivity\(\);\s*await refreshRoomVisits\(\);/, "reload from SharePoint after confirmed delete");
  assert.match(confirm, /showToast\("PACE visit deleted\."\)/);
  assert.doesNotMatch(app.slice(app.indexOf("TRASH_ICON_SVG ="), app.indexOf("async function loadRecentActivity")), /\.remove\(\)|splice|filter\(.*visit/, "no optimistic hiding before SharePoint confirms");
  assert.match(html, /id="deleteVisitOverlay"[^>]*role="alertdialog"/);
  assert.match(html, /id="confirmDeleteVisitBtn"[^>]*>\s*Delete Visit\s*</, "destructive button says Delete");
  assert.match(html, /id="cancelDeleteVisitBtn"/);
  assert.match(html, /<script src="visit-corrections\.js"><\/script>/);
  const recentSection = html.slice(html.indexOf('data-screen="recent"'));
  assert.ok(recentSection.indexOf('id="deleteVisitOverlay"') > 0 && recentSection.indexOf('id="deleteVisitOverlay"') < recentSection.indexOf("</section>"), "dialog lives inside the Recent Activity screen");

  // After a confirmed edit the app still returns to a refreshed Recent list.
  assert.match(app, /PACE_DATA\.updateVisit\(STATE\.editingVisitId, entry\)/);
}

function verifyAuthorizationNotAdminOnly() {
  const gate = functionBody(app, /function isPaceAuthorized\(/);
  assert.match(gate, /APP_USERS\.decide\("PACE", true\)\.allowed/, "reuses the existing PACE authorization decision");
  const flows = [
    gate,
    functionBody(app, /function requestDeleteVisit\(/),
    functionBody(app, /async function beginEditVisit\(/),
    correctionsSrc,
    functionBody(graphSrc, /async deletePaceVisit\(/),
    functionBody(paceDataSrc, /async deleteVisit\(/)
  ].join("\n");
  const code = flows.replace(/\/\*[\s\S]*?\*\//g, "").replace(/\/\/.*$/gm, "");
  assert.doesNotMatch(code, /isAdmin|Administrator|AUTH\.role|\.role\b|\brole\b|IEP_Users2/i, "no role/Administrator logic anywhere in edit or delete");
  assert.match(functionBody(app, /async function beginEditVisit\(/), /isPaceAuthorized\(\)/, "edit is gated by the same check");
  // No second role system: the only decide("PACE") consumers are boot's gate and the one helper.
  assert.equal((app.match(/APP_USERS\.decide\("PACE", true\)/g) || []).length, 2);
  // boot's gate is unchanged.
  assert.match(app, /const paceDecision = APP_USERS\.decide\("PACE", true\);\s*if \(!paceDecision\.allowed\) \{/);
}

function verifyGraphSurface() {
  const withoutComments = src => src.replace(/\/\*[\s\S]*?\*\//g, "").replace(/\/\/.*$/gm, "");
  const g = withoutComments(graphSrc);
  assert.equal((g.match(/method:\s*"DELETE"/g) || []).length, 1, "one DELETE transport");
  assert.equal((g.match(/this\._delete\(/g) || []).length, 1, "one caller of _delete: deletePaceVisit");
  const del = withoutComments(functionBody(graphSrc, /async deletePaceVisit\(/));
  assert.match(del, /CONFIG\.LISTS\.paceVisits/, "targets IEP_Pace_Visits only");
  assert.doesNotMatch(del, /\$filter|findListItemByDisplayField|getListItems|forEach|\bfor\s*\(|\.map\(|Student|Submission|Entry ID/, "no bulk / by-name / by-Submission-ID delete");
  assert.doesNotMatch(g, /"@microsoft\.graph\.conflictBehavior"|deletedDateTime|isDeleted/, "no soft-delete scheme");
  // Edit is PATCH against the existing item; never delete + create.
  const upd = withoutComments(functionBody(paceDataSrc, /async updateVisit\(/));
  assert.doesNotMatch(upd, /createVisit|savePaceVisit|deleteVisit|deletePaceVisit|_post|\bPOST\b/, "edit is never delete + create");
  assert.match(withoutComments(functionBody(graphSrc, /async updatePaceVisit\(/)), /updateMappedListItem\("IEP_Pace_Visits", itemId,/);
  assert.match(g, /async updateMappedListItem\([\s\S]*?this\._patch\(`sites\/\$\{siteId\}\/lists\/\$\{listId\}\/items\/\$\{itemId\}\/fields`/);
}

function verifyRegressionsUntouched() {
  // Start Visit / Mark Complete path.
  assert.match(paceDataSrc, /async createVisit\(entry\) \{/);
  assert.match(functionBody(paceDataSrc, /async completeVisit\(/), /"Time Out":\s+entry\.timeOut/);
  assert.match(app, /PACE_DATA\.completeVisit\(/);
  assert.match(app, /openExitScreen/);
  // Required Notes at save time, still ahead of the new edit guard.
  const save = app.slice(app.indexOf('getElementById("saveEntryBtn").addEventListener'));
  assert.ok(save.indexOf("Add a brief note before submitting this PACE visit.") > 0);
  assert.ok(save.indexOf("Add a brief note before submitting this PACE visit.") < save.indexOf("isPaceAuthorized()"), "required-Notes check runs before the edit guard");
  assert.match(app, /const entry = buildVisitEntry\(\);/);
  // Starting a live visit still does not need Notes.
  assert.match(app, /buildVisitEntry\(\{ open: true \}\)/);
  const entrySrc = app.slice(app.indexOf("function buildVisitEntry("));
  assert.match(entrySrc.slice(0, entrySrc.indexOf("\n}\n")), /notes: open \? "" : STATE\.notes/);
  // Full roster module still present and wired.
  assert.ok(fs.existsSync(path.join(ROOT, "student-select.js")));
  assert.match(html, /<script src="student-select\.js"><\/script>/);
  // Authentication + IEP_App_Users layer untouched: migration mode still on.
  assert.match(read("iep-app-users.js"), /ENFORCE_APP_USERS:\s*false\b/);
  assert.match(app, /AUTH\.isAuthenticated/);
  // Service worker: bumped past v17, network-first, new file cached.
  assert.match(sw, /CACHE_NAME = "pace-tracker-shell-v18"/);
  assert.match(sw, /"\.\/visit-corrections\.js"/);
  assert.match(sw, /fetch\(event\.request\)/);
  assert.doesNotMatch(sw, /pace-tracker-shell-v17"/);
}

Promise.resolve()
  .then(verifyEditPath)
  .then(verifyTimeAndDuration)
  .then(verifyNotesRequired)
  .then(verifyDeleteRequest)
  .then(verifyDeleteRefusesBadIds)
  .then(verifyDeleteFailuresKeepRecord)
  .then(verifyRunDeleteGuards)
  .then(() => {
    verifyUiStructure();
    verifyAuthorizationNotAdminOnly();
    verifyGraphSurface();
    verifyRegressionsUntouched();
  })
  .then(() => console.log("PACE staff edit + delete tests passed."))
  .catch(error => { console.error(error); process.exitCode = 1; });
