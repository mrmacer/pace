/* ─────────────────────────────────────────────────────────────────────────
   PATCH A — IEP_App_Users reader coverage

   Plain-script style, matching this project's other tests (no test
   framework — assertions throw, Promise.all + process.exitCode surfaces
   failures). Run with: node tests/iep-app-users.test.js

   Covers only iep-app-users.js in isolation. No app.js/auth.js behavior
   changes yet in PATCH A, so there is nothing end-to-end to test here —
   that arrives with the PACE gate patch.
   ───────────────────────────────────────────────────────────────────────── */

const assert = require("node:assert/strict");
const fs = require("node:fs");
const path = require("node:path");
const vm = require("node:vm");

const ROOT = path.resolve(__dirname, "..");

function loadModule(context) {
  vm.runInContext(fs.readFileSync(path.join(ROOT, "iep-app-users.js"), "utf8"), context, { filename: "iep-app-users.js" });
}

function makeContext(overrides = {}) {
  const context = vm.createContext({
    console,
    APP_MODE: "production",
    CONFIG: { LISTS: { appUsers: "IEP_App_Users" } },
    ...overrides
  });
  return context;
}

async function verifyBooleanNormalization() {
  const context = makeContext();
  loadModule(context);
  const normalize = vm.runInContext("normalizeAppUsersBoolean", context);
  for (const truthy of ["Yes", "yes", " YES ", "true", "1", true, 1]) {
    assert.equal(normalize(truthy), true, `expected truthy for ${JSON.stringify(truthy)}`);
  }
  for (const falsy of ["No", "no", "false", "0", false, 0]) {
    assert.equal(normalize(falsy), false, `expected falsy for ${JSON.stringify(falsy)}`);
  }
  assert.equal(normalize(undefined, true), true, "blank/undefined must use the caller's default");
  assert.equal(normalize("", false), false, "blank string must use the caller's default");
  assert.equal(normalize(null), false, "default with no explicit defaultValue must be false");
}

async function verifyDefaultIsMigrationMode() {
  const context = makeContext();
  loadModule(context);
  const accessConfig = vm.runInContext("ACCESS_MODEL_CONFIG", context);
  assert.equal(accessConfig.ENFORCE_APP_USERS, false, "PATCH A must ship in migration mode, not enforced mode");
}

async function verifyResolveFindsMatchByEmailCaseInsensitive() {
  let calledList = null;
  const context = makeContext({
    GRAPH: {
      async getListItems(listName) {
        calledList = listName;
        return [
          { id: "1", field_1: "Sharon Morgan", field_2: "Behavior Specialist", field_3: true, field_4: "Sharon.Morgan@iu29.org", field_5: "Yes", field_6: "No" },
          { id: "2", field_1: "Kelly Higgins", field_2: "Teacher", field_3: true, field_4: "khiggins@iu29.org", field_5: "No", field_6: "Yes" }
        ];
      },
      async getListSchema() {
        return {
          "Title": "Title",
          "Display Name": "field_1",
          "Role": "field_2",
          "Teacher/Classroom": "field_9",
          "Email": "field_4",
          "PACE": "field_5",
          "Daily Pulse": "field_6",
          "Walkthrough": "field_7",
          "Admin Panel": "field_8"
        };
      }
    }
  });
  loadModule(context);
  const appUsers = vm.runInContext("APP_USERS", context);
  const row = await appUsers.resolve("  SHARON.MORGAN@IU29.ORG  ");
  assert.equal(calledList, "IEP_App_Users");
  assert.ok(row, "expected a matching row");
  assert.equal(row["Display Name"], "Sharon Morgan");
  assert.equal(appUsers.found, true);
  assert.equal(appUsers.lookupError, null);
  // JSON round-trip strips the vm context's separate-realm prototype so
  // deepEqual compares plain structure, not object identity (see the same
  // pattern in recent-activity.test.js / live-visit.test.js).
  assert.deepEqual(JSON.parse(JSON.stringify(appUsers.permissions)), { dailyPulse: false, pace: true, walkthrough: false, adminPanel: false });
}

async function verifyResolveNoMatchReturnsNullCleanly() {
  const context = makeContext({
    GRAPH: {
      async getListItems() { return [{ id: "1", field_4: "someone-else@iu29.org" }]; },
      async getListSchema() { return { "Email": "field_4" }; }
    }
  });
  loadModule(context);
  const appUsers = vm.runInContext("APP_USERS", context);
  const row = await appUsers.resolve("nobody@iu29.org");
  assert.equal(row, null);
  assert.equal(appUsers.found, false);
  assert.equal(appUsers.permissions, null);
  assert.equal(appUsers.lookupError, null, "no match is not an error");
}

async function verifyResolveFailsClosedOnGraphError() {
  const context = makeContext({
    GRAPH: {
      async getListItems() { throw new Error("Graph 403: insufficient privileges"); },
      async getListSchema() { return {}; }
    }
  });
  loadModule(context);
  const appUsers = vm.runInContext("APP_USERS", context);
  const row = await appUsers.resolve("sharon.morgan@iu29.org");
  assert.equal(row, null, "a failed lookup must never resolve to a match");
  assert.equal(appUsers.found, false);
  assert.match(appUsers.lookupError, /insufficient privileges/);
}

async function verifyDemoModeNeverCallsGraph() {
  let graphCalled = false;
  const context = makeContext({
    APP_MODE: "demo",
    GRAPH: {
      async getListItems() { graphCalled = true; return []; },
      async getListSchema() { graphCalled = true; return {}; }
    }
  });
  loadModule(context);
  const appUsers = vm.runInContext("APP_USERS", context);
  const row = await appUsers.resolve("sharon.morgan@iu29.org");
  assert.equal(row, null);
  assert.equal(graphCalled, false, "demo mode must never touch Graph, exactly like AUTH/GRAPH elsewhere in this app");
}

Promise.all([
  verifyBooleanNormalization(),
  verifyDefaultIsMigrationMode(),
  verifyResolveFindsMatchByEmailCaseInsensitive(),
  verifyResolveNoMatchReturnsNullCleanly(),
  verifyResolveFailsClosedOnGraphError(),
  verifyDemoModeNeverCallsGraph()
])
  .then(() => console.log("IEP_App_Users reader (PATCH A) tests passed."))
  .catch(error => {
    console.error(error);
    process.exitCode = 1;
  });
