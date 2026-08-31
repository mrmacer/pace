/* ─────────────────────────────────────────────────────────────────────────
   PATCH C — PACE Room Tracker authorization via IEP_App_Users

   Two layers of coverage, matching this project's existing conventions:
   1. Behavioral: APP_USERS.decide() exercised directly (iep-app-users.js
      already has isolated resolve()/normalize coverage from PATCH A — this
      adds the decision helper PATCH C introduces).
   2. Static: boot()/index.html are grep-checked (like
      patch010-specialists-notes.test.js) rather than executed, since app.js
      queries the live DOM at load time and isn't designed to run under
      node:vm the way the other three apps' app.js files are.

   Run with: node tests/patch-c-authorization.test.js
   ───────────────────────────────────────────────────────────────────────── */

const assert = require("node:assert/strict");
const fs = require("node:fs");
const path = require("node:path");
const vm = require("node:vm");

const ROOT = path.resolve(__dirname, "..");
const appSrc = fs.readFileSync(path.join(ROOT, "app.js"), "utf8");
const html = fs.readFileSync(path.join(ROOT, "index.html"), "utf8");

function loadAppUsers(context) {
  vm.runInContext(fs.readFileSync(path.join(ROOT, "iep-app-users.js"), "utf8"), context, { filename: "iep-app-users.js" });
}

function makeContext(overrides = {}) {
  return vm.createContext({
    console,
    APP_MODE: "production",
    CONFIG: { LISTS: { appUsers: "IEP_App_Users" } },
    ...overrides
  });
}

async function verifyExplicitFlagOverridesLegacyAllow() {
  const context = makeContext({
    GRAPH: {
      async getListItems() { return [{ id: "1", field_1: "sharon.morgan@iu29.org", field_2: "No" }]; },
      async getListSchema() { return { "Email": "field_1", "PACE": "field_2" }; }
    }
  });
  loadAppUsers(context);
  const appUsers = vm.runInContext("APP_USERS", context);
  await appUsers.resolve("sharon.morgan@iu29.org");
  // legacyAllowed=true (today's actual PACE behavior) must be overridden by
  // an explicit "PACE: No" row.
  const decision = appUsers.decide("PACE", true);
  assert.equal(decision.allowed, false);
  assert.equal(decision.reason, "app-users-denied");
}

async function verifyExplicitFlagAllows() {
  const context = makeContext({
    GRAPH: {
      async getListItems() { return [{ id: "1", field_1: "sharon.morgan@iu29.org", field_2: "Yes" }]; },
      async getListSchema() { return { "Email": "field_1", "PACE": "field_2" }; }
    }
  });
  loadAppUsers(context);
  const appUsers = vm.runInContext("APP_USERS", context);
  await appUsers.resolve("sharon.morgan@iu29.org");
  assert.equal(appUsers.decide("PACE", true).allowed, true);
}

async function verifyNoRowFallsBackToLegacyAllowed() {
  const context = makeContext({
    GRAPH: {
      async getListItems() { return [{ id: "1", field_1: "someone-else@iu29.org" }]; },
      async getListSchema() { return { "Email": "field_1" }; }
    }
  });
  loadAppUsers(context);
  const appUsers = vm.runInContext("APP_USERS", context);
  await appUsers.resolve("sharon.morgan@iu29.org");
  // Today's actual PACE behavior: any active IEP_Users2 user is allowed.
  const decision = appUsers.decide("PACE", true);
  assert.equal(decision.allowed, true);
  assert.equal(decision.reason, "legacy-fallback");
}

async function verifyLookupFailureAlwaysDenies() {
  const context = makeContext({
    GRAPH: {
      async getListItems() { throw new Error("Graph 403: insufficient privileges"); },
      async getListSchema() { return {}; }
    }
  });
  loadAppUsers(context);
  const appUsers = vm.runInContext("APP_USERS", context);
  await appUsers.resolve("sharon.morgan@iu29.org");
  const decision = appUsers.decide("PACE", true); // legacyAllowed=true must NOT rescue a lookup failure
  assert.equal(decision.allowed, false);
  assert.equal(decision.reason, "app-users-lookup-failed");
}

async function verifyAdminWithPaceYesAllowed() {
  const context = makeContext({
    GRAPH: {
      async getListItems() { return [{ id: "1", field_1: "admin@iu29.org", field_2: "Yes" }]; },
      async getListSchema() { return { "Email": "field_1", "PACE": "field_2" }; }
    }
  });
  loadAppUsers(context);
  const appUsers = vm.runInContext("APP_USERS", context);
  await appUsers.resolve("admin@iu29.org");
  assert.equal(appUsers.decide("PACE", true).allowed, true);
}

async function verifyBootWiring() {
  // Static checks mirroring patch010-specialists-notes.test.js's approach:
  // app.js is not designed to run under node:vm (top-level
  // document.querySelectorAll against a real DOM), so pin the gate's shape
  // in source instead of executing boot().
  assert.match(appSrc, /await APP_USERS\.resolve\(AUTH\.account\?\.username \|\| ""\)/);
  assert.match(appSrc, /APP_USERS\.decide\("PACE",\s*true\)/);
  assert.match(appSrc, /if \(!paceDecision\.allowed\)/);

  // The gate must run, and return on denial, strictly before nav("home"...)
  // — i.e. before any room is ever shown — and before the diagnostic
  // button (which only makes sense once inside the app) is unhidden.
  const gateIndex = appSrc.indexOf("APP_USERS.resolve(AUTH.account");
  const homeNavIndex = appSrc.indexOf('nav("home", "forward")');
  const diagnosticRevealIndex = appSrc.indexOf('runDiagnosticBtn")?.classList.remove("hidden")');
  assert.ok(gateIndex >= 0 && gateIndex < homeNavIndex, "PACE gate must run before nav(\"home\")");
  assert.ok(gateIndex < diagnosticRevealIndex, "PACE gate must run before the diagnostic button is revealed");

  // Denial must route through showUnauthorized() (same screen/copy as the
  // existing IEP_Users2 unauthorized path) — not a bespoke dead end.
  assert.match(appSrc, /showUnauthorized\("You are signed in, but your account does not currently have access to PACE Room Tracker\."\)/);

  // Consistent IEP Skook template present in the unauthorized screen.
  assert.match(html, /data-screen="unauthorized"[\s\S]*?IEP Skook/);
  assert.match(html, /id="unauthorizedSignedInAs"/);
  assert.match(html, /Contact an IEP Skook administrator/);
}

Promise.all([
  verifyExplicitFlagOverridesLegacyAllow(),
  verifyExplicitFlagAllows(),
  verifyNoRowFallsBackToLegacyAllowed(),
  verifyLookupFailureAlwaysDenies(),
  verifyAdminWithPaceYesAllowed(),
  verifyBootWiring()
])
  .then(() => console.log("PATCH C (PACE authorization) tests passed."))
  .catch(error => {
    console.error(error);
    process.exitCode = 1;
  });
