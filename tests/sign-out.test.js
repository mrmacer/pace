/* ─────────────────────────────────────────────────────────────────────────
   PACE Room Tracker — focused sign-out coverage

   Pins the existing MSAL helper's redirect request and the app-owned local
   session cleanup without constructing a browser or contacting Microsoft.
   ───────────────────────────────────────────────────────────────────────── */

const assert = require("node:assert/strict");
const fs = require("node:fs");
const path = require("node:path");
const vm = require("node:vm");

const ROOT = path.resolve(__dirname, "..");
const authSource = fs.readFileSync(path.join(ROOT, "auth.js"), "utf8");
const htmlSource = fs.readFileSync(path.join(ROOT, "index.html"), "utf8");
const appSource = fs.readFileSync(path.join(ROOT, "app.js"), "utf8");

async function verifyProductionLogout() {
  const removed = [];
  const storage = new Map([
    ["paceTracker_lastRoom", "pace-room-2"],
    ["paceTracker_visitContext", '{"123":{"room":"pace-room-2"}}'],
    ["unrelated-setting", "keep-me"]
  ]);
  const redirects = [];
  const context = vm.createContext({
    console,
    APP_MODE: "production",
    CONFIG: { STORAGE_KEYS: { LAST_ROOM: "paceTracker_lastRoom", VISIT_CONTEXT: "paceTracker_visitContext" } },
    window: { location: { origin: "https://pace-room-tracker.vercel.app" } },
    localStorage: {
      removeItem(key) { removed.push(key); storage.delete(key); },
      getItem(key) { return storage.get(key) || null; }
    },
    msal: {}
  });
  vm.runInContext(authSource, context, { filename: "auth.js" });
  const auth = vm.runInContext("AUTH", context);
  auth.account = { username: "pilot@example.test" };
  auth.staffId = "staff-1";
  auth.staffName = "Pilot User";
  auth.role = "Behavior Specialist";
  auth._client = {
    logoutRedirect(request) {
      redirects.push(request);
      return Promise.resolve();
    }
  };

  await auth.logout();

  assert.deepEqual(removed, ["paceTracker_lastRoom", "paceTracker_visitContext"]);
  assert.equal(storage.get("unrelated-setting"), "keep-me");
  assert.equal(auth.account, null);
  assert.equal(auth.staffId, null);
  assert.equal(auth.staffName, null);
  assert.equal(auth.role, null);
  assert.deepEqual(JSON.parse(JSON.stringify(redirects)), [{
    account: { username: "pilot@example.test" },
    postLogoutRedirectUri: "https://pace-room-tracker.vercel.app"
  }]);
}

function verifyAuthenticatedShellWiring() {
  assert.match(htmlSource, /class="auth-toolbar hidden"[^>]*id="authToolbar"/);
  assert.match(htmlSource, /id="signedInIndicator"/);
  assert.match(htmlSource, /id="signOutBtn"[^>]*>Sign Out<\/button>/);
  assert.match(appSource, /authToolbar\.classList\.remove\("hidden"\)/);
  assert.match(appSource, /AUTH\.logout\(\)/);
  assert.match(appSource, /if \(APP_MODE === "demo" \|\| !AUTH\.account\) return/);
}

Promise.all([verifyProductionLogout(), verifyAuthenticatedShellWiring()])
  .then(() => console.log("Sign-out tests passed."))
  .catch(error => {
    console.error(error);
    process.exitCode = 1;
  });
