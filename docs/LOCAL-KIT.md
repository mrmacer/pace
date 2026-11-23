# Local kit (`~/.pace/`) — how to rebuild it

R-E keeps local verification tooling **outside** this repository, in
`~/.pace/` (on Windows: `C:\Users\<you>\.pace\`). Files can't be copied
between PCs, so this page is the source for rebuilding the kit exactly. Any
agent that finds `~/.pace/` missing should rebuild it from here. Whenever the
kit changes, update this file in the same piece of work.

Nothing in `~/.pace/` is ever committed. The project code never depends on it.

## Layout

```
~/.pace/
  demo/
    tour.js                   ONE command: sandbox copy → every test file → visible browser tour
  tests/
    student-history.test.js   Node-only checks for the STUDENT HISTORY PATCH
  tools/
    doc-audit.js              checks the house documentation style on every function
  sandbox/
    tools/                    npm folder holding playwright-core@1.48.2
    work/                     throwaway copy of the project under test (recreated every run)
    report/                   tour.log (last line starts "DONE:") and shots/ (recreated every run)
```

## Requirements

- Node.js 20 or newer (built with 24.21.0).
- Microsoft Edge installed. Playwright drives it with `channel: "msedge"`, so
  no browser download is needed.

## Rebuild steps (Git Bash or any POSIX shell)

```bash
mkdir -p ~/.pace/demo ~/.pace/tests ~/.pace/tools ~/.pace/sandbox/tools
cd ~/.pace/sandbox/tools && npm init -y >/dev/null && npm install playwright-core@1.48.2 --no-audit --no-fund
# then create ~/.pace/demo/tour.js, ~/.pace/tests/student-history.test.js and
# ~/.pace/tools/doc-audit.js with the exact contents below
```

## Using it

```bash
node ~/.pace/demo/tour.js            # visible Edge window, paced so a person can watch (~3 min)
node ~/.pace/demo/tour.js --headless # same checks, no window
```

The tour finds the repo (`PACE_REPO`, else `~/Projects/pace`, `~/pace/pace`,
`~/pace`), copies it to `~/.pace/sandbox/work`, and runs the repo's
`tests/*.test.js` and `~/.pace/tests/*.test.js` against that copy, and runs
`~/.pace/tools/doc-audit.js` on the app scripts and the kit scripts (every
function must carry the house documentation block). It then
serves the copy on a random localhost port and drives a fresh, throwaway Edge
profile in demo mode (fake students, zero Microsoft traffic, which it
checks). A caption in the corner says what each step is doing. Results are in
`~/.pace/sandbox/report/tour.log` and screenshots in
`~/.pace/sandbox/report/shots/`.

The agent runs this itself as a background task. It watches for the `DONE:`
line, looks at the screenshots, and reports. R-E is never asked to click
through.

## `~/.pace/demo/tour.js`

~~~~js
///////////////////////////////////////////////////////////////////////////////////////////////
// Author: R-E Miller
// Creation Date: November 23, 2026
// Filename: tour.js
// Purpose: Visible, sandboxed browser tour of PACE Room Tracker (local kit only; never committed).
//          Usage: node ~/.pace/demo/tour.js [--headless] [--fast]. Copies the repo (PACE_REPO, else
//          ~/Projects/pace, ~/pace/pace, ~/pace) into ~/.pace/sandbox/work, runs the repo's tests and
//          ~/.pace/tests/*.test.js against that copy, then drives a VISIBLE Microsoft Edge window
//          (fresh throwaway profile, demo mode only) through the features like a user would. The repo
//          folder itself is never touched. Report: ~/.pace/sandbox/report/tour.log (last line starts
//          "DONE:") and shots/.
///////////////////////////////////////////////////////////////////////////////////////////////
const fs = require("node:fs");
const os = require("node:os");
const path = require("node:path");
const http = require("node:http");
const { spawnSync } = require("node:child_process");

const HOME = os.homedir();
const KIT = path.join(HOME, ".pace");
const WORK = path.join(KIT, "sandbox", "work");
const REPORT = path.join(KIT, "sandbox", "report");
const SHOTS = path.join(REPORT, "shots");
const LOG = path.join(REPORT, "tour.log");
const HEADLESS = process.argv.includes("--headless");
const FAST = process.argv.includes("--fast");
const { chromium } = require(path.join(KIT, "sandbox", "tools", "node_modules", "playwright-core"));

fs.rmSync(REPORT, { recursive: true, force: true });
fs.mkdirSync(SHOTS, { recursive: true });
///////////////////////////////////////////////////////////////////////////////////////////////
// Function Name: log
// Description: Prints one report line and appends it to tour.log.
// Parameters: string line - text to record - input
///////////////////////////////////////////////////////////////////////////////////////////////
const log = line => { console.log(line); fs.appendFileSync(LOG, line + "\n"); };

///////////////////////////////////////////////////////////////////////////////////////////////
// Function Name: findRepo
// Description: Finds the PACE repo folder (PACE_REPO, then the usual checkout locations).
// Parameters: none
///////////////////////////////////////////////////////////////////////////////////////////////
function findRepo() {
  const candidates = [process.env.PACE_REPO, path.join(HOME, "Projects", "pace"), path.join(HOME, "pace", "pace"), path.join(HOME, "pace")];
  const repo = candidates.find(dir => dir && fs.existsSync(path.join(dir, "index.html")) && fs.existsSync(path.join(dir, "pace-data.js")));
  if (!repo) throw new Error("PACE repo not found: set PACE_REPO");
  return repo;
}

///////////////////////////////////////////////////////////////////////////////////////////////
// Function Name: makeSandboxCopy
// Description: Recreates ~/.pace/sandbox/work as a copy of the repo, without .git, node_modules,
//              .vercel, or .env files.
// Parameters: string repo - path of the repo to copy - input
///////////////////////////////////////////////////////////////////////////////////////////////
function makeSandboxCopy(repo) {
  fs.rmSync(WORK, { recursive: true, force: true });
  const skip = new Set([".git", "node_modules", ".vercel"]);
  fs.cpSync(repo, WORK, {
    recursive: true,
    filter: src => {
      const first = path.relative(repo, src).split(path.sep)[0];
      return !first || (!skip.has(first) && !first.startsWith(".env"));
    }
  });
}

///////////////////////////////////////////////////////////////////////////////////////////////
// Function Name: runTestFiles
// Description: Runs every repo and kit *.test.js file against the sandbox copy and logs PASS/FAIL.
// Parameters: none
///////////////////////////////////////////////////////////////////////////////////////////////
function runTestFiles() {
  const files = [];
  const repoTests = path.join(WORK, "tests");
  if (fs.existsSync(repoTests)) fs.readdirSync(repoTests).filter(f => f.endsWith(".test.js")).sort().forEach(f => files.push(path.join(repoTests, f)));
  const kitTests = path.join(KIT, "tests");
  if (fs.existsSync(kitTests)) fs.readdirSync(kitTests).filter(f => f.endsWith(".test.js")).sort().forEach(f => files.push(path.join(kitTests, f)));
  const failed = [];
  for (const file of files) {
    const run = spawnSync(process.execPath, [file, WORK], { cwd: WORK, encoding: "utf8", timeout: 120000 });
    const label = file.startsWith(WORK) ? `repo  ${path.relative(WORK, file)}` : `kit   ${path.relative(KIT, file)}`;
    if (run.status === 0) log(`  PASS ${label}`);
    else {
      failed.push(path.basename(file));
      const why = `${run.stderr || ""}${run.stdout || ""}`.split(/\r?\n/).find(l => /Error|assert/i.test(l)) || "exit " + run.status;
      log(`  FAIL ${label} — ${why.trim().slice(0, 160)}`);
    }
  }
  // House documentation style: every app script (sandbox copy) and every kit script.
  const audit = path.join(KIT, "tools", "doc-audit.js");
  const targets = [
    ["app scripts", [WORK]],
    ["kit demo/tour.js", [path.join(KIT, "demo"), "tour.js"]],
    ["kit tools/doc-audit.js", [path.join(KIT, "tools"), "doc-audit.js"]],
    ["kit tests", [path.join(KIT, "tests"), ...fs.readdirSync(path.join(KIT, "tests")).filter(f => f.endsWith(".js"))]]
  ];
  let checks = files.length;
  for (const [label, args] of targets) {
    checks++;
    const run = spawnSync(process.execPath, [audit, ...args], { encoding: "utf8" });
    if (run.status === 0) log(`  PASS docs  ${label}`);
    else {
      failed.push(`docs: ${label}`);
      log(`  FAIL docs  ${label} — ${(run.stdout || "").split(/\r?\n/).filter(Boolean).slice(0, 3).join("; ").slice(0, 200)}`);
    }
  }
  return { total: checks, failed };
}

///////////////////////////////////////////////////////////////////////////////////////////////
// Function Name: serve
// Description: Serves a folder over HTTP on a random localhost port (no caching).
// Parameters: string root - folder to serve - input
///////////////////////////////////////////////////////////////////////////////////////////////
function serve(root) {
  const TYPES = { ".html": "text/html; charset=utf-8", ".js": "text/javascript", ".css": "text/css", ".png": "image/png", ".webmanifest": "application/manifest+json" };
  const server = http.createServer((req, res) => {
    const p = decodeURIComponent(new URL(req.url, "http://x").pathname);
    const file = path.join(root, p === "/" ? "index.html" : p);
    if (!file.startsWith(root) || !fs.existsSync(file) || fs.statSync(file).isDirectory()) { res.writeHead(404); return res.end(); }
    res.writeHead(200, { "Content-Type": TYPES[path.extname(file)] || "application/octet-stream", "Cache-Control": "no-store" });
    fs.createReadStream(file).pipe(res);
  });
  return new Promise(resolve => server.listen(0, "127.0.0.1", () => resolve(server)));
}

// ── seed data (fake demo students only) ────────────────────────────────────
///////////////////////////////////////////////////////////////////////////////////////////////
// Function Name: local
// Description: Formats a Date as a local YYYY-MM-DD string (the stored date format).
// Parameters: Date d - date to format - input
///////////////////////////////////////////////////////////////////////////////////////////////
const local = d => `${d.getFullYear()}-${String(d.getMonth() + 1).padStart(2, "0")}-${String(d.getDate()).padStart(2, "0")}`;
///////////////////////////////////////////////////////////////////////////////////////////////
// Function Name: day
// Description: Returns the local YYYY-MM-DD date a number of days from today.
// Parameters: number n - days from today (negative = earlier) - input
///////////////////////////////////////////////////////////////////////////////////////////////
const day = n => { const d = new Date(); d.setDate(d.getDate() + n); return local(d); };
///////////////////////////////////////////////////////////////////////////////////////////////
// Function Name: hhmmAgo
// Description: Returns the 24-hour HH:MM time a number of minutes ago (the stored time format).
// Parameters: number minutes - minutes before now - input
///////////////////////////////////////////////////////////////////////////////////////////////
const hhmmAgo = minutes => { const d = new Date(Date.now() - minutes * 60000); return `${String(d.getHours()).padStart(2, "0")}:${String(d.getMinutes()).padStart(2, "0")}`; };
const TODAY = day(0), YESTERDAY = day(-1);
///////////////////////////////////////////////////////////////////////////////////////////////
// Function Name: visit
// Description: Builds one fake demo visit row, keyed by SharePoint display names.
// Parameters: string id - demo row id - input
//             string room - room label - input
//             string sid - Student ID - input
//             string name - student name - input
//             string date - YYYY-MM-DD - input
//             string tin - Time In HH:MM - input
//             string tout - Time Out HH:MM, or "" for an open visit - input
//             string reason - Reason text - input
//             string support - Intervention Used text - input
//             boolean|null scm - SCM answer - input
//             string notes - Notes text - input
//             string staff - Behavior Specialist(s) - input
//             string teacher - Teacher Came From - input
///////////////////////////////////////////////////////////////////////////////////////////////
const visit = (id, room, sid, name, date, tin, tout, reason, support, scm, notes, staff, teacher) => ({
  id, demo: true, Room: room, "Student ID": sid, Student: name, Date: date, "Time In": tin, "Time Out": tout,
  Reason: reason, "Intervention Used": support, "SCM Used": scm, Notes: notes, "Staff Member": staff,
  "Teacher Came From": teacher, createdAt: new Date().toISOString()
});
const SEED = { paceVisits: [
  visit("s1", "PACE Room 1", "SIM-001", "Alex R.", day(-7), "10:05", "10:30", "Disruption", "Calm space", false, "A week ago: needed a quiet space after group work. Back to class in 15 minutes.", "Dana Fielding", "Mrs. Ashford"),
  visit("s2", "PACE Room 1", "SIM-001", "Alex R.", day(-3), "09:00", "09:25", "Elopement", "De-escalation conversation, Break / reset", true, "Left the classroom without permission.\nUsed SCM briefly, then a walk and a reset break.", "Marcus Webb, Priya Anand", "Mr. Bellamy"),
  visit("s3", "PACE Room 2", "SIM-001", "Alex R.", YESTERDAY, "13:10", "13:40", "Peer conflict", "Restorative conversation", false, "Yesterday: argument with a peer at lunch. Restorative conversation went well.", "Priya Anand", "Ms. Castillo"),
  visit("s4", "PACE Room 1", "SIM-003", "Taylor S.", YESTERDAY, "10:00", "10:15", "Needs a Break", "Break / reset", false, "Asked for a break. Short reset, back to class.", "Dana Fielding", "Mr. Delgado"),
  visit("s5", "PACE Room 1", "SIM-005", "Morgan T.", YESTERDAY, "14:20", "14:50", "Verbal aggression", "De-escalation conversation, Parent/guardian contact", false, "Called parent after the visit.", "Marcus Webb", "Mrs. Ashford"),
  visit("s6", "PACE Room 1", "SIM-001", "Alex R.", TODAY, hhmmAgo(150), hhmmAgo(130), "Defiance / refusal", "Modified task", false, "Today: refused a writing task; a modified task helped.", "Dana Fielding", "Mrs. Ashford"),
  visit("s7", "PACE Room 1", "SIM-002", "Jordan M.", TODAY, hhmmAgo(12), "", "Disruption", "", null, "Live notes in progress.", "Marcus Webb", "Mr. Delgado"),
  visit("s8", "Other", "SIM-002", "Jordan M.", day(-2), "11:00", "11:20", "Emotional dysregulation", "Sensory support", false, "Wiggle Room visit two days ago.", "Priya Anand", "Mr. Bellamy"),
  // Older visits (a year+ ago), shaped like older rows: no room; one with the student's ID, one
  // with none. Plus a student who isn't on the demo roster at all (findable by name only).
  { ...visit("y1", "", "SIM-001", "Alex R.", day(-365), "10:00", "10:20", "Disruption", "Calm space", false, "Needed a quiet space after recess.", "Dana Fielding", "Mrs. Ashford"), Room: undefined },
  { ...visit("y2", "", "", "Alex R.", day(-400), "13:00", "13:30", "Peer conflict", "", null, "Argument over a game.", "", ""), Room: undefined, "Student ID": undefined },
  { ...visit("y3", "", "", "Sam Avery", day(-380), "09:10", "09:40", "Needs a Break", "Break / reset", false, "Asked for a break before a test.", "Marcus Webb", "Mr. Delgado"), Room: undefined, "Student ID": undefined },
  // Ten more completed visits today in Room 1, so today has 11 (beyond the old 10-card cap).
  ...Array.from({ length: 10 }, (_, i) => visit(`t${i}`, "PACE Room 1", "SIM-003", "Taylor S.", TODAY, hhmmAgo(600 - i * 30), hhmmAgo(590 - i * 30), "Needs a Break", "Break / reset", false, `Short break ${i + 1}.`, "Dana Fielding", "Mr. Delgado"))
] };
const A_YEAR_AGO = day(-365), EARLIEST = day(-400);

(async () => {
  const repo = findRepo();
  const [ty, tm, td] = TODAY.split("-");
  log(`PACE tour ${tm}-${td}-${ty} ${new Date().toLocaleTimeString("en-US")} — repo ${repo}`);
  log(`Sandbox copy: ${WORK}`);
  makeSandboxCopy(repo);

  log("Test files (run against the sandbox copy):");
  const tests = runTestFiles();

  const server = await serve(WORK);
  const BASE = `http://127.0.0.1:${server.address().port}/`;
  log(`Browser tour (${HEADLESS ? "headless" : "visible"} Edge, throwaway profile, demo mode) at ${BASE}?demo=1`);
  const browser = await chromium.launch({
    channel: "msedge", headless: HEADLESS, slowMo: FAST || HEADLESS ? 0 : 120,
    args: ["--window-size=1220,960", "--window-position=40,30"]
  });
  const context = await browser.newContext({ viewport: { width: 1180, height: 820 } });
  // Records every sideways scroll of #app (screen + focused element at that moment), so a
  // "shifted sideways" failure says what caused it.
  await context.addInitScript(() => {
    window.__appScrolls = [];
    document.addEventListener("DOMContentLoaded", () => {
      const app = document.getElementById("app");
      if (!app) return;
      app.addEventListener("scroll", () => {
        const focus = document.activeElement;
        window.__appScrolls.push(`${app.scrollLeft}px on "${document.querySelector(".screen.active")?.dataset.screen}" (focus: ${focus?.id || focus?.className || focus?.tagName})`);
      });
    });
  });
  const page = await context.newPage();
  const problems = [], microsoft = [];
  page.on("pageerror", e => problems.push(`page error: ${e}`));
  page.on("console", m => { if (m.type() === "error" && !/Failed to load resource/.test(m.text())) problems.push(`console: ${m.text()}`); });
  page.on("response", r => { if (r.status() >= 400 && !/favicon/.test(r.url())) problems.push(`${r.status()} ${r.url()}`); });
  page.on("request", r => { if (/microsoft|msal|sharepoint|office|live\.com|windows\.net/i.test(new URL(r.url()).hostname)) microsoft.push(r.url()); });

  let passed = 0, failed = 0, shotNo = 0, stoppedAt = "";
  ///////////////////////////////////////////////////////////////////////////////////////////////
  // Function Name: pause
  // Description: Waits so a person watching can follow (shortened in --fast / --headless).
  // Parameters: number ms - pause length in milliseconds - input
  ///////////////////////////////////////////////////////////////////////////////////////////////
  const pause = ms => page.waitForTimeout(FAST || HEADLESS ? Math.min(ms, 150) : ms);
  ///////////////////////////////////////////////////////////////////////////////////////////////
  // Function Name: expect
  // Description: Records one check as ok or FAIL; a FAIL stops the current step.
  // Parameters: boolean cond - whether the check passed - input
  //             string msg - what was checked - input
  ///////////////////////////////////////////////////////////////////////////////////////////////
  const expect = (cond, msg) => {
    if (cond) { passed++; log(`    ok   ${msg}`); return; }
    failed++; log(`    FAIL ${msg}`);
    const err = new Error(msg); err.fromExpect = true; throw err;
  };
  // Screens slide horizontally (220 ms). Wait until every screen is either fully in view or fully
  // off to one side, so the tour never taps or types into a screen that is still moving.
  // (Checking positions alone isn't enough: on a slide's first frame the incoming screen still
  // reads as fully off to the side. So also wait until no CSS transition is running or pending.)
  ///////////////////////////////////////////////////////////////////////////////////////////////
  // Function Name: settled
  // Description: Waits until no screen is sliding and no CSS transition is running.
  // Parameters: none
  ///////////////////////////////////////////////////////////////////////////////////////////////
  const settled = async () => {
    await page.evaluate(() => new Promise(r => requestAnimationFrame(() => requestAnimationFrame(r))));
    await page.waitForFunction(() =>
      document.getAnimations().every(a => !(a instanceof CSSTransition) || (a.playState !== "running" && !a.pending)) &&
      [...document.querySelectorAll(".screen")].every(s => {
        const x = new DOMMatrix(getComputedStyle(s).transform).m41;
        return Math.abs(x) < 1 || Math.abs(Math.abs(x) - s.offsetWidth) < 1;
      }), null, { timeout: 6000 });
  };
  // Every tap/type/fill first waits for slides to finish, like a person would (no hand audits).
  for (const action of ["click", "type", "fill"]) {
    const raw = page[action].bind(page);
    page[action] = async (...args) => { await settled(); return raw(...args); };
  }
  ///////////////////////////////////////////////////////////////////////////////////////////////
  // Function Name: notShifted
  // Description: Checks the app isn't scrolled sideways, reporting what caused it if it is.
  // Parameters: string when - label for the report line - input
  ///////////////////////////////////////////////////////////////////////////////////////////////
  const notShifted = async when => {
    const state = await page.evaluate(() => ({ left: document.getElementById("app").scrollLeft, scrolls: window.__appScrolls || [] }));
    expect(state.left === 0, `screen is not shifted sideways (${when})${state.left ? ` — ${state.scrolls.join("; ")}` : ""}`);
  };
  ///////////////////////////////////////////////////////////////////////////////////////////////
  // Function Name: shot
  // Description: Takes a numbered screenshot after the screen settles (with a sideways check).
  // Parameters: string name - screenshot file label - input
  //             Object options - { check } false skips the settle/shift check - input
  ///////////////////////////////////////////////////////////////////////////////////////////////
  const shot = async (name, { check = true } = {}) => {
    if (check) { await settled(); await page.waitForTimeout(250); await notShifted(name); }
    const file = path.join(SHOTS, `${String(++shotNo).padStart(2, "0")}-${name}.png`);
    await page.screenshot({ path: file });
    log(`    shot ${path.basename(file)}`);
  };
  ///////////////////////////////////////////////////////////////////////////////////////////////
  // Function Name: caption
  // Description: Shows the current tour step in a small caption in the demo banner's corner.
  // Parameters: string text - step description - input
  ///////////////////////////////////////////////////////////////////////////////////////////////
  const caption = async text => {
    await page.evaluate(t => {
      let el = document.getElementById("tourCaption");
      if (!el) {
        el = document.createElement("div");
        el.id = "tourCaption";
        // Top-right, inside the demo banner's empty right side, so it never covers app content.
        el.style.cssText = "position:fixed;right:8px;top:5px;z-index:99999;max-width:300px;padding:5px 10px;border-radius:8px;background:rgba(15,23,42,.9);color:#fff;font:600 12px/1.3 system-ui,sans-serif;pointer-events:none;white-space:nowrap;overflow:hidden;text-overflow:ellipsis";
        document.body.appendChild(el);
      }
      el.textContent = "Tour: " + t;
    }, text).catch(() => {});
  };
  ///////////////////////////////////////////////////////////////////////////////////////////////
  // Function Name: screenIs
  // Description: Waits until the named screen is active, settled, and no toast covers it.
  // Parameters: string name - data-screen value - input
  ///////////////////////////////////////////////////////////////////////////////////////////////
  const screenIs = async name => {
    try {
      await page.waitForFunction(n => document.querySelector(".screen.active")?.dataset.screen === n, name, { timeout: 6000 });
      await settled();
      // A toast (e.g. "… is now in PACE.") sits over the bottom action button for ~2.6 s. Tapping
      // through it makes Playwright retry from other scroll positions, so wait for it like a person.
      await page.waitForFunction(() => document.getElementById("toast")?.classList.contains("hidden"), null, { timeout: 6000 });
      return true;
    } catch { return false; }
  };
  ///////////////////////////////////////////////////////////////////////////////////////////////
  // Function Name: text
  // Description: Returns the trimmed text of the first matching element.
  // Parameters: string selector - CSS selector - input
  ///////////////////////////////////////////////////////////////////////////////////////////////
  const text = async selector => (await page.textContent(selector) || "").trim();
  ///////////////////////////////////////////////////////////////////////////////////////////////
  // Function Name: visible
  // Description: Reports whether the first matching element is visible.
  // Parameters: string selector - CSS selector - input
  ///////////////////////////////////////////////////////////////////////////////////////////////
  const visible = selector => page.isVisible(selector);
  ///////////////////////////////////////////////////////////////////////////////////////////////
  // Function Name: count
  // Description: Counts the elements matching a selector.
  // Parameters: string selector - CSS selector - input
  ///////////////////////////////////////////////////////////////////////////////////////////////
  const count = selector => page.locator(selector).count();
  ///////////////////////////////////////////////////////////////////////////////////////////////
  // Function Name: step
  // Description: Runs one named tour step; after a failure, later steps are skipped and logged.
  // Parameters: string title - step name - input
  //             Function body - async step actions - input
  ///////////////////////////////////////////////////////////////////////////////////////////////
  const step = async (title, body) => {
    if (stoppedAt) { log(`  SKIP ${title}`); return; }
    log(`  STEP ${title}`);
    await caption(title);
    try { await body(); await pause(700); }
    catch (err) {
      stoppedAt = title;
      if (!err.fromExpect) { failed++; log(`    FAIL ${String(err.message).split("\n")[0]}`); }
      await shot("FAILED", { check: false }).catch(() => {});
    }
  };

  await step("Open the app in demo mode with sample visits", async () => {
    await page.goto(`${BASE}?demo=1`);
    await page.evaluate(seed => { localStorage.clear(); localStorage.setItem("paceRoomTrackerDemoData", JSON.stringify(seed)); }, SEED);
    await page.goto(`${BASE}?demo=1`);
    expect(await screenIs("home"), "home screen (room picker) shows");
    expect(await visible("#demoBanner"), "demo banner is showing (no SharePoint)");
    await caption("Pick PACE Room 1");
    await pause(600);
    await page.click('.room-card[data-room="pace-room-1"]');
    expect(await screenIs("room"), "PACE Room 1 opens");
    await page.waitForSelector("#currentlyInPace .current-card");
    expect((await text("#currentlyInPace")).includes("Jordan M."), "Jordan M. is Currently in PACE");
    expect(await count("#currentlyInPace [data-history-item-id]") === 1, "card has a Previous visits button");
    await shot("room-currently-in-pace");
  });

  await step("Previous visits for a student in PACE", async () => {
    await page.click("#currentlyInPace [data-history-item-id]");
    expect(await screenIs("history"), "Visit History opens");
    await page.waitForSelector("#historyList .history-card");
    expect(await text("#historyStudentName") === "Jordan M.", "title is Jordan M.");
    expect(/^2 visits on 2 days/.test(await text("#historySummary")), `summary: ${await text("#historySummary")}`);
    expect((await text("#historyList")).includes("In progress"), "today's open visit shows In progress");
    expect((await text("#historyList")).includes("Other"), "the 'Other' location visit is listed");
    expect(await count("#historyList .visit-details:not(.hidden)") === 0, "details start collapsed");
    expect(!(await text("#historyList")).includes("Wiggle Room visit"), "notes aren't on the page before Details");
    await pause(600);
    await page.locator("#historyList [data-history-details]").nth(1).click();
    const details = await text('#historyList [data-history-panel="1"]');
    expect(details.includes("Wiggle Room visit two days ago.") && details.includes("Sensory support"), "Details shows notes and support");
    await shot("history-jordan-details");
    await page.click("#historyBackBtn");
    expect(await screenIs("room"), "Back returns to the room");
  });

  await step("Recent Activity starts on today", async () => {
    await page.click('[data-screen="room"] [data-nav="recent"]');
    expect(await screenIs("recent"), "Recent Activity opens");
    await page.waitForSelector("#recentList .recent-card");
    expect(await text("#recentScope") === "Today · PACE Room 1", `scope: ${await text("#recentScope")}`);
    expect(await page.inputValue("#recentDateInput") === TODAY, "date field shows today");
    expect(await count("#recentList .recent-card") === 11, `all 11 of today's visits listed, no 10-visit cap (${await count("#recentList .recent-card")})`);
    expect(await count("#recentList .recent-edit-btn") === 11 && await count("#recentList .recent-delete-btn") === 11, "every one of today's cards can be edited and deleted");
    expect(await page.isDisabled("#recentNextDayBtn"), "Next day is off on today");
    expect(await visible("#recentTodayHint") && !(await visible("#recentTodayBtn")), "today hint shown, no Back to today");
    await shot("recent-today");
  });

  await step("Prev day: yesterday is view-only", async () => {
    await page.click("#recentPrevDayBtn");
    await page.waitForFunction(d => document.getElementById("recentDateInput").value === d && document.querySelector("#recentList .recent-card"), YESTERDAY);
    const scope = await text("#recentScope");
    expect(!scope.startsWith("Today") && scope.endsWith("· PACE Room 1"), `scope: ${scope}`);
    const [yy, ym, yd] = YESTERDAY.split("-");
    expect(/^[A-Z][a-z]+day, \d{2}-\d{2}-\d{4} · /.test(scope) && scope.includes(`${ym}-${yd}-${yy}`), `date shown in US format MM-DD-YYYY: ${scope}`);
    expect(await count("#recentList .recent-card") === 2, "both Room 1 visits from yesterday listed");
    const list = await text("#recentList");
    expect(list.includes("Taylor S.") && list.includes("Morgan T.") && !list.includes("Alex R."), "only Room 1 students (Alex was in Room 2)");
    expect(await count("#recentList .recent-edit-btn, #recentList .recent-delete-btn") === 0, "no Edit or Delete on a past day");
    expect(!list.includes("Called parent"), "no notes on Recent cards");
    expect(await visible("#recentPastHint") && await visible("#recentTodayBtn"), "view-only hint and Back to today shown");
    expect(!(await page.isDisabled("#recentNextDayBtn")), "Next day is available");
    await shot("recent-yesterday");
  });

  await step("Pick any date with the date field", async () => {
    await page.fill("#recentDateInput", day(-7));
    await page.dispatchEvent("#recentDateInput", "change");
    await page.waitForFunction(() => document.querySelector("#recentList .recent-card"));
    expect((await text("#recentList")).includes("Alex R.") && await count("#recentList .recent-card") === 1, "a week ago shows Alex's visit");
    await pause(500);
    await page.click("#recentPrevDayBtn");
    await page.waitForFunction(() => /No completed PACE visits on/.test(document.getElementById("recentList").textContent));
    expect(true, "an empty day says so");
    await shot("recent-empty-day");
    await page.fill("#recentDateInput", A_YEAR_AGO);
    await page.dispatchEvent("#recentDateInput", "change");
    await page.waitForFunction(() => /Alex R\./.test(document.getElementById("recentList").textContent));
    expect((await text("#recentScope")).endsWith("· All PACE rooms"), `a day a year ago shows (older rows have no room): ${await text("#recentScope")}`);
    expect(await count("#recentList .recent-edit-btn, #recentList .recent-delete-btn") === 0, "older day is view-only");
    await shot("recent-a-year-ago");
    await page.fill("#recentDateInput", day(5));
    await page.dispatchEvent("#recentDateInput", "change");
    await page.waitForFunction(d => document.getElementById("recentDateInput").value === d, TODAY);
    expect(true, "a future date snaps back to today");
    await page.waitForSelector("#recentList .recent-edit-btn");
  });

  await step("Visit history from a Recent card", async () => {
    await page.locator("#recentList .recent-card", { hasText: "Alex R." }).locator("[data-recent-history]").click();
    expect(await screenIs("history"), "Visit History opens");
    await page.waitForSelector("#historyList .history-card");
    const summary = await text("#historySummary");
    expect(/^6 visits on 6 days/.test(summary), `summary: ${summary}`);
    const [ey, em, ed] = EARLIEST.split("-");
    const earliest = `${em}-${ed}-${ey}`; // US MM-DD-YYYY, as the app shows it
    expect(summary.endsWith(`since ${earliest}`), "history reaches back to the oldest visit (400 days ago)");
    expect(await count("#historyList .history-day") === 6, "grouped into 6 days, newest first");
    expect((await page.locator("#historyList .history-day-label").first().textContent()).trim().endsWith("· Today"), "first group is today");
    expect((await text("#historyList")).includes("quiet space after recess") === false, "notes stay hidden until Details");
    const list = await text("#historyList");
    expect(list.includes("PACE Room 1") && list.includes("PACE Room 2"), "visits from both rooms");
    expect(!list.includes("Taylor") && !list.includes("Morgan"), "no other students");
    await shot("history-alex");
    await page.locator("#historyList [data-history-details]").nth(2).click();
    const panel = '#historyList [data-history-panel="2"]';
    const details = await text(panel);
    expect(details.includes("Left the classroom without permission.") && details.includes("Marcus Webb") && details.includes("Priya Anand"), "Details: notes and both specialists");
    expect(/SCM\s*Yes/.test(details) && details.includes("Mr. Bellamy"), "Details: SCM Yes, teacher came from");
    const lines = await page.locator(`${panel} .visit-details-notes-text`).evaluate(el => el.innerText.split("\n").length);
    expect(lines === 2, "multi-line notes keep their line break");
    await page.locator(`${panel}`).scrollIntoViewIfNeeded();
    await shot("history-alex-details");
    await page.locator("#historyList [data-history-details]").nth(2).click();
    expect(!(await visible(panel)), "Hide details closes it");
    // Open the visit from a year ago to prove it's the real row.
    await page.locator("#historyList [data-history-details]").nth(4).click();
    const old = await text('#historyList [data-history-panel="4"]');
    expect(old.includes("Needed a quiet space after recess."), "the visit from a year ago opens with its notes");
    await page.locator('#historyList [data-history-panel="4"]').scrollIntoViewIfNeeded();
    await shot("history-alex-a-year-ago");
    await page.click("#historyBackBtn");
    expect(await screenIs("recent"), "Back returns to Recent");
  });

  await step("Look up a student's visit history", async () => {
    await page.click("#findStudentHistoryBtn");
    expect(await screenIs("historySearch"), "student lookup opens");
    await page.waitForSelector("#historyStudentGrid .student-card");
    await page.waitForFunction(() => document.querySelectorAll("#historyStudentGrid .student-card").length === 7);
    expect(true, "the 6 roster students plus 1 name found only on visits listed");
    await page.type("#historyStudentSearch", "tay", { delay: FAST ? 0 : 120 });
    expect(await count("#historyStudentGrid .student-card") === 1, "search narrows to Taylor S.");
    await shot("lookup-search");
    await page.click("#historyStudentGrid .student-card");
    expect(await screenIs("history"), "Visit History opens");
    await page.waitForSelector("#historyList .history-card");
    expect(await text("#historyStudentName") === "Taylor S." && /^11 visits on 2 days/.test(await text("#historySummary")), `Taylor's history: ${await text("#historySummary")}`);
    await page.click("#historyBackBtn");
    expect(await screenIs("historySearch"), "Back returns to the lookup");
    await page.fill("#historyStudentSearch", "");
    await page.type("#historyStudentSearch", "riley", { delay: FAST ? 0 : 120 });
    await page.click("#historyStudentGrid .student-card");
    expect(await screenIs("history"), "Visit History opens");
    await page.waitForFunction(() => /No PACE visits are recorded/.test(document.getElementById("historyList").textContent));
    expect(true, "a student with no visits says so");
    await page.click("#historyBackBtn");
    expect(await screenIs("historySearch"), "Back returns to the lookup");
    await page.click("#historySearchBackBtn");
    expect(await screenIs("recent"), "back on Recent");
  });

  await step("Search by name finds a student with only older visits", async () => {
    await page.click("#findStudentHistoryBtn");
    expect(await screenIs("historySearch"), "student lookup opens");
    await page.waitForFunction(() => document.querySelectorAll("#historyStudentGrid .student-card").length === 7);
    await page.type("#historyStudentSearch", "avery", { delay: FAST ? 0 : 120 });
    expect(await count("#historyStudentGrid .student-card") === 1, "name search finds Sam Avery");
    expect(await text("#historyStudentGrid") === "Sam Avery", "just the name, no extra labels");
    await shot("lookup-by-name");
    await page.click("#historyStudentGrid .student-card");
    expect(await screenIs("history"), "Visit History opens");
    await page.waitForSelector("#historyList .history-card");
    expect(await text("#historyStudentName") === "Sam Avery" && /^1 visit on 1 day/.test(await text("#historySummary")), `Sam Avery: ${await text("#historySummary")}`);
    await page.click("#historyList [data-history-details]");
    expect((await text('#historyList [data-history-panel="0"]')).includes("Asked for a break before a test."), "their visit opens with notes");
    await shot("history-older-visits");
    await page.click("#historyBackBtn");
    expect(await screenIs("historySearch"), "Back returns to the lookup");
    await page.click("#historySearchBackBtn");
    expect(await screenIs("recent"), "back on Recent");
  });

  await step("Bug fix: students in PACE stay listed after visiting Recent", async () => {
    await page.click('[data-screen="recent"] .btn-back');
    expect(await screenIs("room"), "back in the room");
    await notShifted("before the wait");
    await page.evaluate(() => { document.querySelector("#currentlyInPace .current-card").dataset.tourMark = "before"; });
    await caption("Waiting 32 seconds for the room's automatic refresh…");
    await page.waitForTimeout(32000);
    expect(await count('#currentlyInPace .current-card[data-tour-mark="before"]') === 0, "the room list refreshed on its own");
    expect((await text("#currentlyInPace")).includes("Jordan M."), "Jordan M. is still Currently in PACE");
    await shot("room-after-refresh");
  });

  await step("A new live visit shows up in history", async () => {
    await page.click("#startLiveVisitBtn");
    expect(await screenIs("specialist"), "Start Visit asks for the specialist");
    await page.waitForSelector('#specialistGrid [data-specialist="Dana Fielding"]');
    await page.click('#specialistGrid [data-specialist="Dana Fielding"]');
    await page.click("#specialistContinueBtn");
    expect(await screenIs("student"), "student search");
    await page.waitForSelector("#studentGrid .student-card");
    await page.type("#studentSearch", "casey", { delay: FAST ? 0 : 110 });
    await page.click("#studentGrid .student-card");
    expect(await screenIs("cameFrom"), "teacher came from");
    await page.waitForSelector('#cameFromGrid [data-came-from="Mrs. Ashford"]');
    await page.click('#cameFromGrid [data-came-from="Mrs. Ashford"]');
    expect(await screenIs("reason"), "reason");
    await page.click('#reasonChips .chip[data-value="Needs a Break"]');
    await page.click("#reasonNextBtn");
    expect(await screenIs("visitinfo"), "visit details");
    await page.fill("#visitTimeInInput", hhmmAgo(20));
    await page.fill("#liveVisitNotesInput", "Tour: started live.");
    await page.click("#visitInfoNextBtn");
    expect(await screenIs("room"), "visit started, back in the room");
    await page.waitForFunction(() => /Casey L\./.test(document.getElementById("currentlyInPace").textContent));
    expect(true, "Casey L. is Currently in PACE");
    await caption("Mark Casey's visit complete");
    await page.locator("#currentlyInPace .current-card", { hasText: "Casey L." }).locator(".current-card-exit").click();
    expect(await screenIs("exit"), "Mark Visit Complete");
    await page.click("#confirmExitBtn");
    expect(await screenIs("support"), "support");
    await page.click('#supportChips .chip[data-value="Calm space"]');
    await page.click("#supportNextBtn");
    expect(await screenIs("scm"), "SCM");
    await page.click("#scmNoBtn");
    expect(await screenIs("notes"), "notes");
    expect(await page.inputValue("#noteText") === "Tour: started live.", "live notes carried into completion");
    await page.click("#notesNextBtn");
    expect(await screenIs("confirm"), "confirm");
    await page.click("#saveEntryBtn");
    expect(await screenIs("room"), "completed, back in the room");
    await page.click('[data-screen="room"] [data-nav="recent"]');
    expect(await screenIs("recent"), "Recent Activity opens");
    await page.waitForFunction(() => /Casey L\./.test(document.getElementById("recentList").textContent));
    expect(await count("#recentList .recent-card") === 12, "today now shows all 12 visits");
    await page.locator("#recentList .recent-card", { hasText: "Casey L." }).locator("[data-recent-history]").click();
    expect(await screenIs("history"), "Visit History opens");
    await page.waitForSelector("#historyList .history-card");
    expect(/^1 visit on 1 day/.test(await text("#historySummary")), "Casey's history has the new visit");
    await page.click("#historyList [data-history-details]");
    const details = await text('#historyList [data-history-panel="0"]');
    expect(details.includes("Tour: started live.") && details.includes("Calm space") && details.includes("Needs a Break"), "its details are complete");
    await shot("history-new-visit");
    await page.click("#historyBackBtn");
    expect(await screenIs("recent"), "Back returns to Recent");
    await page.click('[data-screen="recent"] .btn-back');
    expect(await screenIs("room"), "back in the room");
  });

  await step("A visit logged after the fact for yesterday shows on that day", async () => {
    await page.click("#logCompletedVisitBtn");
    expect(await screenIs("specialist"), "Log Completed Visit asks for the specialist");
    await page.waitForSelector("#specialistGrid [data-specialist]");
    if (await page.isDisabled("#specialistContinueBtn")) await page.click('#specialistGrid [data-specialist="Dana Fielding"]');
    await page.click("#specialistContinueBtn");
    expect(await screenIs("student"), "student search");
    await page.waitForSelector("#studentGrid .student-card");
    await page.type("#studentSearch", "morgan", { delay: FAST ? 0 : 110 });
    await page.click("#studentGrid .student-card");
    expect(await screenIs("cameFrom"), "teacher came from");
    await page.waitForSelector('#cameFromGrid [data-came-from="Mr. Bellamy"]');
    await page.click('#cameFromGrid [data-came-from="Mr. Bellamy"]');
    expect(await screenIs("visitinfo"), "visit details");
    await page.fill("#visitDateInput", YESTERDAY);
    await page.fill("#visitTimeInInput", "08:00");
    await page.fill("#visitTimeOutInput", "08:20");
    await page.click("#visitInfoNextBtn");
    expect(await screenIs("reason"), "reason");
    await page.click('#reasonChips .chip[data-value="Disruption"]');
    await page.click("#reasonNextBtn");
    expect(await screenIs("support"), "support");
    await page.click('#supportChips .chip[data-value="Break / reset"]');
    await page.click("#supportNextBtn");
    expect(await screenIs("scm"), "SCM");
    await page.click("#scmYesBtn");
    expect(await screenIs("notes"), "notes");
    await page.fill("#noteText", "Tour: logged after the fact for yesterday.");
    await page.click("#notesNextBtn");
    expect(await screenIs("confirm"), "confirm");
    expect((await text("#confirmCard")).includes("PACE Room 1"), "confirm shows the room");
    await page.click("#saveEntryBtn");
    expect(await screenIs("student"), "saved, back to student search");
    await page.click('[data-screen="student"] .btn-back');
    expect(await screenIs("specialist"), "back to the specialist screen");
    await page.click('[data-screen="specialist"] .btn-back');
    expect(await screenIs("room"), "back in the room");
    await page.click('[data-screen="room"] [data-nav="recent"]');
    expect(await screenIs("recent"), "Recent Activity opens");
    await page.waitForSelector("#recentList .recent-card");
    await page.click("#recentPrevDayBtn");
    await page.waitForFunction(() => document.querySelectorAll("#recentList .recent-card").length === 3);
    expect(true, "yesterday now lists 3 visits");
    await shot("recent-yesterday-with-logged");
    await page.locator("#recentList .recent-card", { hasText: "Morgan T." }).first().locator("[data-recent-history]").click();
    expect(await screenIs("history"), "Visit History opens");
    await page.waitForSelector("#historyList .history-card");
    expect(/^2 visits on 1 day/.test(await text("#historySummary")), `Morgan: ${await text("#historySummary")}`);
    await page.locator("#historyList [data-history-details]").nth(1).click();
    expect((await text('#historyList [data-history-panel="1"]')).includes("Tour: logged after the fact for yesterday."), "the logged visit's notes show");
    await page.click("#historyBackBtn");
    expect(await screenIs("recent"), "back on Recent (still yesterday)");
  });

  await step("Today's visits can still be edited", async () => {
    await page.click("#recentTodayBtn");
    await page.waitForSelector("#recentList .recent-edit-btn");
    await page.locator("#recentList .recent-edit-btn").first().click();
    expect(await screenIs("visitinfo"), "Edit opens the visit");
    expect(await page.isDisabled("#visitDateInput") && await page.inputValue("#visitDateInput") === TODAY, "date is locked to today");
    await shot("edit-today");
    await page.click("#visitEditCancelBtn");
    expect(await screenIs("recent"), "Cancel edit returns to Recent");
  });

  await step("Phone width: no sideways scrolling", async () => {
    await page.setViewportSize({ width: 390, height: 844 });
    await page.waitForSelector("#recentList .recent-card");
    ///////////////////////////////////////////////////////////////////////////////////////////////
    // Function Name: overflow
    // Description: Measures sideways overflow (px) of one screen's content area at phone width.
    // Parameters: string screen - data-screen value - input
    ///////////////////////////////////////////////////////////////////////////////////////////////
    const overflow = async screen => page.evaluate(s => { const el = document.querySelector(`[data-screen="${s}"] .panel-inner`); return el.scrollWidth - el.clientWidth; }, screen);
    expect(await overflow("recent") <= 0, "Recent fits a phone");
    await shot("phone-recent");
    await page.locator("#recentList [data-recent-history]").first().click();
    expect(await screenIs("history"), "Visit History opens");
    await page.waitForSelector("#historyList .history-card");
    await page.locator("#historyList [data-history-details]").first().click();
    expect(await overflow("history") <= 0, "Visit History fits a phone");
    await shot("phone-history");
    await page.click("#historyBackBtn");
    await page.setViewportSize({ width: 1180, height: 820 });
  });

  await step("Nothing sent to Microsoft, no errors", async () => {
    expect(microsoft.length === 0, `no Microsoft requests (${microsoft.length})`);
    expect(problems.length === 0, problems.length ? `errors: ${problems.join(" | ")}` : "no page or console errors");
  });

  const total = passed + failed;
  await caption(stoppedAt ? `Stopped at: ${stoppedAt}` : `Finished: all ${total} checks passed`);
  if (!HEADLESS) await page.waitForTimeout(6000);
  await browser.close();
  server.close();
  const testLine = `test files ${tests.total - tests.failed.length}/${tests.total} passed${tests.failed.length ? ` (failed: ${tests.failed.join(", ")})` : ""}`;
  log(`DONE: browser checks ${passed}/${total} passed${stoppedAt ? ` (stopped at "${stoppedAt}")` : ""}; ${testLine}`);
  process.exit(stoppedAt || tests.failed.length ? 1 : 0);
})().catch(err => { log(`DONE: tour crashed — ${err.stack || err}`); process.exit(1); });
~~~~

## `~/.pace/tests/student-history.test.js`

~~~~js
///////////////////////////////////////////////////////////////////////////////////////////////
// Author: R-E Miller
// Creation Date: November 23, 2026
// Filename: student-history.test.js
// Purpose: Local-only checks for the STUDENT HISTORY PATCH (never committed). Usage:
//          node ~/.pace/tests/student-history.test.js <path-to-project-copy>. Covers the pure
//          VISIT_HISTORY rules, the Graph per-student and name-list reads, PACE_DATA in demo and
//          production, and read-only / privacy structure checks.
///////////////////////////////////////////////////////////////////////////////////////////////
const assert = require("node:assert/strict");
const fs = require("node:fs");
const path = require("node:path");
const vm = require("node:vm");

const ROOT = path.resolve(process.argv[2] || path.join(require("node:os").homedir(), ".pace", "sandbox", "work"));
///////////////////////////////////////////////////////////////////////////////////////////////
// Function Name: plain
// Description: Copies a value out of a vm sandbox as plain JSON so deep-equal works.
// Parameters: any x - value to copy - input
///////////////////////////////////////////////////////////////////////////////////////////////
const plain = x => JSON.parse(JSON.stringify(x));
///////////////////////////////////////////////////////////////////////////////////////////////
// Function Name: src
// Description: Reads one project file's source text from the copy under test.
// Parameters: string f - file name - input
///////////////////////////////////////////////////////////////////////////////////////////////
const src = f => fs.readFileSync(path.join(ROOT, f), "utf8");
const H = require(path.join(ROOT, "visit-history.js"));

///////////////////////////////////////////////////////////////////////////////////////////////
// Function Name: local
// Description: Formats a Date as a local YYYY-MM-DD string.
// Parameters: Date d - date to format - input
///////////////////////////////////////////////////////////////////////////////////////////////
const local = d => `${d.getFullYear()}-${String(d.getMonth() + 1).padStart(2, "0")}-${String(d.getDate()).padStart(2, "0")}`;
const now = new Date();
const TODAY = local(now);
const YESTERDAY = H.shiftDate(TODAY, -1);
const TOMORROW = H.shiftDate(TODAY, 1);

/* ── pure date rules ── */
assert.equal(H.shiftDate("2026-03-01", -1), "2026-02-28");
assert.equal(H.shiftDate("2028-03-01", -1), "2028-02-29");
assert.equal(H.shiftDate("2026-12-31", 1), "2027-01-01");
assert.equal(H.shiftDate("2026-03-08", 1), "2026-03-09", "DST weekend never skips a day");
assert.equal(H.shiftDate("2026-11-01", -1), "2026-10-31");
assert.equal(H.shiftDate("bad", 1), "");
assert.equal(H.isIsoDate("2026-02-30"), false);
assert.equal(H.clampViewDate(TOMORROW, TODAY), TODAY, "future clamps to today");
assert.equal(H.clampViewDate("", TODAY), TODAY);
assert.equal(H.clampViewDate("garbage", TODAY), TODAY);
assert.equal(H.clampViewDate(YESTERDAY, TODAY), YESTERDAY);
assert.ok(H.dayLimit() >= 1e6, "no daily cap, today included");

/* ── student matching: ID when both sides have one, else exact name ── */
///////////////////////////////////////////////////////////////////////////////////////////////
// Function Name: v
// Description: Builds a test visit row, overriding any fields given.
// Parameters: Object o - fields to override - input
///////////////////////////////////////////////////////////////////////////////////////////////
const v = (o = {}) => ({ id: "1", Student: "Alex R.", "Student ID": "SIM-001", Date: TODAY, "Time In": "09:00", "Time Out": "09:30", ...o });
assert.equal(H.belongsToStudent(v(), { id: "SIM-001", name: "Alex R." }), true);
assert.equal(H.belongsToStudent(v({ "Student ID": "SIM-999" }), { id: "SIM-001", name: "Alex R." }), false, "same name, different ID = different student");
assert.equal(H.belongsToStudent(v({ "Student ID": "SIM-001", Student: "Alex Renamed" }), { id: "SIM-001", name: "Alex R." }), true, "same ID counts even if the name was corrected");
assert.equal(H.belongsToStudent(v({ "Student ID": "" }), { id: "SIM-001", name: "alex r." }), true, "row without an ID matches exact name, case-insensitive");
const lookup = H.lookupStudents([{ id: "SIM-001", name: "Alex R." }], ["alex r.", "Sam Former", " Sam Former ", "", null]);
assert.deepEqual(JSON.parse(JSON.stringify(lookup)), [{ id: "SIM-001", name: "Alex R." }, { id: "name:sam former", name: "Sam Former", nameOnly: true }], "roster + names found on visits, deduped case-insensitively");
assert.equal(H.belongsToStudent(v({ "Student ID": "" }), { id: "SIM-001", name: "Alex" }), false, "never a partial match");
assert.equal(H.belongsToStudent(v(), { id: "", name: "Alex R." }), true, "no id on lookup side falls back to name");
assert.equal(H.belongsToStudent(v(), { id: "", name: "" }), false);

const rows = [
  v({ id: "a", Date: YESTERDAY, "Time In": "13:00" }),
  v({ id: "b", Date: TODAY, "Time In": "9:05" }),
  v({ id: "c", Date: TODAY, "Time In": "10:15" }),
  v({ id: "d", Student: "Jordan M.", "Student ID": "SIM-002" }),
  v({ id: "e", Date: `${YESTERDAY}T00:00:00Z`, "Time In": "08:00", "Time Out": "" })
];
const mine = H.forStudent(rows, { id: "SIM-001", name: "Alex R." });
assert.deepEqual(mine.map(r => r.id), ["c", "b", "a", "e"], "newest first, single-digit hour sorts correctly, other students excluded, open rows kept");
const groups = H.groupByDate(mine);
assert.deepEqual(groups.map(g => [g.date, g.visits.length]), [[TODAY, 2], [YESTERDAY, 2]]);
assert.deepEqual(H.summary(mine), { visits: 4, days: 2, earliest: YESTERDAY });
assert.deepEqual(H.summary([]), { visits: 0, days: 0, earliest: "" });

/* ── detail model: full details, but never ids/emails/metadata ── */
const detail = H.detailModel(v({
  Reason: "Disruption, Elopement", "Intervention Used": "Calm space", "SCM Used": false,
  Notes: "  Line one\nLine two  ", "Behavior Specialist": "Dana Fielding, Marcus Webb",
  "Teacher Came From": "Mr. Bellamy", "PACE Room": "pace-room-2", Email: "x@y", "Entry ID": "zzz", Duration: 30
}));
assert.deepEqual(detail.reasons, ["Disruption", "Elopement"]);
assert.deepEqual(detail.supports, ["Calm space"]);
assert.equal(detail.scm, "No");
assert.equal(detail.notes, "Line one\nLine two");
assert.deepEqual(detail.specialists, ["Dana Fielding", "Marcus Webb"]);
assert.equal(detail.duration, 30);
assert.equal(detail.open, false);
const serialized = JSON.stringify(detail);
for (const leak of ["x@y", "zzz", "SIM-001"]) assert.ok(!serialized.includes(leak), `${leak} leaked into detailModel`);
assert.equal(H.detailModel(v({ "SCM Used": null })).scm, "");
assert.equal(H.detailModel(v({ "Time Out": "" })).open, true);

/* ── Graph: per-student read is server-filtered, escaped, paged ── */
///////////////////////////////////////////////////////////////////////////////////////////////
// Function Name: graphStudentRead
// Description: Checks the Graph per-student read, the per-date read, and the name-list read.
// Parameters: none
///////////////////////////////////////////////////////////////////////////////////////////////
async function graphStudentRead() {
  const ctx = vm.createContext({
    console, APP_MODE: "production",
    CONFIG: { SITE: "s", LISTS: { paceVisits: "IEP_Pace_Visits" }, ROOMS: [] },
    AUTH: { acquireGraphToken: async () => "t" },
    fetch: async () => { throw new Error("network"); }
  });
  vm.runInContext(src("graph.js") + "\nthis.G = GRAPH;", ctx);
  const G = ctx.G;
  G.getSiteId = async () => "site"; G.getListId = async () => "list";
  G.getListSchema = async () => ({ Student: "Title", Date: "VisitDate", "Time In": "TimeIn" });
  const calls = [];
  G._get = async (p, headers) => {
    calls.push({ p, headers });
    if (calls.length === 1) return { value: [{ id: "1", fields: { Title: "Ja'de O'Brien", TimeIn: "09:00" } }], "@odata.nextLink": "https://graph.microsoft.com/v1.0/next-page" };
    return { value: [{ id: "2", fields: { Title: "Ja'de O'Brien", TimeIn: "10:00" } }] };
  };
  const out = await G.getPaceVisitsForStudentByDisplayName("Ja'de O'Brien");
  assert.equal(calls.length, 2, "follows paging");
  assert.equal(calls[1].p, "next-page");
  const decoded = decodeURIComponent(calls[0].p);
  assert.ok(decoded.includes("$filter=fields/Title eq 'Ja''de O''Brien'"), decoded);
  assert.equal(calls[0].headers.Prefer, "HonorNonIndexedQueriesWarningMayFailRandomly");
  assert.deepEqual(plain(out).map(r => [r.id, r.Student, r["Time In"]]), [["1", "Ja'de O'Brien", "09:00"], ["2", "Ja'de O'Brien", "10:00"]]);
  await assert.rejects(() => G.getPaceVisitsForStudentByDisplayName("  "), /student is required/);
  // Per-date read is unchanged after the refactor.
  calls.length = 0;
  G._get = async (p, headers) => { calls.push({ p, headers }); return { value: [] }; };
  await G.getPaceVisitsForDateByDisplayName("2026-11-23");
  const d = decodeURIComponent(calls[0].p);
  assert.ok(d.includes("fields/VisitDate ge '2026-11-23T00:00:00Z' and fields/VisitDate lt '2026-11-24T00:00:00Z'"));
  await assert.rejects(() => G.getPaceVisitsForDateByDisplayName("nope"), /valid visit date/);
  // All-years name list: only the Student column, every page, distinct names.
  calls.length = 0;
  G._get = async p => {
    calls.push({ p });
    if (calls.length === 1) return { value: [{ fields: { Title: "Alex R." } }, { fields: { Title: " alex r. " } }, { fields: {} }], "@odata.nextLink": "https://graph.microsoft.com/v1.0/page-2" };
    return { value: [{ fields: { Title: "Former Student" } }] };
  };
  const names = await G.getPaceVisitStudentNames();
  assert.deepEqual(plain(names), ["Alex R.", "Former Student"]);
  assert.equal(calls.length, 2, "name list follows paging");
  assert.ok(calls[0].p.includes("$expand=fields($select=Title)"), "downloads only the Student column");
  assert.ok(!calls[0].p.includes("$filter"), "no date or other filter: every school year");
}

/* ── PACE_DATA.getStudentVisits in demo (localStorage only, zero Graph) and production ── */
///////////////////////////////////////////////////////////////////////////////////////////////
// Function Name: providerStudentRead
// Description: Checks PACE_DATA history reads in demo (zero Graph) and production.
// Parameters: none
///////////////////////////////////////////////////////////////////////////////////////////////
async function providerStudentRead() {
  const store = new Map();
  const localStorage = { getItem: k => store.get(k) ?? null, setItem: (k, x) => store.set(k, String(x)), removeItem: k => store.delete(k) };
  let graphCalls = 0;
  const base = {
    console, localStorage, crypto: { randomUUID: () => "u" },
    window: { location: { search: "?demo=1", hostname: "localhost" } },
    fetch: async () => { graphCalls++; throw new Error("no network in demo"); }
  };
  const ctx = vm.createContext({ ...base, URLSearchParams });
  for (const f of ["config.js", "demo-data.js", "graph.js", "recent-activity.js", "visit-history.js", "visit-workflow.js", "visit-corrections.js", "pace-data.js"]) {
    vm.runInContext(src(f).replace(/^const (\w+) =/gm, "var $1 ="), ctx, { filename: f });
  }
  ctx.localStorage.setItem("paceRoomTrackerDemoData", JSON.stringify({ paceVisits: [
    { id: "p1", Room: "PACE Room 1", "Student ID": "SIM-001", Student: "Alex R.", Date: YESTERDAY, "Time In": "09:00", "Time Out": "09:20", Notes: "n" },
    { id: "p2", Room: "PACE Room 2", "Student ID": "SIM-001", Student: "Alex R.", Date: TODAY, "Time In": "11:00", "Time Out": "11:10", Notes: "n" },
    { id: "p3", Room: "PACE Room 1", "Student ID": "SIM-003", Student: "Taylor S.", Date: TODAY, "Time In": "11:00", "Time Out": "11:10", Notes: "n" }
  ] }));
  const hist = await ctx.PACE_DATA.getStudentVisits({ id: "SIM-001", name: "Alex R." });
  assert.deepEqual(plain(hist).map(r => [r.id, r["PACE Room"]]), [["p2", "pace-room-2"], ["p1", "pace-room-1"]], "demo: newest first, all rooms, room normalized");
  assert.equal(graphCalls, 0, "demo never touches Graph");
  await assert.rejects(() => ctx.PACE_DATA.getStudentVisits({ id: "x", name: "" }), /student is required/);
  // Past-day Recent read (existing provider path, new date)
  const past = await ctx.PACE_DATA.getRecentVisits({ date: YESTERDAY, room: "pace-room-1", limit: 500 });
  assert.deepEqual(plain(past.visits).map(r => r.id), ["p1"]);
  assert.deepEqual(plain(await ctx.PACE_DATA.getVisitStudentNames()).sort(), ["Alex R.", "Taylor S."], "demo name list comes from local visits");
  assert.equal(graphCalls, 0, "demo never touches Graph");

  // Production branch: uses the per-student Graph read, then narrows by id.

  const prod = vm.createContext({ ...base, URLSearchParams, window: { location: { search: "?production=1", hostname: "localhost" } } });
  for (const f of ["config.js", "demo-data.js", "graph.js", "recent-activity.js", "visit-history.js", "visit-workflow.js", "visit-corrections.js", "pace-data.js"]) {
    vm.runInContext(src(f).replace(/^const (\w+) =/gm, "var $1 ="), prod, { filename: f });
  }
  let askedFor = null;
  const A_YEAR_AGO = H.shiftDate(TODAY, -365);
  prod.GRAPH.getPaceVisitsForStudentByDisplayName = async name => { askedFor = name; return [
    { id: "10", Room: "PACE Room 1", "Student ID": "77", Student: "Sam Lee", Date: YESTERDAY, "Time In": "08:00", "Time Out": "08:30" },
    { id: "11", Room: "Other", "Student ID": "88", Student: "Sam Lee", Date: TODAY, "Time In": "08:00", "Time Out": "08:30" },
    { id: "12", Room: "", "Student ID": "", Student: "Sam Lee", Date: TODAY, "Time In": "07:00", "Time Out": "07:30" },
    { id: "13", "Student ID": "5", Student: "Sam Lee", Date: A_YEAR_AGO, "Time In": "10:00", "Time Out": "10:20" },
    { id: "14", Student: "sam lee", Date: `${H.shiftDate(TODAY, -500)}T04:00:00Z`, "Time In": "10:00", "Time Out": "10:20" }
  ]; };
  const prodHist = await prod.PACE_DATA.getStudentVisits({ id: "77", name: "Sam Lee" });
  assert.equal(askedFor, "Sam Lee");
  assert.deepEqual(plain(prodHist).map(r => r.id), ["12", "10", "14"],
    "other students' IDs (88, 5) excluded; rows with no ID matched by name, any date");
  // A name found only on visit rows (no roster ID) gets every row with that name.
  assert.deepEqual(plain(await prod.PACE_DATA.getStudentVisits({ id: "", name: "Sam Lee" })).map(r => r.id), ["11", "12", "10", "13", "14"]);
  // Search list: whole roster (any state); a roster failure yields [] and never blocks.
  prod.ROSTER = { loaded: true, getAll: () => [{ id: "77", name: "Sam Lee", active: false, paceEnabled: false }] };
  assert.equal(plain(await prod.PACE_DATA.getAllStudents()).length, 1, "inactive / non-PACE students are searchable too");
  prod.ROSTER = { loaded: false, refresh: async () => { throw new Error("roster down"); }, getAll: () => [] };
  assert.deepEqual(plain(await prod.PACE_DATA.getAllStudents()), []);
  // All-years name list is cached for 10 minutes.
  let nameReads = 0;
  prod.GRAPH.getPaceVisitStudentNames = async () => { nameReads++; return ["Sam Lee", "Former Student"]; };
  await prod.PACE_DATA.getVisitStudentNames(); await prod.PACE_DATA.getVisitStudentNames();
  assert.equal(nameReads, 1, "name list cached");
}

/* ── structural: Recent stays privacy-limited and today-only for writes ── */
///////////////////////////////////////////////////////////////////////////////////////////////
// Function Name: structural
// Description: Checks Recent stays privacy-limited and the history code never writes.
// Parameters: none
///////////////////////////////////////////////////////////////////////////////////////////////
function structural() {
  const app = src("app.js");
  const recentBlock = app.slice(app.indexOf("async function loadRecentActivity"), app.indexOf("/* ── DEMO: reset control"));
  for (const bad of ["Notes", "visitDetailsHtml", "detailModel", "Student ID"]) assert.ok(!recentBlock.includes(bad), `${bad} in Recent render path`);
  assert.match(recentBlock, /\$\{isToday \? `<button class="recent-edit-btn"/, "Edit/Delete only rendered for today");
  assert.ok(!recentBlock.includes("STATE.paceVisits = result.visits"), "Recent must not clobber the room's open-visit cache");
  const html = src("index.html");
  assert.match(html, /<script src="recent-activity\.js"><\/script>\s*<script src="visit-history\.js"><\/script>/);
  assert.match(src("sw.js"), /"\.\/visit-history\.js"/);
  assert.ok(Number(src("sw.js").match(/shell-v(\d+)/)[1]) >= 22);
  const historyBlock = app.slice(app.indexOf("/* ── STUDENT HISTORY PATCH"), app.indexOf("/* ── TEMPORARY PATCH 003"));
  assert.doesNotMatch(historyBlock, /(PACE_DATA|GRAPH|DemoStorage)\.(createVisit|updateVisit|deleteVisit|completeVisit|savePaceVisit|updatePaceVisit|deletePaceVisit)\([^)]/, "history block never writes");
  assert.match(historyBlock, /panel\.innerHTML = opening \? visitDetailsHtml\(visit\) : ""/, "details built only on tap");
}

(async () => {
  await graphStudentRead();
  await providerStudentRead();
  structural();
  console.log("student-history tests: all passed");
})().catch(err => { console.error(err); process.exit(1); });
~~~~

## `~/.pace/tools/doc-audit.js`

~~~~js
///////////////////////////////////////////////////////////////////////////////////////////////
// Author: R-E Miller
// Creation Date: November 23, 2026
// Filename: doc-audit.js
// Purpose: Read-only check that every PACE script follows the house documentation style: a file
//          header (Author / Creation Date / Filename / Purpose) and a Function Name /
//          Description / Parameters block directly above every named function, with 95-slash
//          bars. Usage: node ~/.pace/tools/doc-audit.js <dir> [file ...]. Exits 1 on problems.
///////////////////////////////////////////////////////////////////////////////////////////////
const fs = require("fs");
const path = require("path");
const dir = process.argv[2];
const files = process.argv.length > 3 ? process.argv.slice(3) : fs.readdirSync(dir).filter(f => f.endsWith(".js"));
const BAR = "/".repeat(95);
const problems = [];
const MODULES = new Set(["RECENT_ACTIVITY", "PACE_VISIT_CORRECTIONS", "VISIT_HISTORY"]);
const KEYWORDS = new Set(["if", "for", "while", "switch", "catch", "function", "return", "with"]);

for (const file of files) {
  const lines = fs.readFileSync(path.join(dir, file), "utf8").split(/\r?\n/);
  // File header
  const head = lines.slice(0, 6).join("\n");
  if (lines[0] !== BAR || !/\/\/ Author: /.test(head) || !/\/\/ Creation Date: /.test(head) || !new RegExp(`// Filename: ${file.replace(".", "\\.")}`).test(head) || !/\/\/ Purpose: /.test(head)) {
    problems.push(`${file}:1 file header missing or incomplete (Author / Creation Date / Filename / Purpose)`);
  }
  // Bar lines must be exactly 95 slashes
  lines.forEach((l, i) => { if (/^\s*\/{20,}\s*$/.test(l) && l.trim() !== BAR) problems.push(`${file}:${i + 1} separator bar is ${l.trim().length} slashes, not 95`); });

  lines.forEach((line, i) => {
    const code = line.replace(/\/\/.*$/, "");
    const indent = line.match(/^\s*/)[0].length;
    let name = null;
    let m;
    if ((m = code.match(/^\s*(?:async\s+)?function\s*\*?\s*([A-Za-z_$][\w$]*)\s*\(/))) name = m[1];
    else if ((m = code.match(/^\s*(?:const|let|var)\s+([A-Za-z_$][\w$]*)\s*=\s*(?:async\s*)?(?:\([^)]*\)|[A-Za-z_$][\w$]*)\s*=>/))) name = m[1];
    else if ((m = code.match(/^\s*(?:async\s+)?([A-Za-z_$][\w$]*)\s*\([^)]*\)\s*\{\s*(?:.*\})?\s*$/)) && !KEYWORDS.has(m[1]) && indent >= 2 && indent <= 4) name = m[1];
    if (!name || MODULES.has(name)) return;
    // Look back for the header block, skipping blank and plain comment lines
    let j = i - 1;
    while (j >= 0 && lines[j].trim() === "") j--;
    if (j < 0 || lines[j].trim() !== BAR) { problems.push(`${file}:${i + 1} ${name}: no header block directly above`); return; }
    const block = [];
    let k = j - 1;
    while (k >= 0 && lines[k].trim() !== BAR) { block.unshift(lines[k].trim()); k--; }
    const text = block.join("\n");
    if (!new RegExp(`^// Function Name: ${name.replace(/\$/g, "\\$")}$`, "m").test(text)) problems.push(`${file}:${i + 1} ${name}: "Function Name:" missing or doesn't match`);
    if (!/^\/\/ Description: \S/m.test(text)) problems.push(`${file}:${i + 1} ${name}: "Description:" missing`);
    if (!/^\/\/ Parameters: \S/m.test(text)) problems.push(`${file}:${i + 1} ${name}: "Parameters:" missing`);
    else {
      const params = text.slice(text.indexOf("// Parameters:"));
      if (!/^\/\/ Parameters: none\b/m.test(params) && !/- (input|output|input\/output)\b/.test(params.replace(/\n\/\/\s*/g, " "))) {
        problems.push(`${file}:${i + 1} ${name}: Parameters don't end with "- input" / "- output"`);
      }
    }
  });
}
console.log(problems.length ? problems.join("\n") : "No problems found.");
console.log(`\n${problems.length} problem(s) in ${files.length} file(s).`);
~~~~
