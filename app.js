///////////////////////////////////////////////////////////////////////////////////////////////
// Author: R-E Miller & Greg Macer
// Creation Date: August 20, 2026
// Filename: app.js
// Purpose: Browser-side controller for the PACE Room Tracker kiosk app. Drives the horizontal-
//          panel workflow (Home → Room → Specialist → Student search → Teacher Came From →
//          Visit Info → Reason → Support → SCM → Notes → Confirm → Save, plus Room's own Exit
//          and Recent panels), owns screen navigation and the transient STATE object, and
//          translates user actions into a provider-neutral visit object. Talks to the backend
//          only through PACE_DATA so demo mode and production mode share identical UI behavior.
//          Save returns to Student Search (not Room) so the same specialist can log another
//          visit right away; see the save handler. PATCH 006 removed the old homeroom-Teacher-
//          grouped Student screen — see README.md for the phase-by-phase build notes.
///////////////////////////////////////////////////////////////////////////////////////////////

/* ── Screen navigation (deterministic slide, no history stack needed —
   every call states its own direction) ─────────────────────────────────── */

const SCREENS = {};
document.querySelectorAll(".screen").forEach(el => { SCREENS[el.dataset.screen] = el; });
let currentScreenName = "signin";

///////////////////////////////////////////////////////////////////////////////////////////////
// Function Name: nav
// Description: Activates one screen and slides the previous screen out of view.
// Parameters: string name      - data-screen value of the destination panel - input
//             string direction - "forward" or "back" animation direction (default "forward") -
//                                 input
///////////////////////////////////////////////////////////////////////////////////////////////
function nav(name, direction = "forward") {
  const next = SCREENS[name];
  const current = SCREENS[currentScreenName];
  if (!next || next === current) return;

  next.classList.add("notransition");
  next.classList.toggle("left", direction === "back");
  next.classList.remove("active");
  void next.offsetWidth; // force reflow so the jump above applies with no animation
  next.classList.remove("notransition");

  requestAnimationFrame(() => {
    next.classList.add("active");
    next.classList.remove("left");
    if (current) {
      current.classList.remove("active");
      current.classList.toggle("left", direction === "forward");
    }
  });

  currentScreenName = name;
}

// Generic back-navigation wiring for every [data-nav] element EXCEPT
// "recent" (that one triggers an async fetch and always moves forward —
// wired separately below). Several screens intentionally share the same
// data-nav target (e.g. both the Reason screen's Back button and the
// Duplicate screen's Cancel button point at "student"), so this must use
// querySelectorAll, not querySelector, or only the first one wires up.
const BACK_NAV_PRERENDER = {
  specialist: () => renderSpecialistGrid(),
  student: () => renderStudentGrid(),
  cameFrom: () => renderCameFromGrid(),
  visitinfo: () => renderVisitInfo(),
  notes: () => renderNotesScreen()
};
document.querySelectorAll("[data-nav]").forEach(btn => {
  if (btn.dataset.nav === "recent") return;
  btn.addEventListener("click", () => {
    // Read the target at click time. Same-day editing temporarily changes
    // the Specialist screen's Cancel destination from Room to Recent.
    const target = btn.dataset.nav;
    if (target === "recent" && isEditingVisit()) {
      resetTrip();
      document.querySelector('[data-screen="specialist"] .btn-back').dataset.nav = "room";
      nav("recent", "back");
      loadRecentActivity();
      return;
    }
    if (BACK_NAV_PRERENDER[target]) BACK_NAV_PRERENDER[target]();
    nav(target, "back");
  });
});

/* ── Toast ────────────────────────────────────────────────────────────── */

let toastTimer = null;
///////////////////////////////////////////////////////////////////////////////////////////////
// Function Name: showToast
// Description: Displays a short success or error message without interrupting the workflow.
// Parameters: string msg  - message text to display - input
//             string kind - "ok" or "error" styling variant (default "ok") - input
///////////////////////////////////////////////////////////////////////////////////////////////
function showToast(msg, kind = "ok") {
  const el = document.getElementById("toast");
  el.textContent = msg;
  el.classList.toggle("error", kind === "error");
  el.classList.remove("hidden");
  clearTimeout(toastTimer);
  toastTimer = setTimeout(() => el.classList.add("hidden"), 2600);
}

/* ── Time helpers ─────────────────────────────────────────────────────── */

///////////////////////////////////////////////////////////////////////////////////////////////
// Function Name: nowHHMM
// Description: Returns the current local time in the HTML time-input format.
// Parameters: none
///////////////////////////////////////////////////////////////////////////////////////////////
function nowHHMM() {
  const d = new Date();
  return `${String(d.getHours()).padStart(2, "0")}:${String(d.getMinutes()).padStart(2, "0")}`;
}
// PATCH 001: this used to be `new Date().toISOString().slice(0,10)`, which
// converts to UTC first — in any US timezone that rolls to the WRONG local
// calendar date in the evening (e.g. 8:30 PM Eastern is already past
// midnight UTC). Build the date from local getters instead so "today"
// always means the staff member's local today, never a UTC-shifted one.
///////////////////////////////////////////////////////////////////////////////////////////////
// Function Name: todayISODate
// Description: Returns today's local calendar date as YYYY-MM-DD.
// Parameters: none
///////////////////////////////////////////////////////////////////////////////////////////////
function todayISODate() {
  const d = new Date();
  return `${d.getFullYear()}-${String(d.getMonth() + 1).padStart(2, "0")}-${String(d.getDate()).padStart(2, "0")}`;
}
// "August 20, 2026" from a "YYYY-MM-DD" string, parsed as LOCAL date parts
// (never `new Date("YYYY-MM-DD")`, which parses as UTC midnight and can
// print the wrong day in negative-UTC-offset timezones — the same class of
// bug todayISODate() above was fixed for).
///////////////////////////////////////////////////////////////////////////////////////////////
// Function Name: fmtLongDate
// Description: Formats an ISO date for human-readable confirmation screens.
// Parameters: string dateStr - ISO "YYYY-MM-DD" date to format - input
///////////////////////////////////////////////////////////////////////////////////////////////
function fmtLongDate(dateStr) {
  const d = dateOnly(dateStr);
  if (!d) return "—";
  const [y, m, day] = d.split("-").map(Number);
  return new Date(y, m - 1, day).toLocaleDateString(undefined, { year: "numeric", month: "long", day: "numeric" });
}
// SharePoint may hand back "Date" as a bare "YYYY-MM-DD" (what this app
// writes) or as a full ISO datetime (if the column is a true Date/Time
// field) — normalize to just the date part before comparing/parsing it.
///////////////////////////////////////////////////////////////////////////////////////////////
// Function Name: dateOnly
// Description: Extracts the date portion from a SharePoint or local date value.
// Parameters: string value - raw date or datetime value from the provider - input
///////////////////////////////////////////////////////////////////////////////////////////////
function dateOnly(value) {
  return String(value || "").slice(0, 10);
}
///////////////////////////////////////////////////////////////////////////////////////////////
// Function Name: fmt12h
// Description: Converts a 24-hour HH:MM value to a compact local 12-hour label.
// Parameters: string hhmm - 24-hour "HH:MM" time value - input
///////////////////////////////////////////////////////////////////////////////////////////////
function fmt12h(hhmm) {
  if (!hhmm || !hhmm.includes(":")) return "—";
  let [h, m] = hhmm.split(":").map(Number);
  const ampm = h >= 12 ? "PM" : "AM";
  h = h % 12; if (h === 0) h = 12;
  return `${h}:${String(m).padStart(2, "0")} ${ampm}`;
}
// Elapsed minutes between an "HH:MM" time-in (assumed today, or the given
// date) and now. Handles the rare midnight-crossing open visit gracefully.
///////////////////////////////////////////////////////////////////////////////////////////////
// Function Name: elapsedMinutesSince
// Description: Calculates elapsed minutes from a visit start until the current local time.
// Parameters: string dateStr - visit date, or empty to assume today - input
//             string hhmm    - "HH:MM" start time - input
///////////////////////////////////////////////////////////////////////////////////////////////
function elapsedMinutesSince(dateStr, hhmm) {
  if (!hhmm) return 0;
  const [h, m] = hhmm.split(":").map(Number);
  const d = dateOnly(dateStr);
  const base = d ? new Date(d + "T00:00:00") : new Date();
  base.setHours(h, m, 0, 0);
  let diff = Math.round((Date.now() - base.getTime()) / 60000);
  if (diff < 0) diff += 24 * 60;
  return diff;
}
///////////////////////////////////////////////////////////////////////////////////////////////
// Function Name: minutesBetween
// Description: Calculates the duration between two same-day HH:MM values.
// Parameters: string hhmmIn  - "HH:MM" start time - input
//             string hhmmOut - "HH:MM" end time - input
///////////////////////////////////////////////////////////////////////////////////////////////
function minutesBetween(hhmmIn, hhmmOut) {
  const [ih, im] = hhmmIn.split(":").map(Number);
  const [oh, om] = hhmmOut.split(":").map(Number);
  let dur = (oh * 60 + om) - (ih * 60 + im);
  if (dur < 0) dur += 24 * 60;
  return dur;
}

// PATCH 001: the ONE reusable duration calculation for the new completed-
// visit workflow (Visit Info + Confirm screens both call this). Unlike
// minutesBetween() above — which wraps past midnight and stays in place
// for the legacy open→close Exit flow's elapsed-time display — this is
// deliberately same-day-only per spec: a Time Out at or before Time In is
// invalid, not a visit that "wrapped around," so this returns null rather
// than guessing.
///////////////////////////////////////////////////////////////////////////////////////////////
// Function Name: visitDurationMinutes
// Description: Returns a valid visit duration or null when the time pair is incomplete/invalid.
// Parameters: string timeIn  - "HH:MM" visit start time - input
//             string timeOut - "HH:MM" visit end time - input
///////////////////////////////////////////////////////////////////////////////////////////////
function visitDurationMinutes(timeIn, timeOut) {
  return PACE_VISIT_WORKFLOW.durationMinutes(timeIn, timeOut);
}

/* ── App state ────────────────────────────────────────────────────────── */

// Mutable state for the currently visible room and visit workflow. The object is intentionally
// centralized because the UI is implemented as multiple static panels rather than a component
// framework: a property is changed by event handlers and read by render functions; resetTrip()
// clears visit-specific fields between entries while preserving the selected room and the
// sticky specialist selection.
const STATE = {
  room: null,            // Internal room id selected on the Home screen.
  paceVisits: [],        // Cached today's IEP_Pace_Visits rows (display-name keyed).
  // PATCH 004: intentionally separate from AUTH's authenticated Microsoft
  // identity — staffMembers is the selected Behavior Specialist(s) (PATCH
  // 005: loaded dynamically from IEP_Users2, see pace-data.js's
  // getSpecialists()), never conflated with who actually signed in.
  // Sticky for the whole room session (not cleared by resetTrip()) since
  // the same specialist(s) typically log several visits in a row — see
  // enterRoom()/"Change Staff" for where it DOES reset.
  // PATCH 010: one or more — was a single string (STATE.staffMember)
  // before multi-select. See README "PATCH 010" for the full writeup.
  staffMembers: [],
  student: null,         // Normalized roster record selected for the visit.
  // PATCH 006: who the student was physically with immediately before
  // this PACE visit — NOT the roster's homeroom `student.teacher` (that's
  // preserved on the roster object but no longer drives navigation; see
  // README). Deliberately its own property, not reused from the old
  // homeroom-grouping STATE.teacher (removed this patch), so the two
  // concepts can never be confused.
  cameFromTeacher: null,
  reasons: [],            // Selected behavior/reason labels.
  supports: [],           // Selected intervention/support labels.
  scmUsed: null,          // Explicit SCM answer; null means unanswered.
  notes: "",              // Required completion note, when applicable.
  timeIn: "",             // Local HH:MM start time.
  timeOut: "",  // PATCH 001: collected on the Visit Info screen, before save
  date: "",               // Local YYYY-MM-DD visit date.
  submissionId: null,     // Stable id used to protect retries.
  workflowMode: null,     // One of PACE_VISIT_WORKFLOW.MODES.
  completionTarget: null, // Open row being completed.
  editingVisitId: null,   // SharePoint/demo id being edited.
  editingOriginalDate: "", // Original date captured at edit start.
  editingRoom: "",        // Original/internal room id for the edit.
  editOriginal: null,    // STAFF CORRECTIONS: values shown when the edit began; the baseline for "what changed"
  deleteTarget: null,    // STAFF CORRECTIONS: the Recent Activity visit awaiting delete confirmation
  deleting: false,        // Prevents duplicate delete submissions.
  saving: false,          // Prevents duplicate create/update submissions.
  exitTarget: null,       // the open-visit row being closed
  duplicateTarget: null   // the open-visit row that blocked a new entry
};

// Clears everything specific to ONE visit-in-progress. Deliberately does
// NOT touch STATE.staffMembers (sticky across a room session — see STATE
// declaration above) or STATE.room.
///////////////////////////////////////////////////////////////////////////////////////////////
// Function Name: resetTrip
// Description: Clears transient workflow state while preserving the selected room.
// Parameters: none
///////////////////////////////////////////////////////////////////////////////////////////////
function resetTrip() {
  STATE.student = null;
  STATE.cameFromTeacher = null;
  STATE.reasons = [];
  STATE.supports = [];
  STATE.scmUsed = null;
  STATE.notes = "";
  STATE.timeIn = "";
  STATE.timeOut = "";
  STATE.date = "";
  STATE.submissionId = null;
  STATE.workflowMode = null;
  STATE.completionTarget = null;
  STATE.editingVisitId = null;
  STATE.editingOriginalDate = "";
  STATE.editingRoom = "";
  STATE.editOriginal = null;
  STATE.exitTarget = null;
  STATE.duplicateTarget = null;
}

///////////////////////////////////////////////////////////////////////////////////////////////
// Function Name: isEditingVisit
// Description: Reports whether the workflow is editing an existing completed visit.
// Parameters: none
///////////////////////////////////////////////////////////////////////////////////////////////
function isEditingVisit() {
  return Boolean(STATE.editingVisitId);
}

///////////////////////////////////////////////////////////////////////////////////////////////
// Function Name: isLiveStart
// Description: Reports whether the current workflow creates an open live visit.
// Parameters: none
///////////////////////////////////////////////////////////////////////////////////////////////
function isLiveStart() {
  return STATE.workflowMode === PACE_VISIT_WORKFLOW.MODES.LIVE_START;
}

///////////////////////////////////////////////////////////////////////////////////////////////
// Function Name: isCompletingVisit
// Description: Reports whether the current workflow is completing an existing open visit.
// Parameters: none
///////////////////////////////////////////////////////////////////////////////////////////////
function isCompletingVisit() {
  return STATE.workflowMode === PACE_VISIT_WORKFLOW.MODES.COMPLETING;
}

// STAFF CORRECTIONS: the ONE authorization check for editing and deleting
// visits — the same PACE gate boot() enforces (a signed-in user with an
// active IEP_Users2 record whose IEP_App_Users PACE permission allows PACE),
// re-evaluated from the state boot() already resolved. It is deliberately
// NOT a role check: any PACE-authorized staff member may correct a visit.
// Demo mode has no accounts and touches only local fake data.
///////////////////////////////////////////////////////////////////////////////////////////////
// Function Name: isPaceAuthorized
// Description: Returns whether the authenticated staff member passed the PACE access gate.
// Parameters: none
///////////////////////////////////////////////////////////////////////////////////////////////
function isPaceAuthorized() {
  if (APP_MODE === "demo") return true;
  return AUTH.isAuthenticated && APP_USERS.decide("PACE", true).allowed;
}

///////////////////////////////////////////////////////////////////////////////////////////////
// Function Name: beginVisitWorkflow
// Description: Starts a new live-start or completed-entry workflow from the room screen.
// Parameters: string mode - workflow mode to start (see PACE_VISIT_WORKFLOW.MODES) - input
///////////////////////////////////////////////////////////////////////////////////////////////
function beginVisitWorkflow(mode) {
  resetTrip();
  STATE.workflowMode = mode;
  openSpecialistScreen({ preserveTrip: true });
}

// "Remember last room" uses a demo-specific key in demo mode (spec: keep
// simulated state fully separate from anything a production deployment
// would persist), and the real key otherwise.
///////////////////////////////////////////////////////////////////////////////////////////////
// Function Name: roomStorageKey
// Description: Returns the local-storage key used to remember this device's last room.
// Parameters: none
///////////////////////////////////////////////////////////////////////////////////////////////
function roomStorageKey() {
  return APP_MODE === "demo" ? DEMO_CONFIG.roomStorageKey : CONFIG.STORAGE_KEYS.LAST_ROOM;
}

/* ── Boot ─────────────────────────────────────────────────────────────── */

// Consistent IEP Skook access-denied template (same wording used across
// the other three apps in the suite). Used both for "no active IEP_Users2
// record" and "IEP_App_Users PACE = No / lookup failed" — both are signed
// in, both stop before any room/visit data loads.
///////////////////////////////////////////////////////////////////////////////////////////////
// Function Name: showUnauthorized
// Description: Shows the authorization failure panel and its user-facing explanation.
// Parameters: string message - user-facing explanation of the access denial - input
///////////////////////////////////////////////////////////////////////////////////////////////
function showUnauthorized(message) {
  document.getElementById("unauthorizedMsg").textContent = message;
  const signedInAs = AUTH.staffName || AUTH.displayName || "";
  const el = document.getElementById("unauthorizedSignedInAs");
  if (el) el.textContent = signedInAs ? `Signed in as: ${signedInAs}` : "";
  nav("unauthorized", "forward");
}

///////////////////////////////////////////////////////////////////////////////////////////////
// Function Name: boot
// Description: Initializes the application and chooses its initial authenticated screen. In
//              production this initializes MSAL, verifies the staff directory record, resolves
//              the optional application-permission record, and loads the roster. In demo mode
//              startDemoMode() is used instead, which avoids authentication and all Microsoft
//              network access. Errors are converted into the unauthorized/loading UI state.
// Parameters: none
///////////////////////////////////////////////////////////////////////////////////////////////
async function boot() {
  if (APP_MODE === "demo") {
    await startDemoMode();
    return;
  }

  await AUTH.init();

  if (!AUTH.account) {
    nav("signin", "forward");
    return;
  }

  nav("loading", "forward");
  document.getElementById("loadingStatus").textContent = "Signing you in…";

  if (AUTH.isUnauthorized) {
    showUnauthorized(AUTH.lookupError || "Your account is not approved for PACE Room Tracker.");
    return;
  }
  if (!AUTH.isAuthenticated) {
    // account present but staff lookup didn't resolve; loadStaffFromSharePoint already ran in AUTH.init
    showUnauthorized(AUTH.lookupError || "Unable to verify your account.");
    return;
  }

  // PATCH C: authorization now also runs through IEP_App_Users (email-
  // keyed), on top of the "active IEP_Users2 record" check above. See
  // iep-app-users.js for the reader/decide() and its MIGRATION MODE
  // contract: no matching row falls back to legacyAllowed=true (today's
  // actual PACE behavior — any active IEP_Users2 user), a lookup failure
  // always denies. This runs before any room/visit data is shown.
  await APP_USERS.resolve(AUTH.account?.username || "");
  const paceDecision = APP_USERS.decide("PACE", true);
  if (!paceDecision.allowed) {
    showUnauthorized("You are signed in, but your account does not currently have access to PACE Room Tracker.");
    return;
  }

  // Production-only authenticated shell control. Never the raw email, and
  // demoBanner/authToolbar are never touched on the demo path, so the two
  // modes' indicators can't ever both show.
  const authToolbar = document.getElementById("authToolbar");
  const indicator = document.getElementById("signedInIndicator");
  if (authToolbar && indicator && AUTH.isAuthenticated) {
    indicator.textContent = `Signed in as ${AUTH.staffName || "you"}`;
    authToolbar.classList.remove("hidden");
  }

  // TEMPORARY PATCH 003 DIAGNOSTIC — only ever unhidden here, i.e. only
  // after a real signed-in + authorized production boot. Never reached in
  // demo mode (this whole branch of boot() is production-only).
  document.getElementById("runDiagnosticBtn")?.classList.remove("hidden");

  const lastRoom = localStorage.getItem(roomStorageKey());
  if (lastRoom && CONFIG.ROOMS.some(r => r.id === lastRoom)) {
    await enterRoom(lastRoom, "forward");
  } else {
    nav("home", "forward");
  }
}

// Demo mode entry point: no MSAL, no Graph, no sign-in screen at all —
// straight to Room selection (or straight into the last-used room) with a
// fake staff identity already attached to STATE for "Submitted By".
///////////////////////////////////////////////////////////////////////////////////////////////
// Function Name: startDemoMode
// Description: Initializes the local-only demo provider and enters the home panel.
// Parameters: none
///////////////////////////////////////////////////////////////////////////////////////////////
async function startDemoMode() {
  document.body.classList.add("demo-mode");
  document.getElementById("demoBanner").classList.remove("hidden");
  document.getElementById("resetDemoBtn")?.classList.remove("hidden");
  AUTH.staffName = DEMO_USER.name;

  const savedRoom = localStorage.getItem(roomStorageKey());
  if (savedRoom && CONFIG.ROOMS.some(r => r.id === savedRoom)) {
    await enterRoom(savedRoom, "forward");
  } else {
    nav("home", "forward");
  }
}

document.getElementById("signInBtn").addEventListener("click", () => AUTH.login());
document.getElementById("signOutFromUnauthBtn").addEventListener("click", () => AUTH.logout());
document.getElementById("signOutBtn")?.addEventListener("click", () => {
  if (APP_MODE === "demo" || !AUTH.account) return;
  AUTH.logout();
});

/* ── HOME → ROOM ──────────────────────────────────────────────────────── */

document.querySelectorAll(".room-card").forEach(btn => {
  btn.addEventListener("click", () => enterRoom(btn.dataset.room, "forward"));
});

// PATCH 011: the CONFIG.ROOMS entry for whatever room is "in effect" right
// now — the room being edited if a same-day edit is in progress, otherwise
// the selected room. Centralizes the one lookup every "is this a physical
// PACE room, or the non-room 'Other' location" copy decision below reads,
// rather than repeating `CONFIG.ROOMS.find(...)` at each call site.
///////////////////////////////////////////////////////////////////////////////////////////////
// Function Name: currentRoomConfig
// Description: Returns the configured metadata for the currently selected room.
// Parameters: none
///////////////////////////////////////////////////////////////////////////////////////////////
function currentRoomConfig() {
  return CONFIG.ROOMS.find(r => r.id === (STATE.editingRoom || STATE.room));
}

///////////////////////////////////////////////////////////////////////////////////////////////
// Function Name: enterRoom
// Description: Selects a room and opens its operational workspace. Persists the selected room
//              and updates all room-specific labels before loading its open visits.
// Parameters: string roomId    - internal room id from CONFIG.ROOMS - input
//             string direction - "forward" or "back" screen transition direction - input
///////////////////////////////////////////////////////////////////////////////////////////////
async function enterRoom(roomId, direction) {
  STATE.room = roomId;
  // PATCH 004: a (re-)entered room is a fresh session — re-confirm the
  // specialist(s) rather than silently carrying one over from a previous
  // room or app launch. resetTrip() also clears any stale student/
  // came-from-teacher left over from an incomplete attempt.
  STATE.staffMembers = [];
  resetTrip();
  localStorage.setItem(roomStorageKey(), roomId);
  const room = CONFIG.ROOMS.find(r => r.id === roomId);
  document.getElementById("roomTitle").textContent = (room?.label || roomId).toUpperCase();
  // PATCH 001: single room-state/theme mechanism — everything colored by
  // room (badges, accents) reads from this one attribute via CSS custom
  // properties (see styles.css `body[data-room="..."]`), rather than any
  // screen styling itself individually.
  document.body.dataset.room = roomId;
  updateRoomBadges();
  // PATCH 011: "Other" isn't a physical PACE room, so the room screen's
  // copy that assumes one ("...entering PACE now", "Currently in PACE")
  // needs to say something else there.
  const isPhysicalRoom = room?.isPhysicalRoom !== false;
  document.getElementById("startVisitCaption").textContent =
    isPhysicalRoom ? "Student is entering PACE now" : "Support is starting now";
  document.getElementById("currentlyInPaceHeading").textContent =
    isPhysicalRoom ? "Currently in PACE" : "Currently receiving support";
  nav("room", direction);
  await refreshRoomVisits();
}

// Keeps every `.room-badge` element (one per workflow screen) in sync with
// the currently selected room — "PACE ROOM 1 · YELLOW HALL" — a single
// function driving every instance rather than per-screen logic.
///////////////////////////////////////////////////////////////////////////////////////////////
// Function Name: updateRoomBadges
// Description: Updates room labels and contextual wording across all workflow panels.
// Parameters: none
///////////////////////////////////////////////////////////////////////////////////////////////
function updateRoomBadges() {
  const room = CONFIG.ROOMS.find(r => r.id === STATE.room);
  const text = room ? `${room.label.toUpperCase()} · ${room.hallway.toUpperCase()}` : "";
  document.querySelectorAll(".room-badge").forEach(el => { el.textContent = text; });
}

///////////////////////////////////////////////////////////////////////////////////////////////
// Function Name: refreshRoomVisits
// Description: Loads today's open visits for the selected room from the active provider.
// Parameters: none
///////////////////////////////////////////////////////////////////////////////////////////////
async function refreshRoomVisits() {
  const listEl = document.getElementById("currentlyInPace");
  listEl.innerHTML = `<p class="empty-hint">Loading…</p>`;
  try {
    STATE.paceVisits = await PACE_DATA.getVisits({ date: todayISODate() });
  } catch (err) {
    console.error("Failed to load PACE visits:", err);
    listEl.innerHTML = `<p class="empty-hint">Unable to load current activity.</p>`;
    return;
  }
  renderCurrentlyInPace();
}

///////////////////////////////////////////////////////////////////////////////////////////////
// Function Name: openVisitsForRoom
// Description: Filters the current visit cache to open visits belonging to one room.
// Parameters: string roomId - internal room id to filter open visits for - input
///////////////////////////////////////////////////////////////////////////////////////////////
function openVisitsForRoom(roomId) {
  return PACE_VISIT_WORKFLOW.openForRoom(STATE.paceVisits, roomId);
}

///////////////////////////////////////////////////////////////////////////////////////////////
// Function Name: renderCurrentlyInPace
// Description: Renders open visits and their same-row completion controls.
// Parameters: none
///////////////////////////////////////////////////////////////////////////////////////////////
function renderCurrentlyInPace() {
  const listEl = document.getElementById("currentlyInPace");
  const open = openVisitsForRoom(STATE.room);
  if (open.length === 0) {
    const isPhysicalRoom = currentRoomConfig()?.isPhysicalRoom !== false;
    listEl.innerHTML = `<p class="empty-hint">${isPhysicalRoom ? "No students currently in PACE." : "No one currently logged as receiving support."}</p>`;
    return;
  }
  listEl.innerHTML = open.map(v => `
    <div class="current-card">
      <div class="current-card-info">
        <span class="current-card-name">${escHtml(v.Student || "")}${v.demo ? ' <span class="badge-simulated">SIMULATED</span>' : ""}</span>
        <span class="current-card-meta">In: ${escHtml(fmt12h(v["Time In"]))} · ${elapsedMinutesSince(dateOnly(v.Date), v["Time In"])} min</span>
        ${v.Reason ? `<span class="current-card-meta">${escHtml(v.Reason)}</span>` : ""}
        ${RECENT_ACTIVITY.formatSpecialistsCompact(v["Behavior Specialist"] || v["Staff Member"]) ? `<span class="current-card-meta">${escHtml(RECENT_ACTIVITY.formatSpecialistsCompact(v["Behavior Specialist"] || v["Staff Member"]))}</span>` : ""}
      </div>
      <button class="current-card-exit" data-item-id="${escHtml(v.id)}">MARK COMPLETE</button>
    </div>`).join("");

  listEl.querySelectorAll(".current-card-exit").forEach(btn => {
    btn.addEventListener("click", () => {
      const item = open.find(v => String(v.id) === btn.dataset.itemId);
      if (item) openExitScreen(item);
    });
  });
}

// Keep elapsed-minute labels fresh without a full reload.
setInterval(() => { if (currentScreenName === "room") renderCurrentlyInPace(); }, 30000);

document.getElementById("startLiveVisitBtn").addEventListener("click", async () => {
  // Refresh first so duplicate-open protection includes visits started from
  // another session since this room screen was opened.
  await refreshRoomVisits();
  beginVisitWorkflow(PACE_VISIT_WORKFLOW.MODES.LIVE_START);
});
document.getElementById("logCompletedVisitBtn").addEventListener("click", () => {
  beginVisitWorkflow(PACE_VISIT_WORKFLOW.MODES.COMPLETED_ENTRY);
});

/* ── BEHAVIOR SPECIALIST (PATCH 004; PATCH 005: dynamic from IEP_Users2) ── */

// Loaded once per Specialist screen visit (via PACE_DATA.getSpecialists(),
// which branches on APP_MODE) and rendered synchronously from here on —
// same caching pattern as cachedStudents below.
let cachedSpecialists = [];

///////////////////////////////////////////////////////////////////////////////////////////////
// Function Name: openSpecialistScreen
// Description: Loads available specialists and opens the multi-select specialist panel.
// Parameters: boolean preserveTrip - keep in-progress visit state instead of resetting it
//                                    (default false) - input
///////////////////////////////////////////////////////////////////////////////////////////////
async function openSpecialistScreen({ preserveTrip = false } = {}) {
  if (!preserveTrip) resetTrip(); // new visit; edit mode preloads and preserves the existing row
  document.querySelector('[data-screen="specialist"] .btn-back').dataset.nav = isEditingVisit() ? "recent" : "room";
  document.getElementById("specialistGrid").innerHTML = `<p class="empty-hint">Loading…</p>`;
  nav("specialist", "forward");
  try {
    cachedSpecialists = await PACE_DATA.getSpecialists();
  } catch (err) {
    console.error("Failed to load Behavior Specialists:", err);
    cachedSpecialists = [];
  }
  renderSpecialistGrid();
}

// PATCH 010: multi-select. Tapping a card toggles it in/out of
// STATE.staffMembers; nothing navigates forward until CONTINUE is tapped
// (previously, tapping the one allowed specialist navigated immediately —
// see git history for the prior single-select renderSpecialistGrid()/
// selectSpecialist()). Any previously-selected name not present in the
// freshly loaded live list (e.g. picked, then that person went inactive)
// is still shown so a return visit to this screen doesn't silently drop it.
///////////////////////////////////////////////////////////////////////////////////////////////
// Function Name: renderSpecialistGrid
// Description: Renders selectable specialists and preserves the current selections.
// Parameters: none
///////////////////////////////////////////////////////////////////////////////////////////////
function renderSpecialistGrid() {
  const grid = document.getElementById("specialistGrid");
  const missing = STATE.staffMembers.filter(name => !cachedSpecialists.includes(name));
  const specialists = [...missing, ...cachedSpecialists];
  if (specialists.length === 0) {
    grid.innerHTML = `<p class="empty-hint">No Behavior Specialists found.</p>`;
  } else {
    grid.innerHTML = specialists.map(name => `
      <button class="student-card${STATE.staffMembers.includes(name) ? " selected" : ""}" data-specialist="${escHtml(name)}">${escHtml(name)}</button>
    `).join("");
    grid.querySelectorAll("[data-specialist]").forEach(btn => {
      btn.addEventListener("click", () => toggleSpecialist(btn.dataset.specialist));
    });
  }
  updateSpecialistContinueState();
}

///////////////////////////////////////////////////////////////////////////////////////////////
// Function Name: updateSpecialistContinueState
// Description: Enables specialist continuation only when at least one specialist is selected.
// Parameters: none
///////////////////////////////////////////////////////////////////////////////////////////////
function updateSpecialistContinueState() {
  const count = STATE.staffMembers.length;
  document.getElementById("specialistSelectedCount").textContent =
    count === 0 ? "" : `${count} selected`;
  document.getElementById("specialistContinueBtn").disabled = count === 0;
}

///////////////////////////////////////////////////////////////////////////////////////////////
// Function Name: toggleSpecialist
// Description: Adds or removes one specialist from the current visit.
// Parameters: string name - specialist name to toggle in STATE.staffMembers - input
///////////////////////////////////////////////////////////////////////////////////////////////
function toggleSpecialist(name) {
  const idx = STATE.staffMembers.indexOf(name);
  if (idx === -1) STATE.staffMembers.push(name); else STATE.staffMembers.splice(idx, 1);
  renderSpecialistGrid();
}

document.getElementById("specialistContinueBtn").addEventListener("click", () => {
  if (STATE.staffMembers.length === 0) return; // defensive: button is disabled at zero already
  openStudentScreen();
});

document.getElementById("changeStaffBtn").addEventListener("click", () => {
  STATE.staffMembers = [];
  renderSpecialistGrid();
  nav("specialist", "back");
});

/* ── STUDENT SELECT (PATCH 006: search-first — homeroom-Teacher-grouped
   selection removed; Teacher is not reliable for "which classroom was
   the student physically coming from") ──────────────────────────────── */

// Loaded once per Specialist→Student entry (via PACE_DATA.getStudents(),
// which branches on APP_MODE) and searched synchronously from here on —
// no re-fetch per keystroke.
let cachedStudents = [];

// Lowercase, strip punctuation ("Ja'de" still matches "jade"), collapse to
// a plain space-joined string — used by Teacher Came From search. Student
// search uses the identical rule via PACE_STUDENT_SELECT.normalize().
///////////////////////////////////////////////////////////////////////////////////////////////
// Function Name: normalizeSearchText
// Description: Normalizes user search text for case-insensitive matching.
// Parameters: string str - raw search text to normalize - input
///////////////////////////////////////////////////////////////////////////////////////////////
function normalizeSearchText(str) {
  return String(str || "").toLowerCase().replace(/[^a-z0-9\s]/g, "").trim();
}

///////////////////////////////////////////////////////////////////////////////////////////////
// Function Name: openStudentScreen
// Description: Loads the eligible roster and opens the search-first student panel.
// Parameters: none
///////////////////////////////////////////////////////////////////////////////////////////////
async function openStudentScreen() {
  document.getElementById("specialistContextLine").textContent = STATE.staffMembers.join(", ");
  document.getElementById("studentSearch").value = "";
  document.getElementById("studentRosterCount").textContent = "";
  document.getElementById("studentGrid").innerHTML = `<p class="empty-hint">Loading students…</p>`;
  nav("student", "forward");
  try {
    cachedStudents = await PACE_DATA.getStudents();
  } catch (err) {
    console.error("Failed to load students:", err);
    cachedStudents = [];
  }
  renderStudentGrid();
}

// The card label is presentation only ("Last, First" when the roster has
// separate name fields); the student's identity/value remains s.name,
// resolved by id in selectStudent() exactly as before.
///////////////////////////////////////////////////////////////////////////////////////////////
// Function Name: studentCardHtml
// Description: Builds the safe button markup for one eligible student.
// Parameters: Object s - roster student record to render - input
///////////////////////////////////////////////////////////////////////////////////////////////
function studentCardHtml(s) {
  return `<button class="student-card${STATE.student?.id === s.id ? " selected" : ""}" data-student-id="${escHtml(s.id)}">${escHtml(PACE_STUDENT_SELECT.displayLabel(s))}</button>`;
}
///////////////////////////////////////////////////////////////////////////////////////////////
// Function Name: wireStudentCards
// Description: Connects rendered student buttons to student selection handling.
// Parameters: Element grid - container holding rendered student-card buttons - input
///////////////////////////////////////////////////////////////////////////////////////////////
function wireStudentCards(grid) {
  grid.querySelectorAll(".student-card").forEach(btn => {
    btn.addEventListener("click", () => selectStudent(btn.dataset.studentId));
  });
}

///////////////////////////////////////////////////////////////////////////////////////////////
// Function Name: renderStudentGrid
// Description: Renders the complete eligible roster, optionally filtered by search text.
// Parameters: string filter - search text to filter the roster by (default "") - input
///////////////////////////////////////////////////////////////////////////////////////////////
function renderStudentGrid(filter = "") {
  const grid = document.getElementById("studentGrid");
  document.getElementById("specialistContextLine").textContent = STATE.staffMembers.join(", ");

  // Every eligible (Active + PACE Enabled) student is shown alphabetically
  // on open; the search box only FILTERS that complete roster. Nothing is
  // capped or hidden until searched — see student-select.js.
  const model = PACE_STUDENT_SELECT.renderModel(cachedStudents, filter);
  document.getElementById("studentRosterCount").textContent = model.countText;

  if (model.emptyMessage) {
    grid.innerHTML = `<p class="empty-hint">${escHtml(model.emptyMessage)}</p>` +
      `<p class="empty-hint">${escHtml(model.emptyHint)}</p>`;
    return;
  }
  grid.innerHTML = model.students.map(studentCardHtml).join("");
  wireStudentCards(grid);
}

document.getElementById("studentSearch").addEventListener("input", e => renderStudentGrid(e.target.value));

///////////////////////////////////////////////////////////////////////////////////////////////
// Function Name: selectStudent
// Description: Stores the selected student and advances or blocks on duplicate-open checks.
// Parameters: string studentId - id of the selected roster student - input
///////////////////////////////////////////////////////////////////////////////////////////////
function selectStudent(studentId) {
  const student = cachedStudents.find(s => s.id === studentId);
  if (!student) return;

  if (isLiveStart()) {
    // Duplicate protection applies to a new visit, never to the completed
    // row currently being edited.
    const openForRoom = openVisitsForRoom(STATE.room);
    const existing = openForRoom.find(v => String(v["Student ID"] || "").trim() === String(student.id).trim());
    if (existing) {
      STATE.duplicateTarget = existing;
      document.getElementById("duplicateMsg").textContent =
        `${student.name} entered PACE at ${fmt12h(existing["Time In"])}.`;
      nav("duplicate", "forward");
      return;
    }
  }

  // PATCH 001: completed-visit model — selecting a student no longer
  // creates or auto-times anything. Date defaults to today (only if this
  // is a genuinely fresh trip — resetTrip() already cleared it), but Time
  // In/Time Out are left blank for staff to enter after the fact on the
  // Visit Info screen. STATE.date/timeIn/timeOut persist across Back/Next
  // navigation from here on since nothing else resets them.
  STATE.student = student;
  // PATCH 006: a (re-)selected student always starts a fresh "came from"
  // answer — that value belongs to this specific visit/student pairing,
  // not something that should ever silently carry over from whoever was
  // selected before.
  if (!isEditingVisit()) STATE.cameFromTeacher = null;
  if (!STATE.date) STATE.date = todayISODate();
  if (!isEditingVisit()) {
    STATE.submissionId = (typeof crypto !== "undefined" && crypto.randomUUID) ? crypto.randomUUID() : `pace-${Date.now()}`;
  }

  openCameFromScreen();
}

/* ── TEACHER CAME FROM (PATCH 006) — who the student was physically with
   immediately before this PACE visit. NOT the roster's homeroom Teacher,
   and NOT the removed homeroom-grouped Teacher screen. Combines two
   SharePoint sources via PACE_DATA.getTeachers() — see that function for
   exactly which ones. ──────────────────────────────────────────────────── */

let cachedCameFromTeachers = [];

///////////////////////////////////////////////////////////////////////////////////////////////
// Function Name: openCameFromScreen
// Description: Loads teacher-origin options and opens the Teacher Came From panel.
// Parameters: none
///////////////////////////////////////////////////////////////////////////////////////////////
async function openCameFromScreen() {
  document.getElementById("cameFromGrid").innerHTML = `<p class="empty-hint">Loading…</p>`;
  document.getElementById("cameFromSearch").value = "";
  document.getElementById("cameFromOtherWrap").classList.add("hidden");
  document.getElementById("cameFromOtherInput").value = "";
  nav("cameFrom", "forward");
  try {
    cachedCameFromTeachers = await PACE_DATA.getTeachers();
  } catch (err) {
    console.error("Failed to load teachers:", err);
    cachedCameFromTeachers = [];
  }
  renderCameFromGrid();
}

///////////////////////////////////////////////////////////////////////////////////////////////
// Function Name: renderCameFromGrid
// Description: Renders searchable teacher-origin choices plus the manual-entry option.
// Parameters: string filter - search text to filter teacher options by (default "") - input
///////////////////////////////////////////////////////////////////////////////////////////////
function renderCameFromGrid(filter = "") {
  const grid = document.getElementById("cameFromGrid");
  const q = normalizeSearchText(filter);
  let teachers = q ? cachedCameFromTeachers.filter(t => normalizeSearchText(t).includes(q)) : cachedCameFromTeachers.slice();
  if (STATE.cameFromTeacher && !cachedCameFromTeachers.includes(STATE.cameFromTeacher)) {
    const currentMatches = !q || normalizeSearchText(STATE.cameFromTeacher).includes(q);
    if (currentMatches) teachers.unshift(STATE.cameFromTeacher);
  }

  const namedCards = teachers.length > 0
    ? teachers.map(t => `<button class="student-card${STATE.cameFromTeacher === t ? " selected" : ""}" data-came-from="${escHtml(t)}">${escHtml(t)}</button>`).join("")
    : `<p class="empty-hint">No matching teachers.</p>`;
  // Spec §10: always reachable, regardless of search text, so the app is
  // never blocked while the teacher roster is still being completed.
  const otherCard = `<button class="student-card other-option" id="cameFromOtherBtn">Other / Not Listed</button>`;
  grid.innerHTML = namedCards + otherCard;

  grid.querySelectorAll("[data-came-from]").forEach(btn => {
    btn.addEventListener("click", () => selectCameFromTeacher(btn.dataset.cameFrom));
  });
  document.getElementById("cameFromOtherBtn").addEventListener("click", () => {
    document.getElementById("cameFromOtherWrap").classList.remove("hidden");
    document.getElementById("cameFromOtherInput").focus();
  });
}

document.getElementById("cameFromSearch").addEventListener("input", e => renderCameFromGrid(e.target.value));

///////////////////////////////////////////////////////////////////////////////////////////////
// Function Name: selectCameFromTeacher
// Description: Stores the selected teacher origin and advances to visit details.
// Parameters: string name - selected or manually entered teacher name - input
///////////////////////////////////////////////////////////////////////////////////////////////
function selectCameFromTeacher(name) {
  STATE.cameFromTeacher = name;
  if (isLiveStart()) {
    openReasonScreen("forward");
  } else {
    renderVisitInfo();
    nav("visitinfo", "forward");
  }
}

document.getElementById("cameFromOtherContinueBtn").addEventListener("click", () => {
  const val = document.getElementById("cameFromOtherInput").value.trim();
  if (!val) {
    showToast("Please enter a teacher name.", "error");
    return;
  }
  // Spec §10: visit data only — never written back into IEP_Users2.
  selectCameFromTeacher(val);
});

document.getElementById("dupViewBtn").addEventListener("click", () => {
  if (STATE.duplicateTarget) openExitScreen(STATE.duplicateTarget);
});

/* ── VISIT INFO: date / time in / time out (PATCH 001) ────────────────── */

///////////////////////////////////////////////////////////////////////////////////////////////
// Function Name: openReasonScreen
// Description: Opens the behavior/reason multi-select panel for the current visit.
// Parameters: string direction - "forward" or "back" screen transition direction (default
//                                 "forward") - input
///////////////////////////////////////////////////////////////////////////////////////////////
function openReasonScreen(direction = "forward") {
  document.querySelector('[data-screen="reason"] .btn-back').dataset.nav = isLiveStart() ? "cameFrom" : "visitinfo";
  document.getElementById("reasonNextBtn").textContent = isLiveStart() ? "Next" : "Next";
  renderChipGrid("reasonChips", CONFIG.REASON_OPTIONS, STATE.reasons, () => updateNextEnabled("reason"));
  updateNextEnabled("reason");
  nav("reason", direction);
}

// Populates the screen from STATE every time it's shown — forward from
// Student or backward from Reason — so nothing is ever silently reset by
// navigating. Re-validates on every input change so the error clears the
// moment the values become valid again, without waiting for a Next tap.
///////////////////////////////////////////////////////////////////////////////////////////////
// Function Name: renderVisitInfo
// Description: Hydrates and renders date/time fields for a new or edited visit.
// Parameters: none
///////////////////////////////////////////////////////////////////////////////////////////////
function renderVisitInfo() {
  document.getElementById("visitStudentName").textContent = STATE.student?.name || "";
  const dateInput = document.getElementById("visitDateInput");
  dateInput.value = STATE.date || todayISODate();
  dateInput.disabled = isEditingVisit();
  dateInput.min = isEditingVisit() ? todayISODate() : "";
  dateInput.max = isEditingVisit() ? todayISODate() : "";
  document.getElementById("visitDateEditHint").classList.toggle("hidden", !isEditingVisit());
  document.getElementById("visitEditCancelBtn").classList.toggle("hidden", !isEditingVisit());
  document.getElementById("editIdentityHint").classList.toggle("hidden", !isEditingVisit());

  const roomCard = document.getElementById("visitRoomCard");
  const roomInput = document.getElementById("visitRoomInput");
  roomCard.classList.toggle("hidden", !isEditingVisit());
  roomInput.innerHTML = CONFIG.ROOMS.map(room =>
    `<option value="${escHtml(room.id)}"${room.id === (STATE.editingRoom || STATE.room) ? " selected" : ""}>${escHtml(room.label)}</option>`
  ).join("");
  document.getElementById("visitTimeInInput").value = STATE.timeIn || "";
  document.getElementById("visitTimeOutInput").value = STATE.timeOut || "";
  document.getElementById("visitTimeOutCard").classList.toggle("hidden", isLiveStart());
  document.getElementById("visitDurationHint").classList.toggle("hidden", isLiveStart());
  document.querySelector('[data-screen="visitinfo"] .btn-back').dataset.nav = isLiveStart() ? "reason" : "cameFrom";
  document.getElementById("visitInfoNextBtn").textContent = isLiveStart() ? "START PACE VISIT" : "Next";
  document.getElementById("visitTimeError").classList.add("hidden");
  updateVisitDurationHint();
}

///////////////////////////////////////////////////////////////////////////////////////////////
// Function Name: updateVisitDurationHint
// Description: Updates the visible duration hint from the current time inputs.
// Parameters: none
///////////////////////////////////////////////////////////////////////////////////////////////
function updateVisitDurationHint() {
  const hintEl = document.getElementById("visitDurationHint");
  const dur = visitDurationMinutes(STATE.timeIn, STATE.timeOut);
  hintEl.textContent = dur !== null ? `Duration: ${dur} minute${dur !== 1 ? "s" : ""}` : "";
}

// Shared by the Next button AND the defensive re-check right before final
// save (spec: validate "before advancing... OR before final submission").
// Returns a human-readable message, or null if everything's valid.
///////////////////////////////////////////////////////////////////////////////////////////////
// Function Name: validateVisitInfo
// Description: Validates the Visit Details panel before navigation or final submission. Live
//              starts may omit Time Out because the visit is intentionally open, but all
//              completed-entry and edit workflows require a positive same-day duration.
//              Same-day editing additionally requires the date to equal the device's local
//              date, which prevents the Recent Activity editor from changing history.
// Parameters: none
///////////////////////////////////////////////////////////////////////////////////////////////
function validateVisitInfo() {
  const date = document.getElementById("visitDateInput").value;
  const timeIn = document.getElementById("visitTimeInInput").value;
  const timeOut = document.getElementById("visitTimeOutInput").value;
  if (!date) return "Please enter the visit date.";
  if (isEditingVisit() && date !== todayISODate()) return "Only today's visits can be edited.";
  if (!timeIn) return "Please enter Time In.";
  if (isLiveStart()) return null;
  if (!timeOut) return "Please enter Time Out.";
  if (visitDurationMinutes(timeIn, timeOut) === null) return "Time Out must be later than Time In.";
  return null;
}

["visitDateInput", "visitTimeInInput", "visitTimeOutInput"].forEach(id => {
  document.getElementById(id).addEventListener("input", () => {
    STATE.date = document.getElementById("visitDateInput").value;
    STATE.timeIn = document.getElementById("visitTimeInInput").value;
    STATE.timeOut = document.getElementById("visitTimeOutInput").value;
    document.getElementById("visitTimeError").classList.add("hidden");
    updateVisitDurationHint();
  });
});

document.getElementById("visitRoomInput").addEventListener("change", event => {
  if (!isEditingVisit()) return;
  const room = CONFIG.ROOMS.find(item => item.id === event.target.value);
  if (!room) return;
  STATE.editingRoom = room.id;
});

document.getElementById("visitEditCancelBtn").addEventListener("click", () => {
  if (!isEditingVisit()) return;
  resetTrip();
  document.querySelector('[data-screen="specialist"] .btn-back').dataset.nav = "room";
  nav("recent", "back");
  loadRecentActivity();
});

document.getElementById("visitInfoNextBtn").addEventListener("click", () => {
  const error = validateVisitInfo();
  const errorEl = document.getElementById("visitTimeError");
  if (error) {
    // Block advancement, preserve everything already entered, and never
    // silently alter the staff member's times — just say what's wrong.
    errorEl.textContent = error;
    errorEl.classList.remove("hidden");
    return;
  }
  errorEl.classList.add("hidden");
  if (isLiveStart()) {
    saveLiveVisit();
    return;
  }
  openReasonScreen("forward");
});

/* ── REASON / SUPPORT chips (shared renderer) ────────────────────────── */

///////////////////////////////////////////////////////////////////////////////////////////////
// Function Name: renderChipGrid
// Description: Renders reusable toggle chips for reasons or support options.
// Parameters: string containerId - id of the element to render chips into - input
//             Array options      - available chip labels - input
//             Array selectedArr  - currently selected labels, toggled in place - input/output
//             Function onChange  - callback invoked after a chip selection changes - input
///////////////////////////////////////////////////////////////////////////////////////////////
function renderChipGrid(containerId, options, selectedArr, onChange) {
  const el = document.getElementById(containerId);
  el.innerHTML = options.map(opt => `<button type="button" class="chip${selectedArr.includes(opt) ? " selected" : ""}" data-value="${escHtml(opt)}">${escHtml(opt)}</button>`).join("");
  el.querySelectorAll(".chip").forEach(chip => {
    chip.addEventListener("click", () => {
      const val = chip.dataset.value;
      const idx = selectedArr.indexOf(val);
      if (idx === -1) selectedArr.push(val); else selectedArr.splice(idx, 1);
      chip.classList.toggle("selected");
      onChange();
    });
  });
}

///////////////////////////////////////////////////////////////////////////////////////////////
// Function Name: updateNextEnabled
// Description: Enables the next button when the named workflow step is complete.
// Parameters: string which - "reason" or "support" workflow step to check - input
///////////////////////////////////////////////////////////////////////////////////////////////
function updateNextEnabled(which) {
  if (which === "reason") document.getElementById("reasonNextBtn").disabled = STATE.reasons.length === 0;
  if (which === "support") document.getElementById("supportNextBtn").disabled = STATE.supports.length === 0;
}

document.getElementById("reasonNextBtn").addEventListener("click", () => {
  if (isLiveStart()) {
    if (!STATE.date) STATE.date = todayISODate();
    if (!STATE.timeIn) STATE.timeIn = nowHHMM();
    renderVisitInfo();
    nav("visitinfo", "forward");
    return;
  }
  renderChipGrid("supportChips", CONFIG.SUPPORT_OPTIONS, STATE.supports, () => updateNextEnabled("support"));
  document.querySelector('[data-screen="support"] .btn-back').dataset.nav = "reason";
  updateNextEnabled("support");
  nav("support", "forward");
});

document.getElementById("supportNextBtn").addEventListener("click", () => {
  document.getElementById("scmYesBtn").classList.toggle("selected", STATE.scmUsed === true);
  document.getElementById("scmNoBtn").classList.toggle("selected", STATE.scmUsed === false);
  nav("scm", "forward");
});

/* ── SCM ──────────────────────────────────────────────────────────────── */

///////////////////////////////////////////////////////////////////////////////////////////////
// Function Name: setScm
// Description: Stores the explicit SCM answer and advances to required notes.
// Parameters: boolean value - true for SCM used, false for not used - input
///////////////////////////////////////////////////////////////////////////////////////////////
function setScm(value) {
  STATE.scmUsed = value;
  document.getElementById("scmYesBtn").classList.toggle("selected", value === true);
  document.getElementById("scmNoBtn").classList.toggle("selected", value === false);
  setTimeout(() => {
    renderNotesScreen();
    nav("notes", "forward");
  }, 140); // brief pause so the selection state is visible
}
document.getElementById("scmYesBtn").addEventListener("click", () => setScm(true));
document.getElementById("scmNoBtn").addEventListener("click", () => setScm(false));

/* ── NOTES ────────────────────────────────────────────────────────────── */

// CURRENT-PATCH: this screen is never reached by a live-start visit (see
// isLiveStart()'s branch in reasonNextBtn's handler — a live start jumps
// straight from Reason to Visit Info and saves via saveLiveVisit(), which
// never sends Notes at all). Every path that DOES reach this screen —
// Mark Complete, an after-the-fact completed entry, or editing an
// existing completed visit (Recent Activity only ever surfaces already-
// completed visits for edit — see RECENT_ACTIVITY.select()) — represents
// a completed record, so Notes is unconditionally required here; no
// optional/hide-behind-a-button state is needed any more.
///////////////////////////////////////////////////////////////////////////////////////////////
// Function Name: renderNotesScreen
// Description: Hydrates the required notes field for a new visit or same-day edit.
// Parameters: none
///////////////////////////////////////////////////////////////////////////////////////////////
function renderNotesScreen() {
  const noteText = document.getElementById("noteText");
  noteText.value = STATE.notes || "";
  // PATCH 011: "Other" visits are exactly the ones staff asked to be able
  // to explain in more detail (Wiggle Room, a classroom check-in, a walk,
  // etc.), since there's no physical room to imply it — nudge for that
  // here rather than adding a second required field just for this case.
  const isPhysicalRoom = currentRoomConfig()?.isPhysicalRoom !== false;
  noteText.placeholder = isPhysicalRoom
    ? "Add a brief note about this visit…"
    : "Describe what happened and where (e.g., Wiggle Room, classroom check-in, a walk)…";
  const errorEl = document.getElementById("notesError");
  if (errorEl) errorEl.classList.add("hidden");
}

document.getElementById("noteText").addEventListener("input", event => {
  STATE.notes = event.target.value;
  document.getElementById("notesError")?.classList.add("hidden");
});
document.getElementById("notesNextBtn").addEventListener("click", () => {
  const trimmed = document.getElementById("noteText").value.trim();
  if (!trimmed) {
    const el = document.getElementById("notesError");
    if (el) { el.textContent = "Add a brief note before submitting this PACE visit."; el.classList.remove("hidden"); }
    document.getElementById("noteText").focus();
    return;
  }
  document.getElementById("notesError")?.classList.add("hidden");
  STATE.notes = trimmed;
  renderConfirmCard();
  nav("confirm", "forward");
});

/* ── CONFIRM ──────────────────────────────────────────────────────────── */

///////////////////////////////////////////////////////////////////////////////////////////////
// Function Name: buildVisitEntry
// Description: Converts STATE into the provider-neutral visit payload consumed by PACE_DATA
//              (room, identity, timing, taxonomy, and completion fields). Only reads STATE;
//              does not mutate it.
// Parameters: boolean open - true to build an open live-start row, omitting completion-only
//                             fields (default false) - input
///////////////////////////////////////////////////////////////////////////////////////////////
function buildVisitEntry({ open = false } = {}) {
  return {
    id: STATE.submissionId,
    paceRoom: isEditingVisit() ? (STATE.editingRoom || STATE.room) : STATE.room,
    staffMembers: [...STATE.staffMembers],
    cameFromTeacher: STATE.cameFromTeacher,
    studentId: STATE.student?.id || "",
    studentName: STATE.student?.name || "",
    date: STATE.date,
    timeIn: STATE.timeIn,
    timeOut: open ? "" : STATE.timeOut,
    durationMinutes: open ? null : visitDurationMinutes(STATE.timeIn, STATE.timeOut),
    behaviors: STATE.reasons,
    interventions: open ? [] : STATE.supports,
    scmUsed: open ? null : STATE.scmUsed,
    notes: open ? "" : STATE.notes,
    submittedByName: AUTH.staffName || AUTH.displayName,
    timestamp: new Date().toISOString(),
    // STAFF CORRECTIONS: present only while editing an existing visit;
    // PACE_DATA.updateVisit() uses it to send just the changed fields.
    original: isEditingVisit() ? STATE.editOriginal : undefined
  };
}

///////////////////////////////////////////////////////////////////////////////////////////////
// Function Name: saveLiveVisit
// Description: Persists a live-start visit as an open row. This path intentionally does not
//              collect Notes, SCM, or interventions, since staff may not know those completion
//              details when a student first enters. The resulting provider id is retained by
//              the visit row so Mark Complete can patch the same record later. Provider
//              failures are surfaced through the save UI.
// Parameters: none
///////////////////////////////////////////////////////////////////////////////////////////////
async function saveLiveVisit() {
  if (STATE.saving || !isLiveStart()) return;
  if (!STATE.student || STATE.staffMembers.length === 0 || !STATE.cameFromTeacher || STATE.reasons.length === 0 || !STATE.date || !STATE.timeIn) {
    showToast("Please complete the live visit details.", "error");
    return;
  }
  STATE.saving = true;
  const btn = document.getElementById("visitInfoNextBtn");
  btn.disabled = true;
  btn.textContent = "Starting…";
  try {
    const isPhysicalRoom = currentRoomConfig()?.isPhysicalRoom !== false;
    await PACE_DATA.createVisit(buildVisitEntry({ open: true }));
    await refreshRoomVisits();
    const studentName = STATE.student.name;
    resetTrip();
    btn.disabled = false;
    btn.textContent = "START PACE VISIT";
    STATE.saving = false;
    nav("room", "back");
    showToast(isPhysicalRoom ? `${studentName} is now in PACE.` : `${studentName}'s support has started.`);
  } catch (err) {
    console.error("PACE live visit start failed:", err);
    STATE.saving = false;
    btn.disabled = false;
    btn.textContent = "START PACE VISIT";
    showToast("Visit could not be started. Try again.", "error");
  }
}

// PATCH 001: read-only summary of the COMPLETE visit — editing now happens
// by tapping Back to the relevant earlier screen (Visit Info for date/
// times), not inline here. Keeps this screen to "verify in seconds, then
// tap the one dominant Save button," per spec.
///////////////////////////////////////////////////////////////////////////////////////////////
// Function Name: renderConfirmCard
// Description: Renders the final review card before creating or updating a visit.
// Parameters: none
///////////////////////////////////////////////////////////////////////////////////////////////
function renderConfirmCard() {
  document.getElementById("confirmTitle").textContent = isCompletingVisit() ? "Review Completion" : (isEditingVisit() ? "Review Changes" : "Ready To Log");
  document.getElementById("saveEntryBtn").textContent = isCompletingVisit() ? "COMPLETE VISIT" : (isEditingVisit() ? "UPDATE PACE VISIT" : "SAVE PACE VISIT");
  const visitRoomId = (isEditingVisit() || isCompletingVisit()) ? (STATE.editingRoom || STATE.room) : STATE.room;
  const room = CONFIG.ROOMS.find(r => r.id === visitRoomId);
  const roomLine = room ? `${room.label} — ${room.hallway}` : (visitRoomId || "");
  const dur = visitDurationMinutes(STATE.timeIn, STATE.timeOut);
  const card = document.getElementById("confirmCard");
  card.innerHTML = `
    <div class="confirm-row"><span class="confirm-row-label">Room</span><span class="confirm-row-value">${escHtml(roomLine)}</span></div>
    <div class="confirm-row"><span class="confirm-row-label">${STATE.staffMembers.length > 1 ? "Behavior Specialists" : "Behavior Specialist"}</span><span class="confirm-row-value">${STATE.staffMembers.map(escHtml).join("<br>")}</span></div>
    <div class="confirm-row"><span class="confirm-row-label">Student</span><span class="confirm-row-value">${escHtml(STATE.student?.name || "")}</span></div>
    <div class="confirm-row"><span class="confirm-row-label">Teacher Came From</span><span class="confirm-row-value">${escHtml(STATE.cameFromTeacher || "")}</span></div>
    <div class="confirm-row"><span class="confirm-row-label">Visit</span><span class="confirm-row-value">${escHtml(fmtLongDate(STATE.date))}</span></div>
    <div class="confirm-row">
      <span class="confirm-row-label">Time</span>
      <span class="confirm-row-value">${escHtml(fmt12h(STATE.timeIn))} → ${escHtml(fmt12h(STATE.timeOut))}${dur !== null ? ` · ${dur} min` : ""}</span>
    </div>
    <div class="confirm-row"><span class="confirm-row-label">Reason</span><span class="confirm-row-value">${escHtml(STATE.reasons.join(", "))}</span></div>
    <div class="confirm-row"><span class="confirm-row-label">Support</span><span class="confirm-row-value">${escHtml(STATE.supports.join(", "))}</span></div>
    <div class="confirm-row"><span class="confirm-row-label">SCM</span><span class="confirm-row-value">${STATE.scmUsed ? "Yes" : "No"}</span></div>
    ${STATE.notes ? `<div class="confirm-row"><span class="confirm-row-label">Notes</span><span class="confirm-row-value">${escHtml(STATE.notes)}</span></div>` : ""}
  `;
}

document.getElementById("saveEntryBtn").addEventListener("click", async () => {
  if (STATE.saving) return; // never allow a second tap to fire a duplicate submission

  // Defensive re-check (spec: validate "before advancing... OR before
  // final submission") — the Visit Info screen already gated this once,
  // but re-verify STATE directly in case anything upstream changed it.
  if (!STATE.date || !STATE.timeIn || !STATE.timeOut || visitDurationMinutes(STATE.timeIn, STATE.timeOut) === null) {
    if (isCompletingVisit()) {
      renderExitCard();
      nav("exit", "back");
    } else {
      renderVisitInfo();
      nav("visitinfo", "back");
    }
    showToast("Please double-check the visit date and times.", "error");
    return;
  }
  if (isEditingVisit() && (STATE.editingOriginalDate !== todayISODate() || STATE.date !== todayISODate())) {
    showToast("Only today's visits can be edited.", "error");
    return;
  }

  // CURRENT-PATCH: a completed visit may never save with an unknown/blank
  // SCM value (an open live-start visit is the one legitimate exception,
  // but that path saves via saveLiveVisit() above, never this handler).
  // The SCM screen already makes this structurally hard to bypass — the
  // only way forward from it is tapping Yes or No — but this is the same
  // defensive re-check pattern as the date/time guard just above, and
  // directly matches the "must require SCM Yes or No before save" rule.
  if (STATE.scmUsed !== true && STATE.scmUsed !== false) {
    document.getElementById("scmYesBtn").classList.remove("selected");
    document.getElementById("scmNoBtn").classList.remove("selected");
    nav("scm", "back");
    showToast("Please record whether SCM was used.", "error");
    return;
  }

  // CURRENT-PATCH: no completed/edited/after-the-fact PACE visit may ever
  // save with blank or whitespace-only Notes — a live-start visit never
  // reaches this handler at all (see the Notes screen comment above), so
  // this can only ever block a genuinely completed record. Same defensive
  // re-check pattern as the date/time and SCM guards just above: the
  // Notes screen already makes this hard to bypass forward, this only
  // catches it if STATE changed upstream since then. Nothing already
  // entered is cleared — the user lands back on Notes with their draft
  // (specialists, student, reason, support, SCM, times) fully intact.
  if (!STATE.notes || !STATE.notes.trim()) {
    nav("notes", "back");
    renderNotesScreen();
    const el = document.getElementById("notesError");
    if (el) { el.textContent = "Add a brief note before submitting this PACE visit."; el.classList.remove("hidden"); }
    document.getElementById("noteText")?.focus();
    showToast("Add a brief note before submitting this PACE visit.", "error");
    return;
  }

  // STAFF CORRECTIONS: editing an existing visit re-checks PACE
  // authorization at the moment of the write, and refuses a save that
  // would change nothing rather than sending an empty PATCH.
  if (isEditingVisit()) {
    if (!isPaceAuthorized()) {
      showToast("Your account is not authorized to edit PACE visits.", "error");
      return;
    }
    if (!PACE_VISIT_CORRECTIONS.hasChanges(STATE.editOriginal, buildVisitEntry())) {
      showToast("No changes to save. Use Back to change something, or Cancel.", "error");
      return;
    }
  }

  STATE.saving = true;
  const btn = document.getElementById("saveEntryBtn");
  btn.disabled = true; btn.textContent = "Saving…";

  const entry = buildVisitEntry();

  try {
    const wasEditing = isEditingVisit();
    const wasCompleting = isCompletingVisit();
    if (wasCompleting) {
      await PACE_DATA.completeVisit(STATE.completionTarget.id, entry);
    } else if (wasEditing) {
      await PACE_DATA.updateVisit(STATE.editingVisitId, entry);
    } else {
      await PACE_DATA.createVisit(entry);
    }
    btn.disabled = false;
    btn.textContent = wasCompleting ? "COMPLETE VISIT" : (wasEditing ? "UPDATE PACE VISIT" : "SAVE PACE VISIT");
    STATE.saving = false;
    document.getElementById("savedOverlayMessage").textContent = wasCompleting ? "PACE visit completed" : (wasEditing ? "PACE visit updated" : "PACE visit saved");
    document.getElementById("savedOverlay").classList.remove("hidden");
    await refreshRoomVisits();
    setTimeout(async () => {
      document.getElementById("savedOverlay").classList.add("hidden");
      if (wasCompleting) {
        resetTrip();
        nav("room", "back");
        return;
      }
      if (wasEditing) {
        resetTrip();
        nav("recent", "back");
        await loadRecentActivity();
        return;
      }
      // PATCH 004/006: return to Student Search, not all the way back to
      // Room — the same specialist (kept, resetTrip() doesn't touch it)
      // is likely about to log another visit right away. Room +
      // specialist identity stay intact; student/came-from/visit-specific
      // fields are cleared. cachedStudents isn't re-fetched — same
      // pattern as before, avoids an extra round-trip for the common
      // back-to-back case.
      resetTrip();
      document.getElementById("studentSearch").value = "";
      renderStudentGrid();
      nav("student", "back");
    }, 900);
  } catch (err) {
    console.error("PACE save failed:", err);
    STATE.saving = false;
    btn.disabled = false;
    btn.textContent = isCompletingVisit() ? "COMPLETE VISIT" : (isEditingVisit() ? "UPDATE PACE VISIT" : "SAVE PACE VISIT");
    document.getElementById("syncErrorDetail").textContent = err.message || "Please check the connection and try again.";
    document.getElementById("syncErrorOverlay").classList.remove("hidden");
  }
});

document.getElementById("tryAgainBtn").addEventListener("click", () => {
  document.getElementById("syncErrorOverlay").classList.add("hidden");
  document.getElementById("saveEntryBtn").click();
});
document.getElementById("returnFromErrorBtn").addEventListener("click", () => {
  document.getElementById("syncErrorOverlay").classList.add("hidden");
});

/* ── EXIT WORKFLOW ────────────────────────────────────────────────────── */

///////////////////////////////////////////////////////////////////////////////////////////////
// Function Name: openExitScreen
// Description: Opens the completion workflow for a selected open visit.
// Parameters: Object item - open-visit provider row to complete - input
///////////////////////////////////////////////////////////////////////////////////////////////
function openExitScreen(item) {
  if (!item?.id) return;
  resetTrip();
  STATE.workflowMode = PACE_VISIT_WORKFLOW.MODES.COMPLETING;
  STATE.completionTarget = item;
  STATE.exitTarget = item;
  STATE.student = { id: String(item.id), name: item.Student || "" };
  // Read-only context for Mark Complete — the specialist(s) chosen at Start
  // Visit are carried through as-is; this screen never lets staff reselect
  // them (see README "PATCH 010" / brief "Preserve specialists through
  // live visit"). Falls back to the current sticky STATE.staffMembers only
  // if the row itself has nothing recorded (e.g. no column exists yet).
  const storedSpecialists = RECENT_ACTIVITY.specialistNames(item["Behavior Specialist"] || item["Staff Member"]);
  STATE.staffMembers = storedSpecialists.length > 0 ? storedSpecialists : [...STATE.staffMembers];
  STATE.cameFromTeacher = item["Teacher Came From"] || "";
  STATE.date = dateOnly(item.Date) || todayISODate();
  STATE.timeIn = item["Time In"] || "";
  STATE.timeOut = nowHHMM();
  STATE.reasons = RECENT_ACTIVITY.editModel(item).reasons;
  STATE.supports = RECENT_ACTIVITY.editModel(item).supports;
  STATE.scmUsed = typeof item["SCM Used"] === "boolean" ? item["SCM Used"] : null;
  STATE.notes = item.Notes || "";
  STATE.editingRoom = item["PACE Room"] || STATE.room;
  renderExitCard();
  nav("exit", "forward");
}

///////////////////////////////////////////////////////////////////////////////////////////////
// Function Name: renderExitCard
// Description: Renders the read-only context and editable completion time fields.
// Parameters: none
///////////////////////////////////////////////////////////////////////////////////////////////
function renderExitCard() {
  const item = STATE.completionTarget;
  if (!item) return;
  const dur = visitDurationMinutes(STATE.timeIn, STATE.timeOut);
  document.getElementById("exitCard").innerHTML = `
    <div class="confirm-row"><span class="confirm-row-label">Student</span><span class="confirm-row-value">${escHtml(item.Student || "")}</span></div>
    <div class="confirm-row"><span class="confirm-row-label">Entered</span><span class="confirm-row-value">${escHtml(fmt12h(item["Time In"]))}</span></div>
    <div class="confirm-row"><span class="confirm-row-label">Reason</span><span class="confirm-row-value">${escHtml(item.Reason || "")}</span></div>
    ${STATE.staffMembers.length > 0 ? `<div class="confirm-row"><span class="confirm-row-label">${STATE.staffMembers.length > 1 ? "Behavior Specialists" : "Behavior Specialist"}</span><span class="confirm-row-value">${STATE.staffMembers.map(escHtml).join("<br>")}</span></div>` : ""}
    <div class="confirm-row"><span class="confirm-row-label">Duration</span><span class="confirm-row-value">${dur === null ? "—" : `${dur} minute${dur !== 1 ? "s" : ""}`}</span></div>
  `;
  document.getElementById("exitTimeOutInput").value = STATE.timeOut;
}

document.getElementById("exitTimeOutInput").addEventListener("input", event => {
  STATE.timeOut = event.target.value;
  document.getElementById("exitTimeError").classList.add("hidden");
  renderExitCard();
});

document.getElementById("confirmExitBtn").addEventListener("click", () => {
  const item = STATE.completionTarget;
  if (!item) return;
  if (visitDurationMinutes(STATE.timeIn, STATE.timeOut) === null) {
    const error = document.getElementById("exitTimeError");
    error.textContent = "Time Out must be later than Time In.";
    error.classList.remove("hidden");
    return;
  }
  renderChipGrid("supportChips", CONFIG.SUPPORT_OPTIONS, STATE.supports, () => updateNextEnabled("support"));
  updateNextEnabled("support");
  document.querySelector('[data-screen="support"] .btn-back').dataset.nav = "exit";
  nav("support", "forward");
});
document.getElementById("exitTryAgainBtn").addEventListener("click", () => {
  document.getElementById("exitSyncErrorOverlay").classList.add("hidden");
  document.getElementById("confirmExitBtn").click();
});
document.getElementById("exitReturnBtn").addEventListener("click", () => {
  document.getElementById("exitSyncErrorOverlay").classList.add("hidden");
});

/* ── RECENT (privacy-limited cards + same-day edit entry point) ──────── */

///////////////////////////////////////////////////////////////////////////////////////////////
// Function Name: beginEditVisit
// Description: Loads a completed visit into the shared editing workflow. Populates
//              STATE.editOriginal with the pre-edit values and navigates into the workflow
//              once specialist/student/teacher options have loaded.
// Parameters: Object visit - display-name-keyed provider row from Recent Activity - input
///////////////////////////////////////////////////////////////////////////////////////////////
async function beginEditVisit(visit) {
  if (!isPaceAuthorized()) {
    showToast("Your account is not authorized to edit PACE visits.", "error");
    return;
  }
  if (!RECENT_ACTIVITY.isEditableToday(visit, todayISODate())) {
    showToast("Only today's visits can be edited.", "error");
    return;
  }

  const currentRoom = STATE.room;
  const edit = RECENT_ACTIVITY.editModel(visit, {
    fallbackRoom: currentRoom,
    validRooms: CONFIG.ROOMS.map(room => room.id)
  });
  resetTrip();
  STATE.workflowMode = PACE_VISIT_WORKFLOW.MODES.EDITING;
  STATE.editingVisitId = edit.itemId;
  STATE.editingOriginalDate = edit.date;
  STATE.editingRoom = edit.room;
  STATE.submissionId = edit.submissionId;
  STATE.date = edit.date;
  STATE.timeIn = edit.timeIn;
  STATE.timeOut = edit.timeOut;
  STATE.reasons = edit.reasons;
  STATE.supports = edit.supports;
  STATE.scmUsed = edit.scmUsed;
  STATE.notes = edit.notes;
  STATE.cameFromTeacher = edit.cameFromTeacher || null;
  STATE.staffMembers = edit.specialists || [];
  // STAFF CORRECTIONS: remember exactly what staff were shown, so save can
  // write only what they actually change (see PACE_VISIT_CORRECTIONS).
  STATE.editOriginal = {
    paceRoom: edit.room,
    studentName: edit.student,
    date: edit.date,
    timeIn: edit.timeIn,
    timeOut: edit.timeOut,
    behaviors: [...edit.reasons],
    interventions: [...edit.supports],
    scmUsed: edit.scmUsed,
    notes: edit.notes,
    staffMembers: [...(edit.specialists || [])],
    cameFromTeacher: edit.cameFromTeacher || ""
  };
  document.querySelector('[data-screen="specialist"] .btn-back').dataset.nav = "recent";

  const [specialists, students, teachers] = await Promise.allSettled([
    PACE_DATA.getSpecialists(),
    PACE_DATA.getStudents(),
    PACE_DATA.getTeachers()
  ]);
  cachedSpecialists = specialists.status === "fulfilled" ? specialists.value : [];
  cachedStudents = students.status === "fulfilled" ? students.value : [];
  cachedCameFromTeachers = teachers.status === "fulfilled" ? teachers.value : [];

  const studentName = edit.student;
  STATE.student = cachedStudents.find(student => student.name === studentName) || {
    id: `edit-current-${STATE.editingVisitId}`,
    name: studentName || "Student"
  };
  if (!cachedStudents.some(student => student.id === STATE.student.id)) cachedStudents.unshift(STATE.student);

  renderSpecialistGrid();
  renderStudentGrid();
  renderCameFromGrid();
  renderNotesScreen();
  renderVisitInfo();
  nav("visitinfo", "forward");
}

/* ── STAFF CORRECTIONS: delete a visit ──────────────────────────────────
   Two steps, never one: the trash icon only OPENS a confirmation naming the
   visit; nothing is sent until the dialog's Delete button is pressed. The
   destructive write targets the exact SharePoint item id carried by the
   Recent Activity record, is re-authorized at the moment it runs, and the
   list is only changed after SharePoint confirms — a failure leaves the
   record on screen. See visit-corrections.js runDelete(). */

const TRASH_ICON_SVG = `<svg viewBox="0 0 24 24" width="22" height="22" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round" aria-hidden="true"><path d="M3 6h18"/><path d="M8 6V4a1 1 0 0 1 1-1h6a1 1 0 0 1 1 1v2"/><path d="M19 6l-1 14a2 2 0 0 1-2 2H8a2 2 0 0 1-2-2L5 6"/><path d="M10 11v6"/><path d="M14 11v6"/></svg>`;

///////////////////////////////////////////////////////////////////////////////////////////////
// Function Name: requestDeleteVisit
// Description: Opens the destructive-action confirmation for a same-day visit.
// Parameters: Object visit - display-name-keyed provider row to delete - input
///////////////////////////////////////////////////////////////////////////////////////////////
function requestDeleteVisit(visit) {
  if (!isPaceAuthorized()) {
    showToast("Your account is not authorized to delete PACE visits.", "error");
    return;
  }
  if (!RECENT_ACTIVITY.isEditableToday(visit, todayISODate())) {
    showToast("Only today's visits can be deleted.", "error");
    return;
  }
  const summary = PACE_VISIT_CORRECTIONS.deleteSummary(visit, {
    longDate: fmtLongDate,
    time: fmt12h,
    roomLabel: roomId => CONFIG.ROOMS.find(room => room.id === roomId)?.label || roomId
  });
  STATE.deleteTarget = visit;
  document.getElementById("deleteVisitSummary").innerHTML = [summary.student, summary.date, summary.time, summary.room]
    .filter(Boolean)
    .map((line, index) => `<p class="${index === 0 ? "delete-summary-name" : "delete-summary-line"}">${escHtml(line)}</p>`)
    .join("");
  document.getElementById("deleteVisitError").classList.add("hidden");
  document.getElementById("deleteVisitOverlay").classList.remove("hidden");
  document.getElementById("cancelDeleteVisitBtn").focus();
}

///////////////////////////////////////////////////////////////////////////////////////////////
// Function Name: closeDeleteConfirm
// Description: Closes the delete confirmation without modifying provider data.
// Parameters: none
///////////////////////////////////////////////////////////////////////////////////////////////
function closeDeleteConfirm() {
  STATE.deleteTarget = null;
  STATE.deleting = false;
  const confirmBtn = document.getElementById("confirmDeleteVisitBtn");
  confirmBtn.disabled = false;
  confirmBtn.textContent = "Delete Visit";
  document.getElementById("cancelDeleteVisitBtn").disabled = false;
  document.getElementById("deleteVisitOverlay").classList.add("hidden");
}

document.getElementById("cancelDeleteVisitBtn").addEventListener("click", () => {
  if (STATE.deleting) return;
  closeDeleteConfirm();
});

document.getElementById("confirmDeleteVisitBtn").addEventListener("click", async () => {
  if (STATE.deleting || !STATE.deleteTarget) return;
  STATE.deleting = true;
  const confirmBtn = document.getElementById("confirmDeleteVisitBtn");
  const cancelBtn = document.getElementById("cancelDeleteVisitBtn");
  const errorEl = document.getElementById("deleteVisitError");
  confirmBtn.disabled = true;
  confirmBtn.textContent = "Deleting…";
  cancelBtn.disabled = true;
  errorEl.classList.add("hidden");

  const result = await PACE_VISIT_CORRECTIONS.runDelete({
    visit: STATE.deleteTarget,
    confirmed: true, // only this button ever passes true
    isAuthorized: isPaceAuthorized,
    deleteVisit: itemId => PACE_DATA.deleteVisit(itemId),
    onSuccess: async () => {
      closeDeleteConfirm();
      showToast("PACE visit deleted.");
      await loadRecentActivity();
      await refreshRoomVisits();
    }
  });

  if (result.status === "deleted") return;
  STATE.deleting = false;
  confirmBtn.disabled = false;
  confirmBtn.textContent = "Delete Visit";
  cancelBtn.disabled = false;
  errorEl.textContent = result.status === "unauthorized"
    ? "Your account is not authorized to delete PACE visits."
    : result.status === "invalid"
      ? "This visit could not be identified. Nothing was deleted."
      : PACE_VISIT_CORRECTIONS.deleteFailureMessage(result.error);
  errorEl.classList.remove("hidden");
});

///////////////////////////////////////////////////////////////////////////////////////////////
// Function Name: loadRecentActivity
// Description: Reads today's completed visits and renders their privacy-safe cards. The
//              provider performs the bounded read, while RECENT_ACTIVITY enforces completion,
//              sorting, room scope, and the display allow-list. Notes, IDs, email addresses,
//              and raw timestamps are intentionally never rendered here. Provider failures
//              result in the retry/error panel.
// Parameters: none
///////////////////////////////////////////////////////////////////////////////////////////////
async function loadRecentActivity() {
  const listEl = document.getElementById("recentList");
  const refreshBtn = document.getElementById("recentRefreshBtn");
  const scopeEl = document.getElementById("recentScope");
  listEl.innerHTML = `<p class="empty-hint">Loading…</p>`;
  scopeEl.textContent = "Today";
  refreshBtn.disabled = true;
  try {
    const result = await PACE_DATA.getRecentVisits({
      date: todayISODate(),
      room: STATE.room,
      limit: 10
    });
    STATE.paceVisits = result.visits;

    const room = CONFIG.ROOMS.find(item => item.id === STATE.room);
    scopeEl.textContent = result.roomScoped && room
      ? `Today · ${room.label}`
      : "Today · All PACE rooms";

    if (result.visits.length === 0) {
      listEl.innerHTML = `<p class="empty-hint">No PACE activity today yet.</p>`;
      return;
    }

    listEl.innerHTML = result.visits.map((visit, index) => {
      const card = RECENT_ACTIVITY.cardModel(visit, { roomScoped: result.roomScoped });
      const cardRoom = CONFIG.ROOMS.find(room => room.id === card.room)?.label || card.room;
      const timeAndDuration = [
        fmt12h(card.displayTime),
        card.duration === null ? "Duration unavailable" : `${card.duration} min`
      ].join(" · ");
      return `<article class="recent-card">
        <h3 class="recent-card-name">${escHtml(card.student)}</h3>
        <p class="recent-card-time">${escHtml(timeAndDuration)}</p>
        ${card.reason ? `<p class="recent-card-reason">${escHtml(card.reason)}</p>` : ""}
        ${card.specialist ? `<p class="recent-card-specialist">${escHtml(card.specialist)}</p>` : ""}
        ${cardRoom ? `<p class="recent-card-room">${escHtml(cardRoom)}</p>` : ""}
        <div class="recent-card-actions">
          <button class="recent-edit-btn" type="button" data-recent-edit="${index}">Edit today's visit</button>
          <button class="recent-delete-btn" type="button" data-recent-delete="${index}" aria-label="Delete this visit" title="Delete this visit">${TRASH_ICON_SVG}</button>
        </div>
      </article>`;
    }).join("");
    listEl.querySelectorAll("[data-recent-delete]").forEach(button => {
      button.addEventListener("click", () => {
        const visit = result.visits[Number(button.dataset.recentDelete)];
        if (visit) requestDeleteVisit(visit);
      });
    });
    listEl.querySelectorAll("[data-recent-edit]").forEach(button => {
      button.addEventListener("click", async () => {
        const visit = result.visits[Number(button.dataset.recentEdit)];
        if (!visit) return;
        button.disabled = true;
        button.textContent = "Opening…";
        await beginEditVisit(visit);
        if (currentScreenName === "recent") {
          button.disabled = false;
          button.textContent = "Edit today's visit";
        }
      });
    });
  } catch (err) {
    console.error("Failed to load recent activity:", err);
    scopeEl.textContent = "Today";
    listEl.innerHTML = `<p class="empty-hint recent-error">Recent activity could not be loaded. Try again.</p>`;
  } finally {
    refreshBtn.disabled = false;
  }
}

document.querySelector('[data-nav="recent"]').addEventListener("click", () => {
  nav("recent", "forward");
  loadRecentActivity();
});

document.getElementById("recentRefreshBtn").addEventListener("click", loadRecentActivity);

/* ── DEMO: reset control (spec §11 — visible only in demo mode) ──────── */

document.getElementById("resetDemoBtn")?.addEventListener("click", () => {
  if (APP_MODE !== "demo") return; // defensive: this control only exists/works in demo mode
  const confirmed = confirm("Clear all demo PACE records stored on this device?");
  if (!confirmed) return;
  DemoStorage.reset();
  showToast("Demo data cleared.");
  refreshRoomVisits();
  if (currentScreenName === "recent") loadRecentActivity();
});

/* ── TEMPORARY PATCH 003 DIAGNOSTIC (production-only, read-only) ─────────
   Remove this whole block (plus the button/overlay markup in index.html
   and diagnostic.js) once field mapping is reconciled and confirmed. ──── */

document.getElementById("runDiagnosticBtn")?.addEventListener("click", async () => {
  if (APP_MODE === "demo") return; // defensive: button is never even shown in demo mode
  const overlay  = document.getElementById("diagnosticOverlay");
  const status   = document.getElementById("diagnosticStatus");
  const output   = document.getElementById("diagnosticOutput");
  const copyBtn  = document.getElementById("copyDiagnosticBtn");

  overlay.classList.remove("hidden");
  status.textContent = "Running diagnostic (read-only)…";
  output.classList.add("hidden");
  copyBtn.classList.add("hidden");

  try {
    const result = await Diagnostic.run();
    output.value = Diagnostic.formatReport(result);
    output.classList.remove("hidden");
    copyBtn.classList.remove("hidden");
    status.textContent = "Done — nothing was created, changed, or deleted.";
  } catch (err) {
    console.error("Diagnostic failed:", err.message || err);
    status.textContent = `Diagnostic failed: ${err.message || err}`;
  }
});

document.getElementById("copyDiagnosticBtn")?.addEventListener("click", async () => {
  const output = document.getElementById("diagnosticOutput");
  try {
    await navigator.clipboard.writeText(output.value);
    showToast("Diagnostic report copied.");
  } catch {
    // Clipboard API can be blocked in some contexts — fall back to a
    // manual-select so the report is still easy to copy.
    output.removeAttribute("readonly");
    output.focus();
    output.select();
    output.setAttribute("readonly", "");
    showToast("Select-all applied — copy with your keyboard shortcut.");
  }
});

document.getElementById("closeDiagnosticBtn")?.addEventListener("click", () => {
  document.getElementById("diagnosticOverlay").classList.add("hidden");
});

/* ── Misc ─────────────────────────────────────────────────────────────── */

///////////////////////////////////////////////////////////////////////////////////////////////
// Function Name: escHtml
// Description: Escapes dynamic text before inserting it into application HTML.
// Parameters: string str - raw text to HTML-escape - input
///////////////////////////////////////////////////////////////////////////////////////////////
function escHtml(str) {
  return String(str ?? "").replace(/[&<>"']/g, c => ({ "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;", "'": "&#39;" }[c]));
}

if ("serviceWorker" in navigator) {
  window.addEventListener("load", () => navigator.serviceWorker.register("sw.js").catch(err => console.warn("SW registration failed:", err)));
}

boot();
