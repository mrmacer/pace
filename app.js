/* ─────────────────────────────────────────────────────────────────────────
   PACE Room Tracker — App controller
   Horizontal-panel kiosk workflow: Home → Room → Specialist → Student
   (search-first) → Teacher Came From → Visit Info → Reason → Support →
   SCM → Notes → Confirm → Save, plus Room's own Exit and Recent panels.
   Save returns to Student Search (not Room) — see the save handler.
   PATCH 006 removed the old homeroom-Teacher-grouped Student screen; see
   README.md for the phase-by-phase build notes.
   ───────────────────────────────────────────────────────────────────────── */

/* ── Screen navigation (deterministic slide, no history stack needed —
   every call states its own direction) ─────────────────────────────────── */

const SCREENS = {};
document.querySelectorAll(".screen").forEach(el => { SCREENS[el.dataset.screen] = el; });
let currentScreenName = "signin";

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
function showToast(msg, kind = "ok") {
  const el = document.getElementById("toast");
  el.textContent = msg;
  el.classList.toggle("error", kind === "error");
  el.classList.remove("hidden");
  clearTimeout(toastTimer);
  toastTimer = setTimeout(() => el.classList.add("hidden"), 2600);
}

/* ── Time helpers ─────────────────────────────────────────────────────── */

function nowHHMM() {
  const d = new Date();
  return `${String(d.getHours()).padStart(2, "0")}:${String(d.getMinutes()).padStart(2, "0")}`;
}
// PATCH 001: this used to be `new Date().toISOString().slice(0,10)`, which
// converts to UTC first — in any US timezone that rolls to the WRONG local
// calendar date in the evening (e.g. 8:30 PM Eastern is already past
// midnight UTC). Build the date from local getters instead so "today"
// always means the staff member's local today, never a UTC-shifted one.
function todayISODate() {
  const d = new Date();
  return `${d.getFullYear()}-${String(d.getMonth() + 1).padStart(2, "0")}-${String(d.getDate()).padStart(2, "0")}`;
}
// "August 20, 2026" from a "YYYY-MM-DD" string, parsed as LOCAL date parts
// (never `new Date("YYYY-MM-DD")`, which parses as UTC midnight and can
// print the wrong day in negative-UTC-offset timezones — the same class of
// bug todayISODate() above was fixed for).
function fmtLongDate(dateStr) {
  const d = dateOnly(dateStr);
  if (!d) return "—";
  const [y, m, day] = d.split("-").map(Number);
  return new Date(y, m - 1, day).toLocaleDateString(undefined, { year: "numeric", month: "long", day: "numeric" });
}
// SharePoint may hand back "Date" as a bare "YYYY-MM-DD" (what this app
// writes) or as a full ISO datetime (if the column is a true Date/Time
// field) — normalize to just the date part before comparing/parsing it.
function dateOnly(value) {
  return String(value || "").slice(0, 10);
}
function fmt12h(hhmm) {
  if (!hhmm || !hhmm.includes(":")) return "—";
  let [h, m] = hhmm.split(":").map(Number);
  const ampm = h >= 12 ? "PM" : "AM";
  h = h % 12; if (h === 0) h = 12;
  return `${h}:${String(m).padStart(2, "0")} ${ampm}`;
}
// Elapsed minutes between an "HH:MM" time-in (assumed today, or the given
// date) and now. Handles the rare midnight-crossing open visit gracefully.
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
function visitDurationMinutes(timeIn, timeOut) {
  return PACE_VISIT_WORKFLOW.durationMinutes(timeIn, timeOut);
}

/* ── App state ────────────────────────────────────────────────────────── */

const STATE = {
  room: null,
  paceVisits: [],       // cached today's IEP_Pace_Visits rows (display-name keyed)
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
  student: null,         // roster entry
  // PATCH 006: who the student was physically with immediately before
  // this PACE visit — NOT the roster's homeroom `student.teacher` (that's
  // preserved on the roster object but no longer drives navigation; see
  // README). Deliberately its own property, not reused from the old
  // homeroom-grouping STATE.teacher (removed this patch), so the two
  // concepts can never be confused.
  cameFromTeacher: null,
  reasons: [],
  supports: [],
  scmUsed: null,
  notes: "",
  timeIn: "",
  timeOut: "",  // PATCH 001: collected on the Visit Info screen, before save
  date: "",
  submissionId: null,
  workflowMode: null,
  completionTarget: null,
  editingVisitId: null,
  editingOriginalDate: "",
  editingRoom: "",
  saving: false,
  exitTarget: null,      // the open-visit row being closed
  duplicateTarget: null  // the open-visit row that blocked a new entry
};

// Clears everything specific to ONE visit-in-progress. Deliberately does
// NOT touch STATE.staffMembers (sticky across a room session — see STATE
// declaration above) or STATE.room.
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
  STATE.exitTarget = null;
  STATE.duplicateTarget = null;
}

function isEditingVisit() {
  return Boolean(STATE.editingVisitId);
}

function isLiveStart() {
  return STATE.workflowMode === PACE_VISIT_WORKFLOW.MODES.LIVE_START;
}

function isCompletingVisit() {
  return STATE.workflowMode === PACE_VISIT_WORKFLOW.MODES.COMPLETING;
}

function beginVisitWorkflow(mode) {
  resetTrip();
  STATE.workflowMode = mode;
  openSpecialistScreen({ preserveTrip: true });
}

// "Remember last room" uses a demo-specific key in demo mode (spec: keep
// simulated state fully separate from anything a production deployment
// would persist), and the real key otherwise.
function roomStorageKey() {
  return APP_MODE === "demo" ? DEMO_CONFIG.roomStorageKey : CONFIG.STORAGE_KEYS.LAST_ROOM;
}

/* ── Boot ─────────────────────────────────────────────────────────────── */

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
    document.getElementById("unauthorizedMsg").textContent = AUTH.lookupError || "Your account is not approved for PACE Room Tracker.";
    nav("unauthorized", "forward");
    return;
  }
  if (!AUTH.isAuthenticated) {
    // account present but staff lookup didn't resolve; loadStaffFromSharePoint already ran in AUTH.init
    document.getElementById("unauthorizedMsg").textContent = AUTH.lookupError || "Unable to verify your account.";
    nav("unauthorized", "forward");
    return;
  }

  // PATCH 002: production's subtle identity indicator (Home screen only —
  // see index.html). Never the raw email, and demoBanner is never touched
  // on this path, so the two modes' indicators can't ever both show.
  const indicator = document.getElementById("signedInIndicator");
  if (indicator) {
    indicator.textContent = `Signed in as ${AUTH.staffName || "you"}`;
    indicator.classList.remove("hidden");
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

/* ── HOME → ROOM ──────────────────────────────────────────────────────── */

document.querySelectorAll(".room-card").forEach(btn => {
  btn.addEventListener("click", () => enterRoom(btn.dataset.room, "forward"));
});

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
  nav("room", direction);
  await refreshRoomVisits();
}

// Keeps every `.room-badge` element (one per workflow screen) in sync with
// the currently selected room — "PACE ROOM 1 · YELLOW HALL" — a single
// function driving every instance rather than per-screen logic.
function updateRoomBadges() {
  const room = CONFIG.ROOMS.find(r => r.id === STATE.room);
  const text = room ? `${room.label.toUpperCase()} · ${room.hallway.toUpperCase()}` : "";
  document.querySelectorAll(".room-badge").forEach(el => { el.textContent = text; });
}

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

function openVisitsForRoom(roomId) {
  return PACE_VISIT_WORKFLOW.openForRoom(STATE.paceVisits, roomId);
}

function renderCurrentlyInPace() {
  const listEl = document.getElementById("currentlyInPace");
  const open = openVisitsForRoom(STATE.room);
  if (open.length === 0) {
    listEl.innerHTML = `<p class="empty-hint">No students currently in PACE.</p>`;
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

function updateSpecialistContinueState() {
  const count = STATE.staffMembers.length;
  document.getElementById("specialistSelectedCount").textContent =
    count === 0 ? "" : `${count} selected`;
  document.getElementById("specialistContinueBtn").disabled = count === 0;
}

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
const STUDENT_SUGGESTION_LIMIT = 8;

// Lowercase, strip punctuation ("Ja'de" still matches "jade"), collapse to
// a plain space-joined string — shared by Student and Teacher Came From
// search so both behave the same way.
function normalizeSearchText(str) {
  return String(str || "").toLowerCase().replace(/[^a-z0-9\s]/g, "").trim();
}

// Matches first name, last name, "First Last", AND "Last First" — not
// just the single combined display name — so a last-name-only or
// reversed-order search still finds the right student.
function studentSearchCorpus(s) {
  const parts = [s.name, s.firstName, s.lastName];
  if (s.firstName || s.lastName) parts.push(`${s.lastName || ""} ${s.firstName || ""}`);
  return normalizeSearchText(parts.filter(Boolean).join(" "));
}

async function openStudentScreen() {
  document.getElementById("specialistContextLine").textContent = STATE.staffMembers.join(", ");
  document.getElementById("studentSearch").value = "";
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

function studentCardHtml(s) {
  return `<button class="student-card${STATE.student?.id === s.id ? " selected" : ""}" data-student-id="${escHtml(s.id)}">${escHtml(s.name)}</button>`;
}
function wireStudentCards(grid) {
  grid.querySelectorAll(".student-card").forEach(btn => {
    btn.addEventListener("click", () => selectStudent(btn.dataset.studentId));
  });
}

function renderStudentGrid(filter = "") {
  const grid = document.getElementById("studentGrid");
  document.getElementById("specialistContextLine").textContent = STATE.staffMembers.join(", ");
  const q = normalizeSearchText(filter);

  // Blank search: a bounded set of alphabetical suggestions, never the
  // whole roster — "staff should not need to scroll through the entire
  // roster" is the hard requirement here, so this stays capped regardless
  // of how large the real production roster grows.
  if (!q) {
    let suggestions = cachedStudents.slice().sort((a, b) => a.name.localeCompare(b.name)).slice(0, STUDENT_SUGGESTION_LIMIT);
    if (isEditingVisit() && STATE.student) {
      suggestions = [STATE.student, ...suggestions.filter(s => s.id !== STATE.student.id)].slice(0, STUDENT_SUGGESTION_LIMIT);
    }
    grid.innerHTML = `<p class="empty-hint">Start typing a student's name…</p>` + suggestions.map(studentCardHtml).join("");
    wireStudentCards(grid);
    return;
  }

  const students = cachedStudents.filter(s => studentSearchCorpus(s).includes(q));
  if (students.length === 0) {
    grid.innerHTML = `<p class="empty-hint">No matching PACE-enabled student found.</p>`;
    return;
  }
  grid.innerHTML = students.map(studentCardHtml).join("");
  wireStudentCards(grid);
}

document.getElementById("studentSearch").addEventListener("input", e => renderStudentGrid(e.target.value));

function selectStudent(studentId) {
  const student = cachedStudents.find(s => s.id === studentId);
  if (!student) return;

  if (isLiveStart()) {
    // Duplicate protection applies to a new visit, never to the completed
    // row currently being edited.
    const openForRoom = openVisitsForRoom(STATE.room);
    const existing = openForRoom.find(v => v.Student === student.name);
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

function updateVisitDurationHint() {
  const hintEl = document.getElementById("visitDurationHint");
  const dur = visitDurationMinutes(STATE.timeIn, STATE.timeOut);
  hintEl.textContent = dur !== null ? `Duration: ${dur} minute${dur !== 1 ? "s" : ""}` : "";
}

// Shared by the Next button AND the defensive re-check right before final
// save (spec: validate "before advancing... OR before final submission").
// Returns a human-readable message, or null if everything's valid.
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

function renderNotesScreen() {
  const noteText = document.getElementById("noteText");
  const addNoteBtn = document.getElementById("addNoteBtn");
  noteText.value = STATE.notes || "";
  const showEditor = isEditingVisit() || Boolean(STATE.notes);
  noteText.classList.toggle("hidden", !showEditor);
  addNoteBtn.classList.toggle("hidden", showEditor);
}

document.getElementById("addNoteBtn").addEventListener("click", (e) => {
  document.getElementById("noteText").classList.remove("hidden");
  document.getElementById("noteText").focus();
  e.target.classList.add("hidden");
});
document.getElementById("noteText").addEventListener("input", event => {
  STATE.notes = event.target.value;
});
document.getElementById("notesNextBtn").addEventListener("click", () => {
  STATE.notes = document.getElementById("noteText").value.trim();
  renderConfirmCard();
  nav("confirm", "forward");
});

/* ── CONFIRM ──────────────────────────────────────────────────────────── */

function buildVisitEntry({ open = false } = {}) {
  return {
    id: STATE.submissionId,
    paceRoom: isEditingVisit() ? (STATE.editingRoom || STATE.room) : STATE.room,
    staffMembers: [...STATE.staffMembers],
    cameFromTeacher: STATE.cameFromTeacher,
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
    timestamp: new Date().toISOString()
  };
}

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
    await PACE_DATA.createVisit(buildVisitEntry({ open: true }));
    await refreshRoomVisits();
    const studentName = STATE.student.name;
    resetTrip();
    btn.disabled = false;
    btn.textContent = "START PACE VISIT";
    STATE.saving = false;
    nav("room", "back");
    showToast(`${studentName} is now in PACE.`);
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

async function beginEditVisit(visit) {
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
        <button class="recent-edit-btn" type="button" data-recent-edit="${index}">Edit today's visit</button>
      </article>`;
    }).join("");
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

function escHtml(str) {
  return String(str ?? "").replace(/[&<>"']/g, c => ({ "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;", "'": "&#39;" }[c]));
}

if ("serviceWorker" in navigator) {
  window.addEventListener("load", () => navigator.serviceWorker.register("sw.js").catch(err => console.warn("SW registration failed:", err)));
}

boot();
