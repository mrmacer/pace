/* ─────────────────────────────────────────────────────────────────────────
   PACE Room Tracker — App controller
   Horizontal-panel kiosk workflow: Home → Room → Student → Reason →
   Support → SCM → Notes → Confirm → Save, plus Room's own Exit and Recent
   panels. See README.md for the phase-by-phase build notes.
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
  student: () => renderStudentGrid()
};
document.querySelectorAll("[data-nav]").forEach(btn => {
  const target = btn.dataset.nav;
  if (target === "recent") return;
  btn.addEventListener("click", () => {
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
function todayISODate() {
  return new Date().toISOString().slice(0, 10);
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

/* ── App state ────────────────────────────────────────────────────────── */

const STATE = {
  room: null,
  paceVisits: [],       // cached IEP_Pace_Visits rows (display-name keyed) for the active room's day
  student: null,         // roster entry
  reasons: [],
  supports: [],
  scmUsed: null,
  notes: "",
  timeIn: "",
  date: "",
  submissionId: null,
  saving: false,
  exitTarget: null,      // the open-visit row being closed
  duplicateTarget: null  // the open-visit row that blocked a new entry
};

function resetTrip() {
  STATE.student = null;
  STATE.reasons = [];
  STATE.supports = [];
  STATE.scmUsed = null;
  STATE.notes = "";
  STATE.timeIn = "";
  STATE.date = "";
  STATE.submissionId = null;
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
  localStorage.setItem(roomStorageKey(), roomId);
  const label = CONFIG.ROOMS.find(r => r.id === roomId)?.label || roomId;
  document.getElementById("roomTitle").textContent = label.toUpperCase();
  nav("room", direction);
  await refreshRoomVisits();
}

async function refreshRoomVisits() {
  const listEl = document.getElementById("currentlyInPace");
  listEl.innerHTML = `<p class="empty-hint">Loading…</p>`;
  try {
    STATE.paceVisits = await PACE_DATA.getVisits();
  } catch (err) {
    console.error("Failed to load PACE visits:", err);
    listEl.innerHTML = `<p class="empty-hint">Unable to load current activity.</p>`;
    return;
  }
  renderCurrentlyInPace();
}

function openVisitsForRoom(roomId) {
  return STATE.paceVisits
    .filter(v => v["PACE Room"] === roomId && !String(v["Time Out"] || "").trim())
    .sort((a, b) => (a["Time In"] || "").localeCompare(b["Time In"] || ""));
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
      </div>
      <button class="current-card-exit" data-item-id="${escHtml(v.id)}">EXIT →</button>
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

document.getElementById("newStudentBtn").addEventListener("click", () => openStudentScreen());

/* ── STUDENT SELECT ───────────────────────────────────────────────────── */

// Loaded once per visit to the Student screen (via PACE_DATA.getStudents(),
// which itself branches on APP_MODE) and filtered synchronously from here
// on so typing in the search box doesn't re-fetch on every keystroke.
let cachedStudents = [];

async function openStudentScreen() {
  resetTrip();
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

function renderStudentGrid(filter = "") {
  const grid = document.getElementById("studentGrid");
  const q = filter.trim().toLowerCase();
  const students = cachedStudents.filter(s => !q || s.name.toLowerCase().includes(q));
  if (students.length === 0) {
    grid.innerHTML = `<p class="empty-hint">No matching students.</p>`;
    return;
  }
  grid.innerHTML = students.map(s => `<button class="student-card" data-student-id="${escHtml(s.id)}">${escHtml(s.name)}</button>`).join("");
  grid.querySelectorAll(".student-card").forEach(btn => {
    btn.addEventListener("click", () => selectStudent(btn.dataset.studentId));
  });
}

document.getElementById("studentSearch").addEventListener("input", e => renderStudentGrid(e.target.value));

function selectStudent(studentId) {
  const student = cachedStudents.find(s => s.id === studentId);
  if (!student) return;

  // Duplicate protection (spec section 32) — check this room's cached open
  // visits list for the same student before starting a new entry.
  const openForRoom = openVisitsForRoom(STATE.room);
  const existing = openForRoom.find(v => v.Student === student.name);
  if (existing) {
    STATE.duplicateTarget = existing;
    document.getElementById("duplicateMsg").textContent =
      `${student.name} entered PACE at ${fmt12h(existing["Time In"])}.`;
    nav("duplicate", "forward");
    return;
  }

  STATE.student = student;
  STATE.date = todayISODate();
  STATE.timeIn = nowHHMM();
  STATE.submissionId = (typeof crypto !== "undefined" && crypto.randomUUID) ? crypto.randomUUID() : `pace-${Date.now()}`;

  renderChipGrid("reasonChips", CONFIG.REASON_OPTIONS, STATE.reasons, () => updateNextEnabled("reason"));
  updateNextEnabled("reason");
  nav("reason", "forward");
}

document.getElementById("dupViewBtn").addEventListener("click", () => {
  if (STATE.duplicateTarget) openExitScreen(STATE.duplicateTarget);
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
  renderChipGrid("supportChips", CONFIG.SUPPORT_OPTIONS, STATE.supports, () => updateNextEnabled("support"));
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
  setTimeout(() => nav("notes", "forward"), 140); // brief pause so the selection state is visible
}
document.getElementById("scmYesBtn").addEventListener("click", () => setScm(true));
document.getElementById("scmNoBtn").addEventListener("click", () => setScm(false));

/* ── NOTES ────────────────────────────────────────────────────────────── */

document.getElementById("addNoteBtn").addEventListener("click", (e) => {
  document.getElementById("noteText").classList.remove("hidden");
  document.getElementById("noteText").focus();
  e.target.classList.add("hidden");
});
document.getElementById("notesNextBtn").addEventListener("click", () => {
  STATE.notes = document.getElementById("noteText").value.trim();
  renderConfirmCard();
  nav("confirm", "forward");
});

/* ── CONFIRM ──────────────────────────────────────────────────────────── */

function renderConfirmCard() {
  const roomLabel = CONFIG.ROOMS.find(r => r.id === STATE.room)?.label || STATE.room;
  const card = document.getElementById("confirmCard");
  card.innerHTML = `
    <div class="confirm-row"><span class="confirm-row-label">Student</span><span class="confirm-row-value">${escHtml(STATE.student?.name || "")}</span></div>
    <div class="confirm-row"><span class="confirm-row-label">Room</span><span class="confirm-row-value">${escHtml(roomLabel)}</span></div>
    <div class="confirm-row">
      <span class="confirm-row-label">Entered</span>
      <span class="confirm-row-value" id="confirmTimeInDisplay" style="cursor:pointer;text-decoration:underline dotted">${escHtml(fmt12h(STATE.timeIn))} ✏️</span>
    </div>
    <input type="time" id="confirmTimeInEdit" class="hidden" value="${escHtml(STATE.timeIn)}" style="font-size:16px;padding:8px;border-radius:8px;border:1.5px solid var(--border)">
    <div class="confirm-row"><span class="confirm-row-label">Reason</span><span class="confirm-row-value">${escHtml(STATE.reasons.join(", "))}</span></div>
    <div class="confirm-row"><span class="confirm-row-label">Support</span><span class="confirm-row-value">${escHtml(STATE.supports.join(", "))}</span></div>
    <div class="confirm-row"><span class="confirm-row-label">SCM</span><span class="confirm-row-value">${STATE.scmUsed ? "Yes" : "No"}</span></div>
    ${STATE.notes ? `<div class="confirm-row"><span class="confirm-row-label">Note</span><span class="confirm-row-value">${escHtml(STATE.notes)}</span></div>` : ""}
  `;
  const display = document.getElementById("confirmTimeInDisplay");
  const editor  = document.getElementById("confirmTimeInEdit");
  display.addEventListener("click", () => { display.classList.add("hidden"); editor.classList.remove("hidden"); editor.focus(); });
  editor.addEventListener("change", () => {
    if (editor.value) STATE.timeIn = editor.value;
    renderConfirmCard();
  });
}

document.getElementById("saveEntryBtn").addEventListener("click", async () => {
  if (STATE.saving) return; // never allow a second tap to fire a duplicate submission
  STATE.saving = true;
  const btn = document.getElementById("saveEntryBtn");
  btn.disabled = true; btn.textContent = "Saving…";

  const entry = {
    id: STATE.submissionId,
    paceRoom: STATE.room,
    studentName: STATE.student.name,
    date: STATE.date,
    timeIn: STATE.timeIn,
    behaviors: STATE.reasons,
    interventions: STATE.supports,
    scmUsed: STATE.scmUsed,
    notes: STATE.notes,
    submittedByName: AUTH.staffName || AUTH.displayName,
    timestamp: new Date().toISOString()
  };

  try {
    await PACE_DATA.createVisit(entry);
    btn.disabled = false; btn.textContent = "SAVE PACE ENTRY";
    STATE.saving = false;
    document.getElementById("savedOverlay").classList.remove("hidden");
    await refreshRoomVisits();
    setTimeout(() => {
      document.getElementById("savedOverlay").classList.add("hidden");
      resetTrip();
      nav("room", "back");
    }, 900);
  } catch (err) {
    console.error("PACE save failed:", err);
    STATE.saving = false;
    btn.disabled = false; btn.textContent = "SAVE PACE ENTRY";
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
  STATE.exitTarget = item;
  renderExitCard();
  nav("exit", "forward");
}

function renderExitCard() {
  const item = STATE.exitTarget;
  const now = nowHHMM();
  const dur = minutesBetween(item["Time In"] || "00:00", now);
  document.getElementById("exitCard").innerHTML = `
    <div class="confirm-row"><span class="confirm-row-label">Student</span><span class="confirm-row-value">${escHtml(item.Student || "")}</span></div>
    <div class="confirm-row"><span class="confirm-row-label">Entered</span><span class="confirm-row-value">${escHtml(fmt12h(item["Time In"]))}</span></div>
    <div class="confirm-row"><span class="confirm-row-label">Current time</span><span class="confirm-row-value">${escHtml(fmt12h(now))}</span></div>
    <div class="confirm-row"><span class="confirm-row-label">Duration</span><span class="confirm-row-value">${dur} minute${dur !== 1 ? "s" : ""}</span></div>
  `;
}

document.getElementById("confirmExitBtn").addEventListener("click", async () => {
  const item = STATE.exitTarget;
  if (!item) return;
  const btn = document.getElementById("confirmExitBtn");
  btn.disabled = true; btn.textContent = "Closing…";
  const timeOut = nowHHMM();
  try {
    await PACE_DATA.closeVisit(item.id, timeOut);
    btn.disabled = false; btn.textContent = "CONFIRM EXIT";
    showToast(`${item.Student} exited PACE.`);
    STATE.exitTarget = null;
    STATE.duplicateTarget = null;
    await refreshRoomVisits();
    nav("room", "back");
  } catch (err) {
    console.error("PACE exit sync failed:", err);
    btn.disabled = false; btn.textContent = "CONFIRM EXIT";
    document.getElementById("exitSyncErrorDetail").textContent = err.message || "Please check the connection and try again.";
    document.getElementById("exitSyncErrorOverlay").classList.remove("hidden");
  }
});
document.getElementById("exitTryAgainBtn").addEventListener("click", () => {
  document.getElementById("exitSyncErrorOverlay").classList.add("hidden");
  document.getElementById("confirmExitBtn").click();
});
document.getElementById("exitReturnBtn").addEventListener("click", () => {
  document.getElementById("exitSyncErrorOverlay").classList.add("hidden");
});

/* ── RECENT (read-only, today's activity for this room) ──────────────── */

document.querySelector('[data-nav="recent"]').addEventListener("click", async () => {
  nav("recent", "forward");
  const listEl = document.getElementById("recentList");
  listEl.innerHTML = `<p class="empty-hint">Loading…</p>`;
  try {
    STATE.paceVisits = await PACE_DATA.getVisits();
  } catch (err) {
    listEl.innerHTML = `<p class="empty-hint">Unable to load recent activity.</p>`;
    return;
  }
  const today = todayISODate();
  const entries = STATE.paceVisits
    .filter(v => v["PACE Room"] === STATE.room && dateOnly(v.Date) === today)
    .sort((a, b) => (b["Submitted At"] || "").localeCompare(a["Submitted At"] || ""))
    .slice(0, 20);

  if (entries.length === 0) {
    listEl.innerHTML = `<p class="empty-hint">No PACE activity today yet.</p>`;
    return;
  }
  listEl.innerHTML = entries.map(v => {
    const timeOut = String(v["Time Out"] || "").trim();
    const dur = timeOut ? `${minutesBetween(v["Time In"], timeOut)} min` : "open";
    const subParts = [`${fmt12h(v["Time In"])} – ${timeOut ? fmt12h(timeOut) : "—"}`];
    if (v.Behavior) subParts.push(v.Behavior);
    if (v["SCM Used"]) subParts.push("SCM");
    return `<div class="recent-row">
      <div class="recent-row-top"><span>${escHtml(v.Student || "")}${v.demo ? ' <span class="badge-simulated">SIMULATED</span>' : ""}</span><span>${dur}</span></div>
      <div class="recent-row-sub">${escHtml(subParts.join(" · "))}</div>
    </div>`;
  }).join("");
});

/* ── DEMO: reset control (spec §11 — visible only in demo mode) ──────── */

document.getElementById("resetDemoBtn")?.addEventListener("click", () => {
  if (APP_MODE !== "demo") return; // defensive: this control only exists/works in demo mode
  const confirmed = confirm("Clear all demo PACE records stored on this device?");
  if (!confirmed) return;
  DemoStorage.reset();
  showToast("Demo data cleared.");
  refreshRoomVisits();
  if (currentScreenName === "recent") document.querySelector('[data-nav="recent"]').click();
});

/* ── Misc ─────────────────────────────────────────────────────────────── */

function escHtml(str) {
  return String(str ?? "").replace(/[&<>"']/g, c => ({ "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;", "'": "&#39;" }[c]));
}

if ("serviceWorker" in navigator) {
  window.addEventListener("load", () => navigator.serviceWorker.register("sw.js").catch(err => console.warn("SW registration failed:", err)));
}

boot();
