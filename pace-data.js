/* ─────────────────────────────────────────────────────────────────────────
   PACE Room Tracker — Unified data adapter

   The single interface app.js talks to. Every read/write branches on
   APP_MODE here, in exactly one place — app.js never checks APP_MODE or
   calls GRAPH/DemoStorage directly, so the two modes render and behave
   identically (spec: "The UX should be identical to production mode. Only
   the backend changes.").
   ───────────────────────────────────────────────────────────────────────── */

// PATCH 010: same comma-join convention as Reason/Intervention Used
// (entry.behaviors.join(", ") etc., below and in graph.js) — see
// recent-activity.js's specialistNames()/selections() for the read side.
function joinSpecialists(staffMembers) {
  return Array.isArray(staffMembers) ? staffMembers.filter(Boolean).join(", ") : (staffMembers || "");
}

function paceDataTodayISODate() {
  const now = new Date();
  return `${now.getFullYear()}-${String(now.getMonth() + 1).padStart(2, "0")}-${String(now.getDate()).padStart(2, "0")}`;
}

// The current production list may not yet contain writable text columns
// for room/specialist/teacher. Keep only that non-student operational
// context on this iPad, keyed by the SharePoint item id, so a just-created
// open visit remains room-scoped after refresh. SharePoint remains the
// visit source of truth; this never stores student names or visit details.
function loadVisitContext() {
  try { return JSON.parse(localStorage.getItem(CONFIG.STORAGE_KEYS.VISIT_CONTEXT)) || {}; }
  catch { return {}; }
}

function rememberVisitContext(itemId, entry) {
  if (!itemId || APP_MODE === "demo") return;
  const context = loadVisitContext();
  context[String(itemId)] = {
    room: entry.paceRoom || "",
    specialist: joinSpecialists(entry.staffMembers),
    teacher: entry.cameFromTeacher || ""
  };
  localStorage.setItem(CONFIG.STORAGE_KEYS.VISIT_CONTEXT, JSON.stringify(context));
}

function enrichVisitContext(visits) {
  if (APP_MODE === "demo") return visits;
  const context = loadVisitContext();
  return visits.map(visit => {
    const saved = context[String(visit.id)] || {};
    return {
      ...visit,
      "PACE Room": visit["PACE Room"] || saved.room || "",
      "Behavior Specialist": visit["Behavior Specialist"] || visit["Staff Member"] || saved.specialist || "",
      "Teacher Came From": visit["Teacher Came From"] || saved.teacher || ""
    };
  });
}

// CURRENT-PATCH: the read-side half of the room slug<->label boundary
// conversion (write-side: graph.js's savePaceVisit()/updatePaceVisit(),
// demo-data.js's createVisit()). Every consumer of a visit object —
// visit-workflow.js's openForRoom(), recent-activity.js's room filter and
// editModel(), app.js's openExitScreen()/beginEditVisit() — compares
// visit["PACE Room"] against a raw slug (STATE.room/STATE.editingRoom).
// Normalizing here, in the one function both getVisits() callers share,
// means none of those call sites need to know the live SharePoint value is
// actually a label. paceRoomIdForLabel() is a no-op on an already-slug
// value (e.g. the local visit-context fallback in enrichVisitContext(), or
// a legacy row saved before this patch), so this is safe to apply
// unconditionally to both modes.
function normalizeRoomOnRead(visits) {
  return visits.map(visit => ({
    ...visit,
    "PACE Room": paceRoomIdForLabel(visit["PACE Room"])
  }));
}

const PACE_DATA = {
  // Active, PACE-eligible students. Demo students are always considered
  // PACE-eligible (there's no separate "PACE Enabled" concept in the fake
  // roster); production defers to ROSTER's real Active + PACE Enabled
  // filtering, loading it on first use if it hasn't been already.
  async getStudents() {
    if (APP_MODE === "demo") {
      return DEMO_STUDENTS.filter(s => s.active);
    }
    if (!ROSTER.loaded) await ROSTER.refresh();
    return ROSTER.getPaceEnabled();
  },

  // PACE visit rows, display-field-name keyed (see demo-data.js for why
  // the shapes match). Supplying a date keeps the shared iPad read bounded
  // to that local calendar day; production applies the date in Graph and
  // demo filters only the localStorage copy.
  async getVisits({ date } = {}) {
    if (APP_MODE === "demo") {
      const visits = DemoStorage.getVisits();
      const filtered = date
        ? visits.filter(visit => String(visit.Date || "").slice(0, 10) === String(date).slice(0, 10))
        : visits;
      return normalizeRoomOnRead(filtered);
    }
    const visits = date
      ? await GRAPH.getPaceVisitsForDateByDisplayName(date)
      : await GRAPH.getPaceVisitsByDisplayName();
    return normalizeRoomOnRead(enrichVisitContext(visits));
  },

  // PATCH 008: the existing provider remains the only UI data boundary.
  // Recent Activity receives today's bounded rows, then applies the tested
  // completed/newest/room rules in RECENT_ACTIVITY. The `roomScoped` flag
  // tells the UI whether room labels would be redundant.
  async getRecentVisits({ date, room, limit = 10 } = {}) {
    const visits = await this.getVisits({ date });
    return RECENT_ACTIVITY.select(visits, { date, room, limit });
  },

  async createVisit(entry) {
    if (APP_MODE === "demo") return DemoStorage.createVisit(entry);
    const result = await GRAPH.savePaceVisit(entry);
    rememberVisitContext(result?.id, entry);
    return result;
  },

  // Same-day correction path. The UI can only discover today's rows, and
  // the provider independently rejects any payload that is not for the
  // device's local today before choosing localStorage or Graph.
  async updateVisit(id, entry) {
    if (!id) throw new Error("A visit is required for editing.");
    if (String(entry?.date || "").slice(0, 10) !== paceDataTodayISODate()) {
      throw new Error("Only today's visits can be edited.");
    }
    if (APP_MODE === "demo") {
      return DemoStorage.updateVisit(id, {
        // CURRENT-PATCH: label, matching createVisit() and production's
        // updatePaceVisit() — see config.js's paceRoomLabelForId().
        "PACE Room":         paceRoomLabelForId(entry.paceRoom),
        "Student":           entry.studentName || "",
        "Date":              entry.date || "",
        "Time In":           entry.timeIn || "",
        "Time Out":          entry.timeOut || "",
        "Duration":          entry.durationMinutes ?? null,
        "Reason":            Array.isArray(entry.behaviors) ? entry.behaviors.join(", ") : (entry.behaviors || ""),
        "Intervention Used": Array.isArray(entry.interventions) ? entry.interventions.join(", ") : (entry.interventions || ""),
        // CURRENT-PATCH fix: null-safe, matching completeVisit() below —
        // was coercing an unset SCM to `false` (see graph.js's identical fix).
        "SCM Used":          entry.scmUsed == null ? null : entry.scmUsed === true,
        "Notes":             entry.notes || "",
        "Staff Member":      joinSpecialists(entry.staffMembers),
        "Teacher Came From": entry.cameFromTeacher || ""
      });
    }
    const result = await GRAPH.updatePaceVisit(id, entry);
    rememberVisitContext(id, entry);
    return result;
  },

  // Completes an existing open record. This path never calls createVisit:
  // Time Out and completion-only fields are patched onto the same id.
  async completeVisit(id, entry) {
    if (!id) throw new Error("An open visit is required for completion.");
    const patch = {
      "Time Out":          entry.timeOut || "",
      "Duration":          entry.durationMinutes ?? null,
      "Intervention Used": Array.isArray(entry.interventions) ? entry.interventions.join(", ") : (entry.interventions || ""),
      "SCM Used":          entry.scmUsed == null ? null : entry.scmUsed === true,
      "Notes":             entry.notes || ""
    };
    if (APP_MODE === "demo") return DemoStorage.updateVisit(id, patch);
    return GRAPH.completePaceVisit(id, entry);
  },

  async closeVisit(id, timeOut) {
    if (APP_MODE === "demo") return DemoStorage.updateVisit(id, { "Time Out": timeOut });
    return GRAPH.closePaceVisit(id, timeOut);
  },

  // PATCH 005: Behavior Specialist names, sorted alphabetically. Demo uses
  // its own fake roster (DEMO_SPECIALISTS); production loads Active rows
  // with Role = "Behavior Specialist" from IEP_Users2 via the new
  // GRAPH.getActiveUsersByRole(). Falls back to the old hardcoded
  // CONFIG.BEHAVIOR_SPECIALISTS list ONLY if the live call fails or
  // returns nothing (e.g. the IU29 Behavior Specialist rows haven't been
  // added to IEP_Users2 yet) — this fallback is meant to be TEMPORARY;
  // remove CONFIG.BEHAVIOR_SPECIALISTS and this catch once dynamic
  // loading is confirmed working against real production data (can't be
  // verified from here — no live Graph session in this environment).
  async getSpecialists() {
    if (APP_MODE === "demo") return DEMO_SPECIALISTS.slice();
    try {
      const specialists = await GRAPH.getActiveUsersByRole("Behavior Specialist");
      if (specialists.length > 0) return specialists.map(s => s.name);
      console.warn("IEP_Users2 returned zero Active \"Behavior Specialist\" rows — falling back to CONFIG.BEHAVIOR_SPECIALISTS.");
    } catch (err) {
      console.error("Failed to load Behavior Specialists from IEP_Users2:", err.message || err);
    }
    return CONFIG.BEHAVIOR_SPECIALISTS.slice();
  },

  // PATCH 006 originally combined this with unique Teacher values from the
  // student roster ("Source B"). PATCH 007 removes that: production data
  // confirmed IEP_Users2 and the roster's Teacher column use different
  // naming conventions for the same person (e.g. "Andruchek" vs.
  // "Andruchek, B."), which surfaced as apparent duplicates in the picker.
  // Surnames can legitimately collide between two different people, so no
  // fuzzy/surname-based merge is attempted here or anywhere in this
  // codebase — instead, IEP_Users2 is now the SOLE authoritative source.
  // Active "Teacher"-role rows only, sorted alphabetically by Name.
  //
  // If this query fails or returns nothing, this returns an empty array —
  // deliberately NOT falling back to roster Teacher values, since that
  // fallback is exactly what would reintroduce the duplicate/inconsistent
  // naming problem. The UI still always offers "Other / Not Listed"
  // regardless of this list's length (see app.js's renderCameFromGrid()),
  // so the workflow is never blocked by an empty or failed query.
  //
  // Adding a teacher later still needs no code change — see README
  // "Future teacher additions."
  async getTeachers() {
    if (APP_MODE === "demo") return DEMO_CAME_FROM_TEACHERS.slice();
    try {
      const teachers = await GRAPH.getActiveUsersByRole("Teacher");
      return teachers.map(t => t.name).sort((a, b) => a.localeCompare(b));
    } catch (err) {
      console.error("Failed to load Teacher-role users from IEP_Users2:", err.message || err);
      return [];
    }
  }
};
