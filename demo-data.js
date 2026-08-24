/* ─────────────────────────────────────────────────────────────────────────
   PACE Room Tracker — Demo mode data (no IU29 data, ever)

   Everything in this file is fake: a simulated roster, a simulated staff
   identity, and a localStorage-backed record store. Nothing here reads or
   writes anything real. See config.js for APP_MODE / DEMO_CONFIG and
   pace-data.js for how this plugs into the same UI code path production
   uses.
   ───────────────────────────────────────────────────────────────────────── */

// PATCH 004: `teacher` added so demo mirrors the real Teacher → Student
// grouping (roster.js now reads the same live column). Fake teacher names,
// same as the students — nothing here is a real IU29 person.
const DEMO_STUDENTS = [
  { id: "SIM-001", name: "Alex R.",   teacher: "Mrs. Ashford", active: true },
  { id: "SIM-002", name: "Jordan M.", teacher: "Mrs. Ashford", active: true },
  { id: "SIM-003", name: "Taylor S.", teacher: "Mr. Bellamy",  active: true },
  { id: "SIM-004", name: "Casey L.",  teacher: "Mr. Bellamy",  active: true },
  { id: "SIM-005", name: "Morgan T.", teacher: "Ms. Castillo", active: true },
  { id: "SIM-006", name: "Riley C.",  teacher: "Ms. Castillo", active: true }
];

const DEMO_USER = {
  id: "demo-user",
  name: "Demo Staff",
  role: "PACE Staff"
};

// PATCH 005: Behavior Specialists now come from a real SharePoint people
// list (IEP_Users2, Role = "Behavior Specialist") in production — so, same
// as DEMO_STUDENTS, demo mode gets its OWN fake names rather than reusing
// the real 12 from CONFIG.BEHAVIOR_SPECIALISTS (that constant is now a
// production-only fallback — see pace-data.js).
const DEMO_SPECIALISTS = ["Dana Fielding", "Marcus Webb", "Priya Anand"];

// PATCH 006: simulates PACE_DATA.getTeachers()'s two-source production
// combination (IEP_Users2 Role="Teacher" + unique roster Teacher values) —
// "Mr. Delgado" has no PACE-eligible students of his own (unlike the other
// three, who are also each a DEMO_STUDENTS homeroom teacher), so the
// dedupe/merge behavior gets exercised the same way real data would.
const DEMO_CAME_FROM_TEACHERS = ["Mrs. Ashford", "Mr. Bellamy", "Ms. Castillo", "Mr. Delgado"];

const DemoStorage = {
  load() {
    try {
      return JSON.parse(localStorage.getItem(DEMO_CONFIG.storageKey)) || { paceVisits: [] };
    } catch {
      return { paceVisits: [] };
    }
  },

  save(data) {
    localStorage.setItem(DEMO_CONFIG.storageKey, JSON.stringify(data));
  },

  getVisits() {
    return this.load().paceVisits || [];
  },

  // `entry` is the exact same shape app.js already builds for
  // GRAPH.savePaceVisit() (paceRoom, studentName, date, timeIn, timeOut,
  // behaviors[], interventions[], scmUsed, notes, submittedByName,
  // timestamp). Stored using the same SharePoint *display*-field-name keys
  // GRAPH uses, so every render function in app.js (which reads
  // v["PACE Room"], v.Student, v["Time In"], etc.) works identically
  // against demo and real data — only this file and pace-data.js know demo
  // mode exists.
  //
  // PATCH 001: completed-visit model — entry.timeOut now arrives already
  // filled in (collected before save, not added later via updateVisit).
  //
  // PATCH 003 (demo parity): renamed "Behavior"->"Reason" and
  // "Interventions"->"Intervention Used", and added "Duration", to match
  // GRAPH.savePaceVisit()'s live-schema-reconciled keys — app.js's render
  // code reads these by the same names for both modes, so the two must
  // stay in lockstep. "PACE Room"/"SCM Used"/"Submitted By"/"Submitted At"
  // are kept here even though the live SharePoint list currently has no
  // matching column for them (see README) — demo still models the full
  // logical visit regardless of what production can persist today.
  createVisit(entry) {
    const data = this.load();
    const id = entry.id || crypto.randomUUID();

    // Mirror GRAPH.savePaceVisit's duplicate-submission guard so a retried
    // tap can't create two rows for the same entry.
    const existing = data.paceVisits.find(v => v.id === id);
    if (existing) return { duplicatePrevented: true, existingItem: existing };

    const visit = {
      id,
      demo: true,
      "PACE Room":          entry.paceRoom || "",
      "Student":            entry.studentName || "",
      "Date":               entry.date || "",
      "Time In":            entry.timeIn || "",
      "Time Out":           entry.timeOut || "",
      "Duration":           entry.durationMinutes ?? null,
      "Reason":             Array.isArray(entry.behaviors) ? entry.behaviors.join(", ") : (entry.behaviors || ""),
      "Intervention Used":  Array.isArray(entry.interventions) ? entry.interventions.join(", ") : (entry.interventions || ""),
      "SCM Used":           entry.scmUsed === true,
      "Notes":              entry.notes || "",
      // PATCH 004: kept here for demo parity/completeness even though
      // production doesn't persist "Staff Member" today (it's a
      // Person-type column this project has no write infrastructure for —
      // see README). Demo still models the full logical visit.
      "Staff Member":       entry.staffMember || "",
      // PATCH 006: replaces the old homeroom "Teacher" (STATE.teacher,
      // removed along with the homeroom-grouped Student screen) — this is
      // specifically who the student physically came from for THIS visit,
      // not their roster homeroom teacher.
      "Teacher Came From":  entry.cameFromTeacher || "",
      "Submitted By":       entry.submittedByName || DEMO_USER.name,
      "Submitted At":       entry.timestamp || new Date().toISOString(),
      createdAt:            new Date().toISOString()
    };

    data.paceVisits.push(visit);
    this.save(data);
    return visit;
  },

  // `patch` uses the same display-field-name keys as the stored row, e.g.
  // { "Time Out": "14:05" } — kept consistent with createVisit() above.
  updateVisit(id, patch) {
    const data = this.load();
    const visit = data.paceVisits.find(item => item.id === id);
    if (!visit) throw new Error("Demo visit not found.");
    Object.assign(visit, patch, { modifiedAt: new Date().toISOString() });
    this.save(data);
    return visit;
  },

  reset() {
    localStorage.removeItem(DEMO_CONFIG.storageKey);
  }
};
