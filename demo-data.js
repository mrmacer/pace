/* ─────────────────────────────────────────────────────────────────────────
   PACE Room Tracker — Demo mode data (no IU29 data, ever)

   Everything in this file is fake: a simulated roster, a simulated staff
   identity, and a localStorage-backed record store. Nothing here reads or
   writes anything real. See config.js for APP_MODE / DEMO_CONFIG and
   pace-data.js for how this plugs into the same UI code path production
   uses.
   ───────────────────────────────────────────────────────────────────────── */

const DEMO_STUDENTS = [
  { id: "SIM-001", name: "Alex R.",   active: true },
  { id: "SIM-002", name: "Jordan M.", active: true },
  { id: "SIM-003", name: "Taylor S.", active: true },
  { id: "SIM-004", name: "Casey L.",  active: true },
  { id: "SIM-005", name: "Morgan T.", active: true },
  { id: "SIM-006", name: "Riley C.",  active: true }
];

const DEMO_USER = {
  id: "demo-user",
  name: "Demo Staff",
  role: "PACE Staff"
};

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
  // GRAPH.savePaceVisit() (paceRoom, studentName, date, timeIn, behaviors[],
  // interventions[], scmUsed, notes, submittedByName, timestamp). Stored
  // using the same SharePoint *display*-field-name keys GRAPH uses, so
  // every render function in app.js (which reads v["PACE Room"], v.Student,
  // v["Time In"], etc.) works identically against demo and real data —
  // only this file and pace-data.js know demo mode exists.
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
      "PACE Room":     entry.paceRoom || "",
      "Student":       entry.studentName || "",
      "Date":          entry.date || "",
      "Time In":       entry.timeIn || "",
      "Time Out":      "",
      "Behavior":      Array.isArray(entry.behaviors) ? entry.behaviors.join(", ") : (entry.behaviors || ""),
      "Interventions": Array.isArray(entry.interventions) ? entry.interventions.join(", ") : (entry.interventions || ""),
      "SCM Used":      entry.scmUsed === true,
      "Notes":         entry.notes || "",
      "Submitted By":  entry.submittedByName || DEMO_USER.name,
      "Submitted At":  entry.timestamp || new Date().toISOString(),
      createdAt:       new Date().toISOString()
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
