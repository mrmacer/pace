/* ─────────────────────────────────────────────────────────────────────────
   PACE Room Tracker — Unified data adapter

   The single interface app.js talks to. Every read/write branches on
   APP_MODE here, in exactly one place — app.js never checks APP_MODE or
   calls GRAPH/DemoStorage directly, so the two modes render and behave
   identically (spec: "The UX should be identical to production mode. Only
   the backend changes.").
   ───────────────────────────────────────────────────────────────────────── */

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

  // Every PACE visit row, display-field-name keyed (see demo-data.js for
  // why the shapes match).
  async getVisits() {
    if (APP_MODE === "demo") return DemoStorage.getVisits();
    return GRAPH.getPaceVisitsByDisplayName();
  },

  async createVisit(entry) {
    if (APP_MODE === "demo") return DemoStorage.createVisit(entry);
    return GRAPH.savePaceVisit(entry);
  },

  async closeVisit(id, timeOut) {
    if (APP_MODE === "demo") return DemoStorage.updateVisit(id, { "Time Out": timeOut });
    return GRAPH.closePaceVisit(id, timeOut);
  }
};
