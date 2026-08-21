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
  }
};
