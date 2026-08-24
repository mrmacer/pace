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
  },

  // PATCH 006: "Teacher Came From" picklist for the new screen of the same
  // name — deliberately NOT the same list as the old homeroom-grouped
  // Teacher screen (removed this patch). Combines two sources per spec:
  //   Source A — Active "Teacher"-role rows in IEP_Users2 (same
  //              GRAPH.getActiveUsersByRole() the Specialist picker uses).
  //   Source B — unique, non-blank `teacher` values already present on the
  //              PACE-eligible roster (reuses this.getStudents() rather
  //              than a second/different roster fetch — Active + PACE
  //              Enabled only, since that's the only roster query this
  //              project has; a teacher whose only students aren't
  //              PACE-enabled won't appear from this source, but still
  //              will via Source A if they also have a Teacher-role row).
  // Trimmed, case-insensitively deduped (first-seen casing wins), sorted
  // alphabetically. Adding a teacher later needs no code change — see
  // README "Future teacher additions."
  async getTeachers() {
    if (APP_MODE === "demo") return DEMO_CAME_FROM_TEACHERS.slice();

    const [roleTeachers, students] = await Promise.all([
      GRAPH.getActiveUsersByRole("Teacher").catch(err => {
        console.error("Failed to load Teacher-role users from IEP_Users2:", err.message || err);
        return [];
      }),
      this.getStudents().catch(err => {
        console.error("Failed to load roster teacher names:", err.message || err);
        return [];
      })
    ]);

    const seen = new Map(); // lowercase key -> first-seen display casing
    const add = name => {
      const trimmed = String(name || "").trim();
      if (!trimmed) return;
      const key = trimmed.toLowerCase();
      if (!seen.has(key)) seen.set(key, trimmed);
    };
    roleTeachers.forEach(u => add(u.name));
    students.forEach(s => add(s.teacher));

    return [...seen.values()].sort((a, b) => a.localeCompare(b));
  }
};
