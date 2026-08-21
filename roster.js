/* ─────────────────────────────────────────────────────────────────────────
   PACE Room Tracker — Student roster provider

   Mirrors MAC Walkthrough's STUDENT_ROSTER (app.js): loads
   IEP_Students_2026_27 live from SharePoint, normalizes it, and exposes
   only PACE-enabled + Active students. No student data is ever hard-coded
   in source — this module only knows HOW to load the roster.
   ───────────────────────────────────────────────────────────────────────── */

function normalizeRosterBoolean(value, defaultValue) {
  if (value === undefined || value === null || value === "") return defaultValue;
  if (typeof value === "boolean") return value;
  if (typeof value === "number") return value === 1;
  const s = String(value).trim().toLowerCase();
  return s === "yes" || s === "true" || s === "1";
}

const ROSTER = {
  _students: [],
  loading:   false,
  loaded:    false,
  error:     null,

  _normalize(row) {
    const firstName = String(row["Student First Name"] || "").trim();
    const lastName  = String(row["Student Last Name"]  || "").trim();
    const legacyName = String(row["Student Name"] || row["StudentName"] || row["Name"] || "").trim();
    const name = (firstName || lastName) ? `${firstName} ${lastName}`.trim() : legacyName;
    if (!name) return null;
    return {
      id:          row.id != null ? String(row.id) : name,
      firstName, lastName, name,
      // PATCH 004: previously read nowhere in this file even though both
      // columns exist on the live roster list — needed now to group the
      // Student screen by Teacher. Classroom carried along too (spec:
      // "may remain available for future filtering or display").
      teacher:     String(row["Teacher"]   || "").trim(),
      classroom:   String(row["Classroom"] || "").trim(),
      active:      normalizeRosterBoolean(row["Active"],       true),
      paceEnabled: normalizeRosterBoolean(row["PACE Enabled"], false)
    };
  },

  async refresh() {
    this.loading = true;
    this.error   = null;
    try {
      const listName = CONFIG.LISTS.students;
      const [raw, schema] = await Promise.all([
        GRAPH.getListItems(listName).catch(async (err) => {
          // Defensive fallback, mirroring MAC Walkthrough's behavior if the
          // production list name is ever unset/renamed.
          console.warn("Primary roster list failed, trying fallback:", err.message);
          return GRAPH.getListItems(CONFIG.STUDENT_ROSTER_FALLBACK);
        }),
        GRAPH.getListSchema(listName).catch(() => GRAPH.getListSchema(CONFIG.STUDENT_ROSTER_FALLBACK))
      ]);
      const inv = {};
      Object.entries(schema).forEach(([display, internal]) => { inv[internal] = display; });
      this._students = raw
        .map(item => {
          const row = { id: item.id };
          Object.entries(item).forEach(([k, v]) => { if (k !== "id") row[inv[k] || k] = v; });
          return this._normalize(row);
        })
        .filter(Boolean)
        .sort((a, b) => a.name.localeCompare(b.name));
      this.loaded = true;
    } catch (err) {
      console.error("Student roster load failed:", err.message || String(err));
      this.error = err.message || "Unable to load the student roster.";
    } finally {
      this.loading = false;
    }
    return this._students;
  },

  getPaceEnabled() { return this._students.filter(s => s.active && s.paceEnabled); },
  find(id)         { return this._students.find(s => s.id === id) || null; },
  findByName(name) { return this._students.find(s => s.name === name) || null; }
};
