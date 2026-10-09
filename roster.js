///////////////////////////////////////////////////////////////////////////////////////////////
// Author: R-E Miller & Greg Macer
// Creation Date: August 20, 2026
// Filename: roster.js
// Purpose: Loads and normalizes the SharePoint student roster for the PACE Room Tracker
//          kiosk, mirroring MAC Walkthrough's STUDENT_ROSTER pattern. Owns schema inversion
//          and eligibility normalization — student-select.js handles presentation/search and
//          app.js handles navigation. No student data is hard-coded in source; getPaceEnabled()
//          returns only rows with Active=true and PACE Enabled=true, and no UI code should
//          duplicate or weaken that filter.
///////////////////////////////////////////////////////////////////////////////////////////////

///////////////////////////////////////////////////////////////////////////////////////////////
// Function Name: normalizeRosterBoolean
// Description: Normalizes SharePoint boolean-like roster values with a safe default.
// Parameters: any value - the raw SharePoint field value to interpret - input
//             any defaultValue - value returned when value is undefined/null/empty - input
///////////////////////////////////////////////////////////////////////////////////////////////
function normalizeRosterBoolean(value, defaultValue) {
  if (value === undefined || value === null || value === "") return defaultValue;
  if (typeof value === "boolean") return value;
  if (typeof value === "number") return value === 1;
  const s = String(value).trim().toLowerCase();
  return s === "yes" || s === "true" || s === "1";
}

// SharePoint roster cache and normalization facade. _students holds the normalized roster
// records; loading/loaded track refresh state; error holds the latest refresh error message.
const ROSTER = {
  _students: [],
  loading:   false,
  loaded:    false,
  error:     null,

  ///////////////////////////////////////////////////////////////////////////////////////////////
  // Function Name: _normalize
  // Description: Converts one schema-mapped SharePoint row into the UI student shape.
  // Parameters: Object row - schema-mapped SharePoint roster row - input
  ///////////////////////////////////////////////////////////////////////////////////////////////
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

  ///////////////////////////////////////////////////////////////////////////////////////////////
  // Function Name: refresh
  // Description: Loads, maps, normalizes, and sorts the current student roster.
  // Parameters: none
  ///////////////////////////////////////////////////////////////////////////////////////////////
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

  ///////////////////////////////////////////////////////////////////////////////////////////////
  // Function Name: getPaceEnabled
  // Description: Returns students marked both Active and PACE Enabled.
  // Parameters: none
  ///////////////////////////////////////////////////////////////////////////////////////////////
  getPaceEnabled() { return this._students.filter(s => s.active && s.paceEnabled); },

  ///////////////////////////////////////////////////////////////////////////////////////////////
  // Function Name: find
  // Description: Finds one normalized student by stable roster id.
  // Parameters: string id - the roster id to look up - input
  ///////////////////////////////////////////////////////////////////////////////////////////////
  find(id)         { return this._students.find(s => s.id === id) || null; },

  ///////////////////////////////////////////////////////////////////////////////////////////////
  // Function Name: findByName
  // Description: Finds one normalized student by exact display name.
  // Parameters: string name - the exact display name to look up - input
  ///////////////////////////////////////////////////////////////////////////////////////////////
  findByName(name) { return this._students.find(s => s.name === name) || null; }
};
