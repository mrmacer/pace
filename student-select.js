///////////////////////////////////////////////////////////////////////////////////////////////
// Author: R-E Miller & Greg Macer
// Creation Date: September 21, 2026
// Filename: student-select.js
// Purpose: Pure search, sorting, labeling, and view-model logic for the Select Student panel
//          of the PACE Room Tracker kiosk app. Receives an already-eligible roster from
//          ROSTER.getPaceEnabled() and never changes student identity, eligibility, or
//          persistence values — it only orders, filters, and labels the students it is given.
//          displayLabel() is presentation only: the stable student id remains the value app.js
//          uses for selection and duplicate-open detection, and the student's `name` written to
//          SharePoint is never modified.
///////////////////////////////////////////////////////////////////////////////////////////////

// Pure Select Student presentation API — search, sort, and label the eligible roster.
const PACE_STUDENT_SELECT = {
  // Empty filtered-result message.
  NO_MATCH_MESSAGE: "No matching PACE-enabled student found.",
  // Guidance shown for an unexpected absence.
  NO_MATCH_HINT: "Check the student's Active and PACE Enabled settings if you expected them to appear.",
  // Message for an empty eligible roster.
  EMPTY_ROSTER_MESSAGE: "No PACE-enabled students are available.",

  // Lowercase, strip punctuation ("Ja'de" still matches "jade"), collapse to
  // plain text. Same rule Teacher Came From search uses in app.js.
  ///////////////////////////////////////////////////////////////////////////////////////////////
  // Function Name: normalize
  // Description: Normalizes punctuation and case for student search matching.
  // Parameters: string str - raw text to normalize for matching - input
  ///////////////////////////////////////////////////////////////////////////////////////////////
  normalize(str) {
    return String(str || "").toLowerCase().replace(/[^a-z0-9\s]/g, "").trim();
  },

  // Matches first name, last name, "First Last", AND "Last First".
  ///////////////////////////////////////////////////////////////////////////////////////////////
  // Function Name: searchCorpus
  // Description: Builds the searchable first-name/last-name corpus for one student.
  // Parameters: object student - student record to build a search corpus from - input
  ///////////////////////////////////////////////////////////////////////////////////////////////
  searchCorpus(student) {
    const parts = [student.name, student.firstName, student.lastName];
    if (student.firstName || student.lastName) {
      parts.push(`${student.lastName || ""} ${student.firstName || ""}`);
    }
    return this.normalize(parts.filter(Boolean).join(" "));
  },

  // Last name, then first name; falls back to the combined display name for
  // rows without separate first/last fields. Deterministic tie-breaks.
  ///////////////////////////////////////////////////////////////////////////////////////////////
  // Function Name: sortKey
  // Description: Builds the deterministic last-name-first sort key.
  // Parameters: object student - student record to build a sort key from - input
  ///////////////////////////////////////////////////////////////////////////////////////////////
  sortKey(student) {
    return student.firstName || student.lastName
      ? `${student.lastName || ""} ${student.firstName || ""}`.trim()
      : String(student.name || "");
  },

  ///////////////////////////////////////////////////////////////////////////////////////////////
  // Function Name: sort
  // Description: Returns a non-mutating alphabetical copy of the supplied roster.
  // Parameters: array students - roster to sort alphabetically - input
  ///////////////////////////////////////////////////////////////////////////////////////////////
  sort(students) {
    const collate = (a, b) => a.localeCompare(b, undefined, { sensitivity: "base" });
    return (Array.isArray(students) ? students : []).slice().sort((a, b) =>
      collate(this.sortKey(a), this.sortKey(b)) ||
      collate(String(a.name || ""), String(b.name || "")) ||
      collate(String(a.id || ""), String(b.id || ""))
    );
  },

  // Visible card label only — never used as the student's identity/value.
  ///////////////////////////////////////////////////////////////////////////////////////////////
  // Function Name: displayLabel
  // Description: Formats the visible student card label without changing identity data.
  // Parameters: object student - student record to format a label for - input
  ///////////////////////////////////////////////////////////////////////////////////////////////
  displayLabel(student) {
    return student.firstName && student.lastName
      ? `${student.lastName}, ${student.firstName}`
      : String(student.name || "");
  },

  // Search always runs against the COMPLETE eligible roster it is handed.
  ///////////////////////////////////////////////////////////////////////////////////////////////
  // Function Name: filter
  // Description: Filters the complete eligible roster using the normalized search query.
  // Parameters: array students - eligible roster to filter - input
  //             string query - search text to match against - input
  ///////////////////////////////////////////////////////////////////////////////////////////////
  filter(students, query) {
    const q = this.normalize(query);
    const sorted = this.sort(students);
    return q ? sorted.filter(student => this.searchCorpus(student).includes(q)) : sorted;
  },

  ///////////////////////////////////////////////////////////////////////////////////////////////
  // Function Name: countText
  // Description: Formats the roster count shown beside the search field.
  // Parameters: number total - total eligible student count - input
  //             number shown - count of students currently visible - input
  //             boolean isFiltering - whether a search query is active - input
  ///////////////////////////////////////////////////////////////////////////////////////////////
  countText(total, shown, isFiltering) {
    const noun = total === 1 ? "student" : "students";
    return isFiltering ? `${shown} of ${total} ${noun}` : `${total} ${noun} available`;
  },

  ///////////////////////////////////////////////////////////////////////////////////////////////
  // Function Name: renderModel
  // Description: Builds the pure view model consumed by app.js's student grid.
  // Parameters: array students - eligible roster to build a view model from - input
  //             string query - search text, defaults to empty - input
  ///////////////////////////////////////////////////////////////////////////////////////////////
  renderModel(students, query = "") {
    const total = Array.isArray(students) ? students.length : 0;
    const isFiltering = Boolean(this.normalize(query));
    const visible = this.filter(students, query);
    const emptyMessage = total === 0
      ? this.EMPTY_ROSTER_MESSAGE
      : (visible.length === 0 ? this.NO_MATCH_MESSAGE : "");
    return {
      total,
      shown: visible.length,
      isFiltering,
      students: visible,
      countText: total === 0 ? "" : this.countText(total, visible.length, isFiltering),
      emptyMessage,
      emptyHint: emptyMessage ? this.NO_MATCH_HINT : ""
    };
  }
};

if (typeof module !== "undefined" && module.exports) module.exports = PACE_STUDENT_SELECT;
