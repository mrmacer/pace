/* ─────────────────────────────────────────────────────────────────────────
   PACE Room Tracker — Select Student presentation/filtering rules

   Pure functions only: no DOM, Graph, localStorage, or STATE access. app.js
   maps renderModel()'s output onto the Select Student screen. Eligibility
   (Active + PACE Enabled, from IEP_Students_2026_27) is decided upstream by
   ROSTER.getPaceEnabled(); this module never widens or narrows that list —
   it only orders, filters, and labels the students it is given.

   Nothing here changes what is written to SharePoint: the student's `name`
   (used for the visit's Student value and duplicate-open checks) is never
   modified; only the card's visible label is formatted.
   ───────────────────────────────────────────────────────────────────────── */

const PACE_STUDENT_SELECT = {
  NO_MATCH_MESSAGE: "No matching PACE-enabled student found.",
  NO_MATCH_HINT: "Check the student's Active and PACE Enabled settings if you expected them to appear.",
  EMPTY_ROSTER_MESSAGE: "No PACE-enabled students are available.",

  // Lowercase, strip punctuation ("Ja'de" still matches "jade"), collapse to
  // plain text. Same rule Teacher Came From search uses in app.js.
  normalize(str) {
    return String(str || "").toLowerCase().replace(/[^a-z0-9\s]/g, "").trim();
  },

  // Matches first name, last name, "First Last", AND "Last First".
  searchCorpus(student) {
    const parts = [student.name, student.firstName, student.lastName];
    if (student.firstName || student.lastName) {
      parts.push(`${student.lastName || ""} ${student.firstName || ""}`);
    }
    return this.normalize(parts.filter(Boolean).join(" "));
  },

  // Last name, then first name; falls back to the combined display name for
  // rows without separate first/last fields. Deterministic tie-breaks.
  sortKey(student) {
    return student.firstName || student.lastName
      ? `${student.lastName || ""} ${student.firstName || ""}`.trim()
      : String(student.name || "");
  },

  sort(students) {
    const collate = (a, b) => a.localeCompare(b, undefined, { sensitivity: "base" });
    return (Array.isArray(students) ? students : []).slice().sort((a, b) =>
      collate(this.sortKey(a), this.sortKey(b)) ||
      collate(String(a.name || ""), String(b.name || "")) ||
      collate(String(a.id || ""), String(b.id || ""))
    );
  },

  // Visible card label only — never used as the student's identity/value.
  displayLabel(student) {
    return student.firstName && student.lastName
      ? `${student.lastName}, ${student.firstName}`
      : String(student.name || "");
  },

  // Search always runs against the COMPLETE eligible roster it is handed.
  filter(students, query) {
    const q = this.normalize(query);
    const sorted = this.sort(students);
    return q ? sorted.filter(student => this.searchCorpus(student).includes(q)) : sorted;
  },

  countText(total, shown, isFiltering) {
    const noun = total === 1 ? "student" : "students";
    return isFiltering ? `${shown} of ${total} ${noun}` : `${total} ${noun} available`;
  },

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
