///////////////////////////////////////////////////////////////////////////////////////////////
// Author: R-E Miller & Greg Macer
// Creation Date: November 23, 2026
// Filename: visit-history.js
// Purpose: Pure rules for looking back at earlier PACE visits — the Recent Activity day picker
//          and a single student's Visit History. No DOM, network, authentication, or storage
//          access belongs here. Past days are read-only: same-day editing and deletion remain
//          governed by RECENT_ACTIVITY.isEditableToday() and PACE_DATA.updateVisit()'s own
//          today-only guard, which this module never relaxes. detailModel() is the one place
//          a visit's full details (including Notes) are shaped for display, and app.js only
//          renders it after a deliberate "Details" tap — cards stay privacy-limited.
///////////////////////////////////////////////////////////////////////////////////////////////

const VISIT_HISTORY = (() => {
  const recent = typeof RECENT_ACTIVITY !== "undefined" ? RECENT_ACTIVITY : require("./recent-activity.js");

  ///////////////////////////////////////////////////////////////////////////////////////////////
  // Function Name: text
  // Description: Normalizes a value to a trimmed string, treating null/undefined as empty.
  // Parameters: any value - the value to coerce and trim - input
  ///////////////////////////////////////////////////////////////////////////////////////////////
  const text = value => String(value ?? "").trim();

  ///////////////////////////////////////////////////////////////////////////////////////////////
  // Function Name: dateOnly
  // Description: Extracts a YYYY-MM-DD date key from a stored date value.
  // Parameters: any value - the stored date value to normalize - input
  ///////////////////////////////////////////////////////////////////////////////////////////////
  const dateOnly = value => text(value).slice(0, 10);

  ///////////////////////////////////////////////////////////////////////////////////////////////
  // Function Name: isIsoDate
  // Description: Reports whether a value is a well-formed, real YYYY-MM-DD calendar date.
  // Parameters: any value - candidate date string - input
  ///////////////////////////////////////////////////////////////////////////////////////////////
  function isIsoDate(value) {
    const raw = text(value);
    if (!/^\d{4}-\d{2}-\d{2}$/.test(raw)) return false;
    const [y, m, d] = raw.split("-").map(Number);
    const check = new Date(Date.UTC(y, m - 1, d));
    return check.getUTCFullYear() === y && check.getUTCMonth() === m - 1 && check.getUTCDate() === d;
  }

  // Calendar arithmetic is done in UTC on purpose: the input is already a
  // LOCAL calendar date, and UTC has no daylight-saving gaps, so adding a
  // day can never skip or repeat a date the way local-time math can.
  ///////////////////////////////////////////////////////////////////////////////////////////////
  // Function Name: shiftDate
  // Description: Returns the YYYY-MM-DD date a whole number of days before/after another.
  // Parameters: string isoDate - starting local calendar date - input
  //             number days - days to move (negative = earlier) - input
  ///////////////////////////////////////////////////////////////////////////////////////////////
  function shiftDate(isoDate, days) {
    if (!isIsoDate(isoDate)) return "";
    const [y, m, d] = isoDate.split("-").map(Number);
    return new Date(Date.UTC(y, m - 1, d + Number(days || 0))).toISOString().slice(0, 10);
  }

  // The day picker may look back, never forward: there is nothing to view
  // in the future and a future date would only show an empty list.
  ///////////////////////////////////////////////////////////////////////////////////////////////
  // Function Name: clampViewDate
  // Description: Returns a valid date to view — the requested one, or today when the request is
  //              missing, malformed, or in the future.
  // Parameters: string requested - date staff asked to view - input
  //             string today - device's local today - input
  ///////////////////////////////////////////////////////////////////////////////////////////////
  function clampViewDate(requested, today) {
    const date = dateOnly(requested);
    if (!isIsoDate(date) || date > dateOnly(today)) return dateOnly(today);
    return date;
  }

  ///////////////////////////////////////////////////////////////////////////////////////////////
  // Function Name: sameName
  // Description: Exact, case-insensitive, whitespace-trimmed name comparison (never partial).
  // Parameters: any a - first name - input
  //             any b - second name - input
  ///////////////////////////////////////////////////////////////////////////////////////////////
  const sameName = (a, b) => Boolean(text(a)) && text(a).toLowerCase() === text(b).toLowerCase();

  // Student identity: Student IDs don't change, so when both the row and the
  // lookup have one, the ID decides (two students can share a name). Rows
  // with no ID (e.g. written by MAC Walkthrough) match on the exact name,
  // case-insensitive. Never a partial match.
  ///////////////////////////////////////////////////////////////////////////////////////////////
  // Function Name: belongsToStudent
  // Description: Reports whether one visit row belongs to the given student.
  // Parameters: object visit - display-name-keyed visit row - input
  //             object student - { id, name } of the student being viewed - input
  ///////////////////////////////////////////////////////////////////////////////////////////////
  function belongsToStudent(visit, student) {
    const rowId = text(visit?.["Student ID"]);
    const wantId = text(student?.id);
    if (rowId && wantId) return rowId === wantId;
    return sameName(visit?.Student, student?.name);
  }

  // The lookup list: every student on the roster, plus any name found on a
  // visit row that isn't on the roster, so every student with visits can be
  // found by name. Name-only entries have no roster id (matched by name).
  ///////////////////////////////////////////////////////////////////////////////////////////////
  // Function Name: lookupStudents
  // Description: Merges roster students with names found on visit rows, one entry per name.
  // Parameters: array roster - roster students - input
  //             array visitNames - every distinct Student value found on visit rows - input
  ///////////////////////////////////////////////////////////////////////////////////////////////
  function lookupStudents(roster, visitNames) {
    const list = (Array.isArray(roster) ? roster : []).slice();
    const seen = new Set(list.map(s => text(s?.name).toLowerCase()).filter(Boolean));
    (Array.isArray(visitNames) ? visitNames : []).forEach(raw => {
      const name = text(raw);
      const key = name.toLowerCase();
      if (!name || seen.has(key)) return;
      seen.add(key);
      list.push({ id: `name:${key}`, name, nameOnly: true });
    });
    return list;
  }

  ///////////////////////////////////////////////////////////////////////////////////////////////
  // Function Name: newestFirst
  // Description: Comparator ordering visits by date, then time in, newest first.
  // Parameters: object a - first visit - input
  //             object b - second visit - input
  ///////////////////////////////////////////////////////////////////////////////////////////////
  function newestFirst(a, b) {
    const key = visit => `${dateOnly(visit?.Date)} ${text(visit?.["Time In"]).slice(0, 5).padStart(5, "0")}`;
    return key(b).localeCompare(key(a));
  }

  ///////////////////////////////////////////////////////////////////////////////////////////////
  // Function Name: forStudent
  // Description: Returns one student's visits (open and completed), newest first.
  // Parameters: array visits - candidate visit rows - input
  //             object student - { id, name } of the student being viewed - input
  ///////////////////////////////////////////////////////////////////////////////////////////////
  function forStudent(visits, student) {
    return (Array.isArray(visits) ? visits : [])
      .filter(visit => belongsToStudent(visit, student))
      .slice()
      .sort(newestFirst);
  }

  ///////////////////////////////////////////////////////////////////////////////////////////////
  // Function Name: groupByDate
  // Description: Groups already-sorted visits into consecutive { date, visits } day buckets.
  // Parameters: array visits - visits sorted newest first - input
  ///////////////////////////////////////////////////////////////////////////////////////////////
  function groupByDate(visits) {
    const groups = [];
    (Array.isArray(visits) ? visits : []).forEach(visit => {
      const date = dateOnly(visit?.Date);
      const last = groups[groups.length - 1];
      if (last && last.date === date) last.visits.push(visit);
      else groups.push({ date, visits: [visit] });
    });
    return groups;
  }

  // Every day, today included, shows ALL of that day's completed visits (the
  // original Recent Activity showed only the 10 newest for today).
  ///////////////////////////////////////////////////////////////////////////////////////////////
  // Function Name: dayLimit
  // Description: Returns how many cards the Recent Activity day view shows: no cap.
  // Parameters: none
  ///////////////////////////////////////////////////////////////////////////////////////////////
  function dayLimit() {
    return Number.MAX_SAFE_INTEGER;
  }

  // The full-details shape, shown only after a deliberate Details tap. Never
  // includes item ids, Entry ID, Student ID, emails, or submission metadata.
  ///////////////////////////////////////////////////////////////////////////////////////////////
  // Function Name: detailModel
  // Description: Shapes one visit's full, read-only details for the Details panel.
  // Parameters: object visit - display-name-keyed visit row - input
  ///////////////////////////////////////////////////////////////////////////////////////////////
  function detailModel(visit) {
    const timeIn = text(visit?.["Time In"]).slice(0, 5);
    const timeOut = text(visit?.["Time Out"]).slice(0, 5);
    const scm = recent.editModel(visit).scmUsed;
    return {
      student: text(visit?.Student) || "Student",
      date: dateOnly(visit?.Date),
      timeIn,
      timeOut,
      open: !timeOut,
      duration: timeOut ? recent.cardModel(visit).duration : null,
      room: text(visit?.["PACE Room"]),
      reasons: recent.selections(visit?.Reason),
      supports: recent.selections(visit?.["Intervention Used"]),
      scm: scm === true ? "Yes" : (scm === false ? "No" : ""),
      specialists: recent.specialistNames(visit?.["Behavior Specialist"] || visit?.["Staff Member"]),
      cameFromTeacher: text(visit?.["Teacher Came From"]),
      notes: String(visit?.Notes ?? "").trim()
    };
  }

  ///////////////////////////////////////////////////////////////////////////////////////////////
  // Function Name: summary
  // Description: Builds the one-line count summary for a student's history.
  // Parameters: array visits - the student's visits, newest first - input
  ///////////////////////////////////////////////////////////////////////////////////////////////
  function summary(visits) {
    const list = Array.isArray(visits) ? visits : [];
    const days = new Set(list.map(visit => dateOnly(visit?.Date)).filter(Boolean)).size;
    return {
      visits: list.length,
      days,
      earliest: list.length ? dateOnly(list[list.length - 1]?.Date) : ""
    };
  }

  return {
    isIsoDate, shiftDate, clampViewDate, belongsToStudent, lookupStudents, forStudent,
    groupByDate, dayLimit, detailModel, summary
  };
})();

if (typeof module !== "undefined" && module.exports) module.exports = VISIT_HISTORY;
