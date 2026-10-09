///////////////////////////////////////////////////////////////////////////////////////////////
// Author: R-E Miller & Greg Macer
// Creation Date: August 25, 2026
// Filename: visit-workflow.js
// Purpose: Pure workflow helpers shared by the UI and automated tests. These functions
//          contain no DOM, storage, authentication, or network access. Keeping duration
//          and open-row rules here prevents app.js and test fixtures from silently
//          implementing different visit semantics.
///////////////////////////////////////////////////////////////////////////////////////////////
const PACE_VISIT_WORKFLOW = {
  // Visit lifecycle modes
  MODES: Object.freeze({
    LIVE_START: "live-start",
    COMPLETED_ENTRY: "completed-entry",
    COMPLETING: "completing",
    EDITING: "editing"
  }),

  // NOTES MINIMUM: the one reusable rule for completed-visit Notes. A
  // completed visit (Mark Complete, after-the-fact entry, or an edit of a
  // completed visit) needs at least NOTES_MIN_MEANINGFUL letters/numbers.
  // Whitespace and punctuation never count. This only measures — it never
  // rewrites, trims into, or generates note text; callers save what the
  // user typed.
  NOTES_MIN_MEANINGFUL: 20,
  NOTES_REQUIREMENT_MESSAGE: "Please enter at least 20 letters or numbers describing the visit or support provided.",

  ///////////////////////////////////////////////////////////////////////////////////////////////
  // Function Name: countMeaningfulChars
  // Description: Counts the letters and numbers in a Notes value after trimming; whitespace and
  //              punctuation are ignored.
  // Parameters: any notes - the Notes text to measure - input
  ///////////////////////////////////////////////////////////////////////////////////////////////
  countMeaningfulChars(notes) {
    const matches = String(notes ?? "").trim().match(/[\p{L}\p{N}]/gu);
    return matches ? matches.length : 0;
  },

  ///////////////////////////////////////////////////////////////////////////////////////////////
  // Function Name: validateCompletionNotes
  // Description: The single Notes rule for completed visits: at least NOTES_MIN_MEANINGFUL
  //              letters/numbers. Returns { valid, count, required, message } and never
  //              alters the text.
  // Parameters: any notes - the Notes text to validate - input
  ///////////////////////////////////////////////////////////////////////////////////////////////
  validateCompletionNotes(notes) {
    const count = this.countMeaningfulChars(notes);
    const required = this.NOTES_MIN_MEANINGFUL;
    return {
      valid: count >= required,
      count,
      required,
      message: count >= required ? "" : this.NOTES_REQUIREMENT_MESSAGE
    };
  },

  ///////////////////////////////////////////////////////////////////////////////////////////////
  // Function Name: durationMinutes
  // Description: Calculates a positive same-day duration or returns null when invalid.
  // Parameters: string timeIn - HH:MM visit start time - input
  //             string timeOut - HH:MM visit end time - input
  ///////////////////////////////////////////////////////////////////////////////////////////////
  durationMinutes(timeIn, timeOut) {
    if (!timeIn || !timeOut) return null;
    const [ih, im] = String(timeIn).split(":").map(Number);
    const [oh, om] = String(timeOut).split(":").map(Number);
    if (![ih, im, oh, om].every(Number.isFinite)) return null;
    const duration = (oh * 60 + om) - (ih * 60 + im);
    return duration > 0 ? duration : null;
  },

  ///////////////////////////////////////////////////////////////////////////////////////////////
  // Function Name: isOpen
  // Description: Reports whether a visit has not yet received a Time Out value.
  // Parameters: object visit - visit record to check - input
  ///////////////////////////////////////////////////////////////////////////////////////////////
  isOpen(visit) {
    return Boolean(visit) && !String(visit["Time Out"] || "").trim();
  },

  ///////////////////////////////////////////////////////////////////////////////////////////////
  // Function Name: openForRoom
  // Description: Returns open visits for one room, sorted by entry time.
  // Parameters: array visits - visit records to filter - input
  //             string roomId - PACE Room identifier to match - input
  ///////////////////////////////////////////////////////////////////////////////////////////////
  openForRoom(visits, roomId) {
    return (Array.isArray(visits) ? visits : [])
      .filter(visit => visit?.["PACE Room"] === roomId && this.isOpen(visit))
      .sort((a, b) => String(a["Time In"] || "").localeCompare(String(b["Time In"] || "")));
  }
};

if (typeof module !== "undefined" && module.exports) module.exports = PACE_VISIT_WORKFLOW;
