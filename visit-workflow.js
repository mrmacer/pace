/* Pure workflow helpers shared by the UI and automated tests. */
const PACE_VISIT_WORKFLOW = {
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

  countMeaningfulChars(notes) {
    const matches = String(notes ?? "").trim().match(/[\p{L}\p{N}]/gu);
    return matches ? matches.length : 0;
  },

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

  durationMinutes(timeIn, timeOut) {
    if (!timeIn || !timeOut) return null;
    const [ih, im] = String(timeIn).split(":").map(Number);
    const [oh, om] = String(timeOut).split(":").map(Number);
    if (![ih, im, oh, om].every(Number.isFinite)) return null;
    const duration = (oh * 60 + om) - (ih * 60 + im);
    return duration > 0 ? duration : null;
  },

  isOpen(visit) {
    return Boolean(visit) && !String(visit["Time Out"] || "").trim();
  },

  openForRoom(visits, roomId) {
    return (Array.isArray(visits) ? visits : [])
      .filter(visit => visit?.["PACE Room"] === roomId && this.isOpen(visit))
      .sort((a, b) => String(a["Time In"] || "").localeCompare(String(b["Time In"] || "")));
  }
};

if (typeof module !== "undefined" && module.exports) module.exports = PACE_VISIT_WORKFLOW;
