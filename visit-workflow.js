/* Pure workflow helpers shared by the UI and automated tests. */
const PACE_VISIT_WORKFLOW = {
  MODES: Object.freeze({
    LIVE_START: "live-start",
    COMPLETED_ENTRY: "completed-entry",
    COMPLETING: "completing",
    EDITING: "editing"
  }),

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
