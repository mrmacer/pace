/* ─────────────────────────────────────────────────────────────────────────
   PACE Room Tracker — Recent Activity selection + privacy-safe card models

   This file contains no UI, Graph, or localStorage access. PACE_DATA owns
   the backend choice; app.js receives only the small, today-scoped result
   and converts each selected visit into the explicitly allowed card fields.
   Keeping the rules pure also makes the shared-iPad privacy boundary easy to
   test without a Microsoft session.
   ───────────────────────────────────────────────────────────────────────── */

const RECENT_ACTIVITY = (() => {
  const text = value => String(value ?? "").trim();

  function dateOnly(value) {
    return text(value).slice(0, 10);
  }

  function personDisplayName(value) {
    if (Array.isArray(value)) return value.map(personDisplayName).filter(Boolean).join(", ");
    if (value && typeof value === "object") {
      return text(value.displayName || value.LookupValue || value.lookupValue || value.name || value.title);
    }
    return text(value);
  }

  function sortKey(visit) {
    const explicit = text(visit["Submitted At"] || visit.Created || visit.createdAt);
    const parsed = explicit ? Date.parse(explicit) : NaN;
    if (Number.isFinite(parsed)) return parsed;

    const date = dateOnly(visit.Date);
    const time = text(visit["Time Out"] || visit["Time In"]);
    const fallback = date && time ? Date.parse(`${date}T${time}:00`) : NaN;
    return Number.isFinite(fallback) ? fallback : 0;
  }

  // Room scoping is used only when today's data actually contains room
  // values. The live production schema currently drops "PACE Room" on
  // writes; treating blank room values as a mismatch caused the old false
  // empty state. If room values exist, never mix another room into the list.
  function select(visits, { date, room = "", limit = 10 } = {}) {
    const targetDate = dateOnly(date);
    const targetRoom = text(room);
    const today = (Array.isArray(visits) ? visits : [])
      .filter(visit => dateOnly(visit.Date) === targetDate);

    const rowsWithRoom = today.filter(visit => text(visit["PACE Room"]));
    const roomScoped = Boolean(targetRoom && rowsWithRoom.length > 0);
    const candidates = roomScoped
      ? today.filter(visit => text(visit["PACE Room"]) === targetRoom)
      : today;

    // Recent Activity is a completed-visit view. Legacy/open rows without
    // both times are intentionally excluded, but remain valid elsewhere.
    const completed = candidates.filter(visit =>
      text(visit["Time In"]) && text(visit["Time Out"])
    );

    const safeLimit = Math.max(1, Number(limit) || 10);
    return {
      roomScoped,
      visits: completed
        .map((visit, index) => ({ visit, index, key: sortKey(visit) }))
        .sort((a, b) => b.key - a.key || b.index - a.index)
        .slice(0, safeLimit)
        .map(item => item.visit)
    };
  }

  function durationMinutes(visit) {
    const storedText = text(visit.Duration);
    const stored = Number(storedText);
    if (storedText && Number.isFinite(stored) && stored >= 0) return Math.round(stored);

    const timeIn = text(visit["Time In"]);
    const timeOut = text(visit["Time Out"]);
    if (!/^\d{1,2}:\d{2}$/.test(timeIn) || !/^\d{1,2}:\d{2}$/.test(timeOut)) return null;
    const [inHour, inMinute] = timeIn.split(":").map(Number);
    const [outHour, outMinute] = timeOut.split(":").map(Number);
    let duration = (outHour * 60 + outMinute) - (inHour * 60 + inMinute);
    if (duration < 0) duration += 24 * 60;
    return Number.isFinite(duration) ? duration : null;
  }

  // This object is the privacy allow-list for each card. Notes, IDs,
  // emails, metadata, and raw timestamps never enter the render path.
  function cardModel(visit, { roomScoped = false } = {}) {
    const timeOut = text(visit["Time Out"]);
    return {
      student: text(visit.Student) || "Student",
      displayTime: timeOut || text(visit["Time In"]),
      duration: durationMinutes(visit),
      reason: text(visit.Reason),
      specialist: formatSpecialistsCompact(visit["Behavior Specialist"] || visit["Staff Member"]),
      room: roomScoped ? "" : text(visit["PACE Room"])
    };
  }

  function isEditableToday(visit, date) {
    return Boolean(text(visit?.id)) && dateOnly(visit?.Date) === dateOnly(date);
  }

  function selections(value) {
    if (Array.isArray(value)) return value.map(item => text(item)).filter(Boolean);
    return text(value).split(/[,;]/).map(item => item.trim()).filter(Boolean);
  }

  // PATCH 010: Behavior Specialist(s) share the exact same "comma/
  // semicolon-joined text column, parsed back into an array" convention
  // already used above for Reason/Intervention Used — see
  // GRAPH.savePaceVisit()/updatePaceVisit() (joins STATE.staffMembers with
  // ", ") and app.js's STATE.staffMembers (an array, not a single string,
  // since PATCH 010). personDisplayName() runs first so a defensively-
  // resolved Person-field object/array (never expected in practice — see
  // README "Staff Member is a Person field" — but harmless if it ever
  // happens) still normalizes to plain names before splitting.
  function specialistNames(value) {
    return selections(personDisplayName(value));
  }

  // Compact display for space-constrained cards (Recent Activity, Currently
  // in PACE): full names up to 2, "First Name +N" beyond that. Screens with
  // room to spare (Confirm, Mark Complete) use specialistNames() directly
  // instead, one name per line — see app.js.
  function formatSpecialistsCompact(value) {
    const names = specialistNames(value);
    if (names.length === 0) return "";
    if (names.length <= 2) return names.join(", ");
    return `${names[0]} +${names.length - 1}`;
  }

  function booleanValue(value) {
    if (value === true || value === 1) return true;
    if (value === false || value === 0) return false;
    const normalized = text(value).toLowerCase();
    if (["true", "yes", "1"].includes(normalized)) return true;
    if (["false", "no", "0"].includes(normalized)) return false;
    return null;
  }

  // Full details are exposed only to the deliberate editor entry point;
  // cardModel above remains the shared-iPad display allow-list.
  function editModel(visit, { fallbackRoom = "", validRooms = [] } = {}) {
    const storedRoom = text(visit?.["PACE Room"]);
    return {
      itemId: text(visit?.id),
      submissionId: text(visit?.["Entry ID"] || visit?.id),
      date: dateOnly(visit?.Date),
      timeIn: text(visit?.["Time In"]).slice(0, 5),
      timeOut: text(visit?.["Time Out"]).slice(0, 5),
      reasons: selections(visit?.Reason),
      supports: selections(visit?.["Intervention Used"]),
      scmUsed: booleanValue(visit?.["SCM Used"]),
      notes: String(visit?.Notes ?? ""),
      cameFromTeacher: text(visit?.["Teacher Came From"]),
      // PATCH 010: an array (like `reasons`/`supports` above), not the
      // compact display string cardModel() uses — see app.js's
      // beginEditVisit(), which hydrates STATE.staffMembers from this.
      specialists: specialistNames(visit?.["Behavior Specialist"] || visit?.["Staff Member"]),
      student: text(visit?.Student) || "Student",
      room: validRooms.includes(storedRoom) ? storedRoom : text(fallbackRoom)
    };
  }

  return { select, cardModel, isEditableToday, editModel, selections, specialistNames, formatSpecialistsCompact };
})();

if (typeof module !== "undefined" && module.exports) module.exports = RECENT_ACTIVITY;
