///////////////////////////////////////////////////////////////////////////////////////////////
// Author: R-E Miller & Greg Macer
// Creation Date: August 25, 2026
// Filename: recent-activity.js
// Purpose: Pure selection and view-model rules for the PACE Room Tracker Recent Activity
//          panel. No DOM, network, authentication, or storage access belongs here, so room
//          scoping, sorting, privacy filtering, and edit hydration stay testable without a
//          browser or Microsoft session. cardModel() is an explicit privacy allow-list — notes,
//          IDs, email addresses, and raw submission timestamps must not be added casually.
///////////////////////////////////////////////////////////////////////////////////////////////

const RECENT_ACTIVITY = (() => {
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
  function dateOnly(value) {
    return text(value).slice(0, 10);
  }

  ///////////////////////////////////////////////////////////////////////////////////////////////
  // Function Name: personDisplayName
  // Description: Converts SharePoint person shapes or strings into display text.
  // Parameters: any value - a person field value (string, object, or array) - input
  ///////////////////////////////////////////////////////////////////////////////////////////////
  function personDisplayName(value) {
    if (Array.isArray(value)) return value.map(personDisplayName).filter(Boolean).join(", ");
    if (value && typeof value === "object") {
      return text(value.displayName || value.LookupValue || value.lookupValue || value.name || value.title);
    }
    return text(value);
  }

  ///////////////////////////////////////////////////////////////////////////////////////////////
  // Function Name: sortKey
  // Description: Builds a stable newest-first sort key from visit timestamps.
  // Parameters: object visit - the visit record to derive a sort key from - input
  ///////////////////////////////////////////////////////////////////////////////////////////////
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
  // values. Historically the live schema had no matching column at all
  // (any "PACE Room" value came back blank), which is why this still
  // falls back to an unscoped list rather than treating a blank room as a
  // mismatch — see README "Known gaps." ROOM-FIELD-NAME PATCH: the column
  // now exists (confirmed named "Room" — see pace-data.js's
  // normalizeRoomOnRead()/readPaceRoomValue()), so `visit["PACE Room"]`
  // here is already the normalized slug for every current row; this
  // fallback now only matters for older rows saved before the column
  // existed. If room values exist, never mix another room into the list.
  ///////////////////////////////////////////////////////////////////////////////////////////////
  // Function Name: select
  // Description: Selects today's completed visits with optional room scope and limit.
  // Parameters: array visits - the full list of visit records to filter - input
  //             object options - { date, room, limit } selection options - input
  ///////////////////////////////////////////////////////////////////////////////////////////////
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

  ///////////////////////////////////////////////////////////////////////////////////////////////
  // Function Name: durationMinutes
  // Description: Returns stored duration or calculates it from visit times.
  // Parameters: object visit - the visit record to read duration/times from - input
  ///////////////////////////////////////////////////////////////////////////////////////////////
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
  ///////////////////////////////////////////////////////////////////////////////////////////////
  // Function Name: cardModel
  // Description: Produces the privacy-limited fields allowed on a Recent card.
  // Parameters: object visit - the visit record to project into a card - input
  //             object options - { roomScoped } - whether room is already implied - input
  ///////////////////////////////////////////////////////////////////////////////////////////////
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

  ///////////////////////////////////////////////////////////////////////////////////////////////
  // Function Name: isEditableToday
  // Description: Reports whether a visit has an id and belongs to the requested day.
  // Parameters: object visit - the visit record to check - input
  //             string date - the day the visit must fall on - input
  ///////////////////////////////////////////////////////////////////////////////////////////////
  function isEditableToday(visit, date) {
    return Boolean(text(visit?.id)) && dateOnly(visit?.Date) === dateOnly(date);
  }

  ///////////////////////////////////////////////////////////////////////////////////////////////
  // Function Name: selections
  // Description: Parses comma/semicolon-delimited multi-choice text into an array.
  // Parameters: any value - a delimited string or array of choice values - input
  ///////////////////////////////////////////////////////////////////////////////////////////////
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
  ///////////////////////////////////////////////////////////////////////////////////////////////
  // Function Name: specialistNames
  // Description: Normalizes specialist values into individual display names.
  // Parameters: any value - a Person field, string, or array of specialist values - input
  ///////////////////////////////////////////////////////////////////////////////////////////////
  function specialistNames(value) {
    return selections(personDisplayName(value));
  }

  // Compact display for space-constrained cards (Recent Activity, Currently
  // in PACE): full names up to 2, "First Name +N" beyond that. Screens with
  // room to spare (Confirm, Mark Complete) use specialistNames() directly
  // instead, one name per line — see app.js.
  ///////////////////////////////////////////////////////////////////////////////////////////////
  // Function Name: formatSpecialistsCompact
  // Description: Compresses specialist names for narrow room and Recent Activity cards.
  // Parameters: any value - a Person field, string, or array of specialist values - input
  ///////////////////////////////////////////////////////////////////////////////////////////////
  function formatSpecialistsCompact(value) {
    const names = specialistNames(value);
    if (names.length === 0) return "";
    if (names.length <= 2) return names.join(", ");
    return `${names[0]} +${names.length - 1}`;
  }

  ///////////////////////////////////////////////////////////////////////////////////////////////
  // Function Name: booleanValue
  // Description: Converts common SharePoint boolean representations to true/false/null.
  // Parameters: any value - a boolean-ish value (bool, number, or string) - input
  ///////////////////////////////////////////////////////////////////////////////////////////////
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
  ///////////////////////////////////////////////////////////////////////////////////////////////
  // Function Name: editModel
  // Description: Hydrates the full, deliberate same-day editing model from one row.
  // Parameters: object visit - the visit record to hydrate for editing - input
  //             object options - { fallbackRoom, validRooms } - input
  ///////////////////////////////////////////////////////////////////////////////////////////////
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
