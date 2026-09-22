///////////////////////////////////////////////////////////////////////////////////////////////
// Author: R-E Miller & Greg Macer
// Creation Date: September 21, 2026
// Filename: visit-corrections.js
// Purpose: Pure business rules for same-day editing and deletion of PACE visits;
//          app.js supplies authorization and the data adapter supplies
//          persistence callbacks. Corrections are differential (only changed
//          fields are written), completed visits must retain both times and
//          notes, and deletion requires explicit confirmation plus a fresh
//          authorization check immediately before persistence.
///////////////////////////////////////////////////////////////////////////////////////////////
// PACE Room Tracker — staff correction rules (edit + delete)
//
// Pure logic only: no DOM, Graph, localStorage, or STATE access. app.js and
// pace-data.js supply everything this module needs, which keeps the rules
// testable without a Microsoft session.
//
// Permission model: there is none in here on purpose. Any staff member who
// already passed the PACE authorization gate may correct or delete a visit;
// app.js supplies that decision through `isAuthorized`. This module never
// looks at a role.
//
// EDIT — buildEditChanges() returns ONLY the fields that actually differ
// from what the staff member was shown when the edit started, so a
// Notes-only correction never rewrites Time In / Time Out / Duration.
// Duration is recomputed (through the caller-supplied canonical function,
// never a second algorithm here) only when Time In or Time Out changed.
//
// DELETE — runDelete() targets exactly one SharePoint item id. It never
// deletes without explicit confirmation, re-checks authorization at the
// moment of the destructive write, and never reports success unless the
// data layer confirmed it.

const PACE_VISIT_CORRECTIONS = (() => {
  ///////////////////////////////////////////////////////////////////////////////////////////////
  // Function Name: text
  // Description: Coerces a value to a trimmed string, treating null/undefined as an empty string.
  // Parameters: any value - the value to normalize - input
  ///////////////////////////////////////////////////////////////////////////////////////////////
  const text = value => String(value ?? "").trim();

  ///////////////////////////////////////////////////////////////////////////////////////////////
  // Function Name: list
  // Description: Normalizes an array-like input into an array of trimmed, non-empty strings.
  // Parameters: any value - the value to normalize into a list - input
  ///////////////////////////////////////////////////////////////////////////////////////////////
  const list = value => (Array.isArray(value) ? value : []).map(text).filter(Boolean);

  ///////////////////////////////////////////////////////////////////////////////////////////////
  // Function Name: collate
  // Description: Compares two strings using locale-aware ordering, for use as an Array.sort
  //              comparator.
  // Parameters: string a - the first string to compare - input
  //             string b - the second string to compare - input
  ///////////////////////////////////////////////////////////////////////////////////////////////
  const collate = (a, b) => a.localeCompare(b);

  ///////////////////////////////////////////////////////////////////////////////////////////////
  // Function Name: sameSet
  // Description: Reports whether two array-like inputs contain the same set of trimmed string
  //              values, ignoring order.
  // Parameters: any a - the first collection to compare - input
  //             any b - the second collection to compare - input
  ///////////////////////////////////////////////////////////////////////////////////////////////
  const sameSet = (a, b) => {
    const x = list(a).sort(collate);
    const y = list(b).sort(collate);
    return x.length === y.length && x.every((item, index) => item === y[index]);
  };

  // `original` is the snapshot taken when the edit began; `edited` is the
  // full logical entry app.js builds at save time (same shape either way):
  //   { paceRoom, studentName, date, timeIn, timeOut, behaviors[],
  //     interventions[], scmUsed, notes, staffMembers[], cameFromTeacher }
  // The result contains only changed fields, plus `visitDate` — a guard
  // input for the same-day rule that is never written to SharePoint.
  ///////////////////////////////////////////////////////////////////////////////////////////////
  // Function Name: buildEditChanges
  // Description: Computes only the visit fields that differ between the original and edited
  //              entry, recomputing duration when Time In or Time Out changed.
  // Parameters: object original - the visit snapshot taken when the edit began - input
  //             object edited - the full edited visit entry built by app.js at save time - input
  //             object options - object providing the canonical durationMinutes(timeIn, timeOut)
  //                               function, used only when times change - input
  ///////////////////////////////////////////////////////////////////////////////////////////////
  function buildEditChanges(original, edited, { durationMinutes } = {}) {
    if (!original || !edited) throw new Error("An original visit and an edited visit are required.");
    const changes = {};

    if (text(edited.paceRoom) !== text(original.paceRoom)) changes.paceRoom = text(edited.paceRoom);
    if (text(edited.studentName) !== text(original.studentName)) changes.studentName = text(edited.studentName);
    if (text(edited.date) !== text(original.date)) changes.date = text(edited.date);

    let timesChanged = false;
    if (text(edited.timeIn) !== text(original.timeIn)) { changes.timeIn = text(edited.timeIn); timesChanged = true; }
    if (text(edited.timeOut) !== text(original.timeOut)) { changes.timeOut = text(edited.timeOut); timesChanged = true; }
    if (timesChanged) {
      if (typeof durationMinutes !== "function") throw new Error("The canonical duration function is required when times change.");
      changes.durationMinutes = durationMinutes(text(edited.timeIn), text(edited.timeOut));
    }

    if (!sameSet(edited.behaviors, original.behaviors)) changes.behaviors = list(edited.behaviors);
    if (!sameSet(edited.interventions, original.interventions)) changes.interventions = list(edited.interventions);
    if (edited.scmUsed !== original.scmUsed) changes.scmUsed = edited.scmUsed;
    if (text(edited.notes) !== text(original.notes)) changes.notes = text(edited.notes);
    if (!sameSet(edited.staffMembers, original.staffMembers)) changes.staffMembers = list(edited.staffMembers);
    if (text(edited.cameFromTeacher) !== text(original.cameFromTeacher)) changes.cameFromTeacher = text(edited.cameFromTeacher);

    changes.visitDate = text(edited.date);
    return changes;
  }

  ///////////////////////////////////////////////////////////////////////////////////////////////
  // Function Name: hasChanges
  // Description: Reports whether an edit changes at least one persisted visit field, ignoring
  //              the visitDate guard field.
  // Parameters: object original - the visit snapshot taken when the edit began - input
  //             object edited - the full edited visit entry built by app.js at save time - input
  ///////////////////////////////////////////////////////////////////////////////////////////////
  function hasChanges(original, edited) {
    const changes = buildEditChanges(original, edited, { durationMinutes: () => null });
    return Object.keys(changes).some(key => key !== "visitDate");
  }

  // A completed visit needs meaningful Notes and both times; a correction
  // may never turn a completed visit into an open one.
  ///////////////////////////////////////////////////////////////////////////////////////////////
  // Function Name: completedVisitProblem
  // Description: Validates that a completed visit correction retains meaningful notes and both
  //              times, returning a user-facing error when it does not.
  // Parameters: object edited - the edited visit entry to validate - input
  ///////////////////////////////////////////////////////////////////////////////////////////////
  function completedVisitProblem(edited) {
    if (!text(edited?.notes)) return "Add a brief note before submitting this PACE visit.";
    if (!text(edited?.timeIn) || !text(edited?.timeOut)) return "A completed visit needs both Time In and Time Out.";
    return "";
  }

  // What the confirmation shows, so staff can tell exactly which record is
  // about to be removed. `format` supplies the app's own date/time/room
  // formatters; nothing here decides identity — the SharePoint item id does.
  ///////////////////////////////////////////////////////////////////////////////////////////////
  // Function Name: deleteSummary
  // Description: Builds the confirmation summary shown to staff before a destructive visit
  //              deletion.
  // Parameters: object visit - the visit record to summarize - input
  //             object format - optional app-supplied longDate/time/roomLabel formatters - input
  ///////////////////////////////////////////////////////////////////////////////////////////////
  function deleteSummary(visit, format = {}) {
    const fmtDate = format.longDate || text;
    const fmtTime = format.time || text;
    const roomLabel = format.roomLabel || text;
    const timeIn = text(visit?.["Time In"]);
    const timeOut = text(visit?.["Time Out"]);
    return {
      student: text(visit?.Student) || "Student",
      date: fmtDate(text(visit?.Date).slice(0, 10)),
      time: timeIn && timeOut ? `${fmtTime(timeIn)} – ${fmtTime(timeOut)}` : fmtTime(timeIn || timeOut),
      room: roomLabel(text(visit?.["PACE Room"]))
    };
  }

  // Delete orchestration. Every guard runs before deleteVisit(), which is
  // called at most once with the item id and nothing else.
  //   confirmed    — true only from the confirmation dialog's Delete button
  //   isAuthorized — the app's canonical PACE authorization check, evaluated
  //                  now, at the moment of the destructive write
  //   deleteVisit  — (itemId) => Promise; resolves only on confirmed success
  //   onSuccess    — called after (and only after) deleteVisit resolved
  ///////////////////////////////////////////////////////////////////////////////////////////////
  // Function Name: runDelete
  // Description: Applies confirmation and authorization guards, then deletes exactly one visit
  //              and reports the outcome.
  // Parameters: object visit - the visit record targeted for deletion - input
  //             boolean confirmed - true only when the confirmation dialog's Delete button was
  //                                  pressed - input
  //             function isAuthorized - the app's canonical PACE authorization check, called at
  //                                      the moment of the destructive write - input
  //             function deleteVisit - deletes the visit by item id, resolving only on confirmed
  //                                     success - input
  //             function onSuccess - optional callback invoked after a successful delete - input
  ///////////////////////////////////////////////////////////////////////////////////////////////
  async function runDelete({ visit, confirmed, isAuthorized, deleteVisit, onSuccess }) {
    if (confirmed !== true) return { status: "cancelled" };
    if (typeof isAuthorized !== "function" || !isAuthorized()) return { status: "unauthorized" };
    const itemId = text(visit?.id);
    if (!itemId) return { status: "invalid" };
    try {
      await deleteVisit(itemId);
    } catch (error) {
      return { status: "failed", error };
    }
    if (typeof onSuccess === "function") await onSuccess();
    return { status: "deleted", itemId };
  }

  // Short, safe wording for the failure toast — never a raw Graph body.
  ///////////////////////////////////////////////////////////////////////////////////////////////
  // Function Name: deleteFailureMessage
  // Description: Maps a deletion failure to a safe, non-sensitive user-facing error message.
  // Parameters: object error - the error thrown by the delete operation - input
  ///////////////////////////////////////////////////////////////////////////////////////////////
  function deleteFailureMessage(error) {
    const status = Number(error?.status);
    if (status === 401 || status === 403) return "You don't have permission to delete this visit in SharePoint.";
    if (status === 404) return "This visit no longer exists in SharePoint. Tap Refresh to update the list.";
    return "The visit could not be deleted. Nothing was removed. Check the connection and try again.";
  }

  return { buildEditChanges, hasChanges, completedVisitProblem, deleteSummary, runDelete, deleteFailureMessage };
})();

if (typeof module !== "undefined" && module.exports) module.exports = PACE_VISIT_CORRECTIONS;
