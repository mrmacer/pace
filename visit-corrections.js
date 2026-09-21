/* ─────────────────────────────────────────────────────────────────────────
   PACE Room Tracker — staff correction rules (edit + delete)

   Pure logic only: no DOM, Graph, localStorage, or STATE access. app.js and
   pace-data.js supply everything this module needs, which keeps the rules
   testable without a Microsoft session.

   Permission model: there is none in here on purpose. Any staff member who
   already passed the PACE authorization gate may correct or delete a visit;
   app.js supplies that decision through `isAuthorized`. This module never
   looks at a role.

   EDIT — buildEditChanges() returns ONLY the fields that actually differ
   from what the staff member was shown when the edit started, so a
   Notes-only correction never rewrites Time In / Time Out / Duration.
   Duration is recomputed (through the caller-supplied canonical function,
   never a second algorithm here) only when Time In or Time Out changed.

   DELETE — runDelete() targets exactly one SharePoint item id. It never
   deletes without explicit confirmation, re-checks authorization at the
   moment of the destructive write, and never reports success unless the
   data layer confirmed it.
   ───────────────────────────────────────────────────────────────────────── */

const PACE_VISIT_CORRECTIONS = (() => {
  const text = value => String(value ?? "").trim();
  const list = value => (Array.isArray(value) ? value : []).map(text).filter(Boolean);
  const collate = (a, b) => a.localeCompare(b);
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

  function hasChanges(original, edited) {
    const changes = buildEditChanges(original, edited, { durationMinutes: () => null });
    return Object.keys(changes).some(key => key !== "visitDate");
  }

  // A completed visit needs meaningful Notes and both times; a correction
  // may never turn a completed visit into an open one.
  function completedVisitProblem(edited) {
    if (!text(edited?.notes)) return "Add a brief note before submitting this PACE visit.";
    if (!text(edited?.timeIn) || !text(edited?.timeOut)) return "A completed visit needs both Time In and Time Out.";
    return "";
  }

  // What the confirmation shows, so staff can tell exactly which record is
  // about to be removed. `format` supplies the app's own date/time/room
  // formatters; nothing here decides identity — the SharePoint item id does.
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
  function deleteFailureMessage(error) {
    const status = Number(error?.status);
    if (status === 401 || status === 403) return "You don't have permission to delete this visit in SharePoint.";
    if (status === 404) return "This visit no longer exists in SharePoint. Tap Refresh to update the list.";
    return "The visit could not be deleted. Nothing was removed. Check the connection and try again.";
  }

  return { buildEditChanges, hasChanges, completedVisitProblem, deleteSummary, runDelete, deleteFailureMessage };
})();

if (typeof module !== "undefined" && module.exports) module.exports = PACE_VISIT_CORRECTIONS;
