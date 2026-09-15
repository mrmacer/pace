/* ─────────────────────────────────────────────────────────────────────────
   REQUIRED NOTES PATCH — every completed PACE visit must carry non-empty
   Notes; an open live-start visit never needs them, since it never even
   reaches the Notes screen.

   Static/structural style, matching this project's own established
   convention for app.js-level concerns (see patch010-specialists-notes.
   test.js — app.js is never loaded into node:vm anywhere in this suite;
   its DOM-wiring behavior is proven via source-order/structural assertions
   instead, since the actual validation logic lives inline in event
   handlers rather than in a separately-loadable pure module like
   visit-workflow.js/pace-data.js).

   Run with: node tests/patch-notes-required.test.js
   ───────────────────────────────────────────────────────────────────────── */

const assert = require("node:assert/strict");
const fs = require("node:fs");
const path = require("node:path");

const ROOT = path.resolve(__dirname, "..");
const app = fs.readFileSync(path.join(ROOT, "app.js"), "utf8");
const html = fs.readFileSync(path.join(ROOT, "index.html"), "utf8");
const graph = fs.readFileSync(path.join(ROOT, "graph.js"), "utf8");
const paceData = fs.readFileSync(path.join(ROOT, "pace-data.js"), "utf8");

/* ── UX: Notes marked required in the operational form ──────────────── */

assert.match(html, /for="noteText">Notes\s*<span class="req">\*<\/span><\/label>/, 'Notes field must be visibly marked required ("Notes *"), matching the existing .req convention already used for Visit Date/Time In/Time Out/Teacher Name');
assert.match(html, /<p class="visit-time-error hidden" id="notesError" role="alert"><\/p>/, "an inline error element must exist for Notes, using the same convention as visitTimeError/exitTimeError");
assert.doesNotMatch(html, /id="addNoteBtn"/, 'the old "+ Add Note (Optional)" toggle must be removed — Notes is no longer optional on any screen that reaches it');
assert.doesNotMatch(html, /Optional note/, 'no "optional" placeholder text should remain');

/* ── 1/2/3: trim semantics, no invented minimum length ───────────────── */

// Notes screen's own Next button: trims first, then checks the trimmed
// value — same trim semantics, applied at the point of entry.
assert.match(app, /const trimmed = document\.getElementById\("noteText"\)\.value\.trim\(\);\s*\n\s*if \(!trimmed\)/, "notesNextBtn must block advancing to Confirm on blank/whitespace-only Notes, using trim semantics");
// saveEntryBtn's defensive re-check: same trim semantics, re-verified
// directly against STATE in case anything changed since the Notes screen.
assert.match(app, /if \(!STATE\.notes \|\| !STATE\.notes\.trim\(\)\)/, "saveEntryBtn must defensively re-check Notes with the same trim semantics, mirroring the existing SCM/date-time re-check pattern");
assert.doesNotMatch(app, /STATE\.notes\.trim\(\)\.length\s*[<>]=?\s*\d/, "must not invent a minimum character count — trim-to-non-empty only");
assert.doesNotMatch(app, /notes.{0,40}length\s*[<>]=?\s*[1-9]\d*/i, "no arbitrary minimum length check anywhere near notes validation");

/* ── 4: blocked validation performs NO Graph write ───────────────────── */

{
  const handlerStart = app.indexOf('document.getElementById("saveEntryBtn").addEventListener("click"');
  assert.ok(handlerStart >= 0, "saveEntryBtn click handler must exist");
  const savingIdx = app.indexOf("STATE.saving = true;", handlerStart);
  const notesCheckIdx = app.indexOf("if (!STATE.notes || !STATE.notes.trim())", handlerStart);
  const buildEntryIdx = app.indexOf("const entry = buildVisitEntry();", handlerStart);
  const createIdx = app.indexOf("PACE_DATA.createVisit(entry)", handlerStart);
  const completeIdx = app.indexOf("PACE_DATA.completeVisit(", handlerStart);
  const updateIdx = app.indexOf("PACE_DATA.updateVisit(", handlerStart);

  assert.ok(notesCheckIdx > handlerStart, "notes check must be inside the saveEntryBtn handler");
  assert.ok(notesCheckIdx < savingIdx, "notes check must run BEFORE STATE.saving is set — a rejected save must never even enter the saving state");
  assert.ok(notesCheckIdx < buildEntryIdx, "notes check must run before the entry payload is even built");
  assert.ok(notesCheckIdx < createIdx && notesCheckIdx < completeIdx && notesCheckIdx < updateIdx,
    "notes check must run before every single Graph-write call (create, complete, and update) — blocking one must never leave another reachable");

  // The check's own block must return before reaching any write path.
  const checkBlockEnd = app.indexOf("STATE.saving = true;", notesCheckIdx);
  const checkBlock = app.slice(notesCheckIdx, checkBlockEnd);
  assert.match(checkBlock, /return;/, "the notes-required block must return, not fall through toward a write");
}

/* ── 5: entered form data remains intact after a validation failure ─── */

{
  const scmCheckIdx = app.indexOf("STATE.scmUsed !== true && STATE.scmUsed !== false");
  const notesCheckIdx = app.indexOf("if (!STATE.notes || !STATE.notes.trim())", scmCheckIdx);
  const checkBlockEnd = app.indexOf("STATE.saving = true;", notesCheckIdx);
  const checkBlock = app.slice(notesCheckIdx, checkBlockEnd);
  for (const field of ["STATE.reasons", "STATE.supports", "STATE.scmUsed", "STATE.student", "STATE.staffMembers", "STATE.cameFromTeacher", "STATE.date", "STATE.timeIn", "STATE.timeOut"]) {
    assert.doesNotMatch(checkBlock, new RegExp(`${field}\\s*=[^=]`), `the notes-required block must never reset ${field} — everything already entered must survive the bounce back to Notes`);
  }
  assert.doesNotMatch(checkBlock, /resetTrip\(\)/, "the notes-required block must never call resetTrip() — that would discard the whole in-progress visit, not just prompt for Notes");
}

/* ── 6: starting an open visit can still succeed without Notes ──────── */

{
  const fnStart = app.indexOf("async function saveLiveVisit()");
  const fnEnd = app.indexOf("\n}", fnStart);
  const fnBody = app.slice(fnStart, fnEnd);
  assert.doesNotMatch(fnBody, /STATE\.notes/, "saveLiveVisit() (the Start Visit path) must contain no Notes validation at all — it's architecturally unreachable from the Notes screen");
  assert.match(app, /notes: open \? "" : STATE\.notes/, "buildVisitEntry() must still send no Notes for an open (live-start) visit — unchanged from before this patch");
}

/* ── 7/8/9: completing (Mark Complete), after-the-fact entry, and edit
   all share the SAME unconditional check — none of the three branches is
   exempted ─────────────────────────────────────────────────────────── */

{
  const notesCheckIdx = app.indexOf("if (!STATE.notes || !STATE.notes.trim())", app.indexOf('document.getElementById("saveEntryBtn").addEventListener'));
  const branchIdx = app.indexOf("if (wasCompleting)", notesCheckIdx);
  assert.ok(notesCheckIdx > 0 && branchIdx > notesCheckIdx,
    "the notes check must run BEFORE the branch that decides completeVisit/updateVisit/createVisit — it is unconditional across Mark Complete, editing, and after-the-fact entry alike, not scoped to only one of them");
}

/* ── 10: existing historical records are not modified ───────────────── */

assert.doesNotMatch(app, /for\s*\([^)]*of\s+(recentVisits|allVisits|visits)\)[\s\S]{0,200}updateMappedListItem/, "no bulk/loop update over existing records was introduced");
assert.doesNotMatch(graph, /forEach[\s\S]{0,100}(updateMappedListItem|_patch)/, "graph.js must not gain any bulk-update capability — updates remain one specific itemId at a time");
assert.match(app, /PACE_DATA\.updateVisit\(STATE\.editingVisitId, entry\)/, "editing still targets only the one specific item being actively edited");
assert.match(app, /PACE_DATA\.completeVisit\(STATE\.completionTarget\.id, entry\)/, "completing still targets only the one specific open visit being closed");

/* ── 11: existing dedupe/submission-ID behavior is unchanged ────────── */

assert.match(app, /id: STATE\.submissionId/, "buildVisitEntry() must still carry the same submissionId field this patch did not touch");
assert.match(app, /if \(STATE\.saving\) return; \/\/ never allow a second tap to fire a duplicate submission/, "the existing re-entrancy/dedupe guard on saveEntryBtn must be unchanged");
assert.match(paceData, /async createVisit\(entry\) \{/, "pace-data.js's createVisit() signature/dedupe path is untouched by this patch");

console.log("Required-Notes patch tests passed.");
