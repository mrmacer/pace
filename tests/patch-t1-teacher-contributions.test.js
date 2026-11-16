/* PATCH T1 — synthetic Teacher Referral & Contribution System coverage
 * Run with: node tests/patch-t1-teacher-contributions.test.js
 */

const assert = require("node:assert/strict");
const fs = require("node:fs");
const path = require("node:path");
const vm = require("node:vm");

const ROOT = path.resolve(__dirname, "..");
const moduleSource = fs.readFileSync(path.join(ROOT, "teacher-contributions.js"), "utf8");

function makeContext(appMode = "demo") {
  const data = new Map();
  let sequence = 0;
  const context = vm.createContext({
    console,
    APP_MODE: appMode,
    crypto: { randomUUID: () => `synthetic-${++sequence}` },
    localStorage: {
      getItem: key => data.has(key) ? data.get(key) : null,
      setItem: (key, value) => data.set(key, value),
      removeItem: key => data.delete(key)
    }
  });
  vm.runInContext(moduleSource, context, { filename: "teacher-contributions.js" });
  return { context, api: vm.runInContext("TEACHER_CONTRIBUTIONS", context) };
}

function visit(overrides = {}) {
  return {
    id: "visit-001",
    studentId: "SIM-001",
    studentName: "Alex R.",
    cameFromTeacher: "Mrs. Ashford",
    date: "2026-11-12",
    timeIn: "09:10",
    behaviors: ["Escalation"],
    notes: "BI notes stay separate.",
    ...overrides
  };
}

const { api } = makeContext();
const assigned = api.createAssignment({ visitId: "visit-001", entry: visit(), visitStatus: "Active" });
assert.equal(assigned.duplicatePrevented, false);
assert.equal(assigned.contribution.status, "Requested");
assert.equal(api.getNotifications().filter(e => e.eventType === "TEACHER_CONTRIBUTION_REQUESTED" && !e.suppressedDuplicate).length, 1);

// Repeating visit-start handling must not create a second assignment or invitation.
const duplicate = api.createAssignment({ visitId: "visit-001", entry: visit(), visitStatus: "Active" });
assert.equal(duplicate.duplicatePrevented, true);
assert.equal(api.getAll().filter(c => c.visitId === "visit-001").length, 1);
assert.equal(api.getNotifications().filter(e => e.suppressedDuplicate).length, 1);

// Teacher visibility is assignment-scoped.
const teacher = api.teacherByName("Mrs. Ashford");
const otherTeacher = api.teacherByName("Mr. Bellamy");
assert.equal(api.getForTeacher(teacher.teacherId).length, 1);
assert.equal(api.getForTeacher(otherTeacher.teacherId).length, 0);
assert.equal(api.isTeacherAuthorized(assigned.contribution.contributionId, teacher.teacherId), true);
assert.equal(api.isTeacherAuthorized(assigned.contribution.contributionId, otherTeacher.teacherId), false);

// Drafts persist without requiring the final fields.
const draft = api.saveDraft(assigned.contribution.contributionId, teacher.teacherId, {
  classroomContext: "Student had difficulty transitioning.",
  strategiesAttempted: "",
  followUp: "Please check in tomorrow."
});
assert.equal(draft.status, "Draft");
assert.equal(api.getById(draft.contributionId).classroomContext, "Student had difficulty transitioning.");
assert.throws(() => api.submit(draft.contributionId, teacher.teacherId, {
  classroomContext: "Student had difficulty transitioning.", strategiesAttempted: "", followUp: ""
}), /classroom context and strategies attempted/);

// Completion generates one reminder while the contribution is incomplete.
const completed = api.markVisitCompleted("visit-001");
assert.equal(completed.visitStatus, "Completed");
assert.equal(completed.status, "Draft");
assert.equal(api.getNotifications().filter(e => e.eventType === "TEACHER_CONTRIBUTION_REMINDER" && !e.suppressedDuplicate).length, 1);

const submitted = api.submit(draft.contributionId, teacher.teacherId, {
  classroomContext: "Student had difficulty transitioning from math.",
  strategiesAttempted: "Visual timer and a quiet prompt.",
  followUp: "Please check in tomorrow."
});
assert.equal(submitted.contribution.status, "Submitted");
assert.equal(submitted.contribution.versions.length, 1);

// Repeated completion does not send another reminder; the suppression is visible in the log.
api.markVisitCompleted("visit-001");
assert.equal(api.getNotifications().filter(e => e.eventType === "TEACHER_CONTRIBUTION_REMINDER" && !e.suppressedDuplicate).length, 1);
assert.ok(api.getNotifications().some(e => e.eventType === "TEACHER_CONTRIBUTION_REMINDER" && e.suppressedDuplicate));

// Corrections create a new version and retain version 1.
api.requestCorrection(draft.contributionId, teacher.teacherId, "Additional classroom context");
const corrected = api.submit(draft.contributionId, teacher.teacherId, {
  classroomContext: "Additional context: the transition followed an unexpected schedule change.",
  strategiesAttempted: "Visual timer, quiet prompt, and a two-minute preview.",
  followUp: "Review the transition plan with the classroom team."
});
assert.equal(corrected.contribution.status, "Submitted");
assert.equal(corrected.contribution.versions.length, 2);
assert.match(corrected.contribution.versions[0].fields.classroomContext, /difficulty transitioning from math/);
assert.match(corrected.contribution.versions[1].correctionReason, /Additional classroom context/);

// Reassignment supersedes the old assignment and creates a fresh invitation.
const replacement = api.reassign(draft.contributionId, "demo-admin", otherTeacher.teacherId);
assert.equal(replacement.original.status, "Superseded");
assert.equal(replacement.replacement.teacherId, otherTeacher.teacherId);
assert.equal(api.isTeacherAuthorized(draft.contributionId, teacher.teacherId), false);
assert.equal(api.isTeacherAuthorized(replacement.replacement.contributionId, otherTeacher.teacherId), true);
assert.equal(api.getAll().filter(c => c.visitId === "visit-001").length, 2);

// Unresolved/manual teachers remain visible to admins but never receive a notification.
const attention = api.createAssignment({
  visitId: "visit-002",
  entry: visit({ id: "visit-002", cameFromTeacher: "Other / Not Listed", studentName: "Jordan M." }),
  visitStatus: "Active"
});
assert.equal(attention.contribution.status, "Needs Attention");
assert.equal(attention.contribution.teacherId, "");
assert.equal(api.getNotifications().some(e => e.contributionId === attention.contribution.contributionId), false);

// The BI note is copied for display context but is never used as teacher note content.
assert.equal(attention.contribution.biNotes, "BI notes stay separate.");
assert.notEqual(attention.contribution.biNotes, attention.contribution.classroomContext);

// Prototype isolation: a production-mode context cannot use the synthetic provider.
const production = makeContext("production").api;
assert.throws(() => production.getAll(), /demo-only/);

// Prototype config must fail closed even on the production hostname or query flag.
const configSource = fs.readFileSync(path.join(ROOT, "config.js"), "utf8");
for (const location of [
  { hostname: "localhost", search: "?production=1" },
  { hostname: "pace-demo-two.vercel.app", search: "?production=1" },
  { hostname: "pace-room-tracker.vercel.app", search: "" }
]) {
  const context = vm.createContext({ console, URLSearchParams, window: { location } });
  vm.runInContext(`${configSource}\nthis.testMode = APP_MODE;`, context, { filename: "config.js" });
  assert.equal(context.testMode, "demo", `prototype must remain demo for ${location.hostname}${location.search}`);
}

const ui = fs.readFileSync(path.join(ROOT, "teacher-contributions-ui.js"), "utf8");
assert.doesNotMatch(moduleSource, /\b(fetch|XMLHttpRequest|Mail\.Send)\b/, "synthetic provider must not contain network or mail calls");
assert.doesNotMatch(ui, /\b(fetch|XMLHttpRequest|Mail\.Send)\b/, "synthetic UI must not contain network or mail calls");

console.log("PATCH T1 teacher contribution tests passed.");
