///////////////////////////////////////////////////////////////////////////////////////////////
// PATCH T1 — synthetic Teacher Referral & Contribution System
// Purpose: Demo-only contribution data model, local persistence, authorization simulation,
//          versioned corrections, assignment/reassignment, and notification event logging.
//          This module deliberately has no Graph, SharePoint, Outlook, SMTP, or Power Automate
//          integration. It is the seam a future Microsoft 365 provider can replace.
///////////////////////////////////////////////////////////////////////////////////////////////

(function (root, factory) {
  const api = factory();
  if (typeof module !== "undefined" && module.exports) module.exports = api;
  root.TEACHER_CONTRIBUTIONS = api;
})(typeof globalThis !== "undefined" ? globalThis : window, function () {
  const CONFIG = {
    storageKey: "paceRoomTrackerTeacherContributionsDemoData",
    statuses: ["Requested", "Draft", "Submitted", "Correction requested", "Superseded", "Needs Attention"],
    deliveryStatus: "Simulated — not sent"
  };

  // Every identity here is fictional and intentionally uses the reserved
  // .invalid domain. These records never come from IEP_Users2 in demo mode.
  const DEMO_TEACHERS = [
    { teacherId: "demo-teacher-ashford", displayName: "Mrs. Ashford", fictionalEmail: "mrs.ashford@demo.invalid", active: true, role: "Teacher" },
    { teacherId: "demo-teacher-bellamy", displayName: "Mr. Bellamy", fictionalEmail: "mr.bellamy@demo.invalid", active: true, role: "Teacher" },
    { teacherId: "demo-teacher-castillo", displayName: "Ms. Castillo", fictionalEmail: "ms.castillo@demo.invalid", active: true, role: "Teacher" },
    { teacherId: "demo-teacher-delgado", displayName: "Mr. Delgado", fictionalEmail: "mr.delgado@demo.invalid", active: true, role: "Teacher" },
    { teacherId: "demo-teacher-rivera", displayName: "Jordan Rivera", fictionalEmail: "jordan.rivera@demo.invalid", active: true, role: "Teacher" }
  ];

  const REQUIRED_FIELDS = ["classroomContext", "strategiesAttempted"];
  const NOTE_FIELDS = ["classroomContext", "strategiesAttempted", "followUp"];

  function now() {
    return new Date().toISOString();
  }

  function makeId(prefix) {
    const uuid = typeof crypto !== "undefined" && typeof crypto.randomUUID === "function"
      ? crypto.randomUUID()
      : `${Date.now()}-${Math.random().toString(16).slice(2)}`;
    return `${prefix}-${uuid}`;
  }

  function clone(value) {
    return JSON.parse(JSON.stringify(value));
  }

  function text(value) {
    return String(value ?? "").trim();
  }

  function assertDemoOnly() {
    if (typeof APP_MODE !== "undefined" && APP_MODE !== "demo") {
      throw new Error("Teacher contribution prototype is demo-only; live provider access is disabled.");
    }
  }

  function emptyStore() {
    return { contributions: [], notifications: [] };
  }

  const Storage = {
    load() {
      assertDemoOnly();
      try {
        const parsed = JSON.parse(localStorage.getItem(CONFIG.storageKey));
        return parsed && Array.isArray(parsed.contributions) && Array.isArray(parsed.notifications)
          ? parsed
          : emptyStore();
      } catch {
        return emptyStore();
      }
    },

    save(data) {
      assertDemoOnly();
      localStorage.setItem(CONFIG.storageKey, JSON.stringify(data));
    },

    reset() {
      assertDemoOnly();
      localStorage.removeItem(CONFIG.storageKey);
    }
  };

  function teacherById(teacherId) {
    return DEMO_TEACHERS.find(t => t.teacherId === text(teacherId)) || null;
  }

  function teacherByName(displayName) {
    const target = text(displayName).toLowerCase();
    return DEMO_TEACHERS.find(t => t.displayName.toLowerCase() === target) || null;
  }

  function fieldsFrom(contribution) {
    return NOTE_FIELDS.reduce((out, key) => {
      out[key] = contribution?.[key] || "";
      return out;
    }, {});
  }

  function requiredProblems(fields) {
    return REQUIRED_FIELDS.filter(key => !text(fields?.[key]));
  }

  function makeDueAt(visitDate) {
    const base = visitDate && /^\d{4}-\d{2}-\d{2}$/.test(String(visitDate))
      ? new Date(`${visitDate}T16:00:00`)
      : new Date(Date.now() + 24 * 60 * 60 * 1000);
    base.setDate(base.getDate() + 1);
    return base.toISOString();
  }

  function assignmentKey(contribution) {
    return `${text(contribution.visitId)}|${text(contribution.teacherId)}|${contribution.assignmentVersion || 1}`;
  }

  function notificationKey(eventType, contribution) {
    return `${eventType}|${assignmentKey(contribution)}`;
  }

  function findContribution(data, contributionId) {
    return data.contributions.find(c => c.contributionId === text(contributionId)) || null;
  }

  function findByVisit(data, visitId) {
    return data.contributions.find(c => c.visitId === text(visitId) && c.status !== "Superseded") || null;
  }

  function eventFor(data, eventType, contribution) {
    const key = notificationKey(eventType, contribution);
    return data.notifications.find(event => event.deduplicationKey === key && !event.suppressedDuplicate) || null;
  }

  function appendNotification(data, contribution, eventType, reason = "") {
    const deduplicationKey = notificationKey(eventType, contribution);
    const prior = eventFor(data, eventType, contribution);
    const suppressed = Boolean(prior);
    const event = {
      notificationId: makeId("notification"),
      contributionId: contribution.contributionId,
      visitId: contribution.visitId,
      teacherId: contribution.teacherId || "",
      recipient: contribution.teacherEmail || "Needs attention",
      eventType,
      createdAt: now(),
      simulatedDeliveryStatus: suppressed ? `Suppressed — ${reason || "duplicate"}` : CONFIG.deliveryStatus,
      deduplicationKey,
      suppressedDuplicate: suppressed,
      suppressionReason: suppressed ? (reason || "Duplicate event") : ""
    };
    data.notifications.push(event);
    return event;
  }

  function contributionSnapshot(contribution) {
    return {
      version: contribution.currentVersion || 1,
      classroomContext: contribution.classroomContext || "",
      strategiesAttempted: contribution.strategiesAttempted || "",
      followUp: contribution.followUp || "",
      submittedBy: contribution.submittedBy || "",
      submittedByEmail: contribution.submittedByEmail || "",
      submittedAt: contribution.submittedAt || "",
      correctionReason: contribution.correctionReason || ""
    };
  }

  function authorizeTeacher(contribution, teacherId) {
    if (!contribution) throw new Error("Teacher contribution not found.");
    if (contribution.status === "Superseded") throw new Error("This assignment has been superseded.");
    if (!teacherById(teacherId) || contribution.teacherId !== text(teacherId)) {
      throw new Error("This teacher is not assigned to the contribution.");
    }
  }

  function createAssignment({ visitId, entry, visitStatus = "Active" } = {}) {
    assertDemoOnly();
    const safeVisitId = text(visitId || entry?.id);
    if (!safeVisitId) throw new Error("A stable synthetic visit ID is required.");
    const data = Storage.load();
    const existing = findByVisit(data, safeVisitId);
    if (existing) {
      appendNotification(data, existing, "TEACHER_CONTRIBUTION_REQUESTED", "duplicate assignment");
      Storage.save(data);
      return { contribution: clone(existing), duplicatePrevented: true };
    }

    const teacher = teacherByName(entry?.cameFromTeacher);
    const createdAt = now();
    const contribution = {
      contributionId: makeId("contribution"),
      visitId: safeVisitId,
      teacherId: teacher?.teacherId || "",
      teacherName: teacher?.displayName || text(entry?.cameFromTeacher) || "Other / Not Listed",
      teacherEmail: teacher?.fictionalEmail || "",
      teacherRole: teacher?.role || "",
      studentId: text(entry?.studentId),
      studentName: text(entry?.studentName),
      visitDate: text(entry?.date),
      referralReason: Array.isArray(entry?.behaviors) ? entry.behaviors.join(", ") : text(entry?.behaviors),
      biNotes: text(entry?.notes),
      timeIn: text(entry?.timeIn),
      visitStatus,
      assignmentVersion: 1,
      status: teacher ? "Requested" : "Needs Attention",
      classroomContext: "",
      strategiesAttempted: "",
      followUp: "",
      createdAt,
      updatedAt: createdAt,
      submittedAt: "",
      submittedBy: "",
      submittedByEmail: "",
      dueAt: makeDueAt(entry?.date),
      reminderSentAt: "",
      correctionReason: "",
      correctionRequestedAt: "",
      versions: [],
      originalContributionId: "",
      supersededAt: "",
      supersededBy: "",
      notificationHistory: []
    };

    data.contributions.push(contribution);
    if (teacher) {
      const event = appendNotification(data, contribution, "TEACHER_CONTRIBUTION_REQUESTED");
      contribution.notificationHistory.push(event.notificationId);
    }
    Storage.save(data);
    return { contribution: clone(contribution), duplicatePrevented: false };
  }

  function markVisitCompleted(visitId) {
    assertDemoOnly();
    const data = Storage.load();
    const contribution = findByVisit(data, visitId);
    if (!contribution) return null;
    contribution.visitStatus = "Completed";
    contribution.updatedAt = now();

    if (contribution.status === "Submitted") {
      const event = appendNotification(data, contribution, "TEACHER_CONTRIBUTION_REMINDER", "already submitted");
      contribution.notificationHistory.push(event.notificationId);
    } else if (contribution.reminderSentAt) {
      const event = appendNotification(data, contribution, "TEACHER_CONTRIBUTION_REMINDER", "duplicate reminder");
      contribution.notificationHistory.push(event.notificationId);
    } else if (contribution.teacherId) {
      const event = appendNotification(data, contribution, "TEACHER_CONTRIBUTION_REMINDER");
      contribution.reminderSentAt = event.createdAt;
      contribution.notificationHistory.push(event.notificationId);
    }
    Storage.save(data);
    return clone(contribution);
  }

  function saveDraft(contributionId, teacherId, fields) {
    assertDemoOnly();
    const data = Storage.load();
    const contribution = findContribution(data, contributionId);
    authorizeTeacher(contribution, teacherId);
    if (contribution.status === "Submitted") throw new Error("Request a correction before editing submitted notes.");
    NOTE_FIELDS.forEach(key => { contribution[key] = String(fields?.[key] ?? ""); });
    contribution.status = "Draft";
    contribution.updatedAt = now();
    Storage.save(data);
    return clone(contribution);
  }

  function submit(contributionId, teacherId, fields) {
    assertDemoOnly();
    const data = Storage.load();
    const contribution = findContribution(data, contributionId);
    authorizeTeacher(contribution, teacherId);
    if (contribution.status === "Submitted") {
      return { contribution: clone(contribution), duplicatePrevented: true };
    }
    const problems = requiredProblems(fields);
    if (problems.length) throw new Error("Complete the classroom context and strategies attempted fields before submitting.");

    NOTE_FIELDS.forEach(key => { contribution[key] = String(fields?.[key] ?? ""); });
    const submittedAt = now();
    const teacher = teacherById(teacherId);
    contribution.currentVersion = (contribution.currentVersion || 0) + 1;
    contribution.status = "Submitted";
    contribution.submittedAt = submittedAt;
    contribution.submittedBy = teacher?.displayName || "Demo Teacher";
    contribution.submittedByEmail = teacher?.fictionalEmail || "";
    contribution.updatedAt = submittedAt;
    contribution.versions.push({
      version: contribution.currentVersion,
      submittedAt,
      submittedBy: contribution.submittedBy,
      submittedByEmail: contribution.submittedByEmail,
      correctionReason: contribution.correctionReason || (contribution.currentVersion === 1 ? "Initial submission" : ""),
      fields: fieldsFrom(contribution)
    });
    contribution.correctionReason = "";
    contribution.correctionRequestedAt = "";
    Storage.save(data);
    return { contribution: clone(contribution), duplicatePrevented: false };
  }

  function requestCorrection(contributionId, teacherId, reason) {
    assertDemoOnly();
    const data = Storage.load();
    const contribution = findContribution(data, contributionId);
    authorizeTeacher(contribution, teacherId);
    if (contribution.status !== "Submitted") throw new Error("Only submitted contributions can enter correction mode.");
    const correctionReason = text(reason);
    if (!correctionReason) throw new Error("A correction reason is required.");
    contribution.status = "Correction requested";
    contribution.correctionReason = correctionReason;
    contribution.correctionRequestedAt = now();
    contribution.updatedAt = contribution.correctionRequestedAt;
    Storage.save(data);
    return clone(contribution);
  }

  function reassign(contributionId, teacherId, newTeacherId) {
    assertDemoOnly();
    const data = Storage.load();
    const original = findContribution(data, contributionId);
    if (!original) throw new Error("Teacher contribution not found.");
    if (original.status === "Superseded") throw new Error("This assignment has already been superseded.");
    const newTeacher = teacherById(newTeacherId);
    if (!newTeacher || !newTeacher.active) throw new Error("Choose an active demo teacher.");
    if (original.teacherId === newTeacher.teacherId) throw new Error("Choose a different teacher.");

    const createdAt = now();
    const replacement = {
      ...clone(original),
      contributionId: makeId("contribution"),
      teacherId: newTeacher.teacherId,
      teacherName: newTeacher.displayName,
      teacherEmail: newTeacher.fictionalEmail,
      teacherRole: newTeacher.role,
      assignmentVersion: (original.assignmentVersion || 1) + 1,
      status: "Requested",
      classroomContext: "",
      strategiesAttempted: "",
      followUp: "",
      createdAt,
      updatedAt: createdAt,
      submittedAt: "",
      submittedBy: "",
      submittedByEmail: "",
      reminderSentAt: "",
      correctionReason: "",
      correctionRequestedAt: "",
      versions: [],
      originalContributionId: original.contributionId,
      supersededAt: "",
      supersededBy: "",
      notificationHistory: []
    };
    original.status = "Superseded";
    original.supersededAt = createdAt;
    original.supersededBy = replacement.contributionId;
    original.updatedAt = createdAt;
    data.contributions.push(replacement);
    const event = appendNotification(data, replacement, "TEACHER_CONTRIBUTION_REQUESTED");
    replacement.notificationHistory.push(event.notificationId);
    Storage.save(data);
    return { original: clone(original), replacement: clone(replacement) };
  }

  function getForTeacher(teacherId) {
    assertDemoOnly();
    const data = Storage.load();
    return clone(data.contributions.filter(c => c.teacherId === text(teacherId) && c.status !== "Superseded"));
  }

  function getById(contributionId) {
    assertDemoOnly();
    const contribution = findContribution(Storage.load(), contributionId);
    return contribution ? clone(contribution) : null;
  }

  function getByVisit(visitId) {
    assertDemoOnly();
    const contribution = findByVisit(Storage.load(), visitId);
    return contribution ? clone(contribution) : null;
  }

  function getAll() {
    assertDemoOnly();
    return clone(Storage.load().contributions);
  }

  function getNotifications() {
    assertDemoOnly();
    return clone(Storage.load().notifications).sort((a, b) => String(b.createdAt).localeCompare(String(a.createdAt)));
  }

  function summary() {
    const all = getAll().filter(c => c.status !== "Superseded");
    return {
      total: all.length,
      requested: all.filter(c => c.status === "Requested").length,
      draft: all.filter(c => c.status === "Draft").length,
      submitted: all.filter(c => c.status === "Submitted").length,
      reminders: getNotifications().filter(e => e.eventType === "TEACHER_CONTRIBUTION_REMINDER" && !e.suppressedDuplicate).length,
      overdue: all.filter(c => c.status !== "Submitted" && c.dueAt && new Date(c.dueAt).getTime() < Date.now()).length,
      needsAttention: all.filter(c => c.status === "Needs Attention").length
    };
  }

  function isTeacherAuthorized(contributionId, teacherId) {
    const contribution = getById(contributionId);
    return Boolean(contribution && contribution.status !== "Superseded" && contribution.teacherId === text(teacherId));
  }

  return {
    CONFIG,
    DEMO_TEACHERS: clone(DEMO_TEACHERS),
    NOTE_FIELDS: NOTE_FIELDS.slice(),
    REQUIRED_FIELDS: REQUIRED_FIELDS.slice(),
    Storage,
    teacherById,
    teacherByName,
    createAssignment,
    markVisitCompleted,
    saveDraft,
    submit,
    requestCorrection,
    reassign,
    getForTeacher,
    getById,
    getByVisit,
    getAll,
    getNotifications,
    summary,
    isTeacherAuthorized,
    requiredProblems,
    fieldsFrom
  };
});
