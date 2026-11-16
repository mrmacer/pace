///////////////////////////////////////////////////////////////////////////////////////////////
// PATCH T1 — synthetic Teacher Referral & Contribution System UI
// Purpose: Focused demo teacher view, administrative contribution dashboard, notification
//          center, and MAC Walkthrough PACE Log preview. All data comes from the demo-only
//          TEACHER_CONTRIBUTIONS provider; no network calls are made here.
///////////////////////////////////////////////////////////////////////////////////////////////

(function (root) {
  const UI = {
    selectedTeacherId: "demo-teacher-ashford",
    activeContributionId: null,
    draftTimer: null,
    editorMode: "draft"
  };

  function esc(value) {
    if (typeof escHtml === "function") return escHtml(value);
    return String(value ?? "").replace(/[&<>"']/g, c => ({ "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;", "'": "&#39;" }[c]));
  }

  function text(value) { return String(value ?? "").trim(); }

  function dateLabel(value) {
    const raw = text(value);
    if (!raw) return "—";
    const date = new Date(raw.length === 10 ? `${raw}T12:00:00` : raw);
    return Number.isNaN(date.getTime()) ? raw : date.toLocaleDateString(undefined, { month: "short", day: "numeric", year: "numeric" });
  }

  function dateTimeLabel(value) {
    const raw = text(value);
    if (!raw) return "—";
    const date = new Date(raw);
    return Number.isNaN(date.getTime()) ? raw : date.toLocaleString(undefined, { month: "short", day: "numeric", hour: "numeric", minute: "2-digit" });
  }

  function isOverdue(contribution) {
    return contribution.status !== "Submitted" && contribution.dueAt && new Date(contribution.dueAt).getTime() < Date.now();
  }

  function statusClass(status) {
    return {
      Submitted: "submitted",
      Draft: "draft",
      "Correction requested": "correction",
      "Needs Attention": "attention"
    }[status] || "";
  }

  function statusBadge(contribution) {
    const label = isOverdue(contribution) ? `${contribution.status} · Overdue` : contribution.status;
    return `<span class="teacher-status ${statusClass(contribution.status)}${isOverdue(contribution) ? " overdue" : ""}">${esc(label)}</span>`;
  }

  function empty(message) { return `<p class="empty-hint">${esc(message)}</p>`; }

  function navigate(screen) {
    if (typeof nav === "function") nav(screen, "forward");
  }

  function toast(message, kind = "ok") {
    if (typeof showToast === "function") showToast(message, kind);
  }

  function init() {
    const tools = document.getElementById("teacherDemoTools");
    if (!tools || typeof APP_MODE !== "undefined" && APP_MODE !== "demo") return;
    tools.classList.remove("hidden");

    document.getElementById("openTeacherViewBtn")?.addEventListener("click", () => {
      renderTeacherAccountOptions();
      renderTeacherLanding();
      navigate("teacherView");
    });
    document.getElementById("openTeacherAdminBtn")?.addEventListener("click", () => {
      renderAdmin();
      navigate("teacherAdmin");
    });
    document.getElementById("openNotificationCenterBtn")?.addEventListener("click", () => {
      renderNotifications();
      navigate("notificationCenter");
    });
    document.getElementById("openWalkthroughPreviewBtn")?.addEventListener("click", () => {
      renderWalkthroughPreview();
      navigate("walkthroughPreview");
    });
    document.getElementById("teacherViewBackBtn")?.addEventListener("click", () => {
      closeEditor();
      navigate("home");
    });
    document.getElementById("teacherAdminBackBtn")?.addEventListener("click", () => navigate("home"));
    document.getElementById("notificationCenterBackBtn")?.addEventListener("click", () => navigate("home"));
    document.getElementById("walkthroughPreviewBackBtn")?.addEventListener("click", () => navigate("home"));
    document.getElementById("demoTeacherAccountSelect")?.addEventListener("change", event => {
      UI.selectedTeacherId = event.target.value;
      closeEditor();
      renderTeacherLanding();
    });
    document.getElementById("teacherAdminStatusFilter")?.addEventListener("change", renderAdmin);
    document.getElementById("teacherAdminTeacherFilter")?.addEventListener("change", renderAdmin);
  }

  function renderTeacherAccountOptions() {
    const select = document.getElementById("demoTeacherAccountSelect");
    if (!select) return;
    select.innerHTML = TEACHER_CONTRIBUTIONS.DEMO_TEACHERS.map(teacher =>
      `<option value="${esc(teacher.teacherId)}">${esc(teacher.displayName)} · ${esc(teacher.fictionalEmail)}</option>`
    ).join("");
    select.value = UI.selectedTeacherId;
  }

  function renderTeacherLanding() {
    const list = document.getElementById("teacherLandingList");
    const editor = document.getElementById("teacherContributionEditor");
    if (!list || !editor) return;
    const teacher = TEACHER_CONTRIBUTIONS.teacherById(UI.selectedTeacherId);
    const assignments = TEACHER_CONTRIBUTIONS.getForTeacher(UI.selectedTeacherId);
    editor.classList.add("hidden");
    list.classList.remove("hidden");
    const groupOrder = ["Requested", "Draft", "Submitted", "Correction requested"];
    const grouped = groupOrder.map(status => {
      const items = assignments.filter(c => c.status === status);
      return items.length ? `<div class="teacher-group-heading">${esc(status)}</div>${items.map(renderTeacherAssignmentCard).join("")}` : "";
    }).join("");
    list.innerHTML = `<div class="section-label">${esc(teacher?.displayName || "Demo Teacher")} · Assigned contributions</div>` +
      (assignments.length ? grouped : empty("No synthetic contributions are assigned to this teacher yet. Start a demo visit from Home."));
    list.querySelectorAll("[data-teacher-open]").forEach(button => button.addEventListener("click", () => openEditor(button.dataset.teacherOpen)));
  }

  function renderTeacherAssignmentCard(contribution) {
    const visitState = contribution.visitStatus === "Completed" ? "Completed" : "Active";
    const action = contribution.status === "Submitted" ? "View contribution" : "Open contribution";
    return `<article class="teacher-assignment-card">
      <div class="teacher-card-heading"><div><div class="teacher-card-title">${esc(contribution.studentName || "Synthetic student")}</div><div class="teacher-card-subtitle">Visit ${esc(dateLabel(contribution.visitDate))} · ${esc(visitState)}</div></div>${statusBadge(contribution)}</div>
      <div class="teacher-card-meta"><span>Due ${esc(dateLabel(contribution.dueAt))}</span><span>${contribution.reminderSentAt ? "Reminder generated" : "No reminder yet"}</span></div>
      <div class="teacher-card-actions"><button class="btn btn-secondary btn-lg" type="button" data-teacher-open="${esc(contribution.contributionId)}">${action}</button></div>
    </article>`;
  }

  function formValue(id) { return document.getElementById(id)?.value || ""; }

  function currentFields() {
    return {
      classroomContext: formValue("teacherClassroomContext"),
      strategiesAttempted: formValue("teacherStrategiesAttempted"),
      followUp: formValue("teacherFollowUp")
    };
  }

  function setSaveState(label, className = "") {
    const state = document.getElementById("teacherSaveState");
    if (!state) return;
    state.textContent = label;
    state.className = `teacher-save-state ${className}`;
  }

  async function saveDraftFromForm({ rerender = false, quiet = false } = {}) {
    if (!UI.activeContributionId) return;
    try {
      const contribution = TEACHER_CONTRIBUTIONS.saveDraft(UI.activeContributionId, UI.selectedTeacherId, currentFields());
      setSaveState("Saved", "saved");
      if (!quiet) toast("Teacher draft saved.");
      if (rerender) renderTeacherLanding();
      return contribution;
    } catch (err) {
      setSaveState("Save failed", "failed");
      if (!quiet) toast(err.message || "Teacher draft could not be saved.", "error");
      return null;
    }
  }

  function scheduleDraftSave() {
    setSaveState("Unsaved", "unsaved");
    clearTimeout(UI.draftTimer);
    UI.draftTimer = setTimeout(() => saveDraftFromForm({ quiet: true }), 550);
  }

  function openEditor(contributionId) {
    const contribution = TEACHER_CONTRIBUTIONS.getById(contributionId);
    if (!contribution || !TEACHER_CONTRIBUTIONS.isTeacherAuthorized(contributionId, UI.selectedTeacherId)) {
      toast("This contribution is no longer assigned to the selected teacher.", "error");
      renderTeacherLanding();
      return;
    }
    UI.activeContributionId = contributionId;
    UI.editorMode = contribution.status === "Submitted" ? "submitted" : "draft";
    const list = document.getElementById("teacherLandingList");
    const editor = document.getElementById("teacherContributionEditor");
    list.classList.add("hidden");
    editor.classList.remove("hidden");
    const readOnly = contribution.status === "Submitted";
    editor.innerHTML = `<div class="teacher-note-form">
      <button class="btn-text" id="teacherEditorBackBtn" type="button">&larr; Back to assignments</button>
      <div class="teacher-card-heading"><div><h3>${esc(contribution.studentName)}</h3><div class="teacher-card-subtitle">Visit ${esc(dateLabel(contribution.visitDate))} · ${esc(contribution.visitStatus)}</div></div>${statusBadge(contribution)}</div>
      ${contribution.visitStatus === "Completed" ? `<div class="field-hint">The visit is complete, but this contribution remains available for completion.</div>` : `<div class="field-hint">You may save a draft while the student is in PACE and return later.</div>`}
      <div class="form-card"><label class="walk-field-label" for="teacherClassroomContext">What happened before the student was referred to PACE? <span class="req">*</span></label><textarea class="note-textarea" id="teacherClassroomContext" ${readOnly ? "readonly" : ""}>${esc(contribution.classroomContext)}</textarea></div>
      <div class="form-card"><label class="walk-field-label" for="teacherStrategiesAttempted">What strategies or interventions did you try before the referral? <span class="req">*</span></label><textarea class="note-textarea" id="teacherStrategiesAttempted" ${readOnly ? "readonly" : ""}>${esc(contribution.strategiesAttempted)}</textarea></div>
      <div class="form-card"><label class="walk-field-label" for="teacherFollowUp">Is there anything else the PACE team should know or follow up on?</label><textarea class="note-textarea" id="teacherFollowUp" ${readOnly ? "readonly" : ""}>${esc(contribution.followUp)}</textarea></div>
      <div id="teacherSaveState" class="teacher-save-state">${readOnly ? `Submitted ${esc(dateTimeLabel(contribution.submittedAt))}` : "Not saved yet"}</div>
      ${readOnly ? `<div class="teacher-form-actions"><button class="btn btn-secondary btn-lg" id="teacherCorrectionBtn" type="button">Request / Make Correction</button></div><div class="form-card hidden" id="teacherCorrectionPanel"><label class="walk-field-label" for="teacherCorrectionReason">Correction reason <span class="req">*</span></label><input class="form-input-lg" id="teacherCorrectionReason" maxlength="240" placeholder="Additional classroom context"><button class="btn btn-primary btn-lg" id="teacherStartCorrectionBtn" type="button" style="margin-top:10px">Open Correction</button></div>` : `<div class="teacher-form-actions"><button class="btn btn-secondary btn-lg" id="teacherSaveDraftBtn" type="button">Save Draft</button><button class="btn btn-primary btn-lg" id="teacherSubmitBtn" type="button">Submit Contribution</button></div>`}
    </div>`;

    document.getElementById("teacherEditorBackBtn").addEventListener("click", () => {
      saveDraftFromForm({ quiet: true });
      closeEditor();
      renderTeacherLanding();
    });
    if (readOnly) {
      document.getElementById("teacherCorrectionBtn").addEventListener("click", () => {
        document.getElementById("teacherCorrectionPanel").classList.remove("hidden");
        document.getElementById("teacherCorrectionReason").focus();
      });
      document.getElementById("teacherStartCorrectionBtn").addEventListener("click", () => {
        const reason = text(formValue("teacherCorrectionReason"));
        if (!reason) {
          toast("A correction reason is required.", "error");
          return;
        }
        try {
          TEACHER_CONTRIBUTIONS.requestCorrection(contributionId, UI.selectedTeacherId, reason);
          toast("Correction mode opened.");
          openEditor(contributionId);
        } catch (err) { toast(err.message || "Correction could not be opened.", "error"); }
      });
      return;
    }
    ["teacherClassroomContext", "teacherStrategiesAttempted", "teacherFollowUp"].forEach(id => document.getElementById(id).addEventListener("input", scheduleDraftSave));
    document.getElementById("teacherSaveDraftBtn").addEventListener("click", () => saveDraftFromForm({ rerender: true }));
    document.getElementById("teacherSubmitBtn").addEventListener("click", () => {
      clearTimeout(UI.draftTimer);
      try {
        const result = TEACHER_CONTRIBUTIONS.submit(contributionId, UI.selectedTeacherId, currentFields());
        toast(result.duplicatePrevented ? "Contribution already submitted." : "Teacher contribution submitted.");
        closeEditor();
        renderTeacherLanding();
      } catch (err) { toast(err.message || "Contribution could not be submitted.", "error"); }
    });
  }

  function closeEditor() {
    clearTimeout(UI.draftTimer);
    UI.activeContributionId = null;
    const editor = document.getElementById("teacherContributionEditor");
    const list = document.getElementById("teacherLandingList");
    editor?.classList.add("hidden");
    list?.classList.remove("hidden");
  }

  function renderAdmin() {
    const summaryEl = document.getElementById("teacherAdminSummary");
    const list = document.getElementById("teacherAdminList");
    if (!summaryEl || !list) return;
    const summary = TEACHER_CONTRIBUTIONS.summary();
    const metrics = [["total", "Requested"], ["requested", "Awaiting submission"], ["draft", "Draft saved"], ["submitted", "Submitted"], ["reminders", "Reminders generated"], ["overdue", "Overdue"], ["needsAttention", "Needs attention"]];
    summaryEl.innerHTML = metrics.map(([key, label]) => `<div class="teacher-summary-card"><strong>${summary[key]}</strong><span>${esc(label)}</span></div>`).join("");

    const statusFilter = document.getElementById("teacherAdminStatusFilter")?.value || "";
    const teacherFilter = document.getElementById("teacherAdminTeacherFilter")?.value || "";
    const teacherSelect = document.getElementById("teacherAdminTeacherFilter");
    if (teacherSelect && teacherSelect.options.length === 1) {
      teacherSelect.innerHTML += TEACHER_CONTRIBUTIONS.DEMO_TEACHERS.map(t => `<option value="${esc(t.teacherId)}">${esc(t.displayName)}</option>`).join("");
      teacherSelect.value = teacherFilter;
    }
    const all = TEACHER_CONTRIBUTIONS.getAll().filter(c => (!statusFilter || c.status === statusFilter) && (!teacherFilter || c.teacherId === teacherFilter));
    list.innerHTML = all.length ? all.map(renderAdminCard).join("") : empty("No synthetic teacher contributions match these filters.");
    list.querySelectorAll("[data-admin-detail]").forEach(button => button.addEventListener("click", () => renderAdminDetails(button.dataset.adminDetail)));
  }

  function renderAdminCard(contribution) {
    return `<article class="teacher-admin-card"><div class="teacher-card-heading"><div><div class="teacher-card-title">${esc(contribution.studentName || "Synthetic student")}</div><div class="teacher-card-subtitle">${esc(contribution.teacherName)} · ${esc(dateLabel(contribution.visitDate))}</div></div>${statusBadge(contribution)}</div><div class="teacher-card-meta"><span>Visit: ${esc(contribution.visitStatus)}</span><span>${contribution.reminderSentAt ? "Reminder generated" : "No reminder"}</span><span>Updated ${esc(dateTimeLabel(contribution.updatedAt))}</span></div><div class="teacher-card-actions"><button class="btn btn-secondary btn-lg" type="button" data-admin-detail="${esc(contribution.contributionId)}">View details</button></div></article>`;
  }

  function renderAdminDetails(contributionId) {
    const contribution = TEACHER_CONTRIBUTIONS.getById(contributionId);
    const details = document.getElementById("teacherAdminDetails");
    if (!contribution || !details) return;
    const notifications = TEACHER_CONTRIBUTIONS.getNotifications().filter(event => event.contributionId === contributionId);
    const currentVersion = contribution.versions?.[contribution.versions.length - 1];
    details.classList.remove("hidden");
    details.innerHTML = `<div class="teacher-card-heading"><div><div class="teacher-card-title">${esc(contribution.studentName)}</div><div class="teacher-card-subtitle">${esc(contribution.teacherName)} · ${esc(contribution.teacherEmail || "No resolved email")}</div></div><button class="btn-text" id="closeTeacherAdminDetails" type="button">Close</button></div>
      ${statusBadge(contribution)}
      <div class="teacher-detail-section"><h4>Original referral context</h4><div class="teacher-detail-note">${esc(contribution.referralReason || "No synthetic referral reason recorded.")}</div></div>
      <div class="teacher-detail-section"><h4>Current teacher contribution</h4><div class="teacher-detail-note"><strong>Classroom context</strong>\n${esc(contribution.classroomContext || "—")}\n\n<strong>Strategies attempted</strong>\n${esc(contribution.strategiesAttempted || "—")}\n\n<strong>Follow-up</strong>\n${esc(contribution.followUp || "—")}</div></div>
      <div class="teacher-detail-section"><h4>Submission and correction history</h4>${contribution.versions?.length ? contribution.versions.map(version => `<div class="teacher-history-item"><strong>Version ${version.version}</strong> · ${esc(dateTimeLabel(version.submittedAt))}<br>${esc(version.correctionReason || "Initial submission")}<br><span class="teacher-card-subtitle">${esc(version.submittedBy || "Unknown author")}</span></div>`).join("") : `<div class="teacher-card-subtitle">No submitted versions yet.</div>`}</div>
      <div class="teacher-detail-section"><h4>Notification history</h4>${notifications.length ? notifications.map(event => `<div class="teacher-history-item"><strong>${esc(event.eventType)}</strong> · ${esc(dateTimeLabel(event.createdAt))}<br>${esc(event.simulatedDeliveryStatus)}${event.suppressedDuplicate ? ` · ${esc(event.suppressionReason)}` : ""}</div>`).join("") : `<div class="teacher-card-subtitle">No notification events.</div>`}</div>
      <div class="teacher-detail-section"><h4>Administrative reassignment</h4><div class="teacher-form-actions"><select class="form-input-lg" id="teacherReassignSelect">${TEACHER_CONTRIBUTIONS.DEMO_TEACHERS.map(t => `<option value="${esc(t.teacherId)}" ${t.teacherId === contribution.teacherId ? "selected" : ""}>${esc(t.displayName)}</option>`).join("")}</select><button class="btn btn-secondary btn-lg" id="teacherReassignBtn" type="button">Reassign</button></div></div>`;
    document.getElementById("closeTeacherAdminDetails").addEventListener("click", () => details.classList.add("hidden"));
    document.getElementById("teacherReassignBtn").addEventListener("click", () => {
      const newTeacherId = document.getElementById("teacherReassignSelect").value;
      try {
        TEACHER_CONTRIBUTIONS.reassign(contributionId, "demo-admin", newTeacherId);
        toast("Assignment superseded and reassigned. Simulated notification logged.");
        renderAdmin();
        renderAdminDetails(TEACHER_CONTRIBUTIONS.getAll().find(c => c.originalContributionId === contributionId)?.contributionId);
      } catch (err) { toast(err.message || "Assignment could not be reassigned.", "error"); }
    });
  }

  function renderNotifications() {
    const list = document.getElementById("notificationCenterList");
    if (!list) return;
    const events = TEACHER_CONTRIBUTIONS.getNotifications();
    list.innerHTML = events.length ? events.map(event => `<article class="notification-card${event.suppressedDuplicate ? " suppressed" : ""}"><div class="notification-event-type">${esc(event.eventType)}</div><div>Recipient: <strong>${esc(event.recipient)}</strong></div><div>Contribution: <code>${esc(event.contributionId)}</code></div><div>Created: ${esc(dateTimeLabel(event.createdAt))}</div><div class="notification-event-status">${esc(event.simulatedDeliveryStatus)}</div>${event.suppressedDuplicate ? `<div>Suppressed duplicate: ${esc(event.suppressionReason)}</div>` : ""}</article>`).join("") : empty("No simulated notifications yet. Start a synthetic visit from Home.");
  }

  function renderWalkthroughPreview() {
    const list = document.getElementById("walkthroughPreviewList");
    if (!list) return;
    const contributions = TEACHER_CONTRIBUTIONS.getAll().filter(c => c.status !== "Superseded");
    list.innerHTML = contributions.length ? contributions.map(contribution => `<article class="walkthrough-preview-card" tabindex="0" data-walkthrough-detail="${esc(contribution.contributionId)}"><div class="teacher-card-heading"><div><div class="teacher-card-title">${esc(contribution.studentName)}</div><div class="teacher-card-subtitle">${esc(dateLabel(contribution.visitDate))} · ${esc(contribution.teacherName)}</div></div><span class="teacher-status ${statusClass(contribution.status)}">Teacher Notes: ${esc(contribution.status)}</span></div><div class="walkthrough-preview-detail hidden" id="walkthrough-detail-${esc(contribution.contributionId)}"><strong>Current contribution</strong><p class="teacher-detail-note">${esc(contribution.classroomContext || "—")}\n\n${esc(contribution.strategiesAttempted || "—")}\n\n${esc(contribution.followUp || "")}</p><strong>Correction history</strong>${contribution.versions?.length ? contribution.versions.map(version => `<p class="teacher-card-subtitle">Version ${version.version} · ${esc(dateTimeLabel(version.submittedAt))} · ${esc(version.correctionReason || "Initial submission")}</p>`).join("") : `<p class="teacher-card-subtitle">No submitted versions yet.</p>`}</div></article>`).join("") : empty("No synthetic PACE visits have teacher contributions yet.");
    list.querySelectorAll("[data-walkthrough-detail]").forEach(card => {
      const toggle = () => document.getElementById(`walkthrough-detail-${card.dataset.walkthroughDetail}`)?.classList.toggle("hidden");
      card.addEventListener("click", toggle);
      card.addEventListener("keydown", event => { if (event.key === "Enter" || event.key === " ") { event.preventDefault(); toggle(); } });
    });
  }

  function showAssignmentConfirmation(visitId) {
    if (typeof APP_MODE !== "undefined" && APP_MODE !== "demo") return;
    const contribution = TEACHER_CONTRIBUTIONS.getByVisit(visitId);
    if (!contribution) return;
    if (contribution.status === "Needs Attention") {
      toast("Teacher contribution needs attention — no notification generated.", "error");
    } else {
      toast(`Teacher contribution requested for ${contribution.teacherName}. Simulation only.`);
    }
  }

  function refreshVisible() {
    if (typeof currentScreenName === "undefined") return;
    if (currentScreenName === "teacherView") { renderTeacherAccountOptions(); renderTeacherLanding(); }
    if (currentScreenName === "teacherAdmin") renderAdmin();
    if (currentScreenName === "notificationCenter") renderNotifications();
    if (currentScreenName === "walkthroughPreview") renderWalkthroughPreview();
  }

  root.TEACHER_CONTRIBUTIONS_UI = {
    init,
    showAssignmentConfirmation,
    refreshVisible,
    renderTeacherLanding,
    renderAdmin,
    renderNotifications,
    renderWalkthroughPreview
  };
})(typeof globalThis !== "undefined" ? globalThis : window);
