/* ─────────────────────────────────────────────────────────────────────────
   IEP Skook Access Model v1 — IEP_App_Users reader (PATCH A)

   NEW authorization layer. IEP_Users2 (auth.js) remains this app's identity
   / Active / Role source — this file does not touch it and does not change
   anything AUTH already does.

   IEP_App_Users is a separate, currently-unpopulated SharePoint list that
   holds the explicit per-application permission flags (Daily Pulse / PACE /
   Walkthrough / Admin Panel) plus Teacher/Classroom and Display Name. Email
   is the join key between the signed-in Microsoft account, IEP_Users2, and
   this list. Internal SharePoint field names are never hard-coded — this
   file resolves them live via GRAPH.getListSchema(), the same pattern every
   other list read in this app already uses (see roster.js, graph.js).

   THIS PATCH DOES NOT CHANGE APP BEHAVIOR. Nothing in app.js or auth.js
   calls APP_USERS yet — this file only defines the reader so a later patch
   (the PACE gate) can call it. Loading this script has zero runtime effect
   until something invokes APP_USERS.resolve().

   MIGRATION MODE (temporary — see ACCESS_MODEL_CONFIG.ENFORCE_APP_USERS
   below): IEP_App_Users is not populated yet. While ENFORCE_APP_USERS is
   false, a later patch's gate is expected to treat "no matching row" as
   "fall back to this app's existing IEP_Users2-based authorization"
   instead of locking everyone out. Flip ENFORCE_APP_USERS to true only
   once every active staff member has a row in IEP_App_Users — after that,
   a missing row means denied, full stop, no fallback. This flag is
   intentionally NOT exposed in any UI — code-only, changed by a developer.
   ───────────────────────────────────────────────────────────────────────── */

const ACCESS_MODEL_CONFIG = {
  // TEMPORARY MIGRATION SWITCH.
  //   false = MIGRATION MODE — no IEP_App_Users row for a user falls back
  //           to this app's pre-existing authorization behavior.
  //   true  = ENFORCED MODE  — no IEP_App_Users row means denied, always.
  // Flip once IEP_App_Users is confirmed populated for all active staff.
  ENFORCE_APP_USERS: false
};

function normalizeAppUsersBoolean(value, defaultValue = false) {
  if (value === undefined || value === null || value === "") return defaultValue;
  if (typeof value === "boolean") return value;
  if (typeof value === "number") return value === 1;
  return ["yes", "true", "1"].includes(String(value).trim().toLowerCase());
}

const APP_USERS = {
  _row: null,
  _checked: false,
  lookupError: null,

  // Looks up the IEP_App_Users row whose Email matches (case/whitespace-
  // insensitive). Returns null — and records lookupError — on any Graph or
  // list failure; a failed lookup must never be treated as a match by a
  // caller. Demo mode never touches Graph at all, mirroring auth.js/graph.js.
  async resolve(email) {
    this._checked = true;
    this._row = null;
    this.lookupError = null;
    const target = String(email || "").toLowerCase().trim();
    if (!target) return null;

    if (typeof APP_MODE !== "undefined" && APP_MODE === "demo") {
      return null;
    }

    try {
      const [items, schema] = await Promise.all([
        GRAPH.getListItems(CONFIG.LISTS.appUsers),
        GRAPH.getListSchema(CONFIG.LISTS.appUsers)
      ]);
      const displayByInternal = {};
      Object.entries(schema).forEach(([display, internal]) => { displayByInternal[internal] = display; });
      const rows = items.map(item => {
        const row = { id: item.id };
        Object.entries(item).forEach(([key, value]) => {
          if (key !== "id") row[displayByInternal[key] || key] = value;
        });
        return row;
      });
      const match = rows.find(row => String(row["Email"] || "").toLowerCase().trim() === target) || null;
      this._row = match;
      return match;
    } catch (err) {
      console.error("IEP_App_Users lookup failed:", err.message || String(err));
      this.lookupError = err.message || "Unable to verify application permissions.";
      return null;
    }
  },

  get found() { return this._row !== null; },

  permissionsFromRow(row) {
    if (!row) return null;
    return {
      dailyPulse:  normalizeAppUsersBoolean(row["Daily Pulse"]),
      pace:        normalizeAppUsersBoolean(row["PACE"]),
      walkthrough: normalizeAppUsersBoolean(row["Walkthrough"]),
      adminPanel:  normalizeAppUsersBoolean(row["Admin Panel"])
    };
  },

  get permissions()      { return this.permissionsFromRow(this._row); },
  get teacherClassroom() { return this._row ? String(this._row["Teacher/Classroom"] || "").trim() : ""; },
  get displayName()      { return this._row ? String(this._row["Display Name"] || "").trim() : ""; },
  get role()             { return this._row ? String(this._row["Role"] || "").trim() : ""; },

  // PATCH C: generic authorization decision, callable after resolve() has
  // run (uses the outcome it just recorded). An explicit IEP_App_Users flag
  // is authoritative whenever a row exists — even if it disagrees with
  // legacyAllowed. No row falls back to legacyAllowed (this app's own
  // pre-existing behavior) unless ENFORCE_APP_USERS is on. A failed lookup
  // (this.lookupError set) always denies — never falls back to legacyAllowed.
  decide(permissionKey, legacyAllowed) {
    if (this.lookupError) return { allowed: false, reason: "app-users-lookup-failed" };
    if (this._row) {
      const allowed = normalizeAppUsersBoolean(this._row[permissionKey]);
      return { allowed, reason: allowed ? "app-users-permission" : "app-users-denied" };
    }
    if (ACCESS_MODEL_CONFIG.ENFORCE_APP_USERS) return { allowed: false, reason: "enforced-no-row" };
    return { allowed: Boolean(legacyAllowed), reason: legacyAllowed ? "legacy-fallback" : "legacy-denied" };
  }
};
