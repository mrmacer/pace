///////////////////////////////////////////////////////////////////////////////////////////////
// Author: R-E Miller & Greg Macer
// Creation Date: August 31, 2026
// Filename: iep-app-users.js
// Purpose: Reads explicit per-application permissions from IEP_App_Users and converts them
//          into a small authorization decision model, separate from auth.js (which handles
//          identity and active-staff status only). IEP_App_Users is joined to the signed-in
//          user by Email, and internal SharePoint field names are resolved live via
//          GRAPH.getListSchema() rather than hard-coded. ACCESS_MODEL_CONFIG.ENFORCE_APP_USERS
//          is a temporary migration switch: while false, a missing IEP_App_Users row falls
//          back to this app's pre-existing authorization instead of denying access; leaving
//          it false after migration is an access-control bypass, so it must be flipped to
//          true only once every active staff member has a row in IEP_App_Users.
///////////////////////////////////////////////////////////////////////////////////////////////

const ACCESS_MODEL_CONFIG = {
  // TEMPORARY MIGRATION SWITCH.
  //   false = MIGRATION MODE — no IEP_App_Users row for a user falls back
  //           to this app's pre-existing authorization behavior.
  //   true  = ENFORCED MODE  — no IEP_App_Users row means denied, always.
  // Flip once IEP_App_Users is confirmed populated for all active staff.
  // REVIEW: leaving this false permits the legacy authorization fallback;
  // switch only after every active staff member has an IEP_App_Users row.
  ENFORCE_APP_USERS: false
};

///////////////////////////////////////////////////////////////////////////////////////////////
// Function Name: normalizeAppUsersBoolean
// Description: Normalizes an application permission value from its various SharePoint
//              representations (boolean, number, or string) into a real boolean.
// Parameters: any value - raw permission value read from SharePoint - input
//             boolean defaultValue - value to return when value is empty (default false) - input
///////////////////////////////////////////////////////////////////////////////////////////////
function normalizeAppUsersBoolean(value, defaultValue = false) {
  if (value === undefined || value === null || value === "") return defaultValue;
  if (typeof value === "boolean") return value;
  if (typeof value === "number") return value === 1;
  return ["yes", "true", "1"].includes(String(value).trim().toLowerCase());
}

// Application-specific permission reader and decision facade.
// _row: matched IEP_App_Users row, if one exists (else null).
// _checked: whether resolve() has run for this session.
// lookupError: error message that caused a fail-closed decision, if any.
const APP_USERS = {
  _row: null,
  _checked: false,
  lookupError: null,

  // Looks up the IEP_App_Users row whose Email matches (case/whitespace-
  // insensitive). Returns null — and records lookupError — on any Graph or
  // list failure; a failed lookup must never be treated as a match by a
  // caller. Demo mode never touches Graph at all, mirroring auth.js/graph.js.
  ///////////////////////////////////////////////////////////////////////////////////////////////
  // Function Name: resolve
  // Description: Resolves one signed-in email to its explicit IEP_App_Users permission row,
  //              recording the matched row (or a lookup error) for later use by decide().
  // Parameters: string email - signed-in user's email to match against IEP_App_Users - input
  ///////////////////////////////////////////////////////////////////////////////////////////////
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

  ///////////////////////////////////////////////////////////////////////////////////////////////
  // Function Name: found
  // Description: Reports whether resolve() found a matching IEP_App_Users row.
  // Parameters: none
  ///////////////////////////////////////////////////////////////////////////////////////////////
  get found() { return this._row !== null; },

  ///////////////////////////////////////////////////////////////////////////////////////////////
  // Function Name: permissionsFromRow
  // Description: Converts one IEP_App_Users row into the app's boolean permission model
  //              (Daily Pulse / PACE / Walkthrough / Admin Panel).
  // Parameters: object row - matched IEP_App_Users row, or null - input
  ///////////////////////////////////////////////////////////////////////////////////////////////
  permissionsFromRow(row) {
    if (!row) return null;
    return {
      dailyPulse:  normalizeAppUsersBoolean(row["Daily Pulse"]),
      pace:        normalizeAppUsersBoolean(row["PACE"]),
      walkthrough: normalizeAppUsersBoolean(row["Walkthrough"]),
      adminPanel:  normalizeAppUsersBoolean(row["Admin Panel"])
    };
  },

  ///////////////////////////////////////////////////////////////////////////////////////////////
  // Function Name: permissions
  // Description: Returns the boolean permission object for the currently matched row.
  // Parameters: none
  ///////////////////////////////////////////////////////////////////////////////////////////////
  get permissions()      { return this.permissionsFromRow(this._row); },

  ///////////////////////////////////////////////////////////////////////////////////////////////
  // Function Name: teacherClassroom
  // Description: Returns the matched row's Teacher/Classroom field, or an empty string.
  // Parameters: none
  ///////////////////////////////////////////////////////////////////////////////////////////////
  get teacherClassroom() { return this._row ? String(this._row["Teacher/Classroom"] || "").trim() : ""; },

  ///////////////////////////////////////////////////////////////////////////////////////////////
  // Function Name: displayName
  // Description: Returns the matched row's Display Name field, or an empty string.
  // Parameters: none
  ///////////////////////////////////////////////////////////////////////////////////////////////
  get displayName()      { return this._row ? String(this._row["Display Name"] || "").trim() : ""; },

  ///////////////////////////////////////////////////////////////////////////////////////////////
  // Function Name: role
  // Description: Returns the matched row's Role field, or an empty string.
  // Parameters: none
  ///////////////////////////////////////////////////////////////////////////////////////////////
  get role()             { return this._row ? String(this._row["Role"] || "").trim() : ""; },

  // PATCH C: generic authorization decision, callable after resolve() has
  // run (uses the outcome it just recorded). An explicit IEP_App_Users flag
  // is authoritative whenever a row exists — even if it disagrees with
  // legacyAllowed. No row falls back to legacyAllowed (this app's own
  // pre-existing behavior) unless ENFORCE_APP_USERS is on. A failed lookup
  // (this.lookupError set) always denies — never falls back to legacyAllowed.
  ///////////////////////////////////////////////////////////////////////////////////////////////
  // Function Name: decide
  // Description: Applies explicit IEP_App_Users permissions, migration fallback, and
  //              lookup-failure denial to produce a single allow/deny authorization decision.
  // Parameters: string permissionKey - permission flag name to check (e.g. "PACE") - input
  //             boolean legacyAllowed - this app's pre-existing authorization result,
  //                                     used as the migration-mode fallback - input
  ///////////////////////////////////////////////////////////////////////////////////////////////
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
