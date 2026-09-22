///////////////////////////////////////////////////////////////////////////////////////////////
// Author: R-E Miller & Greg Macer
// Creation Date: August 20, 2026
// Filename: graph.js
// Purpose: Narrow Microsoft Graph client for the three SharePoint lists used by PACE Room
//          Tracker (users, students, PACE visits), owning authentication transport, list/schema
//          discovery, display-name field mapping, pagination, and visit CRUD. Callers always
//          supply SharePoint display names; this module resolves them to internal names from
//          the live column schema and never hard-codes internal field names, so reads are
//          converted back to display-name keys and the UI stays independent of SharePoint's
//          internal naming rules. Every network method calls assertGraphAllowed() before
//          acquiring a token, so demo mode fails closed at the transport boundary rather than
//          relying on each caller to remember a mode check.
///////////////////////////////////////////////////////////////////////////////////////////////

// Hard safety block (demo mode): every Graph call funnels through _get/
// _post/_patch below, so guarding those three is a single choke point that
// covers every higher-level method transitively (getSiteId, getListId,
// findUserByEmail, savePaceVisit, closePaceVisit, everything) — no call
// site needs its own check, and none can be added later that forgets one.
///////////////////////////////////////////////////////////////////////////////////////////////
// Function Name: assertGraphAllowed
// Description: Throws when any code path attempts Microsoft Graph access while the app is
//              running in demo mode.
// Parameters: none
///////////////////////////////////////////////////////////////////////////////////////////////
function assertGraphAllowed() {
  if (typeof APP_MODE !== "undefined" && APP_MODE === "demo") {
    throw new Error("Demo mode safety block: Microsoft Graph access is disabled.");
  }
}

///////////////////////////////////////////////////////////////////////////////////////////////
// Function Name: isWritableSharePointField
// Description: Rejects SharePoint system, read-only, and underscore-prefixed fields before a
//              write payload is mapped to internal column names.
// Parameters: string internalName - internal SharePoint column name to check - input
///////////////////////////////////////////////////////////////////////////////////////////////
function isWritableSharePointField(internalName) {
  const blocked = new Set([
    "id", "ContentType", "Modified", "Created", "Author", "Editor",
    "AuthorLookupId", "EditorLookupId", "_UIVersionString", "Attachments",
    "Edit", "LinkTitle", "LinkTitleNoMenu", "ItemChildCount", "FolderChildCount",
    "_ComplianceFlags", "_ComplianceTag", "_ComplianceTagWrittenTime", "_ComplianceTagUserId"
  ]);
  if (!internalName) return false;
  if (blocked.has(internalName)) return false;
  if (internalName.startsWith("_")) return false;
  return true;
}

// Microsoft Graph gateway and SharePoint field-mapping service. Properties: _BASE (Graph API
// root), _SITE (configured SharePoint site locator), _siteId (cached resolved site id),
// _listIdCache (list-name to id cache), _schemaCache (list schema cache).
const GRAPH = {
  _BASE:        "https://graph.microsoft.com/v1.0",
  _SITE:        CONFIG.SITE,
  _siteId:      null,
  _listIdCache: {},
  _schemaCache: {},

  ///////////////////////////////////////////////////////////////////////////////////////////////
  // Function Name: _get
  // Description: Performs an authenticated Microsoft Graph GET request and returns the parsed
  //              JSON response body, throwing when Graph returns a non-2xx status.
  // Parameters: string path - relative Graph API path without the base URL - input
  //             Object additionalHeaders - extra request headers to merge in (optional) - input
  ///////////////////////////////////////////////////////////////////////////////////////////////
  async _get(path, additionalHeaders = {}) {
    assertGraphAllowed();
    const token = await AUTH.acquireGraphToken();
    const resp  = await fetch(`${this._BASE}/${path}`, {
      headers: { Authorization: `Bearer ${token}`, ...additionalHeaders }
    });
    if (!resp.ok) {
      const body = await resp.text().catch(() => "");
      // REVIEW: response bodies are useful while diagnosing Graph failures,
      // but may contain sensitive metadata on a shared kiosk console.
      console.error("Graph GET failed", path, resp.status, body);
      throw new Error(`Graph ${resp.status}: ${body}`);
    }
    return resp.json();
  },

  ///////////////////////////////////////////////////////////////////////////////////////////////
  // Function Name: _post
  // Description: Performs an authenticated Microsoft Graph POST request with a JSON body and
  //              returns the parsed response, throwing when Graph rejects the request.
  // Parameters: string path - relative Graph API path - input
  //             Object body - JSON-serializable request body to send - input
  ///////////////////////////////////////////////////////////////////////////////////////////////
  async _post(path, body) {
    assertGraphAllowed();
    const token = await AUTH.acquireGraphToken();
    const resp  = await fetch(`${this._BASE}/${path}`, {
      method: "POST",
      headers: { Authorization: `Bearer ${token}`, "Content-Type": "application/json" },
      body: JSON.stringify(body)
    });
    if (!resp.ok) {
      const text = await resp.text().catch(() => "");
      // REVIEW: avoid retaining raw Graph error bodies in production logs.
      console.error("Graph POST failed", path, resp.status, text);
      throw new Error(`Graph POST ${resp.status}: ${text}`);
    }
    return resp.json();
  },

  ///////////////////////////////////////////////////////////////////////////////////////////////
  // Function Name: _patch
  // Description: Performs an authenticated Microsoft Graph PATCH request, returning the parsed
  //              response JSON or null on an HTTP 204, and throwing when Graph rejects the patch.
  // Parameters: string path - relative Graph API path - input
  //             Object body - JSON-serializable field patch to send - input
  ///////////////////////////////////////////////////////////////////////////////////////////////
  async _patch(path, body) {
    assertGraphAllowed();
    const token = await AUTH.acquireGraphToken();
    const resp  = await fetch(`${this._BASE}/${path}`, {
      method: "PATCH",
      headers: { Authorization: `Bearer ${token}`, "Content-Type": "application/json" },
      body: JSON.stringify(body)
    });
    if (!resp.ok) {
      const text = await resp.text().catch(() => "");
      // REVIEW: sanitize this diagnostic payload before production use.
      console.error("Graph PATCH failed", { path, status: resp.status, response: text });
      throw new Error(`Graph PATCH ${resp.status}: ${text}`);
    }
    if (resp.status === 204) return null;
    return resp.json();
  },

  // STAFF CORRECTIONS: single-item DELETE. A successful Graph delete is
  // HTTP 204 with no body, so nothing is parsed on success; any non-2xx
  // throws with the status attached (never a silent success).
  ///////////////////////////////////////////////////////////////////////////////////////////////
  // Function Name: _delete
  // Description: Performs an authenticated Microsoft Graph DELETE for one list resource,
  //              throwing (with the HTTP status attached) on any non-2xx response.
  // Parameters: string path - relative Graph API path of the resource to delete - input
  ///////////////////////////////////////////////////////////////////////////////////////////////
  async _delete(path) {
    assertGraphAllowed();
    const token = await AUTH.acquireGraphToken();
    const resp  = await fetch(`${this._BASE}/${path}`, {
      method: "DELETE",
      headers: { Authorization: `Bearer ${token}` }
    });
    if (!resp.ok) {
      const text = await resp.text().catch(() => "");
      // REVIEW: deletion failures should expose status, not raw SharePoint data.
      console.error("Graph DELETE failed", { path, status: resp.status, response: text });
      const error = new Error(`Graph DELETE ${resp.status}`);
      error.status = resp.status;
      throw error;
    }
    return null;
  },

  ///////////////////////////////////////////////////////////////////////////////////////////////
  // Function Name: getSiteId
  // Description: Resolves and caches the configured SharePoint site id, reusing the cached
  //              value on subsequent calls.
  // Parameters: none
  ///////////////////////////////////////////////////////////////////////////////////////////////
  async getSiteId() {
    if (this._siteId) return this._siteId;
    const data = await this._get(`sites/${this._SITE}`);
    this._siteId = data.id;
    return data.id;
  },

  ///////////////////////////////////////////////////////////////////////////////////////////////
  // Function Name: getListId
  // Description: Resolves and caches a SharePoint list id, matching by either its display or
  //              internal name.
  // Parameters: string listName - configured SharePoint list name to resolve - input
  ///////////////////////////////////////////////////////////////////////////////////////////////
  async getListId(listName) {
    if (this._listIdCache[listName]) return this._listIdCache[listName];
    const siteId = await this.getSiteId();
    const data   = await this._get(`sites/${siteId}/lists?$select=id,name,displayName`);
    const match  = (data.value || []).find(list => {
      const names = [list.name, list.displayName].map(v => String(v || "").toLowerCase().trim());
      return names.includes(String(listName).toLowerCase().trim());
    });
    if (!match) throw new Error(`SharePoint list not found: ${listName}`);
    this._listIdCache[listName] = match.id;
    return match.id;
  },

  ///////////////////////////////////////////////////////////////////////////////////////////////
  // Function Name: getListItems
  // Description: Reads every item in a list, following Graph's @odata.nextLink pagination until
  //              exhausted, and returns each row with its stable `id` plus field values.
  // Parameters: string listName - configured SharePoint list name to read - input
  ///////////////////////////////////////////////////////////////////////////////////////////////
  async getListItems(listName) {
    const siteId = await this.getSiteId();
    const listId = await this.getListId(listName);
    // Graph pages list items (~200/page). IEP_Pace_Visits grows all school
    // year, so "Currently in PACE" must follow @odata.nextLink to the end
    // or open visits created after page 1 would silently vanish from the
    // room dashboard. nextLink is a full absolute URL — fetch it directly
    // rather than re-deriving a relative path.
    let path  = `sites/${siteId}/lists/${listId}/items?$expand=fields&$top=200`;
    let items = [];
    while (path) {
      const data = await this._get(path);
      items = items.concat(data.value || []);
      const nextLink = data["@odata.nextLink"];
      path = nextLink ? nextLink.replace(`${this._BASE}/`, "") : null;
    }
    // item.id is the stable SharePoint list-item id — required for the
    // Time-Out PATCH later; never derive it from any display value.
    return items.map(item => ({ id: item.id, ...item.fields }));
  },

  ///////////////////////////////////////////////////////////////////////////////////////////////
  // Function Name: getListSchema
  // Description: Reads and caches the display-name to internal-name field schema for a list from
  //              its live SharePoint column definitions.
  // Parameters: string listName - configured SharePoint list name whose schema to read - input
  ///////////////////////////////////////////////////////////////////////////////////////////////
  async getListSchema(listName) {
    if (this._schemaCache[listName]) return this._schemaCache[listName];
    const siteId = await this.getSiteId();
    const listId = await this.getListId(listName);
    const data   = await this._get(`sites/${siteId}/lists/${listId}/columns?$select=name,displayName`);
    const schema = {};
    (data.value || []).forEach(col => { if (col.displayName && col.name) schema[col.displayName] = col.name; });
    this._schemaCache[listName] = schema;
    return schema;
  },

  ///////////////////////////////////////////////////////////////////////////////////////////////
  // Function Name: mapFields
  // Description: Maps a set of display-name keyed field values to writable live SharePoint
  //              internal field names, dropping any unmapped or read-only fields.
  // Parameters: string listName - list whose live column schema should be used - input
  //             Object displayFields - display-name keyed values to map - input
  ///////////////////////////////////////////////////////////////////////////////////////////////
  async mapFields(listName, displayFields) {
    const schema = await this.getListSchema(listName);
    const mapped = {};
    Object.entries(displayFields).forEach(([displayName, value]) => {
      if (value === undefined || value === null) return;
      const internalName = schema[displayName];
      if (!internalName) { console.warn(`UNMAPPED FIELD — display: "${displayName}"`); return; }
      if (!isWritableSharePointField(internalName)) { console.warn(`SKIPPED (read-only) — "${displayName}"`); return; }
      mapped[internalName] = value;
    });
    return mapped;
  },

  ///////////////////////////////////////////////////////////////////////////////////////////////
  // Function Name: createListItem
  // Description: Creates a new SharePoint list item from a payload of already internal-name
  //              mapped fields.
  // Parameters: string listName - configured SharePoint list name to create the item in - input
  //             Object fields - internal-name keyed field values to write - input
  ///////////////////////////////////////////////////////////////////////////////////////////////
  async createListItem(listName, fields) {
    const siteId = await this.getSiteId();
    const listId = await this.getListId(listName);
    return this._post(`sites/${siteId}/lists/${listId}/items`, { fields });
  },

  ///////////////////////////////////////////////////////////////////////////////////////////////
  // Function Name: createMappedListItem
  // Description: Maps display-name keyed fields to internal names and creates one SharePoint
  //              list item from the result.
  // Parameters: string listName - configured SharePoint list name to create the item in - input
  //             Object displayFields - display-name keyed values to write - input
  ///////////////////////////////////////////////////////////////////////////////////////////////
  async createMappedListItem(listName, displayFields) {
    const fields = await this.mapFields(listName, displayFields);
    return this.createListItem(listName, fields);
  },

  ///////////////////////////////////////////////////////////////////////////////////////////////
  // Function Name: updateMappedListItem
  // Description: Maps display-name keyed fields to internal names and patches one existing
  //              SharePoint list item with the result.
  // Parameters: string listName - configured SharePoint list name containing the item - input
  //             string itemId - SharePoint list item id to patch - input
  //             Object displayFields - display-name keyed values to write - input
  ///////////////////////////////////////////////////////////////////////////////////////////////
  async updateMappedListItem(listName, itemId, displayFields) {
    const siteId = await this.getSiteId();
    const listId = await this.getListId(listName);
    const fields = await this.mapFields(listName, displayFields);
    return this._patch(`sites/${siteId}/lists/${listId}/items/${itemId}/fields`, fields);
  },

  ///////////////////////////////////////////////////////////////////////////////////////////////
  // Function Name: findListItemByDisplayField
  // Description: Finds the first list item whose mapped display-name field equals a given value.
  // Parameters: string listName - configured SharePoint list name to search - input
  //             string displayFieldName - display-name field to compare - input
  //             * value - value to match against the field, compared as trimmed strings - input
  ///////////////////////////////////////////////////////////////////////////////////////////////
  async findListItemByDisplayField(listName, displayFieldName, value) {
    const [items, schema] = await Promise.all([this.getListItems(listName), this.getListSchema(listName)]);
    const internalName = schema[displayFieldName];
    if (!internalName) throw new Error(`Column not found in ${listName}: ${displayFieldName}`);
    return items.find(item => String(item[internalName] || "").trim() === String(value || "").trim()) || null;
  },

  ///////////////////////////////////////////////////////////////////////////////////////////////
  // Function Name: findUserByEmail
  // Description: Finds the approved IEP_Users2 staff row matching an account email address,
  //              checking several tolerated email field name aliases.
  // Parameters: string email - account email address to match, case/whitespace-insensitive - input
  ///////////////////////////////////////////////////////////////////////////////////////////////
  async findUserByEmail(email) {
    const users  = await this.getListItems(CONFIG.LISTS.users);
    const target = String(email || "").toLowerCase().trim();
    const match = users.find(user => {
      const possible = [user.Email, user.email, user.EMail, user.Email0, user.EmailAddress, user.Email_x0020_Address, user["Email Address"]];
      return possible.some(v => String(v || "").toLowerCase().trim() === target);
    }) || null;
    if (!match) throw new Error("No matching IEP_Users2 account found for: " + target);
    return match;
  },

  // PATCH 005: reusable data-provider for "Active users with a given
  // Role" — used for the Behavior Specialist picker, but written generic
  // in case another role-scoped list is ever needed. Deliberately mirrors
  // AUTH.loadStaffFromSharePoint()'s exact field-candidate reads
  // (field_1/Name/Title, field_2/Role, field_3/Active) rather than
  // resolving IEP_Users2 via getListSchema() the way IEP_Pace_Visits/
  // IEP_Students_2026_27 are — this list's internal names were already
  // being read this defensive way by proven, working production auth
  // code, so the new method matches it instead of introducing a second,
  // untested resolution path for the same list. Never reads/returns any
  // email field — callers only ever see name/role/active/id.
  ///////////////////////////////////////////////////////////////////////////////////////////////
  // Function Name: getActiveUsersByRole
  // Description: Returns active IEP_Users2 people whose role matches exactly, sorted by name.
  // Parameters: string role - role name to match exactly (case-insensitive) - input
  ///////////////////////////////////////////////////////////////////////////////////////////////
  async getActiveUsersByRole(role) {
    const users  = await this.getListItems(CONFIG.LISTS.users);
    const target = String(role || "").trim().toLowerCase();
    return users
      .map(user => {
        const rawRole = user.field_2 || user.Role || "";
        const userRole = Array.isArray(rawRole) ? rawRole[0] : rawRole;
        const active = user.field_3 ?? user.Active;
        const isActive = active === true || active === "Yes" || active === "true" || active === 1;
        const name = String(user.field_1 || user.Name || user.Title || "").trim();
        return { id: user.id, name, role: String(userRole || "").trim(), active: isActive };
      })
      .filter(u => u.name && u.active && u.role.toLowerCase() === target)
      .sort((a, b) => a.name.localeCompare(b.name));
  },

  /* ── PACE-specific helpers ─────────────────────────────────────────────── */

  // Every list item, mapped back to DISPLAY field names via the schema, so
  // the rest of the app never touches internal SharePoint names.
  ///////////////////////////////////////////////////////////////////////////////////////////////
  // Function Name: getPaceVisitsByDisplayName
  // Description: Reads every PACE visit item and converts its field keys from internal names
  //              back to display names.
  // Parameters: none
  ///////////////////////////////////////////////////////////////////////////////////////////////
  async getPaceVisitsByDisplayName() {
    const [items, schema] = await Promise.all([
      this.getListItems(CONFIG.LISTS.paceVisits),
      this.getListSchema(CONFIG.LISTS.paceVisits)
    ]);
    const inv = {};
    Object.entries(schema).forEach(([display, internal]) => { inv[internal] = display; });
    return items.map(item => {
      const row = { id: item.id };
      Object.entries(item).forEach(([k, v]) => { if (k !== "id") row[inv[k] || k] = v; });
      return row;
    });
  },

  // PATCH 008: bounded Recent Activity/current-room read. The caller
  // supplies the iPad's LOCAL calendar date; Graph filters the real
  // IEP_Pace_Visits Date column before any rows reach the device. This is
  // deliberately separate from the legacy all-items reader above so the
  // shared iPad never needs a school-year history just to show today.
  //
  // A two-sided range is used instead of string equality because a
  // SharePoint Date column may serialize as either a date or midnight ISO
  // datetime. Paging is still followed, but only within that one day.
  ///////////////////////////////////////////////////////////////////////////////////////////////
  // Function Name: getPaceVisitsForDateByDisplayName
  // Description: Reads only one local calendar day's PACE visits from Graph, using a two-sided
  //              date-range filter, and converts field keys back to display names.
  // Parameters: string localDate - local calendar date (YYYY-MM-DD) to filter visits by - input
  ///////////////////////////////////////////////////////////////////////////////////////////////
  async getPaceVisitsForDateByDisplayName(localDate) {
    const date = String(localDate || "").slice(0, 10);
    if (!/^\d{4}-\d{2}-\d{2}$/.test(date)) throw new Error("A valid visit date is required.");

    const [siteId, listId, schema] = await Promise.all([
      this.getSiteId(),
      this.getListId(CONFIG.LISTS.paceVisits),
      this.getListSchema(CONFIG.LISTS.paceVisits)
    ]);
    const dateField = schema.Date;
    if (!dateField) throw new Error("IEP_Pace_Visits Date column was not found.");

    const [year, month, day] = date.split("-").map(Number);
    const nextDate = new Date(Date.UTC(year, month - 1, day + 1)).toISOString().slice(0, 10);
    const filter = `fields/${dateField} ge '${date}T00:00:00Z' and fields/${dateField} lt '${nextDate}T00:00:00Z'`;
    let path = `sites/${siteId}/lists/${listId}/items?$expand=fields&$filter=${encodeURIComponent(filter)}&$top=200`;
    const items = [];

    while (path) {
      const data = await this._get(path, { Prefer: "HonorNonIndexedQueriesWarningMayFailRandomly" });
      items.push(...(data.value || []));
      const nextLink = data["@odata.nextLink"];
      path = nextLink ? nextLink.replace(`${this._BASE}/`, "") : null;
    }

    const inverseSchema = {};
    Object.entries(schema).forEach(([display, internal]) => { inverseSchema[internal] = display; });
    return items.map(item => {
      const row = { id: item.id, Created: item.createdDateTime || item.fields?.Created || "" };
      Object.entries(item.fields || {}).forEach(([key, value]) => {
        row[inverseSchema[key] || key] = value;
      });
      return row;
    });
  },

  // Creates a new PACE visit.
  //
  // PATCH 003: reconciled against the LIVE IEP_Pace_Visits schema (see the
  // in-app diagnostic's report, and README "Known gaps" for the full
  // writeup). Two keys were flat-out wrong and are corrected here:
  //   "Behavior"     -> "Reason"            (live column is "Reason")
  //   "Interventions"-> "Intervention Used" (live column is "Intervention Used")
  // "Duration" is a real Number column, previously never sent — added.
  // The old "Entry ID" pre-save lookup below is REMOVED: "Entry ID" isn't
  // a real column on the live list, so that lookup threw every single
  // time (silently, via the .catch below it used to have) and cost a full
  // list fetch before every save for zero benefit. STATE.saving (app.js)
  // remains the functional duplicate-tap guard.
  //
  // Still sent despite no live column currently matching — kept because
  // mapFields() already drops unmapped keys harmlessly (a console.warn,
  // nothing more), and if a matching column is ever added to the list,
  // these start working with no code change: "Entry ID", "PACE Room",
  // "SCM Used", "Submitted By". "Submitted At" also has no live column,
  // but SharePoint's own system "Created" timestamp already covers that
  // same need automatically, so nothing is actually lost by leaving it.
  //
  // PATCH 004: added Behavior Specialist (STATE.staffMember) + Teacher to
  // the app's logical entry, but NEITHER is written to SharePoint here:
  //   - Teacher: no "Teacher" column exists on IEP_Pace_Visits at all
  //     (confirmed by Patch 003's live schema dump). Stays UI/state only,
  //     per instruction not to invent a column.
  //   - Behavior Specialist -> would-be "Staff Member": same personOrGroup
  //     column/limitation noted below for the authenticated user — the
  //     app has no infrastructure to resolve ANY plain name (specialist
  //     or signed-in user) to a SharePoint site-user id, so this is a
  //     second, independent confirmation of the same gap, not a new one.
  // NOT sent: "Staff Member" — it's a Person-type column and writing one
  // requires new infrastructure (resolving a name to a SharePoint
  // site-user id) this project doesn't have yet.
  //
  // PATCH 010: STATE.staffMember (single string) -> STATE.staffMembers
  // (array, multi-select) — "Behavior Specialist" now receives all
  // selected names, comma-joined, same convention as "Reason"/
  // "Intervention Used" just below. This is still the SAME speculative,
  // harmlessly-dropped-if-missing text column PATCH 004/005 already used
  // for a single name — multi-select does not change whether that column
  // exists; see README "Live IEP_Pace_Visits schema" for its confirmed
  // status. The Person-type "Staff Member" column remains unwritten for
  // exactly the same reason as before: still no site-user resolution
  // infrastructure, now doubly true for a list of names.
  //
  // PATCH 001: completed-visit model — Time Out is collected before save
  // and written here at creation (the existing "Time Out" field
  // closePaceVisit already used for the old open→close flow). A row
  // created by the current workflow is never "open."
  // CURRENT-PATCH: sends the human-readable label ("PACE Room 1"/"PACE
  // Room 2" — see config.js's paceRoomLabelForId()), not the raw internal
  // slug. Written here at CREATE time (Start Visit or Log Completed
  // Visit) — never deferred to completion.
  //
  // ROOM-FIELD-NAME PATCH: the confirmed live display name for this
  // column is "Room" (see config.js's PACE_ROOM_FIELD_CANDIDATES) — sent
  // under all three tolerated aliases with the identical value.
  // mapFields() silently drops whichever keys don't match a real column
  // (a console.warn, nothing more), so only the one real column actually
  // receives this write; the other two are no-ops.
  ///////////////////////////////////////////////////////////////////////////////////////////////
  // Function Name: savePaceVisit
  // Description: Creates a new PACE visit list item using the live display-name field mapping,
  //              sending both confirmed and speculative fields (harmlessly dropped if unmapped).
  // Parameters: Object entry - logical visit entry (room, student, times, reason, etc.) - input
  ///////////////////////////////////////////////////////////////////////////////////////////////
  async savePaceVisit(entry) {
    const roomValue = paceRoomLabelForId(entry.paceRoom);
    return this.createMappedListItem("IEP_Pace_Visits", {
      "Entry ID":           entry.id,
      "Room":               roomValue,
      "PACE Room":          roomValue,
      "Pace Room":          roomValue,
      "Student ID":         entry.studentId || "",
      "Student":            entry.studentName  || "",
      "Date":               entry.date         || "",
      "Time In":            entry.timeIn       || "",
      "Time Out":           entry.timeOut      || "",
      "Duration":           entry.durationMinutes ?? undefined,
      "Reason":             Array.isArray(entry.behaviors)     ? entry.behaviors.join(", ")     : (entry.behaviors || ""),
      "Intervention Used":  Array.isArray(entry.interventions) ? entry.interventions.join(", ") : (entry.interventions || ""),
      "SCM Used":           entry.scmUsed == null ? undefined : entry.scmUsed === true,
      "Notes":              entry.notes        || "",
      "Behavior Specialist": Array.isArray(entry.staffMembers) ? entry.staffMembers.join(", ") : (entry.staffMembers || ""),
      // PATCH 006: the teacher/classroom the student physically came from
      // immediately before this PACE visit — NOT the roster's homeroom
      // Teacher. Sent speculatively, same as the other fields above with
      // no confirmed live column yet: mapFields() drops it harmlessly (a
      // console.warn, nothing more) if "Teacher Came From" doesn't exist
      // on IEP_Pace_Visits yet, and starts working with zero code change
      // the moment that column is added. Manual SharePoint action may be
      // required — see README "Known gaps."
      "Teacher Came From":  entry.cameFromTeacher || "",
      "Submitted By":       entry.submittedByName || "",
      "Submitted At":       entry.timestamp    || new Date().toISOString()
    });
  },

  // PATCH 009: updates the existing SharePoint list item selected from
  // today's Recent Activity. Uses the same display-name mapping as CREATE,
  // so only confirmed writable columns are patched; missing logical fields
  // are dropped by mapFields() exactly as they are on initial save.
  ///////////////////////////////////////////////////////////////////////////////////////////////
  // Function Name: updatePaceVisit
  // Description: Applies a same-day correction patch to an existing PACE visit, sending only the
  //              fields present on `entry` so omitted fields are left untouched.
  // Parameters: string itemId - SharePoint list item id of the visit to patch - input
  //             Object entry - partial logical visit entry with only the changed fields - input
  ///////////////////////////////////////////////////////////////////////////////////////////////
  async updatePaceVisit(itemId, entry) {
    if (!itemId) throw new Error("A SharePoint visit item is required for editing.");
    // CURRENT-PATCH: label, not slug — see savePaceVisit() above. Same-day
    // edits can change the room via the Visit Info screen's room select.
    // ROOM-FIELD-NAME PATCH: all three tolerated aliases — see savePaceVisit().
    //
    // STAFF CORRECTIONS: a field whose key is ABSENT from `entry` is left
    // undefined here, and mapFields() drops undefined values — so a
    // correction that only changed Notes PATCHes only Notes and never
    // touches Time In / Time Out / Duration. A full entry (every key
    // present) produces exactly the same payload as before.
    const has = key => entry[key] !== undefined;
    const joined = value => Array.isArray(value) ? value.join(", ") : (value || "");
    const roomValue = has("paceRoom") ? paceRoomLabelForId(entry.paceRoom) : undefined;
    return this.updateMappedListItem("IEP_Pace_Visits", itemId, {
      "Room":              roomValue,
      "PACE Room":         roomValue,
      "Pace Room":         roomValue,
      "Student ID":        has("studentId") ? (entry.studentId || "") : undefined,
      "Student":           has("studentName") ? (entry.studentName || "") : undefined,
      "Date":              has("date") ? (entry.date || "") : undefined,
      "Time In":           has("timeIn") ? (entry.timeIn || "") : undefined,
      "Time Out":          has("timeOut") ? (entry.timeOut || "") : undefined,
      "Duration":          entry.durationMinutes ?? undefined,
      "Reason":            has("behaviors") ? joined(entry.behaviors) : undefined,
      "Intervention Used": has("interventions") ? joined(entry.interventions) : undefined,
      // CURRENT-PATCH fix: was `entry.scmUsed === true`, which silently
      // coerced a null/unknown SCM value to `false` — inconsistent with the
      // null-safe pattern already used by savePaceVisit()/completePaceVisit()
      // below, and directly against the "do not treat blank as No" rule.
      // In practice app.js's save-time guard no longer allows this path to
      // run with a non-boolean scmUsed, but this stays null-safe regardless.
      "SCM Used":          entry.scmUsed == null ? undefined : entry.scmUsed === true,
      "Notes":             has("notes") ? (entry.notes || "") : undefined,
      "Behavior Specialist": has("staffMembers") ? joined(entry.staffMembers) : undefined,
      "Teacher Came From": has("cameFromTeacher") ? (entry.cameFromTeacher || "") : undefined
    });
  },

  // STAFF CORRECTIONS: deletes exactly ONE IEP_Pace_Visits item, addressed
  // only by its SharePoint list-item id. Never by student, date, room, or
  // any search/bulk criteria. SharePoint list-item ids are integers, so
  // anything else (empty, a path fragment, a GUID) is refused before any
  // network call — an odd value can never widen the request URL.
  ///////////////////////////////////////////////////////////////////////////////////////////////
  // Function Name: deletePaceVisit
  // Description: Deletes exactly one PACE visit after validating that its item id is a numeric
  //              SharePoint list-item id.
  // Parameters: string itemId - SharePoint list item id of the visit to delete - input
  ///////////////////////////////////////////////////////////////////////////////////////////////
  async deletePaceVisit(itemId) {
    const id = String(itemId ?? "").trim();
    if (!/^\d+$/.test(id)) throw new Error("A valid SharePoint visit item is required for deletion.");
    const siteId = await this.getSiteId();
    const listId = await this.getListId(CONFIG.LISTS.paceVisits);
    await this._delete(`sites/${siteId}/lists/${listId}/items/${id}`);
    return { deleted: true, id };
  },

  // Legacy open-visit close path — Recent Activity corrections use the
  // separate full-field updatePaceVisit() method above.
  ///////////////////////////////////////////////////////////////////////////////////////////////
  // Function Name: closePaceVisit
  // Description: Retains the legacy minimal close operation for open visits, patching only the
  //              Time Out field.
  // Parameters: string itemId - SharePoint list item id of the visit to close - input
  //             string timeOut - time-out value to write - input
  ///////////////////////////////////////////////////////////////////////////////////////////////
  async closePaceVisit(itemId, timeOut) {
    return this.updateMappedListItem("IEP_Pace_Visits", itemId, { "Time Out": timeOut });
  },

  ///////////////////////////////////////////////////////////////////////////////////////////////
  // Function Name: completePaceVisit
  // Description: Completes an open visit by patching completion fields (time out, duration,
  //              interventions, SCM used, notes) on the same row.
  // Parameters: string itemId - SharePoint list item id of the visit to complete - input
  //             Object entry - logical completion fields to apply - input
  ///////////////////////////////////////////////////////////////////////////////////////////////
  async completePaceVisit(itemId, entry) {
    if (!itemId) throw new Error("A SharePoint visit item is required for completion.");
    return this.updateMappedListItem("IEP_Pace_Visits", itemId, {
      "Time Out":          entry.timeOut || "",
      "Duration":          entry.durationMinutes ?? undefined,
      "Intervention Used": Array.isArray(entry.interventions) ? entry.interventions.join(", ") : (entry.interventions || ""),
      "SCM Used":          entry.scmUsed == null ? undefined : entry.scmUsed === true,
      "Notes":             entry.notes || ""
    });
  }
};
