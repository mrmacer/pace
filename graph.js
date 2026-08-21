/* ─────────────────────────────────────────────────────────────────────────
   PACE Room Tracker — Microsoft Graph client

   Deliberately narrow: this app only ever touches the three lists in
   CONFIG.LISTS (users, students, paceVisits) — see spec section 29, "Data
   Separation." No Daily Pulse / Walkthrough / Check-In list is referenced
   anywhere in this file.

   Field-mapping pattern is copied from MAC Walkthrough's graph.js: callers
   always write by SharePoint *display* name; mapFields() resolves the
   internal name live from the list's column schema. Internal names are
   never hard-coded, per the project's explicit instruction not to invent
   them.
   ───────────────────────────────────────────────────────────────────────── */

// Hard safety block (demo mode): every Graph call funnels through _get/
// _post/_patch below, so guarding those three is a single choke point that
// covers every higher-level method transitively (getSiteId, getListId,
// findUserByEmail, savePaceVisit, closePaceVisit, everything) — no call
// site needs its own check, and none can be added later that forgets one.
function assertGraphAllowed() {
  if (typeof APP_MODE !== "undefined" && APP_MODE === "demo") {
    throw new Error("Demo mode safety block: Microsoft Graph access is disabled.");
  }
}

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

const GRAPH = {
  _BASE:        "https://graph.microsoft.com/v1.0",
  _SITE:        CONFIG.SITE,
  _siteId:      null,
  _listIdCache: {},
  _schemaCache: {},

  async _get(path) {
    assertGraphAllowed();
    const token = await AUTH.acquireGraphToken();
    const resp  = await fetch(`${this._BASE}/${path}`, { headers: { Authorization: `Bearer ${token}` } });
    if (!resp.ok) {
      const body = await resp.text().catch(() => "");
      console.error("Graph GET failed", path, resp.status, body);
      throw new Error(`Graph ${resp.status}: ${body}`);
    }
    return resp.json();
  },

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
      console.error("Graph POST failed", path, resp.status, text);
      throw new Error(`Graph POST ${resp.status}: ${text}`);
    }
    return resp.json();
  },

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
      console.error("Graph PATCH failed", { path, status: resp.status, response: text });
      throw new Error(`Graph PATCH ${resp.status}: ${text}`);
    }
    if (resp.status === 204) return null;
    return resp.json();
  },

  async getSiteId() {
    if (this._siteId) return this._siteId;
    const data = await this._get(`sites/${this._SITE}`);
    this._siteId = data.id;
    return data.id;
  },

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

  async createListItem(listName, fields) {
    const siteId = await this.getSiteId();
    const listId = await this.getListId(listName);
    return this._post(`sites/${siteId}/lists/${listId}/items`, { fields });
  },

  async createMappedListItem(listName, displayFields) {
    const fields = await this.mapFields(listName, displayFields);
    return this.createListItem(listName, fields);
  },

  async updateMappedListItem(listName, itemId, displayFields) {
    const siteId = await this.getSiteId();
    const listId = await this.getListId(listName);
    const fields = await this.mapFields(listName, displayFields);
    return this._patch(`sites/${siteId}/lists/${listId}/items/${itemId}/fields`, fields);
  },

  async findListItemByDisplayField(listName, displayFieldName, value) {
    const [items, schema] = await Promise.all([this.getListItems(listName), this.getListSchema(listName)]);
    const internalName = schema[displayFieldName];
    if (!internalName) throw new Error(`Column not found in ${listName}: ${displayFieldName}`);
    return items.find(item => String(item[internalName] || "").trim() === String(value || "").trim()) || null;
  },

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

  /* ── PACE-specific helpers ─────────────────────────────────────────────── */

  // Every list item, mapped back to DISPLAY field names via the schema, so
  // the rest of the app never touches internal SharePoint names.
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

  // Creates a new PACE visit. entry.id is a stable per-submission id
  // (reused across a retry after a failed sync) — used as the duplicate-
  // detection key, same pattern as MAC Walkthrough's savePaceVisit/
  // saveWalkthrough. "Return Status" is intentionally not written — that
  // question isn't part of this app's entry flow (see README "Known gap").
  //
  // PATCH 001: completed-visit model — Time Out is now collected before
  // save and written here at creation (same existing "Time Out" field
  // closePaceVisit already used for the old open→close flow; no new field
  // introduced). A row created by the current workflow is never "open."
  async savePaceVisit(entry) {
    const existing = await this.findListItemByDisplayField("IEP_Pace_Visits", "Entry ID", entry.id).catch(() => null);
    if (existing) {
      console.warn("PACE visit already exists in SharePoint:", entry.id);
      return { duplicatePrevented: true, existingItem: existing };
    }
    return this.createMappedListItem("IEP_Pace_Visits", {
      "Entry ID":      entry.id,
      "PACE Room":     entry.paceRoom     || "",
      "Student":       entry.studentName  || "",
      "Date":          entry.date         || "",
      "Time In":       entry.timeIn       || "",
      "Time Out":      entry.timeOut      || "",
      "Behavior":      Array.isArray(entry.behaviors)     ? entry.behaviors.join(", ")     : (entry.behaviors || ""),
      "Interventions": Array.isArray(entry.interventions) ? entry.interventions.join(", ") : (entry.interventions || ""),
      "SCM Used":      entry.scmUsed === true,
      "Notes":         entry.notes        || "",
      "Submitted By":  entry.submittedByName || "",
      "Submitted At":  entry.timestamp    || new Date().toISOString()
    });
  },

  // Closes an open visit — the ONLY field this app is ever allowed to patch
  // on an existing record (spec section 19/21).
  async closePaceVisit(itemId, timeOut) {
    return this.updateMappedListItem("IEP_Pace_Visits", itemId, { "Time Out": timeOut });
  }
};
