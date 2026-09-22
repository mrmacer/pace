///////////////////////////////////////////////////////////////////////////////////////////////
// Author: R-E Miller & Greg Macer
// Creation Date: August 20, 2026
// Filename: auth.js
// Purpose: Encapsulates MSAL initialization, Entra redirect sign-in, silent Graph token
//          acquisition, IEP_Users2 staff verification, and sign-out cleanup for the PACE Room
//          Tracker kiosk app. Authentication proves Microsoft identity, while the active
//          IEP_Users2 lookup supplies the application's staff gate; demo mode must never
//          construct MSAL, process a redirect, or request a token. Reuses the same Entra app
//          registration already proven in MAC Walkthrough (same tenant, same "IEP Skool"
//          backend), computing redirectUri from window.location.origin so sign-in works from
//          whatever origin the kiosk iPad actually loads from. MANUAL CONFIG REQUIRED: an IU29
//          admin must add this app's deployed origin (and a localhost origin for local
//          testing) to the Entra app registration's "Single-page application" redirect URI
//          allow-list, or sign-in will fail with an AADSTS50011 redirect mismatch — this is
//          the same app registration MAC Walkthrough uses, so it is safe to extend, not
//          replace.
///////////////////////////////////////////////////////////////////////////////////////////////

const MSAL_CONFIG = {
  auth: {
    clientId: "145a3fc7-5cff-4d03-96c7-577e17980110",
    authority: "https://login.microsoftonline.com/3276761c-22db-462b-a930-172d155bd795",
    // REVIEW: keep redirect origins restricted in the Entra registration;
    // this value is intentionally derived from the current host.
    redirectUri: window.location.origin
  },
  cache: {
    // localStorage (not sessionStorage) so the MSAL session survives the
    // iPad sleeping/waking and the app reloading — section 22/23 of spec.
    cacheLocation: "localStorage",
    storeAuthStateInCookie: false
  }
};

// Authentication/session facade used by app.js and GRAPH. Tracks the MSAL client (_client),
// the signed-in Microsoft account, the resolved IEP_Users2 staff identity (staffId, staffName,
// role), and the latest staff verification failure (lookupError).
const AUTH = {
  _client:     null,
  account:     null,
  staffId:     null,
  staffName:   null,
  role:        null,
  lookupError: null,

  ///////////////////////////////////////////////////////////////////////////////////////////////
  // Function Name: init
  // Description: Initializes the MSAL client, restores any signed-in Microsoft account from a
  //              redirect or cache, and loads the corresponding staff record; does nothing in
  //              demo mode.
  // Parameters: none
  ///////////////////////////////////////////////////////////////////////////////////////////////
  async init() {
    // Demo mode never touches Microsoft auth at all — no MSAL client is
    // constructed, no redirect handling, no token/account state. app.js's
    // demo boot path doesn't even call this, but guard it here too so it's
    // safe by construction regardless of call site.
    if (typeof APP_MODE !== "undefined" && APP_MODE === "demo") {
      console.warn("Demo mode: skipping Microsoft authentication entirely.");
      return;
    }

    this._client = new msal.PublicClientApplication(MSAL_CONFIG);

    try {
      const response = await this._client.handleRedirectPromise();
      if (response?.account) {
        this.account = response.account;
      } else {
        const accounts = this._client.getAllAccounts();
        if (accounts.length > 0) this.account = accounts[0];
      }
    } catch (err) {
      console.error("MSAL error:", err);
    }

    if (this.account) {
      await this.loadStaffFromSharePoint();
    }
  },

  // Same gate MAC Walkthrough uses: a signed-in Microsoft account must also
  // resolve to an Active row in IEP_Users2 before the app is usable. This
  // keeps the append-only PACE workflow limited to approved IU29 staff
  // (spec section 19/22), reusing proven logic rather than a new scheme.
  ///////////////////////////////////////////////////////////////////////////////////////////////
  // Function Name: loadStaffFromSharePoint
  // Description: Resolves the signed-in account's email against IEP_Users2 via Graph and
  //              requires an Active staff row before granting access; sets staffId/staffName/
  //              role on success or lookupError on failure.
  // Parameters: none
  ///////////////////////////////////////////////////////////////////////////////////////////////
  async loadStaffFromSharePoint() {
    const email = (this.account.username || "").toLowerCase().trim();
    try {
      const user = await GRAPH.findUserByEmail(email);
      if (!user) {
        this.staffId = null; this.staffName = null; this.role = null;
        this.lookupError = "Your IEP Skook account was not found.";
        return null;
      }

      const staffId = user.Title || user.UserID || user.User_x0020_ID || null;
      const name    = user.field_1 || user.Name || user.Title || "";
      const rawRole = user.field_2 || user.Role || "";
      const role    = Array.isArray(rawRole) ? rawRole[0] : rawRole;
      const active  = user.field_3 ?? user.Active;
      const isActive = active === true || active === "Yes" || active === "true" || active === 1;

      if (!isActive) {
        this.staffId = null; this.staffName = null; this.role = null;
        this.lookupError = "Your IEP Skook account is inactive. Contact an administrator.";
        return null;
      }

      this.staffId    = staffId;
      this.staffName  = name || this.displayName;
      this.role       = role;
      this.lookupError = null;
      return { staffId, name: this.staffName, role };
    } catch (err) {
      console.error("Failed to load staff record from SharePoint:", err);
      this.lookupError = err.message || "Unable to verify your IEP Skook account.";
      return null;
    }
  },

  get isAuthenticated() { return this.account !== null && this.staffId !== null; },
  get isUnauthorized()  { return this.account !== null && this.staffId === null; },
  get displayName()     { return this.account?.name || this.account?.username || ""; },

  ///////////////////////////////////////////////////////////////////////////////////////////////
  // Function Name: acquireGraphToken
  // Description: Silently acquires a Microsoft Graph access token for the signed-in account,
  //              falling back to an interactive redirect if silent acquisition fails.
  // Parameters: none
  ///////////////////////////////////////////////////////////////////////////////////////////////
  async acquireGraphToken() {
    // Belt-and-suspenders: GRAPH's own assertGraphAllowed() already blocks
    // demo mode before this could ever be reached, but never request a
    // token in demo mode under any circumstance.
    if (typeof APP_MODE !== "undefined" && APP_MODE === "demo") {
      throw new Error("Demo mode safety block: token acquisition is disabled.");
    }
    const request = { scopes: ["User.Read", "Sites.ReadWrite.All"], account: this.account };
    try {
      const resp = await this._client.acquireTokenSilent(request);
      return resp.accessToken;
    } catch (err) {
      await this._client.acquireTokenRedirect(request);
      return null;
    }
  },

  ///////////////////////////////////////////////////////////////////////////////////////////////
  // Function Name: login
  // Description: Starts the Entra ID redirect sign-in flow requesting the app's required Graph
  //              scopes.
  // Parameters: none
  ///////////////////////////////////////////////////////////////////////////////////////////////
  login() {
    this._client.loginRedirect({ scopes: ["User.Read", "Sites.ReadWrite.All"] });
  },

  ///////////////////////////////////////////////////////////////////////////////////////////////
  // Function Name: _clearLocalSessionState
  // Description: Removes app-owned keys (last room, visit context) from localStorage while
  //              leaving unrelated stored settings untouched.
  // Parameters: none
  ///////////////////////////////////////////////////////////////////////////////////////////////
  _clearLocalSessionState() {
    if (typeof localStorage === "undefined" || typeof CONFIG === "undefined") return;
    const keys = CONFIG.STORAGE_KEYS || {};
    [keys.LAST_ROOM, keys.VISIT_CONTEXT]
      .filter(Boolean)
      .forEach(key => localStorage.removeItem(key));
  },

  ///////////////////////////////////////////////////////////////////////////////////////////////
  // Function Name: logout
  // Description: Clears in-memory authentication and staff state, removes app-owned local
  //              session keys, and redirects through MSAL sign-out; no-ops in demo mode.
  // Parameters: none
  ///////////////////////////////////////////////////////////////////////////////////////////////
  logout() {
    // Demo mode never constructs an MSAL client and must not clear its
    // simulated data or attempt a Microsoft redirect.
    if (typeof APP_MODE !== "undefined" && APP_MODE === "demo") return;

    const account = this.account;
    this.account = null;
    this.staffId = null;
    this.staffName = null;
    this.role = null;
    this.lookupError = null;
    this._clearLocalSessionState();

    if (!account || !this._client) return;
    return this._client.logoutRedirect({
      account,
      postLogoutRedirectUri: window.location.origin
    });
  }
};
