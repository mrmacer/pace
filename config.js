/* ─────────────────────────────────────────────────────────────────────────
   PACE Room Tracker — Configuration
   Standalone iPad kiosk app. Shares SharePoint backend with MAC Walkthrough
   but has its own tiny UI. See README.md for the architecture summary.

   Values below (site, list names, option lists, MSAL app registration) are
   copied from the existing, proven MAC Walkthrough project
   (/Projects/01-IU29/MAC-Walkthrough/config.js, auth.js, graph.js) rather
   than invented. Do not hand-guess SharePoint internal field names anywhere
   in this app — GRAPH.mapFields() always resolves display name → internal
   name live from the list schema.
   ───────────────────────────────────────────────────────────────────────── */

const CONFIG = {
  APP_NAME: "PACE Room Tracker",
  ORG_NAME: "IU29",

  // Same SharePoint site as MAC Walkthrough — PACE data is shared backend,
  // not a separate database (see project mission / section 14 of the spec).
  SITE: "siu29.sharepoint.com:/sites/IEP_Skook:",

  LISTS: {
    users:       "IEP_Users2",             // staff directory — Active + Role gate
    students:    "IEP_Students_2026_27",   // 2026–27 production roster
    paceVisits:  "IEP_Pace_Visits"         // shared with MAC Walkthrough / admin dashboard
  },

  // Fallback roster list name, only used if LISTS.students is ever reset to
  // a placeholder — mirrors the defensive fallback in MAC Walkthrough.
  STUDENT_ROSTER_FALLBACK: "IEP_Skook_Pilot_Students",

  // Local storage keys (device-local only — never student data at rest
  // beyond the current in-progress entry, per section 23 of the spec).
  STORAGE_KEYS: {
    LAST_ROOM: "paceTracker_lastRoom"
  },

  // Room identifiers. IMPORTANT: these raw slugs (not display labels) are
  // what MAC Walkthrough already writes into the "PACE Room" SharePoint
  // column (see its app.js: fd.get("paceRoom") is stored as-is, e.g.
  // "pace-room-1"). Keep using the same raw values so existing/future
  // admin reports keep working across both apps.
  ROOMS: [
    { id: "pace-room-1", label: "PACE Room 1" },
    { id: "pace-room-2", label: "PACE Room 2" }
  ],

  // Reused verbatim from MAC Walkthrough's PACE_BEHAVIOR_OPTIONS /
  // PACE_INTERVENTION_OPTIONS (config.js) so the two apps share one
  // taxonomy. "Other" intentionally omitted — PACE Room Tracker keeps
  // typing to an absolute minimum and this is a kiosk, not a report tool;
  // an admin can capture unusual cases in MAC Walkthrough directly.
  REASON_OPTIONS: [
    "Disruption", "Defiance / refusal", "Physical aggression", "Verbal aggression",
    "Elopement", "Unsafe behavior", "Peer conflict", "Property damage",
    "Transition difficulty", "Emotional dysregulation"
  ],

  SUPPORT_OPTIONS: [
    "De-escalation conversation", "Calm space", "Sensory support",
    "Restorative conversation", "Problem-solving conference", "Break / reset",
    "Check-in / check-out", "Parent/guardian contact", "Counselor support",
    "Admin support", "Modified task", "Return-to-class plan"
  ]
};
