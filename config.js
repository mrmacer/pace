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

// ── App mode ────────────────────────────────────────────────────────────
// "demo"       — no Microsoft sign-in, no Graph calls, simulated roster,
//                records saved only to this browser's localStorage. Safe
//                to deploy publicly for hands-on testing.
// "production" — the real thing: MSAL sign-in, live Graph calls, writes to
//                IEP_Pace_Visits. See auth.js/graph.js for the guards that
//                key off this flag.
//
// PATCH 002: resolved once at load, from the URL, instead of a hardcoded
// literal — but every downstream `APP_MODE === "demo"` check (auth.js,
// graph.js, pace-data.js, app.js) is completely unchanged, since this still
// produces one plain "demo" | "production" string in the same constant.
// Rule, in order:
//   1. ?demo=1        forces demo, on ANY host — including the production
//      URL itself, so it stays possible to demo the app publicly without
//      ever touching real IU29 data.
//   2. ?production=1  forces production, for local/manual verification
//      against the real tenant from a dev server.
//   3. Otherwise: production ONLY on the known deployed production
//      hostname; everything else (localhost, a Vercel preview URL, a
//      not-yet-listed custom domain) defaults to demo. Fails toward "no
//      access to real data," never the other way.
function resolveAppMode() {
  const params = new URLSearchParams(window.location.search);
  if (params.has("demo")) return "demo";
  if (params.has("production")) return "production";
  const PRODUCTION_HOSTNAMES = ["pace-room-tracker.vercel.app"];
  return PRODUCTION_HOSTNAMES.includes(window.location.hostname) ? "production" : "demo";
}
const APP_MODE = resolveAppMode(); // "demo" | "production"

const DEMO_CONFIG = {
  enabled: APP_MODE === "demo",
  storageKey: "paceRoomTrackerDemoData",
  roomStorageKey: "paceRoomTrackerDemoRoom"
};

const CONFIG = {
  APP_NAME: "PACE Room Tracker",
  ORG_NAME: "IU29",

  // Same SharePoint site as MAC Walkthrough — PACE data is shared backend,
  // not a separate database (see project mission / section 14 of the spec).
  SITE: "siu29.sharepoint.com:/sites/IEP_Skook:",

  LISTS: {
    users:       "IEP_Users2",             // staff directory — Active + Role gate
    students:    "IEP_Students_2026_27",   // 2026–27 production roster
    paceVisits:  "IEP_Pace_Visits",        // shared with MAC Walkthrough / admin dashboard
    appUsers:    "IEP_App_Users"           // PATCH A: new cross-app permission registry — see iep-app-users.js
  },

  // Fallback roster list name, only used if LISTS.students is ever reset to
  // a placeholder — mirrors the defensive fallback in MAC Walkthrough.
  STUDENT_ROSTER_FALLBACK: "IEP_Skook_Pilot_Students",

  // Local storage keys (device-local only — never student data at rest
  // beyond the current in-progress entry, per section 23 of the spec).
  STORAGE_KEYS: {
    LAST_ROOM: "paceTracker_lastRoom",
    VISIT_CONTEXT: "paceTracker_visitContext"
  },

  // Room identifiers. `id` (a raw slug) is what every internal comparison
  // in this app uses — STATE.room, STATE.editingRoom, visit-workflow.js's
  // openForRoom(), recent-activity.js's room filter/editModel(), and every
  // room <select>/CSS hook. `label` is the human-readable string.
  //
  // CURRENT-PATCH: `label` (e.g. "PACE Room 1") is what actually gets
  // written into the existing SharePoint "PACE Room" display column — NOT
  // the raw slug. Conversion happens in exactly two places, both far from
  // this array: graph.js's savePaceVisit()/updatePaceVisit() convert
  // id -> label on the way out (paceRoomLabelForId(), below), and
  // pace-data.js's getVisits() converts label -> id on the way back in
  // (paceRoomIdForLabel()) so every internal comparison keeps working on
  // slugs unchanged. demo-data.js mirrors the same id -> label conversion
  // for parity. hallway/color remain presentation-only, never sent anywhere.
  //
  // PATCH 011: added "other" — a third, non-physical-room location for
  // support that never touches a PACE room (Wiggle Room, a classroom
  // check-in, a walk, etc.), requested by staff whose day often wasn't
  // being captured at all. Room is a plain text SharePoint column, so a
  // third label needs no SharePoint schema change. `isPhysicalRoom: false`
  // is the one flag app.js reads to swap "PACE room"-specific copy (action
  // captions, the Currently-in-PACE heading/empty-state, the start-visit
  // toast, the Notes placeholder) for wording that doesn't imply a
  // physical room — see app.js's currentRoomConfig(). Every other consumer
  // (recent-activity.js, visit-workflow.js, the room <select>, duplicate-
  // open-visit protection) already reads CONFIG.ROOMS generically and
  // needed zero changes for a third entry.
  ROOMS: [
    { id: "pace-room-1", label: "PACE Room 1", hallway: "Yellow Hall", color: "yellow", isPhysicalRoom: true },
    { id: "pace-room-2", label: "PACE Room 2", hallway: "Green Hall",  color: "green",  isPhysicalRoom: true },
    { id: "other", label: "Other", hallway: "Outside PACE Room", color: "gray", isPhysicalRoom: false }
  ],

  // Reused verbatim from MAC Walkthrough's PACE_BEHAVIOR_OPTIONS /
  // PACE_INTERVENTION_OPTIONS (config.js) so the two apps share one
  // taxonomy. "Other" intentionally omitted — PACE Room Tracker keeps
  // typing to an absolute minimum and this is a kiosk, not a report tool;
  // an admin can capture unusual cases in MAC Walkthrough directly.
  // PATCH 006: added "Needs a Break" and "Other" — all prior entries left
  // exactly as they were (same wording, same order), nothing renamed or
  // removed. "Other" is a plain chip like every other option here; there's
  // no free-text follow-up mechanism on the Reason screen today, and
  // adding one for just this one value isn't warranted (see README).
  REASON_OPTIONS: [
    "Disruption", "Defiance / refusal", "Physical aggression", "Verbal aggression",
    "Elopement", "Unsafe behavior", "Peer conflict", "Property damage",
    "Transition difficulty", "Emotional dysregulation",
    "Needs a Break", "Other"
  ],

  SUPPORT_OPTIONS: [
    "De-escalation conversation", "Calm space", "Sensory support",
    "Restorative conversation", "Problem-solving conference", "Break / reset",
    "Check-in / check-out", "Parent/guardian contact", "Counselor support",
    "Admin support", "Modified task", "Return-to-class plan"
  ],

  // PATCH 004 introduced this as the live picklist. PATCH 005 replaced
  // that: production now loads Active "Behavior Specialist" rows from
  // IEP_Users2 dynamically (see pace-data.js's getSpecialists()) — this
  // array is now ONLY a TEMPORARY fallback for when that live call fails
  // or returns nothing (e.g. before the SharePoint rows exist yet).
  // TODO: remove this once dynamic loading is confirmed against real
  // production data — see README "Known gaps."
  BEHAVIOR_SPECIALISTS: [
    "Sharon Morgan", "Carl Stine", "Robyn Seiler", "Kelly Higgins",
    "Emma Brady", "Katrina Quinn", "Patrick Denmon", "Amber Clews",
    "Kelly Marchetti", "Bruce Andruchek", "Nicole Williams", "Luke Prescott"
  ]
};

// CURRENT-PATCH: the room slug <-> label boundary conversion referenced in
// the ROOMS comment above. Kept here (not pace-data.js) since both
// graph.js and demo-data.js are plain <script> globals loaded before
// pace-data.js, and this is pure config-driven lookup with no APP_MODE
// branching of its own.
function paceRoomLabelForId(id) {
  const room = CONFIG.ROOMS.find(r => r.id === id);
  return room ? room.label : (id || "");
}
function paceRoomIdForLabel(value) {
  const raw = String(value ?? "").trim();
  if (!raw) return "";
  const room = CONFIG.ROOMS.find(r => r.label === raw || r.id === raw);
  // Unrecognized values (e.g. a legacy row, or a manual SharePoint edit)
  // are passed through as-is rather than silently dropped — never guess,
  // never discard data this app doesn't understand.
  return room ? room.id : raw;
}

// ROOM-FIELD-NAME PATCH: the user manually confirmed the live
// IEP_Pace_Visits column for room identity is actually named "Room" —
// not "PACE Room", the name the prior patch (and Patch 003's diagnostic,
// run before this column existed) assumed. Every place this app reads a
// room value off a raw visit object tries these display names in order;
// "Room" is the one confirmed live today, "PACE Room"/"Pace Room" stay as
// tolerated aliases so a future rename either direction doesn't silently
// break persistence again. graph.js's/demo-data.js's write side sends all
// three (mapFields()/DemoStorage drop whichever key isn't real — see
// README "Known gaps" — so this costs nothing beyond a harmless
// console.warn per unmatched alias).
const PACE_ROOM_FIELD_CANDIDATES = ["Room", "PACE Room", "Pace Room"];
