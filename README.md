# PACE Room Tracker — IU29

Standalone, kiosk-style iPad web app for logging PACE room visits. Built to
live permanently on a PACE room iPad: **tap → tap → tap → save**, no
scrolling forms. It is a separate project from MAC Walkthrough but writes
into the same SharePoint backend so PACE data stays visible to
administrators through MAC Walkthrough's dashboard/reports.

```
PACE ROOM iPad → PACE Room Tracker → Microsoft Graph → IEP_Pace_Visits → MAC Walkthrough
```

## What was reused from MAC Walkthrough

Inspected at `~/Projects/01-IU29/MAC-Walkthrough` before writing any code,
per this project's build instructions. Reused directly (not reinvented):

- **MSAL / Entra config** (`auth.js`) — same tenant, same app registration
  (clientId `145a3fc7-…`), same scopes (`User.Read`, `Sites.ReadWrite.All`).
- **Graph client pattern** (`graph.js`) — resolve site → list → **column
  schema by display name**, never hard-coded internal names. Copied
  verbatim from MAC Walkthrough's `GRAPH.mapFields()`/`getListSchema()`.
- **SharePoint site**: `siu29.sharepoint.com:/sites/IEP_Skook:`
- **Lists used** (and *only* these three — see "Data separation" below):
  - `IEP_Users2` — staff directory, gates sign-in (must be Active).
  - `IEP_Students_2026_27` — student roster (`CONFIG.LISTS.students`),
    filtered to `Active` + `PACE Enabled`.
  - `IEP_Pace_Visits` — the shared PACE record list.
- **`IEP_Pace_Visits` field mapping** (display names, exactly as MAC
  Walkthrough's `GRAPH.savePaceVisit` writes them): `Entry ID`, `PACE
  Room`, `Student`, `Date`, `Time In`, `Time Out`, `Behavior`,
  `Interventions`, `SCM Used`, `Notes`, `Submitted By`, `Submitted At`.
- **`PACE Room` values** — kept as the same raw slugs MAC Walkthrough
  already writes (`pace-room-1` / `pace-room-2`), not display labels, so
  existing/future admin reports keep working across both apps.
- **Behavior / Intervention taxonomy** — copied from MAC Walkthrough's
  `PACE_BEHAVIOR_OPTIONS` / `PACE_INTERVENTION_OPTIONS` (`Other` dropped —
  this kiosk minimizes typing; unusual cases can still be captured in MAC
  Walkthrough).
- **Duplicate-submission protection** — same pattern: a stable per-entry
  `Entry ID` (UUID) generated once and reused across a retry, checked
  against the list before create.
- **IU29 branding** — logo copied from `assets/iu29-logo.png`; brand blue
  (`#2563eb`) kept, everything else simplified for a kiosk.

Nothing else was copied — no Daily Pulse, Walkthrough, Student Check-In,
Reports, Dashboard, Form Lab, Setup, or sidebar nav. This app has no
navigation menu at all; the workflow itself *is* the navigation.

## Demo mode

`APP_MODE` in `config.js` is currently `"demo"` — the deployed/testing
build requires **no Microsoft sign-in and makes zero calls to Microsoft
Graph, MSAL, or SharePoint**. It's safe to share the URL publicly for
hands-on testing.

- **Fake roster** (`demo-data.js`, `DEMO_STUDENTS`) — six `SIM-00N` /
  first-name-plus-initial students, no real IU29 names.
- **Local-only storage** (`DemoStorage` in `demo-data.js`) — PACE visits
  are saved to `localStorage` under `paceRoomTrackerDemoData`, keyed by
  the exact same SharePoint *display* field names production writes
  (`"PACE Room"`, `"Student"`, `"Time In"`, …), so every render function in
  `app.js` works identically in both modes without a single `if (APP_MODE)`
  inside the UI code itself.
- **Unified adapter** (`pace-data.js`, `PACE_DATA`) — the only place that
  branches on `APP_MODE`. `app.js` calls `PACE_DATA.getStudents()` /
  `.getVisits()` / `.createVisit()` / `.closeVisit()` and never touches
  `GRAPH`, `ROSTER`, or `DemoStorage` directly.
- **Hard safety block** — `GRAPH._get/_post/_patch` (the only three
  functions that ever call `fetch()` against Microsoft endpoints) throw
  immediately in demo mode via `assertGraphAllowed()`, so this is a single
  choke point that covers every higher-level Graph method transitively,
  not a per-call-site check that a future call could forget.
  `AUTH.init()`/`AUTH.acquireGraphToken()` are guarded the same way, so
  MSAL is never constructed and no token is ever requested.
- **Demo banner** — persistent, fixed top strip (`#demoBanner`), shown only
  when `APP_MODE === "demo"`; every screen's own top offset shifts down to
  clear it (see `body.demo-mode .screen` in `styles.css`).
- **`SIMULATED` badge** — shown on any visit row with `demo: true` (Currently
  in PACE + Recent), so nothing reads as a real record.
- **Reset Demo Data** — bottom of the Recent screen, demo-only, confirms
  before calling `DemoStorage.reset()`.
- **Verified**: room→student→reason→support→SCM→notes→confirm→save→Currently
  in PACE→reload-persists→Exit→Time-Out-updates→Recent→Reset, all exercised
  in-browser with zero network requests recorded throughout (confirmed via
  the browser's network panel — nothing to `login.microsoftonline.com`,
  `graph.microsoft.com`, or `siu29.sharepoint.com`).

**To switch to the real thing:** set `APP_MODE = "production"` in
`config.js` and complete "Manual configuration required" below first. No
production auth/Graph code was removed or altered to build demo mode —
only guarded — so this is a one-line flip, not a rebuild.

**Deploying the demo publicly**: this project isn't yet linked to a Vercel
project (unlike MAC Walkthrough). Deploying makes the app reachable at a
real public URL, so that step wasn't run automatically — see the
"Deploying" section below and run `vercel deploy --prod` (or `vercel link`
first, if this is the first deploy) when ready.

## Known gaps / decisions made without further data model changes

- **`Return Status`** exists as a column on `IEP_Pace_Visits` (MAC
  Walkthrough's form asks it) but this kiosk's entry flow does not include
  that question — the spec's Reason → Support → SCM → Notes flow has no
  return-status step. This app never writes that field; it's left blank on
  rows it creates. An admin can still fill it in later via MAC Walkthrough
  or SharePoint directly.
- **Device identity** (spec §31, e.g. "PACE Room 1 iPad") — no SharePoint
  column exists for this on `IEP_Pace_Visits` today, and the project
  instructions say not to invent internal field names. The room iPad's
  identity is inferred from which room the staff member picked (stored
  locally per-device for "remember last room"), but it is **not** written
  to SharePoint. If per-device provenance is wanted later, an admin should
  add a `Device` (single line text) column to `IEP_Pace_Visits` and it can
  be wired in with one line in `graph.js`.
- **Pagination**: added `@odata.nextLink` following to `GRAPH.getListItems`
  (MAC Walkthrough's version does not page) — `IEP_Pace_Visits` will grow
  past Graph's ~200-item page size within a school year, and "Currently in
  PACE" must never silently miss a recent entry.

## Manual configuration required before first sign-in

1. **Entra redirect URI** — an IU29 admin must add this app's deployed
   origin (e.g. `https://pace-room-tracker.vercel.app`) — and
   `http://localhost:*` for local testing — to the **same** Entra app
   registration MAC Walkthrough uses (Single-page application platform,
   clientId `145a3fc7-5cff-4d03-96c7-577e17980110`). Without this,
   sign-in fails with `AADSTS50011`.
2. Confirm the signed-in test account has an **Active** row in
   `IEP_Users2` and the roster list has students flagged **PACE Enabled**.

No other manual SharePoint changes are required — this app only reads/
writes the same list MAC Walkthrough already uses successfully.

## Running locally

Static site, no build step (same as MAC Walkthrough):

```bash
npx serve .
```

Then open the printed `http://localhost:...` URL on an iPad (or Safari
responsive mode) and add the localhost origin to the Entra redirect URI
allow-list for testing.

## Deploying

Any static host works. To use Vercel (as MAC Walkthrough does):

```bash
vercel deploy
```

No `vercel.json` is needed — this is a plain static app like MAC
Walkthrough's `index.html`/`app.js` bundle.

## Phase status (per project build plan)

- [x] Phase 1 — Architecture inspection + SharePoint connection wiring
- [x] Phase 2 — Room-selection shell + horizontal panel navigation
- [x] Phase 3 — Student selection + automatic Time In
- [x] Phase 4 — Reason / support / SCM workflow
- [x] Phase 5 — SharePoint CREATE (`GRAPH.savePaceVisit`)
- [x] Phase 6 — Currently in PACE + EXIT workflow (Time Out PATCH only)
- [x] Phase 7 — Recent entries (read-only, today, this room)
- [x] Phase 8 — iPad/PWA optimization (manifest, standalone display, app
      shell service worker, safe-area insets, 44px+ touch targets)
- [ ] Phase 9 — Security/permissions verification — **needs a live sign-in
      test** against the real tenant (cannot be exercised headlessly); see
      "Manual configuration required" above before attempting it.

## File map

```
index.html    — all screens as static markup, one <section class="screen">
                 per step; app.js shows/hides + slides between them
styles.css    — kiosk styling, big tap targets, slide-transition CSS
config.js     — APP_MODE/DEMO_CONFIG, site/list names, room ids, options
demo-data.js  — fake roster, fake staff identity, localStorage-backed store
auth.js       — MSAL init/login/logout, IEP_Users2 staff gate (demo-guarded)
graph.js      — Graph client: schema-mapped CRUD, scoped to 3 lists only
                 (demo-guarded via assertGraphAllowed())
roster.js     — loads + filters the PACE-enabled student roster
pace-data.js  — PACE_DATA: the one place demo vs. production is decided
app.js        — screen navigation + all panel logic/state (mode-agnostic —
                 talks only to PACE_DATA)
manifest.webmanifest, sw.js — PWA install + offline app-shell caching
```
