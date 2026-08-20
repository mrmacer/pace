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
config.js     — site/list names, room ids, reason/support option lists
auth.js       — MSAL init/login/logout, IEP_Users2 staff gate
graph.js      — Graph client: schema-mapped CRUD, scoped to 3 lists only
roster.js     — loads + filters the PACE-enabled student roster
app.js        — screen navigation + all panel logic/state
manifest.webmanifest, sw.js — PWA install + offline app-shell caching
```
