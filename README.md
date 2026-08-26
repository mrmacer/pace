# PACE Room Tracker — IU29

Standalone, kiosk-style iPad web app for logging PACE room visits. Built to
live permanently on a PACE room iPad: **tap → tap → tap → save**, no
scrolling forms. It is a separate project from MAC Walkthrough but writes
into the same SharePoint backend so PACE data stays visible to
administrators through MAC Walkthrough's dashboard/reports.

```
PACE ROOM iPad → PACE Room Tracker → Microsoft Graph → IEP_Pace_Visits → MAC Walkthrough
```

PATCH 008 exposes two intentionally separate entry paths:

```
START VISIT:
Room → Specialist → Student → Teacher Came From → Reason → Time In → Start

LOG COMPLETED VISIT:
Room → Specialist → Student → Teacher Came From → Date/Times → Reason →
Support → SCM → Notes → Confirm → Save

MARK COMPLETE (an existing open row):
Currently in PACE → Time Out → Support → SCM → Notes → Confirm → Complete
```

Starting creates one open `IEP_Pace_Visits` row with completion fields
blank. Mark Complete patches that same item id; it never creates a second
row. Completed-entry saves still return to Student Search for rapid
back-to-back logging. A live start or completion returns to the room so
the Currently in PACE list updates immediately.

**PATCH 004 introduced a homeroom-Teacher-grouped Student screen (Room →
Specialist → Teacher → filtered Student list) — PATCH 006 removed it.** The
roster's `Teacher` column is homeroom/roster ownership, which turned out
not to reliably represent which classroom a student was *physically*
coming from at the moment of a PACE visit, so grouping student selection by
it was the wrong model. Student selection is now search-first across the
full eligible roster, and a new, separate "Teacher Came From" question
(right after picking the student, stored as `STATE.cameFromTeacher`, never
conflated with the roster's homeroom `student.teacher`) captures that
instead. `getTeacherList()`/`renderTeacherGrid()`/`studentsForTeacher()`/
`selectTeacher()` and the old Teacher screen's markup were deleted outright
rather than kept-but-hidden (unlike "Currently in PACE" in Patch 001) —
they were fully superseded by the new Teacher-Came-From screen's similar
but functionally different logic, so keeping both would have meant two
confusing "teacher grid" concepts side by side. The roster still reads and
keeps `student.teacher`/`student.classroom` internally (harmless, may be
useful elsewhere later) — neither is displayed anywhere in the PACE
workflow.

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
- **`IEP_Pace_Visits` field mapping** — *initially* copied from MAC
  Walkthrough's `GRAPH.savePaceVisit` display names, but **PATCH 003
  discovered the live list actually in use is not identical** to what MAC
  Walkthrough's source implied — see "Known gaps" below for the corrected,
  schema-verified mapping. Lesson: MAC Walkthrough's *source code* was
  inspected for this, never the *live schema* directly, until Patch 003's
  in-app diagnostic actually queried it.
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

Local/preview URLs default to demo mode; `?demo=1` also forces demo on any
host. Demo requires **no Microsoft sign-in and makes zero calls to
Microsoft Graph, MSAL, or SharePoint**.

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
  `.getVisits()` / `.getRecentVisits()` / `.createVisit()` / `.completeVisit()` /
  `.updateVisit()` and never touches
  `GRAPH`, `ROSTER`, or `DemoStorage` directly.
- **Recent Activity (PATCH 008)** — opening the existing screen, or tapping
  its small **Refresh** button, makes a fresh provider read for the iPad's
  local **today** and shows at most the 10 newest completed visits. In
  production, the `IEP_Pace_Visits` Date filter is applied in the Graph
  list request, so historical rows are never downloaded for this view.
  In demo, the same provider method reads and filters localStorage and
  makes no Microsoft request. If today's rows contain `PACE Room` values,
  the selected room is enforced. If all room values are blank (the current
  production schema cannot persist that logical field), the list falls
  back to today's completed visits across rooms instead of showing a false
  empty state. Cards are an allow-list only: student display name, Time
  Out, duration, reason, optional Behavior Specialist, and room only when
  the result is not already room-scoped. Missing-Time-Out legacy/open rows
  are excluded. No notes, IDs, emails, raw timestamps, history, exports,
  student filtering, or analytics are rendered.
- **Same-day editing (PATCH 009)** — every Recent card has an explicit
  **Edit today's visit** action. It preloads the existing Specialist →
  Student → Teacher Came From → Visit Details → Reason → Support → SCM →
  Notes → Confirm workflow, patches the original row, then returns to a
  freshly loaded Recent list. The date is locked to local today in the UI,
  and `PACE_DATA.updateVisit()` independently rejects a non-today payload.
  Cards remain privacy-limited; notes and the other full details appear
  only after staff deliberately enter the editor. Demo updates the same
  localStorage row with zero Microsoft traffic. Production patches every
  writable live `IEP_Pace_Visits` field; the known missing columns and
  Person-field specialist limitation below still apply.
- **Live visits + same-row completion (PATCH 008)** — demo creates an open
  localStorage row (`Time Out`, Duration, Intervention, SCM, and Notes are
  blank), renders it in Currently in PACE for its room, and later updates
  that exact row through `PACE_DATA.completeVisit()`. The completed row
  disappears from Currently in PACE and becomes eligible for Recent
  Activity immediately. Stable entry IDs plus the UI saving lock prevent
  rapid double-tap duplicates.
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
- **`SIMULATED` badge** — shown on demo rows in Currently in PACE; the
  persistent demo banner identifies Recent Activity without adding a
  non-operational field to its privacy-limited cards.
- **Reset Demo Data** — bottom of the Recent screen, demo-only, confirms
  before calling `DemoStorage.reset()`.
- **Automated verification** covers open creation, refresh persistence,
  room separation, duplicate-entry-id protection, same-id completion,
  transition from Currently in PACE to Recent, and zero Microsoft calls.

The known production hostname selects production automatically. Use
`?production=1` only for an intentional local production test after the
manual configuration below is complete.

**Deployment**: this directory is linked to the Vercel project
`pace-room-tracker`. Deploy manually with `vercel deploy --prod` after
local verification; code changes are not deployed automatically.

## Known gaps / decisions made without further data model changes

### Live `IEP_Pace_Visits` schema (verified by PATCH 003's in-app diagnostic — authoritative)

| Concept | Live column | Type | Status |
|---|---|---|---|
| Student | `Student` | text | correct, unchanged |
| Date | `Date` | **dateTime** (not text — see caveat below) | correct, unchanged |
| Time In / Time Out | `Time In` / `Time Out` | text | correct, unchanged |
| Duration | `Duration` | number | **fixed in PATCH 003** — was never sent; now sends whole minutes from `visitDurationMinutes()` |
| Reason/Behavior | `Reason` | choice | **fixed in PATCH 003** — app was sending `"Behavior"`, which doesn't exist |
| Interventions/Support | `Intervention Used` | choice | **fixed in PATCH 003** — app was sending `"Interventions"`, which doesn't exist |
| Notes | `Notes` | text | **PATCH 010: Single line vs. Multiple lines still unconfirmed — see below** |
| Staff identity | `Staff Member` | **personOrGroup** | **not written — see below** |
| Behavior Specialist(s) | `Behavior Specialist` | *(unconfirmed — sent speculatively)* | **PATCH 010: now one or more, comma-joined — see below** |
| Entry ID (dedupe key) | *(none)* | — | **not a real column — see below** |
| **PACE Room** | *(none)* | — | **not a real column — see below, most significant gap** |
| SCM | *(none)* | — | **not a real column — see below** |
| Submitted At | *(none)* | — | not needed — SharePoint's own system `Created` timestamp already covers this |
| Return Status | `Return Status` | choice | exists, intentionally unused — see below |

### Open gaps this app cannot close on its own

- **`PACE Room` has no matching live column at all.** This is the most
  significant open gap: every production visit this app has ever written
  has no way to say which room it came from. The app still sends
  `"PACE Room": entry.paceRoom` (harmless — silently dropped if unmapped,
  and starts working immediately with zero code change the moment a
  matching column exists), but **this needs a decision**: either an admin
  adds a room-identifying column, or there's an existing mechanism for
  this the app hasn't been told about. One direct, currently-live
  consequence: Recent Activity cannot reliably enforce per-room scoping
  for real production rows. PATCH 008 detects that all room values are
  blank and falls back to today's visits across rooms instead of showing
  a false empty state; adding the column restores true room scoping and
  makes room edits persist.
- **`Staff Member` is a Person field (`personOrGroup`), not text.**
  Writing to it requires resolving the signed-in user to a SharePoint
  site-user id first (a separate Graph call this project has no
  infrastructure for) and a different payload shape
  (`StaffMemberLookupId`, not a plain string) — `mapFields()` only does
  flat display-name→internal-name value assignment, nothing Person-field
  aware. Per instruction, this was reported rather than improvised. The
  app still sends `"Submitted By"` (also not a real column — dropped
  harmlessly) for the same reason: **no staff identity is currently
  captured anywhere in production SharePoint.**
- **PATCH 004 confirms the same `Staff Member` gap a second, independent
  way**: the selected Behavior Specialist (`STATE.staffMember`) has the
  identical problem — no infrastructure to resolve a plain name to a
  SharePoint site-user id. **PATCH 005** made the specialist *list itself*
  dynamic (loaded from `IEP_Users2`, `Active` + `Role = "Behavior
  Specialist"`, via `GRAPH.getActiveUsersByRole()`/
  `PACE_DATA.getSpecialists()`) and confirmed authorization was never the
  blocker here — `AUTH.isAuthenticated` only checks `Active`, never
  `Role`, so any Active Behavior Specialist row already signs in
  successfully with zero auth code changes. `CONFIG.BEHAVIOR_SPECIALISTS`
  (the old hardcoded 12) is now a **temporary fallback only**, used if the
  live `IEP_Users2` query fails or returns zero rows — remove it (and the
  fallback branch in `pace-data.js`'s `getSpecialists()`) once dynamic
  loading is confirmed against real IEP_Users2 data with actual Behavior
  Specialist rows added. Still true regardless: **no `Teacher` column exists on
  `IEP_Pace_Visits` at all.** Both Specialist and Teacher are fully
  functional app state (drive the new Specialist → Teacher → Student
  screens, shown on Confirm) but are **not** written to SharePoint —
  intentionally, per instruction not to invent a column or hack around a
  Person field. Demo mode still stores both (`"Staff Member"`, `"Teacher"`)
  for parity/completeness, same as the other not-yet-mappable fields above.
- **No SCM-related column exists in the live list at all** — not `SCM
  Used`, not `SCM`, nothing. The app's "Was SCM Required?" screen answer
  is not persisted in production today. Still sent (`"SCM Used"`) for the
  same forward-compatibility reason as above.
- **`Entry ID` is not a real column.** The old SharePoint-side duplicate
  lookup (`findListItemByDisplayField(..., "Entry ID", ...)`) was
  therefore throwing on *every single save*, silently swallowed by a
  `.catch`, after fetching the *entire* `IEP_Pace_Visits` list first for
  no benefit. **Removed in PATCH 003.** `STATE.saving` (`app.js`) remains
  the actual duplicate-tap guard, exactly as it already was in practice.
- **`Reason`/`Intervention Used` are Choice-type columns**, and this kiosk
  lets staff select *multiple* reasons/supports (joined with `", "` into
  one string, unchanged behavior from before). Whether these Choice
  columns allow multiple selections wasn't captured by the diagnostic's
  schema query — if either is single-select-only, a save with more than
  one reason/support chosen may be rejected by Graph. Watch for this on
  the next manual production test.
- **`Date` is a true `dateTime` column**, not text. The app writes a bare
  `"YYYY-MM-DD"` string (as before) — Graph/SharePoint accepts this
  (confirmed populated on the test row) and this app's own reads are
  unaffected (`dateOnly()` already normalizes defensively). The only
  caveat: SharePoint's *own* native list views may display the stored
  value with a timezone-shifted time-of-day artifact (common for
  date-only values in a DateTime column) — cosmetic in SharePoint's UI
  only, not a data-correctness issue for this app.
- **`Return Status`** exists (`choice`) but this kiosk's flow has no
  return-status question — never written, left blank on every row. An
  admin can fill it in via MAC Walkthrough or SharePoint directly.
- **Device identity** (spec §31, e.g. "PACE Room 1 iPad") — no SharePoint
  column exists for this either. Inferred from the selected room and kept
  local-only (remembers last room per device), never written to
  SharePoint.
- **Pagination**: `GRAPH.getListItems` follows `@odata.nextLink` (MAC
  Walkthrough's version does not) — `IEP_Pace_Visits` will grow past
  Graph's ~200-item page size within a school year, and "Currently in
  PACE" must never silently miss a recent entry.
- **PATCH 006: `"Teacher Came From"` has no confirmed live column yet.**
  Sent speculatively (same pattern as `PACE Room`/`SCM Used`/etc. above) as
  a plain string — harmless if the column doesn't exist, starts working
  immediately with zero code change once it does. **Manual SharePoint
  action recommended**: add a column to `IEP_Pace_Visits` named exactly
  `Teacher Came From`, type **Single line of text**. Not created
  automatically, per instruction. Run the Patch 003 diagnostic (still in
  place) after adding it to confirm the exact internal name Graph assigns.
- **`PACE_DATA.getTeachers()`** (the "Teacher Came From" picklist) —
  **`IEP_Users2` is the sole authoritative source** for this, same as it
  already is for Behavior Specialists: Active rows with `Role = "Teacher"`,
  `Name` field sorted alphabetically, via the same `GRAPH.
  getActiveUsersByRole()` the Specialist picker uses. **The student
  roster's `Teacher` column is deliberately NOT used to build this list.**
  Patch 006 originally combined both sources; **Patch 007 removed the
  roster side** after production data showed `IEP_Users2` and the
  roster's `Teacher` column use different naming conventions for the same
  person (e.g. `"Andruchek"` vs. `"Andruchek, B."`), which surfaced as
  apparent duplicates. No fuzzy or surname-based merge is attempted
  anywhere in this codebase — two people can legitimately share a
  surname — so reconciling those naming conventions is a data-governance
  decision for IU29, not something this app should guess at. If the
  `IEP_Users2` query fails or returns zero rows, the picker shows **no**
  named teachers rather than silently falling back to the roster (that
  fallback is exactly what would reintroduce the problem) — "Other / Not
  Listed" (with manual entry) is always available regardless, so the
  workflow is never blocked. **To add a selectable teacher, no deployment
  is needed**: add a row to `IEP_Users2` with `Role = Teacher`,
  `Active = Yes`, a `Name`, and any placeholder `User ID` — `Email` can
  stay blank if this person is only ever selected as a PACE origin
  teacher. They appear automatically the next time someone opens the
  Teacher Came From screen.

### PATCH 010 — multiple Behavior Specialists + expanded Notes

**Multiple Behavior Specialists.** The specialist screen now supports
selecting one or more people (`STATE.staffMembers`, an array — replaces the
old single-value `STATE.staffMember`). Tapping a card toggles it; nothing
navigates forward until the new CONTINUE button is tapped (previously,
tapping the one allowed specialist advanced immediately). Screen title is
now "Select Behavior Specialist(s)"; "Behavior Specialist" terminology was
kept rather than switching to "Behavior Interventionist" — it's already
consistent with everything else in this codebase (`IEP_Users2`'s
`Role = "Behavior Specialist"`, `CONFIG.BEHAVIOR_SPECIALISTS`,
`getSpecialists()`, the diagnostic's concept list) and with the live
`Role` value itself, which nothing here should drift out of sync with.

Multiple names are joined with `", "` and split back with the same
`selections()` helper already used for `Reason`/`Intervention Used` (see
`recent-activity.js`'s `specialistNames()`) — this is the exact same
established multi-value-in-a-text-column convention this codebase already
used for those two fields, not a new pattern. Written to the **same**
speculative `Behavior Specialist` text column PATCH 004/005 already used
for a single name (see the schema table above) — multi-select doesn't
change whether that column exists on the live list; it was already
recommended, still not confirmed. The `Staff Member` Person field remains
unwritten for the same reason as always (no site-user resolution
infrastructure) — now doubly true for a list of names, not a new gap.

Display: Confirm and Mark Complete show every name on its own line under a
"Behavior Specialist" (or, for 2+, "Behavior Specialists") label. Recent
Activity and Currently in PACE — both space-constrained — show up to 2 names
in full and "First Name +N" beyond that (`RECENT_ACTIVITY.formatSpecialistsCompact()`).
Mark Complete never re-asks for specialists: they're carried over read-only
from the open visit (`openExitScreen()`), exactly as `completeVisit()`'s
SharePoint patch already never touched that field.

Older single-specialist records still render correctly — `specialistNames()`
on a plain one-name string just returns a one-element array.

**Expanded Notes.** The actual limitation was **entirely frontend**: a
`maxlength="500"` attribute on `#noteText` in `index.html`, with no
corresponding limit anywhere else — no JS truncation (`STATE.notes` is only
ever `.trim()`'d, never sliced) and no truncation in `DemoStorage`/`GRAPH`'s
payload builders (both pass `entry.notes` straight through). That attribute
is now removed; the textarea (still `rows="6"`, unchanged) accepts natural
multi-paragraph typing with no arbitrary cap, and no character counter was
added (there is no confirmed *meaningful* limit left to display — see next
paragraph).

**Notes column type — still unconfirmed, don't guess.** The PATCH 003
diagnostic confirmed `Notes` exists and is Graph-`text`-typed, but Graph
represents both SharePoint "Single line of text" (hard 255-character cap)
and "Multiple lines of text" (no such cap) as the same `text` facet — the
diagnostic's original type classification could not, and still cannot from
existing data, tell them apart. **This patch extends the diagnostic**
(`diagnostic.js`) to also capture `text.allowMultipleLines`/`text.maxLength`
for every text column, so running it once against the real tenant will give
a conclusive answer next time — until then, treat notes over roughly 250
characters as *at risk* of being silently truncated by SharePoint on save if
the live column does turn out to be Single line. Two ways forward, neither
performed automatically:
- Run the (now-enhanced) diagnostic while signed in to production — instant,
  read-only, authoritative answer.
- Or just convert `Notes` to **Multiple lines of text** in SharePoint list
  settings regardless — a safe, purely-additive change (existing values are
  preserved; nothing currently reading a "Single line" `Notes` value breaks)
  that removes the question either way.

### Temporary diagnostic tool

PATCH 003 added a production-only, read-only "Run SharePoint Diagnostic"
button (Home screen) — see `diagnostic.js`, marked `TEMPORARY PATCH 003
DIAGNOSTIC` everywhere it touches `index.html`/`app.js`/`sw.js`. Remove all
four marked spots once the mapping above is confirmed stable in
production and no longer needed.

## Manual configuration required before first sign-in

1. **Entra redirect URI** — an IU29 admin must add this app's deployed
   origin (e.g. `https://pace-room-tracker.vercel.app`) — and
   `http://localhost:*` for local testing — to the **same** Entra app
   registration MAC Walkthrough uses (Single-page application platform,
   clientId `145a3fc7-5cff-4d03-96c7-577e17980110`). Without this,
   sign-in fails with `AADSTS50011`.
2. Confirm the signed-in test account has an **Active** row in
   `IEP_Users2` and the roster list has students flagged **PACE Enabled**.

3. **Recommended before multi-iPad live use:** add writable Single line of
   text columns named `PACE Room`, `Behavior Specialist`, and
   `Teacher Came From`, plus a Yes/No column named `SCM Used`, to
   `IEP_Pace_Visits`. The app schema-maps these display names automatically.
   Until then, this iPad keeps only room/specialist/teacher context locally
   by SharePoint item id (never student data), so its own newly started
   visits stay room-scoped after refresh. Another iPad cannot infer the room
   for an open row whose SharePoint record has no room value.
   **PATCH 010 note:** `Behavior Specialist` now receives a comma-joined
   list of one or more names — a couple of names comfortably fits Single
   line of text's 255-character cap, but if you expect a PACE room to
   regularly have several specialists logged on one visit, use **Multiple
   lines of text** instead for headroom; either works with zero code
   changes on this end.
4. **Recommended, low-effort:** convert the existing `Notes` column to
   **Multiple lines of text** if it isn't already (unconfirmed either way —
   see "PATCH 010" below; the enhanced diagnostic will tell you for
   certain). Safe, additive, and removes any risk of a long note being
   silently cut off at save time.

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
- [x] Phase 6 — Currently in PACE + full same-row completion workflow
- [x] Phase 7 — Recent entries (read-only, today, selected-room preference)
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
recent-activity.js — pure Recent selection, sorting, room scope, card allow-list
visit-workflow.js — pure workflow modes, duration validation, open-room selection
pace-data.js  — PACE_DATA: the one place demo vs. production is decided
app.js        — screen navigation + all panel logic/state (mode-agnostic —
                 talks only to PACE_DATA)
manifest.webmanifest, sw.js — PWA install + offline app-shell caching
```
