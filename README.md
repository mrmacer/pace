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
- **`PACE Room` values** — **CURRENT-PATCH: changed.** Now writes the
  human-readable label (`PACE Room 1` / `PACE Room 2`), not the raw
  `pace-room-1` / `pace-room-2` slug MAC Walkthrough uses internally — see
  "Known gaps" below for why. Every internal comparison in this app
  (`STATE.room`, `visit-workflow.js`, `recent-activity.js`) still compares
  the raw slug; the label conversion happens only at the two read/write
  boundary functions in `config.js` (`paceRoomLabelForId()` /
  `paceRoomIdForLabel()`).
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

### PATCH 011 — "Other" location (support outside a physical PACE room)

Staff reported that a lot of their day never touched a physical PACE room at
all — de-escalating a student in the Wiggle Room, a classroom check-in, a
walk — and none of that time was being captured, even though it's exactly
the kind of intervention parents and admin want documented. This patch adds
**Other** as a third entry on the Home screen's room picker
(`CONFIG.ROOMS`), alongside PACE Room 1/2, so that time gets logged through
the same Start Visit / Log Completed Visit / Mark Complete workflow as a
room visit — just not tied to a physical room.

**No SharePoint schema change needed.** `Room` is a plain text column (see
the schema table below); the app already writes whatever
`CONFIG.ROOMS[].label` is via `paceRoomLabelForId()`, so a third label
("Other") persists exactly the same way "PACE Room 1"/"PACE Room 2" do.

**Almost every consumer needed zero code changes.** `recent-activity.js`,
`visit-workflow.js`, the same-day-edit room `<select>`, and duplicate-open-
visit protection all already read `CONFIG.ROOMS` generically rather than
assuming exactly two rooms — confirmed by inspection before writing this
patch, the same way the ROOM-FIELD-NAME PATCH confirmed its own blast
radius. The one new flag, `isPhysicalRoom: false` on the `other` entry, is
read in exactly one place — `app.js`'s `currentRoomConfig()` — to swap a
handful of PACE-room-specific strings for wording that doesn't imply a
physical room:
- Room screen: "Student is entering PACE now" → "Support is starting now";
  "Currently in PACE" → "Currently receiving support"; the empty-state
  message and the Start Visit toast follow the same split.
- Notes screen: the placeholder becomes "Describe what happened and where
  (e.g., Wiggle Room, classroom check-in, a walk)…" — Notes was already
  required on every completed visit (see "PATCH 010" below), so this reuses
  that existing required field rather than adding a second one; it's the
  detail staff specifically asked for to help parents understand what
  happened.

Everything else about the workflow is identical to a room visit: Specialist
→ Student → Teacher Came From → Reason/Support chips → SCM → Notes →
Confirm, same duplicate-open-visit protection, same same-day edit/delete,
same Recent Activity scoping. Styled with a neutral gray theme
(`body[data-room="other"]`) deliberately distinct from both room colors, so
staff can't mistake it for a third physical room at a glance.

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
| **Room** | `Room` | text (Yes/No confirmed separately below for SCM) | **ROOM-FIELD-NAME PATCH: CONFIRMED EXISTING — manually verified by the user, not yet re-confirmed by the in-app diagnostic.** The live display name is `Room`, not `PACE Room` (the name every prior patch, including the diagnostic's own concept list, assumed). The app now writes `"PACE Room 1"`/`"PACE Room 2"` (label, not slug) under all three of `Room`/`PACE Room`/`Pace Room` — `mapFields()` keeps only the one that's real. Run the diagnostic again to get this row's type/internal-name confirmed the same authoritative way as the rest of this table. |
| SCM Used | *(unconfirmed exact name — user reports it as working)* | Yes/No | **ROOM-FIELD-NAME PATCH: CONFIRMED WORKING IN PRODUCTION** — the user reports SCM is already reaching SharePoint and appearing in the MAC Walkthrough admin report. This patch did not change SCM's write logic (still `entry.scmUsed == null ? undefined : entry.scmUsed === true`) or its required-before-completion validation (see the prior patch entry below) — only verified it wasn't regressed. Run the diagnostic to get the exact confirmed display/internal name and type into this table. |
| Submitted At | *(none)* | — | not needed — SharePoint's own system `Created` timestamp already covers this |
| Return Status | `Return Status` | choice | exists, intentionally unused — see below |

### Open gaps this app cannot close on its own

- **RESOLVED by the ROOM-FIELD-NAME PATCH — `PACE Room` was never missing,
  it was misnamed.** Every prior patch (including this one's own
  immediately-preceding version) assumed the room column would be called
  `PACE Room`, because that's what Patch 003's diagnostic was told to look
  for, run before this column existed at all. The user has since manually
  confirmed the live column's actual display name is **`Room`**. The app
  now writes the room label (`"PACE Room 1"` / `"PACE Room 2"`, via
  `paceRoomLabelForId()`) under all three of `Room`, `PACE Room`, and
  `Pace Room` in the same `mapFields()` call — whichever of the three is
  the real column keeps the write, the other two are silently dropped
  (harmless `console.warn`s) — so a future rename in either direction
  doesn't silently break persistence again. Written at visit **creation**
  time for both Start Visit and Log Completed Visit — never deferred to
  Mark Complete — and Mark Complete's own patch payload never re-sends any
  room-field alias, so the originally-recorded room is preserved
  automatically by SharePoint's partial-update semantics. On the read
  side, `pace-data.js`'s `readPaceRoomValue()` checks the same three
  aliases in order (a real visit read back from Graph is keyed by
  whatever `GRAPH.getListSchema()` says the live display name is —
  `Room`, confirmed) before `normalizeRoomOnRead()` converts the label
  back to the internal slug every other part of this app compares against.
  **Still not diagnostic-confirmed** — the user's manual confirmation is
  the only evidence so far; run the (now-updated) PATCH 003 diagnostic to
  get this into the schema table above the same authoritative way as every
  other row. Recent Activity's per-room scoping, "Currently in PACE," and
  same-day editing all depend on this and were re-verified working after
  the rename (see the "ROOM-FIELD-NAME PATCH" changelog entry below for
  the verification list) — no changes were needed in any of those files,
  since they all only ever read the already-normalized `"PACE Room"` key
  `PACE_DATA.getVisits()` produces, never a raw SharePoint field name.
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
- **RESOLVED — SCM is confirmed reaching SharePoint in production.** Prior
  patches (through the one immediately preceding this one) documented no
  SCM-related column as existing at all. The user has since confirmed SCM
  is already being written successfully and appearing in the MAC
  Walkthrough admin report — the exact live display/internal name and
  column type are still not diagnostic-confirmed (the schema table above
  shows this as the user's manual report, not a diagnostic result), but
  the write path itself is proven working and this patch made no changes
  to it. Unchanged from the prior patch: this is a *required* answer
  app-side regardless — a completed visit (Mark Complete or Log Completed
  Visit) cannot reach Save without an explicit Yes/No SCM answer (the SCM
  screen only advances via `setScm(true)`/`setScm(false)`, and `app.js`'s
  save handler independently re-checks `STATE.scmUsed` is strictly
  `true`/`false` before saving, redirecting back to the SCM screen
  otherwise). An **open** live-start visit is the one exception — it may
  still save with SCM unanswered, since staff often don't know yet — and
  `"SCM Used"` sends `undefined`/`null` in that case rather than false, so
  a blank cell is never misread as "No." `GRAPH.updatePaceVisit()` and the
  demo equivalent's earlier null-to-`false` coercion bug remains fixed.
  This patch only re-verified none of the above regressed — see the
  "ROOM-FIELD-NAME PATCH" changelog entry below.
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

### CURRENT PATCH — PACE Room + SCM required production data

**Room: label instead of slug, written at creation, preserved through
completion.** `"PACE Room"` now sends `entry.paceRoom` through
`config.js`'s new `paceRoomLabelForId()` before it reaches
`GRAPH.savePaceVisit()`/`updatePaceVisit()` (and the `demo-data.js`
equivalent), so the value is `"PACE Room 1"`/`"PACE Room 2"` — matching
`CONFIG.ROOMS[].label` exactly — not the internal `pace-room-1`/
`pace-room-2` slug. This was already being written at Start Visit
creation time (`saveLiveVisit()` → `PACE_DATA.createVisit()`), not deferred
to Mark Complete — that didn't need to change. Mark Complete's patch
payload (`GRAPH.completePaceVisit()`) still never includes `"PACE Room"`,
so the room recorded at Start Visit is preserved by SharePoint's own
partial-update semantics — that didn't need to change either. Log
Completed Visit already wrote `"PACE Room"` at creation via the same
`savePaceVisit()` path.

Every internal comparison — `visit-workflow.js`'s `openForRoom()`,
`recent-activity.js`'s room filter/`editModel()`, `app.js`'s room `<select>`
and `openExitScreen()`/`beginEditVisit()` — still compares the raw slug,
completely unchanged. The label only exists at the SharePoint boundary:
`pace-data.js`'s `getVisits()` now runs every visit through a new
`normalizeRoomOnRead()` (using `paceRoomIdForLabel()`) immediately after
reading, in both production and demo, converting the label back to a slug
before anything else in the app ever sees it. `paceRoomIdForLabel()` is a
no-op on an already-slug value, so this is also safe against legacy rows
saved before this patch.

**SCM: required before a visit can complete, in both directions.** The SCM
Yes/No screen was already structurally unbypassable for every non-live-start
flow (completed-entry, Mark Complete, same-day edit) — the only way off it
going forward is `setScm(true)`/`setScm(false)`, and there is no other route
to Notes/Confirm. This patch adds an explicit, defensive re-check in
`app.js`'s `saveEntryBtn` handler (the same pattern already used for the
date/time re-check just above it): if `STATE.scmUsed` isn't strictly `true`
or `false` when Save is tapped, the save is blocked and the flow is sent
back to the SCM screen. A **live-start** (Start Visit) save is unaffected —
it goes through `saveLiveVisit()`, a different function, and legitimately
saves with `scmUsed: null` for an open visit. Also fixed:
`GRAPH.updatePaceVisit()` and `pace-data.js`'s demo `updateVisit()` used to
send `entry.scmUsed === true`, silently coercing an unset SCM to `false` on
the same-day-edit path — both now use the same null-safe pattern
(`entry.scmUsed == null ? undefined/null : entry.scmUsed === true`) already
used by create/complete, so an edited row can never gain a fabricated "No."

**Superseded by the ROOM-FIELD-NAME PATCH immediately below:** this
section originally said neither `PACE Room` nor `SCM Used` was found on
the live list by the last diagnostic run (PATCH 003), and that every
change here was purely speculative/forward-compatible. The user has since
manually confirmed both are live — `SCM` under some name this app hasn't
independently confirmed, and room identity under the display name `Room`,
not `PACE Room`. The SCM-required validation and null-safety fixes
described above are unaffected and still accurate; only the room *field
name* assumption was wrong — see below for the fix. This patch did not
add, rename, or modify any SharePoint column, list, or permission.

### ROOM-FIELD-NAME PATCH — confirmed live column is `Room`, not `PACE Room`

The user manually confirmed the live `IEP_Pace_Visits` schema now contains
columns named `Room` and `SCM`, with SCM already reaching SharePoint and
appearing in the MAC Walkthrough admin report. The room column's real
display name is `Room` — every prior patch (including the one immediately
above) assumed `PACE Room`, because that's what Patch 003's diagnostic was
told to look for, run before this column existed.

**Compatibility over a hard rename.** Rather than replace `"PACE Room"`
with `"Room"` outright, the write side now sends the same room-label value
under all three of `Room`, `PACE Room`, and `Pace Room` in one
`mapFields()` call (`GRAPH.savePaceVisit()`/`updatePaceVisit()`,
`demo-data.js`'s `createVisit()`) — `config.js`'s new
`PACE_ROOM_FIELD_CANDIDATES` is the single source of truth for this list.
`mapFields()` already silently drops any key that isn't a real column (a
`console.warn`, nothing more — the same mechanism every speculative field
in this app already relies on), so only the one real column (`Room`,
today) actually receives the write; the other two cost nothing beyond
harmless console noise. If the column is ever renamed again in either
direction, persistence doesn't silently break a second time.

**Read side.** `GRAPH.getPaceVisitsByDisplayName()`/
`getPaceVisitsForDateByDisplayName()` already key every field by whatever
the *live* schema calls it (via `getListSchema()`), so a real production
row comes back keyed `Room`, not `PACE Room` — the app's own internal
`"PACE Room"` key (used everywhere in `visit-workflow.js`,
`recent-activity.js`, and `app.js`) would have silently stopped finding
any value the moment the column was actually confirmed. New
`pace-data.js` function `readPaceRoomValue()` checks the same three
candidates, in the same order, and `enrichVisitContext()`/
`normalizeRoomOnRead()` both now use it instead of a hardcoded
`visit["PACE Room"]` read. Because every consumer of a visit object only
ever reads the *already-normalized* `"PACE Room"` key these two functions
produce, **zero changes were needed in `app.js`, `visit-workflow.js`, or
`recent-activity.js`** — confirmed by inspection, not assumption, before
implementing.

**Demo parity.** `demo-data.js`'s `createVisit()` and `pace-data.js`'s
demo `updateVisit()` branch both renamed their stored key from
`"PACE Room"` to `"Room"`, matching the confirmed live name exactly — so
demo mode exercises the same alias-resolution path production does,
rather than continuing to hide the exact bug this patch fixes.

**Diagnostic tool.** `diagnostic.js`'s `DIAGNOSTIC_CONCEPTS` now includes
`"Room"` alongside the existing `"PACE Room"`/`"Pace Room"` entries, so a
re-run reports all three by name (FOUND for whichever is real, NOT FOUND
for the others) instead of only ever checking the name that turned out to
be wrong. The diagnostic remains strictly read-only.

**Verification (demo mode, in-browser, zero Graph calls):**
- Room 1 Start Visit → stored `Room: "PACE Room 1"` ✓
- Room 2 Start Visit → stored `Room: "PACE Room 2"` ✓
- Room 1 Log Completed Visit → stored `Room: "PACE Room 1"` ✓
- Room 2 Log Completed Visit → stored `Room: "PACE Room 2"` ✓
- Start Visit in Room 1 → Mark Complete → still `Room: "PACE Room 1"` ✓
- Start Visit in Room 2 → Mark Complete → still `Room: "PACE Room 2"` ✓
- SCM Yes and SCM No both still save correctly (unchanged by this patch) ✓
- An open visit created in Room 1, after a full page reload, still appears
  under Room 1's "Currently in PACE" (proves the read-side alias
  resolution round-trips correctly, not just the write) ✓
- Recent Activity's per-room scoping unaffected — reads the same
  already-normalized `"PACE Room"` key as before ✓
- No new console errors; zero `graph.microsoft.com` requests throughout ✓

**Older records.** A row saved before the `Room` column existed has no
value under any of the three candidates — `readPaceRoomValue()` correctly
returns `""` for it, exactly as before. Nothing infers, backfills, or
modifies a historical record's room from Behavior Specialist, Teacher, or
timestamp data; per instruction, older rows simply remain blank/unknown.

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

3. **Room and SCM: confirmed to already exist — no admin action needed for
   these two.** `Room` (room identity — see "ROOM-FIELD-NAME PATCH" below;
   the app writes to `Room`/`PACE Room`/`Pace Room` and only the real one
   keeps the value) and SCM (exact display name still unconfirmed by this
   app, but the user reports it working) are both confirmed live per the
   user, not this app's own tooling. **Still recommended:** run the PATCH
   003 diagnostic once, signed in to production, to get `Room` and SCM's
   exact display/internal names and types into this README with the same
   authority as every other confirmed row. **Still needed:** writable
   Single line of text columns named `Behavior Specialist` and
   `Teacher Came From` on `IEP_Pace_Visits` — the app schema-maps these
   display names automatically, no code change needed once they exist.
   Until then, this iPad keeps specialist/teacher context locally by
   SharePoint item id (never student data); room context no longer needs
   this local fallback now that `Room` is confirmed writable, though the
   fallback mechanism itself hasn't been removed (harmless — see
   `pace-data.js`'s `enrichVisitContext()`).
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
