# Agent session log

Newest first. One entry per working session or piece of work. Quote requests
word for word, with student information replaced by `[student]`. Never edit
an old entry; correct it in a new one. Rules: [README.md](README.md).

Entry template:

```
## MM-DD-YYYY — short title
**Agent:** <tool / model>, **machine:** <optional>
### Asked (verbatim)
> ...
### Decisions and why
### What changed (files)
### Verified how
### Open / follow-ups
### Commit text handed over
```

---

## 11-23-2026 (later still) — Full history, search by name, US formats

**Agent:** Claude Code (Claude Opus 5.5), Windows PC.

### Asked (verbatim)

> what is the one fail? is it working now? does it work as anticipated?

> also - it must have FULL history - not just 10 days. like eeven stuff from last school year

> they need to be able to search by NAME. state issued student IDs do NOT change between school years

> why am i seing stuff that says "last school year"? this list is just s list. it contains all students from all school years. student IDs and names dont change. i just want to see the date of the incident - i dont need any bullshit stating it was from last year

> ^ or was that justnotes

> you dont search by ID tho you search by student name..

> also we are in USA. the proper format is MM-DD-YYYY. change all formatting, language, tone to USA english and all numbers according... imperial, times, dates, etc

### Answers and decisions

- **The one failing check** (`tests/live-visit-notes.test.js`: "the sign-out
  handler must invalidate the active-note session") is a Windows line-ending
  problem in the test, not an app bug. The test searches `app.js` for the
  exact text `discardActiveNoteDrafts();\n  AUTH.logout();` with a Unix line
  break. With `core.autocrlf=true` the file has `\r\n`, so it can't match. The
  regex check just above it (any whitespace) passes. Proof: on a sandbox copy,
  changing only `app.js`'s line endings makes the test pass. Not changed:
  it's Greg's test.
- **"Last school year" text R-E saw** was the agent's fake demo notes and tour
  captions, not app labels. The app shows only each visit's date. The agent
  had added a "Past visits only" tag in the lookup; it was removed. All
  year wording was removed from the demo data and captions.
- **Correction:** the agent first claimed Student IDs change between school
  years and made the name the matching key. R-E corrected this: the roster is
  one list holding all students from all years, and IDs and names don't
  change. Final rule: staff search by **name**. After a name is picked, a
  visit matches by `Student ID` when both have one (keeps two same-named
  students apart), otherwise by the exact name (rows written by MAC
  Walkthrough have no ID).
- **Full history:** no date limit anywhere. Visit History returns every visit
  ever recorded for the student. The day picker goes back to any date. Today
  no longer stops at 10 visits; every day lists all of its visits. The
  search lists every roster student (any Active / PACE Enabled state), plus
  any other name found on visit rows (one read of only the `Student` column,
  cached 10 minutes), so every student with visits can be found by name.
- **US formats:** the app shows dates as MM-DD-YYYY ("11-23-2026",
  "Monday, 11-23-2026"), always, instead of following the device's language.
  Stored dates stay YYYY-MM-DD because SharePoint and the date picker require
  it. The page is marked `lang="en-US"`. Times were already 12-hour AM/PM. The
  app has no measurements, so nothing needed converting to imperial. The
  iPad's own date picker shows dates in the iPad's region format (US iPads:
  MM/DD/YYYY); the app can't change that picker. The agent notes now use
  MM-DD-YYYY too.

### What changed (files)

- `visit-history.js`: `belongsToStudent()` (ID when both have one, else exact
  name), `lookupStudents()`; `dayLimit()` uncapped.
- `graph.js`: `getPaceVisitStudentNames()`. `roster.js`: `getAll()`.
- `pace-data.js`: `getAllStudents()`, `getVisitStudentNames()`.
- `app.js`: the search uses the whole roster plus names found on visits;
  US date formatting in `fmtLongDate()` / `fmtWeekdayDate()`.
- `index.html`: `lang="en-US"`. `README.md`: STUDENT HISTORY PATCH section.
- Local kit: unit tests and tour updated (visits a year or more ago, one with
  the student's ID and one with none, a student found only by name on visits,
  11 visits today, a US date-format check).

### Verified how

- Visible tour, final run: `DONE: browser checks 136/136 passed; test files
  13/14 passed (failed: live-visit-notes.test.js)`. That one is the line-ending
  test problem above. The agent viewed the screenshots: US dates, history back
  400 days, a visit from a year ago opening with notes, name search with no
  extra labels, and 11 visits listed for today.
- Not yet verified: a signed-in production run on a real iPad.

---

## 11-23-2026 (later) — Browser test of the history feature

**Agent:** Claude Code (Claude Opus 5.5), Windows PC.

### Asked (verbatim)

> can we test this in browser to make sure it works

The agent started a local server and told R-E which links to click. R-E
corrected that:

> i need YOU to control the browser - we have done this before. it opens in my IDE not in my default browser. you run the tests.. remember this must be sandboxes

> dont mention my other projects in this project please

### Decisions and why

- The agent drives the browser itself: a new local-kit script,
  `~/.pace/demo/tour.js` (see `docs/LOCAL-KIT.md`). In one command it copies
  the repo into `~/.pace/sandbox/work`, runs every test file against that
  copy, then drives a **visible** Microsoft Edge window. The browser uses a
  fresh throwaway profile, runs in demo mode only, and the tour checks that no
  request goes to Microsoft. A caption shows what each step is doing. The
  earlier headless-only script was folded into it.
- The tour covers the whole feature like a staff member would:
  - Previous visits from a student in PACE
  - Recent Activity today, yesterday view-only, picking a date, an empty day,
    and a future date snapping back to today
  - Visit History from a card, with Details including multi-line notes
  - Student lookup, including a student with no visits
  - Waiting 32 s for the room's real auto-refresh (the bug fix)
  - A brand-new live visit started and completed, then seen in history
  - A visit logged after the fact for yesterday, then seen on that day
  - Today's edit still working
  - Phone width
- Rules added to CONTEXT.md: work only on `dev`; prove changes in the browser
  yourself; keep PACE self-contained (no mention of R-E's other projects).

### Findings while testing

- Early runs showed the app shifted sideways in some screenshots. Cause: the
  tour clicked or typed while a screen was still sliding in, or tapped through
  a toast. Focusing an element mid-slide scrolls the app's `overflow: hidden`
  container, and it stays shifted until reload. The tour now waits for slides
  and toasts to finish, like a person would, and checks for a sideways shift
  at every screenshot.
- Small pre-existing quirk, not changed: for about 2.6 s after Start Visit,
  the "… is now in PACE." toast sits over the bottom action button.
- The app itself could avoid the sideways shift entirely with
  `#app { overflow: clip; }`. Not changed; only worth doing if it's seen
  on an iPad.

### What changed (files)

- Repo: `docs/agent-notes/CONTEXT.md` (rules), this entry,
  `docs/LOCAL-KIT.md` (now documents `demo/tour.js`).
- Local kit only: `~/.pace/demo/tour.js` added,
  `~/.pace/tests/student-history.e2e.js` removed (folded into the tour).

### Verified how

- Visible tour, final run: `DONE: browser checks 117/117 passed; test files
  13/14 passed (failed: live-visit-notes.test.js)`. That one failure is the
  known pre-existing one. The agent viewed all 14 screenshots: no shifted
  screens, nothing covered, layout correct at iPad landscape and phone width.

---

## 11-23-2026 — View students' visits from previous days; start agent notes

**Agent:** Claude Code (Claude Opus 5.5), Windows PC.

### Asked (verbatim)

> Kelly, Robyn, Carl, Bruce, Nicole and Amber suggested:
> Be able to view previous days for a student.. add maybe a student search? Or - if the student was seen and has records today - make something in that students profile that says "previous visits"? Or in the recents, give an option to choose what day? Or both maybe?
>
> go through the app to make sure you fully understand. use sandboxes if you must. the BIs downstairs want to be able to view students visits from previous days. make that happen. thanks

Mid-task follow-up:

> also -= please make agent noters that save chats to the gitrhub, save context, etc. i want to note EVERYTHING that is said so any new agents on new machines know the full scope and context

The agent pointed out that `mrmacer/pace` is a **public** repo and asked where
the notes should live. Options offered: a separate private notes repo, curated
notes in this repo, or a verbatim auto-logged transcript in this repo.
**R-E chose: "In pace repo, curated only"**, meaning agent-written notes in
`docs/agent-notes/`, verbatim requests, never student names, no raw
transcripts.

### Decisions and why

- **Built all three suggestions**, since the BIs said "or both maybe":
  1. a day picker on Recent Activity (Prev day / date / Next day, never in
     the future, "Back to today");
  2. "Look up a student's visit history" (roster search) → a new Visit
     History screen;
  3. "Visit history" on every Recent card and "Previous visits" on every
     Currently in PACE card.
- **Past days are read-only.** Edit/Delete stay today-only, matching the
  existing rule ("For an older entry, contact an administrator") and
  `PACE_DATA.updateVisit()`'s own guard.
- **Today unchanged:** still the 10 newest, with Edit + Delete. Past days list
  every completed visit that day. Same room scoping as today.
- **Privacy:** Recent cards keep the existing allow-list with no notes on any
  day. Full details including notes appear only in Visit History, after a
  deliberate "Details" tap, and the details markup is built only on that tap.
  This matches the existing pattern (notes visible only after deliberately
  entering the editor). The BIs need the notes to understand past visits,
  which is why notes are reachable at all.
- **Production reads stay bounded:** a new Graph read filters on the confirmed
  `Student` column server-side, so it doesn't download the whole list. The
  result is then narrowed by `Student ID` when available, so same-named
  students stay apart; rows without an id fall back to an exact name match.
  The paging and mapping code is shared with the existing per-date read via
  `GRAPH._getPaceVisitsWhere()`.
- **Fixed a real bug found while reading:** `loadRecentActivity()` overwrote
  `STATE.paceVisits` (the room's open-visit cache) with completed-only rows.
  After Recent → Back, the room's 30-second refresh dropped students still in
  PACE and discarded their unsaved live notes. Removed that overwrite.
- **Agent notes:** `docs/agent-notes/` (this LOG, CONTEXT, README), a root
  `CLAUDE.md` and `AGENTS.md` that point every agent here, and
  `docs/LOCAL-KIT.md` (R-E's standing rule: document how to rebuild the
  local `~/.pace/` kit).

### What changed (files)

- New `visit-history.js`: pure rules (date stepping/clamping, student
  matching, newest-first grouping, details model).
- `graph.js`: `getPaceVisitsForStudentByDisplayName()` plus the shared
  `_getPaceVisitsWhere()`. The per-date read's behavior is unchanged.
- `pace-data.js`: `getStudentVisits()`.
- `app.js`: day picker, Visit History, roster lookup, Details panel, entry
  buttons, stale-response guards, and the `STATE.paceVisits` bug fix.
- `index.html`: day bar and lookup button on Recent; new `historySearch` and
  `history` screens; `<script src="visit-history.js">`.
- `styles.css`: styles for the above. `sw.js`: cache `v22`, added
  `visit-history.js`.
- `README.md`: "STUDENT HISTORY PATCH" section; Recent description and file
  map updated.
- New docs: `docs/agent-notes/*`, `CLAUDE.md`, `AGENTS.md`,
  `docs/LOCAL-KIT.md`.

### Verified how

- Repo's existing checks run against a copy in `~/.pace/sandbox/work`: all
  unchanged. One check, `tests/live-visit-notes.test.js` (sign-out
  assertion), already failed on the untouched `main` checkout before any
  change.
- New local-only checks in `~/.pace/` (see `docs/LOCAL-KIT.md`): pure rules,
  Graph filter escaping (`O'Brien`) and paging, demo/production provider
  behavior, read-only structure; and a headless Microsoft Edge run in demo
  mode clicking through every new path at 1024×768 and 390×844. Zero
  Microsoft requests, no console errors.
- **Not verified:** a signed-in production run against real SharePoint.

### Open / follow-ups

- Production check on an iPad: day picker, Visit History, Details.
- Ask an IU29 admin to index `Date` and `Student` on `IEP_Pace_Visits`.
- Investigate the pre-existing `live-visit-notes` failure.

---

## Before 11-23-2026 — history reconstructed from git (no notes were kept)

- 08-20-2026 → 08-31 (Greg): initial app, demo mode, production SharePoint
  activation, temporary diagnostic, specialists/teachers, student search +
  Teacher Came From (PATCH 005–006), PATCH 007–010 (teacher-list cleanup,
  room-scoped Recent, multiple specialists, longer notes), PACE room + SCM
  required, Room field-name fix, IEP_App_Users authorization, sign-out.
- 09-15-2026 → 09-22: Notes required on completed visits; full roster on
  Select Student; same-day edit/delete (STAFF CORRECTIONS); "Other"
  location (PATCH 011); bug fixes for GitHub issues #5 and #8.
- 10-09-2026: duplicate-save fix ("Duplicate Save Risk #2", retry
  idempotency); minimum Notes length (later removed).
- 10-21-2026: issue #10, hard 20-character Notes minimum removed (now just
  "not blank").
- 11-04-2026 (Greg, PR #12): live visit notes (notes before start + autosave
  while in PACE); diagnostic made metadata-only.
