# PACE Room Tracker — context for agents

_Last updated: 11-23-2026. Keep this current; history belongs in [LOG.md](LOG.md)._

## What this is

A kiosk-style iPad web app used by IU29's PACE rooms (Skook) to log student
support visits: **tap → tap → tap → save**. Plain static HTML/CSS/JS with no
build step and no framework. It writes to the same SharePoint backend as the
separate **MAC Walkthrough** app (`IEP_Pace_Visits` list, via Microsoft Graph),
so administrators see PACE data in MAC Walkthrough's reports.

- **Users:** Behavior Specialists / Behavior Interventionists ("BIs") who
  staff PACE Room 1 (Yellow Hall), PACE Room 2 (Green Hall), and "Other"
  (support outside a PACE room).
- **Deployed:** Vercel project `pace-room-tracker`, deployed by hand
  (`vercel deploy --prod`). Pushing does not deploy.
- **Modes:** demo (default for local/preview hosts or `?demo=1`; fake data in
  localStorage, zero Microsoft traffic) and production (the deployed hostname
  or `?production=1`). Only `pace-data.js` branches on the mode.

## People

- **R-E Miller** (GitHub `R-E-Miller`): owner, works mainly through Claude Code
  on Windows. Commits and pushes personally using GitHub Desktop.
- **Greg Macer** (GitHub `mrmacer`, owns the `mrmacer/pace` remote): co-author.
  Recent work arrives as PRs made with Codex.
- The BIs who request features (first names appear in LOG entries as quoted).

## Standing rules from the owner (apply to every agent)

- **Never run `git commit` or `git push`.** When work is done, give R-E a
  commit title and description to paste into GitHub Desktop.
- **Work on `dev`, never `main`.** Before changing anything, confirm the branch
  is `dev` and that it already contains `origin/main`
  (`git merge-base --is-ancestor origin/main HEAD`). R-E merges `dev` into
  `main` personally.
- **Prove changes in a browser yourself.** Don't hand R-E a link to click
  through (links open inside their IDE). Run the visible browser tour from the
  local kit (see [../LOCAL-KIT.md](../LOCAL-KIT.md)): it tests a sandbox copy
  in demo mode with a throwaway Edge profile. Look at the screenshots, then
  report the results.
- **Keep PACE self-contained.** Don't mention R-E's other projects anywhere in
  this repo, its notes, or replies about it.
- **The repo holds only the project.** Scratch tools, sandboxes, and local
  verification kits live outside the repo in `~/.pace/`. Ask before adding any
  file to the repo that isn't part of the project. The one sanctioned
  exception is [../LOCAL-KIT.md](../LOCAL-KIT.md), which says how to rebuild
  `~/.pace/` on a new PC. Keep it current whenever the kit changes.
- **Keep these agent notes current** (see [README.md](README.md)): add a LOG
  entry at the end of every piece of work.
- **Never put student information** in the repo, notes, commits, or PR text.
- **US conventions everywhere** (the district is in the USA): dates shown as
  MM-DD-YYYY, 12-hour times with AM/PM, US English spelling and tone, imperial
  units. This applies in the app, the notes, and replies. Stored dates stay
  YYYY-MM-DD because SharePoint and the date picker require it.
- **No extra labels the staff didn't ask for.** For example, don't tag
  visits or students by school year; show the date.

## How the code is organized

See the "File map" at the end of the root `README.md`. In short:
`app.js` (UI controller, one `STATE` object, screens slide horizontally),
`pace-data.js` (the only demo/production switch), `graph.js` (Graph client,
fields addressed by SharePoint *display* name and resolved from the live
schema), and pure, DOM-free rule modules: `recent-activity.js`,
`visit-history.js`, `visit-workflow.js`, `visit-corrections.js`,
`student-select.js`.

Conventions worth keeping:
- Every file and function starts with the bordered `// Function Name: /
  Description: / Parameters:` header block.
- Changes are labelled as named patches in comments and in the README
  (e.g. `PATCH 011`, `STAFF CORRECTIONS`, `STUDENT HISTORY PATCH`).
- When a shipped JS/CSS/HTML file changes or a new one is added, bump
  `CACHE_NAME` in `sw.js` and add new files to its `SHELL_FILES`. Installed
  iPads won't pick up the change otherwise.
- Recent Activity cards follow a privacy allow-list (`RECENT_ACTIVITY.cardModel`):
  no notes, IDs, emails, or raw timestamps on cards. Full details (including
  notes) appear only behind a deliberate action (the edit flow, or the
  **Details** tap in Visit History).
- Editing and deleting are **today-only**. `PACE_DATA.updateVisit()` enforces
  this independently of the UI.

## Current feature set (as of 11-23-2026)

Start Visit (live, open row) → Mark Complete on the same row; Log Completed
Visit (after the fact); multiple Behavior Specialists; "Teacher Came From";
required Notes on completion (at least one non-blank character); live visit
notes with autosave; same-day edit and delete from Recent Activity; "Other"
location; IEP_App_Users permission gate; a temporary read-only SharePoint
diagnostic button; and, new, **viewing earlier days and a per-student Visit
History** (STUDENT HISTORY PATCH).

## Open items / known gaps

- SharePoint columns `Behavior Specialist` and `Teacher Came From` may not
  exist on the live list yet (written speculatively; see README "Known gaps").
- `Staff Member` is a Person field and is never written.
- Whether `Notes` is "Multiple lines of text" is unconfirmed.
- The per-date and per-student Graph reads filter on non-indexed columns.
  An admin should index `Date` and `Student` on `IEP_Pace_Visits` before it
  passes 5,000 items, or those reads will start failing.
- The temporary PATCH 003 diagnostic is still in place (marked for removal).
- `tests/live-visit-notes.test.js` (in the repo) fails on Windows checkouts
  only. One assertion searches for a `\n` line break, but `core.autocrlf`
  gives `\r\n`. The app code is correct, and the test passes with LF line
  endings. Fix belongs in the test (use the `\s*` regex already on the line above).
- Visit History: staff search by **name**. The roster list holds all students
  from all years, and Student IDs and names don't change. After a name is
  picked, a visit matches by `Student ID` when both have one, otherwise by the
  exact name (rows written by MAC Walkthrough have no ID). No date limit.
- STUDENT HISTORY PATCH is verified in demo mode only. It still needs a
  signed-in production check on a real iPad.
