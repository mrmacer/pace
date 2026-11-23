# PACE Room Tracker — instructions for Claude

**Before starting any work, read:**
1. [docs/agent-notes/CONTEXT.md](docs/agent-notes/CONTEXT.md): what this app is, who's involved, the standing rules, and open issues.
2. [docs/agent-notes/LOG.md](docs/agent-notes/LOG.md): what has been asked and decided in every earlier session (newest first).
3. [README.md](README.md): the detailed patch history and SharePoint schema notes.

**At the end of every piece of work, before handing it over:**
- Add an entry at the top of `docs/agent-notes/LOG.md`, following the template
  and rules in [docs/agent-notes/README.md](docs/agent-notes/README.md). Quote
  every request, follow-up, and correction from the people you worked with
  word for word, and record your decisions and reasons, what changed, how you
  verified it, and what's still open.
- Update `docs/agent-notes/CONTEXT.md` if anything it describes changed.
- This repo is **public**: never write student information (use `[student]`),
  email addresses, or secrets in these notes, in code comments, or in commit
  text.

**Working rules from the owner:**
- Never run `git commit` or `git push`. Hand over a commit title and description instead.
- Keep local tooling and sandboxes outside the repo in `~/.pace/`. Rebuild it from
  [docs/LOCAL-KIT.md](docs/LOCAL-KIT.md) if it's missing, and update that file
  whenever the kit changes.
