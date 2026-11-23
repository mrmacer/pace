# Agent notes

This folder is the shared memory for every AI agent (Claude Code, Codex, or any
other) that works on PACE Room Tracker, on any machine. A new agent should be
able to read these files and know the full scope, history, and context of the
project without the chat that produced it.

| File | What it holds | Who updates it |
|---|---|---|
| [CONTEXT.md](CONTEXT.md) | The current picture: what the app is, who's involved, how it's built, standing rules, open issues. Rewritten in place as things change. | Any agent, whenever something it describes changes |
| [LOG.md](LOG.md) | A running record of every working session, newest first: what was asked (word for word), what was decided and why, what changed, what's still open. Append-only. | Every agent, at the end of every piece of work |

## Rules for writing here

This repository is **public**. Anything written here is visible to anyone.

1. **Record what was said.** Quote each request from the people working on the
   project word for word, including follow-ups and corrections made mid-task.
   Then record the agent's decisions, the reasons, and anything rejected.
2. **Never write student information.** No student names, initials, IDs, or
   anything that could identify a student, even if it was typed in chat.
   Replace it with `[student]`. Visit notes and SharePoint data never go here.
3. **No secrets or personal contact details.** No email addresses, tokens,
   tenant secrets, or passwords. Staff first names and public GitHub usernames
   are fine.
4. **Facts, not guesses.** If something wasn't verified, say so.
5. **Never rewrite past LOG entries.** If an earlier entry turns out to be
   wrong, add a new entry that corrects it.
