# Chat Handoffs

This directory stores agent-chat handoff documents created by the `!handoff`
command and consumed by `!recall`. The protocol lives in
`.cursor/rules/handoff-recall.mdc` and is loaded automatically into every
Cursor agent session in this workspace.

## Quick reference

| Command | Effect |
|---------|--------|
| `!handoff` | Save the current chat as a new handoff file. |
| `!handoff <slug>` | Same, but use a custom kebab-case slug. |
| `!recall` | Load the most recent handoff into the current chat. |
| `!recall list` | List all available handoffs (newest first). |
| `!recall <substring>` | Load the newest handoff whose filename matches. |

## File naming

`YYYY-MM-DD-HHMMSS-<slug>.md` (UTC). Lexical sort = chronological sort, so
"newest" is always `ls -1 | tail -n 1`.

## Git policy

Handoff `.md` files are git-ignored by default — they are personal session
artifacts, often containing in-flight thinking. If you want to share one,
either commit it explicitly with `git add -f` or copy its contents into a
proper doc (e.g. an issue, PR description, or design note).
