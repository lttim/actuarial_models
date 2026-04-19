# Actuary SME -- review verdicts and evidence

This directory holds the **personal session artifacts** of the Actuary
SME review workflow defined by
[`.cursor/rules/actuary-sme-protocol.mdc`](../rules/actuary-sme-protocol.mdc)
and the skill at
[`annuity_model/.cursor/skills/actuary-sme/SKILL.md`](../../annuity_model/.cursor/skills/actuary-sme/SKILL.md).

Files are gitignored except for this README and the `.gitkeep`. The
verdicts are personal -- like `.cursor/handoffs/` -- and not committed.

## What lives here

| Path | Purpose | Lifecycle |
|---|---|---|
| `iter-<N>-<UTC>-<scope>.md` | One verdict per iteration of an SME loop. YAML frontmatter is the parser's source of truth; markdown body is the human-readable narrative. | Created by the loop; kept indefinitely as audit trail. |
| `_evidence-current.md` | Transient evidence pack the SME subagent reads. Overwritten at the start of every iteration. | Lifecycle: per-iteration. |
| `README.md` | This file. | Tracked in git. |
| `.gitkeep` | Keeps the directory present in git. | Tracked. |

## Triggering a review

Three equivalent ways:

1. **Explicit command** in chat: `!actuaryreview` (also accepts `full`,
   `<product>`, or `status` arguments).
2. **Natural language** in chat: any phrasing containing "actuary
   review", "have the actuary review", "ask the actuary SME",
   "actuarial review please", etc. The rule routes these through the
   same orchestration as the explicit command.
3. **Auto-fired**: when the AI agent's session edited any file in the
   CALCULATION or TOLERANCE branches of
   `annuity_model/docs/AI_AGENT_PREFLIGHT.md`. The rule's auto-trigger
   globs cover engines, builders, parity_constants,
   actuarial_benchmarks, and product subpackages.

## Reading verdicts

The most recent verdict is the one with the highest `iter-<N>` value
for the most recent UTC timestamp. To list recent verdicts, use the
explicit `!actuaryreview status` command in chat -- it prints the last
~10 verdict files newest-first.

For a verdict file:

- The YAML frontmatter at the top is the structured payload (verdict,
  scope, findings, prior-finding regression check). The orchestration
  loop parses this directly.
- The markdown body below is the human-readable narrative grouped by
  finding category (sign, band, sensitivity, closed-form,
  cross-product, assumption, methodology, documentation).

## When the loop stops without APPROVE

Three escalation conditions cause the loop to stop and print a chain
of verdict file paths instead of declaring task complete:

1. **MAX_ITERATIONS exceeded** (default 5; override with
   `ACTUARY_REVIEW_MAX_ITER`).
2. **Recurring finding**: the parent agent's fix did not resolve a
   prior `[AGENT-FIXABLE]` finding (regression).
3. **Human judgment required**: a `[NEEDS-HUMAN-JUDGMENT]` finding
   inside a `BLOCK` verdict (e.g., "is the synthetic CSO placeholder
   fit for production?").

In all three cases the user is **not** prompted -- the chain is
printed and the user can pick it up when they next look at the chat.

## Difference from `.cursor/handoffs/`

| Aspect | `.cursor/handoffs/` | `.cursor/actuary-reviews/` |
|---|---|---|
| Purpose | Cross-session continuity (chat handoff) | Per-change actuarial-fitness verdict |
| Trigger | User types `!handoff` | User types `!actuaryreview` (or natural-language equivalent), or auto-fired by the rule |
| Audience | The next agent session | The current loop's parent agent + future reviewer |
| Cardinality | One file per handoff | One file per loop iteration (chain) |
| Required action on read | None (read at session start) | The loop reads the prior verdict to verify whether prior findings were addressed |
