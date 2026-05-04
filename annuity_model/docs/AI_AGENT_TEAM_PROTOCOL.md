# AI Agent Team Protocol

This protocol defines how autonomous AI agents should staff and coordinate work
in `annuity_model`. It is platform-neutral: Codex, Cursor, Claude Code, and
future agents should all follow the same operating model.

The user should only need to describe the feature or fix. The orchestrator
agent is responsible for staffing the team, assigning scoped ownership,
collecting review evidence, running gates, and declaring completion only after
the Team Run Packet is coherent.

## Source of Truth

Use these files together:

- `docs/AI_AGENT_PREFLIGHT.md` -- task classification and hard repo rules.
- `scripts/agent_team_router.py` -- deterministic staffing and gate selection.
- `scripts/agent_preflight.py` -- Team Run Packet writer and optional gate runner.
- This file -- role authority, dynamic staffing rules, and evidence contract.

If these disagree, fix the docs/scripts in the same change. Do not silently
choose whichever instruction is convenient.

## Core Team

The default team catalog is encoded in `scripts/agent_team_router.py`.

| Role | Authority | Typical trigger |
|---|---|---|
| Orchestrator / Integrator | integration owner | Every autonomous task |
| Model Engineer | scoped write | Engines, product definitions, liability/scenario logic |
| Excel Builder Engineer | scoped write | Workbook builders, formula layouts, validator-facing workbook structure |
| Validation Engineer | scoped write | Tests, parity, smoke, mutation, security, coverage evidence |
| Actuarial Peer Reviewer | read-only | Calculation, Excel, portfolio, assumption, or methodology risk |
| UX Reviewer / Builder | scoped write | Streamlit UI, state, AppTest surfaces |
| Docs / Governance Steward | scoped write | Agent instructions, docs inventory, release/governance artifacts |

Review and judgment roles remain read-only. Write-capable roles must stay within
their assigned owned paths unless the orchestrator updates the Team Run Packet.

## Dynamic Specialist Roles

The orchestrator may create new specialist roles automatically when the task
requires them. Examples include:

- Security Reviewer
- Packaging Engineer
- Data Governance Reviewer
- Performance Engineer
- Release Manager
- Migration Planner

Dynamic roles do not require user approval. They do require a role contract in
the Team Run Packet with:

- purpose
- authority level
- read-only vs write-capable status
- owned paths
- dependencies
- expected artifact
- acceptance checks
- rationale for staffing the role

If a dynamic role proves useful repeatedly, promote it into the documented
catalog in `scripts/agent_team_router.py` and update this protocol.

## Automatic Staffing Rules

Run the router before implementation:

```bash
python scripts/agent_preflight.py \
  --objective "<feature or fix>" \
  --write-packet
```

When explicit changed files are known:

```bash
python scripts/agent_preflight.py \
  --objective "<feature or fix>" \
  --changed-files annuity_model/pricing_ui.py annuity_model/tests/ui/test_apptest_spia.py \
  --write-packet
```

The router automatically requires multi-agent staffing when it detects:

- multiple subsystems
- calculation, Excel, portfolio, assumption, or packaging surfaces
- UI plus state changes
- governance or release changes
- validation-heavy work
- broad refactors

Default concurrency is capped at five non-orchestrator roles. The orchestrator
may staff more roles only when the Team Run Packet explains why the work is
independent and useful.

All staffing flows through the orchestrator. Specialist agents may recommend
additional help, but they do not independently spawn more agents.

## Team Run Packet

Every multi-agent run must produce a packet under `.agent-team-runs/`.
The directory is gitignored because packets are session evidence, but important
excerpts should be copied into the PR description or release record.

The packet records:

- objective
- selected surfaces
- selected roles and rationale
- owned paths and write scopes
- validation gates selected by the router
- per-agent outputs
- review findings
- unresolved risks
- orchestrator integration summary
- final signoff

The orchestrator cannot claim completion for a multi-agent task until the
packet is updated with final gates and unresolved risks.

## Completion Protocol

Before completion:

1. Run `scripts/agent_preflight.py --run-gates --write-packet` against the
   final diff or explicit changed files.
2. Trigger Actuarial Peer Review for calculation, Excel, portfolio, assumption,
   or tolerance-facing changes.
3. Ensure every write-capable role’s work is integrated by the orchestrator.
4. Record gate exit codes, review findings, and residual risk in the packet.
5. Attach or summarize packet evidence in the final response or PR.

The four canonical gates in `annuity_model/AGENTS.md` remain mandatory where
applicable. This protocol adds automatic staffing and evidence capture; it does
not replace parity, smoke, or rendered-contract checks.
