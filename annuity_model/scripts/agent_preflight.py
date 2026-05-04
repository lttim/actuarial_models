"""Autonomous AI-team preflight and Team Run Packet writer.

The router in :mod:`agent_team_router` decides which specialist roles and
validation gates a task needs. This script is the repo-native command an
orchestrator agent runs to make that decision auditable, write a packet, and
optionally execute the selected gates.

Typical usage from the repository root::

    python annuity_model/scripts/agent_preflight.py \\
        --objective "Add RILA fee sensitivity" \\
        --write-packet

Before completion, add ``--run-gates`` to execute the selected gates and append
their exit codes to the same packet.
"""

from __future__ import annotations

import argparse
import datetime as _dt
import json
import os

# Reviewed: this local orchestration tool invokes fixed git/gate command vectors with shell=False.
import subprocess  # nosec B404
import sys
from dataclasses import asdict, dataclass
from pathlib import Path

from agent_team_router import GateSpec, StaffingPlan, build_staffing_plan, gate_specs_for_plan

ANNUITY_DIR = Path(__file__).resolve().parents[1]
REPO_ROOT = ANNUITY_DIR.parent
DEFAULT_PACKET_DIR = REPO_ROOT / ".agent-team-runs"


@dataclass(frozen=True, slots=True)
class GateResult:
    gate_id: str
    label: str
    command: tuple[str, ...]
    cwd: str
    exit_code: int
    stdout_tail: str
    stderr_tail: str


def _git_changed_files(*, base: str | None, head: str | None, staged: bool) -> tuple[str, ...]:
    if bool(base) != bool(head):
        raise ValueError("Provide both --base and --head, or neither.")
    if base and head:
        cmd = ["git", "diff", "--name-only", f"{base}..{head}"]
        # Reviewed: fixed git command vector; refs are developer/CI supplied and shell=False.
        out = subprocess.check_output(cmd, cwd=REPO_ROOT, text=True)  # nosec B603
        return tuple(sorted(line.strip() for line in out.splitlines() if line.strip()))
    elif staged:
        cmd = ["git", "diff", "--cached", "--name-only"]
        # Reviewed: fixed git command vector for local staged files; shell=False.
        out = subprocess.check_output(cmd, cwd=REPO_ROOT, text=True)  # nosec B603
        return tuple(sorted(line.strip() for line in out.splitlines() if line.strip()))
    else:
        cmd = ["git", "diff", "--name-only", "HEAD"]
        # Reviewed: fixed git command vector for local working tree discovery; shell=False.
        out = subprocess.check_output(cmd, cwd=REPO_ROOT, text=True)  # nosec B603
        tracked = {line.strip() for line in out.splitlines() if line.strip()}
        untracked_cmd = ["git", "ls-files", "--others", "--exclude-standard"]
        # Reviewed: fixed git command vector for local untracked file discovery; shell=False.
        untracked_out = subprocess.check_output(  # nosec B603
            untracked_cmd, cwd=REPO_ROOT, text=True
        )
        untracked = {line.strip() for line in untracked_out.splitlines() if line.strip()}
        return tuple(sorted(tracked | untracked))


def _normalize_repo_path(path: str) -> str:
    """Return a stable repo-relative path for CLI supplied changed files."""
    raw = Path(path)
    root_relative_prefixes = {
        ".devcontainer",
        ".github",
        "actuarial_parity_kit",
        "annuity_model",
    }
    root_relative_files = {
        ".pre-commit-config.yaml",
        "AGENTS.md",
        "CONTRIBUTING.md",
        "DOCUMENTATION_MAP.md",
        "Dockerfile",
        "Justfile",
        "PROJECT_DEVELOPMENT_GUIDE.md",
        "README.md",
    }
    if raw.is_absolute():
        candidate = raw
    elif raw.parts and (
        raw.parts[0] in root_relative_prefixes or raw.parts[0] in root_relative_files
    ):
        candidate = REPO_ROOT / raw
    else:
        candidate = Path.cwd() / raw
    try:
        return candidate.resolve().relative_to(REPO_ROOT).as_posix()
    except ValueError:
        return path.replace("\\", "/")


def resolve_changed_files(args: argparse.Namespace) -> tuple[str, ...]:
    explicit = tuple(sorted({_normalize_repo_path(p) for p in args.changed_files if p}))
    if explicit:
        return explicit
    return _git_changed_files(base=args.base, head=args.head, staged=args.staged)


def _utc_slug() -> str:
    return _dt.datetime.now(_dt.UTC).strftime("%Y-%m-%d-%H%M%S")


def _objective_slug(objective: str) -> str:
    cleaned = "".join(ch.lower() if ch.isalnum() else "-" for ch in objective.strip())
    parts = [p for p in cleaned.split("-") if p]
    return "-".join(parts[:6]) or "agent-preflight"


def packet_paths(packet_dir: Path, objective: str) -> tuple[Path, Path]:
    stem = f"{_utc_slug()}-{_objective_slug(objective)}"
    return packet_dir / f"{stem}.md", packet_dir / f"{stem}.json"


def _render_role_table(plan: StaffingPlan) -> str:
    rows = [
        "| Role | Type | Authority | Read-only | Expected artifact |",
        "|---|---|---|---:|---|",
    ]
    for role in plan.roles:
        role_type = "dynamic" if role.dynamic else "core"
        rows.append(
            "| "
            + " | ".join(
                [
                    role.display_name,
                    role_type,
                    role.authority,
                    "yes" if role.read_only else "no",
                    role.expected_artifact.replace("|", "\\|"),
                ]
            )
            + " |"
        )
    return "\n".join(rows)


def _render_gate_table(plan: StaffingPlan, gate_results: tuple[GateResult, ...]) -> str:
    by_id = {result.gate_id: result for result in gate_results}
    rows = [
        "| Gate | Command | Status |",
        "|---|---|---|",
    ]
    for spec in gate_specs_for_plan(plan):
        result = by_id.get(spec.gate_id)
        status = "not run" if result is None else f"exit {result.exit_code}"
        command = " ".join(spec.command)
        rows.append(f"| `{spec.gate_id}` | `{command}` | {status} |")
    return "\n".join(rows)


def render_packet_markdown(plan: StaffingPlan, gate_results: tuple[GateResult, ...] = ()) -> str:
    changed = "\n".join(f"- `{path}`" for path in plan.changed_files) or "- none detected"
    rationale = "\n".join(f"- {item}" for item in plan.staffing_rationale)
    surfaces = ", ".join(plan.surfaces) or "none detected"
    unresolved = (
        "- Assumption guardrail may require waiver evidence when placeholder data is in scope.\n"
        "- Branch-protected CI evidence is external to local preflight and must be attached by the orchestrator."
    )
    return f"""# Team Run Packet

- **Created:** {_dt.datetime.now(_dt.UTC).isoformat(timespec="seconds").replace("+00:00", "Z")}
- **Objective:** {plan.objective or "(not provided)"}
- **Surfaces:** {surfaces}
- **Multi-agent required:** {plan.multi_agent_required}
- **Recommended concurrency:** {plan.recommended_concurrency}
- **Soft cap exceeded:** {plan.soft_cap_exceeded}

## Staffing Rationale

{rationale}

## Selected Team

{_render_role_table(plan)}

## Changed Files

{changed}

## Validation Gates

{_render_gate_table(plan, gate_results)}

## Per-Agent Outputs

Agents must append their role-specific outputs here or link to their own
review/evidence artifacts. Review agents remain read-only. Write-capable agents
must stay inside their declared owned paths unless the orchestrator updates this
packet first.

## Review Findings

- pending

## Unresolved Risks

{unresolved}

## Orchestrator Integration Summary

- pending

## Final Signoff

- pending
"""


def _packet_json(plan: StaffingPlan, gate_results: tuple[GateResult, ...]) -> dict[str, object]:
    return {
        "plan": plan.to_dict(),
        "gate_results": [asdict(result) for result in gate_results],
    }


def write_packet(
    plan: StaffingPlan,
    *,
    packet_dir: Path,
    gate_results: tuple[GateResult, ...] = (),
    markdown_path: Path | None = None,
    json_path: Path | None = None,
) -> tuple[Path, Path]:
    if markdown_path is not None:
        md_path = markdown_path
        json_target = json_path or markdown_path.with_suffix(".json")
    elif json_path is not None:
        json_target = json_path
        md_path = json_path.with_suffix(".md")
    else:
        md_path, json_target = packet_paths(packet_dir, plan.objective)
    md_path.parent.mkdir(parents=True, exist_ok=True)
    json_target.parent.mkdir(parents=True, exist_ok=True)
    md_path.write_text(render_packet_markdown(plan, gate_results), encoding="utf-8")
    json_target.write_text(
        json.dumps(_packet_json(plan, gate_results), indent=2) + "\n", encoding="utf-8"
    )
    return md_path, json_target


def _changed_files_for_command(
    command: tuple[str, ...], changed_files: tuple[str, ...]
) -> tuple[str, ...]:
    if "scripts/mutmut_pr_gate.py" not in command:
        return changed_files
    return tuple(
        path.removeprefix("annuity_model/")
        for path in changed_files
        if path.startswith("annuity_model/")
    )


def _expand_command(command: tuple[str, ...], changed_files: tuple[str, ...]) -> tuple[str, ...]:
    expanded: list[str] = []
    command_changed_files = _changed_files_for_command(command, changed_files)
    for part in command:
        if part == "python":
            expanded.append(sys.executable)
        elif part == "{changed_files}":
            expanded.extend(command_changed_files)
        else:
            expanded.append(part)
    return tuple(expanded)


def run_gate(spec: GateSpec, changed_files: tuple[str, ...]) -> GateResult:
    cwd = REPO_ROOT / spec.cwd
    env = os.environ.copy()
    env.update(spec.env)
    command = _expand_command(spec.command, changed_files)
    # Reviewed: gate specs are repo-defined local validation commands run with shell=False.
    proc = subprocess.run(  # nosec B603
        command,
        cwd=cwd,
        env=env,
        text=True,
        capture_output=True,
        check=False,
    )
    stdout_tail = "\n".join(proc.stdout.splitlines()[-40:])
    stderr_tail = "\n".join(proc.stderr.splitlines()[-40:])
    return GateResult(
        gate_id=spec.gate_id,
        label=spec.label,
        command=command,
        cwd=str(cwd),
        exit_code=proc.returncode,
        stdout_tail=stdout_tail,
        stderr_tail=stderr_tail,
    )


def run_selected_gates(plan: StaffingPlan) -> tuple[GateResult, ...]:
    return tuple(run_gate(spec, plan.changed_files) for spec in gate_specs_for_plan(plan))


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(
        description="Route, packetize, and optionally run agent preflight gates."
    )
    parser.add_argument("--objective", default="", help="Short task objective.")
    parser.add_argument(
        "--changed-files", nargs="*", default=(), help="Repo-relative paths to route."
    )
    parser.add_argument("--base", default=None, help="Base commit for git diff routing.")
    parser.add_argument("--head", default=None, help="Head commit for git diff routing.")
    parser.add_argument(
        "--staged", action="store_true", help="Route staged changes instead of working-tree diff."
    )
    parser.add_argument(
        "--run-gates", action="store_true", help="Execute selected gates and record exit codes."
    )
    parser.add_argument(
        "--write-packet", action="store_true", help="Write markdown and JSON Team Run Packet files."
    )
    parser.add_argument(
        "--packet-markdown",
        type=Path,
        default=None,
        help="Existing or target markdown Team Run Packet to update instead of creating a new timestamped packet.",
    )
    parser.add_argument(
        "--packet-json",
        type=Path,
        default=None,
        help="Existing or target JSON Team Run Packet to update instead of creating a new timestamped packet.",
    )
    parser.add_argument(
        "--packet-dir", type=Path, default=DEFAULT_PACKET_DIR, help="Directory for packet output."
    )
    parser.add_argument("--json", action="store_true", help="Print machine-readable plan/results.")
    args = parser.parse_args(argv)

    changed_files = resolve_changed_files(args)
    plan = build_staffing_plan(changed_files, objective=args.objective)
    gate_results = run_selected_gates(plan) if args.run_gates else ()

    packet_written: tuple[Path, Path] | None = None
    if args.write_packet or args.packet_markdown is not None or args.packet_json is not None:
        packet_written = write_packet(
            plan,
            packet_dir=args.packet_dir,
            gate_results=gate_results,
            markdown_path=args.packet_markdown,
            json_path=args.packet_json,
        )

    payload = _packet_json(plan, gate_results)
    if packet_written is not None:
        payload["packet_markdown"] = str(packet_written[0])
        payload["packet_json"] = str(packet_written[1])

    if args.json:
        print(json.dumps(payload, indent=2, sort_keys=True))
    else:
        print(f"Surfaces: {', '.join(plan.surfaces) or '(none detected)'}")
        print(f"Multi-agent required: {plan.multi_agent_required}")
        print("Roles: " + ", ".join(role.display_name for role in plan.roles))
        print("Gates: " + (", ".join(plan.gate_ids) or "(none selected)"))
        if packet_written is not None:
            print(f"Team Run Packet: {packet_written[0]}")
        for result in gate_results:
            print(f"{result.gate_id}: exit {result.exit_code}")

    return 1 if any(result.exit_code != 0 for result in gate_results) else 0


if __name__ == "__main__":
    raise SystemExit(main())
