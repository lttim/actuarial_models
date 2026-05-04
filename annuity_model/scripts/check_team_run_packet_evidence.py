"""Require Team Run Packet evidence for broad or high-risk changes.

The packet directory is intentionally gitignored, so this guard supports two
evidence locations:

* local packet markdown/JSON under ``.agent-team-runs/`` for pre-commit and
  autonomous-agent handoff; and
* PR body excerpts in CI, where ignored local packets are unavailable.
"""

from __future__ import annotations

import argparse
import json
import os

# Reviewed: local/CI guard invokes fixed git command vectors with shell=False.
import subprocess  # nosec B404
from pathlib import Path
from typing import Any

try:
    from agent_team_router import StaffingPlan, build_staffing_plan
except ModuleNotFoundError:  # pragma: no cover - import path used when imported as a package
    from scripts.agent_team_router import StaffingPlan, build_staffing_plan

ANNUITY_DIR = Path(__file__).resolve().parents[1]
REPO_ROOT = ANNUITY_DIR.parent
DEFAULT_PACKET_DIR = REPO_ROOT / ".agent-team-runs"

REQUIRED_PR_HEADINGS = (
    "Team Run Packet Evidence",
    "Selected Roles",
    "Validation Gates",
    "Review Findings",
    "Unresolved Risks",
    "Final Signoff",
)


def _git_changed_files(base_sha: str | None, head_sha: str | None, staged: bool) -> tuple[str, ...]:
    if bool(base_sha) != bool(head_sha):
        raise ValueError("Provide both base_sha and head_sha, or neither.")
    if base_sha and head_sha:
        cmd = ["git", "diff", "--name-only", f"{base_sha}..{head_sha}"]
    elif staged:
        cmd = ["git", "diff", "--cached", "--name-only"]
    else:
        cmd = ["git", "diff", "--name-only", "HEAD"]
    # Reviewed: fixed git command vector; refs are CI/developer supplied and shell=False.
    out = subprocess.check_output(cmd, cwd=REPO_ROOT, text=True)  # nosec B603
    return tuple(sorted(line.strip() for line in out.splitlines() if line.strip()))


def _normalize_repo_path(path: str) -> str:
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
    return _git_changed_files(args.base, args.head, staged=args.staged)


def _latest_packet_json(packet_dir: Path) -> Path | None:
    packets = sorted(packet_dir.glob("*.json"), key=lambda path: path.stat().st_mtime)
    return packets[-1] if packets else None


def _load_packet(packet_json: Path | None, packet_dir: Path) -> tuple[Path, dict[str, Any]]:
    selected = packet_json or _latest_packet_json(packet_dir)
    if selected is None:
        raise ValueError(f"No Team Run Packet JSON found under {packet_dir}")
    return selected, json.loads(selected.read_text(encoding="utf-8"))


def _role_ids(plan: StaffingPlan) -> set[str]:
    return {role.role_id for role in plan.roles}


def _validate_packet_json(plan: StaffingPlan, payload: dict[str, Any]) -> list[str]:
    errors: list[str] = []
    packet_plan = payload.get("plan")
    if not isinstance(packet_plan, dict):
        return ["packet JSON is missing a plan object"]

    if packet_plan.get("multi_agent_required") is not True:
        errors.append("packet plan does not mark multi_agent_required=true")

    packet_changed = set(packet_plan.get("changed_files") or ())
    missing_changed = set(plan.changed_files) - packet_changed
    if missing_changed:
        errors.append(
            "packet changed_files does not cover current diff: "
            + ", ".join(sorted(missing_changed))
        )

    packet_roles = {role.get("role_id") for role in packet_plan.get("roles") or []}
    missing_roles = _role_ids(plan) - packet_roles
    if missing_roles:
        errors.append("packet is missing selected roles: " + ", ".join(sorted(missing_roles)))

    packet_gates = set(packet_plan.get("gate_ids") or ())
    missing_gates = set(plan.gate_ids) - packet_gates
    if missing_gates:
        errors.append("packet is missing selected gates: " + ", ".join(sorted(missing_gates)))

    gate_results = payload.get("gate_results") or []
    result_by_gate = {
        result.get("gate_id"): result
        for result in gate_results
        if isinstance(result, dict) and result.get("gate_id")
    }
    deferred_results = payload.get("deferred_gate_results") or []
    deferred_by_gate = {
        result.get("gate_id"): result
        for result in deferred_results
        if isinstance(result, dict) and result.get("gate_id")
    }
    incomplete_deferrals = sorted(
        gate_id
        for gate_id, result in deferred_by_gate.items()
        if gate_id in plan.gate_ids
        and (not result.get("reason") or not result.get("next_validation"))
    )
    if incomplete_deferrals:
        errors.append(
            "packet has incomplete deferred gate evidence: " + ", ".join(incomplete_deferrals)
        )

    missing_gate_results = set(plan.gate_ids) - set(result_by_gate) - set(deferred_by_gate)
    if missing_gate_results:
        errors.append(
            "packet is missing gate result evidence: " + ", ".join(sorted(missing_gate_results))
        )
    failing_gates = sorted(
        gate_id
        for gate_id, result in result_by_gate.items()
        if gate_id in plan.gate_ids and result.get("exit_code") != 0
    )
    if failing_gates:
        errors.append("packet records failing gates: " + ", ".join(failing_gates))

    return errors


def _section_after_heading(text: str, heading: str) -> str:
    lines = text.splitlines()
    lower_heading = heading.lower()
    for idx, line in enumerate(lines):
        if line.lstrip("# ").strip().lower() == lower_heading:
            collected: list[str] = []
            for later in lines[idx + 1 :]:
                if later.startswith("#"):
                    break
                collected.append(later)
            return "\n".join(collected).strip()
    return ""


def _validate_packet_markdown(packet_markdown: Path) -> list[str]:
    text = packet_markdown.read_text(encoding="utf-8")
    errors: list[str] = []
    for heading in ("Per-Agent Outputs", "Review Findings", "Orchestrator Integration Summary"):
        section = _section_after_heading(text, heading)
        pending_marker = any(
            line.strip().lower() in {"pending", "- pending"} for line in section.splitlines()
        )
        if not section or pending_marker:
            errors.append(f"packet markdown section is incomplete: {heading}")
    signoff = _section_after_heading(text, "Final Signoff")
    signoff_pending = any(
        line.strip().lower() in {"pending", "- pending"} for line in signoff.splitlines()
    )
    if not signoff or signoff_pending or "complete" not in signoff.lower():
        errors.append("packet markdown final signoff must be complete")
    return errors


def _read_pr_body(args: argparse.Namespace) -> str:
    if args.pr_body:
        return args.pr_body
    if args.pr_body_file:
        return args.pr_body_file.read_text(encoding="utf-8")
    event_path = Path(args.github_event_path) if args.github_event_path else None
    if event_path and event_path.exists():
        event = json.loads(event_path.read_text(encoding="utf-8"))
        pr = event.get("pull_request") or {}
        body = pr.get("body")
        return body if isinstance(body, str) else ""
    return ""


def _validate_pr_body(plan: StaffingPlan, body: str) -> list[str]:
    errors: list[str] = []
    if not body.strip():
        return ["PR body is missing Team Run Packet evidence"]
    missing_headings = [heading for heading in REQUIRED_PR_HEADINGS if heading not in body]
    if missing_headings:
        errors.append("PR body is missing evidence headings: " + ", ".join(missing_headings))
    missing_roles = [role.display_name for role in plan.roles if role.display_name not in body]
    if missing_roles:
        errors.append("PR body is missing selected role evidence: " + ", ".join(missing_roles))
    missing_gates = [gate_id for gate_id in plan.gate_ids if gate_id not in body]
    if missing_gates:
        errors.append("PR body is missing validation gate evidence: " + ", ".join(missing_gates))
    has_pending_marker = any(
        line.strip().lower() in {"pending", "- pending"} for line in body.splitlines()
    )
    if has_pending_marker:
        errors.append("PR body evidence still contains 'pending'")
    if "COMPLETE" not in body and "Complete" not in body:
        errors.append("PR body final signoff must state COMPLETE")
    return errors


def evaluate_evidence(
    changed_files: tuple[str, ...],
    *,
    packet_json: Path | None = None,
    packet_dir: Path = DEFAULT_PACKET_DIR,
    pr_body: str = "",
) -> tuple[StaffingPlan, list[str]]:
    plan = build_staffing_plan(changed_files)
    if not plan.multi_agent_required:
        return plan, []

    if pr_body:
        return plan, _validate_pr_body(plan, pr_body)

    try:
        selected_json, payload = _load_packet(packet_json, packet_dir)
    except (OSError, ValueError, json.JSONDecodeError) as exc:
        return plan, [str(exc)]
    errors = _validate_packet_json(plan, payload)
    errors.extend(_validate_packet_markdown(selected_json.with_suffix(".md")))
    return plan, errors


def _print_result(plan: StaffingPlan, errors: list[str]) -> None:
    print(f"Surfaces: {', '.join(plan.surfaces) or '(none detected)'}")
    print(f"Multi-agent required: {plan.multi_agent_required}")
    if not plan.multi_agent_required:
        print("OK: Team Run Packet evidence is not required for this diff.")
        return
    if not errors:
        print("OK: Team Run Packet evidence is complete for this diff.")
        return
    print("Team Run Packet evidence is incomplete for this broad/high-risk diff:")
    for error in errors:
        print(f"  - {error}")


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(
        description="Validate Team Run Packet evidence for broad/high-risk changes."
    )
    parser.add_argument("--changed-files", nargs="*", default=(), help="Repo-relative paths.")
    parser.add_argument("--base", default=None, help="Base commit for git diff validation.")
    parser.add_argument("--head", default=None, help="Head commit for git diff validation.")
    parser.add_argument("--staged", action="store_true", help="Validate staged changes.")
    parser.add_argument("--packet-json", type=Path, default=None, help="Packet JSON to validate.")
    parser.add_argument(
        "--packet-dir", type=Path, default=DEFAULT_PACKET_DIR, help="Packet directory."
    )
    parser.add_argument("--pr-body", default="", help="Pull request body text to validate.")
    parser.add_argument("--pr-body-file", type=Path, default=None, help="File containing PR body.")
    parser.add_argument(
        "--github-event-path",
        default=None,
        help="GitHub event payload path; defaults to GITHUB_EVENT_PATH in Actions.",
    )
    args = parser.parse_args(argv)
    if args.github_event_path is None:
        args.github_event_path = os.environ.get("GITHUB_EVENT_PATH")

    changed_files = resolve_changed_files(args)
    pr_body = _read_pr_body(args)
    try:
        plan, errors = evaluate_evidence(
            changed_files,
            packet_json=args.packet_json,
            packet_dir=args.packet_dir,
            pr_body=pr_body,
        )
    except (OSError, ValueError, json.JSONDecodeError) as exc:
        plan = build_staffing_plan(changed_files)
        errors = [str(exc)]
    _print_result(plan, errors)
    return 1 if errors else 0


if __name__ == "__main__":
    raise SystemExit(main())
