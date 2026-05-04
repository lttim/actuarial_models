from __future__ import annotations

import json
import sys
from pathlib import Path

REPO_ROOT = Path(__file__).resolve().parents[1]
SCRIPTS = REPO_ROOT / "scripts"
sys.path.insert(0, str(SCRIPTS))

import agent_preflight  # noqa: E402
import agent_team_router as router  # noqa: E402


def test_packet_markdown_contains_roles_gates_and_changed_files() -> None:
    plan = router.build_staffing_plan(
        ["annuity_model/pricing_ui.py"],
        objective="Improve Streamlit demo flow",
    )

    text = agent_preflight.render_packet_markdown(plan)

    assert "# Team Run Packet" in text
    assert "UX Reviewer / Builder" in text
    assert "`ui_apptest`" in text
    assert "`annuity_model/pricing_ui.py`" in text
    assert "Orchestrator Integration Summary" in text


def test_write_packet_emits_markdown_and_json(tmp_path: Path) -> None:
    plan = router.build_staffing_plan(
        ["annuity_model/rila_projection.py"],
        objective="RILA model review",
    )

    md_path, json_path = agent_preflight.write_packet(plan, packet_dir=tmp_path)

    assert md_path.exists()
    assert json_path.exists()
    payload = json.loads(json_path.read_text())
    assert payload["plan"]["objective"] == "RILA model review"
    assert payload["plan"]["multi_agent_required"] is True
    assert any(role["role_id"] == "actuarial_peer_reviewer" for role in payload["plan"]["roles"])


def test_write_packet_can_update_existing_packet_paths(tmp_path: Path) -> None:
    plan = router.build_staffing_plan(
        ["annuity_model/docs/index.md"],
        objective="Docs refresh",
    )
    md_path = tmp_path / "team-run.md"
    json_path = tmp_path / "team-run.json"

    written_md, written_json = agent_preflight.write_packet(
        plan,
        packet_dir=tmp_path / "unused",
        markdown_path=md_path,
        json_path=json_path,
    )

    assert written_md == md_path
    assert written_json == json_path
    assert "Docs refresh" in md_path.read_text()
    assert json.loads(json_path.read_text())["plan"]["objective"] == "Docs refresh"


def test_expand_command_substitutes_python_and_changed_files() -> None:
    command = ("python", "scripts/check_test_update_required.py", "{changed_files}")

    expanded = agent_preflight._expand_command(
        command,
        ("annuity_model/rila_projection.py", "annuity_model/alm_excel_ladder.py"),
    )

    assert expanded[0] == sys.executable
    assert expanded[-2:] == (
        "annuity_model/rila_projection.py",
        "annuity_model/alm_excel_ladder.py",
    )


def test_mutmut_gate_receives_annuity_relative_paths() -> None:
    command = ("python", "scripts/mutmut_pr_gate.py", "--touched-files", "{changed_files}")

    expanded = agent_preflight._expand_command(
        command,
        (
            "annuity_model/rila_projection.py",
            "annuity_model/alm_excel_ladder.py",
            ".github/CODEOWNERS",
        ),
    )

    assert expanded[-2:] == ("rila_projection.py", "alm_excel_ladder.py")
    assert ".github/CODEOWNERS" not in expanded


def test_default_git_changed_files_includes_untracked(monkeypatch) -> None:
    calls: list[tuple[str, ...]] = []

    def fake_check_output(cmd, **kwargs):  # noqa: ANN001, ANN202
        calls.append(tuple(cmd))
        if cmd == ["git", "diff", "--name-only", "HEAD"]:
            return "annuity_model/pricing_ui.py\n"
        if cmd == ["git", "ls-files", "--others", "--exclude-standard"]:
            return "annuity_model/tests/test_new_behavior.py\n"
        raise AssertionError(f"unexpected command: {cmd}")

    monkeypatch.setattr(agent_preflight.subprocess, "check_output", fake_check_output)

    changed = agent_preflight._git_changed_files(base=None, head=None, staged=False)

    assert changed == (
        "annuity_model/pricing_ui.py",
        "annuity_model/tests/test_new_behavior.py",
    )
    assert calls == [
        ("git", "diff", "--name-only", "HEAD"),
        ("git", "ls-files", "--others", "--exclude-standard"),
    ]


def test_main_json_dry_run_does_not_write_packet(
    tmp_path: Path,
    capsys,
) -> None:
    rc = agent_preflight.main(
        [
            "--objective",
            "Docs refresh",
            "--changed-files",
            "annuity_model/docs/index.md",
            "--packet-dir",
            str(tmp_path),
            "--json",
        ]
    )

    assert rc == 0
    assert not list(tmp_path.iterdir())
    payload = json.loads(capsys.readouterr().out)
    assert payload["plan"]["surfaces"] == ["docs_governance"]
    assert payload["gate_results"] == []
