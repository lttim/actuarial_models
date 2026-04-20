from __future__ import annotations

from scripts.check_test_update_required import evaluate_changed_files


def test_guard_passes_when_no_behavior_files_changed() -> None:
    result = evaluate_changed_files(
        [
            "README.md",
            "annuity_model/docs/index.md",
            "annuity_model/scripts/render_parity_contract.py",
        ]
    )
    assert result.ok
    assert result.behavior_files == ()


def test_guard_blocks_behavior_change_without_tests() -> None:
    result = evaluate_changed_files(["annuity_model/pricing_ui.py"])
    assert not result.ok
    assert result.behavior_files == ("annuity_model/pricing_ui.py",)
    assert result.test_files == ()


def test_guard_passes_when_behavior_and_tests_change_together() -> None:
    result = evaluate_changed_files(
        [
            "annuity_model/portfolio_runner.py",
            "annuity_model/tests/parity/test_portfolio_golden.py",
        ]
    )
    assert result.ok
    assert result.behavior_files == ("annuity_model/portfolio_runner.py",)
    assert result.test_files == ("annuity_model/tests/parity/test_portfolio_golden.py",)
