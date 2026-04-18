# Model change log

A change is "model-impacting" iff it modifies any of:

* `parity_constants.py` (any tolerance or epsilon value).
* Mortality table loader, q_x source, or generational projection.
* Yield-curve construction, interpolation, or extrapolation policy.
* ALM disinvestment / reinvestment rule, borrowing policy, or rebalance
  policy.
* RILA crediting (cap, floor, participation, fee, segment window).
* SPIA payment timing / cessation policy.
* Term Life premium / claim cashflow definition.

Every entry MUST include:

| Field             | Required content                                                |
|-------------------|------------------------------------------------------------------|
| Date              | ISO `YYYY-MM-DD`                                                 |
| Version / PR      | Semver tag if released, plus PR link                             |
| Author / reviewer | The CODEOWNER who approved                                       |
| Summary           | One sentence: what changed                                       |
| Justification     | Why -- regulatory, bug fix, methodology improvement              |
| Parity trace      | Path to before/after CSVs from `scripts/parity_trace.py`         |
| Backward compat   | Whether existing scenarios reproduce within the prior tolerance  |

Entries are appended at the **bottom**; never edit historical entries.

---

## 2026-04-18 -- Tolerance constants centralised (no numeric change)

- **PR:** hardening-roadmap (this branch)
- **Author / reviewer:** lttim
- **Summary:** Created `parity_constants.py` as the single source of truth for
  tolerance / epsilon values. Refactored `tests/parity/test_alm_parity.py`,
  `test_rila_parity.py`, `test_term_parity.py`, and `excel_formula_sim.py`
  to import from the new module. No numeric values changed.
- **Justification:** Eliminates the doc-vs-test drift class of bug. Makes
  CI's "tolerance was secretly weakened" check possible
  (`scripts/render_parity_contract.py --check`).
- **Parity trace:** N/A (refactor only). Verified by full
  `pytest tests/parity` green and `scripts/render_parity_contract.py --check`
  green.
- **Backward compat:** Yes -- bit-for-bit identical numerics to prior tag.

## 2026-04-18 -- RILA PV tolerance tightened from 1e-3 to 1e-4

- **PR:** hardening-roadmap (P0 doc-drift fix)
- **Author / reviewer:** lttim
- **Summary:** Brought `tests/parity/test_rila_parity.py` PV assertions in
  line with `docs/rila_parity_contract.md` (`1e-4`).
- **Justification:** The contract has stated `1e-4` since v1.0; the test was
  silently looser. No production scenario observed at the looser tolerance,
  so the tightening is a documentation-conformance fix, not a methodology
  change.
- **Parity trace:** Existing parity tests pass at the tighter tolerance.
- **Backward compat:** Yes.
