# Assumption governance standard

This standard adds governance metadata on top of the technical
artifact registry in [`data_registry.py`](../data_registry.py).

## Required metadata for assumption sets

Each governed assumption set should have:

- `artifact_name`: maps to `data_registry.REGISTRY` entry name.
- `assumption_family`: mortality, lapse, expense, yield curve, index scenario.
- `approval_id`: committee minute ID or documented approval reference.
- `approved_by`: owner role or committee.
- `challenged_by`: independent challenge role or reviewer.
- `approval_date`: ISO date.
- `valid_from` / `valid_to`: applicability period.
- `intended_use`: pricing, valuation, stress testing, or development only.
- `status`: approved, provisional, deprecated.
- `requires_waiver_for_release`: bool.
- `notes`: free-text caveats and residual risk.

## Operating policies

1. **No silent assumption swaps:** any artifact version change requires a documented governance record.
2. **Placeholder controls:** synthetic/placeholder sets must be tagged as release-restricted and require waiver evidence.
3. **Challenger evidence:** Tier 1 assumptions require explicit independent challenger metadata.
4. **Expiry discipline:** assumptions cannot be used beyond `valid_to` without renewal.
5. **Traceability:** release artifacts must identify the exact assumption versions used.

## Suggested file format

Governed assumptions should be tracked in a machine-readable file,
for example `annuity_model/data/assumptions/assumption_approvals.yaml`.

Minimal entry shape:

```yaml
- artifact_name: cso_2017_ult_male_nonsmoker_qx
  assumption_family: mortality
  approval_id: MRC-2026-041
  approved_by: chief_actuary
  challenged_by: model_risk_review
  approval_date: 2026-04-19
  valid_from: 2026-04-19
  valid_to: 2027-04-19
  intended_use: development_only
  status: provisional
  requires_waiver_for_release: true
  notes: Synthetic placeholder; replace with licensed table for production.
```

## Enforcement hooks

- Use [`scripts/check_assumption_release_guardrails.py`](../scripts/check_assumption_release_guardrails.py)
  as a release gate.
- Keep branch protection and CODEOWNERS aligned to ensure review quality on assumption changes.
