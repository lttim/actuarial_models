# Assumption release guardrail runbook

This runbook enforces release-time controls around synthetic/placeholder assumptions.

## Command

```bash
cd annuity_model
python scripts/check_assumption_release_guardrails.py
```

## Behavior

- **Passes** if no placeholder assumptions are detected in `data_registry`.
- **Passes with waiver** if placeholders are detected and `.release/assumption_waiver.md` exists.
- **Fails** if placeholders are detected without waiver evidence.

## Waiver flow

1. Copy template:
   `cp docs/release_assumption_waiver.md .release/assumption_waiver.md`
2. Complete all required fields in `.release/assumption_waiver.md`.
3. Obtain explicit approver and challenger signoff.
4. Re-run guardrail command and proceed only if pass criteria are met.

## Operational guidance

- Prefer replacing placeholders with approved artifacts over using waivers.
- Treat waivers as time-bound, with explicit expiry and remediation owner.
