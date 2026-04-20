# Independent challenge activation checklist

This checklist operationalizes the second-reviewer path already scaffolded
in CODEOWNERS and branch-protection profiles.

## Objective

Enable enforceable independent challenge for Tier 1 model changes.

## Activation steps

1. Create or onboard a second human/team reviewer.
2. Update [`.github/CODEOWNERS`](../../.github/CODEOWNERS):
   - uncomment and populate second-owner entries for parity-critical paths.
3. Apply branch protection profile with required reviews:
   - use [`.github/branch-protection.with-second-reviewer.json`](../../.github/branch-protection.with-second-reviewer.json).
4. Validate with a test PR that touches parity-critical files.
5. Record activation date and approver in governance notes.

## Readiness checks

- Required status checks still match workflow job names.
- Required conversation resolution remains enabled.
- Review-from-CODEOWNERS rule is active on `main`.

## Evidence to retain

- API response from branch protection update command.
- Screenshot/log showing required reviewer enforcement.
- First merged PR under second-reviewer enforcement.
