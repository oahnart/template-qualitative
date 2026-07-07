# Consultant-safe EBX-Q Template v10.1

V10.1 preserves all 95 EBX rows, 44 operational templates, and the six customer workbook columns from v10.

## Improvements
- Prefer exact mapped-item metrics and topic-specific indicators.
- Reject control characters, prose/table fragments, and year-as-value metric artifacts.
- Fail QA on exact duplicate final answers and malformed metric sentences.
- Keep all six columns while bounding row height and wrapping only user-readable content columns.

## Validation
- `node scripts/validate_consultant_safe_v10_1.mjs`
- `node scripts/test_consultant_safe_v10_1_profiles.mjs`
- `node scripts/regress_consultant_safe_v10_1_large.mjs`
