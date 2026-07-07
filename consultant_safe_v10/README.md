# Consultant-safe EBX-Q Template v10

## Purpose
V10 replaces the 27-item qualitative template with the 95-item master workbook `source/26.07.06 ESG-정성_v2.xlsx`.

## Operational Profile Model
- Policy source-of-truth: `policies/template_profiles.json` (template-profiles-v1).
- Matrix: 11 sector overlays x 4 operational profiles (대기업, 중견, 중소, 비상장) = 44 templates.
- 비상장 remains a compatibility profile for unlisted-company disclosure and does not imply company size.
- Policy precedence: master row -> v9 row rules -> topic profile -> sector overlay -> size profile -> explicit row policy override.

## What Changed
- Uses the v10 master workbook as the source-of-truth for EBX-Q-001 through EBX-Q-095.
- Preserves v9 row-level rules for the original 27 items.
- Applies 17 topic profiles to the 68 new items and adds sector- and size-specific operational controls.
- Keeps the six-column customer output schema unchanged.
- Rejects duplicate operational profiles during build.

## Validation
- `node scripts/validate_consultant_safe_v10.mjs`
- `node scripts/test_consultant_safe_v10_profiles.mjs`
- `node scripts/regress_consultant_safe_v10_large.mjs`

## Build Stats
- Source v9 JSON templates read: 44
- v10 JSON templates generated: 44
- Rows per template: 95
- Fatal issues: 0
- Warnings: 1
