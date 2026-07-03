# Consultant-safe EBX-Q Template v6

## Purpose
V6 keeps the v5 workbook family but replaces duplicated generic control columns with compact decision fields used by deterministic output generation.

## Key Changes
- Builds a new `consultant_safe_v6` release without overwriting `consultant_safe_v5`.
- Adds `Answer Type`, `Evidence Priority`, `Evidence Slots`, `Metric Role`, `Style Rule`, `Sentence Blueprint`, `Forbidden Tokens`, and `QA Severity`.
- Preserves the established output schema: `EBX Indicator`, `Field`, `Original Answer`, `Original Answer Metadata`, `Style Template Applied`, and `Final Answer`.
- Moves QA detail to sidecar JSON reports.

## Build Stats
- Source workbooks: 44
- V6 workbooks: 44
- Fatal issues: 0
- Warnings: 0
