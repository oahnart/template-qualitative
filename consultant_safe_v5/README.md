# Consultant-safe EBX-Q Template v5

## Purpose
This release keeps the v4 report-ready workbook shape and adds deterministic evidence-grounding and QA controls.

## Key Changes
- Builds a new `consultant_safe_v5` release without overwriting `consultant_safe_v4`.
- Adds `Evidence Selection Rules`, `Metric Evidence Requirements`, `Company Naming Rule`, `Anti-Repetition Group`, and `QA Severity`.
- Requires generated company output to use row-specific evidence, supported figures when available, Korean report-facing company names, and non-repetitive final prose.
- Supports v5 output workbooks with a clean answer sheet plus `QA_Audit`.

## Build Stats
- Source workbooks: 44
- V5 workbooks: 44
- Fatal issues: 0
- Warnings: 0
