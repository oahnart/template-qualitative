# Consultant-safe EBX-Q Template v4

## Purpose
This release keeps the consultant-safe controls from v3, then adds field-path, style-template, and report-ready final answer controls.

## Key Changes
- Builds a new `consultant_safe_v4` release without overwriting `consultant_safe_v3`.
- Adds `Field Path`, `Final Answer Requirements`, `Report-Ready Guardrails`, `Coverage Handling`, and `Metric Handling` to every template workbook.
- Requires company output `Field` to use `area / pillar / item`.
- Requires `Style Template Applied` to describe content slots, style options, sentence patterns, anti-repetition guidance, coverage treatment, and metric handling.
- Requires `Final Answer` to remove source names, page citations, PDF references, reviewer notes, and audit-trace wording.

## Build Stats
- Source workbooks: 44
- V4 workbooks: 44
- Fatal issues: 0
- Warnings: 0
