# Consultant-safe EBX-Q Template v4

## Purpose
This release keeps the consultant-safe controls from v3, then adds field-path, style-template, and report-ready final answer controls.

## Key Changes
- Builds a new `consultant_safe_v4` release without overwriting `consultant_safe_v3`.
- Adds `Field Path`, `Final Answer Requirements`, `Report-Ready Guardrails`, `Coverage Handling`, and `Metric Handling` to every template workbook.
- Requires company output `Field` to use `area / pillar / item`.
- Requires `Style Template Applied` to describe the selected style option, selected sentence pattern, applied content slots, anti-repetition guidance, coverage treatment, and metric handling.
- Requires `Final Answer` to remove source names, page citations, PDF references, reviewer notes, and audit-trace wording.

## Rule Placement
- Template columns hold row-level generation rules used by downstream output: `Field Path`, `Final Answer Requirements`, `Report-Ready Guardrails`, `Coverage Handling`, and `Metric Handling`.
- `README.md` documents the release purpose, generated artifacts, and where QA lives.
- `CHECKLIST.md` is a manual handoff checklist. It is not executable; automated checks live in the build/fill scripts.
- The build writes release workbooks only to `consultant_safe_v4/`; it does not create or refresh `final_template/template_qualitative/`.

## Handoff QA
Before handing off a generated company workbook, review `CHECKLIST.md`.
Use the template columns as the source of row-level output rules, then use the checklist as the final human-readable review layer.

## Build Stats
- Source workbooks: 44
- V4 workbooks: 44
- Fatal issues: 0
- Warnings: 0
