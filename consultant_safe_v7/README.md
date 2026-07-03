# Consultant-safe EBX-Q Template v7

## Purpose
V7 keeps the v6 workbook family but changes the style-control columns so Final Answer prose is less repetitive and easier for non-specialist clients to read.

## Key Changes
- Builds a new `consultant_safe_v7` release without overwriting `consultant_safe_v6`.
- Replaces metric/formula-oriented controls with `Answer Intent`, `Opening Strategy`, `Evidence Weave`, `Required Facts`, `Plain-Language Avoid List`, `Style Guardrails`, and `QA Severity`.
- Preserves the customer output schema: `EBX Indicator`, `Field`, `Original Answer`, `Original Answer Metadata`, `Style Template Applied`, and `Final Answer`.
- Instructs generators to avoid `quantitative`, `정량`, and `định lượng` in customer-facing answers.

## Build Stats
- Source workbooks: 44
- V7 workbooks: 44
- Fatal issues: 0
- Warnings: 0
