# Consultant-safe EBX-Q Template v9

## Purpose
V9 keeps v8 JSON templates and doubles the Final Answer length policy while keeping the prose evidence-bounded and customer-safe.

## What Changed
- Raises evidence selection from 6 to 12 candidate sentences.
- Raises sufficient-answer QA targets to 8 sentences and 900 Korean characters.
- Doubles coverage-aware answer length targets from v8.
- Allows at least 5 metric records for metric-heavy rows.
- Keeps the six-column customer output schema unchanged.
- Keeps JSON templates as the source of truth for row-level behavior.
- Keeps JavaScript as the execution engine for reading data, applying rules, QA, and exporting Excel.

## Build Stats
- Source v8 JSON templates read: 44
- v9 JSON templates generated: 44
- Fatal issues: 0
- Warnings: 0
