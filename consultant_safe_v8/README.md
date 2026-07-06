# Consultant-safe EBX-Q Template v8

## Purpose
V8 keeps v7.2 JSON templates and adds evidence-bounded length policy so Final Answer prose is fuller without writing beyond available company evidence.

## What Changed
- Raises evidence selection from 4 to 6 candidate sentences.
- Raises sufficient-answer QA targets to 4 sentences and 450 Korean characters.
- Adds `lengthPolicy` to every EBX row for coverage-aware answer length.
- Allows up to 3 metric records for metric-heavy rows.
- Keeps the six-column customer output schema unchanged.
- Keeps JSON templates as the source of truth for row-level behavior.
- Keeps JavaScript as the execution engine for reading data, applying rules, QA, and exporting Excel.

## Build Stats
- Source v7.2 JSON templates read: 44
- v8 JSON templates generated: 44
- Fatal issues: 0
- Warnings: 0
