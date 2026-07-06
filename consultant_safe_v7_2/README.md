# Consultant-safe EBX-Q Template v7.2

## Purpose
V7.2 keeps v7.1 JSON templates and adds row-level rule fields so the fill script can read generation decisions from template data.

## What Changed
- Adds `preferredStyle`, `styleOptions`, `sentencePlan`, `evidenceSelection`, `metricHints`, `qaRules`, and `forbiddenTerms` to every EBX row.
- Keeps the six-column customer output schema unchanged.
- Keeps JSON templates as the source of truth for row-level behavior.
- Keeps JavaScript as a thin execution engine for reading data, applying rules, QA, and exporting Excel.

## Build Stats
- Source v7.1 JSON templates read: 44
- v7.2 JSON templates generated: 44
- Fatal issues: 0
- Warnings: 0
