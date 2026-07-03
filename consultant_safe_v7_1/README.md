# Consultant-safe EBX-Q Template v7.1

## Purpose
V7.1 changes the template source-of-truth from Excel workbooks to JSON files.

## What Changed
- Keeps v7 answer controls and customer output schema.
- Converts each sector/size template workbook into a full-column JSON template under `templates/`.
- Preserves the Excel header list in `headers` and row values using the same header names.
- Keeps internal camelCase aliases for generator fields so existing scripts can read the JSON safely.
- Company filling reads JSON templates directly and only creates Excel for final customer delivery.
- Future template improvements should edit JSON, not regenerate 44 Excel workbooks.

## Build Stats
- Source workbooks read once: 44
- JSON templates generated: 44
- Fatal issues: 0
- Warnings: 0
