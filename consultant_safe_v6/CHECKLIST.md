# consultant_safe_v6 Handoff Checklist

- [ ] Customer workbook preserves the established output columns: `EBX Indicator`, `Field`, `Original Answer`, `Original Answer Metadata`, `Style Template Applied`, and `Final Answer`.
- [ ] `Style Template Applied` contains the selected v6 decision fields, not a dump of all style options or sentence patterns.
- [ ] No extra QA sheet appears in the customer workbook; QA details are in the sidecar JSON.
- [ ] Sidecar QA JSON has zero fatal findings.
- [ ] `Final Answer` has no EBX codes, source/page/PDF/reviewer/audit wording, report headers, table/index fragments, or raw OCR labels.
- [ ] Metric-supported rows include supported figures or target/status wording where available.
- [ ] Style selection is not collapsed into one option for all rows unless the template explicitly requires it.
