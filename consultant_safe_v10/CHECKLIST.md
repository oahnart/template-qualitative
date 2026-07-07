# consultant_safe_v10 Handoff Checklist

- [ ] Use `policies/template_profiles.json` as the operational policy source-of-truth and generated JSON files as build artifacts.
- [ ] Rebuild v10 if `source/26.07.06 ESG-정성_v2.xlsx` changes.
- [ ] Static validator confirms 44 templates, 95 rows, 4 distinct size hashes per sector, and 11 distinct sector hashes per size.
- [ ] Four profile fixtures select the exact requested size and reject invalid size values.
- [ ] Eleven large-company regressions have zero fail, fatal, and topic-mismatched metric rows.
- [ ] Customer workbook preserves the six established output columns.
- [ ] Customer workbook has 95 EBX rows.
- [ ] Sidecar QA JSON has zero fatal findings before handoff.
- [ ] Rows without company evidence remain marked UNKNOWN and do not invent facts.
