# Version Log

## consultant_safe_v1
- Source: `ebx_template_review_44`, version `2026-06-15`.
- Created release package for consultant nhập liệu.
- Removed public-facing real-data column: `동종 실제사례(완성문·회사명 마스킹·수치유지)`.
- Renamed reworded example column to `동종산업 재서술 예시`.
- Added concise Vietnamese guidance to `안내` sheets.
- Added scale-specific `규모지침` for `대기업`, `중견`, `중소`, and `비상장`.
- Added QA checks that fail when the 4 scale guidance variants are duplicated inside a sector.
- Added warnings for SV fallback mapping and topic-review borderline/missing items.
- Masked numeric sequences in consultant-facing example columns.
- Added README, checklist, release index, and QA report.
