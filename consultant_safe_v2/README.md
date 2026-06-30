# Consultant-safe EBX-Q Template v2

## Purpose
This release is for ESG qualitative input using EBX-Q. It keeps the consultant-safe data controls from v1, but removes the single fixed draft sentence as the main writing anchor. AI and consultants should compose from content slots, style options, sentence patterns, and anti-repetition rules.

## Quick Use
1. Choose the workbook by sector and company size.
2. In sheet `EBX-Q 템플릿`, use `작성 소재/Content Slots` as the required facts to collect. Do not treat it as final prose.
3. Pick one of the four `문체 옵션/Style Options`: formal, concise, evidence-led, or narrative.
4. Use `문장 패턴/Sentence Patterns` only as a structural pattern. Do not copy the same opening or sentence order across all 27 items.
5. Use `동종산업 범위 참고` and `동종산업 재서술 참고` only to understand topic scope, not as company facts or wording.
6. For missing metric data, write `미집계` or `해당사항 없음`; never leave blank, estimate, or enter `0` for unavailable data.
7. Before reviewer handoff, run the checklist in `CHECKLIST.md`.

## Release Contents
- 44 consultant-safe v2 workbooks: 11 sectors x 4 company sizes.
- The old `빈칸초안` answer-draft column is replaced by slots, style options, sentence patterns, and anti-repetition rules.
- The real-data column from source remains excluded.
- Scale-specific guidance remains enforced for 대기업, 중견, 중소, and 비상장.
- QA now checks style-option coverage and duplicate sentence-pattern text within each workbook.

## Build Stats
- Workbooks: 44
- Source JSON: 44
- Scale guidance sectors checked: 11
- QA fatal issues: 0
- QA warnings: 0
