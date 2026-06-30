import crypto from "node:crypto";
import fs from "node:fs/promises";
import path from "node:path";
import { fileURLToPath } from "node:url";
import { SpreadsheetFile, Workbook } from "@oai/artifact-tool";

const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);
const legacyReleaseDir = path.resolve(__dirname, "..");
const repoRoot = path.resolve(legacyReleaseDir, "..");
const releaseDir = path.join(repoRoot, "consultant_safe_v2");

const SOURCE_VERSION = "2026-06-15";
const RELEASE_NAME = "consultant_safe_v2";
const METRIC_EBXS = new Set([
  "EBX-Q-007",
  "EBX-Q-011",
  "EBX-Q-015",
  "EBX-Q-019",
  "EBX-Q-023",
  "EBX-Q-027",
]);
const EXPECTED_SIZES = ["대기업", "중견", "중소", "비상장"];
const STYLE_KEYS = ["formal", "concise", "evidence-led", "narrative"];
const SAFE_HEADERS = [
  "ebx",
  "area",
  "pillar",
  "item",
  "공개성",
  "작성 소재/Content Slots",
  "작성지침",
  "문체 옵션/Style Options",
  "문장 패턴/Sentence Patterns",
  "반복 방지/Anti-Repetition Rules",
  "동종산업 범위 참고",
  "동종산업 재서술 참고",
  "규모지침",
  "비고",
];

const SCALE_GUIDANCE = {
  대기업: [
    "대기업: describe governance, policy, consolidation scope, KPI, timeline, and accountable function in enough detail for reviewer traceability.",
    "Use public evidence when available, then reconcile it with internal source documents before submission.",
  ],
  중견: [
    "중견: focus on material policy/activity, applicable scope, owner, and 1-3 KPIs currently tracked.",
    "If the system is still being standardized, state the roadmap and what remains incomplete instead of overclaiming.",
  ],
  중소: [
    "중소: a concise qualitative answer is acceptable; prioritize operating process, owner, current status, incidents, and internal evidence.",
    "If quantitative data is unavailable, write 미집계 or 해당사항 없음; do not estimate.",
  ],
  비상장: [
    "비상장: emphasize internal evidence and due-diligence readiness; do not assume public-company disclosure duties.",
    "Clearly separate non-public internal scope from externally disclosed scope; use metrics only when internally verifiable.",
  ],
};

const METRIC_SCALE_GUIDANCE = {
  대기업: "Metric rows: include measurement scope, period, formula/source system, and owner; do not leave blank or convert untracked data into 0.",
  중견: "Metric rows: state current KPI and tracking scope; if the series is incomplete, describe the missing part instead of entering 0.",
  중소: "Metric rows: if not tracked, write 미집계 or 해당사항 없음; do not leave blank, estimate, or enter 0 for missing data.",
  비상장: "Metric rows: only enter numbers supported by internal evidence; for due diligence, state source and never enter 0 for untracked data.",
};

async function exists(dir) {
  try {
    const stat = await fs.stat(dir);
    return stat.isDirectory();
  } catch {
    return false;
  }
}

async function resolveSourceDir() {
  const candidates = [
    process.env.EBX_SOURCE_DIR,
    path.join(repoRoot, "ebx_template_review_44"),
    path.join(repoRoot, "..", "template-gen-44-new", "ebx_template_review_44"),
    path.join(repoRoot, "..", "template-gen-44", "ebx_template_review_44"),
    path.join(repoRoot, "..", "template-gen-44", "ebx_template_review_44_vi"),
  ].filter(Boolean);

  for (const candidate of candidates) {
    const resolved = path.resolve(candidate);
    if (await exists(resolved)) return resolved;
  }

  throw new Error(`Source directory not found. Set EBX_SOURCE_DIR or restore ${path.join(repoRoot, "ebx_template_review_44")}.`);
}

function parseCsv(text) {
  const rows = [];
  let row = [];
  let cur = "";
  let quoted = false;
  for (let i = 0; i < text.length; i += 1) {
    const ch = text[i];
    if (ch === "\"") {
      if (quoted && text[i + 1] === "\"") {
        cur += "\"";
        i += 1;
      } else {
        quoted = !quoted;
      }
    } else if (ch === "," && !quoted) {
      row.push(cur);
      cur = "";
    } else if ((ch === "\n" || ch === "\r") && !quoted) {
      if (ch === "\r" && text[i + 1] === "\n") i += 1;
      row.push(cur);
      if (row.some((value) => value !== "")) rows.push(row);
      row = [];
      cur = "";
    } else {
      cur += ch;
    }
  }
  if (cur || row.length) {
    row.push(cur);
    rows.push(row);
  }
  const headers = rows[0].map((header) => header.replace(/^\uFEFF/, ""));
  return rows.slice(1).map((values) =>
    Object.fromEntries(headers.map((header, index) => [header, values[index] ?? ""])),
  );
}

function bulletList(values) {
  if (!values) return "";
  if (Array.isArray(values)) return values.map((value) => `• ${value}`).join("\n");
  return String(values);
}

function hashText(text) {
  return crypto.createHash("sha256").update(text, "utf8").digest("hex");
}

function maskNumerals(text) {
  return String(text ?? "")
    .replace(/\d+(?:[.,]\d+)?(?:\s?[%건명개회월년시간원억조])?/g, "[ ]")
    .replace(/\[\s+\]/g, "[ ]");
}

function safeExample(value) {
  const text = Array.isArray(value) ? value.join("\n") : String(value ?? "");
  return maskNumerals(text);
}

function safeReword(value) {
  const text = String(value ?? "");
  if (!text || text.startsWith("(모델 재서술 미채택") || text.startsWith("(동종 실제사례 미확보")) {
    return "(Reference unavailable: use structure/guidance only and fill with internal company data.)";
  }
  return maskNumerals(text.replace(/\(예시·실데이터 기반 재서술·수치 마스킹\)/g, "(reference only, numbers masked)"));
}

function filePrefixFromMeta(meta) {
  return `${meta.sasb_sector}_${meta.sasb_name}_${meta.size}`;
}

function ebxNumber(item) {
  const match = String(item.ebx ?? "").match(/(\d+)$/);
  return match ? Number(match[1]) : 0;
}

function classifyItem(item) {
  const text = `${item.ebx} ${item.pillar ?? ""} ${item.item ?? ""} ${item.description ?? ""}`;
  if (METRIC_EBXS.has(item.ebx) || /현황|지표|성과|사고|고충|처리/.test(text)) return "metrics";
  if (/전략|비전|목표|로드맵|중장기/.test(text)) return "strategy";
  if (/거버넌스|조직|책임|의사결정|위원회|운영 체계/.test(text)) return "governance";
  if (/리스크|침해|정보보안|개인정보|안전|환경/.test(text)) return "risk";
  if (/정책|방침|관리|프로세스|교육|품질|윤리/.test(text)) return "policy";
  return "general";
}

function extractPlaceholders(item) {
  const draft = String(item["빈칸초안"] ?? "");
  return [...new Set(draft.match(/\[[^\]]+\]/g) ?? [])];
}

function buildWarningNote(meta, item, topicFlags) {
  const notes = [];
  if (meta.mapping_confidence === "low" || meta.fallback_used) {
    notes.push("⚠ Mapping confidence low/fallback: consultant must confirm sector before filling.");
  }
  const topicFlag = topicFlags.get(`${meta.sasb_sector}|${item.ebx}`);
  if (topicFlag?.verdict === "borderline") {
    notes.push("⚠ Review note: industry reference was marked borderline; use only as structure, not as a model answer.");
  }
  if (topicFlag?.verdict === "미확보") {
    notes.push("⚠ Review note: real example was unavailable; prioritize structure and internal evidence.");
  }
  return notes.join(" / ");
}

function buildGuidance(item) {
  const lines = Array.isArray(item["작성지침"]) ? [...item["작성지침"]] : [String(item["작성지침"] ?? "")];
  lines.unshift("Do not write the final answer by copying one fixed sentence. Choose a style option, then compose from slots and verified evidence.");
  if (METRIC_EBXS.has(item.ebx)) {
    lines.push("If metric data is not tracked, write 미집계 or 해당사항 없음; do not estimate or invent figures.");
  }
  return bulletList(lines.filter(Boolean));
}

function buildContentSlots(item) {
  const outline = Array.isArray(item["구조개요"]) ? item["구조개요"] : [String(item["구조개요"] ?? "")].filter(Boolean);
  const placeholders = extractPlaceholders(item);
  const lines = [
    `Disclosure item: ${item.item}`,
    ...outline.map((value, index) => `Required material ${index + 1}: ${value}`),
    placeholders.length
      ? `Company data slots to fill: ${placeholders.join(", ")}`
      : "Company data slots to fill: internal fact, period, scope, owner, and evidence source.",
    "Evidence slot: internal document/source name, reporting period, organizational boundary, responsible team/person.",
  ];
  if (METRIC_EBXS.has(item.ebx)) {
    lines.push("Metric slot: value, unit, aggregation method, measurement period, and reason if 미집계/해당사항 없음.");
  }
  lines.push("Do not convert these slots into a final sentence until company-specific facts are supplied.");
  return bulletList(lines);
}

function buildStyleOptions(meta, item) {
  const kind = classifyItem(item);
  const size = meta.size;
  const detailBias = {
    대기업: "full governance/KPI/evidence detail",
    중견: "policy, owner, key KPI, and improvement roadmap",
    중소: "short qualitative facts with owner and evidence",
    비상장: "internal evidence and due-diligence readiness",
  }[size] ?? "company-specific evidence";

  const lines = [
    `formal: ESG report tone; cover ${detailBias}; suitable when source evidence is complete.`,
    "concise: 1-2 compact sentences; use when the company has limited evidence or the item only needs a factual status.",
    "evidence-led: start with reporting period/source/scope, then explain the policy, activity, or metric; best for reviewer traceability.",
    "narrative: start from business context or materiality, then connect to governance, action, and next step; best for strategy/policy items.",
  ];
  if (kind === "metrics") {
    lines[2] = "evidence-led: start with period, boundary, unit, and source system, then state value/status; if unavailable, state 미집계/해당사항 없음.";
  }
  if (kind === "risk") {
    lines[3] = "narrative: start from risk exposure/context, then explain control, monitoring, incident status, and improvement action.";
  }
  return bulletList(lines);
}

function buildSentencePatterns(item) {
  const kind = classifyItem(item);
  const tag = `${item.ebx} ${item.item}`;
  const common = [
    `Pattern A (${tag}): [기준연도/범위] 기준 [핵심 사실]을 먼저 제시하고 [근거/owner]로 연결한다.`,
    `Pattern B (${tag}): [배경/중요성] -> [정책/프로세스] -> [실행/성과] 순서로 구성한다.`,
    `Pattern C (${tag}): [담당 조직/책임자]를 주어로 시작하고 [활동] [주기] [개선 계획]을 이어간다.`,
  ];
  const byKind = {
    metrics: `Pattern D (${tag}): [측정값/미집계 상태] -> [산식/단위] -> [집계 범위] -> [검증 근거] 순서로 쓴다.`,
    strategy: `Pattern D (${tag}): [비전/목표] -> [중장기 과제] -> [로드맵] -> [성과 점검 방식] 순서로 쓴다.`,
    governance: `Pattern D (${tag}): [의사결정 기구] -> [역할/R&R] -> [보고 주기] -> [최고책임자 관여] 순서로 쓴다.`,
    risk: `Pattern D (${tag}): [식별된 리스크] -> [통제 활동] -> [모니터링/사고 현황] -> [후속 조치] 순서로 쓴다.`,
    policy: `Pattern D (${tag}): [정책/방침] -> [적용 범위] -> [운영 절차] -> [교육/점검] 순서로 쓴다.`,
    general: `Pattern D (${tag}): [현황] -> [운영 방식] -> [증빙] -> [개선 또는 다음 단계] 순서로 쓴다.`,
  };
  return bulletList([...common, byKind[kind]]);
}

function buildAntiRepeatRules(item) {
  const orders = [
    "evidence -> scope -> action -> owner",
    "context -> policy -> activity -> evidence",
    "owner -> process -> KPI/status -> next step",
    "risk/materiality -> control -> monitoring -> result",
  ];
  const openers = [
    "start with reporting period or source",
    "start with responsible organization",
    "start with materiality/context",
    "start with current status or metric",
  ];
  const index = ebxNumber(item);
  return bulletList([
    `Preferred variation for this row: ${orders[index % orders.length]}; avoid using the same order as the previous item.`,
    `Opening choice: ${openers[index % openers.length]}; do not start every answer with 회사는/당사는.`,
    "Limit repeated endings such as 추진하고 있습니다, 관리하고 있습니다, 운영하고 있습니다 across nearby rows.",
    "Change sentence length: mix one short factual sentence with one evidence/detail sentence when enough data exists.",
    "Never copy 동종산업 references as wording; they are only topic/scope references.",
  ]);
}

function buildScaleGuidance(meta, item) {
  const sizeGuidance = SCALE_GUIDANCE[meta.size];
  if (!sizeGuidance) {
    throw new Error(`Unsupported size "${meta.size}" in ${meta.sasb_sector}.`);
  }
  const lines = [...sizeGuidance];
  if (METRIC_EBXS.has(item.ebx)) {
    lines.push(METRIC_SCALE_GUIDANCE[meta.size]);
  }
  return bulletList(lines);
}

function buildScaleGuideSummary(size) {
  const lines = SCALE_GUIDANCE[size];
  if (!lines) {
    throw new Error(`Unsupported size "${size}".`);
  }
  return `${size}. ${lines.join(" ")}`;
}

async function writeMarkdownDocs(summary) {
  const readme = `# Consultant-safe EBX-Q Template v2

## Purpose
This release is for ESG qualitative input using EBX-Q. It keeps the consultant-safe data controls from v1, but removes the single fixed draft sentence as the main writing anchor. AI and consultants should compose from content slots, style options, sentence patterns, and anti-repetition rules.

## Quick Use
1. Choose the workbook by sector and company size.
2. In sheet \`EBX-Q 템플릿\`, use \`작성 소재/Content Slots\` as the required facts to collect. Do not treat it as final prose.
3. Pick one of the four \`문체 옵션/Style Options\`: formal, concise, evidence-led, or narrative.
4. Use \`문장 패턴/Sentence Patterns\` only as a structural pattern. Do not copy the same opening or sentence order across all 27 items.
5. Use \`동종산업 범위 참고\` and \`동종산업 재서술 참고\` only to understand topic scope, not as company facts or wording.
6. For missing metric data, write \`미집계\` or \`해당사항 없음\`; never leave blank, estimate, or enter \`0\` for unavailable data.
7. Before reviewer handoff, run the checklist in \`CHECKLIST.md\`.

## Release Contents
- 44 consultant-safe v2 workbooks: 11 sectors x 4 company sizes.
- The old \`빈칸초안\` answer-draft column is replaced by slots, style options, sentence patterns, and anti-repetition rules.
- The real-data column from source remains excluded.
- Scale-specific guidance remains enforced for 대기업, 중견, 중소, and 비상장.
- QA now checks style-option coverage and duplicate sentence-pattern text within each workbook.

## Build Stats
- Workbooks: ${summary.workbooks}
- Source JSON: ${summary.sourceJson}
- Scale guidance sectors checked: ${summary.scaleGuidanceChecked}
- QA fatal issues: ${summary.fatalIssues}
- QA warnings: ${summary.warningIssues}
`;

  const checklist = `# Checklist Before Reviewer Handoff

## Workbook-level
- [ ] Correct sector and company size workbook selected.
- [ ] \`작성 소재/Content Slots\` used as data requirements, not copied as final prose.
- [ ] At least two different style options are used across the workbook where evidence allows.
- [ ] The same opening, same sentence order, and repeated endings are not used across all 27 items.
- [ ] Source facts come from internal evidence; no invented numbers, targets, certificates, organizations, or processes.

## EBX-Q item-level
- [ ] Every required slot is either filled with verified company data or marked \`미집계/해당사항 없음\`.
- [ ] Metric rows Q-007/Q-011/Q-015/Q-019/Q-023/Q-027 do not use \`0\` for missing data.
- [ ] The chosen style option fits the available evidence: formal for complete evidence, concise for limited evidence, evidence-led for auditable data, narrative for strategy/policy context.
- [ ] 동종산업 reference columns are used only for topic scope and are not copied as wording.
- [ ] Rows with warnings in \`비고\` are reviewed carefully.

## Final handoff
- [ ] Reviewer checked all rows with \`⚠\`.
- [ ] All figures and claims are traceable to internal source documents.
- [ ] Company names, titles, certifications, and reporting periods match official internal records.
- [ ] The filled workbook is saved as a customer/company copy, not over the template.
`;

  const changelog = `# Version Log

## ${RELEASE_NAME}
- Source: \`ebx_template_review_44\`, version \`${SOURCE_VERSION}\`.
- Builds a new \`consultant_safe_v2\` release without overwriting \`consultant_safe_v1\`.
- Replaced the single \`빈칸초안\` writing anchor with content slots, four style options, sentence patterns, and anti-repetition rules.
- Kept consultant-safe controls: removed real-data source column, masked numbers in reference columns, and preserved metric missing-data safeguards.
- Added QA checks for required style keys and duplicate sentence-pattern text within a workbook.
`;

  await fs.writeFile(path.join(releaseDir, "README.md"), readme, "utf8");
  await fs.writeFile(path.join(releaseDir, "CHECKLIST.md"), checklist, "utf8");
  await fs.writeFile(path.join(releaseDir, "CHANGELOG.md"), changelog, "utf8");
}

function createGuideRows(meta, fileName) {
  const mappingText = meta.mapping_confidence === "low" || meta.fallback_used
    ? `산업='${meta.industry}' / 주요산업='${meta.business}' -> SASB ${meta.sasb_sector}(${meta.sasb_name}), confidence=low/fallback. Consultant must confirm sector before filling.`
    : `산업='${meta.industry}' / 주요산업='${meta.business}' -> SASB ${meta.sasb_sector}(${meta.sasb_name}), confidence=${meta.mapping_confidence}.`;

  return [
    [`Consultant-safe EBX-Q Template v2 - ${filePrefixFromMeta(meta)}`, ""],
    ["Source", `Built from ${fileName}; source version ${SOURCE_VERSION}. Source files are not modified.`],
    ["Purpose", "Safe workbook for ESG qualitative input. V2 prevents same-style AI prose by using slots, style options, sentence patterns, and anti-repetition rules."],
    ["How to fill slots", "Fill with verified internal facts. Do not invent figures, targets, organizations, certificates, or processes."],
    ["Style selection", "Choose formal, concise, evidence-led, or narrative per item. Do not use one style for all 27 rows."],
    ["Missing data", "Write 미집계 when not tracked or 해당사항 없음 when not applicable. Do not estimate or enter 0 for unavailable data."],
    ["Industry references", "Reference columns show topic scope only. Do not copy as company facts or prose style."],
    ["Sector mapping", mappingText],
    ["Company size", buildScaleGuideSummary(meta.size)],
    ["Before reviewer", "Read CHECKLIST.md, check all ⚠ notes, and reconcile every figure/claim with internal sources."],
  ];
}

async function buildWorkbook(sourceDir, sourceFile, topicFlags) {
  const sourcePath = path.join(sourceDir, sourceFile);
  const raw = await fs.readFile(sourcePath, "utf8");
  const data = JSON.parse(raw);
  const meta = data._meta;
  const outName = sourceFile.replace(".json", "_consultant_safe_v2.xlsx");
  const outPath = path.join(releaseDir, outName);

  const workbook = Workbook.create();
  const guide = workbook.worksheets.add("안내");
  const sheet = workbook.worksheets.add("EBX-Q 템플릿");

  const guideRows = createGuideRows(meta, sourceFile);
  guide.getRangeByIndexes(0, 0, guideRows.length, 2).values = guideRows;
  guide.getRange("A1:B1").merge();
  guide.getRange("A1").format = {
    fill: "#1F4E79",
    font: { bold: true, color: "#FFFFFF", size: 15 },
  };
  guide.getRange(`A2:A${guideRows.length}`).format = {
    fill: "#D9EAF7",
    font: { bold: true, color: "#17365D" },
  };
  guide.getRange(`A1:B${guideRows.length}`).format.wrapText = true;
  guide.getRange(`A1:B${guideRows.length}`).format.verticalAlignment = "top";
  guide.getRange(`A1:A${guideRows.length}`).format.columnWidthPx = 170;
  guide.getRange(`B1:B${guideRows.length}`).format.columnWidthPx = 840;
  guide.freezePanes.freezeRows(1);
  guide.showGridLines = false;

  const scaleGuidanceByEbx = new Map();
  const rowDiagnostics = [];
  const rows = data.items.map((item) => {
    const scaleGuidance = buildScaleGuidance(meta, item);
    const contentSlots = buildContentSlots(item);
    const styleOptions = buildStyleOptions(meta, item);
    const sentencePatterns = buildSentencePatterns(item);
    const antiRepeatRules = buildAntiRepeatRules(item);
    scaleGuidanceByEbx.set(item.ebx, scaleGuidance);
    rowDiagnostics.push({
      ebx: item.ebx,
      legacyDraft: String(item["빈칸초안"] ?? ""),
      contentSlots,
      styleOptions,
      sentencePatterns,
      antiRepeatRules,
    });

    const baseNote = String(item["비고"] ?? "");
    const warning = buildWarningNote(meta, item, topicFlags);
    const note = [baseNote, warning].filter(Boolean).join(" / ");
    return [
      item.ebx,
      item.area,
      item.pillar,
      item.item,
      item["공개성"],
      contentSlots,
      buildGuidance(item),
      styleOptions,
      sentencePatterns,
      antiRepeatRules,
      safeExample(item["동종예시"]),
      safeReword(item["동종실제재서술"]),
      scaleGuidance,
      note,
    ];
  });

  sheet.getRangeByIndexes(0, 0, 1, SAFE_HEADERS.length).values = [SAFE_HEADERS];
  sheet.getRangeByIndexes(1, 0, rows.length, SAFE_HEADERS.length).values = rows;
  sheet.getRange("A1:N1").format = {
    fill: "#1F4E79",
    font: { bold: true, color: "#FFFFFF" },
  };
  sheet.getRange("A1:N28").format.wrapText = true;
  sheet.getRange("A1:N28").format.verticalAlignment = "top";
  sheet.getRange("F1:J28").format.fill = "#FFF2CC";
  sheet.getRange("K1:L28").format.fill = "#E2F0D9";
  sheet.getRange("M1:N28").format.fill = "#FCE4D6";
  sheet.getRange("A1:A28").format.columnWidthPx = 92;
  sheet.getRange("B1:B28").format.columnWidthPx = 72;
  sheet.getRange("C1:C28").format.columnWidthPx = 130;
  sheet.getRange("D1:D28").format.columnWidthPx = 220;
  sheet.getRange("E1:E28").format.columnWidthPx = 98;
  sheet.getRange("F1:F28").format.columnWidthPx = 420;
  sheet.getRange("G1:G28").format.columnWidthPx = 360;
  sheet.getRange("H1:H28").format.columnWidthPx = 410;
  sheet.getRange("I1:I28").format.columnWidthPx = 470;
  sheet.getRange("J1:J28").format.columnWidthPx = 390;
  sheet.getRange("K1:K28").format.columnWidthPx = 360;
  sheet.getRange("L1:L28").format.columnWidthPx = 360;
  sheet.getRange("M1:M28").format.columnWidthPx = 390;
  sheet.getRange("N1:N28").format.columnWidthPx = 320;
  sheet.freezePanes.freezeRows(1);
  sheet.freezePanes.freezeColumns(4);
  sheet.showGridLines = false;

  const output = await SpreadsheetFile.exportXlsx(workbook);
  await output.save(outPath);
  return {
    outName,
    outPath,
    sourceFile,
    meta,
    rowCount: rows.length,
    rowDiagnostics,
    scaleGuidanceHash: hashText([...scaleGuidanceByEbx.entries()].map(([ebx, guidance]) => `${ebx}:${guidance}`).join("\n")),
    metricGuidance: scaleGuidanceByEbx.get("EBX-Q-007") ?? "",
  };
}

async function main() {
  await fs.mkdir(releaseDir, { recursive: true });
  const sourceDir = await resolveSourceDir();
  const sourceFiles = (await fs.readdir(sourceDir))
    .filter((file) => file.endsWith(".json") && !file.startsWith("_"))
    .sort();
  const topicRows = parseCsv(await fs.readFile(path.join(sourceDir, "_topic_review.csv"), "utf8"));
  const topicFlags = new Map(topicRows.map((row) => [`${row.sector}|${row.ebx}`, row]));
  const outputs = [];

  for (const file of sourceFiles) {
    outputs.push(await buildWorkbook(sourceDir, file, topicFlags));
  }

  const indexRows = [
    "file,sector_code,sector_name,size,mapping_confidence,source_json,rows,sv_mapping_warning,scale_guidance_hash",
    ...outputs.map(({ outName, sourceFile, meta, rowCount, scaleGuidanceHash }) => [
      outName,
      meta.sasb_sector,
      meta.sasb_name,
      meta.size,
      meta.mapping_confidence,
      sourceFile,
      rowCount,
      meta.mapping_confidence === "low" || meta.fallback_used ? "yes" : "no",
      scaleGuidanceHash,
    ].map((value) => `"${String(value).replaceAll("\"", "\"\"")}"`).join(",")),
  ];
  await fs.writeFile(path.join(releaseDir, "RELEASE_INDEX.csv"), `${indexRows.join("\n")}\n`, "utf8");

  const qa = await runQa(sourceDir, outputs);
  await fs.writeFile(path.join(releaseDir, "QA_REPORT.json"), `${JSON.stringify(qa, null, 2)}\n`, "utf8");
  await writeMarkdownDocs({
    workbooks: outputs.length,
    sourceJson: sourceFiles.length,
    scaleGuidanceChecked: qa.scaleGuidance.checkedSectors,
    fatalIssues: qa.fatal.length,
    warningIssues: qa.warnings.length,
  });

  console.log(JSON.stringify({
    releaseDir,
    sourceDir,
    workbooks: outputs.length,
    fatal: qa.fatal.length,
    warnings: qa.warnings.length,
  }, null, 2));
  if (qa.fatal.length) {
    process.exitCode = 1;
  }
}

async function runQa(sourceDir, outputs) {
  const fatal = [];
  const warnings = [];
  const seenFiles = new Set(outputs.map((output) => output.outName));
  if (seenFiles.size !== 44) fatal.push(`Expected 44 workbooks, found ${seenFiles.size}.`);
  if (SAFE_HEADERS.includes("빈칸초안")) fatal.push("V2 headers must not include the old 빈칸초안 column.");

  for (const output of outputs) {
    const sourceData = JSON.parse(await fs.readFile(path.join(sourceDir, output.sourceFile), "utf8"));
    if (sourceData.items.length !== 27) fatal.push(`${output.outName}: source item count is ${sourceData.items.length}, expected 27.`);
    const ids = sourceData.items.map((item) => item.ebx);
    if (new Set(ids).size !== 27) fatal.push(`${output.outName}: duplicate EBX ids.`);
    const expectedIds = Array.from({ length: 27 }, (_, index) => `EBX-Q-${String(index + 1).padStart(3, "0")}`);
    for (const id of expectedIds) {
      if (!ids.includes(id)) fatal.push(`${output.outName}: missing ${id}.`);
    }
    const jsonText = JSON.stringify(sourceData);
    if (!jsonText.includes("[ ]")) warnings.push(`${output.outName}: source JSON has no [ ] placeholder, unusual.`);
    if (!EXPECTED_SIZES.includes(output.meta.size)) {
      fatal.push(`${output.outName}: unsupported size ${output.meta.size}.`);
    }
    if (!output.metricGuidance.includes("0") || !/do not leave blank|do not enter 0|instead of entering 0|never enter 0/.test(output.metricGuidance)) {
      fatal.push(`${output.outName}: EBX-Q-007 metric scale guidance does not warn against blank/zero substitution.`);
    }
    if ((output.meta.mapping_confidence === "low" || output.meta.fallback_used) && output.meta.sasb_sector !== "SV") {
      warnings.push(`${output.outName}: low/fallback mapping outside SV.`);
    }

    const patternTexts = output.rowDiagnostics.map((row) => row.sentencePatterns);
    const uniquePatternTexts = new Set(patternTexts);
    if (uniquePatternTexts.size !== output.rowDiagnostics.length) {
      fatal.push(`${output.outName}: duplicate sentence pattern text within workbook.`);
    }
    for (const row of output.rowDiagnostics) {
      const missingStyles = STYLE_KEYS.filter((style) => !row.styleOptions.includes(`${style}:`));
      if (missingStyles.length) {
        fatal.push(`${output.outName} ${row.ebx}: missing style options ${missingStyles.join(", ")}.`);
      }
      if (row.legacyDraft && row.contentSlots.trim() === row.legacyDraft.trim()) {
        fatal.push(`${output.outName} ${row.ebx}: content slots still equal old 빈칸초안 draft.`);
      }
      if (!row.contentSlots.includes("Required material") || !row.contentSlots.includes("Evidence slot")) {
        fatal.push(`${output.outName} ${row.ebx}: content slots are incomplete.`);
      }
    }
  }

  const bySector = new Map();
  for (const output of outputs) {
    if (!bySector.has(output.meta.sasb_sector)) bySector.set(output.meta.sasb_sector, []);
    bySector.get(output.meta.sasb_sector).push(output);
  }

  for (const [sector, sectorOutputs] of bySector.entries()) {
    const sizes = new Set(sectorOutputs.map((output) => output.meta.size));
    for (const size of EXPECTED_SIZES) {
      if (!sizes.has(size)) fatal.push(`${sector}: missing ${size} consultant-safe workbook.`);
    }
    const hashes = new Set(sectorOutputs.map((output) => output.scaleGuidanceHash));
    if (hashes.size !== EXPECTED_SIZES.length) {
      fatal.push(`${sector}: expected 4 distinct scale guidance variants, found ${hashes.size}.`);
    }
    const bySize = Object.fromEntries(sectorOutputs.map((output) => [output.meta.size, output.scaleGuidanceHash]));
    if (bySize["중소"] && bySize["비상장"] && bySize["중소"] === bySize["비상장"]) {
      fatal.push(`${sector}: 중소 and 비상장 scale guidance are identical.`);
    }
    if (bySize["대기업"] && bySize["중견"] && bySize["대기업"] === bySize["중견"]) {
      fatal.push(`${sector}: 대기업 and 중견 scale guidance are identical.`);
    }
  }

  return {
    generatedAt: new Date().toISOString(),
    sourceVersion: SOURCE_VERSION,
    releaseName: RELEASE_NAME,
    expectedHeaders: SAFE_HEADERS,
    styleOptions: {
      required: STYLE_KEYS,
      checkedWorkbooks: outputs.length,
    },
    scaleGuidance: {
      expectedSizes: EXPECTED_SIZES,
      checkedSectors: bySector.size,
    },
    fatal,
    warnings,
  };
}

await main();
