import fs from "node:fs/promises";
import path from "node:path";
import { fileURLToPath } from "node:url";
import { FileBlob, SpreadsheetFile } from "@oai/artifact-tool";

const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);
const repoRoot = path.resolve(__dirname, "..");
const sourceDir = path.join(repoRoot, "consultant_safe_v5");
const releaseDir = path.join(repoRoot, "consultant_safe_v6");

const RELEASE_NAME = "consultant_safe_v6";
const SOURCE_RELEASE = "consultant_safe_v5";
const TEMPLATE_SHEET = "EBX-Q 템플릿";
const GUIDE_SHEET = "안내";

const BASE_COLUMN_COUNT = 14;
const V6_HEADERS = [
  "Answer Type",
  "Evidence Priority",
  "Evidence Slots",
  "Metric Role",
  "Style Rule",
  "Sentence Blueprint",
  "Forbidden Tokens",
  "QA Severity",
];

async function ensureDir(dir) {
  await fs.mkdir(dir, { recursive: true });
}

async function cleanupInspectSidecar(file) {
  await fs.rm(`${file}.inspect.ndjson`, { force: true });
}

function csvEscape(value) {
  return `"${String(value ?? "").replaceAll('"', '""')}"`;
}

function indexByHeader(headers, name) {
  return headers.findIndex((header) => String(header ?? "") === name);
}

function ebxNumber(row, idx) {
  return Number(String(row[idx.ebx] ?? "").match(/\d+/)?.[0] ?? 0);
}

function answerTypeFor(row, idx) {
  const num = ebxNumber(row, idx);
  if ([1, 8, 16, 20, 24].includes(num)) return "strategy-policy";
  if ([2, 5, 9, 13, 17, 21, 25].includes(num)) return "governance-organization";
  if ([3, 6, 10, 14, 18, 22, 26].includes(num)) return "risk-control";
  if ([7, 11, 15, 19, 23, 27].includes(num)) return "metric-status";
  return "policy-management";
}

function buildFieldPath(row, idx) {
  return [row[idx.area], row[idx.pillar], row[idx.item]]
    .map((value) => String(value ?? "").trim())
    .filter(Boolean)
    .join(" / ");
}

function evidencePriority(type) {
  const priorities = {
    "strategy-policy": "1. company strategy, policy, target, roadmap; 2. governance owner; 3. disclosed KPI or target year.",
    "governance-organization": "1. responsible body or owner; 2. role/R&R; 3. reporting or decision cycle; 4. executive oversight.",
    "risk-control": "1. identified risk; 2. control or prevention action; 3. monitoring process; 4. remediation or follow-up.",
    "metric-status": "1. disclosed metric or incident status; 2. period/unit/boundary; 3. trend; 4. management interpretation.",
    "policy-management": "1. policy or standard; 2. implementation process; 3. governance owner; 4. supported metric if available.",
  };
  return priorities[type] ?? priorities["policy-management"];
}

function metricRole(type) {
  if (type === "metric-status") return "Required when disclosed: include period, unit, boundary, and trend; never convert blanks to zero.";
  if (type === "strategy-policy" || type === "policy-management") return "Use supported targets or KPI values when available; otherwise keep the paragraph qualitative.";
  return "Optional supporting evidence only; metric presence must not force the style to evidence-led.";
}

function styleRule(type) {
  const rules = {
    "strategy-policy": "narrative: start with business direction or policy context, then target, execution, and review.",
    "governance-organization": "governance: start with the responsible body, then role, reporting flow, and oversight.",
    "risk-control": "risk-control: start with risk or control scope, then prevention, monitoring, and follow-up.",
    "metric-status": "metric-led: start with disclosed figure/status when available, then explain boundary and management meaning.",
    "policy-management": "balanced-policy: start with policy/standard, then operating process and performance evidence.",
  };
  return rules[type] ?? rules["policy-management"];
}

function sentenceBlueprint(type) {
  const blueprints = {
    "strategy-policy": "[company policy/strategy] -> [target or roadmap] -> [execution body/activity] -> [performance review or limitation].",
    "governance-organization": "[responsible body] -> [role/R&R] -> [reporting or decision cadence] -> [executive/board oversight].",
    "risk-control": "[risk/control scope] -> [prevention or mitigation activity] -> [monitoring] -> [corrective action or next control].",
    "metric-status": "[metric/incident status] -> [period/unit/boundary] -> [trend or result] -> [management interpretation].",
    "policy-management": "[policy/standard] -> [implementation process] -> [responsible owner] -> [supported performance fact].",
  };
  return blueprints[type] ?? blueprints["policy-management"];
}

function qaSeverity(type) {
  if (type === "metric-status") return "High: figures/status must be supported and no missing-value-to-zero conversion is allowed.";
  if (type === "risk-control") return "High: risk rows must include control and monitoring evidence.";
  if (type === "governance-organization") return "Medium: responsible body and decision flow must be clear.";
  return "Medium: final prose must be report-ready, evidence-grounded, and non-repetitive.";
}

function evidenceSlots(row, idx, type) {
  const fieldPath = buildFieldPath(row, idx);
  return [
    `Field Path: ${fieldPath}`,
    `Answer Type: ${type}`,
    "Use 2-4 row-specific evidence sentences before adding connective prose.",
    "Reject source titles, page labels, OCR navigation, table/index rows, template instructions, and incomplete metric fragments.",
    "Repair obvious OCR spacing before final output; do not pass broken Korean spacing into Final Answer.",
    "Keep source traces outside the customer workbook.",
  ].join("\n");
}

function forbiddenTokens() {
  return [
    "EBX codes",
    "source/PDF/page/citation/audit/reviewer wording",
    "report navigation headers",
    "GRI/index/table labels",
    "Step labels and raw OCR fragments",
    "incomplete metric fragments such as sentences starting with '등이 확인되어'",
    "table/list headings such as 구분, 교육 내용, 점검 시기, Topic No.",
    "broken OCR spacing such as 사 항, 구 현, 직 접, 집 행",
    "awkward fallback grammar such as 책임와, 조직와, 정책를",
    "English/underscore company fallback names",
  ].join("; ");
}

function v6Enhancements(row, idx) {
  const type = answerTypeFor(row, idx);
  return [
    type,
    evidencePriority(type),
    evidenceSlots(row, idx, type),
    metricRole(type),
    styleRule(type),
    sentenceBlueprint(type),
    forbiddenTokens(),
    qaSeverity(type),
  ];
}

function guideRows(title, sourceFile) {
  return [
    [`Consultant-safe EBX-Q Template v6 - ${title}`, ""],
    ["Source", `Built from ${sourceFile}; source release ${SOURCE_RELEASE}.`],
    ["Purpose", "V6 replaces broad v5 control text with row-level decision fields while preserving the established output workbook columns."],
    ["Customer output", "Generated output keeps the v5 column schema: EBX Indicator, Field, Original Answer, Original Answer Metadata, Style Template Applied, and Final Answer."],
    ["QA", "QA is written as a sidecar JSON report, not a workbook sheet, and treats EBX codes, source traces, OCR/table artifacts, and metric misses as fatal."],
    ["Style selection", "Style is selected from row intent and template answer type; metric support can add a metric sentence but must not force all rows to one style."],
  ];
}

async function writeDocs(summary) {
  const readme = `# Consultant-safe EBX-Q Template v6

## Purpose
V6 keeps the v5 workbook family but replaces duplicated generic control columns with compact decision fields used by deterministic output generation.

## Key Changes
- Builds a new \`${RELEASE_NAME}\` release without overwriting \`${SOURCE_RELEASE}\`.
- Adds \`Answer Type\`, \`Evidence Priority\`, \`Evidence Slots\`, \`Metric Role\`, \`Style Rule\`, \`Sentence Blueprint\`, \`Forbidden Tokens\`, and \`QA Severity\`.
- Preserves the established output schema: \`EBX Indicator\`, \`Field\`, \`Original Answer\`, \`Original Answer Metadata\`, \`Style Template Applied\`, and \`Final Answer\`.
- Moves QA detail to sidecar JSON reports.

## Build Stats
- Source workbooks: ${summary.sourceWorkbooks}
- V6 workbooks: ${summary.v6Workbooks}
- Fatal issues: ${summary.fatal.length}
- Warnings: ${summary.warnings.length}
`;

  const checklist = `# consultant_safe_v6 Handoff Checklist

- [ ] Customer workbook preserves the established output columns: \`EBX Indicator\`, \`Field\`, \`Original Answer\`, \`Original Answer Metadata\`, \`Style Template Applied\`, and \`Final Answer\`.
- [ ] \`Style Template Applied\` contains the selected v6 decision fields, not a dump of all style options or sentence patterns.
- [ ] No extra QA sheet appears in the customer workbook; QA details are in the sidecar JSON.
- [ ] Sidecar QA JSON has zero fatal findings.
- [ ] \`Final Answer\` has no EBX codes, source/page/PDF/reviewer/audit wording, report headers, table/index fragments, or raw OCR labels.
- [ ] Metric-supported rows include supported figures or target/status wording where available.
- [ ] Style selection is not collapsed into one option for all rows unless the template explicitly requires it.
`;

  const changelog = `# Version Log

## ${RELEASE_NAME}
- Source release: \`${SOURCE_RELEASE}\`.
- Replaces duplicated v5 control columns with decision-oriented v6 fields.
- Creates v6 workbook filenames ending in \`_consultant_safe_v6.xlsx\`.
- Preserves the established output workbook schema and uses sidecar JSON QA.
`;

  await fs.writeFile(path.join(releaseDir, "README.md"), readme, "utf8");
  await fs.writeFile(path.join(releaseDir, "CHECKLIST.md"), checklist, "utf8");
  await fs.writeFile(path.join(releaseDir, "CHANGELOG.md"), changelog, "utf8");
}

async function buildWorkbook(sourceFile) {
  const inputPath = path.join(sourceDir, sourceFile);
  const outName = sourceFile.replace("_consultant_safe_v5.xlsx", "_consultant_safe_v6.xlsx");
  const outPath = path.join(releaseDir, outName);
  const workbook = await SpreadsheetFile.importXlsx(await FileBlob.load(inputPath));

  const guide = workbook.worksheets.getItem(GUIDE_SHEET);
  const title = outName
    .replace(/^EBX_Q_템플릿_/, "")
    .replace("_consultant_safe_v6.xlsx", "");
  const guideValues = guideRows(title, sourceFile);
  guide.getRangeByIndexes(0, 0, guideValues.length, 2).values = guideValues;
  guide.getRange(`A1:B${guideValues.length}`).format.wrapText = true;
  guide.getRange(`A1:B${guideValues.length}`).format.verticalAlignment = "top";
  guide.getRange("A1:B1").merge();
  guide.getRange("A1").format = {
    fill: "#174A5A",
    font: { bold: true, color: "#FFFFFF", size: 15 },
  };
  guide.getRange(`A2:A${guideValues.length}`).format = {
    fill: "#DCECF0",
    font: { bold: true, color: "#143743" },
  };

  const sheet = workbook.worksheets.getItem(TEMPLATE_SHEET);
  const sourceValues = sheet.getRange("A1:X28").values;
  const sourceHeaders = sourceValues[0];
  const idx = {
    ebx: indexByHeader(sourceHeaders, "ebx"),
    area: indexByHeader(sourceHeaders, "area"),
    pillar: indexByHeader(sourceHeaders, "pillar"),
    item: indexByHeader(sourceHeaders, "item"),
  };

  const baseHeaders = sourceHeaders.slice(0, BASE_COLUMN_COUNT);
  const baseRows = sourceValues.slice(1).map((row) => row.slice(0, BASE_COLUMN_COUNT));
  const outputHeaders = [...baseHeaders, ...V6_HEADERS];
  const outputRows = baseRows.map((row, index) => [
    ...row,
    ...v6Enhancements(sourceValues[index + 1], idx),
  ]);

  sheet.getRangeByIndexes(0, 0, 1, outputHeaders.length).values = [outputHeaders];
  sheet.getRangeByIndexes(1, 0, outputRows.length, outputHeaders.length).values = outputRows;
  if (sourceHeaders.length > outputHeaders.length) {
    const clearWidth = sourceHeaders.length - outputHeaders.length;
    const blanks = Array.from({ length: outputRows.length + 1 }, () => Array(clearWidth).fill(""));
    sheet.getRangeByIndexes(0, outputHeaders.length, outputRows.length + 1, clearWidth).values = blanks;
  }
  sheet.getRangeByIndexes(0, 0, outputRows.length + 1, outputHeaders.length).format.wrapText = true;
  sheet.getRangeByIndexes(0, 0, outputRows.length + 1, outputHeaders.length).format.verticalAlignment = "top";
  sheet.getRangeByIndexes(0, BASE_COLUMN_COUNT, 1, V6_HEADERS.length).format = {
    fill: "#174A5A",
    font: { bold: true, color: "#FFFFFF" },
  };
  sheet.getRange("O:V").format.columnWidthPx = 320;

  const exported = await SpreadsheetFile.exportXlsx(workbook);
  await exported.save(outPath);
  await cleanupInspectSidecar(outPath);
  return { sourceFile, outName };
}

async function main() {
  await ensureDir(releaseDir);
  const sourceFiles = (await fs.readdir(sourceDir))
    .filter((file) => file.endsWith("_consultant_safe_v5.xlsx"))
    .sort();

  const outputs = [];
  for (const file of sourceFiles) {
    outputs.push(await buildWorkbook(file));
  }

  const fatal = [];
  const warnings = [];
  if (sourceFiles.length !== 44) fatal.push(`Expected 44 v5 source workbooks, found ${sourceFiles.length}.`);
  if (outputs.length !== 44) fatal.push(`Expected 44 v6 workbooks, generated ${outputs.length}.`);

  const releaseIndex = [
    "file,source_file,release_name",
    ...outputs.map((row) => [row.outName, row.sourceFile, RELEASE_NAME].map(csvEscape).join(",")),
  ].join("\n");
  await fs.writeFile(path.join(releaseDir, "RELEASE_INDEX.csv"), `${releaseIndex}\n`, "utf8");

  const qa = {
    generatedAt: new Date().toISOString(),
    releaseName: RELEASE_NAME,
    sourceRelease: SOURCE_RELEASE,
    sourceWorkbooks: sourceFiles.length,
    v6Workbooks: outputs.length,
    fatal,
    warnings,
  };
  await fs.writeFile(path.join(releaseDir, "QA_REPORT.json"), `${JSON.stringify(qa, null, 2)}\n`, "utf8");
  await writeDocs({ sourceWorkbooks: sourceFiles.length, v6Workbooks: outputs.length, fatal, warnings });

  console.log(JSON.stringify(qa, null, 2));
  if (fatal.length) process.exitCode = 1;
}

await main();
