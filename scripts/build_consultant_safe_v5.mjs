import fs from "node:fs/promises";
import path from "node:path";
import { fileURLToPath } from "node:url";
import { FileBlob, SpreadsheetFile } from "@oai/artifact-tool";

const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);
const repoRoot = path.resolve(__dirname, "..");
const sourceDir = path.join(repoRoot, "consultant_safe_v4");
const releaseDir = path.join(repoRoot, "consultant_safe_v5");

const RELEASE_NAME = "consultant_safe_v5";
const SOURCE_RELEASE = "consultant_safe_v4";

async function ensureDir(dir) {
  await fs.mkdir(dir, { recursive: true });
}

async function cleanupInspectSidecar(file) {
  await fs.rm(`${file}.inspect.ndjson`, { force: true });
}

function csvEscape(value) {
  return `"${String(value ?? "").replaceAll('"', '""')}"`;
}

function findHeaderIndex(headers, matcher) {
  return headers.findIndex((header) => matcher(String(header ?? "")));
}

function buildFieldPath(row, idx) {
  return [row[idx.area], row[idx.pillar], row[idx.item]]
    .map((value) => String(value ?? "").trim())
    .filter(Boolean)
    .join(" / ");
}

function antiRepetitionGroup(row, idx) {
  const ebx = String(row[idx.ebx] ?? "");
  const field = `${row[idx.pillar] ?? ""} ${row[idx.item] ?? ""}`;
  if (/001|008|020|024|전략|목표|정책/.test(`${ebx} ${field}`)) return "strategy-policy";
  if (/002|005|013|017|021|025|거버넌스|조직|책임|체계/.test(`${ebx} ${field}`)) return "governance-organization";
  if (/003|006|010|014|018|022|026|리스크|위험/.test(`${ebx} ${field}`)) return "risk-control";
  if (/007|011|015|019|023|027|성과|현황|지표|사고|위반/.test(`${ebx} ${field}`)) return "metrics-status";
  return "evidence-narrative";
}

function qaSeverity(row, idx) {
  const group = antiRepetitionGroup(row, idx);
  if (group === "metrics-status") return "High: metric and incident rows must use disclosed figures when present.";
  if (group === "risk-control") return "High: risk rows must describe identified risk, control, monitoring, and response.";
  if (group === "governance-organization") return "Medium: governance rows must identify responsible bodies and decision flow.";
  return "Medium: report-ready prose must remain evidence-grounded and non-repetitive.";
}

function v5TemplateEnhancements(row, idx) {
  const fieldPath = buildFieldPath(row, idx);
  const group = antiRepetitionGroup(row, idx);
  return [
    [
      "EVIDENCE SELECTION RULES:",
      "1. Select 2-4 company-specific evidence sentences before adding any connective prose.",
      "2. Prefer sentences with named policy, owner, process, target, period, scope, KPI, or incident status.",
      "3. Drop source titles, page labels, OCR navigation text, appendix labels, and generic report boilerplate.",
      "4. Do not fill answer depth with reusable generic sentences.",
      `Field Path: ${fieldPath}`,
    ].join("\n"),
    [
      "METRIC EVIDENCE REQUIREMENTS:",
      "1. If qualitative evidence or quantitative CSV contains disclosed numbers for this EBX row, include at least one supported figure.",
      "2. Keep period, unit, boundary, and trend when available.",
      "3. If metric evidence is expected but no disclosed number is available, state the business disclosure limitation without reviewer/source wording.",
      "4. Never convert blanks, dashes, or missing values to 0.",
    ].join("\n"),
    [
      "COMPANY NAMING RULE:",
      "1. Use the Korean report-facing company name in Final Answer when available.",
      "2. Do not use underscored IDs or English fallback names inside Korean prose.",
      "3. Keep source names and file names only in metadata, not Final Answer.",
    ].join("\n"),
    group,
    qaSeverity(row, idx),
  ];
}

function v5GuideRows(title, sourceFile) {
  return [
    [`Consultant-safe EBX-Q Template v5 - ${title}`, ""],
    ["Source", `Built from ${sourceFile}; source release ${SOURCE_RELEASE}.`],
    ["Purpose", "V5 keeps v4 report-ready controls and adds deterministic evidence selection, metric inclusion, company naming, repetition, and QA severity controls."],
    ["Evidence first", "Company output must select row-specific company evidence before adding connective prose."],
    ["Metric discipline", "Rows with qualitative+quantitative support must include disclosed figures when available, or state a business limitation without reviewer/source wording."],
    ["Naming", "Final Answer should use Korean report-facing company names such as 삼성전자 or LG전자 when available."],
    ["Anti-repetition", "The same Final Answer sentence may not appear in more than two EBX rows in a generated workbook."],
    ["Output", "Company-facing v5 output includes a clean answer sheet and QA_Audit sheet."],
    ["Before handoff", "Review QA_Audit for fatal findings, metric misses, forbidden language, company-name fallback, short answers, and repeated sentences."],
  ];
}

async function writeDocs(summary) {
  const readme = `# Consultant-safe EBX-Q Template v5

## Purpose
This release keeps the v4 report-ready workbook shape and adds deterministic evidence-grounding and QA controls.

## Key Changes
- Builds a new \`${RELEASE_NAME}\` release without overwriting \`${SOURCE_RELEASE}\`.
- Adds \`Evidence Selection Rules\`, \`Metric Evidence Requirements\`, \`Company Naming Rule\`, \`Anti-Repetition Group\`, and \`QA Severity\`.
- Requires generated company output to use row-specific evidence, supported figures when available, Korean report-facing company names, and non-repetitive final prose.
- Supports v5 output workbooks with a clean answer sheet plus \`QA_Audit\`.

## Build Stats
- Source workbooks: ${summary.sourceWorkbooks}
- V5 workbooks: ${summary.v5Workbooks}
- Fatal issues: ${summary.fatal.length}
- Warnings: ${summary.warnings.length}
`;

  const checklist = `# consultant_safe_v5 Handoff Checklist

- [ ] Use one final company workbook with \`consultant_safe_v5\` and \`QA_Audit\` sheets.
- [ ] Keep clean output headers in English.
- [ ] Keep original evidence and final answers in Korean where source evidence is Korean.
- [ ] Use Korean report-facing company names in Final Answer.
- [ ] Each \`Field\` value uses \`area / pillar / item\`.
- [ ] \`Style Template Applied\` shows selected style, selected pattern, content slots, anti-repetition treatment, coverage treatment, and metric treatment.
- [ ] \`Final Answer\` has no source names, page citations, PDF references, reviewer notes, or audit-trace wording.
- [ ] For qualitative+quantitative rows, include disclosed figures when available.
- [ ] No exact Final Answer sentence appears in more than two EBX rows.
- [ ] \`QA_Audit\` contains no fatal findings before handoff.
`;

  const changelog = `# Version Log

## ${RELEASE_NAME}
- Source release: \`${SOURCE_RELEASE}\`.
- Adds deterministic evidence selection, metric inclusion, company naming, anti-repetition, and QA audit controls.
- Creates v5 workbook filenames ending in \`_consultant_safe_v5.xlsx\`.
- Supports a two-sheet company output flow: clean answer sheet plus \`QA_Audit\`.
`;

  await fs.writeFile(path.join(releaseDir, "README.md"), readme, "utf8");
  await fs.writeFile(path.join(releaseDir, "CHECKLIST.md"), checklist, "utf8");
  await fs.writeFile(path.join(releaseDir, "CHANGELOG.md"), changelog, "utf8");
}

async function buildWorkbook(sourceFile) {
  const inputPath = path.join(sourceDir, sourceFile);
  const outName = sourceFile.replace("_consultant_safe_v4.xlsx", "_consultant_safe_v5.xlsx");
  const outPath = path.join(releaseDir, outName);

  const workbook = await SpreadsheetFile.importXlsx(await FileBlob.load(inputPath));
  const guide = workbook.worksheets.getItem("안내");
  const title = outName
    .replace(/^EBX_Q_템플릿_/, "")
    .replace("_consultant_safe_v5.xlsx", "");
  const guideRows = v5GuideRows(title, sourceFile);

  guide.getRangeByIndexes(0, 0, guideRows.length, 2).values = guideRows;
  guide.getRange(`A1:B${guideRows.length}`).format.wrapText = true;
  guide.getRange(`A1:B${guideRows.length}`).format.verticalAlignment = "top";
  guide.getRange("A1:B1").merge();
  guide.getRange("A1").format = {
    fill: "#0F5B4F",
    font: { bold: true, color: "#FFFFFF", size: 15 },
  };
  guide.getRange(`A2:A${guideRows.length}`).format = {
    fill: "#DDEFEA",
    font: { bold: true, color: "#0F3F37" },
  };
  guide.getRange(`A1:A${guideRows.length}`).format.columnWidthPx = 190;
  guide.getRange(`B1:B${guideRows.length}`).format.columnWidthPx = 920;

  const templateSheet = workbook.worksheets.getItem("EBX-Q 템플릿");
  const values = templateSheet.getRange("A1:S28").values;
  const headers = values[0];
  const idx = {
    ebx: findHeaderIndex(headers, (header) => header === "ebx"),
    area: findHeaderIndex(headers, (header) => header === "area"),
    pillar: findHeaderIndex(headers, (header) => header === "pillar"),
    item: findHeaderIndex(headers, (header) => header === "item"),
  };

  const v5Headers = [
    "Evidence Selection Rules",
    "Metric Evidence Requirements",
    "Company Naming Rule",
    "Anti-Repetition Group",
    "QA Severity",
  ];
  const v5Rows = values.slice(1).map((row) => v5TemplateEnhancements(row, idx));
  const startCol = headers.length;
  templateSheet.getRangeByIndexes(0, startCol, 1, v5Headers.length).values = [v5Headers];
  templateSheet.getRangeByIndexes(1, startCol, v5Rows.length, v5Headers.length).values = v5Rows;
  templateSheet.getRangeByIndexes(0, startCol, v5Rows.length + 1, v5Headers.length).format.wrapText = true;
  templateSheet.getRangeByIndexes(0, startCol, v5Rows.length + 1, v5Headers.length).format.verticalAlignment = "top";
  templateSheet.getRangeByIndexes(0, startCol, 1, v5Headers.length).format = {
    fill: "#0F5B4F",
    font: { bold: true, color: "#FFFFFF" },
  };
  templateSheet.getRange("T:W").format.columnWidthPx = 360;
  templateSheet.getRange("X:X").format.columnWidthPx = 300;

  const exported = await SpreadsheetFile.exportXlsx(workbook);
  await exported.save(outPath);
  await cleanupInspectSidecar(outPath);
  return { sourceFile, outName };
}

async function main() {
  await ensureDir(releaseDir);

  const sourceFiles = (await fs.readdir(sourceDir))
    .filter((file) => file.endsWith("_consultant_safe_v4.xlsx"))
    .sort();

  const outputs = [];
  for (const file of sourceFiles) {
    outputs.push(await buildWorkbook(file));
  }

  const fatal = [];
  const warnings = [];
  if (sourceFiles.length !== 44) fatal.push(`Expected 44 v4 source workbooks, found ${sourceFiles.length}.`);
  if (outputs.length !== 44) fatal.push(`Expected 44 v5 workbooks, generated ${outputs.length}.`);

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
    v5Workbooks: outputs.length,
    fatal,
    warnings,
  };
  await fs.writeFile(path.join(releaseDir, "QA_REPORT.json"), `${JSON.stringify(qa, null, 2)}\n`, "utf8");
  await writeDocs({ sourceWorkbooks: sourceFiles.length, v5Workbooks: outputs.length, fatal, warnings });

  console.log(JSON.stringify(qa, null, 2));
  if (fatal.length) process.exitCode = 1;
}

await main();
