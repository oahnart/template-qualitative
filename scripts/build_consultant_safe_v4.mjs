import fs from "node:fs/promises";
import path from "node:path";
import { fileURLToPath } from "node:url";
import { FileBlob, SpreadsheetFile } from "@oai/artifact-tool";

const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);
const repoRoot = path.resolve(__dirname, "..");
const sourceDir = path.join(repoRoot, "consultant_safe_v3");
const releaseDir = path.join(repoRoot, "consultant_safe_v4");
const finalTemplateDir = path.join(repoRoot, "final_template", "template_qualitative");

const RELEASE_NAME = "consultant_safe_v4";
const SOURCE_RELEASE = "consultant_safe_v3";
const SOURCE_VERSION = "2026-07-02";

async function ensureDir(dir) {
  await fs.mkdir(dir, { recursive: true });
}

async function cleanupInspectSidecar(file) {
  await fs.rm(`${file}.inspect.ndjson`, { force: true });
}

function csvEscape(value) {
  return `"${String(value ?? "").replaceAll('"', '""')}"`;
}

function v4GuideRows(title, sourceFile) {
  return [
    [`Consultant-safe EBX-Q Template v4 - ${title}`, ""],
    ["Source", `Built from ${sourceFile}; source release ${SOURCE_RELEASE}; v4 source version ${SOURCE_VERSION}.`],
    ["Purpose", "Safe workbook for ESG qualitative input. V4 keeps v3 answer-depth controls and adds report-ready final answer guardrails."],
    ["Field path", "Downstream output must use Field = area / pillar / item, for example: general / strategy / ESG vision and mid-to-long-term strategy."],
    ["Style template applied", "Company output must describe the applied Content Slots, Style Options, Sentence Patterns, anti-repetition guidance, coverage treatment, and metric handling."],
    ["Final answer rule", "Final Answer must be ESG report-ready prose only. Do not include source names, page citations, audit trace language, reviewer notes, or source/reviewer wording."],
    ["Answer depth", "For SUFFICIENT evidence, final Korean answers should usually be 4-6 sentences and include scope, governance/process, activity, KPI/evidence, and any business-relevant limitation."],
    ["Partial evidence", "For PARTIAL evidence, write only supported company facts and state missing disclosure as a business limitation without source/reviewer phrasing."],
    ["Metric evidence", "For qualitative+quantitative rows, include supported figures with period, unit, and scope when useful; never estimate or turn missing data into 0."],
    ["Writing style", "Choose the actual style per row: evidence-led, governance-first, metric-first, risk-control-monitoring, gap-aware, or assertive consultant tone."],
    ["Final output", "Company-facing v4 output is a single clean workbook with English headers and Korean evidence/final answers."],
    ["Industry references", "Reference columns show topic scope only. Do not copy them as company facts or final prose."],
    ["Before handoff", "Check field paths, style template detail, final answer cleanliness, answer depth, missing components, and every supported figure before handoff."],
  ];
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

function v4TemplateEnhancements(row, idx) {
  const contentSlots = row[idx.contentSlots] ?? "";
  const styleOptions = row[idx.styleOptions] ?? "";
  const sentencePatterns = row[idx.sentencePatterns] ?? "";
  const antiRepetition = row[idx.antiRepetition] ?? "";
  return [
    buildFieldPath(row, idx),
    [
      "Use ESG report-ready Korean prose.",
      "Cover the relevant content slots before adding optional context.",
      "Prefer 4-6 sentences when evidence is sufficient; keep partial rows factual and bounded.",
      "Do not include source/page/reviewer/audit-trace wording in Final Answer.",
      `Content Slots: ${contentSlots}`,
    ].join("\n"),
    [
      "Final Answer may use company facts and supported figures only.",
      "Remove source names, source pages, PDF references, citation language, and reviewer follow-up wording.",
      "Do not invent figures, targets, organizations, certifications, incidents, or processes.",
      `Style Options: ${styleOptions}`,
      `Sentence Patterns: ${sentencePatterns}`,
      `Anti-Repetition: ${antiRepetition}`,
    ].join("\n"),
    [
      "SUFFICIENT: write a complete report paragraph with scope, governance/process, activity, and KPI or limitation where relevant.",
      "PARTIAL: disclose supported facts only and phrase missing data as a disclosure limitation, not as an audit note.",
      "NO DATA: state that the company has not disclosed or has not yet managed the item, without source tracing.",
    ].join("\n"),
    [
      "When quantitative support exists, include period, unit, boundary, and trend if available.",
      "When quantitative support is unavailable, do not convert blanks, dashes, or missing values to zero.",
      "Keep metrics inside the report prose, not as source evidence notes.",
    ].join("\n"),
  ];
}

async function writeDocs(summary) {
  const readme = `# Consultant-safe EBX-Q Template v4

## Purpose
This release keeps the consultant-safe controls from v3, then adds field-path, style-template, and report-ready final answer controls.

## Key Changes
- Builds a new \`${RELEASE_NAME}\` release without overwriting \`${SOURCE_RELEASE}\`.
- Adds \`Field Path\`, \`Final Answer Requirements\`, \`Report-Ready Guardrails\`, \`Coverage Handling\`, and \`Metric Handling\` to every template workbook.
- Requires company output \`Field\` to use \`area / pillar / item\`.
- Requires \`Style Template Applied\` to describe content slots, style options, sentence patterns, anti-repetition guidance, coverage treatment, and metric handling.
- Requires \`Final Answer\` to remove source names, page citations, PDF references, reviewer notes, and audit-trace wording.

## Build Stats
- Source workbooks: ${summary.sourceWorkbooks}
- V4 workbooks: ${summary.v4Workbooks}
- Fatal issues: ${summary.fatal.length}
- Warnings: ${summary.warnings.length}
`;

  const checklist = `# consultant_safe_v4 Handoff Checklist

- [ ] Use one final company workbook as the user-facing deliverable.
- [ ] Keep column headers in English.
- [ ] Keep original evidence and final answers in Korean.
- [ ] Populate metadata with source PDF/page/evidence type/RAG trace when available; leave blank when unavailable.
- [ ] Output header \`Area\` has been replaced by \`Field\`.
- [ ] Output header \`Writing Style Template\` has been replaced by \`Style Template Applied\`.
- [ ] Each \`Field\` value uses \`area / pillar / item\`.
- [ ] \`Style Template Applied\` includes Content Slots, Style Options, Sentence Patterns, anti-repetition guidance, coverage treatment, and metric handling.
- [ ] \`Final Answer\` has no source names, page citations, PDF references, reviewer notes, or audit-trace wording.
- [ ] For SUFFICIENT rows, final answers are normally 4-6 Korean sentences.
- [ ] For PARTIAL rows, final answers explain disclosed facts and disclosure limitations without reviewer/source wording.
- [ ] For qualitative+quantitative rows, include supported figures or clearly explain why figures are not used.
- [ ] Do not invent figures, targets, organizations, certifications, incidents, or processes.
- [ ] Do not convert missing values or dash markers into 0.
`;

  const changelog = `# Version Log

## ${RELEASE_NAME}
- Source release: \`${SOURCE_RELEASE}\`.
- Adds report-ready output controls on top of v3 safety controls.
- Creates v4 workbook filenames ending in \`_consultant_safe_v4.xlsx\`.
- Supports the v4 single-file company output flow.
`;

  await fs.writeFile(path.join(releaseDir, "README.md"), readme, "utf8");
  await fs.writeFile(path.join(releaseDir, "CHECKLIST.md"), checklist, "utf8");
  await fs.writeFile(path.join(releaseDir, "CHANGELOG.md"), changelog, "utf8");
}

async function buildWorkbook(sourceFile) {
  const inputPath = path.join(sourceDir, sourceFile);
  const outName = sourceFile.replace("_consultant_safe_v3.xlsx", "_consultant_safe_v4.xlsx");
  const outPath = path.join(releaseDir, outName);
  const finalOutPath = path.join(finalTemplateDir, outName);

  const workbook = await SpreadsheetFile.importXlsx(await FileBlob.load(inputPath));
  const guide = workbook.worksheets.getItem("안내");
  const title = outName
    .replace(/^EBX_Q_템플릿_/, "")
    .replace("_consultant_safe_v4.xlsx", "");
  const rows = v4GuideRows(title, sourceFile);

  guide.getRangeByIndexes(0, 0, rows.length, 2).values = rows;
  guide.getRange(`A1:B${rows.length}`).format.wrapText = true;
  guide.getRange(`A1:B${rows.length}`).format.verticalAlignment = "top";
  guide.getRange("A1:B1").merge();
  guide.getRange("A1").format = {
    fill: "#174A7C",
    font: { bold: true, color: "#FFFFFF", size: 15 },
  };
  guide.getRange(`A2:A${rows.length}`).format = {
    fill: "#D9EAF7",
    font: { bold: true, color: "#17365D" },
  };
  guide.getRange(`A1:A${rows.length}`).format.columnWidthPx = 180;
  guide.getRange(`B1:B${rows.length}`).format.columnWidthPx = 900;

  const templateSheet = workbook.worksheets.getItem("EBX-Q 템플릿");
  const templateValues = templateSheet.getRange("A1:N28").values;
  const headers = templateValues[0];
  const idx = {
    area: findHeaderIndex(headers, (header) => header === "area"),
    pillar: findHeaderIndex(headers, (header) => header === "pillar"),
    item: findHeaderIndex(headers, (header) => header === "item"),
    contentSlots: findHeaderIndex(headers, (header) => header.includes("Content Slots")),
    styleOptions: findHeaderIndex(headers, (header) => header.includes("Style Options")),
    sentencePatterns: findHeaderIndex(headers, (header) => header.includes("Sentence Patterns")),
    antiRepetition: findHeaderIndex(headers, (header) => header.includes("Anti-Repetition")),
  };
  const v4Headers = [
    "Field Path",
    "Final Answer Requirements",
    "Report-Ready Guardrails",
    "Coverage Handling",
    "Metric Handling",
  ];
  const v4Rows = templateValues.slice(1).map((row) => v4TemplateEnhancements(row, idx));
  templateSheet.getRangeByIndexes(0, headers.length, 1, v4Headers.length).values = [v4Headers];
  templateSheet.getRangeByIndexes(1, headers.length, v4Rows.length, v4Headers.length).values = v4Rows;
  templateSheet.getRangeByIndexes(0, headers.length, v4Rows.length + 1, v4Headers.length).format.wrapText = true;
  templateSheet.getRangeByIndexes(0, headers.length, v4Rows.length + 1, v4Headers.length).format.verticalAlignment = "top";
  templateSheet.getRangeByIndexes(0, headers.length, 1, v4Headers.length).format = {
    fill: "#174A7C",
    font: { bold: true, color: "#FFFFFF" },
  };
  templateSheet.getRange("O:S").format.columnWidthPx = 360;

  const exported = await SpreadsheetFile.exportXlsx(workbook);
  await exported.save(outPath);
  await exported.save(finalOutPath);
  await cleanupInspectSidecar(outPath);
  await cleanupInspectSidecar(finalOutPath);
  return { sourceFile, outName, outPath, finalOutPath };
}

async function main() {
  await ensureDir(releaseDir);
  await ensureDir(finalTemplateDir);

  const sourceFiles = (await fs.readdir(sourceDir))
    .filter((file) => file.endsWith("_consultant_safe_v3.xlsx"))
    .sort();

  const outputs = [];
  for (const file of sourceFiles) {
    outputs.push(await buildWorkbook(file));
  }

  const fatal = [];
  const warnings = [];
  if (sourceFiles.length !== 44) fatal.push(`Expected 44 v3 source workbooks, found ${sourceFiles.length}.`);
  if (outputs.length !== 44) fatal.push(`Expected 44 v4 workbooks, generated ${outputs.length}.`);

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
    v4Workbooks: outputs.length,
    finalTemplateDir,
    fatal,
    warnings,
  };
  await fs.writeFile(path.join(releaseDir, "QA_REPORT.json"), `${JSON.stringify(qa, null, 2)}\n`, "utf8");
  await writeDocs({ sourceWorkbooks: sourceFiles.length, v4Workbooks: outputs.length, fatal, warnings });

  console.log(JSON.stringify(qa, null, 2));
  if (fatal.length) process.exitCode = 1;
}

await main();
