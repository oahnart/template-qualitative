import fs from "node:fs/promises";
import path from "node:path";
import { fileURLToPath } from "node:url";
import { FileBlob, SpreadsheetFile } from "@oai/artifact-tool";

const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);
const repoRoot = path.resolve(__dirname, "..");
const sourceDir = path.join(repoRoot, "consultant_safe_v2");
const releaseDir = path.join(repoRoot, "consultant_safe_v3");
const finalTemplateDir = path.join(repoRoot, "final_template", "template_qualitative");

const RELEASE_NAME = "consultant_safe_v3";
const SOURCE_RELEASE = "consultant_safe_v2";
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

function v3GuideRows(title, sourceFile) {
  return [
    [`Consultant-safe EBX-Q Template v3 - ${title}`, ""],
    ["Source", `Built from ${sourceFile}; source release ${SOURCE_RELEASE}; v3 source version ${SOURCE_VERSION}.`],
    ["Purpose", "Safe workbook for ESG qualitative input. V3 keeps v2 controls and adds consultant-ready answer-depth guidance."],
    ["Safety rule", "Safe does not mean minimal. Safe means source-backed, gap-aware, and explicit about what is known and not known."],
    ["Answer depth", "For SUFFICIENT evidence, final Korean answers should usually be 4-6 sentences and include scope, governance/process, activity, KPI/evidence, and any reviewer gap."],
    ["Partial evidence", "For PARTIAL evidence, disclose what the source supports, state the missing disclosure plainly, and tell the reviewer what to verify next."],
    ["Metric evidence", "For qualitative+quantitative rows, include at least one quantified evidence point when the source supports it; never estimate or turn missing data into 0."],
    ["Writing style", "Choose the actual style per row: evidence-led, governance-first, metric-first, risk-control-monitoring, gap-aware, or assertive consultant tone."],
    ["Final output", "Company-facing v3 output is a single clean workbook with English headers and Korean evidence/final answers."],
    ["Industry references", "Reference columns show topic scope only. Do not copy them as company facts or final prose."],
    ["Before reviewer", "Check trace metadata, answer depth, missing components, and every source-backed figure before handoff."],
  ];
}

async function writeDocs(summary) {
  const readme = `# Consultant-safe EBX-Q Template v3

## Purpose
This release keeps the consultant-safe controls from v2, but adds explicit answer-depth guidance for company-specific final responses.

## Key Changes
- Builds a new \`${RELEASE_NAME}\` release without overwriting \`${SOURCE_RELEASE}\`.
- Adds guidance that safe answers should be source-backed and sufficiently detailed, not merely short.
- Defines depth rules for SUFFICIENT, PARTIAL, and metric-backed evidence.
- Supports one clean company-facing output workbook with English headers and Korean evidence/final answers.

## Build Stats
- Source workbooks: ${summary.sourceWorkbooks}
- V3 workbooks: ${summary.v3Workbooks}
- Fatal issues: ${summary.fatal.length}
- Warnings: ${summary.warnings.length}
`;

  const checklist = `# consultant_safe_v3 Handoff Checklist

- [ ] Use one final company workbook as the user-facing deliverable.
- [ ] Keep column headers in English.
- [ ] Keep original evidence and final answers in Korean.
- [ ] Populate metadata with source PDF/page/evidence type/RAG trace when available; leave blank when unavailable.
- [ ] For SUFFICIENT rows, final answers are normally 4-6 Korean sentences.
- [ ] For PARTIAL rows, final answers explain the disclosed facts, missing facts, and reviewer follow-up.
- [ ] For qualitative+quantitative rows, include supported figures or clearly explain why figures are not used.
- [ ] Do not invent figures, targets, organizations, certifications, incidents, or processes.
- [ ] Do not convert missing values or dash markers into 0.
`;

  const changelog = `# Version Log

## ${RELEASE_NAME}
- Source release: \`${SOURCE_RELEASE}\`.
- Adds consultant-ready answer-depth guidance on top of v2 safety controls.
- Creates v3 workbook filenames ending in \`_consultant_safe_v3.xlsx\`.
- Supports the v3 single-file company output flow.
`;

  await fs.writeFile(path.join(releaseDir, "README.md"), readme, "utf8");
  await fs.writeFile(path.join(releaseDir, "CHECKLIST.md"), checklist, "utf8");
  await fs.writeFile(path.join(releaseDir, "CHANGELOG.md"), changelog, "utf8");
}

async function buildWorkbook(sourceFile) {
  const inputPath = path.join(sourceDir, sourceFile);
  const outName = sourceFile.replace("_consultant_safe_v2.xlsx", "_consultant_safe_v3.xlsx");
  const outPath = path.join(releaseDir, outName);
  const finalOutPath = path.join(finalTemplateDir, outName);

  const workbook = await SpreadsheetFile.importXlsx(await FileBlob.load(inputPath));
  const guide = workbook.worksheets.getItem("안내");
  const title = outName
    .replace(/^EBX_Q_템플릿_/, "")
    .replace("_consultant_safe_v3.xlsx", "");
  const rows = v3GuideRows(title, sourceFile);

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
    .filter((file) => file.endsWith("_consultant_safe_v2.xlsx"))
    .sort();

  const outputs = [];
  for (const file of sourceFiles) {
    outputs.push(await buildWorkbook(file));
  }

  const fatal = [];
  const warnings = [];
  if (sourceFiles.length !== 44) fatal.push(`Expected 44 v2 source workbooks, found ${sourceFiles.length}.`);
  if (outputs.length !== 44) fatal.push(`Expected 44 v3 workbooks, generated ${outputs.length}.`);

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
    v3Workbooks: outputs.length,
    finalTemplateDir,
    fatal,
    warnings,
  };
  await fs.writeFile(path.join(releaseDir, "QA_REPORT.json"), `${JSON.stringify(qa, null, 2)}\n`, "utf8");
  await writeDocs({ sourceWorkbooks: sourceFiles.length, v3Workbooks: outputs.length, fatal, warnings });

  console.log(JSON.stringify(qa, null, 2));
  if (fatal.length) process.exitCode = 1;
}

await main();
