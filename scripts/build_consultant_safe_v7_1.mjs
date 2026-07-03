import fs from "node:fs/promises";
import path from "node:path";
import { fileURLToPath } from "node:url";
import { FileBlob, SpreadsheetFile } from "@oai/artifact-tool";

const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);
const repoRoot = path.resolve(__dirname, "..");
const sourceDir = path.join(repoRoot, "consultant_safe_v7");
const releaseDir = path.join(repoRoot, "consultant_safe_v7_1");
const templateJsonDir = path.join(releaseDir, "templates");

const RELEASE_NAME = "consultant_safe_v7_1";
const SOURCE_RELEASE = "consultant_safe_v7";
const TEMPLATE_SHEET = "EBX-Q 템플릿";

const FIELD_ALIASES = {
  "Answer Type": "answerType",
  "Answer Intent": "answerIntent",
  "Opening Strategy": "openingStrategy",
  "Evidence Weave": "evidenceWeave",
  "Required Facts": "requiredFacts",
  "Plain-Language Avoid List": "avoidList",
  "Style Guardrails": "styleGuardrails",
  "QA Severity": "qaSeverity",
};

async function ensureDirs() {
  await fs.mkdir(templateJsonDir, { recursive: true });
}

function parseTemplateName(file) {
  const parts = file.replace(/\.xlsx$/i, "").split("_");
  return {
    sector: parts[3] ?? "",
    sectorName: parts[4] ?? "",
    size: parts[5] ?? "",
  };
}

function jsonTemplateName({ sector, size }) {
  return `${sector}_${size}_${RELEASE_NAME}.json`;
}

function normalizeTemplateRow(row, headers) {
  const out = {};
  for (const [index, header] of headers.entries()) {
    const key = String(header ?? "").trim();
    if (!key) continue;
    const value = row[index] ?? "";
    out[key] = value;
    if (FIELD_ALIASES[key]) out[FIELD_ALIASES[key]] = value;
  }
  return out;
}

async function buildJsonTemplate(sourceFile) {
  const sourcePath = path.join(sourceDir, sourceFile);
  const workbook = await SpreadsheetFile.importXlsx(await FileBlob.load(sourcePath));
  const sheet = workbook.worksheets.getItem(TEMPLATE_SHEET);
  const values = sheet.getRange("A1:V28").values;
  const headers = values[0].map((header) => String(header ?? ""));
  const meta = parseTemplateName(sourceFile);
  const outName = jsonTemplateName(meta);
  const output = {
    releaseName: RELEASE_NAME,
    sourceRelease: SOURCE_RELEASE,
    sourceWorkbook: sourceFile,
    sector: meta.sector,
    sectorName: meta.sectorName,
    size: meta.size,
    headers,
    rows: values.slice(1).map((row) => normalizeTemplateRow(row, headers)),
  };
  await fs.writeFile(path.join(templateJsonDir, outName), `${JSON.stringify(output, null, 2)}\n`, "utf8");
  return { ...meta, sourceWorkbook: sourceFile, jsonTemplate: path.join("templates", outName), rows: output.rows.length };
}

async function writeDocs(summary) {
  const readme = `# Consultant-safe EBX-Q Template v7.1

## Purpose
V7.1 changes the template source-of-truth from Excel workbooks to JSON files.

## What Changed
- Keeps v7 answer controls and customer output schema.
- Converts each sector/size template workbook into a full-column JSON template under \`templates/\`.
- Preserves the Excel header list in \`headers\` and row values using the same header names.
- Keeps internal camelCase aliases for generator fields so existing scripts can read the JSON safely.
- Company filling reads JSON templates directly and only creates Excel for final customer delivery.
- Future template improvements should edit JSON, not regenerate 44 Excel workbooks.

## Build Stats
- Source workbooks read once: ${summary.sourceWorkbooks}
- JSON templates generated: ${summary.jsonTemplates}
- Fatal issues: ${summary.fatal.length}
- Warnings: ${summary.warnings.length}
`;

  const checklist = `# consultant_safe_v7_1 Handoff Checklist

- [ ] Use JSON files in \`consultant_safe_v7_1/templates\` as the template source-of-truth.
- [ ] Do not regenerate 44 template Excel workbooks for routine prompt/rule improvements.
- [ ] Customer workbook preserves the six established output columns.
- [ ] Sidecar QA JSON has zero fatal findings before handoff.
`;

  const changelog = `# Version Log

## consultant_safe_v7_1
- Source release: \`${SOURCE_RELEASE}\`.
- Converts v7 Excel templates into JSON templates.
- Keeps generator-safe aliases in each row.
- Updates company fill flow to read JSON templates and export only the final company workbook.
`;

  await fs.writeFile(path.join(releaseDir, "README.md"), readme, "utf8");
  await fs.writeFile(path.join(releaseDir, "CHECKLIST.md"), checklist, "utf8");
  await fs.writeFile(path.join(releaseDir, "CHANGELOG.md"), changelog, "utf8");
}

async function main() {
  await ensureDirs();
  const sourceFiles = (await fs.readdir(sourceDir))
    .filter((file) => file.endsWith("_consultant_safe_v7.xlsx"))
    .sort();

  const templates = [];
  for (const file of sourceFiles) {
    templates.push(await buildJsonTemplate(file));
  }

  const fatal = [];
  const warnings = [];
  if (sourceFiles.length !== 44) fatal.push(`Expected 44 v7 source workbooks, found ${sourceFiles.length}.`);
  if (templates.length !== 44) fatal.push(`Expected 44 v7.1 JSON templates, generated ${templates.length}.`);
  for (const template of templates) {
    if (template.rows !== 27) fatal.push(`${template.jsonTemplate}: expected 27 EBX rows, found ${template.rows}.`);
  }

  const index = {
    generatedAt: new Date().toISOString(),
    releaseName: RELEASE_NAME,
    sourceRelease: SOURCE_RELEASE,
    templates,
  };
  await fs.writeFile(path.join(releaseDir, "TEMPLATE_INDEX.json"), `${JSON.stringify(index, null, 2)}\n`, "utf8");

  const qa = {
    generatedAt: new Date().toISOString(),
    releaseName: RELEASE_NAME,
    sourceRelease: SOURCE_RELEASE,
    sourceWorkbooks: sourceFiles.length,
    jsonTemplates: templates.length,
    fatal,
    warnings,
  };
  await fs.writeFile(path.join(releaseDir, "QA_REPORT.json"), `${JSON.stringify(qa, null, 2)}\n`, "utf8");
  await writeDocs({ sourceWorkbooks: sourceFiles.length, jsonTemplates: templates.length, fatal, warnings });

  console.log(JSON.stringify(qa, null, 2));
  if (fatal.length) process.exitCode = 1;
}

await main();
