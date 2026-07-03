import fs from "node:fs/promises";
import path from "node:path";
import { fileURLToPath } from "node:url";
import { FileBlob, SpreadsheetFile } from "@oai/artifact-tool";

const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);
const repoRoot = path.resolve(__dirname, "..");
const sourceDir = path.join(repoRoot, "consultant_safe_v6");
const releaseDir = path.join(repoRoot, "consultant_safe_v7");

const RELEASE_NAME = "consultant_safe_v7";
const SOURCE_RELEASE = "consultant_safe_v6";
const TEMPLATE_SHEET = "EBX-Q 템플릿";
const GUIDE_SHEET = "안내";
const BASE_COLUMN_COUNT = 14;

const V7_HEADERS = [
  "Answer Type",
  "Answer Intent",
  "Opening Strategy",
  "Evidence Weave",
  "Required Facts",
  "Plain-Language Avoid List",
  "Style Guardrails",
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
  if ([7, 11, 15, 19, 23, 27].includes(num)) return "status-performance";
  return "policy-management";
}

function fieldPath(row, idx) {
  return [row[idx.area], row[idx.pillar], row[idx.item]]
    .map((value) => String(value ?? "").trim())
    .filter(Boolean)
    .join(" / ");
}

function answerIntent(type) {
  const intents = {
    "strategy-policy": "Explain the policy or strategic direction first, then connect targets, governance, and execution evidence.",
    "governance-organization": "Explain who is responsible, what they decide or oversee, and how reporting or escalation works.",
    "risk-control": "Explain the risk/control scope, then prevention, monitoring, and follow-up actions.",
    "status-performance": "Explain the reported status in plain language, then weave figures into the management meaning.",
    "policy-management": "Explain the standard or policy, then show implementation ownership and supporting performance facts.",
  };
  return intents[type] ?? intents["policy-management"];
}

function openingStrategy(type) {
  const openings = {
    "strategy-policy": "Open with business direction, policy commitment, or strategic objective. Do not open with a metric label.",
    "governance-organization": "Open with the responsible committee, executive, department, or governance body.",
    "risk-control": "Open with the material risk, control scope, or prevention process.",
    "status-performance": "Open with the business context or reported status. Put numbers after the context sentence.",
    "policy-management": "Open with the policy, standard, or operating principle.",
  };
  return openings[type] ?? openings["policy-management"];
}

function evidenceWeave(row, idx, type) {
  return [
    `Field Path: ${fieldPath(row, idx)}`,
    "Use 2-4 row-specific evidence sentences, repaired for obvious OCR spacing.",
    "If a figure is needed, introduce it after the business context; avoid formulaic metric-first openings.",
    "Reject page labels, report navigation, table/index rows, raw headings, and long list fragments.",
    `Intent: ${answerIntent(type)}`,
  ].join("\n");
}

function requiredFacts(row, idx, type) {
  const num = ebxNumber(row, idx);
  if (num === 27) {
    return "Reported ethics/compliance incident status; report intake or violation cases; investigation/action or disciplinary outcomes; compliance/ethics training completions when disclosed; period, unit, trend, and management interpretation.";
  }
  const facts = {
    "strategy-policy": "Policy/strategy; target or roadmap when disclosed; implementation owner or activity; limitation if disclosure is partial.",
    "governance-organization": "Responsible body/owner; role; reporting or decision flow; oversight level.",
    "risk-control": "Risk/control scope; preventive or mitigation action; monitoring cadence or process; follow-up or escalation.",
    "status-performance": "Reported status; period and unit for figures; boundary when disclosed; trend or management interpretation.",
    "policy-management": "Policy/standard; operating process; responsible owner; supported figure or target when relevant.",
  };
  return facts[type] ?? facts["policy-management"];
}

function avoidList() {
  return [
    "Do not use: quantitative, 정량, 정량 지표, 정량 목표, định lượng.",
    "Do not use source/page/PDF/audit/reviewer wording in Final Answer.",
    "Do not expose EBX codes, raw indicator names, or template terms.",
    "Do not keep OCR/table headings such as Corporate Governance, 구분, 추진체계와 주요성과, 코드 공시 항목.",
    "For ethics/compliance incident rows, use plain words such as report intake, actions, training completions, or committee/self-check activity when supported.",
  ].join("\n");
}

function styleGuardrails(row, idx, type) {
  const num = ebxNumber(row, idx);
  const base = [
    "Final Answer should read like consultant-ready prose, not a copied table row.",
    "Avoid repeating the same opening pattern across rows.",
    "Use plain wording that non-specialist clients can understand.",
    "Keep the customer output schema unchanged.",
  ];
  if (type === "status-performance") {
    base.push("Figures must support the same EBX topic; do not match metrics by page alone.");
  }
  if (num === 27) {
    base.push("For ethics/compliance incident status, do not omit disclosed numbers for report intake, actions, training completions, committee meetings, or self-check participation.");
  }
  return base.join("\n");
}

function qaSeverity(type) {
  if (type === "status-performance") return "High: topic-fit figures, non-repetitive opening, no technical metric label.";
  if (type === "risk-control") return "High: must include risk/control and monitoring evidence.";
  if (type === "governance-organization") return "Medium: responsibility and decision flow must be clear.";
  return "Medium: prose must be natural, evidence-grounded, and non-repetitive.";
}

function v7Enhancements(row, idx) {
  const type = answerTypeFor(row, idx);
  return [
    type,
    answerIntent(type),
    openingStrategy(type),
    evidenceWeave(row, idx, type),
    requiredFacts(row, idx, type),
    avoidList(),
    styleGuardrails(row, idx, type),
    qaSeverity(type),
  ];
}

function guideRows(title, sourceFile) {
  return [
    [`Consultant-safe EBX-Q Template v7 - ${title}`, ""],
    ["Source", `Built from ${sourceFile}; source release ${SOURCE_RELEASE}.`],
    ["Purpose", "V7 improves Final Answer prose by replacing formulaic metric/template controls with plain-language evidence and opening controls."],
    ["Customer output", "Generated output keeps the same six columns: EBX Indicator, Field, Original Answer, Original Answer Metadata, Style Template Applied, Final Answer."],
    ["Metric wording", "Final Answer must not use technical labels such as quantitative, 정량, or định lượng. Figures are woven into business prose after context."],
    ["QA", "Sidecar JSON fails repeated openings, technical metric labels, source traces, OCR/table artifacts, and topic-mismatched figures."],
  ];
}

async function writeDocs(summary) {
  const readme = `# Consultant-safe EBX-Q Template v7

## Purpose
V7 keeps the v6 workbook family but changes the style-control columns so Final Answer prose is less repetitive and easier for non-specialist clients to read.

## Key Changes
- Builds a new \`${RELEASE_NAME}\` release without overwriting \`${SOURCE_RELEASE}\`.
- Replaces metric/formula-oriented controls with \`Answer Intent\`, \`Opening Strategy\`, \`Evidence Weave\`, \`Required Facts\`, \`Plain-Language Avoid List\`, \`Style Guardrails\`, and \`QA Severity\`.
- Preserves the customer output schema: \`EBX Indicator\`, \`Field\`, \`Original Answer\`, \`Original Answer Metadata\`, \`Style Template Applied\`, and \`Final Answer\`.
- Instructs generators to avoid \`quantitative\`, \`정량\`, and \`định lượng\` in customer-facing answers.

## Build Stats
- Source workbooks: ${summary.sourceWorkbooks}
- V7 workbooks: ${summary.v7Workbooks}
- Fatal issues: ${summary.fatal.length}
- Warnings: ${summary.warnings.length}
`;

  const checklist = `# consultant_safe_v7 Handoff Checklist

- [ ] Customer workbook preserves the six established output columns.
- [ ] \`Final Answer\` does not contain \`quantitative\`, \`정량\`, \`định lượng\`, source/page/PDF/reviewer wording, EBX codes, raw headings, or OCR/table fragments.
- [ ] Figures are introduced after business context, not through repeated metric-label openings.
- [ ] Same opening pattern does not appear more than two times.
- [ ] Metric/status rows use figures that fit the EBX topic, not merely the same PDF page.
- [ ] Sidecar QA JSON has zero fatal findings before handoff.
`;

  const changelog = `# Version Log

## ${RELEASE_NAME}
- Source release: \`${SOURCE_RELEASE}\`.
- Replaces v6 formula-style controls with plain-language answer controls.
- Creates v7 workbook filenames ending in \`_consultant_safe_v7.xlsx\`.
- Preserves the customer output schema and uses sidecar JSON QA.
`;

  await fs.writeFile(path.join(releaseDir, "README.md"), readme, "utf8");
  await fs.writeFile(path.join(releaseDir, "CHECKLIST.md"), checklist, "utf8");
  await fs.writeFile(path.join(releaseDir, "CHANGELOG.md"), changelog, "utf8");
}

async function buildWorkbook(sourceFile) {
  const inputPath = path.join(sourceDir, sourceFile);
  const outName = sourceFile.replace("_consultant_safe_v6.xlsx", "_consultant_safe_v7.xlsx");
  const outPath = path.join(releaseDir, outName);
  const workbook = await SpreadsheetFile.importXlsx(await FileBlob.load(inputPath));

  const guide = workbook.worksheets.getItem(GUIDE_SHEET);
  const title = outName
    .replace(/^EBX_Q_템플릿_/, "")
    .replace("_consultant_safe_v7.xlsx", "");
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
  const sourceValues = sheet.getRange("A1:V28").values;
  const sourceHeaders = sourceValues[0];
  const idx = {
    ebx: indexByHeader(sourceHeaders, "ebx"),
    area: indexByHeader(sourceHeaders, "area"),
    pillar: indexByHeader(sourceHeaders, "pillar"),
    item: indexByHeader(sourceHeaders, "item"),
  };

  const baseHeaders = sourceHeaders.slice(0, BASE_COLUMN_COUNT);
  const baseRows = sourceValues.slice(1).map((row) => row.slice(0, BASE_COLUMN_COUNT));
  const outputHeaders = [...baseHeaders, ...V7_HEADERS];
  const outputRows = baseRows.map((row, index) => [
    ...row,
    ...v7Enhancements(sourceValues[index + 1], idx),
  ]);

  sheet.getRangeByIndexes(0, 0, 1, outputHeaders.length).values = [outputHeaders];
  sheet.getRangeByIndexes(1, 0, outputRows.length, outputHeaders.length).values = outputRows;
  sheet.getRangeByIndexes(0, 0, outputRows.length + 1, outputHeaders.length).format.wrapText = true;
  sheet.getRangeByIndexes(0, 0, outputRows.length + 1, outputHeaders.length).format.verticalAlignment = "top";
  sheet.getRangeByIndexes(0, BASE_COLUMN_COUNT, 1, V7_HEADERS.length).format = {
    fill: "#174A5A",
    font: { bold: true, color: "#FFFFFF" },
  };
  sheet.getRange("O:V").format.columnWidthPx = 340;

  const exported = await SpreadsheetFile.exportXlsx(workbook);
  await exported.save(outPath);
  await cleanupInspectSidecar(outPath);
  return { sourceFile, outName };
}

async function main() {
  await ensureDir(releaseDir);
  const sourceFiles = (await fs.readdir(sourceDir))
    .filter((file) => file.endsWith("_consultant_safe_v6.xlsx"))
    .sort();

  const outputs = [];
  for (const file of sourceFiles) {
    outputs.push(await buildWorkbook(file));
  }

  const fatal = [];
  const warnings = [];
  if (sourceFiles.length !== 44) fatal.push(`Expected 44 v6 source workbooks, found ${sourceFiles.length}.`);
  if (outputs.length !== 44) fatal.push(`Expected 44 v7 workbooks, generated ${outputs.length}.`);

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
    v7Workbooks: outputs.length,
    fatal,
    warnings,
  };
  await fs.writeFile(path.join(releaseDir, "QA_REPORT.json"), `${JSON.stringify(qa, null, 2)}\n`, "utf8");
  await writeDocs({ sourceWorkbooks: sourceFiles.length, v7Workbooks: outputs.length, fatal, warnings });

  console.log(JSON.stringify(qa, null, 2));
  if (fatal.length) process.exitCode = 1;
}

await main();
