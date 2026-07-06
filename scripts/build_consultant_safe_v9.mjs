import fs from "node:fs/promises";
import path from "node:path";
import { fileURLToPath } from "node:url";

const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);
const repoRoot = path.resolve(__dirname, "..");
const sourceDir = path.join(repoRoot, "consultant_safe_v8");
const sourceTemplateDir = path.join(sourceDir, "templates");
const releaseDir = path.join(repoRoot, "consultant_safe_v9");
const templateJsonDir = path.join(releaseDir, "templates");

const RELEASE_NAME = "consultant_safe_v9";
const SOURCE_RELEASE = "consultant_safe_v8";

const BASE_QA_RULES = {
  requireKorean: true,
  minSentences: 8,
  minCharsWarn: 900,
  forbidTechnicalMetricWording: true,
  forbidEbxCode: true,
  forbidSourceTrace: true,
  forbidOcrArtifacts: true,
  forbidGenericV6FallbackGrammar: true,
  requireMetricNumberWhenMetricExpected: true,
};

const BASE_LENGTH_POLICY = {
  sufficient: {
    minSentences: 8,
    minCharsWarn: 900,
    targetMinChars: 900,
    targetMaxChars: 1500,
    maxPieces: 12,
    maxLength: 2800,
  },
  partial: {
    minSentences: 6,
    minCharsWarn: 640,
    targetMinChars: 640,
    targetMaxChars: 1000,
    maxPieces: 8,
    maxLength: 2800,
  },
  unknown: {
    minSentences: 6,
    minCharsWarn: 760,
    targetMinChars: 760,
    weakEvidenceMinChars: 540,
    targetMaxChars: 1300,
    maxPieces: 10,
    maxLength: 2800,
  },
};

const BASE_FORBIDDEN_TERMS = [
  "EBX-Q",
  "Source",
  "PDF",
  "page",
  "reviewer",
  "audit",
  "trace",
  "metadata",
  "quantitative",
  "TC-SC-",
  "GRI",
  "ESRS",
  "TCFD",
];

const METRIC_RULES = {
  1: {
    required: true,
    regexes: ["carbon", "Scope", "RE100", "water", "recycled", "tanso", "sujawon", "jaehwalyong"],
    topic: "strategy target or roadmap metric",
  },
  4: {
    required: false,
    regexes: ["LTIR", "injur", "accident", "safety", "anjeon"],
    topic: "health and safety metric",
  },
  7: {
    required: true,
    regexes: ["LTIR", "injur", "accident", "safety", "anjeon"],
    topic: "health and safety performance metric",
  },
  11: {
    required: true,
    regexes: ["grievance", "processing", "gochung"],
    topic: "grievance metric",
  },
  15: {
    required: true,
    regexes: ["service centers", "complaint", "consumer", "VOC", "product", "quality", "service"],
    negativeRegexes: ["Compliance", "Anti-fraud", "Privacy", "GHG", "Scope"],
    topic: "customer service or complaint metric",
  },
  19: {
    required: true,
    regexes: ["privacy", "information request", "gaeinjeongbo", "jeongbo"],
    negativeRegexes: ["Compliance", "Anti-fraud", "GHG", "Scope", "LTIR"],
    topic: "privacy or information request metric",
  },
  23: {
    required: true,
    regexes: ["GHG", "Scope", "emission", "energy", "waste", "water", "ongas", "baechul", "energy", "waste", "yongsu"],
    topic: "environmental performance metric",
  },
  27: {
    required: true,
    regexes: ["compliance", "fraud", "ethics", "report", "compliance", "bujeong", "yunli", "junbeop"],
    topic: "ethics or compliance performance metric",
  },
};

function ebxNumber(rowOrEbx) {
  return Number(String(rowOrEbx?.ebx ?? rowOrEbx ?? "").match(/\d+/)?.[0] ?? 0);
}

function answerTypeFromEbx(ebx) {
  const num = ebxNumber(ebx);
  if ([1, 8, 16, 20, 24].includes(num)) return "strategy-policy";
  if ([2, 5, 9, 13, 17, 21, 25].includes(num)) return "governance-organization";
  if ([3, 6, 10, 14, 18, 22, 26].includes(num)) return "risk-control";
  if ([7, 11, 15, 19, 23, 27].includes(num)) return "status-performance";
  return "policy-management";
}

function preferredStyleFor(answerType) {
  if (answerType === "strategy-policy") return "narrative";
  if (answerType === "governance-organization") return "governance";
  if (answerType === "risk-control") return "risk-control";
  if (answerType === "status-performance") return "status-performance";
  return "balanced-policy";
}

function sentencePlanFor(answerType) {
  const plans = {
    "strategy-policy": {
      openingProfile: "business-direction",
      closingProfile: "goal-governance-execution",
      fallbackProfile: "policy-management",
    },
    "governance-organization": {
      openingProfile: "owner-decision-flow",
      closingProfile: "owner-reporting-supervision",
      fallbackProfile: "policy-management",
    },
    "risk-control": {
      openingProfile: "risk-prevention-control",
      closingProfile: "prevention-monitoring-followup",
      fallbackProfile: "policy-management",
    },
    "status-performance": {
      openingProfile: "status-and-performance",
      closingProfile: "scale-trend-management-gap",
      fallbackProfile: "policy-management",
    },
    "policy-management": {
      openingProfile: "policy-standard-procedure",
      closingProfile: "policy-owner-performance",
      fallbackProfile: "policy-management",
    },
  };
  return plans[answerType] ?? plans["policy-management"];
}

function evidenceSelectionFor(row) {
  return {
    keywordFields: ["item", "question_title"],
    minEvidenceSentences: 2,
    maxEvidenceSentences: 12,
    preferNumbersWhenMetricExpected: Boolean(METRIC_RULES[ebxNumber(row)]?.required),
    rejectSourceTrace: true,
    rejectNavigationText: true,
    rejectLongListFragments: true,
  };
}

function enrichRow(row) {
  const answerType = row.answerType || row["Answer Type"] || answerTypeFromEbx(row.ebx);
  const num = ebxNumber(row);
  const evidenceSelection = {
    ...evidenceSelectionFor(row),
    ...(row.evidenceSelection ?? {}),
    maxEvidenceSentences: 12,
  };
  const metricHints = {
    required: false,
    regexes: [],
    negativeRegexes: [],
    minScore: 8,
    maxRecords: 2,
    topic: "",
    ...(METRIC_RULES[num] ?? {}),
    ...(row.metricHints ?? {}),
  };
  metricHints.maxRecords = metricHints.required
    ? Math.max(Number(metricHints.maxRecords ?? 2), 5)
    : Number(metricHints.maxRecords ?? 2);
  return {
    ...row,
    answerType,
    preferredStyle: row.preferredStyle || preferredStyleFor(answerType),
    styleOptions: row.styleOptions || ["formal", "concise", "evidence-led", "narrative"],
    sentencePlan: row.sentencePlan || sentencePlanFor(answerType),
    evidenceSelection,
    metricHints,
    qaRules: {
      ...(row.qaRules ?? {}),
      ...BASE_QA_RULES,
    },
    lengthPolicy: BASE_LENGTH_POLICY,
    forbiddenTerms: row.forbiddenTerms || BASE_FORBIDDEN_TERMS,
  };
}

async function ensureDirs() {
  await fs.mkdir(templateJsonDir, { recursive: true });
}

function v9TemplateName(file) {
  return file.replace(/consultant_safe_v8\.json$/i, `${RELEASE_NAME}.json`);
}

async function buildJsonTemplate(sourceFile) {
  const sourcePath = path.join(sourceTemplateDir, sourceFile);
  const template = JSON.parse(await fs.readFile(sourcePath, "utf8"));
  const outputName = v9TemplateName(sourceFile);
  const output = {
    ...template,
    releaseName: RELEASE_NAME,
    sourceRelease: SOURCE_RELEASE,
    sourceJsonTemplate: path.join("templates", sourceFile),
    ruleModel: {
      version: "template-rules-v1",
      purpose: "Move row-specific generation, metric, and QA decisions into JSON data while keeping code as the execution engine.",
    },
    rows: (template.rows ?? []).map(enrichRow),
  };
  await fs.writeFile(path.join(templateJsonDir, outputName), `${JSON.stringify(output, null, 2)}\n`, "utf8");
  return {
    sector: template.sector ?? "",
    sectorName: template.sectorName ?? "",
    size: template.size ?? "",
    sourceJsonTemplate: path.join("templates", sourceFile),
    jsonTemplate: path.join("templates", outputName),
    rows: output.rows.length,
  };
}

async function writeDocs(summary) {
  const readme = `# Consultant-safe EBX-Q Template v9

## Purpose
V9 keeps v8 JSON templates and doubles the Final Answer length policy while keeping the prose evidence-bounded and customer-safe.

## What Changed
- Raises evidence selection from 6 to 12 candidate sentences.
- Raises sufficient-answer QA targets to 8 sentences and 900 Korean characters.
- Doubles coverage-aware answer length targets from v8.
- Allows at least 5 metric records for metric-heavy rows.
- Keeps the six-column customer output schema unchanged.
- Keeps JSON templates as the source of truth for row-level behavior.
- Keeps JavaScript as the execution engine for reading data, applying rules, QA, and exporting Excel.

## Build Stats
- Source v8 JSON templates read: ${summary.sourceTemplates}
- v9 JSON templates generated: ${summary.jsonTemplates}
- Fatal issues: ${summary.fatal.length}
- Warnings: ${summary.warnings.length}
`;

  const checklist = `# consultant_safe_v9 Handoff Checklist

- [ ] Use JSON files in \`consultant_safe_v9/templates\` as the template source-of-truth.
- [ ] Prefer editing row-level rule fields before editing fill code.
- [ ] Customer workbook preserves the six established output columns.
- [ ] Sidecar QA JSON has zero fatal findings before handoff.
- [ ] Review QA length summary for short sufficient rows, below-double rows, and generic top-up rows.
`;

  const changelog = `# Version Log

## consultant_safe_v9
- Source release: \`${SOURCE_RELEASE}\`.
- Doubles coverage-aware answer length policy from v8.
- Raises evidence and metric budgets for longer evidence-backed Final Answer prose.
- Keeps v8 safety controls and customer workbook schema.
`;

  await fs.writeFile(path.join(releaseDir, "README.md"), readme, "utf8");
  await fs.writeFile(path.join(releaseDir, "CHECKLIST.md"), checklist, "utf8");
  await fs.writeFile(path.join(releaseDir, "CHANGELOG.md"), changelog, "utf8");
}

async function main() {
  await ensureDirs();
  const sourceFiles = (await fs.readdir(sourceTemplateDir))
    .filter((file) => file.endsWith("_consultant_safe_v8.json"))
    .sort();

  const templates = [];
  for (const file of sourceFiles) {
    templates.push(await buildJsonTemplate(file));
  }

  const fatal = [];
  const warnings = [];
  if (sourceFiles.length !== 44) fatal.push(`Expected 44 v8 JSON templates, found ${sourceFiles.length}.`);
  if (templates.length !== 44) fatal.push(`Expected 44 v9 JSON templates, generated ${templates.length}.`);
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
    sourceTemplates: sourceFiles.length,
    jsonTemplates: templates.length,
    fatal,
    warnings,
  };
  await fs.writeFile(path.join(releaseDir, "QA_REPORT.json"), `${JSON.stringify(qa, null, 2)}\n`, "utf8");
  await writeDocs({ sourceTemplates: sourceFiles.length, jsonTemplates: templates.length, fatal, warnings });

  console.log(JSON.stringify(qa, null, 2));
  if (fatal.length) process.exitCode = 1;
}

await main();
