import fs from "node:fs/promises";
import path from "node:path";
import crypto from "node:crypto";
import { fileURLToPath } from "node:url";

const __dirname = path.dirname(fileURLToPath(import.meta.url));
const repoRoot = path.resolve(__dirname, "..");
const sourceDir = path.join(repoRoot, "consultant_safe_v10");
const releaseDir = path.join(repoRoot, "consultant_safe_v10_1");
const sourceTemplateDir = path.join(sourceDir, "templates");
const templateDir = path.join(releaseDir, "templates");
const RELEASE_NAME = "consultant_safe_v10_1";
const SOURCE_RELEASE = "consultant_safe_v10";
const EXPECTED_TEMPLATE_COUNT = 44;
const EXPECTED_ROW_COUNT = 95;

const ARTIFACT_EVIDENCE_GUARDS = [
  "보고\\s*프레임워크",
  "보고\\s*위치",
  "보고기간",
  "SPECIAL\\s*REPORT",
  "ESG\\s*DATA\\s*PACK",
  "TCFD",
  "GRI",
  "SASB",
  "Global\\s*Standard",
  "Topic\\s*Standard",
  "Material\\s*Topic",
  "다양성 및 인권경영\\s+DEI",
  "안전보건\\s+안전보건",
  "정보보호 및 개인정보보호\\s+정보보호",
  "윤리준법경영\\s+윤리준법경영",
  "임직원 및 이사회 대상 환경 교육",
  "옥상 태양열 설비",
  "건축물 일체형 태양광",
  "냉각 효율 제고.*전력 효율지수",
  "^\\s*\\d+\\s*\\([a-z]\\)",
  "^\\s*p\\.\\s*\\d+",
];

const ARTIFACT_METRIC_GUARDS = [
  "SPECIAL\\s*REPORT",
  "ESG\\s*DATA\\s*PACK",
  "TCFD",
  "GRI",
  "SASB",
  "Global\\s*Standard",
  "Topic\\s*Standard",
  "Material\\s*Topic",
  "보고\\s*위치",
  "보고기간",
  "p\\.\\s*\\d+",
  "^\\s*\\d+\\s*\\([a-z]\\)",
  "^\\s*\\d+[\\).]\\s",
  "임직원 및 이사회 대상 환경 교육",
  "옥상 태양열 설비",
  "건축물 일체형 태양광",
  "냉각 효율 제고.*전력 효율지수",
];

const ARTIFACT_AVOID_SENTENCE =
  "Do not keep report section headings, question prompts, framework/index labels, page-reference rows, or equipment/table fragments in Final Answer.";

function v101Name(file) {
  return file.replace(/consultant_safe_v10\.json$/i, "consultant_safe_v10_1.json");
}

function operationalProjection(row) {
  return {
    ebx: row.ebx,
    answerType: row.answerType,
    evidenceSelection: row.evidenceSelection,
    metricHints: row.metricHints,
    qaRules: row.qaRules,
    lengthPolicy: row.lengthPolicy,
    sizeProfile: row.sizeProfile,
    sectorProfile: row.sectorProfile,
  };
}

function hashOperationalRows(rows) {
  return crypto
    .createHash("sha256")
    .update(JSON.stringify(rows.map(operationalProjection)))
    .digest("hex");
}

function appendUnique(target, additions) {
  const seen = new Set(target);
  for (const item of additions) {
    if (!seen.has(item)) {
      target.push(item);
      seen.add(item);
    }
  }
  return target;
}

function appendAvoidList(current) {
  const text = String(current || "").trim();
  if (text.includes("question prompts, framework/index labels")) return text;
  return text ? `${text}\n${ARTIFACT_AVOID_SENTENCE}` : ARTIFACT_AVOID_SENTENCE;
}

function applyArtifactGuards(template) {
  for (const row of template.rows ?? []) {
    row.evidenceSelection ??= {};
    row.evidenceSelection.negativeTopicHints = appendUnique(
      Array.isArray(row.evidenceSelection.negativeTopicHints)
        ? row.evidenceSelection.negativeTopicHints
        : [],
      ARTIFACT_EVIDENCE_GUARDS,
    );

    row.metricHints ??= {};
    row.metricHints.negativeRegexes = appendUnique(
      Array.isArray(row.metricHints.negativeRegexes) ? row.metricHints.negativeRegexes : [],
      ARTIFACT_METRIC_GUARDS,
    );

    row.avoidList = appendAvoidList(row.avoidList);
    if (row["Plain-Language Avoid List"] !== undefined) {
      row["Plain-Language Avoid List"] = appendAvoidList(row["Plain-Language Avoid List"]);
    }
  }
  return template;
}

async function main() {
  const sourceIndex = JSON.parse(await fs.readFile(path.join(sourceDir, "TEMPLATE_INDEX.json"), "utf8"));
  await fs.rm(releaseDir, { recursive: true, force: true });
  await fs.mkdir(templateDir, { recursive: true });
  await fs.cp(path.join(sourceDir, "source"), path.join(releaseDir, "source"), { recursive: true });
  await fs.cp(path.join(sourceDir, "policies"), path.join(releaseDir, "policies"), { recursive: true });

  const fatal = [];
  const templates = [];
  for (const entry of sourceIndex.templates ?? []) {
    const sourceFile = path.basename(entry.jsonTemplate);
    const outputFile = v101Name(sourceFile);
    const sourceTemplate = JSON.parse(await fs.readFile(path.join(sourceTemplateDir, sourceFile), "utf8"));
    const outputTemplate = {
      ...sourceTemplate,
      releaseName: RELEASE_NAME,
      sourceRelease: SOURCE_RELEASE,
      sourceJsonTemplate: path.join("templates", sourceFile),
      ruleModel: {
        ...(sourceTemplate.ruleModel ?? {}),
        version: "template-rules-v2.1",
        purpose: "Preserve the v10 95-item/profile model while applying v10.1 metric-selection, QA, and workbook-layout safety controls.",
      },
    };
    applyArtifactGuards(outputTemplate);
    if ((outputTemplate.rows ?? []).length !== EXPECTED_ROW_COUNT) {
      fatal.push(`${sourceFile}: expected ${EXPECTED_ROW_COUNT} rows.`);
    }
    await fs.writeFile(path.join(templateDir, outputFile), `${JSON.stringify(outputTemplate, null, 2)}\n`, "utf8");
    templates.push({
      ...entry,
      sourceJsonTemplate: path.join("templates", sourceFile),
      jsonTemplate: path.join("templates", outputFile),
      operationalHash: hashOperationalRows(outputTemplate.rows ?? []),
    });
  }
  if (templates.length !== EXPECTED_TEMPLATE_COUNT) fatal.push(`Expected ${EXPECTED_TEMPLATE_COUNT} templates, found ${templates.length}.`);

  const index = {
    ...sourceIndex,
    generatedAt: new Date().toISOString(),
    releaseName: RELEASE_NAME,
    sourceRelease: SOURCE_RELEASE,
    templates,
  };
  const qa = {
    generatedAt: new Date().toISOString(),
    releaseName: RELEASE_NAME,
    sourceRelease: SOURCE_RELEASE,
    sourceWorkbook: index.sourceWorkbook,
    masterRows: EXPECTED_ROW_COUNT,
    sourceTemplates: (sourceIndex.templates ?? []).length,
    jsonTemplates: templates.length,
    profileVersion: index.profileVersion,
    sizeProfiles: [...new Set(templates.map((entry) => entry.size))],
    sectorProfiles: [...new Set(templates.map((entry) => entry.sector))],
    operationalVariationValidated: sourceIndex.templates?.every((entry) => entry.operationalHash),
    fatal,
    warnings: [],
  };
  const readme = `# Consultant-safe EBX-Q Template v10.1

V10.1 preserves all 95 EBX rows, 44 operational templates, and the six customer workbook columns from v10.

## Improvements
- Prefer exact mapped-item metrics and topic-specific indicators.
- Reject control characters, prose/table fragments, and year-as-value metric artifacts.
- Fail QA on exact duplicate final answers and malformed metric sentences.
- Keep all six columns while bounding row height and wrapping only user-readable content columns.

## Validation
- \`node scripts/validate_consultant_safe_v10_1.mjs\`
- \`node scripts/test_consultant_safe_v10_1_profiles.mjs\`
- \`node scripts/regress_consultant_safe_v10_1_large.mjs\`
`;
  const checklist = `# consultant_safe_v10_1 Handoff Checklist

- [ ] Exactly 44 templates and 95 rows per template.
- [ ] Excel headers remain: EBX Indicator, Field, Original Answer, Original Answer Metadata, Style Template Applied, Final Answer.
- [ ] No control characters, malformed metrics, exact duplicate answers, formula errors, fail, or fatal rows.
- [ ] All 11 large-company regression cases stay within their warning baselines.
- [ ] Full-sheet render remains bounded and readable.
`;
  const changelog = `# Version Log

## consultant_safe_v10_1
- Source release: \`${SOURCE_RELEASE}\`.
- Preserves the v10 template/profile matrix and six-column Excel schema.
- Adds exact metric mapping priority and safer metric record filtering.
- Adds malformed-metric, control-character, and duplicate-answer QA gates.
- Bounds Excel row height without removing, hiding, or renaming columns.
`;

  await fs.writeFile(path.join(releaseDir, "TEMPLATE_INDEX.json"), `${JSON.stringify(index, null, 2)}\n`, "utf8");
  await fs.writeFile(path.join(releaseDir, "QA_REPORT.json"), `${JSON.stringify(qa, null, 2)}\n`, "utf8");
  await fs.writeFile(path.join(releaseDir, "README.md"), readme, "utf8");
  await fs.writeFile(path.join(releaseDir, "CHECKLIST.md"), checklist, "utf8");
  await fs.writeFile(path.join(releaseDir, "CHANGELOG.md"), changelog, "utf8");
  console.log(JSON.stringify(qa, null, 2));
  if (fatal.length) process.exitCode = 1;
}

await main();
