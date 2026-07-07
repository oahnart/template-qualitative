import fs from "node:fs/promises";
import path from "node:path";
import crypto from "node:crypto";
import { fileURLToPath } from "node:url";
import { FileBlob, SpreadsheetFile } from "@oai/artifact-tool";

const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);
const repoRoot = path.resolve(__dirname, "..");
const sourceDir = path.join(repoRoot, "consultant_safe_v9");
const sourceTemplateDir = path.join(sourceDir, "templates");
const releaseDir = path.join(repoRoot, "consultant_safe_v10");
const templateJsonDir = path.join(releaseDir, "templates");
const masterWorkbookPath = path.join(releaseDir, "source", "26.07.06 ESG-정성_v2.xlsx");
const profileConfigPath = path.join(releaseDir, "policies", "template_profiles.json");
const PROFILE_CONFIG = JSON.parse(await fs.readFile(profileConfigPath, "utf8"));

const RELEASE_NAME = "consultant_safe_v10";
const SOURCE_RELEASE = "consultant_safe_v9";
const EXPECTED_TEMPLATE_COUNT = 44;
const EXPECTED_ROW_COUNT = 95;
const OLD_ROW_COUNT = 27;
const MASTER_HEADERS = ["영역", "카테고리", "구분 (4 Pillars)", "항목", "설명", "예시"];
const EXPECTED_SIZES = ["대기업", "중견", "중소", "비상장"];
const EXPECTED_SECTORS = ["CG", "EM", "FB", "FN", "HC", "IF", "RR", "RT", "SV", "TC", "TR"];

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

function ebxId(index) {
  return `EBX-Q-${String(index + 1).padStart(3, "0")}`;
}

function normalize(value) {
  return String(value ?? "").replace(/\s+/g, " ").trim();
}

function escapeRegex(value) {
  return String(value).replace(/[.*+?^${}()|[\]\\]/g, "\\$&");
}

function unique(values) {
  return [...new Set((values ?? []).map(normalize).filter(Boolean))];
}

function mergeLengthPolicies(...policies) {
  const result = {};
  for (const policy of policies.filter(Boolean)) {
    for (const [coverage, values] of Object.entries(policy)) {
      result[coverage] = { ...(result[coverage] ?? {}), ...(values ?? {}) };
    }
  }
  return result;
}

function topicProfileFor(row) {
  const number = Number(String(row.ebx ?? "").match(/\d+/)?.[0] ?? 0);
  return PROFILE_CONFIG.topicProfiles.find((profile) => number >= profile.from && number <= profile.to) ?? null;
}

function resolvedProfiles(row, templateMeta, answerType) {
  const size = PROFILE_CONFIG.sizeProfiles[templateMeta.size];
  const sector = PROFILE_CONFIG.sectorProfiles[templateMeta.sector];
  if (!size) throw new Error(`Missing size profile: ${templateMeta.size}`);
  if (!sector) throw new Error(`Missing sector profile: ${templateMeta.sector}`);
  const sizeLengthPolicy = mergeLengthPolicies(size.lengthPolicy.default, size.lengthPolicy[answerType]);
  return { size, sector, topic: topicProfileFor(row), sizeLengthPolicy };
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
  return crypto.createHash("sha256")
    .update(JSON.stringify(rows.map(operationalProjection)))
    .digest("hex");
}

function answerTypeFromPillar(pillar) {
  if (/Strategy|전략/i.test(pillar)) return "strategy-policy";
  if (/Governance|거버넌스/i.test(pillar)) return "governance-organization";
  if (/Risk Management|위험 관리/i.test(pillar)) return "risk-control";
  if (/Metrics|지표/i.test(pillar)) return "status-performance";
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

function evidenceSelectionFor(row, answerType) {
  return {
    keywordFields: ["item", "question_title"],
    minEvidenceSentences: 2,
    maxEvidenceSentences: 12,
    preferNumbersWhenMetricExpected: answerType === "status-performance",
    rejectSourceTrace: true,
    rejectNavigationText: true,
    rejectLongListFragments: true,
    topicHints: [row.category, row.item].filter(Boolean),
  };
}

function metricRegexesFor(row, answerType) {
  const genericTokens = /^(?:환경|사회|거버넌스|관리|현황|체계|목표|정책|활동|관련|운영|성과|이행|책임|전략|리스크|지표)$/i;
  const base = [row.category, row.item]
    .flatMap((value) => normalize(value).split(/[\s/·,()]+/))
    .filter((token) => token.length >= 2)
    .filter((token) => !genericTokens.test(token))
    .slice(0, 8)
    .map(escapeRegex);
  const thematic = [];
  const haystack = `${row.category} ${row.item} ${row.description}`;
  if (/온실가스|탄소|기후|에너지/i.test(haystack)) thematic.push("GHG", "Scope", "emission", "energy", "온실가스", "탄소", "에너지");
  if (/안전|재해|보건/i.test(haystack)) thematic.push("LTIR", "injur", "accident", "safety", "안전", "재해");
  if (/인권|고충|노동/i.test(haystack)) thematic.push("grievance", "human rights", "labor", "인권", "고충", "노동");
  if (/제품|품질|고객|소비자/i.test(haystack)) thematic.push("product", "quality", "complaint", "customer", "품질", "고객", "소비자");
  if (/정보|보안|개인정보/i.test(haystack)) thematic.push("privacy", "security", "information", "개인정보", "정보보호", "보안");
  if (/폐기물|자원|순환/i.test(haystack)) thematic.push("waste", "recycled", "폐기물", "재활용", "자원");
  if (/물|수자원|폐수|수질/i.test(haystack)) thematic.push("water", "wastewater", "용수", "수자원", "폐수");
  if (/윤리|준법|컴플라이언스|부패/i.test(haystack)) thematic.push("compliance", "fraud", "ethics", "윤리", "준법", "부패");
  return [...new Set([...thematic.map(escapeRegex), ...base])].slice(0, answerType === "status-performance" ? 14 : 8);
}

function metricHintsFor(row, answerType) {
  const required = answerType === "status-performance";
  return {
    required,
    regexes: metricRegexesFor(row, answerType),
    negativeRegexes: [],
    minScore: required ? 8 : 10,
    maxRecords: required ? 5 : 2,
    topic: `${row.category} ${row.item}`.trim(),
  };
}

function contentSlotsFor(row, answerType) {
  const facts = answerType === "status-performance"
    ? ["핵심 성과지표(KPI) 정의", "최근 실적 수치", "목표 대비 달성도", "전년 대비 추세"]
    : answerType === "risk-control"
      ? ["주요 리스크 정의", "식별 및 평가 절차", "예방·완화 활동", "모니터링 및 후속 조치"]
      : answerType === "governance-organization"
        ? ["담당 조직", "의사결정 권한", "보고 및 감독 체계", "역할과 책임"]
        : ["추진 배경·중요성", "정책·방침 선언", "중장기 목표", "이행 로드맵"];
  return [
    `• Disclosure item: ${row.item}`,
    ...facts.map((fact, index) => `• Required material ${index + 1}: ${fact}`),
    `• Category: ${row.category}`,
    `• Description: ${row.description}`,
    `• Company data slots to fill: [회사명], [기준연도], [담당 조직], [핵심 활동], [성과 또는 미집계 상태]`,
    "• Do not convert these slots into a final sentence until company-specific facts are supplied.",
  ].join("\n");
}

function instructionsFor(row) {
  return [
    "• Do not write the final answer by copying one fixed sentence. Choose a style option, then compose from slots and verified evidence.",
    "• [ ] 안의 회사 실제 값·명칭을 내부 문서에 근거해 채울 것(가상 수치 금지).",
    `• ${row.description}`,
    "• 원천에 없는 내용은 지어내지 말 것 — 미보유 시 '해당사항 없음/미집계' 명시.",
    "• If metric data is not tracked, write 미집계 or 해당사항 없음; do not estimate or invent figures.",
  ].join("\n");
}

function sentencePatternsFor(row) {
  return [
    `• Pattern A (${row.ebx} ${row.item}): [기준연도/범위] 기준 [핵심 사실]을 먼저 제시하고 [담당 조직/관리 방식]으로 연결한다.`,
    `• Pattern B (${row.ebx} ${row.item}): [배경/중요성] -> [정책/프로세스] -> [실행/성과] 순서로 구성한다.`,
    `• Pattern C (${row.ebx} ${row.item}): [담당 조직/책임자]를 주어로 시작하고 [활동] [주기] [개선 계획]을 이어간다.`,
    `• Pattern D (${row.ebx} ${row.item}): [측정값/미집계 상태] -> [산식/단위] -> [집계 범위] -> [검증 근거] 순서로 쓴다.`,
  ].join("\n");
}

function antiRepetitionRulesFor(row) {
  return [
    "• Preferred variation for this row: context -> policy -> activity -> evidence; avoid using the same order as the previous item.",
    "• Opening choice: start with responsible organization; do not start every answer with 회사는/당사는.",
    "• Limit repeated endings such as 추진하고 있습니다, 관리하고 있습니다, 운영하고 있습니다 across nearby rows.",
    "• Change sentence length: mix one short factual sentence with one evidence/detail sentence when enough data exists.",
    `• Never copy the example for '${row.item}' as final wording; it is only a scope reference.`,
  ].join("\n");
}

function answerIntentFor(answerType) {
  if (answerType === "strategy-policy") return "Explain the policy or strategic direction first, then connect targets, governance, and execution evidence.";
  if (answerType === "governance-organization") return "Explain the responsible organization, decision rights, reporting flow, and oversight structure.";
  if (answerType === "risk-control") return "Explain the risk identification, prevention controls, monitoring, and follow-up actions.";
  if (answerType === "status-performance") return "Explain the reported status in plain language, then weave figures into the management meaning.";
  return "Explain the management approach, owner, operating procedure, and evidence boundary.";
}

function openingStrategyFor(answerType) {
  if (answerType === "strategy-policy") return "Start from business context or strategic direction, then connect to goals and execution.";
  if (answerType === "governance-organization") return "Start with the owner or decision body, then describe reporting and accountability.";
  if (answerType === "risk-control") return "Start with the material risk or control objective, then explain prevention and monitoring.";
  if (answerType === "status-performance") return "Start with status/performance scope, then present figures only when supported.";
  return "Start with policy scope, then explain the operating method and limitation.";
}

function evidenceWeaveFor(row) {
  return [
    `Field Path: ${row.area} / ${row.category} / ${row.pillar} / ${row.item}`,
    "Use 2-4 row-specific evidence sentences, repaired for obvious OCR spacing.",
    "If a figure is needed, introduce it after the business context; avoid formulaic metric-first openings.",
    "Reject page labels, report navigation, table/index rows, raw headings, and long list fragments.",
    "Intent: Explain the reported status in plain language, then weave figures into the management meaning.",
  ].join("\n");
}

function requiredFactsFor(row, answerType) {
  const core = ["회사명", "기준연도", "보고 범위", "담당 조직"];
  if (answerType === "status-performance") core.push("지표명", "실적값", "단위", "집계 기준");
  if (answerType === "risk-control") core.push("식별 리스크", "예방 활동", "후속 조치");
  if (answerType === "governance-organization") core.push("의사결정 기구", "보고 경로", "감독 책임");
  if (answerType === "strategy-policy") core.push("정책·목표", "목표연도", "이행 과제");
  return core.join(", ");
}

function avoidListFor() {
  return [
    "Do not use: quantitative, 정량, 정량 지표, 정량 목표, định lượng.",
    "Do not use source/page/PDF/audit/reviewer wording in Final Answer.",
    "Do not expose EBX codes, raw indicator names, or template terms.",
    "Do not keep OCR/table headings such as Corporate Governance, 구분, 추진체계와 주요성과, 코드 공시 항목.",
    "If evidence is missing, state the limitation plainly and do not invent company-specific facts.",
  ].join("\n");
}

function styleGuardrailsFor(row) {
  return [
    `Write in Korean ESG-report style for ${row.item}.`,
    "Use company-specific evidence first; if unavailable, mark the information gap plainly.",
    "Do not reuse example wording as if it were company evidence.",
  ].join("\n");
}

function industryScopeReferenceFor(row) {
  return `예) 이 항목('${row.item}')에서는 ${row.category} 범위에서 ${row.description} 회사별 확인 자료가 있는 경우 정책, 조직, 리스크 관리, 성과를 구분해 기술합니다. (예시·실데이터 아님)`;
}

function industryRewriteReferenceFor(row) {
  return row.example ? `${row.example} (작성 참고용 예시·실데이터 아님)` : "";
}

function sizeGuidanceFor(size) {
  const profile = PROFILE_CONFIG.sizeProfiles[size];
  return [
    `• ${size || "대기업"}: ${profile?.guidance ?? "Use proportionate company-specific disclosure."}`,
    size === "비상장"
      ? "• Use available internal or voluntary disclosure evidence; do not assume listed-company filing duties."
      : "• Use public evidence when available, then reconcile it with internal source documents before submission.",
  ].join("\n");
}

function createGeneratedRow(masterRow, templateMeta) {
  const answerType = answerTypeFromPillar(masterRow.pillar);
  const row = {
    ebx: masterRow.ebx,
    area: masterRow.area,
    category: masterRow.category,
    pillar: masterRow.pillar,
    item: masterRow.item,
    description: masterRow.description,
    example: masterRow.example,
    "공개성": "🔴내부(서술)",
    "작성 소재/Content Slots": contentSlotsFor(masterRow, answerType),
    "작성지침": instructionsFor(masterRow),
    "문체 옵션/Style Options": "• formal: ESG report tone; cover full governance/KPI/evidence detail; suitable when source evidence is complete.\n• concise: 1-2 compact sentences; use when the company has limited evidence or the item only needs a factual status.\n• evidence-led: start with reporting period/scope, then explain the policy, activity, or metric; best for reviewer traceability.\n• narrative: start from business context or materiality, then connect to governance, action, and next step; best for strategy/policy items.",
    "문장 패턴/Sentence Patterns": sentencePatternsFor(masterRow),
    "반복 방지/Anti-Repetition Rules": antiRepetitionRulesFor(masterRow),
    "동종산업 범위 참고": industryScopeReferenceFor(masterRow),
    "동종산업 재서술 참고": industryRewriteReferenceFor(masterRow),
    "규모지침": sizeGuidanceFor(templateMeta.size),
    "비고": "v10 master 95-item template generated from 26.07.06 ESG-정성_v2.xlsx.",
    "Answer Type": answerType,
    answerType,
    "Answer Intent": answerIntentFor(answerType),
    answerIntent: answerIntentFor(answerType),
    "Opening Strategy": openingStrategyFor(answerType),
    openingStrategy: openingStrategyFor(answerType),
    "Evidence Weave": evidenceWeaveFor(masterRow),
    evidenceWeave: evidenceWeaveFor(masterRow),
    "Required Facts": requiredFactsFor(masterRow, answerType),
    requiredFacts: requiredFactsFor(masterRow, answerType),
    "Plain-Language Avoid List": avoidListFor(),
    avoidList: avoidListFor(),
    "Style Guardrails": styleGuardrailsFor(masterRow),
    styleGuardrails: styleGuardrailsFor(masterRow),
    "QA Severity": answerType === "status-performance" ? "metric-sensitive" : "standard",
    qaSeverity: answerType === "status-performance" ? "metric-sensitive" : "standard",
  };
  return enrichRow(row, templateMeta);
}

function enrichRow(row, templateMeta) {
  const answerType = row.answerType || row["Answer Type"] || answerTypeFromPillar(row.pillar);
  const profiles = resolvedProfiles(row, templateMeta, answerType);
  const overrides = row.policyOverrides ?? {};
  const topicHints = profiles.topic?.evidenceHints ?? [];
  const negativeTopicHints = profiles.topic?.negativeHints ?? [];
  const evidenceSelection = {
    ...evidenceSelectionFor(row, answerType),
    ...(row.evidenceSelection ?? {}),
    topicHints: unique([
      ...(row.evidenceSelection?.topicHints ?? []),
      ...topicHints,
    ]),
    negativeTopicHints: unique([
      ...(row.evidenceSelection?.negativeTopicHints ?? []),
      ...negativeTopicHints,
    ]),
    sectorHints: unique(profiles.sector.evidenceHints),
    ...profiles.size.evidenceSelection,
    ...(overrides.evidenceSelection ?? {}),
  };
  const metricHints = {
    ...metricHintsFor(row, answerType),
    ...(row.metricHints ?? {}),
    regexes: unique([
      ...(row.metricHints?.regexes ?? metricRegexesFor(row, answerType)),
      ...(profiles.topic?.metricRegexes ?? []),
    ]),
    negativeRegexes: unique([
      ...(row.metricHints?.negativeRegexes ?? []),
      ...(profiles.topic?.negativeMetricRegexes ?? []),
    ]),
    sectorHints: unique(profiles.sector.metricHints),
    ...profiles.size.metricHints,
    ...(overrides.metricHints ?? {}),
  };
  metricHints.required = answerType === "status-performance";
  const lengthPolicy = mergeLengthPolicies(
    BASE_LENGTH_POLICY,
    row.lengthPolicy,
    profiles.sizeLengthPolicy,
    overrides.lengthPolicy,
  );
  const sufficientPolicy = lengthPolicy.sufficient ?? {};
  const qaRules = {
    ...BASE_QA_RULES,
    ...(row.qaRules ?? {}),
    minSentences: Number(sufficientPolicy.minSentences ?? 3),
    minCharsWarn: Number(sufficientPolicy.minCharsWarn ?? 270),
    ...(overrides.qaRules ?? {}),
  };
  const styleOptionsBySize = {
    "대기업": ["formal", "evidence-led", "narrative", "concise"],
    "중견": ["formal", "concise", "evidence-led", "narrative"],
    "중소": ["concise", "evidence-led", "formal"],
    "비상장": ["evidence-led", "concise", "formal", "narrative"],
  };
  const sizeGuardrail = templateMeta.size === "비상장"
    ? "Do not assume exchange filings, listed-company committees, or public-market reporting duties."
    : profiles.size.guidance;
  return {
    ...row,
    "규모지침": sizeGuidanceFor(templateMeta.size),
    "Answer Type": row["Answer Type"] || answerType,
    answerType,
    "Answer Intent": row["Answer Intent"] || answerIntentFor(answerType),
    answerIntent: row.answerIntent || row["Answer Intent"] || answerIntentFor(answerType),
    "Opening Strategy": row["Opening Strategy"] || openingStrategyFor(answerType),
    openingStrategy: row.openingStrategy || row["Opening Strategy"] || openingStrategyFor(answerType),
    "Evidence Weave": row["Evidence Weave"] || evidenceWeaveFor(row),
    evidenceWeave: row.evidenceWeave || row["Evidence Weave"] || evidenceWeaveFor(row),
    "Required Facts": row["Required Facts"] || requiredFactsFor(row, answerType),
    requiredFacts: row.requiredFacts || row["Required Facts"] || requiredFactsFor(row, answerType),
    "Plain-Language Avoid List": row["Plain-Language Avoid List"] || avoidListFor(),
    avoidList: row.avoidList || row["Plain-Language Avoid List"] || avoidListFor(),
    "Style Guardrails": `${row["Style Guardrails"] || styleGuardrailsFor(row)}\n${sizeGuardrail}`,
    styleGuardrails: `${row.styleGuardrails || row["Style Guardrails"] || styleGuardrailsFor(row)}\n${sizeGuardrail}`,
    "QA Severity": row["QA Severity"] || row.qaSeverity || (answerType === "status-performance" ? "metric-sensitive" : "standard"),
    qaSeverity: row.qaSeverity || row["QA Severity"] || (answerType === "status-performance" ? "metric-sensitive" : "standard"),
    preferredStyle: row.preferredStyle || preferredStyleFor(answerType),
    styleOptions: overrides.styleOptions || styleOptionsBySize[templateMeta.size],
    sentencePlan: row.sentencePlan || sentencePlanFor(answerType),
    evidenceSelection,
    metricHints,
    qaRules,
    lengthPolicy,
    sizeProfile: {
      id: templateMeta.size,
      description: profiles.size.description,
    },
    sectorProfile: {
      id: templateMeta.sector,
      name: profiles.sector.name,
      evidenceHints: unique(profiles.sector.evidenceHints),
      metricHints: unique(profiles.sector.metricHints),
    },
    forbiddenTerms: row.forbiddenTerms || BASE_FORBIDDEN_TERMS,
  };
}

async function readMasterRows() {
  const input = await FileBlob.load(masterWorkbookPath);
  const workbook = await SpreadsheetFile.importXlsx(input);
  const sheet = workbook.worksheets.getItemAt(0);
  const values = sheet.getRange(`A1:F${EXPECTED_ROW_COUNT + 1}`).values;
  const headers = values[0].map(normalize);
  return {
    headers,
    rows: values.slice(1).map((row, index) => ({
      ebx: ebxId(index),
      area: normalize(row[0]),
      category: normalize(row[1]),
      pillar: normalize(row[2]),
      item: normalize(row[3]),
      description: normalize(row[4]),
      example: normalize(row[5]),
    })),
  };
}

function validateMaster({ headers, rows }, referenceRows) {
  const fatal = [];
  const warnings = [];
  if (headers.join("|") !== MASTER_HEADERS.join("|")) {
    fatal.push(`Master headers mismatch. Expected ${MASTER_HEADERS.join(", ")}, found ${headers.join(", ")}.`);
  }
  if (rows.length !== EXPECTED_ROW_COUNT) fatal.push(`Expected ${EXPECTED_ROW_COUNT} master rows, found ${rows.length}.`);
  const blankCells = [];
  for (const [rowIndex, row] of rows.entries()) {
    for (const key of ["area", "category", "pillar", "item", "description", "example"]) {
      if (!row[key]) blankCells.push(`${row.ebx}.${key}`);
    }
    if (row.ebx !== ebxId(rowIndex)) fatal.push(`Unexpected EBX id at row ${rowIndex + 2}: ${row.ebx}.`);
  }
  if (blankCells.length) fatal.push(`Master contains blank cells: ${blankCells.slice(0, 20).join(", ")}${blankCells.length > 20 ? "..." : ""}`);
  const itemCounts = new Map();
  for (const row of rows) itemCounts.set(row.item, (itemCounts.get(row.item) ?? 0) + 1);
  const duplicates = [...itemCounts.entries()].filter(([item, count]) => item && count > 1);
  if (duplicates.length) fatal.push(`Duplicate master item names: ${duplicates.map(([item]) => item).join(", ")}`);
  for (let index = 0; index < OLD_ROW_COUNT; index += 1) {
    const master = rows[index];
    const old = referenceRows[index];
    if (!old) {
      fatal.push(`Missing v9 reference row ${index + 1}.`);
      continue;
    }
    if (master.ebx !== old.ebx || master.item !== old.item) {
      fatal.push(`Old row mismatch at ${master.ebx}: master '${master.item}' vs v9 '${old.ebx} ${old.item}'.`);
    }
  }
  const categories = rows.reduce((acc, row) => {
    acc[row.category] = (acc[row.category] ?? 0) + 1;
    return acc;
  }, {});
  if (categories["ESG 경영"] === 3) {
    warnings.push("ESG 경영 has 3 pillar rows by design; no Metrics row is added so total remains 95.");
  }
  return { fatal, warnings };
}

async function ensureDirs() {
  await fs.mkdir(templateJsonDir, { recursive: true });
}

function validateProfileConfig() {
  const fatal = [];
  const warnings = [];
  for (const size of EXPECTED_SIZES) {
    if (!PROFILE_CONFIG.sizeProfiles[size]) fatal.push(`Missing size profile: ${size}.`);
  }
  for (const sector of EXPECTED_SECTORS) {
    if (!PROFILE_CONFIG.sectorProfiles[sector]) fatal.push(`Missing sector profile: ${sector}.`);
  }
  const coveredRows = new Map();
  for (const profile of PROFILE_CONFIG.topicProfiles ?? []) {
    for (const value of [...(profile.metricRegexes ?? []), ...(profile.negativeMetricRegexes ?? [])]) {
      try {
        new RegExp(String(value), "i");
      } catch {
        fatal.push(`Invalid metric regex in topic ${profile.id}: ${value}`);
      }
    }
    for (let row = Number(profile.from); row <= Number(profile.to); row += 1) {
      if (coveredRows.has(row)) fatal.push(`Topic profiles overlap at EBX-Q-${String(row).padStart(3, "0")}.`);
      coveredRows.set(row, profile.id);
    }
  }
  for (let row = 28; row <= EXPECTED_ROW_COUNT; row += 1) {
    if (!coveredRows.has(row)) fatal.push(`No topic profile for EBX-Q-${String(row).padStart(3, "0")}.`);
  }
  if ((PROFILE_CONFIG.topicProfiles ?? []).length !== 17) {
    warnings.push(`Expected 17 topic profiles, found ${(PROFILE_CONFIG.topicProfiles ?? []).length}.`);
  }
  return { fatal, warnings };
}

function validateOperationalVariation(templates) {
  const fatal = [];
  for (const sector of EXPECTED_SECTORS) {
    const group = templates.filter((template) => template.sector === sector);
    if (group.length !== EXPECTED_SIZES.length) {
      fatal.push(`${sector}: expected ${EXPECTED_SIZES.length} size templates, found ${group.length}.`);
      continue;
    }
    const sizes = new Set(group.map((template) => template.size));
    const hashes = new Set(group.map((template) => template.operationalHash));
    if (sizes.size !== EXPECTED_SIZES.length) fatal.push(`${sector}: size labels are incomplete or duplicated.`);
    if (hashes.size !== EXPECTED_SIZES.length) fatal.push(`${sector}: operational size profiles are not all distinct.`);
  }
  const largeHashes = new Set(templates
    .filter((template) => template.size === "대기업")
    .map((template) => template.operationalHash));
  if (largeHashes.size !== EXPECTED_SECTORS.length) {
    fatal.push("Sector overlays do not produce 11 distinct large-company operational templates.");
  }
  return fatal;
}

function v10TemplateName(file) {
  return file.replace(/consultant_safe_v9\.json$/i, `${RELEASE_NAME}.json`);
}

function buildRowsForTemplate(template, masterRows) {
  const sourceRowsByItem = new Map((template.rows ?? []).map((row) => [row.item, row]));
  return masterRows.map((masterRow) => {
    const existing = sourceRowsByItem.get(masterRow.item);
    if (existing) {
      return enrichRow({
        ...existing,
        ebx: masterRow.ebx,
        area: masterRow.area,
        category: masterRow.category,
        pillar: masterRow.pillar,
        item: masterRow.item,
        description: masterRow.description,
        example: masterRow.example,
      }, template);
    }
    return createGeneratedRow(masterRow, template);
  });
}

async function buildJsonTemplate(sourceFile, masterRows) {
  const sourcePath = path.join(sourceTemplateDir, sourceFile);
  const template = JSON.parse(await fs.readFile(sourcePath, "utf8"));
  const outputName = v10TemplateName(sourceFile);
  const outputRows = buildRowsForTemplate(template, masterRows);
  const output = {
    ...template,
    releaseName: RELEASE_NAME,
    sourceRelease: SOURCE_RELEASE,
    sourceWorkbook: path.join("source", path.basename(masterWorkbookPath)),
    sourceJsonTemplate: path.join("templates", sourceFile),
    ruleModel: {
      version: "template-rules-v2",
      purpose: "Use the v10 95-item master workbook as template source-of-truth while preserving v9 row-level rules for the first 27 items.",
    },
    rows: outputRows,
  };
  await fs.writeFile(path.join(templateJsonDir, outputName), `${JSON.stringify(output, null, 2)}\n`, "utf8");
  return {
    sector: template.sector ?? "",
    sectorName: template.sectorName ?? "",
    size: template.size ?? "",
    sourceJsonTemplate: path.join("templates", sourceFile),
    jsonTemplate: path.join("templates", outputName),
    rows: output.rows.length,
    operationalHash: hashOperationalRows(output.rows),
  };
}

async function writeDocs(summary) {
  const readme = `# Consultant-safe EBX-Q Template v10

## Purpose
V10 replaces the 27-item qualitative template with the 95-item master workbook \`source/${path.basename(masterWorkbookPath)}\`.

## Operational Profile Model
- Policy source-of-truth: \`policies/template_profiles.json\` (${PROFILE_CONFIG.version}).
- Matrix: 11 sector overlays x 4 operational profiles (대기업, 중견, 중소, 비상장) = 44 templates.
- 비상장 remains a compatibility profile for unlisted-company disclosure and does not imply company size.
- Policy precedence: master row -> v9 row rules -> topic profile -> sector overlay -> size profile -> explicit row policy override.

## What Changed
- Uses the v10 master workbook as the source-of-truth for EBX-Q-001 through EBX-Q-095.
- Preserves v9 row-level rules for the original 27 items.
- Applies 17 topic profiles to the 68 new items and adds sector- and size-specific operational controls.
- Keeps the six-column customer output schema unchanged.
- Rejects duplicate operational profiles during build.

## Validation
- \`node scripts/validate_consultant_safe_v10.mjs\`
- \`node scripts/test_consultant_safe_v10_profiles.mjs\`
- \`node scripts/regress_consultant_safe_v10_large.mjs\`

## Build Stats
- Source v9 JSON templates read: ${summary.sourceTemplates}
- v10 JSON templates generated: ${summary.jsonTemplates}
- Rows per template: ${EXPECTED_ROW_COUNT}
- Fatal issues: ${summary.fatal.length}
- Warnings: ${summary.warnings.length}
`;

  const checklist = `# consultant_safe_v10 Handoff Checklist

- [ ] Use \`policies/template_profiles.json\` as the operational policy source-of-truth and generated JSON files as build artifacts.
- [ ] Rebuild v10 if \`source/${path.basename(masterWorkbookPath)}\` changes.
- [ ] Static validator confirms 44 templates, 95 rows, 4 distinct size hashes per sector, and 11 distinct sector hashes per size.
- [ ] Four profile fixtures select the exact requested size and reject invalid size values.
- [ ] Eleven large-company regressions have zero fail, fatal, and topic-mismatched metric rows.
- [ ] Customer workbook preserves the six established output columns.
- [ ] Customer workbook has ${EXPECTED_ROW_COUNT} EBX rows.
- [ ] Sidecar QA JSON has zero fatal findings before handoff.
- [ ] Rows without company evidence remain marked UNKNOWN and do not invent facts.
`;

  const changelog = `# Version Log

## consultant_safe_v10
- Source release: \`${SOURCE_RELEASE}\`.
- Replaces the 27-item qualitative row set with the 95-item master workbook.
- Keeps EBX-Q-001 through EBX-Q-027 item names and IDs unchanged.
- Adds EBX-Q-028 through EBX-Q-095 from the master workbook.
- Adds ${PROFILE_CONFIG.version} with 17 topic profiles, 11 sector overlays, and 4 operational size/disclosure profiles.
- Requires exact sector and size template selection; invalid explicit sizes no longer fall back to 대기업.
- Keeps v9 safety controls and the six-column customer workbook schema.
`;

  await fs.writeFile(path.join(releaseDir, "README.md"), readme, "utf8");
  await fs.writeFile(path.join(releaseDir, "CHECKLIST.md"), checklist, "utf8");
  await fs.writeFile(path.join(releaseDir, "CHANGELOG.md"), changelog, "utf8");
}

async function main() {
  await ensureDirs();
  const master = await readMasterRows();
  const sourceFiles = (await fs.readdir(sourceTemplateDir))
    .filter((file) => file.endsWith("_consultant_safe_v9.json"))
    .sort();
  const referenceTemplate = JSON.parse(await fs.readFile(path.join(sourceTemplateDir, sourceFiles[0]), "utf8"));
  const validation = validateMaster(master, referenceTemplate.rows ?? []);

  const templates = [];
  for (const file of sourceFiles) {
    templates.push(await buildJsonTemplate(file, master.rows));
  }

  const profileValidation = validateProfileConfig();
  const fatal = [...validation.fatal, ...profileValidation.fatal];
  const warnings = [...validation.warnings, ...profileValidation.warnings];
  if (sourceFiles.length !== EXPECTED_TEMPLATE_COUNT) fatal.push(`Expected ${EXPECTED_TEMPLATE_COUNT} v9 JSON templates, found ${sourceFiles.length}.`);
  if (templates.length !== EXPECTED_TEMPLATE_COUNT) fatal.push(`Expected ${EXPECTED_TEMPLATE_COUNT} v10 JSON templates, generated ${templates.length}.`);
  for (const template of templates) {
    if (template.rows !== EXPECTED_ROW_COUNT) fatal.push(`${template.jsonTemplate}: expected ${EXPECTED_ROW_COUNT} EBX rows, found ${template.rows}.`);
  }
  fatal.push(...validateOperationalVariation(templates));

  const index = {
    generatedAt: new Date().toISOString(),
    releaseName: RELEASE_NAME,
    sourceRelease: SOURCE_RELEASE,
    sourceWorkbook: path.join("source", path.basename(masterWorkbookPath)),
    rowCount: EXPECTED_ROW_COUNT,
    oldRowCount: OLD_ROW_COUNT,
    newRowCount: EXPECTED_ROW_COUNT - OLD_ROW_COUNT,
    profileVersion: PROFILE_CONFIG.version,
    templates,
  };
  await fs.writeFile(path.join(releaseDir, "TEMPLATE_INDEX.json"), `${JSON.stringify(index, null, 2)}\n`, "utf8");

  const qa = {
    generatedAt: new Date().toISOString(),
    releaseName: RELEASE_NAME,
    sourceRelease: SOURCE_RELEASE,
    sourceWorkbook: path.join("source", path.basename(masterWorkbookPath)),
    masterRows: master.rows.length,
    sourceTemplates: sourceFiles.length,
    jsonTemplates: templates.length,
    profileVersion: PROFILE_CONFIG.version,
    sizeProfiles: EXPECTED_SIZES,
    sectorProfiles: EXPECTED_SECTORS,
    operationalVariationValidated: !fatal.some((item) => /operational|profile|size labels|sector overlays/i.test(item)),
    fatal,
    warnings,
  };
  await fs.writeFile(path.join(releaseDir, "QA_REPORT.json"), `${JSON.stringify(qa, null, 2)}\n`, "utf8");
  await writeDocs({ sourceTemplates: sourceFiles.length, jsonTemplates: templates.length, fatal, warnings });

  console.log(JSON.stringify(qa, null, 2));
  if (fatal.length) process.exitCode = 1;
}

await main();
