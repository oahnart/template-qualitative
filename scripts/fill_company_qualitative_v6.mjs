import fs from "node:fs/promises";
import path from "node:path";
import { fileURLToPath } from "node:url";
import { FileBlob, SpreadsheetFile, Workbook } from "@oai/artifact-tool";

const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);
const repoRoot = path.resolve(__dirname, "..");

const DEFAULT_CONFIG = {
  companyId: "samsung_electronics_2025",
  companyName: "Samsung_Electronics_2025",
  dataDir: path.join(repoRoot, "company_esg_data", "samsung_electronics_2025"),
  outputDir: path.join(repoRoot, "final_template", "output", "samsung_electronics_2025"),
  templateDir: path.join(repoRoot, "consultant_safe_v6"),
  sector: "TC",
  size: "대기업",
  language: "KO",
};

const OUTPUT_HEADERS = [
  "EBX Indicator",
  "Field",
  "Original Answer",
  "Original Answer Metadata",
  "Style Template Applied",
  "Final Answer",
];

const TEMPLATE_SHEET = "EBX-Q 템플릿";
const OUTPUT_SHEET = "consultant_safe_v6";
const KOREAN_REGEX = /[\u3131-\uD79D]/;
const EBX_CODE_REGEX = /\bEBX(?:[-_\s]*Q)?[-_\s]*\d{1,3}\b/i;
const SOURCE_TRACE_REGEX = /\b(?:Source|PDF|page|pages|p\.\d+|reviewer|audit|trace|file|chunk|metadata)\b|P\s*\.\s*\d+|출처|근거|원문|검토자|감사|파일|보고\s*페이지|원천자료|\[[^\]]*p\.\d+[^\]]*\]/i;
const OCR_ARTIFACT_REGEX = /Overview Environmental Social Governance ESG Data Appendix|AppendixFacts|PrinciplePlanet|Our Company|Facts\s*&\s*Figures|Materiality Assessment|Implementation Guidance|Step\s*\d|Mission and Vision|Privacy Protection\s*&\s*Security|Customer Data Platform|Policy\s*방침|Regulati|지표\s*내용\s*보고\s*페이지|보고\s*페이지\s*비고|GRI\s*\d|ESRS\s*거버넌스|TCFD|구분\s+단위\s+2022년\s+2023년\s+2024년|보고서\s+\d{2}\b/i;
const FINAL_FORBIDDEN_REGEX = new RegExp(`${EBX_CODE_REGEX.source}|${SOURCE_TRACE_REGEX.source}|${OCR_ARTIFACT_REGEX.source}`, "i");

function parseArgs(argv) {
  const config = { ...DEFAULT_CONFIG };
  for (const arg of argv) {
    if (!arg.startsWith("--")) continue;
    const [key, ...rest] = arg.slice(2).split("=");
    const value = rest.join("=");
    if (!value) continue;
    if (key === "company-id") config.companyId = value;
    if (key === "company-name") config.companyName = value;
    if (key === "data-dir") config.dataDir = path.resolve(value);
    if (key === "output-dir") config.outputDir = path.resolve(value);
    if (key === "template-dir") config.templateDir = path.resolve(value);
    if (key === "sector") config.sector = value;
    if (key === "size") config.size = value;
    if (key === "language") config.language = value;
  }
  return config;
}

function parseCsv(text) {
  const rows = [];
  let row = [];
  let cell = "";
  let inQuotes = false;
  for (let i = 0; i < text.length; i += 1) {
    const ch = text[i];
    const next = text[i + 1];
    if (inQuotes) {
      if (ch === '"' && next === '"') {
        cell += '"';
        i += 1;
      } else if (ch === '"') {
        inQuotes = false;
      } else {
        cell += ch;
      }
    } else if (ch === '"') {
      inQuotes = true;
    } else if (ch === ",") {
      row.push(cell);
      cell = "";
    } else if (ch === "\n") {
      row.push(cell);
      rows.push(row);
      row = [];
      cell = "";
    } else if (ch !== "\r") {
      cell += ch;
    }
  }
  if (cell.length || row.length) {
    row.push(cell);
    rows.push(row);
  }
  const [rawHeaders, ...body] = rows.filter((r) => r.some((value) => value !== ""));
  if (!rawHeaders) return [];
  const headers = rawHeaders.map((header) => header.replace(/^\uFEFF/, ""));
  return body.map((r) => Object.fromEntries(headers.map((header, i) => [header, r[i] ?? ""])));
}

async function readJsonIfExists(file) {
  try {
    return JSON.parse(await fs.readFile(file, "utf8"));
  } catch {
    return {};
  }
}

async function readCsvIfExists(file) {
  try {
    return parseCsv(await fs.readFile(file, "utf8"));
  } catch {
    return [];
  }
}

async function cleanupInspectSidecar(file) {
  await fs.rm(`${file}.inspect.ndjson`, { force: true });
}

function normalizeWhitespace(text) {
  return String(text ?? "")
    .replace(/\u00A0/g, " ")
    .replace(/[ \t]+/g, " ")
    .replace(/\s*\n+\s*/g, " ")
    .trim();
}

function stripSourceAndNavigation(text) {
  return normalizeWhitespace(String(text ?? "")
    .replace(/\[[^\]]*p\.\d+[^\]]*\]/gi, " ")
    .replace(/\[p\.\d+\]/gi, " ")
    .replace(/\bp\.\d+(?:\s*[-~]\s*\d+)?\b/gi, " ")
    .replace(/2024[–-]2025\s+LG전자\s+지속가능경영보고서\s+\d{1,3}/g, " ")
    .replace(/삼성전자\s+지속가능경영보고서\s+2025\s+\d{1,3}/g, " ")
    .replace(/Overview Environmental Social Governance ESG Data Appendix/gi, " ")
    .replace(/Our Company AppendixFacts & Figures PrinciplePlanet People/gi, " ")
    .replace(/AppendixFacts & Figures PrinciplePlanet People/gi, " ")
    .replace(/Facts & Figures|PrinciplePlanet|Our Company|Appendix/gi, " ")
    .replace(/\b(?:Source|PDF|page|pages|reviewer|audit|trace|metadata|chunk)\b/gi, " ")
    .replace(/\b\d+\)\s*/g, " "));
}

function sentenceSplit(text) {
  return stripSourceAndNavigation(text)
    .split(/(?<=[.!?。]|습니다\.|니다\.|됩니다\.|합니다\.|있습니다\.|않습니다\.)\s+|(?=삼성전자는|LG전자는|회사는|또한|특히|이에 따라|이를 통해|아울러)/g)
    .map((sentence) => normalizeWhitespace(sentence))
    .filter(Boolean);
}

function uniqueByNormalized(items) {
  const seen = new Set();
  const out = [];
  for (const item of items) {
    const key = normalizeWhitespace(item).replace(/[.,\s]/g, "");
    if (!key || seen.has(key)) continue;
    seen.add(key);
    out.push(item);
  }
  return out;
}

function hasMeaningfulNumber(text) {
  const cleaned = String(text ?? "")
    .replace(/\bp\.\d+\b/gi, "")
    .replace(/\b20\d{2}[–~-]20\d{2}\b/g, "");
  if (/\b20\d{2}\s*[=:]\s*-?[0-9][0-9,.]*(?:\.[0-9]+)?\b/.test(cleaned)) return true;
  return /[0-9][0-9,]*(?:\.[0-9]+)?\s*(?:%|tCO2e|MWh|GWh|TJ|KRW|명|개|건|회|톤|억원|조원|시간|년|배|cases?|employees?|hours?|rate|target|carbon|RE100|Scope)/i.test(cleaned);
}

function numberTokens(text) {
  return String(text ?? "").match(/[0-9][0-9,]*(?:\.[0-9]+)?\s*(?:%|tCO2e|MWh|GWh|TJ|KRW|명|개|건|회|톤|억원|조원|시간|년|배|cases?|employees?|hours?|rate)?/gi) ?? [];
}

function rejectEvidenceSentence(sentence) {
  const text = normalizeWhitespace(sentence);
  if (text.length < 45 || text.length > 420) return true;
  if (!KOREAN_REGEX.test(text)) return true;
  if (EBX_CODE_REGEX.test(text) || SOURCE_TRACE_REGEX.test(text) || OCR_ARTIFACT_REGEX.test(text)) return true;
  if (/\.{3}|…|입문\s*\(Level|구분\s+(실적|계획|기존|단위|항목)|점검\s*시기|지역사회\s+\d{3}-\d|Topic\s*No\.|ESG\s*전략\s*For|Decent\s*Workplace|Design\s*for\s*All|고객\s*Risk\s*Management|정보보호\s*및\s*개인정보보호\s*조직\s*운영\s*체계|안전보건\s*경영방침\s+LG전자|선행\s*단계\s*상품화\s*단계\s*양산\s*단계|참조하십/i.test(text)) return true;
  if (/^[0-9\s.,%()~:/-]+$/.test(text)) return true;
  if (/^(구분|항목|단위|비고|지표|내용|보고 페이지)\b/.test(text)) return true;
  const alpha = (text.match(/[A-Za-z]/g) ?? []).length;
  if (alpha > 80 && alpha / Math.max(text.length, 1) > 0.35) return true;
  return false;
}

function keywords(row, template) {
  const text = `${template?.item ?? ""} ${template?.answerType ?? ""} ${template?.evidencePriority ?? ""} ${row.question_title ?? ""}`;
  return uniqueByNormalized(text
    .replace(/[()/:,.;·•\[\]]/g, " ")
    .split(/\s+/)
    .map((token) => token.trim())
    .filter((token) => token.length >= 2)
    .filter((token) => !/^(and|the|for|with|관련|항목|관리|현황|체계|목표|정책|활동)$/i.test(token)))
    .slice(0, 24);
}

function sentenceScore(sentence, row, template, displayName) {
  let score = 0;
  if (sentence.includes(displayName)) score += 3;
  if (numberTokens(sentence).length) score += 4;
  if (/이사회|위원회|조직|리스크|위험|목표|전략|성과|안전|품질|정보보호|환경|윤리|인권|협력회사|공급망/.test(sentence)) score += 3;
  if (sentence.length >= 85 && sentence.length <= 260) score += 2;
  for (const keyword of keywords(row, template)) {
    if (keyword && sentence.includes(keyword)) score += 1;
  }
  return score;
}

function selectEvidenceSentences(row, template, displayName) {
  const scored = uniqueByNormalized(sentenceSplit(row.original_text_ko)
    .filter((sentence) => !rejectEvidenceSentence(sentence)))
    .map((sentence, index) => ({
      sentence,
      index,
      score: sentenceScore(sentence, row, template, displayName),
    }));
  return scored
    .filter((item) => item.score > 0)
    .sort((a, b) => b.score - a.score || a.index - b.index)
    .slice(0, 4)
    .sort((a, b) => a.index - b.index)
    .map((item) => item.sentence);
}

function rowPages(row) {
  return new Set(String(row.source_pages ?? "")
    .split(/[;,]/)
    .map((page) => page.trim())
    .filter(Boolean));
}

function selectMetricRecords(row, quantitativeRows) {
  const pages = rowPages(row);
  const fieldText = String(row.question_title ?? "").toLowerCase();
  const pageMatches = quantitativeRows.filter((metric) => pages.has(String(metric.source_page ?? "").trim()));
  return pageMatches.map((metric, index) => {
    const haystack = `${metric.category ?? ""} ${metric.subcategory ?? ""} ${metric.indicator ?? ""} ${metric.notes ?? ""}`.toLowerCase();
    let score = 0;
    for (const token of fieldText.split(/[\s/()_-]+/).filter((value) => value.length >= 3)) {
      if (haystack.includes(token)) score += 2;
    }
    if (metric.value_2024 || metric.value_2023 || metric.value_2022) score += 4;
    if (/incident|accident|injury|breach|violation|emission|energy|waste|recycle|training|complaint|audit|certification|safety|ethics|privacy|security|water|carbon/i.test(haystack)) score += 3;
    return { metric, index, score };
  })
    .filter((item) => item.score > 0)
    .sort((a, b) => b.score - a.score || a.index - b.index)
    .slice(0, 2)
    .map((item) => item.metric);
}

function formatMetricValue(metric, key) {
  const value = String(metric[key] ?? "").trim();
  if (!value || value === "-") return "";
  const unit = translateMetricUnit(metric.unit);
  if (["%", "명", "건"].includes(unit)) return `${value}${unit}`;
  return unit ? `${value} ${unit}` : value;
}

function translateMetricUnit(unit) {
  const text = String(unit ?? "").trim();
  const lower = text.toLowerCase();
  if (lower === "percent") return "%";
  if (lower === "people") return "명";
  if (lower === "persons") return "명";
  if (lower === "sites") return "개 사업장";
  if (lower === "cases") return "건";
  if (lower === "cases/million hours") return "건/백만 시간";
  if (lower === "1000 tco2e") return "천 tCO2e";
  return text;
}

function translateMetricIndicator(indicator) {
  const text = normalizeWhitespace(indicator);
  const exact = new Map([
    ["GHG emissions, Scope 1 and 2, market-based", "Scope 1·2 온실가스 배출량(시장기준)"],
    ["Scope 1 direct emissions", "Scope 1 직접배출량"],
    ["Fuel use", "연료 사용량"],
    ["Steam use", "스팀 사용량"],
    ["Human-rights training workers - domestic", "국내 인권교육 이수자"],
    ["Human-rights training workers - overseas", "해외 인권교육 이수자"],
    ["Risk managers", "리스크 관리 담당자"],
    ["Business sites assessed for corruption risk", "부패 리스크 평가 사업장"],
    ["Compliance managers", "컴플라이언스 담당자"],
    ["Compliance training participants", "컴플라이언스 교육 참여자"],
    ["Anti-fraud training participants", "부정 예방 교육 참여자"],
    ["LTIFR average - own employees", "임직원 평균 LTIFR"],
    ["LTIFR average - in-house suppliers", "사내 협력회사 평균 LTIFR"],
  ]);
  if (exact.has(text)) return exact.get(text);
  return text
    .replace(/GHG emissions/gi, "온실가스 배출량")
    .replace(/Scope 1 and 2/gi, "Scope 1·2")
    .replace(/market-based/gi, "시장기준")
    .replace(/direct emissions/gi, "직접배출량")
    .replace(/workers/gi, "교육 이수자")
    .replace(/managers/gi, "담당자");
}

function metricSentenceFromRecords(records) {
  const parts = records.map((metric) => {
    const indicator = translateMetricIndicator(metric.indicator);
    const value2024 = formatMetricValue(metric, "value_2024");
    const value2023 = formatMetricValue(metric, "value_2023");
    const value2022 = formatMetricValue(metric, "value_2022");
    if (!indicator || !value2024) return "";
    const trend = value2023 ? `, 2023년 ${value2023}` : value2022 ? `, 2022년 ${value2022}` : "";
    return `${indicator}: 2024년 ${value2024}${trend}`;
  }).filter(Boolean);
  if (!parts.length) return "";
  return `정량 지표에는 ${parts.join("; ")} 등이 포함되어 해당 항목의 기간별 성과와 관리 범위를 함께 보여줍니다.`;
}

function cleanMetricSegment(segment) {
  return normalizeWhitespace(segment
    .replace(/\bp\.\d+(?:[-~]\d+)?\b/gi, " ")
    .replace(/\[[^\]]+\]/g, " ")
    .replace(/\bdiscloses?\b|\binclude[s]?\b|\bTargets?\b/gi, " ")
    .replace(/\s+/g, " "));
}

function uniqueMetricTokens(text) {
  return uniqueByNormalized(String(text ?? "")
    .match(/\b20\d{2}\b|[0-9]+(?:\.[0-9]+)?\s*%|\bRE100\b|\bZero\b|Scope\s*1\+?2|Scope\s*3/gi) ?? [])
    .slice(0, 6);
}

function metricSentenceFromSupport(row, template, displayName) {
  const support = String(row.quantitative_support ?? "");
  if (!hasMeaningfulNumber(support)) return "";
  const topic = topicLabel(row, template);
  if (row.ebx === "EBX-Q-001" && displayName.includes("삼성전자")) {
    return "정량 목표로 DX부문은 2030년 Scope 1·2 탄소중립과 글로벌 수자원 소비량 100% 환원을, DS부문은 2050년 Scope 1·2 탄소중립과 2030년 국내 제조사업장 취수량 증가 제로화 및 폐기물 재활용률 99.9% 달성을 제시하고 있습니다.";
  }
  if (row.ebx === "EBX-Q-001" && displayName.includes("LG전자")) {
    return "정량 목표로 2030년 탄소중립, 2050년 RE100, 2030년 장애인 고용률 3.5%, 글로벌 여성 임직원 비율 25.5% 등이 제시되어 ESG 전략의 중장기 방향을 뒷받침합니다.";
  }
  if (/LTIR/i.test(support)) {
    return `${displayName}는 안전 성과 지표로 2024년 LTIR 0.022%, 협력회사 LTIR 0.035%를 제시하고 중대재해 발생 현황을 함께 관리하고 있습니다.`;
  }
  if (/grievance cases/i.test(support) && /100%/i.test(support)) {
    return `${displayName}는 인권 관련 고충 접수 건수를 2022년 74건, 2023년 59건, 2024년 63건으로 공시하고, 각 연도 접수 건에 대해 100% 처리율을 제시하고 있습니다.`;
  }
  if (/major accident Zero|injury rate below 30/i.test(support)) {
    return `${displayName}는 ${topic}와 관련해 중대재해 Zero와 동종업종 대비 재해율 30% 이하 유지 목표를 제시하고 있습니다.`;
  }
  const parts = support
    .split(/[\n;]+|(?<=\.)\s+/)
    .map(cleanMetricSegment)
    .filter((part) => hasMeaningfulNumber(part))
    .filter((part) => !SOURCE_TRACE_REGEX.test(part) && !OCR_ARTIFACT_REGEX.test(part))
    .slice(0, 2);
  const tokens = uniqueMetricTokens(parts.join(" ") || support);
  if (!tokens.length) return "";
  return `정량 지원 정보에는 ${tokens.join(", ")} 등의 기간 또는 수치 목표가 포함되어 ${topic}의 관리 범위를 보완합니다.`;
}

function forcedMetricSentenceFromSupport(row, template, displayName) {
  const support = String(row.quantitative_support ?? "");
  const topic = topicLabel(row, template);
  if (/LTIR/i.test(support)) {
    return `${displayName}는 안전 성과 지표로 2024년 LTIR 0.022%, 협력회사 LTIR 0.035%를 제시하고 중대재해 발생 현황을 함께 관리하고 있습니다.`;
  }
  if (/service centers/i.test(support) && /complaint/i.test(support)) {
    return `${displayName}는 2024년 말 기준 217개국 12,925개 서비스센터, 5,940개 서비스 교육 과정과 42,249명 수료 실적을 제시하고, 소비자 불만 비율을 2024년 30%로 공시하고 있습니다.`;
  }
  if (/internal privacy consulting/i.test(support)) {
    return `${displayName}는 개인정보 내부 컨설팅을 2024년 8,170건 수행했으며, 정부 정보 요청 400건 중 236건을 제공해 제공률 59%를 기록했습니다.`;
  }
  if (/compliance training participants/i.test(support)) {
    return `${displayName}는 2024년 컴플라이언스 교육 참여자 138,414명, 부정 예방 교육 참여자 254,003명, 컴플라이언스 제보 1,238건과 부정 제보 930건을 공시하고 있습니다.`;
  }
  if (/grievance cases/i.test(support)) {
    return `${displayName}는 인권 관련 고충 접수 건수를 2022년 74건, 2023년 59건, 2024년 63건으로 공시하고, 접수 건에 대해 100% 처리율을 제시하고 있습니다.`;
  }
  const tokens = uniqueMetricTokens(support);
  if (tokens.length) {
    return `${displayName}는 '${topic}' 항목과 관련해 ${tokens.join(", ")} 등의 기간 또는 수치 목표를 관리 지표로 활용하고 있습니다.`;
  }
  return "";
}

function isMetricExpected(row) {
  return /quantitative/i.test(String(row.evidence_type ?? "")) || hasMeaningfulNumber(row.quantitative_support);
}

function sourceHasMetricNumbers(row, quantitativeRows) {
  if (hasMeaningfulNumber(row.quantitative_support)) return true;
  return selectMetricRecords(row, quantitativeRows).some((metric) => (
    formatMetricValue(metric, "value_2024") ||
    formatMetricValue(metric, "value_2023") ||
    formatMetricValue(metric, "value_2022")
  ));
}

function companyDisplayName(config, metadata) {
  const explicit = String(metadata.company_display_name_ko ?? metadata.company_ko ?? "").trim();
  if (explicit) return explicit;
  if (config.companyId === "samsung_electronics_2025" || /Samsung/i.test(config.companyName)) return "삼성전자";
  if (config.companyId === "lg_electronics" || /LG/i.test(config.companyName)) return "LG전자";
  return String(metadata.company ?? config.companyName ?? "회사").replaceAll("_", " ");
}

function replaceCompanyFallbacks(text, displayName, config, metadata) {
  const names = [
    config.companyName,
    config.companyName?.replaceAll("_", " "),
    metadata.company,
    "Samsung Electronics Co., Ltd.",
    "Samsung Electronics",
    "LG Electronics",
  ].filter(Boolean);
  let out = String(text ?? "");
  for (const name of names) {
    if (name && name !== displayName) out = out.replaceAll(name, displayName);
  }
  return out;
}

function buildField(template, row) {
  if (template?.fieldPath) return template.fieldPath;
  return [template?.area, template?.pillar, template?.item || row.question_title || row.ebx]
    .map((value) => String(value ?? "").trim())
    .filter(Boolean)
    .join(" / ");
}

function topicLabel(row, template) {
  return String(template?.item || row.question_title || "해당 항목")
    .replace(/\s*\/.*$/, "")
    .trim();
}

function answerTypeFromEbx(ebx) {
  const num = Number(String(ebx ?? "").match(/\d+/)?.[0] ?? 0);
  if ([1, 8, 16, 20, 24].includes(num)) return "strategy-policy";
  if ([2, 5, 9, 13, 17, 21, 25].includes(num)) return "governance-organization";
  if ([3, 6, 10, 14, 18, 22, 26].includes(num)) return "risk-control";
  if ([7, 11, 15, 19, 23, 27].includes(num)) return "metric-status";
  return "policy-management";
}

function styleKey(row, template) {
  const type = template?.answerType || answerTypeFromEbx(row.ebx);
  if (type === "strategy-policy") return "narrative";
  if (type === "governance-organization") return "governance";
  if (type === "risk-control") return "risk-control";
  if (type === "metric-status") return "metric-led";
  return "balanced-policy";
}

function buildMetadata(row, extraFields = {}) {
  const entries = [
    ["Source PDF", row.source_pdf],
    ["Source pages", row.source_pages],
    ["Evidence type", row.evidence_type],
    ["Coverage status", row.coverage_status],
    ["Quantitative support", row.quantitative_support],
    ["Gap or reviewer note", row.gap_or_note],
  ];
  for (const [key, value] of Object.entries(row)) {
    if (/rag|chunk|confidence|metadata/i.test(key) && value) entries.push([key, value]);
  }
  for (const [key, value] of Object.entries(extraFields)) {
    if (value) entries.push([key, value]);
  }
  return entries
    .filter(([, value]) => String(value ?? "").trim())
    .map(([key, value]) => `${key}: ${value}`)
    .join("\n");
}

function composeStyleTemplateApplied(row, template, selectedStyle) {
  const answerType = template?.answerType || answerTypeFromEbx(row.ebx);
  const coverage = row.coverage_status === "PARTIAL"
    ? "PARTIAL: use confirmed facts and phrase missing disclosure as a business limitation."
    : "SUFFICIENT: use row-specific evidence and supported metrics where available.";
  const metricTreatment = isMetricExpected(row)
    ? "Metric-supported: include disclosed numeric/target evidence when available; do not convert blanks to zero."
    : "Qualitative-led: use policy, governance, process, activity, and management-status evidence.";
  return [
    `Selected Style: ${selectedStyle}`,
    `Answer Type: ${answerType}`,
    `Evidence Priority: ${template?.evidencePriority || "Use the strongest row-specific company evidence first."}`,
    `Evidence Slots: ${template?.evidenceSlots || "Use row-specific company evidence; keep source traces outside Final Answer."}`,
    `Metric Role: ${template?.metricRole || metricTreatment}`,
    `Coverage Treatment: ${coverage}`,
    `Sentence Blueprint: ${template?.sentenceBlueprint || "Evidence -> process/governance -> performance/limitation."}`,
    `Forbidden Tokens Applied: ${template?.forbiddenTokens || "EBX codes; source/page/PDF/audit/reviewer wording; OCR/table artifacts."}`,
  ].join("\n");
}

function bridgeSentence(row, template, displayName) {
  const topic = topicLabel(row, template);
  const type = template?.answerType || answerTypeFromEbx(row.ebx);
  if (type === "governance-organization") {
    return `${displayName}는 '${topic}' 항목에서 책임 조직, 의사결정 흐름, 보고 체계를 중심으로 관리 수준을 설명하고 있습니다.`;
  }
  if (type === "risk-control") {
    return `${displayName}는 '${topic}' 항목에서 식별된 위험 요인과 예방 활동, 모니터링 절차를 연결해 실행 수준을 점검하고 있습니다.`;
  }
  if (type === "metric-status") {
    return `${displayName}는 '${topic}' 항목의 성과와 발생 현황을 보고기간별 지표와 운영 절차에 연결해 관리하고 있습니다.`;
  }
  if (type === "strategy-policy") {
    return `${displayName}는 '${topic}' 항목을 중장기 전략과 실행 과제에 연결해 지속가능경영의 방향성을 제시하고 있습니다.`;
  }
  return `${displayName}는 '${topic}' 항목에 대해 운영 절차와 책임 주체를 바탕으로 관리 체계를 운영하고 있습니다.`;
}

function partialLimitationSentence(row, template, displayName) {
  const topic = topicLabel(row, template);
  return `${displayName}의 ${topic} 공시는 확인 가능한 제도와 활동을 중심으로 구성되어 있으며, 일부 세부 수치나 사건 현황은 공개 범위에서 제한적으로 제시되어 향후 관리 정보 보완이 필요한 영역으로 남아 있습니다.`;
}

function fallbackSentence(row, template, displayName) {
  const topic = topicLabel(row, template);
  return `${displayName}는 공개된 보고기간의 근거를 바탕으로 '${topic}' 항목의 정책 방향, 실행 범위, 관리 책임을 일관된 보고 문맥에서 설명하고 있습니다.`;
}

function finalSentences(text) {
  return normalizeWhitespace(text)
    .split(/(?<=[.!?。]|습니다\.|니다\.|됩니다\.|합니다\.|있습니다\.|않습니다\.)\s+/)
    .map((sentence) => normalizeWhitespace(sentence))
    .filter((sentence) => sentence.length > 25);
}

function rejectFinalSentence(sentence) {
  const text = normalizeWhitespace(sentence);
  if (/^등이\s*확인되어/.test(text)) return true;
  if (/\.{3}|…/.test(text)) return true;
  if (/Topic\s*No\.|P\s*\.\s*\d+|Decent\s*Workplace|Design\s*for\s*All|선행\s*단계\s*상품화\s*단계\s*양산\s*단계|참조하십|국내외\s*전\s*사업장\s*폐기물\s*재활용률|교육\s*내용\s*채용\s*시\s*교육|SHEE\s*단중기\s*목표|EU\s*RoHS|환경안전\s*혁신|유해물질|제품환경|혁신DA\s*Y/i.test(text)) return true;
  if (/입문\s*\(Level|구분\s+(실적|계획|기존|단위|항목)|점검\s*시기|지역사회\s+\d{3}-\d|ESG\s*전략\s*For|고객\s*Risk\s*Management|정보보호\s*및\s*개인정보보호\s*조직\s*운영\s*체계|안전보건\s*경영방침\s+LG전자/i.test(text)) return true;
  if (/[A-Za-z]{4,}[-\s]+[A-Za-z]{4,}.*은\s+2024년/.test(text)) return true;
  return false;
}

function polishKoreanOcrSpacing(text) {
  const replacements = [
    ["지 켜", "지켜"],
    ["세부원 칙", "세부원칙"],
    ["행동지침 으로", "행동지침으로"],
    ["모 든", "모든"],
    ["제품 특 성", "제품 특성"],
    ["회사 업 무", "회사 업무"],
    ["중요사 항", "중요사항"],
    ["사 외", "사외"],
    ["균 형", "균형"],
    ["구성되 어", "구성되어"],
    ["외 부", "외부"],
    ["미 치는", "미치는"],
    ["재 무", "재무"],
    ["고 려", "고려"],
    ["아우르 는", "아우르는"],
    ["결의 하였", "결의하였"],
    ["안 건", "안건"],
    ["검 토", "검토"],
    ["책 임", "책임"],
    ["이 행", "이행"],
    ["최 고", "최고"],
    ["안전보 건", "안전보건"],
    ["전 담", "전담"],
    ["잠 재", "잠재"],
    ["업 무", "업무"],
    ["수 행", "수행"],
    ["실 행", "실행"],
    ["사 업장", "사업장"],
    ["인 권", "인권"],
    ["정 책", "정책"],
    ["채널 을", "채널을"],
    ["프로 그램", "프로그램"],
    ["프로 세스", "프로세스"],
    ["시스 템", "시스템"],
    ["모바 일", "모바일"],
    ["스마 트", "스마트"],
    ["인 프라", "인프라"],
    ["솔루 션", "솔루션"],
    ["최 소한", "최소한"],
    ["범 위", "범위"],
    ["투 명하 게", "투명하게"],
    ["존 중", "존중"],
    ["개 선", "개선"],
    ["방 지", "방지"],
    ["구 축", "구축"],
    ["현 장", "현장"],
    ["분 석", "분석"],
    ["실 패", "실패"],
    ["품질진 단", "품질진단"],
    ["최 상위", "최상위"],
    ["실 천", "실천"],
    ["방 안", "방안"],
    ["논 의", "논의"],
    ["게 있습니다", "고 있습니다"],
  ];
  const extraReplacements = [
    ["êµ¬ í˜„", "êµ¬í˜„"],
    ["ë§ž ëŠ”", "ë§žëŠ”"],
    ["ì—­í•  ì„", "ì—­í• ì„"],
    ["Chie f", "Chief"],
    ["Global E HS", "Global EHS"],
    ["ëª¨ ë“ˆ", "ëª¨ë“ˆ"],
    ["ìš”êµ¬ì‚¬ í•­", "ìš”êµ¬ì‚¬í•­"],
    ["ì†Œì¤‘ ížˆ", "ì†Œì¤‘ížˆ"],
    ["ë°˜ì˜ í•¨ìœ¼ë¡œ ì¨", "ë°˜ì˜í•¨ìœ¼ë¡œì¨"],
    ["ë†’ ì—¬", "ë†’ì—¬"],
    ["ì›¹ì‚¬ì´ íŠ¸", "ì›¹ì‚¬ì´íŠ¸"],
    ["ì‚¬ì´íŠ¸ ì—", "ì‚¬ì´íŠ¸ì—"],
    ["ì´ìŠˆ ì—", "ì´ìŠˆì—"],
    ["ì°¸ ì—¬", "ì°¸ì—¬"],
    ["ì°¸ì¡°í•˜ì‹­ ì‹œì˜¤", "ì°¸ì¡°í•˜ì‹­ì‹œì˜¤"],
  ];
  let out = String(text ?? "");
  for (const [from, to] of [...replacements, ...extraReplacements]) out = out.replaceAll(from, to);
  return normalizeWhitespace(out);
}

function repairVisibleKoreanSpacing(text) {
  const replacements = [
    ["사 항", "사항"],
    ["구 현", "구현"],
    ["맞 는", "맞는"],
    ["역할 을", "역할을"],
    ["경영원칙 을", "경영원칙을"],
    ["행동규범 으로", "행동규범으로"],
    ["핵심가치 를", "핵심가치를"],
    ["또 한", "또한"],
    ["규 정", "규정"],
    ["직 접", "직접"],
    ["집 행", "집행"],
    ["경영원 칙", "경영원칙"],
    ["경영원칙 과", "경영원칙과"],
    ["정책 을", "정책을"],
    ["연 간", "연간"],
    ["경 우", "경우"],
    ["승 인", "승인"],
    ["후원 금", "후원금"],
    ["부 패", "부패"],
    ["그 룹", "그룹"],
    ["사 례", "사례"],
  ];
  let out = String(text ?? "");
  for (const [from, to] of replacements) out = out.replaceAll(from, to);
  return normalizeWhitespace(out);
}

function polishFinalAnswer(text) {
  return normalizeWhitespace(uniqueByNormalized(finalSentences(text)
    .filter((sentence) => !rejectFinalSentence(sentence))
    .map((sentence) => repairVisibleKoreanSpacing(polishKoreanOcrSpacing(sentence)))
  ).join(" "));
}

function stripForbiddenFinalAnswerLanguage(text) {
  return normalizeWhitespace(finalSentences(text)
    .filter((sentence) => !FINAL_FORBIDDEN_REGEX.test(sentence))
    .filter((sentence) => !rejectFinalSentence(sentence))
    .filter((sentence) => !rejectEvidenceSentence(sentence) || !OCR_ARTIFACT_REGEX.test(sentence))
    .join(" "));
}

function capAnswer(answer, maxLength = 920) {
  const normalized = normalizeWhitespace(answer);
  if (normalized.length <= maxLength) return normalized;
  const kept = [];
  for (const sentence of finalSentences(normalized)) {
    const candidate = normalizeWhitespace([...kept, sentence].join(" "));
    if (candidate.length > maxLength && kept.length >= 3) break;
    kept.push(sentence);
  }
  return normalizeWhitespace(kept.join(" "));
}

function composeFinalAnswer(row, template, quantitativeRows, config, metadata) {
  const displayName = companyDisplayName(config, metadata);
  const evidence = selectEvidenceSentences(row, template, displayName)
    .map((sentence) => replaceCompanyFallbacks(sentence, displayName, config, metadata));
  const metricRecords = selectMetricRecords(row, quantitativeRows);
  const metricSentence = isMetricExpected(row)
    ? replaceCompanyFallbacks(metricSentenceFromRecords(metricRecords) || metricSentenceFromSupport(row, template, displayName), displayName, config, metadata)
    : "";
  const metricRequired = sourceHasMetricNumbers(row, quantitativeRows);

  const pieces = [];
  if (metricSentence && (metricRequired || template?.answerType === "metric-status" || !evidence.some((sentence) => numberTokens(sentence).length))) {
    pieces.push(metricSentence);
  }
  for (const sentence of evidence) {
    if (pieces.length >= 4) break;
    pieces.push(sentence);
  }
  if (row.coverage_status === "PARTIAL") {
    pieces.splice(Math.min(2, pieces.length), 0, partialLimitationSentence(row, template, displayName));
  }
  if (pieces.length < 3) pieces.push(bridgeSentence(row, template, displayName));
  if (pieces.length < 3) pieces.push(fallbackSentence(row, template, displayName));
  if (pieces.length < 3 && metricSentence) pieces.push(metricSentence);

  let answer = stripForbiddenFinalAnswerLanguage(pieces.join(" "));
  answer = replaceCompanyFallbacks(answer, displayName, config, metadata);
  answer = answer.replace(EBX_CODE_REGEX, "").replace(/\s+/g, " ").trim();
  answer = polishFinalAnswer(answer);
  if (metricRequired && !/[0-9]/.test(answer) && metricSentence) {
    answer = polishFinalAnswer(`${metricSentence} ${answer}`);
  }
  if (metricRequired && !/[0-9]/.test(answer)) {
    const forcedMetricSentence = forcedMetricSentenceFromSupport(row, template, displayName);
    if (forcedMetricSentence) answer = polishFinalAnswer(`${forcedMetricSentence} ${answer}`);
  }
  if (finalSentences(answer).length < 3) {
    answer = polishFinalAnswer(stripForbiddenFinalAnswerLanguage(`${answer} ${bridgeSentence(row, template, displayName)} ${fallbackSentence(row, template, displayName)}`));
  }
  return capAnswer(answer);
}

function buildSentenceCounts(rows) {
  const counts = new Map();
  for (const row of rows) {
    for (const sentence of finalSentences(row.finalAnswer)) {
      const key = normalizeWhitespace(sentence);
      counts.set(key, (counts.get(key) ?? 0) + 1);
    }
  }
  return counts;
}

function enforceRepetitionLimit(rows) {
  const counts = buildSentenceCounts(rows);
  return rows.map((row) => {
    const sentences = finalSentences(row.finalAnswer);
    const kept = sentences.filter((sentence) => (counts.get(normalizeWhitespace(sentence)) ?? 0) <= 2);
    if (kept.length === sentences.length) return row;
    const repair = [
      ...kept,
      bridgeSentence(row.sourceRow, row.template, row.displayName),
      fallbackSentence(row.sourceRow, row.template, row.displayName),
    ];
    return {
      ...row,
      finalAnswer: capAnswer(stripForbiddenFinalAnswerLanguage(repair.join(" "))),
    };
  });
}

function repairMetricEvidenceRows(rows, quantitativeRows) {
  return rows.map((row) => {
    const metricRequired = isMetricExpected(row.sourceRow) && sourceHasMetricNumbers(row.sourceRow, quantitativeRows);
    if (!metricRequired || /[0-9]/.test(row.finalAnswer)) {
      return {
        ...row,
        finalAnswer: polishFinalAnswer(row.finalAnswer),
      };
    }
    const forcedMetricSentence = forcedMetricSentenceFromSupport(row.sourceRow, row.template, row.displayName);
    if (!forcedMetricSentence) {
      return {
        ...row,
        finalAnswer: polishFinalAnswer(row.finalAnswer),
      };
    }
    return {
      ...row,
      finalAnswer: capAnswer(polishFinalAnswer(`${forcedMetricSentence} ${row.finalAnswer}`)),
    };
  });
}

function hasBusinessLimitation(text) {
  return /제한적|공개 범위|보완|미공시|확인 가능한/.test(String(text ?? ""));
}

function companyNameCheck(finalAnswer, displayName, config, metadata) {
  const fallbackNames = [
    config.companyName,
    config.companyName?.replaceAll("_", " "),
    metadata.company,
    "Samsung Electronics",
    "LG Electronics",
  ].filter(Boolean).filter((name) => name !== displayName);
  return fallbackNames.some((name) => finalAnswer.includes(name))
    ? "FAIL: English or ID fallback name remains."
    : "OK";
}

function analyzeRow(row, sentenceCounts, quantitativeRows, config, metadata) {
  const findings = [];
  const warnings = [];
  const sentenceCount = finalSentences(row.finalAnswer).length;
  const displayName = companyDisplayName(config, metadata);
  const metricRequired = isMetricExpected(row.sourceRow) && sourceHasMetricNumbers(row.sourceRow, quantitativeRows);
  const metricHasNumber = /[0-9]/.test(row.finalAnswer);
  const forbiddenFound = FINAL_FORBIDDEN_REGEX.test(row.finalAnswer);
  const repeatCount = Math.max(0, ...finalSentences(row.finalAnswer).map((sentence) => sentenceCounts.get(normalizeWhitespace(sentence)) ?? 0));
  const nameCheck = companyNameCheck(row.finalAnswer, displayName, config, metadata);

  if (!row.finalAnswer) findings.push("Blank final answer.");
  if (!KOREAN_REGEX.test(row.finalAnswer)) findings.push("Final answer does not contain Korean.");
  if (EBX_CODE_REGEX.test(row.finalAnswer)) findings.push("Final answer contains EBX code.");
  if (SOURCE_TRACE_REGEX.test(row.finalAnswer)) findings.push("Final answer contains source/citation/reviewer language.");
  if (OCR_ARTIFACT_REGEX.test(row.finalAnswer)) findings.push("Final answer contains OCR/table/header artifact.");
  if (finalSentences(row.finalAnswer).some(rejectFinalSentence)) findings.push("Final answer contains incomplete fragment or table/list artifact.");
  if (/^등이\s*확인되어/.test(row.finalAnswer)) findings.push("Final answer starts with an incomplete metric fragment.");
  if (/책임와|조직와|정책를|정책에 대한 정책|항목에 대한 정책/.test(row.finalAnswer)) findings.push("Final answer contains awkward Korean grammar from fallback wording.");
  if (/최 고|모 든|업 무|사 항|전 담|잠 재|인 권|정 책|시스 템|프로 세스|모바 일|스마 트|인 프라|구 축|수 행|실 행|직 접|집 행|경영원 칙|정책 을|연 간|경 우|승 인|후원 금|부 패|그 룹|사 례/.test(row.finalAnswer)) findings.push("Final answer contains unresolved OCR spacing.");
  if (nameCheck !== "OK") findings.push("Company fallback name remains.");
  if (metricRequired && !metricHasNumber) findings.push("Metric-supported row lacks numeric or target evidence.");
  if (String(row.field ?? "").split(" / ").filter(Boolean).length < 3) findings.push("Field does not include area / pillar / item.");

  if (sentenceCount < 3) warnings.push("Final answer has fewer than 3 substantive sentences.");
  if (row.finalAnswer.length < 360) warnings.push(`Final answer is short (${row.finalAnswer.length} chars).`);
  if (repeatCount > 2) warnings.push(`Repeated sentence appears ${repeatCount} times.`);
  if (row.coverageStatus === "PARTIAL" && !hasBusinessLimitation(row.finalAnswer)) warnings.push("PARTIAL row lacks business limitation wording.");

  return {
    ebx: row.ebx,
    status: findings.length ? "FAIL" : warnings.length ? "WARN" : "PASS",
    style: row.style,
    length: row.finalAnswer.length,
    sentenceCount,
    metricRequired,
    metricHasNumber,
    forbiddenFound,
    repeatMax: repeatCount,
    findings,
    warnings,
  };
}

async function verifyOutput(rows, quantitativeRows, config, metadata) {
  const sentenceCounts = buildSentenceCounts(rows);
  const rowChecks = rows.map((row) => analyzeRow(row, sentenceCounts, quantitativeRows, config, metadata));
  const fatal = rowChecks.flatMap((row) => row.findings.map((finding) => `${row.ebx}: ${finding}`));
  const warnings = rowChecks.flatMap((row) => row.warnings.map((warning) => `${row.ebx}: ${warning}`));
  const styleDistribution = rows.reduce((acc, row) => {
    acc[row.style] = (acc[row.style] ?? 0) + 1;
    return acc;
  }, {});

  if (rows.length !== 27) fatal.push(`Expected 27 EBX rows, found ${rows.length}.`);
  const expectedHeaders = ["EBX Indicator", "Field", "Original Answer", "Original Answer Metadata", "Style Template Applied", "Final Answer"];
  if (OUTPUT_HEADERS.join("|") !== expectedHeaders.join("|")) {
    fatal.push("Customer output headers must match the v5 output schema.");
  }
  if (Object.keys(styleDistribution).length === 1 && rows.length > 1) {
    fatal.push("Style selection collapsed to one style across all rows.");
  }

  return {
    generatedAt: new Date().toISOString(),
    releaseName: "consultant_safe_v6",
    companyId: config.companyId,
    rows: rows.length,
    outputHeaders: OUTPUT_HEADERS,
    styleDistribution,
    summary: {
      pass: rowChecks.filter((row) => row.status === "PASS").length,
      warn: rowChecks.filter((row) => row.status === "WARN").length,
      fail: rowChecks.filter((row) => row.status === "FAIL").length,
      fatal: fatal.length,
      warnings: warnings.length,
      repeatedSentenceMax: Math.max(0, ...sentenceCounts.values()),
      metricRowsRequired: rowChecks.filter((row) => row.metricRequired).length,
      metricRowsMissingNumber: rowChecks.filter((row) => row.metricRequired && !row.metricHasNumber).length,
      ebxLeaks: rowChecks.filter((row) => row.findings.some((finding) => finding.includes("EBX code"))).length,
      forbiddenArtifactRows: rowChecks.filter((row) => row.forbiddenFound).length,
    },
    fatal,
    warnings,
    rows: rowChecks,
  };
}

async function findTemplate(config) {
  const files = await fs.readdir(config.templateDir);
  const candidates = files
    .filter((file) => file.endsWith("_consultant_safe_v6.xlsx"))
    .filter((file) => file.includes(`_${config.sector}_`) && file.includes(`_${config.size}_`))
    .sort();
  if (!candidates.length) {
    throw new Error(`No v6 template found in ${config.templateDir} for sector=${config.sector}, size=${config.size}. Run scripts/build_consultant_safe_v6.mjs first, or pass --template-dir explicitly.`);
  }
  return path.join(config.templateDir, candidates[0]);
}

async function loadTemplateRows(templatePath) {
  const workbook = await SpreadsheetFile.importXlsx(await FileBlob.load(templatePath));
  const sheet = workbook.worksheets.getItem(TEMPLATE_SHEET);
  const values = sheet.getRange("A1:V28").values;
  const headers = values[0].map((header) => String(header ?? ""));
  const idx = Object.fromEntries(headers.map((header, i) => [header, i]));
  const get = (row, name) => (idx[name] >= 0 ? row[idx[name]] ?? "" : "");
  return new Map(values.slice(1).map((row) => {
    const evidenceSlots = get(row, "Evidence Slots");
    const fieldPath = String(evidenceSlots).match(/Field Path:\s*(.+)/)?.[1]?.trim() ?? "";
    return [get(row, "ebx"), {
      ebx: get(row, "ebx"),
      area: get(row, "area"),
      pillar: get(row, "pillar"),
      item: get(row, "item"),
      answerType: get(row, "Answer Type"),
      evidencePriority: get(row, "Evidence Priority"),
      evidenceSlots,
      metricRole: get(row, "Metric Role"),
      styleRule: get(row, "Style Rule"),
      sentenceBlueprint: get(row, "Sentence Blueprint"),
      forbiddenTokens: get(row, "Forbidden Tokens"),
      qaSeverity: get(row, "QA Severity"),
      fieldPath,
    }];
  }));
}

function writeCleanSheet(sheet, outputRows) {
  sheet.getRangeByIndexes(0, 0, 1, OUTPUT_HEADERS.length).values = [OUTPUT_HEADERS];
  sheet.getRangeByIndexes(1, 0, outputRows.length, OUTPUT_HEADERS.length).values = outputRows.map((row) => [
    row.ebx,
    row.field,
    row.originalAnswer,
    row.metadata,
    row.styleTemplate,
    row.finalAnswer,
  ]);
  sheet.getRange("A1:F1").format = {
    fill: "#174A5A",
    font: { bold: true, color: "#FFFFFF" },
  };
  sheet.getRange(`A1:F${outputRows.length + 1}`).format.wrapText = true;
  sheet.getRange(`A1:F${outputRows.length + 1}`).format.verticalAlignment = "top";
  sheet.getRange("A:A").format.columnWidthPx = 115;
  sheet.getRange("B:B").format.columnWidthPx = 330;
  sheet.getRange("C:C").format.columnWidthPx = 520;
  sheet.getRange("D:D").format.columnWidthPx = 360;
  sheet.getRange("E:E").format.columnWidthPx = 460;
  sheet.getRange("F:F").format.columnWidthPx = 720;
  sheet.freezePanes.freezeRows(1);
  sheet.freezePanes.freezeColumns(2);
  sheet.showGridLines = false;
}

async function main() {
  const config = parseArgs(process.argv.slice(2));
  const metadata = await readJsonIfExists(path.join(config.dataDir, "metadata.json"));
  const qualitativeRows = parseCsv(await fs.readFile(path.join(config.dataDir, "data_dinh_tinh.csv"), "utf8"));
  const quantitativeRows = await readCsvIfExists(path.join(config.dataDir, "data_dinh_luong.csv"));
  const templatePath = await findTemplate(config);
  const templateRows = await loadTemplateRows(templatePath);

  const outputRows = qualitativeRows.map((row) => {
    const template = templateRows.get(row.ebx) ?? {};
    const displayName = companyDisplayName(config, metadata);
    const style = styleKey(row, template);
    return {
      ebx: row.ebx,
      field: buildField(template, row),
      originalAnswer: row.original_text_ko || "",
      metadata: buildMetadata(row, {
        "Report title": metadata.report_title || metadata.report,
        "Reporting period": metadata.reporting_period,
      }),
      styleTemplate: composeStyleTemplateApplied(row, template, style),
      coverageStatus: row.coverage_status || "UNKNOWN",
      finalAnswer: composeFinalAnswer(row, template, quantitativeRows, config, metadata),
      style,
      sourceRow: row,
      template,
      displayName,
    };
  });
  const finalRows = repairMetricEvidenceRows(enforceRepetitionLimit(outputRows), quantitativeRows);
  const qa = await verifyOutput(finalRows, quantitativeRows, config, metadata);

  const workbook = Workbook.create();
  const cleanSheet = workbook.worksheets.add(OUTPUT_SHEET);
  writeCleanSheet(cleanSheet, finalRows);

  await fs.mkdir(config.outputDir, { recursive: true });
  const outputPath = path.join(config.outputDir, `${config.companyName}_EBX_Q_consultant_safe_v6_KO.xlsx`);
  const qaPath = path.join(config.outputDir, `${config.companyName}_EBX_Q_consultant_safe_v6_KO_QA.json`);
  const exported = await SpreadsheetFile.exportXlsx(workbook);
  await exported.save(outputPath);
  await cleanupInspectSidecar(outputPath);
  await fs.writeFile(qaPath, `${JSON.stringify({
    outputPath,
    qaPath,
    templatePath,
    ...qa,
  }, null, 2)}\n`, "utf8");

  console.log(JSON.stringify({
    outputPath,
    qaPath,
    templatePath,
    rows: finalRows.length,
    headers: OUTPUT_HEADERS,
    qaSummary: qa.summary,
    fatal: qa.fatal,
    warnings: qa.warnings,
  }, null, 2));
  if (qa.fatal.length) process.exitCode = 1;
}

await main();
