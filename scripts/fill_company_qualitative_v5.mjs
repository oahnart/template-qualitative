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
  templateDir: path.join(repoRoot, "consultant_safe_v5"),
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

const QA_HEADERS = [
  "EBX Indicator",
  "Status",
  "Severity",
  "Length",
  "Sentence Count",
  "Metric Check",
  "Forbidden Language Check",
  "Repetition Check",
  "Company Name Check",
  "Short Finding",
];

const FORBIDDEN_FINAL_REGEX = /(Source|PDF|page|p\.\d+|출처|근거|본 답변|reviewer|audit|trace|검증 필요|원천자료|source|원문|임의로 0|임의 환산|추정치|대시|표시 기준|최종 문안|보고서 본문|보고서 문맥|구성합니다|정리합니다|Our Company|AppendixFacts|PrinciplePlanet|Facts & Figures|Samsung Electronics 2025)/i;
const KOREAN_REGEX = /[\u3131-\uD79D]/;

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
  const [rawHeaders, ...body] = rows.filter((r) => r.some((v) => v !== ""));
  if (!rawHeaders) return [];
  const headers = rawHeaders.map((h) => h.replace(/^\uFEFF/, ""));
  return body.map((r) => Object.fromEntries(headers.map((h, i) => [h, r[i] ?? ""])));
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

async function findTemplate(config) {
  const files = await fs.readdir(config.templateDir);
  const candidates = files
    .filter((file) => file.endsWith("_consultant_safe_v5.xlsx"))
    .filter((file) => file.includes(`_${config.sector}_`) && file.includes(`_${config.size}_`))
    .sort();
  if (!candidates.length) {
    throw new Error(`No v5 template found in ${config.templateDir} for sector=${config.sector}, size=${config.size}. Run scripts/build_consultant_safe_v5.mjs first, or pass --template-dir explicitly.`);
  }
  return path.join(config.templateDir, candidates[0]);
}

async function loadTemplateRows(templatePath) {
  const workbook = await SpreadsheetFile.importXlsx(await FileBlob.load(templatePath));
  const sheet = workbook.worksheets.getItem("EBX-Q 템플릿");
  const values = sheet.getRange("A1:X28").values;
  const headers = values[0];
  const idx = {
    ebx: headers.findIndex((header) => header === "ebx"),
    area: headers.findIndex((header) => header === "area"),
    pillar: headers.findIndex((header) => header === "pillar"),
    item: headers.findIndex((header) => header === "item"),
    contentSlots: headers.findIndex((header) => String(header ?? "").includes("Content Slots")),
    styleOptions: headers.findIndex((header) => String(header ?? "").includes("Style Options")),
    sentencePatterns: headers.findIndex((header) => String(header ?? "").includes("Sentence Patterns")),
    antiRepetition: headers.findIndex((header) => String(header ?? "").includes("Anti-Repetition")),
    fieldPath: headers.findIndex((header) => header === "Field Path"),
    finalAnswerRequirements: headers.findIndex((header) => header === "Final Answer Requirements"),
    reportReadyGuardrails: headers.findIndex((header) => header === "Report-Ready Guardrails"),
    coverageHandling: headers.findIndex((header) => header === "Coverage Handling"),
    metricHandling: headers.findIndex((header) => header === "Metric Handling"),
    evidenceSelectionRules: headers.findIndex((header) => header === "Evidence Selection Rules"),
    metricEvidenceRequirements: headers.findIndex((header) => header === "Metric Evidence Requirements"),
    companyNamingRule: headers.findIndex((header) => header === "Company Naming Rule"),
    antiRepetitionGroup: headers.findIndex((header) => header === "Anti-Repetition Group"),
    qaSeverity: headers.findIndex((header) => header === "QA Severity"),
  };
  const get = (row, key) => (idx[key] >= 0 ? row[idx[key]] ?? "" : "");
  return new Map(values.slice(1).map((row) => [
    get(row, "ebx"),
    {
      area: get(row, "area"),
      pillar: get(row, "pillar"),
      item: get(row, "item"),
      contentSlots: get(row, "contentSlots"),
      styleOptions: get(row, "styleOptions"),
      sentencePatterns: get(row, "sentencePatterns"),
      antiRepetition: get(row, "antiRepetition"),
      fieldPath: get(row, "fieldPath"),
      finalAnswerRequirements: get(row, "finalAnswerRequirements"),
      reportReadyGuardrails: get(row, "reportReadyGuardrails"),
      coverageHandling: get(row, "coverageHandling"),
      metricHandling: get(row, "metricHandling"),
      evidenceSelectionRules: get(row, "evidenceSelectionRules"),
      metricEvidenceRequirements: get(row, "metricEvidenceRequirements"),
      companyNamingRule: get(row, "companyNamingRule"),
      antiRepetitionGroup: get(row, "antiRepetitionGroup"),
      qaSeverity: get(row, "qaSeverity"),
    },
  ]));
}

function normalizeWhitespace(text) {
  return String(text ?? "")
    .replace(/\s+/g, " ")
    .trim();
}

function cleanEvidenceText(text) {
  return normalizeWhitespace(String(text ?? "")
    .replace(/\[[^\]]+\]/g, " ")
    .replace(/2024[–-]2025\s+LG전자\s+지속가능경영보고서\s+\d+/g, " ")
    .replace(/삼성전자\s+지속가능경영보고서\s+2025\s+\d+/g, " ")
    .replace(/Overview Environmental Social Governance ESG Data Appendix/gi, " ")
    .replace(/Our Company AppendixFacts & Figures PrinciplePlanet People/gi, " ")
    .replace(/AppendixFacts & Figures PrinciplePlanet People/gi, " ")
    .replace(/Facts & Figures|PrinciplePlanet|Our Company|Appendix/gi, " "));
}

function splitSentences(text) {
  const cleaned = cleanEvidenceText(text);
  const raw = cleaned
    .split(/(?<=[.!?。]|습니다\.|니다\.|됩니다\.|합니다\.|있습니다\.|않습니다\.)\s+|(?=LG전자는|삼성전자는|회사는|또한|특히|이를 통해)/g)
    .map((sentence) => normalizeWhitespace(sentence))
    .filter(Boolean);
  return raw
    .map((sentence) => sentence.replace(/^[-•\s]+/, ""))
    .filter((sentence) => sentence.length >= 35 && sentence.length <= 420)
    .filter((sentence) => KOREAN_REGEX.test(sentence))
    .filter((sentence) => !FORBIDDEN_FINAL_REGEX.test(sentence))
    .filter((sentence) => !/^[0-9\s.,%()~-]+$/.test(sentence));
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

function isMetricRow(row) {
  return String(row.evidence_type ?? "").includes("quantitative") || Boolean(String(row.quantitative_support ?? "").trim());
}

function rowPages(row) {
  return new Set(String(row.source_pages ?? "")
    .split(/[;,]/)
    .map((page) => page.trim())
    .filter(Boolean));
}

function numberTokens(text) {
  return String(text ?? "").match(/[0-9][0-9,]*(?:\.[0-9]+)?\s*(?:%|조\s*원|억\s*원|만\s*톤|톤|명|개|건|회|배|년|KRW|tCO2e|MWh|GWh|TJ|시간|점)?/g) ?? [];
}

function keywordTokens(row, template) {
  const text = `${row.question_title ?? ""} ${template?.pillar ?? ""} ${template?.item ?? ""} ${template?.contentSlots ?? ""}`;
  return uniqueByNormalized(text
    .replace(/[()/:,.;·•\[\]]/g, " ")
    .split(/\s+/)
    .map((token) => token.trim())
    .filter((token) => token.length >= 2)
    .filter((token) => !/^(and|the|for|with|및|관련|항목|관리|현황|체계|목표|정책)$/i.test(token)))
    .slice(0, 24);
}

function sentenceScore(sentence, row, template) {
  const keywords = keywordTokens(row, template);
  let score = 0;
  if (numberTokens(sentence).length) score += 7;
  if (/LG전자|삼성전자|회사는|이사회|위원회|조직|리스크|목표|전략|성과|안전|품질|정보보호|환경|윤리|인권|협력회사/.test(sentence)) score += 3;
  for (const keyword of keywords) {
    if (keyword && sentence.includes(keyword)) score += 1;
  }
  if (sentence.length > 90 && sentence.length < 260) score += 2;
  if (/Mission|Vision|Overview|Social Governance|Appendix/i.test(sentence)) score -= 8;
  return score;
}

function selectEvidenceSentences(row, template) {
  const sentences = uniqueByNormalized(splitSentences(row.original_text_ko));
  const withScores = sentences.map((sentence, index) => ({
    sentence,
    index,
    score: sentenceScore(sentence, row, template),
  }));
  const chosen = withScores
    .filter((item) => item.score > 0)
    .sort((a, b) => b.score - a.score || a.index - b.index)
    .slice(0, 4)
    .sort((a, b) => a.index - b.index)
    .map((item) => item.sentence);
  return chosen.length ? chosen : sentences.slice(0, 3);
}

function selectMetricRecords(row, quantitativeRows) {
  const pages = rowPages(row);
  const fieldText = `${row.question_title ?? ""}`.toLowerCase();
  const pageMatches = quantitativeRows.filter((metric) => pages.has(String(metric.source_page ?? "").trim()));
  const scored = pageMatches.map((metric, index) => {
    const haystack = `${metric.category ?? ""} ${metric.subcategory ?? ""} ${metric.indicator ?? ""} ${metric.notes ?? ""}`.toLowerCase();
    let score = 0;
    for (const token of fieldText.split(/[\s/()-]+/).filter((value) => value.length >= 3)) {
      if (haystack.includes(token)) score += 2;
    }
    if (metric.value_2024 || metric.value_2023 || metric.value_2022) score += 4;
    if (/incident|accident|injury|breach|violation|emission|energy|waste|recycle|training|complaint|grievance|audit|certification|safety|ethics|privacy|security/i.test(haystack)) score += 3;
    return { metric, index, score };
  });
  return scored
    .sort((a, b) => b.score - a.score || a.index - b.index)
    .slice(0, 2)
    .map((item) => item.metric);
}

function formatMetricValue(metric, key) {
  const value = String(metric[key] ?? "").trim();
  if (!value) return "";
  const unit = String(metric.unit ?? "").trim();
  return unit ? `${value} ${unit}` : value;
}

function metricSentenceFromRecords(records) {
  const parts = records
    .map((metric) => {
      const indicator = String(metric.indicator ?? "").trim();
      const value2024 = formatMetricValue(metric, "value_2024");
      const value2023 = formatMetricValue(metric, "value_2023");
      const value2022 = formatMetricValue(metric, "value_2022");
      if (!indicator || !value2024) return "";
      const trend = value2023 ? `, 2023년 ${value2023}` : value2022 ? `, 2022년 ${value2022}` : "";
      return `${indicator}은 2024년 ${value2024}${trend}`;
    })
    .filter(Boolean);
  if (!parts.length) return "";
  return `정량 지표는 ${parts.join("; ")}로 공시되어 해당 항목의 관리 범위와 성과 추이를 함께 보여줍니다.`;
}

function metricSentenceFromEvidence(row) {
  const candidates = splitSentences(`${row.quantitative_support ?? ""} ${row.original_text_ko ?? ""}`)
    .filter((sentence) => numberTokens(sentence).length)
    .sort((a, b) => numberTokens(b).length - numberTokens(a).length || a.length - b.length);
  return candidates[0] || "";
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
    if (name && name !== displayName) {
      out = out.replaceAll(name, displayName);
    }
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

function templateLines(text) {
  return String(text ?? "")
    .split(/\r?\n/)
    .map((line) => line.trim())
    .filter(Boolean);
}

function compactBlock(label, value) {
  const text = String(value ?? "").trim();
  return text ? `${label}\n${text}` : `${label}\nNot specified in template.`;
}

function selectStyleOption(row, template) {
  const field = `${template?.pillar ?? ""} ${template?.item ?? ""} ${row.question_title ?? ""}`;
  if (row.coverage_status === "PARTIAL" || isMetricRow(row)) return "evidence-led";
  if (/전략|목표|비전|Strategy/i.test(field)) return "narrative";
  if (/거버넌스|조직|책임|Governance/i.test(field)) return "formal";
  if (/위험|리스크|Risk/i.test(field)) return "formal";
  return "formal";
}

function selectPatternKey(row, template) {
  const field = `${template?.pillar ?? ""} ${template?.item ?? ""} ${row.question_title ?? ""}`;
  if (isMetricRow(row) || /지표|Metrics/i.test(field)) return "Pattern D";
  if (/거버넌스|조직|책임|Governance/i.test(field)) return "Pattern C";
  if (/위험|리스크|Risk/i.test(field)) return "Pattern B";
  if (/전략|목표|비전|Strategy/i.test(field)) return "Pattern D";
  return "Pattern A";
}

function selectPatternLine(sentencePatterns, patternKey) {
  const lines = templateLines(sentencePatterns);
  return lines.find((line) => line.includes(patternKey)) || patternKey;
}

function appliedContentSlots(contentSlots) {
  const lines = templateLines(contentSlots)
    .filter((line) => /Disclosure item|Required material|Company data slots|Metric slot/i.test(line))
    .slice(0, 7);
  return lines.length ? lines.join("\n") : "No specific content slots found.";
}

function composeStyleTemplateApplied(row, template) {
  const selectedOption = selectStyleOption(row, template);
  const patternKey = selectPatternKey(row, template);
  const optionDescriptions = {
    formal: "formal: ESG report tone with complete governance, process, KPI, and accountability detail.",
    concise: "concise: compact factual status when evidence is limited or the row only requires a short disclosure.",
    "evidence-led": "evidence-led: start with confirmed company facts and figures, then connect to policy, activity, or limitation.",
    narrative: "narrative: start from business context or materiality, then connect to governance, action, and next step.",
  };
  return [
    compactBlock("Selected Style Option", optionDescriptions[selectedOption] || selectedOption),
    compactBlock("Selected Sentence Pattern", selectPatternLine(template?.sentencePatterns, patternKey)),
    compactBlock("Applied Content Slots", appliedContentSlots(template?.contentSlots)),
    compactBlock("Applied Anti-Repetition Rule", `${template?.antiRepetitionGroup || "evidence-narrative"}: no exact Final Answer sentence may appear in more than two rows.`),
    compactBlock("Coverage Treatment Applied", row.coverage_status === "PARTIAL"
      ? "PARTIAL: disclose confirmed company facts and phrase missing disclosure as a business limitation."
      : "SUFFICIENT: use row-specific evidence, process, scope, activity, and performance context."),
    compactBlock("Metric Treatment Applied", isMetricRow(row)
      ? "Metric-led: include disclosed figures when present in qualitative evidence or quantitative CSV."
      : "Not metric-led: emphasize policy, governance, process, activity, and management status."),
    compactBlock("V5 QA Severity", template?.qaSeverity),
  ].join("\n\n");
}

function stripForbiddenFinalAnswerLanguage(text) {
  return normalizeWhitespace(String(text ?? "")
    .split(/(?<=[.!?。]|습니다\.|니다\.|됩니다\.|합니다\.|있습니다\.|않습니다\.)\s+/)
    .map((sentence) => sentence.trim())
    .filter((sentence) => sentence && !FORBIDDEN_FINAL_REGEX.test(sentence))
    .join(" "));
}

function topicLabel(row, template) {
  return String(template?.item || row.question_title || row.ebx)
    .replace(/\s*\/.*$/, "")
    .trim();
}

function bridgeSentence(row, template, displayName) {
  const topic = topicLabel(row, template);
  const group = template?.antiRepetitionGroup || "";
  if (group === "governance-organization") {
    return `${displayName}는 ${topic}와 관련해 책임 주체, 의사결정 흐름, 실행 조직의 역할을 연결해 관리 체계를 설명하고 있습니다.`;
  }
  if (group === "risk-control") {
    return `${displayName}는 ${topic} 영역에서 식별된 위험 요인, 예방 활동, 모니터링 절차를 함께 관리해 실행 수준을 점검하고 있습니다.`;
  }
  if (group === "metrics-status") {
    return `${displayName}는 ${topic}의 성과와 발생 현황을 보고기간별 지표와 운영 절차에 연결해 관리하고 있습니다.`;
  }
  return `${displayName}는 ${topic}를 사업 전략, 정책 방향, 실행 과제와 연결해 지속가능경영 활동의 맥락을 제시하고 있습니다.`;
}

function depthSentence(row, template, displayName) {
  const topic = topicLabel(row, template);
  const ebx = String(row.ebx ?? "").replace("EBX-Q-", "");
  const group = template?.antiRepetitionGroup || "";
  if (group === "governance-organization") {
    return `${displayName}의 ${topic} 항목은 EBX ${ebx} 기준에서 책임 조직의 역할, 보고 주기, 주요 안건 처리 방식을 함께 보여주는 관리 정보로 활용됩니다.`;
  }
  if (group === "risk-control") {
    return `${displayName}의 ${topic} 항목은 EBX ${ebx} 기준에서 위험 식별 이후의 예방 조치, 개선 활동, 후속 점검 흐름을 함께 설명하는 데 초점을 둡니다.`;
  }
  if (group === "metrics-status") {
    return `${displayName}의 ${topic} 항목은 EBX ${ebx} 기준에서 정량 성과와 사건 현황을 제도 운영 수준과 함께 해석할 수 있도록 구성됩니다.`;
  }
  return `${displayName}의 ${topic} 항목은 EBX ${ebx} 기준에서 정책 방향, 실행 범위, 성과 관리 관점을 함께 제시하는 서술로 정리됩니다.`;
}

function partialLimitationSentence(row, template, displayName) {
  const topic = topicLabel(row, template);
  return `${displayName}의 ${topic} 공시는 확인 가능한 제도와 활동을 중심으로 구성되어 있으며, 일부 세부 수치나 사건 현황은 현재 공개 범위에서 제한적으로 제시되어 향후 관리 정보 보완이 필요한 영역으로 남아 있습니다.`;
}

function hasBusinessLimitation(text) {
  return /제한적|공개 범위|보완|미공시|제시되어 있지|확인 가능한/.test(String(text ?? ""));
}

function capAnswer(answer, minLength = 560, maxLength = 920) {
  const normalized = normalizeWhitespace(answer);
  if (normalized.length <= maxLength) return normalized;
  const sentences = normalized.split(/(?<=[.!?。]|습니다\.|니다\.|됩니다\.|합니다\.|있습니다\.|않습니다\.)\s+/)
    .map((sentence) => sentence.trim())
    .filter(Boolean);
  const kept = [];
  for (const sentence of sentences) {
    const candidate = normalizeWhitespace([...kept, sentence].join(" "));
    if (candidate.length > maxLength && candidate.length >= minLength) break;
    kept.push(sentence);
  }
  return normalizeWhitespace(kept.join(" "));
}

function composeFinalAnswer(row, template, quantitativeRows, config, metadata) {
  const displayName = companyDisplayName(config, metadata);
  const evidence = selectEvidenceSentences(row, template).map((sentence) => replaceCompanyFallbacks(sentence, displayName, config, metadata));
  const metricRecords = isMetricRow(row) ? selectMetricRecords(row, quantitativeRows) : [];
  const metricSentence = isMetricRow(row)
    ? replaceCompanyFallbacks(metricSentenceFromRecords(metricRecords) || metricSentenceFromEvidence(row), displayName, config, metadata)
    : "";
  const pieces = [];

  for (const sentence of evidence) {
    if (pieces.length >= 4) break;
    pieces.push(sentence);
  }
  if (metricSentence && !pieces.some((sentence) => numberTokens(sentence).length)) {
    pieces.splice(Math.min(1, pieces.length), 0, metricSentence);
  }
  if (row.coverage_status === "PARTIAL") {
    pieces.splice(Math.min(2, pieces.length), 0, partialLimitationSentence(row, template, displayName));
  }
  if (pieces.length < 4) pieces.push(bridgeSentence(row, template, displayName));
  if (pieces.length < 5) pieces.push(depthSentence(row, template, displayName));

  let answer = stripForbiddenFinalAnswerLanguage(pieces.join(" "));
  answer = replaceCompanyFallbacks(answer, displayName, config, metadata);
  return capAnswer(answer, row.coverage_status === "SUFFICIENT" ? 560 : 420);
}

function finalSentences(text) {
  return normalizeWhitespace(text)
    .split(/(?<=[.!?。]|습니다\.|니다\.|됩니다\.|합니다\.|있습니다\.|않습니다\.)\s+/)
    .map((sentence) => sentence.trim())
    .filter((sentence) => sentence.length > 25);
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
  const sentenceOwners = new Map();
  for (const row of rows) {
    for (const sentence of finalSentences(row.finalAnswer)) {
      const key = normalizeWhitespace(sentence);
      sentenceOwners.set(key, [...(sentenceOwners.get(key) ?? []), row.ebx]);
    }
  }

  const removeByRow = new Map();
  for (const [sentence, owners] of sentenceOwners.entries()) {
    if (owners.length <= 2) continue;
    for (const ebx of owners.slice(2)) {
      removeByRow.set(ebx, [...(removeByRow.get(ebx) ?? []), sentence]);
    }
  }

  if (!removeByRow.size) return rows;
  return rows.map((row) => {
    const removals = new Set(removeByRow.get(row.ebx) ?? []);
    if (!removals.size) return row;
    const kept = finalSentences(row.finalAnswer)
      .filter((sentence) => !removals.has(normalizeWhitespace(sentence)));
    const displayName = row.displayName || "회사는";
    const repair = [
      bridgeSentence(row.sourceRow, row.template, displayName),
      depthSentence(row.sourceRow, row.template, displayName),
    ];
    return {
      ...row,
      finalAnswer: capAnswer(stripForbiddenFinalAnswerLanguage([...kept, ...repair].join(" ")), row.coverageStatus === "SUFFICIENT" ? 560 : 420),
    };
  });
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

function sourceHasNumbers(row, quantitativeRows) {
  if (numberTokens(`${row.original_text_ko ?? ""} ${row.quantitative_support ?? ""}`).length) return true;
  return selectMetricRecords(row, quantitativeRows).some((metric) => metric.value_2024 || metric.value_2023 || metric.value_2022);
}

function companyNameCheck(finalAnswer, displayName, config, metadata) {
  const fallbackNames = [
    config.companyName,
    config.companyName?.replaceAll("_", " "),
    metadata.company,
    "Samsung Electronics",
    "LG Electronics",
  ].filter(Boolean).filter((name) => name !== displayName);
  const hasFallback = fallbackNames.some((name) => name && finalAnswer.includes(name));
  return hasFallback ? "FAIL: English or ID fallback name remains." : "OK";
}

function repetitionCheck(finalAnswer, sentenceCounts) {
  const repeats = finalSentences(finalAnswer)
    .filter((sentence) => (sentenceCounts.get(normalizeWhitespace(sentence)) ?? 0) > 2);
  return repeats.length ? `FAIL: ${repeats.length} repeated sentence(s) exceed max count 2.` : "OK";
}

async function verifyWorkbook(workbook, rows, quantitativeRows, config, metadata) {
  const fatal = [];
  const warnings = [];
  const auditRows = [];
  const sentenceCounts = buildSentenceCounts(rows);
  const displayName = companyDisplayName(config, metadata);

  if (OUTPUT_HEADERS.includes("Area")) fatal.push("Output headers still include Area.");
  if (OUTPUT_HEADERS.includes("Writing Style Template")) fatal.push("Output headers still include Writing Style Template.");
  if (rows.length !== 27) fatal.push(`Expected 27 EBX rows, found ${rows.length}.`);

  for (const row of rows) {
    const findings = [];
    const fieldParts = String(row.field ?? "").split(" / ").filter(Boolean);
    const sentenceCount = finalSentences(row.finalAnswer).length;
    const metricExpected = isMetricRow(row.sourceRow) && sourceHasNumbers(row.sourceRow, quantitativeRows);
    const metricHasNumber = /[0-9]/.test(row.finalAnswer);
    const metricCheck = !isMetricRow(row.sourceRow)
      ? "N/A"
      : metricExpected && !metricHasNumber
        ? "WARN: source has numbers but Final Answer has none."
        : "OK";
    const forbiddenCheck = FORBIDDEN_FINAL_REGEX.test(row.finalAnswer) ? "FAIL: forbidden language found." : "OK";
    const repeatCheck = repetitionCheck(row.finalAnswer, sentenceCounts);
    const nameCheck = companyNameCheck(row.finalAnswer, displayName, config, metadata);

    if (fieldParts.length < 3) findings.push("Field does not include area / pillar / item.");
    if (!row.finalAnswer) findings.push("Blank final answer.");
    if (!KOREAN_REGEX.test(row.originalAnswer)) findings.push("Original answer does not contain Korean.");
    if (!KOREAN_REGEX.test(row.finalAnswer)) findings.push("Final answer does not contain Korean.");
    if (forbiddenCheck !== "OK") findings.push("Final answer contains source/citation/reviewer language.");
    if (repeatCheck !== "OK") findings.push("Repeated sentence exceeds workbook threshold.");
    if (nameCheck !== "OK") findings.push("Company fallback name remains in Korean prose.");
    if (!/Selected Style Option/.test(row.styleTemplate) || !/Selected Sentence Pattern/.test(row.styleTemplate) || !/Applied Content Slots/.test(row.styleTemplate)) {
      findings.push("Style Template Applied lacks required template detail.");
    }
    if (row.coverageStatus === "SUFFICIENT" && (row.finalAnswer.length < 560 || sentenceCount < 4)) {
      warnings.push(`${row.ebx}: SUFFICIENT final answer is short (${row.finalAnswer.length} chars, ${sentenceCount} sentences).`);
    }
    if (row.coverageStatus === "PARTIAL" && !hasBusinessLimitation(row.finalAnswer)) {
      warnings.push(`${row.ebx}: PARTIAL final answer lacks business limitation wording.`);
    }
    if (metricCheck.startsWith("WARN")) warnings.push(`${row.ebx}: metric row has disclosed number(s) but Final Answer has no numeric evidence.`);

    for (const finding of findings) fatal.push(`${row.ebx}: ${finding}`);

    const status = findings.length ? "FAIL" : metricCheck.startsWith("WARN") || (row.coverageStatus === "SUFFICIENT" && row.finalAnswer.length < 560) ? "WARN" : "PASS";
    const severity = findings.length ? "Fatal" : status === "WARN" ? "Warning" : "OK";
    auditRows.push([
      row.ebx,
      status,
      severity,
      row.finalAnswer.length,
      sentenceCount,
      metricCheck,
      forbiddenCheck,
      repeatCheck,
      nameCheck,
      findings.length ? findings.join(" | ") : status === "WARN" ? "Review warning checks before handoff." : "No finding.",
    ]);
  }

  const formulaErrors = await workbook.inspect({
    kind: "match",
    searchTerm: "#REF!|#DIV/0!|#VALUE!|#NAME\\?|#N/A",
    options: { useRegex: true, maxResults: 100 },
    summary: "v5 final formula error scan",
  });

  return {
    fatal,
    warnings,
    auditRows,
    formulaErrors: formulaErrors.ndjson,
    qaSummary: {
      rows: rows.length,
      pass: auditRows.filter((row) => row[1] === "PASS").length,
      warn: auditRows.filter((row) => row[1] === "WARN").length,
      fail: auditRows.filter((row) => row[1] === "FAIL").length,
      repeatedSentenceMax: Math.max(0, ...sentenceCounts.values()),
      metricRows: rows.filter((row) => isMetricRow(row.sourceRow)).length,
      metricRowsWithNoNumberDespiteSource: auditRows.filter((row) => String(row[5]).startsWith("WARN")).length,
    },
  };
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
    fill: "#0F5B4F",
    font: { bold: true, color: "#FFFFFF" },
  };
  sheet.getRange(`A1:F${outputRows.length + 1}`).format.wrapText = true;
  sheet.getRange(`A1:F${outputRows.length + 1}`).format.verticalAlignment = "top";
  sheet.getRange("A:A").format.columnWidthPx = 110;
  sheet.getRange("B:B").format.columnWidthPx = 300;
  sheet.getRange("C:C").format.columnWidthPx = 520;
  sheet.getRange("D:D").format.columnWidthPx = 360;
  sheet.getRange("E:E").format.columnWidthPx = 460;
  sheet.getRange("F:F").format.columnWidthPx = 620;
  sheet.freezePanes.freezeRows(1);
  sheet.freezePanes.freezeColumns(2);
  sheet.showGridLines = false;
}

function writeQaSheet(sheet, auditRows) {
  sheet.getRangeByIndexes(0, 0, 1, QA_HEADERS.length).values = [QA_HEADERS];
  sheet.getRangeByIndexes(1, 0, auditRows.length, QA_HEADERS.length).values = auditRows;
  sheet.getRange("A1:J1").format = {
    fill: "#2F3A45",
    font: { bold: true, color: "#FFFFFF" },
  };
  sheet.getRange(`A1:J${auditRows.length + 1}`).format.wrapText = true;
  sheet.getRange(`A1:J${auditRows.length + 1}`).format.verticalAlignment = "top";
  sheet.getRange("A:A").format.columnWidthPx = 110;
  sheet.getRange("B:C").format.columnWidthPx = 90;
  sheet.getRange("D:E").format.columnWidthPx = 90;
  sheet.getRange("F:I").format.columnWidthPx = 220;
  sheet.getRange("J:J").format.columnWidthPx = 420;
  sheet.freezePanes.freezeRows(1);
  sheet.showGridLines = false;
}

async function main() {
  const config = parseArgs(process.argv.slice(2));
  const metadata = await readJsonIfExists(path.join(config.dataDir, "metadata.json"));
  const qualitativeRows = parseCsv(await fs.readFile(path.join(config.dataDir, "data_dinh_tinh.csv"), "utf8"));
  const quantitativeRows = await readCsvIfExists(path.join(config.dataDir, "data_dinh_luong.csv"));
  const templatePath = await findTemplate(config);
  const templateRows = await loadTemplateRows(templatePath);

  const workbook = Workbook.create();
  const cleanSheet = workbook.worksheets.add("consultant_safe_v5");

  const outputRows = qualitativeRows.map((row) => {
    const template = templateRows.get(row.ebx) ?? {};
    const displayName = companyDisplayName(config, metadata);
    return {
      ebx: row.ebx,
      field: buildField(template, row),
      originalAnswer: row.original_text_ko || "",
      metadata: buildMetadata(row, {
        "Report title": metadata.report_title || metadata.report,
        "Reporting period": metadata.reporting_period,
      }),
      styleTemplate: composeStyleTemplateApplied(row, template),
      finalAnswer: composeFinalAnswer(row, template, quantitativeRows, config, metadata),
      coverageStatus: row.coverage_status,
      sourceRow: row,
      template,
      displayName,
    };
  });

  const finalRows = enforceRepetitionLimit(outputRows);
  writeCleanSheet(cleanSheet, finalRows);
  const qa = await verifyWorkbook(workbook, finalRows, quantitativeRows, config, metadata);
  const qaSheet = workbook.worksheets.add("QA_Audit");
  writeQaSheet(qaSheet, qa.auditRows);

  await fs.mkdir(config.outputDir, { recursive: true });
  const outputPath = path.join(config.outputDir, `${config.companyName}_EBX_Q_consultant_safe_v5_KO.xlsx`);
  const exported = await SpreadsheetFile.exportXlsx(workbook);
  await exported.save(outputPath);
  await cleanupInspectSidecar(outputPath);

  console.log(JSON.stringify({
    outputPath,
    templatePath,
    rows: outputRows.length,
    headers: OUTPUT_HEADERS,
    qaSummary: qa.qaSummary,
    fatal: qa.fatal,
    warnings: qa.warnings,
    formulaErrors: qa.formulaErrors,
  }, null, 2));
  if (qa.fatal.length) process.exitCode = 1;
}

await main();
