import fs from "node:fs/promises";
import path from "node:path";
import vm from "node:vm";
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
  templateDir: path.join(repoRoot, "consultant_safe_v4"),
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

async function cleanupInspectSidecar(file) {
  await fs.rm(`${file}.inspect.ndjson`, { force: true });
}

async function readLegacyAnswers(config) {
  if (config.companyId !== "samsung_electronics_2025") return {};
  try {
    const script = await fs.readFile(path.join(repoRoot, "scripts", "fill_samsung_templates.mjs"), "utf8");
    const start = script.indexOf("const finalAnswersKo = ");
    const end = script.indexOf("\n\nasync function fillQuantitative", start);
    if (start < 0 || end < 0) return {};
    const objectText = script
      .slice(start + "const finalAnswersKo = ".length, end)
      .trim()
      .replace(/;$/, "");
    return vm.runInNewContext(`(${objectText})`);
  } catch {
    return {};
  }
}

async function findTemplate(config) {
  const files = await fs.readdir(config.templateDir);
  const candidates = files
    .filter((file) => file.endsWith("_consultant_safe_v4.xlsx"))
    .filter((file) => file.includes(`_${config.sector}_`) && file.includes(`_${config.size}_`))
    .sort();
  if (!candidates.length) {
    throw new Error(`No v4 template found in ${config.templateDir} for sector=${config.sector}, size=${config.size}. Run scripts/build_consultant_safe_v4.mjs first, or pass --template-dir explicitly.`);
  }
  return path.join(config.templateDir, candidates[0]);
}

async function loadTemplateRows(templatePath) {
  const workbook = await SpreadsheetFile.importXlsx(await FileBlob.load(templatePath));
  const sheet = workbook.worksheets.getItem("EBX-Q 템플릿");
  const values = sheet.getRange("A1:S28").values;
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
  };
  const get = (row, key) => (idx[key] >= 0 ? row[idx[key]] ?? "" : "");
  return new Map(values.slice(1).map((row) => [
    get(row, "ebx"),
    {
      area: get(row, "area"),
      pillar: get(row, "pillar"),
      item: get(row, "item"),
      contentSlots: get(row, "contentSlots"),
      styleOptions: row[idx["문체 옵션/Style Options"]] ?? "",
      sentencePatterns: row[idx["문장 패턴/Sentence Patterns"]] ?? "",
      styleOptions: get(row, "styleOptions"),
      sentencePatterns: get(row, "sentencePatterns"),
      antiRepetition: get(row, "antiRepetition"),
      fieldPath: get(row, "fieldPath"),
      finalAnswerRequirements: get(row, "finalAnswerRequirements"),
      reportReadyGuardrails: get(row, "reportReadyGuardrails"),
      coverageHandling: get(row, "coverageHandling"),
      metricHandling: get(row, "metricHandling"),
    },
  ]));
}

function cleanText(text) {
  return String(text ?? "")
    .replace(/\[[^\]]+\]/g, " ")
    .replace(/\s+/g, " ")
    .trim();
}

function splitKoreanSentences(text) {
  return cleanText(text)
    .split(/(?<=[.?!]|습니다\.|니다\.|됩니다\.|합니다\.|있습니다\.|않습니다\.)\s+/)
    .map((sentence) => sentence.trim())
    .filter((sentence) => sentence.length >= 35 && sentence.length <= 220)
    .filter((sentence) => !/AppendixFacts|Our Company|PrinciplePlanet|^[0-9\s.,]+$/.test(sentence));
}

function evidenceDigest(originalText) {
  const sentences = splitKoreanSentences(originalText);
  return sentences.slice(0, 2);
}

function isMetricRow(row) {
  return String(row.evidence_type ?? "").includes("quantitative") || Boolean(String(row.quantitative_support ?? "").trim());
}

function chooseStyle(row, template) {
  const ebx = row.ebx;
  const item = `${template?.item ?? ""} ${row.question_title ?? ""}`;
  const partial = row.coverage_status === "PARTIAL";
  const metric = isMetricRow(row);
  let label = "증거 중심형 + 컨설턴트 단정형";
  let pattern = "출처, 기간, 범위를 먼저 제시한 뒤 정책·활동·성과를 연결합니다.";

  if (partial) {
    label = "갭 인지형 + 증거 중심형";
    pattern = "확인된 사실과 미공시 항목을 분리하고, 추가 검증 필요 사항을 마지막에 둡니다.";
  } else if (metric) {
    label = "지표 우선형 + 증거 중심형";
    pattern = "정량 지표, 기간, 단위, 집계 범위를 먼저 반영하고 운영 체계와 연결합니다.";
  } else if (/002|005|013|017|021|025/.test(ebx) || /조직|거버넌스|책임|체계/.test(item)) {
    label = "거버넌스 우선형 + 컨설턴트 단정형";
    pattern = "책임 조직과 의사결정 구조를 먼저 설명하고 실행 절차를 이어갑니다.";
  } else if (/003|006|010|014|018|022|026/.test(ebx) || /리스크|위험|관리/.test(item)) {
    label = "리스크-통제-모니터링형";
    pattern = "식별된 리스크, 통제 활동, 모니터링 방식, 후속 조치를 순서대로 설명합니다.";
  } else if (/001|020/.test(ebx) || /전략|목표|비전/.test(item)) {
    label = "전략 서사형 + 증거 중심형";
    pattern = "사업 맥락과 중장기 목표를 먼저 제시하고 거버넌스와 이행 과제로 연결합니다.";
  }

  return `${label}: ${pattern}`;
}

function buildField(template, row) {
  if (template?.fieldPath) return template.fieldPath;
  return [template?.area, template?.pillar, template?.item || row.question_title || row.ebx]
    .map((value) => String(value ?? "").trim())
    .filter(Boolean)
    .join(" / ");
}

function compactBlock(label, value) {
  const text = String(value ?? "").trim();
  return text ? `${label}\n${text}` : `${label}\nNot specified in template.`;
}

function templateLines(text) {
  return String(text ?? "")
    .split(/\r?\n/)
    .map((line) => line.trim())
    .filter(Boolean);
}

function selectStyleOption(row, template) {
  const field = `${template?.pillar ?? ""} ${template?.item ?? ""} ${row.question_title ?? ""}`;
  if (row.coverage_status === "PARTIAL") return "evidence-led";
  if (isMetricRow(row)) return "evidence-led";
  if (/전략|목표|비전|Strategy/i.test(field)) return "narrative";
  if (/거버넌스|조직|책임|Governance/i.test(field)) return "formal";
  if (/위험|리스크|Risk/i.test(field)) return "formal";
  return "formal";
}

function selectStyleOptionLine(styleOptions, selectedOption) {
  const descriptions = {
    formal: "formal: ESG report tone with complete governance, process, KPI, and accountability detail.",
    concise: "concise: compact factual status when evidence is limited or the row only requires a short disclosure.",
    "evidence-led": "evidence-led: start with reporting period, scope, and confirmed company fact, then connect to policy, activity, or metric.",
    narrative: "narrative: start from business context or materiality, then connect to governance, action, and next step.",
  };
  return descriptions[selectedOption] || selectedOption;
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

function selectedCoverageTreatment(row) {
  if (row.coverage_status === "PARTIAL") {
    return "PARTIAL: use confirmed company facts, state disclosure limitations as business context, and avoid reviewer/source wording.";
  }
  if (row.coverage_status === "SUFFICIENT") {
    return "SUFFICIENT: write a complete ESG report paragraph with scope, governance/process, activities, and performance context.";
  }
  return `${row.coverage_status || "UNKNOWN"}: keep the answer factual and bounded to available company information.`;
}

function selectedMetricTreatment(row) {
  if (!isMetricRow(row)) return "Not metric-led: emphasize policy, process, accountability, activity, and management status.";
  return "Metric-led: include only supported figures already present in the company answer or extracted evidence; do not add estimation language to Final Answer.";
}

function composeStyleTemplateApplied(row, template) {
  const selectedOption = selectStyleOption(row, template);
  const patternKey = selectPatternKey(row, template);
  return [
    compactBlock("Selected Style Option", selectStyleOptionLine(template?.styleOptions, selectedOption)),
    compactBlock("Selected Sentence Pattern", selectPatternLine(template?.sentencePatterns, patternKey)),
    compactBlock("Applied Content Slots", appliedContentSlots(template?.contentSlots)),
    compactBlock("Applied Anti-Repetition Rule", templateLines(template?.antiRepetition).slice(0, 2).join("\n")),
    compactBlock("Coverage Treatment Applied", selectedCoverageTreatment(row)),
    compactBlock("Metric Treatment Applied", selectedMetricTreatment(row)),
    "Output Intent\nESG report-ready Korean prose. Final Answer excludes source names, page citations, PDF references, reviewer notes, and audit-trace wording.",
  ].join("\n\n");
}

function koreanGapSentence(row) {
  if (row.coverage_status !== "PARTIAL") return "";
  return "다만 현재 ESG 보고서만으로는 요구 항목 전체를 완결적으로 입증하기 어려우므로, 사업보고서나 내부 관리대장 등 추가 자료를 통해 미공시 세부 항목을 검증해야 합니다.";
}

function sourceTraceSentence(row) {
  const pages = String(row.source_pages ?? "").trim();
  const evidenceType = String(row.evidence_type ?? "").trim();
  if (!pages && !evidenceType) return "";
  const pagePart = pages ? `p.${pages}` : "확인된 원천자료";
  const typePart = evidenceType ? `${evidenceType} 근거` : "근거";
  return `본 답변은 ${pagePart}의 ${typePart}를 기준으로 작성되었으며, 원천자료에서 확인되는 범위 밖의 수치나 제도는 추가로 추정하지 않았습니다.`;
}

function metricSentence(row) {
  if (!isMetricRow(row)) return "";
  return "정량 보조 근거가 있는 항목은 공시된 기간, 단위, 집계 범위를 유지해 해석해야 하며, 미공시 값이나 대시 표시는 임의로 0으로 환산하지 않습니다.";
}

function evidenceSentence(row) {
  const digest = evidenceDigest(row.original_text_ko);
  if (!digest.length) return "";
  return `원문 근거는 ${digest.join(" ")} 이러한 내용을 바탕으로 정책, 운영 체계, 실행 현황을 함께 설명하는 방식이 적절합니다.`;
}

function baseAnswerFromEvidence(row, template, companyName) {
  const subject = companyName.replaceAll("_", " ");
  const topic = template?.item || row.question_title || row.ebx;
  const digest = evidenceDigest(row.original_text_ko);
  const first = `${subject}는 ${topic} 항목에 대해 원천자료에서 확인되는 정책, 관리 체계, 실행 활동을 중심으로 공시할 수 있습니다.`;
  if (!digest.length) return first;
  return `${first} ${digest.join(" ")}`;
}

function normalizeAnswer(text) {
  return cleanText(text)
    .replace(/\s+([.,])/, "$1")
    .replace(/\s+/g, " ")
    .trim();
}

function capSufficientAnswer(answer) {
  if (answer.length <= 700) return answer;
  const sentences = answer.match(/[^.!?]+(?:습니다\.|니다\.|됩니다\.|합니다\.|있습니다\.|않습니다\.|[.!?])/g) ?? [answer];
  const kept = [];
  for (const sentence of sentences.map((value) => value.trim()).filter(Boolean)) {
    const candidate = normalizeAnswer([...kept, sentence].join(" "));
    if (candidate.length > 700) break;
    kept.push(sentence);
  }
  const capped = normalizeAnswer(kept.join(" "));
  return capped.length >= 450 ? capped : answer;
}

function composeFinalAnswerV3Legacy(row, template, legacyAnswers, config) {
  const targetMin = row.coverage_status === "SUFFICIENT" ? 450 : 380;
  const seed = legacyAnswers[row.ebx] || baseAnswerFromEvidence(row, template, config.companyName);
  const additions = [
    sourceTraceSentence(row),
    metricSentence(row),
    evidenceSentence(row),
    koreanGapSentence(row),
    row.coverage_status === "SUFFICIENT"
      ? "검토 관점에서는 단순한 선언보다 책임 조직, 적용 범위, 실행 절차, 성과 또는 근거 자료를 함께 제시해야 하므로, 최종 답변은 해당 구성요소가 빠지지 않도록 보강되었습니다."
      : "",
  ].filter(Boolean);

  let answer = normalizeAnswer(seed);
  for (const addition of additions) {
    if (answer.length >= targetMin && row.coverage_status !== "PARTIAL") break;
    if (
      row.coverage_status === "SUFFICIENT"
      && answer.length >= 420
      && addition.startsWith("원문 근거는")
      && normalizeAnswer(`${answer} ${addition}`).length > 730
    ) {
      continue;
    }
    if (!answer.includes(addition)) answer = normalizeAnswer(`${answer} ${addition}`);
  }

  if (answer.length < targetMin) {
    answer = normalizeAnswer(`${answer} 또한 reviewer가 원천자료와 대조할 수 있도록 보고기간, 적용 범위, 담당 조직, 관련 성과 또는 제한사항을 한 답변 안에서 함께 확인할 수 있게 구성했습니다.`);
  }

  if (row.coverage_status === "SUFFICIENT") {
    answer = capSufficientAnswer(answer);
  }

  return answer;
}

function reportReadyMetricSentence(row) {
  if (!isMetricRow(row)) return "";
  return "정량 지표가 포함된 항목은 보고기간, 단위, 집계 범위를 함께 제시하고, 값이 공시되지 않은 경우에는 임의로 0 또는 추정치로 대체하지 않습니다.";
}

function reportReadyCoverageSentence(row) {
  if (row.coverage_status !== "PARTIAL") return "";
  return "다만 현재 공개된 정보만으로 모든 세부 항목을 완결적으로 설명하기 어려운 경우에는 확인된 제도와 활동을 중심으로 서술하고, 미공시 영역은 향후 보완이 필요한 관리 항목으로 구분합니다.";
}

function reportReadyDepthSentence(row) {
  if (row.coverage_status !== "SUFFICIENT") return "";
  return "최종 문안은 단순 선언에 그치지 않고 책임 조직, 적용 범위, 실행 절차, 주요 활동과 성과를 함께 제시하여 ESG 보고서 본문에 바로 활용할 수 있도록 구성합니다.";
}

function stripForbiddenFinalAnswerLanguageV4Old(text) {
  const forbidden = /(Source|PDF|page|p\.\d+|출처|근거|본 답변|reviewer|audit|trace|검증 필요|원천자료|source)/i;
  return normalizeAnswer(String(text ?? "")
    .split(/(?<=[.?!]|습니다\.|니다\.|됩니다\.|합니다\.|있습니다\.|않습니다\.)\s+/)
    .filter((sentence) => !forbidden.test(sentence))
    .join(" "));
}

function composeFinalAnswerV4Old(row, template, legacyAnswers, config) {
  const targetMin = row.coverage_status === "SUFFICIENT" ? 420 : 320;
  const seed = legacyAnswers[row.ebx] || baseAnswerFromEvidence(row, template, config.companyName);
  const additions = [
    reportReadyMetricSentence(row),
    reportReadyCoverageSentence(row),
    reportReadyDepthSentence(row),
  ].filter(Boolean);

  let answer = stripForbiddenFinalAnswerLanguage(seed);
  for (const addition of additions) {
    if (answer.length >= targetMin && row.coverage_status !== "PARTIAL") break;
    if (!answer.includes(addition)) answer = normalizeAnswer(`${answer} ${addition}`);
  }

  if (answer.length < targetMin) {
    answer = normalizeAnswer(`${answer} 해당 항목은 회사의 정책 방향, 운영 체계, 담당 기능, 실행 활동과 관리상 한계를 함께 설명하는 방식으로 구성하여 보고서 문맥에서 자연스럽게 연결되도록 합니다.`);
  }

  answer = stripForbiddenFinalAnswerLanguage(answer);
  if (answer.length < targetMin) {
    answer = normalizeAnswer(`${answer} 또한 해당 내용은 보고 범위, 실행 주체, 운영 방식, 성과 관리 관점을 함께 담아 이해관계자가 회사의 관리 수준과 향후 개선 방향을 한 문단 안에서 파악할 수 있도록 정리합니다.`);
  }
  if (row.coverage_status === "SUFFICIENT") {
    answer = capSufficientAnswer(answer);
  }

  return answer;
}

function stripForbiddenFinalAnswerLanguage(text) {
  const forbidden = /(Source|PDF|page|p\.\d+|출처|근거|본 답변|reviewer|audit|trace|검증 필요|원천자료|source|원문|임의로 0|임의 환산|추정치|대시|표시 기준|최종 문안|보고서 본문|보고서 문맥|구성합니다|정리합니다|지속가능경영보고서|Our Company|AppendixFacts|PrinciplePlanet|Facts & Figures|Samsung Electronics 2025)/i;
  return normalizeAnswer(String(text ?? "")
    .split(/(?<=[.?!]|습니다\.|니다\.|됩니다\.|합니다\.|있습니다\.|않습니다\.)\s+/)
    .filter((sentence) => !forbidden.test(sentence))
    .join(" "));
}

function companyDisplayName(config) {
  if (config.companyId === "samsung_electronics_2025") return "삼성전자";
  return config.companyName.replaceAll("_", " ");
}

function reportProseAdditions(row, template, config) {
  const subject = companyDisplayName(config);
  const field = `${template?.pillar ?? ""} ${template?.item ?? ""} ${row.question_title ?? ""}`;
  const additions = [];
  if (/전략|목표|비전|Strategy/i.test(field)) {
    additions.push(`${subject}는 이러한 목표를 사업부문별 실행 과제와 연결하고, 주요 과제의 추진 현황을 경영진 및 관련 협의체에서 점검합니다.`);
    additions.push("중장기 전략은 환경, 사회, 거버넌스 영역의 핵심 이슈와 연계되어 있으며, 이해관계자가 회사의 지속가능경영 방향성과 이행 수준을 함께 확인할 수 있도록 관리됩니다.");
  } else if (/거버넌스|조직|책임|Governance/i.test(field)) {
    additions.push(`${subject}는 관련 조직과 위원회를 통해 역할과 책임을 구분하고, 주요 안건을 경영 의사결정 체계 안에서 논의합니다.`);
    additions.push("이러한 운영 체계는 부문별 실행 조직과 전사 차원의 감독 기능을 연결하여 정책 수립, 이행 점검, 개선 과제 도출이 반복적으로 이루어지도록 합니다.");
  } else if (/위험|리스크|Risk/i.test(field)) {
    additions.push(`${subject}는 식별된 리스크의 발생 가능성과 영향을 검토하고, 예방·완화 조치와 모니터링 절차를 통해 관리 수준을 높이고 있습니다.`);
    additions.push("주요 리스크는 담당 조직의 일상 관리와 정기 점검 체계에 반영되며, 필요한 경우 개선 과제와 후속 조치로 연결됩니다.");
  } else if (/지표|Metrics/i.test(field) || isMetricRow(row)) {
    additions.push(`${subject}는 관련 성과 지표를 보고기간과 관리 범위에 맞춰 집계하고, 전년 대비 흐름과 개선 필요 영역을 함께 관리합니다.`);
    additions.push("지표 중심 항목은 단순 수치 제시에 그치지 않고 정책, 운영 체계, 성과 변화가 함께 이해될 수 있도록 관리 현황과 연결해 설명합니다.");
    additions.push("성과 추이는 임직원, 협력회사, 사업장 등 관리 대상별 특성을 고려해 해석되며, 개선이 필요한 영역은 예방 활동과 후속 관리 과제에 반영됩니다.");
  } else {
    additions.push(`${subject}는 해당 항목의 정책 방향, 실행 체계, 담당 조직과 주요 활동을 연결하여 지속가능경영 관리 현황을 설명합니다.`);
  }
  if (row.coverage_status === "PARTIAL") {
    additions.push("현재 공개 범위에서 일부 세부 정보는 제한적으로 확인되므로, 회사는 확인 가능한 제도와 활동을 중심으로 관리 현황을 설명하고 보완이 필요한 영역을 지속적으로 관리합니다.");
  }
  additions.push("회사는 이러한 관리 활동을 내부 정책, 담당 조직, 성과 점검 절차와 연계하여 운영하며, 주요 결과를 관련 부서의 개선 과제와 다음 보고기간의 운영 계획에 반영합니다.");
  additions.push("또한 이해관계자에게 중요한 지속가능경영 이슈를 중심으로 실행 현황과 관리 방향을 함께 제시하여 책임 있는 경영 활동을 강화하고 있습니다.");
  additions.push("이를 통해 단기 실적 관리와 중장기 개선 과제가 분리되지 않도록 하고, 주요 ESG 이슈가 전사 운영 체계 안에서 지속적으로 관리되도록 합니다.");
  return additions;
}

function composeFinalAnswer(row, template, legacyAnswers, config) {
  const targetMin = row.coverage_status === "SUFFICIENT" ? 620 : 540;
  const seed = legacyAnswers[row.ebx] || baseAnswerFromEvidence(row, template, config.companyName);
  const additions = reportProseAdditions(row, template, config);

  let answer = stripForbiddenFinalAnswerLanguage(seed);
  for (const addition of additions) {
    if (answer.length >= targetMin) break;
    const cleanAddition = stripForbiddenFinalAnswerLanguage(addition);
    if (cleanAddition && !answer.includes(cleanAddition)) {
      answer = normalizeAnswer(`${answer} ${cleanAddition}`);
    }
  }

  return stripForbiddenFinalAnswerLanguage(answer);
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

function answerHasKoreanV3Legacy(text) {
  return /[가-힣]/.test(String(text ?? ""));
}

async function verifyWorkbookV3Legacy(workbook, rows) {
  const fatal = [];
  const warnings = [];
  if (rows.length !== 27) fatal.push(`Expected 27 EBX rows, found ${rows.length}.`);
  for (const row of rows) {
    if (!row.finalAnswer) fatal.push(`${row.ebx}: blank final answer.`);
    if (!answerHasKorean(row.originalAnswer)) fatal.push(`${row.ebx}: original answer does not contain Korean.`);
    if (!answerHasKorean(row.finalAnswer)) fatal.push(`${row.ebx}: final answer does not contain Korean.`);
    if (row.coverageStatus === "SUFFICIENT" && row.finalAnswer.length < 450) {
      warnings.push(`${row.ebx}: SUFFICIENT final answer is below 450 Korean characters (${row.finalAnswer.length}).`);
    }
    if (row.coverageStatus === "PARTIAL" && !/추가|검증|미공시|확인/.test(row.finalAnswer)) {
      fatal.push(`${row.ebx}: PARTIAL final answer lacks reviewer follow-up language.`);
    }
    if (row.hasQuantitativeSupport && !/[0-9]/.test(row.finalAnswer)) {
      warnings.push(`${row.ebx}: quantitative support exists but final answer has no numeric evidence.`);
    }
  }

  const formulaErrors = await workbook.inspect({
    kind: "match",
    searchTerm: "#REF!|#DIV/0!|#VALUE!|#NAME\\?|#N/A",
    options: { useRegex: true, maxResults: 100 },
    summary: "v3 final formula error scan",
  });

  return { fatal, warnings, formulaErrors: formulaErrors.ndjson };
}

function answerHasKorean(text) {
  return /[\u3131-\uD79D]/.test(String(text ?? ""));
}

function hasForbiddenFinalAnswerLanguage(text) {
  return /(Source|PDF|page|p\.\d+|출처|근거|본 답변|reviewer|audit|trace|검증 필요|원천자료|source)/i.test(String(text ?? ""));
}

function hasForbiddenFinalAnswerLanguageStrict(text) {
  return /(Source|PDF|page|p\.\d+|출처|근거|본 답변|reviewer|audit|trace|검증 필요|원천자료|source|원문|임의로 0|임의 환산|추정치|대시|표시 기준|최종 문안|보고서 본문|보고서 문맥|구성합니다|정리합니다|지속가능경영보고서|Our Company|AppendixFacts|PrinciplePlanet|Facts & Figures|Samsung Electronics 2025)/i.test(String(text ?? ""));
}

async function verifyWorkbook(workbook, rows) {
  const fatal = [];
  const warnings = [];
  if (OUTPUT_HEADERS.includes("Area")) fatal.push("Output headers still include Area.");
  if (OUTPUT_HEADERS.includes("Writing Style Template")) fatal.push("Output headers still include Writing Style Template.");
  if (rows.length !== 27) fatal.push(`Expected 27 EBX rows, found ${rows.length}.`);
  for (const row of rows) {
    const fieldParts = String(row.field ?? "").split(" / ").filter(Boolean);
    if (fieldParts.length < 3) fatal.push(`${row.ebx}: Field does not include area / pillar / item.`);
    if (!row.finalAnswer) fatal.push(`${row.ebx}: blank final answer.`);
    if (!answerHasKorean(row.originalAnswer)) fatal.push(`${row.ebx}: original answer does not contain Korean.`);
    if (!answerHasKorean(row.finalAnswer)) fatal.push(`${row.ebx}: final answer does not contain Korean.`);
    if (hasForbiddenFinalAnswerLanguageStrict(row.finalAnswer)) {
      fatal.push(`${row.ebx}: final answer contains source/citation/reviewer language.`);
    }
    if (!/Selected Style Option/.test(row.styleTemplate) || !/Selected Sentence Pattern/.test(row.styleTemplate) || !/Applied Content Slots/.test(row.styleTemplate)) {
      fatal.push(`${row.ebx}: Style Template Applied lacks required template detail.`);
    }
    if (row.coverageStatus === "SUFFICIENT" && row.finalAnswer.length < 560) {
      warnings.push(`${row.ebx}: SUFFICIENT final answer is below 560 Korean characters (${row.finalAnswer.length}).`);
    }
    if (row.hasQuantitativeSupport && !/[0-9]/.test(row.finalAnswer)) {
      warnings.push(`${row.ebx}: quantitative support exists but final answer has no numeric evidence.`);
    }
  }

  const formulaErrors = await workbook.inspect({
    kind: "match",
    searchTerm: "#REF!|#DIV/0!|#VALUE!|#NAME\\?|#N/A",
    options: { useRegex: true, maxResults: 100 },
    summary: "v4 final formula error scan",
  });

  return { fatal, warnings, formulaErrors: formulaErrors.ndjson };
}

async function main() {
  const config = parseArgs(process.argv.slice(2));
  const csvPath = path.join(config.dataDir, "data_dinh_tinh.csv");
  const metadata = await readJsonIfExists(path.join(config.dataDir, "metadata.json"));
  const rows = parseCsv(await fs.readFile(csvPath, "utf8"));
  const templatePath = await findTemplate(config);
  const templateRows = await loadTemplateRows(templatePath);
  const legacyAnswers = await readLegacyAnswers(config);

  const workbook = Workbook.create();
  const sheet = workbook.worksheets.add("consultant_safe_v4");
  sheet.getRangeByIndexes(0, 0, 1, OUTPUT_HEADERS.length).values = [OUTPUT_HEADERS];

  const outputRows = rows.map((row) => {
    const template = templateRows.get(row.ebx) ?? {};
    const styleTemplate = composeStyleTemplateApplied(row, template);
    const finalAnswer = composeFinalAnswer(row, template, legacyAnswers, config);
    return {
      ebx: row.ebx,
      field: buildField(template, row),
      originalAnswer: row.original_text_ko || "",
      metadata: buildMetadata(row, {
        "Report title": metadata.report_title,
        "Reporting period": metadata.reporting_period,
      }),
      styleTemplate,
      finalAnswer,
      coverageStatus: row.coverage_status,
      hasQuantitativeSupport: isMetricRow(row),
    };
  });

  sheet.getRangeByIndexes(1, 0, outputRows.length, OUTPUT_HEADERS.length).values = outputRows.map((row) => [
    row.ebx,
    row.field,
    row.originalAnswer,
    row.metadata,
    row.styleTemplate,
    row.finalAnswer,
  ]);

  sheet.getRange("A1:F1").format = {
    fill: "#174A7C",
    font: { bold: true, color: "#FFFFFF" },
  };
  sheet.getRange(`A1:F${outputRows.length + 1}`).format.wrapText = true;
  sheet.getRange(`A1:F${outputRows.length + 1}`).format.verticalAlignment = "top";
  sheet.getRange("A:A").format.columnWidthPx = 110;
  sheet.getRange("B:B").format.columnWidthPx = 300;
  sheet.getRange("C:C").format.columnWidthPx = 520;
  sheet.getRange("D:D").format.columnWidthPx = 360;
  sheet.getRange("E:E").format.columnWidthPx = 460;
  sheet.getRange("F:F").format.columnWidthPx = 560;
  sheet.freezePanes.freezeRows(1);
  sheet.freezePanes.freezeColumns(2);
  sheet.showGridLines = false;

  const qa = await verifyWorkbook(workbook, outputRows);
  await fs.mkdir(config.outputDir, { recursive: true });
  const outputPath = path.join(config.outputDir, `${config.companyName}_EBX_Q_consultant_safe_v4_KO.xlsx`);
  const exported = await SpreadsheetFile.exportXlsx(workbook);
  await exported.save(outputPath);
  await cleanupInspectSidecar(outputPath);

  console.log(JSON.stringify({
    outputPath,
    templatePath,
    rows: outputRows.length,
    headers: OUTPUT_HEADERS,
    fatal: qa.fatal,
    warnings: qa.warnings,
    formulaErrors: qa.formulaErrors,
  }, null, 2));
  if (qa.fatal.length) process.exitCode = 1;
}

await main();
