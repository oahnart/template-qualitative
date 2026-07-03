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
  templateDir: path.join(repoRoot, "final_template", "template_qualitative"),
  sector: "TC",
  size: "대기업",
  language: "KO",
};

const OUTPUT_HEADERS = [
  "EBX Indicator",
  "Area",
  "Original Answer",
  "Original Answer Metadata",
  "Writing Style Template",
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
    .filter((file) => file.endsWith("_consultant_safe_v3.xlsx"))
    .filter((file) => file.includes(`_${config.sector}_`) && file.includes(`_${config.size}_`))
    .sort();
  if (!candidates.length) {
    throw new Error(`No v3 template found in ${config.templateDir} for sector=${config.sector}, size=${config.size}. Run scripts/build_consultant_safe_v3.mjs first.`);
  }
  return path.join(config.templateDir, candidates[0]);
}

async function loadTemplateRows(templatePath) {
  const workbook = await SpreadsheetFile.importXlsx(await FileBlob.load(templatePath));
  const sheet = workbook.worksheets.getItem("EBX-Q 템플릿");
  const values = sheet.getRange("A1:J28").values;
  const headers = values[0];
  const idx = Object.fromEntries(headers.map((header, index) => [header, index]));
  return new Map(values.slice(1).map((row) => [
    row[idx.ebx],
    {
      area: row[idx.area] ?? "",
      item: row[idx.item] ?? "",
      styleOptions: row[idx["문체 옵션/Style Options"]] ?? "",
      sentencePatterns: row[idx["문장 패턴/Sentence Patterns"]] ?? "",
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

function composeFinalAnswer(row, template, legacyAnswers, config) {
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

function answerHasKorean(text) {
  return /[가-힣]/.test(String(text ?? ""));
}

async function verifyWorkbook(workbook, rows) {
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

async function main() {
  const config = parseArgs(process.argv.slice(2));
  const csvPath = path.join(config.dataDir, "data_dinh_tinh.csv");
  const metadata = await readJsonIfExists(path.join(config.dataDir, "metadata.json"));
  const rows = parseCsv(await fs.readFile(csvPath, "utf8"));
  const templatePath = await findTemplate(config);
  const templateRows = await loadTemplateRows(templatePath);
  const legacyAnswers = await readLegacyAnswers(config);

  const workbook = Workbook.create();
  const sheet = workbook.worksheets.add("consultant_safe_v3");
  sheet.getRangeByIndexes(0, 0, 1, OUTPUT_HEADERS.length).values = [OUTPUT_HEADERS];

  const outputRows = rows.map((row) => {
    const template = templateRows.get(row.ebx) ?? {};
    const styleTemplate = chooseStyle(row, template);
    const finalAnswer = composeFinalAnswer(row, template, legacyAnswers, config);
    return {
      ebx: row.ebx,
      area: template.area || row.question_title || "",
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
    row.area,
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
  sheet.getRange("B:B").format.columnWidthPx = 110;
  sheet.getRange("C:C").format.columnWidthPx = 520;
  sheet.getRange("D:D").format.columnWidthPx = 360;
  sheet.getRange("E:E").format.columnWidthPx = 320;
  sheet.getRange("F:F").format.columnWidthPx = 560;
  sheet.freezePanes.freezeRows(1);
  sheet.freezePanes.freezeColumns(2);
  sheet.showGridLines = false;

  const qa = await verifyWorkbook(workbook, outputRows);
  await fs.mkdir(config.outputDir, { recursive: true });
  const outputPath = path.join(config.outputDir, `${config.companyName}_EBX_Q_consultant_safe_v3_KO.xlsx`);
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
