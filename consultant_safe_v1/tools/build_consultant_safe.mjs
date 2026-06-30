import crypto from "node:crypto";
import fs from "node:fs/promises";
import path from "node:path";
import { fileURLToPath } from "node:url";
import { SpreadsheetFile, Workbook } from "@oai/artifact-tool";

const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);
const releaseDir = path.resolve(__dirname, "..");
const repoRoot = path.resolve(releaseDir, "..");

const SOURCE_VERSION = "2026-06-15";
const METRIC_EBXS = new Set([
  "EBX-Q-007",
  "EBX-Q-011",
  "EBX-Q-015",
  "EBX-Q-019",
  "EBX-Q-023",
  "EBX-Q-027",
]);
const EXPECTED_SIZES = ["대기업", "중견", "중소", "비상장"];
const SAFE_HEADERS = [
  "ebx",
  "area",
  "pillar",
  "item",
  "공개성",
  "구조개요",
  "빈칸초안",
  "작성지침",
  "동종산업 예시",
  "동종산업 재서술 예시",
  "규모지침",
  "비고",
];

const SCALE_GUIDANCE = {
  대기업: [
    "대기업: mô tả đầy đủ governance, chính sách, phạm vi hợp nhất, KPI, mốc thời gian và bộ phận chịu trách nhiệm.",
    "Ưu tiên nêu bằng chứng/nguồn công khai nếu có, đồng thời đối chiếu với dữ liệu nội bộ trước khi nộp reviewer.",
  ],
  중견: [
    "중견: mô tả chính sách và hoạt động trọng yếu, phạm vi áp dụng, owner phụ trách và 1-3 KPI đang theo dõi.",
    "Nếu hệ thống chưa đầy đủ như doanh nghiệp lớn, nêu lộ trình chuẩn hóa và phần còn đang hoàn thiện.",
  ],
  중소: [
    "중소: có thể viết định tính gọn hơn; ưu tiên quy trình đang vận hành, người phụ trách, tình trạng phát sinh và bằng chứng nội bộ.",
    "Nếu thiếu dữ liệu định lượng, ghi rõ \"미집계\" hoặc \"해당사항 없음\"; không tự ước lượng.",
  ],
  비상장: [
    "비상장: nhấn mạnh dữ liệu nội bộ và mức sẵn sàng cho due diligence; không giả định nghĩa vụ công bố đại chúng.",
    "Nêu rõ phạm vi chưa công bố ra bên ngoài; metric chỉ dùng khi có nguồn nội bộ kiểm chứng được.",
  ],
};

const METRIC_SCALE_GUIDANCE = {
  대기업: "Dòng chỉ số: phải ghi phạm vi đo, kỳ đo, công thức hoặc nguồn hệ thống; không để trống và không biến \"chưa thống kê\" thành 0.",
  중견: "Dòng chỉ số: ghi KPI hiện có và phạm vi theo dõi; nếu chưa đủ chuỗi số liệu, nêu rõ phần thiếu thay vì điền 0.",
  중소: "Dòng chỉ số: nếu chưa thống kê, ghi \"미집계\" hoặc \"해당사항 없음\"; không để trống, không ước lượng và không điền 0 thay cho dữ liệu thiếu.",
  비상장: "Dòng chỉ số: chỉ ghi số khi có chứng từ nội bộ; nếu dùng cho due diligence, ghi rõ nguồn và không điền 0 cho phần chưa thống kê.",
};

async function exists(dir) {
  try {
    const stat = await fs.stat(dir);
    return stat.isDirectory();
  } catch {
    return false;
  }
}

async function resolveSourceDir() {
  const candidates = [
    process.env.EBX_SOURCE_DIR,
    path.join(repoRoot, "ebx_template_review_44"),
    path.join(repoRoot, "..", "template-gen-44-new", "ebx_template_review_44"),
    path.join(repoRoot, "..", "template-gen-44", "ebx_template_review_44"),
    path.join(repoRoot, "..", "template-gen-44", "ebx_template_review_44_vi"),
  ].filter(Boolean);

  for (const candidate of candidates) {
    const resolved = path.resolve(candidate);
    if (await exists(resolved)) return resolved;
  }

  throw new Error(`Source directory not found. Set EBX_SOURCE_DIR or restore ${path.join(repoRoot, "ebx_template_review_44")}.`);
}

function parseCsv(text) {
  const rows = [];
  let row = [];
  let cur = "";
  let quoted = false;
  for (let i = 0; i < text.length; i += 1) {
    const ch = text[i];
    if (ch === "\"") {
      if (quoted && text[i + 1] === "\"") {
        cur += "\"";
        i += 1;
      } else {
        quoted = !quoted;
      }
    } else if (ch === "," && !quoted) {
      row.push(cur);
      cur = "";
    } else if ((ch === "\n" || ch === "\r") && !quoted) {
      if (ch === "\r" && text[i + 1] === "\n") i += 1;
      row.push(cur);
      if (row.some((value) => value !== "")) rows.push(row);
      row = [];
      cur = "";
    } else {
      cur += ch;
    }
  }
  if (cur || row.length) {
    row.push(cur);
    rows.push(row);
  }
  const headers = rows[0].map((header) => header.replace(/^\uFEFF/, ""));
  return rows.slice(1).map((values) =>
    Object.fromEntries(headers.map((header, index) => [header, values[index] ?? ""])),
  );
}

function bulletList(values) {
  if (!values) return "";
  if (Array.isArray(values)) return values.map((value) => `• ${value}`).join("\n");
  return String(values);
}

function hashText(text) {
  return crypto.createHash("sha256").update(text, "utf8").digest("hex");
}

function maskNumerals(text) {
  return String(text ?? "")
    .replace(/\d+(?:[.,]\d+)?(?:\s?[%건명개회월년시간원억조])?/g, "[ ]")
    .replace(/\[\s+\]/g, "[ ]");
}

function safeExample(value) {
  const text = Array.isArray(value) ? value.join("\n") : String(value ?? "");
  return maskNumerals(text);
}

function safeReword(value) {
  const text = String(value ?? "");
  if (!text || text.startsWith("(모델 재서술 미채택") || text.startsWith("(동종 실제사례 미확보")) {
    return "(Ví dụ chưa có — dùng 구조개요/작성지침 và điền theo dữ liệu nội bộ.)";
  }
  return maskNumerals(text.replace(/\(예시·실데이터 기반 재서술·수치 마스킹\)/g, "(ví dụ tham khảo, đã mask số)"));
}

function filePrefixFromMeta(meta) {
  return `${meta.sasb_sector}_${meta.sasb_name}_${meta.size}`;
}

function buildWarningNote(meta, item, topicFlags) {
  const notes = [];
  if (meta.mapping_confidence === "low" || meta.fallback_used) {
    notes.push("⚠ Mapping confidence thấp/fallback: consultant cần xác nhận ngành trước khi nhập.");
  }
  const topicFlag = topicFlags.get(`${meta.sasb_sector}|${item.ebx}`);
  if (topicFlag?.verdict === "borderline") {
    notes.push("⚠ Review note: ví dụ ngành từng bị đánh dấu borderline; chỉ dùng cấu trúc, không xem là đáp án chuẩn.");
  }
  if (topicFlag?.verdict === "미확보") {
    notes.push("⚠ Review note: ví dụ thực tế chưa 확보; ưu tiên cấu trúc và dữ liệu nội bộ.");
  }
  return notes.join(" / ");
}

function buildGuidance(item) {
  const lines = Array.isArray(item["작성지침"]) ? [...item["작성지침"]] : [String(item["작성지침"] ?? "")];
  if (METRIC_EBXS.has(item.ebx)) {
    lines.push("Nếu chưa thống kê chỉ số, ghi rõ \"미집계\" hoặc \"해당사항 없음\"; không ước lượng hoặc bịa số liệu.");
  }
  return bulletList(lines.filter(Boolean));
}

function buildScaleGuidance(meta, item) {
  const sizeGuidance = SCALE_GUIDANCE[meta.size];
  if (!sizeGuidance) {
    throw new Error(`Unsupported size "${meta.size}" in ${meta.sasb_sector}.`);
  }
  const lines = [...sizeGuidance];
  if (METRIC_EBXS.has(item.ebx)) {
    lines.push(METRIC_SCALE_GUIDANCE[meta.size]);
  }
  return bulletList(lines);
}

function buildScaleGuideSummary(size) {
  const lines = SCALE_GUIDANCE[size];
  if (!lines) {
    throw new Error(`Unsupported size "${size}".`);
  }
  return `${size}. ${lines.join(" ")}`;
}

async function writeMarkdownDocs(summary) {
  const readme = `# Consultant-safe EBX-Q Template v1

## Mục đích
Bộ này dành cho consultant nhập liệu ESG định tính theo EBX-Q. Đây là bản an toàn được dựng từ source \`${SOURCE_VERSION}\`, đã loại cột J chứa dữ liệu thực tế của công ty cùng ngành và đã bổ sung hướng dẫn riêng theo 4 quy mô.

## Cách dùng nhanh
1. Chọn workbook theo ngành và quy mô doanh nghiệp.
2. Vào sheet \`EBX-Q 템플릿\`, dùng \`빈칸초안\` làm câu khung và điền mọi vị trí \`[ ]\` bằng dữ liệu nội bộ.
3. Đọc \`작성지침\` và \`규모지침\` trước khi nhập; \`규모지침\` khác nhau cho \`대기업\`, \`중견\`, \`중소\`, \`비상장\`.
4. Với chỉ số chưa thống kê, ghi \`미집계\` hoặc \`해당사항 없음\`; không để trống, không ước lượng, không điền \`0\` thay cho dữ liệu thiếu.
5. Chỉ dùng \`동종산업 예시\` và \`동종산업 재서술 예시\` để tham khảo cấu trúc diễn đạt, không copy như dữ liệu của doanh nghiệp.
6. Trước khi gửi reviewer, chạy checklist trong \`CHECKLIST.md\`.

## Nội dung phát hành
- 44 workbook consultant-safe, đủ 11 sector x 4 quy mô.
- Cột thực dữ liệu \`동종 실제사례(완성문·회사명 마스킹·수치유지)\` đã bị loại.
- Cột \`규모지침\` được sinh riêng theo từng quy mô và được QA chống trùng giữa 4 quy mô trong cùng sector.
- Các workbook sector SV có cảnh báo mapping confidence thấp/fallback.
- Các mục từng bị review borderline có cảnh báo để consultant dùng ví dụ thận trọng.

## Thống kê build
- Workbooks: ${summary.workbooks}
- Source JSON: ${summary.sourceJson}
- Scale guidance variants checked: ${summary.scaleGuidanceChecked}
- QA fatal issues: ${summary.fatalIssues}
- QA warnings: ${summary.warningIssues}
`;

  const checklist = `# Checklist Nhập Liệu Trước Khi Gửi Reviewer

## Checklist theo workbook
- [ ] Chọn đúng ngành và quy mô doanh nghiệp.
- [ ] Đọc dòng "Quy mô" trong sheet \`안내\` và cột \`규모지침\` trong sheet \`EBX-Q 템플릿\`.
- [ ] Nếu workbook thuộc SV 서비스, đã xác nhận lại ngành/mapping với reviewer.
- [ ] Không sửa các cột định danh: \`ebx\`, \`area\`, \`pillar\`, \`item\`.
- [ ] Không thêm lại cột J hoặc nội dung \`실데이터 기반·복사금지\`.

## Checklist theo từng EBX-Q item
- [ ] Tất cả \`[ ]\` trong câu dùng để nộp đã được điền bằng dữ liệu nội bộ hoặc đổi thành \`미집계/해당사항 없음\`.
- [ ] Nội dung đã khớp quy mô: \`대기업\` đầy đủ governance/KPI; \`중견\` có chính sách, owner và KPI chính; \`중소\` có thể định tính gọn; \`비상장\` nhấn mạnh dữ liệu nội bộ/due diligence.
- [ ] Không bịa số, mục tiêu, năm, tổ chức, chứng chỉ hoặc quy trình nếu không có nguồn nội bộ.
- [ ] Với Q-007/Q-011/Q-015/Q-019/Q-023/Q-027, nếu không có thống kê thì ghi \`미집계\` hoặc \`해당사항 없음\`; không để trống hoặc điền \`0\` thay cho dữ liệu thiếu.
- [ ] Ví dụ ngành chỉ được dùng để tham khảo cấu trúc, không copy thành câu trả lời.
- [ ] Các dòng có cảnh báo trong \`비고\` đã được đọc và xử lý thận trọng.

## Checklist trước khi bàn giao
- [ ] Một reviewer ESG đã đọc các mục có \`⚠\`.
- [ ] Các số liệu đã đối chiếu với nguồn nội bộ.
- [ ] Các tên tổ chức/chức danh/chứng chỉ đã đúng theo văn bản doanh nghiệp.
- [ ] File đã lưu thành bản riêng cho khách hàng/doanh nghiệp, không ghi đè template gốc.
`;

  const changelog = `# Version Log

## consultant_safe_v1
- Source: \`ebx_template_review_44\`, version \`${SOURCE_VERSION}\`.
- Created release package for consultant nhập liệu.
- Removed public-facing real-data column: \`동종 실제사례(완성문·회사명 마스킹·수치유지)\`.
- Renamed reworded example column to \`동종산업 재서술 예시\`.
- Added concise Vietnamese guidance to \`안내\` sheets.
- Added scale-specific \`규모지침\` for \`대기업\`, \`중견\`, \`중소\`, and \`비상장\`.
- Added QA checks that fail when the 4 scale guidance variants are duplicated inside a sector.
- Added warnings for SV fallback mapping and topic-review borderline/missing items.
- Masked numeric sequences in consultant-facing example columns.
- Added README, checklist, release index, and QA report.
`;

  await fs.writeFile(path.join(releaseDir, "README.md"), readme, "utf8");
  await fs.writeFile(path.join(releaseDir, "CHECKLIST.md"), checklist, "utf8");
  await fs.writeFile(path.join(releaseDir, "CHANGELOG.md"), changelog, "utf8");
}

function createGuideRows(meta, fileName) {
  const mappingText = meta.mapping_confidence === "low" || meta.fallback_used
    ? `산업='${meta.industry}' / 주요산업='${meta.business}' → SASB ${meta.sasb_sector}(${meta.sasb_name}), confidence=low/fallback. Consultant phải xác nhận ngành trước khi nhập.`
    : `산업='${meta.industry}' / 주요산업='${meta.business}' → SASB ${meta.sasb_sector}(${meta.sasb_name}), confidence=${meta.mapping_confidence}.`;

  return [
    [`Consultant-safe EBX-Q Template — ${filePrefixFromMeta(meta)}`, ""],
    ["Nguồn", `Dựng từ ${fileName}; source version ${SOURCE_VERSION}. Source gốc không bị sửa.`],
    ["Mục đích", "Workbook dành cho consultant nhập liệu ESG định tính. Đây là bản an toàn, không chứa cột J thực dữ liệu của công ty khác."],
    ["Cách điền [ ]", "Điền bằng dữ liệu nội bộ có nguồn. Không bịa số liệu, mục tiêu, tổ chức, chứng chỉ hoặc quy trình."],
    ["Khi thiếu dữ liệu", "Ghi rõ \"미집계\" nếu chưa thống kê hoặc \"해당사항 없음\" nếu không áp dụng. Không được lượng hoá."],
    ["Ví dụ ngành", "Cột ví dụ chỉ dùng để tham khảo cấu trúc diễn đạt; không copy làm dữ liệu của doanh nghiệp."],
    ["Mapping ngành", mappingText],
    ["Quy mô", buildScaleGuideSummary(meta.size)],
    ["Trước khi gửi reviewer", "Đọc CHECKLIST.md, kiểm tra các dòng có ⚠ trong cột 비고, và đối chiếu mọi số liệu với nguồn nội bộ."],
  ];
}

async function buildWorkbook(sourceDir, sourceFile, topicFlags) {
  const sourcePath = path.join(sourceDir, sourceFile);
  const raw = await fs.readFile(sourcePath, "utf8");
  const data = JSON.parse(raw);
  const meta = data._meta;
  const outName = sourceFile.replace(".json", "_consultant_safe.xlsx");
  const outPath = path.join(releaseDir, outName);

  const workbook = Workbook.create();
  const guide = workbook.worksheets.add("안내");
  const sheet = workbook.worksheets.add("EBX-Q 템플릿");

  const guideRows = createGuideRows(meta, sourceFile);
  guide.getRangeByIndexes(0, 0, guideRows.length, 2).values = guideRows;
  guide.getRange("A1:B1").merge();
  guide.getRange("A1").format = {
    fill: "#1F4E79",
    font: { bold: true, color: "#FFFFFF", size: 15 },
  };
  guide.getRange(`A2:A${guideRows.length}`).format = {
    fill: "#D9EAF7",
    font: { bold: true, color: "#17365D" },
  };
  guide.getRange(`A1:B${guideRows.length}`).format.wrapText = true;
  guide.getRange(`A1:B${guideRows.length}`).format.verticalAlignment = "top";
  guide.getRange("A1:A9").format.columnWidthPx = 170;
  guide.getRange("B1:B9").format.columnWidthPx = 820;
  guide.freezePanes.freezeRows(1);
  guide.showGridLines = false;

  const scaleGuidanceByEbx = new Map();
  const rows = data.items.map((item) => {
    const scaleGuidance = buildScaleGuidance(meta, item);
    scaleGuidanceByEbx.set(item.ebx, scaleGuidance);
    const baseNote = String(item["비고"] ?? "");
    const warning = buildWarningNote(meta, item, topicFlags);
    const note = [baseNote, warning].filter(Boolean).join(" / ");
    return [
      item.ebx,
      item.area,
      item.pillar,
      item.item,
      item["공개성"],
      bulletList(item["구조개요"]),
      item["빈칸초안"] ?? "",
      buildGuidance(item),
      safeExample(item["동종예시"]),
      safeReword(item["동종실제재서술"]),
      scaleGuidance,
      note,
    ];
  });

  sheet.getRangeByIndexes(0, 0, 1, SAFE_HEADERS.length).values = [SAFE_HEADERS];
  sheet.getRangeByIndexes(1, 0, rows.length, SAFE_HEADERS.length).values = rows;
  sheet.getRange("A1:L1").format = {
    fill: "#1F4E79",
    font: { bold: true, color: "#FFFFFF" },
  };
  sheet.getRange("A1:L28").format.wrapText = true;
  sheet.getRange("A1:L28").format.verticalAlignment = "top";
  sheet.getRange("G1:H28").format.fill = "#FFF2CC";
  sheet.getRange("I1:J28").format.fill = "#E2F0D9";
  sheet.getRange("K1:L28").format.fill = "#FCE4D6";
  sheet.getRange("A1:A28").format.columnWidthPx = 92;
  sheet.getRange("B1:B28").format.columnWidthPx = 72;
  sheet.getRange("C1:C28").format.columnWidthPx = 130;
  sheet.getRange("D1:D28").format.columnWidthPx = 220;
  sheet.getRange("E1:E28").format.columnWidthPx = 98;
  sheet.getRange("F1:F28").format.columnWidthPx = 260;
  sheet.getRange("G1:G28").format.columnWidthPx = 330;
  sheet.getRange("H1:H28").format.columnWidthPx = 370;
  sheet.getRange("I1:I28").format.columnWidthPx = 370;
  sheet.getRange("J1:J28").format.columnWidthPx = 370;
  sheet.getRange("K1:K28").format.columnWidthPx = 420;
  sheet.getRange("L1:L28").format.columnWidthPx = 320;
  sheet.freezePanes.freezeRows(1);
  sheet.freezePanes.freezeColumns(4);
  sheet.showGridLines = false;

  const output = await SpreadsheetFile.exportXlsx(workbook);
  await output.save(outPath);
  return {
    outName,
    outPath,
    sourceFile,
    meta,
    rowCount: rows.length,
    scaleGuidanceHash: hashText([...scaleGuidanceByEbx.entries()].map(([ebx, guidance]) => `${ebx}:${guidance}`).join("\n")),
    metricGuidance: scaleGuidanceByEbx.get("EBX-Q-007") ?? "",
  };
}

async function main() {
  await fs.mkdir(releaseDir, { recursive: true });
  const sourceDir = await resolveSourceDir();
  const sourceFiles = (await fs.readdir(sourceDir))
    .filter((file) => file.endsWith(".json") && !file.startsWith("_"))
    .sort();
  const topicRows = parseCsv(await fs.readFile(path.join(sourceDir, "_topic_review.csv"), "utf8"));
  const topicFlags = new Map(topicRows.map((row) => [`${row.sector}|${row.ebx}`, row]));
  const outputs = [];

  for (const file of sourceFiles) {
    outputs.push(await buildWorkbook(sourceDir, file, topicFlags));
  }

  const indexRows = [
    "file,sector_code,sector_name,size,mapping_confidence,source_json,rows,sv_mapping_warning,scale_guidance_hash",
    ...outputs.map(({ outName, sourceFile, meta, rowCount, scaleGuidanceHash }) => [
      outName,
      meta.sasb_sector,
      meta.sasb_name,
      meta.size,
      meta.mapping_confidence,
      sourceFile,
      rowCount,
      meta.mapping_confidence === "low" || meta.fallback_used ? "yes" : "no",
      scaleGuidanceHash,
    ].map((value) => `"${String(value).replaceAll("\"", "\"\"")}"`).join(",")),
  ];
  await fs.writeFile(path.join(releaseDir, "RELEASE_INDEX.csv"), `${indexRows.join("\n")}\n`, "utf8");

  const qa = await runQa(sourceDir, outputs);
  await fs.writeFile(path.join(releaseDir, "QA_REPORT.json"), `${JSON.stringify(qa, null, 2)}\n`, "utf8");
  await writeMarkdownDocs({
    workbooks: outputs.length,
    sourceJson: sourceFiles.length,
    scaleGuidanceChecked: qa.scaleGuidance.checkedSectors,
    fatalIssues: qa.fatal.length,
    warningIssues: qa.warnings.length,
  });

  console.log(JSON.stringify({
    releaseDir,
    sourceDir,
    workbooks: outputs.length,
    fatal: qa.fatal.length,
    warnings: qa.warnings.length,
  }, null, 2));
  if (qa.fatal.length) {
    process.exitCode = 1;
  }
}

async function runQa(sourceDir, outputs) {
  const fatal = [];
  const warnings = [];
  const seenFiles = new Set(outputs.map((output) => output.outName));
  if (seenFiles.size !== 44) fatal.push(`Expected 44 workbooks, found ${seenFiles.size}.`);

  for (const output of outputs) {
    const sourceData = JSON.parse(await fs.readFile(path.join(sourceDir, output.sourceFile), "utf8"));
    if (sourceData.items.length !== 27) fatal.push(`${output.outName}: source item count is ${sourceData.items.length}, expected 27.`);
    const ids = sourceData.items.map((item) => item.ebx);
    if (new Set(ids).size !== 27) fatal.push(`${output.outName}: duplicate EBX ids.`);
    const expectedIds = Array.from({ length: 27 }, (_, index) => `EBX-Q-${String(index + 1).padStart(3, "0")}`);
    for (const id of expectedIds) {
      if (!ids.includes(id)) fatal.push(`${output.outName}: missing ${id}.`);
    }
    const jsonText = JSON.stringify(sourceData);
    if (!jsonText.includes("[ ]")) warnings.push(`${output.outName}: source JSON has no [ ] placeholder, unusual.`);
    if (!EXPECTED_SIZES.includes(output.meta.size)) {
      fatal.push(`${output.outName}: unsupported size ${output.meta.size}.`);
    }
    if (!output.metricGuidance.includes("0") || !/không để trống|không điền 0|thay vì điền 0|thay cho dữ liệu thiếu/.test(output.metricGuidance)) {
      fatal.push(`${output.outName}: EBX-Q-007 metric scale guidance does not warn against blank/zero substitution.`);
    }
    if ((output.meta.mapping_confidence === "low" || output.meta.fallback_used) && output.meta.sasb_sector !== "SV") {
      warnings.push(`${output.outName}: low/fallback mapping outside SV.`);
    }
  }

  const bySector = new Map();
  for (const output of outputs) {
    if (!bySector.has(output.meta.sasb_sector)) bySector.set(output.meta.sasb_sector, []);
    bySector.get(output.meta.sasb_sector).push(output);
  }

  for (const [sector, sectorOutputs] of bySector.entries()) {
    const sizes = new Set(sectorOutputs.map((output) => output.meta.size));
    for (const size of EXPECTED_SIZES) {
      if (!sizes.has(size)) fatal.push(`${sector}: missing ${size} consultant-safe workbook.`);
    }
    const hashes = new Set(sectorOutputs.map((output) => output.scaleGuidanceHash));
    if (hashes.size !== EXPECTED_SIZES.length) {
      fatal.push(`${sector}: expected 4 distinct scale guidance variants, found ${hashes.size}.`);
    }
    const bySize = Object.fromEntries(sectorOutputs.map((output) => [output.meta.size, output.scaleGuidanceHash]));
    if (bySize["중소"] && bySize["비상장"] && bySize["중소"] === bySize["비상장"]) {
      fatal.push(`${sector}: 중소 and 비상장 scale guidance are identical.`);
    }
    if (bySize["대기업"] && bySize["중견"] && bySize["대기업"] === bySize["중견"]) {
      fatal.push(`${sector}: 대기업 and 중견 scale guidance are identical.`);
    }
  }

  return {
    generatedAt: new Date().toISOString(),
    sourceVersion: SOURCE_VERSION,
    expectedHeaders: SAFE_HEADERS,
    scaleGuidance: {
      expectedSizes: EXPECTED_SIZES,
      checkedSectors: bySector.size,
    },
    fatal,
    warnings,
  };
}

await main();
