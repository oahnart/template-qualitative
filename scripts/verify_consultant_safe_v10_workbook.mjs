import fs from "node:fs/promises";
import path from "node:path";
import { fileURLToPath } from "node:url";
import { FileBlob, SpreadsheetFile } from "@oai/artifact-tool";

const __dirname = path.dirname(fileURLToPath(import.meta.url));
const repoRoot = path.resolve(__dirname, "..");
const inputPath = path.resolve(process.argv[2] ?? "");
const previewPath = path.resolve(process.argv[3] ?? path.join(repoRoot, ".tmp", "consultant_safe_v10_verify", "preview.png"));
if (!process.argv[2]) throw new Error("Usage: node scripts/verify_consultant_safe_v10_workbook.mjs <xlsx> [preview.png]");

const workbook = await SpreadsheetFile.importXlsx(await FileBlob.load(inputPath));
const sheet = workbook.worksheets.getItem("consultant_safe_v10");
const values = sheet.getRange("A1:F96").values;
const expectedHeaders = ["EBX Indicator", "Field", "Original Answer", "Original Answer Metadata", "Style Template Applied", "Final Answer"];
const issuePattern = /LG전자은|현황와|지표을|횟수은|안건 수은|목표은|관리은|tons|cases\/million hours|조직명|수행업무|주요 안건|구분\s+(?:주요 장비|운영주기)|샘플링 장비|분석 장비\s+ICP|전처리 장비|Corporate Governance|Strategy\s+Risk Management|업무협약의\s+네 가지\s+중점 분야|01\s+02\s+03\s+04|가이드라인\s+전문\s+바로가기|홈페이지\s+환경정책\s+링크|보고 페이지|코드 공시 항목|추진체계와 주요성과|Topic No\./i;
const issueRows = values.slice(1).flatMap((row) => {
  const matches = String(row[5] ?? "").match(issuePattern);
  return matches ? [{ ebx: row[0], matches: [...new Set(matches)] }] : [];
});
const formulaErrors = await workbook.inspect({
  kind: "match",
  searchTerm: "#REF!|#DIV/0!|#VALUE!|#NAME\\?|#N/A",
  options: { useRegex: true, maxResults: 100 },
  summary: "formula error scan",
});
const formulaErrorFound = /#REF!|#DIV\/0!|#VALUE!|#NAME\?|#N\/A/.test(formulaErrors.ndjson);
const fatal = [];
if (values.length !== 96) fatal.push(`Expected 96 rows including header, found ${values.length}.`);
if (values[0].join("|") !== expectedHeaders.join("|")) fatal.push("Workbook headers changed.");
if (issueRows.length) fatal.push(`Found ${issueRows.length} OCR, grammar, or untranslated-unit issue rows.`);
if (formulaErrorFound) fatal.push("Formula error found.");

await fs.mkdir(path.dirname(previewPath), { recursive: true });
const preview = await workbook.render({ sheetName: "consultant_safe_v10", range: "A1:F2", scale: 0.5, format: "png" });
await fs.writeFile(previewPath, new Uint8Array(await preview.arrayBuffer()));
const summary = { inputPath, previewPath, rows: values.length - 1, issueRows, formulaErrorFound, fatal };
console.log(JSON.stringify(summary, null, 2));
if (fatal.length) process.exitCode = 1;
