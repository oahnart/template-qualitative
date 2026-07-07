import fs from "node:fs/promises";
import path from "node:path";
import { fileURLToPath } from "node:url";
import { FileBlob, SpreadsheetFile } from "@oai/artifact-tool";

const __dirname = path.dirname(fileURLToPath(import.meta.url));
const repoRoot = path.resolve(__dirname, "..");
const inputPath = path.resolve(process.argv[2] ?? "");
const previewPath = path.resolve(process.argv[3] ?? path.join(repoRoot, ".tmp", "consultant_safe_v10_1_verify", "preview.png"));
if (!process.argv[2]) throw new Error("Usage: node scripts/verify_consultant_safe_v10_1_workbook.mjs <xlsx> [preview.png]");

const expectedHeaders = ["EBX Indicator", "Field", "Original Answer", "Original Answer Metadata", "Style Template Applied", "Final Answer"];
const controlChars = /[\x00-\x08\x0B\x0C\x0E-\x1F]/;
const malformedMetric = /(?:202[0-9]\uB144\s+2,0[0-9]{2}\b)|(?:[,;]\s*(?:\uC740|\uB294|\uC774|\uAC00)\s+202[0-9]\uB144)/;
const workbook = await SpreadsheetFile.importXlsx(await FileBlob.load(inputPath));
const sheet = workbook.worksheets.getItem("consultant_safe_v10_1");
const values = sheet.getRange("A1:F96").values;
const answers = values.slice(1).map((row) => String(row[5] ?? ""));
const formulaErrors = await workbook.inspect({
  kind: "match",
  searchTerm: "#REF!|#DIV/0!|#VALUE!|#NAME\\?|#N/A",
  options: { useRegex: true, maxResults: 100 },
  summary: "formula error scan",
});
const fatal = [];
if (values.length !== 96) fatal.push(`Expected 96 rows including header, found ${values.length}.`);
if (values[0].join("|") !== expectedHeaders.join("|")) fatal.push("Workbook headers changed.");
if (answers.some((answer) => controlChars.test(answer))) fatal.push("Control character found in Final Answer.");
if (answers.some((answer) => malformedMetric.test(answer))) fatal.push("Malformed metric wording found in Final Answer.");
if (new Set(answers).size !== answers.length) fatal.push("Exact duplicate Final Answer found.");
if (/#REF!|#DIV\/0!|#VALUE!|#NAME\?|#N\/A/.test(formulaErrors.ndjson)) fatal.push("Formula error found.");

await fs.mkdir(path.dirname(previewPath), { recursive: true });
const preview = await workbook.render({ sheetName: "consultant_safe_v10_1", range: "A1:F96", scale: 0.15, format: "png" });
await fs.writeFile(previewPath, new Uint8Array(await preview.arrayBuffer()));
const summary = { inputPath, previewPath, rows: values.length - 1, headers: values[0], fatal };
console.log(JSON.stringify(summary, null, 2));
if (fatal.length) process.exitCode = 1;
