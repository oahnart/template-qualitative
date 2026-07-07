import fs from "node:fs/promises";
import path from "node:path";
import { spawn } from "node:child_process";
import { fileURLToPath } from "node:url";
import { FileBlob, SpreadsheetFile } from "@oai/artifact-tool";

const __dirname = path.dirname(fileURLToPath(import.meta.url));
const repoRoot = path.resolve(__dirname, "..");
const fixtureDir = path.join(repoRoot, "tests", "fixtures", "consultant_safe_v10", "profile_matrix");
const outputRoot = path.join(repoRoot, ".tmp", "consultant_safe_v10_1_profile_tests");
const cases = JSON.parse(await fs.readFile(path.join(fixtureDir, "cases.json"), "utf8")).cases;

function runFill(args) {
  return new Promise((resolve) => {
    const child = spawn(process.execPath, [path.join(repoRoot, "scripts", "fill_company_qualitative_v10_1.mjs"), ...args], {
      cwd: repoRoot,
      stdio: ["ignore", "pipe", "pipe"],
    });
    let stdout = "";
    let stderr = "";
    child.stdout.on("data", (chunk) => { stdout += chunk; });
    child.stderr.on("data", (chunk) => { stderr += chunk; });
    child.on("close", (code) => resolve({ code, stdout, stderr }));
  });
}

function assert(condition, message) {
  if (!condition) throw new Error(message);
}

await fs.mkdir(outputRoot, { recursive: true });
const results = [];
for (const fixture of cases) {
  const outputDir = path.join(outputRoot, fixture.id);
  const companyName = `Profile_Fixture_${fixture.id}`;
  const run = await runFill([
    `--company-id=profile_fixture_${fixture.id}`,
    `--company-name=${companyName}`,
    `--data-dir=${fixtureDir}`,
    `--output-dir=${outputDir}`,
    "--sector=TC",
    `--size=${fixture.size}`,
  ]);
  assert(run.code === 0, `${fixture.id}: fill failed: ${run.stderr || run.stdout}`);
  const qaPath = path.join(outputDir, `${companyName}_EBX_Q_consultant_safe_v10_1_KO_QA.json`);
  const xlsxPath = path.join(outputDir, `${companyName}_EBX_Q_consultant_safe_v10_1_KO.xlsx`);
  const qa = JSON.parse(await fs.readFile(qaPath, "utf8"));
  assert(qa.rowCount === 95, `${fixture.id}: expected 95 rows.`);
  assert(qa.summary.fail === 0 && qa.summary.fatal === 0, `${fixture.id}: QA fail/fatal must be zero.`);
  assert(qa.summary.topicMismatchedMetricRows === 0, `${fixture.id}: topic-mismatched metric detected.`);
  assert(qa.templatePath.includes(`TC_${fixture.size}_consultant_safe_v10_1.json`), `${fixture.id}: wrong template selected.`);
  assert(qa.summary.controlCharacterRows === 0, `${fixture.id}: control character reached output.`);
  assert(qa.summary.malformedMetricRows === 0, `${fixture.id}: malformed metric reached output.`);
  assert(qa.summary.duplicateFinalAnswerGroups === 0, `${fixture.id}: duplicate final answers detected.`);

  const workbook = await SpreadsheetFile.importXlsx(await FileBlob.load(xlsxPath));
  const sheet = workbook.worksheets.getItem("consultant_safe_v10_1");
  const values = sheet.getRange("A1:F96").values;
  assert(values.length === 96, `${fixture.id}: workbook row count mismatch.`);
  assert(values[0].join("|") === "EBX Indicator|Field|Original Answer|Original Answer Metadata|Style Template Applied|Final Answer", `${fixture.id}: output headers changed.`);
  const byEbx = new Map(values.slice(1).map((row) => [row[0], row]));
  const q28Style = String(byEbx.get("EBX-Q-028")?.[4] ?? "");
  const q36Answer = String(byEbx.get("EBX-Q-036")?.[5] ?? "");
  const q72Answer = String(byEbx.get("EBX-Q-072")?.[5] ?? "");
  assert(q28Style.includes(`\"id\":\"${fixture.size}\"`), `${fixture.id}: size profile missing from applied style.`);
  assert(q28Style.includes(`\"maxEvidenceSentences\":${fixture.expectedMaxEvidence}`), `${fixture.id}: evidence limit mismatch.`);
  assert(q28Style.includes(`\"maxRecords\":${fixture.expectedMaxMetrics}`), `${fixture.id}: metric limit mismatch.`);
  assert(!/샘플링 장비|분석 장비\s+ICP|구분\s+주요 장비/.test(q36Answer), `${fixture.id}: OCR/table residue remains.`);
  assert(/추가 확인/.test(q72Answer) && !/정보보호|개인정보|제품 보안|사이버/.test(q72Answer), `${fixture.id}: mismatched evidence was not converted to an information gap.`);
  const formulaErrors = await workbook.inspect({
    kind: "match",
    searchTerm: "#REF!|#DIV/0!|#VALUE!|#NAME\\?|#N/A",
    options: { useRegex: true, maxResults: 100 },
    summary: `${fixture.id} formula error scan`,
  });
  assert(!/#REF!|#DIV\/0!|#VALUE!|#NAME\?|#N\/A/.test(formulaErrors.ndjson), `${fixture.id}: formula error detected.`);
  const preview = await workbook.render({ sheetName: "consultant_safe_v10_1", range: "A1:F6", scale: 0.5, format: "png" });
  await fs.writeFile(path.join(outputDir, "preview.png"), new Uint8Array(await preview.arrayBuffer()));
  results.push({
    id: fixture.id,
    size: fixture.size,
    pass: qa.summary.pass,
    warn: qa.summary.warn,
    fail: qa.summary.fail,
    fatal: qa.summary.fatal,
    templatePath: qa.templatePath,
  });
}

const invalidRun = await runFill([
  "--company-id=profile_fixture_invalid",
  "--company-name=Profile_Fixture_invalid",
  `--data-dir=${fixtureDir}`,
  `--output-dir=${path.join(outputRoot, "invalid")}`,
  "--sector=TC",
  "--size=INVALID_SIZE",
]);
assert(invalidRun.code !== 0, "Invalid size must fail instead of falling back to 대기업.");
assert(/No exact v10\.1 JSON template/.test(invalidRun.stderr), "Invalid-size error message is not explicit.");

const summary = { cases: results, invalidSizeRejected: true };
await fs.writeFile(path.join(outputRoot, "summary.json"), `${JSON.stringify(summary, null, 2)}\n`, "utf8");
console.log(JSON.stringify(summary, null, 2));
