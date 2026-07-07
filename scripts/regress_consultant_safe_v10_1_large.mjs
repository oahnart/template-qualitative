import fs from "node:fs/promises";
import path from "node:path";
import { spawn } from "node:child_process";
import { fileURLToPath } from "node:url";

const __dirname = path.dirname(fileURLToPath(import.meta.url));
const repoRoot = path.resolve(__dirname, "..");
const outputRoot = path.join(repoRoot, ".tmp", "consultant_safe_v10_1_large_regression");
const warningBaselines = { RR: 63, FB: 69, IF: 48, RT: 44, TR: 65, EM: 49, TC: 36, CG: 61, SV: 66, HC: 51, FN: 69 };
const companies = [
  { sector: "RR", id: "doosan_fuel_cell_large_RR_renewable_resources_alternative_energy_2025", name: "Doosan_Fuel_Cell_2025" },
  { sector: "FB", id: "hitejinro_large_FB_food_beverage_2025", name: "HiteJinro_2025" },
  { sector: "IF", id: "hyundai_eandc_large_IF_infrastructure_2025", name: "Hyundai_EandC_2025" },
  { sector: "RT", id: "hyundai_steel_large_RT_resource_processing_manufacturing_2025", name: "Hyundai_Steel_2025" },
  { sector: "TR", id: "korean_air_large_TR_transportation_2025", name: "Korean_Air_2025" },
  { sector: "EM", id: "korea_zinc_large_EM_extractives_minerals_2025", name: "Korea_Zinc_2025" },
  { sector: "TC", id: "lg_electronics_large_Technology_Communications", name: "LG_Electronics_2025" },
  { sector: "CG", id: "lg_household_health_care_large_CG_consumer_goods_2025", name: "LG_Household_Health_Care_2025" },
  { sector: "SV", id: "naver_large_SV_services_2025", name: "NAVER_2025" },
  { sector: "HC", id: "samsung_biologics_large_HC_healthcare_2025", name: "Samsung_Biologics_2025" },
  { sector: "FN", id: "shinhan_financial_group_large_FN_finance_2025", name: "Shinhan_Financial_Group_2025" },
];

function runFill(company, outputDir) {
  const args = [
    path.join(repoRoot, "scripts", "fill_company_qualitative_v10_1.mjs"),
    `--company-id=${company.id}`,
    `--company-name=${company.name}`,
    `--data-dir=${path.join(repoRoot, "company_esg_data", company.id)}`,
    `--output-dir=${outputDir}`,
    `--sector=${company.sector}`,
    "--size=\uB300\uAE30\uC5C5",
  ];
  return new Promise((resolve) => {
    const child = spawn(process.execPath, args, { cwd: repoRoot, stdio: ["ignore", "pipe", "pipe"] });
    let stdout = "";
    let stderr = "";
    child.stdout.on("data", (chunk) => { stdout += chunk; });
    child.stderr.on("data", (chunk) => { stderr += chunk; });
    child.on("close", (code) => resolve({ code, stdout, stderr }));
  });
}

await fs.rm(outputRoot, { recursive: true, force: true });
await fs.mkdir(outputRoot, { recursive: true });
const results = [];
const failures = [];
for (const company of companies) {
  const outputDir = path.join(outputRoot, company.sector);
  await fs.mkdir(outputDir, { recursive: true });
  const run = await runFill(company, outputDir);
  if (run.code !== 0) failures.push(`${company.sector}: fill exited ${run.code}: ${run.stderr || run.stdout}`);
  const qaFile = (await fs.readdir(outputDir)).find((file) => file.endsWith("_QA.json"));
  if (!qaFile) {
    failures.push(`${company.sector}: QA file not generated.`);
    continue;
  }
  const qa = JSON.parse(await fs.readFile(path.join(outputDir, qaFile), "utf8"));
  if (qa.rowCount !== 95) failures.push(`${company.sector}: expected 95 rows.`);
  if (qa.summary.fail !== 0 || qa.summary.fatal !== 0) failures.push(`${company.sector}: fail=${qa.summary.fail}, fatal=${qa.summary.fatal}.`);
  if (qa.summary.topicMismatchedMetricRows !== 0) failures.push(`${company.sector}: topic-mismatched metric detected.`);
  if (qa.summary.technicalMetricWordingRows !== 0) failures.push(`${company.sector}: technical metric wording detected.`);
  if (qa.summary.controlCharacterRows !== 0) failures.push(`${company.sector}: control character detected.`);
  if (qa.summary.malformedMetricRows !== 0) failures.push(`${company.sector}: malformed metric detected.`);
  if (qa.summary.duplicateFinalAnswerGroups !== 0) failures.push(`${company.sector}: duplicate final answer detected.`);
  if (qa.summary.warnings > warningBaselines[company.sector]) failures.push(`${company.sector}: warning regression (${qa.summary.warnings} > ${warningBaselines[company.sector]}).`);
  if (!qa.templatePath.includes(`${company.sector}_\uB300\uAE30\uC5C5_consultant_safe_v10_1.json`)) failures.push(`${company.sector}: wrong template selected.`);
  if (company.sector === "TC" && qa.summary.genericTopUpRows > 2) failures.push(`TC: generic top-up regression.`);
  if (qa.summary.repeatedSentenceMax > 4) failures.push(`${company.sector}: repeated sentence regression.`);
  results.push({
    sector: company.sector,
    companyId: company.id,
    pass: qa.summary.pass,
    warn: qa.summary.warn,
    fail: qa.summary.fail,
    fatal: qa.summary.fatal,
    warnings: qa.summary.warnings,
    controlCharacterRows: qa.summary.controlCharacterRows,
    malformedMetricRows: qa.summary.malformedMetricRows,
    duplicateFinalAnswerGroups: qa.summary.duplicateFinalAnswerGroups,
    sufficientWithoutEvidenceRows: qa.summary.sufficientWithoutEvidenceRows,
    repeatedSentenceMax: qa.summary.repeatedSentenceMax,
    genericTopUpRows: qa.summary.genericTopUpRows,
  });
  console.log(`${company.sector}: pass=${qa.summary.pass} warn=${qa.summary.warn} fatal=${qa.summary.fatal}`);
}

const summary = { companies: results.length, results, failures };
await fs.writeFile(path.join(outputRoot, "summary.json"), `${JSON.stringify(summary, null, 2)}\n`, "utf8");
console.log(JSON.stringify(summary, null, 2));
if (failures.length) process.exitCode = 1;
