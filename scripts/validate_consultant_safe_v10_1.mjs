import fs from "node:fs/promises";
import path from "node:path";
import crypto from "node:crypto";
import { fileURLToPath } from "node:url";

const __dirname = path.dirname(fileURLToPath(import.meta.url));
const repoRoot = path.resolve(__dirname, "..");
const releaseDir = path.join(repoRoot, "consultant_safe_v10_1");
const templateDir = path.join(releaseDir, "templates");
const profileConfig = JSON.parse(await fs.readFile(path.join(releaseDir, "policies", "template_profiles.json"), "utf8"));
const index = JSON.parse(await fs.readFile(path.join(releaseDir, "TEMPLATE_INDEX.json"), "utf8"));

const expectedSectors = ["CG", "EM", "FB", "FN", "HC", "IF", "RR", "RT", "SV", "TC", "TR"];
const expectedSizes = ["대기업", "중견", "중소", "비상장"];
const fatal = [];
const warnings = [];
const templates = [];

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

function hashRows(rows) {
  return crypto.createHash("sha256").update(JSON.stringify(rows.map(operationalProjection))).digest("hex");
}

function validateRegexes(values, label) {
  for (const value of values ?? []) {
    try {
      new RegExp(String(value), "i");
    } catch {
      fatal.push(`${label}: invalid regex ${value}`);
    }
  }
}

for (const entry of index.templates ?? []) {
  const templatePath = path.join(releaseDir, entry.jsonTemplate);
  const template = JSON.parse(await fs.readFile(templatePath, "utf8"));
  const rows = template.rows ?? [];
  if (rows.length !== 95) fatal.push(`${entry.jsonTemplate}: expected 95 rows, found ${rows.length}.`);
  const sizePolicy = profileConfig.sizeProfiles[entry.size];
  if (!sizePolicy) fatal.push(`${entry.jsonTemplate}: missing size policy ${entry.size}.`);
  if (!profileConfig.sectorProfiles[entry.sector]) fatal.push(`${entry.jsonTemplate}: missing sector policy ${entry.sector}.`);
  for (const [indexInTemplate, row] of rows.entries()) {
    const expectedEbx = `EBX-Q-${String(indexInTemplate + 1).padStart(3, "0")}`;
    if (row.ebx !== expectedEbx) fatal.push(`${entry.jsonTemplate}: expected ${expectedEbx}, found ${row.ebx}.`);
    if (row.sizeProfile?.id !== entry.size) fatal.push(`${entry.jsonTemplate} ${row.ebx}: wrong size profile.`);
    if (row.sectorProfile?.id !== entry.sector) fatal.push(`${entry.jsonTemplate} ${row.ebx}: wrong sector profile.`);
    if (sizePolicy && Number(row.evidenceSelection?.maxEvidenceSentences) !== Number(sizePolicy.evidenceSelection.maxEvidenceSentences)) {
      fatal.push(`${entry.jsonTemplate} ${row.ebx}: evidence limit does not match size profile.`);
    }
    if (sizePolicy && Number(row.metricHints?.maxRecords) !== Number(sizePolicy.metricHints.maxRecords)) {
      fatal.push(`${entry.jsonTemplate} ${row.ebx}: metric limit does not match size profile.`);
    }
    validateRegexes(row.evidenceSelection?.topicHints, `${entry.jsonTemplate} ${row.ebx} topicHints`);
    validateRegexes(row.evidenceSelection?.negativeTopicHints, `${entry.jsonTemplate} ${row.ebx} negativeTopicHints`);
    validateRegexes(row.metricHints?.regexes, `${entry.jsonTemplate} ${row.ebx} metricHints`);
    validateRegexes(row.metricHints?.negativeRegexes, `${entry.jsonTemplate} ${row.ebx} negativeMetricHints`);
  }
  const operationalHash = hashRows(rows);
  if (entry.operationalHash && entry.operationalHash !== operationalHash) {
    fatal.push(`${entry.jsonTemplate}: index operational hash does not match template content.`);
  }
  templates.push({ ...entry, operationalHash });
}

if (templates.length !== 44) fatal.push(`Expected 44 templates, found ${templates.length}.`);
for (const sector of expectedSectors) {
  const group = templates.filter((template) => template.sector === sector);
  if (group.length !== 4) fatal.push(`${sector}: expected four templates, found ${group.length}.`);
  if (new Set(group.map((template) => template.size)).size !== expectedSizes.length) fatal.push(`${sector}: incomplete size matrix.`);
  if (new Set(group.map((template) => template.operationalHash)).size !== expectedSizes.length) fatal.push(`${sector}: operational size hashes are not distinct.`);
}
for (const size of expectedSizes) {
  const group = templates.filter((template) => template.size === size);
  if (new Set(group.map((template) => template.operationalHash)).size !== expectedSectors.length) {
    fatal.push(`${size}: sector overlays do not produce 11 distinct operational hashes.`);
  }
}

const summary = {
  releaseName: index.releaseName,
  templates: templates.length,
  rowsPerTemplate: 95,
  sectors: expectedSectors.length,
  sizes: expectedSizes.length,
  profileVersion: profileConfig.version,
  fatal,
  warnings,
};
console.log(JSON.stringify(summary, null, 2));
if (fatal.length) process.exitCode = 1;
