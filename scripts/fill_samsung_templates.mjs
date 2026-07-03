import fs from "node:fs/promises";
import path from "node:path";
import { FileBlob, SpreadsheetFile } from "@oai/artifact-tool";

const root = "F:/MY_PROJECT/template-qualitative";
const dataDir = path.join(root, "company_esg_data", "samsung_electronics_2025");
const outputDir = path.join(root, "final_template", "output", "samsung_electronics_2025");

const quantitativeTemplate = path.join(root, "final_template", "quantitative", "ESG_정량.xlsx");
const qualitativeTemplate = path.join(
  root,
  "final_template",
  "template_qualitative",
  "EBX_Q_템플릿_TC_기술·통신_대기업_consultant_safe_v2.xlsx",
);

const outputQuantitative = path.join(outputDir, "Samsung_Electronics_2025_ESG_정량_filled.xlsx");
const outputQualitative = path.join(outputDir, "Samsung_Electronics_2025_EBX_Q_qualitative_with_final_KO.xlsx");

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

function toValue(raw) {
  if (raw == null || raw === "") return null;
  const trimmed = String(raw).trim();
  if (trimmed === "-") return "-";
  if (/^>\s*\d+(\.\d+)?$/.test(trimmed)) return trimmed;
  const num = Number(trimmed.replace(/,/g, ""));
  return Number.isFinite(num) ? num : trimmed;
}

function convert(raw, mode) {
  const value = toValue(raw);
  if (value == null || typeof value === "string") return value;
  switch (mode) {
    case "krw_trillion_to_million":
      return value * 1_000_000;
    case "krw_billion_to_million":
      return value * 1_000;
    case "krw_100m_to_million":
      return value * 100;
    case "thousand_to_unit":
      return value * 1_000;
    case "gwh_to_tj":
      return value * 3.6;
    default:
      return value;
  }
}

function buildIndex(rows) {
  const byIndicator = new Map();
  for (const row of rows) byIndicator.set(row.indicator, row);
  return byIndicator;
}

function sourceText(row, sourceYear) {
  return row ? `p.${row.source_page} ${row.report} (${sourceYear} data)` : "";
}

function basisText(row, extra = "") {
  if (!row) return "";
  const parts = [row.notes, `unit in source: ${row.unit}`];
  if (extra) parts.push(extra);
  return parts.filter(Boolean).join("; ");
}

function getYearValues(row, mode = "identity") {
  return {
    y2023: convert(row?.value_2022, mode),
    y2024: convert(row?.value_2023, mode),
    y2025: convert(row?.value_2024, mode),
  };
}

function getLatestOnly(row, mode = "identity") {
  return {
    y2023: null,
    y2024: null,
    y2025: convert(row?.value_2024, mode),
  };
}

function derivedNonRenewableEnergy(rowTotal, rowRenewable) {
  const out = {};
  for (const [templateYear, sourceYear] of [
    ["y2023", "value_2022"],
    ["y2024", "value_2023"],
    ["y2025", "value_2024"],
  ]) {
    const total = toValue(rowTotal?.[sourceYear]);
    const renewable = toValue(rowRenewable?.[sourceYear]);
    out[templateYear] =
      typeof total === "number" && typeof renewable === "number" ? (total - renewable) * 3.6 : null;
  }
  return out;
}

function writeMappedRow(sheet, excelRow, values, source, basis) {
  sheet.getRange(`G${excelRow}:O${excelRow}`).values = [[
    values.y2023,
    source?.s2023 ?? "",
    basis?.b2023 ?? "",
    values.y2024,
    source?.s2024 ?? "",
    basis?.b2024 ?? "",
    values.y2025,
    source?.s2025 ?? "",
    basis?.b2025 ?? "",
  ]];
}

function writeLogRows(sheet, startRow, rows) {
  if (!rows.length) return;
  sheet.getRangeByIndexes(startRow - 1, 0, rows.length, rows[0].length).values = rows;
}

const finalAnswersKo = {
  "EBX-Q-001":
    "삼성전자는 인재와 기술을 기반으로 최고의 제품과 서비스를 창출해 인류사회에 기여한다는 경영 방향 아래 지속가능경영을 추진하고 있습니다. 2024년 보고 기준으로 이사회 산하 지속가능경영위원회와 부문별 협의체가 주요 ESG 전략과 이행 성과를 감독하며, 기후변화·수자원·자원순환 등 중대 주제별 중장기 목표를 설정했습니다. 주요 목표로는 DX부문 2030년 Scope 1·2 탄소중립, DS부문 2050년 Scope 1·2 탄소중립, DX부문 2030년 글로벌 수자원 소비량 100% 환원, DS부문 2030년 국내 제조사업장 취수량 증가 제로화, DX부문 2050년 제품 플라스틱 부품 100% 재활용 플라스틱 적용, DS부문 2030년 국내 제조사업장 폐기물 재활용률 99.9% 달성이 제시되어 있습니다.",
  "EBX-Q-002":
    "삼성전자의 지속가능경영 활동은 최고 의사결정기구인 이사회가 감독합니다. 2021년 7월 이사회 산하 지속가능경영위원회를 설치하여 환경, 사회, 지배구조 분야의 주주환원 정책과 중장기 지속가능경영 전략 등을 심의하고 있으며, 위원회는 사외이사 전원이 참여하는 구조입니다. 2025년 3월 말 기준 이사회는 사내이사 3인과 사외이사 6인으로 구성되어 있고, 이사회 의장은 대표이사와 분리되어 사외이사가 맡고 있습니다. 실무 차원에서는 DX부문 지속가능경영협의회와 DS부문 ESG경영협의회 등 부문별 회의체가 ESG 현안을 논의하고 의사결정을 지원합니다.",
  "EBX-Q-003":
    "삼성전자는 이중 중대성 평가를 통해 ESG 관련 영향, 리스크, 기회를 식별하고 관리합니다. 회사는 ESRS가 제시한 이슈와 산업 특성을 반영한 총 115개 이슈를 검토한 뒤, 회사 전략과의 연관성, 글로벌 이니셔티브, 동종업계 주요 이슈, 밸류체인 분석 결과 등을 기준으로 평가하여 핵심 지속가능경영 주제를 도출했습니다. 기후변화 영역에서는 TCFD와 CDP 기준을 참고해 리스크와 기회 풀을 구성하고, 시나리오 분석 및 내외부 이해관계자 설문, 유관부서 라운드테이블을 통해 주요 리스크와 대응 방안을 확정합니다. 식별된 리스크는 이사회, 지속가능경영위원회, 부문별 협의체 및 환경 분야 회의체를 통해 관리됩니다.",
  "EBX-Q-004":
    "삼성전자는 안전하고 건강한 근로환경 조성을 위해 부문별 최고안전보건책임자(CSO)를 중심으로 안전보건 정책과 실행 체계를 운영합니다. 각 부문은 사업장 위험요인을 점검하고 예방 활동을 추진하며, 임직원과 협력회사 안전보건 수준을 함께 관리합니다. 정량 성과로는 임직원 LTIR이 2022년 0.033%, 2023년 0.023%, 2024년 0.022%로 공시되었고, 협력회사 LTIR은 2022년 0.046%, 2023년 0.064%, 2024년 0.035%로 공시되었습니다. 중대 산업재해는 보고서상 2022년부터 2024년까지 대시 표시로 공시되어 별도 수치로 환산하지 않았습니다.",
  "EBX-Q-005":
    "삼성전자는 DX부문과 DS부문별로 최고안전보건책임자(CSO)를 두고 안전보건 관리 책임을 부여하고 있습니다. DX부문은 Global EHS실을 중심으로, DS부문은 제조·인프라 관련 조직을 중심으로 국내외 사업장의 환경안전 기준 수립과 실행을 지원합니다. 안전보건 관련 주요 사안은 부문별 회의체와 전담 조직을 통해 검토되며, 필요 시 경영진 및 상위 거버넌스 체계에 보고되어 개선 활동으로 연결됩니다.",
  "EBX-Q-006":
    "삼성전자는 사업장 안전 리스크를 사전에 식별하고 예방하기 위해 현장 점검, 위험성 평가, 교육 및 개선 조치를 운영합니다. 사업장별 유해·위험요인을 확인하고 개선 필요 사항을 도출하며, 협력회사에 대해서도 환경안전 진단, 컨설팅, 교육 등을 제공해 공급망 내 안전보건 리스크를 관리합니다. 이러한 활동은 부문별 CSO와 EHS 전담 조직이 주관하고, 정기 회의체를 통해 이행 현황을 점검하는 방식으로 운영됩니다.",
  "EBX-Q-007":
    "삼성전자는 산업재해 발생 현황과 안전 성과를 LTIR 등 지표로 관리하고 있습니다. 임직원 LTIR은 2022년 0.033%, 2023년 0.023%, 2024년 0.022%로 낮아졌으며, 협력회사 LTIR은 2022년 0.046%, 2023년 0.064%, 2024년 0.035%로 공시되었습니다. 보고서의 중대 산업재해 항목은 2022년, 2023년, 2024년 모두 대시로 표시되어 있어 이를 0으로 임의 환산하지 않고 원문 표시 기준을 유지하는 것이 적절합니다.",
  "EBX-Q-008":
    "삼성전자는 인권 기본 원칙과 관련 정책을 바탕으로 임직원과 공급망의 노동인권을 관리합니다. 이사회 산하 지속가능경영위원회, 지속가능경영협의회, ESG경영협의회 및 노동인권협의회가 노동인권 이슈를 감독하며, People팀은 인권존중 활동, 정책 수립, 인권 리스크 평가와 RBA 제3자 점검 등 실사 프로그램을 관장합니다. 또한 회사는 HRRA 등 인권 리스크 평가 절차를 통해 주요 인권 영향을 파악하고, 식별된 이슈를 정책과 개선 활동에 반영하고 있습니다.",
  "EBX-Q-009":
    "삼성전자의 노동 및 인권 관리 체계는 인권 기본 원칙, 고충처리 정책, 환경안전 방침 등 관련 정책과 전담 조직을 기반으로 운영됩니다. 노동인권협의회는 People팀, 상생협력센터, Global EHS실, 법무실, IR팀, 지속가능경영추진센터 등 관련 부서가 참여하는 협의체로서 국내외 사업장과 공급망의 노동인권 이슈를 논의합니다. 결사의 자유와 단체교섭 보장, 생활임금 격차 분석, 고충처리 채널 운영, 인권 실사 등도 관리 체계의 주요 요소입니다.",
  "EBX-Q-010":
    "삼성전자는 인권 리스크를 식별하기 위해 사업장과 공급망을 대상으로 인권 실사 및 리스크 평가를 수행합니다. People팀은 임직원 관련 일상적 인권 리스크를 관리하고, 평가 결과를 회사 정책과 표준 관행에 반영해 조직 내에 전파합니다. 식별된 이슈는 사안의 경중과 시급성에 따라 노동인권협의회, 사업위기관리 회의체, 지속가능경영협의회 또는 지속가능경영위원회에 보고되어 개선 및 대응 조치로 연결됩니다.",
  "EBX-Q-011":
    "삼성전자는 임직원 및 이해관계자의 인권·노동 관련 고충을 접수하고 처리하기 위한 다양한 채널을 운영합니다. 2024년에는 DX 및 DS부문 기준 고충 33,148건이 접수되었고, 2024년 말 기준 처리율은 98.7%로 공시되었습니다. 접수 채널 비중은 핫라인 45.1%, 온라인 29.5%, 오프라인 19.3%, 근로자 대표기구 6.1%로 나타났습니다. 또한 2024년 말 기준 글로벌 임직원의 42.7%가 단체협약의 적용을 받으며, 전 세계에 33개 노동조합과 45개 노사협의회가 운영되고 있습니다.",
  "EBX-Q-012":
    "삼성전자는 제품 안전과 품질을 고객 신뢰와 사업 지속성의 핵심 요소로 보고 품질 관리 정책과 프로세스를 운영합니다. 제품 개발, 생산, 서비스 단계에서 품질과 안전 관련 요구사항을 점검하고, 제품 책임(PL) 리스크 예방 및 대응 체계를 통해 고객에게 제공되는 제품과 서비스의 안전성을 관리합니다. 또한 고객의 소리(VOC)와 서비스 채널을 활용하여 품질 이슈를 파악하고 개선 활동에 반영합니다.",
  "EBX-Q-013":
    "삼성전자는 제품 품질과 안전 관리를 위해 품질 관련 전담 조직과 사업부별 책임 체계를 운영합니다. 관련 조직은 제품 개발·생산·서비스 단계에서 품질 기준 준수 여부를 점검하고, 품질 이슈가 발생할 경우 원인 분석과 개선 조치를 수행합니다. 고객 접점에서 접수되는 VOC와 서비스 데이터를 활용해 품질 개선 과제를 도출하고, 관련 부서와 연계하여 제품 안전 및 서비스 품질 향상을 추진합니다.",
  "EBX-Q-014":
    "삼성전자는 제품 품질 및 안전 리스크를 개발 단계부터 사후 서비스 단계까지 관리합니다. 제품 책임 리스크를 예방하기 위해 품질 검증, 안전성 점검, 고객 VOC 분석, 서비스 대응 프로세스를 운영하며, 식별된 이슈는 원인 분석과 재발 방지 조치로 연결됩니다. 이러한 체계는 제품과 서비스의 품질 문제를 조기에 발견하고 고객 피해를 최소화하기 위한 내부 관리 절차로 활용됩니다.",
  "EBX-Q-015":
    "삼성전자는 제품 품질 및 안전 관련 리스크 관리 체계, PL 예방 및 대응 절차, VOC 및 서비스 채널 운영 현황을 공시하고 있습니다. 2024년 말 기준 217개국에서 12,925개 서비스센터를 운영했으며, 2024년 서비스 교육 과정 5,940개와 교육 이수 42,249명을 공시했습니다. 또한 부정 제보 항목 내 소비자 민원 비중은 2022년 34%, 2023년 36%, 2024년 30%로 제시되어 있습니다. 다만 ESG 보고서 자체에는 제품 안전 사고 또는 리콜 건수에 대한 3개년 표가 별도로 공시되어 있지 않아, 해당 항목은 추가 내부 자료 또는 사업보고서 확인이 필요합니다.",
  "EBX-Q-016":
    "삼성전자는 개인정보 보호와 정보보안을 중요한 경영 과제로 인식하고 관련 정책과 관리 체계를 운영합니다. 회사는 개인정보 처리와 보호를 위한 기준을 마련하고, 제품·서비스 및 내부 업무 과정에서 개인정보와 정보자산을 안전하게 관리하기 위한 절차를 적용합니다. 또한 개인정보 보호 관련 내부 컨설팅, 임직원 인식 제고, 보안 점검 및 인증 활동을 통해 데이터 보호 수준을 지속적으로 관리합니다.",
  "EBX-Q-017":
    "삼성전자는 개인정보 보호와 정보보안 업무를 담당하는 전담 조직을 통해 관련 정책 수립, 내부 점검, 컨설팅 및 사고 대응 체계를 운영합니다. 개인정보와 정보보안 관련 부서는 사업부 및 관련 기능 조직과 협업하여 제품·서비스 운영 과정의 보호 조치를 검토하고, 법규 준수와 보안 리스크 관리를 지원합니다. 주요 이슈는 내부 관리 체계를 통해 검토되며 필요한 개선 조치로 연결됩니다.",
  "EBX-Q-018":
    "삼성전자는 개인정보 및 정보보안 리스크를 예방하기 위해 내부 컨설팅, 보안 점검, 인증 관리, 사고 대응 프로세스를 운영합니다. 개인정보 처리 과정과 서비스 운영 환경에서 발생할 수 있는 위험을 사전에 검토하고, 관련 법규와 내부 기준에 따라 보호 조치를 적용합니다. 또한 정부기관 정보 제공 요청 등 외부 요청에 대해서는 관계 법령에 따라 검토·대응하며, 제공 여부와 처리 현황을 관리합니다.",
  "EBX-Q-019":
    "삼성전자는 개인정보 보호와 정보보안 관리 체계, 인증 및 사고 대응 프로세스를 공시하고 있으며, 정량적으로는 사내 개인정보 컨설팅 건수와 정부기관 정보 제공 요청 대응 현황을 제시하고 있습니다. 사내 개인정보 컨설팅 건수는 2022년 5,858건, 2023년 8,302건, 2024년 8,170건이며, 정부기관 정보 제공 요청은 2022년 187건, 2023년 594건, 2024년 400건으로 공시되었습니다. 이 중 제공 건수는 각각 126건, 456건, 236건이고 제공률은 67%, 77%, 59%입니다. 다만 ESG 보고서에는 개인정보 침해 건수 또는 사이버보안 사고 건수가 명시적으로 공시되어 있지 않아, 해당 수치는 별도 자료 확인이 필요합니다.",
  "EBX-Q-020":
    "삼성전자는 환경경영을 지속가능경영의 핵심 축으로 설정하고 기후변화, 수자원, 자원순환 등 주요 환경 이슈별 중장기 목표를 운영합니다. 회사는 2030년까지 공정가스 저감과 수자원 보전 등 환경경영 과제에 총 7조 원 이상을 투자할 계획을 제시했습니다. 주요 목표로는 DX부문 2030년 Scope 1·2 탄소중립, DS부문 2050년 Scope 1·2 탄소중립, DX부문 2030년 글로벌 수자원 소비량 100% 환원, DS부문 2030년 국내 제조사업장 취수량 증가 제로화, 자원순환 관련 재활용 플라스틱 및 폐기물 재활용 목표가 포함됩니다.",
  "EBX-Q-021":
    "삼성전자의 환경경영은 이사회와 이사회 산하 지속가능경영위원회가 감독하며, 경영진과 부문별 협의체가 실행을 담당합니다. DX부문 지속가능경영협의회와 DS부문 ESG경영협의회는 환경경영 계획과 이행 성과를 검토하고, Global EHS실 및 부문별 환경 전담 조직은 사업장의 환경 기준 수립과 실행을 지원합니다. 기후변화 대응과 환경 투자 등 주요 사안은 경영진 책임 하에 추진되며, 관련 KPI와 이행 현황은 내부 회의체를 통해 관리됩니다.",
  "EBX-Q-022":
    "삼성전자는 기후변화, 수자원, 자원순환 등 환경 리스크와 영향을 식별하고 주제별 대응 전략을 운영합니다. 기후변화 영역에서는 시나리오 분석을 통해 사업, 전략, 재무계획에 미칠 수 있는 리스크와 기회를 평가하고, 온실가스 감축과 재생에너지 확대 등 대응 조치를 추진합니다. 수자원 영역에서는 지역별 수자원 리스크를 평가하고 AWS 인증 확대 등을 추진하며, 자원순환 영역에서는 원료 조달부터 생산, 사용, 폐기, 재활용까지 제품 전 과정의 환경 영향을 관리합니다.",
  "EBX-Q-023":
    "삼성전자는 환경 성과를 온실가스, 에너지, 재생에너지, 폐기물, 수자원, 오염물질 등 주요 지표로 관리하고 있습니다. Scope 1·2 온실가스 배출량(시장기준)은 2022년 15,053천 tCO2e, 2023년 13,291천 tCO2e, 2024년 14,889천 tCO2e이며, Scope 3 배출량은 2024년 105,612천 tCO2e로 공시되었습니다. 사업장 에너지 사용량은 2024년 38,772GWh, 재생에너지 사용량은 10,069GWh, 재생에너지 전환율은 31.4%입니다. 2024년 폐기물 발생량은 1,348,979톤, 폐기물 재활용률은 98%이고, 용수 취수량은 188,540천 톤, 용수 재사용량은 125,463천 톤으로 공시되었습니다. 환경 규제 위반 건수는 2022년 2건, 2023년 1건, 2024년 2건입니다.",
  "EBX-Q-024":
    "삼성전자는 준법과 윤리에 기반한 경영 원칙을 바탕으로 윤리경영과 반부패 정책을 운영합니다. 회사는 글로벌 반부패 및 뇌물방지 정책을 수립해 임직원과 거래 업체에 적용하며, 부패와 뇌물 수수에 대해서는 무관용 원칙을 적용합니다. 이러한 정책은 법과 윤리를 준수하는 기업 문화를 정착시키고 지속가능한 성장을 지원하기 위한 기준으로 활용됩니다.",
  "EBX-Q-025":
    "삼성전자는 이사회와 주요 산하 위원회, 전사 Compliance팀, 경영진단팀, 외부 독립조직인 삼성 준법감시위원회를 중심으로 준법·윤리경영 체계를 운영합니다. Compliance팀장인 준법지원인은 이사회와 경영위원회에 참석해 회사 의사결정을 지원하고 주요 사안을 이사회에 보고합니다. 삼성 준법감시위원회는 주요 관계사의 준법 감시와 통제 기능을 강화하고 제도 개선 의견을 제시하며, 회사는 윤리경영 사이트, 이메일, 전화, 팩스 등 다양한 제보 채널을 통해 부정 및 법 위반 제보를 접수합니다.",
  "EBX-Q-026":
    "삼성전자는 CPMS(Compliance Program Management System)를 활용해 부패방지, 공정거래, 지식재산권, 개인정보보호, 인권과 노사, 환경안전 등 주요 분야의 준법 리스크를 관리합니다. 국내외 모든 사업장을 대상으로 준법과 윤리 점검을 실시하고, 점검 결과는 연 1회 이상 이사회에 보고되어 개선 활동에 반영됩니다. 또한 대외 후원금 사용, 신규 업체 등록 및 계약 체결 과정에서 반부패 검토 절차를 운영하여 거래 초기 단계부터 부패 가능성을 관리합니다.",
  "EBX-Q-027":
    "삼성전자는 준법·윤리 관련 교육, 제보 채널, 점검 및 조사 절차를 운영하고 있으며 관련 정량 지표를 일부 공시하고 있습니다. 컴플라이언스 교육 참여자는 2022년 126,867명, 2023년 138,742명, 2024년 138,414명이고, 부정 예방 교육 참여자는 2022년 254,045명, 2023년 254,511명, 2024년 254,003명입니다. 컴플라이언스 제보는 2022년 1,098건, 2023년 1,400건, 2024년 1,238건이며, 부정 제보 건수는 2022년 999건, 2023년 892건, 2024년 930건으로 공시되었습니다. 다만 ESG 보고서 자체에는 제재, 소송, 벌금 또는 징계 조치에 대한 3개년 상세 표가 포함되어 있지 않고, 관련 법적 제재 정보는 사업보고서의 별도 페이지를 참조하도록 안내되어 있어 추가 확인이 필요합니다.",
};

async function fillQuantitative() {
  const csv = parseCsv(await fs.readFile(path.join(dataDir, "data_dinh_luong.csv"), "utf8"));
  const byIndicator = buildIndex(csv);
  const workbook = await SpreadsheetFile.importXlsx(await FileBlob.load(quantitativeTemplate));
  const sheet = workbook.worksheets.getItem("정량 데이터");
  sheet.getRange("G2:O211").values = Array.from({ length: 210 }, () => Array(9).fill(null));

  const mappings = [
    [597, "Female employee ratio", "identity", "Directly mapped to female employee ratio."],
    [619, "Employees with disabilities", "identity", "Domestic employees; source includes subsidiary-type standard workplace employees from 2023."],
    [645, "Welfare expenses", "krw_billion_to_million", "Converted KRW billion to KRW million."],
    [670, "Serious industrial accidents", "identity", "Source discloses dash markers; retained as source marker, not normalized to zero."],
    [673, "LTIR", "identity", "Mapped to lost-time injury rate disclosed by Samsung."],
    [685, "Grievance cases received", "identity", "2024 only disclosed; placed in 2025 reporting-year column."],
    [689, "Total training expense", "krw_100m_to_million", "Converted KRW 100 million to KRW million."],
    [742, "Annual recycled plastic use", "identity", "Mapped recycled plastic use to recycled/renewable raw material use."],
    [781, "NOx emissions", "identity", "Direct air pollutant mapping."],
    [782, "SOx emissions", "identity", "Direct air pollutant mapping."],
    [784, "Volatile organic compound emissions", "identity", "Direct VOC mapping."],
    [751, "Scope 1 direct emissions", "thousand_to_unit", "Converted 1000 tCO2e to tCO2e."],
    [752, "Scope 2 indirect emissions, market-based", "thousand_to_unit", "Converted 1000 tCO2e to tCO2e; market-based."],
    [762, "Electricity use", "gwh_to_tj", "Converted GWh to TJ."],
    [767, "General waste generated", "identity", "Direct waste mapping."],
    [768, "Hazardous waste generated", "identity", "Direct waste mapping."],
    [769, "Waste recycled", "identity", "Direct waste mapping."],
    [789, "BOD emissions", "identity", "Direct water pollutant mapping."],
    [790, "SS emissions", "identity", "Direct water pollutant mapping."],
    [795, "Municipal/surface water withdrawal", "thousand_to_unit", "Converted 1000 tons to tons."],
    [796, "Groundwater withdrawal", "thousand_to_unit", "Converted 1000 tons to tons."],
    [799, "Water reuse", "thousand_to_unit", "Converted 1000 tons to tons."],
    [809, "Revenue", "krw_trillion_to_million", "Converted KRW trillion to KRW million."],
    [812, "Personnel expenses", "krw_trillion_to_million", "Converted KRW trillion to KRW million."],
    [813, "Purchase cost to suppliers", "krw_trillion_to_million", "Converted KRW trillion to KRW million."],
    [814, "Social contribution cost", "krw_trillion_to_million", "Converted KRW trillion to KRW million."],
    [815, "Taxes and public dues", "krw_trillion_to_million", "Converted KRW trillion to KRW million."],
    [831, "Anti-fraud training participants", "identity", "Mapped to ethics/anti-fraud training completion count."],
  ];

  const rawValues = sheet.getRange("A1:O211").values;
  const rowByTemplateId = new Map(rawValues.slice(1).map((row, i) => [String(row[0]), i + 2]));
  const log = [
    ["template_id", "excel_row", "source_indicator", "status", "conversion_or_note"],
  ];

  for (const [templateId, indicator, mode, note] of mappings) {
    const row = byIndicator.get(indicator);
    const excelRow = rowByTemplateId.get(String(templateId));
    if (!row || !excelRow) {
      log.push([templateId, excelRow ?? "", indicator, "not_written", row ? "template row not found" : "source indicator not found"]);
      continue;
    }
    const values = row.value_2022 || row.value_2023 ? getYearValues(row, mode) : getLatestOnly(row, mode);
    writeMappedRow(
      sheet,
      excelRow,
      values,
      {
        s2023: values.y2023 == null ? "" : sourceText(row, "2022"),
        s2024: values.y2024 == null ? "" : sourceText(row, "2023"),
        s2025: values.y2025 == null ? "" : sourceText(row, "2024"),
      },
      {
        b2023: values.y2023 == null ? "" : basisText(row, note),
        b2024: values.y2024 == null ? "" : basisText(row, note),
        b2025: values.y2025 == null ? "" : basisText(row, note),
      },
    );
    log.push([templateId, excelRow, indicator, "written", note]);
  }

  const totalEnergy = byIndicator.get("Facility energy use");
  const renewableEnergy = byIndicator.get("Renewable energy use");
  const nonRenewableValues = derivedNonRenewableEnergy(totalEnergy, renewableEnergy);
  const energyRow = rowByTemplateId.get("756");
  if (energyRow) {
    const note = "Derived as (facility energy use - renewable energy use), converted from GWh to TJ.";
    writeMappedRow(
      sheet,
      energyRow,
      nonRenewableValues,
      {
        s2023: sourceText(totalEnergy, "2022") + "; " + sourceText(renewableEnergy, "2022"),
        s2024: sourceText(totalEnergy, "2023") + "; " + sourceText(renewableEnergy, "2023"),
        s2025: sourceText(totalEnergy, "2024") + "; " + sourceText(renewableEnergy, "2024"),
      },
      { b2023: note, b2024: note, b2025: note },
    );
    log.push([756, energyRow, "Facility energy use - Renewable energy use", "written", note]);
  }

  const boardRows = [
    [845, 3, "Board internal directors", "From qualitative evidence p.6: board comprised 3 internal directors as of 2025-03."],
    [846, 6, "Board outside directors", "From qualitative evidence p.6: board comprised 6 outside directors as of 2025-03."],
  ];
  for (const [templateId, value, indicator, note] of boardRows) {
    const excelRow = rowByTemplateId.get(String(templateId));
    if (!excelRow) continue;
    writeMappedRow(
      sheet,
      excelRow,
      { y2023: null, y2024: null, y2025: value },
      { s2023: "", s2024: "", s2025: "p.6 Samsung Electronics Sustainability Report 2025" },
      { b2023: "", b2024: "", b2025: note },
    );
    log.push([templateId, excelRow, indicator, "written", note]);
  }

  const logSheet = workbook.worksheets.add("Samsung_mapping_log");
  writeLogRows(logSheet, 1, log);
  logSheet.getRange("A1:E1").format = { font: { bold: true } };
  logSheet.getRange("A:E").format.columnWidthPx = 220;

  await fs.mkdir(outputDir, { recursive: true });
  const exported = await SpreadsheetFile.exportXlsx(workbook);
  await exported.save(outputQuantitative);
  return { workbook, output: outputQuantitative, written: log.filter((r) => r[3] === "written").length };
}

async function fillQualitative() {
  const csv = parseCsv(await fs.readFile(path.join(dataDir, "data_dinh_tinh.csv"), "utf8"));
  const byEbx = new Map(csv.map((row) => [row.ebx, row]));
  const workbook = await SpreadsheetFile.importXlsx(await FileBlob.load(qualitativeTemplate));
  const sheet = workbook.worksheets.getItem("EBX-Q 템플릿");
  const headers = [
    "Samsung company",
    "Samsung coverage",
    "Samsung source pages",
    "Evidence type",
    "Samsung evidence / filled slots",
    "Quantitative support",
    "Gap or reviewer note",
    "Source PDF",
    "Samsung final response (KO)",
  ];
  sheet.getRange("O1:W1").values = [headers];
  sheet.getRange("O1:W1").format = { font: { bold: true }, fill: "#D9EAF7" };
  sheet.getRange("O:W").format.columnWidthPx = 230;

  const templateRows = sheet.getRange("A2:A28").values.flat();
  for (let i = 0; i < templateRows.length; i += 1) {
    const ebx = templateRows[i];
    const row = byEbx.get(ebx);
    if (!row) continue;
    const excelRow = i + 2;
    sheet.getRange(`O${excelRow}:W${excelRow}`).values = [[
      "Samsung Electronics Co., Ltd.",
      row.coverage_status,
      row.source_pages,
      row.evidence_type,
      row.original_text_ko,
      row.quantitative_support,
      row.gap_or_note,
      row.source_pdf,
      finalAnswersKo[ebx] ?? "",
    ]];
  }

  const sourceSheet = workbook.worksheets.add("Samsung_qual_source");
  const sourceHeaders = Object.keys(csv[0]);
  writeLogRows(sourceSheet, 1, [sourceHeaders, ...csv.map((row) => sourceHeaders.map((h) => row[h]))]);
  sourceSheet.getRange("A1:I1").format = { font: { bold: true } };
  sourceSheet.getRange("A:I").format.columnWidthPx = 200;

  await fs.mkdir(outputDir, { recursive: true });
  const exported = await SpreadsheetFile.exportXlsx(workbook);
  await exported.save(outputQualitative);
  return { workbook, output: outputQualitative, written: csv.length };
}

async function verifyWorkbook(workbook, label) {
  const errors = await workbook.inspect({
    kind: "match",
    searchTerm: "#REF!|#DIV/0!|#VALUE!|#NAME\\?|#N/A",
    options: { useRegex: true, maxResults: 100 },
    summary: `${label} formula error scan`,
  });
  return errors.ndjson;
}

const quantitative = await fillQuantitative();
const qualitative = await fillQualitative();
const quantitativeErrors = await verifyWorkbook(quantitative.workbook, "quantitative");
const qualitativeErrors = await verifyWorkbook(qualitative.workbook, "qualitative");

const summary = {
  outputDir,
  quantitative: {
    output: quantitative.output,
    mapped_rows_written: quantitative.written,
    formula_errors: quantitativeErrors,
  },
  qualitative: {
    output: qualitative.output,
    ebx_rows_written: qualitative.written,
    formula_errors: qualitativeErrors,
  },
};

await fs.writeFile(path.join(outputDir, "run_summary.json"), JSON.stringify(summary, null, 2), "utf8");
console.log(JSON.stringify(summary, null, 2));
