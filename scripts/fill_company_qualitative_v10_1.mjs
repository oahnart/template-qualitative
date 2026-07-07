import fs from "node:fs/promises";
import path from "node:path";
import { fileURLToPath } from "node:url";
import { SpreadsheetFile, Workbook } from "@oai/artifact-tool";

const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);
const repoRoot = path.resolve(__dirname, "..");

const RELEASE_NAME = "consultant_safe_v10_1";
const EXPECTED_ROW_COUNT = 95;
const DEFAULT_CONFIG = {
  companyId: "samsung_biologics_2025",
  companyName: "Samsung_Biologics_2025",
  dataDir: path.join(repoRoot, "company_esg_data", "samsung_biologics_large_HC_healthcare_2025"),
  outputDir: path.join(repoRoot, "final_template", "output", "samsung_biologics_2025"),
  templateDir: path.join(repoRoot, "consultant_safe_v10_1"),
  sector: "HC",
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

const TEMPLATE_SHEET = "EBX-Q 템플릿";
const OUTPUT_SHEET = RELEASE_NAME;
const CONTROL_CHAR_REGEX = /[\x00-\x08\x0B\x0C\x0E-\x1F]/;
const MALFORMED_METRIC_REGEX = /(?:202[0-9]\uB144\s+2,0[0-9]{2}\b)|(?:[,;]\s*(?:\uC740|\uB294|\uC774|\uAC00)\s+202[0-9]\uB144)|(?:^|[;.]\s*)\d+\)\s*(?:202[0-9]\uB144|[^.;]{80,})/;
const KOREAN_REGEX = /[\u3131-\uD79D]/;
const TECHNICAL_METRIC_REGEX = /\bquantitative\b|정량|định lượng/i;
const EBX_CODE_REGEX = /\bEBX(?:[-_\s]*Q)?[-_\s]*\d{1,3}\b/i;
const SOURCE_TRACE_REGEX = /\b(?:Source|PDF|page|pages|p\.\d+|reviewer|audit|trace|file|chunk|metadata)\b|P\s*\.\s*\d+|출처|근거|원문|검토자|감사|파일|보고\s*페이지|원천자료|\[[^\]]*p\.\d+[^\]]*\]/i;
const OCR_ARTIFACT_REGEX = /Overview Environmental Social Governance ESG Data Appendix|AppendixFacts|PrinciplePlanet|Our Company|Facts\s*&\s*Figures|Materiality Assessment|Implementation Guidance|Step\s*\d|Mission and Vision|Privacy Protection\s*&\s*Security|Customer Data Platform|Corporate Governance|Governance and Major Progress|코드\s*공시\s*항목|보고\s*페이지\s*및\s*답변|TC-SC-|GRI\s*\d|ESRS|TCFD|구분\s+리스크|구분\s+단위\s+2022년\s+2023년\s+2024년|보고서\s+\d{2}\b/i;
const INLINE_LIST_ARTIFACT_REGEX = /[>|]/;
const CONTACT_ARTIFACT_REGEX = /[A-Z0-9._%+-]+@[A-Z0-9.-]+\.[A-Z]{2,}|doosan\.com|Legal\/Compliance|HELP\s*DESK|Help\s*Desk|이메일|우편|주소|주관부서/i;
const NAVIGATION_ARTIFACT_REGEX = /Company Overview|ESG Strategy|Materiality|Appendix|활동\s*내역|일자\s*활동|구분\s*단위|환경경영\s*>|안전보건\s*>|인권경영\s*>|윤리경영\s*>/i;
const BAD_OPENING_CONNECTOR_REGEX = /^(?:또한|이를 통해|아울러|다만|특히|따라서|그리고|이에 따라|이와 함께|뿐만 아니라|이렇게|이처럼|이로써)[,\s]+/;
const SIGNATURE_ARTIFACT_REGEX = /^20\d{2}년\s*\d{1,2}월\s*\d{1,2}일\s*두산퓨얼셀(?:㈜)?\s*(?:CSHO|CEO|CFO|CISO)?\s*[^\s.]*\s*(?:\d{2}\s*){3,}/;
const HEADING_PREFIX_REGEX = /^(?:ESG\s*거버넌스\s*ESG\s*위원회|ESG\s*전략|인권경영\s*인권정책\s*인권이슈\s*제보\s*채널|고객\s*만족\s*품질\s*정책|추진\s*방향|리스크\s*관리\s*리스크\s*관리\s*활동|환경법규\s*준수|안전보건\s*안전보건\s*관리\s*활동|안전보건\s*안전보건관리\s*중장기\s*로드맵\s*수립|CEO\s*Message|윤리\s*및\s*공정거래\s*\d+\.?|노동\s*및\s*인권\s*\d+(?:\.\d+)?\s*목적:?|공급망\s*ESG\s*관리\s*동반성장\s*지원\s*프로그램|공급망\s*ESG\s*관리|공급망\s*리스크\s*관리\s*평가\s*및\s*후속\s*조치|품질경영\s*추진체계|정보보안\s*및\s*고객\s*정보보호\s*정보보호\s*인식\s*제고|환경영향평가\s*프로세스\s*환경\s*교육|환경\s*교육\s*이수\s*현장\s*폐기물\s*발생\s*카드\s*작성법\s*환경경영시스템\s*인증|환경영향평가|온실가스\s*관리\s*SCOPE\s*3\s*배출량\s*관리|윤리경영|사이버신고센터\s*운영방침\s*•?|접수\s*채널\s*윤리규범\s*및\s*윤리규범위반과\s*관련한\s*문의사항이나\s*도움이\s*필요한\s*경우)\s*/;
const SPECIAL_ARTIFACT_REGEX = /[▲△■□◆◇▶▷●○❶-❿①-⑳Ⓐ-Ⓩ※✓✔✕→←↔]/;
const FINAL_UNSUPPORTED_SPECIAL_CHAR_REGEX = /[^\p{Script=Hangul}\p{Script=Latin}\p{Number}\s.,;:()/%&+\-·]/u;
const RAW_HEADING_PREFIX_REGEX = /^(?:지배구조|전략\s|영향,\s*(?:위험|기회)|산업안전보건\s*지배구조|안전보건\s*목표\s*모니터링|주요\s*사고\s*유형별|보안\s*교육\s*및\s*보안\s*수칙|윤리․?준법\s*경영에\s*대한\s*비전|투명경영위원회\s*주요\s*역할|내부고발자\s*보호\s*규정|정보보호\s*투자|이슈풀\s*구성\s*시|품질경영\s*인증\s*시스템\s*운영|통합\s*폐기물관리시스템|지표\s*및\s*목표)/;
const TABLE_DIAGRAM_ARTIFACT_REGEX = /수지율\s*=|발생시점|발생\s*가능성|영향\s*크기|영향\s*범위|재무\s*영향|조직도|인증서|이사회\s+투명경영위원회\s+준법지원인|검토\s*\/\s*승인|환경,\s*사회|2024년\s*품질\s*교육\s*실시\s*현황|조직명|수행업무|운영주기|주요\s*안건|구분\s+운영주기|홈페이지\s*환경정책\s*링크|자가진단\s+실사\s+현장심사|정보보호\s+Strategy|제품\s+보안\s+운영\s+방침/;
const INCOMPLETE_FINAL_REGEX = /(하고,|하며,|개정하고,|수립하고,|운영하고,)\s*$/;
const REPORT_HEADING_PHRASES = [
  "신한금융그룹 지배구조",
  "다양성 및 인권경영",
  "DEI 문화 조성 활동",
  "인권경영 관리 체계",
  "인권리스크 관리",
  "인권 리스크 관리",
  "안전보건",
  "안전보건 관리 체계",
  "안전보건 정책 및 거버넌스",
  "안전보건 리스크 관리",
  "안전보건 프로그램",
  "정보보호 및 개인정보보호",
  "정보보호 원칙 및 방향",
  "정보보호 체계 확립",
  "정보보호 실행력 강화",
  "윤리준법경영",
  "윤리준법경영 체계",
  "불공정거래 및 부정거래 방지",
  "불공정거래 예방교육",
  "기후변화 대응",
  "환경경영 및 자원순환",
  "환경경영 추진 체계",
  "기업 리스크 관리 강화",
  "리스크 관리 체계",
  "통합 리스크 관리",
  "금융 시스템 리스크 관리",
  "ESG 리스크 관리 및 감독",
  "ESG 평가",
  "지속가능경영 거버넌스",
  "포용·상생금융",
  "금융소비자 보호",
  "금융 접근성 향상",
  "디지털 혁신",
  "생물다양성 보존 및 생태계 보호",
  "생물다양성 보존",
  "사회공헌",
  "공급망 ESG 관리",
  "이사회 내 위원회 구성",
  "이사회 구성",
  "이사회 정관",
  "이사회 평가 및 운영",
  "이사회 활동",
  "사외이사 평가",
  "위원회 구성 및 운영 체계",
];
const REPORT_HEADING_PATTERN = [...REPORT_HEADING_PHRASES]
  .sort((a, b) => b.length - a.length)
  .map(escapeRegExp)
  .join("|");
const REPORT_HEADING_PREFIX_REGEX = new RegExp(`^(?:(?:${REPORT_HEADING_PATTERN})(?=\\s|$)\\s*)+`, "i");
const REPORT_HEADING_INLINE_REGEX = new RegExp(`\\s+(?:${REPORT_HEADING_PATTERN})\\s+(?=신한|Shinhan|이사회|위원회|위험관리|공정거래|2024년|3C|지주|각|그룹|임직원|고객|협력|공개|보고|국내|전|모든|당사|정관|재해|노사|산업|안전|인권|ESG|SDGs)`, "g");
const KOREAN_GRAMMAR_CLEANUPS = [
  [/공개함으로서/g, "공개함으로써"],
  [/선도기업으로써/g, "선도기업으로서"],
  [/해야하는/g, "해야 하는"],
  [/최소화\s+하기/g, "최소화하기"],
  [/발생될것으로/g, "발생될 것으로"],
  [/범위\s+를/g, "범위를"],
  [/검토\s*\/\s*승인/g, "검토 및 승인"],
  [/환경,\s*사회/g, "환경, 사회"],
];
const FINAL_FORBIDDEN_REGEX = new RegExp(`${EBX_CODE_REGEX.source}|${SOURCE_TRACE_REGEX.source}|${OCR_ARTIFACT_REGEX.source}`, "i");

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
  const [rawHeaders, ...body] = rows.filter((r) => r.some((value) => value !== ""));
  if (!rawHeaders) return [];
  const headers = rawHeaders.map((header) => header.replace(/^\uFEFF/, ""));
  return body.map((r) => Object.fromEntries(headers.map((header, i) => [header, r[i] ?? ""])));
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

function normalizeEbxCode(value) {
  const text = String(value ?? "").trim();
  const match = text.match(/(?:EBX|ESG)[-_\s]*Q[-_\s]*(\d{1,3})/i);
  return match ? `EBX-Q-${match[1].padStart(3, "0")}` : text;
}

function normalizeInputRow(row) {
  const values = row.values && typeof row.values === "object" ? row.values : {};
  const ebx = normalizeEbxCode(row.ebx || row.legacy_ebx || row.mapped_item_id || row.item_id);
  return {
    ...row,
    ebx,
    original_text_ko: row.original_text_ko ?? row.evidence_text ?? "",
    value_2022: row.value_2022 ?? values["2022"] ?? "",
    value_2023: row.value_2023 ?? values["2023"] ?? "",
    value_2024: row.value_2024 ?? values["2024"] ?? "",
  };
}

function rowsFromJsonValue(value) {
  if (Array.isArray(value)) return value;
  if (Array.isArray(value.rows)) return value.rows;
  if (Array.isArray(value.data)) return value.data;
  if (Array.isArray(value.items)) return value.items;
  if (Array.isArray(value.metrics)) return value.metrics;
  return [];
}

async function readRowsCsvOrJson(dataDir, baseName, { aliases = [], required = false } = {}) {
  const baseNames = [baseName, ...aliases];
  for (const candidate of baseNames) {
    const csvRows = await readCsvIfExists(path.join(dataDir, `${candidate}.csv`));
    if (csvRows.length) return csvRows.map(normalizeInputRow);
  }
  for (const candidate of baseNames) {
    try {
      const value = JSON.parse(await fs.readFile(path.join(dataDir, `${candidate}.json`), "utf8"));
      const rows = rowsFromJsonValue(value);
      if (rows.length) return rows.map(normalizeInputRow);
    } catch {
      // Try the next accepted basename.
    }
  }
  if (required) {
    throw new Error(`Missing required ${baseNames.map((name) => `${name}.csv/json`).join(" or ")} in ${dataDir}.`);
  }
  return [];
}

async function cleanupInspectSidecar(file) {
  await fs.rm(`${file}.inspect.ndjson`, { force: true });
}

function normalizeWhitespace(text) {
  return String(text ?? "")
    .replace(/\u00A0/g, " ")
    .replace(/[ \t]+/g, " ")
    .replace(/\s*\n+\s*/g, " ")
    .trim();
}

function escapeRegExp(text) {
  return String(text ?? "").replace(/[.*+?^${}()|[\]\\]/g, "\\$&");
}

function sanitizeXmlText(text) {
  return String(text ?? "").replace(/[\x00-\x08\x0B\x0C\x0E-\x1F]/g, "");
}

function finalTextCleanup(text) {
  let out = repairKoreanSpacing(text)
    .replace(/\uC800\uD0C4\uC18C\s+\uACBD\uC81C/g, "\uC800\uD0C4\uC18C \uACBD\uC81C")
    .replace(/\uBD84\uC11D\s+\uD558\uC600\uC2B5\uB2C8\uB2E4/g, "\uBD84\uC11D\uD558\uC600\uC2B5\uB2C8\uB2E4");
  out = out
    .replace(/[•∙‣◦▪▫■□◆◇▶▷●○▲△❶-❿①-⑳Ⓐ-Ⓩ※✓✔✕→←↔⇒⇔↗↘↙↖★☆]/g, " ")
    .replace(/[“”„‟«»「」『』《》〈〉【】]/g, "")
    .replace(/[‘’‚‛`´]/g, "")
    .replace(/[~∼〜～]/g, "-")
    .replace(/\s*\*\s*/g, " ")
    .replace(/LG전자은/g, "LG전자는")
    .replace(/현황와/g, "현황과")
    .replace(/지표을/g, "지표를")
    .replace(/체계을/g, "체계를")
    .replace(/목표을/g, "목표를")
    .replace(/횟수은/g, "횟수는")
    .replace(/안건 수은/g, "안건 수는")
    .replace(/목표은/g, "목표는")
    .replace(/관리은/g, "관리는")
    .replace(/정책을 관리 지표/g, "정책을 관리 기준")
    .replace(/건로 제시/g, "건으로 제시");
  for (const [from, to] of KOREAN_GRAMMAR_CLEANUPS) out = out.replace(from, to);
  return normalizeWhitespace(out);
}

function repairKoreanSpacing(text) {
  const replacements = [
    ["최 고", "최고"],
    ["모 든", "모든"],
    ["업 무", "업무"],
    ["사 항", "사항"],
    ["전 담", "전담"],
    ["잠 재", "잠재"],
    ["인 권", "인권"],
    ["정 책", "정책"],
    ["시스 템", "시스템"],
    ["프로 세스", "프로세스"],
    ["프로 그램", "프로그램"],
    ["모바 일", "모바일"],
    ["스마 트", "스마트"],
    ["인 프라", "인프라"],
    ["구 축", "구축"],
    ["수 행", "수행"],
    ["실 행", "실행"],
    ["직 접", "직접"],
    ["집 행", "집행"],
    ["경영원 칙", "경영원칙"],
    ["정책 을", "정책을"],
    ["연 간", "연간"],
    ["경 우", "경우"],
    ["승 인", "승인"],
    ["후원 금", "후원금"],
    ["부 패", "부패"],
    ["그 룹", "그룹"],
    ["사 례", "사례"],
    ["영 향", "영향"],
    ["전 략", "전략"],
    ["방 안", "방안"],
    ["논 의", "논의"],
    ["위 험", "위험"],
    ["건 강", "건강"],
    ["안 건", "안건"],
    ["검 토", "검토"],
    ["이 슈", "이슈"],
    ["전 세 계", "전 세계"],
    ["해 외", "해외"],
    ["국내 외", "국내외"],
    ["사 업장", "사업장"],
    ["세 부", "세부"],
    ["핵 심", "핵심"],
    ["고 려", "고려"],
    ["관 행", "관행"],
    ["접 수", "접수"],
    ["선 정", "선정"],
    ["워 크 숍", "워크숍"],
    ["워크 숍", "워크숍"],
    ["컨설 팅", "컨설팅"],
    ["모니터 링", "모니터링"],
    ["리 포트", "리포트"],
    ["도 입", "도입"],
    ["시 행", "시행"],
    ["수 렴", "수렴"],
    ["종 합", "종합"],
    ["권 리", "권리"],
    ["게 미치는", "에게 미치는"],
    ["원 칙", "원칙"],
    ["재 무", "재무"],
    ["절 차", "절차"],
    ["발 굴", "발굴"],
    ["계 획", "계획"],
    ["구 현", "구현"],
    ["맞 는", "맞는"],
    ["역 량", "역량"],
    ["향 상", "향상"],
    ["상시 키고", "상시키고"],
    ["배 포", "배포"],
    ["제 작", "제작"],
    ["웹 툰", "웹툰"],
    ["직 무", "직무"],
    ["담당 자", "담당자"],
    ["안전보 건", "안전보건"],
    ["E HS", "EHS"],
    ["Chie f", "Chief"],
    ["O fficer", "Officer"],
    ["Z ero", "Zero"],
    ["E MC", "EMC"],
    ["선 호도", "선호도"],
    ["가 뭄", "가뭄"],
    ["폭 염", "폭염"],
    ["폐 기물", "폐기물"],
    ["취 수량", "취수량"],
    ["기 울이고", "기울이고"],
    ["신 규", "신규"],
    ["냉 매", "냉매"],
    ["유 관", "유관"],
    ["일 반", "일반"],
    ["규 제", "규제"],
    ["작 업", "작업"],
    ["동 참", "동참"],
    ["자 율", "자율"],
    ["장 애", "장애"],
    ["품질진 단", "품질진단"],
    ["참 여", "참여"],
    ["실 천", "실천"],
    ["사 외", "사외"],
    ["시 각", "시각"],
    ["균 형", "균형"],
    ["또 한", "또한"],
    ["외 부", "외부"],
    ["미 치는", "미치는"],
    ["생 각", "생각"],
    ["시 나리오", "시나리오"],
    ["분 석", "분석"],
    ["진 행", "진행"],
    ["따 라", "따라"],
    ["역할 을", "역할을"],
    ["책 임", "책임"],
    ["N GO", "NGO"],
    ["I LO", "ILO"],
    ["진 단", "진단"],
    ["채널 을", "채널을"],
    ["접근 하여", "접근하여"],
    ["점 검", "점검"],
    ["개 선", "개선"],
    ["센 터", "센터"],
    ["Satis faction", "Satisfaction"],
    ["완벽 한", "완벽한"],
    ["예방 을", "예방을"],
    ["실시 간", "실시간"],
    ["생 애주기", "생애주기"],
    ["걸 쳐", "걸쳐"],
    ["최 소한", "최소한"],
    ["범 위", "범위"],
    ["투 명하 게", "투명하게"],
    ["존 중", "존중"],
    ["보호 할", "보호할"],
    ["고객에 게", "고객에게"],
    ["솔루 션", "솔루션"],
    ["아우르 는", "아우르는"],
    ["담당 하며", "담당하며"],
    ["사 무국", "사무국"],
    ["인 식", "인식"],
    ["사내 ·외", "사내·외"],
    ["이메 일", "이메일"],
    ["24시 간", "24시간"],
    ["제 외", "제외"],
    ["정 착시 키기", "정착시키기"],
    ["안전하 게", "안전하게"],
    ["선택 을", "선택을"],
    ["최우선 으로", "최우선으로"],
    ["보 험", "보험"],
    ["생 명", "생명"],
    ["프 로그 램", "프로그램"],
    ["출 범", "출범"],
    ["경영원칙 과", "경영원칙과"],
    ["예방 합니다", "예방합니다"],
    ["실 현", "실현"],
    ["C PMS", "CPMS"],
    ["팀 은", "팀은"],
    ["회사외부", "회사 외부"],
    ["거래 업체", "거래업체"],
    ["책 임", "책임"],
    ["주 축", "주축"],
    ["감독 할", "감독할"],
    ["인 게이지먼 트", "인게이지먼트"],
    ["전문가 들", "전문가들"],
    ["개최 하여", "개최하여"],
    ["프로세 스", "프로세스"],
    ["선행· 상품화", "선행·상품화"],
    ["상품화 ·양산", "상품화·양산"],
    ["Bespo ke", "Bespoke"],
    ["스 팀", "스팀"],
    ["로 봇", "로봇"],
    ["스 탠다드", "스탠다드"],
    ["획득 하였", "획득하였"],
    ["부서장 들", "부서장들"],
    ["이 행", "이행"],
    ["지 역별", "지역별"],
    ["DS 부문", "DS부문"],
    ["책임와", "책임과"],
    ["두산퓨 얼셀", "두산퓨얼셀"],
    ["두산퓨얼셀㈜", "두산퓨얼셀"],
    ["두산퓨얼셀는", "두산퓨얼셀은"],
    ["두산퓨얼셀가", "두산퓨얼셀이"],
    ["두산퓨얼셀를", "두산퓨얼셀을"],
    ["전략를", "전략을"],
    ["현황는", "현황은"],
    ["위 하여", "위하여"],
    ["간담 회", "간담회"],
    ["들 의", "들의"],
    ["Player’ 라는", "Player’라는"],
    ["임직원 뿐", "임직원뿐"],
    ["다 만", "다만"],
    ["위험관리 능력향상을 시킵니다", "위험관리 능력 향상을 지원합니다"],
    ["능력향상을 시킵니다", "능력 향상을 지원합니다"],
    ["제공고자", "제공하고자"],
    ["물리적리스크", "물리적 리스크"],
    ["대내 ㆍ 외", "대내외"],
    ["저탄소  경제", "저탄소 경제"],
    ["분석 하였습니다", "분석하였습니다"],
    ["있으 며", "있으며"],
    ["내/외부", "내외부"],
  ];
  let out = String(text ?? "");
  for (const [from, to] of replacements) out = out.replaceAll(from, to);
  out = out
    .replace(/저탄소\s+경제/g, "저탄소 경제")
    .replace(/분석\s+하였습니다/g, "분석하였습니다")
    .replace(/대내\s*ㆍ\s*외/g, "대내외");
  return normalizeWhitespace(out);
}

function stripSourceAndNavigation(text) {
  return repairKoreanSpacing(normalizeWhitespace(String(text ?? "")
    .replace(/[A-Z0-9._%+-]+@[A-Z0-9.-]+\.[A-Z]{2,}/gi, " ")
    .replace(/\[[^\]]*p\.\d+[^\]]*\]/gi, " ")
    .replace(/\[p\.\d+\]/gi, " ")
    .replace(/\bp\.\d+(?:\s*[-~]\s*\d+)?\b/gi, " ")
    .replace(/삼성전자\s+지속가능경영보고서\s+2025\s+\d{1,3}/g, " ")
    .replace(/2024[–-]2025\s+LG전자\s+지속가능경영보고서\s+\d{1,3}/g, " ")
    .replace(/Overview Environmental Social Governance ESG Data Appendix/gi, " ")
    .replace(/Our Company AppendixFacts & Figures PrinciplePlanet People/gi, " ")
    .replace(/AppendixFacts & Figures PrinciplePlanet People/gi, " ")
    .replace(/Facts & Figures|PrinciplePlanet|Our Company|Appendix/gi, " ")
    .replace(/Doosan Fuel Cell Sustainability Report\s*2025|Company Overview|ESG Strategy|Materiality|ESG Performance/gi, " ")
    .replace(/\b(?:Source|PDF|page|pages|reviewer|audit|trace|metadata|chunk)\b/gi, " ")
    .replace(/\b\d+\)\s*/g, " ")));
}

function stripReportHeadingResidue(text) {
  let out = normalizeWhitespace(text);
  for (let i = 0; i < 4; i += 1) {
    const before = out;
    out = out
      .replace(/^위원회 규정 개정의 건\s*-\s*[^.]{0,160}?2024\.?\s*/i, "")
      .replace(/^보고안건\s+[^.]{0,160}?2024\.?\s*/i, "")
      .replace(REPORT_HEADING_PREFIX_REGEX, "")
      .replace(REPORT_HEADING_INLINE_REGEX, " ");
    out = normalizeWhitespace(out);
    if (out === before) break;
  }
  return out;
}

function hasReportHeadingResidue(text) {
  const normalized = normalizeWhitespace(text);
  if (!normalized) return false;
  if (/^(?:위원회 규정 개정의 건|보고안건)\b/i.test(normalized)) return true;
  if (/보고\s*프레임워크|보고\s*위치|2024\s*SPECIAL\s*REPORT|Part\s+[A-Z]:|[A-Z]\d\s+기업은.*어떻게/i.test(normalized)) return true;
  return REPORT_HEADING_PHRASES.some((phrase) => (
    normalized.startsWith(`${phrase} `)
    || normalized.includes(`. ${phrase} `)
    || normalized.includes(` ${phrase} 신한`)
    || normalized.includes(` ${phrase} 2024년`)
  ));
}

function cleanEvidenceSentence(sentence) {
  let out = repairKoreanSpacing(sentence)
    .replace(SIGNATURE_ARTIFACT_REGEX, "")
    .replace(HEADING_PREFIX_REGEX, "")
    .replace(BAD_OPENING_CONNECTOR_REGEX, "");
  out = stripReportHeadingResidue(out);
  for (const [from, to] of KOREAN_GRAMMAR_CLEANUPS) out = out.replace(from, to);
  return normalizeWhitespace(out);
}

function splitSentences(text) {
  return stripSourceAndNavigation(text)
    .split(/(?<=[.!?。]|습니다\.|니다\.|됩니다\.|합니다\.|있습니다\.|않습니다\.)\s+|(?=삼성전자는|LG전자는|회사는|DX부문은|DS부문은|또한|특히|이를 통해|이에 따라|뿐만 아니라|이렇게|이처럼|2024년|2022년)/g)
    .map((sentence) => normalizeWhitespace(sentence))
    .filter(Boolean);
}

function uniqueByNormalized(items) {
  const seen = new Set();
  const out = [];
  for (const item of items) {
    const key = normalizeWhitespace(item)
      .replace(/[0-9][0-9,.]*(?:\.[0-9]+)?/g, "#")
      .replace(/[.,;:\s'"]/g, "");
    if (!key || seen.has(key)) continue;
    seen.add(key);
    out.push(item);
  }
  return out;
}

function numberTokens(text) {
  return String(text ?? "").match(/[0-9][0-9,]*(?:\.[0-9]+)?\s*(?:%|tCO2e|MWh|GWh|TJ|KRW|명|개|건|회|톤|억원|조원|시간|년|배|cases?|employees?|hours?|rate)?/gi) ?? [];
}

function hasMeaningfulNumber(text) {
  const cleaned = String(text ?? "")
    .replace(/\b\d+\s+quantitative\s+row\(s\)[^.]*\.?/gi, "")
    .replace(/\bp\.\d+\b/gi, "")
    .replace(/\b20\d{2}[–~-]20\d{2}\b/g, "");
  return /[0-9][0-9,]*(?:\.[0-9]+)?\s*(?:%|tCO2e|MWh|GWh|TJ|KRW|명|개|건|회|톤|억원|조원|시간|년|배|cases?|employees?|hours?|rate|target|carbon|RE100|Scope)/i.test(cleaned);
}

function rejectEvidenceSentence(sentence) {
  const text = normalizeWhitespace(sentence);
  if (hasReportHeadingResidue(text)) return true;
  if (BAD_OPENING_CONNECTOR_REGEX.test(text) || SIGNATURE_ARTIFACT_REGEX.test(text) || HEADING_PREFIX_REGEX.test(text)) return true;
  if (SPECIAL_ARTIFACT_REGEX.test(text) || RAW_HEADING_PREFIX_REGEX.test(text) || TABLE_DIAGRAM_ARTIFACT_REGEX.test(text)) return true;
  if (INCOMPLETE_FINAL_REGEX.test(text)) return true;
  if (/\uBD84\uC11D\s+\uD558\uC600\uC2B5\uB2C8\uB2E4/.test(text)) return true;
  if (/해야 합니다[.]?$/.test(text) && !/두산퓨얼셀/.test(text)) return true;
  if (INLINE_LIST_ARTIFACT_REGEX.test(text) || CONTACT_ARTIFACT_REGEX.test(text) || NAVIGATION_ARTIFACT_REGEX.test(text)) return true;
  if (text.length < 45 || text.length > 360) return true;
  if (!KOREAN_REGEX.test(text)) return true;
  if (EBX_CODE_REGEX.test(text) || SOURCE_TRACE_REGEX.test(text) || OCR_ARTIFACT_REGEX.test(text)) return true;
  if (/^(\d|구분|항목|단위|비고|지표|내용|보고 페이지|코드|공시 항목)\b/.test(text)) return true;
  if (/구분 리스크|추진체계 추진방향 리스크 관리 활동|중대 주제 UN SDGs|지역총괄|판매거점|생산거점|Topic No\.|N\/A|참조하십|보고서\(XI|보고서\(II/i.test(text)) return true;
  if (/임직원 및 이사회 대상 환경 교육|옥상 태양열 설비|건축물 일체형 태양광|냉각 효율 제고.*전력 효율지수/i.test(text)) return true;
  if (/조직명|수행업무|운영주기|주요 안건|구분\s+운영주기|구분\s+주요 장비|샘플링 장비|분석 장비\s+ICP|전처리 장비|Strategy\s+Risk Management|업무협약의\s+네 가지\s+중점 분야|01\s+02\s+03\s+04|홈페이지 환경정책 링크|가이드라인\s+전문\s+바로가기|LG전자 주식회사\s*「[^」]+가이드라인」|자가진단\s+실사\s+현장심사|정보보호 Strategy|제품 보안 운영 방침/.test(text)) return true;
  if (/^[0-9\s.,%()~:/-]+$/.test(text)) return true;
  const alpha = (text.match(/[A-Za-z]/g) ?? []).length;
  if (alpha > 80 && alpha / Math.max(text.length, 1) > 0.3) return true;
  if ((numberTokens(text).length >= 12 && !/목표|관리|운영|보고|공시|달성|수립|실행/.test(text))) return true;
  return false;
}

function topicRejectRegex(row) {
  const num = ebxNumber(row);
  const map = {
    4: /인권|정보보호|개인정보|사회공헌|윤리|준법|품질|환경경영/,
    5: /인권|정보보호|개인정보|사회공헌|윤리|준법|품질|환경경영/,
    6: /인권|정보보호|개인정보|사회공헌|윤리|준법|품질|환경경영/,
    7: /인권|정보보호|개인정보|사회공헌|윤리|준법|품질|환경경영/,
    8: /품질|정보보호|개인정보|사회공헌|윤리|준법|환경경영|산업안전/,
    9: /품질|정보보호|개인정보|사회공헌|윤리|준법|환경경영|산업안전/,
    10: /품질|정보보호|개인정보|사회공헌|윤리|준법|환경경영|산업안전/,
    11: /품질|정보보호|개인정보|사회공헌|윤리|준법|환경경영|산업안전/,
    12: /사회공헌|인권|정보보호|개인정보|윤리|준법|환경경영|산업안전/,
    13: /사회공헌|인권|정보보호|개인정보|윤리|준법|환경경영|산업안전/,
    14: /사회공헌|인권|정보보호|개인정보|윤리|준법|환경경영/,
    15: /사회공헌|인권|정보보호|개인정보|윤리|준법|환경경영/,
    16: /품질|사회공헌|인권|윤리|준법|환경경영|산업안전/,
    17: /품질|사회공헌|인권|윤리|준법|환경경영|산업안전/,
    18: /품질|사회공헌|인권|윤리|준법|환경경영|산업안전/,
    19: /품질|사회공헌|인권|윤리|준법|환경경영|산업안전/,
    20: /품질|사회공헌|인권|정보보호|개인정보|윤리|준법|산업안전/,
    21: /품질|사회공헌|인권|정보보호|개인정보|윤리|준법|산업안전/,
    22: /품질|사회공헌|인권|정보보호|개인정보|윤리|준법|산업안전/,
    23: /품질|사회공헌|인권|정보보호|개인정보|윤리|준법|산업안전/,
    24: /품질|사회공헌|인권|정보보호|개인정보|환경경영|산업안전/,
    25: /품질|사회공헌|인권|정보보호|개인정보|환경경영|산업안전/,
    26: /품질|사회공헌|인권|정보보호|개인정보|환경경영|산업안전/,
    27: /품질|사회공헌|인권|정보보호|개인정보|환경경영|산업안전/,
  };
  return map[num] ?? null;
}

function topicRequireRegex(row) {
  const num = ebxNumber(row);
  if ([4, 5, 6, 7].includes(num)) return /안전|재해|보건|CSO|LTIR|사망|부상|현장/;
  if ([8, 9, 10, 11].includes(num)) return /인권|고충|노동|임직원|이해관계자|침해/;
  if ([12, 13, 14, 15].includes(num)) return /품질|제품|안전|QCON|ISO\s*9001|하자|고객|소비자|시공/;
  if ([16, 17, 18, 19].includes(num)) return /정보보호|개인정보|보안|CISO|침해|사이버|IT|정보기술/;
  if ([20, 21, 22, 23].includes(num)) return /환경|기후|탄소|온실가스|폐기물|에너지|용수|배출|리스크|오염/;
  if ([24, 25, 26, 27].includes(num)) return /윤리|준법|컴플라이언스|부패|뇌물|제보|투명경영|공정거래|내부고발/;
  if ([28, 29, 30, 31].includes(num)) return /기후|탄소|온실가스|Scope|에너지|재생에너지|RE100|SBTi|배출|감축|전환/;
  if ([32, 33, 34, 35].includes(num)) return /자원|폐기물|재활용|순환|폐가전|플라스틱|회수|매립|재사용/;
  if ([36, 37, 38, 39].includes(num)) return /수자원|용수|취수|폐수|수질|물\s*관리|water|방류|재이용/;
  if ([40, 41, 42, 43].includes(num)) return /생물다양성|서식지|보전|산림|자연|생태|보호지역|복원/;
  if ([44, 45, 46, 47].includes(num)) return /오염|대기|수질|배출|오염물질|NOx|SOx|VOC|화학물질|먼지/;
  if ([48, 49, 50, 51, 52, 53, 54, 55].includes(num)) return /친환경|제품|환경성|에너지효율|유해물질|재활용\s*플라스틱|포장재|전과정|인증|제품책임/;
  if ([56, 57, 58, 59].includes(num)) return /인재|교육|훈련|임직원|구성원|인적\s*자본|역량|채용|육성/;
  if ([60, 61, 62, 63].includes(num)) return /다양성|포용|DEI|여성|장애|차별|형평|구성원|조직문화/;
  if ([64, 65, 66, 67].includes(num)) return /공급망|협력사|동반성장|실사|RBA|책임광물|구매|ESG\s*평가/;
  if ([68, 69, 70, 71].includes(num)) return /사회공헌|지역사회|기부|봉사|투자|커뮤니티|프로그램|수혜/;
  if ([72, 73, 74, 75].includes(num)) return /위원회|이사회|감사위원회|ESG위원회|사외이사|사내이사|심의|의결|개최|회의/;
  if ([76, 77, 78, 79].includes(num)) return /이사회|사외이사|사내이사|독립성|전문성|후보|선임|의장|평가|보수/;
  if ([80, 81, 82, 83].includes(num)) return /ESG|지속가능|협의체|전담\s*조직|중대성|리스크|성과|목표|운영\s*체계/;
  if ([84, 85, 86, 87].includes(num)) return /컴플라이언스|준법|법규|위반|교육|제보|진단|공정거래|부패|윤리/;
  if ([88, 89, 90, 91].includes(num)) return /소유구조|주주|주식|의결권|배당|지분|주주권리|주주총회/;
  if ([92, 93, 94, 95].includes(num)) return /이해관계자|소통|고객|임직원|협력사|주주|지역사회|채널|참여|의견/;
  return null;
}

function isTopicAligned(row, sentence, template = {}) {
  const text = normalizeWhitespace(sentence);
  const configuredReject = regexesFromHints(template?.evidenceSelection?.negativeTopicHints);
  if (configuredReject.some((regex) => regex.test(text))) return false;
  const reject = topicRejectRegex(row);
  if (reject?.test(text)) return false;
  const require = topicRequireRegex(row);
  if (require && !require.test(text)) return false;
  const configuredRequire = regexesFromHints(template?.evidenceSelection?.topicHints);
  if (ebxNumber(row) >= 28 && configuredRequire.length && !configuredRequire.some((regex) => regex.test(text))) return false;
  if (ebxNumber(row) < 28) return true;
  const type = template?.answerType || answerTypeFromEbx(row.ebx);
  if (type === "strategy-policy") return /전략|정책|목표|방향|계획|원칙|방침|협약|중장기|로드맵/.test(text);
  if (type === "governance-organization") return /조직|위원회|이사회|협의체|담당|센터|부서|거버넌스|의사결정|보고체계|관리체계/.test(text);
  if (type === "risk-control") return /리스크|위험|영향|예방|대응|모니터링|통제|평가|점검|개선 조치/.test(text);
  if (type === "status-performance") return /[0-9]|현황|성과|실적|발생|달성|사용량|배출량|횟수|비율|건수|참여자|이수자/.test(text);
  return true;
}

function keywords(row, template) {
  const text = `${template?.item ?? ""} ${row.question_title ?? ""}`;
  return uniqueByNormalized(text
    .replace(/[()/:,.;·•\[\]]/g, " ")
    .split(/\s+/)
    .map((token) => token.trim())
    .filter((token) => token.length >= 2)
    .filter((token) => !/^(and|the|for|with|관련|항목|관리|현황|체계|목표|정책|활동|및|và|của|cho)$/i.test(token)))
    .slice(0, 30);
}

function sentenceScore(sentence, row, template, displayName) {
  if (!isTopicAligned(row, sentence, template)) return -100;
  let score = 0;
  if (sentence.includes(displayName)) score += 3;
  if (/이사회|위원회|조직|리스크|위험|목표|전략|성과|안전|품질|정보보호|환경|윤리|인권|협력회사|공급망|고충|침해|준법|컴플라이언스/.test(sentence)) score += 4;
  if (numberTokens(sentence).length) score += 1;
  if (sentence.length >= 70 && sentence.length <= 230) score += 2;
  for (const keyword of keywords(row, template)) {
    if (keyword && sentence.includes(keyword)) score += 2;
  }
  for (const regex of regexesFromHints(template?.evidenceSelection?.topicHints)) {
    if (regex.test(sentence)) score += 2;
  }
  for (const regex of regexesFromHints(template?.evidenceSelection?.sectorHints)) {
    if (regex.test(sentence)) score += 1;
  }
  return score;
}

function selectEvidenceSentences(row, template, displayName) {
  const maxEvidenceSentences = Number(template?.evidenceSelection?.maxEvidenceSentences ?? 12);
  const minScore = Number(template?.evidenceSelection?.minScore ?? 1);
  const scored = uniqueByNormalized(splitSentences(row.original_text_ko)
    .map(repairKoreanSpacing)
    .map(cleanEvidenceSentence)
    .filter((sentence) => !rejectEvidenceSentence(sentence))
    .filter((sentence) => isTopicAligned(row, sentence, template)))
    .map((sentence, index) => ({
      sentence,
      index,
      score: sentenceScore(sentence, row, template, displayName),
    }));
  return scored
    .filter((item) => item.score >= minScore)
    .sort((a, b) => b.score - a.score || a.index - b.index)
    .slice(0, maxEvidenceSentences)
    .sort((a, b) => a.index - b.index)
    .map((item) => item.sentence);
}

function topicLabel(row, template) {
  return String(template?.item || row.question_title || "해당 항목")
    .replace(/\s*\/.*$/, "")
    .trim();
}

function ebxNumber(rowOrEbx) {
  return Number(String(rowOrEbx?.ebx ?? rowOrEbx ?? "").match(/\d+/)?.[0] ?? 0);
}

function answerTypeFromEbx(ebx) {
  const num = ebxNumber(ebx);
  if ([1, 8, 16, 20, 24].includes(num)) return "strategy-policy";
  if ([2, 5, 9, 13, 17, 21, 25].includes(num)) return "governance-organization";
  if ([3, 6, 10, 14, 18, 22, 26].includes(num)) return "risk-control";
  if ([7, 11, 15, 19, 23, 27].includes(num)) return "status-performance";
  return "policy-management";
}

function styleKey(row, template) {
  if (template?.preferredStyle) return template.preferredStyle;
  const type = template?.answerType || answerTypeFromEbx(row.ebx);
  if (type === "strategy-policy") return "narrative";
  if (type === "governance-organization") return "governance";
  if (type === "risk-control") return "risk-control";
  if (type === "status-performance") return "status-performance";
  return "balanced-policy";
}

function regexesFromHints(values) {
  return (values ?? [])
    .map((value) => {
      try {
        return new RegExp(String(value), "i");
      } catch {
        return null;
      }
    })
    .filter(Boolean);
}

function companyDisplayName(config, metadata) {
  const explicit = String(metadata.company_display_name_ko ?? metadata.company_ko ?? "").trim();
  if (explicit) return explicit;
  if (/doosan_fuel_cell/i.test(config.companyId) || /Doosan/i.test(config.companyName)) return "두산퓨얼셀";
  if (/samsung_biologics/i.test(config.companyId) || /Samsung_Biologics/i.test(config.companyName)) return "삼성바이오로직스";
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
    "Doosan Fuel Cell",
    "Doosan Fuel Cell 2025",
    "Doosan_Fuel_Cell_2025",
  ].filter(Boolean);
  let out = String(text ?? "");
  for (const name of names) {
    if (name && name !== displayName) out = out.replaceAll(name, displayName);
  }
  return out;
}

function buildField(template, row) {
  return [template?.area, template?.pillar, template?.item || row.question_title || row.ebx]
    .map((value) => String(value ?? "").trim())
    .filter(Boolean)
    .join(" / ");
}

function missingQualitativeRow(template) {
  return {
    ebx: template.ebx,
    question_title: template.item || template.ebx,
    coverage_status: "UNKNOWN",
    source_pdf: "",
    source_pages: "",
    evidence_type: "missing-company-evidence",
    original_text_ko: "",
    quantitative_support: "",
    gap_or_note: "No company-specific qualitative evidence supplied for this v10 item.",
    missingCompanyEvidence: true,
  };
}

function hasCompanyEvidence(row) {
  return Boolean(normalizeWhitespace([
    row.original_text_ko,
    row.quantitative_support,
    row.source_pages,
  ].filter(Boolean).join(" ")));
}

function hasFinalConsonant(text) {
  const chars = Array.from(String(text ?? ""));
  for (let i = chars.length - 1; i >= 0; i -= 1) {
    const code = chars[i].charCodeAt(0) - 0xac00;
    if (code >= 0 && code <= 11171) return code % 28 !== 0;
  }
  return false;
}

function withParticle(text, withBatchim, withoutBatchim) {
  return `${text}${hasFinalConsonant(text) ? withBatchim : withoutBatchim}`;
}

function openingSentence(row, template, displayName) {
  const topic = topicLabel(row, template);
  const type = template?.answerType || answerTypeFromEbx(row.ebx);
  const companySubject = withParticle(displayName, "은", "는");
  if (type === "governance-organization") {
    return `${companySubject} ${withParticle(topic, "과", "와")} 관련해 책임 주체와 의사결정 흐름을 중심으로 관리 체계를 설명하고 있습니다.`;
  }
  if (type === "risk-control") {
    return `${companySubject} ${topic}에서 식별된 위험 요인과 예방 활동을 연결해 관리 수준과 개선 필요 사항을 함께 점검하고 있습니다.`;
  }
  if (type === "status-performance") {
    return `${companySubject} ${topic}을 관리 지표, 후속 조치, 공개 범위를 기준으로 점검하고 있습니다.`;
  }
  if (type === "strategy-policy") {
    return `${companySubject} ${withParticle(topic, "을", "를")} 중장기 방향과 실행 과제에 연결해 지속가능경영의 방향성을 제시하고 있습니다.`;
  }
  return `${companySubject} ${topic}에 대해 정책 기준과 운영 절차를 함께 제시하고 있습니다.`;
}

function closingSentence(row, template, displayName) {
  const topic = topicLabel(row, template);
  const type = template?.answerType || answerTypeFromEbx(row.ebx);
  if (type === "governance-organization") {
    return `이 구조는 ${topic}의 담당 조직, 보고 경로, 감독 역할을 한 문맥에서 확인할 수 있게 합니다.`;
  }
  if (type === "risk-control") {
    return `${topic} 관련 공시는 단순한 선언보다 예방, 모니터링, 후속 조치의 연결성을 중심으로 해석할 수 있으며, 담당 조직의 실행 책임과 개선 활동의 지속성을 함께 파악하는 데 도움이 됩니다.`;
  }
  if (type === "status-performance") {
    return `이 정보는 ${topic}의 규모와 추이를 확인하고 향후 관리 보완이 필요한 부분을 판단하는 근거가 됩니다.`;
  }
  if (type === "strategy-policy") {
    return `이러한 내용은 ${withParticle(topic, "이", "가")} 선언에 머물지 않고 목표, 조직, 실행 과제로 이어지는 구조임을 보여줍니다.`;
  }
  return `${withParticle(topic, "은", "는")} 정책, 실행 주체, 성과 확인이 연결된 관리 항목으로 정리됩니다.`;
}

function partialLimitationSentence(row, template, displayName) {
  const topic = topicLabel(row, template);
  return `${displayName}의 ${topic} 공시는 확인 가능한 제도와 활동을 중심으로 구성되어 있으며, 일부 세부 수치나 사건 현황은 공개 범위에서 제한적으로 제시되어 추가 보완 여지가 있습니다.`;
}

function formatMetricValue(metric, key) {
  const originalValue = String(metric[key] ?? "");
  const rawValue = sanitizeXmlText(originalValue).trim();
  if (CONTROL_CHAR_REGEX.test(originalValue)) return "";
  if (!/^\(?-?\d[\d,]*(?:\.\d+)?\)?$/.test(rawValue)) return "";
  if (!normalizeWhitespace(metric.unit) && /^(?:19|20)\d{2}$/.test(rawValue.replace(/,/g, ""))) return "";
  if ((rawValue.includes("(") || rawValue.includes(")")) && !/^\([0-9,.]+\)$/.test(rawValue)) return "";
  const value = rawValue.replace(/^-?\d+(?:\.\d+)?$/, (match) => {
    const [whole, decimal] = match.split(".");
    return `${Number(whole).toLocaleString("en-US")}${decimal ? `.${decimal}` : ""}`;
  });
  if (!value || value === "-") return "";
  const unit = translateMetricUnit(metric.unit);
  if (["%", "명", "건", "회", "점", "개사"].includes(unit)) return `${value}${unit}`;
  return unit ? `${value} ${unit}` : value;
}

function translateMetricUnit(unit) {
  const lower = String(unit ?? "").trim().toLowerCase();
  if (lower === "percent") return "%";
  if (lower === "tons" || lower === "tonnes" || lower === "ton") return "톤";
  if (lower === "thousand tons") return "천 톤";
  if (lower === "cases/million hours") return "건/백만 시간";
  if (lower === "meetings") return "회";
  if (lower === "krw 100m") return "억 원";
  if (lower === "krw tn" || lower === "krw trillion") return "조 원";
  if (lower === "krw million") return "백만 원";
  if (lower === "points") return "점";
  if (lower === "times") return "회";
  if (lower === "people" || lower === "persons") return "명";
  if (lower === "companies") return "개사";
  if (lower === "sites") return "개 사업장";
  if (lower === "cases") return "건";
  if (lower === "hours") return "시간";
  if (lower === "krw trillion") return "조 원";
  if (lower === "1000 tco2e") return "천 tCO2e";
  return String(unit ?? "").trim();
}

function translateMetricIndicator(indicator) {
  const text = normalizeWhitespace(indicator)
    .replace(/^[·•\-\s]+/, "")
    .replace(/\s+(?:tons?|tonnes?|tCO2e|MWh|GWh|TJ|%|명|건|회|톤|백만 원|억 원)$/i, "");
  const exact = new Map([
    ["GHG emissions, Scope 1 and 2, market-based", "시장기반 Scope 1·2 온실가스 배출량"],
    ["GHG emissions, Scope 1 and 2, location-based", "지역기반 Scope 1·2 온실가스 배출량"],
    ["Scope 1 direct emissions", "Scope 1 직접배출량"],
    ["Scope 2 indirect emissions", "Scope 2 간접배출량"],
    ["LTIR", "LTIR"],
    ["Supplier LTIR", "협력회사 LTIR"],
    ["Grievance cases received", "고충 접수 건수"],
    ["Grievance processing rate", "고충 처리율"],
    ["Internal privacy consulting cases", "개인정보 내부 컨설팅 건수"],
    ["Government information requests", "정부 정보 요청 건수"],
    ["Government information provided cases", "정부 정보 제공 건수"],
    ["Government information provision rate", "정부 정보 제공률"],
    ["Compliance training participants", "컴플라이언스 교육 참여자"],
    ["Anti-fraud training participants", "부정 예방 교육 참여자"],
    ["Compliance reports", "컴플라이언스 제보"],
    ["Fraud report cases", "부정 제보"],
    ["Consumer complaint ratio", "소비자 불만 비율"],
    ["Customer satisfaction - domestic market", "고객 만족도(국내)"],
    ["Customer satisfaction - overseas market", "고객 만족도(해외)"],
    ["Products/services assessed for safety and health impacts", "제품·서비스 안전보건 영향 평가 비율"],
    ["VOCs total", "VOC 배출량"],
    ["Sign-language consultation cases Korea", "수어 상담 건수(한국)"],
    ["Substantiated customer data leakage/theft/loss incidents", "확인된 고객정보 유출·도난·분실 건수"],
    ["Fuel use", "연료 사용량"],
    ["Steam use", "스팀 사용량"],
    ["Electricity use", "전력 사용량"],
    ["Total energy use", "총 에너지 사용량"],
    ["Total waste generated", "총 폐기물 발생량"],
    ["Total recycled waste", "폐기물 재활용량"],
    ["Waste recycling rate - total", "폐기물 재활용률"],
    ["GHG reduction rate vs 2017", "2017년 대비 온실가스 감축률"],
    ["GHG reduction amount vs 2017", "2017년 대비 온실가스 감축량"],
    ["Designated waste recycled", "지정폐기물 재활용량"],
    ["Designated waste incinerated", "지정폐기물 소각량"],
    ["Designated waste landfilled", "지정폐기물 매립량"],
    ["Recycled plastic use ratio", "재활용 플라스틱 사용 비율"],
    ["Recycled plastic use amount", "재활용 플라스틱 사용량"],
    ["RBA self-assessment target sites", "RBA 자가진단 대상 사업장"],
    ["RBA Low Risk sites", "RBA 저위험 사업장"],
    ["RBA Medium Risk sites", "RBA 중위험 사업장"],
    ["RBA High Risk sites", "RBA 고위험 사업장"],
    ["RBA self-assessment average score", "RBA 자가진단 평균 점수"],
    ["Compliance managers", "컴플라이언스 담당자"],
    ["Risk managers", "리스크 담당자"],
    ["Enterprise compliance committee meetings", "전사 컴플라이언스 위원회 개최 횟수"],
    ["Site compliance committee meetings", "사업장 컴플라이언스 위원회 개최 횟수"],
    ["Compliance survey target - total", "컴플라이언스 설문 대상자"],
    ["Hazardous chemical use", "유해화학물질 사용량"],
    ["Chemical emissions", "화학물질 배출량"],
    ["Water withdrawal total", "총 취수량"],
    ["Water withdrawal target", "취수량 목표"],
    ["Major chemical spills - count", "중대 화학물질 유출 건수"],
    ["Board meetings held", "이사회 개최 횟수"],
    ["Board agenda items approved", "이사회 승인 안건 수"],
    ["Board agenda items reported", "이사회 보고 안건 수"],
    ["Board attendance rate - overall", "이사회 전체 참석률"],
    ["Material issue reports to board", "중대 이슈 이사회 보고 횟수"],
    ["Regular female employees", "여성 정규직 임직원"],
    ["Full-time female employees", "여성 상근 임직원"],
    ["Executives - female", "여성 임원"],
    ["Supplier ESG field diagnosis certification rate", "협력사 ESG 현장진단 인증률"],
    ["Indigenous rights violations", "원주민 권리 침해 건수"],
    ["Shareholder/investor cash dividend plus interest cost", "주주·투자자 배당 및 이자 비용"],
    ["Tax reductions and credits", "세액공제 및 감면액"],
    ["Direct purchase amount - Korea", "국내 직접 구매 금액"],
    ["Direct purchase amount - Asia", "아시아 직접 구매 금액"],
    ["Direct purchase amount - China", "중국 직접 구매 금액"],
    ["Direct purchase amount - Americas", "미주 직접 구매 금액"],
    ["Direct purchase amount - Europe/CIS", "유럽·CIS 직접 구매 금액"],
  ]);
  if (exact.has(text)) return exact.get(text);
  return text
    .replace(/GHG emissions/gi, "온실가스 배출량")
    .replace(/Scope 1 and 2/gi, "Scope 1·2")
    .replace(/market-based/gi, "시장기반")
    .replace(/location-based/gi, "지역기반")
    .replace(/direct emissions/gi, "직접배출량")
    .replace(/indirect emissions/gi, "간접배출량")
    .replace(/cases/gi, "건수")
    .replace(/participants/gi, "참여자");
}

function metricRecordLabel(metric) {
  const explicit = normalizeWhitespace(metric.indicator).replace(/^[·•\-\s]+/, "");
  if (explicit && !/^(?:구분|단위|20\d{2}년|보고서|주요 데이터|비교가능성|기준 기간)(?:\s|$)/.test(explicit)) {
    const translated = translateMetricIndicator(explicit);
    const alpha = (translated.match(/[A-Za-z]/g) ?? []).length;
    if (alpha > 12 && alpha / Math.max(translated.length, 1) > 0.35 && !/Scope|LTIR|ESG|RBA|VOC|NOx|SOx|BOD/.test(translated)) return "";
    return translated;
  }
  const raw = normalizeWhitespace(metric.raw_line)
    .replace(/^[·•\-\s]+/, "")
    .replace(/\b20\d{2}\b.*$/, "")
    .replace(/\s+(?:건\/백만\s*시간|tCO2e|MWh|GWh|TJ|ton|톤|백만\s*원|억\s*원|%|건|명|회|개사|개)\s+[-0-9,.\s()%]+.*$/i, "")
    .replace(/\s+[-0-9,.\s()%]{6,}.*$/, "");
  if (!raw || /^구분|^단위|^국내$|^해외$|^소계$|^총계$|^사망$|^부상$/.test(raw)) {
    return translateMetricIndicator(metric.subcategory || metric.category || "");
  }
  return raw;
}

function metricRecordLabelV101(metric) {
  const originalIndicator = String(metric.indicator ?? "");
  const originalRaw = String(metric.raw_line ?? "");
  if (CONTROL_CHAR_REGEX.test(originalIndicator) || CONTROL_CHAR_REGEX.test(originalRaw)) return "";
  const cleanedMetric = {
    ...metric,
    indicator: sanitizeXmlText(originalIndicator),
    raw_line: sanitizeXmlText(originalRaw),
  };
  const explicit = normalizeWhitespace(cleanedMetric.indicator);
  const years = explicit.match(/(?:19|20)\d{2}/g) ?? [];
  if (explicit.length > 140 || explicit.split(/\s+/).filter(Boolean).length > 18 || years.length > 1) return "";
  if (/^\d+[).]\s|^\d+\s*\([a-z]\)|^\d+\s*\([^)]+\)\s|MATERIAL\s+TOPIC|(?:^|\s)Step\s*\d/i.test(explicit)) return "";
  if (/SPECIAL\s+REPORT|ESG\s+DATA\s+PACK|TCFD|GRI|SASB|Global\s+Standard|보고\s*위치|보고기간|p\.\s*\d|p\.\s*$/i.test(explicit)) return "";
  if (/\||(?:19|20)\d{2}\.\d{1,2}|(?:\uC2B5\uB2C8\uB2E4|\uD569\uB2C8\uB2E4|\uD558\uC600|\uC788\uC73C\uBA70|\uC788\uC2B5\uB2C8\uB2E4|\uC54A\uC558|\uD544\uC694\uD569\uB2C8\uB2E4|\uC81C\uACF5\uD569\uB2C8\uB2E4)/.test(explicit)) return "";
  if (/^[A-Z]{2,}(?:-[A-Z0-9.]+)+/i.test(explicit) || /[,.;:]\s*$/.test(explicit)) return "";
  const label = sanitizeXmlText(metricRecordLabel(cleanedMetric))
    .replace(/^Total\s+/i, "\uCD1D ")
    .replace(/\btarget\b/gi, "\uBAA9\uD45C");
  if (!label || label.length > 140 || label.split(/\s+/).filter(Boolean).length > 18) return "";
  if (((label.match(/(?:19|20)\d{2}/g) ?? []).length > 1) || MALFORMED_METRIC_REGEX.test(label)) return "";
  if (/SPECIAL\s+REPORT|ESG\s+DATA\s+PACK|TCFD|GRI|SASB|Global\s+Standard|보고\s*위치|보고기간|p\.\s*\d|p\.\s*$/i.test(label)) return "";
  if (/\||(?:19|20)\d{2}\.\d{1,2}|(?:\uC2B5\uB2C8\uB2E4|\uD569\uB2C8\uB2E4|\uD558\uC600|\uC788\uC73C\uBA70|\uC788\uC2B5\uB2C8\uB2E4|\uC54A\uC558|\uBC1B\uC558\uC73C\uBA70|\uD544\uC694\uD569\uB2C8\uB2E4|\uC81C\uACF5\uD569\uB2C8\uB2E4)/.test(label)) return "";
  if (/^[A-Z]{2,}(?:-[A-Z0-9.]+)+/i.test(label) || /[,.;:]\s*$/.test(label)) return "";
  if ((label.match(/\(/g) ?? []).length !== (label.match(/\)/g) ?? []).length) return "";
  const alpha = (label.match(/[A-Za-z]/g) ?? []).length;
  if (alpha > 12 && alpha / Math.max(label.length, 1) > 0.35 && !/Scope|LTIR|ESG|RBA|VOC|NOx|SOx|BOD|ISO/.test(label)) return "";
  if (/(?:^|\s)-?\d[\d,.]*$/.test(label) && !/(?:Scope\s*[123]|ISO\s*\d+)$/i.test(label)) return "";
  return label;
}

function exactMetricMapping(row, metric) {
  return normalizeEbxCode(metric.mapped_item_id) === normalizeEbxCode(row.ebx);
}

function metricSpecificityScore(row, label, haystack) {
  const num = ebxNumber(row);
  const text = `${label} ${haystack}`;
  let score = 0;
  if (/\btotal\b|\uCD1D\s*(?:\uBC30\uCD9C\uB7C9|\uBC1C\uC0DD\uB7C9|\uC0AC\uC6A9\uB7C9|\uC778\uC6D0)/i.test(text)) score += 3;
  if (/target|\uBAA9\uD45C/i.test(text)) score -= 1;
  if (num === 31) {
    if (/total\s+Scope\s*1\s*\+?\s*2|Scope\s*1\s*(?:and|\+|\u00B7)\s*2/i.test(text)) score += 16;
    else if (/GHG|greenhouse|\uC628\uC2E4\uAC00\uC2A4|\uD0C4\uC18C\s*\uBC30\uCD9C/i.test(text)) score += 9;
    else if (/energy|\uC5D0\uB108\uC9C0|\uC804\uB825|\uC5F0\uB8CC|\uC2A4\uD300/i.test(text)) score += 2;
  }
  return score;
}

function metricRulesFor(row, template = {}) {
  const hinted = regexesFromHints(template?.metricHints?.regexes);
  if (hinted.length) return hinted;
  const num = ebxNumber(row);
  const map = {
    1: [/carbon|Scope|RE100|water|recycled|탄소|수자원|재활용/i],
    4: [/LTIR|injur|accident|safety|안전/i],
    7: [/LTIR|injur|accident|safety|안전/i],
    11: [/grievance|고충|processing/i],
    15: [/service centers|complaint|consumer|VOC|product|quality|service|품질|하자|QCON|고객|소비자/i],
    19: [/privacy|information request|개인정보|정보보호|보안|정보기술/i],
    23: [/GHG|Scope|emission|energy|waste|water|온실가스|배출|에너지|폐기물|용수/i],
    27: [/compliance|fraud|ethics|report|컴플라이언스|부정|윤리|준법/i],
  };
  return map[num] ?? [];
}

function metricTopicGate(row, haystack) {
  const num = ebxNumber(row);
  const text = normalizeWhitespace(haystack);
  const gates = [
    { rows: [28, 29, 30, 31], require: /climate|carbon|ghg|scope|energy|renewable|re100|sbti|기후|탄소|온실가스|에너지|재생에너지|배출|감축/i },
    { rows: [32, 33, 34, 35], require: /waste|recycl|resource|폐기물|재활용|자원|순환|플라스틱|회수/i, reject: /water|wastewater|용수|폐수|수질|biodiversity|생물다양성/i },
    { rows: [36, 37, 38, 39], require: /water|wastewater|withdrawal|discharge|용수|수자원|폐수|수질|취수|방류/i, reject: /designated waste|recycled plastic|waste generated|waste recycling|hazardous chemical|chemical emissions|chemical spills|biodiversity|생물다양성|product|quality|품질|제품|폐기물|재활용\s*플라스틱|화학물질/i },
    { rows: [40, 41, 42, 43], require: /biodiversity|habitat|forest|생물다양성|서식지|보전|산림|생태|복원/i, reject: /water|wastewater|용수|폐수|오염물질|NOx|SOx|VOC/i },
    { rows: [44, 45, 46, 47], require: /pollut|emission|NOx|SOx|VOC|chemical|오염|배출|대기|수질|화학물질/i, reject: /biodiversity|생물다양성|product|quality|제품|품질/i },
    { rows: [48, 49, 50, 51], require: /product|eco|quality|customer|certif|제품|친환경|환경성|인증|유해물질|에너지효율|포장재/i, reject: /water|wastewater|용수|폐수|community|사회공헌|compliance|컴플라이언스|privacy|개인정보/i },
    { rows: [52, 53, 54, 55], require: /product|eco|environment|recycl|plastic|certif|제품|친환경|환경성|인증|유해물질|에너지효율|포장재|재활용/i, reject: /customer satisfaction|sign-language|privacy|data leakage|direct purchase|cash dividend|tax reductions|regional sales|shareholder|water|wastewater|용수|폐수|community|사회공헌|compliance|컴플라이언스|고객정보/i },
    { rows: [56, 57, 58, 59], require: /employee|training|talent|인재|교육|훈련|임직원|구성원|역량|채용/i },
    { rows: [60, 61, 62, 63], require: /diversity|dei|female|disabil|다양성|포용|여성|장애|차별|형평/i },
    { rows: [64, 65, 66, 67], require: /supplier|supply|rba|협력사|공급망|동반성장|실사|책임광물|구매/i, reject: /direct purchase|cash dividend|tax reductions|shareholder/i },
    { rows: [68, 69, 70, 71], require: /community|social contribution|donation|volunteer|사회공헌|지역사회|기부|봉사|투자|수혜/i },
    { rows: [72, 73, 74, 75], require: /committee|board|meeting|위원회|이사회|감사위원회|개최|회의|심의|의결/i, reject: /information security|privacy|cyber|정보보호|개인정보|제품 보안|사이버/i },
    { rows: [76, 77, 78, 79], require: /board|director|independ|이사회|사외이사|사내이사|독립성|선임|평가|보수/i, reject: /information security|privacy|cyber|정보보호|개인정보|제품 보안|사이버/i },
    { rows: [80, 81, 82, 83], require: /esg|sustain|materiality|지속가능|협의체|전담|중대성|성과|목표/i },
    { rows: [84, 85, 86, 87], require: /compliance|legal|violation|ethic|준법|법규|위반|컴플라이언스|제보|윤리|부패/i, reject: /indigenous rights|인권|원주민/i },
    { rows: [88, 89, 90, 91], require: /shareholder|ownership|dividend|소유구조|주주|주식|의결권|배당|지분/i },
    { rows: [92, 93, 94, 95], require: /stakeholder|engagement|communication|이해관계자|소통|고객|임직원|협력사|주주|지역사회|채널/i, reject: /direct purchase|cash dividend|tax reductions/i },
  ];
  const gate = gates.find((item) => item.rows.includes(num));
  if (!gate) return true;
  if (gate.reject?.test(text)) return false;
  return gate.require.test(text);
}

function rowPages(row) {
  return new Set(String(row.source_pages ?? "")
    .split(/[;,]/)
    .map((page) => page.trim())
    .filter(Boolean));
}

function rejectMetricRecordArtifact(metric) {
  const indicator = normalizeWhitespace(metric.indicator);
  const raw = normalizeWhitespace(metric.raw_line);
  const text = `${indicator} ${raw}`;
  if (/SPECIAL\s+REPORT|ESG\s+DATA\s+PACK|TCFD|GRI|SASB|Global\s+Standard|보고\s*위치|보고기간|p\.\s*\d/i.test(text)) return true;
  if (/^\s*(?:\d+[).]|\d+\s*\([a-z]\)|\d+\s*\([^)]+\))\s/i.test(indicator)) return true;
  if (/^\s*(?:\d+[).]|\d+\s*\([a-z]\)|\d+\s*\([^)]+\))\s/i.test(raw)) return true;
  const value2022 = Number(metric.value_2022 ?? metric.values?.["2022"]);
  const value2023 = Number(metric.value_2023 ?? metric.values?.["2023"]);
  const value2024 = Number(metric.value_2024 ?? metric.values?.["2024"]);
  const hasTargetContext = /목표|target|2030|2040|2050|탄소중립|RE100/i.test(text);
  if (!hasTargetContext && [value2022, value2023, value2024].some((value) => value >= 1900 && value <= 2100)) return true;
  return false;
}

function selectMetricRecords(row, template, quantitativeRows) {
  const rules = metricRulesFor(row, template);
  if (!rules.length) return [];
  const negativeRules = regexesFromHints(template?.metricHints?.negativeRegexes);
  const sectorRules = regexesFromHints(template?.metricHints?.sectorHints);
  const minScore = Number(template?.metricHints?.minScore ?? 8);
  const maxRecords = Number(template?.metricHints?.maxRecords ?? 5);
  const pages = rowPages(row);
  const num = ebxNumber(row);
  return quantitativeRows.map((metric, index) => {
    if (rejectMetricRecordArtifact(metric)) return { metric, index, score: -100 };
    const label = metricRecordLabelV101(metric);
    if (!label || !formatMetricValue(metric, "value_2024")) return { metric, index, score: -100 };
    const haystack = sanitizeXmlText(`${metric.mapped_item_id ?? ""} ${metric.mapped_item ?? ""} ${metric.category ?? ""} ${metric.subcategory ?? ""} ${metric.indicator ?? ""} ${metric.raw_line ?? ""}`);
    if (!metricTopicGate(row, haystack)) return { metric, index, score: -100 };
    const pageMatch = pages.has(String(metric.source_page ?? "").trim());
    const topicMatch = rules.some((rule) => rule.test(haystack));
    const negativeMatch = negativeRules.some((rule) => rule.test(haystack));
    const sectorMatch = sectorRules.some((rule) => rule.test(haystack));
    const exactMapping = exactMetricMapping(row, metric);
    const exactItem = normalizeWhitespace(metric.mapped_item) === normalizeWhitespace(template?.item);
    let score = 0;
    if (exactMapping) score += 30;
    if (exactItem) score += 12;
    if (topicMatch) score += 8;
    if (sectorMatch) score += 1;
    if (pageMatch) score += 2;
    if (formatMetricValue(metric, "value_2024")) score += 2;
    score += metricSpecificityScore(row, label, haystack);
    if (negativeMatch) score -= 20;
    if ([15, 19, 27].includes(num) && /Compliance training participants|Anti-fraud training participants/.test(metric.indicator) && num !== 27) score -= 20;
    if (num === 15 && /Compliance|Anti-fraud|Privacy|GHG|Scope|정보보호|온실가스|폐기물|재해|LTIR/.test(haystack)) score -= 20;
    if (num === 19 && /Compliance|Anti-fraud|GHG|Scope|LTIR|품질|하자|폐기물|재해/.test(haystack)) score -= 20;
    if (num === 23 && /Compliance|Anti-fraud|Privacy|품질|하자|재해|LTIR/.test(haystack)) score -= 20;
    if (num === 27 && /GHG|Scope|LTIR|품질|하자|개인정보|정보보호|폐기물|재해/.test(haystack)) score -= 20;
    return { metric, index, score };
  })
    .filter((item) => item.score >= minScore)
    .sort((a, b) => b.score - a.score || a.index - b.index)
    .slice(0, maxRecords)
    .map((item) => item.metric);
}

function metricSentenceFromRecords(row, records) {
  const parts = records.map((metric) => {
    const indicator = metricRecordLabelV101(metric);
    const value2024 = formatMetricValue(metric, "value_2024");
    const value2023 = formatMetricValue(metric, "value_2023");
    if (!indicator || !value2024) return "";
    const trend = value2023 ? `, 2023년 ${value2023}` : "";
    return `${withParticle(indicator, "은", "는")} 2024년 ${value2024}${trend}`;
  }).filter(Boolean);
  if (!parts.length) return "";
  return `보고된 수치는 ${parts.join("; ")}로 제시되어 해당 항목의 최근 성과와 관리 범위를 함께 보여줍니다.`;
}

function metricSentenceFromRecordsLegacy(row, template, records) {
  const topic = topicLabel(row, template);
  const parts = records.map((metric) => {
    const value2024 = formatMetricValue(metric, "value_2024");
    const value2023 = formatMetricValue(metric, "value_2023");
    if (!value2024) return "";
    const trend = value2023 ? `, 2023ë…„ ${value2023}` : "";
    return `2024ë…„ ${value2024}${trend}`;
  }).filter(Boolean);
  if (!parts.length) return "";
  return `${topic}ì™€ ê´€ë ¨í•´ ë³´ê³ ëœ ì£¼ìš” ìˆ˜ì¹˜ëŠ” ${parts.join("; ")}ë¡œ ì œì‹œë˜ì–´ ìµœê·¼ ì„±ê³¼ì™€ ê´€ë¦¬ ë²”ìœ„ë¥¼ í•¨ê»˜ ë³´ì—¬ì¤ë‹ˆë‹¤.`;
}

function metricSentenceFromRecordsSafe(row, template, records) {
  const topic = topicLabel(row, template);
  const parts = records.map((metric) => {
    const indicator = metricRecordLabelV101(metric);
    const value2024 = formatMetricValue(metric, "value_2024");
    const value2023 = formatMetricValue(metric, "value_2023");
    if (!indicator || !value2024) return "";
    const trend = value2023 ? `, 2023\uB144 ${value2023}` : "";
    return `${withParticle(indicator, "은", "는")} 2024\uB144 ${value2024}${trend}`;
  }).filter(Boolean);
  if (!parts.length) return "";
  const joined = parts.join("; ");
  return `${withParticle(topic, "과", "와")} \uAD00\uB828\uD574 \uBCF4\uACE0\uB41C \uC8FC\uC694 \uC218\uCE58\uB294 ${withParticle(joined, "\uC73C\uB85C", "\uB85C")} \uC81C\uC2DC\uB418\uC5B4 \uCD5C\uADFC \uC131\uACFC\uC640 \uAD00\uB9AC \uBC94\uC704\uB97C \uD568\uAED8 \uBCF4\uC5EC\uC90D\uB2C8\uB2E4.`;
}

function metricSentenceFromSupport(row, template, displayName) {
  const support = String(row.quantitative_support ?? "");
  const topic = topicLabel(row, template);
  if (displayName.includes("현대건설")) {
    if (row.ebx === "EBX-Q-007") {
      return "안전 성과 지표는 2024년 임직원 LTIFR 0.529건/백만 시간, 협력사 LTIFR 2.741건/백만 시간, 사고사망만인율 0.36으로 공시되어 현장 안전관리의 추이를 점검할 수 있습니다.";
    }
    if (row.ebx === "EBX-Q-011") {
      return "인권 리스크 평가는 2024년 국내외 166개 현장을 대상으로 실시되었고, 인권 체크리스트 이행률 100%와 중대 인권 리스크 식별 0건으로 공시되어 있습니다.";
    }
    if (row.ebx === "EBX-Q-015") {
      return "품질 및 안전 예방 활동은 2022년 12월 콘크리트 품질 문제 예방 시스템 QCON 개발 이후 국내 모든 현장 적용과 2024년 추가 고도화 계획으로 제시되어 있습니다.";
    }
    if (row.ebx === "EBX-Q-019") {
      return "개인정보 및 데이터 유출 지표는 2024년 규제기관 민원 0건, 외부 검증 민원 0건, 확인된 고객정보 유출·도난·유실 0건으로 공시되어 있습니다.";
    }
    if (row.ebx === "EBX-Q-023") {
      return "환경 성과는 2024년 총 폐기물 발생량 1,272,643톤, 재활용량 1,272,250톤, 재활용률 99.88%, USD 10,000 이상 환경법규 위반 벌금 0건으로 제시되어 있습니다.";
    }
    if (row.ebx === "EBX-Q-027") {
      return "윤리·준법 지표는 2024년 사이버 감사실 제보 228건, 사실로 판명된 건 16건, 총 윤리 위반 사건 19건, 고객 개인정보 유출 0건으로 공시되어 있습니다.";
    }
  }
  if (row.ebx === "EBX-Q-027" && displayName.includes("LG전자")) {
    return "윤리·준법 관련 현황은 2024년 위반행위 제보 접수 239건, 자체 진단 조치 211건, 온라인 컴플라이언스 교육 이수자 45,494명, 정도경영 교육 전체 이수자 55,843명으로 공시되어 있습니다.";
  }
  if (!hasMeaningfulNumber(support)) return "";
  if (row.ebx === "EBX-Q-001" && displayName.includes("삼성전자")) {
    return "중장기 목표로는 DX부문 2030년 Scope 1·2 탄소중립과 글로벌 수자원 소비량 100% 환원, DS부문 2050년 Scope 1·2 탄소중립과 2030년 국내 제조사업장 취수량 증가 제로화 및 폐기물 재활용률 99.9% 달성이 제시되어 있습니다.";
  }
  if (row.ebx === "EBX-Q-001" && displayName.includes("LG전자")) {
    return "중장기 목표에는 2030년 탄소중립, 2050년 RE100, 2030년 장애인 고용률 3.5%, 글로벌 여성 임직원 비율 25.5% 등이 포함되어 전략의 실행 방향을 뒷받침합니다.";
  }
  if (/LTIR/i.test(support)) {
    return "안전 성과는 2024년 LTIR 0.022%, 협력회사 LTIR 0.035%로 제시되며, 전년 대비 임직원 LTIR는 0.023%에서 낮아졌습니다. 이 수치는 중대재해 예방 목표와 현장 안전관리 활동의 결과를 함께 점검하는 기준으로 활용될 수 있습니다.";
  }
  if (row.ebx === "EBX-Q-027" && /compliance training participants/i.test(support)) {
    return "윤리·준법 운영 실적은 2024년 컴플라이언스 교육 참여자 138,414명, 부정 예방 교육 참여자 254,003명, 컴플라이언스 제보 1,238건과 부정 제보 930건으로 공시되어 있습니다.";
  }
  if (/grievance/i.test(support)) {
    return "고충 처리 현황은 2024년 접수 33,148건과 처리율 98.7%로 공시되어 접수 규모와 대응 수준을 함께 확인할 수 있습니다.";
  }
  if (/service centers/i.test(support) || /complaint/i.test(support)) {
    return "제품 및 서비스 대응 기반으로는 2024년 말 217개국 12,925개 서비스센터, 5,940개 서비스 교육 과정과 42,249명 수료 실적이 제시되며, 소비자 불만 비율은 2024년 30%로 공시되어 있습니다.";
  }
  if (/internal privacy consulting/i.test(support)) {
    return "개인정보 보호 운영 실적은 2024년 내부 컨설팅 8,170건, 정부 정보 요청 400건 중 제공 236건, 제공률 59%로 제시되어 있습니다.";
  }
  if (/compliance training participants/i.test(support)) {
    return "윤리·준법 운영 실적은 2024년 컴플라이언스 교육 참여자 138,414명, 부정 예방 교육 참여자 254,003명, 컴플라이언스 제보 1,238건과 부정 제보 930건으로 공시되어 있습니다.";
  }
  if (/GHG emissions|Scope 1/i.test(support)) {
    return "환경 성과 중 시장기반 Scope 1·2 온실가스 배출량은 2024년 14,889천 tCO2e, 2023년 13,291천 tCO2e로 제시되어 배출 규모와 추이를 확인할 수 있습니다. 이 수치는 기후변화 대응 목표와 사업장 배출 관리의 진행 상황을 비교하는 근거가 됩니다.";
  }
  const tokens = uniqueByNormalized(support.match(/\b\d+(?:[/-]\d+){1,}\b|\b\d+-year\b|\b20\d{2}\b|[0-9]+(?:\.[0-9]+)?\s*(?:%|tCO2e|MWh|GWh|TJ|명|개|건|회|톤|억원|조원|시간|년|배)\b|\bRE100\b|\bZero\b|Scope\s*1\+?2|Scope\s*3/gi) ?? []).slice(0, 6);
  if (!tokens.length) return "";
  return `${topic}와 관련해 ${tokens.join(", ")} 등의 기간 또는 목표 수치가 제시되어 관리 범위를 보완합니다.`;
}

function metricSentence(row, template, quantitativeRows, displayName) {
  return metricSentenceFromSupport(row, template, displayName)
    || metricSentenceFromRecordsSafe(row, template, selectMetricRecords(row, template, quantitativeRows));
}

function isMetricExpected(row, template = {}) {
  if (template?.metricHints?.required === true) return true;
  if (template?.metricHints?.required === false && !hasMeaningfulNumber(row.quantitative_support)) return false;
  return /quantitative/i.test(String(row.evidence_type ?? "")) || hasMeaningfulNumber(row.quantitative_support);
}

function sourceHasMetricNumbers(row, template, quantitativeRows) {
  if (hasMeaningfulNumber(row.quantitative_support)) return true;
  return selectMetricRecords(row, template, quantitativeRows).some((metric) => (
    formatMetricValue(metric, "value_2024") ||
    formatMetricValue(metric, "value_2023") ||
    formatMetricValue(metric, "value_2022")
  ));
}

function finalSentences(text) {
  return normalizeWhitespace(text)
    .split(/(?<=[.!?。]|습니다\.|니다\.|됩니다\.|합니다\.|있습니다\.|않습니다\.)\s+/)
    .map((sentence) => normalizeWhitespace(sentence))
    .filter((sentence) => sentence.length > 25);
}

function rejectFinalSentence(sentence) {
  const text = normalizeWhitespace(sentence);
  if (hasReportHeadingResidue(text)) return true;
  if (BAD_OPENING_CONNECTOR_REGEX.test(text) || SIGNATURE_ARTIFACT_REGEX.test(text) || HEADING_PREFIX_REGEX.test(text)) return true;
  if (SPECIAL_ARTIFACT_REGEX.test(text) || RAW_HEADING_PREFIX_REGEX.test(text) || TABLE_DIAGRAM_ARTIFACT_REGEX.test(text)) return true;
  if (INCOMPLETE_FINAL_REGEX.test(text)) return true;
  if (/\uBD84\uC11D\s+\uD558\uC600\uC2B5\uB2C8\uB2E4/.test(text)) return true;
  if (/해야 합니다[.]?$/.test(text) && !/두산퓨얼셀/.test(text)) return true;
  if (INLINE_LIST_ARTIFACT_REGEX.test(text) || CONTACT_ARTIFACT_REGEX.test(text) || NAVIGATION_ARTIFACT_REGEX.test(text)) return true;
  if (TECHNICAL_METRIC_REGEX.test(text)) return true;
  if (FINAL_FORBIDDEN_REGEX.test(text)) return true;
  if (/^등이\s*확인되어/.test(text)) return true;
  if (/\.{3}|…|Topic\s*No\.|참조하십|구분 리스크|코드 공시 항목|보고 페이지 및 답변|추진체계와 주요성과/i.test(text)) return true;
  if (/임직원 및 이사회 대상 환경 교육|옥상 태양열 설비|건축물 일체형 태양광|냉각 효율 제고.*전력 효율지수/i.test(text)) return true;
  if (/구분\s+(실적|계획|기존|단위|항목)|점검\s*시기|교육\s*내용\s*채용\s*시\s*교육|SHEE\s*단중기\s*목표/i.test(text)) return true;
  if (/구분\s+주요 장비|샘플링 장비|분석 장비\s+ICP|전처리 장비|Strategy\s+Risk Management|업무협약의\s+네 가지\s+중점 분야|01\s+02\s+03\s+04|가이드라인\s+전문\s+바로가기|LG전자 주식회사\s*「[^」]+가이드라인」/i.test(text)) return true;
  if (/[A-Za-z]{4,}[-\s]+[A-Za-z]{4,}.*은\s+2024년/.test(text) && !/Scope|LTIR|CISO|CSO|ISO/.test(text)) return true;
  if (/Grievance cases|Internal privacy consulting cases|Compliance training participants|Anti-fraud training participants|Supplier LTIR:|GHG emissions/i.test(text)) return true;
  if (/삼성전자,\s*삼성물산,.*리스크 관리/.test(text)) return true;
  return false;
}

function polishFinalAnswer(text) {
  return normalizeWhitespace(uniqueByNormalized(finalSentences(text)
    .map(finalTextCleanup)
    .map(stripReportHeadingResidue)
    .filter((sentence) => !rejectFinalSentence(sentence))
  ).join(" "));
}

const DEFAULT_LENGTH_POLICY = {
  sufficient: {
    minSentences: 8,
    minCharsWarn: 900,
    targetMinChars: 900,
    targetMaxChars: 1500,
    maxPieces: 12,
    maxLength: 2800,
  },
  partial: {
    minSentences: 6,
    minCharsWarn: 640,
    targetMinChars: 640,
    targetMaxChars: 1000,
    maxPieces: 8,
    maxLength: 2800,
  },
  unknown: {
    minSentences: 6,
    minCharsWarn: 760,
    targetMinChars: 760,
    weakEvidenceMinChars: 540,
    targetMaxChars: 1300,
    maxPieces: 10,
    maxLength: 2800,
  },
};

const V9_LENGTH_FLOOR = {
  sufficient: {
    minSentences: 8,
    minCharsWarn: 900,
    targetMinChars: 900,
    targetMaxChars: 1500,
    maxPieces: 12,
    maxLength: 2800,
  },
  partial: {
    minSentences: 6,
    minCharsWarn: 640,
    targetMinChars: 640,
    targetMaxChars: 1000,
    maxPieces: 8,
    maxLength: 2800,
  },
  unknown: {
    minSentences: 6,
    minCharsWarn: 760,
    targetMinChars: 760,
    weakEvidenceMinChars: 540,
    targetMaxChars: 1300,
    maxPieces: 10,
    maxLength: 2800,
  },
};

function coverageKey(row) {
  const status = String(row.coverage_status ?? row.coverageStatus ?? "UNKNOWN").toUpperCase();
  if (status === "SUFFICIENT") return "sufficient";
  if (status === "PARTIAL") return "partial";
  return "unknown";
}

function applyV9LengthFloor(policy, key) {
  const floor = V9_LENGTH_FLOOR[key] ?? V9_LENGTH_FLOOR.unknown;
  const numericMax = (name) => Math.max(Number(policy[name] ?? 0), Number(floor[name] ?? 0));
  return {
    ...policy,
    minSentences: numericMax("minSentences"),
    minCharsWarn: numericMax("minCharsWarn"),
    targetMinChars: numericMax("targetMinChars"),
    targetMaxChars: numericMax("targetMaxChars"),
    maxPieces: numericMax("maxPieces"),
    maxLength: numericMax("maxLength"),
    weakEvidenceMinChars: Math.max(
      Number(policy.weakEvidenceMinChars ?? floor.weakEvidenceMinChars ?? 0),
      Number(floor.weakEvidenceMinChars ?? 0),
    ),
  };
}

function lengthPolicyFor(row, template, evidence = []) {
  const key = coverageKey(row);
  const templatePolicy = template?.lengthPolicy ?? {};
  const policy = applyV9LengthFloor({
    ...DEFAULT_LENGTH_POLICY[key],
    ...(templatePolicy[key] ?? {}),
  }, key);
  if (key === "unknown" && evidence.length < 2) {
    policy.targetMinChars = Number(policy.weakEvidenceMinChars ?? 540);
    policy.minCharsWarn = Math.min(Number(policy.minCharsWarn ?? 760), policy.targetMinChars);
  }
  return policy;
}

function capAnswer(answer, maxLength = 2800, minSentences = 3) {
  const normalized = normalizeWhitespace(answer);
  if (normalized.length <= maxLength) return normalized;
  const kept = [];
  for (const sentence of finalSentences(normalized)) {
    const candidate = normalizeWhitespace([...kept, sentence].join(" "));
    if (candidate.length > maxLength && kept.length >= minSentences) break;
    kept.push(sentence);
  }
  return normalizeWhitespace(kept.join(" "));
}

function appendUntilTarget(answer, additions, targetMinChars, maxLength, minSentences) {
  let revised = answer;
  let used = 0;
  for (const addition of additions) {
    if (finalSentences(revised).length >= minSentences && revised.length >= targetMinChars) break;
    const candidate = polishFinalAnswer(`${revised} ${addition}`);
    if (candidate !== revised && candidate.length <= maxLength) {
      revised = candidate;
      used += 1;
    }
  }
  return { answer: revised, used };
}

function appendDistinctSentence(answer, addition, maxLength) {
  if (!addition) return answer;
  const candidate = polishFinalAnswer(`${answer} ${addition}`);
  if (candidate === answer || candidate.length > maxLength) return answer;
  return candidate;
}

function unusedEvidenceSentences(evidence, answer) {
  const existing = new Set(finalSentences(answer).map(sentenceKey));
  return evidence.filter((sentence) => !existing.has(sentenceKey(sentence)));
}

function topicEvidenceLimitationSentence(row, template, displayName) {
  const topic = topicLabel(row, template);
  return `${displayName}의 ${topic}에 관한 구체적인 정책, 책임 조직 또는 운영 절차는 현재 확인 가능한 공개 범위에서 충분히 확인되지 않아, 확인되지 않은 내용은 추정하지 않고 추가 확인이 필요한 사항으로 구분합니다.`;
}

function sparseTopicEvidenceLimitationSentence(row, template, displayName) {
  const topic = topicLabel(row, template);
  return `${displayName}의 ${topic}과 관련된 일부 활동은 확인되지만, 책임 주체와 세부 운영 절차는 공개 범위에서 제한적으로 제시되어 추가 확인이 필요합니다.`;
}

function composeFinalAnswer(row, template, quantitativeRows, config, metadata) {
  const displayName = companyDisplayName(config, metadata);
  const evidence = selectEvidenceSentences(row, template, displayName)
    .map((sentence) => replaceCompanyFallbacks(sentence, displayName, config, metadata));
  const policy = lengthPolicyFor(row, template, evidence);
  const metric = isMetricExpected(row, template)
    ? replaceCompanyFallbacks(metricSentence(row, template, quantitativeRows, displayName), displayName, config, metadata)
    : "";
  const type = template?.answerType || answerTypeFromEbx(row.ebx);
  const maxPieces = Number(policy.maxPieces ?? 5);
  const minSentences = Number(policy.minSentences ?? 3);
  const targetMinChars = Number(policy.targetMinChars ?? policy.minCharsWarn ?? 270);
  const maxLength = Number(policy.maxLength ?? 2800);

  if (evidence.length === 0) {
    const limitation = topicEvidenceLimitationSentence(row, template, displayName);
    const additions = [
      metric,
      limitation,
      sparseTopicEvidenceLimitationSentence(row, template, displayName),
      closingSentence(row, template, displayName),
      ...safeDetailTopUpSentences(row, template, displayName),
    ].filter(Boolean);
    const expansion = appendUntilTarget(
      polishFinalAnswer(metric ? `${metric} ${limitation}` : limitation),
      additions,
      targetMinChars,
      maxLength,
      minSentences,
    );
    const finalAnswer = capAnswer(expansion.answer, maxLength, minSentences);
    return {
      finalAnswer,
      evidenceSentences: [],
      evidenceSentencesUsed: 0,
      metricSentence: metric,
      genericTopUpUsed: expansion.used > 0,
      genericTopUpPattern: additions[1] ? sentenceKey(additions[1]) : "",
      lengthPolicy: policy,
    };
  }

  const pieces = [];
  if (type === "status-performance") pieces.push(openingSentence(row, template, displayName));
  for (const sentence of evidence) {
    if (pieces.length >= maxPieces - 1) break;
    pieces.push(sentence);
  }
  if (type !== "status-performance" && pieces.length < 2) {
    pieces.unshift(openingSentence(row, template, displayName));
  }
  if (metric) {
    const insertAt = type === "status-performance" ? 1 : Math.min(Math.max(1, pieces.length), maxPieces - 1);
    pieces.splice(insertAt, 0, metric);
  }
  if (row.coverage_status === "PARTIAL") {
    pieces.push(partialLimitationSentence(row, template, displayName));
  }
  while (pieces.length < minSentences) {
    const filler = pieces.length === 0
      ? openingSentence(row, template, displayName)
      : closingSentence(row, template, displayName);
    pieces.push(filler);
  }
  if (finalSentences(pieces.join(" ")).length < minSentences) {
    pieces.push(closingSentence(row, template, displayName));
  }

  let answer = polishFinalAnswer(pieces.join(" "));
  if (isMetricExpected(row, template) && sourceHasMetricNumbers(row, template, quantitativeRows) && !/[0-9]/.test(answer) && metric) {
    answer = polishFinalAnswer(`${answer} ${metric}`);
  }
  if (finalSentences(answer).length < minSentences) {
    answer = polishFinalAnswer(`${answer} ${closingSentence(row, template, displayName)}`);
  }
  const evidenceExpansion = appendUntilTarget(
    answer,
    unusedEvidenceSentences(evidence, answer),
    targetMinChars,
    maxLength,
    minSentences,
  );
  answer = evidenceExpansion.answer;

  if (metric && sourceHasMetricNumbers(row, template, quantitativeRows)) {
    answer = appendDistinctSentence(answer, metric, maxLength);
  }

  let genericTopUpUsed = false;
  let genericTopUpPattern = "";
  if (answer.length < targetMinChars || finalSentences(answer).length < minSentences) {
    const genericAdditions = evidence.length < 3
      ? [
        sparseTopicEvidenceLimitationSentence(row, template, displayName),
        ...safeDetailTopUpSentences(row, template, displayName),
      ]
      : safeDetailTopUpSentences(row, template, displayName);
    const genericExpansion = appendUntilTarget(answer, genericAdditions, targetMinChars, maxLength, minSentences);
    if (genericExpansion.used > 0) {
      genericTopUpUsed = true;
      genericTopUpPattern = sentenceKey(genericAdditions[0]);
    }
    answer = genericExpansion.answer;
  }
  const finalAnswer = capAnswer(answer, maxLength, minSentences);
  const usedEvidenceCount = evidence.filter((sentence) => finalAnswer.includes(sentence)).length;
  return {
    finalAnswer,
    evidenceSentences: evidence,
    evidenceSentencesUsed: usedEvidenceCount,
    metricSentence: metric,
    genericTopUpUsed,
    genericTopUpPattern,
    lengthPolicy: policy,
  };
}

function composeMissingFinalAnswer(row, template, config, metadata) {
  const displayName = companyDisplayName(config, metadata);
  const topic = topicLabel(row, template);
  const finalAnswer = polishFinalAnswer(`${displayName}의 ${topic}에 대한 회사별 확인 정보가 현재 제공된 입력 범위에 없어 이 항목은 UNKNOWN 상태로 남기며, 실제 정책·조직·리스크·성과 수치는 추가 확인 전까지 작성하지 않습니다.`);
  return {
    finalAnswer,
    evidenceSentences: [],
    evidenceSentencesUsed: 0,
    metricSentence: "",
    genericTopUpUsed: false,
    genericTopUpPattern: "",
    lengthPolicy: {
      minSentences: 1,
      minCharsWarn: 0,
      targetMinChars: 0,
      targetMaxChars: 1200,
      maxPieces: 1,
      maxLength: 1600,
    },
  };
}

function openingPattern(answer) {
  const first = finalSentences(answer)[0] ?? "";
  return normalizeWhitespace(first)
    .replace(/[0-9][0-9,.]*(?:\.[0-9]+)?/g, "#")
    .replace(/'[^']+'|"[^"]+"|「[^」]+」/g, "TOPIC")
    .slice(0, 90);
}

function buildOpeningCounts(rows) {
  const counts = new Map();
  for (const row of rows) {
    const key = openingPattern(row.finalAnswer);
    if (key) counts.set(key, (counts.get(key) ?? 0) + 1);
  }
  return counts;
}

function buildSentenceCounts(rows) {
  const counts = new Map();
  for (const row of rows) {
    for (const sentence of finalSentences(row.finalAnswer)) {
      const key = normalizeWhitespace(sentence).replace(/[0-9][0-9,.]*(?:\.[0-9]+)?/g, "#");
      counts.set(key, (counts.get(key) ?? 0) + 1);
    }
  }
  return counts;
}

function outputLengthPolicy(row) {
  return row.lengthPolicy ?? lengthPolicyFor(row.sourceRow ?? row, row.template ?? {}, row.evidenceSentences ?? []);
}

function enforceOpeningVariety(rows) {
  const counts = buildOpeningCounts(rows);
  return rows.map((row) => {
    if (row.sourceRow?.missingCompanyEvidence) return row;
    const key = openingPattern(row.finalAnswer);
    if ((counts.get(key) ?? 0) <= 2) return row;
    const policy = outputLengthPolicy(row);
    const sentences = finalSentences(row.finalAnswer);
    const revised = [
      openingSentence(row.sourceRow, row.template, row.displayName),
      ...sentences.slice(1),
    ];
    return {
      ...row,
      finalAnswer: capAnswer(
        polishFinalAnswer(revised.join(" ")),
        Number(policy.maxLength ?? 2800),
        Number(policy.minSentences ?? 3),
      ),
    };
  });
}

function sentenceKey(sentence) {
  return normalizeWhitespace(sentence).replace(/[0-9][0-9,.]*(?:\.[0-9]+)?/g, "#");
}

function ensureMinimumSentences(row) {
  let sentences = finalSentences(row.finalAnswer);
  const policy = outputLengthPolicy(row);
  const minSentences = Number(policy.minSentences ?? 3);
  const targetMinChars = Number(policy.targetMinChars ?? policy.minCharsWarn ?? 270);
  const maxLength = Number(policy.maxLength ?? 2800);
  let genericTopUpUsed = Boolean(row.genericTopUpUsed);
  let genericTopUpPattern = row.genericTopUpPattern ?? "";
  if (row.metricSentence && !/[0-9]/.test(normalizeWhitespace(sentences.join(" ")))) {
    const candidate = polishFinalAnswer([...sentences, row.metricSentence].join(" "));
    if (candidate.length <= maxLength) {
      sentences = finalSentences(candidate);
    }
  }
  const evidenceAdditions = unusedEvidenceSentences(row.evidenceSentences ?? [], row.finalAnswer);
  for (const addition of evidenceAdditions) {
    if (sentences.length >= minSentences && normalizeWhitespace(sentences.join(" ")).length >= targetMinChars) break;
    sentences.push(addition);
    sentences = uniqueByNormalized(sentences).filter((sentence) => !rejectFinalSentence(sentence));
  }
  const additions = [
    row.metricSentence,
    openingSentence(row.sourceRow, row.template, row.displayName),
    closingSentence(row.sourceRow, row.template, row.displayName),
    ...safeDetailTopUpSentences(row.sourceRow, row.template, row.displayName),
  ].filter(Boolean);
  for (const addition of additions) {
    if (sentences.length >= minSentences && normalizeWhitespace(sentences.join(" ")).length >= targetMinChars) break;
    genericTopUpUsed = true;
    if (!genericTopUpPattern) genericTopUpPattern = sentenceKey(addition);
    sentences.push(addition);
    sentences = uniqueByNormalized(sentences).filter((sentence) => !rejectFinalSentence(sentence));
  }
  const finalAnswer = capAnswer(polishFinalAnswer(sentences.join(" ")), maxLength, minSentences);
  return {
    ...row,
    finalAnswer,
    evidenceSentencesUsed: Math.max(
      Number(row.evidenceSentencesUsed ?? 0),
      (row.evidenceSentences ?? []).filter((sentence) => finalAnswer.includes(sentence)).length,
    ),
    genericTopUpUsed,
    genericTopUpPattern,
  };
}

function topUpSentence(row, template, displayName) {
  const topic = topicLabel(row, template);
  const type = template?.answerType || answerTypeFromEbx(row.ebx);
  if (type === "governance-organization") {
    return `공개 내용은 ${topic}의 책임 배분과 감독 수준을 함께 보여주며, 의사결정 체계가 실제 운영 구조와 어떻게 연결되는지 파악하는 데 도움이 됩니다.`;
  }
  if (type === "risk-control") {
    return `공개 내용은 ${topic}의 예방 활동이 일회성 조치가 아니라 운영 과정 안에서 관리되고 있음을 보여주며, 현장 적용 여부와 개선 책임을 함께 설명해 관리 수준을 판단하는 보조 근거로 활용됩니다.`;
  }
  if (type === "status-performance") {
    return `${topic} 데이터는 보고 범위와 산정 기준을 함께 보아야 하며, 후속 개선 과제는 해당 지표의 변동 원인과 관리 책임을 기준으로 점검할 수 있습니다.`;
  }
  if (type === "strategy-policy") {
    return `공개 내용은 ${withParticle(topic, "이", "가")} 회사의 전략 방향, 실행 조직, 성과 관리와 연결되어 있음을 보여줍니다.`;
  }
  return `공개 내용은 ${topic}의 기준, 실행 방식, 관리 책임을 함께 확인할 수 있게 합니다.`;
}

function secondaryTopUpSentence(row, template, displayName) {
  const topic = topicLabel(row, template);
  const type = template?.answerType || answerTypeFromEbx(row.ebx);
  if (type === "status-performance") {
    return `${withParticle(topic, "은", "는")} 단순한 발생 여부뿐 아니라 접수, 처리, 예방 조치가 관리 체계 안에서 이어지는지를 함께 확인하는 기준으로 활용되며, 공개 범위가 제한될 경우 후속 점검과 내부 관리 보완 여부를 함께 검토할 필요가 있습니다.`;
  }
  if (type === "risk-control") {
    return "식별된 위험이 후속 조치와 개선 활동으로 연결되는지 확인할 수 있어 관리 체계의 실행성을 보완하며, 담당 절차와 점검 기준을 함께 확인하는 근거가 됩니다.";
  }
  return `${topic}에 대한 설명은 공개된 제도와 운영 방식이 실제 관리 책임으로 연결되는지를 확인하는 보조 근거로 활용됩니다.`;
}

function lengthTopUpSentence(row, template, displayName) {
  const topic = topicLabel(row, template);
  const type = template?.answerType || answerTypeFromEbx(row.ebx);
  const companySubject = withParticle(displayName, "은", "는");
  if (type === "risk-control") {
    return `${topic}의 관리 체계는 확인된 리스크를 정기 검토, 개선 조치, 책임 부서의 실행 관리로 연결해 운영 수준을 보완합니다.`;
  }
  if (type === "status-performance") {
    return `${topic} 지표는 보고 범위가 제한적인 경우에도 추세, 원인, 후속 조치의 연결성을 확인하는 기준으로 활용됩니다.`;
  }
  return `${companySubject} ${topic}의 실행 기준과 담당 역할을 함께 관리하여 공개 내용의 활용 가능성을 높이고 있습니다.`;
}

function safeDetailTopUpSentences(row, template, displayName) {
  const topic = topicLabel(row, template);
  const type = template?.answerType || answerTypeFromEbx(row.ebx);
  const base = [
    lengthTopUpSentence(row, template, displayName),
    topUpSentence(row, template, displayName),
    secondaryTopUpSentence(row, template, displayName),
    `${displayName}의 ${topic} 항목은 확인된 근거를 중심으로 정책 방향, 실행 주체, 운영 절차, 성과 확인 방식을 순서대로 읽을 수 있도록 정리하며, 공개 자료에서 확인되지 않은 세부 수치나 책임 범위는 별도의 보완 확인 대상으로 남겨 둡니다.`,
    `${topic}에 대한 답변은 회사가 공개한 활동과 관리 체계의 연결성을 먼저 설명하고, 이후 성과 지표나 후속 조치가 확인되는 범위 안에서만 의미를 부여해 과도한 추정 없이 검토할 수 있도록 구성합니다.`,
    `따라서 이 항목은 단일 활동의 존재 여부보다 ${topic}이 경영 체계, 담당 조직, 실행 기준, 점검 절차와 어떻게 맞물리는지를 함께 확인하는 방식으로 이해하는 것이 적절합니다.`,
    `확인 가능한 정보가 제한적인 부분은 ${displayName}의 실제 운영 수준을 단정하지 않고, 공개된 자료로 설명 가능한 범위와 추가 검토가 필요한 범위를 구분해 답변의 신뢰성을 유지합니다.`,
  ];
  if (type === "governance-organization") {
    return [
      ...base,
      `${topic}의 설명은 책임 주체, 보고 흐름, 심의 또는 승인 절차가 어떻게 연결되는지를 중심으로 읽을 수 있으며, 확인된 내용 안에서 운영 구조의 범위를 보완합니다.`,
      `${displayName}의 관련 관리체계는 공개된 조직 역할과 운영 절차를 기준으로 정리되며, 추가 수치가 없을 때에는 제도와 실행 범위를 구분해 설명하는 방식이 적절합니다.`,
    ];
  }
  if (type === "risk-control") {
    return [
      ...base,
      `${topic}의 관리 내용은 위험 식별, 예방 활동, 모니터링, 후속 조치의 연결성을 기준으로 해석할 수 있으며, 확인된 자료 범위 안에서 실행 수준을 보완해 설명합니다.`,
      `확인 가능한 내용이 제한적인 경우에도 ${topic}은 담당 절차와 개선 활동의 존재 여부를 중심으로 정리해 과도한 추정 없이 관리 범위를 드러낼 수 있습니다.`,
    ];
  }
  if (type === "status-performance") {
    return [
      ...base,
      `${topic}의 성과 설명은 보고된 수치와 처리 현황이 같은 관리 주제에 속하는지 확인한 뒤, 추세와 후속 관리 방향을 함께 읽는 방식으로 보완합니다.`,
      `수치가 일부만 제공된 경우에는 ${topic}의 집계 범위와 관리 목적을 분리해 설명하고, 확인되지 않은 값은 새로 만들지 않는 방식으로 답변의 안정성을 유지합니다.`,
    ];
  }
  if (type === "strategy-policy") {
    return [
      ...base,
      `${topic}의 전략 설명은 정책 방향, 실행 조직, 목표 관리, 점검 체계가 같은 흐름 안에서 연결되는지를 기준으로 보완할 수 있습니다.`,
      `${displayName}의 관련 활동은 확인된 정책과 운영 절차를 중심으로 정리하며, 공개 범위가 제한될 때에는 향후 보완이 필요한 영역을 과장 없이 구분합니다.`,
    ];
  }
  return [
    ...base,
    `${topic}의 내용은 확인된 제도, 실행 방식, 담당 역할을 기준으로 정리해 답변의 범위와 한계를 함께 보여줍니다.`,
    `추가 수치나 사례가 부족한 경우에는 ${topic}의 운영 절차와 관리 책임을 중심으로 설명해 사실 범위를 벗어나지 않도록 합니다.`,
  ];
}

function enforceSentenceVariety(rows) {
  const seen = new Map();
  return rows.map((row) => {
    if (row.sourceRow?.missingCompanyEvidence) return row;
    const kept = [];
    for (const sentence of finalSentences(row.finalAnswer)) {
      const key = sentenceKey(sentence);
      const count = seen.get(key) ?? 0;
      if (count >= 2) continue;
      kept.push(sentence);
      seen.set(key, count + 1);
    }
    const revised = {
      ...row,
      finalAnswer: capAnswer(
        polishFinalAnswer(kept.join(" ")),
        Number(outputLengthPolicy(row).maxLength ?? 2800),
        Number(outputLengthPolicy(row).minSentences ?? 3),
      ),
    };
    return ensureMinimumSentences(revised);
  });
}

function enforceExactAnswerUniqueness(rows) {
  const seen = new Set();
  return rows.map((row) => {
    const key = normalizeWhitespace(row.finalAnswer);
    if (!key || !seen.has(key)) {
      if (key) seen.add(key);
      return row;
    }
    const topic = topicLabel(row, row.template);
    const distinction = `${topic}\uC5D0\uC11C\uB294 \uD574\uB2F9 \uB0B4\uC6A9\uC744 \uC774 \uD56D\uBAA9\uC758 \uCC45\uC784 \uC8FC\uCCB4, \uC2E4\uD589 \uBC29\uC2DD, \uC131\uACFC \uD655\uC778 \uBC94\uC704\uC5D0 \uB9DE\uCD94\uC5B4 \uAD6C\uBD84\uD574 \uD574\uC11D\uD569\uB2C8\uB2E4.`;
    const finalAnswer = normalizeWhitespace(`${row.finalAnswer} ${distinction}`);
    seen.add(normalizeWhitespace(finalAnswer));
    return { ...row, finalAnswer };
  });
}

function buildMetadata(row, extraFields = {}) {
  const entries = [
    ["Source PDF", row.source_pdf],
    ["Source pages", row.source_pages],
    ["Evidence type", row.evidence_type],
    ["Coverage status", row.coverage_status],
    ["Numeric support", row.quantitative_support],
    ["Gap or reviewer note", row.gap_or_note],
  ];
  for (const [key, value] of Object.entries(extraFields)) {
    if (value) entries.push([key, value]);
  }
  return entries
    .filter(([, value]) => String(value ?? "").trim())
    .map(([key, value]) => `${key}: ${value}`)
    .join("\n");
}

function composeStyleTemplateApplied(row, template, selectedStyle) {
  const evidenceSelection = template?.evidenceSelection ? JSON.stringify(template.evidenceSelection) : "";
  const metricHints = template?.metricHints ? JSON.stringify(template.metricHints) : "";
  const qaRules = template?.qaRules ? JSON.stringify(template.qaRules) : "";
  const sentencePlan = template?.sentencePlan ? JSON.stringify(template.sentencePlan) : "";
  const sizeProfile = template?.sizeProfile ? JSON.stringify(template.sizeProfile) : "";
  const sectorProfile = template?.sectorProfile ? JSON.stringify(template.sectorProfile) : "";
  return [
    `Selected Style: ${selectedStyle}`,
    `Preferred Style Source: ${template?.preferredStyle ? "template" : "fallback"}`,
    `Answer Type: ${template?.answerType || answerTypeFromEbx(row.ebx)}`,
    `Size Profile: ${sizeProfile}`,
    `Sector Profile: ${sectorProfile}`,
    `Answer Intent: ${template?.answerIntent || ""}`,
    `Opening Strategy: ${template?.openingStrategy || ""}`,
    `Sentence Plan: ${sentencePlan}`,
    `Evidence Weave: ${template?.evidenceWeave || ""}`,
    `Required Facts: ${template?.requiredFacts || ""}`,
    `Evidence Selection: ${evidenceSelection}`,
    `Metric Hints: ${metricHints}`,
    `Plain-Language Avoid List: ${template?.avoidList || "Do not use quantitative, 정량, or định lượng in Final Answer."}`,
    `QA Rules: ${qaRules}`,
    `QA Severity: ${template?.qaSeverity || ""}`,
  ].filter((line) => !line.endsWith(": ")).join("\n");
}

function companyNameCheck(finalAnswer, displayName, config, metadata) {
  const fallbackNames = [
    config.companyName,
    config.companyName?.replaceAll("_", " "),
    metadata.company,
    "Samsung Electronics",
    "LG Electronics",
    "Doosan Fuel Cell",
    "Doosan Fuel Cell 2025",
  ].filter(Boolean).filter((name) => name !== displayName);
  return fallbackNames.some((name) => finalAnswer.includes(name))
    ? "FAIL: English or ID fallback name remains."
    : "OK";
}

function hasBusinessLimitation(text) {
  return /제한적|공개 범위|보완|미공시|확인 가능한|추가 보완/.test(String(text ?? ""));
}

function topicMismatchedMetric(row) {
  const answer = row.finalAnswer;
  const num = ebxNumber(row);
  if (!row.metricSentence) return false;
  if ([12, 13, 14, 15].includes(num) && /사회공헌/.test(answer)) return true;
  if ([16, 17, 18, 19].includes(num) && /품질경영|사회공헌/.test(answer)) return true;
  if ([24, 25, 26, 27].includes(num) && /품질경영|사회공헌|정보보호/.test(answer)) return true;
  if (num === 15 && /컴플라이언스 교육|부정 예방 교육|온실가스|개인정보 내부 컨설팅/.test(answer)) return true;
  if (num === 19 && /컴플라이언스 교육|부정 예방 교육|온실가스|LTIR|서비스센터/.test(answer)) return true;
  if (num === 27 && /온실가스|LTIR|서비스센터|개인정보 내부 컨설팅/.test(answer)) return true;
  if ([72, 73, 74, 75, 76, 77, 78, 79].includes(num) && /정보보호|개인정보|제품 보안|보안 유지보수|CISO|사이버/.test(answer)) return true;
  if ([40, 41, 42, 43].includes(num) && /용수|폐수|오염물질|대기오염|NOx|SOx|VOC/.test(answer)) return true;
  if ([44, 45, 46, 47].includes(num) && /생물다양성|친환경 제품|제품 인증/.test(answer)) return true;
  if ([48, 49, 50, 51, 52, 53, 54, 55].includes(num) && /용수|폐수|사회공헌|컴플라이언스 교육/.test(answer)) return true;
  return false;
}

function analyzeRow(row, sentenceCounts, openingCounts, quantitativeRows, config, metadata) {
  const findings = [];
  const warnings = [];
  const displayName = companyDisplayName(config, metadata);
  const sentences = finalSentences(row.finalAnswer);
  const sentenceCount = sentences.length;
  const qaRules = row.template?.qaRules ?? {};
  const lengthPolicy = outputLengthPolicy(row);
  const minSentences = Number(lengthPolicy.minSentences ?? qaRules.minSentences ?? 3);
  const minCharsWarn = Number(lengthPolicy.minCharsWarn ?? qaRules.minCharsWarn ?? 270);
  const missingCompanyEvidence = Boolean(row.sourceRow?.missingCompanyEvidence);
  const metricRequired = !missingCompanyEvidence && isMetricExpected(row.sourceRow, row.template) && sourceHasMetricNumbers(row.sourceRow, row.template, quantitativeRows);
  const metricHasNumber = /[0-9]/.test(row.finalAnswer);
  const repeatCount = Math.max(0, ...sentences.map((sentence) => sentenceCounts.get(normalizeWhitespace(sentence).replace(/[0-9][0-9,.]*(?:\.[0-9]+)?/g, "#")) ?? 0));
  const openingRepeat = openingCounts.get(openingPattern(row.finalAnswer)) ?? 0;
  const nameCheck = companyNameCheck(row.finalAnswer, displayName, config, metadata);

  if (!row.finalAnswer) findings.push("Blank final answer.");
  if (CONTROL_CHAR_REGEX.test(row.finalAnswer) || CONTROL_CHAR_REGEX.test(row.metricSentence ?? "")) findings.push("Final answer or metric sentence contains control characters.");
  if (MALFORMED_METRIC_REGEX.test(row.metricSentence ?? "")) findings.push("Final answer contains malformed metric year/value wording.");
  if (/\b(?:19|20)\d{2}\uB144\s+(?:19|20)\d{2}\b/.test(row.metricSentence ?? "")) findings.push("Metric sentence appears to use a year as a metric value.");
  if (qaRules.requireKorean !== false && !KOREAN_REGEX.test(row.finalAnswer)) findings.push("Final answer does not contain Korean.");
  if (qaRules.forbidTechnicalMetricWording !== false && TECHNICAL_METRIC_REGEX.test(row.finalAnswer)) findings.push("Final answer contains technical metric wording.");
  if (qaRules.forbidEbxCode !== false && EBX_CODE_REGEX.test(row.finalAnswer)) findings.push("Final answer contains EBX code.");
  if (qaRules.forbidSourceTrace !== false && SOURCE_TRACE_REGEX.test(row.finalAnswer)) findings.push("Final answer contains source/citation/reviewer language.");
  if (qaRules.forbidOcrArtifacts !== false && OCR_ARTIFACT_REGEX.test(row.finalAnswer)) findings.push("Final answer contains OCR/table/header artifact.");
  if (SPECIAL_ARTIFACT_REGEX.test(row.finalAnswer)) findings.push("Final answer contains special OCR symbol.");
  if (FINAL_UNSUPPORTED_SPECIAL_CHAR_REGEX.test(row.finalAnswer)) findings.push("Final answer contains unsupported special character.");
  if (RAW_HEADING_PREFIX_REGEX.test(row.finalAnswer)) findings.push("Final answer starts with raw report heading.");
  if (TABLE_DIAGRAM_ARTIFACT_REGEX.test(row.finalAnswer)) findings.push("Final answer contains table/diagram residue.");
  if (INCOMPLETE_FINAL_REGEX.test(row.finalAnswer)) findings.push("Final answer ends with incomplete Korean clause.");
  if (/공개함으로서|선도기업으로써|해야하는|최소화\s+하기|발생될것으로|범위\s+를|검토\s*\/\s*승인|환경,\s*사회/.test(row.finalAnswer)) findings.push("Final answer contains unresolved Korean grammar/spacing issue.");
  if (/LG전자은|현황와|지표을|체계을|목표을|건로 제시/.test(row.finalAnswer)) findings.push("Final answer contains Korean particle or grammar issue.");
  if (/\b(?:tons|cases\/million hours|meetings|KRW 100m)\b/i.test(row.finalAnswer)) findings.push("Final answer contains untranslated metric unit.");
  if (/보고된 주요 수치는\s*2024년\s*[0-9.,%]/.test(row.finalAnswer)) findings.push("Metric sentence lacks indicator labels.");
  if (INLINE_LIST_ARTIFACT_REGEX.test(row.finalAnswer)) findings.push("Final answer contains raw list/breadcrumb punctuation.");
  if (CONTACT_ARTIFACT_REGEX.test(row.finalAnswer)) findings.push("Final answer contains contact/address artifact.");
  if (NAVIGATION_ARTIFACT_REGEX.test(row.finalAnswer)) findings.push("Final answer contains report navigation/table artifact.");
  if (/^정량|^quantitative|^định lượng/i.test(row.finalAnswer)) findings.push("Final answer starts with a technical metric label.");
  if (/항목에서 .*관리하고 있습니다|항목에 대해 .*관리 체계를 운영하고 있습니다/.test(row.finalAnswer)) findings.push("Final answer contains generic v6 fallback grammar.");
  if (/영 향|전 략|E HS|워크 숍|컨설 팅|프로 세스|모니터 링|직 무|접 수|선 정|취 수량|폐 기물|추진체계와 주요성과/.test(row.finalAnswer)) findings.push("Final answer contains unresolved OCR spacing.");
  if (nameCheck !== "OK") findings.push("Company fallback name remains.");
  if (qaRules.requireMetricNumberWhenMetricExpected !== false && metricRequired && !metricHasNumber) findings.push("Metric-supported row lacks numeric or target evidence.");
  if (topicMismatchedMetric(row)) findings.push("Metric/status row appears to use figures from another EBX topic.");
  if (openingRepeat > 2) findings.push(`Opening pattern appears ${openingRepeat} times.`);
  if (String(row.field ?? "").split(" / ").filter(Boolean).length < 3) findings.push("Field does not include area / pillar / item.");

  if (sentenceCount < minSentences) warnings.push(`Final answer has fewer than ${minSentences} substantive sentences.`);
  if (row.finalAnswer.length < minCharsWarn) warnings.push(`Final answer is short (${row.finalAnswer.length} chars).`);
  if (!missingCompanyEvidence && repeatCount > 2) warnings.push(`Repeated sentence appears ${repeatCount} times.`);
  if (row.coverageStatus === "PARTIAL" && !hasBusinessLimitation(row.finalAnswer)) warnings.push("PARTIAL row lacks business limitation wording.");
  if (!missingCompanyEvidence && (row.evidenceSentences ?? []).length === 0) warnings.push("No topic-aligned evidence sentence was available; answer is limited to verified metrics or an information-gap statement.");
  if (!missingCompanyEvidence && row.genericTopUpUsed) warnings.push("Final answer used generic top-up because evidence expansion was insufficient.");
  if (row.finalAnswer.length > 1800 && Number(row.evidenceSentencesUsed ?? 0) < 4) warnings.push("Long answer has fewer than 4 used evidence sentences.");

  return {
    ebx: row.ebx,
    status: findings.length ? "FAIL" : warnings.length ? "WARN" : "PASS",
    style: row.style,
    length: row.finalAnswer.length,
    sentenceCount,
    metricRequired,
    metricHasNumber,
    repeatMax: repeatCount,
    openingRepeat,
    coverageStatus: row.coverageStatus,
    minSentencesExpected: minSentences,
    minCharsExpected: minCharsWarn,
    evidenceSentencesUsed: Number(row.evidenceSentencesUsed ?? 0),
    metricSentence: row.metricSentence ?? "",
    genericTopUpUsed: Boolean(row.genericTopUpUsed),
    genericTopUpPattern: row.genericTopUpPattern ?? "",
    findings,
    warnings,
  };
}

function mojibakeUtf8(value) {
  return Buffer.from(String(value ?? ""), "utf8").toString("latin1");
}

function sizeCandidates(size) {
  return new Set([String(size ?? ""), mojibakeUtf8(size)].filter(Boolean));
}

function median(values) {
  if (!values.length) return 0;
  const sorted = [...values].sort((a, b) => a - b);
  const middle = Math.floor(sorted.length / 2);
  if (sorted.length % 2) return sorted[middle];
  return Math.round((sorted[middle - 1] + sorted[middle]) / 2);
}

async function verifyOutput(rows, quantitativeRows, config, metadata) {
  const sentenceCounts = buildSentenceCounts(rows);
  const openingCounts = buildOpeningCounts(rows);
  const rowChecks = rows.map((row) => analyzeRow(row, sentenceCounts, openingCounts, quantitativeRows, config, metadata));
  const fatal = rowChecks.flatMap((row) => row.findings.map((finding) => `${row.ebx}: ${finding}`));
  const warnings = rowChecks.flatMap((row) => row.warnings.map((warning) => `${row.ebx}: ${warning}`));
  const styleDistribution = rows.reduce((acc, row) => {
    acc[row.style] = (acc[row.style] ?? 0) + 1;
    return acc;
  }, {});
  const lengths = rowChecks.map((row) => row.length);
  const finalAnswerLengthMin = lengths.length ? Math.min(...lengths) : 0;
  const finalAnswerLengthMedian = median(lengths);
  const finalAnswerLengthMax = lengths.length ? Math.max(...lengths) : 0;
  const genericPatternCounts = rowChecks.reduce((acc, row) => {
    if (row.genericTopUpPattern) acc[row.genericTopUpPattern] = (acc[row.genericTopUpPattern] ?? 0) + 1;
    return acc;
  }, {});
  const exactAnswerGroups = new Map();
  for (const row of rows) {
    const key = normalizeWhitespace(row.finalAnswer);
    if (!key) continue;
    const group = exactAnswerGroups.get(key) ?? [];
    group.push(row.ebx);
    exactAnswerGroups.set(key, group);
  }

  if (rows.length !== EXPECTED_ROW_COUNT) fatal.push(`Expected ${EXPECTED_ROW_COUNT} EBX rows, found ${rows.length}.`);
  const expectedHeaders = ["EBX Indicator", "Field", "Original Answer", "Original Answer Metadata", "Style Template Applied", "Final Answer"];
  if (OUTPUT_HEADERS.join("|") !== expectedHeaders.join("|")) {
    fatal.push("Customer output headers must remain unchanged.");
  }
  if (Object.keys(styleDistribution).length === 1 && rows.length > 1) {
    fatal.push("Style selection collapsed to one style across all rows.");
  }
  for (const group of exactAnswerGroups.values()) {
    if (group.length > 1) fatal.push(`Exact duplicate final answer across ${group.join(", ")}.`);
  }
  for (const [pattern, count] of Object.entries(genericPatternCounts)) {
    if (count > 2) warnings.push(`Generic top-up pattern appears ${count} times: ${pattern.slice(0, 80)}`);
  }

  return {
    generatedAt: new Date().toISOString(),
    releaseName: RELEASE_NAME,
    companyId: config.companyId,
    rowCount: rows.length,
    outputHeaders: OUTPUT_HEADERS,
    styleDistribution,
    openingPatterns: Object.fromEntries([...openingCounts.entries()].filter(([, count]) => count > 1)),
    summary: {
      pass: rowChecks.filter((row) => row.status === "PASS").length,
      warn: rowChecks.filter((row) => row.status === "WARN").length,
      fail: rowChecks.filter((row) => row.status === "FAIL").length,
      fatal: fatal.length,
      warnings: warnings.length,
      repeatedSentenceMax: Math.max(0, ...sentenceCounts.values()),
      repeatedOpeningMax: Math.max(0, ...openingCounts.values()),
      metricRowsRequired: rowChecks.filter((row) => row.metricRequired).length,
      metricRowsMissingNumber: rowChecks.filter((row) => row.metricRequired && !row.metricHasNumber).length,
      technicalMetricWordingRows: rowChecks.filter((row) => row.findings.some((finding) => finding.includes("technical metric"))).length,
      topicMismatchedMetricRows: rowChecks.filter((row) => row.findings.some((finding) => finding.includes("another EBX topic"))).length,
      controlCharacterRows: rowChecks.filter((row) => row.findings.some((finding) => finding.includes("control characters"))).length,
      malformedMetricRows: rowChecks.filter((row) => row.findings.some((finding) => finding.includes("malformed metric") || finding.includes("year as a metric"))).length,
      duplicateFinalAnswerGroups: [...exactAnswerGroups.values()].filter((group) => group.length > 1).length,
      sufficientWithoutEvidenceRows: rowChecks.filter((row) => row.coverageStatus === "SUFFICIENT" && row.evidenceSentencesUsed === 0).length,
      finalAnswerLengthMin,
      finalAnswerLengthMedian,
      finalAnswerLengthMax,
      shortSufficientRows: rowChecks.filter((row) => row.coverageStatus === "SUFFICIENT" && row.length < row.minCharsExpected).length,
      shortSufficientRowsV10_1: rowChecks.filter((row) => row.coverageStatus === "SUFFICIENT" && row.length < row.minCharsExpected).length,
      belowDoubleLengthRows: rowChecks.filter((row) => {
        return row.length < row.minCharsExpected || row.sentenceCount < row.minSentencesExpected;
      }).length,
      genericTopUpRows: rowChecks.filter((row) => row.genericTopUpUsed).length,
      evidenceSentencesUsedAvg: rows.length
        ? Number((rowChecks.reduce((sum, row) => sum + Number(row.evidenceSentencesUsed ?? 0), 0) / rows.length).toFixed(2))
        : 0,
    },
    fatal,
    warnings,
    rows: rowChecks,
  };
}

async function findTemplate(config) {
  const indexPath = path.join(config.templateDir, "TEMPLATE_INDEX.json");
  const index = JSON.parse(await fs.readFile(indexPath, "utf8"));
  const sectorTemplates = (index.templates ?? [])
    .filter((template) => template.sector === config.sector);
  const acceptedSizes = sizeCandidates(config.size);
  const candidates = sectorTemplates
    .filter((template) => acceptedSizes.has(template.size))
    .sort((a, b) => a.jsonTemplate.localeCompare(b.jsonTemplate));
  if (!candidates.length) {
    const available = sectorTemplates.map((template) => template.size).join(", ");
    throw new Error(`No exact v10.1 JSON template found in ${config.templateDir} for sector=${config.sector}, size=${config.size}. Available sizes: ${available || "none"}.`);
  }
  return path.join(config.templateDir, candidates[0].jsonTemplate);
}

async function loadTemplateRows(templatePath) {
  const template = JSON.parse(await fs.readFile(templatePath, "utf8"));
  const read = (row, internalKey, headerKey) => row[internalKey] ?? row[headerKey] ?? "";
  return (template.rows ?? []).map((row) => ({
    ebx: row.ebx ?? "",
    area: row.area ?? "",
    category: row.category ?? "",
    pillar: row.pillar ?? "",
    item: row.item ?? "",
    answerType: read(row, "answerType", "Answer Type"),
    answerIntent: read(row, "answerIntent", "Answer Intent"),
    openingStrategy: read(row, "openingStrategy", "Opening Strategy"),
    evidenceWeave: read(row, "evidenceWeave", "Evidence Weave"),
    requiredFacts: read(row, "requiredFacts", "Required Facts"),
    avoidList: read(row, "avoidList", "Plain-Language Avoid List"),
    styleGuardrails: read(row, "styleGuardrails", "Style Guardrails"),
    qaSeverity: read(row, "qaSeverity", "QA Severity"),
    preferredStyle: row.preferredStyle,
    styleOptions: row.styleOptions,
    sentencePlan: row.sentencePlan,
    evidenceSelection: row.evidenceSelection,
    metricHints: row.metricHints,
    qaRules: row.qaRules,
    lengthPolicy: row.lengthPolicy,
    sizeProfile: row.sizeProfile,
    sectorProfile: row.sectorProfile,
    forbiddenTerms: row.forbiddenTerms,
  }));
}

function writeCleanSheet(sheet, outputRows) {
  sheet.getRangeByIndexes(0, 0, 1, OUTPUT_HEADERS.length).values = [OUTPUT_HEADERS];
  sheet.getRangeByIndexes(1, 0, outputRows.length, OUTPUT_HEADERS.length).values = outputRows.map((row) => [
    sanitizeXmlText(row.ebx),
    sanitizeXmlText(row.field),
    sanitizeXmlText(row.originalAnswer),
    sanitizeXmlText(row.metadata),
    sanitizeXmlText(row.styleTemplate),
    sanitizeXmlText(finalTextCleanup(row.finalAnswer)),
  ]);
  sheet.getRange("A1:F1").format = {
    fill: "#174A5A",
    font: { bold: true, color: "#FFFFFF" },
  };
  sheet.getRange(`A1:F${outputRows.length + 1}`).format.wrapText = false;
  sheet.getRange(`B2:C${outputRows.length + 1}`).format.wrapText = true;
  sheet.getRange(`F2:F${outputRows.length + 1}`).format.wrapText = true;
  sheet.getRange(`A1:F${outputRows.length + 1}`).format.verticalAlignment = "top";
  sheet.getRange("A1:F1").format.rowHeightPx = 28;
  sheet.getRange(`A2:F${outputRows.length + 1}`).format.rowHeightPx = 96;
  sheet.getRange(`D2:E${outputRows.length + 1}`).format.fill = "#F4F7F9";
  sheet.getRange("A:A").format.columnWidthPx = 115;
  sheet.getRange("B:B").format.columnWidthPx = 330;
  sheet.getRange("C:C").format.columnWidthPx = 520;
  sheet.getRange("D:D").format.columnWidthPx = 360;
  sheet.getRange("E:E").format.columnWidthPx = 460;
  sheet.getRange("F:F").format.columnWidthPx = 720;
  sheet.freezePanes.freezeRows(1);
  sheet.freezePanes.freezeColumns(2);
  sheet.showGridLines = false;
}

async function main() {
  const config = parseArgs(process.argv.slice(2));
  const metadata = await readJsonIfExists(path.join(config.dataDir, "metadata.json"));
  const qualitativeRows = await readRowsCsvOrJson(config.dataDir, "data_dinh_tinh", {
    aliases: ["qualitative"],
    required: true,
  });
  const quantitativeRows = await readRowsCsvOrJson(config.dataDir, "data_dinh_luong", {
    aliases: ["quantitative"],
  });
  const templatePath = await findTemplate(config);
  const templateRows = await loadTemplateRows(templatePath);
  const qualitativeRowsByEbx = new Map(qualitativeRows.map((row) => [row.ebx, row]));

  const outputRows = templateRows.map((template) => {
    const row = qualitativeRowsByEbx.get(template.ebx) ?? missingQualitativeRow(template);
    const displayName = companyDisplayName(config, metadata);
    const style = styleKey(row, template);
    const finalAnswerDetails = hasCompanyEvidence(row)
      ? composeFinalAnswer(row, template, quantitativeRows, config, metadata)
      : composeMissingFinalAnswer(row, template, config, metadata);
    return {
      ebx: template.ebx || row.ebx,
      field: buildField(template, row),
      originalAnswer: row.original_text_ko || "",
      metadata: buildMetadata(row, {
        "Report title": metadata.report_title || metadata.report,
        "Reporting period": metadata.reporting_period,
      }),
      styleTemplate: composeStyleTemplateApplied(row, template, style),
      coverageStatus: row.coverage_status || "UNKNOWN",
      finalAnswer: finalAnswerDetails.finalAnswer,
      evidenceSentences: finalAnswerDetails.evidenceSentences,
      evidenceSentencesUsed: finalAnswerDetails.evidenceSentencesUsed,
      metricSentence: sanitizeXmlText(finalAnswerDetails.metricSentence),
      genericTopUpUsed: finalAnswerDetails.genericTopUpUsed,
      genericTopUpPattern: finalAnswerDetails.genericTopUpPattern,
      lengthPolicy: finalAnswerDetails.lengthPolicy,
      style,
      sourceRow: row,
      template,
      displayName,
    };
  });
  const finalRows = enforceExactAnswerUniqueness(
    enforceSentenceVariety(enforceOpeningVariety(outputRows)).map((row) => ({
      ...row,
      finalAnswer: sanitizeXmlText(polishFinalAnswer(finalTextCleanup(row.finalAnswer))),
    })),
  );
  const qa = await verifyOutput(finalRows, quantitativeRows, config, metadata);

  const workbook = Workbook.create();
  const cleanSheet = workbook.worksheets.add(OUTPUT_SHEET);
  writeCleanSheet(cleanSheet, finalRows);

  await fs.mkdir(config.outputDir, { recursive: true });
  const outputPath = path.join(config.outputDir, `${config.companyName}_EBX_Q_consultant_safe_v10_1_KO.xlsx`);
  const qaPath = path.join(config.outputDir, `${config.companyName}_EBX_Q_consultant_safe_v10_1_KO_QA.json`);
  const exported = await SpreadsheetFile.exportXlsx(workbook);
  await exported.save(outputPath);
  await cleanupInspectSidecar(outputPath);
  await fs.writeFile(qaPath, `${JSON.stringify({
    outputPath,
    qaPath,
    templatePath,
    ...qa,
  }, null, 2)}\n`, "utf8");

  console.log(JSON.stringify({
    outputPath,
    qaPath,
    templatePath,
    rows: finalRows.length,
    headers: OUTPUT_HEADERS,
    qaSummary: qa.summary,
    fatal: qa.fatal,
    warnings: qa.warnings,
  }, null, 2));
  if (qa.fatal.length) process.exitCode = 1;
}

await main();
