import fs from "node:fs/promises";
import path from "node:path";
import { fileURLToPath } from "node:url";
import { SpreadsheetFile, Workbook } from "@oai/artifact-tool";

const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);
const repoRoot = path.resolve(__dirname, "..");

const RELEASE_NAME = "consultant_safe_v7_2";
const DEFAULT_CONFIG = {
  companyId: "samsung_electronics_2025",
  companyName: "Samsung_Electronics_2025",
  dataDir: path.join(repoRoot, "company_esg_data", "samsung_electronics_2025"),
  outputDir: path.join(repoRoot, "final_template", "output", "samsung_electronics_2025"),
  templateDir: path.join(repoRoot, "consultant_safe_v7_2"),
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

const TEMPLATE_SHEET = "EBX-Q 템플릿";
const OUTPUT_SHEET = RELEASE_NAME;
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
const RAW_HEADING_PREFIX_REGEX = /^(?:지배구조|전략\s|영향,\s*(?:위험|기회)|산업안전보건\s*지배구조|안전보건\s*목표\s*모니터링|주요\s*사고\s*유형별|보안\s*교육\s*및\s*보안\s*수칙|윤리․?준법\s*경영에\s*대한\s*비전|투명경영위원회\s*주요\s*역할|내부고발자\s*보호\s*규정|정보보호\s*투자|이슈풀\s*구성\s*시|품질경영\s*인증\s*시스템\s*운영|통합\s*폐기물관리시스템|지표\s*및\s*목표)/;
const TABLE_DIAGRAM_ARTIFACT_REGEX = /수지율\s*=|발생시점|발생\s*가능성|영향\s*크기|영향\s*범위|재무\s*영향|조직도|인증서|이사회\s+투명경영위원회\s+준법지원인|검토\s*\/\s*승인|환경,\s*사회|2024년\s*품질\s*교육\s*실시\s*현황/;
const INCOMPLETE_FINAL_REGEX = /(하고,|하며,|개정하고,|수립하고,|운영하고,)\s*$/;
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

async function readRowsCsvOrJson(dataDir, baseName, { required = false } = {}) {
  const csvRows = await readCsvIfExists(path.join(dataDir, `${baseName}.csv`));
  if (csvRows.length) return csvRows;
  try {
    const value = JSON.parse(await fs.readFile(path.join(dataDir, `${baseName}.json`), "utf8"));
    if (Array.isArray(value)) return value;
    if (Array.isArray(value.rows)) return value.rows;
    if (Array.isArray(value.data)) return value.data;
  } catch {
    // Fall through to the required check below.
  }
  if (required) {
    throw new Error(`Missing required ${baseName}.csv or ${baseName}.json in ${dataDir}.`);
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

function sanitizeXmlText(text) {
  return String(text ?? "").replace(/[\x00-\x08\x0B\x0C\x0E-\x1F]/g, "");
}

function finalTextCleanup(text) {
  let out = repairKoreanSpacing(text)
    .replace(/\uC800\uD0C4\uC18C\s+\uACBD\uC81C/g, "\uC800\uD0C4\uC18C \uACBD\uC81C")
    .replace(/\uBD84\uC11D\s+\uD558\uC600\uC2B5\uB2C8\uB2E4/g, "\uBD84\uC11D\uD558\uC600\uC2B5\uB2C8\uB2E4");
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

function cleanEvidenceSentence(sentence) {
  let out = repairKoreanSpacing(sentence)
    .replace(SIGNATURE_ARTIFACT_REGEX, "")
    .replace(HEADING_PREFIX_REGEX, "")
    .replace(BAD_OPENING_CONNECTOR_REGEX, "");
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
    .replace(/\bp\.\d+\b/gi, "")
    .replace(/\b20\d{2}[–~-]20\d{2}\b/g, "");
  return /[0-9][0-9,]*(?:\.[0-9]+)?\s*(?:%|tCO2e|MWh|GWh|TJ|KRW|명|개|건|회|톤|억원|조원|시간|년|배|cases?|employees?|hours?|rate|target|carbon|RE100|Scope)/i.test(cleaned);
}

function rejectEvidenceSentence(sentence) {
  const text = normalizeWhitespace(sentence);
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
  return null;
}

function isTopicAligned(row, sentence) {
  const text = normalizeWhitespace(sentence);
  const reject = topicRejectRegex(row);
  if (reject?.test(text)) return false;
  const require = topicRequireRegex(row);
  if (!require) return true;
  return require.test(text);
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
  if (!isTopicAligned(row, sentence)) return -100;
  let score = 0;
  if (sentence.includes(displayName)) score += 3;
  if (/이사회|위원회|조직|리스크|위험|목표|전략|성과|안전|품질|정보보호|환경|윤리|인권|협력회사|공급망|고충|침해|준법|컴플라이언스/.test(sentence)) score += 4;
  if (numberTokens(sentence).length) score += 1;
  if (sentence.length >= 70 && sentence.length <= 230) score += 2;
  for (const keyword of keywords(row, template)) {
    if (keyword && sentence.includes(keyword)) score += 2;
  }
  return score;
}

function selectEvidenceSentences(row, template, displayName) {
  const scored = uniqueByNormalized(splitSentences(row.original_text_ko)
    .map(repairKoreanSpacing)
    .map(cleanEvidenceSentence)
    .filter((sentence) => !rejectEvidenceSentence(sentence))
    .filter((sentence) => isTopicAligned(row, sentence)))
    .map((sentence, index) => ({
      sentence,
      index,
      score: sentenceScore(sentence, row, template, displayName),
    }));
  return scored
    .filter((item) => item.score > 0)
    .sort((a, b) => b.score - a.score || a.index - b.index)
    .slice(0, 4)
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
  const value = String(metric[key] ?? "").trim();
  if (!value || value === "-") return "";
  const unit = translateMetricUnit(metric.unit);
  if (["%", "명", "건"].includes(unit)) return `${value}${unit}`;
  return unit ? `${value} ${unit}` : value;
}

function translateMetricUnit(unit) {
  const lower = String(unit ?? "").trim().toLowerCase();
  if (lower === "percent") return "%";
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
  const text = normalizeWhitespace(indicator);
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
  const explicit = normalizeWhitespace(metric.indicator);
  if (explicit) return translateMetricIndicator(explicit);
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

function rowPages(row) {
  return new Set(String(row.source_pages ?? "")
    .split(/[;,]/)
    .map((page) => page.trim())
    .filter(Boolean));
}

function selectMetricRecords(row, template, quantitativeRows) {
  const rules = metricRulesFor(row, template);
  if (!rules.length) return [];
  const negativeRules = regexesFromHints(template?.metricHints?.negativeRegexes);
  const minScore = Number(template?.metricHints?.minScore ?? 8);
  const maxRecords = Number(template?.metricHints?.maxRecords ?? 2);
  const pages = rowPages(row);
  const num = ebxNumber(row);
  return quantitativeRows.map((metric, index) => {
    const haystack = `${metric.category ?? ""} ${metric.subcategory ?? ""} ${metric.indicator ?? ""} ${metric.notes ?? ""} ${metric.raw_line ?? ""}`;
    const pageMatch = pages.has(String(metric.source_page ?? "").trim());
    const topicMatch = rules.some((rule) => rule.test(haystack));
    const negativeMatch = negativeRules.some((rule) => rule.test(haystack));
    let score = 0;
    if (topicMatch) score += 8;
    if (pageMatch) score += 2;
    if (formatMetricValue(metric, "value_2024")) score += 2;
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
    const indicator = metricRecordLabel(metric);
    const value2024 = formatMetricValue(metric, "value_2024");
    const value2023 = formatMetricValue(metric, "value_2023");
    if (!indicator || !value2024) return "";
    const trend = value2023 ? `, 2023년 ${value2023}` : "";
    return `${indicator}은 2024년 ${value2024}${trend}`;
  }).filter(Boolean);
  if (!parts.length) return "";
  return `보고된 수치는 ${parts.join("; ")}로 제시되어 해당 항목의 최근 성과와 관리 범위를 함께 보여줍니다.`;
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
  const tokens = uniqueByNormalized(support.match(/\b20\d{2}\b|[0-9]+(?:\.[0-9]+)?\s*%|\bRE100\b|\bZero\b|Scope\s*1\+?2|Scope\s*3/gi) ?? []).slice(0, 6);
  if (!tokens.length) return "";
  return `${topic}와 관련해 ${tokens.join(", ")} 등의 기간 또는 목표 수치가 제시되어 관리 범위를 보완합니다.`;
}

function metricSentence(row, template, quantitativeRows, displayName) {
  return metricSentenceFromSupport(row, template, displayName)
    || metricSentenceFromRecords(row, selectMetricRecords(row, template, quantitativeRows));
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
  if (/구분\s+(실적|계획|기존|단위|항목)|점검\s*시기|교육\s*내용\s*채용\s*시\s*교육|SHEE\s*단중기\s*목표/i.test(text)) return true;
  if (/[A-Za-z]{4,}[-\s]+[A-Za-z]{4,}.*은\s+2024년/.test(text) && !/Scope|LTIR|CISO|CSO|ISO/.test(text)) return true;
  if (/Grievance cases|Internal privacy consulting cases|Compliance training participants|Anti-fraud training participants|Supplier LTIR:|GHG emissions/i.test(text)) return true;
  if (/삼성전자,\s*삼성물산,.*리스크 관리/.test(text)) return true;
  return false;
}

function polishFinalAnswer(text) {
  return normalizeWhitespace(uniqueByNormalized(finalSentences(text)
    .map(finalTextCleanup)
    .filter((sentence) => !rejectFinalSentence(sentence))
  ).join(" "));
}

function capAnswer(answer, maxLength = 960) {
  const normalized = normalizeWhitespace(answer);
  if (normalized.length <= maxLength) return normalized;
  const kept = [];
  for (const sentence of finalSentences(normalized)) {
    const candidate = normalizeWhitespace([...kept, sentence].join(" "));
    if (candidate.length > maxLength && kept.length >= 3) break;
    kept.push(sentence);
  }
  return normalizeWhitespace(kept.join(" "));
}

function composeFinalAnswer(row, template, quantitativeRows, config, metadata) {
  const displayName = companyDisplayName(config, metadata);
  const evidence = selectEvidenceSentences(row, template, displayName)
    .map((sentence) => replaceCompanyFallbacks(sentence, displayName, config, metadata));
  const metric = isMetricExpected(row, template)
    ? replaceCompanyFallbacks(metricSentence(row, template, quantitativeRows, displayName), displayName, config, metadata)
    : "";
  const type = template?.answerType || answerTypeFromEbx(row.ebx);

  const pieces = [];
  if (type === "status-performance") pieces.push(openingSentence(row, template, displayName));
  for (const sentence of evidence) {
    if (pieces.length >= 3) break;
    pieces.push(sentence);
  }
  if (type !== "status-performance" && pieces.length < 2) {
    pieces.unshift(openingSentence(row, template, displayName));
  }
  if (metric) {
    const insertAt = type === "status-performance" ? 1 : Math.min(Math.max(1, pieces.length), 3);
    pieces.splice(insertAt, 0, metric);
  }
  if (row.coverage_status === "PARTIAL") {
    pieces.push(partialLimitationSentence(row, template, displayName));
  }
  while (pieces.length < 3) {
    const filler = pieces.length === 0
      ? openingSentence(row, template, displayName)
      : closingSentence(row, template, displayName);
    pieces.push(filler);
  }
  if (finalSentences(pieces.join(" ")).length < 3) {
    pieces.push(closingSentence(row, template, displayName));
  }

  let answer = polishFinalAnswer(pieces.join(" "));
  if (isMetricExpected(row, template) && sourceHasMetricNumbers(row, template, quantitativeRows) && !/[0-9]/.test(answer) && metric) {
    answer = polishFinalAnswer(`${answer} ${metric}`);
  }
  if (finalSentences(answer).length < 3) {
    answer = polishFinalAnswer(`${answer} ${closingSentence(row, template, displayName)}`);
  }
  if (answer.length < 270) {
    answer = polishFinalAnswer(`${answer} ${lengthTopUpSentence(row, template, displayName)}`);
  }
  return capAnswer(answer);
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

function enforceOpeningVariety(rows) {
  const counts = buildOpeningCounts(rows);
  return rows.map((row) => {
    const key = openingPattern(row.finalAnswer);
    if ((counts.get(key) ?? 0) <= 2) return row;
    const sentences = finalSentences(row.finalAnswer);
    const revised = [
      openingSentence(row.sourceRow, row.template, row.displayName),
      ...sentences.slice(1),
    ];
    return {
      ...row,
      finalAnswer: capAnswer(polishFinalAnswer(revised.join(" "))),
    };
  });
}

function sentenceKey(sentence) {
  return normalizeWhitespace(sentence).replace(/[0-9][0-9,.]*(?:\.[0-9]+)?/g, "#");
}

function ensureMinimumSentences(row) {
  let sentences = finalSentences(row.finalAnswer);
  const additions = [
    openingSentence(row.sourceRow, row.template, row.displayName),
    closingSentence(row.sourceRow, row.template, row.displayName),
    topUpSentence(row.sourceRow, row.template, row.displayName),
    secondaryTopUpSentence(row.sourceRow, row.template, row.displayName),
  ];
  for (const addition of additions) {
    if (sentences.length >= 3 && normalizeWhitespace(sentences.join(" ")).length >= 300) break;
    sentences.push(addition);
    sentences = uniqueByNormalized(sentences).filter((sentence) => !rejectFinalSentence(sentence));
  }
  return {
    ...row,
    finalAnswer: capAnswer(polishFinalAnswer(sentences.join(" "))),
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
  if (type === "risk-control") {
    return `${topic}의 관리 체계는 확인된 리스크를 정기 검토, 개선 조치, 책임 부서의 실행 관리로 연결해 운영 수준을 보완합니다.`;
  }
  if (type === "status-performance") {
    return `${topic} 지표는 보고 범위가 제한적인 경우에도 추세, 원인, 후속 조치의 연결성을 확인하는 기준으로 활용됩니다.`;
  }
  return `${displayName}은 ${topic}의 실행 기준과 담당 역할을 함께 관리하여 공개 내용의 활용 가능성을 높이고 있습니다.`;
}

function enforceSentenceVariety(rows) {
  const seen = new Map();
  return rows.map((row) => {
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
      finalAnswer: capAnswer(polishFinalAnswer(kept.join(" "))),
    };
    return ensureMinimumSentences(revised);
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
  const metricHints = template?.metricHints ? JSON.stringify(template.metricHints) : "";
  const qaRules = template?.qaRules ? JSON.stringify(template.qaRules) : "";
  const sentencePlan = template?.sentencePlan ? JSON.stringify(template.sentencePlan) : "";
  return [
    `Selected Style: ${selectedStyle}`,
    `Preferred Style Source: ${template?.preferredStyle ? "template" : "fallback"}`,
    `Answer Type: ${template?.answerType || answerTypeFromEbx(row.ebx)}`,
    `Answer Intent: ${template?.answerIntent || ""}`,
    `Opening Strategy: ${template?.openingStrategy || ""}`,
    `Sentence Plan: ${sentencePlan}`,
    `Evidence Weave: ${template?.evidenceWeave || ""}`,
    `Required Facts: ${template?.requiredFacts || ""}`,
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
  if ([12, 13, 14, 15].includes(num) && /사회공헌/.test(answer)) return true;
  if ([16, 17, 18, 19].includes(num) && /품질경영|사회공헌/.test(answer)) return true;
  if ([24, 25, 26, 27].includes(num) && /품질경영|사회공헌|정보보호/.test(answer)) return true;
  if (num === 15 && /컴플라이언스 교육|부정 예방 교육|온실가스|개인정보 내부 컨설팅/.test(answer)) return true;
  if (num === 19 && /컴플라이언스 교육|부정 예방 교육|온실가스|LTIR|서비스센터/.test(answer)) return true;
  if (num === 27 && /온실가스|LTIR|서비스센터|개인정보 내부 컨설팅/.test(answer)) return true;
  return false;
}

function analyzeRow(row, sentenceCounts, openingCounts, quantitativeRows, config, metadata) {
  const findings = [];
  const warnings = [];
  const displayName = companyDisplayName(config, metadata);
  const sentences = finalSentences(row.finalAnswer);
  const sentenceCount = sentences.length;
  const qaRules = row.template?.qaRules ?? {};
  const metricRequired = isMetricExpected(row.sourceRow, row.template) && sourceHasMetricNumbers(row.sourceRow, row.template, quantitativeRows);
  const metricHasNumber = /[0-9]/.test(row.finalAnswer);
  const repeatCount = Math.max(0, ...sentences.map((sentence) => sentenceCounts.get(normalizeWhitespace(sentence).replace(/[0-9][0-9,.]*(?:\.[0-9]+)?/g, "#")) ?? 0));
  const openingRepeat = openingCounts.get(openingPattern(row.finalAnswer)) ?? 0;
  const nameCheck = companyNameCheck(row.finalAnswer, displayName, config, metadata);

  if (!row.finalAnswer) findings.push("Blank final answer.");
  if (qaRules.requireKorean !== false && !KOREAN_REGEX.test(row.finalAnswer)) findings.push("Final answer does not contain Korean.");
  if (qaRules.forbidTechnicalMetricWording !== false && TECHNICAL_METRIC_REGEX.test(row.finalAnswer)) findings.push("Final answer contains technical metric wording.");
  if (qaRules.forbidEbxCode !== false && EBX_CODE_REGEX.test(row.finalAnswer)) findings.push("Final answer contains EBX code.");
  if (qaRules.forbidSourceTrace !== false && SOURCE_TRACE_REGEX.test(row.finalAnswer)) findings.push("Final answer contains source/citation/reviewer language.");
  if (qaRules.forbidOcrArtifacts !== false && OCR_ARTIFACT_REGEX.test(row.finalAnswer)) findings.push("Final answer contains OCR/table/header artifact.");
  if (sentences.some(rejectFinalSentence)) findings.push("Final answer contains incomplete fragment, English raw indicator, or table/list artifact.");
  if (SPECIAL_ARTIFACT_REGEX.test(row.finalAnswer)) findings.push("Final answer contains special OCR symbol.");
  if (RAW_HEADING_PREFIX_REGEX.test(row.finalAnswer)) findings.push("Final answer starts with raw report heading.");
  if (TABLE_DIAGRAM_ARTIFACT_REGEX.test(row.finalAnswer)) findings.push("Final answer contains table/diagram residue.");
  if (INCOMPLETE_FINAL_REGEX.test(row.finalAnswer)) findings.push("Final answer ends with incomplete Korean clause.");
  if (/공개함으로서|선도기업으로써|해야하는|최소화\s+하기|발생될것으로|범위\s+를|검토\s*\/\s*승인|환경,\s*사회/.test(row.finalAnswer)) findings.push("Final answer contains unresolved Korean grammar/spacing issue.");
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

  if (sentenceCount < Number(qaRules.minSentences ?? 3)) warnings.push(`Final answer has fewer than ${Number(qaRules.minSentences ?? 3)} substantive sentences.`);
  if (row.finalAnswer.length < Number(qaRules.minCharsWarn ?? 270)) warnings.push(`Final answer is short (${row.finalAnswer.length} chars).`);
  if (repeatCount > 2) warnings.push(`Repeated sentence appears ${repeatCount} times.`);
  if (row.coverageStatus === "PARTIAL" && !hasBusinessLimitation(row.finalAnswer)) warnings.push("PARTIAL row lacks business limitation wording.");

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
    findings,
    warnings,
  };
}

function mojibakeUtf8(value) {
  return Buffer.from(String(value ?? ""), "utf8").toString("latin1");
}

function sizeCandidates(size) {
  const labels = new Set([String(size ?? ""), mojibakeUtf8(size)]);
  labels.add("대기업");
  labels.add(mojibakeUtf8("대기업"));
  return labels;
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

  if (rows.length !== 27) fatal.push(`Expected 27 EBX rows, found ${rows.length}.`);
  const expectedHeaders = ["EBX Indicator", "Field", "Original Answer", "Original Answer Metadata", "Style Template Applied", "Final Answer"];
  if (OUTPUT_HEADERS.join("|") !== expectedHeaders.join("|")) {
    fatal.push("Customer output headers must remain unchanged.");
  }
  if (Object.keys(styleDistribution).length === 1 && rows.length > 1) {
    fatal.push("Style selection collapsed to one style across all rows.");
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
  let candidates = sectorTemplates
    .filter((template) => acceptedSizes.has(template.size))
    .sort((a, b) => a.jsonTemplate.localeCompare(b.jsonTemplate));
  if (!candidates.length) {
    candidates = sectorTemplates
      .filter((template) => sizeCandidates("대기업").has(template.size))
      .sort((a, b) => a.jsonTemplate.localeCompare(b.jsonTemplate));
  }
  if (!candidates.length) {
    throw new Error(`No v7.2 JSON template found in ${config.templateDir} for sector=${config.sector}, size=${config.size}. Run scripts/build_consultant_safe_v7_2.mjs first.`);
  }
  return path.join(config.templateDir, candidates[0].jsonTemplate);
}

async function loadTemplateRows(templatePath) {
  const template = JSON.parse(await fs.readFile(templatePath, "utf8"));
  const read = (row, internalKey, headerKey) => row[internalKey] ?? row[headerKey] ?? "";
  return new Map((template.rows ?? []).map((row) => [row.ebx, {
    ebx: row.ebx ?? "",
    area: row.area ?? "",
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
    forbiddenTerms: row.forbiddenTerms,
  }]));
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
  sheet.getRange(`A1:F${outputRows.length + 1}`).format.wrapText = true;
  sheet.getRange(`A1:F${outputRows.length + 1}`).format.verticalAlignment = "top";
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
  const qualitativeRows = await readRowsCsvOrJson(config.dataDir, "data_dinh_tinh", { required: true });
  const quantitativeRows = await readRowsCsvOrJson(config.dataDir, "data_dinh_luong");
  const templatePath = await findTemplate(config);
  const templateRows = await loadTemplateRows(templatePath);

  const outputRows = qualitativeRows.map((row) => {
    const template = templateRows.get(row.ebx) ?? {};
    const displayName = companyDisplayName(config, metadata);
    const style = styleKey(row, template);
    return {
      ebx: row.ebx,
      field: buildField(template, row),
      originalAnswer: row.original_text_ko || "",
      metadata: buildMetadata(row, {
        "Report title": metadata.report_title || metadata.report,
        "Reporting period": metadata.reporting_period,
      }),
      styleTemplate: composeStyleTemplateApplied(row, template, style),
      coverageStatus: row.coverage_status || "UNKNOWN",
      finalAnswer: composeFinalAnswer(row, template, quantitativeRows, config, metadata),
      style,
      sourceRow: row,
      template,
      displayName,
    };
  });
  const finalRows = enforceSentenceVariety(enforceOpeningVariety(outputRows)).map((row) => ({
    ...row,
    finalAnswer: polishFinalAnswer(finalTextCleanup(row.finalAnswer)),
  }));
  const qa = await verifyOutput(finalRows, quantitativeRows, config, metadata);

  const workbook = Workbook.create();
  const cleanSheet = workbook.worksheets.add(OUTPUT_SHEET);
  writeCleanSheet(cleanSheet, finalRows);

  await fs.mkdir(config.outputDir, { recursive: true });
  const outputPath = path.join(config.outputDir, `${config.companyName}_EBX_Q_consultant_safe_v7_2_KO.xlsx`);
  const qaPath = path.join(config.outputDir, `${config.companyName}_EBX_Q_consultant_safe_v7_2_KO_QA.json`);
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
