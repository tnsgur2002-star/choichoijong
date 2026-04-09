const fs = require("fs");
const path = require("path");
const CFB = require("cfb");
const XLSX = require("xlsx");
const pdfParse = require("pdf-parse");
const AdmZip = require("adm-zip");

const COMPLAINANT_FIELDS = [
  "성명",
  "주민등록번호",
  "주소",
  "전화번호",
  "휴대전화번호",
  "전자우편주소",
];

const RESPONDENT_FIELDS = [
  "성명",
  "연락처",
  "주소",
  "사업장명",
  "사업장 주소",
  "사업장전화번호",
  "근로자 수",
];

function normalizeKey(key) {
  return String(key || "")
    .replace(/\s+/g, " ")
    .trim();
}

function getField(sectionMap, aliases) {
  for (const alias of aliases) {
    const found = sectionMap.get(normalizeKey(alias));
    if (typeof found !== "undefined") return String(found).trim();
  }
  return "";
}

function sectionRange(text, startPattern, endPattern) {
  const start = text.search(startPattern);
  if (start < 0) return "";
  const end = endPattern ? text.slice(start + 1).search(endPattern) : -1;
  if (end < 0) return text.slice(start);
  return text.slice(start, start + 1 + end);
}

function parseBracketPairs(text) {
  const result = new Map();
  const re = /<([^<>]+)><([^<>]*)>/g;
  let m;
  while ((m = re.exec(text)) !== null) {
    const key = normalizeKey(m[1]);
    const val = String(m[2] || "").trim();
    if (!result.has(key) || (result.get(key) === "" && val !== "")) {
      result.set(key, val);
    }
  }
  return result;
}

function stripXmlToText(xml) {
  return xml
    .replace(/<w:br[^>]*\/>/gi, "\n")
    .replace(/<hp:lineBreak[^>]*\/>/gi, "\n")
    .replace(/<\/?(?:w:p|hp:p)[^>]*>/gi, "\n")
    .replace(/<[^>]+>/g, " ")
    .replace(/&lt;/g, "<")
    .replace(/&gt;/g, ">")
    .replace(/&amp;/g, "&")
    .replace(/\r/g, "")
    .replace(/[ \t]+\n/g, "\n")
    .replace(/\n{3,}/g, "\n\n");
}

function parseHwpPrvText(filePath) {
  const cfb = CFB.read(filePath, { type: "file" });
  const prv = CFB.find(cfb, "Root Entry/PrvText");
  if (!prv || !prv.content) {
    throw new Error("HWP의 PrvText 스트림을 찾지 못했습니다.");
  }

  const buf = prv.content;
  let text = "";
  for (let i = 0; i + 1 < buf.length; i += 2) {
    const ch = buf.readUInt16LE(i);
    if (ch === 13 || ch === 10) text += "\n";
    else if (ch === 9) text += "\t";
    else if (ch >= 32) text += String.fromCharCode(ch);
  }
  return text;
}

function parseHwpxText(filePath) {
  const zip = new AdmZip(filePath);
  const entries = zip
    .getEntries()
    .filter((e) => !e.isDirectory)
    .map((e) => e.entryName);

  const sectionNames = entries
    .filter((name) => /^Contents\/section\d+\.xml$/i.test(name))
    .sort((a, b) => a.localeCompare(b, "ko"));

  if (sectionNames.length === 0) {
    throw new Error("HWPX의 Contents/section*.xml 파일을 찾지 못했습니다.");
  }

  const textParts = sectionNames.map((name) => {
    const xml = zip.readAsText(name, "utf8");
    return stripXmlToText(xml);
  });
  return textParts.join("\n");
}

async function parsePdfText(filePath) {
  const raw = fs.readFileSync(filePath);
  const parsed = await pdfParse(raw);
  return String(parsed.text || "");
}

function cleanText(text) {
  return String(text || "")
    .replace(/\u0000/g, "")
    .replace(/\r/g, "")
    .replace(/[ \t]+\n/g, "\n")
    .replace(/\n{3,}/g, "\n\n")
    .trim();
}

function buildClaimSummary(details) {
  const v = (x) => {
    const s = String(x || "").trim();
    return s ? s : "미기재";
  };

  return [
    `진정인은 ${v(details.joinDate)}에 입사하여 ${v(details.leaveDate)}에 퇴사(또는 퇴사 예정)하였다고 진술하고 있습니다.`,
    `체불 금품 내역은 체불임금총액 ${v(details.unpaidWageTotal)}, 체불퇴직금총액 ${v(details.unpaidRetirementTotal)}, 기타체불금액 ${v(details.otherUnpaidTotal)}으로 기재되어 있습니다.`,
    `업무내용은 ${v(details.jobDescription)}이며, 임금지급일은 ${v(details.payday)}로 작성되어 있습니다.`,
    `진정 상세 내용은 다음과 같습니다: ${v(details.detailContent)}`,
  ].join("\n");
}

function extractStructuredData(fullText) {
  const text = cleanText(fullText);
  const complainantSection = sectionRange(text, /1\.\s*진정인/, /2\.\s*피진정인/);
  const respondentSection = sectionRange(text, /2\.\s*피진정인/, /3\.\s*진정\s*내용/);
  const claimSection = sectionRange(text, /3\.\s*진정\s*내용/, /\(\s*.*고용노동\(지\)청장 귀하/);

  const complainantMap = parseBracketPairs(complainantSection);
  const respondentMap = parseBracketPairs(respondentSection);
  const claimMap = parseBracketPairs(claimSection);

  const complainant = {
    성명: getField(complainantMap, ["성명", "성 명"]),
    주민등록번호: getField(complainantMap, ["주민등록번호"]),
    주소: getField(complainantMap, ["주소", "주 소"]),
    전화번호: getField(complainantMap, ["전화번호", "전 화 번 호"]),
    휴대전화번호: getField(complainantMap, ["휴대전화번호"]),
    전자우편주소: getField(complainantMap, ["전자우편주소"]),
  };

  const respondent = {
    성명: getField(respondentMap, ["성명", "성 명"]),
    연락처: getField(respondentMap, ["연락처", "연 락 처"]),
    주소: getField(respondentMap, ["주소", "주 소"]),
    사업장명: getField(respondentMap, ["사업장명", "사 업 장 명"]),
    "사업장 주소": getField(respondentMap, ["사업장 주소", "사업장 주소 (실근무장소)"]),
    사업장전화번호: getField(respondentMap, ["사업장전화번호"]),
    "근로자 수": getField(respondentMap, ["근로자 수"]),
  };

  const claimDetails = {
    joinDate: getField(claimMap, ["입사일", "입 사 일"]),
    leaveDate: getField(claimMap, ["퇴사일", "퇴 사 일"]),
    unpaidWageTotal: getField(claimMap, ["체불임금총액"]),
    unpaidRetirementTotal: getField(claimMap, ["체불퇴직금총액", "체불퇴직금액"]),
    otherUnpaidTotal: getField(claimMap, ["기타체불금액"]),
    jobDescription: getField(claimMap, ["업무내용", "업 무 내 용"]),
    payday: getField(claimMap, ["임금지급일", "임금 지급일"]),
    detailContent: getField(claimMap, ["내용", "내 용"]),
  };

  const claimSummary = buildClaimSummary(claimDetails);

  return {
    complainant,
    respondent,
    claimSummary,
    rawText: text,
  };
}

function writeWorkbook(outPath, data) {
  const wb = XLSX.utils.book_new();

  const ws1 = XLSX.utils.json_to_sheet([data.complainant], {
    header: COMPLAINANT_FIELDS,
  });
  const ws2 = XLSX.utils.json_to_sheet([data.respondent], {
    header: RESPONDENT_FIELDS,
  });
  const ws3 = XLSX.utils.aoa_to_sheet([["진정요지"], [data.claimSummary]]);

  ws1["!cols"] = [
    { wch: 12 },
    { wch: 16 },
    { wch: 32 },
    { wch: 16 },
    { wch: 16 },
    { wch: 28 },
  ];
  ws2["!cols"] = [
    { wch: 12 },
    { wch: 16 },
    { wch: 28 },
    { wch: 20 },
    { wch: 30 },
    { wch: 18 },
    { wch: 10 },
  ];
  ws3["!cols"] = [{ wch: 120 }];

  XLSX.utils.book_append_sheet(wb, ws1, "진정인");
  XLSX.utils.book_append_sheet(wb, ws2, "피진정인");
  XLSX.utils.book_append_sheet(wb, ws3, "진정요지");
  XLSX.writeFile(wb, outPath);
}

function getDefaultOutPath(inPath) {
  const dir = path.dirname(inPath);
  const base = path.basename(inPath, path.extname(inPath));
  return path.join(dir, `${base}_정리.xlsx`);
}

function getMissingFields(data) {
  return {
    complainant: COMPLAINANT_FIELDS.filter((k) => !String(data.complainant[k] || "").trim()),
    respondent: RESPONDENT_FIELDS.filter((k) => !String(data.respondent[k] || "").trim()),
  };
}

async function extractTextByType(filePath) {
  const ext = path.extname(filePath).toLowerCase();
  if (ext === ".hwp") return parseHwpPrvText(filePath);
  if (ext === ".hwpx") return parseHwpxText(filePath);
  if (ext === ".pdf") return parsePdfText(filePath);
  throw new Error(`지원하지 않는 파일 형식입니다: ${ext}`);
}

async function convertFile(inputPath, outputPath) {
  const inPath = path.resolve(inputPath);
  if (!fs.existsSync(inPath)) {
    throw new Error(`입력 파일이 없습니다: ${inPath}`);
  }

  const outPath = outputPath ? path.resolve(outputPath) : getDefaultOutPath(inPath);
  const text = await extractTextByType(inPath);
  const data = extractStructuredData(text);

  writeWorkbook(outPath, data);
  const debugPath = outPath.replace(/\.xlsx$/i, "_extracted.txt");
  fs.writeFileSync(debugPath, data.rawText, "utf8");

  return {
    outputPath: outPath,
    debugPath,
    missing: getMissingFields(data),
  };
}

module.exports = {
  COMPLAINANT_FIELDS,
  RESPONDENT_FIELDS,
  convertFile,
};
