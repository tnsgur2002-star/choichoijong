const fs = require("fs");
const path = require("path");
const CFB = require("cfb");
const XLSX = require("xlsx");
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

function escapeRegExp(str) {
  return String(str).replace(/[.*+?^${}()|[\]\\]/g, "\\$&");
}

function looseLabelPattern(label) {
  const compact = String(label || "").replace(/\s+/g, "");
  return compact
    .split("")
    .map((ch) => escapeRegExp(ch))
    .join("\\s*");
}

function compactText(text) {
  return String(text || "")
    .replace(/\r/g, "\n")
    .replace(/\t/g, " ")
    .replace(/\u00a0/g, " ")
    .replace(/[ ]{2,}/g, " ")
    .replace(/\n{3,}/g, "\n\n")
    .trim();
}

function extractValueByAliases(text, aliases, allAliases) {
  const source = compactText(text);
  if (!source) return "";

  const stopPattern = allAliases
    .map((x) => escapeRegExp(x))
    .sort((a, b) => b.length - a.length)
    .join("|");

  for (const alias of aliases) {
    const aliasEsc = escapeRegExp(alias);
    const re = new RegExp(
      `${aliasEsc}\\s*[:：]?\\s*([\\s\\S]{0,120}?)(?=\\s*(?:${stopPattern})\\s*[:：]?|\\n{2,}|$)`,
      "i"
    );
    const m = source.match(re);
    if (!m) continue;
    const value = String(m[1] || "")
      .replace(/^[\s:：\-·•]+/, "")
      .replace(/\s+/g, " ")
      .trim();
    if (value) return value;
  }

  return "";
}

function getFieldWithFallback(sectionMap, aliases, sectionText, allAliases) {
  const fromMap = getField(sectionMap, aliases);
  if (fromMap) return fromMap;
  return extractValueByAliases(sectionText, aliases, allAliases);
}

function extractLooseBetween(text, starts, ends, maxLen = 140) {
  const startPat = starts.map(looseLabelPattern).join("|");
  const endPat = ends.length ? ends.map(looseLabelPattern).join("|") : "";
  const re = new RegExp(
    `(?:${startPat})\\s*[:：]?\\s*([\\s\\S]{1,${maxLen}}?)` +
      (endPat ? `(?=(?:${endPat})|\\n{2,}|$)` : `(?=\\n{2,}|$)`),
    "i"
  );
  const m = String(text || "").match(re);
  if (!m) return "";
  return String(m[1] || "")
    .replace(/^[\s:：\-·•()]+/, "")
    .replace(/\s+/g, " ")
    .trim();
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
  const pdfParse = require("pdf-parse");
  const result = await pdfParse(raw);
  return String(result?.text || "");
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

function extractPatternFallbacks(text) {
  const src = compactText(text);

  const rrn = (src.match(/\b\d{6}-\d{7}\b/) || [])[0] || "";
  const email = (src.match(/\b[A-Z0-9._%+-]+@[A-Z0-9.-]+\.[A-Z]{2,}\b/i) || [])[0] || "";
  const phones = src.match(/\b0\d{1,2}-\d{3,4}-\d{4}\b/g) || [];
  const mobile = phones.find((p) => /^01\d-/.test(p)) || "";
  const landlines = phones.filter((p) => !/^01\d-/.test(p));
  const dates = src.match(/\b\d{4}[.\-\/]\s*\d{1,2}[.\-\/]\s*\d{1,2}\.?\b/g) || [];
  const bigNumbers = src.match(/\b\d{5,}\b/g) || [];

  return {
    rrn,
    email,
    mobile,
    landlines,
    dates,
    bigNumbers,
  };
}

function applyPatternFallbacks(data, text) {
  const f = extractPatternFallbacks(text);
  const src = String(text || "");

  if (!data.complainant["주민등록번호"] && f.rrn) data.complainant["주민등록번호"] = f.rrn;
  if (!data.complainant["전자우편주소"] && f.email) data.complainant["전자우편주소"] = f.email;
  if (!data.complainant["휴대전화번호"] && f.mobile) data.complainant["휴대전화번호"] = f.mobile;
  if (!data.complainant["전화번호"] && f.landlines[0]) data.complainant["전화번호"] = f.landlines[0];

  if (!data.respondent["연락처"] && f.landlines[1]) data.respondent["연락처"] = f.landlines[1];
  if (!data.respondent["사업장전화번호"]) {
    data.respondent["사업장전화번호"] = f.landlines[2] || f.landlines[1] || "";
  }

  const lines = String(text || "").split(/\n+/).map((x) => x.trim()).filter(Boolean);
  const claim = data.claimDetails;
  if (!claim.joinDate && f.dates[0]) claim.joinDate = f.dates[0];
  if (!claim.leaveDate && f.dates[1]) claim.leaveDate = f.dates[1];
  const amountCandidates = f.bigNumbers.filter((n) => {
    const only = String(n).replace(/\D/g, "");
    return only.length >= 8;
  });
  if (!claim.unpaidWageTotal && amountCandidates[0]) claim.unpaidWageTotal = amountCandidates[0];
  if (!claim.unpaidRetirementTotal && amountCandidates[1]) claim.unpaidRetirementTotal = amountCandidates[1];
  if (!claim.otherUnpaidTotal && amountCandidates[2]) claim.otherUnpaidTotal = amountCandidates[2];
  if (!claim.detailContent && lines.length) {
    claim.detailContent = lines.slice(0, 6).join(" ").slice(0, 300);
  }

  if (!data.complainant["성명"]) {
    data.complainant["성명"] =
      extractLooseBetween(src, ["성명", "성 명"], ["주민등록번호"], 50) ||
      "";
  }
  if (!data.complainant["주소"]) {
    data.complainant["주소"] =
      extractLooseBetween(src, ["주소", "주 소"], ["전화번호", "전 화 번 호", "휴대전화번호", "전자우편주소"], 120) ||
      "";
  }

  if (!data.respondent["성명"]) {
    data.respondent["성명"] =
      extractLooseBetween(src, ["피진정인 성명", "피진정인 성 명", "성명", "성 명"], ["연락처", "연 락 처"], 50) ||
      "";
  }
  if (!data.respondent["주소"]) {
    data.respondent["주소"] =
      extractLooseBetween(src, ["피진정인 주소", "주소", "주 소"], ["사업체구분", "사업장명", "사 업 장 명"], 120) ||
      "";
  }
  if (!data.respondent["사업장명"]) {
    data.respondent["사업장명"] =
      extractLooseBetween(src, ["사업장명", "사 업 장 명"], ["사업장주소", "사업장 주소", "사업장전화번호"], 100) ||
      "";
  }
  if (!data.respondent["사업장 주소"]) {
    data.respondent["사업장 주소"] =
      extractLooseBetween(src, ["사업장주소", "사업장 주소"], ["사업장전화번호", "근로자수"], 120) ||
      "";
  }
  if (!data.respondent["근로자 수"]) {
    data.respondent["근로자 수"] =
      extractLooseBetween(src, ["근로자수", "근로자 수"], ["진정 내용", "입사일"], 20) ||
      ((src.match(/(\d{1,3}(?:,\d{3})*|\d+)\s*명/) || [])[0] || "");
  }

  if (!claim.jobDescription) {
    claim.jobDescription =
      extractLooseBetween(src, ["업무내용", "업 무 내 용"], ["임금지급일", "근로계약방법", "내용"], 140) || "";
  }
  if (!claim.payday) {
    claim.payday =
      extractLooseBetween(src, ["임금지급일", "임금 지급일"], ["근로계약방법", "내용", "파일첨부"], 60) || "";
  }
}

function extractStructuredData(fullText) {
  const text = cleanText(fullText);
  const complainantSection = sectionRange(text, /1\.\s*진정인/, /2\.\s*피진정인/);
  const respondentSection = sectionRange(text, /2\.\s*피진정인/, /3\.\s*진정\s*내용/);
  const claimSection = sectionRange(text, /3\.\s*진정\s*내용/, /\(\s*.*고용노동\(지\)청장 귀하/);

  const complainantMap = parseBracketPairs(complainantSection);
  const respondentMap = parseBracketPairs(respondentSection);
  const claimMap = parseBracketPairs(claimSection);

  const complainantAliases = [
    "성명", "성 명", "주민등록번호", "주소", "주 소", "전화번호", "전 화 번 호", "휴대전화번호", "전자우편주소",
  ];
  const respondentAliases = [
    "성명", "성 명", "연락처", "연 락 처", "주소", "주 소", "사업장명", "사 업 장 명",
    "사업장 주소", "사업장 주소 (실근무장소)", "사업장전화번호", "근로자 수",
  ];
  const claimAliases = [
    "입사일", "입 사 일", "퇴사일", "퇴 사 일", "체불임금총액", "체불퇴직금총액", "체불퇴직금액",
    "기타체불금액", "업무내용", "업 무 내 용", "임금지급일", "임금 지급일", "내용", "내 용",
  ];

  const complainantText = complainantSection || text;
  const respondentText = respondentSection || text;
  const claimText = claimSection || text;

  const complainant = {
    성명: getFieldWithFallback(complainantMap, ["성명", "성 명"], complainantText, complainantAliases),
    주민등록번호: getFieldWithFallback(complainantMap, ["주민등록번호"], complainantText, complainantAliases),
    주소: getFieldWithFallback(complainantMap, ["주소", "주 소"], complainantText, complainantAliases),
    전화번호: getFieldWithFallback(complainantMap, ["전화번호", "전 화 번 호"], complainantText, complainantAliases),
    휴대전화번호: getFieldWithFallback(complainantMap, ["휴대전화번호"], complainantText, complainantAliases),
    전자우편주소: getFieldWithFallback(complainantMap, ["전자우편주소"], complainantText, complainantAliases),
  };

  const respondent = {
    성명: getFieldWithFallback(respondentMap, ["성명", "성 명"], respondentText, respondentAliases),
    연락처: getFieldWithFallback(respondentMap, ["연락처", "연 락 처"], respondentText, respondentAliases),
    주소: getFieldWithFallback(respondentMap, ["주소", "주 소"], respondentText, respondentAliases),
    사업장명: getFieldWithFallback(respondentMap, ["사업장명", "사 업 장 명"], respondentText, respondentAliases),
    "사업장 주소": getFieldWithFallback(
      respondentMap,
      ["사업장 주소", "사업장 주소 (실근무장소)"],
      respondentText,
      respondentAliases
    ),
    사업장전화번호: getFieldWithFallback(respondentMap, ["사업장전화번호"], respondentText, respondentAliases),
    "근로자 수": getFieldWithFallback(respondentMap, ["근로자 수"], respondentText, respondentAliases),
  };

  const claimDetails = {
    joinDate: getFieldWithFallback(claimMap, ["입사일", "입 사 일"], claimText, claimAliases),
    leaveDate: getFieldWithFallback(claimMap, ["퇴사일", "퇴 사 일"], claimText, claimAliases),
    unpaidWageTotal: getFieldWithFallback(claimMap, ["체불임금총액"], claimText, claimAliases),
    unpaidRetirementTotal: getFieldWithFallback(
      claimMap,
      ["체불퇴직금총액", "체불퇴직금액"],
      claimText,
      claimAliases
    ),
    otherUnpaidTotal: getFieldWithFallback(claimMap, ["기타체불금액"], claimText, claimAliases),
    jobDescription: getFieldWithFallback(claimMap, ["업무내용", "업 무 내 용"], claimText, claimAliases),
    payday: getFieldWithFallback(claimMap, ["임금지급일", "임금 지급일"], claimText, claimAliases),
    detailContent: getFieldWithFallback(claimMap, ["내용", "내 용"], claimText, claimAliases),
  };

  const result = {
    complainant,
    respondent,
    claimDetails,
    claimSummary: "",
    rawText: text,
  };

  applyPatternFallbacks(result, text);
  result.claimSummary = buildClaimSummary(result.claimDetails);

  return result;
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
