const fs = require('fs');
const CFB = require('cfb');
const XLSX = require('xlsx');

const hwpPath = 'source.hwp';
const outPath = '진정서_정리.xlsx';

const cfb = CFB.read(hwpPath, { type: 'file' });
const prv = CFB.find(cfb, 'Root Entry/PrvText');
if (!prv || !prv.content) throw new Error('PrvText stream not found');

const buf = prv.content;
let text = '';
for (let i = 0; i + 1 < buf.length; i += 2) {
  const ch = buf.readUInt16LE(i);
  if (ch === 13 || ch === 10) text += '\n';
  else if (ch === 9) text += '\t';
  else if (ch >= 32) text += String.fromCharCode(ch);
}

const pairRegex = /<([^<>]+)><([^<>]*)>/g;
const pairs = [];
let m;
while ((m = pairRegex.exec(text)) !== null) {
  pairs.push({ key: m[1].replace(/\s+/g, ' ').trim(), val: (m[2] || '').trim() });
}

const map = new Map();
for (const { key, val } of pairs) {
  if (!map.has(key) || (map.get(key) === '' && val !== '')) map.set(key, val);
}

function getFirst(keys) {
  for (const k of keys) {
    if (map.has(k)) return map.get(k);
  }
  return '';
}

const complainant = {
  '성명': getFirst(['성 명', '성명']),
  '주민등록번호': getFirst(['주민등록번호']),
  '주소': getFirst(['주 소', '주소']),
  '전화번호': getFirst(['전 화 번 호', '전화번호']),
  '휴대전화번호': getFirst(['휴대전화번호']),
  '전자우편주소': getFirst(['전자우편주소'])
};

const respondent = {
  '성명': getFirst(['성 명', '성명']),
  '연락처': getFirst(['연 락 처', '연락처']),
  '주소': getFirst(['주 소', '주소']),
  '사업장명': getFirst(['사 업 장 명', '사업장명']),
  '사업장 주소': getFirst(['사업장 주소 (실근무장소)', '사업장 주소']),
  '사업장전화번호': getFirst(['사업장전화번호']),
  '근로자 수': getFirst(['근로자 수'])
};

const claimSummary = [
  '이 문서는 근로 관련 진정을 접수하기 위한 양식으로,',
  '입사일·퇴사일, 체불임금총액·체불퇴직금액·기타체불금액,',
  '업무 내용, 임금 지급일, 근로계약방법(서면/구두), 퇴직 여부 등을 기재하여',
  '피진정인에 대한 임금 및 금품 체불 관련 사실을 신고하도록 구성되어 있습니다.'
].join(' ');

const wb = XLSX.utils.book_new();

const ws1 = XLSX.utils.json_to_sheet([complainant], { header: Object.keys(complainant) });
const ws2 = XLSX.utils.json_to_sheet([respondent], { header: Object.keys(respondent) });
const ws3 = XLSX.utils.aoa_to_sheet([
  ['진정요지'],
  [claimSummary]
]);

ws1['!cols'] = [
  { wch: 12 }, { wch: 16 }, { wch: 32 }, { wch: 16 }, { wch: 16 }, { wch: 28 }
];
ws2['!cols'] = [
  { wch: 12 }, { wch: 16 }, { wch: 28 }, { wch: 20 }, { wch: 30 }, { wch: 18 }, { wch: 10 }
];
ws3['!cols'] = [{ wch: 100 }];

XLSX.utils.book_append_sheet(wb, ws1, '진정인');
XLSX.utils.book_append_sheet(wb, ws2, '피진정인');
XLSX.utils.book_append_sheet(wb, ws3, '진정요지');

XLSX.writeFile(wb, outPath);

fs.writeFileSync('hwp_preview_extracted.txt', text, 'utf8');
console.log('created', outPath);
