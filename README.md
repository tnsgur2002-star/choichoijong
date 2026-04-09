# HWP/PDF -> XLSX 변환기 (진정서 양식)

`hwp`, `hwpx`, `pdf` 파일에서 텍스트를 추출해서 아래 3개 시트로 엑셀을 생성합니다.

- `진정인` 시트: 성명, 주민등록번호, 주소, 전화번호, 휴대전화번호, 전자우편주소
- `피진정인` 시트: 성명, 연락처, 주소, 사업장명, 사업장 주소, 사업장전화번호, 근로자 수
- `진정요지` 시트: 문단형 요약

## 설치

```bash
npm install
```

## CLI 실행

```bash
npm run convert -- "<입력파일>"
```

예시:

```bash
npm run convert -- "source.hwp"
npm run convert -- "sample.pdf" "sample_정리.xlsx"
```

## 출력

- 기본 출력 파일명: `<원본파일명>_정리.xlsx`
- 같은 경로에 `<원본파일명>_정리_extracted.txt`도 생성(추출 텍스트 확인용)

## 웹 실행 (드래그앤드롭 업로드)

```bash
npm run web
```

브라우저에서 아래 주소를 열면 됩니다.

```text
http://localhost:3000
```

웹에서는 파일 업로드 후 자동으로 변환된 XLSX 다운로드가 시작됩니다.

## Vercel 배포

- 루트 페이지: `index.html`
- 서버리스 API: `api/convert.js`
- 배포 설정: `vercel.json`

Vercel에서 이 저장소를 Import하면 `/api/convert` 함수와 정적 페이지가 함께 배포됩니다.

## 참고

- `.hwp`는 `PrvText` 기반으로 추출합니다. 문서 구조에 따라 일부 값이 비어 있을 수 있습니다.
- 스캔형 PDF(이미지 PDF)는 OCR이 없으면 텍스트 추출이 제한됩니다.
