#!/usr/bin/env node

const path = require("path");
const { convertFile } = require("./converter");

function usage() {
  console.log("사용법:");
  console.log("  node src/index.js <입력파일(hwp|hwpx|pdf)> [출력파일.xlsx]");
  console.log("");
  console.log("예시:");
  console.log("  node src/index.js source.hwp");
  console.log("  node src/index.js sample.pdf sample_정리.xlsx");
}

function printMissing(missing) {
  if (!missing.complainant.length && !missing.respondent.length) return;
  console.log("\n비어있는 항목:");
  if (missing.complainant.length) {
    console.log(`- 진정인: ${missing.complainant.join(", ")}`);
  }
  if (missing.respondent.length) {
    console.log(`- 피진정인: ${missing.respondent.join(", ")}`);
  }
}

async function main() {
  const inArg = process.argv[2];
  const outArg = process.argv[3];
  if (!inArg) {
    usage();
    process.exit(1);
  }

  const inputPath = path.resolve(inArg);
  const outputPath = outArg ? path.resolve(outArg) : undefined;
  const result = await convertFile(inputPath, outputPath);

  console.log(`완료: ${result.outputPath}`);
  console.log(`추출원문: ${result.debugPath}`);
  printMissing(result.missing);
}

main().catch((err) => {
  console.error("오류:", err.message);
  process.exit(1);
});
