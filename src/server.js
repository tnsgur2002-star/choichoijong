const fs = require("fs");
const path = require("path");
const express = require("express");
const multer = require("multer");
const { convertFile } = require("./converter");

const app = express();
const PORT = Number(process.env.PORT || 3000);
const ROOT_DIR = path.resolve(__dirname, "..");
const TMP_DIR = path.join(ROOT_DIR, "tmp");

if (!fs.existsSync(TMP_DIR)) {
  fs.mkdirSync(TMP_DIR, { recursive: true });
}

const storage = multer.diskStorage({
  destination: (_req, _file, cb) => cb(null, TMP_DIR),
  filename: (_req, file, cb) => {
    const safe = `${Date.now()}_${Math.random().toString(36).slice(2, 8)}${path.extname(file.originalname)}`;
    cb(null, safe);
  },
});

const upload = multer({
  storage,
  limits: {
    fileSize: 30 * 1024 * 1024,
  },
  fileFilter: (_req, file, cb) => {
    const ext = path.extname(file.originalname).toLowerCase();
    if ([".hwp", ".hwpx", ".pdf"].includes(ext)) return cb(null, true);
    cb(new Error("지원 파일 형식은 .hwp, .hwpx, .pdf 입니다."));
  },
});

function cleanup(paths) {
  for (const p of paths) {
    if (!p) continue;
    fs.unlink(p, () => {});
  }
}

app.use(express.static(path.join(ROOT_DIR, "public")));

app.post("/api/convert", upload.single("document"), async (req, res) => {
  if (!req.file) {
    res.status(400).json({ message: "업로드된 파일이 없습니다." });
    return;
  }

  const inputPath = req.file.path;
  const baseName = path.basename(req.file.originalname, path.extname(req.file.originalname));
  const outputName = `${baseName}_정리.xlsx`;
  const outputPath = path.join(TMP_DIR, `${Date.now()}_${Math.random().toString(36).slice(2, 8)}.xlsx`);

  try {
    await convertFile(inputPath, outputPath);
    res.download(outputPath, outputName, (err) => {
      cleanup([inputPath, outputPath, outputPath.replace(/\.xlsx$/i, "_extracted.txt")]);
      if (err && !res.headersSent) {
        res.status(500).json({ message: "파일 다운로드 중 오류가 발생했습니다." });
      }
    });
  } catch (err) {
    cleanup([inputPath, outputPath, outputPath.replace(/\.xlsx$/i, "_extracted.txt")]);
    res.status(400).json({ message: err.message || "변환에 실패했습니다." });
  }
});

app.use((err, _req, res, _next) => {
  res.status(400).json({ message: err.message || "요청 처리 중 오류가 발생했습니다." });
});

app.listen(PORT, () => {
  console.log(`웹 변환기 실행: http://localhost:${PORT}`);
});
