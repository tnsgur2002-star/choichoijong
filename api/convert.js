const fs = require("fs");
const os = require("os");
const path = require("path");
const formidable = require("formidable");
const { convertFile } = require("../src/converter");

module.exports = async function handler(req, res) {
  if (req.method !== "POST") {
    res.status(405).json({ message: "Method Not Allowed" });
    return;
  }

  const form = formidable({
    multiples: false,
    uploadDir: os.tmpdir(),
    keepExtensions: true,
    maxFileSize: 30 * 1024 * 1024,
    filter: ({ originalFilename }) => {
      const ext = path.extname(originalFilename || "").toLowerCase();
      return [".hwp", ".hwpx", ".pdf"].includes(ext);
    },
  });

  let parsed;
  try {
    parsed = await new Promise((resolve, reject) => {
      form.parse(req, (err, fields, files) => {
        if (err) return reject(err);
        resolve({ fields, files });
      });
    });
  } catch (err) {
    res.status(400).json({ message: err.message || "업로드 파싱에 실패했습니다." });
    return;
  }

  const fileObj = parsed.files?.document;
  const file = Array.isArray(fileObj) ? fileObj[0] : fileObj;

  if (!file || !file.filepath) {
    res.status(400).json({ message: "업로드된 파일이 없습니다." });
    return;
  }

  const inputPath = file.filepath;
  const baseName = path.basename(file.originalFilename || "result", path.extname(file.originalFilename || ""));
  const outputPath = path.join(os.tmpdir(), `${Date.now()}_${Math.random().toString(36).slice(2, 8)}.xlsx`);
  const downloadName = `${baseName || "result"}_정리.xlsx`;

  try {
    await convertFile(inputPath, outputPath);
    const data = fs.readFileSync(outputPath);
    res.setHeader(
      "Content-Type",
      "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    );
    res.setHeader("Content-Disposition", `attachment; filename*=UTF-8''${encodeURIComponent(downloadName)}`);
    res.status(200).send(data);
  } catch (err) {
    res.status(400).json({ message: err.message || "변환에 실패했습니다." });
  } finally {
    fs.unlink(inputPath, () => {});
    fs.unlink(outputPath, () => {});
    fs.unlink(outputPath.replace(/\.xlsx$/i, "_extracted.txt"), () => {});
  }
};

module.exports.config = {
  api: {
    bodyParser: false,
  },
};
