const fs = require("fs");
const os = require("os");
const path = require("path");

module.exports = async function handler(req, res) {
  try {
    const formidableLib = require("formidable");
    const { convertFile } = require("../src/converter");

    if (req.method !== "POST") {
      res.status(405).json({ message: "Method Not Allowed" });
      return;
    }

    const formOptions = {
      multiples: false,
      uploadDir: os.tmpdir(),
      keepExtensions: true,
      maxFileSize: 30 * 1024 * 1024,
      filter: ({ originalFilename, mimetype }) => {
        const ext = path.extname(originalFilename || "").toLowerCase();
        if ([".hwp", ".hwpx", ".pdf"].includes(ext)) return true;
        if ((mimetype || "").toLowerCase() === "application/pdf") return true;
        return false;
      },
    };

    const form =
      typeof formidableLib.formidable === "function"
        ? formidableLib.formidable(formOptions)
        : new formidableLib.IncomingForm(formOptions);

    const parsed = await new Promise((resolve, reject) => {
      form.parse(req, (err, fields, files) => {
        if (err) {
          reject(err);
          return;
        }
        resolve({ fields, files });
      });
    });

    const fileObj =
      parsed.files?.document ||
      parsed.files?.file ||
      parsed.files?.upload ||
      Object.values(parsed.files || {})[0];
    const file = Array.isArray(fileObj) ? fileObj[0] : fileObj;

    if (!file || !file.filepath) {
      res.status(400).json({ message: "업로드된 파일이 없습니다." });
      return;
    }

    const inputPath = file.filepath;
    const baseName = path.basename(
      file.originalFilename || "result",
      path.extname(file.originalFilename || "")
    );
    const outputPath = path.join(
      os.tmpdir(),
      `${Date.now()}_${Math.random().toString(36).slice(2, 8)}.xlsx`
    );
    const downloadName = `${baseName || "result"}_정리.xlsx`;

    await convertFile(inputPath, outputPath);

    const data = fs.readFileSync(outputPath);
    res.setHeader(
      "Content-Type",
      "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    );
    res.setHeader(
      "Content-Disposition",
      `attachment; filename*=UTF-8''${encodeURIComponent(downloadName)}`
    );
    res.status(200).send(data);

    fs.unlink(inputPath, () => {});
    fs.unlink(outputPath, () => {});
    fs.unlink(outputPath.replace(/\.xlsx$/i, "_extracted.txt"), () => {});
  } catch (err) {
    console.error("api/convert error:", err);
    if (!res.headersSent) {
      res.status(500).json({ message: err.message || "서버 내부 오류가 발생했습니다." });
    }
  }
};

module.exports.config = {
  api: {
    bodyParser: false,
  },
};
