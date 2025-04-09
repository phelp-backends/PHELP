const express = require("express");
const multer = require("multer");
const cors = require("cors");
const path = require("path");
const { exec } = require("child_process");
const fs = require("fs");

const app = express();
const PORT = 5000;

app.use(cors());
app.use(express.json());

// Serve uploaded and converted files
app.use("/uploads", express.static("uploads"));
app.use("/converted", express.static("converted"));

// Ensure folders exist
if (!fs.existsSync("uploads")) fs.mkdirSync("uploads");
if (!fs.existsSync("converted")) fs.mkdirSync("converted");

// Multer config
const storage = multer.diskStorage({
  destination: (req, file, cb) => cb(null, "uploads/"),
  filename: (req, file, cb) => cb(null, `${Date.now()}-${file.originalname}`),
});

const upload = multer({
  storage,
  limits: { fileSize: 10 * 1024 * 1024 }, // 10MB limit
  fileFilter: (req, file, cb) => {
    const allowed = ["application/vnd.openxmlformats-officedocument.wordprocessingml.document"];
    allowed.includes(file.mimetype)
      ? cb(null, true)
      : cb(new Error("Only .docx files allowed"));
  },
});

// Upload & Convert
app.post("/upload", upload.single("file"), (req, res) => {
  if (!req.file)
    return res.status(400).json({ success: false, message: "No file uploaded" });

  const inputPath = path.join(__dirname, "uploads", req.file.filename);
  const outputFileName = req.file.filename.replace(".docx", ".pptx");
  const outputPath = path.join(__dirname, "converted", outputFileName);

  const command = `python convert.py "${inputPath}" "${outputPath}"`;

  exec(command, (err, stdout, stderr) => {
    if (err) {
      console.error("❌ Conversion Error:", stderr);
      return res.status(500).json({
        success: false,
        message: "Conversion failed",
        error: stderr,
      });
    }

    res.status(200).json({
      success: true,
      message: "Converted successfully!",
      downloadUrl: `/converted/${outputFileName}`, // frontend uses this URL
      log: stdout.trim(),
    });
  });
});

app.listen(PORT, () =>
  console.log(`🚀 Server running at http://localhost:${PORT}`)
);
