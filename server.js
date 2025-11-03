// server.js
import express from "express";
import multer from "multer";
import ExcelJS from "exceljs";
import xlsx from "xlsx";
import fs from "fs";
import path from "path";
import { fileURLToPath } from "url";

const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);

const app = express();
const upload = multer({ dest: "uploads/" });

app.use(express.static("public"));

// Cột tương ứng các process
const processes = [
  { name: "Extrusion", colIndex: 50, suffix: "1" },
  { name: "UV", colIndex: 51, suffix: "2" },
  { name: "UV+Bigsheet", colIndex: 52, suffix: "3" },
  { name: "Profiling", colIndex: 53, suffix: "4" },
  { name: "Profiling+Bevel", colIndex: 54, suffix: "5" },
  { name: "Packaging", colIndex: 55, suffix: null },
  { name: "Profiling+Bevel+Packaging", colIndex: 56, suffix: null },
  { name: "Padding+Packaging", colIndex: 57, suffix: null }
];

app.post("/upload", upload.single("file"), async (req, res) => {
  try {
    if (!req.file) return res.status(400).send("Chưa upload file Excel!");

    const wb = xlsx.readFile(req.file.path);
    const sheetYC = wb.Sheets["YC XUẤT BOM"];
    const sheetBOM = wb.Sheets["Thông tin khai BOM"];

    if (!sheetYC || !sheetBOM) {
      return res.status(400).send("Không tìm thấy sheet YC XUẤT BOM hoặc Thông tin khai BOM!");
    }

    const dataYC = xlsx.utils.sheet_to_json(sheetYC, { header: 1, defval: "" });
    const dataBOM = xlsx.utils.sheet_to_json(sheetBOM, { header: 1, defval: "" });

    // ==== Lấy danh sách mã đầu 5 trong YC (bắt đầu từ C3) ====
    const listYC = new Set();
    for (let r = 2; r < dataYC.length; r++) {
      const row = dataYC[r] || [];
      const ma5 = row[2];
      if (ma5 && `${ma5}`.trim() !== "") listYC.add(`${ma5}`.trim());
    }

    // ==== Duyệt dữ liệu BOM (bắt đầu từ hàng 5) ====
    const results = [];
    const minVersionMap = {}; // { ma5: versionMin }

    // Lần 1: tìm version nhỏ nhất của từng mã đầu 5
    for (let r = 4; r < dataBOM.length; r++) {
      const row = dataBOM[r] || [];
      const ma5 = `${row[2] || ""}`.trim();
      const version = `${row[3] || ""}`.trim();
      if (!ma5 || !version) continue;
      if (!listYC.has(ma5)) continue;

      const verNum = parseInt(version.replace(/\D/g, "")) || 0;
      if (!minVersionMap[ma5] || verNum < minVersionMap[ma5]) {
        minVersionMap[ma5] = verNum;
      }
    }

    // Lần 2: xuất dữ liệu theo dấu X
    for (let r = 4; r < dataBOM.length; r++) {
      const row = dataBOM[r] || [];
      const ma5 = `${row[2] || ""}`.trim();
      let version = `${row[3] || ""}`.trim();
      if (!ma5 || !version) continue;
      if (!listYC.has(ma5)) continue;

      for (const proc of processes) {
        const cellVal = row[proc.colIndex];
        const hasX = cellVal && `${cellVal}`.trim().toUpperCase() === "X";
        if (!hasX) continue;

        let finalCode;
        if (proc.suffix) {
          finalCode = "4" + ma5.slice(1) + proc.suffix;
        } else {
          finalCode = ma5;
        }

  
        // xử lý version: nếu là 3 → 4, nếu là 4 → 5
        let versionNum = parseInt(version.replace(/\D/g, "")) || 0;
        if (versionNum === 3) versionNum = 4;
        else if (versionNum === 4) versionNum = 5;

        // ánh xạ routing name sang routing no
        const routingMap = {
          "Extrusion": 1,
          "UV": 2,
          "UV+Bigsheet": 11,
          "Profiling": 4,
          "Packaging": 6,
          "Profiling+Bevel": 12,
          "Padding+Packaging": 8
        };

        const routingNo = routingMap[proc.name] || "";

        results.push({
          "mã đầu 5": ma5,
          inventoryid: finalCode,
          inventoryname: "",
          routingname: proc.name,
          version: versionNum,
          description: "",
          mftimes: 100,
          no: "",
          routingno: routingNo,
        });

        // ✅ Clone PACKAGING thêm dòng version 99 nếu là version nhỏ nhất
        if (proc.name === "Packaging") {
          const verNum = parseInt(version.replace(/\D/g, "")) || 0;
          if (verNum === minVersionMap[ma5]) {
            results.push({
              "mã đầu 5": ma5,
              inventoryid: ma5,
              inventoryname: "",
              routingname: "Packaging",
              version: 99,
              description: "",
              mftimes: 100,
              no: "",
              routingno: routingNo,
            });
          }
        }
      }
    }

    // ==== Xuất file Excel ====
    const outWb = new ExcelJS.Workbook();
    const outWs = outWb.addWorksheet("result");

    outWs.addRow([
      "mã đầu 5",
      "InventoryID",
      "Inventory Name",
      "Version",
      "Description",
      "No",
      "Routing No",
      "Routing Name",
      "MFTimes",
    ]);

    results.forEach((r) =>
      outWs.addRow([
        r["mã đầu 5"],
        r.inventoryid,
        r.inventoryname,
        r.version,
        r.description,
        r.no,
        r.routingno,
        r.routingname,
        r.mftimes,
      ])
    );

    const outputPath = path.join(__dirname, `result_${Date.now()}.xlsx`);
    await outWb.xlsx.writeFile(outputPath);
    fs.unlinkSync(req.file.path);

    res.download(outputPath, "Routing_result.xlsx", (err) => {
      if (err) console.error(err);
      fs.unlinkSync(outputPath);
    });
  } catch (err) {
    console.error("❌ Lỗi xử lý:", err);
    if (req?.file?.path) fs.unlinkSync(req.file.path);
    res.status(500).send("Lỗi xử lý file Excel.");
  }
});

const PORT = 3000;
app.listen(PORT, () =>
  console.log(`✅ Server chạy tại http://localhost:${PORT}`)
);
