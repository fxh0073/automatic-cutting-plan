import ExcelJS from "exceljs";
import fs from "node:fs/promises";
import path from "node:path";

const outputFile = path.resolve(process.argv[2] ?? "examples/下料工艺卡_示例.xlsx");
await fs.mkdir(path.dirname(outputFile), { recursive: true });

const workbook = new ExcelJS.Workbook();
const sheet = workbook.addWorksheet("锯床");
sheet.showGridLines = false;
sheet.mergeCells("A1:G1");
sheet.getCell("A1").value = "型材下料工艺卡（脱敏示例）";
sheet.getCell("A1").font = { name: "Microsoft YaHei", size: 16, bold: true, color: { argb: "FFFFFFFF" } };
sheet.getCell("A1").alignment = { horizontal: "center", vertical: "middle" };
sheet.getCell("A1").fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FF1F4E78" } };
sheet.getRow(1).height = 28;

sheet.mergeCells("A2:G2");
sheet.getCell("A2").value = "工艺要求：长度尺寸公差：（-1,-3）；对角线允差：2；工件堆放整齐。";
sheet.getCell("A2").font = { name: "Microsoft YaHei", italic: true, color: { argb: "FF666666" } };
sheet.getCell("A2").alignment = { wrapText: true, vertical: "middle" };
sheet.getRow(2).height = 24;

sheet.addRow([]);
sheet.addRow(["序号", "部件名称", "材质", "规格尺寸", "数量", "图号", "备注"]);
const header = sheet.getRow(4);
header.height = 24;
header.eachCell((cell) => {
  cell.font = { name: "Microsoft YaHei", bold: true };
  cell.alignment = { horizontal: "center", vertical: "middle", wrapText: true };
  cell.fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FFD9EAF7" } };
  cell.border = { top: { style: "thin" }, bottom: { style: "thin" }, left: { style: "thin" }, right: { style: "thin" } };
});

const rows = [
  [1, "示例框架", "6063-T5", "L=2700", 4, "EX-001", "示例数据"],
  [2, "示例框架", "6063-T5", "L=1950", 6, "EX-001", "示例数据"],
  [3, "示例框架", "6063-T5", "L=1250", 8, "EX-001", "示例数据"],
  [4, "示例框架", "6063-T5", "L=735", 10, "EX-001", "示例数据"],
  [5, "示例框架", "6063-T5", "L=480", 6, "EX-001", "示例数据"],
];
for (const values of rows) sheet.addRow(values);
for (let row = 5; row <= 9; row += 1) {
  sheet.getRow(row).eachCell((cell) => {
    cell.font = { name: "Microsoft YaHei" };
    cell.alignment = { vertical: "middle" };
    cell.border = { top: { style: "thin", color: { argb: "FFB7C9D6" } }, bottom: { style: "thin", color: { argb: "FFB7C9D6" } }, left: { style: "thin", color: { argb: "FFB7C9D6" } }, right: { style: "thin", color: { argb: "FFB7C9D6" } } };
  });
  sheet.getCell(`A${row}`).alignment = { horizontal: "center", vertical: "middle" };
  sheet.getCell(`C${row}`).alignment = { horizontal: "center", vertical: "middle" };
  sheet.getCell(`D${row}`).alignment = { horizontal: "center", vertical: "middle" };
  sheet.getCell(`E${row}`).alignment = { horizontal: "center", vertical: "middle" };
}
sheet.columns = [
  { key: "no", width: 8 },
  { key: "part", width: 18 },
  { key: "material", width: 14 },
  { key: "spec", width: 16 },
  { key: "qty", width: 10 },
  { key: "drawing", width: 14 },
  { key: "remark", width: 18 },
];
sheet.views = [{ state: "frozen", ySplit: 4, topLeftCell: "A5", activeCell: "A5" }];
workbook.calcProperties = { fullCalcOnLoad: true, forceFullCalc: true, calcMode: "auto" };
await workbook.xlsx.writeFile(outputFile);
console.log(JSON.stringify({ outputFile, rows: rows.length }, null, 2));
