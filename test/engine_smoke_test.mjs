import assert from "node:assert/strict";
import { execFile } from "node:child_process";
import fs from "node:fs/promises";
import os from "node:os";
import path from "node:path";
import { fileURLToPath } from "node:url";
import { promisify } from "node:util";
import ExcelJS from "exceljs";

const execFileAsync = promisify(execFile);
const repoRoot = path.resolve(path.dirname(fileURLToPath(import.meta.url)), "..");
const inputFile = path.join(repoRoot, "examples", "下料工艺卡_示例.xlsx");
const tempDir = await fs.mkdtemp(path.join(os.tmpdir(), "automatic-cutting-test-"));
const outputFile = path.join(tempDir, "示例下料方案.xlsx");

try {
  const { stdout } = await execFileAsync(process.execPath, [
    path.join(repoRoot, "src", "build_cutting_plan.mjs"),
    "--input", inputFile,
    "--output", outputFile,
    "--stock-length", "6000",
    "--kerf", "3",
    "--length-step", "5",
    "--candidates", "300",
  ], { cwd: repoRoot, maxBuffer: 1024 * 1024, encoding: "utf8" });

  const summary = JSON.parse(stdout);
  assert.equal(summary.recordCount, 4);
  assert.equal(summary.totalCount, 28);
  assert.equal(summary.rawBars, 7);
  assert.equal(summary.patternCount, 4);

  const workbook = new ExcelJS.Workbook();
  await workbook.xlsx.readFile(outputFile);
  assert.deepEqual(workbook.worksheets.map((sheet) => sheet.name), ["锯床", "下料方案", "切割明细", "校核"]);
  assert.equal(workbook.getWorksheet("校核").getCell("H9").value.formula, 'IF(G9=0,"OK","CHECK")');
  console.log("engine_smoke_test: PASS");
} finally {
  await fs.rm(tempDir, { recursive: true, force: true });
}
