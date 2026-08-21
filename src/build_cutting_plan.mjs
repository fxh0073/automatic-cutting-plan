import fs from "node:fs/promises";
import path from "node:path";
import { fileURLToPath } from "node:url";
import ExcelJS from "exceljs";

function columnName(index) {
  let n = index;
  let out = "";
  while (n > 0) {
    const remainder = (n - 1) % 26;
    out = String.fromCharCode(65 + remainder) + out;
    n = Math.floor((n - 1) / 26);
  }
  return out;
}

function parseCellAddress(address) {
  const match = String(address).replace(/\$/g, "").toUpperCase().match(/^([A-Z]+)(\d+)$/);
  if (!match) throw new Error(`无效单元格地址：${address}`);
  let column = 0;
  for (const letter of match[1]) column = column * 26 + letter.charCodeAt(0) - 64;
  return { row: Number(match[2]), column };
}

function parseRangeAddress(address) {
  const parts = String(address).split(":");
  const start = parseCellAddress(parts[0]);
  const end = parseCellAddress(parts[1] ?? parts[0]);
  return {
    startRow: Math.min(start.row, end.row),
    startColumn: Math.min(start.column, end.column),
    endRow: Math.max(start.row, end.row),
    endColumn: Math.max(start.column, end.column),
  };
}

function argb(color) {
  if (typeof color !== "string") return "FF000000";
  const value = color.replace(/^#/, "").trim();
  if (/^[0-9A-Fa-f]{8}$/.test(value)) return value.toUpperCase();
  if (/^[0-9A-Fa-f]{6}$/.test(value)) return `FF${value.toUpperCase()}`;
  return "FF000000";
}

function readableCellValue(value) {
  if (value == null) return null;
  if (typeof value === "object") {
    if (Object.prototype.hasOwnProperty.call(value, "result")) return value.result;
    if (Array.isArray(value.richText)) return value.richText.map((part) => part.text ?? "").join("");
    if (Object.prototype.hasOwnProperty.call(value, "text")) return value.text;
  }
  return value;
}

class ExcelRangeFormat {
  constructor(range) {
    this.range = range;
  }

  eachCell(callback) {
    for (let row = this.range.startRow; row <= this.range.endRow; row += 1) {
      for (let column = this.range.startColumn; column <= this.range.endColumn; column += 1) {
        callback(this.range.sheet.raw.getCell(row, column), row, column);
      }
    }
  }

  set fill(color) {
    const fill = { type: "pattern", pattern: "solid", fgColor: { argb: argb(color) } };
    this.eachCell((cell) => { cell.fill = fill; });
  }

  set font(options = {}) {
    this.eachCell((cell) => {
      const previous = cell.font ?? {};
      const next = {
        ...previous,
        name: options.typeface ?? options.name ?? previous.name ?? "宋体",
        size: options.fontSize ?? options.size ?? previous.size ?? 10,
      };
      if (options.bold !== undefined) next.bold = options.bold;
      if (options.italic !== undefined) next.italic = options.italic;
      if (options.color) next.color = { argb: argb(options.color) };
      cell.font = next;
    });
  }

  set horizontalAlignment(value) {
    this.eachCell((cell) => { cell.alignment = { ...cell.alignment, horizontal: value }; });
  }

  set verticalAlignment(value) {
    this.eachCell((cell) => { cell.alignment = { ...cell.alignment, vertical: value }; });
  }

  set wrapText(value) {
    this.eachCell((cell) => { cell.alignment = { ...cell.alignment, wrapText: value }; });
  }

  set rowHeight(value) {
    for (let row = this.range.startRow; row <= this.range.endRow; row += 1) this.range.sheet.raw.getRow(row).height = value;
  }

  set columnWidth(value) {
    for (let column = this.range.startColumn; column <= this.range.endColumn; column += 1) this.range.sheet.raw.getColumn(column).width = value;
  }

  set numberFormat(value) {
    this.eachCell((cell) => { cell.numFmt = value; });
  }

  set borders(spec = {}) {
    const style = spec.style ?? "thin";
    const color = { argb: argb(spec.color) };
    const edge = { style, color };
    this.eachCell((cell, row, column) => {
      const outsideOnly = spec.preset === "outside";
      cell.border = {
        top: !outsideOnly || row === this.range.startRow ? edge : undefined,
        bottom: !outsideOnly || row === this.range.endRow ? edge : undefined,
        left: !outsideOnly || column === this.range.startColumn ? edge : undefined,
        right: !outsideOnly || column === this.range.endColumn ? edge : undefined,
      };
    });
  }
}

class ExcelRangeAdapter {
  constructor(sheet, address) {
    this.sheet = sheet;
    this.address = address;
    Object.assign(this, parseRangeAddress(address));
    this.conditionalFormats = {
      add: (type, config = {}) => {
        if (type !== "containsText" || typeof this.sheet.raw.addConditionalFormatting !== "function") return;
        try {
          this.sheet.raw.addConditionalFormatting({
            ref: this.address,
            rules: [{
              type: "containsText",
              operator: "containsText",
              text: config.text ?? "",
              style: {
                fill: { type: "pattern", pattern: "solid", bgColor: { argb: argb(config.format?.fill) }, fgColor: { argb: argb(config.format?.fill) } },
                font: { bold: Boolean(config.format?.font?.bold), color: { argb: argb(config.format?.font?.color) } },
              },
            }],
          });
        } catch {
          // Conditional formatting is presentation-only; do not block export.
        }
      },
    };
  }

  get values() {
    const rows = [];
    for (let row = this.startRow; row <= this.endRow; row += 1) {
      const values = [];
      for (let column = this.startColumn; column <= this.endColumn; column += 1) values.push(readableCellValue(this.sheet.raw.getCell(row, column).value));
      rows.push(values);
    }
    return rows;
  }

  set values(matrix) {
    for (let row = this.startRow; row <= this.endRow; row += 1) {
      for (let column = this.startColumn; column <= this.endColumn; column += 1) {
        const value = matrix?.[row - this.startRow]?.[column - this.startColumn] ?? null;
        this.sheet.raw.getCell(row, column).value = value;
      }
    }
  }

  set formulas(matrix) {
    for (let row = this.startRow; row <= this.endRow; row += 1) {
      for (let column = this.startColumn; column <= this.endColumn; column += 1) {
        const formula = matrix?.[row - this.startRow]?.[column - this.startColumn];
        this.sheet.raw.getCell(row, column).value = formula ? { formula: String(formula).replace(/^=/, "") } : null;
      }
    }
  }

  get format() {
    return new ExcelRangeFormat(this);
  }

  merge() {
    this.sheet.raw.mergeCells(this.address);
  }
}

class ExcelWorksheetAdapter {
  constructor(raw) {
    this.raw = raw;
    this.name = raw.name;
    this.freezePanes = {
      freezeRows: (count) => {
        this.raw.views = [{ state: "frozen", ySplit: count, topLeftCell: `A${count + 1}`, activeCell: "A1" }];
      },
    };
  }

  get showGridLines() {
    return this.raw.showGridLines;
  }

  set showGridLines(value) {
    this.raw.showGridLines = value;
  }

  getRange(address) {
    return new ExcelRangeAdapter(this, address);
  }

  getUsedRange() {
    const endRow = Math.max(this.raw.actualRowCount || 1, 1);
    const endColumn = Math.max(this.raw.actualColumnCount || 1, 1);
    return this.getRange(`A1:${columnName(endColumn)}${endRow}`);
  }
}

class ExcelWorksheetCollectionAdapter {
  constructor(rawWorkbook) {
    this.rawWorkbook = rawWorkbook;
  }

  add(name) {
    return new ExcelWorksheetAdapter(this.rawWorkbook.addWorksheet(name));
  }

  getItem(name) {
    const sheet = this.rawWorkbook.getWorksheet(name);
    if (!sheet) throw new Error(`工作表不存在：${name}`);
    return new ExcelWorksheetAdapter(sheet);
  }

  getItemAt(index) {
    const sheet = this.rawWorkbook.worksheets[index];
    if (!sheet) throw new Error(`工作表索引不存在：${index}`);
    return new ExcelWorksheetAdapter(sheet);
  }
}

class ExcelWorkbookAdapter {
  constructor() {
    this.raw = new ExcelJS.Workbook();
    this.worksheets = new ExcelWorksheetCollectionAdapter(this.raw);
  }

  async load(filePath) {
    await this.raw.xlsx.readFile(filePath);
    this.raw.calcProperties = { fullCalcOnLoad: true, forceFullCalc: true, calcMode: "auto" };
  }

  async save(filePath) {
    await this.raw.xlsx.writeFile(filePath);
  }
}

function readCliArgs(argv) {
  const args = {};
  for (let i = 0; i < argv.length; i += 1) {
    const token = argv[i];
    if (!token.startsWith("--")) continue;
    const key = token.slice(2).replace(/-([a-z])/g, (_, letter) => letter.toUpperCase());
    const next = argv[i + 1];
    args[key] = next && !next.startsWith("--") ? (++i, next) : true;
  }
  return args;
}

const cli = readCliArgs(process.argv.slice(2));
const scriptDir = path.dirname(fileURLToPath(import.meta.url));
const root = process.env.CUT_ROOT ?? path.resolve(scriptDir, "..");
const workRoot = process.env.CUT_WORK_DIR ?? path.join(process.env.TEMP ?? process.env.TMP ?? root, "自动下料程序");
const workDir = path.join(path.resolve(workRoot), `work-${process.pid}`);
const sourceFile = path.resolve(String(cli.input ?? process.env.CUT_SOURCE_FILE ?? path.join(root, "下料工艺卡.xlsx")));
const defaultOutput = path.join(path.dirname(sourceFile), `自动下料方案_${new Date().toISOString().replace(/[-:TZ.]/g, "").slice(0, 14)}.xlsx`);
const outputFile = path.resolve(String(cli.output ?? process.env.CUT_OUTPUT_FILE ?? path.join(root, "outputs", "019fefcd-0539-7033-a009-797d7b331f5f", process.env.CUT_OUTPUT_FILE_NAME ?? path.basename(defaultOutput))));
const outputDir = path.dirname(outputFile);
const inspectOnly = cli.inspectOnly === true || process.env.CUT_INSPECT_ONLY === "1";
const stockLength = Number(cli.stockLength ?? process.env.CUT_STOCK_LENGTH ?? 6000);
const kerf = Number(cli.kerf ?? process.env.CUT_KERF ?? 3);
const candidateCount = Number(cli.candidates ?? process.env.CUT_CANDIDATES ?? 2000);
const lengthStep = Number(cli.lengthStep ?? process.env.CUT_LENGTH_STEP ?? 5);

const colors = {
  border: "#808080",
  title: "#1F4E78",
  section: "#D9EAF7",
  label: "#F2F2F2",
  input: "#FFF2CC",
  formula: "#E2F0D9",
  warning: "#FCE4D6",
  text: "#000000",
};

function cellText(value) {
  return value == null ? "" : String(value).trim();
}

function toNumber(value) {
  if (typeof value === "number") return value;
  const match = cellText(value).replace(/,/g, "").match(/-?\d+(?:\.\d+)?/);
  return match ? Number(match[0]) : NaN;
}

function parseLength(value) {
  if (typeof value === "number" && Number.isFinite(value)) return value;
  const text = cellText(value).replace(/＝/g, "=");
  const match = text.match(/(?:L|长度|长)\s*[:=]?\s*(\d+(?:\.\d+)?)/i) || text.match(/(\d+(?:\.\d+)?)\s*mm/i);
  return match ? Number(match[1]) : NaN;
}

function normalizedHeader(value) {
  return cellText(value).replace(/[\s()（）\[\]【】]/g, "").toLowerCase();
}

function findColumnMap(rows) {
  const defaults = { no: 0, part: 1, material: 2, spec: 3, qty: 4, drawing: 5, remark: 6 };
  for (let rowIndex = 0; rowIndex < Math.min(rows.length, 30); rowIndex += 1) {
    const row = rows[rowIndex] ?? [];
    const map = {};
    for (let col = 0; col < row.length; col += 1) {
      const header = normalizedHeader(row[col]);
      if (!header) continue;
      if (header.includes("序号")) map.no = col;
      else if (header.includes("部件名称") || header === "部件") map.part = col;
      else if (header.includes("材质") || header.includes("材料")) map.material = col;
      else if (header.includes("规格") || header.includes("尺寸")) map.spec = col;
      else if (header.includes("数量") || header.includes("件数")) map.qty = col;
      else if (header.includes("图号") || header.includes("图纸")) map.drawing = col;
      else if (header.includes("备注") || header.includes("说明")) map.remark = col;
    }
    if (["no", "part", "material", "spec", "qty"].every((key) => Number.isInteger(map[key]))) {
      return { ...defaults, ...map, headerRow: rowIndex };
    }
  }
  return { ...defaults, headerRow: -1 };
}

function parseCard(rows) {
  const columns = findColumnMap(rows);
  const records = [];
  const instructions = [];
  for (let rowIndex = Math.max(0, columns.headerRow + 1); rowIndex < rows.length; rowIndex += 1) {
    const row = rows[rowIndex] ?? [];
    const rowText = row.map(cellText).filter(Boolean).join(" ");
    if (/工艺要求|长度尺寸公差|对角线允差|工件堆放/i.test(rowText)) {
      instructions.push(rowText);
      continue;
    }
    const no = toNumber(row[columns.no]);
    const part = cellText(row[columns.part]);
    const material = cellText(row[columns.material]);
    const length = parseLength(row[columns.spec]);
    const qty = toNumber(row[columns.qty]);
    if (!Number.isFinite(no) || !part || !material || !Number.isFinite(length) || length <= 0 || !Number.isInteger(qty) || qty <= 0) continue;
    records.push({
      no,
      part,
      material,
      length,
      qty,
      drawing: cellText(row[columns.drawing]),
      remark: cellText(row[columns.remark]),
    });
  }
  return {
    records,
    instruction: instructions[0] || "长度尺寸公差：（-1,-3）；对角线允差：2；工件堆放整齐。",
    headerRow: columns.headerRow,
    columns,
  };
}

function seeded(seed) {
  let state = seed >>> 0;
  return () => {
    state = (1664525 * state + 1013904223) >>> 0;
    return state / 4294967296;
  };
}

function makeItems(records) {
  const items = [];
  for (const record of records) {
    for (let i = 1; i <= record.qty; i += 1) {
      items.push({ ...record, instance: i, cost: record.length + kerf });
    }
  }
  return items;
}

function bestFit(order, capacity) {
  const bins = [];
  for (const item of order) {
    let bestIndex = -1;
    let bestRemaining = Infinity;
    for (let i = 0; i < bins.length; i += 1) {
      const remaining = bins[i].remaining - item.cost;
      if (remaining >= 0 && remaining < bestRemaining) {
        bestIndex = i;
        bestRemaining = remaining;
      }
    }
    if (bestIndex < 0) {
      bins.push({ remaining: capacity - item.cost, items: [item] });
    } else {
      bins[bestIndex].remaining = bestRemaining;
      bins[bestIndex].items.push(item);
    }
  }
  return bins;
}

function firstFit(order, capacity) {
  const bins = [];
  for (const item of order) {
    let placed = false;
    for (const bin of bins) {
      if (bin.remaining >= item.cost) {
        bin.remaining -= item.cost;
        bin.items.push(item);
        placed = true;
        break;
      }
    }
    if (!placed) bins.push({ remaining: capacity - item.cost, items: [item] });
  }
  return bins;
}

function score(bins) {
  const waste = bins.reduce((sum, b) => sum + b.remaining, 0);
  const largestWaste = Math.max(...bins.map((b) => b.remaining));
  const patternCount = new Set(bins.map((b) => {
    const counts = new Map();
    for (const item of b.items) counts.set(item.length, (counts.get(item.length) ?? 0) + 1);
    return [...counts.entries()].sort((a, b) => b[0] - a[0]).map(([length, qty]) => `${length}×${qty}`).join(",");
  })).size;
  return [bins.length, waste, patternCount, largestWaste];
}

function better(a, b) {
  if (!b) return true;
  const sa = score(a);
  const sb = score(b);
  for (let i = 0; i < sa.length; i += 1) {
    if (sa[i] !== sb[i]) return sa[i] < sb[i];
  }
  return false;
}

function summarizeBins(bins) {
  return bins.map((bin, index) => {
    const counts = new Map();
    for (const item of bin.items) counts.set(item.length, (counts.get(item.length) ?? 0) + 1);
    const cuts = [...counts.entries()].sort((a, b) => b[0] - a[0]);
    const finishedLength = bin.items.reduce((sum, item) => sum + item.length, 0);
    const cutCount = bin.items.length;
    return {
      id: index + 1,
      stockLength,
      cuts,
      finishedLength,
      cutCount,
      kerfLength: cutCount * kerf,
      remaining: stockLength - finishedLength - cutCount * kerf,
      utilization: (finishedLength + cutCount * kerf) / stockLength,
    };
  });
}

function aggregatePatterns(bins) {
  const map = new Map();
  for (const bin of bins) {
    const key = bin.cuts.map(([length, qty]) => `${length}×${qty}`).join(",");
    const current = map.get(key);
    if (current) {
      current.barQty += 1;
    } else {
      map.set(key, { cuts: bin.cuts, barQty: 1 });
    }
  }
  const patterns = [...map.values()];
  patterns.sort((a, b) => {
    const aLen = a.cuts.reduce((s, [length, qty]) => s + length * qty, 0);
    const bLen = b.cuts.reduce((s, [length, qty]) => s + length * qty, 0);
    return bLen - aLen || a.cuts[0][0] - b.cuts[0][0];
  });
  return patterns.map((pattern, index) => {
    const finishedLength = pattern.cuts.reduce((sum, [length, qty]) => sum + length * qty, 0);
    const cutCount = pattern.cuts.reduce((sum, [, qty]) => sum + qty, 0);
    const consumedLength = finishedLength + cutCount * kerf;
    const recommendedLength = Math.ceil(consumedLength / lengthStep) * lengthStep;
    if (recommendedLength > stockLength) throw new Error(`方案 ${index + 1} 超出最长原料定尺`);
    return {
      scheme: index + 1,
      stockLength: recommendedLength,
      maxStockLength: stockLength,
      ...pattern,
      finishedLength,
      cutCount,
      consumedLength,
      tailLength: recommendedLength - consumedLength,
      patternText: pattern.cuts.map(([length, qty]) => `${length}×${qty}`).join(", "),
    };
  });
}

function excelCol(index) {
  let n = index + 1;
  let out = "";
  while (n > 0) {
    const r = (n - 1) % 26;
    out = String.fromCharCode(65 + r) + out;
    n = Math.floor((n - 1) / 26);
  }
  return out;
}

function setFont(range, options = {}) {
  range.format.font = { typeface: "宋体", fontSize: 10, color: colors.text, ...options };
}

function setBorders(range, color = colors.border) {
  range.format.borders = { preset: "all", style: "thin", color };
}

function styleTableHeader(range) {
  range.format.fill = colors.section;
  setFont(range, { bold: true });
  range.format.horizontalAlignment = "center";
  range.format.wrapText = true;
  setBorders(range);
}

function styleLabel(range) {
  range.format.fill = colors.label;
  setFont(range, { bold: true });
  range.format.horizontalAlignment = "center";
  setBorders(range);
}

function styleFormula(range) {
  range.format.fill = colors.formula;
  setFont(range);
  setBorders(range);
}

function applyBodyStyle(sheet, rangeAddress) {
  const range = sheet.getRange(rangeAddress);
  setFont(range);
  range.format.horizontalAlignment = "center";
  range.format.wrapText = false;
  setBorders(range);
}

function makeCutDetailRows(patterns) {
  const rows = [];
  for (const pattern of patterns) {
    for (const [length, qty] of pattern.cuts) {
      rows.push([pattern.scheme, pattern.stockLength, pattern.barQty, length, qty, null, null, null, null, null, "按方案号组合切割"]);
    }
  }
  return rows;
}

function getCuttingSheet(workbook) {
  try {
    return workbook.worksheets.getItem("锯床");
  } catch {
    for (let i = 0; i < 30; i += 1) {
      let sheet;
      try { sheet = workbook.worksheets.getItemAt(i); } catch { break; }
      if (sheet && /锯床|下料/.test(sheet.name)) return sheet;
    }
  }
  throw new Error("未找到‘锯床’工作表，请确认输入文件为下料工艺卡。");
}

function hasWorksheet(workbook, name) {
  try {
    workbook.worksheets.getItem(name);
    return true;
  } catch {
    return false;
  }
}

async function main() {
  if (!Number.isFinite(stockLength) || stockLength <= 0 || !Number.isFinite(kerf) || kerf < 0) throw new Error("原料定尺或锯缝参数无效");
  if (!Number.isFinite(lengthStep) || lengthStep <= 0) throw new Error("原料长度取整步长无效");
  if (path.resolve(sourceFile) === path.resolve(outputFile)) throw new Error("输出文件不能覆盖输入工艺卡，请选择新的输出文件名。");
  await fs.mkdir(workDir, { recursive: true });
  await fs.mkdir(outputDir, { recursive: true });

  const workbook = new ExcelWorkbookAdapter();
  await workbook.load(sourceFile);
  if (inspectOnly) {
    const sheets = [];
    for (let i = 0; i < 20; i += 1) {
      let sheet;
      try { sheet = workbook.worksheets.getItemAt(i); } catch { break; }
      if (!sheet) break;
      const name = sheet.name;
      const used = sheet.getUsedRange();
      sheets.push({ index: i, name, usedRange: used.address });
    }
    console.log(JSON.stringify({ sourceFile, sheets }, null, 2));
    return;
  }
  const sourceSheet = getCuttingSheet(workbook);
  const sourceRange = sourceSheet.getUsedRange(true);
  const sourceRows = sourceRange?.values ?? [];
  const card = parseCard(sourceRows);
  const records = card.records;
  if (!records.length) throw new Error("未从锯床工作表解析到有效下料规格");
  const materialKeys = new Set(records.map((r) => `${r.part}||${r.material}`));
  if (materialKeys.size > 1) throw new Error("当前版本要求一次处理同一部件、同一材质的型材；请将不同部件或材质拆分后分别运行。");
  for (const name of ["下料方案", "切割明细", "校核"]) {
    if (hasWorksheet(workbook, name)) throw new Error(`输入文件已包含‘${name}’工作表，请选择原始下料工艺卡，不要使用上一次输出文件作为输入。`);
  }

  const items = makeItems(records);
  const baseOrder = [...items].sort((a, b) => b.cost - a.cost || b.length - a.length || a.no - b.no);
  let best = null;
  const candidates = [bestFit(baseOrder, stockLength), firstFit(baseOrder, stockLength)];
  for (let seed = 1; seed <= candidateCount; seed += 1) {
    const rand = seeded(seed);
    const order = [...items].sort((a, b) => b.cost - a.cost || rand() - 0.5);
    candidates.push(bestFit(order, stockLength));
    if (seed <= Math.ceil(candidateCount / 4)) {
      const shuffled = [...items].sort(() => rand() - 0.5);
      candidates.push(bestFit(shuffled, stockLength));
      candidates.push(firstFit(shuffled, stockLength));
    }
  }
  for (const candidate of candidates) if (better(candidate, best)) best = candidate;

  const bins = summarizeBins(best);
  const patterns = aggregatePatterns(bins);
  const detailRows = makeCutDetailRows(patterns);
  const part = records[0].part;
  const material = records[0].material;
  const sourceBaseName = path.basename(sourceFile);
  const drawingList = [...new Set(records.map((r) => r.drawing).filter(Boolean))];
  const drawing = drawingList.length === 1 ? drawingList[0] : "多图号";
  const drawingNote = drawingList.length ? `图号：${drawingList.join("、")}；` : "";
  const instruction = card.instruction;

  const planSheet = workbook.worksheets.add("下料方案");
  const detailSheet = workbook.worksheets.add("切割明细");
  const checkSheet = workbook.worksheets.add("校核");
  for (const sheet of [planSheet, detailSheet, checkSheet]) sheet.showGridLines = false;

  const checkStart = 5;
  const checkEnd = checkStart + records.length - 1;
  const checkTotal = checkEnd + 1;
  const detailStart = 5;
  const detailEnd = detailStart + detailRows.length - 1;
  const planStart = 9;
  const planEnd = planStart + patterns.length - 1;
  const planTotal = planEnd + 1;

  // 下料方案 sheet: template-like summary and aggregated cutting patterns.
  planSheet.getRange("A1:L1").merge();
  planSheet.getRange("A1").values = [[`${part} 型材优化切割下料方案`]];
  setFont(planSheet.getRange("A1"), { bold: true, fontSize: 18, color: "#FFFFFF" });
  planSheet.getRange("A1:L1").format.fill = colors.title;
  planSheet.getRange("A1:L1").format.horizontalAlignment = "center";
  planSheet.getRange("A1:L1").format.rowHeight = 30;

  planSheet.getRange("A2:L4").format.rowHeight = 22;
  planSheet.getRange("A2:L4").format.wrapText = true;
  planSheet.getRange("A2:L4").format.horizontalAlignment = "center";
  planSheet.getRange("A2:L4").format.verticalAlignment = "center";
  planSheet.getRange("A2:L4").format.borders = { preset: "all", style: "thin", color: colors.border };
  planSheet.getRange("A2:A4").values = [["部件名称"], ["最长原料定尺(mm)"], ["工艺要求"]];
  planSheet.getRange("C2:C3").values = [["材质"], ["锯缝(mm)"]];
  planSheet.getRange("E2:E3").values = [["图号"], ["优化方式"]];
  planSheet.getRange("G2:G3").values = [["套数"], ["编制依据"]];
  planSheet.getRange("I2:I3").values = [["方案根数"], ["算法说明"]];
  styleLabel(planSheet.getRange("A2:A4"));
  styleLabel(planSheet.getRange("C2:C3"));
  styleLabel(planSheet.getRange("E2:E3"));
  styleLabel(planSheet.getRange("G2:G3"));
  styleLabel(planSheet.getRange("I2:I3"));
  planSheet.getRange("B2").values = [[part]];
  planSheet.getRange("B3").values = [[stockLength]];
  planSheet.getRange("B4:L4").merge();
  planSheet.getRange("B4").values = [[instruction]];
  planSheet.getRange("D2").values = [[material]];
  planSheet.getRange("D3").values = [[kerf]];
  planSheet.getRange("F2").values = [[drawing]];
  planSheet.getRange("F3:I3").merge();
  planSheet.getRange("F3").values = [[`${drawingNote}按${stockLength.toLocaleString("en-US")} mm装载约束组合；推荐原料长度按消耗量向上取${lengthStep} mm`]];
  planSheet.getRange("H2").values = [["1套"]];
  planSheet.getRange("H3").values = [[sourceBaseName]];
  planSheet.getRange("J2").formulas = [[`=SUM(C${planStart}:C${planEnd})`]];
  planSheet.getRange("J3:L3").merge();
  planSheet.getRange("J3").values = [["切割组合为建议方案；锯缝参数变更后需重新优化"]];
  for (const range of ["B2:B3", "D2:D3", "F2", "H2:H3", "J2", "B4", "F3", "J3"]) {
    styleFormula(planSheet.getRange(range));
  }
  planSheet.getRange("B3:D3").format.fill = colors.input;
  planSheet.getRange("D3").format.fill = colors.input;
  planSheet.getRange("A2:L4").format.borders = { preset: "all", style: "thin", color: colors.border };

  planSheet.getRange("A5:L6").format.rowHeight = 24;
  planSheet.getRange("A5:L6").format.wrapText = true;
  planSheet.getRange("A5:L6").format.borders = { preset: "all", style: "thin", color: colors.border };
  planSheet.getRange("A5:L6").values = [
    ["需求规格数", null, "成品总件数", null, "原料根数", null, "原料总长度(mm)", null, "成品利用率", null, "综合占用率", null],
    ["成品总长度(mm)", null, "锯缝总占用(mm)", null, "尾料总长度(mm)", null, "余料率", null, "校核差异(件)", null, "校核状态", null],
  ];
  for (const range of ["A5:A6", "C5:C6", "E5:E6", "G5:G6", "I5:I6", "K5:K6"]) styleLabel(planSheet.getRange(range));
  planSheet.getRange("B5").formulas = [[`=COUNTA('校核'!$A$${checkStart}:$A$${checkEnd})`]];
  planSheet.getRange("D5").formulas = [[`=SUM('校核'!$E$${checkStart}:$E$${checkEnd})`]];
  planSheet.getRange("F5").formulas = [[`=SUM(C${planStart}:C${planEnd})`]];
  planSheet.getRange("H5").formulas = [[`=SUM(J${planStart}:J${planEnd})`]];
  planSheet.getRange("J5").formulas = [["=B6/H5"]];
  planSheet.getRange("L5").formulas = [["=(B6+D6)/H5"]];
  planSheet.getRange("B6").formulas = [[`=SUM('校核'!$I$${checkStart}:$I$${checkEnd})`]];
  planSheet.getRange("D6").formulas = [[`=SUM('切割明细'!$J$${detailStart}:$J$${detailEnd})`]];
  planSheet.getRange("F6").formulas = [[`=SUM(K${planStart}:K${planEnd})`]];
  planSheet.getRange("H6").formulas = [["=F6/H5"]];
  planSheet.getRange("J6").formulas = [[`='校核'!$G$${checkTotal}`]];
  planSheet.getRange("L6").formulas = [["=IF(J6=0,\"OK\",\"CHECK\")"]];
  styleFormula(planSheet.getRange("B5:B6"));
  styleFormula(planSheet.getRange("D5:D6"));
  styleFormula(planSheet.getRange("F5:F6"));
  styleFormula(planSheet.getRange("H5:H6"));
  styleFormula(planSheet.getRange("J5:J6"));
  styleFormula(planSheet.getRange("L5:L6"));
  planSheet.getRange("J5:L5").format.numberFormat = "0.00%";
  planSheet.getRange("H6").format.numberFormat = "0.00%";

  const planHeaders = [["方案号", "推荐原料长度(mm)", "原料根数", "单根下料方案", "单根件数", "单根成品长度(mm)", "单根锯缝(mm)", "单根尾料(mm)", "成品利用率", "总原料长度(mm)", "总尾料(mm)", "执行备注"]];
  planSheet.getRange("A8:L8").values = planHeaders;
  styleTableHeader(planSheet.getRange("A8:L8"));
  planSheet.getRange(`A${planStart}:D${planEnd}`).values = patterns.map((p) => [p.scheme, p.stockLength, p.barQty, p.patternText]);
  for (let r = planStart; r <= planEnd; r += 1) {
    planSheet.getRange(`E${r}`).formulas = [[`=SUMIF('切割明细'!$A$${detailStart}:$A$${detailEnd},A${r},'切割明细'!$E$${detailStart}:$E$${detailEnd})`]];
    planSheet.getRange(`F${r}`).formulas = [[`=SUMIF('切割明细'!$A$${detailStart}:$A$${detailEnd},A${r},'切割明细'!$G$${detailStart}:$G$${detailEnd})`]];
    planSheet.getRange(`G${r}`).formulas = [[`=E${r}*$D$3`]];
    planSheet.getRange(`H${r}`).formulas = [[`=B${r}-F${r}-G${r}`]];
    planSheet.getRange(`I${r}`).formulas = [[`=F${r}/B${r}`]];
    planSheet.getRange(`J${r}`).formulas = [[`=B${r}*C${r}`]];
    planSheet.getRange(`K${r}`).formulas = [[`=C${r}*H${r}`]];
  }
  planSheet.getRange(`L${planStart}:L${planEnd}`).values = patterns.map(() => ["按此方案逐根执行"]);
  applyBodyStyle(planSheet, `A${planStart}:L${planEnd}`);
  planSheet.getRange(`D${planStart}:D${planEnd}`).format.wrapText = true;
  planSheet.getRange(`L${planStart}:L${planEnd}`).format.wrapText = true;
  planSheet.getRange(`A${planStart}:L${planEnd}`).format.rowHeight = 26;
  planSheet.getRange(`B${planStart}:H${planEnd}`).format.numberFormat = "#,##0";
  planSheet.getRange(`I${planStart}:I${planEnd}`).format.numberFormat = "0.0%";
  planSheet.getRange(`J${planStart}:K${planEnd}`).format.numberFormat = "#,##0";
  planSheet.getRange(`A${planTotal}:L${planTotal}`).values = [["合计", null, null, "—", null, null, null, "—", null, null, null, "—"]];
  planSheet.getRange(`C${planTotal}`).formulas = [[`=SUM(C${planStart}:C${planEnd})`]];
  planSheet.getRange(`E${planTotal}`).formulas = [[`=SUM('切割明细'!$F$${detailStart}:$F$${detailEnd})`]];
  planSheet.getRange(`F${planTotal}`).formulas = [["=B6"]];
  planSheet.getRange(`G${planTotal}`).formulas = [["=D6"]];
  planSheet.getRange(`I${planTotal}`).formulas = [["=J5"]];
  planSheet.getRange(`J${planTotal}`).formulas = [["=H5"]];
  planSheet.getRange(`K${planTotal}`).formulas = [["=F6"]];
  styleTableHeader(planSheet.getRange(`A${planTotal}:L${planTotal}`));
  planSheet.getRange(`A${planTotal}:L${planTotal}`).format.fill = colors.formula;
  planSheet.getRange(`B${planTotal}:H${planTotal}`).format.numberFormat = "#,##0";
  planSheet.getRange(`I${planTotal}`).format.numberFormat = "0.0%";
  planSheet.getRange(`J${planTotal}:K${planTotal}`).format.numberFormat = "#,##0";

  // 切割明细 sheet: normalized data used by the formula-driven checks.
  detailSheet.getRange("A1:K1").merge();
  detailSheet.getRange("A1").values = [[`${part} 切割明细（按方案号展开）`]];
  detailSheet.getRange("A1:K1").format.fill = colors.title;
  setFont(detailSheet.getRange("A1"), { bold: true, fontSize: 16, color: "#FFFFFF" });
  detailSheet.getRange("A1:K1").format.horizontalAlignment = "center";
  detailSheet.getRange("A1:K1").format.rowHeight = 28;
  detailSheet.getRange("A2:K2").merge();
  detailSheet.getRange("A2").values = [["本表为下料方案的可追溯展开表：合计数量、成品长度和锯缝占用均由公式计算。"]];
  setFont(detailSheet.getRange("A2"), { italic: true, color: "#666666" });
  detailSheet.getRange("A2:K2").format.wrapText = true;
  detailSheet.getRange("A2:K2").format.rowHeight = 24;
  const detailHeaders = [["方案号", "推荐原料长度(mm)", "原料根数", "成品长度(mm)", "单根数量", "合计数量", "单根成品长度(mm)", "合计成品长度(mm)", "单根锯缝(mm)", "合计锯缝(mm)", "执行备注"]];
  detailSheet.getRange("A4:K4").values = detailHeaders;
  styleTableHeader(detailSheet.getRange("A4:K4"));
  detailSheet.getRange(`A${detailStart}:E${detailEnd}`).values = detailRows.map((r) => r.slice(0, 5));
  detailSheet.getRange(`K${detailStart}:K${detailEnd}`).values = detailRows.map((r) => [r[10]]);
  for (let r = detailStart; r <= detailEnd; r += 1) {
    detailSheet.getRange(`F${r}`).formulas = [[`=C${r}*E${r}`]];
    detailSheet.getRange(`G${r}`).formulas = [[`=D${r}*E${r}`]];
    detailSheet.getRange(`H${r}`).formulas = [[`=G${r}*C${r}`]];
    detailSheet.getRange(`I${r}`).formulas = [[`=E${r}*'下料方案'!$D$3`]];
    detailSheet.getRange(`J${r}`).formulas = [[`=I${r}*C${r}`]];
  }
  applyBodyStyle(detailSheet, `A${detailStart}:K${detailEnd}`);
  detailSheet.getRange(`A${detailStart}:K${detailEnd}`).format.rowHeight = 20;
  detailSheet.getRange(`B${detailStart}:J${detailEnd}`).format.numberFormat = "#,##0";
  detailSheet.getRange(`K${detailStart}:K${detailEnd}`).format.wrapText = true;

  // 校核 sheet: line-by-line reconciliation to the original card.
  checkSheet.getRange("A1:I1").merge();
  checkSheet.getRange("A1").values = [[`${part} 下料数量校核`]];
  checkSheet.getRange("A1:I1").format.fill = colors.title;
  setFont(checkSheet.getRange("A1"), { bold: true, fontSize: 16, color: "#FFFFFF" });
  checkSheet.getRange("A1:I1").format.horizontalAlignment = "center";
  checkSheet.getRange("A1:I1").format.rowHeight = 28;
  checkSheet.getRange("A2:I2").merge();
  checkSheet.getRange("A2").values = [[`来源：${sourceBaseName} / ${sourceSheet.name}；计划数量按成品长度从切割明细中 SUMIF 汇总。差异为 0 表示数量完全匹配。`]];
  setFont(checkSheet.getRange("A2"), { italic: true, color: "#666666" });
  checkSheet.getRange("A2:I2").format.wrapText = true;
  checkSheet.getRange("A2:I2").format.rowHeight = 24;
  const checkHeaders = [["工艺卡序号", "成品长度(mm)", "部件名称", "材质", "工艺卡需求(件)", "计划数量(件)", "差异(件)", "校核状态", "需求总长度(mm)"]];
  checkSheet.getRange("A4:I4").values = checkHeaders;
  styleTableHeader(checkSheet.getRange("A4:I4"));
  checkSheet.getRange(`A${checkStart}:E${checkEnd}`).values = records.map((r) => [r.no, r.length, r.part, r.material, r.qty]);
  for (let r = checkStart; r <= checkEnd; r += 1) {
    checkSheet.getRange(`F${r}`).formulas = [[`=SUMIF('切割明细'!$D$${detailStart}:$D$${detailEnd},B${r},'切割明细'!$F$${detailStart}:$F$${detailEnd})`]];
    checkSheet.getRange(`G${r}`).formulas = [[`=F${r}-E${r}`]];
    checkSheet.getRange(`H${r}`).formulas = [[`=IF(G${r}=0,"OK","CHECK")`]];
    checkSheet.getRange(`I${r}`).formulas = [[`=B${r}*E${r}`]];
  }
  applyBodyStyle(checkSheet, `A${checkStart}:I${checkEnd}`);
  checkSheet.getRange(`A${checkStart}:I${checkEnd}`).format.rowHeight = 20;
  checkSheet.getRange(`B${checkStart}:G${checkEnd}`).format.numberFormat = "#,##0";
  checkSheet.getRange(`I${checkStart}:I${checkEnd}`).format.numberFormat = "#,##0";
  checkSheet.getRange(`A${checkTotal}:I${checkTotal}`).values = [["合计", null, "—", "—", null, null, null, null, null]];
  checkSheet.getRange(`E${checkTotal}`).formulas = [[`=SUM(E${checkStart}:E${checkEnd})`]];
  checkSheet.getRange(`F${checkTotal}`).formulas = [[`=SUM(F${checkStart}:F${checkEnd})`]];
  checkSheet.getRange(`G${checkTotal}`).formulas = [[`=SUM(G${checkStart}:G${checkEnd})`]];
  checkSheet.getRange(`H${checkTotal}`).formulas = [[`=IF(G${checkTotal}=0,"OK","CHECK")`]];
  checkSheet.getRange(`I${checkTotal}`).formulas = [[`=SUM(I${checkStart}:I${checkEnd})`]];
  styleTableHeader(checkSheet.getRange(`A${checkTotal}:I${checkTotal}`));
  checkSheet.getRange(`A${checkTotal}:I${checkTotal}`).format.fill = colors.formula;
  checkSheet.getRange(`E${checkTotal}:I${checkTotal}`).format.numberFormat = "#,##0";
  checkSheet.getRange(`H${checkStart}:H${checkTotal}`).conditionalFormats.add("containsText", { text: "CHECK", format: { fill: colors.warning, font: { bold: true, color: "#9C0006" } } });

  // Column sizing and freezing, kept compact and consistent with the source/template.
  const planWidths = { A: 8, B: 14, C: 10, D: 50, E: 10, F: 16, G: 13, H: 13, I: 12, J: 16, K: 13, L: 18 };
  for (const [col, width] of Object.entries(planWidths)) planSheet.getRange(`${col}1:${col}${planTotal}`).format.columnWidth = width;
  planSheet.freezePanes.freezeRows(8);
  const detailWidths = { A: 8, B: 14, C: 10, D: 14, E: 10, F: 10, G: 16, H: 16, I: 13, J: 13, K: 18 };
  for (const [col, width] of Object.entries(detailWidths)) detailSheet.getRange(`${col}1:${col}${detailEnd}`).format.columnWidth = width;
  detailSheet.freezePanes.freezeRows(4);
  const checkWidths = { A: 11, B: 14, C: 14, D: 12, E: 14, F: 14, G: 11, H: 12, I: 16 };
  for (const [col, width] of Object.entries(checkWidths)) checkSheet.getRange(`${col}1:${col}${checkTotal}`).format.columnWidth = width;
  checkSheet.freezePanes.freezeRows(4);

  const debug = {
    sourceFile,
    sourceSheet: sourceSheet.name,
    headerRow: card.headerRow + 1,
    stockLength,
    kerf,
    recordCount: records.length,
    totalCount: records.reduce((s, r) => s + r.qty, 0),
    totalLength: records.reduce((s, r) => s + r.length * r.qty, 0),
    lowerBound: Math.ceil((records.reduce((s, r) => s + r.length * r.qty, 0) + kerf * items.length) / stockLength),
    rawBars: bins.length,
    patternCount: patterns.length,
    totalWaste: patterns.reduce((s, p) => s + p.tailLength * p.barQty, 0),
    patterns,
  };
  await fs.writeFile(path.join(workDir, "plan-debug.json"), JSON.stringify(debug, null, 2), "utf8");

  workbook.raw.calcProperties = { fullCalcOnLoad: true, forceFullCalc: true, calcMode: "auto" };
  await workbook.save(outputFile);
  console.log(JSON.stringify({
    outputFile,
    recordCount: debug.recordCount,
    totalCount: debug.totalCount,
    totalLength: debug.totalLength,
    lowerBound: debug.lowerBound,
    rawBars: debug.rawBars,
    patternCount: debug.patternCount,
    totalWaste: debug.totalWaste,
    formulas: "Excel formulas will recalculate when opened in Excel or WPS",
  }, null, 2));
}

await main();
