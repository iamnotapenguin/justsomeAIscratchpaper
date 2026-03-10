const ROWS = 100;
const COLS = 26;
const STORAGE_KEY = "sheetforge-workbook-v5";
const DEFAULT_COL_WIDTH = 95;
const MAX_HISTORY = 120;

const FUNCTIONS_CATALOG = [
  { name: "SUM", syntax: "SUM(v1, v2, ... | A1:B9)", desc: "Sum of numeric args/ranges." },
  { name: "AVG", syntax: "AVG(...) / AVERAGE(...) ", desc: "Average of numeric args/ranges." },
  { name: "MIN", syntax: "MIN(...) ", desc: "Minimum numeric value." },
  { name: "MAX", syntax: "MAX(...) ", desc: "Maximum numeric value." },
  { name: "COUNT", syntax: "COUNT(...) ", desc: "Count numeric values." },
  { name: "COUNTA", syntax: "COUNTA(...) ", desc: "Count non-empty values." },
  { name: "MEDIAN", syntax: "MEDIAN(...) ", desc: "Median of numeric values." },
  { name: "PRODUCT", syntax: "PRODUCT(...) ", desc: "Multiply numeric values." },
  { name: "ABS", syntax: "ABS(x)", desc: "Absolute value." },
  { name: "ROUND", syntax: "ROUND(x, digits)", desc: "Round to digits." },
  { name: "ROUNDUP", syntax: "ROUNDUP(x, digits)", desc: "Round away from zero." },
  { name: "ROUNDDOWN", syntax: "ROUNDDOWN(x, digits)", desc: "Round toward zero." },
  { name: "INT", syntax: "INT(x)", desc: "Floor integer." },
  { name: "CEILING", syntax: "CEILING(x, step)", desc: "Round up to nearest step." },
  { name: "FLOOR", syntax: "FLOOR(x, step)", desc: "Round down to nearest step." },
  { name: "MOD", syntax: "MOD(a,b)", desc: "Remainder." },
  { name: "POW", syntax: "POW(a,b)", desc: "Power a^b." },
  { name: "SQRT", syntax: "SQRT(x)", desc: "Square root." },
  { name: "EXP", syntax: "EXP(x)", desc: "e^x." },
  { name: "LN", syntax: "LN(x)", desc: "Natural logarithm." },
  { name: "LOG", syntax: "LOG(x, base)", desc: "Logarithm base b (default 10)." },
  { name: "LOG10", syntax: "LOG10(x)", desc: "Base-10 logarithm." },
  { name: "SIN", syntax: "SIN(x)", desc: "Sine." },
  { name: "COS", syntax: "COS(x)", desc: "Cosine." },
  { name: "TAN", syntax: "TAN(x)", desc: "Tangent." },
  { name: "ASIN", syntax: "ASIN(x)", desc: "Arc sine." },
  { name: "ACOS", syntax: "ACOS(x)", desc: "Arc cosine." },
  { name: "ATAN", syntax: "ATAN(x)", desc: "Arc tangent." },
  { name: "PI", syntax: "PI()", desc: "π constant." },
  { name: "RAND", syntax: "RAND()", desc: "Random number [0,1)." },
  { name: "RANDBETWEEN", syntax: "RANDBETWEEN(min,max)", desc: "Random integer in range." },
  { name: "IF", syntax: "IF(cond, a, b)", desc: "Conditional expression." },
  { name: "AND", syntax: "AND(a,b,...)", desc: "Logical and." },
  { name: "OR", syntax: "OR(a,b,...)", desc: "Logical or." },
  { name: "NOT", syntax: "NOT(a)", desc: "Logical negation." },
  { name: "LEN", syntax: "LEN(text)", desc: "String length." },
  { name: "UPPER", syntax: "UPPER(text)", desc: "Uppercase text." },
  { name: "LOWER", syntax: "LOWER(text)", desc: "Lowercase text." },
  { name: "TRIM", syntax: "TRIM(text)", desc: "Trim + collapse spaces." },
  { name: "LEFT", syntax: "LEFT(text, n)", desc: "Left n chars." },
  { name: "RIGHT", syntax: "RIGHT(text, n)", desc: "Right n chars." },
  { name: "MID", syntax: "MID(text, start, n)", desc: "Substring." },
  { name: "CONCAT", syntax: "CONCAT(a,b,...)", desc: "Concatenate text." },
  { name: "TODAY", syntax: "TODAY()", desc: "Current date (YYYY-MM-DD)." },
  { name: "NOW", syntax: "NOW()", desc: "Current date-time (ISO)." },
  { name: "YEAR", syntax: "YEAR(date)", desc: "Year from date." },
  { name: "MONTH", syntax: "MONTH(date)", desc: "Month from date." },
  { name: "DAY", syntax: "DAY(date)", desc: "Day from date." },
  { name: "HOUR", syntax: "HOUR(date)", desc: "Hour from date." },
  { name: "MINUTE", syntax: "MINUTE(date)", desc: "Minute from date." },
  { name: "SECOND", syntax: "SECOND(date)", desc: "Second from date." },
  { name: "DATE", syntax: "DATE(y,m,d)", desc: "Build date string." }
];

const sheetTable = document.getElementById("sheetTable");
const sheetWrap = document.getElementById("sheetWrap");
const fillHandle = document.getElementById("fillHandle");
const formulaInput = document.getElementById("formulaInput");
const cellAddress = document.getElementById("cellAddress");
const statusText = document.getElementById("statusText");
const selectionMetrics = document.getElementById("selectionMetrics");
const sheetTabs = document.getElementById("sheetTabs");
const addSheetBtn = document.getElementById("addSheetBtn");
const functionsCatalogBtn = document.getElementById("functionsCatalogBtn");
const functionsCatalogPanel = document.getElementById("functionsCatalogPanel");
const functionsCatalogCloseBtn = document.getElementById("functionsCatalogCloseBtn");
const catalogSearchInput = document.getElementById("catalogSearchInput");
const catalogList = document.getElementById("catalogList");

const saveBtn = document.getElementById("saveBtn");
const loadBtn = document.getElementById("loadBtn");
const newSheetBtn = document.getElementById("newSheetBtn");
const undoBtn = document.getElementById("undoBtn");
const redoBtn = document.getElementById("redoBtn");
const downloadCsvBtn = document.getElementById("downloadCsvBtn");
const downloadXlsxBtn = document.getElementById("downloadXlsxBtn");
const importCsvInput = document.getElementById("importCsvInput");
const importXlsxInput = document.getElementById("importXlsxInput");

const freezeRowsInput = document.getElementById("freezeRowsInput");
const freezeColsInput = document.getElementById("freezeColsInput");
const applyFreezeBtn = document.getElementById("applyFreezeBtn");

const boldBtn = document.getElementById("boldBtn");
const italicBtn = document.getElementById("italicBtn");
const textColorInput = document.getElementById("textColorInput");
const fillColorInput = document.getElementById("fillColorInput");

const workbook = { sheets: [], activeSheetId: null };
let activeCell = "A1";
let selectionRange = { start: "A1", end: "A1" };
let isSelecting = false;
let isFillDragging = false;
let fillPreviewCell = null;
let isRestoringHistory = false;
const undoStack = [];
const redoStack = [];

function hasXlsxLib() { return typeof XLSX !== "undefined"; }
function toColName(index) { let n = index + 1; let s = ""; while (n > 0) { const r = (n - 1) % 26; s = String.fromCharCode(65 + r) + s; n = Math.floor((n - 1) / 26);} return s; }
function colNameToIndex(name) { let n = 0; for (const ch of name) n = n * 26 + (ch.charCodeAt(0) - 64); return n - 1; }
function toAddress(row, col) { return `${toColName(col)}${row + 1}`; }
function parseAddress(address) { const m = /^([A-Z]+)(\d+)$/i.exec(address.trim()); if (!m) return null; return { col: colNameToIndex(m[1].toUpperCase()), row: Number(m[2]) - 1 }; }

function normalizeRange(startAddr, endAddr) {
  const a = parseAddress(startAddr); const b = parseAddress(endAddr); if (!a || !b) return null;
  return { rowStart: Math.min(a.row, b.row), rowEnd: Math.max(a.row, b.row), colStart: Math.min(a.col, b.col), colEnd: Math.max(a.col, b.col) };
}

function defaultCellModel() { return { raw: "", computed: "", style: { bold: false, italic: false, textColor: "#1f2933", fillColor: "#ffffff" } }; }
function createSheet(name) { return { id: crypto.randomUUID(), name, cells: {}, freezeRows: 0, freezeCols: 0, colWidths: Array.from({ length: COLS }, () => DEFAULT_COL_WIDTH) }; }

function ensureSheetDefaults(sheet) {
  if (!Array.isArray(sheet.colWidths) || sheet.colWidths.length !== COLS) {
    sheet.colWidths = Array.from({ length: COLS }, (_, i) => Number(sheet.colWidths?.[i]) || DEFAULT_COL_WIDTH);
  }
  sheet.freezeRows = Math.max(0, Math.min(ROWS - 1, Number(sheet.freezeRows) || 0));
  sheet.freezeCols = Math.max(0, Math.min(COLS - 1, Number(sheet.freezeCols) || 0));
}

function snapshotState() { return JSON.stringify({ workbook, activeCell, selectionRange }); }
function syncUndoRedoButtons() { undoBtn.disabled = undoStack.length === 0; redoBtn.disabled = redoStack.length === 0; }
function recordHistory() {
  if (isRestoringHistory) return;
  undoStack.push(snapshotState());
  if (undoStack.length > MAX_HISTORY) undoStack.shift();
  redoStack.length = 0;
  syncUndoRedoButtons();
}

function restoreState(snap) {
  const parsed = JSON.parse(snap);
  workbook.sheets = parsed.workbook.sheets || [];
  workbook.activeSheetId = parsed.workbook.activeSheetId;
  activeCell = parsed.activeCell || "A1";
  selectionRange = parsed.selectionRange || { start: activeCell, end: activeCell };
  ensureActiveSheet();
  workbook.sheets.forEach(ensureSheetDefaults);
  refreshView();
}

function undo() { if (!undoStack.length) return; redoStack.push(snapshotState()); isRestoringHistory = true; restoreState(undoStack.pop()); isRestoringHistory = false; statusText.textContent = "Undo"; syncUndoRedoButtons(); }
function redo() { if (!redoStack.length) return; undoStack.push(snapshotState()); isRestoringHistory = true; restoreState(redoStack.pop()); isRestoringHistory = false; statusText.textContent = "Redo"; syncUndoRedoButtons(); }

function getActiveSheet() { return workbook.sheets.find((s) => s.id === workbook.activeSheetId) ?? workbook.sheets[0]; }
function ensureActiveSheet() {
  if (!workbook.sheets.length) {
    const first = createSheet("Sheet 1"); workbook.sheets.push(first); workbook.activeSheetId = first.id;
  }
  if (!workbook.activeSheetId || !getActiveSheet()) workbook.activeSheetId = workbook.sheets[0].id;
}
function getCellModelFromSheet(sheet, addr) { if (!sheet.cells[addr]) sheet.cells[addr] = defaultCellModel(); return sheet.cells[addr]; }
function getCellModel(addr) { return getCellModelFromSheet(getActiveSheet(), addr); }

function buildGrid() {
  sheetTable.innerHTML = "";
  const headerRow = document.createElement("tr");
  const corner = document.createElement("th");
  corner.className = "header-cell row-header corner";
  headerRow.appendChild(corner);

  for (let c = 0; c < COLS; c += 1) {
    const th = document.createElement("th");
    th.className = "header-cell resizable";
    th.dataset.col = String(c);
    th.textContent = toColName(c);
    const resize = document.createElement("span");
    resize.className = "resize-handle";
    resize.dataset.col = String(c);
    resize.addEventListener("mousedown", startColumnResize);
    th.appendChild(resize);
    headerRow.appendChild(th);
  }
  sheetTable.appendChild(headerRow);

  for (let r = 0; r < ROWS; r += 1) {
    const tr = document.createElement("tr");
    const rowHeader = document.createElement("th");
    rowHeader.className = "row-header";
    rowHeader.textContent = String(r + 1);
    rowHeader.dataset.row = String(r);
    tr.appendChild(rowHeader);

    for (let c = 0; c < COLS; c += 1) {
      const td = document.createElement("td");
      const addr = toAddress(r, c);
      td.className = "sheet-cell";
      td.dataset.address = addr;
      td.dataset.row = String(r);
      td.dataset.col = String(c);
      td.contentEditable = "true";
      td.spellcheck = false;
      td.addEventListener("focus", () => onCellFocus(addr));
      td.addEventListener("input", () => onCellInput(addr, td.textContent ?? ""));
      td.addEventListener("keydown", onCellKeyDown);
      td.addEventListener("mousedown", onCellMouseDown);
      td.addEventListener("mouseenter", onCellMouseEnter);
      tr.appendChild(td);
    }
    sheetTable.appendChild(tr);
  }
  document.addEventListener("mouseup", onGlobalMouseUp);
}

function getCellElement(address) { return sheetTable.querySelector(`[data-address="${address}"]`); }
function setActiveRange(start, end = start) { selectionRange = { start, end }; renderRangeSelection(); }
function onCellMouseDown(event) { if (event.button !== 0) return; const addr = event.currentTarget.dataset.address; setActiveRange(addr, addr); isSelecting = true; }
function onCellMouseEnter(event) { if (!isSelecting) return; setActiveRange(selectionRange.start, event.currentTarget.dataset.address); }
function onGlobalMouseUp() { isSelecting = false; if (isFillDragging) { applyFillFromSelection(); isFillDragging = false; fillPreviewCell = null; renderRangeSelection(); } }

function onCellFocus(address) {
  activeCell = address;
  cellAddress.textContent = address;
  formulaInput.value = getCellModel(address).raw;
  syncFormatInputs(getCellModel(address).style);
  setActiveRange(address, address);
  updateMetrics();
}

function onCellInput(address, userText) { recordHistory(); getCellModel(address).raw = userText; recalculateSheet(); formulaInput.value = getCellModel(address).raw; }
function onCellKeyDown(event) {
  const parsed = parseAddress(event.target.dataset.address); if (!parsed) return;
  let target = null;
  if (event.key === "ArrowRight") target = toAddress(parsed.row, Math.min(parsed.col + 1, COLS - 1));
  if (event.key === "ArrowLeft") target = toAddress(parsed.row, Math.max(parsed.col - 1, 0));
  if (event.key === "ArrowDown" || event.key === "Enter") target = toAddress(Math.min(parsed.row + 1, ROWS - 1), parsed.col);
  if (event.key === "ArrowUp") target = toAddress(Math.max(parsed.row - 1, 0), parsed.col);
  if (target) { event.preventDefault(); getCellElement(target)?.focus(); }
}

function syncFormatInputs(style) {
  boldBtn.style.background = style.bold ? "#dce9ff" : "#fff";
  italicBtn.style.background = style.italic ? "#dce9ff" : "#fff";
  textColorInput.value = style.textColor;
  fillColorInput.value = style.fillColor;
}

function applyStyleToCell(address, patch) { recordHistory(); const m = getCellModel(address); m.style = { ...m.style, ...patch }; renderCell(address); syncFormatInputs(m.style); }

function tokenize(expr) {
  const tokens = []; let i = 0;
  while (i < expr.length) {
    const ch = expr[i];
    if (/\s/.test(ch)) { i += 1; continue; }
    if (/[+\-*/(),:]/.test(ch)) { tokens.push({ type: ch, value: ch }); i += 1; continue; }
    if (ch === '"') {
      let j = i + 1; let out = "";
      while (j < expr.length) {
        if (expr[j] === '"' && expr[j + 1] === '"') { out += '"'; j += 2; continue; }
        if (expr[j] === '"') break;
        out += expr[j]; j += 1;
      }
      if (j >= expr.length || expr[j] !== '"') throw new Error("Unclosed string");
      tokens.push({ type: "STRING", value: out });
      i = j + 1;
      continue;
    }
    if (/\d|\./.test(ch)) {
      let j = i; while (j < expr.length && /[\d.]/.test(expr[j])) j += 1;
      const n = Number(expr.slice(i, j)); if (!Number.isFinite(n)) throw new Error("Invalid number");
      tokens.push({ type: "NUMBER", value: n }); i = j; continue;
    }
    if (/[A-Za-z_]/.test(ch)) {
      let j = i; while (j < expr.length && /[A-Za-z0-9_]/.test(expr[j])) j += 1;
      const raw = expr.slice(i, j).toUpperCase();
      if (/^[A-Z]+\d+$/.test(raw)) tokens.push({ type: "CELL", value: raw });
      else tokens.push({ type: "IDENT", value: raw });
      i = j; continue;
    }
    throw new Error(`Unexpected character: ${ch}`);
  }
  tokens.push({ type: "EOF", value: null });
  return tokens;
}

class FormulaParser {
  constructor(tokens) { this.tokens = tokens; this.pos = 0; }
  current() { return this.tokens[this.pos]; }
  eat(type) { if (this.current().type !== type) throw new Error(`Expected ${type}`); return this.tokens[this.pos++]; }
  parse() { const node = this.parseExpr(); if (this.current().type !== "EOF") throw new Error("Unexpected token"); return node; }
  parseExpr() { let node = this.parseTerm(); while (["+", "-"].includes(this.current().type)) { const op = this.eat(this.current().type).type; node = { type: "Binary", op, left: node, right: this.parseTerm() }; } return node; }
  parseTerm() { let node = this.parseFactor(); while (["*", "/"].includes(this.current().type)) { const op = this.eat(this.current().type).type; node = { type: "Binary", op, left: node, right: this.parseFactor() }; } return node; }
  parseFactor() {
    if (["+", "-"].includes(this.current().type)) { const op = this.eat(this.current().type).type; return { type: "Unary", op, value: this.parseFactor() }; }
    if (this.current().type === "NUMBER") return { type: "Number", value: this.eat("NUMBER").value };
    if (this.current().type === "STRING") return { type: "String", value: this.eat("STRING").value };
    if (this.current().type === "CELL") {
      const start = this.eat("CELL").value;
      if (this.current().type === ":") { this.eat(":"); return { type: "Range", start, end: this.eat("CELL").value }; }
      return { type: "Cell", ref: start };
    }
    if (this.current().type === "IDENT") {
      const name = this.eat("IDENT").value;
      this.eat("(");
      const args = [];
      if (this.current().type !== ")") { args.push(this.parseExpr()); while (this.current().type === ",") { this.eat(","); args.push(this.parseExpr()); } }
      this.eat(")");
      return { type: "Call", name, args };
    }
    if (this.current().type === "(") { this.eat("("); const node = this.parseExpr(); this.eat(")"); return node; }
    throw new Error("Invalid expression");
  }
}

function toNumber(v) { const n = Number(v); return Number.isFinite(n) ? n : NaN; }
function isTruthy(v) {
  if (Array.isArray(v)) return v.some(isTruthy);
  if (typeof v === "string") return v.length > 0 && v !== "0";
  if (typeof v === "number") return v !== 0 && !Number.isNaN(v);
  return Boolean(v);
}
function flattenValues(v) { return Array.isArray(v) ? v.flatMap(flattenValues) : [v]; }
function flattenNumbers(v) { return flattenValues(v).map(toNumber).filter((n) => Number.isFinite(n)); }
function flattenNonEmpty(v) { return flattenValues(v).filter((x) => !(x === "" || x === null || x === undefined)); }

function evaluateReferencedValue(ref, sheet, stack) {
  const p = parseAddress(ref);
  if (!p || p.row < 0 || p.row >= ROWS || p.col < 0 || p.col >= COLS) return 0;
  const addr = toAddress(p.row, p.col);
  if (stack.has(addr)) return 0;
  stack.add(addr);
  const raw = getCellModelFromSheet(sheet, addr).raw;
  const out = evaluateFormula(raw, sheet, stack);
  stack.delete(addr);
  return out;
}

function collectRangeValues(startAddr, endAddr, sheet, stack) {
  const range = normalizeRange(startAddr, endAddr);
  if (!range) return [];
  const out = [];
  for (let r = range.rowStart; r <= range.rowEnd; r += 1) {
    for (let c = range.colStart; c <= range.colEnd; c += 1) out.push(evaluateReferencedValue(toAddress(r, c), sheet, stack));
  }
  return out;
}

function evaluateFunction(name, args) {
  const nums = () => flattenNumbers(args);
  const text = (i = 0) => String(flattenValues(args[i] ?? "")[0] ?? "");
  const narg = (i = 0, d = NaN) => { const n = toNumber(flattenValues(args[i] ?? d)[0]); return Number.isFinite(n) ? n : d; };

  if (name === "SUM") return nums().reduce((a, n) => a + n, 0);
  if (name === "AVG" || name === "AVERAGE") { const v = nums(); return v.length ? v.reduce((a, n) => a + n, 0) / v.length : 0; }
  if (name === "MIN") { const v = nums(); return v.length ? Math.min(...v) : 0; }
  if (name === "MAX") { const v = nums(); return v.length ? Math.max(...v) : 0; }
  if (name === "COUNT") return nums().length;
  if (name === "COUNTA") return flattenNonEmpty(args).length;
  if (name === "MEDIAN") { const v = nums().sort((a, b) => a - b); if (!v.length) return 0; const mid = Math.floor(v.length / 2); return v.length % 2 ? v[mid] : (v[mid - 1] + v[mid]) / 2; }
  if (name === "PRODUCT") { const v = nums(); return v.length ? v.reduce((a, n) => a * n, 1) : 0; }

  if (name === "ABS") return Math.abs(narg(0, 0));
  if (name === "ROUND") { const d = narg(1, 0); const p = 10 ** d; return Math.round(narg(0, 0) * p) / p; }
  if (name === "ROUNDUP") { const d = narg(1, 0); const p = 10 ** d; const v = narg(0, 0) * p; return (v >= 0 ? Math.ceil(v) : Math.floor(v)) / p; }
  if (name === "ROUNDDOWN") { const d = narg(1, 0); const p = 10 ** d; const v = narg(0, 0) * p; return (v >= 0 ? Math.floor(v) : Math.ceil(v)) / p; }
  if (name === "INT") return Math.floor(narg(0, 0));
  if (name === "CEILING") { const step = narg(1, 1) || 1; return Math.ceil(narg(0, 0) / step) * step; }
  if (name === "FLOOR") { const step = narg(1, 1) || 1; return Math.floor(narg(0, 0) / step) * step; }
  if (name === "MOD") return narg(0, 0) % narg(1, 1);
  if (name === "POW") return narg(0, 0) ** narg(1, 1);
  if (name === "SQRT") return Math.sqrt(narg(0, 0));
  if (name === "EXP") return Math.exp(narg(0, 0));
  if (name === "LN") return Math.log(narg(0, 1));
  if (name === "LOG") { const x = narg(0, 1); const b = narg(1, 10); return Math.log(x) / Math.log(b); }
  if (name === "LOG10") return Math.log10(narg(0, 1));
  if (name === "SIN") return Math.sin(narg(0, 0));
  if (name === "COS") return Math.cos(narg(0, 0));
  if (name === "TAN") return Math.tan(narg(0, 0));
  if (name === "ASIN") return Math.asin(narg(0, 0));
  if (name === "ACOS") return Math.acos(narg(0, 0));
  if (name === "ATAN") return Math.atan(narg(0, 0));
  if (name === "PI") return Math.PI;
  if (name === "RAND") return Math.random();
  if (name === "RANDBETWEEN") { const a = Math.floor(narg(0, 0)); const b = Math.floor(narg(1, 1)); const lo = Math.min(a, b); const hi = Math.max(a, b); return Math.floor(Math.random() * (hi - lo + 1)) + lo; }

  if (name === "IF") return isTruthy(args[0]) ? (args[1] ?? "") : (args[2] ?? "");
  if (name === "AND") return args.every(isTruthy) ? 1 : 0;
  if (name === "OR") return args.some(isTruthy) ? 1 : 0;
  if (name === "NOT") return isTruthy(args[0]) ? 0 : 1;

  if (name === "LEN") return text(0).length;
  if (name === "UPPER") return text(0).toUpperCase();
  if (name === "LOWER") return text(0).toLowerCase();
  if (name === "TRIM") return text(0).trim().replace(/\s+/g, " ");
  if (name === "LEFT") return text(0).slice(0, Math.max(0, Math.floor(narg(1, 1))));
  if (name === "RIGHT") { const n = Math.max(0, Math.floor(narg(1, 1))); const t = text(0); return t.slice(Math.max(0, t.length - n)); }
  if (name === "MID") { const t = text(0); const start = Math.max(1, Math.floor(narg(1, 1))); const n = Math.max(0, Math.floor(narg(2, 1))); return t.slice(start - 1, start - 1 + n); }
  if (name === "CONCAT") return flattenValues(args).join("");

  if (name === "TODAY") return new Date().toISOString().slice(0, 10);
  if (name === "NOW") return new Date().toISOString();
  if (name === "YEAR") return new Date(text(0)).getFullYear();
  if (name === "MONTH") return new Date(text(0)).getMonth() + 1;
  if (name === "DAY") return new Date(text(0)).getDate();
  if (name === "HOUR") return new Date(text(0)).getHours();
  if (name === "MINUTE") return new Date(text(0)).getMinutes();
  if (name === "SECOND") return new Date(text(0)).getSeconds();
  if (name === "DATE") {
    const y = Math.floor(narg(0, 1970));
    const m = Math.max(1, Math.floor(narg(1, 1)));
    const d = Math.max(1, Math.floor(narg(2, 1)));
    return new Date(Date.UTC(y, m - 1, d)).toISOString().slice(0, 10);
  }

  throw new Error(`Unknown function: ${name}`);
}

function evalAst(ast, sheet, stack) {
  if (ast.type === "Number") return ast.value;
  if (ast.type === "String") return ast.value;
  if (ast.type === "Unary") {
    const n = toNumber(evalAst(ast.value, sheet, stack));
    if (!Number.isFinite(n)) return NaN;
    return ast.op === "-" ? -n : n;
  }
  if (ast.type === "Binary") {
    const a = toNumber(evalAst(ast.left, sheet, stack));
    const b = toNumber(evalAst(ast.right, sheet, stack));
    if (!Number.isFinite(a) || !Number.isFinite(b)) return NaN;
    if (ast.op === "+") return a + b;
    if (ast.op === "-") return a - b;
    if (ast.op === "*") return a * b;
    if (ast.op === "/") return b === 0 ? NaN : a / b;
    return NaN;
  }
  if (ast.type === "Cell") return evaluateReferencedValue(ast.ref, sheet, stack);
  if (ast.type === "Range") return collectRangeValues(ast.start, ast.end, sheet, stack);
  if (ast.type === "Call") {
    const args = ast.args.map((a) => evalAst(a, sheet, stack));
    return evaluateFunction(ast.name, args);
  }
  return NaN;
}

function evaluateFormula(raw, sheet = getActiveSheet(), stack = new Set()) {
  if (!raw.startsWith("=")) return raw;
  try {
    const ast = new FormulaParser(tokenize(raw.slice(1).trim())).parse();
    const v = evalAst(ast, sheet, stack);
    if (v === undefined || v === null || (typeof v === "number" && Number.isNaN(v))) return "#ERR";
    return String(v);
  } catch {
    return "#ERR";
  }
}

function renderCell(addr) {
  const el = getCellElement(addr);
  if (!el) return;
  const m = getCellModel(addr);
  el.textContent = m.computed;
  el.style.fontWeight = m.style.bold ? "700" : "400";
  el.style.fontStyle = m.style.italic ? "italic" : "normal";
  el.style.color = m.style.textColor;
  el.style.backgroundColor = m.style.fillColor;
}

function recalculateSheet() {
  const sheet = getActiveSheet();
  Object.keys(sheet.cells).forEach((addr) => { sheet.cells[addr].computed = evaluateFormula(sheet.cells[addr].raw, sheet, new Set([addr])); });
  for (let r = 0; r < ROWS; r += 1) for (let c = 0; c < COLS; c += 1) renderCell(toAddress(r, c));
  applyColumnWidths();
  applyFrozenPanes();
  renderRangeSelection();
  updateMetrics();
}

function applyColumnWidths() {
  const sheet = getActiveSheet();
  for (let c = 0; c < COLS; c += 1) {
    const w = Math.max(48, Number(sheet.colWidths[c]) || DEFAULT_COL_WIDTH);
    sheet.colWidths[c] = w;
    const h = sheetTable.querySelector(`.header-cell[data-col="${c}"]`);
    if (h) h.style.width = `${w}px`;
    sheetTable.querySelectorAll(`.sheet-cell[data-col="${c}"]`).forEach((cell) => {
      cell.style.width = `${w}px`;
      cell.style.minWidth = `${w}px`;
      cell.style.maxWidth = `${w}px`;
    });
  }
}

function getLeftOffsetForCol(col) { const s = getActiveSheet(); let left = 46; for (let i = 0; i < col; i += 1) left += Number(s.colWidths[i]) || DEFAULT_COL_WIDTH; return left; }
function applyFrozenPanes() {
  const s = getActiveSheet(); const fr = s.freezeRows; const fc = s.freezeCols; const headerH = 28;
  sheetTable.querySelectorAll(".frozen-cell").forEach((n) => { n.classList.remove("frozen-cell"); n.style.top = ""; n.style.left = ""; n.style.zIndex = ""; });
  sheetTable.querySelectorAll(".row-header[data-row]").forEach((h) => {
    const r = Number(h.dataset.row);
    if (r < fr) { h.classList.add("frozen-cell"); h.style.top = `${headerH + r * 28}px`; h.style.left = "0px"; h.style.zIndex = "6"; }
  });
  for (let c = 0; c < fc; c += 1) {
    const h = sheetTable.querySelector(`.header-cell[data-col="${c}"]`);
    if (h) { h.classList.add("frozen-cell"); h.style.left = `${getLeftOffsetForCol(c)}px`; h.style.top = "0px"; h.style.zIndex = "6"; }
  }
  sheetTable.querySelectorAll(".sheet-cell").forEach((cell) => {
    const r = Number(cell.dataset.row); const c = Number(cell.dataset.col);
    if (r < fr) { cell.classList.add("frozen-cell"); cell.style.top = `${headerH + r * 28}px`; }
    if (c < fc) { cell.classList.add("frozen-cell"); cell.style.left = `${getLeftOffsetForCol(c)}px`; }
    if (r < fr || c < fc) cell.style.zIndex = r < fr && c < fc ? "5" : "2";
  });
  freezeRowsInput.value = String(fr); freezeColsInput.value = String(fc);
}

function startColumnResize(event) {
  event.preventDefault(); event.stopPropagation();
  const col = Number(event.target.dataset.col); const s = getActiveSheet(); const startX = event.clientX; const startW = Number(s.colWidths[col]) || DEFAULT_COL_WIDTH;
  recordHistory();
  const onMove = (e) => { s.colWidths[col] = Math.max(48, startW + (e.clientX - startX)); applyColumnWidths(); applyFrozenPanes(); renderRangeSelection(); };
  const onUp = () => { document.removeEventListener("mousemove", onMove); document.removeEventListener("mouseup", onUp); };
  document.addEventListener("mousemove", onMove); document.addEventListener("mouseup", onUp);
}

function renderRangeSelection() {
  sheetTable.querySelectorAll(".range-selected").forEach((n) => n.classList.remove("range-selected"));
  sheetTable.querySelectorAll(".selected").forEach((n) => n.classList.remove("selected"));
  const range = normalizeRange(selectionRange.start, selectionRange.end); if (!range) return;
  for (let r = range.rowStart; r <= range.rowEnd; r += 1) for (let c = range.colStart; c <= range.colEnd; c += 1) getCellElement(toAddress(r, c))?.classList.add("range-selected");
  getCellElement(activeCell)?.classList.add("selected");
  positionFillHandle(range);
}

function positionFillHandle(range) {
  const target = getCellElement(toAddress(range.rowEnd, range.colEnd));
  if (!target) { fillHandle.style.display = "none"; return; }
  const t = target.getBoundingClientRect(); const wrap = sheetWrap.getBoundingClientRect();
  fillHandle.style.display = "block";
  fillHandle.style.left = `${t.right - wrap.left + sheetWrap.scrollLeft - 4}px`;
  fillHandle.style.top = `${t.bottom - wrap.top + sheetWrap.scrollTop - 4}px`;
}

function setFormulaForActiveCell(raw) { recordHistory(); getCellModel(activeCell).raw = raw; recalculateSheet(); getCellElement(activeCell)?.focus(); }
function updateMetrics() {
  const range = normalizeRange(selectionRange.start, selectionRange.end); if (!range) return;
  let cnt = 0; let sum = 0;
  for (let r = range.rowStart; r <= range.rowEnd; r += 1) for (let c = range.colStart; c <= range.colEnd; c += 1) { const n = toNumber(getCellModel(toAddress(r, c)).computed); if (Number.isFinite(n)) { cnt += 1; sum += n; }}
  selectionMetrics.textContent = cnt ? `Selected: ${cnt} num • Sum ${sum}` : `Range ${selectionRange.start}:${selectionRange.end}`;
}

function renderSheetTabs() {
  sheetTabs.innerHTML = "";
  workbook.sheets.forEach((sheet) => {
    const btn = document.createElement("button");
    btn.className = `sheet-tab ${sheet.id === workbook.activeSheetId ? "active" : ""}`;
    btn.textContent = sheet.name;
    btn.addEventListener("click", () => switchSheet(sheet.id));
    btn.addEventListener("dblclick", () => renameSheet(sheet.id));
    sheetTabs.appendChild(btn);
  });
}

function switchSheet(id) { workbook.activeSheetId = id; activeCell = "A1"; selectionRange = { start: activeCell, end: activeCell }; refreshView(); }
function addSheet() { recordHistory(); const s = createSheet(`Sheet ${workbook.sheets.length + 1}`); workbook.sheets.push(s); switchSheet(s.id); }
function renameSheet(id) { const s = workbook.sheets.find((x) => x.id === id); if (!s) return; const next = window.prompt("Rename sheet", s.name); if (!next) return; recordHistory(); s.name = next.trim() || s.name; renderSheetTabs(); }
function applyFreeze() { recordHistory(); const s = getActiveSheet(); s.freezeRows = Math.max(0, Math.min(ROWS - 1, Number(freezeRowsInput.value) || 0)); s.freezeCols = Math.max(0, Math.min(COLS - 1, Number(freezeColsInput.value) || 0)); applyFrozenPanes(); }

function saveWorkbook() { localStorage.setItem(STORAGE_KEY, JSON.stringify(workbook)); statusText.textContent = "Workbook saved locally"; }
function loadWorkbook() {
  const raw = localStorage.getItem(STORAGE_KEY); if (!raw) return;
  recordHistory();
  const parsed = JSON.parse(raw);
  workbook.sheets = parsed.sheets || []; workbook.activeSheetId = parsed.activeSheetId;
  ensureActiveSheet(); workbook.sheets.forEach(ensureSheetDefaults);
  activeCell = "A1"; selectionRange = { start: activeCell, end: activeCell }; refreshView();
}
function clearWorkbook() { recordHistory(); workbook.sheets = [createSheet("Sheet 1")]; workbook.activeSheetId = workbook.sheets[0].id; activeCell = "A1"; selectionRange = { start: activeCell, end: activeCell }; refreshView(); }

function toCsv() {
  const lines = [];
  for (let r = 0; r < ROWS; r += 1) {
    const vals = [];
    for (let c = 0; c < COLS; c += 1) vals.push(escapeCsv(getCellModel(toAddress(r, c)).raw));
    lines.push(vals.join(","));
  }
  return lines.join("\n");
}
function escapeCsv(text) { if (text.includes(",") || text.includes("\"") || text.includes("\n")) return `"${text.replaceAll("\"", "\"\"")}"`; return text; }
function downloadCsv() { const blob = new Blob([toCsv()], { type: "text/csv;charset=utf-8" }); const url = URL.createObjectURL(blob); const a = document.createElement("a"); a.href = url; a.download = `${getActiveSheet().name.replaceAll(/\s+/g, "-").toLowerCase() || "sheet"}.csv`; a.click(); URL.revokeObjectURL(url); }

function parseCsvLine(line) {
  const out = []; let cur = ""; let quoted = false;
  for (let i = 0; i < line.length; i += 1) {
    const ch = line[i];
    if (ch === '"') { if (quoted && line[i + 1] === '"') { cur += '"'; i += 1; } else quoted = !quoted; continue; }
    if (ch === "," && !quoted) { out.push(cur); cur = ""; continue; }
    cur += ch;
  }
  out.push(cur);
  return out;
}

function importCsvText(text) {
  recordHistory();
  const s = getActiveSheet();
  s.cells = {};
  text.split(/\r?\n/).slice(0, ROWS).forEach((line, r) => {
    parseCsvLine(line).slice(0, COLS).forEach((value, c) => { if (value) getCellModel(toAddress(r, c)).raw = value; });
  });
  recalculateSheet();
}

function convertSheetToXlsxWorksheet(sheet) {
  const ws = {}; let hasData = false;
  for (let r = 0; r < ROWS; r += 1) {
    for (let c = 0; c < COLS; c += 1) {
      const model = sheet.cells[toAddress(r, c)]; if (!model?.raw) continue;
      const ref = XLSX.utils.encode_cell({ r, c });
      const raw = model.raw;
      if (raw.startsWith("=") && raw.length > 1) ws[ref] = { f: raw.slice(1) };
      else { const n = Number(raw); ws[ref] = Number.isFinite(n) && raw.trim() !== "" ? { t: "n", v: n } : { t: "s", v: raw }; }
      hasData = true;
    }
  }
  ws["!ref"] = hasData ? `A1:${toAddress(ROWS - 1, COLS - 1)}` : "A1";
  return ws;
}

function downloadXlsx() {
  if (!hasXlsxLib()) return;
  const wb = XLSX.utils.book_new();
  workbook.sheets.forEach((sheet) => XLSX.utils.book_append_sheet(wb, convertSheetToXlsxWorksheet(sheet), sheet.name.slice(0, 31) || "Sheet"));
  XLSX.writeFile(wb, "sheetforge-workbook.xlsx");
}

function sanitizeSheetName(name, idx) { const cleaned = (name || "").replace(/[\\/?*\[\]:]/g, " ").trim(); return (cleaned || `Sheet ${idx + 1}`).slice(0, 31); }
function importXlsxWorkbook(buffer) {
  if (!hasXlsxLib()) return;
  recordHistory();
  const xlsxWb = XLSX.read(buffer, { type: "array", cellFormula: true });
  const nextSheets = [];
  xlsxWb.SheetNames.forEach((name, idx) => {
    const ws = xlsxWb.Sheets[name];
    const next = createSheet(sanitizeSheetName(name, idx));
    const range = ws["!ref"] ? XLSX.utils.decode_range(ws["!ref"]) : null;
    if (range) {
      for (let r = range.s.r; r <= Math.min(range.e.r, ROWS - 1); r += 1) {
        for (let c = range.s.c; c <= Math.min(range.e.c, COLS - 1); c += 1) {
          const cell = ws[XLSX.utils.encode_cell({ r, c })];
          if (!cell) continue;
          getCellModelFromSheet(next, toAddress(r, c)).raw = cell.f ? `=${cell.f}` : cell.v == null ? "" : String(cell.v);
        }
      }
    }
    nextSheets.push(next);
  });
  workbook.sheets = nextSheets.length ? nextSheets : [createSheet("Sheet 1")];
  workbook.activeSheetId = workbook.sheets[0].id;
  activeCell = "A1";
  selectionRange = { start: activeCell, end: activeCell };
  refreshView();
}

function startFillDrag(event) { event.preventDefault(); isFillDragging = true; fillPreviewCell = selectionRange.end; }
function trackFillPreview(event) {
  if (!isFillDragging) return;
  const target = document.elementFromPoint(event.clientX, event.clientY);
  const cell = target?.closest?.(".sheet-cell");
  if (!cell) return;
  fillPreviewCell = cell.dataset.address;
  selectionRange.end = fillPreviewCell;
  renderRangeSelection();
}
function applyFillFromSelection() {
  const src = normalizeRange(selectionRange.start, selectionRange.end);
  const end = parseAddress(fillPreviewCell || selectionRange.end);
  if (!src || !end) return;

  let dest = null;
  if (end.row > src.rowEnd) dest = { rowStart: src.rowEnd + 1, rowEnd: end.row, colStart: src.colStart, colEnd: src.colEnd };
  else if (end.row < src.rowStart) dest = { rowStart: end.row, rowEnd: src.rowStart - 1, colStart: src.colStart, colEnd: src.colEnd };
  else if (end.col > src.colEnd) dest = { rowStart: src.rowStart, rowEnd: src.rowEnd, colStart: src.colEnd + 1, colEnd: end.col };
  else if (end.col < src.colStart) dest = { rowStart: src.rowStart, rowEnd: src.rowEnd, colStart: end.col, colEnd: src.colStart - 1 };
  if (!dest) return;

  recordHistory();
  const sheet = getActiveSheet();
  const srcH = src.rowEnd - src.rowStart + 1;
  const srcW = src.colEnd - src.colStart + 1;

  for (let r = dest.rowStart; r <= dest.rowEnd; r += 1) {
    for (let c = dest.colStart; c <= dest.colEnd; c += 1) {
      const srcRow = src.rowStart + ((r - dest.rowStart) % srcH);
      const srcCol = src.colStart + ((c - dest.colStart) % srcW);
      const from = getCellModelFromSheet(sheet, toAddress(srcRow, srcCol));
      const to = getCellModelFromSheet(sheet, toAddress(r, c));
      to.raw = from.raw;
      to.style = { ...from.style };
    }
  }
  recalculateSheet();
}

function renderFunctionCatalog(filter = "") {
  const q = filter.trim().toUpperCase();
  const items = FUNCTIONS_CATALOG.filter((f) => !q || f.name.includes(q));
  catalogList.innerHTML = "";
  items.forEach((f) => {
    const div = document.createElement("div");
    div.className = "catalog-item";
    div.innerHTML = `<code>${f.name}</code> — ${f.syntax}<p>${f.desc}</p>`;
    catalogList.appendChild(div);
  });
}

function refreshView() {
  renderSheetTabs();
  recalculateSheet();
  formulaInput.value = getCellModel(activeCell).raw;
  cellAddress.textContent = activeCell;
  syncUndoRedoButtons();
}

formulaInput.addEventListener("keydown", (event) => {
  if (event.key === "Enter") {
    event.preventDefault();
    setFormulaForActiveCell(formulaInput.value);
  }
});

functionsCatalogBtn.addEventListener("click", () => { functionsCatalogPanel.classList.toggle("hidden"); renderFunctionCatalog(catalogSearchInput.value); });
functionsCatalogCloseBtn.addEventListener("click", () => functionsCatalogPanel.classList.add("hidden"));
catalogSearchInput.addEventListener("input", () => renderFunctionCatalog(catalogSearchInput.value));

document.addEventListener("keydown", (event) => {
  const isUndo = (event.ctrlKey || event.metaKey) && !event.shiftKey && event.key.toLowerCase() === "z";
  const isRedo = (event.ctrlKey || event.metaKey) && (event.key.toLowerCase() === "y" || (event.shiftKey && event.key.toLowerCase() === "z"));
  if (isUndo) { event.preventDefault(); undo(); }
  if (isRedo) { event.preventDefault(); redo(); }
});

document.addEventListener("mousemove", trackFillPreview);
fillHandle.addEventListener("mousedown", startFillDrag);
sheetWrap.addEventListener("scroll", () => renderRangeSelection());

saveBtn.addEventListener("click", saveWorkbook);
loadBtn.addEventListener("click", loadWorkbook);
newSheetBtn.addEventListener("click", clearWorkbook);
undoBtn.addEventListener("click", undo);
redoBtn.addEventListener("click", redo);
downloadCsvBtn.addEventListener("click", downloadCsv);
downloadXlsxBtn.addEventListener("click", downloadXlsx);
addSheetBtn.addEventListener("click", addSheet);
applyFreezeBtn.addEventListener("click", applyFreeze);

importCsvInput.addEventListener("change", async (event) => { const file = event.target.files?.[0]; if (!file) return; importCsvText(await file.text()); event.target.value = ""; });
importXlsxInput.addEventListener("change", async (event) => { const file = event.target.files?.[0]; if (!file) return; importXlsxWorkbook(await file.arrayBuffer()); event.target.value = ""; });

boldBtn.addEventListener("click", () => applyStyleToCell(activeCell, { bold: !getCellModel(activeCell).style.bold }));
italicBtn.addEventListener("click", () => applyStyleToCell(activeCell, { italic: !getCellModel(activeCell).style.italic }));
textColorInput.addEventListener("input", () => applyStyleToCell(activeCell, { textColor: textColorInput.value }));
fillColorInput.addEventListener("input", () => applyStyleToCell(activeCell, { fillColor: fillColorInput.value }));

buildGrid();
ensureActiveSheet();
workbook.sheets.forEach(ensureSheetDefaults);
selectionRange = { start: activeCell, end: activeCell };
renderFunctionCatalog("");
refreshView();
statusText.textContent = hasXlsxLib() ? "Ready" : "Ready (XLSX unavailable offline)";
