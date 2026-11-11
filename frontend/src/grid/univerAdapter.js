// Nơi duy nhất được import Univer. Chỉ dùng gói Apache-2.0, không có gói Pro nào.
import { createUniver, IUndoRedoService, LocaleType, mergeLocales } from '@univerjs/presets';
import { UniverSheetsCorePreset } from '@univerjs/preset-sheets-core';
import coreVi from '@univerjs/preset-sheets-core/locales/vi-VN';
import { UniverSheetsConditionalFormattingPreset } from '@univerjs/preset-sheets-conditional-formatting';
import cfVi from '@univerjs/preset-sheets-conditional-formatting/locales/vi-VN';
import { UniverSheetsDataValidationPreset } from '@univerjs/preset-sheets-data-validation';
import dvVi from '@univerjs/preset-sheets-data-validation/locales/vi-VN';
import { UniverSheetsDrawingPreset } from '@univerjs/preset-sheets-drawing';
import drawVi from '@univerjs/preset-sheets-drawing/locales/vi-VN';
import '@univerjs/preset-sheets-core/lib/index.css';
import '@univerjs/preset-sheets-conditional-formatting/lib/index.css';
import '@univerjs/preset-sheets-data-validation/lib/index.css';
import '@univerjs/preset-sheets-drawing/lib/index.css';
import './univerOverrides.css';
import { parseTypedNumber } from './typedNumber.js';
import { NUMFMT_LOCALE, usesOtherDecimal } from '../numberFormat.js';
import { cellsFromHtml, cellsFromText, interpretTable } from './pastedTable.js';

const HA = { left: 1, center: 2, right: 3 };
const VA = { top: 1, center: 2, bottom: 3 };
const BORDER = {
  thin: 1, hair: 2, dotted: 3, dashed: 4, dashDot: 5, dashDotDot: 6, double: 7, medium: 8,
  mediumDashed: 9, mediumDashDot: 10, mediumDashDotDot: 11, slantDashDot: 12, thick: 13,
};
const TOLERANCE = 1e-6;
const EMU_PER_PX = 9525;
const CALC_DEADLINE_MS = 60000;
// Có giá trị mới mà bộ tính chưa báo bắt đầu: coi là đang chờ tính trong khoảng này (T-DEV-08).
const CALC_START_GRACE_MS = 1000;
const CALC_STABLE_MS = 4000;
const DEFAULT_COL_CHARS = 8;
const GENERAL_MAX_CHARS = 11;

function toStyle(s) {
  const o = {};
  if (s.ff) o.ff = s.ff;
  if (s.fs) o.fs = s.fs;
  if (s.bl) o.bl = 1;
  if (s.it) o.it = 1;
  if (s.ul) o.ul = { s: 1 };
  if (s.st) o.st = { s: 1 };
  if (s.fc) o.cl = { rgb: s.fc };
  if (s.bg) o.bg = { rgb: s.bg };
  if (HA[s.ha]) o.ht = HA[s.ha];
  if (VA[s.va]) o.vt = VA[s.va];
  if (s.wrap) o.tb = 3;
  if (s.nf) o.n = { pattern: s.nf };
  if (s.bd) {
    o.bd = {};
    for (const [side, [style, color]] of Object.entries(s.bd)) o.bd[side] = { s: BORDER[style] ?? 1, cl: { rgb: color } };
  }
  return o;
}

// Mẫu số cho ô General có phần lẻ: làm tròn cho vừa cột như Excel và dấu thập phân theo ngôn ngữ (Q80).
// Chỉ đổi cách hiện, giá trị trong ô giữ nguyên. Số nguyên không cần mẫu. Trả null nếu không cần.
function generalPattern(value, colChars) {
  if (typeof value !== 'number' || !Number.isFinite(value) || Number.isInteger(value)) return null;
  const chars = Math.min(colChars, GENERAL_MAX_CHARS);
  const abs = Math.abs(value);
  if (abs >= 1e11 || abs < 1e-4) return '0.00E+00';
  const intDigits = abs < 1 ? 1 : Math.floor(Math.log10(abs)) + 1;
  const room = chars - intDigits - 1 - (value < 0 ? 1 : 0);
  if (room < 1) return intDigits > chars ? '0.00E+00' : '0';
  const shown = Number(abs.toFixed(room)).toString();
  const decimals = shown.includes('.') ? shown.split('.')[1].length : 0;
  return decimals ? `0.${'0'.repeat(decimals)}` : '0';
}

function colCharsOf(sheet) {
  const widths = new Map();
  for (const [lo, hi, width] of sheet.col_groups) {
    if (width) for (let k = lo; k <= hi && k < sheet.cols + 3; k++) widths.set(k, Math.floor(width));
  }
  return (c) => widths.get(c) ?? DEFAULT_COL_CHARS;
}

function toSnapshot(model) {
  const styles = Object.fromEntries(model.styles.map((s, i) => [`s${i}`, toStyle(s)]));
  const generalStyle = (i, pattern) => {
    const id = `g${i ?? ''}_${pattern}`;
    styles[id] ||= { ...(i === undefined ? {} : styles[`s${i}`]), n: { pattern } };
    return id;
  };
  const sheets = {};
  const order = [];
  model.sheets.forEach((sh, i) => {
    const id = `sheet${i}`;
    order.push(id);
    const cellData = {};
    const colChars = colCharsOf(sh);
    for (const c of sh.cells) {
      const cell = {};
      if (c.f) cell.f = c.f;
      if (c.v !== undefined && c.v !== null) {
        cell.v = c.v;
        if (typeof c.v === 'boolean') cell.t = 3;
      }
      const hasPattern = c.s !== undefined && model.styles[c.s].nf;
      const pattern = hasPattern ? null : generalPattern(c.f ? c.excel_v : c.v, colChars(c.c));
      if (pattern) cell.s = generalStyle(c.s, pattern);
      else if (c.s !== undefined) cell.s = `s${c.s}`;
      (cellData[c.r] ||= {})[c.c] = cell;
    }
    const columnData = {};
    for (const [lo, hi, width, hidden] of sh.col_groups) {
      for (let k = lo; k <= hi && k < sh.cols + 3; k++) {
        columnData[k] = { ...(width ? { w: Math.round(width * 7 + 5) } : {}), ...(hidden ? { hd: 1 } : {}) };
      }
    }
    const rowData = {};
    for (const [k, [height, hidden]] of Object.entries(sh.row_dims)) {
      rowData[k] = { ...(height ? { h: (height * 4) / 3 } : {}), ...(hidden ? { hd: 1 } : {}) };
    }
    sheets[id] = {
      id,
      name: sh.name,
      hidden: sh.hidden ? 1 : 0,
      rowCount: sh.rows + 5,
      columnCount: sh.cols + 3,
      cellData,
      columnData,
      rowData,
      mergeData: sh.merges.map(([a, b, c, d]) => ({ startRow: a, startColumn: b, endRow: c, endColumn: d })),
      freeze: sh.freeze ? { xSplit: sh.freeze.cols, ySplit: sh.freeze.rows, startRow: sh.freeze.rows, startColumn: sh.freeze.cols } : undefined,
    };
  });
  return { id: 'workbook', name: 'bieu-mau', appVersion: '1', locale: LocaleType.VI_VN, styles, sheetOrder: order, sheets };
}

function a1(r, c) {
  let col = '';
  for (let n = c + 1; n > 0; n = Math.floor((n - 1) / 26)) col = String.fromCharCode(65 + ((n - 1) % 26)) + col;
  return `${col}${r + 1}`;
}

function firstCellOf(sqref) {
  return sqref.split(/[\s:]/)[0].replaceAll('$', '');
}

// Quy tắc "giá trị ô so với số" đổi thành công thức tương đương, tương đối theo ô đầu vùng.
function cellIsFormula(rule) {
  const cell = firstCellOf(rule.sqref);
  const [f1, f2] = rule.formulas;
  switch (rule.operator) {
    case 'lessThan': return `${cell}<${f1}`;
    case 'lessThanOrEqual': return `${cell}<=${f1}`;
    case 'greaterThan': return `${cell}>${f1}`;
    case 'greaterThanOrEqual': return `${cell}>=${f1}`;
    case 'equal': return `${cell}=${f1}`;
    case 'notEqual': return `${cell}<>${f1}`;
    case 'between': return `AND(${cell}>=${f1},${cell}<=${f2})`;
    case 'notBetween': return `OR(${cell}<${f1},${cell}>${f2})`;
    default: return null;
  }
}

function sameValue(saved, computed) {
  if (saved === null || saved === undefined || saved === '') return computed === null || computed === undefined || computed === '' || computed === 0;
  if (typeof saved === 'boolean') return computed === saved || computed === (saved ? 1 : 0) || String(computed).toUpperCase() === String(saved).toUpperCase();
  if (typeof saved === 'number') {
    const b = Number(computed);
    return !Number.isNaN(b) && computed !== '' && Math.abs(saved - b) <= TOLERANCE * Math.max(1, Math.abs(saved));
  }
  return String(saved) === String(computed);
}

function applyRules(univerAPI, wb, model) {
  const problems = [];
  wb.getSheets().forEach((ws, i) => {
    const sh = model.sheets.find((s) => s.name === ws.getSheetName()) ?? model.sheets[i];
    for (const rule of [...sh.cf].sort((x, y) => x.priority - y.priority)) {
      try {
        const formula = rule.type === 'cellIs' ? cellIsFormula(rule) : rule.formulas[0];
        if (!formula) continue;
        const ranges = rule.sqref.split(/\s+/).map((ref) => ws.getRange(ref).getRange());
        let b = ws.newConditionalFormattingRule().whenFormulaSatisfied(`=${formula}`);
        if (rule.style.fc) b = b.setFontColor(rule.style.fc);
        if (rule.style.bg) b = b.setBackground(rule.style.bg);
        if (rule.style.bl) b = b.setBold(true);
        if (rule.style.it) b = b.setItalic(true);
        ws.addConditionalFormattingRule(b.setRanges(ranges).build());
      } catch (e) {
        problems.push(`${sh.name}!${rule.sqref}: ${e}`);
      }
    }
    for (const dv of sh.dv) {
      if (dv.type !== 'list' || !dv.f1) continue;
      try {
        const range = ws.getRange(dv.sqref.split(/\s+/)[0]);
        const inline = dv.f1.match(/^"(.*)"$/s);
        let rule;
        if (inline) {
          rule = univerAPI.newDataValidation().requireValueInList(inline[1].split(',')).build();
        } else {
          const m = dv.f1.match(/^(?:'?([^'!]+)'?!)?(.+)$/);
          const source = m[1] ? wb.getSheetByName(m[1]) : ws;
          rule = univerAPI.newDataValidation().requireValueInRange(source.getRange(m[2].replaceAll('$', ''))).build();
        }
        range.setDataValidation(rule);
      } catch (e) {
        problems.push(`${sh.name}!${dv.sqref}: ${e}`);
      }
    }
  });
  return problems;
}

async function applyImages(wb, model) {
  const problems = [];
  const sheets = wb.getSheets();
  for (let i = 0; i < sheets.length; i++) {
    const ws = sheets[i];
    const sh = model.sheets.find((s) => s.name === ws.getSheetName());
    for (const im of sh?.images ?? []) {
      try {
        let b = ws.newOverGridImage().setSource(im.src, 'BASE64').setColumn(im.col).setRow(im.row);
        if (im.cx && im.cy) b = b.setWidth(Math.round(im.cx / EMU_PER_PX)).setHeight(Math.round(im.cy / EMU_PER_PX));
        ws.insertImages([await b.buildAsync()]);
      } catch (e) {
        problems.push(`${sh.name}: ${e}`);
      }
    }
  }
  return problems;
}

// Chữ người dùng gõ có phải là số đang hiện khi sửa (đã bị cắt còn 12 chữ số) được sửa thêm không: F2, nhấp đúp, thanh công thức.
// Không dựa vào sự kiện bắt đầu sửa vì không đáng tin (đo 10/11). Phần đầu và phần cuối trùng gần hết là sửa trên số đang hiện;
// gõ thẳng một số mới thì không, vì số mới thay hẳn số cũ, không có chữ số nào bị cắt.
function editsShown(typed, shown) {
  const a = String(typed ?? '').trim().replace(',', '.');
  let head = 0;
  while (head < a.length && head < shown.length && a[head] === shown[head]) head++;
  let tail = 0;
  while (tail < a.length - head && tail < shown.length - head && a[a.length - 1 - tail] === shown[shown.length - 1 - tail]) tail++;
  return head + tail >= shown.length - 3;
}

export function createGrid(container, model, { readOnly = false, onAmbiguousNumber, onDecimalSwapped, onHiddenDigits, onLockedCell, onPasted, onChange, onEditDone, onCalculating } = {}) {
  const { univer, univerAPI } = createUniver({
    locale: LocaleType.VI_VN,
    locales: { [LocaleType.VI_VN]: mergeLocales(coreVi, cfVi, dvVi, drawVi) },
    presets: [
      UniverSheetsCorePreset({
        container,
        header: true,
        toolbar: false,
        contextMenu: false,
        formulaBar: true,
        footer: { sheetBar: true, statisticBar: false, menus: false, zoomSlider: false, addSheetButtonConfig: { show: false } },
        formula: { initialFormulaComputing: 0 }, // tự tính lại, không dựa vào giá trị đã lưu
      }),
      UniverSheetsConditionalFormattingPreset(),
      UniverSheetsDataValidationPreset(),
      UniverSheetsDrawingPreset(),
    ],
  });
  const wb = univerAPI.createWorkbook(toSnapshot(model));
  // Cách hiện số theo dấu thập phân của app (numberFormat.js, Q122). Chữ giao diện vẫn tiếng Việt.
  wb.setNumfmtLocal(NUMFMT_LOCALE);
  const undoRedo = univer.__getInjector().get(IUndoRedoService);
  let lastTyped = null;
  let undoBatch = null;
  let typedCount = 0;
  const endUndoBatch = () => {
    undoBatch?.dispose();
    undoBatch = null;
  };
  const problems = applyRules(univerAPI, wb, model);
  const imagesDone = applyImages(wb, model).then((p) => {
    problems.push(...p);
    wb.setEditable(!readOnly);
  });

  const sheetIdByName = (name) => wb.getSheetByName(name)?.getSheetId();
  const rawValue = (snapshot, sheetId, r, c) => snapshot.sheets[sheetId]?.cellData?.[r]?.[c]?.v ?? null;

  // Bộ tính chạy nền và có thể chạy nhiều lượt. Tính xong khi mọi ô công thức đều có kết quả,
  // hoặc số ô chưa có kết quả không giảm nữa (công thức thật sự ra rỗng thì phải dừng, không chờ mãi).
  // Xét cả ô không có số đã lưu: file do app ghi không có số nào, nếu bỏ qua thì coi như tính xong ngay.
  async function waitForCalculation() {
    const pending = (snap) => {
      let n = 0;
      for (const sh of model.sheets) {
        const id = sheetIdByName(sh.name);
        for (const c of sh.cells) if (c.f && c.excel_v !== '' && rawValue(snap, id, c.r, c.c) === null) n++;
      }
      return n;
    };
    let last = Infinity;
    let stableSince = Date.now();
    const deadline = Date.now() + CALC_DEADLINE_MS;
    for (;;) {
      if (disposed) return null;
      const snap = wb.save();
      const n = pending(snap);
      if (n === 0 || Date.now() > deadline) return snap;
      if (n < last) {
        last = n;
        stableSince = Date.now();
      } else if (Date.now() - stableSince > CALC_STABLE_MS) {
        return snap;
      }
      await new Promise((r) => setTimeout(r, 100));
    }
  }

  // Mẫu số của ô General phụ thuộc giá trị, nên phải tính lại sau mỗi lần tính. Dùng lệnh không vào lịch sử hoàn tác.
  const fixedPattern = new Set();
  model.sheets.forEach((sh) => sh.cells.forEach((c) => {
    if (c.s !== undefined && model.styles[c.s].nf) fixedPattern.add(`${sh.name}!${c.r},${c.c}`);
  }));
  const applied = new Map();
  let refreshTimer = null;
  let disposed = false;
  function refreshGeneralPatterns() {
    if (disposed) return;
    const snap = wb.save();
    for (const ws of wb.getSheets()) {
      const name = ws.getSheetName();
      const sh = model.sheets.find((x) => x.name === name);
      const colChars = sh ? colCharsOf(sh) : () => DEFAULT_COL_CHARS;
      const data = snap.sheets[ws.getSheetId()]?.cellData ?? {};
      const cellValue = {};
      let changed = false;
      for (const [r, row] of Object.entries(data)) {
        for (const [c, cell] of Object.entries(row)) {
          const key = `${name}!${r},${c}`;
          if (fixedPattern.has(key) || typeof cell?.v !== 'number') continue;
          const pattern = generalPattern(cell.v, colChars(Number(c))) ?? (applied.has(key) ? '0' : null);
          if (!pattern || applied.get(key) === pattern) continue;
          applied.set(key, pattern);
          (cellValue[r] ||= {})[c] = { s: { n: { pattern } } };
          changed = true;
        }
      }
      if (changed) univerAPI.executeCommand('sheet.mutation.set-range-values', { unitId: wb.getId(), subUnitId: ws.getSheetId(), cellValue });
    }
  }
  const scheduleRefresh = () => {
    clearTimeout(refreshTimer);
    refreshTimer = setTimeout(refreshGeneralPatterns, 50);
  };
  const calcListener = univerAPI.getFormula().calculationResultApplied(scheduleRefresh);

  // T-DEV-08: biết bộ tính còn chạy hay không, chỉ qua API công khai (calculationStart, calculationEnd).
  // Trước đây "tính xong" nghĩa là ô công thức có giá trị, nên sau một lần sửa, ô còn giữ số cũ vẫn bị coi là xong:
  // đo 10/11, bảng 18000 × 10 hiện số cũ khoảng 4 giây, không báo gì. QC tháng, bảng 300 × 25: dưới 0,1 giây.
  // Ô có giá trị mới thì đánh dấu chờ tính ngay, để không có khoảng hở trước khi bộ tính kịp báo bắt đầu.
  const formulaApi = univerAPI.getFormula();
  let calcRunning = false;
  let calcPendingAt = 0;
  let calcHold = false; // chỉ cho test: giả lập bộ tính không bao giờ báo xong
  let resultPending = false; // có lượt tính đã chạy mà settled() chưa chờ kết quả được ghi vào lưới
  let calcTimeoutMs = CALC_DEADLINE_MS;
  let lastBusy = false;
  const calcWaiters = new Set();
  const calcBusy = () => calcRunning || (calcPendingAt > 0 && Date.now() - calcPendingAt < CALC_START_GRACE_MS);
  function notifyCalc() {
    if (disposed) return;
    if (calcPendingAt && Date.now() - calcPendingAt >= CALC_START_GRACE_MS) calcPendingAt = 0;
    const busy = calcBusy();
    if (busy !== lastBusy) {
      lastBusy = busy;
      onCalculating?.(busy);
    }
    if (!busy) for (const wake of [...calcWaiters]) wake();
  }
  const calcStartListener = formulaApi.calculationStart(() => {
    calcRunning = true;
    resultPending = true;
    calcPendingAt = 0;
    notifyCalc();
  });
  const calcEndListener = formulaApi.calculationEnd(() => {
    if (calcHold) return;
    calcRunning = false;
    notifyCalc();
  });
  function markCalcPending() {
    if (calcRunning) return;
    calcPendingAt = Date.now();
    notifyCalc();
    setTimeout(notifyCalc, CALC_START_GRACE_MS + 20);
  }
  // Chờ bộ tính xong. Trả { ok: true } hoặc { ok: false, timedOut: true } khi quá hạn: không bao giờ chờ mãi.
  function waitCalcIdle(timeoutMs = calcTimeoutMs) {
    if (!calcBusy()) return Promise.resolve({ ok: true });
    return new Promise((resolve) => {
      let tick = null;
      let timer = null;
      const finish = (result) => {
        clearInterval(tick);
        clearTimeout(timer);
        calcWaiters.delete(wake);
        resolve(result);
      };
      const wake = () => finish({ ok: true });
      calcWaiters.add(wake);
      tick = setInterval(notifyCalc, 100);
      timer = setTimeout(() => finish({ ok: false, timedOut: true }), timeoutMs);
    });
  }

  const lockedAt = (name, r, c) => {
    const sh = model.sheets.find((x) => x.name === name);
    if (!sh?.protected) return false;
    return !sh.cells.some((x) => x.r === r && x.c === c && x.unlocked);
  };
  const a1Of = a1;
  const patternOf = (name, r, c) => {
    const cell = model.sheets.find((x) => x.name === name)?.cells.find((x) => x.r === r && x.c === c);
    return cell && cell.s !== undefined ? model.styles[cell.s].nf ?? null : null;
  };
  univerAPI.addEvent(univerAPI.Event.BeforeSheetEditStart, (p) => {
    const name = p.worksheet.getSheetName();
    if (lockedAt(name, p.row, p.column)) {
      p.cancel = true;
      onLockedCell?.({ sheet: name, cell: a1Of(p.row, p.column) });
    }
  });
  // Chỉ coi là sửa khi người dùng đã gõ. F2 rồi Enter không được đổi số, kể cả khi ô đang hiện làm tròn.
  let typedInEditor = false;
  let hiddenAsked = null;
  univerAPI.addEvent(univerAPI.Event.SheetEditChanging, () => {
    typedInEditor = true;
    hiddenAsked = null; // đổi gì sau khi đã hỏi thì phải hỏi lại
  });
  univerAPI.addEvent(univerAPI.Event.BeforeSheetEditEnd, (p) => {
    if (!p.isConfirm || !typedInEditor) return;
    const text = p.value?.toPlainText?.();
    const parsed = parseTypedNumber(text);
    if (parsed.kind === 'ambiguous' || parsed.kind === 'unclear') {
      p.cancel = true;
      onAmbiguousNumber?.({ sheet: p.worksheet.getSheetName(), cell: a1Of(p.row, p.column), text: parsed.text, unclear: parsed.kind === 'unclear' });
      return;
    }
    // Q122: số cũ có nhiều chữ số hơn ô đang sửa hiện. Lần Enter đầu chỉ hỏi; lần sau, không đổi gì, mới ghi.
    // Lúc đang sửa, getCellData() trả chuỗi đã cắt (đo 10/11: "3.141592654"), nên đọc số gốc từ bản lưu của sổ.
    if (parsed.kind !== 'number') return;
    const full = rawValue(wb.save(), p.worksheet.getSheetId(), p.row, p.column);
    if (parsed.kind !== 'number' || typeof full !== 'number') return;
    const shown = Number(full.toPrecision(12));
    if (shown === full || parsed.value === full || !editsShown(text, String(shown))) return;
    const key = `${p.worksheet.getSheetId()}!${p.row},${p.column}`;
    if (hiddenAsked?.key === key && hiddenAsked.text === text) {
      hiddenAsked = null;
      return;
    }
    p.cancel = true;
    hiddenAsked = { key, text };
    onHiddenDigits?.({ sheet: p.worksheet.getSheetName(), cell: a1Of(p.row, p.column), full, shown, next: parsed.value });
  });
  univerAPI.addEvent(univerAPI.Event.BeforeSheetEditEnd, (p) => {
    lastTyped = p.isConfirm && typedInEditor ? { ws: p.worksheet, row: p.row, column: p.column, text: p.value?.toPlainText?.() } : null;
    // Ghi ô và sửa số ở dưới phải là một bước hoàn tác: hoàn tác về số cũ, làm lại ra số đúng.
    endUndoBatch();
    if (lastTyped && parseTypedNumber(lastTyped.text).kind === 'number') undoBatch = undoRedo.beginUndoRedoGroup(wb.getId(), `typed-${++typedCount}`, 'replace');
  });
  univerAPI.addEvent(univerAPI.Event.SheetEditEnded, (p) => {
    // Gõ phím thẳng vào ô không qua sự kiện bắt đầu sửa, nên ô khóa phải được trả lại ở đây. Lưu vẫn bị backend từ chối.
    if (p.isConfirm && lockedAt(p.worksheet.getSheetName(), p.row, p.column)) {
      typedInEditor = false;
      endUndoBatch();
      lastTyped = null;
      onLockedCell?.({ sheet: p.worksheet.getSheetName(), cell: a1Of(p.row, p.column) });
      setTimeout(() => univerAPI.undo(), 0);
      return;
    }
    if (p.isConfirm && lastTyped && lastTyped.row === p.row && lastTyped.column === p.column) {
      const parsed = parseTypedNumber(lastTyped.text);
      // Thư viện hiểu số theo kiểu Mỹ, 5,5 thành 55. Đổi về đúng số theo chữ người dùng đã gõ.
      const current = p.worksheet.getRange(p.row, p.column).getCellData()?.v;
      if (parsed.kind === 'number' && current !== parsed.value) p.worksheet.getRange(p.row, p.column).setValue(parsed.value);
      // Gõ số lẻ bằng dấu khác dấu của app, ví dụ 5,5 khi app dùng dấu chấm: đã ghi 5.5, báo để người dùng biết (Q122).
      if (parsed.kind === 'number' && usesOtherDecimal(lastTyped.text)) onDecimalSwapped?.({ sheet: p.worksheet.getSheetName(), cell: a1Of(p.row, p.column), text: lastTyped.text.trim(), value: parsed.value });
      // Thư viện tự đoán mẫu số khi gõ, có thể làm ô 0.00 hiện 5.51 thành 6. Trả lại đúng mẫu của biểu mẫu.
      const nf = patternOf(p.worksheet.getSheetName(), p.row, p.column);
      if (parsed.kind === 'number' && nf) {
        univerAPI.executeCommand('sheet.mutation.set-range-values', { unitId: wb.getId(), subUnitId: p.worksheet.getSheetId(), cellValue: { [p.row]: { [p.column]: { s: { n: { pattern: nf } } } } } });
      }
    }
    endUndoBatch();
    lastTyped = null;
    typedInEditor = false;
    hiddenAsked = null;
    if (p.isConfirm) onEditDone?.();
    scheduleRefresh();
  });

  // Báo App khi nội dung hoặc trạng thái đang gõ đổi, gộp lại vì mỗi lần kiểm phải đọc cả bảng.
  let changeTimer = null;
  univerAPI.addEvent(univerAPI.Event.CommandExecuted, (e) => {
    if (e.id === 'sheet.mutation.set-range-values') markCalcPending();
    if (e.id !== 'sheet.mutation.set-range-values' && e.id !== 'sheet.operation.set-cell-edit-visible') return;
    clearTimeout(changeTimer);
    changeTimer = setTimeout(() => {
      if (!disposed) onChange?.();
    }, 150);
  });

  // Dán từ ngoài app: chỉ lấy giá trị, vì lưu chỉ ghi giá trị và số phải được hiểu đúng (Q80). Chép trong lưới để thư viện làm.
  univerAPI.addEvent(univerAPI.Event.BeforeClipboardPaste, (p) => {
    const { html, text } = p;
    if (html?.includes('data-copy-id')) return;
    const fromExcel = Boolean(html && /urn:schemas-microsoft-com:office:excel/.test(html));
    const cells = (html?.includes('<table') ? cellsFromHtml(html) : null) ?? (text ? cellsFromText(text) : null);
    if (!cells?.length) return;
    p.cancel = true;
    const ws = p.worksheet;
    const name = ws.getSheetName();
    const at = ws.getSelection()?.getActiveRange();
    if (!at) return;
    const top = at.getRow();
    const left = at.getColumn();
    const width = Math.max(...cells.map((row) => row.length));
    if (top + cells.length > ws.getMaxRows() || left + width > ws.getMaxColumns()) {
      onPasted?.({ sheet: name, tooBig: true, rows: cells.length, cols: width });
      return;
    }
    for (let r = 0; r < cells.length; r++) {
      for (let c = 0; c < width; c++) {
        if (lockedAt(name, top + r, left + c)) {
          onLockedCell?.({ sheet: name, cell: a1Of(top + r, left + c) });
          return;
        }
      }
    }
    const read = interpretTable(cells, { fromExcel, plain: fromExcel && text ? cellsFromText(text) : null });
    if (read.narrow.length) {
      onPasted?.({ sheet: name, narrow: read.narrow });
      return;
    }
    const data = read.rows.map((row) => Array.from({ length: width }, (_, c) => {
      const v = row[c] ?? null;
      return typeof v === 'string' ? { v, t: 1 } : { v }; // chữ từ ngoài không được thành công thức
    }));
    ws.getRange(top, left, cells.length, width).setValues(data);
    scheduleRefresh();
    const place = (x) => ({ cell: a1Of(top + x.r, left + x.c), text: x.text, value: x.value });
    onPasted?.({ sheet: name, guessed: read.guessed.map(place), unread: read.unread.map(place), rounded: read.rounded.map(place), swapped: read.swapped.map(place), unclear: read.unclear.map(place), decimal: read.decimal, hadEvidence: read.hadEvidence, fromExcel: read.fromExcel });
  });

  function collectEdits() {
    const snap = wb.save();
    const edits = [];
    for (const ws of wb.getSheets()) {
      const name = ws.getSheetName();
      const sh = model.sheets.find((x) => x.name === name);
      const orig = new Map((sh?.cells ?? []).map((c) => [`${c.r},${c.c}`, c]));
      const data = snap.sheets[ws.getSheetId()]?.cellData ?? {};
      const seen = new Set();
      for (const [r, row] of Object.entries(data)) {
        for (const [c, cell] of Object.entries(row)) {
          const key = `${r},${c}`;
          seen.add(key);
          const before = orig.get(key);
          const addr = a1(Number(r), Number(c));
          // Công thức chép bằng kéo có thể chỉ lưu mã dùng chung, nên hỏi lại đúng ô để có chữ công thức.
          const formula = cell?.f || (cell?.si ? ws.getRange(Number(r), Number(c)).getFormula() : null) || null;
          if (formula) {
            if (formula !== before?.f) edits.push({ sheet: name, cell: addr, formula: formula.replace(/^=/, '') });
            continue;
          }
          const value = cell?.v ?? null;
          if (before?.f) {
            edits.push({ sheet: name, cell: addr, value });
          } else {
            const old = before?.v ?? null;
            const same = value === old || (typeof old === 'boolean' && Number(old) === value);
            if (!same) edits.push({ sheet: name, cell: addr, value });
          }
        }
      }
      for (const [key, before] of orig) {
        if (!seen.has(key) && (before.f || before.v !== undefined)) {
          const [r, c] = key.split(',').map(Number);
          edits.push({ sheet: name, cell: a1(r, c), value: null });
        }
      }
    }
    return edits;
  }

  return {
    problems,
    collectEdits,
    isEditing: () => wb.isCellEditing(),
    // Xác nhận ô đang gõ dở. Trả false nếu ô vẫn đang mở, ví dụ số nhập nhằng đang chờ người dùng gõ lại.
    async commitEditing() {
      if (!wb.isCellEditing()) return true;
      await wb.endEditingAsync(true);
      await new Promise((r) => setTimeout(r, 50));
      return !wb.isCellEditing();
    },
    formulaValues() {
      const snap = wb.save();
      const out = [];
      for (const sh of model.sheets) {
        const id = sheetIdByName(sh.name);
        for (const c of sh.cells) if (c.f) out.push({ sheet: sh.name, cell: a1(c.r, c.c), value: rawValue(snap, id, c.r, c.c) });
      }
      return out;
    },
    undo: () => univerAPI.undo(),
    redo: () => univerAPI.redo(),
    formula: (sheet, ref) => wb.getSheetByName(sheet).getRange(ref).getFormula(),
    // Chờ ảnh, chờ bộ tính xong rồi kết quả được ghi vào lưới. Trả { ok } như waitCalcIdle: ok false là quá hạn, số có thể còn cũ.
    settled: async () => {
      await imagesDone;
      const calc = await waitCalcIdle();
      if (calc.ok && resultPending) {
        resultPending = false;
        await formulaApi.onCalculationResultApplied(5000).catch(() => {});
      }
      await waitForCalculation();
      refreshGeneralPatterns();
      return calc;
    },
    isCalculating: () => calcBusy(),
    // Chỉ cho test (T-DEV-08): giữ bộ tính như không bao giờ báo xong, và rút ngắn hạn chờ.
    testHoldCalculation(on) {
      calcHold = Boolean(on);
      calcRunning = calcHold; // bật: như bộ tính đã bắt đầu mà không bao giờ báo xong
      notifyCalc();
    },
    testCalcTimeout(ms) {
      calcTimeoutMs = ms ?? CALC_DEADLINE_MS;
    },
    visibleSheets: () => wb.getSheets().filter((ws) => !ws.isSheetHidden()).map((ws) => ws.getSheetName()),
    value(sheet, ref) {
      const ws = wb.getSheetByName(sheet);
      const range = ws.getRange(ref);
      return rawValue(wb.save(), ws.getSheetId(), range.getRow(), range.getColumn());
    },
    display: (sheet, ref) => wb.getSheetByName(sheet).getRange(ref).getDisplayValue(),
    allDisplay: () => model.sheets.flatMap((sh) => sh.cells.filter((c) => c.f || typeof c.v === 'number').map((c) => ({
      cell: `${sh.name}!${a1(c.r, c.c)}`,
      text: wb.getSheetByName(sh.name).getRange(c.r, c.c).getDisplayValue(),
    }))),
    stats: () => ({ formulas: model.sheets.reduce((n, sh) => n + sh.cells.filter((c) => c.f).length, 0) }),
    cfCount: (sheet) => wb.getSheetByName(sheet).getConditionalFormattingRules().length,
    dvType: (sheet, ref) => wb.getSheetByName(sheet).getRange(ref).getDataValidation()?.getCriteriaType() ?? null,
    activeSheet: () => wb.getActiveSheet().getSheetName(),
    selection: () => wb.getActiveSheet().getSelection()?.getActiveRange()?.getA1Notation() ?? null,
    goTo(sheet, ref) {
      const ws = wb.getSheetByName(sheet);
      if (!ws) return;
      wb.setActiveSheet(ws);
      if (ref) ws.getRange(ref).activate();
    },
    async compareWithSaved() {
      await imagesDone;
      const snap = await waitForCalculation();
      if (!snap) return [];
      refreshGeneralPatterns();
      const out = [];
      for (const sh of model.sheets) {
        const id = sheetIdByName(sh.name);
        for (const c of sh.cells) {
          // Ô không có số đã lưu (file do app ghi, chưa qua Excel) thì không có gì để so.
          if (!c.f || c.excel_v === null || c.excel_v === undefined) continue;
          const computed = rawValue(snap, id, c.r, c.c);
          if (!sameValue(c.excel_v, computed)) out.push({ sheet: sh.name, cell: a1(c.r, c.c), saved: c.excel_v ?? null, computed });
        }
      }
      return out;
    },
    destroy() {
      disposed = true;
      clearTimeout(refreshTimer);
      clearTimeout(changeTimer);
      calcListener.dispose();
      calcStartListener.dispose();
      calcEndListener.dispose();
      for (const wake of [...calcWaiters]) wake();
      endUndoBatch();
      univer.dispose();
    },
  };
}
