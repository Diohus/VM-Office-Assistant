// Nơi duy nhất được import Univer. Chỉ dùng gói Apache-2.0, không có gói Pro nào.
import { createUniver, LocaleType, mergeLocales } from '@univerjs/presets';
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

const HA = { left: 1, center: 2, right: 3 };
const VA = { top: 1, center: 2, bottom: 3 };
const BORDER = {
  thin: 1, hair: 2, dotted: 3, dashed: 4, dashDot: 5, dashDotDot: 6, double: 7, medium: 8,
  mediumDashed: 9, mediumDashDot: 10, mediumDashDotDot: 11, slantDashDot: 12, thick: 13,
};
const TOLERANCE = 1e-6;
const EMU_PER_PX = 9525;
const CALC_DEADLINE_MS = 60000;
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

export function createGrid(container, model) {
  const { univer, univerAPI } = createUniver({
    locale: LocaleType.VI_VN,
    locales: { [LocaleType.VI_VN]: mergeLocales(coreVi, cfVi, dvVi, drawVi) },
    presets: [
      UniverSheetsCorePreset({
        container,
        header: false,
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
  const problems = applyRules(univerAPI, wb, model);
  const imagesDone = applyImages(wb, model).then((p) => {
    problems.push(...p);
    wb.setEditable(false);
  });

  const sheetIdByName = (name) => wb.getSheetByName(name)?.getSheetId();
  const rawValue = (snapshot, sheetId, r, c) => snapshot.sheets[sheetId]?.cellData?.[r]?.[c]?.v ?? null;

  // Bộ tính chạy nền và có thể chạy nhiều lượt. Tính xong khi mọi ô công thức có số đã lưu đều có kết quả,
  // hoặc số ô chưa có kết quả không giảm nữa (công thức thật sự ra rỗng thì phải dừng, không chờ mãi).
  async function waitForCalculation() {
    const pending = (snap) => {
      let n = 0;
      for (const sh of model.sheets) {
        const id = sheetIdByName(sh.name);
        for (const c of sh.cells) if (c.f && c.excel_v !== null && c.excel_v !== undefined && c.excel_v !== '' && rawValue(snap, id, c.r, c.c) === null) n++;
      }
      return n;
    };
    let last = Infinity;
    let stableSince = Date.now();
    const deadline = Date.now() + CALC_DEADLINE_MS;
    for (;;) {
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

  return {
    problems,
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
      const out = [];
      for (const sh of model.sheets) {
        const id = sheetIdByName(sh.name);
        for (const c of sh.cells) {
          if (!c.f) continue;
          const computed = rawValue(snap, id, c.r, c.c);
          if (!sameValue(c.excel_v, computed)) out.push({ sheet: sh.name, cell: a1(c.r, c.c), saved: c.excel_v ?? null, computed });
        }
      }
      return out;
    },
    destroy() {
      univer.dispose();
    },
  };
}
