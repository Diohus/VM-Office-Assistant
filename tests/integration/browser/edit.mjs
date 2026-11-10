// Luồng sửa ô trên Edge thật, gõ phím thật. Ghi kết quả vào ../out/edit.json.
import { chromium } from 'playwright-core';
import fs from 'node:fs';
import path from 'node:path';

const BASE = process.env.VM_BASE || 'http://127.0.0.1:8089';
const FIX = path.resolve('..', '..', '..', 'fixtures', 'templates');
const OUT = path.resolve('..', 'out');
const ROOT = path.resolve('..', '..', '..');
fs.mkdirSync(OUT, { recursive: true });
const sleep = (ms) => new Promise((r) => setTimeout(r, ms));
const base = new URL(BASE);
const QC = path.join(FIX, 'GIA-QC-01_da_dien.xlsx');
const SHEET = 'KQ_QC_Thang';

const res = { outbound: [], page_errors: [], csp_violations: [], typed: {} };
const browser = await chromium.launch({ channel: 'msedge', headless: true });
const ctx = await browser.newContext({ viewport: { width: 1366, height: 768 }, acceptDownloads: true });
const page = await ctx.newPage();
page.on('request', (r) => {
  const u = new URL(r.url());
  if (u.protocol.startsWith('http') && u.host !== base.host) res.outbound.push(r.url());
});
page.on('pageerror', (e) => res.page_errors.push(String(e.stack || e).slice(0, 500)));
page.on('console', (m) => { if (m.type() === 'error' && /Content Security Policy|Refused to/i.test(m.text())) res.csp_violations.push(m.text().slice(0, 300)); });

const value = (ref, sheet = SHEET) => page.evaluate(([s, r]) => window.vmGrid.value(s, r), [sheet, ref]);
const display = (ref, sheet = SHEET) => page.evaluate(([s, r]) => window.vmGrid.display(s, r), [sheet, ref]);
const noticeText = (id) => page.$eval(`[data-notice=${id}] .notice-text`, (e) => e.firstChild.textContent).catch(() => null);

async function waitReady() {
  await page.waitForFunction(() => window.vmGridReady, null, { timeout: 90000 });
  await page.evaluate(() => window.vmGrid.settled());
  await sleep(300);
}

async function open(file) {
  await page.evaluate(() => { window.vmGridReady = null; });
  await page.setInputFiles('[data-testid=file-input]', file);
  await waitReady();
}

// Gõ vào ô bằng bàn phím: chọn ô, gõ, Enter.
async function type(ref, text, sheet = SHEET) {
  await page.evaluate(([s, r]) => window.vmGrid.goTo(s, r), [sheet, ref]);
  await sleep(150);
  await page.keyboard.type(text);
  await page.keyboard.press('Enter');
  await sleep(400);
  await page.keyboard.press('Escape');
  await page.evaluate(() => window.vmGrid.settled());
  await sleep(200);
}

async function saveAs(name) {
  await page.evaluate(() => { window.vmGridReady = null; });
  await page.click('.sidebar .side-btn:has-text("Lưu")');
  if (name && await page.$('[data-testid=name-input]')) {
    await page.fill('[data-testid=name-input]', name);
    await page.click('.dialog-actions button:has-text("Tiếp tục")');
  }
  await waitReady();
}

await page.goto(BASE);
await page.waitForSelector('.welcome');
await open(QC);
await page.screenshot({ path: path.join(OUT, 'sua_truoc.png') });

// Sửa số: công thức tính lại ngay.
res.b39_before = await value('B39');
await type('B8', '5,5');
res.typed.comma = { value: await value('B8'), display: await display('B8'), notice: await noticeText('decimal') };
res.b39_after = await value('B39');
res.b42_after = await value('B42');
await type('B9', '5.51');
res.typed.dot = await value('B9');
res.typed.dot_display = await display('B9');
res.formula_b39 = await page.evaluate((s) => window.vmGrid.formula(s, 'B39'), SHEET);

// Nhập nhằng: không ghi, có hỏi lại.
const b10 = await value('B10');
await type('B10', '1.234');
res.typed.ambiguous_dot = { value: await value('B10'), before: b10, notice: await noticeText('ambiguous') };
await page.keyboard.press('Escape');
await sleep(200);
await type('B11', '1,234');
res.typed.ambiguous_comma = { notice: await noticeText('ambiguous') };
await page.keyboard.press('Escape');
await sleep(200);
await type('B12', '1.234,5');
res.typed.both = await value('B12');
// Q122: app dùng dấu chấm. Gõ dấu phẩy là số lẻ, có báo, không bao giờ thành số nghìn. 1,5,5 không ghi.
res.typed.swapped = {};
for (const [ref, text] of [['C8', '0,5'], ['C9', '12,5']]) {
  await type(ref, text);
  res.typed.swapped[text] = { value: await value(ref), notice: await noticeText('decimal') };
}
const c10 = await value('C10');
await type('C10', '1,5,5');
res.typed.unclear = { value: await value('C10'), before: c10, notice: await noticeText('ambiguous') };
await page.keyboard.press('Escape');
// Chờ ô thật sự thoát chế độ sửa: 1 trong 4 lượt ngày 10/11, chữ gõ ngay sau đó không vào ô kế tiếp.
await page.waitForFunction(() => !window.vmGrid.isEditing(), null, { timeout: 5000 });
await sleep(200);
await type('I8', 'abc');
res.typed.text = await value('I8');
await type('I9', '=SUM(B8:B10)');
res.typed.formula_value = await value('I9');
res.typed.formula_text = await page.evaluate((s) => window.vmGrid.formula(s, 'I9'), SHEET);
await type('I10', '1,5');
await page.waitForTimeout(300);
res.general_comma = await display('I10');
await type('I11', '=1/3');
res.general_formula_display = await display('I11');
await page.screenshot({ path: path.join(OUT, 'sua_sau.png') });

// Hoàn tác, làm lại.
const undoBefore = await value('I11');
await page.click('.editbar button:has-text("Hoàn tác")');
await sleep(500);
res.after_undo = await value('I11');
await page.click('.editbar button:has-text("Làm lại")');
await sleep(500);
res.after_redo = await value('I11');
res.undo_expected = undoBefore;

// Q86: mở file khác khi có sửa dở thì hỏi. "Ở lại" giữ nguyên.
res.dirty_count = await page.evaluate(() => window.vmGrid.collectEdits().length);
const [chooser] = await Promise.all([page.waitForEvent('filechooser'), page.click('.sidebar .side-btn:has-text("Mở")')]);
await chooser.setFiles(path.join(FIX, 'GIA-QC-01_mau.xlsx'));
await page.waitForSelector('[data-dialog="Có sửa chưa lưu"]');
res.unsaved_dialog = await page.textContent('.dialog-body');
await page.click('.dialog-actions button:has-text("Ở lại")');
res.after_stay_dirty = await page.evaluate(() => window.vmGrid.collectEdits().length);

// Q86: đóng tab thì trình duyệt cảnh báo.
let beforeUnload = 0;
page.on('dialog', async (d) => { if (d.type() === 'beforeunload') beforeUnload++; await d.dismiss(); });
await page.close({ runBeforeUnload: true }).catch(() => {});
await sleep(500);
res.beforeunload_dialogs = beforeUnload;
await browser.close();

// Phiên mới cho phần lưu và phiên bản.
const browser2 = await chromium.launch({ channel: 'msedge', headless: true });
const ctx2 = await browser2.newContext({ viewport: { width: 1366, height: 768 }, acceptDownloads: true });
const p2 = await ctx2.newPage();
p2.on('pageerror', (e) => res.page_errors.push(String(e.stack || e).slice(0, 500)));
await runVersions(p2);
await browser2.close();
fs.writeFileSync(path.join(OUT, 'edit.json'), JSON.stringify(res, null, 1));

async function runVersions(pg) {
  const val = (ref) => pg.evaluate(([s, r]) => window.vmGrid.value(s, r), [SHEET, ref]);
  const ready = async () => {
    await pg.waitForFunction(() => window.vmGridReady, null, { timeout: 90000 });
    await pg.evaluate(() => window.vmGrid.settled());
    await sleep(300);
  };
  const typeIn = async (ref, text) => {
    await pg.evaluate(([s, r]) => window.vmGrid.goTo(s, r), [SHEET, ref]);
    await sleep(150);
    await pg.keyboard.type(text);
    await pg.keyboard.press('Enter');
    await sleep(400);
  };
  await pg.goto(BASE);
  await pg.waitForSelector('.welcome');
  await pg.setInputFiles('[data-testid=file-input]', QC);
  await ready();
  res.v1_b39 = await val('B39');
  await typeIn('B8', '500');
  await pg.evaluate(() => { window.vmGridReady = null; });
  await pg.click('.sidebar .side-btn:has-text("Lưu")');
  await pg.fill('[data-testid=name-input]', 'Người Giả');
  await pg.click('.dialog-actions button:has-text("Tiếp tục")');
  await ready();
  res.v2_state = await pg.textContent('[data-testid=doc-state]');
  res.saved_notice = await pg.$eval('[data-notice=saved] .notice-text', (e) => e.firstChild.textContent);
  res.v2_b8 = await val('B8');
  res.v2_b39 = await val('B39');
  res.v2_mismatch_notice = await pg.$('[data-notice=mismatch]') !== null;
  res.v2_dirty = await pg.evaluate(() => window.vmGrid.collectEdits().length);

  // Xuất bản đã ghi: phải có cảnh báo Q48.
  await pg.click('.sidebar .side-btn:has-text("Xuất")');
  res.export_warning = await pg.textContent('.dialog-body');
  const [dl] = await Promise.all([pg.waitForEvent('download'), pg.click('.dialog-actions button:has-text("Tải file Excel")')]);
  await dl.saveAs(path.join(OUT, 'xuat_ban2.xlsx'));

  // Nút PDF thật trên giao diện.
  await pg.click('.sidebar .side-btn:has-text("Xuất")');
  const [pdf] = await Promise.all([pg.waitForEvent('download', { timeout: 120000 }), pg.click('.dialog-actions button:has-text("Tải PDF")')]);
  await pdf.saveAs(path.join(OUT, 'xuat_ban2.pdf'));
  res.pdf_name = pdf.suggestedFilename();
  res.pdf_error = await pg.$('[data-notice=open-error]') !== null;

  // Danh sách phiên bản, xem bản cũ chỉ đọc.
  await pg.click('.sidebar .side-btn:has-text("Phiên bản")');
  await pg.waitForSelector('[data-testid=versions] li');
  res.versions_text = await pg.textContent('[data-testid=versions]');
  await pg.evaluate(() => { window.vmGridReady = null; });
  await pg.click('[data-version="1"] button:has-text("Xem")');
  await ready();
  res.old_state = await pg.textContent('[data-testid=doc-state]');
  res.old_b8 = await val('B8');
  const before = await val('B9');
  await pg.evaluate(([s]) => window.vmGrid.goTo(s, 'B9'), [SHEET]);
  await pg.keyboard.type('123');
  await pg.keyboard.press('Enter');
  await sleep(400);
  res.old_readonly = { before, after: await val('B9') };
  res.old_has_save = await pg.$('.sidebar .side-btn:has-text("Lưu")') !== null;
  await pg.screenshot({ path: path.join(OUT, 'xem_ban_cu.png') });

  // Khôi phục tạo bản mới.
  await pg.evaluate(() => { window.vmGridReady = null; });
  await pg.click('.editbar button:has-text("Về bản mới nhất")');
  await ready();
  await pg.waitForSelector('[data-version="1"] button:has-text("Khôi phục")');
  await pg.evaluate(() => { window.vmGridReady = null; });
  await pg.click('[data-version="1"] button:has-text("Khôi phục")');
  res.restore_dialog = await pg.textContent('.dialog-body');
  await pg.click('.dialog-actions button:has-text("Khôi phục")');
  await ready();
  res.v3_state = await pg.textContent('[data-testid=doc-state]');
  res.v3_b8 = await val('B8');
  res.restored_notice = await pg.$eval('[data-notice=restored] .notice-text', (e) => e.firstChild.textContent);

  // Tải lại bản đã xuất, chưa qua Excel: không báo lệch sai, lưới tự tính.
  await pg.evaluate(() => { window.vmGridReady = null; });
  await pg.setInputFiles('[data-testid=file-input]', path.join(OUT, 'xuat_ban2.xlsx'));
  await ready();
  res.reupload_mismatch = await pg.$('[data-notice=mismatch]') !== null;
  res.reupload_b39 = await val('B39');
  res.reupload_display = await pg.evaluate((s) => window.vmGrid.display(s, 'B39'), SHEET);

  // Số rất lớn, số rất nhỏ ở cột rất hẹp: hiện như Excel, không phải dấu thăng.
  await pg.evaluate(() => { window.vmGridReady = null; });
  await pg.setInputFiles('[data-testid=file-input]', path.join(ROOT, 'fixtures', 'generated', 'GIA-SO_LON_cot_hep.xlsx'));
  await ready();
  res.narrow = await pg.evaluate(() => ['A1', 'A2', 'A3', 'A4'].map((a) => window.vmGrid.display('So', a)));
  await pg.screenshot({ path: path.join(OUT, 'so_lon_cot_hep.png') });

  // Ô khóa: không vào được chế độ sửa, có lý do.
  await pg.evaluate(() => { window.vmGridReady = null; });
  await pg.setInputFiles('[data-testid=file-input]', path.join(FIX, 'GIA-X19_sheet_khoa.xlsx'));
  await ready();
  const lockedBefore = await val('B8');
  await pg.evaluate(([s]) => window.vmGrid.goTo(s, 'B8'), [SHEET]);
  await pg.keyboard.type('9');
  await sleep(300);
  await pg.keyboard.press('Enter');
  await sleep(300);
  res.locked = { before: lockedBefore, after: await val('B8'), notice: await pg.$eval('[data-notice=locked] .notice-text', (e) => e.firstChild.textContent).catch(() => null) };
}
