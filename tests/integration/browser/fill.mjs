// T-DEV-09 mốc 2: giao diện lát 1 nối AI (Q133). Đối chiếu docs/product/acceptance_lat1.md. Chỉ bộ thử 8089.
import { chromium } from 'playwright-core';
import fs from 'node:fs';
import path from 'node:path';

const BASE = process.env.VM_BASE || 'http://127.0.0.1:8089';
const FIX = path.resolve('..', '..', '..', 'fixtures', 'templates');
const SRC = path.join(FIX, 'nguon_lat1');
const OUT = path.resolve('..', 'out');
const SHEET = 'KQ_QC_Thang';
const sleep = (ms) => new Promise((r) => setTimeout(r, ms));
const res = { page_errors: [], outbound: [] };

const browser = await chromium.launch({ channel: 'msedge', headless: true });
const ctx = await browser.newContext({ viewport: { width: 1600, height: 900 } });
const pg = await ctx.newPage();
pg.on('pageerror', (e) => res.page_errors.push(String(e.message).slice(0, 200)));
pg.on('request', (r) => { if (!r.url().startsWith(BASE) && !r.url().startsWith('data:') && !r.url().startsWith('blob:')) res.outbound.push(r.url().slice(0, 80)); });

const val = (ref, sheet = SHEET) => pg.evaluate(([s, r]) => window.vmGrid.value(s, r), [sheet, ref]);
const errorText = () => pg.$eval('[data-notice=open-error] .notice-text', (e) => e.firstChild.textContent).catch(() => null);
const noticeText = (id) => pg.$eval(`[data-notice=${id}] .notice-text`, (e) => e.firstChild.textContent).catch(() => null);
const applyLabel = () => pg.textContent('[data-testid=fill-apply]');
const state = () => pg.textContent('[data-testid=doc-state]');
async function ready() {
  await pg.waitForFunction(() => window.vmGridReady, null, { timeout: 90000 });
  await pg.evaluate(() => window.vmGrid.settled());
  await sleep(300);
}
async function open(file) {
  await pg.goto(BASE);
  await pg.waitForSelector('.welcome');
  await pg.evaluate(() => { window.vmGridReady = null; });
  await pg.setInputFiles('[data-testid=file-input]', file);
  await ready();
}
async function chooseSource(file) {
  const [chooser] = await Promise.all([pg.waitForEvent('filechooser'), pg.click('[data-testid=fill-button]')]);
  await chooser.setFiles(file);
  await pg.waitForSelector('[data-dialog="Điền từ file"]');
}
async function startFill(file) {
  await chooseSource(file);
  await pg.click('.dialog-actions button:has-text("Điền")');
  await pg.waitForSelector('[data-testid=fill-panel], [data-notice=open-error]', { timeout: 120000 });
}
async function nameIfAsked() {
  await sleep(150);
  if (await pg.$('[data-testid=name-input]')) {
    await pg.fill('[data-testid=name-input]', 'Người Giả');
    await pg.click('.dialog-actions button:has-text("Tiếp tục")');
  }
}
async function applyNow() {
  await pg.evaluate(() => { window.vmGridReady = null; });
  await pg.click('[data-testid=fill-apply]');
  await nameIfAsked();
  await pg.waitForFunction(() => window.vmGridReady || document.querySelector('[data-notice=fill-again]') || document.querySelector('[data-notice=open-error]'), null, { timeout: 120000 });
  await sleep(500);
}
const rowButton = (i, label) => pg.click(`[data-testid=fill-rows] .fill-row >> nth=${i} >> button:has-text("${label}")`);

// 1. Tài liệu không phải QC mẫu: báo ngay, không mở hộp chọn file.
await open(path.join(FIX, 'GIA-QC-01_da_dien.xlsx'));
await pg.click('[data-testid=fill-button]');
await sleep(300);
res.not_qc = { error: await errorText(), dialog: await pg.$('[data-dialog="Điền từ file"]') !== null };

// 6.3: có sửa chưa lưu rồi bấm Điền từ file: hỏi như M2. Chọn Ở lại để không mở hộp chọn file.
await open(path.join(FIX, 'GIA-QC-01_mau.xlsx'));
await pg.evaluate(([s]) => window.vmGrid.goTo(s, 'I8'), [SHEET]);
await sleep(150);
await pg.keyboard.type('ghi chu');
await pg.keyboard.press('Enter');
await sleep(300);
await pg.click('[data-testid=fill-button]');
await sleep(300);
res.unsaved_asks = await pg.$('[data-dialog="Có sửa chưa lưu"]') !== null;
await pg.click('.dialog-actions button:has-text("Ở lại")');
await sleep(200);

// 1.1, 1.2, 1.5: hộp hỏi lại nêu cả hai tên file, có câu nhắc file giả; Hủy thì không có gì.
await open(path.join(FIX, 'GIA-QC-01_mau.xlsx'));
const docId = await pg.evaluate(() => window.vmGridReady.split(':')[0]);
await chooseSource(path.join(FIX, 'GIA-NGUON_2026-09.xlsx'));
res.confirm = await pg.$eval('[data-dialog="Điền từ file"]', (e) => e.textContent);
await pg.click('.dialog-actions button:has-text("Hủy")');
await sleep(300);
res.cancel = { panel: await pg.$('[data-testid=fill-panel]') !== null, state: await state(), open_batches: await pg.evaluate(async (id) => (await fetch(`/api/documents/${id}/fill`)).json(), docId) };

// 2.1, 3.1, 3.3, 4.1: X01, khung duyệt; lưới không đổi khi chưa Ghi.
await startFill(path.join(FIX, 'GIA-NGUON_2026-09.xlsx'));
res.x01 = {
  header: await pg.textContent('[data-testid=fill-header]'),
  source: await pg.textContent('[data-testid=fill-source]'),
  counts: await pg.textContent('[data-testid=fill-counts]'),
  mapping: await pg.$$eval('[data-testid=fill-mapping] li', (l) => l.length),
  rows: await pg.$$eval('[data-testid=fill-rows] .fill-row', (l) => l.length),
  b8_before: await val('B8'),
  apply_label: await applyLabel(),
};
await pg.screenshot({ path: path.join(OUT, 'fill_panel.png') });
// 3.10: nhận 5 dòng, bỏ nhận 1, nhận lại.
for (let i = 0; i < 5; i++) await rowButton(i, 'Nhận dòng');
res.x01.after_5 = await applyLabel();
await rowButton(0, 'Nhận dòng');
res.x01.after_unaccept = await applyLabel();
await rowButton(0, 'Nhận dòng');
// 3.4, 3.5, 3.5a: phím. Enter trong lưới không nhận; Ctrl+A, Ctrl+Enter, Shift+Enter không nhận; giữ Enter chỉ nhận 1 dòng.
await pg.evaluate(([s]) => window.vmGrid.goTo(s, 'B20'), [SHEET]);
await pg.click('[data-testid=grid]', { position: { x: 400, y: 400 } }); // bàn phím ở lưới
await sleep(150);
await pg.keyboard.press('Enter');
await sleep(200);
res.keys = { grid_enter: await applyLabel() };
await pg.focus('[data-testid=fill-rows]');
for (let i = 0; i < 5; i++) await pg.keyboard.press('ArrowDown');
await pg.keyboard.press('Control+A');
await pg.keyboard.press('Control+Enter');
await pg.keyboard.press('Shift+Enter');
res.keys.modifiers = await applyLabel();
await pg.keyboard.down('Enter');
for (let i = 0; i < 20; i++) await pg.keyboard.down('Enter'); // phím lặp khi giữ Enter
await pg.keyboard.up('Enter');
await sleep(200);
res.keys.hold_enter = await applyLabel();
res.x01.b8_still = await val('B8');
await applyNow();
res.x01.after_apply = { state: await state(), b8: await val('B8'), notice: await noticeText('fill-done'), panel: await pg.$('[data-testid=fill-panel]') !== null };
res.x01.versions = await pg.evaluate(async (id) => (await fetch(`/api/documents/${id}/versions`)).json(), docId);

// 5.3, D10: lô X01 còn chờ; tạo lô khác chạm cùng ô thì báo.
await chooseSource(path.join(SRC, 'X07_so_dang_chu_dau_phay.xlsx'));
await pg.click('.dialog-actions button:has-text("Điền")');
await pg.waitForSelector('[data-notice=open-error]', { timeout: 120000 });
res.d10 = await errorText();
await pg.click('[data-notice=open-error] button:has-text("Đóng")').catch(() => {});

// 3.8: mở lại lần điền còn chờ từ thông báo.
await pg.goto(BASE);
await pg.waitForSelector('.welcome');
await pg.evaluate(() => { window.vmGridReady = null; });
await pg.click(`[data-testid=recent] button[data-document="${docId}"]`);
await ready();
await sleep(800);
res.reopen = { notice: await noticeText('fill-open') };
await pg.click('[data-notice=fill-open] button:has-text("Xem vị trí")');
await pg.click('[data-notice=fill-open] .place-list button');
await pg.waitForSelector('[data-testid=fill-panel]');
res.reopen.counts = await pg.textContent('[data-testid=fill-counts]');

// D4 trên X04 và X10, ô ghi đè (D6): tài liệu mới.
await open(path.join(FIX, 'GIA-QC-01_mau.xlsx'));
await startFill(path.join(SRC, 'X04_mot_ngay_hai_lan_chay.xlsx'));
const red = await pg.$eval('.fill-cell[data-color=red]', (e) => ({ id: e.dataset.change, text: e.textContent }));
res.x04 = { red_text: red.text };
await pg.$eval(`.fill-cell[data-change="${red.id}"]`, (e) => e.scrollIntoView({ block: 'center' })); // ảnh phải có ô đỏ (PM b)
await sleep(200);
res.x04.red_in_view = await pg.$eval(`.fill-cell[data-change="${red.id}"]`, (e) => { const r = e.getBoundingClientRect(); return r.top >= 0 && r.bottom <= window.innerHeight; });
await pg.screenshot({ path: path.join(OUT, 'fill_o_do.png') });
await pg.click(`.fill-cell[data-change="${red.id}"] button:has-text("Chọn lần đọc 2")`);
res.x04.label = await applyLabel();
const redA1 = (red.text.match(/\(([A-Z]+\d+)\)/) || [])[1];
await applyNow();
res.x04.after = { value: await val(redA1), a1: redA1, versions: await pg.evaluate(async () => { const id = window.vmGridReady.split(':')[0]; return (await fetch(`/api/documents/${id}/versions`)).json(); }) };

await open(path.join(FIX, 'GIA-QC-01_mau.xlsx'));
await startFill(path.join(SRC, 'X10_so_vo_ly.xlsx'));
const reds = await pg.$$eval('.fill-cell[data-color=red]', (l) => l.map((e) => e.dataset.change));
res.x10 = { reds: reds.length };
await pg.fill(`.fill-cell[data-change="${reds[0]}"] input`, '1,234');
await pg.click(`.fill-cell[data-change="${reds[0]}"] button:has-text("Nhận số gõ")`);
res.x10.ambiguous = await pg.$eval(`.fill-cell[data-change="${reds[0]}"] .fill-problem`, (e) => e.textContent).catch(() => null);
res.x10.label_after_ambiguous = await applyLabel();
await pg.fill(`.fill-cell[data-change="${reds[0]}"] input`, '5.5');
await pg.click(`.fill-cell[data-change="${reds[0]}"] button:has-text("Nhận số gõ")`);
await pg.click(`.fill-cell[data-change="${reds[1]}"] button:has-text("Để trống")`);
res.x10.label = await applyLabel();
await applyNow();
res.x10.after = { notice: await noticeText('fill-done'), versions: await pg.evaluate(async () => { const id = window.vmGridReady.split(':')[0]; return (await fetch(`/api/documents/${id}/versions`)).json(); }) };

// D6, 5.1, 5.2: ô đã có số nằm riêng, không nhận theo dòng.
await open(path.join(FIX, 'GIA-QC-01_mau.xlsx'));
await pg.evaluate(([s]) => window.vmGrid.goTo(s, 'B8'), [SHEET]);
await sleep(150);
await pg.keyboard.type('9.99');
await pg.keyboard.press('Enter');
await sleep(300);
await pg.evaluate(() => { window.vmGridReady = null; });
await pg.click('.sidebar .side-btn:has-text("Lưu")');
await nameIfAsked();
await ready();
await startFill(path.join(FIX, 'GIA-NGUON_2026-09.xlsx'));
res.d6 = { section: await pg.$eval('[data-testid=fill-overwrites]', (e) => e.textContent).catch(() => null) };
await rowButton(0, 'Nhận dòng');
res.d6.label_after_row = await applyLabel();
res.d6.b8_choice = await pg.$eval('[data-testid=fill-overwrites] li button.on', (e) => e.textContent).catch(() => null);
// 6.3 lúc Ghi: có sửa chưa lưu thì hỏi trước, không ghi.
await pg.evaluate(([s]) => window.vmGrid.goTo(s, 'I10'), [SHEET]);
await pg.click('[data-testid=grid]', { position: { x: 400, y: 400 } });
await pg.evaluate(([s]) => window.vmGrid.goTo(s, 'I10'), [SHEET]);
await sleep(150);
await pg.keyboard.type('ghi chu');
await pg.keyboard.press('Enter');
await sleep(300);
await pg.click('[data-testid=fill-apply]');
await sleep(300);
res.d6.apply_unsaved_asks = await pg.$('[data-dialog="Có sửa chưa lưu"]') !== null;
await pg.click('.dialog-actions button:has-text("Ở lại")');
res.d6.state_after = await state();

const docNow = () => pg.evaluate(() => window.vmGridReady.split(':')[0]);
const api = (url) => pg.evaluate(async (u) => (await fetch(u)).json(), url);
const noticesText = () => pg.$eval('[data-testid=fill-notices]', (e) => e.textContent).catch(() => '');

// N8: lô 0 ô (X13 sai tháng): khung nói không có ô nào, không có nút Ghi, không báo còn ô chờ.
await open(path.join(FIX, 'GIA-QC-01_mau.xlsx'));
await startFill(path.join(SRC, 'X13_sai_thang_10.xlsx'));
res.n8 = {
  empty: await pg.$eval('[data-testid=fill-empty]', (e) => e.textContent).catch(() => null),
  header: await pg.$('[data-testid=fill-header]') !== null,
  apply: await pg.$('[data-testid=fill-apply]') !== null,
  notices: await noticesText(),
  open: (await api(`/api/documents/${await docNow()}/fill`)).length,
};

// K2: X03 thiếu 3 ngày. K3: X17 cột giống người bệnh.
await open(path.join(FIX, 'GIA-QC-01_mau.xlsx'));
await startFill(path.join(SRC, 'X03_thieu_3_ngay.xlsx'));
res.k2 = { notices: await noticesText(), common: await pg.$eval('[data-testid=fill-common]', (e) => e.textContent).catch(() => null) };
await open(path.join(FIX, 'GIA-QC-01_mau.xlsx'));
await startFill(path.join(SRC, 'X17_cot_nguoi_benh_gia.xlsx'));
res.k3 = { notices: await noticesText() };

// N5 (5.5): Bỏ lần điền này. Hỏi lại có số ô còn chờ; Hủy giữ lô; Đồng ý: không bản mới, không câu nhắc, lô mới mở được ngay.
const docN5 = await docNow();
const pendingBefore = await pg.$$eval('[data-testid=fill-panel] [data-change]', (l) => l.length);
await pg.click('[data-testid=fill-drop]');
await pg.waitForSelector('[data-dialog="Bỏ lần điền này"]');
res.n5 = { ask: await pg.textContent('[data-dialog="Bỏ lần điền này"]'), pending_before: pendingBefore };
await pg.click('.dialog-actions button:has-text("Hủy")');
await sleep(200);
res.n5.after_cancel = { panel: await pg.$('[data-testid=fill-panel]') !== null, open: (await api(`/api/documents/${docN5}/fill`)).length };
await pg.click('[data-testid=fill-drop]');
await pg.click('.dialog-actions button:has-text("Bỏ lần điền")');
await pg.waitForSelector('[data-notice=fill-dropped]', { timeout: 30000 });
res.n5.after_drop = {
  panel: await pg.$('[data-testid=fill-panel]') !== null,
  notice: await noticeText('fill-dropped'),
  versions: (await api(`/api/documents/${docN5}/versions`)).length,
  open: (await api(`/api/documents/${docN5}/fill`)).length,
};
await pg.goto(BASE);
await pg.waitForSelector('.welcome');
await pg.evaluate(() => { window.vmGridReady = null; });
await pg.click(`[data-testid=recent] button[data-document="${docN5}"]`);
await ready();
await sleep(800);
res.n5.reopen_notice = await noticeText('fill-open');
await startFill(path.join(SRC, 'X17_cot_nguoi_benh_gia.xlsx'));
res.n5.refill = { panel: await pg.$('[data-testid=fill-panel]') !== null, error: await errorText() };

// PO/BA 11/11: khung của tài liệu A đang mở, mở tài liệu B (cùng tên, không tải lại trang) thì khung đóng, không Ghi được lô của A vào đâu.
await open(path.join(FIX, 'GIA-QC-01_mau.xlsx'));
const docA = await pg.evaluate(() => window.vmGridReady.split(':')[0]);
await startFill(path.join(FIX, 'GIA-NGUON_2026-09.xlsx'));
res.other_doc = { panel_before: await pg.$('[data-testid=fill-panel]') !== null };
await pg.evaluate(() => { window.vmGridReady = null; });
await pg.setInputFiles('[data-testid=file-input]', path.join(FIX, 'GIA-QC-01_mau.xlsx'));
await ready();
const docB = await pg.evaluate(() => window.vmGridReady.split(':')[0]);
res.other_doc.same_doc = docA === docB;
res.other_doc.panel_after = await pg.$('[data-testid=fill-panel]') !== null;
res.other_doc.apply_button = await pg.$('[data-testid=fill-apply]') !== null;
res.other_doc.versions_a = (await pg.evaluate(async (id) => (await fetch(`/api/documents/${id}/versions`)).json(), docA)).length;
res.other_doc.versions_b = (await pg.evaluate(async (id) => (await fetch(`/api/documents/${id}/versions`)).json(), docB)).length;
res.other_doc.open_a = (await pg.evaluate(async (id) => (await fetch(`/api/documents/${id}/fill`)).json(), docA)).length;

// 6.2: mất kết nối khi tạo lô: câu tiếng Việt, không có khung, biểu mẫu không đổi.
await open(path.join(FIX, 'GIA-QC-01_mau.xlsx'));
await pg.route('**/api/documents/*/fill', (route) => (route.request().method() === 'POST' ? route.abort() : route.continue()));
await startFill(path.join(FIX, 'GIA-NGUON_2026-09.xlsx'));
res.net = { error: await errorText(), panel: await pg.$('[data-testid=fill-panel]') !== null, state: await state() };
await pg.unroute('**/api/documents/*/fill');

await browser.close();
fs.writeFileSync(path.join(OUT, 'fill.json'), JSON.stringify(res, null, 1));
