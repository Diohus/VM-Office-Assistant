// Không mất sửa, không xuất sai số (Q86, Q92), hoàn tác, lệch số, dán từ Excel. Edge thật, gõ phím thật. Ghi ../out/guards.json.
import { chromium } from 'playwright-core';
import fs from 'node:fs';
import path from 'node:path';

const BASE = process.env.VM_BASE || 'http://127.0.0.1:8089';
const FIX = path.resolve('..', '..', '..', 'fixtures', 'templates');
const OUT = path.resolve('..', 'out');
const QC = path.join(FIX, 'GIA-QC-01_da_dien.xlsx');
const OTHER = path.join(FIX, 'GIA-QC-01_mau.xlsx');
const SHEET = 'KQ_QC_Thang';
const UNSAVED = '[data-dialog="Có sửa chưa lưu"]';
const sleep = (ms) => new Promise((r) => setTimeout(r, ms));

const res = { page_errors: [] };
const browser = await chromium.launch({ channel: 'msedge', headless: true });
const ctx = await browser.newContext({ viewport: { width: 1366, height: 768 }, acceptDownloads: true });
const pg = await ctx.newPage();
pg.on('pageerror', (e) => res.page_errors.push(String(e.stack || e).slice(0, 500)));

const val = (ref, sheet = SHEET) => pg.evaluate(([s, r]) => window.vmGrid.value(s, r), [sheet, ref]);
const dirty = () => pg.evaluate(() => window.vmGrid.collectEdits().length);
const state = () => pg.textContent('[data-testid=doc-state]');
const docName = () => pg.textContent('[data-testid=doc-name]');
const noticeText = (id) => pg.$eval(`[data-notice=${id}] .notice-text`, (e) => e.firstChild.textContent).catch(() => null);
const clearReady = () => pg.evaluate(() => { window.vmGridReady = null; });
const button = (text) => pg.click(`.dialog-actions button:has-text("${text}")`);
async function ready() {
  await pg.waitForFunction(() => window.vmGridReady, null, { timeout: 90000 });
  // R-1021: mỗi lần mở file, lưới phải thật sự được vẽ (có canvas lớn trong .grid-host), không chỉ có mô hình bảng.
  await pg.waitForFunction(() => [...document.querySelectorAll('.grid-host canvas')].some((c) => c.getBoundingClientRect().height > 100), null, { timeout: 15000 })
    .catch(() => { throw new Error('lưới trắng: .grid-host không có canvas sau khi mở file (R-1021)'); });
  await pg.evaluate(() => window.vmGrid.settled());
  await sleep(300);
}
async function open(file) {
  await clearReady();
  await pg.setInputFiles('[data-testid=file-input]', file);
  await ready();
}
async function typeNoEnter(ref, text, sheet = SHEET) {
  await pg.evaluate(([s, r]) => window.vmGrid.goTo(s, r), [sheet, ref]);
  await sleep(150);
  await pg.keyboard.type(text);
  await sleep(150);
}
async function type(ref, text, sheet = SHEET) {
  await typeNoEnter(ref, text, sheet);
  await pg.keyboard.press('Enter');
  await sleep(400);
  await pg.evaluate(() => window.vmGrid.settled());
  await sleep(150);
}
async function nameIfAsked() {
  await sleep(100);
  if (await pg.$('[data-testid=name-input]')) {
    await pg.fill('[data-testid=name-input]', 'Người Giả');
    await button('Tiếp tục');
  }
}
async function download(label, file) {
  const [d] = await Promise.all([pg.waitForEvent('download', { timeout: 120000 }), button(label)]);
  await d.saveAs(path.join(OUT, file));
  return d.suggestedFilename();
}
const sidebar = (text) => pg.click(`.sidebar .side-btn:has-text("${text}")`);

await pg.goto(BASE);
await pg.waitForSelector('.welcome');
await open(QC);
res.b8_original = await val('B8');
res.b9_original = await val('B9');

// 5.10: đang gõ dở, chưa Enter, mở file khác: số được tính là sửa chưa lưu.
await typeNoEnter('B20', '7,5');
res.midtype_editing = await pg.evaluate(() => window.vmGrid.isEditing());
await pg.setInputFiles('[data-testid=file-input]', OTHER);
await pg.waitForSelector(UNSAVED);
await button('Ở lại');
res.midtype_kept = await val('B20');
res.midtype_name = await docName();

// 5.10: đang gõ dở một số nhập nhằng: dừng lại và báo, không mở hộp, không mất.
await typeNoEnter('B21', '1.234');
await sidebar('Xuất');
await sleep(400);
res.midtype_ambiguous = { dialog: await pg.$('.dialog') !== null, notice: await noticeText('ambiguous'), editing: await pg.evaluate(() => window.vmGrid.isEditing()) };
await pg.keyboard.press('Escape');
await sleep(200);
res.midtype_ambiguous.after_esc = await val('B21');

// Mục 5b.1: có sửa chưa lưu, bấm Xuất, chọn Lưu: file ra đúng số đang thấy.
await type('B8', '500');
res.q92_grid_b39 = await val('B39');
await sidebar('Xuất');
await pg.waitForSelector(UNSAVED);
await clearReady();
await button('Lưu');
await nameIfAsked();
await pg.waitForSelector('[data-dialog="Xuất file"]', { timeout: 90000 });
res.q92_saved_state = await state();
fs.writeFileSync(path.join(OUT, 'q92_luu_grid.json'), JSON.stringify(await pg.evaluate(() => window.vmGrid.formulaValues())));
res.q92_saved_warning = await pg.textContent('.dialog-body');
res.q92_xlsx_name = await download('Tải file Excel', 'q92_luu.xlsx');
await sidebar('Xuất');
res.q92_pdf_name = await download('Tải PDF', 'q92_luu.pdf');

// Mục 5b.1: chọn Bỏ: lưới về bản đã lưu, không trắng, xuất được cả PDF.
await type('B9', '600');
await sidebar('Xuất');
await pg.waitForSelector(UNSAVED);
await clearReady();
await button('Bỏ thay đổi');
await pg.waitForSelector('[data-dialog="Xuất file"]', { timeout: 90000 });
res.q92_drop_b9 = await val('B9');
res.q92_drop_b8 = await val('B8');
res.q92_drop_dirty = await dirty();
res.q92_drop_pdf = await download('Tải PDF', 'q92_bo.pdf');

// Mục 5b.2: Xem bản cũ: Ở lại giữ sửa; Lưu thì sang bản cũ, về lại đúng bản mới nhất, không cũ số bản.
await sidebar('Phiên bản');
await pg.waitForSelector('[data-testid=versions] li');
await type('B10', '700');
await pg.click('[data-version="1"] button:has-text("Xem")');
await pg.waitForSelector(UNSAVED);
await button('Ở lại');
res.q92_stay_b10 = await val('B10');
await pg.click('[data-version="1"] button:has-text("Xem")');
await pg.waitForSelector(UNSAVED);
await clearReady();
await button('Lưu');
await nameIfAsked();
await pg.waitForFunction(() => document.querySelector('[data-testid=doc-state]')?.textContent.includes('chỉ đọc'), null, { timeout: 90000 });
await ready();
res.q92_old_state = await state();
res.q92_old_b8 = await val('B8');
await clearReady();
await pg.click('.editbar button:has-text("Về bản mới nhất")');
await ready();
res.q92_back_state = await state();
res.q92_back_b10 = await val('B10');

// Mục 5b.2: Khôi phục khi có sửa: Lưu rồi khôi phục, không báo bản đã cũ, mọi sửa nằm trong một bản.
await type('B12', '900');
await pg.click('[data-version="1"] button:has-text("Khôi phục")');
await button('Khôi phục');
await pg.waitForSelector(UNSAVED);
await clearReady();
await button('Lưu');
await nameIfAsked();
await pg.waitForFunction(() => /Đã tạo bản/.test(document.querySelector('[data-notice=restored]')?.textContent ?? ''), null, { timeout: 90000 });
await ready();
res.q92_restore_state = await state();
res.q92_restore_b8 = await val('B8');
res.q92_restore_error = await pg.$('[data-notice=open-error]') !== null;
res.q92_versions = await pg.$$eval('[data-testid=versions] li .v-meta:last-child', (l) => l.map((e) => e.textContent));

// Mục 5.9: lưu lỗi: không mở file kia, sửa còn nguyên, có báo lỗi.
await type('B13', '123');
await pg.route('**/api/documents/*/versions', (route) => (route.request().method() === 'POST' ? route.fulfill({ status: 500, contentType: 'application/json', body: '{"code":"X","message":"Giả lập lưu lỗi."}' }) : route.continue()));
const nameBefore = await docName();
await pg.setInputFiles('[data-testid=file-input]', OTHER);
await pg.waitForSelector(UNSAVED);
await button('Lưu');
await nameIfAsked();
await pg.waitForSelector('[data-notice=open-error]');
res.save_fail = { error: await noticeText('open-error'), same_doc: (await docName()) === nameBefore, b13: await val('B13'), dirty: await dirty() };
await pg.unroute('**/api/documents/*/versions');

// Q124: máy chủ trả lỗi không có câu của app thì không được báo là mất kết nối. 502 (API tắt sau máy chủ trung gian) mới là mất kết nối.
for (const [status, key] of [[400, 'server_400'], [502, 'server_502']]) {
  await pg.route('**/api/documents/*/versions', (route) => (route.request().method() === 'POST' ? route.fulfill({ status, contentType: 'text/html', body: '<html>loi</html>' }) : route.continue()));
  await sidebar('Lưu');
  await nameIfAsked();
  await pg.waitForSelector('[data-notice=open-error]');
  res[key] = { error: await noticeText('open-error'), dirty: await dirty() };
  await pg.unroute('**/api/documents/*/versions');
  await pg.click('[data-notice=open-error] button:has-text("Đóng")');
  await sleep(200);
}

// Mục 5.11: hoàn tác hết rồi mở file khác: không còn gì để hỏi, không mất sửa nào.
await pg.click('.editbar button:has-text("Hoàn tác")');
await sleep(500);
res.undo_all_dirty = await dirty();
res.undo_all_list = await pg.evaluate(() => window.vmGrid.collectEdits());
await clearReady();
await pg.setInputFiles('[data-testid=file-input]', OTHER);
await sleep(300);
res.undo_all_dialog = await pg.$(UNSAVED) !== null;
if (res.undo_all_dialog) await button('Bỏ thay đổi');
await ready();
res.undo_all_name = await docName();

// Mục 2.6: hoàn tác, làm lại 20 thao tác, đúng từng bước, kể cả số gõ bằng dấu phẩy và ô công thức.
await open(QC);
const steps = [];
const track = async () => ({ b39: await val('B39'), cells: await pg.evaluate((s) => Array.from({ length: 20 }, (_, i) => window.vmGrid.value(s, `B${8 + i}`)), SHEET) });
steps.push(await track());
for (let i = 0; i < 20; i++) {
  await type(`B${8 + i}`, i % 2 ? `${i + 100},5` : `${i + 100}`);
  steps.push(await track());
}
const walk = async (label, from, to) => {
  let wrong = 0;
  const dir = from > to ? -1 : 1;
  for (let k = from + dir; dir < 0 ? k >= to : k <= to; k += dir) {
    await pg.click(`.editbar button:has-text("${label}")`);
    await sleep(350);
    await pg.evaluate(() => window.vmGrid.settled());
    const now = await track();
    if (JSON.stringify(now) !== JSON.stringify(steps[k])) wrong++;
  }
  return wrong;
};
res.undo20 = { wrong_undo: await walk('Hoàn tác', 20, 0), wrong_redo: null };
res.undo20.wrong_redo = await walk('Làm lại', 0, 20);
res.undo20.end_b39 = await val('B39');
res.undo20.expected_end_b39 = steps[20].b39;

// Mục 2.10: số Excel đã lưu khác số lưới tính: báo đúng ô.
await pg.evaluate(() => { window.onbeforeunload = null; });
await pg.goto(BASE);
await pg.waitForSelector('.welcome');
await open(path.join(OUT, 'lech_so.xlsx'));
res.mismatch = { text: await noticeText('mismatch'), list: await pg.evaluate(() => window.vmGrid.compareWithSaved()) };

// Mục 3.10: dán từ Excel. Nội dung clipboard thật lấy từ Excel (excel_clipboard.json). Trình duyệt tự động không có clipboard máy, nên gửi sự kiện dán.
const clip = JSON.parse(fs.readFileSync('excel_clipboard.json', 'utf8'));
await open(QC);
const paste = async (ref, data, sheet = SHEET) => {
  await pg.evaluate(([s, r]) => window.vmGrid.goTo(s, r), [sheet, ref]);
  await sleep(150);
  await pg.evaluate(({ html, text }) => {
    const dt = new DataTransfer();
    if (html) dt.setData('text/html', html);
    dt.setData('text/plain', text);
    document.activeElement.dispatchEvent(new ClipboardEvent('paste', { clipboardData: dt, bubbles: true, cancelable: true }));
  }, data);
  await sleep(600);
  await pg.evaluate(() => window.vmGrid.settled());
};
await paste('I8', clip);
res.paste_excel = { values: await pg.evaluate((s) => Array.from({ length: 10 }, (_, i) => window.vmGrid.value(s, `I${8 + i}`)), SHEET), notice: await noticeText('pasted'), dirty: await dirty() };
await paste('J8', { text: '1,234,567\n1,234\n5.47\n' });
res.paste_text_evidence = { values: await pg.evaluate((s) => [8, 9, 10].map((r) => window.vmGrid.value(s, `J${r}`)), SHEET), notice: await noticeText('pasted') };
await paste('K8', { text: '1,234\n1.234\n' });
res.paste_text_none = { values: await pg.evaluate((s) => [8, 9].map((r) => window.vmGrid.value(s, `K${r}`)), SHEET), notice: await noticeText('pasted') };
await paste('L8', { text: '=WEBSERVICE("http://example.invalid")\n' });
res.paste_formula = { value: await val('L8'), formula: await pg.evaluate((s) => window.vmGrid.formula(s, 'L8'), SHEET) };
await pg.click('.editbar button:has-text("Hoàn tác")');
await sleep(300);

await pg.evaluate(() => { window.onbeforeunload = null; });
await open(path.join(FIX, 'GIA-X19_sheet_khoa.xlsx')).catch(() => {});
res.after_paste_dialog = await pg.$(UNSAVED) !== null;
if (res.after_paste_dialog) {
  await clearReady();
  await button('Bỏ thay đổi');
  await ready();
}
const lockedBefore = await val('B8');
await paste('B8', { text: '9\n' });
res.paste_locked = { before: lockedBefore, after: await val('B8'), notice: await noticeText('locked') };

// Mục 5b.3, 5b.4: Ctrl+S là Lưu, không mở hộp của trình duyệt; thanh trên báo có sửa chưa lưu.
await pg.goto(BASE);
await pg.waitForSelector('.welcome');
await open(QC);
await pg.evaluate(() => {
  const original = KeyboardEvent.prototype.preventDefault;
  KeyboardEvent.prototype.preventDefault = function patched() {
    if (this.key?.toLowerCase() === 's' && this.ctrlKey) window.__ctrlS = (window.__ctrlS || 0) + 1;
    return original.call(this);
  };
});
const dirtyShown = () => pg.$('[data-testid=dirty-state]').then((e) => e !== null);
res.bar_clean = await dirtyShown();
await type('B8', '6,5');
await sleep(300);
res.bar_dirty = await dirtyShown();
await clearReady();
await pg.keyboard.press('Control+s');
await nameIfAsked();
await ready();
res.ctrl_s = { state: await state(), prevented: await pg.evaluate(() => window.__ctrlS ?? 0), b8: await val('B8'), saved: await noticeText('saved') };
res.bar_after_save = await dirtyShown();
await type('B9', '6,6');
await sleep(300);
res.bar_dirty_again = { shown: await dirtyShown(), saved_notice: await noticeText('saved') };
await pg.click('.editbar button:has-text("Hoàn tác")');
await sleep(500);
res.bar_after_undo = await dirtyShown();
await typeNoEnter('B10', '6,7');
await clearReady();
await pg.keyboard.press('Control+s');
await nameIfAsked();
await ready();
res.ctrl_s_midtype = { state: await state(), b10: await val('B10') };
await pg.keyboard.press('Control+s');
await sleep(400);
res.ctrl_s_nothing = { state: await state(), notice: await noticeText('nothing') };

// Mục 7.3: lệch giả ở số lưới gửi đi. Cách này cũng dùng được khi nghiệm thu, xem m2_report.md.
await pg.evaluate(() => {
  const real = window.vmGrid.formulaValues;
  window.vmGrid.formulaValues = () => real().map((v) => (v.cell === 'B39' ? { ...v, value: 999 } : v));
});
await sidebar('Xuất');
res.export_text = await pg.$$eval('.dialog-body p', (l) => l.map((e) => e.textContent));
await button('Tải PDF');
await pg.waitForSelector('[data-notice=pdf-mismatch]', { timeout: 120000 });
res.pdf_fake_mismatch = await noticeText('pdf-mismatch');

// Mục 5b.5: đóng trình duyệt, mở lại: từ màn chào về đúng file, đúng bản mới nhất.
const savedId = await pg.evaluate(() => fetch('/api/documents').then((r) => r.json()).then((l) => l[0].document_id));
const ctx3 = await browser.newContext({ viewport: { width: 1366, height: 768 } });
const pg3 = await ctx3.newPage();
pg3.on('pageerror', (e) => res.page_errors.push(String(e.stack || e).slice(0, 500)));
await pg3.goto(BASE);
await pg3.waitForSelector('[data-testid=recent]');
res.back_row = await pg3.$eval(`[data-testid=recent] li:has([data-document="${savedId}"])`, (e) => e.textContent);
await pg3.screenshot({ path: path.join(OUT, 'man_chao_quay_lai.png') });
await pg3.click(`[data-document="${savedId}"]`);
await pg3.waitForFunction(() => window.vmGridReady, null, { timeout: 90000 });
await pg3.evaluate(() => window.vmGrid.settled());
res.back = { state: await pg3.textContent('[data-testid=doc-state]'), b8: await pg3.evaluate((s) => window.vmGrid.value(s, 'B8'), SHEET), b10: await pg3.evaluate((s) => window.vmGrid.value(s, 'B10'), SHEET) };
await pg3.evaluate(() => { window.vmGridReady = null; });
await pg3.setInputFiles('[data-testid=file-input]', QC);
await pg3.waitForFunction(() => window.vmGridReady, null, { timeout: 90000 });
await pg3.evaluate(() => window.vmGrid.settled());
res.reupload_original = { state: await pg3.textContent('[data-testid=doc-state]'), b8: await pg3.evaluate((s) => window.vmGrid.value(s, 'B8'), SHEET) };
const rows = await pg3.evaluate(() => fetch('/api/documents').then((r) => r.json()));
res.reupload_kept = rows.some((r) => r.document_id === savedId && r.version === 3);
await ctx3.close();

// Dòng chặn Q104: F2 rồi Enter không gõ gì không đổi số. Có gõ thì vẫn là sửa.
await pg.goto(BASE);
await pg.waitForSelector('.welcome');
await open(QC);
await type('I20', '3,14159265358979');
await type('I21', '1,2340');
await clearReady();
await sidebar('Lưu');
await nameIfAsked();
await ready();
const f2Enter = async (ref) => {
  await pg.evaluate(([s, r]) => window.vmGrid.goTo(s, r), [SHEET, ref]);
  await sleep(150);
  await pg.keyboard.press('F2');
  await sleep(200);
  if (ref === 'I20') await pg.screenshot({ path: path.join(OUT, 'f2_thanh_cong_thuc.png') });
  await pg.keyboard.press('Enter');
  await sleep(400);
};
res.f2 = { before: await val('I20'), ambiguous_before: await val('I21') };
await f2Enter('I20');
res.f2.after = await val('I20');
res.f2.dirty = await dirty();
await f2Enter('I21');
res.f2.ambiguous_after = await val('I21');
res.f2.ambiguous_notice = await noticeText('ambiguous');
res.f2.ambiguous_dirty = await dirty();
await pg.evaluate(([s]) => window.vmGrid.goTo(s, 'B9'), [SHEET]);
await sleep(200);
await pg.screenshot({ path: path.join(OUT, 'thanh_cong_thuc_dau_phay.png') });
await pg.keyboard.press('F2');
await sleep(200);
await pg.keyboard.type('12');
await pg.keyboard.press('Enter');
await sleep(400);
res.f2.typed_after = await val('B9');

// Q122: F2 trên ô có nhiều chữ số hơn ô đang sửa hiện (12 chữ số), gõ thêm. Lần Enter đầu chỉ hỏi, số cũ còn.
// Gõ thêm sau khi hỏi thì phải hỏi lại. Enter lần nữa, không đổi gì, mới ghi.
const editing = () => pg.evaluate(() => window.vmGrid.isEditing());
await pg.evaluate(([s]) => window.vmGrid.goTo(s, 'I20'), [SHEET]);
await sleep(150);
await pg.keyboard.press('F2');
await sleep(200);
await pg.keyboard.press('End');
await pg.keyboard.type('1');
await pg.keyboard.press('Enter');
await sleep(400);
res.hidden = { first: { value: await val('I20'), notice: await noticeText('hidden-digits'), editing: await editing() } };
await pg.screenshot({ path: path.join(OUT, 'f2_mat_chu_so.png') });
await pg.keyboard.type('2');
await pg.keyboard.press('Enter');
await sleep(400);
res.hidden.changed = { value: await val('I20'), notice: await noticeText('hidden-digits'), editing: await editing() };
await pg.keyboard.press('Enter');
await sleep(400);
await pg.evaluate(() => window.vmGrid.settled());
res.hidden.second = { value: await val('I20'), notice: await noticeText('hidden-digits'), editing: await editing() };
// Gõ thẳng vào ô (không F2) thì thay hẳn số cũ, không có gì bị cắt, không hỏi.
await type('I22', '2.71828182845905');
await type('I22', '9');
res.hidden.direct = { value: await val('I22'), notice: await noticeText('hidden-digits') };

// F2 rồi gõ ra một số có thể hiểu hai cách: hỏi lại, số cũ còn nguyên.
await pg.evaluate(([s]) => window.vmGrid.goTo(s, 'B11'), [SHEET]);
await sleep(150);
await pg.keyboard.press('F2');
await sleep(200);
await pg.keyboard.type('1');
await pg.keyboard.press('Enter');
await sleep(300);
res.f2.typed_ambiguous = { value: await val('B11'), notice: await noticeText('ambiguous'), editing: await pg.evaluate(() => window.vmGrid.isEditing()) };
await pg.keyboard.press('Escape');
await sleep(200);

// Chữ số lẫn lộn: không để thư viện đoán.
await typeNoEnter('B12', '5,5.1,2');
await pg.keyboard.press('Enter');
await sleep(300);
res.unclear = { value: await val('B12'), notice: await noticeText('ambiguous'), editing: await pg.evaluate(() => window.vmGrid.isEditing()) };
await pg.keyboard.press('Escape');
await sleep(200);

// Thông báo nhập nhằng tự đóng khi ô đã ghi được.
await typeNoEnter('B10', '1.234');
await pg.keyboard.press('Enter');
await sleep(300);
const firstAmbiguous = await noticeText('ambiguous');
await pg.keyboard.press('Escape');
await sleep(200);
await type('B10', '1234');
res.ambiguous_close = { first: firstAmbiguous, after_fix: await noticeText('ambiguous'), value: await val('B10') };

// Dòng chặn Q104: số hiện lệch số thật. Clipboard thật của Excel (excel_clipboard_lech.json), giá trị gốc không có trong đó.
const clipRounded = JSON.parse(fs.readFileSync('excel_clipboard_lech.json', 'utf8'));
await paste('L8', clipRounded);
res.paste_rounded = { values: await pg.evaluate((s) => Array.from({ length: 8 }, (_, i) => window.vmGrid.value(s, `L${8 + i}`)), SHEET), notice: await noticeText('pasted') };
// Bấm Xem vị trí làm tiêu điểm nằm ở nút, các lần dán sau sẽ không vào lưới. Đọc xong thì trả tiêu điểm về lưới như trước.
const placesOf = async (id) => {
  await pg.evaluate(() => { window.__focusBefore = document.activeElement; });
  await pg.click(`[data-notice=${id}] .notice-text button:has-text("Xem vị trí")`);
  const places = await pg.$$eval(`[data-notice=${id}] .place-list button`, (l) => l.map((e) => e.textContent));
  await pg.evaluate(() => window.__focusBefore?.focus());
  await sleep(100);
  return places;
};
res.paste_rounded.places = await placesOf('pasted');
// Q113: ô General cột hẹp, clipboard thật của Excel (excel_clipboard_general.json). Số gốc chỉ có trong phần chữ thường.
const clipGeneral = JSON.parse(fs.readFileSync('excel_clipboard_general.json', 'utf8'));
await paste('N8', clipGeneral);
res.paste_general = {
  values: await pg.evaluate((s) => Array.from({ length: 7 }, (_, i) => [window.vmGrid.value(s, `N${8 + i}`), window.vmGrid.value(s, `O${8 + i}`)]), SHEET),
  notice: await noticeText('pasted'),
  places: await placesOf('pasted'),
};
// Cùng clipboard nhưng mất phần chữ thường: mọi ô số đều phải báo.
await paste('N20', { html: clipGeneral.html, text: '' });
res.paste_general_no_plain = { notice: await noticeText('pasted'), places: await placesOf('pasted') };
// Q125: dán chữ thuần (Notepad) báo giống gõ tay.
await paste('K30', { text: ['5,5', '0,5', '12,75', '1,5,5', ''].join(String.fromCharCode(10)) });
res.paste_plain_swapped = { values: await pg.evaluate((s) => [30, 31, 32, 33].map((r) => window.vmGrid.value(s, `K${r}`)), SHEET), notice: await noticeText('pasted'), places: await placesOf('pasted') };
// Q125: bấm Đóng trên thông báo rồi gõ ngay: phím phải vào lưới, không mất phím nào (gõ 12 không được thành 2 hay trống).
await pg.evaluate(([s]) => window.vmGrid.goTo(s, 'K35'), [SHEET]);
await sleep(150);
await pg.click('[data-notice=pasted] button:has-text("Đóng")');
await pg.keyboard.type('12');
await pg.keyboard.press('Enter');
await sleep(400);
res.close_then_type = await val('K35');
await paste('L30', { html: '<html xmlns:x="urn:schemas-microsoft-com:office:excel"><body><table><tr><td align=right>9</td></tr><tr><td class=xl68 align=center>######</td></tr></table></body></html>', text: '9\n######\n' });
res.paste_narrow = { value: await val('L30'), notice: await noticeText('pasted') };

// Hộp xuất của file đời cũ chưa sửa.
await pg.goto(BASE);
await pg.waitForSelector('.welcome');
await open(path.join(FIX, 'GIA-QC-01_da_dien_doi_cu.xls'));
await sidebar('Xuất');
res.legacy_export = await pg.$$eval('.dialog-body p', (l) => l.map((e) => e.textContent));
await button('Hủy');

// Mục 6: biểu mẫu thứ hai, để tools/excel_check/q48_grid.py so số lưới với Excel.
await clearReady();
await open(path.join(FIX, 'GIA-KHO-02_tieu_hao.xlsx'));
const kho = await pg.evaluate(() => window.vmGrid.activeSheet());
res.kho_sheet = kho;
res.kho_c4_before = await val('C4', kho);
await type('C4', '77', kho);
await clearReady();
await sidebar('Lưu');
await nameIfAsked();
await ready();
res.kho_state = await state();
fs.writeFileSync(path.join(OUT, 'kho_grid.json'), JSON.stringify(await pg.evaluate(() => window.vmGrid.formulaValues())));
await sidebar('Xuất');
res.kho_name = await download('Tải file Excel', 'kho_xuat.xlsx');
await sidebar('Xuất');
await download('Tải PDF', 'kho_xuat.pdf');
res.kho_pdf_error = await noticeText('pdf-mismatch');

// Q124: quá 5000 ô sửa thì báo ngay ở trình duyệt, không gửi đi; các sửa còn trên màn hình.
await clearReady();
await pg.goto(BASE);
await pg.waitForSelector('.welcome');
await open(path.resolve('..', '..', '..', 'fixtures', 'generated', 'GIA-LON_300x25.xlsx'));
const bigBlock = Array.from({ length: 150 }, (_, r) => Array.from({ length: 25 }, (_, k) => String(r * 100 + k)).join(String.fromCharCode(9))).join(String.fromCharCode(10));
await paste('A2', { text: bigBlock }, 'Lon');
await paste('A152', { text: bigBlock }, 'Lon');
let savePosts = 0;
pg.on('request', (r) => { if (r.method() === 'POST' && r.url().includes('/versions')) savePosts++; });
await sidebar('Lưu');
await nameIfAsked();
await pg.waitForSelector('[data-notice=open-error]');
res.too_many = { error: await noticeText('open-error'), dirty: await dirty(), state: await state(), posts: savePosts };
// Làm đúng cách câu báo chỉ: Xuất, Bỏ thay đổi, Hủy. Phải bỏ hết sửa trong một bước, về bản đang lưu.
await sidebar('Xuất');
await pg.waitForSelector(UNSAVED);
await button('Bỏ thay đổi');
await pg.waitForSelector('.dialog-actions button:has-text("Hủy")');
await button('Hủy');
await sleep(300);
res.too_many_discard = { dirty: await dirty(), state: await state(), a2: await val('A2', 'Lon') };

await browser.close();
fs.writeFileSync(path.join(OUT, 'guards.json'), JSON.stringify(res, null, 1));
