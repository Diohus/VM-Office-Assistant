// Điều khiển app thật trên Edge có sẵn, không tải trình duyệt. Ghi kết quả vào ../out/browser.json và ảnh chụp để PO/BA xem.
import { chromium } from 'playwright-core';
import { execFileSync } from 'node:child_process';
import crypto from 'node:crypto';
import fs from 'node:fs';
import path from 'node:path';

const BASE = process.env.VM_BASE || 'http://127.0.0.1:8089';
const API_CONTAINER = `${process.env.VM_PROJECT || 'vm-office-test'}-api-1`;
const ROOT = path.resolve('..', '..', '..');
const FIX = path.join(ROOT, 'fixtures', 'templates');
const OUT = path.resolve('..', 'out');
fs.mkdirSync(OUT, { recursive: true });
const sleep = (ms) => new Promise((r) => setTimeout(r, ms));
const base = new URL(BASE);

const res = { files: {}, outbound: [], page_errors: [], csp_violations: [], error_sequence: {} };
const browser = await chromium.launch({ channel: 'msedge', headless: true });
const ctx = await browser.newContext({ viewport: { width: 1366, height: 768 }, acceptDownloads: true });
const page = await ctx.newPage();
page.on('request', (r) => {
  const u = new URL(r.url());
  if (u.protocol.startsWith('http') && u.host !== base.host) res.outbound.push(r.url());
  if (u.protocol.startsWith('ws')) res.outbound.push(r.url());
});
page.on('pageerror', (e) => res.page_errors.push(String(e.stack || e).slice(0, 500)));
page.on('console', (m) => { if (m.type() === 'error' && /Content Security Policy|Refused to/i.test(m.text())) res.csp_violations.push(m.text().slice(0, 300)); });

const notices = () => page.$$eval('[data-notice]', (els) => els.map((e) => ({ id: e.dataset.notice, text: e.querySelector('.notice-text').firstChild.textContent })));

async function open(file) {
  await page.evaluate(() => { window.vmGridReady = null; });
  await page.setInputFiles('[data-testid=file-input]', file);
  await page.waitForFunction(() => window.vmGridReady, null, { timeout: 120000 });
  await sleep(400);
}

async function expectError(file) {
  await page.setInputFiles('[data-testid=file-input]', file);
  await page.waitForSelector('[data-notice=open-error]', { timeout: 60000 });
  const text = await page.$eval('[data-notice=open-error] .notice-text', (e) => e.textContent);
  await page.click('[data-notice=open-error] button:has-text("Đóng")');
  return text;
}

async function inspect(name, file) {
  const t0 = Date.now();
  await open(file);
  const seconds = (Date.now() - t0) / 1000;
  const r = await page.evaluate(async () => ({
    sheets: window.vmGrid.visibleSheets(),
    problems: window.vmGrid.problems,
    mismatches: await window.vmGrid.compareWithSaved(),
    stats: window.vmGrid.stats(),
  }));
  res.files[name] = { seconds, sheets: r.sheets, problems: r.problems, mismatches: r.mismatches, formulas: r.stats.formulas, notices: await notices() };
  await page.screenshot({ path: path.join(OUT, `${name}.png`) });
  return res.files[name];
}

await page.goto(BASE);
await page.waitForSelector('.welcome');
await page.screenshot({ path: path.join(OUT, 'man_chao.png') });

// Biểu mẫu trống: mục chưa hiện, nhận xét, xem vị trí, đóng và xem lại thông báo.
const qcPath = path.join(FIX, 'GIA-QC-01_mau.xlsx');
const qc = await inspect('GIA-QC-01_mau.xlsx', qcPath);
await page.click('[data-notice=unshown] button:has-text("Xem vị trí")');
await sleep(200);
await page.screenshot({ path: path.join(OUT, 'xem_vi_tri.png') });
qc.comment_text = await page.$eval('[data-notice=unshown] .place-list', (e) => e.textContent);
await page.click('[data-notice=unshown] .place-list button:has-text("B7")');
await sleep(500);
qc.jump_comment = await page.evaluate(() => ({ sheet: window.vmGrid.activeSheet(), cell: window.vmGrid.selection() }));
await page.click('[data-notice=unshown] .place-list button:has-text("Bieu_do")');
await sleep(500);
qc.jump_chart = await page.evaluate(() => ({ sheet: window.vmGrid.activeSheet(), cell: window.vmGrid.selection() }));
await page.click('[data-notice=unshown] > button:has-text("Đóng")');
qc.after_close = (await notices()).length;
await page.click('.reopen button');
qc.after_reopen = (await notices()).length;

// Bố cục.
res.layout = await page.evaluate(() => {
  const box = (s) => document.querySelector(s).getBoundingClientRect();
  return {
    page_scroll_x: document.documentElement.scrollWidth > document.documentElement.clientWidth,
    topbar_h: box('.topbar').height,
    sidebar_w: box('.sidebar').width,
    grid_share: box('.grid-host').width * box('.grid-host').height / (innerWidth * innerHeight),
    sidebar_labels: [...document.querySelectorAll('.side-btn')].map((b) => b.textContent.trim()),
    visible_text: document.body.innerText.slice(0, 1500),
  };
});

// Xuất: bấm nút thật, so hash file tải về.
const [download] = await Promise.all([page.waitForEvent('download'), page.click('.sidebar a.side-btn')]);
const dlPath = path.join(OUT, 'xuat.xlsx');
await download.saveAs(dlPath);
res.download = { name: download.suggestedFilename(), sha256: crypto.createHash('sha256').update(fs.readFileSync(dlPath)).digest('hex') };

// Bản xem không sửa được: gõ phím vào ô thật.
await page.evaluate(() => window.vmGrid.goTo('KQ_QC_Thang', 'B8'));
const before = await page.evaluate(() => window.vmGrid.value('KQ_QC_Thang', 'B8'));
const canvasBox = await page.$eval('.grid-host canvas', (c) => { const b = c.getBoundingClientRect(); return { x: b.x, y: b.y, w: b.width, h: b.height }; });
await page.mouse.dblclick(canvasBox.x + 200, canvasBox.y + 220);
await page.keyboard.type('123');
await page.keyboard.press('Enter');
await sleep(300);
await page.evaluate(() => window.vmGrid.goTo('KQ_QC_Thang', 'B8'));
await page.keyboard.type('123');
await page.keyboard.press('Enter');
await sleep(300);
res.read_only = { before, after: await page.evaluate(() => window.vmGrid.value('KQ_QC_Thang', 'B8')) };

// Biểu mẫu đã điền.
const dd = await inspect('GIA-QC-01_da_dien.xlsx', path.join(FIX, 'GIA-QC-01_da_dien.xlsx'));
Object.assign(dd, await page.evaluate(() => ({
  B39: window.vmGrid.value('KQ_QC_Thang', 'B39'),
  cf_rules: window.vmGrid.cfCount('KQ_QC_Thang'),
  dv_H8: window.vmGrid.dvType('KQ_QC_Thang', 'H8'),
})));

const kho = await inspect('GIA-KHO-02_tieu_hao.xlsx', path.join(FIX, 'GIA-KHO-02_tieu_hao.xlsx'));
await page.evaluate(() => window.vmGrid.goTo('HamTinh', 'A1'));
await sleep(600);
await page.screenshot({ path: path.join(OUT, 'kho_HamTinh.png') });
kho.display_narrow = await page.evaluate(() => Object.fromEntries(['C5', 'C6', 'C7', 'C15'].map((a) => [a, window.vmGrid.display('HamTinh', a)])));
kho.values_narrow = await page.evaluate(() => Object.fromEntries(['C5', 'C15'].map((a) => [a, window.vmGrid.value('HamTinh', a)])));
kho.display_all = await page.evaluate(() => window.vmGrid.allDisplay());
await inspect('GIA-NGUON_2026-09.xlsx', path.join(FIX, 'GIA-NGUON_2026-09.xlsx'));
await inspect('GIA-X18_ham_mang.xlsx', path.join(FIX, 'GIA-X18_ham_mang.xlsx'));
await inspect('x18_canary.xlsx', process.env.VM_CANARY_FILE);
const x18b = await inspect('GIA-X18b_lien_ket_ngoai.xlsx', path.join(FIX, 'GIA-X18b_lien_ket_ngoai.xlsx'));
x18b.display_general = await page.evaluate(() => [window.vmGrid.display('KQ_QC_Thang', 'I8'), window.vmGrid.display('KQ_QC_Thang', 'I9'), window.vmGrid.display('KQ_QC_Thang', 'B8')]);
x18b.sheet_list_button = await page.$$eval('[data-u-comp="sheet-bar-append-button"]', (els) => els.filter((e) => e.getClientRects().length > 0).length);
x18b.body_text = await page.evaluate(() => document.body.innerText);
x18b.btn_color = await page.evaluate(() => getComputedStyle(document.documentElement).getPropertyValue('--navy-2').trim());
await inspect('GIA-X19_khong_co_vba.xlsm', path.join(FIX, 'GIA-X19_khong_co_vba.xlsm'));
await inspect('GIA-X19_sheet_khoa.xlsx', path.join(FIX, 'GIA-X19_sheet_khoa.xlsx'));

// Lỗi liên tiếp, rồi mở file đúng trên cùng một trang.
fs.writeFileSync(path.join(OUT, 'ghi_chu.txt'), 'khong phai excel');
for (const name of ['GIA-X19_co_mat_khau.xlsx', 'GIA-X19_hong.xlsx', 'GIA-QC-01_da_dien_doi_cu.xls']) res.error_sequence[name] = await expectError(path.join(FIX, name));
res.error_sequence['ghi_chu.txt'] = await expectError(path.join(OUT, 'ghi_chu.txt'));
await open(qcPath);
res.recovered_doc = await page.textContent('[data-testid=doc-name]');

// Mất kết nối với máy chủ.
await page.route('**/api/documents**', (route) => route.abort());
res.network_down = await expectError(qcPath);
await page.unroute('**/api/documents**');

// Tắt API thật: web trả trang lỗi HTML, người dùng vẫn phải nhận được lời dễ hiểu, không treo.
execFileSync('docker', ['stop', '-t', '2', API_CONTAINER]);
try {
  res.backend_down = await expectError(qcPath);
} finally {
  execFileSync('docker', ['start', API_CONTAINER]);
}
for (let i = 0; i < 60; i++) {
  const ok = await fetch(`${BASE}/api/health`).then((r) => r.ok).catch(() => false);
  if (ok) break;
  await sleep(500);
}
await open(qcPath);
res.after_backend_back = await page.textContent('[data-testid=doc-name]');

await browser.close();
fs.writeFileSync(path.join(OUT, 'browser.json'), JSON.stringify(res, null, 1));
