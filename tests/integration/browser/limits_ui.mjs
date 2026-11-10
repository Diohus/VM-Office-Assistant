// T-DEV-05: ca giới hạn trên giao diện, chỉ bộ 8089. Không nằm trong run.sh. Chạy: node limits_ui.mjs
// GH-02: cùng file ở 2 tab, sửa cả hai rồi lưu. GH-06: sửa hơn 5.000 ô rồi lưu.
import { chromium } from 'playwright-core';
import fs from 'node:fs';
import path from 'node:path';

const BASE = process.env.VM_BASE || 'http://127.0.0.1:8089';
if (BASE.includes('8088')) throw new Error('Chỉ chạy trên bộ thử 8089');
const FIX = path.resolve('..', '..', '..', 'fixtures', 'templates');
const OUT = path.resolve('..', 'out');
const SHEET = 'KQ_QC_Thang';
const sleep = (ms) => new Promise((r) => setTimeout(r, ms));
const res = {};

const browser = await chromium.launch({ channel: 'msedge', headless: true });
const ctx = await browser.newContext({ viewport: { width: 1366, height: 768 } });

async function ready(pg) {
  await pg.waitForFunction(() => window.vmGridReady, null, { timeout: 90000 });
  await pg.evaluate(() => window.vmGrid.settled());
  await sleep(300);
}
async function type(pg, ref, text) {
  await pg.evaluate(([s, r]) => window.vmGrid.goTo(s, r), [SHEET, ref]);
  await sleep(150);
  await pg.keyboard.type(text);
  await pg.keyboard.press('Enter');
  await sleep(400);
  await pg.evaluate(() => window.vmGrid.settled());
}
async function save(pg) {
  await pg.evaluate(() => { window.vmGridReady = null; });
  await pg.click('.sidebar .side-btn:has-text("Lưu")');
  await sleep(150);
  if (await pg.$('[data-testid=name-input]')) {
    await pg.fill('[data-testid=name-input]', 'Người Giả');
    await pg.click('.dialog-actions button:has-text("Tiếp tục")');
  }
  await sleep(2500);
}
const notices = (pg) => pg.$$eval('[data-notice]', (l) => l.map((e) => ({ id: e.dataset.notice, text: e.querySelector('.notice-text').firstChild.textContent })));
const state = (pg) => pg.textContent('[data-testid=doc-state]');

// GH-02
const a = await ctx.newPage();
await a.goto(BASE);
await a.waitForSelector('.welcome');
await a.evaluate(() => { window.vmGridReady = null; });
await a.setInputFiles('[data-testid=file-input]', path.join(FIX, 'GIA-QC-01_mau.xlsx'));
await ready(a);
const docId = await a.evaluate(() => new URL(location.href).searchParams.get('doc') || null);
const b = await ctx.newPage();
await b.goto(BASE);
await b.waitForSelector('[data-testid=recent] button');
await b.click('[data-testid=recent] li:first-child button');
await ready(b);
await type(a, 'B8', '6.1');
await save(a);
res.gh02_a_state = await state(a);
await type(b, 'B9', '6.2');
await save(b);
res.gh02_b = { state: await state(b), notices: await notices(b), b9_on_screen: await b.evaluate((s) => window.vmGrid.value(s, 'B9'), SHEET), dirty: await b.evaluate(() => window.vmGrid.collectEdits().length) };
await b.screenshot({ path: path.join(OUT, 'gh02_tab_b.png') });
res.gh02_doc = docId;
await a.close();
await b.close();

// GH-06: dán 84 dòng × 60 cột số (5.040 ô) rồi lưu.
const c = await ctx.newPage();
await c.goto(BASE);
await c.waitForSelector('.welcome');
await c.setInputFiles('[data-testid=file-input]', path.join(FIX, 'GIA-QC-01_mau.xlsx'));
await ready(c);
const rows = Array.from({ length: 84 }, (_, r) => Array.from({ length: 60 }, (_, k) => String(r * 100 + k)).join('\t')).join('\n');
await c.evaluate(([s]) => window.vmGrid.goTo(s, 'A60'), [SHEET]);
await sleep(200);
const t0 = Date.now();
await c.evaluate((text) => {
  const dt = new DataTransfer();
  dt.setData('text/plain', text);
  document.activeElement.dispatchEvent(new ClipboardEvent('paste', { clipboardData: dt, bubbles: true, cancelable: true }));
}, rows);
await sleep(1500);
await c.evaluate(() => window.vmGrid.settled());
res.gh06_paste_ms = Date.now() - t0;
res.gh06_dirty_before = await c.evaluate(() => window.vmGrid.collectEdits().length);
res.gh06_paste_notices = await notices(c);
await save(c);
res.gh06_after = { state: await state(c), notices: await notices(c), dirty: await c.evaluate(() => window.vmGrid.collectEdits().length) };
await c.screenshot({ path: path.join(OUT, 'gh06_luu_5000.png') });

// GH-06b: biểu mẫu giả đủ lớn (fixtures/generated/GIA-LON_300x25.xlsx), dán 2 lần 150 × 25 = 7.500 ô rồi lưu.
const d = await ctx.newPage();
await d.goto(BASE);
await d.waitForSelector('.welcome');
await d.setInputFiles('[data-testid=file-input]', path.resolve('..', '..', '..', 'fixtures', 'generated', 'GIA-LON_300x25.xlsx'));
await ready(d);
const block = Array.from({ length: 150 }, (_, r) => Array.from({ length: 25 }, (_, k) => String(r * 100 + k)).join(String.fromCharCode(9))).join(String.fromCharCode(10));
for (const start of ['A2', 'A152']) {
  await d.evaluate(([r]) => window.vmGrid.goTo('Lon', r), [start]);
  await sleep(200);
  await d.evaluate((text) => {
    const dt = new DataTransfer();
    dt.setData('text/plain', text);
    document.activeElement.dispatchEvent(new ClipboardEvent('paste', { clipboardData: dt, bubbles: true, cancelable: true }));
  }, block);
  await sleep(1500);
  await d.evaluate(() => window.vmGrid.settled());
}
res.gh06b_dirty_before = await d.evaluate(() => window.vmGrid.collectEdits().length);
await save(d);
res.gh06b_after = { state: await state(d), notices: await notices(d), dirty: await d.evaluate(() => window.vmGrid.collectEdits().length) };
await d.screenshot({ path: path.join(OUT, 'gh06b_luu_7500.png') });

await browser.close();
fs.writeFileSync(path.join(OUT, 'limits_ui.json'), JSON.stringify(res, null, 1));
console.log(JSON.stringify(res, null, 1));
