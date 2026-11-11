// T-DEV-06b: nút Hoàn tác, Làm lại chỉ sáng khi còn bước để làm. Thư viện giữ tối đa 50 bước; sau khi Lưu không còn gì để Làm lại.
import { chromium } from 'playwright-core';
import fs from 'node:fs';
import path from 'node:path';

const BASE = process.env.VM_BASE || 'http://127.0.0.1:8089';
const FIX = path.resolve('..', '..', '..', 'fixtures', 'templates');
const OUT = path.resolve('..', 'out');
const SHEET = 'KQ_QC_Thang';
const sleep = (ms) => new Promise((r) => setTimeout(r, ms));
const res = { page_errors: [] };

const browser = await chromium.launch({ channel: 'msedge', headless: true });
const pg = await (await browser.newContext({ viewport: { width: 1366, height: 768 } })).newPage();
pg.on('pageerror', (e) => res.page_errors.push(String(e.message).slice(0, 200)));
const buttons = () => pg.evaluate(() => Object.fromEntries(['Hoàn tác', 'Làm lại'].map((t) => [t, ![...document.querySelectorAll('.editbar button')].find((b) => b.textContent === t)?.disabled])));
const dirty = () => pg.evaluate(() => window.vmGrid.collectEdits().length);
async function type(ref, text) {
  await pg.evaluate(([s, r]) => window.vmGrid.goTo(s, r), [SHEET, ref]);
  await sleep(100);
  await pg.keyboard.type(text);
  await pg.keyboard.press('Enter');
  await sleep(250);
}

await pg.goto(BASE);
await pg.waitForSelector('.welcome');
await pg.setInputFiles('[data-testid=file-input]', path.join(FIX, 'GIA-QC-01_mau.xlsx'));
await pg.waitForFunction(() => window.vmGridReady, null, { timeout: 90000 });
await pg.evaluate(() => window.vmGrid.settled());
res.open = await buttons();
await type('B8', '5.5');
res.after_type = await buttons();
await pg.click('.editbar button:has-text("Hoàn tác")');
await sleep(400);
res.after_undo = { ...(await buttons()), dirty: await dirty() };
await pg.click('.editbar button:has-text("Làm lại")');
await sleep(400);
res.after_redo = { ...(await buttons()), dirty: await dirty() };
// Lưu: lưới được tạo lại, không còn gì để Hoàn tác hay Làm lại.
await pg.evaluate(() => { window.vmGridReady = null; });
await pg.click('.sidebar .side-btn:has-text("Lưu")');
await sleep(150);
if (await pg.$('[data-testid=name-input]')) {
  await pg.fill('[data-testid=name-input]', 'Người Giả');
  await pg.click('.dialog-actions button:has-text("Tiếp tục")');
}
await pg.waitForFunction(() => window.vmGridReady, null, { timeout: 90000 });
await sleep(300);
res.after_save = { ...(await buttons()), state: await pg.textContent('[data-testid=doc-state]') };
// Quá 50 bước: gõ 55 ô, Hoàn tác tới khi hết. Nút phải tắt khi không còn bước, dù vẫn còn sửa.
for (let i = 0; i < 55; i++) await type(`${i < 28 ? 'C' : 'D'}${8 + (i % 28)}`, String(10 + i));
let clicks = 0;
while ((await buttons())['Hoàn tác'] && clicks < 80) {
  await pg.click('.editbar button:has-text("Hoàn tác")');
  await sleep(60);
  clicks++;
}
res.limit = { clicks, ...(await buttons()), dirty: await dirty() };

await browser.close();
fs.writeFileSync(path.join(OUT, 'undo.json'), JSON.stringify(res, null, 1));
