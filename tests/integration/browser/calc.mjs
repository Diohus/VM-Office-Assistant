// T-DEV-08: sau khi sửa, ô công thức không được hiện số cũ mà không báo; chờ bộ tính xong trước khi dùng số; bộ tính không xong thì có hạn chờ.
import { chromium } from 'playwright-core';
import fs from 'node:fs';
import path from 'node:path';

const BASE = process.env.VM_BASE || 'http://127.0.0.1:8089';
const ROOT = path.resolve('..', '..', '..');
const OUT = path.resolve('..', 'out');
const sleep = (ms) => new Promise((r) => setTimeout(r, ms));
const res = { page_errors: [] };

const browser = await chromium.launch({ channel: 'msedge', headless: true });
const ctx = await browser.newContext({ viewport: { width: 1366, height: 768 }, acceptDownloads: true });
const pg = await ctx.newPage();
pg.on('pageerror', (e) => res.page_errors.push(String(e.message).slice(0, 200)));

async function open(file) {
  await pg.goto(BASE);
  await pg.waitForSelector('.welcome');
  await pg.evaluate(() => { window.vmGridReady = null; });
  await pg.setInputFiles('[data-testid=file-input]', file);
  await pg.waitForFunction(() => window.vmGridReady, null, { timeout: 300000 });
  await pg.evaluate(() => window.vmGrid.settled());
}
const val = (sheet, ref) => pg.evaluate(([s, r]) => window.vmGrid.value(s, r), [sheet, ref]);
async function typeEnter(sheet, ref, text) {
  await pg.evaluate(([s, r]) => window.vmGrid.goTo(s, r), [sheet, ref]);
  await sleep(150);
  await pg.keyboard.type(text);
  await pg.keyboard.press('Enter');
}
const notice = () => pg.$eval('[data-notice=open-error] .notice-text', (e) => e.firstChild.textContent).catch(() => null);

// 1. Bảng lớn (bộ tính mất vài giây): ngay sau Enter, settled() phải chờ tới khi ô công thức có số mới. Có dòng "Đang tính lại".
await open(path.join(ROOT, 'fixtures', 'generated', 'GIA-GAN_200000_o.xlsx'));
res.big = { b3_before: await val('Lon', 'B3'), j3_before: await val('Lon', 'J3') };
await typeEnter('Lon', 'B3', '1000');
// Theo dõi song song: lúc B3 đã mới mà J3 còn cũ (số cũ trên màn hình) thì có dòng "Đang tính lại" chưa.
// Độ trễ thay đổi theo lượt (đo 10-11/11: 0,3 tới 4 giây), nên đo khoảng số cũ hiện mà không có dòng báo, không đo một thời điểm.
Object.assign(res.big, await pg.evaluate(async ([oldJ3]) => {
  const t0 = performance.now();
  let settledResult = null;
  const waiting = window.vmGrid.settled().then((r) => { settledResult = r; });
  let staleMs = 0; let unannouncedMs = 0; let maxUnannounced = 0; let shownText = null; let last = performance.now();
  while (settledResult === null || performance.now() - t0 < 300) {
    const now = performance.now();
    const stale = window.vmGrid.value('Lon', 'B3') === 1000 && window.vmGrid.value('Lon', 'J3') === oldJ3;
    const banner = document.querySelector('[data-testid=calc-status]')?.textContent ?? null;
    if (banner) shownText = banner;
    if (stale) staleMs += now - last;
    if (stale && !banner) { unannouncedMs += now - last; maxUnannounced = Math.max(maxUnannounced, unannouncedMs); } else unannouncedMs = 0;
    last = now;
    if (now - t0 > 60000) break;
    await new Promise((r) => setTimeout(r, 20));
  }
  await waiting;
  return { settled: settledResult, settled_ms: Math.round(performance.now() - t0), stale_ms: Math.round(staleMs), max_unannounced_ms: Math.round(maxUnannounced), status_text: shownText };
}, [res.big.j3_before]));
res.big.j3_after_settled = await val('Lon', 'J3');
await sleep(400);
res.big.status_after = await pg.$('[data-testid=calc-status]') !== null;

// 2. Bộ tính không bao giờ báo xong: Tải PDF phải dừng sau hạn chờ, báo đúng lý do, app không treo.
await open(path.join(ROOT, 'fixtures', 'templates', 'GIA-QC-01_da_dien.xlsx'));
await pg.evaluate(() => { window.vmGrid.testCalcTimeout(3000); window.vmGrid.testHoldCalculation(true); });
await sleep(500);
res.stuck = { status_text: await pg.$eval('[data-testid=calc-status]', (e) => e.textContent).catch(() => null) };
// settled() phải chờ tới hạn rồi báo quá hạn, không được coi là xong ngay.
res.stuck.settled = await pg.evaluate(async () => { const t = performance.now(); const r = await window.vmGrid.settled(); return { ...r, ms: Math.round(performance.now() - t) }; });
const t0 = Date.now();
let downloaded = false;
pg.once('download', () => { downloaded = true; });
await pg.click('.sidebar .side-btn:has-text("Xuất")');
await pg.click('.dialog-actions button:has-text("Tải PDF")');
await pg.waitForSelector('[data-notice=open-error]', { timeout: 30000 }).catch(() => { res.stuck.no_notice = true; });
res.stuck.ms = Date.now() - t0;
res.stuck.error = await notice();
res.stuck.downloaded = downloaded;
// App không treo: gõ được ô khác ngay sau đó.
if (!res.stuck.no_notice) await pg.click('[data-notice=open-error] button:has-text("Đóng")');
await typeEnter('KQ_QC_Thang', 'I12', '42');
await sleep(400);
res.stuck.i12 = await val('KQ_QC_Thang', 'I12');
await pg.evaluate(() => { window.vmGrid.testHoldCalculation(false); window.vmGrid.testCalcTimeout(null); });
await sleep(1500);
res.stuck.status_after_release = await pg.$('[data-testid=calc-status]') !== null;

await browser.close();
fs.writeFileSync(path.join(OUT, 'calc.json'), JSON.stringify(res, null, 1));
