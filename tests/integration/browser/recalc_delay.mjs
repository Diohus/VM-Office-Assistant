// T-DEV-08: đo độ trễ tính lại sau khi sửa một ô, và lưới có hiện số cũ không. Chỉ bộ thử 8089. Không nằm trong run.sh.
// Chạy: node recalc_delay.mjs. Kết quả: ../out/recalc_delay.json
import { chromium } from 'playwright-core';
import fs from 'node:fs';
import path from 'node:path';

const BASE = process.env.VM_BASE || 'http://127.0.0.1:8089';
if (BASE.includes('8088')) throw new Error('Chỉ chạy trên bộ thử 8089');
const ROOT = path.resolve('..', '..', '..');
const OUT = path.resolve('..', 'out');
const sleep = (ms) => new Promise((r) => setTimeout(r, ms));
const CASES = [
  // [tên, file, sheet, ô sửa, giá trị, ô công thức theo dõi]
  ['QC tháng', path.join(ROOT, 'fixtures', 'templates', 'GIA-QC-01_da_dien.xlsx'), 'KQ_QC_Thang', 'B8', '50', ['B39', 'B42']],
  ['300 x 25 có công thức', path.join(ROOT, 'fixtures', 'generated', 'GIA-CONG_THUC_300x25.xlsx'), 'Bang', 'B3', '1000', ['Y3', 'B303', 'Y303']],
  ['18000 x 10 gần giới hạn', path.join(ROOT, 'fixtures', 'generated', 'GIA-GAN_200000_o.xlsx'), 'Lon', 'B3', '1000', ['J3']],
];
const REPEAT = Number(process.env.REPEAT || 3);

const browser = await chromium.launch({ channel: 'msedge', headless: true });
const ctx = await browser.newContext({ viewport: { width: 1366, height: 768 } });
const res = { base: BASE, started: new Date().toISOString(), cases: [] };
for (const [name, file, sheet, ref, text, watch] of CASES) {
  if (!fs.existsSync(file)) { res.cases.push({ name, skipped: 'thiếu file, chạy make_limit_files.py' }); continue; }
  const runs = [];
  for (let i = 0; i < REPEAT; i++) {
    const pg = await ctx.newPage();
    await pg.goto(BASE);
    await pg.waitForSelector('.welcome');
    await pg.setInputFiles('[data-testid=file-input]', file);
    await pg.waitForFunction(() => window.vmGridReady, null, { timeout: 300000 });
    await pg.evaluate(() => window.vmGrid.settled());
    const before = await pg.evaluate(([s, w]) => w.map((x) => window.vmGrid.value(s, x)), [sheet, watch]);
    await pg.evaluate(([s, r]) => window.vmGrid.goTo(s, r), [sheet, ref]);
    await sleep(150);
    await pg.keyboard.type(text + String(i));
    const t0 = Date.now();
    await pg.keyboard.press('Enter');
    // Theo dõi: lúc nào ô đã sửa có số mới mà ô công thức còn số cũ (số cũ hiện trên màn hình), và lúc nào mọi ô công thức đổi xong.
    const timeline = await pg.evaluate(async ([s, r, w, old]) => {
      const t0 = performance.now();
      let staleFrom = null; let staleTo = null; let doneAt = null;
      while (performance.now() - t0 < 60000) {
        const edited = window.vmGrid.value(s, r);
        const now = w.map((x) => window.vmGrid.value(s, x));
        const changed = now.every((v, k) => v !== old[k]);
        const editedIn = typeof edited === 'number';
        if (editedIn && !changed && staleFrom === null) staleFrom = performance.now() - t0;
        if (changed) { doneAt = performance.now() - t0; if (staleFrom !== null) staleTo = doneAt; break; }
        await new Promise((ok) => setTimeout(ok, 20));
      }
      return { staleFrom, staleTo, doneAt };
    }, [sheet, ref, watch, before]);
    const settledStart = Date.now();
    await pg.evaluate(() => window.vmGrid.settled());
    runs.push({ ...timeline, stale_ms: timeline.staleFrom === null ? 0 : Math.round(timeline.staleTo - timeline.staleFrom), done_ms: timeline.doneAt === null ? null : Math.round(timeline.doneAt), enter_to_poll_ms: Date.now() - t0 - (Date.now() - settledStart) });
    await pg.close();
  }
  res.cases.push({ name, runs, max_stale_ms: Math.max(...runs.map((x) => x.stale_ms)), max_done_ms: Math.max(...runs.map((x) => x.done_ms ?? 60000)) });
  console.log(name, JSON.stringify(res.cases.at(-1)));
}
await browser.close();
fs.writeFileSync(path.join(OUT, 'recalc_delay.json'), JSON.stringify(res, null, 1));
