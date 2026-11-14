// R-1021 (lưới trắng khi máy nặng): 4 trình duyệt chạy song song, mỗi trình duyệt mở file 10 lần, mỗi lần trong một ngữ cảnh mới.
// Trước khi sửa (chờ một khung hình rồi đọc hostRef): đo 7/40 lần lưới trắng. Mỗi lần mở: phải có canvas lớn trong .grid-host.
// Chỉ bộ thử 8089, chỉ file giả.
import { chromium } from 'playwright-core';
import fs from 'node:fs';
import path from 'node:path';

const BASE = process.env.VM_BASE || 'http://127.0.0.1:8089';
const FILE = path.resolve('..', '..', '..', 'fixtures', 'templates', 'chat_lat2', 'GIA-QC-01_chat.xlsx');
const OUT = path.resolve('..', 'out');
const WORKERS = 4;
const RUNS = Number(process.env.GRID_LOAD_RUNS || 10);

async function worker(w) {
  const browser = await chromium.launch({ channel: 'msedge', headless: true });
  const rows = [];
  try {
    for (let i = 0; i < RUNS; i++) {
      const ctx = await browser.newContext({ viewport: { width: 1366, height: 657 } });
      const pg = await ctx.newPage();
      const errors = [];
      pg.on('pageerror', (e) => errors.push(String(e.message).slice(0, 160)));
      const row = { w, i };
      try {
        await pg.goto(BASE);
        await pg.waitForSelector('.welcome', { timeout: 60000 });
        await pg.setInputFiles('[data-testid=file-input]', FILE);
        await pg.waitForFunction(() => window.vmGridReady, null, { timeout: 120000 });
        row.drawn = await pg.waitForFunction(() => [...document.querySelectorAll('.grid-host canvas')].some((c) => c.getBoundingClientRect().height > 100), null, { timeout: 15000 })
          .then(() => true, () => false);
        row.host_children = await pg.$eval('.grid-host', (e) => e.childElementCount);
        row.error_notice = await pg.$eval('body', (b) => b.innerText.includes('App chưa làm xong việc này'));
      } catch (e) {
        row.failed = String(e.message).slice(0, 200);
      }
      row.page_errors = errors;
      rows.push(row);
      await ctx.close();
    }
  } finally {
    await browser.close();
  }
  return rows;
}

const all = (await Promise.all(Array.from({ length: WORKERS }, (_, w) => worker(w)))).flat();
fs.writeFileSync(path.join(OUT, 'grid_load.json'), JSON.stringify({ runs: all.length, rows: all }, null, 1));
