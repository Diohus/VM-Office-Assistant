// T-DEV-05: ca giới hạn nặng trên giao diện, CHỈ bộ thử 8089. Không nằm trong run.sh. Chạy: node limits_heavy.mjs [ca...]
// Ca: big (GH-05a/f), tabs (GH-01), pdf (GH-04), paste (GH-11b), stop (GH-08a, GH-08c). Kết quả: ../out/limits_heavy.json
import { chromium } from 'playwright-core';
import { execSync } from 'node:child_process';
import fs from 'node:fs';
import path from 'node:path';

const BASE = process.env.VM_BASE || 'http://127.0.0.1:8089';
const PROJECT = process.env.VM_PROJECT || 'vm-office-test';
if (BASE.includes('8088') || !PROJECT.endsWith('test')) throw new Error('Chỉ chạy trên bộ thử 8089');
const FIX = path.resolve('..', '..', '..', 'fixtures', 'templates');
const GEN = path.resolve('..', '..', '..', 'fixtures', 'generated');
const OUT = path.resolve('..', 'out');
const sleep = (ms) => new Promise((r) => setTimeout(r, ms));
const only = process.argv.slice(2);
const want = (c) => !only.length || only.includes(c);
const res = { base: BASE, started: new Date().toISOString() };

function stats() {
  const out = execSync(`docker stats --no-stream --format "{{.Name}} {{.MemUsage}} {{.CPUPerc}}"`, { encoding: 'utf8' });
  return Object.fromEntries(out.trim().split('\n').filter((l) => l.startsWith(`${PROJECT}-`)).map((l) => {
    const [name, mem, , , cpu] = l.split(' ');
    return [name.replace(`${PROJECT}-`, '').replace('-1', ''), `${mem} ${cpu}`];
  }));
}
const docker = (cmd) => execSync(`docker ${cmd}`, { encoding: 'utf8', timeout: 120000 }).trim();

const browser = await chromium.launch({ channel: 'msedge', headless: true });
const ctx = await browser.newContext({ viewport: { width: 1366, height: 768 }, acceptDownloads: true });

async function openFile(pg, file, timeout = 300000) {
  await pg.goto(BASE);
  await pg.waitForSelector('.welcome');
  await pg.evaluate(() => { window.vmGridReady = null; });
  const t0 = Date.now();
  await pg.setInputFiles('[data-testid=file-input]', file);
  await pg.waitForFunction(() => window.vmGridReady || document.querySelector('[data-notice=open-error]'), null, { timeout });
  const err = await pg.$eval('[data-notice=open-error] .notice-text', (e) => e.firstChild.textContent).catch(() => null);
  if (!err) await pg.evaluate(() => window.vmGrid.settled());
  return { ms: Date.now() - t0, error: err };
}
async function typeCell(pg, sheet, ref, text) {
  await pg.evaluate(([s, r]) => window.vmGrid.goTo(s, r), [sheet, ref]);
  await sleep(150);
  await pg.keyboard.type(text);
  await pg.keyboard.press('Enter');
  await sleep(300);
}
async function save(pg, timeout = 300000) {
  await pg.evaluate(() => { window.vmGridReady = null; });
  const t0 = Date.now();
  await pg.click('.sidebar .side-btn:has-text("Lưu")');
  await sleep(150);
  if (await pg.$('[data-testid=name-input]')) {
    await pg.fill('[data-testid=name-input]', 'Người Giả');
    await pg.click('.dialog-actions button:has-text("Tiếp tục")');
  }
  await pg.waitForFunction(() => window.vmGridReady || document.querySelector('[data-notice=open-error]'), null, { timeout });
  const err = await pg.$eval('[data-notice=open-error] .notice-text', (e) => e.firstChild.textContent).catch(() => null);
  return { ms: Date.now() - t0, state: await pg.textContent('[data-testid=doc-state]'), error: err };
}
async function download(pg, label, timeout = 300000) {
  const t0 = Date.now();
  await pg.click('.sidebar .side-btn:has-text("Xuất")');
  await pg.waitForSelector(`.dialog-actions button:has-text("${label}")`, { timeout: 30000 });
  try {
    const [d] = await Promise.all([pg.waitForEvent('download', { timeout }), pg.click(`.dialog-actions button:has-text("${label}")`)]);
    const file = path.join(OUT, `heavy_${Date.now()}_${d.suggestedFilename()}`);
    await d.saveAs(file);
    const bytes = fs.statSync(file).size;
    fs.unlinkSync(file);
    return { ms: Date.now() - t0, bytes };
  } catch (e) {
    const err = await pg.$$eval('[data-notice] .notice-text', (l) => l.map((x) => x.firstChild.textContent).join(' | ')).catch(() => null);
    return { ms: Date.now() - t0, failed: String(e.message).slice(0, 120), notice: err };
  }
}
const heap = (pg) => pg.evaluate(() => Math.round((performance.memory?.usedJSHeapSize ?? 0) / 1048576));

// GH-05a/f: biểu mẫu giả gần giới hạn ô (18002 dòng × 10 cột, 18000 ô công thức).
if (want('big')) {
  const pg = await ctx.newPage();
  const r = { file: 'GIA-GAN_200000_o.xlsx', before: stats() };
  r.open = await openFile(pg, path.join(GEN, 'GIA-GAN_200000_o.xlsx'));
  r.after_open = stats();
  if (!r.open.error) {
    r.heap_mb = await heap(pg);
    r.j3_before = await pg.evaluate(() => window.vmGrid.value('Lon', 'J3'));
    const t0 = Date.now();
    await typeCell(pg, 'Lon', 'B3', '1000');
    await pg.evaluate(() => window.vmGrid.settled());
    r.settled_ms = Date.now() - t0;
    r.j3_at_settled = await pg.evaluate(() => window.vmGrid.value('Lon', 'J3'));
    // settled() trả về trước khi bộ tính xong (đo 10/11), nên chờ tới khi J3 thật sự đổi.
    await pg.waitForFunction((old) => window.vmGrid.value('Lon', 'J3') !== old, r.j3_before, { timeout: 120000, polling: 100 });
    r.recalc_ms = Date.now() - t0;
    r.j3_after = await pg.evaluate(() => window.vmGrid.value('Lon', 'J3'));
    r.save = await save(pg);
    r.xlsx = await download(pg, 'Tải file Excel');
    r.pdf = await download(pg, 'Tải PDF', 400000);
    r.after_all = stats();
  }
  res.big = r;
  await pg.close();
  console.log('big', JSON.stringify(r));
}

// GH-01: 2, 5, 10 tab, mỗi tab một biểu mẫu, sửa một ô rồi lưu.
if (want('tabs')) {
  const files = ['GIA-QC-01_mau.xlsx', 'GIA-KHO-02_tieu_hao.xlsx', 'GIA-NGUON_2026-09.xlsx', 'GIA-X18_ham_mang.xlsx', 'GIA-X18b_lien_ket_ngoai.xlsx']; // không dùng X19_sheet_khoa: B8 bị khóa, không có sửa để lưu
  const cell = { 'GIA-QC-01_mau.xlsx': ['KQ_QC_Thang', 'B8'], 'GIA-KHO-02_tieu_hao.xlsx': [null, 'C4'], 'GIA-NGUON_2026-09.xlsx': [null, 'H3'], 'GIA-X18_ham_mang.xlsx': ['KQ_QC_Thang', 'B8'], 'GIA-X18b_lien_ket_ngoai.xlsx': ['KQ_QC_Thang', 'B8'] };
  res.tabs = {};
  for (const n of [2, 5, 10]) {
    const pages = [];
    const r = { open_ms: [], save: [], errors: [] };
    for (let i = 0; i < n; i++) {
      const pg = await ctx.newPage();
      pages.push(pg);
      const name = files[i % files.length];
      const o = await openFile(pg, path.join(FIX, name));
      r.open_ms.push(o.ms);
      if (o.error) r.errors.push(`${name}: ${o.error}`);
    }
    for (let i = 0; i < n; i++) {
      const pg = pages[i];
      const name = files[i % files.length];
      const [sheet0, ref] = cell[name];
      const sheet = sheet0 ?? (await pg.evaluate(() => window.vmGrid.activeSheet()));
      await typeCell(pg, sheet, ref, String(100 + i));
      const s = await save(pg);
      r.save.push(s.ms);
      if (s.error) r.errors.push(`${name} lưu: ${s.error}`);
      r.saved_value_ok = (r.saved_value_ok ?? true) && (await pg.evaluate(([sh, rf]) => window.vmGrid.value(sh, rf), [sheet, ref])) === 100 + i;
    }
    r.heap_mb_total = (await Promise.all(pages.map(heap))).reduce((a, b) => a + b, 0);
    r.stats = stats();
    for (const pg of pages) await pg.close();
    res.tabs[n] = r;
    console.log('tabs', n, JSON.stringify(r));
  }
}

// GH-04: 5 tab đã sửa, bấm Tải PDF gần như cùng lúc.
if (want('pdf')) {
  const pages = [];
  for (let i = 0; i < 5; i++) {
    const pg = await ctx.newPage();
    pages.push(pg);
    await openFile(pg, path.join(FIX, 'GIA-QC-01_da_dien.xlsx'));
    await typeCell(pg, 'KQ_QC_Thang', 'B8', String(5 + i / 10));
    await save(pg);
  }
  const t0 = Date.now();
  const results = await Promise.all(pages.map((pg) => download(pg, 'Tải PDF', 400000)));
  res.pdf_concurrent = { wall_ms: Date.now() - t0, each: results, stats: stats() };
  for (const pg of pages) await pg.close();
  console.log('pdf', JSON.stringify(res.pdf_concurrent));
}

// GH-11b: dán vùng lớn vào biểu mẫu lớn: 1000 × 10 (quá 5000 ô, phải bị chặn khi lưu), rồi 400 × 10 lưu được.
if (want('paste')) {
  const pg = await ctx.newPage();
  const r = {};
  r.open = await openFile(pg, path.join(GEN, 'GIA-GAN_200000_o.xlsx'));
  const block = (rows, cols) => Array.from({ length: rows }, (_, i) => Array.from({ length: cols }, (_, k) => String(i * 10 + k)).join(String.fromCharCode(9))).join(String.fromCharCode(10));
  const paste = async (ref, text) => {
    await pg.evaluate(([r0]) => window.vmGrid.goTo('Lon', r0), [ref]);
    await sleep(200);
    const t0 = Date.now();
    await pg.evaluate((t) => { const dt = new DataTransfer(); dt.setData('text/plain', t); document.activeElement.dispatchEvent(new ClipboardEvent('paste', { clipboardData: dt, bubbles: true, cancelable: true })); }, text);
    await pg.evaluate(() => window.vmGrid.settled());
    return Date.now() - t0;
  };
  r.paste_10000_ms = await paste('A20', block(1000, 10));
  r.dirty_10000 = await pg.evaluate(() => window.vmGrid.collectEdits().length);
  r.save_10000 = await save(pg, 60000).catch((e) => ({ failed: String(e.message).slice(0, 100) }));
  r.save_10000.notice = await pg.$eval('[data-notice=open-error] .notice-text', (e) => e.firstChild.textContent.slice(0, 80)).catch(() => null);
  await pg.click('.sidebar .side-btn:has-text("Xuất")');
  await pg.waitForSelector('[data-dialog="Có sửa chưa lưu"]');
  await pg.click('.dialog-actions button:has-text("Bỏ thay đổi")');
  await pg.waitForSelector('.dialog-actions button:has-text("Hủy")', { timeout: 120000 });
  await pg.click('.dialog-actions button:has-text("Hủy")');
  await sleep(500);
  r.paste_4000_ms = await paste('A20', block(400, 10));
  r.save_4000 = await save(pg);
  r.stats = stats();
  res.paste = r;
  await pg.close();
  console.log('paste', JSON.stringify(r));
}

// GH-08a: tắt api khi có sửa chưa lưu. GH-08c: tắt converter ngay khi bấm Tải PDF.
if (want('stop')) {
  const pg = await ctx.newPage();
  const r = {};
  await openFile(pg, path.join(FIX, 'GIA-QC-01_mau.xlsx'));
  await typeCell(pg, 'KQ_QC_Thang', 'B8', '6.25');
  docker(`stop ${PROJECT}-api-1`);
  r.api_down = await save(pg, 60000);
  r.api_down.b8_on_screen = await pg.evaluate(() => window.vmGrid.value('KQ_QC_Thang', 'B8'));
  docker(`start ${PROJECT}-api-1`);
  for (let i = 0; i < 60; i++) {
    const ok = await pg.evaluate(async (b) => (await fetch(`${b}/api/health`).catch(() => null))?.ok, BASE);
    if (ok) break;
    await sleep(1000);
  }
  await pg.click('[data-notice=open-error] button:has-text("Đóng")').catch(() => {});
  r.api_back = await save(pg);
  r.api_back.b8 = await pg.evaluate(() => window.vmGrid.value('KQ_QC_Thang', 'B8'));
  const t0 = Date.now();
  await pg.click('.sidebar .side-btn:has-text("Xuất")');
  await pg.waitForSelector('.dialog-actions button:has-text("Tải PDF")');
  await pg.click('.dialog-actions button:has-text("Tải PDF")');
  await sleep(300);
  docker(`stop ${PROJECT}-converter-1`);
  await pg.waitForSelector('[data-notice=open-error], [data-notice=pdf-mismatch]', { timeout: 180000 }).catch(() => {});
  r.converter_down = { ms: Date.now() - t0, notice: await pg.$$eval('[data-notice] .notice-text', (l) => l.map((x) => x.firstChild.textContent).join(' | ')).catch(() => null) };
  docker(`start ${PROJECT}-converter-1`);
  r.xlsx_while_converter_restarting = await download(pg, 'Tải file Excel', 60000);
  res.stop = r;
  await pg.close();
  console.log('stop', JSON.stringify(r));
}

await browser.close();
res.finished = new Date().toISOString();
const outFile = path.join(OUT, 'limits_heavy.json');
const prev = fs.existsSync(outFile) ? JSON.parse(fs.readFileSync(outFile, 'utf8')) : {};
fs.writeFileSync(outFile, JSON.stringify({ ...prev, ...res }, null, 1));
