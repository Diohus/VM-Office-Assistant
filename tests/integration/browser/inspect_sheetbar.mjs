// Công cụ xem nhanh thanh sheet và kiểu hiện số, khi cần dò giao diện của thư viện lưới. Không thuộc bộ test.
import { chromium } from 'playwright-core';
import path from 'node:path';

const BASE = process.env.VM_BASE || 'http://127.0.0.1:8089';
const FIX = path.resolve('..', '..', '..', 'fixtures', 'templates');
const browser = await chromium.launch({ channel: 'msedge', headless: true });
const page = await browser.newPage({ viewport: { width: 1366, height: 768 } });
await page.goto(BASE);
await page.setInputFiles('[data-testid=file-input]', path.join(FIX, process.argv[2] || 'GIA-QC-01_mau.xlsx'));
await page.waitForFunction(() => window.vmGridReady, null, { timeout: 60000 });
await new Promise((r) => setTimeout(r, 2000));
const html = await page.evaluate(() => {
  const walker = document.createTreeWalker(document.body, NodeFilter.SHOW_TEXT);
  let tab = null;
  while (walker.nextNode()) if (walker.currentNode.textContent.includes('KQ_QC_Thang') && !walker.currentNode.parentElement.closest('.notices')) { tab = walker.currentNode.parentElement; break; }
  if (!tab) return document.querySelector('.grid-host').innerHTML.slice(0, 500);
  let bar = tab;
  for (let i = 0; i < 6 && bar.parentElement; i++) bar = bar.parentElement;
  return bar.outerHTML.replace(/<svg.*?<\/svg>/gs, '<svg/>').slice(0, 2600);
});
console.log(html);
await browser.close();
