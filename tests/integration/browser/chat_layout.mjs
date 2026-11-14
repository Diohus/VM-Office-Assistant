// Lát 2, Owner thử 13/11 (O1, O2): docs/product/acceptance_lat2.md mục 1.5, 1.6, 1.7. Chỉ bộ thử 8089, model là máy giả fakeai.
// Cửa sổ như máy Owner: màn 1366x768, trình duyệt phóng to, thu phóng 100%: vùng trang còn khoảng 1366x657.
// Mọi cú bấm đi bằng chuột theo tọa độ (page.mouse), không gọi click trên phần tử. Trước khi bấm, kiểm phần tử nằm ở tọa độ đó
// đúng là nút (không phải ô gõ hay phần khác vẽ đè), và nút nằm trọn trong khung chat và trong cửa sổ.
import { chromium } from 'playwright-core';
import fs from 'node:fs';
import path from 'node:path';

const BASE = process.env.VM_BASE || 'http://127.0.0.1:8089';
const FILE = path.resolve('..', '..', '..', 'fixtures', 'templates', 'chat_lat2', 'GIA-QC-01_chat.xlsx');
const OUT = path.resolve('..', 'out');
const SHEET = 'KQ_QC_Thang';
const VIEW = { width: 1366, height: 657 };
const Y2 = 'Glucose mức 1 ngày 3/9 là 5,62 FAKE:edit';
const sleep = (ms) => new Promise((r) => setTimeout(r, ms));
const res = { page_errors: [], outbound: [], clicks: [], selections: {} };

const browser = await chromium.launch({ channel: 'msedge', headless: true });
const ctx = await browser.newContext({ viewport: VIEW, deviceScaleFactor: 1 });
const pg = await ctx.newPage();
pg.on('pageerror', (e) => res.page_errors.push(String(e.message).slice(0, 200)));
pg.on('request', (r) => {
  if (!r.url().startsWith(BASE) && !r.url().startsWith('data:') && !r.url().startsWith('blob:')) res.outbound.push(r.url().slice(0, 80));
});

const val = (ref) => pg.evaluate(([s, r]) => window.vmGrid.value(s, r), [SHEET, ref]);
const turns = () => pg.$$eval('[data-testid=chat] .chat-turn', (els) => els.length);
const lastTurn = () => pg.$eval('[data-testid=chat] .chat-turn:last-child', (el) => ({
  text: el.innerText,
  codes: [...el.querySelectorAll('[data-code]')].map((x) => x.dataset.code),
  summary: el.querySelectorAll('.chat-summary li').length,
  skipped: [...el.querySelectorAll('.chat-skipped li')].map((x) => `${x.dataset.cell}:${x.dataset.code}`),
  already: [...el.querySelectorAll('.chat-already li')].map((x) => `${x.dataset.cell}:${x.dataset.code}`),
  model: el.querySelector('.chat-model')?.innerText ?? null,
}));

// Bấm bằng chuột vào giữa phần tử. Ghi lại: nút có nằm trọn trong cửa sổ và khung chat không, phần tử ở tọa độ đó là gì.
async function mouseClick(selector, label) {
  const info = await pg.evaluate(([sel, view]) => {
    const el = document.querySelector(sel);
    if (!el) return { found: false };
    const r = el.getBoundingClientRect();
    const chat = document.querySelector('[data-testid=chat]')?.getBoundingClientRect();
    const x = r.left + r.width / 2, y = r.top + r.height / 2;
    const hit = document.elementFromPoint(x, y);
    const inside = (b) => b && r.left >= b.left - 0.5 && r.right <= b.right + 0.5 && r.top >= b.top - 0.5 && r.bottom <= b.bottom + 0.5;
    return {
      found: true, x, y, text: el.innerText?.trim() ?? '', disabled: Boolean(el.disabled),
      in_view: inside({ left: 0, top: 0, right: view.width, bottom: view.height }),
      in_chat: chat ? inside(chat) : null,
      hit_ok: Boolean(hit && (hit === el || el.contains(hit))),
      hit: hit ? `${hit.tagName.toLowerCase()}${hit.dataset?.testid ? `[${hit.dataset.testid}]` : ''}.${hit.className}` : null,
    };
  }, [selector, VIEW]);
  res.clicks.push({ label, ...info });
  if (!info.found) throw new Error(`không thấy ${label}`);
  if (!info.in_view || !info.hit_ok || info.in_chat === false) throw new Error(`${label}: không bấm được bằng chuột (y=${Math.round(info.y)}, ở tọa độ đó là ${info.hit})`);
  await pg.mouse.click(info.x, info.y);
  return info;
}

async function ready() {
  await pg.waitForFunction(() => window.vmGridReady, null, { timeout: 90000 });
  // R-1021: mỗi lần mở file, lưới phải thật sự được vẽ (có canvas lớn trong .grid-host), không chỉ có mô hình bảng.
  await pg.waitForFunction(() => [...document.querySelectorAll('.grid-host canvas')].some((c) => c.getBoundingClientRect().height > 100), null, { timeout: 15000 })
    .catch(() => { throw new Error('lưới trắng: .grid-host không có canvas sau khi mở file (R-1021)'); });
  await pg.evaluate(() => window.vmGrid.settled());
  await sleep(300);
}
// Gõ vào ô chat bằng bàn phím sau khi bấm chuột vào ô, rồi bấm Gửi bằng chuột. Trả 'preview' hoặc 'turn'.
async function ask(text, label) {
  const before = await turns();
  await mouseClick('[data-testid=chat-input]', `ô gõ (${label})`);
  await pg.keyboard.press('Control+a');
  await pg.keyboard.press('Delete');
  await pg.keyboard.insertText(text);
  await mouseClick('[data-testid=chat-send]', `Gửi (${label})`);
  await pg.waitForFunction((n) => document.querySelector('[data-testid=chat-preview]') || document.querySelectorAll('[data-testid=chat] .chat-turn').length > n,
    before, { timeout: 90000 });
  return (await pg.$('[data-testid=chat-preview]')) ? 'preview' : 'turn';
}
async function waitTurn(before) {
  await pg.waitForFunction((n) => document.querySelectorAll('[data-testid=chat] .chat-turn').length > n, before, { timeout: 60000 });
  await sleep(300);
}
// Lưới chính là canvas lớn nhất trong vùng lưới.
// Chờ lưới vẽ xong (tối đa 15 giây). Lần chạy đủ tầng 1 ngày 13/11 có một lần lưới trống không có canvas: nếu vẫn trống thì
// báo rõ, không bỏ qua.
async function gridBox() {
  await pg.waitForFunction(() => [...document.querySelectorAll('.grid-host canvas')].some((c) => c.getBoundingClientRect().height > 100), null, { timeout: 15000 })
    .catch(() => { throw new Error('lưới trống: không có canvas nào trong vùng lưới sau 15 giây'); });
  const boxes = await pg.$$eval('.grid-host canvas', (els) => els.map((e) => e.getBoundingClientRect()).map((r) => ({ x: r.left, y: r.top, width: r.width, height: r.height })));
  return boxes.sort((a, b) => b.width * b.height - a.width * a.height)[0];
}
// Mọi nút đang hiện trong khung chat: chữ, có bị khóa không.
// Nút "Hoàn tác lượt này" của lượt cũ nằm trong danh sách lượt, có câu giải thích riêng khi bị khóa: không thuộc phép kiểm 1.6.
const chatButtons = () => pg.$$eval('[data-testid=chat] button', (els) => els.filter((e) => e.offsetParent !== null && !e.closest('.chat-turns')).map((e) => ({ text: e.innerText.trim(), disabled: e.disabled })));

try {
await pg.goto(BASE);
await pg.waitForSelector('.welcome');
await pg.setInputFiles('[data-testid=file-input]', FILE);
await ready();
await mouseClick('[data-testid=chat-button]', 'mở khung chat');
await pg.waitForSelector('[data-testid=chat]');
res.chat_box = await pg.$eval('[data-testid=chat]', (e) => { const r = e.getBoundingClientRect(); return { width: r.width, height: r.height }; });

// 1.7 (O2): mọi cách chọn đều tới màn xem trước, không có CHAT_BAD_INPUT. "Không gửi" thì trả lại câu đã gõ.
const selections = [
  ['chưa chọn gì', async () => {}],
  ['bấm chuột một ô', async () => { const b = await gridBox(); await pg.mouse.click(b.x + b.width / 2, b.y + b.height / 2); }],
  ['một ô B10', async () => pg.evaluate((s) => window.vmGrid.goTo(s, 'B10'), SHEET)],
  ['ô cột H (H10)', async () => pg.evaluate((s) => window.vmGrid.goTo(s, 'H10'), SHEET)],
  ['ô gộp tiêu đề A1', async () => pg.evaluate((s) => window.vmGrid.goTo(s, 'A1'), SHEET)],
  ['cả cột H', async () => pg.evaluate((s) => window.vmGrid.goTo(s, 'H:H'), SHEET)],
  // Chuột trên tiêu đề cột, tiêu đề hàng, góc trên trái: Univer ghi vùng "B:B", "3:3", "1:65" (nguyên nhân O2 tìm được 13/11).
  ['tiêu đề một cột (chuột)', async () => { const b = await gridBox(); await pg.mouse.click(b.x + 260, b.y + 6); }],
  ['tiêu đề một hàng (chuột)', async () => { const b = await gridBox(); await pg.mouse.click(b.x + 8, b.y + 100); }],
  ['cả sheet (góc trên trái)', async () => { const b = await gridBox(); await pg.mouse.click(b.x + 8, b.y + 6); }],
];
for (const [label, select] of selections) {
  await select();
  await sleep(250);
  const kind = await ask(Y2, label);
  const entry = { kind, turn: kind === 'turn' ? await lastTurn() : null };
  if (kind === 'preview') {
    entry.sheets = await pg.$$eval('[data-testid=chat-sheets] li', (els) => els.map((e) => e.textContent));
    await mouseClick('[data-testid=chat-cancel-preview]', `Không gửi (${label})`);
    await sleep(200);
    entry.preview_gone = (await pg.$('[data-testid=chat-preview]')) === null;
    entry.input_back = await pg.$eval('[data-testid=chat-input]', (e) => e.value);
  }
  res.selections[label] = entry;
}

// N2 (ChatEditPlan 0.2): chọn sẵn vùng số B8:C37 thì gửi kèm A1:C37 chỉ để đọc (tiêu đề, cột Ngày); AI chỉ ghi trong B8:C37.
await pg.evaluate((s) => window.vmGrid.goTo(s, 'B8:C37'), SHEET);
await sleep(250);
const a10Before = await val('A10');
const b10Start = await val('B10');
res.n2_kind = await ask('Glucose mức 1 ngày 3/9 là 5,62 FAKE:read_only', 'chọn sẵn B8:C37');
if (res.n2_kind === 'preview') {
  res.n2_preview = {
    sheets: await pg.$$eval('[data-testid=chat-sheets] li', (els) => els.map((e) => e.textContent)),
    context: await pg.$$eval('[data-testid=chat-preview-context]', (els) => els.map((e) => e.textContent)),
    prompt: await pg.textContent('[data-testid=chat-prompt]'),
  };
  const n = await turns();
  await mouseClick('[data-testid=chat-send-preview]', 'Đồng ý gửi (B8:C37)');
  await waitTurn(n);
}
res.n2 = { turn: await lastTurn(), a10: await val('A10'), a10_before: a10Before, b10: await val('B10'), b10_before: b10Start };
if (res.n2.turn.summary) {
  await mouseClick('[data-testid=chat] .chat-turn:last-child [data-undo-turn]', 'Hoàn tác lượt này (B8:C37)');
  await sleep(400);
}

// O2 gốc: ô công thức chưa có kết quả (bộ tính còn chạy, ví dụ ngay sau lượt AI ghi công thức). Trước đây gửi v null và vmai
// từ chối cả lượt. Giữ bộ tính như đang chạy, hạn chờ ngắn, ghi công thức vào ô trống H40.
await pg.evaluate((s) => window.vmGrid.goTo(s, 'B10'), SHEET);
await pg.evaluate(() => { window.vmGrid.testCalcTimeout(3000); window.vmGrid.testHoldCalculation(true); });
await pg.evaluate((s) => window.vmGrid.applyChatOps(s, [{ op: 'set_formula', cell: 'H40', formula: '=B10*2' }], {}), SHEET);
res.pending_formula_value = await val('H40');
// N12: trong lúc chuẩn bị (ở đây: chờ bộ tính, rồi chờ màn xem trước), khung không được nói "AI đang làm…".
const pendingTurns = await turns();
await mouseClick('[data-testid=chat-input]', 'ô gõ (công thức chưa tính)');
await pg.keyboard.press('Control+a');
await pg.keyboard.press('Delete');
await pg.keyboard.insertText(Y2);
await mouseClick('[data-testid=chat-send]', 'Gửi (công thức chưa tính)');
await pg.waitForSelector('[data-testid=chat-wait]', { timeout: 10000 });
res.preparing = await pg.$eval('[data-testid=chat-wait] span', (e) => ({ text: e.textContent, phase: e.dataset.phase }));
await pg.waitForFunction((n) => document.querySelector('[data-testid=chat-preview]') || document.querySelectorAll('[data-testid=chat] .chat-turn').length > n, pendingTurns, { timeout: 90000 });
const kindPending = (await pg.$('[data-testid=chat-preview]')) ? 'preview' : 'turn';
res.selections['công thức chưa tính'] = { kind: kindPending, turn: kindPending === 'turn' ? await lastTurn() : null };
if (kindPending === 'preview') await mouseClick('[data-testid=chat-cancel-preview]', 'Không gửi (công thức chưa tính)');
await pg.evaluate(() => { window.vmGrid.testHoldCalculation(false); window.vmGrid.testCalcTimeout(null); });
await pg.evaluate(() => window.vmGrid.undo());
await sleep(300);

// 1.5, 1.6: màn xem trước dài hơn khung (cả KQ_QC_Thang cùng GioiHan). Nút trong khung bấm được bằng chuột; không còn nút
// Gửi chính hay ô gõ đang khóa; phần thân cuộn.
const b = await gridBox();
await pg.mouse.click(b.x + 8, b.y + 6);
await sleep(250);
res.long_kind = await ask(Y2, 'xem trước dài');
res.long = await pg.evaluate(() => {
  const body = document.querySelector('[data-testid=chat-preview-body]');
  return {
    body_scrolls: body.scrollHeight > body.clientHeight + 4,
    main_send: document.querySelector('[data-testid=chat-send]') !== null,
    input: document.querySelector('[data-testid=chat-input]') !== null,
    sheets: [...document.querySelectorAll('[data-testid=chat-sheets] li')].map((e) => e.dataset.sheet),
  };
});
res.long_buttons = await chatButtons();
// 1.6 (PM): nút bị khóa trong danh sách lượt phải trông bị khóa thật: mờ, con trỏ không phải bàn tay.
res.turn_disabled_look = await pg.$$eval('[data-testid=chat] .chat-turns button:disabled', (els) => els.map((e) => ({ text: e.innerText.trim(), opacity: Number(getComputedStyle(e).opacity), cursor: getComputedStyle(e).cursor })));
await pg.screenshot({ path: path.join(OUT, 'chat_o1_xem_truoc.png') });
const b10Before = await val('B10');
let before = await turns();
await mouseClick('[data-testid=chat-send-preview]', 'Đồng ý gửi');
await waitTurn(before);
res.after_send = { turn: await lastTurn(), b10: await val('B10'), b10_before: b10Before };
await mouseClick('[data-testid=chat] .chat-turn:last-child [data-undo-turn]', 'Hoàn tác lượt này');
await sleep(400);
res.after_undo_b10 = await val('B10');

// 1.5: câu trả lời dài (AI hỏi lại).
res.long_ask_kind = await ask('Sửa giúp tôi FAKE:long_ask', 'câu trả lời dài');
res.long_ask = await lastTurn();
await mouseClick('[data-testid=chat-input]', 'ô gõ sau câu trả lời dài');
await pg.screenshot({ path: path.join(OUT, 'chat_o1_tra_loi_dai.png') });

// 1.5: tóm tắt 30 ô.
const nums = Array.from({ length: 30 }, (_, i) => `4.${String(i + 1).padStart(2, '0')}`).join(' ');
res.many_kind = await ask(`Glucose mức 1 từ ngày 1/9 đến 30/9 lần lượt là ${nums} FAKE:many`, 'tóm tắt 30 ô');
if (res.many_kind === 'preview') {
  before = await turns();
  await mouseClick('[data-testid=chat-send-preview]', 'Đồng ý gửi (30 ô)');
  await waitTurn(before);
}
res.many = await lastTurn();
await pg.screenshot({ path: path.join(OUT, 'chat_o1_30_o.png') });
await mouseClick('[data-testid=chat] .chat-turn:last-child [data-undo-turn]', 'Hoàn tác lượt này (30 ô)');
await sleep(400);
res.many_undo_b8 = await val('B8');

// K4 (PM chọn 2): AI ghi vào dòng thống kê ngoài thân bảng (I38:I44) vẫn được áp, nhưng tóm tắt tách nhóm có câu cảnh báo.
await pg.evaluate((s) => window.vmGrid.goTo(s, 'A1'), SHEET);
res.outside_kind = await ask('Ghi chạy lại khi ALT lớn hơn 44 FAKE:outside', 'ô ngoài thân bảng'); // hằng số 44, "chạy lại" phải có trong câu (Q135 (i))
if (res.outside_kind === 'preview') {
  before = await turns();
  await mouseClick('[data-testid=chat-send-preview]', 'Đồng ý gửi (ngoài thân bảng)');
  await waitTurn(before);
}
res.outside = await pg.$eval('[data-testid=chat] .chat-turn:last-child', (el) => ({
  normal: [...el.querySelectorAll('.chat-summary li:not([data-outside])')].map((x) => x.dataset.cell),
  outside: [...el.querySelectorAll('[data-testid=chat-outside] li')].map((x) => x.dataset.cell),
  warn: el.querySelector('[data-testid=chat-outside] .chat-warn')?.textContent ?? null,
  skipped: [...el.querySelectorAll('.chat-skipped li')].map((x) => `${x.dataset.cell}:${x.dataset.code}`),
}));
res.outside.i38_formula = await pg.evaluate((s) => window.vmGrid.formula(s, 'I38'), SHEET);
if (res.outside.normal.length + res.outside.outside.length) {
  await mouseClick('[data-testid=chat] .chat-turn:last-child [data-undo-turn]', 'Hoàn tác lượt này (ngoài thân bảng)');
  await sleep(400);
}

// N7: ô đã đúng như yêu cầu không tính vào "không áp", nêu riêng.
const b12 = await val('B12');
res.already_kind = await ask(`Glucose mức 1 ngày 3/9 là 5,62, ngày 5/9 là ${b12} FAKE:already:${b12}`, 'ô đã đúng sẵn');
if (res.already_kind === 'preview') {
  before = await turns();
  await mouseClick('[data-testid=chat-send-preview]', 'Đồng ý gửi (đã đúng sẵn)');
  await waitTurn(before);
}
res.already = { turn: await lastTurn(), b12, b12_after: await val('B12') };
await mouseClick('[data-testid=chat] .chat-turn:last-child [data-undo-turn]', 'Hoàn tác lượt này (đã đúng sẵn)');
await sleep(400);

// N4: biểu mẫu chưa điền được bằng Điền từ file (file chat): câu không nói "App mở Điền từ file".
res.fill_kind = await ask('Nhập số từ file máy xét nghiệm FAKE:fill_file', 'cần Điền từ file');
res.fill = { turn: await lastTurn(), fill_panel: (await pg.$('.fill-panel')) !== null, page: await pg.textContent('body') };

// 6.5: lượt hỏng ở bước gửi (mạng, hết giờ): câu trở lại ô gõ nguyên văn. Lỗi do người dùng (câu có tên): câu vẫn trong ô gõ.
const failMsg = 'Glucose mức 1 ngày 3/9 là 5,62 FAKE:echo_500';
res.retry_kind = await ask(failMsg, 'lỗi khi gửi');
if (res.retry_kind === 'preview') {
  before = await turns();
  await mouseClick('[data-testid=chat-send-preview]', 'Đồng ý gửi (lỗi khi gửi)');
  await waitTurn(before);
}
res.retry = { turn: await lastTurn(), input: await pg.$eval('[data-testid=chat-input]', (e) => e.value), message: failMsg };
const personMsg = 'Hôm nay Trần Giả Định chạy, Glucose mức 1 ngày 5/9 là 5,40';
res.person_kind = await ask(personMsg, 'câu có tên');
res.person = { turn: await lastTurn(), input: await pg.$eval('[data-testid=chat-input]', (e) => e.value), message: personMsg };
await mouseClick('[data-testid=chat-input]', 'ô gõ (xóa câu có tên)');
await pg.keyboard.press('Control+a');
await pg.keyboard.press('Delete');

// 1.5: Hủy khi đang chờ.
await mouseClick('[data-testid=chat-input]', 'ô gõ (chờ)');
await pg.keyboard.insertText('Glucose mức 1 ngày 3/9 là 5,62 FAKE:slow');
await mouseClick('[data-testid=chat-send]', 'Gửi (chờ)');
await pg.waitForSelector('[data-testid=chat-wait]', { timeout: 30000 });
await mouseClick('[data-testid=chat-wait] button', 'Hủy');
await sleep(500);
res.cancel = { turn: await lastTurn(), wait: (await pg.$('[data-testid=chat-wait]')) !== null };

// 1.5: đóng và mở lại khung.
await mouseClick('[data-testid=chat] .chat-head button', 'Đóng khung chat');
await sleep(200);
res.closed = (await pg.$('[data-testid=chat]')) === null;
await mouseClick('[data-testid=chat-button]', 'mở lại khung chat');
await pg.waitForSelector('[data-testid=chat]');
res.reopened = true;
await sleep(8000); // kết quả về muộn của lượt đã Hủy không được áp

// 2.13 (Q135 (f)): sheet phụ được công thức tham chiếu nhưng không vừa trần (DuLieu_Nam 3.010 ô) thì trình duyệt không gửi.
await pg.goto(BASE);
await pg.waitForSelector('.welcome');
await pg.evaluate(() => { window.vmGridReady = null; });
await pg.setInputFiles('[data-testid=file-input]', path.resolve('..', '..', '..', 'fixtures', 'templates', 'chat_lat2', 'GIA-QC-01_chat_sheet_lon.xlsx'));
await ready();
res.big_snapshot = await pg.evaluate(() => {
  const s = window.vmGrid.chatSnapshot([]);
  const area = (r) => { const m = /^([A-Z]+)(\d+):([A-Z]+)(\d+)$/.exec(r); const col = (x) => [...x].reduce((n, ch) => n * 26 + ch.charCodeAt(0) - 64, 0); return (Number(m[4]) - Number(m[2]) + 1) * (col(m[3]) - col(m[1]) + 1); };
  return { sheets: s.request.sheets.map((x) => x.name), total: s.request.sheets.reduce((n, x) => n + area(x.context_region ?? x.region), 0), skipped: s.skipped };
});
await mouseClick('[data-testid=chat-button]', 'mở khung chat (sheet lớn)');
await pg.waitForSelector('[data-testid=chat]');
res.big_kind = await ask(Y2, 'sheet phụ quá lớn');
res.big_preview = res.big_kind === 'preview'
  ? { sheets: await pg.$$eval('[data-testid=chat-sheets] li', (els) => els.map((e) => e.dataset.sheet)) }
  : { turn: await lastTurn() };
} catch (e) {
  res.failed = String(e.message).slice(0, 300);
  await pg.screenshot({ path: path.join(OUT, 'chat_layout_loi.png') }).catch(() => {});
  process.exitCode = 1;
} finally {
  await browser.close();
  fs.writeFileSync(path.join(OUT, 'chat_layout.json'), JSON.stringify(res, null, 1));
}
