// Lát 2 (Q135): giao diện khung chat. Đối chiếu docs/product/acceptance_lat2.md. Chỉ bộ thử 8089, model là máy giả fakeai
// (câu trả lời chọn theo dấu FAKE:... trong câu gõ, tests/integration/fakeai/fakeai.py).
import { chromium } from 'playwright-core';
import fs from 'node:fs';
import path from 'node:path';

const BASE = process.env.VM_BASE || 'http://127.0.0.1:8089';
const FILE = path.resolve('..', '..', '..', 'fixtures', 'templates', 'chat_lat2', 'GIA-QC-01_chat.xlsx');
const OUT = path.resolve('..', 'out');
const SHEET = 'KQ_QC_Thang';
const sleep = (ms) => new Promise((r) => setTimeout(r, ms));
const res = { page_errors: [], outbound: [], llm_requests: [] };

const browser = await chromium.launch({ channel: 'msedge', headless: true });
const ctx = await browser.newContext({ viewport: { width: 1600, height: 900 } });
const pg = await ctx.newPage();
pg.on('pageerror', (e) => res.page_errors.push(String(e.message).slice(0, 200)));
pg.on('request', (r) => {
  if (!r.url().startsWith(BASE) && !r.url().startsWith('data:') && !r.url().startsWith('blob:')) res.outbound.push(r.url().slice(0, 80));
  if (r.url().includes('/send')) res.llm_requests.push(r.url().split('/api/')[1].split('/')[0]);
});

const val = (ref) => pg.evaluate(([s, r]) => window.vmGrid.value(s, r), [SHEET, ref]);
const bg = (ref) => pg.evaluate(([s, r]) => window.vmGrid.background(s, r), [SHEET, ref]);
const aiCells = () => pg.evaluate(() => window.vmGrid.aiCells().map((c) => `${c.sheet}!${c.cell}`));
const turns = () => pg.$$eval('[data-testid=chat] .chat-turn', (els) => els.length);
// Trạng thái nút "Hoàn tác lượt này" của một lượt: ẩn (đã hoàn tác), bấm được, hay khóa (kèm kiểu hiển thị).
const turnUndo = (id) => pg.$eval(`[data-testid=chat] .chat-turn[data-turn="${id}"]`, (el) => {
  const b = el.querySelector('[data-undo-turn]');
  const cs = b ? getComputedStyle(b) : null;
  return {
    undone: el.querySelector('[data-testid=chat-turn-undone]')?.textContent ?? null,
    button: b ? (b.disabled ? 'disabled' : 'enabled') : null,
    opacity: cs ? Number(cs.opacity) : null, cursor: cs?.cursor ?? null,
    text: el.innerText,
  };
});
const lastTurnId = () => pg.$eval('[data-testid=chat] .chat-turn:last-child', (el) => el.dataset.turn);
const lastTurn = () => pg.$eval('[data-testid=chat] .chat-turn:last-child', (el) => ({
  text: el.innerText,
  codes: [...el.querySelectorAll('[data-code]')].map((x) => x.dataset.code),
  summary: [...el.querySelectorAll('.chat-summary li')].map((x) => x.dataset.cell),
  skipped: [...el.querySelectorAll('.chat-skipped li')].map((x) => `${x.dataset.cell}:${x.dataset.code}`),
  model: el.querySelector('.chat-model')?.innerText ?? null,
  edited: el.querySelector('[data-edited]')?.dataset.edited ?? null,
  undo: el.querySelector('[data-undo-turn]') ? !el.querySelector('[data-undo-turn]').disabled : null,
}));

async function ready() {
  await pg.waitForFunction(() => window.vmGridReady, null, { timeout: 90000 });
  // R-1021: mỗi lần mở file, lưới phải thật sự được vẽ (có canvas lớn trong .grid-host), không chỉ có mô hình bảng.
  await pg.waitForFunction(() => [...document.querySelectorAll('.grid-host canvas')].some((c) => c.getBoundingClientRect().height > 100), null, { timeout: 15000 })
    .catch(() => { throw new Error('lưới trắng: .grid-host không có canvas sau khi mở file (R-1021)'); });
  await pg.evaluate(() => window.vmGrid.settled());
  await sleep(300);
}
async function open() {
  await pg.goto(BASE);
  await pg.waitForSelector('.welcome');
  await pg.evaluate(() => { window.vmGridReady = null; });
  await pg.setInputFiles('[data-testid=file-input]', FILE);
  await ready();
}
// Gõ và bấm Gửi. Trả 'preview' nếu màn xem trước hiện, 'turn' nếu đã có lượt mới.
async function ask(text) {
  const before = await turns();
  await pg.fill('[data-testid=chat-input]', text);
  await pg.click('[data-testid=chat-send]');
  await pg.waitForFunction((n) => document.querySelector('[data-testid=chat-preview]') || document.querySelectorAll('[data-testid=chat] .chat-turn').length > n,
    before, { timeout: 60000 });
  return (await pg.$('[data-testid=chat-preview]')) ? 'preview' : 'turn';
}
async function sendPreview() {
  const before = await turns();
  await pg.click('[data-testid=chat-send-preview]');
  await pg.waitForFunction((n) => document.querySelectorAll('[data-testid=chat] .chat-turn').length > n, before, { timeout: 60000 });
  await sleep(300);
}
// Gõ tay vào ô: bỏ con trỏ khỏi ô chat, bấm vào lưới để bàn phím về lưới, rồi chọn ô, gõ, Enter.
async function typeIn(ref, text) {
  await pg.evaluate(() => document.activeElement?.blur());
  const box = await pg.locator('.main canvas').first().boundingBox();
  await pg.mouse.click(box.x + box.width / 2, box.y + box.height / 2);
  await pg.evaluate(([s, r]) => window.vmGrid.goTo(s, r), [SHEET, ref]);
  await sleep(150);
  await pg.keyboard.type(text);
  await pg.keyboard.press('Enter');
  await sleep(500);
}
async function nameIfAsked() {
  await sleep(150);
  if (await pg.$('[data-testid=name-input]')) {
    await pg.fill('[data-testid=name-input]', 'Người Giả');
    await pg.click('.dialog-actions button:has-text("Tiếp tục")');
  }
}

try {
await open();
await pg.click('[data-testid=chat-button]');
await pg.waitForSelector('[data-testid=chat]');
res.legend_before = (await pg.$('.chat-legend')) !== null; // N5: chưa có ô tím thì không có chú thích
res.trial = (await pg.$('[data-testid=chat-trial]')) !== null;
const b10Before = await val('B10');
const bgBefore = await bg('B10');

// 2.1, 2.10-2.12: lượt đầu hiện xem trước, liệt kê sheet gửi kèm; không có sheet ẩn, sheet biểu đồ, tên nhân viên.
res.first = await ask('Glucose mức 1 ngày 3/9 là 5,62 FAKE:edit');
res.preview = {
  sheets: await pg.$$eval('[data-testid=chat-sheets] li', (els) => els.map((e) => e.dataset.sheet)),
  sheet_lines: await pg.$$eval('[data-testid=chat-sheets] li', (els) => els.map((e) => e.textContent)),
  prompt: await pg.textContent('[data-testid=chat-prompt]'),
  not_sent: await pg.textContent('[data-testid=chat-not-sent]').catch(() => null),
  hash: await pg.textContent('[data-testid=chat-hash]'),
};
await pg.screenshot({ path: path.join(OUT, 'chat_xem_truoc.png') });
// 2.3: Không gửi thì không có yêu cầu nào tới send, lưới không đổi.
await pg.click('[data-testid=chat-preview] button:has-text("Không gửi")');
await sleep(300);
res.not_sent = { sends: res.llm_requests.length, b10: await val('B10'), turns: await turns() };

// 3.1-3.2, 3.7: gửi lại, xem trước hiện lại (chưa đồng ý lần nào), rồi Gửi.
res.again = await ask('Glucose mức 1 ngày 3/9 là 5,62 FAKE:edit');
await sendPreview();
res.legend = await pg.textContent('.chat-legend').catch(() => null);
const editTurn = await lastTurnId();
res.undo_stacks_readable = await pg.evaluate(() => window.vmGrid.undoStacksReadable());
res.edit = { turn: await lastTurn(), b10: await val('B10'), bg: await bg('B10'), h10: await val('H10'), ai: await aiCells(), b10_before: b10Before, bg_before: bgBefore };
await pg.screenshot({ path: path.join(OUT, 'chat_da_sua.png') });

// 3.4: Hoàn tác lượt này bỏ cả lượt và dấu; Làm lại đưa lại.
await pg.click('[data-testid=chat] .chat-turn:last-child [data-undo-turn]');
await sleep(400);
res.undo = { b10: await val('B10'), bg: await bg('B10'), ai: await aiCells(), turn: await turnUndo(editTurn) };
await pg.evaluate(() => window.vmGrid.redo());
await sleep(400);
res.redo = { b10: await val('B10'), bg: await bg('B10'), ai: await aiCells(), undo_enabled: (await lastTurn()).undo, turn: await turnUndo(editTurn) };

// 2.5, 3.8: lượt hai cùng vùng không hiện xem trước; AI hỏi lại thì hiện câu của model.
res.second = await ask('Sửa giúp tôi');
await sleep(300);
res.ask = await lastTurn();

// Q135 (e2): model đề xuất nhưng cổng chặn hết: câu của app và lý do từng ô, KHÔNG có câu của model ("Đã ghi xong.").
const h10Before = await val('H10');
res.blocked_kind = await ask('Glucose mức 1 ngày 3/9 là 5,62 FAKE:all_blocked');
res.blocked = { turn: await lastTurn(), h10: await val('H10'), h10_before: h10Before };

// 2.8: câu có tên người: không tạo lượt gọi model.
const sendsBefore = res.llm_requests.length;
res.person_kind = await ask('Hôm nay Trần Giả Định chạy, Glucose mức 1 ngày 5/9 là 5,40');
res.person = { turn: await lastTurn(), sends: res.llm_requests.length - sendsBefore };

// 1.2-1.3: đang chờ thì không gõ được lượt hai; Hủy thì bảng không đổi và kết quả về muộn không được áp.
await pg.fill('[data-testid=chat-input]', 'Glucose mức 1 ngày 3/9 là 5,62 FAKE:slow');
await pg.click('[data-testid=chat-send]');
await pg.waitForSelector('[data-testid=chat-wait] span[data-phase=model]', { timeout: 30000 }); // N12: đã qua bước chuẩn bị, đang chờ AI
res.waiting_text = await pg.textContent('[data-testid=chat-wait] span');
res.waiting = { input_disabled: await pg.$eval('[data-testid=chat-input]', (e) => e.disabled), send_disabled: await pg.$eval('[data-testid=chat-send]', (e) => e.disabled) };
const cellsBeforeCancel = await aiCells();
await pg.click('[data-testid=chat-wait] button');
await sleep(8000); // llm hết hạn sau 5 giây; api trả về sau đó
res.cancel = { turn: await lastTurn(), ai_before: cellsBeforeCancel, ai_after: await aiCells(), wait: (await pg.$('[data-testid=chat-wait]')) !== null };

// 3.5: gõ tay vào ô AI đã sửa thì ô đó mất dấu.
await typeIn('B10', '5.7');
res.typed = { b10: await val('B10'), bg: await bg('B10'), ai: await aiCells(), turn: await turnUndo(editTurn) };
await pg.evaluate(() => window.vmGrid.undo());
await sleep(400);
res.typed_undo = { b10: await val('B10'), ai: await aiCells() };

// Q135 (h): I10 có ghi chú (chữ không gửi cho AI). AI ghi "chạy lại" vào I8, I9, I10: I8, I9 được ghi, I10 giữ nguyên, có câu báo.
// Ghi chú dài hơn 80 ký tự: vmai không gửi (ô [chữ]). Chữ ngắn được gửi như nhãn nên AI thấy được và được ghi đè có tóm tắt.
const NOTE = 'ghi chú giả dài của người dùng: mẫu đã chạy lại hai lần vì máy báo lỗi hút mẫu, chờ kiểm tra lại sau';
res.note = NOTE;
await typeIn('I10', NOTE);
// Q135 (h2): ghi chú ngắn (vmai gửi như nhãn, AI thấy được) cũng không được ghi đè.
res.short_note = 'ghi chú ngắn';
await typeIn('I11', res.short_note);
res.col_i_kind = await ask('Ghi chạy lại vào ghi chú ngày 1/9 đến 3/9 FAKE:col_i');
res.preview_col_i = res.col_i_kind === 'preview' ? await pg.textContent('[data-testid=chat-prompt]') : '';
if (res.col_i_kind === 'preview') await sendPreview();
res.col_i = { turn: await lastTurn(), i8: await val('I8'), i9: await val('I9'), i10: await val('I10'), i11: await val('I11'), ai: await aiCells() };

// 5.1-5.2: Lưu khi còn ô AI sửa: hỏi xác nhận; Xem lại thì không lưu; Lưu bản mới thì lịch sử ghi số ô.
await pg.click('.sidebar .side-btn:has-text("Lưu")');
await pg.waitForSelector('.dialog', { timeout: 10000 });
res.save_confirm = await pg.textContent('.dialog');
await pg.click('.dialog-actions button:has-text("Xem lại")');
await sleep(300);
res.after_review = await pg.textContent('[data-testid=doc-state]');
await pg.click('.sidebar .side-btn:has-text("Lưu")');
await pg.waitForSelector('.dialog');
await pg.evaluate(() => { window.vmGridReady = null; });
await pg.click('.dialog-actions button:has-text("Lưu bản mới")');
await nameIfAsked();
await ready();
res.legend_after_save = (await pg.$('.chat-legend')) !== null;
res.saved = { state: await pg.textContent('[data-testid=doc-state]'), b10: await val('B10'), bg: await bg('B10'), ai: await aiCells() };
await pg.click('.sidebar .side-btn:has-text("Phiên bản")');
await pg.waitForSelector('[data-testid=versions]');
res.versions = await pg.textContent('[data-testid=versions]');

} catch (e) {
  // Ghi lại những gì đã có và lượt cuối, để biết hỏng ở đâu; vẫn báo lỗi cho pytest.
  res.failed = String(e.message).slice(0, 300);
  res.last_turn = await lastTurn().catch(() => null);
  await pg.screenshot({ path: path.join(OUT, 'chat_loi.png') }).catch(() => {});
  process.exitCode = 1;
} finally {
  await browser.close();
  fs.writeFileSync(path.join(OUT, 'chat.json'), JSON.stringify(res, null, 1));
}
