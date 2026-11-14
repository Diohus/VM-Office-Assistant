// Khung chat lát 2 (Q135). API: docs/engineering/chat_api.md. Câu chữ: chat_texts trong doc_codes.json (AI Engineer, PO/BA duyệt).
// AI sửa thẳng trên lưới, ô có nền tím, cả lượt là một bước Hoàn tác, chỉ thành bản mới khi người dùng bấm Lưu.
import React, { useEffect, useRef, useState } from 'react';
import { ApiError, chatHealth, chatPreview, chatSend } from '../api.js';
import { formatNumber } from '../numberFormat.js';
import { SETUP_ERRORS, chatText, notSentLine, planHeadline, previewKey, rejectedLines, summaryText } from './chatModel.js';

const NETWORK_CODE = 'CHAT_MODEL_UNAVAILABLE';

function errorCode(e) {
  if (e instanceof ApiError && typeof e.code === 'string' && e.code.startsWith('CHAT_')) return e.code;
  return NETWORK_CODE;
}

function Turn({ turn, texts, grid, undoTick }) {
  const t = (code, vars) => chatText(texts, code, vars);
  // 'top': bấm được; 'covered': có sửa khác sau lượt này (nút khóa, câu giải thích); 'undone': đã hoàn tác (ẩn nút, PM 13/11).
  const state = turn.applied?.length ? (grid?.turnState?.(turn.turn, turn.sheet, turn.applied) ?? 'covered') : null;
  return (
    <li className="chat-turn" data-turn={turn.id}>
      {turn.message && <div className="chat-you">{turn.message}</div>}
      {turn.error && (
        <div className="chat-error" role="alert" data-code={turn.error}>
          {t(turn.error, { 'số ô': turn.count })}{SETUP_ERRORS.has(turn.error) ? ` ${t('CHAT_UI_ERROR_CODE', { 'mã': turn.error })}` : ''}
        </div>
      )}
      {turn.stopped && <div className="chat-info">{t('CHAT_UI_STOPPED')}</div>}
      {turn.headline && <div className="chat-info" data-code={turn.headlineCode}>{turn.headline}</div>}
      {turn.applied?.length > 0 && (
        <>
          <div className="chat-info" data-edited={turn.applied.length}>{t('CHAT_UI_EDITED', { N: turn.applied.length, sheet: turn.sheet })}</div>
          {/* 1.5 (O1): nút Hoàn tác đứng ngay dưới dòng đầu của lượt, không nằm sau tóm tắt dài (30 ô đẩy nút ra khỏi màn hình). */}
          <div className="chat-hint">{t('CHAT_UI_EDITED_HINT')}</div>
          {state === 'undone'
            ? <div className="chat-hint" data-testid="chat-turn-undone" data-tick={undoTick}>{t('CHAT_UI_TURN_UNDONE')}</div>
            : <button type="button" className="btn-line chat-undo" disabled={state !== 'top'} data-undo-turn={turn.id} data-tick={undoTick}
                onClick={() => grid?.undo()}>{t('CHAT_UI_UNDO_TURN')}</button>}
          {state === 'covered' && <div className="chat-hint">{t('CHAT_UI_UNDO_TURN_DISABLED')}</div>}
          {/* K4 (PM chọn 2): ô AI ghi ngoài các dòng dữ liệu của bảng (dòng tổng, dòng thống kê) vẫn được áp, nhưng tách riêng ngay sau dòng "AI đã sửa" và nút Hoàn tác, trước danh sách ô (PO/BA 4.16;
              nút Hoàn tác đứng trước nhóm này để nhóm dài không đẩy nút ra sau ô gõ), kèm câu cảnh báo để người dùng xem kỹ. Cờ outside_data_rows do vmai đặt cho từng dòng tóm tắt. */}
          {turn.summary.some((s) => s.outside) && (() => {
            const out = turn.summary.filter((s) => s.outside);
            return (
              <div className="chat-outside" data-testid="chat-outside">
                <div className="chat-warn">{t('CHAT_UI_OUTSIDE_BODY', { N: out.length, 'danh sách ô': out.map((s) => (s.rowLabel ? `${s.cell} (${s.rowLabel})` : s.cell)).join(', ') })}</div>
                <ul className="chat-summary">
                  {out.map((s) => <li key={s.cell} data-cell={s.cell} data-outside="1">{s.text}</li>)}
                </ul>
              </div>
            );
          })()}
          <ul className="chat-summary">
            {turn.summary.filter((s) => !s.outside).map((s) => <li key={s.cell} data-cell={s.cell}>{s.text}</li>)}
          </ul>
        </>
      )}
      {turn.skipped?.length > 0 && (
        <>
          {turn.applied?.length > 0 && <div className="chat-info">{t('CHAT_UI_PARTIAL', { K: turn.applied.length, M: turn.skipped.length })}</div>}
          <ul className="chat-skipped">
            {turn.skipped.map((s) => <li key={`${s.cell}-${s.code}`} data-code={s.code} data-cell={s.cell}>{s.text}</li>)}
          </ul>
        </>
      )}
      {/* N7: ô đã đúng như yêu cầu không phải "không áp", nêu riêng sau cùng. */}
      {turn.already?.length > 0 && (
        <ul className="chat-already" data-testid="chat-already">
          {turn.already.map((s) => <li key={s.cell} data-code={s.code} data-cell={s.cell}>{s.text}</li>)}
        </ul>
      )}
      {turn.modelSaid && (
        <div className="chat-model"><span>{t('CHAT_UI_MODEL_SAID')}</span> {turn.modelSaid}</div>
      )}
    </li>
  );
}

export default function ChatPanel({ doc, gridRef, texts, undoTick, onClose, onOpenFill, onBusy }) {
  // Q135 (g): sheet có biểu đồ không gửi kèm.
  const chartSheets = [...new Set((doc.unshown ?? []).filter((u) => u.kind === 'chart' && u.sheet).map((u) => u.sheet))];
  const t = (code, vars) => chatText(texts, code, vars);
  const [message, setMessage] = useState('');
  const [turns, setTurns] = useState([]);
  const [preview, setPreview] = useState(null); // { data, snapshot, key, message }
  const [waiting, setWaiting] = useState(false); // false | 'prepare' (chụp vùng, chờ xem trước) | 'model' (đã gửi AI)
  const [trial, setTrial] = useState(false);
  const accepted = useRef(null); // khóa xem trước người dùng đã đồng ý gửi trong phiên này
  const abort = useRef(null);
  const listRef = useRef(null);

  useEffect(() => {
    chatHealth().then((h) => setTrial(Boolean(h.chat_test_settings))).catch(() => {});
  }, []);
  // Mở tài liệu khác: lượt cũ không còn thuộc về lưới đang hiện.
  useEffect(() => {
    setTurns([]);
    setPreview(null);
    accepted.current = null;
  }, [doc.document_id]);
  useEffect(() => {
    // Lượt mới hiện từ đầu lượt (câu đã gõ, dòng "AI đã sửa", nút Hoàn tác); tóm tắt dài cuộn tiếp bên dưới. Chỉ cuộn danh sách lượt.
    const list = listRef.current, last = list?.lastElementChild;
    if (list && last) list.scrollTop = last.offsetTop - list.offsetTop;
  }, [turns]);
  useEffect(() => onBusy?.(Boolean(waiting)), [waiting, onBusy]);

  const hasMarks = (gridRef.current?.aiCells?.() ?? []).length > 0;
  const addTurn = (turn) => setTurns((list) => [...list, { id: list.length + 1, ...turn }]);

  const send = async (p) => {
    const ctrl = new AbortController();
    abort.current = ctrl;
    setWaiting('model');
    try {
      const res = await chatSend(p.data.preview_id, ctrl.signal);
      if (ctrl.signal.aborted) return; // đã Hủy: kết quả về muộn không được áp (1.3)
      if (res.chat_test_settings) setTrial(true);
      const plan = res.plan;
      const grid = gridRef.current;
      const turn = { message: p.message, sheet: plan.active_sheet };
      const skipped = rejectedLines(texts, plan);
      if (plan.status === 'edit' && plan.ops.length) {
        const done = grid.applyChatOps(plan.active_sheet, plan.ops, p.snapshot.sent);
        turn.applied = done.applied;
        turn.turn = done.turn;
        // Tóm tắt do app viết: chữ cũ thật từ bản chụp, số theo dấu thập phân của app (không dùng summary.text của vmai).
        const lines = new Map((plan.summary ?? []).map((s) => [s.cell, s]));
        turn.summary = plan.ops.filter((op) => done.applied.includes(op.cell))
          .map((op) => ({ cell: op.cell, outside: Boolean(lines.get(op.cell)?.outside_data_rows), rowLabel: lines.get(op.cell)?.row_label ?? null,
            text: summaryText(texts, lines.get(op.cell), op, p.snapshot.sent[op.cell] ?? null, formatNumber) }));
        for (const cell of done.changed) skipped.push({ cell, code: 'CHAT_UI_CELL_CHANGED', text: t('CHAT_UI_CELL_CHANGED', { 'ô': cell }) });
        for (const cell of done.locked) skipped.push({ cell, code: 'CHAT_UI_CELL_LOCKED', text: t('CHAT_UI_CELL_LOCKED', { 'ô': cell }) });
        for (const cell of done.hasText) skipped.push({ cell, code: 'CHAT_CELL_HAS_TEXT', text: t('CHAT_CELL_HAS_TEXT', { 'ô': cell }) });
        if (!done.applied.length) {
          const stale = done.changed.length === plan.ops.length;
          turn.headlineCode = stale ? 'CHAT_UI_STALE' : 'CHAT_UI_EDITED_NONE';
          turn.headline = t(turn.headlineCode);
        }
      } else {
        turn.headline = planHeadline(texts, plan) ?? t('CHAT_UI_EDITED_NONE');
        // N4: biểu mẫu chưa điền được bằng Điền từ file (lát 1 chỉ biểu mẫu QC mẫu) thì app không mở Điền từ file, chỉ hiện băng
        // báo. Câu không được nói "App mở Điền từ file": dùng câu chung "chưa làm được"; băng của app nêu lý do.
        if (plan.status === 'unsupported' && plan.unsupported === 'fill_from_file' && !doc.fill_supported) turn.headline = t('CHAT_UNSUPPORTED');
        turn.headlineCode = plan.status === 'unsupported' ? `CHAT_UNSUPPORTED_${String(plan.unsupported ?? '').toUpperCase()}` : plan.reason ?? plan.status;
        if (plan.status === 'unsupported' && plan.unsupported === 'fill_from_file') onOpenFill?.();
      }
      // N7: ô đã đúng sẵn không tính vào "không áp".
      turn.already = skipped.filter((s) => s.code === 'CHAT_NO_CHANGE');
      turn.skipped = skipped.filter((s) => s.code !== 'CHAT_NO_CHANGE');
      // Q135 (e2): câu của model chỉ hiện khi model KHÔNG đề xuất thao tác nào (hỏi lại, từ chối, chưa làm được). Model có đề
      // xuất mà bị chặn hết thì chỉ có câu của app và lý do từng ô: model có thể nói "đã ghi" trong khi bảng không đổi.
      const proposedNothing = !(plan.ops?.length) && !(plan.rejected_ops?.length);
      if (proposedNothing && ['ask', 'refuse', 'unsupported'].includes(plan.status) && plan.model_message) turn.modelSaid = plan.model_message;
      addTurn(turn);
    } catch (e) {
      if (ctrl.signal.aborted) return;
      addTurn({ message: p.message, error: errorCode(e) });
      // 6.5 (PM thay cho nút "Gửi lại"): lượt hỏng ở bước gửi (mạng, hết giờ, trần...) trả nguyên câu về ô gõ để bấm Gửi lại.
      // Lỗi ở bước xem trước (câu có tên, quá 3.000 ô) vốn đã giữ câu, vì ô gõ chỉ bị xóa khi xem trước thành công.
      setMessage((current) => current || p.message);
    } finally {
      if (abort.current === ctrl) abort.current = null;
      setWaiting(false);
    }
  };

  const ask = async () => {
    const grid = gridRef.current;
    if (!grid || waiting || preview) return;
    if (!(await grid.commitEditing())) return;
    const text = message;
    // Hủy trong lúc chờ tính hay chờ bản xem trước cũng dừng cả lượt: không gửi tiếp (1.3).
    const ctrl = new AbortController();
    abort.current = ctrl;
    setWaiting('prepare');
    // O2: chụp vùng sau khi bộ tính xong (có hạn chờ), để ô công thức mới gõ hoặc do lượt trước ghi có kết quả thật.
    await grid.settled?.();
    if (ctrl.signal.aborted) return;
    const snapshot = grid.chatSnapshot(chartSheets);
    if (snapshot.tooLarge) {
      abort.current = null;
      setWaiting(false);
      addTurn({ message: text, error: 'CHAT_REGION_TOO_LARGE', count: snapshot.tooLarge.toLocaleString('vi-VN') });
      return;
    }
    const body = { ...snapshot.request, base_version: doc.version, message: text };
    let data;
    try {
      data = await chatPreview(doc.document_id, body);
    } catch (e) {
      if (!ctrl.signal.aborted) addTurn({ message: text, error: errorCode(e), count: e.detail?.items?.[0] });
      if (abort.current === ctrl) abort.current = null;
      setWaiting(false);
      return;
    }
    if (abort.current === ctrl) abort.current = null;
    setWaiting(false);
    if (ctrl.signal.aborted) return;
    if (data.chat_test_settings) setTrial(true);
    setMessage('');
    const p = { data, snapshot, message: text, key: previewKey(body, data) };
    // Xem trước ở lượt đầu và mỗi khi vùng hay ô bị che đổi (2.5, 2.6).
    if (accepted.current === p.key) send(p);
    else setPreview(p);
  };

  const stop = () => {
    abort.current?.abort();
    abort.current = null;
    setWaiting(false);
    addTurn({ message: '', stopped: true });
  };

  return (
    <aside className="chat" aria-label="Chat AI" data-testid="chat">
      <div className="chat-head">
        <h2>Chat AI</h2>
        <button type="button" className="btn-plain" onClick={onClose}>Đóng</button>
      </div>
      {trial && <div className="chat-trial" data-testid="chat-trial">{t('CHAT_UI_TRIAL_LABEL')}</div>}
      {/* N5: chú thích chỉ hiện khi lưới đang có ô tím (vẽ lại theo undoTick: Hoàn tác, Làm lại, sửa, Lưu). */}
      {hasMarks && <div className="chat-legend" data-testid="chat-legend"><span className="chat-swatch" aria-hidden="true" />{t('CHAT_UI_LEGEND')}</div>}
      <ul className="chat-turns" ref={listRef}>
        {turns.map((turn) => <Turn key={turn.id} turn={turn} texts={texts} grid={gridRef.current} undoTick={undoTick} />)}
      </ul>
      {preview && (
        // O1: nội dung xem trước cuộn trong khung; mã băm và hàng nút nằm dưới, không bao giờ bị đẩy ra ngoài hay bị phần khác đè.
        <div className="chat-preview" data-testid="chat-preview" role="dialog" aria-label={t('CHAT_UI_PREVIEW_TITLE')}>
          <div className="chat-preview-body" data-testid="chat-preview-body">
            <div className="chat-info">{t('CHAT_UI_PREVIEW_TITLE')}</div>
            <ul className="chat-sheets" data-testid="chat-sheets">
              {preview.snapshot.request.sheets.map((s) => <li key={s.name} data-sheet={s.name}>{t('CHAT_UI_PREVIEW_SHEET', { sheet: s.name, 'danh sách cột': s.region })}</li>)}
            </ul>
            {/* 0.2 (N2): phần gửi kèm chỉ để đọc (tiêu đề, cột Ngày) và vùng AI được sửa. */}
            {preview.snapshot.request.sheets.filter((s) => s.context_region).map((s) => (
              <div key={`ctx-${s.name}`} className="chat-info" data-testid="chat-preview-context">{t('CHAT_UI_PREVIEW_CONTEXT', { 'vùng': s.context_region, 'vùng sửa': s.region })}</div>
            ))}
            <pre className="chat-prompt" data-testid="chat-prompt">{preview.data.prompt}</pre>
            <div className="chat-info" data-testid="chat-not-sent">{notSentLine(texts, preview.data.masked)}</div>
          </div>
          <div className="chat-hint" data-testid="chat-hash">{preview.data.prompt_sha256.slice(0, 12)}</div>
          <div className="chat-actions" data-testid="chat-preview-actions">
            <button type="button" className="btn" data-testid="chat-send-preview" onClick={() => {
              accepted.current = preview.key;
              const p = preview;
              setPreview(null);
              send(p);
            }}>{t('CHAT_UI_PREVIEW_SEND')}</button>
            <button type="button" className="btn-line" data-testid="chat-cancel-preview" onClick={() => {
              setMessage(preview.message); // Không gửi: trả lại câu đã gõ để sửa, không phải gõ lại
              setPreview(null);
            }}>{t('CHAT_UI_PREVIEW_CANCEL')}</button>
          </div>
        </div>
      )}
      {waiting && (
        <div className="chat-wait" data-testid="chat-wait">
          <span data-phase={waiting}>{t(waiting === 'prepare' ? 'CHAT_UI_PREPARING' : 'CHAT_UI_WAITING')}</span>{/* N12: lúc chuẩn bị chưa gửi gì cho AI */}
          <button type="button" className="btn-line" onClick={stop}>{t('CHAT_UI_STOP')}</button>
        </div>
      )}
      {/* 1.6 (O1): đang xem trước thì không có ô gõ và nút Gửi chính (trước đây hiện mờ, trông như bấm được). Bước tiếp chỉ
          còn hai nút của khung xem trước. */}
      {!preview && <div className="chat-input">
        <textarea className="text-input" value={message} onChange={(e) => setMessage(e.target.value)} placeholder={t('CHAT_UI_INPUT_HINT')}
          aria-label="Yêu cầu cho AI" data-testid="chat-input" rows={3} maxLength={1200} disabled={waiting}
          onKeyDown={(e) => {
            e.stopPropagation();
            if (e.key === 'Enter' && (e.ctrlKey || e.metaKey)) {
              e.preventDefault();
              ask();
            }
          }} />
        <button type="button" className="btn" onClick={ask} disabled={waiting} data-testid="chat-send">Gửi</button>
      </div>}
    </aside>
  );
}
