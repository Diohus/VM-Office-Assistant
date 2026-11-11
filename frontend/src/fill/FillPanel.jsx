// Khung duyệt lát 1 (Q133, docs/engineering/lat1_plan.md mục 4). Đề xuất chỉ nằm ở đây, không vào lưới cho tới khi Ghi (4.1).
import React, { useEffect, useMemo, useRef, useState } from 'react';
import { buildView, rejectRow, toggleRow, typedChoice, writeCount } from './fillModel.js';

const COLOR_LABEL = { yellow: 'Vàng', red: 'Đỏ' };
const TYPED_PROBLEM = {
  ambiguous: 'Số này có thể hiểu hai cách (số nghìn hoặc số lẻ) nên app chưa nhận. Hãy gõ không có dấu nghìn, dùng dấu chấm cho số lẻ.',
  unclear: 'Đây không phải một số rõ ràng nên app chưa nhận. Hãy gõ lại, dùng dấu chấm cho số lẻ.',
  other: 'Hãy gõ một số, dùng dấu chấm cho số lẻ.',
};

function Reasons({ cell }) {
  return (
    <ul className="fill-reasons">
      {cell.reasons.map((r) => <li key={r.code} data-code={r.code}>{r.text}</li>)}
    </ul>
  );
}

function RedCell({ cell, choice, setChoice }) {
  const [draft, setDraft] = useState('');
  const [problem, setProblem] = useState(null);
  const pick = (c) => {
    setProblem(null);
    setChoice(cell.id, c);
  };
  const typed = () => {
    const r = typedChoice(draft);
    if (r.ok) pick(r.choice);
    else setProblem(TYPED_PROBLEM[r.kind]);
  };
  return (
    <div className="fill-red" data-red={cell.id}>
      {cell.readings.length > 1 && cell.readings.map((r) => (
        <button key={r.index} type="button" className={`btn-line ${choice?.action === 'red_choice' && choice.index === r.index ? 'on' : ''}`}
          onClick={() => pick({ action: 'red_choice', index: r.index })}>Chọn lần đọc {r.index + 1}: {r.text}</button>
      ))}
      {cell.value != null && (
        <button type="button" className={`btn-line ${choice?.action === 'red_accept' ? 'on' : ''}`} onClick={() => pick({ action: 'red_accept' })}>Nhận số này: {cell.value}</button>
      )}
      <span className="fill-typed">
        <input className="text-input" aria-label={`Gõ số khác cho ô ${cell.a1}`} value={draft} onChange={(e) => setDraft(e.target.value)} maxLength={30}
          onKeyDown={(e) => { if (e.key === 'Enter') { e.preventDefault(); e.stopPropagation(); typed(); } }} />
        <button type="button" className={`btn-line ${choice?.action === 'red_typed' ? 'on' : ''}`} onClick={typed}>Nhận số gõ</button>
      </span>
      <button type="button" className={`btn-line ${choice?.action === 'red_blank' ? 'on' : ''}`} onClick={() => pick({ action: 'red_blank' })}>Để trống</button>
      {choice && <button type="button" className="btn-plain" onClick={() => { setProblem(null); setChoice(cell.id, null); }}>Bỏ chọn</button>}
      {problem && <div className="fill-problem" role="alert">{problem}</div>}
    </div>
  );
}

export default function FillPanel({ fill, texts, busy, reviewAgain, onApply, onClose, onGo, onDrop }) {
  const view = useMemo(() => buildView(fill, texts), [fill, texts]);
  const [choices, setChoices] = useState({});
  const [sel, setSel] = useState(0);
  const listRef = useRef(null);

  // Lô mới hay vừa ghi: bỏ lựa chọn cũ, chỉ giữ các ô cần duyệt lại.
  useEffect(() => {
    setChoices({});
    setSel(0);
  }, [fill]);

  // Ô bị đổi từ chỗ khác (3.8a): duyệt lại từng ô, số hiện tại → số mới; nhận lại thì gửi số đã thấy.
  const again = Object.keys(reviewAgain ?? {});
  const rows = view.rows.map((r) => ({ ...r, cells: r.cells.filter((c) => !again.includes(c.id)) })).filter((r) => r.cells.length);
  const fromRows = view.rows.flatMap((r) => r.cells.filter((c) => again.includes(c.id)).map((c) => ({ ...c, row: r.key })));
  const overwrites = [
    ...fromRows.map((c) => ({ ...c, previous: reviewAgain[c.id] ?? '', seen: reviewAgain[c.id] == null ? '' : String(reviewAgain[c.id]) })),
    ...view.overwrites.map((c) => (again.includes(c.id) ? { ...c, previous: reviewAgain[c.id] ?? '', seen: reviewAgain[c.id] == null ? '' : String(reviewAgain[c.id]) } : c)),
  ];

  const setChoice = (id, c) => setChoices((prev) => {
    const next = { ...prev };
    if (c) next[id] = c;
    else delete next[id];
    return next;
  });
  const n = writeCount(choices);

  // Bàn phím chỉ khi tiêu điểm ở danh sách dòng (3.5a): Enter trong lưới vẫn là của lưới.
  const onKeyDown = (e) => {
    if (e.target !== listRef.current) return;
    if (e.key === 'ArrowDown') { e.preventDefault(); setSel((i) => Math.min(i + 1, rows.length - 1)); return; }
    if (e.key === 'ArrowUp') { e.preventDefault(); setSel((i) => Math.max(i - 1, 0)); return; }
    if (e.key !== 'Enter') return;
    e.preventDefault();
    // Không có phím duyệt nhiều dòng (3.4): bỏ Ctrl, Shift, Alt. Giữ Enter chỉ nhận một dòng (3.5).
    if (e.repeat || e.ctrlKey || e.shiftKey || e.altKey || e.metaKey) return;
    const row = rows[sel];
    if (!row) return;
    setChoices((prev) => {
      const yellow = row.cells.filter((c) => c.color === 'yellow' && !c.overwrite);
      return yellow.every((c) => prev[c.id]?.action === 'accept') ? prev : toggleRow(row, prev);
    });
    setSel((i) => Math.min(i + 1, rows.length - 1));
  };

  return (
    <aside className="fill-panel" aria-label="Điền từ file" data-testid="fill-panel">
      <div className="fill-head">
        <h2>Điền từ file</h2>
        <button type="button" className="btn-plain" onClick={onClose}>Đóng</button>
      </div>
      <div className="fill-source" data-testid="fill-source">File nguồn: <b>{view.sourceName}</b></div>
      {view.empty ? (
        // N8: lô 0 ô kết thúc ngay ở máy chủ; khung chỉ nói lý do, không nói "các số dưới đây".
        <p className="fill-header" data-testid="fill-empty">Không có ô nào điền được từ file {view.sourceName}. Biểu mẫu không bị đổi. Lý do ở dưới.</p>
      ) : (
        <>
          <p className="fill-header" data-testid="fill-header">{view.header}</p>
          <p className="fill-counts" data-testid="fill-counts">
            Còn chờ: {view.counts.yellow} ô vàng, {view.counts.red} ô đỏ.
          </p>
        </>
      )}
      {view.common.length > 0 && (
        <div className="fill-common" data-testid="fill-common">
          Áp dụng cho mọi ô vàng:
          <ul className="fill-reasons">{view.common.map((r) => <li key={r.code} data-code={r.code}>{r.text}</li>)}</ul>
        </div>
      )}
      {view.notices.length > 0 && (
        <ul className="fill-notices" data-testid="fill-notices">
          {view.notices.map((x) => <li key={x.code} data-code={x.code}>{x.text} {x.action}</li>)}
        </ul>
      )}
      <details className="fill-mapping" data-testid="fill-mapping" open>
        <summary>Cột nguồn → cột biểu mẫu</summary>
        <ul>
          {view.mapping.map((m) => <li key={m.target}>{m.source || '(không rõ)'} → {m.target}: {m.samples.join(', ')}</li>)}
        </ul>
      </details>

      {overwrites.length > 0 && (
        <section className="fill-overwrites" data-testid="fill-overwrites">
          <h3>Ô đã có số: duyệt từng ô</h3>
          <ul>
            {overwrites.map((c) => {
              const on = choices[c.id];
              return (
                <li key={c.id} data-change={c.id}>
                  <button type="button" className="btn-plain" onClick={() => onGo(c.a1)}>{c.row} · {c.col} ({c.a1})</button>
                  <span className="fill-old-new">{c.previous === '' || c.previous == null ? '(trống)' : String(c.previous)} → {c.value ?? '?'}</span>
                  {c.color === 'red' ? <RedCell cell={c} choice={on} setChoice={(id, x) => setChoice(id, x && c.seen !== undefined ? { ...x, seen: c.seen } : x)} /> : (
                    <span>
                      <button type="button" className={`btn-line ${on?.action === 'accept' ? 'on' : ''}`} onClick={() => setChoice(c.id, on?.action === 'accept' ? null : { action: 'accept', ...(c.seen !== undefined ? { seen: c.seen } : {}) })}>Nhận ô</button>
                      <button type="button" className={`btn-line ${on?.action === 'reject' ? 'on' : ''}`} onClick={() => setChoice(c.id, on?.action === 'reject' ? null : { action: 'reject' })}>Bỏ ô</button>
                    </span>
                  )}
                  <Reasons cell={c} />
                </li>
              );
            })}
          </ul>
        </section>
      )}

      <div className="fill-rows" ref={listRef} tabIndex={0} onKeyDown={onKeyDown} data-testid="fill-rows" aria-label="Danh sách đề xuất theo dòng. Mũi tên lên xuống để chọn dòng, Enter để nhận dòng.">
        {rows.map((r, i) => (
          <div key={r.key} className={`fill-row ${i === sel ? 'sel' : ''}`} data-row={r.key}>
            <div className="fill-row-head">
              <b>{r.key}</b>
              <button type="button" className="btn-line" onClick={() => { setSel(i); setChoices((prev) => toggleRow(r, prev)); }}>Nhận dòng</button>
              <button type="button" className="btn-line" onClick={() => { setSel(i); setChoices((prev) => rejectRow(r, prev)); }}>Bỏ dòng</button>
            </div>
            {r.cells.map((c) => {
              const on = choices[c.id];
              return (
                <div key={c.id} className={`fill-cell ${c.color}`} data-change={c.id} data-color={c.color}>
                  <button type="button" className="btn-plain" onClick={() => onGo(c.a1)}>{c.col} ({c.a1})</button>
                  <span className="fill-color">{COLOR_LABEL[c.color]}</span>
                  <span className="fill-value">Sẽ ghi: {c.value ?? '(chưa có)'}</span>
                  <span className="fill-src">Trong nguồn: {c.source.join(' / ')}</span>
                  {c.color === 'yellow' && (
                    <span className="fill-state" data-state={on?.action ?? 'pending'}>{on?.action === 'accept' ? 'Đã nhận' : on?.action === 'reject' ? 'Đã bỏ' : 'Chưa duyệt'}</span>
                  )}
                  {c.color === 'red' && <RedCell cell={c} choice={on} setChoice={setChoice} />}
                  <Reasons cell={c} />
                </div>
              );
            })}
          </div>
        ))}
      </div>

      {!view.empty && (
        <div className="fill-foot">
          <button type="button" className="btn" data-testid="fill-apply" disabled={busy || Object.keys(choices).length === 0} onClick={() => onApply(choices)}>
            Ghi {n} ô thành bản mới
          </button>
          <div className="fill-note">Ghi là tạo bản mới ngay, không cần bấm Lưu. Ô chưa duyệt vẫn chờ ở đây.</div>
          {/* N5: câu D10 bảo bỏ lần điền đang mở; đây là chỗ làm việc đó. Hỏi lại trước khi bỏ. */}
          <button type="button" className="btn-plain fill-drop" data-testid="fill-drop" disabled={busy} onClick={() => onDrop(view.pendingCount)}>
            Bỏ lần điền này
          </button>
        </div>
      )}
    </aside>
  );
}
