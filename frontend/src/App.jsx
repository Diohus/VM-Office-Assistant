import React, { useCallback, useEffect, useRef, useState } from 'react';
import { ApiError, exportUrl, fetchGrid, uploadDocument } from './api.js';
import { createGrid } from './grid/index.js';
import { documentNotices, mismatchNotice } from './notices.js';

const ACCEPT = '.xlsx,.xlsm,.xls,application/vnd.openxmlformats-officedocument.spreadsheetml.sheet';

function Icon({ d }) {
  return <svg viewBox="0 0 24 24" aria-hidden="true"><path d={d} /></svg>;
}

function NoticeBar({ notice, onClose, onGo }) {
  const [open, setOpen] = useState(false);
  return (
    <div className={`notice ${notice.kind}`} role={notice.kind === 'error' ? 'alert' : 'status'} data-notice={notice.id}>
      <div className="notice-text">
        {notice.text}
        {notice.places?.length > 0 && (
          <button type="button" className="btn-plain" onClick={() => setOpen(!open)}>Xem vị trí</button>
        )}
        {open && (
          <ul className="place-list">
            {notice.places.map((p, i) => (
              <li key={i}>
                <button type="button" className="btn-plain" onClick={() => onGo(p)}>{p.label}</button>
              </li>
            ))}
          </ul>
        )}
      </div>
      <button type="button" className="btn-plain" onClick={onClose}>Đóng</button>
    </div>
  );
}

export default function App() {
  const [doc, setDoc] = useState(null);
  const [busy, setBusy] = useState(false);
  const [error, setError] = useState(null);
  const [notices, setNotices] = useState([]);
  const [closed, setClosed] = useState(new Set());
  const inputRef = useRef(null);
  const hostRef = useRef(null);
  const gridRef = useRef(null);

  useEffect(() => () => gridRef.current?.destroy(), []);

  const openFile = useCallback(async (file) => {
    setBusy(true);
    setError(null);
    try {
      const meta = await uploadDocument(file);
      const model = await fetchGrid(meta.document_id);
      gridRef.current?.destroy();
      gridRef.current = null;
      setDoc(meta);
      setClosed(new Set());
      setNotices(documentNotices(meta));
      // Chờ React dựng khung lưới rồi mới gắn lưới vào.
      await new Promise((r) => requestAnimationFrame(r));
      const grid = createGrid(hostRef.current, model);
      gridRef.current = grid;
      window.vmGrid = grid;
      const mismatches = await grid.compareWithSaved();
      const note = mismatchNotice(mismatches);
      if (note) setNotices((list) => [...list, note]);
      window.vmGridReady = meta.document_id;
    } catch (e) {
      setError(e instanceof ApiError ? e.message : 'App chưa mở được file này. File của bạn không bị đổi. Hãy thử lại, hoặc hỏi người quản lý app nếu vẫn không được.');
    } finally {
      setBusy(false);
    }
  }, []);

  const pick = () => inputRef.current?.click();
  const onChosen = (e) => {
    const file = e.target.files?.[0];
    e.target.value = '';
    if (file) openFile(file);
  };
  const go = (place) => gridRef.current?.goTo(place.sheet, place.cell);
  const visible = notices.filter((n) => !closed.has(n.id));
  const hiddenCount = notices.length - visible.length;

  return (
    <div className="app">
      <header className="topbar">
        <div className="logo" title="Logo" aria-hidden="true" />
        <div className="doc-name" data-testid="doc-name">{doc ? doc.name : 'VM-Office-Assistant'}</div>
        {doc && <div className="doc-state">Đang xem, chưa sửa được ở bản này</div>}
      </header>
      {busy && <div className="progress" role="progressbar" aria-label="Đang mở file" />}
      <div className="body">
        <nav className="sidebar" aria-label="Chức năng">
          <button type="button" className="side-btn" onClick={pick}>
            <Icon d="M4 19V6h6l2 2h8v11H4z" />Mở
          </button>
          {doc && (
            <a className="side-btn" href={exportUrl(doc.document_id)} download style={{ textDecoration: 'none' }}>
              <Icon d="M12 4v11m0 0l-4-4m4 4l4-4M5 20h14" />Xuất
            </a>
          )}
        </nav>
        <main className="main">
          <input ref={inputRef} type="file" accept={ACCEPT} hidden onChange={onChosen} data-testid="file-input" />
          {error && (
            <div className="notices">
              <NoticeBar notice={{ id: 'open-error', kind: 'error', text: error }} onClose={() => setError(null)} onGo={go} />
            </div>
          )}
          {doc && (
            <div className="notices">
              {visible.map((n) => (
                <NoticeBar key={`${doc.document_id}-${n.id}`} notice={n} onGo={go} onClose={() => setClosed(new Set([...closed, n.id]))} />
              ))}
              {hiddenCount > 0 && (
                <div className="reopen">
                  <button type="button" className="btn-plain" onClick={() => setClosed(new Set())}>Xem lại {hiddenCount} thông báo</button>
                </div>
              )}
            </div>
          )}
          {doc ? (
            <div className="grid-host" ref={hostRef} data-testid="grid" />
          ) : (
            <div className="welcome">
              <h1>Mở biểu mẫu để bắt đầu</h1>
              <p>App mở được file .xlsx. File gốc của bạn không bị đổi.</p>
              <button type="button" className="btn" onClick={pick} disabled={busy}>Mở file</button>
            </div>
          )}
        </main>
      </div>
    </div>
  );
}
