import React, { useCallback, useEffect, useRef, useState } from 'react';
import { ApiError, EDITS_TOO_MANY, MAX_EDITS, exportPdf, exportUrl, fetchDocument, fetchGrid, fetchRecent, fetchVersions, restoreVersion, saveVersion, uploadDocument } from './api.js';
import { createGrid } from './grid/index.js';
import { ambiguousNumberNotice, decimalSwappedNotice, documentNotices, hiddenDigitsNotice, EXPORT_UNCHANGED, EXPORT_UNCHANGED_LEGACY, EXPORT_WARNING, lockedCellNotice, mismatchNotice, pastedNotice, pdfMismatchNotice, refusedNotice, SAVE_FAILED_NOTE, savedNotice } from './notices.js';

const ACCEPT = '.xlsx,.xlsm,.xls,application/vnd.openxmlformats-officedocument.spreadsheetml.sheet';
const USER_KEY = 'vm_user_name';

function readUser() {
  try {
    return localStorage.getItem(USER_KEY) || '';
  } catch {
    return '';
  }
}

function rememberUser(name) {
  try {
    localStorage.setItem(USER_KEY, name);
  } catch {
    // Trình duyệt chặn bộ nhớ thì lần sau hỏi lại tên, không ảnh hưởng dữ liệu.
  }
}

function Icon({ d }) {
  return <svg viewBox="0 0 24 24" aria-hidden="true"><path d={d} /></svg>;
}

// Nút trên thông báo không lấy bàn phím khi bấm chuột: bấm Đóng hay Xem vị trí rồi gõ tiếp thì phím vẫn vào lưới (Q125).
// Trước đây bàn phím rơi ra ngoài trang, phím gõ ngay sau đó bị mất.
const keepFocus = (e) => e.preventDefault();

function NoticeBar({ notice, onClose, onGo }) {
  const [open, setOpen] = useState(false);
  return (
    <div className={`notice ${notice.kind}`} role={notice.kind === 'error' ? 'alert' : 'status'} data-notice={notice.id}>
      <div className="notice-text">
        {notice.text}
        {notice.places?.length > 0 && (
          <button type="button" className="btn-plain" onMouseDown={keepFocus} onClick={() => setOpen(!open)}>Xem vị trí</button>
        )}
        {open && (
          <ul className="place-list">
            {notice.places.map((p, i) => (
              <li key={i}>
                <button type="button" className="btn-plain" onMouseDown={keepFocus} onClick={() => onGo(p)}>{p.label}</button>
              </li>
            ))}
          </ul>
        )}
      </div>
      <button type="button" className="btn-plain" onMouseDown={keepFocus} onClick={onClose}>Đóng</button>
    </div>
  );
}

function Dialog({ title, children, actions }) {
  // Đóng hộp thì trả bàn phím về chỗ cũ, thường là lưới, để gõ tiếp được ngay.
  useEffect(() => {
    const before = document.activeElement;
    return () => before?.focus?.();
  }, []);
  return (
    <div className="dialog-back" role="dialog" aria-modal="true" aria-label={title} data-dialog={title}>
      <div className="dialog">
        <h2>{title}</h2>
        <div className="dialog-body">{children}</div>
        <div className="dialog-actions">
          {actions.map((a) => (
            <button key={a.label} type="button" className={a.primary ? 'btn' : 'btn-line'} onClick={a.onClick} disabled={a.disabled}>{a.label}</button>
          ))}
        </div>
      </div>
    </div>
  );
}

export default function App() {
  const [doc, setDoc] = useState(null);
  const [viewing, setViewing] = useState(null); // số bản cũ đang xem, null là đang ở bản mới nhất
  const [busy, setBusy] = useState(false);
  const [error, setError] = useState(null);
  const [notices, setNotices] = useState([]);
  const [closed, setClosed] = useState(new Set());
  const [versions, setVersions] = useState(null);
  const [dialog, setDialog] = useState(null);
  const [user, setUser] = useState(readUser);
  const [nameDraft, setNameDraft] = useState('');
  const [dirty, setDirty] = useState(false);
  const [recent, setRecent] = useState([]);
  const inputRef = useRef(null);
  const hostRef = useRef(null);
  const gridRef = useRef(null);
  const docRef = useRef(null);
  const shortcutSave = useRef(null);

  docRef.current = doc;

  const isDirty = () => {
    const grid = gridRef.current;
    return grid ? grid.isEditing() || grid.collectEdits().length > 0 : false;
  };
  const pushNotice = useCallback((n) => {
    setNotices((list) => [...list.filter((x) => x.id !== n.id), n]);
    setClosed((set) => {
      const next = new Set(set);
      next.delete(n.id);
      return next;
    });
  }, []);

  useEffect(() => () => gridRef.current?.destroy(), []);

  // Màn chào: quay lại file đã làm, khỏi tải lại. Lỗi thì chỉ không có danh sách, vẫn mở file được.
  useEffect(() => {
    fetchRecent().then(setRecent).catch(() => {});
  }, []);

  // Ctrl+S là Lưu của app, không mở hộp lưu trang web của trình duyệt.
  useEffect(() => {
    const onKey = (e) => {
      if ((e.ctrlKey || e.metaKey) && !e.altKey && e.key.toLowerCase() === 's') {
        e.preventDefault();
        shortcutSave.current?.();
      }
    };
    window.addEventListener('keydown', onKey, true);
    return () => window.removeEventListener('keydown', onKey, true);
  }, []);

  // Q86: sửa chưa lưu không được mất im lặng. Đóng tab thì trình duyệt cảnh báo.
  useEffect(() => {
    const warn = (e) => {
      if (isDirty()) {
        e.preventDefault();
        e.returnValue = '';
      }
    };
    window.addEventListener('beforeunload', warn);
    return () => window.removeEventListener('beforeunload', warn);
  }, []);

  const showGrid = useCallback(async (meta, { oldVersion = null, keep = [] } = {}) => {
    const model = await fetchGrid(meta.document_id, oldVersion);
    const sheetBefore = gridRef.current && docRef.current?.document_id === meta.document_id ? gridRef.current.activeSheet() : null;
    gridRef.current?.destroy();
    gridRef.current = null;
    setDoc(meta);
    setViewing(oldVersion);
    setClosed(new Set());
    setDirty(false);
    setNotices([...documentNotices(meta), ...keep]);
    await new Promise((r) => requestAnimationFrame(r));
    const grid = createGrid(hostRef.current, model, {
      readOnly: oldVersion !== null,
      onAmbiguousNumber: (info) => pushNotice(ambiguousNumberNotice(info)),
      onDecimalSwapped: (info) => pushNotice(decimalSwappedNotice(info)),
      onHiddenDigits: (info) => pushNotice(hiddenDigitsNotice(info)),
      onLockedCell: (info) => pushNotice(lockedCellNotice(info)),
      // Ô đã ghi hoặc đã bỏ: lời báo "chưa ghi" của lần gõ trước không còn đúng.
      onEditDone: () => setNotices((list) => list.filter((n) => n.id !== 'ambiguous' && n.id !== 'hidden-digits')),
      onChange: () => {
        const now = isDirty();
        setDirty(now);
        // Có sửa mới thì "Đã lưu bản N" không còn đúng với thứ đang thấy.
        if (now) setNotices((list) => list.filter((n) => n.id !== 'saved' && n.id !== 'restored'));
      },
      onPasted: (info) => {
        const note = pastedNotice(info);
        if (note) pushNotice(note);
      },
    });
    gridRef.current = grid;
    window.vmGrid = grid;
    if (sheetBefore) grid.goTo(sheetBefore, null);
    const mismatches = await grid.compareWithSaved();
    // Trong lúc chờ tính, người dùng có thể đã mở file khác: kết quả này không còn thuộc về lưới đang hiện.
    if (gridRef.current !== grid) return;
    const note = mismatchNotice(mismatches);
    if (note) setNotices((list) => [...list, note]);
    window.vmGridReady = `${meta.document_id}:${meta.version}:${oldVersion ?? ''}`;
  }, [pushNotice]);

  const run = useCallback(async (work, note = '') => {
    setBusy(true);
    setError(null);
    try {
      await work();
    } catch (e) {
      if (e instanceof ApiError && e.code === 'WRITE_REFUSED') {
        pushNotice(refusedNotice(e.detail?.cells ?? [], e.message));
      } else if (e instanceof ApiError && e.code === 'PDF_MISMATCH') {
        pushNotice(pdfMismatchNotice(e.detail?.cells ?? [], e.message));
      } else {
        setError((e instanceof ApiError ? e.message : 'App chưa làm xong việc này. File của bạn không bị đổi. Hãy thử lại, hoặc hỏi người quản lý app nếu vẫn không được.') + note);
      }
    } finally {
      setBusy(false);
    }
  }, [pushNotice]);

  const refreshVersions = useCallback(async (id) => {
    setVersions(await fetchVersions(id));
  }, []);

  const save = useCallback(async (author) => {
    const edits = gridRef.current.collectEdits();
    if (!edits.length) {
      pushNotice({ id: 'nothing', kind: 'info', text: 'Chưa có gì thay đổi nên app không tạo bản mới.', places: [] });
      return true;
    }
    // Q124: quá giới hạn thì báo ngay, không gửi đi. Lỗi làm dừng việc tiếp theo (mở file khác...), các sửa còn trên màn hình.
    if (edits.length > MAX_EDITS) throw new ApiError(EDITS_TOO_MANY, 'EDITS_TOO_MANY');
    const meta = await saveVersion(docRef.current.document_id, { base_version: docRef.current.version, author, edits });
    await showGrid(meta, { keep: [savedNotice(meta.version, edits.length)] });
    if (versions) await refreshVersions(meta.document_id);
    return true;
  }, [pushNotice, showGrid, versions, refreshVersions]);

  // Chạy `next` khi không còn sửa dở; có sửa dở thì hỏi lưu, bỏ, hoặc ở lại (Q86, Q92).
  // Ô đang gõ dở được xác nhận trước. Nếu số nhập nhằng thì dừng ở đó: app đã báo và ô vẫn mở, không mất gì.
  // keepsGrid: việc tiếp theo vẫn dùng lưới hiện tại, nên "Bỏ" phải nạp lại bản đã lưu trước khi làm tiếp.
  const guard = async (next, { keepsGrid = false } = {}) => {
    if (gridRef.current && !(await gridRef.current.commitEditing())) return;
    if (!isDirty()) {
      next();
      return;
    }
    setDialog({ kind: 'unsaved', next, keepsGrid });
  };

  const askName = (then) => {
    if (user) {
      then(user);
      return;
    }
    setNameDraft('');
    setDialog({ kind: 'name', then });
  };

  // Chọn file trước rồi mới hỏi: hủy hộp chọn file thì không có gì thay đổi.
  const pick = () => inputRef.current?.click();
  const onChosen = (e) => {
    const file = e.target.files?.[0];
    e.target.value = '';
    if (file) {
      guard(() => run(async () => {
        const meta = await uploadDocument(file);
        setVersions(null);
        await showGrid(meta);
      }));
    }
  };
  const onSave = async () => {
    if (gridRef.current && !(await gridRef.current.commitEditing())) return;
    askName((author) => run(() => save(author), SAVE_FAILED_NOTE));
  };
  const openExportWarning = () => setDialog({ kind: 'export', written: Boolean(docRef.current.written) });
  const onExport = () => guard(openExportWarning, { keepsGrid: true });
  const openRecent = (id) => run(async () => {
    const meta = await fetchDocument(id);
    setVersions(null);
    await showGrid(meta);
  });
  const onVersions = () => {
    if (versions) setVersions(null);
    else run(() => refreshVersions(doc.document_id));
  };
  const viewOld = (n) => guard(() => run(() => showGrid(docRef.current, { oldVersion: n })));
  const backToLatest = () => run(() => showGrid(doc));
  const restore = (n) => setDialog({ kind: 'restore', version: n });
  const go = (place) => gridRef.current?.goTo(place.sheet, place.cell);

  const visible = notices.filter((n) => !closed.has(n.id));
  const hiddenCount = notices.length - visible.length;
  const canEdit = doc && viewing === null;
  const shownVersion = viewing ?? doc?.version;
  shortcutSave.current = () => {
    if (canEdit && !busy && !dialog) onSave();
  };

  return (
    <div className="app">
      <header className="topbar">
        <div className="logo" title="Logo" aria-hidden="true" />
        <div className="doc-name" data-testid="doc-name">{doc ? doc.name : 'VM-Office-Assistant'}</div>
        {doc && <div className="doc-state" data-testid="doc-state">{viewing !== null ? `Đang xem bản ${viewing}, chỉ đọc` : `Bản ${doc.version}`}</div>}
        {doc && viewing === null && dirty && <div className="doc-dirty" data-testid="dirty-state">Có sửa chưa lưu</div>}
        {user && <div className="doc-user">{user}</div>}
      </header>
      {busy && <div className="progress" role="progressbar" aria-label="Đang xử lý" />}
      <div className="body">
        <nav className="sidebar" aria-label="Chức năng" onMouseDown={(e) => e.preventDefault()}>
          <button type="button" className="side-btn" onClick={pick}>
            <Icon d="M4 19V6h6l2 2h8v11H4z" />Mở
          </button>
          {canEdit && (
            <button type="button" className="side-btn" onClick={onSave} disabled={busy}>
              <Icon d="M5 4h11l3 3v13H5zM8 4v5h7V4M8 20v-6h8v6" />Lưu
            </button>
          )}
          {doc && (
            <button type="button" className="side-btn" onClick={onExport}>
              <Icon d="M12 4v11m0 0l-4-4m4 4l4-4M5 20h14" />Xuất
            </button>
          )}
          {doc && (
            <button type="button" className="side-btn" onClick={onVersions}>
              <Icon d="M12 7v5l3 2M4 12a8 8 0 1 0 2.5-5.8M4 4v4h4" />Phiên bản
            </button>
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
                <NoticeBar key={`${doc.document_id}-${n.id}-${n.text}`} notice={n} onGo={go} onClose={() => setClosed(new Set([...closed, n.id]))} />
              ))}
              {hiddenCount > 0 && (
                <div className="reopen">
                  <button type="button" className="btn-plain" onClick={() => setClosed(new Set())}>Xem lại {hiddenCount} thông báo</button>
                </div>
              )}
            </div>
          )}
          {canEdit && (
            <div className="editbar" onMouseDown={(e) => e.preventDefault()}>
              <button type="button" className="btn-line" onClick={() => gridRef.current?.undo()}>Hoàn tác</button>
              <button type="button" className="btn-line" onClick={() => gridRef.current?.redo()}>Làm lại</button>
            </div>
          )}
          {viewing !== null && (
            <div className="editbar" onMouseDown={(e) => e.preventDefault()}>
              <span>Đang xem bản {viewing}, chỉ đọc.</span>
              <button type="button" className="btn-line" onClick={backToLatest}>Về bản mới nhất</button>
            </div>
          )}
          {doc ? (
            <div className="grid-host" ref={hostRef} data-testid="grid" />
          ) : (
            <div className="welcome">
              <h1>Mở biểu mẫu để bắt đầu</h1>
              <p>App mở được file Excel .xlsx, .xlsm và .xls. File gốc của bạn không bị đổi.</p>
              <button type="button" className="btn" onClick={pick} disabled={busy}>Mở file</button>
              {recent.length > 0 && (
                <div className="recent" data-testid="recent">
                  <h2>Làm tiếp file đã mở trong app</h2>
                  <ul>
                    {recent.map((r) => (
                      <li key={r.document_id}>
                        <button type="button" className="btn-plain" data-document={r.document_id} onClick={() => openRecent(r.document_id)} disabled={busy}>{r.name}</button>
                        <span className="v-meta"> Bản {r.version} · {new Date(r.updated_at).toLocaleString('vi-VN')}{r.author ? ` · ${r.author}` : ''}</span>
                      </li>
                    ))}
                  </ul>
                </div>
              )}
            </div>
          )}
        </main>
        {versions && doc && (
          <aside className="versions" aria-label="Phiên bản" data-testid="versions">
            <h2>Phiên bản</h2>
            <ul>
              {versions.map((v) => (
                <li key={v.version} data-version={v.version} className={v.version === shownVersion ? 'current' : ''}>
                  <div className="v-title">Bản {v.version}{v.version === doc.version ? ' (mới nhất)' : ''}</div>
                  <div className="v-meta">{new Date(v.created_at).toLocaleString('vi-VN')}{v.author ? ` · ${v.author}` : ''}</div>
                  <div className="v-meta">{v.summary}</div>
                  {v.version !== doc.version && (
                    <div className="v-actions">
                      <button type="button" className="btn-plain" onClick={() => viewOld(v.version)}>Xem</button>
                      <button type="button" className="btn-plain" onClick={() => restore(v.version)}>Khôi phục bản này</button>
                    </div>
                  )}
                </li>
              ))}
            </ul>
          </aside>
        )}
      </div>

      {dialog?.kind === 'unsaved' && (
        <Dialog
          title="Có sửa chưa lưu"
          actions={[
            { label: 'Lưu', primary: true, onClick: () => { const { next } = dialog; setDialog(null); askName((a) => run(async () => { await save(a); next(); }, SAVE_FAILED_NOTE)); } },
            { label: 'Bỏ thay đổi', onClick: () => {
              const { next, keepsGrid } = dialog;
              setDialog(null);
              if (keepsGrid) run(async () => { await showGrid(docRef.current); next(); });
              else next();
            } },
            { label: 'Ở lại', onClick: () => setDialog(null) },
          ]}
        >
          Bạn có thay đổi chưa lưu. Lưu thành bản mới, bỏ các thay đổi này, hay ở lại để sửa tiếp?
        </Dialog>
      )}
      {dialog?.kind === 'name' && (
        <Dialog
          title="Tên của bạn"
          actions={[
            { label: 'Tiếp tục', primary: true, disabled: !nameDraft.trim(), onClick: () => { const name = nameDraft.trim(); rememberUser(name); setUser(name); const { then } = dialog; setDialog(null); then(name); } },
            { label: 'Hủy', onClick: () => setDialog(null) },
          ]}
        >
          <label>App ghi tên này vào lịch sử để biết ai đã sửa. Tên bạn tự gõ, bản thử nghiệm chưa kiểm tra đăng nhập.
            <input className="text-input" value={nameDraft} onChange={(e) => setNameDraft(e.target.value)} autoFocus data-testid="name-input" maxLength={120} />
          </label>
        </Dialog>
      )}
      {dialog?.kind === 'export' && (
        <Dialog
          title="Xuất file"
          actions={[
            { label: 'Tải PDF', onClick: () => { setDialog(null); run(async () => {
              await gridRef.current.settled();
              const { blob, name } = await exportPdf(doc.document_id, { version: viewing, grid_values: gridRef.current.formulaValues() });
              const url = URL.createObjectURL(blob);
              const a = document.createElement('a');
              a.href = url;
              a.download = name;
              document.body.appendChild(a);
              a.click();
              a.remove();
              URL.revokeObjectURL(url);
            }); } },
            { label: 'Tải file Excel', primary: true, onClick: () => { setDialog(null); const a = document.createElement('a'); a.href = exportUrl(doc.document_id, viewing); a.download = ''; document.body.appendChild(a); a.click(); a.remove(); } },
            { label: 'Hủy', onClick: () => setDialog(null) },
          ]}
        >
          {(dialog.written ? EXPORT_WARNING : doc.legacy_original_sha256 ? EXPORT_UNCHANGED_LEGACY : EXPORT_UNCHANGED).map((text) => <p key={text}>{text}</p>)}
        </Dialog>
      )}
      {dialog?.kind === 'restore' && (
        <Dialog
          title={`Khôi phục bản ${dialog.version}`}
          actions={[
            { label: 'Khôi phục', primary: true, onClick: () => { const n = dialog.version; setDialog(null); guardedRestore(n); } },
            { label: 'Hủy', onClick: () => setDialog(null) },
          ]}
        >
          Sẽ tạo bản mới với nội dung bản {dialog.version}. Các bản hiện có không bị xóa.
        </Dialog>
      )}
    </div>
  );

  function guardedRestore(n) {
    guard(() => askName((author) => run(async () => {
      const current = docRef.current;
      const meta = await restoreVersion(current.document_id, n, { base_version: current.version, author });
      await showGrid(meta, { keep: [{ id: 'restored', kind: 'info', text: `Đã tạo bản ${meta.version} với nội dung bản ${n}. Các bản cũ vẫn còn.`, places: [] }] });
      await refreshVersions(meta.document_id);
    })));
  }
}
