import React, { useCallback, useEffect, useRef, useState } from 'react';
import { ApiError, applyFill, cancelFill, createFill, editsTooMany, MAX_EDITS, exportPdf, exportUrl, fetchDocument, fetchFill, fetchGrid, fetchOpenFills, fetchRecent, fetchVersions, openFill, restoreVersion, saveVersion, uploadDocument, uploadSource } from './api.js';
import texts from '@vmai-texts';
import FillPanel from './fill/FillPanel.jsx';
import { toDecisions } from './fill/fillModel.js';
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

// Lát 1 nối AI (Q133). Câu nhắc file giả: nguyên văn PO/BA (acceptance_lat1 1.5).
const FILL_FAKE_ONLY = 'Bản thử này chỉ dùng với file giả đội gửi kèm. Anh/chị đừng chọn file kết quả thật của khoa.';
const FILL_WRONG_DOC = 'Lần điền này thuộc một biểu mẫu khác với biểu mẫu đang mở nên app chưa ghi. Không biểu mẫu nào bị đổi. Hãy mở lại đúng biểu mẫu rồi Ghi.';
const FILL_ONLY_QC = 'Bản thử này chỉ điền được biểu mẫu QC mẫu (GIA-QC-01_mau.xlsx). Biểu mẫu của anh/chị không bị đổi. Hãy mở biểu mẫu QC mẫu đội gửi kèm rồi bấm Điền từ file.';

const CALCULATING_TEXT = 'Đang tính lại các ô công thức, số trong các ô đó có thể chưa mới...';
const CALC_TIMEOUT_TEXT = 'App chưa tính xong các ô công thức nên chưa xuất PDF, để PDF không có số cũ. Biểu mẫu của anh/chị không bị đổi. Hãy chờ dòng "Đang tính lại" biến mất rồi bấm Tải PDF lại.';

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
  // T-DEV-08: bộ tính đang chạy thì ô công thức có thể còn số cũ. Chỉ hiện khi tính lâu hơn 300 ms, để không nhấp nháy.
  const [calculating, setCalculating] = useState(false);
  const [undoRedo, setUndoRedo] = useState({ undos: 0, redos: 0 });
  const calcTimer = useRef(null);
  const [error, setError] = useState(null);
  const [notices, setNotices] = useState([]);
  const [closed, setClosed] = useState(new Set());
  const [versions, setVersions] = useState(null);
  const [dialog, setDialog] = useState(null);
  const [user, setUser] = useState(readUser);
  const [nameDraft, setNameDraft] = useState('');
  const [dirty, setDirty] = useState(false);
  const [recent, setRecent] = useState([]);
  // Lát 1: lô đang duyệt, ô cần duyệt lại (3.8a), lệnh hủy lúc đang tạo lô (6.1).
  const [fill, setFill] = useState(null);
  const [reviewAgain, setReviewAgain] = useState({});
  const fillInputRef = useRef(null);
  const fillAbort = useRef(null);
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
    // Mở tài liệu khác: khung Điền từ file của tài liệu cũ đóng lại; lô cũ vẫn chờ, mở lại tài liệu đó thì có câu nhắc (PO/BA 11/11).
    if (docRef.current?.document_id !== meta.document_id) {
      setFill(null);
      setReviewAgain({});
    }
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
      onUndoRedo: (st) => setUndoRedo({ undos: st?.undos ?? 0, redos: st?.redos ?? 0 }),
      onCalculating: (on) => {
        clearTimeout(calcTimer.current);
        if (on) calcTimer.current = setTimeout(() => setCalculating(true), 300);
        else setCalculating(false);
      },
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
    if (oldVersion === null) {
      fetchOpenFills(meta.document_id).then((open) => {
        if (!open.length || gridRef.current !== grid) return;
        const first = open[0];
        pushNotice({ id: 'fill-open', kind: 'info', text: `Còn một lần điền từ file ${first.source_name} chưa xong: có ô đang chờ duyệt.`, places: [{ label: 'Mở lại lần điền này', fillId: first.batch_id }] });
      }).catch(() => {});
    }
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
        // Câu quá 5000 ô đã nói các sửa còn trên màn hình ở đầu câu, không nối thêm ở cuối (PO/BA lượt 3).
        const tail = e instanceof ApiError && e.code === 'EDITS_TOO_MANY' ? '' : note;
        setError((e instanceof ApiError ? e.message : 'App chưa làm xong việc này. File của anh/chị không bị đổi. Hãy thử lại, hoặc hỏi người quản lý app nếu vẫn không được.') + tail);
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
    if (edits.length > MAX_EDITS) throw new ApiError(editsTooMany(edits.length), 'EDITS_TOO_MANY');
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
  const reopenFill = (batchId) => run(async () => {
    const b = await fetchFill(batchId);
    setReviewAgain({});
    setVersions(null);
    setFill(b);
    setNotices((list) => list.filter((n) => n.id !== 'fill-open'));
  });
  const go = (place) => (place.fillId ? reopenFill(place.fillId) : gridRef.current?.goTo(place.sheet, place.cell));
  // Nút Điền từ file: chỉ biểu mẫu QC mẫu (báo trước, không gọi API); có sửa chưa lưu thì hỏi trước (6.3).
  const onFill = () => {
    if (!docRef.current?.fill_supported) {
      setError(FILL_ONLY_QC);
      return;
    }
    guard(() => fillInputRef.current?.click());
  };
  const onFillChosen = (e) => {
    const file = e.target.files?.[0];
    e.target.value = '';
    if (file) setDialog({ kind: 'fill-confirm', file });
  };
  const startFill = (file) => {
    const ctrl = new AbortController();
    fillAbort.current = ctrl;
    setDialog({ kind: 'fill-wait' });
    let batchId = null;
    run(async () => {
      try {
        const meta = docRef.current;
        const src = await uploadSource(meta.document_id, file, ctrl.signal);
        if (ctrl.signal.aborted) return;
        const created = await createFill(meta.document_id, { source_id: src.source_id, base_version: meta.version }, ctrl.signal);
        batchId = created.batch_id;
        if (ctrl.signal.aborted) {
          await cancelFill(batchId).catch(() => {});
          return;
        }
        const opened = await openFill(batchId);
        setReviewAgain({});
        setVersions(null);
        setFill(opened);
      } catch (err) {
        // Người dùng đã bấm Hủy: không báo lỗi mạng, xóa lô nháp nếu đã có (6.1).
        if (ctrl.signal.aborted) {
          if (batchId) await cancelFill(batchId).catch(() => {});
          return;
        }
        throw err;
      } finally {
        fillAbort.current = null;
        setDialog((d) => (d?.kind === 'fill-wait' ? null : d));
      }
    });
  };
  // Ghi: có sửa chưa lưu thì hỏi trước (6.3). Chỉ gửi change_id và lựa chọn; số lấy từ lô ở máy chủ (4.6).
  const applyChoices = (choices) => guard(() => askName((author) => run(async () => {
    const current = fill;
    // Lô phải thuộc tài liệu đang hiện; máy chủ kiểm lại bằng document_id (FILL_WRONG_DOCUMENT).
    if (current.document_id !== docRef.current.document_id) {
      setFill(null);
      throw new ApiError(FILL_WRONG_DOC, 'FILL_WRONG_DOCUMENT');
    }
    const r = await applyFill(current.batch_id, { document_id: docRef.current.document_id, base_version: docRef.current.version, author, decisions: toDecisions(choices) });
    const again = r.review_again ?? {};
    const keep = [];
    const left = r.pending;
    if (r.version) keep.push({ id: 'fill-done', kind: 'info', text: `Đã ghi ${r.written} ô từ file ${current.source_name} thành bản ${r.version.version}, còn ${left} ô chưa ghi.`, places: [] });
    if (Object.keys(again).length) keep.push({ id: 'fill-again', kind: 'warn', text: `${Object.keys(again).length} ô đã bị đổi ở chỗ khác từ lúc tạo đề xuất nên app chưa ghi. Hãy xem số hiện tại → số mới trong khung Điền từ file rồi Nhận lại nếu vẫn muốn ghi.`, places: [] });
    if (r.version) await showGrid(r.version, { keep });
    else keep.forEach(pushNotice);
    setReviewAgain(again);
    const b = await fetchFill(current.batch_id);
    setFill(b.status === 'done' ? null : b);
  })));

  // N5: bỏ lần điền đang mở. Không tạo bản mới; các ô đã ghi giữ nguyên; ô của lô không còn bị giữ (D10).
  const dropFill = () => run(async () => {
    const current = fill;
    const r = await cancelFill(current.batch_id);
    setFill(null);
    setReviewAgain({});
    setNotices((list) => list.filter((n) => n.id !== 'fill-open' && n.id !== 'fill-again'));
    pushNotice({ id: 'fill-dropped', kind: 'info', text: `Đã bỏ lần điền từ file ${current.source_name}: ${r.dropped ?? 0} ô không ghi. Biểu mẫu không bị đổi.`, places: [] });
  });

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
          {canEdit && (
            <button type="button" className="side-btn" onClick={onFill} disabled={busy} data-testid="fill-button">
              <Icon d="M4 5h7v14H4zM13 12h7M17 8l4 4-4 4" />Điền từ file
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
          <input ref={fillInputRef} type="file" accept=".xlsx" hidden onChange={onFillChosen} data-testid="fill-input" />
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
              <button type="button" className="btn-line" disabled={!undoRedo.undos} onClick={() => gridRef.current?.undo()}>Hoàn tác</button>
              <button type="button" className="btn-line" disabled={!undoRedo.redos} onClick={() => gridRef.current?.redo()}>Làm lại</button>
              {calculating && <span className="calc-status" role="status" data-testid="calc-status">{CALCULATING_TEXT}</span>}
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
              <p>App mở được file Excel .xlsx, .xlsm và .xls. File gốc của anh/chị không bị đổi.</p>
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
        {fill && doc && viewing === null && (
          <FillPanel fill={fill} texts={texts} busy={busy} reviewAgain={reviewAgain} onApply={applyChoices}
            onClose={() => setFill(null)} onGo={(a1) => gridRef.current?.goTo(fill.batch.changes[0]?.target.sheet, a1)}
            onDrop={(pending) => setDialog({ kind: 'fill-drop', pending })} />
        )}
        {versions && doc && !fill && (
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
          Anh/chị có thay đổi chưa lưu. Lưu thành bản mới, bỏ các thay đổi này, hay ở lại để sửa tiếp?
        </Dialog>
      )}
      {dialog?.kind === 'name' && (
        <Dialog
          title="Tên của anh/chị"
          actions={[
            { label: 'Tiếp tục', primary: true, disabled: !nameDraft.trim(), onClick: () => { const name = nameDraft.trim(); rememberUser(name); setUser(name); const { then } = dialog; setDialog(null); then(name); } },
            { label: 'Hủy', onClick: () => setDialog(null) },
          ]}
        >
          <label>App ghi tên này vào lịch sử để biết ai đã sửa. Tên anh/chị tự gõ, bản thử nghiệm chưa kiểm tra đăng nhập.
            <input className="text-input" value={nameDraft} onChange={(e) => setNameDraft(e.target.value)} autoFocus data-testid="name-input" maxLength={120} />
          </label>
        </Dialog>
      )}
      {dialog?.kind === 'fill-confirm' && (
        <Dialog
          title="Điền từ file"
          actions={[
            { label: 'Điền', primary: true, onClick: () => startFill(dialog.file) },
            { label: 'Hủy', onClick: () => setDialog(null) },
          ]}
        >
          <p>Điền biểu mẫu <b>{doc?.name}</b> bằng số từ file <b>{dialog.file.name}</b>?</p>
          <p>App chỉ đề xuất, chưa ghi gì vào biểu mẫu. Anh/chị duyệt từng dòng rồi mới Ghi.</p>
          <p className="fill-warning">{FILL_FAKE_ONLY}</p>
        </Dialog>
      )}
      {dialog?.kind === 'fill-drop' && (
        <Dialog
          title="Bỏ lần điền này"
          actions={[
            { label: 'Bỏ lần điền', primary: true, onClick: () => { setDialog(null); dropFill(); } },
            { label: 'Hủy', onClick: () => setDialog(null) },
          ]}
        >
          <p>Bỏ lần điền từ file <b>{fill?.source_name}</b>? {dialog.pending} ô còn chờ sẽ bị bỏ, app không ghi các ô này.</p>
          <p>Các ô đã ghi trước đó vẫn giữ trong các bản đã lưu. Muốn điền lại thì bấm Điền từ file.</p>
        </Dialog>
      )}
      {dialog?.kind === 'fill-wait' && (
        <Dialog title="Đang tạo đề xuất" actions={[{ label: 'Hủy', onClick: () => { fillAbort.current?.abort(); setDialog(null); } }]}>
          Đang tạo đề xuất từ file nguồn. Biểu mẫu chưa bị đổi.
        </Dialog>
      )}
      {dialog?.kind === 'export' && (
        <Dialog
          title="Xuất file"
          actions={[
            { label: 'Tải PDF', onClick: () => { setDialog(null); run(async () => {
              const calc = await gridRef.current.settled();
              if (calc && !calc.ok) throw new ApiError(CALC_TIMEOUT_TEXT, 'CALC_TIMEOUT');
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
