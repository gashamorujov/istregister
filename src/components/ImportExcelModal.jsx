import { useRef, useState } from 'react';
import ExcelJS from 'exceljs';
import {
  CloseIcon, DocIcon, CheckIcon, ClipboardIcon, WarningIcon, ArrowLeftIcon,
} from './Icons';
import {
  TARGET_FIELDS, FIELD_LABELS, rowKey, detectColumnMapping, HEADER_MATCH_THRESHOLD,
  emptyFieldsOf, hasIdentitySignal,
} from '../lib/importMapping';

// Generous upper bound on how many spreadsheet/clipboard columns we scan per
// row when looking for header text or reading values — comfortably above the
// ~15 columns REGİSTR-2026 actually uses, so a reordered or slightly widened
// sheet still gets read correctly.
const COL_SCAN_LIMIT = 40;

const STEP_TITLES = {
  choose: 'Import',
  clipboard: 'Clipboard ilə idxal',
  file: 'Excel sənədi ilə idxal',
  preview: 'Önizləmə (Preview)',
};

// Approximate preview-table column widths, echoing the real widths used in
// the site's own grid so the preview reads as close to "how it will look on
// the site" as a plain HTML table reasonably can.
const PREVIEW_COL_WIDTHS = {
  fullName: 210, serial: 110, idNumber: 120, birthDate: 100, phone: 140,
  email: 190, rank: 170, fullNameId: 200, rank2: 130, courseCode: 100,
  startDate: 110, finishDate: 110, note: 120, date: 100,
};

function cleanVal(v) {
  return String(v === null || v === undefined ? '' : v).replace(/\[object Object\]/g, '').trim();
}

// Reads one ExcelJS cell into plain display text, handling dates, formulas
// (using their cached result), and rich text runs.
function parseCellText(cell) {
  if (!cell) return '';
  const v = cell.value;
  if (v === null || v === undefined) return '';
  if (v instanceof Date && !isNaN(v.getTime())) {
    const d = String(v.getUTCDate()).padStart(2, '0');
    const m = String(v.getUTCMonth() + 1).padStart(2, '0');
    return `${d}.${m}.${v.getUTCFullYear()}`;
  }
  if (v.richText && Array.isArray(v.richText)) return v.richText.map((rt) => rt.text || '').join('').trim();
  if (v.result !== undefined && v.result !== null) {
    if (v.result instanceof Date && !isNaN(v.result.getTime())) {
      const d = String(v.result.getUTCDate()).padStart(2, '0');
      const m = String(v.result.getUTCMonth() + 1).padStart(2, '0');
      return `${d}.${m}.${v.result.getUTCFullYear()}`;
    }
    return String(v.result).trim();
  }
  if (v.text) return String(v.text).trim();
  return cleanVal(v);
}

// Scans the first few rows of a worksheet for the one whose cell text best
// matches our known column headers — by wording, never by column letter —
// and returns the winning row number plus its column→field mapping.
function findHeaderInSheet(sheet) {
  let best = null;
  for (let r = 1; r <= 5; r++) {
    const row = sheet.getRow(r);
    const texts = [];
    for (let c = 1; c <= COL_SCAN_LIMIT; c++) texts.push(parseCellText(row.getCell(c)));
    const { colIndexToField, matchedFields } = detectColumnMapping(texts);
    if (matchedFields.size >= HEADER_MATCH_THRESHOLD && (!best || matchedFields.size > best.matchedFields.size)) {
      best = { rowNumber: r, colIndexToField, matchedFields };
    }
  }
  return best;
}

export default function ImportExcelModal({ existingKeys, onConfirm, onCancel }) {
  const [step, setStep] = useState('choose'); // 'choose' | 'clipboard' | 'file' | 'preview'
  const [source, setSource] = useState(''); // which input step preview should return to
  const [fileName, setFileName] = useState('');
  const [clipboardText, setClipboardText] = useState('');
  const [previewRows, setPreviewRows] = useState(null);
  const [newCount, setNewCount] = useState(0);
  const [skippedCount, setSkippedCount] = useState(0);
  const [emptyCount, setEmptyCount] = useState(0);
  const [missingCols, setMissingCols] = useState([]);
  const [usedFallback, setUsedFallback] = useState(false);
  const [error, setError] = useState('');
  const [loading, setLoading] = useState(false);
  const [dragOver, setDragOver] = useState(false);
  const inputRef = useRef(null);

  const processParsed = (records) => {
    if (!records || records.length === 0) { setError('Heç bir məlumat tapılmadı.'); return; }
    const result = records.map((rec) => ({
      ...rec,
      IS_NEW: !existingKeys.has(rowKey(rec)),
      EMPTY_FIELDS: emptyFieldsOf(rec),
    }));
    const newRows = result.filter((r) => r.IS_NEW);
    setNewCount(newRows.length);
    setSkippedCount(result.length - newRows.length);
    setEmptyCount(result.filter((r) => r.EMPTY_FIELDS.length > 0).length);
    setPreviewRows(result);
    setStep('preview');
  };

  const resetOutcome = () => {
    setError(''); setPreviewRows(null); setNewCount(0); setSkippedCount(0);
    setEmptyCount(0); setMissingCols([]); setUsedFallback(false);
  };

  // ---------- Excel Document path ----------

  const handleFile = async (file) => {
    resetOutcome();
    if (!file) return;
    setFileName(file.name);
    setSource('file');
    setLoading(true);
    try {
      const workbook = new ExcelJS.Workbook();
      await workbook.xlsx.load(await file.arrayBuffer());

      // Only the REGİSTR-2026 worksheet is ever read. Every other sheet in the
      // workbook — visible, hidden, or helper — is left completely untouched.
      const sheet = workbook.getWorksheet('REGİSTR-2026');
      if (!sheet) { setError('Faylda "REGİSTR-2026" səhifəsi tapılmadı.'); setLoading(false); return; }

      const header = findHeaderInSheet(sheet);
      if (!header) {
        setError('REGİSTR-2026 səhifəsinin sütun başlıqları tanınmadı. Fayl strukturunu yoxlayın.');
        setLoading(false);
        return;
      }

      const parsed = [];
      sheet.eachRow((row, rowNumber) => {
        if (rowNumber <= header.rowNumber) return;
        const texts = [];
        for (let c = 1; c <= COL_SCAN_LIMIT; c++) texts.push(parseCellText(row.getCell(c)));
        const rec = {};
        header.colIndexToField.forEach((field, idx) => {
          rec[field] = texts[idx] || '';
        });
        // A row only counts if it actually identifies someone — guards against
        // a value dragged/filled far past the real data in just one column.
        if (!hasIdentitySignal(rec)) return;
        // Guard against a stray repeated header row (e.g. a printed multi-page export)
        if (detectColumnMapping(texts).matchedFields.size >= HEADER_MATCH_THRESHOLD) return;
        parsed.push(rec);
      });

      const missing = TARGET_FIELDS.filter((f) => !header.matchedFields.has(f));
      setMissingCols(missing.map((f) => FIELD_LABELS[f]));
      processParsed(parsed);
    } catch (err) {
      console.error('Excel parse xətası:', err);
      setError('Excel faylı oxunarkən xəta baş verdi. Düzgün .xlsx fayl seçdiyinizə əmin olun.');
    } finally {
      setLoading(false);
    }
  };

  const handleDrop = (e) => {
    e.preventDefault();
    setDragOver(false);
    const file = e.dataTransfer.files?.[0];
    if (file) handleFile(file);
  };

  // ---------- Clipboard path ----------

  const handleClipboardParse = () => {
    resetOutcome();
    if (!clipboardText.trim()) { setError('Kopyaladığınız mətn boşdur.'); return; }
    setSource('clipboard');

    const lines = clipboardText.split(/\r?\n/).filter((l) => l.trim() !== '');
    if (lines.length === 0) { setError('Kopyaladığınız mətn boşdur.'); return; }

    let headerIdx = -1;
    let mapping = null;
    const scanLimit = Math.min(lines.length, 5);
    for (let i = 0; i < scanLimit; i++) {
      const cells = lines[i].split('\t');
      const { colIndexToField, matchedFields } = detectColumnMapping(cells);
      if (matchedFields.size >= HEADER_MATCH_THRESHOLD && (!mapping || matchedFields.size > mapping.matchedFields.size)) {
        headerIdx = i; mapping = { colIndexToField, matchedFields };
      }
    }

    let fellBack = false;
    if (!mapping) {
      // No recognizable header text in the pasted content — fall back to the
      // sheet's own canonical column order (№ then the 14 site fields) so a
      // data-only paste still imports. The preview flags this so it can be
      // double-checked before confirming.
      fellBack = true;
      const colIndexToField = new Map();
      TARGET_FIELDS.forEach((f, i) => colIndexToField.set(i + 1, f));
      mapping = { colIndexToField, matchedFields: new Set(TARGET_FIELDS) };
      headerIdx = -1;
    }

    const records = [];
    for (let i = headerIdx + 1; i < lines.length; i++) {
      const cells = lines[i].split('\t');
      if (detectColumnMapping(cells).matchedFields.size >= HEADER_MATCH_THRESHOLD) continue; // stray repeated header
      const rec = {};
      mapping.colIndexToField.forEach((field, idx) => {
        rec[field] = cells[idx] !== undefined ? cleanVal(cells[idx]) : '';
      });
      if (hasIdentitySignal(rec)) records.push(rec);
    }

    setUsedFallback(fellBack);
    setMissingCols(TARGET_FIELDS.filter((f) => !mapping.matchedFields.has(f)).map((f) => FIELD_LABELS[f]));
    processParsed(records);
  };

  // ---------- Shared ----------

  const handleConfirm = () => {
    const newRows = (previewRows || [])
      .filter((r) => r.IS_NEW)
      .map((r) => {
        const copy = { ...r };
        delete copy.IS_NEW;
        delete copy.EMPTY_FIELDS;
        return copy;
      });
    onConfirm(newRows);
  };

  const goBack = () => {
    setError('');
    if (step === 'preview') {
      setPreviewRows(null);
      setStep(source === 'clipboard' ? 'clipboard' : 'file');
    } else {
      setStep('choose');
    }
  };

  return (
    <div className="modal-overlay" onClick={onCancel}>
      <div className="modal" style={{ maxWidth: step === 'preview' ? 1000 : 640 }} onClick={(e) => e.stopPropagation()}>
        <div className="modal-header">
          <h2>{STEP_TITLES[step]}</h2>
          <button className="modal-close" onClick={onCancel} aria-label="Bağla"><CloseIcon /></button>
        </div>
        <div className="modal-body">

          {step === 'choose' && (
            <div className="import-choice-wrap">
              <p className="import-choice-lead">Məlumatları necə idxal etmək istəyirsiniz?</p>
              <div className="import-choice-grid">
                <button type="button" className="import-choice-card" onClick={() => { setSource('clipboard'); setStep('clipboard'); }}>
                  <span className="import-choice-icon"><ClipboardIcon /></span>
                  <span className="import-choice-title">Clipboard</span>
                  <span className="import-choice-desc">Excel-dən kopyaladığınız (Ctrl+A → Ctrl+C) məlumatları birbaşa yapışdırın</span>
                </button>
                <button type="button" className="import-choice-card" onClick={() => { setSource('file'); setStep('file'); }}>
                  <span className="import-choice-icon"><DocIcon /></span>
                  <span className="import-choice-title">Excel Document</span>
                  <span className="import-choice-desc">.xlsx faylı yükləyin — yalnız "REGİSTR-2026" səhifəsi oxunacaq</span>
                </button>
              </div>
            </div>
          )}

          {step === 'clipboard' && (
            <div className="import-step">
              <p className="import-clipboard-hint">
                Excel-də <b>REGİSTR-2026</b> səhifəsini açın, <b>Ctrl+A</b> ilə hamısını seçin, <b>Ctrl+C</b> edin və aşağıya yapışdırın.
              </p>
              <textarea
                className="import-clipboard-area"
                rows={9}
                autoFocus
                placeholder="Excel-dən kopyaladığınız mətni buraya yapışdırın (Ctrl+V)..."
                value={clipboardText}
                onChange={(e) => setClipboardText(e.target.value)}
              />
              <button className="btn-primary import-clipboard-btn" onClick={handleClipboardParse} disabled={!clipboardText.trim()}>
                <ClipboardIcon /> Məlumatları analiz et
              </button>
              {error && <div className="import-parse-error"><WarningIcon /> {error}</div>}
            </div>
          )}

          {step === 'file' && (
            <div className="import-step">
              <div
                className={`import-dropzone ${dragOver ? 'drag' : ''}`}
                onClick={() => !loading && inputRef.current?.click()}
                onDragOver={(e) => { e.preventDefault(); if (!loading) setDragOver(true); }}
                onDragLeave={() => setDragOver(false)}
                onDrop={(e) => { if (loading) { e.preventDefault(); return; } handleDrop(e); }}
              >
                <div className="import-dropzone-icon"><DocIcon /></div>
                <div className="import-dropzone-text">{loading ? 'Fayl analiz edilir...' : 'Excel faylı seçin və ya buraya sürükləyin'}</div>
                <div className="import-dropzone-sub">.xlsx — yalnız "REGİSTR-2026" səhifəsi oxunacaq</div>
                <input
                  ref={inputRef}
                  type="file"
                  accept=".xlsx"
                  style={{ display: 'none' }}
                  onChange={(e) => handleFile(e.target.files?.[0])}
                />
              </div>
              {fileName && !loading && !error && <div className="import-file-name"><CheckIcon /> {fileName}</div>}
              {error && <div className="import-parse-error"><WarningIcon /> {error}</div>}
            </div>
          )}

          {step === 'preview' && previewRows && (
            <>
              <div className="import-summary-stats">
                <div className="import-stat import-stat-new"><b>{newCount}</b><span>yeni məlumat əlavə olunacaq</span></div>
                {skippedCount > 0 && (
                  <div className="import-stat import-stat-dup"><b>{skippedCount}</b><span>artıq mövcuddur (atlanacaq)</span></div>
                )}
                {emptyCount > 0 && (
                  <div className="import-stat import-stat-warn"><b>{emptyCount}</b><span>sətirdə boş xana var</span></div>
                )}
              </div>

              {missingCols.length > 0 && (
                <div className="import-note-warn">
                  <WarningIcon />
                  <span>Bu sütunlar mənbədə tapılmadı: {missingCols.join(', ')}</span>
                </div>
              )}
              {usedFallback && (
                <div className="import-note-warn">
                  <WarningIcon />
                  <span>Sütun başlıqları tanınmadı — məlumatlar defolt sütun ardıcıllığı ilə oxundu. Nəticəni diqqətlə yoxlayın.</span>
                </div>
              )}

              <div className="import-preview-table-wrap">
                <table className="import-preview-table import-preview-table-full">
                  <colgroup>
                    <col style={{ width: 78 }} />
                    {TARGET_FIELDS.map((f) => <col key={f} style={{ width: PREVIEW_COL_WIDTHS[f] || 140 }} />)}
                  </colgroup>
                  <thead>
                    <tr>
                      <th>Status</th>
                      {TARGET_FIELDS.map((f) => <th key={f}>{FIELD_LABELS[f]}</th>)}
                    </tr>
                  </thead>
                  <tbody>
                    {previewRows.length === 0 && (
                      <tr><td colSpan={TARGET_FIELDS.length + 1} className="import-empty">Heç bir məlumat tapılmadı.</td></tr>
                    )}
                    {previewRows.map((r, i) => (
                      <tr key={i}>
                        <td>
                          {r.IS_NEW
                            ? <span className="import-badge-new">Yeni</span>
                            : <span className="import-badge-new import-badge-existing">Mövcud</span>}
                        </td>
                        {TARGET_FIELDS.map((f) => {
                          const val = r[f];
                          const isEmpty = !String(val || '').trim();
                          return (
                            <td key={f} className={isEmpty ? 'import-cell-empty' : (r.IS_NEW ? 'import-colnew' : '')}>
                              {isEmpty ? <span className="import-no-field">boş</span> : val}
                            </td>
                          );
                        })}
                      </tr>
                    ))}
                  </tbody>
                </table>
              </div>
            </>
          )}

        </div>
        <div className="modal-footer">
          {step === 'choose' ? (
            <button className="btn-secondary" onClick={onCancel}>Ləğv et</button>
          ) : (
            <button className="btn-secondary" onClick={goBack}><ArrowLeftIcon /> Geri</button>
          )}
          {step === 'preview' && (
            <button className="btn-primary" disabled={newCount === 0} onClick={handleConfirm}>
              <CheckIcon /> Təsdiqlə ({newCount})
            </button>
          )}
        </div>
      </div>
    </div>
  );
}
