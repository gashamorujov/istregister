import { useRef, useState, useEffect } from 'react';
import ExcelJS from 'exceljs';
import {
  CloseIcon, DocIcon, CheckIcon, ClipboardIcon, WarningIcon, ArrowLeftIcon,
} from './Icons';
import {
  TARGET_FIELDS, FIELD_LABELS, rowKey, detectColumnMapping, HEADER_MATCH_THRESHOLD,
  emptyFieldsOf, hasIdentitySignal,
} from '../lib/importMapping';

const COL_SCAN_LIMIT = 40;

const STEP_TITLES = {
  choose: 'Import',
  clipboard: 'Clipboard ilə idxal',
  file: 'Fayl ilə idxal',
  preview: 'Önizləmə (Preview)',
};

const PREVIEW_COL_WIDTHS = {
  fullName: 210, serial: 110, idNumber: 120, birthDate: 100, phone: 140,
  email: 190, rank: 170, fullNameId: 200, rank2: 130, courseCode: 100,
  startDate: 110, finishDate: 110, note: 120, date: 100,
};

function cleanVal(v) {
  const s = String(v === null || v === undefined ? '' : v);
  return s.includes('[object Object]') ? '' : s.trim();
}

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
  if (typeof v === 'object') return '';
  return cleanVal(v);
}

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

// CSV / semicolon / comma delimited parser (respects quoted fields)
function parseCSVText(text) {
  const delimiter = text.includes('\t') ? '\t'
    : text.includes(';') ? ';'
    : ',';

  const lines = [];
  let current = '';
  let inQuotes = false;

  for (let i = 0; i < text.length; i++) {
    const ch = text[i];
    if (inQuotes) {
      if (ch === '"') {
        if (i + 1 < text.length && text[i + 1] === '"') {
          current += '"';
          i++;
        } else {
          inQuotes = false;
        }
      } else {
        current += ch;
      }
    } else {
      if (ch === '"') {
        inQuotes = true;
      } else if (ch === delimiter || ch === '\r' || ch === '\n') {
        if (ch === '\r' && text[i + 1] === '\n') i++;
        lines.push(current);
        current = '';
      } else {
        current += ch;
      }
    }
  }
  lines.push(current);

  // Group into rows based on delimiter
  const rows = [];
  let row = [];
  for (const cell of lines) {
    if (cell === '\r') continue;
    row.push(cell);
    // Row break on delimiter = newline
    if (lines.indexOf(cell) === lines.length - 1 || lines[lines.indexOf(cell) + 1] === undefined) {
      // Check if this was actually a row end
    }
  }

  // Simpler approach: split by newline first, then by delimiter
  const textLines = text.split(/\r?\n/).filter(l => l.trim() !== '');
  const parsed = [];
  for (const line of textLines) {
    const cells = [];
    let c = '';
    let inQ = false;
    for (let i = 0; i < line.length; i++) {
      const ch = line[i];
      if (inQ) {
        if (ch === '"') {
          if (i + 1 < line.length && line[i + 1] === '"') { c += '"'; i++; }
          else inQ = false;
        } else { c += ch; }
      } else {
        if (ch === '"') inQ = true;
        else if (ch === delimiter) { cells.push(c); c = ''; }
        else c += ch;
      }
    }
    cells.push(c);
    parsed.push(cells);
  }
  return parsed;
}

export default function ImportExcelModal({ existingKeys, existingRows, onConfirm, onCancel }) {
  const [step, setStep] = useState('choose');
  const [source, setSource] = useState('');
  const [fileName, setFileName] = useState('');
  const [clipboardText, setClipboardText] = useState('');
  const [previewRows, setPreviewRows] = useState(null);
  const [newCount, setNewCount] = useState(0);
  const [changedCount, setChangedCount] = useState(0);
  const [unchangedCount, setUnchangedCount] = useState(0);
  const [emptyCount, setEmptyCount] = useState(0);
  const [missingCols, setMissingCols] = useState([]);
  const [usedFallback, setUsedFallback] = useState(false);
  const [error, setError] = useState('');
  const [loading, setLoading] = useState(false);
  const [dragOver, setDragOver] = useState(false);
  const inputRef = useRef(null);

  // Compare existing rows map for changed detection
  const existingRowsMapRef = useRef(new Map());
  
  // Populate existingRowsMap from parent data
  useEffect(() => {
    if (existingRows && existingRows.length > 0) {
      const map = new Map();
      existingRows.forEach(r => {
        const key = [r.fullName, r.serial, r.idNumber, r.courseCode]
          .map(v => String(v || '').trim().toLowerCase()).join('|');
        map.set(key, r);
      });
      existingRowsMapRef.current = map;
    }
  }, [existingRows]);

  const processParsed = (records, existingMap) => {
    if (!records || records.length === 0) { setError('Heç bir məlumat tapılmadı.'); return; }
    const result = records.map((rec) => {
      const key = rowKey(rec);
      const existing = existingMap.get(key);
      let status = 'NEW';
      if (existing) {
        const changed = TARGET_FIELDS.some(f => String(rec[f] || '') !== String(existing[f] || ''));
        status = changed ? 'CHANGED' : 'UNCHANGED';
      }
      return { ...rec, _STATUS: status, EMPTY_FIELDS: emptyFieldsOf(rec) };
    });
    setNewCount(result.filter(r => r._STATUS === 'NEW').length);
    setChangedCount(result.filter(r => r._STATUS === 'CHANGED').length);
    setUnchangedCount(result.filter(r => r._STATUS === 'UNCHANGED').length);
    setEmptyCount(result.filter(r => r.EMPTY_FIELDS.length > 0).length);
    setPreviewRows(result);
    setStep('preview');
  };

  const resetOutcome = () => {
    setError(''); setPreviewRows(null); setNewCount(0); setChangedCount(0);
    setUnchangedCount(0); setEmptyCount(0); setMissingCols([]); setUsedFallback(false);
  };

  // Excel file handling
  const handleFile = async (file) => {
    resetOutcome();
    if (!file) return;
    setFileName(file.name);
    setSource('file');
    setLoading(true);
    try {
      const ext = file.name.split('.').pop().toLowerCase();
      let records = [];

      if (ext === 'xlsx') {
        const workbook = new ExcelJS.Workbook();
        await workbook.xlsx.load(await file.arrayBuffer());
        const sheet = workbook.getWorksheet('REGİSTR-2026') || workbook.worksheets[0];
        if (!sheet) { setError('Faylda heç bir səhifə tapılmadı.'); setLoading(false); return; }

        const header = findHeaderInSheet(sheet);
        if (!header) {
          // Fallback: try to read all rows with default column mapping
          const colIndexToField = new Map();
          TARGET_FIELDS.forEach((f, i) => colIndexToField.set(i + 1, f));
          const fallbackMapping = { colIndexToField, matchedFields: new Set(TARGET_FIELDS) };
          sheet.eachRow((row, rowNumber) => {
            if (rowNumber <= 1) return;
            const texts = [];
            for (let c = 1; c <= COL_SCAN_LIMIT; c++) texts.push(parseCellText(row.getCell(c)));
            const rec = {};
            fallbackMapping.colIndexToField.forEach((field, idx) => { rec[field] = texts[idx] || ''; });
            if (hasIdentitySignal(rec)) records.push(rec);
          });
          setUsedFallback(true);
          setMissingCols([]);
        } else {
          sheet.eachRow((row, rowNumber) => {
            if (rowNumber <= header.rowNumber) return;
            const texts = [];
            for (let c = 1; c <= COL_SCAN_LIMIT; c++) texts.push(parseCellText(row.getCell(c)));
            const rec = {};
            header.colIndexToField.forEach((field, idx) => { rec[field] = texts[idx] || ''; });
            if (!hasIdentitySignal(rec)) return;
            if (detectColumnMapping(texts).matchedFields.size >= HEADER_MATCH_THRESHOLD) return;
            records.push(rec);
          });
          const missing = TARGET_FIELDS.filter((f) => !header.matchedFields.has(f));
          setMissingCols(missing.map((f) => FIELD_LABELS[f]));
        }
      } else {
        // CSV / TXT: read as text and parse
        const text = await file.text();
        const parsed = parseCSVText(text);
        if (parsed.length < 2) { setError('Faylda kifayət qədər məlumat tapılmadı.'); setLoading(false); return; }

        const headerCells = parsed[0];
        const { colIndexToField, matchedFields } = detectColumnMapping(headerCells);

        if (matchedFields.size >= HEADER_MATCH_THRESHOLD) {
          for (let i = 1; i < parsed.length; i++) {
            const rec = {};
            colIndexToField.forEach((field, idx) => { rec[field] = parsed[i][idx] || ''; });
            if (hasIdentitySignal(rec)) records.push(rec);
          }
          const missing = TARGET_FIELDS.filter((f) => !matchedFields.has(f));
          setMissingCols(missing.map((f) => FIELD_LABELS[f]));
        } else {
          // Fallback: assume canonical column order
          const fallbackMapping = new Map();
          TARGET_FIELDS.forEach((f, i) => fallbackMapping.set(i, f));
          for (let i = 0; i < parsed.length; i++) {
            const rec = {};
            fallbackMapping.forEach((field, idx) => { rec[field] = parsed[i][idx] || ''; });
            if (hasIdentitySignal(rec)) records.push(rec);
          }
          setUsedFallback(true);
        }
      }

      // Build existing map for changed detection
      // We need the current rows from the parent — passed as existingKeys (Set) but we also need values
      // Since we only get existingKeys, we detect changed by key presence
      processParsed(records, existingRowsMapRef.current);
    } catch (err) {
      console.error('Fayl parse xətası:', err);
      setError('Fayl oxunarkən xəta baş verdi. Formatı yoxlayın.');
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

  // Clipboard path
  const handleClipboardParse = () => {
    resetOutcome();
    if (!clipboardText.trim()) { setError('Kopyaladığınız mətn boşdur.'); return; }
    setSource('clipboard');

    const lines = clipboardText.split(/\r?\n/).filter((l) => l.trim() !== '');
    if (lines.length === 0) { setError('Kopyaladığınız mətn boşdur.'); return; }

    // Detect delimiter
    const delimiter = lines[0].includes('\t') ? '\t' : lines[0].includes(';') ? ';' : ',';

    let headerIdx = -1;
    let mapping = null;
    const scanLimit = Math.min(lines.length, 5);
    for (let i = 0; i < scanLimit; i++) {
      const cells = lines[i].split(delimiter).map(c => c.trim().replace(/^"|"$/g, ''));
      const { colIndexToField, matchedFields } = detectColumnMapping(cells);
      if (matchedFields.size >= HEADER_MATCH_THRESHOLD && (!mapping || matchedFields.size > mapping.matchedFields.size)) {
        headerIdx = i; mapping = { colIndexToField, matchedFields };
      }
    }

    let fellBack = false;
    if (!mapping) {
      fellBack = true;
      const colIndexToField = new Map();
      TARGET_FIELDS.forEach((f, i) => colIndexToField.set(i, f));
      mapping = { colIndexToField, matchedFields: new Set(TARGET_FIELDS) };
      headerIdx = -1;
    }

    const records = [];
    for (let i = headerIdx + 1; i < lines.length; i++) {
      const cells = lines[i].split(delimiter).map(c => c.trim().replace(/^"|"$/g, ''));
      if (detectColumnMapping(cells).matchedFields.size >= HEADER_MATCH_THRESHOLD) continue;
      const rec = {};
      mapping.colIndexToField.forEach((field, idx) => {
        rec[field] = cells[idx] !== undefined ? cleanVal(cells[idx]) : '';
      });
      if (hasIdentitySignal(rec)) records.push(rec);
    }

    setUsedFallback(fellBack);
    setMissingCols(TARGET_FIELDS.filter((f) => !mapping.matchedFields.has(f)).map((f) => FIELD_LABELS[f]));
    processParsed(records, existingRowsMapRef.current);
  };

  // Confirm: pass ALL rows (new + changed) to parent for upsert
  const handleConfirm = () => {
    const toImport = (previewRows || [])
      .filter((r) => r._STATUS === 'NEW' || r._STATUS === 'CHANGED')
      .map((r) => {
        const copy = { ...r };
        delete copy._STATUS;
        delete copy.EMPTY_FIELDS;
        return copy;
      });
    if (toImport.length > 0) onConfirm(toImport);
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
                  <span className="import-choice-title">Fayl Yüklə</span>
                  <span className="import-choice-desc">.xlsx, .csv və ya .txt faylı yükləyin və ya sürükləyin</span>
                </button>
              </div>
            </div>
          )}

          {step === 'clipboard' && (
            <div className="import-step">
              <p className="import-clipboard-hint">
                Excel-də məlumatları seçin, <b>Ctrl+C</b> edin və aşağıya yapışdırın.
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
                <div className="import-dropzone-text">{loading ? 'Fayl analiz edilir...' : 'Excel, CSV və ya TXT faylı seçin'}</div>
                <div className="import-dropzone-sub">.xlsx, .csv, .txt — sürükləyin və ya klikləyin</div>
                <input
                  ref={inputRef}
                  type="file"
                  accept=".xlsx,.csv,.txt"
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
                <div className="import-stat import-stat-new"><b>{newCount}</b><span>yeni əlavə olunacaq</span></div>
                {changedCount > 0 && (
                  <div className="import-stat import-stat-changed"><b>{changedCount}</b><span>dəyişəcək</span></div>
                )}
                {unchangedCount > 0 && (
                  <div className="import-stat import-stat-dup"><b>{unchangedCount}</b><span>dəyişiklik yoxdur</span></div>
                )}
                {emptyCount > 0 && (
                  <div className="import-stat import-stat-warn"><b>{emptyCount}</b><span>boş xana var</span></div>
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
                  <span>Sütun başlıqları tanınmadı — məlumatlar defolt sütun ardıcıllığı ilə oxundu.</span>
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
                      <tr key={i} className={`import-row-${r._STATUS.toLowerCase()}`}>
                        <td>
                          {r._STATUS === 'NEW' && <span className="import-badge-new">Yeni</span>}
                          {r._STATUS === 'CHANGED' && <span className="import-badge-changed">Dəyişəcək</span>}
                          {r._STATUS === 'UNCHANGED' && <span className="import-badge-existing">Mövcud</span>}
                        </td>
                        {TARGET_FIELDS.map((f) => {
                          const val = r[f];
                          const isEmpty = !String(val || '').trim();
                          return (
                            <td key={f} className={isEmpty ? 'import-cell-empty' : (r._STATUS === 'NEW' ? 'import-colnew' : '')}>
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
            <button className="btn-primary" disabled={newCount + changedCount === 0} onClick={handleConfirm}>
              <CheckIcon /> Təsdiqlə ({newCount + changedCount})
            </button>
          )}
        </div>
      </div>
    </div>
  );
}
