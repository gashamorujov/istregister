import { useRef, useState } from 'react';
import ExcelJS from 'exceljs';
import { IMPORT_FIELDS } from './SpreadsheetTable';
import { CloseIcon, DocIcon, CheckIcon } from './Icons';

function rowKey(row) {
  return [row.fullName, row.serial, row.idNumber, row.courseCode]
    .map(v => String(v || '').trim().toLowerCase()).join('|');
}

function parseCellText(cell) {
  if (cell === null || cell === undefined) return '';
  const v = cell.value;
  if (v === null || v === undefined) return '';

  // Date object - format as dd.mm.yyyy
  if (v instanceof Date && !isNaN(v.getTime())) {
    const d = String(v.getUTCDate()).padStart(2, '0');
    const m = String(v.getUTCMonth() + 1).padStart(2, '0');
    const y = v.getUTCFullYear();
    return d + '.' + m + '.' + y;
  }

  // Rich text
  if (v.richText && Array.isArray(v.richText)) {
    return v.richText.map(rt => rt.text || '').join('').trim();
  }

  // Formula result
  if (v.result !== undefined && v.result !== null) {
    if (v.result instanceof Date && !isNaN(v.result.getTime())) {
      const d = String(v.result.getUTCDate()).padStart(2, '0');
      const m = String(v.result.getUTCMonth() + 1).padStart(2, '0');
      return d + '.' + m + '.' + v.result.getUTCFullYear();
    }
    return String(v.result).trim();
  }

  // Hyperlink
  if (v.text) return String(v.text).trim();

  return String(v).trim();
}

function matchField(label) {
  const low = label.toLowerCase().trim();
  return IMPORT_FIELDS.find(f => {
    const fLow = f.label.toLowerCase();
    return low === fLow || low.includes(fLow) || fLow.includes(low);
  });
}

export default function ImportExcelModal({ existingKeys, onConfirm, onCancel }) {
  const [fileName, setFileName] = useState('');
  const [previewRows, setPreviewRows] = useState(null);
  const [newCount, setNewCount] = useState(0);
  const [skippedCount, setSkippedCount] = useState(0);
  const [error, setError] = useState('');
  const [dragOver, setDragOver] = useState(false);
  const inputRef = useRef(null);

  const handleFile = async (file) => {
    setError('');
    setPreviewRows(null);
    setNewCount(0);
    setSkippedCount(0);
    if (!file) return;
    setFileName(file.name);
    try {
      const workbook = new ExcelJS.Workbook();
      await workbook.xlsx.load(await file.arrayBuffer());

      // Find "REGİSTR-2026" sheet, fallback to first sheet
      let sheet = workbook.getWorksheet('REGİSTR-2026');
      if (!sheet) sheet = workbook.worksheets[0];
      if (!sheet) { setError('Faylda vərəq tapılmadı.'); return; }

      // Find header row - check first 3 rows for row with matching labels
      let headerRowNum = 1;
      for (let r = 1; r <= 3; r++) {
        const testRow = sheet.getRow(r);
        let matchCount = 0;
        testRow.eachCell({ includeEmpty: true }, (cell) => {
          const label = parseCellText(cell);
          if (label && matchField(label)) matchCount++;
        });
        if (matchCount >= 3) { headerRowNum = r; break; }
      }

      const headerRow = sheet.getRow(headerRowNum);
      const colMap = {};
      headerRow.eachCell({ includeEmpty: true }, (cell, colNumber) => {
        const label = parseCellText(cell);
        const found = matchField(label);
        if (found) colMap[colNumber] = found.field;
      });

      if (Object.keys(colMap).length === 0) {
        setError('Başlıqlar tanınmadı. Düzgün Excel faylı seçdiyinizə əmin olun.');
        return;
      }

      const parsed = [];
      sheet.eachRow((row, rowNumber) => {
        if (rowNumber <= headerRowNum) return;
        const rec = {};
        let hasValue = false;
        Object.entries(colMap).forEach(([cn, field]) => {
          const v = parseCellText(row.getCell(Number(cn)));
          rec[field] = v;
          if (v) hasValue = true;
        });
        if (!hasValue) return;
        parsed.push(rec);
      });

      const result = parsed.map(rec => ({ ...rec, IS_NEW: !existingKeys.has(rowKey(rec)) }));
      const newRows = result.filter(r => r.IS_NEW);
      setNewCount(newRows.length);
      setSkippedCount(result.length - newRows.length);
      setPreviewRows(result);
    } catch (err) {
      console.error('Excel parse xətası:', err);
      setError('Excel faylı oxunarkən xəta baş verdi. Düzgün .xlsx fayl seçdiyinizə əmin olun.');
    }
  };

  const handleDrop = (e) => {
    e.preventDefault();
    setDragOver(false);
    const file = e.dataTransfer.files?.[0];
    if (file) handleFile(file);
  };

  const handleConfirm = () => {
    const newRows = (previewRows || [])
      .filter(r => r.IS_NEW)
      .map(r => {
        const copy = { ...r };
        delete copy.IS_NEW;
        return copy;
      });
    onConfirm(newRows);
  };

  return (
    <div className="modal-overlay" onClick={onCancel}>
      <div className="modal" style={{ maxWidth: 820 }} onClick={e => e.stopPropagation()}>
        <div className="modal-header">
          <h2>Excel Import</h2>
          <button className="modal-close" onClick={onCancel} aria-label="Bağla"><CloseIcon /></button>
        </div>
        <div className="modal-body">
          {!previewRows ? (
            <>
              <div
                className={`import-dropzone ${dragOver ? 'drag' : ''}`}
                onClick={() => inputRef.current?.click()}
                onDragOver={(e) => { e.preventDefault(); setDragOver(true); }}
                onDragLeave={() => setDragOver(false)}
                onDrop={handleDrop}
              >
                <div className="import-dropzone-icon"><DocIcon /></div>
                <div className="import-dropzone-text">Excel faylı seçin və ya buraya sürükləyin</div>
                <div className="import-dropzone-sub">.xlsx formatında — yalnız REGİSTR-2026 səhifəsi oxunacaq</div>
                <input
                  ref={inputRef}
                  type="file"
                  accept=".xlsx"
                  style={{ display: 'none' }}
                  onChange={(e) => handleFile(e.target.files?.[0])}
                />
              </div>
              {fileName && <div className="import-file-name"><CheckIcon /> {fileName}</div>}
              {error && <div className="import-parse-error"><CloseIcon /> {error}</div>}
            </>
          ) : (
            <>
              <p className="import-summary">
                <b>{newCount}</b> yeni məlumat əlavə olunacaq,
                {skippedCount > 0 ? <> <b style={{ color: '#b45309' }}>{skippedCount}</b> məlumat artıq mövcuddur və atlanacaq.</> : ' artıq mövcud məlumat yoxdur.'}
              </p>
              <div className="import-preview-table-wrap">
                <table className="import-preview-table">
                  <thead>
                    <tr>
                      <th>Status</th>
                      <th>Soyad, Ad və Ata adı</th>
                      <th>Seriya</th>
                      <th>Fərdi ID</th>
                      <th>Course Code</th>
                      <th>Başlama</th>
                      <th>Bitmə</th>
                    </tr>
                  </thead>
                  <tbody>
                    {previewRows.length === 0 && (
                      <tr><td colSpan={7} className="import-empty">Faylda heç bir məlumat tapılmadı.</td></tr>
                    )}
                    {previewRows.map((r, i) => (
                      <tr key={i}>
                        <td>
                          {r.IS_NEW
                            ? <span className="import-badge-new">Yeni</span>
                            : <span className="import-badge-new" style={{ background: '#fef3c7', color: '#92400e' }}>Mövcud</span>}
                        </td>
                        <td className={r.IS_NEW ? 'import-colnew' : ''}>
                          {r.fullName || <span className="import-no-field">—</span>}
                        </td>
                        <td>{r.serial || '—'}</td>
                        <td>{r.idNumber || '—'}</td>
                        <td>{r.courseCode || '—'}</td>
                        <td>{r.startDate || '—'}</td>
                        <td>{r.finishDate || '—'}</td>
                      </tr>
                    ))}
                  </tbody>
                </table>
              </div>
            </>
          )}
        </div>
        <div className="modal-footer">
          <button className="btn-secondary" onClick={() => { if (previewRows) { setPreviewRows(null); setFileName(''); } else onCancel(); }}>
            {previewRows ? 'Geri' : 'Ləğv'}
          </button>
          {previewRows && (
            <button className="btn-primary" disabled={newCount === 0} onClick={handleConfirm}>
              <CheckIcon /> Təsdiqlə ({newCount})
            </button>
          )}
        </div>
      </div>
    </div>
  );
}
