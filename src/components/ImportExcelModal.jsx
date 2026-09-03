import { useRef, useState } from 'react';
import ExcelJS from 'exceljs';
import { IMPORT_FIELDS } from './SpreadsheetTable';

function rowKey(row) {
  return [row.fullName, row.serial, row.idNumber, row.courseCode]
    .map(v => String(v || '').trim().toLowerCase()).join('|');
}

function parseCellText(cell) {
  if (cell === null || cell === undefined) return '';
  if (typeof cell === 'object' && cell.text) return String(cell.text).trim();
  if (cell && typeof cell === 'object' && 'result' in cell) return String(cell.result || '').trim();
  if (cell && typeof cell === 'object' && 'richText' in cell) {
    return String((cell.richText || []).map(r => r.text).join('')).trim();
  }
  if (cell && typeof cell === 'object' && 'hyperlinks' in cell) return String(cell.hyperlinks?.text || '').trim();
  return String(cell).trim();
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
      const sheet = workbook.worksheets[0];
      if (!sheet) { setError('Faylda vərəq tapılmadı.'); return; }

      const headerRow = sheet.getRow(1);
      const colMap = {};
      headerRow.eachCell({ includeEmpty: true }, (cell, colNumber) => {
        const label = parseCellText(cell.value);
        const found = IMPORT_FIELDS.find(f => f.label.toLowerCase() === label.toLowerCase() || f.field.toLowerCase() === label.toLowerCase());
        if (found) colMap[colNumber] = found.field;
      });

      const hasAnyField = Object.keys(colMap).length > 0;
      const parsed = [];
      sheet.eachRow((row, rowNumber) => {
        if (rowNumber === 1) return;
        if (hasAnyField) {
          const rec = {};
          let hasValue = false;
          Object.entries(colMap).forEach(([cn, field]) => {
            const cell = row.getCell(Number(cn));
            const v = parseCellText(cell.value);
            rec[field] = v;
            if (v) hasValue = true;
          });
          if (!hasValue) return;
          parsed.push(rec);
        } else {
          const fullName = parseCellText(row.getCell(1).value);
          if (!fullName) return;
          const rec = {
            fullName,
            serial: parseCellText(row.getCell(2).value),
            idNumber: parseCellText(row.getCell(3).value),
            birthDate: parseCellText(row.getCell(4).value),
            phone: parseCellText(row.getCell(5).value),
            email: parseCellText(row.getCell(6).value),
            rank: parseCellText(row.getCell(7).value),
            fullNameId: parseCellText(row.getCell(8).value),
            rank2: parseCellText(row.getCell(9).value),
            courseCode: parseCellText(row.getCell(10).value),
            startDate: parseCellText(row.getCell(11).value),
            finishDate: parseCellText(row.getCell(12).value),
            note: parseCellText(row.getCell(13).value),
            date: parseCellText(row.getCell(14).value),
          };
          parsed.push(rec);
        }
      });

      // Mark new vs existing
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
          <h2>📥 Excel Import</h2>
          <button className="modal-close" onClick={onCancel}>✕</button>
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
                <div className="import-dropzone-icon">📄</div>
                <div className="import-dropzone-text">Excel faylı seçin və ya buraya sürükləyin</div>
                <div className="import-dropzone-sub">.xlsx formatında — yalnız yeni məlumatlar əlavə olunur</div>
                <input
                  ref={inputRef}
                  type="file"
                  accept=".xlsx"
                  style={{ display: 'none' }}
                  onChange={(e) => handleFile(e.target.files?.[0])}
                />
              </div>
              {fileName && <div className="import-file-name">✓ {fileName}</div>}
              {error && <div className="import-parse-error">⚠️ {error}</div>}
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
              ✅ Təsdiqlə ({newCount})
            </button>
          )}
        </div>
      </div>
    </div>
  );
}
