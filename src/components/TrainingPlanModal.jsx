import { useState } from 'react';
import teachers from '../data/teachers';
import { CloseIcon, ImportIcon } from './Icons';

export default function TrainingPlanModal({ groups, onConfirm, onCancel }) {
  const [entries, setEntries] = useState({});

  const update = (code, field, val) => {
    setEntries(prev => ({
      ...prev,
      [code]: { ...(prev[code] || {}), [field]: val }
    }));
  };

  const allValid = groups.length > 0 && groups.every(g => {
    const e = entries[g.courseCode] || {};
    return e.teacher && e.teacher.trim() !== '' && e.groupNum && e.groupNum.trim() !== '';
  });

  const handleConfirm = () => {
    if (!allValid) return;
    onConfirm(entries);
  };

  return (
    <div className="modal-overlay" onClick={onCancel}>
      <div className="modal modal-wide" onClick={e => e.stopPropagation()}>
        <div className="modal-header">
          <h2>Training Plan - Yükləmə</h2>
          <button className="modal-close" onClick={onCancel} aria-label="Bağla"><CloseIcon /></button>
        </div>
        <div className="modal-body">
          <p className="modal-info">
            <strong>{groups.length}</strong> kurs qrupu tapıldı. Hər kurs üçün <strong>müəllim</strong> və <strong>qrup nömrəsi</strong> daxil edin:
          </p>
          <div className="course-group-list">
            {groups.map((g, idx) => {
              const e = entries[g.courseCode] || {};
              const hasTeacher = e.teacher && e.teacher.trim() !== '';
              const hasGroup = e.groupNum && e.groupNum.trim() !== '';
              return (
                <div className="course-group-row" key={`${g.courseCode}-${idx}`}>
                  <div className="course-group-info">
                    <strong>{g.courseName}</strong>
                    <span className="course-group-meta">
                      {g.courseCode} · {g.startDate || '—'} → {g.finishDate || '—'} · {g.studentCount} nəfər
                    </span>
                  </div>
                  <div className="tp-fields">
                    <div className="tp-field">
                      <label className="tp-label">Qrup nömrəsi *</label>
                      <input
                        className="tp-input"
                        type="text"
                        placeholder="006/26 GST"
                        value={e.groupNum || ''}
                        onChange={ev => update(g.courseCode, 'groupNum', ev.target.value)}
                      />
                    </div>
                    <div className="tp-field">
                      <label className="tp-label">Müəllim *</label>
                      <select
                        className={`tp-select ${hasTeacher ? '' : 'tp-required'}`}
                        value={e.teacher || ''}
                        onChange={ev => update(g.courseCode, 'teacher', ev.target.value)}
                      >
                        <option value="">— Müəllim seçin (məcburi) —</option>
                        {teachers.map(t => <option key={t} value={t}>{t}</option>)}
                      </select>
                    </div>
                  </div>
                </div>
              );
            })}
          </div>
        </div>
        <div className="modal-footer">
          <button className="btn-secondary" onClick={onCancel}>Ləğv</button>
          <button className="btn-primary" disabled={!allValid} onClick={handleConfirm}>
            <ImportIcon /> Yüklə (Excel)
          </button>
        </div>
      </div>
    </div>
  );
}

