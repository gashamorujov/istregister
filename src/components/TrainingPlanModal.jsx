import { useState } from 'react';
import teachers from '../data/teachers';

export default function TrainingPlanModal({ groups, onConfirm, onCancel }) {
  const [selections, setSelections] = useState({});

  const handleSelect = (code, teacher) => {
    setSelections(prev => ({ ...prev, [code]: teacher }));
  };

  const allSelected = groups.length > 0 && groups.every(g => selections[g.courseCode]);

  return (
    <div className="modal-overlay" onClick={onCancel}>
      <div className="modal" onClick={(e) => e.stopPropagation()}>
        <div className="modal-header">
          <h2>Training Plan — Müəllim Seçimi</h2>
          <button className="modal-close" onClick={onCancel}>✕</button>
        </div>
        <div className="modal-body">
          <p className="modal-info">
            <strong>{groups.length}</strong> kurs qrupu tapıldı. Hər kurs üçün müəllim seçin:
          </p>
          <div className="course-group-list">
            {groups.map((g, idx) => (
              <div className="course-group-row" key={`${g.courseCode}-${idx}`}>
                <div className="course-group-info">
                  <strong>{g.courseName}</strong>
                  <span className="course-group-meta">
                    {g.courseCode} · {g.startDate || '—'} → {g.finishDate || '—'} · {g.studentCount} nəfər
                  </span>
                </div>
                <select className="teacher-select" value={selections[g.courseCode] || ''} onChange={(e) => handleSelect(g.courseCode, e.target.value)}>
                  <option value="">— Müəllim seçin —</option>
                  {teachers.map(t => <option key={t} value={t}>{t}</option>)}
                </select>
              </div>
            ))}
          </div>
        </div>
        <div className="modal-footer">
          <button className="btn-secondary" onClick={onCancel}>Ləğv</button>
          <button className="btn-primary" disabled={!allSelected} onClick={() => onConfirm(selections)}>Yüklə (Excel)</button>
        </div>
      </div>
    </div>
  );
}
