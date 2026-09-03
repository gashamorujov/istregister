import { useState } from 'react';
import teachers from '../data/teachers';

export default function TrainingPlanModal({ groups, onConfirm, onCancel }) {
  const [selections, setSelections] = useState({});

  const handleSelect = (courseCode, teacher) => {
    setSelections(prev => ({ ...prev, [courseCode]: teacher }));
  };

  const allSelected = groups.every(g => selections[g.courseCode]);

  return (
    <div className="modal-overlay" onClick={onCancel}>
      <div className="modal" onClick={(e) => e.stopPropagation()}>
        <div className="modal-header">
          <h2>Training Plan - Müəllim Seçimi</h2>
          <button className="modal-close" onClick={onCancel}>×</button>
        </div>
        <div className="modal-body">
          <p className="modal-info">
            Filtr edilmiş məlumatlarda <strong>{groups.length}</strong> kurs qrupu tapıldı. 
            Hər kurs üçün müəllim seçin:
          </p>
          <div className="course-group-list">
            {groups.map((group, idx) => (
              <div className="course-group-row" key={`${group.courseCode}-${idx}`}>
                <div className="course-group-info">
                  <strong>{group.courseName}</strong>
                  <span className="course-group-meta">
                    Kod: {group.courseCode} | Tarix: {group.startDate || '—'} - {group.finishDate || '—'} | Tələbə: {group.studentCount} nəfər
                  </span>
                </div>
                <select
                  className="teacher-select"
                  value={selections[group.courseCode] || ''}
                  onChange={(e) => handleSelect(group.courseCode, e.target.value)}
                >
                  <option value="">— Müəllim seçin —</option>
                  {teachers.map(t => (
                    <option key={t} value={t}>{t}</option>
                  ))}
                </select>
              </div>
            ))}
          </div>
        </div>
        <div className="modal-footer">
          <button className="btn btn-secondary" onClick={onCancel}>Ləğv et</button>
          <button
            className="btn btn-primary"
            disabled={!allSelected}
            onClick={() => onConfirm(selections)}
          >
            Yüklə (Excel)
          </button>
        </div>
      </div>
    </div>
  );
}
