import { useState } from 'react';

export default function TrainingPlanModal({ groups, onConfirm, onCancel }) {
  const [groupNumbers, setGroupNumbers] = useState({});

  const handleGroupNum = (code, val) => {
    setGroupNumbers(prev => ({ ...prev, [code]: val }));
  };

  const allSelected = groups.length > 0 && groups.every(g =>
    groupNumbers[g.courseCode] && groupNumbers[g.courseCode].trim() !== ''
  );

  return (
    <div className="modal-overlay" onClick={onCancel}>
      <div className="modal" onClick={(e) => e.stopPropagation()}>
        <div className="modal-header">
          <h2>Training Plan — Qrup Nömrəsi</h2>
          <button className="modal-close" onClick={onCancel}>✕</button>
        </div>
        <div className="modal-body">
          <p className="modal-info">
            <strong>{groups.length}</strong> kurs qrupu tapıldı. Hər kurs üçün <strong>qrup nömrəsi</strong> yazın:
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
                <input
                  className="group-number-input"
                  type="text"
                  placeholder="Qrup nömrəsi (məs. 006/26 GST)"
                  value={groupNumbers[g.courseCode] || ''}
                  onChange={(e) => handleGroupNum(g.courseCode, e.target.value)}
                />
              </div>
            ))}
          </div>
        </div>
        <div className="modal-footer">
          <button className="btn-secondary" onClick={onCancel}>Ləğv</button>
          <button className="btn-primary" disabled={!allSelected} onClick={() => onConfirm(groupNumbers)}>
            Yüklə (Excel)
          </button>
        </div>
      </div>
    </div>
  );
}
