import { useEffect, useRef } from 'react';

export default function ContextMenu({ x, y, onClose, onInsertAbove, onInsertBelow, onDelete, onTrainingPlan }) {
  const ref = useRef(null);

  useEffect(() => {
    const h = (e) => { if (ref.current && !ref.current.contains(e.target)) onClose(); };
    const k = (e) => { if (e.key === 'Escape') onClose(); };
    document.addEventListener('mousedown', h);
    document.addEventListener('keydown', k);
    return () => { document.removeEventListener('mousedown', h); document.removeEventListener('keydown', k); };
  }, [onClose]);

  return (
    <div ref={ref} className="context-menu" style={{ left: x, top: y }} onContextMenu={(e) => e.preventDefault()}>
      <div className="context-menu-header">Əməliyyatlar</div>
      <div className="context-menu-item" onClick={onInsertAbove}>
        <span className="menu-icon">↑</span> Sətrin üstünə əlavə et
      </div>
      <div className="context-menu-item" onClick={onInsertBelow}>
        <span className="menu-icon">↓</span> Sətrin altına əlavə et
      </div>
      <div className="context-menu-divider" />
      <div className="context-menu-item highlight" onClick={onTrainingPlan}>
        <span className="menu-icon">📋</span> Training Plan
      </div>
      <div className="context-menu-divider" />
      <div className="context-menu-item danger" onClick={onDelete}>
        <span className="menu-icon">🗑</span> Sətri sil
      </div>
    </div>
  );
}
