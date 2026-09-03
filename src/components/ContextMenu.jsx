import { useEffect, useRef } from 'react';

export default function ContextMenu({ x, y, onClose, onInsertAbove, onInsertBelow, onDelete }) {
  const ref = useRef(null);

  useEffect(() => {
    const handleClick = (e) => {
      if (ref.current && !ref.current.contains(e.target)) onClose();
    };
    const handleKey = (e) => {
      if (e.key === 'Escape') onClose();
    };
    document.addEventListener('mousedown', handleClick);
    document.addEventListener('keydown', handleKey);
    return () => {
      document.removeEventListener('mousedown', handleClick);
      document.removeEventListener('keydown', handleKey);
    };
  }, [onClose]);

  return (
    <div
      ref={ref}
      className="context-menu"
      style={{ left: x, top: y }}
      onContextMenu={(e) => e.preventDefault()}
    >
      <div className="context-menu-header">Sətir əməliyyatları</div>
      <div className="context-menu-item" onClick={onInsertAbove}>
        <span className="menu-icon">↑</span> Sətrin üstünə əlavə et
      </div>
      <div className="context-menu-item" onClick={onInsertBelow}>
        <span className="menu-icon">↓</span> Sətrin altına əlavə et
      </div>
      <div className="context-menu-divider" />
      <div className="context-menu-item danger" onClick={onDelete}>
        <span className="menu-icon">🗑</span> Sətri sil
      </div>
    </div>
  );
}
