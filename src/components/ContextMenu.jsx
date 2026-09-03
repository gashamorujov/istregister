import { useEffect, useRef } from 'react';

export default function ContextMenu({ x, y, onClose, onTrainingPlan, onProtocol }) {
  const ref = useRef(null);

  useEffect(() => {
    const handleClick = (e) => {
      if (ref.current && !ref.current.contains(e.target)) {
        onClose();
      }
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
      <div className="context-menu-item" onClick={onTrainingPlan}>
        📋 Training Plan
      </div>
      <div className="context-menu-item disabled">
        📄 Protokol <span className="badge-soon">tezliklə</span>
      </div>
    </div>
  );
}
