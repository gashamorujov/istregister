import { useState, useRef, useEffect } from 'react';

export default function FilterPanel({ field, headerName, values, selected, onApply, onClose }) {
  const [searchText, setSearchText] = useState('');
  const [checked, setChecked] = useState(new Set(selected));
  const panelRef = useRef(null);

  useEffect(() => {
    setChecked(new Set(selected));
  }, [selected]);

  useEffect(() => {
    const handleClick = (e) => {
      if (panelRef.current && !panelRef.current.contains(e.target)) {
        onClose();
      }
    };
    document.addEventListener('mousedown', handleClick);
    return () => document.removeEventListener('mousedown', handleClick);
  }, [onClose]);

  const filteredValues = values.filter(v =>
    v.toLowerCase().includes(searchText.toLowerCase())
  );

  const toggleValue = (val) => {
    setChecked(prev => {
      const next = new Set(prev);
      if (next.has(val)) next.delete(val);
      else next.add(val);
      return next;
    });
  };

  const toggleAll = () => {
    if (checked.size === filteredValues.length) {
      setChecked(new Set());
    } else {
      setChecked(new Set(filteredValues));
    }
  };

  const handleApplyClick = () => {
    onApply(field, Array.from(checked));
  };

  return (
    <div className="filter-panel-overlay">
      <div className="filter-panel" ref={panelRef}>
        <div className="filter-panel-header">
          <h3>{headerName}</h3>
          <button className="filter-close" onClick={onClose}>×</button>
        </div>
        <div className="filter-search">
          <input
            type="text"
            placeholder="Axtar..."
            value={searchText}
            onChange={(e) => setSearchText(e.target.value)}
            autoFocus
          />
        </div>
        <div className="filter-actions">
          <label className="filter-checkbox">
            <input
              type="checkbox"
              checked={checked.size === filteredValues.length && filteredValues.length > 0}
              onChange={toggleAll}
            />
            Hamısı ({filteredValues.length})
          </label>
        </div>
        <div className="filter-values">
          {filteredValues.map(val => (
            <label className="filter-checkbox" key={val}>
              <input
                type="checkbox"
                checked={checked.has(val)}
                onChange={() => toggleValue(val)}
              />
              {val}
            </label>
          ))}
          {filteredValues.length === 0 && (
            <div className="filter-empty">Nəticə tapılmadı</div>
          )}
        </div>
        <div className="filter-panel-footer">
          <button className="btn btn-secondary" onClick={onClose}>Ləğv</button>
          <button className="btn btn-primary" onClick={handleApplyClick}>Təsdiq et</button>
        </div>
      </div>
    </div>
  );
}
