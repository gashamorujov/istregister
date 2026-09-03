import { memo } from 'react';
import { FilterIcon } from './Icons';

// Custom AG Grid header that shows a filter icon button (visible on hover,
// always visible/active when the column has an active filter).
function GridHeader(props) {
  const colId = props.column?.getColDef?.()?.field || props.column?.getColId?.();
  const params = props.activeRef || {};
  const active = !!(params.value && colId && params.value.has(colId));

  const handleClick = (e) => {
    e.stopPropagation();
    props.onOpenFilter(colId);
  };

  return (
    <div className="header-filter-wrap">
      <span className="header-label-text">{props.displayName}</span>
      <button
        type="button"
        className={`header-filter-btn ${active ? 'active' : ''}`}
        onClick={handleClick}
        title="Sütun üzrə filtrlə"
        aria-label={`${props.displayName} üzrə filtr`}
      >
        <FilterIcon />
      </button>
    </div>
  );
}

export default memo(GridHeader);
