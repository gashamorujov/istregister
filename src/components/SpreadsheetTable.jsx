import { useCallback, useMemo, useRef, useState } from 'react';
import { AgGridReact } from 'ag-grid-react';
import 'ag-grid-community/styles/ag-grid.css';
import 'ag-grid-community/styles/ag-theme-quartz.css';
import { ClientSideRowModelModule } from 'ag-grid-community';
import studentsData from '../data/registrData.json';
import logoUrl from '../assets/ist-logo.png?url';
import ContextMenu from './ContextMenu';
import FilterPanel from './FilterPanel';
import TrainingPlanModal from './TrainingPlanModal';
import ImportExcelModal from './ImportExcelModal';
import { getUniqueCourseGroups, generateTrainingPlan } from '../lib/excelGenerator';
import { rowKey, FIELD_LABELS } from '../lib/importMapping';
import {
  SearchIcon, CloseIcon, ResetFilterIcon, ImportIcon, WarningIcon,
} from './Icons';

function createEmptyRow(idx) {
  return {
    _id: `empty-${Date.now()}-${idx}`,
    fullName: '', serial: '', idNumber: '', birthDate: '',
    phone: '', email: '', rank: '', fullNameId: '', rank2: '',
    courseCode: '', startDate: '', finishDate: '', note: '', date: '',
  };
}

function getInitialState() {
  const data = studentsData.map((r, i) => ({ ...r, _id: `reg-${i}` }));
  for (let i = 0; i < 50; i++) data.push(createEmptyRow(i));
  return data;
}

export default function SpreadsheetTable() {
  const [rowData, setRowData] = useState(() => getInitialState());
  const [searchText, setSearchText] = useState('');
  const [columnFilters, setColumnFilters] = useState({});
  const [menuState, setMenuState] = useState(null);
  const [activeFilterColumn, setActiveFilterColumn] = useState(null);
  const [modalOpen, setModalOpen] = useState(false);
  const [modalGroups, setModalGroups] = useState([]);
  const [filteredForTemplate, setFilteredForTemplate] = useState([]);
  const [importOpen, setImportOpen] = useState(false);
  const [confirmDelete, setConfirmDelete] = useState(null);
  const gridRef = useRef(null);
  const touchTimer = useRef(null);
  const currentRef = useRef(rowData);

  const filteredData = useMemo(() => {
    let data = rowData;
    const af = Object.entries(columnFilters);
    if (af.length > 0) {
      data = data.filter(row =>
        af.every(([field, vals]) => {
          if (!vals || vals.length === 0) return true;
          return vals.some(v => String(row[field] || '').toLowerCase().includes(v.toLowerCase()));
        })
      );
    }
    if (searchText.trim()) {
      const q = searchText.toLowerCase().trim();
      data = data.filter(row =>
        Object.values(row).some(v => String(v || '').toLowerCase().includes(q))
      );
    }
    return data;
  }, [rowData, columnFilters, searchText]);

  const onCellValueChanged = useCallback((event) => {
    const { data, colDef, newValue } = event;
    const field = colDef.field;
    const newRows = currentRef.current.map(r => r._id === data._id ? { ...r, [field]: newValue ?? '' } : r);
    currentRef.current = newRows; setRowData(newRows);
  }, []);

  const onContextMenu = useCallback((e) => {
    e.preventDefault();
    const gridApi = gridRef.current?.api;
    if (!gridApi) return;
    let rowIndex = -1;
    const target = e.target.closest('.ag-row');
    if (target) {
      const rowNode = gridApi.getRowNode(target.getAttribute('row-id'));
      if (rowNode) rowIndex = rowNode.rowIndex;
    }
    if (rowIndex === -1) {
      const rect = gridRef.current?.eGridDiv?.getBoundingClientRect();
      if (rect) rowIndex = Math.floor((e.clientY - rect.top) / 40);
    }
    const x = Math.min(e.clientX, window.innerWidth - 240);
    const y = Math.min(e.clientY, window.innerHeight - 180);
    setMenuState({ x, y, rowIndex: rowIndex >= 0 ? rowIndex : currentRef.current.length - 1 });
  }, []);

  const handleTouchStart = useCallback((e) => {
    const touch = e.touches[0];
    touchTimer.current = setTimeout(() => setMenuState({ x: 80, y: touch.clientY || 100, rowIndex: 0 }), 600);
  }, []);

  const handleTouchEnd = useCallback(() => { if (touchTimer.current) clearTimeout(touchTimer.current); }, []);

  const insertRow = useCallback((position) => {
    if (!menuState) return;
    const empty = createEmptyRow(0);
    const newRows = [...currentRef.current];
    newRows.splice(position === 'above' ? menuState.rowIndex : menuState.rowIndex + 1, 0, empty);
    currentRef.current = newRows; setRowData(newRows); setMenuState(null);
  }, [menuState]);

  // Open the delete confirmation dialog for a specific row index
  const requestDelete = useCallback((rowIndex) => {
    const api = gridRef.current?.api;
    let row = null;
    if (api) {
      api.forEachNode(n => { if (n.rowIndex === rowIndex && !row) row = n.data; });
    }
    setConfirmDelete({
      rowIndex,
      name: row && row.fullName ? String(row.fullName) : '',
    });
  }, []);

  const confirmDeleteRow = useCallback(() => {
    if (!confirmDelete) return;
    const newRows = [...currentRef.current];
    if (confirmDelete.rowIndex >= 0 && confirmDelete.rowIndex < newRows.length) {
      newRows.splice(confirmDelete.rowIndex, 1);
      currentRef.current = newRows; setRowData(newRows);
    }
    setConfirmDelete(null);
  }, [confirmDelete]);

  const handleTrainingPlan = useCallback(() => {
    setMenuState(null);
    const groups = getUniqueCourseGroups(filteredData);
    if (groups.length === 0) { alert('Filtrlənmiş məlumatda kurs tapılmadı.'); return; }
    setFilteredForTemplate(filteredData); setModalGroups(groups); setModalOpen(true);
  }, [filteredData]);

  const handleConfirmTrainingPlan = useCallback(async (entries) => {
    try {
      const resp = await fetch(logoUrl);
      const logoBuffer = await resp.arrayBuffer();
      await generateTrainingPlan(filteredForTemplate, entries, logoBuffer);
      setModalOpen(false);
    } catch (err) {
      console.error('Training Plan xətası:', err);
      alert('Training Plan yaradılmasında xəta baş verdi.');
    }
  }, [filteredForTemplate]);

  // Import: append only brand-new rows (dedup by key), preserving all existing data
  const handleImportConfirm = useCallback((newRows) => {
    if (!newRows || newRows.length === 0) return;
    const existing = new Set(currentRef.current.map(r => rowKey(r)));
    const added = newRows.filter(r => !existing.has(rowKey(r))).map((r, i) => ({ ...r, _id: `import-${Date.now()}-${i}` }));
    if (added.length === 0) { setImportOpen(false); return; }
    const merged = [...currentRef.current, ...added];
    currentRef.current = merged; setRowData(merged);
    setImportOpen(false);
  }, []);

  const handleFilterApply = useCallback((field, values) => {
    setColumnFilters(prev => {
      const next = { ...prev };
      if (values && values.length > 0) next[field] = values; else delete next[field];
      return next;
    });
    setActiveFilterColumn(null);
  }, []);

  const resetFilters = useCallback(() => {
    setColumnFilters({});
    setActiveFilterColumn(null);
  }, []);

  // Synchronized unique values: computed from data filtered by ALL OTHER columns
  // (excluding the current field), so options cascade sequentially.
  const getSynchronizedValues = useCallback((field) => {
    const af = Object.entries(columnFilters).filter(([f]) => f !== field);
    let data = rowData;
    if (af.length > 0) {
      data = data.filter(row =>
        af.every(([f, vals]) => {
          if (!vals || vals.length === 0) return true;
          return vals.some(v => String(row[f] || '').toLowerCase().includes(v.toLowerCase()));
        })
      );
    }
    const vals = new Set();
    data.forEach(r => { const v = r[field]; if (v) vals.add(String(v)); });
    return Array.from(vals).sort();
  }, [rowData, columnFilters]);

  // Clicking a column header opens the filter directly (no extra button)
  const handleHeaderClick = useCallback((event) => {
    const colId = event.column?.getColId?.();
    if (colId) setActiveFilterColumn(colId);
  }, []);

  const makeCol = (field, cfg = {}) => ({
    field,
    headerName: FIELD_LABELS[field],
    editable: true,
    ...cfg,
  });

  const columnDefs = useMemo(() => [
    makeCol('fullName', { width: 260, minWidth: 150 }),
    makeCol('serial', { width: 120, minWidth: 90 }),
    makeCol('idNumber', { width: 130, minWidth: 100 }),
    makeCol('birthDate', { width: 110, minWidth: 90 }),
    makeCol('phone', { width: 175, minWidth: 120 }),
    makeCol('email', { width: 260, minWidth: 180 }),
    makeCol('rank', { width: 200, minWidth: 140 }),
    makeCol('fullNameId', { width: 260, minWidth: 150 }),
    makeCol('rank2', { width: 160, minWidth: 120 }),
    makeCol('courseCode', { width: 110, minWidth: 80 }),
    makeCol('startDate', { width: 130, minWidth: 100 }),
    makeCol('finishDate', { width: 130, minWidth: 100 }),
    makeCol('note', { width: 140, minWidth: 100 }),
    makeCol('date', { width: 120, minWidth: 90 }),
  ], []);

  const defaultColDef = useMemo(() => ({
    sortable: false, resizable: true, filter: false, suppressMovable: true, suppressMenu: true,
  }), []);

  return (
    <div className="spreadsheet-root">
      <div className="toolbar">
        <div className="toolbar-left">
          <span className="row-count">{filteredData.length} / {rowData.length} sətir</span>
        </div>

        <div className="toolbar-center">
          <div className="search-box">
            <span className="search-icon"><SearchIcon /></span>
            <input type="text" className="search-input" placeholder="Axtar..." value={searchText} onChange={(e) => setSearchText(e.target.value)} />
            {searchText && <button className="search-clear" onClick={() => setSearchText('')} aria-label="Axtarışı təmizlə"><CloseIcon /></button>}
          </div>
        </div>

        <div className="toolbar-right">
          <div className="control-group">
            <button
              className={`btn-control reset ${Object.keys(columnFilters).length > 0 ? 'active' : ''}`}
              onClick={resetFilters}
              disabled={Object.keys(columnFilters).length === 0}
              title="Filtirləri sıfırla"
            >
              <ResetFilterIcon />
            </button>
            <button className="btn-control import" onClick={() => setImportOpen(true)} title="Excel-dən yeni məlumat idxal et">
              <ImportIcon /> Import
            </button>
          </div>
        </div>
      </div>

      {Object.keys(columnFilters).length > 0 && (
        <div className="active-filters-bar">
          {Object.entries(columnFilters).map(([field, values]) => (
            <span className="filter-chip" key={field}>
              {FIELD_LABELS[field] || field}: {values.length}
              <button onClick={() => setColumnFilters(prev => { const n = { ...prev }; delete n[field]; return n; })} aria-label={`${field} filtrini sil`}>
                <CloseIcon />
              </button>
            </span>
          ))}
        </div>
      )}

      <div className="ag-theme-quartz grid-wrap"
        onContextMenu={onContextMenu}
        onTouchStart={handleTouchStart}
        onTouchEnd={handleTouchEnd}
        onTouchMove={handleTouchEnd}
      >
        <AgGridReact
          ref={gridRef}
          rowData={filteredData}
          columnDefs={columnDefs}
          defaultColDef={defaultColDef}
          modules={[ClientSideRowModelModule]}
          enableCellTextSelection={true}
          rowHeight={40}
          headerHeight={44}
          suppressRowHoverHighlight={false}
          singleClickEdit={true}
          onCellValueChanged={onCellValueChanged}
          onColumnHeaderClicked={handleHeaderClick}
          animateRows={false}
        />
      </div>

      {menuState && (
        <ContextMenu
          x={menuState.x} y={menuState.y}
          onClose={() => setMenuState(null)}
          onInsertAbove={() => insertRow('above')}
          onInsertBelow={() => insertRow('below')}
          onDelete={() => requestDelete(menuState.rowIndex)}
          onTrainingPlan={handleTrainingPlan}
        />
      )}

      {activeFilterColumn && (
        <FilterPanel
          field={activeFilterColumn}
          headerName={FIELD_LABELS[activeFilterColumn] || activeFilterColumn}
          values={getSynchronizedValues(activeFilterColumn)}
          selected={columnFilters[activeFilterColumn] || []}
          onApply={handleFilterApply}
          onClose={() => setActiveFilterColumn(null)}
        />
      )}

      {confirmDelete && (
        <div className="modal-overlay" onClick={() => setConfirmDelete(null)}>
          <div className="modal confirm-modal" onClick={e => e.stopPropagation()}>
            <div className="confirm-body">
              <div className="confirm-icon"><WarningIcon /></div>
              <div className="confirm-title">Sətir silinsin?</div>
              <div className="confirm-message">
                {confirmDelete.name
                  ? `"${confirmDelete.name}" məlumatı silinəcək.`
                  : 'Bu sətir tamamilə silinəcək.'} Bu əməliyyat geri qaytarıla bilməz.
              </div>
            </div>
            <div className="confirm-actions">
              <button className="btn btn-secondary" onClick={() => setConfirmDelete(null)}>Ləğv et</button>
              <button className="btn btn-danger" onClick={confirmDeleteRow}>Sil</button>
            </div>
          </div>
        </div>
      )}

      {modalOpen && (
        <TrainingPlanModal
          groups={modalGroups}
          onConfirm={handleConfirmTrainingPlan}
          onCancel={() => setModalOpen(false)}
        />
      )}

      {importOpen && (
        <ImportExcelModal
          existingKeys={new Set(currentRef.current.map(r => rowKey(r)))}
          onConfirm={handleImportConfirm}
          onCancel={() => setImportOpen(false)}
        />
      )}
    </div>
  );
}
