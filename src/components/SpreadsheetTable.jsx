import { useCallback, useMemo, useRef, useState, useEffect } from 'react';
import { AgGridReact } from 'ag-grid-react';
import 'ag-grid-community/styles/ag-theme-quartz.css';
import { ClientSideRowModelModule, TextEditorModule, CellStyleModule, RowApiModule, ValidationModule } from 'ag-grid-community';
import ContextMenu from './ContextMenu';
import FilterPanel from './FilterPanel';
import TrainingPlanModal from './TrainingPlanModal';
import ImportExcelModal from './ImportExcelModal';
import { getUniqueCourseGroups, generateTrainingPlan } from '../lib/excelGenerator';
import { FIELD_LABELS, TARGET_FIELDS } from '../lib/importMapping';
import useFirebaseData from '../lib/useFirebaseData';
import logoUrl from '../assets/ist-logo.png?url';
import {
  SearchIcon, CloseIcon, ResetFilterIcon, ImportIcon, WarningIcon,
  UndoIcon, RedoIcon, PlusIcon, WifiOffIcon, WifiIcon,
} from './Icons';

function createEmptyRow(idx) {
  return {
    _id: `empty-${Date.now()}-${idx}`,
    fullName: '', serial: '', idNumber: '', birthDate: '',
    phone: '', email: '', rank: '', fullNameId: '', rank2: '',
    courseCode: '', startDate: '', finishDate: '', note: '', date: '',
  };
}

export default function SpreadsheetTable() {
  const {
    rows, setRows, loading, connected, lastSync,
    canUndo, canRedo,
    addRow, deleteRow, updateCell, importRows, undo, redo,
  } = useFirebaseData();

  const [searchText, setSearchText] = useState('');
  const [columnFilters, setColumnFilters] = useState({});
  const [menuState, setMenuState] = useState(null);
  const [activeFilterColumn, setActiveFilterColumn] = useState(null);
  const [modalOpen, setModalOpen] = useState(false);
  const [modalGroups, setModalGroups] = useState([]);
  const [filteredForTemplate, setFilteredForTemplate] = useState([]);
  const [importOpen, setImportOpen] = useState(false);
  const [confirmDelete, setConfirmDelete] = useState(null);
  const [toast, setToast] = useState(null);
  const gridRef = useRef(null);
  const touchTimer = useRef(null);
  const currentRef = useRef(rows);

  useEffect(() => {
    currentRef.current = rows;
  }, [rows]);

  // Toast helper
  const showToast = useCallback((msg, type = 'info') => {
    setToast({ msg, type });
    setTimeout(() => setToast(null), 3000);
  }, []);

  // Filtered data
  const filteredData = useMemo(() => {
    let data = rows;
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
  }, [rows, columnFilters, searchText]);

  // Cell value changed
  const onCellValueChanged = useCallback((event) => {
    const { data, colDef, newValue } = event;
    const field = colDef.field;
    updateCell(data._id, field, newValue ?? '');
  }, [updateCell]);

  // Context menu
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

  // Touch support
  const handleTouchStart = useCallback((e) => {
    const touch = e.touches[0];
    touchTimer.current = setTimeout(() => setMenuState({ x: 80, y: touch.clientY || 100, rowIndex: 0 }), 600);
  }, []);
  const handleTouchEnd = useCallback(() => { if (touchTimer.current) clearTimeout(touchTimer.current); }, []);

  // Row operations
  const insertRow = useCallback((position) => {
    if (!menuState) return;
    addRow(position, menuState.rowIndex);
    setMenuState(null);
  }, [menuState, addRow]);

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
    deleteRow(confirmDelete.rowIndex);
    setConfirmDelete(null);
    showToast('Sətir silindi', 'info');
  }, [confirmDelete, deleteRow, showToast]);

  // Training plan
  const handleTrainingPlan = useCallback(() => {
    if (!menuState) return;
    const gridApi = gridRef.current?.api;
    if (!gridApi) return;
    let row = null;
    gridApi.forEachNode(n => { if (n.rowIndex === menuState.rowIndex && !row) row = n.data; });
    if (!row || (!row.courseCode && !row.startDate)) {
      showToast('Bu sətirdə kurs məlumatı yoxdur', 'info');
      setMenuState(null);
      return;
    }
    const data = currentRef.current;
    const groups = getUniqueCourseGroups(data);
    if (groups.length === 0) { setMenuState(null); return; }
    setModalGroups(groups);
    setFilteredForTemplate(data);
    setModalOpen(true);
    setMenuState(null);
  }, [menuState, showToast]);

  const handleConfirmTrainingPlan = useCallback(async (entries) => {
    setModalOpen(false);
    try {
      await generateTrainingPlan(filteredForTemplate, entries, logoUrl);
      showToast('Training Plan yükləndi', 'info');
    } catch (err) {
      console.error('Training plan error:', err);
      showToast('Xəta baş verdi', 'info');
    }
  }, [filteredForTemplate, showToast]);

  // Filter handling
  const getSynchronizedValues = useCallback((field) => {
    const vals = new Set();
    rows.forEach(r => {
      const v = String(r[field] || '').trim();
      if (v) vals.add(v);
    });
    return Array.from(vals).sort();
  }, [rows]);

  const handleHeaderClick = useCallback((e) => {
    const col = e.column;
    if (!col) return;
    const field = col.colDef?.field;
    if (field) setActiveFilterColumn(field);
  }, []);

  const handleFilterApply = useCallback((field, selectedValues) => {
    setColumnFilters(prev => {
      const next = { ...prev };
      if (selectedValues.length === 0) {
        delete next[field];
      } else {
        next[field] = selectedValues;
      }
      return next;
    });
    setActiveFilterColumn(null);
  }, []);

  const resetFilters = useCallback(() => {
    setColumnFilters({});
    setSearchText('');
  }, []);

  // Import confirm
  const handleImportConfirm = useCallback((newRecords) => {
    importRows(newRecords);
    setImportOpen(false);
    showToast(`${newRecords.length} yeni sətir əlavə edildi`, 'info');
  }, [importRows, showToast]);

  // Keyboard shortcuts
  useEffect(() => {
    const handler = (e) => {
      if ((e.ctrlKey || e.metaKey) && e.key === 'z' && !e.shiftKey) {
        e.preventDefault();
        undo();
      }
      if ((e.ctrlKey || e.metaKey) && (e.key === 'y' || (e.key === 'z' && e.shiftKey))) {
        e.preventDefault();
        redo();
      }
    };
    window.addEventListener('keydown', handler);
    return () => window.removeEventListener('keydown', handler);
  }, [undo, redo]);

  // Column definitions
  const sanitize = useCallback((v) => {
    if (v === null || v === undefined) return '';
    if (typeof v === 'object') {
      if (v.richText && Array.isArray(v.richText)) return v.richText.map(rt => rt.text || '').join('');
      if (v.text !== undefined) return String(v.text);
      if (v instanceof Date && !isNaN(v.getTime())) {
        const dd = String(v.getUTCDate()).padStart(2, '0');
        const mm = String(v.getUTCMonth() + 1).padStart(2, '0');
        return dd + '.' + mm + '.' + v.getUTCFullYear();
      }
      return '';
    }
    const s = String(v);
    return s.includes('[object Object]') ? '' : s;
  }, []);

  const columnDefs = useMemo(() => [
    ...TARGET_FIELDS.map(field => ({
      headerName: FIELD_LABELS[field],
      field,
      editable: true,
      width: field === 'fullName' ? 240 : field === 'phone' ? 160 : field === 'email' ? 210 : field === 'rank' ? 180 : field === 'fullNameId' ? 220 : 130,
      minWidth: 90,
      cellEditor: 'agTextCellEditor',
      cellClass: 'editable-cell',
      valueFormatter: (params) => sanitize(params.value),
    })),
  ], [sanitize]);

  const defaultColDef = useMemo(() => ({
    sortable: true,
    resizable: true,
    filter: false,
    suppressMovable: true,
    floatingFilter: false,
  }), []);

  // Loading screen
  if (loading) {
    return (
      <div className="spreadsheet-root">
        <div className="loading-screen">
          <div className="loading-spinner" />
          <div className="loading-text">Məlumatlar yüklənir...</div>
          <div className="loading-sub">Firebase bağlantısı qurulur</div>
        </div>
      </div>
    );
  }

  return (
    <div className="spreadsheet-root">
      {/* Toolbar */}
      <div className="toolbar">
        <div className="toolbar-left">
          <span className="row-count">{filteredData.length} / {rows.length} sətir</span>
          <div className="toolbar-divider" />
          <button
            className={`btn-control ${canUndo ? '' : 'disabled'}`}
            onClick={undo}
            disabled={!canUndo}
            title="Geri al (Ctrl+Z)"
          >
            <UndoIcon />
          </button>
          <button
            className={`btn-control ${canRedo ? '' : 'disabled'}`}
            onClick={redo}
            disabled={!canRedo}
            title="İrəli get (Ctrl+Y)"
          >
            <RedoIcon />
          </button>
        </div>

        <div className="toolbar-center">
          <div className="search-box">
            <span className="search-icon"><SearchIcon /></span>
            <input
              type="text"
              className="search-input"
              placeholder="Axtar..."
              value={searchText}
              onChange={(e) => setSearchText(e.target.value)}
            />
            {searchText && (
              <button className="search-clear" onClick={() => setSearchText('')} aria-label="Axtarışı təmizlə">
                <CloseIcon />
              </button>
            )}
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
            <button className="btn-control add-row" onClick={() => addRow('below', rows.length - 1)} title="Yeni sətir əlavə et">
              <PlusIcon />
            </button>
          </div>
          <div className="toolbar-divider" />
          <div className="connection-status">
            {connected
              ? <span className="status-online"><WifiIcon /> Canlı</span>
              : <span className="status-offline"><WifiOffIcon /> Kəsildi</span>
            }
          </div>
        </div>
      </div>

      {/* Active filter chips */}
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

      {/* AG Grid */}
      <div
        className={`ag-theme-quartz grid-wrap ${loading ? 'grid-loading' : ''}`}
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
          modules={[ClientSideRowModelModule, TextEditorModule, CellStyleModule, RowApiModule, ValidationModule]}
          enableCellTextSelection={true}
          rowHeight={40}
          headerHeight={44}
          suppressRowHoverHighlight={false}
          singleClickEdit={false}
          onCellValueChanged={onCellValueChanged}
          animateRows={false}
          getRowId={params => params.data._id}
        />
      </div>

      {/* Toast notification */}
      {toast && (
        <div className={`toast toast-${toast.type}`}>
          {toast.msg}
        </div>
      )}

      {/* Context Menu */}
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

      {/* Filter Panel */}
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

      {/* Delete Confirmation */}
      {confirmDelete && (
        <div className="modal-overlay" onClick={() => setConfirmDelete(null)}>
          <div className="modal confirm-modal" onClick={e => e.stopPropagation()}>
            <div className="confirm-body">
              <div className="confirm-icon"><WarningIcon /></div>
              <div className="confirm-title">Sətir silinsin?</div>
              <div className="confirm-message">
                {confirmDelete.name
                  ? `"${confirmDelete.name}" məlumatı silinəcək.`
                  : 'Bu sətir tamamilə silinəcək.'}
              </div>
            </div>
            <div className="confirm-actions">
              <button className="btn btn-secondary" onClick={() => setConfirmDelete(null)}>Ləğv et</button>
              <button className="btn btn-danger" onClick={confirmDeleteRow}>Sil</button>
            </div>
          </div>
        </div>
      )}

      {/* Training Plan Modal */}
      {modalOpen && (
        <TrainingPlanModal
          groups={modalGroups}
          onConfirm={handleConfirmTrainingPlan}
          onCancel={() => setModalOpen(false)}
        />
      )}

      {/* Import Modal */}
      {importOpen && (
        <ImportExcelModal
          existingKeys={new Set(rows.map(r => [r.fullName, r.serial, r.idNumber, r.courseCode].map(v => String(v || '').trim().toLowerCase()).join('|')))}
          existingRows={rows}
          onConfirm={handleImportConfirm}
          onCancel={() => setImportOpen(false)}
        />
      )}
    </div>
  );
}
