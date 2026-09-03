import { useCallback, useMemo, useRef, useState, useEffect } from 'react';
import { AgGridReact } from 'ag-grid-react';
import 'ag-grid-community/styles/ag-grid.css';
import 'ag-grid-community/styles/ag-theme-quartz.css';
import { ClientSideRowModelModule } from 'ag-grid-community';
import studentsData from '../data/registrData.json';
import templateXlsx from '../data/Training_plan_template.xlsx?url';
import ContextMenu from './ContextMenu';
import FilterPanel from './FilterPanel';
import TrainingPlanModal from './TrainingPlanModal';
import { getUniqueCourseGroups, generateTrainingPlan } from '../lib/excelGenerator';

function createEmptyRow(idx) {
  return {
    _id: `empty-${Date.now()}-${idx}`,
    _no: 0,
    fullName: '', serial: '', idNumber: '', birthDate: '',
    phone: '', email: '', rank: '', fullNameId: '', rank2: '',
    courseCode: '', startDate: '', finishDate: '', note: '', date: '',
  };
}

function renumber(rows) {
  return rows.map((row, i) => ({ ...row, _no: i + 1 }));
}

function getInitialState() {
  const data = studentsData.map((r, i) => ({ ...r, _no: i + 1, _id: `reg-${i}` }));
  for (let i = 0; i < 50; i++) data.push(createEmptyRow(i));
  return data;
}

export default function SpreadsheetTable() {
  const [rowData, setRowData] = useState(() => getInitialState());
  const [searchText, setSearchText] = useState('');
  const [debouncedSearch, setDebouncedSearch] = useState('');
  const [columnFilters, setColumnFilters] = useState({});
  const [menuState, setMenuState] = useState(null);
  const [activeFilterColumn, setActiveFilterColumn] = useState(null);
  const [history, setHistory] = useState({ past: [], future: [] });
  const [modalOpen, setModalOpen] = useState(false);
  const [modalGroups, setModalGroups] = useState([]);
  const [filteredForTemplate, setFilteredForTemplate] = useState([]);
  const gridRef = useRef(null);
  const touchTimer = useRef(null);
  const currentRef = useRef(rowData);

  useEffect(() => {
    const t = setTimeout(() => setDebouncedSearch(searchText), 250);
    return () => clearTimeout(t);
  }, [searchText]);

  const filteredData = useMemo(() => {
    let data = rowData;
    const af = Object.entries(columnFilters);
    if (af.length > 0) {
      data = data.filter(row =>
        af.every(([field, vals]) => {
          if (!vals || vals.length === 0) return true;
          const cv = String(row[field] || '').toLowerCase();
          return vals.some(v => cv.includes(v.toLowerCase()));
        })
      );
    }
    if (debouncedSearch.trim()) {
      const q = debouncedSearch.toLowerCase().trim();
      data = data.filter(row =>
        Object.values(row).some(v => String(v || '').toLowerCase().includes(q))
      );
    }
    return data;
  }, [rowData, columnFilters, debouncedSearch]);

  const pushHistory = useCallback(() => {
    setHistory(prev => ({
      past: [...prev.past.slice(-99), [...currentRef.current]],
      future: [],
    }));
  }, []);

  const undo = useCallback(() => {
    setHistory(prev => {
      if (prev.past.length === 0) return prev;
      const previous = prev.past[prev.past.length - 1];
      setRowData(previous);
      currentRef.current = previous;
      return { past: prev.past.slice(0, -1), future: [[...previous], ...prev.future] };
    });
  }, []);

  const redo = useCallback(() => {
    setHistory(prev => {
      if (prev.future.length === 0) return prev;
      const next = prev.future[0];
      setRowData(next);
      currentRef.current = next;
      return { past: [...prev.past, [...currentRef.current]], future: prev.future.slice(1) };
    });
  }, []);

  useEffect(() => {
    const h = (e) => {
      if ((e.ctrlKey || e.metaKey) && e.key === 'z' && !e.shiftKey) { e.preventDefault(); undo(); }
      if ((e.ctrlKey || e.metaKey) && (e.key === 'y' || (e.key === 'z' && e.shiftKey))) { e.preventDefault(); redo(); }
    };
    window.addEventListener('keydown', h);
    return () => window.removeEventListener('keydown', h);
  }, [undo, redo]);

  const onCellValueChanged = useCallback((event) => {
    const { data, colDef, newValue } = event;
    const field = colDef.field;
    const newRows = currentRef.current.map(r => r._id === data._id ? { ...r, [field]: newValue ?? '' } : r);
    currentRef.current = newRows;
    setRowData(newRows);
    setHistory(prev => ({ past: [...prev.past.slice(-99), [...newRows]], future: [] }));
  }, []);

  const onContextMenu = useCallback((e) => {
    e.preventDefault();
    const gridApi = gridRef.current?.api;
    if (!gridApi) return;
    let rowIndex = -1;
    const target = e.target.closest('.ag-row');
    if (target) {
      const rowId = target.getAttribute('row-id');
      const rowNode = gridApi.getRowNode(rowId);
      if (rowNode) rowIndex = rowNode.rowIndex;
    }
    if (rowIndex === -1) {
      const rect = gridRef.current?.eGridDiv?.getBoundingClientRect();
      if (rect) rowIndex = Math.floor((e.clientY - rect.top) / 36);
    }
    const x = Math.min(e.clientX, window.innerWidth - 240);
    const y = Math.min(e.clientY, window.innerHeight - 180);
    setMenuState({ x, y, rowIndex: rowIndex >= 0 ? rowIndex : currentRef.current.length - 1 });
  }, []);

  const handleTouchStart = useCallback((e) => {
    const touch = e.touches[0];
    touchTimer.current = setTimeout(() => {
      setMenuState({ x: 80, y: touch.clientY || 100, rowIndex: 0 });
    }, 600);
  }, []);

  const handleTouchEnd = useCallback(() => {
    if (touchTimer.current) clearTimeout(touchTimer.current);
  }, []);

  const insertRow = useCallback((position) => {
    if (!menuState) return;
    pushHistory();
    const empty = createEmptyRow(0);
    const newRows = [...currentRef.current];
    const insertAt = position === 'above' ? menuState.rowIndex : menuState.rowIndex + 1;
    newRows.splice(insertAt, 0, empty);
    const renumbered = renumber(newRows);
    currentRef.current = renumbered;
    setRowData(renumbered);
    setMenuState(null);
    setHistory(prev => ({ past: [...prev.past.slice(-99), [...renumbered]], future: [] }));
  }, [menuState, pushHistory]);

  const deleteRow = useCallback(() => {
    if (!menuState) return;
    pushHistory();
    const newRows = [...currentRef.current];
    if (menuState.rowIndex >= 0 && menuState.rowIndex < newRows.length) {
      newRows.splice(menuState.rowIndex, 1);
      const renumbered = renumber(newRows);
      currentRef.current = renumbered;
      setRowData(renumbered);
    }
    setMenuState(null);
    setHistory(prev => ({ past: [...prev.past.slice(-99), [...currentRef.current]], future: [] }));
  }, [menuState, pushHistory]);

  const handleTrainingPlan = useCallback(async () => {
    setMenuState(null);
    const groups = getUniqueCourseGroups(filteredData);
    if (groups.length === 0) {
      alert('Filtrlənmiş məlumatda kurs tapılmadı.');
      return;
    }
    setFilteredForTemplate(filteredData);
    setModalGroups(groups);
    setModalOpen(true);
  }, [filteredData]);

  const handleConfirmTrainingPlan = useCallback(async (selections) => {
    try {
      const resp = await fetch(templateXlsx);
      const buf = await resp.arrayBuffer();
      await generateTrainingPlan(filteredForTemplate, selections, buf);
      setModalOpen(false);
    } catch (err) {
      console.error('Training Plan xətası:', err);
      alert('Training Plan yaradılmasında xəta baş verdi.');
    }
  }, [filteredForTemplate]);

  const handleFilterApply = useCallback((field, values) => {
    setColumnFilters(prev => {
      const next = { ...prev };
      if (values && values.length > 0) next[field] = values;
      else delete next[field];
      return next;
    });
    setActiveFilterColumn(null);
  }, []);

  const getUniqueValues = useCallback((field) => {
    const vals = new Set();
    rowData.forEach(r => {
      const v = r[field];
      if (v !== null && v !== undefined && v !== '') vals.add(String(v));
    });
    return Array.from(vals).sort();
  }, [rowData]);

  const handleColumnHeaderDblClick = useCallback((event) => {
    const colId = event.column?.getId?.();
    if (colId && colId !== '_no') setActiveFilterColumn(colId);
  }, []);

  const canUndo = history.past.length > 0;
  const canRedo = history.future.length > 0;

  const columnDefs = useMemo(() => [
    { field: '_no', headerName: '№', width: 55, editable: false, pinned: 'left', cellClass: 'cell-readonly cell-center cell-no', headerClass: 'header-no' },
    { field: 'fullName', headerName: 'Soyad, Ad və Ata adı', width: 250, editable: true },
    { field: 'serial', headerName: 'Seriya nömrəsi', width: 115, editable: true },
    { field: 'idNumber', headerName: 'Fərdi ID nömrəsi', width: 120, editable: true },
    { field: 'birthDate', headerName: 'Doğum tarixi', width: 105, editable: true },
    { field: 'phone', headerName: 'Telefon', width: 160, editable: true },
    { field: 'email', headerName: 'Email', width: 240, editable: true },
    { field: 'rank', headerName: 'Rank (Working Diploma)', width: 190, editable: true },
    { field: 'fullNameId', headerName: 'Full Name (ID)', width: 240, editable: false, cellClass: 'cell-readonly' },
    { field: 'rank2', headerName: 'Rank / Vəzifə', width: 150, editable: true },
    { field: 'courseCode', headerName: 'Course Code', width: 105, editable: true },
    { field: 'startDate', headerName: 'Başlama tarixi', width: 120, editable: true },
    { field: 'finishDate', headerName: 'Bitmə tarixi', width: 120, editable: true },
    { field: 'note', headerName: 'Qeyd', width: 130, editable: true },
    { field: 'date', headerName: 'Tarix', width: 90, editable: true },
  ], []);

  const defaultColDef = useMemo(() => ({
    sortable: false, resizable: true, filter: false, suppressMovable: true, suppressMenu: true,
  }), []);

  return (
    <div className="spreadsheet-root">
      <div className="toolbar">
        <div className="toolbar-left">
          <div className="search-box">
            <span className="search-icon">🔍</span>
            <input type="text" className="search-input" placeholder="Axtar..." value={searchText} onChange={(e) => setSearchText(e.target.value)} />
            {searchText && <button className="search-clear" onClick={() => setSearchText('')}>✕</button>}
          </div>
        </div>
        <div className="toolbar-center">
          <div className="undo-redo-group">
            <button className="btn-toolbar" onClick={undo} disabled={!canUndo} title="Geri (Ctrl+Z)">↩</button>
            <button className="btn-toolbar" onClick={redo} disabled={!canRedo} title="İrəli (Ctrl+Y)">↪</button>
          </div>
        </div>
        <div className="toolbar-right">
          <span className="row-count">{filteredData.length} / {rowData.length} sətir</span>
          {Object.keys(columnFilters).length > 0 && (
            <button className="btn-clear-filters" onClick={() => setColumnFilters({})}>✕ Filtri təmizlə</button>
          )}
        </div>
      </div>

      {Object.keys(columnFilters).length > 0 && (
        <div className="active-filters-bar">
          {Object.entries(columnFilters).map(([field, values]) => {
            const labels = { _no:'№', fullName:'Soyad, Ad', serial:'Seriya', idNumber:'ID', birthDate:'Doğum', phone:'Telefon', email:'Email', rank:'Rank', fullNameId:'Full Name', rank2:'Vəzifə', courseCode:'Kurs kodu', startDate:'Başlama', finishDate:'Bitmə', note:'Qeyd', date:'Tarix' };
            return (
              <span className="filter-chip" key={field}>
                {labels[field] || field}: {values.length}
                <button onClick={() => setColumnFilters(prev => { const n = { ...prev }; delete n[field]; return n; })}>✕</button>
              </span>
            );
          })}
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
          rowHeight={36}
          headerHeight={40}
          suppressRowHoverHighlight={false}
          editType="fullRow"
          singleClickEdit={true}
          onCellValueChanged={onCellValueChanged}
          onColumnHeaderClicked={handleColumnHeaderDblClick}
          stopEditingWhenCellsLoseFocus={true}
          suppressCellFocus={false}
          animateRows={false}
        />
      </div>

      {menuState && (
        <ContextMenu
          x={menuState.x}
          y={menuState.y}
          onClose={() => setMenuState(null)}
          onInsertAbove={() => insertRow('above')}
          onInsertBelow={() => insertRow('below')}
          onDelete={deleteRow}
          onTrainingPlan={handleTrainingPlan}
        />
      )}

      {activeFilterColumn && (
        <FilterPanel
          field={activeFilterColumn}
          headerName={activeFilterColumn}
          values={getUniqueValues(activeFilterColumn)}
          selected={columnFilters[activeFilterColumn] || []}
          onApply={handleFilterApply}
          onClose={() => setActiveFilterColumn(null)}
        />
      )}

      {modalOpen && (
        <TrainingPlanModal
          groups={modalGroups}
          onConfirm={handleConfirmTrainingPlan}
          onCancel={() => setModalOpen(false)}
        />
      )}
    </div>
  );
}
