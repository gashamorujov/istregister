import { useCallback, useMemo, useRef, useState, useEffect } from 'react';
import { AgGridReact } from 'ag-grid-react';
import 'ag-grid-community/styles/ag-grid.css';
import 'ag-grid-community/styles/ag-theme-quartz.css';
import { ClientSideRowModelModule } from 'ag-grid-community';
import studentsData from '../data/registrData.json';
import ContextMenu from './ContextMenu';
import FilterPanel from './FilterPanel';

// Empty row template
function createEmptyRow(idx) {
  return {
    _id: `empty-${Date.now()}-${idx}`,
    _no: 0,
    fullName: '',
    serial: '',
    idNumber: '',
    birthDate: '',
    phone: '',
    email: '',
    rank: '',
    fullNameId: '',
    rank2: '',
    courseCode: '',
    startDate: '',
    finishDate: '',
    note: '',
    date: '',
  };
}

// Renumber _no field
function renumber(rows) {
  return rows.map((row, i) => ({ ...row, _no: i + 1 }));
}

// Initial state: original data + 50 empty rows
function getInitialState() {
  const data = studentsData.map((r, i) => ({ ...r, _no: i + 1, _id: `reg-${i}` }));
  for (let i = 0; i < 50; i++) {
    data.push(createEmptyRow(i));
  }
  return data;
}

const EDITABLE_FIELDS = [
  'fullName', 'serial', 'idNumber', 'birthDate', 'phone',
  'email', 'rank', 'rank2', 'courseCode', 'startDate',
  'finishDate', 'note', 'date',
];

export default function SpreadsheetTable() {
  const [rowData, setRowData] = useState(() => getInitialState());
  const [searchText, setSearchText] = useState('');
  const [debouncedSearch, setDebouncedSearch] = useState('');
  const [columnFilters, setColumnFilters] = useState({});
  const [menuState, setMenuState] = useState(null);
  const [activeFilterColumn, setActiveFilterColumn] = useState(null);
  const [history, setHistory] = useState({ past: [], future: [] });
  const gridRef = useRef(null);
  const touchTimer = useRef(null);

  // Current state ref for undo/redo
  const currentRef = useRef(rowData);

  // Debounce search
  useEffect(() => {
    const t = setTimeout(() => setDebouncedSearch(searchText), 250);
    return () => clearTimeout(t);
  }, [searchText]);

  // Filter data
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

  // Push history snapshot
  const pushHistory = useCallback(() => {
    setHistory(prev => ({
      past: [...prev.past.slice(-99), [...currentRef.current]],
      future: [],
    }));
  }, []);

  // Undo
  const undo = useCallback(() => {
    setHistory(prev => {
      if (prev.past.length === 0) return prev;
      const previous = prev.past[prev.past.length - 1];
      const newPast = prev.past.slice(0, -1);
      setRowData(previous);
      currentRef.current = previous;
      return { past: newPast, future: [[...currentRef.current], ...prev.future] };
    });
  }, []);

  // Redo
  const redo = useCallback(() => {
    setHistory(prev => {
      if (prev.future.length === 0) return prev;
      const next = prev.future[0];
      const newFuture = prev.future.slice(1);
      setRowData(next);
      currentRef.current = next;
      return { past: [...prev.past, [...currentRef.current]], future: newFuture };
    });
  }, []);

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

  // Cell value changed
  const onCellValueChanged = useCallback((event) => {
    const { data, colDef, newValue } = event;
    const field = colDef.field;
    const newRows = currentRef.current.map(r => {
      if (r._id === data._id) {
        return { ...r, [field]: newValue ?? '' };
      }
      return r;
    });
    currentRef.current = newRows;
    setRowData(newRows);
    // Add to history
    setHistory(prev => ({
      past: [...prev.past.slice(-99), [...newRows]],
      future: [],
    }));
  }, []);

  // Context menu
  const onContextMenu = useCallback((e) => {
    e.preventDefault();
    const gridApi = gridRef.current?.api;
    if (!gridApi) return;

    // Get clicked row
    const target = e.target.closest('.ag-row');
    let rowIndex = -1;
    if (target) {
      const rowComp = gridApi.getRowNode(target.getAttribute('row-id'));
      if (rowComp) {
        rowIndex = rowComp.rowIndex;
      }
    }

    // Fallback: use mouse position
    if (rowIndex === -1) {
      const rect = gridRef.current?.eGridDiv?.getBoundingClientRect();
      if (rect) {
        const y = e.clientY - rect.top;
        const rowHeight = 36;
        rowIndex = Math.floor(y / rowHeight);
      }
    }

    const x = Math.min(e.clientX, window.innerWidth - 240);
    const y = Math.min(e.clientY, window.innerHeight - 140);
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

  // Insert row
  const insertRow = useCallback((position) => {
    if (!menuState) return;
    pushHistory();
    const idx = menuState.rowIndex;
    const empty = createEmptyRow(0);
    const newRows = [...currentRef.current];
    const insertAt = position === 'above' ? idx : idx + 1;
    newRows.splice(insertAt, 0, empty);
    const renumbered = renumber(newRows);
    currentRef.current = renumbered;
    setRowData(renumbered);
    setMenuState(null);
    setHistory(prev => ({
      past: [...prev.past.slice(-99), [...renumbered]],
      future: [],
    }));
  }, [menuState, pushHistory]);

  // Delete row
  const deleteRow = useCallback(() => {
    if (!menuState) return;
    pushHistory();
    const idx = menuState.rowIndex;
    const newRows = [...currentRef.current];
    if (idx >= 0 && idx < newRows.length) {
      newRows.splice(idx, 1);
      const renumbered = renumber(newRows);
      currentRef.current = renumbered;
      setRowData(renumbered);
    }
    setMenuState(null);
    setHistory(prev => ({
      past: [...prev.past.slice(-99), [...currentRef.current]],
      future: [],
    }));
  }, [menuState, pushHistory]);

  // Filter apply
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

  // Double-click header to open filter
  const handleColumnHeaderDblClick = useCallback((event) => {
    const colId = event.column?.getId?.();
    if (colId && colId !== '_no') {
      setActiveFilterColumn(colId);
    }
  }, []);

  const canUndo = history.past.length > 0;
  const canRedo = history.future.length > 0;

  const columnDefs = useMemo(() => [
    {
      field: '_no',
      headerName: '№',
      width: 55,
      editable: false,
      pinned: 'left',
      cellClass: 'cell-readonly cell-center cell-no',
      headerClass: 'header-no',
    },
    {
      field: 'fullName',
      headerName: 'Soyad, Ad və Ata adı',
      width: 250,
      editable: true,
    },
    {
      field: 'serial',
      headerName: 'Seriya nömrəsi',
      width: 115,
      editable: true,
    },
    {
      field: 'idNumber',
      headerName: 'Fərdi ID nömrəsi',
      width: 120,
      editable: true,
    },
    {
      field: 'birthDate',
      headerName: 'Doğum tarixi',
      width: 105,
      editable: true,
    },
    {
      field: 'phone',
      headerName: 'Telefon',
      width: 160,
      editable: true,
    },
    {
      field: 'email',
      headerName: 'Email',
      width: 240,
      editable: true,
    },
    {
      field: 'rank',
      headerName: 'Rank (Working Diploma)',
      width: 190,
      editable: true,
    },
    {
      field: 'fullNameId',
      headerName: 'Full Name (ID)',
      width: 240,
      editable: false,
      cellClass: 'cell-readonly',
    },
    {
      field: 'rank2',
      headerName: 'Rank / Vəzifə',
      width: 150,
      editable: true,
    },
    {
      field: 'courseCode',
      headerName: 'Course Code',
      width: 105,
      editable: true,
    },
    {
      field: 'startDate',
      headerName: 'Başlama tarixi',
      width: 120,
      editable: true,
    },
    {
      field: 'finishDate',
      headerName: 'Bitmə tarixi',
      width: 120,
      editable: true,
    },
    {
      field: 'note',
      headerName: 'Qeyd',
      width: 130,
      editable: true,
    },
    {
      field: 'date',
      headerName: 'Tarix',
      width: 90,
      editable: true,
    },
  ], []);

  const defaultColDef = useMemo(() => ({
    sortable: false,
    resizable: true,
    filter: false,
    suppressMovable: true,
    suppressMenu: true,
  }), []);

  return (
    <div className="spreadsheet-root">
      {/* Toolbar */}
      <div className="toolbar">
        <div className="toolbar-left">
          <div className="search-box">
            <span className="search-icon">🔍</span>
            <input
              type="text"
              className="search-input"
              placeholder="Axtar..."
              value={searchText}
              onChange={(e) => setSearchText(e.target.value)}
            />
            {searchText && (
              <button className="search-clear" onClick={() => setSearchText('')}>✕</button>
            )}
          </div>
        </div>
        <div className="toolbar-center">
          <div className="undo-redo-group">
            <button
              className="btn-toolbar"
              onClick={undo}
              disabled={!canUndo}
              title="Geri (Ctrl+Z)"
            >
              ↩
            </button>
            <button
              className="btn-toolbar"
              onClick={redo}
              disabled={!canRedo}
              title="İrəli (Ctrl+Y)"
            >
              ↪
            </button>
          </div>
        </div>
        <div className="toolbar-right">
          <span className="row-count">
            {filteredData.length} / {rowData.length} sətir
          </span>
          {Object.keys(columnFilters).length > 0 && (
            <button className="btn-clear-filters" onClick={() => setColumnFilters({})}>
              ✕ Filtri təmizlə
            </button>
          )}
        </div>
      </div>

      {/* Active filter chips */}
      {Object.keys(columnFilters).length > 0 && (
        <div className="active-filters-bar">
          {Object.entries(columnFilters).map(([field, values]) => {
            const col = [
              { field: '_no', headerName: '№' },
              { field: 'fullName', headerName: 'Soyad, Ad' },
              { field: 'serial', headerName: 'Seriya' },
              { field: 'idNumber', headerName: 'ID' },
              { field: 'birthDate', headerName: 'Doğum' },
              { field: 'phone', headerName: 'Telefon' },
              { field: 'email', headerName: 'Email' },
              { field: 'rank', headerName: 'Rank' },
              { field: 'fullNameId', headerName: 'Full Name' },
              { field: 'rank2', headerName: 'Vəzifə' },
              { field: 'courseCode', headerName: 'Kurs kodu' },
              { field: 'startDate', headerName: 'Başlama' },
              { field: 'finishDate', headerName: 'Bitmə' },
              { field: 'note', headerName: 'Qeyd' },
              { field: 'date', headerName: 'Tarix' },
            ].find(c => c.field === field);
            return (
              <span className="filter-chip" key={field}>
                {col?.headerName || field}: {values.length}
                <button onClick={() => setColumnFilters(prev => {
                  const n = { ...prev };
                  delete n[field];
                  return n;
                })}>✕</button>
              </span>
            );
          })}
        </div>
      )}

      {/* Hint */}
      <div className="hint-bar">
        💡 Başlığa cüt klik — filtr açılır · Sağ klik — sətir əlavə/çıxar · Boşluqlar doldurulur
      </div>

      {/* Table */}
      <div
        className="ag-theme-quartz grid-wrap"
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
          rowSelection={undefined}
          animateRows={false}
        />
      </div>

      {/* Context Menu */}
      {menuState && (
        <ContextMenu
          x={menuState.x}
          y={menuState.y}
          onClose={() => setMenuState(null)}
          onInsertAbove={() => insertRow('above')}
          onInsertBelow={() => insertRow('below')}
          onDelete={deleteRow}
        />
      )}

      {/* Filter Panel */}
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
    </div>
  );
}
