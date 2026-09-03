import { useCallback, useMemo, useRef, useState, useEffect } from 'react';
import { AgGridReact } from 'ag-grid-react';
import 'ag-grid-community/styles/ag-grid.css';
import 'ag-grid-community/styles/ag-theme-quartz.css';
import { ClientSideRowModelModule } from 'ag-grid-community';
import ContextMenu from './components/ContextMenu';
import TrainingPlanModal from './components/TrainingPlanModal';
import FilterPanel from './components/FilterPanel';
import { generateTrainingPlan, getUniqueCourseGroups } from './lib/excelGenerator';
import studentsData from './data/registrData.json';
import { getCourseName } from './data/courses';
import './App.css';

const COLUMNS = [
  { field: 'no', headerName: '№', width: 55 },
  { field: 'fullName', headerName: 'Soyad, Ad və Ata adı', width: 250 },
  { field: 'serial', headerName: 'Seriya nömrəsi', width: 115 },
  { field: 'idNumber', headerName: 'Fərdi identifikasiya nömrəsi', width: 130 },
  { field: 'birthDate', headerName: 'Doğum tarixi', width: 105 },
  { field: 'phone', headerName: 'Telefon nömrəsi', width: 175 },
  { field: 'email', headerName: 'email', width: 250 },
  { field: 'rank', headerName: 'Rank (Working Diploma)', width: 190 },
  { field: 'fullNameId', headerName: 'Full Name (ID)', width: 250 },
  { field: 'rank2', headerName: 'Rank / Vəzifə', width: 150 },
  { field: 'courseCode', headerName: 'Course Code', width: 105 },
  { field: 'startDate', headerName: 'Course Start Date', width: 135 },
  { field: 'finishDate', headerName: 'Course Finish Date', width: 135 },
  { field: 'note', headerName: 'Qeyd', width: 115 },
  { field: 'date', headerName: 'Tarix', width: 85 },
];

function App() {
  const [searchText, setSearchText] = useState('');
  const [debouncedSearch, setDebouncedSearch] = useState('');
  const [columnFilters, setColumnFilters] = useState({});
  const [menuState, setMenuState] = useState(null);
  const [modalOpen, setModalOpen] = useState(false);
  const [modalGroups, setModalGroups] = useState([]);
  const [activeFilterColumn, setActiveFilterColumn] = useState(null);
  const [filteredRowData, setFilteredRowData] = useState(studentsData);
  const gridRef = useRef(null);
  const touchTimer = useRef(null);

  // Debounce search
  useEffect(() => {
    const timer = setTimeout(() => setDebouncedSearch(searchText), 250);
    return () => clearTimeout(timer);
  }, [searchText]);

  // Filter data
  const applyFilters = useCallback(() => {
    let data = [...studentsData];
    
    const activeFilters = Object.entries(columnFilters);
    if (activeFilters.length > 0) {
      data = data.filter(row => {
        return activeFilters.every(([field, filterValues]) => {
          if (!filterValues || filterValues.length === 0) return true;
          const cellValue = String(row[field] || '').toLowerCase();
          return filterValues.some(v => cellValue.includes(v.toLowerCase()));
        });
      });
    }
    
    if (debouncedSearch.trim()) {
      const q = debouncedSearch.toLowerCase().trim();
      data = data.filter(row =>
        Object.values(row).some(val =>
          String(val || '').toLowerCase().includes(q)
        )
      );
    }
    
    setFilteredRowData(data);
  }, [columnFilters, debouncedSearch]);

  useEffect(() => {
    applyFilters();
  }, [applyFilters]);

  // Context menu
  const onContextMenu = useCallback((e) => {
    e.preventDefault();
    const x = Math.min(e.clientX, window.innerWidth - 220);
    const y = Math.min(e.clientY, window.innerHeight - 120);
    setMenuState({ x, y });
  }, []);

  const handleTouchStart = useCallback((e) => {
    touchTimer.current = setTimeout(() => {
      setMenuState({ x: 100, y: 100 });
    }, 600);
  }, []);

  const handleTouchEnd = useCallback(() => {
    if (touchTimer.current) clearTimeout(touchTimer.current);
  }, []);

  const handleTrainingPlan = () => {
    setMenuState(null);
    const groups = getUniqueCourseGroups(filteredRowData);
    setModalGroups(groups);
    setModalOpen(true);
  };

  const handleProtocol = () => setMenuState(null);

  const handleConfirmTrainingPlan = async (selections) => {
    try {
      await generateTrainingPlan(filteredRowData, selections);
      setModalOpen(false);
    } catch (error) {
      console.error('Training Plan xətası:', error);
      alert('Training Plan yaradılmasında xəta baş verdi.');
    }
  };

  const getUniqueValues = useCallback((field) => {
    const values = new Set();
    studentsData.forEach(row => {
      const val = row[field];
      if (val !== null && val !== undefined && val !== '') {
        values.add(String(val));
      }
    });
    return Array.from(values).sort();
  }, []);

  const handleFilterApply = useCallback((field, values) => {
    setColumnFilters(prev => {
      const next = { ...prev };
      if (values && values.length > 0) {
        next[field] = values;
      } else {
        delete next[field];
      }
      return next;
    });
    setActiveFilterColumn(null);
  }, []);

  const handleColumnHeaderDblClick = useCallback((event) => {
    const colId = event.column?.getId?.();
    if (colId) {
      setActiveFilterColumn(colId);
    }
  }, []);

  const columnDefs = useMemo(() => COLUMNS.map(col => ({
    field: col.field,
    headerName: col.headerName,
    width: col.width,
    sortable: true,
    resizable: true,
  })), []);

  const defaultColDef = useMemo(() => ({
    sortable: true,
    resizable: true,
  }), []);

  return (
    <div className="app">
      <header className="app-header">
        <div className="header-left">
          <h1>🗂️ İSTREGISTER</h1>
          <span className="header-subtitle">Tədris Reyestri İdarəetmə Sistemi</span>
        </div>
        <div className="header-right">
          <span className="header-org">Industrial Support and Training MMC</span>
        </div>
      </header>
      
      <div className="toolbar">
        <div className="search-container">
          <input
            type="text"
            className="search-input"
            placeholder="🔍 Bütün sütunlar üzrə axtar..."
            value={searchText}
            onChange={(e) => setSearchText(e.target.value)}
          />
          {searchText && (
            <button className="search-clear" onClick={() => setSearchText('')}>×</button>
          )}
        </div>
        <div className="toolbar-stats">
          <span className="stat-count">
            <strong>{filteredRowData.length}</strong> / {studentsData.length} qeyd
          </span>
          {Object.keys(columnFilters).length > 0 && (
            <button className="btn btn-ghost" onClick={() => setColumnFilters({})}>
              ✕ Filtrləri təmizlə
            </button>
          )}
        </div>
      </div>

      {Object.keys(columnFilters).length > 0 && (
        <div className="active-filters">
          {Object.entries(columnFilters).map(([field, values]) => (
            <span className="filter-chip" key={field}>
              {COLUMNS.find(c => c.field === field)?.headerName}: {values.length}
              <button onClick={() => setColumnFilters(prev => {
                const next = { ...prev };
                delete next[field];
                return next;
              })}>×</button>
            </span>
          ))}
        </div>
      )}

      <div className="filter-hint">
        💡 Sütun başlığına iki dəfə klikləyin — filtr açılır | Cədvəldə sağ klik / uzun basış → Training Plan
      </div>

      <div
        className="ag-theme-quartz grid-container"
        onContextMenu={onContextMenu}
        onTouchStart={handleTouchStart}
        onTouchEnd={handleTouchEnd}
        onTouchMove={handleTouchEnd}
      >
        <AgGridReact
          ref={gridRef}
          rowData={filteredRowData}
          columnDefs={columnDefs}
          defaultColDef={defaultColDef}
          modules={[ClientSideRowModelModule]}
          enableCellTextSelection={true}
          rowHeight={38}
          headerHeight={42}
          suppressRowHoverHighlight={false}
          onColumnHeaderClicked={handleColumnHeaderDblClick}
        />
      </div>

      <footer className="app-footer">
        <span>© 2026 Industrial Support and Training MMC — Tədris Reyestri İdarəetmə Sistemi</span>
      </footer>

      {menuState && (
        <ContextMenu
          x={menuState.x}
          y={menuState.y}
          onClose={() => setMenuState(null)}
          onTrainingPlan={handleTrainingPlan}
          onProtocol={handleProtocol}
        />
      )}

      {modalOpen && (
        <TrainingPlanModal
          groups={modalGroups}
          onConfirm={handleConfirmTrainingPlan}
          onCancel={() => setModalOpen(false)}
        />
      )}

      {activeFilterColumn && (
        <FilterPanel
          field={activeFilterColumn}
          headerName={COLUMNS.find(c => c.field === activeFilterColumn)?.headerName}
          values={getUniqueValues(activeFilterColumn)}
          selected={columnFilters[activeFilterColumn] || []}
          onApply={handleFilterApply}
          onClose={() => setActiveFilterColumn(null)}
        />
      )}
    </div>
  );
}

export default App;
