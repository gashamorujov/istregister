import SpreadsheetTable from './components/SpreadsheetTable';
import './App.css';

function App() {
  return (
    <div className="app">
      <header className="app-header">
        <div className="header-brand">
          <span className="brand-icon">🗂️</span>
          <div className="brand-text">
            <h1>İSTREGISTER</h1>
            <span className="brand-sub">Tədris Reyestri İdarəetmə Sistemi</span>
          </div>
        </div>
        <div className="header-right">
          <span className="org-label">Industrial Support and Training MMC</span>
        </div>
      </header>

      <SpreadsheetTable />

      <footer className="app-footer">
        <span>© 2026 Industrial Support and Training MMC · İSTREGISTER</span>
      </footer>
    </div>
  );
}

export default App;
