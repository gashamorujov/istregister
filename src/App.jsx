import { useState } from 'react';
import SpreadsheetTable from './components/SpreadsheetTable';
import LoginScreen from './components/LoginScreen';
import AdminPanel, { getAccessCode } from './components/AdminPanel';
import './App.css';

function App() {
  const [access, setAccess] = useState(false);
  const [isAdmin, setIsAdmin] = useState(false);
  const [showAdmin, setShowAdmin] = useState(false);

  const handleLogin = (code) => {
    if (code === '__ADMIN__') {
      setIsAdmin(true);
      setAccess(true);
      return true;
    }
    if (code === getAccessCode()) {
      setAccess(true);
      setIsAdmin(false);
      return true;
    }
    return false;
  };

  if (!access) {
    return <LoginScreen onLogin={handleLogin} />;
  }

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
          <button className="btn-admin-panel" onClick={() => setShowAdmin(true)}>⚙️ Admin Panel</button>
        </div>
      </header>

      <SpreadsheetTable />

      <footer className="app-footer">
        <span>© 2026 Industrial Support and Training MMC · İSTREGISTER</span>
      </footer>

      {showAdmin && <AdminPanel onClose={() => setShowAdmin(false)} />}
    </div>
  );
}

export default App;
