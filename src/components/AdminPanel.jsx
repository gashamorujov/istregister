import { useState, useEffect } from 'react';

const DEFAULT_CODE = '0706';

function getCode() {
  return localStorage.getItem('istregister_passcode') || DEFAULT_CODE;
}

function setCode(code) {
  localStorage.setItem('istregister_passcode', code);
}

export function getAccessCode() {
  return getCode();
}

export default function AdminPanel({ onClose }) {
  const [currentCode, setCurrentCode] = useState('');
  const [newCode, setNewCode] = useState('');
  const [confirmCode, setConfirmCode] = useState('');
  const [msg, setMsg] = useState('');
  const [loading, setLoading] = useState(false);

  const handleSave = (e) => {
    e.preventDefault();
    setMsg('');

    if (currentCode !== getCode()) {
      setMsg('Cari kod yanlışdır.');
      return;
    }
    if (!newCode || newCode.length < 4) {
      setMsg('Yeni kod minimum 4 simvol olmalıdır.');
      return;
    }
    if (newCode !== confirmCode) {
      setMsg('Yeni kodlar uyğun gəlmir.');
      return;
    }

    setLoading(true);
    setTimeout(() => {
      setCode(newCode);
      setMsg('✅ Kod uğurla dəyişdirildi.');
      setCurrentCode('');
      setNewCode('');
      setConfirmCode('');
      setLoading(false);
    }, 500);
  };

  return (
    <div className="admin-overlay" onClick={onClose}>
      <div className="admin-panel" onClick={(e) => e.stopPropagation()}>
        <div className="admin-header">
          <h2>⚙️ Admin Panel</h2>
          <button className="admin-close" onClick={onClose}>✕</button>
        </div>

        <form className="admin-form" onSubmit={handleSave}>
          <label className="admin-label">Cari giriş kodu</label>
          <input
            type="password"
            inputMode="numeric"
            className="admin-input"
            placeholder="••••"
            value={currentCode}
            onChange={(e) => setCurrentCode(e.target.value)}
            autoFocus
          />

          <label className="admin-label">Yeni kod</label>
          <input
            type="password"
            inputMode="numeric"
            className="admin-input"
            placeholder="••••"
            value={newCode}
            onChange={(e) => setNewCode(e.target.value)}
          />

          <label className="admin-label">Yeni kodu təsdiqlə</label>
          <input
            type="password"
            inputMode="numeric"
            className="admin-input"
            placeholder="••••"
            value={confirmCode}
            onChange={(e) => setConfirmCode(e.target.value)}
          />

          {msg && <div className={`admin-msg ${msg.startsWith('✅') ? 'ok' : 'err'}`}>{msg}</div>}

          <button className="btn-admin-save" type="submit" disabled={loading}>
            {loading ? 'Yadda saxlanılır...' : 'Kodu dəyişdir'}
          </button>
        </form>
      </div>
    </div>
  );
}
