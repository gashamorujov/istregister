import { useState, useEffect, useRef, useCallback } from 'react';
import {
  recordsRef,
  singleRecordRef,
  onValue,
  set as fbSet,
  update as fbUpdate,
  remove as fbRemove,
  push as fbPush,
  db,
  ref,
} from './firebase';

function generateId() {
  return `rec-${Date.now()}-${Math.random().toString(36).slice(2, 8)}`;
}

const EMPTY_ROW = {
  fullName: '', serial: '', idNumber: '', birthDate: '',
  phone: '', email: '', rank: '', fullNameId: '', rank2: '',
  courseCode: '', startDate: '', finishDate: '', note: '', date: '',
};

const MAX_HISTORY = 50;

export default function useFirebaseData() {
  const [rows, setRows] = useState([]);
  const [loading, setLoading] = useState(true);
  const [connected, setConnected] = useState(true);

  const historyRef = useRef([]);
  const futureRef = useRef([]);
  const skipNextRef = useRef(false);
  const mountedRef = useRef(true);

  useEffect(() => {
    mountedRef.current = true;
    return () => { mountedRef.current = false; };
  }, []);

  // Subscribe to Firebase records
  useEffect(() => {
    const unsubscribe = onValue(recordsRef(), (snapshot) => {
      if (!mountedRef.current) return;
      if (skipNextRef.current) {
        skipNextRef.current = false;
      } else {
        const val = snapshot.val();
        if (val) {
          const sanitize = (v) => {
            if (v === null || v === undefined) return '';
            if (typeof v === 'object') {
              if (v.richText && Array.isArray(v.richText)) return v.richText.map(rt => rt.text || '').join('');
              if (v.text !== undefined) return String(v.text);
              return '';
            }
            const s = String(v);
            return s.includes('[object Object]') ? '' : s;
          };
          const list = Object.entries(val).map(([id, data]) => {
            const clean = {};
            for (const [k, v] of Object.entries(data)) {
              clean[k] = sanitize(v);
            }
            return { _id: id, ...clean };
          });
          setRows(list);
        } else {
          setRows([]);
        }
      }
      setLoading(false);
    }, (error) => {
      console.error('Firebase read error:', error);
      setLoading(false);
    });

    return () => unsubscribe();
  }, []);

  // Connection status via .info/connected
  useEffect(() => {
    const connRef = ref(db, '.info/connected');
    const unsub = onValue(connRef, (snap) => {
      if (mountedRef.current) {
        setConnected(snap.val() === true);
      }
    }, () => {});
    return () => unsub();
  }, []);

  const pushHistory = useCallback((snapshot) => {
    historyRef.current = [...historyRef.current.slice(-MAX_HISTORY + 1), snapshot];
    futureRef.current = [];
  }, []);

  const flushToFirebase = useCallback(async (newRows) => {
    try {
      const data = {};
      newRows.forEach(row => {
        const { _id, ...rest } = row;
        data[_id || generateId()] = rest;
      });
      skipNextRef.current = true;
      await fbSet(recordsRef(), data);
    } catch (err) {
      console.error('Firebase write error:', err);
    }
  }, []);

  const addRow = useCallback((position, afterIndex) => {
    setRows(prev => {
      const newRow = { _id: generateId(), ...EMPTY_ROW };
      const next = [...prev];
      const idx = afterIndex !== undefined ? afterIndex : prev.length - 1;
      if (position === 'above') {
        next.splice(idx, 0, newRow);
      } else {
        next.splice(idx + 1, 0, newRow);
      }
      pushHistory(prev);
      flushToFirebase(next);
      return next;
    });
  }, [pushHistory, flushToFirebase]);

  const deleteRow = useCallback((index) => {
    setRows(prev => {
      const next = prev.filter((_, i) => i !== index);
      pushHistory(prev);
      flushToFirebase(next);
      return next;
    });
  }, [pushHistory, flushToFirebase]);

  const updateCell = useCallback((rowId, field, value) => {
    setRows(prev => {
      const next = prev.map(r => r._id === rowId ? { ...r, [field]: value } : r);
      pushHistory(prev);
      const { _id, ...data } = next.find(r => r._id === rowId) || {};
      if (_id) {
        fbUpdate(singleRecordRef(_id), data).catch(err => console.error('Cell save failed:', err));
      }
      return next;
    });
  }, [pushHistory]);

  const importRows = useCallback((newRecords) => {
    setRows(prev => {
      const next = [...prev, ...newRecords.map(r => ({
        _id: generateId(),
        ...r,
      }))];
      pushHistory(prev);
      flushToFirebase(next);
      return next;
    });
  }, [pushHistory, flushToFirebase]);

  const undo = useCallback(() => {
    if (historyRef.current.length === 0) return;
    setRows(() => {
      const snapshot = historyRef.current[historyRef.current.length - 1];
      historyRef.current = historyRef.current.slice(0, -1);
      futureRef.current = [...futureRef.current, rows];
      flushToFirebase(snapshot);
      return snapshot;
    });
  }, [rows, flushToFirebase]);

  const redo = useCallback(() => {
    if (futureRef.current.length === 0) return;
    setRows(() => {
      const snapshot = futureRef.current[futureRef.current.length - 1];
      futureRef.current = futureRef.current.slice(0, -1);
      historyRef.current = [...historyRef.current, rows];
      flushToFirebase(snapshot);
      return snapshot;
    });
  }, [rows, flushToFirebase]);

  return {
    rows,
    loading,
    connected,
    canUndo: historyRef.current.length > 0,
    canRedo: futureRef.current.length > 0,
    addRow,
    deleteRow,
    updateCell,
    importRows,
    undo,
    redo,
  };
}
