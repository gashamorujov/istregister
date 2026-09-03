import { useState, useCallback, useRef } from 'react';

const MAX_HISTORY = 100;

export function useUndoRedo(initialState) {
  const pastRef = useRef([]);
  const futureRef = useRef([]);
  const [state, setStateInternal] = useState(initialState);
  const stateRef = useRef(state);

  const setState = useCallback((newState) => {
    const prev = stateRef.current;
    pastRef.current = [...pastRef.current.slice(-(MAX_HISTORY - 1)), prev];
    futureRef.current = [];
    stateRef.current = newState;
    setStateInternal(newState);
  }, []);

  const undo = useCallback(() => {
    if (pastRef.current.length === 0) return;
    const previous = pastRef.current[pastRef.current.length - 1];
    pastRef.current = pastRef.current.slice(0, -1);
    futureRef.current = [stateRef.current, ...futureRef.current];
    stateRef.current = previous;
    setStateInternal(previous);
  }, []);

  const redo = useCallback(() => {
    if (futureRef.current.length === 0) return;
    const next = futureRef.current[0];
    futureRef.current = futureRef.current.slice(1);
    pastRef.current = [...pastRef.current, stateRef.current];
    stateRef.current = next;
    setStateInternal(next);
  }, []);

  const canUndo = pastRef.current.length > 0;
  const canRedo = futureRef.current.length > 0;

  return {
    state,
    setState,
    undo,
    redo,
    canUndo,
    canRedo,
  };
}

export function renumberRows(rows) {
  return rows.map((row, i) => ({ ...row, _no: i + 1 }));
}
