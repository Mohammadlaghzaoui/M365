import { useState, useCallback } from 'react';

const PREFIX = 'workpilot:';

// Optional cross-device sync hook (wired by cloudStore at runtime to avoid a cycle).
let onSaved: (() => void) | null = null;
export function setSaveHook(fn: (() => void) | null) { onSaved = fn; }

export function load<T>(key: string, fallback: T): T {
  try {
    const raw = localStorage.getItem(PREFIX + key);
    return raw ? (JSON.parse(raw) as T) : fallback;
  } catch {
    return fallback;
  }
}

export function save<T>(key: string, value: T): void {
  localStorage.setItem(PREFIX + key, JSON.stringify(value));
  if (key !== 'cloud-sync' && key !== 'session') onSaved?.();
}

export function useLocalStorage<T>(key: string, fallback: T) {
  const [value, setValue] = useState<T>(() => load(key, fallback));
  const set = useCallback(
    (next: T | ((prev: T) => T)) => {
      setValue((prev) => {
        const v = typeof next === 'function' ? (next as (p: T) => T)(prev) : next;
        save(key, v);
        return v;
      });
    },
    [key],
  );
  return [value, set] as const;
}

export function uid(): string {
  return Date.now().toString(36) + Math.random().toString(36).slice(2, 8);
}
