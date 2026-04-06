import { useSyncExternalStore } from "react";
import type { ZodType } from "zod";

type StorageSubscriber = () => void;

const listenersByKey = new Map<string, Set<StorageSubscriber>>();

function getStorageApi() {
  return typeof globalThis.localStorage === "undefined"
    ? null
    : globalThis.localStorage;
}

function decodeStoredValue<T>(stored: string | null, fallback: T, schema?: ZodType<T>): T {
  if (stored === null) return fallback;

  const parsed = JSON.parse(stored) as unknown;
  if (!schema) return parsed as T;

  const result = schema.safeParse(parsed);
  return result.success ? result.data : fallback;
}

function readStoredSnapshot<T>(key: string, fallback: T, schema?: ZodType<T>) {
  const storage = getStorageApi();
  if (!storage) return fallback;

  try {
    return decodeStoredValue(storage.getItem(key), fallback, schema);
  } catch {
    return fallback;
  }
}

function notifySubscribers(key: string) {
  const listeners = listenersByKey.get(key);
  if (!listeners) return;
  for (const listener of listeners) {
    listener();
  }
}

function subscribeToKey(key: string, onStoreChange: StorageSubscriber) {
  const listeners = listenersByKey.get(key) ?? new Set<StorageSubscriber>();
  listeners.add(onStoreChange);
  listenersByKey.set(key, listeners);

  const handleStorage = (event: StorageEvent) => {
    if (event.key === null || event.key === key) {
      onStoreChange();
    }
  };

  if (typeof window !== "undefined") {
    window.addEventListener("storage", handleStorage);
  }

  return () => {
    listeners.delete(onStoreChange);
    if (listeners.size === 0) {
      listenersByKey.delete(key);
    }

    if (typeof window !== "undefined") {
      window.removeEventListener("storage", handleStorage);
    }
  };
}

function persistValue<T>(key: string, value: T) {
  const storage = getStorageApi();
  if (!storage) return;

  try {
    storage.setItem(key, JSON.stringify(value));
    notifySubscribers(key);
  } catch {}
}

export function useLocalStorage<T>(key: string, defaultValue: T): [T, (value: T) => void];
export function useLocalStorage<T>(key: string, defaultValue: T, schema: ZodType<T>): [T, (value: T) => void];
export function useLocalStorage<T>(key: string, defaultValue: T, schema?: ZodType<T>): [T, (value: T) => void] {
  const value = useSyncExternalStore(
    (onStoreChange) => subscribeToKey(key, onStoreChange),
    () => readStoredSnapshot(key, defaultValue, schema),
    () => defaultValue,
  );

  const setValue = (nextValue: T) => {
    persistValue(key, nextValue);
  };

  return [value, setValue];
}
