import { useState, useEffect } from "react";
import type { ZodType } from "zod";

function parseStoredValue<T>(stored: string | null, defaultValue: T, schema?: ZodType<T>): T {
  if (!stored) return defaultValue;

  const parsed = JSON.parse(stored) as unknown;
  if (!schema) return parsed as T;

  const result = schema.safeParse(parsed);
  return result.success ? result.data : defaultValue;
}

export function useLocalStorage<T>(key: string, defaultValue: T): [T, (value: T) => void];
export function useLocalStorage<T>(key: string, defaultValue: T, schema: ZodType<T>): [T, (value: T) => void];
export function useLocalStorage<T>(key: string, defaultValue: T, schema?: ZodType<T>): [T, (value: T) => void] {
  const [value, setValue] = useState<T>(() => {
    try {
      const stored = localStorage.getItem(key);
      return parseStoredValue(stored, defaultValue, schema);
    } catch {
      return defaultValue;
    }
  });

  useEffect(() => {
    try {
      localStorage.setItem(key, JSON.stringify(value));
    } catch {}
  }, [key, value]);

  return [value, setValue];
}
