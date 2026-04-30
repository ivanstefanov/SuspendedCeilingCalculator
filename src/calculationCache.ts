const DB_NAME = "scc-calculation-cache";
const DB_VERSION = 1;
const STORE_NAME = "entries";
const LOCAL_STORAGE_PREFIX = "scc-calculation-cache:";

export type CalculationCacheKind = "room-table-cuts" | "cut-plan" | "material-takeoff";

type CacheEntry<T> = {
  key: string;
  value: T;
  updatedAt: number;
};

let dbPromise: Promise<IDBDatabase | null> | null = null;

function openCacheDb(): Promise<IDBDatabase | null> {
  if (typeof indexedDB === "undefined") return Promise.resolve(null);
  if (dbPromise) return dbPromise;

  dbPromise = new Promise((resolve) => {
    const request = indexedDB.open(DB_NAME, DB_VERSION);

    request.onupgradeneeded = () => {
      const db = request.result;
      if (!db.objectStoreNames.contains(STORE_NAME)) {
        db.createObjectStore(STORE_NAME, { keyPath: "key" });
      }
    };
    request.onsuccess = () => resolve(request.result);
    request.onerror = () => resolve(null);
    request.onblocked = () => resolve(null);
  });

  return dbPromise;
}

function stableStringify(value: unknown): string {
  if (value === null || typeof value !== "object") return JSON.stringify(value);
  if (Array.isArray(value)) return `[${value.map(stableStringify).join(",")}]`;

  const record = value as Record<string, unknown>;
  return `{${Object.keys(record)
    .sort()
    .map((key) => `${JSON.stringify(key)}:${stableStringify(record[key])}`)
    .join(",")}}`;
}

function fallbackStorageKey(key: string): string {
  return `${LOCAL_STORAGE_PREFIX}${key}`;
}

function hashString(value: string): string {
  let hash = 5381;
  for (let index = 0; index < value.length; index += 1) {
    hash = (hash * 33) ^ value.charCodeAt(index);
  }
  return (hash >>> 0).toString(36);
}

export function buildCalculationCacheKey(kind: CalculationCacheKind, input: unknown): string {
  const serialized = stableStringify(input);
  return `${kind}:v2:${serialized.length}:${hashString(serialized)}`;
}

export async function readCalculationCache<T>(key: string): Promise<T | null> {
  const db = await openCacheDb();
  if (db) {
    const cached = await new Promise<CacheEntry<T> | null>((resolve) => {
      const transaction = db.transaction(STORE_NAME, "readonly");
      const store = transaction.objectStore(STORE_NAME);
      const request = store.get(key);
      request.onsuccess = () => resolve((request.result as CacheEntry<T> | undefined) ?? null);
      request.onerror = () => resolve(null);
    });
    if (cached) return cached.value;
  }

  try {
    const raw = localStorage.getItem(fallbackStorageKey(key));
    return raw ? (JSON.parse(raw) as CacheEntry<T>).value : null;
  } catch {
    return null;
  }
}

export async function writeCalculationCache<T>(key: string, value: T): Promise<void> {
  const entry: CacheEntry<T> = { key, value, updatedAt: Date.now() };
  const db = await openCacheDb();
  if (db) {
    await new Promise<void>((resolve) => {
      const transaction = db.transaction(STORE_NAME, "readwrite");
      const store = transaction.objectStore(STORE_NAME);
      const request = store.put(entry);
      request.onsuccess = () => resolve();
      request.onerror = () => resolve();
    });
    return;
  }

  try {
    localStorage.setItem(fallbackStorageKey(key), JSON.stringify(entry));
  } catch {
    // Cache writes must never block the calculator flow.
  }
}
