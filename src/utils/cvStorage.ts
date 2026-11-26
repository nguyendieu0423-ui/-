import type { Candidate, SavedRirekishoRecord } from '../types/crm';

const IDB_NAME = 'tokutei_crm_cv_idb';
const IDB_VERSION = 1;
const IDB_STORE = 'candidate_cvs';

const LOCAL_STORAGE_SAFE_BYTES = 3_400_000;

export interface PersistCandidatesResult {
  mode: 'localStorage' | 'indexeddb_offload' | 'error';
  byteSize: number;
  offloadedCount: number;
  errorMessage?: string;
}

const isIndexedDbAvailable = (): boolean =>
  typeof window !== 'undefined' && typeof window.indexedDB !== 'undefined';

const openCvDatabase = (): Promise<IDBDatabase> =>
  new Promise((resolve, reject) => {
    if (!isIndexedDbAvailable()) {
      reject(new Error('IndexedDB không khả dụng trên trình duyệt này.'));
      return;
    }
    const req = window.indexedDB.open(IDB_NAME, IDB_VERSION);
    req.onupgradeneeded = () => {
      const db = req.result;
      if (!db.objectStoreNames.contains(IDB_STORE)) {
        db.createObjectStore(IDB_STORE, { keyPath: 'id' });
      }
    };
    req.onsuccess = () => resolve(req.result);
    req.onerror = () => reject(req.error || new Error('Không mở được IndexedDB.'));
  });

export const saveCandidateCvsToIdb = async (
  records: Array<{ id: string; savedRirekisho: SavedRirekishoRecord }>
): Promise<void> => {
  if (!isIndexedDbAvailable() || records.length === 0) return;
  let db: IDBDatabase | null = null;
  try {
    db = await openCvDatabase();
    await new Promise<void>((resolve, reject) => {
      const tx = db!.transaction(IDB_STORE, 'readwrite');
      const store = tx.objectStore(IDB_STORE);
      records.forEach((rec) => {
        if (rec.id && rec.savedRirekisho) {
          store.put({
            id: rec.id,
            savedRirekisho: rec.savedRirekisho,
            updatedAt: rec.savedRirekisho.updatedAt || new Date().toISOString(),
          });
        }
      });
      tx.oncomplete = () => resolve();
      tx.onerror = () => reject(tx.error || new Error('Lỗi lưu CV vào IndexedDB.'));
    });
  } finally {
    if (db) db.close();
  }
};

export const loadCandidateCvsFromIdb = async (): Promise<Record<string, SavedRirekishoRecord>> => {
  if (!isIndexedDbAvailable()) return {};
  let db: IDBDatabase | null = null;
  try {
    db = await openCvDatabase();
    return await new Promise<Record<string, SavedRirekishoRecord>>((resolve) => {
      const tx = db!.transaction(IDB_STORE, 'readonly');
      const store = tx.objectStore(IDB_STORE);
      const req = store.getAll();
      req.onsuccess = () => {
        const result: Record<string, SavedRirekishoRecord> = {};
        const items = Array.isArray(req.result) ? req.result : [];
        items.forEach((entry: { id?: string; savedRirekisho?: SavedRirekishoRecord }) => {
          if (entry && entry.id && entry.savedRirekisho) {
            result[entry.id] = entry.savedRirekisho;
          }
        });
        resolve(result);
      };
      req.onerror = () => resolve({});
    });
  } catch {
    return {};
  } finally {
    if (db) db.close();
  }
};

export const compressBase64ImageStringIfHeavy = async (
  dataUrl: string,
  maxBytes = 90_000
): Promise<string> => {
  if (!dataUrl || !dataUrl.startsWith('data:image/') || dataUrl.length <= maxBytes) {
    return dataUrl;
  }
  if (typeof window === 'undefined' || typeof document === 'undefined') {
    return dataUrl;
  }

  return await new Promise<string>((resolve) => {
    const img = new Image();
    img.onload = () => {
      try {
        const maxSide = 360;
        const scale = Math.min(1, maxSide / Math.max(img.width || 1, img.height || 1));
        const width = Math.max(1, Math.round(img.width * scale));
        const height = Math.max(1, Math.round(img.height * scale));
        const canvas = document.createElement('canvas');
        canvas.width = width;
        canvas.height = height;
        const ctx = canvas.getContext('2d');
        if (!ctx) {
          resolve(dataUrl);
          return;
        }
        ctx.fillStyle = '#FFFFFF';
        ctx.fillRect(0, 0, width, height);
        ctx.drawImage(img, 0, 0, width, height);
        const compressed = canvas.toDataURL('image/jpeg', 0.72);
        resolve(compressed.length < dataUrl.length ? compressed : dataUrl);
      } catch {
        resolve(dataUrl);
      }
    };
    img.onerror = () => resolve(dataUrl);
    img.src = dataUrl;
  });
};

export const compressSavedRirekishoInCandidates = async (
  list: Candidate[]
): Promise<Candidate[]> => {
  if (!Array.isArray(list) || list.length === 0) return [];
  const next: Candidate[] = [];
  for (const cand of list) {
    const photo = cand.savedRirekisho?.draft?.photoDataUrl;
    if (photo && photo.startsWith('data:image/') && photo.length > 90_000) {
      const compressedPhoto = await compressBase64ImageStringIfHeavy(photo, 90_000);
      next.push({
        ...cand,
        savedRirekisho: {
          ...cand.savedRirekisho!,
          draft: {
            ...cand.savedRirekisho!.draft,
            photoDataUrl: compressedPhoto,
          },
        },
      });
    } else {
      next.push(cand);
    }
  }
  return next;
};

const stripHeavyPhotosOnly = (list: Candidate[]): Candidate[] =>
  list.map((c) => {
    if (!c.savedRirekisho?.draft?.photoDataUrl) return c;
    return {
      ...c,
      savedRirekisho: {
        ...c.savedRirekisho,
        draft: {
          ...c.savedRirekisho.draft,
          photoDataUrl: '',
        },
      },
    };
  });

const stripEntireSavedRirekisho = (list: Candidate[]): Candidate[] =>
  list.map((c) => {
    if (!c.savedRirekisho) return c;
    const rest = { ...c };
    delete rest.savedRirekisho;
    return rest;
  });

export const persistCandidatesWithIdbFallback = async (
  storageKey: string,
  list: Candidate[]
): Promise<PersistCandidatesResult> => {
  if (typeof localStorage === 'undefined') {
    return { mode: 'localStorage', byteSize: 0, offloadedCount: 0 };
  }

  const serialized = JSON.stringify(list);
  const byteSize = serialized.length;

  const tryWriteLocal = (payloadStr: string): { ok: boolean; err?: string } => {
    try {
      localStorage.setItem(storageKey, payloadStr);
      return { ok: true };
    } catch (e) {
      const msg = e instanceof Error ? e.message : String(e);
      return { ok: false, err: msg };
    }
  };

  if (byteSize <= LOCAL_STORAGE_SAFE_BYTES) {
    const direct = tryWriteLocal(serialized);
    if (direct.ok) {
      return { mode: 'localStorage', byteSize, offloadedCount: 0 };
    }
  }

  // Offload savedRirekisho (and portrait base64 photos) into IndexedDB
  const cvRecords = list
    .filter((c) => Boolean(c.id && c.savedRirekisho))
    .map((c) => ({ id: c.id, savedRirekisho: c.savedRirekisho! }));

  if (cvRecords.length > 0) {
    try {
      await saveCandidateCvsToIdb(cvRecords);
    } catch {
      // Continue attempting compact write even if IDB fails
    }
  }

  const withoutPhotosStr = JSON.stringify(stripHeavyPhotosOnly(list));
  if (withoutPhotosStr.length <= LOCAL_STORAGE_SAFE_BYTES) {
    const photoStrippedWrite = tryWriteLocal(withoutPhotosStr);
    if (photoStrippedWrite.ok) {
      return {
        mode: 'indexeddb_offload',
        byteSize,
        offloadedCount: cvRecords.length,
      };
    }
  }

  const withoutCvStr = JSON.stringify(stripEntireSavedRirekisho(list));
  const cvStrippedWrite = tryWriteLocal(withoutCvStr);
  if (cvStrippedWrite.ok) {
    return {
      mode: 'indexeddb_offload',
      byteSize,
      offloadedCount: cvRecords.length,
    };
  }

  return {
    mode: 'error',
    byteSize,
    offloadedCount: cvRecords.length,
    errorMessage:
      cvStrippedWrite.err ||
      'Dung lượng dữ liệu vượt quá giới hạn lưu trữ trình duyệt (localStorage QuotaExceededError).',
  };
};

export const persistJsonWithQuotaCheck = (
  storageKey: string,
  value: unknown
): { ok: boolean; errorMessage?: string } => {
  if (typeof localStorage === 'undefined') return { ok: true };
  try {
    localStorage.setItem(storageKey, JSON.stringify(value));
    return { ok: true };
  } catch (e) {
    return {
      ok: false,
      errorMessage:
        e instanceof Error
          ? e.message
          : 'Không đủ dung lượng localStorage để lưu dữ liệu.',
    };
  }
};
