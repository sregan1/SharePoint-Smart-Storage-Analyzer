import { StoredReport } from '../models/models';

const DB_NAME = 'SmartStorageAnalyzerHistory';
const DB_VERSION = 1;
const STORE_NAME = 'reports';
const MAX_REPORTS = 10;

export class ReportHistoryService {
  private openDb(): Promise<IDBDatabase> {
    return new Promise((resolve, reject) => {
      const req = indexedDB.open(DB_NAME, DB_VERSION);
      req.onupgradeneeded = (e) => {
        const db = (e.target as IDBOpenDBRequest).result;
        if (!db.objectStoreNames.contains(STORE_NAME)) {
          const store = db.createObjectStore(STORE_NAME, { keyPath: 'id' });
          store.createIndex('timestamp', 'timestamp', { unique: false });
        }
      };
      req.onsuccess = (e) => resolve((e.target as IDBOpenDBRequest).result);
      req.onerror = (e) => reject((e.target as IDBOpenDBRequest).error);
    });
  }

  async add(report: StoredReport): Promise<void> {
    const db = await this.openDb();
    return new Promise((resolve, reject) => {
      const tx = db.transaction(STORE_NAME, 'readwrite');
      const store = tx.objectStore(STORE_NAME);
      const fail = (err: any): void => { db.close(); reject(err); };

      const countReq = store.count();
      countReq.onsuccess = () => {
        const doAdd = (): void => {
          store.put(report);
          tx.oncomplete = () => { db.close(); resolve(); };
          tx.onerror = () => fail(tx.error);
        };

        // Trim to exactly MAX_REPORTS oldest-first, not just one — a single
        // delete only keeps the store from growing when it was already at
        // the limit; if it were ever more than one over (a version that
        // shipped without this trim, direct DB access, etc.), that excess
        // would never shrink back down since each add() would delete one
        // and add one, net zero.
        let toDelete = countReq.result - MAX_REPORTS + 1;
        if (toDelete > 0) {
          const cursorReq = store.index('timestamp').openCursor(null, 'next');
          cursorReq.onsuccess = (e) => {
            const cursor = (e.target as IDBRequest<IDBCursorWithValue>).result;
            if (cursor && toDelete > 0) {
              store.delete(cursor.primaryKey);
              toDelete--;
              cursor.continue();
            } else {
              doAdd();
            }
          };
          cursorReq.onerror = () => fail(cursorReq.error);
        } else {
          doAdd();
        }
      };
      countReq.onerror = () => fail(countReq.error);
    });
  }

  async getAll(): Promise<StoredReport[]> {
    const db = await this.openDb();
    return new Promise((resolve, reject) => {
      const tx = db.transaction(STORE_NAME, 'readonly');
      const req = tx.objectStore(STORE_NAME).index('timestamp').getAll();
      req.onsuccess = () => {
        db.close();
        resolve((req.result as StoredReport[]).reverse()); // newest first
      };
      req.onerror = () => { db.close(); reject(req.error); };
    });
  }

  async delete(id: string): Promise<void> {
    const db = await this.openDb();
    return new Promise((resolve, reject) => {
      const tx = db.transaction(STORE_NAME, 'readwrite');
      tx.objectStore(STORE_NAME).delete(id);
      tx.oncomplete = () => { db.close(); resolve(); };
      tx.onerror = () => { db.close(); reject(tx.error); };
    });
  }
}
