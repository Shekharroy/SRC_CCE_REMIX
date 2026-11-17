import fs from 'fs';
import path from 'path';
import { MongoClient, type Db } from 'mongodb';
import type {
  DatabaseStatusInfo,
  GeneratedReportSnapshot,
  UploadedDocumentRecord,
} from '../src/shared/types.ts';

const DEFAULT_MONGO_URI =
  'mongodb://127.0.0.1:27017/?directConnection=true&serverSelectionTimeoutMS=2000&appName=mongosh+2.14.1';
const DB_NAME = process.env.MONGODB_DB || 'SRC_CCE';
const MONGO_URI = process.env.MONGODB_URI || DEFAULT_MONGO_URI;

const DATA_DIR = path.resolve(process.cwd(), 'data');
const FILES_DIR = path.resolve(DATA_DIR, 'files');
const LOCAL_DB_FILE = path.resolve(DATA_DIR, 'src_cce_store.json');

const SAFE_ID_PATTERN = /^(doc|rep)_[a-zA-Z0-9_-]{4,80}$/;

export function isValidRecordId(id: string): boolean {
  return typeof id === 'string' && SAFE_ID_PATTERN.test(id);
}

function resolveSafeFilePath(fileName: string): string | null {
  const base = path.basename(fileName);
  if (!base || base === '.' || base === '..') return null;
  const resolved = path.resolve(FILES_DIR, base);
  if (!resolved.startsWith(FILES_DIR + path.sep)) {
    return null;
  }
  return resolved;
}

interface LocalStoreShape {
  uploaded_documents: Array<UploadedDocumentRecord & { storedFileName?: string }>;
  generated_reports: GeneratedReportSnapshot[];
}

class SrcCceDatabase {
  private client: MongoClient | null = null;
  private db: Db | null = null;
  private mode: 'mongodb' | 'local-fallback' = 'local-fallback';
  private lastError: string | undefined;
  private initialized = false;
  private initPromise: Promise<void> | null = null;
  private writeQueue: Promise<unknown> = Promise.resolve();

  async init(forceReconnect = false): Promise<void> {
    if (this.initialized && !forceReconnect) return;
    if (this.initPromise && !forceReconnect) return this.initPromise;

    this.initPromise = (async () => {
      this.ensureLocalDirs();

      if (forceReconnect && this.client) {
        await this.client.close().catch(() => {});
        this.client = null;
        this.db = null;
      }

      try {
        const client = new MongoClient(MONGO_URI, {
          serverSelectionTimeoutMS: 2000,
          connectTimeoutMS: 2000,
        });
        await client.connect();
        const db = client.db(DB_NAME);
        await db.command({ ping: 1 });
        await Promise.all([
          db.collection('uploaded_documents').createIndex({ id: 1 }, { unique: true }),
          db.collection('uploaded_documents').createIndex({ uploadedAt: -1 }),
          db.collection('generated_reports').createIndex({ id: 1 }, { unique: true }),
          db.collection('generated_reports').createIndex({ createdAt: -1 }),
        ]);

        this.client = client;
        this.db = db;
        this.mode = 'mongodb';
        this.lastError = undefined;
        console.log(`[SRC_CCE] Connected to MongoDB database "${DB_NAME}"`);
      } catch (err) {
        this.mode = 'local-fallback';
        this.lastError = err instanceof Error ? err.message : 'MongoDB connection unavailable';
        console.log(
          `[SRC_CCE] MongoDB not reachable at ${MONGO_URI.split('?')[0]} — using persistent local storage at ./data/ (${this.lastError})`
        );
      }

      this.initialized = true;
    })();

    try {
      await this.initPromise;
    } finally {
      this.initPromise = null;
    }
  }

  private enqueueWrite<T>(fn: () => Promise<T> | T): Promise<T> {
    const run = this.writeQueue.then(() => fn());
    this.writeQueue = run.catch(() => {});
    return run;
  }

  private ensureLocalDirs(): void {
    if (!fs.existsSync(DATA_DIR)) fs.mkdirSync(DATA_DIR, { recursive: true });
    if (!fs.existsSync(FILES_DIR)) fs.mkdirSync(FILES_DIR, { recursive: true });
    if (!fs.existsSync(LOCAL_DB_FILE)) {
      const initial: LocalStoreShape = {
        uploaded_documents: [],
        generated_reports: [],
      };
      fs.writeFileSync(LOCAL_DB_FILE, JSON.stringify(initial, null, 2), 'utf-8');
    }
  }

  private readLocalStore(): LocalStoreShape {
    this.ensureLocalDirs();
    try {
      const raw = fs.readFileSync(LOCAL_DB_FILE, 'utf-8');
      const parsed = JSON.parse(raw) as Partial<LocalStoreShape>;
      return {
        uploaded_documents: Array.isArray(parsed.uploaded_documents) ? parsed.uploaded_documents : [],
        generated_reports: Array.isArray(parsed.generated_reports) ? parsed.generated_reports : [],
      };
    } catch {
      return { uploaded_documents: [], generated_reports: [] };
    }
  }

  private writeLocalStore(data: LocalStoreShape): void {
    this.ensureLocalDirs();
    const tmpFile = `${LOCAL_DB_FILE}.${process.pid}.${Date.now()}.tmp`;
    fs.writeFileSync(tmpFile, JSON.stringify(data, null, 2), 'utf-8');
    fs.renameSync(tmpFile, LOCAL_DB_FILE);
  }

  async getStatus(forceReconnect = false): Promise<DatabaseStatusInfo> {
    await this.init(forceReconnect);
    const docs = await this.listDocuments();
    const reports = await this.listReports();

    return {
      mode: this.mode,
      databaseName: DB_NAME,
      connected: this.mode === 'mongodb',
      connectionLabel:
        this.mode === 'mongodb'
          ? `MongoDB (${DB_NAME})`
          : `Local Persistent Store (${DB_NAME} schema)`,
      documentsCount: docs.length,
      reportsCount: reports.length,
      lastError: this.lastError,
    };
  }

  async saveUploadedDocument(
    doc: UploadedDocumentRecord,
    rawBuffer?: Buffer
  ): Promise<UploadedDocumentRecord> {
    await this.init();
    if (!isValidRecordId(doc.id)) {
      throw new Error('Invalid document ID format.');
    }

    const sanitizedFile = doc.fileName.replace(/[^a-zA-Z0-9._-]/g, '_').slice(0, 100);
    const safeDiskName = `${doc.id}_${sanitizedFile || 'upload.dat'}`;
    const safeDiskPath = resolveSafeFilePath(safeDiskName);

    if (rawBuffer && rawBuffer.length > 0 && safeDiskPath) {
      this.ensureLocalDirs();
      fs.writeFileSync(safeDiskPath, rawBuffer);
    }

    if (this.mode === 'mongodb' && this.db) {
      const col = this.db.collection('uploaded_documents');
      await col.insertOne({
        ...doc,
        storedFileName: rawBuffer && safeDiskPath ? safeDiskName : undefined,
        fileBase64:
          rawBuffer && rawBuffer.length <= 8 * 1024 * 1024
            ? rawBuffer.toString('base64')
            : undefined,
      });
      return doc;
    }

    return this.enqueueWrite(() => {
      const store = this.readLocalStore();
      store.uploaded_documents.unshift({
        ...doc,
        storedFileName: rawBuffer && safeDiskPath ? safeDiskName : undefined,
      });
      this.writeLocalStore(store);
      return doc;
    });
  }

  async listDocuments(): Promise<UploadedDocumentRecord[]> {
    await this.init();
    if (this.mode === 'mongodb' && this.db) {
      const col = this.db.collection('uploaded_documents');
      const items = await col
        .find({}, { projection: { fileBase64: 0, storedFileName: 0, _id: 0 } })
        .sort({ uploadedAt: -1 })
        .toArray();
      return items as unknown as UploadedDocumentRecord[];
    }

    const store = this.readLocalStore();
    return store.uploaded_documents
      .map(({ storedFileName: _ignored, ...rest }) => rest)
      .sort((a, b) => b.uploadedAt.localeCompare(a.uploadedAt));
  }

  async getDocumentById(id: string): Promise<UploadedDocumentRecord | null> {
    if (!isValidRecordId(id)) return null;
    await this.init();
    if (this.mode === 'mongodb' && this.db) {
      const item = await this.db
        .collection('uploaded_documents')
        .findOne({ id }, { projection: { fileBase64: 0, storedFileName: 0, _id: 0 } });
      return (item as unknown as UploadedDocumentRecord) || null;
    }

    const store = this.readLocalStore();
    const found = store.uploaded_documents.find((d) => d.id === id);
    if (!found) return null;
    const { storedFileName: _ignored, ...rest } = found;
    return rest;
  }

  async getDocumentFileBuffer(
    id: string
  ): Promise<{ buffer: Buffer; fileName: string; mimeType: string } | null> {
    if (!isValidRecordId(id)) return null;
    await this.init();

    if (this.mode === 'mongodb' && this.db) {
      const item = await this.db.collection('uploaded_documents').findOne({ id });
      if (!item) return null;
      if (typeof item.fileBase64 === 'string') {
        return {
          buffer: Buffer.from(item.fileBase64, 'base64'),
          fileName: String(item.fileName),
          mimeType: String(item.mimeType || 'application/octet-stream'),
        };
      }
      if (typeof item.storedFileName === 'string') {
        const filePath = resolveSafeFilePath(item.storedFileName);
        if (filePath && fs.existsSync(filePath)) {
          return {
            buffer: fs.readFileSync(filePath),
            fileName: String(item.fileName),
            mimeType: String(item.mimeType || 'application/octet-stream'),
          };
        }
      }
      return null;
    }

    const store = this.readLocalStore();
    const found = store.uploaded_documents.find((d) => d.id === id);
    if (!found || !found.storedFileName) return null;
    const filePath = resolveSafeFilePath(found.storedFileName);
    if (!filePath || !fs.existsSync(filePath)) return null;
    return {
      buffer: fs.readFileSync(filePath),
      fileName: found.fileName,
      mimeType: found.mimeType || 'application/octet-stream',
    };
  }

  async deleteDocument(id: string): Promise<boolean> {
    if (!isValidRecordId(id)) return false;
    await this.init();

    if (this.mode === 'mongodb' && this.db) {
      const existing = await this.db.collection('uploaded_documents').findOne({ id });
      if (existing && typeof existing.storedFileName === 'string') {
        const filePath = resolveSafeFilePath(existing.storedFileName);
        if (filePath && fs.existsSync(filePath)) fs.unlinkSync(filePath);
      }
      const res = await this.db.collection('uploaded_documents').deleteOne({ id });
      return res.deletedCount > 0;
    }

    return this.enqueueWrite(() => {
      const store = this.readLocalStore();
      const idx = store.uploaded_documents.findIndex((d) => d.id === id);
      if (idx === -1) return false;
      const removed = store.uploaded_documents[idx];
      if (removed.storedFileName) {
        const filePath = resolveSafeFilePath(removed.storedFileName);
        if (filePath && fs.existsSync(filePath)) fs.unlinkSync(filePath);
      }
      store.uploaded_documents.splice(idx, 1);
      this.writeLocalStore(store);
      return true;
    });
  }

  async saveGeneratedReport(report: GeneratedReportSnapshot): Promise<GeneratedReportSnapshot> {
    await this.init();
    if (!isValidRecordId(report.id)) {
      throw new Error('Invalid report ID format.');
    }

    if (this.mode === 'mongodb' && this.db) {
      await this.db.collection('generated_reports').insertOne({ ...report });
      return report;
    }

    return this.enqueueWrite(() => {
      const store = this.readLocalStore();
      store.generated_reports.unshift(report);
      this.writeLocalStore(store);
      return report;
    });
  }

  async listReports(): Promise<GeneratedReportSnapshot[]> {
    await this.init();
    if (this.mode === 'mongodb' && this.db) {
      const items = await this.db
        .collection('generated_reports')
        .find({}, { projection: { _id: 0 } })
        .sort({ createdAt: -1 })
        .toArray();
      return items as unknown as GeneratedReportSnapshot[];
    }

    const store = this.readLocalStore();
    return [...store.generated_reports].sort((a, b) =>
      b.createdAt.localeCompare(a.createdAt)
    );
  }

  async getReportById(id: string): Promise<GeneratedReportSnapshot | null> {
    if (!isValidRecordId(id)) return null;
    await this.init();

    if (this.mode === 'mongodb' && this.db) {
      const item = await this.db
        .collection('generated_reports')
        .findOne({ id }, { projection: { _id: 0 } });
      return (item as unknown as GeneratedReportSnapshot) || null;
    }

    const store = this.readLocalStore();
    return store.generated_reports.find((r) => r.id === id) || null;
  }

  async deleteReport(id: string): Promise<boolean> {
    if (!isValidRecordId(id)) return false;
    await this.init();

    if (this.mode === 'mongodb' && this.db) {
      const res = await this.db.collection('generated_reports').deleteOne({ id });
      return res.deletedCount > 0;
    }

    return this.enqueueWrite(() => {
      const store = this.readLocalStore();
      const prevLen = store.generated_reports.length;
      store.generated_reports = store.generated_reports.filter((r) => r.id !== id);
      if (store.generated_reports.length === prevLen) return false;
      this.writeLocalStore(store);
      return true;
    });
  }
}

export const srcCceDb = new SrcCceDatabase();
