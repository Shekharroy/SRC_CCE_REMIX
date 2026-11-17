import 'dotenv/config';
import crypto from 'crypto';
import fs from 'fs';
import path from 'path';
import express, { type NextFunction, type Request, type Response } from 'express';
import multer from 'multer';
import { createServer as createViteServer } from 'vite';
import { isValidRecordId, srcCceDb } from './server/db.ts';
import {
  buildCustomReportBuffer,
  buildSampleTemplateBuffer,
  buildWorkbookBuffer,
} from './server/excelExport.ts';
import { importPublicGoogleSheet, parseUploadedDocument } from './server/parser.ts';
import {
  buildAllReports,
  DEFAULT_CRM_NAME,
  getTodayIsoDate,
} from './src/shared/reportEngine.ts';
import type {
  CustomReportConfig,
  EnquiryRow,
  GeneratedReportSnapshot,
  UploadedDocumentRecord,
} from './src/shared/types.ts';

const PORT = Number(process.env.PORT || 3000);

function safeAttachmentHeader(fileName: string): string {
  const asciiSafe = fileName
    .replace(/[\r\n"\\]/g, '_')
    .replace(/[^\x20-\x7E]/g, '_')
    .slice(0, 120) || 'download';
  const encoded = encodeURIComponent(fileName.replace(/[\r\n]/g, ''));
  return `attachment; filename="${asciiSafe}"; filename*=UTF-8''${encoded}`;
}

const upload = multer({
  storage: multer.memoryStorage(),
  limits: {
    fileSize: 25 * 1024 * 1024, // 25 MB per file
    files: 15,
  },
});

async function startServer() {
  await srcCceDb.init();

  const app = express();
  app.disable('x-powered-by');
  app.use((_req: Request, res: Response, next: NextFunction) => {
    res.setHeader('X-Content-Type-Options', 'nosniff');
    res.setHeader('Referrer-Policy', 'strict-origin-when-cross-origin');
    next();
  });

  app.use(express.json({ limit: '30mb' }));
  app.use(express.urlencoded({ extended: true, limit: '30mb' }));

  app.get('/api/health', async (req: Request, res: Response) => {
    try {
      const reconnect = req.query.reconnect === '1';
      const status = await srcCceDb.getStatus(reconnect);
      res.json({ ok: true, database: status });
    } catch (err) {
      res.status(500).json({
        ok: false,
        error: err instanceof Error ? err.message : 'Database health check failed',
      });
    }
  });

  app.get('/api/documents', async (_req: Request, res: Response) => {
    try {
      const documents = await srcCceDb.listDocuments();
      res.json({ documents });
    } catch (err) {
      res.status(500).json({
        error: err instanceof Error ? err.message : 'Failed to list documents',
      });
    }
  });

  app.post(
    '/api/upload',
    upload.array('files', 15),
    async (req: Request, res: Response) => {
      try {
        const files = req.files as Express.Multer.File[] | undefined;
        if (!files || files.length === 0) {
          res.status(400).json({ error: 'No files were uploaded.' });
          return;
        }

        const savedDocs: UploadedDocumentRecord[] = [];
        const errors: Array<{ fileName: string; message: string }> = [];

        for (const file of files) {
          const docId = `doc_${Date.now()}_${crypto.randomBytes(3).toString('hex')}`;
          try {
            const parsed = await parseUploadedDocument({
              buffer: file.buffer,
              fileName: file.originalname,
              mimeType: file.mimetype,
              docId,
            });

            const docRecord: UploadedDocumentRecord = {
              id: docId,
              fileName: file.originalname,
              fileKind: parsed.fileKind,
              mimeType: file.mimetype || 'application/octet-stream',
              sizeBytes: file.size,
              uploadedAt: new Date().toISOString(),
              rowCount: parsed.rows.length,
              extractionMethod: parsed.extractionMethod,
              rows: parsed.rows,
              hasStoredFile: true,
            };

            await srcCceDb.saveUploadedDocument(docRecord, file.buffer);
            savedDocs.push(docRecord);
          } catch (fileErr) {
            errors.push({
              fileName: file.originalname,
              message:
                fileErr instanceof Error
                  ? fileErr.message
                  : 'Unable to extract records from file',
            });
          }
        }

        if (savedDocs.length === 0 && errors.length > 0) {
          res.status(422).json({
            error: errors.map((e) => `${e.fileName}: ${e.message}`).join('; '),
            errors,
          });
          return;
        }

        res.json({
          documents: savedDocs,
          errors,
        });
      } catch (err) {
        res.status(500).json({
          error: err instanceof Error ? err.message : 'File upload failed',
        });
      }
    }
  );

  app.post('/api/ingest-text', async (req: Request, res: Response) => {
    try {
      const { fileName, content } = req.body as {
        fileName?: string;
        content?: string;
      };
      if (!content || !content.trim()) {
        res.status(400).json({ error: 'Text content is required.' });
        return;
      }

      const docId = `doc_${Date.now()}_${crypto.randomBytes(3).toString('hex')}`;
      const resolvedName =
        (fileName || '').trim() || `Pasted_Notes_${getTodayIsoDate()}.txt`;
      const buffer = Buffer.from(content, 'utf-8');

      const parsed = await parseUploadedDocument({
        buffer,
        fileName: resolvedName,
        mimeType: 'text/plain',
        docId,
      });

      const docRecord: UploadedDocumentRecord = {
        id: docId,
        fileName: resolvedName,
        fileKind: 'text',
        mimeType: 'text/plain',
        sizeBytes: buffer.length,
        uploadedAt: new Date().toISOString(),
        rowCount: parsed.rows.length,
        extractionMethod: parsed.extractionMethod,
        rows: parsed.rows,
        hasStoredFile: true,
      };

      await srcCceDb.saveUploadedDocument(docRecord, buffer);
      res.json({ document: docRecord });
    } catch (err) {
      res.status(422).json({
        error: err instanceof Error ? err.message : 'Text ingestion failed',
      });
    }
  });

  app.post('/api/import-google-sheet', async (req: Request, res: Response) => {
    try {
      const { url, label } = req.body as { url?: string; label?: string };
      if (!url || !url.trim()) {
        res.status(400).json({ error: 'Google Sheets URL is required.' });
        return;
      }

      const docId = `doc_${Date.now()}_${crypto.randomBytes(3).toString('hex')}`;
      const imported = await importPublicGoogleSheet({
        url: url.trim(),
        customLabel: label,
        docId,
      });

      const docRecord: UploadedDocumentRecord = {
        id: docId,
        fileName: imported.fileName,
        fileKind: 'google-sheet',
        mimeType: 'text/csv',
        sizeBytes: imported.buffer.length,
        uploadedAt: new Date().toISOString(),
        rowCount: imported.rows.length,
        extractionMethod: 'google-sheets',
        sourceUrl: url.trim(),
        rows: imported.rows,
        hasStoredFile: true,
      };

      await srcCceDb.saveUploadedDocument(docRecord, imported.buffer);
      res.json({ document: docRecord });
    } catch (err) {
      const msg = err instanceof Error ? err.message : 'Google Sheet import failed.';
      const status = msg.includes('private') || msg.includes('sign-in') ? 403 : 400;
      res.status(status).json({ error: msg });
    }
  });

  app.get('/api/documents/:id/download', async (req: Request, res: Response) => {
    try {
      const id = String(req.params.id);
      if (!isValidRecordId(id)) {
        res.status(400).json({ error: 'Invalid document ID.' });
        return;
      }
      const fileData = await srcCceDb.getDocumentFileBuffer(id);
      if (!fileData) {
        res.status(404).json({ error: 'Stored document file not found.' });
        return;
      }
      res.setHeader('Content-Type', fileData.mimeType);
      res.setHeader('Content-Disposition', safeAttachmentHeader(fileData.fileName));
      res.send(fileData.buffer);
    } catch (err) {
      res.status(500).json({
        error: err instanceof Error ? err.message : 'Failed to download document',
      });
    }
  });

  app.delete('/api/documents/:id', async (req: Request, res: Response) => {
    try {
      const id = String(req.params.id);
      if (!isValidRecordId(id)) {
        res.status(400).json({ error: 'Invalid document ID.' });
        return;
      }
      const deleted = await srcCceDb.deleteDocument(id);
      if (!deleted) {
        res.status(404).json({ error: 'Document not found.' });
        return;
      }
      res.json({ ok: true });
    } catch (err) {
      res.status(500).json({
        error: err instanceof Error ? err.message : 'Failed to delete document',
      });
    }
  });

  app.get('/api/reports', async (_req: Request, res: Response) => {
    try {
      const reports = await srcCceDb.listReports();
      res.json({ reports });
    } catch (err) {
      res.status(500).json({
        error: err instanceof Error ? err.message : 'Failed to list saved reports',
      });
    }
  });

  app.post('/api/reports/save', async (req: Request, res: Response) => {
    try {
      const body = req.body as {
        title?: string;
        targetDate?: string;
        crmTarget?: string;
        strictDateMatch?: boolean;
        sourceDocIds?: string[];
        sourceFileNames?: string[];
        rows?: EnquiryRow[];
      };

      const rows = Array.isArray(body.rows) ? body.rows : [];
      if (rows.length === 0) {
        res.status(400).json({ error: 'Cannot save an empty report snapshot.' });
        return;
      }

      const targetDate = body.targetDate || getTodayIsoDate();
      const crmTarget = body.crmTarget || DEFAULT_CRM_NAME;
      const bundle = buildAllReports(rows, {
        targetDate,
        crmTarget,
        strictDateMatch: body.strictDateMatch,
      });

      const activeCounters = bundle.report4CounterWise.summaries.filter(
        (c) => c.totalEnquiries > 0 && c.counterName !== 'Unassigned'
      ).length;

      const snapshot: GeneratedReportSnapshot = {
        id: `rep_${Date.now()}_${crypto.randomBytes(3).toString('hex')}`,
        title:
          (body.title || '').trim() ||
          `SRC_CCE Report (${targetDate})`,
        createdAt: new Date().toISOString(),
        targetDate,
        crmTarget,
        sourceDocIds: body.sourceDocIds || [],
        sourceFileNames: body.sourceFileNames || [],
        totalRecords: rows.length,
        summaryMetrics: {
          totalEnquiries: rows.length,
          todayFollowups: bundle.report1DseFollowup.totalFollowups,
          crmCalls: bundle.report3CrmDailyCall.totalCalls,
          crmConversions: bundle.report3CrmDailyCall.totalConversions,
          activeCounters,
          totalConverted: bundle.report4CounterWise.totalConverted,
        },
        bundle,
      };

      await srcCceDb.saveGeneratedReport(snapshot);
      res.json({ report: snapshot });
    } catch (err) {
      res.status(500).json({
        error: err instanceof Error ? err.message : 'Failed to save report snapshot',
      });
    }
  });

  app.delete('/api/reports/:id', async (req: Request, res: Response) => {
    try {
      const id = String(req.params.id);
      if (!isValidRecordId(id)) {
        res.status(400).json({ error: 'Invalid report ID.' });
        return;
      }
      const deleted = await srcCceDb.deleteReport(id);
      if (!deleted) {
        res.status(404).json({ error: 'Report snapshot not found.' });
        return;
      }
      res.json({ ok: true });
    } catch (err) {
      res.status(500).json({
        error: err instanceof Error ? err.message : 'Failed to delete report snapshot',
      });
    }
  });

  app.post('/api/export/excel', async (req: Request, res: Response) => {
    try {
      const body = req.body as {
        reportId?: string;
        rows?: EnquiryRow[];
        targetDate?: string;
        crmTarget?: string;
        strictDateMatch?: boolean;
        scope?: 'all' | 'report1' | 'report2' | 'report3' | 'report4';
      };

      const scope = body.scope || 'all';
      let bundle;
      let rows: EnquiryRow[] = [];

      if (body.reportId) {
        if (!isValidRecordId(body.reportId)) {
          res.status(400).json({ error: 'Invalid report ID.' });
          return;
        }
        const snapshot = await srcCceDb.getReportById(body.reportId);
        if (!snapshot) {
          res.status(404).json({ error: 'Saved report snapshot not found.' });
          return;
        }
        bundle = snapshot.bundle;
        rows = snapshot.bundle.report2PrincipleModel.rows || [];
      } else {
        rows = Array.isArray(body.rows) ? body.rows : [];
        if (rows.length === 0) {
          res.status(400).json({ error: 'No enquiry rows available to export.' });
          return;
        }
        bundle = buildAllReports(rows, {
          targetDate: body.targetDate || getTodayIsoDate(),
          crmTarget: body.crmTarget || DEFAULT_CRM_NAME,
          strictDateMatch: body.strictDateMatch,
        });
      }

      const buffer = await buildWorkbookBuffer(bundle, rows, scope);
      const scopeLabel =
        scope === 'all'
          ? 'Complete_4_Reports'
          : scope === 'report1'
          ? '1_DSE_Today_Followup'
          : scope === 'report2'
          ? '2_Principle_Model_Wise'
          : scope === 'report3'
          ? '3_CRM_Chandrashekhar_Report'
          : '4_Counter_Wise_Summary';

      const fileName = `SRC_CCE_${scopeLabel}_${bundle.targetDate}.xlsx`;
      res.setHeader(
        'Content-Type',
        'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
      );
      res.setHeader('Content-Disposition', safeAttachmentHeader(fileName));
      res.send(buffer);
    } catch (err) {
      res.status(500).json({
        error: err instanceof Error ? err.message : 'Failed to generate Excel export',
      });
    }
  });

  app.post('/api/export/custom-excel', async (req: Request, res: Response) => {
    try {
      const body = req.body as {
        rows?: EnquiryRow[];
        config?: CustomReportConfig;
      };
      const rows = Array.isArray(body.rows) ? body.rows : [];
      if (rows.length === 0) {
        res.status(400).json({ error: 'No enquiry rows provided for custom export.' });
        return;
      }
      if (!body.config) {
        res.status(400).json({ error: 'Custom report configuration is required.' });
        return;
      }

      const buffer = await buildCustomReportBuffer(rows, body.config);
      const safeTitle =
        (body.config.title || 'Custom_Report')
          .replace(/[^a-zA-Z0-9_-]+/g, '_')
          .slice(0, 40) || 'Custom_Report';
      const fileName = `SRC_CCE_${safeTitle}_${getTodayIsoDate()}.xlsx`;

      res.setHeader(
        'Content-Type',
        'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
      );
      res.setHeader('Content-Disposition', safeAttachmentHeader(fileName));
      res.send(buffer);
    } catch (err) {
      res.status(500).json({
        error: err instanceof Error ? err.message : 'Failed to generate custom Excel export',
      });
    }
  });

  app.get('/api/sample-template', async (_req: Request, res: Response) => {
    try {
      const buffer = await buildSampleTemplateBuffer();
      const fileName = `SRC_CCE_Sample_Enquiry_Sheet_${getTodayIsoDate()}.xlsx`;
      res.setHeader(
        'Content-Type',
        'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
      );
      res.setHeader('Content-Disposition', safeAttachmentHeader(fileName));
      res.send(buffer);
    } catch (err) {
      res.status(500).json({
        error: err instanceof Error ? err.message : 'Failed to generate sample Excel template',
      });
    }
  });

  // Structured JSON error handler (catches Multer size/file limits & malformed JSON bodies)
  app.use((err: unknown, _req: Request, res: Response, next: NextFunction) => {
    if (res.headersSent) {
      next(err);
      return;
    }
    if (err instanceof multer.MulterError) {
      const status = err.code === 'LIMIT_FILE_SIZE' ? 413 : 400;
      const msg =
        err.code === 'LIMIT_FILE_SIZE'
          ? 'File exceeds maximum 25 MB upload limit.'
          : err.code === 'LIMIT_FILE_COUNT'
          ? 'Too many files uploaded at once (max 15).'
          : `Upload error: ${err.message}`;
      res.status(status).json({ error: msg });
      return;
    }
    if (err instanceof SyntaxError && 'body' in err) {
      res.status(400).json({ error: 'Malformed JSON request payload.' });
      return;
    }
    res.status(500).json({
      error: err instanceof Error ? err.message : 'Internal server error.',
    });
  });

  // Vite dev middleware vs production static serving
  if (process.env.NODE_ENV !== 'production') {
    const vite = await createViteServer({
      server: { middlewareMode: true },
      appType: 'spa',
    });
    app.use(vite.middlewares);
  } else {
    const distDir = path.resolve(process.cwd(), 'dist');
    app.use(express.static(distDir));
    app.get('*', (_req, res) => {
      const indexFile = path.join(distDir, 'index.html');
      if (fs.existsSync(indexFile)) {
        res.sendFile(indexFile);
      } else {
        res.status(404).send('Production build not found. Run npm run build.');
      }
    });
  }

  app.listen(PORT, '0.0.0.0', () => {
    console.log(`SRC_CCE Reporting Server running at http://0.0.0.0:${PORT}`);
  });
}

startServer().catch((err) => {
  console.error('Failed to start server:', err);
  process.exit(1);
});
