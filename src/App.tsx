import { useCallback, useEffect, useMemo, useState } from 'react';
import { Download } from 'lucide-react';
import { CounterDirectoryView } from './components/CounterDirectoryView';
import { DocumentsArchiveView } from './components/DocumentsArchiveView';
import { ReportsSection } from './components/ReportsSection';
import { UploadPanel } from './components/UploadPanel';
import {
  buildAllReports,
  DEFAULT_CRM_NAME,
  getTodayIsoDate,
} from './shared/reportEngine';
import type {
  CustomReportConfig,
  DatabaseStatusInfo,
  EnquiryRow,
  GeneratedReportSnapshot,
  UploadedDocumentRecord,
} from './shared/types';

type ActiveSection = 'workspace' | 'documents' | 'directory';

export default function App() {
  const [activeSection, setActiveSection] = useState<ActiveSection>('workspace');
  const [dbStatus, setDbStatus] = useState<DatabaseStatusInfo | null>(null);
  const [documents, setDocuments] = useState<UploadedDocumentRecord[]>([]);
  const [selectedDocIds, setSelectedDocIds] = useState<string[]>([]);
  const [savedReports, setSavedReports] = useState<GeneratedReportSnapshot[]>([]);
  const [loadedSnapshot, setLoadedSnapshot] = useState<GeneratedReportSnapshot | null>(null);

  const [targetDate, setTargetDate] = useState<string>(getTodayIsoDate());
  const [strictDateMatch, setStrictDateMatch] = useState<boolean>(false);
  const [crmTarget, setCrmTarget] = useState<string>(DEFAULT_CRM_NAME);

  const [isInitialLoading, setIsInitialLoading] = useState<boolean>(true);
  const [isUploading, setIsUploading] = useState<boolean>(false);
  const [isSavingSnapshot, setIsSavingSnapshot] = useState<boolean>(false);
  const [snapshotSavedFeedback, setSnapshotSavedFeedback] = useState<boolean>(false);
  const [uploadError, setUploadError] = useState<string | null>(null);
  const [activeFilteredRows, setActiveFilteredRows] = useState<EnquiryRow[] | null>(null);

  const fetchAllState = useCallback(async () => {
    try {
      const [healthRes, docsRes, repsRes] = await Promise.all([
        fetch('/api/health'),
        fetch('/api/documents'),
        fetch('/api/reports'),
      ]);

      if (healthRes.ok) {
        const healthData = await healthRes.json();
        if (healthData.database) setDbStatus(healthData.database);
      }
      if (docsRes.ok) {
        const docsData = await docsRes.json();
        const docList: UploadedDocumentRecord[] = Array.isArray(docsData.documents)
          ? docsData.documents
          : [];
        setDocuments(docList);
        setSelectedDocIds((prev) => {
          if (prev.length === 0) return docList.map((d) => d.id);
          const validIds = new Set(docList.map((d) => d.id));
          const kept = prev.filter((id) => validIds.has(id));
          return kept.length > 0 ? kept : docList.map((d) => d.id);
        });
      }
      if (repsRes.ok) {
        const repsData = await repsRes.json();
        setSavedReports(Array.isArray(repsData.reports) ? repsData.reports : []);
      }
    } catch (err) {
      setUploadError(
        err instanceof Error ? err.message : 'Unable to load saved SRC_CCE records.'
      );
    } finally {
      setIsInitialLoading(false);
    }
  }, []);

  useEffect(() => {
    void fetchAllState();
  }, [fetchAllState]);

  const activeRows: EnquiryRow[] = useMemo(() => {
    if (loadedSnapshot) {
      return loadedSnapshot.bundle.report2PrincipleModel.rows || [];
    }
    const selectedSet = new Set(selectedDocIds);
    return documents
      .filter((d) => selectedSet.has(d.id))
      .flatMap((d) => d.rows || []);
  }, [documents, selectedDocIds, loadedSnapshot]);

  const bundle = useMemo(() => {
    if (loadedSnapshot) {
      return loadedSnapshot.bundle;
    }
    return buildAllReports(activeRows, {
      targetDate,
      crmTarget,
      strictDateMatch,
    });
  }, [activeRows, targetDate, crmTarget, strictDateMatch, loadedSnapshot]);

  const activeCountersCount = useMemo(() => {
    return bundle.report4CounterWise.summaries.filter(
      (c) => c.counterName !== 'Unassigned' && c.totalEnquiries > 0
    ).length;
  }, [bundle.report4CounterWise.summaries]);

  const handleUploadFiles = async (fileList: FileList | File[]) => {
    setIsUploading(true);
    setUploadError(null);
    setLoadedSnapshot(null);

    try {
      const formData = new FormData();
      Array.from(fileList).forEach((file) => {
        formData.append('files', file);
      });

      const res = await fetch('/api/upload', {
        method: 'POST',
        body: formData,
      });

      const data = await res.json();
      if (!res.ok) {
        throw new Error(data.error || 'Upload failed.');
      }

      const newlyAdded: UploadedDocumentRecord[] = Array.isArray(data.documents)
        ? data.documents
        : [];

      if (Array.isArray(data.errors) && data.errors.length > 0) {
        setUploadError(
          `Partial upload warnings: ${data.errors
            .map((e: { fileName: string; message: string }) => `${e.fileName}: ${e.message}`)
            .join('; ')}`
        );
      }

      if (newlyAdded.length > 0) {
        setSelectedDocIds((prev) => [
          ...newlyAdded.map((d) => d.id),
          ...prev,
        ]);
      }

      await fetchAllState();
    } catch (err) {
      setUploadError(err instanceof Error ? err.message : 'File upload failed.');
    } finally {
      setIsUploading(false);
    }
  };

  const handleIngestText = async (fileName: string, content: string) => {
    setIsUploading(true);
    setUploadError(null);
    setLoadedSnapshot(null);

    try {
      const res = await fetch('/api/ingest-text', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ fileName, content }),
      });

      const text = await res.text();
      const data = text ? JSON.parse(text) : {};
      if (!res.ok) {
        throw new Error(data.error || 'Failed to extract records from text.');
      }

      if (data.document?.id) {
        setSelectedDocIds((prev) => [data.document.id, ...prev]);
      }
      await fetchAllState();
    } catch (err) {
      setUploadError(
        err instanceof Error ? err.message : 'Text extraction failed.'
      );
    } finally {
      setIsUploading(false);
    }
  };

  const handleImportGoogleSheet = async (url: string, label: string) => {
    setIsUploading(true);
    setUploadError(null);
    setLoadedSnapshot(null);

    try {
      const res = await fetch('/api/import-google-sheet', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ url, label }),
      });

      const text = await res.text();
      const data = text ? JSON.parse(text) : {};
      if (!res.ok) {
        throw new Error(data.error || 'Failed to import Google Sheet.');
      }

      if (data.document?.id) {
        setSelectedDocIds((prev) => [data.document.id, ...prev]);
      }
      await fetchAllState();
    } catch (err) {
      setUploadError(
        err instanceof Error ? err.message : 'Google Sheet import failed.'
      );
    } finally {
      setIsUploading(false);
    }
  };

  const handleToggleDocSelection = (docId: string) => {
    setLoadedSnapshot(null);
    setSelectedDocIds((prev) =>
      prev.includes(docId) ? prev.filter((id) => id !== docId) : [...prev, docId]
    );
  };

  const handleSelectAllDocs = () => {
    setLoadedSnapshot(null);
    setSelectedDocIds(documents.map((d) => d.id));
  };

  const handleDeleteDocument = async (id: string) => {
    try {
      const res = await fetch(`/api/documents/${encodeURIComponent(id)}`, {
        method: 'DELETE',
      });
      if (res.ok) {
        setSelectedDocIds((prev) => prev.filter((item) => item !== id));
        await fetchAllState();
      }
    } catch (err) {
      setUploadError(
        err instanceof Error ? err.message : 'Failed to delete document.'
      );
    }
  };

  const handleSaveSnapshot = async (title: string) => {
    if (activeRows.length === 0) return;
    setIsSavingSnapshot(true);
    setUploadError(null);

    try {
      const selectedSet = new Set(selectedDocIds);
      const sourceDocs = documents.filter((d) => selectedSet.has(d.id));
      const res = await fetch('/api/reports/save', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({
          title,
          targetDate,
          crmTarget,
          strictDateMatch,
          sourceDocIds: sourceDocs.map((d) => d.id),
          sourceFileNames: sourceDocs.map((d) => d.fileName),
          rows: activeRows,
        }),
      });
      const data = await res.json();
      if (!res.ok) {
        throw new Error(data.error || 'Failed to save snapshot.');
      }

      setSnapshotSavedFeedback(true);
      setTimeout(() => setSnapshotSavedFeedback(false), 2500);
      await fetchAllState();
    } catch (err) {
      setUploadError(
        err instanceof Error ? err.message : 'Failed to save report snapshot.'
      );
    } finally {
      setIsSavingSnapshot(false);
    }
  };

  const handleDeleteReport = async (id: string) => {
    try {
      const res = await fetch(`/api/reports/${encodeURIComponent(id)}`, {
        method: 'DELETE',
      });
      if (res.ok) {
        if (loadedSnapshot?.id === id) setLoadedSnapshot(null);
        await fetchAllState();
      }
    } catch (err) {
      setUploadError(
        err instanceof Error ? err.message : 'Failed to delete saved report.'
      );
    }
  };

  const triggerBlobDownload = async (res: Response, fallbackName: string) => {
    const blob = await res.blob();
    const disposition = res.headers.get('Content-Disposition') || '';
    const match = disposition.match(/filename="([^"]+)"/);
    const filename = match?.[1] || fallbackName;
    const url = URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url;
    a.download = filename;
    document.body.appendChild(a);
    a.click();
    a.remove();
    URL.revokeObjectURL(url);
  };

  const handleExportExcel = async (
    scope: 'all' | 'report1' | 'report2' | 'report3' | 'report4',
    overrideRows?: EnquiryRow[]
  ) => {
    const rowsToExport = overrideRows ?? activeFilteredRows;
    if (bundle.totalRecords === 0) return;
    if (rowsToExport && rowsToExport.length === 0) return;
    try {
      const res = await fetch('/api/export/excel', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify(
          rowsToExport
            ? {
                rows: rowsToExport,
                targetDate,
                crmTarget,
                strictDateMatch,
                scope,
              }
            : loadedSnapshot
            ? { reportId: loadedSnapshot.id, scope }
            : {
                rows: activeRows,
                targetDate,
                crmTarget,
                strictDateMatch,
                scope,
              }
        ),
      });
      if (!res.ok) {
        const errData = await res.json().catch(() => ({}));
        throw new Error(errData.error || 'Excel export failed.');
      }
      await triggerBlobDownload(res, `SRC_CCE_${scope}_${targetDate}.xlsx`);
    } catch (err) {
      setUploadError(err instanceof Error ? err.message : 'Excel export failed.');
    }
  };

  const handleExportCustomExcel = async (config: CustomReportConfig) => {
    if (activeRows.length === 0) return;
    try {
      const res = await fetch('/api/export/custom-excel', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ rows: activeRows, config }),
      });
      if (!res.ok) {
        const errData = await res.json().catch(() => ({}));
        throw new Error(errData.error || 'Custom Excel export failed.');
      }
      await triggerBlobDownload(res, `SRC_CCE_Custom_Report_${targetDate}.xlsx`);
    } catch (err) {
      setUploadError(
        err instanceof Error ? err.message : 'Custom Excel export failed.'
      );
    }
  };

  const handleExportSavedReport = async (reportId: string) => {
    try {
      const res = await fetch('/api/export/excel', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ reportId, scope: 'all' }),
      });
      if (!res.ok) {
        throw new Error('Failed to export saved report.');
      }
      await triggerBlobDownload(res, `SRC_CCE_Saved_Report_${reportId}.xlsx`);
    } catch (err) {
      setUploadError(
        err instanceof Error ? err.message : 'Failed to export saved report.'
      );
    }
  };

  const handleNavClick = (section: ActiveSection, anchorId?: string) => {
    setActiveSection(section);
    if (anchorId) {
      requestAnimationFrame(() => {
        const el = document.getElementById(anchorId);
        if (el) el.scrollIntoView({ behavior: 'smooth', block: 'start' });
      });
    }
  };

  return (
    <div className="min-h-screen bg-slate-50 text-slate-900 flex flex-col">
      {/* Top Bar Contract: 3 zones separated by gap-8 */}
      <header className="sticky top-0 z-20 bg-white border-b border-slate-200">
        <div className="max-w-[1440px] mx-auto px-6 py-3.5 flex items-center justify-between gap-8">
          <a
            href="#workspace"
            onClick={(e) => {
              e.preventDefault();
              handleNavClick('workspace', 'workspace');
            }}
            className="text-base font-bold tracking-tight text-slate-900 whitespace-nowrap shrink-0 font-display"
          >
            SRC CCE Analytics
          </a>

          <nav className="hidden md:flex items-center gap-7 text-sm font-medium text-slate-600">
            <button
              type="button"
              onClick={() => handleNavClick('workspace', 'workspace')}
              className={`hover:text-slate-900 transition-colors whitespace-nowrap shrink-0 cursor-pointer ${
                activeSection === 'workspace' ? 'text-slate-900 underline underline-offset-8 decoration-blue-700 decoration-2' : ''
              }`}
            >
              Workspace
            </button>
            <button
              type="button"
              onClick={() => handleNavClick('workspace', 'reports')}
              className="hover:text-slate-900 transition-colors whitespace-nowrap shrink-0 cursor-pointer"
            >
              Reports
            </button>
            <button
              type="button"
              onClick={() => handleNavClick('documents')}
              className={`hover:text-slate-900 transition-colors whitespace-nowrap shrink-0 cursor-pointer ${
                activeSection === 'documents' ? 'text-slate-900 underline underline-offset-8 decoration-blue-700 decoration-2' : ''
              }`}
            >
              Documents &amp; DB
            </button>
            <button
              type="button"
              onClick={() => handleNavClick('directory')}
              className={`hover:text-slate-900 transition-colors whitespace-nowrap shrink-0 cursor-pointer ${
                activeSection === 'directory' ? 'text-slate-900 underline underline-offset-8 decoration-blue-700 decoration-2' : ''
              }`}
            >
              Counter Directory
            </button>
          </nav>

          <div className="flex items-center gap-3 shrink-0">
            <button
              type="button"
              disabled={
                bundle.totalRecords === 0 ||
                (activeFilteredRows !== null && activeFilteredRows.length === 0)
              }
              onClick={() => handleExportExcel('all')}
              className="inline-flex items-center gap-2 px-4 py-2 text-xs font-medium text-white bg-blue-700 hover:bg-blue-800 disabled:opacity-40 rounded-md transition-colors whitespace-nowrap shrink-0 cursor-pointer"
            >
              <Download className="w-3.5 h-3.5" />
              <span>{activeFilteredRows ? 'Export Filtered Excel' : 'Export All Excel'}</span>
            </button>
          </div>
        </div>
      </header>

      <main className="flex-1 max-w-[1440px] w-full mx-auto px-6 py-6 space-y-6">
        {loadedSnapshot && (
          <div className="bg-blue-50 border border-blue-200 rounded-md px-4 py-3 flex flex-wrap items-center justify-between gap-3 text-xs text-blue-900">
            <span>
              Viewing saved snapshot: <strong>{loadedSnapshot.title}</strong> ({loadedSnapshot.totalRecords} records, saved {loadedSnapshot.createdAt.slice(0, 10)})
            </span>
            <button
              type="button"
              onClick={() => setLoadedSnapshot(null)}
              className="font-medium text-blue-800 hover:underline cursor-pointer"
            >
              Return to Live Uploaded Documents
            </button>
          </div>
        )}

        {/* 4-Column KPI Strip */}
        <section className="grid grid-cols-1 sm:grid-cols-2 lg:grid-cols-4 bg-white border border-slate-200 rounded-lg divide-y sm:divide-y-0 sm:divide-x divide-slate-200">
          <div className="p-5">
            <p className="text-xs font-medium text-slate-500">Total Enquiries Loaded</p>
            <p className="mt-1.5 text-2xl font-semibold tracking-tight font-mono tabular-nums text-slate-900">
              {isInitialLoading ? '...' : bundle.totalRecords}
            </p>
            <p className="mt-1 text-xs text-slate-500">
              {documents.length} source files in SRC_CCE
            </p>
          </div>

          <div className="p-5">
            <p className="text-xs font-medium text-slate-500">DSE Follow-ups ({targetDate})</p>
            <p className="mt-1.5 text-2xl font-semibold tracking-tight font-mono tabular-nums text-slate-900">
              {isInitialLoading ? '...' : bundle.report1DseFollowup.totalFollowups}
            </p>
            <p className="mt-1 text-xs text-slate-500">
              Across {bundle.report1DseFollowup.dseCount} DSEs with remarks
            </p>
          </div>

          <div className="p-5">
            <p className="text-xs font-medium text-slate-500">CRM {crmTarget} Calls</p>
            <p className="mt-1.5 text-2xl font-semibold tracking-tight font-mono tabular-nums text-slate-900">
              {isInitialLoading ? '...' : bundle.report3CrmDailyCall.totalCalls}
            </p>
            <p className="mt-1 text-xs text-slate-500">
              {bundle.report3CrmDailyCall.totalConversions} conversions ({bundle.report3CrmDailyCall.conversionRate}%)
            </p>
          </div>

          <div className="p-5">
            <p className="text-xs font-medium text-slate-500">Counter Coverage &amp; Conversions</p>
            <p className="mt-1.5 text-2xl font-semibold tracking-tight font-mono tabular-nums text-slate-900">
              {isInitialLoading ? '...' : `${activeCountersCount} / 7`}
            </p>
            <p className="mt-1 text-xs text-slate-500">
              {bundle.report4CounterWise.totalConverted} total converted ({bundle.report4CounterWise.overallConversionRate}%)
            </p>
          </div>
        </section>

        {activeSection === 'workspace' && (
          <>
            <UploadPanel
              documents={documents}
              selectedDocIds={selectedDocIds}
              onToggleDocSelection={handleToggleDocSelection}
              onSelectAllDocs={handleSelectAllDocs}
              onUploadFiles={handleUploadFiles}
              onIngestText={handleIngestText}
              onImportGoogleSheet={handleImportGoogleSheet}
              isUploading={isUploading}
              uploadError={uploadError}
              onDismissError={() => setUploadError(null)}
            />

            <ReportsSection
              bundle={bundle}
              rawRows={activeRows}
              targetDate={targetDate}
              onChangeTargetDate={(d) => {
                setLoadedSnapshot(null);
                setTargetDate(d);
              }}
              strictDateMatch={strictDateMatch}
              onToggleStrictDate={(s) => {
                setLoadedSnapshot(null);
                setStrictDateMatch(s);
              }}
              crmTarget={crmTarget}
              onChangeCrmTarget={(c) => {
                setLoadedSnapshot(null);
                setCrmTarget(c);
              }}
              onExportExcel={handleExportExcel}
              onExportCustomExcel={handleExportCustomExcel}
              onSaveSnapshot={handleSaveSnapshot}
              isSavingSnapshot={isSavingSnapshot}
              snapshotSavedFeedback={snapshotSavedFeedback}
              onActiveFilteredRowsChange={setActiveFilteredRows}
              onOpenUpload={() => {
                const el = document.getElementById('workspace');
                if (el) el.scrollIntoView({ behavior: 'smooth' });
              }}
            />
          </>
        )}

        {activeSection === 'documents' && (
          <DocumentsArchiveView
            dbStatus={dbStatus}
            documents={documents}
            reports={savedReports}
            onDeleteDocument={handleDeleteDocument}
            onDeleteReport={handleDeleteReport}
            onLoadReportSnapshot={(snapshot) => {
              setLoadedSnapshot(snapshot);
              setActiveSection('workspace');
            }}
            onExportSavedReport={handleExportSavedReport}
          />
        )}

        {activeSection === 'directory' && (
          <CounterDirectoryView counterWiseReport={bundle.report4CounterWise} />
        )}
      </main>

      <footer className="border-t border-slate-200 bg-white mt-8">
        <div className="max-w-[1440px] mx-auto px-6 py-4 flex flex-col sm:flex-row items-center justify-between gap-2 text-xs text-slate-500">
          <span>SRC CCE Analytics · Dealership Enquiry &amp; CRM Reporting System</span>
          <span>
            Database: {dbStatus?.databaseName || 'SRC_CCE'} ({dbStatus?.mode === 'mongodb' ? 'MongoDB Connected' : 'Local Persistent Store'})
          </span>
        </div>
      </footer>
    </div>
  );
}
