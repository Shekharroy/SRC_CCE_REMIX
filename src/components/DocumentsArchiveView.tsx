import { Database, Download, Trash2, FolderOpen } from 'lucide-react';
import type {
  DatabaseStatusInfo,
  GeneratedReportSnapshot,
  UploadedDocumentRecord,
} from '../shared/types';

interface DocumentsArchiveViewProps {
  dbStatus: DatabaseStatusInfo | null;
  documents: UploadedDocumentRecord[];
  reports: GeneratedReportSnapshot[];
  onDeleteDocument: (id: string) => Promise<void>;
  onDeleteReport: (id: string) => Promise<void>;
  onLoadReportSnapshot: (snapshot: GeneratedReportSnapshot) => void;
  onExportSavedReport: (reportId: string) => Promise<void>;
}

function formatBytes(bytes: number): string {
  if (bytes < 1024) return `${bytes} B`;
  if (bytes < 1024 * 1024) return `${(bytes / 1024).toFixed(1)} KB`;
  return `${(bytes / (1024 * 1024)).toFixed(2)} MB`;
}

export function DocumentsArchiveView({
  dbStatus,
  documents,
  reports,
  onDeleteDocument,
  onDeleteReport,
  onLoadReportSnapshot,
  onExportSavedReport,
}: DocumentsArchiveViewProps) {
  return (
    <section id="documents" className="space-y-6">
      {/* Database Connection Overview */}
      <div className="bg-white border border-slate-200 rounded-lg p-5 flex flex-col md:flex-row md:items-center justify-between gap-4">
        <div className="flex items-start gap-3">
          <div className="w-9 h-9 rounded-md bg-slate-100 border border-slate-200 flex items-center justify-center text-slate-700 shrink-0">
            <Database className="w-4 h-4" />
          </div>
          <div>
            <h2 className="text-sm font-semibold text-slate-900">
              SRC_CCE Document &amp; Report Storage
            </h2>
            <p className="text-xs text-slate-600 mt-0.5">
              {dbStatus
                ? `${dbStatus.connectionLabel} · Collections: uploaded_documents (${documents.length}), generated_reports (${reports.length})`
                : 'Connecting to SRC_CCE storage...'}
            </p>
          </div>
        </div>
        <div className="text-xs font-mono text-slate-500 whitespace-nowrap">
          DB: {dbStatus?.databaseName || 'SRC_CCE'} · Mode: {dbStatus?.mode === 'mongodb' ? 'MongoDB Direct' : 'Local Persistent Fallback'}
        </div>
      </div>

      {/* Stored Uploaded Documents */}
      <div className="bg-white border border-slate-200 rounded-lg">
        <div className="px-5 py-4 border-b border-slate-200 flex items-center justify-between">
          <div>
            <h3 className="text-sm font-semibold text-slate-900">
              Stored Source Documents ({documents.length})
            </h3>
            <p className="text-xs text-slate-500 mt-0.5">
              Original uploaded Excel, CSV, Text, and Image files archived in SRC_CCE.
            </p>
          </div>
        </div>

        <div className="overflow-x-auto">
          <table className="w-full text-left border-collapse text-xs">
            <thead>
              <tr className="bg-slate-50 border-b border-slate-200 text-slate-600 font-medium">
                <th className="py-2.5 px-4">File Name</th>
                <th className="py-2.5 px-4">Format</th>
                <th className="py-2.5 px-4">Extraction</th>
                <th className="py-2.5 px-4 text-right">Rows</th>
                <th className="py-2.5 px-4 text-right">Size</th>
                <th className="py-2.5 px-4">Uploaded At</th>
                <th className="py-2.5 px-4 text-right">Actions</th>
              </tr>
            </thead>
            <tbody className="divide-y divide-slate-200">
              {documents.length === 0 ? (
                <tr>
                  <td colSpan={7} className="py-8 text-center text-slate-500">
                    No uploaded documents stored in SRC_CCE yet.
                  </td>
                </tr>
              ) : (
                documents.map((doc) => (
                  <tr key={doc.id} className="hover:bg-slate-50/80">
                    <td className="py-2.5 px-4 font-medium text-slate-900">
                      {doc.fileName}
                    </td>
                    <td className="py-2.5 px-4 text-slate-600 capitalize">
                      {doc.fileKind}
                    </td>
                    <td className="py-2.5 px-4 text-slate-600">
                      {doc.extractionMethod === 'gemini-ai'
                        ? 'AI Document Extraction'
                        : doc.extractionMethod === 'spreadsheet'
                        ? 'Spreadsheet Parser'
                        : 'Delimited Text Parser'}
                    </td>
                    <td className="py-2.5 px-4 text-right font-mono tabular-nums font-medium text-slate-900">
                      {doc.rowCount}
                    </td>
                    <td className="py-2.5 px-4 text-right font-mono tabular-nums text-slate-600">
                      {formatBytes(doc.sizeBytes)}
                    </td>
                    <td className="py-2.5 px-4 font-mono tabular-nums text-slate-600 whitespace-nowrap">
                      {doc.uploadedAt.slice(0, 16).replace('T', ' ')}
                    </td>
                    <td className="py-2.5 px-4 text-right">
                      <div className="inline-flex items-center gap-3">
                        <a
                          href={`/api/documents/${encodeURIComponent(doc.id)}/download`}
                          download
                          className="inline-flex items-center gap-1 text-blue-700 hover:underline font-medium"
                        >
                          <Download className="w-3.5 h-3.5" />
                          <span>File</span>
                        </a>
                        <button
                          type="button"
                          onClick={() => onDeleteDocument(doc.id)}
                          className="inline-flex items-center gap-1 text-red-700 hover:underline font-medium cursor-pointer"
                        >
                          <Trash2 className="w-3.5 h-3.5" />
                          <span>Delete</span>
                        </button>
                      </div>
                    </td>
                  </tr>
                ))
              )}
            </tbody>
          </table>
        </div>
      </div>

      {/* Saved Report Snapshots */}
      <div className="bg-white border border-slate-200 rounded-lg">
        <div className="px-5 py-4 border-b border-slate-200">
          <h3 className="text-sm font-semibold text-slate-900">
            Saved Report Snapshots ({reports.length})
          </h3>
          <p className="text-xs text-slate-500 mt-0.5">
            Historical 4-report bundles persisted in SRC_CCE. Reload any snapshot or download its multi-sheet Excel workbook.
          </p>
        </div>

        <div className="overflow-x-auto">
          <table className="w-full text-left border-collapse text-xs">
            <thead>
              <tr className="bg-slate-50 border-b border-slate-200 text-slate-600 font-medium">
                <th className="py-2.5 px-4">Report Title</th>
                <th className="py-2.5 px-4">Target Date</th>
                <th className="py-2.5 px-4 text-right">Enquiries</th>
                <th className="py-2.5 px-4 text-right">Follow-ups</th>
                <th className="py-2.5 px-4 text-right">CRM Calls</th>
                <th className="py-2.5 px-4 text-right">Converted</th>
                <th className="py-2.5 px-4">Saved At</th>
                <th className="py-2.5 px-4 text-right">Actions</th>
              </tr>
            </thead>
            <tbody className="divide-y divide-slate-200">
              {reports.length === 0 ? (
                <tr>
                  <td colSpan={8} className="py-8 text-center text-slate-500">
                    No saved report snapshots yet. Click &quot;Save Report to DB&quot; in the Reports toolbar to archive a snapshot.
                  </td>
                </tr>
              ) : (
                reports.map((rep) => (
                  <tr key={rep.id} className="hover:bg-slate-50/80">
                    <td className="py-2.5 px-4 font-medium text-slate-900">
                      {rep.title}
                    </td>
                    <td className="py-2.5 px-4 font-mono tabular-nums text-slate-600">
                      {rep.targetDate}
                    </td>
                    <td className="py-2.5 px-4 text-right font-mono tabular-nums font-medium text-slate-900">
                      {rep.summaryMetrics.totalEnquiries}
                    </td>
                    <td className="py-2.5 px-4 text-right font-mono tabular-nums text-blue-700">
                      {rep.summaryMetrics.todayFollowups}
                    </td>
                    <td className="py-2.5 px-4 text-right font-mono tabular-nums text-slate-700">
                      {rep.summaryMetrics.crmCalls}
                    </td>
                    <td className="py-2.5 px-4 text-right font-mono tabular-nums text-emerald-700">
                      {rep.summaryMetrics.totalConverted}
                    </td>
                    <td className="py-2.5 px-4 font-mono tabular-nums text-slate-600 whitespace-nowrap">
                      {rep.createdAt.slice(0, 16).replace('T', ' ')}
                    </td>
                    <td className="py-2.5 px-4 text-right">
                      <div className="inline-flex items-center gap-3">
                        <button
                          type="button"
                          onClick={() => onLoadReportSnapshot(rep)}
                          className="inline-flex items-center gap-1 text-slate-700 hover:underline font-medium cursor-pointer"
                        >
                          <FolderOpen className="w-3.5 h-3.5" />
                          <span>View</span>
                        </button>
                        <button
                          type="button"
                          onClick={() => onExportSavedReport(rep.id)}
                          className="inline-flex items-center gap-1 text-blue-700 hover:underline font-medium cursor-pointer"
                        >
                          <Download className="w-3.5 h-3.5" />
                          <span>Excel</span>
                        </button>
                        <button
                          type="button"
                          onClick={() => onDeleteReport(rep.id)}
                          className="inline-flex items-center gap-1 text-red-700 hover:underline font-medium cursor-pointer"
                        >
                          <Trash2 className="w-3.5 h-3.5" />
                          <span>Delete</span>
                        </button>
                      </div>
                    </td>
                  </tr>
                ))
              )}
            </tbody>
          </table>
        </div>
      </div>
    </section>
  );
}
