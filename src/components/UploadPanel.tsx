import { useRef, useState, type DragEvent, type ChangeEvent } from 'react';
import {
  FileSpreadsheet,
  FileText,
  Image as ImageIcon,
  Upload,
  Download,
  ClipboardPaste,
  Link2,
  Loader2,
  X,
} from 'lucide-react';
import type { UploadedDocumentRecord } from '../shared/types';

interface UploadPanelProps {
  documents: UploadedDocumentRecord[];
  selectedDocIds: string[];
  onToggleDocSelection: (docId: string) => void;
  onSelectAllDocs: () => void;
  onUploadFiles: (files: FileList | File[]) => Promise<void>;
  onIngestText: (fileName: string, content: string) => Promise<void>;
  onImportGoogleSheet: (url: string, label: string) => Promise<void>;
  isUploading: boolean;
  uploadError: string | null;
  onDismissError: () => void;
}

export function UploadPanel({
  documents,
  selectedDocIds,
  onToggleDocSelection,
  onSelectAllDocs,
  onUploadFiles,
  onIngestText,
  onImportGoogleSheet,
  isUploading,
  uploadError,
  onDismissError,
}: UploadPanelProps) {
  const fileInputRef = useRef<HTMLInputElement>(null);
  const [isDragging, setIsDragging] = useState(false);
  const [showPasteBox, setShowPasteBox] = useState(false);
  const [showSheetBox, setShowSheetBox] = useState(false);
  const [pasteTitle, setPasteTitle] = useState('');
  const [pasteContent, setPasteContent] = useState('');
  const [sheetUrl, setSheetUrl] = useState('');
  const [sheetLabel, setSheetLabel] = useState('');

  const handleDragOver = (e: DragEvent<HTMLDivElement>) => {
    e.preventDefault();
    setIsDragging(true);
  };

  const handleDragLeave = () => {
    setIsDragging(false);
  };

  const handleDrop = async (e: DragEvent<HTMLDivElement>) => {
    e.preventDefault();
    setIsDragging(false);
    if (e.dataTransfer.files && e.dataTransfer.files.length > 0) {
      await onUploadFiles(e.dataTransfer.files);
    }
  };

  const handleFileChange = async (e: ChangeEvent<HTMLInputElement>) => {
    if (e.target.files && e.target.files.length > 0) {
      await onUploadFiles(e.target.files);
      e.target.value = '';
    }
  };

  const handlePasteSubmit = async () => {
    if (!pasteContent.trim()) return;
    await onIngestText(pasteTitle.trim() || 'Pasted_CRM_Notes.txt', pasteContent);
    setPasteTitle('');
    setPasteContent('');
    setShowPasteBox(false);
  };

  const handleSheetSubmit = async () => {
    if (!sheetUrl.trim()) return;
    await onImportGoogleSheet(sheetUrl.trim(), sheetLabel.trim());
    setSheetUrl('');
    setSheetLabel('');
    setShowSheetBox(false);
  };

  return (
    <section
      id="workspace"
      className="bg-white border border-slate-200 rounded-lg p-6"
    >
      <div className="flex flex-col lg:flex-row lg:items-center lg:justify-between gap-4 pb-5 border-b border-slate-200">
        <div>
          <h1 className="text-xl font-semibold tracking-tight text-slate-900">
            Document Ingestion &amp; Multi-Format Report Builder
          </h1>
          <p className="mt-1 text-sm text-slate-600 max-w-2xl">
            Upload Excel sheets, CSV exports, plain text call logs, register images, or Google Sheets links. Records are normalized across all 7 counters and persisted in SRC_CCE.
          </p>
        </div>
        <div className="flex flex-wrap items-center gap-3 shrink-0">
          <button
            type="button"
            onClick={() => {
              setShowSheetBox((v) => !v);
              setShowPasteBox(false);
            }}
            className="inline-flex items-center gap-2 px-4 py-2 text-xs font-medium text-slate-700 bg-white border border-slate-300 rounded-md hover:bg-slate-50 transition-colors whitespace-nowrap shrink-0 cursor-pointer"
          >
            <Link2 className="w-3.5 h-3.5" />
            <span>{showSheetBox ? 'Close Sheet Link' : 'Google Sheets Link'}</span>
          </button>
          <button
            type="button"
            onClick={() => {
              setShowPasteBox((v) => !v);
              setShowSheetBox(false);
            }}
            className="inline-flex items-center gap-2 px-4 py-2 text-xs font-medium text-slate-700 bg-white border border-slate-300 rounded-md hover:bg-slate-50 transition-colors whitespace-nowrap shrink-0 cursor-pointer"
          >
            <ClipboardPaste className="w-3.5 h-3.5" />
            <span>{showPasteBox ? 'Close Text Input' : 'Paste Text Notes'}</span>
          </button>
          <a
            href="/api/sample-template"
            download
            className="inline-flex items-center gap-2 px-4 py-2 text-xs font-medium text-slate-700 bg-slate-100 hover:bg-slate-200 rounded-md transition-colors whitespace-nowrap shrink-0"
          >
            <Download className="w-3.5 h-3.5" />
            <span>Download Sample Excel</span>
          </a>
        </div>
      </div>

      {uploadError && (
        <div className="mt-4 flex items-start justify-between gap-3 px-4 py-3 bg-red-50 border border-red-200 rounded-md text-sm text-red-800">
          <span>{uploadError}</span>
          <button
            type="button"
            onClick={onDismissError}
            className="text-red-600 hover:text-red-800 p-0.5 cursor-pointer"
            aria-label="Dismiss error"
          >
            <X className="w-4 h-4" />
          </button>
        </div>
      )}

      <div
        onDragOver={handleDragOver}
        onDragLeave={handleDragLeave}
        onDrop={handleDrop}
        className={`mt-5 border border-dashed rounded-md p-6 transition-colors ${
          isDragging
            ? 'border-blue-600 bg-blue-50/50'
            : 'border-slate-300 bg-slate-50/60 hover:border-slate-400'
        }`}
      >
        <input
          ref={fileInputRef}
          type="file"
          multiple
          accept=".xlsx,.xls,.csv,.txt,.tsv,.log,.png,.jpg,.jpeg,.webp"
          onChange={handleFileChange}
          className="hidden"
        />

        <div className="flex flex-col md:flex-row items-start md:items-center justify-between gap-4">
          <div className="flex items-start gap-4">
            <div className="w-10 h-10 rounded-md bg-white border border-slate-200 flex items-center justify-center text-blue-700 shrink-0">
              {isUploading ? (
                <Loader2 className="w-5 h-5 animate-spin" />
              ) : (
                <Upload className="w-5 h-5" />
              )}
            </div>
            <div>
              <p className="text-sm font-medium text-slate-900">
                {isUploading
                  ? 'Extracting rows and storing document in SRC_CCE...'
                  : 'Drop Excel, CSV, Text, or Image files here to generate reports'}
              </p>
              <div className="mt-1.5 flex flex-wrap items-center gap-x-3 gap-y-1 text-xs text-slate-500">
                <span className="inline-flex items-center gap-1">
                  <FileSpreadsheet className="w-3.5 h-3.5 text-emerald-700" />
                  Excel (.xlsx, .xls) &amp; CSV
                </span>
                <span aria-hidden="true">·</span>
                <span className="inline-flex items-center gap-1">
                  <FileText className="w-3.5 h-3.5 text-slate-600" />
                  Text (.txt, .tsv)
                </span>
                <span aria-hidden="true">·</span>
                <span className="inline-flex items-center gap-1">
                  <ImageIcon className="w-3.5 h-3.5 text-blue-700" />
                  Images (.png, .jpg, .webp)
                </span>
              </div>
            </div>
          </div>

          <button
            type="button"
            disabled={isUploading}
            onClick={() => fileInputRef.current?.click()}
            className="px-4 py-2 text-xs font-medium text-white bg-blue-700 hover:bg-blue-800 disabled:opacity-50 rounded-md transition-colors whitespace-nowrap shrink-0 cursor-pointer"
          >
            {isUploading ? 'Processing...' : 'Select Files'}
          </button>
        </div>
      </div>

      {showSheetBox && (
        <div className="mt-4 pt-4 border-t border-slate-200 space-y-2.5">
          <div className="flex flex-col sm:flex-row gap-3">
            <input
              type="url"
              value={sheetUrl}
              onChange={(e) => setSheetUrl(e.target.value)}
              placeholder="Paste Google Sheets link (https://docs.google.com/spreadsheets/d/.../edit#gid=0)"
              className="flex-1 px-3 py-2 text-sm border border-slate-300 rounded-md focus:outline-none focus:ring-2 focus:ring-blue-600"
            />
            <input
              type="text"
              value={sheetLabel}
              onChange={(e) => setSheetLabel(e.target.value)}
              placeholder="Sheet label (optional)"
              className="sm:w-52 px-3 py-2 text-sm border border-slate-300 rounded-md focus:outline-none focus:ring-2 focus:ring-blue-600"
            />
            <button
              type="button"
              disabled={isUploading || !sheetUrl.trim()}
              onClick={handleSheetSubmit}
              className="px-4 py-2 text-xs font-medium text-white bg-blue-700 hover:bg-blue-800 disabled:opacity-50 rounded-md transition-colors whitespace-nowrap shrink-0 cursor-pointer"
            >
              Import Google Sheet
            </button>
          </div>
          <p className="text-xs text-slate-500">
            Supports shared Google Sheets links (&quot;Anyone with the link can view&quot;). For private sheets tied directly to your Gmail account, review and approve the Google Sheets access card when prompted.
          </p>
        </div>
      )}

      {showPasteBox && (
        <div className="mt-4 pt-4 border-t border-slate-200 space-y-3">
          <div className="flex flex-col sm:flex-row gap-3">
            <input
              type="text"
              value={pasteTitle}
              onChange={(e) => setPasteTitle(e.target.value)}
              placeholder="Document name (e.g. Today_Call_Notes.txt)"
              className="flex-1 px-3 py-2 text-sm border border-slate-300 rounded-md focus:outline-none focus:ring-2 focus:ring-blue-600"
            />
            <button
              type="button"
              disabled={isUploading || !pasteContent.trim()}
              onClick={handlePasteSubmit}
              className="px-4 py-2 text-xs font-medium text-white bg-blue-700 hover:bg-blue-800 disabled:opacity-50 rounded-md transition-colors whitespace-nowrap shrink-0 cursor-pointer"
            >
              Extract &amp; Save to SRC_CCE
            </button>
          </div>
          <textarea
            rows={4}
            value={pasteContent}
            onChange={(e) => setPasteContent(e.target.value)}
            placeholder="Paste tab-separated rows, CSV lines, or plain text CRM notes (e.g. Customer Raj Kumar 9835011223 Bolero Neo DSE Pramod Chandrahiya Followup today called by Chandrashekhar Kumar - requested EMI quote)..."
            className="w-full px-3 py-2 text-sm border border-slate-300 rounded-md focus:outline-none focus:ring-2 focus:ring-blue-600 font-mono"
          />
        </div>
      )}

      {documents.length > 0 && (
        <div className="mt-5 pt-4 border-t border-slate-200 flex flex-wrap items-center justify-between gap-3">
          <div className="flex flex-wrap items-center gap-2 text-xs text-slate-600">
            <span className="font-medium text-slate-800">Active Source Files:</span>
            {documents.map((doc) => {
              const active = selectedDocIds.includes(doc.id);
              return (
                <button
                  key={doc.id}
                  type="button"
                  onClick={() => onToggleDocSelection(doc.id)}
                  className={`px-2.5 py-1 text-xs rounded border transition-colors whitespace-nowrap cursor-pointer ${
                    active
                      ? 'bg-blue-50 border-blue-300 text-blue-800 font-medium'
                      : 'bg-white border-slate-200 text-slate-500 hover:text-slate-800'
                  }`}
                >
                  {doc.fileName} ({doc.rowCount})
                </button>
              );
            })}
          </div>
          <button
            type="button"
            onClick={onSelectAllDocs}
            className="text-xs font-medium text-blue-700 hover:underline whitespace-nowrap shrink-0 cursor-pointer"
          >
            {selectedDocIds.length === documents.length ? 'All Selected' : 'Select All'}
          </button>
        </div>
      )}
    </section>
  );
}
