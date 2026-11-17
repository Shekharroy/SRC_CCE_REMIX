import { useEffect, useMemo, useState } from 'react';
import { Download, Search, X } from 'lucide-react';
import { SortableHeader, sortItems, toggleSortState, type SortDirection, type SortState } from './SortableHeader';
import { StatusText } from './StatusText';
import type { EnquiryRow } from '../shared/types';

export interface DrillDownContext {
  title: string;
  subtitle: string;
  rows: EnquiryRow[];
  initialFocusedRowId?: string;
}

interface DrillDownDrawerProps {
  context: DrillDownContext | null;
  onClose: () => void;
  onExportSlice: (title: string, rows: EnquiryRow[]) => Promise<void>;
}

type DrillSortKey =
  | 'enquiryNo'
  | 'date'
  | 'customerName'
  | 'mobile'
  | 'counterName'
  | 'dseName'
  | 'principleModel'
  | 'status'
  | 'crmName';

export function DrillDownDrawer({
  context,
  onClose,
  onExportSlice,
}: DrillDownDrawerProps) {
  const [statusFilter, setStatusFilter] = useState<string>('ALL');
  const [modelFilter, setModelFilter] = useState<string>('ALL');
  const [searchQuery, setSearchQuery] = useState<string>('');
  const [sort, setSort] = useState<SortState<DrillSortKey> | null>(null);
  const [focusedRowId, setFocusedRowId] = useState<string | null>(null);
  const [isExporting, setIsExporting] = useState(false);

  useEffect(() => {
    setStatusFilter('ALL');
    setModelFilter('ALL');
    setSearchQuery('');
    setSort(null);
    setFocusedRowId(context?.initialFocusedRowId || context?.rows[0]?.id || null);
  }, [context]);

  useEffect(() => {
    if (!context) return;
    const onKeyDown = (e: KeyboardEvent) => {
      if (e.key === 'Escape') onClose();
    };
    window.addEventListener('keydown', onKeyDown);
    return () => window.removeEventListener('keydown', onKeyDown);
  }, [context, onClose]);

  const availableStatuses = useMemo(() => {
    if (!context) return [];
    return Array.from(new Set(context.rows.map((r) => r.status))).sort();
  }, [context]);

  const availableModels = useMemo(() => {
    if (!context) return [];
    return Array.from(new Set(context.rows.map((r) => r.principleModel || 'Unspecified Model'))).sort();
  }, [context]);

  const filteredRows = useMemo(() => {
    if (!context) return [];
    const q = searchQuery.trim().toLowerCase();
    const list = context.rows.filter((r) => {
      if (statusFilter !== 'ALL' && r.status !== statusFilter) return false;
      if (modelFilter !== 'ALL' && (r.principleModel || 'Unspecified Model') !== modelFilter) {
        return false;
      }
      if (q) {
        const hay = `${r.enquiryNo} ${r.customerName} ${r.mobile} ${r.dseName} ${r.counterName} ${r.principleModel} ${r.crmName} ${r.remarks}`.toLowerCase();
        if (!hay.includes(q)) return false;
      }
      return true;
    });

    return sortItems(list, sort, {
      enquiryNo: (r) => r.enquiryNo,
      date: (r) => r.followupDate || r.calledDate || r.enquiryDate || '',
      customerName: (r) => r.customerName,
      mobile: (r) => r.mobile,
      counterName: (r) => r.counterName,
      dseName: (r) => r.dseName,
      principleModel: (r) => r.principleModel,
      status: (r) => r.status,
      crmName: (r) => r.crmName,
    });
  }, [context, statusFilter, modelFilter, searchQuery, sort]);

  const metrics = useMemo(() => {
    const total = filteredRows.length;
    const followup = filteredRows.filter((r) => r.status === 'Follow-up').length;
    const hot = filteredRows.filter((r) => r.status === 'Hot').length;
    const converted = filteredRows.filter((r) => r.isConverted).length;
    const convRate = total > 0 ? Number(((converted / total) * 100).toFixed(1)) : 0;
    return { total, followup, hot, converted, convRate };
  }, [filteredRows]);

  const focusedRow = useMemo(() => {
    if (!focusedRowId) return null;
    return (
      filteredRows.find((r) => r.id === focusedRowId) ||
      context?.rows.find((r) => r.id === focusedRowId) ||
      null
    );
  }, [focusedRowId, filteredRows, context]);

  if (!context) return null;

  const handleSort = (key: DrillSortKey, defaultDir: SortDirection = 'asc') => {
    setSort((prev) => toggleSortState(prev, key, defaultDir));
  };

  const handleExport = async () => {
    if (filteredRows.length === 0) return;
    setIsExporting(true);
    try {
      await onExportSlice(context.title, filteredRows);
    } finally {
      setIsExporting(false);
    }
  };

  return (
    <div
      className="fixed inset-0 z-40 flex justify-end bg-slate-900/35"
      role="dialog"
      aria-modal="true"
      aria-label={`Drill-down analysis: ${context.title}`}
      onClick={onClose}
    >
      <div
        className="w-full max-w-5xl bg-white h-full border-l border-slate-200 flex flex-col overflow-hidden"
        onClick={(e) => e.stopPropagation()}
      >
        {/* Drawer Header */}
        <div className="px-6 py-4 border-b border-slate-200 flex flex-wrap items-center justify-between gap-4 bg-white">
          <div className="min-w-0">
            <p className="text-xs font-medium text-blue-700">Data Point Drill-Down Analysis</p>
            <h2 className="text-base font-semibold text-slate-900 truncate mt-0.5">
              {context.title}
            </h2>
            <p className="text-xs text-slate-500 mt-0.5">{context.subtitle}</p>
          </div>

          <div className="flex items-center gap-2.5 shrink-0">
            <button
              type="button"
              disabled={filteredRows.length === 0 || isExporting}
              onClick={handleExport}
              className="inline-flex items-center gap-1.5 px-3.5 py-2 text-xs font-medium text-white bg-blue-700 hover:bg-blue-800 disabled:opacity-40 rounded-md transition-colors whitespace-nowrap cursor-pointer"
            >
              <Download className="w-3.5 h-3.5" />
              <span>{isExporting ? 'Exporting...' : `Export Slice (${filteredRows.length}) (.xlsx)`}</span>
            </button>
            <button
              type="button"
              onClick={onClose}
              aria-label="Close drill-down drawer"
              className="p-2 text-slate-500 hover:text-slate-900 border border-slate-200 rounded-md hover:bg-slate-50 transition-colors cursor-pointer"
            >
              <X className="w-4 h-4" />
            </button>
          </div>
        </div>

        {/* Summary KPI Strip */}
        <div className="grid grid-cols-2 sm:grid-cols-5 border-b border-slate-200 divide-x divide-slate-200 bg-slate-50/70">
          <div className="px-4 py-3">
            <p className="text-xs text-slate-500">Matched Rows</p>
            <p className="text-lg font-semibold font-mono tabular-nums text-slate-900 mt-0.5">
              {metrics.total}
            </p>
          </div>
          <div className="px-4 py-3">
            <p className="text-xs text-slate-500">Follow-up</p>
            <p className="text-lg font-semibold font-mono tabular-nums text-blue-700 mt-0.5">
              {metrics.followup}
            </p>
          </div>
          <div className="px-4 py-3">
            <p className="text-xs text-slate-500">Hot Leads</p>
            <p className="text-lg font-semibold font-mono tabular-nums text-amber-700 mt-0.5">
              {metrics.hot}
            </p>
          </div>
          <div className="px-4 py-3">
            <p className="text-xs text-slate-500">Converted / Booked</p>
            <p className="text-lg font-semibold font-mono tabular-nums text-emerald-700 mt-0.5">
              {metrics.converted}
            </p>
          </div>
          <div className="px-4 py-3">
            <p className="text-xs text-slate-500">Conversion Rate</p>
            <p className="text-lg font-semibold font-mono tabular-nums text-slate-900 mt-0.5">
              {metrics.convRate}%
            </p>
          </div>
        </div>

        {/* Sub-filters inside drawer */}
        <div className="px-6 py-3 border-b border-slate-200 flex flex-wrap items-center gap-3 bg-white">
          <div className="relative flex-1 min-w-[200px]">
            <Search className="w-3.5 h-3.5 text-slate-400 absolute left-3 top-1/2 -translate-y-1/2" />
            <input
              type="text"
              value={searchQuery}
              onChange={(e) => setSearchQuery(e.target.value)}
              placeholder="Filter inside drill-down (customer, mobile, remarks)..."
              className="w-full pl-8 pr-3 py-1.5 text-xs border border-slate-300 rounded-md focus:outline-none focus:ring-2 focus:ring-blue-600"
            />
          </div>

          <select
            value={modelFilter}
            onChange={(e) => setModelFilter(e.target.value)}
            className="px-2.5 py-1.5 text-xs border border-slate-300 rounded-md bg-white focus:outline-none focus:ring-2 focus:ring-blue-600"
          >
            <option value="ALL">All Models ({availableModels.length})</option>
            {availableModels.map((m) => (
              <option key={m} value={m}>
                {m}
              </option>
            ))}
          </select>

          <select
            value={statusFilter}
            onChange={(e) => setStatusFilter(e.target.value)}
            className="px-2.5 py-1.5 text-xs border border-slate-300 rounded-md bg-white focus:outline-none focus:ring-2 focus:ring-blue-600"
          >
            <option value="ALL">All Statuses ({availableStatuses.length})</option>
            {availableStatuses.map((s) => (
              <option key={s} value={s}>
                {s}
              </option>
            ))}
          </select>
        </div>

        {/* Scrollable Content: Record Table + Focused Enquiry Inspector */}
        <div className="flex-1 overflow-y-auto p-6 space-y-5">
          {focusedRow && (
            <div className="border border-slate-200 rounded-md bg-slate-50/70 p-4 space-y-3">
              <div className="flex flex-wrap items-center justify-between gap-2 border-b border-slate-200 pb-2.5">
                <div>
                  <p className="text-xs font-medium text-slate-500">
                    Selected Record Detail · <span className="font-mono">{focusedRow.enquiryNo}</span>
                  </p>
                  <h3 className="text-sm font-semibold text-slate-900 mt-0.5">
                    {focusedRow.customerName} {focusedRow.mobile ? `· ${focusedRow.mobile}` : ''}
                  </h3>
                </div>
                <StatusText status={focusedRow.status} />
              </div>

              <div className="grid grid-cols-2 sm:grid-cols-4 gap-3 text-xs">
                <div>
                  <p className="text-slate-500">Counter / DSE</p>
                  <p className="font-medium text-slate-900 mt-0.5">
                    {focusedRow.counterName} · {focusedRow.dseName}
                  </p>
                </div>
                <div>
                  <p className="text-slate-500">Principle Model</p>
                  <p className="font-medium text-slate-900 mt-0.5">
                    {focusedRow.principleModel}
                  </p>
                </div>
                <div>
                  <p className="text-slate-500">CRM / Caller</p>
                  <p className="font-medium text-slate-900 mt-0.5">
                    {focusedRow.crmName || 'Unassigned'}
                  </p>
                </div>
                <div>
                  <p className="text-slate-500">Key Dates (Enq / Follow / Call)</p>
                  <p className="font-mono tabular-nums text-slate-900 mt-0.5">
                    {focusedRow.enquiryDate || '-'} / {focusedRow.followupDate || '-'} / {focusedRow.calledDate || '-'}
                  </p>
                </div>
              </div>

              <div className="text-xs pt-1">
                <span className="text-slate-500">Remarks: </span>
                <span className="text-slate-800 leading-relaxed">
                  {focusedRow.remarks || 'No remarks recorded.'}
                </span>
                {focusedRow.sourceFileName && (
                  <span className="text-slate-400 ml-2">
                    (Source: {focusedRow.sourceFileName})
                  </span>
                )}
              </div>
            </div>
          )}

          <div className="overflow-x-auto border border-slate-200 rounded-md">
            <table className="w-full text-left border-collapse text-xs">
              <thead>
                <tr className="bg-slate-50 border-b border-slate-200 text-slate-600 font-medium">
                  <SortableHeader label="Enquiry No" sortKey="enquiryNo" sort={sort} onSort={handleSort} />
                  <SortableHeader label="Date" sortKey="date" sort={sort} onSort={handleSort} />
                  <SortableHeader label="Customer" sortKey="customerName" sort={sort} onSort={handleSort} />
                  <SortableHeader label="Mobile" sortKey="mobile" sort={sort} onSort={handleSort} />
                  <SortableHeader label="Counter" sortKey="counterName" sort={sort} onSort={handleSort} />
                  <SortableHeader label="DSE Name" sortKey="dseName" sort={sort} onSort={handleSort} />
                  <SortableHeader label="Principle Model" sortKey="principleModel" sort={sort} onSort={handleSort} />
                  <SortableHeader label="Status" sortKey="status" sort={sort} onSort={handleSort} />
                  <SortableHeader label="CRM / Caller" sortKey="crmName" sort={sort} onSort={handleSort} />
                </tr>
              </thead>
              <tbody className="divide-y divide-slate-200">
                {filteredRows.length === 0 ? (
                  <tr>
                    <td colSpan={9} className="py-8 text-center text-slate-500">
                      No enquiry rows match the current drill-down filters.
                    </td>
                  </tr>
                ) : (
                  filteredRows.map((row) => {
                    const isFocused = row.id === focusedRow?.id;
                    return (
                      <tr
                        key={row.id}
                        onClick={() => setFocusedRowId(row.id)}
                        className={`cursor-pointer transition-colors ${
                          isFocused ? 'bg-blue-50/70' : 'hover:bg-slate-50/80'
                        }`}
                      >
                        <td className="py-2 px-3 font-mono tabular-nums text-slate-600 whitespace-nowrap">
                          {row.enquiryNo}
                        </td>
                        <td className="py-2 px-3 font-mono tabular-nums text-slate-600 whitespace-nowrap">
                          {row.followupDate || row.calledDate || row.enquiryDate || '-'}
                        </td>
                        <td className="py-2 px-3 font-medium text-slate-900 whitespace-nowrap">
                          {row.customerName}
                        </td>
                        <td className="py-2 px-3 font-mono tabular-nums text-slate-600 whitespace-nowrap">
                          {row.mobile || '-'}
                        </td>
                        <td className="py-2 px-3 text-slate-600 whitespace-nowrap">
                          {row.counterName}
                        </td>
                        <td className="py-2 px-3 text-slate-800 whitespace-nowrap">
                          {row.dseName}
                        </td>
                        <td className="py-2 px-3 text-slate-800 whitespace-nowrap">
                          {row.principleModel}
                        </td>
                        <td className="py-2 px-3">
                          <StatusText status={row.status} />
                        </td>
                        <td className="py-2 px-3 text-slate-700 whitespace-nowrap">
                          {row.crmName || '-'}
                        </td>
                      </tr>
                    );
                  })
                )}
              </tbody>
            </table>
          </div>
        </div>
      </div>
    </div>
  );
}
