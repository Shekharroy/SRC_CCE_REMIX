import { Fragment, useEffect, useMemo, useState } from 'react';
import {
  ChevronDown,
  ChevronRight,
  Download,
  Save,
  Search,
  Check,
  FileSpreadsheet,
  RotateCcw,
} from 'lucide-react';
import {
  ALL_NORMALIZED_STATUSES,
  buildAllReports,
  COUNTER_DEFINITIONS,
  filterEnquiryRows,
} from '../shared/reportEngine';
import type {
  AllReportsBundle,
  CanonicalCounterName,
  CounterDseBreakdown,
  CounterSummary,
  CrmDailyCallSummary,
  CustomReportConfig,
  DseFollowupGroup,
  EnquiryRow,
  NormalizedStatus,
  PrincipleModelSummary,
} from '../shared/types';
import { CustomReportBuilder } from './CustomReportBuilder';
import { DrillDownDrawer, type DrillDownContext } from './DrillDownDrawer';
import { SortableHeader, sortItems, toggleSortState, type SortState } from './SortableHeader';
import { StatusText } from './StatusText';

export type ActiveReportTab = 'report1' | 'report2' | 'report3' | 'report4' | 'custom';

interface ReportsSectionProps {
  bundle: AllReportsBundle;
  rawRows: EnquiryRow[];
  targetDate: string;
  onChangeTargetDate: (date: string) => void;
  strictDateMatch: boolean;
  onToggleStrictDate: (strict: boolean) => void;
  crmTarget: string;
  onChangeCrmTarget: (crm: string) => void;
  onExportExcel: (
    scope: 'all' | 'report1' | 'report2' | 'report3' | 'report4',
    overrideRows?: EnquiryRow[]
  ) => Promise<void>;
  onExportCustomExcel: (config: CustomReportConfig) => Promise<void>;
  onSaveSnapshot: (title: string) => Promise<void>;
  isSavingSnapshot: boolean;
  snapshotSavedFeedback: boolean;
  onOpenUpload: () => void;
  onActiveFilteredRowsChange?: (rows: EnquiryRow[] | null) => void;
}

type R1GroupSortKey = 'dseName' | 'counterName' | 'totalFollowups' | 'hotCount' | 'convertedCount';
type R1RowSortKey =
  | 'dseName'
  | 'counterName'
  | 'followupDate'
  | 'customerName'
  | 'mobile'
  | 'principleModel'
  | 'status'
  | 'remarks';
type R2SortKey =
  | 'principleModel'
  | 'totalEnquiries'
  | 'openCount'
  | 'followupCount'
  | 'hotCount'
  | 'warmCount'
  | 'coldCount'
  | 'bookedCount'
  | 'convertedCount'
  | 'lostCount'
  | 'conversionRate'
  | 'topCounter';
type R3DailySortKey =
  | 'date'
  | 'totalCalled'
  | 'followupCount'
  | 'hotCount'
  | 'convertedCount'
  | 'lostCount'
  | 'remarks';
type R3CallSortKey =
  | 'calledDate'
  | 'customerName'
  | 'mobile'
  | 'counterName'
  | 'dseName'
  | 'principleModel'
  | 'status'
  | 'remarks';
type R3ConvSortKey =
  | 'conversionDate'
  | 'enquiryNo'
  | 'customerName'
  | 'mobile'
  | 'counterName'
  | 'dseName'
  | 'principleModel'
  | 'status'
  | 'remarks';
type R4SortKey =
  | 'counterName'
  | 'totalEnquiries'
  | 'openCount'
  | 'followupCount'
  | 'hotCount'
  | 'warmCount'
  | 'coldCount'
  | 'bookedCount'
  | 'convertedCount'
  | 'lostCount'
  | 'conversionRate';

export function ReportsSection({
  bundle,
  rawRows,
  targetDate,
  onChangeTargetDate,
  strictDateMatch,
  onToggleStrictDate,
  crmTarget,
  onChangeCrmTarget,
  onExportExcel,
  onExportCustomExcel,
  onSaveSnapshot,
  isSavingSnapshot,
  snapshotSavedFeedback,
  onOpenUpload,
  onActiveFilteredRowsChange,
}: ReportsSectionProps) {
  const [activeTab, setActiveTab] = useState<ActiveReportTab>('report1');
  const [searchQuery, setSearchQuery] = useState('');
  const [counterFilter, setCounterFilter] = useState<string>('ALL');
  const [dseFilter, setDseFilter] = useState<string>('ALL');
  const [modelFilter, setModelFilter] = useState<string>('ALL');
  const [statusFilter, setStatusFilter] = useState<string>('ALL');

  const [expandedCounters, setExpandedCounters] = useState<Record<string, boolean>>({
    Chandrahiya: true,
  });
  const [snapshotTitle, setSnapshotTitle] = useState('');
  const [drillDown, setDrillDown] = useState<DrillDownContext | null>(null);

  // Column sort states per table
  const [r1GroupSort, setR1GroupSort] = useState<SortState<R1GroupSortKey> | null>(null);
  const [r1RowSort, setR1RowSort] = useState<SortState<R1RowSortKey> | null>(null);
  const [r2Sort, setR2Sort] = useState<SortState<R2SortKey> | null>(null);
  const [r3DailySort, setR3DailySort] = useState<SortState<R3DailySortKey> | null>(null);
  const [r3CallSort, setR3CallSort] = useState<SortState<R3CallSortKey> | null>(null);
  const [r3ConvSort, setR3ConvSort] = useState<SortState<R3ConvSortKey> | null>(null);
  const [r4Sort, setR4Sort] = useState<SortState<R4SortKey> | null>(null);

  const hasActiveFilters = Boolean(
    counterFilter !== 'ALL' ||
      dseFilter !== 'ALL' ||
      modelFilter !== 'ALL' ||
      statusFilter !== 'ALL' ||
      searchQuery.trim()
  );

  // Available options for multi-facet filter dropdowns
  const availableCounters = useMemo(() => {
    const base = COUNTER_DEFINITIONS.map((c) => c.name) as string[];
    if (rawRows.some((r) => r.counterName === 'Unassigned')) {
      base.push('Unassigned');
    }
    return base;
  }, [rawRows]);

  const availableDses = useMemo(() => {
    const set = new Set<string>();
    for (const r of rawRows) {
      if (counterFilter !== 'ALL' && r.counterName !== counterFilter) continue;
      if (r.dseName) set.add(r.dseName);
    }
    return Array.from(set).sort((a, b) => a.localeCompare(b));
  }, [rawRows, counterFilter]);

  const availableModels = useMemo(() => {
    const set = new Set<string>();
    for (const r of rawRows) {
      set.add((r.principleModel || '').trim() || 'Unspecified Model');
    }
    return Array.from(set).sort((a, b) => a.localeCompare(b));
  }, [rawRows]);

  const filteredRawRows = useMemo(() => {
    if (!hasActiveFilters) return rawRows;
    return filterEnquiryRows(rawRows, {
      counter: counterFilter,
      dse: dseFilter,
      principleModel: modelFilter,
      status: statusFilter,
      searchQuery,
    });
  }, [rawRows, hasActiveFilters, counterFilter, dseFilter, modelFilter, statusFilter, searchQuery]);

  useEffect(() => {
    onActiveFilteredRowsChange?.(hasActiveFilters ? filteredRawRows : null);
  }, [hasActiveFilters, filteredRawRows, onActiveFilteredRowsChange]);

  // Keep bundle responsive to global filters
  const displayBundle = useMemo(() => {
    if (!hasActiveFilters) return bundle;
    return buildAllReports(filteredRawRows, {
      targetDate,
      crmTarget,
      strictDateMatch,
    });
  }, [hasActiveFilters, bundle, filteredRawRows, targetDate, crmTarget, strictDateMatch]);

  const resetAllFilters = () => {
    setCounterFilter('ALL');
    setDseFilter('ALL');
    setModelFilter('ALL');
    setStatusFilter('ALL');
    setSearchQuery('');
  };

  const toggleCounterExpand = (name: CanonicalCounterName) => {
    setExpandedCounters((prev) => ({ ...prev, [name]: !prev[name] }));
  };

  // Sorted collections for Reports 1 - 4
  const sortedR1Groups = useMemo(() => {
    return sortItems(displayBundle.report1DseFollowup.groups, r1GroupSort, {
      dseName: (g) => g.dseName,
      counterName: (g) => g.counterName,
      totalFollowups: (g) => g.totalFollowups,
      hotCount: (g) => g.hotCount,
      convertedCount: (g) => g.convertedCount,
    });
  }, [displayBundle.report1DseFollowup.groups, r1GroupSort]);

  const sortedR1Rows = useMemo(() => {
    return sortItems(displayBundle.report1DseFollowup.flatRows, r1RowSort, {
      dseName: (r) => r.dseName,
      counterName: (r) => r.counterName,
      followupDate: (r) => r.followupDate || targetDate,
      customerName: (r) => r.customerName,
      mobile: (r) => r.mobile,
      principleModel: (r) => r.principleModel,
      status: (r) => r.status,
      remarks: (r) => r.remarks,
    });
  }, [displayBundle.report1DseFollowup.flatRows, r1RowSort, targetDate]);

  const sortedR2Summaries = useMemo(() => {
    return sortItems(displayBundle.report2PrincipleModel.summaries, r2Sort, {
      principleModel: (m) => m.principleModel,
      totalEnquiries: (m) => m.totalEnquiries,
      openCount: (m) => m.openCount,
      followupCount: (m) => m.followupCount,
      hotCount: (m) => m.hotCount,
      warmCount: (m) => m.warmCount,
      coldCount: (m) => m.coldCount,
      bookedCount: (m) => m.bookedCount,
      convertedCount: (m) => m.convertedCount,
      lostCount: (m) => m.lostCount,
      conversionRate: (m) => m.conversionRate,
      topCounter: (m) => m.topCounter,
    });
  }, [displayBundle.report2PrincipleModel.summaries, r2Sort]);

  const sortedR3Daily = useMemo(() => {
    return sortItems(displayBundle.report3CrmDailyCall.dailySummaries, r3DailySort, {
      date: (d) => d.date,
      totalCalled: (d) => d.totalCalled,
      followupCount: (d) => d.followupCount,
      hotCount: (d) => d.hotCount,
      convertedCount: (d) => d.convertedCount,
      lostCount: (d) => d.lostCount,
      remarks: (d) => d.remarksList.join(' | '),
    });
  }, [displayBundle.report3CrmDailyCall.dailySummaries, r3DailySort]);

  const sortedR3CallRows = useMemo(() => {
    return sortItems(displayBundle.report3CrmDailyCall.callRows, r3CallSort, {
      calledDate: (r) => r.calledDate || r.followupDate || r.enquiryDate || '',
      customerName: (r) => r.customerName,
      mobile: (r) => r.mobile,
      counterName: (r) => r.counterName,
      dseName: (r) => r.dseName,
      principleModel: (r) => r.principleModel,
      status: (r) => r.status,
      remarks: (r) => r.remarks,
    });
  }, [displayBundle.report3CrmDailyCall.callRows, r3CallSort]);

  const sortedR3Conversions = useMemo(() => {
    return sortItems(displayBundle.report3CrmDailyCall.conversionList, r3ConvSort, {
      conversionDate: (r) => r.conversionDate || r.calledDate || r.enquiryDate || '',
      enquiryNo: (r) => r.enquiryNo,
      customerName: (r) => r.customerName,
      mobile: (r) => r.mobile,
      counterName: (r) => r.counterName,
      dseName: (r) => r.dseName,
      principleModel: (r) => r.principleModel,
      status: (r) => r.status,
      remarks: (r) => r.remarks,
    });
  }, [displayBundle.report3CrmDailyCall.conversionList, r3ConvSort]);

  const sortedR4Counters = useMemo(() => {
    return sortItems(displayBundle.report4CounterWise.summaries, r4Sort, {
      counterName: (c) => c.counterName,
      totalEnquiries: (c) => c.totalEnquiries,
      openCount: (c) => c.openCount,
      followupCount: (c) => c.followupCount,
      hotCount: (c) => c.hotCount,
      warmCount: (c) => c.warmCount,
      coldCount: (c) => c.coldCount,
      bookedCount: (c) => c.bookedCount,
      convertedCount: (c) => c.convertedCount,
      lostCount: (c) => c.lostCount,
      conversionRate: (c) => c.conversionRate,
    });
  }, [displayBundle.report4CounterWise.summaries, r4Sort]);

  const maxModelCount = useMemo(() => {
    return Math.max(1, ...displayBundle.report2PrincipleModel.summaries.map((s) => s.totalEnquiries));
  }, [displayBundle.report2PrincipleModel.summaries]);

  // Drill-down helpers
  const drillIntoDseGroup = (g: DseFollowupGroup, filterStatus?: 'Hot' | 'CONVERTED_OR_BOOKED') => {
    const matched = g.rows.filter((r) => {
      if (filterStatus === 'Hot') return r.status === 'Hot';
      if (filterStatus === 'CONVERTED_OR_BOOKED') return r.isConverted;
      return true;
    });
    setDrillDown({
      title: `${g.dseName} (${g.counterName})${
        filterStatus === 'Hot'
          ? ' · Hot Leads'
          : filterStatus === 'CONVERTED_OR_BOOKED'
          ? ' · Converted / Booked'
          : ''
      }`,
      subtitle: `First Report Follow-up Slice · ${matched.length} matching enquiries`,
      rows: matched,
    });
  };

  const drillIntoModel = (
    m: PrincipleModelSummary,
    status?: NormalizedStatus | 'CONVERTED_OR_BOOKED'
  ) => {
    const matched = displayBundle.report2PrincipleModel.rows.filter((r) => {
      if (((r.principleModel || '').trim() || 'Unspecified Model') !== m.principleModel) {
        return false;
      }
      if (status === 'CONVERTED_OR_BOOKED') return r.isConverted;
      if (status) return r.status === status;
      return true;
    });
    setDrillDown({
      title: `${m.principleModel}${
        status ? ` · ${status === 'CONVERTED_OR_BOOKED' ? 'Converted / Booked' : status}` : ''
      }`,
      subtitle: `Principle Model Drill-Down · Top Counter: ${m.topCounter} · ${matched.length} rows`,
      rows: matched,
    });
  };

  const drillIntoCrmDay = (
    d: CrmDailyCallSummary,
    status?: NormalizedStatus | 'CONVERTED_OR_BOOKED'
  ) => {
    const matched = displayBundle.report3CrmDailyCall.callRows.filter((r) => {
      const cDate = r.calledDate || r.followupDate || r.enquiryDate || 'Unspecified Date';
      if (cDate !== d.date) return false;
      if (status === 'CONVERTED_OR_BOOKED') return r.isConverted;
      if (status) return r.status === status;
      return true;
    });
    setDrillDown({
      title: `CRM ${displayBundle.report3CrmDailyCall.crmName} · ${d.date}${
        status ? ` · ${status === 'CONVERTED_OR_BOOKED' ? 'Converted' : status}` : ''
      }`,
      subtitle: `Daily Called Data Slice · ${matched.length} call entries`,
      rows: matched,
    });
  };

  const drillIntoCounterMetric = (
    counter: CounterSummary,
    dse?: CounterDseBreakdown,
    status?: NormalizedStatus | 'CONVERTED_OR_BOOKED'
  ) => {
    const matched = filteredRawRows.filter((r) => {
      if (r.counterName !== counter.counterName) return false;
      if (dse && r.dseName !== dse.dseName) return false;
      if (status === 'CONVERTED_OR_BOOKED') return r.isConverted;
      if (status) return r.status === status;
      return true;
    });
    const statusLabel = status
      ? status === 'CONVERTED_OR_BOOKED'
        ? 'Converted / Booked'
        : status
      : 'All Statuses';
    setDrillDown({
      title: `${counter.counterName}${dse ? ` · ${dse.dseName}` : ''} · ${statusLabel}`,
      subtitle: `Counter Wise Drill-Down · ${matched.length} matching enquiries`,
      rows: matched,
    });
  };

  if (bundle.totalRecords === 0) {
    return (
      <section
        id="reports"
        className="bg-white border border-slate-200 rounded-lg p-12 text-center"
      >
        <div className="mx-auto w-12 h-12 rounded-md bg-slate-100 border border-slate-200 flex items-center justify-center text-slate-500">
          <FileSpreadsheet className="w-6 h-6" />
        </div>
        <h2 className="mt-4 text-base font-semibold text-slate-900">
          No Enquiry Records Loaded Yet
        </h2>
        <p className="mt-1.5 text-sm text-slate-600 max-w-md mx-auto">
          Upload an Excel sheet, CSV file, text log, or register image above—or download the sample Excel sheet to inspect all 4 reports immediately.
        </p>
        <div className="mt-6 flex flex-wrap items-center justify-center gap-3">
          <button
            type="button"
            onClick={onOpenUpload}
            className="px-4 py-2 text-xs font-medium text-white bg-blue-700 hover:bg-blue-800 rounded-md transition-colors whitespace-nowrap cursor-pointer"
          >
            Upload Document
          </button>
          <a
            href="/api/sample-template"
            download
            className="px-4 py-2 text-xs font-medium text-slate-700 bg-slate-100 hover:bg-slate-200 rounded-md transition-colors whitespace-nowrap"
          >
            Download Sample Excel
          </a>
        </div>
      </section>
    );
  }

  const exportRowsArg = hasActiveFilters ? filteredRawRows : undefined;

  return (
    <section id="reports" className="bg-white border border-slate-200 rounded-lg">
      {/* Toolbar: Segmented Tabs + Controls */}
      <div className="p-5 border-b border-slate-200 space-y-4">
        <div className="flex flex-col xl:flex-row xl:items-center justify-between gap-4">
          <div className="flex flex-wrap items-center gap-1 p-1 bg-slate-100 rounded-md">
            <button
              type="button"
              onClick={() => setActiveTab('report1')}
              className={`px-3 py-1.5 text-xs font-medium rounded transition-colors whitespace-nowrap shrink-0 cursor-pointer ${
                activeTab === 'report1'
                  ? 'bg-white text-slate-900 shadow-xs'
                  : 'text-slate-600 hover:text-slate-900'
              }`}
            >
              1. DSE Today Follow-up ({displayBundle.report1DseFollowup.totalFollowups})
            </button>
            <button
              type="button"
              onClick={() => setActiveTab('report2')}
              className={`px-3 py-1.5 text-xs font-medium rounded transition-colors whitespace-nowrap shrink-0 cursor-pointer ${
                activeTab === 'report2'
                  ? 'bg-white text-slate-900 shadow-xs'
                  : 'text-slate-600 hover:text-slate-900'
              }`}
            >
              2. Principle Model Wise ({displayBundle.report2PrincipleModel.totalModels})
            </button>
            <button
              type="button"
              onClick={() => setActiveTab('report3')}
              className={`px-3 py-1.5 text-xs font-medium rounded transition-colors whitespace-nowrap shrink-0 cursor-pointer ${
                activeTab === 'report3'
                  ? 'bg-white text-slate-900 shadow-xs'
                  : 'text-slate-600 hover:text-slate-900'
              }`}
            >
              3. CRM Chandrashekhar ({displayBundle.report3CrmDailyCall.totalCalls})
            </button>
            <button
              type="button"
              onClick={() => setActiveTab('report4')}
              className={`px-3 py-1.5 text-xs font-medium rounded transition-colors whitespace-nowrap shrink-0 cursor-pointer ${
                activeTab === 'report4'
                  ? 'bg-white text-slate-900 shadow-xs'
                  : 'text-slate-600 hover:text-slate-900'
              }`}
            >
              4. Counter Wise Summary (7)
            </button>
            <button
              type="button"
              onClick={() => setActiveTab('custom')}
              className={`px-3 py-1.5 text-xs font-medium rounded transition-colors whitespace-nowrap shrink-0 cursor-pointer ${
                activeTab === 'custom'
                  ? 'bg-white text-slate-900 shadow-xs'
                  : 'text-slate-600 hover:text-slate-900'
              }`}
            >
              5. Custom Report Builder
            </button>
          </div>

          <div className="flex flex-wrap items-center gap-2.5 shrink-0">
            <div className="flex items-center gap-1.5">
              <input
                type="text"
                value={snapshotTitle}
                onChange={(e) => setSnapshotTitle(e.target.value)}
                placeholder="Report title (optional)"
                className="w-40 px-2.5 py-1.5 text-xs border border-slate-300 rounded-md focus:outline-none focus:ring-2 focus:ring-blue-600"
              />
              <button
                type="button"
                disabled={isSavingSnapshot}
                onClick={() => onSaveSnapshot(snapshotTitle)}
                className="inline-flex items-center gap-1.5 px-3 py-1.5 text-xs font-medium text-slate-700 bg-white border border-slate-300 hover:bg-slate-50 rounded-md transition-colors whitespace-nowrap shrink-0 cursor-pointer"
              >
                {snapshotSavedFeedback ? (
                  <>
                    <Check className="w-3.5 h-3.5 text-emerald-700" />
                    <span className="text-emerald-700">Saved to SRC_CCE</span>
                  </>
                ) : (
                  <>
                    <Save className="w-3.5 h-3.5" />
                    <span>{isSavingSnapshot ? 'Saving...' : 'Save Report to DB'}</span>
                  </>
                )}
              </button>
            </div>

            {activeTab !== 'custom' && (
              <button
                type="button"
                disabled={displayBundle.totalRecords === 0}
                onClick={() => onExportExcel(activeTab, exportRowsArg)}
                className="inline-flex items-center gap-1.5 px-3.5 py-1.5 text-xs font-medium text-white bg-blue-700 hover:bg-blue-800 disabled:opacity-40 rounded-md transition-colors whitespace-nowrap shrink-0 cursor-pointer"
              >
                <Download className="w-3.5 h-3.5" />
                <span>
                  {hasActiveFilters
                    ? 'Download Filtered (.xlsx)'
                    : 'Download This Report (.xlsx)'}
                </span>
              </button>
            )}
          </div>
        </div>

        {/* Multi-Dimensional Interactive Filter Bar */}
        {activeTab !== 'custom' && (
          <div className="space-y-3 pt-1">
            <div className="grid grid-cols-1 sm:grid-cols-2 lg:grid-cols-6 gap-2.5">
              <div className="relative lg:col-span-2">
                <Search className="w-4 h-4 text-slate-400 absolute left-3 top-1/2 -translate-y-1/2" />
                <input
                  type="text"
                  value={searchQuery}
                  onChange={(e) => setSearchQuery(e.target.value)}
                  placeholder="Search customer, DSE, model, counter, remarks..."
                  className="w-full pl-9 pr-3 py-1.5 text-xs border border-slate-300 rounded-md focus:outline-none focus:ring-2 focus:ring-blue-600"
                />
              </div>

              <select
                aria-label="Filter by Counter Name"
                value={counterFilter}
                onChange={(e) => {
                  setCounterFilter(e.target.value);
                  setDseFilter('ALL');
                }}
                className="px-2.5 py-1.5 text-xs border border-slate-300 rounded-md bg-white focus:outline-none focus:ring-2 focus:ring-blue-600"
              >
                <option value="ALL">All Counters (7)</option>
                {availableCounters.map((c) => (
                  <option key={c} value={c}>
                    {c}
                  </option>
                ))}
              </select>

              <select
                aria-label="Filter by DSE Name"
                value={dseFilter}
                onChange={(e) => setDseFilter(e.target.value)}
                className="px-2.5 py-1.5 text-xs border border-slate-300 rounded-md bg-white focus:outline-none focus:ring-2 focus:ring-blue-600"
              >
                <option value="ALL">All DSEs ({availableDses.length})</option>
                {availableDses.map((d) => (
                  <option key={d} value={d}>
                    {d}
                  </option>
                ))}
              </select>

              <select
                aria-label="Filter by Principle Model"
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
                aria-label="Filter by Report Status"
                value={statusFilter}
                onChange={(e) => setStatusFilter(e.target.value)}
                className="px-2.5 py-1.5 text-xs border border-slate-300 rounded-md bg-white focus:outline-none focus:ring-2 focus:ring-blue-600"
              >
                <option value="ALL">All Statuses</option>
                {ALL_NORMALIZED_STATUSES.map((st) => (
                  <option key={st} value={st}>
                    {st}
                  </option>
                ))}
                <option value="CONVERTED_OR_BOOKED">Converted / Booked</option>
              </select>
            </div>

            <div className="flex flex-wrap items-center justify-between gap-3 text-xs">
              <div className="flex flex-wrap items-center gap-3 text-slate-600">
                <span>
                  Showing{' '}
                  <strong className="font-mono tabular-nums text-slate-900">
                    {filteredRawRows.length}
                  </strong>{' '}
                  of{' '}
                  <strong className="font-mono tabular-nums text-slate-900">
                    {rawRows.length}
                  </strong>{' '}
                  enquiries
                </span>
                {hasActiveFilters && (
                  <>
                    <span aria-hidden="true">·</span>
                    <button
                      type="button"
                      onClick={resetAllFilters}
                      className="inline-flex items-center gap-1 font-medium text-blue-700 hover:underline cursor-pointer"
                    >
                      <RotateCcw className="w-3 h-3" />
                      <span>Reset Filters</span>
                    </button>
                  </>
                )}
              </div>

              <div className="flex flex-wrap items-center gap-3">
                {activeTab === 'report1' && (
                  <>
                    <label className="inline-flex items-center gap-2 text-slate-600">
                      <span>Target Date:</span>
                      <input
                        type="date"
                        value={targetDate}
                        onChange={(e) => onChangeTargetDate(e.target.value)}
                        className="px-2.5 py-1 text-xs font-mono border border-slate-300 rounded-md focus:outline-none focus:ring-2 focus:ring-blue-600"
                      />
                    </label>
                    <label className="inline-flex items-center gap-1.5 text-slate-600 cursor-pointer">
                      <input
                        type="checkbox"
                        checked={strictDateMatch}
                        onChange={(e) => onToggleStrictDate(e.target.checked)}
                        className="rounded border-slate-300 text-blue-700 focus:ring-blue-600"
                      />
                      <span>Strict date match only</span>
                    </label>
                  </>
                )}

                {activeTab === 'report3' && (
                  <label className="inline-flex items-center gap-2 text-slate-600">
                    <span>CRM Caller:</span>
                    <select
                      value={crmTarget}
                      onChange={(e) => onChangeCrmTarget(e.target.value)}
                      className="px-2.5 py-1 text-xs border border-slate-300 rounded-md bg-white focus:outline-none focus:ring-2 focus:ring-blue-600"
                    >
                      {displayBundle.report3CrmDailyCall.availableCrmNames.map((name) => (
                        <option key={name} value={name}>
                          {name}
                        </option>
                      ))}
                      <option value="ALL">All CRM Callers</option>
                    </select>
                  </label>
                )}
              </div>
            </div>
          </div>
        )}
      </div>

      {/* REPORT 1: DSE Name Wise Today Followup Data List with Remarks */}
      {activeTab === 'report1' && (
        <div className="p-5 space-y-6">
          <div>
            <h2 className="text-base font-semibold text-slate-900">
              First Report: DSE Name Wise Today Followup Data List with Remarks
            </h2>
            <p className="text-xs text-slate-500 mt-0.5">
              Target Date: {displayBundle.report1DseFollowup.targetDate} · {displayBundle.report1DseFollowup.dseCount} DSEs · {displayBundle.report1DseFollowup.totalFollowups} follow-up entries · Click any column header to sort or any row/metric to drill down.
            </p>
          </div>

          {/* DSE Summary Matrix */}
          <div className="overflow-x-auto border border-slate-200 rounded-md">
            <table className="w-full text-left border-collapse text-xs">
              <thead>
                <tr className="bg-slate-50 border-b border-slate-200 text-slate-600 font-medium">
                  <SortableHeader
                    label="DSE Name"
                    sortKey="dseName"
                    sort={r1GroupSort}
                    onSort={(k, d) => setR1GroupSort((p) => toggleSortState(p, k, d))}
                  />
                  <SortableHeader
                    label="Counter"
                    sortKey="counterName"
                    sort={r1GroupSort}
                    onSort={(k, d) => setR1GroupSort((p) => toggleSortState(p, k, d))}
                  />
                  <SortableHeader
                    label="Follow-ups"
                    sortKey="totalFollowups"
                    sort={r1GroupSort}
                    onSort={(k, d) => setR1GroupSort((p) => toggleSortState(p, k, d))}
                    align="right"
                    defaultDirection="desc"
                  />
                  <SortableHeader
                    label="Hot Leads"
                    sortKey="hotCount"
                    sort={r1GroupSort}
                    onSort={(k, d) => setR1GroupSort((p) => toggleSortState(p, k, d))}
                    align="right"
                    defaultDirection="desc"
                  />
                  <SortableHeader
                    label="Converted / Booked"
                    sortKey="convertedCount"
                    sort={r1GroupSort}
                    onSort={(k, d) => setR1GroupSort((p) => toggleSortState(p, k, d))}
                    align="right"
                    defaultDirection="desc"
                  />
                  <th className="py-2.5 px-3 text-right">Actions</th>
                </tr>
              </thead>
              <tbody className="divide-y divide-slate-200">
                {sortedR1Groups.length === 0 ? (
                  <tr>
                    <td colSpan={6} className="py-6 text-center text-slate-500">
                      No follow-up records found for current filters and date ({targetDate}).
                    </td>
                  </tr>
                ) : (
                  sortedR1Groups.map((g) => (
                    <tr
                      key={`${g.dseName}-${g.counterName}`}
                      onClick={() => drillIntoDseGroup(g)}
                      className="hover:bg-slate-50/80 cursor-pointer transition-colors"
                    >
                      <td className="py-2.5 px-3 font-medium text-blue-700 hover:underline">
                        {g.dseName}
                      </td>
                      <td className="py-2.5 px-3 text-slate-600">{g.counterName}</td>
                      <td className="py-2.5 px-3 text-right font-mono tabular-nums font-medium text-slate-900">
                        {g.totalFollowups}
                      </td>
                      <td
                        onClick={(e) => {
                          if (g.hotCount > 0) {
                            e.stopPropagation();
                            drillIntoDseGroup(g, 'Hot');
                          }
                        }}
                        className="py-2.5 px-3 text-right font-mono tabular-nums text-amber-700 hover:underline"
                      >
                        {g.hotCount}
                      </td>
                      <td
                        onClick={(e) => {
                          if (g.convertedCount > 0) {
                            e.stopPropagation();
                            drillIntoDseGroup(g, 'CONVERTED_OR_BOOKED');
                          }
                        }}
                        className="py-2.5 px-3 text-right font-mono tabular-nums text-emerald-700 hover:underline"
                      >
                        {g.convertedCount}
                      </td>
                      <td className="py-2.5 px-3 text-right whitespace-nowrap">
                        <button
                          type="button"
                          onClick={(e) => {
                            e.stopPropagation();
                            setDseFilter((prev) => (prev === g.dseName ? 'ALL' : g.dseName));
                          }}
                          className="text-xs font-medium text-blue-700 hover:underline cursor-pointer mr-3"
                        >
                          {dseFilter === g.dseName ? 'Show All' : 'Filter DSE'}
                        </button>
                        <button
                          type="button"
                          onClick={(e) => {
                            e.stopPropagation();
                            drillIntoDseGroup(g);
                          }}
                          className="text-xs font-medium text-slate-700 hover:text-slate-900 hover:underline cursor-pointer"
                        >
                          Drill Down
                        </button>
                      </td>
                    </tr>
                  ))
                )}
              </tbody>
            </table>
          </div>

          {/* Detailed Follow-up Rows with Remarks */}
          <div className="overflow-x-auto border border-slate-200 rounded-md">
            <table className="w-full text-left border-collapse text-xs">
              <thead>
                <tr className="bg-slate-50 border-b border-slate-200 text-slate-600 font-medium">
                  <th className="py-2.5 px-3 w-10">Sl</th>
                  <SortableHeader
                    label="DSE Name"
                    sortKey="dseName"
                    sort={r1RowSort}
                    onSort={(k, d) => setR1RowSort((p) => toggleSortState(p, k, d))}
                  />
                  <SortableHeader
                    label="Counter"
                    sortKey="counterName"
                    sort={r1RowSort}
                    onSort={(k, d) => setR1RowSort((p) => toggleSortState(p, k, d))}
                  />
                  <SortableHeader
                    label="Followup Date"
                    sortKey="followupDate"
                    sort={r1RowSort}
                    onSort={(k, d) => setR1RowSort((p) => toggleSortState(p, k, d))}
                  />
                  <SortableHeader
                    label="Customer"
                    sortKey="customerName"
                    sort={r1RowSort}
                    onSort={(k, d) => setR1RowSort((p) => toggleSortState(p, k, d))}
                  />
                  <SortableHeader
                    label="Mobile"
                    sortKey="mobile"
                    sort={r1RowSort}
                    onSort={(k, d) => setR1RowSort((p) => toggleSortState(p, k, d))}
                  />
                  <SortableHeader
                    label="Principle Model"
                    sortKey="principleModel"
                    sort={r1RowSort}
                    onSort={(k, d) => setR1RowSort((p) => toggleSortState(p, k, d))}
                  />
                  <SortableHeader
                    label="Status"
                    sortKey="status"
                    sort={r1RowSort}
                    onSort={(k, d) => setR1RowSort((p) => toggleSortState(p, k, d))}
                  />
                  <SortableHeader
                    label="Remarks"
                    sortKey="remarks"
                    sort={r1RowSort}
                    onSort={(k, d) => setR1RowSort((p) => toggleSortState(p, k, d))}
                    className="min-w-[240px]"
                  />
                </tr>
              </thead>
              <tbody className="divide-y divide-slate-200">
                {sortedR1Rows.length === 0 ? (
                  <tr>
                    <td colSpan={9} className="py-6 text-center text-slate-500">
                      No follow-up rows match the active filters.
                    </td>
                  </tr>
                ) : (
                  sortedR1Rows.map((row, idx) => (
                    <tr
                      key={row.id}
                      onClick={() =>
                        setDrillDown({
                          title: `${row.customerName} (${row.enquiryNo})`,
                          subtitle: `DSE: ${row.dseName} · Counter: ${row.counterName} · Follow-up Context`,
                          rows: sortedR1Rows,
                          initialFocusedRowId: row.id,
                        })
                      }
                      className="hover:bg-slate-50/80 cursor-pointer transition-colors"
                    >
                      <td className="py-2.5 px-3 font-mono tabular-nums text-slate-500">
                        {idx + 1}
                      </td>
                      <td className="py-2.5 px-3 font-medium text-slate-900 whitespace-nowrap">
                        {row.dseName}
                      </td>
                      <td className="py-2.5 px-3 text-slate-600 whitespace-nowrap">
                        {row.counterName}
                      </td>
                      <td className="py-2.5 px-3 font-mono tabular-nums text-slate-700 whitespace-nowrap">
                        {row.followupDate || targetDate}
                      </td>
                      <td className="py-2.5 px-3 font-medium text-blue-700 hover:underline whitespace-nowrap">
                        {row.customerName}
                      </td>
                      <td className="py-2.5 px-3 font-mono tabular-nums text-slate-600 whitespace-nowrap">
                        {row.mobile || '-'}
                      </td>
                      <td className="py-2.5 px-3 text-slate-800 whitespace-nowrap">
                        {row.principleModel}
                      </td>
                      <td className="py-2.5 px-3">
                        <StatusText status={row.status} />
                      </td>
                      <td className="py-2.5 px-3 text-slate-700 leading-relaxed">
                        {row.remarks || '-'}
                      </td>
                    </tr>
                  ))
                )}
              </tbody>
            </table>
          </div>
        </div>
      )}

      {/* REPORT 2: Principle Model Wise Enquiry Report */}
      {activeTab === 'report2' && (
        <div className="p-5 space-y-6">
          <div>
            <h2 className="text-base font-semibold text-slate-900">
              Second Report: Principle Model Wise Enquiry Report
            </h2>
            <p className="text-xs text-slate-500 mt-0.5">
              {displayBundle.report2PrincipleModel.totalModels} models · {displayBundle.report2PrincipleModel.totalEnquiries} total enquiries · {displayBundle.report2PrincipleModel.totalConverted} converted/booked · Click any column to sort or click any model or status count to drill down.
            </p>
          </div>

          <div className="overflow-x-auto border border-slate-200 rounded-md">
            <table className="w-full text-left border-collapse text-xs">
              <thead>
                <tr className="bg-slate-50 border-b border-slate-200 text-slate-600 font-medium">
                  <th className="py-2.5 px-3 w-10">Sl</th>
                  <SortableHeader
                    label="Principle Model"
                    sortKey="principleModel"
                    sort={r2Sort}
                    onSort={(k, d) => setR2Sort((p) => toggleSortState(p, k, d))}
                    className="min-w-[180px]"
                  />
                  <th className="py-2.5 px-3 min-w-[120px]">Share</th>
                  <SortableHeader
                    label="Total"
                    sortKey="totalEnquiries"
                    sort={r2Sort}
                    onSort={(k, d) => setR2Sort((p) => toggleSortState(p, k, d))}
                    align="right"
                    defaultDirection="desc"
                  />
                  <SortableHeader
                    label="Open"
                    sortKey="openCount"
                    sort={r2Sort}
                    onSort={(k, d) => setR2Sort((p) => toggleSortState(p, k, d))}
                    align="right"
                    defaultDirection="desc"
                  />
                  <SortableHeader
                    label="Follow-up"
                    sortKey="followupCount"
                    sort={r2Sort}
                    onSort={(k, d) => setR2Sort((p) => toggleSortState(p, k, d))}
                    align="right"
                    defaultDirection="desc"
                  />
                  <SortableHeader
                    label="Hot"
                    sortKey="hotCount"
                    sort={r2Sort}
                    onSort={(k, d) => setR2Sort((p) => toggleSortState(p, k, d))}
                    align="right"
                    defaultDirection="desc"
                  />
                  <SortableHeader
                    label="Warm"
                    sortKey="warmCount"
                    sort={r2Sort}
                    onSort={(k, d) => setR2Sort((p) => toggleSortState(p, k, d))}
                    align="right"
                    defaultDirection="desc"
                  />
                  <SortableHeader
                    label="Cold"
                    sortKey="coldCount"
                    sort={r2Sort}
                    onSort={(k, d) => setR2Sort((p) => toggleSortState(p, k, d))}
                    align="right"
                    defaultDirection="desc"
                  />
                  <SortableHeader
                    label="Booked"
                    sortKey="bookedCount"
                    sort={r2Sort}
                    onSort={(k, d) => setR2Sort((p) => toggleSortState(p, k, d))}
                    align="right"
                    defaultDirection="desc"
                  />
                  <SortableHeader
                    label="Converted"
                    sortKey="convertedCount"
                    sort={r2Sort}
                    onSort={(k, d) => setR2Sort((p) => toggleSortState(p, k, d))}
                    align="right"
                    defaultDirection="desc"
                  />
                  <SortableHeader
                    label="Lost"
                    sortKey="lostCount"
                    sort={r2Sort}
                    onSort={(k, d) => setR2Sort((p) => toggleSortState(p, k, d))}
                    align="right"
                    defaultDirection="desc"
                  />
                  <SortableHeader
                    label="Conv %"
                    sortKey="conversionRate"
                    sort={r2Sort}
                    onSort={(k, d) => setR2Sort((p) => toggleSortState(p, k, d))}
                    align="right"
                    defaultDirection="desc"
                  />
                  <SortableHeader
                    label="Top Counter"
                    sortKey="topCounter"
                    sort={r2Sort}
                    onSort={(k, d) => setR2Sort((p) => toggleSortState(p, k, d))}
                  />
                </tr>
              </thead>
              <tbody className="divide-y divide-slate-200">
                {sortedR2Summaries.length === 0 ? (
                  <tr>
                    <td colSpan={14} className="py-6 text-center text-slate-500">
                      No Principle Model rows match the active filters.
                    </td>
                  </tr>
                ) : (
                  sortedR2Summaries.map((m, idx) => {
                    const pctWidth = Math.max(4, Math.round((m.totalEnquiries / maxModelCount) * 100));
                    return (
                      <tr
                        key={m.principleModel}
                        onClick={() => drillIntoModel(m)}
                        className="cursor-pointer hover:bg-slate-50/80 transition-colors"
                      >
                        <td className="py-2.5 px-3 font-mono tabular-nums text-slate-500">
                          {idx + 1}
                        </td>
                        <td className="py-2.5 px-3 font-medium text-blue-700 hover:underline">
                          {m.principleModel}
                        </td>
                        <td className="py-2.5 px-3">
                          <div className="w-24 h-2 bg-slate-100 rounded-xs overflow-hidden">
                            <div
                              className="h-full bg-blue-700"
                              style={{ width: `${pctWidth}%` }}
                            />
                          </div>
                        </td>
                        <td className="py-2.5 px-3 text-right font-mono tabular-nums font-semibold text-slate-900">
                          {m.totalEnquiries}
                        </td>
                        <td
                          onClick={(e) => {
                            if (m.openCount > 0) {
                              e.stopPropagation();
                              drillIntoModel(m, 'Open');
                            }
                          }}
                          className="py-2.5 px-3 text-right font-mono tabular-nums text-slate-600 hover:underline"
                        >
                          {m.openCount}
                        </td>
                        <td
                          onClick={(e) => {
                            if (m.followupCount > 0) {
                              e.stopPropagation();
                              drillIntoModel(m, 'Follow-up');
                            }
                          }}
                          className="py-2.5 px-3 text-right font-mono tabular-nums text-blue-700 hover:underline"
                        >
                          {m.followupCount}
                        </td>
                        <td
                          onClick={(e) => {
                            if (m.hotCount > 0) {
                              e.stopPropagation();
                              drillIntoModel(m, 'Hot');
                            }
                          }}
                          className="py-2.5 px-3 text-right font-mono tabular-nums text-amber-700 hover:underline"
                        >
                          {m.hotCount}
                        </td>
                        <td
                          onClick={(e) => {
                            if (m.warmCount > 0) {
                              e.stopPropagation();
                              drillIntoModel(m, 'Warm');
                            }
                          }}
                          className="py-2.5 px-3 text-right font-mono tabular-nums text-amber-600 hover:underline"
                        >
                          {m.warmCount}
                        </td>
                        <td
                          onClick={(e) => {
                            if (m.coldCount > 0) {
                              e.stopPropagation();
                              drillIntoModel(m, 'Cold');
                            }
                          }}
                          className="py-2.5 px-3 text-right font-mono tabular-nums text-slate-500 hover:underline"
                        >
                          {m.coldCount}
                        </td>
                        <td
                          onClick={(e) => {
                            if (m.bookedCount > 0) {
                              e.stopPropagation();
                              drillIntoModel(m, 'Booked');
                            }
                          }}
                          className="py-2.5 px-3 text-right font-mono tabular-nums text-emerald-700 hover:underline"
                        >
                          {m.bookedCount}
                        </td>
                        <td
                          onClick={(e) => {
                            if (m.convertedCount > 0) {
                              e.stopPropagation();
                              drillIntoModel(m, 'CONVERTED_OR_BOOKED');
                            }
                          }}
                          className="py-2.5 px-3 text-right font-mono tabular-nums font-medium text-emerald-700 hover:underline"
                        >
                          {m.convertedCount}
                        </td>
                        <td
                          onClick={(e) => {
                            if (m.lostCount > 0) {
                              e.stopPropagation();
                              drillIntoModel(m, 'Lost');
                            }
                          }}
                          className="py-2.5 px-3 text-right font-mono tabular-nums text-red-700 hover:underline"
                        >
                          {m.lostCount}
                        </td>
                        <td className="py-2.5 px-3 text-right font-mono tabular-nums font-medium text-slate-900">
                          {m.conversionRate}%
                        </td>
                        <td className="py-2.5 px-3 text-slate-600 whitespace-nowrap">
                          {m.topCounter}
                        </td>
                      </tr>
                    );
                  })
                )}
              </tbody>
            </table>
          </div>
        </div>
      )}

      {/* REPORT 3: Daily Called Data Count Report of CRM Chandrashekhar Kumar with Remarks & Daily Lead Conversion List */}
      {activeTab === 'report3' && (
        <div className="p-5 space-y-6">
          <div>
            <h2 className="text-base font-semibold text-slate-900">
              Third Report: Daily Called Data Count Report of CRM {displayBundle.report3CrmDailyCall.crmName} with Remarks &amp; Daily Lead Conversion List
            </h2>
            <p className="text-xs text-slate-500 mt-0.5">
              Total Called Records: {displayBundle.report3CrmDailyCall.totalCalls} · Converted Leads: {displayBundle.report3CrmDailyCall.totalConversions} · Conversion Rate: {displayBundle.report3CrmDailyCall.conversionRate}% · Sort any column or click any Call Date / entry to drill down.
            </p>
          </div>

          {/* 3A. Daily Called Count Summary */}
          <div className="space-y-2">
            <h3 className="text-xs font-semibold text-slate-700">
              3A. Daily Called Data Count &amp; Consolidated Remarks
            </h3>
            <div className="overflow-x-auto border border-slate-200 rounded-md">
              <table className="w-full text-left border-collapse text-xs">
                <thead>
                  <tr className="bg-slate-50 border-b border-slate-200 text-slate-600 font-medium">
                    <SortableHeader
                      label="Call Date"
                      sortKey="date"
                      sort={r3DailySort}
                      onSort={(k, d) => setR3DailySort((p) => toggleSortState(p, k, d))}
                    />
                    <SortableHeader
                      label="Called Count"
                      sortKey="totalCalled"
                      sort={r3DailySort}
                      onSort={(k, d) => setR3DailySort((p) => toggleSortState(p, k, d))}
                      align="right"
                      defaultDirection="desc"
                    />
                    <SortableHeader
                      label="Follow-up"
                      sortKey="followupCount"
                      sort={r3DailySort}
                      onSort={(k, d) => setR3DailySort((p) => toggleSortState(p, k, d))}
                      align="right"
                      defaultDirection="desc"
                    />
                    <SortableHeader
                      label="Hot"
                      sortKey="hotCount"
                      sort={r3DailySort}
                      onSort={(k, d) => setR3DailySort((p) => toggleSortState(p, k, d))}
                      align="right"
                      defaultDirection="desc"
                    />
                    <SortableHeader
                      label="Converted"
                      sortKey="convertedCount"
                      sort={r3DailySort}
                      onSort={(k, d) => setR3DailySort((p) => toggleSortState(p, k, d))}
                      align="right"
                      defaultDirection="desc"
                    />
                    <SortableHeader
                      label="Lost"
                      sortKey="lostCount"
                      sort={r3DailySort}
                      onSort={(k, d) => setR3DailySort((p) => toggleSortState(p, k, d))}
                      align="right"
                      defaultDirection="desc"
                    />
                    <SortableHeader
                      label="Remarks Summary"
                      sortKey="remarks"
                      sort={r3DailySort}
                      onSort={(k, d) => setR3DailySort((p) => toggleSortState(p, k, d))}
                      className="min-w-[280px]"
                    />
                  </tr>
                </thead>
                <tbody className="divide-y divide-slate-200">
                  {sortedR3Daily.length === 0 ? (
                    <tr>
                      <td colSpan={7} className="py-6 text-center text-slate-500">
                        No call log rows matched CRM &quot;{displayBundle.report3CrmDailyCall.crmName}&quot; with active filters.
                      </td>
                    </tr>
                  ) : (
                    sortedR3Daily.map((d) => (
                      <tr
                        key={d.date}
                        onClick={() => drillIntoCrmDay(d)}
                        className="hover:bg-slate-50/80 cursor-pointer transition-colors"
                      >
                        <td className="py-2.5 px-3 font-mono tabular-nums font-medium text-blue-700 hover:underline whitespace-nowrap">
                          {d.date}
                        </td>
                        <td className="py-2.5 px-3 text-right font-mono tabular-nums font-semibold text-slate-900">
                          {d.totalCalled}
                        </td>
                        <td
                          onClick={(e) => {
                            if (d.followupCount > 0) {
                              e.stopPropagation();
                              drillIntoCrmDay(d, 'Follow-up');
                            }
                          }}
                          className="py-2.5 px-3 text-right font-mono tabular-nums text-blue-700 hover:underline"
                        >
                          {d.followupCount}
                        </td>
                        <td
                          onClick={(e) => {
                            if (d.hotCount > 0) {
                              e.stopPropagation();
                              drillIntoCrmDay(d, 'Hot');
                            }
                          }}
                          className="py-2.5 px-3 text-right font-mono tabular-nums text-amber-700 hover:underline"
                        >
                          {d.hotCount}
                        </td>
                        <td
                          onClick={(e) => {
                            if (d.convertedCount > 0) {
                              e.stopPropagation();
                              drillIntoCrmDay(d, 'CONVERTED_OR_BOOKED');
                            }
                          }}
                          className="py-2.5 px-3 text-right font-mono tabular-nums font-medium text-emerald-700 hover:underline"
                        >
                          {d.convertedCount}
                        </td>
                        <td
                          onClick={(e) => {
                            if (d.lostCount > 0) {
                              e.stopPropagation();
                              drillIntoCrmDay(d, 'Lost');
                            }
                          }}
                          className="py-2.5 px-3 text-right font-mono tabular-nums text-red-700 hover:underline"
                        >
                          {d.lostCount}
                        </td>
                        <td className="py-2.5 px-3 text-slate-600 leading-relaxed">
                          {d.remarksList.join(' | ') || '-'}
                        </td>
                      </tr>
                    ))
                  )}
                </tbody>
              </table>
            </div>
          </div>

          {/* 3B. Detailed Call Remarks Log */}
          <div className="space-y-2">
            <h3 className="text-xs font-semibold text-slate-700">
              3B. Individual Call Log with Remarks ({sortedR3CallRows.length})
            </h3>
            <div className="overflow-x-auto border border-slate-200 rounded-md">
              <table className="w-full text-left border-collapse text-xs">
                <thead>
                  <tr className="bg-slate-50 border-b border-slate-200 text-slate-600 font-medium">
                    <th className="py-2.5 px-3 w-10">Sl</th>
                    <SortableHeader
                      label="Call Date"
                      sortKey="calledDate"
                      sort={r3CallSort}
                      onSort={(k, d) => setR3CallSort((p) => toggleSortState(p, k, d))}
                    />
                    <SortableHeader
                      label="Customer"
                      sortKey="customerName"
                      sort={r3CallSort}
                      onSort={(k, d) => setR3CallSort((p) => toggleSortState(p, k, d))}
                    />
                    <SortableHeader
                      label="Mobile"
                      sortKey="mobile"
                      sort={r3CallSort}
                      onSort={(k, d) => setR3CallSort((p) => toggleSortState(p, k, d))}
                    />
                    <SortableHeader
                      label="Counter"
                      sortKey="counterName"
                      sort={r3CallSort}
                      onSort={(k, d) => setR3CallSort((p) => toggleSortState(p, k, d))}
                    />
                    <SortableHeader
                      label="DSE Name"
                      sortKey="dseName"
                      sort={r3CallSort}
                      onSort={(k, d) => setR3CallSort((p) => toggleSortState(p, k, d))}
                    />
                    <SortableHeader
                      label="Principle Model"
                      sortKey="principleModel"
                      sort={r3CallSort}
                      onSort={(k, d) => setR3CallSort((p) => toggleSortState(p, k, d))}
                    />
                    <SortableHeader
                      label="Status"
                      sortKey="status"
                      sort={r3CallSort}
                      onSort={(k, d) => setR3CallSort((p) => toggleSortState(p, k, d))}
                    />
                    <SortableHeader
                      label="Remarks"
                      sortKey="remarks"
                      sort={r3CallSort}
                      onSort={(k, d) => setR3CallSort((p) => toggleSortState(p, k, d))}
                      className="min-w-[220px]"
                    />
                  </tr>
                </thead>
                <tbody className="divide-y divide-slate-200">
                  {sortedR3CallRows.length === 0 ? (
                    <tr>
                      <td colSpan={9} className="py-5 text-center text-slate-500">
                        No matching call log entries.
                      </td>
                    </tr>
                  ) : (
                    sortedR3CallRows.map((row, idx) => (
                      <tr
                        key={row.id}
                        onClick={() =>
                          setDrillDown({
                            title: `${row.customerName} (${row.enquiryNo})`,
                            subtitle: `CRM ${row.crmName} Call Entry · ${row.counterName} · ${row.dseName}`,
                            rows: sortedR3CallRows,
                            initialFocusedRowId: row.id,
                          })
                        }
                        className="hover:bg-slate-50/80 cursor-pointer transition-colors"
                      >
                        <td className="py-2 px-3 font-mono tabular-nums text-slate-500">
                          {idx + 1}
                        </td>
                        <td className="py-2 px-3 font-mono tabular-nums text-slate-700 whitespace-nowrap">
                          {row.calledDate || row.followupDate || row.enquiryDate || '-'}
                        </td>
                        <td className="py-2 px-3 font-medium text-blue-700 hover:underline whitespace-nowrap">
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
                        <td className="py-2 px-3 text-slate-700">{row.remarks || '-'}</td>
                      </tr>
                    ))
                  )}
                </tbody>
              </table>
            </div>
          </div>

          {/* 3C. Daily Lead Conversion List */}
          <div className="space-y-2">
            <h3 className="text-xs font-semibold text-slate-700">
              3C. Daily Lead Conversion List ({sortedR3Conversions.length})
            </h3>
            <div className="overflow-x-auto border border-slate-200 rounded-md">
              <table className="w-full text-left border-collapse text-xs">
                <thead>
                  <tr className="bg-slate-50 border-b border-slate-200 text-slate-600 font-medium">
                    <th className="py-2.5 px-3 w-10">Sl</th>
                    <SortableHeader
                      label="Conversion Date"
                      sortKey="conversionDate"
                      sort={r3ConvSort}
                      onSort={(k, d) => setR3ConvSort((p) => toggleSortState(p, k, d))}
                    />
                    <SortableHeader
                      label="Enquiry No"
                      sortKey="enquiryNo"
                      sort={r3ConvSort}
                      onSort={(k, d) => setR3ConvSort((p) => toggleSortState(p, k, d))}
                    />
                    <SortableHeader
                      label="Customer"
                      sortKey="customerName"
                      sort={r3ConvSort}
                      onSort={(k, d) => setR3ConvSort((p) => toggleSortState(p, k, d))}
                    />
                    <SortableHeader
                      label="Mobile"
                      sortKey="mobile"
                      sort={r3ConvSort}
                      onSort={(k, d) => setR3ConvSort((p) => toggleSortState(p, k, d))}
                    />
                    <SortableHeader
                      label="Counter"
                      sortKey="counterName"
                      sort={r3ConvSort}
                      onSort={(k, d) => setR3ConvSort((p) => toggleSortState(p, k, d))}
                    />
                    <SortableHeader
                      label="DSE Name"
                      sortKey="dseName"
                      sort={r3ConvSort}
                      onSort={(k, d) => setR3ConvSort((p) => toggleSortState(p, k, d))}
                    />
                    <SortableHeader
                      label="Principle Model"
                      sortKey="principleModel"
                      sort={r3ConvSort}
                      onSort={(k, d) => setR3ConvSort((p) => toggleSortState(p, k, d))}
                    />
                    <SortableHeader
                      label="Stage"
                      sortKey="status"
                      sort={r3ConvSort}
                      onSort={(k, d) => setR3ConvSort((p) => toggleSortState(p, k, d))}
                    />
                    <SortableHeader
                      label="Remarks"
                      sortKey="remarks"
                      sort={r3ConvSort}
                      onSort={(k, d) => setR3ConvSort((p) => toggleSortState(p, k, d))}
                      className="min-w-[200px]"
                    />
                  </tr>
                </thead>
                <tbody className="divide-y divide-slate-200">
                  {sortedR3Conversions.length === 0 ? (
                    <tr>
                      <td colSpan={10} className="py-5 text-center text-slate-500">
                        No converted or booked leads found in current filtered dataset.
                      </td>
                    </tr>
                  ) : (
                    sortedR3Conversions.map((row, idx) => (
                      <tr
                        key={row.id}
                        onClick={() =>
                          setDrillDown({
                            title: `${row.customerName} (${row.enquiryNo})`,
                            subtitle: `Converted Lead · ${row.counterName} · ${row.dseName}`,
                            rows: sortedR3Conversions,
                            initialFocusedRowId: row.id,
                          })
                        }
                        className="hover:bg-slate-50/80 cursor-pointer transition-colors"
                      >
                        <td className="py-2 px-3 font-mono tabular-nums text-slate-500">
                          {idx + 1}
                        </td>
                        <td className="py-2 px-3 font-mono tabular-nums text-slate-700 whitespace-nowrap">
                          {row.conversionDate || row.calledDate || row.enquiryDate || '-'}
                        </td>
                        <td className="py-2 px-3 font-mono tabular-nums text-slate-600 whitespace-nowrap">
                          {row.enquiryNo}
                        </td>
                        <td className="py-2 px-3 font-medium text-blue-700 hover:underline whitespace-nowrap">
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
                        <td className="py-2 px-3 text-slate-700">{row.remarks || '-'}</td>
                      </tr>
                    ))
                  )}
                </tbody>
              </table>
            </div>
          </div>
        </div>
      )}

      {/* REPORT 4: Counter Wise Enquiry Report with Status Summaries */}
      {activeTab === 'report4' && (
        <div className="p-5 space-y-6">
          <div>
            <h2 className="text-base font-semibold text-slate-900">
              Fourth Report: Counter Wise Enquiry Report with Status Summaries
            </h2>
            <p className="text-xs text-slate-500 mt-0.5">
              All 7 dealership counters · {displayBundle.report4CounterWise.totalEnquiries} total enquiries · {displayBundle.report4CounterWise.totalConverted} converted/booked ({displayBundle.report4CounterWise.overallConversionRate}% conversion rate). Click any column header to sort, click arrow to expand DSEs, or click any count cell to drill down.
            </p>
          </div>

          <div className="overflow-x-auto border border-slate-200 rounded-md">
            <table className="w-full text-left border-collapse text-xs">
              <thead>
                <tr className="bg-slate-50 border-b border-slate-200 text-slate-600 font-medium">
                  <th className="py-2.5 px-3 w-8"></th>
                  <SortableHeader
                    label="Counter Name"
                    sortKey="counterName"
                    sort={r4Sort}
                    onSort={(k, d) => setR4Sort((p) => toggleSortState(p, k, d))}
                  />
                  <SortableHeader
                    label="Total"
                    sortKey="totalEnquiries"
                    sort={r4Sort}
                    onSort={(k, d) => setR4Sort((p) => toggleSortState(p, k, d))}
                    align="right"
                    defaultDirection="desc"
                  />
                  <SortableHeader
                    label="Open"
                    sortKey="openCount"
                    sort={r4Sort}
                    onSort={(k, d) => setR4Sort((p) => toggleSortState(p, k, d))}
                    align="right"
                    defaultDirection="desc"
                  />
                  <SortableHeader
                    label="Follow-up"
                    sortKey="followupCount"
                    sort={r4Sort}
                    onSort={(k, d) => setR4Sort((p) => toggleSortState(p, k, d))}
                    align="right"
                    defaultDirection="desc"
                  />
                  <SortableHeader
                    label="Hot"
                    sortKey="hotCount"
                    sort={r4Sort}
                    onSort={(k, d) => setR4Sort((p) => toggleSortState(p, k, d))}
                    align="right"
                    defaultDirection="desc"
                  />
                  <SortableHeader
                    label="Warm"
                    sortKey="warmCount"
                    sort={r4Sort}
                    onSort={(k, d) => setR4Sort((p) => toggleSortState(p, k, d))}
                    align="right"
                    defaultDirection="desc"
                  />
                  <SortableHeader
                    label="Cold"
                    sortKey="coldCount"
                    sort={r4Sort}
                    onSort={(k, d) => setR4Sort((p) => toggleSortState(p, k, d))}
                    align="right"
                    defaultDirection="desc"
                  />
                  <SortableHeader
                    label="Booked"
                    sortKey="bookedCount"
                    sort={r4Sort}
                    onSort={(k, d) => setR4Sort((p) => toggleSortState(p, k, d))}
                    align="right"
                    defaultDirection="desc"
                  />
                  <SortableHeader
                    label="Converted"
                    sortKey="convertedCount"
                    sort={r4Sort}
                    onSort={(k, d) => setR4Sort((p) => toggleSortState(p, k, d))}
                    align="right"
                    defaultDirection="desc"
                  />
                  <SortableHeader
                    label="Lost"
                    sortKey="lostCount"
                    sort={r4Sort}
                    onSort={(k, d) => setR4Sort((p) => toggleSortState(p, k, d))}
                    align="right"
                    defaultDirection="desc"
                  />
                  <SortableHeader
                    label="Conv %"
                    sortKey="conversionRate"
                    sort={r4Sort}
                    onSort={(k, d) => setR4Sort((p) => toggleSortState(p, k, d))}
                    align="right"
                    defaultDirection="desc"
                  />
                </tr>
              </thead>
              <tbody className="divide-y divide-slate-200">
                {sortedR4Counters.map((counter) => {
                  const isExpanded = Boolean(expandedCounters[counter.counterName]);
                  const sortedDses = r4Sort
                    ? sortItems(counter.dseBreakdowns, r4Sort, {
                        counterName: (d) => d.dseName,
                        totalEnquiries: (d) => d.totalEnquiries,
                        openCount: (d) => d.openCount,
                        followupCount: (d) => d.followupCount,
                        hotCount: (d) => d.hotCount,
                        warmCount: (d) => d.warmCount,
                        coldCount: (d) => d.coldCount,
                        bookedCount: (d) => d.bookedCount,
                        convertedCount: (d) => d.convertedCount,
                        lostCount: (d) => d.lostCount,
                        conversionRate: (d) => d.conversionRate,
                      })
                    : counter.dseBreakdowns;

                  return (
                    <Fragment key={counter.counterName}>
                      <tr className="bg-white hover:bg-slate-50 transition-colors">
                        <td
                          onClick={() => toggleCounterExpand(counter.counterName)}
                          className="py-2.5 px-3 text-slate-500 cursor-pointer"
                        >
                          {isExpanded ? (
                            <ChevronDown className="w-4 h-4" />
                          ) : (
                            <ChevronRight className="w-4 h-4" />
                          )}
                        </td>
                        <td
                          onClick={() => toggleCounterExpand(counter.counterName)}
                          className="py-2.5 px-3 font-semibold text-slate-900 whitespace-nowrap cursor-pointer"
                        >
                          {counter.counterNumber <= 7
                            ? `${counter.counterNumber}. ${counter.counterName}`
                            : counter.counterName}
                        </td>
                        <td
                          onClick={() =>
                            counter.totalEnquiries > 0 && drillIntoCounterMetric(counter)
                          }
                          className="py-2.5 px-3 text-right font-mono tabular-nums font-semibold text-blue-700 hover:underline cursor-pointer"
                        >
                          {counter.totalEnquiries}
                        </td>
                        <td
                          onClick={() =>
                            counter.openCount > 0 && drillIntoCounterMetric(counter, undefined, 'Open')
                          }
                          className="py-2.5 px-3 text-right font-mono tabular-nums text-slate-600 hover:underline cursor-pointer"
                        >
                          {counter.openCount}
                        </td>
                        <td
                          onClick={() =>
                            counter.followupCount > 0 &&
                            drillIntoCounterMetric(counter, undefined, 'Follow-up')
                          }
                          className="py-2.5 px-3 text-right font-mono tabular-nums text-blue-700 hover:underline cursor-pointer"
                        >
                          {counter.followupCount}
                        </td>
                        <td
                          onClick={() =>
                            counter.hotCount > 0 && drillIntoCounterMetric(counter, undefined, 'Hot')
                          }
                          className="py-2.5 px-3 text-right font-mono tabular-nums text-amber-700 hover:underline cursor-pointer"
                        >
                          {counter.hotCount}
                        </td>
                        <td
                          onClick={() =>
                            counter.warmCount > 0 && drillIntoCounterMetric(counter, undefined, 'Warm')
                          }
                          className="py-2.5 px-3 text-right font-mono tabular-nums text-amber-600 hover:underline cursor-pointer"
                        >
                          {counter.warmCount}
                        </td>
                        <td
                          onClick={() =>
                            counter.coldCount > 0 && drillIntoCounterMetric(counter, undefined, 'Cold')
                          }
                          className="py-2.5 px-3 text-right font-mono tabular-nums text-slate-500 hover:underline cursor-pointer"
                        >
                          {counter.coldCount}
                        </td>
                        <td
                          onClick={() =>
                            counter.bookedCount > 0 &&
                            drillIntoCounterMetric(counter, undefined, 'Booked')
                          }
                          className="py-2.5 px-3 text-right font-mono tabular-nums text-emerald-700 hover:underline cursor-pointer"
                        >
                          {counter.bookedCount}
                        </td>
                        <td
                          onClick={() =>
                            counter.convertedCount > 0 &&
                            drillIntoCounterMetric(counter, undefined, 'Converted')
                          }
                          className="py-2.5 px-3 text-right font-mono tabular-nums font-medium text-emerald-700 hover:underline cursor-pointer"
                        >
                          {counter.convertedCount}
                        </td>
                        <td
                          onClick={() =>
                            counter.lostCount > 0 && drillIntoCounterMetric(counter, undefined, 'Lost')
                          }
                          className="py-2.5 px-3 text-right font-mono tabular-nums text-red-700 hover:underline cursor-pointer"
                        >
                          {counter.lostCount}
                        </td>
                        <td className="py-2.5 px-3 text-right font-mono tabular-nums font-semibold text-slate-900">
                          {counter.conversionRate}%
                        </td>
                      </tr>
                      {isExpanded &&
                        sortedDses.map((dse) => (
                          <tr
                            key={`${counter.counterName}-${dse.dseName}`}
                            className="bg-slate-50/70 text-slate-700"
                          >
                            <td className="py-2 px-3"></td>
                            <td
                              onClick={() =>
                                dse.totalEnquiries > 0 && drillIntoCounterMetric(counter, dse)
                              }
                              className="py-2 px-3 pl-8 text-blue-700 hover:underline whitespace-nowrap cursor-pointer"
                            >
                              {dse.dseName}
                            </td>
                            <td
                              onClick={() =>
                                dse.totalEnquiries > 0 && drillIntoCounterMetric(counter, dse)
                              }
                              className="py-2 px-3 text-right font-mono tabular-nums hover:underline cursor-pointer"
                            >
                              {dse.totalEnquiries}
                            </td>
                            <td
                              onClick={() =>
                                dse.openCount > 0 && drillIntoCounterMetric(counter, dse, 'Open')
                              }
                              className="py-2 px-3 text-right font-mono tabular-nums hover:underline cursor-pointer"
                            >
                              {dse.openCount}
                            </td>
                            <td
                              onClick={() =>
                                dse.followupCount > 0 &&
                                drillIntoCounterMetric(counter, dse, 'Follow-up')
                              }
                              className="py-2 px-3 text-right font-mono tabular-nums hover:underline cursor-pointer"
                            >
                              {dse.followupCount}
                            </td>
                            <td
                              onClick={() =>
                                dse.hotCount > 0 && drillIntoCounterMetric(counter, dse, 'Hot')
                              }
                              className="py-2 px-3 text-right font-mono tabular-nums hover:underline cursor-pointer"
                            >
                              {dse.hotCount}
                            </td>
                            <td
                              onClick={() =>
                                dse.warmCount > 0 && drillIntoCounterMetric(counter, dse, 'Warm')
                              }
                              className="py-2 px-3 text-right font-mono tabular-nums hover:underline cursor-pointer"
                            >
                              {dse.warmCount}
                            </td>
                            <td
                              onClick={() =>
                                dse.coldCount > 0 && drillIntoCounterMetric(counter, dse, 'Cold')
                              }
                              className="py-2 px-3 text-right font-mono tabular-nums hover:underline cursor-pointer"
                            >
                              {dse.coldCount}
                            </td>
                            <td
                              onClick={() =>
                                dse.bookedCount > 0 &&
                                drillIntoCounterMetric(counter, dse, 'Booked')
                              }
                              className="py-2 px-3 text-right font-mono tabular-nums hover:underline cursor-pointer"
                            >
                              {dse.bookedCount}
                            </td>
                            <td
                              onClick={() =>
                                dse.convertedCount > 0 &&
                                drillIntoCounterMetric(counter, dse, 'Converted')
                              }
                              className="py-2 px-3 text-right font-mono tabular-nums hover:underline cursor-pointer"
                            >
                              {dse.convertedCount}
                            </td>
                            <td
                              onClick={() =>
                                dse.lostCount > 0 && drillIntoCounterMetric(counter, dse, 'Lost')
                              }
                              className="py-2 px-3 text-right font-mono tabular-nums hover:underline cursor-pointer"
                            >
                              {dse.lostCount}
                            </td>
                            <td className="py-2 px-3 text-right font-mono tabular-nums">
                              {dse.conversionRate}%
                            </td>
                          </tr>
                        ))}
                    </Fragment>
                  );
                })}
              </tbody>
            </table>
          </div>
        </div>
      )}

      {activeTab === 'custom' && (
        <CustomReportBuilder
          rows={rawRows}
          onExportCustomExcel={onExportCustomExcel}
          onDrillDown={setDrillDown}
        />
      )}

      <DrillDownDrawer
        context={drillDown}
        onClose={() => setDrillDown(null)}
        onExportSlice={async (_title, rows) => {
          await onExportExcel('all', rows);
        }}
      />
    </section>
  );
}
