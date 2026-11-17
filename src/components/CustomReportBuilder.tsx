import { useMemo, useState } from 'react';
import { Download, RotateCcw } from 'lucide-react';
import {
  ALL_NORMALIZED_STATUSES,
  buildCustomReport,
  COUNTER_DEFINITIONS,
  CUSTOM_REPORT_COLUMNS,
} from '../shared/reportEngine';
import type {
  CanonicalCounterName,
  CustomReportColumnKey,
  CustomReportConfig,
  CustomReportGroupBy,
  CustomReportGroupRow,
  EnquiryRow,
  NormalizedStatus,
} from '../shared/types';
import type { DrillDownContext } from './DrillDownDrawer';
import { SortableHeader, sortItems, toggleSortState, type SortDirection, type SortState } from './SortableHeader';
import { StatusText } from './StatusText';

interface CustomReportBuilderProps {
  rows: EnquiryRow[];
  onExportCustomExcel: (config: CustomReportConfig) => Promise<void>;
  onDrillDown?: (ctx: DrillDownContext) => void;
}

const DEFAULT_COLUMNS: CustomReportColumnKey[] = [
  'enquiryNo',
  'enquiryDate',
  'customerName',
  'mobile',
  'counterName',
  'dseName',
  'principleModel',
  'status',
  'followupDate',
  'crmName',
  'remarks',
];

type GroupSortKey =
  | 'groupKey'
  | 'total'
  | 'openCount'
  | 'followupCount'
  | 'hotCount'
  | 'convertedCount'
  | 'lostCount'
  | 'conversionRate';

export function CustomReportBuilder({
  rows,
  onExportCustomExcel,
  onDrillDown,
}: CustomReportBuilderProps) {
  const [title, setTitle] = useState('Custom Dealership Report');
  const [selectedCounters, setSelectedCounters] = useState<CanonicalCounterName[]>([]);
  const [selectedDses, setSelectedDses] = useState<string[]>([]);
  const [selectedStatuses, setSelectedStatuses] = useState<NormalizedStatus[]>([]);
  const [selectedModels, setSelectedModels] = useState<string[]>([]);
  const [crmName, setCrmName] = useState<string>('ALL');
  const [dateFrom, setDateFrom] = useState<string>('');
  const [dateTo, setDateTo] = useState<string>('');
  const [searchQuery, setSearchQuery] = useState<string>('');
  const [groupBy, setGroupBy] = useState<CustomReportGroupBy>('counterName');
  const [columns, setColumns] = useState<CustomReportColumnKey[]>(DEFAULT_COLUMNS);
  const [isExporting, setIsExporting] = useState(false);

  const [groupSort, setGroupSort] = useState<SortState<GroupSortKey> | null>(null);
  const [rowSort, setRowSort] = useState<SortState<CustomReportColumnKey> | null>(null);

  const availableDses = useMemo(() => {
    const set = new Set<string>();
    for (const r of rows) {
      if (selectedCounters.length > 0 && !selectedCounters.includes(r.counterName)) continue;
      if (r.dseName) set.add(r.dseName);
    }
    return Array.from(set).sort();
  }, [rows, selectedCounters]);

  const availableModels = useMemo(() => {
    const set = new Set<string>();
    for (const r of rows) {
      if (r.principleModel) set.add(r.principleModel);
    }
    return Array.from(set).sort();
  }, [rows]);

  const availableCrms = useMemo(() => {
    const set = new Set<string>(['Chandrashekhar Kumar']);
    for (const r of rows) {
      if (r.crmName) set.add(r.crmName);
    }
    return Array.from(set).sort();
  }, [rows]);

  const config: CustomReportConfig = useMemo(
    () => ({
      title: title.trim() || 'Custom Report',
      counters: selectedCounters,
      dseNames: selectedDses,
      statuses: selectedStatuses,
      principleModels: selectedModels,
      crmName,
      dateFrom,
      dateTo,
      searchQuery,
      groupBy,
      columns,
    }),
    [
      title,
      selectedCounters,
      selectedDses,
      selectedStatuses,
      selectedModels,
      crmName,
      dateFrom,
      dateTo,
      searchQuery,
      groupBy,
      columns,
    ]
  );

  const result = useMemo(() => buildCustomReport(rows, config), [rows, config]);

  const sortedGroups = useMemo(() => {
    return sortItems(result.groups, groupSort, {
      groupKey: (g) => g.groupKey,
      total: (g) => g.total,
      openCount: (g) => g.openCount,
      followupCount: (g) => g.followupCount,
      hotCount: (g) => g.hotCount,
      convertedCount: (g) => g.convertedCount,
      lostCount: (g) => g.lostCount,
      conversionRate: (g) => g.conversionRate,
    });
  }, [result.groups, groupSort]);

  const sortedRows = useMemo(() => {
    const accessors = Object.fromEntries(
      CUSTOM_REPORT_COLUMNS.map((c) => [c.key, (r: EnquiryRow) => r[c.key] || ''])
    ) as Record<CustomReportColumnKey, (item: EnquiryRow) => string>;
    return sortItems(result.rows, rowSort, accessors);
  }, [result.rows, rowSort]);

  const toggleCounter = (name: CanonicalCounterName) => {
    setSelectedCounters((prev) =>
      prev.includes(name) ? prev.filter((c) => c !== name) : [...prev, name]
    );
  };

  const toggleDse = (name: string) => {
    setSelectedDses((prev) =>
      prev.includes(name) ? prev.filter((d) => d !== name) : [...prev, name]
    );
  };

  const toggleStatus = (status: NormalizedStatus) => {
    setSelectedStatuses((prev) =>
      prev.includes(status) ? prev.filter((s) => s !== status) : [...prev, status]
    );
  };

  const toggleModel = (model: string) => {
    setSelectedModels((prev) =>
      prev.includes(model) ? prev.filter((m) => m !== model) : [...prev, model]
    );
  };

  const toggleColumn = (key: CustomReportColumnKey) => {
    setColumns((prev) =>
      prev.includes(key)
        ? prev.length > 1
          ? prev.filter((c) => c !== key)
          : prev
        : [...prev, key]
    );
  };

  const handleReset = () => {
    setTitle('Custom Dealership Report');
    setSelectedCounters([]);
    setSelectedDses([]);
    setSelectedStatuses([]);
    setSelectedModels([]);
    setCrmName('ALL');
    setDateFrom('');
    setDateTo('');
    setSearchQuery('');
    setGroupBy('counterName');
    setColumns(DEFAULT_COLUMNS);
    setGroupSort(null);
    setRowSort(null);
  };

  const handleDownload = async () => {
    setIsExporting(true);
    try {
      await onExportCustomExcel(config);
    } finally {
      setIsExporting(false);
    }
  };

  const handleDrillGroup = (g: CustomReportGroupRow, statusFilter?: NormalizedStatus | 'CONVERTED_OR_BOOKED') => {
    if (!onDrillDown) return;
    const matched = result.rows.filter((row) => {
      let key = 'Unspecified';
      if (groupBy === 'counterName') key = row.counterName || 'Unassigned';
      else if (groupBy === 'dseName') key = row.dseName || 'Unassigned';
      else if (groupBy === 'principleModel') key = row.principleModel || 'Unspecified Model';
      else if (groupBy === 'status') key = row.status;
      else if (groupBy === 'crmName') key = row.crmName || 'Unassigned Caller';
      else if (groupBy === 'date') {
        key = row.followupDate || row.calledDate || row.enquiryDate || 'Unspecified Date';
      }
      if (key !== g.groupKey) return false;
      if (statusFilter === 'CONVERTED_OR_BOOKED') return row.isConverted;
      if (statusFilter) return row.status === statusFilter;
      return true;
    });
    onDrillDown({
      title: `${g.groupKey}${statusFilter ? ` · ${statusFilter === 'CONVERTED_OR_BOOKED' ? 'Converted / Booked' : statusFilter}` : ''}`,
      subtitle: `Custom Report (${title}) · ${matched.length} matching records`,
      rows: matched,
    });
  };

  const columnLabelMap = new Map(CUSTOM_REPORT_COLUMNS.map((c) => [c.key, c.label]));

  return (
    <div className="p-5 space-y-6">
      <div className="flex flex-col lg:flex-row lg:items-center justify-between gap-4">
        <div>
          <h2 className="text-base font-semibold text-slate-900">
            Dynamic Custom Report Builder
          </h2>
          <p className="text-xs text-slate-500 mt-0.5">
            Combine filters across Counters, DSE Names, Statuses, Principle Models, CRM Callers, and Date Range, sort any column, drill down into groups, and export colorful Excel workbooks with charts.
          </p>
        </div>

        <div className="flex flex-wrap items-center gap-2.5 shrink-0">
          <button
            type="button"
            onClick={handleReset}
            className="inline-flex items-center gap-1.5 px-3 py-1.5 text-xs font-medium text-slate-700 bg-white border border-slate-300 hover:bg-slate-50 rounded-md transition-colors whitespace-nowrap cursor-pointer"
          >
            <RotateCcw className="w-3.5 h-3.5" />
            <span>Reset Filters</span>
          </button>
          <button
            type="button"
            disabled={result.totalMatched === 0 || isExporting}
            onClick={handleDownload}
            className="inline-flex items-center gap-1.5 px-3.5 py-1.5 text-xs font-medium text-white bg-blue-700 hover:bg-blue-800 disabled:opacity-50 rounded-md transition-colors whitespace-nowrap cursor-pointer"
          >
            <Download className="w-3.5 h-3.5" />
            <span>{isExporting ? 'Exporting...' : 'Download Custom Excel (.xlsx)'}</span>
          </button>
        </div>
      </div>

      {/* Configuration Controls */}
      <div className="border border-slate-200 rounded-md p-4 bg-slate-50/60 space-y-4">
        <div className="grid grid-cols-1 sm:grid-cols-2 lg:grid-cols-5 gap-3">
          <div>
            <label className="block text-xs font-medium text-slate-700 mb-1">
              Report Title
            </label>
            <input
              type="text"
              value={title}
              onChange={(e) => setTitle(e.target.value)}
              className="w-full px-2.5 py-1.5 text-xs bg-white border border-slate-300 rounded-md focus:outline-none focus:ring-2 focus:ring-blue-600"
            />
          </div>

          <div>
            <label className="block text-xs font-medium text-slate-700 mb-1">
              Group Summary By
            </label>
            <select
              value={groupBy}
              onChange={(e) => setGroupBy(e.target.value as CustomReportGroupBy)}
              className="w-full px-2.5 py-1.5 text-xs bg-white border border-slate-300 rounded-md focus:outline-none focus:ring-2 focus:ring-blue-600"
            >
              <option value="counterName">Counter Name</option>
              <option value="dseName">DSE Name</option>
              <option value="principleModel">Principle Model</option>
              <option value="status">Status</option>
              <option value="crmName">CRM / Caller</option>
              <option value="date">Date</option>
              <option value="none">None (Detail Rows Only)</option>
            </select>
          </div>

          <div>
            <label className="block text-xs font-medium text-slate-700 mb-1">
              CRM / Caller Filter
            </label>
            <select
              value={crmName}
              onChange={(e) => setCrmName(e.target.value)}
              className="w-full px-2.5 py-1.5 text-xs bg-white border border-slate-300 rounded-md focus:outline-none focus:ring-2 focus:ring-blue-600"
            >
              <option value="ALL">All Callers</option>
              {availableCrms.map((c) => (
                <option key={c} value={c}>
                  {c}
                </option>
              ))}
              <option value="__unassigned__">Unassigned Caller</option>
            </select>
          </div>

          <div>
            <label className="block text-xs font-medium text-slate-700 mb-1">
              From Date
            </label>
            <input
              type="date"
              value={dateFrom}
              onChange={(e) => setDateFrom(e.target.value)}
              className="w-full px-2.5 py-1.5 text-xs font-mono bg-white border border-slate-300 rounded-md focus:outline-none focus:ring-2 focus:ring-blue-600"
            />
          </div>

          <div>
            <label className="block text-xs font-medium text-slate-700 mb-1">
              To Date
            </label>
            <input
              type="date"
              value={dateTo}
              onChange={(e) => setDateTo(e.target.value)}
              className="w-full px-2.5 py-1.5 text-xs font-mono bg-white border border-slate-300 rounded-md focus:outline-none focus:ring-2 focus:ring-blue-600"
            />
          </div>
        </div>

        <div className="grid grid-cols-1 lg:grid-cols-2 xl:grid-cols-4 gap-4 pt-2 border-t border-slate-200">
          {/* Counters Filter */}
          <div>
            <p className="text-xs font-medium text-slate-700 mb-1.5">
              Counters {selectedCounters.length > 0 ? `(${selectedCounters.length})` : '(All)'}
            </p>
            <div className="flex flex-wrap gap-1.5">
              {[...COUNTER_DEFINITIONS.map((c) => c.name), 'Unassigned' as CanonicalCounterName].map(
                (name) => {
                  const active = selectedCounters.includes(name);
                  return (
                    <button
                      key={name}
                      type="button"
                      onClick={() => toggleCounter(name)}
                      className={`px-2 py-1 text-xs rounded border transition-colors cursor-pointer ${
                        active
                          ? 'bg-blue-700 border-blue-700 text-white'
                          : 'bg-white border-slate-300 text-slate-700 hover:bg-slate-100'
                      }`}
                    >
                      {name}
                    </button>
                  );
                }
              )}
            </div>
          </div>

          {/* DSE Names Filter */}
          <div>
            <p className="text-xs font-medium text-slate-700 mb-1.5">
              DSE Names {selectedDses.length > 0 ? `(${selectedDses.length})` : '(All)'}
            </p>
            <div className="flex flex-wrap gap-1.5 max-h-24 overflow-y-auto">
              {availableDses.map((dse) => {
                const active = selectedDses.includes(dse);
                return (
                  <button
                    key={dse}
                    type="button"
                    onClick={() => toggleDse(dse)}
                    className={`px-2 py-1 text-xs rounded border transition-colors cursor-pointer ${
                      active
                        ? 'bg-blue-700 border-blue-700 text-white'
                        : 'bg-white border-slate-300 text-slate-700 hover:bg-slate-100'
                    }`}
                  >
                    {dse}
                  </button>
                );
              })}
            </div>
          </div>

          {/* Statuses Filter */}
          <div>
            <p className="text-xs font-medium text-slate-700 mb-1.5">
              Statuses {selectedStatuses.length > 0 ? `(${selectedStatuses.length})` : '(All)'}
            </p>
            <div className="flex flex-wrap gap-1.5">
              {ALL_NORMALIZED_STATUSES.map((st) => {
                const active = selectedStatuses.includes(st);
                return (
                  <button
                    key={st}
                    type="button"
                    onClick={() => toggleStatus(st)}
                    className={`px-2 py-1 text-xs rounded border transition-colors cursor-pointer ${
                      active
                        ? 'bg-blue-700 border-blue-700 text-white'
                        : 'bg-white border-slate-300 text-slate-700 hover:bg-slate-100'
                    }`}
                  >
                    {st}
                  </button>
                );
              })}
            </div>
          </div>

          {/* Models Filter */}
          <div>
            <p className="text-xs font-medium text-slate-700 mb-1.5">
              Principle Models {selectedModels.length > 0 ? `(${selectedModels.length})` : '(All)'}
            </p>
            <div className="flex flex-wrap gap-1.5 max-h-24 overflow-y-auto">
              {availableModels.map((model) => {
                const active = selectedModels.includes(model);
                return (
                  <button
                    key={model}
                    type="button"
                    onClick={() => toggleModel(model)}
                    className={`px-2 py-1 text-xs rounded border transition-colors cursor-pointer ${
                      active
                        ? 'bg-blue-700 border-blue-700 text-white'
                        : 'bg-white border-slate-300 text-slate-700 hover:bg-slate-100'
                    }`}
                  >
                    {model}
                  </button>
                );
              })}
            </div>
          </div>
        </div>

        {/* Columns + Search */}
        <div className="pt-2 border-t border-slate-200 flex flex-col lg:flex-row lg:items-center justify-between gap-3">
          <div className="flex flex-wrap items-center gap-1.5">
            <span className="text-xs font-medium text-slate-700 mr-1">Columns:</span>
            {CUSTOM_REPORT_COLUMNS.map((col) => {
              const active = columns.includes(col.key);
              return (
                <button
                  key={col.key}
                  type="button"
                  onClick={() => toggleColumn(col.key)}
                  className={`px-2 py-1 text-xs rounded border transition-colors cursor-pointer ${
                    active
                      ? 'bg-slate-900 border-slate-900 text-white'
                      : 'bg-white border-slate-300 text-slate-600 hover:bg-slate-100'
                  }`}
                >
                  {col.label}
                </button>
              );
            })}
          </div>

          <input
            type="text"
            value={searchQuery}
            onChange={(e) => setSearchQuery(e.target.value)}
            placeholder="Keyword filter..."
            className="w-full lg:w-56 px-2.5 py-1.5 text-xs bg-white border border-slate-300 rounded-md focus:outline-none focus:ring-2 focus:ring-blue-600"
          />
        </div>
      </div>

      {/* Live Output Metrics */}
      <div className="flex flex-wrap items-center gap-4 text-xs text-slate-600">
        <span>
          Matched Records: <strong className="font-mono tabular-nums text-slate-900">{result.totalMatched}</strong>
        </span>
        <span aria-hidden="true">·</span>
        <span>
          Converted / Booked: <strong className="font-mono tabular-nums text-emerald-700">{result.convertedMatched}</strong>
        </span>
        <span aria-hidden="true">·</span>
        <span>
          Conversion Rate: <strong className="font-mono tabular-nums text-slate-900">{result.conversionRate}%</strong>
        </span>
      </div>

      {/* Grouped Summary Table */}
      {groupBy !== 'none' && (
        <div className="overflow-x-auto border border-slate-200 rounded-md">
          <table className="w-full text-left border-collapse text-xs">
            <thead>
              <tr className="bg-slate-50 border-b border-slate-200 text-slate-600 font-medium">
                <SortableHeader
                  label={`Group (${groupBy})`}
                  sortKey="groupKey"
                  sort={groupSort}
                  onSort={(k, d) => setGroupSort((p) => toggleSortState(p, k, d))}
                />
                <SortableHeader
                  label="Total"
                  sortKey="total"
                  sort={groupSort}
                  onSort={(k, d) => setGroupSort((p) => toggleSortState(p, k, d))}
                  align="right"
                  defaultDirection="desc"
                />
                <SortableHeader
                  label="Open"
                  sortKey="openCount"
                  sort={groupSort}
                  onSort={(k, d) => setGroupSort((p) => toggleSortState(p, k, d))}
                  align="right"
                  defaultDirection="desc"
                />
                <SortableHeader
                  label="Follow-up"
                  sortKey="followupCount"
                  sort={groupSort}
                  onSort={(k, d) => setGroupSort((p) => toggleSortState(p, k, d))}
                  align="right"
                  defaultDirection="desc"
                />
                <SortableHeader
                  label="Hot"
                  sortKey="hotCount"
                  sort={groupSort}
                  onSort={(k, d) => setGroupSort((p) => toggleSortState(p, k, d))}
                  align="right"
                  defaultDirection="desc"
                />
                <SortableHeader
                  label="Converted / Booked"
                  sortKey="convertedCount"
                  sort={groupSort}
                  onSort={(k, d) => setGroupSort((p) => toggleSortState(p, k, d))}
                  align="right"
                  defaultDirection="desc"
                />
                <SortableHeader
                  label="Lost"
                  sortKey="lostCount"
                  sort={groupSort}
                  onSort={(k, d) => setGroupSort((p) => toggleSortState(p, k, d))}
                  align="right"
                  defaultDirection="desc"
                />
                <SortableHeader
                  label="Conv %"
                  sortKey="conversionRate"
                  sort={groupSort}
                  onSort={(k, d) => setGroupSort((p) => toggleSortState(p, k, d))}
                  align="right"
                  defaultDirection="desc"
                />
              </tr>
            </thead>
            <tbody className="divide-y divide-slate-200">
              {sortedGroups.length === 0 ? (
                <tr>
                  <td colSpan={8} className="py-6 text-center text-slate-500">
                    No records match the selected custom filters.
                  </td>
                </tr>
              ) : (
                sortedGroups.map((g) => (
                  <tr
                    key={g.groupKey}
                    onClick={() => handleDrillGroup(g)}
                    className="hover:bg-slate-50/80 cursor-pointer transition-colors"
                  >
                    <td className="py-2.5 px-3 font-medium text-blue-700 hover:underline">
                      {g.groupKey}
                    </td>
                    <td className="py-2.5 px-3 text-right font-mono tabular-nums font-semibold text-slate-900">
                      {g.total}
                    </td>
                    <td
                      onClick={(e) => {
                        if (g.openCount > 0) {
                          e.stopPropagation();
                          handleDrillGroup(g, 'Open');
                        }
                      }}
                      className="py-2.5 px-3 text-right font-mono tabular-nums text-slate-600 hover:underline"
                    >
                      {g.openCount}
                    </td>
                    <td
                      onClick={(e) => {
                        if (g.followupCount > 0) {
                          e.stopPropagation();
                          handleDrillGroup(g, 'Follow-up');
                        }
                      }}
                      className="py-2.5 px-3 text-right font-mono tabular-nums text-blue-700 hover:underline"
                    >
                      {g.followupCount}
                    </td>
                    <td
                      onClick={(e) => {
                        if (g.hotCount > 0) {
                          e.stopPropagation();
                          handleDrillGroup(g, 'Hot');
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
                          handleDrillGroup(g, 'CONVERTED_OR_BOOKED');
                        }
                      }}
                      className="py-2.5 px-3 text-right font-mono tabular-nums text-emerald-700 font-medium hover:underline"
                    >
                      {g.convertedCount}
                    </td>
                    <td
                      onClick={(e) => {
                        if (g.lostCount > 0) {
                          e.stopPropagation();
                          handleDrillGroup(g, 'Lost');
                        }
                      }}
                      className="py-2.5 px-3 text-right font-mono tabular-nums text-red-700 hover:underline"
                    >
                      {g.lostCount}
                    </td>
                    <td className="py-2.5 px-3 text-right font-mono tabular-nums font-medium text-slate-900">
                      {g.conversionRate}%
                    </td>
                  </tr>
                ))
              )}
            </tbody>
          </table>
        </div>
      )}

      {/* Custom Columns Detail Table */}
      <div className="overflow-x-auto border border-slate-200 rounded-md">
        <table className="w-full text-left border-collapse text-xs">
          <thead>
            <tr className="bg-slate-50 border-b border-slate-200 text-slate-600 font-medium">
              <th className="py-2.5 px-3 w-10">Sl</th>
              {columns.map((colKey) => (
                <SortableHeader
                  key={colKey}
                  label={columnLabelMap.get(colKey) || colKey}
                  sortKey={colKey}
                  sort={rowSort}
                  onSort={(k: CustomReportColumnKey, d: SortDirection = 'asc') =>
                    setRowSort((p) => toggleSortState(p, k, d))
                  }
                />
              ))}
            </tr>
          </thead>
          <tbody className="divide-y divide-slate-200">
            {sortedRows.length === 0 ? (
              <tr>
                <td colSpan={columns.length + 1} className="py-6 text-center text-slate-500">
                  No rows match your custom filter criteria.
                </td>
              </tr>
            ) : (
              sortedRows.map((row, idx) => (
                <tr
                  key={row.id}
                  onClick={() =>
                    onDrillDown?.({
                      title: `${row.customerName} (${row.enquiryNo})`,
                      subtitle: `${row.counterName} · ${row.dseName} · ${row.principleModel}`,
                      rows: sortedRows,
                      initialFocusedRowId: row.id,
                    })
                  }
                  className="hover:bg-slate-50/80 cursor-pointer transition-colors"
                >
                  <td className="py-2 px-3 font-mono tabular-nums text-slate-500">
                    {idx + 1}
                  </td>
                  {columns.map((colKey) => (
                    <td key={colKey} className="py-2 px-3 text-slate-800">
                      {colKey === 'status' ? (
                        <StatusText status={row.status} />
                      ) : (
                        <span
                          className={
                            colKey === 'mobile' ||
                            colKey === 'enquiryNo' ||
                            colKey.endsWith('Date')
                              ? 'font-mono tabular-nums whitespace-nowrap'
                              : colKey === 'remarks'
                              ? 'leading-relaxed'
                              : 'whitespace-nowrap'
                          }
                        >
                          {row[colKey] || '-'}
                        </span>
                      )}
                    </td>
                  ))}
                </tr>
              ))
            )}
          </tbody>
        </table>
      </div>
    </div>
  );
}
