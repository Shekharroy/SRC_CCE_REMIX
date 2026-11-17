import ExcelJS from 'exceljs';
import {
  ALL_NORMALIZED_STATUSES,
  buildCustomReport,
  COUNTER_DEFINITIONS,
  CUSTOM_REPORT_COLUMNS,
  getTodayIsoDate,
} from '../src/shared/reportEngine.ts';
import type {
  AllReportsBundle,
  CustomReportConfig,
  EnquiryRow,
} from '../src/shared/types.ts';
import {
  embedNativeChartsInXlsx,
  sheetRangeRef,
  type EmbeddedChartSpec,
  type InternalHyperlinkSpec,
} from './ooxmlCharts.ts';

interface SheetTheme {
  tabColor: string;
  headerBg: string;
  bannerBg: string;
  subtitleBg: string;
}

const THEMES: Record<string, SheetTheme> = {
  executive: {
    tabColor: '1D4ED8',
    headerBg: '1D4ED8',
    bannerBg: '0F172A',
    subtitleBg: 'E0E7FF',
  },
  dse: {
    tabColor: '0D9488',
    headerBg: '0F766E',
    bannerBg: '134E4A',
    subtitleBg: 'CCFBF1',
  },
  model: {
    tabColor: '4338CA',
    headerBg: '4338CA',
    bannerBg: '1E1B4B',
    subtitleBg: 'E0E7FF',
  },
  crm: {
    tabColor: '059669',
    headerBg: '047857',
    bannerBg: '064E3B',
    subtitleBg: 'D1FAE5',
  },
  counter: {
    tabColor: 'D97706',
    headerBg: 'B45309',
    bannerBg: '451A03',
    subtitleBg: 'FEF3C7',
  },
  raw: {
    tabColor: '475569',
    headerBg: '334155',
    bannerBg: '0F172A',
    subtitleBg: 'F1F5F9',
  },
};

function sanitizeExcelCell(val: unknown): unknown {
  if (typeof val !== 'string') return val;
  if (/^[=+\-@\t\r]/.test(val)) {
    if (/^-\d+(\.\d+)?$/.test(val.trim())) return val;
    return `'${val}`;
  }
  return val;
}

function getStatusStyle(statusText: string): { bg: string; fg: string } | null {
  const s = String(statusText || '').toLowerCase().trim();
  if (s === 'converted' || s === 'booked' || s === 'converted / retailed') {
    return { bg: 'DCFCE7', fg: '166534' };
  }
  if (s === 'hot') {
    return { bg: 'FFEDD5', fg: '9A3412' };
  }
  if (s === 'follow-up') {
    return { bg: 'DBEAFE', fg: '1E40AF' };
  }
  if (s === 'warm') {
    return { bg: 'FEF9C3', fg: '854D0E' };
  }
  if (s === 'lost') {
    return { bg: 'FEE2E2', fg: '991B1B' };
  }
  if (s === 'open' || s === 'cold') {
    return { bg: 'F1F5F9', fg: '475569' };
  }
  return null;
}

interface StyledTableResult {
  sheetName: string;
  ws: ExcelJS.Worksheet;
  startDataRow: number;
  endDataRow: number;
}

function addDrillDownLink(
  hyperlinks: InternalHyperlinkSpec[],
  sourceWs: ExcelJS.Worksheet,
  sourceSheetName: string,
  col1Based: number,
  row1Based: number,
  targetSheetName: string,
  targetCell: string,
  tooltip: string
): void {
  const cell = sourceWs.getRow(row1Based).getCell(col1Based);
  const existingFont = cell.font || { name: 'Calibri', size: 10 };
  cell.font = {
    ...existingFont,
    color: { argb: 'FF1D4ED8' },
    underline: true,
    bold: existingFont.bold ?? true,
  };
  const safeTarget = `'${targetSheetName.replace(/'/g, "''")}'!${targetCell}`;
  hyperlinks.push({
    sheetName: sourceSheetName,
    cellRef: cell.address,
    location: safeTarget,
    display: String(cell.value ?? ''),
    tooltip,
  });
}

function addStyledSheet(params: {
  wb: ExcelJS.Workbook;
  sheetName: string;
  title: string;
  subtitle: string;
  theme: SheetTheme;
  headers: string[];
  rows: Array<Record<string, unknown>>;
  totalsRow?: Record<string, unknown>;
  statusColumnKeys?: string[];
  numericColumnKeys?: string[];
  enableAutoFilter?: boolean;
  backToExecutive?: { hyperlinks: InternalHyperlinkSpec[] };
}): StyledTableResult {
  const {
    wb,
    sheetName,
    title,
    subtitle,
    theme,
    headers,
    rows,
    totalsRow,
    statusColumnKeys = ['Status', 'Stage'],
    numericColumnKeys = [],
    enableAutoFilter = true,
    backToExecutive,
  } = params;

  const safeName = sheetName.replace(/[:\\/?*\[\]]/g, '').slice(0, 31) || 'Sheet';
  const ws = wb.addWorksheet(safeName, {
    properties: { tabColor: { argb: `FF${theme.tabColor}` } },
    views: [{ state: 'frozen', ySplit: 4 }],
  });

  const colCount = Math.max(headers.length, 1);

  // Row 1: Title Banner
  ws.mergeCells(1, 1, 1, colCount);
  const titleCell = ws.getCell(1, 1);
  titleCell.value = title;
  titleCell.font = { name: 'Calibri', size: 13, bold: true, color: { argb: 'FFFFFFFF' } };
  titleCell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: `FF${theme.bannerBg}` } };
  titleCell.alignment = { vertical: 'middle', horizontal: 'left', indent: 1 };
  ws.getRow(1).height = 28;

  // Row 2: Subtitle Banner
  ws.mergeCells(2, 1, 2, colCount);
  const subCell = ws.getCell(2, 1);
  subCell.value = subtitle;
  subCell.font = { name: 'Calibri', size: 10, italic: true, color: { argb: 'FF1E293B' } };
  subCell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: `FF${theme.subtitleBg}` } };
  subCell.alignment = { vertical: 'middle', horizontal: 'left', indent: 1 };
  ws.getRow(2).height = 20;

  // Row 3: Optional Back Navigation Link or Spacer
  if (backToExecutive) {
    ws.getRow(3).height = 18;
    const backCell = ws.getCell(3, 1);
    backCell.value = '<- Back to Executive Summary';
    backCell.font = { name: 'Calibri', size: 9, bold: true, underline: true, color: { argb: 'FF1D4ED8' } };
    backCell.alignment = { vertical: 'middle', horizontal: 'left' };
    backToExecutive.hyperlinks.push({
      sheetName: safeName,
      cellRef: 'A3',
      location: "'Executive Summary'!A1",
      display: '<- Back to Executive Summary',
      tooltip: 'Return to Executive Summary & Interactive Slicers',
    });
  } else {
    ws.getRow(3).height = 8;
  }

  // Row 4: Table Headers
  const headerRow = ws.getRow(4);
  headerRow.height = 24;
  headers.forEach((h, idx) => {
    const cell = headerRow.getCell(idx + 1);
    cell.value = h;
    cell.font = { name: 'Calibri', size: 10, bold: true, color: { argb: 'FFFFFFFF' } };
    cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: `FF${theme.headerBg}` } };
    cell.alignment = {
      vertical: 'middle',
      horizontal: numericColumnKeys.includes(h) ? 'right' : 'left',
      wrapText: true,
    };
    cell.border = {
      top: { style: 'thin', color: { argb: 'FFCBD5E1' } },
      bottom: { style: 'medium', color: { argb: `FF${theme.bannerBg}` } },
      left: { style: 'thin', color: { argb: 'FFCBD5E1' } },
      right: { style: 'thin', color: { argb: 'FFCBD5E1' } },
    };
  });

  const effectiveRows =
    rows.length > 0
      ? rows
      : [
          Object.fromEntries(
            headers.map((h, i) => [h, i === 0 ? 'No records matching criteria' : ''])
          ),
        ];

  const startDataRow = 5;
  effectiveRows.forEach((rowObj, rIdx) => {
    const excelRow = ws.getRow(startDataRow + rIdx);
    excelRow.height = 20;
    const isEven = rIdx % 2 === 1;
    const rowBg = isEven ? 'FFF8FAFC' : 'FFFFFFFF';

    headers.forEach((h, cIdx) => {
      const rawVal = rowObj[h];
      const safeVal = sanitizeExcelCell(rawVal);
      const cell = excelRow.getCell(cIdx + 1);
      cell.value = (safeVal === undefined || safeVal === null ? '' : safeVal) as ExcelJS.CellValue;

      const isNumeric = typeof safeVal === 'number' || numericColumnKeys.includes(h);
      const statusColors = statusColumnKeys.includes(h)
        ? getStatusStyle(String(safeVal ?? ''))
        : null;

      cell.font = {
        name: 'Calibri',
        size: 10,
        bold: Boolean(statusColors) || h === 'Total Enquiries' || h === 'Total',
        color: { argb: statusColors ? `FF${statusColors.fg}` : 'FF0F172A' },
      };
      cell.fill = {
        type: 'pattern',
        pattern: 'solid',
        fgColor: { argb: statusColors ? `FF${statusColors.bg}` : rowBg },
      };
      cell.alignment = {
        vertical: 'middle',
        horizontal: isNumeric ? 'right' : 'left',
        wrapText: h.toLowerCase().includes('remark') || h.toLowerCase().includes('detail'),
      };
      cell.border = {
        top: { style: 'thin', color: { argb: 'FFE2E8F0' } },
        bottom: { style: 'thin', color: { argb: 'FFE2E8F0' } },
        left: { style: 'thin', color: { argb: 'FFE2E8F0' } },
        right: { style: 'thin', color: { argb: 'FFE2E8F0' } },
      };
    });
  });

  const endDataRow = startDataRow + effectiveRows.length - 1;

  if (totalsRow && rows.length > 0) {
    const tRow = ws.getRow(endDataRow + 1);
    tRow.height = 22;
    headers.forEach((h, cIdx) => {
      const val = sanitizeExcelCell(totalsRow[h] ?? '');
      const cell = tRow.getCell(cIdx + 1);
      cell.value = val as ExcelJS.CellValue;
      const isNumeric = typeof val === 'number' || numericColumnKeys.includes(h);
      cell.font = { name: 'Calibri', size: 10, bold: true, color: { argb: 'FF0F172A' } };
      cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFE0E7FF' } };
      cell.alignment = { vertical: 'middle', horizontal: isNumeric ? 'right' : 'left' };
      cell.border = {
        top: { style: 'thin', color: { argb: 'FF64748B' } },
        bottom: { style: 'double', color: { argb: 'FF0F172A' } },
        left: { style: 'thin', color: { argb: 'FFCBD5E1' } },
        right: { style: 'thin', color: { argb: 'FFCBD5E1' } },
      };
    });
  }

  // Auto-fit column widths
  headers.forEach((h, cIdx) => {
    let maxLen = h.length;
    for (const r of effectiveRows) {
      const v = r[h] !== undefined && r[h] !== null ? String(r[h]) : '';
      if (v.length > maxLen) maxLen = v.length;
    }
    if (totalsRow && totalsRow[h] !== undefined) {
      const tv = String(totalsRow[h]);
      if (tv.length > maxLen) maxLen = tv.length;
    }
    ws.getColumn(cIdx + 1).width = Math.min(Math.max(maxLen + 4, 12), 46);
  });

  if (enableAutoFilter && rows.length > 0) {
    ws.autoFilter = {
      from: { row: 4, column: 1 },
      to: { row: endDataRow, column: colCount },
    };
  }

  return {
    sheetName: safeName,
    ws,
    startDataRow,
    endDataRow,
  };
}

export async function buildWorkbookBuffer(
  bundle: AllReportsBundle,
  rawRows: EnquiryRow[],
  scope: 'all' | 'report1' | 'report2' | 'report3' | 'report4' = 'all'
): Promise<Buffer> {
  const wb = new ExcelJS.Workbook();
  wb.creator = 'SRC CCE Analytics';
  wb.created = new Date();

  const charts: EmbeddedChartSpec[] = [];
  const hyperlinks: InternalHyperlinkSpec[] = [];
  const backLinkCfg = scope === 'all' ? { hyperlinks } : undefined;
  let execMeta: StyledTableResult | null = null;
  let r1Meta: StyledTableResult | null = null;
  let r1ListMeta: StyledTableResult | null = null;
  let r2Meta: StyledTableResult | null = null;
  let r3Meta: StyledTableResult | null = null;
  let r3CallMeta: StyledTableResult | null = null;
  let r4Meta: StyledTableResult | null = null;
  let r4MatrixMeta: StyledTableResult | null = null;
  let rawMeta: StyledTableResult | null = null;

  const firstRowByDseInR1 = new Map<string, number>();
  const firstRowByDateInR3B = new Map<string, number>();
  const firstRowByCounterInR4B = new Map<string, number>();
  const firstRowByModelInRaw = new Map<string, number>();
  const firstRowByDseInRaw = new Map<string, number>();

  if (scope === 'all') {
    const execHeaders = ['Section', 'Metric', 'Details'];
    const execRows = [
      {
        Section: 'Report Generated At',
        Metric: bundle.generatedAt.slice(0, 19).replace('T', ' '),
        Details: `Target Date: ${bundle.targetDate} | CRM Focus: ${bundle.crmTarget}`,
      },
      {
        Section: 'Total Normalized Enquiries',
        Metric: bundle.totalRecords,
        Details: 'Combined across uploaded Excel, CSV, Text, Image & Google Sheets sources',
      },
      {
        Section: '1. DSE Today Follow-up',
        Metric: bundle.report1DseFollowup.totalFollowups,
        Details: `${bundle.report1DseFollowup.dseCount} DSEs with scheduled follow-ups`,
      },
      {
        Section: '2. Principle Model Wise Enquiry',
        Metric: bundle.report2PrincipleModel.totalModels,
        Details: `${bundle.report2PrincipleModel.totalEnquiries} enquiries, ${bundle.report2PrincipleModel.totalConverted} converted/booked`,
      },
      {
        Section: `3. CRM Calls (${bundle.report3CrmDailyCall.crmName})`,
        Metric: bundle.report3CrmDailyCall.totalCalls,
        Details: `${bundle.report3CrmDailyCall.totalConversions} conversions (${bundle.report3CrmDailyCall.conversionRate}% conversion rate)`,
      },
      {
        Section: '4. Counter Wise Enquiry Summary',
        Metric: bundle.report4CounterWise.summaries.length,
        Details: `${bundle.report4CounterWise.totalEnquiries} total enquiries, ${bundle.report4CounterWise.overallConversionRate}% overall conversion`,
      },
    ];
    execMeta = addStyledSheet({
      wb,
      sheetName: 'Executive Summary',
      title: 'SRC CCE Analytics — Interactive Dealership Report Dashboard',
      subtitle: `Generated ${bundle.targetDate} · CRM Caller: ${bundle.crmTarget} · Total Enquiries: ${bundle.totalRecords} · Use Dropdowns in B14:B17 or Click Any Report Link`,
      theme: THEMES.executive,
      headers: execHeaders,
      rows: execRows,
      numericColumnKeys: ['Metric'],
      enableAutoFilter: false,
    });
  }

  // Report 1: DSE Name Wise Today Followup Data List with Remarks
  if (scope === 'all' || scope === 'report1') {
    const r1SumHeaders = [
      'Sl No',
      'DSE Name',
      'Counter Name',
      'Today Followup Count',
      'Hot Leads',
      'Converted',
    ];
    const r1Summary = bundle.report1DseFollowup.groups.map((g, i) => ({
      'Sl No': i + 1,
      'DSE Name': g.dseName,
      'Counter Name': g.counterName,
      'Today Followup Count': g.totalFollowups,
      'Hot Leads': g.hotCount,
      'Converted': g.convertedCount,
    }));
    const r1SumTotals = {
      'Sl No': '',
      'DSE Name': 'TOTAL',
      'Counter Name': `${bundle.report1DseFollowup.dseCount} DSEs`,
      'Today Followup Count': bundle.report1DseFollowup.totalFollowups,
      'Hot Leads': bundle.report1DseFollowup.groups.reduce((a, g) => a + g.hotCount, 0),
      'Converted': bundle.report1DseFollowup.groups.reduce((a, g) => a + g.convertedCount, 0),
    };
    r1Meta = addStyledSheet({
      wb,
      sheetName: '1A. DSE Followup Summary',
      title: 'First Report (Summary): DSE Name Wise Follow-up Count',
      subtitle: `Target Date: ${bundle.report1DseFollowup.targetDate} · ${bundle.report1DseFollowup.totalFollowups} follow-ups across ${bundle.report1DseFollowup.dseCount} DSEs · Click any DSE Name to drill down`,
      theme: THEMES.dse,
      headers: r1SumHeaders,
      rows: r1Summary,
      totalsRow: r1SumTotals,
      numericColumnKeys: ['Sl No', 'Today Followup Count', 'Hot Leads', 'Converted'],
      backToExecutive: backLinkCfg,
    });

    if (r1Summary.length > 0) {
      charts.push({
        sheetName: r1Meta.sheetName,
        title: 'DSE Wise Follow-up & Hot Leads',
        chartType: 'col',
        categoriesRef: sheetRangeRef(r1Meta.sheetName, 2, r1Meta.startDataRow, r1Meta.endDataRow),
        series: [
          {
            name: 'Follow-ups',
            valuesRef: sheetRangeRef(r1Meta.sheetName, 4, r1Meta.startDataRow, r1Meta.endDataRow),
            colorHex: '0D9488',
          },
          {
            name: 'Hot Leads',
            valuesRef: sheetRangeRef(r1Meta.sheetName, 5, r1Meta.startDataRow, r1Meta.endDataRow),
            colorHex: 'D97706',
          },
          {
            name: 'Converted',
            valuesRef: sheetRangeRef(r1Meta.sheetName, 6, r1Meta.startDataRow, r1Meta.endDataRow),
            colorHex: '16A34A',
          },
        ],
        anchor: {
          fromCol: 0,
          fromRow: r1Meta.endDataRow + 3,
          toCol: 7,
          toRow: r1Meta.endDataRow + 20,
        },
      });
    }

    const r1ListHeaders = [
      'Sl No',
      'DSE Name',
      'Counter Name',
      'Followup Date',
      'Enquiry No',
      'Enquiry Date',
      'Customer Name',
      'Mobile No',
      'Principle Model',
      'Status',
      'CRM / Caller',
      'Remarks',
    ];
    const r1List = bundle.report1DseFollowup.flatRows.map((row, idx) => {
      const targetRowNum = 5 + idx;
      if (row.dseName && !firstRowByDseInR1.has(row.dseName)) {
        firstRowByDseInR1.set(row.dseName, targetRowNum);
      }
      return {
        'Sl No': idx + 1,
        'DSE Name': row.dseName,
        'Counter Name': row.counterName,
        'Followup Date': row.followupDate || bundle.targetDate,
        'Enquiry No': row.enquiryNo,
        'Enquiry Date': row.enquiryDate,
        'Customer Name': row.customerName,
        'Mobile No': row.mobile,
        'Principle Model': row.principleModel,
        'Status': row.status,
        'CRM / Caller': row.crmName || '-',
        'Remarks': row.remarks || '-',
      };
    });
    r1ListMeta = addStyledSheet({
      wb,
      sheetName: '1. DSE Today Followup List',
      title: 'First Report: DSE Name Wise Today Follow-up Data List with Remarks',
      subtitle: `Target Date: ${bundle.report1DseFollowup.targetDate} · ${r1List.length} records · Filter & Sort enabled on all columns`,
      theme: THEMES.dse,
      headers: r1ListHeaders,
      rows: r1List,
      numericColumnKeys: ['Sl No'],
      backToExecutive: backLinkCfg,
    });

    if (r1Summary.length > 0 && r1Meta && r1ListMeta) {
      r1Summary.forEach((item, idx) => {
        const targetRow = firstRowByDseInR1.get(item['DSE Name']);
        if (targetRow) {
          addDrillDownLink(
            hyperlinks,
            r1Meta!.ws,
            r1Meta!.sheetName,
            2,
            r1Meta!.startDataRow + idx,
            r1ListMeta!.sheetName,
            `A${targetRow}`,
            `Drill down to ${item['DSE Name']} follow-up list`
          );
        }
      });
    }
  }

  // Report 2: Principle Model Wise Enquiry Report
  if (scope === 'all' || scope === 'report2') {
    const r2Headers = [
      'Sl No',
      'Principle Model',
      'Total Enquiries',
      'Open',
      'Follow-up',
      'Hot',
      'Warm',
      'Cold',
      'Booked',
      'Converted / Retailed',
      'Lost',
      'Conversion %',
      'Top Counter',
    ];
    const r2Rows = bundle.report2PrincipleModel.summaries.map((m, idx) => ({
      'Sl No': idx + 1,
      'Principle Model': m.principleModel,
      'Total Enquiries': m.totalEnquiries,
      'Open': m.openCount,
      'Follow-up': m.followupCount,
      'Hot': m.hotCount,
      'Warm': m.warmCount,
      'Cold': m.coldCount,
      'Booked': m.bookedCount,
      'Converted / Retailed': m.convertedCount,
      'Lost': m.lostCount,
      'Conversion %': `${m.conversionRate}%`,
      'Top Counter': m.topCounter,
    }));
    const r2Totals = {
      'Sl No': '',
      'Principle Model': 'TOTAL',
      'Total Enquiries': bundle.report2PrincipleModel.totalEnquiries,
      'Open': bundle.report2PrincipleModel.summaries.reduce((a, m) => a + m.openCount, 0),
      'Follow-up': bundle.report2PrincipleModel.summaries.reduce((a, m) => a + m.followupCount, 0),
      'Hot': bundle.report2PrincipleModel.summaries.reduce((a, m) => a + m.hotCount, 0),
      'Warm': bundle.report2PrincipleModel.summaries.reduce((a, m) => a + m.warmCount, 0),
      'Cold': bundle.report2PrincipleModel.summaries.reduce((a, m) => a + m.coldCount, 0),
      'Booked': bundle.report2PrincipleModel.summaries.reduce((a, m) => a + m.bookedCount, 0),
      'Converted / Retailed': bundle.report2PrincipleModel.totalConverted,
      'Lost': bundle.report2PrincipleModel.summaries.reduce((a, m) => a + m.lostCount, 0),
      'Conversion %':
        bundle.report2PrincipleModel.totalEnquiries > 0
          ? `${Number(
              (
                (bundle.report2PrincipleModel.totalConverted /
                  bundle.report2PrincipleModel.totalEnquiries) *
                100
              ).toFixed(1)
            )}%`
          : '0%',
      'Top Counter': '-',
    };
    r2Meta = addStyledSheet({
      wb,
      sheetName: '2. Principle Model Wise',
      title: 'Second Report: Principle Model Wise Enquiry & Status Summary',
      subtitle: `${bundle.report2PrincipleModel.totalModels} Models · ${bundle.report2PrincipleModel.totalEnquiries} Enquiries · ${bundle.report2PrincipleModel.totalConverted} Converted / Booked`,
      theme: THEMES.model,
      headers: r2Headers,
      rows: r2Rows,
      totalsRow: r2Totals,
      numericColumnKeys: [
        'Sl No',
        'Total Enquiries',
        'Open',
        'Follow-up',
        'Hot',
        'Warm',
        'Cold',
        'Booked',
        'Converted / Retailed',
        'Lost',
        'Conversion %',
      ],
      backToExecutive: backLinkCfg,
    });

    if (r2Rows.length > 0 && r2Meta) {
      charts.push({
        sheetName: r2Meta.sheetName,
        title: 'Principle Model Wise Enquiries vs Conversions',
        chartType: 'bar',
        categoriesRef: sheetRangeRef(r2Meta.sheetName, 2, r2Meta.startDataRow, r2Meta.endDataRow),
        series: [
          {
            name: 'Total Enquiries',
            valuesRef: sheetRangeRef(r2Meta.sheetName, 3, r2Meta.startDataRow, r2Meta.endDataRow),
            colorHex: '4338CA',
          },
          {
            name: 'Hot',
            valuesRef: sheetRangeRef(r2Meta.sheetName, 6, r2Meta.startDataRow, r2Meta.endDataRow),
            colorHex: 'D97706',
          },
          {
            name: 'Converted / Retailed',
            valuesRef: sheetRangeRef(r2Meta.sheetName, 10, r2Meta.startDataRow, r2Meta.endDataRow),
            colorHex: '16A34A',
          },
        ],
        anchor: {
          fromCol: 0,
          fromRow: r2Meta.endDataRow + 3,
          toCol: 9,
          toRow: r2Meta.endDataRow + 22,
        },
      });

      if (scope === 'all') {
        charts.push({
          sheetName: 'Executive Summary',
          title: 'Principle Model Wise Demand',
          chartType: 'bar',
          categoriesRef: sheetRangeRef(r2Meta.sheetName, 2, r2Meta.startDataRow, r2Meta.endDataRow),
          series: [
            {
              name: 'Total Enquiries',
              valuesRef: sheetRangeRef(r2Meta.sheetName, 3, r2Meta.startDataRow, r2Meta.endDataRow),
              colorHex: '4338CA',
            },
            {
              name: 'Converted',
              valuesRef: sheetRangeRef(r2Meta.sheetName, 10, r2Meta.startDataRow, r2Meta.endDataRow),
              colorHex: '16A34A',
            },
          ],
          anchor: {
            fromCol: 2,
            fromRow: 26,
            toCol: 6,
            toRow: 44,
          },
        });
      }
    }
  }

  // Report 3: CRM Chandrashekhar Kumar Daily Called Data Count with Remarks & Daily Lead Conversion List
  if (scope === 'all' || scope === 'report3') {
    const r3DailyHeaders = [
      'Sl No',
      'CRM Name',
      'Call Date',
      'Called Data Count',
      'Follow-up',
      'Hot Leads',
      'Converted Leads',
      'Lost',
      'Remarks Summary',
    ];
    const r3Daily = bundle.report3CrmDailyCall.dailySummaries.map((d, idx) => ({
      'Sl No': idx + 1,
      'CRM Name': bundle.report3CrmDailyCall.crmName,
      'Call Date': d.date,
      'Called Data Count': d.totalCalled,
      'Follow-up': d.followupCount,
      'Hot Leads': d.hotCount,
      'Converted Leads': d.convertedCount,
      'Lost': d.lostCount,
      'Remarks Summary': d.remarksList.join(' | ') || '-',
    }));
    const r3DailyTotals = {
      'Sl No': '',
      'CRM Name': 'TOTAL',
      'Call Date': `${bundle.report3CrmDailyCall.dailySummaries.length} Days`,
      'Called Data Count': bundle.report3CrmDailyCall.totalCalls,
      'Follow-up': bundle.report3CrmDailyCall.dailySummaries.reduce((a, d) => a + d.followupCount, 0),
      'Hot Leads': bundle.report3CrmDailyCall.dailySummaries.reduce((a, d) => a + d.hotCount, 0),
      'Converted Leads': bundle.report3CrmDailyCall.totalConversions,
      'Lost': bundle.report3CrmDailyCall.dailySummaries.reduce((a, d) => a + d.lostCount, 0),
      'Remarks Summary': `Conversion Rate: ${bundle.report3CrmDailyCall.conversionRate}%`,
    };
    r3Meta = addStyledSheet({
      wb,
      sheetName: '3A. CRM Daily Call Count',
      title: `Third Report (Part A): Daily Called Data Count — CRM ${bundle.report3CrmDailyCall.crmName}`,
      subtitle: `Total Calls: ${bundle.report3CrmDailyCall.totalCalls} · Conversions: ${bundle.report3CrmDailyCall.totalConversions} (${bundle.report3CrmDailyCall.conversionRate}%) · Click any Call Date to drill down`,
      theme: THEMES.crm,
      headers: r3DailyHeaders,
      rows: r3Daily,
      totalsRow: r3DailyTotals,
      numericColumnKeys: [
        'Sl No',
        'Called Data Count',
        'Follow-up',
        'Hot Leads',
        'Converted Leads',
        'Lost',
      ],
      backToExecutive: backLinkCfg,
    });

    if (r3Daily.length > 0) {
      charts.push({
        sheetName: r3Meta.sheetName,
        title: `Daily Call Count vs Conversions (${bundle.report3CrmDailyCall.crmName})`,
        chartType: 'col',
        categoriesRef: sheetRangeRef(r3Meta.sheetName, 3, r3Meta.startDataRow, r3Meta.endDataRow),
        series: [
          {
            name: 'Called Data Count',
            valuesRef: sheetRangeRef(r3Meta.sheetName, 4, r3Meta.startDataRow, r3Meta.endDataRow),
            colorHex: '1D4ED8',
          },
          {
            name: 'Hot Leads',
            valuesRef: sheetRangeRef(r3Meta.sheetName, 6, r3Meta.startDataRow, r3Meta.endDataRow),
            colorHex: 'D97706',
          },
          {
            name: 'Converted Leads',
            valuesRef: sheetRangeRef(r3Meta.sheetName, 7, r3Meta.startDataRow, r3Meta.endDataRow),
            colorHex: '059669',
          },
        ],
        anchor: {
          fromCol: 0,
          fromRow: r3Meta.endDataRow + 3,
          toCol: 8,
          toRow: r3Meta.endDataRow + 20,
        },
      });
    }

    const r3CallHeaders = [
      'Sl No',
      'CRM Name',
      'Call Date',
      'Customer Name',
      'Mobile No',
      'Counter Name',
      'DSE Name',
      'Principle Model',
      'Status',
      'Remarks',
    ];
    const r3CallDetail = bundle.report3CrmDailyCall.callRows.map((row, idx) => {
      const cDate = row.calledDate || row.followupDate || row.enquiryDate || 'Unspecified Date';
      if (!firstRowByDateInR3B.has(cDate)) {
        firstRowByDateInR3B.set(cDate, 5 + idx);
      }
      return {
        'Sl No': idx + 1,
        'CRM Name': row.crmName || bundle.report3CrmDailyCall.crmName,
        'Call Date': cDate,
        'Customer Name': row.customerName,
        'Mobile No': row.mobile,
        'Counter Name': row.counterName,
        'DSE Name': row.dseName,
        'Principle Model': row.principleModel,
        'Status': row.status,
        'Remarks': row.remarks || '-',
      };
    });
    r3CallMeta = addStyledSheet({
      wb,
      sheetName: '3B. CRM Call Remarks Log',
      title: `Third Report (Part B): CRM ${bundle.report3CrmDailyCall.crmName} Call Log with Remarks`,
      subtitle: `${r3CallDetail.length} individual call entries · AutoFilter & Column Sorting enabled`,
      theme: THEMES.crm,
      headers: r3CallHeaders,
      rows: r3CallDetail,
      numericColumnKeys: ['Sl No'],
      backToExecutive: backLinkCfg,
    });

    if (r3Daily.length > 0 && r3Meta && r3CallMeta) {
      r3Daily.forEach((d, idx) => {
        const targetRow = firstRowByDateInR3B.get(d['Call Date']);
        if (targetRow) {
          addDrillDownLink(
            hyperlinks,
            r3Meta!.ws,
            r3Meta!.sheetName,
            3,
            r3Meta!.startDataRow + idx,
            r3CallMeta!.sheetName,
            `A${targetRow}`,
            `Drill down to ${d['Call Date']} call log`
          );
        }
      });
    }

    const r3ConvHeaders = [
      'Sl No',
      'Conversion Date',
      'CRM Name',
      'Enquiry No',
      'Customer Name',
      'Mobile No',
      'Counter Name',
      'DSE Name',
      'Principle Model',
      'Status',
      'Remarks',
    ];
    const r3Conversions = bundle.report3CrmDailyCall.conversionList.map((row, idx) => ({
      'Sl No': idx + 1,
      'Conversion Date': row.conversionDate || row.calledDate || row.enquiryDate,
      'CRM Name': row.crmName || bundle.report3CrmDailyCall.crmName,
      'Enquiry No': row.enquiryNo,
      'Customer Name': row.customerName,
      'Mobile No': row.mobile,
      'Counter Name': row.counterName,
      'DSE Name': row.dseName,
      'Principle Model': row.principleModel,
      'Status': row.status,
      'Remarks': row.remarks || '-',
    }));
    addStyledSheet({
      wb,
      sheetName: '3C. Daily Lead Conversion List',
      title: 'Third Report (Part C): Daily Lead Conversion List',
      subtitle: `${r3Conversions.length} converted / booked leads · AutoFilter & Column Sorting enabled`,
      theme: THEMES.crm,
      headers: r3ConvHeaders,
      rows: r3Conversions,
      numericColumnKeys: ['Sl No'],
      backToExecutive: backLinkCfg,
    });
  }

  // Report 4: Counter Wise Enquiry Report with Status Summaries
  if (scope === 'all' || scope === 'report4') {
    const r4CounterHeaders = [
      'Counter No',
      'Counter Name',
      'Total Enquiries',
      'Open',
      'Follow-up',
      'Hot',
      'Warm',
      'Cold',
      'Booked',
      'Converted',
      'Lost',
      'Conversion %',
      'Assigned DSEs',
    ];
    const r4Counter = bundle.report4CounterWise.summaries.map((c) => ({
      'Counter No': c.counterNumber,
      'Counter Name': c.counterName,
      'Total Enquiries': c.totalEnquiries,
      'Open': c.openCount,
      'Follow-up': c.followupCount,
      'Hot': c.hotCount,
      'Warm': c.warmCount,
      'Cold': c.coldCount,
      'Booked': c.bookedCount,
      'Converted': c.convertedCount,
      'Lost': c.lostCount,
      'Conversion %': `${c.conversionRate}%`,
      'Assigned DSEs': c.assignedDseList.join(', '),
    }));
    const r4Totals = {
      'Counter No': '',
      'Counter Name': 'TOTAL',
      'Total Enquiries': bundle.report4CounterWise.totalEnquiries,
      'Open': bundle.report4CounterWise.summaries.reduce((a, c) => a + c.openCount, 0),
      'Follow-up': bundle.report4CounterWise.summaries.reduce((a, c) => a + c.followupCount, 0),
      'Hot': bundle.report4CounterWise.summaries.reduce((a, c) => a + c.hotCount, 0),
      'Warm': bundle.report4CounterWise.summaries.reduce((a, c) => a + c.warmCount, 0),
      'Cold': bundle.report4CounterWise.summaries.reduce((a, c) => a + c.coldCount, 0),
      'Booked': bundle.report4CounterWise.summaries.reduce((a, c) => a + c.bookedCount, 0),
      'Converted': bundle.report4CounterWise.summaries.reduce((a, c) => a + c.convertedCount, 0),
      'Lost': bundle.report4CounterWise.summaries.reduce((a, c) => a + c.lostCount, 0),
      'Conversion %': `${bundle.report4CounterWise.overallConversionRate}%`,
      'Assigned DSEs': '',
    };
    r4Meta = addStyledSheet({
      wb,
      sheetName: '4A. Counter Wise Summary',
      title: 'Fourth Report: Counter Wise Enquiry & Status Summary (All 7 Counters)',
      subtitle: `${bundle.report4CounterWise.totalEnquiries} Enquiries · ${bundle.report4CounterWise.totalConverted} Converted / Booked (${bundle.report4CounterWise.overallConversionRate}%) · Click any Counter Name to drill down`,
      theme: THEMES.counter,
      headers: r4CounterHeaders,
      rows: r4Counter,
      totalsRow: r4Totals,
      numericColumnKeys: [
        'Counter No',
        'Total Enquiries',
        'Open',
        'Follow-up',
        'Hot',
        'Warm',
        'Cold',
        'Booked',
        'Converted',
        'Lost',
        'Conversion %',
      ],
      backToExecutive: backLinkCfg,
    });

    if (r4Counter.length > 0) {
      charts.push({
        sheetName: r4Meta.sheetName,
        title: 'Counter Wise Enquiries vs Hot vs Converted',
        chartType: 'col',
        categoriesRef: sheetRangeRef(r4Meta.sheetName, 2, r4Meta.startDataRow, r4Meta.endDataRow),
        series: [
          {
            name: 'Total Enquiries',
            valuesRef: sheetRangeRef(r4Meta.sheetName, 3, r4Meta.startDataRow, r4Meta.endDataRow),
            colorHex: '1D4ED8',
          },
          {
            name: 'Hot',
            valuesRef: sheetRangeRef(r4Meta.sheetName, 6, r4Meta.startDataRow, r4Meta.endDataRow),
            colorHex: 'D97706',
          },
          {
            name: 'Converted',
            valuesRef: sheetRangeRef(r4Meta.sheetName, 10, r4Meta.startDataRow, r4Meta.endDataRow),
            colorHex: '16A34A',
          },
        ],
        anchor: {
          fromCol: 0,
          fromRow: r4Meta.endDataRow + 3,
          toCol: 6,
          toRow: r4Meta.endDataRow + 21,
        },
      });

      charts.push({
        sheetName: r4Meta.sheetName,
        title: 'Enquiry Share by Counter',
        chartType: 'pie',
        categoriesRef: sheetRangeRef(r4Meta.sheetName, 2, r4Meta.startDataRow, r4Meta.endDataRow),
        series: [
          {
            name: 'Total Enquiries',
            valuesRef: sheetRangeRef(r4Meta.sheetName, 3, r4Meta.startDataRow, r4Meta.endDataRow),
            colorHex: '1D4ED8',
          },
        ],
        anchor: {
          fromCol: 7,
          fromRow: r4Meta.endDataRow + 3,
          toCol: 12,
          toRow: r4Meta.endDataRow + 21,
        },
      });

      if (scope === 'all') {
        charts.push({
          sheetName: 'Executive Summary',
          title: 'Counter Wise Enquiries vs Conversions',
          chartType: 'col',
          categoriesRef: sheetRangeRef(r4Meta.sheetName, 2, r4Meta.startDataRow, r4Meta.endDataRow),
          series: [
            {
              name: 'Total Enquiries',
              valuesRef: sheetRangeRef(r4Meta.sheetName, 3, r4Meta.startDataRow, r4Meta.endDataRow),
              colorHex: '1D4ED8',
            },
            {
              name: 'Converted',
              valuesRef: sheetRangeRef(r4Meta.sheetName, 10, r4Meta.startDataRow, r4Meta.endDataRow),
              colorHex: '16A34A',
            },
          ],
          anchor: {
            fromCol: 0,
            fromRow: 26,
            toCol: 2,
            toRow: 44,
          },
        });
      }
    }

    const r4MatrixHeaders = [
      'Counter Name',
      'DSE Name',
      'Total Enquiries',
      'Open',
      'Follow-up',
      'Hot',
      'Warm',
      'Cold',
      'Booked',
      'Converted',
      'Lost',
      'Conversion %',
    ];
    const r4Matrix: Array<Record<string, unknown>> = [];
    for (const c of bundle.report4CounterWise.summaries) {
      for (const d of c.dseBreakdowns) {
        if (!firstRowByCounterInR4B.has(c.counterName)) {
          firstRowByCounterInR4B.set(c.counterName, 5 + r4Matrix.length);
        }
        r4Matrix.push({
          'Counter Name': c.counterName,
          'DSE Name': d.dseName,
          'Total Enquiries': d.totalEnquiries,
          'Open': d.openCount,
          'Follow-up': d.followupCount,
          'Hot': d.hotCount,
          'Warm': d.warmCount,
          'Cold': d.coldCount,
          'Booked': d.bookedCount,
          'Converted': d.convertedCount,
          'Lost': d.lostCount,
          'Conversion %': `${d.conversionRate}%`,
        });
      }
    }
    r4MatrixMeta = addStyledSheet({
      wb,
      sheetName: '4B. Counter DSE Breakdown',
      title: 'Fourth Report (Detail): Counter & DSE Performance Matrix',
      subtitle: `${r4Matrix.length} DSE rows across ${bundle.report4CounterWise.summaries.length} counters · Filter & Sort enabled on all columns`,
      theme: THEMES.counter,
      headers: r4MatrixHeaders,
      rows: r4Matrix,
      numericColumnKeys: [
        'Total Enquiries',
        'Open',
        'Follow-up',
        'Hot',
        'Warm',
        'Cold',
        'Booked',
        'Converted',
        'Lost',
        'Conversion %',
      ],
      backToExecutive: backLinkCfg,
    });

    if (r4Meta && r4MatrixMeta) {
      r4Counter.forEach((c, idx) => {
        const targetRow = firstRowByCounterInR4B.get(c['Counter Name']);
        if (targetRow) {
          addDrillDownLink(
            hyperlinks,
            r4Meta!.ws,
            r4Meta!.sheetName,
            2,
            r4Meta!.startDataRow + idx,
            r4MatrixMeta!.sheetName,
            `A${targetRow}`,
            `Drill down to ${c['Counter Name']} DSE breakdown`
          );
        }
      });
    }
  }

  if (scope === 'all' && rawRows.length > 0) {
    const rawHeaders = [
      'Sl No',
      'Source File',
      'Enquiry No',
      'Enquiry Date',
      'Customer Name',
      'Mobile',
      'Counter Name',
      'DSE Name',
      'Principle Model',
      'Status',
      'Followup Date',
      'Called Date',
      'Conversion Date',
      'CRM Name',
      'Remarks',
    ];
    const rawExport = rawRows.map((row, idx) => {
      const targetRow = 5 + idx;
      const modelKey = (row.principleModel || '').trim() || 'Unspecified Model';
      if (!firstRowByModelInRaw.has(modelKey)) {
        firstRowByModelInRaw.set(modelKey, targetRow);
      }
      if (row.dseName && !firstRowByDseInRaw.has(row.dseName)) {
        firstRowByDseInRaw.set(row.dseName, targetRow);
      }
      return {
        'Sl No': idx + 1,
        'Source File': row.sourceFileName || '-',
        'Enquiry No': row.enquiryNo,
        'Enquiry Date': row.enquiryDate,
        'Customer Name': row.customerName,
        'Mobile': row.mobile,
        'Counter Name': row.counterName,
        'DSE Name': row.dseName,
        'Principle Model': row.principleModel,
        'Status': row.status,
        'Followup Date': row.followupDate,
        'Called Date': row.calledDate,
        'Conversion Date': row.conversionDate,
        'CRM Name': row.crmName,
        'Remarks': row.remarks,
      };
    });
    rawMeta = addStyledSheet({
      wb,
      sheetName: 'Normalized Data',
      title: 'Normalized Source Data (SRC_CCE)',
      subtitle: `${rawExport.length} unified records · Use column header arrows to Filter by DSE, Model, Counter, or Status & Sort any column`,
      theme: THEMES.raw,
      headers: rawHeaders,
      rows: rawExport,
      numericColumnKeys: ['Sl No'],
      backToExecutive: backLinkCfg,
    });

    // Cross-sheet drill-down links into Normalized Data
    if (r2Meta) {
      bundle.report2PrincipleModel.summaries.forEach((m, idx) => {
        const targetRow = firstRowByModelInRaw.get(m.principleModel);
        if (targetRow) {
          addDrillDownLink(
            hyperlinks,
            r2Meta!.ws,
            r2Meta!.sheetName,
            2,
            r2Meta!.startDataRow + idx,
            rawMeta!.sheetName,
            `A${targetRow}`,
            `Drill down to ${m.principleModel} rows in Normalized Data`
          );
        }
      });
    }
    if (r4MatrixMeta) {
      let rowOffset = 0;
      for (const c of bundle.report4CounterWise.summaries) {
        for (const d of c.dseBreakdowns) {
          const targetRow = firstRowByDseInRaw.get(d.dseName);
          if (targetRow) {
            addDrillDownLink(
              hyperlinks,
              r4MatrixMeta.ws,
              r4MatrixMeta.sheetName,
              2,
              r4MatrixMeta.startDataRow + rowOffset,
              rawMeta.sheetName,
              `A${targetRow}`,
              `Drill down to ${d.dseName} enquiries in Normalized Data`
            );
          }
          rowOffset++;
        }
      }
    }
  }

  // Interactive Slicer Panel & Navigation Links on Executive Summary
  if (execMeta && scope === 'all') {
    const eWs = execMeta.ws;
    eWs.getColumn(1).width = 38;
    eWs.getColumn(2).width = 28;
    eWs.getColumn(3).width = 56;
    eWs.getColumn(4).width = 18;
    eWs.getColumn(5).width = 18;
    eWs.getColumn(6).width = 18;

    // Navigation hyperlinks from Executive Summary Section column (Rows 7-10)
    const sectionTargets: Array<{ row: number; targetSheet: string; tooltip: string }> = [
      { row: 7, targetSheet: '1A. DSE Followup Summary', tooltip: 'Open Report 1: DSE Today Follow-up' },
      { row: 8, targetSheet: '2. Principle Model Wise', tooltip: 'Open Report 2: Principle Model Wise Enquiry' },
      { row: 9, targetSheet: '3A. CRM Daily Call Count', tooltip: 'Open Report 3: CRM Daily Called Data Count' },
      { row: 10, targetSheet: '4A. Counter Wise Summary', tooltip: 'Open Report 4: Counter Wise Enquiry Summary' },
    ];
    for (const sLink of sectionTargets) {
      if (wb.getWorksheet(sLink.targetSheet)) {
        addDrillDownLink(
          hyperlinks,
          eWs,
          execMeta.sheetName,
          1,
          sLink.row,
          sLink.targetSheet,
          'A4',
          sLink.tooltip
        );
      }
    }

    if (rawMeta && rawRows.length > 0) {
      const lookupWs = wb.addWorksheet('LookupLists', { state: 'hidden' });
      const uniqueCounters = [
        'All',
        ...COUNTER_DEFINITIONS.map((c) => c.name),
        ...(rawRows.some((r) => r.counterName === 'Unassigned') ? ['Unassigned'] : []),
      ];
      const uniqueDses = [
        'All',
        ...Array.from(new Set(rawRows.map((r) => r.dseName).filter(Boolean))).sort(),
      ];
      const uniqueModels = [
        'All',
        ...Array.from(new Set(rawRows.map((r) => r.principleModel).filter(Boolean))).sort(),
      ];
      const uniqueStatuses = ['All', ...ALL_NORMALIZED_STATUSES];

      lookupWs.getCell(1, 1).value = 'Counters';
      lookupWs.getCell(1, 2).value = 'DSEs';
      lookupWs.getCell(1, 3).value = 'Models';
      lookupWs.getCell(1, 4).value = 'Statuses';

      uniqueCounters.forEach((v, i) => {
        lookupWs.getCell(i + 2, 1).value = v;
      });
      uniqueDses.forEach((v, i) => {
        lookupWs.getCell(i + 2, 2).value = v;
      });
      uniqueModels.forEach((v, i) => {
        lookupWs.getCell(i + 2, 3).value = v;
      });
      uniqueStatuses.forEach((v, i) => {
        lookupWs.getCell(i + 2, 4).value = v;
      });

      const rawEnd = rawMeta.endDataRow;
      eWs.mergeCells(13, 1, 13, 3);
      const slicerTitle = eWs.getCell(13, 1);
      slicerTitle.value = 'Interactive Slicer & Dynamic Filter (Select Dropdowns in Column B)';
      slicerTitle.font = { name: 'Calibri', size: 11, bold: true, color: { argb: 'FFFFFFFF' } };
      slicerTitle.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF1D4ED8' } };
      slicerTitle.alignment = { vertical: 'middle', horizontal: 'left', indent: 1 };
      eWs.getRow(13).height = 24;

      const filterRows: Array<{
        rowNum: number;
        label: string;
        formulaRange: string;
        hint: string;
      }> = [
        {
          rowNum: 14,
          label: 'Counter Name Filter',
          formulaRange: `=LookupLists!$A$2:$A$${uniqueCounters.length + 1}`,
          hint: 'Filter by Counter (Chandrahiya, Raja Bazar, etc.) or All',
        },
        {
          rowNum: 15,
          label: 'DSE Name Filter',
          formulaRange: `=LookupLists!$B$2:$B$${uniqueDses.length + 1}`,
          hint: 'Filter by DSE Name or All',
        },
        {
          rowNum: 16,
          label: 'Principle Model Filter',
          formulaRange: `=LookupLists!$C$2:$C$${uniqueModels.length + 1}`,
          hint: 'Filter by Principle Model or All',
        },
        {
          rowNum: 17,
          label: 'Report Status Filter',
          formulaRange: `=LookupLists!$D$2:$D$${uniqueStatuses.length + 1}`,
          hint: 'Filter by Status (Hot, Follow-up, Converted, etc.) or All',
        },
      ];

      for (const fr of filterRows) {
        const row = eWs.getRow(fr.rowNum);
        row.height = 21;
        const cLabel = row.getCell(1);
        cLabel.value = fr.label;
        cLabel.font = { name: 'Calibri', size: 10, bold: true, color: { argb: 'FF0F172A' } };
        cLabel.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFF1F5F9' } };
        cLabel.border = {
          top: { style: 'thin', color: { argb: 'FFCBD5E1' } },
          bottom: { style: 'thin', color: { argb: 'FFCBD5E1' } },
          left: { style: 'thin', color: { argb: 'FFCBD5E1' } },
          right: { style: 'thin', color: { argb: 'FFCBD5E1' } },
        };

        const cVal = row.getCell(2);
        cVal.value = 'All';
        cVal.font = { name: 'Calibri', size: 10, bold: true, color: { argb: 'FF1D4ED8' } };
        cVal.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFDBEAFE' } };
        cVal.alignment = { vertical: 'middle', horizontal: 'center' };
        cVal.border = {
          top: { style: 'medium', color: { argb: 'FF1D4ED8' } },
          bottom: { style: 'medium', color: { argb: 'FF1D4ED8' } },
          left: { style: 'medium', color: { argb: 'FF1D4ED8' } },
          right: { style: 'medium', color: { argb: 'FF1D4ED8' } },
        };
        cVal.dataValidation = {
          type: 'list',
          allowBlank: false,
          formulae: [fr.formulaRange],
          showErrorMessage: true,
          errorTitle: 'Invalid Selection',
          error: 'Please select a value from the dropdown list.',
        };

        const cHint = row.getCell(3);
        cHint.value = fr.hint;
        cHint.font = { name: 'Calibri', size: 9, italic: true, color: { argb: 'FF475569' } };
        cHint.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFFFFFFF' } };
        cHint.border = {
          top: { style: 'thin', color: { argb: 'FFE2E8F0' } },
          bottom: { style: 'thin', color: { argb: 'FFE2E8F0' } },
          left: { style: 'thin', color: { argb: 'FFE2E8F0' } },
          right: { style: 'thin', color: { argb: 'FFE2E8F0' } },
        };
      }

      eWs.mergeCells(19, 1, 19, 3);
      const calcHeader = eWs.getCell(19, 1);
      calcHeader.value = 'Dynamic Filtered Metrics (Recalculated Live via COUNTIFS)';
      calcHeader.font = { name: 'Calibri', size: 10, bold: true, color: { argb: 'FFFFFFFF' } };
      calcHeader.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF0F172A' } };
      calcHeader.alignment = { vertical: 'middle', horizontal: 'left', indent: 1 };
      eWs.getRow(19).height = 22;

      const baseCriteria = `'Normalized Data'!$G$5:$G$${rawEnd},IF($B$14="All","*",$B$14),'Normalized Data'!$H$5:$H$${rawEnd},IF($B$15="All","*",$B$15),'Normalized Data'!$I$5:$I$${rawEnd},IF($B$16="All","*",$B$16)`;
      const metricRows: Array<{
        rowNum: number;
        label: string;
        formula: string;
        numFmt?: string;
        note: string;
        bgHex?: string;
      }> = [
        {
          rowNum: 20,
          label: 'Filtered Total Enquiries',
          formula: `COUNTIFS(${baseCriteria},'Normalized Data'!$J$5:$J$${rawEnd},IF($B$17="All","*",$B$17))`,
          note: 'Total records matching B14:B17 filters',
          bgHex: 'FFE0E7FF',
        },
        {
          rowNum: 21,
          label: 'Filtered Hot Leads',
          formula: `IF(OR($B$17="All",$B$17="Hot"),COUNTIFS(${baseCriteria},'Normalized Data'!$J$5:$J$${rawEnd},"Hot"),0)`,
          note: 'Hot status enquiries matching active slicers',
          bgHex: 'FFFFEDD5',
        },
        {
          rowNum: 22,
          label: 'Filtered Follow-up Leads',
          formula: `IF(OR($B$17="All",$B$17="Follow-up"),COUNTIFS(${baseCriteria},'Normalized Data'!$J$5:$J$${rawEnd},"Follow-up"),0)`,
          note: 'Follow-up status enquiries matching active slicers',
          bgHex: 'FFDBEAFE',
        },
        {
          rowNum: 23,
          label: 'Filtered Converted / Booked',
          formula: `IF(OR($B$17="All",$B$17="Converted"),COUNTIFS(${baseCriteria},'Normalized Data'!$J$5:$J$${rawEnd},"Converted"),0)+IF(OR($B$17="All",$B$17="Booked"),COUNTIFS(${baseCriteria},'Normalized Data'!$J$5:$J$${rawEnd},"Booked"),0)`,
          note: 'Booked + Converted leads matching active slicers',
          bgHex: 'FFDCFCE7',
        },
        {
          rowNum: 24,
          label: 'Filtered Lost Leads',
          formula: `IF(OR($B$17="All",$B$17="Lost"),COUNTIFS(${baseCriteria},'Normalized Data'!$J$5:$J$${rawEnd},"Lost"),0)`,
          note: 'Lost status enquiries matching active slicers',
          bgHex: 'FFFEE2E2',
        },
        {
          rowNum: 25,
          label: 'Filtered Conversion %',
          formula: 'IF(B20>0,B23/B20,0)',
          numFmt: '0.0%',
          note: 'Converted & Booked share of Filtered Total',
          bgHex: 'FFF8FAFC',
        },
      ];

      for (const mr of metricRows) {
        const r = eWs.getRow(mr.rowNum);
        r.height = 20;
        const c1 = r.getCell(1);
        c1.value = mr.label;
        c1.font = { name: 'Calibri', size: 10, bold: true, color: { argb: 'FF0F172A' } };
        c1.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: mr.bgHex || 'FFFFFFFF' } };
        c1.border = {
          top: { style: 'thin', color: { argb: 'FFE2E8F0' } },
          bottom: { style: 'thin', color: { argb: 'FFE2E8F0' } },
          left: { style: 'thin', color: { argb: 'FFE2E8F0' } },
          right: { style: 'thin', color: { argb: 'FFE2E8F0' } },
        };

        const c2 = r.getCell(2);
        c2.value = { formula: mr.formula };
        c2.font = { name: 'Calibri', size: 10, bold: true, color: { argb: 'FF0F172A' } };
        c2.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: mr.bgHex || 'FFFFFFFF' } };
        c2.alignment = { vertical: 'middle', horizontal: 'right' };
        if (mr.numFmt) c2.numFmt = mr.numFmt;
        c2.border = {
          top: { style: 'thin', color: { argb: 'FFE2E8F0' } },
          bottom: { style: 'thin', color: { argb: 'FFE2E8F0' } },
          left: { style: 'thin', color: { argb: 'FFE2E8F0' } },
          right: { style: 'thin', color: { argb: 'FFE2E8F0' } },
        };

        const c3 = r.getCell(3);
        c3.value = mr.note;
        c3.font = { name: 'Calibri', size: 9, color: { argb: 'FF475569' } };
        c3.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFFFFFFF' } };
        c3.border = {
          top: { style: 'thin', color: { argb: 'FFE2E8F0' } },
          bottom: { style: 'thin', color: { argb: 'FFE2E8F0' } },
          left: { style: 'thin', color: { argb: 'FFE2E8F0' } },
          right: { style: 'thin', color: { argb: 'FFE2E8F0' } },
        };
      }
    }
  }

  const rawBuf = Buffer.from(await wb.xlsx.writeBuffer());
  return embedNativeChartsInXlsx(rawBuf, charts, hyperlinks);
}

export async function buildSampleTemplateBuffer(): Promise<Buffer> {
  const today = getTodayIsoDate();
  const sampleRows = [
    {
      'Enquiry No': 'ENQ-2026-101',
      'Enquiry Date': today,
      'Customer Name': 'Rajeshwar Prasad',
      'Mobile': '9835011223',
      'Counter Name': 'Chandrahiya',
      'DSE Name': 'Pramod',
      'Principle Model': 'Bolero Neo N10',
      'Status': 'Follow-up',
      'Followup Date': today,
      'Called Date': today,
      'Conversion Date': '',
      'CRM Name': 'Chandrashekhar Kumar',
      'Remarks': 'Customer requested finance quote for 5-year EMI; visit scheduled today evening',
    },
    {
      'Enquiry No': 'ENQ-2026-102',
      'Enquiry Date': today,
      'Customer Name': 'Sunita Devi',
      'Mobile': '9934221109',
      'Counter Name': 'Chandrahiya',
      'DSE Name': 'Mani',
      'Principle Model': 'Scorpio-N Z8',
      'Status': 'Converted',
      'Followup Date': today,
      'Called Date': today,
      'Conversion Date': today,
      'CRM Name': 'Chandrashekhar Kumar',
      'Remarks': 'Booking amount received; retail invoicing completed today',
    },
    {
      'Enquiry No': 'ENQ-2026-103',
      'Enquiry Date': today,
      'Customer Name': 'Vikash Tiwari',
      'Mobile': '9708114455',
      'Counter Name': '',
      'DSE Name': 'Monu',
      'Principle Model': 'XUV 3XO AX5',
      'Status': 'Hot',
      'Followup Date': today,
      'Called Date': today,
      'Conversion Date': '',
      'CRM Name': 'Chandrashekhar Kumar',
      'Remarks': 'Test drive completed; comparing Red vs White color delivery timeline',
    },
    {
      'Enquiry No': 'ENQ-2026-104',
      'Enquiry Date': today,
      'Customer Name': 'Amitabh Singh',
      'Mobile': '9431055667',
      'Counter Name': '',
      'DSE Name': 'Chandani',
      'Principle Model': 'Bolero Pickup FB 1.7',
      'Status': 'Booked',
      'Followup Date': today,
      'Called Date': today,
      'Conversion Date': today,
      'CRM Name': 'Chandrashekhar Kumar',
      'Remarks': 'Down payment Rs 50,000 deposited; delivery planned on Dhanteras',
    },
    {
      'Enquiry No': 'ENQ-2026-105',
      'Enquiry Date': today,
      'Customer Name': 'Deepak Jaiswal',
      'Mobile': '9123456780',
      'Counter Name': 'Chandrahiya',
      'DSE Name': 'Sagar Kumar',
      'Principle Model': 'Thar Roxx MX5',
      'Status': 'Warm',
      'Followup Date': today,
      'Called Date': '',
      'Conversion Date': '',
      'CRM Name': '',
      'Remarks': 'Asked for exchange valuation on old hatchback',
    },
    {
      'Enquiry No': 'ENQ-2026-106',
      'Enquiry Date': today,
      'Customer Name': 'Mukesh Kushwaha',
      'Mobile': '9801234512',
      'Counter Name': '',
      'DSE Name': 'Kumar Sagar',
      'Principle Model': 'Bolero B6 Opt',
      'Status': 'Follow-up',
      'Followup Date': today,
      'Called Date': today,
      'Conversion Date': '',
      'CRM Name': 'Chandrashekhar Kumar',
      'Remarks': 'Called customer; documents pending for Mahindra Finance',
    },
    {
      'Enquiry No': 'ENQ-2026-107',
      'Enquiry Date': today,
      'Customer Name': 'Pankaj Verma',
      'Mobile': '9955443322',
      'Counter Name': '',
      'DSE Name': 'Awnish Kumar',
      'Principle Model': 'Scorpio Classic S11',
      'Status': 'Hot',
      'Followup Date': today,
      'Called Date': today,
      'Conversion Date': '',
      'CRM Name': 'Chandrashekhar Kumar',
      'Remarks': 'Needs quotation on company GST billing',
    },
    {
      'Enquiry No': 'ENQ-2026-108',
      'Enquiry Date': today,
      'Customer Name': 'Ramesh Sah',
      'Mobile': '9835778899',
      'Counter Name': '',
      'DSE Name': 'Abhishekhar Kumar',
      'Principle Model': 'Bolero Pickup ExtraLong',
      'Status': 'Converted',
      'Followup Date': today,
      'Called Date': today,
      'Conversion Date': today,
      'CRM Name': 'Chandrashekhar Kumar',
      'Remarks': 'Full cash payment settled; RC process initiated',
    },
    {
      'Enquiry No': 'ENQ-2026-109',
      'Enquiry Date': today,
      'Customer Name': 'Nitin Patel',
      'Mobile': '9931122334',
      'Counter Name': '',
      'DSE Name': 'Anand Kumar',
      'Principle Model': 'XUV700 AX7L',
      'Status': 'Open',
      'Followup Date': today,
      'Called Date': '',
      'Conversion Date': '',
      'CRM Name': '',
      'Remarks': 'Walk-in showroom enquiry for 7-seater diesel automatic',
    },
    {
      'Enquiry No': 'ENQ-2026-110',
      'Enquiry Date': today,
      'Customer Name': 'Santosh Mahto',
      'Mobile': '9431887766',
      'Counter Name': '',
      'DSE Name': 'Dev Kumar',
      'Principle Model': 'Bolero Neo N8',
      'Status': 'Lost',
      'Followup Date': today,
      'Called Date': today,
      'Conversion Date': '',
      'CRM Name': 'Chandrashekhar Kumar',
      'Remarks': 'Purchased used vehicle locally due to budget constraint',
    },
    {
      'Enquiry No': 'ENQ-2026-111',
      'Enquiry Date': today,
      'Customer Name': 'Harishankar Jha',
      'Mobile': '9852001122',
      'Counter Name': 'Chandrahiya',
      'DSE Name': 'Seibel Administrator',
      'Principle Model': 'XUV 3XO MX2',
      'Status': 'Follow-up',
      'Followup Date': today,
      'Called Date': today,
      'Conversion Date': '',
      'CRM Name': '',
      'Remarks': 'Siebel portal lead assigned; retry at 4 PM',
    },
    {
      'Enquiry No': 'ENQ-2026-112',
      'Enquiry Date': today,
      'Customer Name': 'Kundan Mishra',
      'Mobile': '9934556611',
      'Counter Name': 'Chandrahiya',
      'DSE Name': 'Web Service',
      'Principle Model': 'Thar Roxx AX7',
      'Status': 'Hot',
      'Followup Date': today,
      'Called Date': today,
      'Conversion Date': '',
      'CRM Name': '',
      'Remarks': 'Website enquiry verified by CRM; visiting Chandrahiya showroom tomorrow',
    },
    {
      'Enquiry No': 'ENQ-2026-113',
      'Enquiry Date': today,
      'Customer Name': 'Birendra Yadav',
      'Mobile': '9771223344',
      'Counter Name': 'Raja Bazar',
      'DSE Name': 'Niraj',
      'Principle Model': 'Bolero Pickup FB 1.7',
      'Status': 'Converted',
      'Followup Date': today,
      'Called Date': today,
      'Conversion Date': today,
      'CRM Name': 'Chandrashekhar Kumar',
      'Remarks': 'Converted via Raja Bazar counter; commercial permit documents submitted',
    },
    {
      'Enquiry No': 'ENQ-2026-114',
      'Enquiry Date': today,
      'Customer Name': 'Arvind Pandey',
      'Mobile': '9905112299',
      'Counter Name': 'Raja Bazar',
      'DSE Name': 'Anuj',
      'Principle Model': 'Scorpio Classic S',
      'Status': 'Follow-up',
      'Followup Date': today,
      'Called Date': today,
      'Conversion Date': '',
      'CRM Name': 'Chandrashekhar Kumar',
      'Remarks': 'Discussed rural agri-loan discount scheme',
    },
    {
      'Enquiry No': 'ENQ-2026-115',
      'Enquiry Date': today,
      'Customer Name': 'Sanjay Prasad',
      'Mobile': '9835990011',
      'Counter Name': 'Priemia',
      'DSE Name': 'Alok Das',
      'Principle Model': 'XUV700 AX5',
      'Status': 'Booked',
      'Followup Date': today,
      'Called Date': today,
      'Conversion Date': today,
      'CRM Name': 'Chandrashekhar Kumar',
      'Remarks': 'Priority booking confirmed at Priemia counter',
    },
    {
      'Enquiry No': 'ENQ-2026-116',
      'Enquiry Date': today,
      'Customer Name': 'Manoj Dubey',
      'Mobile': '9431229988',
      'Counter Name': 'Priemia',
      'DSE Name': 'Anurag',
      'Principle Model': 'XUV 3XO AX5',
      'Status': 'Cold',
      'Followup Date': today,
      'Called Date': '',
      'Conversion Date': '',
      'CRM Name': '',
      'Remarks': 'Purchase deferred to next month',
    },
    {
      'Enquiry No': 'ENQ-2026-117',
      'Enquiry Date': today,
      'Customer Name': 'Rakesh Thakur',
      'Mobile': '9852334455',
      'Counter Name': 'Rupdih',
      'DSE Name': 'Ravi',
      'Principle Model': 'Bolero B6',
      'Status': 'Follow-up',
      'Followup Date': today,
      'Called Date': today,
      'Conversion Date': '',
      'CRM Name': 'Chandrashekhar Kumar',
      'Remarks': 'Customer confirmed home visit by Ravi for KYC pickup',
    },
    {
      'Enquiry No': 'ENQ-2026-118',
      'Enquiry Date': today,
      'Customer Name': 'Ashok Kumar',
      'Mobile': '9931445566',
      'Counter Name': 'Rupdih',
      'DSE Name': 'Manish',
      'Principle Model': 'Bolero Neo N10',
      'Status': 'Converted',
      'Followup Date': today,
      'Called Date': today,
      'Conversion Date': today,
      'CRM Name': 'Chandrashekhar Kumar',
      'Remarks': 'Loan sanctioned by SBI; vehicle delivered from Rupdih',
    },
    {
      'Enquiry No': 'ENQ-2026-119',
      'Enquiry Date': today,
      'Customer Name': 'Umesh Ray',
      'Mobile': '9708667788',
      'Counter Name': 'Rupdih',
      'DSE Name': 'Chhotu',
      'Principle Model': 'Bolero Pickup ExtraLong',
      'Status': 'Warm',
      'Followup Date': today,
      'Called Date': '',
      'Conversion Date': '',
      'CRM Name': '',
      'Remarks': 'Comparing payload options',
    },
    {
      'Enquiry No': 'ENQ-2026-120',
      'Enquiry Date': today,
      'Customer Name': 'Prakash Gupta',
      'Mobile': '9955112233',
      'Counter Name': 'Mehsi',
      'DSE Name': 'Jaiprakash',
      'Principle Model': 'Scorpio-N Z6',
      'Status': 'Hot',
      'Followup Date': today,
      'Called Date': today,
      'Conversion Date': '',
      'CRM Name': 'Chandrashekhar Kumar',
      'Remarks': 'High interest; final price negotiation at Mehsi counter',
    },
    {
      'Enquiry No': 'ENQ-2026-121',
      'Enquiry Date': today,
      'Customer Name': 'Dinesh Paswan',
      'Mobile': '9835443322',
      'Counter Name': 'Mehsi',
      'DSE Name': 'Ranjan',
      'Principle Model': 'Bolero B6 Opt',
      'Status': 'Converted',
      'Followup Date': today,
      'Called Date': today,
      'Conversion Date': today,
      'CRM Name': 'Chandrashekhar Kumar',
      'Remarks': 'Converted lead; accessories package added',
    },
    {
      'Enquiry No': 'ENQ-2026-122',
      'Enquiry Date': today,
      'Customer Name': 'Suresh Choudhary',
      'Mobile': '9431665544',
      'Counter Name': 'Rishabh Kumar',
      'DSE Name': 'Rishabh Kumar',
      'Principle Model': 'Bolero Pickup FB 1.7',
      'Status': 'Follow-up',
      'Followup Date': today,
      'Called Date': today,
      'Conversion Date': '',
      'CRM Name': 'Chandrashekhar Kumar',
      'Remarks': 'Requested follow-up call after bank approval',
    },
    {
      'Enquiry No': 'ENQ-2026-123',
      'Enquiry Date': today,
      'Customer Name': 'Mahesh Sharma',
      'Mobile': '9934778800',
      'Counter Name': 'Siswa',
      'DSE Name': 'Siswa',
      'Principle Model': 'Bolero Neo N10',
      'Status': 'Hot',
      'Followup Date': today,
      'Called Date': today,
      'Conversion Date': '',
      'CRM Name': 'Chandrashekhar Kumar',
      'Remarks': 'Siswa counter prospect ready for booking this week',
    },
  ];

  const wb = new ExcelJS.Workbook();
  const headers = Object.keys(sampleRows[0]);
  addStyledSheet({
    wb,
    sheetName: 'SRC_CCE_Enquiries',
    title: 'SRC CCE Sample Dealership Enquiry Sheet',
    subtitle: `Sample dataset across all 7 Counters · Target Date: ${today}`,
    theme: THEMES.executive,
    headers,
    rows: sampleRows,
  });

  return Buffer.from(await wb.xlsx.writeBuffer());
}

export async function buildCustomReportBuffer(
  rawRows: EnquiryRow[],
  config: CustomReportConfig
): Promise<Buffer> {
  const result = buildCustomReport(rawRows, config);
  const wb = new ExcelJS.Workbook();
  wb.creator = 'SRC CCE Analytics';
  wb.created = new Date();
  const charts: EmbeddedChartSpec[] = [];
  const hyperlinks: InternalHyperlinkSpec[] = [];
  let sumMeta: StyledTableResult | null = null;

  if (result.groups.length > 0) {
    const sumHeaders = [
      'Sl No',
      'Group',
      'Total Records',
      'Open',
      'Follow-up',
      'Hot',
      'Converted / Booked',
      'Lost',
      'Conversion %',
    ];
    const groupRows = result.groups.map((g, idx) => ({
      'Sl No': idx + 1,
      'Group': g.groupKey,
      'Total Records': g.total,
      'Open': g.openCount,
      'Follow-up': g.followupCount,
      'Hot': g.hotCount,
      'Converted / Booked': g.convertedCount,
      'Lost': g.lostCount,
      'Conversion %': `${g.conversionRate}%`,
    }));
    sumMeta = addStyledSheet({
      wb,
      sheetName: 'Custom Summary',
      title: `${config.title || 'Custom Report'} — Grouped Summary`,
      subtitle: `Grouped by ${config.groupBy} · ${result.totalMatched} matched records · ${result.convertedMatched} converted (${result.conversionRate}%)`,
      theme: THEMES.model,
      headers: sumHeaders,
      rows: groupRows,
      numericColumnKeys: [
        'Sl No',
        'Total Records',
        'Open',
        'Follow-up',
        'Hot',
        'Converted / Booked',
        'Lost',
        'Conversion %',
      ],
    });

    charts.push({
      sheetName: sumMeta.sheetName,
      title: `${config.title || 'Custom Report'} Breakdown`,
      chartType: 'col',
      categoriesRef: sheetRangeRef(sumMeta.sheetName, 2, sumMeta.startDataRow, sumMeta.endDataRow),
      series: [
        {
          name: 'Total Records',
          valuesRef: sheetRangeRef(sumMeta.sheetName, 3, sumMeta.startDataRow, sumMeta.endDataRow),
          colorHex: '4338CA',
        },
        {
          name: 'Hot',
          valuesRef: sheetRangeRef(sumMeta.sheetName, 6, sumMeta.startDataRow, sumMeta.endDataRow),
          colorHex: 'D97706',
        },
        {
          name: 'Converted / Booked',
          valuesRef: sheetRangeRef(sumMeta.sheetName, 7, sumMeta.startDataRow, sumMeta.endDataRow),
          colorHex: '16A34A',
        },
      ],
      anchor: {
        fromCol: 0,
        fromRow: sumMeta.endDataRow + 3,
        toCol: 8,
        toRow: sumMeta.endDataRow + 21,
      },
    });
  }

  const colMetaMap = new Map(CUSTOM_REPORT_COLUMNS.map((c) => [c.key, c.label]));
  const selectedCols =
    config.columns.length > 0
      ? config.columns
      : CUSTOM_REPORT_COLUMNS.map((c) => c.key);

  const firstRowByGroupKey = new Map<string, number>();
  const detailHeaders = ['Sl No', ...selectedCols.map((k) => colMetaMap.get(k) || k)];
  const detailRows = result.rows.map((row, idx) => {
    let key = 'Unspecified';
    if (config.groupBy === 'counterName') key = row.counterName || 'Unassigned';
    else if (config.groupBy === 'dseName') key = row.dseName || 'Unassigned';
    else if (config.groupBy === 'principleModel') key = row.principleModel || 'Unspecified Model';
    else if (config.groupBy === 'status') key = row.status;
    else if (config.groupBy === 'crmName') key = row.crmName || 'Unassigned Caller';
    else if (config.groupBy === 'date') {
      key = row.followupDate || row.calledDate || row.enquiryDate || 'Unspecified Date';
    }
    if (!firstRowByGroupKey.has(key)) {
      firstRowByGroupKey.set(key, 5 + idx);
    }

    const out: Record<string, unknown> = { 'Sl No': idx + 1 };
    for (const colKey of selectedCols) {
      const label = colMetaMap.get(colKey) || colKey;
      out[label] = row[colKey] || '-';
    }
    return out;
  });

  let detailSheetName =
    (config.title || 'Custom Report').replace(/[:\\/?*\[\]]/g, '').slice(0, 31) || 'Custom Report';
  if (sumMeta && detailSheetName === sumMeta.sheetName) {
    detailSheetName = 'Custom Detail Rows';
  }

  const detailMeta = addStyledSheet({
    wb,
    sheetName: detailSheetName,
    title: config.title || 'Custom Dealership Report',
    subtitle: `${result.totalMatched} matched records · ${result.convertedMatched} converted (${result.conversionRate}%) · Filter & Sort enabled on all columns`,
    theme: THEMES.executive,
    headers: detailHeaders,
    rows: detailRows,
    numericColumnKeys: ['Sl No'],
  });

  if (sumMeta && result.groups.length > 0) {
    result.groups.forEach((g, idx) => {
      const targetRow = firstRowByGroupKey.get(g.groupKey);
      if (targetRow) {
        addDrillDownLink(
          hyperlinks,
          sumMeta!.ws,
          sumMeta!.sheetName,
          2,
          sumMeta!.startDataRow + idx,
          detailMeta.sheetName,
          `A${targetRow}`,
          `Drill down to ${g.groupKey} rows`
        );
      }
    });
  }

  const rawBuf = Buffer.from(await wb.xlsx.writeBuffer());
  return embedNativeChartsInXlsx(rawBuf, charts, hyperlinks);
}

