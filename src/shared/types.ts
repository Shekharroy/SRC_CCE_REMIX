export type CanonicalCounterName =
  | 'Chandrahiya'
  | 'Raja Bazar'
  | 'Priemia'
  | 'Rupdih'
  | 'Mehsi'
  | 'Rishabh Kumar'
  | 'Siswa'
  | 'Unassigned';

export type NormalizedStatus =
  | 'Open'
  | 'Follow-up'
  | 'Hot'
  | 'Warm'
  | 'Cold'
  | 'Booked'
  | 'Converted'
  | 'Lost';

export interface CounterDefinition {
  id: number;
  name: Exclude<CanonicalCounterName, 'Unassigned'>;
  dseNames: string[];
}

export interface EnquiryRow {
  id: string;
  sourceDocId?: string;
  sourceFileName?: string;
  enquiryNo: string;
  enquiryDate: string; // YYYY-MM-DD or raw
  customerName: string;
  mobile: string;
  dseName: string;
  counterName: CanonicalCounterName;
  principleModel: string;
  statusRaw: string;
  status: NormalizedStatus;
  followupDate: string; // YYYY-MM-DD or ''
  calledDate: string; // YYYY-MM-DD or ''
  conversionDate: string; // YYYY-MM-DD or ''
  crmName: string;
  remarks: string;
  isConverted: boolean;
}

export interface DseFollowupGroup {
  dseName: string;
  counterName: CanonicalCounterName;
  totalFollowups: number;
  convertedCount: number;
  hotCount: number;
  rows: EnquiryRow[];
}

export interface DseFollowupReport {
  targetDate: string;
  totalFollowups: number;
  dseCount: number;
  groups: DseFollowupGroup[];
  flatRows: EnquiryRow[];
}

export interface PrincipleModelSummary {
  principleModel: string;
  totalEnquiries: number;
  openCount: number;
  followupCount: number;
  hotCount: number;
  warmCount: number;
  coldCount: number;
  bookedCount: number;
  convertedCount: number;
  lostCount: number;
  conversionRate: number;
  topCounter: string;
}

export interface PrincipleModelReport {
  totalModels: number;
  totalEnquiries: number;
  totalConverted: number;
  summaries: PrincipleModelSummary[];
  rows: EnquiryRow[];
}

export interface CrmDailyCallSummary {
  date: string;
  totalCalled: number;
  convertedCount: number;
  followupCount: number;
  hotCount: number;
  lostCount: number;
  remarksList: string[];
}

export interface CrmDailyCallReport {
  crmName: string;
  availableCrmNames: string[];
  totalCalls: number;
  totalConversions: number;
  conversionRate: number;
  dailySummaries: CrmDailyCallSummary[];
  callRows: EnquiryRow[];
  conversionList: EnquiryRow[];
}

export interface CounterDseBreakdown {
  dseName: string;
  totalEnquiries: number;
  openCount: number;
  followupCount: number;
  hotCount: number;
  warmCount: number;
  coldCount: number;
  bookedCount: number;
  convertedCount: number;
  lostCount: number;
  conversionRate: number;
}

export interface CounterSummary {
  counterNumber: number;
  counterName: CanonicalCounterName;
  assignedDseList: string[];
  totalEnquiries: number;
  openCount: number;
  followupCount: number;
  hotCount: number;
  warmCount: number;
  coldCount: number;
  bookedCount: number;
  convertedCount: number;
  lostCount: number;
  conversionRate: number;
  dseBreakdowns: CounterDseBreakdown[];
}

export interface CounterWiseReport {
  totalEnquiries: number;
  totalConverted: number;
  overallConversionRate: number;
  summaries: CounterSummary[];
}

export interface AllReportsBundle {
  generatedAt: string;
  targetDate: string;
  crmTarget: string;
  totalRecords: number;
  report1DseFollowup: DseFollowupReport;
  report2PrincipleModel: PrincipleModelReport;
  report3CrmDailyCall: CrmDailyCallReport;
  report4CounterWise: CounterWiseReport;
}

export type SourceFileKind = 'excel' | 'csv' | 'text' | 'image' | 'google-sheet';

export interface UploadedDocumentRecord {
  id: string;
  fileName: string;
  fileKind: SourceFileKind;
  mimeType: string;
  sizeBytes: number;
  uploadedAt: string;
  rowCount: number;
  extractionMethod: 'spreadsheet' | 'gemini-ai' | 'text-parser' | 'google-sheets';
  notes?: string;
  sourceUrl?: string;
  rows: EnquiryRow[];
  hasStoredFile: boolean;
}

export type CustomReportColumnKey =
  | 'enquiryNo'
  | 'enquiryDate'
  | 'customerName'
  | 'mobile'
  | 'counterName'
  | 'dseName'
  | 'principleModel'
  | 'status'
  | 'followupDate'
  | 'calledDate'
  | 'conversionDate'
  | 'crmName'
  | 'remarks';

export type CustomReportGroupBy =
  | 'none'
  | 'counterName'
  | 'dseName'
  | 'principleModel'
  | 'status'
  | 'crmName'
  | 'date';

export interface CustomReportConfig {
  title: string;
  counters: CanonicalCounterName[];
  dseNames: string[];
  statuses: NormalizedStatus[];
  principleModels: string[];
  crmName: string;
  dateFrom: string;
  dateTo: string;
  searchQuery: string;
  groupBy: CustomReportGroupBy;
  columns: CustomReportColumnKey[];
}

export interface CustomReportGroupRow {
  groupKey: string;
  total: number;
  openCount: number;
  followupCount: number;
  hotCount: number;
  convertedCount: number;
  lostCount: number;
  conversionRate: number;
}

export interface CustomReportResult {
  config: CustomReportConfig;
  totalMatched: number;
  convertedMatched: number;
  conversionRate: number;
  groups: CustomReportGroupRow[];
  rows: EnquiryRow[];
}

export interface GeneratedReportSnapshot {
  id: string;
  title: string;
  createdAt: string;
  targetDate: string;
  crmTarget: string;
  sourceDocIds: string[];
  sourceFileNames: string[];
  totalRecords: number;
  summaryMetrics: {
    totalEnquiries: number;
    todayFollowups: number;
    crmCalls: number;
    crmConversions: number;
    activeCounters: number;
    totalConverted: number;
  };
  bundle: AllReportsBundle;
}

export interface DatabaseStatusInfo {
  mode: 'mongodb' | 'local-fallback';
  databaseName: string;
  connected: boolean;
  connectionLabel: string;
  documentsCount: number;
  reportsCount: number;
  lastError?: string;
}
