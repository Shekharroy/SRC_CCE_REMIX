import assert from 'node:assert/strict';
import test from 'node:test';
import JSZip from 'jszip';
import * as XLSX from 'xlsx';
import { isValidRecordId } from '../server/db.ts';
import {
  buildCustomReportBuffer,
  buildSampleTemplateBuffer,
  buildWorkbookBuffer,
} from '../server/excelExport.ts';
import {
  extractGoogleSheetIds,
  normalizeRawRows,
  parseUploadedDocument,
} from '../server/parser.ts';
import {
  buildAllReports,
  buildCustomReport,
  COUNTER_DEFINITIONS,
  DEFAULT_CRM_NAME,
  filterEnquiryRows,
  getTodayIsoDate,
  normalizeCrmName,
  normalizeDateString,
  resolveDseAndCounter,
} from '../src/shared/reportEngine.ts';

test('Counter & DSE canonical resolution covers all 7 counters', () => {
  assert.equal(COUNTER_DEFINITIONS.length, 7);
  assert.deepEqual(resolveDseAndCounter('Pramod', ''), {
    dseName: 'Pramod',
    counterName: 'Chandrahiya',
  });
  assert.deepEqual(resolveDseAndCounter('Abhishek Kumar', ''), {
    dseName: 'Abhishekhar Kumar',
    counterName: 'Chandrahiya',
  });
  assert.deepEqual(resolveDseAndCounter('Niraj', ''), {
    dseName: 'Niraj',
    counterName: 'Raja Bazar',
  });
  assert.deepEqual(resolveDseAndCounter('Alok Das', ''), {
    dseName: 'Alok Das',
    counterName: 'Priemia',
  });
  assert.deepEqual(resolveDseAndCounter('Chhotu', ''), {
    dseName: 'Chhotu',
    counterName: 'Rupdih',
  });
  assert.deepEqual(resolveDseAndCounter('Jaiprakash', ''), {
    dseName: 'Jaiprakash',
    counterName: 'Mehsi',
  });
  assert.deepEqual(resolveDseAndCounter('Rishabh Kumar', ''), {
    dseName: 'Rishabh Kumar',
    counterName: 'Rishabh Kumar',
  });
  assert.deepEqual(resolveDseAndCounter('', 'Siswa'), {
    dseName: 'Siswa',
    counterName: 'Siswa',
  });
});

test('Web Service & Seibel Administrator caller rule defaults to Chandrashekhar Kumar unless another caller is mentioned', () => {
  assert.equal(normalizeCrmName('', 'Web Service'), DEFAULT_CRM_NAME);
  assert.equal(normalizeCrmName('Seibel Administrator', 'Seibel Administrator'), DEFAULT_CRM_NAME);
  assert.equal(normalizeCrmName('N/A', 'Web Service'), DEFAULT_CRM_NAME);
  assert.equal(normalizeCrmName('Rahul Verma', 'Web Service'), 'Rahul Verma');
  assert.equal(normalizeCrmName('Rahul Verma', 'Seibel Administrator'), 'Rahul Verma');
  assert.equal(normalizeCrmName('', 'Pramod'), '');
  assert.equal(normalizeCrmName('chandrashekhar', 'Pramod'), DEFAULT_CRM_NAME);
});

test('Date normalization handles ISO, Indian DD/MM/YYYY, and Excel serial dates', () => {
  assert.equal(normalizeDateString('2026-11-16'), '2026-11-16');
  assert.equal(normalizeDateString('16/11/2026'), '2026-11-16');
  assert.equal(normalizeDateString('16-11-2026'), '2026-11-16');
  assert.equal(normalizeDateString(46342), '2026-11-16');
});

test('All 4 reports + Custom Report Builder compute accurately on normalized rows', () => {
  const today = getTodayIsoDate();
  const rows = normalizeRawRows(
    [
      {
        enquiryNo: 'E1',
        enquiryDate: today,
        customerName: 'Asha Devi',
        mobile: '9835000001',
        dseName: 'Web Service',
        principleModel: 'Bolero Neo',
        statusRaw: 'Follow-up',
        followupDate: today,
        calledDate: today,
        crmName: '',
        remarks: 'Needs EMI call back',
      },
      {
        enquiryNo: 'E2',
        enquiryDate: today,
        customerName: 'Mohan Prasad',
        mobile: '9835000002',
        dseName: 'Niraj',
        principleModel: 'Scorpio-N',
        statusRaw: 'Converted',
        followupDate: today,
        calledDate: today,
        conversionDate: today,
        crmName: 'Chandrashekhar Kumar',
        remarks: 'Full payment received',
      },
      {
        enquiryNo: 'E3',
        enquiryDate: today,
        customerName: 'Sita Ram',
        mobile: '9835000003',
        dseName: 'Seibel Administrator',
        principleModel: 'Bolero Neo',
        statusRaw: 'Hot',
        followupDate: today,
        calledDate: today,
        crmName: 'Vikram',
        remarks: 'Assigned to Vikram caller explicitly',
      },
    ],
    'doc_test_1',
    'test.xlsx'
  );

  const bundle = buildAllReports(rows, { targetDate: today, crmTarget: DEFAULT_CRM_NAME });
  assert.equal(bundle.report1DseFollowup.totalFollowups, 3);
  assert.equal(bundle.report2PrincipleModel.totalModels, 2);
  // E1 (Web Service blank caller -> Chandrashekhar) + E2 (explicit Chandrashekhar); E3 omitted (Vikram)
  assert.equal(bundle.report3CrmDailyCall.totalCalls, 2);
  assert.equal(bundle.report3CrmDailyCall.totalConversions, 1);
  assert.equal(bundle.report4CounterWise.summaries.find((c) => c.counterName === 'Chandrahiya')?.totalEnquiries, 2);
  assert.equal(bundle.report4CounterWise.summaries.find((c) => c.counterName === 'Raja Bazar')?.totalEnquiries, 1);

  const custom = buildCustomReport(rows, {
    title: 'Chandrahiya Only',
    counters: ['Chandrahiya'],
    dseNames: [],
    statuses: [],
    principleModels: ['Bolero Neo'],
    crmName: 'ALL',
    dateFrom: '',
    dateTo: '',
    searchQuery: '',
    groupBy: 'dseName',
    columns: ['customerName', 'dseName', 'crmName', 'remarks'],
  });
  assert.equal(custom.totalMatched, 2);
  assert.equal(custom.groups.length, 2);
});

test('Security: ID validation rejects path traversal and NoSQL injection payloads', () => {
  assert.equal(isValidRecordId('doc_1794839234135_ceff86'), true);
  assert.equal(isValidRecordId('rep_1794839234135_abc123'), true);
  assert.equal(isValidRecordId('../etc/passwd'), false);
  assert.equal(isValidRecordId('doc_../../secret'), false);
  assert.equal(isValidRecordId('{"$ne": ""}'), false);
  assert.equal(isValidRecordId(''), false);
});

test('Security: Google Sheets URL parser blocks SSRF and non-docs.google.com hosts', () => {
  const valid = extractGoogleSheetIds(
    'https://docs.google.com/spreadsheets/d/1BxiMVs0XRA5nFMdKvBdBZjgmUUqptlbs74OgvE2upms/edit#gid=12345'
  );
  assert.equal(valid.spreadsheetId, '1BxiMVs0XRA5nFMdKvBdBZjgmUUqptlbs74OgvE2upms');
  assert.equal(valid.gid, '12345');

  assert.throws(() => extractGoogleSheetIds('http://127.0.0.1:3000/api/health'));
  assert.throws(() => extractGoogleSheetIds('https://evil.com/spreadsheets/d/123456789012345'));
  assert.throws(() => extractGoogleSheetIds('javascript:alert(1)'));
});

test('Security & Edge Cases: Parser rejects 0-byte, binary, corrupt xlsx, and unsupported extensions', async () => {
  await assert.rejects(() =>
    parseUploadedDocument({
      buffer: Buffer.alloc(0),
      fileName: 'empty.csv',
      mimeType: 'text/csv',
      docId: 'doc_test_empty',
    })
  );

  await assert.rejects(() =>
    parseUploadedDocument({
      buffer: Buffer.from('MZ\x00\x01\x02binary'),
      fileName: 'malware.exe',
      mimeType: 'application/octet-stream',
      docId: 'doc_test_exe',
    })
  );

  await assert.rejects(() =>
    parseUploadedDocument({
      buffer: Buffer.from([0x00, 0x01, 0x02, 0x03]),
      fileName: 'fake.csv',
      mimeType: 'text/csv',
      docId: 'doc_test_bin_csv',
    })
  );

  await assert.rejects(() =>
    parseUploadedDocument({
      buffer: Buffer.from('not-a-valid-zip-excel-file'),
      fileName: 'corrupt.xlsx',
      mimeType: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
      docId: 'doc_test_corrupt_xlsx',
    })
  );
});

test('Security & Styling: Excel export neutralizes formula injection and embeds colorful styles & OOXML charts', async () => {
  const rows = normalizeRawRows(
    [
      {
        enquiryNo: 'ENQ-SEC',
        customerName: '=CMD|" /C calc"!A0',
        dseName: 'Pramod',
        principleModel: 'Bolero',
        statusRaw: 'Open',
        remarks: '+1+1 @SUM(A1:A2)',
      },
    ],
    'doc_test_sec',
    'sec.csv'
  );
  const bundle = buildAllReports(rows);
  const xlsxBuf = await buildWorkbookBuffer(bundle, rows, 'all');
  const wb = XLSX.read(xlsxBuf, { type: 'buffer' });
  const sheet = wb.Sheets['Normalized Data'];
  const parsed = XLSX.utils.sheet_to_json<Record<string, string>>(sheet, { range: 3 });
  assert.equal(parsed[0]['Customer Name'].startsWith("'="), true);
  assert.equal(parsed[0]['Remarks'].startsWith("'+"), true);

  const zip = await JSZip.loadAsync(xlsxBuf);
  const chartFiles = Object.keys(zip.files).filter((f) => f.startsWith('xl/charts/chart') && f.endsWith('.xml'));
  const drawingFiles = Object.keys(zip.files).filter((f) => f.startsWith('xl/drawings/drawing') && f.endsWith('.xml'));
  assert.ok(chartFiles.length >= 4, `Expected at least 4 embedded charts, found ${chartFiles.length}`);
  assert.ok(drawingFiles.length >= 3, `Expected at least 3 drawings, found ${drawingFiles.length}`);
  const stylesXml = await zip.file('xl/styles.xml')?.async('string');
  assert.ok(stylesXml && stylesXml.includes('fgColor'), 'Expected colored cell fills in xl/styles.xml');

  const customBuf = await buildCustomReportBuffer(rows, {
    title: 'SecCheck',
    counters: [],
    dseNames: [],
    statuses: [],
    principleModels: [],
    crmName: 'ALL',
    dateFrom: '',
    dateTo: '',
    searchQuery: '',
    groupBy: 'dseName',
    columns: ['customerName', 'remarks'],
  });
  const cwb = XLSX.read(customBuf, { type: 'buffer' });
  const cRows = XLSX.utils.sheet_to_json<Record<string, string>>(cwb.Sheets['SecCheck'], { range: 3 });
  assert.equal(cRows[0]['Customer Name'].startsWith("'="), true);
  const cZip = await JSZip.loadAsync(customBuf);
  assert.ok(cZip.file('xl/charts/chart1.xml'), 'Expected custom grouped report to embed chart1.xml');
});

test('Live Server Integration: Health, Security Headers, Upload, Save Snapshot, Export, and Error Handling', async () => {
  const baseUrl = 'http://127.0.0.1:3000';

  const healthRes = await fetch(`${baseUrl}/api/health`);
  assert.equal(healthRes.status, 200);
  assert.equal(healthRes.headers.get('x-content-type-options'), 'nosniff');
  assert.equal(healthRes.headers.get('x-powered-by'), null);
  const healthJson = (await healthRes.json()) as { ok: boolean; database: { databaseName: string } };
  assert.equal(healthJson.ok, true);
  assert.equal(healthJson.database.databaseName, 'SRC_CCE');

  // Break test: path traversal ID on download & delete returns 400
  const badIdRes = await fetch(`${baseUrl}/api/documents/..%2F..%2Fetc%2Fpasswd/download`);
  assert.equal(badIdRes.status, 400);

  // Break test: SSRF URL on google sheet import returns 400
  const ssrfRes = await fetch(`${baseUrl}/api/import-google-sheet`, {
    method: 'POST',
    headers: { 'Content-Type': 'application/json' },
    body: JSON.stringify({ url: 'http://127.0.0.1:3000/api/health' }),
  });
  assert.equal(ssrfRes.status, 400);

  // Break test: malformed JSON payload returns 400 JSON (not HTML stack trace)
  const badJsonRes = await fetch(`${baseUrl}/api/ingest-text`, {
    method: 'POST',
    headers: { 'Content-Type': 'application/json' },
    body: '{invalid-json',
  });
  assert.equal(badJsonRes.status, 400);
  const badJsonBody = (await badJsonRes.json()) as { error: string };
  assert.ok(badJsonBody.error);

  // Full Happy Path: Upload sample Excel -> Save Snapshot -> Export All & Custom Excel -> Delete cleanup
  const sampleBuf = await buildSampleTemplateBuffer();
  const form = new FormData();
  form.append(
    'files',
    new Blob([new Uint8Array(sampleBuf)], {
      type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
    }),
    'live_test_sheet.xlsx'
  );

  const upRes = await fetch(`${baseUrl}/api/upload`, { method: 'POST', body: form });
  assert.equal(upRes.status, 200);
  const upData = (await upRes.json()) as {
    documents: Array<{ id: string; rowCount: number; rows: unknown[] }>;
  };
  assert.equal(upData.documents.length, 1);
  assert.equal(upData.documents[0].rowCount, 23);
  const docId = upData.documents[0].id;

  const saveRes = await fetch(`${baseUrl}/api/reports/save`, {
    method: 'POST',
    headers: { 'Content-Type': 'application/json' },
    body: JSON.stringify({
      title: 'Live Integration Test Snapshot',
      rows: upData.documents[0].rows,
    }),
  });
  assert.equal(saveRes.status, 200);
  const saveData = (await saveRes.json()) as { report: { id: string } };
  const repId = saveData.report.id;

  const expRes = await fetch(`${baseUrl}/api/export/excel`, {
    method: 'POST',
    headers: { 'Content-Type': 'application/json' },
    body: JSON.stringify({ reportId: repId, scope: 'all' }),
  });
  assert.equal(expRes.status, 200);
  const expBuf = Buffer.from(await expRes.arrayBuffer());
  const wb = XLSX.read(expBuf, { type: 'buffer' });
  assert.ok(wb.SheetNames.includes('Executive Summary'));
  assert.ok(wb.SheetNames.includes('1. DSE Today Followup List'));
  assert.ok(wb.SheetNames.includes('2. Principle Model Wise'));
  assert.ok(wb.SheetNames.includes('3A. CRM Daily Call Count'));
  assert.ok(wb.SheetNames.includes('4A. Counter Wise Summary'));
  assert.ok(wb.SheetNames.includes('LookupLists'));

  const byCounter = filterEnquiryRows(upData.documents[0].rows as Parameters<typeof filterEnquiryRows>[0], {
    counter: 'Chandrahiya',
    status: 'Hot',
  });
  assert.ok(byCounter.length >= 1);
  assert.ok(byCounter.every((r) => r.counterName === 'Chandrahiya' && r.status === 'Hot'));

  const expZip = await JSZip.loadAsync(expBuf);
  const expChartFiles = Object.keys(expZip.files).filter((f) => f.startsWith('xl/charts/chart') && f.endsWith('.xml'));
  assert.ok(expChartFiles.length >= 6, `Expected >= 6 charts in complete workbook, found ${expChartFiles.length}`);
  const chartContents = await Promise.all(
    expChartFiles.map((f) => expZip.file(f)!.async('string'))
  );
  assert.ok(chartContents.some((c) => c.includes('<c:barChart>')), 'Expected bar/column charts in export');
  assert.ok(chartContents.some((c) => c.includes('<c:pieChart>')), 'Expected pie chart in Counter Wise Summary');

  const execSheetXml = await expZip.file('xl/worksheets/sheet1.xml')!.async('string');
  assert.ok(execSheetXml.includes('<dataValidations'), 'Expected Data Validation dropdown slicers on Executive Summary');
  assert.ok(execSheetXml.includes('COUNTIFS('), 'Expected dynamic COUNTIFS formulas on Executive Summary');
  assert.ok(execSheetXml.includes('<hyperlinks>') && execSheetXml.includes('location="'), 'Expected internal drill-down hyperlinks');
  const dseSummaryXml = await expZip.file('xl/worksheets/sheet2.xml')!.async('string');
  assert.ok(dseSummaryXml.includes('<autoFilter'), 'Expected AutoFilter sorting/filtering on report sheets');
  assert.ok(dseSummaryXml.includes('<hyperlinks>'), 'Expected drill-down hyperlinks on DSE Followup Summary');

  const customExpRes = await fetch(`${baseUrl}/api/export/custom-excel`, {
    method: 'POST',
    headers: { 'Content-Type': 'application/json' },
    body: JSON.stringify({
      rows: upData.documents[0].rows,
      config: {
        title: 'Chandrahiya Custom',
        counters: ['Chandrahiya'],
        dseNames: [],
        statuses: [],
        principleModels: [],
        crmName: 'ALL',
        dateFrom: '',
        dateTo: '',
        searchQuery: '',
        groupBy: 'dseName',
        columns: ['customerName', 'dseName', 'crmName', 'status'],
      },
    }),
  });
  assert.equal(customExpRes.status, 200);
  const customExpZip = await JSZip.loadAsync(Buffer.from(await customExpRes.arrayBuffer()));
  assert.ok(customExpZip.file('xl/charts/chart1.xml'), 'Expected custom Excel download to include chart1.xml');

  // Clean up test records
  const delRep = await fetch(`${baseUrl}/api/reports/${repId}`, { method: 'DELETE' });
  assert.equal(delRep.status, 200);
  const delDoc = await fetch(`${baseUrl}/api/documents/${docId}`, { method: 'DELETE' });
  assert.equal(delDoc.status, 200);
});
