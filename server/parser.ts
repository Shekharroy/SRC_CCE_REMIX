import crypto from 'crypto';
import * as XLSX from 'xlsx';
import { GoogleGenAI, Type } from '@google/genai';
import {
  DEFAULT_CRM_NAME,
  normalizeCrmName,
  normalizeDateString,
  normalizeStatus,
  resolveDseAndCounter,
} from '../src/shared/reportEngine.ts';
import type { EnquiryRow, SourceFileKind } from '../src/shared/types.ts';

function getAiClient(): GoogleGenAI {
  const apiKey = process.env.GEMINI_API_KEY;
  if (!apiKey || apiKey === 'MY_GEMINI_API_KEY') {
    throw new Error(
      'GEMINI_API_KEY is not configured. Please verify your API key in Settings > Secrets to extract records from images or unstructured text.'
    );
  }
  return new GoogleGenAI({
    apiKey,
    httpOptions: {
      headers: {
        'User-Agent': 'aistudio-build',
      },
    },
  });
}

function normalizeHeaderKey(header: string): string {
  return String(header || '')
    .toLowerCase()
    .replace(/[^a-z0-9]/g, '');
}

const COLUMN_SYNONYMS: Record<keyof RawExtractedRow, string[]> = {
  enquiryNo: ['enquiryno', 'enquiryid', 'enqno', 'leadid', 'optyid', 'opportunityid', 'srno', 'slno', 'id', 'refno'],
  enquiryDate: ['enquirydate', 'enqdate', 'date', 'createddate', 'leaddate', 'optydate'],
  customerName: ['customername', 'prospectname', 'clientname', 'partyname', 'customer', 'name', 'buyername'],
  mobile: ['mobile', 'mobileno', 'phone', 'phoneno', 'contact', 'contactno', 'cell'],
  dseName: ['dsename', 'dse', 'salesexecutive', 'executive', 'executivename', 'salesofficer', 'advisor', 'assignedto', 'consultant'],
  counterName: ['countername', 'counter', 'location', 'branch', 'showroom', 'outlet', 'center', 'centre'],
  principleModel: ['principlemodel', 'principalmodel', 'model', 'vehiclemodel', 'modelname', 'product', 'variant', 'carmodel', 'bikemodel', 'tractor'],
  statusRaw: ['status', 'enquirystatus', 'leadstatus', 'stage', 'disposition', 'callstatus', 'category', 'bucket'],
  followupDate: ['followupdate', 'nextfollowupdate', 'followup', 'nextfollowup', 'todayfollowup', 'followdate', 'scheduleddate'],
  calledDate: ['calleddate', 'calldate', 'lastcalled', 'callingdate', 'dialeddate'],
  conversionDate: ['conversiondate', 'converteddate', 'bookingdate', 'retaildate', 'invoicedate', 'saledate', 'deliverydate'],
  crmName: ['crmname', 'crm', 'caller', 'callername', 'telecaller', 'cce', 'ccename', 'cre', 'crename'],
  remarks: ['remarks', 'remark', 'feedback', 'comment', 'comments', 'note', 'notes', 'discussion', 'summary'],
};

interface RawExtractedRow {
  enquiryNo?: string;
  enquiryDate?: string;
  customerName?: string;
  mobile?: string;
  dseName?: string;
  counterName?: string;
  principleModel?: string;
  statusRaw?: string;
  followupDate?: string;
  calledDate?: string;
  conversionDate?: string;
  crmName?: string;
  remarks?: string;
}

function mapObjectToRawRow(record: Record<string, unknown>): RawExtractedRow {
  const mapped: RawExtractedRow = {};
  const entries = Object.entries(record);

  for (const [key, value] of entries) {
    if (value === null || value === undefined || String(value).trim() === '') continue;
    const norm = normalizeHeaderKey(key);
    if (!norm) continue;

    let matchedField: keyof RawExtractedRow | null = null;
    for (const [field, aliases] of Object.entries(COLUMN_SYNONYMS) as Array<
      [keyof RawExtractedRow, string[]]
    >) {
      if (aliases.includes(norm)) {
        matchedField = field;
        break;
      }
    }

    if (!matchedField) {
      for (const [field, aliases] of Object.entries(COLUMN_SYNONYMS) as Array<
        [keyof RawExtractedRow, string[]]
      >) {
        if (aliases.some((alias) => norm.includes(alias) && alias.length >= 4)) {
          matchedField = field;
          break;
        }
      }
    }

    if (matchedField && !mapped[matchedField]) {
      mapped[matchedField] = String(value).trim();
    }
  }

  return mapped;
}

export function normalizeRawRows(
  rawRows: RawExtractedRow[],
  docId: string,
  fileName: string
): EnquiryRow[] {
  const normalized: EnquiryRow[] = [];

  rawRows.forEach((raw, idx) => {
    const hasMeaningfulContent = Boolean(
      raw.customerName ||
        raw.dseName ||
        raw.principleModel ||
        raw.counterName ||
        raw.remarks ||
        raw.enquiryNo ||
        raw.mobile
    );
    if (!hasMeaningfulContent) return;

    const { dseName, counterName } = resolveDseAndCounter(
      raw.dseName || '',
      raw.counterName || ''
    );
    const { status, isConverted } = normalizeStatus(raw.statusRaw || '', raw.remarks || '');

    const enquiryDate = normalizeDateString(raw.enquiryDate || '');
    const followupDate = normalizeDateString(raw.followupDate || '');
    const calledDate = normalizeDateString(raw.calledDate || '');
    let conversionDate = normalizeDateString(raw.conversionDate || '');
    if (isConverted && !conversionDate) {
      conversionDate = calledDate || followupDate || enquiryDate;
    }

    normalized.push({
      id: `${docId}-row-${idx + 1}-${crypto.randomBytes(2).toString('hex')}`,
      sourceDocId: docId,
      sourceFileName: fileName,
      enquiryNo: (raw.enquiryNo || '').trim() || `ENQ-${idx + 1}`,
      enquiryDate,
      customerName: (raw.customerName || '').trim() || 'Walk-in / Prospect',
      mobile: (raw.mobile || '').trim(),
      dseName,
      counterName,
      principleModel: (raw.principleModel || '').trim() || 'Unspecified Model',
      statusRaw: (raw.statusRaw || '').trim() || status,
      status,
      followupDate,
      calledDate,
      conversionDate,
      crmName: normalizeCrmName(raw.crmName || '', dseName),
      remarks: (raw.remarks || '').trim(),
      isConverted,
    });
  });

  return normalized;
}

function findHeaderRowIndex(matrix: unknown[][]): number {
  const allSynonyms = new Set(Object.values(COLUMN_SYNONYMS).flat());
  let bestIdx = 0;
  let bestScore = -1;

  const limit = Math.min(matrix.length, 15);
  for (let i = 0; i < limit; i++) {
    const row = matrix[i] || [];
    let score = 0;
    for (const cell of row) {
      const norm = normalizeHeaderKey(String(cell ?? ''));
      if (norm && allSynonyms.has(norm)) {
        score += 2;
      } else if (norm && Array.from(allSynonyms).some((s) => s.length >= 4 && norm.includes(s))) {
        score += 1;
      }
    }
    if (score > bestScore) {
      bestScore = score;
      bestIdx = i;
    }
  }
  return bestScore > 0 ? bestIdx : 0;
}

export function parseSpreadsheetBuffer(
  buffer: Buffer,
  docId: string,
  fileName: string
): EnquiryRow[] {
  const workbook = XLSX.read(buffer, { type: 'buffer', cellDates: false });
  const allRaw: RawExtractedRow[] = [];

  for (const sheetName of workbook.SheetNames) {
    const sheet = workbook.Sheets[sheetName];
    if (!sheet) continue;

    const matrix = XLSX.utils.sheet_to_json<unknown[]>(sheet, {
      header: 1,
      defval: '',
      blankrows: false,
    });
    if (matrix.length < 2) continue;

    const headerIdx = findHeaderRowIndex(matrix);
    const headerRow = (matrix[headerIdx] || []).map((cell, cIdx) => {
      const txt = String(cell ?? '').trim();
      return txt || `Column_${cIdx + 1}`;
    });

    for (let r = headerIdx + 1; r < matrix.length; r++) {
      const values = matrix[r] || [];
      const rowObj: Record<string, unknown> = {};
      let nonEmpty = false;
      headerRow.forEach((colName, cIdx) => {
        const val = values[cIdx];
        if (val !== undefined && val !== null && String(val).trim() !== '') {
          nonEmpty = true;
        }
        rowObj[colName] = val;
      });
      if (!nonEmpty) continue;
      allRaw.push(mapObjectToRawRow(rowObj));
    }
  }

  return normalizeRawRows(allRaw, docId, fileName);
}

const EXTRACTION_SCHEMA = {
  type: Type.ARRAY,
  items: {
    type: Type.OBJECT,
    properties: {
      enquiryNo: { type: Type.STRING, description: 'Enquiry or lead ID number if present' },
      enquiryDate: { type: Type.STRING, description: 'Enquiry date in YYYY-MM-DD format if available' },
      customerName: { type: Type.STRING, description: 'Customer or prospect full name' },
      mobile: { type: Type.STRING, description: 'Customer phone/mobile number' },
      dseName: { type: Type.STRING, description: 'DSE / Sales Executive name' },
      counterName: {
        type: Type.STRING,
        description:
          'Counter / showroom location (Chandrahiya, Raja Bazar, Priemia, Rupdih, Mehsi, Rishabh Kumar, Siswa)',
      },
      principleModel: { type: Type.STRING, description: 'Vehicle / Principle Model name' },
      statusRaw: {
        type: Type.STRING,
        description: 'Status such as Open, Follow-up, Hot, Warm, Cold, Booked, Converted, Lost',
      },
      followupDate: { type: Type.STRING, description: 'Followup date in YYYY-MM-DD format' },
      calledDate: { type: Type.STRING, description: 'CRM call date in YYYY-MM-DD format' },
      conversionDate: { type: Type.STRING, description: 'Conversion / retail / booking date in YYYY-MM-DD' },
      crmName: {
        type: Type.STRING,
        description: `CRM / Caller name (e.g. ${DEFAULT_CRM_NAME})`,
      },
      remarks: { type: Type.STRING, description: 'Call remarks, customer feedback, or notes' },
    },
  },
};

const EXTRACTION_INSTRUCTION = `You are a precision data extraction engine for an automotive dealership CRM system (SRC_CCE).
Extract every enquiry, follow-up, call log, or lead conversion record from the provided document.
Known Counters and DSEs:
1. Chandrahiya: Pramod, Mani, Monu, Chandani, Sagar Kumar, Kumar Sagar, Awnish Kumar, Abhishekhar Kumar, Anand Kumar, Dev Kumar, Seibel Administrator, Web Service
2. Raja Bazar: Niraj, Anuj
3. Priemia: Alok Das, Anurag
4. Rupdih: Ravi, Manish, Chhotu
5. Mehsi: Jaiprakash, Ranjan
6. Rishabh Kumar
7. Siswa
Known CRM Caller: Chandrashekhar Kumar.
Rule: For rows where DSE / source is "Web Service" or "Seibel Administrator", if no explicit different caller name is mentioned, set crmName to "Chandrashekhar Kumar"; if an explicit caller name is mentioned, preserve that mentioned caller name.
Return a clean JSON array matching the schema. Normalize dates to YYYY-MM-DD where possible. Do not invent rows that are not present.`;

async function extractWithGemini(
  parts: Array<{ text: string } | { inlineData: { mimeType: string; data: string } }>
): Promise<RawExtractedRow[]> {
  const ai = getAiClient();
  const response = await ai.models.generateContent({
    model: 'gemini-4-hydrogen',
    contents: { parts },
    config: {
      systemInstruction: EXTRACTION_INSTRUCTION,
      responseMimeType: 'application/json',
      responseSchema: EXTRACTION_SCHEMA,
    },
  });

  const text = (response.text ?? '[]').trim();
  const parsed = JSON.parse(text);
  if (!Array.isArray(parsed)) return [];
  return parsed as RawExtractedRow[];
}

const ALLOWED_EXTENSIONS = new Set([
  'xlsx',
  'xls',
  'csv',
  'txt',
  'tsv',
  'log',
  'png',
  'jpg',
  'jpeg',
  'webp',
]);

export async function parseUploadedDocument(params: {
  buffer: Buffer;
  fileName: string;
  mimeType: string;
  docId: string;
}): Promise<{
  fileKind: SourceFileKind;
  extractionMethod: 'spreadsheet' | 'gemini-ai' | 'text-parser';
  rows: EnquiryRow[];
}> {
  const { buffer, fileName, mimeType, docId } = params;
  if (!buffer || buffer.length === 0) {
    throw new Error('Uploaded file is empty (0 bytes).');
  }

  const ext = fileName.split('.').pop()?.toLowerCase() || '';
  if (ext && !ALLOWED_EXTENSIONS.has(ext) && !mimeType.startsWith('image/') && !mimeType.startsWith('text/')) {
    throw new Error(`Unsupported file format ".${ext}". Supported formats: .xlsx, .xls, .csv, .txt, .tsv, .png, .jpg, .webp.`);
  }

  // 1. Excel (.xlsx, .xls)
  if (ext === 'xlsx' || ext === 'xls' || mimeType.includes('spreadsheet') || mimeType.includes('excel')) {
    try {
      const rows = parseSpreadsheetBuffer(buffer, docId, fileName);
      if (rows.length === 0) {
        throw new Error('Spreadsheet parsed successfully, but no recognizable enquiry rows were found.');
      }
      return { fileKind: 'excel', extractionMethod: 'spreadsheet', rows };
    } catch (err) {
      throw new Error(
        err instanceof Error && err.message.includes('no recognizable')
          ? err.message
          : 'Invalid or corrupted Excel workbook. Please re-save as .xlsx or .csv and upload again.'
      );
    }
  }

  // 2. CSV
  if (ext === 'csv' || mimeType === 'text/csv') {
    if (buffer.includes(0)) {
      throw new Error('File appears to contain binary data rather than valid CSV text.');
    }
    const rows = parseSpreadsheetBuffer(buffer, docId, fileName);
    if (rows.length === 0) {
      throw new Error('CSV file contained no valid enquiry rows.');
    }
    return { fileKind: 'csv', extractionMethod: 'spreadsheet', rows };
  }

  // 3. Images (png, jpg, jpeg, webp)
  if (
    mimeType.startsWith('image/') ||
    ['png', 'jpg', 'jpeg', 'webp'].includes(ext)
  ) {
    const resolvedMime = mimeType.startsWith('image/')
      ? mimeType
      : `image/${ext === 'jpg' ? 'jpeg' : ext}`;
    const rawRows = await extractWithGemini([
      {
        inlineData: {
          mimeType: resolvedMime,
          data: buffer.toString('base64'),
        },
      },
      {
        text: 'Extract all dealership enquiry, DSE follow-up, principle model, CRM call count/remarks, and conversion rows visible in this image.',
      },
    ]);
    const rows = normalizeRawRows(rawRows, docId, fileName);
    return { fileKind: 'image', extractionMethod: 'gemini-ai', rows };
  }

  // 4. Text (.txt, .tsv, .log, etc.)
  if (buffer.includes(0)) {
    throw new Error('File contains binary data and cannot be parsed as text.');
  }
  const textContent = buffer.toString('utf-8').trim();
  if (!textContent) {
    throw new Error('Text file is empty.');
  }

  // Try structured TSV/CSV/delimited parsing first if multiple lines contain tabs or pipes or commas
  const lines = textContent.split(/\r?\n/).filter((l) => l.trim().length > 0);
  if (lines.length >= 2 && (lines[0].includes('\t') || lines[0].includes('|') || lines[0].split(',').length >= 4)) {
    const normalizedText = lines[0].includes('|')
      ? lines.map((l) => l.split('|').map((c) => c.trim()).join('\t')).join('\n')
      : textContent;
    const rows = parseSpreadsheetBuffer(Buffer.from(normalizedText, 'utf-8'), docId, fileName);
    if (rows.length > 0) {
      return { fileKind: 'text', extractionMethod: 'text-parser', rows };
    }
  }

  // Unstructured text -> Gemini structured extraction
  const rawRows = await extractWithGemini([
    {
      text: `Extract all enquiry, follow-up, CRM call log, and conversion records from this text:\n\n${textContent}`,
    },
  ]);
  const rows = normalizeRawRows(rawRows, docId, fileName);
  return { fileKind: 'text', extractionMethod: 'gemini-ai', rows };
}

export function extractGoogleSheetIds(rawUrl: string): {
  spreadsheetId: string;
  gid?: string;
} {
  let parsed: URL;
  try {
    parsed = new URL(rawUrl.trim());
  } catch {
    throw new Error('Invalid URL format. Please provide a valid Google Sheets URL.');
  }

  if (parsed.protocol !== 'https:' || parsed.hostname !== 'docs.google.com') {
    throw new Error('Only https://docs.google.com/spreadsheets/... links are allowed.');
  }

  const match = parsed.pathname.match(/^\/spreadsheets\/d\/([a-zA-Z0-9-_]{10,120})/);
  if (!match) {
    throw new Error('Could not locate a Google Spreadsheet ID in the provided link.');
  }

  const gidFromQuery = parsed.searchParams.get('gid');
  const gidFromHash = parsed.hash.match(/gid=(\d+)/)?.[1];
  const gid = gidFromQuery || gidFromHash;

  return {
    spreadsheetId: match[1],
    gid: gid && /^\d+$/.test(gid) ? gid : undefined,
  };
}

export async function importPublicGoogleSheet(params: {
  url: string;
  customLabel?: string;
  docId: string;
}): Promise<{
  fileName: string;
  buffer: Buffer;
  rows: EnquiryRow[];
}> {
  const { spreadsheetId, gid } = extractGoogleSheetIds(params.url);
  const exportUrl = `https://docs.google.com/spreadsheets/d/${spreadsheetId}/export?format=csv${
    gid ? `&gid=${gid}` : ''
  }`;

  const response = await fetch(exportUrl, {
    method: 'GET',
    redirect: 'follow',
    signal: AbortSignal.timeout(12000),
  });

  if (response.status === 401 || response.status === 403) {
    throw new Error(
      'This Google Sheet is private or restricted. Share it as "Anyone with the link can view", or connect your Google Sheets account once enabled.'
    );
  }
  if (!response.ok) {
    throw new Error(`Unable to fetch Google Sheet (HTTP ${response.status}).`);
  }

  const contentType = response.headers.get('content-type') || '';
  const csvText = await response.text();
  if (contentType.includes('text/html') || csvText.trimStart().startsWith('<!DOCTYPE html') || csvText.trimStart().startsWith('<html')) {
    throw new Error(
      'This Google Sheet requires Google sign-in (returned login page). Set sharing to "Anyone with the link can view" or connect your Google account.'
    );
  }

  const buffer = Buffer.from(csvText, 'utf-8');
  const fileName =
    (params.customLabel || '').trim() ||
    `GoogleSheet_${spreadsheetId.slice(0, 8)}${gid ? `_gid_${gid}` : ''}.csv`;

  const rows = parseSpreadsheetBuffer(buffer, params.docId, fileName);
  if (rows.length === 0) {
    throw new Error('The linked Google Sheet tab was empty or had no recognizable CRM/enquiry columns.');
  }

  return { fileName, buffer, rows };
}
