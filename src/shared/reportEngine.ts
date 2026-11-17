import type {
  AllReportsBundle,
  CanonicalCounterName,
  CounterDefinition,
  CounterDseBreakdown,
  CounterSummary,
  CounterWiseReport,
  CrmDailyCallReport,
  CrmDailyCallSummary,
  CustomReportColumnKey,
  CustomReportConfig,
  CustomReportGroupRow,
  CustomReportResult,
  DseFollowupGroup,
  DseFollowupReport,
  EnquiryRow,
  NormalizedStatus,
  PrincipleModelReport,
  PrincipleModelSummary,
} from './types.ts';

export const COUNTER_DEFINITIONS: CounterDefinition[] = [
  {
    id: 1,
    name: 'Chandrahiya',
    dseNames: [
      'Pramod',
      'Mani',
      'Monu',
      'Chandani',
      'Sagar Kumar',
      'Kumar Sagar',
      'Awnish Kumar',
      'Abhishekhar Kumar',
      'Anand Kumar',
      'Dev Kumar',
      'Seibel Administrator',
      'Web Service',
    ],
  },
  {
    id: 2,
    name: 'Raja Bazar',
    dseNames: ['Niraj', 'Anuj'],
  },
  {
    id: 3,
    name: 'Priemia',
    dseNames: ['Alok Das', 'Anurag'],
  },
  {
    id: 4,
    name: 'Rupdih',
    dseNames: ['Ravi', 'Manish', 'Chhotu'],
  },
  {
    id: 5,
    name: 'Mehsi',
    dseNames: ['Jaiprakash', 'Ranjan'],
  },
  {
    id: 6,
    name: 'Rishabh Kumar',
    dseNames: ['Rishabh Kumar'],
  },
  {
    id: 7,
    name: 'Siswa',
    dseNames: ['Siswa'],
  },
];

export const DEFAULT_CRM_NAME = 'Chandrashekhar Kumar';

const DSE_ALIAS_MAP: Record<string, { dse: string; counter: Exclude<CanonicalCounterName, 'Unassigned'> }> = {
  'pramod': { dse: 'Pramod', counter: 'Chandrahiya' },
  'pramod kumar': { dse: 'Pramod', counter: 'Chandrahiya' },
  'mani': { dse: 'Mani', counter: 'Chandrahiya' },
  'monu': { dse: 'Monu', counter: 'Chandrahiya' },
  'monu kumar': { dse: 'Monu', counter: 'Chandrahiya' },
  'chandani': { dse: 'Chandani', counter: 'Chandrahiya' },
  'chandni': { dse: 'Chandani', counter: 'Chandrahiya' },
  'sagar kumar': { dse: 'Sagar Kumar', counter: 'Chandrahiya' },
  'kumar sagar': { dse: 'Kumar Sagar', counter: 'Chandrahiya' },
  'awnish kumar': { dse: 'Awnish Kumar', counter: 'Chandrahiya' },
  'avnish kumar': { dse: 'Awnish Kumar', counter: 'Chandrahiya' },
  'awnish': { dse: 'Awnish Kumar', counter: 'Chandrahiya' },
  'abhishekhar kumar': { dse: 'Abhishekhar Kumar', counter: 'Chandrahiya' },
  'abhishekh kumar': { dse: 'Abhishekhar Kumar', counter: 'Chandrahiya' },
  'abhishek kumar': { dse: 'Abhishekhar Kumar', counter: 'Chandrahiya' },
  'abhishekhar': { dse: 'Abhishekhar Kumar', counter: 'Chandrahiya' },
  'anand kumar': { dse: 'Anand Kumar', counter: 'Chandrahiya' },
  'anand': { dse: 'Anand Kumar', counter: 'Chandrahiya' },
  'dev kumar': { dse: 'Dev Kumar', counter: 'Chandrahiya' },
  'dev': { dse: 'Dev Kumar', counter: 'Chandrahiya' },
  'seibel administrator': { dse: 'Seibel Administrator', counter: 'Chandrahiya' },
  'siebel administrator': { dse: 'Seibel Administrator', counter: 'Chandrahiya' },
  'siebel admin': { dse: 'Seibel Administrator', counter: 'Chandrahiya' },
  'seibel admin': { dse: 'Seibel Administrator', counter: 'Chandrahiya' },
  'web service': { dse: 'Web Service', counter: 'Chandrahiya' },
  'webservice': { dse: 'Web Service', counter: 'Chandrahiya' },
  'niraj': { dse: 'Niraj', counter: 'Raja Bazar' },
  'niraj kumar': { dse: 'Niraj', counter: 'Raja Bazar' },
  'neeraj': { dse: 'Niraj', counter: 'Raja Bazar' },
  'anuj': { dse: 'Anuj', counter: 'Raja Bazar' },
  'anuj kumar': { dse: 'Anuj', counter: 'Raja Bazar' },
  'alok das': { dse: 'Alok Das', counter: 'Priemia' },
  'alok': { dse: 'Alok Das', counter: 'Priemia' },
  'anurag': { dse: 'Anurag', counter: 'Priemia' },
  'anurag kumar': { dse: 'Anurag', counter: 'Priemia' },
  'ravi': { dse: 'Ravi', counter: 'Rupdih' },
  'ravi kumar': { dse: 'Ravi', counter: 'Rupdih' },
  'manish': { dse: 'Manish', counter: 'Rupdih' },
  'manish kumar': { dse: 'Manish', counter: 'Rupdih' },
  'chhotu': { dse: 'Chhotu', counter: 'Rupdih' },
  'chotu': { dse: 'Chhotu', counter: 'Rupdih' },
  'jaiprakash': { dse: 'Jaiprakash', counter: 'Mehsi' },
  'jai prakash': { dse: 'Jaiprakash', counter: 'Mehsi' },
  'jayprakash': { dse: 'Jaiprakash', counter: 'Mehsi' },
  'ranjan': { dse: 'Ranjan', counter: 'Mehsi' },
  'ranjan kumar': { dse: 'Ranjan', counter: 'Mehsi' },
  'rishabh kumar': { dse: 'Rishabh Kumar', counter: 'Rishabh Kumar' },
  'rishabh': { dse: 'Rishabh Kumar', counter: 'Rishabh Kumar' },
  'siswa': { dse: 'Siswa', counter: 'Siswa' },
};

const COUNTER_ALIASES: Record<string, CanonicalCounterName> = {
  'chandrahiya': 'Chandrahiya',
  'chandrahia': 'Chandrahiya',
  'chandrahiya counter': 'Chandrahiya',
  'raja bazar': 'Raja Bazar',
  'rajabazar': 'Raja Bazar',
  'raja bazaar': 'Raja Bazar',
  'priemia': 'Priemia',
  'premia': 'Priemia',
  'rupdih': 'Rupdih',
  'roopdih': 'Rupdih',
  'mehsi': 'Mehsi',
  'rishabh kumar': 'Rishabh Kumar',
  'rishabh': 'Rishabh Kumar',
  'siswa': 'Siswa',
  'siswa bazar': 'Siswa',
};

function cleanKey(val: string): string {
  return val.trim().toLowerCase().replace(/\s+/g, ' ');
}

export function resolveDseAndCounter(
  rawDse: string,
  rawCounter: string
): { dseName: string; counterName: CanonicalCounterName } {
  const cleanedDse = cleanKey(rawDse || '');
  const cleanedCounter = cleanKey(rawCounter || '');

  let matchedDse = (rawDse || '').trim() || 'Unassigned';
  let matchedCounter: CanonicalCounterName = 'Unassigned';

  if (cleanedCounter) {
    if (COUNTER_ALIASES[cleanedCounter]) {
      matchedCounter = COUNTER_ALIASES[cleanedCounter];
    } else {
      for (const [alias, canonical] of Object.entries(COUNTER_ALIASES)) {
        if (cleanedCounter.includes(alias)) {
          matchedCounter = canonical;
          break;
        }
      }
    }
  }

  if (cleanedDse) {
    if (DSE_ALIAS_MAP[cleanedDse]) {
      matchedDse = DSE_ALIAS_MAP[cleanedDse].dse;
      if (matchedCounter === 'Unassigned') {
        matchedCounter = DSE_ALIAS_MAP[cleanedDse].counter;
      }
    } else {
      for (const [alias, info] of Object.entries(DSE_ALIAS_MAP)) {
        if (cleanedDse === alias || cleanedDse.startsWith(alias + ' ') || cleanedDse.includes(alias)) {
          matchedDse = info.dse;
          if (matchedCounter === 'Unassigned') {
            matchedCounter = info.counter;
          }
          break;
        }
      }
    }
  }

  // If counter was Siswa or Rishabh Kumar and DSE wasn't set
  if (matchedDse === 'Unassigned' && (matchedCounter === 'Rishabh Kumar' || matchedCounter === 'Siswa')) {
    matchedDse = matchedCounter;
  }

  return { dseName: matchedDse, counterName: matchedCounter };
}

export function normalizeStatus(rawStatus: string, remarks = ''): {
  status: NormalizedStatus;
  isConverted: boolean;
} {
  const text = `${rawStatus || ''} ${remarks || ''}`.toLowerCase();
  const statusOnly = (rawStatus || '').toLowerCase().trim();

  if (
    statusOnly.includes('convert') ||
    statusOnly.includes('retail') ||
    statusOnly.includes('invoiced') ||
    statusOnly.includes('delivered') ||
    statusOnly.includes('sold') ||
    statusOnly.includes('closed won') ||
    statusOnly === 'won'
  ) {
    return { status: 'Converted', isConverted: true };
  }
  if (statusOnly.includes('book')) {
    return { status: 'Booked', isConverted: true };
  }
  if (statusOnly.includes('lost') || statusOnly.includes('drop') || statusOnly.includes('cancel') || statusOnly.includes('not interested')) {
    return { status: 'Lost', isConverted: false };
  }
  if (statusOnly.includes('hot')) {
    return { status: 'Hot', isConverted: false };
  }
  if (statusOnly.includes('warm')) {
    return { status: 'Warm', isConverted: false };
  }
  if (statusOnly.includes('cold')) {
    return { status: 'Cold', isConverted: false };
  }
  if (statusOnly.includes('follow') || statusOnly.includes('call back') || statusOnly.includes('callback') || statusOnly.includes('pending')) {
    return { status: 'Follow-up', isConverted: false };
  }

  // Secondary check against remarks if status is ambiguous
  if (!statusOnly && (text.includes('converted') || text.includes('retailed') || text.includes('booked'))) {
    return { status: text.includes('booked') ? 'Booked' : 'Converted', isConverted: true };
  }

  return { status: 'Open', isConverted: false };
}

export function normalizeDateString(input: unknown): string {
  if (input === null || input === undefined) return '';
  if (typeof input === 'number' && Number.isFinite(input) && input > 20000 && input < 80000) {
    // Excel serial date
    const utcDays = Math.floor(input - 25569);
    const d = new Date(utcDays * 86400 * 1000);
    if (!Number.isNaN(d.getTime())) {
      return d.toISOString().slice(0, 10);
    }
  }
  const str = String(input).trim();
  if (!str) return '';

  // Already YYYY-MM-DD
  if (/^\d{4}-\d{2}-\d{2}$/.test(str)) {
    return str;
  }

  // YYYY/MM/DD
  const ymdSlash = str.match(/^(\d{4})[\/.](\d{1,2})[\/.](\d{1,2})/);
  if (ymdSlash) {
    return `${ymdSlash[1]}-${ymdSlash[2].padStart(2, '0')}-${ymdSlash[3].padStart(2, '0')}`;
  }

  // DD-MM-YYYY or DD/MM/YYYY
  const dmy = str.match(/^(\d{1,2})[-\/.](\d{1,2})[-\/.](\d{2,4})$/);
  if (dmy) {
    const p1 = Number(dmy[1]);
    const p2 = Number(dmy[2]);
    let yr = Number(dmy[3]);
    if (yr < 100) yr += 2000;
    // Assume DD/MM/YYYY if p1 > 12 or Indian date convention
    let day = p1;
    let month = p2;
    if (p2 > 12 && p1 <= 12) {
      month = p1;
      day = p2;
    }
    if (month >= 1 && month <= 12 && day >= 1 && day <= 31) {
      return `${yr}-${String(month).padStart(2, '0')}-${String(day).padStart(2, '0')}`;
    }
  }

  // Month name formats like 16-Nov-2026 or Nov 16, 2026
  const parsed = new Date(str);
  if (!Number.isNaN(parsed.getTime())) {
    const y = parsed.getFullYear();
    const m = String(parsed.getMonth() + 1).padStart(2, '0');
    const d = String(parsed.getDate()).padStart(2, '0');
    if (y >= 2000 && y <= 2100) {
      return `${y}-${m}-${d}`;
    }
  }

  return str;
}

export function getTodayIsoDate(): string {
  const now = new Date();
  const y = now.getFullYear();
  const m = String(now.getMonth() + 1).padStart(2, '0');
  const d = String(now.getDate()).padStart(2, '0');
  return `${y}-${m}-${d}`;
}

export function isPortalSystemDse(dseName: string): boolean {
  const low = cleanKey(dseName || '');
  return (
    low === 'web service' ||
    low === 'webservice' ||
    low === 'seibel administrator' ||
    low === 'siebel administrator' ||
    low === 'seibel admin' ||
    low === 'siebel admin'
  );
}

const NON_PERSON_CALLER_VALUES = new Set([
  '',
  '-',
  '--',
  'na',
  'n/a',
  'none',
  'null',
  'nil',
  'system',
  'admin',
  'administrator',
  'web service',
  'webservice',
  'seibel administrator',
  'siebel administrator',
  'seibel admin',
  'siebel admin',
]);

export function normalizeCrmName(raw: string, dseName = ''): string {
  const trimmed = (raw || '').trim();
  const low = cleanKey(trimmed);

  // If a real caller name is explicitly mentioned, keep it (omitting any override)
  if (!NON_PERSON_CALLER_VALUES.has(low)) {
    if (low.includes('chandrashekhar') || low.includes('chandra shekhar')) {
      return DEFAULT_CRM_NAME;
    }
    return trimmed;
  }

  // No explicit human caller mentioned: Web Service & Seibel Administrator always default to Chandrashekhar Kumar
  if (isPortalSystemDse(dseName)) {
    return DEFAULT_CRM_NAME;
  }

  return '';
}

export function buildDseTodayFollowupReport(
  rows: EnquiryRow[],
  targetDate: string,
  showAllFollowupsWhenNoExactDateMatch = true
): DseFollowupReport {
  const exactTodayRows = rows.filter(
    (r) =>
      r.followupDate === targetDate ||
      (!r.followupDate && r.status === 'Follow-up' && r.enquiryDate === targetDate)
  );

  // If user uploaded a sheet with dates from another day and exactTodayRows is empty,
  // optionally fall back to all rows that have a followupDate or Follow-up status when flagged,
  // OR strictly filter when targetDate is explicitly requested.
  const effectiveRows =
    exactTodayRows.length > 0
      ? exactTodayRows
      : showAllFollowupsWhenNoExactDateMatch
      ? rows.filter((r) => Boolean(r.followupDate) || r.status === 'Follow-up' || Boolean(r.remarks))
      : [];

  const groupMap = new Map<string, DseFollowupGroup>();

  for (const row of effectiveRows) {
    const key = `${row.dseName}__${row.counterName}`;
    let group = groupMap.get(key);
    if (!group) {
      group = {
        dseName: row.dseName,
        counterName: row.counterName,
        totalFollowups: 0,
        convertedCount: 0,
        hotCount: 0,
        rows: [],
      };
      groupMap.set(key, group);
    }
    group.totalFollowups += 1;
    if (row.isConverted) group.convertedCount += 1;
    if (row.status === 'Hot') group.hotCount += 1;
    group.rows.push(row);
  }

  const groups = Array.from(groupMap.values()).sort(
    (a, b) => b.totalFollowups - a.totalFollowups || a.dseName.localeCompare(b.dseName)
  );

  const flatRows = groups.flatMap((g) => g.rows);

  return {
    targetDate,
    totalFollowups: flatRows.length,
    dseCount: groups.length,
    groups,
    flatRows,
  };
}

export function buildPrincipleModelReport(rows: EnquiryRow[]): PrincipleModelReport {
  const modelMap = new Map<
    string,
    PrincipleModelSummary & { counterCounts: Map<string, number> }
  >();

  for (const row of rows) {
    const model = (row.principleModel || '').trim() || 'Unspecified Model';
    let entry = modelMap.get(model);
    if (!entry) {
      entry = {
        principleModel: model,
        totalEnquiries: 0,
        openCount: 0,
        followupCount: 0,
        hotCount: 0,
        warmCount: 0,
        coldCount: 0,
        bookedCount: 0,
        convertedCount: 0,
        lostCount: 0,
        conversionRate: 0,
        topCounter: '-',
        counterCounts: new Map<string, number>(),
      };
      modelMap.set(model, entry);
    }
    entry.totalEnquiries += 1;
    if (row.status === 'Open') entry.openCount += 1;
    if (row.status === 'Follow-up') entry.followupCount += 1;
    if (row.status === 'Hot') entry.hotCount += 1;
    if (row.status === 'Warm') entry.warmCount += 1;
    if (row.status === 'Cold') entry.coldCount += 1;
    if (row.status === 'Booked') entry.bookedCount += 1;
    if (row.status === 'Converted') entry.convertedCount += 1;
    if (row.status === 'Lost') entry.lostCount += 1;

    const cName = row.counterName || 'Unassigned';
    entry.counterCounts.set(cName, (entry.counterCounts.get(cName) || 0) + 1);
  }

  let totalConverted = 0;
  const summaries: PrincipleModelSummary[] = Array.from(modelMap.values())
    .map((item) => {
      const convertedTotal = item.convertedCount + item.bookedCount;
      totalConverted += convertedTotal;
      let topCounter = '-';
      let maxCount = 0;
      for (const [counter, cnt] of item.counterCounts.entries()) {
        if (cnt > maxCount) {
          maxCount = cnt;
          topCounter = counter;
        }
      }
      return {
        principleModel: item.principleModel,
        totalEnquiries: item.totalEnquiries,
        openCount: item.openCount,
        followupCount: item.followupCount,
        hotCount: item.hotCount,
        warmCount: item.warmCount,
        coldCount: item.coldCount,
        bookedCount: item.bookedCount,
        convertedCount: convertedTotal,
        lostCount: item.lostCount,
        conversionRate:
          item.totalEnquiries > 0
            ? Number(((convertedTotal / item.totalEnquiries) * 100).toFixed(1))
            : 0,
        topCounter,
      };
    })
    .sort((a, b) => b.totalEnquiries - a.totalEnquiries || a.principleModel.localeCompare(b.principleModel));

  return {
    totalModels: summaries.length,
    totalEnquiries: rows.length,
    totalConverted,
    summaries,
    rows,
  };
}

export function buildCrmDailyCallReport(
  rows: EnquiryRow[],
  crmTarget = DEFAULT_CRM_NAME
): CrmDailyCallReport {
  const resolvedRows = rows.map((r) => {
    const resolvedCrm = normalizeCrmName(r.crmName || '', r.dseName);
    return resolvedCrm === r.crmName ? r : { ...r, crmName: resolvedCrm };
  });

  const crmSet = new Set<string>([DEFAULT_CRM_NAME]);
  for (const r of resolvedRows) {
    if (r.crmName) crmSet.add(r.crmName);
  }
  const availableCrmNames = Array.from(crmSet);

  const targetLower = crmTarget.toLowerCase();
  const matchedCalls = resolvedRows.filter((r) => {
    const c = (r.crmName || '').toLowerCase();
    if (!c) return false;
    if (targetLower === 'all') return true;
    return c === targetLower || (c.includes('chandrashekhar') && targetLower.includes('chandrashekhar'));
  });

  const callRows = matchedCalls;

  const dailyMap = new Map<string, CrmDailyCallSummary>();
  for (const row of callRows) {
    const callDate = row.calledDate || row.followupDate || row.enquiryDate || 'Unspecified Date';
    let day = dailyMap.get(callDate);
    if (!day) {
      day = {
        date: callDate,
        totalCalled: 0,
        convertedCount: 0,
        followupCount: 0,
        hotCount: 0,
        lostCount: 0,
        remarksList: [],
      };
      dailyMap.set(callDate, day);
    }
    day.totalCalled += 1;
    if (row.isConverted) day.convertedCount += 1;
    if (row.status === 'Follow-up') day.followupCount += 1;
    if (row.status === 'Hot') day.hotCount += 1;
    if (row.status === 'Lost') day.lostCount += 1;
    if (row.remarks && !day.remarksList.includes(row.remarks)) {
      day.remarksList.push(row.remarks);
    }
  }

  const dailySummaries = Array.from(dailyMap.values()).sort((a, b) =>
    b.date.localeCompare(a.date)
  );

  // Daily Lead Conversion List: converted rows for CRM (or all converted rows if CRM column empty on converted rows)
  const crmConverted = callRows.filter((r) => r.isConverted);
  const conversionList = (crmConverted.length > 0 ? crmConverted : rows.filter((r) => r.isConverted)).sort(
    (a, b) => (b.conversionDate || b.calledDate || b.enquiryDate).localeCompare(a.conversionDate || a.calledDate || a.enquiryDate)
  );

  const totalCalls = callRows.length;
  const totalConversions = crmConverted.length;

  return {
    crmName: crmTarget,
    availableCrmNames,
    totalCalls,
    totalConversions,
    conversionRate:
      totalCalls > 0 ? Number(((totalConversions / totalCalls) * 100).toFixed(1)) : 0,
    dailySummaries,
    callRows,
    conversionList,
  };
}

export function buildCounterWiseReport(rows: EnquiryRow[]): CounterWiseReport {
  const summaries: CounterSummary[] = COUNTER_DEFINITIONS.map((def) => ({
    counterNumber: def.id,
    counterName: def.name,
    assignedDseList: def.dseNames,
    totalEnquiries: 0,
    openCount: 0,
    followupCount: 0,
    hotCount: 0,
    warmCount: 0,
    coldCount: 0,
    bookedCount: 0,
    convertedCount: 0,
    lostCount: 0,
    conversionRate: 0,
    dseBreakdowns: [],
  }));

  const unassignedSummary: CounterSummary = {
    counterNumber: 8,
    counterName: 'Unassigned',
    assignedDseList: [],
    totalEnquiries: 0,
    openCount: 0,
    followupCount: 0,
    hotCount: 0,
    warmCount: 0,
    coldCount: 0,
    bookedCount: 0,
    convertedCount: 0,
    lostCount: 0,
    conversionRate: 0,
    dseBreakdowns: [],
  };

  const summaryMap = new Map<CanonicalCounterName, CounterSummary>();
  for (const s of summaries) {
    summaryMap.set(s.counterName, s);
  }
  summaryMap.set('Unassigned', unassignedSummary);

  const dseMapByCounter = new Map<CanonicalCounterName, Map<string, CounterDseBreakdown>>();
  for (const def of COUNTER_DEFINITIONS) {
    const m = new Map<string, CounterDseBreakdown>();
    for (const dse of def.dseNames) {
      m.set(dse, {
        dseName: dse,
        totalEnquiries: 0,
        openCount: 0,
        followupCount: 0,
        hotCount: 0,
        warmCount: 0,
        coldCount: 0,
        bookedCount: 0,
        convertedCount: 0,
        lostCount: 0,
        conversionRate: 0,
      });
    }
    dseMapByCounter.set(def.name, m);
  }
  dseMapByCounter.set('Unassigned', new Map());

  for (const row of rows) {
    const targetCounter = summaryMap.has(row.counterName) ? row.counterName : 'Unassigned';
    const summary = summaryMap.get(targetCounter)!;
    summary.totalEnquiries += 1;
    if (row.status === 'Open') summary.openCount += 1;
    if (row.status === 'Follow-up') summary.followupCount += 1;
    if (row.status === 'Hot') summary.hotCount += 1;
    if (row.status === 'Warm') summary.warmCount += 1;
    if (row.status === 'Cold') summary.coldCount += 1;
    if (row.status === 'Booked') summary.bookedCount += 1;
    if (row.status === 'Converted') summary.convertedCount += 1;
    if (row.status === 'Lost') summary.lostCount += 1;

    const dseMap = dseMapByCounter.get(targetCounter)!;
    const dseKey = row.dseName || 'Unassigned';
    let dseEntry = dseMap.get(dseKey);
    if (!dseEntry) {
      dseEntry = {
        dseName: dseKey,
        totalEnquiries: 0,
        openCount: 0,
        followupCount: 0,
        hotCount: 0,
        warmCount: 0,
        coldCount: 0,
        bookedCount: 0,
        convertedCount: 0,
        lostCount: 0,
        conversionRate: 0,
      };
      dseMap.set(dseKey, dseEntry);
    }
    dseEntry.totalEnquiries += 1;
    if (row.status === 'Open') dseEntry.openCount += 1;
    if (row.status === 'Follow-up') dseEntry.followupCount += 1;
    if (row.status === 'Hot') dseEntry.hotCount += 1;
    if (row.status === 'Warm') dseEntry.warmCount += 1;
    if (row.status === 'Cold') dseEntry.coldCount += 1;
    if (row.status === 'Booked') dseEntry.bookedCount += 1;
    if (row.status === 'Converted') dseEntry.convertedCount += 1;
    if (row.status === 'Lost') dseEntry.lostCount += 1;
  }

  const finalList = [...summaries];
  if (unassignedSummary.totalEnquiries > 0) {
    finalList.push(unassignedSummary);
  }

  let totalConverted = 0;
  for (const s of finalList) {
    const conv = s.convertedCount + s.bookedCount;
    totalConverted += conv;
    s.conversionRate =
      s.totalEnquiries > 0 ? Number(((conv / s.totalEnquiries) * 100).toFixed(1)) : 0;

    const dseMap = dseMapByCounter.get(s.counterName)!;
    s.dseBreakdowns = Array.from(dseMap.values()).map((d) => ({
      ...d,
      conversionRate:
        d.totalEnquiries > 0
          ? Number((((d.convertedCount + d.bookedCount) / d.totalEnquiries) * 100).toFixed(1))
          : 0,
    }));
  }

  return {
    totalEnquiries: rows.length,
    totalConverted,
    overallConversionRate:
      rows.length > 0 ? Number(((totalConverted / rows.length) * 100).toFixed(1)) : 0,
    summaries: finalList,
  };
}

export function buildAllReports(
  rows: EnquiryRow[],
  options?: {
    targetDate?: string;
    crmTarget?: string;
    strictDateMatch?: boolean;
  }
): AllReportsBundle {
  const targetDate = options?.targetDate || getTodayIsoDate();
  const crmTarget = options?.crmTarget || DEFAULT_CRM_NAME;
  const normalizedRows = rows.map((r) => {
    const resolvedCrm = normalizeCrmName(r.crmName || '', r.dseName);
    return resolvedCrm === r.crmName ? r : { ...r, crmName: resolvedCrm };
  });

  return {
    generatedAt: new Date().toISOString(),
    targetDate,
    crmTarget,
    totalRecords: normalizedRows.length,
    report1DseFollowup: buildDseTodayFollowupReport(
      normalizedRows,
      targetDate,
      !options?.strictDateMatch
    ),
    report2PrincipleModel: buildPrincipleModelReport(normalizedRows),
    report3CrmDailyCall: buildCrmDailyCallReport(normalizedRows, crmTarget),
    report4CounterWise: buildCounterWiseReport(normalizedRows),
  };
}

export const CUSTOM_REPORT_COLUMNS: Array<{
  key: CustomReportColumnKey;
  label: string;
}> = [
  { key: 'enquiryNo', label: 'Enquiry No' },
  { key: 'enquiryDate', label: 'Enquiry Date' },
  { key: 'customerName', label: 'Customer Name' },
  { key: 'mobile', label: 'Mobile No' },
  { key: 'counterName', label: 'Counter Name' },
  { key: 'dseName', label: 'DSE Name' },
  { key: 'principleModel', label: 'Principle Model' },
  { key: 'status', label: 'Status' },
  { key: 'followupDate', label: 'Followup Date' },
  { key: 'calledDate', label: 'Called Date' },
  { key: 'conversionDate', label: 'Conversion Date' },
  { key: 'crmName', label: 'CRM / Caller' },
  { key: 'remarks', label: 'Remarks' },
];

export function buildCustomReport(
  rows: EnquiryRow[],
  config: CustomReportConfig
): CustomReportResult {
  const normalizedRows = rows.map((r) => {
    const resolvedCrm = normalizeCrmName(r.crmName || '', r.dseName);
    return resolvedCrm === r.crmName ? r : { ...r, crmName: resolvedCrm };
  });

  const q = (config.searchQuery || '').trim().toLowerCase();
  const crmFilter = (config.crmName || '').trim().toLowerCase();

  const filtered = normalizedRows.filter((row) => {
    if (config.counters.length > 0 && !config.counters.includes(row.counterName)) {
      return false;
    }
    if (config.dseNames.length > 0 && !config.dseNames.includes(row.dseName)) {
      return false;
    }
    if (config.statuses.length > 0 && !config.statuses.includes(row.status)) {
      return false;
    }
    if (
      config.principleModels.length > 0 &&
      !config.principleModels.includes(row.principleModel)
    ) {
      return false;
    }
    if (crmFilter && crmFilter !== 'all') {
      const rowCrm = (row.crmName || '').toLowerCase();
      if (crmFilter === '__unassigned__') {
        if (rowCrm) return false;
      } else if (rowCrm !== crmFilter) {
        return false;
      }
    }

    const rowDate = row.followupDate || row.calledDate || row.enquiryDate || row.conversionDate;
    if (config.dateFrom && (!rowDate || rowDate < config.dateFrom)) {
      return false;
    }
    if (config.dateTo && (!rowDate || rowDate > config.dateTo)) {
      return false;
    }

    if (q) {
      const hay = `${row.enquiryNo} ${row.customerName} ${row.mobile} ${row.dseName} ${row.counterName} ${row.principleModel} ${row.status} ${row.crmName} ${row.remarks}`.toLowerCase();
      if (!hay.includes(q)) return false;
    }

    return true;
  });

  let convertedMatched = 0;
  for (const r of filtered) {
    if (r.isConverted) convertedMatched += 1;
  }

  const groupsMap = new Map<string, CustomReportGroupRow>();
  if (config.groupBy !== 'none') {
    for (const row of filtered) {
      let key = 'Unspecified';
      if (config.groupBy === 'counterName') key = row.counterName || 'Unassigned';
      else if (config.groupBy === 'dseName') key = row.dseName || 'Unassigned';
      else if (config.groupBy === 'principleModel') key = row.principleModel || 'Unspecified Model';
      else if (config.groupBy === 'status') key = row.status;
      else if (config.groupBy === 'crmName') key = row.crmName || 'Unassigned Caller';
      else if (config.groupBy === 'date') {
        key = row.followupDate || row.calledDate || row.enquiryDate || 'Unspecified Date';
      }

      let g = groupsMap.get(key);
      if (!g) {
        g = {
          groupKey: key,
          total: 0,
          openCount: 0,
          followupCount: 0,
          hotCount: 0,
          convertedCount: 0,
          lostCount: 0,
          conversionRate: 0,
        };
        groupsMap.set(key, g);
      }
      g.total += 1;
      if (row.status === 'Open') g.openCount += 1;
      if (row.status === 'Follow-up') g.followupCount += 1;
      if (row.status === 'Hot') g.hotCount += 1;
      if (row.isConverted) g.convertedCount += 1;
      if (row.status === 'Lost') g.lostCount += 1;
    }
  }

  const groups = Array.from(groupsMap.values())
    .map((g) => ({
      ...g,
      conversionRate: g.total > 0 ? Number(((g.convertedCount / g.total) * 100).toFixed(1)) : 0,
    }))
    .sort((a, b) => b.total - a.total || a.groupKey.localeCompare(b.groupKey));

  return {
    config,
    totalMatched: filtered.length,
    convertedMatched,
    conversionRate:
      filtered.length > 0
        ? Number(((convertedMatched / filtered.length) * 100).toFixed(1))
        : 0,
    groups,
    rows: filtered,
  };
}

export const ALL_NORMALIZED_STATUSES: NormalizedStatus[] = [
  'Open',
  'Follow-up',
  'Hot',
  'Warm',
  'Cold',
  'Booked',
  'Converted',
  'Lost',
];

export interface ReportFilterCriteria {
  counter?: string;
  dse?: string;
  principleModel?: string;
  status?: string;
  searchQuery?: string;
}

export function filterEnquiryRows(
  rows: EnquiryRow[],
  criteria: ReportFilterCriteria
): EnquiryRow[] {
  const counter = (criteria.counter || '').trim();
  const dse = (criteria.dse || '').trim();
  const model = (criteria.principleModel || '').trim();
  const status = (criteria.status || '').trim();
  const q = (criteria.searchQuery || '').trim().toLowerCase();

  return rows.filter((row) => {
    if (counter && counter !== 'ALL' && row.counterName !== counter) {
      return false;
    }
    if (dse && dse !== 'ALL' && row.dseName !== dse) {
      return false;
    }
    if (
      model &&
      model !== 'ALL' &&
      (row.principleModel || 'Unspecified Model') !== model
    ) {
      return false;
    }
    if (status && status !== 'ALL') {
      if (status === 'CONVERTED_OR_BOOKED') {
        if (!row.isConverted) return false;
      } else if (row.status !== status) {
        return false;
      }
    }
    if (q) {
      const hay = `${row.enquiryNo} ${row.customerName} ${row.mobile} ${row.dseName} ${row.counterName} ${row.principleModel} ${row.status} ${row.crmName} ${row.remarks}`.toLowerCase();
      if (!hay.includes(q)) return false;
    }
    return true;
  });
}

