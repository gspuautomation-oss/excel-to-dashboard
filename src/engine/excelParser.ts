import * as XLSX from 'xlsx';
import {
  CashFlowLineCategory,
  CashFlowStatementItem,
  ColumnMappingConfig,
  DetectedReportType,
  GSTInvoice,
  GSTR3BSummary,
  ImportValidationMessage,
  Ledger,
  OutstandingInvoice,
  ReportedStatementSubtotal,
  toPaise,
  Voucher,
  VoucherEntry,
} from '../types/accounting';
import {
  classifyLedgerAndGroup,
  normalizeInvoiceNo,
  normalizeName,
  stripLeadingBulletOrRoman,
} from './ledgerMapping';
import { normalizeVoucherType, parseTallyDate } from './xmlParser';

export interface ParsedExcelResult {
  detectedReportType: DetectedReportType;
  extractedCompanyName?: string;
  detectedCurrency?: string;
  periodStart?: string;
  periodEnd?: string;
  currentPeriodLabel?: string;
  comparativePeriodLabel?: string;
  hasMonthlyGranularity?: boolean;
  ledgers: Ledger[];
  vouchers: Voucher[];
  outstandingInvoices: OutstandingInvoice[];
  gstInvoices: GSTInvoice[];
  gstr3bSummaries: GSTR3BSummary[];
  cashFlowItems: CashFlowStatementItem[];
  reportedSubtotals: ReportedStatementSubtotal[];
  validationMessages: ImportValidationMessage[];
  rawHeaders: string[];
  rawSampleRows: Record<string, string | number>[];
  needsMapping: boolean;
}

function normalizeHeader(h: unknown): string {
  if (h === null || h === undefined) return '';
  if (typeof h === 'number' && h > 35000 && h < 65000) {
    const utcDays = Math.floor(h - 25569);
    const dt = new Date(utcDays * 86400 * 1000);
    return dt.toISOString().slice(0, 10);
  }
  return String(h)
    .trim()
    .replace(/\s+/g, ' ');
}

function parseExcelCellDate(val: unknown): string {
  if (val === null || val === undefined || val === '') return '';
  if (typeof val === 'number') {
    if (val > 25000 && val < 70000) {
      const utcDays = Math.floor(val - 25569);
      const dt = new Date(utcDays * 86400 * 1000);
      return dt.toISOString().slice(0, 10);
    }
  }
  return parseTallyDate(String(val));
}

const MONTH_NAME_MAP: Record<string, string> = {
  jan: '01',
  feb: '02',
  mar: '03',
  apr: '04',
  may: '05',
  jun: '06',
  jul: '07',
  aug: '08',
  sep: '09',
  oct: '10',
  nov: '11',
  dec: '12',
};

/**
 * Checks if a column header represents an ANNUAL or AS-AT period-end column
 * (e.g. "For the year ended 31st March 2026", "As at 31-03-2026", "FY 2025-26", "2025-26", "31-Mar-26").
 * Returns the ending 4-digit year (e.g. 2026 or 2025) if matched, or null.
 */
function parseAnnualColumnYear(headerText: string): { endYear: number; label: string } | null {
  const s = headerText.trim();
  const lower = s.toLowerCase();

  // Explicit FY range: "2025-26", "2025-2026", "FY 2025-26", "FY 25-26", "2024-25"
  const fyRange = lower.match(/\b(20\d{2}|\d{2})\s*[-–/]\s*(20\d{2}|\d{2})\b/);
  if (fyRange) {
    const y1Str = fyRange[1].length === 2 ? `20${fyRange[1]}` : fyRange[1];
    const y1 = Number(y1Str);
    const y2Raw = fyRange[2];
    const y2 = y2Raw.length === 2 ? Number(`${y1Str.slice(0, 2)}${y2Raw}`) : Number(y2Raw);
    if (y2 === y1 + 1 && y1 >= 2000 && y1 <= 2040) {
      return { endYear: y2, label: `FY ${y1}–${String(y2).slice(-2)}` };
    }
  }

  // Full date: "31-03-2026", "31.03.2026", "31/03/2026", "2026-03-31"
  const dmyMatch = lower.match(/\b(31|30|28|29)[./\-](0?3|0?6|0?9|12)[./\-](20\d{2}|\d{2})\b/);
  if (dmyMatch) {
    const yrStr = dmyMatch[3].length === 2 ? `20${dmyMatch[3]}` : dmyMatch[3];
    const yr = Number(yrStr);
    return { endYear: yr, label: `FY ${yr - 1}–${String(yr).slice(-2)}` };
  }
  const isoDateMatch = lower.match(/\b(20\d{2})[./\-](0?3|12)[./\-](31|30)\b/);
  if (isoDateMatch) {
    const yr = Number(isoDateMatch[1]);
    return { endYear: yr, label: `FY ${yr - 1}–${String(yr).slice(-2)}` };
  }

  // Textual date: "31st March 2026", "31 March, 2026", "March 31, 2026", "31-Mar-2026", "31-Mar-26"
  const marDayMatch = lower.match(
    /\b(?:31(?:st)?[\s\-.,/]*(?:mar|march)|(?:mar|march)[\s\-.,/]*31(?:st)?)[\s\-.,/']*(20\d{2}|\d{2})\b/
  );
  if (marDayMatch) {
    const yrStr = marDayMatch[1].length === 2 ? `20${marDayMatch[1]}` : marDayMatch[1];
    const yr = Number(yrStr);
    return { endYear: yr, label: `FY ${yr - 1}–${String(yr).slice(-2)}` };
  }

  // Any header mentioning March + 4-digit year, or "Year ended / As at / As on / FY" + 4-digit year
  if (
    lower.includes('year ended') ||
    lower.includes('as at') ||
    lower.includes('as on') ||
    lower.includes('march') ||
    /\bmar\b/.test(lower) ||
    /\bf\.?y\.?\b/.test(lower)
  ) {
    const yrMatch = lower.match(/\b(20\d{2})\b/);
    if (yrMatch) {
      const yr = Number(yrMatch[1]);
      return { endYear: yr, label: `FY ${yr - 1}–${String(yr).slice(-2)}` };
    }
  }

  return null;
}

/**
 * Parses a strictly monthly column header (e.g. "Apr-25", "May-25", "2025-04").
 * Does NOT match full dates like "31-Mar-26" or "31st March 2026" or "FY 2025-26".
 */
function parseHeaderToYearMonth(headerText: string): string | null {
  if (parseAnnualColumnYear(headerText)) return null;
  const s = headerText.trim().toLowerCase();
  if (
    s.includes('year') ||
    s.includes('as at') ||
    s.includes('as on') ||
    s.includes('31') ||
    s.includes('30') ||
    s.includes('fy')
  ) {
    return null;
  }

  const isoMatch = s.match(/^(20\d{2})[-/](0[1-9]|1[0-2])$/);
  if (isoMatch) return `${isoMatch[1]}-${isoMatch[2]}`;

  const mmmYy = s.match(
    /^(jan|feb|mar|apr|may|jun|jul|aug|sep|oct|nov|dec)[a-z]*[\s\-/']*(\d{2,4})$/
  );
  if (mmmYy) {
    const mm = MONTH_NAME_MAP[mmmYy[1]];
    let yr = mmmYy[2];
    if (yr.length === 2) yr = `20${yr}`;
    return `${yr}-${mm}`;
  }
  return null;
}

/**
 * Filters out audit signature blocks, director sign-offs, UDIN, place/date, and EPS rows.
 */
export function isSignatureOrFooterRow(rawText: string): boolean {
  const n = rawText.trim().toLowerCase();
  if (!n) return true;
  if (
    n.includes('for and on behalf of') ||
    n.includes('as per our report') ||
    n.includes('chartered accountant') ||
    n.includes('firm registration') ||
    n.includes('firm reg') ||
    n.startsWith('frn') ||
    n.includes('membership no') ||
    n.includes('m.no') ||
    n.includes('m. no') ||
    n.includes('udin') ||
    n.startsWith('din:') ||
    n.startsWith('din ') ||
    n.startsWith('din-') ||
    n.includes('director') ||
    n.includes('company secretary') ||
    n.includes('partner') ||
    n.startsWith('place:') ||
    n.startsWith('place :') ||
    n === 'place' ||
    n.startsWith('date:') ||
    n.startsWith('date :') ||
    n === 'date' ||
    n.includes('accompanying notes') ||
    n.includes('significant accounting policies') ||
    n.includes('notes forming part of') ||
    n.includes('notes to financial statements') ||
    n.includes('notes to the financial statements') ||
    n.includes('earnings per equity share') ||
    n.includes('earning per share') ||
    n.startsWith('basic (') ||
    n.startsWith('diluted (') ||
    n === 'basic' ||
    n === 'diluted' ||
    n.includes('authorised share capital') ||
    n.includes('authorized share capital') ||
    n.startsWith('authorised capital') ||
    n.startsWith('authorized capital')
  ) {
    return true;
  }
  return false;
}

/**
 * Detects whether a worksheet is a supporting Note / Schedule / Fixed Asset / Depreciation / Tax / Working sheet
 * rather than a primary Financial Statement (Balance Sheet, Profit & Loss, Cash Flow).
 */
function isSupportingOrNotesWorksheet(sheetName: string, rows2D: unknown[][]): boolean {
  const sLower = sheetName.trim().toLowerCase();
  if (
    /(?:^|[\s\-_.,()])(note|notes|sch|schedule|schedules|sub[\s\-_]*sch|grouping|groupings|working|workings|fixed[\s\-_]*asset|fa[\s\-_]*sch|depreciation|depn|dep[\s\-_]*sch|wdv|ppe|deferred[\s\-_]*tax|dta|dtl|tax[\s\-_]*comp|computation|ratio|ratios|cover|index|contents|policy|policies|annexure)(?:$|[\s\-_.,()\d])/i.test(
      sLower
    ) ||
    /^(?:note|notes|sch|schedule|annex|fa|dep|grouping|working)[\s\-_\d]*/i.test(sLower)
  ) {
    return true;
  }

  // Check top 6 non-empty rows for title banners indicating Notes / Fixed Asset Schedule / Tax Computation
  const topLines: string[] = [];
  for (let r = 0; r < Math.min(8, rows2D.length); r++) {
    const rowText = (rows2D[r] || [])
      .map((c) => String(c ?? '').trim())
      .filter(Boolean)
      .join(' ')
      .toLowerCase();
    if (rowText) topLines.push(rowText);
  }
  const topBanner = topLines.slice(0, 5).join(' ');
  const hasPrimaryStatementTitle =
    topBanner.includes('balance sheet as at') ||
    topBanner.includes('statement of profit and loss') ||
    topBanner.includes('statement of profit & loss') ||
    topBanner.includes('profit and loss account') ||
    topBanner.includes('profit & loss account') ||
    topBanner.includes('cash flow statement') ||
    topBanner.includes('statement of cash flow');

  if (
    !hasPrimaryStatementTitle &&
    (topBanner.includes('notes forming part of') ||
      topBanner.includes('notes to financial statements') ||
      topBanner.includes('notes to the financial statements') ||
      topBanner.includes('notes to accounts') ||
      topBanner.includes('notes on accounts') ||
      topBanner.includes('schedules forming part of') ||
      topBanner.includes('schedule of fixed assets') ||
      topBanner.includes('fixed assets schedule') ||
      topBanner.includes('depreciation schedule') ||
      topBanner.includes('depreciation chart') ||
      topBanner.includes('depreciation as per') ||
      topBanner.includes('computation of total income'))
  ) {
    return true;
  }

  // Check if multiple rows in the first 40 rows start with "Note 1", "Note 2", "Note - 1", "Schedule 1"
  let noteHeadingCount = 0;
  for (let r = 0; r < Math.min(45, rows2D.length); r++) {
    const firstCells = (rows2D[r] || [])
      .slice(0, 3)
      .map((c) => String(c ?? '').trim())
      .filter(Boolean)
      .join(' ');
    if (/^(?:note|schedule)\s*(?:no\.?)?\s*[:\-–]?\s*\d+\b/i.test(firstCells)) {
      noteHeadingCount++;
    }
  }
  if (noteHeadingCount >= 2 && !hasPrimaryStatementTitle) {
    return true;
  }

  return false;
}

/**
 * Strips Roman numerals, numbering, and trailing Schedule III formula parentheticals
 * such as "(I + II)", "(III - IV)", "(V - VI)", "(A + B)" to detect subtotals reliably.
 */
function cleanLabelForSubtotalCheck(rawName: string): string {
  return stripLeadingBulletOrRoman(rawName)
    .toLowerCase()
    .replace(/\([ivx0-9a-z\s+\-*/.,]+\)\s*$/i, '')
    .replace(/[:\-–—]+$/, '')
    .replace(/\s+/g, ' ')
    .trim();
}

export function isStatementSubtotalRow(rawName: string): boolean {
  const cleaned = cleanLabelForSubtotalCheck(rawName);
  const rawLower = rawName.trim().toLowerCase();
  if (!cleaned) return true;

  // CRITICAL SCHEDULE III PROTECTION:
  // Trade Payables child lines in Schedule III begin with:
  // "(A) Total outstanding dues of micro enterprises and small enterprises"
  // "(B) Total outstanding dues of creditors other than micro enterprises and small enterprises"
  // These are LEAF LINE ITEMS under Trade Payables, NEVER statement subtotals!
  if (
    cleaned.includes('total outstanding dues') ||
    cleaned.includes('outstanding dues of') ||
    cleaned.includes('dues of micro') ||
    cleaned.includes('dues of creditor') ||
    cleaned.includes('micro enterprises and small enterprises') ||
    cleaned.includes('micro and small enterprises') ||
    cleaned.includes('creditors other than micro')
  ) {
    return false;
  }

  if (
    cleaned === 'total' ||
    cleaned === 'grand total' ||
    cleaned === 'sub total' ||
    cleaned === 'subtotal' ||
    cleaned === 'sub-total' ||
    cleaned.startsWith('total ') ||
    cleaned.endsWith(' total')
  ) {
    return true;
  }

  if (
    cleaned.includes('profit before exceptional') ||
    cleaned.includes('profit before extraordinary') ||
    cleaned.includes('profit before prior period') ||
    cleaned.includes('profit / (loss) before') ||
    cleaned.includes('profit/(loss) before') ||
    cleaned.includes('profit before tax') ||
    cleaned.includes('loss before tax') ||
    cleaned.includes('profit after tax') ||
    cleaned.includes('profit / (loss) after tax') ||
    cleaned.includes('profit/(loss) after tax') ||
    cleaned.includes('profit for the year') ||
    cleaned.includes('profit for the period') ||
    cleaned.includes('profit / (loss) for the period') ||
    cleaned.includes('profit/(loss) for the period') ||
    cleaned.includes('profit/(loss) for the year') ||
    cleaned.includes('profit from continuing operations') ||
    cleaned.includes('total tax expense') ||
    cleaned.includes('net tax expense') ||
    cleaned.includes('total shareholders') ||
    cleaned.includes('total equity and liabilities') ||
    cleaned.includes('total liabilities and equity') ||
    cleaned.includes('total non-current liabilities') ||
    cleaned.includes('total current liabilities') ||
    cleaned.includes('total non-current assets') ||
    cleaned.includes('total current assets') ||
    cleaned.includes('total assets') ||
    rawLower.includes('(i + ii)') ||
    rawLower.includes('(i+ii)') ||
    rawLower.includes('(iii - iv)') ||
    rawLower.includes('(iii-iv)') ||
    rawLower.includes('(v - vi)') ||
    rawLower.includes('(v-vi)')
  ) {
    return true;
  }

  const exactSubtotals = new Set([
    'gross profit',
    'gross margin',
    'ebitda',
    'operating profit',
    'ebit',
    'profit before tax',
    'pbt',
    'profit before tax (pbt)',
    'profit after tax',
    'pat',
    'profit after tax (pat)',
    'net profit',
    'net income',
    'net profit / (loss)',
    'total revenue',
    'total income',
    'total expenses',
    'total operating expenses',
    'total non-current assets',
    'total current assets',
    'total assets',
    'total equity',
    'shareholders funds',
    'total shareholders equity',
    'total non-current liabilities',
    'total current liabilities',
    'total liabilities',
    'total equity and liabilities',
    'total liabilities and equity',
    'net working capital',
  ]);
  return exactSubtotals.has(cleaned);
}

/**
 * Classifies a Cash Flow Statement row, distinguishing intermediate subtotals
 * ("Operating profit before working capital changes", "Cash generated from operations")
 * from the final section totals ("Net cash from operating activities", etc.).
 */
function classifyCashFlowRow(
  rawName: string,
  activeSection: 'OPERATING' | 'INVESTING' | 'FINANCING'
): { category: CashFlowLineCategory; isSubtotal: boolean } {
  const n = stripLeadingBulletOrRoman(rawName).toLowerCase();

  // Opening & Closing Cash and Net Change (Check these first as they appear at the bottom)
  if (
    n.includes('opening cash') ||
    n.includes('cash and cash equivalents at beginning') ||
    n.includes('cash and cash equivalents at the beginning') ||
    n.includes('cash & cash equivalents at beginning') ||
    n.includes('cash & cash equivalents at the beginning') ||
    n.includes('cash at beginning') ||
    n.includes('cash at the beginning') ||
    n.includes('opening balance of cash')
  ) {
    return { category: 'CF_OPENING_CASH', isSubtotal: true };
  }
  if (
    n.includes('closing cash') ||
    n.includes('cash and cash equivalents at end') ||
    n.includes('cash and cash equivalents at the end') ||
    n.includes('cash & cash equivalents at end') ||
    n.includes('cash & cash equivalents at the end') ||
    n.includes('cash at end') ||
    n.includes('cash at the end') ||
    n.includes('closing balance of cash')
  ) {
    return { category: 'CF_CLOSING_CASH', isSubtotal: true };
  }
  if (
    n.includes('net increase') ||
    n.includes('net decrease') ||
    n.includes('net (decrease)') ||
    n.includes('net change in cash') ||
    n.includes('net cash movement') ||
    n.includes('a+b+c') ||
    n.includes('a + b + c') ||
    (n.includes('net cash flow') && !n.includes('operating') && !n.includes('investing') && !n.includes('financing'))
  ) {
    return { category: 'CF_NET_CHANGE', isSubtotal: true };
  }

  // Intermediate Operating Subtotals (Before Working Capital or Before Taxes)
  if (
    n.includes('operating profit before working capital') ||
    n.includes('cash generated from operations') ||
    n.includes('cash generated from operating activities before') ||
    n.includes('cash flow before working capital') ||
    n.includes('cash flow before extraordinary')
  ) {
    return { category: 'CFO_INTERMEDIATE_SUBTOTAL', isSubtotal: true };
  }

  // Final Operating Cash Flow Total
  if (
    n.includes('net cash from operating') ||
    n.includes('net cash flow from operating') ||
    n.includes('net cash generated from operating') ||
    n.includes('net cash used in operating') ||
    n.includes('net cash from / (used in) operating') ||
    n.includes('net cash generated/(used) in operating') ||
    n.includes('net operating cash flow') ||
    n === 'cash flow from operating activities' ||
    n === 'cash flows from operating activities' ||
    n === 'operating cash flow' ||
    n === 'total operating cash flow' ||
    (activeSection === 'OPERATING' && (n === 'total' || n.startsWith('net cash')))
  ) {
    return { category: 'CFO_OPERATING_TOTAL', isSubtotal: true };
  }

  // Final Investing Cash Flow Total
  if (
    n.includes('net cash from investing') ||
    n.includes('net cash flow from investing') ||
    n.includes('net cash used in investing') ||
    n.includes('net cash flow used in investing') ||
    n.includes('net cash generated from investing') ||
    n.includes('net cash from / (used in) investing') ||
    n.includes('net investing cash flow') ||
    n === 'cash flow from investing activities' ||
    n === 'cash flows from investing activities' ||
    n === 'investing cash flow' ||
    n === 'total investing cash flow' ||
    (activeSection === 'INVESTING' && (n === 'total' || n.startsWith('net cash')))
  ) {
    return { category: 'CFI_INVESTING_TOTAL', isSubtotal: true };
  }

  // Final Financing Cash Flow Total
  if (
    n.includes('net cash from financing') ||
    n.includes('net cash flow from financing') ||
    n.includes('net cash used in financing') ||
    n.includes('net cash flow used in financing') ||
    n.includes('net cash generated from financing') ||
    n.includes('net cash from / (used in) financing') ||
    n.includes('net financing cash flow') ||
    n === 'cash flow from financing activities' ||
    n === 'cash flows from financing activities' ||
    n === 'financing cash flow' ||
    n === 'total financing cash flow' ||
    (activeSection === 'FINANCING' && (n === 'total' || n.startsWith('net cash')))
  ) {
    return { category: 'CFF_FINANCING_TOTAL', isSubtotal: true };
  }

  // Line items inside sections
  if (
    n.includes('capital expenditure') ||
    n.includes('capex') ||
    n.includes('purchase of property') ||
    n.includes('purchase of fixed asset') ||
    n.includes('additions to fixed asset') ||
    n.includes('acquisition of ppe') ||
    (activeSection === 'INVESTING' && (n.includes('fixed asset') || n.includes('property, plant') || n.includes('intangible')))
  ) {
    return { category: 'CFI_CAPEX', isSubtotal: false };
  }
  if (
    activeSection === 'OPERATING' &&
    (n.includes('tax paid') || n.includes('taxes paid') || n.includes('direct tax') || n.includes('income tax')) &&
    !n.includes('profit before tax')
  ) {
    return { category: 'CFO_TAXES_PAID', isSubtotal: false };
  }
  if (
    activeSection === 'OPERATING' &&
    (n.includes('working capital') ||
      n.includes('trade receivable') ||
      n.includes('inventories') ||
      n.includes('trade payable') ||
      n.includes('other current asset') ||
      n.includes('other current liabilit') ||
      n.includes('short-term provision') ||
      n.includes('loans and advances') ||
      n.includes('debtors') ||
      n.includes('creditors'))
  ) {
    return { category: 'CFO_WORKING_CAPITAL_CHANGE', isSubtotal: false };
  }
  if (
    activeSection === 'FINANCING' &&
    (n.includes('borrowing') || n.includes('loan') || n.includes('repayment of debt'))
  ) {
    return { category: 'CFF_BORROWINGS', isSubtotal: false };
  }

  if (activeSection === 'INVESTING') {
    return { category: 'CFI_INVESTING_ITEM', isSubtotal: false };
  }
  if (activeSection === 'FINANCING') {
    return { category: 'CFF_FINANCING_ITEM', isSubtotal: false };
  }
  return { category: 'CFO_OPERATING_ITEM', isSubtotal: false };
}

/**
 * Extracts a clean company name from the top rows of a worksheet if present.
 */
function extractCompanyNameFromRows(rows2D: unknown[][]): string {
  const ignorePatterns = [
    /balance\s*sheet/i,
    /profit\s*(and|&)\s*loss/i,
    /statement\s*of/i,
    /cash\s*flow/i,
    /trial\s*balance/i,
    /particulars/i,
    /note\s*no/i,
    /schedule/i,
    /^cin\s*[:\-]/i,
    /^pan\s*[:\-]/i,
    /^gstin\s*[:\-]/i,
    /amount\s*in/i,
    /\(in\s*rupees/i,
    /\(in\s*rs/i,
    /\(in\s*lakhs/i,
    /for\s*the\s*year\s*ended/i,
    /as\s*at\s*31/i,
  ];

  let fallbackCandidate = '';

  for (let r = 0; r < Math.min(8, rows2D.length); r++) {
    const row = rows2D[r] || [];
    const cells = row
      .map((c) => (typeof c === 'string' ? c.trim() : ''))
      .filter((c) => c.length > 2);
    if (cells.length === 0) continue;

    // Check for explicit "Name of the Company: XYZ"
    for (let i = 0; i < cells.length; i++) {
      const cell = cells[i];
      if (/^(?:name\s*of\s*(?:the\s*)?company|company\s*name|entity\s*name)\s*[:\-]?$/i.test(cell)) {
        if (cells[i + 1] && cells[i + 1].length > 3) {
          return cells[i + 1].trim();
        }
      }
      const inlineMatch = cell.match(
        /^(?:name\s*of\s*(?:the\s*)?company|company\s*name|entity\s*name)\s*[:\-]\s*(.+)$/i
      );
      if (inlineMatch && inlineMatch[1].trim().length > 3) {
        return inlineMatch[1].trim();
      }
    }

    // Check for corporate legal entity suffix (Private Limited, Pvt Ltd, Limited, LLP, Inc, etc.)
    for (const cell of cells) {
      if (ignorePatterns.some((pat) => pat.test(cell))) continue;
      if (
        /\b(private\s+limited|pvt\.?\s*ltd\.?|limited|ltd\.?|llp|inc\.?|corp\.?|llc|enterprise|solutions|technologies|consulting|industries)\b/i.test(
          cell
        )
      ) {
        return cell.replace(/^(?:m\/s\.?\s*)/i, '').trim();
      }
      if (!fallbackCandidate && r <= 2 && cells.length <= 2 && cell.length >= 4 && !/^\d+$/.test(cell)) {
        fallbackCandidate = cell;
      }
    }
  }

  return fallbackCandidate;
}

export function parseExcelOrCsvBuffer(
  buffer: ArrayBuffer,
  sourceFileId: string,
  sourceFilename: string,
  customMapping?: ColumnMappingConfig
): ParsedExcelResult {
  const validationMessages: ImportValidationMessage[] = [];
  let workbook: XLSX.WorkBook;

  try {
    workbook = XLSX.read(buffer, { type: 'array', cellDates: false });
  } catch (err) {
    validationMessages.push({
      severity: 'ERROR',
      code: 'EXCEL_READ_FAILED',
      message: `Unable to read spreadsheet file: ${err instanceof Error ? err.message : String(err)}`,
    });
    return {
      detectedReportType: 'UNKNOWN_UNSUPPORTED',
      ledgers: [],
      vouchers: [],
      outstandingInvoices: [],
      gstInvoices: [],
      gstr3bSummaries: [],
      cashFlowItems: [],
      reportedSubtotals: [],
      validationMessages,
      rawHeaders: [],
      rawSampleRows: [],
      needsMapping: false,
    };
  }

  const ledgers: Ledger[] = [];
  const vouchers: Voucher[] = [];
  const outstandingInvoices: OutstandingInvoice[] = [];
  const gstInvoices: GSTInvoice[] = [];
  const gstr3bSummaries: GSTR3BSummary[] = [];
  const cashFlowItems: CashFlowStatementItem[] = [];
  const reportedSubtotals: ReportedStatementSubtotal[] = [];

  let extractedCompanyName = '';
  let detectedCurrency = 'INR';
  let minDate = '';
  let maxDate = '';
  let detectedCurrentPeriodLabel = '';
  let detectedComparativePeriodLabel = '';
  let hasAnyMonthlyGranularity = false;

  let primaryReportType: DetectedReportType =
    customMapping?.reportCategoryOverride || 'UNKNOWN_UNSUPPORTED';
  let capturedHeaders: string[] = [];
  let capturedSampleRows: Record<string, string | number>[] = [];

  const lowerFilename = sourceFilename.toLowerCase();
  const detectedSheetTypes = new Set<DetectedReportType>();

  // Pre-scan all worksheets in the workbook to identify supporting Notes/Schedules vs authoritative Financial Statements
  const parsedSheetsCache = new Map<string, unknown[][]>();
  const supportingSheets = new Set<string>();
  let bestBsSheet: { name: string; score: number } | null = null;
  let bestPnlSheet: { name: string; score: number } | null = null;
  let bestCfSheet: { name: string; score: number } | null = null;

  for (const sName of workbook.SheetNames) {
    const sh = workbook.Sheets[sName];
    if (!sh) continue;
    const r2d = XLSX.utils.sheet_to_json<unknown[]>(sh, { header: 1, defval: '' });
    if (r2d.length === 0) continue;
    parsedSheetsCache.set(sName, r2d);

    const isSupp = isSupportingOrNotesWorksheet(sName, r2d);
    if (isSupp) {
      supportingSheets.add(sName);
      continue;
    }

    const sLow = sName.trim().toLowerCase();
    const topText = r2d
      .slice(0, 10)
      .map((r) => (r || []).map((c) => String(c ?? '').trim()).join(' '))
      .join(' ')
      .toLowerCase();
    const bodyText = r2d
      .slice(0, 50)
      .map((r) => (r || []).slice(0, 4).map((c) => String(c ?? '').trim()).join(' '))
      .join(' ')
      .toLowerCase();

    let bsScore = 0;
    if (sLow === 'balance sheet' || sLow === 'balancesheet' || sLow === 'bs' || sLow === 'bal sheet') {
      bsScore += 100;
    } else if (sLow.includes('balance sheet') || sLow.includes('balancesheet')) {
      bsScore += 80;
    }
    if (topText.includes('balance sheet as at') || topText.includes('balance sheet as on')) {
      bsScore += 75;
    }
    if (
      (bodyText.includes('equity and liabilities') ||
        bodyText.includes('equity & liabilities') ||
        bodyText.includes('shareholders')) &&
      (bodyText.includes('non-current assets') ||
        bodyText.includes('non current assets') ||
        bodyText.includes('current assets'))
    ) {
      bsScore += 65;
    }

    let pnlScore = 0;
    if (
      sLow === 'profit & loss' ||
      sLow === 'profit and loss' ||
      sLow === 'p&l' ||
      sLow === 'pnl' ||
      sLow === 'pl' ||
      sLow === 'statement of profit and loss'
    ) {
      pnlScore += 100;
    } else if (
      sLow.includes('profit') ||
      sLow.includes('p&l') ||
      sLow.includes('pnl') ||
      sLow.includes('income statement')
    ) {
      pnlScore += 80;
    }
    if (
      topText.includes('statement of profit and loss') ||
      topText.includes('statement of profit & loss') ||
      topText.includes('profit and loss account') ||
      topText.includes('profit & loss account')
    ) {
      pnlScore += 75;
    }
    if (
      bodyText.includes('revenue from operations') &&
      (bodyText.includes('total expenses') ||
        bodyText.includes('profit before tax') ||
        bodyText.includes('employee benefit'))
    ) {
      pnlScore += 65;
    }

    let cfScore = 0;
    if (sLow.includes('cash flow') || sLow.includes('cashflow') || sLow === 'cfs') {
      cfScore += 100;
    }
    if (topText.includes('cash flow statement') || topText.includes('statement of cash flow')) {
      cfScore += 75;
    }
    if (bodyText.includes('operating activities') && bodyText.includes('investing activities')) {
      cfScore += 65;
    }

    if (bsScore >= 60 && bsScore >= pnlScore && bsScore >= cfScore) {
      if (!bestBsSheet || bsScore > bestBsSheet.score) {
        bestBsSheet = { name: sName, score: bsScore };
      }
    } else if (pnlScore >= 60 && pnlScore > bsScore && pnlScore >= cfScore) {
      if (!bestPnlSheet || pnlScore > bestPnlSheet.score) {
        bestPnlSheet = { name: sName, score: pnlScore };
      }
    } else if (cfScore >= 60 && cfScore > bsScore && cfScore > pnlScore) {
      if (!bestCfSheet || cfScore > bestCfSheet.score) {
        bestCfSheet = { name: sName, score: cfScore };
      }
    }
  }

  const hasAuthoritativeFinancialStatements = Boolean(bestBsSheet || bestPnlSheet);

  for (const sheetName of workbook.SheetNames) {
    const rows2D = parsedSheetsCache.get(sheetName);
    if (!rows2D || rows2D.length === 0) continue;

    // 1. Extract Company Name from top rows (even before skipping supporting sheets)
    const sheetCompany = extractCompanyNameFromRows(rows2D);
    if (sheetCompany && !extractedCompanyName) {
      extractedCompanyName = sheetCompany;
    }

    // Skip supporting Notes, Schedules, Fixed Assets, Depreciation, Groupings & Working sheets.
    // Furthermore, when a workbook contains authoritative primary Financial Statements (Balance Sheet / P&L),
    // strictly isolate extraction to ONLY the authoritative BS, P&L, and Cash Flow worksheets so no other tab can contaminate them.
    if (!customMapping?.reportCategoryOverride) {
      if (hasAuthoritativeFinancialStatements) {
        const isAuthoritativeSheet =
          sheetName === bestBsSheet?.name ||
          sheetName === bestPnlSheet?.name ||
          sheetName === bestCfSheet?.name;
        if (!isAuthoritativeSheet) {
          continue;
        }
      } else if (supportingSheets.has(sheetName)) {
        continue;
      }
    }

    // 2. Locate Header Row within top 15 rows
    let headerRowIndex = 0;
    let bestScore = -1;

    for (let r = 0; r < Math.min(15, rows2D.length); r++) {
      const row = rows2D[r] || [];
      const nonEmpty = row.map(normalizeHeader).filter(Boolean);
      const joinedLower = nonEmpty.join(' ').toLowerCase();

      if (joinedLower.includes('usd') || joinedLower.includes('$')) detectedCurrency = 'USD';
      else if (joinedLower.includes('eur') || joinedLower.includes('€')) detectedCurrency = 'EUR';
      else if (joinedLower.includes('gbp') || joinedLower.includes('£')) detectedCurrency = 'GBP';
      else if (joinedLower.includes('aed')) detectedCurrency = 'AED';

      let score = nonEmpty.length;
      if (
        joinedLower.includes('particulars') ||
        joinedLower.includes('ledger') ||
        joinedLower.includes('account') ||
        joinedLower.includes('line item') ||
        joinedLower.includes('description') ||
        joinedLower.includes('voucher') ||
        joinedLower.includes('invoice') ||
        joinedLower.includes('gstin') ||
        joinedLower.includes('debit') ||
        joinedLower.includes('credit') ||
        joinedLower.includes('current') ||
        joinedLower.includes('previous') ||
        joinedLower.includes('note') ||
        joinedLower.includes('as at') ||
        joinedLower.includes('year ended') ||
        joinedLower.includes('2025') ||
        joinedLower.includes('2026')
      ) {
        score += 18;
      }
      if (joinedLower.includes('particulars')) score += 12;
      if (score > bestScore && nonEmpty.length >= 2) {
        bestScore = score;
        headerRowIndex = r;
      }
    }

    // Build composite header row combining headerRowIndex and adjacent header sub-rows (e.g. "As at" + "31st March 2026" + "Rupees")
    const maxCols = Math.max(
      ...(rows2D.slice(0, Math.min(25, rows2D.length)).map((r) => (r ? r.length : 0))),
      1
    );
    const rawHeaderRow: string[] = [];
    let lastIncludedSubHeaderRow = headerRowIndex;

    for (let c = 0; c < maxCols; c++) {
      const parts: string[] = [];
      const startH = Math.max(0, headerRowIndex - 2);
      const endH = Math.min(rows2D.length - 1, headerRowIndex + 2);
      for (let hr = startH; hr <= endH; hr++) {
        const rowCellsNorm = (rows2D[hr] || []).map(normalizeHeader);
        const rowJoinedLower = rowCellsNorm.join(' ').toLowerCase();
        const isSectionStartRow =
          rowJoinedLower.includes('equity and liabilities') ||
          rowJoinedLower.includes('equity & liabilities') ||
          rowJoinedLower.includes('shareholders') ||
          rowJoinedLower.includes('revenue from operations') ||
          rowJoinedLower.includes('cash flow from operating');

        if (hr > headerRowIndex) {
          if (isSectionStartRow) break;
          const looksLikeSubHeader =
            rowJoinedLower.includes('march') ||
            /\bmar\b/.test(rowJoinedLower) ||
            rowJoinedLower.includes('202') ||
            rowJoinedLower.includes('as at') ||
            rowJoinedLower.includes('as on') ||
            rowJoinedLower.includes('year ended') ||
            rowJoinedLower.includes('rs') ||
            rowJoinedLower.includes('inr') ||
            rowJoinedLower.includes('amount') ||
            rowJoinedLower.includes('rupees') ||
            rowJoinedLower.startsWith('col_');
          if (!looksLikeSubHeader) continue;
          if (hr > lastIncludedSubHeaderRow) lastIncludedSubHeaderRow = hr;
        }
        if (hr < headerRowIndex) {
          const prevRowNonEmpty = rowCellsNorm.filter(Boolean);
          // Skip title banner row if it only has 1 cell across the whole sheet
          if (prevRowNonEmpty.length <= 1) continue;
          const looksLikeUpperHeader =
            rowJoinedLower.includes('particulars') ||
            rowJoinedLower.includes('note') ||
            rowJoinedLower.includes('as at') ||
            rowJoinedLower.includes('as on') ||
            rowJoinedLower.includes('year ended') ||
            rowJoinedLower.includes('for the year') ||
            rowJoinedLower.includes('march') ||
            rowJoinedLower.includes('202');
          if (!looksLikeUpperHeader && hr < headerRowIndex - 1) continue;
        }
        const cellText = rowCellsNorm[c];
        if (cellText && !parts.includes(cellText)) {
          parts.push(cellText);
        }
      }
      const primaryCell = normalizeHeader((rows2D[headerRowIndex] || [])[c]);
      const combined = parts.join(' ').trim() || primaryCell || `Col_${c + 1}`;
      rawHeaderRow.push(combined);
    }

    // Deduplicate identical column names by appending index if needed
    const seenHeaderCounts = new Map<string, number>();
    for (let c = 0; c < rawHeaderRow.length; c++) {
      const h = rawHeaderRow[c];
      const count = seenHeaderCounts.get(h) || 0;
      seenHeaderCounts.set(h, count + 1);
      if (count > 0) {
        rawHeaderRow[c] = `${h}_${c + 1}`;
      }
    }

    if (capturedHeaders.length === 0) {
      capturedHeaders = rawHeaderRow;
    }

    const dataRows: Array<{
      rowIndex1Based: number;
      rawCells: unknown[];
      rowObj: Record<string, string | number>;
    }> = [];
    for (let r = lastIncludedSubHeaderRow + 1; r < rows2D.length; r++) {
      const rowArr = rows2D[r] || [];
      const rowObj: Record<string, string | number> = {};
      let hasContent = false;
      for (let c = 0; c < rawHeaderRow.length; c++) {
        const val = rowArr[c];
        if (val !== undefined && val !== null && String(val).trim() !== '') {
          hasContent = true;
        }
        rowObj[rawHeaderRow[c]] =
          typeof val === 'number' || typeof val === 'string' ? val : String(val ?? '');
      }
      if (hasContent) {
        dataRows.push({ rowIndex1Based: r + 1, rawCells: rowArr, rowObj });
      }
    }

    if (capturedSampleRows.length === 0 && dataRows.length > 0) {
      capturedSampleRows = dataRows.slice(0, 6).map((d) => d.rowObj);
    }

    const findCol = (customCol: string | undefined, patterns: RegExp[]): string | undefined => {
      if (customCol && rawHeaderRow.includes(customCol)) return customCol;
      for (const pat of patterns) {
        const found = rawHeaderRow.find((h) => pat.test(h.toLowerCase()));
        if (found) return found;
      }
      return undefined;
    };

    const colDate = findCol(customMapping?.dateCol, [
      /^date$/,
      /invoice\s*date/,
      /vch\s*date/,
      /bill\s*date/,
      /doc\s*date/,
    ]);
    const colVchNo = findCol(customMapping?.voucherNoCol, [
      /voucher\s*no/,
      /vch\s*no/,
      /invoice\s*no/,
      /bill\s*no/,
      /inv\s*no/,
      /doc\s*no/,
    ]);
    const colVchType = findCol(customMapping?.voucherTypeCol, [
      /voucher\s*type/,
      /vch\s*type/,
      /^type$/,
    ]);
    const colLedger = findCol(customMapping?.ledgerNameCol, [
      /particulars/,
      /ledger/,
      /account\s*name/,
      /line\s*item/,
      /head\s*of\s*account/,
      /description/,
      /schedule\s*item/,
      /cash\s*flow\s*item/,
    ]);
    const colNote = findCol(undefined, [/^note/, /note\s*no/, /^sch/, /schedule\s*no/, /^ref/]);
    const colGroup = findCol(customMapping?.parentGroupCol, [
      /parent\s*group/,
      /^group$/,
      /activity\s*category/,
      /^category$/,
    ]);
    const colParty = findCol(customMapping?.partyNameCol, [
      /party/,
      /customer/,
      /supplier/,
      /vendor/,
      /buyer/,
      /trade\s*name/,
      /receiver/,
    ]);
    const colGstin = findCol(customMapping?.gstinCol, [/gstin/, /gst\s*no/, /uin/]);
    const colDebit = findCol(customMapping?.debitCol, [
      /^debit$/,
      /^debit\s*amount/,
      /^dr\.?$/,
      /gross\s*debit/,
    ]);
    const colCredit = findCol(customMapping?.creditCol, [
      /^credit$/,
      /^credit\s*amount/,
      /^cr\.?$/,
      /gross\s*credit/,
    ]);
    const colOpening = findCol(customMapping?.openingBalanceCol, [/opening\s*bal/, /^opening/]);
    const colBudget = findCol(undefined, [/budget/, /target/, /plan/]);
    const colTaxable = findCol(customMapping?.taxableValueCol, [
      /taxable\s*val/,
      /taxable\s*amt/,
      /^taxable/,
    ]);
    const colIgst = findCol(customMapping?.igstCol, [/igst/, /integrated\s*tax/]);
    const colCgst = findCol(customMapping?.cgstCol, [/cgst/, /central\s*tax/]);
    const colSgst = findCol(customMapping?.sgstCol, [/sgst/, /state.*tax/, /utgst/]);
    const colCess = findCol(customMapping?.cessCol, [/cess/]);
    const colDueDate = findCol(customMapping?.dueDateCol, [/due\s*date/, /due\s*on/]);
    const colNarration = findCol(customMapping?.narrationCol, [/narration/, /remarks/]);

    // Identify annual comparative columns vs genuine monthly columns
    const annualCols: Array<{ colName: string; colIdx: number; endYear: number; label: string }> = [];
    const candidateMonthCols: Array<{ colName: string; yearMonth: string }> = [];

    rawHeaderRow.forEach((h, idx) => {
      if (h === colNote || h === colLedger || h === colGroup) return;
      const annual = parseAnnualColumnYear(h);
      if (annual) {
        annualCols.push({ colName: h, colIdx: idx, endYear: annual.endYear, label: annual.label });
      } else {
        const ym = parseHeaderToYearMonth(h);
        if (ym) {
          candidateMonthCols.push({ colName: h, yearMonth: ym });
        }
      }
    });

    // Only treat candidateMonthCols as genuine monthly columns if there are >= 3 distinct months
    // (Prevents 2 comparative annual columns like "Mar-26" and "Mar-25" from ever being treated as 2 months!)
    const monthColMap: Array<{ colName: string; yearMonth: string }> = [];
    if (candidateMonthCols.length >= 3) {
      hasAnyMonthlyGranularity = true;
      for (const mc of candidateMonthCols) {
        monthColMap.push(mc);
        const mStart = `${mc.yearMonth}-01`;
        const mEnd = `${mc.yearMonth}-28`;
        if (!minDate || mStart < minDate) minDate = mStart;
        if (!maxDate || mEnd > maxDate) maxDate = mEnd;
      }
    } else if (candidateMonthCols.length === 2 && annualCols.length === 0) {
      // Two columns like "Mar-26" and "Mar-25" 1 year apart are annual comparative columns!
      const y1 = Number(candidateMonthCols[0].yearMonth.slice(0, 4));
      const y2 = Number(candidateMonthCols[1].yearMonth.slice(0, 4));
      if (Math.abs(y1 - y2) === 1) {
        annualCols.push({
          colName: candidateMonthCols[0].colName,
          colIdx: rawHeaderRow.indexOf(candidateMonthCols[0].colName),
          endYear: y1,
          label: `FY ${y1 - 1}–${String(y1).slice(-2)}`,
        });
        annualCols.push({
          colName: candidateMonthCols[1].colName,
          colIdx: rawHeaderRow.indexOf(candidateMonthCols[1].colName),
          endYear: y2,
          label: `FY ${y2 - 1}–${String(y2).slice(-2)}`,
        });
      }
    }

    // Determine Current Period Column and Prior Period Column
    let colClosing = customMapping?.closingBalanceCol;
    let colPrior = customMapping?.priorPeriodCol;

    if (!colClosing && annualCols.length >= 1) {
      const sortedAnnual = [...annualCols].sort((a, b) => b.endYear - a.endYear);
      colClosing = sortedAnnual[0].colName;
      detectedCurrentPeriodLabel = sortedAnnual[0].label;
      const endYr = sortedAnnual[0].endYear;
      minDate = `${endYr - 1}-04-01`;
      maxDate = `${endYr}-03-31`;
      if (sortedAnnual.length >= 2) {
        colPrior = sortedAnnual[1].colName;
        detectedComparativePeriodLabel = sortedAnnual[1].label;
      }
    }

    if (!colClosing) {
      colClosing = findCol(undefined, [
        /closing\s*bal/,
        /^closing/,
        /current\s*period/,
        /current\s*year/,
        /current\s*month/,
        /actual/,
        /fy\s*2025/,
        /fy\s*25/,
        /pending\s*amount/,
        /balance\s*amount/,
        /^amount$/,
        /^inr$/,
        /^value$/,
      ]);
    }
    if (!colPrior) {
      colPrior = findCol(undefined, [
        /previous\s*period/,
        /prior\s*period/,
        /previous\s*year/,
        /prior\s*year/,
        /last\s*year/,
        /last\s*month/,
        /fy\s*2024/,
        /fy\s*24/,
        /comparative/,
      ]);
    }

    // Identify the best Particulars column(s)
    // In Schedule III, Col A may have Roman numerals (I, II, III) while Col B has Particulars, or vice versa
    const effectiveLedgerCol =
      colLedger ||
      rawHeaderRow.find((h, idx) => {
        if (idx > 2) return false;
        const textCount = dataRows.filter((r) => {
          const v = String(r.rowObj[h] ?? '').trim();
          return v.length > 3 && isNaN(Number(v));
        }).length;
        return textCount >= 3;
      }) ||
      rawHeaderRow[0];

    const ledgerColIdx = Math.max(0, rawHeaderRow.indexOf(effectiveLedgerCol));

    // Helper to extract the descriptive label from a row (checking both ledgerColIdx and adjacent text column if ledgerColIdx is a Roman numeral)
    const extractRowLabel = (rawCells: unknown[]): string => {
      const primary = String(rawCells[ledgerColIdx] ?? '').trim();
      if (
        (!primary || /^(?:[ivx]+|[0-9]+|[a-z]|\([a-z0-9ivx]+\))\.?$/i.test(primary)) &&
        ledgerColIdx + 1 < rawCells.length
      ) {
        const nextCell = String(rawCells[ledgerColIdx + 1] ?? '').trim();
        if (nextCell.length > 2 && isNaN(Number(nextCell))) {
          return primary ? `${primary}. ${nextCell}` : nextCell;
        }
      }
      if (!primary && ledgerColIdx > 0) {
        const prevCell = String(rawCells[ledgerColIdx - 1] ?? '').trim();
        if (prevCell.length > 2 && isNaN(Number(prevCell))) {
          return prevCell;
        }
      }
      // If primary is just a standalone number or bullet with no text next to it, treat as empty label
      if (/^(?:[ivx]+|[0-9]+|[a-z]|\([a-z0-9ivx]+\))\.?$/i.test(primary)) {
        return '';
      }
      return primary;
    };

    // Determine numeric columns after Particulars & Note No. if colClosing / colPrior were not matched by header name
    const noteColIdx = colNote ? rawHeaderRow.indexOf(colNote) : -1;
    const candidateNumericColIndices: number[] = [];
    for (let c = ledgerColIdx + 1; c < rawHeaderRow.length; c++) {
      const hName = rawHeaderRow[c];
      if (c === noteColIdx || hName === colGroup || hName === colBudget) continue;
      if (monthColMap.some((m) => m.colName === hName)) continue;

      // Check if this column looks like a Note No. column (only small integers 1..45 or note strings like "2", "3")
      const numVals = dataRows
        .map((r) => r.rawCells[c])
        .filter((v) => v !== '' && v !== null && v !== undefined && !isNaN(Number(String(v).replace(/,/g, ''))))
        .map((v) => Number(String(v).replace(/,/g, '')));

      if (numVals.length === 0) continue;
      const maxAbs = Math.max(...numVals.map(Math.abs));
      const allSmallSequentialInts =
        maxAbs <= 45 && numVals.every((n) => Number.isInteger(n) && n >= 1 && n <= 45);
      if (allSmallSequentialInts && c <= ledgerColIdx + 2) {
        continue; // Skip Note number column
      }
      candidateNumericColIndices.push(c);
    }

    const effectiveCurrentCol =
      colClosing ||
      (candidateNumericColIndices.length > 0
        ? rawHeaderRow[candidateNumericColIndices[0]]
        : monthColMap.length > 0
        ? monthColMap[monthColMap.length - 1].colName
        : undefined);

    const effectivePriorCol =
      colPrior ||
      (effectiveCurrentCol &&
      candidateNumericColIndices.length >= 2 &&
      rawHeaderRow[candidateNumericColIndices[0]] === effectiveCurrentCol
        ? rawHeaderRow[candidateNumericColIndices[1]]
        : undefined);

    // Determine sheet report type
    const sheetLower = sheetName.toLowerCase();
    const allParticularsLower = dataRows
      .slice(0, 40)
      .map((d) => extractRowLabel(d.rawCells).toLowerCase())
      .join(' ');

    let sheetType: DetectedReportType =
      customMapping?.reportCategoryOverride || 'UNKNOWN_UNSUPPORTED';

    if (sheetType === 'UNKNOWN_UNSUPPORTED') {
      if (bestBsSheet && sheetName === bestBsSheet.name) {
        sheetType = 'BALANCE_SHEET_EXCEL';
      } else if (bestPnlSheet && sheetName === bestPnlSheet.name) {
        sheetType = 'PROFIT_LOSS_EXCEL';
      } else if (bestCfSheet && sheetName === bestCfSheet.name) {
        sheetType = 'CASH_FLOW_EXCEL';
      } else if (
        !bestCfSheet &&
        (sheetLower.includes('cash flow') ||
          sheetLower.includes('cashflow') ||
          sheetLower === 'cfs' ||
          (allParticularsLower.includes('operating activities') &&
            allParticularsLower.includes('investing activities')))
      ) {
        sheetType = 'CASH_FLOW_EXCEL';
      } else if (
        !bestPnlSheet &&
        (sheetLower.includes('profit') ||
          sheetLower.includes('p&l') ||
          sheetLower.includes('pnl') ||
          sheetLower.includes('income') ||
          (allParticularsLower.includes('revenue from operations') &&
            (allParticularsLower.includes('profit before tax') ||
              allParticularsLower.includes('total expenses'))))
      ) {
        sheetType = 'PROFIT_LOSS_EXCEL';
      } else if (
        !bestBsSheet &&
        (sheetLower.includes('balance sheet') ||
          sheetLower.includes('balancesheet') ||
          sheetLower === 'bs' ||
          (allParticularsLower.includes('equity and liabilities') &&
            allParticularsLower.includes('assets')) ||
          (allParticularsLower.includes('share capital') &&
            allParticularsLower.includes('total assets')))
      ) {
        sheetType = 'BALANCE_SHEET_EXCEL';
      } else if (
        !bestCfSheet &&
        (lowerFilename.includes('cash flow') || lowerFilename.includes('cashflow'))
      ) {
        sheetType = 'CASH_FLOW_EXCEL';
      } else if (
        lowerFilename.includes('gstr1') ||
        lowerFilename.includes('gstr-1') ||
        (sheetLower.includes('b2b') && lowerFilename.includes('1'))
      ) {
        sheetType = 'GSTR1_RETURN';
      } else if (
        lowerFilename.includes('gstr2b') ||
        lowerFilename.includes('gstr-2b') ||
        sheetLower.includes('2b')
      ) {
        sheetType = 'GSTR2B_RETURN';
      } else if (
        lowerFilename.includes('gstr3b') ||
        lowerFilename.includes('gstr-3b') ||
        sheetLower.includes('3b')
      ) {
        sheetType = 'GSTR3B_RETURN';
      } else if (
        lowerFilename.includes('outstanding') ||
        lowerFilename.includes('ageing') ||
        lowerFilename.includes('receivable') ||
        lowerFilename.includes('payable') ||
        (colParty && colDueDate)
      ) {
        sheetType = 'OUTSTANDING_BILLS_EXCEL';
      } else if (colTaxable && (colIgst || colCgst || colSgst)) {
        if (lowerFilename.includes('purchase') || sheetLower.includes('purchase')) {
          sheetType = 'PURCHASE_REGISTER_EXCEL';
        } else if (lowerFilename.includes('sale') || sheetLower.includes('sale')) {
          sheetType = 'SALES_REGISTER_EXCEL';
        } else {
          sheetType = 'GST_REGISTER_EXCEL';
        }
      } else if (colDate && (colVchNo || colParty) && (colDebit || colCredit || effectiveCurrentCol)) {
        if (lowerFilename.includes('sale') || sheetLower.includes('sale')) {
          sheetType = 'SALES_REGISTER_EXCEL';
        } else if (lowerFilename.includes('purchase') || sheetLower.includes('purchase')) {
          sheetType = 'PURCHASE_REGISTER_EXCEL';
        } else {
          sheetType = 'DAYBOOK_EXCEL';
        }
      } else if (
        !hasAuthoritativeFinancialStatements &&
        effectiveLedgerCol &&
        (effectiveCurrentCol || colDebit || colCredit || colOpening || monthColMap.length > 0)
      ) {
        if (
          lowerFilename.includes('profit') ||
          lowerFilename.includes('p&l') ||
          lowerFilename.includes('pnl') ||
          (allParticularsLower.includes('revenue') &&
            (allParticularsLower.includes('gross profit') ||
              allParticularsLower.includes('ebitda') ||
              allParticularsLower.includes('net profit') ||
              allParticularsLower.includes('profit after tax')))
        ) {
          sheetType = 'PROFIT_LOSS_EXCEL';
        } else if (
          lowerFilename.includes('balance sheet') ||
          lowerFilename.includes('balancesheet') ||
          (allParticularsLower.includes('assets') &&
            (allParticularsLower.includes('liabilities') || allParticularsLower.includes('equity')))
        ) {
          sheetType = 'BALANCE_SHEET_EXCEL';
        } else if (sheetLower.includes('budget')) {
          sheetType = 'BUDGET_EXCEL';
        } else if (
          sheetLower.includes('trial') ||
          lowerFilename.includes('trial') ||
          (colDebit && colCredit)
        ) {
          sheetType = 'TRIAL_BALANCE_EXCEL';
        }
      }
    }

    if (sheetType !== 'UNKNOWN_UNSUPPORTED') {
      detectedSheetTypes.add(sheetType);
      if (primaryReportType === 'UNKNOWN_UNSUPPORTED') {
        primaryReportType = sheetType;
      }
    }

    const currColIdx = effectiveCurrentCol ? rawHeaderRow.indexOf(effectiveCurrentCol) : -1;
    const priorColIdx = effectivePriorCol ? rawHeaderRow.indexOf(effectivePriorCol) : -1;
    const currColLetter = XLSX.utils.encode_col(Math.max(0, currColIdx >= 0 ? currColIdx : 1));
    const priorColLetter =
      priorColIdx >= 0 ? XLSX.utils.encode_col(priorColIdx) : undefined;

    // ------------------------------------------------------------------------
    // CASE 1: CASH FLOW STATEMENT SHEET (Section 7)
    // ------------------------------------------------------------------------
    if (sheetType === 'CASH_FLOW_EXCEL') {
      let currentSection: 'OPERATING' | 'INVESTING' | 'FINANCING' = 'OPERATING';
      let reachedClosingCash = false;

      for (let i = 0; i < dataRows.length; i++) {
        if (reachedClosingCash) break;
        const { rowIndex1Based, rawCells, rowObj } = dataRows[i];
        const rawItemName = extractRowLabel(rawCells);
        if (!rawItemName || isSignatureOrFooterRow(rawItemName)) continue;

        const lowerName = rawItemName.toLowerCase();
        if (lowerName.includes('investing activities')) currentSection = 'INVESTING';
        else if (lowerName.includes('financing activities')) currentSection = 'FINANCING';
        else if (lowerName.includes('operating activities')) currentSection = 'OPERATING';

        // Support multi-column Cash Flow layout: check effectiveCurrentCol first; if empty on a detail row, check inner numeric columns
        let activeRowCurrColIdx = currColIdx;
        let currentPaise =
          currColIdx >= 0 ? toPaise(rawCells[currColIdx] as string | number | null | undefined) : 0;

        if (currentPaise === 0 && candidateNumericColIndices.length > 1) {
          // Check if value was placed in an inner detail column before currColIdx
          for (const cIdx of candidateNumericColIndices) {
            if (cIdx === priorColIdx) continue;
            const v = toPaise(rawCells[cIdx] as string | number | null | undefined);
            if (v !== 0) {
              currentPaise = v;
              activeRowCurrColIdx = cIdx;
              break;
            }
          }
        }

        const priorPaise =
          priorColIdx >= 0 && rawCells[priorColIdx] !== ''
            ? toPaise(rawCells[priorColIdx] as string | number | null | undefined)
            : undefined;
        const budgetPaise = colBudget && rowObj[colBudget] !== '' ? toPaise(rowObj[colBudget]) : undefined;

        const monthlyValuesPaise: Record<string, number> = {};
        let sumMonthlyPaise = 0;
        for (const mc of monthColMap) {
          const mp = toPaise(rowObj[mc.colName]);
          monthlyValuesPaise[mc.yearMonth] = mp;
          sumMonthlyPaise += mp;
        }

        const resolvedCurrentPaise =
          currentPaise !== 0
            ? currentPaise
            : monthColMap.length > 0
            ? sumMonthlyPaise
            : 0;

        const { category, isSubtotal } = classifyCashFlowRow(rawItemName, currentSection);

        // Skip pure section headers with zero values across all columns
        if (
          resolvedCurrentPaise === 0 &&
          (!priorPaise || priorPaise === 0) &&
          !isSubtotal
        ) {
          continue;
        }

        if (
          resolvedCurrentPaise === 0 &&
          (!priorPaise || priorPaise === 0) &&
          (lowerName.endsWith('activities') || lowerName.endsWith('activities:'))
        ) {
          continue;
        }

        const activeColLetter = XLSX.utils.encode_col(Math.max(0, activeRowCurrColIdx));
        const cellRef = `'${sheetName}'!${activeColLetter}${rowIndex1Based}`;
        const priorCellRef = priorColLetter
          ? `'${sheetName}'!${priorColLetter}${rowIndex1Based}`
          : undefined;

        cashFlowItems.push({
          id: `${sourceFileId}_CF_${sheetName}_${i + 1}`,
          particulars: stripLeadingBulletOrRoman(rawItemName) || rawItemName,
          normalizedParticulars: normalizeName(stripLeadingBulletOrRoman(rawItemName) || rawItemName),
          category,
          currentPeriodPaise: resolvedCurrentPaise,
          previousPeriodPaise: priorPaise,
          budgetPaise,
          monthlyValuesPaise: monthColMap.length > 0 ? monthlyValuesPaise : undefined,
          isSubtotalOrTotal: isSubtotal,
          sourceFileId,
          sourceFilename,
          worksheetName: sheetName,
          cellReference: cellRef,
          priorCellReference: priorCellRef,
          reportingPeriodLabel: detectedCurrentPeriodLabel || 'FY 2025–26',
          comparativePeriodLabel: detectedComparativePeriodLabel || 'FY 2024–25',
        });

        if (isSubtotal) {
          reportedSubtotals.push({
            id: `${sourceFileId}_SUB_CF_${i + 1}`,
            statementType: 'CASH_FLOW',
            label: stripLeadingBulletOrRoman(rawItemName) || rawItemName,
            normalizedLabel: normalizeName(stripLeadingBulletOrRoman(rawItemName) || rawItemName),
            reportedCurrentPaise: resolvedCurrentPaise,
            reportedPreviousPaise: priorPaise,
            reportedBudgetPaise: budgetPaise,
            sourceFileId,
            sourceFilename,
            worksheetName: sheetName,
            cellReference: cellRef,
            priorCellReference: priorCellRef,
            reportingPeriodLabel: detectedCurrentPeriodLabel || 'FY 2025–26',
            comparativePeriodLabel: detectedComparativePeriodLabel || 'FY 2024–25',
          });
        }

        if (category === 'CF_CLOSING_CASH' && resolvedCurrentPaise !== 0) {
          reachedClosingCash = true;
        }
      }
    }
    // ------------------------------------------------------------------------
    // CASE 2: P&L, BALANCE SHEET, TRIAL BALANCE, OR BUDGET SHEET (Sections 4, 5, 6)
    // ------------------------------------------------------------------------
    else if (
      sheetType === 'TRIAL_BALANCE_EXCEL' ||
      sheetType === 'PROFIT_LOSS_EXCEL' ||
      sheetType === 'BALANCE_SHEET_EXCEL' ||
      sheetType === 'BUDGET_EXCEL'
    ) {
      const statementContext: 'PNL' | 'BALANCE_SHEET' | 'TRIAL_BALANCE' =
        sheetType === 'BALANCE_SHEET_EXCEL'
          ? 'BALANCE_SHEET'
          : sheetType === 'PROFIT_LOSS_EXCEL'
          ? 'PNL'
          : 'TRIAL_BALANCE';

      let activeSectionGroup = '';
      let activeScheduleIIIParentCategory:
        | 'Equity'
        | 'LongTermBorrowings'
        | 'ShortTermBorrowings'
        | 'TradePayables'
        | 'OtherCurrentLiabilities'
        | 'FixedAssets'
        | 'Investments'
        | 'NonCurrentAssets'
        | 'Inventory'
        | 'TradeReceivables'
        | 'CashAndEquivalents'
        | 'BankBalances'
        | 'OtherCurrentAssets'
        | null = null;
      let activeScheduleIIIParentTitle = '';
      let bsTopDivision: 'NONE' | 'EQUITY_AND_LIABILITIES' | 'ASSETS' = 'NONE';
      let reachedEndOfStatement = false;
      let seenEquityAndLiabilitiesTotal = false;

      let sectionRunningSumPaise = 0;
      let eqLiabRunningSumPaise = 0;
      let assetsRunningSumPaise = 0;

      const detectScheduleIIIParentHeading = (
        lowerClean: string,
        sectionGroup: string
      ): {
        category:
          | 'Equity'
          | 'LongTermBorrowings'
          | 'ShortTermBorrowings'
          | 'TradePayables'
          | 'OtherCurrentLiabilities'
          | 'FixedAssets'
          | 'Investments'
          | 'NonCurrentAssets'
          | 'Inventory'
          | 'TradeReceivables'
          | 'CashAndEquivalents'
          | 'BankBalances'
          | 'OtherCurrentAssets';
        title: string;
      } | null => {
        // Never treat Trade Payables MSME/Creditor child lines as a new parent heading
        if (
          lowerClean.includes('total outstanding dues') ||
          lowerClean.includes('outstanding dues of') ||
          lowerClean.includes('dues of micro') ||
          lowerClean.includes('dues of creditor') ||
          lowerClean.includes('micro enterprises') ||
          lowerClean.includes('creditors other than')
        ) {
          return null;
        }

        if (sectionGroup === 'Current Liabilities') {
          if (
            lowerClean === 'trade payables' ||
            lowerClean.startsWith('trade payables') ||
            lowerClean === 'trade payable' ||
            lowerClean === 'sundry creditors'
          ) {
            return { category: 'TradePayables', title: 'Trade Payables' };
          }
          if (
            lowerClean === 'short-term borrowings' ||
            lowerClean === 'short term borrowings' ||
            lowerClean.startsWith('short-term borrowings') ||
            lowerClean.startsWith('short term borrowings')
          ) {
            return { category: 'ShortTermBorrowings', title: 'Short-Term Borrowings' };
          }
          if (
            lowerClean === 'other current liabilities' ||
            lowerClean.startsWith('other current liabilities')
          ) {
            return { category: 'OtherCurrentLiabilities', title: 'Other Current Liabilities' };
          }
          if (
            lowerClean === 'short-term provisions' ||
            lowerClean === 'short term provisions' ||
            lowerClean.startsWith('short-term provisions') ||
            lowerClean.startsWith('short term provisions')
          ) {
            return { category: 'OtherCurrentLiabilities', title: 'Short-Term Provisions' };
          }
        } else if (sectionGroup === 'Non-Current Liabilities') {
          if (
            lowerClean.startsWith('long-term borrowings') ||
            lowerClean.startsWith('long term borrowings') ||
            lowerClean.startsWith('deferred tax liabilities') ||
            lowerClean.startsWith('other long-term liabilities') ||
            lowerClean.startsWith('other long term liabilities') ||
            lowerClean.startsWith('long-term provisions') ||
            lowerClean.startsWith('long term provisions')
          ) {
            return { category: 'LongTermBorrowings', title: 'Non-Current Liabilities' };
          }
        } else if (sectionGroup === "Shareholders' Funds") {
          if (
            lowerClean.startsWith('share capital') ||
            lowerClean.startsWith('reserves and surplus') ||
            lowerClean.startsWith('reserves & surplus') ||
            lowerClean.startsWith('money received against share warrants')
          ) {
            return { category: 'Equity', title: "Shareholders' Funds" };
          }
        } else if (sectionGroup === 'Non-Current Assets') {
          if (
            lowerClean.includes('property, plant') ||
            lowerClean.includes('property plant') ||
            lowerClean === 'fixed assets' ||
            lowerClean.startsWith('fixed assets')
          ) {
            return { category: 'FixedAssets', title: 'Property, Plant and Equipment' };
          }
          if (
            lowerClean.startsWith('non-current investments') ||
            lowerClean.startsWith('non current investments')
          ) {
            return { category: 'Investments', title: 'Non-Current Investments' };
          }
          if (
            lowerClean.startsWith('deferred tax assets') ||
            lowerClean.startsWith('long-term loans and advances') ||
            lowerClean.startsWith('long term loans and advances') ||
            lowerClean.startsWith('other non-current assets') ||
            lowerClean.startsWith('other non current assets')
          ) {
            return { category: 'NonCurrentAssets', title: 'Non-Current Assets' };
          }
        } else if (sectionGroup === 'Current Assets') {
          if (lowerClean.startsWith('inventories') || lowerClean === 'inventory') {
            return { category: 'Inventory', title: 'Inventories' };
          }
          if (
            lowerClean.startsWith('trade receivables') ||
            lowerClean === 'trade receivable' ||
            lowerClean === 'sundry debtors'
          ) {
            return { category: 'TradeReceivables', title: 'Trade Receivables' };
          }
          if (
            lowerClean.startsWith('cash and cash equivalents') ||
            lowerClean.startsWith('cash & cash equivalents')
          ) {
            return { category: 'CashAndEquivalents', title: 'Cash and Cash Equivalents' };
          }
          if (
            lowerClean.startsWith('bank balances other than') ||
            lowerClean.startsWith('other bank balances')
          ) {
            return { category: 'BankBalances', title: 'Bank Balances' };
          }
          if (
            lowerClean.startsWith('current investments') ||
            lowerClean.startsWith('short-term loans and advances') ||
            lowerClean.startsWith('short term loans and advances') ||
            lowerClean.startsWith('other current assets')
          ) {
            return { category: 'OtherCurrentAssets', title: 'Other Current Assets' };
          }
        }
        return null;
      };

      const isExplicitZeroCell = (cellVal: unknown): boolean => {
        if (cellVal === 0) return true;
        if (typeof cellVal === 'string') {
          const t = cellVal.trim().toLowerCase();
          return (
            t === '0' ||
            t === '0.0' ||
            t === '0.00' ||
            t === '-' ||
            t === '–' ||
            t === '—' ||
            t === 'nil' ||
            t === '-nil-'
          );
        }
        return false;
      };

      const sheetExtractedRows: Array<{
        ledger: Ledger;
        rawCurrentPaise: number;
        rawPriorPaise?: number;
        isSectionHeaderRow?: boolean;
        isScheduleIIIParentHeading?: boolean;
      }> = [];

      for (let i = 0; i < dataRows.length; i++) {
        if (reachedEndOfStatement) break;
        const { rowIndex1Based, rawCells, rowObj: row } = dataRows[i];
        const rawLedgerName = extractRowLabel(rawCells);
        if (!rawLedgerName) continue;

        // If we encounter Notes or Signature blocks below the statement on the same worksheet, stop immediately
        const rawLowerTrim = rawLedgerName.trim().toLowerCase();
        if (
          (statementContext === 'BALANCE_SHEET' || statementContext === 'PNL') &&
          (rawLowerTrim.includes('notes forming part') ||
            rawLowerTrim.includes('significant accounting policies') ||
            rawLowerTrim.includes('for and on behalf of') ||
            rawLowerTrim.includes('as per our report') ||
            /^(?:note|schedule)\s*(?:no\.?)?\s*[:\-–]?\s*\d+\b/i.test(rawLowerTrim))
        ) {
          if (sheetExtractedRows.length > 0) {
            reachedEndOfStatement = true;
            break;
          }
          continue;
        }

        if (isSignatureOrFooterRow(rawLedgerName)) continue;

        const cleanName = stripLeadingBulletOrRoman(rawLedgerName) || rawLedgerName;
        const lowerClean = cleanName.toLowerCase();
        if (!cleanName || /^[\d\s.,()\-–—/:]+$/.test(cleanName)) continue;

        let isScheduleIIISectionHeader = false;

        // Update active Schedule III section context when encountering section headers
        if (statementContext === 'BALANCE_SHEET') {
          if (
            lowerClean === 'equity and liabilities' ||
            lowerClean === 'equity & liabilities' ||
            lowerClean.includes('equity and liabilities') ||
            lowerClean.includes('equity & liabilities') ||
            lowerClean === 'sources of funds'
          ) {
            if (!lowerClean.startsWith('total')) {
              bsTopDivision = 'EQUITY_AND_LIABILITIES';
              activeScheduleIIIParentCategory = null;
              activeScheduleIIIParentTitle = '';
              isScheduleIIISectionHeader = true;
            }
          } else if (
            (lowerClean.includes('shareholders') ||
              lowerClean.includes("shareholder's") ||
              lowerClean === 'equity' ||
              lowerClean.includes('share application money')) &&
            !lowerClean.startsWith('total')
          ) {
            bsTopDivision = 'EQUITY_AND_LIABILITIES';
            if (activeSectionGroup !== "Shareholders' Funds") {
              activeSectionGroup = "Shareholders' Funds";
              activeScheduleIIIParentCategory = null;
              activeScheduleIIIParentTitle = '';
              sectionRunningSumPaise = 0;
            }
            if (
              lowerClean === "shareholders' funds" ||
              lowerClean === 'shareholders funds' ||
              lowerClean === "shareholder's funds" ||
              lowerClean === 'equity'
            ) {
              isScheduleIIISectionHeader = true;
            }
          } else if (
            (lowerClean.includes('non-current liabilit') ||
              lowerClean.includes('non current liabilit')) &&
            !lowerClean.startsWith('total') &&
            !lowerClean.includes('other non-current liabilit') &&
            !lowerClean.includes('other non current liabilit')
          ) {
            bsTopDivision = 'EQUITY_AND_LIABILITIES';
            if (activeSectionGroup !== 'Non-Current Liabilities') {
              activeSectionGroup = 'Non-Current Liabilities';
              activeScheduleIIIParentCategory = null;
              activeScheduleIIIParentTitle = '';
              sectionRunningSumPaise = 0;
            }
            if (
              lowerClean === 'non-current liabilities' ||
              lowerClean === 'non current liabilities'
            ) {
              isScheduleIIISectionHeader = true;
            }
          } else if (
            (lowerClean === 'current liabilities' ||
              lowerClean.startsWith('current liabilities')) &&
            !lowerClean.includes('non-current') &&
            !lowerClean.includes('non current') &&
            !lowerClean.includes('other current liabilit') &&
            !lowerClean.startsWith('total')
          ) {
            bsTopDivision = 'EQUITY_AND_LIABILITIES';
            if (activeSectionGroup !== 'Current Liabilities') {
              activeSectionGroup = 'Current Liabilities';
              activeScheduleIIIParentCategory = null;
              activeScheduleIIIParentTitle = '';
              sectionRunningSumPaise = 0;
            }
            isScheduleIIISectionHeader = true;
          } else if (
            (lowerClean === 'assets' || lowerClean === 'application of funds') &&
            !lowerClean.startsWith('total')
          ) {
            bsTopDivision = 'ASSETS';
            if (activeSectionGroup !== 'Non-Current Assets') {
              activeSectionGroup = 'Non-Current Assets';
              activeScheduleIIIParentCategory = null;
              activeScheduleIIIParentTitle = '';
              sectionRunningSumPaise = 0;
            }
            isScheduleIIISectionHeader = true;
          } else if (
            (lowerClean.includes('non-current asset') ||
              lowerClean.includes('non current asset') ||
              lowerClean === 'fixed assets') &&
            !lowerClean.startsWith('total') &&
            !lowerClean.includes('other non-current asset') &&
            !lowerClean.includes('other non current asset')
          ) {
            bsTopDivision = 'ASSETS';
            if (activeSectionGroup !== 'Non-Current Assets') {
              activeSectionGroup = 'Non-Current Assets';
              activeScheduleIIIParentCategory = null;
              activeScheduleIIIParentTitle = '';
              sectionRunningSumPaise = 0;
            }
            if (
              lowerClean === 'non-current assets' ||
              lowerClean === 'non current assets' ||
              lowerClean === 'fixed assets'
            ) {
              isScheduleIIISectionHeader = true;
            }
          } else if (
            (lowerClean === 'current assets' || lowerClean.startsWith('current assets')) &&
            !lowerClean.includes('non-current') &&
            !lowerClean.includes('non current') &&
            !lowerClean.includes('other current asset') &&
            !lowerClean.startsWith('total')
          ) {
            bsTopDivision = 'ASSETS';
            if (activeSectionGroup !== 'Current Assets') {
              activeSectionGroup = 'Current Assets';
              activeScheduleIIIParentCategory = null;
              activeScheduleIIIParentTitle = '';
              sectionRunningSumPaise = 0;
            }
            isScheduleIIISectionHeader = true;
          }
        } else if (statementContext === 'PNL') {
          if (lowerClean === 'income' || lowerClean === 'revenue') {
            activeSectionGroup = 'Income';
          } else if (lowerClean === 'expenses' || lowerClean === 'expenditure') {
            activeSectionGroup = 'Expenses';
          } else if (lowerClean.includes('tax expense')) {
            activeSectionGroup = 'Tax Expense';
          }
        }

        // Check if this row is a Schedule III primary category heading (e.g. "(a) Trade payables", "(a) Property, Plant and Equipment")
        let isScheduleIIIParentHeading = false;
        if (statementContext === 'BALANCE_SHEET' && !isScheduleIIISectionHeader) {
          const detectedParent = detectScheduleIIIParentHeading(lowerClean, activeSectionGroup);
          if (detectedParent) {
            activeScheduleIIIParentCategory = detectedParent.category;
            activeScheduleIIIParentTitle = detectedParent.title;
            isScheduleIIIParentHeading = true;
          }
        }

        const explicitGroup = colGroup ? String(row[colGroup] || '').trim() : '';
        const parentGroup =
          explicitGroup && explicitGroup.toLowerCase() !== 'subtotal'
            ? explicitGroup
            : activeSectionGroup || 'Primary';

        const openingPaise = colOpening ? toPaise(row[colOpening]) : 0;
        const drPaise = colDebit ? Math.abs(toPaise(row[colDebit])) : 0;
        const crPaise = colCredit ? Math.abs(toPaise(row[colCredit])) : 0;

        const currCellRaw = currColIdx >= 0 ? rawCells[currColIdx] : undefined;
        const priorCellRaw = priorColIdx >= 0 ? rawCells[priorColIdx] : undefined;
        const hasExplicitZeroCurr = isExplicitZeroCell(currCellRaw);
        const hasExplicitZeroPrior = isExplicitZeroCell(priorCellRaw);

        const rawPriorPaise =
          priorColIdx >= 0 && priorCellRaw !== '' && priorCellRaw !== null && priorCellRaw !== undefined
            ? toPaise(priorCellRaw as string | number | null | undefined)
            : hasExplicitZeroPrior
            ? 0
            : undefined;
        const budgetPaise =
          colBudget && row[colBudget] !== '' ? toPaise(row[colBudget]) : undefined;

        const monthlyReportedPaise: Record<string, number> = {};
        let sumMonthlyPaise = 0;
        for (const mc of monthColMap) {
          const mp = toPaise(row[mc.colName]);
          monthlyReportedPaise[mc.yearMonth] = mp;
          sumMonthlyPaise += mp;
        }

        let rawCurrentPaise = 0;
        if (currColIdx >= 0 && currCellRaw !== '' && currCellRaw !== null && currCellRaw !== undefined) {
          rawCurrentPaise = toPaise(currCellRaw as string | number | null | undefined);
        } else if (colDebit || colCredit) {
          rawCurrentPaise = openingPaise + drPaise - crPaise;
        } else if (monthColMap.length > 0) {
          rawCurrentPaise = sumMonthlyPaise;
        }

        const cellRef = `'${sheetName}'!${currColLetter}${rowIndex1Based}`;
        const priorCellRef = priorColLetter
          ? `'${sheetName}'!${priorColLetter}${rowIndex1Based}`
          : undefined;

        const isExplicitSubtotal =
          explicitGroup.toLowerCase() === 'subtotal' || isStatementSubtotalRow(rawLedgerName);

        if (isExplicitSubtotal) {
          if (rawCurrentPaise !== 0 || (rawPriorPaise !== undefined && rawPriorPaise !== 0)) {
            let normLabel = cleanLabelForSubtotalCheck(rawLedgerName);

            // Check if the next non-empty row in dataRows is ALSO a subtotal row (e.g. Section Total followed immediately by Grand Total)
            let isNextRowAlsoSubtotal = false;
            for (let nextIdx = i + 1; nextIdx < Math.min(i + 3, dataRows.length); nextIdx++) {
              const nextLabel = extractRowLabel(dataRows[nextIdx].rawCells);
              if (!nextLabel) continue;
              if (isStatementSubtotalRow(nextLabel)) {
                isNextRowAlsoSubtotal = true;
              }
              break;
            }

            // Disambiguate generic "Total" rows on Balance Sheet & P&L using section context, running sums & lookahead
            if (normLabel === 'total' || normLabel === 'sub total' || normLabel === 'subtotal' || normLabel === 'sub-total') {
              if (statementContext === 'BALANCE_SHEET') {
                if (bsTopDivision === 'ASSETS') {
                  if (activeSectionGroup === 'Non-Current Assets') {
                    normLabel = 'total non-current assets';
                  } else if (activeSectionGroup === 'Current Assets') {
                    const closerToSection =
                      assetsRunningSumPaise > sectionRunningSumPaise &&
                      Math.abs(rawCurrentPaise - sectionRunningSumPaise) <
                        Math.abs(rawCurrentPaise - assetsRunningSumPaise);
                    if (isNextRowAlsoSubtotal || closerToSection) {
                      normLabel = 'total current assets';
                    } else {
                      normLabel = 'total assets';
                      reachedEndOfStatement = true;
                    }
                  } else {
                    normLabel = 'total assets';
                    reachedEndOfStatement = true;
                  }
                } else {
                  if (activeSectionGroup === "Shareholders' Funds") {
                    normLabel = 'total shareholders funds';
                  } else if (activeSectionGroup === 'Non-Current Liabilities') {
                    normLabel = 'total non-current liabilities';
                  } else if (activeSectionGroup === 'Current Liabilities') {
                    const closerToSection =
                      eqLiabRunningSumPaise > sectionRunningSumPaise &&
                      Math.abs(rawCurrentPaise - sectionRunningSumPaise) <
                        Math.abs(rawCurrentPaise - eqLiabRunningSumPaise);
                    if (isNextRowAlsoSubtotal || closerToSection) {
                      normLabel = 'total current liabilities';
                    } else {
                      normLabel = 'total equity and liabilities';
                      seenEquityAndLiabilitiesTotal = true;
                    }
                  } else if (!seenEquityAndLiabilitiesTotal) {
                    normLabel = 'total equity and liabilities';
                    seenEquityAndLiabilitiesTotal = true;
                  }
                }
              }
            } else {
              if (
                normLabel.includes('total equity and liabilities') ||
                normLabel.includes('total equity & liabilities') ||
                normLabel.includes('total liabilities and equity') ||
                normLabel.includes('total capital and liabilities')
              ) {
                normLabel = 'total equity and liabilities';
                seenEquityAndLiabilitiesTotal = true;
              } else if (
                statementContext === 'BALANCE_SHEET' &&
                (normLabel === 'total assets' ||
                  normLabel === 'total of assets' ||
                  normLabel === 'total asset' ||
                  (bsTopDivision === 'ASSETS' && normLabel === 'grand total'))
              ) {
                normLabel = 'total assets';
                reachedEndOfStatement = true;
              } else if (
                statementContext === 'PNL' &&
                (normLabel.includes('profit after tax') ||
                  normLabel.includes('profit for the year') ||
                  normLabel.includes('profit for the period') ||
                  normLabel.includes('profit / (loss) for the period') ||
                  normLabel.includes('profit/(loss) for the period') ||
                  normLabel.includes('profit/(loss) for the year'))
              ) {
                reachedEndOfStatement = true;
              }
            }

            reportedSubtotals.push({
              id: `${sourceFileId}_SUB_${sheetName}_${i + 1}`,
              statementType: statementContext === 'BALANCE_SHEET' ? 'BALANCE_SHEET' : 'PNL',
              label: cleanName,
              normalizedLabel: normLabel,
              reportedCurrentPaise: rawCurrentPaise,
              reportedPreviousPaise: rawPriorPaise,
              reportedBudgetPaise: budgetPaise,
              sourceFileId,
              sourceFilename,
              worksheetName: sheetName,
              cellReference: cellRef,
              priorCellReference: priorCellRef,
              reportingPeriodLabel: detectedCurrentPeriodLabel || 'FY 2025–26',
              comparativePeriodLabel: detectedComparativePeriodLabel || 'FY 2024–25',
            });
          }
          continue;
        }

        const isTradePayablesChildRow =
          statementContext === 'BALANCE_SHEET' &&
          activeScheduleIIIParentCategory === 'TradePayables' &&
          !isScheduleIIIParentHeading &&
          (lowerClean.includes('total outstanding dues') ||
            lowerClean.includes('outstanding dues of') ||
            lowerClean.includes('dues of micro') ||
            lowerClean.includes('dues of creditor') ||
            lowerClean.includes('micro enterprises') ||
            lowerClean.includes('creditors other than') ||
            lowerClean.includes('msme'));

        // Skip pure section header rows or blank parent heading rows with no numerical amounts,
        // EXCEPT explicit zero / nil child rows under Trade Payables (e.g. MSME = ₹0.00 / '-') so the full parent-child breakdown is preserved!
        if (
          rawCurrentPaise === 0 &&
          openingPaise === 0 &&
          drPaise === 0 &&
          crPaise === 0 &&
          (rawPriorPaise === undefined || rawPriorPaise === 0) &&
          (budgetPaise === undefined || budgetPaise === 0)
        ) {
          if (
            isTradePayablesChildRow &&
            (hasExplicitZeroCurr || hasExplicitZeroPrior || lowerClean.includes('micro'))
          ) {
            // Preserve explicit zero MSME child row under Trade Payables for complete Schedule III audit trail
          } else {
            continue;
          }
        }

        let classification = classifyLedgerAndGroup(
          cleanName,
          parentGroup,
          undefined,
          statementContext
        );

        // Inherit active Schedule III parent category for child sub-rows (e.g. child rows under "(a) Trade payables", "(a) Property, Plant and Equipment", etc.)
        if (
          statementContext === 'BALANCE_SHEET' &&
          activeScheduleIIIParentCategory &&
          !isScheduleIIIParentHeading
        ) {
          if (
            activeScheduleIIIParentCategory === 'TradePayables' &&
            activeSectionGroup === 'Current Liabilities'
          ) {
            classification = 'TradePayables';
          } else if (
            activeScheduleIIIParentCategory === 'ShortTermBorrowings' &&
            activeSectionGroup === 'Current Liabilities' &&
            classification === 'OtherCurrentLiabilities'
          ) {
            classification = 'ShortTermBorrowings';
          } else if (
            activeScheduleIIIParentCategory === 'FixedAssets' &&
            activeSectionGroup === 'Non-Current Assets' &&
            classification === 'NonCurrentAssets'
          ) {
            classification = 'FixedAssets';
          }
        }

        const isCreditNature = [
          'Sales',
          'OtherOperatingIncome',
          'OtherIncome',
          'Equity',
          'LongTermBorrowings',
          'ShortTermBorrowings',
          'TradePayables',
          'TaxLiabilities',
          'OtherCurrentLiabilities',
        ].includes(classification);

        let closingPaise = rawCurrentPaise;
        // Apply natural accounting sign (+Dr / -Cr) when statement numbers are reported in standard unsigned/signed schedule convention
        if (
          currColIdx >= 0 &&
          !String(rawCells[currColIdx]).toLowerCase().includes('dr') &&
          !String(rawCells[currColIdx]).toLowerCase().includes('cr') &&
          !colDebit &&
          !colCredit
        ) {
          if (isCreditNature) {
            // Positive in Equity/Liability/Revenue schedule = Credit (-); Negative in Equity (e.g. deficit in Reserves) = Debit (+)
            closingPaise = -rawCurrentPaise;
          }
        }

        const effectiveRawPriorPaise =
          rawPriorPaise !== undefined
            ? rawPriorPaise
            : isTradePayablesChildRow && priorColIdx >= 0
            ? 0
            : undefined;

        let signedPriorPaise = effectiveRawPriorPaise;
        if (
          signedPriorPaise !== undefined &&
          isCreditNature &&
          !String(rawCells[priorColIdx] || '').toLowerCase().includes('dr')
        ) {
          signedPriorPaise = -effectiveRawPriorPaise!;
        }

        const effectiveOpeningPaise =
          openingPaise !== 0
            ? openingPaise
            : signedPriorPaise !== undefined && sheetType === 'BALANCE_SHEET_EXCEL'
            ? signedPriorPaise
            : 0;

        if (!isScheduleIIISectionHeader) {
          sectionRunningSumPaise += rawCurrentPaise;
          if (bsTopDivision === 'ASSETS') {
            assetsRunningSumPaise += rawCurrentPaise;
          } else {
            eqLiabRunningSumPaise += rawCurrentPaise;
          }
        }

        sheetExtractedRows.push({
          rawCurrentPaise,
          rawPriorPaise: effectiveRawPriorPaise,
          isSectionHeaderRow: isScheduleIIISectionHeader,
          isScheduleIIIParentHeading,
          ledger: {
            id: `${sourceFileId}_TB_${sheetName}_${i + 1}`,
            name:
              activeScheduleIIIParentTitle &&
              !isScheduleIIIParentHeading &&
              activeScheduleIIIParentCategory === 'TradePayables' &&
              !cleanName.toLowerCase().includes('trade payable')
                ? `${activeScheduleIIIParentTitle} — ${cleanName}`
                : cleanName,
            normalizedName: normalizeName(cleanName),
            parentGroup,
            classification,
            openingBalancePaise: effectiveOpeningPaise,
            closingBalanceReportedPaise: closingPaise,
            // Preserve true statement sign in previousPeriodPaise (so negative Deferred Tax Credit or negative Reserve is preserved!)
            previousPeriodPaise: effectiveRawPriorPaise,
            previousPeriodSignedPaise: signedPriorPaise,
            budgetPaise: budgetPaise,
            monthlyReportedPaise: monthColMap.length > 0 ? monthlyReportedPaise : undefined,
            worksheetName: sheetName,
            cellReference: cellRef,
            priorCellReference: priorCellRef,
            reportingPeriodLabel: detectedCurrentPeriodLabel || 'FY 2025–26',
            comparativePeriodLabel: detectedComparativePeriodLabel || 'FY 2024–25',
            statementSourceType: statementContext,
            currency: detectedCurrency,
            tbDebitPaise: drPaise || (closingPaise > 0 ? closingPaise : 0),
            tbCreditPaise: crPaise || (closingPaise < 0 ? Math.abs(closingPaise) : 0),
            sourceFileIds: [sourceFileId],
          },
        });
      }

      // Check for Schedule III section header rows & parent-child subtotal rows on the sheet
      for (let i = 0; i < sheetExtractedRows.length; i++) {
        const item = sheetExtractedRows[i];
        let isParentOfNextRows = false;

        // 1. If this row is a Schedule III section header (e.g. "Shareholders' Funds", "Current Assets") and child rows exist in the same section, treat it as a reported subtotal
        if (item.isSectionHeaderRow) {
          const hasChildrenInSection = sheetExtractedRows.some(
            (other, idx) =>
              idx !== i &&
              !other.isSectionHeaderRow &&
              other.ledger.parentGroup === item.ledger.parentGroup
          );
          if (hasChildrenInSection) {
            isParentOfNextRows = true;
          }
        }

        // 2. Check if this row is a parent heading (e.g. "(a) Property, Plant and Equipment and Intangible Assets" or "(a) Trade payables") whose value equals the sum of the next 1..6 child rows in the same parentGroup
        if (!isParentOfNextRows && i + 1 < sheetExtractedRows.length) {
          let sumChildrenCurr = 0;
          let sumChildrenPrior = 0;
          for (let k = i + 1; k < Math.min(i + 7, sheetExtractedRows.length); k++) {
            if (sheetExtractedRows[k].ledger.parentGroup === item.ledger.parentGroup) {
              sumChildrenCurr += sheetExtractedRows[k].rawCurrentPaise;
              sumChildrenPrior += sheetExtractedRows[k].rawPriorPaise || 0;
              const matchesCurr =
                item.rawCurrentPaise !== 0 &&
                Math.abs(sumChildrenCurr - item.rawCurrentPaise) <= 100 &&
                (k >= i + 2 ||
                  item.isSectionHeaderRow ||
                  item.isScheduleIIIParentHeading ||
                  item.ledger.name.toLowerCase().includes('property, plant') ||
                  item.ledger.name.toLowerCase().includes('trade payable'));
              const matchesPriorWhenCurrZero =
                item.rawCurrentPaise === 0 &&
                (item.rawPriorPaise || 0) !== 0 &&
                Math.abs(sumChildrenPrior - (item.rawPriorPaise || 0)) <= 100 &&
                (k >= i + 2 || item.isScheduleIIIParentHeading);
              if (matchesCurr || matchesPriorWhenCurrZero) {
                isParentOfNextRows = true;
                break;
              }
            } else {
              break;
            }
          }
        }

        if (isParentOfNextRows) {
          reportedSubtotals.push({
            id: `${sourceFileId}_SUB_PARENT_${sheetName}_${i + 1}`,
            statementType: statementContext === 'BALANCE_SHEET' ? 'BALANCE_SHEET' : 'PNL',
            label: item.ledger.name,
            normalizedLabel: `total ${normalizeName(item.ledger.name)}`,
            reportedCurrentPaise: item.rawCurrentPaise,
            reportedPreviousPaise: item.rawPriorPaise,
            reportedBudgetPaise: item.ledger.budgetPaise,
            sourceFileId,
            sourceFilename,
            worksheetName: item.ledger.worksheetName,
            cellReference: item.ledger.cellReference,
            priorCellReference: item.ledger.priorCellReference,
            reportingPeriodLabel: item.ledger.reportingPeriodLabel,
            comparativePeriodLabel: item.ledger.comparativePeriodLabel,
          });
        } else {
          ledgers.push(item.ledger);
        }
      }
    }
    // ------------------------------------------------------------------------
    // CASE 3: OUTSTANDING BILLS EXCEL
    // ------------------------------------------------------------------------
    else if (sheetType === 'OUTSTANDING_BILLS_EXCEL') {
      for (let i = 0; i < dataRows.length; i++) {
        const { rowObj: row } = dataRows[i];
        const party = String(row[colParty || effectiveLedgerCol || ''] || '').trim();
        if (!party || party.toLowerCase().includes('total')) continue;

        const billNo = String(row[colVchNo || ''] || `BILL-${i + 1}`).trim();
        const billDate = parseExcelCellDate(row[colDate || '']) || '2025-04-01';
        const dueDate = parseExcelCellDate(row[colDueDate || '']) || billDate;
        const rawAmt = row[effectiveCurrentCol || colDebit || colCredit || ''];
        const signedPaise = toPaise(rawAmt);
        if (signedPaise === 0) continue;

        const isPayable =
          lowerFilename.includes('payable') ||
          lowerFilename.includes('creditor') ||
          sheetLower.includes('payable') ||
          signedPaise < 0;

        outstandingInvoices.push({
          id: `${sourceFileId}_OB_${i + 1}`,
          partyName: party,
          normalizedPartyName: normalizeName(party),
          category: isPayable ? 'PAYABLE' : 'RECEIVABLE',
          billNumber: billNo,
          billDate,
          dueDate,
          originalAmountPaise: Math.abs(signedPaise),
          adjustedAmountPaise: 0,
          pendingAmountPaise: Math.abs(signedPaise),
          daysOverdue: 0,
          invoiceAgeDays: 0,
          relatedVoucherIds: [],
          sourceFileId,
          sourceFilename,
          isFromBillAllocation: true,
        });
      }
    }
    // ------------------------------------------------------------------------
    // CASE 4: GST RETURNS & SALES / PURCHASE REGISTERS / DAYBOOK
    // ------------------------------------------------------------------------
    else if (
      sheetType === 'GSTR1_RETURN' ||
      sheetType === 'GSTR2B_RETURN' ||
      sheetType === 'GST_REGISTER_EXCEL' ||
      sheetType === 'SALES_REGISTER_EXCEL' ||
      sheetType === 'PURCHASE_REGISTER_EXCEL' ||
      sheetType === 'DAYBOOK_EXCEL'
    ) {
      for (let i = 0; i < dataRows.length; i++) {
        const { rowObj: row } = dataRows[i];
        const partyName = String(row[colParty || effectiveLedgerCol || ''] || '').trim();
        if (partyName.toLowerCase() === 'total' || partyName.toLowerCase() === 'grand total')
          continue;

        const date = parseExcelCellDate(row[colDate || '']);
        if (!date && !partyName) continue;
        const effectiveDate = date || '2025-04-15';
        if (date) {
          if (!minDate || date < minDate) minDate = date;
          if (!maxDate || date > maxDate) maxDate = date;
        }

        const invNo = String(row[colVchNo || ''] || `INV-${i + 1}`).trim();
        const gstin = String(row[colGstin || ''] || '')
          .trim()
          .toUpperCase();
        const taxablePaise = colTaxable ? Math.abs(toPaise(row[colTaxable])) : 0;
        const igstPaise = colIgst ? Math.abs(toPaise(row[colIgst])) : 0;
        const cgstPaise = colCgst ? Math.abs(toPaise(row[colCgst])) : 0;
        const sgstPaise = colSgst ? Math.abs(toPaise(row[colSgst])) : 0;
        const cessPaise = colCess ? Math.abs(toPaise(row[colCess])) : 0;
        const totalTaxPaise = igstPaise + cgstPaise + sgstPaise + cessPaise;

        const drPaise = colDebit ? Math.abs(toPaise(row[colDebit])) : 0;
        const crPaise = colCredit ? Math.abs(toPaise(row[colCredit])) : 0;
        const grossPaise =
          taxablePaise + totalTaxPaise > 0
            ? taxablePaise + totalTaxPaise
            : Math.max(drPaise, crPaise, Math.abs(toPaise(row[effectiveCurrentCol || ''])));

        if (grossPaise === 0) continue;

        if (sheetType === 'GSTR1_RETURN' || sheetType === 'GSTR2B_RETURN') {
          const isCdnr = sheetLower.includes('cdn') || invNo.toUpperCase().startsWith('CN');
          gstInvoices.push({
            id: `${sourceFileId}_GST_${i + 1}`,
            source: sheetType === 'GSTR1_RETURN' ? 'GSTR1' : 'GSTR2B',
            documentType: isCdnr ? 'CREDIT_NOTE' : 'INVOICE',
            partyGstin: gstin,
            partyName: partyName || 'GST Counterparty',
            invoiceNumber: invNo,
            normalizedInvoiceNumber: normalizeInvoiceNo(invNo),
            invoiceDate: effectiveDate,
            returnPeriod: effectiveDate.slice(0, 7),
            supplyType: gstin ? 'B2B' : 'B2C',
            reverseCharge: false,
            itcEligibility: sheetType === 'GSTR2B_RETURN' ? 'ELIGIBLE' : undefined,
            taxableValuePaise: taxablePaise || grossPaise - totalTaxPaise,
            igstPaise,
            cgstPaise,
            sgstPaise,
            cessPaise,
            totalTaxPaise,
            invoiceTotalPaise: grossPaise,
            sourceFileId,
            sourceFilename,
          });
        } else {
          const rawVchType =
            String(row[colVchType || ''] || '') ||
            (sheetType === 'SALES_REGISTER_EXCEL'
              ? 'Sales'
              : sheetType === 'PURCHASE_REGISTER_EXCEL'
              ? 'Purchase'
              : 'Journal');
          const vchType = normalizeVoucherType(rawVchType);
          const isSale = vchType === 'Sales' || sheetType === 'SALES_REGISTER_EXCEL';
          const isPurchase = vchType === 'Purchase' || sheetType === 'PURCHASE_REGISTER_EXCEL';
          const netTaxable = taxablePaise > 0 ? taxablePaise : grossPaise - totalTaxPaise;

          const entries: VoucherEntry[] = [];
          if (isSale) {
            entries.push({
              ledgerName: partyName || 'Sundry Debtors',
              normalizedLedgerName: normalizeName(partyName || 'Sundry Debtors'),
              amountPaise: grossPaise,
              isPartyLedger: true,
              billAllocations: [
                {
                  billName: invNo,
                  billType: 'New Ref',
                  amountPaise: grossPaise,
                },
              ],
            });
            entries.push({
              ledgerName: String(row[effectiveLedgerCol || ''] || 'Sales Account'),
              normalizedLedgerName: normalizeName(
                String(row[effectiveLedgerCol || ''] || 'Sales Account')
              ),
              amountPaise: -netTaxable,
              billAllocations: [],
            });
            if (igstPaise > 0) {
              entries.push({
                ledgerName: 'Output IGST',
                normalizedLedgerName: 'output igst',
                amountPaise: -igstPaise,
                billAllocations: [],
              });
            }
            if (cgstPaise > 0) {
              entries.push({
                ledgerName: 'Output CGST',
                normalizedLedgerName: 'output cgst',
                amountPaise: -cgstPaise,
                billAllocations: [],
              });
            }
            if (sgstPaise > 0) {
              entries.push({
                ledgerName: 'Output SGST',
                normalizedLedgerName: 'output sgst',
                amountPaise: -sgstPaise,
                billAllocations: [],
              });
            }
          } else if (isPurchase) {
            entries.push({
              ledgerName: String(row[effectiveLedgerCol || ''] || 'Purchase Account'),
              normalizedLedgerName: normalizeName(
                String(row[effectiveLedgerCol || ''] || 'Purchase Account')
              ),
              amountPaise: netTaxable,
              billAllocations: [],
            });
            if (igstPaise > 0) {
              entries.push({
                ledgerName: 'Input IGST',
                normalizedLedgerName: 'input igst',
                amountPaise: igstPaise,
                billAllocations: [],
              });
            }
            if (cgstPaise > 0) {
              entries.push({
                ledgerName: 'Input CGST',
                normalizedLedgerName: 'input cgst',
                amountPaise: cgstPaise,
                billAllocations: [],
              });
            }
            if (sgstPaise > 0) {
              entries.push({
                ledgerName: 'Input SGST',
                normalizedLedgerName: 'input sgst',
                amountPaise: sgstPaise,
                billAllocations: [],
              });
            }
            entries.push({
              ledgerName: partyName || 'Sundry Creditors',
              normalizedLedgerName: normalizeName(partyName || 'Sundry Creditors'),
              amountPaise: -grossPaise,
              isPartyLedger: true,
              billAllocations: [
                {
                  billName: invNo,
                  billType: 'New Ref',
                  amountPaise: -grossPaise,
                },
              ],
            });
          } else {
            const targetLedger = String(
              row[effectiveLedgerCol || colParty || ''] || 'General Ledger'
            );
            entries.push({
              ledgerName: targetLedger,
              normalizedLedgerName: normalizeName(targetLedger),
              amountPaise: drPaise > 0 ? drPaise : -crPaise,
              billAllocations: [],
            });
          }

          let totalDr = 0;
          let totalCr = 0;
          for (const e of entries) {
            if (e.amountPaise > 0) totalDr += e.amountPaise;
            else totalCr += Math.abs(e.amountPaise);
          }

          const vchId = `${sourceFileId}_REG_VCH_${i + 1}`;
          vouchers.push({
            id: vchId,
            sourceRecordId: `REG_${vchType}_${invNo}_${effectiveDate}`,
            date: effectiveDate,
            voucherNumber: invNo,
            voucherType: vchType,
            originalVoucherType: rawVchType,
            partyName: partyName || undefined,
            partyGstin: gstin || undefined,
            narration: colNarration ? String(row[colNarration] || '') : undefined,
            entries,
            inventoryMovements: [],
            totalDebitPaise: totalDr,
            totalCreditPaise: totalCr,
            isBalanced: Math.abs(totalDr - totalCr) <= 100,
            sourceFileId,
            sourceFilename,
            isSupplementaryRegister: true,
          });

          if (isSale || isPurchase) {
            gstInvoices.push({
              id: `${vchId}_GST`,
              source: isSale ? 'TALLY_SALES' : 'TALLY_PURCHASE',
              documentType: 'INVOICE',
              partyGstin: gstin,
              partyName: partyName || 'Counterparty',
              invoiceNumber: invNo,
              normalizedInvoiceNumber: normalizeInvoiceNo(invNo),
              invoiceDate: effectiveDate,
              returnPeriod: effectiveDate.slice(0, 7),
              supplyType: gstin ? 'B2B' : 'B2C',
              reverseCharge: false,
              itcEligibility: isPurchase ? 'ELIGIBLE' : undefined,
              taxableValuePaise: netTaxable,
              igstPaise,
              cgstPaise,
              sgstPaise,
              cessPaise,
              totalTaxPaise,
              invoiceTotalPaise: grossPaise,
              sourceFileId,
              sourceFilename,
              linkedVoucherId: vchId,
            });
          }
        }
      }
    } else if (sheetType === 'GSTR3B_RETURN') {
      let outwardTaxable = 0;
      let outwardIgst = 0;
      let outwardCgst = 0;
      let outwardSgst = 0;
      let outwardCess = 0;
      let itcIgst = 0;
      let itcCgst = 0;
      let itcSgst = 0;
      let itcCess = 0;

      for (const { rowObj: row } of dataRows) {
        const desc = String(row[effectiveLedgerCol || rawHeaderRow[0] || ''] || '').toLowerCase();
        const tx = colTaxable ? Math.abs(toPaise(row[colTaxable])) : 0;
        const ig = colIgst ? Math.abs(toPaise(row[colIgst])) : 0;
        const cg = colCgst ? Math.abs(toPaise(row[colCgst])) : 0;
        const sg = colSgst ? Math.abs(toPaise(row[colSgst])) : 0;
        const cs = colCess ? Math.abs(toPaise(row[colCess])) : 0;

        if (desc.includes('outward taxable') || desc.includes('3.1(a)')) {
          outwardTaxable += tx;
          outwardIgst += ig;
          outwardCgst += cg;
          outwardSgst += sg;
          outwardCess += cs;
        } else if (
          desc.includes('all other itc') ||
          desc.includes('net itc') ||
          desc.includes('4(a)(5)') ||
          desc.includes('eligible itc')
        ) {
          itcIgst += ig;
          itcCgst += cg;
          itcSgst += sg;
          itcCess += cs;
        }
      }

      if (outwardTaxable > 0 || itcIgst + itcCgst + itcSgst > 0) {
        gstr3bSummaries.push({
          id: `${sourceFileId}_3B_${sheetName}`,
          returnPeriod: minDate ? minDate.slice(0, 7) : '2025-04',
          outwardTaxableValuePaise: outwardTaxable,
          outwardIgstPaise: outwardIgst,
          outwardCgstPaise: outwardCgst,
          outwardSgstPaise: outwardSgst,
          outwardCessPaise: outwardCess,
          rcmTaxableValuePaise: 0,
          rcmIgstPaise: 0,
          rcmCgstPaise: 0,
          rcmSgstPaise: 0,
          eligibleItcIgstPaise: itcIgst,
          eligibleItcCgstPaise: itcCgst,
          eligibleItcSgstPaise: itcSgst,
          eligibleItcCessPaise: itcCess,
          ineligibleItcPaise: 0,
          reversedItcPaise: 0,
          sourceFileId,
          sourceFilename,
        });
      }
    }
  }

  if (
    !customMapping?.reportCategoryOverride &&
    detectedSheetTypes.size >= 2 &&
    (detectedSheetTypes.has('PROFIT_LOSS_EXCEL') ||
      detectedSheetTypes.has('BALANCE_SHEET_EXCEL') ||
      detectedSheetTypes.has('CASH_FLOW_EXCEL'))
  ) {
    primaryReportType = 'CFO_MONTHLY_PACK_EXCEL';
  }

  const totalExtracted =
    ledgers.length +
    vouchers.length +
    outstandingInvoices.length +
    gstInvoices.length +
    gstr3bSummaries.length +
    cashFlowItems.length;

  const needsMapping = totalExtracted === 0 && capturedHeaders.length > 0;
  if (needsMapping) {
    validationMessages.push({
      severity: 'WARNING',
      code: 'MANUAL_MAPPING_REQUIRED',
      message:
        'Spreadsheet headers could not be automatically mapped to a financial statement or Tally report. Click "Map" on this file to configure statement and column mapping.',
    });
  }

  return {
    detectedReportType: primaryReportType,
    extractedCompanyName: extractedCompanyName || undefined,
    detectedCurrency,
    periodStart: minDate || undefined,
    periodEnd: maxDate || undefined,
    currentPeriodLabel: detectedCurrentPeriodLabel || undefined,
    comparativePeriodLabel: detectedComparativePeriodLabel || undefined,
    hasMonthlyGranularity: hasAnyMonthlyGranularity,
    ledgers,
    vouchers,
    outstandingInvoices,
    gstInvoices,
    gstr3bSummaries,
    cashFlowItems,
    reportedSubtotals,
    validationMessages,
    rawHeaders: capturedHeaders,
    rawSampleRows: capturedSampleRows,
    needsMapping,
  };
}
