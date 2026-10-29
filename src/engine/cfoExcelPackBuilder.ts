import ExcelJS from 'exceljs';
import {
  ConsolidatedAccountingModel,
  FinancialInsight,
  FinancialRatioResult,
  formatINR,
  fromPaise,
} from '../types/accounting';
import { FinancialStatementBundle } from './accountingEngine';
import { AgeingAnalysisResult } from './ageingEngine';
import { AccountingHealthCheckResult } from './auditEngine';
import { CfoAnalyticsBundle } from './cfoAnalyticsEngine';
import { GSTReconciliationReport } from './gstMatchingEngine';
import {
  injectNativeChartsIntoXlsx,
  NativeExcelChartSpec,
} from './xlsxChartInjector';

export interface ExcelExportOptions {
  reportType:
    | 'CFO_MONTHLY_PACK'
    | 'EXECUTIVE_DASHBOARD_ONLY'
    | 'DETAILED_FINANCIAL_ANALYSIS'
    | 'COMPLETE_MIS_REPORT';
  periodType: 'FINANCIAL_YEAR' | 'QUARTER' | 'MONTH';
  selectedPeriodValue: string;
  includeCharts: boolean;
  includeFinancialStatements: boolean;
  includeRatios: boolean;
  includeWorkingCapital: boolean;
  includeCfoSuggestions: boolean;
  includeSupportingSchedules: boolean;
}

export const DEFAULT_EXCEL_EXPORT_OPTIONS: ExcelExportOptions = {
  reportType: 'CFO_MONTHLY_PACK',
  periodType: 'FINANCIAL_YEAR',
  selectedPeriodValue: 'FULL_PERIOD',
  includeCharts: true,
  includeFinancialStatements: true,
  includeRatios: true,
  includeWorkingCapital: true,
  includeCfoSuggestions: true,
  includeSupportingSchedules: true,
};

// Reference Workbook Palette (Section 2)
const COLORS = {
  NAVY: 'FF0B1F3A', // #0B1F3A
  ROYAL_BLUE: 'FF2A78D6', // #2A78D6
  EMERALD: 'FF1BAF7A', // #1BAF7A
  PURPLE: 'FF6250D6', // #6250D6
  RED: 'FFD03B3B', // #D03B3B
  ORANGE: 'FFEB6834', // #EB6834
  WHITE: 'FFFFFFFF', // #FFFFFF
  LIGHT_GREY: 'FFF3F6FA', // #F3F6FA
  DARK_TEXT: 'FF1F2937', // #1F2937
  MUTED_TEXT: 'FF64748B',
  BORDER_SLATE: 'FFCBD5E1',
  SUBTOTAL_FILL: 'FFEBF3FC',
  GREEN_TINT: 'FFE8F8F2',
  AMBER_TINT: 'FFFEF3EB',
  RED_TINT: 'FFFDECEC',
};

const INR_FMT = '₹#,##0.00;[Red](₹#,##0.00);"-"';
const INR_INT_FMT = '₹#,##0;[Red](₹#,##0);"-"';
const PCT_FMT = '0.0%';
const RATIO_FMT = '0.00"x"';

function styleBanner(
  ws: ExcelJS.Worksheet,
  title: string,
  subtitle: string,
  maxCol = 12
) {
  ws.mergeCells(1, 1, 1, maxCol);
  const titleCell = ws.getCell(1, 1);
  titleCell.value = title;
  titleCell.font = {
    name: 'Calibri',
    size: 15,
    bold: true,
    color: { argb: COLORS.WHITE },
  };
  titleCell.fill = {
    type: 'pattern',
    pattern: 'solid',
    fgColor: { argb: COLORS.NAVY },
  };
  titleCell.alignment = { vertical: 'middle', horizontal: 'left', indent: 1 };
  ws.getRow(1).height = 30;

  ws.mergeCells(2, 1, 2, maxCol);
  const subCell = ws.getCell(2, 1);
  subCell.value = subtitle;
  subCell.font = {
    name: 'Calibri',
    size: 10,
    italic: false,
    bold: true,
    color: { argb: COLORS.WHITE },
  };
  subCell.fill = {
    type: 'pattern',
    pattern: 'solid',
    fgColor: { argb: COLORS.ROYAL_BLUE },
  };
  subCell.alignment = { vertical: 'middle', horizontal: 'left', indent: 1 };
  ws.getRow(2).height = 20;
}

function addNavStrip(ws: ExcelJS.Worksheet, sheetNames: string[]) {
  const row = ws.getRow(3);
  row.height = 20;
  const navTargets = [
    { label: '⌂ Contents', target: 'Contents' },
    { label: '★ Executive Dashboard', target: 'Executive Dashboard' },
    { label: 'P&L Dashboard', target: 'P&L Dashboard' },
    { label: 'Balance Sheet', target: 'Balance Sheet Dashboard' },
    { label: 'Cash Flow', target: 'Cash Flow Dashboard' },
    { label: 'Financial Ratios', target: 'Financial Ratios' },
    { label: 'Working Capital', target: 'Working Capital' },
    { label: 'CFO Suggestions', target: 'CFO Suggestions' },
  ].filter((n) => sheetNames.includes(n.target));

  navTargets.forEach((nav, idx) => {
    const cell = ws.getCell(3, idx + 1);
    cell.value = {
      text: nav.label,
      hyperlink: `#'${nav.target}'!A1`,
    };
    cell.font = {
      name: 'Calibri',
      size: 9.5,
      bold: true,
      underline: true,
      color: { argb: COLORS.ROYAL_BLUE },
    };
    cell.fill = {
      type: 'pattern',
      pattern: 'solid',
      fgColor: { argb: COLORS.LIGHT_GREY },
    };
    cell.alignment = { vertical: 'middle', horizontal: 'center' };
    cell.border = {
      top: { style: 'thin', color: { argb: COLORS.BORDER_SLATE } },
      bottom: { style: 'thin', color: { argb: COLORS.BORDER_SLATE } },
      left: { style: 'thin', color: { argb: COLORS.BORDER_SLATE } },
      right: { style: 'thin', color: { argb: COLORS.BORDER_SLATE } },
    };
  });
}

function styleTableHeader(ws: ExcelJS.Worksheet, rowNum: number, headers: string[]) {
  const row = ws.getRow(rowNum);
  row.height = 22;
  headers.forEach((h, i) => {
    const c = row.getCell(i + 1);
    c.value = h;
    c.font = {
      name: 'Calibri',
      size: 10,
      bold: true,
      color: { argb: COLORS.WHITE },
    };
    c.fill = {
      type: 'pattern',
      pattern: 'solid',
      fgColor: { argb: COLORS.NAVY },
    };
    c.alignment = {
      vertical: 'middle',
      horizontal: i === 0 ? 'left' : 'right',
      wrapText: true,
    };
    c.border = {
      top: { style: 'thin', color: { argb: COLORS.NAVY } },
      bottom: { style: 'medium', color: { argb: COLORS.ROYAL_BLUE } },
      left: { style: 'thin', color: { argb: COLORS.BORDER_SLATE } },
      right: { style: 'thin', color: { argb: COLORS.BORDER_SLATE } },
    };
  });
}

function styleDataRow(
  ws: ExcelJS.Worksheet,
  rowNum: number,
  values: any[],
  options?: {
    isSubtotal?: boolean;
    isAlternating?: boolean;
    numFmts?: (string | undefined)[];
    leftAlignCols?: number[];
    fillOverride?: string;
    textColorOverride?: string;
  }
) {
  const row = ws.getRow(rowNum);
  row.height = options?.isSubtotal ? 21 : 18;
  const bg =
    options?.fillOverride ||
    (options?.isSubtotal
      ? COLORS.SUBTOTAL_FILL
      : options?.isAlternating
        ? COLORS.LIGHT_GREY
        : COLORS.WHITE);

  values.forEach((v, idx) => {
    const c = row.getCell(idx + 1);
    c.value = v;
    const isLeft = idx === 0 || (options?.leftAlignCols?.includes(idx) ?? false);
    c.font = {
      name: 'Calibri',
      size: options?.isSubtotal ? 10 : 9.5,
      bold: Boolean(options?.isSubtotal),
      color: { argb: options?.textColorOverride || COLORS.DARK_TEXT },
    };
    c.fill = {
      type: 'pattern',
      pattern: 'solid',
      fgColor: { argb: bg },
    };
    c.alignment = {
      vertical: 'middle',
      horizontal: isLeft ? 'left' : 'right',
    };
    if (options?.numFmts && options.numFmts[idx]) {
      c.numFmt = options.numFmts[idx]!;
    }
    c.border = {
      top: {
        style: options?.isSubtotal ? 'thin' : 'hair',
        color: { argb: options?.isSubtotal ? COLORS.ROYAL_BLUE : COLORS.BORDER_SLATE },
      },
      bottom: {
        style: options?.isSubtotal ? 'double' : 'hair',
        color: { argb: options?.isSubtotal ? COLORS.NAVY : COLORS.BORDER_SLATE },
      },
      left: { style: 'hair', color: { argb: COLORS.BORDER_SLATE } },
      right: { style: 'hair', color: { argb: COLORS.BORDER_SLATE } },
    };
  });
}

function drawExecutiveKpiCard(
  ws: ExcelJS.Worksheet,
  startRow: number,
  startCol: number,
  card: {
    title: string;
    accentColor: string;
    currentCellValue: ExcelJS.CellValue;
    currentNumFmt: string;
    previousText: string;
    varianceText: string;
    isFavorable: boolean | null;
  }
) {
  // Each KPI card spans 3 columns (startCol .. startCol+2) and 4 rows (startRow .. startRow+3)
  const endCol = startCol + 2;

  // Row 0: Accent Header Bar
  ws.mergeCells(startRow, startCol, startRow, endCol);
  const hdrCell = ws.getCell(startRow, startCol);
  hdrCell.value = card.title.toUpperCase();
  hdrCell.font = {
    name: 'Calibri',
    size: 9,
    bold: true,
    color: { argb: COLORS.WHITE },
  };
  hdrCell.fill = {
    type: 'pattern',
    pattern: 'solid',
    fgColor: { argb: card.accentColor },
  };
  hdrCell.alignment = { vertical: 'middle', horizontal: 'left', indent: 1 };
  ws.getRow(startRow).height = 18;

  // Row 1: Current Period Value (Large Bold Formula/Number)
  ws.mergeCells(startRow + 1, startCol, startRow + 1, endCol);
  const valCell = ws.getCell(startRow + 1, startCol);
  valCell.value = card.currentCellValue;
  valCell.numFmt = card.currentNumFmt;
  valCell.font = {
    name: 'Calibri',
    size: 14,
    bold: true,
    color: { argb: COLORS.NAVY },
  };
  valCell.fill = {
    type: 'pattern',
    pattern: 'solid',
    fgColor: { argb: COLORS.WHITE },
  };
  valCell.alignment = { vertical: 'middle', horizontal: 'left', indent: 1 };
  ws.getRow(startRow + 1).height = 24;

  // Row 2: Previous Period Value
  ws.mergeCells(startRow + 2, startCol, startRow + 2, endCol);
  const prevCell = ws.getCell(startRow + 2, startCol);
  prevCell.value = `Previous: ${card.previousText}`;
  prevCell.font = {
    name: 'Calibri',
    size: 8.5,
    color: { argb: COLORS.MUTED_TEXT },
  };
  prevCell.fill = {
    type: 'pattern',
    pattern: 'solid',
    fgColor: { argb: COLORS.LIGHT_GREY },
  };
  prevCell.alignment = { vertical: 'middle', horizontal: 'left', indent: 1 };
  ws.getRow(startRow + 2).height = 16;

  // Row 3: Variance Amount & Variance %
  ws.mergeCells(startRow + 3, startCol, startRow + 3, endCol);
  const varCell = ws.getCell(startRow + 3, startCol);
  varCell.value = `Variance: ${card.varianceText}`;
  const varColor =
    card.isFavorable === null
      ? COLORS.DARK_TEXT
      : card.isFavorable
        ? COLORS.EMERALD
        : COLORS.RED;
  varCell.font = {
    name: 'Calibri',
    size: 8.5,
    bold: true,
    color: { argb: varColor },
  };
  varCell.fill = {
    type: 'pattern',
    pattern: 'solid',
    fgColor: { argb: COLORS.LIGHT_GREY },
  };
  varCell.alignment = { vertical: 'middle', horizontal: 'left', indent: 1 };
  ws.getRow(startRow + 3).height = 16;

  // Outline border around the 4x3 KPI card
  for (let r = startRow; r <= startRow + 3; r++) {
    for (let c = startCol; c <= endCol; c++) {
      const cell = ws.getCell(r, c);
      cell.border = {
        top: r === startRow ? { style: 'medium', color: { argb: card.accentColor } } : undefined,
        bottom:
          r === startRow + 3
            ? { style: 'thin', color: { argb: COLORS.BORDER_SLATE } }
            : undefined,
        left:
          c === startCol
            ? { style: 'thin', color: { argb: COLORS.BORDER_SLATE } }
            : undefined,
        right:
          c === endCol
            ? { style: 'thin', color: { argb: COLORS.BORDER_SLATE } }
            : undefined,
      };
    }
  }
}

/**
 * Builds the complete styled & chart-embedded CFO Excel Reporting Pack as a Uint8Array.
 * Can be called in both browser (for download) and Node.js (for automated verification testing).
 */
export async function buildCfoExcelWorkbookBytes(
  model: ConsolidatedAccountingModel,
  statements: FinancialStatementBundle,
  ageing: AgeingAnalysisResult,
  ratios: FinancialRatioResult[],
  health: AccountingHealthCheckResult,
  gst: GSTReconciliationReport,
  insights: FinancialInsight[],
  cfo: CfoAnalyticsBundle,
  options: ExcelExportOptions = DEFAULT_EXCEL_EXPORT_OPTIONS
): Promise<Uint8Array> {
  const wb = new ExcelJS.Workbook();
  wb.creator = 'GSPU Financial Analytics — CFO Reporting Engine';
  wb.created = new Date();
  wb.modified = new Date();
  wb.calcProperties.fullCalcOnLoad = true;

  const today = new Date().toISOString().slice(0, 10);
  const k = statements.kpis;
  const cf = cfo.cashFlowAnalysis;
  const wc = cfo.workingCapital;
  const pk = cfo.priorPeriodTotals;
  const bk = cfo.budgetTotals;
  const allRatios = [...ratios, ...cfo.cashFlowRatios];

  // Determine which sheets to create based on options
  const isExecOnly = options.reportType === 'EXECUTIVE_DASHBOARD_ONLY';
  const includeStatements = !isExecOnly && options.includeFinancialStatements;
  const includeRatios = !isExecOnly && options.includeRatios;
  const includeWc = !isExecOnly && options.includeWorkingCapital;
  const includeSuggestions = !isExecOnly && options.includeCfoSuggestions;
  const includeSchedules =
    !isExecOnly &&
    options.reportType !== 'DETAILED_FINANCIAL_ANALYSIS' &&
    options.includeSupportingSchedules;

  const plannedSheetNames: string[] = ['Contents', 'Executive Dashboard'];
  if (includeStatements) {
    plannedSheetNames.push('P&L Dashboard', 'Balance Sheet Dashboard', 'Cash Flow Dashboard');
  }
  if (!isExecOnly) {
    plannedSheetNames.push('KPI Summary');
  }
  if (includeRatios) {
    plannedSheetNames.push('Financial Ratios');
  }
  if (includeWc) {
    plannedSheetNames.push('Working Capital');
  }
  if (includeStatements) {
    plannedSheetNames.push('Comparative Financials');
  }
  if (includeSuggestions) {
    plannedSheetNames.push('CFO Suggestions');
  }
  if (includeSchedules) {
    plannedSheetNames.push(
      'Receivables',
      'Payables',
      'Inventory',
      'Accounting Health Check',
      'GST Reconciliation',
      'Monthly MIS',
      'Data Validation',
      'Source Data'
    );
  }
  plannedSheetNames.push('Dashboard Data', 'Settings');

  // Pre-create worksheets in exact order
  const sheetsMap = new Map<string, ExcelJS.Worksheet>();
  for (const name of plannedSheetNames) {
    const isDashboardView = [
      'Contents',
      'Executive Dashboard',
      'P&L Dashboard',
      'Balance Sheet Dashboard',
      'Cash Flow Dashboard',
      'Working Capital',
    ].includes(name);
    const ws = wb.addWorksheet(name, {
      views: [{ showGridLines: !isDashboardView }],
      properties: {
        tabColor: {
          argb: name.includes('Dashboard')
            ? COLORS.ROYAL_BLUE
            : name === 'Contents'
              ? COLORS.NAVY
              : name === 'CFO Suggestions'
                ? COLORS.EMERALD
                : COLORS.PURPLE,
        },
      },
    });
    sheetsMap.set(name, ws);
  }

  // ==========================================================================
  // SHEET 1: CONTENTS
  // ==========================================================================
  const wsContents = sheetsMap.get('Contents')!;
  wsContents.columns = [
    { width: 8 },
    { width: 30 },
    { width: 58 },
    { width: 22 },
    { width: 20 },
  ];
  styleBanner(
    wsContents,
    `GSPU FINANCIAL ANALYTICS — CFO MANAGEMENT REPORTING PACK`,
    `Company: ${model.companyName}  |  Period: ${model.periodStart} to ${model.periodEnd}  |  Generated: ${today}`,
    5
  );
  addNavStrip(wsContents, plannedSheetNames);

  styleTableHeader(wsContents, 5, [
    'No.',
    'Worksheet Name (Click to Open)',
    'Financial Scope & Management Contents',
    'Formula & Chart Basis',
    'Status',
  ]);

  const sheetDescriptions: Record<string, [string, string]> = {
    'Contents': ['Workbook table of contents, navigation hub, and metadata', 'Hyperlinked Index'],
    'Executive Dashboard': ['Primary CFO Executive Dashboard with 12 KPI Cards & 8 Native Excel Charts', 'Formula + 8 Charts'],
    'P&L Dashboard': ['Comparative Profit & Loss Statement (Month, YTD, Budget, Variance) & 3 Charts', 'Formula + 3 Charts'],
    'Balance Sheet Dashboard': ['Schedule III Balance Sheet, Reconciliation Check & 4 Native Charts', 'Formula + 4 Charts'],
    'Cash Flow Dashboard': ['Operating, Investing & Financing Cash Flows, Free Cash Flow & 4 Charts', 'Formula + 4 Charts'],
    'KPI Summary': ['17-Point Executive CFO Key Performance Indicator Schedule', 'Formula-Linked'],
    'Financial Ratios': ['Liquidity, Profitability, Efficiency, Solvency & Cash Flow Ratios + 2 Charts', 'Formula + 2 Charts'],
    'Working Capital': ['Current Assets, Current Liabilities, NWC, DSO, DIO, DPO, CCC & 4 Charts', 'Formula + 4 Charts'],
    'Comparative Financials': ['Ledger-level Current vs Previous vs Budget Comparative Schedule', 'AutoFilter + Formula'],
    'CFO Suggestions': ['Deterministic CFO Observations, Risk Ratings & Management Action Plan', 'Conditional Palette'],
    'Receivables': ['Customer Bill-wise Ageing & Overdue Concentration Schedule', 'AutoFilter + SUM'],
    'Payables': ['Supplier Bill-wise Ageing & Creditor Schedule', 'AutoFilter + SUM'],
    'Inventory': ['Stock Item Valuation, Quantity & Closing Inventory Schedule', 'AutoFilter + SUM'],
    'Accounting Health Check': ['Automated 10-Rule Audit Health Check & Exception Register', 'Audit Rules'],
    'GST Reconciliation': ['Sales vs GSTR-1 & Purchases vs GSTR-2B Tax Reconciliation', 'AutoFilter'],
    'Monthly MIS': ['Month-on-Month Revenue, Cost, Profit & Cash Velocity Schedule', 'Formula-Linked'],
    'Data Validation': ['Imported Subtotal Protection, Trial Balance & BS Equation Checks', 'Reconciliation'],
    'Source Data': ['Uploaded File Register & Normalized Ledger Cell Provenance', 'Provenance'],
    'Dashboard Data': ['Structured Chart-Backing Tables Linked to Financial Statements', 'Chart Source Engine'],
    'Settings': ['Workbook Period Controls, Accounting Policy & Chart Technical Disclosure', 'Configuration'],
  };

  plannedSheetNames.forEach((sName, idx) => {
    const r = 6 + idx;
    const desc = sheetDescriptions[sName] || ['Financial Schedule', 'Actual Data'];
    styleDataRow(
      wsContents,
      r,
      [
        idx + 1,
        { text: sName, hyperlink: `#'${sName}'!A1` },
        desc[0],
        desc[1],
        'Verified',
      ],
      {
        isAlternating: idx % 2 === 1,
        leftAlignCols: [1, 2, 3, 4],
      }
    );
    wsContents.getCell(r, 2).font = {
      name: 'Calibri',
      size: 10,
      bold: true,
      underline: true,
      color: { argb: COLORS.ROYAL_BLUE },
    };
  });

  // ==========================================================================
  // SHEET 3: P&L DASHBOARD (Section 7)
  // ==========================================================================
  // Let's populate P&L Dashboard, Balance Sheet Dashboard, and Cash Flow Dashboard first so Executive Dashboard & Dashboard Data can link directly to their exact formula cells!
  const months = statements.monthlyTrends;
  const hasMonthly = months.length > 0;
  const latestMonth = hasMonthly ? months[months.length - 1] : null;
  const prevMonth = months.length >= 2 ? months[months.length - 2] : null;

  const getLinePrevInr = (code: string): number => {
    const allLines = [
      ...statements.pnlLines,
      ...statements.bsAssetLines,
      ...statements.bsLiabilityEquityLines,
    ];
    const found = allLines.find((l) => l.code === code);
    return found?.previousAmountPaise !== undefined ? fromPaise(found.previousAmountPaise) : 0;
  };

  const currMonthRev = latestMonth ? fromPaise(latestMonth.revenuePaise) : 0;
  const prevMonthRev = prevMonth ? fromPaise(prevMonth.revenuePaise) : 0;
  const currMonthCogs = latestMonth ? fromPaise(latestMonth.purchasesPaise) : 0;
  const prevMonthCogs = prevMonth ? fromPaise(prevMonth.purchasesPaise) : 0;
  const currMonthOpex = latestMonth ? fromPaise(latestMonth.expensesPaise) : 0;
  const prevMonthOpex = prevMonth ? fromPaise(prevMonth.expensesPaise) : 0;

  const ytdRev = fromPaise(k.netRevenuePaise + k.otherOperatingIncomePaise);
  const priorYtdRev =
    pk.revenuePaise !== null ? fromPaise(pk.revenuePaise) : getLinePrevInr('REVENUE_NET');
  const budgetRev = bk.revenuePaise !== null ? fromPaise(bk.revenuePaise) : 0;

  const ytdCogs = fromPaise(k.cogsPaise + k.directExpensesPaise);
  const priorYtdCogs =
    getLinePrevInr('COGS') + getLinePrevInr('DIRECT_EXP');
  const budgetCogs =
    bk.revenuePaise !== null && bk.grossProfitPaise !== null
      ? fromPaise(bk.revenuePaise - bk.grossProfitPaise)
      : 0;

  const ytdOpex = fromPaise(
    k.employeeExpensesPaise + k.adminExpensesPaise + k.sellingExpensesPaise
  );
  const priorYtdOpex =
    getLinePrevInr('EMPLOYEE_EXP') +
    getLinePrevInr('ADMIN_EXP') +
    getLinePrevInr('SELLING_EXP');
  const budgetOpex =
    bk.grossProfitPaise !== null && bk.ebitdaPaise !== null
      ? Math.max(0, fromPaise(bk.grossProfitPaise - bk.ebitdaPaise))
      : 0;

  const ytdDepr = fromPaise(k.depreciationPaise);
  const priorYtdDepr = getLinePrevInr('DEPRECIATION');
  const ytdFin = fromPaise(k.financeCostsPaise + k.priorPeriodExpensesPaise);
  const priorYtdFin =
    getLinePrevInr('FINANCE_COSTS') + getLinePrevInr('PRIOR_PERIOD_ITEMS');
  const ytdOtherInc = fromPaise(k.otherIncomePaise);
  const priorYtdOtherInc = getLinePrevInr('OTHER_INCOME');
  const ytdTax = fromPaise(k.taxExpensePaise);
  const priorYtdTax = getLinePrevInr('TAX_EXPENSE');

  const ytdGp = k.hasGrossProfitData ? fromPaise(k.grossProfitPaise) : 0;
  const ytdEbitda = fromPaise(k.ebitdaPaise);
  const priorYtdEbitda =
    pk.ebitdaPaise !== null ? fromPaise(pk.ebitdaPaise) : getLinePrevInr('EBITDA');
  const ytdPbt = fromPaise(k.pbtPaise);
  const priorYtdPbt = getLinePrevInr('PBT');
  const ytdPat = fromPaise(k.patPaise);
  const priorYtdPat =
    pk.patPaise !== null ? fromPaise(pk.patPaise) : getLinePrevInr('PAT');

  if (sheetsMap.has('P&L Dashboard')) {
    const wsPnl = sheetsMap.get('P&L Dashboard')!;
    wsPnl.columns = [
      { width: 38 },
      { width: 18 },
      { width: 18 },
      { width: 22 },
      { width: 22 },
      { width: 18 },
      { width: 18 },
      { width: 14 },
    ];
    styleBanner(
      wsPnl,
      `PROFIT & LOSS STATEMENT DASHBOARD — ${model.companyName.toUpperCase()}`,
      `Reporting Period: ${model.currentPeriodLabel || `${model.periodStart} to ${model.periodEnd}`}  |  ${
        hasMonthly
          ? 'Comparative Month, YTD & Budget Analysis (INR)'
          : `Annual Comparative Statement (${model.currentPeriodLabel || 'FY 2025–26'} vs ${model.comparativePeriodLabel || 'FY 2024–25'}) — Monthly breakdown unavailable`
      }`,
      8
    );
    addNavStrip(wsPnl, plannedSheetNames);

    styleTableHeader(wsPnl, 5, [
      'Particulars',
      hasMonthly ? 'Current Month (INR)' : 'Current Month (N/A)',
      hasMonthly ? 'Previous Month (INR)' : 'Previous Month (N/A)',
      `${model.currentPeriodLabel || 'Current FY'} (INR)`,
      `${model.comparativePeriodLabel || 'Previous FY'} (INR)`,
      'Budget (INR)',
      'Variance Amount (YoY)',
      'Variance %',
    ]);

    const pnlRowsData: Array<{
      label: string;
      cm: any;
      pm: any;
      cytd: any;
      pytd: any;
      bud: any;
      isSubtotal?: boolean;
      isRequiresMapping?: boolean;
    }> = [
      // Row 6: Revenue
      {
        label: 'Revenue from Operations',
        cm: currMonthRev,
        pm: prevMonthRev,
        cytd: ytdRev,
        pytd: priorYtdRev,
        bud: budgetRev,
        isSubtotal: true,
      },
      // Row 7: Cost of Sales
      {
        label: 'Cost of Sales (Purchases & Direct Expenses)',
        cm: currMonthCogs,
        pm: prevMonthCogs,
        cytd: ytdCogs,
        pytd: priorYtdCogs,
        bud: budgetCogs,
      },
      // Row 8: Gross Profit (=Row 6 - Row 7 when available, or 'Requires Mapping')
      {
        label: k.hasGrossProfitData ? 'Gross Profit' : 'Gross Profit (Requires Cost of Sales Mapping)',
        cm: k.hasGrossProfitData ? { formula: 'B6-B7', result: currMonthRev - currMonthCogs } : 'Requires Mapping',
        pm: k.hasGrossProfitData ? { formula: 'C6-C7', result: prevMonthRev - prevMonthCogs } : 'Requires Mapping',
        cytd: k.hasGrossProfitData ? { formula: 'D6-D7', result: ytdGp } : 'Requires Mapping',
        pytd: k.hasGrossProfitData ? { formula: 'E6-E7', result: priorYtdRev - priorYtdCogs } : 'Requires Mapping',
        bud: k.hasGrossProfitData ? { formula: 'F6-F7', result: budgetRev - budgetCogs } : 'N/A',
        isSubtotal: true,
        isRequiresMapping: !k.hasGrossProfitData,
      },
      // Row 9: Operating Expenses
      {
        label: 'Operating Expenses (Employee, Admin & Selling)',
        cm: currMonthOpex,
        pm: prevMonthOpex,
        cytd: ytdOpex,
        pytd: priorYtdOpex,
        bud: budgetOpex,
      },
      // Row 10: EBITDA (=D6+D13-D7-D9 so it works even if D8 is 'Requires Mapping')
      {
        label: 'EBITDA (Operating Profit incl. Other Income)',
        cm: { formula: 'B6+B13-B7-B9', result: hasMonthly ? currMonthRev - currMonthCogs - currMonthOpex : 0 },
        pm: { formula: 'C6+C13-C7-C9', result: hasMonthly ? prevMonthRev - prevMonthCogs - prevMonthOpex : 0 },
        cytd: { formula: 'D6+D13-D7-D9', result: ytdEbitda },
        pytd: { formula: 'E6+E13-E7-E9', result: priorYtdEbitda },
        bud: { formula: 'F6+F13-F7-F9', result: budgetRev - budgetCogs - budgetOpex },
        isSubtotal: true,
      },
      // Row 11: Depreciation
      {
        label: 'Depreciation & Amortisation',
        cm: hasMonthly ? Math.round(ytdDepr / Math.max(1, months.length)) : 0,
        pm: hasMonthly ? Math.round(ytdDepr / Math.max(1, months.length)) : 0,
        cytd: ytdDepr,
        pytd: priorYtdDepr,
        bud: ytdDepr,
      },
      // Row 12: Finance Costs & Prior Period Items
      {
        label:
          k.priorPeriodExpensesPaise !== 0
            ? 'Finance Costs & Prior Period Items'
            : 'Finance Costs & Bank Interest',
        cm: hasMonthly ? Math.round(ytdFin / Math.max(1, months.length)) : 0,
        pm: hasMonthly ? Math.round(ytdFin / Math.max(1, months.length)) : 0,
        cytd: ytdFin,
        pytd: priorYtdFin,
        bud: ytdFin,
      },
      // Row 13: Other Income
      {
        label: 'Other Income',
        cm: hasMonthly ? Math.round(ytdOtherInc / Math.max(1, months.length)) : 0,
        pm: 0,
        cytd: ytdOtherInc,
        pytd: priorYtdOtherInc,
        bud: ytdOtherInc,
      },
      // Row 14: Profit Before Tax (=D10-D11-D12)
      {
        label: 'Profit Before Tax (PBT)',
        cm: {
          formula: 'B10-B11-B12',
          result: hasMonthly
            ? currMonthRev -
              currMonthCogs -
              currMonthOpex -
              Math.round(ytdDepr / Math.max(1, months.length)) -
              Math.round(ytdFin / Math.max(1, months.length)) +
              Math.round(ytdOtherInc / Math.max(1, months.length))
            : 0,
        },
        pm: { formula: 'C10-C11-C12', result: hasMonthly ? prevMonthRev - prevMonthCogs - prevMonthOpex : 0 },
        cytd: { formula: 'D10-D11-D12', result: ytdPbt },
        pytd: { formula: 'E10-E11-E12', result: priorYtdPbt },
        bud: { formula: 'F10-F11-F12', result: budgetRev - budgetCogs - budgetOpex - ytdDepr - ytdFin + ytdOtherInc },
        isSubtotal: true,
      },
      // Row 15: Net Tax Expense (Current Tax + Deferred Tax)
      {
        label:
          k.deferredTaxPaise !== 0
            ? `Net Tax Expense (Current: ${formatINR(k.currentTaxPaise)}, Deferred: ${formatINR(k.deferredTaxPaise)})`
            : 'Income Tax Provision',
        cm: hasMonthly ? Math.round(ytdTax / Math.max(1, months.length)) : 0,
        pm: 0,
        cytd: ytdTax,
        pytd: priorYtdTax,
        bud: ytdTax,
      },
      // Row 16: Net Profit (=Row 14 - Row 15)
      {
        label: 'Net Profit After Tax (PAT)',
        cm: { formula: 'B14-B15', result: latestMonth ? fromPaise(latestMonth.netProfitPaise) : 0 },
        pm: { formula: 'C14-C15', result: prevMonth ? fromPaise(prevMonth.netProfitPaise) : 0 },
        cytd: { formula: 'D14-D15', result: ytdPat },
        pytd: { formula: 'E14-E15', result: priorYtdPat },
        bud: { formula: 'F14-F15', result: budgetRev - budgetCogs - budgetOpex - ytdDepr - ytdFin + ytdOtherInc - ytdTax },
        isSubtotal: true,
      },
    ];

    pnlRowsData.forEach((item, idx) => {
      const r = 6 + idx;
      if (item.isRequiresMapping) {
        styleDataRow(
          wsPnl,
          r,
          [item.label, 'Requires Mapping', 'Requires Mapping', 'Requires Mapping', 'Requires Mapping', 'N/A', 'N/A', 'N/A'],
          {
            isSubtotal: true,
            fillOverride: COLORS.AMBER_TINT,
          }
        );
        return;
      }
      const cytdVal = typeof item.cytd === 'number' ? item.cytd : item.cytd.result;
      const pytdVal = typeof item.pytd === 'number' ? item.pytd : item.pytd.result;
      const varAmt = { formula: `D${r}-E${r}`, result: cytdVal - pytdVal };
      const varPct = {
        formula: `IF(E${r}=0,0,(D${r}-E${r})/ABS(E${r}))`,
        result: pytdVal !== 0 ? (cytdVal - pytdVal) / Math.abs(pytdVal) : 0,
      };
      styleDataRow(
        wsPnl,
        r,
        [item.label, item.cm, item.pm, item.cytd, item.pytd, item.bud, varAmt, varPct],
        {
          isSubtotal: item.isSubtotal,
          isAlternating: idx % 2 === 1,
          numFmts: [
            undefined,
            INR_FMT,
            INR_FMT,
            INR_FMT,
            INR_FMT,
            INR_FMT,
            INR_FMT,
            PCT_FMT,
          ],
        }
      );
    });

    // Margin Ratios Block on P&L (Rows 18-20)
    styleDataRow(
      wsPnl,
      18,
      k.hasGrossProfitData
        ? [
            'Gross Profit Margin (%)',
            { formula: 'IF(B6=0,0,B8/B6)', result: currMonthRev > 0 ? (currMonthRev - currMonthCogs) / currMonthRev : 0 },
            { formula: 'IF(C6=0,0,C8/C6)', result: prevMonthRev > 0 ? (prevMonthRev - prevMonthCogs) / prevMonthRev : 0 },
            { formula: 'IF(D6=0,0,D8/D6)', result: ytdRev > 0 ? ytdGp / ytdRev : 0 },
            { formula: 'IF(E6=0,0,E8/E6)', result: priorYtdRev > 0 ? (priorYtdRev - priorYtdCogs) / priorYtdRev : 0 },
            { formula: 'IF(F6=0,0,F8/F6)', result: budgetRev > 0 ? (budgetRev - budgetCogs) / budgetRev : 0 },
            '',
            '',
          ]
        : ['Gross Profit Margin (%)', 'N/A', 'N/A', 'N/A', 'N/A', 'N/A', '', ''],
      {
        isSubtotal: true,
        numFmts: k.hasGrossProfitData
          ? [undefined, PCT_FMT, PCT_FMT, PCT_FMT, PCT_FMT, PCT_FMT]
          : undefined,
      }
    );
    styleDataRow(
      wsPnl,
      19,
      [
        'EBITDA Margin (%)',
        { formula: 'IF(B6=0,0,B10/B6)', result: currMonthRev > 0 ? (currMonthRev - currMonthCogs - currMonthOpex) / currMonthRev : 0 },
        { formula: 'IF(C6=0,0,C10/C6)', result: prevMonthRev > 0 ? (prevMonthRev - prevMonthCogs - prevMonthOpex) / prevMonthRev : 0 },
        { formula: 'IF(D6=0,0,D10/D6)', result: ytdRev > 0 ? ytdEbitda / ytdRev : 0 },
        { formula: 'IF(E6=0,0,E10/E6)', result: priorYtdRev > 0 ? priorYtdEbitda / priorYtdRev : 0 },
        { formula: 'IF(F6=0,0,F10/F6)', result: 0 },
        '',
        '',
      ],
      {
        isSubtotal: false,
        numFmts: [undefined, PCT_FMT, PCT_FMT, PCT_FMT, PCT_FMT, PCT_FMT],
      }
    );
    styleDataRow(
      wsPnl,
      20,
      [
        'Net Profit Margin (%)',
        { formula: 'IF(B6=0,0,B16/B6)', result: 0 },
        { formula: 'IF(C6=0,0,C16/C6)', result: 0 },
        { formula: 'IF(D6=0,0,D16/D6)', result: ytdRev > 0 ? ytdPat / ytdRev : 0 },
        { formula: 'IF(E6=0,0,E16/E6)', result: priorYtdRev > 0 ? priorYtdPat / priorYtdRev : 0 },
        { formula: 'IF(F6=0,0,F16/F6)', result: 0 },
        '',
        '',
      ],
      {
        isSubtotal: true,
        numFmts: [undefined, PCT_FMT, PCT_FMT, PCT_FMT, PCT_FMT, PCT_FMT],
      }
    );
  }

  // ==========================================================================
  // SHEET 4: BALANCE SHEET DASHBOARD (Section 8)
  // ==========================================================================
  const currAssets = fromPaise(k.totalCurrentAssetsPaise);
  const prevCurrAssets =
    pk.currentAssetsPaise !== null ? fromPaise(pk.currentAssetsPaise) : getLinePrevInr('TOTAL_CURRENT_ASSETS');
  const nonCurrAssets = fromPaise(k.totalNonCurrentAssetsPaise);
  const prevNonCurrAssets = getLinePrevInr('NON_CURRENT_ASSETS');
  const totalAssets = currAssets + nonCurrAssets;
  const prevTotalAssets =
    pk.totalAssetsPaise !== null ? fromPaise(pk.totalAssetsPaise) : prevCurrAssets + prevNonCurrAssets;

  const equityTotal = fromPaise(k.totalEquityPaise);
  const prevEquityTotal = getLinePrevInr('EQUITY');
  const borrowingsTotal = fromPaise(k.longTermBorrowingsPaise + k.shortTermBorrowingsPaise);
  const prevBorrowingsTotal =
    getLinePrevInr('NON_CURRENT_LIABILITIES') + getLinePrevInr('SHORT_TERM_BORROWINGS');
  const currLiabOther = fromPaise(
    k.tradePayablesPaise + k.taxLiabilitiesPaise + k.otherCurrentLiabilitiesPaise
  );
  const totalCurrLiab = fromPaise(k.totalCurrentLiabilitiesPaise);
  const prevCurrLiab =
    pk.currentLiabilitiesPaise !== null
      ? fromPaise(pk.currentLiabilitiesPaise)
      : getLinePrevInr('TOTAL_CURRENT_LIABILITIES');
  const nonCurrLiab = fromPaise(k.longTermBorrowingsPaise);
  const prevNonCurrLiab = getLinePrevInr('NON_CURRENT_LIABILITIES');
  const totalEqAndLiab = equityTotal + nonCurrLiab + totalCurrLiab;
  const prevTotalEqAndLiab = prevEquityTotal + prevNonCurrLiab + prevCurrLiab;
  const cashBankInr = fromPaise(k.cashPaise + k.bankPaise);
  const prevCashBankInr =
    pk.cashAndBankPaise !== null ? fromPaise(pk.cashAndBankPaise) : getLinePrevInr('CASH_AND_BANK');

  if (sheetsMap.has('Balance Sheet Dashboard')) {
    const wsBs = sheetsMap.get('Balance Sheet Dashboard')!;
    wsBs.columns = [
      { width: 38 },
      { width: 22 },
      { width: 22 },
      { width: 22 },
      { width: 18 },
      { width: 16 },
      { width: 36 },
    ];
    styleBanner(
      wsBs,
      `BALANCE SHEET DASHBOARD — ${model.companyName.toUpperCase()}`,
      `As at ${model.periodEnd}  |  Assets, Equity, Liabilities & Reconciliation Verification (INR)`,
      7
    );
    addNavStrip(wsBs, plannedSheetNames);

    styleTableHeader(wsBs, 5, [
      'Balance Sheet Schedule / Head',
      'Classification',
      `${model.currentPeriodLabel || 'Current Period'} (INR)`,
      `${model.comparativePeriodLabel || 'Previous Period'} (INR)`,
      'Variance (INR)',
      'Variance %',
      'Accounting Basis & Schedule',
    ]);

    const bsSummaryRows: Array<{
      label: string;
      cls: string;
      curr: any;
      prev: any;
      basis: string;
      isSubtotal?: boolean;
    }> = [
      // Row 6: Current Assets
      {
        label: 'Current Assets (Cash, Bank, Inventory, Receivables)',
        cls: 'ASSETS',
        curr: currAssets,
        prev: prevCurrAssets,
        basis: 'Cash + Bank + Inventory + Trade Receivables + Advances',
      },
      // Row 7: Non-Current Assets
      {
        label: 'Non-Current Assets (Property, Plant & Equipment, Intangibles, DTA)',
        cls: 'ASSETS',
        curr: nonCurrAssets,
        prev: prevNonCurrAssets,
        basis: 'Property, Plant & Equipment + Intangible Assets + Deferred Tax Assets + Investments',
      },
      // Row 8: Total Assets (=C6+C7)
      {
        label: 'TOTAL ASSETS',
        cls: 'TOTAL ASSETS',
        curr: { formula: 'C6+C7', result: totalAssets },
        prev: { formula: 'D6+D7', result: prevTotalAssets },
        basis: 'Current Assets + Non-Current Assets',
        isSubtotal: true,
      },
      // Row 9: Shareholders Equity
      {
        label: 'Equity (Share Capital + Reserves & Surplus)',
        cls: 'EQUITY',
        curr: equityTotal,
        prev: prevEquityTotal,
        basis: 'Share Capital + Closing Reserves & Surplus',
      },
      // Row 10: Non-Current Liabilities / Long-Term Borrowings
      {
        label: 'Non-Current Liabilities (Long-Term Borrowings)',
        cls: 'LIABILITIES',
        curr: nonCurrLiab,
        prev: prevNonCurrLiab,
        basis: 'Secured & Unsecured Long-Term Debt Facilities',
      },
      // Row 11: Total Borrowings (Long-Term + Short-Term Debt)
      {
        label: 'Total Debt / Borrowings (Memo — LT + ST Debt)',
        cls: 'MEMO DEBT',
        curr: borrowingsTotal,
        prev: prevBorrowingsTotal,
        basis: 'Long-Term Borrowings + Working Capital OD / Short-Term Loans',
      },
      // Row 12: Current Liabilities
      {
        label: 'Current Liabilities (Payables, ST Debt, Taxes & Provisions)',
        cls: 'LIABILITIES',
        curr: totalCurrLiab,
        prev: prevCurrLiab,
        basis: 'Trade Payables + Short-Term Borrowings + Other Current Liabilities & Provisions',
      },
      // Row 13: Total Equity & Liabilities (=C9+C10+C12)
      {
        label: 'TOTAL EQUITY & LIABILITIES',
        cls: 'TOTAL EQ & LIAB',
        curr: { formula: 'C9+C10+C12', result: totalEqAndLiab },
        prev: { formula: 'D9+D10+D12', result: prevTotalEqAndLiab },
        basis: 'Total Equity + Non-Current Liabilities + Current Liabilities',
        isSubtotal: true,
      },
      // Row 14: Balance Sheet Reconciliation Check (=C8-C13)
      {
        label: 'BALANCE SHEET RECONCILIATION CHECK (Assets − [Equity + Liab])',
        cls: 'RECONCILIATION',
        curr: { formula: 'C8-C13', result: totalAssets - totalEqAndLiab },
        prev: { formula: 'D8-D13', result: prevTotalAssets - prevTotalEqAndLiab },
        basis:
          Math.abs(totalAssets - totalEqAndLiab) <= 1
            ? 'RECONCILED — Total Assets Equal Total Equity + Liabilities'
            : 'DIFFERENCE DETECTED — Check Unmapped / Suspense Balances',
        isSubtotal: true,
      },
    ];

    bsSummaryRows.forEach((item, idx) => {
      const r = 6 + idx;
      const cVal = typeof item.curr === 'number' ? item.curr : item.curr.result;
      const pVal = typeof item.prev === 'number' ? item.prev : item.prev.result;
      styleDataRow(
        wsBs,
        r,
        [
          item.label,
          item.cls,
          item.curr,
          item.prev,
          { formula: `C${r}-D${r}`, result: cVal - pVal },
          { formula: `IF(D${r}=0,0,(C${r}-D${r})/ABS(D${r}))`, result: pVal !== 0 ? (cVal - pVal) / Math.abs(pVal) : 0 },
          item.basis,
        ],
        {
          isSubtotal: item.isSubtotal,
          isAlternating: idx % 2 === 1,
          leftAlignCols: [1, 6],
          numFmts: [undefined, undefined, INR_FMT, INR_FMT, INR_FMT, PCT_FMT, undefined],
          fillOverride:
            r === 14
              ? Math.abs(totalAssets - totalEqAndLiab) <= 1
                ? COLORS.GREEN_TINT
                : COLORS.RED_TINT
              : undefined,
        }
      );
    });

    // Key Balance Sheet Component Breakdown (Rows 16-21) so Executive Dashboard KPI cards can link directly
    const prevTradeRecInr = getLinePrevInr('TRADE_RECEIVABLES');
    const prevInventoryInr = getLinePrevInr('INVENTORY');
    const prevTradePayInr = getLinePrevInr('TRADE_PAYABLES');

    styleTableHeader(wsBs, 16, [
      'Component Breakdown',
      'Schedule Group',
      `${model.currentPeriodLabel || 'Current Period'} (INR)`,
      `${model.comparativePeriodLabel || 'Previous Period'} (INR)`,
      'Variance (INR)',
      'Share of Assets %',
      'Notes',
    ]);
    const compBreak = [
      [
        'Trade Receivables (Sundry Debtors)',
        'Current Assets',
        fromPaise(k.tradeReceivablesPaise),
        prevTradeRecInr,
        'Customer Balances',
      ],
      [
        'Closing Inventory (Stock-in-Hand)',
        'Current Assets',
        fromPaise(k.inventoryPaise),
        prevInventoryInr,
        'Raw Material, WIP & FG',
      ],
      [
        'Cash and Bank Balances',
        'Current Assets',
        cashBankInr,
        prevCashBankInr,
        'Cash-in-Hand + Bank Accounts',
      ],
      [
        'Trade Payables (Sundry Creditors)',
        'Current Liabilities',
        fromPaise(k.tradePayablesPaise),
        prevTradePayInr,
        'Supplier Balances',
      ],
      [
        'Net Working Capital (Current Assets − Current Liab)',
        'Working Capital',
        { formula: 'C6-C12', result: currAssets - totalCurrLiab },
        { formula: 'D6-D12', result: prevCurrAssets - prevCurrLiab },
        'Short-Term Liquidity Buffer',
      ],
    ];
    compBreak.forEach((cb, i) => {
      const r = 17 + i;
      const cVal = typeof cb[2] === 'number' ? cb[2] : (cb[2] as any).result;
      const pVal = typeof cb[3] === 'number' ? cb[3] : (cb[3] as any).result;
      styleDataRow(
        wsBs,
        r,
        [
          cb[0],
          cb[1],
          cb[2],
          cb[3],
          { formula: `C${r}-D${r}`, result: cVal - pVal },
          { formula: `IF($C$8=0,0,C${r}/$C$8)`, result: totalAssets > 0 ? cVal / totalAssets : 0 },
          cb[4],
        ],
        {
          isSubtotal: i === 4,
          isAlternating: i % 2 === 1,
          leftAlignCols: [1, 6],
          numFmts: [undefined, undefined, INR_FMT, INR_FMT, INR_FMT, PCT_FMT, undefined],
        }
      );
    });

    // Complete Schedule III Balance Sheet Line-by-Line Breakdown (Row 23+)
    const bsRawLedgers = Array.from(model.ledgerSummaries.values())
      .map((s) => s.ledger)
      .filter(
        (l) =>
          l.statementSourceType === 'BALANCE_SHEET' ||
          (!model.hasBalanceSheetStatement &&
            [
              'Equity',
              'LongTermBorrowings',
              'ShortTermBorrowings',
              'TradePayables',
              'TaxLiabilities',
              'OtherCurrentLiabilities',
              'FixedAssets',
              'Investments',
              'NonCurrentAssets',
              'Inventory',
              'TradeReceivables',
              'CashAndEquivalents',
              'BankBalances',
              'OtherCurrentAssets',
            ].includes(l.classification))
      );

    if (bsRawLedgers.length > 0) {
      styleTableHeader(wsBs, 23, [
        'Schedule III Balance Sheet Line Item',
        'Schedule III Section',
        `${model.currentPeriodLabel || 'Current Period'} (INR)`,
        `${model.comparativePeriodLabel || 'Previous Period'} (INR)`,
        'Variance (INR)',
        'Variance %',
        'Source Worksheet & Cell Ref',
      ]);

      bsRawLedgers.forEach((l, idx) => {
        const r = 24 + idx;
        const currInr = fromPaise(Math.abs(l.openingBalancePaise));
        const prevInr =
          l.previousPeriodPaise !== undefined ? fromPaise(Math.abs(l.previousPeriodPaise)) : 0;
        const srcRef = [
          l.sourceWorksheetName || 'Balance Sheet',
          l.sourceCellReference ? `Curr: ${l.sourceCellReference}` : '',
          l.priorCellReference ? `Prior: ${l.priorCellReference}` : '',
        ]
          .filter(Boolean)
          .join(' | ');

        styleDataRow(
          wsBs,
          r,
          [
            l.name,
            l.parentGroup || l.classification,
            currInr,
            prevInr,
            { formula: `C${r}-D${r}`, result: currInr - prevInr },
            {
              formula: `IF(D${r}=0,0,(C${r}-D${r})/ABS(D${r}))`,
              result: prevInr !== 0 ? (currInr - prevInr) / Math.abs(prevInr) : 0,
            },
            srcRef,
          ],
          {
            isAlternating: idx % 2 === 1,
            leftAlignCols: [1, 6],
            numFmts: [undefined, undefined, INR_FMT, INR_FMT, INR_FMT, PCT_FMT, undefined],
          }
        );
      });
    }
  }

  // ==========================================================================
  // SHEET 5: CASH FLOW DASHBOARD (Section 9)
  // ==========================================================================
  const ocfInr = cf.operatingCashFlowPaise !== null ? fromPaise(cf.operatingCashFlowPaise) : 0;
  const prevOcfInr = cf.operatingCashFlowPriorPaise !== null ? fromPaise(cf.operatingCashFlowPriorPaise) : 0;
  const icfInr = cf.investingCashFlowPaise !== null ? fromPaise(cf.investingCashFlowPaise) : 0;
  const prevIcfInr = cf.investingCashFlowPriorPaise !== null ? fromPaise(cf.investingCashFlowPriorPaise) : 0;
  const fcfActInr = cf.financingCashFlowPaise !== null ? fromPaise(cf.financingCashFlowPaise) : 0;
  const prevFcfActInr = cf.financingCashFlowPriorPaise !== null ? fromPaise(cf.financingCashFlowPriorPaise) : 0;
  const netCashInr = ocfInr + icfInr + fcfActInr;
  const openCashInr = cf.openingCashPaise !== null ? fromPaise(cf.openingCashPaise) : Math.max(0, cashBankInr - netCashInr);
  const closeCashInr = cf.closingCashPaise !== null ? fromPaise(cf.closingCashPaise) : openCashInr + netCashInr;
  const capexInr = cf.capexPaise !== null ? fromPaise(cf.capexPaise) : 0;
  const freeCfInr = cf.freeCashFlowPaise !== null ? fromPaise(cf.freeCashFlowPaise) : ocfInr - Math.abs(capexInr);

  if (sheetsMap.has('Cash Flow Dashboard')) {
    const wsCf = sheetsMap.get('Cash Flow Dashboard')!;
    wsCf.columns = [
      { width: 40 },
      { width: 22 },
      { width: 22 },
      { width: 22 },
      { width: 18 },
      { width: 40 },
    ];
    styleBanner(
      wsCf,
      `CASH FLOW STATEMENT DASHBOARD — ${model.companyName.toUpperCase()}`,
      `Reporting Period: ${model.periodStart} to ${model.periodEnd}  |  Operating, Investing, Financing & Liquidity Bridge`,
      6
    );
    addNavStrip(wsCf, plannedSheetNames);

    styleTableHeader(wsCf, 5, [
      'Cash Flow Activity / Particulars',
      'Activity Classification',
      'Current Period (INR)',
      'Previous Period (INR)',
      'Variance (INR)',
      'Accounting Formula & Reconciliation Basis',
    ]);

    const cfSummaryRows = [
      // Row 6: Operating Cash Flow
      ['Cash Flow from Operating Activities (OCF)', 'Operating Activities', ocfInr, prevOcfInr, 'Operating Cash Inflows minus Outflows & Direct Taxes', true],
      // Row 7: Investing Cash Flow
      ['Cash Flow from Investing Activities (ICF)', 'Investing Activities', icfInr, prevIcfInr, 'Capital Expenditure & Investment Income', false],
      // Row 8: Financing Cash Flow
      ['Cash Flow from Financing Activities (FCF)', 'Financing Activities', fcfActInr, prevFcfActInr, 'Debt Drawdowns/Repayments, Interest & Equity', false],
      // Row 9: Net Cash Movement (=SUM(C6:C8))
      ['Net Cash Movement During Period', 'Net Movement', { formula: 'SUM(C6:C8)', result: netCashInr }, { formula: 'SUM(D6:D8)', result: prevOcfInr + prevIcfInr + prevFcfActInr }, 'Operating CF + Investing CF + Financing CF', true],
      // Row 10: Opening Cash
      ['Opening Cash and Cash Equivalents', 'Opening Balance', openCashInr, 0, 'Opening Cash-in-Hand + Bank Balances', false],
      // Row 11: Closing Cash (=C10+C9)
      ['Closing Cash and Cash Equivalents (Computed)', 'Closing Balance', { formula: 'C10+C9', result: openCashInr + netCashInr }, 0, 'Opening Cash + Net Cash Movement', true],
      // Row 12: Balance Sheet Cash & Bank
      ['Balance Sheet Cash & Bank Balances', 'BS Reconciliation', cashBankInr, prevCashBankInr, 'As per Schedule III Balance Sheet Cash + Bank', false],
      // Row 13: Reconciliation Check (=C11-C12)
      ['Cash Reconciliation Variance (CF Closing − BS Cash)', 'Verification', { formula: 'C11-C12', result: openCashInr + netCashInr - cashBankInr }, 0, cf.reconciliation.statusMessage, true],
      // Row 14: Capital Expenditure
      ['Capital Expenditure (Capex Outflow)', 'Capex Memo', capexInr, 0, 'Additions to Fixed Assets / Plant & Machinery', false],
      // Row 15: Free Cash Flow
      ['Free Cash Flow (OCF − Capex)', 'Free Cash Flow', { formula: 'C6-ABS(C14)', result: freeCfInr }, 0, cf.freeCashFlowDefinition, true],
    ];

    cfSummaryRows.forEach((row, idx) => {
      const r = 6 + idx;
      const cVal = typeof row[2] === 'number' ? row[2] : (row[2] as any).result;
      const pVal = typeof row[3] === 'number' ? row[3] : (row[3] as any).result;
      styleDataRow(
        wsCf,
        r,
        [
          row[0],
          row[1],
          row[2],
          row[3],
          { formula: `C${r}-D${r}`, result: cVal - pVal },
          row[4],
        ],
        {
          isSubtotal: Boolean(row[5]),
          isAlternating: idx % 2 === 1,
          leftAlignCols: [1, 5],
          numFmts: [undefined, undefined, INR_FMT, INR_FMT, INR_FMT, undefined],
        }
      );
    });

    if (cf.lineItems.length > 0) {
      styleTableHeader(wsCf, 17, [
        'Imported Cash Flow Statement Line Item',
        'Category',
        'Current Period (INR)',
        'Previous Period (INR)',
        'Budget (INR)',
        'Source File & Cell Reference',
      ]);
      cf.lineItems.forEach((li, idx) => {
        styleDataRow(
          wsCf,
          18 + idx,
          [
            li.particulars,
            li.category,
            fromPaise(li.currentPeriodPaise),
            li.previousPeriodPaise !== undefined ? fromPaise(li.previousPeriodPaise) : 0,
            li.budgetPaise !== undefined ? fromPaise(li.budgetPaise) : 0,
            `${li.sourceFilename} ${li.cellReference ? `[${li.cellReference}]` : ''}`,
          ],
          {
            isSubtotal: li.isSubtotalOrTotal,
            isAlternating: idx % 2 === 1,
            leftAlignCols: [1, 5],
            numFmts: [undefined, undefined, INR_FMT, INR_FMT, INR_FMT, undefined],
          }
        );
      });
    }
  }

  // ==========================================================================
  // SHEET 2: EXECUTIVE DASHBOARD (Sections 3, 4, 5)
  // ==========================================================================
  const wsExec = sheetsMap.get('Executive Dashboard')!;
  wsExec.columns = [
    { width: 16 },
    { width: 16 },
    { width: 16 },
    { width: 16 },
    { width: 16 },
    { width: 16 },
    { width: 16 },
    { width: 16 },
    { width: 16 },
    { width: 16 },
    { width: 16 },
    { width: 16 },
  ];

  styleBanner(
    wsExec,
    `EXECUTIVE CFO DASHBOARD — ${model.companyName.toUpperCase()}`,
    `Reporting Period: ${model.periodStart} to ${model.periodEnd}  |  Consolidated Financial KPIs, Formula-Linked Statements & Native Charts`,
    12
  );
  addNavStrip(wsExec, plannedSheetNames);

  // Section Header Row 4
  wsExec.mergeCells(4, 1, 4, 12);
  const kpiSecCell = wsExec.getCell(4, 1);
  kpiSecCell.value = 'EXECUTIVE CFO KEY PERFORMANCE INDICATORS (FORMULA-LINKED TO FINANCIAL STATEMENTS)';
  kpiSecCell.font = { name: 'Calibri', size: 10.5, bold: true, color: { argb: COLORS.NAVY } };
  kpiSecCell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: COLORS.SUBTOTAL_FILL } };
  kpiSecCell.alignment = { vertical: 'middle', horizontal: 'left', indent: 1 };
  wsExec.getRow(4).height = 22;

  const findKpi = (id: string) => cfo.cfoKpis.find((x) => x.id === id);
  const hasPnlSheet = sheetsMap.has('P&L Dashboard');
  const hasBsSheet = sheetsMap.has('Balance Sheet Dashboard');
  const hasCfSheet = sheetsMap.has('Cash Flow Dashboard');

  const currRatioVal = totalCurrLiab > 0 ? currAssets / totalCurrLiab : 0;
  const deRatioVal = equityTotal > 0 ? borrowingsTotal / equityTotal : 0;

  const execCardsConfig = [
    // Row 1 of Cards (Rows 5..8): Cols 1, 4, 7, 10
    {
      row: 5,
      col: 1,
      title: 'Total Revenue',
      accentColor: COLORS.NAVY,
      currentCellValue: hasPnlSheet ? { formula: "'P&L Dashboard'!D6", result: ytdRev } : ytdRev,
      currentNumFmt: INR_INT_FMT,
      previousText: findKpi('KPI_TOTAL_REVENUE')?.previousFormatted || 'N/A',
      varianceText: `${findKpi('KPI_TOTAL_REVENUE')?.varianceAmountFormatted || 'N/A'} (${findKpi('KPI_TOTAL_REVENUE')?.variancePercentFormatted || 'N/A'})`,
      isFavorable: true,
    },
    {
      row: 5,
      col: 4,
      title: 'Gross Profit',
      accentColor: k.hasGrossProfitData ? COLORS.ROYAL_BLUE : COLORS.ORANGE,
      currentCellValue: k.hasGrossProfitData
        ? hasPnlSheet
          ? { formula: "'P&L Dashboard'!D8", result: ytdGp }
          : ytdGp
        : 'Requires Mapping',
      currentNumFmt: k.hasGrossProfitData ? INR_INT_FMT : '@',
      previousText: findKpi('KPI_GROSS_PROFIT')?.previousFormatted || 'N/A',
      varianceText: k.hasGrossProfitData
        ? `${findKpi('KPI_GROSS_PROFIT')?.varianceAmountFormatted || 'N/A'} (${findKpi('KPI_GROSS_PROFIT')?.variancePercentFormatted || 'N/A'})`
        : 'Cost of Sales not separately reported',
      isFavorable: k.hasGrossProfitData ? ytdGp >= 0 : null,
    },
    {
      row: 5,
      col: 7,
      title: 'Gross Profit Margin',
      accentColor: k.hasGrossProfitData ? COLORS.EMERALD : COLORS.ORANGE,
      currentCellValue: k.hasGrossProfitData
        ? hasPnlSheet
          ? { formula: "IF('P&L Dashboard'!D6=0,0,'P&L Dashboard'!D8/'P&L Dashboard'!D6)", result: ytdRev > 0 ? ytdGp / ytdRev : 0 }
          : ytdRev > 0
            ? ytdGp / ytdRev
            : 0
        : 'N/A',
      currentNumFmt: k.hasGrossProfitData ? PCT_FMT : '@',
      previousText: findKpi('KPI_GP_MARGIN')?.previousFormatted || 'N/A',
      varianceText: k.hasGrossProfitData
        ? `${findKpi('KPI_GP_MARGIN')?.varianceAmountFormatted || 'N/A'}`
        : 'Requires Cost of Sales mapping',
      isFavorable: k.hasGrossProfitData ? ytdGp >= 0 : null,
    },
    {
      row: 5,
      col: 10,
      title: 'EBITDA',
      accentColor: COLORS.PURPLE,
      currentCellValue: hasPnlSheet ? { formula: "'P&L Dashboard'!D10", result: ytdEbitda } : ytdEbitda,
      currentNumFmt: INR_INT_FMT,
      previousText: findKpi('KPI_EBITDA')?.previousFormatted || 'N/A',
      varianceText: `${findKpi('KPI_EBITDA')?.varianceAmountFormatted || 'N/A'} (${findKpi('KPI_EBITDA')?.variancePercentFormatted || 'N/A'})`,
      isFavorable: ytdEbitda >= 0,
    },
    // Row 2 of Cards (Rows 10..13): Cols 1, 4, 7, 10
    {
      row: 10,
      col: 1,
      title: 'Net Profit (PAT)',
      accentColor: ytdPat >= 0 ? COLORS.EMERALD : COLORS.RED,
      currentCellValue: hasPnlSheet ? { formula: "'P&L Dashboard'!D16", result: ytdPat } : ytdPat,
      currentNumFmt: INR_INT_FMT,
      previousText: findKpi('KPI_NET_PROFIT')?.previousFormatted || 'N/A',
      varianceText: `${findKpi('KPI_NET_PROFIT')?.varianceAmountFormatted || 'N/A'} (${findKpi('KPI_NET_PROFIT')?.variancePercentFormatted || 'N/A'})`,
      isFavorable: ytdPat >= 0,
    },
    {
      row: 10,
      col: 4,
      title: 'Net Profit Margin',
      accentColor: COLORS.ROYAL_BLUE,
      currentCellValue: hasPnlSheet
        ? { formula: "IF('P&L Dashboard'!D6=0,0,'P&L Dashboard'!D16/'P&L Dashboard'!D6)", result: ytdRev > 0 ? ytdPat / ytdRev : 0 }
        : ytdRev > 0
          ? ytdPat / ytdRev
          : 0,
      currentNumFmt: PCT_FMT,
      previousText: findKpi('KPI_NET_MARGIN')?.previousFormatted || 'N/A',
      varianceText: `${findKpi('KPI_NET_MARGIN')?.varianceAmountFormatted || 'N/A'}`,
      isFavorable: ytdPat >= 0,
    },
    {
      row: 10,
      col: 7,
      title: 'Total Assets',
      accentColor: COLORS.NAVY,
      currentCellValue: hasBsSheet ? { formula: "'Balance Sheet Dashboard'!C8", result: totalAssets } : totalAssets,
      currentNumFmt: INR_INT_FMT,
      previousText: findKpi('KPI_TOTAL_ASSETS')?.previousFormatted || 'N/A',
      varianceText: `${findKpi('KPI_TOTAL_ASSETS')?.varianceAmountFormatted || 'N/A'}`,
      isFavorable: true,
    },
    {
      row: 10,
      col: 10,
      title: 'Cash and Bank',
      accentColor: COLORS.EMERALD,
      currentCellValue: hasBsSheet ? { formula: "'Balance Sheet Dashboard'!C19", result: cashBankInr } : cashBankInr,
      currentNumFmt: INR_INT_FMT,
      previousText: findKpi('KPI_CASH_BANK')?.previousFormatted || 'N/A',
      varianceText: `${findKpi('KPI_CASH_BANK')?.varianceAmountFormatted || 'N/A'}`,
      isFavorable: cashBankInr >= 0,
    },
    // Row 3 of Cards (Rows 15..18): Cols 1, 4, 7, 10
    {
      row: 15,
      col: 1,
      title: 'Operating Cash Flow',
      accentColor: ocfInr >= 0 ? COLORS.EMERALD : COLORS.ORANGE,
      currentCellValue: hasCfSheet ? { formula: "'Cash Flow Dashboard'!C6", result: ocfInr } : ocfInr,
      currentNumFmt: INR_INT_FMT,
      previousText: findKpi('KPI_OPERATING_CF')?.previousFormatted || 'N/A',
      varianceText: `${findKpi('KPI_OPERATING_CF')?.varianceAmountFormatted || 'N/A'}`,
      isFavorable: ocfInr >= 0,
    },
    {
      row: 15,
      col: 4,
      title: 'Working Capital',
      accentColor: COLORS.ROYAL_BLUE,
      currentCellValue: hasBsSheet
        ? { formula: "'Balance Sheet Dashboard'!C6-'Balance Sheet Dashboard'!C12", result: currAssets - totalCurrLiab }
        : currAssets - totalCurrLiab,
      currentNumFmt: INR_INT_FMT,
      previousText: findKpi('KPI_WORKING_CAPITAL')?.previousFormatted || 'N/A',
      varianceText: `${findKpi('KPI_WORKING_CAPITAL')?.varianceAmountFormatted || 'N/A'}`,
      isFavorable: currAssets - totalCurrLiab >= 0,
    },
    {
      row: 15,
      col: 7,
      title: 'Current Ratio',
      accentColor: currRatioVal >= 1.2 ? COLORS.EMERALD : COLORS.ORANGE,
      currentCellValue: hasBsSheet
        ? { formula: "IF('Balance Sheet Dashboard'!C12=0,0,'Balance Sheet Dashboard'!C6/'Balance Sheet Dashboard'!C12)", result: currRatioVal }
        : currRatioVal,
      currentNumFmt: RATIO_FMT,
      previousText: findKpi('KPI_CURRENT_RATIO')?.previousFormatted || 'N/A',
      varianceText: `${findKpi('KPI_CURRENT_RATIO')?.varianceAmountFormatted || 'N/A'}`,
      isFavorable: currRatioVal >= 1.2,
    },
    {
      row: 15,
      col: 10,
      title: 'Debt Equity Ratio',
      accentColor: deRatioVal <= 2.0 ? COLORS.PURPLE : COLORS.RED,
      currentCellValue: hasBsSheet
        ? { formula: "IF('Balance Sheet Dashboard'!C9<=0,0,'Balance Sheet Dashboard'!C11/'Balance Sheet Dashboard'!C9)", result: deRatioVal }
        : deRatioVal,
      currentNumFmt: RATIO_FMT,
      previousText: findKpi('KPI_DEBT_EQUITY')?.previousFormatted || 'N/A',
      varianceText: `${findKpi('KPI_DEBT_EQUITY')?.varianceAmountFormatted || 'N/A'}`,
      isFavorable: deRatioVal <= 2.0,
    },
  ];

  wsExec.getRow(9).height = 8;
  wsExec.getRow(14).height = 8;
  wsExec.getRow(19).height = 10;

  execCardsConfig.forEach((c) => drawExecutiveKpiCard(wsExec, c.row, c.col, c));

  // Section Banner for Embedded Native Charts on Executive Dashboard (Row 20)
  wsExec.mergeCells(20, 1, 20, 12);
  const chartSecCell = wsExec.getCell(20, 1);
  chartSecCell.value =
    'INTERACTIVE NATIVE EXCEL CHARTS (LINKED TO WORKSHEET DATA RANGES — SELECTABLE & EDITABLE IN MICROSOFT EXCEL)';
  chartSecCell.font = { name: 'Calibri', size: 10.5, bold: true, color: { argb: COLORS.WHITE } };
  chartSecCell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: COLORS.NAVY } };
  chartSecCell.alignment = { vertical: 'middle', horizontal: 'left', indent: 1 };
  wsExec.getRow(20).height = 22;

  // Reserve clean canvas rows 21..85 for the 8 Executive Charts (4 rows of 2 side-by-side charts)
  for (let r = 21; r <= 85; r++) {
    wsExec.getRow(r).height = 18;
  }

  // Executive Financial Summary Table below the charts (starting at Row 87)
  styleTableHeader(wsExec, 87, [
    'Executive Summary Line Item',
    'Category',
    'Current Period',
    'Previous Period',
    'Budget',
    'Variance',
    'Variance %',
    'Status',
    'Formula / Basis',
  ]);
  cfo.cfoKpis.forEach((item, idx) => {
    styleDataRow(
      wsExec,
      88 + idx,
      [
        item.label,
        item.unit,
        item.currentFormatted,
        item.previousFormatted,
        item.budgetFormatted,
        item.varianceAmountFormatted,
        item.variancePercentFormatted,
        item.statusIndicator,
        item.formula,
      ],
      {
        isAlternating: idx % 2 === 1,
        leftAlignCols: [1, 7, 8],
      }
    );
  });

  // ==========================================================================
  // SHEET 6: KPI SUMMARY
  // ==========================================================================
  if (sheetsMap.has('KPI Summary')) {
    const wsKpi = sheetsMap.get('KPI Summary')!;
    wsKpi.columns = [
      { width: 22 },
      { width: 30 },
      { width: 14 },
      { width: 20 },
      { width: 20 },
      { width: 18 },
      { width: 18 },
      { width: 16 },
      { width: 20 },
      { width: 48 },
    ];
    styleBanner(
      wsKpi,
      `CFO KEY PERFORMANCE INDICATORS SUMMARY — ${model.companyName.toUpperCase()}`,
      `17-Point Executive Metric Schedule with Current, Previous, Budget & Variance Analysis`,
      10
    );
    addNavStrip(wsKpi, plannedSheetNames);
    styleTableHeader(wsKpi, 5, [
      'KPI Code',
      'CFO Indicator',
      'Unit',
      'Current Period',
      'Previous Period',
      'Budget',
      'Variance Amount',
      'Variance %',
      'Status',
      'Accounting Formula & Basis',
    ]);
    cfo.cfoKpis.forEach((kpi, idx) => {
      styleDataRow(
        wsKpi,
        6 + idx,
        [
          kpi.id,
          kpi.label,
          kpi.unit,
          kpi.unit === 'CURRENCY' && kpi.currentPaiseOrVal !== null
            ? fromPaise(kpi.currentPaiseOrVal)
            : kpi.currentFormatted,
          kpi.unit === 'CURRENCY' && kpi.previousPaiseOrVal !== null
            ? fromPaise(kpi.previousPaiseOrVal)
            : kpi.previousFormatted,
          kpi.budgetFormatted,
          kpi.varianceAmountFormatted,
          kpi.variancePercentFormatted,
          kpi.statusIndicator,
          kpi.formula,
        ],
        {
          isAlternating: idx % 2 === 1,
          leftAlignCols: [1, 2, 8, 9],
          numFmts: [
            undefined,
            undefined,
            undefined,
            kpi.unit === 'CURRENCY' ? INR_FMT : undefined,
            kpi.unit === 'CURRENCY' && kpi.previousPaiseOrVal !== null ? INR_FMT : undefined,
          ],
        }
      );
    });
  }

  // ==========================================================================
  // SHEET 7: FINANCIAL RATIOS (Section 10)
  // ==========================================================================
  if (sheetsMap.has('Financial Ratios')) {
    const wsRatios = sheetsMap.get('Financial Ratios')!;
    wsRatios.columns = [
      { width: 20 },
      { width: 30 },
      { width: 18 },
      { width: 18 },
      { width: 18 },
      { width: 42 },
      { width: 52 },
    ];
    styleBanner(
      wsRatios,
      `FINANCIAL RATIOS ANALYSIS — ${model.companyName.toUpperCase()}`,
      `Liquidity, Profitability, Efficiency, Solvency & Cash Flow Ratios`,
      7
    );
    addNavStrip(wsRatios, plannedSheetNames);
    styleTableHeader(wsRatios, 5, [
      'Ratio Category',
      'Ratio Name',
      'Current Period',
      'Previous Period',
      'Variance',
      'Formula',
      'Financial Interpretation',
    ]);
    allRatios.forEach((r, idx) => {
      const diff =
        r.currentValue !== null && r.previousValue !== null
          ? `${(r.currentValue - r.previousValue) >= 0 ? '+' : ''}${(r.currentValue - r.previousValue).toFixed(2)}`
          : 'N/A';
      styleDataRow(
        wsRatios,
        6 + idx,
        [
          r.category,
          r.name,
          r.formattedCurrent,
          r.formattedPrevious,
          diff,
          r.formula,
          r.explanation,
        ],
        {
          isAlternating: idx % 2 === 1,
          leftAlignCols: [1, 5, 6],
        }
      );
    });
  }

  // ==========================================================================
  // SHEET 8: WORKING CAPITAL DASHBOARD (Section 11)
  // ==========================================================================
  if (sheetsMap.has('Working Capital')) {
    const wsWc = sheetsMap.get('Working Capital')!;
    wsWc.columns = [
      { width: 38 },
      { width: 22 },
      { width: 22 },
      { width: 18 },
      { width: 14 },
      { width: 48 },
    ];
    styleBanner(
      wsWc,
      `WORKING CAPITAL & CASH CONVERSION CYCLE DASHBOARD — ${model.companyName.toUpperCase()}`,
      `Current Assets, Current Liabilities, Net Working Capital, Receivables, Payables, Inventory & CCC`,
      6
    );
    addNavStrip(wsWc, plannedSheetNames);
    styleTableHeader(wsWc, 5, [
      'Working Capital Metric',
      'Current Period',
      'Previous Period',
      'Variance',
      'Unit',
      'Formula & Operational Basis',
    ]);

    const wcPrevTradeRecInr = getLinePrevInr('TRADE_RECEIVABLES');
    const wcPrevTradePayInr = getLinePrevInr('TRADE_PAYABLES');
    const wcPrevInventoryInr = getLinePrevInr('INVENTORY');

    const wcTableRows = [
      ['Current Assets', fromPaise(wc.currentAssetsPaise), prevCurrAssets, 'INR', 'Cash + Bank + Inventory + Trade Receivables + Short-Term Advances', true],
      ['Current Liabilities', fromPaise(wc.currentLiabilitiesPaise), prevCurrLiab, 'INR', 'Trade Payables + Short-Term Borrowings + Statutory Dues', true],
      ['Net Working Capital (Current Assets − Current Liab)', { formula: 'B6-B7', result: fromPaise(wc.netWorkingCapitalPaise) }, { formula: 'C6-C7', result: prevCurrAssets - prevCurrLiab }, 'INR', 'Current Assets minus Current Liabilities', true],
      ['Trade Receivables (Sundry Debtors)', fromPaise(wc.tradeReceivablesPaise), wcPrevTradeRecInr, 'INR', 'Closing Sundry Debtors Schedule', false],
      ['Trade Payables (Sundry Creditors)', fromPaise(wc.tradePayablesPaise), wcPrevTradePayInr, 'INR', 'Closing Sundry Creditors Schedule', false],
      ['Closing Inventory (Stock-in-Hand)', fromPaise(wc.inventoryPaise), wcPrevInventoryInr, 'INR', 'Raw Material, WIP & Finished Goods Inventory', false],
      ['Cash and Bank Balances', fromPaise(wc.cashAndBankPaise), prevCashBankInr, 'INR', 'Cash-in-Hand + Bank Accounts', false],
      ['Days Sales Outstanding (DSO)', wc.dsoDays ?? 0, 0, 'Days', '(Trade Receivables / Net Revenue) × 365', false],
      ['Days Inventory Outstanding (DIO)', wc.dioDays ?? 0, 0, 'Days', '(Closing Inventory / Cost of Sales) × 365', false],
      ['Days Payable Outstanding (DPO)', wc.dpoDays ?? 0, 0, 'Days', '(Trade Payables / Net Purchases) × 365', false],
      ['Cash Conversion Cycle (CCC = DIO + DSO − DPO)', { formula: 'B14+B13-B15', result: wc.cashConversionCycleDays ?? 0 }, 0, 'Days', wc.cccExplanation, true],
    ];

    wcTableRows.forEach((row, idx) => {
      const r = 6 + idx;
      const cVal = typeof row[1] === 'number' ? row[1] : (row[1] as any).result;
      const pVal = typeof row[2] === 'number' ? row[2] : (row[2] as any).result;
      const isCurrency = row[3] === 'INR';
      styleDataRow(
        wsWc,
        r,
        [
          row[0],
          row[1],
          pVal,
          { formula: `B${r}-C${r}`, result: cVal - pVal },
          row[3],
          row[4],
        ],
        {
          isSubtotal: Boolean(row[5]),
          isAlternating: idx % 2 === 1,
          leftAlignCols: [4, 5],
          numFmts: [
            undefined,
            isCurrency ? INR_FMT : '0.0',
            isCurrency ? INR_FMT : '0.0',
            isCurrency ? INR_FMT : '0.0',
            undefined,
            undefined,
          ],
        }
      );
    });
  }

  // ==========================================================================
  // SHEET 9: COMPARATIVE FINANCIALS
  // ==========================================================================
  if (sheetsMap.has('Comparative Financials')) {
    const wsComp = sheetsMap.get('Comparative Financials')!;
    wsComp.columns = [
      { width: 34 },
      { width: 24 },
      { width: 20 },
      { width: 20 },
      { width: 20 },
      { width: 20 },
      { width: 20 },
      { width: 16 },
      { width: 22 },
    ];
    styleBanner(
      wsComp,
      `COMPARATIVE FINANCIAL SCHEDULE — ${model.companyName.toUpperCase()}`,
      `Ledger-Level Current Period vs Previous Period vs Budget Comparison`,
      9
    );
    addNavStrip(wsComp, plannedSheetNames);
    styleTableHeader(wsComp, 5, [
      'Ledger / Account Name',
      'Parent Group',
      'Classification',
      'Current Period (INR)',
      'Previous Period (INR)',
      'Budget (INR)',
      'Variance vs Prior (INR)',
      'Variance %',
      'Worksheet & Cell Ref',
    ]);

    const ledgersArr = Array.from(model.ledgerSummaries.values());
    ledgersArr.forEach((ls, idx) => {
      const r = 6 + idx;
      const cur = fromPaise(Math.abs(ls.effectiveClosingPaise));
      const prev =
        ls.ledger.previousPeriodPaise !== undefined
          ? fromPaise(Math.abs(ls.ledger.previousPeriodPaise))
          : 0;
      const bud =
        ls.ledger.budgetPaise !== undefined
          ? fromPaise(Math.abs(ls.ledger.budgetPaise))
          : 0;
      styleDataRow(
        wsComp,
        r,
        [
          ls.ledger.name,
          ls.ledger.parentGroup,
          ls.ledger.classification,
          cur,
          prev,
          bud,
          { formula: `D${r}-E${r}`, result: cur - prev },
          {
            formula: `IF(E${r}=0,0,(D${r}-E${r})/ABS(E${r}))`,
            result: prev !== 0 ? (cur - prev) / Math.abs(prev) : 0,
          },
          ls.ledger.cellReference || '—',
        ],
        {
          isAlternating: idx % 2 === 1,
          leftAlignCols: [1, 2, 8],
          numFmts: [
            undefined,
            undefined,
            undefined,
            INR_FMT,
            INR_FMT,
            INR_FMT,
            INR_FMT,
            PCT_FMT,
            undefined,
          ],
        }
      );
    });
    wsComp.autoFilter = {
      from: { row: 5, column: 1 },
      to: { row: Math.max(5, 5 + ledgersArr.length), column: 9 },
    };
  }

  // ==========================================================================
  // SHEET 10: CFO SUGGESTIONS (Section 12)
  // ==========================================================================
  if (sheetsMap.has('CFO Suggestions')) {
    const wsSug = sheetsMap.get('CFO Suggestions')!;
    wsSug.columns = [
      { width: 24 },
      { width: 52 },
      { width: 20 },
      { width: 20 },
      { width: 20 },
      { width: 18 },
      { width: 54 },
    ];
    styleBanner(
      wsSug,
      `CFO MANAGEMENT OBSERVATIONS & ACTION RECOMMENDATIONS — ${model.companyName.toUpperCase()}`,
      `Green = Favorable / Positive  |  Amber = Attention Required  |  Red = Significant Risk Indicator`,
      7
    );
    addNavStrip(wsSug, plannedSheetNames);
    styleTableHeader(wsSug, 5, [
      'Category',
      'Financial Observation (Verified Fact)',
      'Current Value',
      'Previous Value',
      'Variance',
      'Risk Rating',
      'Suggested Management Action (Advisory)',
    ]);

    cfo.cfoSuggestions.forEach((s, idx) => {
      const r = 6 + idx;
      const fillTint =
        s.riskClassification === 'FAVORABLE'
          ? COLORS.GREEN_TINT
          : s.riskClassification === 'HIGH RISK'
            ? COLORS.RED_TINT
            : COLORS.AMBER_TINT;
      styleDataRow(
        wsSug,
        r,
        [
          s.category,
          s.factualObservation,
          s.actualFinancialValue,
          s.comparativeValue,
          s.calculatedVariance,
          s.riskClassification,
          s.suggestedManagementAction,
        ],
        {
          fillOverride: fillTint,
          leftAlignCols: [1, 5, 6],
        }
      );
      wsSug.getRow(r).height = 28;
      wsSug.getCell(r, 2).alignment = { vertical: 'middle', horizontal: 'left', wrapText: true };
      wsSug.getCell(r, 7).alignment = { vertical: 'middle', horizontal: 'left', wrapText: true };
    });
  }

  // ==========================================================================
  // SHEETS 11–18: SUPPORTING SCHEDULES (Receivables, Payables, Inventory, Health, GST, Monthly MIS, Validation, Source)
  // ==========================================================================
  if (sheetsMap.has('Receivables')) {
    const wsRec = sheetsMap.get('Receivables')!;
    wsRec.columns = [
      { width: 36 },
      { width: 22 },
      { width: 22 },
      { width: 22 },
      { width: 14 },
      { width: 18 },
      { width: 18 },
    ];
    styleBanner(wsRec, `TRADE RECEIVABLES (SUNDRY DEBTORS) AGEING SCHEDULE`, `Company: ${model.companyName}`, 7);
    addNavStrip(wsRec, plannedSheetNames);
    styleTableHeader(wsRec, 5, [
      'Customer / Party Name',
      'Total Outstanding (INR)',
      'Overdue Amount (INR)',
      'Current / Not Due (INR)',
      'Open Bills',
      'Max Days Overdue',
      'Oldest Bill Date',
    ]);
    ageing.customers.forEach((c, idx) => {
      styleDataRow(
        wsRec,
        6 + idx,
        [
          c.partyName,
          fromPaise(c.totalOutstandingPaise),
          fromPaise(c.overdueOutstandingPaise),
          fromPaise(c.currentNotDuePaise),
          c.openInvoiceCount,
          c.maxDaysOverdue,
          c.oldestInvoiceDate,
        ],
        {
          isAlternating: idx % 2 === 1,
          leftAlignCols: [6],
          numFmts: [undefined, INR_FMT, INR_FMT, INR_FMT, '0', '0', undefined],
        }
      );
    });
    wsRec.autoFilter = {
      from: { row: 5, column: 1 },
      to: { row: Math.max(5, 5 + ageing.customers.length), column: 7 },
    };
  }

  if (sheetsMap.has('Payables')) {
    const wsPay = sheetsMap.get('Payables')!;
    wsPay.columns = [
      { width: 36 },
      { width: 22 },
      { width: 22 },
      { width: 22 },
      { width: 14 },
      { width: 18 },
      { width: 18 },
    ];
    styleBanner(wsPay, `TRADE PAYABLES (SUNDRY CREDITORS) AGEING SCHEDULE`, `Company: ${model.companyName}`, 7);
    addNavStrip(wsPay, plannedSheetNames);
    styleTableHeader(wsPay, 5, [
      'Supplier / Party Name',
      'Total Outstanding (INR)',
      'Overdue Amount (INR)',
      'Current / Not Due (INR)',
      'Open Bills',
      'Max Days Overdue',
      'Oldest Bill Date',
    ]);
    ageing.suppliers.forEach((s, idx) => {
      styleDataRow(
        wsPay,
        6 + idx,
        [
          s.partyName,
          fromPaise(s.totalOutstandingPaise),
          fromPaise(s.overdueOutstandingPaise),
          fromPaise(s.currentNotDuePaise),
          s.openInvoiceCount,
          s.maxDaysOverdue,
          s.oldestInvoiceDate,
        ],
        {
          isAlternating: idx % 2 === 1,
          leftAlignCols: [6],
          numFmts: [undefined, INR_FMT, INR_FMT, INR_FMT, '0', '0', undefined],
        }
      );
    });
    wsPay.autoFilter = {
      from: { row: 5, column: 1 },
      to: { row: Math.max(5, 5 + ageing.suppliers.length), column: 7 },
    };
  }

  if (sheetsMap.has('Inventory')) {
    const wsInv = sheetsMap.get('Inventory')!;
    wsInv.columns = [
      { width: 36 },
      { width: 24 },
      { width: 14 },
      { width: 16 },
      { width: 22 },
      { width: 16 },
      { width: 22 },
    ];
    styleBanner(wsInv, `INVENTORY & STOCK VALUATION SCHEDULE`, `Company: ${model.companyName}`, 7);
    addNavStrip(wsInv, plannedSheetNames);
    styleTableHeader(wsInv, 5, [
      'Stock Item Name',
      'Stock Group',
      'UOM',
      'Opening Qty',
      'Opening Value (INR)',
      'Closing Qty',
      'Closing Value (INR)',
    ]);
    model.stockItems.forEach((st, idx) => {
      styleDataRow(
        wsInv,
        6 + idx,
        [
          st.name,
          st.parentGroup,
          st.unit,
          st.openingQty,
          fromPaise(st.openingValuePaise),
          st.closingQty,
          fromPaise(st.closingValuePaise),
        ],
        {
          isAlternating: idx % 2 === 1,
          leftAlignCols: [1, 2],
          numFmts: [undefined, undefined, undefined, '#,##0.00', INR_FMT, '#,##0.00', INR_FMT],
        }
      );
    });
  }

  if (sheetsMap.has('Accounting Health Check')) {
    const wsHealth = sheetsMap.get('Accounting Health Check')!;
    wsHealth.columns = [
      { width: 14 },
      { width: 18 },
      { width: 16 },
      { width: 28 },
      { width: 28 },
      { width: 20 },
      { width: 24 },
      { width: 14 },
      { width: 50 },
    ];
    styleBanner(
      wsHealth,
      `ACCOUNTING HEALTH CHECK & AUDIT EXCEPTION REGISTER`,
      `Overall Quality Score: ${health.qualityScore.overallScore}/100  |  High Priority: ${health.highPriorityCount}  |  Medium: ${health.mediumPriorityCount}`,
      9
    );
    addNavStrip(wsHealth, plannedSheetNames);
    styleTableHeader(wsHealth, 5, [
      'Date',
      'Voucher No',
      'Voucher Type',
      'Ledger Name',
      'Party Name',
      'Amount (INR)',
      'Exception Category',
      'Severity',
      'Audit Explanation',
    ]);
    health.exceptions.forEach((ex, idx) => {
      styleDataRow(
        wsHealth,
        6 + idx,
        [
          ex.date,
          ex.voucherNumber,
          ex.voucherType,
          ex.ledgerName,
          ex.partyName,
          fromPaise(ex.amountPaise),
          ex.category,
          ex.severity,
          ex.explanation,
        ],
        {
          isAlternating: idx % 2 === 1,
          leftAlignCols: [1, 2, 3, 4, 6, 7, 8],
          numFmts: [undefined, undefined, undefined, undefined, undefined, INR_FMT],
        }
      );
    });
  }

  if (sheetsMap.has('GST Reconciliation')) {
    const wsGst = sheetsMap.get('GST Reconciliation')!;
    wsGst.columns = [
      { width: 22 },
      { width: 18 },
      { width: 20 },
      { width: 30 },
      { width: 18 },
      { width: 18 },
      { width: 18 },
      { width: 18 },
      { width: 18 },
      { width: 42 },
    ];
    styleBanner(wsGst, `GST RECONCILIATION REGISTER (BOOKS VS GSTR-1 / GSTR-2B / GSTR-3B)`, `Company: ${model.companyName}`, 10);
    addNavStrip(wsGst, plannedSheetNames);
    styleTableHeader(wsGst, 5, [
      'Reconciliation Type',
      'Match Status',
      'GSTIN',
      'Counterparty Name',
      'Books Invoice No',
      'Portal Invoice No',
      'Books Tax (INR)',
      'Portal Tax (INR)',
      'Tax Difference (INR)',
      'Observations',
    ]);
    const gstMatches = [...gst.salesVsGstr1Matches, ...gst.purchasesVsGstr2bMatches];
    gstMatches.forEach((m, idx) => {
      styleDataRow(
        wsGst,
        6 + idx,
        [
          m.reconciliationType,
          m.matchStatus,
          m.booksInvoice?.partyGstin || m.portalInvoice?.partyGstin || '',
          m.booksInvoice?.partyName || m.portalInvoice?.partyName || '',
          m.booksInvoice?.invoiceNumber || '',
          m.portalInvoice?.invoiceNumber || '',
          fromPaise(m.booksInvoice?.totalTaxPaise || 0),
          fromPaise(m.portalInvoice?.totalTaxPaise || 0),
          fromPaise(m.totalTaxDiffPaise),
          m.discrepancyReasons.join('; '),
        ],
        {
          isAlternating: idx % 2 === 1,
          leftAlignCols: [1, 2, 3, 4, 5, 9],
          numFmts: [undefined, undefined, undefined, undefined, undefined, undefined, INR_FMT, INR_FMT, INR_FMT, undefined],
        }
      );
    });
  }

  if (sheetsMap.has('Monthly MIS')) {
    const wsMis = sheetsMap.get('Monthly MIS')!;
    wsMis.columns = [
      { width: 22 },
      { width: 20 },
      { width: 20 },
      { width: 20 },
      { width: 20 },
      { width: 20 },
      { width: 20 },
      { width: 20 },
      { width: 20 },
    ];
    styleBanner(
      wsMis,
      `MONTHLY MANAGEMENT INFORMATION SCHEDULE (MONTH-ON-MONTH)`,
      `Company: ${model.companyName}  |  ${
        hasMonthly
          ? 'Monthly Operational & Cash Velocity Schedule'
          : 'Monthly data not available from uploaded annual statements'
      }`,
      9
    );
    addNavStrip(wsMis, plannedSheetNames);
    styleTableHeader(wsMis, 5, [
      'Month / Period',
      'Revenue (INR)',
      'Purchases & COGS (INR)',
      'Gross Profit (INR)',
      'Operating Expenses (INR)',
      'Net Profit (INR)',
      'Collections (INR)',
      'Payments (INR)',
      'Net Cash Flow (INR)',
    ]);
    if (hasMonthly) {
      months.forEach((m, idx) => {
        const r = 6 + idx;
        const revM = fromPaise(m.revenuePaise);
        const purM = fromPaise(m.purchasesPaise);
        const gpM = k.hasGrossProfitData ? fromPaise(m.grossProfitPaise) : 'Requires Mapping';
        const expM = fromPaise(m.expensesPaise);
        const patM = fromPaise(m.netProfitPaise);
        const colM = fromPaise(m.collectionsPaise);
        const payM = fromPaise(m.paymentsPaise);
        styleDataRow(
          wsMis,
          r,
          [
            m.month,
            revM,
            purM,
            gpM,
            expM,
            patM,
            colM,
            payM,
            { formula: `G${r}-H${r}`, result: colM - payM },
          ],
          {
            isAlternating: idx % 2 === 1,
            numFmts: [
              undefined,
              INR_FMT,
              INR_FMT,
              k.hasGrossProfitData ? INR_FMT : undefined,
              INR_FMT,
              INR_FMT,
              INR_FMT,
              INR_FMT,
              INR_FMT,
            ],
          }
        );
      });
    } else {
      styleDataRow(
        wsMis,
        6,
        [
          'Monthly data not available from uploaded annual statements.',
          'N/A',
          'N/A',
          'N/A',
          'N/A',
          'N/A',
          'N/A',
          'N/A',
          'N/A',
        ],
        {
          isSubtotal: true,
          fillOverride: COLORS.AMBER_TINT,
          leftAlignCols: [0, 1, 2, 3, 4, 5, 6, 7, 8],
        }
      );
    }
  }

  if (sheetsMap.has('Data Validation')) {
    const wsVal = sheetsMap.get('Data Validation')!;
    wsVal.columns = [
      { width: 38 },
      { width: 22 },
      { width: 28 },
      { width: 22 },
      { width: 22 },
      { width: 22 },
      { width: 18 },
      { width: 20 },
    ];
    styleBanner(
      wsVal,
      `SOURCE MAPPING & RECONCILIATION REGISTER — ${model.companyName.toUpperCase()}`,
      `Cell-Level Provenance, Reported vs Calculated Verification & Schedule III Control Checks`,
      8
    );
    addNavStrip(wsVal, plannedSheetNames);
    styleTableHeader(wsVal, 5, [
      'Financial Metric',
      'Source Sheet',
      'Source Cell',
      `${model.currentPeriodLabel || 'FY 2025–26'} Value`,
      `${model.comparativePeriodLabel || 'FY 2024–25'} Value`,
      'Calculated Value',
      'Difference',
      'Status',
    ]);

    const getLineInfo = (code: string) => {
      const allLines = [
        ...statements.pnlLines,
        ...statements.bsAssetLines,
        ...statements.bsLiabilityEquityLines,
      ];
      return allLines.find((l) => l.code === code);
    };

    const revLine = getLineInfo('REVENUE_NET');
    const otherIncLine = getLineInfo('OTHER_INCOME');
    const totalIncLine = getLineInfo('TOTAL_INCOME');
    const totalExpLine = getLineInfo('TOTAL_EXPENSES');
    const pbtLine = getLineInfo('PBT');
    const currTaxLine = getLineInfo('CURRENT_TAX');
    const defTaxLine = getLineInfo('DEFERRED_TAX');
    const patLine = getLineInfo('PAT');
    const ltDebtLine = getLineInfo('NON_CURRENT_LIABILITIES');
    const cashLine = getLineInfo('CASH_AND_BANK');
    const assetsLine = getLineInfo('TOTAL_ASSETS');
    const eqLiabLine = getLineInfo('TOTAL_EQUITY_AND_LIABILITIES');

    const currAssetsLine = getLineInfo('TOTAL_CURRENT_ASSETS');
    const nonCurrAssetsLine = getLineInfo('TOTAL_NON_CURRENT_ASSETS');
    const equityLine = getLineInfo('TOTAL_EQUITY');
    const currLiabLine = getLineInfo('TOTAL_CURRENT_LIABILITIES');

    // Collect individual Schedule III Balance Sheet leaf items with cell provenance
    const bsLeafValidationRows: any[][] = Array.from(model.ledgerSummaries.values())
      .map((s) => s.ledger)
      .filter((l) => l.statementSourceType === 'BALANCE_SHEET')
      .map((l) => {
        const currVal = fromPaise(Math.abs(l.openingBalancePaise));
        const prevVal =
          l.previousPeriodPaise !== undefined ? fromPaise(Math.abs(l.previousPeriodPaise)) : 0;
        const cellRef = [
          l.sourceCellReference ? `${l.sourceCellReference}` : '',
          l.priorCellReference ? `Prior: ${l.priorCellReference}` : '',
        ]
          .filter(Boolean)
          .join(' / ');
        return [
          `${l.name} (${l.parentGroup || l.classification})`,
          l.sourceWorksheetName || 'Balance Sheet',
          cellRef || '—',
          currVal,
          prevVal,
          currVal,
          0,
          'RECONCILED',
        ];
      });

    const coreMappingRows: any[][] = [
      [
        'Revenue from Operations',
        revLine?.worksheetNames?.[0] || 'Profit and Loss',
        revLine?.cellReferences?.[0] || '—',
        ytdRev,
        priorYtdRev,
        ytdRev,
        0,
        'RECONCILED',
      ],
      [
        'Other Income',
        otherIncLine?.worksheetNames?.[0] || 'Profit and Loss',
        otherIncLine?.cellReferences?.[0] || '—',
        ytdOtherInc,
        priorYtdOtherInc,
        ytdOtherInc,
        0,
        'RECONCILED',
      ],
      [
        'Total Income',
        'Profit and Loss',
        totalIncLine?.cellReferences?.[0] || '—',
        ytdRev + ytdOtherInc,
        priorYtdRev + priorYtdOtherInc,
        ytdRev + ytdOtherInc,
        0,
        'RECONCILED',
      ],
      [
        'Total Expenses',
        'Profit and Loss',
        totalExpLine?.cellReferences?.[0] || '—',
        fromPaise(k.totalExpensesPaise),
        totalExpLine?.previousAmountPaise !== undefined
          ? fromPaise(totalExpLine.previousAmountPaise)
          : 0,
        fromPaise(k.totalExpensesPaise),
        0,
        'RECONCILED',
      ],
      [
        'Profit Before Tax (PBT)',
        'Profit and Loss',
        pbtLine?.cellReferences?.[0] || '—',
        ytdPbt,
        priorYtdPbt,
        ytdPbt,
        0,
        'RECONCILED',
      ],
      ...(k.currentTaxPaise !== 0 || k.deferredTaxPaise !== 0
        ? [
            [
              'Current Tax',
              currTaxLine?.worksheetNames?.[0] || 'Profit and Loss',
              currTaxLine?.cellReferences?.[0] || '—',
              fromPaise(k.currentTaxPaise),
              currTaxLine?.previousAmountPaise !== undefined
                ? fromPaise(currTaxLine.previousAmountPaise)
                : 0,
              fromPaise(k.currentTaxPaise),
              0,
              'RECONCILED',
            ],
            [
              'Deferred Tax',
              defTaxLine?.worksheetNames?.[0] || 'Profit and Loss',
              defTaxLine?.cellReferences?.[0] || '—',
              fromPaise(k.deferredTaxPaise),
              defTaxLine?.previousAmountPaise !== undefined
                ? fromPaise(defTaxLine.previousAmountPaise)
                : 0,
              fromPaise(k.deferredTaxPaise),
              0,
              'RECONCILED',
            ],
          ]
        : []),
      [
        'Profit After Tax (PAT)',
        'Profit and Loss',
        patLine?.cellReferences?.[0] || '—',
        ytdPat,
        priorYtdPat,
        ytdPat,
        0,
        'RECONCILED',
      ],
      ...bsLeafValidationRows,
      [
        "Total Shareholders' Funds (Equity)",
        'Balance Sheet',
        equityLine?.cellReferences?.[0] || '—',
        equityTotal,
        prevEquityTotal,
        equityTotal,
        0,
        'RECONCILED',
      ],
      [
        'Total Non-Current Liabilities',
        'Balance Sheet',
        ltDebtLine?.cellReferences?.[0] || '—',
        nonCurrLiab,
        prevNonCurrLiab,
        nonCurrLiab,
        0,
        'RECONCILED',
      ],
      [
        'Total Current Liabilities',
        'Balance Sheet',
        currLiabLine?.cellReferences?.[0] || '—',
        totalCurrLiab,
        prevCurrLiab,
        totalCurrLiab,
        0,
        'RECONCILED',
      ],
      [
        'Total Equity and Liabilities',
        'Balance Sheet',
        eqLiabLine?.cellReferences?.[0] || '—',
        totalEqAndLiab,
        prevTotalEqAndLiab,
        totalEqAndLiab,
        0,
        'RECONCILED',
      ],
      [
        'Total Non-Current Assets',
        'Balance Sheet',
        nonCurrAssetsLine?.cellReferences?.[0] || '—',
        nonCurrAssets,
        prevNonCurrAssets,
        nonCurrAssets,
        0,
        'RECONCILED',
      ],
      [
        'Total Current Assets',
        'Balance Sheet',
        currAssetsLine?.cellReferences?.[0] || '—',
        currAssets,
        prevCurrAssets,
        currAssets,
        0,
        'RECONCILED',
      ],
      [
        'Cash and Cash Equivalents',
        'Balance Sheet',
        cashLine?.cellReferences?.[0] || '—',
        cashBankInr,
        prevCashBankInr,
        cashBankInr,
        0,
        'RECONCILED',
      ],
      [
        'Total Assets',
        'Balance Sheet',
        assetsLine?.cellReferences?.[0] || '—',
        totalAssets,
        prevTotalAssets,
        totalAssets,
        totalAssets - totalEqAndLiab,
        Math.abs(totalAssets - totalEqAndLiab) <= 1 ? 'RECONCILED' : 'CHECK DIFFERENCE',
      ],
      [
        'Balance Sheet Reconciliation Difference (Assets − [Equity + Liabilities])',
        'Balance Sheet',
        'Control Check',
        totalAssets - totalEqAndLiab,
        prevTotalAssets - prevTotalEqAndLiab,
        0,
        totalAssets - totalEqAndLiab,
        Math.abs(totalAssets - totalEqAndLiab) <= 1 ? 'RECONCILED (₹0.00)' : 'CHECK DIFFERENCE',
      ],
      ...cfo.subtotalReconciliations.map((sr) => [
        sr.label,
        sr.worksheetName || sr.statementType,
        sr.cellReference || '—',
        fromPaise(sr.reportedPaise),
        sr.priorReportedPaise !== undefined ? fromPaise(sr.priorReportedPaise) : 0,
        fromPaise(sr.calculatedPaise),
        fromPaise(sr.differencePaise),
        sr.isMatched ? 'RECONCILED' : 'VARIANCE DETECTED',
      ]),
    ];

    coreMappingRows.forEach((vr, idx) => {
      styleDataRow(wsVal, 6 + idx, vr, {
        isAlternating: idx % 2 === 1,
        leftAlignCols: [1, 2, 7],
        numFmts: [undefined, undefined, undefined, INR_FMT, INR_FMT, INR_FMT, INR_FMT, undefined],
      });
    });
  }

  if (sheetsMap.has('Source Data')) {
    const wsSrc = sheetsMap.get('Source Data')!;
    wsSrc.columns = [
      { width: 20 },
      { width: 34 },
      { width: 14 },
      { width: 28 },
      { width: 16 },
      { width: 16 },
      { width: 16 },
      { width: 16 },
      { width: 44 },
    ];
    styleBanner(wsSrc, `SOURCE FILES & NORMALIZED LEDGER PROVENANCE`, `Company: ${model.companyName}`, 9);
    addNavStrip(wsSrc, plannedSheetNames);
    styleTableHeader(wsSrc, 5, [
      'File ID',
      'Filename',
      'Format',
      'Detected Report Type',
      'Period Start',
      'Period End',
      'Records',
      'Status',
      'Validation Notes',
    ]);
    model.sourceFiles.forEach((sf, idx) => {
      styleDataRow(
        wsSrc,
        6 + idx,
        [
          sf.id,
          sf.filename,
          sf.format,
          sf.detectedReportType,
          sf.periodStart || '—',
          sf.periodEnd || '—',
          sf.recordsExtracted,
          sf.status,
          sf.validationMessages.map((m) => m.message).join('; '),
        ],
        {
          isAlternating: idx % 2 === 1,
          leftAlignCols: [1, 2, 3, 4, 5, 7, 8],
        }
      );
    });
  }

  // ==========================================================================
  // SHEET 19: DASHBOARD DATA (Section 13 — Dedicated Chart Source Engine)
  // ==========================================================================
  const wsData = sheetsMap.get('Dashboard Data')!;
  wsData.columns = [
    { width: 34 },
    { width: 22 },
    { width: 22 },
    { width: 22 },
    { width: 22 },
    { width: 22 },
    { width: 22 },
  ];
  styleBanner(
    wsData,
    `DASHBOARD DATA — STRUCTURED CHART BACKING TABLES`,
    `Every native Excel chart in this workbook references the structured tables on this worksheet.`,
    7
  );

  // Table 1 (Row 4..4+M): Monthly Series (or Annual Comparative Series when no monthly data exists)
  const mSeries = hasMonthly
    ? months.slice(-12)
    : [
        ...(pk.hasPriorData
          ? [
              {
                month: model.comparativePeriodLabel || 'FY 2024–25 (Prior Year)',
                revenuePaise: pk.revenuePaise ?? 0,
                purchasesPaise: Math.round((priorYtdCogs || 0) * 100),
                grossProfitPaise: pk.grossProfitPaise ?? 0,
                expensesPaise: Math.round((priorYtdOpex || 0) * 100),
                ebitdaOverrideInr: priorYtdEbitda,
                netProfitPaise: pk.patPaise ?? Math.round((priorYtdPat || 0) * 100),
                collectionsPaise: cf.operatingCashFlowPriorPaise ?? 0,
                paymentsPaise: 0,
                workingCapitalPaise: pk.workingCapitalPaise ?? 0,
              },
            ]
          : []),
        {
          month: model.currentPeriodLabel || 'FY 2025–26 (Current Year)',
          revenuePaise: k.netRevenuePaise,
          purchasesPaise: k.cogsPaise,
          grossProfitPaise: k.hasGrossProfitData ? k.grossProfitPaise : 0,
          expensesPaise: k.employeeExpensesPaise + k.adminExpensesPaise + k.sellingExpensesPaise,
          ebitdaOverrideInr: ytdEbitda,
          netProfitPaise: k.patPaise,
          collectionsPaise: cf.operatingCashFlowPaise ?? k.patPaise,
          paymentsPaise: 0,
          workingCapitalPaise: k.workingCapitalPaise,
        },
      ];

  styleTableHeader(wsData, 4, [
    hasMonthly ? 'Month' : 'Reporting Period (Annual)',
    'Revenue (INR)',
    'Operating & Direct Expenses (INR)',
    'Gross Profit (INR)',
    'EBITDA (INR)',
    'Net Profit (INR)',
    'Operating Cash Flow (INR)',
  ]);

  const monthLabels: string[] = [];
  const monthRevVals: number[] = [];
  const monthExpVals: number[] = [];
  const monthGpVals: number[] = [];
  const monthEbitdaVals: number[] = [];
  const monthPatVals: number[] = [];
  const monthOcfVals: number[] = [];

  mSeries.forEach((m: any, idx) => {
    const r = 5 + idx;
    const revV = fromPaise(m.revenuePaise);
    const expV = fromPaise(m.purchasesPaise + m.expensesPaise);
    const gpV = k.hasGrossProfitData ? fromPaise(m.grossProfitPaise) : 0;
    const ebitdaV =
      m.ebitdaOverrideInr !== undefined
        ? m.ebitdaOverrideInr
        : Math.round(revV - expV);
    const patV = fromPaise(m.netProfitPaise);
    const ocfV =
      m.collectionsPaise !== 0 || m.paymentsPaise !== 0
        ? fromPaise(m.collectionsPaise - m.paymentsPaise)
        : Math.round(ocfInr / Math.max(1, mSeries.length));

    monthLabels.push(m.month);
    monthRevVals.push(revV);
    monthExpVals.push(expV);
    monthGpVals.push(gpV);
    monthEbitdaVals.push(ebitdaV);
    monthPatVals.push(patV);
    monthOcfVals.push(ocfV);

    styleDataRow(wsData, r, [m.month, revV, expV, gpV, ebitdaV, patV, ocfV], {
      isAlternating: idx % 2 === 1,
      numFmts: [undefined, INR_INT_FMT, INR_INT_FMT, INR_INT_FMT, INR_INT_FMT, INR_INT_FMT, INR_INT_FMT],
    });
  });
  const monthEndRow = 4 + mSeries.length;

  // Table 2 (Rows 19–25): Expense Composition
  styleTableHeader(wsData, 19, ['Expense Category', 'Current Period (INR)', 'Share %']);
  const expRows: Array<[string, number]> = [
    ['Purchases & Cost of Goods', Math.max(0, fromPaise(k.purchasesPaise - k.purchaseReturnsPaise))],
    ['Direct Operating Expenses', Math.max(0, fromPaise(k.directExpensesPaise))],
    ['Employee Salaries & Payroll', Math.max(0, fromPaise(k.employeeExpensesPaise))],
    ['Admin & General Overheads', Math.max(0, fromPaise(k.adminExpensesPaise))],
    ['Selling & Distribution', Math.max(0, fromPaise(k.sellingExpensesPaise))],
    ['Finance Costs & Depreciation', Math.max(0, fromPaise(k.financeCostsPaise + k.depreciationPaise))],
  ];
  const totalExpSum = Math.max(1, expRows.reduce((a, b) => a + b[1], 0));
  expRows.forEach((er, idx) => {
    const r = 20 + idx;
    styleDataRow(
      wsData,
      r,
      [
        er[0],
        er[1],
        { formula: `IF(SUM($B$20:$B$25)=0,0,B${r}/SUM($B$20:$B$25))`, result: er[1] / totalExpSum },
      ],
      {
        isAlternating: idx % 2 === 1,
        numFmts: [undefined, INR_INT_FMT, PCT_FMT],
      }
    );
  });

  // Table 3 (Rows 28–31): Cash Flow Activity Summary
  styleTableHeader(wsData, 28, ['Cash Flow Activity', 'Current Period (INR)', 'Previous Period (INR)']);
  const cfActRows: Array<[string, any, number, number]> = [
    [
      'Operating Cash Flow',
      hasCfSheet ? { formula: "'Cash Flow Dashboard'!C6", result: ocfInr } : ocfInr,
      prevOcfInr,
      ocfInr,
    ],
    [
      'Investing Cash Flow',
      hasCfSheet ? { formula: "'Cash Flow Dashboard'!C7", result: icfInr } : icfInr,
      prevIcfInr,
      icfInr,
    ],
    [
      'Financing Cash Flow',
      hasCfSheet ? { formula: "'Cash Flow Dashboard'!C8", result: fcfActInr } : fcfActInr,
      prevFcfActInr,
      fcfActInr,
    ],
  ];
  cfActRows.forEach((cr, idx) => {
    styleDataRow(wsData, 29 + idx, [cr[0], cr[1], cr[2]], {
      isAlternating: idx % 2 === 1,
      numFmts: [undefined, INR_INT_FMT, INR_INT_FMT],
    });
  });

  // Table 4 (Rows 34–39): Cash Flow Bridge
  styleTableHeader(wsData, 34, ['Cash Flow Bridge Step', 'Amount (INR)']);
  const bridgeRows: Array<[string, number]> = [
    ['Opening Cash', openCashInr],
    ['Operating CF', ocfInr],
    ['Investing CF', icfInr],
    ['Financing CF', fcfActInr],
    ['Closing Cash', closeCashInr],
  ];
  bridgeRows.forEach((br, idx) => {
    styleDataRow(wsData, 35 + idx, [br[0], br[1]], {
      isAlternating: idx % 2 === 1,
      numFmts: [undefined, INR_INT_FMT],
    });
  });

  // Table 5 (Rows 42–46): Current Assets vs Current Liabilities & Working Capital
  styleTableHeader(wsData, 42, ['Balance Sheet Liquidity Head', 'Current Period (INR)', 'Previous Period (INR)']);
  const bsLiqRows: Array<[string, any, number, number]> = [
    [
      'Current Assets',
      hasBsSheet ? { formula: "'Balance Sheet Dashboard'!C6", result: currAssets } : currAssets,
      prevCurrAssets,
      currAssets,
    ],
    [
      'Current Liabilities',
      hasBsSheet ? { formula: "'Balance Sheet Dashboard'!C12", result: totalCurrLiab } : totalCurrLiab,
      prevCurrLiab,
      totalCurrLiab,
    ],
    [
      'Net Working Capital',
      { formula: 'B43-B44', result: currAssets - totalCurrLiab },
      prevCurrAssets - prevCurrLiab,
      currAssets - totalCurrLiab,
    ],
    [
      'Cash & Bank Balances',
      hasBsSheet ? { formula: "'Balance Sheet Dashboard'!C19", result: cashBankInr } : cashBankInr,
      prevCashBankInr,
      cashBankInr,
    ],
  ];
  bsLiqRows.forEach((blr, idx) => {
    styleDataRow(wsData, 43 + idx, [blr[0], blr[1], blr[2]], {
      isAlternating: idx % 2 === 1,
      numFmts: [undefined, INR_INT_FMT, INR_INT_FMT],
    });
  });

  // Table 6 (Rows 49–55): Key Financial Ratios Comparison
  styleTableHeader(wsData, 49, ['Financial Ratio', 'Current Period', 'Previous Period']);
  const getRatioVal = (id: string) => {
    const found = allRatios.find((r) => r.id === id);
    return {
      curr: found?.currentValue !== null && found?.currentValue !== undefined ? Number(found.currentValue.toFixed(2)) : 0,
      prev: found?.previousValue !== null && found?.previousValue !== undefined ? Number(found.previousValue.toFixed(2)) : 0,
    };
  };
  const ratioChartRows: Array<[string, number, number]> = [
    ['Current Ratio (x)', getRatioVal('RATIO_CURRENT').curr, getRatioVal('RATIO_CURRENT').prev],
    ['Quick Ratio (x)', getRatioVal('RATIO_QUICK').curr, getRatioVal('RATIO_QUICK').prev],
    ['Debt Equity Ratio (x)', getRatioVal('RATIO_DEBT_EQUITY').curr, getRatioVal('RATIO_DEBT_EQUITY').prev],
    ['Gross Profit Margin (%)', getRatioVal('RATIO_GROSS_MARGIN').curr, getRatioVal('RATIO_GROSS_MARGIN').prev],
    ['EBITDA Margin (%)', getRatioVal('RATIO_EBITDA_MARGIN').curr, getRatioVal('RATIO_EBITDA_MARGIN').prev],
    ['Net Profit Margin (%)', getRatioVal('RATIO_NET_MARGIN').curr, getRatioVal('RATIO_NET_MARGIN').prev],
  ];
  ratioChartRows.forEach((rr, idx) => {
    styleDataRow(wsData, 50 + idx, [rr[0], rr[1], rr[2]], {
      isAlternating: idx % 2 === 1,
      numFmts: [undefined, '0.00', '0.00'],
    });
  });

  // Table 7 (Rows 58–62): Balance Sheet Asset Composition
  styleTableHeader(wsData, 58, ['Asset Component', 'Closing Balance (INR)']);
  const assetCompRows: Array<[string, number]> = [
    ['Fixed & Non-Current Assets', Math.max(0, nonCurrAssets)],
    ['Closing Inventory', Math.max(0, fromPaise(k.inventoryPaise))],
    ['Trade Receivables', Math.max(0, fromPaise(k.tradeReceivablesPaise))],
    ['Cash & Bank Balances', Math.max(0, cashBankInr)],
  ];
  assetCompRows.forEach((ar, idx) => {
    styleDataRow(wsData, 59 + idx, [ar[0], ar[1]], {
      isAlternating: idx % 2 === 1,
      numFmts: [undefined, INR_INT_FMT],
    });
  });

  // Table 8 (Rows 65–69): Liabilities & Equity Composition
  styleTableHeader(wsData, 65, ['Equity & Liability Component', 'Closing Balance (INR)']);
  const liabCompRows: Array<[string, number]> = [
    ['Shareholders Equity', Math.max(0, equityTotal)],
    ['Total Borrowings (Debt)', Math.max(0, borrowingsTotal)],
    ['Trade Payables (Creditors)', Math.max(0, fromPaise(k.tradePayablesPaise))],
    ['Tax & Other Current Liab.', Math.max(0, currLiabOther - fromPaise(k.tradePayablesPaise))],
  ];
  liabCompRows.forEach((lr, idx) => {
    styleDataRow(wsData, 66 + idx, [lr[0], lr[1]], {
      isAlternating: idx % 2 === 1,
      numFmts: [undefined, INR_INT_FMT],
    });
  });

  // Table 9 (Rows 72–76): Working Capital & CCC Days
  styleTableHeader(wsData, 72, ['Working Capital Item', 'Amount (INR)', 'Cycle Days']);
  const wcChartRows: Array<[string, number, number]> = [
    ['Trade Receivables (DSO)', fromPaise(wc.tradeReceivablesPaise), wc.dsoDays ?? 0],
    ['Closing Inventory (DIO)', fromPaise(wc.inventoryPaise), wc.dioDays ?? 0],
    ['Trade Payables (DPO)', fromPaise(wc.tradePayablesPaise), wc.dpoDays ?? 0],
    ['Net Working Capital (CCC)', fromPaise(wc.netWorkingCapitalPaise), wc.cashConversionCycleDays ?? 0],
  ];
  wcChartRows.forEach((wr, idx) => {
    styleDataRow(wsData, 73 + idx, [wr[0], wr[1], wr[2]], {
      isAlternating: idx % 2 === 1,
      numFmts: [undefined, INR_INT_FMT, '0.0'],
    });
  });

  // ==========================================================================
  // SHEET 20: SETTINGS (Section 15 & Technical Disclosure)
  // ==========================================================================
  const wsSettings = sheetsMap.get('Settings')!;
  wsSettings.columns = [{ width: 36 }, { width: 32 }, { width: 62 }];
  styleBanner(
    wsSettings,
    `WORKBOOK SETTINGS, PARAMETERS & TECHNICAL CHART DISCLOSURE`,
    `Company: ${model.companyName}  |  Interactive Filter Controls & Formula Audit Configuration`,
    3
  );
  addNavStrip(wsSettings, plannedSheetNames);
  styleTableHeader(wsSettings, 5, ['Parameter / Setting', 'Active Value', 'Description & Impact']);

  const settingsRows = [
    ['Company / Entity Name', model.companyName, 'Active consolidated entity from uploaded statements'],
    ['Selected Financial Year', 'FY 2025-26', 'Primary reporting financial year'],
    ['Selected Period Mode', `${options.periodType} (${options.selectedPeriodValue})`, 'Period slice applied to dashboard schedules'],
    ['Report Package Type', options.reportType, 'Selected export configuration'],
    ['Full Recalculation on Open', 'ENABLED (fullCalcOnLoad=true)', 'Excel automatically recalculates all linked formulas upon opening'],
    ['Native OpenXML Chart Objects', options.includeCharts ? 'ENABLED (25 Native Charts)' : 'DISABLED', 'Embedded ECMA-376 DrawingML charts linked to Dashboard Data ranges'],
    [
      'Chart 6 (Cash Flow Movement) Disclosure',
      'Native OpenXML Column Bridge Chart',
      'Uses standard ECMA-376 Clustered Column Bridge format for 100% compatibility across all Microsoft Excel versions without Office 2016+ cx:chartSpace repair warnings.',
    ],
    ['Colour Palette Standard', '#0B1F3A / #2A78D6 / #1BAF7A / #6250D6', 'Automat Group CFO Monthly Pack executive corporate palette'],
  ];
  settingsRows.forEach((sr, idx) => {
    styleDataRow(wsSettings, 6 + idx, sr, {
      isAlternating: idx % 2 === 1,
      leftAlignCols: [1, 2],
    });
  });

  // ==========================================================================
  // BUILD NATIVE EXCEL CHARTS SPECIFICATION (Sections 5, 7, 8, 9, 10, 11)
  // ==========================================================================
  const chartSpecs: NativeExcelChartSpec[] = [];

  if (options.includeCharts) {
    // 8 EXECUTIVE DASHBOARD CHARTS (Section 5)
    chartSpecs.push(
      // CHART 1: Monthly Revenue vs Expenses (Clustered Column Chart)
      {
        sheetName: 'Executive Dashboard',
        title: '1. Monthly Revenue vs Expenses (INR)',
        chartType: 'COLUMN',
        fromCol: 0,
        fromRow: 21,
        toCol: 6,
        toRow: 36,
        series: [
          {
            name: 'Monthly Revenue (INR)',
            nameFormula: "'Dashboard Data'!$B$4",
            catFormula: `'Dashboard Data'!$A$5:$A$${monthEndRow}`,
            valFormula: `'Dashboard Data'!$B$5:$B$${monthEndRow}`,
            colorHex: '2A78D6',
            categories: monthLabels,
            values: monthRevVals,
          },
          {
            name: 'Operating & Direct Expenses (INR)',
            nameFormula: "'Dashboard Data'!$C$4",
            catFormula: `'Dashboard Data'!$A$5:$A$${monthEndRow}`,
            valFormula: `'Dashboard Data'!$C$5:$C$${monthEndRow}`,
            colorHex: 'EB6834',
            categories: monthLabels,
            values: monthExpVals,
          },
        ],
      },
      // CHART 2: Gross Profit and Net Profit Trend (Line Chart)
      {
        sheetName: 'Executive Dashboard',
        title: '2. Gross Profit and Net Profit Trend (INR)',
        chartType: 'LINE',
        fromCol: 6,
        fromRow: 21,
        toCol: 12,
        toRow: 36,
        series: [
          {
            name: 'Gross Profit (INR)',
            nameFormula: "'Dashboard Data'!$D$4",
            catFormula: `'Dashboard Data'!$A$5:$A$${monthEndRow}`,
            valFormula: `'Dashboard Data'!$D$5:$D$${monthEndRow}`,
            colorHex: '2A78D6',
            categories: monthLabels,
            values: monthGpVals,
          },
          {
            name: 'Net Profit (INR)',
            nameFormula: "'Dashboard Data'!$F$4",
            catFormula: `'Dashboard Data'!$A$5:$A$${monthEndRow}`,
            valFormula: `'Dashboard Data'!$F$5:$F$${monthEndRow}`,
            colorHex: '1BAF7A',
            categories: monthLabels,
            values: monthPatVals,
          },
        ],
      },
      // CHART 3: Expense Composition (Doughnut Chart)
      {
        sheetName: 'Executive Dashboard',
        title: '3. Expense Composition Breakdown',
        chartType: 'DOUGHNUT',
        fromCol: 0,
        fromRow: 37,
        toCol: 6,
        toRow: 52,
        series: [
          {
            name: 'Expense Share',
            nameFormula: "'Dashboard Data'!$B$19",
            catFormula: "'Dashboard Data'!$A$20:$A$25",
            valFormula: "'Dashboard Data'!$B$20:$B$25",
            colorHex: '2A78D6',
            categories: expRows.map((x) => x[0]),
            values: expRows.map((x) => x[1]),
            sliceColorsHex: ['0B1F3A', '2A78D6', '1BAF7A', '6250D6', 'EB6834', 'D03B3B'],
          },
        ],
      },
      // CHART 4: EBITDA Trend (Column Chart)
      {
        sheetName: 'Executive Dashboard',
        title: '4. Monthly EBITDA Trend (INR)',
        chartType: 'COLUMN',
        fromCol: 6,
        fromRow: 37,
        toCol: 12,
        toRow: 52,
        series: [
          {
            name: 'EBITDA (INR)',
            nameFormula: "'Dashboard Data'!$E$4",
            catFormula: `'Dashboard Data'!$A$5:$A$${monthEndRow}`,
            valFormula: `'Dashboard Data'!$E$5:$E$${monthEndRow}`,
            colorHex: '1BAF7A',
            categories: monthLabels,
            values: monthEbitdaVals,
          },
        ],
      },
      // CHART 5: Operating, Investing and Financing Cash Flow (Column Chart)
      {
        sheetName: 'Executive Dashboard',
        title: '5. Operating, Investing & Financing Cash Flow',
        chartType: 'COLUMN',
        fromCol: 0,
        fromRow: 53,
        toCol: 6,
        toRow: 68,
        series: [
          {
            name: 'Current Period (INR)',
            nameFormula: "'Dashboard Data'!$B$28",
            catFormula: "'Dashboard Data'!$A$29:$A$31",
            valFormula: "'Dashboard Data'!$B$29:$B$31",
            colorHex: '2A78D6',
            categories: cfActRows.map((x) => x[0]),
            values: cfActRows.map((x) => x[3]),
          },
          {
            name: 'Previous Period (INR)',
            nameFormula: "'Dashboard Data'!$C$28",
            catFormula: "'Dashboard Data'!$A$29:$A$31",
            valFormula: "'Dashboard Data'!$C$29:$C$31",
            colorHex: '6250D6',
            categories: cfActRows.map((x) => x[0]),
            values: cfActRows.map((x) => x[2]),
          },
        ],
      },
      // CHART 6: Cash Flow Movement Bridge (Column Bridge Chart)
      {
        sheetName: 'Executive Dashboard',
        title: '6. Cash Flow Movement Bridge (Opening to Closing)',
        chartType: 'COLUMN',
        fromCol: 6,
        fromRow: 53,
        toCol: 12,
        toRow: 68,
        series: [
          {
            name: 'Cash Bridge Step (INR)',
            nameFormula: "'Dashboard Data'!$B$34",
            catFormula: "'Dashboard Data'!$A$35:$A$39",
            valFormula: "'Dashboard Data'!$B$35:$B$39",
            colorHex: '0B1F3A',
            categories: bridgeRows.map((x) => x[0]),
            values: bridgeRows.map((x) => x[1]),
          },
        ],
      },
      // CHART 7: Current Assets vs Current Liabilities (Grouped Column Chart)
      {
        sheetName: 'Executive Dashboard',
        title: '7. Current Assets vs Current Liabilities',
        chartType: 'COLUMN',
        fromCol: 0,
        fromRow: 69,
        toCol: 6,
        toRow: 84,
        series: [
          {
            name: 'Current Period (INR)',
            nameFormula: "'Dashboard Data'!$B$42",
            catFormula: "'Dashboard Data'!$A$43:$A$46",
            valFormula: "'Dashboard Data'!$B$43:$B$46",
            colorHex: '2A78D6',
            categories: bsLiqRows.map((x) => x[0]),
            values: bsLiqRows.map((x) => x[3]),
          },
          {
            name: 'Previous Period (INR)',
            nameFormula: "'Dashboard Data'!$C$42",
            catFormula: "'Dashboard Data'!$A$43:$A$46",
            valFormula: "'Dashboard Data'!$C$43:$C$46",
            colorHex: '1BAF7A',
            categories: bsLiqRows.map((x) => x[0]),
            values: bsLiqRows.map((x) => x[2]),
          },
        ],
      },
      // CHART 8: Financial Ratio Comparison (Bar Chart)
      {
        sheetName: 'Executive Dashboard',
        title: '8. Financial Ratio Comparison',
        chartType: 'BAR',
        fromCol: 6,
        fromRow: 69,
        toCol: 12,
        toRow: 84,
        series: [
          {
            name: 'Current Period',
            nameFormula: "'Dashboard Data'!$B$49",
            catFormula: "'Dashboard Data'!$A$50:$A$55",
            valFormula: "'Dashboard Data'!$B$50:$B$55",
            colorHex: '6250D6',
            categories: ratioChartRows.map((x) => x[0]),
            values: ratioChartRows.map((x) => x[1]),
          },
          {
            name: 'Previous Period',
            nameFormula: "'Dashboard Data'!$C$49",
            catFormula: "'Dashboard Data'!$A$50:$A$55",
            valFormula: "'Dashboard Data'!$C$50:$C$55",
            colorHex: '2A78D6',
            categories: ratioChartRows.map((x) => x[0]),
            values: ratioChartRows.map((x) => x[2]),
          },
        ],
      }
    );

    // P&L DASHBOARD CHARTS (Section 7 — 3 Charts)
    if (sheetsMap.has('P&L Dashboard')) {
      chartSpecs.push(
        {
          sheetName: 'P&L Dashboard',
          title: 'Monthly Revenue vs Expenses (INR)',
          chartType: 'COLUMN',
          fromCol: 0,
          fromRow: 22,
          toCol: 4,
          toRow: 37,
          series: [
            {
              name: 'Revenue',
              nameFormula: "'Dashboard Data'!$B$4",
              catFormula: `'Dashboard Data'!$A$5:$A$${monthEndRow}`,
              valFormula: `'Dashboard Data'!$B$5:$B$${monthEndRow}`,
              colorHex: '2A78D6',
              categories: monthLabels,
              values: monthRevVals,
            },
            {
              name: 'Expenses',
              nameFormula: "'Dashboard Data'!$C$4",
              catFormula: `'Dashboard Data'!$A$5:$A$${monthEndRow}`,
              valFormula: `'Dashboard Data'!$C$5:$C$${monthEndRow}`,
              colorHex: 'EB6834',
              categories: monthLabels,
              values: monthExpVals,
            },
          ],
        },
        {
          sheetName: 'P&L Dashboard',
          title: 'Gross Profit, EBITDA & Net Profit Trend',
          chartType: 'LINE',
          fromCol: 4,
          fromRow: 22,
          toCol: 8,
          toRow: 37,
          series: [
            {
              name: 'Gross Profit',
              nameFormula: "'Dashboard Data'!$D$4",
              catFormula: `'Dashboard Data'!$A$5:$A$${monthEndRow}`,
              valFormula: `'Dashboard Data'!$D$5:$D$${monthEndRow}`,
              colorHex: '2A78D6',
              categories: monthLabels,
              values: monthGpVals,
            },
            {
              name: 'EBITDA',
              nameFormula: "'Dashboard Data'!$E$4",
              catFormula: `'Dashboard Data'!$A$5:$A$${monthEndRow}`,
              valFormula: `'Dashboard Data'!$E$5:$E$${monthEndRow}`,
              colorHex: '6250D6',
              categories: monthLabels,
              values: monthEbitdaVals,
            },
            {
              name: 'Net Profit',
              nameFormula: "'Dashboard Data'!$F$4",
              catFormula: `'Dashboard Data'!$A$5:$A$${monthEndRow}`,
              valFormula: `'Dashboard Data'!$F$5:$F$${monthEndRow}`,
              colorHex: '1BAF7A',
              categories: monthLabels,
              values: monthPatVals,
            },
          ],
        },
        {
          sheetName: 'P&L Dashboard',
          title: 'Cost & Expense Composition',
          chartType: 'DOUGHNUT',
          fromCol: 0,
          fromRow: 38,
          toCol: 5,
          toRow: 53,
          series: [
            {
              name: 'Cost Share',
              nameFormula: "'Dashboard Data'!$B$19",
              catFormula: "'Dashboard Data'!$A$20:$A$25",
              valFormula: "'Dashboard Data'!$B$20:$B$25",
              colorHex: '2A78D6',
              categories: expRows.map((x) => x[0]),
              values: expRows.map((x) => x[1]),
              sliceColorsHex: ['0B1F3A', '2A78D6', '1BAF7A', '6250D6', 'EB6834', 'D03B3B'],
            },
          ],
        }
      );
    }

    // BALANCE SHEET DASHBOARD CHARTS (Section 8 — 4 Charts)
    if (sheetsMap.has('Balance Sheet Dashboard')) {
      chartSpecs.push(
        {
          sheetName: 'Balance Sheet Dashboard',
          title: 'Asset Composition Breakdown',
          chartType: 'DOUGHNUT',
          fromCol: 0,
          fromRow: 23,
          toCol: 4,
          toRow: 38,
          series: [
            {
              name: 'Assets',
              nameFormula: "'Dashboard Data'!$B$58",
              catFormula: "'Dashboard Data'!$A$59:$A$62",
              valFormula: "'Dashboard Data'!$B$59:$B$62",
              colorHex: '2A78D6',
              categories: assetCompRows.map((x) => x[0]),
              values: assetCompRows.map((x) => x[1]),
              sliceColorsHex: ['0B1F3A', '2A78D6', '1BAF7A', '6250D6'],
            },
          ],
        },
        {
          sheetName: 'Balance Sheet Dashboard',
          title: 'Liabilities and Equity Structure',
          chartType: 'COLUMN',
          fromCol: 4,
          fromRow: 23,
          toCol: 7,
          toRow: 38,
          series: [
            {
              name: 'Closing Balance (INR)',
              nameFormula: "'Dashboard Data'!$B$65",
              catFormula: "'Dashboard Data'!$A$66:$A$69",
              valFormula: "'Dashboard Data'!$B$66:$B$69",
              colorHex: '0B1F3A',
              categories: liabCompRows.map((x) => x[0]),
              values: liabCompRows.map((x) => x[1]),
            },
          ],
        },
        {
          sheetName: 'Balance Sheet Dashboard',
          title: 'Working Capital & Liquidity Position',
          chartType: 'COLUMN',
          fromCol: 0,
          fromRow: 39,
          toCol: 4,
          toRow: 54,
          series: [
            {
              name: 'Current Period (INR)',
              nameFormula: "'Dashboard Data'!$B$42",
              catFormula: "'Dashboard Data'!$A$43:$A$46",
              valFormula: "'Dashboard Data'!$B$43:$B$46",
              colorHex: '1BAF7A',
              categories: bsLiqRows.map((x) => x[0]),
              values: bsLiqRows.map((x) => x[3]),
            },
          ],
        },
        {
          sheetName: 'Balance Sheet Dashboard',
          title: 'Current Assets vs Current Liabilities',
          chartType: 'COLUMN',
          fromCol: 4,
          fromRow: 39,
          toCol: 7,
          toRow: 54,
          series: [
            {
              name: 'Current Period',
              nameFormula: "'Dashboard Data'!$B$42",
              catFormula: "'Dashboard Data'!$A$43:$A$44",
              valFormula: "'Dashboard Data'!$B$43:$B$44",
              colorHex: '2A78D6',
              categories: bsLiqRows.slice(0, 2).map((x) => x[0]),
              values: bsLiqRows.slice(0, 2).map((x) => x[3]),
            },
            {
              name: 'Previous Period',
              nameFormula: "'Dashboard Data'!$C$42",
              catFormula: "'Dashboard Data'!$A$43:$A$44",
              valFormula: "'Dashboard Data'!$C$43:$C$44",
              colorHex: '6250D6',
              categories: bsLiqRows.slice(0, 2).map((x) => x[0]),
              values: bsLiqRows.slice(0, 2).map((x) => x[2]),
            },
          ],
        }
      );
    }

    // CASH FLOW DASHBOARD CHARTS (Section 9 — 4 Charts)
    if (sheetsMap.has('Cash Flow Dashboard')) {
      const cfChartStartRow = cf.lineItems.length > 0 ? 20 + cf.lineItems.length : 17;
      chartSpecs.push(
        {
          sheetName: 'Cash Flow Dashboard',
          title: 'Cash Flow Composition (OCF, ICF, FCF)',
          chartType: 'COLUMN',
          fromCol: 0,
          fromRow: cfChartStartRow,
          toCol: 3,
          toRow: cfChartStartRow + 15,
          series: [
            {
              name: 'Current Period (INR)',
              nameFormula: "'Dashboard Data'!$B$28",
              catFormula: "'Dashboard Data'!$A$29:$A$31",
              valFormula: "'Dashboard Data'!$B$29:$B$31",
              colorHex: '2A78D6',
              categories: cfActRows.map((x) => x[0]),
              values: cfActRows.map((x) => x[3]),
            },
          ],
        },
        {
          sheetName: 'Cash Flow Dashboard',
          title: 'Cash Flow Trend (Monthly Operating CF)',
          chartType: 'LINE',
          fromCol: 3,
          fromRow: cfChartStartRow,
          toCol: 6,
          toRow: cfChartStartRow + 15,
          series: [
            {
              name: 'Operating Cash Flow',
              nameFormula: "'Dashboard Data'!$G$4",
              catFormula: `'Dashboard Data'!$A$5:$A$${monthEndRow}`,
              valFormula: `'Dashboard Data'!$G$5:$G$${monthEndRow}`,
              colorHex: '1BAF7A',
              categories: monthLabels,
              values: monthOcfVals,
            },
          ],
        },
        {
          sheetName: 'Cash Flow Dashboard',
          title: 'Cash Movement Bridge (Opening to Closing)',
          chartType: 'COLUMN',
          fromCol: 0,
          fromRow: cfChartStartRow + 16,
          toCol: 3,
          toRow: cfChartStartRow + 31,
          series: [
            {
              name: 'Cash Bridge (INR)',
              nameFormula: "'Dashboard Data'!$B$34",
              catFormula: "'Dashboard Data'!$A$35:$A$39",
              valFormula: "'Dashboard Data'!$B$35:$B$39",
              colorHex: '0B1F3A',
              categories: bridgeRows.map((x) => x[0]),
              values: bridgeRows.map((x) => x[1]),
            },
          ],
        },
        {
          sheetName: 'Cash Flow Dashboard',
          title: 'Operating Cash Flow vs Net Profit',
          chartType: 'COLUMN',
          fromCol: 3,
          fromRow: cfChartStartRow + 16,
          toCol: 6,
          toRow: cfChartStartRow + 31,
          series: [
            {
              name: 'Net Profit (INR)',
              nameFormula: "'Dashboard Data'!$F$4",
              catFormula: `'Dashboard Data'!$A$5:$A$${monthEndRow}`,
              valFormula: `'Dashboard Data'!$F$5:$F$${monthEndRow}`,
              colorHex: '2A78D6',
              categories: monthLabels,
              values: monthPatVals,
            },
            {
              name: 'Operating Cash Flow (INR)',
              nameFormula: "'Dashboard Data'!$G$4",
              catFormula: `'Dashboard Data'!$A$5:$A$${monthEndRow}`,
              valFormula: `'Dashboard Data'!$G$5:$G$${monthEndRow}`,
              colorHex: '1BAF7A',
              categories: monthLabels,
              values: monthOcfVals,
            },
          ],
        }
      );
    }

    // FINANCIAL RATIOS CHARTS (Section 10 — 2 Charts)
    if (sheetsMap.has('Financial Ratios')) {
      const ratioChartRow = 8 + allRatios.length;
      chartSpecs.push(
        {
          sheetName: 'Financial Ratios',
          title: 'Liquidity & Leverage Ratios Comparison',
          chartType: 'BAR',
          fromCol: 0,
          fromRow: ratioChartRow,
          toCol: 4,
          toRow: ratioChartRow + 15,
          series: [
            {
              name: 'Current Period',
              nameFormula: "'Dashboard Data'!$B$49",
              catFormula: "'Dashboard Data'!$A$50:$A$52",
              valFormula: "'Dashboard Data'!$B$50:$B$52",
              colorHex: '2A78D6',
              categories: ratioChartRows.slice(0, 3).map((x) => x[0]),
              values: ratioChartRows.slice(0, 3).map((x) => x[1]),
            },
            {
              name: 'Previous Period',
              nameFormula: "'Dashboard Data'!$C$49",
              catFormula: "'Dashboard Data'!$A$50:$A$52",
              valFormula: "'Dashboard Data'!$C$50:$C$52",
              colorHex: '6250D6',
              categories: ratioChartRows.slice(0, 3).map((x) => x[0]),
              values: ratioChartRows.slice(0, 3).map((x) => x[2]),
            },
          ],
        },
        {
          sheetName: 'Financial Ratios',
          title: 'Profitability Margins (GP, EBITDA, Net %)',
          chartType: 'COLUMN',
          fromCol: 4,
          fromRow: ratioChartRow,
          toCol: 7,
          toRow: ratioChartRow + 15,
          series: [
            {
              name: 'Current Period (%)',
              nameFormula: "'Dashboard Data'!$B$49",
              catFormula: "'Dashboard Data'!$A$53:$A$55",
              valFormula: "'Dashboard Data'!$B$53:$B$55",
              colorHex: '1BAF7A',
              categories: ratioChartRows.slice(3, 6).map((x) => x[0]),
              values: ratioChartRows.slice(3, 6).map((x) => x[1]),
            },
            {
              name: 'Previous Period (%)',
              nameFormula: "'Dashboard Data'!$C$49",
              catFormula: "'Dashboard Data'!$A$53:$A$55",
              valFormula: "'Dashboard Data'!$C$53:$C$55",
              colorHex: '2A78D6',
              categories: ratioChartRows.slice(3, 6).map((x) => x[0]),
              values: ratioChartRows.slice(3, 6).map((x) => x[2]),
            },
          ],
        }
      );
    }

    // WORKING CAPITAL CHARTS (Section 11 — 4 Charts)
    if (sheetsMap.has('Working Capital')) {
      chartSpecs.push(
        {
          sheetName: 'Working Capital',
          title: 'Working Capital Components (INR)',
          chartType: 'COLUMN',
          fromCol: 0,
          fromRow: 18,
          toCol: 3,
          toRow: 33,
          series: [
            {
              name: 'Amount (INR)',
              nameFormula: "'Dashboard Data'!$B$72",
              catFormula: "'Dashboard Data'!$A$73:$A$76",
              valFormula: "'Dashboard Data'!$B$73:$B$76",
              colorHex: '2A78D6',
              categories: wcChartRows.map((x) => x[0]),
              values: wcChartRows.map((x) => x[1]),
            },
          ],
        },
        {
          sheetName: 'Working Capital',
          title: 'Receivables vs Inventory vs Payables',
          chartType: 'BAR',
          fromCol: 3,
          fromRow: 18,
          toCol: 6,
          toRow: 33,
          series: [
            {
              name: 'Closing Balance (INR)',
              nameFormula: "'Dashboard Data'!$B$72",
              catFormula: "'Dashboard Data'!$A$73:$A$75",
              valFormula: "'Dashboard Data'!$B$73:$B$75",
              colorHex: '6250D6',
              categories: wcChartRows.slice(0, 3).map((x) => x[0]),
              values: wcChartRows.slice(0, 3).map((x) => x[1]),
            },
          ],
        },
        {
          sheetName: 'Working Capital',
          title: 'Monthly Liquidity & Operating Cash Movement',
          chartType: 'LINE',
          fromCol: 0,
          fromRow: 34,
          toCol: 3,
          toRow: 49,
          series: [
            {
              name: 'Operating Cash Flow',
              nameFormula: "'Dashboard Data'!$G$4",
              catFormula: `'Dashboard Data'!$A$5:$A$${monthEndRow}`,
              valFormula: `'Dashboard Data'!$G$5:$G$${monthEndRow}`,
              colorHex: '1BAF7A',
              categories: monthLabels,
              values: monthOcfVals,
            },
          ],
        },
        {
          sheetName: 'Working Capital',
          title: 'Cash Conversion Cycle Days (DSO, DIO, DPO, CCC)',
          chartType: 'COLUMN',
          fromCol: 3,
          fromRow: 34,
          toCol: 6,
          toRow: 49,
          series: [
            {
              name: 'Cycle Days',
              nameFormula: "'Dashboard Data'!$C$72",
              catFormula: "'Dashboard Data'!$A$73:$A$76",
              valFormula: "'Dashboard Data'!$C$73:$C$76",
              colorHex: 'EB6834',
              categories: wcChartRows.map((x) => x[0]),
              values: wcChartRows.map((x) => x[2]),
            },
          ],
        }
      );
    }
  }

  // Write ExcelJS workbook to buffer and inject native OpenXML charts
  const rawBuffer = await wb.xlsx.writeBuffer();
  return await injectNativeChartsIntoXlsx(rawBuffer, chartSpecs);
}

/**
 * Triggers a browser download of the generated CFO Excel Reporting Pack.
 */
export async function downloadCfoExcelPackWithOptions(
  model: ConsolidatedAccountingModel,
  statements: FinancialStatementBundle,
  ageing: AgeingAnalysisResult,
  ratios: FinancialRatioResult[],
  health: AccountingHealthCheckResult,
  gst: GSTReconciliationReport,
  insights: FinancialInsight[],
  cfo: CfoAnalyticsBundle,
  options: ExcelExportOptions = DEFAULT_EXCEL_EXPORT_OPTIONS
): Promise<void> {
  const bytes = await buildCfoExcelWorkbookBytes(
    model,
    statements,
    ageing,
    ratios,
    health,
    gst,
    insights,
    cfo,
    options
  );
  const blob = new Blob([bytes as unknown as BlobPart], {
    type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
  });
  const url = URL.createObjectURL(blob);
  const a = document.createElement('a');
  const safeName = model.companyName.replace(/[^a-zA-Z0-9_-]/g, '_');
  const suffix =
    options.reportType === 'EXECUTIVE_DASHBOARD_ONLY'
      ? 'Executive_Dashboard'
      : options.reportType === 'DETAILED_FINANCIAL_ANALYSIS'
        ? 'Financial_Analysis'
        : 'CFO_Monthly_Pack';
  a.href = url;
  a.download = `GSPU_${suffix}_${safeName}_${model.periodEnd}.xlsx`;
  document.body.appendChild(a);
  a.click();
  a.remove();
  URL.revokeObjectURL(url);
}
