import {
  AccountClassification,
  CashFlowStatementItem,
  computeSafeVariance,
  ConsolidatedAccountingModel,
  FinancialRatioResult,
  formatINR,
  formatINRCompact,
  fromPaise,
  PolicySettings,
} from '../types/accounting';
import { FinancialStatementBundle } from './accountingEngine';

export interface CfoKpiCardItem {
  id: string;
  label: string;
  unit: 'CURRENCY' | 'PERCENT' | 'RATIO';
  currentPaiseOrVal: number | null;
  previousPaiseOrVal: number | null;
  budgetPaiseOrVal: number | null;
  currentFormatted: string;
  previousFormatted: string;
  budgetFormatted: string;
  varianceAmountFormatted: string;
  variancePercentFormatted: string;
  statusIndicator:
    | 'Reconciled'
    | 'Favorable'
    | 'Attention Required'
    | 'Requires Review'
    | 'Requires Mapping'
    | 'Insufficient Data';
  formula: string;
  classifications?: AccountClassification[];
  sourceCellRef?: string;
  priorCellRef?: string;
}

export interface CashFlowAnalysisResult {
  hasUploadedCashFlowStatement: boolean;
  hasDerivableCashMovement: boolean;
  operatingCashFlowPaise: number | null;
  operatingCashFlowPriorPaise: number | null;
  investingCashFlowPaise: number | null;
  investingCashFlowPriorPaise: number | null;
  financingCashFlowPaise: number | null;
  financingCashFlowPriorPaise: number | null;
  netCashMovementPaise: number | null;
  openingCashPaise: number | null;
  closingCashPaise: number | null;
  capexPaise: number | null;
  freeCashFlowPaise: number | null;
  freeCashFlowDefinition: string;
  reconciliation: {
    computedClosingCashPaise: number | null;
    reportedClosingCashPaise: number | null;
    internalCashFlowDiffPaise: number;
    balanceSheetCashAndBankPaise: number;
    bsVsCfDiffPaise: number | null;
    isReconciledWithBs: boolean;
    statusMessage: string;
  };
  waterfallSteps: Array<{
    name: string;
    valueRupees: number;
    valuePaise: number;
    type: 'START' | 'INFLOW' | 'OUTFLOW' | 'END';
  }>;
  monthlyCashFlowTrends: Array<{
    month: string;
    OperatingCF: number;
    InvestingCF: number;
    FinancingCF: number;
    NetCashMovement: number;
    NetProfit: number;
  }>;
  lineItems: CashFlowStatementItem[];
}

export interface WorkingCapitalAnalysisResult {
  currentAssetsPaise: number;
  currentLiabilitiesPaise: number;
  inventoryPaise: number;
  tradeReceivablesPaise: number;
  tradePayablesPaise: number;
  cashAndBankPaise: number;
  netWorkingCapitalPaise: number;
  priorNetWorkingCapitalPaise: number | null;
  dsoDays: number | null;
  dpoDays: number | null;
  dioDays: number | null;
  cashConversionCycleDays: number | null;
  cccExplanation: string;
  monthlyWorkingCapitalTrend: Array<{
    month: string;
    CurrentAssets: number;
    CurrentLiabilities: number;
    NetWorkingCapital: number;
    Receivables: number;
    Payables: number;
    Inventory: number;
    CCC: number;
  }>;
}

export interface CfoManagementSuggestion {
  id: string;
  category:
    | 'Revenue Performance'
    | 'Gross Margin'
    | 'Operating Expenses'
    | 'Liquidity'
    | 'Operating Cash Flow'
    | 'Working Capital'
    | 'Debt & Leverage'
    | 'Interest Coverage'
    | 'Profit vs Cash Flow'
    | 'Balance Sheet Reconciliation'
    | 'Statement Movement';
  factualObservation: string;
  actualFinancialValue: string;
  comparativeValue: string;
  calculatedVariance: string;
  quantifiedFinancialImpact: string;
  riskClassification: 'HIGH RISK' | 'MODERATE RISK' | 'WATCH' | 'FAVORABLE';
  suggestedManagementAction: string;
  supportingSourceRef: string;
}

export interface ScenarioAssumptions {
  revenueGrowthPercent: number; // e.g. +8%
  grossProfitMarginPercent: number; // e.g. 34%
  operatingExpenseChangePercent: number; // e.g. +4%
  collectionPeriodDays: number; // e.g. 42
  paymentPeriodDays: number; // e.g. 35
  plannedCapexRupees: number; // e.g. 250000
  plannedLoanRepaymentRupees: number; // e.g. 150000
  plannedTaxPaymentRupees: number; // e.g. 120000
}

export interface ScenarioCaseOutput {
  caseName: 'Base Case' | 'Optimistic Case' | 'Conservative Case';
  projectedRevenuePaise: number;
  projectedGrossProfitPaise: number;
  projectedOpexPaise: number;
  projectedEbitdaPaise: number;
  projectedNetProfitPaise: number;
  projectedOperatingCashFlowPaise: number;
  projectedFreeCashFlowPaise: number;
  projectedClosingCashPaise: number;
  projectedWorkingCapitalPaise: number;
}

export interface SubtotalReconciliationRow {
  statementType: 'PNL' | 'BALANCE_SHEET' | 'CASH_FLOW';
  label: string;
  reportedPaise: number;
  priorReportedPaise?: number;
  calculatedPaise: number;
  differencePaise: number;
  isMatched: boolean;
  worksheetName?: string;
  cellReference?: string;
  priorCellReference?: string;
}

export interface BalanceSheetVerificationRow {
  id: string;
  component: string;
  category: 'EQUITY_AND_LIABILITIES' | 'ASSETS' | 'WORKING_CAPITAL' | 'RECONCILIATION';
  reportedSourcePaise: number | null;
  calculatedPaise: number;
  priorPeriodPaise: number | null;
  differencePaise: number;
  cellReference?: string;
  priorCellReference?: string;
  status: 'RECONCILED' | 'REQUIRES REVIEW' | 'INSUFFICIENT DATA';
  classifications?: AccountClassification[];
  formula: string;
}

export interface CfoAnalyticsBundle {
  cfoKpis: CfoKpiCardItem[];
  cashFlowAnalysis: CashFlowAnalysisResult;
  workingCapital: WorkingCapitalAnalysisResult;
  cashFlowRatios: FinancialRatioResult[];
  cfoSuggestions: CfoManagementSuggestion[];
  subtotalReconciliations: SubtotalReconciliationRow[];
  balanceSheetVerificationRows: BalanceSheetVerificationRow[];
  balanceSheetReconciliation: {
    totalAssetsPaise: number;
    totalEquityPaise: number;
    totalLiabilitiesPaise: number;
    totalEquityAndLiabilitiesPaise: number;
    differencePaise: number;
    isReconciled: boolean;
    priorTotalAssetsPaise?: number | null;
    priorTotalEquityAndLiabilitiesPaise?: number | null;
    priorDifferencePaise?: number | null;
    isPriorReconciled?: boolean;
    reportedTotalAssetsPaise?: number | null;
    reportedTotalEquityAndLiabilitiesPaise?: number | null;
  };
  priorPeriodTotals: {
    hasPriorData: boolean;
    revenuePaise: number | null;
    grossProfitPaise: number | null;
    ebitdaPaise: number | null;
    patPaise: number | null;
    totalAssetsPaise: number | null;
    currentAssetsPaise: number | null;
    currentLiabilitiesPaise: number | null;
    cashAndBankPaise: number | null;
    workingCapitalPaise: number | null;
  };
  budgetTotals: {
    hasBudgetData: boolean;
    revenuePaise: number | null;
    grossProfitPaise: number | null;
    ebitdaPaise: number | null;
    patPaise: number | null;
  };
}

export function computeCfoAnalyticsBundle(
  model: ConsolidatedAccountingModel,
  statements: FinancialStatementBundle,
  priorMonthStatements: FinancialStatementBundle | undefined,
  existingRatios: FinancialRatioResult[],
  policy: PolicySettings
): CfoAnalyticsBundle {
  const k = statements.kpis;
  const pk = priorMonthStatements?.kpis;

  // 1. Extract Prior Period & Budget Totals from computed StatementLineItems (which preserve exact signs)
  const getLinePrev = (code: string): number | null => {
    const allLines = [
      ...statements.pnlLines,
      ...statements.bsAssetLines,
      ...statements.bsLiabilityEquityLines,
    ];
    const found = allLines.find((l) => l.code === code);
    return found?.previousAmountPaise !== undefined ? found.previousAmountPaise : null;
  };
  const getLineCellRef = (code: string): string | undefined => {
    const allLines = [
      ...statements.pnlLines,
      ...statements.bsAssetLines,
      ...statements.bsLiabilityEquityLines,
    ];
    const found = allLines.find((l) => l.code === code);
    return found?.cellReferences?.[0];
  };
  const getLinePriorCellRef = (code: string): string | undefined => {
    const allLines = [
      ...statements.pnlLines,
      ...statements.bsAssetLines,
      ...statements.bsLiabilityEquityLines,
    ];
    const found = allLines.find((l) => l.code === code);
    return found?.priorCellReferences?.[0];
  };

  let sumBudgetSales = 0;
  let sumBudgetPurchases = 0;
  let sumBudgetDirectExp = 0;
  let sumBudgetOpex = 0;
  let hasStatementBudget = false;
  let hasStatementPrior = false;

  for (const [, summary] of model.ledgerSummaries.entries()) {
    const l = summary.ledger;
    const cls = l.classification;
    if (l.previousPeriodPaise !== undefined || l.previousPeriodSignedPaise !== undefined) {
      hasStatementPrior = true;
    }
    if (l.budgetPaise !== undefined && l.budgetPaise !== 0) {
      hasStatementBudget = true;
      const b = Math.abs(l.budgetPaise);
      if (cls === 'Sales' || cls === 'OtherOperatingIncome') sumBudgetSales += b;
      else if (cls === 'Purchases') sumBudgetPurchases += b;
      else if (cls === 'DirectExpenses') sumBudgetDirectExp += b;
      else if (
        cls === 'EmployeeExpenses' ||
        cls === 'AdminExpenses' ||
        cls === 'SellingExpenses'
      ) {
        sumBudgetOpex += b;
      }
    }
  }

  const hasPriorData = hasStatementPrior || pk !== undefined;
  const priorRevenuePaise = getLinePrev('REVENUE_NET') ?? (pk ? pk.netRevenuePaise : null);
  const priorGrossProfitPaise = k.hasGrossProfitData
    ? getLinePrev('GROSS_PROFIT') ?? (pk ? pk.grossProfitPaise : null)
    : null;
  const priorEbitdaPaise = getLinePrev('EBITDA') ?? (pk ? pk.ebitdaPaise : null);
  const priorPatPaise = getLinePrev('PAT') ?? (pk ? pk.patPaise : null);
  const priorCurrentAssetsPaise =
    getLinePrev('TOTAL_CURRENT_ASSETS') ?? (pk ? pk.totalCurrentAssetsPaise : null);
  const priorCurrentLiabilitiesPaise =
    getLinePrev('TOTAL_CURRENT_LIABILITIES') ?? (pk ? pk.totalCurrentLiabilitiesPaise : null);
  const priorTotalAssetsPaise =
    getLinePrev('TOTAL_ASSETS') ?? (pk ? pk.totalAssetsPaise : null);
  const priorCashAndBankPaise =
    getLinePrev('CASH_AND_BANK') ?? (pk ? pk.cashPaise + pk.bankPaise : null);
  const priorWorkingCapitalPaise =
    priorCurrentAssetsPaise !== null && priorCurrentLiabilitiesPaise !== null
      ? priorCurrentAssetsPaise - priorCurrentLiabilitiesPaise
      : null;

  const budgetGrossProfitPaise =
    hasStatementBudget && k.hasGrossProfitData
      ? sumBudgetSales - sumBudgetPurchases - sumBudgetDirectExp
      : null;
  const budgetEbitdaPaise =
    hasStatementBudget
      ? sumBudgetSales - sumBudgetPurchases - sumBudgetDirectExp - sumBudgetOpex
      : null;

  // 2. Cash Flow Analysis & Reconciliation (Part 7)
  const cfItems = model.cashFlowItems || [];
  const hasUploadedCf = cfItems.length > 0;

  let cfoTotal: number | null = null;
  let cfoPriorTotal: number | null = null;
  let cfiTotal: number | null = null;
  let cfiPriorTotal: number | null = null;
  let cffTotal: number | null = null;
  let cffPriorTotal: number | null = null;
  let cfNetMove: number | null = null;
  let cfOpening: number | null = null;
  let cfClosing: number | null = null;
  let capexPaise: number | null = null;

  if (hasUploadedCf) {
    let sumOperatingItems = 0;
    let sumInvestingItems = 0;
    let sumFinancingItems = 0;
    let sumCapexItems = 0;
    let hasCapex = false;

    for (const item of cfItems) {
      if (item.category === 'CFO_OPERATING_TOTAL') {
        cfoTotal = item.currentPeriodPaise;
        if (item.previousPeriodPaise !== undefined) cfoPriorTotal = item.previousPeriodPaise;
      } else if (item.category === 'CFI_INVESTING_TOTAL') {
        cfiTotal = item.currentPeriodPaise;
        if (item.previousPeriodPaise !== undefined) cfiPriorTotal = item.previousPeriodPaise;
      } else if (item.category === 'CFF_FINANCING_TOTAL') {
        cffTotal = item.currentPeriodPaise;
        if (item.previousPeriodPaise !== undefined) cffPriorTotal = item.previousPeriodPaise;
      } else if (item.category === 'CF_NET_CHANGE') {
        cfNetMove = item.currentPeriodPaise;
      } else if (item.category === 'CF_OPENING_CASH') {
        cfOpening = item.currentPeriodPaise;
      } else if (item.category === 'CF_CLOSING_CASH') {
        cfClosing = item.currentPeriodPaise;
      } else if (!item.isSubtotalOrTotal) {
        if (
          item.category === 'CFO_OPERATING_ITEM' ||
          item.category === 'CFO_WORKING_CAPITAL_CHANGE' ||
          item.category === 'CFO_TAXES_PAID'
        ) {
          sumOperatingItems += item.currentPeriodPaise;
        } else if (
          item.category === 'CFI_CAPEX' ||
          item.category === 'CFI_INVESTING_ITEM'
        ) {
          sumInvestingItems += item.currentPeriodPaise;
          if (item.category === 'CFI_CAPEX') {
            hasCapex = true;
            sumCapexItems += Math.abs(item.currentPeriodPaise);
          }
        } else if (
          item.category === 'CFF_BORROWINGS' ||
          item.category === 'CFF_FINANCING_ITEM'
        ) {
          sumFinancingItems += item.currentPeriodPaise;
        }
      }
    }

    if (cfoTotal === null) cfoTotal = sumOperatingItems;
    if (cfiTotal === null) cfiTotal = sumInvestingItems;
    if (cffTotal === null) cffTotal = sumFinancingItems;
    if (cfNetMove === null) {
      cfNetMove = (cfoTotal || 0) + (cfiTotal || 0) + (cffTotal || 0);
    }
    if (hasCapex) capexPaise = sumCapexItems;
  } else if (model.hasCompleteVoucherData && statements.monthlyTrends.length > 0) {
    // Derive strictly disclosed operating cash movement from Receipt/Payment vouchers
    // when no explicit Cash Flow Statement was uploaded, or from Fixed Asset / Borrowing ledger deltas
    let totalCollections = 0;
    let totalPayments = 0;
    for (const m of statements.monthlyTrends) {
      totalCollections += m.collectionsPaise;
      totalPayments += m.paymentsPaise;
    }

    // Check Fixed Assets delta (Capex proxy from vouchers) and Borrowings delta
    let fixedAssetDelta = 0;
    let borrowingDelta = 0;
    let openingCashAndBank = 0;
    for (const [, s] of model.ledgerSummaries.entries()) {
      const cls = s.ledger.classification;
      if (cls === 'FixedAssets') {
        fixedAssetDelta += s.periodDebitPaise - s.periodCreditPaise;
      } else if (cls === 'LongTermBorrowings' || cls === 'ShortTermBorrowings') {
        borrowingDelta += s.periodCreditPaise - s.periodDebitPaise;
      } else if (cls === 'CashAndEquivalents' || cls === 'BankBalances') {
        openingCashAndBank += s.openingBalancePaise;
      }
    }

    capexPaise = fixedAssetDelta > 0 ? fixedAssetDelta : 0;
    cfiTotal = fixedAssetDelta !== 0 ? -fixedAssetDelta : 0;
    cffTotal = borrowingDelta - k.financeCostsPaise;
    const totalCashAndBankClosing = k.cashPaise + k.bankPaise;
    cfOpening = openingCashAndBank;
    cfClosing = totalCashAndBankClosing;
    cfNetMove = cfClosing - cfOpening;
    cfoTotal = cfNetMove - (cfiTotal || 0) - (cffTotal || 0);
  }

  const freeCashFlowPaise =
    cfoTotal !== null ? cfoTotal - (capexPaise !== null ? Math.abs(capexPaise) : 0) : null;

  const computedClosingCashPaise =
    cfOpening !== null && cfNetMove !== null ? cfOpening + cfNetMove : cfClosing;
  const bsCashAndBankPaise = k.cashPaise + k.bankPaise;
  const internalCashFlowDiffPaise =
    computedClosingCashPaise !== null && cfClosing !== null
      ? computedClosingCashPaise - cfClosing
      : 0;
  const bsVsCfDiffPaise =
    cfClosing !== null &&
    (model.hasTrialBalanceData || model.hasCompleteVoucherData || model.hasBalanceSheetStatement)
      ? cfClosing - bsCashAndBankPaise
      : null;
  const tol = policy.roundingTolerancePaise ?? 200;
  const isReconciledWithBs =
    bsVsCfDiffPaise !== null
      ? Math.abs(bsVsCfDiffPaise) <= tol && Math.abs(internalCashFlowDiffPaise) <= tol
      : false;

  const waterfallSteps: CashFlowAnalysisResult['waterfallSteps'] = [
    {
      name: 'Opening Cash',
      valuePaise: cfOpening || 0,
      valueRupees: fromPaise(cfOpening || 0),
      type: 'START',
    },
    {
      name: 'Operating Cash Flow',
      valuePaise: cfoTotal || 0,
      valueRupees: fromPaise(cfoTotal || 0),
      type: (cfoTotal || 0) >= 0 ? 'INFLOW' : 'OUTFLOW',
    },
    {
      name: 'Investing Cash Flow',
      valuePaise: cfiTotal || 0,
      valueRupees: fromPaise(cfiTotal || 0),
      type: (cfiTotal || 0) >= 0 ? 'INFLOW' : 'OUTFLOW',
    },
    {
      name: 'Financing Cash Flow',
      valuePaise: cffTotal || 0,
      valueRupees: fromPaise(cffTotal || 0),
      type: (cffTotal || 0) >= 0 ? 'INFLOW' : 'OUTFLOW',
    },
    {
      name: 'Closing Cash',
      valuePaise: cfClosing ?? computedClosingCashPaise ?? 0,
      valueRupees: fromPaise(cfClosing ?? computedClosingCashPaise ?? 0),
      type: 'END',
    },
  ];

  const monthlyCashFlowTrends = statements.monthlyTrends.map((m) => {
    const netMove = m.collectionsPaise - m.paymentsPaise;
    return {
      month: m.month,
      OperatingCF: fromPaise(netMove),
      InvestingCF: 0,
      FinancingCF: 0,
      NetCashMovement: fromPaise(netMove),
      NetProfit: fromPaise(m.netProfitPaise),
    };
  });

  const cashFlowAnalysis: CashFlowAnalysisResult = {
    hasUploadedCashFlowStatement: hasUploadedCf,
    hasDerivableCashMovement: cfoTotal !== null,
    operatingCashFlowPaise: cfoTotal,
    operatingCashFlowPriorPaise: cfoPriorTotal,
    investingCashFlowPaise: cfiTotal,
    investingCashFlowPriorPaise: cfiPriorTotal,
    financingCashFlowPaise: cffTotal,
    financingCashFlowPriorPaise: cffPriorTotal,
    netCashMovementPaise: cfNetMove,
    openingCashPaise: cfOpening,
    closingCashPaise: cfClosing ?? computedClosingCashPaise,
    capexPaise,
    freeCashFlowPaise,
    freeCashFlowDefinition:
      'Free Cash Flow = Cash Flow from Operating Activities − Identified Capital Expenditure (Capex)',
    reconciliation: {
      computedClosingCashPaise,
      reportedClosingCashPaise: cfClosing,
      internalCashFlowDiffPaise,
      balanceSheetCashAndBankPaise: bsCashAndBankPaise,
      bsVsCfDiffPaise,
      isReconciledWithBs,
      statusMessage:
        cfoTotal === null
          ? 'Upload a Cash Flow Statement or complete Cash/Bank vouchers to enable Cash Flow reconciliation.'
          : isReconciledWithBs
          ? 'Reconciled: Opening Cash + Net Cash Movement equals Closing Cash and matches Balance Sheet Cash & Bank.'
          : `Requires Review: Difference of ${formatINR(
              Math.abs(bsVsCfDiffPaise || internalCashFlowDiffPaise)
            )} detected between Cash Flow Closing Cash and Balance Sheet Cash & Bank.`,
    },
    waterfallSteps,
    monthlyCashFlowTrends,
    lineItems: cfItems,
  };

  // 3. Working Capital & Cash Conversion Cycle (Part 9)
  const dsoDays =
    k.netRevenuePaise > 0 && k.tradeReceivablesPaise >= 0
      ? Math.round((k.tradeReceivablesPaise / k.netRevenuePaise) * 365)
      : null;
  const costBaseForPayables = k.netPurchasesPaise > 0 ? k.netPurchasesPaise : k.cogsPaise;
  const dpoDays =
    costBaseForPayables > 0 && k.tradePayablesPaise >= 0
      ? Math.round((k.tradePayablesPaise / costBaseForPayables) * 365)
      : null;
  const costBaseForInventory = k.cogsPaise > 0 ? k.cogsPaise : k.netPurchasesPaise;
  const dioDays =
    costBaseForInventory > 0 && k.inventoryPaise >= 0
      ? Math.round((k.inventoryPaise / costBaseForInventory) * 365)
      : null;
  const cashConversionCycleDays =
    dsoDays !== null && dpoDays !== null && dioDays !== null
      ? dioDays + dsoDays - dpoDays
      : null;

  const monthlyWorkingCapitalTrend = statements.monthlyTrends.map((m) => {
    const rec = fromPaise(k.tradeReceivablesPaise);
    const pay = fromPaise(k.tradePayablesPaise);
    const cashB = fromPaise(k.cashPaise + k.bankPaise);
    const inv = fromPaise(k.inventoryPaise);
    const ca = rec + cashB + inv;
    const cl = pay + fromPaise(k.shortTermBorrowingsPaise);
    const mDso = m.revenuePaise > 0 ? Math.round((k.tradeReceivablesPaise / m.revenuePaise) * 30) : 0;
    const mDpo = m.purchasesPaise > 0 ? Math.round((k.tradePayablesPaise / m.purchasesPaise) * 30) : 0;
    const mDio = dioDays !== null ? dioDays : 0;
    return {
      month: m.month,
      CurrentAssets: Math.round(ca),
      CurrentLiabilities: Math.round(cl),
      NetWorkingCapital: Math.round(fromPaise(m.workingCapitalPaise)),
      Receivables: Math.round(rec),
      Payables: Math.round(pay),
      Inventory: Math.round(inv),
      CCC: mDio + mDso - mDpo,
    };
  });

  const workingCapital: WorkingCapitalAnalysisResult = {
    currentAssetsPaise: k.totalCurrentAssetsPaise,
    currentLiabilitiesPaise: k.totalCurrentLiabilitiesPaise,
    inventoryPaise: k.inventoryPaise,
    tradeReceivablesPaise: k.tradeReceivablesPaise,
    tradePayablesPaise: k.tradePayablesPaise,
    cashAndBankPaise: bsCashAndBankPaise,
    netWorkingCapitalPaise: k.workingCapitalPaise,
    priorNetWorkingCapitalPaise: priorWorkingCapitalPaise,
    dsoDays,
    dpoDays,
    dioDays,
    cashConversionCycleDays,
    cccExplanation:
      cashConversionCycleDays !== null
        ? `CCC (${cashConversionCycleDays} days) = DIO (${dioDays} days) + DSO (${dsoDays} days) − DPO (${dpoDays} days)`
        : 'Requires Revenue, Purchases/COGS, Inventory, Trade Receivables, and Trade Payables to compute Cash Conversion Cycle.',
    monthlyWorkingCapitalTrend,
  };

  // 4. Cash Flow Ratios (Part 8)
  const totalDebtPaise = k.longTermBorrowingsPaise + k.shortTermBorrowingsPaise;
  const ocfMargin =
    cfoTotal !== null && k.netRevenuePaise > 0 ? (cfoTotal / k.netRevenuePaise) * 100 : null;
  const cfToDebt =
    cfoTotal !== null && totalDebtPaise > 0 ? cfoTotal / totalDebtPaise : null;
  const cashConversionRatio =
    cfoTotal !== null && k.ebitdaPaise > 0 ? (cfoTotal / k.ebitdaPaise) * 100 : null;

  const cashFlowRatios: FinancialRatioResult[] = [
    {
      id: 'RATIO_OCF_MARGIN',
      category: 'LIQUIDITY',
      name: 'Operating Cash Flow Margin',
      formula: '(Operating Cash Flow ÷ Net Revenue) × 100',
      numeratorLabel: 'Operating Cash Flow',
      denominatorLabel: 'Net Revenue',
      numeratorPaise: cfoTotal || 0,
      denominatorPaise: k.netRevenuePaise,
      currentValue: ocfMargin,
      previousValue: null,
      unit: 'PERCENT',
      formattedCurrent: ocfMargin !== null ? `${ocfMargin.toFixed(1)}%` : 'Insufficient Data',
      formattedPrevious: 'N/A',
      status: ocfMargin !== null ? 'COMPUTED' : 'INSUFFICIENT_DATA',
      explanation: 'Measures cash generated from core operations per rupee of revenue.',
      accountMappings: ['Operating Cash Flow', 'Sales'],
      isMaterialChange: false,
    },
    {
      id: 'RATIO_CF_TO_DEBT',
      category: 'SOLVENCY',
      name: 'Cash Flow to Debt Ratio',
      formula: 'Operating Cash Flow ÷ Total Borrowings',
      numeratorLabel: 'Operating Cash Flow',
      denominatorLabel: 'Total Borrowings',
      numeratorPaise: cfoTotal || 0,
      denominatorPaise: totalDebtPaise,
      currentValue: cfToDebt,
      previousValue: null,
      unit: 'RATIO',
      formattedCurrent:
        cfToDebt !== null
          ? `${cfToDebt.toFixed(2)}x`
          : totalDebtPaise === 0
          ? 'Zero Debt'
          : 'Insufficient Data',
      formattedPrevious: 'N/A',
      status:
        cfToDebt !== null
          ? 'COMPUTED'
          : totalDebtPaise === 0
          ? 'ZERO_DENOMINATOR'
          : 'INSUFFICIENT_DATA',
      explanation: 'Assesses ability to cover total debt obligations using annual operating cash flow.',
      accountMappings: ['Operating Cash Flow', 'LongTermBorrowings', 'ShortTermBorrowings'],
      isMaterialChange: false,
    },
    {
      id: 'RATIO_CASH_CONVERSION',
      category: 'EFFICIENCY',
      name: 'EBITDA Cash Conversion Ratio',
      formula: '(Operating Cash Flow ÷ EBITDA) × 100',
      numeratorLabel: 'Operating Cash Flow',
      denominatorLabel: 'EBITDA',
      numeratorPaise: cfoTotal || 0,
      denominatorPaise: k.ebitdaPaise,
      currentValue: cashConversionRatio,
      previousValue: null,
      unit: 'PERCENT',
      formattedCurrent:
        cashConversionRatio !== null ? `${cashConversionRatio.toFixed(1)}%` : 'Insufficient Data',
      formattedPrevious: 'N/A',
      status: cashConversionRatio !== null ? 'COMPUTED' : 'INSUFFICIENT_DATA',
      explanation: 'Measures how effectively operating profit (EBITDA) converts into realized cash.',
      accountMappings: ['Operating Cash Flow', 'EBITDA'],
      isMaterialChange: false,
    },
    {
      id: 'RATIO_FREE_CASH_FLOW',
      category: 'LIQUIDITY',
      name: 'Free Cash Flow (FCF)',
      formula: 'Operating Cash Flow − Capital Expenditure',
      numeratorLabel: 'Operating Cash Flow',
      denominatorLabel: 'Capital Expenditure',
      numeratorPaise: freeCashFlowPaise || 0,
      denominatorPaise: capexPaise || 0,
      currentValue: freeCashFlowPaise !== null ? fromPaise(freeCashFlowPaise) : null,
      previousValue: null,
      unit: 'CURRENCY',
      formattedCurrent:
        freeCashFlowPaise !== null ? formatINRCompact(freeCashFlowPaise) : 'Insufficient Data',
      formattedPrevious: 'N/A',
      status: freeCashFlowPaise !== null ? 'COMPUTED' : 'INSUFFICIENT_DATA',
      explanation: cashFlowAnalysis.freeCashFlowDefinition,
      accountMappings: ['Operating Cash Flow', 'FixedAssets'],
      isMaterialChange: false,
    },
  ];

  // 5. Balance Sheet Reconciliation & Subtotal Reconciliations (Part 3.3 & Part 6)
  const totalLiabilitiesOnlyPaise =
    k.longTermBorrowingsPaise + k.totalCurrentLiabilitiesPaise;
  const totalEquityAndLiabilitiesPaise = k.totalEquityPaise + totalLiabilitiesOnlyPaise;
  const bsDiffPaise = k.totalAssetsPaise - totalEquityAndLiabilitiesPaise;
  const isBsReconciled =
    model.sourceFiles.length > 0 &&
    k.totalAssetsPaise > 0 &&
    Math.abs(bsDiffPaise) <= policy.roundingTolerancePaise;

  const subtotalReconciliations: SubtotalReconciliationRow[] = [];
  const seenReconKeys = new Set<string>();

  for (const sub of model.reportedSubtotals || []) {
    const nl = sub.normalizedLabel;
    let calcPaise: number | null = null;
    if (nl === 'gross profit' || nl === 'gross margin') calcPaise = k.grossProfitPaise;
    else if (nl === 'ebitda' || nl === 'operating profit') calcPaise = k.ebitdaPaise;
    else if (nl.includes('profit before prior period') || nl.includes('profit before exceptional')) {
      calcPaise = k.profitBeforePriorPeriodPaise;
    } else if (nl.includes('profit before tax') || nl === 'pbt') calcPaise = k.pbtPaise;
    else if (
      nl.includes('profit after tax') ||
      nl.includes('profit for the period') ||
      nl.includes('profit for the year') ||
      nl === 'pat' ||
      nl === 'net profit' ||
      nl === 'net income'
    ) {
      calcPaise = k.patPaise;
    } else if (nl === 'total income' || nl === 'total revenue') calcPaise = k.totalIncomePaise;
    else if (nl === 'total expenses' || nl === 'total expense') calcPaise = k.totalExpensesPaise;
    else if (nl === 'total assets' || nl === 'total asset') calcPaise = k.totalAssetsPaise;
    else if (
      nl === 'total equity and liabilities' ||
      nl === 'total equity & liabilities' ||
      nl === 'total liabilities'
    ) {
      calcPaise = totalEquityAndLiabilitiesPaise;
    } else if (nl === 'total current assets' || nl === 'current assets') {
      calcPaise = k.totalCurrentAssetsPaise;
    } else if (
      nl === 'total non-current assets' ||
      nl === 'total non current assets' ||
      nl === 'non-current assets' ||
      nl === 'non current assets'
    ) {
      calcPaise = k.totalNonCurrentAssetsPaise;
    } else if (nl === 'total current liabilities' || nl === 'current liabilities') {
      calcPaise = k.totalCurrentLiabilitiesPaise;
    } else if (
      nl === 'total non-current liabilities' ||
      nl === 'total non current liabilities' ||
      nl === 'non-current liabilities' ||
      nl === 'non current liabilities'
    ) {
      calcPaise = k.longTermBorrowingsPaise;
    } else if (
      nl.includes('property, plant') ||
      nl === 'total fixed assets' ||
      nl === 'tangible and intangible assets'
    ) {
      calcPaise = k.fixedAssetsPaise;
    } else if (
      nl === 'total equity' ||
      nl === 'total shareholders funds' ||
      nl === "total shareholders' funds" ||
      nl === 'shareholders funds' ||
      nl === "shareholders' funds"
    ) {
      calcPaise = k.totalEquityPaise;
    }

    if (calcPaise !== null) {
      const diff = sub.reportedCurrentPaise - calcPaise;
      const key = `${sub.statementType}:${sub.label}`;
      if (!seenReconKeys.has(key)) {
        seenReconKeys.add(key);
        subtotalReconciliations.push({
          statementType: sub.statementType,
          label: sub.label,
          reportedPaise: sub.reportedCurrentPaise,
          priorReportedPaise: sub.reportedPreviousPaise,
          calculatedPaise: calcPaise,
          differencePaise: diff,
          isMatched: Math.abs(diff) <= policy.roundingTolerancePaise,
          worksheetName: sub.worksheetName,
          cellReference: sub.cellReference,
          priorCellReference: sub.priorCellReference,
        });
      }
    }
  }

  // Also include Cash Flow statement subtotal reconciliations if Cash Flow was uploaded
  for (const cf of cfItems) {
    if (
      cf.category === 'CFO_OPERATING_TOTAL' ||
      cf.category === 'CFI_INVESTING_TOTAL' ||
      cf.category === 'CFF_FINANCING_TOTAL' ||
      cf.category === 'CF_NET_CHANGE' ||
      cf.category === 'CF_OPENING_CASH' ||
      cf.category === 'CF_CLOSING_CASH'
    ) {
      const calcVal =
        cf.category === 'CFO_OPERATING_TOTAL'
          ? cfoTotal ?? 0
          : cf.category === 'CFI_INVESTING_TOTAL'
          ? cfiTotal ?? 0
          : cf.category === 'CFF_FINANCING_TOTAL'
          ? cffTotal ?? 0
          : cf.category === 'CF_NET_CHANGE'
          ? (cfoTotal ?? 0) + (cfiTotal ?? 0) + (cffTotal ?? 0)
          : cf.category === 'CF_OPENING_CASH'
          ? cfOpening ?? 0
          : computedClosingCashPaise ?? 0;
      const diff = cf.currentPeriodPaise - calcVal;
      subtotalReconciliations.push({
        statementType: 'CASH_FLOW',
        label: cf.particulars,
        reportedPaise: cf.currentPeriodPaise,
        priorReportedPaise: cf.previousPeriodPaise,
        calculatedPaise: calcVal,
        differencePaise: diff,
        isMatched: Math.abs(diff) <= policy.roundingTolerancePaise,
        worksheetName: cf.worksheetName,
        cellReference: cf.cellReference,
        priorCellReference: cf.priorCellReference,
      });
    }
  }

  // Helper to check if a specific subtotal keyword has a mismatch in subtotalReconciliations
  const findSubtotalRecon = (stmtType: 'PNL' | 'BALANCE_SHEET' | 'CASH_FLOW', keywords: string[]) => {
    return subtotalReconciliations.find(
      (sr) =>
        sr.statementType === stmtType &&
        keywords.some((kw) => sr.label.toLowerCase().includes(kw))
    );
  };

  const hasSubtotalMismatch = (
    stmtType: 'PNL' | 'BALANCE_SHEET' | 'CASH_FLOW',
    keywords: string[]
  ): boolean => {
    const found = findSubtotalRecon(stmtType, keywords);
    return found ? !found.isMatched : false;
  };

  // All Balance Sheet subtotals must match AND Total Assets === Total Equity & Liabilities
  const anyBsSubtotalMismatch = subtotalReconciliations.some(
    (sr) => sr.statementType === 'BALANCE_SHEET' && !sr.isMatched
  );
  const isStrictBsReconciled = isBsReconciled && !anyBsSubtotalMismatch;

  // Build Line-Item Verification Rows for Balance Sheet (Part 6)
  const buildBsVerifyRow = (
    id: string,
    component: string,
    category: BalanceSheetVerificationRow['category'],
    calculatedPaise: number,
    priorPeriodPaise: number | null,
    lineCode: string,
    subtotalKeywords: string[],
    classifications: AccountClassification[] | undefined,
    formula: string,
    forceReportedPaise?: number | null,
    forceDifferencePaise?: number
  ): BalanceSheetVerificationRow => {
    const matchedSub =
      subtotalKeywords.length > 0 ? findSubtotalRecon('BALANCE_SHEET', subtotalKeywords) : undefined;
    const reportedSourcePaise =
      forceReportedPaise !== undefined
        ? forceReportedPaise
        : matchedSub
        ? matchedSub.reportedPaise
        : model.hasBalanceSheetStatement
        ? calculatedPaise
        : null;
    const differencePaise =
      forceDifferencePaise !== undefined
        ? forceDifferencePaise
        : reportedSourcePaise !== null
        ? reportedSourcePaise - calculatedPaise
        : 0;
    const isRowMatched =
      model.sourceFiles.length > 0 &&
      k.totalAssetsPaise > 0 &&
      Math.abs(differencePaise) <= policy.roundingTolerancePaise &&
      (category !== 'RECONCILIATION' ? isStrictBsReconciled : isBsReconciled);

    return {
      id,
      component,
      category,
      reportedSourcePaise,
      calculatedPaise,
      priorPeriodPaise,
      differencePaise,
      cellReference: matchedSub?.cellReference || getLineCellRef(lineCode),
      priorCellReference: matchedSub?.priorCellReference || getLinePriorCellRef(lineCode),
      status:
        model.sourceFiles.length === 0 || k.totalAssetsPaise === 0
          ? 'INSUFFICIENT DATA'
          : isRowMatched
          ? 'RECONCILED'
          : 'REQUIRES REVIEW',
      classifications,
      formula,
    };
  };

  const otherCurrentLiabCombinedPaise =
    k.shortTermBorrowingsPaise + k.taxLiabilitiesPaise + k.otherCurrentLiabilitiesPaise;
  const priorOtherCurrentLiabCombinedPaise =
    (getLinePrev('SHORT_TERM_BORROWINGS') || 0) +
    (getLinePrev('TAX_AND_OTHER_LIAB') || 0);
  const otherCurrentAssetsCombinedPaise = k.inventoryPaise + k.otherCurrentAssetsPaise;
  const priorOtherCurrentAssetsCombinedPaise =
    (getLinePrev('INVENTORY') || 0) + (getLinePrev('OTHER_CURRENT_ASSETS') || 0);

  const balanceSheetVerificationRows: BalanceSheetVerificationRow[] = [
    buildBsVerifyRow(
      'BS_VER_EQUITY',
      "Shareholders' Funds (Share Capital + Reserves & Surplus)",
      'EQUITY_AND_LIABILITIES',
      k.totalEquityPaise,
      getLinePrev('EQUITY'),
      'EQUITY',
      ['shareholders', 'total equity'],
      ['Equity'],
      'Share Capital + Reserves and Surplus'
    ),
    buildBsVerifyRow(
      'BS_VER_NON_CURRENT_LIAB',
      'Non-Current Liabilities (Long-Term Borrowings & Provisions)',
      'EQUITY_AND_LIABILITIES',
      k.longTermBorrowingsPaise,
      getLinePrev('NON_CURRENT_LIABILITIES'),
      'NON_CURRENT_LIABILITIES',
      ['non-current liabilities', 'non current liabilities'],
      ['LongTermBorrowings'],
      'Long-Term Borrowings + Other Long-Term Liabilities + Long-Term Provisions'
    ),
    buildBsVerifyRow(
      'BS_VER_TRADE_PAYABLES',
      'Trade Payables (MSME + Other Creditors)',
      'EQUITY_AND_LIABILITIES',
      k.tradePayablesPaise,
      getLinePrev('TRADE_PAYABLES'),
      'TRADE_PAYABLES',
      ['trade payables'],
      ['TradePayables'],
      'Total outstanding dues of MSME + Total outstanding dues of creditors other than MSME'
    ),
    buildBsVerifyRow(
      'BS_VER_OTHER_CURR_LIAB',
      'Other Current Liabilities & Short-Term Provisions',
      'EQUITY_AND_LIABILITIES',
      otherCurrentLiabCombinedPaise,
      hasPriorData ? priorOtherCurrentLiabCombinedPaise : null,
      'TAX_AND_OTHER_LIAB',
      [],
      ['ShortTermBorrowings', 'TaxLiabilities', 'OtherCurrentLiabilities'],
      'Short-Term Borrowings + Other Current Liabilities + Short-Term Provisions'
    ),
    buildBsVerifyRow(
      'BS_VER_TOTAL_CURR_LIAB',
      'Total Current Liabilities',
      'EQUITY_AND_LIABILITIES',
      k.totalCurrentLiabilitiesPaise,
      priorCurrentLiabilitiesPaise,
      'TOTAL_CURRENT_LIABILITIES',
      ['current liabilities'],
      ['ShortTermBorrowings', 'TradePayables', 'TaxLiabilities', 'OtherCurrentLiabilities'],
      'Trade Payables + Other Current Liabilities + Short-Term Provisions + Short-Term Borrowings'
    ),
    buildBsVerifyRow(
      'BS_VER_TOTAL_EQ_LIAB',
      'Total Equity and Liabilities',
      'EQUITY_AND_LIABILITIES',
      totalEquityAndLiabilitiesPaise,
      getLinePrev('TOTAL_EQUITY_AND_LIABILITIES'),
      'TOTAL_EQUITY_AND_LIABILITIES',
      ['total equity and liabilities', 'total equity & liabilities', 'total liabilities'],
      [
        'Equity',
        'LongTermBorrowings',
        'ShortTermBorrowings',
        'TradePayables',
        'TaxLiabilities',
        'OtherCurrentLiabilities',
      ],
      "Shareholders' Funds + Non-Current Liabilities + Current Liabilities"
    ),
    buildBsVerifyRow(
      'BS_VER_NON_CURRENT_ASSETS',
      'Non-Current Assets (PPE, Intangibles, DTA & Long-Term Advances)',
      'ASSETS',
      k.totalNonCurrentAssetsPaise,
      getLinePrev('NON_CURRENT_ASSETS'),
      'NON_CURRENT_ASSETS',
      ['non-current assets', 'non current assets'],
      ['FixedAssets', 'Investments', 'NonCurrentAssets'],
      'Property, Plant & Equipment + Intangible Assets + Deferred Tax Assets + Long-Term Loans & Advances'
    ),
    buildBsVerifyRow(
      'BS_VER_TRADE_RECEIVABLES',
      'Trade Receivables',
      'ASSETS',
      k.tradeReceivablesPaise,
      getLinePrev('TRADE_RECEIVABLES'),
      'TRADE_RECEIVABLES',
      ['trade receivables'],
      ['TradeReceivables'],
      'Closing Balance of Trade Receivables / Sundry Debtors'
    ),
    buildBsVerifyRow(
      'BS_VER_CASH_AND_BANK',
      'Cash and Cash Equivalents',
      'ASSETS',
      bsCashAndBankPaise,
      priorCashAndBankPaise,
      'CASH_AND_BANK',
      ['cash and cash equivalents'],
      ['CashAndEquivalents', 'BankBalances'],
      'Cash-in-Hand + Balances with Scheduled Banks'
    ),
    buildBsVerifyRow(
      'BS_VER_OTHER_CURR_ASSETS',
      'Other Current Assets & Inventories',
      'ASSETS',
      otherCurrentAssetsCombinedPaise,
      hasPriorData ? priorOtherCurrentAssetsCombinedPaise : null,
      'OTHER_CURRENT_ASSETS',
      [],
      ['Inventory', 'OtherCurrentAssets'],
      'Inventories + Short-Term Loans & Advances + Other Current Assets'
    ),
    buildBsVerifyRow(
      'BS_VER_TOTAL_CURR_ASSETS',
      'Total Current Assets',
      'ASSETS',
      k.totalCurrentAssetsPaise,
      priorCurrentAssetsPaise,
      'TOTAL_CURRENT_ASSETS',
      ['current assets'],
      ['Inventory', 'TradeReceivables', 'CashAndEquivalents', 'BankBalances', 'OtherCurrentAssets'],
      'Inventories + Trade Receivables + Cash & Cash Equivalents + Other Current Assets'
    ),
    buildBsVerifyRow(
      'BS_VER_TOTAL_ASSETS',
      'Total Assets',
      'ASSETS',
      k.totalAssetsPaise,
      priorTotalAssetsPaise,
      'TOTAL_ASSETS',
      ['total assets'],
      [
        'FixedAssets',
        'Investments',
        'NonCurrentAssets',
        'Inventory',
        'TradeReceivables',
        'CashAndEquivalents',
        'BankBalances',
        'OtherCurrentAssets',
      ],
      'Non-Current Assets + Current Assets'
    ),
    buildBsVerifyRow(
      'BS_VER_NET_WORKING_CAPITAL',
      'Net Working Capital',
      'WORKING_CAPITAL',
      k.workingCapitalPaise,
      priorWorkingCapitalPaise,
      'TOTAL_CURRENT_ASSETS',
      [],
      [
        'Inventory',
        'TradeReceivables',
        'CashAndEquivalents',
        'BankBalances',
        'OtherCurrentAssets',
        'ShortTermBorrowings',
        'TradePayables',
        'TaxLiabilities',
        'OtherCurrentLiabilities',
      ],
      'Total Current Assets − Total Current Liabilities'
    ),
    buildBsVerifyRow(
      'BS_VER_RECON_DIFF',
      'Balance Sheet Reconciliation Difference (Assets − Equity & Liabilities)',
      'RECONCILIATION',
      bsDiffPaise,
      priorTotalAssetsPaise !== null && getLinePrev('TOTAL_EQUITY_AND_LIABILITIES') !== null
        ? priorTotalAssetsPaise - (getLinePrev('TOTAL_EQUITY_AND_LIABILITIES') ?? 0)
        : null,
      'TOTAL_ASSETS',
      [],
      undefined,
      'Total Assets − Total Equity and Liabilities (Must equal ₹0.00)',
      0,
      bsDiffPaise
    ),
  ];

  // 6. Build 17 CFO Executive KPI Cards (Part 4.1)
  const buildCurrencyKpi = (
    id: string,
    label: string,
    currPaise: number | null,
    prevPaise: number | null,
    budPaise: number | null,
    formula: string,
    classifications?: AccountClassification[],
    requiresReview = false,
    sourceCellRef?: string,
    priorCellRef?: string,
    customMissingLabel?: string
  ): CfoKpiCardItem => {
    if (model.sourceFiles.length === 0 || currPaise === null) {
      return {
        id,
        label,
        unit: 'CURRENCY',
        currentPaiseOrVal: null,
        previousPaiseOrVal: null,
        budgetPaiseOrVal: null,
        currentFormatted: customMissingLabel || 'Insufficient Data',
        previousFormatted: 'N/A',
        budgetFormatted: 'N/A',
        varianceAmountFormatted: 'N/A',
        variancePercentFormatted: 'N/A',
        statusIndicator: customMissingLabel ? 'Requires Review' : 'Insufficient Data',
        formula,
        classifications,
        sourceCellRef,
        priorCellRef,
      };
    }
    const hasPrev = prevPaise !== null && hasPriorData;
    const variance = hasPrev ? computeSafeVariance(currPaise, prevPaise!) : null;
    return {
      id,
      label,
      unit: 'CURRENCY',
      currentPaiseOrVal: currPaise,
      previousPaiseOrVal: hasPrev ? prevPaise : null,
      budgetPaiseOrVal: budPaise,
      currentFormatted: formatINRCompact(currPaise),
      previousFormatted: hasPrev ? formatINRCompact(prevPaise!) : 'N/A',
      budgetFormatted: budPaise !== null ? formatINRCompact(budPaise) : 'N/A',
      varianceAmountFormatted: variance ? formatINRCompact(variance.diffPaise) : 'N/A',
      variancePercentFormatted: variance ? variance.label : 'N/A',
      statusIndicator: requiresReview ? 'Requires Review' : 'Reconciled',
      formula,
      classifications,
      sourceCellRef,
      priorCellRef,
    };
  };

  const buildPercentOrRatioKpi = (
    id: string,
    label: string,
    unit: 'PERCENT' | 'RATIO',
    currVal: number | null,
    prevVal: number | null,
    formula: string,
    isFavorableCheck: (v: number) => boolean,
    customMissingLabel?: string,
    requiresReview = false
  ): CfoKpiCardItem => {
    if (model.sourceFiles.length === 0 || currVal === null || !Number.isFinite(currVal)) {
      return {
        id,
        label,
        unit,
        currentPaiseOrVal: null,
        previousPaiseOrVal: null,
        budgetPaiseOrVal: null,
        currentFormatted: customMissingLabel || 'Insufficient Data',
        previousFormatted: 'N/A',
        budgetFormatted: 'N/A',
        varianceAmountFormatted: 'N/A',
        variancePercentFormatted: 'N/A',
        statusIndicator: customMissingLabel ? 'Requires Review' : 'Insufficient Data',
        formula,
      };
    }
    const suffix = unit === 'PERCENT' ? '%' : 'x';
    const decimals = unit === 'PERCENT' ? 1 : 2;
    const hasPrev = prevVal !== null && Number.isFinite(prevVal) && hasPriorData;
    const diff = hasPrev ? currVal - prevVal! : null;
    return {
      id,
      label,
      unit,
      currentPaiseOrVal: currVal,
      previousPaiseOrVal: hasPrev ? prevVal : null,
      budgetPaiseOrVal: null,
      currentFormatted: `${currVal.toFixed(decimals)}${suffix}`,
      previousFormatted: hasPrev ? `${prevVal!.toFixed(decimals)}${suffix}` : 'N/A',
      budgetFormatted: 'N/A',
      varianceAmountFormatted:
        diff !== null
          ? `${diff >= 0 ? '+' : ''}${diff.toFixed(decimals)}${ unit === 'PERCENT' ? ' pp' : 'x'}`
          : 'N/A',
      variancePercentFormatted:
        diff !== null && prevVal && Math.abs(prevVal) > 0.001
          ? `${diff >= 0 ? '+' : ''}${((diff / Math.abs(prevVal)) * 100).toFixed(1)}%`
          : 'N/A',
      statusIndicator: requiresReview
        ? 'Requires Review'
        : isFavorableCheck(currVal)
        ? 'Reconciled'
        : 'Attention Required',
      formula,
    };
  };

  const gpMarginCurr =
    k.hasGrossProfitData && k.netRevenuePaise > 0
      ? (k.grossProfitPaise / k.netRevenuePaise) * 100
      : null;
  const gpMarginPrev =
    k.hasGrossProfitData &&
    priorRevenuePaise &&
    priorRevenuePaise > 0 &&
    priorGrossProfitPaise !== null
      ? (priorGrossProfitPaise / priorRevenuePaise) * 100
      : null;

  const ebitdaMarginCurr =
    k.netRevenuePaise > 0 ? (k.ebitdaPaise / k.netRevenuePaise) * 100 : null;
  const ebitdaMarginPrev =
    priorRevenuePaise && priorRevenuePaise > 0 && priorEbitdaPaise !== null
      ? (priorEbitdaPaise / priorRevenuePaise) * 100
      : null;

  const npMarginCurr = k.netRevenuePaise > 0 ? (k.patPaise / k.netRevenuePaise) * 100 : null;
  const npMarginPrev =
    priorRevenuePaise && priorRevenuePaise > 0 && priorPatPaise !== null
      ? (priorPatPaise / priorRevenuePaise) * 100
      : null;

  const currentRatioObj = existingRatios.find((r) => r.id === 'RATIO_CURRENT');
  const debtEquityObj = existingRatios.find((r) => r.id === 'RATIO_DEBT_EQUITY');
  const roceObj = existingRatios.find((r) => r.id === 'RATIO_ROCE');

  const cfoOperatingCellRef = cfItems.find((i) => i.category === 'CFO_OPERATING_TOTAL')?.cellReference;
  const cfoOperatingPriorRef = cfItems.find((i) => i.category === 'CFO_OPERATING_TOTAL')?.priorCellReference;

  const pnlIncomeMismatch = hasSubtotalMismatch('PNL', ['total income', 'total revenue']);
  const pnlPatMismatch = hasSubtotalMismatch('PNL', [
    'profit after tax',
    'profit for the period',
    'profit for the year',
    'net profit',
    'profit before tax',
  ]);
  const cfOperatingMismatch = hasSubtotalMismatch('CASH_FLOW', ['operating']);

  const cfoKpis: CfoKpiCardItem[] = [
    buildCurrencyKpi(
      'KPI_TOTAL_REVENUE',
      'Total Revenue',
      k.netRevenuePaise,
      priorRevenuePaise,
      hasStatementBudget ? sumBudgetSales : null,
      'Revenue from Operations = Gross Revenue from Operations − Sales Returns',
      ['Sales', 'SalesReturn', 'OtherOperatingIncome'],
      pnlIncomeMismatch,
      getLineCellRef('REVENUE_NET'),
      getLinePriorCellRef('REVENUE_NET')
    ),
    buildCurrencyKpi(
      'KPI_GROSS_PROFIT',
      'Gross Profit',
      k.hasGrossProfitData ? k.grossProfitPaise : null,
      priorGrossProfitPaise,
      budgetGrossProfitPaise,
      k.hasGrossProfitData
        ? 'Gross Profit = Net Revenue − Cost of Goods Sold − Direct Expenses'
        : 'Requires Mapping: Cost of Goods Sold / Direct Cost breakdown is not separately reported in the uploaded statement.',
      ['Sales', 'Purchases', 'DirectExpenses'],
      !k.hasGrossProfitData,
      getLineCellRef('GROSS_PROFIT'),
      getLinePriorCellRef('GROSS_PROFIT'),
      !k.hasGrossProfitData ? 'Requires Mapping' : undefined
    ),
    buildPercentOrRatioKpi(
      'KPI_GP_MARGIN',
      'Gross Profit Margin',
      'PERCENT',
      gpMarginCurr,
      gpMarginPrev,
      k.hasGrossProfitData
        ? '(Gross Profit ÷ Net Revenue) × 100'
        : 'N/A: Gross Profit is not separately reported without Cost of Sales mapping.',
      (v) => v >= 20,
      !k.hasGrossProfitData ? 'N/A' : undefined,
      !k.hasGrossProfitData
    ),
    buildCurrencyKpi(
      'KPI_EBITDA',
      'EBITDA',
      k.ebitdaPaise,
      priorEbitdaPaise,
      budgetEbitdaPaise,
      'EBITDA = Total Income − Cost of Goods Sold − Employee, Admin & Selling Overheads (equals PBT + Finance Costs + Depreciation + Prior Period Items)',
      ['Sales', 'Purchases', 'EmployeeExpenses', 'AdminExpenses', 'SellingExpenses'],
      pnlIncomeMismatch || pnlPatMismatch,
      getLineCellRef('EBITDA') || getLineCellRef('PBT'),
      getLinePriorCellRef('EBITDA') || getLinePriorCellRef('PBT')
    ),
    buildPercentOrRatioKpi(
      'KPI_EBITDA_MARGIN',
      'EBITDA Margin',
      'PERCENT',
      ebitdaMarginCurr,
      ebitdaMarginPrev,
      '(EBITDA ÷ Net Revenue) × 100',
      (v) => v >= 12,
      undefined,
      pnlIncomeMismatch || pnlPatMismatch
    ),
    buildCurrencyKpi(
      'KPI_NET_PROFIT',
      'Net Profit (PAT)',
      k.patPaise,
      priorPatPaise,
      null,
      'Profit After Tax = Profit Before Tax − Current Tax − Deferred Tax',
      ['Sales', 'Purchases', 'EmployeeExpenses', 'AdminExpenses', 'FinanceCosts', 'CurrentTax', 'DeferredTax', 'TaxExpense'],
      pnlPatMismatch,
      getLineCellRef('PAT'),
      getLinePriorCellRef('PAT')
    ),
    buildPercentOrRatioKpi(
      'KPI_NP_MARGIN',
      'Net Profit Margin',
      'PERCENT',
      npMarginCurr,
      npMarginPrev,
      '(Net Profit ÷ Net Revenue) × 100',
      (v) => v >= 5,
      undefined,
      pnlPatMismatch
    ),
    buildCurrencyKpi(
      'KPI_TOTAL_ASSETS',
      'Total Assets',
      k.totalAssetsPaise,
      priorTotalAssetsPaise,
      null,
      'Total Assets = Non-Current Assets + Current Assets',
      [
        'FixedAssets',
        'Investments',
        'NonCurrentAssets',
        'Inventory',
        'TradeReceivables',
        'CashAndEquivalents',
        'BankBalances',
      ],
      !isStrictBsReconciled,
      getLineCellRef('TOTAL_ASSETS'),
      getLinePriorCellRef('TOTAL_ASSETS')
    ),
    buildCurrencyKpi(
      'KPI_CURRENT_ASSETS',
      'Current Assets',
      k.totalCurrentAssetsPaise,
      priorCurrentAssetsPaise,
      null,
      'Inventory + Trade Receivables + Cash & Bank + Other Current Assets',
      ['Inventory', 'TradeReceivables', 'CashAndEquivalents', 'BankBalances', 'OtherCurrentAssets'],
      !isStrictBsReconciled || hasSubtotalMismatch('BALANCE_SHEET', ['current assets']),
      getLineCellRef('TOTAL_CURRENT_ASSETS'),
      getLinePriorCellRef('TOTAL_CURRENT_ASSETS')
    ),
    buildCurrencyKpi(
      'KPI_CURRENT_LIABILITIES',
      'Current Liabilities',
      k.totalCurrentLiabilitiesPaise,
      priorCurrentLiabilitiesPaise,
      null,
      'Short-Term Borrowings + Trade Payables + Tax & Other Current Liabilities',
      ['ShortTermBorrowings', 'TradePayables', 'TaxLiabilities', 'OtherCurrentLiabilities'],
      !isStrictBsReconciled || hasSubtotalMismatch('BALANCE_SHEET', ['current liabilities']),
      getLineCellRef('TOTAL_CURRENT_LIABILITIES'),
      getLinePriorCellRef('TOTAL_CURRENT_LIABILITIES')
    ),
    buildCurrencyKpi(
      'KPI_CASH_BANK',
      'Cash and Bank',
      bsCashAndBankPaise,
      priorCashAndBankPaise,
      null,
      'Cash and Cash Equivalents + Scheduled Bank Balances',
      ['CashAndEquivalents', 'BankBalances'],
      !isStrictBsReconciled || (hasUploadedCf && !isReconciledWithBs),
      getLineCellRef('CASH_AND_BANK'),
      getLinePriorCellRef('CASH_AND_BANK')
    ),
    buildCurrencyKpi(
      'KPI_NET_WORKING_CAPITAL',
      'Net Working Capital',
      k.workingCapitalPaise,
      priorWorkingCapitalPaise,
      null,
      'Net Working Capital = Current Assets − Current Liabilities',
      [
        'Inventory',
        'TradeReceivables',
        'CashAndEquivalents',
        'BankBalances',
        'OtherCurrentAssets',
        'ShortTermBorrowings',
        'TradePayables',
        'TaxLiabilities',
        'OtherCurrentLiabilities',
      ],
      !isStrictBsReconciled ||
        hasSubtotalMismatch('BALANCE_SHEET', ['current assets']) ||
        hasSubtotalMismatch('BALANCE_SHEET', ['current liabilities']),
      getLineCellRef('TOTAL_CURRENT_ASSETS'),
      getLinePriorCellRef('TOTAL_CURRENT_ASSETS')
    ),
    buildCurrencyKpi(
      'KPI_OPERATING_CF',
      'Operating Cash Flow',
      cfoTotal,
      cfoPriorTotal,
      null,
      'Net Cash Generated from / (Used in) Operating Activities',
      undefined,
      cfOperatingMismatch || (hasUploadedCf && !isReconciledWithBs),
      cfoOperatingCellRef,
      cfoOperatingPriorRef
    ),
    buildCurrencyKpi(
      'KPI_FREE_CF',
      'Free Cash Flow',
      freeCashFlowPaise,
      null,
      null,
      cashFlowAnalysis.freeCashFlowDefinition,
      undefined,
      cfOperatingMismatch || (hasUploadedCf && !isReconciledWithBs)
    ),
    buildPercentOrRatioKpi(
      'KPI_CURRENT_RATIO',
      'Current Ratio',
      'RATIO',
      currentRatioObj?.currentValue ?? null,
      currentRatioObj?.previousValue ?? null,
      'Current Assets ÷ Current Liabilities',
      (v) => v >= 1.25,
      undefined,
      !isStrictBsReconciled
    ),
    buildPercentOrRatioKpi(
      'KPI_DEBT_EQUITY',
      'Debt Equity Ratio',
      'RATIO',
      debtEquityObj?.currentValue ?? null,
      debtEquityObj?.previousValue ?? null,
      'Total Borrowings ÷ Shareholders Equity',
      (v) => v <= 2.0,
      undefined,
      !isStrictBsReconciled
    ),
    buildPercentOrRatioKpi(
      'KPI_ROCE',
      'Return on Capital Employed',
      'PERCENT',
      roceObj?.currentValue ?? null,
      roceObj?.previousValue ?? null,
      '(EBIT ÷ Capital Employed) × 100',
      (v) => v >= 12,
      undefined,
      !isStrictBsReconciled
    ),
  ];

  // 7. Deterministic CFO Suggestions & Management Recommendations (Part 10)
  const cfoSuggestions: CfoManagementSuggestion[] = [];
  const primaryRef = model.sourceFiles.map((f) => f.filename).join(', ') || 'Uploaded Dataset';

  // Rule 1: Revenue movement
  if (priorRevenuePaise && priorRevenuePaise > 0) {
    const revVar = computeSafeVariance(k.netRevenuePaise, priorRevenuePaise);
    if (revVar.diffPaise < 0) {
      cfoSuggestions.push({
        id: 'CFO_SUG_REV_DECLINE',
        category: 'Revenue Performance',
        factualObservation: `Net Revenue declined by ${revVar.label} compared to the comparative period.`,
        actualFinancialValue: formatINR(k.netRevenuePaise),
        comparativeValue: formatINR(priorRevenuePaise),
        calculatedVariance: `${formatINR(revVar.diffPaise)} (${revVar.label})`,
        quantifiedFinancialImpact: `Top-line contraction of ${formatINR(Math.abs(revVar.diffPaise))}`,
        riskClassification: revVar.percent && revVar.percent < -10 ? 'HIGH RISK' : 'MODERATE RISK',
        suggestedManagementAction:
          'Review customer-wise sales volume, credit notes/returns, and pricing realizations across core product lines.',
        supportingSourceRef: primaryRef,
      });
    } else if (revVar.diffPaise > 0) {
      cfoSuggestions.push({
        id: 'CFO_SUG_REV_GROWTH',
        category: 'Revenue Performance',
        factualObservation: `Net Revenue expanded by ${revVar.label} over the comparative period.`,
        actualFinancialValue: formatINR(k.netRevenuePaise),
        comparativeValue: formatINR(priorRevenuePaise),
        calculatedVariance: `+${formatINR(revVar.diffPaise)} (${revVar.label})`,
        quantifiedFinancialImpact: `Incremental revenue generation of ${formatINR(revVar.diffPaise)}`,
        riskClassification: 'FAVORABLE',
        suggestedManagementAction:
          'Ensure incremental revenue is matched by timely debtor collections so receivables do not expand disproportionately.',
        supportingSourceRef: primaryRef,
      });
    }
  }

  // Rule 2: Gross Margin check
  if (gpMarginCurr !== null && gpMarginPrev !== null && gpMarginCurr < gpMarginPrev - 1.5) {
    const bpsDrop = gpMarginCurr - gpMarginPrev;
    const rupeeImpact = Math.round((Math.abs(bpsDrop) / 100) * k.netRevenuePaise);
    cfoSuggestions.push({
      id: 'CFO_SUG_GP_DETERIORATION',
      category: 'Gross Margin',
      factualObservation: `Gross Profit Margin compressed from ${gpMarginPrev.toFixed(1)}% to ${gpMarginCurr.toFixed(1)}%.`,
      actualFinancialValue: `${gpMarginCurr.toFixed(1)}% (${formatINR(k.grossProfitPaise)})`,
      comparativeValue: `${gpMarginPrev.toFixed(1)}% (${formatINR(priorGrossProfitPaise || 0)})`,
      calculatedVariance: `${bpsDrop.toFixed(1)} percentage points`,
      quantifiedFinancialImpact: `Estimated margin erosion of ${formatINR(rupeeImpact)} on current revenue volume`,
      riskClassification: bpsDrop < -4 ? 'HIGH RISK' : 'MODERATE RISK',
      suggestedManagementAction:
        'Audit raw material purchase rates, inward freight/direct overheads, and closing inventory valuation.',
      supportingSourceRef: primaryRef,
    });
  }

  // Rule 3: Operating Expense Ratio check
  const totalIndOpex =
    k.employeeExpensesPaise + k.adminExpensesPaise + k.sellingExpensesPaise;
  if (k.netRevenuePaise > 0 && totalIndOpex > 0) {
    const opexPct = (totalIndOpex / k.netRevenuePaise) * 100;
    if (opexPct > 25) {
      cfoSuggestions.push({
        id: 'CFO_SUG_OPEX_LOAD',
        category: 'Operating Expenses',
        factualObservation: `Employee, Administrative, and Selling overheads consume ${opexPct.toFixed(1)}% of Net Revenue.`,
        actualFinancialValue: formatINR(totalIndOpex),
        comparativeValue: `Net Revenue: ${formatINR(k.netRevenuePaise)}`,
        calculatedVariance: `${opexPct.toFixed(1)}% Overhead-to-Sales Ratio`,
        quantifiedFinancialImpact: `Every 1% reduction in overheads saves ${formatINR(
          Math.round(totalIndOpex * 0.01)
        )}`,
        riskClassification: opexPct > 35 ? 'HIGH RISK' : 'WATCH',
        suggestedManagementAction:
          'Perform cost-center and ledger-level variance analysis on fixed administrative and selling expenses.',
        supportingSourceRef: primaryRef,
      });
    }
  }

  // Rule 4: Liquidity & Current Ratio check
  if (currentRatioObj && currentRatioObj.currentValue !== null) {
    if (currentRatioObj.currentValue < 1.2) {
      cfoSuggestions.push({
        id: 'CFO_SUG_LOW_LIQUIDITY',
        category: 'Liquidity',
        factualObservation: `Current Ratio stands at ${currentRatioObj.currentValue.toFixed(2)}x, below the 1.25x institutional liquidity cushion.`,
        actualFinancialValue: `${currentRatioObj.currentValue.toFixed(2)}x (CA: ${formatINRCompact(
          k.totalCurrentAssetsPaise
        )})`,
        comparativeValue: `Benchmark ≥ 1.25x (CL: ${formatINRCompact(
          k.totalCurrentLiabilitiesPaise
        )})`,
        calculatedVariance: `Net Working Capital: ${formatINR(k.workingCapitalPaise)}`,
        quantifiedFinancialImpact:
          k.workingCapitalPaise < 0
            ? `Working capital deficit of ${formatINR(Math.abs(k.workingCapitalPaise))}`
            : 'Tight short-term liquidity buffer',
        riskClassification: currentRatioObj.currentValue < 1.0 ? 'HIGH RISK' : 'MODERATE RISK',
        suggestedManagementAction:
          'Accelerate realization of >60-day overdue receivables and align supplier payment schedules with cash inflows.',
        supportingSourceRef: primaryRef,
      });
    }
  }

  // Rule 5: Operating Cash Flow vs Net Profit Divergence
  if (cfoTotal !== null) {
    if (cfoTotal < 0) {
      cfoSuggestions.push({
        id: 'CFO_SUG_NEG_OCF',
        category: 'Operating Cash Flow',
        factualObservation: `Operating Cash Flow is negative at ${formatINR(cfoTotal)} for the period.`,
        actualFinancialValue: formatINR(cfoTotal),
        comparativeValue: `Net Profit (PAT): ${formatINR(k.patPaise)}`,
        calculatedVariance: `Cash-to-PAT Gap: ${formatINR(cfoTotal - k.patPaise)}`,
        quantifiedFinancialImpact: `Operating cash burn of ${formatINR(Math.abs(cfoTotal))}`,
        riskClassification: 'HIGH RISK',
        suggestedManagementAction:
          'Inspect working capital absorption in receivables and inventory, and establish weekly cash-flow forecasting.',
        supportingSourceRef: primaryRef,
      });
    } else if (k.patPaise > 0 && cfoTotal < k.patPaise * 0.65) {
      cfoSuggestions.push({
        id: 'CFO_SUG_PAT_VS_OCF',
        category: 'Profit vs Cash Flow',
        factualObservation: `Operating Cash Flow (${formatINRCompact(
          cfoTotal
        )}) trails accounting Net Profit (${formatINRCompact(k.patPaise)}).`,
        actualFinancialValue: `OCF: ${formatINR(cfoTotal)}`,
        comparativeValue: `PAT: ${formatINR(k.patPaise)}`,
        calculatedVariance: `Unconverted Profit: ${formatINR(k.patPaise - cfoTotal)}`,
        quantifiedFinancialImpact: `${formatINR(
          k.patPaise - cfoTotal
        )} of accounting earnings tied up in working capital`,
        riskClassification: 'MODERATE RISK',
        suggestedManagementAction:
          'Reconcile accrual revenue against debtor collections to verify why accounting profit is not converting into cash.',
        supportingSourceRef: primaryRef,
      });
    }
  }

  // Rule 6: Working Capital / DSO stretch
  if (dsoDays !== null && dsoDays > 45) {
    const excessDays = dsoDays - 45;
    const cashLockedPaise = Math.round((excessDays / 365) * k.netRevenuePaise);
    cfoSuggestions.push({
      id: 'CFO_SUG_WC_DSO',
      category: 'Working Capital',
      factualObservation: `Days Sales Outstanding (DSO) is ${dsoDays} days (Cash Conversion Cycle: ${
        cashConversionCycleDays ?? 'N/A'
      } days).`,
      actualFinancialValue: `${dsoDays} Days (Receivables: ${formatINRCompact(
        k.tradeReceivablesPaise
      )})`,
      comparativeValue: 'Standard Credit Policy: 30–45 Days',
      calculatedVariance: `+${excessDays} Days above 45-day threshold`,
      quantifiedFinancialImpact: `Approximately ${formatINR(
        cashLockedPaise
      )} of liquidity locked in excess debtor days`,
      riskClassification: dsoDays > 75 ? 'HIGH RISK' : 'MODERATE RISK',
      suggestedManagementAction:
        'Enforce credit limits on overdue customers and initiate structured follow-up on invoices aged beyond 60 days.',
      supportingSourceRef: primaryRef,
    });
  }

  // Rule 7: Debt & Interest Coverage
  const intCovObj = existingRatios.find((r) => r.id === 'RATIO_INTEREST_COVERAGE');
  if (intCovObj && intCovObj.currentValue !== null && intCovObj.currentValue < 2.5) {
    cfoSuggestions.push({
      id: 'CFO_SUG_INT_COV',
      category: 'Interest Coverage',
      factualObservation: `Interest Coverage Ratio is ${intCovObj.currentValue.toFixed(2)}x.`,
      actualFinancialValue: `${intCovObj.currentValue.toFixed(2)}x (EBIT: ${formatINRCompact(
        k.ebitdaPaise - k.depreciationPaise
      )})`,
      comparativeValue: `Finance Costs: ${formatINRCompact(k.financeCostsPaise)} (Prudent ≥ 3.0x)`,
      calculatedVariance: `${(intCovObj.currentValue - 3.0).toFixed(2)}x vs benchmark`,
      quantifiedFinancialImpact: `Annualized interest burden of ${formatINR(k.financeCostsPaise)}`,
      riskClassification: intCovObj.currentValue < 1.5 ? 'HIGH RISK' : 'MODERATE RISK',
      suggestedManagementAction:
        'Review high-cost working capital credit facilities (OD/CC) and evaluate debt refinancing or working capital release.',
      supportingSourceRef: primaryRef,
    });
  }

  // Rule 8: Balance Sheet Reconciliation Difference
  if (!isBsReconciled && model.sourceFiles.length > 0 && k.totalAssetsPaise > 0) {
    cfoSuggestions.push({
      id: 'CFO_SUG_BS_RECON_DIFF',
      category: 'Balance Sheet Reconciliation',
      factualObservation: `Balance Sheet equation shows a difference of ${formatINR(
        Math.abs(bsDiffPaise)
      )} between Total Assets and Total Equity + Liabilities.`,
      actualFinancialValue: `Total Assets: ${formatINR(k.totalAssetsPaise)}`,
      comparativeValue: `Equity + Liabilities: ${formatINR(totalEquityAndLiabilitiesPaise)}`,
      calculatedVariance: `Difference: ${formatINR(bsDiffPaise)}`,
      quantifiedFinancialImpact: `Unreconciled balance of ${formatINR(Math.abs(bsDiffPaise))}`,
      riskClassification: 'HIGH RISK',
      suggestedManagementAction:
        'Check for unmapped Trial Balance accounts, Suspense ledgers, or opening stock adjustments not reflected in Reserves.',
      supportingSourceRef: primaryRef,
    });
  }

  return {
    cfoKpis,
    cashFlowAnalysis,
    workingCapital,
    cashFlowRatios,
    cfoSuggestions,
    subtotalReconciliations,
    balanceSheetVerificationRows,
    balanceSheetReconciliation: {
      totalAssetsPaise: k.totalAssetsPaise,
      totalEquityPaise: k.totalEquityPaise,
      totalLiabilitiesPaise: totalLiabilitiesOnlyPaise,
      totalEquityAndLiabilitiesPaise,
      differencePaise: bsDiffPaise,
      isReconciled: isStrictBsReconciled,
      priorTotalAssetsPaise,
      priorTotalEquityAndLiabilitiesPaise: getLinePrev('TOTAL_EQUITY_AND_LIABILITIES'),
      priorDifferencePaise:
        priorTotalAssetsPaise !== null && getLinePrev('TOTAL_EQUITY_AND_LIABILITIES') !== null
          ? priorTotalAssetsPaise - (getLinePrev('TOTAL_EQUITY_AND_LIABILITIES') ?? 0)
          : null,
      isPriorReconciled:
        priorTotalAssetsPaise !== null &&
        getLinePrev('TOTAL_EQUITY_AND_LIABILITIES') !== null &&
        Math.abs(priorTotalAssetsPaise - (getLinePrev('TOTAL_EQUITY_AND_LIABILITIES') ?? 0)) <=
          policy.roundingTolerancePaise,
      reportedTotalAssetsPaise:
        (model.reportedSubtotals || []).find((s) => s.normalizedLabel === 'total assets')
          ?.reportedCurrentPaise ?? null,
      reportedTotalEquityAndLiabilitiesPaise:
        (model.reportedSubtotals || []).find(
          (s) =>
            s.normalizedLabel === 'total equity and liabilities' ||
            s.normalizedLabel === 'total equity & liabilities'
        )?.reportedCurrentPaise ?? null,
    },
    priorPeriodTotals: {
      hasPriorData,
      revenuePaise: priorRevenuePaise,
      grossProfitPaise: priorGrossProfitPaise,
      ebitdaPaise: priorEbitdaPaise,
      patPaise: priorPatPaise,
      totalAssetsPaise: priorTotalAssetsPaise,
      currentAssetsPaise: priorCurrentAssetsPaise,
      currentLiabilitiesPaise: priorCurrentLiabilitiesPaise,
      cashAndBankPaise: priorCashAndBankPaise,
      workingCapitalPaise: priorWorkingCapitalPaise,
    },
    budgetTotals: {
      hasBudgetData: hasStatementBudget,
      revenuePaise: hasStatementBudget ? sumBudgetSales : null,
      grossProfitPaise: budgetGrossProfitPaise,
      ebitdaPaise: budgetEbitdaPaise,
      patPaise: null,
    },
  };
}

// ============================================================================
// PART 11 — OPTIONAL FORECAST & SCENARIO ANALYSIS ENGINE
// Clearly labelled PROJECTED / SCENARIO BASED — Never mixed with actuals
// ============================================================================

export function computeScenarioProjections(
  statements: FinancialStatementBundle,
  cfoBundle: CfoAnalyticsBundle,
  assumptions: ScenarioAssumptions
): ScenarioCaseOutput[] {
  const k = statements.kpis;
  const monthCount = Math.max(1, statements.monthlyTrends.length || 1);

  // Baseline monthly run-rate from actual uploaded data
  const baseMonthlyRevPaise = Math.round(k.netRevenuePaise / monthCount);
  const baseMonthlyOpexPaise = Math.round(
    (k.employeeExpensesPaise + k.adminExpensesPaise + k.sellingExpensesPaise) / monthCount
  );
  const baseMonthlyDeprFinPaise = Math.round(
    (k.depreciationPaise + k.financeCostsPaise) / monthCount
  );
  const currentClosingCashPaise = k.cashPaise + k.bankPaise;

  const buildCase = (
    caseName: 'Base Case' | 'Optimistic Case' | 'Conservative Case',
    revDeltaBoost: number,
    gpMarginBoost: number,
    opexBoost: number
  ): ScenarioCaseOutput => {
    const effectiveRevGrowth = (assumptions.revenueGrowthPercent + revDeltaBoost) / 100;
    const effectiveGpMargin = Math.max(
      0,
      Math.min(95, assumptions.grossProfitMarginPercent + gpMarginBoost)
    ) / 100;
    const effectiveOpexGrowth = (assumptions.operatingExpenseChangePercent + opexBoost) / 100;

    const projRevPaise = Math.round(baseMonthlyRevPaise * (1 + effectiveRevGrowth));
    const projGpPaise = Math.round(projRevPaise * effectiveGpMargin);
    const projCogsPaise = projRevPaise - projGpPaise;
    const projOpexPaise = Math.round(baseMonthlyOpexPaise * (1 + effectiveOpexGrowth));
    const projEbitdaPaise = projGpPaise - projOpexPaise;
    const taxPaise = Math.round(assumptions.plannedTaxPaymentRupees * 100);
    const projNetProfitPaise = projEbitdaPaise - baseMonthlyDeprFinPaise - taxPaise;

    // Working capital & cash conversion effect based on collection/payment days
    const projReceivablesPaise = Math.round(
      projRevPaise * (assumptions.collectionPeriodDays / 30)
    );
    const projPayablesPaise = Math.round(projCogsPaise * (assumptions.paymentPeriodDays / 30));
    const wcDeltaPaise =
      projReceivablesPaise -
      projPayablesPaise -
      (Math.round(k.tradeReceivablesPaise / monthCount) -
        Math.round(k.tradePayablesPaise / monthCount));

    const capexPaise = Math.round(assumptions.plannedCapexRupees * 100);
    const loanRepayPaise = Math.round(assumptions.plannedLoanRepaymentRupees * 100);

    const projOperatingCfPaise = projEbitdaPaise - taxPaise - Math.round(wcDeltaPaise * 0.25);
    const projFreeCfPaise = projOperatingCfPaise - capexPaise;
    const projClosingCashPaise =
      currentClosingCashPaise + projFreeCfPaise - loanRepayPaise;
    const projWorkingCapitalPaise =
      k.workingCapitalPaise + projNetProfitPaise - capexPaise;

    return {
      caseName,
      projectedRevenuePaise: projRevPaise,
      projectedGrossProfitPaise: projGpPaise,
      projectedOpexPaise: projOpexPaise,
      projectedEbitdaPaise: projEbitdaPaise,
      projectedNetProfitPaise: projNetProfitPaise,
      projectedOperatingCashFlowPaise: projOperatingCfPaise,
      projectedFreeCashFlowPaise: projFreeCfPaise,
      projectedClosingCashPaise: projClosingCashPaise,
      projectedWorkingCapitalPaise: projWorkingCapitalPaise,
    };
  };

  return [
    buildCase('Base Case', 0, 0, 0),
    buildCase('Optimistic Case', 5, 2.5, -2),
    buildCase('Conservative Case', -5, -2.5, 3),
  ];
}
