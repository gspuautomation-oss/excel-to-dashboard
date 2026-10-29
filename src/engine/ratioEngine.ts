import {
  ConsolidatedAccountingModel,
  FinancialRatioResult,
  formatINR,
  PolicySettings,
} from '../types/accounting';
import { FinancialStatementBundle } from './accountingEngine';

export function computeFinancialRatios(
  model: ConsolidatedAccountingModel,
  statements: FinancialStatementBundle,
  policy: PolicySettings,
  priorStatements?: FinancialStatementBundle
): FinancialRatioResult[] {
  const k = statements.kpis;
  const pk = priorStatements?.kpis;

  // Compute period-average balances where opening balances exist
  let openingInventoryPaise = 0;
  for (const st of model.stockItems) {
    openingInventoryPaise += st.openingValuePaise;
  }
  let openingReceivablesPaise = 0;
  let openingPayablesPaise = 0;
  let openingEquityPaise = 0;
  let openingDebtPaise = 0;

  for (const [, sum] of model.ledgerSummaries.entries()) {
    const cls = sum.ledger.classification;
    if (cls === 'TradeReceivables') openingReceivablesPaise += Math.max(0, sum.openingBalancePaise);
    if (cls === 'TradePayables') openingPayablesPaise += Math.max(0, -sum.openingBalancePaise);
    if (cls === 'Equity') openingEquityPaise += Math.max(0, -sum.openingBalancePaise);
    if (cls === 'LongTermBorrowings' || cls === 'ShortTermBorrowings') {
      openingDebtPaise += Math.max(0, -sum.openingBalancePaise);
    }
  }

  const avgInventoryPaise =
    openingInventoryPaise > 0
      ? Math.round((openingInventoryPaise + k.inventoryPaise) / 2)
      : k.inventoryPaise;

  const avgReceivablesPaise =
    openingReceivablesPaise > 0
      ? Math.round((openingReceivablesPaise + k.tradeReceivablesPaise) / 2)
      : k.tradeReceivablesPaise;

  const avgPayablesPaise =
    openingPayablesPaise > 0
      ? Math.round((openingPayablesPaise + k.tradePayablesPaise) / 2)
      : k.tradePayablesPaise;

  const avgEquityPaise =
    openingEquityPaise > 0
      ? Math.round((openingEquityPaise + k.totalEquityPaise) / 2)
      : k.totalEquityPaise;

  const totalDebtPaise = k.longTermBorrowingsPaise + k.shortTermBorrowingsPaise;
  const capitalEmployedPaise = k.totalAssetsPaise - k.totalCurrentLiabilitiesPaise;
  const openingCapEmployedPaise = openingEquityPaise + openingDebtPaise;
  const avgCapitalEmployedPaise =
    openingCapEmployedPaise > 0
      ? Math.round((openingCapEmployedPaise + capitalEmployedPaise) / 2)
      : capitalEmployedPaise;

  const daysInPeriod = Math.max(30, model.availableMonths.length * 30 || 365);

  const buildRatio = (
    id: string,
    category: FinancialRatioResult['category'],
    name: string,
    formula: string,
    numeratorLabel: string,
    denominatorLabel: string,
    numeratorPaise: number,
    denominatorPaise: number,
    unit: FinancialRatioResult['unit'],
    multiplier: number,
    accountMappings: string[],
    priorNumPaise?: number,
    priorDenPaise?: number,
    customDirectPaiseValue?: number
  ): FinancialRatioResult => {
    if (unit === 'CURRENCY' && customDirectPaiseValue !== undefined) {
      const prevVal = priorNumPaise !== undefined && priorDenPaise !== undefined
        ? (priorNumPaise - priorDenPaise) / 100
        : null;
      return {
        id,
        category,
        name,
        formula,
        numeratorLabel,
        denominatorLabel,
        numeratorPaise,
        denominatorPaise,
        currentValue: customDirectPaiseValue / 100,
        previousValue: prevVal,
        unit,
        formattedCurrent: formatINR(customDirectPaiseValue),
        formattedPrevious:
          priorNumPaise !== undefined && priorDenPaise !== undefined
            ? formatINR(priorNumPaise - priorDenPaise)
            : 'N/A',
        status: 'COMPUTED',
        explanation: `${numeratorLabel} (${formatINR(numeratorPaise)}) minus ${denominatorLabel} (${formatINR(denominatorPaise)}).`,
        accountMappings,
        isMaterialChange: false,
      };
    }

    if (denominatorPaise === 0) {
      return {
        id,
        category,
        name,
        formula,
        numeratorLabel,
        denominatorLabel,
        numeratorPaise,
        denominatorPaise,
        currentValue: null,
        previousValue: null,
        unit,
        formattedCurrent: 'N/A (Zero Denominator)',
        formattedPrevious: 'N/A',
        status: 'ZERO_DENOMINATOR',
        explanation: `Cannot divide by zero: ${denominatorLabel} is ₹0 in the uploaded records.`,
        accountMappings,
        isMaterialChange: false,
      };
    }

    if (denominatorPaise < 0) {
      return {
        id,
        category,
        name,
        formula,
        numeratorLabel,
        denominatorLabel,
        numeratorPaise,
        denominatorPaise,
        currentValue: null,
        previousValue: null,
        unit,
        formattedCurrent: 'N/A (Negative Base)',
        formattedPrevious: 'N/A',
        status: 'NEGATIVE_DENOMINATOR',
        explanation: `${denominatorLabel} is negative (${formatINR(denominatorPaise)}), making a standard ratio misleading.`,
        accountMappings,
        isMaterialChange: false,
      };
    }

    const rawVal = (numeratorPaise / denominatorPaise) * multiplier;
    const prevRawVal =
      priorNumPaise !== undefined && priorDenPaise !== undefined && priorDenPaise > 0
        ? (priorNumPaise / priorDenPaise) * multiplier
        : null;

    const formatVal = (v: number | null) => {
      if (v === null) return 'N/A';
      if (unit === 'RATIO') return `${v.toFixed(2)}x`;
      if (unit === 'PERCENT') return `${v.toFixed(2)}%`;
      if (unit === 'DAYS') return `${Math.round(v)} Days`;
      return formatINR(Math.round(v * 100));
    };

    const isMaterialChange =
      prevRawVal !== null && Math.abs(rawVal - prevRawVal) >= (unit === 'PERCENT' ? 2.0 : 0.25);

    return {
      id,
      category,
      name,
      formula,
      numeratorLabel,
      denominatorLabel,
      numeratorPaise,
      denominatorPaise,
      currentValue: rawVal,
      previousValue: prevRawVal,
      unit,
      formattedCurrent: formatVal(rawVal),
      formattedPrevious: formatVal(prevRawVal),
      status: 'COMPUTED',
      explanation: `${numeratorLabel} (${formatINR(numeratorPaise)}) ÷ ${denominatorLabel} (${formatINR(denominatorPaise)}).`,
      accountMappings,
      isMaterialChange,
    };
  };

  const quickAssetsPaise =
    k.totalCurrentAssetsPaise - k.inventoryPaise - k.otherCurrentAssetsPaise;
  const cashEquivalentsPaise = k.cashPaise + k.bankPaise;

  const ratios: FinancialRatioResult[] = [
    // 6.1 LIQUIDITY RATIOS
    buildRatio(
      'CURRENT_RATIO',
      'LIQUIDITY',
      'Current Ratio',
      'Current Assets / Current Liabilities',
      'Current Assets',
      'Current Liabilities',
      k.totalCurrentAssetsPaise,
      k.totalCurrentLiabilitiesPaise,
      'RATIO',
      1,
      ['Inventories', 'Trade Receivables', 'Cash & Bank', 'Other Current Assets', 'Short-Term Borrowings', 'Trade Payables', 'Statutory & Other Current Liabilities'],
      pk?.totalCurrentAssetsPaise,
      pk?.totalCurrentLiabilitiesPaise
    ),
    buildRatio(
      'QUICK_RATIO',
      'LIQUIDITY',
      'Quick Ratio (Acid-Test)',
      '(Current Assets − Inventory − Other Non-Quick Current Assets) / Current Liabilities',
      'Quick Assets (Receivables + Cash & Bank)',
      'Current Liabilities',
      quickAssetsPaise,
      k.totalCurrentLiabilitiesPaise,
      'RATIO',
      1,
      ['Trade Receivables', 'Cash-in-Hand', 'Bank Balances', 'Current Liabilities'],
      pk ? pk.totalCurrentAssetsPaise - pk.inventoryPaise - pk.otherCurrentAssetsPaise : undefined,
      pk?.totalCurrentLiabilitiesPaise
    ),
    buildRatio(
      'CASH_RATIO',
      'LIQUIDITY',
      'Cash Ratio',
      'Cash and Cash Equivalents / Current Liabilities',
      'Cash & Bank Balances',
      'Current Liabilities',
      cashEquivalentsPaise,
      k.totalCurrentLiabilitiesPaise,
      'RATIO',
      1,
      ['Cash-in-Hand', 'Bank Accounts', 'Current Liabilities'],
      pk ? pk.cashPaise + pk.bankPaise : undefined,
      pk?.totalCurrentLiabilitiesPaise
    ),
    buildRatio(
      'WORKING_CAPITAL',
      'LIQUIDITY',
      'Net Working Capital',
      'Current Assets − Current Liabilities',
      'Current Assets',
      'Current Liabilities',
      k.totalCurrentAssetsPaise,
      k.totalCurrentLiabilitiesPaise,
      'CURRENCY',
      1,
      ['Total Current Assets', 'Total Current Liabilities'],
      pk?.totalCurrentAssetsPaise,
      pk?.totalCurrentLiabilitiesPaise,
      k.workingCapitalPaise
    ),

    // 6.2 PROFITABILITY RATIOS
    buildRatio(
      'GROSS_MARGIN',
      'PROFITABILITY',
      'Gross Profit Margin',
      'Gross Profit / Net Revenue × 100',
      'Gross Profit',
      'Net Revenue',
      k.grossProfitPaise,
      k.netRevenuePaise,
      'PERCENT',
      100,
      ['Sales', 'Sales Returns', 'Purchases', 'Direct Expenses'],
      pk?.grossProfitPaise,
      pk?.netRevenuePaise
    ),
    buildRatio(
      'OPERATING_MARGIN',
      'PROFITABILITY',
      'Operating Profit Margin (EBIT Margin)',
      'Operating Profit (EBIT) / Net Revenue × 100',
      'Operating Profit (EBIT)',
      'Net Revenue',
      k.ebitPaise,
      k.netRevenuePaise,
      'PERCENT',
      100,
      ['Gross Profit', 'Employee Expenses', 'Administrative Expenses', 'Selling Expenses', 'Depreciation'],
      pk?.ebitPaise,
      pk?.netRevenuePaise
    ),
    buildRatio(
      'NET_MARGIN',
      'PROFITABILITY',
      'Net Profit Margin',
      'Profit After Tax / Net Revenue × 100',
      'Profit After Tax (PAT)',
      'Net Revenue',
      k.patPaise,
      k.netRevenuePaise,
      'PERCENT',
      100,
      ['Profit After Tax', 'Net Sales'],
      pk?.patPaise,
      pk?.netRevenuePaise
    ),
    buildRatio(
      'ROE',
      'PROFITABILITY',
      'Return on Equity (ROE)',
      'Profit After Tax / Average Equity × 100',
      'Profit After Tax',
      'Average Shareholders’ Equity',
      k.patPaise,
      avgEquityPaise,
      'PERCENT',
      100,
      ['Profit After Tax', 'Capital Account', 'Reserves & Surplus'],
      pk?.patPaise,
      pk?.totalEquityPaise
    ),
    buildRatio(
      'ROCE',
      'PROFITABILITY',
      'Return on Capital Employed (ROCE)',
      'EBIT / Average Capital Employed × 100',
      'EBIT',
      'Average Capital Employed (Total Assets − Current Liabilities)',
      k.ebitPaise,
      avgCapitalEmployedPaise,
      'PERCENT',
      100,
      ['EBIT', 'Total Assets', 'Current Liabilities'],
      pk?.ebitPaise,
      pk ? pk.totalAssetsPaise - pk.totalCurrentLiabilitiesPaise : undefined
    ),

    // 6.3 EFFICIENCY RATIOS
    buildRatio(
      'INVENTORY_TURNOVER',
      'EFFICIENCY',
      'Inventory Turnover',
      'Cost of Goods Sold / Average Inventory',
      'Cost of Goods Sold (COGS)',
      'Average Inventory',
      k.cogsPaise,
      avgInventoryPaise,
      'RATIO',
      1,
      ['COGS', 'Opening Stock', 'Closing Stock']
    ),
    buildRatio(
      'RECEIVABLES_TURNOVER',
      'EFFICIENCY',
      'Receivables Turnover',
      'Net Credit Sales / Average Trade Receivables',
      'Net Sales',
      'Average Trade Receivables',
      k.netRevenuePaise,
      avgReceivablesPaise,
      'RATIO',
      1,
      ['Net Sales', 'Sundry Debtors']
    ),
    buildRatio(
      'PAYABLES_TURNOVER',
      'EFFICIENCY',
      'Payables Turnover',
      'Credit Purchases / Average Trade Payables',
      'Net Purchases',
      'Average Trade Payables',
      k.netPurchasesPaise,
      avgPayablesPaise,
      'RATIO',
      1,
      ['Net Purchases', 'Sundry Creditors']
    ),
    buildRatio(
      'DEBTOR_COLLECTION_PERIOD',
      'EFFICIENCY',
      'Debtor Collection Period (DSO)',
      'Average Trade Receivables / Net Credit Sales × Days in Period',
      'Average Trade Receivables',
      'Net Credit Sales',
      avgReceivablesPaise,
      k.netRevenuePaise,
      'DAYS',
      daysInPeriod,
      ['Sundry Debtors', 'Net Sales']
    ),
    buildRatio(
      'CREDITOR_PAYMENT_PERIOD',
      'EFFICIENCY',
      'Creditor Payment Period (DPO)',
      'Average Trade Payables / Credit Purchases × Days in Period',
      'Average Trade Payables',
      'Credit Purchases',
      avgPayablesPaise,
      k.netPurchasesPaise,
      'DAYS',
      daysInPeriod,
      ['Sundry Creditors', 'Net Purchases']
    ),

    // 6.4 SOLVENCY RATIOS
    buildRatio(
      'DEBT_EQUITY',
      'SOLVENCY',
      'Debt-to-Equity Ratio',
      'Interest-Bearing Borrowings / Total Equity',
      'Interest-Bearing Borrowings (Long-Term + Short-Term)',
      'Total Shareholders’ Equity',
      totalDebtPaise,
      k.totalEquityPaise,
      'RATIO',
      1,
      ['Long-Term Borrowings', 'Short-Term Borrowings / Bank OD', 'Equity']
    ),
    buildRatio(
      'INTEREST_COVERAGE',
      'SOLVENCY',
      'Interest Coverage Ratio',
      'EBIT / Finance Costs',
      'Operating Profit (EBIT)',
      'Finance Costs & Interest',
      k.ebitPaise,
      k.financeCostsPaise,
      'RATIO',
      1,
      ['EBIT', 'Finance Costs']
    ),
  ];

  // DSCR Guard (Module 6.4: Calculate ONLY when debt-service schedules are supplied)
  if (policy.debtServiceAnnualPaise && policy.debtServiceAnnualPaise > 0) {
    const cashAvailableForDebtServicePaise = k.ebitdaPaise - k.taxExpensePaise;
    ratios.push(
      buildRatio(
        'DSCR',
        'SOLVENCY',
        'Debt Service Coverage Ratio (DSCR)',
        '(EBITDA − Cash Taxes) / Scheduled Principal & Interest Repayments',
        'Cash Available for Debt Service (EBITDA − Tax)',
        'Scheduled Annual Debt Service',
        cashAvailableForDebtServicePaise,
        policy.debtServiceAnnualPaise,
        'RATIO',
        1,
        ['EBITDA', 'Tax Expense', 'Configured Debt-Service Schedule']
      )
    );
  } else {
    ratios.push({
      id: 'DSCR',
      category: 'SOLVENCY',
      name: 'Debt Service Coverage Ratio (DSCR)',
      formula: 'Cash Available for Debt Service / Scheduled Principal & Interest',
      numeratorLabel: 'EBITDA − Tax Expense',
      denominatorLabel: 'Scheduled Debt Service (Principal + Interest)',
      numeratorPaise: k.ebitdaPaise - k.taxExpensePaise,
      denominatorPaise: 0,
      currentValue: null,
      previousValue: null,
      unit: 'RATIO',
      formattedCurrent: 'Requires Schedule',
      formattedPrevious: 'N/A',
      status: 'INSUFFICIENT_DATA',
      explanation:
        'DSCR requires an explicit principal + interest repayment schedule. Per accounting policy, DSCR is never fabricated from balance sheet totals alone. You may enter the annual debt-service amount in Policy Settings above.',
      accountMappings: ['EBITDA', 'Tax Expense', 'Debt-Service Schedule'],
      isMaterialChange: false,
    });
  }

  return ratios;
}
