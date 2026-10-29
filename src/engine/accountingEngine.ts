import {
  AccountClassification,
  computeSafeVariance,
  ConsolidatedAccountingModel,
  StatementLineItem,
} from '../types/accounting';

export interface FinancialStatementBundle {
  periodLabel: string;
  comparativePeriodLabel?: string;
  isIncompleteComparisonWarning: boolean;
  pnlLines: StatementLineItem[];
  bsAssetLines: StatementLineItem[];
  bsLiabilityEquityLines: StatementLineItem[];
  kpis: {
    grossSalesPaise: number;
    salesReturnsPaise: number;
    netRevenuePaise: number;
    otherOperatingIncomePaise: number;
    otherIncomePaise: number;
    totalIncomePaise: number;
    purchasesPaise: number;
    purchaseReturnsPaise: number;
    netPurchasesPaise: number;
    directExpensesPaise: number;
    cogsPaise: number;
    hasGrossProfitData: boolean;
    grossProfitPaise: number;
    employeeExpensesPaise: number;
    adminExpensesPaise: number;
    sellingExpensesPaise: number;
    ebitdaPaise: number;
    operatingEbitdaPaise: number;
    depreciationPaise: number;
    ebitPaise: number;
    financeCostsPaise: number;
    totalExpensesPaise: number;
    profitBeforePriorPeriodPaise: number;
    priorPeriodExpensesPaise: number;
    pbtPaise: number;
    currentTaxPaise: number;
    deferredTaxPaise: number;
    taxExpensePaise: number;
    patPaise: number;
    // Balance Sheet KPIs
    fixedAssetsPaise: number;
    investmentsPaise: number;
    otherNonCurrentAssetsPaise: number;
    totalNonCurrentAssetsPaise: number;
    inventoryPaise: number;
    tradeReceivablesPaise: number;
    cashPaise: number;
    bankPaise: number;
    otherCurrentAssetsPaise: number;
    totalCurrentAssetsPaise: number;
    totalAssetsPaise: number;
    equityBasePaise: number;
    totalEquityPaise: number;
    longTermBorrowingsPaise: number;
    shortTermBorrowingsPaise: number;
    tradePayablesPaise: number;
    taxLiabilitiesPaise: number;
    otherCurrentLiabilitiesPaise: number;
    totalCurrentLiabilitiesPaise: number;
    totalEquityAndLiabilitiesPaise: number;
    workingCapitalPaise: number;
    suspensePaise: number;
    unclassifiedPaise: number;
  };
  monthlyTrends: Array<{
    month: string;
    revenuePaise: number;
    purchasesPaise: number;
    grossProfitPaise: number;
    expensesPaise: number;
    netProfitPaise: number;
    collectionsPaise: number;
    paymentsPaise: number;
    workingCapitalPaise: number;
  }>;
}

interface ClassificationSumResult {
  amountPaise: number;
  priorAmountPaise?: number;
  hasPrior: boolean;
  ledgerNames: string[];
  cellReferences: string[];
  priorCellReferences: string[];
  worksheetNames: string[];
}

/**
 * Computes net movement for P&L or BS classifications within an optional month filter (or full dataset).
 * Credit-nature items (Revenue/Liabilities/Equity) are returned as positive Paise when Credit > Debit.
 * Debit-nature items (Expenses/Assets) are returned as positive Paise when Debit > Credit.
 */
function sumClassificationsForMonths(
  model: ConsolidatedAccountingModel,
  classifications: AccountClassification[],
  nature: 'CREDIT_POSITIVE' | 'DEBIT_POSITIVE',
  allowedMonths?: Set<string>,
  useClosingBalance = false,
  requiredSourceType?: 'PNL' | 'BALANCE_SHEET'
): ClassificationSumResult {
  let totalPaise = 0;
  let priorTotalPaise = 0;
  let hasPrior = false;
  const ledgerNames: string[] = [];
  const cellReferences: string[] = [];
  const priorCellReferences: string[] = [];
  const worksheetNames: string[] = [];

  for (const [, summary] of model.ledgerSummaries.entries()) {
    if (!classifications.includes(summary.ledger.classification)) continue;
    if (requiredSourceType && summary.ledger.statementSourceType !== requiredSourceType) continue;

    let signedPaise = 0;
    if (useClosingBalance || !model.hasCompleteVoucherData || summary.transactionCount === 0) {
      signedPaise = summary.effectiveClosingPaise;
    } else if (allowedMonths && allowedMonths.size > 0) {
      for (const ym of allowedMonths) {
        const m = summary.monthlyNetPaise[ym];
        if (m) signedPaise += m.netPaise;
      }
    } else {
      signedPaise = summary.periodDebitPaise - summary.periodCreditPaise;
    }

    const hasPriorValue =
      !allowedMonths &&
      (summary.ledger.previousPeriodSignedPaise !== undefined ||
        summary.ledger.previousPeriodPaise !== undefined);

    if (signedPaise !== 0 || hasPriorValue) {
      ledgerNames.push(summary.ledger.name);
      if (summary.ledger.cellReference && !cellReferences.includes(summary.ledger.cellReference)) {
        cellReferences.push(summary.ledger.cellReference);
      }
      if (
        summary.ledger.priorCellReference &&
        !priorCellReferences.includes(summary.ledger.priorCellReference)
      ) {
        priorCellReferences.push(summary.ledger.priorCellReference);
      }
      if (summary.ledger.worksheetName && !worksheetNames.includes(summary.ledger.worksheetName)) {
        worksheetNames.push(summary.ledger.worksheetName);
      }
    }

    if (nature === 'CREDIT_POSITIVE') {
      totalPaise += -signedPaise;
    } else {
      totalPaise += signedPaise;
    }

    if (hasPriorValue) {
      hasPrior = true;
      if (summary.ledger.previousPeriodSignedPaise !== undefined) {
        priorTotalPaise +=
          nature === 'CREDIT_POSITIVE'
            ? -summary.ledger.previousPeriodSignedPaise
            : summary.ledger.previousPeriodSignedPaise;
      } else if (summary.ledger.previousPeriodPaise !== undefined) {
        priorTotalPaise += summary.ledger.previousPeriodPaise;
      }
    }
  }

  return {
    amountPaise: totalPaise,
    priorAmountPaise: hasPrior ? priorTotalPaise : undefined,
    hasPrior,
    ledgerNames,
    cellReferences,
    priorCellReferences,
    worksheetNames,
  };
}

export function computeFinancialStatements(
  model: ConsolidatedAccountingModel,
  currentMonths?: string[],
  previousMonths?: string[],
  periodLabel = 'Consolidated Financial Period',
  comparativePeriodLabel?: string
): FinancialStatementBundle {
  const currSet = currentMonths && currentMonths.length > 0 ? new Set(currentMonths) : undefined;
  const prevSet = previousMonths && previousMonths.length > 0 ? new Set(previousMonths) : undefined;

  const isIncompleteComparisonWarning = Boolean(
    currSet && prevSet && currSet.size !== prevSet.size
  );

  // Stock opening & closing values if stockItems exist
  let openingStockPaise = 0;
  let closingStockPaise = 0;
  for (const st of model.stockItems) {
    openingStockPaise += st.openingValuePaise;
    closingStockPaise += st.closingValuePaise;
  }

  const pnlSourceFilter: 'PNL' | undefined = model.hasProfitLossStatement ? 'PNL' : undefined;
  const bsSourceFilter: 'BALANCE_SHEET' | undefined = model.hasBalanceSheetStatement
    ? 'BALANCE_SHEET'
    : undefined;

  const calcPnlMetrics = (mSet?: Set<string>) => {
    const sales = sumClassificationsForMonths(model, ['Sales'], 'CREDIT_POSITIVE', mSet, false, pnlSourceFilter);
    const salesRet = sumClassificationsForMonths(model, ['SalesReturn'], 'DEBIT_POSITIVE', mSet, false, pnlSourceFilter);
    const netRevenuePaise = sales.amountPaise - salesRet.amountPaise;
    const hasPriorRev = sales.hasPrior || salesRet.hasPrior;
    const priorNetRevenuePaise = hasPriorRev
      ? (sales.priorAmountPaise || 0) - (salesRet.priorAmountPaise || 0)
      : undefined;

    const otherOpInc = sumClassificationsForMonths(model, ['OtherOperatingIncome'], 'CREDIT_POSITIVE', mSet, false, pnlSourceFilter);
    const otherInc = sumClassificationsForMonths(model, ['OtherIncome'], 'CREDIT_POSITIVE', mSet, false, pnlSourceFilter);

    const totalIncomePaise =
      netRevenuePaise + otherOpInc.amountPaise + otherInc.amountPaise;
    const hasPriorTotalInc = hasPriorRev || otherOpInc.hasPrior || otherInc.hasPrior;
    const priorTotalIncomePaise = hasPriorTotalInc
      ? (priorNetRevenuePaise || 0) +
        (otherOpInc.priorAmountPaise || 0) +
        (otherInc.priorAmountPaise || 0)
      : undefined;

    const purch = sumClassificationsForMonths(model, ['Purchases'], 'DEBIT_POSITIVE', mSet, false, pnlSourceFilter);
    const purchRet = sumClassificationsForMonths(model, ['PurchaseReturn'], 'CREDIT_POSITIVE', mSet, false, pnlSourceFilter);
    const netPurchasesPaise = purch.amountPaise - purchRet.amountPaise;
    const priorNetPurchasesPaise =
      purch.hasPrior || purchRet.hasPrior
        ? (purch.priorAmountPaise || 0) - (purchRet.priorAmountPaise || 0)
        : undefined;

    const directExp = sumClassificationsForMonths(model, ['DirectExpenses'], 'DEBIT_POSITIVE', mSet, false, pnlSourceFilter);

    // Adjust inventory variation only when computing from raw vouchers/TB without an explicit P&L statement
    const stockDeltaPaise =
      !mSet && !model.hasProfitLossStatement ? openingStockPaise - closingStockPaise : 0;
    const cogsPaise = netPurchasesPaise + stockDeltaPaise;
    const priorCogsPaise = priorNetPurchasesPaise;

    // Check whether Cost of Sales / Purchases / Direct Expenses or a reported Gross Profit line exists
    const reportedGpSubtotal = (model.reportedSubtotals || []).find(
      (s) =>
        s.statementType === 'PNL' &&
        (s.normalizedLabel === 'gross profit' || s.normalizedLabel === 'gross margin')
    );
    const hasGrossProfitData = Boolean(
      reportedGpSubtotal ||
        purch.ledgerNames.length > 0 ||
        purchRet.ledgerNames.length > 0 ||
        directExp.ledgerNames.length > 0 ||
        cogsPaise !== 0 ||
        directExp.amountPaise !== 0
    );

    const grossProfitPaise = reportedGpSubtotal
      ? reportedGpSubtotal.reportedCurrentPaise
      : hasGrossProfitData
      ? netRevenuePaise + otherOpInc.amountPaise - cogsPaise - directExp.amountPaise
      : 0;
    const priorGrossProfitPaise = reportedGpSubtotal?.reportedPreviousPaise !== undefined
      ? reportedGpSubtotal.reportedPreviousPaise
      : hasGrossProfitData && (hasPriorRev || purch.hasPrior || directExp.hasPrior)
      ? (priorNetRevenuePaise || 0) +
        (otherOpInc.priorAmountPaise || 0) -
        (priorCogsPaise || 0) -
        (directExp.priorAmountPaise || 0)
      : undefined;

    const empExp = sumClassificationsForMonths(model, ['EmployeeExpenses'], 'DEBIT_POSITIVE', mSet, false, pnlSourceFilter);
    const adminExp = sumClassificationsForMonths(model, ['AdminExpenses'], 'DEBIT_POSITIVE', mSet, false, pnlSourceFilter);
    const sellExp = sumClassificationsForMonths(model, ['SellingExpenses'], 'DEBIT_POSITIVE', mSet, false, pnlSourceFilter);

    const operatingEbitdaPaise =
      netRevenuePaise +
      otherOpInc.amountPaise -
      cogsPaise -
      directExp.amountPaise -
      empExp.amountPaise -
      adminExp.amountPaise -
      sellExp.amountPaise;

    const ebitdaPaise = operatingEbitdaPaise + otherInc.amountPaise;
    const hasPriorEbitda =
      hasPriorTotalInc || empExp.hasPrior || adminExp.hasPrior || sellExp.hasPrior;
    const priorEbitdaPaise = hasPriorEbitda
      ? (priorTotalIncomePaise || 0) -
        (priorCogsPaise || 0) -
        (directExp.priorAmountPaise || 0) -
        (empExp.priorAmountPaise || 0) -
        (adminExp.priorAmountPaise || 0) -
        (sellExp.priorAmountPaise || 0)
      : undefined;

    const depExp = sumClassificationsForMonths(model, ['Depreciation'], 'DEBIT_POSITIVE', mSet, false, pnlSourceFilter);
    const ebitPaise = ebitdaPaise - depExp.amountPaise;

    const finCost = sumClassificationsForMonths(model, ['FinanceCosts'], 'DEBIT_POSITIVE', mSet, false, pnlSourceFilter);
    const totalExpensesPaise =
      cogsPaise +
      directExp.amountPaise +
      empExp.amountPaise +
      adminExp.amountPaise +
      sellExp.amountPaise +
      depExp.amountPaise +
      finCost.amountPaise;
    const hasPriorTotalExp =
      purch.hasPrior ||
      directExp.hasPrior ||
      empExp.hasPrior ||
      adminExp.hasPrior ||
      sellExp.hasPrior ||
      depExp.hasPrior ||
      finCost.hasPrior;
    const priorTotalExpensesPaise = hasPriorTotalExp
      ? (priorCogsPaise || 0) +
        (directExp.priorAmountPaise || 0) +
        (empExp.priorAmountPaise || 0) +
        (adminExp.priorAmountPaise || 0) +
        (sellExp.priorAmountPaise || 0) +
        (depExp.priorAmountPaise || 0) +
        (finCost.priorAmountPaise || 0)
      : undefined;

    const profitBeforePriorPeriodPaise = totalIncomePaise - totalExpensesPaise;
    const priorProfitBeforePriorPeriodPaise =
      hasPriorTotalInc || hasPriorTotalExp
        ? (priorTotalIncomePaise || 0) - (priorTotalExpensesPaise || 0)
        : undefined;

    const priorPeriodExp = sumClassificationsForMonths(
      model,
      ['PriorPeriodExpenses'],
      'DEBIT_POSITIVE',
      mSet,
      false,
      pnlSourceFilter
    );
    const pbtPaise = profitBeforePriorPeriodPaise - priorPeriodExp.amountPaise;
    const priorPbtPaise =
      priorProfitBeforePriorPeriodPaise !== undefined || priorPeriodExp.hasPrior
        ? (priorProfitBeforePriorPeriodPaise || 0) - (priorPeriodExp.priorAmountPaise || 0)
        : undefined;

    const currentTax = sumClassificationsForMonths(model, ['CurrentTax'], 'DEBIT_POSITIVE', mSet, false, pnlSourceFilter);
    const deferredTax = sumClassificationsForMonths(model, ['DeferredTax'], 'DEBIT_POSITIVE', mSet, false, pnlSourceFilter);
    const generalTax = sumClassificationsForMonths(model, ['TaxExpense'], 'DEBIT_POSITIVE', mSet, false, pnlSourceFilter);

    const taxExpensePaise =
      currentTax.amountPaise + deferredTax.amountPaise + generalTax.amountPaise;
    const hasPriorTax = currentTax.hasPrior || deferredTax.hasPrior || generalTax.hasPrior;
    const priorTaxExpensePaise = hasPriorTax
      ? (currentTax.priorAmountPaise || 0) +
        (deferredTax.priorAmountPaise || 0) +
        (generalTax.priorAmountPaise || 0)
      : undefined;

    const patPaise = pbtPaise - taxExpensePaise;
    const priorPatPaise =
      priorPbtPaise !== undefined || hasPriorTax
        ? (priorPbtPaise || 0) - (priorTaxExpensePaise || 0)
        : undefined;

    return {
      sales,
      salesRet,
      netRevenuePaise,
      priorNetRevenuePaise,
      otherOpInc,
      otherInc,
      totalIncomePaise,
      priorTotalIncomePaise,
      purch,
      purchRet,
      netPurchasesPaise,
      priorNetPurchasesPaise,
      directExp,
      cogsPaise,
      priorCogsPaise,
      hasGrossProfitData,
      grossProfitPaise,
      priorGrossProfitPaise,
      empExp,
      adminExp,
      sellExp,
      operatingEbitdaPaise,
      ebitdaPaise,
      priorEbitdaPaise,
      depExp,
      ebitPaise,
      finCost,
      totalExpensesPaise,
      priorTotalExpensesPaise,
      profitBeforePriorPeriodPaise,
      priorProfitBeforePriorPeriodPaise,
      priorPeriodExp,
      pbtPaise,
      priorPbtPaise,
      currentTax,
      deferredTax,
      generalTax,
      taxExpensePaise,
      priorTaxExpensePaise,
      patPaise,
      priorPatPaise,
    };
  };

  const currPnl = calcPnlMetrics(currSet);
  const prevPnl = prevSet ? calcPnlMetrics(prevSet) : undefined;

  // Helper to locate reported subtotal cell reference if present
  const findSubtotalRef = (stmtType: 'PNL' | 'BALANCE_SHEET', keywords: string[]) => {
    const match = (model.reportedSubtotals || []).find(
      (s) =>
        s.statementType === stmtType &&
        keywords.some((kw) => s.normalizedLabel.includes(kw))
    );
    return match;
  };

  const makeLine = (
    code: string,
    label: string,
    classifications: AccountClassification[],
    currVal: number,
    prevVal: number | undefined,
    formulaExplanation: string,
    ledgerNames: string[],
    isSubtotal = false,
    options?: {
      cellReferences?: string[];
      priorCellReferences?: string[];
      worksheetNames?: string[];
      requiresMapping?: boolean;
      forceAvailable?: boolean;
    }
  ): StatementLineItem => {
    const variance =
      prevVal !== undefined && !options?.requiresMapping
        ? computeSafeVariance(currVal, prevVal)
        : undefined;
    const isAvailable = options?.requiresMapping
      ? false
      : options?.forceAvailable ??
        (isSubtotal || ledgerNames.length > 0 || currVal !== 0 || (prevVal ?? 0) !== 0);
    return {
      code,
      label,
      classificationFilter: classifications,
      amountPaise: currVal,
      previousAmountPaise: prevVal,
      variancePaise: variance?.diffPaise,
      variancePercent: variance?.percent,
      varianceLabel: variance?.label,
      formulaExplanation,
      ledgerNames,
      cellReferences: options?.cellReferences,
      priorCellReferences: options?.priorCellReferences,
      worksheetNames: options?.worksheetNames,
      isSubtotal,
      isAvailable,
      requiresMapping: options?.requiresMapping,
    };
  };

  const totalIncSub = findSubtotalRef('PNL', ['total income', 'total revenue']);
  const totalExpSub = findSubtotalRef('PNL', ['total expenses', 'total expense']);
  const pbtSub = findSubtotalRef('PNL', ['profit before tax', 'pbt']);
  const patSub = findSubtotalRef('PNL', ['profit after tax', 'profit for the period', 'profit for the year', 'net profit']);

  const pnlLines: StatementLineItem[] = [
    makeLine(
      'REVENUE_NET',
      'Revenue from Operations (Net)',
      ['Sales', 'SalesReturn'],
      currPnl.netRevenuePaise,
      prevPnl ? prevPnl.netRevenuePaise : currPnl.priorNetRevenuePaise,
      'Revenue from Operations = Sales / Operating Revenue − Sales Returns (excluding GST Output Tax)',
      [...currPnl.sales.ledgerNames, ...currPnl.salesRet.ledgerNames],
      false,
      {
        cellReferences: [...currPnl.sales.cellReferences, ...currPnl.salesRet.cellReferences],
        priorCellReferences: [
          ...currPnl.sales.priorCellReferences,
          ...currPnl.salesRet.priorCellReferences,
        ],
        worksheetNames: currPnl.sales.worksheetNames,
      }
    ),
    makeLine(
      'OTHER_OP_INCOME',
      'Other Operating Income',
      ['OtherOperatingIncome'],
      currPnl.otherOpInc.amountPaise,
      prevPnl ? prevPnl.otherOpInc.amountPaise : currPnl.otherOpInc.priorAmountPaise,
      'Sum of Direct / Service Operating Income ledgers',
      currPnl.otherOpInc.ledgerNames,
      false,
      {
        cellReferences: currPnl.otherOpInc.cellReferences,
        priorCellReferences: currPnl.otherOpInc.priorCellReferences,
        worksheetNames: currPnl.otherOpInc.worksheetNames,
      }
    ),
    makeLine(
      'OTHER_INCOME',
      'Other Income',
      ['OtherIncome'],
      currPnl.otherInc.amountPaise,
      prevPnl ? prevPnl.otherInc.amountPaise : currPnl.otherInc.priorAmountPaise,
      'Sum of Non-Operating / Other Income ledgers (Interest, Dividends, Misc Income)',
      currPnl.otherInc.ledgerNames,
      false,
      {
        cellReferences: currPnl.otherInc.cellReferences,
        priorCellReferences: currPnl.otherInc.priorCellReferences,
        worksheetNames: currPnl.otherInc.worksheetNames,
      }
    ),
    makeLine(
      'TOTAL_INCOME',
      'Total Income (I + II)',
      ['Sales', 'SalesReturn', 'OtherOperatingIncome', 'OtherIncome'],
      currPnl.totalIncomePaise,
      prevPnl ? prevPnl.totalIncomePaise : currPnl.priorTotalIncomePaise,
      'Total Income = Revenue from Operations + Other Operating Income + Other Income',
      [],
      true,
      {
        cellReferences: totalIncSub?.cellReference ? [totalIncSub.cellReference] : undefined,
        priorCellReferences: totalIncSub?.priorCellReference
          ? [totalIncSub.priorCellReference]
          : undefined,
      }
    ),
    makeLine(
      'COGS',
      'Cost of Goods Sold (Net Purchases & Stock Variation)',
      ['Purchases', 'PurchaseReturn', 'Inventory'],
      currPnl.cogsPaise,
      prevPnl ? prevPnl.cogsPaise : currPnl.priorCogsPaise,
      'COGS = (Purchases − Purchase Returns) + Opening Stock − Closing Stock',
      [...currPnl.purch.ledgerNames, ...currPnl.purchRet.ledgerNames],
      false,
      {
        cellReferences: [...currPnl.purch.cellReferences, ...currPnl.purchRet.cellReferences],
        priorCellReferences: [
          ...currPnl.purch.priorCellReferences,
          ...currPnl.purchRet.priorCellReferences,
        ],
        worksheetNames: currPnl.purch.worksheetNames,
      }
    ),
    makeLine(
      'DIRECT_EXP',
      'Direct Expenses',
      ['DirectExpenses'],
      currPnl.directExp.amountPaise,
      prevPnl ? prevPnl.directExp.amountPaise : currPnl.directExp.priorAmountPaise,
      'Sum of Direct & Manufacturing Expense ledgers',
      currPnl.directExp.ledgerNames,
      false,
      {
        cellReferences: currPnl.directExp.cellReferences,
        priorCellReferences: currPnl.directExp.priorCellReferences,
        worksheetNames: currPnl.directExp.worksheetNames,
      }
    ),
    makeLine(
      'GROSS_PROFIT',
      'Gross Profit',
      ['Sales', 'SalesReturn', 'OtherOperatingIncome', 'Purchases', 'PurchaseReturn', 'DirectExpenses'],
      currPnl.grossProfitPaise,
      prevPnl ? prevPnl.grossProfitPaise : currPnl.priorGrossProfitPaise,
      currPnl.hasGrossProfitData
        ? 'Gross Profit = Revenue from Operations − COGS − Direct Expenses'
        : 'Requires Mapping: Cost of Goods Sold / Direct Cost breakdown is not separately reported in the uploaded statement.',
      [],
      true,
      {
        requiresMapping: !currPnl.hasGrossProfitData,
      }
    ),
    makeLine(
      'EMPLOYEE_EXP',
      'Employee Benefits Expense',
      ['EmployeeExpenses'],
      currPnl.empExp.amountPaise,
      prevPnl ? prevPnl.empExp.amountPaise : currPnl.empExp.priorAmountPaise,
      'Sum of Salaries, Wages, PF, ESIC, Staff Welfare & Director Remuneration ledgers',
      currPnl.empExp.ledgerNames,
      false,
      {
        cellReferences: currPnl.empExp.cellReferences,
        priorCellReferences: currPnl.empExp.priorCellReferences,
        worksheetNames: currPnl.empExp.worksheetNames,
      }
    ),
    makeLine(
      'ADMIN_EXP',
      'Other / Administrative & General Expenses',
      ['AdminExpenses'],
      currPnl.adminExp.amountPaise,
      prevPnl ? prevPnl.adminExp.amountPaise : currPnl.adminExp.priorAmountPaise,
      'Sum of Other Expenses, Rent, Professional/Audit Fees, Office, Electricity & Insurance ledgers',
      currPnl.adminExp.ledgerNames,
      false,
      {
        cellReferences: currPnl.adminExp.cellReferences,
        priorCellReferences: currPnl.adminExp.priorCellReferences,
        worksheetNames: currPnl.adminExp.worksheetNames,
      }
    ),
    makeLine(
      'SELLING_EXP',
      'Selling & Distribution Expenses',
      ['SellingExpenses'],
      currPnl.sellExp.amountPaise,
      prevPnl ? prevPnl.sellExp.amountPaise : currPnl.sellExp.priorAmountPaise,
      'Sum of Marketing, Advertising, Sales Promotion & Commission ledgers',
      currPnl.sellExp.ledgerNames,
      false,
      {
        cellReferences: currPnl.sellExp.cellReferences,
        priorCellReferences: currPnl.sellExp.priorCellReferences,
        worksheetNames: currPnl.sellExp.worksheetNames,
      }
    ),
    makeLine(
      'EBITDA',
      'EBITDA (Earnings Before Interest, Tax, Depreciation & Amortisation)',
      [],
      currPnl.ebitdaPaise,
      prevPnl ? prevPnl.ebitdaPaise : currPnl.priorEbitdaPaise,
      'EBITDA = Total Income − COGS − Direct − Employee − Administrative − Selling Expenses (equals PBT + Finance Costs + Depreciation + Prior Period Items)',
      [],
      true
    ),
    makeLine(
      'DEPRECIATION',
      'Depreciation & Amortisation Expense',
      ['Depreciation'],
      currPnl.depExp.amountPaise,
      prevPnl ? prevPnl.depExp.amountPaise : currPnl.depExp.priorAmountPaise,
      'Sum of Depreciation & Amortisation ledgers',
      currPnl.depExp.ledgerNames,
      false,
      {
        cellReferences: currPnl.depExp.cellReferences,
        priorCellReferences: currPnl.depExp.priorCellReferences,
        worksheetNames: currPnl.depExp.worksheetNames,
      }
    ),
    makeLine(
      'FINANCE_COSTS',
      'Finance Costs',
      ['FinanceCosts'],
      currPnl.finCost.amountPaise,
      prevPnl ? prevPnl.finCost.amountPaise : currPnl.finCost.priorAmountPaise,
      'Sum of Interest on Borrowings, Bank Charges & Finance Cost ledgers',
      currPnl.finCost.ledgerNames,
      false,
      {
        cellReferences: currPnl.finCost.cellReferences,
        priorCellReferences: currPnl.finCost.priorCellReferences,
        worksheetNames: currPnl.finCost.worksheetNames,
      }
    ),
    makeLine(
      'TOTAL_EXPENSES',
      'Total Expenses (IV)',
      [
        'Purchases',
        'PurchaseReturn',
        'DirectExpenses',
        'EmployeeExpenses',
        'AdminExpenses',
        'SellingExpenses',
        'Depreciation',
        'FinanceCosts',
      ],
      currPnl.totalExpensesPaise,
      prevPnl ? prevPnl.totalExpensesPaise : currPnl.priorTotalExpensesPaise,
      'Total Expenses = Cost of Goods Sold + Employee Benefits + Finance Costs + Depreciation & Amortisation + Other Expenses',
      [],
      true,
      {
        cellReferences: totalExpSub?.cellReference ? [totalExpSub.cellReference] : undefined,
        priorCellReferences: totalExpSub?.priorCellReference
          ? [totalExpSub.priorCellReference]
          : undefined,
      }
    ),
    ...(currPnl.priorPeriodExp.amountPaise !== 0 ||
    (currPnl.priorPeriodExp.priorAmountPaise ?? 0) !== 0 ||
    currPnl.priorPeriodExp.ledgerNames.length > 0
      ? [
          makeLine(
            'PROFIT_BEFORE_PRIOR_PERIOD',
            'Profit Before Exceptional & Prior Period Items (III − IV)',
            [],
            currPnl.profitBeforePriorPeriodPaise,
            prevPnl
              ? prevPnl.profitBeforePriorPeriodPaise
              : currPnl.priorProfitBeforePriorPeriodPaise,
            'Profit Before Prior Period Items = Total Income − Total Expenses',
            [],
            true
          ),
          makeLine(
            'PRIOR_PERIOD_ITEMS',
            'Prior Period & Exceptional Items',
            ['PriorPeriodExpenses'],
            currPnl.priorPeriodExp.amountPaise,
            prevPnl
              ? prevPnl.priorPeriodExp.amountPaise
              : currPnl.priorPeriodExp.priorAmountPaise,
            'Prior Period Items / Exceptional Charges deducted before Tax',
            currPnl.priorPeriodExp.ledgerNames,
            false,
            {
              cellReferences: currPnl.priorPeriodExp.cellReferences,
              priorCellReferences: currPnl.priorPeriodExp.priorCellReferences,
              worksheetNames: currPnl.priorPeriodExp.worksheetNames,
            }
          ),
        ]
      : []),
    makeLine(
      'PBT',
      'Profit Before Tax (PBT)',
      [],
      currPnl.pbtPaise,
      prevPnl ? prevPnl.pbtPaise : currPnl.priorPbtPaise,
      'Profit Before Tax = Total Income − Total Expenses − Prior Period & Exceptional Items',
      [],
      true,
      {
        cellReferences: pbtSub?.cellReference ? [pbtSub.cellReference] : undefined,
        priorCellReferences: pbtSub?.priorCellReference
          ? [pbtSub.priorCellReference]
          : undefined,
      }
    ),
    ...(currPnl.currentTax.ledgerNames.length > 0 || currPnl.deferredTax.ledgerNames.length > 0
      ? [
          makeLine(
            'CURRENT_TAX',
            'Current Tax',
            ['CurrentTax'],
            currPnl.currentTax.amountPaise,
            prevPnl ? prevPnl.currentTax.amountPaise : currPnl.currentTax.priorAmountPaise,
            'Current Income Tax provision for the reporting period',
            currPnl.currentTax.ledgerNames,
            false,
            {
              cellReferences: currPnl.currentTax.cellReferences,
              priorCellReferences: currPnl.currentTax.priorCellReferences,
              worksheetNames: currPnl.currentTax.worksheetNames,
            }
          ),
          makeLine(
            'DEFERRED_TAX',
            'Deferred Tax (Credit / Charge)',
            ['DeferredTax'],
            currPnl.deferredTax.amountPaise,
            prevPnl ? prevPnl.deferredTax.amountPaise : currPnl.deferredTax.priorAmountPaise,
            'Deferred Tax expense (positive) or Deferred Tax credit (negative)',
            currPnl.deferredTax.ledgerNames,
            false,
            {
              cellReferences: currPnl.deferredTax.cellReferences,
              priorCellReferences: currPnl.deferredTax.priorCellReferences,
              worksheetNames: currPnl.deferredTax.worksheetNames,
            }
          ),
        ]
      : []),
    makeLine(
      'TAX_EXPENSE',
      'Net Tax Expense',
      ['CurrentTax', 'DeferredTax', 'TaxExpense'],
      currPnl.taxExpensePaise,
      prevPnl ? prevPnl.taxExpensePaise : currPnl.priorTaxExpensePaise,
      'Net Tax Expense = Current Tax + Deferred Tax (Credit reduces net tax)',
      [
        ...currPnl.currentTax.ledgerNames,
        ...currPnl.deferredTax.ledgerNames,
        ...currPnl.generalTax.ledgerNames,
      ],
      currPnl.currentTax.ledgerNames.length > 0 || currPnl.deferredTax.ledgerNames.length > 0,
      {
        cellReferences: [
          ...currPnl.currentTax.cellReferences,
          ...currPnl.deferredTax.cellReferences,
          ...currPnl.generalTax.cellReferences,
        ],
        priorCellReferences: [
          ...currPnl.currentTax.priorCellReferences,
          ...currPnl.deferredTax.priorCellReferences,
          ...currPnl.generalTax.priorCellReferences,
        ],
      }
    ),
    makeLine(
      'PAT',
      'Profit After Tax (PAT)',
      [],
      currPnl.patPaise,
      prevPnl ? prevPnl.patPaise : currPnl.priorPatPaise,
      'Profit After Tax = Profit Before Tax − Net Tax Expense',
      [],
      true,
      {
        cellReferences: patSub?.cellReference ? [patSub.cellReference] : undefined,
        priorCellReferences: patSub?.priorCellReference
          ? [patSub.priorCellReference]
          : undefined,
      }
    ),
  ];

  // Balance Sheet Computation (Closing cumulative position — strictly isolated to BALANCE_SHEET ledgers when a Balance Sheet is uploaded)
  const fixedAssets = sumClassificationsForMonths(
    model,
    ['FixedAssets'],
    'DEBIT_POSITIVE',
    undefined,
    true,
    bsSourceFilter
  );
  const investments = sumClassificationsForMonths(
    model,
    ['Investments'],
    'DEBIT_POSITIVE',
    undefined,
    true,
    bsSourceFilter
  );
  const nonCurrOther = sumClassificationsForMonths(
    model,
    ['NonCurrentAssets'],
    'DEBIT_POSITIVE',
    undefined,
    true,
    bsSourceFilter
  );
  const totalNonCurrentAssetsPaise =
    fixedAssets.amountPaise + investments.amountPaise + nonCurrOther.amountPaise;
  const hasPriorNonCurrAssets =
    fixedAssets.hasPrior || investments.hasPrior || nonCurrOther.hasPrior;
  const priorTotalNonCurrentAssetsPaise = hasPriorNonCurrAssets
    ? (fixedAssets.priorAmountPaise || 0) +
      (investments.priorAmountPaise || 0) +
      (nonCurrOther.priorAmountPaise || 0)
    : undefined;

  const invLedger = sumClassificationsForMonths(
    model,
    ['Inventory'],
    'DEBIT_POSITIVE',
    undefined,
    true,
    bsSourceFilter
  );
  const inventoryPaise =
    !model.hasBalanceSheetStatement && closingStockPaise > 0
      ? closingStockPaise
      : invLedger.amountPaise;
  const priorInventoryPaise = invLedger.priorAmountPaise;

  const receivables = sumClassificationsForMonths(
    model,
    ['TradeReceivables'],
    'DEBIT_POSITIVE',
    undefined,
    true,
    bsSourceFilter
  );
  const cash = sumClassificationsForMonths(
    model,
    ['CashAndEquivalents'],
    'DEBIT_POSITIVE',
    undefined,
    true,
    bsSourceFilter
  );
  const bank = sumClassificationsForMonths(
    model,
    ['BankBalances'],
    'DEBIT_POSITIVE',
    undefined,
    true,
    bsSourceFilter
  );
  const otherCurrAssets = sumClassificationsForMonths(
    model,
    ['OtherCurrentAssets'],
    'DEBIT_POSITIVE',
    undefined,
    true,
    bsSourceFilter
  );

  const totalCurrentAssetsPaise =
    inventoryPaise +
    receivables.amountPaise +
    cash.amountPaise +
    bank.amountPaise +
    otherCurrAssets.amountPaise;
  const hasPriorCurrAssets =
    invLedger.hasPrior ||
    receivables.hasPrior ||
    cash.hasPrior ||
    bank.hasPrior ||
    otherCurrAssets.hasPrior;
  const priorTotalCurrentAssetsPaise = hasPriorCurrAssets
    ? (priorInventoryPaise || 0) +
      (receivables.priorAmountPaise || 0) +
      (cash.priorAmountPaise || 0) +
      (bank.priorAmountPaise || 0) +
      (otherCurrAssets.priorAmountPaise || 0)
    : undefined;

  const totalAssetsPaise = totalNonCurrentAssetsPaise + totalCurrentAssetsPaise;
  const priorTotalAssetsPaise =
    hasPriorNonCurrAssets || hasPriorCurrAssets
      ? (priorTotalNonCurrentAssetsPaise || 0) + (priorTotalCurrentAssetsPaise || 0)
      : undefined;

  const equityBase = sumClassificationsForMonths(
    model,
    ['Equity'],
    'CREDIT_POSITIVE',
    undefined,
    true,
    bsSourceFilter
  );
  const longTermDebt = sumClassificationsForMonths(
    model,
    ['LongTermBorrowings'],
    'CREDIT_POSITIVE',
    undefined,
    true,
    bsSourceFilter
  );
  const shortTermDebt = sumClassificationsForMonths(
    model,
    ['ShortTermBorrowings'],
    'CREDIT_POSITIVE',
    undefined,
    true,
    bsSourceFilter
  );
  const payables = sumClassificationsForMonths(
    model,
    ['TradePayables'],
    'CREDIT_POSITIVE',
    undefined,
    true,
    bsSourceFilter
  );
  const taxLiab = sumClassificationsForMonths(
    model,
    ['TaxLiabilities'],
    'CREDIT_POSITIVE',
    undefined,
    true,
    bsSourceFilter
  );
  const otherCurrLiab = sumClassificationsForMonths(
    model,
    ['OtherCurrentLiabilities'],
    'CREDIT_POSITIVE',
    undefined,
    true,
    bsSourceFilter
  );

  const totalCurrentLiabilitiesPaise =
    shortTermDebt.amountPaise +
    payables.amountPaise +
    taxLiab.amountPaise +
    otherCurrLiab.amountPaise;
  const hasPriorCurrLiab =
    shortTermDebt.hasPrior ||
    payables.hasPrior ||
    taxLiab.hasPrior ||
    otherCurrLiab.hasPrior;
  const priorTotalCurrentLiabilitiesPaise = hasPriorCurrLiab
    ? (shortTermDebt.priorAmountPaise || 0) +
      (payables.priorAmountPaise || 0) +
      (taxLiab.priorAmountPaise || 0) +
      (otherCurrLiab.priorAmountPaise || 0)
    : undefined;

  // In a closed Schedule III Balance Sheet, Reserves & Surplus already includes current period PAT.
  // Only add currPnl.patPaise if an unclosed Trial Balance / Voucher set was uploaded where adding PAT reconciles Assets.
  const rawLiabAndEquityWithoutPat =
    equityBase.amountPaise + longTermDebt.amountPaise + totalCurrentLiabilitiesPaise;
  const rawLiabAndEquityWithPat = rawLiabAndEquityWithoutPat + currPnl.patPaise;
  const shouldAddPatToEquity =
    !model.hasBalanceSheetStatement &&
    Math.abs(totalAssetsPaise - rawLiabAndEquityWithPat) <
      Math.abs(totalAssetsPaise - rawLiabAndEquityWithoutPat);

  const totalEquityPaise = shouldAddPatToEquity
    ? equityBase.amountPaise + currPnl.patPaise
    : equityBase.amountPaise;
  const priorTotalEquityPaise = equityBase.priorAmountPaise;

  const totalEquityAndLiabilitiesPaise =
    totalEquityPaise + longTermDebt.amountPaise + totalCurrentLiabilitiesPaise;
  const priorTotalEquityAndLiabilitiesPaise =
    equityBase.hasPrior || longTermDebt.hasPrior || hasPriorCurrLiab
      ? (priorTotalEquityPaise || 0) +
        (longTermDebt.priorAmountPaise || 0) +
        (priorTotalCurrentLiabilitiesPaise || 0)
      : undefined;

  const suspense = sumClassificationsForMonths(
    model,
    ['Suspense'],
    'DEBIT_POSITIVE',
    undefined,
    true,
    bsSourceFilter
  );
  const unclassified = sumClassificationsForMonths(
    model,
    ['Unclassified'],
    'DEBIT_POSITIVE',
    undefined,
    true,
    bsSourceFilter
  );

  const totalAssetsSub = findSubtotalRef('BALANCE_SHEET', ['total assets', 'total asset']);
  const totalEqLiabSub = findSubtotalRef('BALANCE_SHEET', [
    'total equity and liabilities',
    'total equity & liabilities',
    'total liabilities',
  ]);

  const bsAssetLines: StatementLineItem[] = [
    makeLine(
      'NON_CURRENT_ASSETS',
      'Non-Current Assets (Property, Plant & Equipment, Intangibles, DTA & Investments)',
      ['FixedAssets', 'Investments', 'NonCurrentAssets'],
      totalNonCurrentAssetsPaise,
      priorTotalNonCurrentAssetsPaise,
      'Property, Plant & Equipment + Intangible Assets + Deferred Tax Assets + Long-Term Investments & Other Non-Current Assets',
      [...fixedAssets.ledgerNames, ...investments.ledgerNames, ...nonCurrOther.ledgerNames],
      false,
      {
        cellReferences: [
          ...fixedAssets.cellReferences,
          ...investments.cellReferences,
          ...nonCurrOther.cellReferences,
        ],
        priorCellReferences: [
          ...fixedAssets.priorCellReferences,
          ...investments.priorCellReferences,
          ...nonCurrOther.priorCellReferences,
        ],
      }
    ),
    makeLine(
      'INVENTORY',
      'Inventories / Closing Stock',
      ['Inventory'],
      inventoryPaise,
      priorInventoryPaise,
      'Closing Stock valuation from Stock Items or Inventories line',
      invLedger.ledgerNames,
      false,
      {
        cellReferences: invLedger.cellReferences,
        priorCellReferences: invLedger.priorCellReferences,
      }
    ),
    makeLine(
      'TRADE_RECEIVABLES',
      'Trade Receivables (Sundry Debtors)',
      ['TradeReceivables'],
      receivables.amountPaise,
      receivables.priorAmountPaise,
      'Net Closing Balance of Trade Receivables / Sundry Debtors',
      receivables.ledgerNames,
      false,
      {
        cellReferences: receivables.cellReferences,
        priorCellReferences: receivables.priorCellReferences,
      }
    ),
    makeLine(
      'CASH_AND_BANK',
      'Cash & Cash Equivalents',
      ['CashAndEquivalents', 'BankBalances'],
      cash.amountPaise + bank.amountPaise,
      cash.hasPrior || bank.hasPrior
        ? (cash.priorAmountPaise || 0) + (bank.priorAmountPaise || 0)
        : undefined,
      'Cash and Cash Equivalents + Bank Balances',
      [...cash.ledgerNames, ...bank.ledgerNames],
      false,
      {
        cellReferences: [...cash.cellReferences, ...bank.cellReferences],
        priorCellReferences: [...cash.priorCellReferences, ...bank.priorCellReferences],
      }
    ),
    makeLine(
      'OTHER_CURRENT_ASSETS',
      'Short-Term Loans & Advances / Other Current Assets',
      ['OtherCurrentAssets'],
      otherCurrAssets.amountPaise,
      otherCurrAssets.priorAmountPaise,
      'Short-Term Loans & Advances, Deposits, Prepaid Expenses & Other Current Assets',
      otherCurrAssets.ledgerNames,
      false,
      {
        cellReferences: otherCurrAssets.cellReferences,
        priorCellReferences: otherCurrAssets.priorCellReferences,
      }
    ),
    makeLine(
      'TOTAL_CURRENT_ASSETS',
      'Total Current Assets',
      ['Inventory', 'TradeReceivables', 'CashAndEquivalents', 'BankBalances', 'OtherCurrentAssets'],
      totalCurrentAssetsPaise,
      priorTotalCurrentAssetsPaise,
      'Inventories + Trade Receivables + Cash & Cash Equivalents + Other Current Assets',
      [],
      true
    ),
    makeLine(
      'TOTAL_ASSETS',
      'Total Assets',
      [],
      totalAssetsPaise,
      priorTotalAssetsPaise,
      'Non-Current Assets + Current Assets',
      [],
      true,
      {
        cellReferences: totalAssetsSub?.cellReference ? [totalAssetsSub.cellReference] : undefined,
        priorCellReferences: totalAssetsSub?.priorCellReference
          ? [totalAssetsSub.priorCellReference]
          : undefined,
      }
    ),
  ];

  const bsLiabilityEquityLines: StatementLineItem[] = [
    makeLine(
      'EQUITY',
      'Shareholders’ Funds (Share Capital & Reserves and Surplus)',
      ['Equity'],
      totalEquityPaise,
      priorTotalEquityPaise,
      shouldAddPatToEquity
        ? 'Share Capital + Reserves & Surplus + Current Period Profit After Tax'
        : 'Share Capital + Closing Reserves & Surplus (including period PAT)',
      equityBase.ledgerNames,
      false,
      {
        cellReferences: equityBase.cellReferences,
        priorCellReferences: equityBase.priorCellReferences,
      }
    ),
    makeLine(
      'NON_CURRENT_LIABILITIES',
      'Non-Current Liabilities (Long-Term Borrowings & Provisions)',
      ['LongTermBorrowings'],
      longTermDebt.amountPaise,
      longTermDebt.priorAmountPaise,
      'Long-Term Borrowings + Deferred Tax Liabilities + Long-Term Provisions',
      longTermDebt.ledgerNames,
      false,
      {
        cellReferences: longTermDebt.cellReferences,
        priorCellReferences: longTermDebt.priorCellReferences,
      }
    ),
    makeLine(
      'SHORT_TERM_BORROWINGS',
      'Short-Term Borrowings',
      ['ShortTermBorrowings'],
      shortTermDebt.amountPaise,
      shortTermDebt.priorAmountPaise,
      'Short-Term Borrowings, Bank OD & Cash Credit Facilities',
      shortTermDebt.ledgerNames,
      false,
      {
        cellReferences: shortTermDebt.cellReferences,
        priorCellReferences: shortTermDebt.priorCellReferences,
      }
    ),
    makeLine(
      'TRADE_PAYABLES',
      'Trade Payables (Sundry Creditors)',
      ['TradePayables'],
      payables.amountPaise,
      payables.priorAmountPaise,
      'Closing Balance of Trade Payables / Sundry Creditors',
      payables.ledgerNames,
      false,
      {
        cellReferences: payables.cellReferences,
        priorCellReferences: payables.priorCellReferences,
      }
    ),
    makeLine(
      'TAX_AND_OTHER_LIAB',
      'Other Current Liabilities & Short-Term Provisions',
      ['TaxLiabilities', 'OtherCurrentLiabilities'],
      taxLiab.amountPaise + otherCurrLiab.amountPaise,
      taxLiab.hasPrior || otherCurrLiab.hasPrior
        ? (taxLiab.priorAmountPaise || 0) + (otherCurrLiab.priorAmountPaise || 0)
        : undefined,
      'Other Current Liabilities + Short-Term Provisions + Statutory Dues',
      [...taxLiab.ledgerNames, ...otherCurrLiab.ledgerNames],
      false,
      {
        cellReferences: [...taxLiab.cellReferences, ...otherCurrLiab.cellReferences],
        priorCellReferences: [
          ...taxLiab.priorCellReferences,
          ...otherCurrLiab.priorCellReferences,
        ],
      }
    ),
    makeLine(
      'TOTAL_CURRENT_LIABILITIES',
      'Total Current Liabilities',
      ['ShortTermBorrowings', 'TradePayables', 'TaxLiabilities', 'OtherCurrentLiabilities'],
      totalCurrentLiabilitiesPaise,
      priorTotalCurrentLiabilitiesPaise,
      'Short-Term Borrowings + Trade Payables + Other Current Liabilities & Provisions',
      [],
      true
    ),
    makeLine(
      'TOTAL_EQUITY_AND_LIABILITIES',
      'Total Equity and Liabilities',
      [],
      totalEquityAndLiabilitiesPaise,
      priorTotalEquityAndLiabilitiesPaise,
      'Shareholders’ Funds + Non-Current Liabilities + Current Liabilities',
      [],
      true,
      {
        cellReferences: totalEqLiabSub?.cellReference ? [totalEqLiabSub.cellReference] : undefined,
        priorCellReferences: totalEqLiabSub?.priorCellReference
          ? [totalEqLiabSub.priorCellReference]
          : undefined,
      }
    ),
  ];

  // Monthly Trend series across availableMonths (only when genuine monthly data exists)
  const monthlyTrends = model.availableMonths.map((ym) => {
    const mMetrics = calcPnlMetrics(new Set([ym]));
    const expensesPaise =
      mMetrics.empExp.amountPaise +
      mMetrics.adminExp.amountPaise +
      mMetrics.sellExp.amountPaise +
      mMetrics.finCost.amountPaise +
      mMetrics.depExp.amountPaise;

    let collectionsPaise = 0;
    let paymentsPaise = 0;
    for (const vch of model.deduplicatedVouchers) {
      if (vch.date.slice(0, 7) !== ym) continue;
      if (vch.voucherType === 'Receipt') collectionsPaise += vch.totalCreditPaise;
      if (vch.voucherType === 'Payment') paymentsPaise += vch.totalDebitPaise;
    }

    return {
      month: ym,
      revenuePaise: mMetrics.netRevenuePaise,
      purchasesPaise: mMetrics.netPurchasesPaise,
      grossProfitPaise: mMetrics.grossProfitPaise,
      expensesPaise,
      netProfitPaise: mMetrics.patPaise,
      collectionsPaise,
      paymentsPaise,
      workingCapitalPaise: totalCurrentAssetsPaise - totalCurrentLiabilitiesPaise,
    };
  });

  return {
    periodLabel,
    comparativePeriodLabel,
    isIncompleteComparisonWarning,
    pnlLines,
    bsAssetLines,
    bsLiabilityEquityLines,
    kpis: {
      grossSalesPaise: currPnl.sales.amountPaise,
      salesReturnsPaise: currPnl.salesRet.amountPaise,
      netRevenuePaise: currPnl.netRevenuePaise,
      otherOperatingIncomePaise: currPnl.otherOpInc.amountPaise,
      otherIncomePaise: currPnl.otherInc.amountPaise,
      totalIncomePaise: currPnl.totalIncomePaise,
      purchasesPaise: currPnl.purch.amountPaise,
      purchaseReturnsPaise: currPnl.purchRet.amountPaise,
      netPurchasesPaise: currPnl.netPurchasesPaise,
      directExpensesPaise: currPnl.directExp.amountPaise,
      cogsPaise: currPnl.cogsPaise,
      hasGrossProfitData: currPnl.hasGrossProfitData,
      grossProfitPaise: currPnl.grossProfitPaise,
      employeeExpensesPaise: currPnl.empExp.amountPaise,
      adminExpensesPaise: currPnl.adminExp.amountPaise,
      sellingExpensesPaise: currPnl.sellExp.amountPaise,
      ebitdaPaise: currPnl.ebitdaPaise,
      operatingEbitdaPaise: currPnl.operatingEbitdaPaise,
      depreciationPaise: currPnl.depExp.amountPaise,
      ebitPaise: currPnl.ebitPaise,
      financeCostsPaise: currPnl.finCost.amountPaise,
      totalExpensesPaise: currPnl.totalExpensesPaise,
      profitBeforePriorPeriodPaise: currPnl.profitBeforePriorPeriodPaise,
      priorPeriodExpensesPaise: currPnl.priorPeriodExp.amountPaise,
      pbtPaise: currPnl.pbtPaise,
      currentTaxPaise: currPnl.currentTax.amountPaise,
      deferredTaxPaise: currPnl.deferredTax.amountPaise,
      taxExpensePaise: currPnl.taxExpensePaise,
      patPaise: currPnl.patPaise,
      fixedAssetsPaise: fixedAssets.amountPaise,
      investmentsPaise: investments.amountPaise,
      otherNonCurrentAssetsPaise: nonCurrOther.amountPaise,
      totalNonCurrentAssetsPaise,
      inventoryPaise,
      tradeReceivablesPaise: receivables.amountPaise,
      cashPaise: cash.amountPaise,
      bankPaise: bank.amountPaise,
      otherCurrentAssetsPaise: otherCurrAssets.amountPaise,
      totalCurrentAssetsPaise,
      totalAssetsPaise,
      equityBasePaise: equityBase.amountPaise,
      totalEquityPaise,
      longTermBorrowingsPaise: longTermDebt.amountPaise,
      shortTermBorrowingsPaise: shortTermDebt.amountPaise,
      tradePayablesPaise: payables.amountPaise,
      taxLiabilitiesPaise: taxLiab.amountPaise,
      otherCurrentLiabilitiesPaise: otherCurrLiab.amountPaise,
      totalCurrentLiabilitiesPaise,
      totalEquityAndLiabilitiesPaise,
      workingCapitalPaise: totalCurrentAssetsPaise - totalCurrentLiabilitiesPaise,
      suspensePaise: suspense.amountPaise,
      unclassifiedPaise: unclassified.amountPaise,
    },
    monthlyTrends,
  };
}
