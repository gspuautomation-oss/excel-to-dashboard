import {
  computeSafeVariance,
  ConsolidatedAccountingModel,
  FinancialInsight,
  FinancialRatioResult,
  formatINR,
} from '../types/accounting';
import { FinancialStatementBundle } from './accountingEngine';
import { AgeingAnalysisResult } from './ageingEngine';
import { AccountingHealthCheckResult } from './auditEngine';
import { GSTReconciliationReport } from './gstMatchingEngine';

export function generateFinancialInsights(
  model: ConsolidatedAccountingModel,
  statements: FinancialStatementBundle,
  priorStatements: FinancialStatementBundle | undefined,
  ageing: AgeingAnalysisResult,
  ratios: FinancialRatioResult[],
  health: AccountingHealthCheckResult,
  gst: GSTReconciliationReport
): FinancialInsight[] {
  const insights: FinancialInsight[] = [];
  const k = statements.kpis;
  const pk = priorStatements?.kpis;
  const sourceFileNames = model.sourceFiles.map((f) => f.filename);

  // 1. Revenue Performance Insight
  if (pk && pk.netRevenuePaise > 0 && k.netRevenuePaise > 0) {
    const revVar = computeSafeVariance(k.netRevenuePaise, pk.netRevenuePaise);
    if (revVar.percent !== null) {
      const isUp = revVar.percent >= 0;
      insights.push({
        id: 'INS_REV_COMP',
        category: 'Revenue performance',
        priority: isUp ? 'Positive development' : 'Attention required',
        headline: `Net Revenue ${isUp ? 'increased' : 'declined'} by ${Math.abs(revVar.percent).toFixed(1)}% compared with the comparative period.`,
        narrative: `Calculated Observation: Net Revenue moved from ${formatINR(pk.netRevenuePaise)} in ${statements.comparativePeriodLabel || 'the prior period'} to ${formatINR(k.netRevenuePaise)} in ${statements.periodLabel} (${revVar.label}).`,
        currentValueFormatted: formatINR(k.netRevenuePaise),
        comparativeValueFormatted: formatINR(pk.netRevenuePaise),
        absoluteMovementFormatted: formatINR(revVar.diffPaise),
        percentageMovementFormatted: revVar.label,
        formulaOrRule: '(Current Period Net Sales − Prior Period Net Sales) / Prior Period Net Sales × 100',
        supportingLedgers: statements.pnlLines[0]?.ledgerNames || ['Sales Accounts'],
        sourceFiles: sourceFileNames,
        drillDownCategory: 'Sales',
      });
    }
  } else if (k.netRevenuePaise > 0 && statements.monthlyTrends.length >= 2) {
    const firstM = statements.monthlyTrends[0];
    const lastM = statements.monthlyTrends[statements.monthlyTrends.length - 1];
    const mVar = computeSafeVariance(lastM.revenuePaise, firstM.revenuePaise);
    if (mVar.percent !== null) {
      insights.push({
        id: 'INS_REV_TREND',
        category: 'Revenue performance',
        priority: mVar.percent >= 0 ? 'Positive development' : 'Informational',
        headline: `Monthly revenue changed by ${mVar.label} between ${firstM.month} and ${lastM.month}.`,
        narrative: `Calculated Observation: Monthly net sales recorded ${formatINR(lastM.revenuePaise)} in ${lastM.month} compared with ${formatINR(firstM.revenuePaise)} in ${firstM.month}.`,
        currentValueFormatted: formatINR(lastM.revenuePaise),
        comparativeValueFormatted: formatINR(firstM.revenuePaise),
        absoluteMovementFormatted: formatINR(mVar.diffPaise),
        percentageMovementFormatted: mVar.label,
        formulaOrRule: 'Latest Month Net Revenue vs Opening Month Net Revenue',
        supportingLedgers: statements.pnlLines[0]?.ledgerNames || ['Sales Accounts'],
        sourceFiles: sourceFileNames,
        drillDownCategory: 'Sales',
      });
    }
  }

  // 2. Profitability & Gross Margin Insight
  const gpRatio = ratios.find((r) => r.id === 'GROSS_MARGIN');
  const npRatio = ratios.find((r) => r.id === 'NET_MARGIN');
  if (gpRatio && gpRatio.currentValue !== null) {
    if (gpRatio.previousValue !== null) {
      const ptsDiff = gpRatio.currentValue - gpRatio.previousValue;
      insights.push({
        id: 'INS_GP_MARGIN_COMP',
        category: 'Profitability',
        priority: ptsDiff >= 0 ? 'Positive development' : 'Attention required',
        headline: `Gross profit margin ${ptsDiff >= 0 ? 'improved' : 'declined'} by ${Math.abs(ptsDiff).toFixed(2)} percentage points.`,
        narrative: `Calculated Observation: Gross Margin stands at ${gpRatio.currentValue.toFixed(2)}% (${formatINR(k.grossProfitPaise)}) compared with ${gpRatio.previousValue.toFixed(2)}% in the comparative period.`,
        currentValueFormatted: `${gpRatio.currentValue.toFixed(2)}%`,
        comparativeValueFormatted: `${gpRatio.previousValue.toFixed(2)}%`,
        absoluteMovementFormatted: `${ptsDiff >= 0 ? '+' : ''}${ptsDiff.toFixed(2)} pp`,
        percentageMovementFormatted: `${ptsDiff >= 0 ? '+' : ''}${ptsDiff.toFixed(2)} pp`,
        formulaOrRule: 'Gross Profit / Net Revenue × 100',
        supportingLedgers: ['Sales Accounts', 'Purchase Accounts', 'Direct Expenses'],
        sourceFiles: sourceFileNames,
        drillDownCategory: 'Purchases',
      });
    } else if (npRatio && npRatio.currentValue !== null) {
      insights.push({
        id: 'INS_MARGIN_SUMMARY',
        category: 'Profitability',
        priority: npRatio.currentValue >= 10 ? 'Positive development' : 'Informational',
        headline: `Gross Profit Margin is ${gpRatio.currentValue.toFixed(2)}% and Net Profit Margin is ${npRatio.currentValue.toFixed(2)}%.`,
        narrative: `Calculated Observation: On net revenue of ${formatINR(k.netRevenuePaise)}, the enterprise generated Gross Profit of ${formatINR(k.grossProfitPaise)} and Profit After Tax of ${formatINR(k.patPaise)}.`,
        currentValueFormatted: `${npRatio.currentValue.toFixed(2)}%`,
        comparativeValueFormatted: `${gpRatio.currentValue.toFixed(2)}% (Gross Margin)`,
        absoluteMovementFormatted: formatINR(k.patPaise),
        percentageMovementFormatted: `${npRatio.currentValue.toFixed(2)}%`,
        formulaOrRule: 'PAT / Net Revenue × 100',
        supportingLedgers: ['Sales Accounts', 'Direct & Indirect Expenses'],
        sourceFiles: sourceFileNames,
        drillDownCategory: 'Sales',
      });
    }
  }

  // 3. Cost Control Insight
  const totalIndirectExpPaise =
    k.employeeExpensesPaise + k.adminExpensesPaise + k.sellingExpensesPaise;
  if (k.netRevenuePaise > 0 && totalIndirectExpPaise > 0) {
    const opExShare = (totalIndirectExpPaise / k.netRevenuePaise) * 100;
    const adminShare = (k.adminExpensesPaise / k.netRevenuePaise) * 100;
    if (pk && pk.adminExpensesPaise > 0 && pk.netRevenuePaise > 0) {
      const adminGrowth = computeSafeVariance(k.adminExpensesPaise, pk.adminExpensesPaise);
      const revGrowth = computeSafeVariance(k.netRevenuePaise, pk.netRevenuePaise);
      if (
        adminGrowth.percent !== null &&
        revGrowth.percent !== null &&
        adminGrowth.percent > revGrowth.percent
      ) {
        insights.push({
          id: 'INS_COST_ADMIN_FASTER',
          category: 'Cost control',
          priority: 'Attention required',
          headline: 'Administrative expenses increased faster than revenue during the selected period.',
          narrative: `Calculated Observation: Administrative expenses grew by ${adminGrowth.label} (from ${formatINR(pk.adminExpensesPaise)} to ${formatINR(k.adminExpensesPaise)}) while revenue changed by ${revGrowth.label}.`,
          currentValueFormatted: formatINR(k.adminExpensesPaise),
          comparativeValueFormatted: formatINR(pk.adminExpensesPaise),
          absoluteMovementFormatted: formatINR(adminGrowth.diffPaise),
          percentageMovementFormatted: adminGrowth.label,
          formulaOrRule: 'Administrative Expense Growth % vs Net Revenue Growth %',
          supportingLedgers: statements.pnlLines.find((l) => l.code === 'ADMIN_EXP')?.ledgerNames || ['Administrative Expenses'],
          sourceFiles: sourceFileNames,
          drillDownCategory: 'AdminExpenses',
        });
      }
    } else {
      insights.push({
        id: 'INS_COST_STRUCTURE',
        category: 'Cost control',
        priority: opExShare > 35 ? 'Attention required' : 'Informational',
        headline: `Operating overheads absorb ${opExShare.toFixed(1)}% of Net Revenue (Administrative share: ${adminShare.toFixed(1)}%).`,
        narrative: `Calculated Observation: Combined Employee (${formatINR(k.employeeExpensesPaise)}), Administrative (${formatINR(k.adminExpensesPaise)}), and Selling (${formatINR(k.sellingExpensesPaise)}) expenses total ${formatINR(totalIndirectExpPaise)}.`,
        currentValueFormatted: formatINR(totalIndirectExpPaise),
        comparativeValueFormatted: formatINR(k.netRevenuePaise),
        absoluteMovementFormatted: `${opExShare.toFixed(1)}% of Revenue`,
        percentageMovementFormatted: `${opExShare.toFixed(1)}%`,
        formulaOrRule: '(Employee + Admin + Selling Expenses) / Net Revenue × 100',
        supportingLedgers: statements.pnlLines.find((l) => l.code === 'ADMIN_EXP')?.ledgerNames || ['Indirect Expenses'],
        sourceFiles: sourceFileNames,
        drillDownCategory: 'AdminExpenses',
      });
    }
  }

  // 4. Working Capital & Receivables Insight
  if (ageing.hasBillLevelData && ageing.overdueReceivablesPaise > 0) {
    const overdueShare =
      ageing.totalBillTrackedReceivablesPaise > 0
        ? (ageing.overdueReceivablesPaise / ageing.totalBillTrackedReceivablesPaise) * 100
        : 0;
    insights.push({
      id: 'INS_WC_OVERDUE_REC',
      category: 'Working capital',
      priority: overdueShare > 30 ? 'Significant risk indicator' : 'Attention required',
      headline: `${overdueShare.toFixed(1)}% of bill-tracked trade receivables (${formatINR(ageing.overdueReceivablesPaise)}) are overdue past their credit period.`,
      narrative: `Calculated Observation: Across ${ageing.customersWithOverdueCount} customer account(s), ${formatINR(ageing.overdueReceivablesPaise)} out of ${formatINR(ageing.totalBillTrackedReceivablesPaise)} total tracked receivables exceeds the configured ${ageing.ageingBasis === 'DUE_DATE' ? 'due date' : 'invoice date'} threshold.`,
      currentValueFormatted: formatINR(ageing.overdueReceivablesPaise),
      comparativeValueFormatted: formatINR(ageing.totalBillTrackedReceivablesPaise),
      absoluteMovementFormatted: `${ageing.customersWithOverdueCount} Overdue Customers`,
      percentageMovementFormatted: `${overdueShare.toFixed(1)}% Overdue`,
      formulaOrRule: 'Sum of Unpaid Bills with Days Overdue > 0 / Total Bill-Tracked Receivables',
      supportingLedgers: ageing.customers.slice(0, 5).map((c) => c.partyName),
      sourceFiles: sourceFileNames,
      drillDownCategory: 'RECEIVABLES',
    });
  }

  // 5. Liquidity Insight
  const currRatio = ratios.find((r) => r.id === 'CURRENT_RATIO');
  if (currRatio && currRatio.currentValue !== null) {
    const isLow = currRatio.currentValue < 1.25;
    insights.push({
      id: 'INS_LIQ_CURRENT',
      category: 'Liquidity',
      priority: isLow ? 'Attention required' : 'Positive development',
      headline: `Current Ratio stands at ${currRatio.formattedCurrent} with Net Working Capital of ${formatINR(k.workingCapitalPaise)}.`,
      narrative: `Calculated Observation: Current Assets of ${formatINR(k.totalCurrentAssetsPaise)} (including Cash & Bank of ${formatINR(k.cashPaise + k.bankPaise)}) cover Current Liabilities of ${formatINR(k.totalCurrentLiabilitiesPaise)} by ${currRatio.formattedCurrent}.`,
      currentValueFormatted: currRatio.formattedCurrent,
      comparativeValueFormatted: currRatio.formattedPrevious,
      absoluteMovementFormatted: formatINR(k.workingCapitalPaise),
      percentageMovementFormatted: currRatio.formattedCurrent,
      formulaOrRule: 'Current Assets / Current Liabilities',
      supportingLedgers: currRatio.accountMappings,
      sourceFiles: sourceFileNames,
      drillDownCategory: 'BankBalances',
    });
  }

  // 6. Accounting Exceptions Insight
  if (health.totalExceptions > 0) {
    insights.push({
      id: 'INS_AUDIT_EXCEPTIONS',
      category: 'Accounting exceptions',
      priority: health.highPriorityCount > 0 ? 'Significant risk indicator' : 'Attention required',
      headline: `Accounting Health Check flagged ${health.totalExceptions} exception(s), including ${health.highPriorityCount} high-priority finding(s).`,
      narrative: `Calculated Observation: Deterministic audit rules detected ${health.unbalancedVoucherCount} unbalanced voucher(s), ${health.duplicateVoucherCount} duplicate/suspected voucher(s), ${health.reconciliationDiffCount} Trial Balance difference(s), and ${health.negativeCashDateCount} negative cash date(s).`,
      currentValueFormatted: `${health.totalExceptions} Exceptions`,
      comparativeValueFormatted: `${health.totalVouchersAnalysed} Vouchers Analysed`,
      absoluteMovementFormatted: `${health.highPriorityCount} High Severity`,
      percentageMovementFormatted: `Score: ${health.qualityScore.overallScore}/100`,
      formulaOrRule: 'Deterministic 10-Rule Audit Exception Engine',
      supportingLedgers: Array.from(new Set(health.exceptions.slice(0, 5).map((e) => e.ledgerName))),
      sourceFiles: sourceFileNames,
      drillDownCategory: 'EXCEPTIONS',
    });
  }

  // 7. GST Discrepancies Insight
  if (gst.hasGstData) {
    const itcDiff = gst.summaryCards.bookItcPaise - gst.summaryCards.gstr2bItcPaise;
    if (Math.abs(itcDiff) > 0 || gst.summaryCards.unmatchedInvoiceCount > 0) {
      insights.push({
        id: 'INS_GST_RECON',
        category: 'GST discrepancies',
        priority: Math.abs(itcDiff) > 100000 ? 'Significant risk indicator' : 'Attention required',
        headline:
          itcDiff > 0
            ? `Input Tax Credit in the books exceeds the corresponding GSTR-2B amount by ${formatINR(itcDiff)}.`
            : `GSTR-2B Input Tax Credit exceeds Purchase Books ITC by ${formatINR(Math.abs(itcDiff))}.`,
        narrative: `Calculated Observation: Books ITC is ${formatINR(gst.summaryCards.bookItcPaise)} while GSTR-2B eligible ITC is ${formatINR(gst.summaryCards.gstr2bItcPaise)}, with ${gst.summaryCards.unmatchedInvoiceCount} unmatched or probable-match invoice(s) requiring reconciliation.`,
        currentValueFormatted: formatINR(gst.summaryCards.bookItcPaise),
        comparativeValueFormatted: formatINR(gst.summaryCards.gstr2bItcPaise),
        absoluteMovementFormatted: formatINR(itcDiff),
        percentageMovementFormatted: `${gst.summaryCards.unmatchedInvoiceCount} Unmatched Invoices`,
        formulaOrRule: 'Books Purchase Register Input GST − GSTR-2B Eligible Input GST',
        supportingLedgers: ['Input IGST', 'Input CGST', 'Input SGST'],
        sourceFiles: sourceFileNames,
        drillDownCategory: 'GST',
      });
    }
  }

  // 8. Financial Risk / Solvency Insight
  const deRatio = ratios.find((r) => r.id === 'DEBT_EQUITY');
  if (deRatio && deRatio.currentValue !== null) {
    insights.push({
      id: 'INS_RISK_SOLVENCY',
      category: 'Financial risk',
      priority: deRatio.currentValue > 2.0 ? 'Significant risk indicator' : 'Informational',
      headline: `Debt-to-Equity ratio is ${deRatio.formattedCurrent} on total borrowings of ${formatINR(k.longTermBorrowingsPaise + k.shortTermBorrowingsPaise)}.`,
      narrative: `Calculated Observation: Total interest-bearing debt (${formatINR(k.longTermBorrowingsPaise + k.shortTermBorrowingsPaise)}) compared against Shareholders' Equity (${formatINR(k.totalEquityPaise)}) yields a leverage multiple of ${deRatio.formattedCurrent}.`,
      currentValueFormatted: deRatio.formattedCurrent,
      comparativeValueFormatted: formatINR(k.totalEquityPaise),
      absoluteMovementFormatted: formatINR(k.longTermBorrowingsPaise + k.shortTermBorrowingsPaise),
      percentageMovementFormatted: deRatio.formattedCurrent,
      formulaOrRule: '(Long-Term + Short-Term Borrowings) / Total Equity',
      supportingLedgers: deRatio.accountMappings,
      sourceFiles: sourceFileNames,
      drillDownCategory: 'LongTermBorrowings',
    });
  }

  return insights;
}
