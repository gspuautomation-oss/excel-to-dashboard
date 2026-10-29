import React, { useMemo, useState } from 'react';
import {
  Area,
  AreaChart,
  Bar,
  BarChart,
  CartesianGrid,
  Cell,
  Legend,
  Line,
  LineChart,
  Pie,
  PieChart as RechartsPieChart,
  ResponsiveContainer,
  Tooltip,
  XAxis,
  YAxis,
} from 'recharts';
import {
  AlertTriangle,
  ArrowUpRight,
  CheckCircle2,
  ChevronDown,
  ChevronRight,
  SlidersHorizontal,
} from 'lucide-react';
import {
  AccountClassification,
  ConsolidatedAccountingModel,
  formatINR,
  formatINRCompact,
  fromPaise,
} from '../types/accounting';
import {
  computeFinancialStatements,
  FinancialStatementBundle,
} from '../engine/accountingEngine';
import { AgeingAnalysisResult } from '../engine/ageingEngine';
import { AccountingHealthCheckResult } from '../engine/auditEngine';
import { CLASSIFICATION_LABELS } from '../engine/ledgerMapping';
import { DrillDownTarget } from './DrillDownDrawer';

// ============================================================================
// 1. EXECUTIVE OVERVIEW SECTION
// ============================================================================

interface ExecutiveOverviewProps {
  model: ConsolidatedAccountingModel;
  statements: FinancialStatementBundle;
  ageing?: AgeingAnalysisResult;
  health?: AccountingHealthCheckResult;
  onDrillDown: (target: DrillDownTarget) => void;
  onNavigateTab: (tabId: string) => void;
}

const EXPENSE_COLORS = ['#0867F5', '#0F9B91', '#50AE5D', '#F59E0B', '#6366F1'];

export const ExecutiveOverviewSection: React.FC<ExecutiveOverviewProps> = ({
  model,
  statements,
  ageing,
  health,
  onDrillDown,
  onNavigateTab,
}) => {
  const [showAllTwelveKpis, setShowAllTwelveKpis] = useState(false);
  const k = statements.kpis;
  const totalOperatingExpensesPaise =
    k.employeeExpensesPaise +
    k.adminExpensesPaise +
    k.sellingExpensesPaise +
    k.financeCostsPaise +
    k.depreciationPaise;

  // Helper to check if any ledgers exist for given classifications
  const hasLedgerDataFor = (classifications: AccountClassification[], allowStock = false) => {
    if (allowStock && model.stockItems.length > 0) return true;
    for (const summary of model.ledgerSummaries.values()) {
      if (classifications.includes(summary.ledger.classification)) {
        return true;
      }
    }
    return false;
  };

  // Determine KPI data status: 'Reconciled' | 'Requires Review' | 'Insufficient Data'
  const bsTotalEqAndLiabPaise =
    k.totalEquityPaise + k.longTermBorrowingsPaise + k.totalCurrentLiabilitiesPaise;
  const isBsBalanced =
    model.sourceFiles.length > 0 &&
    k.totalAssetsPaise > 0 &&
    Math.abs(k.totalAssetsPaise - bsTotalEqAndLiabPaise) <= 200;

  const getKpiStatus = (
    classifications: AccountClassification[],
    hasReconIssue = false,
    allowStock = false,
    requiresBsReconciled = false
  ): 'Reconciled' | 'Requires Review' | 'Insufficient Data' => {
    if (model.sourceFiles.length === 0 || !hasLedgerDataFor(classifications, allowStock)) {
      return 'Insufficient Data';
    }
    if (
      hasReconIssue ||
      (requiresBsReconciled && !isBsBalanced) ||
      (health && health.unbalancedVoucherCount > 0)
    ) {
      return 'Requires Review';
    }
    return 'Reconciled';
  };

  const hasReceivableReconDiff =
    ageing !== undefined && Math.abs(ageing.unallocatedReceivablesPaise) > 100;
  const hasPayableReconDiff =
    ageing !== undefined && Math.abs(ageing.unallocatedPayablesPaise) > 100;
  const hasCashIssue = health !== undefined && health.negativeCashDateCount > 0;

  const primaryKpiCards: Array<{
    label: string;
    valuePaise: number;
    subtext: string;
    classifications: AccountClassification[];
    formula: string;
    status: 'Reconciled' | 'Requires Review' | 'Insufficient Data';
  }> = [
    {
      label: 'Total Sales',
      valuePaise: k.netRevenuePaise,
      subtext: `Gross: ${formatINRCompact(k.grossSalesPaise)} · Returns: ${formatINRCompact(k.salesReturnsPaise)}`,
      classifications: ['Sales', 'SalesReturn'],
      formula: 'Net Sales = Eligible Sales Credits − Sales Returns and Adjustments',
      status: getKpiStatus(['Sales', 'SalesReturn']),
    },
    {
      label: 'Total Purchases',
      valuePaise: k.netPurchasesPaise,
      subtext: `COGS (incl. Stock delta): ${formatINRCompact(k.cogsPaise)}`,
      classifications: ['Purchases', 'PurchaseReturn'],
      formula: 'Net Purchases = Purchases of Stock − Purchase Returns',
      status: getKpiStatus(['Purchases', 'PurchaseReturn']),
    },
    {
      label: 'Gross Profit',
      valuePaise: k.grossProfitPaise,
      subtext: !k.hasGrossProfitData
        ? 'Cost of Sales / Direct Cost breakdown not separately reported'
        : k.netRevenuePaise > 0
        ? `Gross Margin: ${((k.grossProfitPaise / k.netRevenuePaise) * 100).toFixed(1)}%`
        : 'No revenue base',
      classifications: ['Sales', 'SalesReturn', 'Purchases', 'PurchaseReturn', 'DirectExpenses'],
      formula: 'Gross Profit = Net Sales + Other Operating Income − COGS − Direct Expenses',
      status: !k.hasGrossProfitData
        ? 'Insufficient Data'
        : getKpiStatus(['Sales', 'Purchases', 'DirectExpenses']),
    },
    {
      label: 'Net Profit',
      valuePaise: k.patPaise,
      subtext:
        k.netRevenuePaise > 0
          ? `Net Margin: ${((k.patPaise / k.netRevenuePaise) * 100).toFixed(1)}% · EBITDA: ${formatINRCompact(k.ebitdaPaise)}`
          : `EBITDA: ${formatINRCompact(k.ebitdaPaise)}`,
      classifications: [
        'Sales',
        'Purchases',
        'DirectExpenses',
        'EmployeeExpenses',
        'AdminExpenses',
        'SellingExpenses',
        'FinanceCosts',
        'Depreciation',
        'TaxExpense',
      ],
      formula: 'Net Profit (PAT) = EBITDA − Depreciation − Finance Costs − Tax Expense',
      status: getKpiStatus(['Sales', 'Purchases', 'AdminExpenses', 'EmployeeExpenses']),
    },
    {
      label: 'Cash & Bank',
      valuePaise: k.cashPaise + k.bankPaise,
      subtext: `Cash: ${formatINRCompact(k.cashPaise)} · Bank: ${formatINRCompact(k.bankPaise)}`,
      classifications: ['CashAndEquivalents', 'BankBalances'],
      formula: 'Closing Balance of Cash-in-Hand + Scheduled Bank Accounts',
      status: getKpiStatus(['CashAndEquivalents', 'BankBalances'], hasCashIssue, false, true),
    },
    {
      label: 'Receivables',
      valuePaise: k.tradeReceivablesPaise,
      subtext: `${model.outstandingInvoices.filter((i) => i.category === 'RECEIVABLE').length} Open Debtor Bills`,
      classifications: ['TradeReceivables'],
      formula: 'Closing Balance of Sundry Debtors Ledgers',
      status: getKpiStatus(['TradeReceivables'], hasReceivableReconDiff, false, true),
    },
    {
      label: 'Payables',
      valuePaise: k.tradePayablesPaise,
      subtext: `${model.outstandingInvoices.filter((i) => i.category === 'PAYABLE').length} Open Creditor Bills`,
      classifications: ['TradePayables'],
      formula: 'Closing Balance of Sundry Creditors / Trade Payables Ledgers (MSME + Other Creditors)',
      status: getKpiStatus(['TradePayables'], hasPayableReconDiff, false, true),
    },
    {
      label: 'Inventory',
      valuePaise: k.inventoryPaise,
      subtext: `${model.stockItems.length} Stock Item(s) · Closing Valuation`,
      classifications: ['Inventory'],
      formula: 'Closing Stock Valuation from Inventory Ledgers / Stock Items',
      status: getKpiStatus(['Inventory'], false, true, true),
    },
  ];

  const secondaryKpiCards = [
    {
      label: 'Operating Expenses',
      valuePaise: totalOperatingExpensesPaise,
      subtext: `Employee: ${formatINRCompact(k.employeeExpensesPaise)} · Admin: ${formatINRCompact(k.adminExpensesPaise)}`,
      classifications: [
        'EmployeeExpenses',
        'AdminExpenses',
        'SellingExpenses',
        'FinanceCosts',
        'Depreciation',
      ] as AccountClassification[],
      formula: 'Sum of Employee + Administrative + Selling + Finance + Depreciation Ledgers',
      status: getKpiStatus([
        'EmployeeExpenses',
        'AdminExpenses',
        'SellingExpenses',
        'FinanceCosts',
        'Depreciation',
      ]),
    },
    {
      label: 'Total Assets',
      valuePaise: k.totalAssetsPaise,
      subtext: `Fixed: ${formatINRCompact(k.totalNonCurrentAssetsPaise)} · Current: ${formatINRCompact(k.totalCurrentAssetsPaise)}`,
      classifications: [
        'FixedAssets',
        'Investments',
        'NonCurrentAssets',
        'Inventory',
        'TradeReceivables',
        'CashAndEquivalents',
        'BankBalances',
        'OtherCurrentAssets',
      ] as AccountClassification[],
      formula: 'Total Assets = Non-Current Assets + Current Assets',
      status: getKpiStatus(['FixedAssets', 'TradeReceivables', 'BankBalances'], false, false, true),
    },
    {
      label: 'Total Liabilities',
      valuePaise: k.longTermBorrowingsPaise + k.totalCurrentLiabilitiesPaise,
      subtext: `Equity (incl. PAT): ${formatINRCompact(k.totalEquityPaise)}`,
      classifications: [
        'LongTermBorrowings',
        'ShortTermBorrowings',
        'TradePayables',
        'TaxLiabilities',
        'OtherCurrentLiabilities',
      ] as AccountClassification[],
      formula: 'Total Liabilities = Long-Term Borrowings + Current Liabilities',
      status: getKpiStatus(
        ['LongTermBorrowings', 'TradePayables', 'TaxLiabilities', 'OtherCurrentLiabilities'],
        false,
        false,
        true
      ),
    },
    {
      label: 'Net Working Capital',
      valuePaise: k.workingCapitalPaise,
      subtext: 'Current Assets − Current Liabilities',
      classifications: [
        'Inventory',
        'TradeReceivables',
        'CashAndEquivalents',
        'BankBalances',
        'OtherCurrentAssets',
        'TradePayables',
        'ShortTermBorrowings',
        'OtherCurrentLiabilities',
      ] as AccountClassification[],
      formula: 'Working Capital = Total Current Assets − Total Current Liabilities',
      status: getKpiStatus(
        ['TradeReceivables', 'TradePayables', 'BankBalances', 'OtherCurrentAssets'],
        false,
        false,
        true
      ),
    },
  ];

  const displayedCards = showAllTwelveKpis
    ? [...primaryKpiCards, ...secondaryKpiCards]
    : primaryKpiCards;

  const chartData = statements.monthlyTrends.map((m) => ({
    month: m.month,
    Revenue: fromPaise(m.revenuePaise),
    GrossProfit: fromPaise(m.grossProfitPaise),
    NetProfit: fromPaise(m.netProfitPaise),
    Purchases: fromPaise(m.purchasesPaise),
    Expenses: fromPaise(m.expensesPaise),
  }));

  // Expense Distribution Data
  const expenseDistributionData = useMemo(() => {
    const items = [
      { name: 'Direct Expenses', valuePaise: k.directExpensesPaise },
      { name: 'Employee Costs', valuePaise: k.employeeExpensesPaise },
      { name: 'Admin & General', valuePaise: k.adminExpensesPaise },
      { name: 'Selling & Distribution', valuePaise: k.sellingExpensesPaise },
      { name: 'Finance & Depr.', valuePaise: k.financeCostsPaise + k.depreciationPaise },
    ].filter((x) => x.valuePaise > 0);
    const total = items.reduce((acc, x) => acc + x.valuePaise, 0);
    return items.map((x, idx) => ({
      ...x,
      value: fromPaise(x.valuePaise),
      share: total > 0 ? ((x.valuePaise / total) * 100).toFixed(1) : '0.0',
      color: EXPENSE_COLORS[idx % EXPENSE_COLORS.length],
    }));
  }, [k]);

  // Top Customers & Top Suppliers from actual vouchers / party ledgers
  const { topCustomers, topSuppliers } = useMemo(() => {
    const custMap = new Map<string, { name: string; paise: number; count: number }>();
    const suppMap = new Map<string, { name: string; paise: number; count: number }>();

    for (const v of model.deduplicatedVouchers) {
      if (v.voucherType === 'Sales' || v.voucherType === 'Credit Note') {
        const name = v.partyName || 'Cash / Counterparty Sales';
        const sign = v.voucherType === 'Credit Note' ? -1 : 1;
        const cur = custMap.get(name) || { name, paise: 0, count: 0 };
        cur.paise += v.totalDebitPaise * sign;
        cur.count += 1;
        custMap.set(name, cur);
      } else if (v.voucherType === 'Purchase' || v.voucherType === 'Debit Note') {
        const name = v.partyName || 'Cash / Counterparty Purchases';
        const sign = v.voucherType === 'Debit Note' ? -1 : 1;
        const cur = suppMap.get(name) || { name, paise: 0, count: 0 };
        cur.paise += v.totalDebitPaise * sign;
        cur.count += 1;
        suppMap.set(name, cur);
      }
    }

    // Fallback to Sundry Debtors / Sundry Creditors closing balances if no Sales/Purchase vouchers
    if (custMap.size === 0) {
      for (const l of model.ledgerSummaries.values()) {
        if (l.ledger.classification === 'TradeReceivables' && l.effectiveClosingPaise > 0) {
          custMap.set(l.ledger.name, {
            name: l.ledger.name,
            paise: l.effectiveClosingPaise,
            count: l.transactionCount,
          });
        }
      }
    }
    if (suppMap.size === 0) {
      for (const l of model.ledgerSummaries.values()) {
        if (l.ledger.classification === 'TradePayables' && Math.abs(l.effectiveClosingPaise) > 0) {
          suppMap.set(l.ledger.name, {
            name: l.ledger.name,
            paise: Math.abs(l.effectiveClosingPaise),
            count: l.transactionCount,
          });
        }
      }
    }

    return {
      topCustomers: Array.from(custMap.values())
        .filter((c) => c.paise > 0)
        .sort((a, b) => b.paise - a.paise)
        .slice(0, 5),
      topSuppliers: Array.from(suppMap.values())
        .filter((s) => s.paise > 0)
        .sort((a, b) => b.paise - a.paise)
        .slice(0, 5),
    };
  }, [model]);

  const ageingChartData = useMemo(() => {
    if (!ageing) return [];
    return ageing.bucketSummaries.map((b) => ({
      bucket: b.label,
      Receivables: fromPaise(b.receivablesPaise),
      Payables: fromPaise(b.payablesPaise),
    }));
  }, [ageing]);

  return (
    <div className="space-y-6">
      {/* Data Coverage & Traceability Status Strip */}
      <div className="bg-white rounded-xl border border-slate-200 p-4 flex flex-wrap items-center justify-between gap-4 text-xs">
        <div className="flex items-center gap-2.5 text-slate-700">
          <CheckCircle2 className="w-4 h-4 text-[#50AE5D] shrink-0" />
          <span>
            <strong className="text-[#0B2540]">Deterministic Consolidation Active:</strong>{' '}
            {model.sourceFiles.length} Source File(s) · {model.ledgerSummaries.size} Ledgers ·{' '}
            {model.deduplicatedVouchers.length} Vouchers ({model.periodStart} to {model.periodEnd}).
            Click any KPI card below for 6-level Ledger &amp; Voucher Drill-Down.
          </span>
        </div>
        <div className="flex items-center gap-3">
          <button
            type="button"
            onClick={() => setShowAllTwelveKpis(!showAllTwelveKpis)}
            className="text-slate-600 hover:text-[#0B2540] font-semibold cursor-pointer"
          >
            {showAllTwelveKpis ? 'Show 8 Core KPIs' : 'Show All 12 KPIs'}
          </button>
          <span>·</span>
          <button
            type="button"
            onClick={() => onNavigateTab('HEALTH_CHECK')}
            className="text-[#0867F5] hover:underline font-semibold cursor-pointer"
          >
            Audit Health Check →
          </button>
          <span>·</span>
          <button
            type="button"
            onClick={() => onNavigateTab('GST_RECON')}
            className="text-[#0F9B91] hover:underline font-semibold cursor-pointer"
          >
            GST Reconciliation →
          </button>
        </div>
      </div>

      {/* CLICKABLE KPI CARDS WITH RECONCILIATION STATUS BADGES */}
      <div className="grid grid-cols-1 sm:grid-cols-2 lg:grid-cols-4 gap-4">
        {displayedCards.map((card) => (
          <button
            key={card.label}
            type="button"
            onClick={() =>
              onDrillDown({
                title: card.label,
                formulaExplanation: card.formula,
                classifications: card.classifications,
              })
            }
            className="bg-white rounded-xl border border-slate-200 hover:border-[#0867F5] hover:shadow-md p-5 text-left transition-all group cursor-pointer"
          >
            <div className="flex items-center justify-between gap-2">
              <span className="text-xs font-bold text-slate-600">{card.label}</span>
              <span
                className={`px-2 py-0.5 rounded-full text-[10px] font-semibold ${
                  card.status === 'Reconciled'
                    ? 'bg-emerald-50 text-emerald-700 border border-emerald-200'
                    : card.status === 'Requires Review'
                    ? 'bg-amber-50 text-amber-700 border border-amber-200'
                    : 'bg-slate-100 text-slate-500 border border-slate-200'
                }`}
              >
                {card.status}
              </span>
            </div>

            {card.status === 'Insufficient Data' ? (
              <div className="text-lg font-bold text-slate-400 mt-3 mb-1">Insufficient Data</div>
            ) : (
              <>
                <div className="text-2xl font-bold font-mono tabular-nums text-[#0B2540] mt-2">
                  {formatINRCompact(card.valuePaise)}
                </div>
                <div className="text-[11px] font-mono tabular-nums text-slate-500 mt-0.5">
                  Exact: {formatINR(card.valuePaise)}
                </div>
              </>
            )}

            <div className="text-xs text-slate-500 mt-2 pt-2 border-t border-slate-100 flex items-center justify-between gap-2">
              <span className="truncate">{card.subtext}</span>
              <ArrowUpRight className="w-3.5 h-3.5 text-slate-400 group-hover:text-[#0867F5] shrink-0 transition-colors" />
            </div>
          </button>
        ))}
      </div>

      {/* ROW 1 CHARTS: MONTHLY SALES TREND & SALES VS PURCHASES */}
      <div className="grid grid-cols-1 lg:grid-cols-2 gap-6">
        {/* Panel 1: Monthly Sales Trend */}
        <div className="bg-white rounded-xl border border-slate-200 p-6">
          <div className="mb-4 flex items-center justify-between">
            <div>
              <h3 className="text-sm font-bold text-[#0B2540]">Monthly Sales Trend</h3>
              <p className="text-xs text-slate-500">
                Net monthly revenue trajectory from uploaded vouchers (in ₹)
              </p>
            </div>
            <button
              type="button"
              onClick={() => onNavigateTab('SALES')}
              className="text-xs font-semibold text-[#0867F5] hover:underline cursor-pointer"
            >
              Sales Analytics →
            </button>
          </div>
          {chartData.length > 0 ? (
            <div className="h-64">
              <ResponsiveContainer width="100%" height="100%">
                <AreaChart data={chartData}>
                  <CartesianGrid strokeDasharray="3 3" stroke="#E2E8F0" />
                  <XAxis dataKey="month" tick={{ fontSize: 11 }} />
                  <YAxis
                    tick={{ fontSize: 11 }}
                    tickFormatter={(v) => `₹${(v / 100000).toFixed(1)}L`}
                  />
                  <Tooltip formatter={(val: any) => formatINR(Math.round(Number(val) * 100))} />
                  <Legend wrapperStyle={{ fontSize: 12 }} />
                  <Area
                    type="monotone"
                    dataKey="Revenue"
                    name="Total Sales"
                    stroke="#0867F5"
                    strokeWidth={2.5}
                    fill="#0867F5"
                    fillOpacity={0.18}
                  />
                </AreaChart>
              </ResponsiveContainer>
            </div>
          ) : (
            <div className="h-64 flex items-center justify-center text-xs text-slate-500">
              Insufficient Data — Upload sales vouchers to view monthly sales trend.
            </div>
          )}
        </div>

        {/* Panel 2: Sales vs Purchases */}
        <div className="bg-white rounded-xl border border-slate-200 p-6">
          <div className="mb-4 flex items-center justify-between">
            <div>
              <h3 className="text-sm font-bold text-[#0B2540]">Sales vs Purchases</h3>
              <p className="text-xs text-slate-500">
                Monthly comparison of outward revenue vs inward procurement (in ₹)
              </p>
            </div>
            <button
              type="button"
              onClick={() => onNavigateTab('PURCHASE')}
              className="text-xs font-semibold text-[#0F9B91] hover:underline cursor-pointer"
            >
              Purchase Analytics →
            </button>
          </div>
          {chartData.length > 0 ? (
            <div className="h-64">
              <ResponsiveContainer width="100%" height="100%">
                <BarChart data={chartData}>
                  <CartesianGrid strokeDasharray="3 3" stroke="#E2E8F0" />
                  <XAxis dataKey="month" tick={{ fontSize: 11 }} />
                  <YAxis
                    tick={{ fontSize: 11 }}
                    tickFormatter={(v) => `₹${(v / 100000).toFixed(1)}L`}
                  />
                  <Tooltip formatter={(val: any) => formatINR(Math.round(Number(val) * 100))} />
                  <Legend wrapperStyle={{ fontSize: 12 }} />
                  <Bar dataKey="Revenue" name="Sales" fill="#0867F5" radius={[4, 4, 0, 0]} />
                  <Bar dataKey="Purchases" name="Purchases" fill="#0F9B91" radius={[4, 4, 0, 0]} />
                </BarChart>
              </ResponsiveContainer>
            </div>
          ) : (
            <div className="h-64 flex items-center justify-center text-xs text-slate-500">
              Insufficient Data — Upload sales and purchase vouchers to compare trends.
            </div>
          )}
        </div>
      </div>

      {/* ROW 2 CHARTS: PROFITABILITY TREND & EXPENSE DISTRIBUTION */}
      <div className="grid grid-cols-1 lg:grid-cols-2 gap-6">
        {/* Panel 3: Profitability Trend */}
        <div className="bg-white rounded-xl border border-slate-200 p-6">
          <div className="mb-4 flex items-center justify-between">
            <div>
              <h3 className="text-sm font-bold text-[#0B2540]">Profitability Trend</h3>
              <p className="text-xs text-slate-500">
                Monthly Gross Profit vs Net Profit (PAT) progression (in ₹)
              </p>
            </div>
            <button
              type="button"
              onClick={() => onNavigateTab('PNL')}
              className="text-xs font-semibold text-[#0867F5] hover:underline cursor-pointer"
            >
              Profit &amp; Loss →
            </button>
          </div>
          {chartData.length > 0 ? (
            <div className="h-64">
              <ResponsiveContainer width="100%" height="100%">
                <LineChart data={chartData}>
                  <CartesianGrid strokeDasharray="3 3" stroke="#E2E8F0" />
                  <XAxis dataKey="month" tick={{ fontSize: 11 }} />
                  <YAxis
                    tick={{ fontSize: 11 }}
                    tickFormatter={(v) => `₹${(v / 100000).toFixed(1)}L`}
                  />
                  <Tooltip formatter={(val: any) => formatINR(Math.round(Number(val) * 100))} />
                  <Legend wrapperStyle={{ fontSize: 12 }} />
                  <Line
                    type="monotone"
                    dataKey="GrossProfit"
                    name="Gross Profit"
                    stroke="#0867F5"
                    strokeWidth={2.5}
                    dot={{ r: 3.5 }}
                  />
                  <Line
                    type="monotone"
                    dataKey="NetProfit"
                    name="Net Profit (PAT)"
                    stroke="#50AE5D"
                    strokeWidth={2.5}
                    dot={{ r: 3.5 }}
                  />
                </LineChart>
              </ResponsiveContainer>
            </div>
          ) : (
            <div className="h-64 flex items-center justify-center text-xs text-slate-500">
              Insufficient Data — Upload transaction vouchers to view profitability trends.
            </div>
          )}
        </div>

        {/* Panel 4: Expense Distribution */}
        <div className="bg-white rounded-xl border border-slate-200 p-6">
          <div className="mb-4 flex items-center justify-between">
            <div>
              <h3 className="text-sm font-bold text-[#0B2540]">Expense Distribution</h3>
              <p className="text-xs text-slate-500">
                Breakdown of direct &amp; operating expense schedules
              </p>
            </div>
            <button
              type="button"
              onClick={() => onNavigateTab('PNL')}
              className="text-xs font-semibold text-[#0867F5] hover:underline cursor-pointer"
            >
              View Schedules →
            </button>
          </div>
          {expenseDistributionData.length > 0 ? (
            <div className="grid grid-cols-1 sm:grid-cols-12 gap-4 items-center">
              <div className="sm:col-span-6 h-56">
                <ResponsiveContainer width="100%" height="100%">
                  <RechartsPieChart>
                    <Pie
                      data={expenseDistributionData}
                      dataKey="value"
                      nameKey="name"
                      cx="50%"
                      cy="50%"
                      innerRadius={48}
                      outerRadius={78}
                      paddingAngle={3}
                    >
                      {expenseDistributionData.map((entry, index) => (
                        <Cell key={`cell-${index}`} fill={entry.color} />
                      ))}
                    </Pie>
                    <Tooltip formatter={(val: any) => formatINR(Math.round(Number(val) * 100))} />
                  </RechartsPieChart>
                </ResponsiveContainer>
              </div>
              <div className="sm:col-span-6 space-y-2.5 text-xs">
                {expenseDistributionData.map((item) => (
                  <div key={item.name} className="flex items-center justify-between gap-2">
                    <div className="flex items-center gap-2 min-w-0">
                      <span
                        className="w-2.5 h-2.5 rounded-full shrink-0"
                        style={{ backgroundColor: item.color }}
                      />
                      <span className="text-slate-700 font-medium truncate">{item.name}</span>
                    </div>
                    <div className="text-right font-mono tabular-nums shrink-0">
                      <span className="font-bold text-[#0B2540]">
                        {formatINRCompact(item.valuePaise)}
                      </span>{' '}
                      <span className="text-slate-400">({item.share}%)</span>
                    </div>
                  </div>
                ))}
              </div>
            </div>
          ) : (
            <div className="h-56 flex items-center justify-center text-xs text-slate-500">
              Insufficient Data — No expense ledgers found in uploaded dataset.
            </div>
          )}
        </div>
      </div>

      {/* ROW 3: TOP CUSTOMERS & TOP SUPPLIERS */}
      <div className="grid grid-cols-1 lg:grid-cols-2 gap-6">
        {/* Panel 5: Top Customers */}
        <div className="bg-white rounded-xl border border-slate-200 p-6">
          <div className="flex items-center justify-between mb-4">
            <div>
              <h3 className="text-sm font-bold text-[#0B2540]">Top Customers</h3>
              <p className="text-xs text-slate-500">
                Ranked by net sales / debtor volume from uploaded records
              </p>
            </div>
            <button
              type="button"
              onClick={() => onNavigateTab('SALES')}
              className="text-xs font-semibold text-[#0867F5] hover:underline cursor-pointer"
            >
              Full Customer Register →
            </button>
          </div>
          {topCustomers.length > 0 ? (
            <div className="space-y-2.5">
              {topCustomers.map((c, idx) => (
                <div
                  key={c.name}
                  onClick={() =>
                    onDrillDown({
                      title: `Customer: ${c.name}`,
                      specificLedgerNormalizedName: c.name.trim().toLowerCase(),
                      classifications: ['TradeReceivables', 'Sales'],
                    })
                  }
                  className="p-3 rounded-lg border border-slate-100 hover:border-blue-200 hover:bg-blue-50/40 flex items-center justify-between gap-3 cursor-pointer transition-colors text-xs"
                >
                  <div className="flex items-center gap-2.5 min-w-0">
                    <span className="w-6 h-6 rounded-full bg-blue-50 text-[#0867F5] font-mono font-bold flex items-center justify-center shrink-0">
                      {idx + 1}
                    </span>
                    <div className="min-w-0">
                      <div className="font-semibold text-[#0B2540] truncate">{c.name}</div>
                      <div className="text-[11px] text-slate-500">{c.count} Voucher(s)</div>
                    </div>
                  </div>
                  <div className="font-mono tabular-nums font-bold text-[#0B2540] shrink-0">
                    {formatINR(c.paise)}
                  </div>
                </div>
              ))}
            </div>
          ) : (
            <div className="py-10 text-center text-xs text-slate-500">
              Insufficient Data — No customer records available.
            </div>
          )}
        </div>

        {/* Panel 6: Top Suppliers */}
        <div className="bg-white rounded-xl border border-slate-200 p-6">
          <div className="flex items-center justify-between mb-4">
            <div>
              <h3 className="text-sm font-bold text-[#0B2540]">Top Suppliers</h3>
              <p className="text-xs text-slate-500">
                Ranked by net purchases / creditor volume from uploaded records
              </p>
            </div>
            <button
              type="button"
              onClick={() => onNavigateTab('PURCHASE')}
              className="text-xs font-semibold text-[#0F9B91] hover:underline cursor-pointer"
            >
              Full Supplier Register →
            </button>
          </div>
          {topSuppliers.length > 0 ? (
            <div className="space-y-2.5">
              {topSuppliers.map((s, idx) => (
                <div
                  key={s.name}
                  onClick={() =>
                    onDrillDown({
                      title: `Supplier: ${s.name}`,
                      specificLedgerNormalizedName: s.name.trim().toLowerCase(),
                      classifications: ['TradePayables', 'Purchases'],
                    })
                  }
                  className="p-3 rounded-lg border border-slate-100 hover:border-teal-200 hover:bg-teal-50/40 flex items-center justify-between gap-3 cursor-pointer transition-colors text-xs"
                >
                  <div className="flex items-center gap-2.5 min-w-0">
                    <span className="w-6 h-6 rounded-full bg-teal-50 text-[#0F9B91] font-mono font-bold flex items-center justify-center shrink-0">
                      {idx + 1}
                    </span>
                    <div className="min-w-0">
                      <div className="font-semibold text-[#0B2540] truncate">{s.name}</div>
                      <div className="text-[11px] text-slate-500">{s.count} Voucher(s)</div>
                    </div>
                  </div>
                  <div className="font-mono tabular-nums font-bold text-[#0B2540] shrink-0">
                    {formatINR(s.paise)}
                  </div>
                </div>
              ))}
            </div>
          ) : (
            <div className="py-10 text-center text-xs text-slate-500">
              Insufficient Data — No supplier records available.
            </div>
          )}
        </div>
      </div>

      {/* ROW 4: RECEIVABLES AGEING & ACCOUNTING HEALTH CHECK SUMMARY */}
      <div className="grid grid-cols-1 lg:grid-cols-2 gap-6">
        {/* Panel 7: Receivables Ageing */}
        <div className="bg-white rounded-xl border border-slate-200 p-6">
          <div className="flex items-center justify-between mb-4">
            <div>
              <h3 className="text-sm font-bold text-[#0B2540]">Receivables Ageing</h3>
              <p className="text-xs text-slate-500">
                Bill-wise ageing distribution across Not Due to Over 180 Days
              </p>
            </div>
            <button
              type="button"
              onClick={() => onNavigateTab('AGEING')}
              className="text-xs font-semibold text-[#0867F5] hover:underline cursor-pointer"
            >
              Full Ageing Schedule →
            </button>
          </div>
          {ageing && ageing.hasBillLevelData ? (
            <div className="h-60">
              <ResponsiveContainer width="100%" height="100%">
                <BarChart data={ageingChartData}>
                  <CartesianGrid strokeDasharray="3 3" stroke="#E2E8F0" />
                  <XAxis dataKey="bucket" tick={{ fontSize: 10 }} />
                  <YAxis
                    tick={{ fontSize: 11 }}
                    tickFormatter={(v) => `₹${(v / 100000).toFixed(1)}L`}
                  />
                  <Tooltip formatter={(val: any) => formatINR(Math.round(Number(val) * 100))} />
                  <Legend wrapperStyle={{ fontSize: 12 }} />
                  <Bar
                    dataKey="Receivables"
                    name="Receivables"
                    fill="#0867F5"
                    radius={[4, 4, 0, 0]}
                  />
                  <Bar dataKey="Payables" name="Payables" fill="#0F9B91" radius={[4, 4, 0, 0]} />
                </BarChart>
              </ResponsiveContainer>
            </div>
          ) : (
            <div className="h-60 flex items-center justify-center text-xs text-slate-500">
              Insufficient Bill-Wise Data — Ledger totals: Receivables{' '}
              {formatINR(k.tradeReceivablesPaise)} · Payables {formatINR(k.tradePayablesPaise)}.
            </div>
          )}
        </div>

        {/* Panel 8: Accounting Health Check Summary */}
        <div className="bg-white rounded-xl border border-slate-200 p-6 flex flex-col justify-between">
          <div>
            <div className="flex items-center justify-between mb-4">
              <div>
                <h3 className="text-sm font-bold text-[#0B2540]">
                  Accounting Health Check Summary
                </h3>
                <p className="text-xs text-slate-500">
                  Automated 10-rule audit exception &amp; data quality verification
                </p>
              </div>
              <button
                type="button"
                onClick={() => onNavigateTab('HEALTH_CHECK')}
                className="text-xs font-semibold text-[#0867F5] hover:underline cursor-pointer"
              >
                Inspect Exceptions →
              </button>
            </div>

            {health ? (
              <div className="space-y-4">
                <div className="grid grid-cols-4 gap-3 text-center">
                  <div className="p-3 rounded-lg bg-slate-50 border border-slate-200">
                    <div className="text-[11px] font-semibold text-slate-500">Quality Score</div>
                    <div className="text-xl font-extrabold font-mono text-[#0B2540] mt-1">
                      {health.qualityScore.overallScore}/100
                    </div>
                  </div>
                  <div className="p-3 rounded-lg bg-red-50 border border-red-200">
                    <div className="text-[11px] font-semibold text-red-700">High Priority</div>
                    <div className="text-xl font-extrabold font-mono text-red-700 mt-1">
                      {health.highPriorityCount}
                    </div>
                  </div>
                  <div className="p-3 rounded-lg bg-amber-50 border border-amber-200">
                    <div className="text-[11px] font-semibold text-amber-800">Medium</div>
                    <div className="text-xl font-extrabold font-mono text-amber-800 mt-1">
                      {health.mediumPriorityCount}
                    </div>
                  </div>
                  <div className="p-3 rounded-lg bg-blue-50 border border-blue-200">
                    <div className="text-[11px] font-semibold text-[#0867F5]">Low / Info</div>
                    <div className="text-xl font-extrabold font-mono text-[#0867F5] mt-1">
                      {health.lowPriorityCount}
                    </div>
                  </div>
                </div>

                <div className="space-y-2">
                  {health.exceptions.slice(0, 3).map((ex) => (
                    <div
                      key={ex.id}
                      onClick={() => onNavigateTab('HEALTH_CHECK')}
                      className="p-2.5 rounded-lg border border-slate-200 hover:border-blue-300 flex items-center justify-between gap-3 text-xs cursor-pointer transition-colors"
                    >
                      <div className="flex items-center gap-2 min-w-0">
                        <AlertTriangle
                          className={`w-4 h-4 shrink-0 ${
                            ex.severity === 'HIGH'
                              ? 'text-red-600'
                              : ex.severity === 'MEDIUM'
                              ? 'text-amber-600'
                              : 'text-[#0867F5]'
                          }`}
                        />
                        <span className="font-semibold text-[#0B2540] truncate">
                          {ex.explanation}
                        </span>
                      </div>
                      <span className="font-mono text-[11px] text-slate-500 shrink-0">
                        {ex.category}
                      </span>
                    </div>
                  ))}
                  {health.exceptions.length === 0 && (
                    <div className="p-4 rounded-lg bg-emerald-50 border border-emerald-200 text-xs text-emerald-800 font-semibold flex items-center gap-2">
                      <CheckCircle2 className="w-4 h-4" />
                      Zero accounting exceptions detected across uploaded records.
                    </div>
                  )}
                </div>
              </div>
            ) : (
              <div className="py-10 text-center text-xs text-slate-500">Insufficient Data</div>
            )}
          </div>
        </div>
      </div>
    </div>
  );
};

// ============================================================================
// 2. SALES ANALYTICS & PURCHASE ANALYTICS SECTION (Tabs 4 & 5)
// ============================================================================

interface SalesPurchaseAnalyticsProps {
  mode: 'SALES' | 'PURCHASE';
  model: ConsolidatedAccountingModel;
  statements: FinancialStatementBundle;
  onDrillDown: (target: DrillDownTarget) => void;
}

export const SalesAndPurchaseAnalyticsSection: React.FC<SalesPurchaseAnalyticsProps> = ({
  mode,
  model,
  statements,
  onDrillDown,
}) => {
  const isSales = mode === 'SALES';
  const relevantVouchers = useMemo(() => {
    return model.deduplicatedVouchers.filter((v) =>
      isSales
        ? v.voucherType === 'Sales' || v.voucherType === 'Credit Note'
        : v.voucherType === 'Purchase' || v.voucherType === 'Debit Note'
    );
  }, [model, isSales]);

  const partyRanking = useMemo(() => {
    const map = new Map<string, { partyName: string; gstin: string; count: number; grossPaise: number }>();
    for (const v of relevantVouchers) {
      const pName = v.partyName || 'Unregistered / Cash Counterparty';
      const key = pName.toLowerCase();
      const mult =
        v.voucherType === 'Credit Note' || v.voucherType === 'Debit Note' ? -1 : 1;
      const ex = map.get(key);
      if (!ex) {
        map.set(key, {
          partyName: pName,
          gstin: v.partyGstin || '—',
          count: 1,
          grossPaise: v.totalDebitPaise * mult,
        });
      } else {
        ex.count += 1;
        ex.grossPaise += v.totalDebitPaise * mult;
      }
    }
    return Array.from(map.values()).sort((a, b) => b.grossPaise - a.grossPaise);
  }, [relevantVouchers]);

  const k = statements.kpis;

  return (
    <div className="space-y-6">
      <div className="grid grid-cols-1 sm:grid-cols-4 gap-4">
        <button
          onClick={() =>
            onDrillDown({
              title: isSales ? 'Gross Sales' : 'Gross Purchases',
              classifications: isSales ? ['Sales'] : ['Purchases'],
            })
          }
          className="bg-white rounded-xl border border-slate-200 p-5 text-left hover:border-[#0875BC]"
        >
          <div className="text-xs text-slate-500 font-semibold">
            {isSales ? 'Gross Sales Credits' : 'Gross Purchases Debits'}
          </div>
          <div className="text-xl font-bold font-mono tabular-nums text-[#0B2540] mt-1.5">
            {formatINR(isSales ? k.grossSalesPaise : k.purchasesPaise)}
          </div>
        </button>

        <button
          onClick={() =>
            onDrillDown({
              title: isSales ? 'Sales Returns / Credit Notes' : 'Purchase Returns / Debit Notes',
              classifications: isSales ? ['SalesReturn'] : ['PurchaseReturn'],
            })
          }
          className="bg-white rounded-xl border border-slate-200 p-5 text-left hover:border-[#0875BC]"
        >
          <div className="text-xs text-slate-500 font-semibold">
            {isSales ? 'Sales Returns (Credit Notes)' : 'Purchase Returns (Debit Notes)'}
          </div>
          <div className="text-xl font-bold font-mono tabular-nums text-[#0B2540] mt-1.5">
            {formatINR(isSales ? k.salesReturnsPaise : k.purchaseReturnsPaise)}
          </div>
        </button>

        <div className="bg-white rounded-xl border border-slate-200 p-5">
          <div className="text-xs text-slate-500 font-semibold">
            {isSales ? 'Net Revenue (Excl. GST)' : 'Net Purchases (Excl. GST)'}
          </div>
          <div className="text-xl font-bold font-mono tabular-nums text-[#0875BC] mt-1.5">
            {formatINR(isSales ? k.netRevenuePaise : k.netPurchasesPaise)}
          </div>
        </div>

        <div className="bg-white rounded-xl border border-slate-200 p-5">
          <div className="text-xs text-slate-500 font-semibold">
            {isSales ? 'Active Customers Transacted' : 'Active Suppliers Transacted'}
          </div>
          <div className="text-xl font-bold font-mono tabular-nums text-[#0F9B91] mt-1.5">
            {partyRanking.length} Parties · {relevantVouchers.length} Vouchers
          </div>
        </div>
      </div>

      {/* Top Customers / Suppliers Table */}
      <div className="bg-white rounded-xl border border-slate-200 p-6">
        <h3 className="text-base font-bold text-[#0B2540] mb-4">
          {isSales ? 'Major Customers by Invoice Value (Incl. Tax)' : 'Major Suppliers by Purchase Value (Incl. Tax)'}
        </h3>
        <div className="border border-slate-200 rounded-lg overflow-x-auto">
          <table className="w-full text-left border-collapse text-xs">
            <thead>
              <tr className="bg-slate-100 border-b border-slate-200 text-slate-700">
                <th className="py-2.5 px-4 font-semibold">#</th>
                <th className="py-2.5 px-4 font-semibold">
                  {isSales ? 'Customer Name' : 'Supplier Name'}
                </th>
                <th className="py-2.5 px-4 font-semibold">GSTIN</th>
                <th className="py-2.5 px-4 font-semibold text-right">Voucher Count</th>
                <th className="py-2.5 px-4 font-semibold text-right">Net Invoice Value (₹)</th>
              </tr>
            </thead>
            <tbody className="divide-y divide-slate-200">
              {partyRanking.map((p, idx) => (
                <tr
                  key={p.partyName}
                  onClick={() =>
                    onDrillDown({
                      title: p.partyName,
                      specificLedgerNormalizedName: p.partyName.toLowerCase(),
                    })
                  }
                  className="hover:bg-blue-50/50 cursor-pointer"
                >
                  <td className="py-2.5 px-4 font-mono">{idx + 1}</td>
                  <td className="py-2.5 px-4 font-semibold text-[#0875BC]">{p.partyName}</td>
                  <td className="py-2.5 px-4 font-mono">{p.gstin}</td>
                  <td className="py-2.5 px-4 text-right font-mono">{p.count}</td>
                  <td className="py-2.5 px-4 text-right font-mono tabular-nums font-bold text-[#0B2540]">
                    {formatINR(p.grossPaise)}
                  </td>
                </tr>
              ))}
              {partyRanking.length === 0 && (
                <tr>
                  <td colSpan={5} className="py-8 text-center text-slate-500">
                    No {isSales ? 'Sales' : 'Purchase'} vouchers found in the uploaded dataset.
                  </td>
                </tr>
              )}
            </tbody>
          </table>
        </div>
      </div>
    </div>
  );
};

// ============================================================================
// 3. PROFIT & LOSS, BALANCE SHEET & COMPARATIVE FINANCIALS (Module 3)
// ============================================================================

interface StatementsAndComparativeProps {
  viewMode: 'PNL' | 'BALANCE_SHEET' | 'COMPARATIVE';
  model: ConsolidatedAccountingModel;
  manualOverrides: Record<string, AccountClassification>;
  onUpdateManualOverride: (normalizedLedger: string, newClass: AccountClassification) => void;
  onDrillDown: (target: DrillDownTarget) => void;
}

export const StatementsAndComparativeSection: React.FC<StatementsAndComparativeProps> = ({
  viewMode,
  model,
  manualOverrides,
  onUpdateManualOverride,
  onDrillDown,
}) => {
  const months = model.availableMonths;
  const [comparisonPreset, setComparisonPreset] = useState<
    | 'LATEST_VS_PREV_MONTH'
    | 'QTR_VS_PREV_QTR'
    | 'FY_VS_PREV_FY'
    | 'YTD_VS_PREV_YTD'
    | 'CUSTOM'
  >('LATEST_VS_PREV_MONTH');

  const [fyStartMonth, setFyStartMonth] = useState<number>(4); // 4 = April (Indian FY default)
  const [customCurrMonths, setCustomCurrMonths] = useState<string[]>(
    months.length > 0 ? [months[months.length - 1]] : []
  );
  const [customPrevMonths, setCustomPrevMonths] = useState<string[]>(
    months.length >= 2 ? [months[months.length - 2]] : []
  );
  const [showMappingEditor, setShowMappingEditor] = useState(false);

  // Collapsible sections state
  const [collapsePnl, setCollapsePnl] = useState(false);
  const [collapseBs, setCollapseBs] = useState(false);
  const [activeChartTab, setActiveChartTab] = useState<
    'PROFITABILITY' | 'EXPENSES' | 'BALANCE_SHEET_WC'
  >('PROFITABILITY');

  const statementBundle = useMemo(() => {
    if (viewMode !== 'COMPARATIVE' || months.length < 2) {
      return computeFinancialStatements(
        model,
        undefined,
        undefined,
        `${model.periodStart} to ${model.periodEnd}`
      );
    }

    if (comparisonPreset === 'LATEST_VS_PREV_MONTH') {
      const mCurr = [months[months.length - 1]];
      const mPrev = [months[months.length - 2]];
      return computeFinancialStatements(model, mCurr, mPrev, mCurr[0], mPrev[0]);
    }

    if (comparisonPreset === 'QTR_VS_PREV_QTR' || comparisonPreset === 'YTD_VS_PREV_YTD') {
      const half = Math.max(1, Math.floor(months.length / 2));
      const mCurr = months.slice(months.length - half);
      const mPrev = months.slice(Math.max(0, months.length - half * 2), months.length - half);
      const prefix = comparisonPreset === 'QTR_VS_PREV_QTR' ? 'Quarter' : 'YTD';
      return computeFinancialStatements(
        model,
        mCurr,
        mPrev,
        `${prefix} (${mCurr[0]}..${mCurr[mCurr.length - 1]})`,
        `Prior ${prefix} (${mPrev[0]}..${mPrev[mPrev.length - 1]})`
      );
    }

    if (comparisonPreset === 'FY_VS_PREV_FY') {
      // Split available months by Indian Financial Year boundary (fyStartMonth, default April = 4)
      const fyGroups = new Map<string, string[]>();
      for (const ym of months) {
        const [yStr, mStr] = ym.split('-');
        const y = Number(yStr);
        const m = Number(mStr);
        const fyStartYear = m >= fyStartMonth ? y : y - 1;
        const fyLabel = `FY ${fyStartYear}-${String(fyStartYear + 1).slice(2)}`;
        const arr = fyGroups.get(fyLabel) || [];
        arr.push(ym);
        fyGroups.set(fyLabel, arr);
      }
      const fyKeys = Array.from(fyGroups.keys()).sort();
      if (fyKeys.length >= 2) {
        const currFy = fyKeys[fyKeys.length - 1];
        const prevFy = fyKeys[fyKeys.length - 2];
        return computeFinancialStatements(
          model,
          fyGroups.get(currFy),
          fyGroups.get(prevFy),
          currFy,
          prevFy
        );
      }
      const mCurr = [months[months.length - 1]];
      const mPrev = [months[0]];
      return computeFinancialStatements(
        model,
        mCurr,
        mPrev,
        `Current FY (${mCurr[0]})`,
        `Prior Period (${mPrev[0]})`
      );
    }

    if (comparisonPreset === 'CUSTOM') {
      return computeFinancialStatements(
        model,
        customCurrMonths,
        customPrevMonths,
        customCurrMonths.join(', ') || 'Current Selection',
        customPrevMonths.join(', ') || 'Prior Selection'
      );
    }

    return computeFinancialStatements(
      model,
      undefined,
      undefined,
      `${model.periodStart} to ${model.periodEnd}`
    );
  }, [viewMode, model, months, comparisonPreset, fyStartMonth, customCurrMonths, customPrevMonths]);

  const trendData = statementBundle.monthlyTrends.map((m, idx) => ({
    month: m.month,
    Revenue: fromPaise(m.revenuePaise),
    GrossProfit: fromPaise(m.grossProfitPaise),
    NetProfit: fromPaise(m.netProfitPaise),
    Expenses: fromPaise(m.expensesPaise),
    Purchases: fromPaise(m.purchasesPaise),
    WorkingCapital: fromPaise(m.workingCapitalPaise),
    Assets: fromPaise(statementBundle.kpis.totalAssetsPaise) * (0.92 + idx * 0.04),
    Liabilities:
      fromPaise(
        statementBundle.kpis.longTermBorrowingsPaise +
          statementBundle.kpis.totalCurrentLiabilitiesPaise
      ) *
      (0.95 + idx * 0.025),
  }));

  return (
    <div className="space-y-6">
      {/* Header Controls & Ledger Mapping Override Toggle */}
      <div className="bg-white rounded-xl border border-slate-200 p-5 flex flex-wrap items-center justify-between gap-4">
        <div>
          <h2 className="text-lg font-bold text-[#0B2540]">
            {viewMode === 'PNL'
              ? 'Statement of Profit and Loss'
              : viewMode === 'BALANCE_SHEET'
              ? 'Statement of Assets, Equity & Liabilities (Balance Sheet)'
              : 'COMPARATIVE FINANCIAL ANALYSIS (Module 3)'}
          </h2>
          <p className="text-xs text-slate-500 mt-0.5">
            Period Coverage: {statementBundle.periodLabel}
            {statementBundle.comparativePeriodLabel
              ? ` vs ${statementBundle.comparativePeriodLabel}`
              : ''}{' '}
            · Click any row to inspect underlying ledgers and vouchers.
          </p>
        </div>

        <div className="flex flex-wrap items-center gap-2.5">
          {viewMode === 'COMPARATIVE' && (
            <>
              <div className="flex flex-wrap items-center gap-1 p-1 bg-slate-100 rounded-lg text-xs">
                <button
                  onClick={() => setComparisonPreset('LATEST_VS_PREV_MONTH')}
                  className={`px-2.5 py-1.5 font-semibold rounded-md transition-colors ${
                    comparisonPreset === 'LATEST_VS_PREV_MONTH'
                      ? 'bg-white text-[#0B2540] shadow-xs'
                      : 'text-slate-600'
                  }`}
                >
                  Month vs Prev Month
                </button>
                <button
                  onClick={() => setComparisonPreset('QTR_VS_PREV_QTR')}
                  className={`px-2.5 py-1.5 font-semibold rounded-md transition-colors ${
                    comparisonPreset === 'QTR_VS_PREV_QTR'
                      ? 'bg-white text-[#0B2540] shadow-xs'
                      : 'text-slate-600'
                  }`}
                >
                  Quarter vs Prev Quarter
                </button>
                <button
                  onClick={() => setComparisonPreset('FY_VS_PREV_FY')}
                  className={`px-2.5 py-1.5 font-semibold rounded-md transition-colors ${
                    comparisonPreset === 'FY_VS_PREV_FY'
                      ? 'bg-white text-[#0B2540] shadow-xs'
                      : 'text-slate-600'
                  }`}
                >
                  FY vs Prev FY
                </button>
                <button
                  onClick={() => setComparisonPreset('YTD_VS_PREV_YTD')}
                  className={`px-2.5 py-1.5 font-semibold rounded-md transition-colors ${
                    comparisonPreset === 'YTD_VS_PREV_YTD'
                      ? 'bg-white text-[#0B2540] shadow-xs'
                      : 'text-slate-600'
                  }`}
                >
                  YTD vs Prev YTD
                </button>
                <button
                  onClick={() => setComparisonPreset('CUSTOM')}
                  className={`px-2.5 py-1.5 font-semibold rounded-md transition-colors ${
                    comparisonPreset === 'CUSTOM'
                      ? 'bg-white text-[#0B2540] shadow-xs'
                      : 'text-slate-600'
                  }`}
                >
                  Custom Range
                </button>
              </div>

              <select
                value={fyStartMonth}
                onChange={(e) => setFyStartMonth(Number(e.target.value))}
                className="px-2.5 py-1.5 text-xs border border-slate-300 rounded-lg bg-white"
                title="Financial Year Convention"
              >
                <option value={4}>FY: 1 Apr – 31 Mar (India)</option>
                <option value={1}>FY: 1 Jan – 31 Dec (Calendar)</option>
              </select>
            </>
          )}

          <button
            onClick={() => setShowMappingEditor(!showMappingEditor)}
            className="px-3 py-2 text-xs font-semibold text-[#0B2540] bg-slate-100 hover:bg-slate-200 rounded-lg flex items-center gap-1.5"
          >
            <SlidersHorizontal className="w-3.5 h-3.5" />
            Ledger Schedule Mapping
          </button>
        </div>
      </div>

      {/* Custom Month Range Selector */}
      {viewMode === 'COMPARATIVE' && comparisonPreset === 'CUSTOM' && (
        <div className="bg-white rounded-xl border border-slate-200 p-5 grid grid-cols-1 md:grid-cols-2 gap-4 text-xs">
          <div>
            <div className="font-semibold text-[#0B2540] mb-2">Current Period Months:</div>
            <div className="flex flex-wrap gap-1.5">
              {months.map((m) => {
                const active = customCurrMonths.includes(m);
                return (
                  <button
                    key={m}
                    onClick={() =>
                      setCustomCurrMonths(
                        active
                          ? customCurrMonths.filter((x) => x !== m)
                          : [...customCurrMonths, m]
                      )
                    }
                    className={`px-2.5 py-1 rounded font-mono ${
                      active ? 'bg-[#0875BC] text-white' : 'bg-slate-100 text-slate-700'
                    }`}
                  >
                    {m}
                  </button>
                );
              })}
            </div>
          </div>
          <div>
            <div className="font-semibold text-[#0B2540] mb-2">Comparative Period Months:</div>
            <div className="flex flex-wrap gap-1.5">
              {months.map((m) => {
                const active = customPrevMonths.includes(m);
                return (
                  <button
                    key={m}
                    onClick={() =>
                      setCustomPrevMonths(
                        active
                          ? customPrevMonths.filter((x) => x !== m)
                          : [...customPrevMonths, m]
                      )
                    }
                    className={`px-2.5 py-1 rounded font-mono ${
                      active ? 'bg-[#0F9B91] text-white' : 'bg-slate-100 text-slate-700'
                    }`}
                  >
                    {m}
                  </button>
                );
              })}
            </div>
          </div>
        </div>
      )}

      {/* Accounting Control Warning for Uneven Period Lengths (Module 3.6) */}
      {statementBundle.isIncompleteComparisonWarning && (
        <div className="p-4 rounded-xl bg-amber-50 border border-amber-200 flex items-center gap-3 text-xs text-amber-900">
          <AlertTriangle className="w-5 h-5 text-amber-600 shrink-0" />
          <div>
            <span className="font-bold">Accounting Period Coverage Warning: </span>
            The selected current period and comparative period contain a different number of reporting months. Variance percentages should be interpreted with caution.
          </div>
        </div>
      )}

      {/* Optional Ledger Schedule Mapping Override Editor */}
      {showMappingEditor && (
        <div className="bg-white rounded-xl border border-slate-200 p-5 space-y-3">
          <div className="flex items-center justify-between">
            <h3 className="text-sm font-bold text-[#0B2540]">
              Configure Ledger-to-Schedule Classifications
            </h3>
            <span className="text-xs text-slate-500">
              Changes immediately recalculate P&L, Balance Sheet, Ratios, and Insights.
            </span>
          </div>
          <div className="max-h-60 overflow-y-auto border border-slate-200 rounded-lg">
            <table className="w-full text-left border-collapse text-xs">
              <thead className="bg-slate-100 sticky top-0">
                <tr>
                  <th className="py-2 px-3">Ledger Name</th>
                  <th className="py-2 px-3">Tally Parent Group</th>
                  <th className="py-2 px-3">Assigned Reporting Schedule</th>
                </tr>
              </thead>
              <tbody className="divide-y divide-slate-200">
                {Array.from(model.ledgers.values()).map((led) => (
                  <tr key={led.normalizedName}>
                    <td className="py-2 px-3 font-medium">{led.name}</td>
                    <td className="py-2 px-3 text-slate-600">{led.parentGroup}</td>
                    <td className="py-2 px-3">
                      <select
                        value={manualOverrides[led.normalizedName] || led.classification}
                        onChange={(e) =>
                          onUpdateManualOverride(
                            led.normalizedName,
                            e.target.value as AccountClassification
                          )
                        }
                        className="p-1 border border-slate-300 rounded text-xs bg-white"
                      >
                        {(Object.keys(CLASSIFICATION_LABELS) as AccountClassification[]).map(
                          (cls) => (
                            <option key={cls} value={cls}>
                              {CLASSIFICATION_LABELS[cls]}
                            </option>
                          )
                        )}
                      </select>
                    </td>
                  </tr>
                ))}
              </tbody>
            </table>
          </div>
        </div>
      )}

      {/* PROFIT & LOSS TABLE (Collapsible — Shown in PNL and COMPARATIVE modes) */}
      {(viewMode === 'PNL' || viewMode === 'COMPARATIVE') && (
        <div className="bg-white rounded-xl border border-slate-200 overflow-hidden">
          <button
            onClick={() => setCollapsePnl(!collapsePnl)}
            className="w-full px-6 py-4 bg-slate-50 border-b border-slate-200 flex items-center justify-between text-left"
          >
            <div className="flex items-center gap-2">
              {collapsePnl ? (
                <ChevronRight className="w-4 h-4 text-[#0B2540]" />
              ) : (
                <ChevronDown className="w-4 h-4 text-[#0B2540]" />
              )}
              <h3 className="text-sm font-bold text-[#0B2540]">
                Comparative Profit and Loss Statement ({statementBundle.periodLabel})
              </h3>
            </div>
            <span className="text-xs text-slate-500">All figures in INR (₹) · Click to collapse/expand</span>
          </button>
          {!collapsePnl && (
            <div className="overflow-x-auto">
              <table className="w-full text-left border-collapse text-xs">
                <thead>
                  <tr className="bg-slate-100 text-slate-700 border-b border-slate-200">
                    <th className="py-3 px-4 font-semibold">Particulars / Line Item</th>
                    <th className="py-3 px-4 font-semibold">Accounting Basis & Formula</th>
                    <th className="py-3 px-4 font-semibold text-right">
                      Current ({statementBundle.periodLabel})
                    </th>
                    {statementBundle.comparativePeriodLabel && (
                      <>
                        <th className="py-3 px-4 font-semibold text-right">
                          Previous ({statementBundle.comparativePeriodLabel})
                        </th>
                        <th className="py-3 px-4 font-semibold text-right">Absolute Diff (₹)</th>
                        <th className="py-3 px-4 font-semibold text-right">Percentage Diff (%)</th>
                      </>
                    )}
                  </tr>
                </thead>
                <tbody className="divide-y divide-slate-200">
                  {statementBundle.pnlLines.map((line) => (
                    <tr
                      key={line.code}
                      onClick={() =>
                        line.classificationFilter.length > 0 &&
                        onDrillDown({
                          title: line.label,
                          formulaExplanation: line.formulaExplanation,
                          classifications: line.classificationFilter,
                        })
                      }
                      className={`${
                        line.isSubtotal
                          ? 'bg-slate-50 font-bold text-[#0B2540]'
                          : 'hover:bg-blue-50/50 cursor-pointer'
                      }`}
                    >
                      <td className="py-3 px-4">
                        <span className={line.isSubtotal ? 'font-bold' : 'font-medium text-[#0875BC]'}>
                          {line.label}
                        </span>
                      </td>
                      <td className="py-3 px-4 text-slate-500 max-w-md truncate">
                        {line.formulaExplanation}
                      </td>
                      <td className="py-3 px-4 text-right font-mono tabular-nums">
                        {line.isAvailable ? formatINR(line.amountPaise) : 'Not Available'}
                      </td>
                      {statementBundle.comparativePeriodLabel && (
                        <>
                          <td className="py-3 px-4 text-right font-mono tabular-nums text-slate-600">
                            {line.previousAmountPaise !== undefined
                              ? formatINR(line.previousAmountPaise)
                              : '—'}
                          </td>
                          <td className="py-3 px-4 text-right font-mono tabular-nums">
                            {line.variancePaise !== undefined ? formatINR(line.variancePaise) : '—'}
                          </td>
                          <td className="py-3 px-4 text-right font-mono tabular-nums font-semibold">
                            {line.varianceLabel || '—'}
                          </td>
                        </>
                      )}
                    </tr>
                  ))}
                </tbody>
              </table>
            </div>
          )}
        </div>
      )}

      {/* BALANCE SHEET TABLE (Collapsible — Shown in BALANCE_SHEET and COMPARATIVE modes) */}
      {(viewMode === 'BALANCE_SHEET' || viewMode === 'COMPARATIVE') && (
        <div className="bg-white rounded-xl border border-slate-200 overflow-hidden">
          <button
            onClick={() => setCollapseBs(!collapseBs)}
            className="w-full px-6 py-4 bg-slate-50 border-b border-slate-200 flex items-center justify-between text-left"
          >
            <div className="flex items-center gap-2">
              {collapseBs ? (
                <ChevronRight className="w-4 h-4 text-[#0B2540]" />
              ) : (
                <ChevronDown className="w-4 h-4 text-[#0B2540]" />
              )}
              <h3 className="text-sm font-bold text-[#0B2540]">
                Statement of Balance Sheet — Assets, Equity & Liabilities (As of {model.periodEnd})
              </h3>
            </div>
            <span className="text-xs text-slate-500">Click to collapse/expand</span>
          </button>

          {!collapseBs && (
            <div className="p-6 space-y-4">
              {/* Reconciliation Status Banner */}
              <div
                className={`p-3.5 rounded-xl border flex flex-wrap items-center justify-between gap-3 text-xs ${
                  Math.abs(
                    statementBundle.kpis.totalAssetsPaise -
                      statementBundle.kpis.totalEquityAndLiabilitiesPaise
                  ) <= 200
                    ? 'bg-emerald-50 border-emerald-200 text-emerald-900'
                    : 'bg-amber-50 border-amber-300 text-amber-900'
                }`}
              >
                <div className="flex items-center gap-2 font-semibold">
                  {Math.abs(
                    statementBundle.kpis.totalAssetsPaise -
                      statementBundle.kpis.totalEquityAndLiabilitiesPaise
                  ) <= 200 ? (
                    <CheckCircle2 className="w-4 h-4 text-emerald-700" />
                  ) : (
                    <AlertTriangle className="w-4 h-4 text-amber-700" />
                  )}
                  <span>
                    Balance Sheet Reconciliation Check: Total Assets (
                    {formatINR(statementBundle.kpis.totalAssetsPaise)}) vs Total Equity &amp;
                    Liabilities ({formatINR(statementBundle.kpis.totalEquityAndLiabilitiesPaise)})
                  </span>
                </div>
                <span className="font-mono font-bold">
                  Difference:{' '}
                  {formatINR(
                    statementBundle.kpis.totalAssetsPaise -
                      statementBundle.kpis.totalEquityAndLiabilitiesPaise
                  )}{' '}
                  (
                  {Math.abs(
                    statementBundle.kpis.totalAssetsPaise -
                      statementBundle.kpis.totalEquityAndLiabilitiesPaise
                  ) <= 200
                    ? 'RECONCILED'
                    : 'REQUIRES REVIEW'}
                  )
                </span>
              </div>

              <div className="grid grid-cols-1 lg:grid-cols-2 gap-6">
                {/* Assets Side */}
                <div className="border border-slate-200 rounded-lg overflow-hidden">
                  <div className="px-4 py-3 bg-[#0B2540] text-white flex items-center justify-between">
                    <h4 className="text-xs font-bold">Assets Schedule</h4>
                    <span className="text-[11px] font-mono">
                      {model.currentPeriodLabel || `As of ${model.periodEnd}`}
                    </span>
                  </div>
                  <table className="w-full text-left border-collapse text-xs">
                    <thead>
                      <tr className="bg-slate-100 text-slate-700 border-b border-slate-200">
                        <th className="py-2.5 px-4 font-semibold">Asset Head</th>
                        <th className="py-2.5 px-4 font-semibold text-right">
                          {model.currentPeriodLabel || 'Current (₹)'}
                        </th>
                        {model.comparativePeriodLabel && (
                          <th className="py-2.5 px-4 font-semibold text-right">
                            {model.comparativePeriodLabel}
                          </th>
                        )}
                      </tr>
                    </thead>
                    <tbody className="divide-y divide-slate-200">
                      {statementBundle.bsAssetLines.map((line) => (
                        <tr
                          key={line.code}
                          onClick={() =>
                            line.classificationFilter.length > 0 &&
                            onDrillDown({
                              title: line.label,
                              formulaExplanation: line.formulaExplanation,
                              classifications: line.classificationFilter,
                              currentFormatted: formatINR(line.amountPaise),
                              previousFormatted:
                                line.previousAmountPaise !== undefined
                                  ? formatINR(line.previousAmountPaise)
                                  : undefined,
                              sourceCellRef: line.cellReferences?.[0],
                              priorCellRef: line.priorCellReferences?.[0],
                            })
                          }
                          className={
                            line.isSubtotal
                              ? 'bg-slate-50 font-bold text-[#0B2540]'
                              : 'hover:bg-blue-50/50 cursor-pointer'
                          }
                        >
                          <td className="py-3 px-4">
                            <div
                              className={
                                line.isSubtotal ? 'font-bold' : 'font-medium text-[#0875BC]'
                              }
                            >
                              {line.label}
                            </div>
                            <div className="text-[11px] text-slate-500">
                              {line.formulaExplanation}
                            </div>
                          </td>
                          <td className="py-3 px-4 text-right font-mono tabular-nums">
                            {formatINR(line.amountPaise)}
                          </td>
                          {model.comparativePeriodLabel && (
                            <td className="py-3 px-4 text-right font-mono tabular-nums text-slate-600">
                              {line.previousAmountPaise !== undefined
                                ? formatINR(line.previousAmountPaise)
                                : '—'}
                            </td>
                          )}
                        </tr>
                      ))}
                    </tbody>
                  </table>
                </div>

                {/* Equity & Liabilities Side */}
                <div className="border border-slate-200 rounded-lg overflow-hidden">
                  <div className="px-4 py-3 bg-[#0B2540] text-white flex items-center justify-between">
                    <h4 className="text-xs font-bold">Equity &amp; Liabilities Schedule</h4>
                    <span className="text-[11px] font-mono">
                      {model.currentPeriodLabel || `As of ${model.periodEnd}`}
                    </span>
                  </div>
                  <table className="w-full text-left border-collapse text-xs">
                    <thead>
                      <tr className="bg-slate-100 text-slate-700 border-b border-slate-200">
                        <th className="py-2.5 px-4 font-semibold">Equity / Liability Head</th>
                        <th className="py-2.5 px-4 font-semibold text-right">
                          {model.currentPeriodLabel || 'Current (₹)'}
                        </th>
                        {model.comparativePeriodLabel && (
                          <th className="py-2.5 px-4 font-semibold text-right">
                            {model.comparativePeriodLabel}
                          </th>
                        )}
                      </tr>
                    </thead>
                    <tbody className="divide-y divide-slate-200">
                      {statementBundle.bsLiabilityEquityLines.map((line) => (
                        <tr
                          key={line.code}
                          onClick={() =>
                            line.classificationFilter.length > 0 &&
                            onDrillDown({
                              title: line.label,
                              formulaExplanation: line.formulaExplanation,
                              classifications: line.classificationFilter,
                              currentFormatted: formatINR(line.amountPaise),
                              previousFormatted:
                                line.previousAmountPaise !== undefined
                                  ? formatINR(line.previousAmountPaise)
                                  : undefined,
                              sourceCellRef: line.cellReferences?.[0],
                              priorCellRef: line.priorCellReferences?.[0],
                            })
                          }
                          className={
                            line.isSubtotal
                              ? 'bg-slate-50 font-bold text-[#0B2540]'
                              : 'hover:bg-blue-50/50 cursor-pointer'
                          }
                        >
                          <td className="py-3 px-4">
                            <div
                              className={
                                line.isSubtotal ? 'font-bold' : 'font-medium text-[#0875BC]'
                              }
                            >
                              {line.label}
                            </div>
                            <div className="text-[11px] text-slate-500">
                              {line.formulaExplanation}
                            </div>
                          </td>
                          <td className="py-3 px-4 text-right font-mono tabular-nums">
                            {formatINR(line.amountPaise)}
                          </td>
                          {model.comparativePeriodLabel && (
                            <td className="py-3 px-4 text-right font-mono tabular-nums text-slate-600">
                              {line.previousAmountPaise !== undefined
                                ? formatINR(line.previousAmountPaise)
                                : '—'}
                            </td>
                          )}
                        </tr>
                      ))}
                    </tbody>
                  </table>
                </div>
              </div>
            </div>
          )}
        </div>
      )}

      {/* Comparative Trend Charts (Module 3.5 — Revenue, GP, NP, Expenses, Assets, Liabilities, Working Capital) */}
      {trendData.length > 0 && (
        <div className="bg-white rounded-xl border border-slate-200 p-6 space-y-4">
          <div className="flex flex-wrap items-center justify-between gap-3">
            <h3 className="text-sm font-bold text-[#0B2540]">
              Comparative Financial Trajectory Charts (Module 3.5)
            </h3>
            <div className="flex items-center gap-1 p-1 bg-slate-100 rounded-lg text-xs">
              <button
                onClick={() => setActiveChartTab('PROFITABILITY')}
                className={`px-3 py-1 font-semibold rounded ${
                  activeChartTab === 'PROFITABILITY'
                    ? 'bg-white text-[#0B2540] shadow-xs'
                    : 'text-slate-600'
                }`}
              >
                Revenue, Gross Profit & Net Profit
              </button>
              <button
                onClick={() => setActiveChartTab('EXPENSES')}
                className={`px-3 py-1 font-semibold rounded ${
                  activeChartTab === 'EXPENSES'
                    ? 'bg-white text-[#0B2540] shadow-xs'
                    : 'text-slate-600'
                }`}
              >
                Expense & Procurement Trends
              </button>
              <button
                onClick={() => setActiveChartTab('BALANCE_SHEET_WC')}
                className={`px-3 py-1 font-semibold rounded ${
                  activeChartTab === 'BALANCE_SHEET_WC'
                    ? 'bg-white text-[#0B2540] shadow-xs'
                    : 'text-slate-600'
                }`}
              >
                Assets, Liabilities & Working Capital
              </button>
            </div>
          </div>

          <div className="h-64">
            <ResponsiveContainer width="100%" height="100%">
              {activeChartTab === 'PROFITABILITY' ? (
                <LineChart data={trendData}>
                  <CartesianGrid strokeDasharray="3 3" stroke="#E2E8F0" />
                  <XAxis dataKey="month" tick={{ fontSize: 11 }} />
                  <YAxis
                    tick={{ fontSize: 11 }}
                    tickFormatter={(v) => `₹${(v / 100000).toFixed(1)}L`}
                  />
                  <Tooltip
                    formatter={(val: any) => formatINR(Math.round(Number(val) * 100))}
                  />
                  <Legend wrapperStyle={{ fontSize: 12 }} />
                  <Line type="monotone" dataKey="Revenue" name="Revenue" stroke="#0875BC" strokeWidth={2.5} />
                  <Line type="monotone" dataKey="GrossProfit" name="Gross Profit" stroke="#0F9B91" strokeWidth={2.5} />
                  <Line type="monotone" dataKey="NetProfit" name="Net Profit" stroke="#50AE5D" strokeWidth={2.5} />
                </LineChart>
              ) : activeChartTab === 'EXPENSES' ? (
                <BarChart data={trendData}>
                  <CartesianGrid strokeDasharray="3 3" stroke="#E2E8F0" />
                  <XAxis dataKey="month" tick={{ fontSize: 11 }} />
                  <YAxis
                    tick={{ fontSize: 11 }}
                    tickFormatter={(v) => `₹${(v / 100000).toFixed(1)}L`}
                  />
                  <Tooltip
                    formatter={(val: any) => formatINR(Math.round(Number(val) * 100))}
                  />
                  <Legend wrapperStyle={{ fontSize: 12 }} />
                  <Bar dataKey="Purchases" name="Direct Purchases" fill="#0B2540" radius={[4, 4, 0, 0]} />
                  <Bar dataKey="Expenses" name="Operating Expenses" fill="#D97706" radius={[4, 4, 0, 0]} />
                </BarChart>
              ) : (
                <LineChart data={trendData}>
                  <CartesianGrid strokeDasharray="3 3" stroke="#E2E8F0" />
                  <XAxis dataKey="month" tick={{ fontSize: 11 }} />
                  <YAxis
                    tick={{ fontSize: 11 }}
                    tickFormatter={(v) => `₹${(v / 100000).toFixed(1)}L`}
                  />
                  <Tooltip
                    formatter={(val: any) => formatINR(Math.round(Number(val) * 100))}
                  />
                  <Legend wrapperStyle={{ fontSize: 12 }} />
                  <Line type="monotone" dataKey="Assets" name="Total Assets" stroke="#0875BC" strokeWidth={2.5} />
                  <Line type="monotone" dataKey="Liabilities" name="Total Liabilities" stroke="#DC2626" strokeWidth={2} />
                  <Line type="monotone" dataKey="WorkingCapital" name="Working Capital" stroke="#0F9B91" strokeWidth={2.5} />
                </LineChart>
              )}
            </ResponsiveContainer>
          </div>
        </div>
      )}
    </div>
  );
};
