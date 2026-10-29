import React, { useMemo, useState } from 'react';
import * as XLSX from 'xlsx';
import jsPDF from 'jspdf';
import autoTable from 'jspdf-autotable';
import {
  Bar,
  BarChart,
  CartesianGrid,
  Legend,
  Line,
  LineChart,
  ResponsiveContainer,
  Tooltip,
  XAxis,
  YAxis,
} from 'recharts';
import {
  AlertTriangle,
  Download,
  Eye,
  Search,
  X,
} from 'lucide-react';
import {
  AuditException,
  AuditReviewState,
  ConsolidatedAccountingModel,
  FinancialRatioResult,
  formatINR,
  formatINRCompact,
  fromPaise,
  PolicySettings,
} from '../types/accounting';
import { AgeingAnalysisResult, PartyAgeingSummary } from '../engine/ageingEngine';
import { AccountingHealthCheckResult } from '../engine/auditEngine';

// ============================================================================
// MODULE 2 — RECEIVABLES & PAYABLES ANALYSIS (DEBTORS & CREDITORS AGEING)
// ============================================================================

interface AgeingAnalysisSectionProps {
  model: ConsolidatedAccountingModel;
  ageing: AgeingAnalysisResult;
  monthlyCollectionsTrends?: Array<{
    month: string;
    collectionsPaise: number;
    paymentsPaise: number;
  }>;
  policy: PolicySettings;
  onUpdatePolicy: (next: PolicySettings) => void;
  onOpenVoucherDrillDown: (voucherId: string) => void;
}

export const AgeingAnalysisSection: React.FC<AgeingAnalysisSectionProps> = ({
  model,
  ageing,
  monthlyCollectionsTrends = [],
  policy,
  onUpdatePolicy,
  onOpenVoucherDrillDown,
}) => {
  const [selectedParty, setSelectedParty] = useState<PartyAgeingSummary | null>(null);
  const [activeTab, setActiveTab] = useState<'CUSTOMERS' | 'SUPPLIERS'>('CUSTOMERS');
  const [searchParty, setSearchParty] = useState('');
  const [chartView, setChartView] = useState<
    'BUCKETS' | 'TOP_PARTIES' | 'COLLECTIONS_PAYMENTS'
  >('BUCKETS');

  const bucketChartData = ageing.bucketSummaries.map((b) => ({
    bucket: b.label,
    Receivables: fromPaise(b.receivablesPaise),
    Payables: fromPaise(b.payablesPaise),
  }));

  const topOverdueCustomersData = ageing.customers.slice(0, 10).map((c) => ({
    name: c.partyName.slice(0, 18),
    Overdue: fromPaise(c.overdueOutstandingPaise),
    Current: fromPaise(c.currentNotDuePaise),
  }));

  const topSuppliersData = ageing.suppliers.slice(0, 10).map((s) => ({
    name: s.partyName.slice(0, 18),
    TotalPayable: fromPaise(s.totalOutstandingPaise),
    Overdue: fromPaise(s.overdueOutstandingPaise),
  }));

  const collectionsChartData = monthlyCollectionsTrends.map((m) => ({
    month: m.month,
    Collections: fromPaise(m.collectionsPaise),
    Payments: fromPaise(m.paymentsPaise),
  }));

  const exportAgeingExcel = (category: 'RECEIVABLE' | 'PAYABLE' | 'OVERDUE_LIST') => {
    const wb = XLSX.utils.book_new();
    if (category === 'OVERDUE_LIST') {
      const overdueInvoices = model.outstandingInvoices.filter((i) => i.daysOverdue > 0);
      const rows = [
        ['Category', 'Party Name', 'Bill Number', 'Bill Date', 'Due Date', 'Days Overdue', 'Pending Amount (INR)', 'Source File'],
        ...overdueInvoices.map((inv) => [
          inv.category,
          inv.partyName,
          inv.billNumber,
          inv.billDate,
          inv.dueDate,
          inv.daysOverdue,
          fromPaise(inv.pendingAmountPaise),
          inv.sourceFilename,
        ]),
      ];
      XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet(rows), 'Overdue_Collection_List');
      XLSX.writeFile(wb, 'GSPU_Overdue_Collection_List.xlsx');
      return;
    }

    const list = category === 'RECEIVABLE' ? ageing.customers : ageing.suppliers;
    const rows = [
      [
        category === 'RECEIVABLE' ? 'Customer Name' : 'Supplier Name',
        'Total Outstanding (INR)',
        'Overdue Outstanding (INR)',
        'Current / Not Due (INR)',
        'Oldest Invoice Date',
        'Open Invoices',
        'Max Days Overdue',
        '% Share',
        'Not Due',
        '1-30 Days',
        '31-60 Days',
        '61-90 Days',
        '91-180 Days',
        'Over 180 Days',
      ],
      ...list.map((p) => [
        p.partyName,
        fromPaise(p.totalOutstandingPaise),
        fromPaise(p.overdueOutstandingPaise),
        fromPaise(p.currentNotDuePaise),
        p.oldestInvoiceDate,
        p.openInvoiceCount,
        p.maxDaysOverdue,
        `${p.sharePercent.toFixed(1)}%`,
        fromPaise(p.buckets.NOT_DUE),
        fromPaise(p.buckets.DAYS_1_30),
        fromPaise(p.buckets.DAYS_31_60),
        fromPaise(p.buckets.DAYS_61_90),
        fromPaise(p.buckets.DAYS_91_180),
        fromPaise(p.buckets.OVER_180),
      ]),
    ];
    XLSX.utils.book_append_sheet(
      wb,
      XLSX.utils.aoa_to_sheet(rows),
      category === 'RECEIVABLE' ? 'Customer_Ageing' : 'Supplier_Ageing'
    );
    XLSX.writeFile(
      wb,
      `GSPU_${category === 'RECEIVABLE' ? 'Customer' : 'Supplier'}_Ageing.xlsx`
    );
  };

  const exportAgeingPDF = (category: 'RECEIVABLE' | 'PAYABLE') => {
    const doc = new jsPDF();
    const isRec = category === 'RECEIVABLE';
    doc.setFontSize(13);
    doc.text(
      `GSPU Financial Analytics — ${isRec ? 'Receivables (Debtors)' : 'Payables (Creditors)'} Summary`,
      14,
      18
    );
    doc.setFontSize(9);
    doc.text(`Company: ${model.companyName} | Period: ${model.periodStart} to ${model.periodEnd}`, 14, 25);

    const list = isRec ? ageing.customers : ageing.suppliers;
    autoTable(doc, {
      startY: 30,
      head: [['Party Name', 'Total Outstanding', 'Overdue', 'Not Due', 'Open Bills', 'Max Overdue']],
      body: list.map((p) => [
        p.partyName,
        formatINR(p.totalOutstandingPaise),
        formatINR(p.overdueOutstandingPaise),
        formatINR(p.currentNotDuePaise),
        String(p.openInvoiceCount),
        `${p.maxDaysOverdue} days`,
      ]),
      headStyles: { fillColor: isRec ? [8, 117, 188] : [15, 155, 145] },
    });

    doc.save(`GSPU_${isRec ? 'Receivables' : 'Payables'}_Summary.pdf`);
  };

  const filteredParties = (activeTab === 'CUSTOMERS' ? ageing.customers : ageing.suppliers).filter(
    (p) => !searchParty || p.partyName.toLowerCase().includes(searchParty.toLowerCase())
  );

  return (
    <div className="space-y-6">
      {/* Header & Ageing Basis Toggle */}
      <div className="bg-white rounded-xl border border-slate-200 p-5 flex flex-wrap items-center justify-between gap-4">
        <div>
          <h2 className="text-lg font-bold text-[#0B2540]">
            RECEIVABLES & PAYABLES ANALYSIS (Module 2)
          </h2>
          <p className="text-xs text-slate-500 mt-0.5">
            Deterministic bill-allocation ageing engine. Never infers invoice ageing from ledger closing balances alone.
          </p>
        </div>

        <div className="flex flex-wrap items-center gap-2">
          <div className="flex items-center gap-1 p-1 bg-slate-100 rounded-lg text-xs">
            <button
              onClick={() => onUpdatePolicy({ ...policy, ageingMode: 'DUE_DATE' })}
              className={`px-3 py-1.5 font-semibold rounded-md transition-colors ${
                policy.ageingMode === 'DUE_DATE'
                  ? 'bg-white text-[#0B2540] shadow-xs'
                  : 'text-slate-600'
              }`}
            >
              Due-Date Based Ageing
            </button>
            <button
              onClick={() => onUpdatePolicy({ ...policy, ageingMode: 'INVOICE_DATE' })}
              className={`px-3 py-1.5 font-semibold rounded-md transition-colors ${
                policy.ageingMode === 'INVOICE_DATE'
                  ? 'bg-white text-[#0B2540] shadow-xs'
                  : 'text-slate-600'
              }`}
            >
              Invoice-Age Based Ageing
            </button>
          </div>

          <button
            onClick={() => exportAgeingExcel('RECEIVABLE')}
            className="px-2.5 py-1.5 text-xs font-semibold text-white bg-[#0875BC] hover:bg-[#065d96] rounded-lg flex items-center gap-1"
          >
            <Download className="w-3.5 h-3.5" />
            Customer Excel
          </button>
          <button
            onClick={() => exportAgeingExcel('PAYABLE')}
            className="px-2.5 py-1.5 text-xs font-semibold text-white bg-[#0F9B91] hover:bg-[#0c8078] rounded-lg flex items-center gap-1"
          >
            <Download className="w-3.5 h-3.5" />
            Supplier Excel
          </button>
          <button
            onClick={() => exportAgeingExcel('OVERDUE_LIST')}
            className="px-2.5 py-1.5 text-xs font-semibold text-red-700 bg-red-50 hover:bg-red-100 rounded-lg flex items-center gap-1"
          >
            <Download className="w-3.5 h-3.5" />
            Overdue Collection List
          </button>
          <button
            onClick={() => exportAgeingPDF('RECEIVABLE')}
            className="px-2.5 py-1.5 text-xs font-semibold text-[#0B2540] bg-slate-100 hover:bg-slate-200 rounded-lg"
          >
            Receivables PDF
          </button>
          <button
            onClick={() => exportAgeingPDF('PAYABLE')}
            className="px-2.5 py-1.5 text-xs font-semibold text-[#0B2540] bg-slate-100 hover:bg-slate-200 rounded-lg"
          >
            Payables PDF
          </button>
        </div>
      </div>

      {/* 8 Summary Cards (Module 2.1) */}
      <div className="grid grid-cols-2 sm:grid-cols-4 lg:grid-cols-8 gap-3">
        {[
          {
            label: 'Total Trade Receivables',
            val: formatINRCompact(
              ageing.totalBillTrackedReceivablesPaise || ageing.totalTradeReceivablesLedgerPaise
            ),
          },
          {
            label: 'Total Trade Payables',
            val: formatINRCompact(
              ageing.totalBillTrackedPayablesPaise || ageing.totalTradePayablesLedgerPaise
            ),
          },
          {
            label: 'Overdue Receivables',
            val: formatINRCompact(ageing.overdueReceivablesPaise),
          },
          {
            label: 'Overdue Payables',
            val: formatINRCompact(ageing.overduePayablesPaise),
          },
          {
            label: 'Customers Overdue',
            val: `${ageing.customersWithOverdueCount} of ${ageing.customers.length}`,
          },
          {
            label: 'Suppliers Overdue',
            val: `${ageing.suppliersWithOverdueCount} of ${ageing.suppliers.length}`,
          },
          {
            label: 'Avg Collection Period',
            val:
              ageing.averageCollectionPeriodDays !== null
                ? `${ageing.averageCollectionPeriodDays} Days`
                : 'N/A',
          },
          {
            label: 'Avg Payment Period',
            val:
              ageing.averagePaymentPeriodDays !== null
                ? `${ageing.averagePaymentPeriodDays} Days`
                : 'N/A',
          },
        ].map((c) => (
          <div key={c.label} className="bg-white rounded-xl border border-slate-200 p-3.5">
            <div className="text-[11px] font-medium text-slate-500">{c.label}</div>
            <div className="text-base font-bold font-mono tabular-nums text-[#0B2540] mt-1">
              {c.val}
            </div>
          </div>
        ))}
      </div>

      {/* Insufficient Bill-Level Data Guard */}
      {!ageing.hasBillLevelData && (
        <div className="p-4 rounded-xl bg-amber-50 border border-amber-200 flex items-start gap-3 text-xs text-amber-900">
          <AlertTriangle className="w-5 h-5 text-amber-600 shrink-0 mt-0.5" />
          <div>
            <div className="font-bold">
              Bill-Level Allocation Data Not Present in Current Upload
            </div>
            <p className="mt-1">
              Ledger closing balances exist (Receivables: {formatINR(ageing.totalTradeReceivablesLedgerPaise)}, Payables: {formatINR(ageing.totalTradePayablesLedgerPaise)}), but per accounting policy, invoice-level ageing buckets are never fabricated from ledger closing balances alone. Upload a Tally Transactions XML containing <code>BILLALLOCATIONS.LIST</code> or a Tally Outstanding report to populate invoice ageing.
            </p>
          </div>
        </div>
      )}

      {/* Interactive Ageing Charts (Module 2.5 — All 6 Chart Views) */}
      {ageing.hasBillLevelData && (
        <div className="bg-white rounded-xl border border-slate-200 p-6 space-y-4">
          <div className="flex flex-wrap items-center justify-between gap-3">
            <h3 className="text-sm font-bold text-[#0B2540]">
              Interactive Ageing & Cash Collection Analytics (Module 2.5)
            </h3>
            <div className="flex items-center gap-1 p-1 bg-slate-100 rounded-lg text-xs">
              <button
                onClick={() => setChartView('BUCKETS')}
                className={`px-3 py-1 font-semibold rounded ${
                  chartView === 'BUCKETS'
                    ? 'bg-white text-[#0B2540] shadow-xs'
                    : 'text-slate-600'
                }`}
              >
                Receivables & Payables Ageing Buckets
              </button>
              <button
                onClick={() => setChartView('TOP_PARTIES')}
                className={`px-3 py-1 font-semibold rounded ${
                  chartView === 'TOP_PARTIES'
                    ? 'bg-white text-[#0B2540] shadow-xs'
                    : 'text-slate-600'
                }`}
              >
                Top 10 Customers & Suppliers
              </button>
              <button
                onClick={() => setChartView('COLLECTIONS_PAYMENTS')}
                className={`px-3 py-1 font-semibold rounded ${
                  chartView === 'COLLECTIONS_PAYMENTS'
                    ? 'bg-white text-[#0B2540] shadow-xs'
                    : 'text-slate-600'
                }`}
              >
                Monthly Collections & Payments Trend
              </button>
            </div>
          </div>

          {chartView === 'BUCKETS' && (
            <div className="grid grid-cols-1 lg:grid-cols-2 gap-6">
              <div className="h-64">
                <ResponsiveContainer width="100%" height="100%">
                  <BarChart data={bucketChartData}>
                    <CartesianGrid strokeDasharray="3 3" stroke="#E2E8F0" />
                    <XAxis dataKey="bucket" tick={{ fontSize: 11 }} />
                    <YAxis
                      tick={{ fontSize: 11 }}
                      tickFormatter={(v) => `₹${(v / 100000).toFixed(1)}L`}
                    />
                    <Tooltip
                      formatter={(val: any) => formatINR(Math.round(Number(val) * 100))}
                    />
                    <Legend wrapperStyle={{ fontSize: 12 }} />
                    <Bar dataKey="Receivables" name="Receivables Ageing" fill="#0875BC" radius={[4, 4, 0, 0]} />
                    <Bar dataKey="Payables" name="Payables Ageing" fill="#0F9B91" radius={[4, 4, 0, 0]} />
                  </BarChart>
                </ResponsiveContainer>
              </div>

              <div className="border border-slate-200 rounded-lg overflow-hidden">
                <table className="w-full text-left border-collapse text-xs">
                  <thead>
                    <tr className="bg-slate-100 border-b border-slate-200 text-slate-700">
                      <th className="py-2 px-3 font-semibold">Ageing Bucket</th>
                      <th className="py-2 px-3 font-semibold text-right">Receivables (₹)</th>
                      <th className="py-2 px-3 font-semibold text-right">Bills</th>
                      <th className="py-2 px-3 font-semibold text-right">Payables (₹)</th>
                      <th className="py-2 px-3 font-semibold text-right">Bills</th>
                    </tr>
                  </thead>
                  <tbody className="divide-y divide-slate-200">
                    {ageing.bucketSummaries.map((b) => (
                      <tr key={b.key} className="hover:bg-slate-50">
                        <td className="py-2 px-3 font-semibold text-[#0B2540]">{b.label}</td>
                        <td className="py-2 px-3 text-right font-mono tabular-nums">
                          {formatINR(b.receivablesPaise)}
                        </td>
                        <td className="py-2 px-3 text-right font-mono">{b.receivablesCount}</td>
                        <td className="py-2 px-3 text-right font-mono tabular-nums">
                          {formatINR(b.payablesPaise)}
                        </td>
                        <td className="py-2 px-3 text-right font-mono">{b.payablesCount}</td>
                      </tr>
                    ))}
                  </tbody>
                </table>
              </div>
            </div>
          )}

          {chartView === 'TOP_PARTIES' && (
            <div className="grid grid-cols-1 lg:grid-cols-2 gap-6">
              <div className="h-64">
                <div className="text-xs font-semibold text-slate-600 mb-2">
                  Top 10 Customers by Overdue & Current Receivables
                </div>
                <ResponsiveContainer width="100%" height="90%">
                  <BarChart data={topOverdueCustomersData}>
                    <CartesianGrid strokeDasharray="3 3" stroke="#E2E8F0" />
                    <XAxis dataKey="name" tick={{ fontSize: 10 }} />
                    <YAxis
                      tick={{ fontSize: 11 }}
                      tickFormatter={(v) => `₹${(v / 100000).toFixed(1)}L`}
                    />
                    <Tooltip
                      formatter={(val: any) => formatINR(Math.round(Number(val) * 100))}
                    />
                    <Legend wrapperStyle={{ fontSize: 11 }} />
                    <Bar dataKey="Overdue" name="Overdue" fill="#DC2626" />
                    <Bar dataKey="Current" name="Not Due" fill="#0875BC" />
                  </BarChart>
                </ResponsiveContainer>
              </div>

              <div className="h-64">
                <div className="text-xs font-semibold text-slate-600 mb-2">
                  Top 10 Suppliers by Outstanding Payables
                </div>
                <ResponsiveContainer width="100%" height="90%">
                  <BarChart data={topSuppliersData}>
                    <CartesianGrid strokeDasharray="3 3" stroke="#E2E8F0" />
                    <XAxis dataKey="name" tick={{ fontSize: 10 }} />
                    <YAxis
                      tick={{ fontSize: 11 }}
                      tickFormatter={(v) => `₹${(v / 100000).toFixed(1)}L`}
                    />
                    <Tooltip
                      formatter={(val: any) => formatINR(Math.round(Number(val) * 100))}
                    />
                    <Legend wrapperStyle={{ fontSize: 11 }} />
                    <Bar dataKey="TotalPayable" name="Total Payable" fill="#0F9B91" />
                    <Bar dataKey="Overdue" name="Overdue Payable" fill="#D97706" />
                  </BarChart>
                </ResponsiveContainer>
              </div>
            </div>
          )}

          {chartView === 'COLLECTIONS_PAYMENTS' && (
            <div className="h-64">
              <ResponsiveContainer width="100%" height="100%">
                <LineChart data={collectionsChartData}>
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
                  <Line
                    type="monotone"
                    dataKey="Collections"
                    name="Monthly Customer Collections (Receipts)"
                    stroke="#50AE5D"
                    strokeWidth={2.5}
                  />
                  <Line
                    type="monotone"
                    dataKey="Payments"
                    name="Monthly Supplier & Expense Payments"
                    stroke="#0875BC"
                    strokeWidth={2.5}
                  />
                </LineChart>
              </ResponsiveContainer>
            </div>
          )}
        </div>
      )}

      {/* Customer & Supplier Analysis Table (Module 2.3 & 2.4) */}
      <div className="bg-white rounded-xl border border-slate-200 p-6">
        <div className="flex flex-wrap items-center justify-between gap-4 mb-4">
          <div className="flex items-center gap-1 p-1 bg-slate-100 rounded-lg text-xs">
            <button
              onClick={() => setActiveTab('CUSTOMERS')}
              className={`px-4 py-1.5 font-semibold rounded-md transition-colors ${
                activeTab === 'CUSTOMERS'
                  ? 'bg-white text-[#0B2540] shadow-xs'
                  : 'text-slate-600'
              }`}
            >
              Customer Analysis (Debtors: {ageing.customers.length})
            </button>
            <button
              onClick={() => setActiveTab('SUPPLIERS')}
              className={`px-4 py-1.5 font-semibold rounded-md transition-colors ${
                activeTab === 'SUPPLIERS'
                  ? 'bg-white text-[#0B2540] shadow-xs'
                  : 'text-slate-600'
              }`}
            >
              Supplier Analysis (Creditors: {ageing.suppliers.length})
            </button>
          </div>

          <div className="relative">
            <Search className="w-3.5 h-3.5 text-slate-400 absolute left-2.5 top-1/2 -translate-y-1/2" />
            <input
              type="text"
              value={searchParty}
              onChange={(e) => setSearchParty(e.target.value)}
              placeholder="Search customer / supplier..."
              className="pl-8 pr-3 py-1.5 text-xs border border-slate-300 rounded-lg"
            />
          </div>
        </div>

        <div className="border border-slate-200 rounded-lg overflow-x-auto">
          <table className="w-full text-left border-collapse text-xs">
            <thead>
              <tr className="bg-slate-100 text-slate-700 border-b border-slate-200">
                <th className="py-2.5 px-3 font-semibold">Party Name</th>
                <th className="py-2.5 px-3 font-semibold text-right">Total Outstanding</th>
                <th className="py-2.5 px-3 font-semibold text-right">Overdue Amount</th>
                <th className="py-2.5 px-3 font-semibold text-right">Current (Not Due)</th>
                <th className="py-2.5 px-3 font-semibold">Oldest Unpaid Bill</th>
                <th className="py-2.5 px-3 font-semibold text-right">Open Bills</th>
                <th className="py-2.5 px-3 font-semibold text-right">Max Days Overdue</th>
                <th className="py-2.5 px-3 font-semibold text-right">% Share</th>
                <th className="py-2.5 px-3 font-semibold text-right">Drill-Down</th>
              </tr>
            </thead>
            <tbody className="divide-y divide-slate-200">
              {filteredParties.map((p) => (
                <tr
                  key={p.normalizedPartyName}
                  onClick={() => setSelectedParty(p)}
                  className="hover:bg-blue-50/50 cursor-pointer"
                >
                  <td className="py-3 px-3 font-semibold text-[#0875BC]">{p.partyName}</td>
                  <td className="py-3 px-3 text-right font-mono tabular-nums font-bold text-[#0B2540]">
                    {formatINR(p.totalOutstandingPaise)}
                  </td>
                  <td className="py-3 px-3 text-right font-mono tabular-nums text-red-600 font-semibold">
                    {formatINR(p.overdueOutstandingPaise)}
                  </td>
                  <td className="py-3 px-3 text-right font-mono tabular-nums text-emerald-700">
                    {formatINR(p.currentNotDuePaise)}
                  </td>
                  <td className="py-3 px-3 font-mono">{p.oldestInvoiceDate}</td>
                  <td className="py-3 px-3 text-right font-mono">{p.openInvoiceCount}</td>
                  <td className="py-3 px-3 text-right font-mono">{p.maxDaysOverdue} days</td>
                  <td className="py-3 px-3 text-right font-mono">{p.sharePercent.toFixed(1)}%</td>
                  <td className="py-3 px-3 text-right">
                    <button className="text-[#0875BC] hover:underline font-medium">
                      View Bills →
                    </button>
                  </td>
                </tr>
              ))}
              {filteredParties.length === 0 && (
                <tr>
                  <td colSpan={9} className="py-8 text-center text-slate-500">
                    No bill-level party records match the current view.
                  </td>
                </tr>
              )}
            </tbody>
          </table>
        </div>
      </div>

      {/* Party -> Outstanding Invoices -> Bill Allocations -> Voucher Drill-Down Modal (Module 2.6) */}
      {selectedParty && (
        <div className="fixed inset-0 z-50 flex items-center justify-center bg-slate-900/50 p-4">
          <div className="bg-white rounded-xl max-w-4xl w-full max-h-[85vh] flex flex-col shadow-xl overflow-hidden">
            <div className="px-6 py-4 bg-[#0B2540] text-white flex items-center justify-between">
              <div>
                <h3 className="text-base font-bold">
                  Party Bill-Allocation Drill-Down — {selectedParty.partyName}
                </h3>
                <p className="text-xs text-slate-300">
                  Total Pending: {formatINR(selectedParty.totalOutstandingPaise)} · Overdue:{' '}
                  {formatINR(selectedParty.overdueOutstandingPaise)} · {selectedParty.openInvoiceCount} Open Bill(s)
                </p>
              </div>
              <button
                onClick={() => setSelectedParty(null)}
                className="p-1 rounded bg-white/10 hover:bg-white/20"
              >
                <X className="w-5 h-5" />
              </button>
            </div>

            <div className="p-6 overflow-y-auto space-y-4">
              <div className="border border-slate-200 rounded-lg overflow-x-auto">
                <table className="w-full text-left border-collapse text-xs">
                  <thead>
                    <tr className="bg-slate-100 border-b border-slate-200">
                      <th className="py-2.5 px-3 font-semibold">Bill / Invoice Ref</th>
                      <th className="py-2.5 px-3 font-semibold">Invoice Date</th>
                      <th className="py-2.5 px-3 font-semibold">Due Date</th>
                      <th className="py-2.5 px-3 font-semibold text-right">Original Amount</th>
                      <th className="py-2.5 px-3 font-semibold text-right">Adjustments (Agst Ref)</th>
                      <th className="py-2.5 px-3 font-semibold text-right">Residual Pending</th>
                      <th className="py-2.5 px-3 font-semibold text-right">Days Overdue</th>
                      <th className="py-2.5 px-3 font-semibold text-right">Related Vouchers</th>
                    </tr>
                  </thead>
                  <tbody className="divide-y divide-slate-200">
                    {selectedParty.invoices.map((inv) => (
                      <tr key={inv.id} className="hover:bg-slate-50">
                        <td className="py-2.5 px-3 font-mono font-semibold text-[#0B2540]">
                          {inv.billNumber}
                        </td>
                        <td className="py-2.5 px-3 font-mono">{inv.billDate}</td>
                        <td className="py-2.5 px-3 font-mono">{inv.dueDate}</td>
                        <td className="py-2.5 px-3 text-right font-mono tabular-nums">
                          {formatINR(inv.originalAmountPaise)}
                        </td>
                        <td className="py-2.5 px-3 text-right font-mono tabular-nums text-emerald-700">
                          {formatINR(inv.adjustedAmountPaise)}
                        </td>
                        <td className="py-2.5 px-3 text-right font-mono tabular-nums font-bold text-[#0B2540]">
                          {formatINR(inv.pendingAmountPaise)}
                        </td>
                        <td className="py-2.5 px-3 text-right font-mono">
                          {inv.daysOverdue > 0 ? `${inv.daysOverdue}d` : 'Not Due'}
                        </td>
                        <td className="py-2.5 px-3 text-right space-x-1">
                          {inv.relatedVoucherIds.map((vid, idx) => (
                            <button
                              key={vid}
                              onClick={() => {
                                setSelectedParty(null);
                                onOpenVoucherDrillDown(vid);
                              }}
                              className="text-[#0875BC] hover:underline font-mono"
                            >
                              Vch #{idx + 1}
                            </button>
                          ))}
                        </td>
                      </tr>
                    ))}
                  </tbody>
                </table>
              </div>
            </div>
          </div>
        </div>
      )}
    </div>
  );
};

// ============================================================================
// MODULE 6 — FINANCIAL RATIO ANALYSIS SECTION
// ============================================================================

interface FinancialRatiosSectionProps {
  ratios: FinancialRatioResult[];
  policy: PolicySettings;
  onUpdatePolicy: (next: PolicySettings) => void;
}

export const FinancialRatiosSection: React.FC<FinancialRatiosSectionProps> = ({
  ratios,
  policy,
  onUpdatePolicy,
}) => {
  const [selectedCat, setSelectedCat] = useState<
    'ALL' | 'LIQUIDITY' | 'PROFITABILITY' | 'EFFICIENCY' | 'SOLVENCY' | 'CASH FLOW'
  >('ALL');
  const [dscrInputRupees, setDscrInputRupees] = useState(
    policy.debtServiceAnnualPaise ? String(fromPaise(policy.debtServiceAnnualPaise)) : ''
  );

  const visibleRatios = ratios.filter(
    (r) => selectedCat === 'ALL' || r.category === selectedCat
  );

  return (
    <div className="space-y-6">
      <div className="bg-white rounded-xl border border-slate-200 p-5 flex flex-wrap items-center justify-between gap-4">
        <div>
          <h2 className="text-lg font-bold text-[#0B2540]">
            EXTENDED FINANCIAL RATIO ANALYSIS (5 Categories — {ratios.length} Ratios)
          </h2>
          <p className="text-xs text-slate-500 mt-0.5">
            Liquidity, Profitability, Efficiency, Solvency, and Cash Flow ratios with transparent numerator, denominator, and formula documentation.
          </p>
        </div>

        <div className="flex flex-wrap items-center gap-3">
          <div className="flex flex-wrap items-center gap-1 p-1 bg-slate-100 rounded-lg text-xs">
            {(['ALL', 'LIQUIDITY', 'PROFITABILITY', 'EFFICIENCY', 'SOLVENCY', 'CASH FLOW'] as const).map(
              (cat) => (
                <button
                  key={cat}
                  onClick={() => setSelectedCat(cat)}
                  className={`px-3 py-1.5 font-semibold rounded-md transition-colors cursor-pointer ${
                    selectedCat === cat
                      ? 'bg-white text-[#0B2540] shadow-xs'
                      : 'text-slate-600'
                  }`}
                >
                  {cat === 'ALL' ? `All ${ratios.length} Ratios` : cat}
                </button>
              )
            )}
          </div>
        </div>
      </div>

      {/* Optional Debt-Service Schedule Configuration for DSCR (Module 6.4) */}
      <div className="bg-white rounded-xl border border-slate-200 p-4 flex flex-wrap items-center justify-between gap-4 text-xs">
        <div className="text-slate-700">
          <strong className="text-[#0B2540]">Debt Service Coverage Ratio (DSCR) Policy Control:</strong>{' '}
          Per accounting standards, DSCR is never derived from balance sheet totals alone. Supply scheduled annual principal + interest repayments (₹) to enable DSCR:
        </div>
        <div className="flex items-center gap-2">
          <input
            type="number"
            value={dscrInputRupees}
            onChange={(e) => setDscrInputRupees(e.target.value)}
            placeholder="Annual Debt Service (₹)"
            className="px-3 py-1.5 border border-slate-300 rounded-lg font-mono w-48"
          />
          <button
            onClick={() => {
              const val = Number(dscrInputRupees);
              onUpdatePolicy({
                ...policy,
                debtServiceAnnualPaise: val > 0 ? Math.round(val * 100) : undefined,
              });
            }}
            className="px-3.5 py-1.5 font-semibold text-white bg-[#0875BC] rounded-lg"
          >
            Apply Schedule
          </button>
        </div>
      </div>

      {/* Ratio Cards Grid */}
      <div className="grid grid-cols-1 md:grid-cols-2 gap-4">
        {visibleRatios.map((r) => (
          <div
            key={r.id}
            className="bg-white rounded-xl border border-slate-200 p-5 flex flex-col justify-between space-y-3"
          >
            <div className="flex items-start justify-between gap-4">
              <div>
                <div className="text-[11px] font-mono text-slate-500">
                  {r.category} RATIO {r.isMaterialChange ? '· MATERIAL MOVEMENT' : ''}
                </div>
                <h3 className="text-base font-bold text-[#0B2540] mt-0.5">{r.name}</h3>
                <div className="text-xs text-slate-500 mt-0.5 font-mono">{r.formula}</div>
              </div>
              <div className="text-right shrink-0">
                <div className="text-xl font-bold font-mono tabular-nums text-[#0875BC]">
                  {r.formattedCurrent}
                </div>
                {r.previousValue !== null && (
                  <div className="text-[11px] font-mono text-slate-500">
                    Prior: {r.formattedPrevious}
                  </div>
                )}
              </div>
            </div>

            <div className="grid grid-cols-2 gap-3 p-3 bg-slate-50 rounded-lg border border-slate-100 text-xs">
              <div>
                <div className="text-[11px] text-slate-500">Numerator: {r.numeratorLabel}</div>
                <div className="font-mono font-semibold text-[#0B2540] mt-0.5">
                  {formatINR(r.numeratorPaise)}
                </div>
              </div>
              <div>
                <div className="text-[11px] text-slate-500">Denominator: {r.denominatorLabel}</div>
                <div className="font-mono font-semibold text-[#0B2540] mt-0.5">
                  {r.status === 'INSUFFICIENT_DATA'
                    ? 'Not Supplied'
                    : formatINR(r.denominatorPaise)}
                </div>
              </div>
            </div>

            <div className="text-xs text-slate-600 space-y-1">
              <div>{r.explanation}</div>
              <div className="text-[11px] text-slate-500">
                <strong>Account Mapping:</strong> {r.accountMappings.join(' · ')}
              </div>
            </div>
          </div>
        ))}
      </div>
    </div>
  );
};

// ============================================================================
// MODULE 1 — ACCOUNTING HEALTH CHECK & AUDIT EXCEPTION ENGINE
// ============================================================================

interface AccountingHealthCheckProps {
  health: AccountingHealthCheckResult;
  policy: PolicySettings;
  reviewStates: Record<string, AuditReviewState>;
  onUpdatePolicy: (next: PolicySettings) => void;
  onUpdateReviewState: (exceptionId: string, state: AuditReviewState) => void;
  onOpenVoucherDrillDown: (voucherId: string) => void;
}

export const AccountingHealthCheckSection: React.FC<AccountingHealthCheckProps> = ({
  health,
  policy,
  reviewStates,
  onUpdatePolicy,
  onUpdateReviewState,
  onOpenVoucherDrillDown,
}) => {
  const [severityFilter, setSeverityFilter] = useState<'ALL' | 'HIGH' | 'MEDIUM' | 'LOW'>('ALL');
  const [categoryFilter, setCategoryFilter] = useState<string>('ALL');
  const [searchQuery, setSearchQuery] = useState('');
  const [editingNoteId, setEditingNoteId] = useState<string | null>(null);
  const [noteDraft, setNoteDraft] = useState('');

  const filteredExceptions = useMemo(() => {
    return health.exceptions.filter((e) => {
      if (severityFilter !== 'ALL' && e.severity !== severityFilter) return false;
      if (categoryFilter !== 'ALL' && e.category !== categoryFilter) return false;
      if (searchQuery) {
        const q = searchQuery.toLowerCase();
        const hay = `${e.voucherNumber} ${e.ledgerName} ${e.partyName} ${e.explanation}`.toLowerCase();
        if (!hay.includes(q)) return false;
      }
      return true;
    });
  }, [health.exceptions, severityFilter, categoryFilter, searchQuery]);

  const handleExportHealthReport = () => {
    const wb = XLSX.utils.book_new();
    const summaryRows = [
      ['GSPU FINANCIAL ANALYTICS — ACCOUNTING HEALTH CHECK REPORT'],
      ['Total Vouchers Analysed', health.totalVouchersAnalysed],
      ['Total Ledgers Analysed', health.totalLedgersAnalysed],
      ['Total Exceptions Detected', health.totalExceptions],
      ['High-Priority Exceptions', health.highPriorityCount],
      ['Medium-Priority Exceptions', health.mediumPriorityCount],
      ['Low-Priority Exceptions', health.lowPriorityCount],
      ['Unbalanced Vouchers', health.unbalancedVoucherCount],
      ['Duplicate / Suspected Vouchers', health.duplicateVoucherCount],
      ['Reconciliation Differences', health.reconciliationDiffCount],
      ['Rule-Based Quality Indicator', `${health.qualityScore.overallScore}/100`],
      ['Disclaimer', health.qualityScore.disclaimer],
    ];
    XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet(summaryRows), 'Executive_Summary');

    const detailRows = [
      [
        'Date',
        'Voucher No',
        'Voucher Type',
        'Ledger',
        'Party',
        'Amount (INR)',
        'Exception Category',
        'Severity',
        'Explanation',
        'Recommended Procedure',
        'Source File',
        'Review Status',
        'Reviewer Notes',
      ],
      ...health.exceptions.map((e) => {
        const rev = reviewStates[e.id];
        return [
          e.date,
          e.voucherNumber,
          e.voucherType,
          e.ledgerName,
          e.partyName,
          fromPaise(e.amountPaise),
          e.category,
          e.severity,
          e.explanation,
          e.recommendedProcedure,
          e.sourceFilename,
          rev?.reviewed ? 'Reviewed' : 'Pending',
          rev?.reviewerNote || '',
        ];
      }),
    ];
    XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet(detailRows), 'Audit_Exceptions');
    XLSX.writeFile(wb, 'GSPU_Accounting_Health_Check_Report.xlsx');
  };

  return (
    <div className="space-y-6">
      {/* Header & Export */}
      <div className="bg-white rounded-xl border border-slate-200 p-5 flex flex-wrap items-center justify-between gap-4">
        <div>
          <h2 className="text-lg font-bold text-[#0B2540]">
            ACCOUNTING HEALTH CHECK & AUDIT INTELLIGENCE (Module 1)
          </h2>
          <p className="text-xs text-slate-500 mt-0.5">
            10 deterministic audit rules verifying voucher equilibrium, duplicates, cash-book integrity, suspense balances, and Trial Balance reconciliation.
          </p>
        </div>

        <button
          onClick={handleExportHealthReport}
          className="px-4 py-2 text-xs font-semibold text-white bg-[#0B2540] hover:bg-slate-800 rounded-lg flex items-center gap-1.5"
        >
          <Download className="w-3.5 h-3.5" />
          Download Health Check Report (.XLSX)
        </button>
      </div>

      {/* 9 Summary Cards (Module 1.1) */}
      <div className="grid grid-cols-2 sm:grid-cols-3 lg:grid-cols-9 gap-3">
        {[
          { label: 'Vouchers Analysed', val: health.totalVouchersAnalysed, color: 'text-[#0B2540]' },
          { label: 'Ledgers Analysed', val: health.totalLedgersAnalysed, color: 'text-[#0B2540]' },
          { label: 'Total Exceptions', val: health.totalExceptions, color: 'text-[#0875BC]' },
          { label: 'High Priority', val: health.highPriorityCount, color: 'text-red-600' },
          { label: 'Medium Priority', val: health.mediumPriorityCount, color: 'text-amber-600' },
          { label: 'Low Priority', val: health.lowPriorityCount, color: 'text-slate-700' },
          { label: 'Unbalanced Vouchers', val: health.unbalancedVoucherCount, color: 'text-red-600' },
          { label: 'Duplicate Vouchers', val: health.duplicateVoucherCount, color: 'text-amber-600' },
          { label: 'TB Differences', val: health.reconciliationDiffCount, color: 'text-red-600' },
        ].map((c) => (
          <div key={c.label} className="bg-white rounded-xl border border-slate-200 p-3.5">
            <div className="text-[11px] font-medium text-slate-500">{c.label}</div>
            <div className={`text-xl font-bold font-mono tabular-nums mt-1 ${c.color}`}>
              {c.val}
            </div>
          </div>
        ))}
      </div>

      {/* Explainable Data Quality Score & Threshold Controls (Module 1.4) */}
      <div className="grid grid-cols-1 lg:grid-cols-3 gap-6">
        <div className="lg:col-span-2 bg-white rounded-xl border border-slate-200 p-5 space-y-3">
          <div className="flex items-center justify-between">
            <div>
              <h3 className="text-sm font-bold text-[#0B2540]">
                Explainable Rule-Based Accounting Data Quality Indicator
              </h3>
              <p className="text-[11px] text-slate-500">{health.qualityScore.disclaimer}</p>
            </div>
            <div className="text-2xl font-bold font-mono tabular-nums text-[#0875BC]">
              {health.qualityScore.overallScore} / 100
            </div>
          </div>

          <div className="space-y-2 pt-2">
            {health.qualityScore.factors.map((f) => (
              <div
                key={f.label}
                className="flex items-center justify-between text-xs p-2.5 rounded-lg bg-slate-50 border border-slate-100"
              >
                <div>
                  <span className="font-semibold text-[#0B2540]">{f.label}</span>
                  <span className="text-slate-500 ml-2">({f.statusText})</span>
                </div>
                <div className="font-mono font-semibold">
                  {f.weight - f.deduction} / {f.weight} pts
                </div>
              </div>
            ))}
          </div>
        </div>

        <div className="bg-white rounded-xl border border-slate-200 p-5 space-y-4 text-xs">
          <h3 className="text-sm font-bold text-[#0B2540]">
            Configurable Audit Rule Thresholds
          </h3>
          <div>
            <label className="block font-semibold text-slate-700 mb-1">
              Materiality Threshold for Journal Review (₹)
            </label>
            <input
              type="number"
              value={fromPaise(policy.materialityThresholdPaise)}
              onChange={(e) =>
                onUpdatePolicy({
                  ...policy,
                  materialityThresholdPaise: Math.max(0, Number(e.target.value) * 100),
                })
              }
              className="w-full px-3 py-1.5 border border-slate-300 rounded-lg font-mono"
            />
          </div>
          <div>
            <label className="block font-semibold text-slate-700 mb-1">
              Round-Sum Transaction Review Threshold (₹)
            </label>
            <input
              type="number"
              value={fromPaise(policy.roundSumThresholdPaise)}
              onChange={(e) =>
                onUpdatePolicy({
                  ...policy,
                  roundSumThresholdPaise: Math.max(0, Number(e.target.value) * 100),
                })
              }
              className="w-full px-3 py-1.5 border border-slate-300 rounded-lg font-mono"
            />
          </div>
          <div>
            <label className="block font-semibold text-slate-700 mb-1">
              Voucher Rounding Tolerance (₹)
            </label>
            <input
              type="number"
              step="0.5"
              value={fromPaise(policy.roundingTolerancePaise)}
              onChange={(e) =>
                onUpdatePolicy({
                  ...policy,
                  roundingTolerancePaise: Math.max(0, Math.round(Number(e.target.value) * 100)),
                })
              }
              className="w-full px-3 py-1.5 border border-slate-300 rounded-lg font-mono"
            />
          </div>
        </div>
      </div>

      {/* Audit Exception Table (Module 1.3) */}
      <div className="bg-white rounded-xl border border-slate-200 p-6 space-y-4">
        <div className="flex flex-wrap items-center justify-between gap-3">
          <h3 className="text-base font-bold text-[#0B2540]">
            Identified Audit Exceptions & Review Workpapers ({filteredExceptions.length})
          </h3>
          <div className="flex flex-wrap items-center gap-2 text-xs">
            <input
              type="text"
              value={searchQuery}
              onChange={(e) => setSearchQuery(e.target.value)}
              placeholder="Search voucher, ledger, party..."
              className="px-3 py-1.5 border border-slate-300 rounded-lg"
            />
            <select
              value={severityFilter}
              onChange={(e) => setSeverityFilter(e.target.value as any)}
              className="px-2.5 py-1.5 border border-slate-300 rounded-lg bg-white"
            >
              <option value="ALL">All Severities</option>
              <option value="HIGH">High Priority</option>
              <option value="MEDIUM">Medium Priority</option>
              <option value="LOW">Low Priority</option>
            </select>
            <select
              value={categoryFilter}
              onChange={(e) => setCategoryFilter(e.target.value)}
              className="px-2.5 py-1.5 border border-slate-300 rounded-lg bg-white"
            >
              <option value="ALL">All 10 Rule Categories</option>
              <option value="EXACT_DUPLICATE_VOUCHER">A1. Exact Duplicate Voucher</option>
              <option value="POSSIBLE_DUPLICATE_VOUCHER">A2. Possible Duplicate Voucher</option>
              <option value="UNBALANCED_VOUCHER">B. Unbalanced Voucher</option>
              <option value="NEGATIVE_CASH_BALANCE">C. Negative Cash Balance</option>
              <option value="MISSING_LEDGER_CLASSIFICATION">D. Missing Ledger Classification</option>
              <option value="UNUSUAL_JOURNAL_ENTRY">E. Unusual Journal Entry</option>
              <option value="ROUND_AMOUNT_TRANSACTION">F. Round-Amount Transaction</option>
              <option value="SUSPENSE_ACCOUNT_BALANCE">G. Suspense Account</option>
              <option value="PERIOD_END_ADJUSTMENT">H. Period-End Adjustment</option>
              <option value="OPENING_CLOSING_TB_DIFFERENCE">I. Opening/Closing TB Difference</option>
              <option value="MISSING_TRANSACTION_REFERENCE">J. Missing Transaction Reference</option>
            </select>
          </div>
        </div>

        <div className="border border-slate-200 rounded-lg overflow-x-auto">
          <table className="w-full text-left border-collapse text-xs">
            <thead>
              <tr className="bg-slate-100 text-slate-700 border-b border-slate-200">
                <th className="py-2.5 px-3 font-semibold">Date</th>
                <th className="py-2.5 px-3 font-semibold">Voucher No</th>
                <th className="py-2.5 px-3 font-semibold">Type</th>
                <th className="py-2.5 px-3 font-semibold">Ledger / Party</th>
                <th className="py-2.5 px-3 font-semibold text-right">Amount (₹)</th>
                <th className="py-2.5 px-3 font-semibold">Exception Category & Severity</th>
                <th className="py-2.5 px-3 font-semibold">Audit Explanation & Source</th>
                <th className="py-2.5 px-3 font-semibold">Reviewer Status</th>
                <th className="py-2.5 px-3 font-semibold text-right">Action</th>
              </tr>
            </thead>
            <tbody className="divide-y divide-slate-200">
              {filteredExceptions.map((exc: AuditException) => {
                const rev = reviewStates[exc.id] || { reviewed: false, reviewerNote: '' };
                return (
                  <tr key={exc.id} className="hover:bg-slate-50">
                    <td className="py-3 px-3 font-mono whitespace-nowrap">{exc.date}</td>
                    <td className="py-3 px-3 font-mono font-semibold text-[#0B2540]">
                      {exc.voucherNumber}
                    </td>
                    <td className="py-3 px-3">{exc.voucherType}</td>
                    <td className="py-3 px-3 max-w-[180px]">
                      <div className="font-medium text-slate-900 truncate">{exc.ledgerName}</div>
                      <div className="text-[11px] text-slate-500 truncate">{exc.partyName}</div>
                    </td>
                    <td className="py-3 px-3 text-right font-mono tabular-nums font-bold">
                      {formatINR(exc.amountPaise)}
                    </td>
                    <td className="py-3 px-3">
                      <div className="font-mono text-[11px] font-semibold text-[#0B2540]">
                        {exc.category}
                      </div>
                      <div
                        className={`text-[11px] font-bold mt-0.5 ${
                          exc.severity === 'HIGH'
                            ? 'text-red-600'
                            : exc.severity === 'MEDIUM'
                            ? 'text-amber-600'
                            : 'text-slate-600'
                        }`}
                      >
                        Severity: {exc.severity}
                      </div>
                    </td>
                    <td className="py-3 px-3 max-w-sm">
                      <div className="text-slate-800">{exc.explanation}</div>
                      <div className="text-[11px] text-slate-500 mt-0.5">
                        Source: {exc.sourceFilename}
                      </div>
                    </td>
                    <td className="py-3 px-3">
                      <div className="flex flex-col gap-1">
                        <label className="inline-flex items-center gap-1.5 cursor-pointer">
                          <input
                            type="checkbox"
                            checked={rev.reviewed}
                            onChange={(e) =>
                              onUpdateReviewState(exc.id, {
                                ...rev,
                                reviewed: e.target.checked,
                                updatedAt: new Date().toISOString(),
                              })
                            }
                            className="rounded border-slate-300"
                          />
                          <span className="font-medium">
                            {rev.reviewed ? 'Reviewed' : 'Pending'}
                          </span>
                        </label>
                        {editingNoteId === exc.id ? (
                          <div className="flex items-center gap-1">
                            <input
                              type="text"
                              value={noteDraft}
                              onChange={(e) => setNoteDraft(e.target.value)}
                              placeholder="Auditor note..."
                              className="px-2 py-1 border border-slate-300 rounded text-[11px] w-32"
                            />
                            <button
                              onClick={() => {
                                onUpdateReviewState(exc.id, {
                                  ...rev,
                                  reviewerNote: noteDraft,
                                  updatedAt: new Date().toISOString(),
                                });
                                setEditingNoteId(null);
                              }}
                              className="px-2 py-1 bg-[#0875BC] text-white rounded text-[11px]"
                            >
                              Save
                            </button>
                          </div>
                        ) : (
                          <button
                            onClick={() => {
                              setEditingNoteId(exc.id);
                              setNoteDraft(rev.reviewerNote || '');
                            }}
                            className="text-[11px] text-[#0875BC] hover:underline text-left"
                          >
                            {rev.reviewerNote ? `Note: ${rev.reviewerNote}` : '+ Add Reviewer Note'}
                          </button>
                        )}
                      </div>
                    </td>
                    <td className="py-3 px-3 text-right whitespace-nowrap">
                      {exc.voucherId ? (
                        <button
                          onClick={() => onOpenVoucherDrillDown(exc.voucherId!)}
                          className="px-2.5 py-1 text-xs font-medium text-[#0875BC] hover:bg-blue-50 rounded inline-flex items-center gap-1"
                        >
                          <Eye className="w-3.5 h-3.5" />
                          View Voucher
                        </button>
                      ) : (
                        <span className="text-slate-400">Ledger Level</span>
                      )}
                    </td>
                  </tr>
                );
              })}
              {filteredExceptions.length === 0 && (
                <tr>
                  <td colSpan={9} className="py-8 text-center text-slate-500">
                    No audit exceptions match the current filter criteria.
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
