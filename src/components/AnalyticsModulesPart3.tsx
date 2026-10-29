import React, { useRef, useState } from 'react';
import * as XLSX from 'xlsx';
import {
  Bar,
  BarChart,
  CartesianGrid,
  Legend,
  ResponsiveContainer,
  Tooltip,
  XAxis,
  YAxis,
} from 'recharts';
import {
  Download,
  Eye,
  FileSpreadsheet,
  Printer,
  Upload,
} from 'lucide-react';
import {
  ConsolidatedAccountingModel,
  FinancialInsight,
  FinancialRatioResult,
  formatINR,
  formatINRCompact,
  fromPaise,
  PolicySettings,
} from '../types/accounting';
import { FinancialStatementBundle } from '../engine/accountingEngine';
import { AgeingAnalysisResult } from '../engine/ageingEngine';
import { AccountingHealthCheckResult } from '../engine/auditEngine';
import { exportCompleteMISExcel, exportMISReportPDF } from '../engine/exportEngine';
import { GSTReconciliationReport } from '../engine/gstMatchingEngine';
import { DrillDownTarget } from './DrillDownDrawer';
import { GspuBrandLogo } from './GspuBrandLogo';

// ============================================================================
// MODULE 7 — GST RECONCILIATION SECTION
// ============================================================================

interface GSTReconciliationSectionProps {
  gst: GSTReconciliationReport;
  policy: PolicySettings;
  onUpdatePolicy: (next: PolicySettings) => void;
  onUploadFiles: (files: FileList | File[]) => void;
  onOpenVoucherDrillDown: (voucherId: string) => void;
}

export const GSTReconciliationSection: React.FC<GSTReconciliationSectionProps> = ({
  gst,
  policy,
  onUpdatePolicy,
  onUploadFiles,
  onOpenVoucherDrillDown,
}) => {
  const gstFileInputRef = useRef<HTMLInputElement | null>(null);
  const [activeReconTab, setActiveReconTab] = useState<
    'RECON_A_SALES_GSTR1' | 'RECON_B_PURCH_GSTR2B' | 'RECON_C_GSTR1_3B' | 'RECON_D_ITC_3B'
  >('RECON_A_SALES_GSTR1');
  const [matchFilter, setMatchFilter] = useState<string>('ALL');

  const exportGstReport = (reportKind: string) => {
    const wb = XLSX.utils.book_new();
    if (
      reportKind === 'SALES_RECON' ||
      reportKind === 'PURCHASE_RECON' ||
      reportKind === 'GSTR2B_VS_BOOKS' ||
      reportKind === 'UNMATCHED' ||
      reportKind === 'GST_EXCEPTIONS'
    ) {
      const list =
        reportKind === 'SALES_RECON'
          ? gst.salesVsGstr1Matches
          : reportKind === 'PURCHASE_RECON' || reportKind === 'GSTR2B_VS_BOOKS'
          ? gst.purchasesVsGstr2bMatches
          : reportKind === 'UNMATCHED'
          ? [...gst.salesVsGstr1Matches, ...gst.purchasesVsGstr2bMatches].filter(
              (m) =>
                m.matchStatus === 'MISSING_IN_PORTAL' ||
                m.matchStatus === 'MISSING_IN_BOOKS'
            )
          : [...gst.salesVsGstr1Matches, ...gst.purchasesVsGstr2bMatches].filter(
              (m) => m.matchStatus !== 'EXACT'
            );
      const rows = [
        [
          'Reconciliation',
          'Match Status',
          'Party GSTIN',
          'Party Name',
          'Books Invoice No',
          'Books Date',
          'Portal Invoice No',
          'Portal Date',
          'Books Taxable (INR)',
          'Portal Taxable (INR)',
          'Books Total GST (INR)',
          'Portal Total GST (INR)',
          'Tax Difference (INR)',
          'Discrepancy Explanation',
        ],
        ...list.map((m) => [
          m.reconciliationType,
          m.matchStatus,
          m.booksInvoice?.partyGstin || m.portalInvoice?.partyGstin || '',
          m.booksInvoice?.partyName || m.portalInvoice?.partyName || '',
          m.booksInvoice?.invoiceNumber || 'Missing in Books',
          m.booksInvoice?.invoiceDate || '',
          m.portalInvoice?.invoiceNumber || 'Missing in Portal',
          m.portalInvoice?.invoiceDate || '',
          fromPaise(m.booksInvoice?.taxableValuePaise || 0),
          fromPaise(m.portalInvoice?.taxableValuePaise || 0),
          fromPaise(m.booksInvoice?.totalTaxPaise || 0),
          fromPaise(m.portalInvoice?.totalTaxPaise || 0),
          fromPaise(m.totalTaxDiffPaise),
          m.discrepancyReasons.join('; '),
        ]),
      ];
      XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet(rows), 'GST_Reconciliation');
    } else if (reportKind === 'GSTR1_VS_3B') {
      const rows = [
        ['Tax Head', 'Tally Books Outward (INR)', 'GSTR-1 (INR)', 'GSTR-3B (INR)', 'Diff GSTR-1 vs 3B (INR)', 'Diff Books vs 3B (INR)'],
        ...gst.gstr1VsGstr3bComparison.map((r) => [
          r.head,
          fromPaise(r.booksOutwardPaise),
          fromPaise(r.gstr1Paise),
          fromPaise(r.gstr3bPaise),
          fromPaise(r.diffGstr1Vs3bPaise),
          fromPaise(r.diffBooksVs3bPaise),
        ]),
      ];
      XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet(rows), 'GSTR1_vs_GSTR3B');
    } else {
      const rows = [
        ['ITC Head', 'Books Eligible ITC (INR)', 'GSTR-2B ITC (INR)', 'GSTR-3B Claimed (INR)', 'Excess Claimed over 2B (INR)', 'Books vs 2B Diff (INR)'],
        ...gst.itcVsGstr3bComparison.map((r) => [
          r.head,
          fromPaise(r.booksEligibleItcPaise),
          fromPaise(r.gstr2bEligibleItcPaise),
          fromPaise(r.gstr3bClaimedItcPaise),
          fromPaise(r.excessClaimedOver2bPaise),
          fromPaise(r.unclaimedBooksVs2bPaise),
        ]),
      ];
      XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet(rows), 'ITC_Summary');
    }
    XLSX.writeFile(wb, `GSPU_GST_${reportKind}.xlsx`);
  };

  const currentMatchList = (
    activeReconTab === 'RECON_A_SALES_GSTR1'
      ? gst.salesVsGstr1Matches
      : gst.purchasesVsGstr2bMatches
  ).filter((m) => (matchFilter === 'ALL' ? true : m.matchStatus === matchFilter));

  const gstChartData = gst.monthlyGstTrends.map((m) => ({
    month: m.month,
    OutputBooks: fromPaise(m.outputGstBooksPaise),
    OutputGSTR1: fromPaise(m.outputGstGstr1Paise),
    InputBooks: fromPaise(m.inputItcBooksPaise),
    InputGSTR2B: fromPaise(m.inputItcGstr2bPaise),
  }));

  return (
    <div className="space-y-6">
      {/* Module 7.1 — Dedicated GST File Upload & Controls */}
      <div className="bg-white rounded-xl border border-slate-200 p-5 flex flex-wrap items-center justify-between gap-4">
        <div>
          <h2 className="text-lg font-bold text-[#0B2540]">
            GST RECONCILIATION & RETURN MATCHING ENGINE (Module 7)
          </h2>
          <p className="text-xs text-slate-500 mt-0.5">
            Local 5-tier invoice matching (Exact, Normalised, Tolerance, Probable, Unmatched) across Tally Books, GSTR-1, GSTR-2B, and GSTR-3B.
          </p>
        </div>

        <div className="flex flex-wrap items-center gap-2.5">
          <div className="flex items-center gap-1.5 text-xs">
            <span className="text-slate-600 font-medium">GST Tolerance (₹):</span>
            <input
              type="number"
              value={fromPaise(policy.gstTolerancePaise)}
              onChange={(e) =>
                onUpdatePolicy({
                  ...policy,
                  gstTolerancePaise: Math.max(0, Math.round(Number(e.target.value) * 100)),
                })
              }
              className="w-16 px-2 py-1 border border-slate-300 rounded font-mono"
            />
          </div>

          <input
            ref={gstFileInputRef}
            type="file"
            multiple
            accept=".json,.xlsx,.xls,.csv,.xml"
            onChange={(e) => {
              if (e.target.files && e.target.files.length > 0) {
                onUploadFiles(e.target.files);
                e.target.value = '';
              }
            }}
            className="hidden"
          />
          <button
            onClick={() => gstFileInputRef.current?.click()}
            className="px-4 py-2 text-xs font-semibold text-white bg-[#0875BC] hover:bg-[#065d96] rounded-lg flex items-center gap-1.5"
          >
            <Upload className="w-3.5 h-3.5" />
            Upload GSTR-1 / 2B / 3B Files
          </button>
        </div>
      </div>

      {/* 8 Summary Cards (Module 7.7) */}
      <div className="grid grid-cols-2 sm:grid-cols-4 lg:grid-cols-8 gap-3">
        {[
          { label: 'Taxable Outward', val: formatINRCompact(gst.summaryCards.taxableOutwardSuppliesPaise) },
          { label: 'Output GST (Books)', val: formatINRCompact(gst.summaryCards.outputGstBooksPaise) },
          { label: 'Input GST (Books)', val: formatINRCompact(gst.summaryCards.inputGstBooksPaise) },
          { label: 'Book Eligible ITC', val: formatINRCompact(gst.summaryCards.bookItcPaise) },
          { label: 'GSTR-2B ITC', val: formatINRCompact(gst.summaryCards.gstr2bItcPaise) },
          { label: 'GSTR-3B ITC Claimed', val: formatINRCompact(gst.summaryCards.gstr3bItcClaimedPaise) },
          { label: 'Reconciliation Diff', val: formatINRCompact(gst.summaryCards.reconciliationDiffPaise) },
          { label: 'Unmatched Invoices', val: String(gst.summaryCards.unmatchedInvoiceCount) },
        ].map((c) => (
          <div key={c.label} className="bg-white rounded-xl border border-slate-200 p-3.5">
            <div className="text-[11px] font-medium text-slate-500">{c.label}</div>
            <div className="text-base font-bold font-mono tabular-nums text-[#0B2540] mt-1">
              {c.val}
            </div>
          </div>
        ))}
      </div>

      {/* All 7 GST Exports Bar (Module 7.9) */}
      <div className="bg-white rounded-xl border border-slate-200 p-4 flex flex-wrap items-center justify-between gap-2 text-xs">
        <span className="font-semibold text-[#0B2540]">Download GST Reconciliation Reports (7 Formats):</span>
        <div className="flex flex-wrap items-center gap-1.5">
          <button
            onClick={() => exportGstReport('SALES_RECON')}
            className="px-2.5 py-1.5 bg-slate-100 hover:bg-slate-200 rounded font-medium flex items-center gap-1"
          >
            <Download className="w-3 h-3" /> 1. Sales Reconciliation
          </button>
          <button
            onClick={() => exportGstReport('PURCHASE_RECON')}
            className="px-2.5 py-1.5 bg-slate-100 hover:bg-slate-200 rounded font-medium flex items-center gap-1"
          >
            <Download className="w-3 h-3" /> 2. Purchase Reconciliation
          </button>
          <button
            onClick={() => exportGstReport('GSTR1_VS_3B')}
            className="px-2.5 py-1.5 bg-slate-100 hover:bg-slate-200 rounded font-medium flex items-center gap-1"
          >
            <Download className="w-3 h-3" /> 3. GSTR-1 vs GSTR-3B
          </button>
          <button
            onClick={() => exportGstReport('GSTR2B_VS_BOOKS')}
            className="px-2.5 py-1.5 bg-slate-100 hover:bg-slate-200 rounded font-medium flex items-center gap-1"
          >
            <Download className="w-3 h-3" /> 4. GSTR-2B vs Books
          </button>
          <button
            onClick={() => exportGstReport('ITC_SUMMARY')}
            className="px-2.5 py-1.5 bg-slate-100 hover:bg-slate-200 rounded font-medium flex items-center gap-1"
          >
            <Download className="w-3 h-3" /> 5. ITC Summary
          </button>
          <button
            onClick={() => exportGstReport('UNMATCHED')}
            className="px-2.5 py-1.5 bg-amber-50 hover:bg-amber-100 text-amber-800 rounded font-medium flex items-center gap-1"
          >
            <Download className="w-3 h-3" /> 6. Unmatched Invoices
          </button>
          <button
            onClick={() => exportGstReport('GST_EXCEPTIONS')}
            className="px-2.5 py-1.5 bg-red-50 hover:bg-red-100 text-red-700 rounded font-medium flex items-center gap-1"
          >
            <Download className="w-3 h-3" /> 7. GST Exception Report
          </button>
        </div>
      </div>

      {/* 4 Reconciliation Tabs (Module 7.3 to 7.6) */}
      <div className="bg-white rounded-xl border border-slate-200 p-6 space-y-5">
        <div className="flex flex-wrap items-center justify-between gap-4">
          <div className="flex flex-wrap items-center gap-1 p-1 bg-slate-100 rounded-lg text-xs">
            <button
              onClick={() => setActiveReconTab('RECON_A_SALES_GSTR1')}
              className={`px-3.5 py-1.5 font-semibold rounded-md ${
                activeReconTab === 'RECON_A_SALES_GSTR1'
                  ? 'bg-white text-[#0B2540] shadow-xs'
                  : 'text-slate-600'
              }`}
            >
              A. Tally Sales vs GSTR-1 ({gst.salesVsGstr1Matches.length})
            </button>
            <button
              onClick={() => setActiveReconTab('RECON_B_PURCH_GSTR2B')}
              className={`px-3.5 py-1.5 font-semibold rounded-md ${
                activeReconTab === 'RECON_B_PURCH_GSTR2B'
                  ? 'bg-white text-[#0B2540] shadow-xs'
                  : 'text-slate-600'
              }`}
            >
              B. Tally Purchases vs GSTR-2B ({gst.purchasesVsGstr2bMatches.length})
            </button>
            <button
              onClick={() => setActiveReconTab('RECON_C_GSTR1_3B')}
              className={`px-3.5 py-1.5 font-semibold rounded-md ${
                activeReconTab === 'RECON_C_GSTR1_3B'
                  ? 'bg-white text-[#0B2540] shadow-xs'
                  : 'text-slate-600'
              }`}
            >
              C. GSTR-1 vs GSTR-3B Liability
            </button>
            <button
              onClick={() => setActiveReconTab('RECON_D_ITC_3B')}
              className={`px-3.5 py-1.5 font-semibold rounded-md ${
                activeReconTab === 'RECON_D_ITC_3B'
                  ? 'bg-white text-[#0B2540] shadow-xs'
                  : 'text-slate-600'
              }`}
            >
              D. Books ITC vs 2B vs 3B
            </button>
          </div>

          {(activeReconTab === 'RECON_A_SALES_GSTR1' ||
            activeReconTab === 'RECON_B_PURCH_GSTR2B') && (
            <select
              value={matchFilter}
              onChange={(e) => setMatchFilter(e.target.value)}
              className="px-3 py-1.5 text-xs border border-slate-300 rounded-lg bg-white"
            >
              <option value="ALL">All 5 Matching Tiers</option>
              <option value="EXACT">1. Exact Matches</option>
              <option value="NORMALISED">2. Normalised Invoice Matches</option>
              <option value="TOLERANCE">3. Tolerance Matches</option>
              <option value="PROBABLE">4. Probable Matches (Needs Review)</option>
              <option value="MISSING_IN_PORTAL">5A. Missing in GST Portal</option>
              <option value="MISSING_IN_BOOKS">5B. Missing in Tally Books</option>
            </select>
          )}
        </div>

        {(activeReconTab === 'RECON_A_SALES_GSTR1' ||
          activeReconTab === 'RECON_B_PURCH_GSTR2B') && (
          <div className="border border-slate-200 rounded-lg overflow-x-auto">
            <table className="w-full text-left border-collapse text-xs">
              <thead>
                <tr className="bg-slate-100 text-slate-700 border-b border-slate-200">
                  <th className="py-2.5 px-3 font-semibold">Match Tier</th>
                  <th className="py-2.5 px-3 font-semibold">Party Name & GSTIN</th>
                  <th className="py-2.5 px-3 font-semibold">Books Invoice</th>
                  <th className="py-2.5 px-3 font-semibold">Portal Invoice</th>
                  <th className="py-2.5 px-3 font-semibold text-right">Books Tax (₹)</th>
                  <th className="py-2.5 px-3 font-semibold text-right">Portal Tax (₹)</th>
                  <th className="py-2.5 px-3 font-semibold text-right">Tax Diff (₹)</th>
                  <th className="py-2.5 px-3 font-semibold">Reconciliation Observations</th>
                  <th className="py-2.5 px-3 font-semibold text-right">Voucher</th>
                </tr>
              </thead>
              <tbody className="divide-y divide-slate-200">
                {currentMatchList.map((m) => (
                  <tr key={m.id} className="hover:bg-slate-50">
                    <td className="py-2.5 px-3 font-mono font-bold">
                      <span
                        className={
                          m.matchStatus === 'EXACT' || m.matchStatus === 'NORMALISED'
                            ? 'text-emerald-700'
                            : m.matchStatus === 'PROBABLE' || m.matchStatus === 'TOLERANCE'
                            ? 'text-amber-700'
                            : 'text-red-600'
                        }
                      >
                        {m.matchStatus}
                      </span>
                    </td>
                    <td className="py-2.5 px-3">
                      <div className="font-medium text-[#0B2540]">
                        {m.booksInvoice?.partyName || m.portalInvoice?.partyName}
                      </div>
                      <div className="font-mono text-[11px] text-slate-500">
                        {m.booksInvoice?.partyGstin || m.portalInvoice?.partyGstin || 'Unregistered'}
                      </div>
                    </td>
                    <td className="py-2.5 px-3 font-mono">
                      {m.booksInvoice ? (
                        <div>
                          <div>{m.booksInvoice.invoiceNumber}</div>
                          <div className="text-[11px] text-slate-500">
                            {m.booksInvoice.invoiceDate}
                          </div>
                        </div>
                      ) : (
                        <span className="text-red-600">Not in Books</span>
                      )}
                    </td>
                    <td className="py-2.5 px-3 font-mono">
                      {m.portalInvoice ? (
                        <div>
                          <div>{m.portalInvoice.invoiceNumber}</div>
                          <div className="text-[11px] text-slate-500">
                            {m.portalInvoice.invoiceDate}
                          </div>
                        </div>
                      ) : (
                        <span className="text-red-600">Not in Portal</span>
                      )}
                    </td>
                    <td className="py-2.5 px-3 text-right font-mono tabular-nums">
                      {formatINR(m.booksInvoice?.totalTaxPaise || 0)}
                    </td>
                    <td className="py-2.5 px-3 text-right font-mono tabular-nums">
                      {formatINR(m.portalInvoice?.totalTaxPaise || 0)}
                    </td>
                    <td className="py-2.5 px-3 text-right font-mono tabular-nums font-bold">
                      {formatINR(m.totalTaxDiffPaise)}
                    </td>
                    <td className="py-2.5 px-3 text-slate-600 max-w-xs">
                      {m.discrepancyReasons.length > 0
                        ? m.discrepancyReasons.join(' · ')
                        : 'Reconciled with zero tax variance'}
                    </td>
                    <td className="py-2.5 px-3 text-right">
                      {m.booksInvoice?.linkedVoucherId ? (
                        <button
                          onClick={() =>
                            onOpenVoucherDrillDown(m.booksInvoice!.linkedVoucherId!)
                          }
                          className="text-[#0875BC] hover:underline inline-flex items-center gap-1"
                        >
                          <Eye className="w-3.5 h-3.5" />
                          View
                        </button>
                      ) : (
                        '—'
                      )}
                    </td>
                  </tr>
                ))}
              </tbody>
            </table>
          </div>
        )}

        {activeReconTab === 'RECON_C_GSTR1_3B' && (
          <div className="border border-slate-200 rounded-lg overflow-x-auto">
            <table className="w-full text-left border-collapse text-xs">
              <thead>
                <tr className="bg-slate-100 text-slate-700 border-b border-slate-200">
                  <th className="py-2.5 px-4 font-semibold">Outward Tax Head</th>
                  <th className="py-2.5 px-4 font-semibold text-right">Tally Sales Books (₹)</th>
                  <th className="py-2.5 px-4 font-semibold text-right">GSTR-1 Return (₹)</th>
                  <th className="py-2.5 px-4 font-semibold text-right">GSTR-3B Table 3.1 (₹)</th>
                  <th className="py-2.5 px-4 font-semibold text-right">Diff: GSTR-1 vs 3B</th>
                  <th className="py-2.5 px-4 font-semibold text-right">Diff: Books vs 3B</th>
                </tr>
              </thead>
              <tbody className="divide-y divide-slate-200">
                {gst.gstr1VsGstr3bComparison.map((row) => (
                  <tr key={row.head} className="hover:bg-slate-50">
                    <td className="py-3 px-4 font-semibold text-[#0B2540]">{row.head}</td>
                    <td className="py-3 px-4 text-right font-mono tabular-nums">
                      {formatINR(row.booksOutwardPaise)}
                    </td>
                    <td className="py-3 px-4 text-right font-mono tabular-nums">
                      {formatINR(row.gstr1Paise)}
                    </td>
                    <td className="py-3 px-4 text-right font-mono tabular-nums">
                      {formatINR(row.gstr3bPaise)}
                    </td>
                    <td className="py-3 px-4 text-right font-mono tabular-nums font-bold">
                      {formatINR(row.diffGstr1Vs3bPaise)}
                    </td>
                    <td className="py-3 px-4 text-right font-mono tabular-nums font-bold">
                      {formatINR(row.diffBooksVs3bPaise)}
                    </td>
                  </tr>
                ))}
              </tbody>
            </table>
          </div>
        )}

        {activeReconTab === 'RECON_D_ITC_3B' && (
          <div className="space-y-3">
            <div className="p-3 bg-slate-50 rounded-lg border border-slate-200 text-xs text-slate-600">
              <strong>Legal Eligibility Notice:</strong> Appearance of an invoice in GSTR-2B is a statutory precondition under Section 16(2)(aa) but is not by itself conclusive proof of ITC eligibility; verify receipt of goods/services and Section 17(5) blocked credit provisions.
            </div>
            <div className="border border-slate-200 rounded-lg overflow-x-auto">
              <table className="w-full text-left border-collapse text-xs">
                <thead>
                  <tr className="bg-slate-100 text-slate-700 border-b border-slate-200">
                    <th className="py-2.5 px-4 font-semibold">Input Tax Credit Head</th>
                    <th className="py-2.5 px-4 font-semibold text-right">Books Purchase ITC (₹)</th>
                    <th className="py-2.5 px-4 font-semibold text-right">GSTR-2B Statement (₹)</th>
                    <th className="py-2.5 px-4 font-semibold text-right">GSTR-3B Table 4 Claimed (₹)</th>
                    <th className="py-2.5 px-4 font-semibold text-right">Claimed vs 2B Variance</th>
                    <th className="py-2.5 px-4 font-semibold text-right">Books vs 2B Variance</th>
                  </tr>
                </thead>
                <tbody className="divide-y divide-slate-200">
                  {gst.itcVsGstr3bComparison.map((row) => (
                    <tr key={row.head} className="hover:bg-slate-50">
                      <td className="py-3 px-4 font-semibold text-[#0B2540]">{row.head}</td>
                      <td className="py-3 px-4 text-right font-mono tabular-nums">
                        {formatINR(row.booksEligibleItcPaise)}
                      </td>
                      <td className="py-3 px-4 text-right font-mono tabular-nums">
                        {formatINR(row.gstr2bEligibleItcPaise)}
                      </td>
                      <td className="py-3 px-4 text-right font-mono tabular-nums">
                        {formatINR(row.gstr3bClaimedItcPaise)}
                      </td>
                      <td className="py-3 px-4 text-right font-mono tabular-nums font-bold">
                        {formatINR(row.excessClaimedOver2bPaise)}
                      </td>
                      <td className="py-3 px-4 text-right font-mono tabular-nums font-bold">
                        {formatINR(row.unclaimedBooksVs2bPaise)}
                      </td>
                    </tr>
                  ))}
                </tbody>
              </table>
            </div>
          </div>
        )}
      </div>

      {/* GST Charts & Supplier Mismatch Table */}
      {gstChartData.length > 0 && (
        <div className="grid grid-cols-1 lg:grid-cols-2 gap-6">
          <div className="bg-white rounded-xl border border-slate-200 p-6">
            <h3 className="text-sm font-bold text-[#0B2540] mb-4">
              Monthly Output GST & Input ITC: Books vs Portal Returns
            </h3>
            <div className="h-64">
              <ResponsiveContainer width="100%" height="100%">
                <BarChart data={gstChartData}>
                  <CartesianGrid strokeDasharray="3 3" stroke="#E2E8F0" />
                  <XAxis dataKey="month" tick={{ fontSize: 11 }} />
                  <YAxis
                    tick={{ fontSize: 11 }}
                    tickFormatter={(v) => `₹${(v / 100000).toFixed(1)}L`}
                  />
                  <Tooltip
                    formatter={(val: any) => formatINR(Math.round(Number(val) * 100))}
                  />
                  <Legend wrapperStyle={{ fontSize: 11 }} />
                  <Bar dataKey="OutputBooks" name="Books Output GST" fill="#0B2540" />
                  <Bar dataKey="OutputGSTR1" name="GSTR-1 Output GST" fill="#0875BC" />
                  <Bar dataKey="InputBooks" name="Books Input ITC" fill="#0F9B91" />
                  <Bar dataKey="InputGSTR2B" name="GSTR-2B Input ITC" fill="#50AE5D" />
                </BarChart>
              </ResponsiveContainer>
            </div>
          </div>

          <div className="bg-white rounded-xl border border-slate-200 p-6">
            <h3 className="text-sm font-bold text-[#0B2540] mb-4">
              Supplier-Wise ITC Discrepancy Analysis
            </h3>
            <div className="border border-slate-200 rounded-lg overflow-x-auto">
              <table className="w-full text-left border-collapse text-xs">
                <thead>
                  <tr className="bg-slate-100 border-b border-slate-200">
                    <th className="py-2 px-3">Supplier Name & GSTIN</th>
                    <th className="py-2 px-3 text-right">Books Tax</th>
                    <th className="py-2 px-3 text-right">GSTR-2B Tax</th>
                    <th className="py-2 px-3 text-right">Difference</th>
                  </tr>
                </thead>
                <tbody className="divide-y divide-slate-200">
                  {gst.supplierMismatches.map((s) => (
                    <tr key={s.partyGstin + s.partyName}>
                      <td className="py-2 px-3">
                        <div className="font-semibold text-[#0B2540]">{s.partyName}</div>
                        <div className="font-mono text-[11px] text-slate-500">{s.partyGstin}</div>
                      </td>
                      <td className="py-2 px-3 text-right font-mono">{formatINR(s.booksTaxPaise)}</td>
                      <td className="py-2 px-3 text-right font-mono">{formatINR(s.portalTaxPaise)}</td>
                      <td className="py-2 px-3 text-right font-mono font-bold text-red-600">
                        {formatINR(s.diffPaise)}
                      </td>
                    </tr>
                  ))}
                  {gst.supplierMismatches.length === 0 && (
                    <tr>
                      <td colSpan={4} className="py-6 text-center text-slate-500">
                        Zero supplier ITC discrepancies detected.
                      </td>
                    </tr>
                  )}
                </tbody>
              </table>
            </div>
          </div>
        </div>
      )}
    </div>
  );
};

// ============================================================================
// MODULE 8 — AUTOMATED FINANCIAL INSIGHTS SECTION
// ============================================================================

interface FinancialInsightsSectionProps {
  insights: FinancialInsight[];
  onDrillDown: (target: DrillDownTarget) => void;
  onNavigateTab: (tabId: string) => void;
}

export const FinancialInsightsSection: React.FC<FinancialInsightsSectionProps> = ({
  insights,
  onDrillDown,
  onNavigateTab,
}) => {
  const [priorityFilter, setPriorityFilter] = useState<string>('ALL');

  const filtered = insights.filter(
    (i) => priorityFilter === 'ALL' || i.priority === priorityFilter
  );

  return (
    <div className="space-y-6">
      <div className="bg-white rounded-xl border border-slate-200 p-5 flex flex-wrap items-center justify-between gap-4">
        <div>
          <h2 className="text-lg font-bold text-[#0B2540]">
            FINANCIAL INSIGHTS — DETERMINISTIC NARRATIVE ENGINE (Module 8)
          </h2>
          <p className="text-xs text-slate-500 mt-0.5">
            100% rules-based financial observations backed by calculated values, formulas, and source files. Zero external AI/LLM calls.
          </p>
        </div>

        <select
          value={priorityFilter}
          onChange={(e) => setPriorityFilter(e.target.value)}
          className="px-3 py-1.5 text-xs border border-slate-300 rounded-lg bg-white"
        >
          <option value="ALL">All Priority Levels ({insights.length})</option>
          <option value="Significant risk indicator">Significant risk indicator</option>
          <option value="Attention required">Attention required</option>
          <option value="Positive development">Positive development</option>
          <option value="Informational">Informational</option>
        </select>
      </div>

      <div className="grid grid-cols-1 gap-4">
        {filtered.map((ins) => (
          <div
            key={ins.id}
            className="bg-white rounded-xl border border-slate-200 p-6 space-y-4"
          >
            <div className="flex flex-wrap items-start justify-between gap-4">
              <div>
                <div className="text-xs text-slate-500 font-semibold">
                  {ins.category} ·{' '}
                  <span
                    className={
                      ins.priority === 'Significant risk indicator'
                        ? 'text-red-600 font-bold'
                        : ins.priority === 'Attention required'
                        ? 'text-amber-600 font-bold'
                        : ins.priority === 'Positive development'
                        ? 'text-emerald-700 font-bold'
                        : 'text-[#0875BC] font-bold'
                    }
                  >
                    {ins.priority}
                  </span>
                </div>
                <h3 className="text-base font-bold text-[#0B2540] mt-1">{ins.headline}</h3>
                <p className="text-xs text-slate-600 mt-1">{ins.narrative}</p>
              </div>

              <button
                onClick={() => {
                  if (ins.drillDownCategory === 'EXCEPTIONS') {
                    onNavigateTab('HEALTH_CHECK');
                  } else if (ins.drillDownCategory === 'GST') {
                    onNavigateTab('GST_RECON');
                  } else if (
                    ins.drillDownCategory === 'RECEIVABLES' ||
                    ins.drillDownCategory === 'PAYABLES'
                  ) {
                    onNavigateTab('AGEING');
                  } else if (ins.drillDownCategory) {
                    onDrillDown({
                      title: ins.headline,
                      formulaExplanation: ins.formulaOrRule,
                      classifications: [ins.drillDownCategory],
                    });
                  }
                }}
                className="px-3.5 py-2 text-xs font-semibold text-[#0875BC] bg-blue-50 hover:bg-blue-100 rounded-lg whitespace-nowrap"
              >
                View Supporting Records →
              </button>
            </div>

            <div className="grid grid-cols-2 sm:grid-cols-4 gap-3 p-3.5 bg-slate-50 rounded-lg border border-slate-100 text-xs">
              <div>
                <div className="text-[11px] text-slate-500">Current Value</div>
                <div className="font-mono font-bold text-[#0B2540] mt-0.5">
                  {ins.currentValueFormatted}
                </div>
              </div>
              <div>
                <div className="text-[11px] text-slate-500">Comparative Value</div>
                <div className="font-mono font-semibold text-slate-700 mt-0.5">
                  {ins.comparativeValueFormatted}
                </div>
              </div>
              <div>
                <div className="text-[11px] text-slate-500">Absolute Movement</div>
                <div className="font-mono font-semibold text-slate-800 mt-0.5">
                  {ins.absoluteMovementFormatted}
                </div>
              </div>
              <div>
                <div className="text-[11px] text-slate-500">Percentage / Indicator</div>
                <div className="font-mono font-bold text-[#0875BC] mt-0.5">
                  {ins.percentageMovementFormatted}
                </div>
              </div>
            </div>

            <div className="flex flex-wrap items-center justify-between gap-2 text-[11px] text-slate-500 pt-1">
              <div>
                <strong>Formula / Rule:</strong> <span className="font-mono">{ins.formulaOrRule}</span>
              </div>
              <div>
                <strong>Source Files:</strong> {ins.sourceFiles.join(', ')}
              </div>
            </div>
          </div>
        ))}
      </div>
    </div>
  );
};

// ============================================================================
// MODULE 5 — AUTOMATIC MONTHLY MIS REPORT GENERATOR (10-PAGE STRUCTURE)
// ============================================================================

interface MISReportSectionProps {
  model: ConsolidatedAccountingModel;
  statements: FinancialStatementBundle;
  ageing: AgeingAnalysisResult;
  ratios: FinancialRatioResult[];
  health: AccountingHealthCheckResult;
  gst: GSTReconciliationReport;
  insights: FinancialInsight[];
  cfo?: any;
  onOpenExcelModal?: () => void;
}

export const MISReportSection: React.FC<MISReportSectionProps> = ({
  model,
  statements,
  ageing,
  ratios,
  health,
  gst,
  insights,
  cfo,
  onOpenExcelModal,
}) => {
  const [selectedMonth, setSelectedMonth] = useState<string>('FULL_PERIOD');
  const [selectedFy, setSelectedFy] = useState<string>('FY 2025-26');
  const [comparativeMode, setComparativeMode] = useState<string>('PRIOR_MONTH');
  const [currencyFormat, setCurrencyFormat] = useState<'EXACT_INR' | 'COMPACT_LAKH_CR'>('EXACT_INR');
  const [enabledSections, setEnabledSections] = useState<Record<number, boolean>>({
    1: true,
    2: true,
    3: true,
    4: true,
    5: true,
    6: true,
    7: true,
    8: true,
    9: true,
    10: true,
  });
  const [isReportGenerated, setIsReportGenerated] = useState(true);

  const fmt = (paise: number) =>
    currencyFormat === 'COMPACT_LAKH_CR' ? formatINRCompact(paise) : formatINR(paise);

  const k = statements.kpis;
  const generatedDate = new Date().toISOString().slice(0, 10);

  return (
    <div className="space-y-6">
      {/* 5.1 Report Generator Configuration */}
      <div className="bg-white rounded-xl border border-slate-200 p-6 space-y-4 no-print">
        <div className="flex flex-wrap items-center justify-between gap-4">
          <div>
            <h2 className="text-lg font-bold text-[#0B2540]">
              MANAGEMENT INFORMATION REPORT GENERATOR (Module 5)
            </h2>
            <p className="text-xs text-slate-500 mt-0.5">
              Select reporting month, financial year, comparative period, report sections, and currency format, then click GENERATE MIS REPORT.
            </p>
          </div>

          <div className="flex flex-wrap items-center gap-2">
            <button
              onClick={() => setIsReportGenerated(true)}
              className="px-4 py-2 text-xs font-bold text-white bg-[#50AE5D] hover:bg-[#43994f] rounded-lg cursor-pointer"
            >
              GENERATE MIS REPORT
            </button>

            <button
              onClick={() =>
                exportMISReportPDF(model, statements, ageing, ratios, health, gst, insights)
              }
              className="px-3.5 py-2 text-xs font-semibold text-white bg-[#0B2540] hover:bg-slate-800 rounded-lg flex items-center gap-1.5"
            >
              <Download className="w-3.5 h-3.5" />
              Export PDF
            </button>

            <button
              onClick={() => {
                if (onOpenExcelModal) {
                  onOpenExcelModal();
                } else {
                  exportCompleteMISExcel(model, statements, ageing, ratios, health, gst, insights, cfo);
                }
              }}
              className="px-3.5 py-2 text-xs font-semibold text-white bg-[#0875BC] hover:bg-[#065d96] rounded-lg flex items-center gap-1.5 cursor-pointer"
            >
              <FileSpreadsheet className="w-3.5 h-3.5" />
              DOWNLOAD CFO EXCEL REPORT (.XLSX)
            </button>

            <button
              onClick={() => window.print()}
              className="px-3.5 py-2 text-xs font-semibold text-slate-700 bg-slate-100 hover:bg-slate-200 rounded-lg flex items-center gap-1.5"
            >
              <Printer className="w-3.5 h-3.5" />
              Print HTML
            </button>
          </div>
        </div>

        {/* ConfigurationSelectors */}
        <div className="grid grid-cols-1 sm:grid-cols-2 lg:grid-cols-4 gap-3 text-xs pt-2 border-t border-slate-100">
          <div>
            <label className="block font-semibold text-slate-600 mb-1">Reporting Month</label>
            <select
              value={selectedMonth}
              onChange={(e) => setSelectedMonth(e.target.value)}
              className="w-full px-3 py-1.5 border border-slate-300 rounded-lg bg-white"
            >
              <option value="FULL_PERIOD">
                Full Period ({model.periodStart} to {model.periodEnd})
              </option>
              {model.availableMonths.map((m) => (
                <option key={m} value={m}>
                  Month: {m}
                </option>
              ))}
            </select>
          </div>

          <div>
            <label className="block font-semibold text-slate-600 mb-1">Financial Year</label>
            <select
              value={selectedFy}
              onChange={(e) => setSelectedFy(e.target.value)}
              className="w-full px-3 py-1.5 border border-slate-300 rounded-lg bg-white"
            >
              <option value="FY 2025-26">FY 2025-26 (Apr – Mar)</option>
              <option value="FY 2024-25">FY 2024-25 (Apr – Mar)</option>
            </select>
          </div>

          <div>
            <label className="block font-semibold text-slate-600 mb-1">Comparative Period</label>
            <select
              value={comparativeMode}
              onChange={(e) => setComparativeMode(e.target.value)}
              className="w-full px-3 py-1.5 border border-slate-300 rounded-lg bg-white"
            >
              <option value="PRIOR_MONTH">Previous Month</option>
              <option value="PRIOR_QUARTER">Previous Quarter</option>
              <option value="PRIOR_FY">Previous Financial Year</option>
            </select>
          </div>

          <div>
            <label className="block font-semibold text-slate-600 mb-1">Currency Display Format</label>
            <select
              value={currencyFormat}
              onChange={(e) => setCurrencyFormat(e.target.value as any)}
              className="w-full px-3 py-1.5 border border-slate-300 rounded-lg bg-white"
            >
              <option value="EXACT_INR">Exact Indian Rupee (₹12,50,000)</option>
              <option value="COMPACT_LAKH_CR">Lakh / Crore Display (₹12.50 Lakh)</option>
            </select>
          </div>
        </div>

        {/* Section Checkboxes */}
        <div className="pt-2">
          <div className="text-xs font-semibold text-slate-600 mb-1.5">
            Included MIS Report Pages (Toggle Sections 1–10):
          </div>
          <div className="flex flex-wrap gap-2 text-xs">
            {[
              { id: 1, label: '1. Cover Page' },
              { id: 2, label: '2. Executive Summary' },
              { id: 3, label: '3. Profitability' },
              { id: 4, label: '4. Balance Sheet' },
              { id: 5, label: '5. Sales & Purchases' },
              { id: 6, label: '6. Receivables & Payables' },
              { id: 7, label: '7. Financial Ratios' },
              { id: 8, label: '8. Accounting Exceptions' },
              { id: 9, label: '9. Management Insights' },
              { id: 10, label: '10. Data Sources' },
            ].map((s) => (
              <label
                key={s.id}
                className="inline-flex items-center gap-1.5 px-2.5 py-1 rounded bg-slate-50 border border-slate-200 cursor-pointer"
              >
                <input
                  type="checkbox"
                  checked={Boolean(enabledSections[s.id])}
                  onChange={(e) =>
                    setEnabledSections({ ...enabledSections, [s.id]: e.target.checked })
                  }
                />
                <span>{s.label}</span>
              </label>
            ))}
          </div>
        </div>
      </div>

      {/* 5.2 10-PAGE MANAGEMENT INFORMATION REPORT PREVIEW */}
      {isReportGenerated && (
        <div className="bg-white rounded-xl border border-slate-200 p-8 space-y-10 shadow-xs">
          {/* PAGE 1 — COVER PAGE */}
          {enabledSections[1] && (
            <div className="border-b-2 border-[#0B2540] pb-8 flex flex-wrap items-center justify-between gap-6">
              <div>
                <GspuBrandLogo />
                <h1 className="text-2xl font-bold text-[#0B2540] mt-4">
                  Management Information Report (MIS)
                </h1>
                <div className="text-sm font-semibold text-[#0875BC] mt-1">
                  Client / Enterprise: {model.companyName}
                </div>
              </div>
              <div className="text-right font-mono text-xs text-slate-600 space-y-1">
                <div>
                  <strong>Financial Year:</strong> {selectedFy}
                </div>
                <div>
                  <strong>Reporting Period:</strong>{' '}
                  {selectedMonth === 'FULL_PERIOD'
                    ? `${model.periodStart} to ${model.periodEnd}`
                    : selectedMonth}
                </div>
                <div>
                  <strong>Date Generated:</strong> {generatedDate}
                </div>
                <div>
                  <strong>Source Files Consolidated:</strong> {model.sourceFiles.length}
                </div>
              </div>
            </div>
          )}

          {/* PAGE 2 — EXECUTIVE FINANCIAL SUMMARY */}
          {enabledSections[2] && (
            <section className="space-y-3">
              <h2 className="text-base font-bold text-[#0B2540] border-b border-slate-200 pb-2">
                Page 2 · Executive Financial Summary
              </h2>
              <div className="grid grid-cols-2 sm:grid-cols-4 gap-4 text-xs">
                {[
                  { label: 'Net Revenue', val: fmt(k.netRevenuePaise) },
                  { label: 'Gross Profit', val: fmt(k.grossProfitPaise) },
                  { label: 'EBITDA', val: fmt(k.ebitdaPaise) },
                  { label: 'Net Profit (PAT)', val: fmt(k.patPaise) },
                  { label: 'Cash & Bank', val: fmt(k.cashPaise + k.bankPaise) },
                  { label: 'Trade Receivables', val: fmt(k.tradeReceivablesPaise) },
                  { label: 'Trade Payables', val: fmt(k.tradePayablesPaise) },
                  { label: 'Net Working Capital', val: fmt(k.workingCapitalPaise) },
                ].map((item) => (
                  <div key={item.label} className="p-3.5 rounded-lg bg-slate-50 border border-slate-200">
                    <div className="text-slate-500">{item.label}</div>
                    <div className="text-base font-bold font-mono tabular-nums text-[#0B2540] mt-1">
                      {item.val}
                    </div>
                  </div>
                ))}
              </div>
            </section>
          )}

          {/* PAGE 3 & 4 — PROFITABILITY & BALANCE SHEET ANALYSIS */}
          <div className="grid grid-cols-1 lg:grid-cols-2 gap-8">
            {enabledSections[3] && (
              <section className="space-y-3">
                <h2 className="text-base font-bold text-[#0B2540] border-b border-slate-200 pb-2">
                  Page 3 · Profitability Analysis
                </h2>
                <table className="w-full text-left border-collapse text-xs">
                  <tbody className="divide-y divide-slate-200">
                    {statements.pnlLines.map((l) => (
                      <tr key={l.code} className={l.isSubtotal ? 'font-bold bg-slate-50' : ''}>
                        <td className="py-2 px-2">{l.label}</td>
                        <td className="py-2 px-2 text-right font-mono tabular-nums">
                          {fmt(l.amountPaise)}
                        </td>
                      </tr>
                    ))}
                  </tbody>
                </table>
              </section>
            )}

            {enabledSections[4] && (
              <section className="space-y-3">
                <h2 className="text-base font-bold text-[#0B2540] border-b border-slate-200 pb-2">
                  Page 4 · Balance Sheet & Liquidity Position
                </h2>
                <table className="w-full text-left border-collapse text-xs">
                  <tbody className="divide-y divide-slate-200">
                    {statements.bsAssetLines.map((l) => (
                      <tr key={l.code} className={l.isSubtotal ? 'font-bold bg-slate-50' : ''}>
                        <td className="py-2 px-2">{l.label}</td>
                        <td className="py-2 px-2 text-right font-mono tabular-nums">
                          {fmt(l.amountPaise)}
                        </td>
                      </tr>
                    ))}
                    {statements.bsLiabilityEquityLines.map((l) => (
                      <tr key={l.code} className={l.isSubtotal ? 'font-bold bg-slate-50' : ''}>
                        <td className="py-2 px-2">{l.label}</td>
                        <td className="py-2 px-2 text-right font-mono tabular-nums">
                          {fmt(l.amountPaise)}
                        </td>
                      </tr>
                    ))}
                  </tbody>
                </table>
              </section>
            )}
          </div>

          {/* PAGE 5 & 6 — SALES/PURCHASES & RECEIVABLES/PAYABLES */}
          <div className="grid grid-cols-1 lg:grid-cols-2 gap-8">
            {enabledSections[5] && (
              <section className="space-y-3">
                <h2 className="text-base font-bold text-[#0B2540] border-b border-slate-200 pb-2">
                  Page 5 · Sales & Purchase Movement
                </h2>
                <table className="w-full text-left border-collapse text-xs">
                  <thead>
                    <tr className="bg-slate-100">
                      <th className="py-2 px-2">Month</th>
                      <th className="py-2 px-2 text-right">Net Revenue</th>
                      <th className="py-2 px-2 text-right">Net Purchases</th>
                      <th className="py-2 px-2 text-right">Net Profit</th>
                    </tr>
                  </thead>
                  <tbody className="divide-y divide-slate-200">
                    {statements.monthlyTrends.map((m) => (
                      <tr key={m.month}>
                        <td className="py-2 px-2 font-mono">{m.month}</td>
                        <td className="py-2 px-2 text-right font-mono">{fmt(m.revenuePaise)}</td>
                        <td className="py-2 px-2 text-right font-mono">{fmt(m.purchasesPaise)}</td>
                        <td className="py-2 px-2 text-right font-mono">{fmt(m.netProfitPaise)}</td>
                      </tr>
                    ))}
                  </tbody>
                </table>
              </section>
            )}

            {enabledSections[6] && (
              <section className="space-y-3">
                <h2 className="text-base font-bold text-[#0B2540] border-b border-slate-200 pb-2">
                  Page 6 · Receivables & Payables Ageing Summary
                </h2>
                <table className="w-full text-left border-collapse text-xs">
                  <thead>
                    <tr className="bg-slate-100">
                      <th className="py-2 px-2">Bucket</th>
                      <th className="py-2 px-2 text-right">Receivables</th>
                      <th className="py-2 px-2 text-right">Payables</th>
                    </tr>
                  </thead>
                  <tbody className="divide-y divide-slate-200">
                    {ageing.bucketSummaries.map((b) => (
                      <tr key={b.key}>
                        <td className="py-2 px-2">{b.label}</td>
                        <td className="py-2 px-2 text-right font-mono">{fmt(b.receivablesPaise)}</td>
                        <td className="py-2 px-2 text-right font-mono">{fmt(b.payablesPaise)}</td>
                      </tr>
                    ))}
                  </tbody>
                </table>
              </section>
            )}
          </div>

          {/* PAGE 7 & 8 — FINANCIAL RATIOS & ACCOUNTING EXCEPTIONS */}
          <div className="grid grid-cols-1 lg:grid-cols-2 gap-8">
            {enabledSections[7] && (
              <section className="space-y-3">
                <h2 className="text-base font-bold text-[#0B2540] border-b border-slate-200 pb-2">
                  Page 7 · Financial Ratios
                </h2>
                <table className="w-full text-left border-collapse text-xs">
                  <tbody className="divide-y divide-slate-200">
                    {ratios.slice(0, 10).map((r) => (
                      <tr key={r.id}>
                        <td className="py-2 px-2">{r.name}</td>
                        <td className="py-2 px-2 text-slate-500">{r.category}</td>
                        <td className="py-2 px-2 text-right font-mono font-bold">
                          {r.formattedCurrent}
                        </td>
                      </tr>
                    ))}
                  </tbody>
                </table>
              </section>
            )}

            {enabledSections[8] && (
              <section className="space-y-3">
                <h2 className="text-base font-bold text-[#0B2540] border-b border-slate-200 pb-2">
                  Page 8 · Accounting Health & Audit Exceptions ({health.totalExceptions})
                </h2>
                <div className="space-y-2 text-xs">
                  {health.exceptions.slice(0, 6).map((e) => (
                    <div key={e.id} className="p-2.5 bg-slate-50 rounded border border-slate-200">
                      <div className="font-semibold text-[#0B2540]">
                        [{e.severity}] {e.category} — Voucher #{e.voucherNumber} ({fmt(e.amountPaise)})
                      </div>
                      <div className="text-slate-600 mt-0.5">{e.explanation}</div>
                    </div>
                  ))}
                </div>
              </section>
            )}
          </div>

          {/* PAGE 9 & 10 — MANAGEMENT INSIGHTS & DATA SOURCES / LIMITATIONS */}
          {enabledSections[9] && (
            <section className="space-y-3">
              <h2 className="text-base font-bold text-[#0B2540] border-b border-slate-200 pb-2">
                Page 9 · Calculated Management Insights
              </h2>
              <div className="grid grid-cols-1 md:grid-cols-2 gap-3 text-xs">
                {insights.map((i) => (
                  <div key={i.id} className="p-3 bg-slate-50 rounded border border-slate-200">
                    <div className="font-bold text-[#0B2540]">{i.headline}</div>
                    <div className="text-slate-600 mt-1">{i.narrative}</div>
                  </div>
                ))}
              </div>
            </section>
          )}

          {enabledSections[10] && (
            <section className="space-y-2 border-t border-slate-200 pt-6 text-xs text-slate-600">
              <h2 className="text-sm font-bold text-[#0B2540]">
                Page 10 · Data Sources, Provenance & Limitations
              </h2>
              <div>
                <strong>Uploaded Files Consolidated:</strong>{' '}
                {model.sourceFiles.map((f) => `${f.filename} (${f.recordsExtracted} records)`).join(', ')}
              </div>
              <div>
                <strong>Scope & Limitations:</strong> All metrics in this Management Information Report are deterministically computed in-browser from the uploaded TallyPrime and GST files. No figures have been fabricated or extrapolated.
              </div>
            </section>
          )}
        </div>
      )}
    </div>
  );
};
