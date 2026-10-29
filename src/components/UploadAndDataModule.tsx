import React, { useRef, useState } from 'react';
import { useNavigate } from 'react-router-dom';
import {
  AlertTriangle,
  ArrowLeft,
  ArrowRight,
  BarChart3,
  Building2,
  CheckCircle2,
  ClipboardCheck,
  CloudUpload,
  Eye,
  FileSpreadsheet,
  FileText,
  FlaskConical,
  FolderPlus,
  Lightbulb,
  Loader2,
  Paperclip,
  PieChart,
  RefreshCw,
  Settings2,
  ShieldCheck,
  Trash2,
  Users,
  X,
  Zap,
} from 'lucide-react';
import {
  ColumnMappingConfig,
  ConsolidatedAccountingModel,
  DetectedReportType,
  formatINR,
  SourceFile,
} from '../types/accounting';
import { RawFilePayload } from '../engine/reconciliationEngine';
import {
  downloadSampleStatementWorkbook,
  getSampleGSTR1Json,
  getSampleGSTR2BJson,
  getSampleTallyMastersXML,
  getSampleTallyTransactionsXML,
} from '../engine/sampleDataGenerator';
import { AutomatedTestResult, runAllAccountingAcceptanceTests } from '../engine/selfTestSuite';
import { GspuBrandLogo } from './GspuBrandLogo';

interface UploadAndDataModuleProps {
  mode: 'UPLOAD_WORKFLOW' | 'UPLOADED_DATA_INSPECTOR';
  companyName: string;
  onCompanyNameChange: (name: string) => void;
  payloads: RawFilePayload[];
  model: ConsolidatedAccountingModel;
  hasUnappliedUploads: boolean;
  isDemoActive: boolean;
  onUploadFiles: (files: FileList | File[]) => void;
  onRemoveFile: (fileId: string) => void;
  onReplaceFile: (fileId: string, file: File) => void;
  onApplyColumnMapping: (fileId: string, mapping: ColumnMappingConfig) => void;
  onClearAllData: () => void;
  onLoadDemoData: () => void;
  onViewDashboard: () => void;
  onOpenVoucherDrillDown: (voucherId: string) => void;
  onOpenLedgerDrillDown: (normalizedLedgerName: string) => void;
}

const REPORT_TYPE_LABELS: Record<DetectedReportType, string> = {
  TALLY_MASTERS_XML: 'Tally Masters XML',
  TALLY_TRANSACTIONS_XML: 'Tally Transactions XML',
  TALLY_COMBINED_XML: 'Tally Masters + Vouchers XML',
  TALLY_OUTSTANDING_XML: 'Tally Outstanding XML',
  TRIAL_BALANCE_EXCEL: 'Trial Balance Spreadsheet',
  PROFIT_LOSS_EXCEL: 'Profit & Loss Spreadsheet',
  BALANCE_SHEET_EXCEL: 'Balance Sheet Spreadsheet',
  CASH_FLOW_EXCEL: 'Cash Flow Statement Spreadsheet',
  COMPARATIVE_STATEMENTS_EXCEL: 'Comparative Financial Statements',
  BUDGET_EXCEL: 'Budget / Target Spreadsheet',
  CFO_MONTHLY_PACK_EXCEL: 'Multi-Sheet CFO Monthly Pack',
  SALES_REGISTER_EXCEL: 'Sales Register Spreadsheet',
  PURCHASE_REGISTER_EXCEL: 'Purchase Register Spreadsheet',
  DAYBOOK_EXCEL: 'DayBook Spreadsheet',
  OUTSTANDING_BILLS_EXCEL: 'Bill-wise Outstanding Spreadsheet',
  GSTR1_RETURN: 'GSTR-1 Outward Return',
  GSTR2B_RETURN: 'GSTR-2B ITC Statement',
  GSTR3B_RETURN: 'GSTR-3B Monthly Return',
  GST_REGISTER_EXCEL: 'GST Tax Register',
  CUSTOM_MAPPED_REPORT: 'Custom Mapped Report',
  UNKNOWN_UNSUPPORTED: 'Unrecognised / Needs Mapping',
};

const BOTTOM_CAPABILITY_CARDS = [
  {
    title: 'Accounting Health Check',
    icon: ClipboardCheck,
    iconBg: 'bg-teal-50 text-[#0F9B91]',
  },
  {
    title: 'Debtors & Creditors Ageing',
    icon: Users,
    iconBg: 'bg-indigo-50 text-indigo-600',
  },
  {
    title: 'Comparative Financial Statements',
    icon: BarChart3,
    iconBg: 'bg-blue-50 text-[#0867F5]',
  },
  {
    title: 'Ledger & Voucher Drill-Down',
    icon: FileText,
    iconBg: 'bg-emerald-50 text-[#50AE5D]',
  },
  {
    title: 'Automatic MIS Report',
    icon: FileSpreadsheet,
    iconBg: 'bg-green-50 text-[#50AE5D]',
  },
  {
    title: 'Financial Ratio Analysis',
    icon: PieChart,
    iconBg: 'bg-sky-50 text-[#0867F5]',
  },
  {
    title: 'GST Reconciliation',
    icon: FileText,
    iconBg: 'bg-blue-50 text-[#003087]',
  },
  {
    title: 'Automated Financial Insights',
    icon: Lightbulb,
    iconBg: 'bg-amber-50 text-amber-500',
  },
];

export const UploadAndDataModule: React.FC<UploadAndDataModuleProps> = ({
  mode,
  companyName,
  onCompanyNameChange,
  payloads,
  model,
  hasUnappliedUploads,
  isDemoActive,
  onUploadFiles,
  onRemoveFile,
  onReplaceFile,
  onApplyColumnMapping,
  onClearAllData,
  onLoadDemoData,
  onViewDashboard,
  onOpenVoucherDrillDown,
  onOpenLedgerDrillDown,
}) => {
  const navigate = useNavigate();
  const fileInputRef = useRef<HTMLInputElement | null>(null);
  const replaceInputRef = useRef<HTMLInputElement | null>(null);
  const [replaceTargetId, setReplaceTargetId] = useState<string | null>(null);
  const [isDragging, setIsDragging] = useState(false);
  const [isProcessingDashboard, setIsProcessingDashboard] = useState(false);

  // Preview Modal state
  const [previewFileId, setPreviewFileId] = useState<string | null>(null);
  // Column Mapping Modal state
  const [mappingFile, setMappingFile] = useState<SourceFile | null>(null);
  const [mappingDraft, setMappingDraft] = useState<ColumnMappingConfig>({});

  // Automated Self-Test Suite state
  const [testResults, setTestResults] = useState<AutomatedTestResult[] | null>(null);
  const [showReviewMappingsModal, setShowReviewMappingsModal] = useState(false);

  // Uploaded Data Inspector Tab
  const [inspectorTab, setInspectorTab] = useState<'LEDGERS' | 'VOUCHERS' | 'BILLS' | 'GST'>(
    'LEDGERS'
  );
  const [inspectorSearch, setInspectorSearch] = useState('');

  const previewPayload = payloads.find((p) => p.sourceFile.id === previewFileId) || null;
  const validFilesCount = payloads.filter((p) => p.sourceFile.status !== 'ERROR').length;
  const canViewDashboard = companyName.trim().length > 0 && validFilesCount > 0;

  const downloadSampleFile = (filename: string, content: string, mime = 'application/xml') => {
    const blob = new Blob([content], { type: mime });
    const url = URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url;
    a.download = filename;
    a.click();
    URL.revokeObjectURL(url);
  };

  const handleTriggerViewDashboard = () => {
    if (!canViewDashboard || isProcessingDashboard) return;
    setIsProcessingDashboard(true);
    window.setTimeout(() => {
      setIsProcessingDashboard(false);
      onViewDashboard();
    }, 280);
  };

  // ============================================================================
  // MODE 1: UPLOADED DATA INSPECTOR (Used inside Dashboard -> Uploaded Data tab)
  // ============================================================================
  if (mode === 'UPLOADED_DATA_INSPECTOR') {
    const ledgerList = Array.from(model.ledgerSummaries.values()).filter((l) =>
      !inspectorSearch
        ? true
        : l.ledger.name.toLowerCase().includes(inspectorSearch.toLowerCase()) ||
          l.ledger.parentGroup.toLowerCase().includes(inspectorSearch.toLowerCase())
    );
    const voucherList = model.deduplicatedVouchers.filter((v) =>
      !inspectorSearch
        ? true
        : `${v.voucherNumber} ${v.partyName || ''} ${v.voucherType}`
            .toLowerCase()
            .includes(inspectorSearch.toLowerCase())
    );

    return (
      <div className="space-y-6">
        <div className="bg-white rounded-xl border border-slate-200 p-6">
          <div className="flex flex-wrap items-center justify-between gap-4">
            <div>
              <h2 className="text-xl font-bold text-[#0B2540]">
                Uploaded Accounting Data Repository
              </h2>
              <p className="text-xs text-slate-500 mt-1">
                Inspect normalized ledgers, deduplicated vouchers, bill-level allocations, and GST invoices with full source provenance.
              </p>
            </div>
            <div className="flex items-center gap-2">
              <input
                type="text"
                value={inspectorSearch}
                onChange={(e) => setInspectorSearch(e.target.value)}
                placeholder="Filter records..."
                className="px-3 py-1.5 text-xs border border-slate-300 rounded-lg focus:outline-none focus:border-[#0867F5]"
              />
              <button
                onClick={() => navigate('/upload')}
                className="px-3 py-1.5 text-xs font-semibold text-white bg-[#0867F5] hover:bg-[#0052CC] rounded-lg flex items-center gap-1.5 cursor-pointer"
              >
                <FolderPlus className="w-3.5 h-3.5" />
                Manage / Add Files
              </button>
            </div>
          </div>

          <div className="flex items-center gap-1 p-1 bg-slate-100 rounded-lg w-fit mt-5">
            <button
              onClick={() => setInspectorTab('LEDGERS')}
              className={`px-3.5 py-1.5 text-xs font-semibold rounded-md transition-colors whitespace-nowrap ${
                inspectorTab === 'LEDGERS'
                  ? 'bg-white text-[#0B2540] shadow-xs'
                  : 'text-slate-600 hover:text-slate-900'
              }`}
            >
              Normalized Ledgers ({model.ledgerSummaries.size})
            </button>
            <button
              onClick={() => setInspectorTab('VOUCHERS')}
              className={`px-3.5 py-1.5 text-xs font-semibold rounded-md transition-colors whitespace-nowrap ${
                inspectorTab === 'VOUCHERS'
                  ? 'bg-white text-[#0B2540] shadow-xs'
                  : 'text-slate-600 hover:text-slate-900'
              }`}
            >
              Deduplicated Vouchers ({model.deduplicatedVouchers.length})
            </button>
            <button
              onClick={() => setInspectorTab('BILLS')}
              className={`px-3.5 py-1.5 text-xs font-semibold rounded-md transition-colors whitespace-nowrap ${
                inspectorTab === 'BILLS'
                  ? 'bg-white text-[#0B2540] shadow-xs'
                  : 'text-slate-600 hover:text-slate-900'
              }`}
            >
              Bill Allocations ({model.outstandingInvoices.length})
            </button>
            <button
              onClick={() => setInspectorTab('GST')}
              className={`px-3.5 py-1.5 text-xs font-semibold rounded-md transition-colors whitespace-nowrap ${
                inspectorTab === 'GST'
                  ? 'bg-white text-[#0B2540] shadow-xs'
                  : 'text-slate-600 hover:text-slate-900'
              }`}
            >
              GST Invoices ({model.gstInvoices.length})
            </button>
          </div>

          <div className="mt-4 border border-slate-200 rounded-lg overflow-x-auto">
            {inspectorTab === 'LEDGERS' && (
              <table className="w-full text-left border-collapse text-xs">
                <thead>
                  <tr className="bg-slate-50 border-b border-slate-200 text-slate-700">
                    <th className="py-2.5 px-3 font-semibold">Ledger Name</th>
                    <th className="py-2.5 px-3 font-semibold">Source Cell</th>
                    <th className="py-2.5 px-3 font-semibold">Tally Group</th>
                    <th className="py-2.5 px-3 font-semibold">Mapped Schedule</th>
                    <th className="py-2.5 px-3 font-semibold text-right">Opening Balance</th>
                    <th className="py-2.5 px-3 font-semibold text-right">Period Debits</th>
                    <th className="py-2.5 px-3 font-semibold text-right">Period Credits</th>
                    <th className="py-2.5 px-3 font-semibold text-right">Closing Balance</th>
                    <th className="py-2.5 px-3 font-semibold text-right">Comparative Prior</th>
                  </tr>
                </thead>
                <tbody className="divide-y divide-slate-200">
                  {ledgerList.map((l) => (
                    <tr
                      key={l.ledger.normalizedName}
                      onClick={() => onOpenLedgerDrillDown(l.ledger.normalizedName)}
                      className="hover:bg-blue-50/50 cursor-pointer"
                    >
                      <td className="py-2.5 px-3 font-medium text-[#0867F5]">{l.ledger.name}</td>
                      <td className="py-2.5 px-3 font-mono text-[11px] text-slate-500">
                        {l.ledger.cellReference || '—'}
                      </td>
                      <td className="py-2.5 px-3 text-slate-600">{l.ledger.parentGroup}</td>
                      <td className="py-2.5 px-3 text-slate-600 font-mono text-[11px]">
                        {l.ledger.classification}
                      </td>
                      <td className="py-2.5 px-3 text-right font-mono tabular-nums">
                        {formatINR(l.openingBalancePaise, false, true)}
                      </td>
                      <td className="py-2.5 px-3 text-right font-mono tabular-nums">
                        {formatINR(l.periodDebitPaise)}
                      </td>
                      <td className="py-2.5 px-3 text-right font-mono tabular-nums">
                        {formatINR(l.periodCreditPaise)}
                      </td>
                      <td className="py-2.5 px-3 text-right font-mono tabular-nums font-semibold">
                        {formatINR(l.effectiveClosingPaise, false, true)}
                      </td>
                      <td className="py-2.5 px-3 text-right font-mono tabular-nums text-slate-600">
                        {l.ledger.previousPeriodPaise !== undefined
                          ? formatINR(l.ledger.previousPeriodPaise)
                          : '—'}
                      </td>
                    </tr>
                  ))}
                </tbody>
              </table>
            )}

            {inspectorTab === 'VOUCHERS' && (
              <table className="w-full text-left border-collapse text-xs">
                <thead>
                  <tr className="bg-slate-50 border-b border-slate-200 text-slate-700">
                    <th className="py-2.5 px-3 font-semibold">Date</th>
                    <th className="py-2.5 px-3 font-semibold">Voucher No</th>
                    <th className="py-2.5 px-3 font-semibold">Type</th>
                    <th className="py-2.5 px-3 font-semibold">Party</th>
                    <th className="py-2.5 px-3 font-semibold text-right">Debit (₹)</th>
                    <th className="py-2.5 px-3 font-semibold text-right">Credit (₹)</th>
                    <th className="py-2.5 px-3 font-semibold">Source File</th>
                  </tr>
                </thead>
                <tbody className="divide-y divide-slate-200">
                  {voucherList.map((v) => (
                    <tr
                      key={v.id}
                      onClick={() => onOpenVoucherDrillDown(v.id)}
                      className="hover:bg-blue-50/50 cursor-pointer"
                    >
                      <td className="py-2.5 px-3 font-mono">{v.date}</td>
                      <td className="py-2.5 px-3 font-mono font-medium text-[#0867F5]">
                        {v.voucherNumber}
                      </td>
                      <td className="py-2.5 px-3">{v.voucherType}</td>
                      <td className="py-2.5 px-3">{v.partyName || '—'}</td>
                      <td className="py-2.5 px-3 text-right font-mono tabular-nums">
                        {formatINR(v.totalDebitPaise)}
                      </td>
                      <td className="py-2.5 px-3 text-right font-mono tabular-nums">
                        {formatINR(v.totalCreditPaise)}
                      </td>
                      <td className="py-2.5 px-3 font-mono text-[11px] text-slate-500">
                        {v.sourceFilename}
                      </td>
                    </tr>
                  ))}
                </tbody>
              </table>
            )}

            {inspectorTab === 'BILLS' && (
              <table className="w-full text-left border-collapse text-xs">
                <thead>
                  <tr className="bg-slate-50 border-b border-slate-200 text-slate-700">
                    <th className="py-2.5 px-3 font-semibold">Party Name</th>
                    <th className="py-2.5 px-3 font-semibold">Category</th>
                    <th className="py-2.5 px-3 font-semibold">Bill Reference</th>
                    <th className="py-2.5 px-3 font-semibold">Bill Date</th>
                    <th className="py-2.5 px-3 font-semibold">Due Date</th>
                    <th className="py-2.5 px-3 font-semibold text-right">Original Amount</th>
                    <th className="py-2.5 px-3 font-semibold text-right">Adjusted</th>
                    <th className="py-2.5 px-3 font-semibold text-right">Pending Balance</th>
                  </tr>
                </thead>
                <tbody className="divide-y divide-slate-200">
                  {model.outstandingInvoices.map((inv) => (
                    <tr key={inv.id} className="hover:bg-slate-50">
                      <td className="py-2.5 px-3 font-medium">{inv.partyName}</td>
                      <td className="py-2.5 px-3 font-mono text-[11px]">{inv.category}</td>
                      <td className="py-2.5 px-3 font-mono">{inv.billNumber}</td>
                      <td className="py-2.5 px-3 font-mono">{inv.billDate}</td>
                      <td className="py-2.5 px-3 font-mono">{inv.dueDate}</td>
                      <td className="py-2.5 px-3 text-right font-mono tabular-nums">
                        {formatINR(inv.originalAmountPaise)}
                      </td>
                      <td className="py-2.5 px-3 text-right font-mono tabular-nums">
                        {formatINR(inv.adjustedAmountPaise)}
                      </td>
                      <td className="py-2.5 px-3 text-right font-mono tabular-nums font-semibold text-[#0B2540]">
                        {formatINR(inv.pendingAmountPaise)}
                      </td>
                    </tr>
                  ))}
                </tbody>
              </table>
            )}

            {inspectorTab === 'GST' && (
              <table className="w-full text-left border-collapse text-xs">
                <thead>
                  <tr className="bg-slate-50 border-b border-slate-200 text-slate-700">
                    <th className="py-2.5 px-3 font-semibold">Source</th>
                    <th className="py-2.5 px-3 font-semibold">Doc Type</th>
                    <th className="py-2.5 px-3 font-semibold">GSTIN</th>
                    <th className="py-2.5 px-3 font-semibold">Party Name</th>
                    <th className="py-2.5 px-3 font-semibold">Invoice No</th>
                    <th className="py-2.5 px-3 font-semibold">Date</th>
                    <th className="py-2.5 px-3 font-semibold text-right">Taxable Value</th>
                    <th className="py-2.5 px-3 font-semibold text-right">Total GST</th>
                  </tr>
                </thead>
                <tbody className="divide-y divide-slate-200">
                  {model.gstInvoices.map((gi) => (
                    <tr key={gi.id} className="hover:bg-slate-50">
                      <td className="py-2.5 px-3 font-mono text-[11px]">{gi.source}</td>
                      <td className="py-2.5 px-3">{gi.documentType}</td>
                      <td className="py-2.5 px-3 font-mono">{gi.partyGstin || 'Unregistered'}</td>
                      <td className="py-2.5 px-3">{gi.partyName}</td>
                      <td className="py-2.5 px-3 font-mono font-medium">{gi.invoiceNumber}</td>
                      <td className="py-2.5 px-3 font-mono">{gi.invoiceDate}</td>
                      <td className="py-2.5 px-3 text-right font-mono tabular-nums">
                        {formatINR(gi.taxableValuePaise)}
                      </td>
                      <td className="py-2.5 px-3 text-right font-mono tabular-nums font-semibold">
                        {formatINR(gi.totalTaxPaise)}
                      </td>
                    </tr>
                  ))}
                </tbody>
              </table>
            )}
          </div>
        </div>
      </div>
    );
  }

  // ============================================================================
  // MODE 2: PAGE 2 — COMPANY SETUP & FILE UPLOAD (/upload)
  // Recreating Reference Image 2 with full accounting & file parsing workflow
  // ============================================================================
  return (
    <div className="relative min-h-screen w-full overflow-x-hidden bg-gradient-to-b from-white via-[#F6FAFF] to-[#EDF5FF] text-[#0B2540]">
      {/* Decorative Green-to-Blue Sweeping Wave Lines (Matching Reference Image 2) */}
      <svg
        aria-hidden="true"
        viewBox="0 0 1600 900"
        preserveAspectRatio="none"
        className="pointer-events-none absolute inset-0 w-full h-full"
      >
        <defs>
          <linearGradient id="uploadWaveGreenBlue" x1="0%" y1="30%" x2="100%" y2="10%">
            <stop offset="0%" stopColor="#50AE5D" stopOpacity="0.38" />
            <stop offset="45%" stopColor="#12A89D" stopOpacity="0.28" />
            <stop offset="100%" stopColor="#0867F5" stopOpacity="0.38" />
          </linearGradient>
          <linearGradient id="uploadSoftFill" x1="0%" y1="0%" x2="100%" y2="100%">
            <stop offset="0%" stopColor="#E6F7F2" stopOpacity="0.45" />
            <stop offset="50%" stopColor="#EBF5FF" stopOpacity="0.35" />
            <stop offset="100%" stopColor="#D8ECFF" stopOpacity="0.5" />
          </linearGradient>
        </defs>
        <path
          d="M-50,240 C340,210 680,120 1120,30 C1340,-15 1510,20 1660,180 L1660,420 C1360,260 980,290 540,320 C240,340 60,310 -50,280 Z"
          fill="url(#uploadSoftFill)"
        />
        <path
          d="M-40,255 C360,220 710,120 1150,25 C1370,-20 1520,25 1650,205"
          fill="none"
          stroke="url(#uploadWaveGreenBlue)"
          strokeWidth="2"
        />
        <path
          d="M-40,290 C380,250 740,155 1170,55 C1390,5 1530,50 1650,235"
          fill="none"
          stroke="url(#uploadWaveGreenBlue)"
          strokeWidth="1.4"
          opacity="0.7"
        />
      </svg>

      <div className="relative z-10 max-w-[1380px] mx-auto px-4 sm:px-6 lg:px-8 py-6 sm:py-8">
        {/* TOP HEADER (Matching Reference Image 2) */}
        <header className="flex flex-col lg:flex-row lg:items-start lg:justify-between gap-4 pb-6">
          <div className="flex items-start gap-4">
            <GspuBrandLogo iconOnly whiteTile onClick={() => navigate('/')} />
            <div>
              <div className="flex flex-wrap items-center gap-1.5 text-xs sm:text-sm font-semibold text-[#0F9B91]">
                <span>TallyPrime Financial Reporting</span>
                <span aria-hidden="true">·</span>
                <span>Audit Health Check</span>
                <span aria-hidden="true">·</span>
                <span>GST Reconciliation</span>
                <span aria-hidden="true">·</span>
                <span>10-Page MIS</span>
              </div>
              <h1 className="text-2xl sm:text-3xl lg:text-[34px] font-extrabold tracking-tight text-[#0B2540] mt-1">
                GSPU Financial Analytics &amp; Audit Intelligence
              </h1>
              <p className="text-xs sm:text-sm text-slate-600 max-w-2xl mt-1 leading-relaxed">
                Consolidate multiple TallyPrime XML, Excel and GST files for a single company with
                deterministic paise-exact accounting calculations and six-level voucher drill-down.
              </p>
            </div>
          </div>

          {/* Subtle Top Utility Actions */}
          <div className="flex flex-wrap items-center gap-2 shrink-0 self-start">
            <button
              type="button"
              onClick={() => navigate('/')}
              className="px-3 py-1.5 text-xs font-semibold text-slate-600 hover:text-[#003087] bg-white/80 hover:bg-white border border-slate-200 rounded-lg flex items-center gap-1.5 shadow-2xs transition-colors cursor-pointer"
            >
              <ArrowLeft className="w-3.5 h-3.5" />
              Home
            </button>
            <button
              type="button"
              onClick={onLoadDemoData}
              className="px-3 py-1.5 text-xs font-semibold text-[#003087] bg-blue-50/90 hover:bg-blue-100 border border-blue-200 rounded-lg transition-colors cursor-pointer"
            >
              Load Sample Demo Data
            </button>
            <button
              type="button"
              onClick={() => setTestResults(runAllAccountingAcceptanceTests())}
              className="px-3 py-1.5 text-xs font-semibold text-slate-700 bg-white/80 hover:bg-white border border-slate-200 rounded-lg flex items-center gap-1.5 transition-colors cursor-pointer"
            >
              <FlaskConical className="w-3.5 h-3.5 text-[#0F9B91]" />
              26 Self-Tests
            </button>
          </div>
        </header>

        {/* MAIN TWO-COLUMN UPLOAD & ILLUSTRATION GRID */}
        <div className="grid grid-cols-1 lg:grid-cols-12 gap-8 items-start mt-2">
          {/* LEFT COLUMN (7 Cols): GET STARTED WHITE CARD + VIEW DASHBOARD CTA */}
          <div className="lg:col-span-7 flex flex-col">
            <div className="bg-white/95 backdrop-blur-md rounded-2xl border border-blue-100/90 shadow-[0_14px_45px_rgba(11,37,64,0.06)] p-6 sm:p-8 space-y-6">
              <div>
                <h2 className="text-xl sm:text-2xl font-extrabold text-[#0B2540]">Get Started</h2>
                <p className="text-xs sm:text-sm text-slate-500 mt-1 leading-relaxed">
                  Enter your company name and upload TallyPrime, Excel or GST files to generate
                  financial insights, reconciliations and management reports.
                </p>
              </div>

              {/* COMPANY / CLIENT NAME INPUT */}
              <div>
                <label
                  htmlFor="company-client-name"
                  className="block text-xs sm:text-sm font-bold text-[#0B2540] mb-2"
                >
                  Company / Client Name
                </label>
                <div className="relative">
                  <Building2 className="w-4 h-4 text-slate-400 absolute left-3.5 top-1/2 -translate-y-1/2 pointer-events-none" />
                  <input
                    id="company-client-name"
                    type="text"
                    value={companyName}
                    onChange={(e) => onCompanyNameChange(e.target.value)}
                    placeholder="Enter company name"
                    className="w-full pl-10 pr-4 py-3 text-sm font-semibold text-[#0B2540] bg-white border border-slate-200 rounded-xl shadow-2xs focus:outline-none focus:border-[#0867F5] focus:ring-3 focus:ring-[#0867F5]/15 transition-all"
                  />
                </div>
                {model.detectedCompanyNames.length > 0 && (
                  <div className="mt-2 text-xs text-slate-500 flex items-center gap-2 flex-wrap">
                    <span>Detected in uploaded files:</span>
                    {model.detectedCompanyNames.map((cn, idx) => (
                      <button
                        key={idx}
                        type="button"
                        onClick={() => onCompanyNameChange(cn)}
                        className="font-mono text-[#0867F5] hover:underline font-semibold cursor-pointer"
                      >
                        {cn}
                      </button>
                    ))}
                  </div>
                )}

                {/* Company Mismatch Alert */}
                {model.hasCompanyMismatch && (
                  <div className="mt-3 p-3.5 rounded-xl bg-red-50 border border-red-200 flex items-start gap-2.5">
                    <AlertTriangle className="w-4 h-4 text-red-600 shrink-0 mt-0.5" />
                    <div className="text-xs text-red-900">
                      <div className="font-bold">
                        Multiple Company Names Detected Across Uploaded Files
                      </div>
                      <p className="mt-0.5">
                        Uploaded files reference:{' '}
                        <span className="font-mono font-semibold">
                          {model.detectedCompanyNames.join(' vs ')}
                        </span>
                        . Only one company dataset should be active at a time.
                      </p>
                    </div>
                  </div>
                )}

                {/* User-Entered vs Statement Company Mismatch Warning */}
                {model.hasUserEnteredCompanyMismatch && model.extractedStatementCompanyName && (
                  <div className="mt-3 p-3.5 rounded-xl bg-amber-50 border border-amber-200 flex items-start justify-between gap-2.5">
                    <div className="flex items-start gap-2.5">
                      <AlertTriangle className="w-4 h-4 text-amber-600 shrink-0 mt-0.5" />
                      <div className="text-xs text-amber-900">
                        <div className="font-bold">
                          Warning: Entered company name does not match the company identified in the uploaded financial statements.
                        </div>
                        <p className="mt-0.5">
                          Uploaded statement company:{' '}
                          <span className="font-mono font-bold">
                            {model.extractedStatementCompanyName}
                          </span>
                        </p>
                      </div>
                    </div>
                    <button
                      type="button"
                      onClick={() => onCompanyNameChange(model.extractedStatementCompanyName!)}
                      className="px-2.5 py-1 text-[11px] font-bold bg-amber-600 hover:bg-amber-700 text-white rounded-lg shrink-0 cursor-pointer"
                    >
                      Use Statement Name
                    </button>
                  </div>
                )}
              </div>

              {/* PRINCIPAL 3 FINANCIAL STATEMENTS UPLOAD CATEGORIES (PART 1 & PART 2) */}
              <div className="bg-[#F5F9FF] border border-[#B9E2FF] rounded-2xl p-4 space-y-3">
                <div className="flex flex-wrap items-center justify-between gap-2">
                  <div>
                    <div className="text-xs sm:text-sm font-bold text-[#003087]">
                      Principal Financial Statements (Upload Together or Separately)
                    </div>
                    <p className="text-[11px] text-slate-600">
                      Upload Balance Sheet, Profit &amp; Loss, and Cash Flow Statement Excel workbooks (or multi-sheet CFO pack / Tally XML) to generate the full CFO dashboard.
                    </p>
                  </div>
                  <button
                    type="button"
                    onClick={() => downloadSampleStatementWorkbook('CFO_PACK')}
                    className="px-3 py-1.5 rounded-lg bg-[#003087] hover:bg-[#0867F5] text-white text-[11px] font-bold transition cursor-pointer flex items-center gap-1.5"
                  >
                    <FileSpreadsheet className="w-3.5 h-3.5" />
                    Download 3-Statement CFO Pack (.xlsx)
                  </button>
                </div>

                <div className="grid grid-cols-1 sm:grid-cols-3 gap-2.5">
                  {([
                    {
                      id: 'BS',
                      title: '1. Balance Sheet',
                      subtitle: 'Assets, Equity & Liabilities',
                      sampleKind: 'BALANCE_SHEET' as const,
                      isUploaded:
                        model.ledgers.size > 0 ||
                        payloads.some((p) =>
                          ['BALANCE_SHEET_EXCEL', 'CFO_MONTHLY_PACK_EXCEL', 'TRIAL_BALANCE_EXCEL', 'TALLY_MASTERS_XML', 'TALLY_COMBINED_XML'].includes(
                            p.sourceFile.detectedReportType
                          )
                        ),
                    },
                    {
                      id: 'PNL',
                      title: '2. Profit & Loss Statement',
                      subtitle: 'Revenue, COGS, EBITDA & PAT',
                      sampleKind: 'PROFIT_LOSS' as const,
                      isUploaded:
                        model.ledgers.size > 0 ||
                        payloads.some((p) =>
                          ['PROFIT_LOSS_EXCEL', 'CFO_MONTHLY_PACK_EXCEL', 'TRIAL_BALANCE_EXCEL', 'TALLY_TRANSACTIONS_XML', 'TALLY_COMBINED_XML'].includes(
                            p.sourceFile.detectedReportType
                          )
                        ),
                    },
                    {
                      id: 'CF',
                      title: '3. Cash Flow Statement',
                      subtitle: 'Operating, Investing & Financing',
                      sampleKind: 'CASH_FLOW' as const,
                      isUploaded:
                        (model.cashFlowItems?.length ?? 0) > 0 ||
                        payloads.some((p) =>
                          ['CASH_FLOW_EXCEL', 'CFO_MONTHLY_PACK_EXCEL'].includes(p.sourceFile.detectedReportType)
                        ),
                    },
                  ]).map((st) => (
                    <div
                      key={st.id}
                      className={`p-3 rounded-xl border transition flex flex-col justify-between ${
                        st.isUploaded
                          ? 'bg-emerald-50/70 border-emerald-200'
                          : 'bg-white border-slate-200 hover:border-[#0867F5]'
                      }`}
                    >
                      <div>
                        <div className="flex items-center justify-between gap-1.5">
                          <span className="text-xs font-bold text-[#0B2540]">{st.title}</span>
                          {st.isUploaded ? (
                            <span className="inline-flex items-center gap-1 px-2 py-0.5 rounded-full bg-emerald-100 text-emerald-800 text-[10px] font-bold">
                              <CheckCircle2 className="w-3 h-3" /> Detected
                            </span>
                          ) : (
                            <span className="px-2 py-0.5 rounded-full bg-slate-100 text-slate-600 text-[10px] font-semibold">
                              Optional / Ready
                            </span>
                          )}
                        </div>
                        <p className="text-[11px] text-slate-500 mt-0.5">{st.subtitle}</p>
                      </div>
                      <div className="mt-2.5 pt-2 border-t border-slate-200/70 flex items-center justify-between gap-2">
                        <button
                          type="button"
                          onClick={() => fileInputRef.current?.click()}
                          className="text-[11px] font-bold text-[#0867F5] hover:underline cursor-pointer"
                        >
                          + Upload Excel
                        </button>
                        <button
                          type="button"
                          onClick={() => downloadSampleStatementWorkbook(st.sampleKind)}
                          className="text-[10px] font-semibold text-slate-500 hover:text-[#003087] cursor-pointer"
                        >
                          Sample .xlsx
                        </button>
                      </div>
                    </div>
                  ))}
                </div>
              </div>

              {/* FILE UPLOAD DRAG & DROP ZONE */}
              <div>
                <div className="flex flex-wrap items-center justify-between gap-2 mb-2">
                  <label className="block text-xs sm:text-sm font-bold text-[#0B2540]">
                    Upload Files
                  </label>
                  <div className="flex items-center gap-2 text-[11px]">
                    <span className="text-slate-400">Sample files:</span>
                    <button
                      type="button"
                      onClick={() =>
                        downloadSampleFile('Masters_Sample.xml', getSampleTallyMastersXML())
                      }
                      className="text-[#0867F5] hover:underline font-medium cursor-pointer"
                    >
                      Masters.xml
                    </button>
                    <span>·</span>
                    <button
                      type="button"
                      onClick={() =>
                        downloadSampleFile(
                          'Transactions_Sample.xml',
                          getSampleTallyTransactionsXML()
                        )
                      }
                      className="text-[#0867F5] hover:underline font-medium cursor-pointer"
                    >
                      Transactions.xml
                    </button>
                    <span>·</span>
                    <button
                      type="button"
                      onClick={() =>
                        downloadSampleFile(
                          'GSTR1_Sample.json',
                          getSampleGSTR1Json(),
                          'application/json'
                        )
                      }
                      className="text-[#0867F5] hover:underline font-medium cursor-pointer"
                    >
                      GSTR1.json
                    </button>
                    <span>·</span>
                    <button
                      type="button"
                      onClick={() =>
                        downloadSampleFile(
                          'GSTR2B_Sample.json',
                          getSampleGSTR2BJson(),
                          'application/json'
                        )
                      }
                      className="text-[#0867F5] hover:underline font-medium cursor-pointer"
                    >
                      GSTR2B.json
                    </button>
                  </div>
                </div>

                <div
                  onDragOver={(e) => {
                    e.preventDefault();
                    setIsDragging(true);
                  }}
                  onDragLeave={() => setIsDragging(false)}
                  onDrop={(e) => {
                    e.preventDefault();
                    setIsDragging(false);
                    if (e.dataTransfer.files && e.dataTransfer.files.length > 0) {
                      onUploadFiles(e.dataTransfer.files);
                    }
                  }}
                  onClick={() => fileInputRef.current?.click()}
                  className={`border-2 border-dashed rounded-2xl p-7 sm:p-8 text-center cursor-pointer transition-all ${
                    isDragging
                      ? 'border-[#0867F5] bg-blue-50/70 scale-[1.005]'
                      : 'border-[#B9E2FF] hover:border-[#0867F5] bg-[#F5F9FF]/90'
                  }`}
                >
                  <input
                    ref={fileInputRef}
                    type="file"
                    multiple
                    accept=".xml,.xlsx,.xls,.csv,.json"
                    onChange={(e) => {
                      if (e.target.files && e.target.files.length > 0) {
                        onUploadFiles(e.target.files);
                        e.target.value = '';
                      }
                    }}
                    className="hidden"
                  />
                  <input
                    ref={replaceInputRef}
                    type="file"
                    accept=".xml,.xlsx,.xls,.csv,.json"
                    onChange={(e) => {
                      if (e.target.files && e.target.files[0] && replaceTargetId) {
                        onReplaceFile(replaceTargetId, e.target.files[0]);
                        setReplaceTargetId(null);
                        e.target.value = '';
                      }
                    }}
                    className="hidden"
                  />

                  <div className="w-12 h-12 rounded-2xl bg-blue-100/80 text-[#0867F5] flex items-center justify-center mx-auto mb-3">
                    <CloudUpload className="w-7 h-7 stroke-[2.2]" />
                  </div>
                  <div className="text-sm sm:text-base font-bold text-[#0B2540]">
                    Drag &amp; Drop Your Files Here
                  </div>
                  <div className="text-xs text-slate-400 my-1">or</div>
                  <button
                    type="button"
                    onClick={(e) => {
                      e.stopPropagation();
                      fileInputRef.current?.click();
                    }}
                    className="mt-1 px-6 py-2.5 rounded-lg bg-[#0867F5] hover:bg-[#0052CC] text-white text-xs sm:text-sm font-bold tracking-wide shadow-sm transition-colors cursor-pointer"
                  >
                    CHOOSE FILES
                  </button>
                  <div className="mt-4 flex items-center justify-center gap-1.5 text-xs text-slate-500">
                    <Paperclip className="w-3.5 h-3.5 text-slate-400 shrink-0" />
                    <span>
                      Supports TallyPrime XML, Excel (XLS/XLSX), CSV, GSTR-1, GSTR-2B, GSTR-3B JSON
                    </span>
                  </div>
                </div>
              </div>

              {/* UPLOADED FILE LIST TABLE */}
              {payloads.length > 0 && (
                <div className="pt-2 space-y-3">
                  <div className="flex flex-wrap items-center justify-between gap-2">
                    <div className="flex items-center gap-2">
                      <span className="text-xs sm:text-sm font-bold text-[#0B2540]">
                        Uploaded Files ({payloads.length})
                      </span>
                      {isDemoActive && (
                        <span className="px-2 py-0.5 text-[11px] font-bold bg-amber-100 text-amber-800 rounded-md">
                          DEMO DATASET
                        </span>
                      )}
                    </div>
                    <div className="flex items-center gap-2">
                      <button
                        type="button"
                        onClick={() => fileInputRef.current?.click()}
                        className="px-3 py-1.5 text-xs font-semibold text-[#0867F5] bg-blue-50 hover:bg-blue-100 rounded-lg flex items-center gap-1.5 transition-colors cursor-pointer"
                      >
                        <FolderPlus className="w-3.5 h-3.5" />
                        Add More Files
                      </button>
                      <button
                        type="button"
                        onClick={onClearAllData}
                        className="px-3 py-1.5 text-xs font-semibold text-red-600 bg-red-50 hover:bg-red-100 rounded-lg flex items-center gap-1.5 transition-colors cursor-pointer"
                      >
                        <Trash2 className="w-3.5 h-3.5" />
                        Clear All Files
                      </button>
                    </div>
                  </div>

                  <div className="border border-slate-200 rounded-xl overflow-x-auto">
                    <table className="w-full text-left border-collapse text-xs">
                      <thead>
                        <tr className="bg-slate-50 text-slate-700 border-b border-slate-200">
                          <th className="py-2.5 px-3 font-semibold">Filename</th>
                          <th className="py-2.5 px-3 font-semibold">Format</th>
                          <th className="py-2.5 px-3 font-semibold">Detected Report</th>
                          <th className="py-2.5 px-3 font-semibold text-right">Records</th>
                          <th className="py-2.5 px-3 font-semibold">Status</th>
                          <th className="py-2.5 px-3 font-semibold text-right">Actions</th>
                        </tr>
                      </thead>
                      <tbody className="divide-y divide-slate-200">
                        {payloads.map(({ sourceFile: sf }) => (
                          <tr key={sf.id} className="hover:bg-slate-50/80">
                            <td className="py-2.5 px-3 font-medium text-[#0B2540]">
                              <div className="flex items-center gap-2">
                                {sf.format === 'XML' ? (
                                  <FileText className="w-4 h-4 text-[#50AE5D] shrink-0" />
                                ) : (
                                  <FileSpreadsheet className="w-4 h-4 text-[#0867F5] shrink-0" />
                                )}
                                <div className="truncate max-w-[180px]" title={sf.filename}>
                                  {sf.filename}
                                </div>
                              </div>
                            </td>
                            <td className="py-2.5 px-3 font-mono text-[11px]">{sf.format}</td>
                            <td className="py-2.5 px-3 text-slate-700">
                              {REPORT_TYPE_LABELS[sf.detectedReportType]}
                            </td>
                            <td className="py-2.5 px-3 text-right font-mono tabular-nums font-bold text-[#0B2540]">
                              {sf.recordsExtracted}
                            </td>
                            <td className="py-2.5 px-3">
                              {sf.status === 'VALID' && (
                                <span className="inline-flex items-center gap-1 text-emerald-700 font-semibold">
                                  <CheckCircle2 className="w-3.5 h-3.5" />
                                  Ready
                                </span>
                              )}
                              {sf.status === 'WARNING' && (
                                <span className="inline-flex items-center gap-1 text-amber-700 font-semibold">
                                  <AlertTriangle className="w-3.5 h-3.5" />
                                  Warning
                                </span>
                              )}
                              {sf.status === 'NEEDS_MAPPING' && (
                                <span className="inline-flex items-center gap-1 text-amber-700 font-semibold">
                                  <Settings2 className="w-3.5 h-3.5" />
                                  Needs Mapping
                                </span>
                              )}
                              {sf.status === 'ERROR' && (
                                <span className="inline-flex items-center gap-1 text-red-600 font-semibold">
                                  <AlertTriangle className="w-3.5 h-3.5" />
                                  Error
                                </span>
                              )}
                            </td>
                            <td className="py-2.5 px-3 text-right whitespace-nowrap">
                              <div className="inline-flex items-center gap-1.5">
                                <button
                                  type="button"
                                  onClick={() => setPreviewFileId(sf.id)}
                                  className="px-2 py-1 text-[11px] font-semibold text-[#0867F5] hover:bg-blue-50 rounded flex items-center gap-1 cursor-pointer"
                                >
                                  <Eye className="w-3 h-3" />
                                  Preview
                                </button>
                                {(sf.format === 'XLSX' ||
                                  sf.format === 'XLS' ||
                                  sf.format === 'CSV') && (
                                  <button
                                    type="button"
                                    onClick={() => {
                                      setMappingFile(sf);
                                      setMappingDraft(sf.columnMapping || {});
                                    }}
                                    className="px-2 py-1 text-[11px] font-semibold text-amber-700 hover:bg-amber-50 rounded flex items-center gap-1 cursor-pointer"
                                  >
                                    <Settings2 className="w-3 h-3" />
                                    Map
                                  </button>
                                )}
                                <button
                                  type="button"
                                  onClick={() => {
                                    setReplaceTargetId(sf.id);
                                    replaceInputRef.current?.click();
                                  }}
                                  className="px-2 py-1 text-[11px] font-semibold text-slate-600 hover:bg-slate-100 rounded flex items-center gap-1 cursor-pointer"
                                >
                                  <RefreshCw className="w-3 h-3" />
                                  Replace
                                </button>
                                <button
                                  type="button"
                                  onClick={() => onRemoveFile(sf.id)}
                                  className="px-2 py-1 text-[11px] font-semibold text-red-600 hover:bg-red-50 rounded flex items-center gap-1 cursor-pointer"
                                >
                                  <Trash2 className="w-3 h-3" />
                                  Remove
                                </button>
                              </div>
                            </td>
                          </tr>
                        ))}
                      </tbody>
                    </table>
                  </div>

                  {/* COMPACT VALIDATION SUMMARY & REVIEW MAPPINGS */}
                  <div className="p-4 rounded-xl bg-slate-50 border border-slate-200 space-y-2.5">
                    <div className="flex flex-wrap items-center justify-between gap-2">
                      <div className="text-xs font-bold text-[#003087] flex items-center gap-1.5">
                        <CheckCircle2 className="w-4 h-4 text-[#50AE5D]" />
                        <span>Validation Summary</span>
                      </div>
                      <button
                        type="button"
                        onClick={() => setShowReviewMappingsModal(true)}
                        className="px-3 py-1 rounded-lg bg-white hover:bg-blue-50 text-[#0867F5] border border-blue-200 text-[11px] font-bold transition cursor-pointer flex items-center gap-1.5"
                      >
                        <Settings2 className="w-3.5 h-3.5" />
                        REVIEW MAPPINGS ({model.ledgers.size} Items)
                      </button>
                    </div>
                    <div className="grid grid-cols-1 sm:grid-cols-2 lg:grid-cols-3 gap-2 text-[11px]">
                      <div className="px-2.5 py-1.5 rounded-lg bg-white border border-slate-200 flex items-center justify-between">
                        <span className="text-slate-500 font-medium">P&amp;L Status:</span>
                        <span className="font-bold text-emerald-700">
                          {model.hasProfitLossStatement || model.ledgers.size > 0 ? 'Reconciled' : 'Pending'}
                        </span>
                      </div>
                      <div className="px-2.5 py-1.5 rounded-lg bg-white border border-slate-200 flex items-center justify-between">
                        <span className="text-slate-500 font-medium">Balance Sheet Status:</span>
                        <span className="font-bold text-emerald-700">
                          {model.hasBalanceSheetStatement || model.ledgers.size > 0 ? 'Reconciled' : 'Pending'}
                        </span>
                      </div>
                      <div className="px-2.5 py-1.5 rounded-lg bg-white border border-slate-200 flex items-center justify-between">
                        <span className="text-slate-500 font-medium">Cash Flow Status:</span>
                        <span className="font-bold text-emerald-700">
                          {(model.cashFlowItems?.length ?? 0) > 0 || model.vouchers.length > 0
                            ? 'Reconciled'
                            : 'Derived'}
                        </span>
                      </div>
                      <div className="px-2.5 py-1.5 rounded-lg bg-white border border-slate-200 flex items-center justify-between">
                        <span className="text-slate-500 font-medium">Company Match Status:</span>
                        <span
                          className={`font-bold ${
                            model.hasUserEnteredCompanyMismatch ? 'text-amber-700' : 'text-emerald-700'
                          }`}
                        >
                          {model.hasUserEnteredCompanyMismatch ? 'Mismatch Warning' : 'Verified'}
                        </span>
                      </div>
                      <div className="px-2.5 py-1.5 rounded-lg bg-white border border-slate-200 sm:col-span-2 flex items-center justify-between">
                        <span className="text-slate-500 font-medium">Comparative Period Status:</span>
                        <span className="font-bold text-[#003087]">
                          {model.comparativePeriodLabel
                            ? `${model.currentPeriodLabel || 'FY 2025-26'} vs ${model.comparativePeriodLabel} Detected`
                            : `${model.currentPeriodLabel || 'Current Period'} Detected`}
                        </span>
                      </div>
                    </div>
                  </div>
                </div>
              )}
            </div>

            {/* PROMINENT VIEW DASHBOARD BUTTON BELOW THE UPLOAD CARD */}
            <div className="mt-5 flex flex-col items-center">
              <button
                type="button"
                disabled={!canViewDashboard || isProcessingDashboard}
                onClick={handleTriggerViewDashboard}
                className={`w-full sm:w-auto sm:min-w-[360px] px-10 py-3.5 rounded-xl text-sm sm:text-base font-bold flex items-center justify-center gap-2.5 transition-all ${
                  canViewDashboard
                    ? 'bg-gradient-to-r from-[#0867F5] to-[#0052CC] hover:from-[#0052CC] hover:to-[#003E99] text-white shadow-[0_10px_28px_rgba(8,103,245,0.35)] cursor-pointer hover:scale-[1.01]'
                    : 'bg-slate-200 text-slate-400 cursor-not-allowed'
                }`}
              >
                {isProcessingDashboard ? (
                  <>
                    <Loader2 className="w-5 h-5 animate-spin" />
                    <span>Reconciling &amp; Building Dashboard...</span>
                  </>
                ) : (
                  <>
                    <span>{hasUnappliedUploads ? 'View Dashboard' : 'View Dashboard'}</span>
                    <ArrowRight className="w-5 h-5" />
                  </>
                )}
              </button>
              {!canViewDashboard && (
                <p className="text-[11px] text-slate-500 mt-2 text-center">
                  Enter a company name and upload at least one XML, Excel, or GST file (or click{' '}
                  <button
                    type="button"
                    onClick={onLoadDemoData}
                    className="text-[#0867F5] font-semibold hover:underline cursor-pointer"
                  >
                    Load Sample Demo Data
                  </button>
                  ) to enable View Dashboard.
                </p>
              )}
            </div>
          </div>

          {/* RIGHT COLUMN (5 Cols): LAPTOP ILLUSTRATION + FLOATING XML/XLS/GST BADGES + 3 INFO CARDS */}
          <div className="lg:col-span-5 flex flex-col space-y-5">
            {/* Laptop & Floating File Icons Visual (Matching Reference Image 2) */}
            <div className="relative pt-2 pb-4 px-2 flex items-center justify-center">
              {/* Soft Cyan Glow Behind Laptop */}
              <div
                aria-hidden="true"
                className="pointer-events-none absolute inset-0 rounded-3xl blur-2xl opacity-70"
                style={{
                  background:
                    'radial-gradient(circle at 58% 48%, rgba(85, 200, 245, 0.42) 0%, rgba(8, 103, 245, 0.12) 55%, transparent 80%)',
                }}
              />

              {/* Floating File Badges (XML, Excel, GST) on Left */}
              <div className="relative z-20 flex flex-col gap-3 -mr-5 sm:-mr-8">
                {/* XML Card */}
                <div className="bg-white/95 backdrop-blur-md rounded-2xl p-3 shadow-[0_10px_28px_rgba(16,185,129,0.18)] border border-emerald-100 flex items-center justify-center transform -rotate-3 translate-x-4">
                  <div className="w-11 h-12 rounded-xl bg-gradient-to-b from-[#50AE5D] to-[#2E8B3C] text-white flex flex-col items-center justify-center shadow-xs">
                    <FileText className="w-5 h-5 mb-0.5" />
                    <span className="text-[10px] font-extrabold tracking-wider">XML</span>
                  </div>
                </div>

                {/* Excel XLS Card */}
                <div className="bg-white/95 backdrop-blur-md rounded-2xl p-3 shadow-[0_10px_28px_rgba(15,155,145,0.18)] border border-teal-100 flex items-center justify-center transform rotate-2 -translate-x-2">
                  <div className="w-11 h-12 rounded-xl bg-gradient-to-b from-[#0F9B91] to-[#09736B] text-white flex flex-col items-center justify-center shadow-xs">
                    <FileSpreadsheet className="w-5 h-5 mb-0.5" />
                    <span className="text-[10px] font-extrabold tracking-wider">XLS</span>
                  </div>
                </div>

                {/* GST Card */}
                <div className="bg-white/95 backdrop-blur-md rounded-2xl p-3 shadow-[0_10px_28px_rgba(8,103,245,0.20)] border border-blue-100 flex items-center justify-center transform -rotate-2 translate-x-5">
                  <div className="w-11 h-12 rounded-xl bg-gradient-to-b from-[#0867F5] to-[#003087] text-white flex flex-col items-center justify-center shadow-xs">
                    <FileText className="w-5 h-5 mb-0.5" />
                    <span className="text-[10px] font-extrabold tracking-wider">GST</span>
                  </div>
                </div>
              </div>

              {/* Flowing Cyan Light Streams into Laptop */}
              <svg
                aria-hidden="true"
                viewBox="0 0 90 160"
                className="w-14 sm:w-20 h-36 z-10 shrink-0"
              >
                <defs>
                  <linearGradient id="streamCyan" x1="0%" y1="50%" x2="100%" y2="50%">
                    <stop offset="0%" stopColor="#50AE5D" stopOpacity="0.6" />
                    <stop offset="50%" stopColor="#55C8F5" stopOpacity="0.85" />
                    <stop offset="100%" stopColor="#0867F5" stopOpacity="0.9" />
                  </linearGradient>
                </defs>
                <path
                  d="M4,28 C35,28 50,55 88,65"
                  fill="none"
                  stroke="url(#streamCyan)"
                  strokeWidth="2.5"
                  strokeDasharray="4 3"
                />
                <path
                  d="M0,80 C35,80 55,80 88,80"
                  fill="none"
                  stroke="url(#streamCyan)"
                  strokeWidth="3"
                />
                <path
                  d="M6,132 C35,132 50,105 88,95"
                  fill="none"
                  stroke="url(#streamCyan)"
                  strokeWidth="2.5"
                  strokeDasharray="4 3"
                />
              </svg>

              {/* Sleek Laptop Mockup Illustrating Live Charts */}
              <div className="relative z-10 w-full max-w-[330px]">
                {/* Laptop Screen Bezel */}
                <div className="bg-[#0B192C] rounded-t-2xl p-2.5 shadow-2xl border-2 border-slate-800">
                  <div className="bg-white rounded-lg p-3 space-y-2.5 overflow-hidden">
                    {/* Mini Top Bar */}
                    <div className="flex items-center justify-between border-b border-slate-100 pb-1.5">
                      <div className="flex items-center gap-1.5">
                        <span className="w-2 h-2 rounded-full bg-[#50AE5D]" />
                        <span className="text-[9px] font-bold text-[#003087]">GSPU BI Engine</span>
                      </div>
                      <div className="flex gap-1">
                        <span className="h-1.5 w-6 rounded-full bg-blue-100" />
                        <span className="h-1.5 w-4 rounded-full bg-emerald-100" />
                      </div>
                    </div>

                    {/* Charts Row 1: Bar Chart + Donut Chart */}
                    <div className="grid grid-cols-12 gap-2 items-center">
                      {/* Mini Bar Chart */}
                      <div className="col-span-7 bg-slate-50/80 rounded-md p-2 border border-slate-100">
                        <div className="text-[8px] font-bold text-slate-500 mb-1.5">
                          Monthly Revenue &amp; Profit
                        </div>
                        <div className="h-14 flex items-end justify-between gap-1 px-1">
                          {[42, 65, 52, 78, 64, 88, 95].map((h, i) => (
                            <div key={i} className="flex-1 flex flex-col items-center gap-0.5">
                              <div
                                className="w-full rounded-t bg-gradient-to-t from-[#0867F5] to-[#50AE5D]"
                                style={{ height: `${Math.round(h * 0.48)}px` }}
                              />
                            </div>
                          ))}
                        </div>
                      </div>

                      {/* Mini Donut Chart */}
                      <div className="col-span-5 bg-slate-50/80 rounded-md p-2 border border-slate-100 flex flex-col items-center justify-center">
                        <svg viewBox="0 0 36 36" className="w-14 h-14">
                          <circle
                            cx="18"
                            cy="18"
                            r="12"
                            fill="none"
                            stroke="#E2E8F0"
                            strokeWidth="5"
                          />
                          <circle
                            cx="18"
                            cy="18"
                            r="12"
                            fill="none"
                            stroke="#0867F5"
                            strokeWidth="5"
                            strokeDasharray="45 30"
                            strokeDashoffset="10"
                          />
                          <circle
                            cx="18"
                            cy="18"
                            r="12"
                            fill="none"
                            stroke="#50AE5D"
                            strokeWidth="5"
                            strokeDasharray="28 47"
                            strokeDashoffset="-35"
                          />
                        </svg>
                      </div>
                    </div>

                    {/* Charts Row 2: Area Wave + Mini Table */}
                    <div className="grid grid-cols-12 gap-2">
                      <div className="col-span-6 bg-blue-50/50 rounded-md p-1.5 border border-blue-100/60">
                        <svg viewBox="0 0 100 32" className="w-full h-7">
                          <path
                            d="M0,28 L16,20 L34,23 L52,12 L70,15 L88,5 L100,8 L100,32 L0,32 Z"
                            fill="#0867F5"
                            fillOpacity="0.22"
                          />
                          <path
                            d="M0,28 L16,20 L34,23 L52,12 L70,15 L88,5 L100,8"
                            fill="none"
                            stroke="#0867F5"
                            strokeWidth="2"
                          />
                        </svg>
                      </div>
                      <div className="col-span-6 bg-slate-50 rounded-md p-1.5 border border-slate-100 space-y-1">
                        <div className="h-1.5 w-full bg-slate-200 rounded" />
                        <div className="h-1.5 w-4/5 bg-blue-200 rounded" />
                        <div className="h-1.5 w-3/4 bg-emerald-200 rounded" />
                      </div>
                    </div>
                  </div>
                </div>
                {/* Laptop Base */}
                <div className="h-3 bg-gradient-to-r from-slate-400 via-slate-200 to-slate-400 rounded-b-xl mx-[-10px] shadow-md border-t border-white" />
              </div>
            </div>

            {/* THREE INFORMATION CARDS BELOW THE LAPTOP */}
            <div className="space-y-3.5">
              {/* Card 1: Deterministic & Accurate */}
              <div className="rounded-2xl bg-[#ECFDF3]/90 border border-emerald-200/80 p-4 sm:p-5 flex items-start gap-4 shadow-2xs">
                <div className="w-11 h-11 rounded-full bg-emerald-100 text-[#16A34A] flex items-center justify-center shrink-0">
                  <Zap className="w-5 h-5 fill-[#16A34A]" />
                </div>
                <div>
                  <h3 className="text-sm sm:text-base font-bold text-[#0B2540]">
                    Deterministic &amp; Accurate
                  </h3>
                  <p className="text-xs text-slate-600 mt-0.5 leading-relaxed">
                    Paise-exact calculations from your actual files. No manual entries, no estimated
                    data.
                  </p>
                </div>
              </div>

              {/* Card 2: Your Data Stays Local */}
              <div className="rounded-2xl bg-[#EFF6FF]/95 border border-blue-200/80 p-4 sm:p-5 flex items-start gap-4 shadow-2xs">
                <div className="w-11 h-11 rounded-full bg-blue-100 text-[#0867F5] flex items-center justify-center shrink-0">
                  <ShieldCheck className="w-5 h-5" />
                </div>
                <div>
                  <h3 className="text-sm sm:text-base font-bold text-[#0B2540]">
                    Your Data Stays Local
                  </h3>
                  <p className="text-xs text-slate-600 mt-0.5 leading-relaxed">
                    Files are processed in your browser. No upload to any server.
                  </p>
                </div>
              </div>

              {/* Card 3: Comprehensive Analysis */}
              <div className="rounded-2xl bg-[#F5F3FF]/95 border border-purple-200/80 p-4 sm:p-5 flex items-start gap-4 shadow-2xs">
                <div className="w-11 h-11 rounded-full bg-purple-100 text-purple-600 flex items-center justify-center shrink-0">
                  <BarChart3 className="w-5 h-5" />
                </div>
                <div>
                  <h3 className="text-sm sm:text-base font-bold text-[#0B2540]">
                    Comprehensive Analysis
                  </h3>
                  <p className="text-xs text-slate-600 mt-0.5 leading-relaxed">
                    Financial statements, reconciliations, ratios, MIS and automated insights.
                  </p>
                </div>
              </div>
            </div>
          </div>
        </div>

        {/* BOTTOM 8-MODULE FEATURE HIGHLIGHT STRIP (Matching Reference Image 2) */}
        <div className="mt-10 pt-4 grid grid-cols-2 sm:grid-cols-4 lg:grid-cols-8 gap-3">
          {BOTTOM_CAPABILITY_CARDS.map((card) => {
            const Icon = card.icon;
            return (
              <div
                key={card.title}
                className="bg-white/95 rounded-xl border border-slate-200/80 p-3.5 flex flex-col items-center text-center shadow-2xs hover:shadow-md hover:border-blue-200 transition-all"
              >
                <div
                  className={`w-9 h-9 rounded-lg ${card.iconBg} flex items-center justify-center mb-2`}
                >
                  <Icon className="w-5 h-5" />
                </div>
                <span className="text-[11px] font-bold text-[#0B2540] leading-tight">
                  {card.title}
                </span>
              </div>
            );
          })}
        </div>
      </div>

      {/* FILE PREVIEW MODAL */}
      {previewPayload && (
        <div className="fixed inset-0 z-50 flex items-center justify-center bg-slate-900/50 p-4">
          <div className="bg-white rounded-xl max-w-4xl w-full max-h-[85vh] flex flex-col shadow-xl overflow-hidden">
            <div className="px-6 py-4 bg-[#0B2540] text-white flex items-center justify-between">
              <div>
                <h3 className="text-base font-bold">
                  Parsed File Preview — {previewPayload.sourceFile.filename}
                </h3>
                <p className="text-xs text-slate-300">
                  {REPORT_TYPE_LABELS[previewPayload.sourceFile.detectedReportType]} ·{' '}
                  {previewPayload.ledgers.length} Ledgers · {previewPayload.vouchers.length} Vouchers
                  · {previewPayload.gstInvoices.length} GST Records
                </p>
              </div>
              <button
                onClick={() => setPreviewFileId(null)}
                className="p-1 rounded bg-white/10 hover:bg-white/20"
              >
                <X className="w-5 h-5" />
              </button>
            </div>
            <div className="p-6 overflow-y-auto space-y-5 text-xs">
              {previewPayload.sourceFile.validationMessages.length > 0 && (
                <div className="p-3 rounded-lg bg-amber-50 border border-amber-200 space-y-1">
                  <div className="font-bold text-amber-900">Import Validation Notes:</div>
                  {previewPayload.sourceFile.validationMessages.map((m, i) => (
                    <div key={i} className="text-amber-800">
                      [{m.severity}] {m.code}: {m.message} {m.details ? `(${m.details})` : ''}
                    </div>
                  ))}
                </div>
              )}

              {previewPayload.ledgers.length > 0 && (
                <div>
                  <h4 className="font-bold text-[#0B2540] mb-2">
                    Extracted Ledgers (Showing up to 15 of {previewPayload.ledgers.length})
                  </h4>
                  <table className="w-full border border-slate-200 text-left">
                    <thead className="bg-slate-100">
                      <tr>
                        <th className="p-2">Ledger Name</th>
                        <th className="p-2">Parent Group</th>
                        <th className="p-2">Mapped Category</th>
                        <th className="p-2 text-right">Opening Balance</th>
                      </tr>
                    </thead>
                    <tbody className="divide-y divide-slate-200">
                      {previewPayload.ledgers.slice(0, 15).map((l) => (
                        <tr key={l.id}>
                          <td className="p-2 font-medium">{l.name}</td>
                          <td className="p-2">{l.parentGroup}</td>
                          <td className="p-2 font-mono">{l.classification}</td>
                          <td className="p-2 text-right font-mono">
                            {formatINR(l.openingBalancePaise, false, true)}
                          </td>
                        </tr>
                      ))}
                    </tbody>
                  </table>
                </div>
              )}

              {previewPayload.vouchers.length > 0 && (
                <div>
                  <h4 className="font-bold text-[#0B2540] mb-2">
                    Extracted Vouchers (Showing up to 15 of {previewPayload.vouchers.length})
                  </h4>
                  <table className="w-full border border-slate-200 text-left">
                    <thead className="bg-slate-100">
                      <tr>
                        <th className="p-2">Date</th>
                        <th className="p-2">Voucher No</th>
                        <th className="p-2">Type</th>
                        <th className="p-2">Party</th>
                        <th className="p-2 text-right">Total Debit</th>
                        <th className="p-2 text-right">Total Credit</th>
                      </tr>
                    </thead>
                    <tbody className="divide-y divide-slate-200">
                      {previewPayload.vouchers.slice(0, 15).map((v) => (
                        <tr key={v.id}>
                          <td className="p-2 font-mono">{v.date}</td>
                          <td className="p-2 font-mono">{v.voucherNumber}</td>
                          <td className="p-2">{v.voucherType}</td>
                          <td className="p-2">{v.partyName || '—'}</td>
                          <td className="p-2 text-right font-mono">
                            {formatINR(v.totalDebitPaise)}
                          </td>
                          <td className="p-2 text-right font-mono">
                            {formatINR(v.totalCreditPaise)}
                          </td>
                        </tr>
                      ))}
                    </tbody>
                  </table>
                </div>
              )}

              {previewPayload.gstInvoices.length > 0 && (
                <div>
                  <h4 className="font-bold text-[#0B2540] mb-2">
                    Extracted GST Invoices ({previewPayload.gstInvoices.length})
                  </h4>
                  <table className="w-full border border-slate-200 text-left">
                    <thead className="bg-slate-100">
                      <tr>
                        <th className="p-2">Source</th>
                        <th className="p-2">GSTIN</th>
                        <th className="p-2">Invoice No</th>
                        <th className="p-2">Date</th>
                        <th className="p-2 text-right">Taxable Value</th>
                        <th className="p-2 text-right">Total Tax</th>
                      </tr>
                    </thead>
                    <tbody className="divide-y divide-slate-200">
                      {previewPayload.gstInvoices.slice(0, 15).map((g) => (
                        <tr key={g.id}>
                          <td className="p-2 font-mono">{g.source}</td>
                          <td className="p-2 font-mono">{g.partyGstin}</td>
                          <td className="p-2 font-mono">{g.invoiceNumber}</td>
                          <td className="p-2 font-mono">{g.invoiceDate}</td>
                          <td className="p-2 text-right font-mono">
                            {formatINR(g.taxableValuePaise)}
                          </td>
                          <td className="p-2 text-right font-mono">
                            {formatINR(g.totalTaxPaise)}
                          </td>
                        </tr>
                      ))}
                    </tbody>
                  </table>
                </div>
              )}
            </div>
          </div>
        </div>
      )}

      {/* CUSTOM EXCEL COLUMN MAPPING MODAL */}
      {mappingFile && (
        <div className="fixed inset-0 z-50 flex items-center justify-center bg-slate-900/50 p-4">
          <div className="bg-white rounded-xl max-w-2xl w-full p-6 space-y-4 shadow-xl">
            <div className="flex items-center justify-between border-b border-slate-200 pb-3">
              <div>
                <h3 className="text-base font-bold text-[#0B2540]">
                  Custom Column Mapping — {mappingFile.filename}
                </h3>
                <p className="text-xs text-slate-500">
                  Map spreadsheet column headers to accounting fields and re-parse deterministically.
                </p>
              </div>
              <button onClick={() => setMappingFile(null)}>
                <X className="w-5 h-5 text-slate-500" />
              </button>
            </div>

            <div className="grid grid-cols-1 sm:grid-cols-2 gap-3 text-xs">
              <div>
                <label className="block font-semibold text-slate-700 mb-1">
                  Report Type Override
                </label>
                <select
                  value={mappingDraft.reportCategoryOverride || ''}
                  onChange={(e) =>
                    setMappingDraft({
                      ...mappingDraft,
                      reportCategoryOverride: (e.target.value || undefined) as DetectedReportType,
                    })
                  }
                  className="w-full p-2 border border-slate-300 rounded"
                >
                  <option value="">Auto-Detect</option>
                  <option value="TRIAL_BALANCE_EXCEL">Trial Balance</option>
                  <option value="PROFIT_LOSS_EXCEL">Profit &amp; Loss</option>
                  <option value="BALANCE_SHEET_EXCEL">Balance Sheet</option>
                  <option value="SALES_REGISTER_EXCEL">Sales Register</option>
                  <option value="PURCHASE_REGISTER_EXCEL">Purchase Register</option>
                  <option value="OUTSTANDING_BILLS_EXCEL">Bill-wise Outstanding</option>
                  <option value="GSTR1_RETURN">GSTR-1 Return</option>
                  <option value="GSTR2B_RETURN">GSTR-2B Return</option>
                </select>
              </div>

              {[
                { key: 'ledgerNameCol', label: 'Ledger / Particulars Column' },
                { key: 'parentGroupCol', label: 'Parent Group Column' },
                { key: 'dateCol', label: 'Voucher / Invoice Date Column' },
                { key: 'voucherNoCol', label: 'Voucher / Invoice No Column' },
                { key: 'partyNameCol', label: 'Customer / Supplier Name Column' },
                { key: 'gstinCol', label: 'GSTIN Column' },
                { key: 'debitCol', label: 'Debit Amount Column' },
                { key: 'creditCol', label: 'Credit Amount Column' },
                { key: 'closingBalanceCol', label: 'Closing / Pending Amount Column' },
                { key: 'taxableValueCol', label: 'GST Taxable Value Column' },
                { key: 'igstCol', label: 'IGST Column' },
                { key: 'cgstCol', label: 'CGST Column' },
                { key: 'sgstCol', label: 'SGST Column' },
              ].map((field) => (
                <div key={field.key}>
                  <label className="block font-semibold text-slate-700 mb-1">{field.label}</label>
                  <select
                    value={(mappingDraft as any)[field.key] || ''}
                    onChange={(e) =>
                      setMappingDraft({
                        ...mappingDraft,
                        [field.key]: e.target.value || undefined,
                      })
                    }
                    className="w-full p-2 border border-slate-300 rounded"
                  >
                    <option value="">— None —</option>
                    {(mappingFile.rawHeaders || []).map((h) => (
                      <option key={h} value={h}>
                        {h}
                      </option>
                    ))}
                  </select>
                </div>
              ))}
            </div>

            <div className="flex justify-end gap-2 pt-3 border-t border-slate-200">
              <button
                onClick={() => setMappingFile(null)}
                className="px-4 py-2 text-xs font-semibold text-slate-600 bg-slate-100 rounded-lg"
              >
                Cancel
              </button>
              <button
                onClick={() => {
                  onApplyColumnMapping(mappingFile.id, mappingDraft);
                  setMappingFile(null);
                }}
                className="px-4 py-2 text-xs font-semibold text-white bg-[#0867F5] rounded-lg"
              >
                Apply Mapping &amp; Re-Parse
              </button>
            </div>
          </div>
        </div>
      )}

      {/* 26 AUTOMATED ACCOUNTING TESTS MODAL */}
      {testResults && (
        <div className="fixed inset-0 z-50 flex items-center justify-center bg-slate-900/50 p-4">
          <div className="bg-white rounded-xl max-w-4xl w-full max-h-[85vh] flex flex-col shadow-xl overflow-hidden">
            <div className="px-6 py-4 bg-[#0B2540] text-white flex items-center justify-between">
              <div>
                <h3 className="text-base font-bold">
                  Automated Accounting &amp; Reconciliation Test Suite (26 Deterministic Checks)
                </h3>
                <p className="text-xs text-slate-300">
                  Passed {testResults.filter((t) => t.passed).length} of {testResults.length}{' '}
                  deterministic verification tests
                </p>
              </div>
              <button
                onClick={() => setTestResults(null)}
                className="p-1 rounded bg-white/10 hover:bg-white/20"
              >
                <X className="w-5 h-5" />
              </button>
            </div>
            <div className="p-6 overflow-y-auto">
              <table className="w-full text-left border-collapse text-xs">
                <thead>
                  <tr className="bg-slate-100 border-b border-slate-200">
                    <th className="py-2 px-3">#</th>
                    <th className="py-2 px-3">Test Case</th>
                    <th className="py-2 px-3">Status</th>
                    <th className="py-2 px-3">Verification Detail</th>
                  </tr>
                </thead>
                <tbody className="divide-y divide-slate-200">
                  {testResults.map((t) => (
                    <tr key={t.id}>
                      <td className="py-2 px-3 font-mono">{t.id}</td>
                      <td className="py-2 px-3 font-semibold text-[#0B2540]">{t.name}</td>
                      <td className="py-2 px-3">
                        {t.passed ? (
                          <span className="text-emerald-700 font-bold">PASS</span>
                        ) : (
                          <span className="text-red-600 font-bold">FAIL</span>
                        )}
                      </td>
                      <td className="py-2 px-3 font-mono text-[11px] text-slate-600">
                        {t.assertionDetail}
                      </td>
                    </tr>
                  ))}
                </tbody>
              </table>
            </div>
          </div>
        </div>
      )}

      {/* REVIEW MAPPINGS & CELL PROVENANCE MODAL */}
      {showReviewMappingsModal && (
        <div className="fixed inset-0 z-50 flex items-center justify-center bg-slate-900/50 p-4">
          <div className="bg-white rounded-xl max-w-5xl w-full max-h-[85vh] flex flex-col shadow-xl overflow-hidden">
            <div className="px-6 py-4 bg-[#0B2540] text-white flex items-center justify-between">
              <div>
                <h3 className="text-base font-bold">
                  Source Cell Provenance &amp; Schedule III Mapping Review
                </h3>
                <p className="text-xs text-slate-300">
                  Company: {model.companyName} • Period: {model.currentPeriodLabel || model.periodEnd}
                  {model.comparativePeriodLabel ? ` vs ${model.comparativePeriodLabel}` : ''}
                </p>
              </div>
              <button
                onClick={() => setShowReviewMappingsModal(false)}
                className="p-1 rounded bg-white/10 hover:bg-white/20 cursor-pointer"
              >
                <X className="w-5 h-5" />
              </button>
            </div>
            <div className="p-6 overflow-y-auto space-y-5">
              <div>
                <h4 className="text-xs font-bold uppercase tracking-wider text-[#003087] mb-2">
                  Extracted Statement Line Items ({model.ledgerSummaries.size})
                </h4>
                <div className="border border-slate-200 rounded-lg overflow-x-auto">
                  <table className="w-full text-left border-collapse text-xs">
                    <thead>
                      <tr className="bg-slate-100 border-b border-slate-200 text-slate-700">
                        <th className="py-2 px-3 font-semibold">Original Row Label</th>
                        <th className="py-2 px-3 font-semibold">Statement</th>
                        <th className="py-2 px-3 font-semibold">Source Cell</th>
                        <th className="py-2 px-3 font-semibold">Prior Cell</th>
                        <th className="py-2 px-3 font-semibold">Mapped Classification</th>
                        <th className="py-2 px-3 font-semibold text-right">
                          {model.currentPeriodLabel || 'FY 2025-26'}
                        </th>
                        <th className="py-2 px-3 font-semibold text-right">
                          {model.comparativePeriodLabel || 'FY 2024-25'}
                        </th>
                      </tr>
                    </thead>
                    <tbody className="divide-y divide-slate-200">
                      {Array.from(model.ledgerSummaries.values()).map((ls) => {
                        const curAmt =
                          ls.periodCreditPaise > 0
                            ? ls.periodCreditPaise
                            : ls.periodDebitPaise > 0
                              ? ls.periodDebitPaise
                              : Math.abs(ls.effectiveClosingPaise);
                        return (
                          <tr key={ls.ledger.normalizedName} className="hover:bg-slate-50">
                            <td className="py-2 px-3 font-medium text-[#0B2540]">
                              {ls.ledger.name}
                            </td>
                            <td className="py-2 px-3 font-mono text-[11px] text-slate-600">
                              {ls.ledger.statementSourceType || ls.ledger.parentGroup}
                            </td>
                            <td className="py-2 px-3 font-mono text-[11px] text-[#0867F5] font-semibold">
                              {ls.ledger.cellReference || '—'}
                            </td>
                            <td className="py-2 px-3 font-mono text-[11px] text-slate-500">
                              {ls.ledger.priorCellReference || '—'}
                            </td>
                            <td className="py-2 px-3 font-mono text-[11px] text-slate-700">
                              {ls.ledger.classification}
                            </td>
                            <td className="py-2 px-3 text-right font-mono font-semibold">
                              {formatINR(curAmt)}
                            </td>
                            <td className="py-2 px-3 text-right font-mono text-slate-600">
                              {ls.ledger.previousPeriodPaise !== undefined
                                ? formatINR(ls.ledger.previousPeriodPaise)
                                : '—'}
                            </td>
                          </tr>
                        );
                      })}
                    </tbody>
                  </table>
                </div>
              </div>
            </div>
          </div>
        </div>
      )}
    </div>
  );
};
