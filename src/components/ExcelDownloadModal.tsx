import React, { useState } from 'react';
import {
  BarChart3,
  CheckCircle2,
  FileSpreadsheet,
  Layers,
  RotateCcw,
  Sparkles,
  X,
} from 'lucide-react';
import {
  DEFAULT_EXCEL_EXPORT_OPTIONS,
  ExcelExportOptions,
} from '../engine/cfoExcelPackBuilder';

type ExcelReportType = ExcelExportOptions['reportType'];
type ExcelExportPeriod = ExcelExportOptions['periodType'];

export interface ExcelPreExportVerificationSummary {
  currentPeriodLabel: string;
  comparativePeriodLabel: string;
  currentAssetsPaise: number;
  nonCurrentAssetsPaise: number;
  totalAssetsPaise: number;
  equityPaise: number;
  nonCurrentLiabilitiesPaise: number;
  currentLiabilitiesPaise: number;
  totalEquityAndLiabilitiesPaise: number;
  differencePaise: number;
  isReconciled: boolean;
  netWorkingCapitalPaise: number;
}

interface ExcelDownloadModalProps {
  isOpen: boolean;
  companyName: string;
  periodLabel: string;
  reconciliationSummary?: ExcelPreExportVerificationSummary;
  onClose: () => void;
  onConfirmDownload: (options: ExcelExportOptions) => Promise<void> | void;
}

const REPORT_TYPES: Array<{
  id: ExcelReportType;
  title: string;
  subtitle: string;
  badge?: string;
}> = [
  {
    id: 'CFO_MONTHLY_PACK',
    title: 'CFO Monthly Pack',
    subtitle: '20-Worksheet Executive CFO Pack with 12 KPI cards & 25 native Excel charts',
    badge: 'Default Recommended',
  },
  {
    id: 'EXECUTIVE_DASHBOARD_ONLY',
    title: 'Executive Dashboard Only',
    subtitle: 'Focused Executive Dashboard, KPI Summary & chart-backing sheets',
  },
  {
    id: 'DETAILED_FINANCIAL_ANALYSIS',
    title: 'Detailed Financial Analysis',
    subtitle: 'P&L, Balance Sheet, Cash Flow, Ratios, Working Capital & CFO Insights',
  },
  {
    id: 'COMPLETE_MIS_REPORT',
    title: 'Complete MIS Excel Report',
    subtitle: 'All 20 worksheets including Ageing, Inventory, Audit & GST schedules',
  },
];

const PERIOD_OPTIONS: Array<{ id: ExcelExportPeriod; label: string }> = [
  { id: 'FINANCIAL_YEAR', label: 'Financial Year' },
  { id: 'QUARTER', label: 'Quarter' },
  { id: 'MONTH', label: 'Month' },
];

export const ExcelDownloadModal: React.FC<ExcelDownloadModalProps> = ({
  isOpen,
  companyName,
  periodLabel,
  reconciliationSummary,
  onClose,
  onConfirmDownload,
}) => {
  const [options, setOptions] = useState<ExcelExportOptions>(DEFAULT_EXCEL_EXPORT_OPTIONS);
  const [isGenerating, setIsGenerating] = useState(false);

  if (!isOpen) return null;

  const fmtInr = (paise: number) =>
    `₹${(paise / 100).toLocaleString('en-IN', {
      minimumFractionDigits: 2,
      maximumFractionDigits: 2,
    })}`;

  const handleSelectReportType = (reportType: ExcelReportType) => {
    if (reportType === 'EXECUTIVE_DASHBOARD_ONLY') {
      setOptions({
        ...options,
        reportType,
        includeCharts: true,
        includeFinancialStatements: true,
        includeRatios: false,
        includeWorkingCapital: false,
        includeCfoSuggestions: false,
        includeSupportingSchedules: false,
      });
    } else if (reportType === 'DETAILED_FINANCIAL_ANALYSIS') {
      setOptions({
        ...options,
        reportType,
        includeCharts: true,
        includeFinancialStatements: true,
        includeRatios: true,
        includeWorkingCapital: true,
        includeCfoSuggestions: true,
        includeSupportingSchedules: false,
      });
    } else {
      setOptions({
        ...DEFAULT_EXCEL_EXPORT_OPTIONS,
        reportType,
        periodType: options.periodType,
      });
    }
  };

  const handleTriggerDownload = async () => {
    setIsGenerating(true);
    try {
      await onConfirmDownload(options);
      onClose();
    } finally {
      setIsGenerating(false);
    }
  };

  return (
    <div className="fixed inset-0 z-50 flex items-center justify-center bg-slate-900/60 backdrop-blur-xs p-4">
      <div className="bg-white rounded-2xl shadow-2xl border border-slate-200 max-w-2xl w-full overflow-hidden">
        {/* Dark Navy Header matching Excel Pack #0B1F3A */}
        <div className="bg-[#0B1F3A] text-white px-6 py-4 flex items-center justify-between">
          <div className="flex items-center gap-3">
            <div className="w-10 h-10 rounded-xl bg-[#2A78D6]/25 border border-[#2A78D6]/50 flex items-center justify-center">
              <FileSpreadsheet className="w-5 h-5 text-[#55C8F5]" />
            </div>
            <div>
              <div className="flex items-center gap-2">
                <h3 className="text-base font-bold tracking-tight">
                  Download CFO Excel Report — Configuration
                </h3>
                <span className="px-2 py-0.5 rounded-md bg-[#1BAF7A]/20 text-[#1BAF7A] border border-[#1BAF7A]/40 text-[10px] font-bold uppercase">
                  OpenXML Charts Enabled
                </span>
              </div>
              <p className="text-xs text-slate-300 mt-0.5">
                {companyName} • Period: {periodLabel}
              </p>
            </div>
          </div>
          <button
            type="button"
            onClick={onClose}
            className="p-1.5 rounded-lg text-slate-300 hover:text-white hover:bg-white/10 transition cursor-pointer"
            aria-label="Close modal"
          >
            <X className="w-5 h-5" />
          </button>
        </div>

        {/* Body */}
        <div className="p-6 space-y-5 max-h-[80vh] overflow-y-auto">
          {/* 1. Report Type */}
          <div>
            <label className="block text-xs font-bold uppercase tracking-wider text-[#0B1F3A] mb-2">
              1. Report Type
            </label>
            <div className="grid grid-cols-1 sm:grid-cols-2 gap-2.5">
              {REPORT_TYPES.map((rt) => {
                const selected = options.reportType === rt.id;
                return (
                  <button
                    key={rt.id}
                    type="button"
                    onClick={() => handleSelectReportType(rt.id)}
                    className={`p-3.5 rounded-xl border text-left transition cursor-pointer ${
                      selected
                        ? 'border-[#2A78D6] bg-blue-50/70 ring-2 ring-[#2A78D6]/20'
                        : 'border-slate-200 hover:border-slate-300 bg-white'
                    }`}
                  >
                    <div className="flex items-center justify-between gap-2">
                      <span className="text-xs font-bold text-[#0B1F3A]">{rt.title}</span>
                      {rt.badge && (
                        <span className="px-1.5 py-0.5 rounded bg-[#0B1F3A] text-white text-[10px] font-semibold">
                          {rt.badge}
                        </span>
                      )}
                    </div>
                    <p className="text-[11px] text-slate-600 mt-1 leading-snug">{rt.subtitle}</p>
                  </button>
                );
              })}
            </div>
          </div>

          {/* 2. Period */}
          <div>
            <label className="block text-xs font-bold uppercase tracking-wider text-[#0B1F3A] mb-2">
              2. Reporting Period Granularity
            </label>
            <div className="grid grid-cols-3 gap-2.5">
              {PERIOD_OPTIONS.map((p) => {
                const active = options.periodType === p.id;
                return (
                  <button
                    key={p.id}
                    type="button"
                    onClick={() => setOptions({ ...options, periodType: p.id })}
                    className={`py-2.5 px-3 rounded-xl border text-xs font-bold transition cursor-pointer ${
                      active
                        ? 'bg-[#0B1F3A] text-white border-[#0B1F3A]'
                        : 'bg-slate-50 text-slate-700 border-slate-200 hover:bg-slate-100'
                    }`}
                  >
                    {p.label}
                  </button>
                );
              })}
            </div>
          </div>

          {/* 3. Include Modules */}
          <div>
            <div className="flex items-center justify-between mb-2">
              <label className="block text-xs font-bold uppercase tracking-wider text-[#0B1F3A]">
                3. Include Workbook Sections &amp; Visualisations
              </label>
              <button
                type="button"
                onClick={() => setOptions(DEFAULT_EXCEL_EXPORT_OPTIONS)}
                className="inline-flex items-center gap-1 text-[11px] font-semibold text-[#2A78D6] hover:underline cursor-pointer"
              >
                <RotateCcw className="w-3 h-3" />
                Reset to Complete CFO Monthly Pack
              </button>
            </div>

            <div className="grid grid-cols-1 sm:grid-cols-2 gap-2.5">
              {[
                {
                  key: 'includeCharts' as const,
                  label: 'Embedded Native Excel Charts (25 Charts)',
                  desc: 'Editable Column, Line, Bar & Doughnut charts linked to Dashboard Data',
                },
                {
                  key: 'includeFinancialStatements' as const,
                  label: 'Financial Statements & Comparative Tabs',
                  desc: 'P&L Dashboard, Balance Sheet Dashboard, Cash Flow & Comparative',
                },
                {
                  key: 'includeRatios' as const,
                  label: 'Financial Ratios (5 Categories)',
                  desc: 'Liquidity, Profitability, Efficiency, Solvency & Cash Flow Ratios',
                },
                {
                  key: 'includeWorkingCapital' as const,
                  label: 'Working Capital & CCC Schedule',
                  desc: 'Current Assets/Liabilities, DSO, DIO, DPO & Cash Conversion Cycle',
                },
                {
                  key: 'includeCfoSuggestions' as const,
                  label: 'CFO Suggestions & Risk Ratings',
                  desc: 'Green / Amber / Red risk-classified CFO observations & actions',
                },
                {
                  key: 'includeSupportingSchedules' as const,
                  label: 'Supporting Schedules & Reconciliations',
                  desc: 'Receivables, Payables, Inventory, Health Check, GST & Validation',
                },
              ].map((item) => {
                const checked = options[item.key];
                return (
                  <label
                    key={item.key}
                    className={`flex items-start gap-2.5 p-3 rounded-xl border cursor-pointer transition ${
                      checked
                        ? 'bg-[#F3F6FA] border-[#2A78D6]/50'
                        : 'bg-white border-slate-200 opacity-75'
                    }`}
                  >
                    <input
                      type="checkbox"
                      checked={checked}
                      onChange={(e) =>
                        setOptions({ ...options, [item.key]: e.target.checked })
                      }
                      className="mt-0.5 rounded border-slate-300 text-[#2A78D6] focus:ring-[#2A78D6]"
                    />
                    <div>
                      <div className="text-xs font-bold text-[#0B1F3A]">{item.label}</div>
                      <div className="text-[11px] text-slate-500 mt-0.5">{item.desc}</div>
                    </div>
                  </label>
                );
              })}
            </div>
          </div>

          {/* Pre-Export Balance Sheet Three-Layer Reconciliation Verification */}
          {reconciliationSummary && (
            <div
              className={`p-3.5 rounded-xl border text-xs ${
                reconciliationSummary.isReconciled
                  ? 'bg-emerald-50/70 border-emerald-200'
                  : 'bg-amber-50/80 border-amber-300'
              }`}
            >
              <div className="flex items-center justify-between gap-2 mb-2">
                <span className="font-bold uppercase tracking-wider text-[#0B1F3A]">
                  Pre-Export Schedule III Balance Sheet Control Check (
                  {reconciliationSummary.currentPeriodLabel})
                </span>
                <span
                  className={`inline-flex items-center gap-1 px-2 py-0.5 rounded-md font-bold text-[10px] uppercase ${
                    reconciliationSummary.isReconciled
                      ? 'bg-[#1BAF7A] text-white'
                      : 'bg-amber-600 text-white'
                  }`}
                >
                  <CheckCircle2 className="w-3 h-3" />
                  {reconciliationSummary.isReconciled
                    ? 'Reconciled (Diff: ₹0.00)'
                    : `Diff: ${fmtInr(reconciliationSummary.differencePaise)}`}
                </span>
              </div>
              <div className="grid grid-cols-2 sm:grid-cols-4 gap-2 text-[11px] text-slate-700">
                <div className="bg-white/90 p-2 rounded-lg border border-slate-200/80">
                  <div className="text-slate-500">Current Assets</div>
                  <div className="font-bold text-[#0B1F3A] mt-0.5">
                    {fmtInr(reconciliationSummary.currentAssetsPaise)}
                  </div>
                </div>
                <div className="bg-white/90 p-2 rounded-lg border border-slate-200/80">
                  <div className="text-slate-500">Non-Current Assets</div>
                  <div className="font-bold text-[#0B1F3A] mt-0.5">
                    {fmtInr(reconciliationSummary.nonCurrentAssetsPaise)}
                  </div>
                </div>
                <div className="bg-white/90 p-2 rounded-lg border border-slate-200/80">
                  <div className="text-slate-500">Total Assets</div>
                  <div className="font-bold text-[#0867F5] mt-0.5">
                    {fmtInr(reconciliationSummary.totalAssetsPaise)}
                  </div>
                </div>
                <div className="bg-white/90 p-2 rounded-lg border border-slate-200/80">
                  <div className="text-slate-500">Net Working Capital</div>
                  <div className="font-bold text-[#0B1F3A] mt-0.5">
                    {fmtInr(reconciliationSummary.netWorkingCapitalPaise)}
                  </div>
                </div>
                <div className="bg-white/90 p-2 rounded-lg border border-slate-200/80">
                  <div className="text-slate-500">Shareholders&apos; Equity</div>
                  <div className="font-bold text-[#0B1F3A] mt-0.5">
                    {fmtInr(reconciliationSummary.equityPaise)}
                  </div>
                </div>
                <div className="bg-white/90 p-2 rounded-lg border border-slate-200/80">
                  <div className="text-slate-500">Non-Current Liabilities</div>
                  <div className="font-bold text-[#0B1F3A] mt-0.5">
                    {fmtInr(reconciliationSummary.nonCurrentLiabilitiesPaise)}
                  </div>
                </div>
                <div className="bg-white/90 p-2 rounded-lg border border-slate-200/80">
                  <div className="text-slate-500">Current Liabilities</div>
                  <div className="font-bold text-[#0B1F3A] mt-0.5">
                    {fmtInr(reconciliationSummary.currentLiabilitiesPaise)}
                  </div>
                </div>
                <div className="bg-white/90 p-2 rounded-lg border border-slate-200/80">
                  <div className="text-slate-500">Total Equity &amp; Liab</div>
                  <div className="font-bold text-[#1BAF7A] mt-0.5">
                    {fmtInr(reconciliationSummary.totalEquityAndLiabilitiesPaise)}
                  </div>
                </div>
              </div>
            </div>
          )}

          {/* Workbook Architecture Summary Box */}
          <div className="p-3.5 rounded-xl bg-[#F3F6FA] border border-slate-200 flex flex-wrap items-center justify-between gap-3 text-xs">
            <div className="flex items-center gap-2.5">
              <BarChart3 className="w-4 h-4 text-[#2A78D6] shrink-0" />
              <span className="text-slate-700">
                <strong>Palette:</strong> Dark Navy <code className="text-[11px]">#0B1F3A</code> • Royal Blue{' '}
                <code className="text-[11px]">#2A78D6</code> • Emerald{' '}
                <code className="text-[11px]">#1BAF7A</code> • Formula Recalculation Enabled
              </span>
            </div>
            <div className="inline-flex items-center gap-1 text-[#1BAF7A] font-bold">
              <CheckCircle2 className="w-3.5 h-3.5" />
              <span>100% Actual Company Data</span>
            </div>
          </div>
        </div>

        {/* Footer */}
        <div className="px-6 py-4 bg-slate-50 border-t border-slate-200 flex items-center justify-between gap-3">
          <div className="flex items-center gap-1.5 text-xs text-slate-500">
            <Layers className="w-3.5 h-3.5 text-[#2A78D6]" />
            <span>Default: Complete CFO Monthly Pack with Charts (.xlsx)</span>
          </div>
          <div className="flex items-center gap-2.5">
            <button
              type="button"
              onClick={onClose}
              className="px-4 py-2 rounded-xl border border-slate-300 text-xs font-semibold text-slate-700 hover:bg-slate-100 cursor-pointer"
            >
              Cancel
            </button>
            <button
              type="button"
              disabled={isGenerating}
              onClick={handleTriggerDownload}
              className="px-5 py-2.5 rounded-xl bg-[#1BAF7A] hover:bg-[#159668] disabled:opacity-60 text-white text-xs font-bold flex items-center gap-2 shadow-sm transition cursor-pointer"
            >
              <Sparkles className="w-4 h-4" />
              <span>
                {isGenerating
                  ? 'Building CFO Excel Pack & Charts...'
                  : 'Download CFO Excel Pack (.xlsx)'}
              </span>
            </button>
          </div>
        </div>
      </div>
    </div>
  );
};
