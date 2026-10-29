import React, { useEffect, useMemo, useRef, useState } from 'react';
import {
  BrowserRouter,
  Navigate,
  Route,
  Routes,
  useNavigate,
} from 'react-router-dom';
import {
  Banknote,
  BarChart3,
  Building2,
  Calculator,
  Calendar,
  CheckCircle2,
  ClipboardCheck,
  Clock,
  Database,
  Download,
  FileSpreadsheet,
  FileText,
  FolderPlus,
  GitCompare,
  Home,
  LayoutDashboard,
  Lightbulb,
  Menu,
  RefreshCw,
  Scale,
  ShoppingCart,
  Sliders,
  TrendingUp,
  Users,
  X,
} from 'lucide-react';
import {
  AccountClassification,
  AuditReviewState,
  ColumnMappingConfig,
  FileFormat,
  PolicySettings,
} from './types/accounting';
import { computeFinancialStatements } from './engine/accountingEngine';
import { computeAgeingAnalysis } from './engine/ageingEngine';
import { runAccountingHealthCheck } from './engine/auditEngine';
import { computeCfoAnalyticsBundle } from './engine/cfoAnalyticsEngine';
import { parseExcelOrCsvBuffer } from './engine/excelParser';
import {
  ExcelExportOptions,
  exportCfoExcelPack,
  exportMISReportPDF,
} from './engine/exportEngine';
import { runGSTReconciliation } from './engine/gstMatchingEngine';
import { parseGSTJson } from './engine/gstParser';
import { generateFinancialInsights } from './engine/insightEngine';
import { computeFinancialRatios } from './engine/ratioEngine';
import {
  buildConsolidatedAccountingModel,
  RawFilePayload,
} from './engine/reconciliationEngine';
import { buildDemoPayloads, DEMO_COMPANY_NAME } from './engine/sampleDataGenerator';
import { parseTallyXML } from './engine/xmlParser';
import {
  ExecutiveOverviewSection,
  SalesAndPurchaseAnalyticsSection,
  StatementsAndComparativeSection,
} from './components/AnalyticsModulesPart1';
import {
  AccountingHealthCheckSection,
  AgeingAnalysisSection,
  FinancialRatiosSection,
} from './components/AnalyticsModulesPart2';
import {
  FinancialInsightsSection,
  GSTReconciliationSection,
  MISReportSection,
} from './components/AnalyticsModulesPart3';
import {
  CashFlowDashboardSection,
  CfoExecutiveDashboardSection,
  CfoSuggestionsAndScenariosSection,
  WorkingCapitalDashboardSection,
} from './components/CfoDashboardModules';
import { DrillDownDrawer, DrillDownTarget } from './components/DrillDownDrawer';
import { ExcelDownloadModal } from './components/ExcelDownloadModal';
import { GspuBrandLogo } from './components/GspuBrandLogo';
import { LandingPage } from './components/LandingPage';
import { UploadAndDataModule } from './components/UploadAndDataModule';

export type NavTabId =
  | 'OVERVIEW'
  | 'PNL'
  | 'BALANCE_SHEET'
  | 'CASH_FLOW'
  | 'WORKING_CAPITAL'
  | 'COMPARATIVE'
  | 'RATIOS'
  | 'CFO_SUGGESTIONS'
  | 'AGEING'
  | 'SALES'
  | 'PURCHASE'
  | 'HEALTH_CHECK'
  | 'GST_RECON'
  | 'MIS_REPORT'
  | 'UPLOADED_DATA';

const SIDEBAR_ITEMS: Array<{
  id: NavTabId;
  label: string;
  icon: React.FC<{ className?: string }>;
}> = [
  { id: 'OVERVIEW', label: 'Executive CFO Dashboard', icon: LayoutDashboard },
  { id: 'PNL', label: 'Profit & Loss Statement', icon: BarChart3 },
  { id: 'BALANCE_SHEET', label: 'Balance Sheet', icon: Scale },
  { id: 'CASH_FLOW', label: 'Cash Flow Statement', icon: Banknote },
  { id: 'WORKING_CAPITAL', label: 'Working Capital & CCC', icon: Clock },
  { id: 'COMPARATIVE', label: 'Comparative Financials', icon: GitCompare },
  { id: 'RATIOS', label: 'Financial Ratios (5 Cats)', icon: Calculator },
  { id: 'CFO_SUGGESTIONS', label: 'CFO Insights & Scenarios', icon: Sliders },
  { id: 'AGEING', label: 'Receivables & Payables', icon: Users },
  { id: 'SALES', label: 'Sales Analytics', icon: TrendingUp },
  { id: 'PURCHASE', label: 'Purchase Analytics', icon: ShoppingCart },
  { id: 'HEALTH_CHECK', label: 'Accounting Health Check', icon: ClipboardCheck },
  { id: 'GST_RECON', label: 'GST Reconciliation', icon: FileSpreadsheet },
  { id: 'MIS_REPORT', label: 'MIS Report & Pack', icon: FileText },
  { id: 'UPLOADED_DATA', label: 'Uploaded Data & Mapping', icon: Database },
];

const STORAGE_KEY_COMPANY = 'gspu_v2_company_name';
const STORAGE_KEY_PAYLOADS = 'gspu_v2_payloads';

function loadStoredPayloads(): RawFilePayload[] {
  try {
    const raw = sessionStorage.getItem(STORAGE_KEY_PAYLOADS);
    if (raw) {
      const parsed = JSON.parse(raw);
      if (Array.isArray(parsed)) return parsed;
    }
  } catch {
    // Ignore storage read errors
  }
  return [];
}

function loadStoredCompanyName(): string {
  try {
    return sessionStorage.getItem(STORAGE_KEY_COMPANY) || '';
  } catch {
    return '';
  }
}

function GspuWorkspaceRoutes() {
  const navigate = useNavigate();

  const [companyName, setCompanyName] = useState<string>(() => loadStoredCompanyName());
  const [stagedPayloads, setStagedPayloads] = useState<RawFilePayload[]>(() =>
    loadStoredPayloads()
  );
  const [committedPayloads, setCommittedPayloads] = useState<RawFilePayload[]>(() =>
    loadStoredPayloads()
  );
  const [hasUnappliedUploads, setHasUnappliedUploads] = useState<boolean>(false);

  // Store raw ArrayBuffers for Excel/CSV files so custom column mappings can re-parse in-memory
  const rawBuffersRef = useRef<Map<string, ArrayBuffer>>(new Map());

  const [activeTab, setActiveTab] = useState<NavTabId>('OVERVIEW');
  const [mobileMenuOpen, setMobileMenuOpen] = useState<boolean>(false);

  // Header Financial Year & Date-Range Filter State
  const [selectedFy, setSelectedFy] = useState<string>('ALL');
  const [filterStartMonth, setFilterStartMonth] = useState<string>('ALL');
  const [filterEndMonth, setFilterEndMonth] = useState<string>('ALL');

  const [policy, setPolicy] = useState<PolicySettings>({
    roundingTolerancePaise: 100, // ₹1.00
    materialityThresholdPaise: 10000000, // ₹1,00,000
    roundSumThresholdPaise: 5000000, // ₹50,000
    gstTolerancePaise: 200, // ₹2.00
    ageingMode: 'DUE_DATE',
    defaultCreditDays: 30,
    fyStartMonth: 4,
  });

  const [manualOverrides, setManualOverrides] = useState<
    Record<string, AccountClassification>
  >({});
  const [reviewStates, setReviewStates] = useState<Record<string, AuditReviewState>>({});
  const [drillDownTarget, setDrillDownTarget] = useState<DrillDownTarget | null>(null);
  const [isExcelModalOpen, setIsExcelModalOpen] = useState<boolean>(false);

  // Persist companyName & payloads in sessionStorage so browser refresh preserves state
  useEffect(() => {
    try {
      sessionStorage.setItem(STORAGE_KEY_COMPANY, companyName);
    } catch {
      // ignore quota errors
    }
  }, [companyName]);

  useEffect(() => {
    try {
      sessionStorage.setItem(STORAGE_KEY_PAYLOADS, JSON.stringify(stagedPayloads));
    } catch {
      // ignore quota errors for huge files
    }
  }, [stagedPayloads]);

  const isDemoActive = useMemo(
    () => stagedPayloads.length > 0 && stagedPayloads.every((p) => p.sourceFile.isDemoData),
    [stagedPayloads]
  );

  // Parse a single uploaded File into a RawFilePayload
  const parseSingleUploadedFile = async (
    file: File,
    existingId?: string,
    customMapping?: ColumnMappingConfig
  ): Promise<RawFilePayload> => {
    const fileId =
      existingId || `FILE_${Date.now()}_${Math.random().toString(36).slice(2, 7)}`;
    const ext = file.name.split('.').pop()?.toUpperCase() || '';
    const format: FileFormat =
      ext === 'XML'
        ? 'XML'
        : ext === 'XLSX'
        ? 'XLSX'
        : ext === 'XLS'
        ? 'XLS'
        : ext === 'CSV'
        ? 'CSV'
        : 'JSON';

    const now = new Date().toISOString();

    if (format === 'XML') {
      const text = await file.text();
      const parsed = parseTallyXML(text, fileId, file.name);
      const count =
        parsed.ledgers.length +
        parsed.vouchers.length +
        parsed.outstandingInvoices.length +
        parsed.stockItems.length;
      const hasErrors = parsed.validationMessages.some((m) => m.severity === 'ERROR');
      const hasWarnings = parsed.validationMessages.some((m) => m.severity === 'WARNING');

      return {
        sourceFile: {
          id: fileId,
          filename: file.name,
          format,
          sizeBytes: file.size,
          uploadedAt: now,
          detectedReportType: parsed.detectedReportType,
          extractedCompanyName: parsed.extractedCompanyName,
          periodStart: parsed.periodStart,
          periodEnd: parsed.periodEnd,
          recordsExtracted: count,
          status: hasErrors ? 'ERROR' : hasWarnings ? 'WARNING' : 'VALID',
          validationMessages: parsed.validationMessages,
        },
        groups: parsed.groups,
        ledgers: parsed.ledgers,
        vouchers: parsed.vouchers,
        outstandingInvoices: parsed.outstandingInvoices,
        stockItems: parsed.stockItems,
        gstInvoices: parsed.gstInvoices,
        gstr3bSummaries: [],
      };
    }

    if (format === 'JSON') {
      const text = await file.text();
      const parsed = parseGSTJson(text, fileId, file.name);
      const count = parsed.gstInvoices.length + parsed.gstr3bSummaries.length;
      const hasErrors = parsed.validationMessages.some((m) => m.severity === 'ERROR');

      return {
        sourceFile: {
          id: fileId,
          filename: file.name,
          format,
          sizeBytes: file.size,
          uploadedAt: now,
          detectedReportType: parsed.detectedReportType,
          periodStart: parsed.returnPeriod ? `${parsed.returnPeriod}-01` : undefined,
          periodEnd: parsed.returnPeriod ? `${parsed.returnPeriod}-28` : undefined,
          recordsExtracted: count,
          status: hasErrors ? 'ERROR' : 'VALID',
          validationMessages: parsed.validationMessages,
        },
        groups: [],
        ledgers: [],
        vouchers: [],
        outstandingInvoices: [],
        stockItems: [],
        gstInvoices: parsed.gstInvoices,
        gstr3bSummaries: parsed.gstr3bSummaries,
      };
    }

    // Excel / CSV
    const buffer = await file.arrayBuffer();
    rawBuffersRef.current.set(fileId, buffer);
    const parsed = parseExcelOrCsvBuffer(buffer, fileId, file.name, customMapping);
    const count =
      parsed.ledgers.length +
      parsed.vouchers.length +
      parsed.outstandingInvoices.length +
      parsed.gstInvoices.length +
      parsed.gstr3bSummaries.length +
      parsed.cashFlowItems.length;
    const hasErrors = parsed.validationMessages.some((m) => m.severity === 'ERROR');

    return {
      sourceFile: {
        id: fileId,
        filename: file.name,
        format,
        sizeBytes: file.size,
        uploadedAt: now,
        detectedReportType: parsed.detectedReportType,
        extractedCompanyName: parsed.extractedCompanyName,
        periodStart: parsed.periodStart,
        periodEnd: parsed.periodEnd,
        recordsExtracted: count,
        status: hasErrors
          ? 'ERROR'
          : parsed.needsMapping
          ? 'NEEDS_MAPPING'
          : parsed.validationMessages.length > 0
          ? 'WARNING'
          : 'VALID',
        validationMessages: parsed.validationMessages,
        rawHeaders: parsed.rawHeaders,
        rawSampleRows: parsed.rawSampleRows,
        columnMapping: customMapping,
      },
      groups: [],
      ledgers: parsed.ledgers,
      vouchers: parsed.vouchers,
      outstandingInvoices: parsed.outstandingInvoices,
      stockItems: [],
      gstInvoices: parsed.gstInvoices,
      gstr3bSummaries: parsed.gstr3bSummaries,
      cashFlowItems: parsed.cashFlowItems,
      reportedSubtotals: parsed.reportedSubtotals,
    };
  };

  const handleUploadFiles = async (fileList: FileList | File[]) => {
    const files = Array.from(fileList);
    if (files.length === 0) return;

    const newPayloads: RawFilePayload[] = [];
    for (const f of files) {
      const p = await parseSingleUploadedFile(f);
      newPayloads.push(p);
    }

    // Never mix Demo dataset or a previous company's financial statements with newly uploaded client files
    let baseStaged = stagedPayloads.filter((p) => !p.sourceFile.isDemoData);

    const newExtractedName = newPayloads.find((p) => p.sourceFile.extractedCompanyName)?.sourceFile
      .extractedCompanyName;
    const existingExtractedName = baseStaged.find((p) => p.sourceFile.extractedCompanyName)?.sourceFile
      .extractedCompanyName;

    const hasNewFullPack = newPayloads.some(
      (p) => p.sourceFile.detectedReportType === 'CFO_MONTHLY_PACK_EXCEL'
    );
    const isDifferentCompanyUpload = Boolean(
      newExtractedName &&
        existingExtractedName &&
        newExtractedName.trim().toLowerCase() !== existingExtractedName.trim().toLowerCase()
    );

    if (hasNewFullPack || isDifferentCompanyUpload) {
      // Uploading a complete CFO / Financial Statement workbook or a different company replaces prior session files
      baseStaged = [];
    } else {
      // Replace any previous upload of the same primary statement report type to prevent double-counting
      const incomingPrimaryTypes = new Set(
        newPayloads
          .map((p) => p.sourceFile.detectedReportType)
          .filter((t) =>
            [
              'BALANCE_SHEET_EXCEL',
              'PROFIT_LOSS_EXCEL',
              'CASH_FLOW_EXCEL',
              'TRIAL_BALANCE_EXCEL',
              'CFO_MONTHLY_PACK_EXCEL',
            ].includes(t)
          )
      );
      if (incomingPrimaryTypes.size > 0) {
        baseStaged = baseStaged.filter(
          (p) =>
            !incomingPrimaryTypes.has(p.sourceFile.detectedReportType) &&
            p.sourceFile.detectedReportType !== 'CFO_MONTHLY_PACK_EXCEL'
        );
      }
    }

    const nextStaged = [...baseStaged, ...newPayloads];

    // Automatically adopt extracted company name from newly uploaded financial statements
    if (newExtractedName) {
      setCompanyName(newExtractedName);
    } else if (companyName === DEMO_COMPANY_NAME) {
      setCompanyName('');
    }

    setStagedPayloads(nextStaged);
    setCommittedPayloads(nextStaged);
    setHasUnappliedUploads(true);
  };

  const handleRemoveFile = (fileId: string) => {
    rawBuffersRef.current.delete(fileId);
    const next = stagedPayloads.filter((p) => p.sourceFile.id !== fileId);
    setStagedPayloads(next);
    setCommittedPayloads(next);
  };

  const handleReplaceFile = async (fileId: string, newFile: File) => {
    const replacement = await parseSingleUploadedFile(newFile, fileId);
    const next = stagedPayloads.map((p) => (p.sourceFile.id === fileId ? replacement : p));
    setStagedPayloads(next);
    setHasUnappliedUploads(true);
  };

  const handleApplyColumnMapping = (fileId: string, mapping: ColumnMappingConfig) => {
    const buf = rawBuffersRef.current.get(fileId);
    const target = stagedPayloads.find((p) => p.sourceFile.id === fileId);
    if (!buf || !target) return;

    const reparsed = parseExcelOrCsvBuffer(buf, fileId, target.sourceFile.filename, mapping);
    const count =
      reparsed.ledgers.length +
      reparsed.vouchers.length +
      reparsed.outstandingInvoices.length +
      reparsed.gstInvoices.length +
      reparsed.gstr3bSummaries.length +
      reparsed.cashFlowItems.length;

    const updatedPayload: RawFilePayload = {
      sourceFile: {
        ...target.sourceFile,
        detectedReportType: reparsed.detectedReportType,
        recordsExtracted: count,
        status: count > 0 ? 'VALID' : 'NEEDS_MAPPING',
        validationMessages: reparsed.validationMessages,
        columnMapping: mapping,
      },
      groups: [],
      ledgers: reparsed.ledgers,
      vouchers: reparsed.vouchers,
      outstandingInvoices: reparsed.outstandingInvoices,
      stockItems: [],
      gstInvoices: reparsed.gstInvoices,
      gstr3bSummaries: reparsed.gstr3bSummaries,
      cashFlowItems: reparsed.cashFlowItems,
      reportedSubtotals: reparsed.reportedSubtotals,
    };

    const next = stagedPayloads.map((p) => (p.sourceFile.id === fileId ? updatedPayload : p));
    setStagedPayloads(next);
    setCommittedPayloads(next);
  };

  const handleClearAllData = () => {
    rawBuffersRef.current.clear();
    setStagedPayloads([]);
    setCommittedPayloads([]);
    setCompanyName('');
    setHasUnappliedUploads(false);
    try {
      sessionStorage.removeItem(STORAGE_KEY_COMPANY);
      sessionStorage.removeItem(STORAGE_KEY_PAYLOADS);
    } catch {
      // ignore
    }
  };

  const handleLoadDemoData = () => {
    const demo = buildDemoPayloads();
    setCompanyName(DEMO_COMPANY_NAME);
    setStagedPayloads(demo);
    setCommittedPayloads(demo);
    setHasUnappliedUploads(false);
  };

  const handleViewOrUpdateDashboard = () => {
    setCommittedPayloads(stagedPayloads);
    setHasUnappliedUploads(false);
    setActiveTab('OVERVIEW');
    navigate('/dashboard');
  };

  // Consolidated Accounting Model & All Module Calculations
  const consolidatedModel = useMemo(
    () =>
      buildConsolidatedAccountingModel(
        companyName,
        committedPayloads.length > 0 ? committedPayloads : stagedPayloads,
        policy,
        manualOverrides
      ),
    [companyName, committedPayloads, stagedPayloads, policy, manualOverrides]
  );

  // Filter months based on Date-Range Filter in Header
  const activeFilterMonths = useMemo(() => {
    const allMonths = consolidatedModel.availableMonths;
    if (allMonths.length === 0) return undefined;
    if (filterStartMonth === 'ALL' && filterEndMonth === 'ALL') return undefined;
    const start = filterStartMonth === 'ALL' ? allMonths[0] : filterStartMonth;
    const end = filterEndMonth === 'ALL' ? allMonths[allMonths.length - 1] : filterEndMonth;
    return allMonths.filter((m) => m >= start && m <= end);
  }, [consolidatedModel.availableMonths, filterStartMonth, filterEndMonth]);

  const financialStatements = useMemo(
    () =>
      computeFinancialStatements(
        consolidatedModel,
        activeFilterMonths,
        undefined,
        activeFilterMonths && activeFilterMonths.length > 0
          ? `${activeFilterMonths[0]} to ${activeFilterMonths[activeFilterMonths.length - 1]}`
          : `${consolidatedModel.periodStart} to ${consolidatedModel.periodEnd}`
      ),
    [consolidatedModel, activeFilterMonths]
  );

  const priorMonthStatements = useMemo(() => {
    const m = consolidatedModel.availableMonths;
    if (m.length >= 2) {
      return computeFinancialStatements(
        consolidatedModel,
        [m[m.length - 2]],
        undefined,
        m[m.length - 2]
      );
    }
    return undefined;
  }, [consolidatedModel]);

  const ageingAnalysis = useMemo(
    () =>
      computeAgeingAnalysis(
        consolidatedModel,
        policy,
        financialStatements.kpis.netRevenuePaise,
        financialStatements.kpis.netPurchasesPaise
      ),
    [consolidatedModel, policy, financialStatements]
  );

  const accountingHealth = useMemo(
    () => runAccountingHealthCheck(consolidatedModel, policy),
    [consolidatedModel, policy]
  );

  const financialRatios = useMemo(
    () =>
      computeFinancialRatios(
        consolidatedModel,
        financialStatements,
        policy,
        priorMonthStatements
      ),
    [consolidatedModel, financialStatements, policy, priorMonthStatements]
  );

  const gstReconciliation = useMemo(
    () => runGSTReconciliation(consolidatedModel, policy),
    [consolidatedModel, policy]
  );

  const financialInsights = useMemo(
    () =>
      generateFinancialInsights(
        consolidatedModel,
        financialStatements,
        priorMonthStatements,
        ageingAnalysis,
        financialRatios,
        accountingHealth,
        gstReconciliation
      ),
    [
      consolidatedModel,
      financialStatements,
      priorMonthStatements,
      ageingAnalysis,
      financialRatios,
      accountingHealth,
      gstReconciliation,
    ]
  );

  const cfoBundle = useMemo(
    () =>
      computeCfoAnalyticsBundle(
        consolidatedModel,
        financialStatements,
        priorMonthStatements,
        financialRatios,
        policy
      ),
    [
      consolidatedModel,
      financialStatements,
      priorMonthStatements,
      financialRatios,
      policy,
    ]
  );

  const handleExportCfoExcel = () => {
    setIsExcelModalOpen(true);
  };

  const handleConfirmExcelDownload = (options: ExcelExportOptions) => {
    exportCfoExcelPack(
      consolidatedModel,
      financialStatements,
      ageingAnalysis,
      financialRatios,
      accountingHealth,
      gstReconciliation,
      financialInsights,
      cfoBundle,
      options
    );
  };

  const handleExportMisPdf = () => {
    exportMISReportPDF(
      consolidatedModel,
      financialStatements,
      ageingAnalysis,
      financialRatios,
      accountingHealth,
      gstReconciliation,
      financialInsights,
      cfoBundle
    );
  };

  // Page 3 — Interactive Dashboard Layout
  const dashboardView = (
    <div className="min-h-screen flex flex-col bg-[#F4F7FB] text-[#0B2540]">
      {/* DASHBOARD HEADER (No user profile / login avatar) */}
      <header className="sticky top-0 z-30 bg-white border-b border-slate-200 px-4 lg:px-6 py-3 flex flex-wrap items-center justify-between gap-3 no-print">
        <div className="flex items-center gap-3 min-w-0">
          <button
            type="button"
            onClick={() => setMobileMenuOpen(!mobileMenuOpen)}
            className="lg:hidden p-2 rounded-lg text-slate-600 hover:bg-slate-100"
            aria-label="Toggle Menu"
          >
            {mobileMenuOpen ? <X className="w-5 h-5" /> : <Menu className="w-5 h-5" />}
          </button>
          <GspuBrandLogo compact onClick={() => navigate('/')} />
        </div>

        {/* Center Context Strip: Entered Company Name · Financial Year Selector · Date-Range Filter */}
        <div className="flex flex-wrap items-center gap-2 sm:gap-3 text-xs text-slate-600">
          {/* Entered Company Name */}
          <div className="inline-flex items-center gap-1.5 px-3 py-1.5 rounded-lg bg-slate-50 border border-slate-200 font-bold text-[#0B2540] max-w-[240px] truncate">
            <Building2 className="w-3.5 h-3.5 text-[#0867F5] shrink-0" />
            <span className="truncate">
              {companyName.trim() || consolidatedModel.companyName || 'Unassigned Company'}
            </span>
          </div>

          {/* Financial Year Selector */}
          <div className="inline-flex items-center gap-1.5 px-2.5 py-1.5 rounded-lg bg-slate-50 border border-slate-200">
            <Calendar className="w-3.5 h-3.5 text-[#0F9B91] shrink-0" />
            <select
              aria-label="Financial Year Selector"
              value={selectedFy}
              onChange={(e) => {
                setSelectedFy(e.target.value);
                setFilterStartMonth('ALL');
                setFilterEndMonth('ALL');
              }}
              className="bg-transparent font-semibold text-[#0B2540] focus:outline-none cursor-pointer"
            >
              <option value="ALL">FY 2025–26 (Full Period)</option>
              <option value="FY2526">FY 2025–26 (Apr–Mar)</option>
              <option value="FY2425">FY 2024–25</option>
            </select>
          </div>

          {/* Date-Range Filter */}
          {consolidatedModel.availableMonths.length > 0 && (
            <div className="hidden md:inline-flex items-center gap-1.5 px-2.5 py-1.5 rounded-lg bg-slate-50 border border-slate-200 font-mono text-[11px]">
              <span className="text-slate-400">Range:</span>
              <select
                aria-label="Period Start Month"
                value={filterStartMonth}
                onChange={(e) => setFilterStartMonth(e.target.value)}
                className="bg-transparent font-semibold text-[#0B2540] focus:outline-none cursor-pointer"
              >
                <option value="ALL">{consolidatedModel.periodStart}</option>
                {consolidatedModel.availableMonths.map((m) => (
                  <option key={m} value={m}>
                    {m}
                  </option>
                ))}
              </select>
              <span>→</span>
              <select
                aria-label="Period End Month"
                value={filterEndMonth}
                onChange={(e) => setFilterEndMonth(e.target.value)}
                className="bg-transparent font-semibold text-[#0B2540] focus:outline-none cursor-pointer"
              >
                <option value="ALL">{consolidatedModel.periodEnd}</option>
                {consolidatedModel.availableMonths.map((m) => (
                  <option key={m} value={m}>
                    {m}
                  </option>
                ))}
              </select>
            </div>
          )}

          {/* Reconciled File Status Pill */}
          <span className="hidden xl:inline-flex items-center gap-1 font-medium text-emerald-700 whitespace-nowrap">
            <CheckCircle2 className="w-3.5 h-3.5" />
            {consolidatedModel.sourceFiles.length} File(s) ·{' '}
            {consolidatedModel.deduplicatedVouchers.length} Vouchers
          </span>
        </div>

        {/* Right Actions: Add More Files · Update Dashboard · DOWNLOAD CFO EXCEL REPORT · DOWNLOAD MIS PDF */}
        <div className="flex flex-wrap items-center gap-2 shrink-0">
          <button
            type="button"
            onClick={() => navigate('/upload')}
            className="px-3 py-2 text-xs font-semibold text-[#0B2540] bg-slate-100 hover:bg-slate-200 rounded-lg flex items-center gap-1.5 transition-colors whitespace-nowrap cursor-pointer"
          >
            <FolderPlus className="w-3.5 h-3.5 text-[#0867F5]" />
            <span>Add More Files</span>
          </button>

          <button
            type="button"
            onClick={handleViewOrUpdateDashboard}
            className={`px-3 py-2 text-xs font-bold rounded-lg flex items-center gap-1.5 transition-colors whitespace-nowrap cursor-pointer ${
              hasUnappliedUploads
                ? 'text-white bg-[#50AE5D] hover:bg-[#43994f] shadow-xs'
                : 'text-[#0867F5] bg-blue-50 hover:bg-blue-100 border border-blue-200'
            }`}
          >
            <RefreshCw className="w-3.5 h-3.5" />
            <span>Update Dashboard</span>
          </button>

          <button
            type="button"
            onClick={handleExportCfoExcel}
            className="px-3 py-2 text-xs font-bold text-white bg-[#50AE5D] hover:bg-[#43994f] rounded-lg flex items-center gap-1.5 transition-colors whitespace-nowrap cursor-pointer shadow-2xs"
          >
            <FileSpreadsheet className="w-3.5 h-3.5" />
            <span>DOWNLOAD CFO EXCEL REPORT</span>
          </button>

          <button
            type="button"
            onClick={handleExportMisPdf}
            className="px-3 py-2 text-xs font-bold text-white bg-[#003087] hover:bg-[#0867F5] rounded-lg flex items-center gap-1.5 transition-colors whitespace-nowrap cursor-pointer shadow-2xs"
          >
            <Download className="w-3.5 h-3.5" />
            <span>DOWNLOAD MIS PDF</span>
          </button>
        </div>
      </header>

      {/* MAIN WORKSPACE: DARK NAVY SIDEBAR + CONTENT VIEWPORT */}
      <div className="flex-1 flex min-h-0">
        {/* Dark Navy Sidebar Navigation */}
        <aside
          className={`${
            mobileMenuOpen ? 'fixed inset-y-0 left-0 z-40 w-64 block' : 'hidden'
          } lg:block lg:w-64 bg-[#0B2540] text-slate-200 shrink-0 flex flex-col justify-between no-print`}
        >
          <div className="p-3 space-y-1 overflow-y-auto">
            <div className="px-3 py-2 flex items-center justify-between text-[11px] font-mono text-slate-400">
              <span>CFO &amp; Audit Modules</span>
              <button
                type="button"
                onClick={() => navigate('/')}
                className="text-slate-400 hover:text-white flex items-center gap-1 cursor-pointer"
                title="Back to Home Page"
              >
                <Home className="w-3 h-3" />
                <span>Home</span>
              </button>
            </div>
            {SIDEBAR_ITEMS.map((item) => {
              const Icon = item.icon;
              const isActive = activeTab === item.id;
              return (
                <button
                  key={item.id}
                  type="button"
                  onClick={() => {
                    setActiveTab(item.id);
                    setMobileMenuOpen(false);
                  }}
                  className={`w-full px-3 py-2.5 rounded-lg text-xs font-medium flex items-center gap-2.5 transition-colors text-left whitespace-nowrap cursor-pointer ${
                    isActive
                      ? 'bg-[#0867F5] text-white font-semibold shadow-xs'
                      : 'text-slate-300 hover:bg-white/10 hover:text-white'
                  }`}
                >
                  <Icon className="w-4 h-4 shrink-0" />
                  <span className="truncate">{item.label}</span>
                </button>
              );
            })}
          </div>

          <div className="p-4 border-t border-white/10 text-[11px] text-slate-400 space-y-2">
            <button
              type="button"
              onClick={() => navigate('/upload')}
              className="w-full py-2 px-3 rounded-lg bg-white/10 hover:bg-white/15 text-white font-semibold flex items-center justify-center gap-1.5 transition-colors cursor-pointer"
            >
              <FolderPlus className="w-3.5 h-3.5 text-[#55C8F5]" />
              <span>Upload / Manage Files</span>
            </button>
            <div className="pt-1">
              <div className="font-semibold text-white">100% Browser CFO Engine</div>
              <div>Zero Cloud · Paise-Exact Arithmetic</div>
            </div>
          </div>
        </aside>

        {/* Main Viewport */}
        <main className="flex-1 min-w-0 p-4 lg:p-6 overflow-y-auto">
          <div className="max-w-[1400px] mx-auto space-y-6">
            {/* Empty State Banner if user navigated to /dashboard before uploading files */}
            {consolidatedModel.sourceFiles.length === 0 && (
              <div className="bg-white rounded-2xl border border-blue-200 p-6 flex flex-wrap items-center justify-between gap-4 shadow-xs">
                <div>
                  <h2 className="text-base font-bold text-[#0B2540]">
                    No Accounting Files Uploaded Yet
                  </h2>
                  <p className="text-xs text-slate-600 mt-1">
                    Upload your Balance Sheet, Profit &amp; Loss, Cash Flow Statement, TallyPrime XML, or GST files on the Upload page, or load the sample CFO dataset to explore all modules.
                  </p>
                </div>
                <div className="flex items-center gap-2.5">
                  <button
                    type="button"
                    onClick={() => navigate('/upload')}
                    className="px-4 py-2 text-xs font-bold text-white bg-[#0867F5] hover:bg-[#0052CC] rounded-lg cursor-pointer"
                  >
                    Go to Upload Page →
                  </button>
                  <button
                    type="button"
                    onClick={handleLoadDemoData}
                    className="px-4 py-2 text-xs font-semibold text-[#003087] bg-blue-50 hover:bg-blue-100 border border-blue-200 rounded-lg cursor-pointer"
                  >
                    Load Sample CFO Dataset
                  </button>
                </div>
              </div>
            )}

            {activeTab === 'UPLOADED_DATA' && (
              <UploadAndDataModule
                mode="UPLOADED_DATA_INSPECTOR"
                companyName={companyName}
                onCompanyNameChange={setCompanyName}
                payloads={stagedPayloads}
                model={consolidatedModel}
                hasUnappliedUploads={hasUnappliedUploads}
                isDemoActive={isDemoActive}
                onUploadFiles={handleUploadFiles}
                onRemoveFile={handleRemoveFile}
                onReplaceFile={handleReplaceFile}
                onApplyColumnMapping={handleApplyColumnMapping}
                onClearAllData={handleClearAllData}
                onLoadDemoData={handleLoadDemoData}
                onViewDashboard={handleViewOrUpdateDashboard}
                onOpenVoucherDrillDown={(vid) =>
                  setDrillDownTarget({
                    title: 'Voucher Drill-Down',
                    specificVoucherId: vid,
                  })
                }
                onOpenLedgerDrillDown={(lname) =>
                  setDrillDownTarget({
                    title: 'Ledger Drill-Down',
                    specificLedgerNormalizedName: lname,
                  })
                }
              />
            )}

            {activeTab === 'OVERVIEW' && (
              <div className="space-y-8">
                <CfoExecutiveDashboardSection
                  model={consolidatedModel}
                  statements={financialStatements}
                  ratios={[...financialRatios, ...cfoBundle.cashFlowRatios]}
                  cfo={cfoBundle}
                  onNavigateTab={(t) => {
                    if (t === 'BS') setActiveTab('BALANCE_SHEET');
                    else setActiveTab(t as NavTabId);
                  }}
                  onDrillDown={setDrillDownTarget}
                  onExportCfoExcel={handleExportCfoExcel}
                  onExportMisPdf={handleExportMisPdf}
                />

                <ExecutiveOverviewSection
                  model={consolidatedModel}
                  statements={financialStatements}
                  ageing={ageingAnalysis}
                  health={accountingHealth}
                  onDrillDown={setDrillDownTarget}
                  onNavigateTab={(t) => setActiveTab(t as NavTabId)}
                />
              </div>
            )}

            {activeTab === 'CASH_FLOW' && (
              <CashFlowDashboardSection
                model={consolidatedModel}
                statements={financialStatements}
                cfo={cfoBundle}
              />
            )}

            {activeTab === 'WORKING_CAPITAL' && (
              <WorkingCapitalDashboardSection
                cfo={cfoBundle}
                statements={financialStatements}
              />
            )}

            {activeTab === 'CFO_SUGGESTIONS' && (
              <div className="space-y-8">
                <CfoSuggestionsAndScenariosSection
                  cfo={cfoBundle}
                  statements={financialStatements}
                />
                <FinancialInsightsSection
                  insights={financialInsights}
                  onDrillDown={setDrillDownTarget}
                  onNavigateTab={(t) => setActiveTab(t as NavTabId)}
                />
              </div>
            )}

            {activeTab === 'SALES' && (
              <SalesAndPurchaseAnalyticsSection
                mode="SALES"
                model={consolidatedModel}
                statements={financialStatements}
                onDrillDown={setDrillDownTarget}
              />
            )}

            {activeTab === 'PURCHASE' && (
              <SalesAndPurchaseAnalyticsSection
                mode="PURCHASE"
                model={consolidatedModel}
                statements={financialStatements}
                onDrillDown={setDrillDownTarget}
              />
            )}

            {(activeTab === 'PNL' ||
              activeTab === 'BALANCE_SHEET' ||
              activeTab === 'COMPARATIVE') && (
              <StatementsAndComparativeSection
                viewMode={activeTab}
                model={consolidatedModel}
                manualOverrides={manualOverrides}
                onUpdateManualOverride={(normLedger, cls) =>
                  setManualOverrides((prev) => ({ ...prev, [normLedger]: cls }))
                }
                onDrillDown={setDrillDownTarget}
              />
            )}

            {activeTab === 'AGEING' && (
              <AgeingAnalysisSection
                model={consolidatedModel}
                ageing={ageingAnalysis}
                monthlyCollectionsTrends={financialStatements.monthlyTrends.map((m) => ({
                  month: m.month,
                  collectionsPaise: m.collectionsPaise,
                  paymentsPaise: m.paymentsPaise,
                }))}
                policy={policy}
                onUpdatePolicy={setPolicy}
                onOpenVoucherDrillDown={(vid) =>
                  setDrillDownTarget({
                    title: 'Bill Allocation Voucher Entry',
                    specificVoucherId: vid,
                  })
                }
              />
            )}

            {activeTab === 'RATIOS' && (
              <FinancialRatiosSection
                ratios={[...financialRatios, ...cfoBundle.cashFlowRatios]}
                policy={policy}
                onUpdatePolicy={setPolicy}
              />
            )}

            {activeTab === 'HEALTH_CHECK' && (
              <AccountingHealthCheckSection
                health={accountingHealth}
                policy={policy}
                reviewStates={reviewStates}
                onUpdatePolicy={setPolicy}
                onUpdateReviewState={(id, st) =>
                  setReviewStates((prev) => ({ ...prev, [id]: st }))
                }
                onOpenVoucherDrillDown={(vid) =>
                  setDrillDownTarget({
                    title: 'Audit Exception Voucher Inspection',
                    specificVoucherId: vid,
                  })
                }
              />
            )}

            {activeTab === 'GST_RECON' && (
              <GSTReconciliationSection
                gst={gstReconciliation}
                policy={policy}
                onUpdatePolicy={setPolicy}
                onUploadFiles={handleUploadFiles}
                onOpenVoucherDrillDown={(vid) =>
                  setDrillDownTarget({
                    title: 'GST Invoice Voucher Entry',
                    specificVoucherId: vid,
                  })
                }
              />
            )}

      {activeTab === 'MIS_REPORT' && (
              <MISReportSection
                model={consolidatedModel}
                statements={financialStatements}
                ageing={ageingAnalysis}
                ratios={[...financialRatios, ...cfoBundle.cashFlowRatios]}
                health={accountingHealth}
                gst={gstReconciliation}
                insights={financialInsights}
                cfo={cfoBundle}
                onOpenExcelModal={() => setIsExcelModalOpen(true)}
              />
            )}
          </div>
        </main>
      </div>

      {/* 6-LEVEL LEDGER & VOUCHER DRILL-DOWN DRAWER */}
      <DrillDownDrawer
        target={drillDownTarget}
        model={consolidatedModel}
        onClose={() => setDrillDownTarget(null)}
      />

      {/* SECTION 17: REPORT DOWNLOAD OPTIONS MODAL */}
      <ExcelDownloadModal
        isOpen={isExcelModalOpen}
        companyName={consolidatedModel.companyName}
        periodLabel={`${consolidatedModel.periodStart} to ${consolidatedModel.periodEnd}`}
        reconciliationSummary={{
          currentPeriodLabel: consolidatedModel.currentPeriodLabel || 'FY 2025–26',
          comparativePeriodLabel: consolidatedModel.comparativePeriodLabel || 'FY 2024–25',
          currentAssetsPaise: financialStatements.kpis.totalCurrentAssetsPaise,
          nonCurrentAssetsPaise: financialStatements.kpis.totalNonCurrentAssetsPaise,
          totalAssetsPaise: financialStatements.kpis.totalAssetsPaise,
          equityPaise: financialStatements.kpis.totalEquityPaise,
          nonCurrentLiabilitiesPaise: financialStatements.kpis.longTermBorrowingsPaise,
          currentLiabilitiesPaise: financialStatements.kpis.totalCurrentLiabilitiesPaise,
          totalEquityAndLiabilitiesPaise: financialStatements.kpis.totalEquityAndLiabilitiesPaise,
          differencePaise: cfoBundle.balanceSheetReconciliation.differencePaise,
          isReconciled: cfoBundle.balanceSheetReconciliation.isReconciled,
          netWorkingCapitalPaise:
            financialStatements.kpis.totalCurrentAssetsPaise -
            financialStatements.kpis.totalCurrentLiabilitiesPaise,
        }}
        onClose={() => setIsExcelModalOpen(false)}
        onConfirmDownload={handleConfirmExcelDownload}
      />
    </div>
  );

  return (
    <Routes>
      {/* PAGE 1: HOME / LANDING PAGE */}
      <Route path="/" element={<LandingPage />} />

      {/* PAGE 2: COMPANY NAME & FILE UPLOAD */}
      <Route
        path="/upload"
        element={
          <UploadAndDataModule
            mode="UPLOAD_WORKFLOW"
            companyName={companyName}
            onCompanyNameChange={setCompanyName}
            payloads={stagedPayloads}
            model={consolidatedModel}
            hasUnappliedUploads={hasUnappliedUploads}
            isDemoActive={isDemoActive}
            onUploadFiles={handleUploadFiles}
            onRemoveFile={handleRemoveFile}
            onReplaceFile={handleReplaceFile}
            onApplyColumnMapping={handleApplyColumnMapping}
            onClearAllData={handleClearAllData}
            onLoadDemoData={handleLoadDemoData}
            onViewDashboard={handleViewOrUpdateDashboard}
            onOpenVoucherDrillDown={(vid) =>
              setDrillDownTarget({
                title: 'Voucher Drill-Down',
                specificVoucherId: vid,
              })
            }
            onOpenLedgerDrillDown={(lname) =>
              setDrillDownTarget({
                title: 'Ledger Drill-Down',
                specificLedgerNormalizedName: lname,
              })
            }
          />
        }
      />

      {/* PAGE 3: INTERACTIVE FINANCIAL DASHBOARD */}
      <Route path="/dashboard" element={dashboardView} />

      {/* Fallback Redirect */}
      <Route path="*" element={<Navigate to="/" replace />} />
    </Routes>
  );
}

export default function App() {
  return (
    <BrowserRouter>
      <GspuWorkspaceRoutes />
    </BrowserRouter>
  );
}
