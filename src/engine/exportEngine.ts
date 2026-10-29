import jsPDF from 'jspdf';
import autoTable from 'jspdf-autotable';
import {
  ConsolidatedAccountingModel,
  FinancialInsight,
  FinancialRatioResult,
  formatINR,
} from '../types/accounting';
import { FinancialStatementBundle } from './accountingEngine';
import { AgeingAnalysisResult } from './ageingEngine';
import { AccountingHealthCheckResult } from './auditEngine';
import { GSTReconciliationReport } from './gstMatchingEngine';
import { CfoAnalyticsBundle } from './cfoAnalyticsEngine';
import {
  buildCfoExcelWorkbookBytes,
  DEFAULT_EXCEL_EXPORT_OPTIONS,
  downloadCfoExcelPackWithOptions,
  ExcelExportOptions,
} from './cfoExcelPackBuilder';

export {
  buildCfoExcelWorkbookBytes,
  DEFAULT_EXCEL_EXPORT_OPTIONS,
  downloadCfoExcelPackWithOptions,
};
export type { ExcelExportOptions };

export function exportCfoExcelPack(
  model: ConsolidatedAccountingModel,
  statements: FinancialStatementBundle,
  ageing: AgeingAnalysisResult,
  ratios: FinancialRatioResult[],
  health: AccountingHealthCheckResult,
  gst: GSTReconciliationReport,
  insights: FinancialInsight[],
  cfo: CfoAnalyticsBundle,
  options: ExcelExportOptions = DEFAULT_EXCEL_EXPORT_OPTIONS
): void {
  void downloadCfoExcelPackWithOptions(
    model,
    statements,
    ageing,
    ratios,
    health,
    gst,
    insights,
    cfo,
    options
  );
}

export function exportCompleteMISExcel(
  model: ConsolidatedAccountingModel,
  statements: FinancialStatementBundle,
  ageing: AgeingAnalysisResult,
  ratios: FinancialRatioResult[],
  health: AccountingHealthCheckResult,
  gst: GSTReconciliationReport,
  insights: FinancialInsight[],
  cfo?: CfoAnalyticsBundle,
  options: ExcelExportOptions = DEFAULT_EXCEL_EXPORT_OPTIONS
): void {
  if (cfo) {
    void downloadCfoExcelPackWithOptions(
      model,
      statements,
      ageing,
      ratios,
      health,
      gst,
      insights,
      cfo,
      options
    );
  }
}

function drawHorizontalBarChart(
  doc: jsPDF,
  startX: number,
  startY: number,
  width: number,
  title: string,
  items: { label: string; valuePaise: number; color: [number, number, number] }[]
): number {
  doc.setFillColor(248, 250, 252);
  const boxHeight = 14 + items.length * 10;
  doc.roundedRect(startX, startY, width, boxHeight, 2, 2, 'F');

  doc.setFont('helvetica', 'bold');
  doc.setFontSize(9.5);
  doc.setTextColor(11, 37, 64);
  doc.text(title, startX + 4, startY + 7);

  const maxAbs = Math.max(...items.map((i) => Math.abs(i.valuePaise)), 1);
  const barAreaX = startX + 52;
  const maxBarW = width - 96;

  items.forEach((item, idx) => {
    const rowY = startY + 13 + idx * 10;
    doc.setFont('helvetica', 'normal');
    doc.setFontSize(8);
    doc.setTextColor(51, 65, 85);
    doc.text(item.label.slice(0, 26), startX + 4, rowY + 4);

    const w = Math.max(2, (Math.abs(item.valuePaise) / maxAbs) * maxBarW);
    doc.setFillColor(item.color[0], item.color[1], item.color[2]);
    doc.roundedRect(barAreaX, rowY, w, 5.5, 1, 1, 'F');

    doc.setFont('helvetica', 'bold');
    doc.setTextColor(15, 23, 42);
    doc.text(formatINR(item.valuePaise), barAreaX + w + 3, rowY + 4.2);
  });

  return startY + boxHeight + 6;
}

export function exportMISReportPDF(
  model: ConsolidatedAccountingModel,
  statements: FinancialStatementBundle,
  ageing: AgeingAnalysisResult,
  ratios: FinancialRatioResult[],
  health: AccountingHealthCheckResult,
  gst: GSTReconciliationReport,
  insights: FinancialInsight[],
  cfo?: CfoAnalyticsBundle
): void {
  const doc = new jsPDF({ orientation: 'portrait', unit: 'mm', format: 'a4' });
  const today = new Date().toISOString().slice(0, 10);

  const renderSectionBanner = (title: string) => {
    doc.setFillColor(0, 48, 135); // GSPU Deep Blue #003087
    doc.rect(0, 0, 210, 18, 'F');
    doc.setTextColor(255, 255, 255);
    doc.setFontSize(10);
    doc.setFont('helvetica', 'bold');
    doc.text('GSPU FINANCIAL ANALYTICS — CFO MANAGEMENT INFORMATION REPORT', 14, 11);
    doc.setFont('helvetica', 'normal');
    doc.setFontSize(8.5);
    doc.text(`${model.companyName.slice(0, 34)} | ${model.periodStart} to ${model.periodEnd}`, 196, 11, {
      align: 'right',
    });

    doc.setTextColor(0, 48, 135);
    doc.setFontSize(13);
    doc.setFont('helvetica', 'bold');
    doc.text(title, 14, 27);
  };

  // SECTION 1: COVER & EXECUTIVE SUMMARY + PROFITABILITY CHART
  renderSectionBanner('1. Executive CFO Summary & Profitability Bridge');

  const chartBottomY = drawHorizontalBarChart(doc, 14, 32, 182, 'Profitability & Cash Generation Overview (Actual INR)', [
    { label: 'Total Revenue', valuePaise: statements.kpis.netRevenuePaise, color: [8, 103, 245] },
    { label: 'Gross Profit', valuePaise: statements.kpis.grossProfitPaise, color: [14, 165, 233] },
    { label: 'EBITDA', valuePaise: statements.kpis.ebitdaPaise, color: [16, 185, 129] },
    { label: 'Net Profit (PAT)', valuePaise: statements.kpis.patPaise, color: [80, 174, 93] },
    {
      label: 'Operating Cash Flow',
      valuePaise: cfo?.cashFlowAnalysis.operatingCashFlowPaise ?? 0,
      color: [0, 48, 135],
    },
  ]);

  if (cfo) {
    autoTable(doc, {
      startY: chartBottomY + 2,
      head: [['CFO KPI Metric', 'Current Period', 'Previous Period', 'Variance', 'Variance %', 'Status']],
      body: cfo.cfoKpis.map((k) => [
        k.label,
        k.currentFormatted,
        k.previousFormatted,
        k.varianceAmountFormatted,
        k.variancePercentFormatted,
        k.statusIndicator,
      ]),
      headStyles: { fillColor: [0, 48, 135], fontSize: 8 },
      bodyStyles: { fontSize: 7.5 },
    });
  } else {
    autoTable(doc, {
      startY: chartBottomY + 2,
      head: [['Executive Indicator', 'Current Period Value', 'Basis / Source']],
      body: [
        ['Company / Client Name', model.companyName, `${model.sourceFiles.length} Uploaded File(s)`],
        ['Reporting Period', `${model.periodStart} to ${model.periodEnd}`, `${model.deduplicatedVouchers.length} Vouchers`],
        ['Net Revenue from Operations', formatINR(statements.kpis.netRevenuePaise), 'Sales − Sales Returns'],
        ['Gross Profit', formatINR(statements.kpis.grossProfitPaise), 'Net Revenue − COGS − Direct Expenses'],
        ['EBITDA', formatINR(statements.kpis.ebitdaPaise), 'Operating Profit before Depreciation & Interest'],
        ['Profit After Tax (Net Profit)', formatINR(statements.kpis.patPaise), 'PBT − Tax Expense'],
      ],
      headStyles: { fillColor: [0, 48, 135] },
    });
  }

  // SECTION 2: PROFIT & LOSS AND EXPENSE CHART
  doc.addPage();
  renderSectionBanner('2. Profit & Loss Analysis & Cost Structure');

  const k = statements.kpis;
  const topExp = [
    { label: 'Purchases & COGS', valuePaise: k.cogsPaise, color: [0, 48, 135] as [number, number, number] },
    { label: 'Direct Expenses', valuePaise: k.directExpensesPaise, color: [8, 103, 245] as [number, number, number] },
    { label: 'Employee Costs', valuePaise: k.employeeExpensesPaise, color: [85, 200, 245] as [number, number, number] },
    { label: 'Admin & General', valuePaise: k.adminExpensesPaise, color: [80, 174, 93] as [number, number, number] },
    { label: 'Finance & Depr.', valuePaise: k.financeCostsPaise + k.depreciationPaise, color: [245, 158, 11] as [number, number, number] },
  ].filter((x) => x.valuePaise > 0);

  const expBottomY =
    topExp.length > 0
      ? drawHorizontalBarChart(doc, 14, 32, 182, 'Major Cost & Expense Heads (Actual INR)', topExp)
      : 32;

  autoTable(doc, {
    startY: expBottomY + 2,
    head: [['Profit & Loss Line Item', 'Current Period', 'Previous Period', 'Variance %', 'Accounting Basis']],
    body: statements.pnlLines.map((l) => [
      l.label,
      formatINR(l.amountPaise),
      l.previousAmountPaise !== undefined ? formatINR(l.previousAmountPaise) : 'N/A',
      l.varianceLabel || 'N/A',
      l.formulaExplanation,
    ]),
    headStyles: { fillColor: [0, 48, 135], fontSize: 8 },
    bodyStyles: { fontSize: 7.5 },
  });

  // SECTION 3: BALANCE SHEET & CASH FLOW ANALYSIS
  doc.addPage();
  renderSectionBanner('3. Balance Sheet, Cash Flow & Working Capital');

  const totalLiabOnlyPaise =
    statements.kpis.longTermBorrowingsPaise +
    statements.kpis.shortTermBorrowingsPaise +
    statements.kpis.totalCurrentLiabilitiesPaise;

  const bsBarBottom = drawHorizontalBarChart(doc, 14, 32, 182, 'Balance Sheet & Cash Flow Composition', [
    { label: 'Total Assets', valuePaise: statements.kpis.totalAssetsPaise, color: [0, 48, 135] },
    { label: 'Total Equity', valuePaise: statements.kpis.totalEquityPaise, color: [80, 174, 93] },
    { label: 'Total Liabilities', valuePaise: totalLiabOnlyPaise, color: [245, 158, 11] },
    { label: 'Cash & Bank Balances', valuePaise: statements.kpis.cashPaise + statements.kpis.bankPaise, color: [8, 103, 245] },
    { label: 'Net Working Capital', valuePaise: statements.kpis.workingCapitalPaise, color: [14, 165, 233] },
  ]);

  autoTable(doc, {
    startY: bsBarBottom + 2,
    head: [['Balance Sheet Schedule', 'Classification', 'Closing Balance']],
    body: [
      ...statements.bsAssetLines.map((l) => [l.label, 'Assets', formatINR(l.amountPaise)]),
      ...statements.bsLiabilityEquityLines.map((l) => [l.label, 'Equity & Liabilities', formatINR(l.amountPaise)]),
    ],
    headStyles: { fillColor: [0, 48, 135], fontSize: 8 },
    bodyStyles: { fontSize: 7.5 },
  });

  if (cfo) {
    const cf = cfo.cashFlowAnalysis;
    autoTable(doc, {
      startY: (doc as any).lastAutoTable.finalY + 8,
      head: [['Cash Flow & Working Capital Metric', 'Value', 'Basis / Reconciliation']],
      body: [
        [
          'Operating Cash Flow (OCF)',
          cf.operatingCashFlowPaise !== null ? formatINR(cf.operatingCashFlowPaise) : 'N/A',
          cf.reconciliation.statusMessage,
        ],
        [
          'Investing Cash Flow (ICF)',
          cf.investingCashFlowPaise !== null ? formatINR(cf.investingCashFlowPaise) : 'N/A',
          'Capex & Investment Activities',
        ],
        [
          'Financing Cash Flow',
          cf.financingCashFlowPaise !== null ? formatINR(cf.financingCashFlowPaise) : 'N/A',
          'Debt & Equity Movements',
        ],
        [
          'Free Cash Flow (FCF)',
          cf.freeCashFlowPaise !== null ? formatINR(cf.freeCashFlowPaise) : 'N/A',
          cf.freeCashFlowDefinition,
        ],
        [
          'Cash Conversion Cycle (CCC)',
          cfo.workingCapital.cashConversionCycleDays !== null
            ? `${cfo.workingCapital.cashConversionCycleDays.toFixed(1)} Days`
            : 'N/A',
          `DSO: ${cfo.workingCapital.dsoDays?.toFixed(0) ?? 'N/A'}d | DIO: ${cfo.workingCapital.dioDays?.toFixed(0) ?? 'N/A'}d | DPO: ${cfo.workingCapital.dpoDays?.toFixed(0) ?? 'N/A'}d`,
        ],
      ],
      headStyles: { fillColor: [8, 103, 245], fontSize: 8 },
      bodyStyles: { fontSize: 7.5 },
    });
  }

  // SECTION 4: FINANCIAL RATIOS & AGEING
  doc.addPage();
  renderSectionBanner('4. Financial Ratios (5 Categories) & Ageing Summary');

  const combinedRatios = cfo ? [...ratios, ...cfo.cashFlowRatios] : ratios;
  autoTable(doc, {
    startY: 32,
    head: [['Category', 'Financial Ratio', 'Current Value', 'Prior Value', 'Formula']],
    body: combinedRatios.map((r) => [r.category, r.name, r.formattedCurrent, r.formattedPrevious, r.formula]),
    headStyles: { fillColor: [0, 48, 135], fontSize: 8 },
    bodyStyles: { fontSize: 7.5 },
  });

  autoTable(doc, {
    startY: (doc as any).lastAutoTable.finalY + 8,
    head: [['Ageing Bucket', 'Receivables (INR)', 'Bills', 'Payables (INR)', 'Bills']],
    body: ageing.bucketSummaries.map((b) => [
      b.label,
      formatINR(b.receivablesPaise),
      String(b.receivablesCount),
      formatINR(b.payablesPaise),
      String(b.payablesCount),
    ]),
    headStyles: { fillColor: [8, 103, 245], fontSize: 8 },
    bodyStyles: { fontSize: 7.5 },
  });

  // SECTION 5: CFO RECOMMENDATIONS, AUDIT & PROVENANCE
  doc.addPage();
  renderSectionBanner('5. CFO Recommendations, Audit & Data Provenance');

  if (cfo && cfo.cfoSuggestions.length > 0) {
    autoTable(doc, {
      startY: 32,
      head: [['Domain', 'Risk Level', 'Factual Calculated Observation', 'CFO Management Action Recommendation']],
      body: cfo.cfoSuggestions.map((s) => [
        s.category,
        s.riskClassification,
        s.factualObservation,
        s.suggestedManagementAction,
      ]),
      headStyles: { fillColor: [0, 48, 135], fontSize: 8 },
      bodyStyles: { fontSize: 7.2 },
    });
  }

  autoTable(doc, {
    startY: cfo && cfo.cfoSuggestions.length > 0 ? (doc as any).lastAutoTable.finalY + 8 : 32,
    head: [['Uploaded Source Filename', 'Format', 'Detected Type', 'Records', 'Period']],
    body: model.sourceFiles.map((f) => [
      f.filename,
      f.format,
      f.detectedReportType,
      String(f.recordsExtracted),
      `${f.periodStart || '—'} to ${f.periodEnd || '—'}`,
    ]),
    headStyles: { fillColor: [80, 174, 93], fontSize: 8 },
    bodyStyles: { fontSize: 7.5 },
  });

  // Dynamic Page Numbering Across All Generated Pages
  const totalPages = doc.getNumberOfPages();
  for (let p = 1; p <= totalPages; p++) {
    doc.setPage(p);
    doc.setFontSize(8);
    doc.setTextColor(100, 116, 139);
    doc.text(
      `Generated on ${today} | GSPU Deterministic CFO Analytics Engine | Page ${p} of ${totalPages}`,
      105,
      290,
      { align: 'center' }
    );
  }

  const safeName = model.companyName.replace(/[^a-zA-Z0-9_-]/g, '_');
  doc.save(`GSPU_CFO_MIS_Report_${safeName}_${model.periodEnd}.pdf`);
}
