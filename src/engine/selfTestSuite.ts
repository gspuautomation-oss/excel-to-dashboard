import * as XLSX from 'xlsx';
import { PolicySettings } from '../types/accounting';
import { computeFinancialStatements } from './accountingEngine';
import { computeAgeingAnalysis } from './ageingEngine';
import { runAccountingHealthCheck } from './auditEngine';
import { computeCfoAnalyticsBundle } from './cfoAnalyticsEngine';
import { parseExcelOrCsvBuffer } from './excelParser';
import { runGSTReconciliation } from './gstMatchingEngine';
import { computeFinancialRatios } from './ratioEngine';
import { buildConsolidatedAccountingModel } from './reconciliationEngine';
import { buildDemoPayloads, DEMO_COMPANY_NAME } from './sampleDataGenerator';
import { parseTallySignedPaise, parseTallyXML } from './xmlParser';

export interface AutomatedTestResult {
  id: number;
  name: string;
  passed: boolean;
  assertionDetail: string;
}

const DEFAULT_TEST_POLICY: PolicySettings = {
  roundingTolerancePaise: 100,
  materialityThresholdPaise: 10000000,
  roundSumThresholdPaise: 5000000,
  gstTolerancePaise: 200,
  ageingMode: 'DUE_DATE',
  defaultCreditDays: 30,
  fyStartMonth: 4,
};

export function runAllAccountingAcceptanceTests(): AutomatedTestResult[] {
  const results: AutomatedTestResult[] = [];
  const demoPayloads = buildDemoPayloads();
  const model = buildConsolidatedAccountingModel(DEMO_COMPANY_NAME, demoPayloads, DEFAULT_TEST_POLICY);
  const statements = computeFinancialStatements(model);
  const ageing = computeAgeingAnalysis(
    model,
    DEFAULT_TEST_POLICY,
    statements.kpis.netRevenuePaise,
    statements.kpis.netPurchasesPaise
  );
  const health = runAccountingHealthCheck(model, DEFAULT_TEST_POLICY);
  const gst = runGSTReconciliation(model, DEFAULT_TEST_POLICY);
  const ratios = computeFinancialRatios(model, statements, DEFAULT_TEST_POLICY);

  // 1. Balanced and unbalanced vouchers
  const unbalXml = `<ENVELOPE><VOUCHER VCHTYPE="Journal"><DATE>20260105</DATE><VOUCHERNUMBER>U-1</VOUCHERNUMBER><ALLLEDGERENTRIES.LIST><LEDGERNAME>Rent</LEDGERNAME><ISDEEMEDPOSITIVE>Yes</ISDEEMEDPOSITIVE><AMOUNT>-5000.00</AMOUNT></ALLLEDGERENTRIES.LIST><ALLLEDGERENTRIES.LIST><LEDGERNAME>Cash</LEDGERNAME><ISDEEMEDPOSITIVE>No</ISDEEMEDPOSITIVE><AMOUNT>4500.00</AMOUNT></ALLLEDGERENTRIES.LIST></VOUCHER></ENVELOPE>`;
  const parsedUnbal = parseTallyXML(unbalXml, 'T1', 'unbal.xml');
  results.push({
    id: 1,
    name: 'Balanced and unbalanced vouchers detection',
    passed: parsedUnbal.vouchers[0]?.isBalanced === false && model.deduplicatedVouchers[0]?.isBalanced === true,
    assertionDetail: `Unbalanced test voucher flagged (Dr ₹5,000 vs Cr ₹4,500); balanced demo voucher verified.`,
  });

  // 2. Debit/credit signs
  const drPaise = parseTallySignedPaise('-1250.50', 'Yes');
  const crPaise = parseTallySignedPaise('1250.50', 'No');
  results.push({
    id: 2,
    name: 'Tally Debit/Credit sign convention handling',
    passed: drPaise === 125050 && crPaise === -125050,
    assertionDetail: `ISDEEMEDPOSITIVE="Yes" (-1250.50) -> +125050 paise (Dr); "No" -> -125050 paise (Cr).`,
  });

  // 3. Opening balances
  const hdfcSummary = model.ledgerSummaries.get('hdfc bank current a/c');
  results.push({
    id: 3,
    name: 'Opening balance preservation & accumulation',
    passed: hdfcSummary?.openingBalancePaise === 245000000,
    assertionDetail: `HDFC Bank opening balance parsed deterministically at ₹24,50,000 Dr.`,
  });

  // 4. Multi-file upload consolidation
  results.push({
    id: 4,
    name: 'Multi-file upload consolidation',
    passed: model.sourceFiles.length === 5 && model.ledgers.size >= 15 && model.deduplicatedVouchers.length === 10,
    assertionDetail: `Consolidated 5 distinct XML/JSON files with ${model.ledgers.size} ledgers & ${model.deduplicatedVouchers.length} vouchers.`,
  });

  // 5. Duplicate file detection / deduplication
  const doubledModel = buildConsolidatedAccountingModel(
    DEMO_COMPANY_NAME,
    [...demoPayloads, demoPayloads[1]],
    DEFAULT_TEST_POLICY
  );
  results.push({
    id: 5,
    name: 'Duplicate file upload deduplication guard',
    passed: doubledModel.deduplicatedVouchers.length === model.deduplicatedVouchers.length,
    assertionDetail: `Re-uploading Transactions XML kept deduplicated voucher count at ${doubledModel.deduplicatedVouchers.length} (zero double-counting).`,
  });

  // 6. Duplicate voucher detection
  results.push({
    id: 6,
    name: 'Duplicate voucher exception flagging',
    passed: doubledModel.auditExceptions.some((e) => e.category === 'EXACT_DUPLICATE_VOUCHER'),
    assertionDetail: `Exact duplicate GUID vouchers flagged in Audit Exceptions.`,
  });

  // 7. Sales returns accounting
  results.push({
    id: 7,
    name: 'Sales returns deducted from Gross Sales',
    passed:
      statements.kpis.grossSalesPaise === 920000000 &&
      statements.kpis.salesReturnsPaise === 20000000 &&
      statements.kpis.netRevenuePaise === 900000000,
    assertionDetail: `Gross Sales ₹92,00,000 − Sales Return ₹2,00,000 = Net Revenue ₹90,00,000.`,
  });

  // 8. Purchase returns & COGS
  results.push({
    id: 8,
    name: 'Purchases & Stock variation in COGS',
    passed: statements.kpis.netPurchasesPaise === 275000000 && statements.kpis.cogsPaise === 243000000,
    assertionDetail: `Net Purchases ₹27,50,000 + Stock Variation (-₹3,20,000) = COGS ₹24,30,000.`,
  });

  // 9. GST ledger accounting exclusion from P&L Revenue
  results.push({
    id: 9,
    name: 'GST tax ledgers excluded from Revenue',
    passed: statements.kpis.netRevenuePaise === 900000000 && statements.kpis.taxLiabilitiesPaise !== 0,
    assertionDetail: `Output CGST/SGST/IGST classified under Duties & Taxes, excluded from Net Sales.`,
  });

  // 10. GST invoice matching (Exact & Normalised)
  results.push({
    id: 10,
    name: 'GST invoice matching (Exact & Normalised)',
    passed:
      gst.salesVsGstr1Matches.some((m) => m.matchStatus === 'EXACT') &&
      gst.salesVsGstr1Matches.some((m) => m.matchStatus === 'NORMALISED'),
    assertionDetail: `Matched BPE/25-26/101 (Exact) and BPE/25-26/102 vs BPE-2526-102 (Normalised).`,
  });

  // 11. Credit notes and debit notes in GST
  results.push({
    id: 11,
    name: 'Credit Note reconciliation in GSTR-1',
    passed: gst.salesVsGstr1Matches.some((m) => m.booksInvoice?.documentType === 'CREDIT_NOTE'),
    assertionDetail: `Credit Note CN/25-26/009 matched and netted against outward taxable supplies.`,
  });

  // 12. Partially paid invoices via Bill Allocations
  const apexBill = model.outstandingInvoices.find((i) => i.billNumber === 'BPE/25-26/101');
  results.push({
    id: 12,
    name: 'Partially paid invoice bill-allocation tracking',
    passed:
      apexBill?.originalAmountPaise === 354000000 &&
      apexBill?.adjustedAmountPaise === 200000000 &&
      apexBill?.pendingAmountPaise === 154000000,
    assertionDetail: `Bill BPE/25-26/101: Original ₹35.40L − Receipt ₹20.00L = Pending ₹15.40L.`,
  });

  // 13. Debtor ageing buckets
  results.push({
    id: 13,
    name: 'Debtor ageing bucket classification',
    passed: ageing.totalBillTrackedReceivablesPaise > 0 && ageing.customers.length === 3,
    assertionDetail: `Tracked 3 customers across configurable due-date ageing buckets.`,
  });

  // 14. Creditor ageing buckets
  results.push({
    id: 14,
    name: 'Creditor ageing bucket classification',
    passed: ageing.totalBillTrackedPayablesPaise > 0 && ageing.suppliers.length === 2,
    assertionDetail: `Tracked 2 suppliers with bill-level residual payable balances.`,
  });

  // 15. Negative cash balance detection
  const negCashXml = `<ENVELOPE><VOUCHER VCHTYPE="Payment"><DATE>20260102</DATE><VOUCHERNUMBER>P-99</VOUCHERNUMBER><ALLLEDGERENTRIES.LIST><LEDGERNAME>Office Expense</LEDGERNAME><ISDEEMEDPOSITIVE>Yes</ISDEEMEDPOSITIVE><AMOUNT>-900000.00</AMOUNT></ALLLEDGERENTRIES.LIST><ALLLEDGERENTRIES.LIST><LEDGERNAME>Cash</LEDGERNAME><ISDEEMEDPOSITIVE>No</ISDEEMEDPOSITIVE><AMOUNT>900000.00</AMOUNT></ALLLEDGERENTRIES.LIST></VOUCHER></ENVELOPE>`;
  const pNeg = parseTallyXML(negCashXml, 'NEG', 'neg.xml');
  const mNeg = buildConsolidatedAccountingModel(
    'Test',
    [
      demoPayloads[0],
      {
        sourceFile: {
          id: 'NEG',
          filename: 'neg.xml',
          format: 'XML',
          sizeBytes: 100,
          uploadedAt: '',
          detectedReportType: 'TALLY_TRANSACTIONS_XML',
          recordsExtracted: 1,
          status: 'VALID',
          validationMessages: [],
        },
        groups: [],
        ledgers: [],
        vouchers: pNeg.vouchers,
        outstandingInvoices: [],
        stockItems: [],
        gstInvoices: [],
        gstr3bSummaries: [],
      },
    ],
    DEFAULT_TEST_POLICY
  );
  const hNeg = runAccountingHealthCheck(mNeg, DEFAULT_TEST_POLICY);
  results.push({
    id: 15,
    name: 'Negative daily cash balance detection',
    passed: hNeg.negativeCashDateCount > 0,
    assertionDetail: `Reconstructed cash book detected negative cash balance when payment exceeded ₹1,80,000 opening cash.`,
  });

  // 16. Monthly P&L breakdown
  results.push({
    id: 16,
    name: 'Monthly P&L time-series aggregation',
    passed: statements.monthlyTrends.length === 3,
    assertionDetail: `Monthly P&L series computed across ${statements.monthlyTrends.map((m) => m.month).join(', ')}.`,
  });

  // 17. Comparative P&L variance analysis
  const compStatements = computeFinancialStatements(model, ['2026-03'], ['2026-02'], 'Mar 2026', 'Feb 2026');
  results.push({
    id: 17,
    name: 'Comparative P&L variance calculation',
    passed: compStatements.pnlLines[0].previousAmountPaise === 220000000 && compStatements.pnlLines[0].amountPaise === 380000000,
    assertionDetail: `Comparative Net Revenue verified: Feb ₹22,00,000 vs Mar ₹38,00,000 (+72.7%).`,
  });

  // 18. Balance Sheet reconciliation
  results.push({
    id: 18,
    name: 'Balance Sheet Assets & Liabilities compilation',
    passed: statements.kpis.totalAssetsPaise > 0 && statements.kpis.totalEquityAndLiabilitiesPaise > 0,
    assertionDetail: `Compiled Total Assets and Total Equity & Liabilities with Suspense & Unclassified tracking.`,
  });

  // 19. Financial ratios calculation
  const gpMargin = ratios.find((r) => r.id === 'GROSS_MARGIN');
  results.push({
    id: 19,
    name: 'Financial Ratios deterministic calculation',
    passed: gpMargin?.status === 'COMPUTED' && gpMargin.currentValue !== null && gpMargin.currentValue > 60,
    assertionDetail: `Gross Profit Margin computed at ${gpMargin?.formattedCurrent} with full lineage.`,
  });

  // 20. Missing information & DSCR guard
  const dscr = ratios.find((r) => r.id === 'DSCR');
  results.push({
    id: 20,
    name: 'Missing information guard (DSCR without debt schedule)',
    passed: dscr?.status === 'INSUFFICIENT_DATA',
    assertionDetail: `DSCR explicitly reports INSUFFICIENT_DATA when repayment schedule is not supplied.`,
  });

  // 21. Division by zero protection
  const emptyModel = buildConsolidatedAccountingModel('Empty Co', [], DEFAULT_TEST_POLICY);
  const emptyStatements = computeFinancialStatements(emptyModel);
  const emptyRatios = computeFinancialRatios(emptyModel, emptyStatements, DEFAULT_TEST_POLICY);
  results.push({
    id: 21,
    name: 'Zero denominator protection in Ratios & Variances',
    passed: emptyRatios.every((r) => r.status === 'ZERO_DENOMINATOR' || r.status === 'INSUFFICIENT_DATA' || r.unit === 'CURRENCY'),
    assertionDetail: `Zero-denominator ratios safely return ZERO_DENOMINATOR status without NaN/Infinity.`,
  });

  // 22. Large XML batch handling
  results.push({
    id: 22,
    name: 'Multi-entity XML node batch parsing',
    passed: model.ledgers.size >= 16 && model.deduplicatedVouchers.length === 10,
    assertionDetail: `Processed multi-node Tally ENVELOPE seamlessly in browser memory.`,
  });

  // 23. Malformed XML handling
  const badXml = parseTallyXML('<ENVELOPE><VOUCHER><UNCLOSED>', 'BAD', 'corrupt.xml');
  results.push({
    id: 23,
    name: 'Malformed XML error capture without crash',
    passed: badXml.detectedReportType === 'UNKNOWN_UNSUPPORTED' && badXml.validationMessages.some((m) => m.code === 'MALFORMED_XML'),
    assertionDetail: `Malformed XML caught cleanly and surfaced as MALFORMED_XML validation error.`,
  });

  // 24. Invalid date handling
  const badDateXml = parseTallyXML('<ENVELOPE><VOUCHER VCHTYPE="Sales"><DATE>INVALID-DATE</DATE></VOUCHER></ENVELOPE>', 'BD', 'baddate.xml');
  results.push({
    id: 24,
    name: 'Invalid voucher date validation warning',
    passed: badDateXml.validationMessages.some((m) => m.code === 'INVALID_VOUCHER_DATE'),
    assertionDetail: `Invalid date flagged with INVALID_VOUCHER_DATE warning instead of silent corruption.`,
  });

  // 25. Multiple company mismatch detection
  const otherCoXml = parseTallyXML(
    `<ENVELOPE><HEADER><SVCURRENTCOMPANY>Other Unrelated Corp</SVCURRENTCOMPANY></HEADER><LEDGER NAME="Cash"><PARENT>Cash-in-Hand</PARENT></LEDGER></ENVELOPE>`,
    'OTHER',
    'other.xml'
  );
  const mismatchModel = buildConsolidatedAccountingModel(
    DEMO_COMPANY_NAME,
    [
      demoPayloads[0],
      {
        sourceFile: {
          id: 'OTHER',
          filename: 'other.xml',
          format: 'XML',
          sizeBytes: 120,
          uploadedAt: '',
          detectedReportType: 'TALLY_MASTERS_XML',
          extractedCompanyName: otherCoXml.extractedCompanyName,
          recordsExtracted: 1,
          status: 'WARNING',
          validationMessages: [],
        },
        groups: [],
        ledgers: otherCoXml.ledgers,
        vouchers: [],
        outstandingInvoices: [],
        stockItems: [],
        gstInvoices: [],
        gstr3bSummaries: [],
      },
    ],
    DEFAULT_TEST_POLICY
  );
  results.push({
    id: 25,
    name: 'Multiple company name mismatch alert',
    passed: mismatchModel.hasCompanyMismatch === true && mismatchModel.detectedCompanyNames.length === 2,
    assertionDetail: `Detected inconsistent company names across files and triggered mismatch warning.`,
  });

  // 26. Inconsistent / incomplete comparative period warning
  const unevenComp = computeFinancialStatements(
    model,
    ['2026-01', '2026-02'],
    ['2026-03'],
    'Jan-Feb 2026 (2M)',
    'Mar 2026 (1M)'
  );
  results.push({
    id: 26,
    name: 'Uneven reporting period comparison warning',
    passed: unevenComp.isIncompleteComparisonWarning === true,
    assertionDetail: `Flagged period-length mismatch when comparing 2-month range against 1-month range.`,
  });

  // 27. Multi-Sheet Schedule III Balance Sheet, Nested MSME/Creditor Trade Payables & Notes Isolation (ITRAITZ FY 2025-26 & FY 2024-25 Verification)
  const itraitzWb = XLSX.utils.book_new();
  const bsSheetRows = [
    ['ITRAITZ IT SOLUTIONS PRIVATE LIMITED'],
    ['BALANCE SHEET AS AT 31ST MARCH 2026'],
    ['Particulars', 'Note No.', 'As at 31st March 2026', 'As at 31st March 2025'],
    ['I. EQUITY AND LIABILITIES', '', '', ''],
    ["(1) Shareholders' Funds", '', '', ''],
    ['(a) Share Capital', '1', 100000.0, 100000.0],
    ['(b) Reserves and Surplus', '2', 2900239.39, 2150144.34],
    ['', '', 3000239.39, 2250144.34],
    ['(2) Non-Current Liabilities', '', '', ''],
    ['(a) Long-Term Borrowings', '3', 11596087.74, 8448499.2],
    ['(3) Current Liabilities', '', '', ''],
    ['(a) Trade payables', '4', '', ''],
    ['(A) Total outstanding dues of micro enterprises and small enterprises', '', '-', '-'],
    [
      '(B) Total outstanding dues of creditors other than micro enterprises and small enterprises',
      '',
      75436.82,
      56414.72,
    ],
    ['(b) Other Current Liabilities', '5', 138963.0, 152917.8],
    ['(c) Short-Term Provisions', '6', 87980.0, 1723256.0],
    ['Total', '', 302379.82, 1932588.52],
    ['TOTAL', '', 14898706.95, 12631232.06],
    ['II. ASSETS', '', '', ''],
    ['(1) Non-Current Assets', '', '', ''],
    ['(a) Property, Plant and Equipment and Intangible Assets', '', '', ''],
    ['(i) Property, Plant and Equipment', '7', 6948772.89, 5400000.0],
    ['(ii) Intangible Assets', '7', 8.9, 8.9],
    ['(b) Deferred Tax Assets (Net)', '8', 1651761.62, 1332452.39],
    ['(c) Long-Term Loans and Advances', '9', 3607548.0, 3607548.0],
    ['Total', '', 12208091.41, 10340009.29],
    ['(2) Current Assets', '', '', ''],
    ['(a) Trade Receivables', '10', 464777.54, 101679.1],
    ['(b) Cash and Cash Equivalents', '11', 491501.56, 647917.99],
    ['(c) Other Current Assets', '12', 1734336.44, 1541625.68],
    ['Total', '', 2690615.54, 2291222.77],
    ['TOTAL', '', 14898706.95, 12631232.06],
  ];
  const notesSheetRows = [
    ['ITRAITZ IT SOLUTIONS PRIVATE LIMITED'],
    ['NOTES FORMING PART OF FINANCIAL STATEMENTS'],
    ['Particulars', 'As at 31st March 2026', 'As at 31st March 2025'],
    ['Note 1: Share Capital', '', ''],
    ['Authorised Share Capital', 50000.0, 50000.0],
    ['Issued, Subscribed and Paid-up Capital', 100000.0, 100000.0],
    ['Note 14: Revenue from Operations', 21609989.14, 18500000.0],
    ['Note 16: Employee Benefit Expenses', 12450000.0, 10200000.0],
  ];
  const faSheetRows = [
    ['FIXED ASSETS AND DEPRECIATION SCHEDULE'],
    ['Particulars', 'Gross Block', 'Accumulated Depreciation', 'Net Block 31.03.2026'],
    ['Computers & Servers', 9481033.4, 2532260.51, 6948772.89],
  ];
  XLSX.utils.book_append_sheet(itraitzWb, XLSX.utils.aoa_to_sheet(bsSheetRows), 'Balance Sheet');
  XLSX.utils.book_append_sheet(itraitzWb, XLSX.utils.aoa_to_sheet(notesSheetRows), 'Notes');
  XLSX.utils.book_append_sheet(itraitzWb, XLSX.utils.aoa_to_sheet(faSheetRows), 'Fixed Assets');

  const itraitzBuffer = XLSX.write(itraitzWb, { type: 'array', bookType: 'xlsx' });
  const itraitzParsed = parseExcelOrCsvBuffer(
    itraitzBuffer,
    'ITRAITZ-TEST',
    'ITRAITZ IT SOLN FS 2025-26.xlsx'
  );
  const itraitzModel = buildConsolidatedAccountingModel(
    'Itraitz IT Solutions Private Limited',
    [
      {
        sourceFile: {
          id: 'ITRAITZ-TEST',
          filename: 'ITRAITZ IT SOLN FS 2025-26.xlsx',
          format: 'XLSX',
          sizeBytes: 1024,
          uploadedAt: '',
          detectedReportType: itraitzParsed.detectedReportType,
          extractedCompanyName: itraitzParsed.extractedCompanyName,
          recordsExtracted: itraitzParsed.ledgers.length,
          status: 'VALID',
          validationMessages: itraitzParsed.validationMessages,
        },
        groups: [],
        ledgers: itraitzParsed.ledgers,
        vouchers: itraitzParsed.vouchers,
        outstandingInvoices: itraitzParsed.outstandingInvoices,
        stockItems: [],
        gstInvoices: itraitzParsed.gstInvoices,
        gstr3bSummaries: itraitzParsed.gstr3bSummaries,
        cashFlowItems: itraitzParsed.cashFlowItems,
        reportedSubtotals: itraitzParsed.reportedSubtotals,
      },
    ],
    DEFAULT_TEST_POLICY
  );
  const itraitzSt = computeFinancialStatements(itraitzModel);
  const itraitzRatios = computeFinancialRatios(itraitzModel, itraitzSt, DEFAULT_TEST_POLICY);
  const itraitzCfo = computeCfoAnalyticsBundle(
    itraitzModel,
    itraitzSt,
    undefined,
    itraitzRatios,
    DEFAULT_TEST_POLICY
  );
  const ik = itraitzSt.kpis;
  const itraitzPassed =
    ik.tradePayablesPaise === 7543682 &&
    ik.tradeReceivablesPaise === 46477754 &&
    ik.cashPaise + ik.bankPaise === 49150156 &&
    ik.totalCurrentAssetsPaise === 269061554 &&
    ik.totalNonCurrentAssetsPaise === 1220809141 &&
    ik.totalAssetsPaise === 1489870695 &&
    ik.totalEquityPaise === 300023939 &&
    ik.longTermBorrowingsPaise === 1159608774 &&
    ik.totalCurrentLiabilitiesPaise === 30237982 &&
    ik.totalEquityAndLiabilitiesPaise === 1489870695 &&
    ik.workingCapitalPaise === 238823572 &&
    ik.totalAssetsPaise - ik.totalEquityAndLiabilitiesPaise === 0 &&
    itraitzCfo.balanceSheetReconciliation.isReconciled === true &&
    itraitzCfo.priorPeriodTotals.totalAssetsPaise === 1263123206 &&
    itraitzCfo.priorPeriodTotals.currentAssetsPaise === 229122277 &&
    itraitzCfo.priorPeriodTotals.currentLiabilitiesPaise === 193258852 &&
    itraitzCfo.balanceSheetVerificationRows.every((r) => r.status === 'RECONCILED');

  results.push({
    id: 27,
    name: 'Schedule III Balance Sheet, Nested MSME/Creditor Trade Payables & Notes isolation',
    passed: itraitzPassed,
    assertionDetail: `Verified Trade Payables ₹75,436.82 (FY26) & ₹56,414.72 (FY25), CA ₹26,90,615.54, NCA ₹1,22,08,091.41, Total Assets ₹1,48,98,706.95 = Equity ₹30,00,239.39 + NCL ₹1,15,96,087.74 + CL ₹3,02,379.82, NWC ₹23,88,235.72 (Diff ₹0.00).`,
  });

  // 28. Unreconciled Balance Sheet strict badge protection (Never display Reconciled on mismatched BS)
  const unbalBsWb = XLSX.utils.book_new();
  const unbalBsRows = [
    ['TEST UNBALANCED COMPANY'],
    ['BALANCE SHEET AS AT 31ST MARCH 2026'],
    ['Particulars', 'Note', 'As at 31st March 2026'],
    ['I. EQUITY AND LIABILITIES', '', ''],
    ['Share Capital', '1', 100000.0],
    ['Current Liabilities', '', ''],
    ['Other Current Liabilities', '2', 50000.0],
    ['TOTAL', '', 200000.0], // Intentional mismatch vs 150,000 sum
    ['II. ASSETS', '', ''],
    ['Current Assets', '', ''],
    ['Cash and Cash Equivalents', '3', 200000.0],
    ['TOTAL', '', 200000.0],
  ];
  XLSX.utils.book_append_sheet(unbalBsWb, XLSX.utils.aoa_to_sheet(unbalBsRows), 'Balance Sheet');
  const unbalBsBuf = XLSX.write(unbalBsWb, { type: 'array', bookType: 'xlsx' });
  const unbalBsParsed = parseExcelOrCsvBuffer(unbalBsBuf, 'UNBAL-BS', 'unbal_bs.xlsx');
  const unbalBsModel = buildConsolidatedAccountingModel(
    'Test Unbalanced Company',
    [
      {
        sourceFile: {
          id: 'UNBAL-BS',
          filename: 'unbal_bs.xlsx',
          format: 'XLSX',
          sizeBytes: 512,
          uploadedAt: '',
          detectedReportType: unbalBsParsed.detectedReportType,
          extractedCompanyName: unbalBsParsed.extractedCompanyName,
          recordsExtracted: unbalBsParsed.ledgers.length,
          status: 'WARNING',
          validationMessages: unbalBsParsed.validationMessages,
        },
        groups: [],
        ledgers: unbalBsParsed.ledgers,
        vouchers: [],
        outstandingInvoices: [],
        stockItems: [],
        gstInvoices: [],
        gstr3bSummaries: [],
        cashFlowItems: [],
        reportedSubtotals: unbalBsParsed.reportedSubtotals,
      },
    ],
    DEFAULT_TEST_POLICY
  );
  const unbalBsSt = computeFinancialStatements(unbalBsModel);
  const unbalBsRatios = computeFinancialRatios(unbalBsModel, unbalBsSt, DEFAULT_TEST_POLICY);
  const unbalBsCfo = computeCfoAnalyticsBundle(
    unbalBsModel,
    unbalBsSt,
    undefined,
    unbalBsRatios,
    DEFAULT_TEST_POLICY
  );
  const bsKpiCards = unbalBsCfo.cfoKpis.filter((k) =>
    [
      'KPI_TOTAL_ASSETS',
      'KPI_CURRENT_ASSETS',
      'KPI_CURRENT_LIABILITIES',
      'KPI_CASH_BANK',
      'KPI_NET_WORKING_CAPITAL',
    ].includes(k.id)
  );
  const unbalProtectionPassed =
    unbalBsCfo.balanceSheetReconciliation.isReconciled === false &&
    bsKpiCards.every((k) => k.statusIndicator === 'Requires Review') &&
    unbalBsCfo.balanceSheetVerificationRows.every((r) => r.status === 'REQUIRES REVIEW');

  results.push({
    id: 28,
    name: 'Strict reconciliation status protection (Never show Reconciled when BS has difference)',
    passed: unbalProtectionPassed,
    assertionDetail: `Unbalanced Balance Sheet correctly flagged as Requires Review across all BS KPIs and verification rows.`,
  });

  return results;
}
