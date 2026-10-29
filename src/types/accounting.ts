export type FileFormat = 'XML' | 'XLSX' | 'XLS' | 'CSV' | 'JSON';

export type DetectedReportType =
  | 'TALLY_MASTERS_XML'
  | 'TALLY_TRANSACTIONS_XML'
  | 'TALLY_COMBINED_XML'
  | 'TALLY_OUTSTANDING_XML'
  | 'TRIAL_BALANCE_EXCEL'
  | 'PROFIT_LOSS_EXCEL'
  | 'BALANCE_SHEET_EXCEL'
  | 'CASH_FLOW_EXCEL'
  | 'COMPARATIVE_STATEMENTS_EXCEL'
  | 'BUDGET_EXCEL'
  | 'CFO_MONTHLY_PACK_EXCEL'
  | 'SALES_REGISTER_EXCEL'
  | 'PURCHASE_REGISTER_EXCEL'
  | 'DAYBOOK_EXCEL'
  | 'OUTSTANDING_BILLS_EXCEL'
  | 'GSTR1_RETURN'
  | 'GSTR2B_RETURN'
  | 'GSTR3B_RETURN'
  | 'GST_REGISTER_EXCEL'
  | 'CUSTOM_MAPPED_REPORT'
  | 'UNKNOWN_UNSUPPORTED';

export type AccountClassification =
  | 'Sales'
  | 'SalesReturn'
  | 'OtherOperatingIncome'
  | 'OtherIncome'
  | 'Purchases'
  | 'PurchaseReturn'
  | 'DirectExpenses'
  | 'EmployeeExpenses'
  | 'AdminExpenses'
  | 'SellingExpenses'
  | 'FinanceCosts'
  | 'Depreciation'
  | 'PriorPeriodExpenses'
  | 'CurrentTax'
  | 'DeferredTax'
  | 'TaxExpense'
  | 'FixedAssets'
  | 'Investments'
  | 'NonCurrentAssets'
  | 'Inventory'
  | 'TradeReceivables'
  | 'CashAndEquivalents'
  | 'BankBalances'
  | 'OtherCurrentAssets'
  | 'Equity'
  | 'LongTermBorrowings'
  | 'ShortTermBorrowings'
  | 'TradePayables'
  | 'TaxLiabilities'
  | 'OtherCurrentLiabilities'
  | 'Suspense'
  | 'Unclassified';

export type VoucherType =
  | 'Sales'
  | 'Purchase'
  | 'Receipt'
  | 'Payment'
  | 'Contra'
  | 'Journal'
  | 'Credit Note'
  | 'Debit Note'
  | 'Other';

export interface ImportValidationMessage {
  severity: 'ERROR' | 'WARNING' | 'INFO';
  code: string;
  message: string;
  details?: string;
}

export interface ColumnMappingConfig {
  dateCol?: string;
  voucherNoCol?: string;
  voucherTypeCol?: string;
  ledgerNameCol?: string;
  parentGroupCol?: string;
  partyNameCol?: string;
  gstinCol?: string;
  debitCol?: string;
  creditCol?: string;
  openingBalanceCol?: string;
  closingBalanceCol?: string;
  priorPeriodCol?: string;
  amountCol?: string;
  taxableValueCol?: string;
  igstCol?: string;
  cgstCol?: string;
  sgstCol?: string;
  cessCol?: string;
  dueDateCol?: string;
  narrationCol?: string;
  reportCategoryOverride?: DetectedReportType;
}

export interface SourceFile {
  id: string;
  filename: string;
  format: FileFormat;
  sizeBytes: number;
  uploadedAt: string;
  detectedReportType: DetectedReportType;
  extractedCompanyName?: string;
  periodStart?: string;
  periodEnd?: string;
  currentPeriodLabel?: string;
  comparativePeriodLabel?: string;
  hasMonthlyGranularity?: boolean;
  recordsExtracted: number;
  status: 'VALID' | 'WARNING' | 'NEEDS_MAPPING' | 'ERROR';
  validationMessages: ImportValidationMessage[];
  rawHeaders?: string[];
  rawSampleRows?: Record<string, string | number>[];
  columnMapping?: ColumnMappingConfig;
  isDemoData?: boolean;
  sheetNames?: string[];
}

export interface LedgerGroup {
  name: string;
  parentGroup: string;
  classification: AccountClassification;
  isRevenue: boolean;
  isDeemedPositive: boolean;
  sourceFileId: string;
}

export interface Ledger {
  id: string;
  name: string;
  normalizedName: string;
  parentGroup: string;
  classification: AccountClassification;
  gstin?: string;
  stateName?: string;
  openingBalancePaise: number; // Positive = Debit, Negative = Credit
  closingBalanceReportedPaise?: number; // From Trial Balance / Balance Sheet if imported (+Dr / -Cr)
  previousPeriodPaise?: number; // Prior period comparative value in natural statement sign (e.g. + for standard line, - for contra/credit tax)
  previousPeriodSignedPaise?: number; // Prior period in +Dr / -Cr convention
  budgetPaise?: number; // Budgeted value if imported from P&L / Budget sheet
  monthlyReportedPaise?: Record<string, number>; // Month key YYYY-MM -> signed paise from multi-column statement
  worksheetName?: string;
  cellReference?: string;
  priorCellReference?: string;
  reportingPeriodLabel?: string;
  comparativePeriodLabel?: string;
  statementSourceType?: 'PNL' | 'BALANCE_SHEET' | 'TRIAL_BALANCE' | 'TALLY_XML';
  entityName?: string;
  currency?: string;
  tbDebitPaise?: number;
  tbCreditPaise?: number;
  sourceFileIds: string[];
  masterGuid?: string;
}

export type CashFlowLineCategory =
  | 'CFO_OPERATING_ITEM'
  | 'CFO_WORKING_CAPITAL_CHANGE'
  | 'CFO_TAXES_PAID'
  | 'CFO_INTERMEDIATE_SUBTOTAL'
  | 'CFO_OPERATING_TOTAL'
  | 'CFI_CAPEX'
  | 'CFI_INVESTING_ITEM'
  | 'CFI_INVESTING_TOTAL'
  | 'CFF_BORROWINGS'
  | 'CFF_FINANCING_ITEM'
  | 'CFF_FINANCING_TOTAL'
  | 'CF_NET_CHANGE'
  | 'CF_OPENING_CASH'
  | 'CF_CLOSING_CASH';

export interface CashFlowStatementItem {
  id: string;
  particulars: string;
  normalizedParticulars: string;
  category: CashFlowLineCategory;
  currentPeriodPaise: number;
  previousPeriodPaise?: number;
  budgetPaise?: number;
  monthlyValuesPaise?: Record<string, number>;
  isSubtotalOrTotal: boolean;
  sourceFileId: string;
  sourceFilename: string;
  worksheetName?: string;
  cellReference?: string;
  priorCellReference?: string;
  reportingPeriodLabel?: string;
  comparativePeriodLabel?: string;
  entityName?: string;
}

export interface ReportedStatementSubtotal {
  id: string;
  statementType: 'PNL' | 'BALANCE_SHEET' | 'CASH_FLOW';
  label: string;
  normalizedLabel: string;
  reportedCurrentPaise: number;
  reportedPreviousPaise?: number;
  reportedBudgetPaise?: number;
  sourceFileId: string;
  sourceFilename: string;
  worksheetName?: string;
  cellReference?: string;
  priorCellReference?: string;
  reportingPeriodLabel?: string;
  comparativePeriodLabel?: string;
}

export interface BillAllocation {
  billName: string;
  billType: 'New Ref' | 'Agst Ref' | 'Advance' | 'On Account';
  amountPaise: number; // Positive = Debit, Negative = Credit
  dueDate?: string;
  creditDays?: number;
}

export interface InventoryMovement {
  stockItemName: string;
  quantity: number;
  unit: string;
  ratePaise: number;
  amountPaise: number; // Positive = Inward/Debit, Negative = Outward/Credit
  godownName?: string;
}

export interface VoucherEntry {
  ledgerName: string;
  normalizedLedgerName: string;
  amountPaise: number; // Positive = Debit, Negative = Credit
  isPartyLedger?: boolean;
  billAllocations: BillAllocation[];
  costCentre?: string;
}

export interface Voucher {
  id: string;
  sourceRecordId: string;
  guid?: string;
  masterId?: string;
  date: string; // YYYY-MM-DD
  voucherNumber: string;
  voucherType: VoucherType;
  originalVoucherType: string;
  partyName?: string;
  partyGstin?: string;
  placeOfSupply?: string;
  narration?: string;
  referenceNumber?: string;
  referenceDate?: string;
  entries: VoucherEntry[];
  inventoryMovements: InventoryMovement[];
  totalDebitPaise: number;
  totalCreditPaise: number;
  isBalanced: boolean;
  sourceFileId: string;
  sourceFilename: string;
  isSupplementaryRegister?: boolean;
}

export interface OutstandingInvoice {
  id: string;
  partyName: string;
  normalizedPartyName: string;
  category: 'RECEIVABLE' | 'PAYABLE';
  billNumber: string;
  billDate: string; // YYYY-MM-DD
  dueDate: string; // YYYY-MM-DD
  originalAmountPaise: number;
  adjustedAmountPaise: number;
  pendingAmountPaise: number;
  daysOverdue: number;
  invoiceAgeDays: number;
  relatedVoucherIds: string[];
  sourceFileId: string;
  sourceFilename: string;
  isFromBillAllocation: boolean;
}

export interface StockItem {
  name: string;
  parentGroup: string;
  unit: string;
  hsnCode?: string;
  gstRate?: number;
  openingQty: number;
  openingValuePaise: number;
  closingQty: number;
  closingValuePaise: number;
  sourceFileId: string;
}

export interface GSTInvoice {
  id: string;
  source: 'TALLY_SALES' | 'TALLY_PURCHASE' | 'GSTR1' | 'GSTR2B';
  documentType: 'INVOICE' | 'CREDIT_NOTE' | 'DEBIT_NOTE';
  partyGstin: string;
  partyName: string;
  invoiceNumber: string;
  normalizedInvoiceNumber: string;
  invoiceDate: string; // YYYY-MM-DD
  returnPeriod: string; // YYYY-MM
  supplyType: 'B2B' | 'B2C' | 'EXP' | 'SEZ' | 'CDNR';
  placeOfSupply?: string;
  reverseCharge: boolean;
  itcEligibility?: 'ELIGIBLE' | 'INELIGIBLE' | 'UNSPECIFIED';
  taxRate?: number;
  taxableValuePaise: number;
  igstPaise: number;
  cgstPaise: number;
  sgstPaise: number;
  cessPaise: number;
  totalTaxPaise: number;
  invoiceTotalPaise: number;
  sourceFileId: string;
  sourceFilename: string;
  linkedVoucherId?: string;
}

export interface GSTR3BSummary {
  id: string;
  returnPeriod: string; // YYYY-MM
  outwardTaxableValuePaise: number;
  outwardIgstPaise: number;
  outwardCgstPaise: number;
  outwardSgstPaise: number;
  outwardCessPaise: number;
  rcmTaxableValuePaise: number;
  rcmIgstPaise: number;
  rcmCgstPaise: number;
  rcmSgstPaise: number;
  eligibleItcIgstPaise: number;
  eligibleItcCgstPaise: number;
  eligibleItcSgstPaise: number;
  eligibleItcCessPaise: number;
  ineligibleItcPaise: number;
  reversedItcPaise: number;
  sourceFileId: string;
  sourceFilename: string;
}

export type AuditExceptionCategory =
  | 'EXACT_DUPLICATE_VOUCHER'
  | 'POSSIBLE_DUPLICATE_VOUCHER'
  | 'UNBALANCED_VOUCHER'
  | 'NEGATIVE_CASH_BALANCE'
  | 'MISSING_LEDGER_CLASSIFICATION'
  | 'UNUSUAL_JOURNAL_ENTRY'
  | 'ROUND_AMOUNT_TRANSACTION'
  | 'SUSPENSE_ACCOUNT_BALANCE'
  | 'PERIOD_END_ADJUSTMENT'
  | 'OPENING_CLOSING_TB_DIFFERENCE'
  | 'MISSING_TRANSACTION_REFERENCE';

export interface AuditException {
  id: string;
  date: string;
  voucherId?: string;
  voucherNumber: string;
  voucherType: string;
  ledgerName: string;
  partyName: string;
  amountPaise: number;
  category: AuditExceptionCategory;
  severity: 'HIGH' | 'MEDIUM' | 'LOW';
  explanation: string;
  recommendedProcedure: string;
  sourceFileId: string;
  sourceFilename: string;
}

export interface AuditReviewState {
  reviewed: boolean;
  reviewerNote: string;
  updatedAt?: string;
}

export interface ComputedLedgerSummary {
  ledger: Ledger;
  openingBalancePaise: number; // +Dr / -Cr
  periodDebitPaise: number;
  periodCreditPaise: number;
  computedClosingPaise: number; // +Dr / -Cr
  effectiveClosingPaise: number; // Uses reconciled TB closing if no complete voucher stream, else computed
  tbDiscrepancyPaise: number;
  transactionCount: number;
  monthlyNetPaise: Record<string, { debitPaise: number; creditPaise: number; netPaise: number }>;
  voucherIds: string[];
}

export interface StatementLineItem {
  code: string;
  label: string;
  classificationFilter: AccountClassification[];
  amountPaise: number;
  previousAmountPaise?: number;
  variancePaise?: number;
  variancePercent?: number | null;
  varianceLabel?: string;
  formulaExplanation: string;
  ledgerNames: string[];
  cellReferences?: string[];
  priorCellReferences?: string[];
  worksheetNames?: string[];
  isSubtotal?: boolean;
  isAvailable: boolean;
  requiresMapping?: boolean;
}

export interface GSTMatchResult {
  id: string;
  reconciliationType: 'SALES_VS_GSTR1' | 'PURCHASE_VS_GSTR2B';
  matchStatus: 'EXACT' | 'NORMALISED' | 'TOLERANCE' | 'PROBABLE' | 'MISSING_IN_PORTAL' | 'MISSING_IN_BOOKS';
  booksInvoice?: GSTInvoice;
  portalInvoice?: GSTInvoice;
  taxableDiffPaise: number;
  igstDiffPaise: number;
  cgstDiffPaise: number;
  sgstDiffPaise: number;
  totalTaxDiffPaise: number;
  discrepancyReasons: string[];
}

export interface FinancialRatioResult {
  id: string;
  category: 'LIQUIDITY' | 'PROFITABILITY' | 'EFFICIENCY' | 'SOLVENCY';
  name: string;
  formula: string;
  numeratorLabel: string;
  denominatorLabel: string;
  numeratorPaise: number;
  denominatorPaise: number;
  currentValue: number | null;
  previousValue: number | null;
  unit: 'RATIO' | 'PERCENT' | 'DAYS' | 'CURRENCY';
  formattedCurrent: string;
  formattedPrevious: string;
  status: 'COMPUTED' | 'ZERO_DENOMINATOR' | 'NEGATIVE_DENOMINATOR' | 'INSUFFICIENT_DATA';
  explanation: string;
  accountMappings: string[];
  isMaterialChange: boolean;
}

export interface FinancialInsight {
  id: string;
  category:
    | 'Revenue performance'
    | 'Profitability'
    | 'Cost control'
    | 'Working capital'
    | 'Liquidity'
    | 'Financial risk'
    | 'Accounting exceptions'
    | 'GST discrepancies';
  priority: 'Positive development' | 'Attention required' | 'Significant risk indicator' | 'Informational';
  headline: string;
  narrative: string;
  currentValueFormatted: string;
  comparativeValueFormatted: string;
  absoluteMovementFormatted: string;
  percentageMovementFormatted: string;
  formulaOrRule: string;
  supportingLedgers: string[];
  sourceFiles: string[];
  drillDownCategory?: AccountClassification | 'EXCEPTIONS' | 'GST' | 'RECEIVABLES' | 'PAYABLES';
}

export interface PolicySettings {
  roundingTolerancePaise: number; // Default 100 paise = ₹1.00
  materialityThresholdPaise: number; // Default 10000000 paise = ₹1,00,000
  roundSumThresholdPaise: number; // Default 5000000 paise = ₹50,000
  gstTolerancePaise: number; // Default 200 paise = ₹2.00
  ageingMode: 'DUE_DATE' | 'INVOICE_DATE';
  defaultCreditDays: number; // Used only when bill due date is absent
  fyStartMonth: number; // 4 = April
  debtServiceAnnualPaise?: number; // Optional user-supplied debt service schedule for DSCR
}

export interface ConsolidatedAccountingModel {
  companyName: string;
  detectedCompanyNames: string[];
  extractedStatementCompanyName?: string;
  hasCompanyMismatch: boolean;
  hasUserEnteredCompanyMismatch?: boolean;
  periodStart: string;
  periodEnd: string;
  currentPeriodLabel?: string;
  comparativePeriodLabel?: string;
  hasMonthlyGranularity?: boolean;
  availableMonths: string[]; // YYYY-MM sorted (ONLY genuine monthly data, never repeated annual data!)
  availableEntities?: string[];
  detectedCurrency?: string;
  sourceFiles: SourceFile[];
  ledgers: Map<string, Ledger>;
  ledgerSummaries: Map<string, ComputedLedgerSummary>;
  vouchers: Voucher[];
  deduplicatedVouchers: Voucher[];
  outstandingInvoices: OutstandingInvoice[];
  stockItems: StockItem[];
  gstInvoices: GSTInvoice[];
  gstr3bSummaries: GSTR3BSummary[];
  cashFlowItems?: CashFlowStatementItem[];
  reportedSubtotals?: ReportedStatementSubtotal[];
  auditExceptions: AuditException[];
  hasCompleteVoucherData: boolean;
  hasTrialBalanceData: boolean;
  hasBalanceSheetStatement?: boolean;
  hasProfitLossStatement?: boolean;
  hasBillLevelAgeingData: boolean;
  hasPriorPeriodData: boolean;
  hasCashFlowStatementData?: boolean;
  hasBudgetData?: boolean;
}

// ============================================================================
// DECIMAL-SAFE PAISE ARITHMETIC & INDIAN NUMBER FORMATTING UTILITIES
// ============================================================================

export function toPaise(amountInRupees: number | string | null | undefined): number {
  if (amountInRupees === null || amountInRupees === undefined || amountInRupees === '') return 0;
  if (typeof amountInRupees === 'number') {
    if (!Number.isFinite(amountInRupees)) return 0;
    return Math.round(amountInRupees * 100);
  }
  const cleaned = String(amountInRupees)
    .replace(/₹/g, '')
    .replace(/,/g, '')
    .replace(/\s+/g, '')
    .trim();
  if (!cleaned) return 0;

  // Handle Dr / Cr suffixes common in Tally Excel exports
  let multiplier = 1;
  let numericPart = cleaned;
  if (/dr\.?$/i.test(numericPart)) {
    numericPart = numericPart.replace(/dr\.?$/i, '').trim();
    multiplier = 1;
  } else if (/cr\.?$/i.test(numericPart)) {
    numericPart = numericPart.replace(/cr\.?$/i, '').trim();
    multiplier = -1;
  } else if (/^\(.*\)$/.test(numericPart)) {
    numericPart = numericPart.slice(1, -1).trim();
    multiplier = -1;
  }

  const parsed = Number(numericPart);
  if (!Number.isFinite(parsed)) return 0;
  return Math.round(parsed * 100) * multiplier;
}

export function fromPaise(paise: number): number {
  return Math.round(paise) / 100;
}

export function formatINR(paise: number, showDecimals = false, showSignDrCr = false): string {
  if (!Number.isFinite(paise)) return '₹0';
  const isNegative = paise < 0;
  const absRupees = Math.abs(paise) / 100;
  const formatted = new Intl.NumberFormat('en-IN', {
    style: 'currency',
    currency: 'INR',
    minimumFractionDigits: showDecimals ? 2 : 0,
    maximumFractionDigits: showDecimals ? 2 : 0,
  }).format(absRupees);

  if (showSignDrCr) {
    if (paise === 0) return formatted;
    return `${formatted} ${paise > 0 ? 'Dr' : 'Cr'}`;
  }
  return isNegative ? `-${formatted}` : formatted;
}

export function formatINRCompact(paise: number): string {
  if (!Number.isFinite(paise)) return '₹0';
  const sign = paise < 0 ? '-' : '';
  const absRupees = Math.abs(paise) / 100;
  if (absRupees >= 1_00_00_000) {
    return `${sign}₹${(absRupees / 1_00_00_000).toFixed(2)} Cr`;
  }
  if (absRupees >= 1_00_000) {
    return `${sign}₹${(absRupees / 1_00_000).toFixed(2)} Lakh`;
  }
  return formatINR(paise, false);
}

export function computeSafeVariance(currentPaise: number, previousPaise: number): {
  diffPaise: number;
  percent: number | null;
  label: string;
} {
  const diffPaise = currentPaise - previousPaise;
  if (previousPaise === 0 && currentPaise === 0) {
    return { diffPaise: 0, percent: 0, label: '0.0%' };
  }
  if (previousPaise === 0) {
    return { diffPaise, percent: null, label: 'New / Base Zero' };
  }
  if ((previousPaise < 0 && currentPaise > 0) || (previousPaise > 0 && currentPaise < 0)) {
    return { diffPaise, percent: null, label: 'Sign Reversal' };
  }
  const pct = (diffPaise / Math.abs(previousPaise)) * 100;
  const sign = pct > 0 ? '+' : '';
  return {
    diffPaise,
    percent: pct,
    label: `${sign}${pct.toFixed(1)}%`,
  };
}
