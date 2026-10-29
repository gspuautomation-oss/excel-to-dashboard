import {
  AccountClassification,
  AuditException,
  CashFlowStatementItem,
  ComputedLedgerSummary,
  ConsolidatedAccountingModel,
  GSTInvoice,
  GSTR3BSummary,
  Ledger,
  LedgerGroup,
  OutstandingInvoice,
  PolicySettings,
  ReportedStatementSubtotal,
  SourceFile,
  StockItem,
  Voucher,
} from '../types/accounting';
import { classifyLedgerAndGroup, normalizeInvoiceNo, normalizeName } from './ledgerMapping';

export interface RawFilePayload {
  sourceFile: SourceFile;
  groups: LedgerGroup[];
  ledgers: Ledger[];
  vouchers: Voucher[];
  outstandingInvoices: OutstandingInvoice[];
  stockItems: StockItem[];
  gstInvoices: GSTInvoice[];
  gstr3bSummaries: GSTR3BSummary[];
  cashFlowItems?: CashFlowStatementItem[];
  reportedSubtotals?: ReportedStatementSubtotal[];
}

export function buildConsolidatedAccountingModel(
  manualCompanyName: string,
  payloads: RawFilePayload[],
  policy: PolicySettings,
  manualClassificationOverrides?: Record<string, AccountClassification>
): ConsolidatedAccountingModel {
  // If multiple payloads exist from different companies, scope to the most recently uploaded company with extractedCompanyName
  let effectivePayloads = payloads;
  const payloadsWithCompany = payloads.filter((p) => p.sourceFile.extractedCompanyName?.trim());
  if (payloadsWithCompany.length > 0) {
    const latestCompanyNorm = normalizeName(
      payloadsWithCompany[payloadsWithCompany.length - 1].sourceFile.extractedCompanyName!
    );
    const matchingPayloads = payloads.filter((p) => {
      const cName = p.sourceFile.extractedCompanyName?.trim();
      return !cName || normalizeName(cName) === latestCompanyNorm;
    });
    if (matchingPayloads.length > 0) {
      effectivePayloads = matchingPayloads;
    }
  }

  const sourceFiles = effectivePayloads.map((p) => p.sourceFile);
  const detectedCompanySet = new Set<string>();
  for (const p of payloads) {
    if (p.sourceFile.extractedCompanyName) {
      detectedCompanySet.add(p.sourceFile.extractedCompanyName.trim());
    }
  }
  const detectedCompanyNames = Array.from(detectedCompanySet);
  const normalizedDetected = new Set(detectedCompanyNames.map((n) => normalizeName(n)));
  const hasCompanyMismatch = normalizedDetected.size > 1;

  const extractedStatementCompanyName =
    effectivePayloads
      .slice()
      .reverse()
      .find((p) => p.sourceFile.extractedCompanyName)?.sourceFile.extractedCompanyName ||
    detectedCompanyNames[0];
  const manualTrimmed = manualCompanyName.trim();
  const isDefaultTemplateName = normalizeName(manualTrimmed).includes('gspu management consulting');
  const hasUserEnteredCompanyMismatch = Boolean(
    manualTrimmed &&
      !isDefaultTemplateName &&
      extractedStatementCompanyName &&
      normalizeName(manualTrimmed) !== normalizeName(extractedStatementCompanyName)
  );

  // Extracted company name from financial statements takes priority over empty or default template name
  const effectiveCompanyName =
    extractedStatementCompanyName && (!manualTrimmed || isDefaultTemplateName)
      ? extractedStatementCompanyName
      : manualTrimmed || extractedStatementCompanyName || 'Client Enterprise';

  // 1. Consolidate Ledgers across all uploaded files with strict 4-Tier Source Priority:
  // Priority 1: Primary Financial Statements (BALANCE_SHEET / PNL)
  // Priority 2: Supporting Notes / Schedules (used only for verification)
  // Priority 3: Trial Balance (only used when primary BS/PNL statements are not uploaded)
  // Priority 4: Vouchers / Registers (only used when BS/PNL and Trial Balance are not uploaded)
  const ledgers = new Map<string, Ledger>();
  let detectedCurrentPeriodLabel: string | undefined;
  let detectedComparativePeriodLabel: string | undefined;
  let hasBalanceSheetStatement = false;
  let hasProfitLossStatement = false;

  for (const sf of sourceFiles) {
    if (sf.detectedReportType === 'BALANCE_SHEET_EXCEL') hasBalanceSheetStatement = true;
    if (sf.detectedReportType === 'PROFIT_LOSS_EXCEL') hasProfitLossStatement = true;
  }
  for (const payload of effectivePayloads) {
    for (const incoming of payload.ledgers) {
      if (incoming.statementSourceType === 'BALANCE_SHEET') hasBalanceSheetStatement = true;
      if (incoming.statementSourceType === 'PNL') hasProfitLossStatement = true;
    }
  }

  for (const payload of effectivePayloads) {
    for (const incoming of payload.ledgers) {
      if (!incoming.normalizedName) continue;
      // Enforce Priority 1 vs Priority 3: If authoritative Balance Sheet or P&L exists, do not mix Trial Balance rows into statement ledgers
      if (
        incoming.statementSourceType === 'TRIAL_BALANCE' &&
        (hasBalanceSheetStatement || hasProfitLossStatement)
      ) {
        continue;
      }
      if (incoming.reportingPeriodLabel && !detectedCurrentPeriodLabel) {
        detectedCurrentPeriodLabel = incoming.reportingPeriodLabel;
      }
      if (incoming.comparativePeriodLabel && !detectedComparativePeriodLabel) {
        detectedComparativePeriodLabel = incoming.comparativePeriodLabel;
      }

      // Prevent P&L and Balance Sheet items or different Schedule III sections (e.g. Non-Current vs Current "Provisions" or "Loans and Advances") from colliding
      let key = incoming.normalizedName;
      const existingCandidate = ledgers.get(key);
      if (existingCandidate) {
        if (
          existingCandidate.statementSourceType &&
          incoming.statementSourceType &&
          existingCandidate.statementSourceType !== incoming.statementSourceType
        ) {
          key = `${incoming.normalizedName}_${incoming.statementSourceType.toLowerCase()}`;
        } else if (
          (incoming.statementSourceType === 'BALANCE_SHEET' ||
            incoming.statementSourceType === 'PNL') &&
          normalizeName(existingCandidate.parentGroup) !== normalizeName(incoming.parentGroup)
        ) {
          key = `${incoming.normalizedName}_${normalizeName(incoming.parentGroup).replace(/\s+/g, '_')}`;
        }
      }

      const existing = ledgers.get(key);
      const resolvedClass = classifyLedgerAndGroup(
        incoming.name,
        incoming.parentGroup,
        manualClassificationOverrides,
        incoming.statementSourceType
      );

      if (!existing) {
        ledgers.set(key, {
          ...incoming,
          normalizedName: key,
          classification: resolvedClass,
          sourceFileIds: [...incoming.sourceFileIds],
        });
      } else {
        // Merge without double-counting opening balances!
        // Masters XML takes precedence for openingBalancePaise; Primary BS/PNL takes precedence over Trial Balance
        const isIncomingMaster =
          payload.sourceFile.detectedReportType === 'TALLY_MASTERS_XML' ||
          payload.sourceFile.detectedReportType === 'TALLY_COMBINED_XML';
        const isIncomingPrimaryStatement =
          incoming.statementSourceType === 'BALANCE_SHEET' ||
          incoming.statementSourceType === 'PNL';
        const isExistingPrimaryStatement =
          existing.statementSourceType === 'BALANCE_SHEET' ||
          existing.statementSourceType === 'PNL';
        const isIncomingTB =
          payload.sourceFile.detectedReportType === 'TRIAL_BALANCE_EXCEL' ||
          payload.sourceFile.detectedReportType === 'BALANCE_SHEET_EXCEL' ||
          payload.sourceFile.detectedReportType === 'PROFIT_LOSS_EXCEL';

        if (isIncomingMaster && incoming.openingBalancePaise !== 0) {
          existing.openingBalancePaise = incoming.openingBalancePaise;
        } else if (existing.openingBalancePaise === 0 && incoming.openingBalancePaise !== 0) {
          existing.openingBalancePaise = incoming.openingBalancePaise;
        }

        if (
          isIncomingTB &&
          incoming.closingBalanceReportedPaise !== undefined &&
          (!isExistingPrimaryStatement || isIncomingPrimaryStatement)
        ) {
          existing.closingBalanceReportedPaise = incoming.closingBalanceReportedPaise;
          existing.tbDebitPaise = incoming.tbDebitPaise;
          existing.tbCreditPaise = incoming.tbCreditPaise;
        } else if (
          existing.closingBalanceReportedPaise === undefined &&
          incoming.closingBalanceReportedPaise !== undefined
        ) {
          existing.closingBalanceReportedPaise = incoming.closingBalanceReportedPaise;
        }

        if (
          incoming.previousPeriodPaise !== undefined &&
          (!isExistingPrimaryStatement || isIncomingPrimaryStatement || existing.previousPeriodPaise === undefined)
        ) {
          existing.previousPeriodPaise = incoming.previousPeriodPaise;
        }
        if (
          incoming.previousPeriodSignedPaise !== undefined &&
          (!isExistingPrimaryStatement || isIncomingPrimaryStatement || existing.previousPeriodSignedPaise === undefined)
        ) {
          existing.previousPeriodSignedPaise = incoming.previousPeriodSignedPaise;
        }
        if (incoming.budgetPaise !== undefined) {
          existing.budgetPaise = incoming.budgetPaise;
        }
        if (incoming.monthlyReportedPaise) {
          existing.monthlyReportedPaise = {
            ...(existing.monthlyReportedPaise || {}),
            ...incoming.monthlyReportedPaise,
          };
        }
        if (incoming.worksheetName && (!existing.worksheetName || isIncomingPrimaryStatement)) {
          existing.worksheetName = incoming.worksheetName;
          existing.cellReference = incoming.cellReference;
          existing.priorCellReference = incoming.priorCellReference;
        }
        if (incoming.reportingPeriodLabel && !existing.reportingPeriodLabel) {
          existing.reportingPeriodLabel = incoming.reportingPeriodLabel;
        }
        if (incoming.comparativePeriodLabel && !existing.comparativePeriodLabel) {
          existing.comparativePeriodLabel = incoming.comparativePeriodLabel;
        }
        if (incoming.statementSourceType && (!existing.statementSourceType || isIncomingPrimaryStatement)) {
          existing.statementSourceType = incoming.statementSourceType;
        }

        if (existing.parentGroup === 'Unclassified' && incoming.parentGroup !== 'Unclassified') {
          existing.parentGroup = incoming.parentGroup;
          existing.classification = resolvedClass;
        }
        if (!existing.gstin && incoming.gstin) existing.gstin = incoming.gstin;
        if (!existing.sourceFileIds.includes(payload.sourceFile.id)) {
          existing.sourceFileIds.push(payload.sourceFile.id);
        }
      }
    }
  }

  // 2. Consolidate & Deduplicate Vouchers
  // Precedence Policy: Complete Tally XML Vouchers > Supplementary Register Excel rows
  const allVouchers: Voucher[] = [];
  let hasPrimaryXmlVouchers = false;
  for (const p of effectivePayloads) {
    for (const v of p.vouchers) {
      if (!v.isSupplementaryRegister) {
        hasPrimaryXmlVouchers = true;
      }
      allVouchers.push(v);
    }
  }

  const deduplicatedVouchers: Voucher[] = [];
  const auditExceptions: AuditException[] = [];

  const seenGuidMap = new Map<string, Voucher>();
  const seenExactCompositeMap = new Map<string, Voucher>();
  const seenPossibleCompositeMap = new Map<string, Voucher>();

  for (const vch of allVouchers) {
    // If primary XML vouchers are present, check if a supplementary register row duplicates an XML voucher
    const normVchNo = normalizeInvoiceNo(vch.voucherNumber);
    const normParty = normalizeName(vch.partyName || '');
    const exactKey = `${vch.voucherType}|${normVchNo}|${vch.date}|${normParty}|${vch.totalDebitPaise}`;
    const possibleKey = `${vch.voucherType}|${vch.date}|${normParty}|${vch.totalDebitPaise}`;

    if (vch.guid && seenGuidMap.has(vch.guid)) {
      const orig = seenGuidMap.get(vch.guid)!;
      auditExceptions.push({
        id: `EXC_DUP_GUID_${vch.id}`,
        date: vch.date,
        voucherId: vch.id,
        voucherNumber: vch.voucherNumber || 'N/A',
        voucherType: vch.voucherType,
        ledgerName: vch.entries[0]?.ledgerName || 'Multiple Ledgers',
        partyName: vch.partyName || 'N/A',
        amountPaise: vch.totalDebitPaise,
        category: 'EXACT_DUPLICATE_VOUCHER',
        severity: 'HIGH',
        explanation: `Exact duplicate Tally GUID (${vch.guid}) already imported from ${orig.sourceFilename}. Excluded from consolidated totals to prevent double-counting.`,
        recommendedProcedure: 'Verify whether overlapping XML exports were uploaded for the same period.',
        sourceFileId: vch.sourceFileId,
        sourceFilename: vch.sourceFilename,
      });
      continue;
    }

    if (normVchNo && seenExactCompositeMap.has(exactKey)) {
      const orig = seenExactCompositeMap.get(exactKey)!;
      // If one is a supplementary register and original is primary XML, skip supplementary silently or log info
      auditExceptions.push({
        id: `EXC_DUP_EXACT_${vch.id}`,
        date: vch.date,
        voucherId: orig.id,
        voucherNumber: vch.voucherNumber || 'N/A',
        voucherType: vch.voucherType,
        ledgerName: vch.entries[0]?.ledgerName || 'Multiple Ledgers',
        partyName: vch.partyName || 'N/A',
        amountPaise: vch.totalDebitPaise,
        category: 'EXACT_DUPLICATE_VOUCHER',
        severity: orig.sourceFileId === vch.sourceFileId ? 'HIGH' : 'MEDIUM',
        explanation:
          orig.sourceFileId === vch.sourceFileId
            ? `Identical voucher number (${vch.voucherNumber}), date, party, and amount found twice inside ${vch.sourceFilename}.`
            : `Overlapping voucher (${vch.voucherNumber}) found across ${orig.sourceFilename} and ${vch.sourceFilename}. Deduplicated automatically.`,
        recommendedProcedure:
          orig.sourceFileId === vch.sourceFileId
            ? 'Inspect both vouchers in TallyPrime to confirm if a transaction was posted twice.'
            : 'Overlapping period upload handled by deduplication engine.',
        sourceFileId: vch.sourceFileId,
        sourceFilename: vch.sourceFilename,
      });

      // Only exclude from consolidation if it came from an overlapping file; if in same file with different GUID, flag as possible duplicate
      if (orig.sourceFileId !== vch.sourceFileId || (!vch.guid && !orig.guid)) {
        continue;
      }
    }

    // Check possible duplicate (same date, same party, same amount, different voucher number in same file)
    if (normParty && vch.totalDebitPaise > 0 && seenPossibleCompositeMap.has(possibleKey)) {
      const prior = seenPossibleCompositeMap.get(possibleKey)!;
      if (prior.voucherNumber !== vch.voucherNumber) {
        auditExceptions.push({
          id: `EXC_DUP_POSS_${vch.id}`,
          date: vch.date,
          voucherId: vch.id,
          voucherNumber: vch.voucherNumber || 'N/A',
          voucherType: vch.voucherType,
          ledgerName: vch.entries[0]?.ledgerName || 'Multiple Ledgers',
          partyName: vch.partyName || 'N/A',
          amountPaise: vch.totalDebitPaise,
          category: 'POSSIBLE_DUPLICATE_VOUCHER',
          severity: 'MEDIUM',
          explanation: `Suspected duplicate: Same date (${vch.date}), party (${vch.partyName}), and amount as Voucher #${prior.voucherNumber}. Flagged for review (not automatically deleted).`,
          recommendedProcedure: 'Check supporting supplier/customer bills to verify if two identical invoices were raised on the same day.',
          sourceFileId: vch.sourceFileId,
          sourceFilename: vch.sourceFilename,
        });
      }
    }

    if (vch.guid) seenGuidMap.set(vch.guid, vch);
    if (normVchNo) seenExactCompositeMap.set(exactKey, vch);
    if (normParty) seenPossibleCompositeMap.set(possibleKey, vch);

    deduplicatedVouchers.push(vch);
  }

  // Sort deduplicated vouchers chronologically
  deduplicatedVouchers.sort((a, b) => a.date.localeCompare(b.date));

  // Ensure every ledger referenced in vouchers exists in `ledgers` map
  for (const vch of deduplicatedVouchers) {
    for (const ent of vch.entries) {
      if (!ledgers.has(ent.normalizedLedgerName)) {
        const inferredClass = classifyLedgerAndGroup(
          ent.ledgerName,
          '',
          manualClassificationOverrides
        );
        ledgers.set(ent.normalizedLedgerName, {
          id: `AUTO_LED_${ent.normalizedLedgerName}`,
          name: ent.ledgerName,
          normalizedName: ent.normalizedLedgerName,
          parentGroup: 'Inferred from Vouchers',
          classification: inferredClass,
          openingBalancePaise: 0,
          sourceFileIds: [vch.sourceFileId],
        });
      } else {
        const existing = ledgers.get(ent.normalizedLedgerName)!;
        if (manualClassificationOverrides && manualClassificationOverrides[ent.normalizedLedgerName]) {
          existing.classification = manualClassificationOverrides[ent.normalizedLedgerName];
        }
      }
    }
  }

  // 3. Compute Ledger Summaries & Reconcile against Trial Balance
  const ledgerSummaries = new Map<string, ComputedLedgerSummary>();
  const monthSet = new Set<string>();
  let minDate = '';
  let maxDate = '';

  for (const [, ledger] of ledgers.entries()) {
    ledgerSummaries.set(ledger.normalizedName, {
      ledger,
      openingBalancePaise: ledger.openingBalancePaise,
      periodDebitPaise: 0,
      periodCreditPaise: 0,
      computedClosingPaise: ledger.openingBalancePaise,
      effectiveClosingPaise: ledger.openingBalancePaise,
      tbDiscrepancyPaise: 0,
      transactionCount: 0,
      monthlyNetPaise: {},
      voucherIds: [],
    });
  }

  for (const vch of deduplicatedVouchers) {
    const ym = vch.date.slice(0, 7);
    monthSet.add(ym);
    if (!minDate || vch.date < minDate) minDate = vch.date;
    if (!maxDate || vch.date > maxDate) maxDate = vch.date;

    for (const ent of vch.entries) {
      const summary = ledgerSummaries.get(ent.normalizedLedgerName);
      if (!summary) continue;

      if (ent.amountPaise > 0) {
        summary.periodDebitPaise += ent.amountPaise;
      } else {
        summary.periodCreditPaise += Math.abs(ent.amountPaise);
      }
      summary.computedClosingPaise += ent.amountPaise;
      summary.transactionCount += 1;
      if (!summary.voucherIds.includes(vch.id)) {
        summary.voucherIds.push(vch.id);
      }

      if (!summary.monthlyNetPaise[ym]) {
        summary.monthlyNetPaise[ym] = { debitPaise: 0, creditPaise: 0, netPaise: 0 };
      }
      if (ent.amountPaise > 0) {
        summary.monthlyNetPaise[ym].debitPaise += ent.amountPaise;
      } else {
        summary.monthlyNetPaise[ym].creditPaise += Math.abs(ent.amountPaise);
      }
      summary.monthlyNetPaise[ym].netPaise += ent.amountPaise;
    }
  }

  const hasCompleteVoucherData = hasPrimaryXmlVouchers || deduplicatedVouchers.length > 0;
  let hasTrialBalanceData = false;
  let hasPriorPeriodFromStatements = false;
  let hasBudgetData = false;

  for (const [, summary] of ledgerSummaries.entries()) {
    const { ledger } = summary;
    if (ledger.previousPeriodPaise !== undefined) {
      hasPriorPeriodFromStatements = true;
    }
    if (ledger.budgetPaise !== undefined) {
      hasBudgetData = true;
    }
    if (ledger.monthlyReportedPaise && summary.transactionCount === 0) {
      for (const [ym, valPaise] of Object.entries(ledger.monthlyReportedPaise)) {
        monthSet.add(ym);
        summary.monthlyNetPaise[ym] = {
          debitPaise: valPaise > 0 ? valPaise : 0,
          creditPaise: valPaise < 0 ? Math.abs(valPaise) : 0,
          netPaise: valPaise,
        };
      }
    }
    if (ledger.closingBalanceReportedPaise !== undefined) {
      hasTrialBalanceData = true;
      const isPrimaryStatementLedger =
        ledger.statementSourceType === 'BALANCE_SHEET' || ledger.statementSourceType === 'PNL';
      if (!isPrimaryStatementLedger && hasCompleteVoucherData && summary.transactionCount > 0) {
        summary.tbDiscrepancyPaise =
          summary.computedClosingPaise - ledger.closingBalanceReportedPaise;
        summary.effectiveClosingPaise = summary.computedClosingPaise;
      } else {
        // Priority 1: Primary Financial Statements (or Trial Balance without vouchers) always preserve their reported closing balance
        if (ledger.tbDebitPaise !== undefined) summary.periodDebitPaise = ledger.tbDebitPaise;
        if (ledger.tbCreditPaise !== undefined) summary.periodCreditPaise = ledger.tbCreditPaise;
        summary.effectiveClosingPaise = ledger.closingBalanceReportedPaise;
      }
    } else {
      summary.effectiveClosingPaise = summary.computedClosingPaise;
    }
  }

  // 4. Reconstruct Bill-Level Outstanding Invoices from BILLALLOCATIONS + Uploaded Outstanding Files
  const outstandingMap = new Map<string, OutstandingInvoice>();

  // First load explicitly uploaded Outstanding Invoices
  for (const p of effectivePayloads) {
    for (const out of p.outstandingInvoices) {
      const key = `${out.normalizedPartyName}|${normalizeInvoiceNo(out.billNumber)}`;
      outstandingMap.set(key, { ...out });
    }
  }

  // Reconstruct from Voucher BillAllocations if not already provided by an explicit Outstanding export
  const billTrackingMap = new Map<
    string,
    {
      partyName: string;
      normalizedPartyName: string;
      billNumber: string;
      billDate: string;
      dueDate: string;
      debitPaise: number;
      creditPaise: number;
      voucherIds: string[];
      sourceFileId: string;
      sourceFilename: string;
    }
  >();

  for (const vch of deduplicatedVouchers) {
    for (const ent of vch.entries) {
      if (!ent.billAllocations || ent.billAllocations.length === 0) continue;
      const led = ledgers.get(ent.normalizedLedgerName);
      const partyName = ent.ledgerName || vch.partyName || 'Counterparty';
      const normParty = normalizeName(partyName);

      for (const ba of ent.billAllocations) {
        if (ba.billType === 'On Account') continue;
        const normBill = normalizeInvoiceNo(ba.billName);
        if (!normBill) continue;
        const key = `${normParty}|${normBill}`;

        let rec = billTrackingMap.get(key);
        if (!rec) {
          const defaultDue = new Date(vch.date);
          defaultDue.setDate(defaultDue.getDate() + (ba.creditDays || policy.defaultCreditDays || 30));
          rec = {
            partyName,
            normalizedPartyName: normParty,
            billNumber: ba.billName,
            billDate: vch.date,
            dueDate: ba.dueDate || defaultDue.toISOString().slice(0, 10),
            debitPaise: 0,
            creditPaise: 0,
            voucherIds: [],
            sourceFileId: vch.sourceFileId,
            sourceFilename: vch.sourceFilename,
          };
          billTrackingMap.set(key, rec);
        }

        if (ba.billType === 'New Ref') {
          rec.billDate = vch.date;
          if (ba.dueDate) rec.dueDate = ba.dueDate;
        }
        if (ba.amountPaise > 0) {
          rec.debitPaise += ba.amountPaise;
        } else {
          rec.creditPaise += Math.abs(ba.amountPaise);
        }
        if (!rec.voucherIds.includes(vch.id)) {
          rec.voucherIds.push(vch.id);
        }
      }
    }
  }

  for (const [key, rec] of billTrackingMap.entries()) {
    if (outstandingMap.has(key)) continue;
    const netPaise = rec.debitPaise - rec.creditPaise;
    if (Math.abs(netPaise) <= policy.roundingTolerancePaise) continue;

    const led = ledgers.get(rec.normalizedPartyName);
    const isPayable =
      led?.classification === 'TradePayables' || netPaise < 0;
    const originalPaise = isPayable ? rec.creditPaise : rec.debitPaise;
    const adjustedPaise = isPayable ? rec.debitPaise : rec.creditPaise;
    const pendingPaise = Math.abs(netPaise);

    outstandingMap.set(key, {
      id: `BILL_ALLOC_${key}`,
      partyName: rec.partyName,
      normalizedPartyName: rec.normalizedPartyName,
      category: isPayable ? 'PAYABLE' : 'RECEIVABLE',
      billNumber: rec.billNumber,
      billDate: rec.billDate,
      dueDate: rec.dueDate,
      originalAmountPaise: originalPaise > 0 ? originalPaise : pendingPaise,
      adjustedAmountPaise: adjustedPaise,
      pendingAmountPaise: pendingPaise,
      daysOverdue: 0,
      invoiceAgeDays: 0,
      relatedVoucherIds: rec.voucherIds,
      sourceFileId: rec.sourceFileId,
      sourceFilename: rec.sourceFilename,
      isFromBillAllocation: true,
    });
  }

  // Reference date for ageing = periodEnd or latest date in dataset
  const refDateStr = maxDate || '2026-03-31';
  const refDateMs = new Date(refDateStr).getTime();

  const outstandingInvoices: OutstandingInvoice[] = [];
  for (const inv of outstandingMap.values()) {
    const dueMs = new Date(inv.dueDate).getTime();
    const billMs = new Date(inv.billDate).getTime();
    const daysOverdue = Math.floor((refDateMs - dueMs) / (1000 * 60 * 60 * 24));
    const invoiceAgeDays = Math.max(0, Math.floor((refDateMs - billMs) / (1000 * 60 * 60 * 24)));
    outstandingInvoices.push({
      ...inv,
      daysOverdue,
      invoiceAgeDays,
    });
  }

  // 5. Consolidate Stock Items, GST Invoices, GSTR-3B Summaries, Cash Flow Items & Reported Subtotals
  const stockItems: StockItem[] = [];
  const gstInvoices: GSTInvoice[] = [];
  const gstr3bSummaries: GSTR3BSummary[] = [];
  const cashFlowItems: CashFlowStatementItem[] = [];
  const reportedSubtotals: ReportedStatementSubtotal[] = [];
  const seenGstKey = new Set<string>();
  const seenCfKey = new Set<string>();

  for (const p of effectivePayloads) {
    stockItems.push(...p.stockItems);
    for (const gi of p.gstInvoices) {
      const k = `${gi.source}|${gi.partyGstin}|${gi.normalizedInvoiceNumber}|${gi.invoiceDate}|${gi.taxableValuePaise}`;
      if (!seenGstKey.has(k)) {
        seenGstKey.add(k);
        gstInvoices.push(gi);
      }
    }
    gstr3bSummaries.push(...p.gstr3bSummaries);
    if (p.cashFlowItems) {
      for (const cf of p.cashFlowItems) {
        const ck = `${cf.normalizedParticulars}|${cf.category}`;
        if (!seenCfKey.has(ck)) {
          seenCfKey.add(ck);
          cashFlowItems.push(cf);
          if (cf.previousPeriodPaise !== undefined) hasPriorPeriodFromStatements = true;
          if (cf.budgetPaise !== undefined) hasBudgetData = true;
          if (cf.reportingPeriodLabel && !detectedCurrentPeriodLabel) {
            detectedCurrentPeriodLabel = cf.reportingPeriodLabel;
          }
          if (cf.comparativePeriodLabel && !detectedComparativePeriodLabel) {
            detectedComparativePeriodLabel = cf.comparativePeriodLabel;
          }
        }
      }
    }
    if (p.reportedSubtotals) {
      for (const sub of p.reportedSubtotals) {
        reportedSubtotals.push(sub);
        if (sub.reportedPreviousPaise !== undefined) hasPriorPeriodFromStatements = true;
      }
    }
  }

  for (const sf of sourceFiles) {
    if (!minDate && sf.periodStart) minDate = sf.periodStart;
    if (!maxDate && sf.periodEnd) maxDate = sf.periodEnd;
  }

  const availableMonths = Array.from(monthSet).sort();

  return {
    companyName: effectiveCompanyName,
    extractedStatementCompanyName,
    detectedCompanyNames,
    hasCompanyMismatch,
    hasUserEnteredCompanyMismatch,
    periodStart: minDate || '2025-04-01',
    periodEnd: maxDate || '2026-03-31',
    currentPeriodLabel: detectedCurrentPeriodLabel || 'FY 2025–26',
    comparativePeriodLabel: detectedComparativePeriodLabel || (hasPriorPeriodFromStatements ? 'FY 2024–25' : undefined),
    availableMonths,
    sourceFiles,
    ledgers,
    ledgerSummaries,
    vouchers: allVouchers,
    deduplicatedVouchers,
    outstandingInvoices,
    stockItems,
    gstInvoices,
    gstr3bSummaries,
    cashFlowItems,
    reportedSubtotals,
    auditExceptions,
    hasCompleteVoucherData,
    hasTrialBalanceData,
    hasBillLevelAgeingData: outstandingInvoices.length > 0,
    hasPriorPeriodData: availableMonths.length >= 2 || hasPriorPeriodFromStatements,
    hasCashFlowStatementData: cashFlowItems.length > 0,
    hasBudgetData,
    hasBalanceSheetStatement,
    hasProfitLossStatement,
  };
}
