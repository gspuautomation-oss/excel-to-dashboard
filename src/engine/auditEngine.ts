import {
  AuditException,
  ConsolidatedAccountingModel,
  formatINR,
  PolicySettings,
} from '../types/accounting';

export interface AccountingHealthCheckResult {
  totalVouchersAnalysed: number;
  totalLedgersAnalysed: number;
  totalExceptions: number;
  highPriorityCount: number;
  mediumPriorityCount: number;
  lowPriorityCount: number;
  unbalancedVoucherCount: number;
  duplicateVoucherCount: number;
  reconciliationDiffCount: number;
  negativeCashDateCount: number;
  exceptions: AuditException[];
  qualityScore: {
    overallScore: number; // 0 to 100
    factors: Array<{
      label: string;
      weight: number;
      deduction: number;
      statusText: string;
    }>;
    disclaimer: string;
  };
}

export function runAccountingHealthCheck(
  model: ConsolidatedAccountingModel,
  policy: PolicySettings
): AccountingHealthCheckResult {
  const exceptions: AuditException[] = [...model.auditExceptions];
  const primaryFilename = model.sourceFiles[0]?.filename || 'Uploaded Dataset';
  const primaryFileId = model.sourceFiles[0]?.id || 'FILE_1';

  // Rule B: Unbalanced Vouchers
  let unbalancedVoucherCount = 0;
  for (const vch of model.deduplicatedVouchers) {
    const diff = Math.abs(vch.totalDebitPaise - vch.totalCreditPaise);
    if (diff > policy.roundingTolerancePaise) {
      unbalancedVoucherCount++;
      exceptions.push({
        id: `EXC_UNBAL_${vch.id}`,
        date: vch.date,
        voucherId: vch.id,
        voucherNumber: vch.voucherNumber || 'Unnumbered',
        voucherType: vch.voucherType,
        ledgerName: vch.entries[0]?.ledgerName || 'Multiple Ledgers',
        partyName: vch.partyName || 'N/A',
        amountPaise: diff,
        category: 'UNBALANCED_VOUCHER',
        severity: 'HIGH',
        explanation: `Total Debits (${formatINR(vch.totalDebitPaise, true)}) do not equal Total Credits (${formatINR(vch.totalCreditPaise, true)}). Difference: ${formatINR(diff, true)}.`,
        recommendedProcedure: 'Inspect the complete voucher entry in TallyPrime and verify whether a ledger line was omitted or truncated during export.',
        sourceFileId: vch.sourceFileId,
        sourceFilename: vch.sourceFilename,
      });
    }
  }

  // Rule C: Negative Cash Balance (reconstructed daily Cash-in-Hand balance when opening balance & complete transactions exist)
  let negativeCashDateCount = 0;
  if (model.hasCompleteVoucherData) {
    const cashLedgers = Array.from(model.ledgerSummaries.values()).filter(
      (s) => s.ledger.classification === 'CashAndEquivalents'
    );

    for (const cashSum of cashLedgers) {
      let runningBalancePaise = cashSum.openingBalancePaise; // +Dr = positive cash
      const dailyNet = new Map<string, { netPaise: number; lastVch: typeof model.deduplicatedVouchers[0] }>();

      for (const vch of model.deduplicatedVouchers) {
        for (const ent of vch.entries) {
          if (ent.normalizedLedgerName === cashSum.ledger.normalizedName) {
            const existing = dailyNet.get(vch.date);
            if (!existing) {
              dailyNet.set(vch.date, { netPaise: ent.amountPaise, lastVch: vch });
            } else {
              existing.netPaise += ent.amountPaise;
              existing.lastVch = vch;
            }
          }
        }
      }

      const sortedDates = Array.from(dailyNet.keys()).sort();
      for (const dt of sortedDates) {
        const day = dailyNet.get(dt)!;
        runningBalancePaise += day.netPaise;
        if (runningBalancePaise < -policy.roundingTolerancePaise) {
          negativeCashDateCount++;
          exceptions.push({
            id: `EXC_NEGCASH_${cashSum.ledger.normalizedName}_${dt}`,
            date: dt,
            voucherId: day.lastVch.id,
            voucherNumber: day.lastVch.voucherNumber || 'N/A',
            voucherType: day.lastVch.voucherType,
            ledgerName: cashSum.ledger.name,
            partyName: day.lastVch.partyName || 'Cash Account',
            amountPaise: Math.abs(runningBalancePaise),
            category: 'NEGATIVE_CASH_BALANCE',
            severity: 'HIGH',
            explanation: `Reconstructed closing cash balance in "${cashSum.ledger.name}" turned negative (${formatINR(runningBalancePaise, true)}) on ${dt}. Physical cash cannot be negative.`,
            recommendedProcedure: 'Verify chronological order of cash receipts and cash payments, or check if opening cash balance was omitted.',
            sourceFileId: day.lastVch.sourceFileId,
            sourceFilename: day.lastVch.sourceFilename,
          });
        }
      }
    }
  }

  // Rule D: Missing Ledger Classification
  for (const [, summary] of model.ledgerSummaries.entries()) {
    if (summary.ledger.classification === 'Unclassified') {
      exceptions.push({
        id: `EXC_UNCLASS_${summary.ledger.normalizedName}`,
        date: model.periodEnd,
        voucherNumber: 'MASTER',
        voucherType: 'Master',
        ledgerName: summary.ledger.name,
        partyName: summary.ledger.parentGroup,
        amountPaise: Math.abs(summary.effectiveClosingPaise),
        category: 'MISSING_LEDGER_CLASSIFICATION',
        severity: Math.abs(summary.effectiveClosingPaise) > policy.materialityThresholdPaise ? 'HIGH' : 'MEDIUM',
        explanation: `Ledger "${summary.ledger.name}" under group "${summary.ledger.parentGroup}" could not be mapped to a standard financial statement schedule.`,
        recommendedProcedure: 'Assign an explicit reporting classification to this ledger or verify its parent group in Tally Masters.',
        sourceFileId: summary.ledger.sourceFileIds[0] || primaryFileId,
        sourceFilename: primaryFilename,
      });
    }
  }

  // Rule E, F, H, J: Voucher-level checks (Unusual Journal Entries, Round-Amount Transactions, Period-End Adjustments, Missing References)
  for (const vch of model.deduplicatedVouchers) {
    const amt = Math.max(vch.totalDebitPaise, vch.totalCreditPaise);

    // Rule E: Unusual Journal Entries exceeding materiality or adjusting Cash/Bank/Sales via Journal
    if (vch.voucherType === 'Journal' && amt >= policy.materialityThresholdPaise) {
      const touchesCashOrSales = vch.entries.some((e) => {
        const led = model.ledgers.get(e.normalizedLedgerName);
        return (
          led?.classification === 'CashAndEquivalents' ||
          led?.classification === 'Sales'
        );
      });
      exceptions.push({
        id: `EXC_JRNL_${vch.id}`,
        date: vch.date,
        voucherId: vch.id,
        voucherNumber: vch.voucherNumber || 'N/A',
        voucherType: vch.voucherType,
        ledgerName: vch.entries.map((e) => e.ledgerName).join(' / '),
        partyName: vch.partyName || 'Journal Entry',
        amountPaise: amt,
        category: 'UNUSUAL_JOURNAL_ENTRY',
        severity: touchesCashOrSales ? 'HIGH' : 'MEDIUM',
        explanation: touchesCashOrSales
          ? `High-value Journal Voucher (${formatINR(amt)}) directly adjusts Cash or Revenue ledgers.`
          : `Material Journal Voucher (${formatINR(amt)}) exceeds configured materiality threshold (${formatINR(policy.materialityThresholdPaise)}).`,
        recommendedProcedure: 'Review journal voucher authorisation, narration, and working papers supporting the adjustment.',
        sourceFileId: vch.sourceFileId,
        sourceFilename: vch.sourceFilename,
      });
    }

    // Rule F: Round-Amount Transactions (multiples of ₹50,000 or configured threshold on non-statutory payments/journals)
    if (
      amt >= policy.roundSumThresholdPaise &&
      amt % 5000000 === 0 &&
      (vch.voucherType === 'Payment' || vch.voucherType === 'Journal' || vch.voucherType === 'Sales')
    ) {
      exceptions.push({
        id: `EXC_ROUND_${vch.id}`,
        date: vch.date,
        voucherId: vch.id,
        voucherNumber: vch.voucherNumber || 'N/A',
        voucherType: vch.voucherType,
        ledgerName: vch.entries[0]?.ledgerName || 'Account',
        partyName: vch.partyName || 'N/A',
        amountPaise: amt,
        category: 'ROUND_AMOUNT_TRANSACTION',
        severity: 'LOW',
        explanation: `Large exact round-sum transaction of ${formatINR(amt)} flagged as an audit review indicator.`,
        recommendedProcedure: 'Inspect underlying invoice or agreement to confirm whether amount is an ad-hoc estimate or contractual milestone.',
        sourceFileId: vch.sourceFileId,
        sourceFilename: vch.sourceFilename,
      });
    }

    // Rule H: Period-End Adjustments (last 5 days of March or reporting period end with high value)
    const isMarchEnd = vch.date.endsWith('-03-28') || vch.date.endsWith('-03-29') || vch.date.endsWith('-03-30') || vch.date.endsWith('-03-31');
    if (
      (isMarchEnd || vch.date === model.periodEnd) &&
      vch.voucherType === 'Journal' &&
      amt >= Math.round(policy.materialityThresholdPaise * 0.5)
    ) {
      exceptions.push({
        id: `EXC_PERIODEND_${vch.id}`,
        date: vch.date,
        voucherId: vch.id,
        voucherNumber: vch.voucherNumber || 'N/A',
        voucherType: vch.voucherType,
        ledgerName: vch.entries.map((e) => e.ledgerName).slice(0, 2).join(', '),
        partyName: vch.partyName || 'Year-End Entry',
        amountPaise: amt,
        category: 'PERIOD_END_ADJUSTMENT',
        severity: 'MEDIUM',
        explanation: `Significant period-end adjustment entry (${formatINR(amt)}) recorded on ${vch.date}.`,
        recommendedProcedure: 'Verify cut-off assertion and supporting schedules for period-end accruals or provisions.',
        sourceFileId: vch.sourceFileId,
        sourceFilename: vch.sourceFilename,
      });
    }

    // Rule J: Missing Transaction References (missing voucher number, or Purchase/Credit Note missing supplier invoice reference, or missing narration on Journal)
    const missingVchNo = !vch.voucherNumber || vch.voucherNumber.trim() === '';
    const missingSupplierRef =
      (vch.voucherType === 'Purchase' || vch.voucherType === 'Debit Note') &&
      !vch.referenceNumber &&
      !vch.isSupplementaryRegister;
    const missingJrnlNarration =
      vch.voucherType === 'Journal' && (!vch.narration || vch.narration.trim().length < 4);

    if (missingVchNo || missingSupplierRef || missingJrnlNarration) {
      exceptions.push({
        id: `EXC_REF_${vch.id}`,
        date: vch.date,
        voucherId: vch.id,
        voucherNumber: vch.voucherNumber || 'MISSING',
        voucherType: vch.voucherType,
        ledgerName: vch.entries[0]?.ledgerName || 'Account',
        partyName: vch.partyName || 'N/A',
        amountPaise: amt,
        category: 'MISSING_TRANSACTION_REFERENCE',
        severity: missingVchNo ? 'MEDIUM' : 'LOW',
        explanation: missingVchNo
          ? 'Voucher is missing a primary voucher number.'
          : missingSupplierRef
          ? `Purchase/Debit Note voucher #${vch.voucherNumber} is missing original supplier invoice reference number.`
          : `Journal voucher #${vch.voucherNumber} lacks an explanatory narration.`,
        recommendedProcedure: 'Ensure all vouchers carry complete reference numbers and audit-ready narrations.',
        sourceFileId: vch.sourceFileId,
        sourceFilename: vch.sourceFilename,
      });
    }
  }

  // Rule G: Suspense Accounts
  for (const [, summary] of model.ledgerSummaries.entries()) {
    if (
      summary.ledger.classification === 'Suspense' &&
      (Math.abs(summary.effectiveClosingPaise) > 0 || summary.periodDebitPaise + summary.periodCreditPaise > 0)
    ) {
      exceptions.push({
        id: `EXC_SUSP_${summary.ledger.normalizedName}`,
        date: model.periodEnd,
        voucherNumber: 'LEDGER-BAL',
        voucherType: 'Suspense',
        ledgerName: summary.ledger.name,
        partyName: 'Suspense Account',
        amountPaise: Math.max(
          Math.abs(summary.effectiveClosingPaise),
          summary.periodDebitPaise,
          summary.periodCreditPaise
        ),
        category: 'SUSPENSE_ACCOUNT_BALANCE',
        severity: Math.abs(summary.effectiveClosingPaise) > 0 ? 'HIGH' : 'MEDIUM',
        explanation:
          Math.abs(summary.effectiveClosingPaise) > 0
            ? `Suspense ledger "${summary.ledger.name}" has an uncleared closing balance of ${formatINR(summary.effectiveClosingPaise, true, true)}.`
            : `Suspense ledger "${summary.ledger.name}" had movements during the period (${formatINR(summary.periodDebitPaise)} Dr / ${formatINR(summary.periodCreditPaise)} Cr) though closing is zero.`,
        recommendedProcedure: 'Investigate all entries posted to Suspense and reclassify to permanent accounting heads before finalising statements.',
        sourceFileId: summary.ledger.sourceFileIds[0] || primaryFileId,
        sourceFilename: primaryFilename,
      });
    }
  }

  // Rule I: Opening and Closing Balance Differences against Trial Balance
  let reconciliationDiffCount = 0;
  for (const [, summary] of model.ledgerSummaries.entries()) {
    if (
      summary.ledger.closingBalanceReportedPaise !== undefined &&
      Math.abs(summary.tbDiscrepancyPaise) > policy.roundingTolerancePaise
    ) {
      reconciliationDiffCount++;
      exceptions.push({
        id: `EXC_TBDIFF_${summary.ledger.normalizedName}`,
        date: model.periodEnd,
        voucherNumber: 'TB-RECON',
        voucherType: 'Reconciliation',
        ledgerName: summary.ledger.name,
        partyName: summary.ledger.parentGroup,
        amountPaise: Math.abs(summary.tbDiscrepancyPaise),
        category: 'OPENING_CLOSING_TB_DIFFERENCE',
        severity: 'HIGH',
        explanation: `Reconstructed closing balance (${formatINR(summary.computedClosingPaise, true, true)}) differs from imported Trial Balance closing (${formatINR(summary.ledger.closingBalanceReportedPaise, true, true)}) by ${formatINR(Math.abs(summary.tbDiscrepancyPaise), true)}.`,
        recommendedProcedure: 'Check if opening balances or partial date-range vouchers were omitted from the transaction XML export.',
        sourceFileId: summary.ledger.sourceFileIds[0] || primaryFileId,
        sourceFilename: primaryFilename,
      });
    }
  }

  const highPriorityCount = exceptions.filter((e) => e.severity === 'HIGH').length;
  const mediumPriorityCount = exceptions.filter((e) => e.severity === 'MEDIUM').length;
  const lowPriorityCount = exceptions.filter((e) => e.severity === 'LOW').length;
  const duplicateVoucherCount = exceptions.filter(
    (e) =>
      e.category === 'EXACT_DUPLICATE_VOUCHER' ||
      e.category === 'POSSIBLE_DUPLICATE_VOUCHER'
  ).length;

  // Explainable Rule-Based Accounting Data Quality Score
  const f1Deduction = Math.min(25, unbalancedVoucherCount * 10);
  const f2Deduction = Math.min(25, reconciliationDiffCount * 8);
  const f3Deduction = Math.min(20, negativeCashDateCount * 10 + (exceptions.some((e) => e.category === 'SUSPENSE_ACCOUNT_BALANCE') ? 10 : 0));
  const f4Deduction = Math.min(15, duplicateVoucherCount * 5);
  const f5Deduction = Math.min(15, mediumPriorityCount * 2 + lowPriorityCount * 1);

  const overallScore = Math.max(
    0,
    100 - (f1Deduction + f2Deduction + f3Deduction + f4Deduction + f5Deduction)
  );

  return {
    totalVouchersAnalysed: model.deduplicatedVouchers.length,
    totalLedgersAnalysed: model.ledgerSummaries.size,
    totalExceptions: exceptions.length,
    highPriorityCount,
    mediumPriorityCount,
    lowPriorityCount,
    unbalancedVoucherCount,
    duplicateVoucherCount,
    reconciliationDiffCount,
    negativeCashDateCount,
    exceptions,
    qualityScore: {
      overallScore,
      factors: [
        {
          label: 'Double-Entry Voucher Equilibrium (Dr = Cr)',
          weight: 25,
          deduction: f1Deduction,
          statusText: unbalancedVoucherCount === 0 ? 'All vouchers balanced' : `${unbalancedVoucherCount} unbalanced voucher(s)`,
        },
        {
          label: 'Trial Balance & Closing Reconciliation',
          weight: 25,
          deduction: f2Deduction,
          statusText: reconciliationDiffCount === 0 ? 'Reconciled within tolerance' : `${reconciliationDiffCount} ledger difference(s)`,
        },
        {
          label: 'Cash-Book Integrity & Suspense Clearance',
          weight: 20,
          deduction: f3Deduction,
          statusText: f3Deduction === 0 ? 'No negative cash or uncleared suspense' : 'Cash/Suspense exceptions detected',
        },
        {
          label: 'Voucher Uniqueness & Deduplication',
          weight: 15,
          deduction: f4Deduction,
          statusText: duplicateVoucherCount === 0 ? 'Zero duplicate vouchers' : `${duplicateVoucherCount} duplicate/suspected voucher(s)`,
        },
        {
          label: 'Reference Completeness & Classification',
          weight: 15,
          deduction: f5Deduction,
          statusText: `${mediumPriorityCount} medium / ${lowPriorityCount} low review indicators`,
        },
      ],
      disclaimer:
        'Rule-based data quality indicator derived solely from uploaded records. This score is an internal analytical aid and does not constitute audit assurance, proof of accounting accuracy, or absence of fraud.',
    },
  };
}
