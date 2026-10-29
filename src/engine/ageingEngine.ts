import {
  ConsolidatedAccountingModel,
  OutstandingInvoice,
  PolicySettings,
} from '../types/accounting';

export type AgeingBucketKey =
  | 'NOT_DUE'
  | 'DAYS_1_30'
  | 'DAYS_31_60'
  | 'DAYS_61_90'
  | 'DAYS_91_180'
  | 'OVER_180';

export interface AgeingBucketSummary {
  key: AgeingBucketKey;
  label: string;
  receivablesPaise: number;
  receivablesCount: number;
  payablesPaise: number;
  payablesCount: number;
}

export interface PartyAgeingSummary {
  partyName: string;
  normalizedPartyName: string;
  category: 'RECEIVABLE' | 'PAYABLE';
  totalOutstandingPaise: number;
  overdueOutstandingPaise: number;
  currentNotDuePaise: number;
  oldestInvoiceDate: string;
  openInvoiceCount: number;
  maxDaysOverdue: number;
  sharePercent: number;
  buckets: Record<AgeingBucketKey, number>;
  invoices: OutstandingInvoice[];
}

export interface AgeingAnalysisResult {
  hasBillLevelData: boolean;
  ageingBasis: 'DUE_DATE' | 'INVOICE_DATE';
  totalTradeReceivablesLedgerPaise: number;
  totalTradePayablesLedgerPaise: number;
  totalBillTrackedReceivablesPaise: number;
  totalBillTrackedPayablesPaise: number;
  unallocatedReceivablesPaise: number;
  unallocatedPayablesPaise: number;
  overdueReceivablesPaise: number;
  overduePayablesPaise: number;
  customersWithOverdueCount: number;
  suppliersWithOverdueCount: number;
  averageCollectionPeriodDays: number | null;
  averagePaymentPeriodDays: number | null;
  bucketSummaries: AgeingBucketSummary[];
  customers: PartyAgeingSummary[];
  suppliers: PartyAgeingSummary[];
}

export const AGEING_BUCKET_LABELS: Record<AgeingBucketKey, string> = {
  NOT_DUE: 'Not Due',
  DAYS_1_30: '1–30 Days',
  DAYS_31_60: '31–60 Days',
  DAYS_61_90: '61–90 Days',
  DAYS_91_180: '91–180 Days',
  OVER_180: 'Over 180 Days',
};

export function classifyAgeingBucket(
  inv: OutstandingInvoice,
  mode: 'DUE_DATE' | 'INVOICE_DATE'
): AgeingBucketKey {
  const days = mode === 'DUE_DATE' ? inv.daysOverdue : inv.invoiceAgeDays;
  if (days <= 0) return 'NOT_DUE';
  if (days <= 30) return 'DAYS_1_30';
  if (days <= 60) return 'DAYS_31_60';
  if (days <= 90) return 'DAYS_61_90';
  if (days <= 180) return 'DAYS_91_180';
  return 'OVER_180';
}

export function computeAgeingAnalysis(
  model: ConsolidatedAccountingModel,
  policy: PolicySettings,
  netCreditSalesPaise: number,
  netCreditPurchasesPaise: number,
  daysInPeriod = 365
): AgeingAnalysisResult {
  const mode = policy.ageingMode;

  let totalTradeReceivablesLedgerPaise = 0;
  let totalTradePayablesLedgerPaise = 0;
  for (const [, sum] of model.ledgerSummaries.entries()) {
    if (sum.ledger.classification === 'TradeReceivables') {
      totalTradeReceivablesLedgerPaise += Math.max(0, sum.effectiveClosingPaise);
    } else if (sum.ledger.classification === 'TradePayables') {
      totalTradePayablesLedgerPaise += Math.max(0, -sum.effectiveClosingPaise);
    }
  }

  const bucketKeys: AgeingBucketKey[] = [
    'NOT_DUE',
    'DAYS_1_30',
    'DAYS_31_60',
    'DAYS_61_90',
    'DAYS_91_180',
    'OVER_180',
  ];

  const bucketMap = new Map<AgeingBucketKey, AgeingBucketSummary>();
  for (const k of bucketKeys) {
    bucketMap.set(k, {
      key: k,
      label: AGEING_BUCKET_LABELS[k],
      receivablesPaise: 0,
      receivablesCount: 0,
      payablesPaise: 0,
      payablesCount: 0,
    });
  }

  const partyMap = new Map<string, PartyAgeingSummary>();
  let totalBillTrackedReceivablesPaise = 0;
  let totalBillTrackedPayablesPaise = 0;
  let overdueReceivablesPaise = 0;
  let overduePayablesPaise = 0;

  for (const inv of model.outstandingInvoices) {
    const bKey = classifyAgeingBucket(inv, mode);
    const bSummary = bucketMap.get(bKey)!;
    const isOverdue = bKey !== 'NOT_DUE';
    const effectiveDays = mode === 'DUE_DATE' ? inv.daysOverdue : inv.invoiceAgeDays;

    if (inv.category === 'RECEIVABLE') {
      bSummary.receivablesPaise += inv.pendingAmountPaise;
      bSummary.receivablesCount += 1;
      totalBillTrackedReceivablesPaise += inv.pendingAmountPaise;
      if (isOverdue) overdueReceivablesPaise += inv.pendingAmountPaise;
    } else {
      bSummary.payablesPaise += inv.pendingAmountPaise;
      bSummary.payablesCount += 1;
      totalBillTrackedPayablesPaise += inv.pendingAmountPaise;
      if (isOverdue) overduePayablesPaise += inv.pendingAmountPaise;
    }

    const pKey = `${inv.category}|${inv.normalizedPartyName}`;
    let party = partyMap.get(pKey);
    if (!party) {
      party = {
        partyName: inv.partyName,
        normalizedPartyName: inv.normalizedPartyName,
        category: inv.category,
        totalOutstandingPaise: 0,
        overdueOutstandingPaise: 0,
        currentNotDuePaise: 0,
        oldestInvoiceDate: inv.billDate,
        openInvoiceCount: 0,
        maxDaysOverdue: 0,
        sharePercent: 0,
        buckets: {
          NOT_DUE: 0,
          DAYS_1_30: 0,
          DAYS_31_60: 0,
          DAYS_61_90: 0,
          DAYS_91_180: 0,
          OVER_180: 0,
        },
        invoices: [],
      };
      partyMap.set(pKey, party);
    }

    party.totalOutstandingPaise += inv.pendingAmountPaise;
    if (isOverdue) {
      party.overdueOutstandingPaise += inv.pendingAmountPaise;
    } else {
      party.currentNotDuePaise += inv.pendingAmountPaise;
    }
    if (inv.billDate < party.oldestInvoiceDate) {
      party.oldestInvoiceDate = inv.billDate;
    }
    party.openInvoiceCount += 1;
    if (effectiveDays > party.maxDaysOverdue) {
      party.maxDaysOverdue = effectiveDays;
    }
    party.buckets[bKey] += inv.pendingAmountPaise;
    party.invoices.push(inv);
  }

  const customers: PartyAgeingSummary[] = [];
  const suppliers: PartyAgeingSummary[] = [];

  for (const party of partyMap.values()) {
    if (party.category === 'RECEIVABLE') {
      party.sharePercent =
        totalBillTrackedReceivablesPaise > 0
          ? (party.totalOutstandingPaise / totalBillTrackedReceivablesPaise) * 100
          : 0;
      customers.push(party);
    } else {
      party.sharePercent =
        totalBillTrackedPayablesPaise > 0
          ? (party.totalOutstandingPaise / totalBillTrackedPayablesPaise) * 100
          : 0;
      suppliers.push(party);
    }
  }

  customers.sort((a, b) => b.totalOutstandingPaise - a.totalOutstandingPaise);
  suppliers.sort((a, b) => b.totalOutstandingPaise - a.totalOutstandingPaise);

  const customersWithOverdueCount = customers.filter((c) => c.overdueOutstandingPaise > 0).length;
  const suppliersWithOverdueCount = suppliers.filter((s) => s.overdueOutstandingPaise > 0).length;

  const effectiveRec =
    totalTradeReceivablesLedgerPaise > 0
      ? totalTradeReceivablesLedgerPaise
      : totalBillTrackedReceivablesPaise;
  const effectivePay =
    totalTradePayablesLedgerPaise > 0
      ? totalTradePayablesLedgerPaise
      : totalBillTrackedPayablesPaise;

  const averageCollectionPeriodDays =
    netCreditSalesPaise > 0 && effectiveRec > 0
      ? Math.round((effectiveRec / netCreditSalesPaise) * daysInPeriod)
      : null;

  const averagePaymentPeriodDays =
    netCreditPurchasesPaise > 0 && effectivePay > 0
      ? Math.round((effectivePay / netCreditPurchasesPaise) * daysInPeriod)
      : null;

  return {
    hasBillLevelData: model.outstandingInvoices.length > 0,
    ageingBasis: mode,
    totalTradeReceivablesLedgerPaise,
    totalTradePayablesLedgerPaise,
    totalBillTrackedReceivablesPaise,
    totalBillTrackedPayablesPaise,
    unallocatedReceivablesPaise: Math.max(
      0,
      totalTradeReceivablesLedgerPaise - totalBillTrackedReceivablesPaise
    ),
    unallocatedPayablesPaise: Math.max(
      0,
      totalTradePayablesLedgerPaise - totalBillTrackedPayablesPaise
    ),
    overdueReceivablesPaise,
    overduePayablesPaise,
    customersWithOverdueCount,
    suppliersWithOverdueCount,
    averageCollectionPeriodDays,
    averagePaymentPeriodDays,
    bucketSummaries: bucketKeys.map((k) => bucketMap.get(k)!),
    customers,
    suppliers,
  };
}
