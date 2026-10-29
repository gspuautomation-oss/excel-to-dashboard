import React, { useMemo, useState } from 'react';
import * as XLSX from 'xlsx';
import {
  ArrowLeft,
  ChevronRight,
  Download,
  FileSpreadsheet,
  Search,
  X,
} from 'lucide-react';
import {
  AccountClassification,
  ComputedLedgerSummary,
  ConsolidatedAccountingModel,
  formatINR,
  fromPaise,
  Voucher,
} from '../types/accounting';
import { CLASSIFICATION_LABELS } from '../engine/ledgerMapping';

export interface DrillDownTarget {
  title: string;
  formulaExplanation?: string;
  classifications?: AccountClassification[];
  specificLedgerNormalizedName?: string;
  specificVoucherId?: string;
  currentFormatted?: string;
  previousFormatted?: string;
  sourceCellRef?: string;
  priorCellRef?: string;
  statusIndicator?: string;
  differenceFormatted?: string;
}

interface DrillDownDrawerProps {
  target: DrillDownTarget | null;
  model: ConsolidatedAccountingModel;
  onClose: () => void;
}

export const DrillDownDrawer: React.FC<DrillDownDrawerProps> = ({
  target,
  model,
  onClose,
}) => {
  const [selectedCategory, setSelectedCategory] = useState<AccountClassification | null>(null);
  const [selectedGroup, setSelectedGroup] = useState<string | null>(null);
  const [selectedLedgerKey, setSelectedLedgerKey] = useState<string | null>(null);
  const [selectedVoucherId, setSelectedVoucherId] = useState<string | null>(null);

  // Interactive controls (Module 4.5)
  const [searchTerm, setSearchTerm] = useState('');
  const [voucherTypeFilter, setVoucherTypeFilter] = useState<string>('ALL');
  const [dateFrom, setDateFrom] = useState('');
  const [dateTo, setDateTo] = useState('');
  const [minAmountRupees, setMinAmountRupees] = useState('');
  const [sortBy, setSortBy] = useState<'DATE_DESC' | 'DATE_ASC' | 'AMOUNT_DESC' | 'AMOUNT_ASC'>('DATE_DESC');

  // Synchronize direct ledger or voucher drill-down
  const activeVoucherId = selectedVoucherId || target?.specificVoucherId || null;
  const activeLedgerKey = selectedLedgerKey || target?.specificLedgerNormalizedName || null;

  const matchingLedgers = useMemo(() => {
    if (!target) return [];
    const list: ComputedLedgerSummary[] = [];
    for (const [, sum] of model.ledgerSummaries.entries()) {
      if (
        target.classifications &&
        target.classifications.length > 0 &&
        !target.classifications.includes(sum.ledger.classification)
      ) {
        continue;
      }
      if (selectedCategory && sum.ledger.classification !== selectedCategory) {
        continue;
      }
      if (selectedGroup && sum.ledger.parentGroup !== selectedGroup) {
        continue;
      }
      list.push(sum);
    }
    return list.sort(
      (a, b) => Math.abs(b.effectiveClosingPaise) - Math.abs(a.effectiveClosingPaise)
    );
  }, [target, model, selectedCategory, selectedGroup]);

  const activeLedgerSummary = useMemo(() => {
    if (!activeLedgerKey) return null;
    return model.ledgerSummaries.get(activeLedgerKey) || null;
  }, [activeLedgerKey, model]);

  const activeVoucher = useMemo(() => {
    if (!activeVoucherId) return null;
    return (
      model.deduplicatedVouchers.find((v) => v.id === activeVoucherId) ||
      model.vouchers.find((v) => v.id === activeVoucherId) ||
      null
    );
  }, [activeVoucherId, model]);

  const filteredVouchers = useMemo(() => {
    let baseVouchers: Voucher[] = [];
    if (activeLedgerSummary) {
      const idSet = new Set(activeLedgerSummary.voucherIds);
      baseVouchers = model.deduplicatedVouchers.filter((v) => idSet.has(v.id));
    } else if (target?.classifications && target.classifications.length > 0) {
      const allowedLedgerNames = new Set(matchingLedgers.map((l) => l.ledger.normalizedName));
      baseVouchers = model.deduplicatedVouchers.filter((v) =>
        v.entries.some((e) => allowedLedgerNames.has(e.normalizedLedgerName))
      );
    } else {
      baseVouchers = model.deduplicatedVouchers;
    }

    const minPaise = minAmountRupees ? Number(minAmountRupees) * 100 : 0;
    const q = searchTerm.trim().toLowerCase();

    const filtered = baseVouchers.filter((v) => {
      if (voucherTypeFilter !== 'ALL' && v.voucherType !== voucherTypeFilter) return false;
      if (dateFrom && v.date < dateFrom) return false;
      if (dateTo && v.date > dateTo) return false;
      if (minPaise > 0 && Math.max(v.totalDebitPaise, v.totalCreditPaise) < minPaise) return false;
      if (q) {
        const hay = `${v.voucherNumber} ${v.partyName || ''} ${v.narration || ''} ${v.entries
          .map((e) => e.ledgerName)
          .join(' ')}`.toLowerCase();
        if (!hay.includes(q)) return false;
      }
      return true;
    });

    return filtered.sort((a, b) => {
      if (sortBy === 'DATE_DESC') return b.date.localeCompare(a.date);
      if (sortBy === 'DATE_ASC') return a.date.localeCompare(b.date);
      if (sortBy === 'AMOUNT_DESC') return b.totalDebitPaise - a.totalDebitPaise;
      return a.totalDebitPaise - b.totalDebitPaise;
    });
  }, [
    activeLedgerSummary,
    target,
    matchingLedgers,
    model,
    voucherTypeFilter,
    dateFrom,
    dateTo,
    minAmountRupees,
    searchTerm,
    sortBy,
  ]);

  // Group summary for hierarchy level: ACCOUNT CATEGORY -> LEDGER GROUP
  const groupBreakdowns = useMemo(() => {
    const map = new Map<
      string,
      { groupName: string; classification: AccountClassification; count: number; closingPaise: number; debitPaise: number; creditPaise: number }
    >();
    for (const l of matchingLedgers) {
      const gKey = `${l.ledger.classification}|${l.ledger.parentGroup}`;
      const ex = map.get(gKey);
      if (!ex) {
        map.set(gKey, {
          groupName: l.ledger.parentGroup,
          classification: l.ledger.classification,
          count: 1,
          closingPaise: l.effectiveClosingPaise,
          debitPaise: l.periodDebitPaise,
          creditPaise: l.periodCreditPaise,
        });
      } else {
        ex.count += 1;
        ex.closingPaise += l.effectiveClosingPaise;
        ex.debitPaise += l.periodDebitPaise;
        ex.creditPaise += l.periodCreditPaise;
      }
    }
    return Array.from(map.values());
  }, [matchingLedgers]);

  if (!target) return null;

  const handleExportSelectedVouchers = () => {
    const rows = [
      [
        'Date',
        'Voucher Number',
        'Voucher Type',
        'Party Name',
        'Ledgers Involved',
        'Total Debit (INR)',
        'Total Credit (INR)',
        'Narration',
        'Source File',
        'Source Record ID',
      ],
      ...filteredVouchers.map((v) => [
        v.date,
        v.voucherNumber,
        v.voucherType,
        v.partyName || '',
        v.entries.map((e) => `${e.ledgerName} (${formatINR(e.amountPaise, true, true)})`).join(' | '),
        fromPaise(v.totalDebitPaise),
        fromPaise(v.totalCreditPaise),
        v.narration || '',
        v.sourceFilename,
        v.sourceRecordId,
      ]),
    ];
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet(rows), 'DrillDown_Vouchers');
    XLSX.writeFile(wb, `GSPU_DrillDown_${target.title.replace(/[^a-zA-Z0-9]/g, '_')}.xlsx`);
  };

  const resetStepBack = () => {
    if (selectedVoucherId) {
      setSelectedVoucherId(null);
    } else if (selectedLedgerKey) {
      setSelectedLedgerKey(null);
    } else if (selectedGroup) {
      setSelectedGroup(null);
    } else if (selectedCategory) {
      setSelectedCategory(null);
    } else {
      onClose();
    }
  };

  return (
    <div className="fixed inset-0 z-50 flex justify-end bg-slate-900/50 backdrop-blur-[1px]">
      <div className="w-full max-w-5xl bg-white h-full flex flex-col shadow-2xl border-l border-slate-200 overflow-hidden">
        {/* Drawer Header */}
        <div className="px-6 py-4 bg-[#0B2540] text-white flex items-center justify-between gap-4 shrink-0">
          <div className="flex items-center gap-3 min-w-0">
            <button
              onClick={resetStepBack}
              className="p-1.5 rounded-lg bg-white/10 hover:bg-white/20 text-white transition-colors"
              title="Back"
            >
              <ArrowLeft className="w-4 h-4" />
            </button>
            <div className="min-w-0">
              <div className="flex items-center gap-1.5 text-xs text-slate-300 flex-wrap">
                <span>KPI Drill-Down</span>
                <ChevronRight className="w-3.5 h-3.5" />
                <button
                  onClick={() => {
                    setSelectedCategory(null);
                    setSelectedGroup(null);
                    setSelectedLedgerKey(null);
                    setSelectedVoucherId(null);
                  }}
                  className="hover:underline text-white font-medium"
                >
                  {target.title}
                </button>
                {selectedGroup && (
                  <>
                    <ChevronRight className="w-3.5 h-3.5" />
                    <button
                      onClick={() => {
                        setSelectedLedgerKey(null);
                        setSelectedVoucherId(null);
                      }}
                      className="hover:underline text-teal-300"
                    >
                      {selectedGroup}
                    </button>
                  </>
                )}
                {activeLedgerSummary && (
                  <>
                    <ChevronRight className="w-3.5 h-3.5" />
                    <button
                      onClick={() => setSelectedVoucherId(null)}
                      className="hover:underline text-teal-300"
                    >
                      {activeLedgerSummary.ledger.name}
                    </button>
                  </>
                )}
                {activeVoucher && (
                  <>
                    <ChevronRight className="w-3.5 h-3.5" />
                    <span className="text-amber-300 font-mono">
                      Voucher #{activeVoucher.voucherNumber || activeVoucher.id}
                    </span>
                  </>
                )}
              </div>
              <h2 className="text-lg font-bold truncate mt-0.5">
                {activeVoucher
                  ? `Complete Accounting Voucher — ${activeVoucher.voucherType} #${activeVoucher.voucherNumber}`
                  : activeLedgerSummary
                  ? `Ledger Account — ${activeLedgerSummary.ledger.name}`
                  : target.title}
              </h2>
            </div>
          </div>

          <div className="flex items-center gap-2 shrink-0">
            <button
              onClick={handleExportSelectedVouchers}
              className="px-3 py-1.5 text-xs font-medium bg-[#0875BC] hover:bg-[#065d96] text-white rounded-lg flex items-center gap-1.5 transition-colors whitespace-nowrap"
            >
              <Download className="w-3.5 h-3.5" />
              Export Transactions
            </button>
            <button
              onClick={onClose}
              className="p-1.5 rounded-lg bg-white/10 hover:bg-white/20 text-white transition-colors"
              aria-label="Close drill-down"
            >
              <X className="w-5 h-5" />
            </button>
          </div>
        </div>

        {/* Audit Trail Formula Banner (Module 4.6) */}
        {target.formulaExplanation && (
          <div className="px-6 py-2.5 bg-slate-50 border-b border-slate-200 flex items-center justify-between gap-4 text-xs">
            <div className="text-slate-700">
              <span className="font-semibold text-[#0B2540]">Audit Trail Formula: </span>
              <span>{target.formulaExplanation}</span>
            </div>
            <span className="text-slate-500 font-mono whitespace-nowrap">
              {matchingLedgers.length} Ledger(s) · {filteredVouchers.length} Voucher(s)
            </span>
          </div>
        )}

        {/* Drawer Body */}
        <div className="flex-1 overflow-y-auto p-6 space-y-6">
          {/* SOURCE PROVENANCE & RECONCILIATION VERIFICATION CARD (Part 7) */}
          {!activeVoucher && (
            <div className="p-4 rounded-xl bg-slate-50 border border-slate-200 space-y-3">
              <div className="flex flex-wrap items-center justify-between gap-2">
                <div>
                  <div className="text-[10px] font-bold uppercase tracking-wider text-[#0867F5]">
                    Source Provenance &amp; Audit Verification Trace
                  </div>
                  <div className="text-sm font-bold text-[#0B2540]">{target.title}</div>
                </div>
                {target.statusIndicator && (
                  <span
                    className={`px-2.5 py-0.5 rounded-full text-[11px] font-bold border ${
                      target.statusIndicator === 'Reconciled' ||
                      target.statusIndicator === 'RECONCILED'
                        ? 'bg-emerald-50 text-emerald-700 border-emerald-200'
                        : target.statusIndicator === 'Insufficient Data' ||
                          target.statusIndicator === 'INSUFFICIENT DATA'
                        ? 'bg-slate-100 text-slate-600 border-slate-200'
                        : 'bg-amber-50 text-amber-800 border-amber-200'
                    }`}
                  >
                    {target.statusIndicator}
                  </span>
                )}
              </div>
              <div className="grid grid-cols-2 sm:grid-cols-4 gap-3 text-xs">
                <div>
                  <div className="text-[11px] text-slate-500">
                    Current ({model.currentPeriodLabel || 'FY 2025-26'})
                  </div>
                  <div className="font-mono font-bold text-slate-900 mt-0.5">
                    {target.currentFormatted ||
                      formatINR(
                        matchingLedgers.reduce(
                          (acc, l) => acc + Math.abs(l.effectiveClosingPaise),
                          0
                        )
                      )}
                  </div>
                </div>
                <div>
                  <div className="text-[11px] text-slate-500">
                    Comparative ({model.comparativePeriodLabel || 'FY 2024-25'})
                  </div>
                  <div className="font-mono font-semibold text-slate-700 mt-0.5">
                    {target.previousFormatted || '—'}
                  </div>
                </div>
                <div>
                  <div className="text-[11px] text-slate-500">Source File(s)</div>
                  <div
                    className="font-medium text-slate-800 mt-0.5 truncate"
                    title={model.sourceFiles.map((f) => f.filename).join(', ')}
                  >
                    {model.sourceFiles.map((f) => f.filename).join(', ') || 'Uploaded Workbook'}
                  </div>
                </div>
                <div>
                  <div className="text-[11px] text-slate-500">Worksheet &amp; Cell Address</div>
                  <div className="font-mono text-slate-700 mt-0.5 truncate">
                    {target.sourceCellRef ||
                      matchingLedgers
                        .map((l) => l.ledger.cellReference)
                        .filter(Boolean)
                        .slice(0, 3)
                        .join(', ') ||
                      '—'}
                  </div>
                </div>
              </div>
            </div>
          )}

          {/* LEVEL 6: COMPLETE VOUCHER DETAILS (Module 4.4) */}
          {activeVoucher ? (
            <div className="space-y-6">
              <div className="grid grid-cols-2 sm:grid-cols-4 gap-4 p-4 bg-slate-50 rounded-lg border border-slate-200">
                <div>
                  <div className="text-xs text-slate-500">Voucher Date</div>
                  <div className="text-sm font-semibold font-mono text-[#0B2540] mt-0.5">
                    {activeVoucher.date}
                  </div>
                </div>
                <div>
                  <div className="text-xs text-slate-500">Voucher Number & Type</div>
                  <div className="text-sm font-semibold font-mono text-[#0B2540] mt-0.5">
                    {activeVoucher.voucherNumber || 'Unnumbered'} · {activeVoucher.voucherType}
                  </div>
                </div>
                <div>
                  <div className="text-xs text-slate-500">Counterparty / GSTIN</div>
                  <div className="text-sm font-semibold text-[#0B2540] mt-0.5 truncate">
                    {activeVoucher.partyName || '—'}
                    {activeVoucher.partyGstin ? ` (${activeVoucher.partyGstin})` : ''}
                  </div>
                </div>
                <div>
                  <div className="text-xs text-slate-500">Source Provenance</div>
                  <div className="text-xs font-mono text-slate-700 mt-0.5 truncate" title={activeVoucher.sourceRecordId}>
                    {activeVoucher.sourceFilename} · ID: {activeVoucher.sourceRecordId}
                  </div>
                </div>
              </div>

              {activeVoucher.narration && (
                <div className="p-3.5 bg-slate-50 rounded-lg border border-slate-200 text-xs text-slate-700">
                  <span className="font-semibold text-[#0B2540]">Voucher Narration: </span>
                  {activeVoucher.narration}
                </div>
              )}

              {/* Complete Double-Entry Accounting Lines */}
              <div>
                <h3 className="text-sm font-semibold text-[#0B2540] mb-2">
                  Complete Double-Entry Ledger Lines
                </h3>
                <div className="border border-slate-200 rounded-lg overflow-hidden">
                  <table className="w-full text-left border-collapse text-xs">
                    <thead>
                      <tr className="bg-slate-100 text-slate-700 border-b border-slate-200">
                        <th className="py-2.5 px-4 font-semibold">Dr / Cr</th>
                        <th className="py-2.5 px-4 font-semibold">Ledger Account</th>
                        <th className="py-2.5 px-4 font-semibold">Classification</th>
                        <th className="py-2.5 px-4 font-semibold text-right">Debit (₹)</th>
                        <th className="py-2.5 px-4 font-semibold text-right">Credit (₹)</th>
                      </tr>
                    </thead>
                    <tbody className="divide-y divide-slate-200">
                      {activeVoucher.entries.map((ent, idx) => {
                        const led = model.ledgers.get(ent.normalizedLedgerName);
                        const isDr = ent.amountPaise > 0;
                        return (
                          <tr key={idx} className="hover:bg-slate-50">
                            <td className="py-2.5 px-4 font-mono font-semibold">
                              {isDr ? 'Dr' : 'Cr'}
                            </td>
                            <td className="py-2.5 px-4">
                              <button
                                onClick={() => {
                                  setSelectedLedgerKey(ent.normalizedLedgerName);
                                  setSelectedVoucherId(null);
                                }}
                                className="font-medium text-[#0875BC] hover:underline text-left"
                              >
                                {ent.ledgerName}
                              </button>
                              {ent.billAllocations.length > 0 && (
                                <div className="mt-1 space-y-0.5 text-[11px] text-slate-500 font-mono">
                                  {ent.billAllocations.map((ba, bIdx) => (
                                    <div key={bIdx}>
                                      Bill Ref: {ba.billName} ({ba.billType}) ·{' '}
                                      {formatINR(Math.abs(ba.amountPaise), true)}
                                      {ba.dueDate ? ` · Due: ${ba.dueDate}` : ''}
                                    </div>
                                  ))}
                                </div>
                              )}
                            </td>
                            <td className="py-2.5 px-4 text-slate-600">
                              {led ? CLASSIFICATION_LABELS[led.classification] : 'Unclassified'}
                            </td>
                            <td className="py-2.5 px-4 text-right font-mono tabular-nums font-medium text-slate-900">
                              {isDr ? formatINR(ent.amountPaise, true) : '—'}
                            </td>
                            <td className="py-2.5 px-4 text-right font-mono tabular-nums font-medium text-slate-900">
                              {!isDr ? formatINR(Math.abs(ent.amountPaise), true) : '—'}
                            </td>
                          </tr>
                        );
                      })}
                    </tbody>
                    <tfoot>
                      <tr className="bg-slate-50 font-semibold border-t border-slate-300 text-xs">
                        <td colSpan={3} className="py-2.5 px-4 text-[#0B2540]">
                          Voucher Total ({activeVoucher.isBalanced ? 'Balanced' : 'Unbalanced Difference'})
                        </td>
                        <td className="py-2.5 px-4 text-right font-mono tabular-nums text-[#0B2540]">
                          {formatINR(activeVoucher.totalDebitPaise, true)}
                        </td>
                        <td className="py-2.5 px-4 text-right font-mono tabular-nums text-[#0B2540]">
                          {formatINR(activeVoucher.totalCreditPaise, true)}
                        </td>
                      </tr>
                    </tfoot>
                  </table>
                </div>
              </div>

              {/* Inventory Movements if present */}
              {activeVoucher.inventoryMovements.length > 0 && (
                <div>
                  <h3 className="text-sm font-semibold text-[#0B2540] mb-2">
                    Inventory Line Items
                  </h3>
                  <div className="border border-slate-200 rounded-lg overflow-hidden">
                    <table className="w-full text-left border-collapse text-xs">
                      <thead>
                        <tr className="bg-slate-100 text-slate-700 border-b border-slate-200">
                          <th className="py-2 px-4 font-semibold">Stock Item</th>
                          <th className="py-2 px-4 font-semibold text-right">Quantity</th>
                          <th className="py-2 px-4 font-semibold text-right">Rate</th>
                          <th className="py-2 px-4 font-semibold text-right">Amount</th>
                        </tr>
                      </thead>
                      <tbody className="divide-y divide-slate-200">
                        {activeVoucher.inventoryMovements.map((im, i) => (
                          <tr key={i}>
                            <td className="py-2 px-4 font-medium">{im.stockItemName}</td>
                            <td className="py-2 px-4 text-right font-mono">
                              {im.quantity} {im.unit}
                            </td>
                            <td className="py-2 px-4 text-right font-mono">
                              {formatINR(im.ratePaise, true)}
                            </td>
                            <td className="py-2 px-4 text-right font-mono">
                              {formatINR(Math.abs(im.amountPaise), true)}
                            </td>
                          </tr>
                        ))}
                      </tbody>
                    </table>
                  </div>
                </div>
              )}
            </div>
          ) : (
            <>
              {/* LEVEL 4: LEDGER DETAIL SUMMARY (Module 4.3) */}
              {activeLedgerSummary ? (
                <div className="space-y-4">
                  <div className="grid grid-cols-2 sm:grid-cols-5 gap-3 p-4 bg-slate-50 rounded-lg border border-slate-200">
                    <div>
                      <div className="text-xs text-slate-500">Parent Group</div>
                      <div className="text-sm font-semibold text-[#0B2540] mt-0.5">
                        {activeLedgerSummary.ledger.parentGroup}
                      </div>
                    </div>
                    <div>
                      <div className="text-xs text-slate-500">Opening Balance</div>
                      <div className="text-sm font-semibold font-mono tabular-nums text-[#0B2540] mt-0.5">
                        {formatINR(activeLedgerSummary.openingBalancePaise, false, true)}
                      </div>
                    </div>
                    <div>
                      <div className="text-xs text-slate-500">Total Debits</div>
                      <div className="text-sm font-semibold font-mono tabular-nums text-[#0B2540] mt-0.5">
                        {formatINR(activeLedgerSummary.periodDebitPaise)}
                      </div>
                    </div>
                    <div>
                      <div className="text-xs text-slate-500">Total Credits</div>
                      <div className="text-sm font-semibold font-mono tabular-nums text-[#0B2540] mt-0.5">
                        {formatINR(activeLedgerSummary.periodCreditPaise)}
                      </div>
                    </div>
                    <div>
                      <div className="text-xs text-slate-500">Closing Balance</div>
                      <div className="text-sm font-bold font-mono tabular-nums text-[#0875BC] mt-0.5">
                        {formatINR(activeLedgerSummary.effectiveClosingPaise, false, true)}
                      </div>
                    </div>
                  </div>

                  {/* Monthly Movements Strip */}
                  {Object.keys(activeLedgerSummary.monthlyNetPaise).length > 0 && (
                    <div>
                      <h4 className="text-xs font-semibold text-slate-600 mb-2">
                        Monthly Ledger Movements ({activeLedgerSummary.transactionCount} Transactions)
                      </h4>
                      <div className="grid grid-cols-2 sm:grid-cols-4 md:grid-cols-6 gap-2">
                        {Object.entries(activeLedgerSummary.monthlyNetPaise)
                          .sort(([a], [b]) => a.localeCompare(b))
                          .map(([ym, mv]) => (
                            <div
                              key={ym}
                              className="p-2.5 rounded border border-slate-200 bg-white text-xs"
                            >
                              <div className="font-mono font-semibold text-slate-700">{ym}</div>
                              <div className="font-mono tabular-nums text-[11px] text-slate-500 mt-1">
                                Dr: {formatINR(mv.debitPaise)}
                              </div>
                              <div className="font-mono tabular-nums text-[11px] text-slate-500">
                                Cr: {formatINR(mv.creditPaise)}
                              </div>
                            </div>
                          ))}
                      </div>
                    </div>
                  )}
                </div>
              ) : (
                /* LEVEL 2 & 3: ACCOUNT CATEGORY -> LEDGER GROUP -> LEDGERS */
                <div className="space-y-5">
                  {groupBreakdowns.length > 1 && !selectedGroup && (
                    <div>
                      <h3 className="text-xs font-semibold text-slate-500 mb-2">
                        Ledger Groups in Selection (Click Group to Filter)
                      </h3>
                      <div className="grid grid-cols-1 sm:grid-cols-3 gap-3">
                        {groupBreakdowns.map((g) => (
                          <button
                            key={`${g.classification}_${g.groupName}`}
                            onClick={() => setSelectedGroup(g.groupName)}
                            className="p-3 rounded-lg border border-slate-200 hover:border-[#0875BC] bg-slate-50/70 text-left transition-colors"
                          >
                            <div className="text-xs font-semibold text-[#0B2540] truncate">
                              {g.groupName}
                            </div>
                            <div className="text-[11px] text-slate-500 mt-0.5">
                              {CLASSIFICATION_LABELS[g.classification]} · {g.count} Ledger(s)
                            </div>
                            <div className="text-sm font-mono font-bold text-[#0875BC] mt-1.5 tabular-nums">
                              {formatINR(g.closingPaise, false, true)}
                            </div>
                          </button>
                        ))}
                      </div>
                    </div>
                  )}

                  <div>
                    <h3 className="text-sm font-semibold text-[#0B2540] mb-2">
                      Underlying Ledger Accounts (Click any ledger to inspect monthly movements & vouchers)
                    </h3>
                    <div className="border border-slate-200 rounded-lg overflow-hidden">
                      <table className="w-full text-left border-collapse text-xs">
                        <thead>
                          <tr className="bg-slate-100 text-slate-700 border-b border-slate-200">
                            <th className="py-2.5 px-3 font-semibold">Original Row / Ledger Label</th>
                            <th className="py-2.5 px-3 font-semibold">Schedule III / Group</th>
                            <th className="py-2.5 px-3 font-semibold">Worksheet &amp; Cell</th>
                            <th className="py-2.5 px-3 font-semibold text-right">
                              Current ({model.currentPeriodLabel || 'Current'})
                            </th>
                            <th className="py-2.5 px-3 font-semibold text-right">
                              Comparative ({model.comparativePeriodLabel || 'Prior'})
                            </th>
                            <th className="py-2.5 px-3 font-semibold text-right">Vouchers</th>
                          </tr>
                        </thead>
                        <tbody className="divide-y divide-slate-200">
                          {matchingLedgers.map((l) => (
                            <tr
                              key={l.ledger.normalizedName}
                              onClick={() => setSelectedLedgerKey(l.ledger.normalizedName)}
                              className="hover:bg-blue-50/60 cursor-pointer transition-colors"
                            >
                              <td className="py-2.5 px-3 font-medium text-[#0875BC]">
                                {l.ledger.name}
                              </td>
                              <td className="py-2.5 px-3 text-slate-600">{l.ledger.parentGroup}</td>
                              <td className="py-2.5 px-3 font-mono text-[11px] text-slate-500">
                                {l.ledger.cellReference || l.ledger.worksheetName || '—'}
                                {l.ledger.priorCellReference
                                  ? ` / ${l.ledger.priorCellReference.split('!').pop()}`
                                  : ''}
                              </td>
                              <td className="py-2.5 px-3 text-right font-mono tabular-nums font-semibold text-[#0B2540]">
                                {formatINR(l.effectiveClosingPaise, false, true)}
                              </td>
                              <td className="py-2.5 px-3 text-right font-mono tabular-nums text-slate-600">
                                {l.ledger.previousPeriodPaise !== undefined
                                  ? formatINR(l.ledger.previousPeriodPaise)
                                  : '—'}
                              </td>
                              <td className="py-2.5 px-3 text-right font-mono">
                                {l.transactionCount}
                              </td>
                            </tr>
                          ))}
                          {matchingLedgers.length === 0 && (
                            <tr>
                              <td colSpan={6} className="py-6 text-center text-slate-500">
                                No matching ledgers found for this selection.
                              </td>
                            </tr>
                          )}
                        </tbody>
                      </table>
                    </div>
                  </div>
                </div>
              )}

              {/* LEVEL 5: INTERACTIVE VOUCHER LIST & FILTERS (Module 4.5) */}
              <div className="space-y-3 pt-2">
                <div className="flex flex-wrap items-center justify-between gap-3">
                  <h3 className="text-sm font-semibold text-[#0B2540]">
                    Supporting Accounting Vouchers ({filteredVouchers.length})
                  </h3>
                  <div className="flex flex-wrap items-center gap-2">
                    <div className="relative">
                      <Search className="w-3.5 h-3.5 text-slate-400 absolute left-2.5 top-1/2 -translate-y-1/2" />
                      <input
                        type="text"
                        value={searchTerm}
                        onChange={(e) => setSearchTerm(e.target.value)}
                        placeholder="Search vch no, party, narration..."
                        className="pl-8 pr-3 py-1.5 text-xs border border-slate-300 rounded-lg focus:outline-none focus:border-[#0875BC]"
                      />
                    </div>
                    <select
                      value={voucherTypeFilter}
                      onChange={(e) => setVoucherTypeFilter(e.target.value)}
                      className="px-2.5 py-1.5 text-xs border border-slate-300 rounded-lg bg-white"
                    >
                      <option value="ALL">All Voucher Types</option>
                      <option value="Sales">Sales</option>
                      <option value="Purchase">Purchase</option>
                      <option value="Receipt">Receipt</option>
                      <option value="Payment">Payment</option>
                      <option value="Journal">Journal</option>
                      <option value="Credit Note">Credit Note</option>
                      <option value="Debit Note">Debit Note</option>
                    </select>
                    <input
                      type="date"
                      value={dateFrom}
                      onChange={(e) => setDateFrom(e.target.value)}
                      className="px-2 py-1.5 text-xs border border-slate-300 rounded-lg"
                      title="From Date"
                    />
                    <input
                      type="date"
                      value={dateTo}
                      onChange={(e) => setDateTo(e.target.value)}
                      className="px-2 py-1.5 text-xs border border-slate-300 rounded-lg"
                      title="To Date"
                    />
                    <input
                      type="number"
                      value={minAmountRupees}
                      onChange={(e) => setMinAmountRupees(e.target.value)}
                      placeholder="Min ₹"
                      className="w-24 px-2.5 py-1.5 text-xs border border-slate-300 rounded-lg"
                    />
                    <select
                      value={sortBy}
                      onChange={(e) => setSortBy(e.target.value as any)}
                      className="px-2.5 py-1.5 text-xs border border-slate-300 rounded-lg bg-white"
                    >
                      <option value="DATE_DESC">Date (Newest)</option>
                      <option value="DATE_ASC">Date (Oldest)</option>
                      <option value="AMOUNT_DESC">Amount (High → Low)</option>
                      <option value="AMOUNT_ASC">Amount (Low → High)</option>
                    </select>
                  </div>
                </div>

                <div className="border border-slate-200 rounded-lg overflow-hidden">
                  <table className="w-full text-left border-collapse text-xs">
                    <thead>
                      <tr className="bg-slate-100 text-slate-700 border-b border-slate-200">
                        <th className="py-2.5 px-3 font-semibold">Date</th>
                        <th className="py-2.5 px-3 font-semibold">Voucher No</th>
                        <th className="py-2.5 px-3 font-semibold">Type</th>
                        <th className="py-2.5 px-3 font-semibold">Party / Particulars</th>
                        <th className="py-2.5 px-3 font-semibold text-right">Amount (₹)</th>
                        <th className="py-2.5 px-3 font-semibold">Source File</th>
                        <th className="py-2.5 px-3 font-semibold text-right">Action</th>
                      </tr>
                    </thead>
                    <tbody className="divide-y divide-slate-200">
                      {filteredVouchers.map((v) => (
                        <tr
                          key={v.id}
                          onClick={() => setSelectedVoucherId(v.id)}
                          className="hover:bg-blue-50/60 cursor-pointer transition-colors"
                        >
                          <td className="py-2.5 px-3 font-mono whitespace-nowrap">{v.date}</td>
                          <td className="py-2.5 px-3 font-mono font-medium text-[#0875BC]">
                            {v.voucherNumber || '—'}
                          </td>
                          <td className="py-2.5 px-3">{v.voucherType}</td>
                          <td className="py-2.5 px-3 max-w-xs truncate">
                            <div className="font-medium text-slate-800 truncate">
                              {v.partyName || v.entries[0]?.ledgerName || '—'}
                            </div>
                            {v.narration && (
                              <div className="text-[11px] text-slate-500 truncate">
                                {v.narration}
                              </div>
                            )}
                          </td>
                          <td className="py-2.5 px-3 text-right font-mono tabular-nums font-semibold">
                            {formatINR(v.totalDebitPaise)}
                          </td>
                          <td className="py-2.5 px-3 text-slate-500 font-mono text-[11px] truncate max-w-[140px]">
                            <span className="inline-flex items-center gap-1">
                              <FileSpreadsheet className="w-3 h-3 shrink-0" />
                              {v.sourceFilename}
                            </span>
                          </td>
                          <td className="py-2.5 px-3 text-right">
                            <button
                              onClick={(e) => {
                                e.stopPropagation();
                                setSelectedVoucherId(v.id);
                              }}
                              className="text-xs font-medium text-[#0875BC] hover:underline whitespace-nowrap"
                            >
                              View Entry →
                            </button>
                          </td>
                        </tr>
                      ))}
                      {filteredVouchers.length === 0 && (
                        <tr>
                          <td colSpan={7} className="py-8 text-center text-slate-500">
                            No vouchers match the current filter criteria.
                          </td>
                        </tr>
                      )}
                    </tbody>
                  </table>
                </div>
              </div>
            </>
          )}
        </div>
      </div>
    </div>
  );
};
