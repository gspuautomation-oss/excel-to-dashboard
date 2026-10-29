import {
  ConsolidatedAccountingModel,
  formatINR,
  GSTInvoice,
  GSTMatchResult,
  PolicySettings,
} from '../types/accounting';
import { normalizeName } from './ledgerMapping';

export interface GSTReconciliationReport {
  hasGstData: boolean;
  summaryCards: {
    taxableOutwardSuppliesPaise: number;
    outputGstBooksPaise: number;
    inputGstBooksPaise: number;
    bookItcPaise: number;
    gstr2bItcPaise: number;
    gstr3bItcClaimedPaise: number;
    reconciliationDiffPaise: number;
    unmatchedInvoiceCount: number;
  };
  salesVsGstr1Matches: GSTMatchResult[];
  purchasesVsGstr2bMatches: GSTMatchResult[];
  gstr1VsGstr3bComparison: Array<{
    head: string;
    booksOutwardPaise: number;
    gstr1Paise: number;
    gstr3bPaise: number;
    diffGstr1Vs3bPaise: number;
    diffBooksVs3bPaise: number;
  }>;
  itcVsGstr3bComparison: Array<{
    head: string;
    booksEligibleItcPaise: number;
    gstr2bEligibleItcPaise: number;
    gstr3bClaimedItcPaise: number;
    excessClaimedOver2bPaise: number;
    unclaimedBooksVs2bPaise: number;
  }>;
  monthlyGstTrends: Array<{
    month: string;
    outputGstBooksPaise: number;
    outputGstGstr1Paise: number;
    inputItcBooksPaise: number;
    inputItcGstr2bPaise: number;
    mismatchPaise: number;
  }>;
  supplierMismatches: Array<{
    partyGstin: string;
    partyName: string;
    booksTaxPaise: number;
    portalTaxPaise: number;
    diffPaise: number;
    issueCount: number;
  }>;
}

function matchInvoiceLists(
  booksList: GSTInvoice[],
  portalList: GSTInvoice[],
  reconciliationType: GSTMatchResult['reconciliationType'],
  tolerancePaise: number
): GSTMatchResult[] {
  const results: GSTMatchResult[] = [];
  const matchedPortalIds = new Set<string>();
  const matchedBooksIds = new Set<string>();

  // Pass 1: EXACT MATCH (same GSTIN, exact invoice number, exact date, exact taxable & tax)
  for (const b of booksList) {
    const exactCandidate = portalList.find(
      (p) =>
        !matchedPortalIds.has(p.id) &&
        p.partyGstin === b.partyGstin &&
        p.invoiceNumber.trim().toUpperCase() === b.invoiceNumber.trim().toUpperCase() &&
        p.invoiceDate === b.invoiceDate &&
        p.taxableValuePaise === b.taxableValuePaise &&
        p.totalTaxPaise === b.totalTaxPaise
    );
    if (exactCandidate) {
      matchedBooksIds.add(b.id);
      matchedPortalIds.add(exactCandidate.id);
      results.push({
        id: `M_EXACT_${b.id}`,
        reconciliationType,
        matchStatus: 'EXACT',
        booksInvoice: b,
        portalInvoice: exactCandidate,
        taxableDiffPaise: 0,
        igstDiffPaise: 0,
        cgstDiffPaise: 0,
        sgstDiffPaise: 0,
        totalTaxDiffPaise: 0,
        discrepancyReasons: [],
      });
    }
  }

  // Pass 2: NORMALISED & TOLERANCE MATCH (same GSTIN + normalized invoice number)
  for (const b of booksList) {
    if (matchedBooksIds.has(b.id)) continue;
    const normCandidate = portalList.find(
      (p) =>
        !matchedPortalIds.has(p.id) &&
        p.partyGstin === b.partyGstin &&
        p.normalizedInvoiceNumber === b.normalizedInvoiceNumber &&
        b.normalizedInvoiceNumber.length > 0
    );
    if (normCandidate) {
      matchedBooksIds.add(b.id);
      matchedPortalIds.add(normCandidate.id);

      const sign = b.documentType === 'CREDIT_NOTE' ? -1 : 1;
      const taxableDiff = (b.taxableValuePaise - normCandidate.taxableValuePaise) * sign;
      const igstDiff = (b.igstPaise - normCandidate.igstPaise) * sign;
      const cgstDiff = (b.cgstPaise - normCandidate.cgstPaise) * sign;
      const sgstDiff = (b.sgstPaise - normCandidate.sgstPaise) * sign;
      const totalTaxDiff = (b.totalTaxPaise - normCandidate.totalTaxPaise) * sign;

      const reasons: string[] = [];
      if (b.invoiceNumber !== normCandidate.invoiceNumber) {
        reasons.push(`Invoice format normalised (${b.invoiceNumber} vs ${normCandidate.invoiceNumber})`);
      }
      if (b.invoiceDate !== normCandidate.invoiceDate) {
        reasons.push(`Date difference (${b.invoiceDate} vs ${normCandidate.invoiceDate})`);
      }
      if (b.returnPeriod !== normCandidate.returnPeriod) {
        reasons.push(`Return period mismatch (${b.returnPeriod} vs ${normCandidate.returnPeriod})`);
      }
      if (Math.abs(taxableDiff) > 0) {
        reasons.push(`Taxable value diff: ${formatINR(taxableDiff, true)}`);
      }
      if (Math.abs(totalTaxDiff) > 0) {
        reasons.push(`Tax diff: ${formatINR(totalTaxDiff, true)}`);
      }

      const isWithinTolerance =
        Math.abs(taxableDiff) <= tolerancePaise && Math.abs(totalTaxDiff) <= tolerancePaise;

      results.push({
        id: `M_NORM_${b.id}`,
        reconciliationType,
        matchStatus:
          Math.abs(taxableDiff) === 0 && Math.abs(totalTaxDiff) === 0
            ? 'NORMALISED'
            : isWithinTolerance
            ? 'TOLERANCE'
            : 'PROBABLE',
        booksInvoice: b,
        portalInvoice: normCandidate,
        taxableDiffPaise: taxableDiff,
        igstDiffPaise: igstDiff,
        cgstDiffPaise: cgstDiff,
        sgstDiffPaise: sgstDiff,
        totalTaxDiffPaise: totalTaxDiff,
        discrepancyReasons: reasons,
      });
    }
  }

  // Pass 3: PROBABLE MATCH (Requires Review — e.g. same GSTIN & exact taxable+tax, or same invoice number with GSTIN mismatch)
  for (const b of booksList) {
    if (matchedBooksIds.has(b.id)) continue;
    const probableCandidate = portalList.find(
      (p) =>
        !matchedPortalIds.has(p.id) &&
        ((p.partyGstin === b.partyGstin &&
          b.partyGstin !== '' &&
          Math.abs(p.taxableValuePaise - b.taxableValuePaise) <= tolerancePaise &&
          Math.abs(p.totalTaxPaise - b.totalTaxPaise) <= tolerancePaise) ||
          (p.normalizedInvoiceNumber === b.normalizedInvoiceNumber &&
            b.normalizedInvoiceNumber.length >= 3 &&
            normalizeName(p.partyName) === normalizeName(b.partyName)))
    );

    if (probableCandidate) {
      matchedBooksIds.add(b.id);
      matchedPortalIds.add(probableCandidate.id);

      const reasons: string[] = [];
      if (b.partyGstin !== probableCandidate.partyGstin) {
        reasons.push(`GSTIN Mismatch (Books: ${b.partyGstin || 'Blank'} vs Portal: ${probableCandidate.partyGstin || 'Blank'})`);
      }
      if (b.normalizedInvoiceNumber !== probableCandidate.normalizedInvoiceNumber) {
        reasons.push(`Invoice No Mismatch (Books: ${b.invoiceNumber} vs Portal: ${probableCandidate.invoiceNumber})`);
      }

      results.push({
        id: `M_PROB_${b.id}`,
        reconciliationType,
        matchStatus: 'PROBABLE',
        booksInvoice: b,
        portalInvoice: probableCandidate,
        taxableDiffPaise: b.taxableValuePaise - probableCandidate.taxableValuePaise,
        igstDiffPaise: b.igstPaise - probableCandidate.igstPaise,
        cgstDiffPaise: b.cgstPaise - probableCandidate.cgstPaise,
        sgstDiffPaise: b.sgstPaise - probableCandidate.sgstPaise,
        totalTaxDiffPaise: b.totalTaxPaise - probableCandidate.totalTaxPaise,
        discrepancyReasons: reasons,
      });
    }
  }

  // Pass 4: Remaining unmatched in Books -> MISSING_IN_PORTAL
  for (const b of booksList) {
    if (matchedBooksIds.has(b.id)) continue;
    results.push({
      id: `M_MISS_PORTAL_${b.id}`,
      reconciliationType,
      matchStatus: 'MISSING_IN_PORTAL',
      booksInvoice: b,
      taxableDiffPaise: b.taxableValuePaise,
      igstDiffPaise: b.igstPaise,
      cgstDiffPaise: b.cgstPaise,
      sgstDiffPaise: b.sgstPaise,
      totalTaxDiffPaise: b.totalTaxPaise,
      discrepancyReasons: [
        reconciliationType === 'SALES_VS_GSTR1'
          ? 'Recorded in Tally Sales Books but missing in uploaded GSTR-1.'
          : 'Recorded in Tally Purchase Books but missing in uploaded GSTR-2B (Follow up with supplier).',
      ],
    });
  }

  // Pass 5: Remaining unmatched in Portal -> MISSING_IN_BOOKS
  for (const p of portalList) {
    if (matchedPortalIds.has(p.id)) continue;
    results.push({
      id: `M_MISS_BOOKS_${p.id}`,
      reconciliationType,
      matchStatus: 'MISSING_IN_BOOKS',
      portalInvoice: p,
      taxableDiffPaise: -p.taxableValuePaise,
      igstDiffPaise: -p.igstPaise,
      cgstDiffPaise: -p.cgstPaise,
      sgstDiffPaise: -p.sgstPaise,
      totalTaxDiffPaise: -p.totalTaxPaise,
      discrepancyReasons: [
        reconciliationType === 'SALES_VS_GSTR1'
          ? 'Present in GSTR-1 return but omitted in Tally Sales Books.'
          : 'Present in GSTR-2B return but not booked in Tally Purchase Register.',
      ],
    });
  }

  return results;
}

export function runGSTReconciliation(
  model: ConsolidatedAccountingModel,
  policy: PolicySettings
): GSTReconciliationReport {
  const tallySales = model.gstInvoices.filter((i) => i.source === 'TALLY_SALES');
  const tallyPurchases = model.gstInvoices.filter((i) => i.source === 'TALLY_PURCHASE');
  const gstr1List = model.gstInvoices.filter((i) => i.source === 'GSTR1');
  const gstr2bList = model.gstInvoices.filter((i) => i.source === 'GSTR2B');

  const salesVsGstr1Matches = matchInvoiceLists(
    tallySales,
    gstr1List,
    'SALES_VS_GSTR1',
    policy.gstTolerancePaise
  );

  const purchasesVsGstr2bMatches = matchInvoiceLists(
    tallyPurchases,
    gstr2bList,
    'PURCHASE_VS_GSTR2B',
    policy.gstTolerancePaise
  );

  const sumGstList = (list: GSTInvoice[], eligibleOnly = false) => {
    let taxable = 0;
    let igst = 0;
    let cgst = 0;
    let sgst = 0;
    let cess = 0;
    for (const item of list) {
      if (eligibleOnly && item.itcEligibility === 'INELIGIBLE') continue;
      const mult = item.documentType === 'CREDIT_NOTE' ? -1 : 1;
      taxable += item.taxableValuePaise * mult;
      igst += item.igstPaise * mult;
      cgst += item.cgstPaise * mult;
      sgst += item.sgstPaise * mult;
      cess += item.cessPaise * mult;
    }
    return {
      taxable,
      igst,
      cgst,
      sgst,
      cess,
      totalTax: igst + cgst + sgst + cess,
    };
  };

  const booksOutward = sumGstList(tallySales);
  const gstr1Sums = sumGstList(gstr1List);
  const booksInward = sumGstList(tallyPurchases, true);
  const gstr2bSums = sumGstList(gstr2bList, true);

  // Aggregate GSTR-3B summaries
  let g3bOutTaxable = 0;
  let g3bOutIgst = 0;
  let g3bOutCgst = 0;
  let g3bOutSgst = 0;
  let g3bOutCess = 0;
  let g3bItcIgst = 0;
  let g3bItcCgst = 0;
  let g3bItcSgst = 0;
  let g3bItcCess = 0;

  for (const s3b of model.gstr3bSummaries) {
    g3bOutTaxable += s3b.outwardTaxableValuePaise;
    g3bOutIgst += s3b.outwardIgstPaise;
    g3bOutCgst += s3b.outwardCgstPaise;
    g3bOutSgst += s3b.outwardSgstPaise;
    g3bOutCess += s3b.outwardCessPaise;
    g3bItcIgst += s3b.eligibleItcIgstPaise;
    g3bItcCgst += s3b.eligibleItcCgstPaise;
    g3bItcSgst += s3b.eligibleItcSgstPaise;
    g3bItcCess += s3b.eligibleItcCessPaise;
  }

  const g3bTotalItc = g3bItcIgst + g3bItcCgst + g3bItcSgst + g3bItcCess;

  const gstr1VsGstr3bComparison = [
    {
      head: 'Taxable Value of Outward Supplies',
      booksOutwardPaise: booksOutward.taxable,
      gstr1Paise: gstr1Sums.taxable,
      gstr3bPaise: g3bOutTaxable,
      diffGstr1Vs3bPaise: gstr1Sums.taxable - g3bOutTaxable,
      diffBooksVs3bPaise: booksOutward.taxable - g3bOutTaxable,
    },
    {
      head: 'Integrated Tax (IGST)',
      booksOutwardPaise: booksOutward.igst,
      gstr1Paise: gstr1Sums.igst,
      gstr3bPaise: g3bOutIgst,
      diffGstr1Vs3bPaise: gstr1Sums.igst - g3bOutIgst,
      diffBooksVs3bPaise: booksOutward.igst - g3bOutIgst,
    },
    {
      head: 'Central Tax (CGST)',
      booksOutwardPaise: booksOutward.cgst,
      gstr1Paise: gstr1Sums.cgst,
      gstr3bPaise: g3bOutCgst,
      diffGstr1Vs3bPaise: gstr1Sums.cgst - g3bOutCgst,
      diffBooksVs3bPaise: booksOutward.cgst - g3bOutCgst,
    },
    {
      head: 'State / UT Tax (SGST)',
      booksOutwardPaise: booksOutward.sgst,
      gstr1Paise: gstr1Sums.sgst,
      gstr3bPaise: g3bOutSgst,
      diffGstr1Vs3bPaise: gstr1Sums.sgst - g3bOutSgst,
      diffBooksVs3bPaise: booksOutward.sgst - g3bOutSgst,
    },
    {
      head: 'Cess',
      booksOutwardPaise: booksOutward.cess,
      gstr1Paise: gstr1Sums.cess,
      gstr3bPaise: g3bOutCess,
      diffGstr1Vs3bPaise: gstr1Sums.cess - g3bOutCess,
      diffBooksVs3bPaise: booksOutward.cess - g3bOutCess,
    },
  ];

  const itcVsGstr3bComparison = [
    {
      head: 'Integrated Tax (IGST) ITC',
      booksEligibleItcPaise: booksInward.igst,
      gstr2bEligibleItcPaise: gstr2bSums.igst,
      gstr3bClaimedItcPaise: g3bItcIgst,
      excessClaimedOver2bPaise: g3bItcIgst - gstr2bSums.igst,
      unclaimedBooksVs2bPaise: booksInward.igst - gstr2bSums.igst,
    },
    {
      head: 'Central Tax (CGST) ITC',
      booksEligibleItcPaise: booksInward.cgst,
      gstr2bEligibleItcPaise: gstr2bSums.cgst,
      gstr3bClaimedItcPaise: g3bItcCgst,
      excessClaimedOver2bPaise: g3bItcCgst - gstr2bSums.cgst,
      unclaimedBooksVs2bPaise: booksInward.cgst - gstr2bSums.cgst,
    },
    {
      head: 'State / UT Tax (SGST) ITC',
      booksEligibleItcPaise: booksInward.sgst,
      gstr2bEligibleItcPaise: gstr2bSums.sgst,
      gstr3bClaimedItcPaise: g3bItcSgst,
      excessClaimedOver2bPaise: g3bItcSgst - gstr2bSums.sgst,
      unclaimedBooksVs2bPaise: booksInward.sgst - gstr2bSums.sgst,
    },
    {
      head: 'Total Input Tax Credit',
      booksEligibleItcPaise: booksInward.totalTax,
      gstr2bEligibleItcPaise: gstr2bSums.totalTax,
      gstr3bClaimedItcPaise: g3bTotalItc,
      excessClaimedOver2bPaise: g3bTotalItc - gstr2bSums.totalTax,
      unclaimedBooksVs2bPaise: booksInward.totalTax - gstr2bSums.totalTax,
    },
  ];

  const unmatchedSalesCount = salesVsGstr1Matches.filter(
    (m) => m.matchStatus === 'MISSING_IN_PORTAL' || m.matchStatus === 'MISSING_IN_BOOKS' || m.matchStatus === 'PROBABLE'
  ).length;
  const unmatchedPurchCount = purchasesVsGstr2bMatches.filter(
    (m) => m.matchStatus === 'MISSING_IN_PORTAL' || m.matchStatus === 'MISSING_IN_BOOKS' || m.matchStatus === 'PROBABLE'
  ).length;

  // Monthly GST Trends
  const monthSet = new Set<string>(model.availableMonths);
  for (const inv of model.gstInvoices) {
    if (inv.returnPeriod) monthSet.add(inv.returnPeriod);
  }
  const sortedMonths = Array.from(monthSet).sort();

  const monthlyGstTrends = sortedMonths.map((ym) => {
    const mSalesBooks = sumGstList(tallySales.filter((i) => i.returnPeriod === ym));
    const mSalesG1 = sumGstList(gstr1List.filter((i) => i.returnPeriod === ym));
    const mPurchBooks = sumGstList(tallyPurchases.filter((i) => i.returnPeriod === ym), true);
    const mPurch2B = sumGstList(gstr2bList.filter((i) => i.returnPeriod === ym), true);
    return {
      month: ym,
      outputGstBooksPaise: mSalesBooks.totalTax,
      outputGstGstr1Paise: mSalesG1.totalTax,
      inputItcBooksPaise: mPurchBooks.totalTax,
      inputItcGstr2bPaise: mPurch2B.totalTax,
      mismatchPaise: Math.abs(mPurchBooks.totalTax - mPurch2B.totalTax),
    };
  });

  // Supplier mismatch breakdown
  const supplierMap = new Map<
    string,
    {
      partyGstin: string;
      partyName: string;
      booksTaxPaise: number;
      portalTaxPaise: number;
      issueCount: number;
    }
  >();

  for (const m of purchasesVsGstr2bMatches) {
    const gstin = m.booksInvoice?.partyGstin || m.portalInvoice?.partyGstin || 'UNREGISTERED';
    const name = m.booksInvoice?.partyName || m.portalInvoice?.partyName || 'Supplier';
    const key = `${gstin}|${normalizeName(name)}`;
    let entry = supplierMap.get(key);
    if (!entry) {
      entry = { partyGstin: gstin, partyName: name, booksTaxPaise: 0, portalTaxPaise: 0, issueCount: 0 };
      supplierMap.set(key, entry);
    }
    entry.booksTaxPaise += m.booksInvoice?.totalTaxPaise || 0;
    entry.portalTaxPaise += m.portalInvoice?.totalTaxPaise || 0;
    if (m.matchStatus !== 'EXACT' && m.matchStatus !== 'NORMALISED') {
      entry.issueCount += 1;
    }
  }

  const supplierMismatches = Array.from(supplierMap.values())
    .map((s) => ({
      ...s,
      diffPaise: s.booksTaxPaise - s.portalTaxPaise,
    }))
    .filter((s) => s.issueCount > 0 || Math.abs(s.diffPaise) > 0)
    .sort((a, b) => Math.abs(b.diffPaise) - Math.abs(a.diffPaise));

  const reconciliationDiffPaise =
    Math.abs(booksOutward.totalTax - gstr1Sums.totalTax) +
    Math.abs(booksInward.totalTax - gstr2bSums.totalTax);

  return {
    hasGstData: model.gstInvoices.length > 0 || model.gstr3bSummaries.length > 0,
    summaryCards: {
      taxableOutwardSuppliesPaise: booksOutward.taxable || gstr1Sums.taxable,
      outputGstBooksPaise: booksOutward.totalTax,
      inputGstBooksPaise: booksInward.totalTax,
      bookItcPaise: booksInward.totalTax,
      gstr2bItcPaise: gstr2bSums.totalTax,
      gstr3bItcClaimedPaise: g3bTotalItc,
      reconciliationDiffPaise,
      unmatchedInvoiceCount: unmatchedSalesCount + unmatchedPurchCount,
    },
    salesVsGstr1Matches,
    purchasesVsGstr2bMatches,
    gstr1VsGstr3bComparison,
    itcVsGstr3bComparison,
    monthlyGstTrends,
    supplierMismatches,
  };
}
