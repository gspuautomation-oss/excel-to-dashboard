import { AccountClassification } from '../types/accounting';

export function normalizeName(raw: string | undefined | null): string {
  if (!raw) return '';
  return raw
    .trim()
    .replace(/\s+/g, ' ')
    .toLowerCase();
}

export function stripLeadingBulletOrRoman(raw: string): string {
  let s = raw.trim();
  for (let pass = 0; pass < 4; pass++) {
    const next = s
      .replace(
        /^(?:[ivx]+\.|[0-9]+\.|[a-z]\)|[a-z]\.|-[\s]*|•[\s]*|\([a-z0-9ivx]+\))[\s.:-]*/i,
        ''
      )
      .trim();
    if (next === s) break;
    s = next;
  }
  return s.replace(/\s+/g, ' ').trim();
}

export function normalizeInvoiceNo(raw: string | undefined | null): string {
  if (!raw) return '';
  return raw
    .toUpperCase()
    .replace(/[^A-Z0-9]/g, '')
    .replace(/^0+/, '');
}

export function classifyLedgerAndGroup(
  ledgerName: string,
  parentGroup: string,
  manualOverrides?: Record<string, AccountClassification>,
  statementContext?: 'PNL' | 'BALANCE_SHEET' | 'TRIAL_BALANCE' | 'TALLY_XML'
): AccountClassification {
  const normLedgerRaw = normalizeName(ledgerName);
  const normLedger = normalizeName(stripLeadingBulletOrRoman(ledgerName));
  const normGroup = normalizeName(parentGroup);

  if (manualOverrides) {
    if (manualOverrides[normLedgerRaw]) return manualOverrides[normLedgerRaw];
    if (manualOverrides[normLedger]) return manualOverrides[normLedger];
  }

  // Check Suspense first
  if (normGroup.includes('suspense') || normLedger.includes('suspense')) {
    return 'Suspense';
  }

  // ==========================================================================
  // 1. CONTEXT-AWARE BALANCE SHEET MAPPING (When parsing a Balance Sheet)
  // ==========================================================================
  if (statementContext === 'BALANCE_SHEET') {
    // 1A. Strict Schedule III Section-Aware Hierarchy (when active section is known)
    if (normGroup === "shareholders' funds" || normGroup === 'shareholders funds') {
      return 'Equity';
    }

    if (normGroup === 'non-current liabilities' || normGroup === 'non current liabilities') {
      return 'LongTermBorrowings';
    }

    if (normGroup === 'current liabilities') {
      if (
        normLedger.includes('short-term borrowing') ||
        normLedger.includes('short term borrowing') ||
        normLedger.includes('current borrowing') ||
        normLedger.includes('working capital loan') ||
        normLedger.includes('working capital od') ||
        normLedger.includes('cash credit') ||
        normLedger.includes('bank od') ||
        normLedger.includes('overdraft')
      ) {
        return 'ShortTermBorrowings';
      }
      if (
        normLedger.includes('trade payable') ||
        normLedger.includes('sundry creditor') ||
        normLedger.includes('accounts payable') ||
        normLedger.includes('dues of micro') ||
        normLedger.includes('dues of creditor') ||
        normLedger.includes('msme') ||
        normLedger.includes('creditor')
      ) {
        return 'TradePayables';
      }
      if (
        normLedger.includes('duties') ||
        normLedger.includes('statutory') ||
        normLedger.includes('gst payable') ||
        normLedger.includes('tds payable') ||
        normLedger.includes('provision for tax') ||
        normLedger.includes('income tax provision')
      ) {
        return 'TaxLiabilities';
      }
      return 'OtherCurrentLiabilities';
    }

    if (normGroup === 'non-current assets' || normGroup === 'non current assets') {
      if (
        normLedger.includes('property, plant') ||
        normLedger.includes('property plant') ||
        normLedger.includes('tangible asset') ||
        normLedger.includes('intangible asset') ||
        normLedger.includes('capital work-in-progress') ||
        normLedger.includes('capital work in progress') ||
        normLedger.includes('cwip') ||
        normLedger.includes('fixed asset') ||
        normLedger.includes('plant &') ||
        normLedger.includes('plant and') ||
        normLedger.includes('machinery') ||
        normLedger.includes('building') ||
        normLedger.includes('furniture') ||
        normLedger.includes('office equipment') ||
        normLedger.includes('computer') ||
        normLedger.includes('vehicle') ||
        normLedger.includes('goodwill') ||
        normLedger.includes('software')
      ) {
        return 'FixedAssets';
      }
      if (
        normLedger.includes('non-current investment') ||
        normLedger.includes('non current investment') ||
        normLedger.includes('long-term investment') ||
        normLedger.includes('long term investment') ||
        (normLedger.includes('investment') && !normLedger.includes('current investment'))
      ) {
        return 'Investments';
      }
      // Deferred Tax Assets, Long-Term Loans and Advances, Security Deposits, Other Non-Current Assets
      return 'NonCurrentAssets';
    }

    if (normGroup === 'current assets') {
      if (
        normLedger.includes('inventori') ||
        normLedger.includes('stock-in-hand') ||
        normLedger.includes('stock in hand') ||
        normLedger.includes('closing stock') ||
        normLedger.includes('raw material') ||
        normLedger.includes('work-in-progress') ||
        normLedger.includes('finished goods')
      ) {
        return 'Inventory';
      }
      if (
        normLedger.includes('trade receivable') ||
        normLedger.includes('sundry debtor') ||
        normLedger.includes('accounts receivable') ||
        normLedger.includes('debtor')
      ) {
        return 'TradeReceivables';
      }
      if (
        normLedger.includes('cash and cash equivalent') ||
        normLedger.includes('cash & cash equivalent') ||
        normLedger.includes('cash and bank') ||
        normLedger.includes('cash & bank') ||
        normLedger.includes('cash on hand') ||
        normLedger.includes('cash in hand') ||
        normLedger.includes('cash-in-hand') ||
        normLedger === 'cash' ||
        normLedger === 'petty cash'
      ) {
        return 'CashAndEquivalents';
      }
      if (
        normLedger.includes('bank balance') ||
        normLedger.includes('other bank balance') ||
        normLedger.includes('balances with bank') ||
        normLedger.includes('current account') ||
        normLedger.includes('fixed deposit') ||
        normLedger.includes('hdfc bank') ||
        normLedger.includes('icici bank') ||
        normLedger.includes('sbi ') ||
        normLedger.includes('axis bank') ||
        normLedger.includes('kotak bank')
      ) {
        return 'BankBalances';
      }
      // Short-Term Loans and Advances, Other Current Assets, Current Investments, Prepaid, Balance with Govt
      return 'OtherCurrentAssets';
    }

    // 1B. Fallback Balance Sheet Classification (when section header was not explicitly present)
    // Deferred Tax Assets vs Deferred Tax Liabilities on Balance Sheet
    if (normLedger.includes('deferred tax asset') || normGroup.includes('deferred tax asset')) {
      return 'NonCurrentAssets';
    }
    if (normLedger.includes('deferred tax liabilit') || normGroup.includes('deferred tax liabilit')) {
      return 'LongTermBorrowings';
    }

    // Equity (Share Capital, Reserves & Surplus, Money received against share warrants)
    if (
      normLedger.includes('share capital') ||
      normLedger.includes('equity share') ||
      normLedger.includes('preference share') ||
      normLedger.includes('reserves and surplus') ||
      normLedger.includes('reserves & surplus') ||
      normLedger.includes('retained earnings') ||
      normLedger.includes('share warrant') ||
      normLedger.includes('share application') ||
      normLedger.includes('capital reserve') ||
      normLedger.includes('general reserve') ||
      normLedger.includes('surplus in statement') ||
      normLedger.includes('profit & loss a/c') ||
      normLedger.includes('profit and loss account') ||
      normGroup.includes('shareholders') ||
      normGroup.includes('capital account') ||
      normGroup.includes('reserves') ||
      normGroup === 'equity'
    ) {
      return 'Equity';
    }

    // Non-Current Liabilities / Long-Term Borrowings
    const isLoansAndAdvancesAsset =
      normLedger.includes('loans and advances') ||
      normLedger.includes('loans & advances') ||
      normGroup.includes('loans and advances') ||
      normGroup.includes('loans & advances');

    if (
      !isLoansAndAdvancesAsset &&
      (normLedger.includes('long-term borrowing') ||
        normLedger.includes('long term borrowing') ||
        normLedger.includes('non-current borrowing') ||
        normLedger.includes('other long-term liabilit') ||
        normLedger.includes('other long term liabilit') ||
        normLedger.includes('long-term provision') ||
        normLedger.includes('long term provision') ||
        normLedger.includes('term loan') ||
        normLedger.includes('secured loan') ||
        normLedger.includes('unsecured loan') ||
        normGroup.includes('non-current liabilit') ||
        normGroup.includes('non current liabilit') ||
        normGroup.includes('long-term borrowing') ||
        normGroup.includes('secured loan') ||
        normGroup.includes('unsecured loan') ||
        normGroup.includes('loans (liability)'))
    ) {
      return 'LongTermBorrowings';
    }

    // Short-Term Borrowings
    if (
      normLedger.includes('short-term borrowing') ||
      normLedger.includes('short term borrowing') ||
      normLedger.includes('current borrowing') ||
      normLedger.includes('working capital loan') ||
      normLedger.includes('working capital od') ||
      normLedger.includes('cash credit') ||
      normLedger.includes('bank od') ||
      normLedger.includes('overdraft') ||
      normGroup.includes('short-term borrowing') ||
      normGroup.includes('bank od') ||
      normGroup.includes('bank occ')
    ) {
      return 'ShortTermBorrowings';
    }

    // Trade Payables
    if (
      normLedger.includes('trade payable') ||
      normLedger.includes('sundry creditor') ||
      normLedger.includes('accounts payable') ||
      normLedger.includes('dues of micro') ||
      normLedger.includes('dues of creditor') ||
      normLedger.includes('msme payable') ||
      normGroup.includes('trade payable') ||
      normGroup.includes('sundry creditor')
    ) {
      return 'TradePayables';
    }

    // Current Liabilities & Provisions
    const isNonCurrGroup =
      normGroup.includes('non-current') || normGroup.includes('non current');
    if (
      normLedger.includes('other current liabilit') ||
      normLedger.includes('short-term provision') ||
      normLedger.includes('short term provision') ||
      normLedger.includes('statutory dues') ||
      normLedger.includes('duties & taxes') ||
      normLedger.includes('duties and taxes') ||
      normLedger.includes('outstanding expense') ||
      normLedger.includes('expenses payable') ||
      normLedger.includes('provision for') ||
      normLedger.includes('advance from customer') ||
      (!isNonCurrGroup && normGroup.includes('current liabilit')) ||
      normGroup.includes('provisions') ||
      normGroup.includes('duties')
    ) {
      if (
        normLedger.includes('duties') ||
        normLedger.includes('statutory') ||
        normLedger.includes('gst payable') ||
        normLedger.includes('tds payable') ||
        normLedger.includes('provision for tax') ||
        normLedger.includes('income tax provision')
      ) {
        return 'TaxLiabilities';
      }
      return 'OtherCurrentLiabilities';
    }

    // Fixed Assets / Property, Plant & Equipment / Intangibles
    if (
      normLedger.includes('property, plant') ||
      normLedger.includes('property plant') ||
      normLedger.includes('tangible asset') ||
      normLedger.includes('intangible asset') ||
      normLedger.includes('capital work-in-progress') ||
      normLedger.includes('capital work in progress') ||
      normLedger.includes('cwip') ||
      normLedger.includes('fixed asset') ||
      normLedger.includes('plant &') ||
      normLedger.includes('plant and') ||
      normLedger.includes('machinery') ||
      normLedger.includes('building') ||
      normLedger.includes('furniture') ||
      normLedger.includes('office equipment') ||
      normLedger.includes('computer') ||
      normLedger.includes('vehicle') ||
      normLedger.includes('goodwill') ||
      normLedger.includes('software') ||
      normGroup.includes('fixed asset') ||
      normGroup.includes('property, plant') ||
      normGroup.includes('tangible') ||
      normGroup.includes('intangible')
    ) {
      return 'FixedAssets';
    }

    // Investments (Non-Current)
    if (
      normLedger.includes('non-current investment') ||
      normLedger.includes('non current investment') ||
      normLedger.includes('long-term investment') ||
      normLedger.includes('long term investment') ||
      (normLedger.includes('investment') && !normLedger.includes('current investment')) ||
      normGroup.includes('investment')
    ) {
      return 'Investments';
    }

    // Other Non-Current Assets & Long-Term Loans and Advances
    if (
      normLedger.includes('long-term loans and advances') ||
      normLedger.includes('long term loans and advances') ||
      normLedger.includes('long-term loans & advances') ||
      normLedger.includes('long term loans & advances') ||
      normLedger.includes('long-term loan') ||
      normLedger.includes('long term loan') ||
      normLedger.includes('other non-current asset') ||
      normLedger.includes('other non current asset') ||
      normLedger.includes('security deposit') ||
      normLedger.includes('unamortised') ||
      normLedger.includes('unamortized') ||
      normGroup.includes('non-current asset') ||
      normGroup.includes('non current asset')
    ) {
      return 'NonCurrentAssets';
    }

    // Inventories
    if (
      normLedger.includes('inventori') ||
      normLedger.includes('stock-in-hand') ||
      normLedger.includes('stock in hand') ||
      normLedger.includes('closing stock') ||
      normLedger.includes('raw material') ||
      normLedger.includes('work-in-progress') ||
      normLedger.includes('finished goods') ||
      normGroup.includes('stock-in-hand') ||
      normGroup.includes('inventori')
    ) {
      return 'Inventory';
    }

    // Trade Receivables
    if (
      normLedger.includes('trade receivable') ||
      normLedger.includes('sundry debtor') ||
      normLedger.includes('accounts receivable') ||
      normGroup.includes('trade receivable') ||
      normGroup.includes('sundry debtor')
    ) {
      return 'TradeReceivables';
    }

    // Cash and Cash Equivalents & Bank Balances
    if (
      normLedger.includes('cash and cash equivalent') ||
      normLedger.includes('cash & cash equivalent') ||
      normLedger.includes('cash and bank') ||
      normLedger.includes('cash & bank') ||
      normLedger.includes('cash on hand') ||
      normLedger.includes('cash in hand') ||
      normLedger.includes('cash-in-hand') ||
      normLedger === 'cash' ||
      normLedger === 'petty cash' ||
      normGroup.includes('cash-in-hand') ||
      normGroup.includes('cash and cash equivalent')
    ) {
      return 'CashAndEquivalents';
    }

    if (
      normLedger.includes('bank balance') ||
      normLedger.includes('balances with bank') ||
      normLedger.includes('current account') ||
      normLedger.includes('fixed deposit') ||
      normLedger.includes('hdfc bank') ||
      normLedger.includes('icici bank') ||
      normLedger.includes('sbi ') ||
      normLedger.includes('axis bank') ||
      normLedger.includes('kotak bank') ||
      normGroup.includes('bank account') ||
      normGroup === 'bank'
    ) {
      return 'BankBalances';
    }

    // Short-Term Loans & Advances and Other Current Assets
    if (
      normLedger.includes('short-term loans and advances') ||
      normLedger.includes('short term loans and advances') ||
      normLedger.includes('short-term loans & advances') ||
      normLedger.includes('short term loans & advances') ||
      normLedger.includes('other current asset') ||
      normLedger.includes('current investment') ||
      normLedger.includes('loans and advances') ||
      normLedger.includes('loans & advances') ||
      normLedger.includes('prepaid') ||
      normLedger.includes('advance') ||
      normLedger.includes('accrued') ||
      normLedger.includes('balance with government') ||
      normLedger.includes('tds receivable') ||
      normLedger.includes('gst input') ||
      (!isNonCurrGroup && normGroup.includes('current asset')) ||
      normGroup.includes('loans & advances') ||
      normGroup.includes('deposits (asset)')
    ) {
      return 'OtherCurrentAssets';
    }

    // Section-based fallback on Balance Sheet so valid rows are not dropped
    if (normGroup.includes('non-current liabilit') || normGroup.includes('non current liabilit')) {
      return 'LongTermBorrowings';
    }
    if (normGroup.includes('current liabilit')) return 'OtherCurrentLiabilities';
    if (normGroup.includes('non-current asset') || normGroup.includes('non current asset')) {
      return 'NonCurrentAssets';
    }
    if (normGroup.includes('current asset')) return 'OtherCurrentAssets';
  }

  // ==========================================================================
  // 2. PROFIT & LOSS AND GENERAL LEDGER MAPPING
  // ==========================================================================

  // Sales & Revenue from Operations
  if (
    normLedger.includes('sales return') ||
    normLedger.includes('return inward') ||
    normLedger.includes('credit note discount')
  ) {
    return 'SalesReturn';
  }
  if (
    normLedger.includes('revenue from operation') ||
    normLedger.includes('sale of product') ||
    normLedger.includes('sale of service') ||
    normLedger.includes('income from operation') ||
    normLedger.includes('operating revenue') ||
    normGroup.includes('sales account') ||
    normGroup === 'sales' ||
    normGroup === 'revenue from operations'
  ) {
    return 'Sales';
  }

  // Other Operating Income vs Other Non-Operating Income
  if (
    normLedger === 'other operating income' ||
    normLedger === 'other operating revenue' ||
    normGroup.includes('direct income') ||
    normLedger.includes('job work income')
  ) {
    return 'OtherOperatingIncome';
  }
  if (
    normLedger === 'other income' ||
    normLedger.includes('other non-operating income') ||
    normLedger.includes('interest income') ||
    normLedger.includes('interest received') ||
    normLedger.includes('dividend income') ||
    normLedger.includes('dividend received') ||
    normLedger.includes('discount received') ||
    normLedger.includes('foreign exchange gain') ||
    normLedger.includes('miscellaneous income') ||
    normLedger.includes('scrap recovery') ||
    normGroup.includes('indirect income') ||
    normGroup.includes('other income')
  ) {
    return 'OtherIncome';
  }

  // Purchases & Cost of Materials
  if (normLedger.includes('purchase return') || normLedger.includes('return outward')) {
    return 'PurchaseReturn';
  }
  if (
    normLedger.includes('cost of materials consumed') ||
    normLedger.includes('cost of material consumed') ||
    normLedger.includes('purchases of stock-in-trade') ||
    normLedger.includes('purchase of stock-in-trade') ||
    normLedger.includes('cost of goods sold') ||
    normGroup.includes('purchase account') ||
    normGroup === 'purchases' ||
    normGroup === 'purchase'
  ) {
    return 'Purchases';
  }

  // Direct Expenses & Changes in Inventories
  if (
    normLedger.includes('changes in inventories') ||
    normLedger.includes('change in inventories') ||
    normLedger.includes('direct expense') ||
    normLedger.includes('cost of services') ||
    normLedger.includes('freight inward') ||
    normLedger.includes('factory wages') ||
    normLedger.includes('power & fuel') ||
    normLedger.includes('power and fuel') ||
    normLedger.includes('carriage inward') ||
    normLedger.includes('manufacturing expense') ||
    normGroup.includes('direct expense')
  ) {
    return 'DirectExpenses';
  }

  // Employee Benefits Expense
  if (
    normLedger.includes('employee benefit') ||
    normLedger.includes('employee cost') ||
    normLedger.includes('salary') ||
    normLedger.includes('salaries') ||
    normLedger.includes('staff welfare') ||
    normLedger.includes('provident fund') ||
    normLedger.includes('contribution to pf') ||
    normLedger.includes('gratuity') ||
    normLedger.includes('esic') ||
    normLedger.includes('bonus') ||
    normLedger.includes('employee') ||
    normLedger.includes('director remuneration') ||
    normLedger.includes('directors remuneration') ||
    normGroup.includes('employee benefit')
  ) {
    return 'EmployeeExpenses';
  }

  // Finance Costs (Must be checked BEFORE any generic income/interest check!)
  if (
    normLedger.includes('finance cost') ||
    normLedger.includes('interest expense') ||
    normLedger.includes('interest on') ||
    normLedger.includes('bank interest') ||
    normLedger.includes('loan interest') ||
    normLedger.includes('borrowing cost') ||
    normLedger.includes('bank charges') ||
    normLedger.includes('processing fee') ||
    normGroup.includes('finance cost')
  ) {
    return 'FinanceCosts';
  }

  // Depreciation and Amortisation
  if (
    normLedger.includes('depreciation') ||
    normLedger.includes('amortisation') ||
    normLedger.includes('amortization') ||
    normGroup.includes('depreciation')
  ) {
    return 'Depreciation';
  }

  // Prior Period Items / Exceptional Items / Extraordinary Items
  if (
    normLedger.includes('prior period') ||
    normLedger.includes('prior-period') ||
    normLedger.includes('exceptional item') ||
    normLedger.includes('extraordinary item') ||
    normGroup.includes('prior period') ||
    normGroup.includes('exceptional')
  ) {
    return 'PriorPeriodExpenses';
  }

  // Tax Expense: Current Tax vs Deferred Tax vs General Tax Expense
  if (
    normLedger.includes('current tax') ||
    normLedger.includes('provision for income tax') ||
    normLedger.includes('provision for tax') ||
    normLedger.includes('corporate income tax') ||
    normLedger.includes('income tax expense') ||
    normLedger.includes('excess/short provision') ||
    normLedger.includes('short/excess provision') ||
    normLedger.includes('earlier year tax') ||
    normLedger.includes('mat credit')
  ) {
    return 'CurrentTax';
  }

  if (
    (normLedger.includes('deferred tax') || normGroup.includes('deferred tax')) &&
    statementContext !== 'BALANCE_SHEET' &&
    !normLedger.includes('deferred tax asset') &&
    !normLedger.includes('deferred tax liabilit')
  ) {
    return 'DeferredTax';
  }

  if (
    normLedger.includes('income tax') ||
    normLedger.includes('corporate tax') ||
    normLedger === 'tax expense' ||
    normGroup.includes('tax expense')
  ) {
    return 'TaxExpense';
  }

  // Selling & Distribution Expenses
  if (
    normLedger.includes('advertisement') ||
    normLedger.includes('marketing') ||
    normLedger.includes('sales promotion') ||
    normLedger.includes('commission paid') ||
    normLedger.includes('brokerage') ||
    normLedger.includes('freight outward') ||
    normLedger.includes('selling') ||
    normLedger.includes('business promotion')
  ) {
    return 'SellingExpenses';
  }

  // Other Expenses / Administrative & General Expenses
  if (
    normLedger === 'other expenses' ||
    normLedger === 'other expense' ||
    normLedger.includes('other operating expense') ||
    normLedger.includes('administrative') ||
    normLedger.includes('general expense') ||
    normGroup.includes('indirect expense') ||
    normGroup.includes('administrative') ||
    normGroup.includes('other expense') ||
    normLedger.includes('rent') ||
    normLedger.includes('audit fee') ||
    normLedger.includes('legal') ||
    normLedger.includes('professional fee') ||
    normLedger.includes('consultancy') ||
    normLedger.includes('printing') ||
    normLedger.includes('stationery') ||
    normLedger.includes('telephone') ||
    normLedger.includes('internet') ||
    normLedger.includes('software expense') ||
    normLedger.includes('it expense') ||
    normLedger.includes('hosting') ||
    normLedger.includes('office expense') ||
    normLedger.includes('repairs') ||
    normLedger.includes('maintenance') ||
    normLedger.includes('insurance') ||
    normLedger.includes('travelling') ||
    normLedger.includes('conveyance') ||
    normLedger.includes('electricity') ||
    normLedger.includes('rates and taxes') ||
    normLedger.includes('rates & taxes') ||
    normLedger.includes('miscellaneous expense')
  ) {
    return 'AdminExpenses';
  }

  // ==========================================================================
  // 3. GENERAL BALANCE SHEET CLASSIFICATIONS (Trial Balance / Tally XML)
  // ==========================================================================
  if (normLedger.includes('deferred tax asset') || normGroup.includes('deferred tax asset')) {
    return 'NonCurrentAssets';
  }
  if (normLedger.includes('deferred tax liabilit') || normGroup.includes('deferred tax liabilit')) {
    return 'LongTermBorrowings';
  }

  if (
    normGroup.includes('fixed asset') ||
    normGroup.includes('property, plant') ||
    normLedger.includes('property, plant') ||
    normLedger.includes('tangible asset') ||
    normLedger.includes('intangible asset') ||
    normLedger.includes('capital work-in-progress')
  ) {
    return 'FixedAssets';
  }
  if (normGroup.includes('investment') || normLedger.includes('non-current investment')) {
    return 'Investments';
  }
  if (
    normGroup.includes('non-current asset') ||
    normGroup.includes('misc. expenses (asset)') ||
    normLedger.includes('long-term loans and advances') ||
    normLedger.includes('other non-current asset')
  ) {
    return 'NonCurrentAssets';
  }
  if (
    normGroup.includes('stock-in-hand') ||
    normGroup.includes('inventory') ||
    normLedger.includes('inventori') ||
    normLedger.includes('closing stock') ||
    normLedger.includes('opening stock')
  ) {
    return 'Inventory';
  }
  if (
    normGroup.includes('sundry debtor') ||
    normGroup.includes('trade receivable') ||
    normGroup.includes('accounts receivable') ||
    normLedger.includes('trade receivable') ||
    normLedger.includes('sundry debtor')
  ) {
    return 'TradeReceivables';
  }
  if (
    normGroup.includes('cash-in-hand') ||
    normGroup === 'cash' ||
    normLedger === 'cash' ||
    normLedger === 'petty cash' ||
    normLedger.includes('cash on hand') ||
    normLedger.includes('cash and cash equivalent') ||
    normLedger.includes('cash & cash equivalent')
  ) {
    return 'CashAndEquivalents';
  }
  if (
    normGroup.includes('bank account') ||
    normGroup === 'bank' ||
    normLedger.includes('bank balance') ||
    normLedger.includes('hdfc bank') ||
    normLedger.includes('icici bank') ||
    normLedger.includes('sbi ') ||
    normLedger.includes('axis bank') ||
    normLedger.includes('kotak bank')
  ) {
    return 'BankBalances';
  }
  if (
    normGroup.includes('current asset') ||
    normGroup.includes('loans & advances (asset)') ||
    normGroup.includes('deposits (asset)') ||
    normLedger.includes('other current asset') ||
    normLedger.includes('short-term loans and advances') ||
    normLedger.includes('security deposit') ||
    normLedger.includes('prepaid') ||
    normLedger.includes('advance to') ||
    normLedger.includes('tds receivable') ||
    normLedger.includes('input igst') ||
    normLedger.includes('input cgst') ||
    normLedger.includes('input sgst')
  ) {
    if (
      normLedger.includes('igst') ||
      normLedger.includes('cgst') ||
      normLedger.includes('sgst') ||
      normLedger.includes('gst')
    ) {
      return 'TaxLiabilities';
    }
    return 'OtherCurrentAssets';
  }

  // Balance Sheet — Equity & Liabilities
  if (
    normGroup.includes('capital account') ||
    normGroup.includes('reserves & surplus') ||
    normGroup.includes('reserves and surplus') ||
    normGroup.includes('retained earnings') ||
    normGroup.includes('share capital') ||
    normGroup.includes('profit & loss a/c') ||
    normLedger.includes('share capital') ||
    normLedger.includes('reserves and surplus') ||
    normLedger.includes('reserves & surplus') ||
    normLedger.includes('profit & loss a/c')
  ) {
    return 'Equity';
  }

  if (
    normGroup.includes('bank od') ||
    normGroup.includes('bank occ') ||
    normGroup.includes('short-term borrowing') ||
    normLedger.includes('short-term borrowing') ||
    normLedger.includes('short term borrowing') ||
    normLedger.includes('working capital od') ||
    normLedger.includes('cash credit')
  ) {
    return 'ShortTermBorrowings';
  }

  if (
    normGroup.includes('loans (liability)') ||
    normGroup.includes('secured loan') ||
    normGroup.includes('unsecured loan') ||
    normGroup.includes('long-term borrowing') ||
    normLedger.includes('long-term borrowing') ||
    normLedger.includes('long term borrowing') ||
    normLedger.includes('other long-term liabilit') ||
    normLedger.includes('long-term provision') ||
    normLedger.includes('term loan')
  ) {
    return 'LongTermBorrowings';
  }

  if (
    normGroup.includes('sundry creditor') ||
    normGroup.includes('trade payable') ||
    normGroup.includes('accounts payable') ||
    normLedger.includes('trade payable') ||
    normLedger.includes('sundry creditor')
  ) {
    return 'TradePayables';
  }

  if (
    normGroup.includes('duties & taxes') ||
    normGroup.includes('duties and taxes') ||
    normLedger.includes('statutory dues') ||
    normLedger.includes('cgst') ||
    normLedger.includes('sgst') ||
    normLedger.includes('igst') ||
    normLedger.includes('tds payable') ||
    normLedger.includes('gst payable') ||
    normLedger.includes('output gst')
  ) {
    return 'TaxLiabilities';
  }

  if (
    normGroup.includes('current liabilit') ||
    normGroup.includes('provisions') ||
    normLedger.includes('other current liabilit') ||
    normLedger.includes('short-term provision') ||
    normLedger.includes('short term provision') ||
    normLedger.includes('payable') ||
    normLedger.includes('outstanding expense') ||
    normLedger.includes('advance from customer')
  ) {
    return 'OtherCurrentLiabilities';
  }

  // Fallback heuristics on ledger name when parent group is missing
  if (normLedger.includes('sales') || normLedger.includes('revenue')) return 'Sales';
  if (normLedger.includes('purchase')) return 'Purchases';
  if (normLedger.includes('debtor') || normLedger.includes('receivable')) return 'TradeReceivables';
  if (normLedger.includes('creditor') || normLedger.includes('vendor')) return 'TradePayables';
  if (normLedger.includes('capital') || normLedger.includes('reserve')) return 'Equity';
  if (
    normLedger.includes('machine') ||
    normLedger.includes('furniture') ||
    normLedger.includes('computer') ||
    normLedger.includes('vehicle') ||
    normLedger.includes('building')
  ) {
    return 'FixedAssets';
  }

  // If on P&L and under Expenses section, map unrecognized expense line to AdminExpenses (Other Expenses)
  if (statementContext === 'PNL') {
    if (normGroup.includes('income') || normGroup.includes('revenue')) return 'OtherIncome';
    if (normGroup.includes('tax')) return 'TaxExpense';
    if (normGroup.includes('expense')) return 'AdminExpenses';
  }

  return 'Unclassified';
}

export const CLASSIFICATION_LABELS: Record<AccountClassification, string> = {
  Sales: 'Revenue from Operations / Sales',
  SalesReturn: 'Sales Returns & Allowances',
  OtherOperatingIncome: 'Other Operating Income',
  OtherIncome: 'Other Income',
  Purchases: 'Purchases / Cost of Materials',
  PurchaseReturn: 'Purchase Returns',
  DirectExpenses: 'Direct / Manufacturing Expenses',
  EmployeeExpenses: 'Employee Benefits Expense',
  AdminExpenses: 'Other Expenses / Administrative & General',
  SellingExpenses: 'Selling & Distribution Expenses',
  FinanceCosts: 'Finance Costs',
  Depreciation: 'Depreciation & Amortisation',
  PriorPeriodExpenses: 'Prior Period & Exceptional Items',
  CurrentTax: 'Current Tax Expense',
  DeferredTax: 'Deferred Tax Charge / (Credit)',
  TaxExpense: 'Income Tax Expense',
  FixedAssets: 'Property, Plant & Equipment & Intangibles',
  Investments: 'Non-Current Investments',
  NonCurrentAssets: 'Deferred Tax Assets & Long-Term Loans/Advances',
  Inventory: 'Inventories / Stock-in-Hand',
  TradeReceivables: 'Trade Receivables (Sundry Debtors)',
  CashAndEquivalents: 'Cash and Cash Equivalents',
  BankBalances: 'Bank Balances',
  OtherCurrentAssets: 'Short-Term Loans, Advances & Other Current Assets',
  Equity: 'Share Capital, Reserves & Surplus',
  LongTermBorrowings: 'Long-Term Borrowings & Non-Current Liabilities',
  ShortTermBorrowings: 'Short-Term Borrowings & Bank OD/CC',
  TradePayables: 'Trade Payables (Sundry Creditors)',
  TaxLiabilities: 'Duties & Taxes (GST / TDS / Statutory)',
  OtherCurrentLiabilities: 'Other Current Liabilities & Short-Term Provisions',
  Suspense: 'Suspense Account',
  Unclassified: 'Unclassified / Unmapped Ledger',
};
