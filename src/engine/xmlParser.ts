import {
  BillAllocation,
  DetectedReportType,
  GSTInvoice,
  ImportValidationMessage,
  InventoryMovement,
  Ledger,
  LedgerGroup,
  OutstandingInvoice,
  StockItem,
  toPaise,
  Voucher,
  VoucherEntry,
  VoucherType,
} from '../types/accounting';
import { classifyLedgerAndGroup, normalizeInvoiceNo, normalizeName } from './ledgerMapping';

export interface ParsedTallyXMLResult {
  detectedReportType: DetectedReportType;
  extractedCompanyName?: string;
  periodStart?: string;
  periodEnd?: string;
  groups: LedgerGroup[];
  ledgers: Ledger[];
  vouchers: Voucher[];
  outstandingInvoices: OutstandingInvoice[];
  stockItems: StockItem[];
  gstInvoices: GSTInvoice[];
  validationMessages: ImportValidationMessage[];
}

export function sanitizeXMLContent(rawXml: string): string {
  // Prevent XXE / unsafe DOCTYPE / external entity expansion
  return rawXml
    .replace(/<!DOCTYPE[^>]*(\[[\s\S]*?\])?>/gi, '')
    .replace(/<!ENTITY[^>]*>/gi, '')
    // Strip invalid control characters (0x00-0x08, 0x0B-0x0C, 0x0E-0x1F) often exported by Tally
    .replace(/[\x00-\x08\x0B\x0C\x0E-\x1F]/g, '')
    .replace(/&#x[0-9a-fA-F]+;/g, '');
}

export function parseTallyDate(raw: string | null | undefined): string {
  if (!raw) return '';
  const cleaned = raw.trim();
  // Tally standard YYYYMMDD (e.g. 20250415)
  if (/^\d{8}$/.test(cleaned)) {
    return `${cleaned.slice(0, 4)}-${cleaned.slice(4, 6)}-${cleaned.slice(6, 8)}`;
  }
  // YYYY-MM-DD
  if (/^\d{4}-\d{2}-\d{2}$/.test(cleaned)) {
    return cleaned;
  }
  // DD-MM-YYYY or DD/MM/YYYY
  const dmyMatch = cleaned.match(/^(\d{1,2})[-/.](\d{1,2})[-/.](\d{4})$/);
  if (dmyMatch) {
    const d = dmyMatch[1].padStart(2, '0');
    const m = dmyMatch[2].padStart(2, '0');
    const y = dmyMatch[3];
    return `${y}-${m}-${d}`;
  }
  // DD-MMM-YYYY or D-MMM-YY (e.g., 15-Apr-2025 or 1-Apr-25)
  const parsed = new Date(cleaned);
  if (!Number.isNaN(parsed.getTime())) {
    const y = parsed.getFullYear();
    const m = String(parsed.getMonth() + 1).padStart(2, '0');
    const d = String(parsed.getDate()).padStart(2, '0');
    return `${y}-${m}-${d}`;
  }
  return '';
}

function getChildText(parent: Element, tagName: string): string {
  for (let i = 0; i < parent.children.length; i++) {
    const child = parent.children[i];
    if (child.tagName.toUpperCase() === tagName.toUpperCase()) {
      return (child.textContent || '').trim();
    }
  }
  const el = parent.getElementsByTagName(tagName)[0];
  return el ? (el.textContent || '').trim() : '';
}

/**
 * TallyPrime Amount Sign Convention:
 * In Tally XML exports:
 * - Negative numbers in <AMOUNT> or <OPENINGBALANCE> orISDEEMEDPOSITIVE="Yes" represent DEBIT (+Paise in our internal model).
 * - Positive numbers in <AMOUNT> with ISDEEMEDPOSITIVE="No" represent CREDIT (-Paise in our internal model).
 */
export function parseTallySignedPaise(rawAmount: string, isDeemedPositiveTag?: string): number {
  if (!rawAmount) return 0;
  const cleaned = rawAmount
    .replace(/,/g, '')
    .replace(/₹/g, '')
    .trim();
  const numMatch = cleaned.match(/[-+]?\d+(\.\d+)?/);
  if (!numMatch) return 0;
  const numericVal = Number(numMatch[0]);
  if (!Number.isFinite(numericVal) || numericVal === 0) return 0;

  const absPaise = Math.round(Math.abs(numericVal) * 100);
  const deemedPos = (isDeemedPositiveTag || '').trim().toUpperCase();

  if (deemedPos === 'YES') {
    return absPaise; // Debit
  }
  if (deemedPos === 'NO') {
    return -absPaise; // Credit
  }
  // When ISDEEMEDPOSITIVE is not present, Tally uses negative sign for Debit and positive for Credit
  return numericVal < 0 ? absPaise : -absPaise;
}

export function normalizeVoucherType(rawType: string): VoucherType {
  const norm = rawType.trim().toLowerCase();
  if (norm.includes('credit note') || norm.includes('sales return')) return 'Credit Note';
  if (norm.includes('debit note') || norm.includes('purchase return')) return 'Debit Note';
  if (norm.includes('sale') || norm.includes('tax invoice')) return 'Sales';
  if (norm.includes('purch')) return 'Purchase';
  if (norm.includes('rcpt') || norm.includes('receipt')) return 'Receipt';
  if (norm.includes('pymt') || norm.includes('payment')) return 'Payment';
  if (norm.includes('cntr') || norm.includes('contra')) return 'Contra';
  if (norm.includes('jrnl') || norm.includes('journal')) return 'Journal';
  return 'Other';
}

export function parseTallyXML(
  rawXml: string,
  sourceFileId: string,
  sourceFilename: string
): ParsedTallyXMLResult {
  const validationMessages: ImportValidationMessage[] = [];
  const sanitized = sanitizeXMLContent(rawXml);

  if (rawXml.includes('<!ENTITY') || rawXml.includes('<!DOCTYPE')) {
    validationMessages.push({
      severity: 'WARNING',
      code: 'XML_DOCTYPE_STRIPPED',
      message: 'Unsafe XML DOCTYPE/ENTITY declarations were sanitized prior to parsing.',
    });
  }

  const parser = new DOMParser();
  const xmlDoc = parser.parseFromString(sanitized, 'text/xml');
  const parserError = xmlDoc.getElementsByTagName('parsererror')[0];

  if (parserError) {
    validationMessages.push({
      severity: 'ERROR',
      code: 'MALFORMED_XML',
      message: 'Malformed XML structure detected.',
      details: (parserError.textContent || '').slice(0, 220),
    });
    return {
      detectedReportType: 'UNKNOWN_UNSUPPORTED',
      groups: [],
      ledgers: [],
      vouchers: [],
      outstandingInvoices: [],
      stockItems: [],
      gstInvoices: [],
      validationMessages,
    };
  }

  // Extract company name
  let extractedCompanyName =
    getChildText(xmlDoc.documentElement, 'SVCURRENTCOMPANY') ||
    getChildText(xmlDoc.documentElement, 'COMPANYNAME') ||
    '';
  if (!extractedCompanyName) {
    const companyEl = xmlDoc.getElementsByTagName('COMPANY')[0] || xmlDoc.getElementsByTagName('REMOTECMPINFO.LIST')[0];
    if (companyEl) {
      extractedCompanyName = companyEl.getAttribute('NAME') || getChildText(companyEl, 'NAME') || '';
    }
  }

  const groups: LedgerGroup[] = [];
  const ledgers: Ledger[] = [];
  const vouchers: Voucher[] = [];
  const outstandingInvoices: OutstandingInvoice[] = [];
  const stockItems: StockItem[] = [];
  const gstInvoices: GSTInvoice[] = [];

  // 1. Parse GROUP elements
  const groupElements = Array.from(xmlDoc.getElementsByTagName('GROUP'));
  for (const grpEl of groupElements) {
    const name = grpEl.getAttribute('NAME') || getChildText(grpEl, 'NAME');
    if (!name) continue;
    const parentGroup = getChildText(grpEl, 'PARENT') || 'Primary';
    const isRevenue = getChildText(grpEl, 'ISREVENUE').toUpperCase() === 'YES';
    const isDeemedPositive = getChildText(grpEl, 'ISDEEMEDPOSITIVE').toUpperCase() === 'YES';
    groups.push({
      name: name.trim(),
      parentGroup: parentGroup.trim(),
      classification: classifyLedgerAndGroup(name, parentGroup),
      isRevenue,
      isDeemedPositive,
      sourceFileId,
    });
  }

  // 2. Parse LEDGER elements
  const ledgerElements = Array.from(xmlDoc.getElementsByTagName('LEDGER'));
  for (const ledEl of ledgerElements) {
    const name = ledEl.getAttribute('NAME') || getChildText(ledEl, 'NAME');
    if (!name) continue;
    const parentGroup = getChildText(ledEl, 'PARENT') || 'Unclassified';
    const rawOpening = getChildText(ledEl, 'OPENINGBALANCE');
    const rawClosing = getChildText(ledEl, 'CLOSINGBALANCE');
    const isDeemedPositive = getChildText(ledEl, 'ISDEEMEDPOSITIVE');
    const gstin = getChildText(ledEl, 'PARTYGSTIN') || getChildText(ledEl, 'GSTIN');
    const stateName = getChildText(ledEl, 'LEDSTATENAME') || getChildText(ledEl, 'STATENAME');
    const guid = getChildText(ledEl, 'GUID') || getChildText(ledEl, 'MASTERID');

    const openingBalancePaise = parseTallySignedPaise(rawOpening, isDeemedPositive);
    const closingBalanceReportedPaise = rawClosing
      ? parseTallySignedPaise(rawClosing, isDeemedPositive)
      : undefined;

    ledgers.push({
      id: `${sourceFileId}_LED_${normalizeName(name)}`,
      name: name.trim(),
      normalizedName: normalizeName(name),
      parentGroup: parentGroup.trim(),
      classification: classifyLedgerAndGroup(name, parentGroup),
      gstin: gstin || undefined,
      stateName: stateName || undefined,
      openingBalancePaise,
      closingBalanceReportedPaise,
      sourceFileIds: [sourceFileId],
      masterGuid: guid || undefined,
    });
  }

  // 3. Parse STOCKITEM elements
  const stockElements = Array.from(xmlDoc.getElementsByTagName('STOCKITEM'));
  for (const stEl of stockElements) {
    const name = stEl.getAttribute('NAME') || getChildText(stEl, 'NAME');
    if (!name) continue;
    const parentGroup = getChildText(stEl, 'PARENT') || 'Primary';
    const unit = getChildText(stEl, 'BASEUNITS') || 'Nos';
    const hsnCode = getChildText(stEl, 'HSNCODE');
    const openingValuePaise = Math.abs(parseTallySignedPaise(getChildText(stEl, 'OPENINGVALUE')));
    const closingValuePaise = Math.abs(parseTallySignedPaise(getChildText(stEl, 'CLOSINGVALUE')));
    const openingQty = parseFloat(getChildText(stEl, 'OPENINGBALANCE')) || 0;
    const closingQty = parseFloat(getChildText(stEl, 'CLOSINGBALANCE')) || 0;

    stockItems.push({
      name: name.trim(),
      parentGroup,
      unit,
      hsnCode: hsnCode || undefined,
      openingQty,
      openingValuePaise,
      closingQty,
      closingValuePaise,
      sourceFileId,
    });
  }

  // 4. Parse VOUCHER elements
  const voucherElements = Array.from(xmlDoc.getElementsByTagName('VOUCHER'));
  let minDate = '';
  let maxDate = '';

  for (let idx = 0; idx < voucherElements.length; idx++) {
    const vchEl = voucherElements[idx];
    const isCancelled = getChildText(vchEl, 'ISCANCELLED').toUpperCase() === 'YES';
    const isOptional = getChildText(vchEl, 'ISOPTIONAL').toUpperCase() === 'YES';
    if (isCancelled || isOptional) continue;

    const rawDate = getChildText(vchEl, 'DATE') || getChildText(vchEl, 'EFFECTIVEDATE');
    const date = parseTallyDate(rawDate);
    if (!date) {
      validationMessages.push({
        severity: 'WARNING',
        code: 'INVALID_VOUCHER_DATE',
        message: `Voucher at index ${idx + 1} has invalid or missing date (${rawDate || 'empty'}).`,
      });
      continue;
    }

    if (!minDate || date < minDate) minDate = date;
    if (!maxDate || date > maxDate) maxDate = date;

    const rawVchType = vchEl.getAttribute('VCHTYPE') || getChildText(vchEl, 'VOUCHERTYPENAME') || 'Journal';
    const voucherType = normalizeVoucherType(rawVchType);
    const voucherNumber = getChildText(vchEl, 'VOUCHERNUMBER') || '';
    const guid = getChildText(vchEl, 'GUID') || vchEl.getAttribute('REMOTEID') || undefined;
    const masterId = getChildText(vchEl, 'MASTERID') || vchEl.getAttribute('VCHKEY') || undefined;
    const partyName = getChildText(vchEl, 'PARTYLEDGERNAME') || getChildText(vchEl, 'PARTYNAME') || '';
    const partyGstin = getChildText(vchEl, 'PARTYGSTIN') || getChildText(vchEl, 'CONSIGNEEGSTIN') || '';
    const placeOfSupply = getChildText(vchEl, 'PLACEOFSUPPLY') || getChildText(vchEl, 'STATENAME') || '';
    const narration = getChildText(vchEl, 'NARRATION') || '';
    const referenceNumber = getChildText(vchEl, 'REFERENCE') || '';
    const referenceDate = parseTallyDate(getChildText(vchEl, 'REFERENCEDATE'));

    // Collect ledger entries from ALLLEDGERENTRIES.LIST or LEDGERENTRIES.LIST
    const entryLists = [
      ...Array.from(vchEl.getElementsByTagName('ALLLEDGERENTRIES.LIST')),
      ...Array.from(vchEl.getElementsByTagName('LEDGERENTRIES.LIST')),
    ];

    const entries: VoucherEntry[] = [];
    for (const entEl of entryLists) {
      const ledgerName = getChildText(entEl, 'LEDGERNAME');
      if (!ledgerName) continue;
      const rawAmt = getChildText(entEl, 'AMOUNT');
      const isDeemedPositive = getChildText(entEl, 'ISDEEMEDPOSITIVE');
      const amountPaise = parseTallySignedPaise(rawAmt, isDeemedPositive);
      if (amountPaise === 0) continue;

      const isPartyLedger = getChildText(entEl, 'ISPARTYLEDGER').toUpperCase() === 'YES';

      // Extract BILLALLOCATIONS.LIST
      const billAllocations: BillAllocation[] = [];
      const billEls = Array.from(entEl.getElementsByTagName('BILLALLOCATIONS.LIST'));
      for (const bEl of billEls) {
        const billName = getChildText(bEl, 'NAME') || voucherNumber;
        const rawBillType = getChildText(bEl, 'BILLTYPE') || 'New Ref';
        const billType: BillAllocation['billType'] =
          rawBillType === 'Agst Ref'
            ? 'Agst Ref'
            : rawBillType === 'Advance'
            ? 'Advance'
            : rawBillType === 'On Account'
            ? 'On Account'
            : 'New Ref';
        const billAmtPaise = parseTallySignedPaise(
          getChildText(bEl, 'AMOUNT'),
          getChildText(bEl, 'ISDEEMEDPOSITIVE') || isDeemedPositive
        );
        const creditPeriodStr = getChildText(bEl, 'BILLCREDITPERIOD');
        let creditDays = 0;
        let dueDate: string | undefined;
        if (creditPeriodStr) {
          const daysMatch = creditPeriodStr.match(/^(\d+)\s*Days?/i);
          if (daysMatch) {
            creditDays = parseInt(daysMatch[1], 10);
            const dt = new Date(date);
            dt.setDate(dt.getDate() + creditDays);
            dueDate = dt.toISOString().slice(0, 10);
          } else {
            const parsedDue = parseTallyDate(creditPeriodStr);
            if (parsedDue) dueDate = parsedDue;
          }
        }
        if (billName && billAmtPaise !== 0) {
          billAllocations.push({
            billName,
            billType,
            amountPaise: billAmtPaise,
            dueDate,
            creditDays: creditDays || undefined,
          });
        }
      }

      // Extract Cost Centre if present
      const ccEl = entEl.getElementsByTagName('COSTCENTREALLOCATIONS.LIST')[0];
      const costCentre = ccEl ? getChildText(ccEl, 'NAME') : undefined;

      entries.push({
        ledgerName: ledgerName.trim(),
        normalizedLedgerName: normalizeName(ledgerName),
        amountPaise,
        isPartyLedger,
        billAllocations,
        costCentre,
      });
    }

    // Extract Inventory entries
    const invLists = [
      ...Array.from(vchEl.getElementsByTagName('ALLINVENTORYENTRIES.LIST')),
      ...Array.from(vchEl.getElementsByTagName('INVENTORYENTRIES.LIST')),
    ];
    const inventoryMovements: InventoryMovement[] = [];
    for (const invEl of invLists) {
      const stockItemName = getChildText(invEl, 'STOCKITEMNAME');
      if (!stockItemName) continue;
      const qtyRaw = getChildText(invEl, 'ACTUALQTY') || getChildText(invEl, 'BILLEDQTY');
      const qtyNum = parseFloat(qtyRaw) || 0;
      const unitMatch = qtyRaw.replace(/[-+0-9.,\s]/g, '').trim() || 'Nos';
      const ratePaise = toPaise(getChildText(invEl, 'RATE').split('/')[0]);
      const amountPaise = parseTallySignedPaise(
        getChildText(invEl, 'AMOUNT'),
        getChildText(invEl, 'ISDEEMEDPOSITIVE')
      );
      inventoryMovements.push({
        stockItemName,
        quantity: qtyNum,
        unit: unitMatch,
        ratePaise,
        amountPaise,
      });
    }

    let totalDebitPaise = 0;
    let totalCreditPaise = 0;
    for (const entry of entries) {
      if (entry.amountPaise > 0) {
        totalDebitPaise += entry.amountPaise;
      } else {
        totalCreditPaise += Math.abs(entry.amountPaise);
      }
    }

    const voucherId = `${sourceFileId}_VCH_${idx + 1}`;
    const voucher: Voucher = {
      id: voucherId,
      sourceRecordId: guid || masterId || `${rawVchType}_${voucherNumber}_${date}_${idx + 1}`,
      guid,
      masterId,
      date,
      voucherNumber,
      voucherType,
      originalVoucherType: rawVchType,
      partyName: partyName || undefined,
      partyGstin: partyGstin || undefined,
      placeOfSupply: placeOfSupply || undefined,
      narration: narration || undefined,
      referenceNumber: referenceNumber || undefined,
      referenceDate: referenceDate || undefined,
      entries,
      inventoryMovements,
      totalDebitPaise,
      totalCreditPaise,
      isBalanced: Math.abs(totalDebitPaise - totalCreditPaise) <= 100,
      sourceFileId,
      sourceFilename,
    };

    vouchers.push(voucher);

    // Extract GST Invoice from Sales / Purchase / Credit Note / Debit Note vouchers
    if (
      voucherType === 'Sales' ||
      voucherType === 'Purchase' ||
      voucherType === 'Credit Note' ||
      voucherType === 'Debit Note'
    ) {
      let taxableValuePaise = 0;
      let igstPaise = 0;
      let cgstPaise = 0;
      let sgstPaise = 0;
      let cessPaise = 0;

      for (const ent of entries) {
        const lName = ent.normalizedLedgerName;
        const absAmt = Math.abs(ent.amountPaise);
        if (lName.includes('igst')) {
          igstPaise += absAmt;
        } else if (lName.includes('cgst')) {
          cgstPaise += absAmt;
        } else if (lName.includes('sgst') || lName.includes('utgst')) {
          sgstPaise += absAmt;
        } else if (lName.includes('cess')) {
          cessPaise += absAmt;
        } else if (!ent.isPartyLedger && normalizeName(partyName) !== lName) {
          taxableValuePaise += absAmt;
        }
      }

      if (taxableValuePaise > 0 || igstPaise + cgstPaise + sgstPaise > 0) {
        const isOutward = voucherType === 'Sales' || voucherType === 'Credit Note';
        const docType =
          voucherType === 'Credit Note'
            ? 'CREDIT_NOTE'
            : voucherType === 'Debit Note'
            ? 'DEBIT_NOTE'
            : 'INVOICE';
        const invNo = (!isOutward && referenceNumber ? referenceNumber : voucherNumber) || `VCH-${idx + 1}`;
        const totalTaxPaise = igstPaise + cgstPaise + sgstPaise + cessPaise;

        gstInvoices.push({
          id: `${voucherId}_GST`,
          source: isOutward ? 'TALLY_SALES' : 'TALLY_PURCHASE',
          documentType: docType,
          partyGstin: partyGstin.trim().toUpperCase(),
          partyName: partyName || 'Unregistered / Cash Party',
          invoiceNumber: invNo,
          normalizedInvoiceNumber: normalizeInvoiceNo(invNo),
          invoiceDate: date,
          returnPeriod: date.slice(0, 7),
          supplyType: partyGstin ? 'B2B' : 'B2C',
          placeOfSupply: placeOfSupply || undefined,
          reverseCharge: false,
          itcEligibility: isOutward ? undefined : 'ELIGIBLE',
          taxableValuePaise,
          igstPaise,
          cgstPaise,
          sgstPaise,
          cessPaise,
          totalTaxPaise,
          invoiceTotalPaise: taxableValuePaise + totalTaxPaise,
          sourceFileId,
          sourceFilename,
          linkedVoucherId: voucherId,
        });
      }
    }
  }

  // 5. Parse dedicated BILLFIXED / OUTSTANDING XML nodes if exported from Tally Outstanding Report
  const billFixedEls = Array.from(xmlDoc.getElementsByTagName('BILLFIXED'));
  for (let i = 0; i < billFixedEls.length; i++) {
    const bf = billFixedEls[i];
    const billDate = parseTallyDate(getChildText(bf, 'BILLDATE'));
    const billRef = getChildText(bf, 'BILLREF');
    const billParty = getChildText(bf, 'BILLPARTY');
    const parentElem = bf.parentElement;
    if (!billRef || !billParty || !parentElem) continue;

    const clRaw = getChildText(parentElem, 'BILLCL');
    const dueRaw = parseTallyDate(getChildText(parentElem, 'BILLDUEDATE')) || billDate;
    const signedPaise = parseTallySignedPaise(clRaw);
    if (signedPaise === 0) continue;

    const category = signedPaise > 0 ? 'RECEIVABLE' : 'PAYABLE';
    const pendingAmountPaise = Math.abs(signedPaise);

    outstandingInvoices.push({
      id: `${sourceFileId}_OUT_${i + 1}`,
      partyName: billParty,
      normalizedPartyName: normalizeName(billParty),
      category,
      billNumber: billRef,
      billDate: billDate || dueRaw || '2025-04-01',
      dueDate: dueRaw || billDate || '2025-04-30',
      originalAmountPaise: pendingAmountPaise,
      adjustedAmountPaise: 0,
      pendingAmountPaise,
      daysOverdue: 0,
      invoiceAgeDays: 0,
      relatedVoucherIds: [],
      sourceFileId,
      sourceFilename,
      isFromBillAllocation: true,
    });
  }

  let detectedReportType: DetectedReportType = 'UNKNOWN_UNSUPPORTED';
  if (vouchers.length > 0 && ledgers.length > 0) {
    detectedReportType = 'TALLY_COMBINED_XML';
  } else if (vouchers.length > 0) {
    detectedReportType = 'TALLY_TRANSACTIONS_XML';
  } else if (ledgers.length > 0 || groups.length > 0 || stockItems.length > 0) {
    detectedReportType = 'TALLY_MASTERS_XML';
  } else if (outstandingInvoices.length > 0) {
    detectedReportType = 'TALLY_OUTSTANDING_XML';
  } else {
    validationMessages.push({
      severity: 'ERROR',
      code: 'NO_TALLY_RECORDS_FOUND',
      message: 'No TallyPrime VOUCHER, LEDGER, GROUP, STOCKITEM, or BILLFIXED records were found in this XML file.',
    });
  }

  return {
    detectedReportType,
    extractedCompanyName: extractedCompanyName || undefined,
    periodStart: minDate || undefined,
    periodEnd: maxDate || undefined,
    groups,
    ledgers,
    vouchers,
    outstandingInvoices,
    stockItems,
    gstInvoices,
    validationMessages,
  };
}
