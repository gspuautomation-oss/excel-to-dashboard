import {
  DetectedReportType,
  GSTInvoice,
  GSTR3BSummary,
  ImportValidationMessage,
  toPaise,
} from '../types/accounting';
import { normalizeInvoiceNo } from './ledgerMapping';
import { parseTallyDate } from './xmlParser';

export interface ParsedGSTJsonResult {
  detectedReportType: DetectedReportType;
  gstin?: string;
  returnPeriod?: string;
  gstInvoices: GSTInvoice[];
  gstr3bSummaries: GSTR3BSummary[];
  validationMessages: ImportValidationMessage[];
}

function formatFpToMonth(fp: string | undefined): string {
  // GST portal fp is usually MMYYYY e.g. "042025" -> "2025-04"
  if (!fp) return '2025-04';
  const cleaned = fp.trim();
  if (/^\d{6}$/.test(cleaned)) {
    const mm = cleaned.slice(0, 2);
    const yyyy = cleaned.slice(2, 6);
    return `${yyyy}-${mm}`;
  }
  if (/^\d{4}-\d{2}$/.test(cleaned)) return cleaned;
  return '2025-04';
}

export function parseGSTJson(
  rawJson: string,
  sourceFileId: string,
  sourceFilename: string
): ParsedGSTJsonResult {
  const validationMessages: ImportValidationMessage[] = [];
  const gstInvoices: GSTInvoice[] = [];
  const gstr3bSummaries: GSTR3BSummary[] = [];

  let parsed: Record<string, any>;
  try {
    parsed = JSON.parse(rawJson);
  } catch (err) {
    validationMessages.push({
      severity: 'ERROR',
      code: 'INVALID_JSON',
      message: `Invalid JSON format in ${sourceFilename}.`,
    });
    return {
      detectedReportType: 'UNKNOWN_UNSUPPORTED',
      gstInvoices,
      gstr3bSummaries,
      validationMessages,
    };
  }

  const root = parsed.data || parsed;
  const fp = root.fp || root.ret_period || root.rtnprd || '042025';
  const returnPeriod = formatFpToMonth(String(fp));
  const companyGstin = root.gstin || '';

  // 1. Check GSTR-3B JSON structure (sup_details & itc_elg)
  if (root.sup_details || root.itc_elg || sourceFilename.toLowerCase().includes('3b')) {
    const osup = root.sup_details?.osup_det || {};
    const isup = root.sup_details?.isup_rev || {};
    const itcAll = Array.isArray(root.itc_elg?.itc_avl) ? root.itc_elg.itc_avl : [];
    const itcNet = root.itc_elg?.itc_net || {};
    const itcRev = Array.isArray(root.itc_elg?.itc_rev) ? root.itc_elg.itc_rev : [];
    const itcInelg = Array.isArray(root.itc_elg?.itc_inelg) ? root.itc_elg.itc_inelg : [];

    let eligIgst = toPaise(itcNet.iamt);
    let eligCgst = toPaise(itcNet.camt);
    let eligSgst = toPaise(itcNet.samt);
    let eligCess = toPaise(itcNet.csamt);

    if (eligIgst === 0 && eligCgst === 0 && eligSgst === 0 && itcAll.length > 0) {
      for (const row of itcAll) {
        eligIgst += toPaise(row.iamt);
        eligCgst += toPaise(row.camt);
        eligSgst += toPaise(row.samt);
        eligCess += toPaise(row.csamt);
      }
    }

    let revTotal = 0;
    for (const r of itcRev) {
      revTotal += toPaise(r.iamt) + toPaise(r.camt) + toPaise(r.samt);
    }
    let inelgTotal = 0;
    for (const r of itcInelg) {
      inelgTotal += toPaise(r.iamt) + toPaise(r.camt) + toPaise(r.samt);
    }

    gstr3bSummaries.push({
      id: `${sourceFileId}_3B`,
      returnPeriod,
      outwardTaxableValuePaise: toPaise(osup.txval),
      outwardIgstPaise: toPaise(osup.iamt),
      outwardCgstPaise: toPaise(osup.camt),
      outwardSgstPaise: toPaise(osup.samt),
      outwardCessPaise: toPaise(osup.csamt),
      rcmTaxableValuePaise: toPaise(isup.txval),
      rcmIgstPaise: toPaise(isup.iamt),
      rcmCgstPaise: toPaise(isup.camt),
      rcmSgstPaise: toPaise(isup.samt),
      eligibleItcIgstPaise: eligIgst,
      eligibleItcCgstPaise: eligCgst,
      eligibleItcSgstPaise: eligSgst,
      eligibleItcCessPaise: eligCess,
      ineligibleItcPaise: inelgTotal,
      reversedItcPaise: revTotal,
      sourceFileId,
      sourceFilename,
    });

    return {
      detectedReportType: 'GSTR3B_RETURN',
      gstin: companyGstin,
      returnPeriod,
      gstInvoices,
      gstr3bSummaries,
      validationMessages,
    };
  }

  // 2. Check GSTR-2B JSON structure (docdata.b2b or b2b with supplier trdnm / itcavl)
  const docdata = root.docdata || root;
  const is2B =
    Boolean(root.docdata) ||
    sourceFilename.toLowerCase().includes('2b') ||
    (Array.isArray(docdata.b2b) && docdata.b2b.some((p: any) => p.trdnm || p.supprd));

  if (Array.isArray(docdata.b2b)) {
    let counter = 0;
    for (const partyEntry of docdata.b2b) {
      const ctin = String(partyEntry.ctin || '').toUpperCase();
      const partyName = partyEntry.trdnm || partyEntry.cname || ctin || 'Registered Party';
      const invList = partyEntry.inv || [];

      for (const inv of invList) {
        counter++;
        const invNo = String(inv.inum || `INV-${counter}`);
        const invDate = parseTallyDate(inv.dt || inv.idt) || `${returnPeriod}-15`;
        const pos = inv.pos || '';
        const rev = String(inv.rev || 'N').toUpperCase() === 'Y';
        const itcAvl = String(inv.itcavl || 'Y').toUpperCase();

        let taxablePaise = toPaise(inv.txval);
        let igstPaise = toPaise(inv.igst || inv.iamt);
        let cgstPaise = toPaise(inv.cgst || inv.camt);
        let sgstPaise = toPaise(inv.sgst || inv.samt);
        let cessPaise = toPaise(inv.cess || inv.csamt);
        let rate = Number(inv.rt) || undefined;

        if (Array.isArray(inv.itms)) {
          taxablePaise = 0;
          igstPaise = 0;
          cgstPaise = 0;
          sgstPaise = 0;
          cessPaise = 0;
          for (const itm of inv.itms) {
            const det = itm.itm_det || itm;
            taxablePaise += toPaise(det.txval);
            igstPaise += toPaise(det.iamt || det.igst);
            cgstPaise += toPaise(det.camt || det.cgst);
            sgstPaise += toPaise(det.samt || det.sgst);
            cessPaise += toPaise(det.csamt || det.cess);
            if (!rate && det.rt) rate = Number(det.rt);
          }
        }

        const totalTaxPaise = igstPaise + cgstPaise + sgstPaise + cessPaise;

        gstInvoices.push({
          id: `${sourceFileId}_GST_${counter}`,
          source: is2B ? 'GSTR2B' : 'GSTR1',
          documentType: 'INVOICE',
          partyGstin: ctin,
          partyName,
          invoiceNumber: invNo,
          normalizedInvoiceNumber: normalizeInvoiceNo(invNo),
          invoiceDate: invDate,
          returnPeriod,
          supplyType: 'B2B',
          placeOfSupply: pos || undefined,
          reverseCharge: rev,
          itcEligibility: is2B ? (itcAvl === 'N' ? 'INELIGIBLE' : 'ELIGIBLE') : undefined,
          taxRate: rate,
          taxableValuePaise: taxablePaise,
          igstPaise,
          cgstPaise,
          sgstPaise,
          cessPaise,
          totalTaxPaise,
          invoiceTotalPaise: taxablePaise + totalTaxPaise,
          sourceFileId,
          sourceFilename,
        });
      }
    }
  }

  // Parse CDNR (Credit / Debit Notes) in GSTR-1 or GSTR-2B
  if (Array.isArray(docdata.cdnr)) {
    let cdnrCount = 0;
    for (const partyEntry of docdata.cdnr) {
      const ctin = String(partyEntry.ctin || '').toUpperCase();
      const partyName = partyEntry.trdnm || partyEntry.cname || ctin || 'Registered Party';
      const ntList = partyEntry.nt || [];
      for (const nt of ntList) {
        cdnrCount++;
        const ntNum = String(nt.ntnum || nt.nt_num || `CN-${cdnrCount}`);
        const ntDate = parseTallyDate(nt.dt || nt.nt_dt) || `${returnPeriod}-20`;
        const isDebit = String(nt.ntty || nt.typ || 'C').toUpperCase() === 'D';

        let taxablePaise = toPaise(nt.txval);
        let igstPaise = toPaise(nt.igst || nt.iamt);
        let cgstPaise = toPaise(nt.cgst || nt.camt);
        let sgstPaise = toPaise(nt.sgst || nt.samt);
        let cessPaise = toPaise(nt.cess || nt.csamt);

        if (Array.isArray(nt.itms)) {
          taxablePaise = 0;
          igstPaise = 0;
          cgstPaise = 0;
          sgstPaise = 0;
          cessPaise = 0;
          for (const itm of nt.itms) {
            const det = itm.itm_det || itm;
            taxablePaise += toPaise(det.txval);
            igstPaise += toPaise(det.iamt || det.igst);
            cgstPaise += toPaise(det.camt || det.cgst);
            sgstPaise += toPaise(det.samt || det.sgst);
            cessPaise += toPaise(det.csamt || det.cess);
          }
        }

        const totalTaxPaise = igstPaise + cgstPaise + sgstPaise + cessPaise;
        gstInvoices.push({
          id: `${sourceFileId}_CDNR_${cdnrCount}`,
          source: is2B ? 'GSTR2B' : 'GSTR1',
          documentType: isDebit ? 'DEBIT_NOTE' : 'CREDIT_NOTE',
          partyGstin: ctin,
          partyName,
          invoiceNumber: ntNum,
          normalizedInvoiceNumber: normalizeInvoiceNo(ntNum),
          invoiceDate: ntDate,
          returnPeriod,
          supplyType: 'CDNR',
          reverseCharge: false,
          itcEligibility: is2B ? 'ELIGIBLE' : undefined,
          taxableValuePaise: taxablePaise,
          igstPaise,
          cgstPaise,
          sgstPaise,
          cessPaise,
          totalTaxPaise,
          invoiceTotalPaise: taxablePaise + totalTaxPaise,
          sourceFileId,
          sourceFilename,
        });
      }
    }
  }

  return {
    detectedReportType: is2B ? 'GSTR2B_RETURN' : gstInvoices.length > 0 ? 'GSTR1_RETURN' : 'UNKNOWN_UNSUPPORTED',
    gstin: companyGstin,
    returnPeriod,
    gstInvoices,
    gstr3bSummaries,
    validationMessages,
  };
}
