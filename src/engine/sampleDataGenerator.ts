import * as XLSX from 'xlsx';
import { toPaise } from '../types/accounting';
import { parseGSTJson } from './gstParser';
import { RawFilePayload } from './reconciliationEngine';
import { parseTallyXML } from './xmlParser';

export const DEMO_COMPANY_NAME = 'DEMO — Bharat Precision Engineering Pvt Ltd (Test Dataset)';

export function getSampleTallyMastersXML(): string {
  return `<?xml version="1.0" encoding="UTF-8"?>
<ENVELOPE>
  <HEADER>
    <TALLYREQUEST>Export Data</TALLYREQUEST>
  </HEADER>
  <BODY>
    <IMPORTDATA>
      <REQUESTDESC>
        <REPORTNAME>All Masters</REPORTNAME>
        <STATICVARIABLES>
          <SVCURRENTCOMPANY>${DEMO_COMPANY_NAME}</SVCURRENTCOMPANY>
        </STATICVARIABLES>
      </REQUESTDESC>
      <REQUESTDATA>
        <TALLYMESSAGE xmlns:UDF="TallyUDF">
          <LEDGER NAME="Share Capital">
            <NAME>Share Capital</NAME>
            <PARENT>Capital Account</PARENT>
            <ISDEEMEDPOSITIVE>No</ISDEEMEDPOSITIVE>
            <OPENINGBALANCE>5000000.00</OPENINGBALANCE>
          </LEDGER>
          <LEDGER NAME="General Reserves &amp; Surplus">
            <NAME>General Reserves &amp; Surplus</NAME>
            <PARENT>Reserves &amp; Surplus</PARENT>
            <ISDEEMEDPOSITIVE>No</ISDEEMEDPOSITIVE>
            <OPENINGBALANCE>3200000.00</OPENINGBALANCE>
          </LEDGER>
          <LEDGER NAME="SBI Term Loan A/c">
            <NAME>SBI Term Loan A/c</NAME>
            <PARENT>Secured Loans</PARENT>
            <ISDEEMEDPOSITIVE>No</ISDEEMEDPOSITIVE>
            <OPENINGBALANCE>2800000.00</OPENINGBALANCE>
          </LEDGER>
          <LEDGER NAME="HDFC Cash Credit OD A/c">
            <NAME>HDFC Cash Credit OD A/c</NAME>
            <PARENT>Bank OD A/c</PARENT>
            <ISDEEMEDPOSITIVE>No</ISDEEMEDPOSITIVE>
            <OPENINGBALANCE>950000.00</OPENINGBALANCE>
          </LEDGER>
          <LEDGER NAME="CNC Machining Centres &amp; Plant">
            <NAME>CNC Machining Centres &amp; Plant</NAME>
            <PARENT>Fixed Assets</PARENT>
            <ISDEEMEDPOSITIVE>Yes</ISDEEMEDPOSITIVE>
            <OPENINGBALANCE>-6400000.00</OPENINGBALANCE>
          </LEDGER>
          <LEDGER NAME="Industrial Bonds Investment">
            <NAME>Industrial Bonds Investment</NAME>
            <PARENT>Investments</PARENT>
            <ISDEEMEDPOSITIVE>Yes</ISDEEMEDPOSITIVE>
            <OPENINGBALANCE>-800000.00</OPENINGBALANCE>
          </LEDGER>
          <LEDGER NAME="HDFC Bank Current A/c">
            <NAME>HDFC Bank Current A/c</NAME>
            <PARENT>Bank Accounts</PARENT>
            <ISDEEMEDPOSITIVE>Yes</ISDEEMEDPOSITIVE>
            <OPENINGBALANCE>-2450000.00</OPENINGBALANCE>
          </LEDGER>
          <LEDGER NAME="Cash">
            <NAME>Cash</NAME>
            <PARENT>Cash-in-Hand</PARENT>
            <ISDEEMEDPOSITIVE>Yes</ISDEEMEDPOSITIVE>
            <OPENINGBALANCE>-180000.00</OPENINGBALANCE>
          </LEDGER>
          <LEDGER NAME="Security Deposits &amp; Prepaid">
            <NAME>Security Deposits &amp; Prepaid</NAME>
            <PARENT>Deposits (Asset)</PARENT>
            <ISDEEMEDPOSITIVE>Yes</ISDEEMEDPOSITIVE>
            <OPENINGBALANCE>-420000.00</OPENINGBALANCE>
          </LEDGER>
          <LEDGER NAME="Apex Automotive Components Ltd">
            <NAME>Apex Automotive Components Ltd</NAME>
            <PARENT>Sundry Debtors</PARENT>
            <PARTYGSTIN>27AABCA1122K1Z5</PARTYGSTIN>
            <ISDEEMEDPOSITIVE>Yes</ISDEEMEDPOSITIVE>
            <OPENINGBALANCE>-1200000.00</OPENINGBALANCE>
          </LEDGER>
          <LEDGER NAME="Vikram Heavy Turbines Pvt Ltd">
            <NAME>Vikram Heavy Turbines Pvt Ltd</NAME>
            <PARENT>Sundry Debtors</PARENT>
            <PARTYGSTIN>27AABCV3344M1Z2</PARTYGSTIN>
            <ISDEEMEDPOSITIVE>Yes</ISDEEMEDPOSITIVE>
            <OPENINGBALANCE>-850000.00</OPENINGBALANCE>
          </LEDGER>
          <LEDGER NAME="Deccan Aerospace Systems">
            <NAME>Deccan Aerospace Systems</NAME>
            <PARENT>Sundry Debtors</PARENT>
            <PARTYGSTIN>29AABCD5566P1Z8</PARTYGSTIN>
            <ISDEEMEDPOSITIVE>Yes</ISDEEMEDPOSITIVE>
            <OPENINGBALANCE>-450000.00</OPENINGBALANCE>
          </LEDGER>
          <LEDGER NAME="Kalyani Forgings &amp; Alloys">
            <NAME>Kalyani Forgings &amp; Alloys</NAME>
            <PARENT>Sundry Creditors</PARENT>
            <PARTYGSTIN>27AABCK7788Q1Z1</PARTYGSTIN>
            <ISDEEMEDPOSITIVE>No</ISDEEMEDPOSITIVE>
            <OPENINGBALANCE>650000.00</OPENINGBALANCE>
          </LEDGER>
          <LEDGER NAME="Ratnagiri Specialty Steels">
            <NAME>Ratnagiri Specialty Steels</NAME>
            <PARENT>Sundry Creditors</PARENT>
            <PARTYGSTIN>27AABCR9900L1Z9</PARTYGSTIN>
            <ISDEEMEDPOSITIVE>No</ISDEEMEDPOSITIVE>
            <OPENINGBALANCE>420000.00</OPENINGBALANCE>
          </LEDGER>
          <LEDGER NAME="Legacy Unreconciled Suspense A/c">
            <NAME>Legacy Unreconciled Suspense A/c</NAME>
            <PARENT>Suspense A/c</PARENT>
            <ISDEEMEDPOSITIVE>No</ISDEEMEDPOSITIVE>
            <OPENINGBALANCE>45000.00</OPENINGBALANCE>
          </LEDGER>
          <LEDGER NAME="Special Project Clearing Head">
            <NAME>Special Project Clearing Head</NAME>
            <PARENT>Miscellaneous Unmapped Group</PARENT>
            <ISDEEMEDPOSITIVE>No</ISDEEMEDPOSITIVE>
            <OPENINGBALANCE>25000.00</OPENINGBALANCE>
          </LEDGER>
          <STOCKITEM NAME="High-Tensile Alloy Billet 42CrMo4">
            <NAME>High-Tensile Alloy Billet 42CrMo4</NAME>
            <PARENT>Raw Materials</PARENT>
            <BASEUNITS>MT</BASEUNITS>
            <HSNCODE>722830</HSNCODE>
            <OPENINGBALANCE>18 MT</OPENINGBALANCE>
            <OPENINGVALUE>-1440000.00</OPENINGVALUE>
            <CLOSINGBALANCE>22 MT</CLOSINGBALANCE>
            <CLOSINGVALUE>-1760000.00</CLOSINGVALUE>
          </STOCKITEM>
        </TALLYMESSAGE>
      </REQUESTDATA>
    </IMPORTDATA>
  </BODY>
</ENVELOPE>`;
}

export function getSampleTallyTransactionsXML(): string {
  return `<?xml version="1.0" encoding="UTF-8"?>
<ENVELOPE>
  <HEADER>
    <TALLYREQUEST>Export Data</TALLYREQUEST>
  </HEADER>
  <BODY>
    <IMPORTDATA>
      <REQUESTDESC>
        <REPORTNAME>Vouchers</REPORTNAME>
        <STATICVARIABLES>
          <SVCURRENTCOMPANY>${DEMO_COMPANY_NAME}</SVCURRENTCOMPANY>
        </STATICVARIABLES>
      </REQUESTDESC>
      <REQUESTDATA>
        <TALLYMESSAGE xmlns:UDF="TallyUDF">
          <!-- 1. Jan 2026 Sales to Apex Automotive -->
          <VOUCHER VCHTYPE="Sales" ACTION="Create">
            <GUID>GSPU-DEMO-VCH-001</GUID>
            <DATE>20260110</DATE>
            <VOUCHERTYPENAME>Sales</VOUCHERTYPENAME>
            <VOUCHERNUMBER>BPE/25-26/101</VOUCHERNUMBER>
            <PARTYLEDGERNAME>Apex Automotive Components Ltd</PARTYLEDGERNAME>
            <PARTYGSTIN>27AABCA1122K1Z5</PARTYGSTIN>
            <PLACEOFSUPPLY>Maharashtra</PLACEOFSUPPLY>
            <NARRATION>Supply of precision CNC machined transmission shafts</NARRATION>
            <ALLLEDGERENTRIES.LIST>
              <LEDGERNAME>Apex Automotive Components Ltd</LEDGERNAME>
              <ISDEEMEDPOSITIVE>Yes</ISDEEMEDPOSITIVE>
              <ISPARTYLEDGER>Yes</ISPARTYLEDGER>
              <AMOUNT>-3540000.00</AMOUNT>
              <BILLALLOCATIONS.LIST>
                <NAME>BPE/25-26/101</NAME>
                <BILLTYPE>New Ref</BILLTYPE>
                <BILLCREDITPERIOD>30 Days</BILLCREDITPERIOD>
                <AMOUNT>-3540000.00</AMOUNT>
              </BILLALLOCATIONS.LIST>
            </ALLLEDGERENTRIES.LIST>
            <ALLLEDGERENTRIES.LIST>
              <LEDGERNAME>Domestic GST Sales</LEDGERNAME>
              <ISDEEMEDPOSITIVE>No</ISDEEMEDPOSITIVE>
              <AMOUNT>3000000.00</AMOUNT>
            </ALLLEDGERENTRIES.LIST>
            <ALLLEDGERENTRIES.LIST>
              <LEDGERNAME>Output CGST 9%</LEDGERNAME>
              <ISDEEMEDPOSITIVE>No</ISDEEMEDPOSITIVE>
              <AMOUNT>270000.00</AMOUNT>
            </ALLLEDGERENTRIES.LIST>
            <ALLLEDGERENTRIES.LIST>
              <LEDGERNAME>Output SGST 9%</LEDGERNAME>
              <ISDEEMEDPOSITIVE>No</ISDEEMEDPOSITIVE>
              <AMOUNT>270000.00</AMOUNT>
            </ALLLEDGERENTRIES.LIST>
          </VOUCHER>

          <!-- 2. Jan 2026 Partial Receipt from Apex Automotive -->
          <VOUCHER VCHTYPE="Receipt" ACTION="Create">
            <GUID>GSPU-DEMO-VCH-002</GUID>
            <DATE>20260124</DATE>
            <VOUCHERTYPENAME>Receipt</VOUCHERTYPENAME>
            <VOUCHERNUMBER>RCPT/25-26/041</VOUCHERNUMBER>
            <PARTYLEDGERNAME>Apex Automotive Components Ltd</PARTYLEDGERNAME>
            <NARRATION>NEFT collection against Bill BPE/25-26/101</NARRATION>
            <ALLLEDGERENTRIES.LIST>
              <LEDGERNAME>HDFC Bank Current A/c</LEDGERNAME>
              <ISDEEMEDPOSITIVE>Yes</ISDEEMEDPOSITIVE>
              <AMOUNT>-2000000.00</AMOUNT>
            </ALLLEDGERENTRIES.LIST>
            <ALLLEDGERENTRIES.LIST>
              <LEDGERNAME>Apex Automotive Components Ltd</LEDGERNAME>
              <ISDEEMEDPOSITIVE>No</ISDEEMEDPOSITIVE>
              <ISPARTYLEDGER>Yes</ISPARTYLEDGER>
              <AMOUNT>2000000.00</AMOUNT>
              <BILLALLOCATIONS.LIST>
                <NAME>BPE/25-26/101</NAME>
                <BILLTYPE>Agst Ref</BILLTYPE>
                <AMOUNT>2000000.00</AMOUNT>
              </BILLALLOCATIONS.LIST>
            </ALLLEDGERENTRIES.LIST>
          </VOUCHER>

          <!-- 3. Jan 2026 Purchase from Kalyani Forgings -->
          <VOUCHER VCHTYPE="Purchase" ACTION="Create">
            <GUID>GSPU-DEMO-VCH-003</GUID>
            <DATE>20260114</DATE>
            <VOUCHERTYPENAME>Purchase</VOUCHERTYPENAME>
            <VOUCHERNUMBER>PUR/25-26/088</VOUCHERNUMBER>
            <REFERENCE>KF/INV/8821</REFERENCE>
            <PARTYLEDGERNAME>Kalyani Forgings &amp; Alloys</PARTYLEDGERNAME>
            <PARTYGSTIN>27AABCK7788Q1Z1</PARTYGSTIN>
            <NARRATION>Purchase of alloy steel forgings</NARRATION>
            <ALLLEDGERENTRIES.LIST>
              <LEDGERNAME>Raw Material Purchases</LEDGERNAME>
              <ISDEEMEDPOSITIVE>Yes</ISDEEMEDPOSITIVE>
              <AMOUNT>-1800000.00</AMOUNT>
            </ALLLEDGERENTRIES.LIST>
            <ALLLEDGERENTRIES.LIST>
              <LEDGERNAME>Input CGST 9%</LEDGERNAME>
              <ISDEEMEDPOSITIVE>Yes</ISDEEMEDPOSITIVE>
              <AMOUNT>-162000.00</AMOUNT>
            </ALLLEDGERENTRIES.LIST>
            <ALLLEDGERENTRIES.LIST>
              <LEDGERNAME>Input SGST 9%</LEDGERNAME>
              <ISDEEMEDPOSITIVE>Yes</ISDEEMEDPOSITIVE>
              <AMOUNT>-162000.00</AMOUNT>
            </ALLLEDGERENTRIES.LIST>
            <ALLLEDGERENTRIES.LIST>
              <LEDGERNAME>Kalyani Forgings &amp; Alloys</LEDGERNAME>
              <ISDEEMEDPOSITIVE>No</ISDEEMEDPOSITIVE>
              <ISPARTYLEDGER>Yes</ISPARTYLEDGER>
              <AMOUNT>2124000.00</AMOUNT>
              <BILLALLOCATIONS.LIST>
                <NAME>KF/INV/8821</NAME>
                <BILLTYPE>New Ref</BILLTYPE>
                <BILLCREDITPERIOD>30 Days</BILLCREDITPERIOD>
                <AMOUNT>2124000.00</AMOUNT>
              </BILLALLOCATIONS.LIST>
            </ALLLEDGERENTRIES.LIST>
          </VOUCHER>

          <!-- 4. Feb 2026 Inter-State Sales to Deccan Aerospace -->
          <VOUCHER VCHTYPE="Sales" ACTION="Create">
            <GUID>GSPU-DEMO-VCH-004</GUID>
            <DATE>20260208</DATE>
            <VOUCHERTYPENAME>Sales</VOUCHERTYPENAME>
            <VOUCHERNUMBER>BPE/25-26/102</VOUCHERNUMBER>
            <PARTYLEDGERNAME>Deccan Aerospace Systems</PARTYLEDGERNAME>
            <PARTYGSTIN>29AABCD5566P1Z8</PARTYGSTIN>
            <PLACEOFSUPPLY>Karnataka</PLACEOFSUPPLY>
            <NARRATION>Supply of aerospace turbine housings</NARRATION>
            <ALLLEDGERENTRIES.LIST>
              <LEDGERNAME>Deccan Aerospace Systems</LEDGERNAME>
              <ISDEEMEDPOSITIVE>Yes</ISDEEMEDPOSITIVE>
              <ISPARTYLEDGER>Yes</ISPARTYLEDGER>
              <AMOUNT>-2832000.00</AMOUNT>
              <BILLALLOCATIONS.LIST>
                <NAME>BPE/25-26/102</NAME>
                <BILLTYPE>New Ref</BILLTYPE>
                <BILLCREDITPERIOD>15 Days</BILLCREDITPERIOD>
                <AMOUNT>-2832000.00</AMOUNT>
              </BILLALLOCATIONS.LIST>
            </ALLLEDGERENTRIES.LIST>
            <ALLLEDGERENTRIES.LIST>
              <LEDGERNAME>Domestic GST Sales</LEDGERNAME>
              <ISDEEMEDPOSITIVE>No</ISDEEMEDPOSITIVE>
              <AMOUNT>2400000.00</AMOUNT>
            </ALLLEDGERENTRIES.LIST>
            <ALLLEDGERENTRIES.LIST>
              <LEDGERNAME>Output IGST 18%</LEDGERNAME>
              <ISDEEMEDPOSITIVE>No</ISDEEMEDPOSITIVE>
              <AMOUNT>432000.00</AMOUNT>
            </ALLLEDGERENTRIES.LIST>
          </VOUCHER>

          <!-- 5. Feb 2026 Credit Note (Sales Return) from Deccan Aerospace -->
          <VOUCHER VCHTYPE="Credit Note" ACTION="Create">
            <GUID>GSPU-DEMO-VCH-005</GUID>
            <DATE>20260218</DATE>
            <VOUCHERTYPENAME>Credit Note</VOUCHERTYPENAME>
            <VOUCHERNUMBER>CN/25-26/009</VOUCHERNUMBER>
            <PARTYLEDGERNAME>Deccan Aerospace Systems</PARTYLEDGERNAME>
            <PARTYGSTIN>29AABCD5566P1Z8</PARTYGSTIN>
            <NARRATION>Sales return for dimensional tolerance variance on 1 unit</NARRATION>
            <ALLLEDGERENTRIES.LIST>
              <LEDGERNAME>Sales Return Account</LEDGERNAME>
              <ISDEEMEDPOSITIVE>Yes</ISDEEMEDPOSITIVE>
              <AMOUNT>-200000.00</AMOUNT>
            </ALLLEDGERENTRIES.LIST>
            <ALLLEDGERENTRIES.LIST>
              <LEDGERNAME>Output IGST 18%</LEDGERNAME>
              <ISDEEMEDPOSITIVE>Yes</ISDEEMEDPOSITIVE>
              <AMOUNT>-36000.00</AMOUNT>
            </ALLLEDGERENTRIES.LIST>
            <ALLLEDGERENTRIES.LIST>
              <LEDGERNAME>Deccan Aerospace Systems</LEDGERNAME>
              <ISDEEMEDPOSITIVE>No</ISDEEMEDPOSITIVE>
              <ISPARTYLEDGER>Yes</ISPARTYLEDGER>
              <AMOUNT>236000.00</AMOUNT>
              <BILLALLOCATIONS.LIST>
                <NAME>BPE/25-26/102</NAME>
                <BILLTYPE>Agst Ref</BILLTYPE>
                <AMOUNT>236000.00</AMOUNT>
              </BILLALLOCATIONS.LIST>
            </ALLLEDGERENTRIES.LIST>
          </VOUCHER>

          <!-- 6. Feb 2026 Purchase from Ratnagiri Specialty Steels (Missing in GSTR-2B to demonstrate GST Reconciliation!) -->
          <VOUCHER VCHTYPE="Purchase" ACTION="Create">
            <GUID>GSPU-DEMO-VCH-006</GUID>
            <DATE>20260220</DATE>
            <VOUCHERTYPENAME>Purchase</VOUCHERTYPENAME>
            <VOUCHERNUMBER>PUR/25-26/094</VOUCHERNUMBER>
            <REFERENCE>RSS/25-26/409</REFERENCE>
            <PARTYLEDGERNAME>Ratnagiri Specialty Steels</PARTYLEDGERNAME>
            <PARTYGSTIN>27AABCR9900L1Z9</PARTYGSTIN>
            <NARRATION>Purchase of stainless steel bars</NARRATION>
            <ALLLEDGERENTRIES.LIST>
              <LEDGERNAME>Raw Material Purchases</LEDGERNAME>
              <ISDEEMEDPOSITIVE>Yes</ISDEEMEDPOSITIVE>
              <AMOUNT>-950000.00</AMOUNT>
            </ALLLEDGERENTRIES.LIST>
            <ALLLEDGERENTRIES.LIST>
              <LEDGERNAME>Input CGST 9%</LEDGERNAME>
              <ISDEEMEDPOSITIVE>Yes</ISDEEMEDPOSITIVE>
              <AMOUNT>-85500.00</AMOUNT>
            </ALLLEDGERENTRIES.LIST>
            <ALLLEDGERENTRIES.LIST>
              <LEDGERNAME>Input SGST 9%</LEDGERNAME>
              <ISDEEMEDPOSITIVE>Yes</ISDEEMEDPOSITIVE>
              <AMOUNT>-85500.00</AMOUNT>
            </ALLLEDGERENTRIES.LIST>
            <ALLLEDGERENTRIES.LIST>
              <LEDGERNAME>Ratnagiri Specialty Steels</LEDGERNAME>
              <ISDEEMEDPOSITIVE>No</ISDEEMEDPOSITIVE>
              <ISPARTYLEDGER>Yes</ISPARTYLEDGER>
              <AMOUNT>1121000.00</AMOUNT>
              <BILLALLOCATIONS.LIST>
                <NAME>RSS/25-26/409</NAME>
                <BILLTYPE>New Ref</BILLTYPE>
                <BILLCREDITPERIOD>15 Days</BILLCREDITPERIOD>
                <AMOUNT>1121000.00</AMOUNT>
              </BILLALLOCATIONS.LIST>
            </ALLLEDGERENTRIES.LIST>
          </VOUCHER>

          <!-- 7. Feb 2026 Operating Expenses Payment -->
          <VOUCHER VCHTYPE="Payment" ACTION="Create">
            <GUID>GSPU-DEMO-VCH-007</GUID>
            <DATE>20260226</DATE>
            <VOUCHERTYPENAME>Payment</VOUCHERTYPENAME>
            <VOUCHERNUMBER>PYMT/25-26/112</VOUCHERNUMBER>
            <NARRATION>Monthly salaries, factory power, and admin expenses</NARRATION>
            <ALLLEDGERENTRIES.LIST>
              <LEDGERNAME>Salaries &amp; Staff Welfare</LEDGERNAME>
              <ISDEEMEDPOSITIVE>Yes</ISDEEMEDPOSITIVE>
              <AMOUNT>-820000.00</AMOUNT>
            </ALLLEDGERENTRIES.LIST>
            <ALLLEDGERENTRIES.LIST>
              <LEDGERNAME>Power &amp; Fuel (Factory)</LEDGERNAME>
              <ISDEEMEDPOSITIVE>Yes</ISDEEMEDPOSITIVE>
              <AMOUNT>-340000.00</AMOUNT>
            </ALLLEDGERENTRIES.LIST>
            <ALLLEDGERENTRIES.LIST>
              <LEDGERNAME>Administrative &amp; Legal Fees</LEDGERNAME>
              <ISDEEMEDPOSITIVE>Yes</ISDEEMEDPOSITIVE>
              <AMOUNT>-290000.00</AMOUNT>
            </ALLLEDGERENTRIES.LIST>
            <ALLLEDGERENTRIES.LIST>
              <LEDGERNAME>Sales Promotion &amp; Marketing</LEDGERNAME>
              <ISDEEMEDPOSITIVE>Yes</ISDEEMEDPOSITIVE>
              <AMOUNT>-150000.00</AMOUNT>
            </ALLLEDGERENTRIES.LIST>
            <ALLLEDGERENTRIES.LIST>
              <LEDGERNAME>HDFC Bank Current A/c</LEDGERNAME>
              <ISDEEMEDPOSITIVE>No</ISDEEMEDPOSITIVE>
              <AMOUNT>1600000.00</AMOUNT>
            </ALLLEDGERENTRIES.LIST>
          </VOUCHER>

          <!-- 8. March 2026 Sales to Vikram Heavy Turbines (Omitted in GSTR-1 to test GSTR-1 reconciliation!) -->
          <VOUCHER VCHTYPE="Sales" ACTION="Create">
            <GUID>GSPU-DEMO-VCH-008</GUID>
            <DATE>20260312</DATE>
            <VOUCHERTYPENAME>Sales</VOUCHERTYPENAME>
            <VOUCHERNUMBER>BPE/25-26/103</VOUCHERNUMBER>
            <PARTYLEDGERNAME>Vikram Heavy Turbines Pvt Ltd</PARTYLEDGERNAME>
            <PARTYGSTIN>27AABCV3344M1Z2</PARTYGSTIN>
            <PLACEOFSUPPLY>Maharashtra</PLACEOFSUPPLY>
            <NARRATION>Precision rotor assemblies batch #9</NARRATION>
            <ALLLEDGERENTRIES.LIST>
              <LEDGERNAME>Vikram Heavy Turbines Pvt Ltd</LEDGERNAME>
              <ISDEEMEDPOSITIVE>Yes</ISDEEMEDPOSITIVE>
              <ISPARTYLEDGER>Yes</ISPARTYLEDGER>
              <AMOUNT>-4484000.00</AMOUNT>
              <BILLALLOCATIONS.LIST>
                <NAME>BPE/25-26/103</NAME>
                <BILLTYPE>New Ref</BILLTYPE>
                <BILLCREDITPERIOD>10 Days</BILLCREDITPERIOD>
                <AMOUNT>-4484000.00</AMOUNT>
              </BILLALLOCATIONS.LIST>
            </ALLLEDGERENTRIES.LIST>
            <ALLLEDGERENTRIES.LIST>
              <LEDGERNAME>Domestic GST Sales</LEDGERNAME>
              <ISDEEMEDPOSITIVE>No</ISDEEMEDPOSITIVE>
              <AMOUNT>3800000.00</AMOUNT>
            </ALLLEDGERENTRIES.LIST>
            <ALLLEDGERENTRIES.LIST>
              <LEDGERNAME>Output CGST 9%</LEDGERNAME>
              <ISDEEMEDPOSITIVE>No</ISDEEMEDPOSITIVE>
              <AMOUNT>342000.00</AMOUNT>
            </ALLLEDGERENTRIES.LIST>
            <ALLLEDGERENTRIES.LIST>
              <LEDGERNAME>Output SGST 9%</LEDGERNAME>
              <ISDEEMEDPOSITIVE>No</ISDEEMEDPOSITIVE>
              <AMOUNT>342000.00</AMOUNT>
            </ALLLEDGERENTRIES.LIST>
          </VOUCHER>

          <!-- 9. March 2026 Year-End Depreciation, Interest & Tax Provision Journal -->
          <VOUCHER VCHTYPE="Journal" ACTION="Create">
            <GUID>GSPU-DEMO-VCH-009</GUID>
            <DATE>20260331</DATE>
            <VOUCHERTYPENAME>Journal</VOUCHERTYPENAME>
            <VOUCHERNUMBER>JRNL/25-26/099</VOUCHERNUMBER>
            <NARRATION>Year-end depreciation, term loan interest and tax provision</NARRATION>
            <ALLLEDGERENTRIES.LIST>
              <LEDGERNAME>Depreciation on Plant &amp; Machinery</LEDGERNAME>
              <ISDEEMEDPOSITIVE>Yes</ISDEEMEDPOSITIVE>
              <AMOUNT>-480000.00</AMOUNT>
            </ALLLEDGERENTRIES.LIST>
            <ALLLEDGERENTRIES.LIST>
              <LEDGERNAME>Interest on Term Loan &amp; CC</LEDGERNAME>
              <ISDEEMEDPOSITIVE>Yes</ISDEEMEDPOSITIVE>
              <AMOUNT>-310000.00</AMOUNT>
            </ALLLEDGERENTRIES.LIST>
            <ALLLEDGERENTRIES.LIST>
              <LEDGERNAME>Income Tax Provision</LEDGERNAME>
              <ISDEEMEDPOSITIVE>Yes</ISDEEMEDPOSITIVE>
              <AMOUNT>-620000.00</AMOUNT>
            </ALLLEDGERENTRIES.LIST>
            <ALLLEDGERENTRIES.LIST>
              <LEDGERNAME>CNC Machining Centres &amp; Plant</LEDGERNAME>
              <ISDEEMEDPOSITIVE>No</ISDEEMEDPOSITIVE>
              <AMOUNT>480000.00</AMOUNT>
            </ALLLEDGERENTRIES.LIST>
            <ALLLEDGERENTRIES.LIST>
              <LEDGERNAME>SBI Term Loan A/c</LEDGERNAME>
              <ISDEEMEDPOSITIVE>No</ISDEEMEDPOSITIVE>
              <AMOUNT>310000.00</AMOUNT>
            </ALLLEDGERENTRIES.LIST>
            <ALLLEDGERENTRIES.LIST>
              <LEDGERNAME>Provision for Income Tax Payable</LEDGERNAME>
              <ISDEEMEDPOSITIVE>No</ISDEEMEDPOSITIVE>
              <AMOUNT>620000.00</AMOUNT>
            </ALLLEDGERENTRIES.LIST>
          </VOUCHER>

          <!-- 10. Intentional Suspected Duplicate Voucher on same day for Audit Health Check verification -->
          <VOUCHER VCHTYPE="Payment" ACTION="Create">
            <GUID>GSPU-DEMO-VCH-010</GUID>
            <DATE>20260329</DATE>
            <VOUCHERTYPENAME>Payment</VOUCHERTYPENAME>
            <VOUCHERNUMBER>PYMT/25-26/150</VOUCHERNUMBER>
            <PARTYLEDGERNAME>Kalyani Forgings &amp; Alloys</PARTYLEDGERNAME>
            <NARRATION>Part settlement of supplier account</NARRATION>
            <ALLLEDGERENTRIES.LIST>
              <LEDGERNAME>Kalyani Forgings &amp; Alloys</LEDGERNAME>
              <ISDEEMEDPOSITIVE>Yes</ISDEEMEDPOSITIVE>
              <ISPARTYLEDGER>Yes</ISPARTYLEDGER>
              <AMOUNT>-500000.00</AMOUNT>
              <BILLALLOCATIONS.LIST>
                <NAME>KF/INV/8821</NAME>
                <BILLTYPE>Agst Ref</BILLTYPE>
                <AMOUNT>-500000.00</AMOUNT>
              </BILLALLOCATIONS.LIST>
            </ALLLEDGERENTRIES.LIST>
            <ALLLEDGERENTRIES.LIST>
              <LEDGERNAME>HDFC Bank Current A/c</LEDGERNAME>
              <ISDEEMEDPOSITIVE>No</ISDEEMEDPOSITIVE>
              <AMOUNT>500000.00</AMOUNT>
            </ALLLEDGERENTRIES.LIST>
          </VOUCHER>
        </TALLYMESSAGE>
      </REQUESTDATA>
    </IMPORTDATA>
  </BODY>
</ENVELOPE>`;
}

export function getSampleGSTR1Json(): string {
  return JSON.stringify(
    {
      gstin: '27AABCB9988D1Z4',
      fp: '022026',
      b2b: [
        {
          ctin: '27AABCA1122K1Z5',
          cname: 'Apex Automotive Components Ltd',
          inv: [
            {
              inum: 'BPE/25-26/101',
              idt: '2026-01-10',
              val: 3540000,
              pos: '27',
              rchrg: 'N',
              itms: [
                {
                  num: 1,
                  itm_det: {
                    txval: 3000000,
                    rt: 18,
                    camt: 270000,
                    samt: 270000,
                    iamt: 0,
                    csamt: 0,
                  },
                },
              ],
            },
          ],
        },
        {
          ctin: '29AABCD5566P1Z8',
          cname: 'Deccan Aerospace Systems',
          inv: [
            {
              inum: 'BPE-2526-102', // Normalised match test (hyphens vs slashes)
              idt: '2026-02-08',
              val: 2832000,
              pos: '29',
              rchrg: 'N',
              itms: [
                {
                  num: 1,
                  itm_det: {
                    txval: 2400000,
                    rt: 18,
                    camt: 0,
                    samt: 0,
                    iamt: 432000,
                    csamt: 0,
                  },
                },
              ],
            },
          ],
        },
      ],
      cdnr: [
        {
          ctin: '29AABCD5566P1Z8',
          cname: 'Deccan Aerospace Systems',
          nt: [
            {
              nt_num: 'CN/25-26/009',
              nt_dt: '2026-02-18',
              ntty: 'C',
              itms: [
                {
                  num: 1,
                  itm_det: {
                    txval: 200000,
                    rt: 18,
                    iamt: 36000,
                    camt: 0,
                    samt: 0,
                  },
                },
              ],
            },
          ],
        },
      ],
    },
    null,
    2
  );
}

export function getSampleGSTR2BJson(): string {
  return JSON.stringify(
    {
      data: {
        gstin: '27AABCB9988D1Z4',
        rtnprd: '022026',
        docdata: {
          b2b: [
            {
              ctin: '27AABCK7788Q1Z1',
              trdnm: 'Kalyani Forgings & Alloys',
              inv: [
                {
                  inum: 'KF/INV/8821',
                  dt: '2026-01-14',
                  txval: 1800000,
                  igst: 0,
                  cgst: 162000,
                  sgst: 162000,
                  cess: 0,
                  itcavl: 'Y',
                },
              ],
            },
            {
              ctin: '27AABCP4411E1Z7',
              trdnm: 'Pune Industrial Tools (Unbooked Portal Invoice)',
              inv: [
                {
                  inum: 'PIT/2026/774',
                  dt: '2026-02-25',
                  txval: 250000,
                  igst: 0,
                  cgst: 22500,
                  sgst: 22500,
                  cess: 0,
                  itcavl: 'Y',
                },
              ],
            },
          ],
        },
      },
    },
    null,
    2
  );
}

export function getSampleGSTR3BJson(): string {
  return JSON.stringify(
    {
      gstin: '27AABCB9988D1Z4',
      ret_period: '032026',
      sup_details: {
        osup_det: {
          txval: 9000000,
          iamt: 396000,
          camt: 612000,
          samt: 612000,
          csamt: 0,
        },
      },
      itc_elg: {
        itc_net: {
          iamt: 0,
          camt: 220000,
          samt: 220000,
          csamt: 0,
        },
      },
    },
    null,
    2
  );
}

export function buildDemoPayloads(): RawFilePayload[] {
  const now = new Date().toISOString();

  const mastersXml = getSampleTallyMastersXML();
  const parsedMasters = parseTallyXML(mastersXml, 'DEMO_FILE_1', 'Masters_Demo.xml');

  const txXml = getSampleTallyTransactionsXML();
  const parsedTx = parseTallyXML(txXml, 'DEMO_FILE_2', 'Transactions_FY25_26_Demo.xml');

  const gstr1Raw = getSampleGSTR1Json();
  const parsedG1 = parseGSTJson(gstr1Raw, 'DEMO_FILE_3', 'GSTR1_Return_Demo.json');

  const gstr2bRaw = getSampleGSTR2BJson();
  const parsedG2B = parseGSTJson(gstr2bRaw, 'DEMO_FILE_4', 'GSTR2B_Statement_Demo.json');

  const gstr3bRaw = getSampleGSTR3BJson();
  const parsedG3B = parseGSTJson(gstr3bRaw, 'DEMO_FILE_5', 'GSTR3B_Summary_Demo.json');

  return [
    {
      sourceFile: {
        id: 'DEMO_FILE_1',
        filename: 'Masters_Demo.xml',
        format: 'XML',
        sizeBytes: mastersXml.length,
        uploadedAt: now,
        detectedReportType: parsedMasters.detectedReportType,
        extractedCompanyName: parsedMasters.extractedCompanyName,
        periodStart: '2025-04-01',
        periodEnd: '2026-03-31',
        recordsExtracted: parsedMasters.ledgers.length + parsedMasters.stockItems.length,
        status: 'VALID',
        validationMessages: parsedMasters.validationMessages,
        isDemoData: true,
      },
      groups: parsedMasters.groups,
      ledgers: parsedMasters.ledgers,
      vouchers: [],
      outstandingInvoices: [],
      stockItems: parsedMasters.stockItems,
      gstInvoices: [],
      gstr3bSummaries: [],
    },
    {
      sourceFile: {
        id: 'DEMO_FILE_2',
        filename: 'Transactions_FY25_26_Demo.xml',
        format: 'XML',
        sizeBytes: txXml.length,
        uploadedAt: now,
        detectedReportType: parsedTx.detectedReportType,
        extractedCompanyName: parsedTx.extractedCompanyName,
        periodStart: parsedTx.periodStart,
        periodEnd: parsedTx.periodEnd,
        recordsExtracted: parsedTx.vouchers.length,
        status: 'VALID',
        validationMessages: parsedTx.validationMessages,
        isDemoData: true,
      },
      groups: parsedTx.groups,
      ledgers: parsedTx.ledgers,
      vouchers: parsedTx.vouchers,
      outstandingInvoices: parsedTx.outstandingInvoices,
      stockItems: parsedTx.stockItems,
      gstInvoices: parsedTx.gstInvoices,
      gstr3bSummaries: [],
    },
    {
      sourceFile: {
        id: 'DEMO_FILE_3',
        filename: 'GSTR1_Return_Demo.json',
        format: 'JSON',
        sizeBytes: gstr1Raw.length,
        uploadedAt: now,
        detectedReportType: parsedG1.detectedReportType,
        periodStart: '2026-01-01',
        periodEnd: '2026-03-31',
        recordsExtracted: parsedG1.gstInvoices.length,
        status: 'VALID',
        validationMessages: parsedG1.validationMessages,
        isDemoData: true,
      },
      groups: [],
      ledgers: [],
      vouchers: [],
      outstandingInvoices: [],
      stockItems: [],
      gstInvoices: parsedG1.gstInvoices,
      gstr3bSummaries: [],
    },
    {
      sourceFile: {
        id: 'DEMO_FILE_4',
        filename: 'GSTR2B_Statement_Demo.json',
        format: 'JSON',
        sizeBytes: gstr2bRaw.length,
        uploadedAt: now,
        detectedReportType: parsedG2B.detectedReportType,
        periodStart: '2026-01-01',
        periodEnd: '2026-03-31',
        recordsExtracted: parsedG2B.gstInvoices.length,
        status: 'VALID',
        validationMessages: parsedG2B.validationMessages,
        isDemoData: true,
      },
      groups: [],
      ledgers: [],
      vouchers: [],
      outstandingInvoices: [],
      stockItems: [],
      gstInvoices: parsedG2B.gstInvoices,
      gstr3bSummaries: [],
    },
    {
      sourceFile: {
        id: 'DEMO_FILE_5',
        filename: 'GSTR3B_Summary_Demo.json',
        format: 'JSON',
        sizeBytes: gstr3bRaw.length,
        uploadedAt: now,
        detectedReportType: parsedG3B.detectedReportType,
        periodStart: '2026-01-01',
        periodEnd: '2026-03-31',
        recordsExtracted: parsedG3B.gstr3bSummaries.length,
        status: 'VALID',
        validationMessages: parsedG3B.validationMessages,
        isDemoData: true,
      },
      groups: [],
      ledgers: [],
      vouchers: [],
      outstandingInvoices: [],
      stockItems: [],
      gstInvoices: [],
      gstr3bSummaries: parsedG3B.gstr3bSummaries,
    },
    {
      sourceFile: {
        id: 'DEMO_FILE_6',
        filename: 'Cash_Flow_Statement_FY25_26_Demo.xlsx',
        format: 'XLSX',
        sizeBytes: 18420,
        uploadedAt: now,
        detectedReportType: 'CASH_FLOW_EXCEL',
        extractedCompanyName: DEMO_COMPANY_NAME,
        periodStart: '2025-04-01',
        periodEnd: '2026-03-31',
        recordsExtracted: 7,
        status: 'VALID',
        validationMessages: [
          {
            severity: 'INFO',
            code: 'CF_DEMO_DETECTED',
            message: 'Auto-detected Cash Flow Statement with Operating, Investing, Capex, and Financing schedules.',
          },
        ],
        isDemoData: true,
        sheetNames: ['Cash Flow Statement'],
      },
      groups: [],
      ledgers: [],
      vouchers: [],
      outstandingInvoices: [],
      stockItems: [],
      gstInvoices: [],
      gstr3bSummaries: [],
      cashFlowItems: [
        {
          id: 'CF_DEMO_1',
          particulars: 'Cash Generated from Operations (Working Capital Adjusted)',
          normalizedParticulars: 'cash generated from operations (working capital adjusted)',
          category: 'CFO_OPERATING_ITEM',
          currentPeriodPaise: toPaise(3150000),
          previousPeriodPaise: toPaise(2480000),
          budgetPaise: toPaise(3000000),
          sourceFileId: 'DEMO_FILE_6',
          sourceFilename: 'Cash_Flow_Statement_FY25_26_Demo.xlsx',
          worksheetName: 'Cash Flow Statement',
          cellReference: 'Cash Flow Statement!B5',
          isSubtotalOrTotal: false,
        },
        {
          id: 'CF_DEMO_2',
          particulars: 'Income Tax Paid (Net of Refunds)',
          normalizedParticulars: 'income tax paid (net of refunds)',
          category: 'CFO_TAXES_PAID',
          currentPeriodPaise: toPaise(-520000),
          previousPeriodPaise: toPaise(-410000),
          budgetPaise: toPaise(-480000),
          sourceFileId: 'DEMO_FILE_6',
          sourceFilename: 'Cash_Flow_Statement_FY25_26_Demo.xlsx',
          worksheetName: 'Cash Flow Statement',
          cellReference: 'Cash Flow Statement!B6',
          isSubtotalOrTotal: false,
        },
        {
          id: 'CF_DEMO_3',
          particulars: 'Capital Expenditure — Purchase of CNC Machinery & Plant',
          normalizedParticulars: 'capital expenditure — purchase of cnc machinery & plant',
          category: 'CFI_CAPEX',
          currentPeriodPaise: toPaise(-1150000),
          previousPeriodPaise: toPaise(-850000),
          budgetPaise: toPaise(-1200000),
          sourceFileId: 'DEMO_FILE_6',
          sourceFilename: 'Cash_Flow_Statement_FY25_26_Demo.xlsx',
          worksheetName: 'Cash Flow Statement',
          cellReference: 'Cash Flow Statement!B9',
          isSubtotalOrTotal: false,
        },
        {
          id: 'CF_DEMO_4',
          particulars: 'Interest & Dividend Income Received',
          normalizedParticulars: 'interest & dividend income received',
          category: 'CFI_INVESTING_ITEM',
          currentPeriodPaise: toPaise(95000),
          previousPeriodPaise: toPaise(72000),
          budgetPaise: toPaise(80000),
          sourceFileId: 'DEMO_FILE_6',
          sourceFilename: 'Cash_Flow_Statement_FY25_26_Demo.xlsx',
          worksheetName: 'Cash Flow Statement',
          cellReference: 'Cash Flow Statement!B10',
          isSubtotalOrTotal: false,
        },
        {
          id: 'CF_DEMO_5',
          particulars: 'Repayment of Term Loan & Finance Costs Paid',
          normalizedParticulars: 'repayment of term loan & finance costs paid',
          category: 'CFF_FINANCING_ITEM',
          currentPeriodPaise: toPaise(-640000),
          previousPeriodPaise: toPaise(-590000),
          budgetPaise: toPaise(-640000),
          sourceFileId: 'DEMO_FILE_6',
          sourceFilename: 'Cash_Flow_Statement_FY25_26_Demo.xlsx',
          worksheetName: 'Cash Flow Statement',
          cellReference: 'Cash Flow Statement!B13',
          isSubtotalOrTotal: false,
        },
        {
          id: 'CF_DEMO_6',
          particulars: 'Opening Cash and Cash Equivalents',
          normalizedParticulars: 'opening cash and cash equivalents',
          category: 'CF_OPENING_CASH',
          currentPeriodPaise: toPaise(1415000),
          previousPeriodPaise: toPaise(713000),
          sourceFileId: 'DEMO_FILE_6',
          sourceFilename: 'Cash_Flow_Statement_FY25_26_Demo.xlsx',
          worksheetName: 'Cash Flow Statement',
          cellReference: 'Cash Flow Statement!B16',
          isSubtotalOrTotal: true,
        },
        {
          id: 'CF_DEMO_7',
          particulars: 'Closing Cash and Cash Equivalents',
          normalizedParticulars: 'closing cash and cash equivalents',
          category: 'CF_CLOSING_CASH',
          currentPeriodPaise: toPaise(2350000),
          previousPeriodPaise: toPaise(1415000),
          sourceFileId: 'DEMO_FILE_6',
          sourceFilename: 'Cash_Flow_Statement_FY25_26_Demo.xlsx',
          worksheetName: 'Cash Flow Statement',
          cellReference: 'Cash Flow Statement!B17',
          isSubtotalOrTotal: true,
        },
      ],
    },
  ];
}

export function downloadSampleStatementWorkbook(
  kind: 'BALANCE_SHEET' | 'PROFIT_LOSS' | 'CASH_FLOW' | 'CFO_PACK'
): void {
  const wb = XLSX.utils.book_new();

  const bsRows = [
    ['Particulars', 'Parent Group', 'Current Year (FY 2025-26)', 'Previous Year (FY 2024-25)', 'Budget (FY 2025-26)'],
    ['Equity Share Capital', 'Capital Account', 5000000, 5000000, 5000000],
    ['General Reserves & Surplus', 'Reserves & Surplus', 4620000, 3180000, 4450000],
    ['Term Loan - HDFC Bank', 'Secured Loans', 2400000, 2900000, 2400000],
    ['Working Capital OD Facility', 'Unsecured Loans', 650000, 820000, 600000],
    ['Sundry Creditors - Raw Materials', 'Sundry Creditors', 1940000, 2150000, 1800000],
    ['Statutory Dues & GST Payable', 'Duties & Taxes', 380000, 310000, 350000],
    ['Outstanding Expenses Payable', 'Provisions', 290000, 260000, 250000],
    ['Total Equity and Liabilities', 'Subtotal', 15280000, 14620000, 14850000],
    ['Plant & CNC Machinery (Net Block)', 'Fixed Assets', 6450000, 5920000, 6500000],
    ['Industrial Shed & Office Equipment', 'Fixed Assets', 1580000, 1690000, 1600000],
    ['Closing Inventory - Raw Material & FG', 'Stock-in-Hand', 2140000, 2380000, 1950000],
    ['Sundry Debtors - Domestic & OEM', 'Sundry Debtors', 2420000, 2890000, 2250000],
    ['HDFC Bank Current Account', 'Bank Accounts', 2120000, 1245000, 2200000],
    ['Cash on Hand', 'Cash-in-Hand', 230000, 170000, 150000],
    ['Security Deposits & Advances', 'Loans & Advances (Asset)', 340000, 325000, 200000],
    ['Total Assets', 'Subtotal', 15280000, 14620000, 14850000],
  ];

  const pnlRows = [
    [
      'Particulars',
      'Parent Group',
      'Current Year (FY 2025-26)',
      'Previous Year (FY 2024-25)',
      'Budget (FY 2025-26)',
      'Apr-25',
      'May-25',
      'Jun-25',
      'Jul-25',
      'Aug-25',
      'Sep-25',
      'Oct-25',
      'Nov-25',
      'Dec-25',
      'Jan-26',
      'Feb-26',
      'Mar-26',
    ],
    ['Domestic OEM Component Sales', 'Sales Accounts', 14800000, 12650000, 14200000, 1120000, 1180000, 1210000, 1190000, 1240000, 1280000, 1310000, 1220000, 1260000, 1240000, 1250000, 1300000],
    ['Export Precision Assembly Sales', 'Sales Accounts', 3600000, 2950000, 3500000, 260000, 280000, 290000, 300000, 310000, 310000, 320000, 290000, 300000, 310000, 310000, 320000],
    ['Total Revenue', 'Subtotal', 18400000, 15600000, 17700000, 1380000, 1460000, 1500000, 1490000, 1550000, 1590000, 1630000, 1510000, 1560000, 1550000, 1560000, 1620000],
    ['Raw Material Alloy Steel Purchases', 'Purchase Accounts', 8920000, 7850000, 8600000, 680000, 710000, 730000, 720000, 750000, 770000, 790000, 730000, 750000, 750000, 760000, 780000],
    ['Direct Factory Power & Machining Wages', 'Direct Expenses', 1940000, 1720000, 1880000, 150000, 155000, 160000, 160000, 165000, 165000, 170000, 160000, 160000, 165000, 165000, 165000],
    ['Gross Profit', 'Subtotal', 7540000, 6030000, 7220000, 550000, 595000, 610000, 610000, 635000, 655000, 670000, 620000, 650000, 635000, 635000, 675000],
    ['Employee Salaries & Engineering Payroll', 'Indirect Expenses', 2480000, 2240000, 2400000, 205000, 205000, 205000, 205000, 208000, 208000, 208000, 208000, 207000, 207000, 207000, 207000],
    ['Administrative, Rent, Logistics & Audit', 'Indirect Expenses', 1320000, 1190000, 1280000, 105000, 108000, 110000, 110000, 112000, 112000, 115000, 108000, 110000, 110000, 110000, 110000],
    ['EBITDA', 'Subtotal', 3740000, 2600000, 3540000, 240000, 282000, 295000, 295000, 315000, 335000, 347000, 304000, 333000, 318000, 318000, 358000],
    ['Depreciation on Plant & Machinery', 'Depreciation', 620000, 580000, 620000, 51000, 51000, 52000, 52000, 52000, 52000, 52000, 52000, 51000, 51000, 52000, 52000],
    ['Bank Term Loan & OD Interest', 'Finance Costs', 340000, 390000, 350000, 29000, 29000, 29000, 28000, 28000, 28000, 28000, 28000, 28000, 28000, 28000, 29000],
    ['Interest & Scrap Recovery Income', 'Indirect Incomes', 180000, 140000, 150000, 15000, 15000, 15000, 15000, 15000, 15000, 15000, 15000, 15000, 15000, 15000, 15000],
    ['Corporate Income Tax Provision', 'Tax Expense', 740000, 440000, 680000, 45000, 54000, 57000, 57000, 62000, 67000, 70000, 60000, 67000, 64000, 63000, 74000],
    ['Net Profit', 'Subtotal', 2220000, 1330000, 2040000, 130000, 163000, 172000, 173000, 188000, 203000, 212000, 179000, 202000, 190000, 190000, 218000],
  ];

  const cfRows = [
    ['Cash Flow Particulars', 'Activity Category', 'Current Year (FY 2025-26)', 'Previous Year (FY 2024-25)', 'Budget (FY 2025-26)'],
    ['Cash Generated from Operating Activities', 'Operating Activities', 3150000, 2480000, 3000000],
    ['Direct Income Taxes Paid', 'Operating Activities', -520000, -410000, -480000],
    ['Capital Expenditure on CNC Plant & Tooling', 'Capital Expenditure', -1150000, -850000, -1200000],
    ['Interest & Investment Income Received', 'Investing Activities', 95000, 72000, 80000],
    ['Term Loan Repayment & Finance Costs Paid', 'Financing Activities', -640000, -590000, -640000],
    ['Opening Cash and Cash Equivalents', 'Opening Cash', 1415000, 713000, 1415000],
    ['Closing Cash and Cash Equivalents', 'Closing Cash', 2350000, 1415000, 2175000],
  ];

  if (kind === 'BALANCE_SHEET') {
    XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet(bsRows), 'Balance Sheet');
    XLSX.writeFile(wb, 'GSPU_Sample_Balance_Sheet_FY25_26.xlsx');
  } else if (kind === 'PROFIT_LOSS') {
    XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet(pnlRows), 'Profit and Loss');
    XLSX.writeFile(wb, 'GSPU_Sample_Profit_And_Loss_FY25_26.xlsx');
  } else if (kind === 'CASH_FLOW') {
    XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet(cfRows), 'Cash Flow Statement');
    XLSX.writeFile(wb, 'GSPU_Sample_Cash_Flow_Statement_FY25_26.xlsx');
  } else {
    XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet(pnlRows), 'Profit and Loss');
    XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet(bsRows), 'Balance Sheet');
    XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet(cfRows), 'Cash Flow Statement');
    XLSX.writeFile(wb, 'GSPU_Sample_CFO_Monthly_Pack_3_Statements.xlsx');
  }
}

