# Domain 7: Tax Compliance (CST, GST, TDS, TCS)

## Purpose
Manage statutory tax compliance for Indian commodity trading — covering pre-GST era (CST forms: C/E1/E2/F) and the current GST era (CGST/SGST/IGST, RCM, IRN, e-invoicing). Also handles TDS (Tax Deducted at Source) and TCS (Tax Collected at Source) applicable to textile/commodity trading.

---

## Key Features

### 7.1 Pre-GST CST Form Management (`frmFormEntry`)
Track statutory Central Sales Tax forms required for inter-state and consignment transactions:

| Form Type | Description |
|---|---|
| C-Form | Tax exemption certificate for inter-state purchases (buyer gives to seller) |
| E1-Form | First sale in transit; exporter gives to intermediate buyer |
| E2-Form | Second sale in transit; used in multi-party chain transactions |
| F-Form | Stock transfer (not sale) between branches across states |

- Form number entry, type classification, and linked bill reference
- Form date and expiry date tracking
- Validity enforcement: System flags expired or expiry-approaching forms
- Outstanding C-Form tracking (received vs required)

### 7.2 GST Compliance

**Sales Side:**
- CGST, SGST, IGST calculation at applicable rates per item
- HSN code capture at transaction level
- IRN (Invoice Reference Number) for e-invoicing eligible invoices
- TCS collection tracking
- Separate GST registers for GSTR-1 data extraction

**Purchase Side:**
- Input Tax Credit (ITC) tracking via CGST/SGST/IGST amounts
- Reverse Charge Mechanism (RCM) for unregistered vendors
- HSN code capture at bill level
- IRN entry for supplier e-invoices
- GSTR-2B reconciliation support

**GST Reports (`frmRptGST`):**
- Sales GST Register (GSTR-1 support)
- Purchase GST Register (GSTR-2 support)
- Outward liability summary
- Inward credit summary
- RCM purchase totals

### 7.3 TDS Compliance (`frmTDSEntry`)
- TDS categories applicable in textile/commodity trading:
  - 194C: Payments to contractors (job work / mill processing)
  - 194Q: Purchases above ₹50 lakh threshold (new provision)
- TDS rate and amount calculation
- Deductee PAN number capture
- TDS certificate / challan number tracking
- TDS circle / jurisdiction details
- Quarterly TDS return support (data for Form 26Q)

### 7.4 TCS Management
- TCS on sale of specified goods above ₹50 lakh threshold
- TCS firm flag per customer account
- TCS receivable account (TcsRecAcCode in Settings)
- TCS deductee master
- TCS collection register

---

## Key Data Captured

| Field | Description |
|---|---|
| `FormNo` | CST form number |
| `FormType` | C / E1 / E2 / F |
| `FormDt`, `FormExpDt` | Form issue date and expiry date |
| `FormBillRef` | Linked purchase or sale bill number |
| `HSNCode` | HSN code for goods classification |
| `SACCode` | SAC code for services (mill processing) |
| `CGST%`, `CGSTAmt` | CGST rate and amount |
| `SGST%`, `SGSTAmt` | SGST rate and amount |
| `IGST%`, `IGSTAmt` | IGST rate and amount |
| `IsRCM` | Reverse charge flag |
| `IRNNo` | E-invoice reference number |
| `TDSRate`, `TDSAmt` | TDS rate and deducted amount |
| `TDSCertNo` | TDS certificate / challan number |
| `TDSCircle` | TDS jurisdiction circle |
| `DeducteePAN` | Vendor/supplier PAN for TDS |
| `IsTcsFirm` | TCS applicability flag on account |
| `TcsRecAcCode` | TCS receivable account code |

---

## Business Rules & Logic

**CST Forms (Pre-GST):**
- C-Form must be collected from buyer for concessional rate sales (CST 2% vs full rate)
- E1-Form required for first in-transit sale
- F-Form required for stock transfers (not sales) to own branches
- System tracks: forms required per bill vs forms actually received
- Forms linked 1:1 to eligible bills for audit trail
- Expiry validation prevents use of expired forms

**GST:**
- CGST + SGST for intra-state transactions (same state buyer and seller)
- IGST for inter-state transactions (different state buyer and seller)
- Rate determined by HSN code and applicable notification
- RCM: Company pays GST directly when buying from unregistered vendor
- IRN: Required for e-invoicing eligible entities (turnover > ₹5 crore); must be obtained from GST portal
- Input Tax Credit (ITC): Available on GST paid on purchases (except for certain restricted categories)

**TDS:**
- Threshold-based: TDS applicable only above statutory limit
- TDS at bill or at payment (whichever is earlier is the trigger)
- Lower deduction / nil deduction: If vendor provides certificate, reduced TDS applied
- Annual Form 26Q filing: Data tracked in system to support quarterly filings

**TCS:**
- Applies on sale of certain specified goods above ₹50 lakh in a financial year per buyer
- TCS rate: 0.1% on excess above ₹50 lakh (textile materials qualify)
- TCS collected at invoice time, deposited quarterly

---

## Reports Available

| Report | Description |
|---|---|
| `rptTaxRptCFormIssue.rpt` | C-Form issued to sellers |
| `rptTaxRptCFormIssueSumm.rpt` | C-Form issued — summary |
| `rptTaxRptCFormIssueShort.rpt` | C-Form issued — short version |
| `rptTaxRptCFormIssueLtr.rpt` | C-Form issue letter |
| `rptTaxRptCFormRec.rpt` | C-Form received from buyers |
| `rptTaxRptCFormRecPartywise.rpt` | C-Form received — party wise |
| `rptTaxRptCFormRecPartywiseSumm.rpt` | C-Form received — summary |
| `rptTaxRptCFormRecPendingSlabSumm.rpt` | C-Form pending by slab |
| `rptTaxRptCFormRecSumm.rpt` | C-Form received — total summary |
| `rptTaxRptCFormRequire.rpt` | C-Form required (not yet received) |
| `rptTaxRptCFormRequireNew.rpt` | C-Form required — updated |
| `rptTaxRptCFormRequirePartyDet.rpt` | C-Form required — party detail |
| `rptTaxRptE1CovrLtr.rpt` | E1-Form cover letter |
| `rptTaxRptE1FormRec.rpt` | E1-Form received |
| `rptTaxRptE1FormRecSumm.rpt` | E1-Form received — summary |
| `rptTaxRptE1RecSumBrokerwise.rpt` | E1-Form received — broker wise |
| `rptTaxRptE1RecSumPartywise.rpt` | E1-Form received — party wise |
| `rptTaxRptE1FormRecPendingSlabSumm.rpt` | E1-Form pending by slab |
| `rptTaxRptE2FormIssu.rpt` | E2-Form issued |
| `rptTaxRptE2FormIssuSumm.rpt` | E2-Form issued — summary |
| `rptTaxRptE2FormIssuPartywise.rpt` | E2-Form issued — party wise |
| `rptTaxRptE2FormIssuPartywiseSumm.rpt` | E2-Form issued — party wise summary |
| `rptTaxRptE2FormRec.rpt` | E2-Form received |
| `rptTaxRptE2FormRecSumm.rpt` | E2-Form received — summary |
| `rptTaxRptE2FormRequireInvoicewise.rpt` | E2-Form required — invoice wise |
| `rptTaxRptFFormIssue.rpt` | F-Form issued |
| `rptTaxRptFFormIssueSumm.rpt` | F-Form issued — summary |
| `rptTaxRptFFormRequire.rpt` | F-Form required |
| `rptTaxRptFFormRequireNew.rpt` | F-Form required — updated |
| `rptTaxRptFFormRequirePartyDet.rpt` | F-Form required — party detail |
| `rptTaxRptCFormIssuedE1NotRecd.rpt` | C-Form issued but E1 not received |
| GST Register (frmRptGST) | GSTR-1 / GSTR-2 format data |

---

## Integration Points

| Integrates With | How |
|---|---|
| Sales | GST/TCS calculated at invoice; forms required per inter-state sale |
| Purchase | GST/TDS calculated at bill; C-Form received from eligible buyers |
| Finance | TDS/TCS account postings in GL |
| Masters | Tax narration master provides rate and account mapping |
| Reporting | Tax registers exported for GST return preparation |
