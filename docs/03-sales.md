# Domain 3: Sales (Pre-GST and GST Variants)

## Purpose
Record commodity sales transactions across multiple trade structures (trade, consignment, depot, stock-in-transit, mill bills) with full tax accounting, outstanding tracking, and material flow linkage.

---

## Key Features

### 3.1 Sales Bill Entry — Pre-GST Variants (`frmSales`, `frmSalesTrade`, `frmSalesMillBill`)

| Type Code | Variant | Description |
|---|---|---|
| `SY` | Trade Sale | Direct sale to customer from company stock |
| `SO` | Consignment Sale | Sale from consignment stock held on behalf of supplier |
| `SD` | Depot Sale | Sale from depot / third-party warehouse arrangement |
| `ST` | SIT Sale | Sale of stock-in-transit goods |
| `SYHNK` | Hank Sale | Sale of yarn hanks (textile-specific variant) |
| `SM` | Mill Bill | Processing fee invoice to customer (mill job work) |

### 3.2 Sales Entry — GST Variants (`frmSales GST`, `frmSalesOtherGST`)
- GST-compliant invoicing
- Invoice number generation (auto or manual with prefix)
- CGST, SGST, IGST calculation with rates and amounts
- HSN code entry
- IRN (Invoice Reference Number) for e-invoicing
- TCS (Tax Collected at Source) flag
- Taxable amount separation from exempted/nil-rated

### 3.3 Sales Return (`frmSalesReturn`, `frmSaleRtnTradeGST`)
- Return of goods from customer
- Credit note generation
- Linkage to original invoice
- Reversal of tax posted on original sale

### 3.4 Sales Confirmation (`frmConfirmationSales`)
- Formal confirmation sent to customer after despatch
- Links to gate pass and delivery

### 3.5 Late Payment Interest Debit Note (`frmLatePayIntDbNt`)
- Auto-calculated interest on overdue receivables
- Debit note raised on customer account
- Interest rate configurable per account
- Outstanding ageing drives calculation

### 3.6 Credit / Debit Note GST (`frmCrDrNoteGST`)
- GST-compliant credit and debit note
- For price adjustments, quality claims, disputes
- Full CGST/SGST/IGST reversal or additional posting

---

## Key Data Captured

| Field | Description |
|---|---|
| `InvNo`, `InvDt` | Invoice number and date |
| `VType` | Sale type: SY / SO / SD / ST / SM / SYHNK |
| `Party` | Customer account code |
| `Broker` | Commission agent account code |
| `Mill` | Processing mill account code (for SM type) |
| `Qty`, `Weight` | Quantity sold, weight |
| `Rate`, `Amt` | Unit rate and total amount |
| `NarrCode` | Narration code (drives tax posting) |
| `TaxCode` | Tax narration for liability posting |
| `CGST%`, `CGST Amt` | CGST rate and amount |
| `SGST%`, `SGST Amt` | SGST rate and amount |
| `IGST%`, `IGST Amt` | IGST rate and amount |
| `HSNCode` | HSN code for GST |
| `IRNNo` | E-invoice reference number |
| `IsRefEntryType` | Reverse charge mechanism flag |
| `IsTCSDeductible` | TCS applicability flag |
| `VBillDt` | Bill date (may differ from invoice date) |
| `VReconDt` | Reconciliation date |
| `VIsAudited` | Audit lock flag (no modification after audit) |
| `vUserName` | User who created the transaction |

---

## Business Rules & Logic

- **Invoice numbering:** Can be purely numeric or with user-defined prefix (e.g., "INV/25-26/001")
- **Tax posting:** Separate voucher created automatically for GST liability (if VAT flag enabled in Settings)
- **Hank sales:** Tracked separately with special rate cards; rate per hank/piece rather than per kg
- **Mill bill (SM):** Shows processing fees as revenue; links to party ledger for the mill
- **Commission:** Broker commission calculated separately as percentage of bill amount (Commission, Cash Discount, or Sales Incentive type)
- **SIT goods:** Only sold if goods are physically received under SIT arrangement
- **Consignment:** Sale reduces consignment liability, adds revenue; consignor account credited
- **Outstanding tracking:** Bill number and date linked to Outstanding module automatically on save
- **Late payment interest:** Auto-calculated based on interest % on account master and overdue days
- **Audit lock:** Once `VIsAudited = True`, transaction cannot be modified or deleted without password
- **IRN / e-invoicing:** IRN number must be obtained from GST portal and entered manually for eligible invoices
- **TCS:** Tax Collected at Source for specified goods (textiles above threshold); deducted from payment

---

## Reports Available

| Report | Description |
|---|---|
| `rptSaleBillPrint.rpt` | Sales bill print format (customer-facing invoice) |
| `rptSalesSummPartywise.rpt` | Sales summary — party wise |
| `rptSalesSummMillwise.rpt` | Sales summary — mill wise |
| `rptSalesSummBrokerPartywise.rpt` | Sales summary — broker and party wise |
| `rptSalesSummBrokerCountwise.rpt` | Sales summary — broker and count wise |
| `rptSalesSummCountwise.rpt` | Sales summary — count (unit) wise |
| `rptSalesSummMillCountwise.rpt` | Sales summary — mill and count wise |
| `rptSalesSummMonthwise.rpt` | Sales summary — month wise |
| `rptSalesSummTotal.rpt` | Sales total summary |
| `rptPurchaseSalesSummMonthwise.rpt` | Combined purchase and sales monthly |
| `rptLatePayDbNtReg.rpt` | Late payment debit note register |
| `rptOutStandingSALatePayInt.rpt` | Outstanding with late payment interest |
| `rptOutstBrkPartyWithUnadjustRec.rpt` | Outstanding with unadjusted receipts |

---

## Integration Points

| Integrates With | How |
|---|---|
| Finance / Accounting | Sale posts revenue to GL; tax posts to GST liability account |
| Outstanding | Bill amount and due date recorded in outstanding module |
| Inventory / Logistics | Goods outward linked to gate pass |
| Tax Compliance | GST/CST applicable based on narration and buyer state |
| Booking | Sales can be linked to prior booking reference |
| Communication | Sales bill PDF sent to customer via WhatsApp |
| Reporting | Sales registers and summaries generated from transaction data |
