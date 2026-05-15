# Domain 4: Purchase (Pre-GST and GST Variants)

## Purpose
Record commodity purchase transactions with full GST/VAT support, TDS management, goods inward tracking, and outstanding payables management. Handles multiple purchase structures common in textile trading.

---

## Key Features

### 4.1 Purchase Bill Entry — Pre-GST Variants (`frmPurchaseTrade`, `frmPurchaseOther`, `frmPurchaseSIT`)

| Type Code | Variant | Description |
|---|---|---|
| `PY` | Trade Purchase | Direct purchase from supplier |
| `PT` | SIT Purchase | Stock-in-transit purchase (goods at supplier's godown) |
| `PO` | Other Purchase | Miscellaneous / non-commodity purchases |
| `PYHNK` | Hank Purchase | Purchase of yarn hanks (textile-specific) |

### 4.2 Purchase Bill Entry — GST Variants (`frmPurchaseTradeGST`, `frmPurchaseOtherGST`)
- GST-compliant purchase bill entry
- Supplier invoice number capture
- CGST, SGST, IGST rates and amounts
- HSN code
- IRN number (for e-invoicing eligible suppliers)
- Reverse Charge Mechanism (RCM) flag for unregistered vendors
- TDS calculation at bill level (or deferred to payment)

### 4.3 Goods Inward / Purchase Inward (`frmPurchaseInward`)
- Receipt of goods into company godown
- Bag / cartoon level tracking
- Weight reconciliation (bill weight vs actual received weight)
- Taxable inward flag (for import duty scenarios)
- Auto-linkage to purchase bill reference
- Supplier lorry receipt (LR) number capture
- Inward date and time

### 4.4 Purchase Return (`frmPurchaseReturn`, `frmPurchRtnTradeGST`)
- Return of goods to supplier
- Debit note generation
- Links to original purchase bill
- Reversal of tax posted on original purchase
- Weight and quantity reconciliation

### 4.5 TDS Entry (`frmTDSEntry`)
- Tax Deducted at Source on eligible payments
- TDS rate by category (e.g., 194C for contractors, 194Q for purchases above threshold)
- TDS amount auto-calculated or manual override
- TDS account posting (TDS payable)
- TDS certificate number tracking
- TDS circle / jurisdiction
- Deductee PAN capture
- Quarterly TDS return support data

---

## Key Data Captured

| Field | Description |
|---|---|
| `PurchInvNo`, `PurchDt` | Supplier invoice number and date |
| `VType` | Purchase type: PY / PT / PO / PYHNK |
| `Supplier` | Vendor/supplier account code |
| `Broker` | Commission agent account code |
| `Qty`, `Weight` | Quantity purchased, weight |
| `Rate`, `Amt` | Unit rate and total amount |
| `NarrCode` | Narration code (drives tax posting) |
| `CGST%`, `CGST Amt` | CGST rate and amount |
| `SGST%`, `SGST Amt` | SGST rate and amount |
| `IGST%`, `IGST Amt` | IGST rate and amount |
| `IsRCM` | Reverse charge mechanism flag |
| `TDSRate`, `TDSAmt` | TDS rate and amount deducted |
| `TDSCircle` | TDS jurisdiction circle |
| `TDSCertNo` | TDS certificate / challan number |
| `IsTaxableInward` | Import duty applicable flag |
| `BillBags`, `BillWeight` | Total bags and weight per bill |
| `LRNo`, `LRDt` | Lorry receipt number and date |
| `VIsAudited` | Audit lock flag |
| `vUserName` | User who created the transaction |

---

## Business Rules & Logic

- **SIT goods:** Held in supplier's warehouse; purchase liability created when goods are in transit before physical receipt
- **TDS:** Mandatory for payments above statutory threshold; posts to TDS liability (payable) account
- **RCM:** When vendor is unregistered under GST, purchasing firm pays GST directly to government; posts to RCM payable account
- **Inward tracking:** Bag-level details captured for weight reconciliation (bill weight vs actual)
- **Purchase return:** Reduces purchase payable and reverses input tax credit (ITC)
- **Brokerage:** Calculated and posted separately; paid to agent on commission basis
- **Outstanding:** Purchase amounts tracked in outstanding module until payment
- **Debit notes:** For price adjustments, quality defects, or short supplies
- **TDS at payment:** If not deducted at bill, TDS is deducted at the time of payment voucher
- **Multiple inwards against one bill:** One purchase bill can have multiple inward entries (partial deliveries)
- **Weight shortfall:** Difference between billed weight and actual received weight tracked and adjusted

---

## Reports Available

| Report | Description |
|---|---|
| `rptPurchaseSummMillwise.rpt` | Purchase summary — mill wise |
| `rptPurchaseSummCountwise.rpt` | Purchase summary — count (unit) wise |
| `rptPurchaseSummMonthwise.rpt` | Purchase summary — month wise |
| `rptPurchSummMonthwiseNew.rpt` | Purchase summary — monthly (revised) |
| `rptPurchaseSummTotal.rpt` | Purchase total summary |
| `rptPurchaseSalesSummMonthwise.rpt` | Combined purchase and sales monthly |
| `rptOutStandingSALatePayIntPurch.rpt` | Outstanding purchase with interest |
| `rptOutstPurchPartywise.rpt` | Outstanding purchase — party wise |
| `rptTaxRptCFormRec.rpt` | C-Form received (pre-GST compliance) |
| `rptTaxRptE1FormRec.rpt` | E1-Form received |
| `rptTaxRptE2FormIssu.rpt` | E2-Form issued |

---

## Integration Points

| Integrates With | How |
|---|---|
| Finance / Accounting | Purchase posts cost and input tax credit to GL |
| Outstanding | Bill amount and due date recorded until payment |
| Inventory / Logistics | Goods inward linked to bag entry and gate pass |
| Tax Compliance | GST/TDS calculated at bill or payment stage |
| Booking | Purchase can be linked to prior booking reference |
| Communication | Purchase debit notes sent to supplier via WhatsApp |
| Reporting | Purchase registers and summaries from transaction data |
