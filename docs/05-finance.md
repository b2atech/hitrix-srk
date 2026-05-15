# Domain 5: Finance & Accounting

## Purpose
Manage cash and bank transactions, journal entries, bank reconciliation, outstanding tracking, and final accounts preparation. This is the core accounting engine — all transactions from Sales, Purchase, and other modules ultimately post through here.

---

## Key Features

### 5.1 Payment Voucher (`frmPayment`, `frmPaymentGST`)
- Cash or bank payment to suppliers/parties
- Party/Supplier account selection
- Against-bill reference (invoice number and date)
- Cheque number, date, and bank
- Discount / adjustment amounts at payment time
- TDS deduction at payment (if not deducted at bill stage)
- Multiple bills settled in one payment
- GST variant: CGST/SGST/IGST for service payments

### 5.2 Receipt Voucher (`frmReceipt`, `frmReceiptGST`, `frmReceiptMill`)
- Cash or bank receipt from customers/parties
- Party/Customer account selection
- Against-bill reference
- Cheque / demand draft / NEFT / RTGS details
- Interest and other charges on receipt
- TCS (Tax Collected at Source) flag
- Mill receipt variant: Separate receipt for mill-related transactions

### 5.3 Journal Voucher (`frmJV`, `frmJVGST`)
- General ledger journal entry
- Debit and credit account selection
- Up to 4 narration lines
- Narration code linking (auto-selects tax accounts)
- Aging analysis capability
- GST variant: Credit/Debit note for GST tax adjustments

### 5.4 Payment-Receipt Combined (`frmPayRec`)
- Combined view for managing payment and receipt together
- Used for quick reconciliation scenarios

### 5.5 Bank Reconciliation (`frmBankReConcilEntry`)
- Match bank statement to ledger entries
- Outstanding cheques list (issued but not cleared)
- Uncleared receipts (deposited but not credited)
- Bank-only transactions (bank charges, interest)
- Reconciliation statement with difference calculation

### 5.6 Multi-Account Journal (`frmMultiAccount`)
- Multi-line journal entry (single narration, multiple debits/credits)
- For complex transactions like salary, expense distribution

### 5.7 Final Accounts (`frmRptFinal`)
- Profit & Loss Statement
  - Gross profit / loss calculation
  - Income accounts summary
  - Expense accounts summary
- Balance Sheet
  - Assets side (current, fixed, loans)
  - Liabilities side (capital, loans, creditors)
  - Equity reconciliation
- Driven by account group hierarchy (TopGroup classification)

### 5.8 Outstanding Management (cross-cutting)
- Automatic creation on bill posting (sales/purchase)
- Aging buckets: 0–30, 30–60, 60–90, 90+ days
- Bill-to-bill reconciliation at payment time
- Outstanding debit note adjustments
- Late payment interest calculation trigger

---

## Key Data Captured

| Field | Description |
|---|---|
| `VNo`, `VDt` | Voucher number and date |
| `VAcCode` | Debit account code |
| `VAcOCode` | Credit account code |
| `Vamt` | Voucher amount |
| `VBillNo`, `VBillDt` | Reference bill/invoice number and date |
| `VNar1`–`VNar4` | Up to 4 narration lines |
| `ChqNo`, `ChqDt` | Cheque number and date |
| `VCGSTRt`, `VCGSTAmt` | CGST rate and amount |
| `VSGSTRt`, `VSGSTAmt` | SGST rate and amount |
| `VIGSTRt`, `VIGSTAmt` | IGST rate and amount |
| `IsDiscount` | Discount flag at payment |
| `TDSRate`, `TDSAmt` | TDS details at payment |
| `IsTCSApplicable` | TCS flag |
| `VIsAudited` | Audit lock flag |
| `vUserName` | User who created the transaction |
| `VReconDt` | Bank reconciliation date |

---

## Business Rules & Logic

- **Double-entry mandatory:** Every voucher requires a debit account and a credit account
- **Bill-to-bill settlement:** At payment time, specific invoices are selected and amounts allocated; partial settlement is allowed
- **Discount at payment:** Cash discount or settlement discount entered at payment time (not on invoice)
- **TDS at source:** Deducted at payment if not already deducted at bill stage; posts to TDS payable account
- **Bank reconciliation:** Tracks cleared vs uncleared cheques by comparing ledger to bank statement
- **Final accounts:** P&L and BS driven by account group `TopGroup` classification (Liability/Asset/P&L/Stock)
- **Account interest:** Monthly interest charged at percentage on outstanding account balance
- **Audit lock:** Once `VIsAudited = True`, transactions cannot be modified/deleted without password protection
- **Fiscal year boundary:** Cannot post to a closed/locked fiscal year
- **Opening stock:** Handled via special account (TopGroup 90018) and separate opening entry
- **Cheque bounce:** Requires reversal voucher (JV) with bank charges

---

## Reports Available

| Report | Description |
|---|---|
| `rptCashBankStatus.rpt` | Cash and bank current balances |
| `rptOutStandingSALatePayInt.rpt` | Outstanding with late payment interest |
| `rptOutstBrokerPartywise.rpt` | Outstanding — broker and party wise |
| `rptOutstBrokerwise.rpt` | Outstanding — broker wise |
| `rptOutstBrokerwiseWhatsApp.rpt` | Outstanding summary for WhatsApp |
| `rptOutstPartywiseWhatsApp.rpt` | Party outstanding for WhatsApp |
| `rptOutStandingAgingBank.rpt` | Aging analysis with bank details |
| `rptOutstGST.rpt` | GST outstanding |
| `rptSMSOsBroker.rpt` | Outstanding SMS — broker wise |
| `rptSMSOsParty.rpt` | Outstanding SMS — party wise |
| Account Ledger (frmRptAccount) | Detailed party ledger with running balance |
| Group Ledger (frmRptAccount) | Summary by account group |
| Party Confirmation | Account confirmation letter with balance |
| Final Accounts (frmRptFinal) | P&L and Balance Sheet |

---

## Integration Points

| Integrates With | How |
|---|---|
| Sales | Receipt voucher settles sales outstanding |
| Purchase | Payment voucher settles purchase outstanding |
| Tax Compliance | TDS/TCS entries affect payable accounts |
| Bank Reconciliation | Cheque-level matching against bank statement |
| Reporting | All GL entries feed into ledger reports and final accounts |
| Communication | Outstanding summaries sent via WhatsApp/SMS to parties/brokers |
