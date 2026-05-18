# HITRIX V2 ↔ Dhanman Comparison Table

**Purpose:** Map HITRIX V2 legacy tables to Dhanman microservices, identify coverage and gaps.

**Updated:** 2026-05-15 | **Deadline:** 2026-07-15 (8 weeks)

---

## 🎯 DEVELOPER QUICK SUMMARY

**Bottom Line:** Dhanman covers 60% of HITRIX functionality. To go live in Wave 1, **5 critical features** MUST be built (2-3 weeks effort). Without these, the system **cannot generate compliant invoices, calculate taxes correctly, or post GL entries**.

### What Works ✅
- Company & Master data (users, accounts, groups) — mostly ready
- Basic Order/Invoice creation — exists but incomplete
- Inventory basics — stock tracking present (but not bag tracking)

### What's Broken ❌ (Must Fix for Wave 1)

| Gap | Impact | Fix Effort | Status |
|-----|--------|-----------|--------|
| **No IRN generation** | Invoices non-GST-compliant; cannot legally issue | 5 days | NOT STARTED |
| **No RCM/TCS tax engine** | Tax calculations wrong; GSTR mismatches | 5-6 days | NOT STARTED |
| **Account master missing fields** | Credit limits don't work; consignment broken | 3 days | NOT STARTED |
| **No auto-GL posting** | Manual journal entries required; AR/AP out of sync | 4 days | NOT STARTED |
| **No price validation** | User entry errors (sell at cost, etc.) | 2-3 days | NOT STARTED |

**Total Wave 1 Build Effort:** ~22-24 days (4.5 weeks). All 5 must complete before go-live UAT.

### What Can Wait (Wave 2) ⏱️
- Interest on late payment (can do manual workaround)
- Bank reconciliation (can do Excel)
- Booking pre-orders (only if pilot needs)
- Bag tracking (physical inventory detail)
- Year-end closing procedures

---

---

## Master Tables

| HITRIX Table | Purpose | Dhanman Service | Dhanman Entity | Confidence | Status | Gaps |
|---|---|---|---|---|---|---|
| **tblMastCompany** | Firm/Company profile (CCode, CName, CGSTIN, bank details, firm type flags) | dhanman-common | Company | 100% | ✅ Present | None identified |
| **tblCompanyYear** | Multi-year partition (financial year, code ranges per year) | dhanman-common | CompanyYear / Period | 75% | ⚠️ Partial | Code ranges may not be needed; Dhanman may use auto-increment |
| **tblMastGroup** | Account groups/chart of accounts Level 1 (AgCode, hierarchy) | dhanman-finance | AccountGroup | 100% | ✅ Present | None identified |
| **tblMastSubGroup** | Sub-groups Level 2 | dhanman-finance | AccountSubGroup | 100% | ✅ Present | None identified |
| **tblMastCompGroup** | Corporate group (multi-company firms) | dhanman-common | CorporateGroup | 100% | ✅ Present | None identified |
| **tblMastAccount** | Ledger accounts (parties, suppliers, banks, expense heads) — 72 columns | dhanman-common + dhanman-sales + dhanman-purchase | Account / Party / Customer / Vendor | 75% | ⚠️ Needs Update | Missing: AcIntPer (interest rate), AcOSLimit (credit limit), AcMillType (transaction type: trade, consignment, depot, mill), AcConsignmentFirm, AcDepotFirm, AcIsPayToMill, TDS flags |
| **tblMastItem** | Products/commodities (name, unit, HSN, GST, brokerage rate) | dhanman-inventory | Product | 75% | ⚠️ Partial | Core product master exists, but end-user price guardrails are missing. Recommend adding `MinSellingPrice` and `MaxPurchasePrice` plus purchase/sales validation to reduce pricing mistakes. |
| **tblMastNarration** | Multi-purpose lookup: Tax codes (T), Godowns (G), Carriers (R), Charges (C) | dhanman-finance + dhanman-inventory | TaxCode / Godown / Carrier / Charge | 75% | ⚠️ Partial | GST rate fields (CGST, SGST, IGST) need to be checked; RCM/TCS account links may be missing |
| **tblMastSetting** | System configuration (default GL accounts, tax codes per VType, GST accounts) | dhanman-common + dhanman-finance | SystemConfiguration | 50% | ❌ Partial | May not exist as single table; likely distributed across services |
| **tblMastTransactionCode** | Transaction type codes (SY, PY, ST, PT, etc.) | dhanman-sales + dhanman-purchase | TransactionType | 75% | ⚠️ Partial | Dunman/Dhanman may use VType field or separate entity |
| **tblMastUser** | User master (UserName, Password, Firm access) | dhanman-common | User | 100% | ✅ Present | None identified |
| **tblUserPrevilage** | User permissions | dhanman-common | UserPermission | 100% | ✅ Present | None identified |
| **tblMastUserCompany** | User-to-Company access mapping | dhanman-common | UserCompanyAccess | 100% | ✅ Present | None identified |

**Master Tables Summary:** ✅ 7 tables (100%), ⚠️ 6 tables (75%), ❌ 1 table (50%)

---

## Transaction Tables

| HITRIX Table | Purpose | Dhanman Service | Dhanman Entity | Confidence | Status | Gaps |
|---|---|---|---|---|---|---|
| **tblBooking** | Trade booking pre-order (BK type, tblBooKingSub for line items) | dhanman-sales | Order / Booking (status=booking?) | 50% | ❓ Unknown | Pre-order workflow unclear; may use Order with status field instead |
| **tblBooKingSub** | Booking line items | dhanman-sales | OrderItem / BookingItem | 50% | ❓ Unknown | Depends on tblBooking implementation |
| **tblBooKingDesp** | Booking despatch records (partial shipments) | dhanman-sales | Delivery / Despatch | 50% | ❓ Unknown | May be merged into Order/Delivery entity |
| **tblSale** | Sales orders (SY, SO, SD, ST, SM, RY types) — multi-purpose | dhanman-sales | Order / Invoice | 75% | ⚠️ Partial | Missing: SlIRNNo (e-Invoice IRN), SlAckNo (IRN ack), interest link (tblIntSale) |
| **tblSaleSub** | Sales line items (qty, rate, godown, GST, bag tracking) | dhanman-sales | OrderItem / InvoiceItem | 75% | ⚠️ Partial | Bag tracking link (SlSubItCtrlNo → tblBags) may not exist |
| **tblIntSale** | Late payment interest on sales (daily interest calc) | dhanman-finance | Interest / DebitNote | 0% | ❌ Missing | Auto-interest calculation likely not in Dhanman |
| **tblPurch** | Purchase orders (PY, PT, PI, PO, VY types) | dhanman-purchase | PurchaseOrder / Invoice | 75% | ⚠️ Partial | Missing: TDS fields, interest link (tblIntDbNote) |
| **tblPurchSub** | Purchase line items | dhanman-purchase | PurchaseOrderItem | 75% | ⚠️ Partial | Bag tracking link may not exist |
| **tblIntDbNote** | Debit note / late payment interest on purchases | dhanman-finance | DebitNote / Interest | 0% | ❌ Missing | Auto-calculation likely not in Dhanman |
| **tblGatePass** | Gate pass / goods issue (GP type, inventory shipment doc) | dhanman-inventory | GoodsIssue / Delivery | 75% | ⚠️ Partial | Implementation level unclear |
| **tblGatePassSub** | Gate pass line items | dhanman-inventory | GoodsIssueItem | 75% | ⚠️ Partial | Depends on tblGatePass |
| **tblVoucher** | Finance vouchers (CP=cash payment, BP=bank payment, CR=cash receipt, BR=bank receipt, JV=journal) | dhanman-finance | JournalEntry / Payment / Receipt | 75% | ⚠️ Partial | May need to split by VType into separate entities |
| **tblBags** | Bag/physical goods tracking (lifecycle: inward → godown → sale) | dhanman-inventory | StockTracking / BagMaster | 0% | ❌ Missing | Bag lifecycle tracking not found in Dhanman |

**Transaction Tables Summary:** ✅ 0 tables (100%), ⚠️ 6 tables (75%), ❌ 3 tables (0%)

---

## Finance Tables

| HITRIX Table | Purpose | Dhanman Service | Dhanman Entity | Confidence | Status | Gaps |
|---|---|---|---|---|---|---|
| **tblVoucher** (detail) | GL journal entries with multi-level detail | dhanman-finance | JournalEntry | 75% | ⚠️ Partial | Auto-creation from sales/purchase invoices unclear |
| Interest on Outstanding | Debit notes for late payment (via tblIntSale, tblIntDbNote) | dhanman-finance | Interest / DebitNote | 0% | ❌ Missing | Daily interest calculation not found |
| Bank Reconciliation | Match bank statement to vouchers | dhanman-finance | BankReconciliation | 50% | ❓ Unknown | Feature level unclear |

**Finance Tables Summary:** ✅ 0 tables, ⚠️ 1 table (75%), ❌ 1 table (0%), ❓ 1 table

---

## GST & Tax Tables

*Note: GST fields are distributed across multiple tables in HITRIX. In Dhanman, they may be consolidated or spread differently.*

| HITRIX Location | Purpose | Dhanman Service | Dhanman Implementation | Confidence | Gaps |
|---|---|---|---|---|---|
| tblMastAccount.AcGSTIN | Party GST registration number | dhanman-common | Account.gstIn | 100% | ✅ Present |
| tblMastNarration (T type) | Tax rates (CGST, SGST, IGST) per commodity/charge | dhanman-finance | TaxCode / TaxRule | 75% | ⚠️ Rates may need mapping |
| tblSale.SlGSTIn | GST amount on sale invoice | dhanman-sales | Order/Invoice.gstAmount | 75% | ⚠️ Calculated or stored? |
| tblPurch.PurGSTIn | GST on purchase invoice | dhanman-purchase | PurchaseOrder.gstAmount | 75% | ⚠️ Calculated or stored? |
| RCM Logic | Reverse Charge Mechanism (unregistered parties) | dhanman-sales | RCMFlag / RCMAmount | 0% | ❌ Missing |
| TCS Logic | Tax Collected at Source (1% on sales) | dhanman-sales | TCSAmount | 0% | ❌ Missing |
| TDS Logic | Tax Deducted at Source (on payments) | dhanman-finance / dhanman-purchase | TDSAmount | 0% | ❌ Missing |
| E-Invoice IRN | Invoice Registration Number (IRP API) | dhanman-sales | Invoice.irnNumber / Invoice.irnQrCode | 0% | ❌ Missing |
| GSTR2A Matching | Match purchase invoices to GST portal | dhanman-purchase | GSTR2AMatch | 0% | ❌ Missing |

**GST & Tax Summary:** ✅ 1 (100%), ⚠️ 3 (75%), ❌ 5 (0%)

---

## Inventory & Warehouse Tables

| HITRIX Table | Purpose | Dhanman Service | Dhanman Entity | Confidence | Status | Gaps |
|---|---|---|---|---|---|---|
| **tblGatePass** | Goods issue / shipment document | dhanman-inventory | GoodsIssue / Delivery | 75% | ⚠️ Partial | Godown/warehouse tracking level unclear |
| **tblGatePassSub** | Line items on gate pass | dhanman-inventory | GoodsIssueItem | 75% | ⚠️ Partial | Same as above |
| **tblBags** | Physical bag/package tracking (lifecycle) | dhanman-inventory | StockTracking / BagMaster | 0% | ❌ Missing | Bag lifecycle tracking not implemented |
| Implicit: Stock Balance | On-hand inventory by item/warehouse | dhanman-inventory | StockBalance / StockLedger | 75% | ⚠️ Partial | Implementation level unclear |
| Implicit: Warehouse Master | Godowns/warehouses (from tblMastNarration.NarrType='G') | dhanman-inventory | Warehouse / Godown | 75% | ⚠️ Partial | May exist in dhanman-app |

**Inventory Tables Summary:** ✅ 0, ⚠️ 4, ❌ 1

---

## Reporting & Admin Tables

| HITRIX Table | Purpose | Dhanman Service | Dhanman Implementation | Confidence | Status |
|---|---|---|---|---|---|
| 51 Stored Procedures | Reports: Trial Balance, Ledger, Sales, Purchase, Stock, GST, etc. | dhanman-finance + dhanman-sales + dhanman-purchase + dhanman-inventory | API endpoints for reports | 50% | ❓ Unknown |
| _Log tables | Audit trail (mirror tables on all transactions) | dhanman-common | AuditLog | 50% | ❓ Unknown |
| Year-end procedures | Financial year closing, consolidation | dhanman-finance | YearEndClosing | 0% | ❌ Missing |

---

## Summary: Coverage by Module

| Module | HITRIX Tables | Dhanman Coverage | Status | Critical Gaps |
|---|---|---|---|---|
| **Masters** | 9 | 8 / 9 (89%) | ✅ Good | Account transaction types, multi-year code ranges |
| **Booking** | 3 | 1-2 / 3 (33%) | ⚠️ Partial | Pre-order workflow, despatch tracking |
| **Sales** | 4 | 2-3 / 4 (50%) | ⚠️ Partial | E-Invoice IRN, interest on late payment, bag tracking |
| **Purchase** | 3 | 1-2 / 3 (33%) | ⚠️ Partial | Interest calculation, TDS deduction |
| **Finance** | 3 | 1 / 3 (33%) | ⚠️ Partial | Interest calculation, bank reconciliation |
| **Inventory** | 5 | 2-3 / 5 (40%) | ⚠️ Partial | Bag lifecycle tracking, godown management |
| **GST & Tax** | Distributed | 2 / 7 (29%) | ❌ Poor | RCM, TCS, TDS, E-Invoice IRN, GSTR2A |
| **Reporting** | 51 SPs | ? | ❓ Unknown | Custom reports, year-end procedures |
| **Audit** | _Log tables | Partial | ⚠️ Partial | Full audit trail |
| **TOTAL** | ~50 tables | ~30 / 50 (60%) | ⚠️ Moderate | **See "Critical for Wave 1" below** |

---

## Critical Gaps for Wave 1 Migration

### 🔴 **MUST BUILD** (Blockers) — Wave 1 Go-Live Requirement

#### 1. **E-Invoice IRN Generation** (tblSale.SlIRNNo)
**Gap Definition:**
- **What's Missing:** IRN (Invoice Registration Number) generation and storage mechanism
- **Dhanman Current State:** No IRN field in Invoice/Order entity; no IRP API integration
- **HITRIX Behavior:** Stores SlIRNNo (IRN), SlAckNo (acknowledgment), SlIRNQRCode (QR code)
- **Business Impact:** GST-registered businesses >50L turnover MUST generate IRN per GST rules; without this, invoices are non-compliant
- **Developer Action Items:**
  1. Add fields to Invoice/Order entity: `irnNumber` (string, unique), `irnQrCode` (string), `irnAckNo` (string), `irnGeneratedAt` (datetime)
  2. Integrate with IRP API (GST Portal) — handle authentication, retry logic, error responses
  3. Auto-generate IRN on invoice finalization (not on creation — must have qty, rate, tax locked)
  4. Display QR code in invoice PDF print
  5. Handle IRN rejection scenarios (invalid GTIN, duplicate reference, etc.)
- **Wave 1 Critical:** YES (GST mandatory)
- **Effort:** 5 days
- **Acceptance Criteria:** IRN successfully generated for test invoice; QR code renders; API calls logged in audit trail

---

#### 2. **Account Master Extensions** (tblMastAccount missing fields)
**Gap Definition:**
- **What's Missing:** 6 critical fields from HITRIX Account master not present in Dhanman
- **Dhanman Current State:** Basic Account/Party/Vendor entities exist but lack:
  - `AcIntPer` (Interest Rate %) — for delayed payment interest calculation
  - `AcOSLimit` (Credit Limit) — maximum outstanding balance allowed
  - `AcMillType` (Transaction Type) — flags for trade, consignment, depot, mill relationship
  - `AcConsignmentFirm` (boolean) — whether this is a consignment supplier
  - `AcDepotFirm` (boolean) — whether this is a depot supplier
  - `AcIsPayToMill` (boolean) — payment routing flag for mill accounts
- **Business Impact:** Without these, purchase/sales workflows cannot enforce credit limits, interest calculations fail, consignment/depot routing breaks
- **Developer Action Items:**
  1. Add 6 fields to Account entity in dhanman-common
  2. Expose in Account edit form (Account detail page)
  3. Validate in purchase/sales: enforce `AcOSLimit` before confirming invoice
  4. Update Account list view to show account type (mill, consignment, depot)
  5. Migrate HITRIX accounts → map tblMastAccount to these new fields
- **Wave 1 Critical:** YES (affects all sales/purchase)
- **Effort:** 3 days
- **Acceptance Criteria:** Account master shows all 6 fields; limit validation blocks invoices over limit; test migrations match source data

---

#### 3. **GST Tax Calculation Engine** (RCM, TCS, correct CGST/SGST/IGST)
**Gap Definition:**
- **What's Missing:** Tax rule engine; RCM (Reverse Charge Mechanism) detection; TCS (Tax Collected at Source) calculation
- **Dhanman Current State:** Basic GST fields exist (gstAmount) but:
  - No RCM flag or auto-detection (unregistered buyer = RCM applies)
  - No TCS calculation (1% collected at source on sales to B2B)
  - CGST/SGST/IGST split may not be correct for all scenarios
  - No tax rule precedence (item tax vs. party tax vs. transaction tax)
- **HITRIX Behavior:** tblMastNarration (T type) defines tax rules; tblSale/tblPurch calculate CGST/SGST/IGST per item and charge
- **Business Impact:** Incorrect tax calculation = GST audit failure; mismatched GSTR liability; payment delays
- **Developer Action Items:**
  1. Build TaxRule entity: tax code, CGST %, SGST %, IGST %, HSN applicability
  2. Implement RCM detection: if (party.gstStatus == 'unregistered') → apply RCM, store `rcmAmount`, create RCM liability GL entry
  3. Implement TCS logic: if (party.isCorporate && transactionType=='B2B') → calculate 1% TCS, store in invoice, create TCS payable GL entry
  4. Tax calculation precedence: item tax > account tax > default tax
  5. Validate CGST + SGST = 18% or IGST = 18% (for 18% items); ensure HSN alignment
  6. Create GL entries automatically: Tax payable (CGST/SGST/IGST separate accounts), RCM recovery, TCS payable
- **Wave 1 Critical:** YES (GST compliance)
- **Effort:** 5-6 days
- **Acceptance Criteria:** Test invoice with unregistered party shows RCM; B2B corporate invoice shows TCS; GL entries created correctly; tax reconciliation matches GSTR 1/2

---

#### 4. **Finance Journal Auto-Creation** (GL entries from sales/purchase)
**Gap Definition:**
- **What's Missing:** Automatic GL journal entry generation when sales/purchase invoices are finalized
- **Dhanman Current State:** JournalEntry entity exists but invoices don't auto-create GL entries
- **HITRIX Behavior:** tblSale/tblPurch → tblVoucher (automatic GL posting on invoice save)
- **Business Impact:** Without automatic GL creation, accountants must manually post invoices to GL → reconciliation delays, errors, duplicate posting risk
- **Developer Action Items:**
  1. On invoice finalize (not on save/draft):
     - Create JournalEntry for Debit: Accounts Receivable (AR) or Accounts Payable (AP)
     - Create JournalEntry for Credit: Sales/COGS or Purchase/Inventory
     - Create JournalEntry for Tax: Tax Payable (CGST/SGST/IGST accounts)
     - Create JournalEntry for Discounts/Charges if applicable
  2. GL account mapping: read from SystemConfiguration or Account.glAccount field
  3. Posting rules:
     - Sales: AR (Debit) vs. Sales Revenue (Credit)
     - Purchase: Inventory/COGS (Debit) vs. AP (Credit)
     - Tax: separate GL lines per tax type (CGST/SGST/IGST)
  4. Error handling: if GL account not found → mark invoice as "GL pending", notify accountant
  5. Audit trail: log user, timestamp, invoice → journal reference
- **Wave 1 Critical:** YES (accounting accuracy)
- **Effort:** 4 days
- **Acceptance Criteria:** Invoice GL journal entries created on finalize; GL accounts match account master config; AR/AP balance = total invoices; GL reconciliation passes

---

#### 5. **Product Price Guardrails** (`MinSellingPrice`, `MaxPurchasePrice`)
**Gap Definition:**
- **What's Missing:** Operational safeguard to prevent pricing errors by end-users
- **Dhanman Current State:** Product entity has base price but no min/max bounds
- **Business Rationale:** HITRIX doesn't explicitly track this, but trading/commodity businesses need price validation to prevent data-entry errors (e.g., selling at cost, buying above market rate)
- **Recommended Behavior:** Define safe price ranges; warn or block invoices outside range
- **Developer Action Items:**
  1. Add to Product entity: `minSellingPrice` (decimal), `maxPurchasePrice` (decimal)
  2. On invoice creation, before line-item save:
     - Check sale item rate >= `minSellingPrice` → warn if lower, block if much lower (20% threshold)
     - Check purchase item rate <= `maxPurchasePrice` → warn if higher, block if much higher
  3. UX: show range in line-item edit modal; allow override with manager approval
  4. Audit: log all overrides for monthly review
- **Wave 1 Critical:** YES (operational safeguard; high error risk)
- **Effort:** 2-3 days
- **Acceptance Criteria:** Price validation triggers on item entry; warnings display; override audit trail created

---

### 🟡 **SHOULD BUILD** (High-Impact) — Recommended for Wave 1 but has workaround

#### 6. **Interest Calculation on Late Payment** (tblIntSale, tblIntDbNote)
**Gap Definition:**
- **What's Missing:** Auto-calculation of interest on outstanding invoices (overdue AR/AP)
- **Dhanman Current State:** No interest/debit-note entity; no scheduled interest calc job
- **HITRIX Behavior:** tblIntSale (sales interest), tblIntDbNote (purchase interest) — daily interest calc, creates debit notes
- **Business Impact:** Overdue receivables/payables not accrued → balance sheet mismatch; customer disputes
- **Developer Action Items:**
  1. Create DebitNote/Interest entity: invoice reference, interest %, days overdue, interest amount, GL account
  2. Scheduled job (daily): scan open invoices; if (daysOverdue > 0) → calculate interest = (amount × rate% × daysOverdue) / 365
  3. Create GL entry: Interest Receivable (Debit) vs. Interest Income (Credit) for sales; Interest Payable vs. Interest Expense for purchases
  4. On invoice payment: close associated interest record
  5. Report: aging analysis including interest accrual
- **Wave 1 Critical:** MAYBE (depends on pilot customer; can use manual workaround for now)
- **Effort:** 5 days
- **Acceptance Criteria:** Interest accrual calculated correctly; GL entries posted; aging report shows interest; GL reconciles with GAAP

---

#### 7. **Bank Reconciliation** (match bank statement to vouchers)
**Gap Definition:**
- **What's Missing:** Bank statement upload and matching tool
- **Dhanman Current State:** No BankReconciliation entity or statement import feature
- **Business Impact:** Accountants cannot easily match bank feeds to GL entries → reconciliation delays
- **Developer Action Items:**
  1. Create BankReconciliation entity: statement line, matching voucher, match status (matched/unmatched/review)
  2. Bank statement import: CSV/Excel upload with bank parser
  3. Auto-match: by amount, date range, GL account
  4. Manual match: UI to drag/drop statement lines to vouchers
  5. Report: reconciled balance vs. GL balance; exceptions
- **Wave 1 Critical:** MEDIUM (can use manual workaround via Excel)
- **Effort:** 5 days
- **Acceptance Criteria:** Statement upload works; auto-match identifies 80%+ of items; reconciliation report matches bank statement; GL balance agrees

---

#### 8. **Booking Pre-Order Workflow** (tblBooking)
**Gap Definition:**
- **What's Missing:** Pre-order/booking order status and workflow (create → despatch → invoice)
- **Dhanman Current State:** Only standard Order exists (create → finalize → deliver); unclear if status-based booking is supported
- **HITRIX Behavior:** tblBooking (booking master), tblBooKingSub (items), tblBooKingDesp (partial despatch)
- **Business Impact:** Commodity traders need booking orders for forward contracts (qty reserved but price/delivery date flexible)
- **Developer Action Items:**
  1. Research: verify if Dhanman Order can have status='booking', or if separate Booking entity needed
  2. If new entity: create Booking master/detail, map fields from tblBooking
  3. Despatch workflow: allow partial shipments from booking (update tblBooKingDesp equivalent)
  4. Invoice from booking: create invoice only for despatched qty
  5. Audit: track booking → despatch → invoice chain
- **Wave 1 Critical:** MEDIUM (only if pilot customer uses bookings; otherwise skip)
- **Effort:** 4 days (research first)
- **Acceptance Criteria:** Booking order created with future delivery date; partial despatch tracked; invoice generated only for shipped qty

---

### 🟢 **CAN DEFER** (Wave 2+) — Non-critical for initial cutover

9. **Bag Tracking** (tblBags lifecycle) — physical inventory control; defer to Wave 2
10. **Consignment/Depot Logistics** — special account types; defer to Wave 2
11. **GSTR2A Portal Matching** — auto-reconciliation with GST portal; defer to Wave 2
12. **TDS Deduction** (Tax Deducted at Source on payments) — unless pilot customer requires; defer to Wave 2
13. **Year-End Financial Closing** — scheduled for end-of-FY1; defer to Wave 2

---

## Wave 1 Go-Live Prerequisites

### MUST-HAVES (Blocking) — Verify in UAT before cutover

| Feature | Acceptance Criteria | Owner | Status |
|---------|---|---|---|
| **E-Invoice IRN** | IRN successfully generated for test invoices; QR code renders in PDF; API errors logged | dhanman-sales | ❓ TODO |
| **Account Master** | All 6 fields populated (AcIntPer, AcOSLimit, AcMillType, etc.); credit limit validation blocks oversales | dhanman-common | ❓ TODO |
| **GST Calculations** | CGST/SGST/IGST split correct for test items; RCM auto-detected for unregistered parties; TCS calculated for B2B; GL entries created | dhanman-finance + dhanman-sales | ❓ TODO |
| **GL Auto-Creation** | Sales/Purchase invoices create JournalEntry on finalize; AR/AP GL balances match invoice totals; GL reconciliation passes | dhanman-finance | ❓ TODO |
| **Masters Data** | Company, Chart of Accounts, Tax codes, Customers/Vendors loaded from HITRIX; validation report shows 0 errors | dhanman-common | ❓ TODO |
| **Opening Balances** | AR/AP/Stock balances reconciled to HITRIX; GL opening balance = HITRIX closing balance for prior FY | dhanman-finance + dhanman-inventory | ❓ TODO |
| **Migration Test** | 100 sample invoices migrated; invoice totals match HITRIX; GL reconciliation passes | dhanman-sales + dhanman-purchase | ❓ TODO |

### NICE-TO-HAVES (Workaround available) — Can be completed post-go-live

| Feature | Workaround | Planned Completion |
|---------|---|---|
| **Interest Calculations** | Manual debit-note creation in first month; auto-calc added in Sprint 2 | 4 weeks post-go-live |
| **Bank Reconciliation** | Manual Excel reconciliation using GL report; Bank Recon feature added in Sprint 3 | 8 weeks post-go-live |
| **Booking Workflow** | Use Order with status field; dedicated Booking entity added if pilot demands it | Evaluate post-go-live |
| **Price Guardrails** | Manual review of invoice prices; validation rules added in Sprint 2 | 4 weeks post-go-live |

---

## Implementation Roadmap (Wave 1)

### Sprint 1 (Days 1-5): Account Master & Price Guardrails

**Task 1.1: Add Account Master Fields** (dhanman-common)
- [ ] Add to Account entity: `interestRate` (decimal), `creditLimit` (decimal), `accountType` (enum: normal/mill/consignment/depot), `isPayToMill` (bool)
- [ ] Database migration: ALTER TABLE accounts ADD COLUMN...
- [ ] API update: GET /accounts/:id include new fields
- [ ] Form update: Account detail page shows 4 new fields
- [ ] Validation: POST /invoices validates account.creditLimit before confirm
- [ ] Test: Create account with limit; attempt invoice exceeding limit → blocked
- [ ] Migration: Write script to backfill HITRIX data

**Task 1.2: Product Price Guardrails** (dhanman-inventory)
- [ ] Add to Product entity: `minSellingPrice`, `maxPurchasePrice` (both decimal, nullable)
- [ ] Database migration: ALTER TABLE products ADD COLUMN...
- [ ] Validation on invoice line-item save: if (type=='sale' && rate < minSellingPrice) → warn; if (rate < 80% of minSellingPrice) → block with override option
- [ ] UI: show price range in line-item modal
- [ ] Audit: log all price overrides to AuditLog
- [ ] Test: attempt sale at cost → warning shown; allow override with reason

---

### Sprint 2 (Days 6-10): GST Tax Engine

**Task 2.1: Build Tax Rule Engine** (dhanman-finance)
- [ ] Create TaxRule entity: code, description, cgstRate, sgstRate, igstRate, hsnApplicable
- [ ] API: GET /tax-rules, POST /tax-rules (admin)
- [ ] Precedence: item tax > account tax > default tax (implement in TaxCalculator service)
- [ ] Test: sale with multiple tax rates → correct CGST+SGST or IGST
- [ ] Migration: Load tblMastNarration (T type) → TaxRule

**Task 2.2: RCM Logic** (dhanman-sales)
- [ ] On invoice save: detect if buyer.gstStatus == 'unregistered' → set rcmApplicable = true
- [ ] RCM calculation: tax amount marked as reverse-chargeable (not seller liability, buyer liability)
- [ ] GL entry: instead of Tax Payable (seller), post RCM Recovery GL entry (buyer responsibility)
- [ ] Invoice display: show "RCM Applicable" badge
- [ ] Test: create sale to unregistered party → RCM flag set; GL entry in recovery account

**Task 2.3: TCS Logic** (dhanman-sales)
- [ ] Identify TCS accounts: if (buyer.isCorporate && buyer.gstStatus == 'registered') → TCS applies
- [ ] TCS calculation: 1% on invoice total (configurable per tax rule)
- [ ] GL entry: TCS Receivable (Debit) vs. TCS Income (Credit)
- [ ] Invoice display: show "TCS 1% = INR XXX"
- [ ] Test: B2B corporate sale → TCS calculated; GL posted to TCS account

**Task 2.4: Tax Calculation Audit** (dhanman-sales + dhanman-purchase)
- [ ] Create TaxAuditLog: invoice → calculated taxes → GL entries (trail for audit)
- [ ] Report: Tax reconciliation (total CGST/SGST posted vs. GSTR 1 liability)
- [ ] Test: 10-invoice batch → tax audit report shows 0 variance

---

### Sprint 3 (Days 11-15): E-Invoice & GL Auto-Creation

**Task 3.1: IRN Field & IRP Integration** (dhanman-sales)
- [ ] Add to Invoice: `irnNumber` (string, unique, indexed), `irnQrCode` (base64), `irnAckNumber`, `irnGeneratedAt` (datetime)
- [ ] Database: new fields + unique constraint on irnNumber
- [ ] IRP API client: integrate with GST IRN API (https://api.invoiceregistration.xyz/...)
  - [ ] Handle auth (HMAC token)
  - [ ] Retry logic (3x with exponential backoff)
  - [ ] Error handling: invalid GTIN, duplicate reference, unregistered seller → log and notify
- [ ] Invoice finalization: on status='finalized', call IRP API → get IRN/QR → store in DB
- [ ] UI: display IRN in invoice PDF; show QR code
- [ ] Fallback: if IRP down, mark invoice as 'pending_irn', add to queue for batch processing
- [ ] Test: finalize invoice → IRN generated; QR renders; invoice PDF displays

**Task 3.2: GL Auto-Creation on Invoice** (dhanman-finance)
- [ ] Create JournalEntry on invoice.finalize():
  - **For Sales:** 
    - Debit: Accounts Receivable = invoice total
    - Credit: Sales Revenue = base amount
    - Credit: Tax Payable (CGST) = cgst amount
    - Credit: Tax Payable (SGST) = sgst amount
    - (or IGST if applicable)
  - **For Purchase:**
    - Debit: Inventory/COGS = base amount
    - Debit: Tax Input (CGST) = cgst amount
    - Debit: Tax Input (SGST) = sgst amount
    - Credit: Accounts Payable = invoice total
- [ ] GL account mapping: read from Account.glAccount or SystemConfiguration.defaultGLAccounts
- [ ] Posting: all entries posted together in single JournalEntry batch (for reconciliation)
- [ ] Error handling: if GL account not found → mark invoice 'GL_PENDING', notify accountant
- [ ] Audit: JournalEntry.referenceType='Invoice', referenceId=invoice.id
- [ ] Test: create+finalize sale → JournalEntry created; AR balance increases; GL report shows entry

**Task 3.3: GL Balance Reconciliation Report** (dhanman-finance)
- [ ] Query: SUM(JournalEntry) for AR account = SUM(Invoice.total) for sales
- [ ] Query: SUM(JournalEntry) for Inventory = SUM(Purchase invoice.amount)
- [ ] Report endpoint: GET /reports/gl-reconciliation?startDate&endDate → show variance (should be 0)
- [ ] UI: add "GL Reconciliation" dashboard widget
- [ ] Test: post 50 invoices → reconciliation report variance = 0

---

### Sprint 4 (Days 16-20): Data Migration & UAT Setup

**Task 4.1: Master Data Migration** (all services)
- [ ] Script: HITRIX → Dhanman
  - [ ] tblMastCompany → Company
  - [ ] tblMastAccount → Account + backfill credit limits
  - [ ] tblMastItem → Product + backfill price bounds
  - [ ] tblMastNarration → TaxRule / Godown / Carrier
  - [ ] tblMastGroup/SubGroup → AccountGroup/SubGroup
- [ ] Validation: row counts match (with mapping doc), amounts reconcile, no nulls in required fields
- [ ] Rollback: all migration scripts must support rollback to pre-migration state

**Task 4.2: Opening Balances Migration** (dhanman-finance)
- [ ] Script: Load HITRIX closing balances (by GL account, by month, by customer)
- [ ] Create Opening Balance JournalEntry (one per GL account)
- [ ] Validate: GL balance after migration = HITRIX FY closing balance
- [ ] Report: Opening Balance Reconciliation (GL vs. HITRIX)

**Task 4.3: UAT Test Plan Setup**
- [ ] Define test cases for all 5 critical features (see UAT Checklist below)
- [ ] Prepare UAT environment with migrated data
- [ ] Create test data: 10 invoices covering RCM, TCS, multi-tax, interest scenarios

**Task 4.4: Documentation**
- [ ] Update deployment runbook with new schema changes
- [ ] Update API docs with new Invoice fields (irnNumber, etc.)
- [ ] Create troubleshooting guide (IRP failures, GL posting errors)

---

## UAT Checklist (Before Go-Live)

**Tester: _____ | Date: _____**

### E-Invoice IRN
- [ ] Create sale invoice, finalize → IRN generated within 30 seconds
- [ ] IRN is unique; second invoice gets different IRN
- [ ] QR code renders in PDF
- [ ] IRP API failure → invoice marked 'pending_irn'; retry works
- [ ] Unregistered seller attempting IRN → error message shown

### Account Master & Credit Limits
- [ ] Create account with credit limit 100K
- [ ] Create sale invoice for 110K to same account → blocked with error "exceeds credit limit"
- [ ] Account shows account type (mill/consignment/etc.)
- [ ] HITRIX account data migrated correctly (credit limits, types match)

### GST Tax Calculations
- [ ] Sale to registered party: CGST/SGST split correctly (no IGST)
- [ ] Sale to unregistered party: RCM flag set; tax marked as buyer liability
- [ ] Sale to corporate: TCS 1% calculated
- [ ] Purchase invoice: tax input posting to Input account (not expense)
- [ ] Tax audit report: variance = 0 for test batch

### GL Auto-Creation
- [ ] Sale invoice posted → AR and Sales Revenue GL entries created
- [ ] Purchase invoice posted → Inventory and AP GL entries created
- [ ] GL reconciliation report: AR total = all open sales invoices; AP total = all open purchase invoices
- [ ] Multi-currency or multi-tax scenario: all GL entries present, none missed

### Price Guardrails
- [ ] Sale at 80% below min price → blocked
- [ ] Sale at 90% below min price → warning shown; user can override with reason
- [ ] Override logged to audit trail
- [ ] Purchase above max price → warning; override tracked

### Data Migration
- [ ] 100 sample invoices migrated; totals match HITRIX exactly
- [ ] All customer/vendor records present; no blanks
- [ ] Opening balances reconciled (GL = HITRIX closing)
- [ ] No duplicate records

---

## Post-Go-Live Rollout (Days 21-28)

**Day 1 (Cutover Day)**
- [ ] Backup HITRIX database
- [ ] Run final migration scripts
- [ ] Activate Dhanman for production customers
- [ ] Monitor IRP API calls, GL posting, error logs (on-call engineer assigned)

**Day 2-7 (First Week)**
- [ ] Monitor daily: IRP failures, GL reconciliation variance, user-reported issues
- [ ] Fix urgent bugs (blocked transactions, API failures)
- [ ] Document all issues in Post-Go-Live Issues Log

**Week 2-4 (First Month)**
- [ ] Backfill: Interest calculations for overdue invoices (manual workaround data)
- [ ] Prepare Sprint 2 items (bank reconciliation, interest automation)
- [ ] Collect user feedback; schedule enhancement planning

---

**Next Step:** Confirm critical gaps with product owner → Assign Sprint 1 tasks → Kick off development
