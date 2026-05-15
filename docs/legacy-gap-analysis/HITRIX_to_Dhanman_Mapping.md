# HITRIX V2 ↔ Dhanman Comparison Table

**Purpose:** Map HITRIX V2 legacy tables to Dhanman microservices, identify coverage and gaps.

**Updated:** 2026-05-15 | **Deadline:** 2026-07-15 (8 weeks)

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
| **tblMastItem** | Products/commodities (name, unit, HSN, GST, brokerage rate) | dhanman-inventory | Product | 100% | ✅ Present | None identified |
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

### 🔴 **MUST BUILD** (Blockers)

1. **E-Invoice IRN Generation** (tblSale.SlIRNNo)
   - Dunman lacks: IRN field, IRN generation API integration
   - Wave 1 Critical: YES (GST mandatory above 50 lakh)
   - Effort: 5 days
   
2. **Interest Calculation** (tblIntSale, tblIntDbNote)
   - Dunman lacks: Debit note entity, auto-interest logic
   - Wave 1 Critical: MAYBE (depends on pilot customer)
   - Effort: 5 days

3. **Account Extensions** (tblMastAccount extra fields)
   - Dunman needs: AcIntPer, AcOSLimit, AcMillType, consignment/depot flags
   - Wave 1 Critical: YES (affects all sales/purchase)
   - Effort: 3 days

4. **GST Tax Engine** (RCM, TCS, correct CGST/SGST/IGST)
   - Dunman lacks: RCM detection, TCS calculation, tax rule engine
   - Wave 1 Critical: YES (GST compliance)
   - Effort: 5 days

### 🟡 **SHOULD BUILD** (High-Impact)

5. **Finance Journal Auto-Creation** (GL entries from sales/purchase)
   - Wave 1 Critical: YES (accounting accuracy)
   - Effort: 4 days

6. **Bank Reconciliation** (match statement to vouchers)
   - Wave 1 Critical: MEDIUM (can use manual workaround)
   - Effort: 5 days

7. **Booking Pre-Order** (tblBooking workflow)
   - Wave 1 Critical: MEDIUM (only if pilot uses bookings)
   - Effort: 4 days (research first)

### 🟢 **CAN DEFER** (Wave 2+)

8. **Bag Tracking** (tblBags lifecycle) — defer to Wave 2
9. **Consignment/Depot** — defer to Wave 2
10. **GSTR2A Matching** — defer to Wave 2
11. **TDS Deduction** — defer to Wave 2 (unless pilot needs)
12. **Year-End Closing** — defer to Wave 2

---

## Wave 1 Go-Live Prerequisites

**Before cutover, MUST have:**
- ✅ E-Invoice IRN generation working
- ✅ Tax calculations (CGST/SGST/IGST) correct within 0.01%
- ✅ Account master enhancements deployed
- ✅ Finance journal entries auto-created
- ✅ Masters loaded and validated
- ✅ Opening balances (AR/AP/stock) reconciled
- ✅ Sample data migration tested

**Can operate without:**
- Interest calculations (manual workaround)
- Bank reconciliation (manual workaround)
- Booking pre-order (use Order with status)
- Bag tracking (optional for Wave 1)

---

**Next Step:** Confirm critical gaps → Estimate Phase 2 effort → Start development sprints
