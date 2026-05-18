# Sales, Purchase, Inventory & Broker Relationship Flow

**Purpose:** Map end-to-end business flows for commodity trading; identify where gaps impact operations.

**Document Date:** 2026-05-16 | **Related Document:** HITRIX_to_Dhanman_Mapping.md

---

## 1. SALES FLOW (Sales Order → Delivery → Invoice)

### 1.1 Standard Sales Flow (Direct Sale to Customer)

```
┌─────────────────────────────────────────────────────────────────┐
│                    SALES ORDER CREATION                         │
├─────────────────────────────────────────────────────────────────┤
│ INPUT:                                                           │
│  • Customer (Account) + Credit Limit validation                 │
│  • Product (Item) + Quantity + Unit (MT, bags, etc.)            │
│  • Rate per unit + Price validation (Min Selling Price)         │
│  • Godown (Warehouse) source                                     │
│  • GST calculation (CGST/SGST or IGST)                          │
│  • Broker details (if applicable)                               │
│                                                                  │
│ HITRIX Tables: tblSale, tblSaleSub                             │
│ Dhanman Entities: Order, OrderItem (⚠️ Partial coverage)        │
└─────────────────────────────────────────────────────────────────┘
                              ↓
┌─────────────────────────────────────────────────────────────────┐
│                 INVENTORY ALLOCATION                            │
├─────────────────────────────────────────────────────────────────┤
│ ACTION:                                                          │
│  • Check stock balance in selected godown                       │
│  • Reserve stock (qty allocated to sale order)                  │
│  • Update stock ledger: (Beginning Qty - Reserved Qty)          │
│  • If stock < required qty → backorder or rejection             │
│                                                                  │
│ HITRIX: Stock tracked via tblBags (physical bag IDs)           │
│         Linked by tblSaleSub.SlSubItCtrlNo → tblBags           │
│ Dhanman: StockBalance / StockLedger (⚠️ Missing bag tracking)  │
│ GAP: Bag lifecycle tracking NOT in Dhanman                     │
└─────────────────────────────────────────────────────────────────┘
                              ↓
┌─────────────────────────────────────────────────────────────────┐
│                    GATE PASS / DISPATCH                         │
├─────────────────────────────────────────────────────────────────┤
│ ACTION:                                                          │
│  • Create Gate Pass (Goods Issue document) — physical shipment  │
│  • Link to Sale Order                                           │
│  • Record actual qty dispatched (may be partial)                │
│  • Bag IDs / Package IDs recorded (if tracking enabled)         │
│  • Update inventory: actual deduction from godown               │
│                                                                  │
│ HITRIX Tables: tblGatePass, tblGatePassSub                     │
│ Dhanman Entity: GoodsIssue / Delivery (⚠️ Partial implementation)│
│ GAP: Godown/warehouse tracking level unclear                   │
└─────────────────────────────────────────────────────────────────┘
                              ↓
┌─────────────────────────────────────────────────────────────────┐
│                  INVOICE GENERATION                             │
├─────────────────────────────────────────────────────────────────┤
│ ACTION (on order finalization):                                 │
│  1. Validate: All despatched qty matches bill qty               │
│  2. Calculate GST: CGST + SGST or IGST (per tax rules)         │
│  3. Check GST Rules: RCM applicable? TCS applicable?            │
│  4. Create Invoice (tblSale finalization)                       │
│  5. Generate E-Invoice: Get IRN from GST portal                │
│  6. Auto-create GL entries: AR debit, Sales revenue credit     │
│  7. Link to broker: record broker commission (if applicable)    │
│                                                                  │
│ HITRIX Tables: tblSale → tblVoucher (auto GL posting)          │
│ Dhanman Issues:                                                 │
│  ❌ GAP 1: IRN generation missing (E-Invoice not generated)    │
│  ❌ GAP 2: RCM/TCS tax logic missing (wrong tax calculation)   │
│  ❌ GAP 3: GL auto-posting missing (manual journal entry)      │
│  ❌ GAP 4: Broker commission handling unclear                  │
└─────────────────────────────────────────────────────────────────┘
                              ↓
┌─────────────────────────────────────────────────────────────────┐
│                   PAYMENT COLLECTION                            │
├─────────────────────────────────────────────────────────────────┤
│ ACTION:                                                          │
│  • Receive customer payment (check, bank transfer, etc.)        │
│  • Create Bank Receipt Voucher (BR type) or Cash Receipt (CR)  │
│  • Match to invoice: update AR to 0 (payment received)         │
│  • If late payment: accrue interest (tblIntSale)               │
│  • Update GL: Cash/Bank debit, AR credit                        │
│                                                                  │
│ HITRIX Tables: tblVoucher (BR/CR types)                        │
│ Dhanman Entity: Payment / Receipt (⚠️ Partial)                  │
│ GAP: Interest accrual missing (Wave 2 feature)                  │
└─────────────────────────────────────────────────────────────────┘
```

**Sales Flow Summary:**
- ✅ **WORKS:** Order creation, basic invoice
- ⚠️ **PARTIAL:** Inventory reservation (bag tracking missing), GoodsIssue implementation unclear
- ❌ **MISSING:** IRN generation, RCM/TCS tax engine, GL auto-posting, broker commission routing

---

## 2. PURCHASE FLOW (Purchase Order → Goods Receipt → Invoice)

### 2.1 Standard Purchase Flow (Direct Purchase from Supplier)

```
┌─────────────────────────────────────────────────────────────────┐
│                 PURCHASE ORDER CREATION                         │
├─────────────────────────────────────────────────────────────────┤
│ INPUT:                                                           │
│  • Supplier (Account) + Credit Limit (AcOSLimit)               │
│  • Product (Item) + Quantity                                    │
│  • Rate per unit + Purchase validation (Max Purchase Price)     │
│  • Destination Godown (Warehouse)                               │
│  • GST + TDS (Tax Deducted at Source) if applicable            │
│  • Broker details (if commission applicable)                    │
│                                                                  │
│ HITRIX Tables: tblPurch, tblPurchSub                           │
│ Dhanman Entities: PurchaseOrder, PurchaseOrderItem             │
│                   (⚠️ TDS fields missing)                        │
└─────────────────────────────────────────────────────────────────┘
                              ↓
┌─────────────────────────────────────────────────────────────────┐
│                  GOODS RECEIPT (GRN)                            │
├─────────────────────────────────────────────────────────────────┤
│ ACTION:                                                          │
│  • Receive physical goods from supplier                         │
│  • Verify qty, quality, GST Invoice from supplier              │
│  • Create GRN (Goods Received Note) — match to PO              │
│  • Record into godown (warehouse location)                      │
│  • Update inventory: add to stock balance                       │
│  • If inspection required: quarantine pending QC check         │
│                                                                  │
│ HITRIX: No explicit GRN table; implicit in gate pass inbound   │
│ Dhanman Entity: GoodsReceipt / GoodsIssue (⚠️ Unclear)         │
│ GAP: GRN workflow not clearly defined in Dhanman               │
└─────────────────────────────────────────────────────────────────┘
                              ↓
┌─────────────────────────────────────────────────────────────────┐
│              PURCHASE INVOICE MATCHING (3-Way Match)            │
├─────────────────────────────────────────────────────────────────┤
│ ACTION (3-way reconciliation):                                  │
│  • PO: ordered qty & rate                                       │
│  • GRN: received qty                                            │
│  • Supplier Invoice: billed qty & rate (may differ!)            │
│  • Match: PO qty ≈ GRN qty ≈ Invoice qty (tolerance ±2%)      │
│  • Calculate variance: (Invoice Qty - GRN Qty) × Rate = variance│
│  • If variance > tolerance → flag for approval                  │
│                                                                  │
│ HITRIX: Manual matching via tblPurch receipt date tracking     │
│ Dhanman: GRN matching feature (status unclear)                 │
│ GAP: 3-way match workflow not clearly documented               │
└─────────────────────────────────────────────────────────────────┘
                              ↓
┌─────────────────────────────────────────────────────────────────┐
│                 PURCHASE INVOICE POSTING                        │
├─────────────────────────────────────────────────────────────────┤
│ ACTION (on invoice finalization):                               │
│  1. Calculate GST Input: Tax INPUT (claimable from GST)        │
│  2. Calculate TDS: Tax Deducted at Source (if supplier TDS%)   │
│  3. RCM Check: if supplier unregistered → RCM applies          │
│  4. Create Invoice (tblPurch finalization)                      │
│  5. Auto-create GL entries: Inventory/COGS debit, AP credit    │
│  6. Create TDS Payable entry (if TDS applicable)               │
│  7. Link to broker: record purchase broker commission           │
│                                                                  │
│ HITRIX Tables: tblPurch → tblVoucher (auto GL posting)         │
│ Dhanman Issues:                                                 │
│  ❌ GAP 1: TDS deduction missing (TDS fields in Account)      │
│  ❌ GAP 2: RCM detection missing (unregistered supplier check) │
│  ❌ GAP 3: GL auto-posting missing (manual entry required)     │
│  ❌ GAP 4: Broker commission handling unclear                  │
│  ⚠️ GAP 5: Tax Input vs. Expense distinction unclear           │
└─────────────────────────────────────────────────────────────────┘
                              ↓
┌─────────────────────────────────────────────────────────────────┐
│                  PAYMENT TO SUPPLIER                            │
├─────────────────────────────────────────────────────────────────┤
│ ACTION:                                                          │
│  • Process payment: Check, Bank Transfer, or Demand Draft      │
│  • Create Bank Payment Voucher (BP) or Check Payment (CP)      │
│  • Deduct TDS at payment (if TDS applicable):                  │
│    TDS = Invoice Amount × TDS% (e.g., 2% for contractors)     │
│    Net Payment = Invoice - TDS                                 │
│  • Match to invoice: update AP to 0 (payment made)             │
│  • Update GL: Bank/Cash credit, AP debit, TDS receivable       │
│  • If interest accrued on late payment: deduct from payment    │
│                                                                  │
│ HITRIX Tables: tblVoucher (BP/CP types), tblIntDbNote          │
│ Dhanman Entity: Payment / PaymentBatch (⚠️ Partial)            │
│ GAP: TDS at payment + Interest accrual missing                 │
└─────────────────────────────────────────────────────────────────┘
```

**Purchase Flow Summary:**
- ✅ **WORKS:** PO creation, basic invoice
- ⚠️ **PARTIAL:** GRN matching (workflow unclear), 3-way match (not documented)
- ❌ **MISSING:** TDS deduction, RCM detection, GL auto-posting, broker commission routing, interest accrual

---

## 3. BROKER RELATIONSHIP & COMMISSION FLOW

### 3.1 What is a Broker in Commodity Trading?

A **broker** is a middleman who facilitates buy/sell transactions between parties. The broker:
- **Matches buyers & sellers** (they don't hold inventory)
- **Charges commission** (% of transaction value or fixed amount)
- **May hold margin** (security deposit until settlement)
- **Is account type:** "Broker" in HITRIX, recorded in tblMastAccount.AcBrokerType

**Broker Flow:**

```
┌────────────────────────────────────────────────────────────┐
│              BROKER COMMISSION STRUCTURE                   │
├────────────────────────────────────────────────────────────┤
│                                                             │
│  SALE via Broker:                                          │
│  ─────────────────                                         │
│  Customer Sale Price:  INR 100 per unit                    │
│  Broker Commission:    1% on transaction value             │
│  Commission Amount:    INR 1 per unit                      │
│                                                             │
│  When recorded:                                            │
│  • Broker charges commission on invoice finalization       │
│  • Recorded in tblMastNarration (C type = Charge)         │
│  • GL Entry: Commission Expense (Debit), AP to Broker      │
│                                                             │
│  PURCHASE via Broker:                                      │
│  ────────────────────                                      │
│  Supplier Rate:       INR 90 per unit                      │
│  Broker Commission:   1% on transaction value              │
│  Commission Amount:   INR 0.90 per unit                    │
│                                                             │
│  When recorded:                                            │
│  • Broker charges commission on purchase invoice           │
│  • Added to Cost of Goods (COGS)                           │
│  • GL Entry: COGS/Inventory (Debit), AP to Broker          │
└────────────────────────────────────────────────────────────┘
```

### 3.2 Broker Commission in Sales (Detailed Flow)

```
STEP 1: Customer places order via broker
        ├─ Broker Account in dhanman-common
        └─ Invoice links to Broker field

STEP 2: On sales invoice finalization
        ├─ Calculate commission: (Sale Total) × (Broker Commission %)
        └─ Example: INR 1,00,000 sale × 1% = INR 1,000 broker commission

STEP 3: GL Entry for broker commission
        ├─ Debit: Commission Expense (GL Account: 6200)
        ├─ Credit: Accounts Payable — Broker (GL Account: 2100)
        └─ Description: "Commission on Sale Invoice #SL001 to Customer XYZ"

STEP 4: Broker Invoice (on-demand)
        ├─ Accumulate all commissions for a broker per month
        ├─ Create single payment voucher (summary invoice to broker)
        └─ Reconcile: Sum of invoices = Broker statement

STEP 5: Payment to Broker
        ├─ Create Bank Payment Voucher (BP type)
        ├─ Match to broker invoices
        └─ Update AP to 0

GAP IN DHANMAN:
❌ No explicit broker commission tracking in Invoice
❌ No automated GL entry for commission
❌ No broker-wise commission summary report
⚠️  Need: Invoice.brokerCommissionAmount, Invoice.brokerAccount
```

### 3.3 Special Broker Cases (Consignment & Depot Models)

```
CONSIGNMENT SALE (Broker holds inventory):
─────────────────────────────────────────
Scenario: Broker stores goods at warehouse, sells on our behalf

HITRIX Fields: tblMastAccount.AcConsignmentFirm = true
              tblSale.SlMillConsignFirm = broker account

Flow:
  ├─ We send goods to broker warehouse (No Sale yet, inventory remains ours)
  ├─ Broker receives via GatePass (outbound to consignment warehouse)
  ├─ Broker sells from warehouse (creates invoice with our consent)
  ├─ We invoice customer (tblSale for consignment lot)
  ├─ Broker deducts commission + expenses
  ├─ Net settlement: Goods value - Commission - Expenses = our revenue
  └─ GL: Consignment inventory tracking separate from normal stock

Dhanman Gap:
❌ Consignment flag missing from Account (AcConsignmentFirm)
❌ Consignment warehouse management not clear
❌ Net settlement invoice (broker's charge note) not handled
→ DEFER to Wave 2


DEPOT MODEL (Broker sells from depot):
──────────────────────────────────────
Scenario: Broker operates as agent for inventory stored at depots

HITRIX Fields: tblMastAccount.AcDepotFirm = true
              tblSale.SlMillDepot = depot location

Flow:
  ├─ Goods stored at depot (could be broker's depot or ours)
  ├─ Broker facilitates sales from depot
  ├─ Each sale creates invoice (customer buys from depot)
  ├─ Broker commission applied (% of sale price)
  ├─ Depot charges storage (if not broker's depot)
  └─ Net settlement on monthly/quarterly basis

Dhanman Gap:
❌ Depot flag missing from Account (AcDepotFirm)
❌ Multi-location inventory management not clear
❌ Depot charges (storage, handling) not tracked separately
→ DEFER to Wave 2
```

---

## 4. IMPACT OF GAPS ON SALES/PURCHASE/INVENTORY FLOWS

### 4.1 Gap Impact Matrix

| Gap | Sales Flow | Purchase Flow | Inventory | Broker Flow | Criticality |
|-----|---|---|---|---|---|
| **E-Invoice IRN Missing** | ❌ Cannot generate GST-compliant invoices | ❌ Same | N/A | ❌ Affects all taxable invoices | 🔴 CRITICAL |
| **Account Extensions Missing** (AcOSLimit, AcIntPer, AcMillType) | ❌ Credit limit not enforced | ❌ Supplier limit not enforced | N/A | ❌ Consignment/Depot flags missing | 🔴 CRITICAL |
| **RCM/TCS Tax Logic Missing** | ❌ Wrong tax calc for unregistered parties / corporates | ❌ Same | N/A | ❌ Affects commission tax treatment | 🔴 CRITICAL |
| **GL Auto-Posting Missing** | ❌ Manual GL entry required | ❌ Manual GL entry required | ❌ Inventory GL not auto-posted | ❌ Commission GL manual | 🔴 CRITICAL |
| **Broker Commission Handling** | ❌ No commission tracking | ⚠️ Purchase broker commission missing | N/A | ❌ Cannot calculate net settlement | 🟡 HIGH |
| **Bag Tracking Missing** | ⚠️ Cannot track individual bags shipped | ⚠️ Cannot track inbound bags | ❌ Physical bag lifecycle lost | ⚠️ Batch tracking only | 🟢 DEFER (Wave 2) |
| **GoodsIssue/GRN Unclear** | ⚠️ Dispatch documentation level unclear | ⚠️ Receipt workflow unclear | ⚠️ Godown management unclear | N/A | 🟡 MEDIUM (clarify first) |
| **Interest Accrual Missing** | ⚠️ Late payment interest not accrued | ⚠️ Same | N/A | N/A | 🟡 MEDIUM (Wave 2 workaround) |

---

## 5. FIXES REQUIRED PER FLOW

### 5.1 Sales Flow Fixes (Wave 1)

**Fix 1: Account Credit Limit Enforcement**
```
BEFORE (HITRIX):
  Invoice total = INR 150,000
  Customer credit limit = INR 100,000
  System: ✅ Allowed (enforced at tblMastAccount level)

AFTER (Dhanman — needs fix):
  Same scenario
  System: ❌ No limit check (AcOSLimit field missing)
  
FIX:
  1. Add Account.creditLimit field
  2. On invoice.save(): if (invoiceTotal > account.creditLimit) → BLOCK
  3. Show error: "Customer credit limit exceeded: INR 100K. Invoice total: INR 150K"
```

**Fix 2: Min Selling Price Validation**
```
BEFORE: User enters sale rate = INR 50 (product min price = INR 100)
AFTER (needs fix):
  System: ❌ Allows entry
  
FIX:
  1. Add Product.minSellingPrice field
  2. On invoice line-item.save(): if (rate < minSellingPrice) → WARN; if (rate < 80% minPrice) → BLOCK
  3. Show warning: "Sale rate (INR 50) below minimum (INR 100). Override requires manager approval"
```

**Fix 3: E-Invoice IRN Generation**
```
BEFORE: Manual IRN generation (copy-paste from IRP portal)
AFTER (needs fix):
  System: ❌ No IRN integration
  
FIX:
  1. Add Invoice fields: irnNumber, irnQrCode, irnAckNumber
  2. On invoice.finalize(): Call IRP API (GST portal)
  3. Auto-receive IRN within 30 seconds
  4. Store QR code; display in PDF
```

**Fix 4: RCM/TCS Tax Calculation**
```
BEFORE: Complex tax logic in HITRIX (tax rule engine)
AFTER (needs fix):
  System: ❌ Basic CGST/SGST only; no RCM or TCS
  
FIX:
  1. Build TaxRule engine with precedence
  2. RCM detection: if (customer.gstStatus == 'unregistered') → rcmApplicable = true
  3. TCS detection: if (customer.isCorporate && gstStatus == 'registered') → tcsApplicable = true
  4. Correct GL posting: RCM liability on buyer; TCS payable on seller
```

**Fix 5: GL Auto-Posting**
```
BEFORE: Manual GL journal entry creation
AFTER (needs fix):
  System: ❌ No auto GL posting from invoice
  
FIX:
  1. On invoice.finalize():
     ├─ Debit: AR (Accounts Receivable) = invoice total
     ├─ Credit: Sales Revenue = base amount
     ├─ Credit: Tax Payable (CGST, SGST, or IGST) = tax amounts
     └─ Create single JournalEntry batch
  2. Link JournalEntry.referenceId = invoice.id
  3. Audit: Track all GL entries with invoice reference
```

**Fix 6: Broker Commission Tracking**
```
BEFORE: Broker commission recorded as tblMastNarration (charge type)
AFTER (needs fix):
  System: ⚠️ No explicit broker field on invoice
  
FIX:
  1. Add Invoice.brokerAccount field (link to broker Account)
  2. Add Invoice.brokerCommissionPercent & brokerCommissionAmount
  3. On invoice.finalize():
     ├─ Calculate commission = total × brokerCommission%
     └─ Create GL entry: Commission Expense (Debit), AP—Broker (Credit)
  4. Broker settlement report: monthly commission summary
```

### 5.2 Purchase Flow Fixes (Wave 1)

**Fix 1: GRN Workflow Clarity**
```
ISSUE: GRN (Goods Received Note) workflow unclear in Dhanman
  
FIX:
  1. Clarify: Is GRN a separate entity or part of PurchaseOrder?
  2. Dhanman should have GoodsReceipt entity with:
     ├─ Linked to PurchaseOrder
     ├─ Record actual received qty (may differ from PO)
     └─ Update inventory: add to godown stock
  3. 3-way match: PO qty ≈ GRN qty ≈ Invoice qty
```

**Fix 2: TDS Deduction at Source**
```
BEFORE: tblMastAccount.AcTDSPer tracks TDS % per supplier
AFTER (needs fix):
  System: ❌ No TDS field in Account; no TDS calculation
  
FIX:
  1. Add Account.tdsPercentage field
  2. On invoice.finalize(): Calculate TDS = invoiceAmount × account.tdsPercentage
  3. GL entry: TDS Receivable (Debit), AP (Credit for net amount)
  4. On payment: Deduct TDS from payment (net = invoice - TDS)
  5. Create TDS payment voucher (quarterly to tax authority)
```

**Fix 3: RCM on Unregistered Supplier**
```
BEFORE: Manually identify unregistered suppliers; apply RCM
AFTER (needs fix):
  System: ❌ No RCM detection
  
FIX:
  1. Check Account.gstStatus == 'unregistered' on invoice creation
  2. If true: Set rcmApplicable = true
  3. Tax liability shifts to buyer (our company)
  4. GL entry: Tax Input (Debit), RCM Payable (Credit)
```

**Fix 4: GL Auto-Posting for Purchases**
```
BEFORE: Manual GL entry creation
AFTER (needs fix):
  System: ❌ No auto GL posting
  
FIX:
  1. On invoice.finalize():
     ├─ Debit: Inventory/COGS = base amount
     ├─ Debit: Tax Input (CGST, SGST, or IGST) = tax amounts
     ├─ Credit: AP (Accounts Payable) = total
     └─ If TDS applicable:
        ├─ Debit: TDS Receivable = tds amount
        └─ Credit: AP (reduce by TDS)
  2. Link JournalEntry to invoice
```

**Fix 5: Broker Commission on Purchases**
```
BEFORE: Broker commission included in invoice or charged separately
AFTER (needs fix):
  System: ⚠️ Unclear handling; may not be separate GL entry
  
FIX:
  1. Add PurchaseOrder.brokerAccount field
  2. Calculate commission = base amount × brokerCommission%
  3. Add to COGS (not separate line item)
  4. GL entry: COGS (Debit), AP—Broker (Credit)
  5. Broker settlement: accumulate and settle monthly
```

### 5.3 Inventory Flow Fixes (Wave 1)

**Fix 1: Stock Reservation on Sales**
```
ISSUE: Stock reserved when sale order created; deducted on GatePass
  
FIX:
  1. On OrderItem.add(): Reserve qty from StockBalance
  2. StockBalance state: (OnHand - Reserved)
  3. On GatePass.create(): Deduct from OnHand (confirm shipment)
  4. Audit: Track reserved → shipped progression
```

**Fix 2: Stock Receipts on Purchase**
```
ISSUE: Stock added on GRN creation (goods received), not invoice
  
FIX:
  1. On GoodsReceipt.create(): Add qty to StockBalance (onhand)
  2. Invoice matching: verify GRN qty ≈ Invoice qty
  3. If variance: flag for approval (invoice qty > GRN qty = over-billing)
  4. Only update GL when invoice matched to GRN
```

**Fix 3: Godown / Warehouse Management**
```
ISSUE: Godowns not clearly mapped; multi-warehouse inventory unclear
  
FIX:
  1. Define Godown/Warehouse entity with:
     ├─ Godown name & location
     ├─ Stock balances per item per godown
     └─ Capacity (if applicable)
  2. Sale order specifies source godown
  3. Purchase order specifies destination godown
  4. GatePass routes between godowns (internal transfers)
  5. Stock report: by-godown view
```

---

## 6. BROKER FLOW FIXES (WAVE 1 & WAVE 2)

### 6.1 Wave 1: Minimum Broker Support

**Fix 1: Account Type — Broker Flag**
```
BEFORE: tblMastAccount has flags for different account types
AFTER (needs fix):
  System: ❌ No explicit broker field in Dhanman Account
  
FIX:
  1. Add Account.accountType = enum: 'CUSTOMER' | 'SUPPLIER' | 'BROKER' | 'EMPLOYEE'
  2. Or add Account.isBroker = boolean
  3. Use for filtering: "List all broker accounts"
```

**Fix 2: Broker Commission on Invoices**
```
BEFORE: Commission tracked per transaction (HITRIX)
AFTER (needs fix):
  System: ❌ No broker commission field on invoice
  
FIX:
  1. Add Invoice.brokerAccount field (link to broker Account)
  2. Add Invoice.brokerCommissionPercent (0-10%)
  3. Add Invoice.brokerCommissionAmount (auto-calculated)
  4. GL entry: Commission Expense (for sales) or COGS addition (for purchases)
```

**Fix 3: Broker Commission GL Posting**
```
BEFORE: Manual GL entry creation
AFTER (needs fix):
  System: ❌ No auto GL for broker commission
  
FIX:
  1. On invoice.finalize() (if broker present):
     ├─ Calculate: commissionAmount = (invoiceTotal - taxes) × brokerCommission%
     ├─ If SALES:
     │  ├─ Debit: Commission Expense (6200)
     │  └─ Credit: AP—Broker (2100)
     ├─ If PURCHASE:
     │  ├─ Debit: COGS/Inventory (add to cost, not separate)
     │  └─ Credit: AP—Broker (2100)
     └─ Create JournalEntry with referenceId = invoice.id
```

**Fix 4: Broker Settlement Report**
```
BEFORE: Manual broker settlement statement
AFTER (needs fix):
  System: ❌ No broker commission summary report
  
FIX:
  1. Report: Broker-wise commission summary
     ├─ Date range filter
     ├─ By month/quarter
     ├─ Total sales invoices via broker
     ├─ Total commission charged
     ├─ Net settlement amount (total - prior payments)
  2. Export: CSV for payment processing
  3. Drill-down: Click broker → list of invoices
```

### 6.2 Wave 2: Advanced Broker Features (DEFER)

- **Consignment Model**: Broker holds inventory on our behalf; settle net of broker commission & expenses
- **Depot Model**: Goods stored at broker's depot; sales & settlement tracking
- **Broker Margin/Security Deposit**: Track initial deposit, adjust with transactions
- **Broker Charges**: Storage, handling, insurance (separate from commission)

---

## 7. REQUIRED SCHEMA CHANGES (DATABASE MIGRATIONS)

### 7.1 Account Master Enhancements

```sql
-- Add broker/consignment/depot fields
ALTER TABLE accounts ADD COLUMN credit_limit DECIMAL(15,2) DEFAULT 0;
ALTER TABLE accounts ADD COLUMN interest_rate DECIMAL(5,2) DEFAULT 0;  -- % per annum
ALTER TABLE accounts ADD COLUMN account_type VARCHAR(20) DEFAULT 'CUSTOMER';  -- CUSTOMER | SUPPLIER | BROKER | EMPLOYEE
ALTER TABLE accounts ADD COLUMN is_consignment BOOLEAN DEFAULT FALSE;
ALTER TABLE accounts ADD COLUMN is_depot BOOLEAN DEFAULT FALSE;
ALTER TABLE accounts ADD COLUMN is_pay_to_mill BOOLEAN DEFAULT FALSE;
ALTER TABLE accounts ADD COLUMN tds_percentage DECIMAL(5,2) DEFAULT 0;

-- Create index for account type queries
CREATE INDEX idx_account_type ON accounts(account_type);
```

### 7.2 Product Master Enhancements

```sql
-- Add price guardrails
ALTER TABLE products ADD COLUMN min_selling_price DECIMAL(15,2) DEFAULT 0;
ALTER TABLE products ADD COLUMN max_purchase_price DECIMAL(15,2) DEFAULT 0;

CREATE INDEX idx_min_price ON products(min_selling_price);
CREATE INDEX idx_max_price ON products(max_purchase_price);
```

### 7.3 Order/Invoice Enhancements

```sql
-- Add IRN, broker, RCM/TCS fields
ALTER TABLE invoices ADD COLUMN irn_number VARCHAR(16) UNIQUE;
ALTER TABLE invoices ADD COLUMN irn_qr_code LONGTEXT;  -- base64 QR code
ALTER TABLE invoices ADD COLUMN irn_ack_number VARCHAR(16);
ALTER TABLE invoices ADD COLUMN irn_generated_at TIMESTAMP;

ALTER TABLE invoices ADD COLUMN broker_account_id BIGINT;  -- FK to Account
ALTER TABLE invoices ADD COLUMN broker_commission_percent DECIMAL(5,2) DEFAULT 0;
ALTER TABLE invoices ADD COLUMN broker_commission_amount DECIMAL(15,2) DEFAULT 0;

ALTER TABLE invoices ADD COLUMN is_rcm_applicable BOOLEAN DEFAULT FALSE;
ALTER TABLE invoices ADD COLUMN rcm_amount DECIMAL(15,2) DEFAULT 0;

ALTER TABLE invoices ADD COLUMN is_tcs_applicable BOOLEAN DEFAULT FALSE;
ALTER TABLE invoices ADD COLUMN tcs_amount DECIMAL(15,2) DEFAULT 0;

ALTER TABLE invoices ADD COLUMN tds_percentage DECIMAL(5,2) DEFAULT 0;  -- from supplier account
ALTER TABLE invoices ADD COLUMN tds_amount DECIMAL(15,2) DEFAULT 0;

-- Indexes
CREATE INDEX idx_irn_number ON invoices(irn_number);
CREATE INDEX idx_broker_account ON invoices(broker_account_id);
```

### 7.4 GoodsIssue/GoodsReceipt Clarification

```sql
-- Clarify: Are these separate entities or just statuses on Order?
-- Recommendation: Separate entities for audit trail

CREATE TABLE goods_issues (
  id BIGINT PRIMARY KEY,
  sales_order_id BIGINT NOT NULL,  -- FK
  issue_date TIMESTAMP DEFAULT NOW(),
  godown_id BIGINT NOT NULL,  -- FK to Warehouse
  status VARCHAR(20),  -- PENDING | PARTIAL | COMPLETED
  created_by BIGINT,
  created_at TIMESTAMP DEFAULT NOW(),
  FOREIGN KEY (sales_order_id) REFERENCES orders(id)
);

CREATE TABLE goods_issue_lines (
  id BIGINT PRIMARY KEY,
  goods_issue_id BIGINT NOT NULL,  -- FK
  order_item_id BIGINT NOT NULL,  -- FK
  qty_issued DECIMAL(12,2),
  created_at TIMESTAMP DEFAULT NOW(),
  FOREIGN KEY (goods_issue_id) REFERENCES goods_issues(id)
);

-- Similar for goods receipts (inbound)
CREATE TABLE goods_receipts (
  id BIGINT PRIMARY KEY,
  purchase_order_id BIGINT NOT NULL,  -- FK
  receipt_date TIMESTAMP DEFAULT NOW(),
  godown_id BIGINT NOT NULL,  -- FK to Warehouse
  status VARCHAR(20),  -- PENDING | PARTIAL | RECEIVED | QC_HOLD
  created_by BIGINT,
  created_at TIMESTAMP DEFAULT NOW(),
  FOREIGN KEY (purchase_order_id) REFERENCES purchase_orders(id)
);
```

---

## 8. TESTING SCENARIOS (UAT)

### 8.1 Sales Flow Test Cases

**Test Case 1: Sale with credit limit enforcement**
```
Setup: 
  Customer = ABC Trading
  Credit Limit = INR 100,000
  Min Selling Price = INR 500/unit

Test Steps:
  1. Create sale order: 250 units @ INR 600 = INR 150,000 total
  2. System should BLOCK: "Customer credit limit INR 100K exceeded. Invoice total INR 150K"
  3. Reduce qty to 150 units = INR 90,000
  4. System should ALLOW
  5. Finalize invoice → IRN should be generated within 30 seconds
  6. GL entries should be auto-created: AR debit, Sales revenue credit

Expected Result: ✅ PASS if credit limit blocks invoice; IRN generates; GL posts
```

**Test Case 2: Sale to unregistered customer with RCM**
```
Setup:
  Customer = XYZ (unregistered, gstStatus = 'unregistered')
  Sale amount = INR 100,000
  CGST/SGST rate = 18% (9% each)

Test Steps:
  1. Create sale invoice to XYZ for INR 100,000
  2. System should detect: RCM applicable = true
  3. Tax calculation: CGST/SGST split, but marked as buyer liability
  4. GL entries:
     ├─ Debit: AR = 100,000 + 18,000 tax = 118,000
     ├─ Credit: Sales Revenue = 100,000
     ├─ Credit: RCM Recovery (not Tax Payable) = 18,000
  5. Invoice display: "RCM Applicable" badge
  6. IRN generation should succeed

Expected Result: ✅ PASS if RCM flag set; GL posted to RCM account; badge displayed
```

**Test Case 3: Sale via broker with commission**
```
Setup:
  Customer = ABC Trading
  Broker = XYZ Broker
  Sale amount = INR 100,000
  Broker commission = 2%

Test Steps:
  1. Create sale invoice with Broker = XYZ Broker, Commission = 2%
  2. System should calculate: Commission = 100,000 × 2% = INR 2,000
  3. GL entries:
     ├─ Debit: AR = 100,000 + tax
     ├─ Credit: Sales Revenue = 100,000
     ├─ Debit: Commission Expense = 2,000
     └─ Credit: AP—Broker = 2,000
  4. Finalize → IRN generated

Expected Result: ✅ PASS if commission calculated; GL entries created; broker AP recorded
```

### 8.2 Purchase Flow Test Cases

**Test Case 1: Purchase with GRN matching**
```
Setup:
  Supplier = Steel Corp
  PO: 10 MT @ INR 50,000/MT = INR 500,000

Test Steps:
  1. Create Purchase Order: 10 MT @ INR 50,000
  2. Receive goods: Create GRN for 9.5 MT (partial receipt; 0.5 MT shortage)
  3. Supplier sends invoice: INR 475,000 (for 9.5 MT)
  4. 3-way match: PO qty=10, GRN qty=9.5, Invoice qty=9.5 → ✅ Match
  5. Finalize invoice → GL posting: Inventory debit, AP credit
  6. Report shortage: Variance = 0.5 MT @ 50K = INR 25K

Expected Result: ✅ PASS if GRN created; 3-way match validated; GL posts correctly
```

**Test Case 2: Purchase with TDS**
```
Setup:
  Supplier = Contractor ABC
  TDS Rate = 2% (contractor)
  Invoice = INR 100,000

Test Steps:
  1. Create purchase invoice: INR 100,000
  2. System should detect: TDS 2% applicable (from supplier account)
  3. TDS amount = 100,000 × 2% = INR 2,000
  4. GL entries:
     ├─ Debit: COGS = 100,000
     ├─ Debit: TDS Receivable = 2,000
     └─ Credit: AP = 100,000 (net of TDS shown separately)
  5. On payment: Net payment = 100,000 - 2,000 = 98,000
  6. Create TDS remittance voucher (quarterly)

Expected Result: ✅ PASS if TDS calculated; GL separated; payment net of TDS
```

### 8.3 Inventory Flow Test Cases

**Test Case 1: Stock reservation on sale → deduction on dispatch**
```
Setup:
  Product = Rice (onhand = 100 MT in Godown A)
  
Test Steps:
  1. Create sale order: 30 MT Rice
  2. Check stock balance: OnHand=100, Reserved=30, Available=70
  3. Create GatePass (dispatch): 30 MT
  4. Check stock balance: OnHand=70, Reserved=0
  5. Invoice & finalize
  6. GL entry: COGS debit, Inventory credit (30 MT × cost)

Expected Result: ✅ PASS if reservation updates stock; dispatch deducts; GL posts
```

---

## 9. SUMMARY: WHAT NEEDS TO BE FIXED

### ✅ Already Works (No Fix Needed)
- Basic order/invoice creation
- User & company master
- Basic inventory stock tracking

### ❌ MUST FIX (Wave 1 Blockers)

1. **Account Master**: Add 6 fields (credit limit, interest rate, account type, consignment/depot flags, TDS %)
2. **Product Master**: Add min selling & max purchase price
3. **E-Invoice IRN**: Integrate with IRP API; generate & store IRN on invoice finalize
4. **GST Tax Engine**: RCM, TCS, CGST/SGST/IGST split with tax rule precedence
5. **GL Auto-Posting**: Auto-create JournalEntry when invoice finalized
6. **Broker Commission**: Track & post broker commission GL entries

### ⚠️ CLARIFY (Wave 1 Scope)

7. **GoodsIssue/GRN Workflow**: Define clear entities & relationships
8. **3-Way Match**: Document PO-GRN-Invoice matching workflow
9. **Godown Management**: Multi-warehouse inventory scope

### 🟢 DEFER (Wave 2)

10. **Bag Tracking**: Physical goods lifecycle tracking
11. **Interest Accrual**: Auto-calculation on late payment (manual workaround for Wave 1)
12. **Bank Reconciliation**: Statement matching tool
13. **Consignment/Depot**: Special account types & settlement
14. **Year-End Closing**: Financial closing procedures

---

**Next Step:** Prioritize fixes 1-6 for Sprint 1-3 execution. Clarify items 7-9 before development. Defer items 10-14 to Wave 2 planning.
