# Gap Analysis: HITRIX (VB6) vs DhanMan (Microservices)

> **Status:** Phase 1 complete (HITRIX documented). Phase 2 pending (DhanMan source code review).  
> This document will be updated once DhanMan source code is provided.

---

## Context

| | HITRIX | DhanMan |
|---|---|---|
| Type | Legacy VB6 monolith | Modern microservices (.NET 9) |
| Domain | Indian textile/commodity trading | Apartment/community management + Finance |
| Database | Microsoft Access / Oracle | PostgreSQL 17 |
| Clients | Textile trading firms | Apartment associations |
| Goal | Source system to migrate from | Target system to migrate to |

The goal of this gap analysis is to identify what DhanMan must **build, extend, or adapt** to fully support HITRIX clients (textile trading firms).

---

## DhanMan Services Relevant to HITRIX

| DhanMan Service | Likely HITRIX Overlap |
|---|---|
| `dhanman-common` | Chart of accounts, organization/company setup, user management |
| `dhanman-sales` | Sales invoices, receipts, customer outstanding |
| `dhanman-purchase` | Purchase bills, payments, vendor outstanding |
| `dhanman-inventory` | Item/commodity master, warehouse, stock tracking |
| `dhanman-payroll` | Minimal overlap (payroll not prominent in HITRIX) |
| `dhanman-community` | No overlap (apartment-specific) |
| `dhanman-document` | Report storage, gate pass PDFs, tax forms |

---

## Preliminary Gap Assessment (by Domain)

> **Confidence level:** Medium — based on HITRIX documentation + DhanMan context doc.  
> Requires DhanMan source code review to confirm exact gaps.

---

### Domain 1: Masters & Configuration

| HITRIX Feature | DhanMan Status | Gap Level | Notes |
|---|---|---|---|
| Company Master (GSTIN/PAN/TAN) | Partial | Medium | DhanMan has company setup but tax reg fields TBD |
| Account / Chart of Accounts | Exists | Low | dhanman-common has chart of accounts |
| Account Group / Hierarchy | Exists | Low | Standard accounting grouping likely present |
| Party / Weaver Group Master | Missing | High | Textile-specific party classification |
| Count / Unit of Measure | Partial | Medium | Inventory UOM likely exists; textile-specific counts TBD |
| Godown / Warehouse Master | Partial | Medium | dhanman-inventory has warehouses |
| Tax Narration Master | Missing | High | HITRIX-specific: narration → tax account mapping |
| Commodity / Item Master with HSN | Partial | Medium | Inventory items exist; textile-specific attributes TBD |
| Transport Master | Missing | High | Not a DhanMan concept |
| Sizer / Consignee Master | Missing | High | Textile-specific third-party warehouse |
| System Settings — Default Tax Accounts | Missing | High | Auto-posting defaults not present in DhanMan |

---

### Domain 2: Booking & Order Management

| HITRIX Feature | DhanMan Status | Gap Level | Notes |
|---|---|---|---|
| Party / Job Work Booking | Missing | Critical | No equivalent in DhanMan |
| Trade Purchase Booking | Missing | Critical | No purchase order concept in DhanMan |
| SIT Purchase Booking | Missing | Critical | Stock-in-transit is not a DhanMan concept |
| Consignment / Depot Booking | Missing | Critical | No consignment workflow |
| Mill Bill Booking | Missing | Critical | No mill/job-work order concept |
| Booking Status Tracking | Missing | Critical | No pre-transaction booking workflow |
| Booking → Bill Conversion | Missing | Critical | Conversion from booking to bill |
| Booking Confirmation WhatsApp | Missing | High | Notification on booking |

---

### Domain 3: Sales

| HITRIX Feature | DhanMan Status | Gap Level | Notes |
|---|---|---|---|
| Basic Sales Invoice | Exists | Low | dhanman-sales has invoice creation |
| Trade Sale Type (SY) | Partial | Medium | DhanMan invoice exists; trade-specific fields TBD |
| Consignment Sale (SO) | Missing | Critical | No consignment liability/sale concept |
| Depot Sale (SD) | Missing | High | No depot arrangement |
| SIT Sale (ST) | Missing | Critical | No stock-in-transit sale |
| Hank Sale (SYHNK) | Missing | High | Textile-specific — yarn hank sales |
| Mill Bill (SM) | Missing | Critical | No mill processing fee invoice |
| GST Invoice (CGST/SGST/IGST) | Partial | Medium | DhanMan likely has GST; textile-specific config TBD |
| IRN / E-Invoicing | Unknown | Medium | e-invoicing integration status unknown |
| HSN Code at Invoice Line | Partial | Medium | May exist in dhanman-sales; needs verification |
| Sales Return with Credit Note | Exists | Low | DhanMan has sales returns |
| Late Payment Interest Debit Note | Missing | High | Auto-calculate and raise interest debit note |
| Credit / Debit Note GST | Partial | Medium | DhanMan may have credit notes |
| Bill-to-Bill Settlement (Receipt) | Partial | Medium | DhanMan has payments; bill matching TBD |
| Broker / Commission Agent | Missing | High | Broker account and commission tracking |
| Prefix-based Invoice Numbering | Unknown | Medium | Custom prefix numbering |
| TCS on Sales | Missing | High | TCS collection not in DhanMan scope |
| Multiple Sale Types per Company | Missing | High | Type codes SY/SO/SD/ST/SM need mapping |

---

### Domain 4: Purchase

| HITRIX Feature | DhanMan Status | Gap Level | Notes |
|---|---|---|---|
| Basic Purchase Bill | Exists | Low | dhanman-purchase has bill creation |
| Trade Purchase (PY) | Partial | Medium | Direct purchase exists; textile attributes TBD |
| SIT Purchase (PT) | Missing | Critical | Stock-in-transit purchase not in DhanMan |
| Consignment Purchase | Missing | Critical | No consignment workflow |
| Hank Purchase (PYHNK) | Missing | High | Textile-specific |
| GST Purchase (CGST/SGST/IGST) | Partial | Medium | DhanMan purchase has GST; config TBD |
| Reverse Charge Mechanism (RCM) | Unknown | High | RCM support in dhanman-purchase TBD |
| Purchase Return / Debit Note | Exists | Low | DhanMan has purchase returns |
| Goods Inward / Receipt Note | Unknown | Medium | Goods inward with bag tracking TBD |
| Weight Reconciliation | Missing | High | Bill weight vs actual weight — textile-specific |
| Bag-level Inward Tracking | Missing | Critical | Bag numbers and weights not in DhanMan |
| TDS on Purchase | Unknown | High | TDS deduction at payment/bill stage |
| Multiple Inwards per Bill | Unknown | Medium | Partial delivery linkage |
| Brokerage on Purchase | Missing | High | Broker commission on purchase side |

---

### Domain 5: Finance & Accounting

| HITRIX Feature | DhanMan Status | Gap Level | Notes |
|---|---|---|---|
| Payment Voucher | Exists | Low | dhanman-common has payments |
| Receipt Voucher | Exists | Low | dhanman-common has receipts |
| Journal Voucher | Exists | Low | DhanMan has manual journal entries |
| Bank Reconciliation | Exists | Low | DhanMan has bank reconciliation with import |
| Bill-to-Bill Settlement | Partial | Medium | DhanMan payment matching; needs bill linkage |
| Discount at Payment | Unknown | Medium | Discount/adjustment at payment time |
| TDS at Payment | Unknown | High | TDS deduction mechanics |
| Outstanding Aging (4 buckets) | Partial | Medium | DhanMan has outstanding; aging buckets TBD |
| Late Payment Interest | Missing | High | Interest accrual on overdue — auto debit note |
| Multi-Account Journal | Unknown | Medium | Multi-line journal entry |
| Final Accounts (P&L / BS) | Exists | Low | DhanMan has P&L, BS, Trial Balance |
| Cash Book / Bank Book | Partial | Medium | Day-book style reports TBD |
| Party Confirmation Letter | Unknown | Medium | Account confirmation letter format |
| Broker-wise Outstanding | Missing | High | Broker-grouped outstanding — textile-specific |
| Mill-wise Statement | Missing | High | Mill-specific statements |

---

### Domain 6: Inventory & Logistics

| HITRIX Feature | DhanMan Status | Gap Level | Notes |
|---|---|---|---|
| Commodity / Item Master | Partial | Medium | dhanman-inventory has products; textile attributes TBD |
| Warehouse / Godown Master | Exists | Low | dhanman-inventory has warehouses |
| Bag Entry (Bag-level Inward) | Missing | Critical | No bag-level tracking in DhanMan |
| Sequential Bag Numbering | Missing | Critical | No bag number concept |
| Individual Bag Weight Capture | Missing | Critical | Weight at bag level — textile-specific |
| Lot / Batch Tracking | Unknown | High | DhanMan inventory batch tracking TBD |
| Gate Pass — Inward | Missing | Critical | No gate pass concept in DhanMan |
| Gate Pass — Outward | Missing | Critical | No outward gate pass |
| Gate Pass Printing | Missing | Critical | PDF gate pass format |
| Gate Pass WhatsApp | Missing | High | Gate pass delivery via WhatsApp |
| Delivery Confirmation | Unknown | Medium | GP delivery confirmation flow |
| Stock Balance by Item | Partial | Medium | Inventory stock balance TBD |
| SIT Stock Tracking | Missing | Critical | Transit stock not in DhanMan |
| Weight Shortfall / Loss | Missing | High | Quality/loss tracking at inward |
| CST Form Entry (C/E1/E2/F) | Missing | High | Pre-GST forms — needed for older transactions |

---

### Domain 7: Tax Compliance

| HITRIX Feature | DhanMan Status | Gap Level | Notes |
|---|---|---|---|
| GST Sales Invoice | Partial | Medium | DhanMan has GST; textile-specific config TBD |
| GST Purchase Bill | Partial | Medium | DhanMan purchase has GST |
| CGST / SGST / IGST Calculation | Exists | Low | DhanMan has GST calculation |
| HSN Code per Transaction | Partial | Medium | HSN may exist in DhanMan sales |
| IRN / E-Invoicing | Unknown | Medium | E-invoicing integration status TBD |
| RCM (Reverse Charge) | Unknown | High | RCM flag and posting TBD |
| GST Register (GSTR-1 / GSTR-2) | Unknown | High | Export data for GST returns |
| TDS Entry and Tracking | Unknown | High | TDS compliance in DhanMan TBD |
| TDS Certificate Management | Missing | High | TDS certificate number tracking |
| TCS Management | Missing | High | TCS collection — not in DhanMan scope |
| CST C-Form Management | Missing | High | Pre-GST form tracking (legacy needed) |
| CST E1/E2/F-Form Tracking | Missing | High | Pre-GST forms for older transactions |

---

### Domain 8: Reporting & Analytics

| HITRIX Feature | DhanMan Status | Gap Level | Notes |
|---|---|---|---|
| Account Ledger | Exists | Low | Customer/vendor ledgers in DhanMan |
| Party Confirmation Letter | Unknown | Medium | Formatted confirmation letter TBD |
| Outstanding Aging Report | Partial | Medium | Basic outstanding in DhanMan; aging buckets TBD |
| Sales Summary (Party/Mill/Broker/Count) | Partial | High | DhanMan has sales reports; textile dimensions missing |
| Purchase Summary (Mill/Count) | Partial | High | Same as sales — textile dimensions missing |
| P&L / Balance Sheet | Exists | Low | DhanMan has financial reports |
| GST Register Report | Unknown | High | GSTR format export TBD |
| C/E1/E2/F-Form Reports | Missing | High | 30+ pre-GST tax reports — not in DhanMan |
| TDS Report | Unknown | High | TDS compliance reporting TBD |
| Stock Balance Report | Partial | Medium | Inventory reports TBD |
| Goods Inward Reports | Missing | High | Inward summary by mill/count |
| Gate Pass Register | Missing | High | No gate pass in DhanMan |
| Bar/Line/Pie Charts | Partial | Low | DhanMan has charts; scope TBD |
| WhatsApp Report Delivery | Missing | High | WhatsApp PDF delivery not in DhanMan |
| SMS Outstanding Reports | Missing | High | SMS integration not in DhanMan |
| Broker-wise Reports | Missing | High | Broker dimension missing across reports |
| Mill-wise Reports | Missing | High | Mill dimension missing |

---

### Domain 9: Communication

| HITRIX Feature | DhanMan Status | Gap Level | Notes |
|---|---|---|---|
| WhatsApp API Integration | Partial | High | DhanMan has FCM; WhatsApp PDF API not present |
| Ledger / Statement via WhatsApp | Missing | High | No PDF-to-WhatsApp flow |
| Gate Pass via WhatsApp | Missing | Critical | No gate pass in DhanMan |
| Booking Confirmation via WhatsApp | Missing | High | No booking module |
| SMS Integration | Missing | High | No SMS gateway in DhanMan |
| Outstanding SMS Reminders | Missing | High | Bulk SMS for overdue — not in DhanMan |
| Check Pass Audit Alert (SMS/WA) | Missing | Medium | Modification alert to admin |

---

### Domain 10: Administration & Security

| HITRIX Feature | DhanMan Status | Gap Level | Notes |
|---|---|---|---|
| User Management | Exists | Low | DhanMan has users with Auth0 |
| Role-based Access Control | Exists | Low | DhanMan has permissions system |
| Per-form Privilege Matrix | Partial | Medium | DhanMan has permissions; form-level granularity TBD |
| Audit Lock (transaction) | Unknown | Medium | VIsAudited concept — financial audit lock TBD |
| User Activity Tracking | Exists | Low | DhanMan tracks created_by user |
| Data Backup | Unknown | Low | OVH server backups in place |
| Fiscal Year Management | Partial | Medium | DhanMan financial year handling TBD |
| Multi-Company at Transaction Level | Exists | Low | DhanMan has multi-company (org/company model) |
| Year-End Carryforward | Unknown | High | Outstanding carryforward automation TBD |
| Master Copy Utility (firm → firm) | Missing | Medium | Onboarding utility for new clients |
| Check Pass Alert | Missing | Medium | Real-time modification alert |

---

## Gap Summary by Priority

### Critical Gaps (must build before first HITRIX client can migrate)
1. **Booking / Pre-order module** — entire domain missing
2. **Bag-level inventory tracking** — bag numbers, individual weights
3. **Gate Pass system** — inward and outward gate passes
4. **SIT (Stock-in-Transit) concept** — across purchase, sales, inventory
5. **Consignment sale/purchase workflows** — SO/PO types
6. **Mill Bill (SM) transaction type** — processing fee invoicing
7. **Broker / Commission Agent** — across sales, purchase, outstanding

### High Priority Gaps (needed for core textile trading operations)
8. **Tax Narration Master** — narration → tax account auto-mapping
9. **Late Payment Interest Debit Note** — auto-calculate and raise
10. **TDS compliance** — entry, certificate, reporting
11. **TCS management** — collection and reporting
12. **Textile-specific sale types** — Hank, Mill, Trade, SIT
13. **Weight reconciliation** — bill weight vs actual
14. **WhatsApp PDF delivery** — ledger, statement, gate pass
15. **Outstanding with broker dimension** — broker-wise aging
16. **GST Register export** — GSTR-1/GSTR-2 format
17. **CST Form tracking** — C/E1/E2/F forms (legacy clients)
18. **Goods inward reports** — mill-wise, count-wise summaries
19. **Year-end carryforward** — outstanding balances

### Medium Priority Gaps (needed for full feature parity)
20. Prefix-based invoice numbering
21. Party / Weaver group master
22. Transport master
23. Sizer / Consignee master
24. Multi-line journal (multi-account JV)
25. Discount at payment
26. Per-form privilege matrix
27. Party confirmation letter format
28. Audit lock on transactions
29. Check pass alert system
30. SMS integration

---

## Next Steps

1. **Provide DhanMan source code access** → Review each service for actual feature state
2. **Confirm gap levels** → Move from "Unknown/Partial" to confirmed state
3. **Prioritize by client requirements** → First client may not need all features
4. **Design new DhanMan modules** for critical gaps (booking, gate pass, bag tracking)
5. **Extend existing DhanMan services** for medium gaps (sale types, TDS, TCS)
