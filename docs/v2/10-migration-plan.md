---
id: v2-migration-plan
title: HITRIX → DhanMan Migration Plan
---

# HITRIX → DhanMan Migration Plan

> Goal: Migrate HITRIX V2 clients (Indian textile trading firms) onto DhanMan within **2–3 months**. This document identifies what DhanMan already covers, what gaps exist, and a phased delivery plan.

---

## DhanMan Service Mapping

DhanMan is a .NET 9 microservices platform with 7 services: `common`, `sales`, `purchase`, `inventory`, `payroll`, `document`, `community/myhome`.

| HITRIX Domain | HITRIX Tables | DhanMan Service | Coverage |
|---|---|---|---|
| Company Setup | `tblMastCompany`, `tblCompanyYear` | `common` | Partial — no multi-firm-in-one-DB, no fiscal year code ranges |
| Account Groups | `tblMastGroup`, `tblMastSubGroup`, `tblMastCompGroup` | `common` (CoA) | Partial — hierarchy exists but textile-specific groups missing |
| Ledger Accounts | `tblMastAccount` | `common` | Partial — 40+ textile-specific flags not in DhanMan |
| Items / Products | `tblMastItem` | `inventory` | Partial — ticket/count, HSN, brokerage rate, exemption rate missing |
| Tax Codes | `tblMastNarration` (type=T) | `common` | Partial — no dual-rate cotton/poly split |
| Godowns | `tblMastNarration` (type=G) | `inventory` | Partial — no godown address |
| Transport | `tblMastNarration` (type=R) | `common` | Minimal |
| Bill Serials | `tblMastBillSerial` | Not present | **GAP** |
| Delivery Addresses | `tblMastDeleAdd` | `common` | Partial — no per-party multiple delivery addresses |
| Interest Rates | `tblMastFirmwiseIntRate` | Not present | **GAP** |
| Trade Booking | `tblBooking` + `tblBooKingSub` | Not present | **CRITICAL GAP** |
| Booking Despatch | `tblBooKingDesp` | Not present | **CRITICAL GAP** |
| Trade Sale (SY) | `tblSale` + `tblSaleSub` | `sales` | Partial — missing GST complexity |
| Consignment Sale (SO) | `tblSale` (VType=SO) | Not present | **CRITICAL GAP** |
| Depot Sale (SD) | `tblSale` (VType=SD) | Not present | **CRITICAL GAP** |
| SIT Sale (ST) | `tblSale` (VType=ST) | Not present | **GAP** |
| Mill Bill (SM) | `tblSale` (VType=SM) | Not present | **GAP** |
| Sales Return (RY) | `tblSale` (VType=RY) | `sales` | Partial |
| Trade Purchase (PY) | `tblPurch` + `tblPurchSub` | `purchase` | Partial |
| SIT Purchase (PT) | `tblPurch` (VType=PT) | Not present | **GAP** |
| Purchase Inward (PI) | `tblPurch` (VType=PI) | `purchase` / `inventory` | Partial |
| Purchase Return (VY) | `tblPurch` (VType=VY) | `purchase` | Partial |
| RCM on Purchase | `PurIsRCMBill`, `PurRCM*` | Not present | **CRITICAL GAP** |
| Gate Pass | `tblGatePass` + `tblGatePassSub` | `inventory` | Partial — no GP-linked sale |
| Bag Tracking | `tblBags` | Not present | **CRITICAL GAP** |
| Cash/Bank Vouchers | `tblVoucher` (CP/BP/CR/BR) | `common` | Present |
| Journal Voucher | `tblVoucher` (JV) | `common` | Present |
| Bank Reconciliation | `VReconDt` on `tblVoucher` | `common` | Partial |
| Receipt Matching | `tblIntSale` (VBillVno link) | Not present | **CRITICAL GAP** |
| TDS on Purchase | `PurTdsJvNo`, `PurTdsRate` | Not present | **CRITICAL GAP** |
| TDS on Sale | `SlTdsRate`, `SlTdsAmt` | Not present | **CRITICAL GAP** |
| TCS on Sale | `SlTcsAmt`, `SlTcsRate` | Not present | **GAP** |
| Interest / Late Pay | `tblIntDbNote`, `tblIntSale` | Not present | **CRITICAL GAP** |
| GST Cr/Dr Note | `tblIntSale` (IsGSTDbNt) | Not present | **GAP** |
| E-Invoice / IRN | `SlIRNNo`, `SlAckNo` | `document` | Partial — no IRP integration |
| GSTR-2A Matching | `PrcPrepareGSTR2A` | Not present | **GAP** |
| GSTR-3B | `PrcPrepareGSTR3B` | Not present | **GAP** |
| Outstanding Ageing | `PrcPrepareOutStangingSale` | `sales` / `purchase` | Partial |
| Mill Statement | `PrcPrepareMillStatement` | Not present | **GAP** |
| Broker Commission | `PrcPrepareBrokSale` | Not present | **GAP** |
| Daily Report | `PrcPrepareDailyReport` | Not present | **GAP** |
| Stock Reports | `PrcPrepareStock`, `PrcPrepareLotwiseStock` | `inventory` | Partial |
| Audit Log | `tblSale_Log` etc. | `common` | Present (audit trail) |
| User Permissions | `tblUserPrevilage` | Auth0 + RBAC | Present |
| Multi-company | `VFirm` on all tables | Partial in common | Architecture gap |
| Year-end | `PrcPrepareYearEnd` | Not present | **GAP** |

---

## Gap Summary by Priority

### P0 — Critical (blocks go-live)

| # | Gap | Impact | DhanMan Work Needed |
|---|---|---|---|
| 1 | **Pre-sale Booking module** | Core workflow for all clients — they book orders weeks before goods arrive | New module in `sales` service |
| 2 | **Receipt-to-invoice matching** | Every payment must be matched to specific bills with discount/TDS/interest | Extend `common` or `sales` |
| 3 | **TDS on purchase and sale** | Legally required; generates JV entries, challan tracking | New in `purchase` + `sales` |
| 4 | **Interest / Late payment debit notes** | Core revenue for trading firms — automatic interest calculation | New module in `sales` |
| 5 | **RCM (Reverse Charge Mechanism)** | Mandatory for certain purchase categories under GST | Extend `purchase` service |
| 6 | **Sale types: Consignment (SO), Depot (SD), SIT (ST), Mill Bill (SM)** | 4 of 5 sale types are missing — only basic trade sale exists | Extend `sales` service |
| 7 | **Bag-level inventory tracking** | Textile-specific — individual bag with weight tracked from inward to sale | New in `inventory` service |

### P1 — High (needed within first month)

| # | Gap | Impact | DhanMan Work Needed |
|---|---|---|---|
| 8 | **GSTR-2A and GSTR-3B** | GST compliance — mandatory filing | New reports in `document` or `sales` |
| 9 | **E-Invoice IRN integration** | Mandatory for firms >₹5Cr turnover | Extend `document` service |
| 10 | **GST Cr/Dr notes on interest** | Required for interest charged with GST | Extend `sales` |
| 11 | **TCS on sale** | Mandatory for firms >₹50L turnover | Extend `sales` |
| 12 | **Multi-delivery addresses per party** | Used in every booking | Extend `common` party master |
| 13 | **Bill serial management** | Per-mill, per-sale-type auto serial numbers | New in `common` or `sales` |
| 14 | **Gate pass linked to sale** | GP creates sale outstanding, must reconcile | Extend `inventory` service |

### P2 — Medium (second month)

| # | Gap | Impact | DhanMan Work Needed |
|---|---|---|---|
| 15 | **Mill statement / commission report** | Monthly settlement report for mills | New report |
| 16 | **Broker commission statement** | Monthly brokerage calculation | New report |
| 17 | **Daily report (email-ready)** | Operations summary sent by email daily | New report + email trigger |
| 18 | **Lot-wise and godown-wise stock** | Physical stock management | Extend `inventory` reports |
| 19 | **Firmwise interest rates** | Per-party interest rate override | Extend `common` party master |
| 20 | **Year-end procedure** | Rolls over balances to new year | Extend `common` admin |
| 21 | **Transaction code configuration** | Maps VType to GST transaction codes | Config in `common` |

---

## 3-Month Delivery Plan

### Month 1 — Foundation & Core Trading

**Week 1–2: Data Migration & Setup**
- [ ] Migrate `tblMastCompany`, `tblCompanyYear` → DhanMan company/tenant setup
- [ ] Migrate `tblMastGroup`, `tblMastAccount` → DhanMan chart of accounts
- [ ] Migrate `tblMastItem` → DhanMan product catalog (extend with textile fields)
- [ ] Migrate `tblMastNarration` (tax codes, godowns, transport) → DhanMan config
- [ ] Set up user accounts and RBAC permissions

**Week 3–4: Sale Types**
- [ ] Implement Consignment Sale (SO) in DhanMan `sales` service
- [ ] Implement SIT Sale (ST)
- [ ] Implement Depot Sale (SD)
- [ ] Implement Mill Bill (SM)
- [ ] Implement Sales Return (RY) with credit note

**Parallel: Purchase Types**
- [ ] Implement SIT Purchase (PT)
- [ ] Implement Purchase Inward (PI) with bag tracking hook
- [ ] Implement RCM flag on purchases

### Month 2 — Compliance & Financial Controls

**Week 5–6: TDS/TCS**
- [ ] TDS deduction at purchase entry (PY/PT) with JV generation
- [ ] TDS deduction at receipt (for applicable parties)
- [ ] TCS collection at sale (SY/SO above threshold)

**Week 5–6: Receipt Matching**
- [ ] Link receipts (CR/BR) to specific sale invoices
- [ ] Settlement logic: Discount, TDS, Interest deduction
- [ ] Outstanding ageing report with due date tracking

**Week 7–8: Interest & Debit Notes**
- [ ] Late payment interest calculation (configurable grace period + rate)
- [ ] Interest debit note generation with GST
- [ ] Interest credit note for returns

**Week 7–8: GST Compliance**
- [ ] GSTR-3B computation from sales/purchase data
- [ ] GSTR-2A upload and matching
- [ ] E-Invoice IRN integration with IRP API
- [ ] RCM computation and reporting

### Month 3 — Booking, Inventory & Reports

**Week 9–10: Booking Module**
- [ ] Pre-sale order entry (party, mill, item, qty, rate)
- [ ] Booking confirmation with mill reference
- [ ] Partial despatch tracking against booking
- [ ] Pending booking reports
- [ ] Booking vs actual comparison

**Week 9–10: Bag Tracking**
- [ ] Individual bag registration at inward (PI)
- [ ] Godown assignment per bag
- [ ] Bag-to-sale linkage (lot tracking)
- [ ] Lot-wise stock report

**Week 11–12: Reporting & Migration Cutover**
- [ ] Mill commission statement
- [ ] Broker commission statement
- [ ] Daily operations report (email)
- [ ] Opening balance migration from HITRIX
- [ ] Parallel run (2 weeks): HITRIX and DhanMan running simultaneously
- [ ] Cutover verification and sign-off

---

## Data Migration Approach

### Strategy: Direct SQL Migration

Since HITRIX uses SQL Server and DhanMan uses PostgreSQL, the migration will be:
1. Extract from HITRIX SQL Server using `sqlcmd` / SSMS
2. Transform (mapping HITRIX column names → DhanMan field names, normalizing enums)
3. Load into DhanMan PostgreSQL via API or direct SQL insert

### Critical Mapping Decisions

| HITRIX Concept | DhanMan Equivalent | Notes |
|---|---|---|
| `VFirm` (4-char) | `CompanyId` (GUID) | Map each VFirm to a DhanMan company |
| `VYear` (numeric) | `FinancialYearId` (GUID) | Map each VYear to a DhanMan fiscal year |
| `AcCode` (int) | `LedgerId` (GUID) | Re-map all account codes |
| `VNo+VType+VYear+VFirm` (composite) | Single `Id` (GUID) | Normalize to single PK |
| `NarrType=T` (tax code) | `TaxCodeId` | Map narration codes to tax config |
| `NarrType=G` (godown) | `WarehouseId` | Map to DhanMan warehouse |
| `tblMastGroup` hierarchy | DhanMan CoA groups | Flatten to 3-level hierarchy |

### Opening Balance Migration

The `PrcPrepareOpBalance` / `PrcPrepareYearEnd` stored procedures handle year-end in HITRIX. For migration:
1. Take trial balance as of cutover date from HITRIX
2. Import as opening balances in DhanMan
3. Outstanding invoices (from `tblSale`/`tblPurch`) migrate as open items

---

## DhanMan Architecture Extensions Required

The following new microservice capabilities or modules need to be built in DhanMan:

```
sales service additions:
├── SaleType enum: TRADE | CONSIGNMENT | DEPOT | SIT | MILL_BILL
├── BookingOrder (pre-sale order management)
├── BookingDespatch (shipment tracking against booking)
├── LatePaymentInterest (configurable grace + rate + GST)
├── ReceiptMatching (invoice-level settlement)
└── GSTCreditDebitNote

purchase service additions:
├── PurchaseType enum: TRADE | SIT | INWARD | OTHER
├── RCMBill (reverse charge processing)
└── TDSOnPurchase (TDS deduction + challan)

inventory service additions:
├── BagTracking (individual bag lifecycle)
├── GodownTransfer
└── LotWiseStock

common service additions:
├── MultiDeliveryAddress (per-party addresses)
├── BillSerial (per-mill serial sequences)
├── FirmwiseInterestRate
├── TransactionCode configuration
└── YearEnd procedure

new reporting capabilities:
├── MillStatement (commission calculation)
├── BrokerStatement (brokerage calculation)
├── DailyOperationsReport (email-ready)
├── GSTR2AMatching
├── GSTR3BComputation
└── OutstandingAgeing (with interest calc)
```

---

## Risk Register

| Risk | Probability | Impact | Mitigation |
|---|---|---|---|
| GST compliance gaps at go-live | Medium | Critical | Run GSTR reports in parallel for 1 month |
| Client resistance to UI change | High | Medium | Provide 1-week hands-on training + help desk |
| Opening balance mismatch | Medium | High | Reconcile trial balance before and after migration |
| TDS configuration errors | Medium | High | Test with past data, verify against last filed return |
| Booking module not ready in time | Low | High | Phase: run bookings in HITRIX until M3 |
| Performance (SQL Server → PostgreSQL) | Low | Medium | Load test with 1 year of client data |
| Multi-company data isolation | Medium | Critical | Strict `CompanyId` filter on all DhanMan queries |
