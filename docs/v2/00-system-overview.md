---
id: v2-overview
title: HITRIX V2 — System Overview
---

# HITRIX V2 — System Overview

> This document covers the **SQL Server Express** edition of HITRIX (codebase: `ABabu_Cd_15-05-2026`, DB: `vtc180426.bak`). This is the production version actively used by clients. The earlier Access-based version is documented under `docs/`.

---

## What HITRIX V2 Is

HITRIX V2 is a multi-company, multi-year VB6 ERP purpose-built for **Indian textile/commodity trading firms** (agents, brokers, depot operators). It manages the complete order-to-cash and procure-to-pay lifecycle for textile commodities (yarn, fabric, hank cotton) traded through:

- **Trade transactions** — direct buy/sell between party and mill
- **SIT (Station-in-Transit)** — goods in transit billing
- **Consignment** — mill places goods with agent, agent sells on behalf
- **Depot** — agent operates a physical depot for a mill
- **Mill Bill** — commission billing from mill to agent

---

## Architecture Evolution: Access → SQL Server

| Dimension | V1 (Access) | V2 (SQL Server) |
|---|---|---|
| Database | Microsoft Access `.mdb` | MS SQL Server Express 2014 |
| Transaction tables | Single `tblVoucher` for everything | Separate `tblSale`, `tblPurch`, `tblVoucher` (finance), `tblGatePass`, `tblBooking` |
| Multi-company | No | Yes — `VFirm` on every transaction |
| Multi-year | `VYear` field | `tblCompanyYear` + code ranges per year |
| Stored procedures | None (all SQL in VB6 code) | 51 stored procedures for all reports |
| Audit log | None | `_Log` mirror tables for all transaction tables |
| GST | Partial (late addition) | Full — CGST/SGST/IGST, RCM, TCS, e-Invoice IRN |
| Booking module | None | Full pre-sale order management |
| Interest calculation | Manual JV | Automated `tblIntDbNote` + `tblIntSale` with configurable grace |
| Bag tracking | `tblBags` — linked to inward | `tblBags` — linked to both inward and sale (full lifecycle) |
| Foreign keys | None (Access limitation) | Real FK constraints enforced at DB level |
| Communication | None | Email (CDO) + WhatsApp notifications |
| E-invoice | None | IRN generation via IRP API; `SlIRNNo`/`SlAckNo` stored on `tblSale` |

---

## Domain Map

```
┌─────────────────────────────────────────────────────────────────┐
│                        HITRIX V2 DOMAINS                        │
├─────────────────┬──────────────────┬──────────────────┬─────────┤
│   MASTERS       │   BOOKING        │   INVENTORY      │  ADMIN  │
│  Company        │  Trade Booking   │  Gate Pass       │  Users  │
│  Groups         │  SIT Booking     │  Bag Tracking    │  Access │
│  Accounts       │  Mill Booking    │  Delivery        │  Backup │
│  Items          │  Purchase Booking│  LR Details      │  Year   │
│  Narrations     │  Despatch        │  Godown Mgmt     │  End    │
│  Settings       │                  │                  │         │
├─────────────────┼──────────────────┼──────────────────┼─────────┤
│   SALES         │   PURCHASE       │   FINANCE        │ REPORTS │
│  Trade (SY)     │  Trade (PY)      │  Cash Payment    │  51 SPs │
│  SIT (ST)       │  SIT (PT)        │  Bank Payment    │  GST    │
│  Consignment(SO)│  Inward (PI)     │  Cash Receipt    │  Stock  │
│  Depot (SD)     │  Other (PO)      │  Bank Receipt    │  Party  │
│  Mill Bill (SM) │  Purchase Return │  Journal (JV)    │  Ledger │
│  Sales Return   │                  │  TDS Entry       │  Daily  │
│  Cr/Debit Note  │                  │  Bank Recon      │  Final  │
│  Late Pay Int   │                  │  Interest Calc   │  Booking│
│  Confirmation   │                  │  OS Adjustment   │         │
├─────────────────┴──────────────────┴──────────────────┴─────────┤
│              GST & TAX COMPLIANCE                                │
│  CGST/SGST/IGST · RCM · TCS · E-Invoice/IRN · GSTR2A · GSTR3B  │
└─────────────────────────────────────────────────────────────────┘
```

---

## Voucher Type Reference

Every transaction carries a `VType` (2-char) code. In V2, Sales and Purchase have their own dedicated tables (`tblSale`, `tblPurch`) in addition to `tblVoucher` for finance.

| VType | Description | Table | Module |
|---|---|---|---|
| `SY` | Trade Sale | `tblSale` + `tblSaleSub` | Sales |
| `SO` | Consignment Sale | `tblSale` + `tblSaleSub` | Sales |
| `SD` | Depot Sale | `tblSale` + `tblSaleSub` | Sales |
| `ST` | SIT Sale | `tblSale` + `tblSaleSub` | Sales |
| `SM` | Mill Bill | `tblSale` + `tblSaleSub` | Sales |
| `RY` | Sales Return (Trade) | `tblSale` + `tblSaleSub` | Sales |
| `PY` | Trade Purchase | `tblPurch` + `tblPurchSub` | Purchase |
| `PT` | SIT Purchase | `tblPurch` + `tblPurchSub` | Purchase |
| `PO` | Other Purchase | `tblPurch` + `tblPurchSub` | Purchase |
| `PI` | Purchase Inward | `tblPurch` + `tblPurchSub` | Purchase |
| `VY` | Purchase Return | `tblPurch` + `tblPurchSub` | Purchase |
| `CP` | Cash Payment | `tblVoucher` | Finance |
| `BP` | Bank Payment | `tblVoucher` | Finance |
| `CR` | Cash Receipt | `tblVoucher` | Finance |
| `BR` | Bank Receipt | `tblVoucher` | Finance |
| `JV` | Journal Voucher | `tblVoucher` | Finance |
| `GP` | Gate Pass | `tblGatePass` + `tblGatePassSub` | Inventory |
| `BK` | Trade Booking | `tblBooking` + `tblBooKingSub` | Booking |
| `BD` | Booking Despatch | `tblBooKingDesp` | Booking |
| `PN` | Debit Note | `tblIntDbNote` | Adjustments |
| `SN` | Credit Note | `tblIntSale` | Adjustments |
| `OP` | Opening Balance | `tblVoucher` | Setup |

---

## Multi-Company Architecture

Every transaction table has a `VFirm` column (nvarchar 4) which is the company code. This allows a single SQL Server database to serve multiple trading firms simultaneously — each firm's data is partitioned by `VFirm`.

`tblCompanyYear` defines the valid years and account code ranges per company-year combination. When a user logs in, they select company + year, which sets:
- `gCCode` / `gDbYr` (global vars)
- `gCSAcCode` / `gCEAcCode` — account code range for that company-year
- `gCSAgCode` / `gCEAgCode` — group code range
- `gCSItCode` / `gCEItCode` — item code range

All queries filter by both `VFirm` and year to ensure data isolation.

---

## Technology Stack

| Component | Technology |
|---|---|
| Application | VB6 (Visual Basic 6.0) |
| Database | MS SQL Server Express 2014 (upgraded to 2019 compatible) |
| DB Access | ADO (ADODB.Connection) via ODBC DSN |
| Reporting | Crystal Reports 8.5 (craxdrt.dll) |
| E-Invoice | WinHTTP (REST calls to IRP portal) |
| QR Code | IDAutomation + Morovia QRCode libraries |
| Email | Microsoft CDO for Windows 2000 |
| Excel Export | Microsoft Excel 12.0 Object Library (Office 2007+) |
| WhatsApp | Custom HTTP integration |
| Communication | SMS + WhatsApp via API |

---

## File Structure

```
ABabu_Cd_15-05-2026/
├── PrjVTC.vbp              # VB6 project file
├── Bascivar.bas            # Global variables and DB connection
├── CommonPro.bas           # Shared utility procedures (80KB)
├── Account.bas             # Core accounting logic (349KB)
├── DataCollection.bas      # Report data collection queries (99KB)
├── Stock.bas               # Stock calculation logic (70KB)
├── AccountRptProc.bas      # Account report procedures
├── ClsSearch.cls           # Search/lookup class
├── frmLogon.frm            # Login screen
├── frmSelCompany.frm       # Company + year selection
├── frmMain.frm             # Main menu / MDI parent
├── frmBooking*.frm         # 8 booking module forms
├── frmSales*.frm           # 9 sales forms
├── frmPurchase*.frm        # 7 purchase forms
├── frmPayment.frm          # Cash/bank payment
├── frmReceipt.frm          # Cash/bank receipt
├── frmJV*.frm              # 3 journal voucher forms
├── frmGatePass.frm         # Gate pass entry
├── frmBagEntry.frm         # Bag-level inventory
├── frmRpt*.frm             # 12 report forms
├── frmMast*.frm            # 13 master data forms
├── frmUtil*.frm            # 7 utility forms
└── frmAdm*.frm             # 2 admin forms
```
