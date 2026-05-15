---
id: intro
title: HITRIX System Overview
slug: /intro
---

# HITRIX VB6 System — Documentation

**System Name:** HITRIX (PrjHITRIX)  
**Technology:** VB6, Microsoft Access / Oracle, Crystal Reports 8.5  
**Domain:** Indian Commodity / Textile Trading ERP  
**Purpose:** This documentation was created to support the gap analysis between HITRIX (legacy) and DhanMan (modern microservices platform). The goal is to identify what DhanMan must build or enhance to onboard HITRIX clients.

---

## Document Index

| File | Domain |
|---|---|
| [01-masters.md](01-masters.md) | Masters & Configuration |
| [02-booking.md](02-booking.md) | Booking & Order Management |
| [03-sales.md](03-sales.md) | Sales (Pre-GST and GST) |
| [04-purchase.md](04-purchase.md) | Purchase (Pre-GST and GST) |
| [05-finance.md](05-finance.md) | Finance & Accounting |
| [06-inventory-logistics.md](06-inventory-logistics.md) | Inventory & Logistics |
| [07-tax-compliance.md](07-tax-compliance.md) | Tax Compliance (CST, GST, TDS) |
| [08-reporting.md](08-reporting.md) | Reporting & Analytics |
| [09-communication.md](09-communication.md) | Communication (SMS, WhatsApp) |
| [10-administration.md](10-administration.md) | Administration & Security |
| [feature-inventory.md](feature-inventory.md) | Full Feature Inventory (flat table) |
| [gap-analysis.md](gap-analysis.md) | Gap Analysis vs DhanMan *(to be completed)* |

---

## System Overview

HITRIX serves **Indian textile/commodity trading firms**. It handles the complete trade cycle from booking to delivery with full Indian tax compliance.

### Key Business Processes

**Sales Flow:**
1. Optional: Party Booking (pre-order tracking)
2. Sales Bill Entry (variants: Trade / Consignment / Depot / SIT / Mill Bill)
3. Goods Outward via Gate Pass
4. Optional: Sales Confirmation to customer
5. Receipt Voucher (against bill)
6. Optional: Late Payment Interest Debit Note (if overdue)

**Purchase Flow:**
1. Optional: Purchase Booking (pre-order tracking)
2. Purchase Bill Entry (variants: Trade / SIT / Other)
3. Goods Inward via Bag Entry
4. Gate Pass Inward (inventory control)
5. TDS Entry at bill or payment time
6. Payment Voucher (against bill, with TDS deduction)

### Technology Stack (VB6 era)
- **Frontend/UI:** VB6 Forms
- **Database:** Microsoft Access (MDB) / Oracle (via ODBC)
- **Reports:** Crystal Reports 8.5 (100+ .rpt files)
- **Communication:** SMS (Sinfini), WhatsApp (aronertech.com API)
- **QR Codes:** IDAutomation + Morovia QRCode libraries
- **Tax:** Pre-GST (CST: C/E1/E2/F Forms) + GST (CGST/SGST/IGST, IRN, RCM)

### Multi-Tenancy
- Separate MDB file per firm (company)
- Central Setup DB for cross-firm user management
- Single HITRIX installation serves multiple trading firms
