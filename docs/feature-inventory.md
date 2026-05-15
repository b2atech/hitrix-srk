# HITRIX Feature Inventory — Flat Table

> Complete list of features identified in the VB6 HITRIX system.  
> Used as input for gap analysis against DhanMan microservices.

| # | Feature | Domain | Sub-feature | Key Notes |
|---|---|---|---|---|
| 1 | Company Master | Masters | Multi-Company | Fiscal year, GSTIN/PAN/TAN/CST, bank details, TDS flag |
| 2 | Account Master (Chart of Accounts) | Masters | Ledger | Opening balance, interest %, broker linkage, mobile for WhatsApp |
| 3 | Account Group Master | Masters | Grouping | 2-level hierarchy, fixed vs dynamic, BS/P&L classification |
| 4 | Party / Weaver Group Master | Masters | Classification | Weaver/party grouping for filtering |
| 5 | Count / Unit of Measure Master | Masters | Reference | Unit definitions (metres, kg, bags, hanks) |
| 6 | Godown / Warehouse Master | Masters | Reference | Storage location definitions |
| 7 | Narration Master | Masters | Reference | Transaction narration text library |
| 8 | Tax Narration Master (Pre-GST) | Masters | Tax | Narration → tax rate + VAT account mapping |
| 9 | Tax Narration Master (GST) | Masters | Tax | Narration → CGST/SGST/IGST account mapping |
| 10 | Item / Commodity Master | Masters | Inventory | Item code, description, unit, HSN code |
| 11 | Other Items Master | Masters | Inventory | Non-standard / miscellaneous items |
| 12 | Transport Master | Masters | Reference | Transportation providers |
| 13 | Sizer / Consignee Master | Masters | Reference | Third-party warehouse handlers |
| 14 | Mill Expense Description | Masters | Reference | Expense categories for mill charges |
| 15 | System Settings — Default Accounts | Masters | Configuration | Auto-posting defaults for purchase/sale/TDS/TCS/brokerage |
| 16 | System Settings — Tax Flags | Masters | Configuration | VAT voucher generation, mill debit note config |
| 17 | Party Booking | Booking | Order | Job work order for weaver party; rate/qty/delivery date |
| 18 | Trade Purchase Booking | Booking | Order | Purchase order from trade supplier |
| 19 | SIT Purchase Booking | Booking | Order | Goods in transit at supplier's godown |
| 20 | Consignment / Depot Booking | Booking | Order | Inward booking for consignment arrangement |
| 21 | Mill Bill Booking | Booking | Order | Job work order to mill; expected completion date |
| 22 | Booking Status Tracking | Booking | Workflow | Pending / Cancelled / Completed status |
| 23 | Booking Cancellation | Booking | Workflow | Cancel with reason; cancellation report |
| 24 | Booking → Bill Conversion | Booking | Workflow | Link booking to formal inward/sales bill |
| 25 | Booking WhatsApp Notification | Booking | Communication | Confirmation PDF sent to party and broker |
| 26 | Trade Sale (SY) | Sales | Bill Entry | Direct sale to customer |
| 27 | Consignment Sale (SO) | Sales | Bill Entry | Sale from consignment stock |
| 28 | Depot Sale (SD) | Sales | Bill Entry | Sale from depot arrangement |
| 29 | SIT Sale (ST) | Sales | Bill Entry | Sale of stock-in-transit goods |
| 30 | Hank Sale (SYHNK) | Sales | Bill Entry | Yarn hank sales — textile specific |
| 31 | Mill Bill (SM) | Sales | Bill Entry | Processing fee invoice |
| 32 | GST Sales Invoice | Sales | GST | CGST/SGST/IGST, HSN, IRN, prefix numbering |
| 33 | Other GST Sales | Sales | GST | Miscellaneous GST sales (non-trade) |
| 34 | Sales Return | Sales | Return | Customer return with credit note |
| 35 | Sale Return Trade GST | Sales | Return | GST-compliant sales return |
| 36 | Sales Confirmation | Sales | Workflow | Despatch confirmation to customer |
| 37 | Late Payment Interest Debit Note | Sales | Receivables | Auto-calculate interest on overdue; raise debit note |
| 38 | Credit / Debit Note GST | Sales | Adjustment | Price adjustment notes with GST reversal |
| 39 | Invoice Auto-Numbering | Sales | System | Numeric or prefix-based invoice numbering |
| 40 | TCS on Sales | Sales | Tax | Tax Collected at Source for eligible buyers |
| 41 | Audit Lock on Bills | Sales | Audit | VIsAudited prevents modification post-audit |
| 42 | Trade Purchase (PY) | Purchase | Bill Entry | Direct purchase from supplier |
| 43 | SIT Purchase (PT) | Purchase | Bill Entry | Stock-in-transit at supplier's warehouse |
| 44 | Other Purchase (PO) | Purchase | Bill Entry | Miscellaneous / non-commodity purchases |
| 45 | Hank Purchase (PYHNK) | Purchase | Bill Entry | Yarn hank purchase — textile specific |
| 46 | GST Trade Purchase | Purchase | GST | CGST/SGST/IGST, HSN, IRN, RCM flag |
| 47 | GST Other Purchase | Purchase | GST | Non-commodity GST purchase |
| 48 | Reverse Charge Mechanism (RCM) | Purchase | GST | Company pays GST for unregistered vendors |
| 49 | Purchase Return | Purchase | Return | Return to supplier with debit note |
| 50 | Purchase Return Trade GST | Purchase | Return | GST-compliant purchase return |
| 51 | Goods Inward / Purchase Inward | Purchase | Receipt | Bag/weight tracking; links to purchase bill |
| 52 | Weight Reconciliation | Purchase | Receipt | Bill weight vs actual received weight |
| 53 | TDS on Purchase | Purchase | Tax | TDS deduction; certificate; circle; PAN capture |
| 54 | Multiple Inwards per Bill | Purchase | Workflow | Partial deliveries against one purchase bill |
| 55 | Payment Voucher (Cash/Bank) | Finance | Transaction | Settle bills; cheque details; discount at payment |
| 56 | Receipt Voucher (Cash/Bank) | Finance | Transaction | Settle sales bills; cheque/NEFT; TCS flag |
| 57 | Mill Receipt | Finance | Transaction | Receipt specific to mill transactions |
| 58 | Journal Voucher | Finance | Transaction | General ledger journal entry; 4 narration lines |
| 59 | Journal Voucher GST | Finance | Transaction | Credit/debit note for tax adjustments |
| 60 | Multi-Account Journal | Finance | Transaction | Multi-line journal for complex allocations |
| 61 | Payment-Receipt Combined View | Finance | Transaction | Unified payment/receipt management screen |
| 62 | Bill-to-Bill Settlement | Finance | Outstanding | Specific invoice selection at payment time |
| 63 | Discount at Payment | Finance | Outstanding | Cash/settlement discount entered at payment |
| 64 | TDS at Payment | Finance | Tax | TDS deducted at payment if not at bill stage |
| 65 | Bank Reconciliation | Finance | Reconciliation | Match ledger to bank statement; uncleared cheques |
| 66 | Outstanding Aging | Finance | Receivables | 0–30, 30–60, 60–90, 90+ day buckets |
| 67 | Late Payment Interest | Finance | Receivables | Monthly interest on overdue balances |
| 68 | Final Accounts — P&L | Finance | Reports | Profit & Loss statement |
| 69 | Final Accounts — Balance Sheet | Finance | Reports | Assets/Liabilities/Equity |
| 70 | Fiscal Year Management | Finance | System | FY start/end; locked years; year-end carryforward |
| 71 | Goods Inward — Bag Entry | Inventory | Receipt | Sequential bag numbers; individual weights; lot tracking |
| 72 | Lorry Receipt Tracking | Inventory | Receipt | LR number and date at inward |
| 73 | Gate Pass — Inward | Inventory | Movement | Documentary control for goods entering godown |
| 74 | Gate Pass — Outward | Inventory | Movement | Documentary control for goods leaving godown |
| 75 | Gate Pass Printing | Inventory | Output | PDF gate pass for truck/driver |
| 76 | Gate Pass WhatsApp Delivery | Inventory | Communication | Gate pass PDF sent to consignee via WhatsApp |
| 77 | GP Delivery Confirmation | Inventory | Delivery | Confirm delivery to final recipient |
| 78 | Stock Balance Tracking | Inventory | Stock | Current quantities and value per item |
| 79 | Weight Shortfall Tracking | Inventory | Quality | Loss/damage at inward stage |
| 80 | SIT Stock Tracking | Inventory | Stock | Goods in transit not yet in godown |
| 81 | CST C-Form Management | Tax | Pre-GST | C-Form required, issued, received per bill |
| 82 | CST E1-Form Management | Tax | Pre-GST | E1-Form for first in-transit sale |
| 83 | CST E2-Form Management | Tax | Pre-GST | E2-Form for second in-transit sale |
| 84 | CST F-Form Management | Tax | Pre-GST | F-Form for stock transfers |
| 85 | Form Expiry Tracking | Tax | Pre-GST | Validity check; expiry alerts |
| 86 | Form ↔ Bill Linkage | Tax | Pre-GST | Forms linked to specific bills for audit |
| 87 | GST Register (Sales) | Tax | GST | GSTR-1 support data |
| 88 | GST Register (Purchase) | Tax | GST | GSTR-2B support data |
| 89 | IRN / E-Invoicing | Tax | GST | IRN number tracking for eligible invoices |
| 90 | HSN Code at Transaction | Tax | GST | HSN captured per bill line |
| 91 | TCS Management | Tax | TCS | TCS firm flag; collection register; deductee |
| 92 | TDS Compliance | Tax | TDS | TDS entries; certificate; quarterly return data |
| 93 | Account Ledger Report | Reporting | Ledger | Party-wise debit/credit with running balance |
| 94 | Group-wise Ledger | Reporting | Ledger | Summary by account group |
| 95 | Party Confirmation Letter | Reporting | Ledger | Account confirmation for party signature |
| 96 | Party Statement | Reporting | Outstanding | Outstanding dues with bill detail |
| 97 | Broker Statement | Reporting | Outstanding | Commission and volume summary |
| 98 | Outstanding Aging Report | Reporting | Outstanding | 4-bucket aging analysis |
| 99 | Outstanding with Interest | Reporting | Outstanding | Dues with accrued late payment interest |
| 100 | Sales Summary — Party wise | Reporting | Sales | Total sales per party |
| 101 | Sales Summary — Mill wise | Reporting | Sales | Total sales per mill |
| 102 | Sales Summary — Broker wise | Reporting | Sales | Commission/volume per broker |
| 103 | Sales Summary — Count wise | Reporting | Sales | Sales by unit/count |
| 104 | Sales Summary — Month wise | Reporting | Sales | Monthly trend |
| 105 | Sales Bill Print | Reporting | Sales | Customer-facing invoice format |
| 106 | Purchase Summary — Mill wise | Reporting | Purchase | Total purchases per mill |
| 107 | Purchase Summary — Month wise | Reporting | Purchase | Monthly trend |
| 108 | Purchase vs Sales Comparison | Reporting | Analytics | Month-wise combined summary |
| 109 | Stock Balance Report | Reporting | Inventory | Current quantities and value |
| 110 | Goods Inward Reports | Reporting | Inventory | Count-wise, mill-wise, total inward summary |
| 111 | Gate Pass Register | Reporting | Logistics | All passes — inward and outward |
| 112 | P&L Statement | Reporting | Financial | Income vs expenses; gross and net profit |
| 113 | Balance Sheet | Reporting | Financial | Assets/Liabilities/Equity |
| 114 | Trial Balance | Reporting | Financial | All accounts debit/credit totals |
| 115 | Day Book / Cash Book / Bank Book | Reporting | Financial | Chronological transaction registers |
| 116 | GST Register Report | Reporting | Tax | Sales/purchase with HSN and tax amounts |
| 117 | C/E1/E2/F-Form Reports | Reporting | Tax | 30+ variants for CST form tracking |
| 118 | TDS Report | Reporting | Tax | Deduction summary and detail |
| 119 | Bar/Line/Pie Charts | Reporting | Analytics | Purchase, sales, combined visual charts |
| 120 | WhatsApp Outstanding Report | Reporting | Communication | Outstanding formatted for WhatsApp |
| 121 | SMS Outstanding Report | Reporting | Communication | Outstanding formatted for SMS |
| 122 | WhatsApp Ledger Delivery | Communication | WhatsApp | Ledger PDF sent to party via WhatsApp API |
| 123 | WhatsApp Statement Delivery | Communication | WhatsApp | Statement PDF sent via WhatsApp |
| 124 | WhatsApp Gate Pass Delivery | Communication | WhatsApp | Gate pass PDF sent to consignee |
| 125 | WhatsApp Booking Confirmation | Communication | WhatsApp | Booking slip sent to party/broker |
| 126 | SMS Outstanding Reminder | Communication | SMS | Bulk SMS for overdue accounts |
| 127 | SMS Payment Confirmation | Communication | SMS | Receipt acknowledgment |
| 128 | Check Pass SMS (Audit Alert) | Communication | Audit | SMS/WhatsApp to admin on delete/modify |
| 129 | User Management | Admin | Security | User creation; admin flag; active/inactive |
| 130 | User Privilege Matrix | Admin | Security | Per-form Add/Modify/Delete/View control |
| 131 | Password Management | Admin | Security | Change; validation; length enforcement |
| 132 | Data Backup | Admin | Data | Full DB backup to external/network path |
| 133 | Data Restore | Admin | Data | Point-in-time restore from backup |
| 134 | Audit Lock | Admin | Audit | Transaction lock post-audit; password to unlock |
| 135 | User Activity Tracking | Admin | Audit | vUserName on every transaction |
| 136 | Multi-Company Support | Infrastructure | Multi-Tenancy | Separate MDB per firm; company selection at login |
| 137 | Multi-Year Support | Infrastructure | System | Multiple fiscal years; year selection at login |
| 138 | Account Transfer Utility | Infrastructure | Migration | Transfer balances between firms/years |
| 139 | Master Copy Utility | Infrastructure | Migration | Copy masters from one company to another |
| 140 | Year-End Carryforward | Infrastructure | Migration | Carry outstanding balances to new fiscal year |
| 141 | Crystal Reports Integration | Infrastructure | Reporting | 100+ pre-built .rpt report files |
| 142 | WhatsApp API Integration | Infrastructure | Integration | External API (aronertech.com) for WhatsApp |
| 143 | SMS API Integration | Infrastructure | Integration | Sinfini alerts (legacy SMS gateway) |
| 144 | QR Code Generation | Infrastructure | Integration | IDAutomation + Morovia QRCode libraries |
| 145 | Access DB Support | Infrastructure | Database | Microsoft Access MDB per company |
| 146 | Oracle DB Support | Infrastructure | Database | Oracle via ODBC DSN (alternate backend) |
