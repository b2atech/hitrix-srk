# Domain 8: Reporting & Analytics

## Purpose
Provide comprehensive operational, compliance, and management reports covering all business domains. Reports are generated via Crystal Reports 8.5 and delivered as printouts, PDFs, or shared via WhatsApp/SMS.

---

## Key Features

### 8.1 Account / Ledger Reports (`frmRptAccount`, `frmRptPartyLdr`)
- **Account Ledger:** Party-wise debit/credit detail with running balance
- **Group-wise Ledger:** Summary by account group
- **Party Confirmation:** Account confirmation letter with balance for customer signature
- **Account Balance List:** Trial balance style listing
- **Multi-Party Ledger:** Multiple parties in one report

### 8.2 Statement Reports (`frmRptStatement`)
- **Party Statement:** Outstanding dues by date range with bill-level detail
- **Broker Statement:** Commission-wise summary
- **Mill Statement:** Processing fees and balances per mill
- **Outstanding Aging:** 0–30, 30–60, 60–90, 90+ day buckets

### 8.3 Financial Reports (`frmRptFinal`)
- **Profit & Loss Statement:** Income vs expenses; gross and net profit
- **Balance Sheet:** Assets/Liabilities/Equity
- **Trial Balance:** All accounts with debit/credit totals
- Configurable for any date range within fiscal year

### 8.4 Register Reports (`frmRptRegi`)
- **Sales Register:** All sales bills in date order with amounts, tax, party, broker
- **Purchase Register:** All purchase bills with similar detail
- **Receipt Register:** All receipts with cheque details
- **Payment Register:** All payments with cheque details
- **JV Register:** Journal voucher listing
- **TDS Register:** TDS deductions with certificate details
- **Bank Book:** Bank account ledger (day-wise)
- **Cash Book:** Cash account ledger (day-wise)
- **Day Book:** All transactions in a date range across all accounts

### 8.5 Tax Reports (`frmRptTaxReports`)
- CST Form registers: C-Form, E1, E2, F-Form (30+ report variants)
- GST Register: Sales / Purchase (frmRptGST)
- TDS Summary and Detail
- TCS Collection Report

### 8.6 Sales Analytics Reports (`frmRptListView`, Crystal Reports)
- **Sales Summary — Party wise:** Total sales per party
- **Sales Summary — Mill wise:** Total sales per mill
- **Sales Summary — Broker wise:** Commission and volume per broker
- **Sales Summary — Count wise:** Sales by unit/count (common in yarn trade)
- **Sales Summary — Month wise:** Monthly trend
- **Purchase vs Sales Comparison:** Month-wise trend

### 8.7 Stock / Inventory Reports (`frmRptStockReports`)
- **Stock Balance by Item:** Current quantity and value per item
- **Goods in Transit (SIT):** Goods not yet received in godown
- **Bag Inventory:** Bag-level tracking
- **Delivery Summary:** Confirmed deliveries
- **Inward Summary:** Goods received — mill wise, count wise, total

### 8.8 Graph / Chart Reports (`frmRptGraph`)
- **Bar Charts:** Purchase, Sales, combined
- **Line Charts:** Purchase, Sales, combined trends
- **Pie Charts:** Purchase, Sales distribution

| Report | Description |
|---|---|
| `rptGraphPurchaseBar.rpt` | Purchase — bar chart |
| `rptGraphPurchaseLine.rpt` | Purchase — line chart |
| `rptGraphPurchasePie.rpt` | Purchase — pie chart |
| `rptGraphSalesBar.rpt` | Sales — bar chart |
| `rptGraphSalesLine.rpt` | Sales — line chart |
| `rptGraphSalesPie.rpt` | Sales — pie chart |
| `rptGraphPurVsSaleBar.rpt` | Purchase vs Sales — bar |
| `rptGraphPurVsSaleLine.rpt` | Purchase vs Sales — line |
| `rptGraphPurVsSalePie.rpt` | Purchase vs Sales — pie |

### 8.9 Outstanding Reports
- **Party-wise Outstanding:** Unpaid bills per party with age
- **Broker-wise Outstanding:** Outstanding grouped by broker
- **Outstanding with Interest:** Bills with accrued late payment interest
- **WhatsApp format:** Compact outstanding for messaging

### 8.10 SMS / WhatsApp Reports (`frmRptSMS`, `frmRptWhatsApp`)
- Party outstanding formatted for WhatsApp delivery
- Broker outstanding formatted for WhatsApp
- SMS outstanding list
- Message delivery status log

---

## Report Delivery Mechanisms

| Mechanism | Description |
|---|---|
| Print | Direct printer output |
| PDF Export | Crystal Reports export to PDF |
| WhatsApp | PDF uploaded to API, link sent to party mobile |
| SMS | Text summary sent via SMS API |
| Screen Preview | On-screen report viewer (crviewer) |

---

## Report Parameter Patterns

Most reports accept common parameters:
- **Date Range:** From date, To date
- **Company Filter:** CompCode (multi-company support)
- **Party Filter:** Single party or all parties
- **Broker Filter:** Single broker or all
- **Mill Filter:** Single mill or all
- **Account Group Filter:** Group-wise filtering
- **Voucher Type Filter:** SY/SO/SD/ST/SM/PY/PT/PO etc.
- **Fiscal Year:** CompYear (e.g., "2526")

---

## Crystal Reports List (complete)

### Booking Reports
- `rptBookingPartyCancel.rpt` — Cancelled bookings
- `rptBookingPartyPending.rpt` — Pending bookings
- `rptBookingPartyPrint.rpt` — Booking slip print
- `rptBookingPartyReg.rpt` — Booking register
- `rptBookingPartyVsDispatch.rpt` — Booked vs despatched

### Inward Reports
- `rptInwardSummCountwise.rpt`
- `rptInwardSummMillwise.rpt`
- `rptInwardSummTotal.rpt`

### Financial / Cash Reports
- `rptCashBankStatus.rpt`
- `rptLatePayDbNtReg.rpt`

### Outstanding Reports
- `rptOutStandingAgingBank.rpt`
- `rptOutStandingSALatePayInt.rpt`
- `rptOutStandingSALatePayIntPurch.rpt`
- `rptOutstBrkPartyWithUnadjustRec.rpt`
- `rptOutstBrokerPartywise.rpt`
- `rptOutstBrokerwise.rpt`
- `rptOutstBrokerwiseWhatsApp.rpt`
- `rptOutstGST.rpt`
- `rptOutstPartywiseWhatsApp.rpt`
- `rptOutstPurchPartywise.rpt`
- `rptOutstSALatePayIntOS.rpt`

### Purchase Reports
- `rptPurchaseSalesSummMonthwise.rpt`
- `rptPurchaseSummCountwise.rpt`
- `rptPurchaseSummMillwise.rpt`
- `rptPurchaseSummMonthwise.rpt`
- `rptPurchaseSummTotal.rpt`
- `rptPurchSummMonthwiseNew.rpt`

### Sales Reports
- `rptSaleBillPrint.rpt`
- `rptSalesSummBrokerCountwise.rpt`
- `rptSalesSummBrokerPartywise.rpt`
- `rptSalesSummCountwise.rpt`
- `rptSalesSummMillCountwise.rpt`
- `rptSalesSummMillwise.rpt`
- `rptSalesSummMonthwise.rpt`
- `rptSalesSummPartyCountwise.rpt`
- `rptSalesSummPartywise.rpt`
- `rptSalesSummTotal.rpt`

### SMS / WhatsApp
- `rptSMSDetails.rpt`
- `rptSMSOsBroker.rpt`
- `rptSMSOsParty.rpt`

### Tax Reports — CST (30+ reports across C/E1/E2/F Form types)
*(see Domain 7 — Tax Compliance for complete list)*

---

## Integration Points

| Integrates With | How |
|---|---|
| All Transaction Modules | Report data sourced from GL and transaction tables |
| Finance | Ledger, trial balance, final accounts from GL |
| Sales / Purchase | Sales and purchase summaries from bill tables |
| Tax Compliance | Tax registers from GST/CST fields in transactions |
| Communication | PDF reports delivered via WhatsApp API and SMS |
| Crystal Reports Engine | 100+ pre-built .rpt files for all report types |
