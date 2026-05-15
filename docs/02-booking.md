# Domain 2: Booking & Order Management

## Purpose
Capture advance orders and bookings before goods are physically received or sold. Acts as the pre-transaction layer — no ledger posting occurs at booking stage. Bookings are tracked for follow-up, cancellation, and conversion to formal inward/sales bills.

---

## Key Features

### 2.1 Party Booking (`frmBookingParty`, `frmBookingParty New`)
- Book goods for weaver parties (job work orders)
- Rate per piece/lot basis
- Expected delivery date tracking
- Broker linkage
- Status tracking: Pending → Completed / Cancelled
- WhatsApp confirmation to party and broker

### 2.2 Trade Purchase Booking (`frmBookingTrade`)
- Purchase order from trade suppliers
- Booking rate and quantity
- Supplier and broker details
- Links to formal purchase inward when received

### 2.3 SIT Purchase Booking (`frmBookingSIT`)
- Book goods currently in Stock-In-Transit at supplier's warehouse
- Goods not yet physically received by company
- Pending until transfer of ownership confirmed

### 2.4 Consignment / Depot Booking (`frmBookingInward`)
- Book inward goods for consignment or depot arrangements
- Third-party warehouse (sizer/consignee) linkage
- Tracks goods before physical receipt into company godown

### 2.5 Mill Bill Booking (`frmBookingMillbill`)
- Book mill processing jobs (job work orders to mills)
- Expected completion date
- Mill name and job description
- Processing fee estimate
- Converts to Mill Bill (SM) transaction on completion

### 2.6 Sales Confirmation (`frmConfirmationSales`)
- Confirmation of goods despatched to customer
- Party and delivery details
- Link to original sales bill

---

## Key Data Captured

| Field | Description |
|---|---|
| `BookVNo`, `BookVDt` | Booking voucher number and date |
| `BookPartyName` | Customer / weaver party |
| `BrokerName` | Commission agent |
| `MillName` | Mill / processor |
| `BookQty`, `BookWeight` | Booked quantity and weight |
| `BookRate`, `BookAmt` | Rate and total booked amount |
| `DeliveryDt` | Expected delivery date |
| `BookStatus` | Pending / Cancelled / Completed |
| `BookType` | SIT / Trade / Consignment / Party / Mill |
| `Remarks` | Free-text notes on booking |

---

## Business Rules & Logic

- Bookings do not post to General Ledger — they are pre-transaction tracking records
- Outstanding booking tracking: Booked quantity vs actual despatched/received is reconciled
- Party booking for job work: Rate per piece/lot (not per weight) is common in textile trade
- Mill bill booking: Tracks expected processing fees; links to actual bill on completion
- Auto-number generation: Booking number generated per booking type
- WhatsApp notifications sent to brokers including party name and booking details
- Cancellation: Booking can be cancelled with reason; cancellation report generated
- Conversion: Bookings are "converted" to inward/sales bills; booking reference carried forward
- Status check: System prevents duplicate booking for same party/date/item combination

---

## Reports Available

| Report | Description |
|---|---|
| `rptBookingPartyPending.rpt` | Party bookings pending — not yet despatched |
| `rptBookingPartyVsDispatch.rpt` | Booked quantity vs actual despatch |
| `rptBookingPartyCancel.rpt` | Cancelled bookings with reasons |
| `rptBookingPartyReg.rpt` | Booking register — all bookings |
| `rptBookingPartyPrint.rpt` | Booking slip printout (customer-facing) |
| `rptInwardSummCountwise.rpt` | Inward summary — count/unit wise |
| `rptInwardSummMillwise.rpt` | Inward summary — mill wise |
| `rptInwardSummTotal.rpt` | Inward summary — total |

---

## Integration Points

| Integrates With | How |
|---|---|
| Sales | Party booking converts to sales bill; booking ref carried forward |
| Purchase | Trade/SIT booking links to purchase inward on receipt |
| Inventory | Mill bill booking links to goods inward on completion |
| Communication | WhatsApp API sends booking confirmation to broker/party |
| Outstanding | Booking references tracked for follow-up in outstanding module |
| Reporting | Booking register and pending reports generated from booking records |
