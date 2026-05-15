# Domain 9: Communication (SMS, WhatsApp)

## Purpose
Send transactional documents, outstanding notices, and business communications to parties, brokers, customers, and suppliers via WhatsApp and SMS. Replaces manual printing and courier for routine trade communications.

---

## Key Features

### 9.1 WhatsApp Integration (`frmRptWhatsApp`)
- PDF report upload and delivery via WhatsApp API (aronertech.com)
- Reports supported for WhatsApp delivery:
  - Account Ledger
  - Party Outstanding Statement
  - Broker Outstanding Statement
  - Gate Pass (inward/outward)
  - Booking Confirmation
  - Sales Bill
- Mobile number per account (primary and secondary)
- File generated locally, uploaded to API, message sent to party

### 9.2 SMS Integration (`frmRptSMS`)
- Bulk SMS to parties (via Sinfini alerts — legacy provider)
- SMS types:
  - Outstanding balance reminders
  - Payment receipt confirmation
  - Bill delivery notification
  - Audit alert (Check Pass SMS)
- Mobile number per account master

### 9.3 Check Pass Alert (Audit Protection)
- SMS / WhatsApp alert sent to admin/owner when:
  - A transaction is modified
  - A transaction is deleted
  - A sensitive operation is performed
- Acts as a real-time notification of changes for audit trail

### 9.4 Report Delivery Flow
```
User triggers report → Crystal Reports generates PDF → 
PDF saved to HiReports_Access folder → 
WhatsApp API called with file path and mobile number → 
Message sent to party
```

---

## Key Data Captured

| Field | Description |
|---|---|
| `Mbl1`, `Mbl2` | Primary and secondary mobile numbers (from account master) |
| `WAMsg` | WhatsApp message text |
| `WAFile` | Local PDF file path for attachment |
| `SMSMsg` | SMS message content |
| `WASendStatus` | Send success / failure |
| `tmpWhatsAppErr` | Temporary table for failed message logging |

---

## Business Rules & Logic

- Mobile numbers stored in account master (`frmMastAccount`); carried to reports automatically
- API errors: Retry mechanism available; failed messages logged
- File path: Reports auto-generated in `HiReports_Access` folder
- Format validation: Mobile number format checked before sending
- WhatsApp: Uses third-party API (aronertech.com); not official WhatsApp Business API
- SMS: Legacy integration via Sinfini; character limit applies
- Check Pass SMS: Cannot be disabled by regular users (admin-only configuration)

---

## Integration Points

| Integrates With | How |
|---|---|
| Account Master | Mobile numbers sourced per party |
| Reporting | PDF reports generated before WhatsApp/SMS delivery |
| Finance | Outstanding statements triggered from payment module |
| Inventory | Gate pass PDF sent via WhatsApp after printing |
| Booking | Booking confirmations sent to party and broker |
| Administration | Check pass alerts sent on sensitive operations |
