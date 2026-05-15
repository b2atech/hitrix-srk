# HITRIX — Booking & Order Management: Domain Analysis

**Source version**: ABabu_Cd_15-05-2026  
**Analysis date**: 2026-05-15  
**Forms analysed**: frmBookingTrade, frmBookingTradeGST, frmBookingTradePurch, frmBookingMillBill, frmBookingMillBillDespatch, frmBookingMillBillDespatchNew, frmBookingParty, frmBookingInward, frmBookingSIT

---

## 1. Overview

Booking in HITRIX represents the **purchase order / indent** stage of a textile trade transaction. The firm places a booking with a mill on behalf of a buyer (party). This creates a binding order record that later flows into despatch and billing. There are multiple booking variants:

| Form | VType | Purpose |
|------|-------|---------|
| frmBookingTradeGST | BK (set from gfrmTypeStr) | GST-era booking with multi-item despatch schedule |
| frmBookingTrade | BT (legacy) | Pre-GST booking via tblPurchSub (non-booking table — see note) |
| frmBookingTradePurch | set externally | Booking linked to purchase entry |
| frmBookingMillBill | BM / similar | Mill bill booking |
| frmBookingMillBillDespatch | BD | Mill bill despatch records |
| frmBookingMillBillDespatchNew | BD | New-style mill bill despatch |
| frmBookingParty | set externally | Party-side booking |
| frmBookingInward | set externally | Inward (receipt at godown) linked to booking |
| frmBookingSIT | set externally | SIT (Stock-in-Transit) booking |

**Note**: `frmBookingTrade` actually writes to `tblPurchSub` (not `tblBooking`). Its `frmTypeStr` is inherited from `gfrmTypeStr`. `frmBookingTradeGST` writes to `tblBooking`.

---

## 2. Primary Database Tables

### tblBooking (primary booking header + line-level rows)
Each row in `tblBooking` is **one despatch schedule line** (not one booking header). The booking header information is repeated on every row; `BkSrNo` distinguishes rows.

| Column | Type | Description |
|--------|------|-------------|
| VNo | int PK | Booking number (auto-generated per firm/year/type) |
| VType | nvarchar(2) PK | Voucher type code (e.g. BK, BT, BM, BD …) |
| Vdt | smalldatetime | Booking date |
| VYear | int PK | Financial year |
| VFirm | nvarchar(4) PK | Firm code (multi-company key) |
| BkParty | int | Party account code (FK tblMastAccount) |
| BkMillCode | int | Mill account code (FK tblMastAccount, AgCode=90029) |
| BkBroker | int | Broker account code (FK tblMastAccount, AgCode=90016) |
| BkRefNo | nvarchar(20) | Mill's reference / order number |
| BkRefDt | smalldatetime | Mill's reference date |
| BkSubAmt | numeric(18,2) | Sub-total amount |
| BkBillAmt | numeric(18,2) | Bill amount |
| BkRoff | numeric(18,2) | Rounding off |
| DueDays | int | Credit days |
| DueDate | smalldatetime | Due date |
| BkEntDt | smalldatetime | Entry date |
| BkUser | nvarchar(20) | User who entered |
| BkItCode | int | Item (count/quality) code (FK tblMastItem) |
| BkBag | int | Total bags booked |
| BkWt | numeric(18,3) | Total weight (kg) |
| BkRt | numeric(18,2) | Rate |
| BkRtPer | numeric(18,2) | Rate per (kg / 100kg etc.) |
| BkRateType | int | Rate type code |
| BkDeleAdd | int | Delivery address code (FK tblMastDeleAdd) |
| BkPartyGroup | int | Company group code (FK tblMastCompGroup) |
| BkExtraCount | nvarchar(50) | Additional item codes text |
| BkNote | nvarchar(100) | Notes |
| BkCbag | int | Cancelled bags |
| BkCWt | numeric(18,3) | Cancelled weight |
| BkCDate | smalldatetime | Cancellation date |
| BkSrNo | int PK | Line serial number (schedule row) |
| BkShDate | smalldatetime | Scheduled despatch date |
| BkShBag | int | Scheduled bags for this line |
| BkShWt | numeric(18,3) | Scheduled weight for this line |
| BkTmpNo1-3 | numeric(18,2) | Temp numeric fields |
| BkTmpChr1-3 | nvarchar | Temp character fields (BkTmpChr1/2 = narrations; BkTmpChr3 max 25) |
| BkBagLess | int | Bags deducted |
| BkIsDirectPayment | int | Flag: direct payment to mill |
| BkIsFrghtInsu | int | Flag: freight/insurance included |
| BkIsInsu | int | Flag: insurance |
| BkMailSendDt | nvarchar(30) | Mail sent date |
| BkmailSendPartryDt | nvarchar(30) | Mail sent to party date |

### tblBookingSub (alternate sub-table)
Used by some booking forms for line details.

| Column | Description |
|--------|-------------|
| VNo,VType,VYear,VFirm,BkSubItSrNo | PK |
| BkSubBag, BkSubWt, BkSubDDate | Despatch details |
| BkSubItCode | Item code |
| BkSubDespDt, BkSubDespBag, BkSubDespWt | Actual despatch |
| BkSubTrnpt | Transport narration code |
| BkSubDespVno, BkSubDespVyear | Link back to despatch voucher |

### tblBookingDesp (despatch tracking sub-table)
| Column | Description |
|--------|-------------|
| VNo,VType,VYear,VFirm,BkSrNo | PK |
| BkDespItSrNo | Item serial |
| BkDespBag, BkDespWt, BkDespDate | Despatch quantities |
| BkDespItCode | Item |
| BkDespShDate | Scheduled date |
| BkDespTmpNo1-3, BkDespTmpChr1-3 | Spare fields |

### tblPurchSub (used by legacy frmBookingTrade)
PK: VNo, VType, VYear, VFirm, PurSubItSrNo

| Column | Description |
|--------|-------------|
| PursubBookNo | Booking reference number |
| PurSubBookDt | Booking date |
| PurSubItCode | Item code |
| PurSubBag | Bags |
| PurSubWt | Weight |
| PurSubRt | Rate |
| PurSubRtPer | Rate per |
| PurSubAmt | Amount |
| PurSubLotNo | Lot number |
| PurSubGodown | Godown narration code |
| PurSubItCtrlNo | Item control number (bigint) |
| PurSubBkVyear, PurSubBkSrNo | Booking year/serial link |

---

## 3. Field-level Inventory (frmBookingTradeGST — the current GST form)

### Header Fields (single-entry areas)

| Control | Index | Field | Purpose | Mandatory |
|---------|-------|-------|---------|-----------|
| txt(0) | 0 | VNo | Booking number (auto-generated) | Yes |
| txt(1) | 1 | VType | Voucher type (set from gfrmTypeStr) | Yes (system) |
| dtpDt(2) | 2 | Vdt | Booking date | Yes |
| txt(3) | 3 | VYear | Financial year (system) | Yes (system) |
| txtCode(6)/txttemp(6) | 6 | BkMillCode | Mill code / name | Yes |
| txtCode(5)/txttemp(5) | 5 | BkParty | Party (buyer) name | Yes |
| txtCode(7)/txttemp(7) | 7 | BkBroker | Broker name | No |
| txttemp(17) | 17 | BkItCode | Item/count name (mill-specific) | Yes |
| txtStdPack | — | — | Standard packing (from tblMastItem.ItStdpack) | Auto-filled |
| txt(8) | 8 | BkRefNo | Mill reference number | No |
| dtpDt(29)/txt(29) | 29 | BkCDate | Cancellation date | No |
| txt(27) | 27 | BkCbag | Cancelled bags | No |
| txt(28) | 28 | BkCWt | Cancelled weight | No |
| txt(18) | 18 | BkBag | Total bags booked | Yes |
| txt(19) | 19 | BkWt | Total weight | Auto (Bag × StdPack) |
| txt(20) | 20 | BkRt | Rate | Yes |
| txt(21) | 21 | BkRtPer | Rate per unit (from master) | Auto-filled |
| Cbo(22) | 22 | BkRateType | Rate type combo | No |
| txt(23)/txttemp(23) | 23 | BkDeleAdd | Delivery address | No |
| txttemp(24) | 24 | BkPartyGroup | Company group | No |
| txt(25) | 25 | BkExtraCount | Extra count text | No |
| txt(26) | 26 | BkNote | Note / narration | No |
| txt(37) | 37 | BkTmpChr1 | Narration line 1 | No |
| txt(38) | 38 | BkTmpChr2 | Narration line 2 | No |
| lblTinNo | — | AcGSTIN | Party's GSTIN (display only) | — |
| lblPan | — | AcPAN | Party's PAN (display only) | — |
| lblMblNo | — | AcContNo | Party's mobile (display only) | — |
| txtCrBalance | — | — | Party's account balance at booking date (display) | — |

### Grid Fields (despatch schedule — tblBooking rows)

Each grid row = one despatch schedule entry (BkSrNo = row index).

| Grid Col | Field | Description |
|----------|-------|-------------|
| 0 | BkSrNo | Auto-numbered serial |
| 1 | BkShDate | Scheduled despatch date |
| 2 | BkShBag | Bags for this schedule |
| 3 | BkShWt | Weight for this schedule |

---

## 4. Business Rules & Validations

### Trigger
- User opens form from menu (gfrmTypeStr is pre-set by the calling menu item).
- Form_Load reads `gfrmTypeStr` into `frmTypeStr` and opens an empty recordset for the current VType/VFirm/VYear.

### Auto-generation
- **VNo** is auto-generated: `GProcGenerateId("TblBooking","Vno","Vtype=... and VFirm=... and VYear=...")`.
- **Vdt** defaults to last entry date (mLastEntryDate) or today.
- **Rate Per** is auto-filled from `tblMastItem.ItStdrateper` when item is selected.
- **Std Packing** is auto-filled from `tblMastItem.ItStdpack`.
- **Broker** is auto-filled from party master (`tblMastAccount.AcBrkCode`) if not already set.

### Key Validations (frmBookingTradeGST)
1. Mill must be registered in `tblMastBillSerial` for the applicable SaleType. If not found → "Mill Not in Firm" error and field cleared.
2. Party must exist in `tblMastAccount` with AgCode in (90015, 90029).
3. Item must belong to the selected mill (`tblMastItem.ItMillCode = MillCode`).
4. Bags, weight, and rate must be > 0 in grid rows.
5. Godown must exist in `tblMastNarration` (NarrType='G').
6. Duplicate booking number check enforced on save.
7. If party's AgCode = 90029 (mill group), mill field is auto-set to party and locked.
8. Cancel date/bags only enabled via Case 5 ("Cancel Bags") button.

### Amount Calculations (CalGridAmt)
```
If Weight = 0 Then Weight = Bags × StdPack
Amount = Weight × Rate / RatePer
```
Total amount = sum of all grid row amounts.

### Date Constraint (GST form)
- Records only allowed from 01/07/2017 onwards (`vdt >= CONVERT(datetime, '01.07.2017', 104)`).

---

## 5. Business Flow

### New Booking Entry
1. User presses **Add** (Case 6): clears form, generates new VNo, sets default date.
2. User enters Mill code (txtCode(6)) or searches by name (txttemp(6)).
   - System validates mill is in `tblMastBillSerial` for firm.
   - Clears item/rate if mill changes.
3. User enters Party (txttemp(5) with full account search dialog).
   - System displays GSTIN, PAN, mobile.
   - Auto-fills broker from party master.
4. User selects Item (txttemp(17)) — filtered by mill.
   - System auto-fills StdPack and RatePer.
5. User enters total bags/weight/rate (header fields).
6. User enters despatch schedule in grid: date, bags, weight per schedule line.
   - Weight auto-calculated as Bags × StdPack if left 0.
7. User enters optional fields: reference number, note, narrations.
8. User presses **Save** (Case 13):
   - ValidateData runs.
   - CalAmount runs.
   - FillTxtFromTemp copies temp field values.
   - SaveRelatedRecord runs (see below).

### Save Process (SaveRelatedRecord)
1. Lookup account codes: PartyCode, BrokerCode, MillCode, ItemCode from name fields.
2. Delete existing records: `DELETE FROM tblBooking WHERE vtype=... AND vno=... AND VFirm=... AND VYear=...`
3. Open empty recordset on tblBooking.
4. For each grid row (i=1 to rows-1):
   - `rstSub.AddNew`
   - Write all header fields (repeated on every row).
   - Write schedule-specific fields: BkShDate, BkShBag, BkShWt from grid columns.
   - Write BkSrNo = i.
   - `rstSub.Update`
5. Close recordset.
6. Remember last entry values (mill, item, rate, date) for next entry default.

### Delete Process (DeleteRelatedRecord)
- `DELETE FROM tblBooking WHERE vtype=... AND vno=... AND VFirm=... AND VYear=...`
- No cascade to tblOutStanding or tblVoucher (booking does not create financial entries).

### Print Process (cmdPrintBill_Click)
1. User sets From/To booking number range.
2. System deletes user's TmpBooking rows: `DELETE FROM TmpBooking WHERE BkUser=...`
3. Inserts into TmpBooking: `INSERT INTO TmpBooking ... SELECT ... FROM tblBooking WHERE vno IN (...)`
4. Opens Crystal Report: `rptBookingPrint_Trade.rpt`
   - Formulas: wCSTNo, wBSTNo
   - Selection: `{TmpBooking.BkUser} = '...'`
5. Print / Screen / Mail options.
6. Mail option: commented-out email code references `SendEmail` (not currently active in form).

---

## 6. Integration Points

### Crystal Reports
| Report | Trigger |
|--------|---------|
| rptBookingPrint_Trade.rpt | cmdPrintBill — Print/Screen/Mail button |

### Linked Forms (F3 create)
| Field | Opens |
|-------|-------|
| Party (txttemp(5)) | frmMastAccount (AgCode=90015) |
| Mill (txttemp(6)) | frmMastAccount (AgCode=90029) |
| Broker (txttemp(7)) | frmMastAccount (AgCode=90016) |
| Item (txttemp(17)) | frmMastItem |
| Party Group (txttemp(24)) | frmMastCompGroup |
| Delivery Address (txttemp(23)) | frmMastDeleAdd |
| Narration txt(37/38) | frmMastNarration (NarrType='N') |

### Notification
- Email to supplier/buyer: **commented out** (dead code in cmdPrintBill).
- WhatsApp/SMS: not present in booking forms.
- IRN (e-invoice): not applicable to bookings.

---

## 7. VType Codes — Booking Domain

| VType | Description | Form |
|-------|-------------|------|
| BK | Trade Booking (GST) | frmBookingTradeGST |
| BT | Trade Booking (legacy, writes tblPurchSub) | frmBookingTrade |
| BM | Mill Bill Booking | frmBookingMillBill |
| BD | Mill Bill Despatch Booking | frmBookingMillBillDespatch / New |
| BP | Booking Party | frmBookingParty |
| BI | Booking Inward | frmBookingInward |
| BS | Booking SIT | frmBookingSIT |

*Note*: The actual VType values are loaded from `gfrmTypeStr` which is set by the calling menu item, so exact codes depend on menu configuration. The above are the conventionally observed values.

---

## 8. Multi-Company / Multi-Year Logic

- Every query is scoped by `VFirm = gCCode AND VYear = gCYear`.
- `gCCode` = current company code (4-char, from tblMastCompany.CCode).
- `gCYear` = current financial year integer.
- On Form_Load, `frmTypeStr = gfrmTypeStr` — the calling menu sets the VType before opening the form.
- `mShowRec` global: if non-empty, the form loads a specific record (used when drilling from another form).

---

## 9. Key Design Observations for Migration

1. **No separate header/detail split in tblBooking**: header data is denormalised onto every schedule row. Migration should normalise into BookingHeader + BookingSchedule.
2. **No financial postings from booking**: tblOutStanding and tblVoucher are NOT written. Booking is purely a commitment register.
3. **BkSrNo is the row counter** within a booking; the despatch schedule rows are just `BkSrNo=1,2,3...` with the same VNo.
4. **Mill validation at entry**: `tblMastBillSerial` links mill to firm and sale type. DhanMan needs an equivalent "mill-firm assignment" concept.
5. **Cancellation is in-line**: cancelled bags/weight/date are columns on the same table, not a separate cancellation document.
6. **Print uses a temp table** (TmpBooking) as a staging area for Crystal Reports — not a concern for DhanMan (reports will query directly).
7. **Last-entry memory**: form remembers last mill/item/rate for the next new entry. This UX pattern should be preserved.

---

## 10. frmBookingTrade (Legacy Pre-GST Form)

This form is **different** from frmBookingTradeGST:
- `tblName = "tblPurchSub"` — writes to the purchase sub-table, not tblBooking.
- Header written to tblPurchSub with `PursubBookNo`, `PurSubBookDt`, item, bag, weight, rate, godown.
- Grid columns: Sr | Bk.No | Book Date | Count | Godown | Lot No | Bag | Weight | Rate | Per | Amount
- Amount formula: `Amount = Weight × Rate / RatePer`
- Delete: `DELETE FROM tblPurchSub WHERE vtype=... vno=... VFirm=... VYear=...`
- No outstanding creation, no voucher creation.
- Tax type selected via `Cbo(22)` (TaxType from tblMastNarration).
- Purchase account (txttemp(6)) is separate from party (txttemp(5)) and mill (txttemp(7)).
- Exempt mill logic: if `AcIsExemptMill=1`, shows extra exempt amount and kg fields (txt(19), txt(20)).

---

## 11. frmBookingMillBill / frmBookingMillBillDespatch

These forms handle the mill-side billing lifecycle:

### frmBookingMillBill
- Records mill's invoice details at the booking stage (mill sends a bill before goods move).
- Likely VType: SM or BM (confirmed from gTaxCodeSM reference in settings).

### frmBookingMillBillDespatch / frmBookingMillBillDespatchNew
- Records actual despatch against a mill bill booking.
- Links back to the booking via `BkDespVno` / `BkSubDespVno`.
- Writes to `tblBooKingDesp` or `tblBooKingSub`.

### tblBooKingDesp fields of note:
- BkDespItSrNo, BkDespBag, BkDespWt, BkDespDate (actual despatch)
- BkDespShDate (scheduled date — reference)
- BkSrNo (links to tblBooking.BkSrNo)

---

## 12. Stored Procedures Related to Booking

From stored_procedures.sql (signatures only, bodies truncated):
- `PrcPrepareBookingPartyVsDesp` — @VFirm, @Fr_dt, @To_dt, @BkNo, @BkYear, @BkSrNo, @Accode → generates booking vs despatch comparison report.
- `PrcPrepareDespatchDetail` — @VFirm, @Fr_dt, @To_dt, @Accode, @Accode_Br, @MillCode, @ItemCode, @Vt… → despatch detail report.
- `PrcPrepareDailyEntryMillBill` — @VFirm, @Fr_dt, @To_dt, @Accode, @Accode_Br, @MillCode, @ItemCode → daily mill bill entry report.
- `PrcPrepareDailyInward` — @VFirm, @Fr_dt, @To_dt, @MillCode, @ItemCode, @Ucode → daily inward report.

---

## 13. frmBookingInward

Handles receipt of goods against a booking at the godown (inward at firm's premises).

Key fields (inferred from `tblBags` schema and bag entry patterns):
- `InVNo, InVType, InVdt, InVYear, InVFirm` — inward voucher identity
- `InItCode` — item received
- `InItSrNo` — serial within inward
- `InBag` — bags received
- `InWt` — weight
- `InGodown` — godown
- `InLotNo` — lot number
- `BagNo` — bag number (barcode / tag)
- `Cartoon`, `CartoonWt` — carton details
- `SlVNo, SlVType, SlVdt, SlVYear, SlVFirm` — linked sale voucher
- `SlItSrNo, SlBillNo` — sale line link

The inward (booking inward / `BI` VType) writes to `tblBags` tracking individual bag-level inventory.

---

## 14. frmBookingSIT (Stock In Transit)

SIT bookings record goods that are in transit — booked but not yet received at the destination godown. Uses VType likely `BS` or `ST`. The form tracks:
- Source (mill/party)
- Transit details (transport, LR number, date)
- Item details (bag, weight, rate)
- Expected arrival date

---

## 15. Summary: Booking Lifecycle

```
Order Placed (frmBookingTradeGST / frmBookingTrade)
        ↓
   tblBooking / tblPurchSub
        ↓
Goods Dispatched (frmBookingMillBillDespatchNew)
        ↓
   tblBooKingSub / tblBooKingDesp
        ↓
Goods Received Inward (frmBookingInward)
        ↓
   tblBags (bag-level tracking)
        ↓
Mill Bill Raised (frmBookingMillBill)
        ↓
   tblBooking with SM-type VType
        ↓
Purchase Invoice (frmPurchaseTradeGST)
        ↓
   tblPurch + tblPurchSub + tblOutStanding + tblVoucher
```
