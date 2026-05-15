# Finance & GST Domain Analysis — HITRIX VB6 ERP

**Source path:** `src/ABabu_Cd_15-05-2026/`
**Primary forms analysed:** frmPayment.frm, frmReceipt.frm, frmJV.frm, frmLatePaymentIntGST.frm, frmCrnDrnGST.frm, frmBankReConcilEntry.frm, frmOsAdjustment.frm, frmTDSEntry.frm, frmIntCalDate.frm
**Supporting infrastructure:** CommonPro.bas (GProcCreateVoucher, GProcCreateOutStanding, GProcGetSettingDetail, GProcGetGSTAccounts), Account.bas (ledger, P&L, balance sheet)
**Primary tables:** tblVoucher, tblOutStanding, tblIntSale, tblIntDbNote, tblInterestCalDate, TblMillRecPay, tblMastSetting, tblMastNarration (GST record)

---

## 1. Overview

The finance domain is the hub through which all monetary flows are recorded. It has five main subsystems:

1. **Receipts & Payments** (frmReceipt / frmPayment) — Cash, Bank, Mill receipt/payment vouchers, each a multi-line entry in `tblVoucher`
2. **Journal Vouchers** (frmJV / frmJVwithTDS) — General debit/credit entries, TDS reverse charges
3. **Credit / Debit Notes with GST** (frmCrnDrnGST) — Customer/supplier adjustments stored in `tblIntSale`
4. **Late Payment Interest** (frmLatePaymentIntGST) — Calculates and posts interest on overdue sale bills; creates JV/debit note entries and GST liability entries
5. **Bank Reconciliation** (frmBankReConcilEntry) — Marks payment/receipt vouchers as bank-reconciled

The critical design principle: **all financial entries flow through `tblVoucher`**. Each Vno can have multiple `VCtrNo` (counter number = line number within the voucher). Receipt/payment vouchers can have one header row (the bank entry) plus N detail rows (one per bill being settled).

The `tblOutStanding` table tracks unpaid balances per bill/party separately. `GProcCreateOutStanding` (in CommonPro.bas) creates records in it when sales/purchases are posted. The receipt/payment forms update `tblSale.SlBillRecdAmt` and `tblPurch.PurBillPaidAmt` to track running paid/received totals.

---

## 2. VType Code Inventory — Finance

| VType | Description | Form | Table |
|-------|-------------|------|-------|
| BR | Bank Receipt | frmReceipt | tblVoucher |
| CR | Cash Receipt | frmReceipt | tblVoucher |
| MR | Mill Receipt (direct mill payment) | frmReceipt | tblVoucher |
| BP | Bank Payment | frmPayment | tblVoucher |
| CP | Cash Payment | frmPayment | tblVoucher |
| JV | Journal Voucher | frmJV | tblVoucher |
| SV | Sale Credit Note (with GST) | frmCrnDrnGST | tblIntSale |
| SI | Sale Debit Note / Interest Debit Note (with GST) | frmCrnDrnGST | tblIntSale |
| PV | Purchase Credit Note (with GST) | frmCrnDrnGST | tblIntSale |
| PX | Purchase Debit Note (with GST) | frmCrnDrnGST | tblIntSale |
| MI | Mill Interest Debit Note (late payment) | frmLatePaymentIntGST | tblIntSale |

**Note:** `0P` is a legacy VType seen in TCS accumulation queries (older advance payment tracking). `PN` is Promissory Note (seen in commented JV discount entries).

---

## 3. Primary Tables

### 3.1 tblVoucher

PK: `VNo + VType + VYear + VFirm + VCtrNo`

| Column | Type | Description |
|--------|------|-------------|
| VNo | int | Voucher number (monthly sequence) |
| VType | nvarchar(2) | BR/CR/MR/BP/CP/JV |
| Vdt | smalldatetime | Voucher date |
| VYear | int | Financial year |
| VFirm | nvarchar(4) | Company code |
| VCtrNo | int | Line counter within voucher (1..n) |
| VCrAcCode | int | Credit account code |
| VDrAcCode | int | Debit account code |
| VAmt | numeric(18,2) | Amount for this line |
| VRefTp | nvarchar(10) | Reference type (Cheque/RTGS/Transfer/D.D./Other) |
| VRefNo | nvarchar(30) | Reference number (cheque no, RTGS ref, etc.) |
| VRefDate | smalldatetime | Reference date (cheque date) |
| VRefBank | nvarchar(50) | Bank name of reference instrument |
| VReconDt | smalldatetime | Bank reconciliation date |
| VNar1–4 | nvarchar(50) | Narration lines (4 fields) |
| VIsRefEntType | int | 1 = reference entry type flag |
| VIsAudited | int | 1 = audited (read-only) |
| VBillNo | nvarchar(20) | Bill number being settled |
| VBillVno | int | Bill Vno being settled |
| VBillType | nvarchar(2) | Bill VType being settled (e.g. SY, SM) |
| VBillDate | smalldatetime | Bill date |
| VBillVYear | int | Bill year |
| VBillFirm | nvarchar(4) | Bill firm (for cross-firm depot receipts) |
| VBillAmt | numeric(18,2) | Original bill amount |
| VTopCrDr | int | Top-level account code (for voucher header) |
| Discount | numeric(18,2) | Discount allowed/received on this bill |
| LateDays | int | Late days on this bill at time of payment |
| Interest | numeric(18,2) | Late payment interest on this bill |
| InterestCredit | numeric(18,2) | Interest credited back (if applicable) |
| JvNoDisc | int | JV Vno for discount entry |
| JvNoInt | int | JV/IntSale Vno for interest debit note |
| LessTDS | numeric(18,2) | TDS deducted on this bill settlement |
| JvNoLessTDS | int | JV Vno for TDS deduction entry |
| VnoList | int | Group number linking all lines of a late-payment interest batch |
| VNoListVYear | int | Financial year of VnoList |
| DueOnDt | smalldatetime | Original due date of the bill |
| Grace | int | Grace days allowed |
| GraceExtra | int | Additional grace days |
| IntRt | numeric(18,2) | Interest rate % |
| IntFromDate | smalldatetime | Interest calculation start date |
| IntParty | int | Party account code for interest posting |
| TdsRt | numeric(18,2) | TDS rate % on settlement |
| TDSOn | numeric(18,2) | TDS base amount |
| DbNtInt | nvarchar(1) | Debit note interest flag ('Y'/'N') |
| DbNtMonthly | int | Monthly interest calculation flag |
| DbNtInBillNo | nvarchar(20) | Related debit note bill number |
| JvNoIntInDepot | int | JV for depot interest posting |
| CgstRt / CgstAmt | numeric(18,2) | CGST rate and amount (on interest) |
| SgstRt / SgstAmt | numeric(18,2) | SGST rate and amount |
| IgstRt / IgstAmt | numeric(18,2) | IGST rate and amount |
| CrDrNoteNo | nvarchar(10) | Credit/debit note number |
| IsGSTDbNt | int | 1 = GST debit note |
| inttodate | smalldatetime | Interest calculation to-date |
| VCrAcCodeTDS | int | TDS credit account |
| VDrAcCodeTDS | int | TDS debit account |
| NoInterest | nvarchar(1) | 'Y' = no interest on this bill |
| SubGroupCd | int | Sub-group code |
| TdsTcsPartyCd | int | Party code for TDS/TCS |

### 3.2 tblOutStanding

Used by `GProcCreateOutStanding` in CommonPro.bas. Created when a sale/purchase is saved.

PK: `Vno + VType + VYear + VFirm + OutBillno`

| Column | Description |
|--------|-------------|
| Vno | Source document Vno |
| VDt | Document date |
| VType | Document type code |
| VYear | Financial year |
| VFirm | Company code |
| Outbillno | Bill number |
| OutAcCode | Party account code |
| OutBrokerCode | Broker code |
| OutMillCode | Mill code |
| OutAmount | Outstanding amount (for most VTypes) |
| OutAmountMillBill | Outstanding amount for Mill Bill (SM VType only) |
| OutDuedays | Due days |
| OutDuedate | Due date |
| OutDbNtAmt | Debit note amount |
| OutBag | Bags |
| OutItShort | Item short code |
| OutTicket | Ticket number |

**Note:** `tblOutStanding` does NOT appear in the schema file (full_schema.txt), suggesting it may have been removed from the main database schema during the GST migration and its functions replaced by computed queries on `tblSale.SlBillRecdAmt` and `tblPurch.PurBillPaidAmt`.

### 3.3 tblIntSale (Credit Notes / Debit Notes / Interest)

PK: `VNo + VType + VYear + VFirm + VCtrNo`

Used by frmCrnDrnGST (VTypes SV/SI/PV/PX) and frmLatePaymentIntGST (VType SI/MI).

| Column | Description |
|--------|-------------|
| VNo | Auto-generated Vno |
| VType | SV/SI/PV/PX/MI |
| Vdt | Document date |
| VYear / VFirm | Scope |
| VCtrNo | Line counter |
| VCrAcCode | Credit account |
| VDrAcCode | Debit account |
| VAmt | Amount |
| VNar1 / VNar2 | Narration |
| Broker | Broker code |
| DbNtInBillDt | Debit note bill date |
| VIsRefEntType | Reference entry type |
| VIsAudited | Audited flag |
| VBillNo / VBillVno / VBillType / VBillDate / VBillVYear | Source bill reference |
| VBillAmt | Source bill amount |
| Interest | Interest amount |
| InterestCredit | Interest credit |
| VnoList | Late payment batch group number |
| VBillFirm | Source bill firm |
| DbNtInBillNo | Related note bill number |
| JvNoIntInDepot | Depot interest JV Vno |
| CgstRt/Amt, SgstRt/Amt, IgstRt/Amt | GST on note/interest |
| CrDrNoteNo | Credit/debit note number |
| ItCode | Item code (for interest debit notes) |
| IsGSTDbNt | 1 = GST debit note |
| SlIRNNo / SlAckNo | IRN for debit note e-invoice |

**Criteria filter:** `frmCrnDrnGST` uses `Criteria = "Vtype='<type>' AND VFirm='<firm>' AND VYear=<year> AND vnolist=0"` — `vnolist=0` ensures standalone notes are shown separately from late-payment-interest-generated notes.

### 3.4 TblMillRecPay

Parallel to tblVoucher, stores mill receipt/payment entries. Used by `GProcCreateMillRecPay` in CommonPro.bas. Structure is similar to tblVoucher (same columns for Vno, VType, VCtrNo, VCrAcCode, VDrAcCode, VAmt, bill reference fields, narrations).

---

## 4. Form: frmReceipt.frm

### 4.1 VType Initialisation (Form_Load)

```vb
frmTypeStr = gfrmTypeStr

Select Case frmTypeStr
  "BR" → Label6 = "Bank Receipts" (or "Depot Receipt" if gCIsDepotFirm=1)
  "CR" → Label6 = "Cash Receipts"; lbl1(4) = "Code:"
  "MR" → Label6 = "Mill Bill Direct Payment"; lbl1(4) = "Mill Code:"
End Select
```

**BillsInFirm construction (multi-firm depot scoping):**
```vb
BillsInFirm = "'" & gCCode & "'"
' Also include all depot child firms:
SELECT CCode FROM tblMastCompany WHERE CDepotMainFirmCompCd = '<gCCode>'
' Each found CCode is appended: BillsInFirm = BillsInFirm + ",'" & CCode & "'"
```

This means a depot main-firm receipt can show outstanding bills from all its depot child companies.

### 4.2 Outstanding Bill Lookup (F1 in grid)

When the user presses F1 in the Bill No column of the receipt grid:

```sql
SELECT tblMastAccount.AcOurCode, tblMastAccount.AcName, SlBillNo, VType,
       FORMAT(Vdt,'dd/MM/yyyy'), SlBillAmt,
       SlBillAmt - SlBillRecdAmt - SlTdsAmt AS Balance,
       VYear, Vno,
       LTRIM(STR(Vno)) + '-' + VType + '-' + LTRIM(STR(VYear)) + '-' + RTRIM(VFirm) AS VnoKey,
       tblMastAccount.AcCode
FROM tblSale, tblMastAccount
WHERE tblSale.SlAcDrCode = tblMastAccount.AcCode
  AND SlBillAmt - SlBillRecdAmt - SlTdsAmt > 0
  AND VFirm IN (<BillsInFirm>)
```

This shows all unpaid sale bills for the current (and depot child) firms. The balance is computed as `SlBillAmt - SlBillRecdAmt - SlTdsAmt` — directly from the sale header, not from tblOutStanding.

### 4.3 Grid Structure (per-bill receipt lines)

Each row in the msGrid represents one bill being settled:

| Col | Content | tblVoucher Field |
|-----|---------|-----------------|
| 0 | VCtrNo (auto) | VCtrNo |
| 1 | Bill No | VBillNo |
| 2 | VBillYear (2 digits) | VBillVYear |
| 3 | Bill Date | VBillDate |
| 4 | Party OurCode | — |
| 5 | Party Name | VCrAcCode (resolved) |
| 6 | Bill Amt | VBillAmt |
| 7 | This Receipt Amt | VAmt |
| 8 | (discount col hidden) | — |
| 9 | Ref Type | VRefTp |
| 10 | Ref No | VRefNo |
| 11 | Ref Bank | VRefBank |
| 12 | VBillType | VBillType |
| 13 | Bill Vno | VBillVno |
| 14 | Interest | Interest |
| 15 | Interest Credit | InterestCredit |
| 16 | JvNoDisc | JvNoDisc |
| 17 | JvNoInt | JvNoInt |
| 18 | LessTDS | LessTDS |
| 19 | JvNoLessTDS | JvNoLessTDS |
| 20 | VnoList | VnoList |
| 21 | DueOnDt | DueOnDt |
| 22 | Grace | Grace |
| 23 | IntRt | IntRt |
| 24 | GraceExtra | GraceExtra |
| 25 | IntFromDate | IntFromDate |
| 26 | IntParty | IntParty |
| 27 | TdsRt | TdsRt |
| 28 | TDSOn | TDSOn |
| 29 | DbNtInt | DbNtInt |
| 30 | DbNtMonthly | DbNtMonthly |
| 31 | VBillFirm | VBillFirm |
| 32 | DbNtInBillNo | DbNtInBillNo |
| 33 | JvNoIntInDepot | JvNoIntInDepot |
| 34 | CgstRt | CgstRt |
| 35 | CgstAmt | CgstAmt |
| 36 | SgstRt | SgstRt |
| 37 | SgstAmt | SgstAmt |
| 38 | IgstRt | IgstRt |
| 39 | IgstAmt | IgstAmt |
| 40 | CrDrNoteNo | CrDrNoteNo |
| 41 | IsGstDbNt | IsGSTDbNt |
| 42 | inttodate | inttodate |
| 43 | VNoListVYear | VNoListVYear |
| 44 | LateDays | LateDays |
| 45 | VRefDate | VRefDate |

### 4.4 Save Flow (Receipt)

```
1. VNo = GProcGenerateIdMonthwise("TblVoucher","Vno", month, "Vdt", criteria)
2. GProcRstOpen rstbl (tblVoucher) in "O" mode
3. SaveRelatedRecord:
   For each grid row with VAmt > 0:
       rstbl.AddNew
       rstbl!VCtrNo = i
       rstbl!VCrAcCode = party account code (from grid party name)
       rstbl!VDrAcCode = bank/cash account code (txt(7))
       rstbl!VAmt = grid col 7
       rstbl!VBillNo = grid col 1
       rstbl!VBillType = grid col 12
       rstbl!VBillVYear = grid col 2 + 2000
       rstbl!VBillVno = grid col 13
       ... (all late payment interest, TDS, GST, discount, VnoList, DueOnDt, Grace... fields)
       rstbl.Update
   If grid row has VAmt=0 but VBillVno>0:
       Write a zero-amount "adjustment" row (marks the bill as referenced)
4. gCn.CommitTrans
5. Log: INSERT tblVoucher_Log (LogTp='A'/'M')
```

**Note:** `GProcCreateVoucher` and `GProcCreateOutStanding` calls are commented out in both frmReceipt and frmPayment. All voucher records are written directly via ADO `rstbl.AddNew/Update` pattern. The outstanding is NOT maintained in tblOutStanding by these forms; outstanding is tracked via the denormalized `SlBillRecdAmt` column in tblSale.

### 4.5 Header Fields (txt() array)

| txt() | Description |
|-------|-------------|
| txt(0) | Vno |
| txt(2) | Vdt (date) |
| txt(3) | VYear |
| txt(7) | VDrAcCode (bank account code) |
| txt(8) | VAmt (total amount) |
| txt(9) | VRefTp (reference type) |
| txt(10) | VRefNo (cheque/RTGS number) |
| txt(12) | VRefBank |
| txt(13–16) | VNar1–4 |
| txt(17) | VReconDt (bank reconciliation date) |
| txt(18) | VIsRefEntType |
| txt(19) | VIsAudited |
| txt(20) | VBillNo (top-level) |
| txt(21) | VBillVno (top-level) |
| txt(22) | VBillType (top-level) |
| txt(23) | VBillDate |
| txt(24) | VBillVYear |
| txt(25) | VBillAmt |
| txt(26) | VTopCrDr |
| txt(27) | Discount |

---

## 5. Form: frmPayment.frm

### 5.1 VType Initialisation

```vb
Case "BP" → Bank Payment (credit bank, debit party)
Case "CP" → Cash Payment (auto-sets account = "Cash In Hand")
```

For CP, `txttemp(6) = "Cash In Hand"` is set automatically — no bank account selection needed.

### 5.2 Outstanding Bill Lookup (Payment against supplier bills)

```sql
SELECT tblMastAccount.AcOurCode, tblMastAccount.AcName, SlBillNo, VType,
       FORMAT(Vdt,'dd/MM/yyyy'), SlBillAmt,
       SlBillAmt - SlBillRecdAmt AS Balance,
       VYear, Vno, ...
FROM tblSale, tblMastAccount
WHERE SlBillAmt - SlBillRecdAmt > 0
AND VFirm IN (<scope>)
```

Note: Payment form uses `SlBillAmt - SlBillRecdAmt` (without `SlTdsAmt` deduction) unlike the receipt form. This is because TDS is a buyer-side deduction, not applicable to outward payment scenarios in the same way.

### 5.3 Grid Structure (Payment)

Simplified vs. receipt — the payment grid per-bill columns include:
- Col 0: VCtrNo | Col 1: Party OurCode | Col 2: Party Name | Col 3: VAmt | Col 5: VRefTp | Col 6: VRefNo | Col 7: SubGroupCd | Col 8: NoInterest | Col 9: VRefBank | Col 10: VBillNo | Col 11: VBillType | Col 12: VBillDate | Col 13: VBillAmt | Col 14: VBillVYear | Col 15: VBillVno | Col 16: VBillFirm

Plus header-level fields: txt(28-45) for late payment fields (LateDays, Interest, InterestCredit, JvNoDisc, JvNoInt, LessTDS, JvNoLessTDS, VnoList, DueOnDt, Grace, IntRt, GraceExtra, IntFromDate, IntParty, TdsRt, TDSOn, DBNtInt, DbNtMonthly).

### 5.4 Delete Logic

```vb
Maxi = MAX(VCtrNo) FROM tblVoucher WHERE VNo=<vno> AND CriteriaStr
For i = 1 To Maxi:
    DELETE FROM tblVoucher WHERE VNo=<vno> AND CriteriaStr AND VCtrNo=i
```

Delete-then-reinsert pattern (no UPDATE) for both payment and receipt.

---

## 6. Form: frmJV.frm (Journal Voucher)

- Table: `tblVoucher`, VType = 'JV'
- Multi-line debit/credit entries
- No fixed party/bill structure — general purpose
- `CrDrNoteNo` field used for credit/debit note cross-reference
- GST fields (CgstRt/Amt, SgstRt/Amt, IgstRt/Amt) can be populated for GST adjustments
- Print: Crystal Report for JV

**frmJVwithTDS.frm:** Extension of frmJV that includes TDS-specific fields (`TdsTcsPartyCd`, `VCrAcCodeTDS`, `VDrAcCodeTDS`). Used for TDS payments to government accounts.

---

## 7. Form: frmLatePaymentIntGST.frm (Late Payment Interest)

### 7.1 Architecture — VnoList-based Navigation

This form navigates by `VnoList` (not `VNo`). A VnoList batch groups all voucher rows that belong to the same interest calculation session. Multiple sale bills can be in the same batch.

```vb
Const IdField = "VnoList"   ' Navigation key
Criteria = "VType='BR' AND VFirm IN (<BillsInFirm>) AND VNoListVYear=<year>"
```

**BillsInFirm** covers the current company and all depot child companies:
```vb
BillsInFirm = "'" & gCCode & "'"
' + all CCode from tblMastCompany WHERE CDepotMainFirmCompCd = '<gCCode>'
```

### 7.2 VnoList ID Generation

```vb
' New VnoList is MAX(VnoList)+1 across all BR vtype records in scope:
CondNewVno = " VType='BR' AND VnoList<>0 AND VFirm IN (<BillsInFirm>) AND VNoListVYear=<year>"
txt(35) = GProcGenerateId(tblName, "VnoList", CondNewVno)
```

### 7.3 Interest Calculation Flow

For each receipt row in the batch:
1. Days outstanding: `LateDays = DATEDIFF(days, DueOnDt, today) - Grace - GraceExtra`
2. Interest: `Interest = (VAmt × IntRt × LateDays) / 36500`
3. GST on interest: calculated based on party's GSTIN (intra/interstate)

**Save sequence:**
```vb
' Phase 1: Update tblVoucher with interest fields
UPDATE tblVoucher SET
    Interest = <calc_interest>,
    LessTDS = <tds_on_interest>,
    DueOnDt = '<due_dt>',
    VnoList = <vnolist>,
    Grace = <grace_days>,
    IntRt = <rate>,
    LateDays = <late_days>,
    GraceExtra = <extra_grace>,
    IntFromDate = <from_dt>,
    IntParty = <party_code>,
    TdsRt = <tds_rate>,
    TDSOn = <tds_base>
WHERE VBillVno=<sale_vno> AND VBillType='<type>' AND VBillVYear=<year> AND VFirm='<firm>'

' Phase 2: Create interest debit note in tblIntSale (VType='SI' or 'MI')
INSERT INTO tblIntSale (VNo, VType='SI', Vdt, ..., CgstRt, CgstAmt, SgstRt, SgstAmt,
    IgstRt, IgstAmt, CrDrNoteNo, IsGSTDbNt=1, VnoList=<vnolist>)

' Update tblVoucher.JvNoInt = newly created tblIntSale.Vno
UPDATE tblVoucher SET JvNoInt = <intsale_vno>, CrDrNoteNo = '<note_no>'
WHERE VnoList=<vnolist> AND VBillVno=<sale_vno>

' Phase 3 (if TDS on interest): Create JV in tblVoucher (VType='JV') for TDS deduction
' Update tblVoucher.JvNoLessTDS = JV Vno
UPDATE tblVoucher SET JvNoLessTDS = <jv_vno>
```

### 7.4 Delete / Reset Flow

```vb
' 1. Delete tblIntSale rows for SI/MI VType with this VnoList
DELETE FROM tblIntSale WHERE VType IN ('SI','MI') AND VFirm='<firm>' AND VYear=<year> AND VnoList=<vnolist>

' 2. Reset tblVoucher interest fields back to zero
UPDATE tblVoucher SET
    Interest=0, LessTDS=0, DueOnDt='', VnoList=0, Grace=0, IntRt=0,
    LateDays=0, JvNoInt=0, JvNoLessTDS=0, GraceExtra=0, IntFromDate=null,
    IntParty=0, TdsRt=0, TDSOn=0, DBNtInt=null, DbNtMonthly=0,
    DbNtInBillNo=null, CrDrNoteNo=null, VBillFirm=null
WHERE VnoList=<vnolist> AND VFirm IN (<BillsInFirm>) AND VNoListVYear=<year>
```

### 7.5 GST on Interest

```vb
' Rate from tblMastSetting: gSgstTaxRate, gCgstTaxRate, gIgstTaxRate
' Determine intra/interstate from party GSTIN
TaxAmtCGST = SubAmt × CDbl(txt(29)) / 100   ' txt(29) = CGST rate
' (Similarly for SGST and IGST)
```

The GST on interest becomes a debit note (tblIntSale.IsGSTDbNt=1) with IRN support.

---

## 8. Form: frmCrnDrnGST.frm (Credit Note / Debit Note with GST)

### 8.1 VType Mapping

| gfrmTypeStr | VType written | Caption |
|-------------|---------------|---------|
| PV | PV | Credit Note With GST (Purchase) |
| PX | PX | Debit Note With GST (Sales) |
| SV | SV | Credit Note With GST (Sales) |
| SI | SI | Debit Note With GST (Sales) |

Wait: the form caption "Credit Note / Debit Note" and the VType labels are:
- PV or PX → `LblSaleType = "Credit Note With GST (Purchase)"` / `"Credit Note With GST (Sales)"`
- SV or SI → `LblSaleType = "Debit Note With GST (Sales)"` / `"Debit Note With GST (Purchase)"`

### 8.2 Table and Criteria

- Table: `tblIntSale`
- Criteria: `Vtype='<type>' AND VFirm='<firm>' AND VYear=<year> AND vnolist=0`
  - The `vnolist=0` filter shows only standalone notes, not late-payment-interest-generated notes

### 8.3 Note Number Generation

```vb
VNo = GProcGenerateIdMonthwise("tblIntSale","Vno", month, "Vdt", criteria)
```

The note number (`CrDrNoteNo`) is a separate formatted string distinct from the Vno.

### 8.4 IRN Support

Same Excel import pattern as frmSalesGST:
```vb
UPDATE tblIntSale SET SlIRNNo='<irn>', SlAckNo='<ack>'
WHERE vtype='<type>' AND Vno='<vno>' AND VFirm='<firm>' AND VYear=<year>
```

### 8.5 GST Calculation

Rate lookup from `tblMastNarration WHERE Narration='G S T'` for CotCGSTRt/CotSGSTRt/CotIGSTRt (cotton) and PolCGSTRt/PolSGSTRt/PolIGSTRt (polyester). Same item-type + GSTIN prefix logic as sales.

---

## 9. tblMastSetting — GST Configuration Table

This single-row settings table (`tblMastSetting`) is read by `GProcGetSettingDetail` into global variables. Finance-relevant globals:

| Global | Description |
|--------|-------------|
| gSgstTaxRate | Default SGST rate % |
| gCgstTaxRate | Default CGST rate % |
| gIgstTaxRate | Default IGST rate % |
| gLPGrase | Late payment grace days |
| gLPIntRt | Late payment interest rate % |
| gTDSRate | TDS rate % |
| gLatePayIntAcCodeRec | Late payment interest receivable account (with GST) |
| gLPIntAcCodeWithoutGST | Late payment interest (without GST) |
| gTDSAcCodeRec | TDS recoverable account |
| gDiscAcCodeRec | Discount receivable account |
| gBrokerageSaleAc | Brokerage sale account |
| gCommissionSaleAc | Commission sale account |
| gRoundOffAc | Round-off account |
| gTcsRec / gTcsPay | TCS receivable / payable accounts |
| gSgstRCMRecCode / gCgstRCMRecCode / gIgstRCMRecCode | RCM GST receivable accounts |
| gTdsOnPurchCode / gTdsOnSalesCode | TDS on purchase / sales accounts |
| gAcCodeRY / gAcCodeVY | Sales return / purchase return accounts |

### 9.1 GST Ledger Accounts (from tblMastNarration)

`GProcGetGSTAccounts` reads `tblMastNarration WHERE Narration='G S T'`:

| Global | Column in tblMastNarration |
|--------|--------------------------|
| gSgstAcCode | CotSGSTAc (or dedicated SGST ledger column) |
| gCgstAcCode | CGST ledger account |
| gIgstAcCode | IGST ledger account |
| gSgstRCMAcCode | RCM SGST payable |
| gCgstRCMAcCode | RCM CGST payable |
| gIgstRCMAcCode | RCM IGST payable |
| CotCGSTRt / CotSGSTRt / CotIGSTRt | Cotton GST rates |
| PolCGSTRt / PolSGSTRt / PolIGSTRt | Polyester/other GST rates |

---

## 10. CommonPro.bas — Finance Infrastructure Functions

### 10.1 GProcCreateOutStanding

```vb
Public Sub GProcCreateOutStanding(Vno, Vdt, Vtype, BillNo, AcCode, BrkCode, Amt, _
    [DueDays], [DueDt], [MillCode], [DbNtAmt], [Bag], [ItemCode])

' Writes to tblOutStanding:
INSERT tblOutStanding:
    Vno, VDt, VType, vYear=gCYear, Outbillno=BillNo, OutAcCode=AcCode,
    OutBrokerCode=BrkCode, OutMillCode=MillCode,
    OutAmount = Amt   (for all VTypes EXCEPT SM)
    OutAmountMillBill = Amt  (for VType='SM' only)
    OutDuedays=DueDays, OutDuedate=DueDt, OutDbNtAmt=DbNtAmt,
    OutBag=Bag, OutItShort=ItemCode
```

**Important:** As seen in the forms, `GProcCreateOutStanding` calls are commented out in the current GST-era forms (frmSalesGST, frmReceipt, frmPayment). The function remains in CommonPro.bas but the outstanding tracking has shifted to the denormalized columns (`SlBillRecdAmt`, `PurBillPaidAmt`) in tblSale/tblPurch.

### 10.2 GProcCreateVoucher

```vb
Public Sub GProcCreateVoucher(frmAction, Vno, VSrno, Vdt, Vtype, AcCode, AcOCode, Amt,
    [BillNo], [BillDt], [RconDt], [Nar1], [Nar2], [Nar3], [Nar4],
    [IsRefEntryType], [MillExpNarrCode], [Ramt], [ShowUnAdj])

' Writes to tblVoucher:
rstbl.AddNew
rstbl!Vno = Vno;  rstbl!VDt = Vdt;  rstbl!VType = VType
rstbl!VCtrNo = VSrno;  rstbl!vAcCode = AcCode;  rstbl!VAcOCode = AcOCode
rstbl!Vamt = Amt;  rstbl!vbillno = BillNo;  rstbl!VBillDt = BillDt
rstbl!VReconDt = RconDt;  VNar1-4;  rstbl!vUserName = gUserName
rstbl!vEntStatus = "A";  rstbl!vYear = gCYear
rstbl!MillExpNarrCode = MillExpNarrCode
rstbl!VRecAmt = Ramt;  rstbl!VisShowUnadjLst = ShowUnAdj
rstbl!VIsRefEntType = IsRefEntryType
rstbl.Update
```

**Note:** This function also uses field names `vAcCode`/`VAcOCode` suggesting an older schema version. Current tblVoucher uses `VCrAcCode`/`VDrAcCode`. The function is still present but called primarily from the older non-GST-era logic.

### 10.3 GProcCreateMillRecPay

```vb
Public Sub GProcCreateMillRecPay(frmAction, Vno, VSrno, Vdt, Vtype, AcCode, AcOCode, Amt,
    [BillNo], [BillDt], [RconDt], [Nar1-4], [MillExpNarrCode])

' Writes to TblMillRecPay (same structure as tblVoucher)
```

Used specifically for mill payment tracking; separates mill transactions from regular vouchers.

### 10.4 GProcGetSettingDetail

```vb
Public Sub GProcGetSettingDetail()
' Reads tblMastSetting single row into 50+ global variables
' Tax codes: gTaxCodeSY/SO/SD/SM/ST/SYHunk/SYExempt, gTaxCodePY/PT/PI/PYHunk/PYExempt
' Transaction codes: gTranCdSY/SO/SD/ST/SYHunk/SYExempt, gTranCdPY/PT/PI/PYHunk/PYExempt
'                    gTranCdRY/RO/VY/VI
' Account codes: gAcCodeSY/ST/SYHank/SYExempt, gAcCodePY/PT/PYHank/PYExempt
' Late payment: gLatePayIntAcCodeRec/RecGST, gTDSAcCodeRec, gDiscAcCodeRec
' GST rates: gSgstTaxRate, gCgstTaxRate, gIgstTaxRate
' Policy: gLPGrase (grace days), gLPIntRt (interest rate), gTDSRate
' Brokerage: gBrokerageSaleAc, gCommissionSaleAc, gRoundOffAc
' TCS: gTcsRec, gTcsPay
' RCM: gSgstRCMRecCode, gCgstRCMRecCode, gIgstRCMRecCode
' TDS: gTdsOnPurchCode, gTdsOnSalesCode
' Return accs: gAcCodeRY, gAcCodeVY
```

Called at login and on company/year switch.

---

## 11. Bank Reconciliation (frmBankReConcilEntry.frm)

- Table: `tblVoucher`, VType = BR or BP
- Sets `VReconDt` on matched voucher lines
- Grid shows all unreconciled BR/BP entries for the period
- On save: `UPDATE tblVoucher SET VReconDt='<recon_dt>' WHERE VNo=<vno> AND VCtrNo=<ctrno>...`
- Used for monthly bank statement matching

---

## 12. OS Adjustment (frmOsAdjustment.frm)

- Table: `tblVoucher`
- Calls `GProcCreateOutStanding` and `GProcCreateVoucher` (one of the few remaining active callers)
- Purpose: manual adjustments to outstanding balances when bills have been partially settled or written off
- Used for opening balance entry or correction entries

---

## 13. Interest Calculation Date (frmIntCalDate.frm)

- Table: `tblInterestCalDate`
- Stores per-party override dates for interest calculation
- Schema: `(VFirm, AcCode, IntCalDt)`
- Used by frmLatePaymentIntGST to determine the effective start date for interest calculation on specific parties/firms

---

## 14. GSTR-2A Reconciliation

`PrcPrepareGSTR2A` stored procedure appears in the procedure list. It compares:
- Purchase invoices in `tblPurch` (what the firm has entered)
- Against the GSTN portal's GSTR-2A data

Output is used for input tax credit reconciliation. The procedure takes GSTIN, date range parameters and produces a reconciliation report.

---

## 15. Account Master (tblMastAccount) — Finance Relevant Fields

Key financial fields used across voucher forms:

| Column | Description |
|--------|-------------|
| AcCode | Primary key |
| AcName | Account name |
| AcOurCode | Short code used for entry (4-char) |
| AcGSTIN | GSTIN (18 chars) |
| AcPAN | PAN number |
| AcIsNonWork | 1 = non-working account (blocks entry) |
| AcIsLessTDSOnRec | 1 = deduct TDS on receipt (Section 194Q) |
| AcIsTDSfrom1stBill | 1 = deduct TDS from first bill |
| AcContNo | Mobile/contact number (used for WhatsApp) |
| AcPartyBank | Party's bank name |
| AcPartyAcNo | Party's bank account number |
| AcPartyBranch | Branch |
| AcPartyIFSC | IFSC code |
| AcDepotHoAc | Depot HO account code (for depot TDS routing) |
| AcMillBillFirm | Mill bill firm code |
| AcIsExemptMill | 1 = exempt mill (for mandi/cess exemption) |
| AcBrkCode | Default broker code for this party |
| AcMsOmsParty | MS/OMS party flag |

---

## 16. GST Filing Reports

### 16.1 GSTR-1 (outward supplies)

Reports use `tmpSale` staging table (populated during print) or direct tblSale queries. The e-invoice JSON export from frmSalesGST produces the GSTR-1 input.

### 16.2 GSTR-2A / 2B (inward supplies)

`PrcPrepareGSTR2A` stored procedure handles comparison of purchase records vs. portal data.

### 16.3 Balance Sheet / P&L

`GProcPrepareProfitLoss` and `GProcPrepareBalanceSheet` in Account.bas:

```vb
' VType patterns for P&L:
' Opening stock: OY (opening purchase sub)
' Purchases: PY (trade), VY (purchase return), PO (consignment)
' Sales: SY (trade), RY (sale return), SO (consignment), SD (depot)

' Balance sheet calls stored procedure with GST account parameters:
@CGSTAc, @SGSTAc, @IGSTAc, @CGSTRcmPayAc, @SGSTRcmPayAc, @IGSTRcmPayAc
```

### 16.4 Account Ledger

`PrcPrepareAccLedger` stored procedure generates party/account ledger from tblVoucher (all VTypes), tblSale, tblPurch, tblIntSale. Accepts `@AcCode`, `@From_dt`, `@To_dt`, `@VFirm` parameters.

---

## 17. Audit Trail

### 17.1 tblVoucher_Log

Shadow of tblVoucher with `LogNo`, `LogTp` ('A'/'M'/'D'), `UserName`, `LogDate`.

```vb
' On delete:
INSERT INTO tblVoucher_Log SELECT tblVoucher.*, <logno> AS LogNo, 'D' AS LogTp, ...
FROM tblVoucher WHERE VNo=<vno> AND CriteriaStr

' On add/modify:
INSERT INTO tblVoucher_Log SELECT tblVoucher.*, <logno> AS LogNo, 'A'/'M' AS LogTp ...
```

Active when: `gCYear >= 2023 AND gCIsLog = 1`.

### 17.2 tblIntSale_Log

Same pattern for tblIntSale (credit/debit notes and interest debit notes).

---

## 18. Key Business Rules

1. **Outstanding tracking is denormalized**: The current system tracks outstanding via `tblSale.SlBillRecdAmt` and `tblPurch.PurBillPaidAmt`. The `tblOutStanding` table and `GProcCreateOutStanding` calls are commented out in the main GST-era forms. The balance displayed to the user is a direct calculation: `SlBillAmt - SlBillRecdAmt - SlTdsAmt`.

2. **Multi-line voucher per single VNo**: One receipt/payment VNo can settle multiple bills. Each bill gets its own VCtrNo row in tblVoucher. All rows share the same VNo/VType/VYear/VFirm.

3. **VnoList is a batch group key for interest**: Late payment interest batches are grouped by VnoList. Navigation in frmLatePaymentIntGST is by VnoList, not VNo. All receipt rows belonging to the same interest calculation session share the same VnoList.

4. **Delete-reinsert pattern**: Both payment and receipt forms use delete-then-reinsert on modify (not UPDATE). The `DeleteRelatedRecord` sub deletes all VCtrNo rows for the VNo, then `SaveRelatedRecord` re-inserts them.

5. **Cross-firm depot receipts (BillsInFirm)**: Both frmReceipt and frmLatePaymentIntGST build `BillsInFirm` to include depot child companies. A receipt in the main firm can settle bills from child depot firms. VBillFirm column in tblVoucher stores which firm's bill is being settled.

6. **Cash payment auto-account**: For CP (Cash Payment), the credit account (cash/bank account) is always auto-set to "Cash In Hand" — no user selection required.

7. **MR (Mill Receipt)**: A direct receipt from mill (for mill bill settlement). Grid shows mill as the party. TblMillRecPay is also updated via GProcCreateMillRecPay for mill-specific tracking.

8. **GST on credit/debit notes**: All credit/debit notes in tblIntSale require a `CrDrNoteNo` number. IRN/e-invoice support is present via `SlIRNNo/SlAckNo` fields. The `vnolist=0` criteria filter distinguishes standalone notes from auto-generated interest notes.

9. **Late payment interest GST**: Interest debit notes carry GST (SI/MI VType in tblIntSale). These are separate from the voucher entry and require independent IRN generation for amounts above the e-invoice threshold.

10. **Bank reconciliation**: `VReconDt` in tblVoucher is the bank reconciliation date. Unreconciled entries have `VReconDt = NULL`. frmBankReConcilEntry updates this date when a cheque/RTGS entry is confirmed against the bank statement.

---

## 19. Multi-Company Finance Logic

- Each VFirm maintains its own voucher number series (monthly sequence scoped by VFirm + VYear + VType)
- Cross-firm transactions (depot → main): `VBillFirm` in tblVoucher records which firm's bill is settled; `BillsInFirm` query scope
- `tblMastCompany.CDepotMainFirmCompCd` defines the parent–child relationship
- Outstanding lookups include `VFirm IN (<BillsInFirm>)` for cross-firm visibility
- `gCIsDepotFirm = 1` flag changes: (a) receipt label, (b) TDS calculation routing (to `AcDepotHoAc`), (c) GST on invoice (shows depot GSTIN), (d) PayTo on invoices

---

## 20. Stored Procedures Used in Finance

| Procedure | Called From | Purpose |
|-----------|-------------|---------|
| PrcPrepareAccLedger | Reports | Full account ledger (party, GST, trading) |
| PrcPrepareAccLedgerBrokSale | Reports | Broker-wise sale ledger |
| PrcPrepareBrokSale | Reports | Brokerage sale summary |
| PrcPrepareClBalance | Reports | Closing balance report |
| PrcPrepareDailyEntryMillBill | Reports | Daily mill bill entry report |
| PrcPrepareDailyReport | Reports | Daily transaction summary |
| PrcPrepareDailyReportMail | Reports | Daily report for email |
| PrcPrepareGSTR2A | Reports | GSTR-2A reconciliation |
| PrcPrepareGetPassWiseSale | Reports | Gate-pass-wise sale report |
| PrcPrepareOutStangingSale | frmReceipt/Sales | Outstanding bill popup |

---

## 21. Integration Points Summary

| Integration | Mechanism | Details |
|-------------|-----------|---------|
| Crystal Reports (voucher print) | Crystl32.OCX / CRAXDRT | tblVoucher SELECT → rptPayment.rpt / rptReceipt.rpt |
| Crystal Reports (interest debit note) | CRAXDRT | tblIntSale JOIN tblVoucher → rptIntDbNote.rpt |
| tblSale.SlBillRecdAmt update | Direct gCn.Execute | Not in Receipt form — updated by a separate update path or via stored procedures |
| Bank reconciliation | VReconDt column | frmBankReConcilEntry sets VReconDt per voucher line |
| GSTR-2A | PrcPrepareGSTR2A SP | Purchase vs. portal reconciliation |
| e-Invoice (credit/debit notes) | Excel import | SlIRNNo/SlAckNo in tblIntSale updated manually |
| Interest calculation dates | tblInterestCalDate | Per-party override dates for interest start |
| Audit trail | tblVoucher_Log, tblIntSale_Log | Year>=2023, gCIsLog=1 |

---

## 22. Migration Design Notes

### Schema gaps to resolve in DhanMan

1. **Outstanding tracking migration**: HITRIX has a deprecated `tblOutStanding` and a live but denormalized `SlBillRecdAmt`/`PurBillPaidAmt` approach. DhanMan should implement a proper outstanding ledger with event sourcing (each receipt/payment creates an event that updates outstanding) rather than a running-total column.

2. **VnoList (interest batch grouping)**: The VnoList concept (grouping multiple receipt rows into a single interest batch) needs a first-class entity in DhanMan — an "interest batch" that groups multiple bills. The current approach of sharing a VnoList integer across tblVoucher rows is fragile.

3. **Multi-line voucher architecture**: DhanMan's voucher model must support one voucher header + N detail lines (one per bill), where each detail line has its own bill reference, payment amount, discount, TDS, interest, and GST fields.

4. **VBillFirm (cross-firm settlement)**: The ability to settle bills from a child depot firm while the receipt is recorded in the main firm needs a cross-entity payment allocation model in DhanMan.

5. **tblMastSetting migration**: The single-row settings table holds ~50 account codes and rates. DhanMan should migrate these to a structured settings API with named configuration keys per company.

6. **GST ledger account codes**: The `tblMastNarration WHERE Narration='G S T'` record stores all GST account codes. DhanMan needs a dedicated GST account configuration section in company settings.

7. **Bank reconciliation**: DhanMan needs a bank statement import/matching feature rather than the manual `VReconDt` update approach.

8. **Credit/debit note IRN**: tblIntSale notes above ₹50,000 require e-invoice IRN. DhanMan should automate IRN generation via IRP API for these notes.

9. **Late payment interest debit notes as tblIntSale vs. tblVoucher**: In HITRIX, the debit note for interest goes to `tblIntSale` (VType SI/MI) while the corresponding receipt row goes to `tblVoucher`. DhanMan should consolidate this into a unified interest charge entity with clear linkage to both the original bills and the interest levy document.

10. **GProcCreateVoucher commented out**: The fact that the main voucher creation function is commented out in production code indicates a migration mid-point in the original system. DhanMan should NOT replicate this fragmented state — all voucher creation must go through a single, tested pathway.

---

*End of Finance & GST Domain Analysis*
