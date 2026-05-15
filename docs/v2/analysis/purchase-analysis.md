# Purchase Domain Analysis — HITRIX VB6 ERP

**Source path:** `src/ABabu_Cd_15-05-2026/`
**Primary forms analysed:** frmPurchaseTradeGST.frm, frmPurchaseInward.frm, frmPurchaseSIT.frm, frmPurchaseReturn.frm, frmPurchaseOther.frm, frmPurchaseTrade.frm (legacy)
**Primary tables:** tblPurch, tblPurchSub, tblPurchSub_Log, tblPurch_Log, tblSitLrDetails, TmpPackingSlip

---

## 1. Overview

The purchase domain records all incoming goods and services across six VType codes that all write to `tblPurch` + `tblPurchSub`. The main GST-era trade purchase form is `frmPurchaseTradeGST.frm` (VTypes PY and PE). Separate forms handle goods inward/outward for factory stock (frmPurchaseInward, VTypes PI and PR), Sale-in-Transit purchases (frmPurchaseSIT), Purchase Return / Shortage (frmPurchaseReturn), and other/sundry purchases (frmPurchaseOther).

Like the sales domain, the purchase form does **not** directly call `GProcCreateOutStanding` or `GProcCreateVoucher`. Payment outstanding is tracked in `tblOutStanding` via the Payment form (frmPayment). The `tblPurch.PurBillPaidAmt` column is updated by payment entries.

A unique feature of purchase is **RCM (Reverse Charge Mechanism)** support: when a mill is not registered under GST, the buyer is required to self-account for GST. These fields (`PurIsRCMBill`, `PurRCMCgstRt`, etc.) are stored in `tblPurch`.

**Capital goods purchase** (`PurIsCapitalGoods`) is a flag on the purchase header — used for ITC (Input Tax Credit) reporting to distinguish between regular and capital goods purchases.

**TDS on purchase** (Section 194Q / 194C): computed on `PurBillAmt` aggregated by party PAN within the financial year.

---

## 2. VType Code Inventory

| VType | Description | Form | Purchase A/C |
|-------|-------------|------|--------------|
| PY | Trade Purchase (GST) | frmPurchaseTradeGST | gAcCodePY |
| PE | Exempt Purchase | frmPurchaseTradeGST | gAcCodePYExempt |
| PO | Consignment Purchase (old) | frmPurchaseTradeGST | — |
| PI | Goods Inward | frmPurchaseInward | gTranCdPI |
| PR | Goods Outward | frmPurchaseInward | — |
| PT | Purchase SIT | frmPurchaseSIT | gAcCodePT / gAcCodeST |
| VY | Purchase Return | frmPurchaseReturn | gAcCodeVY |
| PX | Other Purchase (Other charges) | frmPurchaseOther | varies |

**Legacy VTypes** (frmPurchaseTrade.frm, pre-GST): PY, PE — same tables, same VType codes, different date range criteria.

---

## 3. Primary Tables

### 3.1 tblPurch (Header)

PK: `VNo + VType + VYear + VFirm`

| Column | Type | Size | Description |
|--------|------|------|-------------|
| VNo | int | — | Auto-numbered (monthly or annual) |
| VType | nvarchar | 2 | Transaction type code |
| Vdt | smalldatetime | — | Document date |
| VYear | int | — | Financial year |
| VFirm | nvarchar | 4 | Company code (gCCode) |
| PurAcCrCode | int | — | Mill/supplier credit account (payable) |
| PurAcDrCode | int | — | Purchase debit account |
| PurMillCode | int | — | Mill account code |
| PurBroker | int | — | Broker account code |
| PurBillNo | nvarchar | 20 | Supplier's invoice number |
| PurBillDt | smalldatetime | — | Supplier's invoice date |
| PurTransport | int | — | Transport narration code |
| PurLorryNo | nvarchar | 20 | Vehicle number |
| PurLrNo | nvarchar | 40 | LR number |
| PurLrDate | smalldatetime | — | LR date |
| PurSubAmt | numeric(18,2) | — | Total of all sub-row amounts |
| PurBillAmt | numeric(18,2) | — | Final payable amount |
| PurAdd1 | numeric(18,2) | — | Addition (freight, etc.) |
| PurLess1 | numeric(18,2) | — | Deductions |
| PurExemptAmt | numeric(18,2) | — | Exempt component (cotton cess/mandi tax etc.) |
| PurExemptPerKg | numeric(18,5) | — | Exempt rate per kg (fed to SE sale calculations) |
| PurTaxableAmt | numeric(18,2) | — | Assessable value (SubAmt + Add1 - Less1 - ExemptAmt) |
| PurTaxCode | int | — | Tax master code |
| PurTaxRate | numeric(18,2) | — | SGST rate % (tier 1) |
| PurTaxAmt | numeric(18,2) | — | SGST amount |
| PurTaxableAmt2 | numeric(18,2) | — | Assessable value tier 2 (same value) |
| PurTaxRate2 | numeric(18,2) | — | CGST rate % |
| PurTaxAmt2 | numeric(18,2) | — | CGST amount |
| PurTaxRate3 | numeric(18,2) | — | IGST rate % |
| PurTaxAmt3 | numeric(18,2) | — | IGST amount |
| PurAftTaxAdd1 | numeric(18,2) | — | Post-tax addition |
| PurAftTaxLess1 | numeric(18,2) | — | Post-tax deduction |
| PurRoff | numeric(18,2) | — | Rounding off |
| DueDays | int | — | Credit days |
| DueDate | smalldatetime | — | Due date |
| PurEntDt | smalldatetime | — | Entry timestamp |
| PurUser | nvarchar | 20 | User who entered |
| PurTranCd | nvarchar | 5 | Transaction code |
| PurIsHank | int | — | 1 = Hank/Yarn item type |
| PurTaxPayAc | int | — | Tax payable account |
| PurIsExempt | int | — | 1 = exempt transaction flag |
| PurTmpNo1–3 | numeric(18,2) | — | Spare numeric fields |
| PurTmpChr1–3 | nvarchar | 25–50 | Spare character fields |
| PurIsCapitalGoods | int | — | 1 = capital goods purchase (affects ITC classification) |
| PurBillPaidAmt | numeric(18,2) | — | Running paid amount (updated by Payment form) |
| PurBillPaidDate | smalldatetime | — | Last payment date |
| PurE1FormNo | nvarchar | 20 | E1 form (legacy VAT) |
| PurCFormNo | nvarchar | 20 | C form (legacy VAT) |
| **RCM fields:** | | | |
| PurIsRCMBill | int | — | 1 = Reverse Charge Mechanism bill |
| PurRCMBillNo | nvarchar | 15 | RCM bill/challan number |
| PurRCMBillDt | smalldatetime | — | RCM date |
| PurRCMSaleAc | int | — | RCM sale account code |
| PurRCMCgstRt / PurRCMCgstAmt | numeric(18,2) | — | RCM CGST rate and amount |
| PurRCMSgstRt / PurRCMSgstAmt | numeric(18,2) | — | RCM SGST rate and amount |
| PurRCMIgstRt / PurRCMIgstAmt | numeric(18,2) | — | RCM IGST rate and amount |
| PurRCMBillAmt | numeric(18,2) | — | RCM total bill amount |
| **TDS fields:** | | | |
| PurTdsJvNo | int | — | JV Vno for TDS entry (deprecated/commented out) |
| PurTdsRate | numeric(6,2) | — | TDS rate % |
| PurTdsAmt | numeric(18,2) | — | TDS amount |
| PurTdsAC | int | — | TDS account code (gTdsOnPurchCode) |
| SubGroupcd | int | — | Sub-group code |
| PurTcsOnAmt | numeric(18,2) | — | TCS base amount (legacy) |
| PurTcsRate | numeric(7,3) | — | TCS rate |
| PurTcsAmt | numeric(18,2) | — | TCS amount |
| PurMillPaid | numeric(18,2) | — | Mill payment made |
| PurTdsOnAmt | numeric(18,2) | — | TDS base amount (year>=2023) |

### 3.2 tblPurchSub (Line Items)

PK: `VNo + VType + VYear + VFirm + PurSubItSrNo`

| Column | Type | Description |
|--------|------|-------------|
| PurSubItSrNo | int | Line sequence (1..n) |
| PurSubItCtrlNo | bigint | Item control number (unique across firm/year) — used for cross-module stock tracking |
| PurSubPItCtrlNo | bigint | Parent item ctrl no (booking item ctrl no from tblBookingSub) |
| PurSubItCode | int | Item code (from tblMastItem) |
| PurSubBag | int | Bags received |
| PurSubWt | numeric(18,3) | Weight in kg |
| PurSubRt | numeric(18,5) | Rate |
| PurSubNetRate | numeric(18,5) | Net rate |
| PurSubRtPer | numeric(18,2) | Rate per unit |
| PurSubAmt | numeric(18,2) | Line amount |
| PurSubGodown | int | Godown NarrCode |
| PurSubLotNo | nvarchar(15) | Lot number |
| PurSubBookNo | int | Source booking Vno |
| PurSubBookDt | smalldatetime | Booking date |
| PurSubBkItCtrlNo | int | Booking item ctrl no |
| PurSubIsExempt | int | 1 = this line is exempt |
| PurSubTmpNo1–3 | numeric(18,2) | Spare numeric |
| PurSubTmpChr1–3 | nvarchar(25) | Spare character |
| PurSubIssBag | int | Issued bags |
| PurSubIssWt | numeric(18,3) | Issued weight |
| PurSubBkVyear | int | Booking year |
| PurSubBkSrNo | int | Booking schedule sr no |

**Critical:** `PurSubItCtrlNo` is a unique item-level identifier across the entire company/year. It is this column that `tblSaleSub.SlSubPItCtrlNo` references to link a sale line item back to its purchase line. This is the foundation of the purchase-to-sale traceability chain.

### 3.3 tblPurch_Log / tblPurchSub_Log (Audit Trail)

Same pattern as sales: shadow copies with `LogNo`, `LogTp` ('A'/'M'/'D'), `UserName`, `LogDate`.
Active when: `gCYear >= 2023 AND gCIsLog = 1`.

---

## 4. Form: frmPurchaseTradeGST.frm

### 4.1 Constants

```vb
Const tblName = "tblPurch"
Const IdField = "Vno"
```

**Criteria:** `vtype='<frmTypeStr>' AND VFirm='<gCCode>' AND VYear=<gCYear>` (no GST date gate — handles all years for PY/PE).

### 4.2 VType Initialisation (Form_Load)

```vb
frmTypeStr = gfrmTypeStr

If frmTypeStr = "PE" Then
    txttemp(6) = gAcCodePYExempt  ' Exempt purchase A/C
Else
    txttemp(6) = gAcCodePY         ' Regular purchase A/C
End If
```

For PO (consignment purchase), the credit account is the mill's AcCode (same pattern as SO/SD sales).

### 4.3 Grid (tblPurchSub) Columns

```
Col 0: Sr | Col 1: Bk.No (hidden) | Col 2: Bk.Year (hidden) |
Col 3: Count (Item Name) | Col 4: Godown (hidden) | Col 5: Lot No |
Col 6: Bag | Col 7: Weight | Col 8: Rate | Col 9: Per | Col 10: Amount |
Col 11: (hidden - PurSubItCtrlNo) | Col 12: (hidden - PurSubPItCtrlNo) |
Col 13: (hidden - PurSubBkSrNo)
```

### 4.4 Header Field Mapping (txt() → tblPurch columns)

| txt() index | tblPurch Column | Description |
|-------------|-----------------|-------------|
| txt(0) | VNo | Document number |
| txt(1) | VType | Transaction type |
| dtpDt(2) / txt(2) | Vdt | Document date |
| txt(3) | VYear | Financial year |
| txt(4) | VFirm | Company code |
| txt(5) / txtCode(5) | PurAcCrCode | Mill/Supplier code |
| txttemp(5) | PurAcCrCode name | Supplier name |
| txttemp(6) | PurAcDrCode name | Purchase A/C |
| txttemp(7) | PurMillCode name | Mill name |
| txtCode(7) | PurMillCode our-code | |
| txttemp(8) | PurBroker name | Broker |
| txt(9) | PurBillNo | Supplier bill number |
| dtpDt(10) | PurBillDt | Supplier bill date |
| txt(15) | PurSubAmt | Sub amount |
| txt(16) | PurBillAmt | Bill amount (final) |
| txt(17) | PurAdd1 | Addition |
| txt(18) | PurLess1 | Less |
| txt(19) | PurExemptAmt | Exempt amount |
| txt(20) | PurExemptPerKg | Exempt per kg |
| txt(21) | PurTaxableAmt | Assessable value |
| txt(23) | PurTaxRate | SGST rate % |
| txt(24) | PurTaxAmt | SGST amount |
| txt(26) | PurTaxRate2 | CGST rate % |
| txt(27) | PurTaxAmt2 | CGST amount |
| txt(28) | PurAftTaxAdd1 | After-tax add |
| txt(29) | PurAftTaxLess1 | After-tax less |
| txt(30) | PurRoff | Round off |
| txt(31) | DueDays | Due days |
| dtpDt(32) | DueDate | Due date |
| txt(37) | PurTaxRate3 | IGST rate % |
| txt(38) | PurTaxAmt3 | IGST amount |
| txt(46) | PurIsCapitalGoods | Capital goods flag (chkIsCapitalGoods) |
| txt(63) | PurTdsRate | TDS rate % |
| txt(64) | PurTdsAmt | TDS amount |
| txt(65) / txttemp(65) | PurTdsAC | TDS account |
| txt(68) | PurTcsOnAmt | TCS on amount |
| txt(69) | PurTcsRate | TCS rate |
| txt(70) | PurTcsAmt | TCS amount |
| txt(72) | PurTdsOnAmt | TDS base amount (year>=2023) |

---

## 5. GST Calculation Logic (CalAmount)

### 5.1 Exempt Mill Detection

On mill code entry, `AcIsExemptMill` flag is read from `tblMastAccount`:

```vb
IsExemptMill = GProcGetColumnValue("TblMastAccount","AcName",txttemp(5),"S","ACIsExemptMill","N")
```

If IsExemptMill=1, for each grid row:
- Fetch `ExemptRt` from `tblMastItem` for the item/mill combination
- If `ExemptRt = 100`: entire line amount is exempt
- Otherwise: `wExemptAmt += Round(LineAmt × ExemptRt / 100, 2)`

This calculated `wExemptAmt` is stored in `PurExemptAmt` and the per-kg rate (`PurExemptPerKg = PurExemptAmt / TotWt`) feeds into SE (exempt sale) calculations at the time of sale.

### 5.2 Intra/Interstate GST Detection (Purchase)

Uses the **current company's GSTIN prefix** vs. supplier's GSTIN prefix:

```vb
' gCGStin = current company GSTIN (e.g. "27AAAAA...")
' lblTinNo.Caption = supplier GSTIN (loaded on party entry)

If Left(gCGStin,2) = Left(lblTinNo.Caption,2)   ' Same state as company
   Or Left(lblTinNo.Caption,2) = ""               ' Supplier unregistered
   Or Left(lblTinNo.Caption,1) = "U" Then         ' Composition dealer
    ' INTRASTATE: SGST + CGST
Else
    ' INTERSTATE: IGST only
End If
```

**Key difference from sales:** Sales uses party GSTIN prefix "27" hardcoded (Maharashtra). Purchase compares the first 2 characters of company GSTIN vs. supplier GSTIN — this is more generalizable and correct for all states.

### 5.3 Tax Amount Tolerance Check

The purchase form has a ±₹1 tolerance check before blocking save:

```vb
' SGST validation
If Abs(Round(AssValue × SGST_Rate/100, 2) - SGST_Amt) > 1 And FormAction=AddNew Then
    Warning + SetFocus to tax amount field

' CGST validation
If Abs(Round(AssValue × CGST_Rate/100, 2) - CGST_Amt) > 1 And FormAction=AddNew Then ...

' IGST validation
If Abs(Round(AssValue × IGST_Rate/100, 2) - IGST_Amt) > 1 And FormAction=AddNew Then ...
```

This means the system accepts the supplier's actual invoice tax amount if it differs from computed by less than ₹1 (rounding differences), but warns/blocks if the difference is more than ₹1.

### 5.4 Amount Formula (Purchase)

```
SubAmt     = Sum of grid row amounts
AssValue   = SubAmt + PurAdd1 - PurLess1 - PurExemptAmt
SGST Amt   = AssValue × SGST_Rate / 100   (within ±1 tolerance of typed amount)
CGST Amt   = AssValue × CGST_Rate / 100
IGST Amt   = AssValue × IGST_Rate / 100
TCS Amt    = if date >= 2020/01/10: PurTcsOnAmt × 0.075 / 100   (if no manual TCS)
BillAmt    = SubAmt + SGST + CGST + IGST + TCS + PurAdd1 - PurLess1 + PurRoff + AftTaxAdd1 - AftTaxLess1
```

**Note:** Round-off (PurRoff) in purchase is NOT auto-computed — it is typed manually by the user or remains from prior record, unlike sales where GProcMakeRounding is called automatically.

### 5.5 TDS on Purchases (Section 194Q / 194C)

Applicable from 01/07/2021 for TCS/Depot firms:

```vb
' Cumulative purchase amount (PY + PO vtypes) from this supplier (or same PAN) in current year
TDSLimitAmt = SUM(PurBillAmt) FROM tblPurch WHERE PurAcCrCode=<party> (or same PAN)
              AND Vtype IN ('PY','PO') AND vyear=<year> AND PurBillDt <= today
              AND (not current record)

' Special threshold for mills ZRA, ZRS, ZRP, RJS (from 01/08/2025, year 2025):
'   Threshold = ₹30 crore (300,000,000)
' Standard threshold for all others:
'   Threshold = ₹50 lakh (5,000,000)

If TDSLimitAmt >= Threshold AND TcsAmt=0 Then
    TdsRate = 0.1%
    TdsOnAmt = portion of AssValue above threshold (year>=2023) OR AssValue (older)
    TdsAmt = TdsOnAmt × 0.1 / 100
    TdsAC = gTdsOnPurchCode
End If
```

Stored in: `PurTdsRate`, `PurTdsAmt`, `PurTdsAC`, `PurTdsOnAmt`.

The TDS JV (`PurTdsJvNo`) update code is commented out — TDS entries are handled manually via frmJVwithTDS or frmJV.

---

## 6. RCM (Reverse Charge Mechanism)

When `PurIsRCMBill = 1`, the purchase record stores additional RCM-specific fields:

| Field | Description |
|-------|-------------|
| PurRCMBillNo | RCM self-invoice number |
| PurRCMBillDt | Self-invoice date |
| PurRCMSaleAc | Sale account for RCM output |
| PurRCMCgstRt/Amt | CGST on RCM (self-assessed) |
| PurRCMSgstRt/Amt | SGST on RCM (self-assessed) |
| PurRCMIgstRt/Amt | IGST on RCM (self-assessed) |
| PurRCMBillAmt | RCM self-invoice total |

RCM applies when buying from an unregistered supplier. The purchasing firm must self-calculate GST liability (`gSgstRCMRecCode`, `gCgstRCMRecCode`, `gIgstRCMRecCode` from GProcGetSettingDetail). These account codes are configured in tblMastSetting.

---

## 7. Capital Goods Flag

```vb
txt(46) = chkIsCapitalGoods.Value   ' on ReadFields:
chkIsCapitalGoods.Value = Val(txt(46))
```

`PurIsCapitalGoods = 1` stored in tblPurch. Used in:
- Balance Sheet / ITC computation to distinguish regular vs. capital ITC
- The `GProcPrepareBalanceSheet` stored procedure uses `@CGSTRcmPayAc`, `@SGSTRcmPayAc` etc. to include capital goods in the correct GST credit bucket

---

## 8. Duplicate Bill Number Validation

Before save, the system checks for duplicate supplier bill numbers:

```sql
SELECT vno FROM tblPurch
WHERE PurBillNo='<billno>' AND PurAcCrCode=<supplierCode>
AND vno <> <current_vno> AND Vtype='<type>'
AND VFirm='<firm>' AND VYear=<year>
```

If found, the save is blocked with a duplicate warning.

---

## 9. Delete Protection

The IsRelatedRecord function checks two conditions:

**Condition 1:** A payment voucher references this purchase:
```sql
SELECT Vno,Vdt FROM tblVoucher
WHERE VBillType='<vtype>' AND VBillVno=<vno>
AND VFirm='<firm>' AND VBillVYear=<year>
```

**Condition 2 (PE type only):** A sale sub-row references this purchase's item ctrl no:
```sql
' For each PurSubItCtrlNo in tblPurchSub for this purchase:
SELECT Vno,Vtype FROM tblSale,tblSaleSub
WHERE tblSaleSub.SlSubPItCtrlNo = <PurSubItCtrlNo>
```

If either check finds records, delete is blocked with a message showing the related document number.

---

## 10. Save Flow

```
1. Validate data (duplicate PurBillNo check + tax amount tolerance checks)
2. If Add: VNo = GProcGenerateIdMonthwise("tblPurch","Vno", month, "Vdt", criteria)
          OR GProcGenerateId (depending on form configuration)
3. gCn.BeginTrans
4. GProcRstOpen rstbl ... "O"
5. GProcSaveRecord Me, rstbl, FormAction, MaxNo, ...  [writes tblPurch header]
6. SaveRelatedRecord  [writes tblPurchSub lines via ADO AddNew loop]
7. gCn.CommitTrans
8. Log: INSERT tblPurch_Log (LogTp='A'/'M') + INSERT tblPurchSub_Log
```

**SaveRelatedRecord for tblPurchSub:**
```vb
For each grid row i:
    rstSub.AddNew
    rstSub!PurSubItSrNo = i
    rstSub!PurSubItCode = lookup by ItName + MillCode
    rstSub!PurSubBag    = grid col 6
    rstSub!PurSubWt     = grid col 7
    rstSub!PurSubRt     = grid col 8
    rstSub!PurSubRtPer  = grid col 9
    rstSub!PurSubAmt    = grid col 10
    rstSub!PurSubLotNo  = grid col 5
    rstSub!PurSubGodown = lookup NarrCode for godown name
    rstSub!PurSubPItCtrlNo = grid col 12  (booking item ctrl no)
    rstSub!PurSubBookNo    = grid col 1
    rstSub!PurSubBkVyear   = grid col 2
    rstSub!PurSubBkSrNo    = grid col 13
    ' PurSubItCtrlNo is auto-generated bigint (identity or max+1)
    rstSub.Update
```

---

## 11. Print Flow — Packing Slip (frmPurchaseTradeGST)

The purchase form has a packing slip print feature (separate from the main purchase report):

**Staging:**
```sql
-- 1. Clean
DELETE FROM TmpPackingSlip WHERE UserName='<user>'

-- 2. Populate: join tblPurch + tblMastAccount + tblMastNarration + tblMastItem + tblBags
SELECT a.Vno, a.Vdt, b.AcName, b.AcAdd1..., c.Narration, a.PurBillNo, a.PurBillDt,
       d.ItName, e.SrNo, e.BagNo, e.CartoonWt
FROM tblPurch a, tblMastAccount b, tblMastNarration c, tblMastItem d, tblBags e
WHERE a.vno=<vno> AND a.Vtype='<type>' AND a.VFirm='<firm>' AND a.VYear=<year>
AND a.PurAcCrCode=b.AcCode AND e.InGodown=c.NarrCode AND e.InItCode=d.ItCode

-- 3. Consolidate: for bags with multiple entries (pivot logic)
UPDATE TmpPackingSlip SET SrNo2=..., BagNo2=..., Weight2=...
UPDATE TmpPackingSlip SET SrNo3=..., BagNo3=..., Weight3=...
UPDATE TmpPackingSlip SET SrNo4=..., BagNo4=..., Weight4=...

-- 4. Crystal Report
cryReport.ReportFileName = gReportPath & "rptPackingSlip_Purch.rpt"
```

The packing slip uses `tblBags` (a bag/cartoon-level detail table) to show individual bag numbers and weights.

---

## 12. Form: frmPurchaseInward.frm (Goods Inward / Outward)

- Table: `tblPurch`, VTypes: PI (Goods Inward) and PR (Goods Outward)
- Form caption: PI → "Goods Inward", PR → "Goods Outward"
- Tax code: `gTaxCodePI`
- Transaction code: `gTranCdPI`
- Grid: same tblPurchSub structure as frmPurchaseTradeGST
- LR details: when VType=PI, tblSitLrDetails is populated via `gfrmBagInwVType = frmTypeStr`
- Save: same ADO pattern; tblPurch_Log + tblPurchSub_Log audit
- **PR (Goods Outward)**: records return of goods from factory; same table but VType=PR with gAcCode from settings

---

## 13. Form: frmPurchaseSIT.frm (Purchase in Transit)

- Table: `tblPurch`, VType = PT
- The form also creates `tblSitLrDetails` rows for each LR number entered in the grid
- On save: inserts LR detail rows with `VNo=<purch_vno>, VType='PT', ...` and `SlVno=0` (not yet sold)
- On delete:
  ```sql
  DELETE FROM tblPurchSub WHERE vtype='PT' AND vno=<vno> AND VFirm='<firm>' AND VYear=<year>
  DELETE FROM tblSitLrDetails WHERE vtype='PT' AND vno=<vno> AND ...
  ```
  But first checks `SELECT LrDBag,LrDWt,SlBillNo FROM tblSitLrDetails WHERE SlVno>0 AND vtype='PT' AND vno=<vno>` — if any LR has been billed (SlVno>0), delete is blocked
- LR grid: contains LR No, LR Date, Bag, Weight, SlBillNo, VNo, VType, VYear, VFirm

---

## 14. Form: frmPurchaseReturn.frm (Purchase Return / Shortage)

- Table: `tblPurch`, VType = VY
- Account: `gAcCodeVY` (from tblMastSetting)
- Caption: "Purchase Return / Shortage"
- Grid: same tblPurchSub structure
- Print: "Purchase Return / Shortage Printing" frame → Crystal Report
- Save: same pattern; allows booking linkage (PE type has special exempt handling)
- Outstanding: frmPayment handles payment against VY returns (creates debit note or reduces outstanding)

---

## 15. Form: frmPurchaseOther.frm (Other Purchases)

- Table: `tblPurch`, VType = PX (or from gfrmTypeStr)
- Caption: "Other Purchase"
- Capital goods checkbox: `chkIsCapitalGoods` → `PurIsCapitalGoods` in tblPurch
- Same GST calculation structure as frmPurchaseTradeGST
- Tax code: from `gTaxCodePY` settings (commented out `gTaxCodePYHunk` variant)
- No booking linkage — item control numbers not linked to bookings

---

## 16. Booking Integration

On party/mill selection in frmPurchaseTradeGST, a stored procedure shows pending booking lines. Selected booking lines populate:
- Grid col 1 (Bk.No): booking Vno
- Grid col 2 (Bk.Year): booking VYear
- Grid col 12 (hidden): PurSubPItCtrlNo (booking item ctrl no)
- Grid col 13 (hidden): PurSubBkSrNo

The purchase item control number (`PurSubItCtrlNo`) becomes the linkage point for future sale entries. When the SE (exempt sale) form is used, it fetches `PurExemptPerKg` via the `PurSubItCtrlNo` → `tblSaleSub.SlSubPItCtrlNo` chain.

---

## 17. Key Business Rules

1. **Duplicate supplier bill number**: A bill number from the same supplier cannot be entered twice in the same year. Validation is by `PurBillNo + PurAcCrCode + VYear + VFirm`.

2. **GST rate tolerance**: When entering a purchase, the manually typed GST amounts are accepted if they differ from computed amounts by ≤₹1. This accommodates supplier rounding differences. A difference >₹1 triggers a validation warning.

3. **Exempt mill auto-calculation**: Mills flagged `AcIsExemptMill=1` automatically compute the mandi/cess exempt component at the item level using `tblMastItem.ExemptRt`. This feeds `PurExemptPerKg` for use in subsequent exempt sales.

4. **Capital goods flag**: Checked at time of purchase entry; affects ITC reporting. Cannot be changed after a payment voucher is linked.

5. **RCM self-invoice**: When `PurIsRCMBill=1`, the buyer generates their own GST liability. The RCM tax amounts use separate account codes (`gSgstRCMRecCode`, `gCgstRCMRecCode`) configured in `tblMastSetting`.

6. **Delete protection (SIT)**: A purchase in transit (PT) that has been partially billed (SlVno>0 in tblSitLrDetails) cannot be deleted.

7. **Delete protection (PE/PY)**: A purchase cannot be deleted if any `PurSubItCtrlNo` from its lines appears in `tblSaleSub.SlSubPItCtrlNo` (i.e., goods have already been sold).

8. **TDS special mill threshold**: For mills with codes ZRA, ZRS, ZRP, RJS from 01/08/2025 onwards (FY2025-26), TDS threshold is ₹30 crore instead of the standard ₹50 lakh. This is hardcoded in the CalAmount function.

9. **PurSubItCtrlNo uniqueness**: This is a bigint that provides unique identification across all purchase line items in a firm/year. The Sales Return and Exempt Sale forms rely on this to find the original purchase.

---

## 18. Multi-Company Logic

- All tblPurch/tblPurchSub rows scoped by `VFirm = gCCode`
- No cross-firm purchase consolidation in HITRIX (each firm has its own purchase records)
- Depot firms: `gCIsDepotFirm=1` affects TDS calculation in both purchase and sale
- TDS query uses `vfirm='<gCCode>'` explicitly (no cross-firm TDS aggregation)

---

## 19. Stored Procedures Used

| Procedure | Called From | Purpose |
|-----------|-------------|---------|
| PrcPreparePendingBookingParty | frmPurchaseTradeGST (mill entry) | Returns pending bookings for mill-party combination |
| PrcPrepareDailyInward | Reports | Daily goods inward report (uses PI/PR vtypes) |
| PrcPrepareGodownDelivery | Reports | Godown delivery tracking |
| PrcPrepareGodownDeliveryPending | Reports | Pending delivery items |

---

## 20. Integration Points Summary

| Integration | Mechanism | Details |
|-------------|-----------|---------|
| Crystal Reports (packing slip) | Crystl32.OCX | TmpPackingSlip staging → rptPackingSlip_Purch.rpt |
| tblBags (bag-level detail) | Direct query | Bag numbers and weights from tblBags joined on InVNo/InVType/InVYear/InVFirm |
| Booking linkage | SP PrcPreparePendingBookingParty | Populates grid with pending booking lines |
| Audit log | Direct INSERTs to tblPurch_Log, tblPurchSub_Log | Year>=2023, gCIsLog=1 |
| Payment (outstanding) | tblVoucher (VBillType/VBillVno FK) | Delete blocked if payment voucher references this purchase |
| Exempt sale linkage | PurSubItCtrlNo → SlSubPItCtrlNo | PE purchase → SE sale item-level traceability |
| SIT LR management | tblSitLrDetails | PT purchase creates LR rows; ST sale claims them |
| RCM self-accounting | PurIsRCMBill + separate RCM fields | Creates RCM liability via JV (manual) using gSgstRCMRecCode etc. |

---

## 21. Migration Design Notes

### Schema gaps to resolve in DhanMan

1. **PurSubItCtrlNo (bigint)**: This is the most critical cross-module linkage key in HITRIX. DhanMan's purchase line items need a globally unique ID (UUID or sequence) that survives the migration and is preserved in sale line items.

2. **RCM fields**: DhanMan purchase module needs a dedicated RCM section with self-invoice generation capability. The six RCM rate/amount fields (CGST/SGST/IGST × Rate/Amount) need proper UI.

3. **Capital goods ITC**: DhanMan needs a capital goods flag on purchase lines (not just header) because a single purchase invoice can have mixed capital and regular items.

4. **Exempt mill / ExemptRt**: The item-level `ExemptRt` field (in tblMastItem) and the mill-level `AcIsExemptMill` flag need to be replicated. This drives the `PurExemptPerKg` calculation that feeds SE sales.

5. **TDS special mill threshold**: The ZRA/ZRS/ZRP/RJS mills with ₹30-crore threshold is a business-specific rule added in August 2025. DhanMan needs a configurable TDS threshold per party (not hardcoded).

6. **Packing slip (tblBags)**: The `tblBags` table stores cartoon/bag-level detail beyond the purchase line. DhanMan needs an equivalent bag/packing detail sub-table linked to purchase lines.

7. **SIT (PT) form**: The tblSitLrDetails claim/release pattern for purchase-side LRs must be maintained in DhanMan. A PT purchase creates LR records; a ST sale claims them.

8. **PI/PR (Goods Inward/Outward)**: These forms track factory-level stock movement (not GST invoices). DhanMan should map these to a stock movement entity rather than a purchase invoice entity.

9. **PurBillPaidAmt auto-update**: This column is updated by the Payment form (frmPayment). DhanMan needs a trigger or service to keep this denormalized running total in sync.

10. **Duplicate PurBillNo check**: This validation is at the firm+year+party+billno level. DhanMan must implement this uniqueness constraint at the database level, not just in UI.

---

*End of Purchase Domain Analysis*
