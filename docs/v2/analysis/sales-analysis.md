# Sales Domain Analysis — HITRIX VB6 ERP

**Source path:** `src/ABabu_Cd_15-05-2026/`
**Primary forms analysed:** frmSalesGST.frm, frmSalesReturn.frm, frmSalesSIT.frm, frmSalesMillBill.frm, frmSalesOtherGST.frm, frmConfirmationSales.frm, frmDespatchDepot.frm
**Primary tables:** tblSale, tblSaleSub, tblSaleSub_Log, tblSale_Log, tblSitLrDetails, tblConfirmation, tmpSale

---

## 1. Overview

The sales domain covers six distinct transaction types that all write to `tblSale` + `tblSaleSub`. Each type is distinguished by its `VType` two-character code. The form `frmSalesGST.frm` handles the main GST-era billing (post-01/07/2017) for VTypes SY, SO, SD, SE. Separate forms handle Sale-in-Transit (frmSalesSIT, VType ST), Mill Bill (frmSalesMillBill, VType SM), Sales Return (frmSalesReturn, VType RY), other charges (frmSalesOtherGST, VType SC/SB), and sales confirmation letters (frmConfirmationSales, VType SA stored in `tblConfirmation`).

The sale form does **not** directly create outstanding or double-entry voucher records. Outstanding is tracked in `tblOutStanding` via the Receipt form (frmReceipt). The `tblSale.SlBillRecdAmt` column is updated by receipts. The sale header stores a computed `SlBillAmt` (total payable) and `SlBillRecdAmt` (received to date) for display.

---

## 2. VType Code Inventory

| VType | Description | Form | Account auto-set |
|-------|-------------|------|-----------------|
| SY | Trade Sale (GST) | frmSalesGST | gAcCodeSY (from tblMastSetting) |
| SO | Consignment Sale | frmSalesGST | From mill account (AcCrCode = mill AcCode) |
| SD | Depot Sale | frmSalesGST | From mill account (AcCrCode = depot mill AcCode) |
| SE | Exempt Sale | frmSalesGST | gAcCodeSYExempt |
| ST | Sale in Transit (SIT) | frmSalesSIT | gAcCodeST |
| SM | Mill Bill | frmSalesMillBill | From mill account |
| RY | Sales Return | frmSalesReturn | gAcCodeRY |
| SC | Sales Other (Credit Note / Other charges) | frmSalesOtherGST | varies |
| SB | Another Other sales type | frmSalesOtherGST | varies |
| SA | Sales Confirmation Letter | frmConfirmationSales | tblConfirmation (separate table) |

**Legacy VTypes** (pre-GST, from frmSales.frm): SY, SE share the same codes; frmSales.frm handles the same table with identical VTypes but different Criteria date range (`vdt < 01/07/2017`).

---

## 3. Primary Tables

### 3.1 tblSale (Header)

PK: `VNo + VType + VYear + VFirm`

| Column | Type | Size | Description |
|--------|------|------|-------------|
| VNo | int | — | Auto-generated monthly sequence (GProcGenerateIdMonthwise) |
| VType | nvarchar | 2 | Transaction type code (SY/SO/SD/SE/ST/SM/RY/SC/SB) |
| Vdt | smalldatetime | — | Invoice/voucher date |
| VYear | int | — | Financial year (e.g. 2024 = FY2024-25) |
| VFirm | nvarchar | 4 | Company code (gCCode) |
| SlAcDrCode | int | — | Party account code (debtor) |
| SlAcCrCode | int | — | Sales/Mill credit account code |
| SlMillCode | int | — | Mill account code |
| SlBroker | int | — | Broker account code |
| SlBillNo | nvarchar | 20 | Invoice number (format: SSSS999999 where SSSS = BillSr from tblMastBillSerial) |
| SlBillDt | smalldatetime | — | Bill date (may differ from Vdt) |
| SlTransport | int | — | Transport narration code (from tblMastNarration NarrType='R') |
| SlLorryNo | nvarchar | 20 | Vehicle number |
| SlLrNo | nvarchar | 40 | LR (Lorry Receipt) number |
| SlLrDate | smalldatetime | — | LR date |
| SlSubAmt | numeric(18,2) | — | Total of all sub-rows (gross goods value) |
| SlBillAmt | numeric(18,2) | — | Final payable amount (post-tax, post-rounding) |
| SlAdd1 | numeric(18,2) | — | Charity/add-on amount (item-level computed) |
| SlLess1 | numeric(18,2) | — | Deductions |
| SlExemptAmt | numeric(18,2) | — | Exempt component (SE type: based on purchase PurExemptPerKg × weight) |
| SlExemptPerKg | numeric(18,5) | — | Exempt rate per kg |
| SlTaxableAmt | numeric(18,2) | — | Assessable value (SubAmt + Add1 - Less1 - ExemptAmt) |
| SlTaxCode | int | — | Tax master code |
| SlTaxRate | numeric(18,2) | — | SGST rate % |
| SlTaxAmt | numeric(18,2) | — | SGST amount |
| SlTaxableAmt2 | numeric(18,2) | — | (Same as SlTaxableAmt, stored again for 2nd rate tier) |
| SlTaxRate2 | numeric(18,2) | — | CGST rate % |
| SlTaxAmt2 | numeric(18,2) | — | CGST amount |
| SlTaxRate3 | numeric(18,2) | — | IGST rate % (used when interstate; SGST+CGST set to 0) |
| SlTaxAmt3 | numeric(18,2) | — | IGST amount |
| SlAftTaxAdd1 | numeric(18,2) | — | Post-tax addition |
| SlAftTaxLess1 | numeric(18,2) | — | Post-tax deduction |
| SlRoff | numeric(18,2) | — | Rounding off (GProcMakeRounding) |
| DueDays | int | — | Credit days |
| DueDate | smalldatetime | — | Due date (Vdt + DueDays) |
| SlEntDt | smalldatetime | — | Entry date/time stamp |
| SlUser | nvarchar | 20 | Username who entered |
| SlTranCd | nvarchar | 5 | Transaction code (from tblMastTransactionCode) |
| SlIsHank | int | — | Flag: 1 = Hank/Yarn type item |
| GpVno | int | — | Gate Pass Vno reference |
| GpNo | nvarchar | 20 | Gate Pass number |
| GpDate | smalldatetime | — | Gate Pass date |
| GpVYear | int | — | Gate Pass year |
| SlCharityRt | numeric(18,2) | — | Charity rate per item (from tblMastItem.CharityRt) |
| SlCharityOn | int | — | Charity basis: 0=on Kg, 1=on Pkg |
| SlCrDays | int | — | Alternate credit days field |
| SlTaxPayAc | int | — | Tax payable account code |
| SlIsExempt | int | — | Flag: 1 = exempt transaction |
| SlCFormNo | nvarchar | 20 | C-Form number (legacy VAT) |
| SlCFormRecdDate | date | — | C-Form received date |
| SlCFromNote | nvarchar | 25 | C-Form note |
| SlE1FormNo | nvarchar | 20 | E1-Form number (legacy VAT) |
| SlE1FormRecdDate | date | — | E1-Form date |
| SlE1FromNote | nvarchar | 25 | E1-Form note |
| SlTmpNo1–3 | numeric(18,2) | — | Spare numeric fields |
| SlTmpChr1–3 | nvarchar | 25/64 | Spare character fields; SlTmpChr3 = 64 chars (used for IRN QR data) |
| SlBillRecdAmt | numeric(18,2) | — | Running received amount (updated by receipts) |
| SlBillLateIntAmt | numeric(18,2) | — | Late payment interest accrued |
| SlPrevBrokPd | nvarchar | 50 | Previous brokerage paid note |
| SlTcsOnAmt | numeric(18,2) | — | TCS taxable amount base |
| SlTcsRate | numeric(7,3) | — | TCS rate % |
| SlTcsAmt | numeric(18,2) | — | TCS amount |
| SlIRNNo | nvarchar | 64 | IRN (e-invoice reference number) |
| SlAckNo | nvarchar | 20 | Acknowledgement number for e-invoice |
| SlTdsRate | numeric(5,2) | — | TDS rate % |
| SlTdsAmt | numeric(18,2) | — | TDS amount deducted |
| SlTdsAc | int | — | TDS account code |
| SlMillPaid | numeric(18,2) | — | Mill payment made |
| SlGodown | int | — | Godown code |
| SlSizer | int | — | Sizer code |
| SlTdsOnAmt | numeric(18,2) | — | TDS base amount (stored separately from TdsAmt) |

### 3.2 tblSaleSub (Line Items)

PK: `SlSubItSrNo + Vno + VType + VYear + VFirm`

| Column | Type | Description |
|--------|------|-------------|
| SlSubItSrNo | int | Line sequence number (1..n) |
| Vno | int | Parent sale Vno |
| Vtype | nvarchar(2) | Parent VType |
| Vdt | smalldatetime | Copied from header |
| VYear / VFirm | — | Multi-company/year scope |
| SlSubItCtrlNo | int | Item control number (deprecated, now 0) |
| SlSubPItCtrlNo | int | Parent item ctrl no (= booking item ctrl no from tblBookingSub) |
| SlSubItCode | int | Item code (from tblMastItem) |
| SlSubBag | int | Bags |
| SlSubWt | numeric(18,3) | Weight in kg |
| SlSubRt | numeric | Rate |
| SlSubNetRate | numeric | Net rate (after adjustments) |
| SlSubRtPer | numeric | Rate per unit (e.g. per 100kg) |
| SlSubAmt | numeric(18,2) | Line amount = Wt × Rate / Per |
| SlSubGodown | int | Godown NarrCode |
| SlSubLotNo | nvarchar | Lot number |
| SlSubBookNo | int | Source booking Vno |
| SlSubBookDt | smalldatetime | Source booking date |
| SlSubBkType | nvarchar(2) | Source booking VType |
| SlSubBkVyear | int | Source booking year |
| SlSubBkSrNo | int | Source booking schedule sr no (links to tblBookingSub.BkSrNo) |
| SlSubBkDespSrNo | int | Source booking despatch sr no |

### 3.3 tblSale_Log / tblSaleSub_Log (Audit Trail)

Shadow copies of tblSale/tblSaleSub with additional columns:
- `LogNo` int — sequence within same Vno
- `LogTp` nchar(1) — 'A' (Add), 'M' (Modify), 'D' (Delete)
- `UserName` nvarchar(20)
- `LogDate` smalldatetime

Log entries created when: `gCYear >= 2023 AND gCIsLog = 1` (controlled by company year settings). All three operations (Add, Modify, Delete) are logged.

### 3.4 tblSitLrDetails (SIT LR Details)

Used by frmSalesSIT (VType ST) to track Lorry Receipt details from the source purchase.

| Column | Description |
|--------|-------------|
| LrDItSrNo | Item serial in LR |
| LrDCtrNo | Control number (PK) |
| LrDItCode | Item code |
| LrDBag / LrDWt | Bags and weight |
| LrDLRNo | LR number |
| LrdLRDate | LR date |
| LrdLRAmt | LR amount |
| SlVno / SlVtype / SlVyear / SlVfirm | Linked sale Vno (0 if not yet billed) |
| SlBillNo | Linked bill number |
| VNo / VType / VYear / VFirm | Source purchase document link |

On SIT sale save: `UPDATE tblSitLrDetails SET SlVno=<sale_vno>, SlVtype=<vtype>, ... WHERE LrDLRNo=<lr> AND VNo=<purch_vno>`
On SIT sale delete: `UPDATE tblSitLrDetails SET SlVno=0, SlVtype='', ... WHERE SlVno=<sale_vno>`

### 3.5 tblConfirmation (Sales Confirmation Letters)

Used by frmConfirmationSales (VType SA). Separate from tblSale. Fields include: Vno, Vtype='SA', Vdt, VYear, CfCrAcCode (seller), CfDrAcCode (buyer), CfItemNarr (goods description), CfQty, CfRate, plus narration fields (quality remark, packing, delivery terms, payment terms). Report: `rptSaleConfirmationLetter.rpt`.

---

## 4. Form: frmSalesGST.frm

### 4.1 Constants and Scope

```vb
Const tblName = "tblSale"
Const IdField = "Vno"
Const MaxNo = (large field count)
```

**Criteria filter** (Form_Load):
```
Vtype='<frmTypeStr>' AND VFirm='<gCCode>' AND VYear=<gCYear> AND vdt >= CONVERT(datetime,'01.07.2017',104)
```
This enforces the GST era: only records dated on or after 01/07/2017 are accessible in this form.

### 4.2 VType Initialisation (Form_Load)

```vb
frmTypeStr = gfrmTypeStr   ' set by calling menu before Show

Select Case frmTypeStr
  "SO" → LblSaleType = "Consignment Sale",  TrnCd = gTranCdSO,  SalesAccode = 0 (from mill)
  "SD" → LblSaleType = "Depot Sale",        TrnCd = gTranCdSD,  SalesAccode = 0 (from mill)
  "SY" → SalesAccode = gAcCodeSY,           TrnCd = gTranCdSY
  "SE" → SalesAccode = gAcCodeSYExempt,     TrnCd = gTranCdSYExempt
End Select
```

For SO/SD: the Sales A/C (SlAcCrCode) is taken from the mill master (`tblMastAccount.AcCode` of the selected mill). This means consignment and depot sales credit the mill's account, not a generic sales account.

### 4.3 Header Field Inventory (txt() array → tblSale columns)

The form uses a `txt()` control array bound to `tblSale` via `GProcSaveRecord` (tag-based). From code evidence:

| txt() index | tblSale Column | Label/Caption |
|-------------|----------------|---------------|
| txt(0) | Vno | Doc No |
| txt(1) | VType | (VType — set from frmTypeStr) |
| txt(2) / dtpDt(2) | Vdt | Date |
| txt(3) | VYear | Year |
| txt(4) | VFirm | Firm |
| txt(5) / txtCode(5) | SlAcDrCode (via AcOurCode lookup) | Party Code |
| txtTemp(5) | SlAcDrCode resolved name | Party Name |
| txt(7) / txtCode(7) | SlMillCode | Mill Code |
| txtTemp(7) | Mill name | |
| txtTemp(6) | SlAcCrCode name | Sales A/C |
| txtTemp(8) | SlBroker name | Broker |
| txtCode(8) | Broker our-code | |
| txt(9) | SlBillNo | Invoice No |
| txt(15) | SlSubAmt | Sub Amount |
| txt(16) | SlBillAmt | Bill Amount |
| txt(17) | SlAdd1 | Charity/Add1 |
| txt(18) | SlLess1 | Less1 |
| txt(19) | SlExemptAmt | Exempt Amt |
| txt(20) | SlExemptPerKg | Exempt/Kg |
| txt(21) | SlTaxableAmt | Assessable Value |
| txt(23) | SlTaxRate | SGST Rate % |
| txt(24) | SlTaxAmt | SGST Amount |
| txt(26) | SlTaxRate2 | CGST Rate % |
| txt(27) | SlTaxAmt2 | CGST Amount |
| txt(28) | SlAftTaxAdd1 | After-tax Add |
| txt(29) | SlAftTaxLess1 | After-tax Less |
| txt(30) | SlRoff | Round Off |
| txt(31) | DueDays | Due Days |
| dtpDt(32) | DueDate | Due Date |
| txt(35) / txtTemp(35) | SlTranCd | Transaction Code |
| txt(38) | GpNo | Gate Pass No |
| dtpDt(39) | GpDate | Gate Pass Date |
| txt(41) | SlTaxRate3 | IGST Rate % |
| txt(42) | SlTaxAmt3 | IGST Amount |
| txt(63) | SlTcsOnAmt | TCS On Amount |
| txt(64) | SlTcsRate | TCS Rate % |
| txt(65) | SlTcsAmt | TCS Amount |
| txt(68) | SlTdsRate | TDS Rate % |
| txt(69) | SlTdsAmt | TDS Amount |
| txt(72) | SlGodown | Godown |
| txt(74) | SlTdsOnAmt | TDS Base Amount |
| txtTemp(70) / txtCode(70) | SlTdsAc | TDS Account |

### 4.4 Grid (tblSaleSub) Columns

```
Col 0: Sr | Col 1: Bk.No | Col 2: Book Date | Col 3: Count (Item Name) |
Col 4: Godown | Col 5: Lot No | Col 6: Bag | Col 7: Std.Pack |
Col 8: Weight | Col 9: Rate | Col 10: Per | Col 11: Net Rate |
Col 12: Amount | Col 13: (hidden - SlSubItCtrlNo) | Col 14: (hidden - SlSubPItCtrlNo) |
Col 15: (hidden - SlSubBkType) | Col 16: (hidden - SlSubBkVyear) | Col 17: (hidden - SlSubBkSrNo)
```

Columns 1, 2 (booking reference) and 13–17 (internal linkage) are hidden (ColWidth=0).

---

## 5. Bill Number Generation

Invoice number generation uses `GProcGenerateInvNo(frmTypeStr, SalesAccode, BilSr)` triggered in the Add flow.

**Logic:**
1. Look up `tblMastNarration` WHERE NarrType='X' for the bill type label (e.g. "Trade" → NarrCode for SY, "Depot" → NarrCode for SD, "Consignment" → NarrCode for SO, "Exempt" → NarrCode for SE, "Sale In Transit" → NarrCode for ST)
2. Resolve SalesAccode from the selected Sales A/C name
3. Look up `tblMastBillSerial` WHERE MillCode=SalesAccode AND SaleType=NarrCdBlTp AND VFirm=gCCode → get BillSr (4-char prefix)
4. `GProcGenerateInvNo` constructs the next sequential number in format `BillSr + sequential_number`

**Note:** For SO/SD, SalesAccode = mill account code (credit account is the mill). This means different mills get different invoice series.

**VNo generation:** `GProcGenerateIdMonthwise("TblSale","Vno", Month(date), "Vdt", "Vtype='SY' and VFirm='...' and VYear=...")` — monthly sequence within firm/year/vtype/month.

---

## 6. GST Calculation Logic

### 6.1 Rate Selection (CalAmount function)

GST rates are determined dynamically based on party GSTIN and item type:

```vb
' Get item type (0=Cotton, 1=Polyester/other)
ItemType = GProcGetColumnValue("tblMastItem","ItName", item, "S", "ItType","N", "ItMillcode=<millcode>")

' Check party GSTIN prefix
If Left(lblTinNo.Caption,2) = "27"       ' Maharashtra intrastate
   Or Left(lblTinNo.Caption,1) = "U"     ' Unregistered
   Or Trim(lblTinNo.Caption) = "" Then   ' No GSTIN
    ' INTRASTATE: use SGST + CGST
    If ItemType=0 Then  ' Cotton
        txt(23) = tblMastNarration."G S T".CotCGSTRt   ' SGST
        txt(26) = tblMastNarration."G S T".CotSGSTRt   ' CGST
    Else                ' Polyester/Other
        txt(23) = tblMastNarration."G S T".PolCGSTRt
        txt(26) = tblMastNarration."G S T".PolSGSTRt
    End If
    txt(41) = 0   ' No IGST
Else
    ' INTERSTATE: use IGST only
    If ItemType=0 Then
        txt(41) = tblMastNarration."G S T".CotIGSTRt
    Else
        txt(41) = tblMastNarration."G S T".PolIGSTRt
    End If
    txt(23) = 0  txt(26) = 0
End If
```

### 6.2 Amount Calculation Formula

```
SubAmt     = Sum of all grid row amounts
AssValue   = SubAmt + SlAdd1 - SlLess1 - SlExemptAmt
SGST Amt   = AssValue × SlTaxRate  / 100
CGST Amt   = AssValue × SlTaxRate2 / 100
IGST Amt   = AssValue × SlTaxRate3 / 100
PreRoundAmt = SubAmt + SGST + CGST + IGST + AftTaxAdd1 - AftTaxLess1 + Add1 - Less1
BillAmt    = GProcMakeRounding(PreRoundAmt)    [rounds to nearest integer]
ROff       = BillAmt - PreRoundAmt
```

### 6.3 Exempt Sale (SE) Calculation

For VType SE, line items may reference a purchase booking (`SlSubPItCtrlNo`). The system fetches `PurExemptPerKg` from the linked purchase record (`tblPurch.PurExemptPerKg`):

```vb
ExmptAmt = ExmptAmt + (PurExemptPerKg × LineWeight)
```

This exemption amount is deducted from assessable value before applying GST.

### 6.4 TDS on Sales (Section 194Q)

Applicable from 01/07/2021. Triggered when:
- Company is TCS firm (`gCIsTcsFirm=1`) OR Depot firm (`gCIsDepotFirm=1`)
- Party has `AcIsLessTDSOnRec=1` OR `AcIsTDSfrom1stBill=1`

Calculation:
```vb
' Cumulative sales to this party (same PAN) in this year
TDSLimitAmt = SUM(SlBillAmt) WHERE party/PAN and vtype IN ('SY','SD') AND vdt <= today

If TDSLimitAmt >= 50,00,000 OR depot firm (older years) OR TdsFrom1stBill=1 Then
    TdsRate = 0.1%
    If year >= 2023 Then
        TdsOnAmt = portion of AssValue above the 50-lakh threshold
    Else
        TdsOnAmt = AssValue
    End If
    TdsAmt = TdsOnAmt × 0.1 / 100
    TdsAc  = gTdsOnSalesCode (or depot HO account for depot firms)
End If
```

**Stored fields:** SlTdsOnAmt, SlTdsRate, SlTdsAmt, SlTdsAc. The TDS amount is shown on the invoice but NOT deducted from SlBillAmt (it's a separate receivable adjustment handled at receipt time).

### 6.5 TCS (Tax Collected at Source)

TCS support exists in the schema (SlTcsOnAmt, SlTcsRate, SlTcsAmt) but the calculation block is commented out in the current code as of the May 2026 version. TCS fields are stored as zero. The fields are still passed to `tmpSale` and printed if non-zero.

---

## 7. Mill Validation (VType-Specific)

When party code (txtCode(5)) loses focus, the form validates the mill against a type-specific narration:

```vb
If frmTypeStr = "SY" Then
    ' Mill must be in tblMastNarration as NarrType='T' (Trade mills)
    Valid = (tblMastNarration WHERE Narration=millName AND NarrType='T')
ElseIf frmTypeStr = "SD" Then
    ' Mill must be NarrType='D' (Depot mills)
    Valid = (tblMastNarration WHERE Narration=millName AND NarrType='D')
ElseIf frmTypeStr = "SO" Or frmTypeStr = "SD" Then
    ' Auto-set Sales A/C from mill account
    txtTemp(6) = mill account name
End If
```

For SO and SD, the credit account (SlAcCrCode) is automatically set to the mill's AcCode — there is no separate "Sales A/C" field for these types.

---

## 8. Booking Integration

When a party code is entered for SY/SD/SE/SO, the form calls stored procedure `PrcPreparePendingBookingParty` to show a popup of pending booking lines. Selected booking lines populate grid columns:
- Col 1 (Bk.No): Booking Vno
- Col 2 (Book Date): Booking date
- Col 15 (hidden): Booking VType
- Col 16 (hidden): Booking VYear
- Col 17 (hidden): Booking Sr No (links to tblBookingSub.BkSrNo)
- Col 14 (hidden): SlSubPItCtrlNo (booking item ctrl no)

This creates a full traceability chain: Sale → Booking → (for SE) Purchase.

**Gate Pass integration:** For SY, SO, SD, the form supports Gate Pass number entry. When a GpNo is typed, it looks up `tblGatePass` (VnoBill=0) and populates item/party/godown/rate into the grid automatically.

---

## 9. Outstanding Balance Warning

After party code entry, `PrcPrepareOutStangingSale` is called with parameters including `@Accode` (party code), `@To_dt` (invoice date). It writes results to `TmpGentbl`. The form then queries:

```sql
SELECT VFirmName,VBillNo,VBillDt,Amt1,Amt3,Amt2,amt1+amt3-Amt2 as Os, Amt5
FROM TmpGentbl WHERE UserName='...' AND amt5>60 AND amt1+amt3-Amt2 > amt3
```

If bills older than 60 days exist, a warning popup is shown before proceeding with the new sale.

---

## 10. Save Flow (Add)

```
1. Validate data (ValidateData function)
2. txt(0) = GProcGenerateIdMonthwise("TblSale","Vno", month, "Vdt", criteria)  [new VNo]
3. Resolve godown code from grid row 1: txt(72) = NarrCode for godown name
4. gCn.BeginTrans
5. GProcRstOpen rstbl ... "O"  [open tblSale in updateable mode]
6. GProcSaveRecord Me, rstbl, FormAction, MaxNo, ...  [writes all txt() to tblSale fields via tag binding]
7. SaveRelatedRecord  [iterates grid, AddNew into tblSaleSub via ADO]
8. gCn.CommitTrans
9. If gCYear>=2023 AND gCIsLog=1: INSERT tblSale_Log (LogTp='A') + INSERT tblSaleSub_Log
```

**SaveRelatedRecord detail:**
```vb
For each grid row i = 1 to .rows-1:
    rstSub.AddNew
    rstSub!Vno         = txt(0)
    rstSub!Vtype       = frmTypeStr
    rstSub!SlSubItSrNo = i
    rstSub!SlSubItCode = lookup by ItName + MillCode
    rstSub!SlSubBag    = grid col 6
    rstSub!SlSubWt     = grid col 8
    rstSub!SlSubRt     = grid col 9  (Rate)
    rstSub!SlSubRtPer  = grid col 10 (Per)
    rstSub!SlSubNetRate= grid col 11
    rstSub!SlSubAmt    = grid col 12
    rstSub!SlSubLotNo  = grid col 5
    rstSub!SlSubGodown = lookup NarrCode for godown name
    rstSub!SlSubPItCtrlNo = grid col 14 (booking linkage)
    rstSub!SlSubBookNo = grid col 1   (booking Vno)
    rstSub!SlSubBookDt = grid col 2
    rstSub!SlSubBkType = grid col 15
    rstSub!SlSubBkVyear= grid col 16
    rstSub!SlSubBkSrNo = grid col 17
    rstSub.Update
```

### 10.1 Modify Flow

Delete + re-insert pattern before save:
```vb
gCn.Execute "DELETE FROM tblSaleSub WHERE vtype='<type>' AND vno=<vno> AND VFirm='<firm>' AND VYear=<year>"
' Then SaveRelatedRecord inserts fresh rows
```

Log entry on modify: LogTp='M' appended to tblSale_Log and tblSaleSub_Log.

### 10.2 Delete Flow

```vb
IF MsgBox("Confirm?") = vbYes THEN
    Log: INSERT tblSale_Log (LogTp='D') + INSERT tblSaleSub_Log (LogTp='D')
    DeleteRelatedRecord  ' deletes tblSaleSub rows
    rstbl.Delete         ' deletes tblSale row via ADO
    Navigate to last record
```

**Pre-delete guard:** Checks `tblVoucher WHERE VBillType=<vtype> AND VBillVno=<vno>` — if a receipt/payment already references this sale, delete is blocked with message.

---

## 11. Print Flow

Print is accessed via cmdPrintBill (Index 0=Printer, 1=Screen, 6=WhatsApp).

**Staging approach:**
```sql
-- 1. Clean staging table
DELETE FROM tmpSale WHERE SlUser='<username>'

-- 2. Populate from join
INSERT INTO tmpSale (Vno,Vtype,...,SlIRNNo,SlAckNo,SlTdsAmt)
SELECT a.*,b.*
FROM tblSale a, tblSaleSub b
WHERE a.vno=b.vno AND a.vtype=b.vtype AND a.vyear=b.vyear AND a.vfirm=b.vfirm
AND a.SlBillNo IN (<selected bill numbers>)

-- 3. QR code generation
QrCodeData  ' inserts QR string into SlTmpChr3 field of tmpSale

-- 4. Crystal Report
cryReport.ReportFileName = gReportPath & "rptSaleBillPrintGstQr.rpt"
cryReport.SelectionFormula = "{tmpSale.SlUser}='<username>'"
cryReport.Action = 1  ' Print / Show
```

**Report variants:**
- `rptSaleBillPrintGstQr.rpt` — Standard GST invoice with QR code (SY/SE/SD)
- `rptSaleBillPrintGst.rpt` — GST invoice without QR (used for WhatsApp PDF)
- `rptSaleBillGP.rpt` — Gate Pass copy of invoice
- `rptSaleRtnPrintGst.rpt` — Sales Return (frmSalesReturn)
- `rptSaleBillSITPrint.rpt` — SIT sale (frmSalesSIT)

**Print options (chkInvChallan checkboxes):**
- Index 0: Invoice copy
- Index 1: Duplicate copy
- Index 2: Triplicate copy
- Index 3: Quadruplicate copy
- Index 4: Gate Pass copy (uses rptSaleBillGP.rpt)

**Batch print by date:** `chkIsDatewise = 1` → selects all bills for a given date instead of by bill number range.

**Company formulas passed to Crystal:**
- `wCSTNo`, `wBSTNo`, `wGSTNo`, `wcPan`, `wPayTo`, `wBank`, `wBankAcNo`, `wBankIfsc`, `wPrintFor`
- For SD: depot GSTIN and PAN are fetched from the mill's AcGSTIN/AcPAN fields
- `wPayTo`: for SD = depot main firm name; for others = current company name

---

## 12. WhatsApp PDF Dispatch (Index=6)

```vb
' For each bill in tmpSale:
Set objCrystal = New CRAXDRT.Application
Set objReport = objCrystal.OpenReport(gReportPath & "rptSaleBillPrintGst.rpt", 1)
objReport.Database.LogOnServer "pdsodbc.dll","Hitrix","Vishnu","sa","123456"
objReport.RecordSelectionFormula = "{tmpSale.SlBillNo}='<billno>' ..."

FileNM = Format(Now,"ddMMyyyyhhmmss") & ".pdf"
ExportReportToPDF objReport, gReportPath & FileNM, "foo"

' Send to party
PtyMbl = tblMastAccount.AcContNo for party
SentWhatsApp "PDF", FileNM, "", PtyMbl, PtyName

' Send to broker
BrkMbl = tblMastAccount.AcContNo for broker
SentWhatsApp "PDF", FileNM, "", BrkMbl, BrkName

' Cleanup
Kill gReportPath & "*.pdf"
```

Errors (e.g. missing phone number) are logged to `tmpWhatsAppErr` and shown at end via `GProcWhatsAppError`.

---

## 13. IRN (e-Invoice) Update

A separate button (Index=5, Caption="Update IRN") reads an Excel spreadsheet:

```vb
' Reads Excel, for each row:
UPDATE tblSale SET SlIRNNo='<irn>', SlAckNo='<ack>'
WHERE vtype='<type>' AND SlBillNo='<billno>' AND VFirm='<firm>' AND VYear=<year>
' OR (if matching by party):
WHERE SlAcDrCode=<partyCd> AND SlBillNo='<billno>' ...
```

The IRN number is also written to `SlTmpChr3` (nvarchar 64) as part of QR code data via `QrCodeData` sub, which runs before print.

### QR Code Data Generation

```vb
' QrCodeData constructs the JSON-like QR string from tmpSale fields:
' Includes: GSTINseller, GSTINbuyer, SlBillNo, Vdt, HSN code, SlIRNNo, SlAckNo, item amounts, tax amounts
' Written into tmpSale.SlTmpChr3 before Crystal Report runs
```

---

## 14. e-Invoice JSON Export

frmSalesGST has a JSON export routine (case branch with File #1 Print statements) that generates the e-invoice JSON payload:

```vb
Print #1, Space(15) + Chr(34) + "No" + Chr(34) + ":" + Chr(34) + RsSA.Fields("SlBillNo") + ...
```

This is written to a file on disk for upload to the GST portal.

---

## 15. Form: frmSalesSIT.frm (Sale in Transit)

- Table: `tblSale`, VType = ST
- Grid shows LR details from `tblSitLrDetails` instead of booking lines
- Grid columns: `LR No | LR Date | Bag | Weight | SlBillNo | VNo | VType | VYear | VFirm | LrDItSrNo | LrDItCode | SlVno | SlVtype | SlVyear | SlVfirm`
- On save: `UPDATE tblSitLrDetails SET SlVno=<sale_vno>, SlVtype='ST', SlVyear=<year>, SlVfirm='<firm>', SlBillNo='<billno>' WHERE VNo=<purch_vno> AND LrDLRNo='<lr>'`
- On delete: `UPDATE tblSitLrDetails SET SlVno=0, SlVtype='', ... WHERE SlVno=<sale_vno>`
- SalesAccode = gAcCodeST (from tblMastSetting)
- Print: `rptSaleBillSITPrint.rpt` using same tmpSale staging

**LR Lookup:** When an item is entered in the grid, F2 key opens a multi-select dialog showing pending LR records:
```sql
SELECT LrDLRNo, LrdLRDate, LrDBag, LrDWt, LrDItCode, VNo, VType, VYear, VFirm
FROM tblSitLrDetails
WHERE LrDItCode=<itcode> AND (SlVno=0 OR slVno=<current_vno>)
```

---

## 16. Form: frmSalesMillBill.frm (Mill Bill)

- Table: `tblSale`, VType = SM
- Purpose: records the mill's own invoice to the trading firm (for consignment/depot stock received)
- TaxCode loaded from `gTaxCodeSM`
- Booking linkage: calls `PrcPrepareBookingMillBill` stored procedure to show pending bookings
- On save: `UPDATE tblBookingDesp SET BkDespTmpNo1=<sale_vno>` to mark the booking despatch as billed
- Outstanding: `SELECT SUM(SlBillAmt) FROM tblSale WHERE SlAcDrCode=<party> AND SlAcCrCode=<mill> AND vtype='SM'` — used to show cumulative mill bill outstanding
- For GProcCreateOutStanding: `GProcCreateOutStanding` is NOT called from this form directly. Outstanding for SM type uses the `OutAmountMillBill` column instead of `OutAmount` (set when VType='SM' in CommonPro.bas).

---

## 17. Form: frmSalesReturn.frm (Sales Return)

- Table: `tblSale`, VType = RY
- VNo generation: `GProcGenerateIdMonthwise("TblSale","Vno", month, "Vdt", "Vtype='RY' ...")`
- SalesAccode = `gAcCodeRY`
- Grid: same structure as frmSalesGST grid (booking reference in hidden columns)
- Print: `rptSaleRtnPrintGst.rpt` via tmpSale staging
- IRN: Update IRN button updates `tblSale.SlIRNNo / SlAckNo` from Excel (same pattern as main sales)
- Outstanding linkage: On delete, `tblOutStanding` and `tblVoucher` deletions are commented out — returns affect outstanding through JV entries rather than direct deletion

---

## 18. Form: frmSalesOtherGST.frm (Other/Credit Notes)

- Table: `tblSale`, VType = SC or SB (from gfrmTypeStr)
- For SC: `frmTypeStr = "SC"` used in save condition check
- For SB: `gfrmTypeStr = "SB"` used in search filter
- Uses same tblSale + tblSaleSub structure
- Log audit: same tblSale_Log / tblSaleSub_Log pattern
- Print: `tmpSale` staging with `rptSaleBillPrintGst.rpt` variant; includes SlIRNNo, SlAckNo
- Exempt amount calculation: if frmTypeStr="SE" branch triggers exempt calculation

---

## 19. Form: frmDespatchDepot.frm (Depot Despatch)

- Table: `tblBooking`, IdField = Vno (same as booking)
- Purpose: records physical despatch of goods from depot (updates booking despatch status)
- VType from gfrmTypeStr
- Does NOT create a tblSale record; it updates `tblBookingDesp.BkDespTmpNo1`

---

## 20. Form: frmConfirmationSales.frm (Sales Confirmation Letter)

- Table: `tblConfirmation`, VType = 'SA' (hardcoded in Form_Load)
- Separate table, not tblSale
- Fields: CfCrAcCode (seller), CfDrAcCode (buyer), CfItemNarr, CfQty, CfRate, narration text fields (quality remark, packing, delivery, payment terms)
- Search: `tblConfirmation JOIN tblMastAccount WHERE CfCrAcCode=AcCode AND Vtype='SA'`
- Report: `rptSaleConfirmationLetter.rpt`
- No tax/GST involvement — purely a pre-sale confirmation document

---

## 21. Key Business Rules

1. **Date gate**: frmSalesGST only handles `vdt >= 01/07/2017`. Older pre-GST sales use frmSales.frm (same tables, different Criteria).
2. **Mill-bill serial validation**: Before invoice number generation, the selected Sales A/C must have a matching row in `tblMastBillSerial` (MillCode + SaleType + VFirm). If not found, system shows "Check Bill Sr not given..." and blocks save.
3. **Delete protection**: A sale cannot be deleted if `tblVoucher` has a receipt referencing it (`VBillType=<vtype> AND VBillVno=<vno>`).
4. **SO/SD auto-account**: For consignment and depot, the credit account is automatically set to the mill's AcCode — user cannot override.
5. **GST intra/interstate auto-detection**: The system checks the first two characters of the party's GSTIN (loaded into lblTinNo). Code "27" = Maharashtra. Blank/unregistered = intrastate. Any other state code = interstate → IGST.
6. **TDS Section 194Q threshold**: Cumulative sales (SY + SD vtype) to the same party (or same PAN) in the financial year are summed. TDS kicks in once threshold of ₹50 lakh is crossed, or from first bill for parties flagged `AcIsTDSfrom1stBill=1`.
7. **Charity rate**: Items can have `CharityRt` and `CharityOn` (0=per kg, 1=per bag) in `tblMastItem`. The charity amount is auto-computed per line and summed into `SlAdd1`.
8. **Booking traceability**: Every tblSaleSub line stores the source booking's VNo, Vtype, VYear, BkSrNo in hidden columns, creating full end-to-end booking → sale traceability.
9. **Exempt sale (SE) must link to purchase**: The SE form checks `SlSubPItCtrlNo > 0` to find the purchase exemption rate. Without a purchase linkage, ExmptAmt=0.

---

## 22. Multi-Company Logic

- `VFirm` column on every tblSale/tblSaleSub row = `gCCode` (4-char company code)
- All Criteria strings include `VFirm='<gCCode>'`
- Depot firm: `gCIsDepotFirm=1` global flag changes TDS logic, and for SD-type sales the depot GSTIN/PAN is taken from the mill's account master (`AcGSTIN`, `AcPAN`) rather than the current firm's GSTIN
- The depot's main firm: `gCDepotMainFirm` is a global holding the depot's associated main firm code; used to set `wPayTo` on invoice prints

---

## 23. Legacy Form: frmSales.frm

Pre-GST form (same tblSale, same VTypes: SY, SE, SO, SD).

- Criteria excludes `vdt >= 01/07/2017` — operates on pre-GST records
- No QR code, no IRN, no TDS/TCS
- VType ST uses `gAcCodeST`
- Form_Load for SY/SE: same pattern; for SO/SD: Caption = "Consignment Sale" / "Depot Sale"
- Used for historical data viewing/modification only

---

## 24. Stored Procedures Used

| Procedure | Called From | Purpose |
|-----------|-------------|---------|
| PrcPreparePendingBookingParty | frmSalesGST (txtCode_LostFocus for mill) | Returns pending booking lines for a party+mill combination |
| PrcPrepareOutStangingSale | frmSalesGST (txtCode_LostFocus for party) | Prepares outstanding bill summary for warning popup |
| PrcPrepareLotwiseGodownStock | frmSalesGST (grid item entry, year>=2023) | Shows lot-wise godown stock for item selection |
| PrcPrepareLotwiseStock | frmSalesGST (grid item entry) | Shows lot-wise stock without godown filter |
| PrcPrepareBookingMillBill | frmSalesMillBill | Returns pending booking despatch for mill bill entry |

---

## 25. Integration Points Summary

| Integration | Mechanism | Details |
|-------------|-----------|---------|
| Crystal Reports print | Crystl32.OCX `.Action=1` | tmpSale staging → rptSaleBillPrintGstQr.rpt |
| WhatsApp PDF | CRAXDRT API (OpenReport + ExportReportToPDF) | Per-bill PDF, send via `SentWhatsApp` with mobile from AcContNo |
| e-Invoice IRN | Excel read via Sheet.Cells | Updates SlIRNNo/SlAckNo from GST portal export |
| QR Code | QrCodeData() sub | Generates QR payload into tmpSale.SlTmpChr3 before print |
| Outstanding warning | PrcPrepareOutStangingSale SP → TmpGentbl | Read-only warning; does not block save |
| Booking linkage | PrcPreparePendingBookingParty SP | Populates grid with pending bookings |
| Gate Pass linkage | Direct tblGatePass query | Populates party/item from gate pass for SY/SO/SD |
| Audit log | Direct INSERTs to tblSale_Log, tblSaleSub_Log | Year>=2023, gCIsLog=1 |
| Receipt/Payment | tblVoucher (VBillType/VBillVno FK) | Outstanding tracked in Receipt form; delete blocked if receipt exists |

---

## 26. Migration Design Notes

### Schema gaps to resolve in DhanMan

1. **tmpSale staging table**: DhanMan will need an equivalent server-side or client-side bill rendering endpoint. The Crystal Report uses a flat denormalized join of header + sub rows per user session.

2. **Bill serial management** (`tblMastBillSerial`): DhanMan must replicate the MillCode + SaleType + VFirm → BillSr mapping. Invoice numbering is mill-specific, not company-global.

3. **GST interstate/intrastate auto-detection**: Logic is based on GSTIN prefix "27". DhanMan should implement this as a configurable state code (not hardcoded "27") with the party's state code from `tblMastState`.

4. **Multi-copy print**: The chkInvChallan array (4 copies + gate pass) needs a corresponding multi-copy flag in the DhanMan print API.

5. **TDS 194Q threshold**: DhanMan needs cross-financial-year cumulative tracking per PAN (not just per AcCode) to match the `tblMastAccount.AcPAN`-based join logic.

6. **Legacy frmSales data**: Pre-2017 records in tblSale (VType SY/SO/SD before GST) need to be migrated to DhanMan with a clear era distinction or as read-only archive.

7. **Booking traceability**: The `SlSubBkSrNo → tblBookingSub.BkSrNo` linkage must be preserved in the DhanMan data model for audit trail purposes.

8. **SIT (ST) form**: The tblSitLrDetails linkage (LR details to purchase document) needs a dedicated LR management entity in DhanMan with the same claim/release pattern.

9. **SE Exempt**: Requires purchase-side `PurExemptPerKg` to compute exemption at sale time. DhanMan purchase module must expose this value.

10. **IRN/e-invoice**: The Excel import workflow is manual. DhanMan should offer API-based e-invoice generation (NIC/IRP integration) with direct IRN population.

---

*End of Sales Domain Analysis*
