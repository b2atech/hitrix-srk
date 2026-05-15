Attribute VB_Name = "Stock"
Option Explicit
'======== Utility Year End Procedure ==============
Public Sub GProcUtlYearEnd()
'On Error GoTo ErrorRoutine
Dim gCnNextYr As Connection '- Connection For Next Yr
Dim NextCYr As String
Dim NextDbNm As String
Dim ToAcCode As Long
Dim ToAgCode As Long
Dim ToCGCode As Long
Dim ToItCode As Long
Dim ToNarrCode As Long
'Dim CYr As String
'== Validate Year End already compleded
If GProcGetColumnValue("tblCompanyYear", "CompCode", gCCode, "S", "CompFdt", "S", "CompIsYrEnd=1 and CompYear='" & gDbYr & "'", gSetupCn) <> "" Then
    MsgBox "Year End of " & gCYSDate & " is already completed.", vbOKOnly + vbInformation, "Utility - Year End"
    Exit Sub
End If
'== Create Next Year Db File
NextCYr = Format(gCYEDate, "yy")
NextCYr = NextCYr + Format(CLng(NextCYr) + 1, "00") '- Yr Code
NextDbNm = gCCode & NextCYr
ToAcCode = GProcGenerateId("TblMastAccount", "AcCode") + 500 '- For New A/c Creation ( max(code ) + 500)
ToAgCode = GProcGenerateId("tblMastGroup", "AgCode") + 200 '- Group Creation
ToCGCode = GProcGenerateId("tblMastCompGroup", "CGCode") + 200 '- Comp Group Creation
ToItCode = GProcGenerateId("tblMastItem", "ItCode") + 500 '- Item Creation
ToNarrCode = GProcGenerateId("tblMastNarration", "NarrCode") + 200 '- Narration Creation
gCn.Close
'- [ Access ]
If gBackEndDB = gBackEndAccess Then FileCopy gDbLocation, gDataPath & "\" & NextDbNm & ".mdb"
If gBackEndDB = gBackEndAccess Then
    GProcConnectionOpen gCn
Else
    GProcConnectionOpenORA gCn
End If
'-- insert New Year Details to Comp Year & update current Year ( set YearEnd=1, set Master Start & End Code)
gSetupCn.Execute "update tblCompanyYear set CompIsYrEnd=1,CompEAcCode=" & ToAcCode & ",CompEAgCode=" & ToAgCode _
& ",CompECGCode=" & ToCGCode & ",CompEItCode=" & ToItCode & ", CompENarrCode=" & ToNarrCode _
& " where CompCode='" & gCCode & "' and CompYear='" & gDbYr & "'"
gSetupCn.Execute "insert into tblCompanyYear (CompCode,CompYear,CompFdt,CompTdt,CompSAcCode,CompSAgCode,CompSCGCode,CompSItCode,CompSNarrCode)" _
& " values ('" & gCCode & "','" & NextCYr & "','" & DateAdd("yyyy", 1, gCYSDate) & "','" & DateAdd("yyyy", 1, gCYEDate) & "'," & (ToAcCode + 1) _
& "," & (ToAgCode + 1) & "," & (ToCGCode + 1) & "," & (ToItCode + 1) & "," & (ToNarrCode + 1) & ")"

If gBackEndDB = gBackEndAccess Then
    GProcConnectionOpen gCnNextYr, NextDbNm
Else
    GProcConnectionOpenORA gCnNextYr, NextDbNm
End If
'== Delete Transactions
    gCnNextYr.Execute "delete from tblVoucher" '- Voucher
    gCnNextYr.Execute "delete from TblMillRecPay" '-- Mill Vouchers (Mill Rec Pay)
    gCnNextYr.Execute "delete from tblRecVsSale" '-- Rec Vs Sale
    gCnNextYr.Execute "delete from tblClosingStock " '- Cl. Stock
    gCnNextYr.Execute "delete from tblConfirmation " '- Sale Confirmation
    gCnNextYr.Execute "delete from tblTDS" '- TDS
    '---
    'gCn.Execute "delete From tblBags where InwType is null and InwNo=0 and GPNo=" & txt(0) & " and GpType='GP'"
    '-- delete Purchase Return,Sale Return,Purchase Other
    gCnNextYr.Execute "delete from TblVouSub where vtype in ('RY','RO','VY','VI','PO')"
    gCnNextYr.Execute "delete from tblAddLess where vtype in ('RY','RO','VY','VI','PO')"
'-----To delete Last  years os details in tblVousub
    gCnNextYr.Execute "delete from TblVouSub where vtype in ('SY','ST','SD','SO','SM') and  vyear <= " & gCYear & "  "
    gCnNextYr.Execute "delete from TblAddLess where vtype in ('SY','ST','SD','SO','SM') and  vyear <= " & gCYear & "  "
'-----
    '--- delete GatePass
    gCnNextYr.Execute "delete from TblVouSub where vtype in ('GP')" '- Gate Pass
    gCnNextYr.Execute "delete from tblAddLess where vtype in ('GP')"
    
    '--- delete GatePass
    gCnNextYr.Execute "delete from tblvousub_Book where vtype in ('OP')" '- Gate Pass
    gCnNextYr.Execute "delete from tblAddLess_Book where vtype in ('OP')"
    
    '--- delete Completed Booking
    gCnNextYr.Execute "delete from tblVouSub where Vtype in ('OM','OS','OI') and " _
    & " (VSubItCtrlNo<0  or VSubBookIsCompleted=1 or (VSubBag-VSubIssBag)<=0)"
    gCnNextYr.Execute "delete from tbladdless where Vtype in ('OM') and Vno not in (select tblVousub.vno from tblVousub " _
    & "  where tblVousub.Vtype in ('OM'))" ','OS','OI'
    gCnNextYr.Execute "delete from tbladdless where Vtype in ('OS') and Vno not in (select tblVousub.vno from tblVousub " _
    & "  where tblVousub.Vtype in ('OS'))"
    gCnNextYr.Execute "delete from tbladdless where Vtype in ('OI') and Vno not in (select tblVousub.vno from tblVousub " _
    & "  where tblVousub.Vtype in ('OI'))"
        '--- Update Pending Booking Balance
        gCnNextYr.Execute "update tblVouSub set VSubItCtrlNo=-VSubItCtrlNo,VSubBag=VSubBag-VSubIssBag, " _
        & " VSubWt=VSubWt-VSubIssWt,VSubIssBag=0,VSubIssWt=0 where Vtype in ('OM','OS','OI') " _
    '--  Delete Completed Purchase
    gCnNextYr.Execute "delete from tblVouSub where Vtype in ('PY','PT','PI')  and " _
    & " (VSubItCtrlNo<0  or  (VSubBag-VSubIssBag)<=0)"
    gCnNextYr.Execute "delete from tbladdless where Vtype in ('PY') and Vno not in (select tblVousub.vno from tblVousub " _
    & " where tblVousub.Vtype in ('PY') )"
    gCnNextYr.Execute "delete from tbladdless where Vtype in ('PT') and Vno not in (select tblVousub.vno from tblVousub " _
    & " where tblVousub.Vtype in ('PT') )"
    gCnNextYr.Execute "delete from tbladdless where Vtype in ('PI') and Vno not in (select tblVousub.vno from tblVousub " _
    & " where tblVousub.Vtype in ('PI') )"
        '--- Update Pending Purchase Balance
    gCnNextYr.Execute "update tblVouSub set VSubItCtrlNo=-VSubItCtrlNo,VSubPItCtrlNo=-VSubPItCtrlNo,VSubBkItCtrlNo=-VSubBkItCtrlNo," _
    & " VSubBag=VSubBag-VSubIssBag,VSubWt=VSubWt-VSubIssWt,VSubIssBag=0,VSubIssWt=0 where Vtype in ('PY','PT','PI')"
    '-- delete Other
    'gCnNextYr.Execute "delete from TblVouSub where vtype in ('SY','ST','SO','SD','SM')"
    'gCnNextYr.Execute "delete from tblAddLess where vtype in ('SY','ST','SO','SD','SM')"
    '-- Delete Completed Outstanding
    gCnNextYr.Execute "delete from tblOutStanding where (OutAmount-OutRecAmt)<=0 and (OutDbNtAmt-OutRecDbAmt)<=0"
    '--- Update Pending Outstanding Balance
    gCnNextYr.Execute "update tblOutStanding set OutAmount=OutAmount-OutRecAmt,OutDbNtAmt=OutDbNtAmt-OutRecDbAmt," _
    & "OutRecAmt=0,OutRecDbAmt=0,OutRecFreight=0"
     
    '--- Delete Purchase Outstanding
     gCnNextYr.Execute "delete from tblOutStanding where vtype in ('PY','PT','PO','PI') and vyear=" & gCYear
     
     
     '-- Delete Completed Outstanding Sale
     gCnNextYr.Execute "delete from TblVouSub where vtype in ('SY','ST','SO','SD','SM') and vno not in (select vno from " _
     & " tblOutStanding where vyear=" & gCYear & " and  vtype in ('SY','ST','SO','SD','SM')) and vyear=" & gCYear
     gCnNextYr.Execute "delete from TblAddLess where vtype in ('SY','ST','SO','SD','SM') and vno not in (select vno from " _
     & " tblVouSub where vyear=" & gCYear & " and  vtype in ('SY','ST','SO','SD','SM')) and vyear=" & gCYear

'==== Carry Forward Balances
'GProcUtlCarryForward NextDbNm
MsgBox "Utility Year End is Completed Successfully.", vbOKOnly + vbInformation, "Utility"
Exit Sub
ErrorRoutine:
    GProcErrorHandler
End Sub
'======== Utility Carry Forword Procedure ==============
Public Sub GProcUtlCarryForward(ToDbNm As String)
'On Error GoTo ErrorRoutine
Dim i As Long
Dim TblMstCount As Integer
Dim gCnToYr As Connection '- Connection For Next Yr
Dim Vno As Long
Dim vamt As Double
Dim VYear As Long
Dim AcOCode As Long
Dim rsBalS As Recordset '- Balance Source Recordset
Dim rsBalS1 As Recordset
Dim rsBalT As Recordset '- Balance Target Recordset
Dim rsBalT1 As Recordset '- Balance Target Recordset
Dim X As String
Dim DtStr As String
GProcConnectionOpen gCnToYr, ToDbNm
'=================== Master (No of tables = 4) New Transaction Entry ================================
For TblMstCount = 0 To 4
    Select Case TblMstCount
        Case 0 '- Comp Group
            X = " Select  *  from tblMastCompGroup "
        Case 1 '-A/c Group
            X = " Select  *  from tblMastGroup "
        Case 2 '- A/c
            X = " Select  *  from tblMastAccount "
        Case 3 '- Item
            X = " Select  *  from tblMastItem "
        Case 4 '- Narration
            X = " Select  *  from tblMastNarration "
    End Select
    If GProcRstOpen(rsBalS, X, "R") > 0 Then
        With rsBalS
            Do While Not .EOF
                X = ""
                Select Case TblMstCount
                    Case 0 '- Comp Group
                        X = " Select  *  from tblMastCompGroup where CGCode=" & .Fields("CGCode")
                    Case 1 '-A/c Group
                        X = " Select  *  from tblMastGroup where AgCode=" & .Fields("AgCode")
                    Case 2 '- A/c
                        X = " Select  *  from tblMastAccount where AcCode=" & .Fields("AcCode")
                    Case 3 '- Item
                        X = " Select  *  from tblMastItem where ItCode=" & .Fields("ItCode")
                    Case 4 '- Narration
                        X = " Select  *  from tblMastNarration where NarrCode=" & .Fields("NarrCode")
                End Select
                If GProcRstOpen(rsBalT, X, "O", gCnToYr) = 0 Then '--- Add Master Entry To Next Year
                    rsBalT.AddNew
                    For i = 0 To .Fields.Count - 1
                        rsBalT.Fields(i) = .Fields(i)
                    Next i
                    rsBalT.Update
                End If
                .MoveNext
            Loop
        End With
    End If
Next TblMstCount
'=============== Next Year  Delete All Opening Bal
gCnToYr.Execute "delete from tblVoucher where Vtype='0P'"
AcOCode = GProcGetColumnValue("TblMastAccount", "AcName", "-", "S", "AcCode", "N") '-- "-" A/c
VYear = CLng(gCYear) + 1
'===================  Next Year Closing Stock Balances =====================
'- Carry forward 31/03 cl. stock to Next Yr
If gBackEndDB = gBackEndAccess Then
    DtStr = "StDate=cdate('" & gCYEDate & "')"
Else
    DtStr = "StDate=to_date('" & gCYEDate & "','dd/MM/yyyy')"
End If
X = "select * from tblClosingStock where " & DtStr
Vno = 0
If GProcRstOpen(rsBalS, X, "R") > 0 Then
    With rsBalS
        Vno = Vno + 1
        If gBackEndDB = gBackEndAccess Then
            DtStr = "cdate('" & gCYEDate + 1 & "')"
        Else
            DtStr = "to_date('" & gCYEDate + 1 & "','dd/MM/yyyy')"
        End If
        gCnToYr.Execute "insert into tblVoucher (VNo,VCtrNo,VDt,VType,VAcCode,VAcOCode,VAmt,vUserName,vYear)  Values " _
        & "(" & Vno & ",1," & DtStr & ",'0P'," & .Fields("StAcCode") & "," & AcOCode & ",-" & .Fields("StAmount") & ",'" & gUserName & "'," & VYear & ")"
    End With
End If
'===================  Next Year Account Opening Balances =====================
gCn.Execute "delete from tmpselection where UserName='" & gUserName & "'"
If gBackEndDB = gBackEndAccess Then
    gCn.Execute "insert into tmpselection (dt1,Dt2,UserName) values ('" & gCYSDate & "','" & gCYEDate & "','" & gUserName & "')"
Else
    gCn.Execute "insert into tmpselection (dt1,dt2,UserName) values (to_date('" & gCYSDate & "','dd/MM/yyyy'),to_Date('" & gCYEDate & "','dd/MM/yyyy'),'" & gUserName & "') "
End If
If gBackEndDB = gBackEndAccess Then
    DtStr = "cdate('" & gCYEDate + 1 & "')"
Else
    DtStr = "to_date('" & gCYEDate + 1 & "','dd/MM/yyyy')"
End If
X = "select * from qryAccountBalance where TopGroup in (90001,90002)  and  UserName='" & gUserName & "'"
If GProcRstOpen(rsBalS, X, "R") > 0 Then
    With rsBalS
        Do While Not .EOF
            Vno = Vno + 1
            If .Fields("Debit") > 0 Then
                vamt = CDbl(.Fields("Debit") * -1)
            Else
                vamt = CDbl(.Fields("Credit"))
            End If
            gCnToYr.Execute "insert into tblVoucher (VNo,VCtrNo,VDt,VType,VAcCode,VAcOCode,VAmt,vUserName,vYear)  Values " _
            & "(" & Vno & ",1," & DtStr & ",'0P'," & .Fields("AcCode") & "," & AcOCode & "," & vamt & ",'" & gUserName & "'," & VYear & ")"
            .MoveNext
        Loop
    End With
End If
'================  Booking Balances ===============================
'X = " Select  *  from tblVouSub where Vtype in ('OM','OS','OI')  and " _
'& "  VSubItCtrlNo>0 and VSubBookIsCompleted=0 and (VSubBag-VSubIssBag)>0"
'If GProcRstOpen(rsBalS, X, "R") > 0 Then
'    With rsBalS
'        Do While Not .EOF
'            X = "Select * from Tblvousub Where VYear=" & .Fields("VYear") & " and Vtype='" & .Fields("VType") & "' and VSubItCtrlNo=-" & .Fields("VSubItCtrlNo")
'            If GProcRstOpen(rsBalT, X, "O", gCnToYr) > 0 Then '--- Update TblVouSub Balance
'                rsBalT!vsubBAG = .Fields("VSubBag") - .Fields("VSubIssBag") '- Bag
'                rsBalT!vsubwt = .Fields("VSubWt") - .Fields("VSubIssWt") '-Wt
'                rsBalT.Update
'            Else
'                X = "Select * from Tbladdless Where VYear=" & .Fields("VYear") & " and VType='" & .Fields("VType") & "' and VNo=" & .Fields("VNo")
'                GProcRstOpen rsBalS1, X, "R"
'                X = "Select * from Tbladdless Where VYear=" & .Fields("VYear") & " and VType='" & .Fields("VType") & "' and VNo=" & .Fields("VNo")
'                GProcRstOpen rsBalT1, X, "O", gCnToYr
'                rsBalT1.AddNew
'                For i = 0 To rsBalS1.Fields.Count - 1 '- Add TblAddLess Entry
'                    rsBalT1.Fields(i) = rsBalS1.Fields(i)
'                Next i
'                rsBalT1.Update
'                rsBalT.AddNew
'                For i = 0 To .Fields.Count - 1 '- Add TblVouSub Entry
'                    rsBalT.Fields(i) = .Fields(i)
'                Next i
'                'rsBalT.Update
'                rsBalT!vsubBAG = .Fields("VSubBag") - .Fields("VSubIssBag") '- Bag
'                rsBalT!vsubwt = .Fields("VSubWt") - .Fields("VSubIssWt") '-Wt
'                rsBalT!VSubItCtrlNo = -1 * .Fields("VSubItCtrlNo") '- It Ctrl No
'                rsBalT!VSubPItCtrlNo = -1 * .Fields("VSubPItCtrlNo") '-P It Ctrl No
'                rsBalT.Update
'            End If
'            .MoveNext
'        Loop
'    End With
'End If
'================  Purchase Balances ===============================
'X = " Select  *  from tblVouSub where Vtype in ('PY','PT','PI')  and " _
'& "  VSubItCtrlNo>0 and (VSubBag-VSubIssBag)>0"
'If GProcRstOpen(rsBalS, X, "R") > 0 Then
'    With rsBalS
'        Do While Not .EOF
'            X = "Select * from Tblvousub Where VYear=" & .Fields("VYear") & " and Vtype='" & .Fields("VType") & "' and VSubItCtrlNo=-" & .Fields("VSubItCtrlNo")
'            If GProcRstOpen(rsBalT, X, "O", gCnToYr) > 0 Then '--- Update TblVouSub Balance
'                rsBalT!vsubBAG = .Fields("VSubBag") - .Fields("VSubIssBag") '- Bag
'                rsBalT!vsubwt = .Fields("VSubWt") - .Fields("VSubIssWt") '-Wt
'                rsBalT.Update
'            Else
'                X = "Select * from Tbladdless Where VYear=" & .Fields("VYear") & " and VType='" & .Fields("VType") & "' and VNo=" & .Fields("VNo")
'                GProcRstOpen rsBalS1, X, "R"
'                X = "Select * from Tbladdless Where VYear=" & .Fields("VYear") & " and VType='" & .Fields("VType") & "' and VNo=" & .Fields("VNo")
'                GProcRstOpen rsBalT1, X, "O", gCnToYr
'                rsBalT1.AddNew
'                For i = 0 To rsBalS1.Fields.Count - 1 '- Add TblAddLess Entry
'                    rsBalT1.Fields(i) = rsBalS1.Fields(i)
'                Next i
'                rsBalT1.Update
'                rsBalT.AddNew
'                For i = 0 To .Fields.Count - 1 '- Add TblVouSub Entry
'                    rsBalT.Fields(i) = .Fields(i)
'                Next i
'                'rsBalT.Update
'                rsBalT!vsubBAG = .Fields("VSubBag") - .Fields("VSubIssBag") '- Bag
'                rsBalT!vsubwt = .Fields("VSubWt") - .Fields("VSubIssWt") '-Wt
'                rsBalT!VSubItCtrlNo = -1 * .Fields("VSubItCtrlNo") '- It Ctrl No
'                rsBalT!VSubPItCtrlNo = -1 * .Fields("VSubPItCtrlNo") '-P It Ctrl No
'                rsBalT.Update
'            End If
'            .MoveNext
'        Loop
'    End With
'End If
''-- Delete Purchase in Next Year for (Complete & No Vs GatePass in Next Year)
'X = " Select  *  from tblVouSub where Vtype in ('PY','PT','PI')  and VSubItCtrlNo>0 and (VSubBag-VSubIssBag)<=0 "
'If GProcRstOpen(rsBalS, X, "R") > 0 Then
'    With rsBalS
'        Do While Not .EOF
'            gCnToYr.Execute "delete from tblVouSub where Vno not in ( select vno from tblVouSub as tblGp where " _
'            & " VSubPItCtrlNo=-" & .Fields("VSubItCtrlNo") & " and  tblGp.Vtype='GP' and  tblGp.VYear=" & .Fields("VYear") & ") " _
'            & " and VYear=" & .Fields("VYear") & " and Vtype='" & .Fields("VType") & "' and VNo=" & .Fields("VNo")
'            rsBalS.MoveNext
'        Loop
'    End With
'End If
'================  Outstanding Balances ===============================
X = " Select  *  from tblOutStanding where ((OutAmount-OutRecAmt)>0  or " _
& "  (OutDbNtAmt-OutRecDbAmt)>0)"
If GProcRstOpen(rsBalS, X, "R") > 0 Then
    With rsBalS
        Do While Not .EOF
            X = "Select * from tblOutStanding Where VYear=" & .Fields("VYear") & " and Vtype='" & .Fields("VType") & "' and VNo=" & .Fields("VNo")
            If GProcRstOpen(rsBalT, X, "O", gCnToYr) = 0 Then
                rsBalT.AddNew
                For i = 0 To .Fields.Count - 1 '- Add TblOutstanding Entry
                    rsBalT.Fields(i) = .Fields(i)
                Next i
            End If '--- Update TblOutstanding Balance
            rsBalT!OutAmount = .Fields("OutAmount") - .Fields("OutRecAmt") '- Out Amt
            rsBalT!OutDbNtAmt = .Fields("OutDbNtAmt") - .Fields("OutRecDbAmt") '-DbNt Amt
            rsBalT!OutRecAmt = 0 '-Rec. Amt
            rsBalT!OutRecDbAmt = 0 '-Rec. DbNt Amt
            rsBalT!OutRecFreight = 0 '-Rec. Freight Amt
            rsBalT.Update
            .MoveNext
        Loop
    End With
End If
'-- Delete Outstanding in Next Year for (Complete & No Vs Receipt in Next Year)
X = " Select  *  from tblOutStanding where ((OutAmount-OutRecAmt)<=0  and " _
& "  (OutDbNtAmt-OutRecDbAmt)<=0)"
If GProcRstOpen(rsBalS, X, "R") > 0 Then
    With rsBalS
        Do While Not .EOF
            gCnToYr.Execute "delete from tblOutStanding where Vno not in ( select vno from tblRecVsSale where " _
            & " InvNo=" & .Fields("VNo") & " and  InvTp='" & .Fields("VType") & "' and  InvYear=" & .Fields("VYear") & ") " _
            & " and VYear=" & .Fields("VYear") & " and Vtype='" & .Fields("VType") & "' and VNo=" & .Fields("VNo")
            rsBalS.MoveNext
        Loop
    End With
End If
MsgBox "Utility Carry Forward Balances is Completed Successfully.", vbOKOnly + vbInformation, "Utility"
Exit Sub
ErrorRoutine:
    GProcErrorHandler
End Sub
'===== Update It Ctrl Voucher Sub Stock
Public Sub GProcUpdateItCtrlStock(pItCtrlNo As Long, pBag As Long, pWt As Double, Optional pModFlag As Boolean, Optional pDeleteFlag As Boolean)
If pDeleteFlag = True Then '--Delete
        gCn.Execute "update tblVousub set VSubIssBag=(VSubIssBag-" & pBag & "),VSubIssWt=(VSubIssWt-" & pWt & ")" _
        & " ,VSubTmpBag=0,VSubTmpWt=0 where VSubItCtrlNo=" & pItCtrlNo
Else
    If pModFlag = True Then '--Modify
        gCn.Execute "update tblVousub set VSubTmpBag=" & pBag & ",VSubTmpWt=" & pWt & " where VSubItCtrlNo=" & pItCtrlNo
    Else '--Update Stock
        gCn.Execute "update tblVousub set VSubIssBag=(VSubIssBag-VSubTmpBag+" & pBag & "),VSubIssWt=(VSubIssWt-VSubTmpWt+" & pWt & ")" _
        & " ,VSubTmpBag=0,VSubTmpWt=0 where VSubItCtrlNo=" & pItCtrlNo
    End If
End If
End Sub
'==== Update Setup Booking Sub It Ctrl Stock
Public Sub GProcUpdateSetupBkItCtrlStock(pItCtrlNo As Long, pBag As Long, pWt As Double, Optional pModFlag As Boolean, Optional pDeleteFlag As Boolean)
If pDeleteFlag = True Then '--Delete
        gSetupCn.Execute "update tblBookingSub set BkSubIssBag=(BkSubIssBag-" & pBag & "),BkSubIssWt=(BkSubIssWt-" & pWt & ")" _
        & " ,BkSubTmpBag=0,BkSubTmpWt=0 where BkSubItCtrlNo=" & pItCtrlNo
Else
    If pModFlag = True Then '--Modify
        gSetupCn.Execute "update tblBookingSub set BkSubTmpBag=" & pBag & ",BkSubTmpWt=" & pWt & " where BkSubItCtrlNo=" & pItCtrlNo
    Else '--Update Stock
        gSetupCn.Execute "update tblBookingsub set BkSubIssBag=(BkSubIssBag-BkSubTmpBag+" & pBag & "),BkSubIssWt=(BkSubIssWt-BkSubTmpWt+" & pWt & ")" _
        & " ,BkSubTmpBag=0,BkSubTmpWt=0 where BkSubItCtrlNo=" & pItCtrlNo
    End If
End If
End Sub
'==== Update Voucher Sub Stock
Public Sub GProcUpdateVSubStock(pOutVno As Long, POutVYear As Long, pOutVType As String, Optional pModFlag As Boolean, Optional pDeleteFlag As Boolean)
If pDeleteFlag = True Then '--Delete
    If gBackEndDB = gBackEndAccess Then '-Access
        gCn.Execute "update tblVousub,tblvousub tblOutSub set tblVousub.VSubIssBag=tblVousub.VSubIssBag-tblOutSub.VSubBag," _
        & " tblVousub.VSubIssWt=tblVousub.VSubIssWt-tblOutSub.VSubWt" _
        & "  where tblOutSub.VSUBPITCTRLNO=tblVousub.VSUBITCTRLNO and tblOutSub.Vno=" & pOutVno & " and tblOutSub.Vtype='" & pOutVType & "' and  " _
        & " tblOutSub.VYear=" & POutVYear
    Else '--Oracle
        gCn.Execute "update tblVousub set VSubIssBag=VSubIssBag-nvl((select VSubBag from tblvousub tblOutSub where " _
        & " tblOutSub.VSUBPITCTRLNO=tblVousub.VSUBITCTRLNO and tblOutSub.Vno=" & pOutVno & " and tblOutSub.Vtype='" & pOutVType & "' and tblOutSub.Vyear=" & POutVYear & "),0) , " _
        & " VSubIssWt=VSubIssWt-nvl((select VSubWt from tblvousub  tblOutSub where " _
        & " tblOutSub.VSUBPITCTRLNO=tblVousub.VSUBITCTRLNO and tblOutSub.Vno=" & pOutVno & " and tblOutSub.Vtype='" & pOutVType & "' and tblOutSub.Vyear=" & POutVYear & "),0) " _
        & " where VSUBITCTRLNO in (select VSUBPITCTRLNO from tblvousub tblOutSub where tblOutSub.Vno=" & pOutVno & " and tblOutSub.Vtype='" & pOutVType & "' and tblOutSub.Vyear=" & POutVYear & ")"
    End If
Else
    If pModFlag = True Then '--Modify
        If gBackEndDB = gBackEndAccess Then '-Access
            gCn.Execute "update tblVousub,tblvousub tblOutSub set tblVousub.VSubIssBag=tblVousub.VSubIssBag-tblOutSub.VSubBag," _
            & " tblVousub.VSubIssWt=(tblVousub.VSubIssWt-tblOutSub.VSubWt)" _
            & " where tblOutSub.VSUBPITCTRLNO=tblVousub.VSUBITCTRLNO and tblOutSub.Vno=" & pOutVno & " and tblOutSub.Vtype='" & pOutVType & "' and  " _
            & " tblOutSub.VYear=" & POutVYear
        Else '--Oracle
            gCn.Execute "update tblVousub set VSubIssBag=VSubIssBag-nvl((select sum(VSubBag) from tblvousub tblOutSub where " _
            & " tblOutSub.VSUBPITCTRLNO=tblVousub.VSUBITCTRLNO and tblOutSub.Vno=" & pOutVno & " and tblOutSub.Vtype='" & pOutVType & "' and tblOutSub.Vyear=" & POutVYear & "),0) , " _
            & " VSubIssWt=VSubIssWt-nvl((select sum(VSubWt) from tblvousub  tblOutSub where " _
            & " tblOutSub.VSUBPITCTRLNO=tblVousub.VSUBITCTRLNO and tblOutSub.Vno=" & pOutVno & " and tblOutSub.Vtype='" & pOutVType & "' and tblOutSub.Vyear=" & POutVYear & "),0)  " _
            & " where VSUBITCTRLNO in (select VSUBPITCTRLNO from tblvousub tblOutSub where tblOutSub.Vno=" & pOutVno & " and tblOutSub.Vtype='" & pOutVType & "' and tblOutSub.Vyear=" & POutVYear & ")"
        End If
    Else '--Update Stock
        If gBackEndDB = gBackEndAccess Then '--Access
            gCn.Execute "update tblVousub,tblVousub tblOutSub set tblVousub.VSubIssBag=tblVousub.VSubIssBag+tblOutSub.VSubBag,tblVousub.VSubIssWt=tblVousub.VSubIssWt+tblOutSub.VSubWt " _
            & " where tblOutSub.VSUBPITCTRLNO=tblVousub.VSUBITCTRLNO and tblOutSub.Vno=" & pOutVno & " and tblOutSub.Vtype='" & pOutVType & "' and  " _
            & " tblOutSub.VYear=" & POutVYear
        Else
            gCn.Execute "update tblVousub set VSubIssBag=VSubIssBag+ nvl((select VSubBag from tblvousub tblOutSub where " _
            & " tblOutSub.VSUBPITCTRLNO=tblVousub.VSUBITCTRLNO and tblOutSub.Vno=" & pOutVno & " and tblOutSub.Vtype='" & pOutVType & "' and tblOutSub.Vyear=" & POutVYear & "),0) , " _
            & " VSubIssWt=VSubIssWt+nvl((select VSubWt from tblvousub  tblOutSub where " _
            & " tblOutSub.VSUBPITCTRLNO=tblVousub.VSUBITCTRLNO and tblOutSub.Vno=" & pOutVno & " and tblOutSub.Vtype='" & pOutVType & "' and tblOutSub.Vyear=" & POutVYear & "),0)" _
            & " where VSUBITCTRLNO in (select VSUBPITCTRLNO from tblvousub tblOutSub where tblOutSub.Vno=" & pOutVno & " and tblOutSub.Vtype='" & pOutVType & "' and tblOutSub.Vyear=" & POutVYear & ")"
        End If
    End If
End If
'-- Update Null To 0
 If gBackEndDB = gBackEndAccess Then '-Access
    gCn.Execute "update tblVousub set VSubIssBag=0 where VSubIssBag is Null "
    gCn.Execute "update tblVousub set VSubIssWt=0 where VSubIssWt is Null "
 Else
    gCn.Execute "update tblVousub set VSubIssBag=0 where VSubIssBag is Null"
    gCn.Execute "update tblVousub set VSubIssWt=0 where VSubIssWt is Null "
 End If
End Sub
''-----Update Voucher Sub Stock
'Public Sub GProcUpdateVSubStock(pVno As Long, pVdt As Date, pVtype As String, pOutVno As Long, POutVYear As Long, pOutVType As String, Optional pModFlag As Boolean, Optional pDeleteFlag As Boolean)
'If pDeleteFlag = True Then '--Delete
'    If gBackEndDB = gBackEndAccess Then '-Access
'        gCn.Execute "update tblVousub,tblOutSub set tblVousub.VSubIssBag=tblVousub.VSubIssBag-tblOutSub.VSubBag," _
'        & " tblVousub.VSubIssWt=tblVousub.VSubIssWt-tblOutSub.VSubWt" _
'        & " from tblvousub,tblvousub tblOutSub where tblOutSub.VSUBPITCTRLNO=tblVousub.VSUBITCTRLNO and tblOutSub.Vno=" & pOutVno & " and tblOutSub.Vtype='" & pOutVType & "' and  " _
'        & " tblOutSub.VYear=" & POutVYear & " and tblVousub.Vno=" & pVno & " and tblVousub.Vtype='" & pVtype & "' and tblVousub.Vdt=cdate('" & pVdt & "')"
'    Else '--Oracle
'        gCn.Execute "update tblVousub set VSubIssBag=VSubIssBag-nvl((select VSubBag from tblvousub tblOutSub where " _
'        & " tblOutSub.VSUBPITCTRLNO=tblVousub.VSUBITCTRLNO and tblOutSub.Vno=" & pOutVno & " and tblOutSub.Vtype='" & pOutVType & "' and tblOutSub.Vyear=" & POutVYear & "),0) , " _
'        & " VSubIssWt=VSubIssWt-nvl((select VSubWt from tblvousub  tblOutSub where " _
'        & " tblOutSub.VSUBPITCTRLNO=tblVousub.VSUBITCTRLNO and tblOutSub.Vno=" & pOutVno & " and tblOutSub.Vtype='" & pOutVType & "' and tblOutSub.Vyear=" & POutVYear & "),0) " _
'        & " where Vno=" & pVno & " and Vtype='" & pVtype & "' and Vdt=to_date('" & pVdt & "','dd/MM/yyyy')"
'    End If
'Else
'    If pModFlag = True Then '--Modify
'        If gBackEndDB = gBackEndAccess Then '-Access
'            gCn.Execute "update tblVousub,tblOutSub set tblVousub.VSubIssBag=tblVousub.VSubIssBag-tblOutSub.VSubBag," _
'            & " tblVousub.VSubIssWt=tblVousub.VSubIssWt-tblOutSub.VSubWt" _
'            & " from tblvousub,tblvousub tblOutSub where tblOutSub.VSUBPITCTRLNO=tblVousub.VSUBITCTRLNO and tblOutSub.Vno=" & pOutVno & " and tblOutSub.Vtype='" & pOutVType & "' and  " _
'            & " tblOutSub.VYear=" & POutVYear & " and tblVousub.Vno=" & pVno & " and tblVousub.Vtype='" & pVtype & "' and tblVousub.Vdt=cdate('" & pVdt & "')"
'        Else '--Oracle
'            gCn.Execute "update tblVousub set VSubIssBag=VSubIssBag-nvl((select VSubBag from tblvousub tblOutSub where " _
'            & " tblOutSub.VSUBPITCTRLNO=tblVousub.VSUBITCTRLNO and tblOutSub.Vno=" & pOutVno & " and tblOutSub.Vtype='" & pOutVType & "' and tblOutSub.Vyear=" & POutVYear & "),0) , " _
'            & " VSubIssWt=VSubIssWt-nvl((select VSubWt from tblvousub  tblOutSub where " _
'            & " tblOutSub.VSUBPITCTRLNO=tblVousub.VSUBITCTRLNO and tblOutSub.Vno=" & pOutVno & " and tblOutSub.Vtype='" & pOutVType & "' and tblOutSub.Vyear=" & POutVYear & "),0)  " _
'            & " where Vno=" & pVno & " and Vtype='" & pVtype & "' and Vdt=to_date('" & pVdt & "','dd/MM/yyyy')"
'        End If
'    Else '--Update Stock
'        If gBackEndDB = gBackEndAccess Then '--Access
'            gCn.Execute "update tblVousub,tblOutSub set tblVousub.VSubIssBag=tblVousub.VSubIssBag+tblOutSub.VSubBag,tblVousub.VSubIssWt=tblVousub.VSubIssWt+tblOutSub.VSubWt " _
'            & " from tblvousub,tblvousub tblOutSub where tblOutSub.VSUBPITCTRLNO=tblVousub.VSUBITCTRLNO and tblOutSub.Vno=" & pOutVno & " and tblOutSub.Vtype='" & pOutVType & "' and  " _
'            & " tblOutSub.VYear=" & POutVYear & " and tblVousub.Vno=" & pVno & " and tblVousub.Vtype='" & pVtype & "' and tblVousub.Vdt=Cdate('" & pVdt & "')"
'        Else
'            gCn.Execute "update tblVousub set VSubIssBag=VSubIssBag+ nvl((select VSubBag from tblvousub tblOutSub where " _
'            & " tblOutSub.VSUBPITCTRLNO=tblVousub.VSUBITCTRLNO and tblOutSub.Vno=" & pOutVno & " and tblOutSub.Vtype='" & pOutVType & "' and tblOutSub.Vyear=" & POutVYear & "),0) , " _
'            & " VSubIssWt=VSubIssWt+nvl((select VSubWt from tblvousub  tblOutSub where " _
'            & " tblOutSub.VSUBPITCTRLNO=tblVousub.VSUBITCTRLNO and tblOutSub.Vno=" & pOutVno & " and tblOutSub.Vtype='" & pOutVType & "' and tblOutSub.Vyear=" & POutVYear & "),0)" _
'            & " where Vno=" & pVno & " and Vtype='" & pVtype & "' and Vdt=to_date('" & pVdt & "','dd/MM/yyyy')"
'        End If
'    End If
'End If
''-- Update Null To 0
' If gBackEndDB = gBackEndAccess Then '-Access
'    gCn.Execute "update tblVousub set VSubIssBag=0 where VSubIssBag is Null and Vno=" & pVno & " and Vtype='" & pVtype & "' and tblVousub.Vdt=Cdate('" & pVdt & "')"
'    gCn.Execute "update tblVousub set VSubIssWt=0 where VSubIssWt is Null and Vno=" & pVno & " and Vtype='" & pVtype & "' and tblVousub.Vdt=Cdate('" & pVdt & "')"
' Else
'    gCn.Execute "update tblVousub set VSubIssBag=0 where VSubIssBag is Null and Vno=" & pVno & " and Vtype='" & pVtype & "' and Vdt=to_date('" & pVdt & "','dd/MM/yyyy')"
'    gCn.Execute "update tblVousub set VSubIssWt=0 where VSubIssWt is Null and Vno=" & pVno & " and Vtype='" & pVtype & "' and Vdt=to_date('" & pVdt & "','dd/MM/yyyy')"
' End If
'End Sub
'''=== Update Outstanding Balance
''Public Sub GProcUpdateOutstandingBal(RecVno As Long, RecVType As String, RecYear As Long, Optional pModFlag As Boolean, Optional pDeleteFlag As Boolean)
''Dim RsInv As Recordset
''Dim X As String
''Dim Y As String
''Dim rsOS As Recordset
''If pDeleteFlag = True Then '--Delete
''   gCn.Execute "update tblOutstanding,tblRecVsSale set OutRecAmt=(OutRecAmt-(tblRecVsSale.AdjAmt+tblRecVsSale.Discount)),OutRecDbAmt=(OutRecDbAmt-tblRecVsSale.AdjDbNtAmt)" _
''   & "  where tblOutstanding.Vno=tblRecVsSale.InvNo and tblOutstanding.Vtype=tblRecVsSale.InvTp and " _
''   & " tblOutstanding.VYear=tblRecVsSale.InvYear and " _
''   & " RecVno=" & RecVno & " and RecVtype='" & RecVType & "' and RecVyear=" & RecYear
''Else
'''---TDS-194Q
''   If RecVType = "BP" Then
''      gCn.Execute "update tblOutstanding,tblRecVsSale set OutRecAmt=OutRecAmt-tblRecVsSale.AdjAmt,OutRecDbAmt=OutRecDbAmt-tblRecVsSale.AdjDbNtAmt where " _
''      & " tblOutstanding.Vno=tblRecVsSale.InvNo and tblOutstanding.Vtype=tblRecVsSale.InvTp and " _
''      & " tblOutstanding.VYear=tblRecVsSale.InvYear and " _
''      & " RecVno=" & RecVno & " and RecVtype='" & RecVType & "' and RecVyear=" & RecYear & " and tblrecVsSale.Invtp <> 'SM'"
''
''   Else
''      X = "select * from tblRecVsSale where RecVno=" & RecVno & " and RecVtype='" & RecVType & "' and RecVyear=" & RecYear
''      If GProcRstOpen(RsInv, X, "R") > 0 Then
''          RsInv.MoveFirst
''          Do While Not RsInv.EOF
''             gCn.Execute "update tblOutstanding set OutRecAmt=0 ,OutRecDbAmt= 0  " _
''             & " where Vno=" & RsInv!InvNo & " and " _
''             & "  Vtype='" & RsInv!InvTp & "' And VYear = " & RsInv!InvYear
''             RsInv.MoveNext
''          Loop
''
''          RsInv.MoveFirst
''          Do While Not RsInv.EOF
''             Y = "select sum(AdjAmt+Discount),sum(AdjDbNtAmt) from tblRecVsSale " _
''                 & " where InvNo=" & RsInv!InvNo & " and " _
''                 & "  InvTp='" & RsInv!InvTp & "' And InvYear = " & RsInv!InvYear
''
''             If GProcRstOpen(rsOS, Y, "R") > 0 Then
''                gCn.Execute "update tblOutstanding set OutRecAmt= " & rsOS.Fields(0) & "  , " _
''                & " OutRecDbAmt= " & rsOS.Fields(1) & "   " _
''                & " where Vno=" & RsInv!InvNo & " and " _
''                & "  Vtype='" & RsInv!InvTp & "' And VYear = " & RsInv!InvYear
''             End If
''             RsInv.MoveNext
''          Loop
''      End If
''   End If
'''--End TDS-194Q
''End If
''End Sub
''Public Sub GProcUpdateOutstandingBal(RecVno As Long, RecVType As String, RecYear As Long, Optional pModFlag As Boolean, Optional pDeleteFlag As Boolean)
''Dim RsInv As Recordset
''Dim X As String
''Dim Y As String
''Dim rsOS As Recordset
''If pDeleteFlag = True Then '--Delete
''   gCn.Execute "update tblOutstanding,tblRecVsSale set OutRecAmt=(OutRecAmt-(tblRecVsSale.AdjAmt+tblRecVsSale.Discount)),OutRecDbAmt=(OutRecDbAmt-tblRecVsSale.AdjDbNtAmt)" _
''   & "  where tblOutstanding.Vno=tblRecVsSale.InvNo and tblOutstanding.Vtype=tblRecVsSale.InvTp and " _
''   & " tblOutstanding.VYear=tblRecVsSale.InvYear and " _
''   & " RecVno=" & RecVno & " and RecVtype='" & RecVType & "' and RecVyear=" & RecYear
''Else
'''---TDS-194Q
'''   If RecVType = "BP" Then
'''      gCn.Execute "update tblOutstanding,tblRecVsSale set OutRecAmt=OutRecAmt-tblRecVsSale.AdjAmt,OutRecDbAmt=OutRecDbAmt-tblRecVsSale.AdjDbNtAmt where " _
'''      & " tblOutstanding.Vno=tblRecVsSale.InvNo and tblOutstanding.Vtype=tblRecVsSale.InvTp and " _
'''      & " tblOutstanding.VYear=tblRecVsSale.InvYear and " _
'''      & " RecVno=" & RecVno & " and RecVtype='" & RecVType & "' and RecVyear=" & RecYear & " and tblrecVsSale.Invtp <> 'SM'"
'''
'''   Else
''      X = "select * from tblRecVsSale where RecVno=" & RecVno & " and RecVtype='" & RecVType & "' and RecVyear=" & RecYear
''      If GProcRstOpen(RsInv, X, "R") > 0 Then
''          RsInv.MoveFirst
''          Do While Not RsInv.EOF
''             gCn.Execute "update tblOutstanding set OutRecAmt=0 ,OutRecDbAmt= 0  " _
''             & " where Vno=" & RsInv!InvNo & " and " _
''             & "  Vtype='" & RsInv!InvTp & "' And VYear = " & RsInv!InvYear
''             RsInv.MoveNext
''          Loop
''
''          RsInv.MoveFirst
''          Do While Not RsInv.EOF
''             If pModFlag = True Then
''                Y = "select IIf(IsNull(sum(AdjAmt+Discount)), 0, sum(AdjAmt+Discount)),IIf(IsNull(sum(AdjDbNtAmt)), 0, sum(AdjDbNtAmt))  from tblRecVsSale " _
''                    & " where InvNo=" & RsInv!InvNo & " and " _
''                    & "  InvTp='" & RsInv!InvTp & "' And InvYear = " & RsInv!InvYear & " and " _
''                    & " RecVno<>" & RecVno & " and RecVtype<>'" & RecVType & "' and RecVyear<>" & RecYear
''             Else
''                Y = "select sum(AdjAmt+Discount),sum(AdjDbNtAmt) from tblRecVsSale " _
''                    & " where InvNo=" & RsInv!InvNo & " and " _
''                    & "  InvTp='" & RsInv!InvTp & "' And InvYear = " & RsInv!InvYear
''             End If
''             If GProcRstOpen(rsOS, Y, "R") > 0 Then
''                gCn.Execute "update tblOutstanding set OutRecAmt= " & rsOS.Fields(0) & "  , " _
''                & " OutRecDbAmt= " & rsOS.Fields(1) & "   " _
''                & " where Vno=" & RsInv!InvNo & " and " _
''                & "  Vtype='" & RsInv!InvTp & "' And VYear = " & RsInv!InvYear
''             End If
''             RsInv.MoveNext
''          Loop
''      End If
'''   End If
'''--End TDS-194Q
''End If
''End Sub
Public Sub GProcUpdateOutstandingBal(RecVno As Long, RecVType As String, RecYear As Long, Optional pModFlag As Boolean, Optional pDeleteFlag As Boolean)
Dim RsInv As Recordset
Dim X As String
Dim Y As String
Dim rsOS As Recordset
If pDeleteFlag = True Then '--Delete
   gCn.Execute "update tblOutstanding,tblRecVsSale set OutRecAmt=(OutRecAmt-(tblRecVsSale.AdjAmt+tblRecVsSale.Discount)),OutRecDbAmt=(OutRecDbAmt-tblRecVsSale.AdjDbNtAmt)" _
   & "  where tblOutstanding.Vno=tblRecVsSale.InvNo and tblOutstanding.Vtype=tblRecVsSale.InvTp and " _
   & " tblOutstanding.VYear=tblRecVsSale.InvYear and " _
   & " RecVno=" & RecVno & " and RecVtype='" & RecVType & "' and RecVyear=" & RecYear
Else
'---TDS-194Q
'   If RecVType = "BP" Then
'      gCn.Execute "update tblOutstanding,tblRecVsSale set OutRecAmt=OutRecAmt-tblRecVsSale.AdjAmt,OutRecDbAmt=OutRecDbAmt-tblRecVsSale.AdjDbNtAmt where " _
'      & " tblOutstanding.Vno=tblRecVsSale.InvNo and tblOutstanding.Vtype=tblRecVsSale.InvTp and " _
'      & " tblOutstanding.VYear=tblRecVsSale.InvYear and " _
'      & " RecVno=" & RecVno & " and RecVtype='" & RecVType & "' and RecVyear=" & RecYear & " and tblrecVsSale.Invtp <> 'SM'"
'
'   Else
      X = "select * from tblRecVsSale where RecVno=" & RecVno & " and RecVtype='" & RecVType & "' and RecVyear=" & RecYear
      If GProcRstOpen(RsInv, X, "R") > 0 Then
          RsInv.MoveFirst
          Do While Not RsInv.EOF
             gCn.Execute "update tblOutstanding set OutRecAmt=0 ,OutRecDbAmt= 0  " _
             & " where Vno=" & RsInv!InvNo & " and " _
             & "  Vtype='" & RsInv!InvTp & "' And VYear = " & RsInv!InvYear
             RsInv.MoveNext
          Loop

          RsInv.MoveFirst
          Do While Not RsInv.EOF
             If pModFlag = True Then
                Y = "select IIf(IsNull(sum(AdjAmt+Discount)), 0, sum(AdjAmt+Discount)),IIf(IsNull(sum(AdjDbNtAmt)), 0, sum(AdjDbNtAmt))  from tblRecVsSale " _
                    & " where InvNo=" & RsInv!InvNo & " and " _
                    & "  InvTp='" & RsInv!InvTp & "' And InvYear = " & RsInv!InvYear & " and " _
                    & " cstr(RecVno)+RecVtype+cstr(RecVyear) <> '" & CStr(RsInv!RecVno) + RsInv!RecVType + CStr(RsInv!RecVYear) & "'  "
             Else
                Y = "select sum(AdjAmt+Discount),sum(AdjDbNtAmt) from tblRecVsSale " _
                    & " where InvNo=" & RsInv!InvNo & " and " _
                    & "  InvTp='" & RsInv!InvTp & "' And InvYear = " & RsInv!InvYear
             End If
             If GProcRstOpen(rsOS, Y, "R") > 0 Then
                gCn.Execute "update tblOutstanding set OutRecAmt= " & rsOS.Fields(0) & "  , " _
                & " OutRecDbAmt= " & rsOS.Fields(1) & "   " _
                & " where Vno=" & RsInv!InvNo & " and " _
                & "  Vtype='" & RsInv!InvTp & "' And VYear = " & RsInv!InvYear
             End If
             RsInv.MoveNext
          Loop
'      End If
   End If
'--End TDS-194Q
End If
End Sub

'=== Update Balances
Public Sub GprocUtlUpdateBalance(pToDt As Date)
'On Error GoTo ErrorHandler
Dim X As String
Dim Y As String
Dim RsP As Recordset, RsTmp As Recordset
'================Outstanding Vs Receipt
gCn.BeginTrans
gCn.Execute "Update tblOutStanding set OutRecAMT = 0, OutRecDbAMT = 0"
If (GProcRstOpen(RsP, "select * from tblRecVsSale ", "O") > 0) Then
   RsP.MoveFirst
   Do While Not RsP.EOF
'--TDS-194Q  set OutRecAMT = OutRecAmt+" & RsP!AdjAmt & "+" & RsP!Discount & "
      gCn.Execute "Update tblOutStanding set OutRecAMT = OutRecAmt+" & RsP!AdjAmt & "+" & RsP!Discount & "  , OutRecDbAMT = OutRecDbAmt+" & RsP!AdjDbNtAmt & " where vtype='" & RsP!InvTp & "' and vno=" & RsP!InvNo & " and VYear=" & RsP!InvYear
      RsP.MoveNext
   Loop
End If
gCn.CommitTrans
RsP.Close
'================ It Ctrl Stock (Purchase Vs Gate Pass ,Booking Vs Sale,Purchase stock)
    If gBackEndDB = gBackEndAccess Then
        X = " Vdt<=cdate('" & pToDt & "')" 'Vdt>=cdate('" & pFrmDt & "') and
    Else '-- Oracle
        X = " Vdt<=to_date('" & CDate(pToDt) & "','dd/MM/yyyy')" 'Vdt>=to_date('" & CDate(pFrmDt) & "','dd/MM/yyyy') and
    End If
    X = X + " and vtype in ('PI','PY','PT','OM','OS','OI','OT') "
    '--- Update to p It Ctrl No Stock
    If gBackEndDB = gBackEndAccess Then
        gCn.Execute "Update tblVousub set VSubTmpBag = 0, VSubTmpWt = 0 where " & X
        If (GProcRstOpen(RsP, "select * from tblVousub where " & X, "O") > 0) Then
            RsP.MoveFirst
            Do While Not RsP.EOF '-issue bag,Wt
                 If (GProcRstOpen(RsTmp, "select sum(iif(vtype= 'RY',(VSUBBAG*-1), VSUBBAG)),sum(iif(vtype= 'RY',(VSUBWt*-1), VSUBWt)) from tblVousub where VSUBPITCTRLNO=" & RsP.Fields("VSUBITCTRLNO"), "O") > 0) Then
                 
                'If (GProcRstOpen(RsTmp, "select sum(VSUBBAG),sum(VSUBWt) from tblVousub where VSUBPITCTRLNO=" & RsP.Fields("VSUBITCTRLNO"), "O") > 0) Then
                    RsP.Fields("vsubissbag") = IIf(IsNumeric(RsTmp.Fields(0)), RsTmp.Fields(0), 0)
                    RsP.Fields("vsubisswt") = IIf(IsNumeric(RsTmp.Fields(1)), RsTmp.Fields(1), 0)
                    RsP.Update
                End If
                RsP.MoveNext
            Loop
        End If
        RsP.Close
    Else
        gCn.Execute "update tblVousub set VSubTmpBag=0,VSubTmpWt=0,vsubissbag=(select Sum (VSUBBAG) from tblVousub sub where " _
        & " sub.VSUBPITCTRLNO=tblVousub.VSUBITCTRLNO ),vsubisswt=(select Sum (VSUBWt) from tblVousub sub where " _
        & " sub.VSUBPITCTRLNO=tblVousub.VSUBITCTRLNO ) " _
        & " where " & X
    End If
    '-- update to 0 if null
    gCn.Execute "update tblVousub set vsubissbag=0 where vsubissbag is null and " & X
    gCn.Execute "update tblVousub set vsubisswt=0 where vsubisswt is null and " & X
    
'--For Party Booking in tblVousub_Booking
    X = " Vdt<=cdate('" & pToDt & "')" 'Vdt>=cdate('" & pFrmDt & "') and
    X = X + " and vtype = 'OP'"
    gCn.Execute "Update tblvousub_Book set VSubTmpBag = 0, VSubTmpWt = 0 where " & X
    If (GProcRstOpen(RsP, "select * from tblvousub_Book where " & X, "O") > 0) Then
        RsP.MoveFirst
        Do While Not RsP.EOF '-issue bag,Wt
           If (GProcRstOpen(RsTmp, "select sum(VSUBBAG),sum(VSUBWt) from tblVousub where VSubBkItCtrlNo=" & RsP.Fields("VSUBITCTRLNO"), "O") > 0) Then
              RsP.Fields("vsubissbag") = IIf(IsNumeric(RsTmp.Fields(0)), RsTmp.Fields(0), 0)
              RsP.Fields("vsubisswt") = IIf(IsNumeric(RsTmp.Fields(1)), RsTmp.Fields(1), 0)
              RsP.Update
           End If
           RsP.MoveNext
        Loop
    End If
    RsP.Close
    
    '-- update to 0 if null
    gCn.Execute "update tblvousub_Book set vsubissbag=0 where vsubissbag is null and " & X
    gCn.Execute "update tblvousub_Book set vsubisswt=0 where vsubisswt is null and " & X
    
    
    
'================ Setup mdb Booking It Ctrl Stock ( Party Booking Vs Sale)
ProcUtlUpdateSetupBookingStock pToDt
 '===========================================
    MsgBox "Balances are updated successfully.", vbInformation + vbOKOnly, App.Title
Exit Sub
ErrorHandler:
    GProcErrorHandler
End Sub
'================ Setup mdb Booking It Ctrl Stock ( Party Booking Vs Sale)
Private Sub ProcUtlUpdateSetupBookingStock(ToDt As Date)
'On Error GoTo ErrorHandler
Dim X As String
Dim Y As String
Dim RsP As Recordset, RsTmp As Recordset
    gSetupCn.BeginTrans
    X = " Vdt<=cdate('" & ToDt & "') and Vtype='OP'"
    '--- Update to p It Ctrl No Stock
    gSetupCn.Execute "Update tblBookingSub set Bksubissbag=0,Bksubisswt=0,BkSubTmpBag = 0, BkSubTmpWt = 0 where " & X
    If (GProcRstOpen(RsP, "select * from tblBookingSub where " & X, "O", gSetupCn) > 0) Then
        RsP.MoveFirst
        Do While Not RsP.EOF '-issue bag,Wt
            '--- insert to setup db Table For All Firm
            GProcInsertToSetupDbTbl True, "select sum(VSUBBAG),sum(VSUBWt),'" & gUserName & "' from tblVousub " _
            & " where VSUBBKITCTRLNO=" & RsP.Fields("BkSUBITCTRLNO") & " having Sum(VSubBag)>0 ", "tmpGenTbl", "Amt1,amt2,Username", " username='" & gUserName & "'", ""
            gSetupCn.Execute "update tmpGenTbl set Amt1=0 where Amt2 is null and UserName='" & gUserName & "'"
            gSetupCn.Execute "update tmpGenTbl set Amt2=0 where Amt2 is null and UserName='" & gUserName & "'"
            If (GProcRstOpen(RsTmp, "select sum(Amt1),sum(Amt2) from tmpGenTbl where UserName='" & gUserName & "'", "O", gSetupCn) > 0) Then
                RsP.Fields("Bksubissbag") = IIf(IsNumeric(RsTmp.Fields(0)), RsTmp.Fields(0), 0)
                RsP.Fields("Bksubisswt") = IIf(IsNumeric(RsTmp.Fields(1)), RsTmp.Fields(1), 0)
                RsP.Update
            End If
            RsP.MoveNext
        Loop
    End If
    RsP.Close
    '-- update to 0 if null
    gSetupCn.Execute "update tblBookingSub set Bksubissbag=0 where Bksubissbag is null and " & X
    gSetupCn.Execute "update tblBookingSub set Bksubisswt=0 where Bksubisswt is null and " & X
    gSetupCn.CommitTrans
 Exit Sub
ErrorHandler:
    MsgBox "Party Booking Balance Updation is not Successed.", vbOKOnly, "Balance Updation Error"
End Sub
'===== Collect Purchase Stock Vs GatePass Datewise
Public Sub GprocCollectPurStock(pFrmDt As Date, pToDt As Date, Optional pIsInwardTypeWise As Boolean)
Dim X As String
gCn.Execute "Delete from tmpGenTbl where USERNAME='" & gUserName & "'"
'--- Inward Type Wise (Trade(PY),Sit (PT),Consig [PI,MillType = consign],Depot [PI,MillType <> consign] )
If pIsInwardTypeWise = True Then
    '-- Opening Add
    If gBackEndDB = gBackEndAccess Then
        X = " Vdt<cdate('" & pFrmDt & "')"
    Else '-- Oracle
        X = "Vdt<to_date('" & CDate(pFrmDt) & "','dd/MM/yyyy')"
    End If
    gCn.Execute "insert into tmpGenTbl ( VSRNO,VNO,Vtype,VAcCode,Nar1,UserName,Amt1,Amt2) select  ACMILLTYPE,VSubItCtrlNo,Vtype,VsubItCode,ItName,'" & gUserName & "'," _
    & "sum(vsubbag),sum(vsubwt) from tblVousub,tblMastItem,TblMastAccount " _
    & " where ITMILLCODE=ACCODE and vtype in ('PI','PY','PT','RY','RO') and tblVouSub.VsubItCode = tblMastItem.ItCode and " & X & " group by ACMILLTYPE,Vtype,VsubItCode,ItName,VSubItCtrlNo"
    '-- Inward
    If gBackEndDB = gBackEndAccess Then
        X = " Vdt>= cdate('" & pFrmDt & "')  and  Vdt<=cdate('" & pToDt & "')"
    Else '-- Oracle
        X = " Vdt>= to_date('" & CDate(pFrmDt) & "','dd/MM/yyyy')  and Vdt<=to_date('" & CDate(pToDt) & "','dd/MM/yyyy')"
    End If
    gCn.Execute "insert into tmpGenTbl (VSrNo,VNO,Vtype,VAcCode,Nar1,UserName,Amt3,Amt4) select ACMILLTYPE,VSubItCtrlNo,VType,VsubItCode,ItName,'" & gUserName & "'," _
    & "sum(vsubbag),sum(vsubwt) from tblVousub,tblMastItem,TblMastAccount " _
    & " where ITMILLCODE=ACCODE and vtype in ('PI','PY','PT','RY','RO') and tblVouSub.VsubItCode = tblMastItem.ItCode and " & X & " group by ACMILLTYPE,Vtype,VsubItCode,ItName,VSubItCtrlNo"
    '--- Outward
    If gBackEndDB = gBackEndAccess Then
        X = " tblVousub.Vdt>= cdate('" & pFrmDt & "')"
    Else '-- Oracle
        X = " tblVousub.Vdt>= to_date('" & CDate(pFrmDt) & "','dd/MM/yyyy')"
    End If
    gCn.Execute "insert into tmpGenTbl (VSrno,Vtype,VAcCode,Nar1,UserName,Amt5,Amt6) select VSrNo,tmpGenTbl.Vtype,VsubItCode,ItName,'" & gUserName & "'," _
    & "sum(vsubbag),sum(vsubwt) from tblVousub,tblMastItem,tmpGenTbl " _
    & " where tblVousub.vtype in ('GP','VY','VI') and tblVouSub.VsubItCode = tblMastItem.ItCode and VSUBPITCTRLNO=tmpGenTbl.Vno" _
    & " and Username='" & gUserName & "'  and " & X & " group by VsubItCode,ItName,vSrNo,TmpGenTbl.Vtype"
    '-- Opening Less
    If gBackEndDB = gBackEndAccess Then
        X = " tblVousub.Vdt< cdate('" & pFrmDt & "')"
    Else '-- Oracle
        X = " tblVousub.Vdt< to_date('" & CDate(pFrmDt) & "','dd/MM/yyyy')"
    End If
    gCn.Execute "insert into tmpGenTbl (VSrno,Vtype,VAcCode,Nar1,UserName,Amt1,Amt2) select VSrNo,tmpGenTbl.Vtype,VsubItCode,ItName,'" & gUserName & "'," _
    & "-sum(vsubbag),-sum(vsubwt) from tblVousub,tblMastItem,tmpGenTbl " _
    & " where tblVousub.vtype in ('GP','VY','VI') and tblVouSub.VsubItCode = tblMastItem.ItCode and VSUBPITCTRLNO= " _
    & "  tmpGenTbl.Vno and Username='" & gUserName & "' and " & X & " group by VsubItCode,ItName,vSrNo,TmpGenTbl.Vtype"
    '-- delete record having opening,inward,outward=0
    gCn.Execute "delete from tmpGenTbl  where Username='" & gUserName & "' and VAcCode in (select VAcCode from tmpGenTbl where username='" & gUserName & "' group by VAcCode having (sum(amt1+amt3-amt5))=0 and (sum(amt2+amt4-amt6))=0)"
Else
''    '-- Opening Add
''    If gBackEndDB = gBackEndAccess Then
''        X = " Vdt<cdate('" & pFrmDt & "')"
''    Else '-- Oracle
''        X = "Vdt<to_date('" & CDate(pFrmDt) & "','dd/MM/yyyy')"
''    End If
''    gCn.Execute "insert into tmpGenTbl (VNO,VAcCode,Nar1,UserName,Amt1,Amt2) select  VSubItCtrlNo,VsubItCode,ItName,'" & gUserName & "'," _
''    & "sum(vsubbag),sum(vsubwt) from tblVousub,tblMastItem " _
''    & " where vtype in ('PI','PY','PT','RY','RO') and tblVouSub.VsubItCode = tblMastItem.ItCode and " & X & " group by VsubItCode,ItName,VSubItCtrlNo"
''    '-- Inward
''    If gBackEndDB = gBackEndAccess Then
''        X = " Vdt>= cdate('" & pFrmDt & "')  and  Vdt<=cdate('" & pToDt & "')"
''    Else '-- Oracle
''        X = " Vdt>= to_date('" & CDate(pFrmDt) & "','dd/MM/yyyy')  and Vdt<=to_date('" & CDate(pToDt) & "','dd/MM/yyyy')"
''    End If
''    gCn.Execute "insert into tmpGenTbl (VNO,VAcCode,Nar1,UserName,Amt3,Amt4) select VSubItCtrlNo,VsubItCode,ItName,'" & gUserName & "'," _
''    & "sum(vsubbag),sum(vsubwt) from tblVousub,tblMastItem " _
''    & " where vtype in ('PI','PY','PT','RY','RO') and tblVouSub.VsubItCode = tblMastItem.ItCode and " & X & " group by VsubItCode,ItName,VSubItCtrlNo"
''    '--- Outward
''    If gBackEndDB = gBackEndAccess Then
''        X = " Vdt>= cdate('" & pFrmDt & "')  and  Vdt<=cdate('" & pToDt & "')"
''    Else '-- Oracle
''        X = " Vdt>= to_date('" & CDate(pFrmDt) & "','dd/MM/yyyy')  and Vdt<=to_date('" & CDate(pToDt) & "','dd/MM/yyyy')"
''    End If
''    gCn.Execute "insert into tmpGenTbl (VAcCode,Nar1,UserName,Amt5,Amt6) select VsubItCode,ItName,'" & gUserName & "'," _
''    & "sum(vsubbag),sum(vsubwt) from tblVousub,tblMastItem " _
''    & " where vtype in ('GP','VY','VI') and tblVouSub.VsubItCode = tblMastItem.ItCode and VSUBPITCTRLNO in" _
''    & " ( select Vno from tmpGenTbl where Username='" & gUserName & "')  and " & X & " group by VsubItCode,ItName"
''    '-- Opening Less
''    If gBackEndDB = gBackEndAccess Then
''        X = " Vdt< cdate('" & pFrmDt & "')"
''    Else '-- Oracle
''        X = " Vdt< to_date('" & CDate(pFrmDt) & "','dd/MM/yyyy')"
''    End If
''    gCn.Execute "insert into tmpGenTbl (VAcCode,Nar1,UserName,Amt1,Amt2) select VsubItCode,ItName,'" & gUserName & "'," _
''    & "-sum(vsubbag),-sum(vsubwt) from tblVousub,tblMastItem " _
''    & " where vtype in ('GP','VY','VI') and tblVouSub.VsubItCode = tblMastItem.ItCode and VSUBPITCTRLNO in" _
''    & " ( select Vno from tmpGenTbl where Username='" & gUserName & "')  and " & X & " group by VsubItCode,ItName"
''    '-- delete record having opening,inward,outward=0
''    gCn.Execute "delete from tmpGenTbl  where Username='" & gUserName & "' and nar1 in (select nar1 from tmpGenTbl where username='" & gUserName & "' group by nar1 having (sum(amt1+amt3-amt5))=0 and (sum(amt2+amt4-amt6))=0)"

    '-- Opening Add
    If gBackEndDB = gBackEndAccess Then
        X = " Vdt<cdate('" & pFrmDt & "')"
    Else '-- Oracle
        X = "Vdt<to_date('" & CDate(pFrmDt) & "','dd/MM/yyyy')"
    End If
    gCn.Execute "insert into tmpGenTbl (VNO,VAcCode,Nar1,nar2,UserName,Amt1,Amt2) select  VSubItCtrlNo,VsubItCode,ItName,AcName,'" & gUserName & "'," _
    & "sum(vsubbag),sum(vsubwt) from tblVousub,tblMastItem,tblMastAccount" _
    & " where ITMILLCODE=ACCODE and vtype in ('PI','PY','PT','RY','RO') and tblVouSub.VsubItCode = tblMastItem.ItCode and " & X & " group by VsubItCode,ItName,Acname,VSubItCtrlNo"
    '-- Inward
    If gBackEndDB = gBackEndAccess Then
        X = " Vdt>= cdate('" & pFrmDt & "')  and  Vdt<=cdate('" & pToDt & "')"
    Else '-- Oracle
        X = " Vdt>= to_date('" & CDate(pFrmDt) & "','dd/MM/yyyy')  and Vdt<=to_date('" & CDate(pToDt) & "','dd/MM/yyyy')"
    End If
    gCn.Execute "insert into tmpGenTbl (VNO,VAcCode,Nar1,Nar2,UserName,Amt3,Amt4) select VSubItCtrlNo,VsubItCode,ItName,acname,'" & gUserName & "'," _
    & "sum(vsubbag),sum(vsubwt) from tblVousub,tblMastItem,tblMastAccount " _
    & " where ITMILLCODE=ACCODE and vtype in ('PI','PY','PT','RY','RO') and tblVouSub.VsubItCode = tblMastItem.ItCode and " & X & " group by VsubItCode,ItName,acname,VSubItCtrlNo"
    '--- Outward
    If gBackEndDB = gBackEndAccess Then
        X = " Vdt>= cdate('" & pFrmDt & "')  and  Vdt<=cdate('" & pToDt & "')"
    Else '-- Oracle
        X = " Vdt>= to_date('" & CDate(pFrmDt) & "','dd/MM/yyyy')  and Vdt<=to_date('" & CDate(pToDt) & "','dd/MM/yyyy')"
    End If
    gCn.Execute "insert into tmpGenTbl (VAcCode,Nar1,Nar2,UserName,Amt5,Amt6) select VsubItCode,ItName,AcName,'" & gUserName & "'," _
    & "sum(vsubbag),sum(vsubwt) from tblVousub,tblMastItem,tblMastAccount " _
    & " where ITMILLCODE=ACCODE and vtype in ('GP','VY','VI') and tblVouSub.VsubItCode = tblMastItem.ItCode and VSUBPITCTRLNO in" _
    & " ( select Vno from tmpGenTbl where Username='" & gUserName & "')  and " & X & " group by VsubItCode,ItName,Acname"
    '-- Opening Less
    If gBackEndDB = gBackEndAccess Then
        X = " Vdt< cdate('" & pFrmDt & "')"
    Else '-- Oracle
        X = " Vdt< to_date('" & CDate(pFrmDt) & "','dd/MM/yyyy')"
    End If
    gCn.Execute "insert into tmpGenTbl (VAcCode,Nar1,Nar2,UserName,Amt1,Amt2) select VsubItCode,ItName,AcName,'" & gUserName & "'," _
    & "-sum(vsubbag),-sum(vsubwt) from tblVousub,tblMastItem,tblMastAccount " _
    & " where ITMILLCODE=ACCODE and vtype in ('GP','VY','VI') and tblVouSub.VsubItCode = tblMastItem.ItCode and VSUBPITCTRLNO in" _
    & " ( select Vno from tmpGenTbl where Username='" & gUserName & "')  and " & X & " group by VsubItCode,ItName,Acname"
    '-- delete record having opening,inward,outward=0
'    gCn.Execute "delete from tmpGenTbl  where Username='" & gUserName & "' and nar1 in (select nar1 from tmpGenTbl where username='" & gUserName & "' group by nar1 having (sum(amt1+amt3-amt5))=0 and (sum(amt2+amt4-amt6))=0)"


End If

End Sub
'==== Mill Count Stock
Public Sub GprocPrepareCountStock(pFrmDt As Date, pToDt As Date, MillName As String)
Dim X As String
Dim Y As String
Dim MillCode As Long
MillCode = GProcGetColumnValue("tblMastAccount", "AcName", MillName, "S", "AcCode", "N")
gCn.BeginTrans
gCn.Execute "Delete from tmpGenTbl where USERNAME='" & gUserName & "'"
'-- Opening Add
If gBackEndDB = gBackEndAccess Then
    X = " Vdt<cdate('" & pFrmDt & "')"
Else '-- Oracle
    X = "Vdt<to_date('" & CDate(pFrmDt) & "','dd/MM/yyyy')"
End If
gCn.Execute "insert into tmpGenTbl (VAcCode,VACOCODE,Nar1,UserName,Amt1,Amt2) select ITMILLCODE,VsubItCode,ItName,'" & gUserName & "'," _
& "sum(vsubbag),sum(vsubwt) from tblVousub,tblMastItem " _
& " where vtype in ('PI') and ITMILLCODE=" & MillCode & " and tblVouSub.VsubItCode = tblMastItem.ItCode and " & X & " group by ITMILLCODE,VsubItCode,ItName"
''PY','PT'
'-- Opening Less
If gBackEndDB = gBackEndAccess Then
    X = " Vdt< cdate('" & pFrmDt & "')"
Else '-- Oracle
    X = " Vdt< to_date('" & CDate(pFrmDt) & "','dd/MM/yyyy')"
End If
gCn.Execute "insert into tmpGenTbl (VAcCode,VACOCODE,Nar1,UserName,Amt1,Amt2) select ITMILLCODE,VsubItCode,ItName,'" & gUserName & "'," _
& "-sum(vsubbag),-sum(vsubwt) from tblVousub,tblMastItem " _
& " where vtype in ('SD','SO') and tblVouSub.VsubItCode = tblMastItem.ItCode " _
& " and " & X & " and ITMILLCODE=" & MillCode & " group by ITMILLCODE,VsubItCode,ItName"
'-- Inward
If gBackEndDB = gBackEndAccess Then
    X = " Vdt>= cdate('" & pFrmDt & "')  and  Vdt<=cdate('" & pToDt & "')"
Else '-- Oracle
    X = " Vdt>= to_date('" & CDate(pFrmDt) & "','dd/MM/yyyy')  and Vdt<=to_date('" & CDate(pToDt) & "','dd/MM/yyyy')"
End If
gCn.Execute "insert into tmpGenTbl (VAcCode,VACOCODE,Nar1,UserName,Amt3,Amt4) select ITMILLCODE,VsubItCode,ItName,'" & gUserName & "'," _
& "sum(vsubbag),sum(vsubwt) from tblVousub,tblMastItem " _
& " where vtype in ('PI') and ITMILLCODE=" & MillCode & " and tblVouSub.VsubItCode = tblMastItem.ItCode and " & X & " group by ITMILLCODE,VsubItCode,ItName"
'--- Outward
If gBackEndDB = gBackEndAccess Then
    X = " Vdt>= cdate('" & pFrmDt & "')  and  Vdt<=cdate('" & pToDt & "')"
Else '-- Oracle
    X = " Vdt>= to_date('" & CDate(pFrmDt) & "','dd/MM/yyyy')  and Vdt<=to_date('" & CDate(pToDt) & "','dd/MM/yyyy')"
End If
gCn.Execute "insert into tmpGenTbl (VAcCode,VACOCODE,Nar1,UserName,Amt5,Amt6) select ITMILLCODE,VsubItCode,ItName,'" & gUserName & "'," _
& "sum(vsubbag),sum(vsubwt) from tblVousub,tblMastItem " _
& " where vtype in ('SO','SD') and ITMILLCODE=" & MillCode & " and tblVouSub.VsubItCode = tblMastItem.ItCode " _
& "  and " & X & " group by ITMILLCODE,VsubItCode,ItName"
gCn.CommitTrans
End Sub
'==== Mill Count Stock (Spl)
Public Sub GprocPrepareCountStockSpl(pFrmDt As Date, pToDt As Date)
Dim X As String
Dim Y As String
Dim MillCode As Long
'MillCode = GProcGetColumnValue("tblMastAccount", "AcName", MillName, "S", "AcCode", "N")
gCn.BeginTrans
gCn.Execute "Delete from tmpGenTbl where USERNAME='" & gUserName & "'"
'-- Opening Add
X = " Vdt<cdate('" & pFrmDt & "')"
gCn.Execute "insert into tmpGenTbl (VAcCode,VACOCODE,Nar1,UserName,Amt1,Amt2) select ITMILLCODE,VsubItCode,ItName,'" & gUserName & "'," _
& "sum(vsubbag),sum(vsubwt) from tblVousub,tblMastItem,tblMastAccount" _
& " where vtype in ('PI') and tblVouSub.VsubItCode = tblMastItem.ItCode and tblMastItem.ITMILLCODE = tblMastAccount.Accode and tblMastAccount.AcMillType = 2 and  " & X & " group by ITMILLCODE,VsubItCode,ItName"
''PY','PT'
'-- Opening Less
X = " Vdt< cdate('" & pFrmDt & "')"
gCn.Execute "insert into tmpGenTbl (VAcCode,VACOCODE,Nar1,UserName,Amt1,Amt2) select ITMILLCODE,VsubItCode,ItName,'" & gUserName & "'," _
& "-sum(vsubbag),-sum(vsubwt) from tblVousub,tblMastItem,tblMastAccount" _
& " where vtype in ('SD','SO') and tblVouSub.VsubItCode = tblMastItem.ItCode " _
& " and " & X & " and tblVouSub.VsubItCode = tblMastItem.ItCode And tblMastItem.ITMILLCODE = tblMastAccount.AcCode And tblMastAccount.AcMillType = 2 group by ITMILLCODE,VsubItCode,ItName"

'& " and " & X & " and ITMILLCODE=" & MillCode & " group by ITMILLCODE,VsubItCode,ItName"

'-- Inward
X = " Vdt>= cdate('" & pFrmDt & "')  and  Vdt<=cdate('" & pToDt & "')"
gCn.Execute "insert into tmpGenTbl (VAcCode,VACOCODE,Nar1,UserName,Amt3,Amt4) select ITMILLCODE,VsubItCode,ItName,'" & gUserName & "'," _
& "sum(vsubbag),sum(vsubwt) from tblVousub,tblMastItem,tblMastAccount " _
& " where vtype in ('PI') and  tblVouSub.VsubItCode = tblMastItem.ItCode and tblMastItem.ITMILLCODE = tblMastAccount.AcCode And tblMastAccount.AcMillType = 2  and " & X & " group by ITMILLCODE,VsubItCode,ItName"
'--- Outward
X = " Vdt>= cdate('" & pFrmDt & "')  and  Vdt<=cdate('" & pToDt & "')"
gCn.Execute "insert into tmpGenTbl (VAcCode,VACOCODE,Nar1,UserName,Amt5,Amt6) select ITMILLCODE,VsubItCode,ItName,'" & gUserName & "'," _
& "sum(vsubbag),sum(vsubwt) from tblVousub,tblMastItem,tblMastAccount " _
& " where vtype in ('SO','SD') and tblVouSub.VsubItCode = tblMastItem.ItCode And tblMastItem.ITMILLCODE = tblMastAccount.AcCode And tblMastAccount.AcMillType = 2 and tblVouSub.VsubItCode = tblMastItem.ItCode " _
& "  and " & X & " group by ITMILLCODE,VsubItCode,ItName"
gCn.CommitTrans
End Sub

'==== Item Ledger
Public Sub GprocPrepareStkItemLedger(pFrmDt As Date, pToDt As Date, Item As String, Mill As String)
Dim CondStr As String
Dim DtStr As String
Dim OpStkStr As String
gCn.BeginTrans
gCn.Execute "Delete from tmpGenTbl where USERNAME='" & gUserName & "'"
'==== Opening
    '-- Date
    If gBackEndDB = gBackEndAccess Then
        DtStr = " BILLDT<cdate('" & pFrmDt & "')"
    Else '-- Oracle
        DtStr = "BILLDT<to_date('" & CDate(pFrmDt) & "','dd/MM/yyyy')"
    End If
    '-- Item
    If Item <> "" Then CondStr = " and Item='" & Item & "'"
    '-- Mill
    If Mill <> "" Then CondStr = CondStr + " and MILL='" & Mill & "'"
    '-- Opening Stock str (In - Out)
    If gBackEndDB = gBackEndAccess Then
        OpStkStr = " iif(isnull(Sum(InBag)),0,Sum(InBag)) - iif(isnull(Sum(OutBag)),0,Sum(OutBag)), iif(isnull(Sum(InWt)),0,Sum(InWt)) - iif(isnull(Sum(OutWt)),0,Sum(OutWt)) "
    Else
        OpStkStr = " nvl(Sum (InBag),0) - nvl(Sum(OutBag),0),nvl(Sum(InWt),0) - nvl(Sum(OutWt),0)"
    End If
    gCn.Execute "insert into tmpGenTbl (Vtype,VBillDt,VSrNo,Nar1,Nar2,VAcCode,Amt1,Amt2,UserName) select 'OP'," _
    & IIf(gBackEndDB = gBackEndAccess, "cdate('" & pFrmDt & "')", "to_date('" & pFrmDt & "','dd/MM/yyyy')") & ",-1,ItemMill," _
    & "'Opening Balance',ItCode," & OpStkStr & ",'" & gUserName & "'" _
    & " from QryStkItemLedger where " & DtStr & CondStr & " group by ItemMill,ItCode having (sum(inbag)-sum(Outbag)<>0) and (sum(inWt)-sum(OutWt)<>0)  "
'===In / Out
    If gBackEndDB = gBackEndAccess Then
        DtStr = " BILLDT>= cdate('" & pFrmDt & "')  and  BILLDT<=cdate('" & pToDt & "')"
    Else '-- Oracle
        DtStr = " BILLDT>= to_date('" & CDate(pFrmDt) & "','dd/MM/yyyy')  and BILLDT<=to_date('" & CDate(pToDt) & "','dd/MM/yyyy')"
    End If
    gCn.Execute "insert into tmpGenTbl (Vtype,VBillNo,VBillDt,VSrNo,Nar1,Nar2,VAcCode,Amt1,Amt2,Amt3,amt4,UserName) " _
    & " select Vtype,BillNo,billDt,Orderno,ItemMill,particulars,Itcode,InBag,inWt,OutBag,OutWt,'" & gUserName & "'" _
    & " from QryStkItemLedger where " & DtStr & CondStr
gCn.CommitTrans
End Sub
'==== Daily Stock To Mill
Public Sub GprocPrepareDailyStkToMill(pFromDt As Date, pToDt As Date, MillAcCode As Long)
Dim CondStr As String
Dim DtStr As String
Dim RsTmp As Recordset
Dim X As String
Dim OutBag As Long
gCn.BeginTrans
gCn.Execute "Delete from tmpGenTbl where USERNAME='" & gUserName & "'"
'==== Opening
    '-- Date
    If gBackEndDB = gBackEndAccess Then
        DtStr = " tblvouSub.Vdt<cdate('" & pToDt & "')"
    Else '-- Oracle
        DtStr = " tblvouSub.Vdt<to_date('" & CDate(pToDt) & "','dd/MM/yyyy')"
    End If
    '-- Mill
    If MillAcCode > 0 Then CondStr = " where ITMILLCODE=" & MillAcCode
    '-- Opening Stock  In (PI) -  Out (SD,SO)
    '--In
    gCn.Execute "insert into tmpGenTbl (VSRNO,vAcCode,Amt1,UserName) select 1,Itcode,(" _
    & " select " & IIf(gBackEndDB = gBackEndAccess, "iif(isnull(Sum(VSubBag)),0,Sum(VSubBag))", "nvl(Sum(VSubBag), 0)") & " from tblvouSub  where VSUBITCODE=tblmastItem.ItCode and Vtype in ('PI','RO') " _
    & " and " & DtStr & " ) ,'" & gUserName & "' from tblmastItem " & CondStr
    '-- Out
    If gBackEndDB = gBackEndAccess Then
        '- Update Oening Balance
        X = " Select * from tmpGenTbl where userName='" & gUserName & "'"
        If GProcRstOpen(RsTmp, X, "O") > 0 Then
            With RsTmp
                .MoveFirst
                OutBag = 0
                Do While Not .EOF
                    OutBag = GProcGetColumnValue("tmpGenTbl,tblvouSub ", "UserName", gUserName, "S", "Sum(VSubBag)", "N", "tmpGenTbl.vAcCode=VSUBITCODE  and " _
                    & " tblvouSub.vtype in ('SD','SO','VI') and vAcCode=" & .Fields("VAcCode") & " and " & DtStr & " and tmpGenTbl.username='" & gUserName & "'")
                    .Fields("amt1") = .Fields("amt1") - OutBag
                    .Update
                    .MoveNext
                Loop
            End With
        End If
    Else
        gCn.Execute "update tmpGenTbl set amt1=amt1- (select  nvl(Sum(VSubBag), 0)  from " _
        & " tblVousub where  tmpGenTbl.vAcCode=VSUBITCODE and  tblvouSub.vtype in ('SD','SO','VI') " _
        & "  and " & DtStr & " )  where username='" & gUserName & "'"
    End If
    '=== Receipts
    '-- Date Condition
    If gBackEndDB = gBackEndAccess Then
        CondStr = " tblvouSub.Vdt=cdate('" & pToDt & "')"
    Else '-- Oracle
        CondStr = " tblvouSub.Vdt=to_date('" & CDate(pToDt) & "','dd/MM/yyyy')"
    End If
    '-- Mill Condition
    If MillAcCode > 0 Then CondStr = CondStr + " and ITMILLCODE=" & MillAcCode
    '=== Inward
    gCn.Execute "insert into tmpGenTbl (VSRNO,vAcCode,Nar1,Dt1,Amt2,UserName) select 2,VSUBITCODE, VSUBLRNO,VSUBLRDATE,VSubBag,'" & gUserName & "' from " _
    & " tblvouSub,tblmastItem where VSUBITCODE=tblmastItem.ItCode and tblvouSub.Vtype in ('PI','RO')  and " & CondStr
    '===Sales
    gCn.Execute "insert into tmpGenTbl (VSRNO,vAcCode,VBillno,Amt3,UserName) select 3,VSUBITCODE," _
    & "AdInvNo,tblvouSub.VsubBag,'" & gUserName & "' from tblvouSub ,tblAddLess,tblmastItem where " _
    & " tblAddLess.vType=tblvouSub.Vtype and tblAddLess.vNo=tblvouSub.VNo and tblAddLess.vYear=tblvouSub.VYear and " _
    & " tblvouSub.VSUBITCODE=tblmastItem.ItCode and tblvouSub.Vtype in ('SD','SO','VI') and " & CondStr
    gCn.Execute "delete from tmpGenTbl where amt1=0 and amt2=0 and amt3=0 and username='" & gUserName & "'"
    gCn.CommitTrans
End Sub
'----SRK Transfer Balance
'======== Utility Carry Forword Procedure ==============
Public Sub GProcUtlCarryForwardOpening(ToDbNm As String)
'On Error GoTo ErrorRoutine
Dim i As Long
Dim TblMstCount As Integer
Dim gCnToYr As Connection '- Connection For Next Yr
Dim Vno As Long
Dim vamt As Double
Dim VYear As Long
Dim AcOCode As Long
Dim rsBalS As Recordset '- Balance Source Recordset
Dim rsBalS1 As Recordset
Dim rsBalT As Recordset '- Balance Target Recordset
Dim rsBalT1 As Recordset '- Balance Target Recordset
Dim X As String
Dim DtStr As String
GProcConnectionOpen gCnToYr, ToDbNm
'=================== Master (No of tables = 4) New Transaction Entry ================================
For TblMstCount = 0 To 4
    Select Case TblMstCount
        Case 0 '- Comp Group
            X = " Select  *  from tblMastCompGroup "
        Case 1 '-A/c Group
            X = " Select  *  from tblMastGroup "
        Case 2 '- A/c
            X = " Select  *  from tblMastAccount "
        Case 3 '- Item
            X = " Select  *  from tblMastItem "
        Case 4 '- Narration
            X = " Select  *  from tblMastNarration "
    End Select
    If GProcRstOpen(rsBalS, X, "R") > 0 Then
        With rsBalS
            Do While Not .EOF
                X = ""
                Select Case TblMstCount
                    Case 0 '- Comp Group
                        X = " Select  *  from tblMastCompGroup where CGCode=" & .Fields("CGCode")
                    Case 1 '-A/c Group
                        X = " Select  *  from tblMastGroup where AgCode=" & .Fields("AgCode")
                    Case 2 '- A/c
                        X = " Select  *  from tblMastAccount where AcCode=" & .Fields("AcCode")
                        
                    Case 3 '- Item
                        X = " Select  *  from tblMastItem where ItCode=" & .Fields("ItCode")
                    Case 4 '- Narration
                        X = " Select  *  from tblMastNarration where NarrCode=" & .Fields("NarrCode")
                End Select
                
                       If GProcRstOpen(rsBalT, X, "O", gCnToYr) = 0 Then '--- Add Master Entry To Next Year
                              rsBalT.AddNew
                               For i = 0 To .Fields.Count - 1
                                  If rsBalT!AcName = !AcName Then '---- check If Account Name Already Present Add "Old"
                                     rsBalT.Fields(1) = "Old " + .Fields(1)
                                  End If
                                     rsBalT.Fields(i) = .Fields(i)
                               Next i
                              rsBalT.Update
                        End If
                .MoveNext
            Loop
        End With
    End If
Next TblMstCount
'=============== Next Year  Delete All Opening Bal
gCnToYr.Execute "delete from tblVoucher where Vtype='0P'"
AcOCode = GProcGetColumnValue("TblMastAccount", "AcName", "-", "S", "AcCode", "N") '-- "-" A/c
VYear = CLng(gCYear) + 1
'===================  Next Year Closing Stock Balances =====================
'- Carry forward 31/03 cl. stock to Next Yr
If gBackEndDB = gBackEndAccess Then
    DtStr = "StDate=cdate('" & gCYEDate & "')"
Else
    DtStr = "StDate=to_date('" & gCYEDate & "','dd/MM/yyyy')"
End If
X = "select * from tblClosingStock where " & DtStr
Vno = 0
If GProcRstOpen(rsBalS, X, "R") > 0 Then
    With rsBalS
        Vno = Vno + 1
        If gBackEndDB = gBackEndAccess Then
            DtStr = "cdate('" & gCYEDate + 1 & "')"
        Else
            DtStr = "to_date('" & gCYEDate + 1 & "','dd/MM/yyyy')"
        End If
        gCnToYr.Execute "insert into tblVoucher (VNo,VCtrNo,VDt,VType,VAcCode,VAcOCode,VAmt,vUserName,vYear)  Values " _
        & "(" & Vno & ",1," & DtStr & ",'0P'," & .Fields("StAcCode") & "," & AcOCode & ",-" & .Fields("StAmount") & ",'" & gUserName & "'," & VYear & ")"
    End With
End If
'===================  Next Year Account Opening Balances =====================
gCn.Execute "delete from tmpselection where UserName='" & gUserName & "'"
If gBackEndDB = gBackEndAccess Then
    gCn.Execute "insert into tmpselection (dt1,Dt2,UserName) values ('" & gCYSDate & "','" & gCYEDate & "','" & gUserName & "')"
Else
    gCn.Execute "insert into tmpselection (dt1,dt2,UserName) values (to_date('" & gCYSDate & "','dd/MM/yyyy'),to_Date('" & gCYEDate & "','dd/MM/yyyy'),'" & gUserName & "') "
End If
If gBackEndDB = gBackEndAccess Then
    DtStr = "cdate('" & gCYEDate + 1 & "')"
Else
    DtStr = "to_date('" & gCYEDate + 1 & "','dd/MM/yyyy')"
End If
X = "select * from qryAccountBalance where TopGroup in (90001,90002)  and  UserName='" & gUserName & "'"
If GProcRstOpen(rsBalS, X, "R") > 0 Then
    With rsBalS
        Do While Not .EOF
            Vno = Vno + 1
            If .Fields("Debit") > 0 Then
                vamt = CDbl(.Fields("Debit") * -1)
            Else
                vamt = CDbl(.Fields("Credit"))
            End If
            gCnToYr.Execute "insert into tblVoucher (VNo,VCtrNo,VDt,VType,VAcCode,VAcOCode,VAmt,vUserName,vYear)  Values " _
            & "(" & Vno & ",1," & DtStr & ",'0P'," & .Fields("AcCode") & "," & AcOCode & "," & vamt & ",'" & gUserName & "'," & VYear & ")"
            .MoveNext
        Loop
    End With
End If
MsgBox "Utility Carry Forward Opening Balances A/c is Completed Successfully.", vbOKOnly + vbInformation, "Utility"
Exit Sub
ErrorRoutine:
    GProcErrorHandler
End Sub

Public Sub GProcUtlCarryForwardOutstanding(ToDbNm As String)
'On Error GoTo ErrorRoutine
Dim i As Long
Dim TblMstCount As Integer
Dim gCnToYr As Connection '- Connection For Next Yr
Dim Vno As Long
Dim vamt As Double
Dim VYear As Long
Dim AcOCode As Long
Dim rsBalS As Recordset '- Balance Source Recordset
Dim rsBalS1 As Recordset
Dim rsBalT As Recordset '- Balance Target Recordset
Dim rsBalT1 As Recordset '- Balance Target Recordset
Dim X As String
Dim DtStr As String
Dim XX As String
Dim RstLstRecDt As Recordset
GProcConnectionOpen gCnToYr, ToDbNm
X = " Select  *  from tblOutStanding where ((OutAmount-OutRecAmt)>0  or " _
& "  (OutDbNtAmt-OutRecDbAmt)>0)"
If GProcRstOpen(rsBalS, X, "R") > 0 Then
    With rsBalS
        Do While Not .EOF
            X = "Select * from tblOutStanding Where VYear=" & .Fields("VYear") & " and Vtype='" & .Fields("VType") & "' and VNo=" & .Fields("VNo")
            If GProcRstOpen(rsBalT, X, "O", gCnToYr) = 0 Then
                rsBalT.AddNew
                For i = 0 To .Fields.Count - 1 '- Add TblOutstanding Entry
                    rsBalT.Fields(i) = .Fields(i)
                Next i
            End If '--- Update TblOutstanding Balance
            rsBalT!OutAmount = .Fields("OutAmount") - .Fields("OutRecAmt") '- Out Amt
            rsBalT!OutDbNtAmt = .Fields("OutDbNtAmt") - .Fields("OutRecDbAmt") '-DbNt Amt
            rsBalT!OutRecAmt = 0 '-Rec. Amt
            rsBalT!OutRecDbAmt = 0 '-Rec. DbNt Amt
            rsBalT!OutRecFreight = 0 '-Rec. Freight Amt
            rsBalT!OutAcCode = .Fields("OutAcCode")
            rsBalT!OutBrokerCode = .Fields("OutBrokerCode")
'            rsBalT!OutCrt = .Fields("OutCrt")
'            rsBalT!OutSrt = .Fields("OutSrt")
'            rsBalT!OutIrt = .Fields("OutIrt")
            XX = "select AdGpDt from tblAddLess where vno =  " & .Fields(0) & "  and vtype = '" & .Fields(1) & "'  and vyear = " & .Fields(2) & ""
            If (GProcRstOpen(RstLstRecDt, XX, "R", gCn) > 0) Then
               rsBalT!OutPurBillDt = RstLstRecDt.Fields(0)
            End If
            rsBalT.Update
            .MoveNext
        Loop
    End With
End If
'-- Delete Outstanding in Next Year for (Complete & No Vs Receipt in Next Year)
X = " Select  *  from tblOutStanding where ((OutAmount-OutRecAmt)<=0  and " _
& "  (OutDbNtAmt-OutRecDbAmt)<=0)"
If GProcRstOpen(rsBalS, X, "R") > 0 Then
    With rsBalS
        Do While Not .EOF
            gCnToYr.Execute "delete from tblOutStanding where Vno not in ( select vno from tblRecVsSale where " _
            & " InvNo=" & .Fields("VNo") & " and  InvTp='" & .Fields("VType") & "' and  InvYear=" & .Fields("VYear") & ") " _
            & " and VYear=" & .Fields("VYear") & " and Vtype='" & .Fields("VType") & "' and VNo=" & .Fields("VNo")
            rsBalS.MoveNext
        Loop
    End With
End If
'--- Delete Purchase Outstanding
'gCnToYr.Execute "delete from tblOutStanding where vtype in ('PY','PT','PO','PI') and vyear=" & gCYear

MsgBox "Utility Carry Forward OutStanding is Completed Successfully.", vbOKOnly + vbInformation, "Utility"
Exit Sub
End Sub

Public Sub GProcUtlCarryForwardStock(ToDbNm As String)
'''On Error GoTo ErrorRoutine
''Dim i As Long
''Dim TblMstCount As Integer
''Dim gCnToYr As Connection '- Connection For Next Yr
''Dim Vno As Long
''Dim vamt As Double
''Dim VYear As Long
''Dim AcOCode As Long
''Dim rsBalS As Recordset '- Balance Source Recordset
''Dim rsBalS1 As Recordset
''Dim rsBalT As Recordset '- Balance Target Recordset
''Dim rsBalT1 As Recordset '- Balance Target Recordset
''Dim X As String
''Dim DtStr As String
''GProcConnectionOpen gCnToYr, ToDbNm
''X = " Select  *  from tblVouSub where Vtype in ('PY','PT','PI')  and " _
''& "  VSubItCtrlNo>0 and (VSubBag-VSubIssBag)>0  order by vtype,vno"
''If GProcRstOpen(rsBalS, X, "R") > 0 Then
''    With rsBalS
''        Do While Not .EOF
''            X = "Select * from Tblvousub Where VYear=" & .Fields("VYear") & " and Vtype='" & .Fields("VType") & "' and VSubItCtrlNo=-" & .Fields("VSubItCtrlNo")
''            If GProcRstOpen(rsBalT, X, "O", gCnToYr) > 0 Then '--- Update TblVouSub Balance
''                rsBalT!vsubBAG = .Fields("VSubBag") - .Fields("VSubIssBag") '- Bag
''                rsBalT!vsubwt = .Fields("VSubWt") - .Fields("VSubIssWt") '-Wt
''                rsBalT.Update
''            Else
''                X = "Select * from Tbladdless Where VYear=" & .Fields("VYear") & " and VType='" & .Fields("VType") & "' and VNo=" & .Fields("VNo")
''                GProcRstOpen rsBalS1, X, "R"
''                X = "Select * from Tbladdless Where VYear=" & .Fields("VYear") & " and VType='" & .Fields("VType") & "' and VNo=" & .Fields("VNo")
''
''                If GProcRstOpen(rsBalT1, X, "O", gCnToYr) = 0 Then
''                    rsBalT1.AddNew
''                    For i = 0 To rsBalS1.Fields.Count - 1 '- Add TblAddLess Entry
''                        rsBalT1.Fields(i) = rsBalS1.Fields(i)
''                    Next i
''                    rsBalT1.Update
''                End If
''                rsBalT.AddNew
''                For i = 0 To .Fields.Count - 1 '- Add TblVouSub Entry
''                    rsBalT.Fields(i) = .Fields(i)
''                Next i
''                'rsBalT.Update
''                rsBalT!vsubAMT = (.Fields("VSubAmt") / .Fields("VSubBag")) * (.Fields("VSubBag") - .Fields("VSubIssBag"))
''                DoEvents
''                rsBalT!vsubBAG = .Fields("VSubBag") - .Fields("VSubIssBag") '- Bag
''                rsBalT!vsubwt = .Fields("VSubWt") - .Fields("VSubIssWt") '-Wt
''                rsBalT!VSubItCtrlNo = -1 * .Fields("VSubItCtrlNo") '- It Ctrl No
''                rsBalT!VSubPItCtrlNo = -1 * .Fields("VSubPItCtrlNo") '-P It Ctrl No
''                rsBalT.Update
''            End If
''            .MoveNext
''        Loop
''    End With
''End If
''
'''----For Closing Stock Value Transfer
'''---For TblVouSub - VsubAmt
''
''GProcConnectionOpen gCnToYr, ToDbNm
''X = " Select  *  from tblVouSub where Vtype in ('PY','PT')  and " _
''& "  VSubItCtrlNo>0 and (VSubBag-VSubIssBag)>0  order by vtype,vno"
''If GProcRstOpen(rsBalS, X, "R") > 0 Then
''    With rsBalS
''        Do While Not .EOF
''            X = "Select * from Tblvousub Where VYear=" & .Fields("VYear") & " and Vtype='" & .Fields("VType") & "' and VSubItCtrlNo=-" & .Fields("VSubItCtrlNo")
''            If GProcRstOpen(rsBalT, X, "O", gCnToYr) > 0 Then '--- Update TblVouSub Balance
''               rsBalT!vsubAMT = (.Fields("VSubAmt") / .Fields("VSubBag")) * (.Fields("VSubBag") - .Fields("VSubIssBag"))
''               DoEvents
''               rsBalT.Update
''            End If
''            .MoveNext
''        Loop
''    End With
''End If
'''---For TblAddless - AdSubAmt,AdOtherLess etc
''
''Dim rsLastYearAddless As Recordset
''Dim rsNextYearAddless As Recordset
''Dim rsNextYearVouSub As Recordset
''Dim Y As String
''Dim TAdd As Double
''Dim TLess As Double
''Dim TAddAftTax As Double
''Dim TLessAftTax As Double
''Dim TTax As Double
''Dim TBillAmt As Double
''
''
''GProcConnectionOpen gCnToYr, ToDbNm
''X = "select * from tblAddless where vno in(select vno from tblVouSub where  Vtype in ('PY')  and  VSubItCtrlNo>0 and (VSubBag-VSubIssBag)>0    ) and vtype in('PY') and Vyear <= " & gCYear & " order by vtype,vno  "
''If GProcRstOpen(rsLastYearAddless, X, "R") > 0 Then
''    With rsLastYearAddless
''        Do While Not .EOF
''            X = "Select * from TblAddless Where VYear=" & .Fields("VYear") & " and Vtype='" & .Fields("VType") & "' and Vno=" & .Fields("Vno")
''            If GProcRstOpen(rsNextYearAddless, X, "O", gCnToYr) > 0 Then '--- Update TblVouSub Balance
''                Y = "Select sum(VsubAmt) as NextYeaeSubAmt from tblVouSub Where VYear=" & .Fields("VYear") & " and Vtype='" & .Fields("VType") & "' and Vno=" & .Fields("Vno")
''                If GProcRstOpen(rsNextYearVouSub, Y, "O", gCnToYr) > 0 Then '--- Update TblVouSub Balance
''                   TAdd = (rsNextYearVouSub.Fields("NextYeaeSubAmt") * rsLastYearAddless.Fields("AdOtherAdd")) / rsLastYearAddless.Fields("AdSubAmt")
''                   TLess = (rsNextYearVouSub.Fields("NextYeaeSubAmt") * rsLastYearAddless.Fields("AdOtherLess")) / rsLastYearAddless.Fields("AdSubAmt")
''                   TAddAftTax = (rsNextYearVouSub.Fields("NextYeaeSubAmt") * rsLastYearAddless.Fields("AdAftTaxAdd")) / rsLastYearAddless.Fields("AdSubAmt")
''                   TLessAftTax = (rsNextYearVouSub.Fields("NextYeaeSubAmt") * rsLastYearAddless.Fields("AdAftTaxLess")) / rsLastYearAddless.Fields("AdSubAmt")
''                   TTax = (rsNextYearVouSub.Fields("NextYeaeSubAmt") * rsLastYearAddless.Fields("AdTaxAmt")) / rsLastYearAddless.Fields("AdSubAmt")
''
''                   TBillAmt = rsNextYearVouSub.Fields("NextYeaeSubAmt") + TAdd - TLess + TTax + TAddAftTax - TLessAftTax
''
''                   rsNextYearAddless!AdSubAmt = rsNextYearVouSub.Fields("NextYeaeSubAmt")
''                   rsNextYearAddless!AdOtherAdd = TAdd
''                   rsNextYearAddless!AdOtherLess = TLess
''                   rsNextYearAddless!AdAftTaxAdd = TAddAftTax
''                   rsNextYearAddless!AdAftTaxLess = TLessAftTax
''                   rsNextYearAddless!AdTaxAmt = TTax
''                   rsNextYearAddless!AdRoff = 0
''                   rsNextYearAddless!AdBillAmt = TBillAmt
''                   DoEvents
''                   rsNextYearAddless.Update
''                End If
''            End If
''            .MoveNext
''        Loop
''    End With
''End If
'On Error GoTo ErrorRoutine
Dim i As Long
Dim TblMstCount As Integer
Dim gCnToYr As Connection '- Connection For Next Yr
Dim Vno As Long
Dim vamt As Double
Dim VYear As Long
Dim AcOCode As Long
Dim rsBalS As Recordset '- Balance Source Recordset
Dim rsBalS1 As Recordset
Dim rsBalT As Recordset '- Balance Target Recordset
Dim rsBalT1 As Recordset '- Balance Target Recordset
Dim X As String
Dim DtStr As String
GProcConnectionOpen gCnToYr, ToDbNm
X = " Select  *  from tblVouSub where Vtype in ('PY','PT','PI')  and " _
& "  VSubItCtrlNo>0 and (VSubBag-VSubIssBag)>0  order by vtype,vno"
If GProcRstOpen(rsBalS, X, "R") > 0 Then
    With rsBalS
        Do While Not .EOF
            X = "Select * from Tblvousub Where VYear=" & .Fields("VYear") & " and Vtype='" & .Fields("VType") & "' and VSubItCtrlNo=-" & .Fields("VSubItCtrlNo")
            If GProcRstOpen(rsBalT, X, "O", gCnToYr) > 0 Then '--- Update TblVouSub Balance
                rsBalT!vsubBAG = .Fields("VSubBag") - .Fields("VSubIssBag") '- Bag
                rsBalT!vsubwt = .Fields("VSubWt") - .Fields("VSubIssWt") '-Wt
                rsBalT.Update
            Else
                X = "Select * from Tbladdless Where VYear=" & .Fields("VYear") & " and VType='" & .Fields("VType") & "' and VNo=" & .Fields("VNo")
                GProcRstOpen rsBalS1, X, "R"
                X = "Select * from Tbladdless Where VYear=" & .Fields("VYear") & " and VType='" & .Fields("VType") & "' and VNo=" & .Fields("VNo")
                
                If GProcRstOpen(rsBalT1, X, "O", gCnToYr) = 0 Then
                    rsBalT1.AddNew
                    For i = 0 To rsBalS1.Fields.Count - 1 '- Add TblAddLess Entry
                        rsBalT1.Fields(i) = rsBalS1.Fields(i)
                    Next i
                    rsBalT1.Update
                End If
                rsBalT.AddNew
                For i = 0 To .Fields.Count - 1 '- Add TblVouSub Entry
                    rsBalT.Fields(i) = .Fields(i)
                Next i
                'rsBalT.Update
                rsBalT!vsubAMT = (.Fields("VSubAmt") / .Fields("VSubBag")) * (.Fields("VSubBag") - .Fields("VSubIssBag"))
                DoEvents
                rsBalT!vsubBAG = .Fields("VSubBag") - .Fields("VSubIssBag") '- Bag
                rsBalT!vsubwt = .Fields("VSubWt") - .Fields("VSubIssWt") '-Wt
                rsBalT!VSubItCtrlNo = -1 * .Fields("VSubItCtrlNo") '- It Ctrl No
                rsBalT!VSubPItCtrlNo = -1 * .Fields("VSubPItCtrlNo") '-P It Ctrl No
                rsBalT.Update
            End If
            .MoveNext
        Loop
    End With
End If

'----For Closing Stock Value Transfer
'---For TblVouSub - VsubAmt

GProcConnectionOpen gCnToYr, ToDbNm
X = " Select  *  from tblVouSub where Vtype in ('PY','PT')  and " _
& "  VSubItCtrlNo>0 and (VSubBag-VSubIssBag)>0  order by vtype,vno"
If GProcRstOpen(rsBalS, X, "R") > 0 Then
    With rsBalS
        Do While Not .EOF
            X = "Select * from Tblvousub Where VYear=" & .Fields("VYear") & " and Vtype='" & .Fields("VType") & "' and VSubItCtrlNo=-" & .Fields("VSubItCtrlNo")
            If GProcRstOpen(rsBalT, X, "O", gCnToYr) > 0 Then '--- Update TblVouSub Balance
               rsBalT!vsubAMT = (.Fields("VSubAmt") / .Fields("VSubBag")) * (.Fields("VSubBag") - .Fields("VSubIssBag"))
               DoEvents
               rsBalT.Update
            End If
            .MoveNext
        Loop
    End With
End If
'---For TblAddless - AdSubAmt,AdOtherLess etc

Dim rsLastYearAddless As Recordset
Dim rsNextYearAddless As Recordset
Dim rsNextYearVouSub As Recordset
Dim Y As String
Dim TAdd As Double
Dim TLess As Double
Dim TAddAftTax As Double
Dim TLessAftTax As Double
Dim TTax As Double
Dim TBillAmt As Double
Dim TTcs As Double


GProcConnectionOpen gCnToYr, ToDbNm
X = "select * from tblAddless where vno in(select vno from tblVouSub where  Vtype in ('PY')  and  VSubItCtrlNo>0 and (VSubBag-VSubIssBag)>0    ) and vtype in('PY') and Vyear <= " & gCYear & " order by vtype,vno  "
If GProcRstOpen(rsLastYearAddless, X, "R") > 0 Then
    With rsLastYearAddless
        Do While Not .EOF
            X = "Select * from TblAddless Where VYear=" & .Fields("VYear") & " and Vtype='" & .Fields("VType") & "' and Vno=" & .Fields("Vno")
            If GProcRstOpen(rsNextYearAddless, X, "O", gCnToYr) > 0 Then '--- Update TblVouSub Balance
                Y = "Select sum(VsubAmt) as NextYeaeSubAmt from tblVouSub Where VYear=" & .Fields("VYear") & " and Vtype='" & .Fields("VType") & "' and Vno=" & .Fields("Vno")
                If GProcRstOpen(rsNextYearVouSub, Y, "O", gCnToYr) > 0 Then '--- Update TblVouSub Balance
                   If IsNull(rsNextYearVouSub.Fields("NextYeaeSubAmt")) Then
                   Else
                        TAdd = (rsNextYearVouSub.Fields("NextYeaeSubAmt") * rsLastYearAddless.Fields("AdOtherAdd")) / rsLastYearAddless.Fields("AdSubAmt")
                        TLess = (rsNextYearVouSub.Fields("NextYeaeSubAmt") * rsLastYearAddless.Fields("AdOtherLess")) / rsLastYearAddless.Fields("AdSubAmt")
                        TAddAftTax = (rsNextYearVouSub.Fields("NextYeaeSubAmt") * rsLastYearAddless.Fields("AdAftTaxAdd")) / rsLastYearAddless.Fields("AdSubAmt")
                        TLessAftTax = (rsNextYearVouSub.Fields("NextYeaeSubAmt") * rsLastYearAddless.Fields("AdAftTaxLess")) / rsLastYearAddless.Fields("AdSubAmt")
                        TTax = (rsNextYearVouSub.Fields("NextYeaeSubAmt") * rsLastYearAddless.Fields("AdTaxAmt")) / rsLastYearAddless.Fields("AdSubAmt")
                        TTcs = (rsNextYearVouSub.Fields("NextYeaeSubAmt") * rsLastYearAddless.Fields("AdTcsAmt")) / rsLastYearAddless.Fields("AdSubAmt")
                        
                        TBillAmt = rsNextYearVouSub.Fields("NextYeaeSubAmt") + TAdd + TTax + TAddAftTax - TLessAftTax + TTcs
                        
                        rsNextYearAddless!AdSubAmt = rsNextYearVouSub.Fields("NextYeaeSubAmt")
                        rsNextYearAddless!AdOtherAdd = TAdd
                        rsNextYearAddless!AdOtherLess = TLess
                        rsNextYearAddless!AdAftTaxAdd = TAddAftTax
                        rsNextYearAddless!AdAftTaxLess = TLessAftTax
                        rsNextYearAddless!AdTaxAmt = TTax
                        rsNextYearAddless!AdRoff = 0
                        rsNextYearAddless!AdBillAmt = TBillAmt
                        rsNextYearAddless!AdTcsAmt = TTcs
                   End If
                   DoEvents
                   rsNextYearAddless.Update
                End If
            End If
            .MoveNext
        Loop
    End With
End If


MsgBox "Utility Carry Forward Stock is Completed Successfully.", vbOKOnly + vbInformation, "Utility"
Exit Sub
End Sub



''''On Error GoTo ErrorRoutine
'''Dim i As Long
'''Dim TblMstCount As Integer
'''Dim gCnToYr As Connection '- Connection For Next Yr
'''Dim Vno As Long
'''Dim vamt As Double
'''Dim VYear As Long
'''Dim AcOCode As Long
'''Dim rsBalS As Recordset '- Balance Source Recordset
'''Dim rsBalS1 As Recordset
'''Dim rsBalT As Recordset '- Balance Target Recordset
'''Dim rsBalT1 As Recordset '- Balance Target Recordset
'''Dim X As String
'''Dim DtStr As String
'''GProcConnectionOpen gCnToYr, ToDbNm
'''X = " Select  *  from tblVouSub where Vtype in ('PY','PT','PI')  and " _
'''& "  VSubItCtrlNo>0 and (VSubBag-VSubIssBag)>0  order by vtype,vno"
'''If GProcRstOpen(rsBalS, X, "R") > 0 Then
'''    With rsBalS
'''        Do While Not .EOF
'''            X = "Select * from Tblvousub Where VYear=" & .Fields("VYear") & " and Vtype='" & .Fields("VType") & "' and VSubItCtrlNo=-" & .Fields("VSubItCtrlNo")
'''            If GProcRstOpen(rsBalT, X, "O", gCnToYr) > 0 Then '--- Update TblVouSub Balance
'''                rsBalT!vsubBAG = .Fields("VSubBag") - .Fields("VSubIssBag") '- Bag
'''                rsBalT!vsubwt = .Fields("VSubWt") - .Fields("VSubIssWt") '-Wt
'''                rsBalT.Update
'''            Else
'''                X = "Select * from Tbladdless Where VYear=" & .Fields("VYear") & " and VType='" & .Fields("VType") & "' and VNo=" & .Fields("VNo")
'''                GProcRstOpen rsBalS1, X, "R"
'''                X = "Select * from Tbladdless Where VYear=" & .Fields("VYear") & " and VType='" & .Fields("VType") & "' and VNo=" & .Fields("VNo")
'''
'''                If GProcRstOpen(rsBalT1, X, "O", gCnToYr) = 0 Then
'''                    rsBalT1.AddNew
'''                    For i = 0 To rsBalS1.Fields.Count - 1 '- Add TblAddLess Entry
'''                        rsBalT1.Fields(i) = rsBalS1.Fields(i)
'''                    Next i
'''                    rsBalT1.Update
'''                End If
'''                rsBalT.AddNew
'''                For i = 0 To .Fields.Count - 1 '- Add TblVouSub Entry
'''                    rsBalT.Fields(i) = .Fields(i)
'''                Next i
'''                'rsBalT.Update
'''                rsBalT!vsubBAG = .Fields("VSubBag") - .Fields("VSubIssBag") '- Bag
'''                rsBalT!vsubwt = .Fields("VSubWt") - .Fields("VSubIssWt") '-Wt
'''                rsBalT!VSubItCtrlNo = -1 * .Fields("VSubItCtrlNo") '- It Ctrl No
'''                rsBalT!VSubPItCtrlNo = -1 * .Fields("VSubPItCtrlNo") '-P It Ctrl No
'''                rsBalT.Update
'''            End If
'''            .MoveNext
'''        Loop
'''    End With
'''End If
''''-- Delete Purchase in Next Year for (Complete & No Vs GatePass in Next Year)
''''X = " Select  *  from tblVouSub where Vtype in ('PY','PT','PI')  and VSubItCtrlNo>0 and (VSubBag-VSubIssBag)<=0 "
''''If GProcRstOpen(rsBalS, X, "R") > 0 Then
''''    With rsBalS
''''        Do While Not .EOF
''''            gCnToYr.Execute "delete from tblVouSub where Vno not in ( select vno from tblVouSub as tblGp where " _
''''            & " VSubPItCtrlNo=-" & .Fields("VSubItCtrlNo") & " and  tblGp.Vtype='GP' and  tblGp.VYear=" & .Fields("VYear") & ") " _
''''            & " and VYear=" & .Fields("VYear") & " and Vtype='" & .Fields("VType") & "' and VNo=" & .Fields("VNo")
''''            rsBalS.MoveNext
''''        Loop
''''    End With
''''End If
'''MsgBox "Utility Carry Forward Stock is Completed Successfully.", vbOKOnly + vbInformation, "Utility"
'''Exit Sub
''End Sub


Public Sub GProcUtlCarryForwardTINNo(ToDbNm As String)
'On Error GoTo ErrorRoutine
Dim i As Long
Dim TblMstCount As Integer
Dim gCnToYr As Connection '- Connection For Next Yr
Dim Vno As Long
Dim vamt As Double
Dim VYear As Long
Dim AcOCode As Long
Dim rsBalS As Recordset '- Balance Source Recordset
Dim rsBalS1 As Recordset
Dim rsBalT As Recordset '- Balance Target Recordset
Dim rsBalT1 As Recordset '- Balance Target Recordset
Dim X As String
Dim DtStr As String
GProcConnectionOpen gCnToYr, ToDbNm
'=================== Master (No of tables = 4) New Transaction Entry ================================
X = " Select  *  from tblMastAccount "
If GProcRstOpen(rsBalS, X, "R") > 0 Then
    With rsBalS
         Do While Not .EOF
         
            X = " Select  *  from tblMastAccount where AcName= '" & .Fields("AcName") & "'"
            If GProcRstOpen(rsBalT, X, "O", gCnToYr) <> 0 Then
                If IsNull(rsBalT.Fields(4)) = True Then
                    rsBalT.Fields(4) = .Fields(4)
                End If
                If IsNull(rsBalT.Fields(5)) = True Then
                    rsBalT.Fields(5) = .Fields(5)
                End If
                If IsNull(rsBalT.Fields(7)) = True Then
                    rsBalT.Fields(7) = .Fields(7)
                End If
                If IsNull(rsBalT.Fields(8)) = True Then
                    rsBalT.Fields(8) = .Fields(8)
                End If
                If IsNull(rsBalT.Fields(9)) = True Then
                    rsBalT.Fields(9) = .Fields(9)
                End If
                If IsNull(rsBalT.Fields(10)) = True Then
                    rsBalT.Fields(10) = .Fields(10)
                End If
                If IsNull(rsBalT.Fields(11)) = True Then
                    rsBalT.Fields(11) = .Fields(11)
                End If
                If IsNull(rsBalT.Fields(12)) = True Then
                    rsBalT.Fields(12) = .Fields(12)
                End If
                If IsNull(rsBalT.Fields(13)) = True Then
                    rsBalT.Fields(13) = .Fields(13)
                End If
                If IsNull(rsBalT.Fields(14)) = True Then
                   rsBalT.Fields(14) = .Fields(14)
                End If
                If IsNull(rsBalT.Fields(15)) = True Then
                   rsBalT.Fields(15) = .Fields(15)
                End If
                If IsNull(rsBalT.Fields(22)) = True Then
                   rsBalT.Fields(22) = .Fields(22)
                End If
                If IsNull(rsBalT.Fields(23)) = True Then
                   rsBalT.Fields(23) = .Fields(23)
                End If
                If IsNull(rsBalT.Fields(36)) = True Then
                   rsBalT.Fields(36) = .Fields(36)
                End If
                If IsNull(rsBalT.Fields(37)) = True Then
                   rsBalT.Fields(37) = .Fields(37)
                End If
                If IsNull(rsBalT.Fields(38)) = True Then
                   rsBalT.Fields(38) = .Fields(38)
                End If
                If IsNull(rsBalT.Fields(39)) = True Then
                   rsBalT.Fields(39) = .Fields(39)
                End If
                If IsNull(rsBalT.Fields(40)) = True Then
                   rsBalT.Fields(40) = .Fields(40)
                End If
                rsBalT.Update
             End If
             .MoveNext
         Loop
    End With
    MsgBox ("Data Transfer Successfully.............")
End If
End Sub

Public Sub GProcStockValutation(pToDt As Date)
                 Dim X1 As String
                 gCn.Execute "delete from tmpGenTbl2 where username='" & gUserName & "'"
                 gCn.Execute "delete from tmpStockCal "
                  
                  X1 = "SELECT  '" & gUserName & "', tblVousub.VSubItCode, tblVousub.VSubItCtrlNo,   sum(iif(tblVousub_Out.vtype='RY',-tblVousub_Out.vsubbag,tblVousub_Out.vsubBag)), sum(iif(tblVousub_Out.vtype= 'RY',(tblVousub_Out.VSUBWt*-1), tblVousub_Out.VSUBWt))  " _
                            & "" _
                            & " FROM tblVousub, tblVousub  tblVousub_Out  where  tblVousub.Vtype in ('PY','PT','PI') and tblVousub.Vdt<=cdate('" & pToDt & "')    and tblVousub_Out.Vdt<=cdate('" & pToDt & "') and  tblVousub.VSubItCtrlNo=tblVousub_Out.VSubPItCtrlNo AND TBLVOUSUB.VSUBTMP3 <> 1" _
                            & " GROUP BY tblVousub.VSubItCode,tblVousub.VSubItCtrlNo"
                 gCn.Execute "insert into tmpStockCal (Username,ItemCode,vno,amt1,amt2) " & X1
                  
                  X1 = "SELECT  '" & gUserName & "',tblVousub.vno, tblVousub.Vdt, tblVousub.vyear, tblVousub.VSubItCtrlNo, iif(isnull(tmpStockCal.Amt1),0,tmpStockCal.Amt1),iif(isnull(tmpStockCal.Amt2),0,tmpStockCal.Amt2) , " _
                            & " tblVousub.VSubTaxableAmt / tblVousub.VSubWt ,tblVousub.VSubWt,tblVousub.VSubItCode  " _
                            & " FROM tblVousub LEFT outer JOIN tmpStockCal ON tblVousub.VSubItCtrlNo=tmpStockCal.Vno " _
                            & " WHERE TBLVOUSUB.VSUBTMP3 <> 1 AND tblVousub.vtype In ('PY','PT','PI') and tblVousub.Vdt<=cdate('" & pToDt & "')  " _
                            & " "
                  
'                  X1 = "SELECT  '" & gUserName & "',   tblVousub.VSubItCtrlNo, iif(isnull(tmpStockCal.Amt1),0,tmpStockCal.Amt1),iif(isnull(tmpStockCal.Amt2),0,tmpStockCal.Amt2) , " _
'                            & " tblVousub.VsubAmt / tblVousub.VSubWt ,tblVousub.VSubWt,tblVousub.VSubItCode, (tblVousub.VSubWt - iif(isnull(tmpStockCal.Amt2),0,tmpStockCal.Amt2)) * (tblVousub.VsubAmt / tblVousub.VSubWt)  ,(tblVousub.VSubWt - iif(isnull(tmpStockCal.Amt2),0,tmpStockCal.Amt2)) * (tblVousub.VsubAmt / tblVousub.VSubWt)      " _
'                            & " FROM tblVousub LEFT JOIN tmpStockCal ON tblVousub.VSubItCtrlNo=tmpStockCal.Vno " _
'                            & " WHERE tblVousub.vtype In ('PY','PT','PI') and tblVousub.Vdt<=cdate('" & pToDt & "')  " _
'                            & " "
                 
                 gCn.Execute "insert into tmpgentbl2 (Username,Amt7,vdt,vyear,vno,amt1,amt2,Amt3,Amt4,VACCODE) " & X1

                 gCn.Execute "update  tmpgentbl2  set Amt10 = (amt4 - amt2) * amt3 where  username='" & gUserName & "'"
End Sub
Public Sub GprocPrepareSalePurSummGST(pFromDt As Date, pToDt As Date, SlPu As String)
Dim X1 As String
gCn.Execute "delete from tmpGenTbl2 where username='" & gUserName & "'"

If SlPu = "S" Then
   X1 = "Select AcName,AcAdd1,AcAdd2,AcCity,AcPAN,sum(AdSubAmt),sum(VsubCGSTAmt+VsubSGSTAmt+VsubIGSTAmt),sum(AdBillAmt),'" & gUserName & "' from tblAddLess,TblVouSub,tblMastAccount where tblAddLess.Vdt >= cdate('" & pFromDt & "')  and  tblAddLess.Vdt<=cdate('" & pToDt & "') " _
        & " and tblAddLess.Vno = tblVouSub.Vno and tblAddLess.Vtype = tblVouSub.VType and tblAddLess.Vyear = tblVouSub.Vyear " _
        & " and tblAddLess.Vtype in ('SY','SO') and tblAddLess.AdAcDrCode = tblMastAccount.AcCode group by AcName,AcAdd1,AcAdd2,AcCity,AcPAN"
Else
   X1 = "Select AcName,AcAdd1,AcAdd2,AcCity,AcPAN,sum(VSubTaxableAmt),sum(VsubCGSTAmt+VsubSGSTAmt+VsubIGSTAmt),sum(AdBillAmt),'" & gUserName & "' from tblAddLess,TblVouSub,tblMastAccount where tblAddLess.Vdt >= cdate('" & pFromDt & "')  and  tblAddLess.Vdt<=cdate('" & pToDt & "') " _
        & " and tblAddLess.Vno = tblVouSub.Vno and tblAddLess.Vtype = tblVouSub.VType and tblAddLess.Vyear = tblVouSub.Vyear " _
        & " and tblAddLess.Vtype in ('PY') and tblAddLess.AdAcCrCode = tblMastAccount.AcCode group by AcName,AcAdd1,AcAdd2,AcCity,AcPAN"
End If

gCn.Execute "insert into tmpGenTbl2 (nar1,nar2,nar3,nar4,nar5,amt1,Amt2,Amt3,Username) " & X1

End Sub


Public Sub GProcUtlCarryForwardPartyBooking(ToDbNm As String)
Dim i As Long
Dim TblMstCount As Integer
Dim gCnToYr As Connection '- Connection For Next Yr
Dim Vno As Long
Dim vamt As Double
Dim VYear As Long
Dim AcOCode As Long
Dim rsBalS As Recordset '- Balance Source Recordset
Dim rsBalS1 As Recordset
Dim rsBalT As Recordset '- Balance Target Recordset
Dim rsBalT1 As Recordset '- Balance Target Recordset
Dim X As String
Dim DtStr As String
GProcConnectionOpen gCnToYr, ToDbNm

gCnToYr.Execute "delete from tblvousub_Book where vyear = " & gCYear & " and vtype = 'OP'"
gCnToYr.Execute "delete from tblAddLess_Book where vyear = " & gCYear & " and vtype = 'OP'"


X = " Select  *  from tblvousub_Book where Vtype = 'OP' and " _
& "  VSubItCtrlNo>0 and VSubBookIsCompleted=0 and (VSubBag-VSubIssBag)>0"
If GProcRstOpen(rsBalS, X, "R") > 0 Then
    With rsBalS
        Do While Not .EOF
            X = "Select * from tblvousub_Book Where VYear=" & .Fields("VYear") & " and Vtype='" & .Fields("VType") & "' and VSubItCtrlNo=-" & .Fields("VSubItCtrlNo")
            If GProcRstOpen(rsBalT, X, "O", gCnToYr) > 0 Then '--- Update TblVouSub Balance
                rsBalT!vsubBAG = .Fields("VSubBag") - .Fields("VSubIssBag") '- Bag
                rsBalT!vsubwt = .Fields("VSubWt") - .Fields("VSubIssWt") '-Wt
                rsBalT.Update
            Else
                X = "Select * from tblAddLess_Book Where VYear=" & .Fields("VYear") & " and VType='" & .Fields("VType") & "' and VNo=" & .Fields("VNo")
                GProcRstOpen rsBalS1, X, "R"
                X = "Select * from tblAddLess_Book Where VYear=" & .Fields("VYear") & " and VType='" & .Fields("VType") & "' and VNo=" & .Fields("VNo")
                If GProcRstOpen(rsBalT1, X, "O", gCnToYr) = 0 Then
                    rsBalT1.AddNew
                    For i = 0 To rsBalS1.Fields.Count - 1 '- Add TblAddLess Entry
                        rsBalT1.Fields(i) = rsBalS1.Fields(i)
                    Next i
                    rsBalT1.Update
                    rsBalT.AddNew
                    For i = 0 To .Fields.Count - 1 '- Add TblVouSub Entry
                        rsBalT.Fields(i) = .Fields(i)
                    Next i
                    'rsBalT.Update
                    rsBalT!vsubBAG = .Fields("VSubBag") - .Fields("VSubIssBag") '- Bag
                    rsBalT!vsubwt = .Fields("VSubWt") - .Fields("VSubIssWt") '-Wt
                    rsBalT!VSubItCtrlNo = -1 * .Fields("VSubItCtrlNo") '- It Ctrl No
                    rsBalT!VSubPItCtrlNo = -1 * .Fields("VSubPItCtrlNo") '-P It Ctrl No
                    rsBalT.Update
                Else
                    rsBalT.AddNew
                    For i = 0 To .Fields.Count - 1 '- Add TblVouSub Entry
                        rsBalT.Fields(i) = .Fields(i)
                    Next i
                    'rsBalT.Update
                    rsBalT!vsubBAG = .Fields("VSubBag") - .Fields("VSubIssBag") '- Bag
                    rsBalT!vsubwt = .Fields("VSubWt") - .Fields("VSubIssWt") '-Wt
                    rsBalT!VSubItCtrlNo = -1 * .Fields("VSubItCtrlNo") '- It Ctrl No
                    rsBalT!VSubPItCtrlNo = -1 * .Fields("VSubPItCtrlNo") '-P It Ctrl No
                    rsBalT.Update
                End If
            End If
            .MoveNext
        Loop
    End With
End If

MsgBox "Utility Carry Forward Party Booking is Completed Successfully.", vbOKOnly + vbInformation, "Utility"

Exit Sub
End Sub
Public Sub GProcUtlCarryForwardTradeBooking(ToDbNm As String)
Dim i As Long
Dim TblMstCount As Integer
Dim gCnToYr As Connection '- Connection For Next Yr
Dim Vno As Long
Dim vamt As Double
Dim VYear As Long
Dim AcOCode As Long
Dim rsBalS As Recordset '- Balance Source Recordset
Dim rsBalS1 As Recordset
Dim rsBalT As Recordset '- Balance Target Recordset
Dim rsBalT1 As Recordset '- Balance Target Recordset
Dim X As String
Dim DtStr As String
GProcConnectionOpen gCnToYr, ToDbNm

gCnToYr.Execute "delete from tblvousub where vyear = " & gCYear & " and vtype = 'OT'"
gCnToYr.Execute "delete from tblAddLess where vyear = " & gCYear & " and vtype = 'OT'"

gCnToYr.Execute "update tblvousub set VSubPItCtrlNo = 0,VSubBookNo = 0 where vyear = " & gCYear & " and vtype = 'PY'"

X = " Select  *  from tblVouSub where Vtype = 'OT' and " _
& "  VSubItCtrlNo>0 and VSubBookIsCompleted=0 and (VSubBag-VSubIssBag)>0"
If GProcRstOpen(rsBalS, X, "R") > 0 Then
    With rsBalS
        Do While Not .EOF
            X = "Select * from Tblvousub Where VYear=" & .Fields("VYear") & " and Vtype='" & .Fields("VType") & "' and VSubItCtrlNo=-" & .Fields("VSubItCtrlNo")
            If GProcRstOpen(rsBalT, X, "O", gCnToYr) > 0 Then '--- Update TblVouSub Balance
                rsBalT!vsubBAG = .Fields("VSubBag") - .Fields("VSubIssBag") '- Bag
                rsBalT!vsubwt = .Fields("VSubWt") - .Fields("VSubIssWt") '-Wt
                rsBalT.Update
            Else
                X = "Select * from Tbladdless Where VYear=" & .Fields("VYear") & " and VType='" & .Fields("VType") & "' and VNo=" & .Fields("VNo")
                GProcRstOpen rsBalS1, X, "R"
                X = "Select * from Tbladdless Where VYear=" & .Fields("VYear") & " and VType='" & .Fields("VType") & "' and VNo=" & .Fields("VNo")
                If GProcRstOpen(rsBalT1, X, "O", gCnToYr) = 0 Then
                    rsBalT1.AddNew
                    For i = 0 To rsBalS1.Fields.Count - 1 '- Add TblAddLess Entry
                        rsBalT1.Fields(i) = rsBalS1.Fields(i)
                    Next i
                    rsBalT1.Update
                    rsBalT.AddNew
                    For i = 0 To .Fields.Count - 1 '- Add TblVouSub Entry
                        rsBalT.Fields(i) = .Fields(i)
                    Next i
                    'rsBalT.Update
                    rsBalT!vsubBAG = .Fields("VSubBag") - .Fields("VSubIssBag") '- Bag
                    rsBalT!vsubwt = .Fields("VSubWt") - .Fields("VSubIssWt") '-Wt
                    rsBalT!VSubItCtrlNo = -1 * .Fields("VSubItCtrlNo") '- It Ctrl No
                    rsBalT!VSubPItCtrlNo = -1 * .Fields("VSubPItCtrlNo") '-P It Ctrl No
                    rsBalT.Update
                Else
                    rsBalT.AddNew
                    For i = 0 To .Fields.Count - 1 '- Add TblVouSub Entry
                        rsBalT.Fields(i) = .Fields(i)
                    Next i
                    'rsBalT.Update
                    rsBalT!vsubBAG = .Fields("VSubBag") - .Fields("VSubIssBag") '- Bag
                    rsBalT!vsubwt = .Fields("VSubWt") - .Fields("VSubIssWt") '-Wt
                    rsBalT!VSubItCtrlNo = -1 * .Fields("VSubItCtrlNo") '- It Ctrl No
                    rsBalT!VSubPItCtrlNo = -1 * .Fields("VSubPItCtrlNo") '-P It Ctrl No
                    rsBalT.Update
                End If
            End If
            .MoveNext
        Loop
    End With
End If

MsgBox "Utility Carry Forward Trade Purchase Booking is Completed Successfully.", vbOKOnly + vbInformation, "Utility"

Exit Sub
End Sub


