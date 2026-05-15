Attribute VB_Name = "Account"
Option Explicit
Dim CommAmt As Double
Dim CommOn As Double
Dim Frm1 As Form
Dim CommBag As Long
'== Procedure For Mill Bill Ledger
'[ Insert To Table -> TmpGenTbl   From -> View qryAccLedgerMillBill ]
Public Sub GProcPrepareMillBillLedger(pFromDt As Date, pToDate As Date, Optional pAcCode As Long)
Dim X As String '- Condition string
'-delete from TmpGenTbl
gCn.Execute "delete from TmpGenTbl where Username='" & gUserName & "'"
'- insert Opening
If gBackEndDB = gBackEndAccess Then
    X = " vdt<cdate('" & pFromDt & "')"
Else
    X = " vdt<to_date('" & pFromDt & "','dd/MM/yyyy')"
End If
If pAcCode > 0 Then X = X & " and AcCode=" & pAcCode
If gBackEndDB = gBackEndAccess Then
    X = " select '" & gUserName & "',cdate('" & pFromDt & "') ,'OP',AcName,'Opening Balance',AgName,0,iif(sum(Debit-Credit)>0,sum(debit-Credit),0) as Debit," _
    & " iif(sum(Credit-Debit)>0,sum(Credit-debit),0) as Credit from qryAccLedgerMillBill " _
    & " where " & X & " group by AcName,AgName having sum(Debit-Credit)<>0"
Else '- Oracle
    X = " select '" & gUserName & "',to_date('" & pFromDt & "','dd/MM/yyyy') ,'OP',AcName,'Opening Balance',AgName,0,decode(sign(sum(Debit-Credit)),1,sum(debit-Credit),0) as Debit," _
    & " decode(sign(sum(Credit-Debit)),1,sum(Credit-debit),0) as Credit from qryAccLedgerMillBill " _
    & " where " & X & " group by AcName,AgName having sum(Debit-Credit)<>0"
End If
gCn.Execute " insert into TmpGenTbl (UserName,VbillDt,Vtype,Nar1,Nar2,Nar3,VSrno,amt1,amt2) " & X
'- insert Dr ,Cr
If gBackEndDB = gBackEndAccess Then
    X = " vdt>=cdate('" & pFromDt & "') and Vdt<=Cdate('" & pToDate & "')"
Else
    X = " vdt>=to_date('" & pFromDt & "','dd/MM/yyyy') and Vdt<=to_date('" & pToDate & "','dd/MM/yyyy')"
End If
If pAcCode > 0 Then X = X & " and AcCode=" & pAcCode
gCn.Execute " insert into TmpGenTbl (UserName,Vno,vbillNo,VbillDt,Vtype,Nar1,Nar2,Nar3,VSrno,Nar7,Nar8,Nar9,Nar10,amt1,amt2) " _
& " select '" & gUserName & "',Vno,BillNo,Vdt,Vtype,AcName,AcOName,AgName,Orderno,Nar1,Nar2,Nar3,Nar4,Debit,Credit from qryAccLedgerMillBill " _
& " where " & X
'--- Update to set Narration = Receipt Vs Mill Bill Inv. Details
If gBackEndDB = gBackEndAccess Then
    gCn.Execute "update TmpGenTbl ,tmpRptTbl set TmpGenTbl.Nar13= tmpRptTbl.Nar1  where tmpRptTbl.VType=TmpGenTbl.Vtype " _
    & " and tmpRptTbl.Vno=TmpGenTbl.Vno and tmpRptTbl.VDt=TmpGenTbl.VDt and tmpRptTbl.UserName=TmpGenTbl.UserName and TmpGenTbl.Vtype='MR' and TmpGenTbl.UserName='" & gUserName & "'"
Else
    gCn.Execute "update TmpGenTbl set Nar13=(select Nar1  from tmpRptTbl where tmpRptTbl.VType=TmpGenTbl.Vtype " _
    & " and tmpRptTbl.Vno=TmpGenTbl.Vno and tmpRptTbl.VDt=TmpGenTbl.VDt and tmpRptTbl.UserName=TmpGenTbl.UserName ) where TmpGenTbl.Vtype='MR' and TmpGenTbl.UserName='" & gUserName & "'"
End If
End Sub
'--- Get Balance Sheet Record In Tmp Tbl Of Given Top Group Code
Private Sub GetBalRecordInTmpTbl(TopGrpcode As Long, ToDt As Date)
Dim X As String
Dim x1 As String

'--- Insert Into tmpRptTbl
'--- With out & with Schedule ( For 1 To 5 Group Head) ,closing Stock,Profit &Loss of Yr
gCn.Execute "delete from tmpRptTbl where VFirm = '" & Trim(gCCode) & "'  and UserName='" & gUserName & "'"

'--- Group Head & A/c H2 WithOut Schedule
x1 = "insert into TmpRptTbl (Vtype,Nar1,Amt1,UserName,VFirm,Nar2,Nar3,Nar4,Nar5,Nar6,Nar7,Nar8,Nar9,Nar10,Nar11,Nar12,Amt2,Amt3,vno) "
X = x1 + " select distinct 'B' as Vtype,H2 as Nar,case when H2Type='A' then ClBal_Cr-ClBal_Dr else 0 end as Bal,'" & gUserName & "','" & Trim(gCCode) & "',H1,H2,'' as H3, " _
   & "'' as H4,'' as H5,'' as H6,H1TYPE, H2TYPE,'-' as H3Type,'-' as H4type,'-' as H5Type,0 as IsSchedule,case when H2Type='A' then OpBal_Cr-OpBal_Dr else 0 end " _
   & ",qryAccountGrouping.s2 from qryAccountGrouping,TmpClosingBalance where TopGroup = TopGrp and qryAccountGrouping.AcCode =TmpClosingBalance.AcCode and PfPer=0 and H2<>'' and (isnull(schedul,'') =  '' or (H2Type='G' and ScheduleHno=2 ) ) " _
   & "and TopGroup=" & TopGrpcode & " and VFirm = '" & Trim(gCCode) & "' and UserName= '" & gUserName & "' and  qryAccountGrouping.Agcode = TmpClosingBalance.Agcode  "
gCn.Execute X

'--- Group Head & A/c H3 WithOut Schedule
X = x1 & "select distinct 'B','  '+H3,case when H3Type='A' then ClBal_Cr-ClBal_Dr else 0 end ,'" & gUserName & "','" & Trim(gCCode) & "',H1,H2,H3,'','','' as H6,H1TYPE, H2TYPE,H3TYPE,'-','-' as H5Type,0,case when H3Type='A' then OpBal_Cr-OpBal_Dr else 0 end,qryAccountGrouping.s2 from qryAccountGrouping,TmpClosingBalance where TopGroup = TopGrp and qryAccountGrouping.AcCode =TmpClosingBalance.AcCode and PfPer=0 and  H3 <>'' and ( isnull(schedul,'') =  '' or (H3Type='G' and ScheduleHno=3 ) ) and TopGroup=" & TopGrpcode & " and VFirm = '" & Trim(gCCode) & "' and UserName= '" & gUserName & "' and  qryAccountGrouping.Agcode = TmpClosingBalance.Agcode"
gCn.Execute X
'--- Group Head & A/c H4 WithOut Schedule



gCn.Execute x1 & " select distinct 'B','    '+H4,case when H4Type='A' then ClBal_Cr-ClBal_Dr else 0 end ,'" & gUserName & "','" & Trim(gCCode) & "',H1,H2,H3,H4,'','' as H6,H1TYPE, H2TYPE,H3TYPE,H4Type,'-' as H5Type,0,case when H4Type='A' then OpBal_Cr-OpBal_Dr else 0 end,qryAccountGrouping.s2  from qryAccountGrouping,TmpClosingBalance where TopGroup = TopGrp and qryAccountGrouping.AcCode =TmpClosingBalance.AcCode and PfPer=0 and  H4<>'' and (isnull(schedul,'') =  '' or (H4Type='G' and ScheduleHno=4 ) ) and TopGroup=" & TopGrpcode & " and VFirm = '" & Trim(gCCode) & "' and UserName= '" & gUserName & "' and  qryAccountGrouping.Agcode = TmpClosingBalance.Agcode"



'--- Group Head & A/c H5 WithOut Schedule
gCn.Execute x1 & " select distinct 'B','        '+H5,case when H5Type='A' then ClBal_Cr-ClBal_Dr else 0 end,'" & gUserName & "','" & Trim(gCCode) & "',H1,H2,H3,H4,H5,'' as H6,H1TYPE, H2TYPE,H3TYPE,H4Type,H5Type,0,case when H5Type='A' then OpBal_Cr-OpBal_Dr else 0 end,qryAccountGrouping.s2 from qryAccountGrouping,TmpClosingBalance where TopGroup = TopGrp and qryAccountGrouping.AcCode =TmpClosingBalance.AcCode and PfPer=0 and H5<>'' and (isnull(schedul,'') =  '' or (H5Type='G' and ScheduleHno=5 ) ) and TopGroup=" & TopGrpcode & "  and VFirm = '" & Trim(gCCode) & "' and UserName= '" & gUserName & "' and  qryAccountGrouping.Agcode = TmpClosingBalance.Agcode "
'--- Group Head & A/c H6 WithOut Schedule
gCn.Execute x1 & " select distinct 'B','            '+H6 ,case when H6Type='A' then ClBal_Cr-ClBal_Dr else 0 end ,'" & gUserName & "','" & Trim(gCCode) & "',H1,H2,H3,H4,H5,H6,H1TYPE, H2TYPE,H3TYPE,H4Type,H5Type,0,case when H6Type='A' then OpBal_Cr-OpBal_Dr else 0 end,qryAccountGrouping.s2 from qryAccountGrouping,TmpClosingBalance where TopGroup = TopGrp and qryAccountGrouping.AcCode =TmpClosingBalance.AcCode and PfPer=0 and H6 <>'' and isnull(schedul,'') =  '' and TopGroup=" & TopGrpcode & " and VFirm = '" & Trim(gCCode) & "' and UserName= '" & gUserName & "' and  qryAccountGrouping.Agcode = TmpClosingBalance.Agcode"
'--- Group Head & A/c H2 With Schedule
'gCn.Execute X1 & " select distinct 'B' ,'As Per Schedule '+Schedul ,sum(ClBal_Cr-ClBal_Dr) ,'" & gUserName & "','" & Trim(gCCode) & "',H1,schedul ,'','','','' as H6,H1TYPE, '0','-','-','-' as H5Type,1,sum(OpBal_Cr-OpBal_Dr),qryAccountGrouping.s2 from qryAccountGrouping,TmpClosingBalance where TopGroup = TopGrp and qryAccountGrouping.AcCode =TmpClosingBalance.AcCode and PfPer=0 and ScheduleHNo=1 and isnull(schedul,'') <> '' and TopGroup=" & TopGrpcode & " and VFirm = '" & Trim(gCCode) & "' and UserName= '" & gUserName & "' and  qryAccountGrouping.Agcode = TmpClosingBalance.Agcode group by Schedul,H1,H1Type,username,VFirm,qryAccountGrouping.s2 "
gCn.Execute x1 & " select distinct 'B' ,'As Per Annexure ' ,sum(ClBal_Cr-ClBal_Dr) ,'" & gUserName & "','" & Trim(gCCode) & "',H1,schedul ,'','','','' as H6,H1TYPE, '0','-','-','-' as H5Type,1,sum(OpBal_Cr-OpBal_Dr),qryAccountGrouping.s2 from qryAccountGrouping,TmpClosingBalance where TopGroup = TopGrp and qryAccountGrouping.AcCode =TmpClosingBalance.AcCode and PfPer=0 and ScheduleHNo=1 and isnull(schedul,'') <> '' and TopGroup=" & TopGrpcode & " and VFirm = '" & Trim(gCCode) & "' and UserName= '" & gUserName & "' and  qryAccountGrouping.Agcode = TmpClosingBalance.Agcode group by Schedul,H1,H1Type,username,VFirm,qryAccountGrouping.s2 "


'--- Group Head & A/c H3 With Schedule
'gCn.Execute X1 & " select distinct 'B','  As Per Schedule '+Schedul ,sum(ClBal_Cr-ClBal_Dr) ,'" & gUserName & "','" & Trim(gCCode) & "',H1,H2,schedul ,'','','' as H6,H1TYPE, H2TYPE,'0','-','-' as H5Type,1,sum(OpBal_Cr-OpBal_Dr),qryAccountGrouping.s2 from qryAccountGrouping,TmpClosingBalance where TopGroup = TopGrp and qryAccountGrouping.AcCode =TmpClosingBalance.AcCode and PfPer=0 and ScheduleHNo=2 and isnull(schedul,'') <> '' and TopGroup=" & TopGrpcode & " and VFirm = '" & Trim(gCCode) & "' and UserName= '" & gUserName & "' and  qryAccountGrouping.Agcode = TmpClosingBalance.Agcode group by Schedul,H1,H2,H1Type,H2Type,username,VFirm,qryAccountGrouping.s2 "
gCn.Execute x1 & " select distinct 'B','  As Per Annexure ' ,sum(ClBal_Cr-ClBal_Dr) ,'" & gUserName & "','" & Trim(gCCode) & "',H1,H2,schedul ,'','','' as H6,H1TYPE, H2TYPE,'0','-','-' as H5Type,1,sum(OpBal_Cr-OpBal_Dr),qryAccountGrouping.s2 from qryAccountGrouping,TmpClosingBalance where TopGroup = TopGrp and qryAccountGrouping.AcCode =TmpClosingBalance.AcCode and PfPer=0 and ScheduleHNo=2 and isnull(schedul,'') <> '' and TopGroup=" & TopGrpcode & " and VFirm = '" & Trim(gCCode) & "' and UserName= '" & gUserName & "' and  qryAccountGrouping.Agcode = TmpClosingBalance.Agcode group by Schedul,H1,H2,H1Type,H2Type,username,VFirm,qryAccountGrouping.s2 "

'--- Group Head & A/c H4 With Schedule
'gCn.Execute X1 & " select distinct 'B','    As Per Schedule '+Schedul,sum(ClBal_Cr-ClBal_Dr),'" & gUserName & "','" & Trim(gCCode) & "',H1,H2,H3,schedul ,'','' as H6,H1TYPE, H2TYPE,H3TYPE,'0' ,'-' as H5Type,1,sum(OpBal_Cr-OpBal_Dr),qryAccountGrouping.s2 from qryAccountGrouping,TmpClosingBalance where TopGroup = TopGrp and qryAccountGrouping.AcCode =TmpClosingBalance.AcCode and PfPer=0 and ScheduleHNo=3 and isnull(schedul,'') <>  '' and TopGroup=" & TopGrpcode & "  and VFirm = '" & Trim(gCCode) & "' and UserName= '" & gUserName & "' and  qryAccountGrouping.Agcode = TmpClosingBalance.Agcode Group by Schedul,H1,H2,H3,H1Type,H2Type,H3Type,username,VFirm,qryAccountGrouping.s2 "
gCn.Execute x1 & " select distinct 'B','    As Per Annexure ',sum(ClBal_Cr-ClBal_Dr),'" & gUserName & "','" & Trim(gCCode) & "',H1,H2,H3,schedul ,'','' as H6,H1TYPE, H2TYPE,H3TYPE,'0' ,'-' as H5Type,1,sum(OpBal_Cr-OpBal_Dr),qryAccountGrouping.s2 from qryAccountGrouping,TmpClosingBalance where TopGroup = TopGrp and qryAccountGrouping.AcCode =TmpClosingBalance.AcCode and PfPer=0 and ScheduleHNo=3 and isnull(schedul,'') <>  '' and TopGroup=" & TopGrpcode & "  and VFirm = '" & Trim(gCCode) & "' and UserName= '" & gUserName & "' and  qryAccountGrouping.Agcode = TmpClosingBalance.Agcode Group by Schedul,H1,H2,H3,H1Type,H2Type,H3Type,username,VFirm,qryAccountGrouping.s2 "


'--- Group Head & A/c H5 With Schedule
'gCn.Execute X1 & " select distinct 'B','        As Per Schedule '+Schedul ,sum(ClBal_Cr-ClBal_Dr),'" & gUserName & "','" & Trim(gCCode) & "',H1,H2,H3,H4,schedul as  H5,'' as H6,H1TYPE, H2TYPE,H3TYPE,H4Type,'0' as H5Type,1,sum(OpBal_Cr-OpBal_Dr),qryAccountGrouping.s2 from qryAccountGrouping,TmpClosingBalance where TopGroup = TopGrp and qryAccountGrouping.AcCode =TmpClosingBalance.AcCode and PfPer=0 and ScheduleHNo=4 and isnull(schedul,'') <>  ''  and TopGroup=" & TopGrpcode & "  and VFirm = '" & Trim(gCCode) & "' and UserName= '" & gUserName & "' and  qryAccountGrouping.Agcode = TmpClosingBalance.Agcode Group by Schedul,H1,H2,H3,H4,H1Type,H2Type,H3Type,H4Type,username,VFirm,qryAccountGrouping.s2 "
gCn.Execute x1 & " select distinct 'B','        As Per Annexure ' ,sum(ClBal_Cr-ClBal_Dr),'" & gUserName & "','" & Trim(gCCode) & "',H1,H2,H3,H4,schedul as  H5,'' as H6,H1TYPE, H2TYPE,H3TYPE,H4Type,'0' as H5Type,1,sum(OpBal_Cr-OpBal_Dr),qryAccountGrouping.s2 from qryAccountGrouping,TmpClosingBalance where TopGroup = TopGrp and qryAccountGrouping.AcCode =TmpClosingBalance.AcCode and PfPer=0 and ScheduleHNo=4 and isnull(schedul,'') <>  ''  and TopGroup=" & TopGrpcode & "  and VFirm = '" & Trim(gCCode) & "' and UserName= '" & gUserName & "' and  qryAccountGrouping.Agcode = TmpClosingBalance.Agcode Group by Schedul,H1,H2,H3,H4,H1Type,H2Type,H3Type,H4Type,username,VFirm,qryAccountGrouping.s2 "


'--- Group Head & A/c H6 With Schedule
'gCn.Execute X1 & " select distinct 'B','            As Per Schedule '+Schedul,sum(ClBal_Cr-ClBal_Dr) ,'" & gUserName & "','" & Trim(gCCode) & "',H1,H2,H3,H4,H5,schedul as H6 ,H1TYPE, H2TYPE,H3TYPE,H4Type,H5Type,1,sum(OpBal_Cr-OpBal_Dr),qryAccountGrouping.s2 from qryAccountGrouping,TmpClosingBalance where TopGroup = TopGrp and qryAccountGrouping.AcCode =TmpClosingBalance.AcCode and PfPer=0 and isnull(schedul,'') <>  ''  and ScheduleHNo=5 and TopGroup=" & TopGrpcode & " and VFirm = '" & Trim(gCCode) & "' and UserName= '" & gUserName & "' and  qryAccountGrouping.Agcode = TmpClosingBalance.Agcode group by Schedul,H1,H2,H3,H4,H5,H1Type,H2Type,H3Type,H4Type,H5Type,username,VFirm,qryAccountGrouping.s2 "
gCn.Execute x1 & " select distinct 'B','            As Per Annexure ',sum(ClBal_Cr-ClBal_Dr) ,'" & gUserName & "','" & Trim(gCCode) & "',H1,H2,H3,H4,H5,schedul as H6 ,H1TYPE, H2TYPE,H3TYPE,H4Type,H5Type,1,sum(OpBal_Cr-OpBal_Dr),qryAccountGrouping.s2 from qryAccountGrouping,TmpClosingBalance where TopGroup = TopGrp and qryAccountGrouping.AcCode =TmpClosingBalance.AcCode and PfPer=0 and isnull(schedul,'') <>  ''  and ScheduleHNo=5 and TopGroup=" & TopGrpcode & " and VFirm = '" & Trim(gCCode) & "' and UserName= '" & gUserName & "' and  qryAccountGrouping.Agcode = TmpClosingBalance.Agcode group by Schedul,H1,H2,H3,H4,H5,H1Type,H2Type,H3Type,H4Type,H5Type,username,VFirm,qryAccountGrouping.s2 "

'--- Profit & Loss of Current Yr

If TopGrpcode = 90002 Then
    Dim pl As Double
    Dim tmp11 As New ADODB.Recordset
    pl = GProcGetColumnValue("TmpClosingBalance", "topgrp", "90003", "N", "sum(ClBal_Cr-ClBal_Dr)", "N", "VFirm = '" & Trim(gCCode) & "' and username='" & gUserName & "' group by topgrp ", gCn)
    pl = pl - GProcGetColumnValue("TmpClosingBalance", "topgrp", "90001", "N", "sum(ClBal_Dr)", "N", "agcode = 90018 and VFirm = '" & Trim(gCCode) & "' and username='" & gUserName & "' group by topgrp ", gCn)
    pl = pl + GProcGetColumnValue("TblClosingStock", "VFirm", gCCode, "S", "sum(stamount)", "N", "StDate= '" & Format(ToDt, "yyyy/mm/dd") & "' group by VFirm  ", gCn)
'    GProcRstOpen tmp11, "Select dbo.FncPLCalculation(" & gCCode & " ,'" & Format(gCYSDate, "yyyy/mm/dd") & " ','" & Format(ToDt, "yyyy/mm/dd") & "')"
    If GProcGetColumnValue("TmpRptTbl", "username", gUserName, "S", "Amt1", "N", "VFirm = '" & gCCode & "' and nar1 = '  Profit And Loss Of Current Year'") <> 0 Then
       gCn.Execute "Update TmpRptTbl set amt1 = amt1 + " & pl & " where VFirm = '" & gCCode & "' and nar1 = '  Profit And Loss Of Current Year' and username = '" & gUserName & "' "
    Else
      gCn.Execute x1 & " values ('B','  Profit And Loss Of Current Year'," & pl & ",'" & gUserName & "','" & Trim(gCCode) & "','Liability','Capital Account','Profit And Loss Of Current Year','','','','G','G','A','-','-',0,0,2)"
    End If
Else
    '--- Closing Stock Group Head
'    gCn.Execute X1 & " select distinct 'B','  '+tblMastGroup.AgName,0,'" & gUserName & "','" & Trim(gCCode) & "',tblG1.AgName as H1,tblMastGroup.AgName as H2,'' as H3,'' as H4,'' as H5,'' as H6,'G' as H1Type,'G' as H2Type,'-' as H3Type,'-' As H4Type,'-' as H5Type,0,0,170 from TblClosingStock,tblMastAccount,tblMastGroup TblG1,TblMastGroup where " _
'    & " VFirm = '" & gCCode & "' and tblMastGroup.Agcode=TblMastAccount.AgCode and stAccode=AcCode and tblG1.Agcode=tblMastGroup.Gpcode and  tblMastGroup.TopGroup=" & TopGrpcode & " and StAmount>0 and StDate= '" & Format(ToDt, "yyyy/mm/dd") & "'"
'    '--- Closing Stock A/c
'
'    gCn.Execute X1 & " select 'B','    '+A.AcName,-StAmount,'" & gUserName & "','" & Trim(gCCode) & "',tblG1.AgName as H1,tblMastGroup.AgName as H2,A.AcName as H3,'' as H4,'' as H5,'' as H6,'G' as H1Type,'G' as H2Type,'A' as H3Type,'-' As H4Type,'-' as H5Type,0,-B.OpBal_Dr,170 from TblClosingStock,tblMastAccount A,tblMastGroup TblG1,TblMastGroup,TmpClosingBalance B where " _
'    & " TblClosingStock.VFirm = '" & gCCode & "' and  UserName = '" & gUserName & "' and  A.AcName = B.AcName and  tblMastGroup.Agcode=A.AgCode and stAccode=A.AcCode and tblG1.Agcode=tblMastGroup.Gpcode and  tblMastGroup.TopGroup=" & TopGrpcode & " and StDate='" & Format(ToDt, "yyyy/mm/dd") & "'"

End If

End Sub
'--- Profit Loss
Public Sub GProcPrepareProfitLoss(pDtFr As Date, pDtTo As Date)
Dim X As String, mGnm(5) As String
Dim RsTmp_S As New ADODB.Recordset
Dim RsTmp_T As New ADODB.Recordset
Dim i As Long, RCount_T As Long, RCount_S As Long, dt1 As Date
Dim mGp As Double, t1 As Integer, t2 As Integer, P1 As Integer, P2 As Integer
Dim OpStk As Double
Dim PurStk As Double
Dim PurRtnStk As Double
Dim SaleStk As Double
Dim SaleRtnStk As Double
Dim ClStk As Double

'--Yarn OP Stock Qty
X = "select sum(PurSubBag) from tblPurchSub where tblPurchSub.vtype = 'OY' and " _
& " tblPurchSub.VDt = ('" & Format(gCYSDate, "yyyy/mm/dd") & "')  and tblPurchSub.vfirm =  '" & gCCode & "'"
If GProcRstOpen(RsTmp_S, X, "R", gCn) > 0 Then
   If IsNull(RsTmp_S.Fields(0)) = False Then
      OpStk = RsTmp_S.Fields(0)
   Else
      OpStk = 0
   End If
End If
DoEvents
DoEvents
DoEvents
'--Yarn OP Stock Qty + Add Purch Before From Date

X = "select sum(PurSubBag) from tblPurchSub,tblPurch where tblPurch.vno = tblPurchSub.Vno and tblPurch.vtype = tblPurchSub.Vtype and tblPurch.vyear = tblPurchSub.vyear and tblPurch.vfirm = tblPurchSub.Vfirm and tblPurch.vtype = 'PY' and " _
& " tblPurch.PurBillDt < ('" & Format(pDtFr, "yyyy/mm/dd") & "')  and tblPurch.Vyear = " & gCYear & " and tblPurch.vfirm =  '" & gCCode & "'"
If GProcRstOpen(RsTmp_S, X, "R", gCn) > 0 Then
   If IsNull(RsTmp_S.Fields(0)) = False Then
      OpStk = OpStk + RsTmp_S.Fields(0)
   End If
End If
DoEvents
DoEvents
DoEvents

'--Yarn OP Stock Qty - Less Sales Before From Date

X = "select sum(SlSubBag) from tblSaleSub where tblSaleSub.vtype = 'SY' and " _
& " tblSaleSub.Vdt < ('" & Format(pDtFr, "yyyy/mm/dd") & "') and tblSaleSub.Vyear = " & gCYear & "  and tblSaleSub.vfirm =  '" & gCCode & "'"
If GProcRstOpen(RsTmp_S, X, "R", gCn) > 0 Then
   If IsNull(RsTmp_S.Fields(0)) = False Then
      OpStk = OpStk - RsTmp_S.Fields(0)
   End If
End If
DoEvents
DoEvents
DoEvents
'--Yarn Purchase Qty
X = "select sum(PurSubBag) from tblPurchSub,tblPurch where tblPurch.vno = tblPurchSub.Vno and tblPurch.vtype = tblPurchSub.Vtype and tblPurch.vyear = tblPurchSub.vyear and tblPurch.vfirm = tblPurchSub.Vfirm and tblPurch.vtype = 'PY' and " _
& " tblPurch.PurBillDt >= ('" & Format(pDtFr, "yyyy/mm/dd") & "')  and tblPurch.PurBillDt <= ('" & Format(pDtTo, "yyyy/mm/dd") & "')  and tblPurch.vfirm =  '" & gCCode & "'"
If GProcRstOpen(RsTmp_S, X, "R", gCn) > 0 Then
   If IsNull(RsTmp_S.Fields(0)) = False Then
      PurStk = RsTmp_S.Fields(0)
   Else
      PurStk = 0
   End If
End If
                 
DoEvents
DoEvents
DoEvents

'--Yarn Purchase Return Qty
X = "select sum(PurSubBag) from tblPurchSub,tblPurch where tblPurch.vno = tblPurchSub.Vno and tblPurch.vtype = tblPurchSub.Vtype and tblPurch.vyear = tblPurchSub.vyear and tblPurch.vfirm = tblPurchSub.Vfirm and tblPurch.vtype = 'VY' and " _
& " tblPurch.PurBillDt >= ('" & Format(pDtFr, "yyyy/mm/dd") & "')  and tblPurch.PurBillDt <= ('" & Format(pDtTo, "yyyy/mm/dd") & "')  and tblPurch.vfirm =  '" & gCCode & "'"
If GProcRstOpen(RsTmp_S, X, "R", gCn) > 0 Then
   If IsNull(RsTmp_S.Fields(0)) = False Then
      PurRtnStk = RsTmp_S.Fields(0)
   Else
      PurRtnStk = 0
   End If
End If
                 
DoEvents
DoEvents
DoEvents



'--Yarn Sale Qty

X = "select sum(SlSubBag) from tblSaleSub where tblSaleSub.vtype = 'SY' and " _
& " tblSaleSub.Vdt >= ('" & Format(pDtFr, "yyyy/mm/dd") & "')  and tblSaleSub.Vdt <= ('" & Format(pDtTo, "yyyy/mm/dd") & "')  and tblSaleSub.vfirm =  '" & gCCode & "'"
If GProcRstOpen(RsTmp_S, X, "R", gCn) > 0 Then
   If IsNull(RsTmp_S.Fields(0)) = False Then
      SaleStk = RsTmp_S.Fields(0)
   Else
      SaleStk = 0
   End If
End If
                 
                 
'--Yarn Sale Rtn Qty

X = "select sum(SlSubBag) from tblSaleSub where tblSaleSub.vtype = 'RY' and " _
& " tblSaleSub.Vdt >= ('" & Format(pDtFr, "yyyy/mm/dd") & "')  and tblSaleSub.Vdt <= ('" & Format(pDtTo, "yyyy/mm/dd") & "')  and tblSaleSub.vfirm =  '" & gCCode & "'"
If GProcRstOpen(RsTmp_S, X, "R", gCn) > 0 Then
   If IsNull(RsTmp_S.Fields(0)) = False Then
      SaleRtnStk = RsTmp_S.Fields(0)
   Else
      SaleRtnStk = 0
   End If
End If
                 
If gCIBOffice <> "I" And gCYear = 2020 Then '---For Fire Case
   If gCCode = "KAG" Then
      PurRtnStk = PurRtnStk + 589
   End If
   If gCCode = "VTC" Then
      PurRtnStk = PurRtnStk + 190
   End If

End If
                 
                 
                 
                 
DoEvents
DoEvents
DoEvents
                 
ClStk = OpStk + PurStk - SaleStk + SaleRtnStk - PurRtnStk
                 
'-- end Stock Calculation
                 
                 
                 gCmd.CommandText = "PrcPrepareClBalance"
                 gCmd.Parameters.Refresh
'                 gCmd.Parameters("@From_dt") = "'" & Format(FromDt, "yyyy/mm/dd") & "'"
'                 If ChkConsolidate.Value = 1 Then
'                    gCmd.Parameters("@VFirm") = ""
'                 Else
                    gCmd.Parameters("@VFirm") = Trim(gCCode)
'                 End If
                 gCmd.Parameters("@Fr_dt") = Format(pDtFr, "yyyy/mm/dd")
                 gCmd.Parameters("@To_dt") = Format(pDtTo, "yyyy/mm/dd")
                 gCmd.Parameters("@FyStart_dt") = Format(gCYSDate, "yyyy/mm/dd")
                 gCmd.Parameters("@UCode") = gUserName
                 gCmd.Parameters("@VFirm") = Trim(gCCode)
                 gCmd.Parameters("@CGSTAc") = gCgstAcCode
                 gCmd.Parameters("@SGSTAc") = gSgstAcCode
                 gCmd.Parameters("@IGSTAc") = gIgstAcCode
                 gCmd.Parameters("@TcsRec") = gTcsRec
                 gCmd.Parameters("@TcsPay") = gTcsPay
                 gCmd.Parameters("@CGSTRcmPayAc") = gCgstRCMAcCode
                 gCmd.Parameters("@SGSTRcmPayAc") = gSgstRCMAcCode
                 gCmd.Parameters("@IGSTRcmPayAc") = gIgstRCMAcCode
                 gCmd.Parameters("@CGSTRcmRecAc") = gCgstRCMRecCode
                 gCmd.Parameters("@SGSTRcmRecAc") = gSgstRCMRecCode
                 gCmd.Parameters("@IGSTRcmRecAc") = gIgstRCMRecCode
                 gCmd.Parameters("@RoundingAc") = gRoundOffAc
                 gCmd.Execute
'gProcPrepareTrialBalance pDtFr, pDtTo
If pDtFr <> gCYSDate Then
   gCn.Execute "update TmpClosingBalance set ClBal_Dr = 0,ClBal_Cr=0 where Vfirm = '" & Trim(gCCode) & "'  and UserName='" & gUserName & "'"
   gCn.Execute "update TmpClosingBalance set ClBal_Dr = Debit-Credit where Vfirm = '" & Trim(gCCode) & "' and UserName='" & gUserName & "' and Debit > Credit"
   gCn.Execute "update TmpClosingBalance set ClBal_Cr = Credit-Debit where Vfirm = '" & Trim(gCCode) & "' and UserName='" & gUserName & "' and Debit <= Credit"
End If
gCn.Execute "delete from TmpClosingBalance where Vfirm = '" & Trim(gCCode) & "' and UserName='" & gUserName & "' and ClBal_Cr =ClBal_Dr"


t1 = 0
t2 = 0
P1 = 0
P2 = 0
mGp = 0
gCn.Execute "delete from tmpGenTbl where Vfirm = '" & Trim(gCCode) & "' and UserName='" & gUserName & "'"
'--- Insert Top Debit Side In TmpGenTbl
dt1 = GProcGetColumnValue("tmpselection", "username", gUserName, "S", "dt1", "S", "Vfirm = '" & Trim(gCCode) & "'")
If dt1 = gCYSDate Then
X = "select h3,h4,h5,h6,h3type,h4type,h5type,h6type,ClBal_Dr  from TmpClosingBalance,qryaccountgrouping Where TmpClosingBalance.AcCode = qryaccountgrouping.AcCode " _
  & "and ((qryaccountgrouping.topgroup = 90003 and h2 = 'Trading & Manufacturing A/C') or (qryaccountgrouping.topgroup = 90001 and h3 ='Stock In Hand'))  and Vfirm = '" & Trim(gCCode) & "'   and UserName = '" & gUserName & "' and ClBal_Dr > 0  Order by h2,s3,h3,h4,h5,h6"
Else
X = "select h3,h4,h5,h6,h3type,h4type,h5type,h6type,ClBal_Dr,h2,s3  from TmpClosingBalance,qryaccountgrouping Where TmpClosingBalance.AcCode = qryaccountgrouping.AcCode " _
  & "and ((qryaccountgrouping.topgroup = 90003 and qryaccountgrouping.topgroup <> 90018 and h2 = 'Trading & Manufacturing A/C') or (qryaccountgrouping.topgroup = 90001 and h3 ='Stock In Hand')) and Vfirm = '" & Trim(gCCode) & "' and  UserName = '" & gUserName & "' and ClBal_Dr > 0  " _
  & "Union All Select Agname,' '+Acname,'','','G','A','','',stamount,'','' From tblClosingStock, tblMastAccount, tblMastGroup, tmpselection " _
  & "where tblClosingStock.stAcCode =tblMastAccount.AcCode and tblMastGroup.Agcode=tblMastAccount.AgCode and tmpselection.Dt1 =tblClosingStock.StDate+1 and tmpselection.UserName = '" & gUserName & "' and tblClosingStock.Vfirm = '" & gCCode & "'" _
  & " and tblClosingStock.StDate >= ('" & Format(gCYSDate, "yyyy/mm/dd") & "')  and tblClosingStock.StDate <= ('" & Format(gCYEDate, "yyyy/mm/dd") & "')  " _
  & " Order by h2,s3,h3,h4,h5,h6 "
End If
If GProcRstOpen(RsTmp_S, X, "R", gCn) > 0 Then
    With RsTmp_S
        For i = 0 To 4
          mGnm(i) = ""
        Next i
        .MoveFirst
        t1 = 0
        Do While Not .EOF
           t1 = t1 + 1
           If !H3Type = "G" Then
                gCn.Execute "insert into tmpGenTbl (vno,VSrNo,Nar1,username,VFirm) " _
                & " values (1," & t1 & ",'" & !h3 & "','" & gUserName & "','" & gCCode & "')"
                mGnm(1) = !h3
                Do While Not .EOF
                    If mGnm(1) <> !h3 Then Exit Do
                    t1 = t1 + 1
                    If !h4type = "G" Then
                         gCn.Execute "insert into tmpGenTbl (vno,VSrNo,Nar1,username,VFirm) " _
                         & " values (1," & t1 & ",'   " & !h4 & "','" & gUserName & "','" & gCCode & "')"
                         mGnm(2) = !h4
                         Do While Not .EOF
                            If mGnm(2) <> !h4 Or mGnm(1) <> !h3 Then Exit Do
                            t1 = t1 + 1
                            If !h5type = "G" Then
                                 gCn.Execute "insert into tmpGenTbl (vno,VSrNo,Nar1,username,VFirm) " _
                                 & " values (1," & t1 & ",'      " & !H5 & "','" & gUserName & "','" & gCCode & "')"
                                 mGnm(3) = !H5
                                 Do While Not .EOF
                                    If mGnm(3) <> !H5 Or mGnm(2) <> !h4 Or mGnm(1) <> !h3 Then Exit Do
                                    t1 = t1 + 1
                                    gCn.Execute "insert into tmpGenTbl (vno,VSrNo,Nar1,Amt2,username,VFirm) " _
                                    & " values (1," & t1 & ",'         " & !H6 & "'," & !ClBal_Dr & ",'" & gUserName & "','" & gCCode & "')"
                                    mGp = mGp - !ClBal_Dr
                                     .MoveNext
                                 Loop
                            Else
                                 gCn.Execute "insert into tmpGenTbl (vno,VSrNo,Nar1,Amt2,username,VFirm) " _
                                 & " values (1," & t1 & ",'      " & !H5 & "'," & !ClBal_Dr & ",'" & gUserName & "','" & gCCode & "')"
                                  mGp = mGp - !ClBal_Dr
                                 .MoveNext
                            End If
                         Loop
                    Else
                         If !h4 = "Yarn Purchase" Then
                            gCn.Execute "insert into tmpGenTbl (vno,VSrNo,Nar1,Amt2,Amt10,username,VFirm) " _
                            & " values (1," & t1 & ",'   " & IIf(!h3 = "Stock In Hand", "Opening ", "") + !h4 & "'," & !ClBal_Dr & "," & PurStk & ",'" & gUserName & "','" & gCCode & "')"
                         ElseIf Left(!h4, 12) = "Sales Return" Then
                            gCn.Execute "insert into tmpGenTbl (vno,VSrNo,Nar1,Amt2,Amt10,username,VFirm) " _
                            & " values (1," & t1 & ",'   " & IIf(!h3 = "Stock In Hand", "Opening ", "") + !h4 & "'," & !ClBal_Dr & "," & SaleRtnStk & ",'" & gUserName & "','" & gCCode & "')"
                         
                         ElseIf Trim(!h4) = "Yarn Stock" Then
                            gCn.Execute "insert into tmpGenTbl (vno,VSrNo,Nar1,Amt2,Amt10,username,VFirm) " _
                            & " values (1," & t1 & ",'   " & IIf(!h3 = "Stock In Hand", "Opening ", "") + !h4 & "'," & !ClBal_Dr & "," & OpStk & ",'" & gUserName & "','" & gCCode & "')"
                         Else
                            gCn.Execute "insert into tmpGenTbl (vno,VSrNo,Nar1,Amt2,username,VFirm) " _
                            & " values (1," & t1 & ",'   " & IIf(!h3 = "Stock In Hand", "Opening ", "") + !h4 & "'," & !ClBal_Dr & ",'" & gUserName & "','" & gCCode & "')"
                         End If
                         mGp = mGp - !ClBal_Dr
                         .MoveNext
                    End If
                Loop
           Else
                gCn.Execute "insert into tmpGenTbl (vno,VSrNo,Nar1,Amt2,username,Vfirm) " _
                & " values (1," & t1 & ",'" & !h3 & "'," & !ClBal_Dr & ",'" & gUserName & "','" & gCCode & "')"
                mGp = mGp - !ClBal_Dr
                .MoveNext
           End If
        Loop
    End With
End If
'--- Insert Top Credit Side In TmpGenTbl
X = "select h3,h4,h5,h6,h3type,h4type,h5type,h6type,ClBal_Cr,Schedul  from TmpClosingBalance,qryaccountgrouping Where TmpClosingBalance.AcCode = qryaccountgrouping.AcCode " _
  & "and qryaccountgrouping.topgroup = 90003  and Vfirm = '" & Trim(gCCode) & "'  and UserName = '" & gUserName & "' and h2 = 'Trading & Manufacturing A/C' and ClBal_Cr > 0  " _
  & "Union All Select Agname,'Closing '+Acname,'','','G','A','','',stamount,Schedule as Schedul From tblClosingStock, tblMastAccount, tblMastGroup, tmpselection " _
  & "where tblClosingStock.StAmount > 0 and tblClosingStock.stAcCode =tblMastAccount.AcCode and tblMastGroup.Agcode=tblMastAccount.AgCode and tmpselection.Dt2 =tblClosingStock.StDate and tmpselection.UserName = '" & gUserName & "' and tmpselection.VFirm  =tblClosingStock.VFirm    and tblClosingStock.vFirm = '" & gCCode & "'" _
  & " and tblClosingStock.StDate >= ('" & Format(gCYSDate, "yyyy/mm/dd") & "')  and tblClosingStock.StDate <= ('" & Format(gCYEDate, "yyyy/mm/dd") & "') " _
  & "Order by Schedul,h3,h4,h5,h6"
If GProcRstOpen(RsTmp_S, X, "R", gCn) > 0 Then
    With RsTmp_S
        For i = 0 To 4
          mGnm(i) = ""
        Next i
        .MoveFirst
        t2 = 0
        Do While Not .EOF
           t2 = t2 + 1
           If !H3Type = "G" Then
                If (t2 > t1) Then
                    gCn.Execute "insert into tmpGenTbl (vno,VSrNo,Nar2,username,VFirm) values (1," & t2 & ",'" & !h3 & "','" & gUserName & "','" & gCCode & "')"
                Else
                    gCn.Execute "Update tmpGenTbl set Nar2 ='" & !h3 & "' where vno = 1 and Vsrno = " & t2
                End If
                mGnm(1) = !h3
                Do While Not .EOF
                    If mGnm(1) <> !h3 Then Exit Do
                    t2 = t2 + 1
                    If !h4type = "G" Then
                        If (t2 > t1) Then
                            gCn.Execute "insert into tmpGenTbl (vno,VSrNo,Nar2,username,VFirm) values (1," & t2 & ",'   " & !h4 & "','" & gUserName & "','" & gCCode & "')"
                        Else
                            gCn.Execute "Update tmpGenTbl set Nar2 ='   " & !h4 & "' where vno = 1 and Vsrno = " & t2
                        End If
                         mGnm(2) = !h4
                         Do While Not .EOF
                            If mGnm(2) <> !h4 Or mGnm(1) <> !h3 Then Exit Do
                            t2 = t2 + 1
                            If !h5type = "G" Then
                                If (t2 > t1) Then
                                    gCn.Execute "insert into tmpGenTbl (vno,VSrNo,Nar2,username,VFirm) values (1," & t2 & ",'      " & !H5 & "','" & gUserName & "','" & gCCode & "')"
                                Else
                                    gCn.Execute "Update tmpGenTbl set Nar2 ='      " & !H5 & "' where vno = 1 and Vsrno = " & t2
                                End If
                                 mGnm(3) = !H5
                                 Do While Not .EOF
                                    If mGnm(3) <> !H5 Or mGnm(2) <> !h4 Or mGnm(1) <> !h3 Then Exit Do
                                    t2 = t2 + 1
                                    If (t2 > t1) Then
                                        gCn.Execute "insert into tmpGenTbl (vno,VSrNo,Nar2,Amt4,username,VFirm) values (1," & t2 & ",'         " & !H6 & "'," & !ClBal_Cr & ",'" & gUserName & "','" & gCCode & "')"
                                    Else
                                        gCn.Execute "Update tmpGenTbl set Nar2 ='         " & !H6 & "',Amt4 = " & !ClBal_Cr & " where vno = 1 and Vsrno = " & t2
                                    End If
                                    mGp = mGp + !ClBal_Cr
                                     .MoveNext
                                 Loop
                            Else
                                If (t2 > t1) Then
                                    gCn.Execute "insert into tmpGenTbl (vno,VSrNo,Nar2,Amt4,username,VFirm) values (1," & t2 & ",'      " & !H5 & "'," & !ClBal_Cr & ",'" & gUserName & "','" & gCCode & "')"
                                Else
                                    gCn.Execute "Update tmpGenTbl set Nar2 ='      " & !H5 & "',Amt4 = " & !ClBal_Cr & " where vno = 1 and Vsrno = " & t2
                                End If
                                  mGp = mGp + !ClBal_Cr
                                 .MoveNext
                            End If
                         Loop
                    Else
                        If (t2 > t1) Then
                            If !h4 = "Yarn Sales" Then
                               gCn.Execute "insert into tmpGenTbl (vno,VSrNo,Nar2,Amt4,Amt11,username,VFirm) values (1," & t2 & ",'   " & !h4 & "'," & !ClBal_Cr & "," & SaleStk & " ,'" & gUserName & "','" & gCCode & "')"
                            ElseIf Left(!h4, 15) = "Purchase Return" Then
                               gCn.Execute "insert into tmpGenTbl (vno,VSrNo,Nar2,Amt4,Amt11,username,VFirm) values (1," & t2 & ",'   " & !h4 & "'," & !ClBal_Cr & "," & PurRtnStk & " ,'" & gUserName & "','" & gCCode & "')"
                            ElseIf Left(!h4, 7) = "Closing" Then
                               gCn.Execute "insert into tmpGenTbl (vno,VSrNo,Nar2,Amt4,Amt11,username,VFirm) values (1," & t2 & ",'   " & !h4 & "'," & !ClBal_Cr & "," & ClStk & " ,'" & gUserName & "','" & gCCode & "')"
                            Else
                               gCn.Execute "insert into tmpGenTbl (vno,VSrNo,Nar2,Amt4,username,VFirm) values (1," & t2 & ",'   " & !h4 & "'," & !ClBal_Cr & ",'" & gUserName & "','" & gCCode & "')"
                            End If
                        Else
                            If !h4 = "Yarn Sales" Then
                               gCn.Execute "Update tmpGenTbl set Nar2 ='   " & !h4 & "',Amt4 = " & !ClBal_Cr & ",Amt11 = " & SaleStk & " where vno = 1 and Vsrno = " & t2
                            ElseIf Left(!h4, 15) = "Purchase Return" Then
                               gCn.Execute "Update tmpGenTbl set Nar2 ='   " & !h4 & "',Amt4 = " & !ClBal_Cr & ",Amt11 = " & PurRtnStk & " where vno = 1 and Vsrno = " & t2
                            
                            
                            ElseIf Left(!h4, 7) = "Closing" Then
                               gCn.Execute "Update tmpGenTbl set Nar2 ='   " & !h4 & "',Amt4 = " & !ClBal_Cr & ",Amt11 = " & ClStk & " where vno = 1 and Vsrno = " & t2
                            Else
                               gCn.Execute "Update tmpGenTbl set Nar2 ='   " & !h4 & "',Amt4 = " & !ClBal_Cr & " where vno = 1 and Vsrno = " & t2
                            End If
                        End If
                         mGp = mGp + !ClBal_Cr
                         .MoveNext
                    End If
                Loop
           Else
                If (t2 > t1) Then
                    gCn.Execute "insert into tmpGenTbl (vno,VSrNo,Nar2,Amt4,username,VFirm) values (1," & t2 & ",'" & !h3 & "'," & !ClBal_Cr & ",'" & gUserName & "','" & gCCode & "')"
                Else
                    gCn.Execute "Update tmpGenTbl set Nar2 ='" & !h3 & "',Amt4 = " & !ClBal_Cr & " where vno = 1 and Vsrno = " & t2
                End If
                mGp = mGp + !ClBal_Cr
                .MoveNext
           End If
        Loop
    End With
End If
If mGp > 0 Then
    t1 = t1 + 1
    If (t1 > t2) Then
        gCn.Execute "insert into tmpGenTbl (vno,VSrNo,Nar1,Amt2,username,VFirm) values (1," & t1 & ",'Gross Profit'," & Abs(mGp) & ",'" & gUserName & "','" & gCCode & "')"
    Else
        gCn.Execute "Update tmpGenTbl set Nar1 ='Gross Profit',Amt2 = " & Abs(mGp) & " where vno = 1 and Vsrno = " & t1
    End If
    P2 = P2 + 1
    gCn.Execute "insert into tmpGenTbl (vno,VSrNo,Nar2,Amt4,username,VFirm) values (2," & P2 & ",'Gross Profit'," & Abs(mGp) & ",'" & gUserName & "','" & gCCode & "')"
Else
    t2 = t2 + 1
    If (t2 > t1) Then
        gCn.Execute "insert into tmpGenTbl (vno,VSrNo,Nar2,Amt4,username,VFirm) values (1," & t2 & ",'Gross Loss'," & Abs(mGp) & ",'" & gUserName & "','" & gCCode & "')"
    Else
        gCn.Execute "Update tmpGenTbl set Nar2 ='Gross Loss',Amt4 = " & Abs(mGp) & " where vno = 1 and Vsrno = " & t2
    End If
    P1 = P1 + 1
    gCn.Execute "insert into tmpGenTbl (vno,VSrNo,Nar1,Amt2,username,VFirm) values (2," & P1 & ",'Gross Loss'," & Abs(mGp) & ",'" & gUserName & "','" & gCCode & "')"
End If

'--- Insert Bottom Debit Side In TmpGenTbl
X = "select h3,h4,h5,h6,h3type,h4type,h5type,h6type,ClBal_Dr  from TmpClosingBalance,qryaccountgrouping Where TmpClosingBalance.AcCode = qryaccountgrouping.AcCode " _
  & "and qryaccountgrouping.topgroup = 90003 and Vfirm = '" & Trim(gCCode) & "'  and username = '" & gUserName & "' and h2 = 'Profit & Loss A/C' and ClBal_Dr > 0  Order by h2,h3,h4,h5,h6"
If GProcRstOpen(RsTmp_S, X, "R", gCn) > 0 Then
    With RsTmp_S
        For i = 0 To 4
          mGnm(i) = ""
        Next i
        .MoveFirst
        Do While Not .EOF
           P1 = P1 + 1
           If !H3Type = "G" Then
                If (P1 > P2) Then
                  gCn.Execute "insert into tmpGenTbl (vno,VSrNo,Nar1,username,VFirm) " _
                  & " values (2," & P1 & ",'" & !h3 & "','" & gUserName & "','" & gCCode & "')"
                Else
                  gCn.Execute "Update tmpGenTbl set Nar1='" & !h3 & "' where vno = 2 and vsrno = " & P1
                End If
                mGnm(1) = !h3
                Do While Not .EOF
                    If mGnm(1) <> !h3 Then Exit Do
                    P1 = P1 + 1
                    If !h4type = "G" Then
                         gCn.Execute "insert into tmpGenTbl (vno,VSrNo,Nar1,username,VFirm) " _
                         & " values (2," & P1 & ",'   " & !h4 & "','" & gUserName & "','" & gCCode & "')"
                         mGnm(2) = !h4
                         Do While Not .EOF
                            If mGnm(2) <> !h4 Or mGnm(1) <> !h3 Then Exit Do
                            P1 = P1 + 1
                            If !h5type = "G" Then
                                 gCn.Execute "insert into tmpGenTbl (vno,VSrNo,Nar1,username,VFirm) " _
                                 & " values (2," & P1 & ",'      " & !H5 & "','" & gUserName & "','" & gCCode & "')"
                                 mGnm(3) = !H5
                                 Do While Not .EOF
                                    If mGnm(3) <> !H5 Or mGnm(2) <> !h4 Or mGnm(1) <> !h3 Then Exit Do
                                    P1 = P1 + 1
                                    gCn.Execute "insert into tmpGenTbl (vno,VSrNo,Nar1,Amt2,username,VFirm) " _
                                    & " values (2," & P1 & ",'         " & !H6 & "'," & !ClBal_Dr & ",'" & gUserName & "','" & gCCode & "')"
                                    mGp = mGp - !ClBal_Dr
                                     .MoveNext
                                 Loop
                            Else
                                 gCn.Execute "insert into tmpGenTbl (vno,VSrNo,Nar1,Amt2,username,VFirm) " _
                                 & " values (2," & P1 & ",'      " & !H5 & "'," & !ClBal_Dr & ",'" & gUserName & "','" & gCCode & "')"
                                  mGp = mGp - !ClBal_Dr
                                 .MoveNext
                            End If
                         Loop
                    Else
                         gCn.Execute "insert into tmpGenTbl (vno,VSrNo,Nar1,Amt2,username,VFirm) " _
                         & " values (2," & P1 & ",'   " & !h4 & "'," & !ClBal_Dr & ",'" & gUserName & "','" & gCCode & "')"
                         mGp = mGp - !ClBal_Dr
                         .MoveNext
                    End If
                Loop
           Else
                gCn.Execute "insert into tmpGenTbl (vno,VSrNo,Nar1,Amt2,username,VFirm) " _
                & " values (2," & P1 & ",'" & !h3 & "'," & !ClBal_Dr & ",'" & gUserName & "','" & gCCode & "')"
                mGp = mGp - !ClBal_Dr
                .MoveNext
           End If
        Loop
    End With
End If
'--- Insert Bottom Credit Side In TmpGenTbl
X = "select h3,h4,h5,h6,h3type,h4type,h5type,h6type,ClBal_Cr  from TmpClosingBalance,qryaccountgrouping Where TmpClosingBalance.AcCode = qryaccountgrouping.AcCode " _
  & "and qryaccountgrouping.topgroup = 90003 and Vfirm = '" & Trim(gCCode) & "'  and username = '" & gUserName & "' and h2 = 'Profit & Loss A/C' and ClBal_Cr > 0  Order by h2,h3,h4,h5,h6"
If GProcRstOpen(RsTmp_S, X, "R", gCn) > 0 Then
    With RsTmp_S
        For i = 0 To 4
          mGnm(i) = ""
        Next i
        .MoveFirst
        Do While Not .EOF
           P2 = P2 + 1
           If !H3Type = "G" Then
                If (P2 > P1) Then
                    gCn.Execute "insert into tmpGenTbl (vno,VSrNo,Nar2,username,VFirm) values (2," & P2 & ",'" & !h3 & "','" & gUserName & "','" & gCCode & "')"
                Else
                    gCn.Execute "Update tmpGenTbl set Nar2 ='" & !h3 & "' where vno = 2 and vsrno = " & P2
                End If
                mGnm(1) = !h3
                Do While Not .EOF
                    If mGnm(1) <> !h3 Then Exit Do
                    P2 = P2 + 1
                    If !h4type = "G" Then
                        If (P2 > P1) Then
                            gCn.Execute "insert into tmpGenTbl (vno,VSrNo,Nar2,username,VFirm) values (2," & P2 & ",'   " & !h4 & "','" & gUserName & "','" & gCCode & "')"
                        Else
                            gCn.Execute "Update tmpGenTbl set Nar2 ='   " & !h4 & "' where vno = 2 and vsrno = " & P2
                        End If
                         mGnm(2) = !h4
                         Do While Not .EOF
                            If mGnm(2) <> !h4 Or mGnm(1) <> !h3 Then Exit Do
                            P2 = P2 + 1
                            If !h5type = "G" Then
                                If (P2 > P1) Then
                                    gCn.Execute "insert into tmpGenTbl (vno,VSrNo,Nar2,username,VFirm) values (2," & P2 & ",'      " & !H5 & "','" & gUserName & "','" & gCCode & "')"
                                Else
                                    gCn.Execute "Update tmpGenTbl set Nar2 ='      " & !H5 & "' where vno = 2 and vsrno = " & P2
                                End If
                                 mGnm(3) = !H5
                                 Do While Not .EOF
                                    If mGnm(3) <> !H5 Or mGnm(2) <> !h4 Or mGnm(1) <> !h3 Then Exit Do
                                    P2 = P2 + 1
                                    If (P2 > P1) Then
                                        gCn.Execute "insert into tmpGenTbl (vno,VSrNo,Nar2,Amt4,username,VFirm) values (2," & P2 & ",'         " & !H6 & "'," & !ClBal_Cr & ",'" & gUserName & "','" & gCCode & "')"
                                    Else
                                        gCn.Execute "Update tmpGenTbl set Nar2 ='         " & !H6 & "',Amt4 = " & !ClBal_Cr & " where vno = 2 and vsrno = " & P2
                                    End If
                                    mGp = mGp + !ClBal_Cr
                                     .MoveNext
                                 Loop
                            Else
                                If (P2 > P1) Then
                                    gCn.Execute "insert into tmpGenTbl (vno,VSrNo,Nar2,Amt4,username,VFirm) values (2," & P2 & ",'      " & !H5 & "'," & !ClBal_Cr & ",'" & gUserName & "','" & gCCode & "')"
                                Else
                                    gCn.Execute "Update tmpGenTbl set Nar2 ='      " & !H5 & "',Amt4 = " & !ClBal_Cr & " where vno = 2 and vsrno = " & P2
                                End If
                                  mGp = mGp + !ClBal_Cr
                                 .MoveNext
                            End If
                         Loop
                    Else
                        If (P2 > P1) Then
                            gCn.Execute "insert into tmpGenTbl (vno,VSrNo,Nar2,Amt4,username,VFirm) values (2," & P2 & ",'   " & !h4 & "'," & !ClBal_Cr & ",'" & gUserName & "','" & gCCode & "')"
                        Else
                            gCn.Execute "Update tmpGenTbl set Nar2 ='   " & !h4 & "',Amt4 = " & !ClBal_Cr & " where vno = 2 and vsrno = " & P2
                        End If
                         mGp = mGp + !ClBal_Cr
                         .MoveNext
                    End If
                Loop
           Else
                If (P2 > P1) Then
                    gCn.Execute "insert into tmpGenTbl (vno,VSrNo,Nar2,Amt4,username,VFirm) values (2," & P2 & ",'" & !h3 & "'," & !ClBal_Cr & ",'" & gUserName & "','" & gCCode & "')"
                Else
                    gCn.Execute "Update tmpGenTbl set Nar2 ='" & !h3 & "',Amt4 = " & !ClBal_Cr & " where vno = 2 and vsrno = " & P2
                End If
                mGp = mGp + !ClBal_Cr
                .MoveNext
           End If
        Loop
    End With
End If
If mGp > 0 Then
    P1 = P1 + 1
    If (P1 > P2) Then
        gCn.Execute "insert into tmpGenTbl (vno,VSrNo,Nar1,Amt2,username,VFirm) values (2," & P1 & ",'Net Profit'," & Abs(mGp) & ",'" & gUserName & "','" & gCCode & "')"
    Else
        gCn.Execute "Update tmpGenTbl set Nar1 ='Net Profit',Amt2 = " & Abs(mGp) & " where vno = 2 and vsrno = " & P1
    End If
Else
    P2 = P2 + 1
    If (P2 > P1) Then
        gCn.Execute "insert into tmpGenTbl (vno,VSrNo,Nar2,Amt4,username,VFirm) values (2," & P2 & ",'Net Loss'," & Abs(mGp) & ",'" & gUserName & "','" & gCCode & "')"
    Else
        gCn.Execute "Update tmpGenTbl set Nar2 ='Net Loss',Amt4 = " & Abs(mGp) & " where vno = 2 and vsrno = " & P2
    End If
End If
GProcRstClose RsTmp_S
'RsTmp_T.Close
End Sub
'---- Balance Sheet
Public Sub GProcPrepareBalanceSheet(FromDt As Date, ToDate As Date)
Dim RsTmp_T As New ADODB.Recordset
Dim RsTmp_S As New ADODB.Recordset
Dim RsTmp_K As New ADODB.Recordset
Dim RCount_T As Long, RCount_S As Long, i As Long
Dim X As String, BalQuery As String
Dim StkCd As Long
StkCd = GProcGetColumnValue("TblMastAccount", "AcName", "Yarn Stock", "S", "AcCode", "N")

                 gCmd.CommandText = "PrcPrepareClBalance"
                 gCmd.Parameters.Refresh
'                 gCmd.Parameters("@From_dt") = "'" & Format(FromDt, "yyyy/mm/dd") & "'"
'                 If ChkConsolidate.Value = 1 Then
'                    gCmd.Parameters("@VFirm") = ""
'                 Else
                    gCmd.Parameters("@VFirm") = Trim(gCCode)
'                 End If
                 gCmd.Parameters("@Fr_dt") = Format(FromDt, "yyyy/mm/dd")
                 gCmd.Parameters("@To_dt") = Format(ToDate, "yyyy/mm/dd")
                 gCmd.Parameters("@FyStart_dt") = Format(gCYSDate, "yyyy/mm/dd")
                 gCmd.Parameters("@UCode") = gUserName
                 gCmd.Parameters("@VFirm") = Trim(gCCode)
                 gCmd.Parameters("@CGSTAc") = gCgstAcCode
                 gCmd.Parameters("@SGSTAc") = gSgstAcCode
                 gCmd.Parameters("@IGSTAc") = gIgstAcCode
                 gCmd.Parameters("@TcsRec") = gTcsRec
                 gCmd.Parameters("@TcsPay") = gTcsPay
                 gCmd.Parameters("@CGSTRcmPayAc") = gCgstRCMAcCode
                 gCmd.Parameters("@SGSTRcmPayAc") = gSgstRCMAcCode
                 gCmd.Parameters("@IGSTRcmPayAc") = gIgstRCMAcCode
                 gCmd.Parameters("@CGSTRcmRecAc") = gCgstRCMRecCode
                 gCmd.Parameters("@SGSTRcmRecAc") = gSgstRCMRecCode
                 gCmd.Parameters("@IGSTRcmRecAc") = gIgstRCMRecCode
                 gCmd.Parameters("@RoundingAc") = gRoundOffAc
                 gCmd.Execute

gCn.Execute "delete from TmpClosingBalance where ClBal_Dr + ClBal_Cr = 0 and UserName='" & gUserName & "' and Vfirm = '" & gCCode & "'  "


'--- For Opening Stock = 0
X = "select *  from TmpClosingBalance  where UserName='" & gUserName & "' and Vfirm = '" & gCCode & "' and AcName = 'Yarn Stock'"
If GProcRstOpen(RsTmp_K, X, "O", gCn) = 0 Then
   gCn.BeginTrans
   gCn.Execute "insert into TmpClosingBalance (Vsrno,Vfirm,VFirmName,AcCode,AcName,AgCode,AgName,UserName,AcOurCode,TopGrp) " _
               & " Values (2,'" & gCCode & "','" & gCName & "'," & StkCd & ",'Yarn Stock',90018,'Stock In Hand','" & gUserName & "','AINI',90001)"

   gCn.CommitTrans

End If

gCn.Execute "delete from tmpGenTbl where UserName='" & gUserName & "' and Vfirm = '" & gCCode & "'  "
'=== Insert For Liability Side
GetBalRecordInTmpTbl 90002, ToDate
BalQuery = "Select Vtype,Nar1,Amt1,UserName,VFirm,Amt2,amt3 from tmpRptTbl " _
& " where UserName='" & gUserName & "' and VFirm = '" & gCCode & "'  order by vno,Nar2,Nar9,Nar3,Nar10,Nar4,Nar11,Nar5,Nar12,Nar6,Nar7"
If GProcRstOpen(RsTmp_S, BalQuery, "O", gCn) > 0 Then
    With RsTmp_S
        .MoveFirst
        Do While Not .EOF
           gCn.Execute "insert into tmpGenTbl (vno,Vtype,Nar13,Amt1,username,VFirm,Amt11,amt3) " _
           & " values (" & .AbsolutePosition & ",'" & .Fields(0) & "','" & Left(.Fields(1), 50) & "'," & .Fields(2) & ",'" & .Fields(3) & "','" & .Fields(4) & "'," & .Fields(5) & "," & .Fields(6) & ")"
           .MoveNext
        Loop
    End With
End If
'=== UpDate For Asset Side

'--- For Closing Stock

X = "select StAcCode,StAmount   from tblMastAccount,tblClosingStock  where tblMastAccount.AcCode = tblClosingStock.StAcCode " _
    & " and tblMastAccount.AcName = 'Yarn Stock' and tblClosingStock.stDate = ('" & Format(ToDate, "yyyy/mm/dd") & "')  and tblClosingStock.Vfirm = '" & gCCode & "'"

If GProcRstOpen(RsTmp_S, X, "O", gCn) > 0 Then
   If RsTmp_S.Fields(1) = 0 Then
      gCn.Execute "delete TmpClosingBalance  where UserName='" & gUserName & "' and Vfirm = '" & gCCode & "' and AcCode = " & RsTmp_S.Fields(0) & " "
   Else
      gCn.Execute "update TmpClosingBalance set ClBal_Dr = " & RsTmp_S.Fields(1) & " where UserName='" & gUserName & "' and Vfirm = '" & gCCode & "' and AcCode = " & RsTmp_S.Fields(0) & " "
   End If
Else
   gCn.Execute "delete TmpClosingBalance  where UserName='" & gUserName & "' and Vfirm = '" & gCCode & "' and AcCode = " & RsTmp_S.Fields(0) & " "
End If
'-- end Closing Stock



RCount_T = GProcRstOpen(RsTmp_T, "Select * from TmpGenTbl where UserName='" & gUserName & "' and VFirm = '" & gCCode & "' and Vtype='B' order by vno ", "R", gCn)
GetBalRecordInTmpTbl 90001, ToDate



BalQuery = "Select Vtype,Nar1,Amt1,UserName,VFirm,Amt2,Amt3,Vno from tmpRptTbl " _
& " where UserName='" & gUserName & "' and VFirm = '" & gCCode & "' order by vno,Nar2,Nar9,Nar3,Nar10,Nar4,Nar11,Nar5,Nar12,Nar6,Nar7"
RCount_S = GProcRstOpen(RsTmp_S, BalQuery, "O", gCn)
If RCount_S > 0 Then
    With RsTmp_S
        .MoveFirst
        Do While Not .EOF
           i = i + 1
           If i > RCount_T Then '- Right Side > Left Side 'Nar1,Nar2,Nar3,Nar4,Nar5,Nar6,Nar8,Nar9,Nar10,Nar11,Nar12,
                gCn.Execute "insert into tmpGenTbl (vno,Vtype,Nar7,Amt2,UserName,VFirm,Amt12,Amt11,Amt4) values (" & i & ",'" & .Fields(0) & "'," _
                & "'" & .Fields(1) & "',-1*" & .Fields(2) & ",'" & .Fields(3) & "','" & .Fields(4) & "'," & .Fields(5) & "," & .Fields(7) & "," & .Fields(6) & "*-1)"
           Else
                gCn.Execute "update tmpGenTbl set Vtype='" & .Fields("Vtype") & "',Nar7='" & .Fields("Nar1") & "',Amt2=-1*" & .Fields("amt1") & ",Amt4=-1*" & .Fields("amt3") & ",Amt12=" & .Fields("amt2") & ",Amt11=" & .Fields("Vno") & "  where " _
                & " username='" & .Fields("UserName") & "' and VFirm = '" & .Fields("VFirm") & "'   and vno=" & i
           
           End If
            If Not RsTmp_T.EOF Then RsTmp_T.MoveNext
            .MoveNext
        Loop
    End With
End If
GProcRstClose RsTmp_S
GProcRstClose RsTmp_T
Exit Sub
ErrorHandler:
   GProcErrorHandler
End Sub
 
 
 
 
 
 
 
 
'--- Account Balance On Date
Public Function GProcGetAccountBalance(pAname As String, pOnDate As Date, Optional IsOnlyOPBalance As Boolean) As Double
Dim rsBal As Recordset
Dim rsOP As Recordset
Dim X As String
Dim AcCode As Long
'AcCode = GProcGetColumnValue("TblMastAccount", "AcName", pAname, "S", "AcCode", "N")
'If IsOnlyOPBalance = False Then
'   If gBackEndDB = gBackEndAccess Then
'       X = "Select sum(VAmt) from tblvoucher where VYear=" & gCYear & " and vtype<>'0P' and  VAcCode=" & AcCode & " and Vdt<=cdate('" & pOnDate & "')"
'   Else '-- Oracle
'        X = "Select sum(VAmt) from tblvoucher where VYear=" & gCYear & " and vtype<>'0P' and VAcCode=" & AcCode & " and Vdt<=to_date('" & pOnDate & "','dd/MM/yyyy')"
'   End If
'   GProcRstOpen rsBal, X, "R"
'    If IsNumeric(rsBal.Fields(0)) = True Then
'        GProcGetAccountBalance = CDbl(rsBal.Fields(0))
'    End If
'    X = "Select VAmt from tblvoucher where VYear=" & gCYear & " and VAcCode=" & AcCode & " and vtype='0P'"
'    GProcRstOpen rsOP, X, "R"
'    If IsNumeric(rsOP.Fields(0)) = True Then
'        GProcGetAccountBalance = GProcGetAccountBalance + CDbl(rsOP.Fields(0))
'    End If
'Else
'    X = "Select VAmt from tblvoucher where VYear=" & gCYear & " and VAcCode=" & AcCode & " and vtype='0P'"
'    GProcRstOpen rsBal, X, "R"
'    If IsNumeric(rsBal.Fields(0)) = True Then
'        GProcGetAccountBalance = CDbl(rsBal.Fields(0))
'    End If
'End If
End Function
'------------------------- Prepare VAT Computation Details
Public Sub GProcPrepareVATComputDet(FromDt As Date, ToDt As Date, CrystalRpt As CrystalReport)
Dim CondStr As String
Dim x1 As String
gCn.Execute "delete from tmpGenTbl where Username='" & gUserName & "'"
If gBackEndDB = gBackEndAccess Then
    CondStr = " tbladdless.Vdt>= cdate('" & FromDt & "')  and  tbladdless.Vdt<=cdate('" & ToDt & "')"
Else '-- Oracle
    CondStr = " tbladdless.Vdt>= to_date('" & CDate(FromDt) & "','dd/MM/yyyy')  and tbladdless.Vdt<=to_date('" & CDate(ToDt) & "','dd/MM/yyyy')"
End If
    '===================================== Sales
    '-------- Cosign/Depot
    If gBackEndDB = gBackEndAccess Then  '-- Access
        gCn.Execute "insert into tmpGenTbl (Vno,Nar1,Nar2,Amt1,Amt2,Amt3,Amt4,Amt5,Amt6,Amt7,UserName,VSRNO) " _
        & " select 1,'Consignment /Depot' as SAType,AcName as Account,sum(VSubBag) as bags,sum(iif(Narration='V A T',AdBillamt-AdTaxAmt,0)) as NetVATSA," _
        & "sum(iif(Narration='V A T',ADEXEMPTAMT,0)) as ExptAmt,sum(iif(Narration='V A T',ADTaxAMT,0)) as VAT," _
        & "sum(iif(Narration='C S T',AdBillamt-AdTaxAmt,0)) as NetCSTSA," _
        & "sum(iif(Narration='C S T',ADTaxAMT,0)) as CST,0 as CForm,'" & gUserName & "',1" _
        & " from tblAddLess,TblMastAccount,tblMastNarration,tblVouSub where NarrCode= ADTAXCODE and AdMillCode=AcCode and " & CondStr & " and tblAddLess.vtype in ('SO','SD') " _
        & " and tblAddLess.Vtype=tblVouSub.Vtype and tblAddLess.VNo=tblVouSub.VNo and tblAddLess.VYear=tblVouSub.VYear" _
        & " group by acname order by acname"
    Else
        gCn.Execute "insert into tmpGenTbl (Vno,Nar1,Nar2,Amt1,Amt2,Amt3,Amt4,Amt5,Amt6,Amt7,UserName,VSRNO) " _
        & " select 1,'Consignment /Depot' as SAType,AcName as Account,sum(VSubBag) as bags,sum(decode(Narration,'V A T',AdBillamt-AdTaxAmt,0)) as NetVATSA," _
        & "sum(decode(Narration,'V A T',ADEXEMPTAMT,0)) as ExptAmt,sum(decode(Narration,'V A T',ADTaxAMT,0)) as VAT," _
        & "sum(decode(Narration,'C S T',AdBillamt-AdTaxAmt,0)) as NetCSTSA," _
        & "sum(decode(Narration,'C S T',ADTaxAMT,0)) as CST,0 as CForm,'" & gUserName & "',1" _
        & " from tblAddLess,TblMastAccount,tblMastNarration,tblVouSub where NarrCode= ADTAXCODE and AdMillCode=AcCode and " & CondStr & " and tblAddLess.vtype in ('SO','SD') " _
        & " and tblAddLess.Vtype=tblVouSub.Vtype and tblAddLess.VNo=tblVouSub.VNo and tblAddLess.VYear=tblVouSub.VYear" _
        & " group by acname order by acname"
    End If
    '--- Sales A/c (SY Without Exmpt)
    If gBackEndDB = gBackEndAccess Then  '-- Access
        gCn.Execute "insert into tmpGenTbl (Vno,Nar1,Nar2,Amt1,Amt2,Amt3,Amt4,Amt5,Amt6,Amt7,UserName) " _
        & " select 2,'Sales Account','Sales Account',sum(VSubBag) as bags,sum(iif(Narration='V A T',AdBillamt-AdTaxAmt,0)) as NetVATSA," _
        & "sum(iif(Narration='V A T',ADEXEMPTAMT,0)) as ExptAmt,sum(iif(Narration='V A T',ADTaxAMT,0)) as VAT," _
        & "sum(iif(Narration='C S T',AdBillamt-AdTaxAmt,0)) as NetCSTSA," _
        & "sum(iif(Narration='C S T',ADTaxAMT,0)) as CST,0 as CForm,'" & gUserName & "'" _
        & " from tblAddLess,tblMastNarration,tblVouSub where " & CondStr & " and  NarrCode= ADTAXCODE " _
        & " and tblAddLess.Vtype=tblVouSub.Vtype and tblAddLess.VNo=tblVouSub.VNo and tblAddLess.VYear=tblVouSub.VYear" _
        & " and  tblAddLess.vtype in ('SY') and ADEXEMPTAMT=0"
    Else
        gCn.Execute "insert into tmpGenTbl (Vno,Nar1,Nar2,Amt1,Amt2,Amt3,Amt4,Amt5,Amt6,Amt7,UserName) " _
        & " select 2,'Sales Account','Sales Account',sum(VSubBag) as bags,sum(decode(Narration,'V A T',AdBillamt-AdTaxAmt,0)) as NetVATSA," _
        & "sum(decode(Narration,'V A T',ADEXEMPTAMT,0)) as ExptAmt,sum(decode(Narration,'V A T',ADTaxAMT,0)) as VAT," _
        & "sum(decode(Narration,'C S T',AdBillamt-AdTaxAmt,0)) as NetCSTSA," _
        & "sum(decode(Narration,'C S T',ADTaxAMT,0)) as CST,0 as CForm,'" & gUserName & "'" _
        & " from tblAddLess,tblMastNarration,tblVouSub where " & CondStr & " and  NarrCode= ADTAXCODE " _
        & " and tblAddLess.Vtype=tblVouSub.Vtype and tblAddLess.VNo=tblVouSub.VNo and tblAddLess.VYear=tblVouSub.VYear" _
        & " and  tblAddLess.vtype in ('SY') and ADEXEMPTAMT=0"
     End If
    '--- SIT Sales
    If gBackEndDB = gBackEndAccess Then  '-- Access
        gCn.Execute " insert into tmpGenTbl (Vno,Nar1,Nar2,Amt1,Amt2,Amt3,Amt4,Amt5,Amt6,Amt7,UserName) " _
        & " select 3,'Sale In Transit Sales','Sale In Transit Sales',sum(VSubBag) as bags,sum(iif(Narration='V A T',AdBillamt-AdTaxAmt,0)) as NetVATSA," _
        & "sum(iif(Narration='V A T',ADEXEMPTAMT,0)) as ExptAmt,sum(iif(Narration='V A T',ADTaxAMT,0)) as VAT," _
        & "sum(iif(Narration='C S T',AdBillamt-AdTaxAmt,0)) as NetCSTSA," _
        & "sum(iif(Narration='C S T',ADTaxAMT,0)) as CST,sum(ADBillAMT) as CForm,'" & gUserName & "'" _
        & " from tblAddLess,tblMastNarration,tblVouSub where " & CondStr & " and NarrCode= ADTAXCODE " _
        & " and tblAddLess.Vtype=tblVouSub.Vtype and tblAddLess.VNo=tblVouSub.VNo and tblAddLess.VYear=tblVouSub.VYear" _
        & " and tblAddLess.vtype in ('ST') having sum(ADBillAMT)>0"
    Else
        gCn.Execute "insert into tmpGenTbl (Vno,Nar1,Nar2,Amt1,Amt2,Amt3,Amt4,Amt5,Amt6,Amt7,UserName) " _
        & " select 3,'Sale In Transit Sales','Sale In Transit Sales',sum(VSubBag) as bags,sum(decode(Narration,'V A T',AdBillamt-AdTaxAmt,0)) as NetVATSA," _
        & "sum(decode(Narration,'V A T',ADEXEMPTAMT,0)) as ExptAmt,sum(decode(Narration,'V A T',ADTaxAMT,0)) as VAT," _
        & "sum(decode(Narration,'C S T',AdBillamt-AdTaxAmt,0)) as NetCSTSA," _
        & "sum(decode(Narration,'C S T',ADTaxAMT,0)) as CST,sum(ADBillAMT) as CForm,'" & gUserName & "'" _
        & " from tblAddLess,tblMastNarration,tblVouSub where " & CondStr & " and  NarrCode= ADTAXCODE " _
        & " and tblAddLess.Vtype=tblVouSub.Vtype and tblAddLess.VNo=tblVouSub.VNo and tblAddLess.VYear=tblVouSub.VYear" _
        & " and tblAddLess.vtype in ('ST') having sum(ADBillAMT)>0"
    End If
    '--- Exmpted Sales (SY)
    If gBackEndDB = gBackEndAccess Then  '-- Access
        gCn.Execute "insert into tmpGenTbl (Vno,Nar1,Nar2,Amt1,Amt2,Amt3,Amt4,Amt5,Amt6,Amt7,UserName,VSRNO) " _
        & " select 4,'Exempted Sales','Exempted Sales @ ' & ACEXEMPTRATE,sum(VSubBag) as bags,sum(iif(Narration='V A T',AdBillamt-AdTaxAmt,0)) as NetVATSA," _
        & "sum(iif(Narration='V A T',ADEXEMPTAMT,0)) as ExptAmt,sum(iif(Narration='V A T',ADTaxAMT,0)) as VAT," _
        & "sum(iif(Narration='C S T',AdBillamt-AdTaxAmt,0)) as NetCSTSA," _
        & "sum(iif(Narration='C S T',ADTaxAMT,0)) as CST,0 as CForm,'" & gUserName & "',1" _
        & " from tblAddLess,tblMastNarration,tblVouSub,tblMastAccount where " & CondStr & " and  NarrCode= ADTAXCODE " _
        & " and tblAddLess.Vtype=tblVouSub.Vtype and tblAddLess.VNo=tblVouSub.VNo and tblAddLess.VYear=tblVouSub.VYear" _
        & " and  AdMillCode=AcCode and tblAddLess.vtype in ('SY')  and ADEXEMPTAMT>0 " _
        & "group by ACEXEMPTRATE order by ACEXEMPTRATE"
    Else
        gCn.Execute "insert into tmpGenTbl (Vno,Nar1,Nar2,Amt1,Amt2,Amt3,Amt4,Amt5,Amt6,Amt7,UserName,VSRNO) " _
        & " select 4,'Exempted Sales','Exempted Sales @ ' || ACEXEMPTRATE,sum(VSubBag) as bags,sum(decode(Narration,'V A T',AdBillamt-AdTaxAmt,0)) as NetVATSA," _
        & "sum(decode(Narration,'V A T',ADEXEMPTAMT,0)) as ExptAmt,sum(decode(Narration,'V A T',ADTaxAMT,0)) as VAT," _
        & "sum(decode(Narration,'C S T',AdBillamt-AdTaxAmt,0)) as NetCSTSA," _
        & "sum(decode(Narration,'C S T',ADTaxAMT,0)) as CST,0 as CForm,'" & gUserName & "',1" _
        & " from tblAddLess,tblMastNarration,tblVouSub,tblMastAccount where " & CondStr & " and  NarrCode= ADTAXCODE " _
        & " and tblAddLess.Vtype=tblVouSub.Vtype and tblAddLess.VNo=tblVouSub.VNo and tblAddLess.VYear=tblVouSub.VYear" _
        & " and  AdMillCode=AcCode and tblAddLess.vtype in ('SY')  and ADEXEMPTAMT>0 " _
        & "group by ACEXEMPTRATE order by ACEXEMPTRATE"
    End If
    '-- Report
    With CrystalRpt
        If gBackEndDB = gBackEndAccess Then
            .DataFiles(0) = gDbLocation
        Else '-- Oracle
            .Connect = "DSN=" + gODBCDSN + " ;UID=Scott;PWD=" + gODBCPwd + ";DBQ=" + gODBCHostStr
        End If
        .Formulas(4) = "wFrToDt = 'From : ' & '" & FromDt & "' & '  To ' & '" & ToDt & "'"
        x1 = "{tmpGenTbl.Username}='" & gUserName & "'"
        .SelectionFormula = x1
        .ReportFileName = gReportPath & "rptTaxVATComputDetSA.rpt"
        .Action = 1
    End With
    '===================================== Purchase
    gCn.Execute "delete from tmpGenTbl where Username='" & gUserName & "'"
    '-------- Cosign/Depot
    If gBackEndDB = gBackEndAccess Then  '-- Access
        gCn.Execute "insert into tmpGenTbl (Vno,Nar1,Nar2,Amt1,Amt2,Amt3,Amt4,Amt5,Amt6,Amt7,Amt8,UserName,VSRNO) " _
        & " select 5,'Consignment /Depot' as Type,AcName as Account,sum(VSubBag) as bags,sum(iif(Narration='V A T' and VSUBCTRNO=1,AdBillamt-AdTaxAmt,0)) as NetVAT," _
        & "sum(iif(Narration='V A T' and VSUBCTRNO=1,ADEXEMPTAMT,0)) as ExptAmt,sum(iif(Narration='V A T' and VSUBCTRNO=1,ADTaxAMT,0)) as VAT," _
        & "sum(iif(Narration='C S T' and VSUBCTRNO=1,AdBillamt-AdTaxAmt,0)) as NetCST," _
        & "sum(iif(Narration='C S T' and VSUBCTRNO=1,ADTaxAMT,0)) as CST,0 as CForm,sum(iif(VSUBCTRNO=1,adBillAmt,0)) as Consign,'" & gUserName & "',1" _
        & " from tblAddLess,TblMastAccount,tblMastNarration,tblVouSub where NarrCode= ADTAXCODE and AdMillCode=AcCode and " & CondStr & " and tblAddLess.vtype in ('PI') " _
        & " and tblAddLess.Vtype=tblVouSub.Vtype and tblAddLess.VNo=tblVouSub.VNo and tblAddLess.VYear=tblVouSub.VYear" _
        & " group by acname order by acname"
    Else
        gCn.Execute "insert into tmpGenTbl (Vno,Nar1,Nar2,Amt1,Amt2,Amt3,Amt4,Amt5,Amt6,Amt7,Amt8,UserName,VSRNO) " _
        & " select 5,'Consignment /Depot' as Type,AcName as Account,sum(VSubBag) as bags,sum(decode(VSUBCTRNO,1,decode(Narration,'V A T',AdBillamt-AdTaxAmt,0),0)) as NetVAT," _
        & "sum(decode(VSUBCTRNO,1,decode(Narration,'V A T',ADEXEMPTAMT,0),0)) as ExptAmt,sum(decode(VSUBCTRNO,1,decode(Narration,'V A T',ADTaxAMT,0),0)) as VAT," _
        & "sum(decode(VSUBCTRNO,1,decode(Narration,'C S T',AdBillamt-AdTaxAmt,0),0)) as NetCST," _
        & "sum(decode(VSUBCTRNO,1,decode(Narration,'C S T',ADTaxAMT,0),0)) as CST,0 as CForm,sum(decode(VSUBCTRNO,1,adBillAmt,0)) as Consign,'" & gUserName & "',1" _
        & " from tblAddLess,TblMastAccount,tblMastNarration,tblVouSub where NarrCode= ADTAXCODE and AdMillCode=AcCode and " & CondStr & " and tblAddLess.vtype in ('PI') " _
        & " and tblAddLess.Vtype=tblVouSub.Vtype and tblAddLess.VNo=tblVouSub.VNo and tblAddLess.VYear=tblVouSub.VYear" _
        & " group by acname order by acname"
    End If
    '--- Purchase A/c (PY Without Exmpt)
    If gBackEndDB = gBackEndAccess Then  '-- Access
        gCn.Execute "insert into tmpGenTbl (Vno,Nar1,Nar2,Amt1,Amt2,Amt3,Amt4,Amt5,Amt6,Amt7,Amt8,UserName) " _
        & " select 6,'Purchase Account','Purchase Account',sum(VSubBag) as bags,sum(iif(Narration='V A T' and VSUBCTRNO=1,AdBillamt-AdTaxAmt,0)) as NetVAT," _
        & "sum(iif(Narration='V A T' and VSUBCTRNO=1,ADEXEMPTAMT,0)) as ExptAmt,sum(iif(Narration='V A T' and VSUBCTRNO=1,ADTaxAMT,0)) as VAT," _
        & "sum(iif(Narration='C S T' and VSUBCTRNO=1,AdBillamt-AdTaxAmt,0)) as NetCST," _
        & "sum(iif(Narration='C S T' and VSUBCTRNO=1,ADTaxAMT,0)) as CST,0 as CForm,0,'" & gUserName & "'" _
        & " from tblAddLess,tblMastNarration,tblVouSub where " & CondStr & " and  NarrCode= ADTAXCODE " _
        & " and tblAddLess.Vtype=tblVouSub.Vtype and tblAddLess.VNo=tblVouSub.VNo and tblAddLess.VYear=tblVouSub.VYear" _
        & " and  tblAddLess.vtype in ('PY') and ADEXEMPTAMT=0"
    Else
        gCn.Execute "insert into tmpGenTbl (Vno,Nar1,Nar2,Amt1,Amt2,Amt3,Amt4,Amt5,Amt6,Amt7,Amt8,UserName) " _
        & " select 6,'Purchase Account','Purchase Account',sum(VSubBag) as bags,sum(decode(VSUBCTRNO,1,decode(Narration,'V A T',AdBillamt-AdTaxAmt,0),0)) as NetVAT," _
        & "sum(decode(VSUBCTRNO,1,decode(Narration,'V A T',ADEXEMPTAMT,0),0)) as ExptAmt,sum(decode(VSUBCTRNO,1,decode(Narration,'V A T',ADTaxAMT,0),0)) as VAT," _
        & "sum(decode(VSUBCTRNO,1,decode(Narration,'C S T',AdBillamt-AdTaxAmt,0),0)) as NetCST," _
        & "sum(decode(VSUBCTRNO,1,decode(Narration,'C S T',ADTaxAMT,0),0)) as CST,0 as CForm,0,'" & gUserName & "'" _
        & " from tblAddLess,tblMastNarration,tblVouSub where " & CondStr & " and  NarrCode= ADTAXCODE " _
        & " and tblAddLess.Vtype=tblVouSub.Vtype and tblAddLess.VNo=tblVouSub.VNo and tblAddLess.VYear=tblVouSub.VYear" _
        & " and  tblAddLess.vtype in ('PY') and ADEXEMPTAMT=0"
    End If
    '--- SIT Purchase
    If gBackEndDB = gBackEndAccess Then  '-- Access
        gCn.Execute "insert into tmpGenTbl (Vno,Nar1,Nar2,Amt1,Amt2,Amt3,Amt4,Amt5,Amt6,Amt7,Amt8,UserName) " _
        & " select 7,'Sale In Transit Purchase','Sale In Transit Purchase',sum(VSubBag) as bags,sum(iif(Narration='V A T' and VSUBCTRNO=1,AdBillamt-AdTaxAmt,0)) as NetVAT," _
        & "sum(iif(Narration='V A T' and VSUBCTRNO=1,ADEXEMPTAMT,0)) as ExptAmt,sum(iif(Narration='V A T' and VSUBCTRNO=1,ADTaxAMT,0)) as VAT," _
        & "sum(iif(Narration='C S T' and VSUBCTRNO=1,AdBillamt-AdTaxAmt,0)) as NetCST," _
        & "sum(iif(Narration='C S T' and VSUBCTRNO=1,ADTaxAMT,0)) as CST,sum(iif(VSUBCTRNO=1,ADBillAMT,0)) as CForm,0,'" & gUserName & "'" _
        & " from tblAddLess,tblMastNarration,tblVouSub where " & CondStr & " and  NarrCode= ADTAXCODE " _
        & " and tblAddLess.Vtype=tblVouSub.Vtype and tblAddLess.VNo=tblVouSub.VNo and tblAddLess.VYear=tblVouSub.VYear" _
        & " and tblAddLess.vtype in ('PT') having sum(ADBillAMT)>0"
    Else
        gCn.Execute "insert into tmpGenTbl (Vno,Nar1,Nar2,Amt1,Amt2,Amt3,Amt4,Amt5,Amt6,Amt7,Amt8,UserName) " _
        & " select 7,'Sale In Transit Purchase','Sale In Transit Purchase',sum(VSubBag) as bags,sum(decode(VSUBCTRNO,1,decode(Narration,'V A T',AdBillamt-AdTaxAmt,0),0)) as NetVAT," _
        & "sum(decode(VSUBCTRNO,1,decode(Narration,'V A T',ADEXEMPTAMT,0),0)) as ExptAmt,sum(decode(VSUBCTRNO,1,decode(Narration,'V A T',ADTaxAMT,0),0)) as VAT," _
        & "sum(decode(VSUBCTRNO,1,decode(Narration,'C S T',AdBillamt-AdTaxAmt,0),0)) as NetCST," _
        & "sum(decode(VSUBCTRNO,1,decode(Narration,'C S T',ADTaxAMT,0),0)) as CST,sum(decode(VSUBCTRNO,1,ADBillAMT,0)) as CForm,0,'" & gUserName & "'" _
        & " from tblAddLess,tblMastNarration,tblVouSub where " & CondStr & " and  NarrCode= ADTAXCODE " _
        & " and tblAddLess.Vtype=tblVouSub.Vtype and tblAddLess.VNo=tblVouSub.VNo and tblAddLess.VYear=tblVouSub.VYear" _
        & " and tblAddLess.vtype in ('PT') having sum(ADBillAMT)>0"
    End If
    '----- Purchase Other
    If gBackEndDB = gBackEndAccess Then  '-- Access
        gCn.Execute "insert into tmpGenTbl (Vno,Nar1,Nar2,Amt1,Amt2,Amt3,Amt4,Amt5,Amt6,Amt7,Amt8,UserName) " _
        & " select 8,'Purchase Other','Purchase Other',0 as bags,sum(iif(Narration='V A T',AdBillamt-AdTaxAmt,0)) as NetVAT," _
        & "sum(iif(Narration='V A T',ADEXEMPTAMT,0)) as ExptAmt,sum(iif(Narration='V A T',ADTaxAMT,0)) as VAT," _
        & "sum(iif(Narration='C S T',AdBillamt-AdTaxAmt,0)) as NetCST," _
        & "sum(iif(Narration='C S T',ADTaxAMT,0)) as CST,0 as CForm,0,'" & gUserName & "'" _
        & " from tblAddLess,tblMastNarration where " & CondStr & " and  NarrCode= ADTAXCODE " _
        & " and tblAddLess.vtype in ('PO') having sum(ADBillAMT)>0"
    Else
        gCn.Execute "insert into tmpGenTbl (Vno,Nar1,Nar2,Amt1,Amt2,Amt3,Amt4,Amt5,Amt6,Amt7,Amt8,UserName) " _
        & " select 8,'Purchase Other','Purchase Other',0 as bags,sum(decode(Narration,'V A T',AdBillamt-AdTaxAmt,0)) as NetVAT," _
        & "sum(decode(Narration,'V A T',ADEXEMPTAMT,0)) as ExptAmt,sum(decode(Narration,'V A T',ADTaxAMT,0)) as VAT," _
        & "sum(decode(Narration,'C S T',AdBillamt-AdTaxAmt,0)) as NetCST," _
        & "sum(decode(Narration,'C S T',ADTaxAMT,0)) as CST,0 as CForm,0,'" & gUserName & "'" _
        & " from tblAddLess,tblMastNarration where " & CondStr & " and  NarrCode= ADTAXCODE " _
        & " and tblAddLess.vtype in ('PO') having sum(ADBillAMT)>0"
    End If
    '--- Exmpted Purchase (PY)
    If gBackEndDB = gBackEndAccess Then  '-- Access
        gCn.Execute "insert into tmpGenTbl (Vno,Nar1,Nar2,Amt1,Amt2,Amt3,Amt4,Amt5,Amt6,Amt7,Amt8,UserName,VSRNO) " _
        & " select 9,'Exempted Purchase','Exempted Purchase @ ' & ACEXEMPTRATE,sum(VSubBag) as bags,sum(iif(Narration='V A T' and VSUBCTRNO=1,AdBillamt-AdTaxAmt,0)) as NetVAT," _
        & "sum(iif(Narration='V A T' and VSUBCTRNO=1,ADEXEMPTAMT,0)) as ExptAmt,sum(iif(Narration='V A T' and VSUBCTRNO=1,ADTaxAMT,0)) as VAT," _
        & "sum(iif(Narration='C S T' and VSUBCTRNO=1,AdBillamt-AdTaxAmt,0)) as NetCST," _
        & "sum(iif(Narration='C S T' and VSUBCTRNO=1,ADTaxAMT,0)) as CST,0 as CForm,0,'" & gUserName & "',1" _
        & " from tblAddLess,tblMastNarration,tblVouSub,tblMastAccount where " & CondStr & " and  NarrCode= ADTAXCODE " _
        & " and tblAddLess.Vtype=tblVouSub.Vtype and tblAddLess.VNo=tblVouSub.VNo and tblAddLess.VYear=tblVouSub.VYear" _
        & " and  AdMillCode=AcCode and tblAddLess.vtype in ('PY')  and ADEXEMPTAMT>0 " _
        & "group by ACEXEMPTRATE order by ACEXEMPTRATE"
    Else
        gCn.Execute "insert into tmpGenTbl (Vno,Nar1,Nar2,Amt1,Amt2,Amt3,Amt4,Amt5,Amt6,Amt7,Amt8,UserName,VSRNO) " _
        & " select 9,'Exempted Purchase','Exempted Purchase @ ' || ACEXEMPTRATE,sum(VSubBag) as bags,sum(decode(VSUBCTRNO,1,decode(Narration,'V A T',AdBillamt-AdTaxAmt,0),0)) as NetVAT," _
        & "sum(decode(VSUBCTRNO,1,decode(Narration,'V A T',ADEXEMPTAMT,0),0)) as ExptAmt,sum(decode(VSUBCTRNO,1,decode(Narration,'V A T',ADTaxAMT,0),0)) as VAT," _
        & "sum(decode(VSUBCTRNO,1,decode(Narration,'C S T',AdBillamt-AdTaxAmt,0),0)) as NetCST," _
        & "sum(decode(VSUBCTRNO,1,decode(Narration,'C S T',ADTaxAMT,0),0)) as CST,0 as CForm,0,'" & gUserName & "',1" _
        & " from tblAddLess,tblMastNarration,tblVouSub,tblMastAccount where " & CondStr & " and  NarrCode= ADTAXCODE " _
        & " and tblAddLess.Vtype=tblVouSub.Vtype and tblAddLess.VNo=tblVouSub.VNo and tblAddLess.VYear=tblVouSub.VYear" _
        & " and  AdMillCode=AcCode and tblAddLess.vtype in ('PY')  and ADEXEMPTAMT>0 " _
        & "group by ACEXEMPTRATE order by ACEXEMPTRATE"
    End If
    With CrystalRpt
        If gBackEndDB = gBackEndAccess Then
            .DataFiles(0) = gDbLocation
        Else '-- Oracle
            .Connect = "DSN=" + gODBCDSN + " ;UID=Scott;PWD=" + gODBCPwd + ";DBQ=" + gODBCHostStr
        End If
        .Formulas(4) = "wFrToDt = 'From : ' & '" & FromDt & "' & '  To ' & '" & ToDt & "'"
        x1 = "{tmpGenTbl.Username}='" & gUserName & "'"
        .SelectionFormula = x1
        .ReportFileName = gReportPath & "rptTaxVATComputDetPU.rpt"
        .Action = 1
    End With
End Sub
'==================================================== Reports Procedure
'------ Out Standing Sale Late Pay Interest
Public Sub GProcOutStandingSALatePayInt(FromDt As Date, ToDt As Date, AllFirmFlag As Boolean, InterestP As Double, GraseDays As Long, Party As String, IntLateFrmInvDt As Long, IsIntDbNtRpt As Boolean, Optional IsMillBillRec As Boolean)
Dim CondStr As String
Dim X As String
'-- Date
If gBackEndDB = gBackEndAccess Then
    CondStr = " RECVDT>= cdate('" & FromDt & "')  and RECVDT<= cdate('" & ToDt & "')"
Else '-- Oracle
    CondStr = " RECVDT>= to_date('" & FromDt & "','dd/MM/yyyy') and RECVDT<= to_date('" & ToDt & "','dd/MM/yyyy')"
End If
'-- Type
If IsMillBillRec = False Then
    CondStr = CondStr + " and (RECVTYPE='BR' or RECVTYPE='CR' ) "
Else
    CondStr = CondStr + " and RECVTYPE='MR' " '-- Mill Bill Receipt
End If
If Party <> "" Then CondStr = CondStr & " and tblmastaccount.AcName='" & Party & "'" '--Party
'-- Is Interest DbNt
If IsIntDbNtRpt = True Then
    CondStr = CondStr & " and IntDbNtDate is Null" '-IntDbNtDt is Null
End If
X = " select ADJNO,RecVdt,tblmastaccount.AcName,InvBillNo,InvDt,OutAmount,AdjAmt, BILLAMT-AdjAmt,INTRECEIVED,RecVdt- " & IIf(IntLateFrmInvDt = 1, "INVDT-", "OUTDUEDATE-") & GraseDays & ",'" & gUserName & "'," & InterestP & "," & GraseDays & ", " _
& " OUTDUEDATE,tblmastaccount_Br.AcName from tblmastaccount,tblmastaccount tblmastaccount_Br,TblOutstanding,TblRecVsSale where tblmastaccount.accode=PartyCode and " & CondStr & " And " _
& " vno=Invno and vtype=Invtp and vYear=InvYear and TblOutstanding.OutBrokerCode =  tblmastaccount_Br.accode"
'--- insert to setup db Table For All Firm
GProcInsertToSetupDbTbl AllFirmFlag, X, "tmpGenTbl", "Vno,vdt,Nar1,VBillNo,VBillDt,Amt1,amt2,Amt3,Amt4,amt5,username,amt6,amt7,Dt1,Nar3", " username='" & gUserName & "'", "Nar2"
End Sub
'---- Outstanding Late Pay Interest Report Outstanding
Public Sub GProcOutstandingSALatePayIntOS(FromDt As Date, ToDt As Date, AllFirmFlag As Boolean, Broker As String, Party As String, Optional IsMillBillRec As Boolean)
Dim CondStr As String
Dim X As String
'-- Date
If gBackEndDB = gBackEndAccess Then
    CondStr = " INTDBNTDATE>= cdate('" & FromDt & "')  and INTDBNTDATE<= cdate('" & ToDt & "')"
Else '-- Oracle
    CondStr = " INTDBNTDATE>= to_date('" & FromDt & "','dd/MM/yyyy') and INTDBNTDATE<= to_date('" & ToDt & "','dd/MM/yyyy')"
End If
'-- In complete
CondStr = CondStr + " and ISINTDBNTCOMPLETE=0 "
'-- Type
If IsMillBillRec = False Then
    CondStr = CondStr + " and (RECVTYPE='BR' or RECVTYPE='CR' ) "
Else
    CondStr = CondStr + " and RECVTYPE='MR' " '-- Mill Bill Receipt
End If
If Party <> "" Then CondStr = CondStr & " and tblMastAccount.AcName='" & Party & "'" '--Party
If Broker <> "" Then CondStr = CondStr & " and tblAcc_Brk.AcName='" & Broker & "'" '--Broker
X = " select INTDBNTDATE,INVBILLNO,INVDT,INTDBNTAMT,INTDBNTRECDAMT,ISINTDBNTCOMPLETE,tblMastAccount.AcName," _
& "tblAcc_Brk.AcName,'" & gUserName & "' from tblmastaccount,tblmastaccount tblAcc_Brk,TblOutstanding,TblRecVsSale " _
& " where tblmastaccount.accode=PartyCode and tblAcc_Brk.accode=OUTBROKERCODE and " & CondStr & " And " _
& " vno=Invno and vtype=Invtp and vYear=InvYear "
'--- insert to setup db Table For All Firm
GProcInsertToSetupDbTbl AllFirmFlag, X, "tmpGenTbl", "vdt,VBillNo,VBillDt,Amt1,amt2,Amt3,Nar2,Nar1,username", " username='" & gUserName & "'", "Nar3"
End Sub
'-------------- Procedure For OutStanding Aging
Public Sub GProcPrepareAging(OnDt As Date, Optional BrkAcName As String, Optional DueDaysFrom As Long, Optional DueDaysTo As Long)
Dim CondStr As String
Dim BrkCode As Long
Dim DueDayStr As String
BrkCode = GProcGetColumnValue("tblMastAccount", "AcName", BrkAcName, "T", "AcCode", "N")
If gBackEndDB = gBackEndAccess Then
    CondStr = " OUTDUEDATE<= cdate('" & OnDt & "') and Vtype in ('SY','SM','SD','SO','ST')"
    DueDayStr = "cdate('" & CDate(OnDt) & "') " & " - cdate(outDueDate) "
Else '-- Oracle
    CondStr = " OUTDUEDATE<= to_date('" & OnDt & "','dd/MM/yyyy') and Vtype in ('SY','SM','SD','SO','ST')"
    DueDayStr = "to_date('" & CDate(OnDt) & "','dd/MM/yyyy') - outDueDate"
End If
'-- Broker
If BrkCode > 0 Then CondStr = CondStr & " and OUTBROKERCODE=" & BrkCode
gCn.Execute "delete from tmpGenTbl where UserName='" & gUserName & "'"
If DueDaysTo = 0 Then
    '-- 0-15 days
    gCn.Execute "insert into tmpGentbl (VACCODE,VACOCODE,Amt1,Amt2,Amt3,Amt4,UserName)" _
    & " select OUTBROKERCODE,OUTACCODE,Sum(OUTAMOUNT-OUTRECAMT) as Due15,0 as Due30,0 as Due45,0 as " _
    & " DueAbove,'" & gUserName & "' from tbloutstanding where " & CondStr & " and " & DueDayStr & " > 0 and " & DueDayStr & " <30 group by OUTBROKERCODE,OUTACCODE"
    '-- 15-30 days
    gCn.Execute "insert into tmpGentbl (VACCODE,VACOCODE,Amt1,Amt2,Amt3,Amt4,UserName)" _
    & " select OUTBROKERCODE,OUTACCODE,0,Sum(OUTAMOUNT-OUTRECAMT),0,0 " _
    & ",'" & gUserName & "' from tbloutstanding where " & CondStr & " and " & DueDayStr & ">=30 and " & DueDayStr & "<45 group by OUTBROKERCODE,OUTACCODE"
    '-- 30 -45 days
    gCn.Execute "insert into tmpGentbl (VACCODE,VACOCODE,Amt1,Amt2,Amt3,Amt4,UserName)" _
    & " select OUTBROKERCODE,OUTACCODE,0,0,Sum(OUTAMOUNT-OUTRECAMT),0" _
    & ",'" & gUserName & "' from tbloutstanding where " & CondStr & " and " & DueDayStr & ">=45 and " & DueDayStr & "<60 group by OUTBROKERCODE,OUTACCODE"
    '-- 45 above days
    gCn.Execute "insert into tmpGentbl (VACCODE,VACOCODE,Amt1,Amt2,Amt3,Amt4,UserName)" _
    & " select OUTBROKERCODE,OUTACCODE,0,0,0,Sum(OUTAMOUNT-OUTRECAMT)" _
    & ",'" & gUserName & "' from tbloutstanding where " & CondStr & " and " & DueDayStr & ">=60 group by OUTBROKERCODE,OUTACCODE"
    
    gCn.Execute "delete from tmpGenTbl where username='" & gUserName & "' and Amt1=0 and Amt2=0 and " _
    & " amt3=0 and amt4=0"
Else
    '--  FromDays to ToDays
    gCn.Execute "insert into tmpGentbl (VACCODE,VACOCODE,Amt1,Vdt,Amt2,Amt3,UserName)" _
    & " select OUTBROKERCODE,OUTACCODE,Sum(OUTAMOUNT-OUTRECAMT) as Due," & IIf(gBackEndDB = gBackEndAccess, CDate(OnDt), "to_date('" & CDate(OnDt) & "','dd/MM/yyyy')") & " as OnDt," _
    & DueDaysFrom & "," & DueDaysTo & ",'" & gUserName & "' from tbloutstanding where " & CondStr & " and (" & DueDayStr & ") >= " & DueDaysFrom & "  and (" & DueDayStr & " )<= " & DueDaysTo & " group by OUTBROKERCODE,OUTACCODE"
    
    gCn.Execute "delete from tmpGenTbl where username='" & gUserName & "' and Amt1=0 "
End If
End Sub


'============= Common Outstanding
Public Function GProcCollectCommonOutstanding(w_Fr As Date, w_To As Date, Optional w_MillName As String, Optional w_BrkName As String, Optional w_PartyName As String, Optional CurrentFirmFlag As Boolean, Optional ShowUnAdjustReceipt As Boolean)
'''Dim mTempCn As Connection, mCurrFirmDbNm As String
'''Dim RstNewOutst As Recordset, RstOldOutst As Recordset, RstCompMast As Recordset, RstTmpAcMast As Recordset
'''Dim X1 As String, i As Integer, j As Long
'''Dim X As String
''''----- if ShowUnAdjustReceipt then insert UnAdjusted Receipt InTmpGenTbl
'''If ShowUnAdjustReceipt = True Then
'''    If gBackEndDB = gBackEndAccess Then
'''        X1 = " Vdt>=cdate('" & w_Fr & "')  and Vdt<=Cdate('" & w_To & "')"
'''    Else
'''        X1 = " Vdt>=to_Date('" & w_Fr & "','dd/MM/yyyy')  and Vdt<=to_date('" & w_To & "','dd/MM/yyyy')"
'''    End If
'''    If w_PartyName <> "" Then X1 = X1 & " and Party='" & w_PartyName & "'"
'''    X = "select Vno,Vdt,Vtype,Party,CrAccount,Vamt,UNADJUSTAMT,VNAR1,VNAR2,VNAR3,'" & gUserName & "' from QryRegUnadjustReceipt where UnadjustAmt>0 and " & X1
'''    '--- insert to setup db Table For All Firm
'''    GProcInsertToSetupDbTbl Not CurrentFirmFlag, X, "tmpGenTbl", "Vno,Vdt,Vtype,Nar1,Nar2,Amt1,Amt2,Nar3,Nar4,Nar5,username", "username='" & gUserName & "'", "Nar6"
'''End If
''''---- Insert Outstanding Records In TmpTblOutstanding
'''gSetupCn.BeginTrans
'''gSetupCn.Execute "Delete from tmptbloutstanding where username = '" & gUserName & "'"
'''GProcRstOpen RstNewOutst, "Select * from tmptbloutstanding ", "O", gSetupCn
'''GProcRstOpen RstCompMast, "Select * from tblmastcompany " & IIf(CurrentFirmFlag = True, " where cname='" & gCName & "'", ""), "R", gSetupCn
''''----New Temparary Connection
'''Dim PartyGp As String
'''mCurrFirmDbNm = gDbNm
'''RstCompMast.MoveFirst
'''Do While Not RstCompMast.EOF
'''   gDbNm = RstCompMast!CCode + Right(mCurrFirmDbNm, 4)
'''   If gBackEndDB = gBackEndORA Then
'''        GProcConnectionOpenORA mTempCn
'''        X1 = "where vdt <= to_date('" & Format(w_To, "dd/MM/yyyy") & "','dd/MM/yyyy')"
'''        If w_Fr <> gCYSDate Then X1 = X1 + " and vdt >= to_date('" & Format(w_Fr, "dd/MM/yyyy") & "','dd/MM/yyyy')"
'''   Else
'''        GProcConnectionOpen mTempCn
'''        X1 = "where format(vdt,'yyyy/MM/dd') <= '" & Format(w_To, "yyyy/MM/dd") & "'"
'''        If w_Fr <> gCYSDate Then X1 = X1 + " and format(vdt,'yyyy/MM/dd') >= '" & Format(w_Fr, "yyyy/MM/dd") & "'"
'''   End If
'''   If w_MillName <> "" Then X1 = X1 + " and OutMillCode = " & GProcGetColumnValue("tblmastaccount", "Acname", w_MillName, "S", "accode", "N", , mTempCn)
'''   If w_BrkName <> "" Then X1 = X1 + " and OutBrokerCode = " & GProcGetColumnValue("tblmastaccount", "Acname", w_BrkName, "S", "accode", "N", , mTempCn)
'''   If w_PartyName <> "" Then X1 = X1 + " and OutAcCode = " & GProcGetColumnValue("tblmastaccount", "Acname", w_PartyName, "S", "accode", "N", , mTempCn)
'''
'''   If gBackEndDB = gBackEndAccess Then
'''       If gSelectedMenu = "Mill Bill Outstanding Broker with Partywise" Then '"Mill Bill Outstanding Broker with Partywise"
'''            X = "Select * from tbloutstanding " & X1 & " and  vtype in ('SM') and ((OutAmount - OutRecAmt > 0) or (OutDbNtAmt - OutRecDbAmt > 0) )"
'''       Else
'''            X = "Select * from tbloutstanding " & X1 & " and  vtype in ('SD','SY','SM','ST','SO') and ((OutAmount - OutRecAmt > 0) or (OutDbNtAmt - OutRecDbAmt > 0) )"
'''       End If
'''   Else '-- Oracle
'''        If gSelectedMenu = "Mill Bill Outstanding Broker with Partywise" Then '"Mill Bill Outstanding Broker with Partywise"
'''            X = "Select * from tbloutstanding " & X1 & " and  vtype in ('SM') and ((OutAmount - OutRecAmt > 0) or (OutDbNtAmt - OutRecDbAmt > 0) )"
'''        Else  'vtype<> SM
'''            X = "Select * from tbloutstanding " & X1 & " and  vtype in ('SD','SY','ST','SO') and ((OutAmount - OutRecAmt > 0) or (OutDbNtAmt - OutRecDbAmt > 0) )"
'''        End If
'''   End If
'''   If (GProcRstOpen(RstOldOutst, X, "R", mTempCn) > 0) Then
'''       With RstOldOutst
'''            .MoveFirst
'''Loop1:
'''            Do While Not .EOF
'''               'If (!Vtype <> "SO") And (!Vtype <> "SY") And (!Vtype <> "ST") Then
'''               '   .MoveNext
'''               '   GoTo Loop1
'''               'End If
'''               'If (!OutAmount - !OutRecAmt > 0) Or (!OutDbNtAmt - !OutRecDbAmt > 0) Then
'''                  RstNewOutst.AddNew
'''                  For i = 0 To 20
'''                     If (i = 5) Or (i = 6) Or (i = 11) Then
'''                        RstNewOutst.Fields(i) = GProcGetColumnValue("tblmastaccount", "accode", .Fields(i), "N", "acname", "S", , mTempCn) & vbNullString
'''                        If i = 6 Then '- Broker Phone
'''                            RstNewOutst.Fields("OutBrkPhone") = GProcGetColumnValue("tblmastaccount", "acCode", .Fields(i), "N", "AcContNo", "S", , mTempCn) & vbNullString
'''                            RstNewOutst.Fields("OutBrokerPhoneSMS") = GProcGetColumnValue("tblmastaccount", "acCode", .Fields(i), "N", "AcFaxNo", "S", , mTempCn) & vbNullString
'''                        End If
'''                        If i = 5 Then '- Party Phone
''''                            RstNewOutst.Fields("OutPartyPhone") = GProcGetColumnValue("tblmastaccount", "acCode", .Fields(i), "N", "'('+'Ph : '+AcContNo+')'", "S", , mTempCn) & vbNullString
'''                            RstNewOutst.Fields("OutPartyPhone") = GProcGetColumnValue("tblmastaccount", "acCode", .Fields(i), "N", " iif(isnull(AcContPer),'Ph : '+AcContNo+')','('+AcContPer+' - Ph : '+AcContNo+')')", "S", , mTempCn) & vbNullString
'''                            RstNewOutst.Fields("OutPartyPhoneSMS") = GProcGetColumnValue("tblmastaccount", "acCode", .Fields(i), "N", "AcFaxNo", "S", , mTempCn) & vbNullString
'''
'''                            RstNewOutst.Fields("OutCstTin") = GProcGetColumnValue("tblmastaccount", "acCode", .Fields(i), "N", "AcCST", "S", , mTempCn) & vbNullString
'''
'''                        End If
'''                     Else
'''                        RstNewOutst.Fields(i) = .Fields(i)
'''                     End If
'''                  Next i
'''                  PartyGp = GProcGetColumnValue("tblmastaccount", "accode", .Fields(5), "N", "AcCmpCode", "S", , mTempCn) & vbNullString
'''                  RstNewOutst!OutPartyGroup = GProcGetColumnValue("tblmastcompGroup", "CgCode", PartyGp, "N", "Cgname", "S", , mTempCn) & vbNullString
'''                  RstNewOutst!UserName = gUserName
'''                  RstNewOutst!VCompSNm = Left(gDbNm, 4)
'''                  RstNewOutst.Update
'''               'End If
'''               .MoveNext
'''            Loop
'''       End With
'''   End If
'''   mTempCn.Close
'''   RstCompMast.MoveNext
'''Loop
''''RstOldOutst.Close
''''RstNewOutst.Close
'''If RstNewOutst.State = adStateOpen Then RstNewOutst.Close
'''If RstCompMast.State = adStateOpen Then RstCompMast.Close
'''If RstOldOutst.State = adStateOpen Then RstOldOutst.Close
'''gSetupCn.CommitTrans
'''gDbNm = mCurrFirmDbNm
'''GProcConnectionOpen gCn
End Function
'============= Purchase Outstanding
Public Function GProcPurchaseOutstanding(w_Fr As Date, w_To As Date, Optional w_BrkName As String, Optional w_PartyName As String, Optional w_MillName As String, Optional CurrentFirmFlag As Boolean)
'''Dim mTempCn As Connection, mCurrFirmDbNm As String
'''Dim RstNewOutst As Recordset, RstOldOutst As Recordset, RstCompMast As Recordset, RstTmpAcMast As Recordset
'''Dim X1 As String, i As Integer, j As Long
'''Dim X As String
'''
''''---For Unadjusted Payments
'''
'''
'''    If gBackEndDB = gBackEndAccess Then
'''        X1 = " Vdt>=cdate('" & w_Fr & "')  and Vdt<=Cdate('" & w_To & "')"
'''    Else
'''        X1 = " Vdt>=to_Date('" & w_Fr & "','dd/MM/yyyy')  and Vdt<=to_date('" & w_To & "','dd/MM/yyyy')"
'''    End If
'''    If w_PartyName <> "" Then X1 = X1 & " and Party='" & w_PartyName & "'"
'''    X = "select Vno,Vdt,Vtype,Party,CrAccount,Vamt,UNADJUSTAMT,VNAR1,VNAR2,VNAR3,'" & gUserName & "' from QryRegUnadjustReceipt where vtype in('BP','CP') and UnadjustAmt>0 and " & X1
'''    '--- insert to setup db Table For All Firm
'''    GProcInsertToSetupDbTbl Not CurrentFirmFlag, X, "tmpGenTbl", "Vno,Vdt,Vtype,Nar1,Nar2,Amt1,Amt2,Nar3,Nar4,Nar5,username", "username='" & gUserName & "'", "Nar6"
'''
''''---End For Unadjusted Payments
'''
''''---- Insert Outstanding Records In TmpTblOutstanding
'''gSetupCn.BeginTrans
'''gSetupCn.Execute "Delete from tmptbloutstanding where username = '" & gUserName & "'"
'''GProcRstOpen RstNewOutst, "Select * from tmptbloutstanding ", "O", gSetupCn
'''GProcRstOpen RstCompMast, "Select * from tblmastcompany " & IIf(CurrentFirmFlag = True, " where cname='" & gCName & "'", ""), "R", gSetupCn
''''----New Temparary Connection
'''mCurrFirmDbNm = gDbNm
'''RstCompMast.MoveFirst
'''Do While Not RstCompMast.EOF
'''   gDbNm = RstCompMast!CCode + Right(mCurrFirmDbNm, 4)
'''   If gBackEndDB = gBackEndORA Then
'''        GProcConnectionOpenORA mTempCn
'''        X1 = "where vdt <= to_date('" & Format(w_To, "dd/MM/yyyy") & "','dd/MM/yyyy')"
'''        If w_Fr <> gCYSDate Then X1 = X1 + " and vdt >= to_date('" & Format(w_Fr, "dd/MM/yyyy") & "','dd/MM/yyyy')"
'''   Else
'''        GProcConnectionOpen mTempCn
'''        X1 = "where format(vdt,'yyyy/MM/dd') <= '" & Format(w_To, "yyyy/MM/dd") & "'"
'''        If w_Fr <> gCYSDate Then X1 = X1 + " and format(vdt,'yyyy/MM/dd') >= '" & Format(w_Fr, "dd/MM/yyyy") & "'"
'''   End If
'''   If w_MillName <> "" Then X1 = X1 + " and OutMillCode = " & GProcGetColumnValue("tblmastaccount", "Acname", w_MillName, "S", "accode", "N", , mTempCn)
'''   If w_BrkName <> "" Then X1 = X1 + " and OutBrokerCode = " & GProcGetColumnValue("tblmastaccount", "Acname", w_BrkName, "S", "accode", "N", , mTempCn)
'''   If w_PartyName <> "" Then X1 = X1 + " and OutAcCode = " & GProcGetColumnValue("tblmastaccount", "Acname", w_PartyName, "S", "accode", "N", , mTempCn)
'''   X = "Select * from tbloutstanding " & X1 & " and  vtype in ('PY','PT','PO','PI') and ((OutAmount - OutRecAmt > 0) or (OutDbNtAmt - OutRecDbAmt > 0) )"
'''   If (GProcRstOpen(RstOldOutst, X, "R", mTempCn) > 0) Then
'''       With RstOldOutst
'''            .MoveFirst
'''Loop1:
'''            Do While Not .EOF
'''                RstNewOutst.AddNew
'''                For i = 0 To 20
'''                   If (i = 5) Or (i = 6) Or (i = 11) Then
'''                      RstNewOutst.Fields(i) = GProcGetColumnValue("tblmastaccount", "accode", .Fields(i), "N", "acname", "S", , mTempCn) & vbNullString
'''                   Else
'''                      RstNewOutst.Fields(i) = .Fields(i)
'''                   End If
'''                Next i
'''                RstNewOutst!UserName = gUserName
'''                RstNewOutst!VCompSNm = Left(gDbNm, 4)
'''                RstNewOutst.Update
'''               .MoveNext
'''            Loop
'''       End With
'''   End If
'''   mTempCn.Close
'''   RstCompMast.MoveNext
'''Loop
'''If RstNewOutst.State = adStateOpen Then RstNewOutst.Close
'''If RstCompMast.State = adStateOpen Then RstCompMast.Close
'''If RstOldOutst.State = adStateOpen Then RstOldOutst.Close
'''gSetupCn.CommitTrans
'''End Function
'''Public Function GProcPurchaseOutstandingMillBill(w_Fr As Date, w_To As Date, Optional w_BrkName As String, Optional w_PartyName As String, Optional w_MillName As String, Optional CurrentFirmFlag As Boolean)
'''Dim mTempCn As Connection, mCurrFirmDbNm As String
'''Dim RstNewOutst As Recordset, RstOldOutst As Recordset, RstCompMast As Recordset, RstTmpAcMast As Recordset
'''Dim X1 As String, i As Integer, j As Long
'''Dim X As String
'''
''''''---For Unadjusted Payments
'''''
'''''
'''''    If gBackEndDB = gBackEndAccess Then
'''''        X1 = " Vdt>=cdate('" & w_Fr & "')  and Vdt<=Cdate('" & w_To & "')"
'''''    Else
'''''        X1 = " Vdt>=to_Date('" & w_Fr & "','dd/MM/yyyy')  and Vdt<=to_date('" & w_To & "','dd/MM/yyyy')"
'''''    End If
'''''    If w_PartyName <> "" Then X1 = X1 & " and Party='" & w_PartyName & "'"
'''''    X = "select Vno,Vdt,Vtype,Party,CrAccount,Vamt,UNADJUSTAMT,VNAR1,VNAR2,VNAR3,'" & gUserName & "' from QryRegUnadjustReceipt where vtype in('BP','CP') and UnadjustAmt>0 and " & X1
'''''    '--- insert to setup db Table For All Firm
'''''    GProcInsertToSetupDbTbl Not CurrentFirmFlag, X, "tmpGenTbl", "Vno,Vdt,Vtype,Nar1,Nar2,Amt1,Amt2,Nar3,Nar4,Nar5,username", "username='" & gUserName & "'", "Nar6"
'''''
''''''---End For Unadjusted Payments
'''
''''---- Insert Outstanding Records In TmpTblOutstanding
'''gSetupCn.BeginTrans
'''gSetupCn.Execute "Delete from tmptbloutstanding where username = '" & gUserName & "'"
'''GProcRstOpen RstNewOutst, "Select * from tmptbloutstanding ", "O", gSetupCn
'''GProcRstOpen RstCompMast, "Select * from tblmastcompany " & IIf(CurrentFirmFlag = True, " where cname='" & gCName & "'", ""), "R", gSetupCn
''''----New Temparary Connection
'''mCurrFirmDbNm = gDbNm
'''RstCompMast.MoveFirst
'''Do While Not RstCompMast.EOF
'''   gDbNm = RstCompMast!CCode + Right(mCurrFirmDbNm, 4)
'''   If gBackEndDB = gBackEndORA Then
'''        GProcConnectionOpenORA mTempCn
'''        X1 = "where vdt <= to_date('" & Format(w_To, "dd/MM/yyyy") & "','dd/MM/yyyy')"
'''        If w_Fr <> gCYSDate Then X1 = X1 + " and vdt >= to_date('" & Format(w_Fr, "dd/MM/yyyy") & "','dd/MM/yyyy')"
'''   Else
'''        GProcConnectionOpen mTempCn
'''        X1 = "where format(vdt,'yyyy/MM/dd') <= '" & Format(w_To, "yyyy/MM/dd") & "'"
'''        If w_Fr <> gCYSDate Then X1 = X1 + " and format(vdt,'dd/MM/yyyy') >= '" & Format(w_Fr, "dd/MM/yyyy") & "'"
'''   End If
'''   If w_MillName <> "" Then X1 = X1 + " and OutMillCode = " & GProcGetColumnValue("tblmastaccount", "Acname", w_MillName, "S", "accode", "N", , mTempCn)
'''   If w_BrkName <> "" Then X1 = X1 + " and OutBrokerCode = " & GProcGetColumnValue("tblmastaccount", "Acname", w_BrkName, "S", "accode", "N", , mTempCn)
'''   If w_PartyName <> "" Then X1 = X1 + " and OutAcCode = " & GProcGetColumnValue("tblmastaccount", "Acname", w_PartyName, "S", "accode", "N", , mTempCn)
'''   X = "Select * from tbloutstanding " & X1 & " and  vtype = 'SM' and (OutAmountMillBill - OutRecAmtMillBill > 0) "
'''   If (GProcRstOpen(RstOldOutst, X, "R", mTempCn) > 0) Then
'''       With RstOldOutst
'''            .MoveFirst
'''Loop1:
'''            Do While Not .EOF
'''                RstNewOutst.AddNew
'''                For i = 0 To 20
'''                   If (i = 5) Or (i = 6) Or (i = 11) Then
'''                      RstNewOutst.Fields(i) = GProcGetColumnValue("tblmastaccount", "accode", .Fields(i), "N", "acname", "S", , mTempCn) & vbNullString
'''                   Else
'''                      RstNewOutst.Fields(i) = .Fields(i)
'''                   End If
'''                Next i
'''                RstNewOutst!OutAmountMillBill = .Fields("OutAmountMillBill")
'''                RstNewOutst!OutRecAmtMillBill = .Fields("OutRecAmtMillBill")
'''                RstNewOutst!UserName = gUserName
'''                RstNewOutst!VCompSNm = Left(gDbNm, 4)
'''                RstNewOutst.Update
'''               .MoveNext
'''            Loop
'''       End With
'''   End If
'''   mTempCn.Close
'''   RstCompMast.MoveNext
'''Loop
'''If RstNewOutst.State = adStateOpen Then RstNewOutst.Close
'''If RstCompMast.State = adStateOpen Then RstCompMast.Close
'''If RstOldOutst.State = adStateOpen Then RstOldOutst.Close
'''gSetupCn.CommitTrans
End Function
'---- Form Issue /Receive Data
Public Sub GProcFormIssueReceive(FromDt As Date, ToDt As Date, FormRptType As String, Mill As String, Party As String, Broker As String, AllFirmFlag As Boolean, Optional OptStr As String)
Dim CondStr As String
Dim SYr As Long
Dim EYr As Long
Dim X As String
Dim SlbFDt As Date
Dim SlbTDt As Date
Dim i As Long
Select Case FormRptType
    
    Case "CR"   '-- "C Form Receivable"
        If gBackEndDB = gBackEndAccess Then
            CondStr = " tbladdLess.vdt>= cdate('" & FromDt & "')  and tbladdLess.vdt<= cdate('" & ToDt & "')"
            If OptStr = "Pending" Then '-- Pending
               CondStr = CondStr + " and (ADCFORMNO='' or ADCFORMNO is null)" ' (ADCFORMNO is null or ADCFORMNO='') " '-
            ElseIf OptStr = "Receive" Then '-- Receive
                CondStr = CondStr + " and ADCFORMNO <>''"
            End If
        Else '-- Oracle
            CondStr = " tbladdLess.vdt>= to_date('" & FromDt & "','dd/MM/yyyy') and tbladdLess.vdt<= to_date('" & ToDt & "','dd/MM/yyyy')"
            If OptStr = "Pending" Then '-- Pending
                CondStr = CondStr + " and ADCFORMNO is null "
            ElseIf OptStr = "Receive" Then '-- Receive
                CondStr = CondStr + " and ADCFORMNO is not null "
            End If
        End If
    
        If Mill <> "" Then CondStr = CondStr + " and tblmill.AcName='" & Mill & "'" '--Mill
        If Party <> "" Then CondStr = CondStr + " and TblMastAccount.AcName='" & Party & "'" '--Party
        If Broker <> "" Then CondStr = CondStr + " and tblBrk.AcName='" & Broker & "'" '--Broker
        X = " select tbladdLess.VType,ADINVNO,tbladdLess.Vdt,ADBILLAMT,TblMastAccount.AcName,tblBrk.AcName,tblBrk.AcContno,tblMill.AcName, " _
        & " (select sum(VSubBag) from tblvousub where tblvousub.Vno=tbladdless.vno and tblvousub.VType=tbladdless.vType and tblvousub.Vdt=tbladdless.vdt ), " _
        & " (select " & IIf(gBackEndDB = gBackEndAccess, "cdbl(sum(VSubWt))", "sum(VSubWt)") & " from tblvousub where tblvousub.Vno=tbladdless.vno and tblvousub.VType=tbladdless.vType and tblvousub.Vdt=tbladdless.vdt ),'" & gUserName & "'" _
        & " ,tblMastCompGroup.CGName " _
        & " ,ADCFORMNO,ADtmp8,TblMastAccount.AcCst,TblMastAccount.AcContno,tblmastaccount.accity,tblmastaccount.acADD1 from tbladdLess,tblmastaccount,tblmastaccount tblBrk,tblmastaccount  tblMill,tblOutstanding,tblMastNarration,tblMastCompGroup where" _
        & " tblOutstanding.vNo=tbladdLess.vno and tblOutstanding.vType=tbladdLess.vType and tblOutstanding.vdt=tbladdLess.vdt" _
        & " and ADACDRCODE=tblmastaccount.Accode and ADMILLCODE=tblMill.AcCode and tblMastAccount.AcCmpCode = tblMastCompGroup.CGCode  and  " _
        & " OUTBROKERCODE=tblBrk.AcCode and ADTAXCODE=NARRCODE and (instr(narration,'C Form')>0 or instr(narration,'C S T')>0)  and tbladdLess.Vtype in ('SM','SY','ST','SD','SO')  and " & CondStr
        '--- insert to setup db Table For All Firm
        GProcInsertToSetupDbTbl AllFirmFlag, X, "tmpGenTbl", "Vtype,VbillNo,VBillDt,Amt1,Nar2,Nar1,Nar11,Nar4,Amt2,Amt3,username,Nar12,Nar6,Nar13,nar7,nar8,nar9,nar10 ", " username='" & gUserName & "'", "Nar3", "Nar5"
 Case "CI"   '-- "C Form Issuable"
        If gBackEndDB = gBackEndAccess Then
              CondStr = " tbladdLess.AdGpDt>= cdate('" & FromDt & "')  and tbladdLess.AdGpDt<= cdate('" & ToDt & "')"
              If OptStr = "Pending" Then '-- Pending
                 CondStr = CondStr + " and (ADCFORMNO ='' or ADCFORMNO is null) "
              ElseIf OptStr = "Receive" Then '-- Receive
                  CondStr = CondStr + " and ADCFORMNO <>'' "
              End If
        Else '-- Oracle
              CondStr = " tbladdLess.AdGpDt>= to_date('" & FromDt & "','dd/MM/yyyy') and tbladdLess.AdGpDt<= to_date('" & ToDt & "','dd/MM/yyyy')"
              If OptStr = "Pending" Then '-- Pending
                  CondStr = CondStr + " and ADCFORMNO is null "
              ElseIf OptStr = "Receive" Then '-- Receive
                  CondStr = CondStr + " and ADCFORMNO is not null "
              End If
        End If
        If Party <> "" Then CondStr = CondStr + " and TblMastAccount.AcName='" & Party & "'" '--Party
        If Broker <> "" Then CondStr = CondStr + " and tblBrk.AcName='" & Broker & "'" '--Broker
        
        X = " select 1,tblAddless.Vtype,tblAddless.AdFormRecdDate,ADINVNO,tbladdLess.AdGpdt,format(tbladdLess.AdGpdt,'DD/MM/YY'),ADSUBAMT,ADTAXAMT,ADBILLAMT,TblMastAccount.AcName,TblMastAccount.AcState,TblMastAccount.AcFaxNo,TblMastAccount.AcEmail,tblmastaccount.ACADD1 " & IIf(gBackEndDB = gBackEndAccess, " + ", " || ") & " tblmastaccount.AcAdd2,TblMastAccount.ACCITY ,TblMastAccount.ACBST,TblMastAccount.ACCST,tblBrk.AcName," _
        & " (select sum(VSubBag) from tblvousub where tblvousub.Vno=tbladdless.vno and tblvousub.VType=tbladdless.vType and tblvousub.Vdt=tbladdless.vdt ), " _
        & " (select " & IIf(gBackEndDB = gBackEndAccess, "cdbl(sum(VSubWt))", "sum(VSubWt)") & " from tblvousub where tblvousub.Vno=tbladdless.vno and tblvousub.VType=tbladdless.vType and tblvousub.Vdt=tbladdless.vdt ),'" & gUserName & "'" _
        & " ,ADCFORMNO,ADE1FORMNO from tbladdLess,tblmastaccount,tblMastNarration,tblmastaccount  tblBrk,tblOutStanding  where" _
        & " tblOutstanding.vNo=tbladdLess.vno and tblOutstanding.vType=tbladdLess.vType and tblOutstanding.vdt=tbladdLess.vdt" _
        & " and ADACCRCODE=tblmastaccount.Accode and OUTBROKERCODE=tblBrk.AcCode and tblOutstanding.vNo=tbladdLess.vno and tblOutstanding.vType=tbladdLess.vType and tblOutstanding.vdt=tbladdLess.vdt  " _
        & " and OUTBROKERCODE=tblBrk.AcCode and ADTAXCODE=NARRCODE and instr(narration,'C Form')>0 and tbladdLess.Vtype in ('PO','PT','PY')  and " & CondStr
        '--- insert to setup db Table For All Firm
        GProcInsertToSetupDbTbl AllFirmFlag, X, "tmpGenTbl", "vno,Vtype,Dt1,VbillNo,VBillDt,Nar13,Amt5,Amt4,Amt1,Nar2,Nar10,Nar11,Nar12,Nar1,Nar3,Nar4,Nar5,Nar8,Amt2,Amt3,username,Nar7,Nar9", "username='" & gUserName & "'", "Nar6"
 Case "FI"  '---  F Form Issuable
        If gBackEndDB = gBackEndAccess Then
              CondStr = " tbladdLess.AdGpDt>= cdate('" & FromDt & "')  and tbladdLess.AdGpDt<= cdate('" & ToDt & "')"
              If OptStr = "Pending" Then '-- Pending
                 CondStr = CondStr + " and (ADCFORMNO ='' or ADCFORMNO is null) "
              ElseIf OptStr = "Receive" Then '-- Receive
                  CondStr = CondStr + " and ADCFORMNO <>'' "
              End If
        Else '-- Oracle
              CondStr = " tbladdLess.AdGpDt>= to_date('" & FromDt & "','dd/MM/yyyy') and tbladdLess.AdGpDt<= to_date('" & ToDt & "','dd/MM/yyyy')"
              If OptStr = "Pending" Then '-- Pending
                  CondStr = CondStr + " and ADCFORMNO is null "
              ElseIf OptStr = "Receive" Then '-- Receive
                  CondStr = CondStr + " and ADCFORMNO is not null "
              End If
        End If
        If Party <> "" Then CondStr = CondStr + " and TblMastAccount.AcName='" & Party & "'" '--Party
        X = " select 1,Vtype,tblAddless.AdFormRecdDate,ADINVNO,tbladdLess.AdGpdt,ADBILLAMT,TblMastAccount.AcName,TblMastAccount.AcState,TblMastAccount.AcFaxNo,TblMastAccount.AcEmail,tblmastaccount.ACADD1 " & IIf(gBackEndDB = gBackEndAccess, " + ", " || ") & " tblmastaccount.AcAdd2,TblMastAccount.ACCITY ,TblMastAccount.ACBST,TblMastAccount.ACCST," _
        & " (select sum(VSubBag) from tblvousub where tblvousub.Vno=tbladdless.vno and tblvousub.VType=tbladdless.vType and tblvousub.Vdt=tbladdless.vdt ), " _
        & " (select " & IIf(gBackEndDB = gBackEndAccess, "cdbl(sum(VSubWt))", "sum(VSubWt)") & " from tblvousub where tblvousub.Vno=tbladdless.vno and tblvousub.VType=tbladdless.vType and tblvousub.Vdt=tbladdless.vdt ),'" & gUserName & "'" _
        & " ,ADCFORMNO from tbladdLess,tblmastaccount,tblMastNarration where" _
        & "  ADACCRCODE=tblmastaccount.Accode " _
        & " and ADTAXCODE=NARRCODE and tbladdLess.Vtype in ('PI')  and " & CondStr
        '--- insert to setup db Table For All Firm
        GProcInsertToSetupDbTbl AllFirmFlag, X, "tmpGenTbl", "vno,Vtype,Dt1,VbillNo,VBillDt,Amt1,Nar2,Nar10,Nar11,Nar12,Nar1,Nar3,Nar4,Nar5,Amt2,Amt3,username,Nar7", "username='" & gUserName & "'", "Nar6"
 Case "FR" '--- F form Requirement
        If gBackEndDB = gBackEndAccess Then
            CondStr = " tbladdLess.AdGpDt>= cdate('" & FromDt & "')  and tbladdLess.AdGpDt<= cdate('" & ToDt & "')"
        Else '-- Oracle
            CondStr = " tbladdLess.AdGpDt>= to_date('" & FromDt & "','dd/MM/yyyy') and tbladdLess.AdGpDt<= to_date('" & ToDt & "','dd/MM/yyyy')"
        End If
        X = " select Vtype,tblAddless.AdFormRecdDate,ADINVNO,tbladdLess.AdGpdt,format(tbladdLess.AdGpdt,'DD/MM/YY'),ADBILLAMT,TblMastAccount.AcName,TblMastAccount.AcState,TblMastAccount.AcFaxNo,TblMastAccount.AcEmail,tblmastaccount.ACADD1 " & IIf(gBackEndDB = gBackEndAccess, " + ", " || ") & " tblmastaccount.AcAdd2,TblMastAccount.ACCITY ,TblMastAccount.ACBST,TblMastAccount.ACCST," _
        & " (select sum(VSubBag) from tblvousub where tblvousub.Vno=tbladdless.vno and tblvousub.VType=tbladdless.vType and tblvousub.Vdt=tbladdless.vdt ), " _
        & " (select " & IIf(gBackEndDB = gBackEndAccess, "cdbl(sum(VSubWt))", "sum(VSubWt)") & " from tblvousub where tblvousub.Vno=tbladdless.vno and tblvousub.VType=tbladdless.vType and tblvousub.Vdt=tbladdless.vdt ),'" & gUserName & "'" _
        & " from tbladdLess,tblmastaccount where" _
        & "  ADACCRCODE=tblmastaccount.Accode " _
        & " and  tbladdLess.Vtype in ('PI')  and " & CondStr
        '--- insert to setup db Table For All Firm
        GProcInsertToSetupDbTbl AllFirmFlag, X, "tmpGenTbl", "Vtype,Dt1,VbillNo,VBillDt,Nar8,Amt1,Nar2,Nar10,Nar11,Nar12,Nar1,Nar3,Nar4,Nar5,Amt2,Amt3,username", "username='" & gUserName & "'", "Nar6"
' Case "E1R" '--"E1 From Receivable"
'        If gBackEndDB = gBackEndAccess Then
'              CondStr = " tbladdLess.AdGpDt>= cdate('" & FromDt & "')  and tbladdLess.AdGpDt<= cdate('" & ToDt & "')"
'              If OptStr = "Pending" Then '-- Pending
'                 CondStr = CondStr + " and (ADE1FORMNO ='' or ADE1FORMNO is null)"
'              ElseIf OptStr = "Receive" Then '-- Receive
'                  CondStr = CondStr + " and ADE1FORMNO <>'' "
'              End If
'        Else '-- Oracle
'              CondStr = " tbladdLess.AdGpDt>= to_date('" & FromDt & "','dd/MM/yyyy') and tbladdLess.AdGpDt<= to_date('" & ToDt & "','dd/MM/yyyy')"
'              If OptStr = "Pending" Then '-- Pending
'                  CondStr = CondStr + " and ADE1FORMNO is null "
'              ElseIf OptStr = "Receive" Then '-- Receive
'                  CondStr = CondStr + " and ADE1FORMNO is not null "
'              End If
'        End If
'        If Party <> "" Then CondStr = CondStr + " and TblMastAccount.AcName='" & Party & "'" '--Party
'        X = " select 1,tblAddless.Vtype,tblAddless.AdFormE1RecdDate,ADINVNO,tbladdLess.AdGpdt,ADBILLAMT,TblMastAccount.AcName,tblMill.AcName,TblMastAccount.ACCITY ,TblMastAccount.ACBST,TblMastAccount.ACCST,tblBrk.AcName,   " _
'        & " (select sum(tblvousub.VSubBag) from tblvousub where tblvousub.Vno=tbladdless.vno and tblvousub.VType=tbladdless.vType and tblvousub.Vdt=tbladdless.vdt ) as TotBag, " _
'        & " (select " & IIf(gBackEndDB = gBackEndAccess, "cdbl(sum(VSubWt))", "sum(VSubWt)") & " from tblvousub where tblvousub.Vno=tbladdless.vno and tblvousub.VType=tbladdless.vType and tblvousub.Vdt=tbladdless.vdt ) AS TotWt,'" & gUserName & "'" _
'        & " ,ADE1FORMNO from tbladdLess,tblmastaccount,tblmastaccount  tblMill,tblMastNarration, tblmastaccount  tblBrk,tblOutStanding  where" _
'        & "  ADACCRCODE=tblmastaccount.Accode and ADMILLCODE=tblMill.AcCode and OUTBROKERCODE=tblBrk.AcCode and tblOutstanding.vNo=tbladdLess.vno and tblOutstanding.vType=tbladdLess.vType and tblOutstanding.vdt=tbladdLess.vdt" _
'        & " and ADTAXCODE=NARRCODE and instr(narration,'C Form')>0 and tbladdLess.Vtype in ('PT')  and " & CondStr
'        '--- insert to setup db Table For All Firm
'        GProcInsertToSetupDbTbl AllFirmFlag, X, "tmpGenTbl", "Vno,Vtype,Dt1,VbillNo,VBillDt,Amt1,Nar2,Nar1,Nar3,Nar4,Nar5,Nar8,Amt2,Amt3,username,Nar7", "username='" & gUserName & "'", "Nar6"
  
 Case "E1R" '--"E1 From Receivable"
        If gBackEndDB = gBackEndAccess Then
              CondStr = " tbladdLess.AdGpDt>= cdate('" & FromDt & "')  and tbladdLess.AdGpDt<= cdate('" & ToDt & "')"
              If OptStr = "Pending" Then '-- Pending
                 CondStr = CondStr + " and (ADE1FORMNO ='' or ADE1FORMNO is null)"
              ElseIf OptStr = "Receive" Then '-- Receive
                  CondStr = CondStr + " and ADE1FORMNO <>'' "
              End If
        Else '-- Oracle
              CondStr = " tbladdLess.AdGpDt>= to_date('" & FromDt & "','dd/MM/yyyy') and tbladdLess.AdGpDt<= to_date('" & ToDt & "','dd/MM/yyyy')"
              If OptStr = "Pending" Then '-- Pending
                  CondStr = CondStr + " and ADE1FORMNO is null "
              ElseIf OptStr = "Receive" Then '-- Receive
                  CondStr = CondStr + " and ADE1FORMNO is not null "
              End If
        End If
        If Party <> "" Then CondStr = CondStr + " and TblMastAccount.AcName='" & Party & "'" '--Party
        If Broker <> "" Then CondStr = CondStr + " and TblBrk.AcName='" & Broker & "'" '--Party
        X = " select 1,tblAddless.Vtype,tblAddless.AdFormE1RecdDate,ADINVNO,tbladdLess.AdGpdt,ADBILLAMT,TblMastAccount.AcName,TblMastAccount.Acadd1,TblMastAccount.Acadd2,tblMill.AcName,TblMastAccount.ACCITY ,TblMastAccount.ACBST,TblMastAccount.ACCST,tblBrk.AcName,   " _
        & " (select sum(tblvousub.VSubBag) from tblvousub where tblvousub.Vno=tbladdless.vno and tblvousub.VType=tbladdless.vType and tblvousub.Vdt=tbladdless.vdt ) as TotBag, " _
        & " (select " & IIf(gBackEndDB = gBackEndAccess, "cdbl(sum(VSubWt))", "sum(VSubWt)") & " from tblvousub where tblvousub.Vno=tbladdless.vno and tblvousub.VType=tbladdless.vType and tblvousub.Vdt=tbladdless.vdt ) AS TotWt,'" & gUserName & "'" _
        & " ,ADE1FORMNO from tbladdLess,tblmastaccount,tblmastaccount  tblMill,tblMastNarration, tblmastaccount  tblBrk,tblOutStanding  where" _
        & "  ADACCRCODE=tblmastaccount.Accode and ADMILLCODE=tblMill.AcCode and OUTBROKERCODE=tblBrk.AcCode and tblOutstanding.vNo=tbladdLess.vno and tblOutstanding.vType=tbladdLess.vType and tblOutstanding.vdt=tbladdLess.vdt" _
        & " and ADTAXCODE=NARRCODE and instr(narration,'C Form')>0 and tbladdLess.Vtype in ('PT') and tbladdless.AdGpNo = 0 and " & CondStr
       
        '--- insert to setup db Table For All Firm
        GProcInsertToSetupDbTbl AllFirmFlag, X, "tmpGenTbl", "Vno,Vtype,Dt1,VbillNo,VBillDt,Amt1,Nar2,Nar9,nar10,Nar1,Nar3,Nar4,Nar5,Nar8,Amt2,Amt3,username,Nar7", "username='" & gUserName & "'", "Nar6"
  
'' Case "E2R" '--"E2 From Receivable"
''        If gBackEndDB = gBackEndAccess Then
''              CondStr = " tbladdLess.AdGpDt>= cdate('" & FromDt & "')  and tbladdLess.AdGpDt<= cdate('" & ToDt & "')"
''              If OptStr = "Pending" Then '-- Pending
''                 CondStr = CondStr + " and (ADE1FORMNO ='' or ADE1FORMNO is null)"
''              ElseIf OptStr = "Receive" Then '-- Receive
''                  CondStr = CondStr + " and ADE1FORMNO <>'' "
''              End If
''        Else '-- Oracle
''              CondStr = " tbladdLess.AdGpDt>= to_date('" & FromDt & "','dd/MM/yyyy') and tbladdLess.AdGpDt<= to_date('" & ToDt & "','dd/MM/yyyy')"
''              If OptStr = "Pending" Then '-- Pending
''                  CondStr = CondStr + " and ADE1FORMNO is null "
''              ElseIf OptStr = "Receive" Then '-- Receive
''                  CondStr = CondStr + " and ADE1FORMNO is not null "
''              End If
''        End If
''        If Party <> "" Then CondStr = CondStr + " and TblMastAccount.AcName='" & Party & "'" '--Party
''        X = " select 1,tblAddless.Vtype,ADINVNO,tbladdLess.AdGpdt,ADBILLAMT,TblMastAccount.AcName,tblMill.AcName,TblMastAccount.ACCITY ,TblMastAccount.ACBST,TblMastAccount.ACCST,tblBrk.AcName,   " _
''        & " (select sum(tblvousub.VSubBag) from tblvousub where tblvousub.Vno=tbladdless.vno and tblvousub.VType=tbladdless.vType and tblvousub.Vdt=tbladdless.vdt ) as TotBag, " _
''        & " (select " & IIf(gBackEndDB = gBackEndAccess, "cdbl(sum(VSubWt))", "sum(VSubWt)") & " from tblvousub where tblvousub.Vno=tbladdless.vno and tblvousub.VType=tbladdless.vType and tblvousub.Vdt=tbladdless.vdt ) AS TotWt,'" & gUserName & "'" _
''        & " ,ADE1FORMNO from tbladdLess,tblmastaccount,tblmastaccount  tblMill,tblMastNarration, tblmastaccount  tblBrk,tblOutStanding  where" _
''        & "  ADACCRCODE=tblmastaccount.Accode and ADMILLCODE=tblMill.AcCode and OUTBROKERCODE=tblBrk.AcCode and tblOutstanding.vNo=tbladdLess.vno and tblOutstanding.vType=tbladdLess.vType and tblOutstanding.vdt=tbladdLess.vdt" _
''        & " and ADTAXCODE=NARRCODE and instr(narration,'C Form')>0 and tbladdLess.Vtype in ('PT') and tbladdless.AdGpNo = 1 and " & CondStr
''
'''        & " and ADTAXCODE=NARRCODE and instr(narration,'C Form')>0 and tbladdLess.Vtype in ('PT')  and " & CondStr
''        '--- insert to setup db Table For All Firm
''        GProcInsertToSetupDbTbl AllFirmFlag, X, "tmpGenTbl", "Vno,Vtype,VbillNo,VBillDt,Amt1,Nar2,Nar1,Nar3,Nar4,Nar5,Nar8,Amt2,Amt3,username,Nar7", "username='" & gUserName & "'", "Nar6"
     Case "E2I"   '-- "E2 Form Issuable"
        If gBackEndDB = gBackEndAccess Then
            CondStr = " tbladdLess.vdt>= cdate('" & FromDt & "')  and tbladdLess.vdt<= cdate('" & ToDt & "')"
            If OptStr = "Pending" Then '-- Pending
               CondStr = CondStr + " and (ADE1FORMNO='' or ADE1FORMNO is null)" ' (ADCFORMNO is null or ADCFORMNO='') " '-
            ElseIf OptStr = "Receive" Then '-- Receive
                CondStr = CondStr + " and ADE1FORMNO <>''"
            End If
        Else '-- Oracle
            CondStr = " tbladdLess.vdt>= to_date('" & FromDt & "','dd/MM/yyyy') and tbladdLess.vdt<= to_date('" & ToDt & "','dd/MM/yyyy')"
            If OptStr = "Pending" Then '-- Pending
                CondStr = CondStr + " and ADE1FORMNO is null "
            ElseIf OptStr = "Receive" Then '-- Receive
                CondStr = CondStr + " and ADE1FORMNO is not null "
            End If
        End If
    
        If Mill <> "" Then CondStr = CondStr + " and tblmill.AcName='" & Mill & "'" '--Mill
        If Party <> "" Then CondStr = CondStr + " and TblMastAccount.AcName='" & Party & "'" '--Party
        If Broker <> "" Then CondStr = CondStr + " and tblBrk.AcName='" & Broker & "'" '--Broker
        X = " select tbladdLess.VType,tblAddless.AdFormE1RecdDate,ADINVNO,tbladdLess.Vdt,format(tbladdLess.Vdt,'DD/MM/YY'),ADBILLAMT,TblMastAccount.AcName,TblMastAccount.AcState,TblMastAccount.AcFaxNo,TblMastAccount.AcEmail,tblBrk.AcName,tblMill.AcName, " _
        & " (select sum(VSubBag) from tblvousub where tblvousub.Vno=tbladdless.vno and tblvousub.VType=tbladdless.vType and tblvousub.Vdt=tbladdless.vdt ), " _
        & " (select " & IIf(gBackEndDB = gBackEndAccess, "cdbl(sum(VSubWt))", "sum(VSubWt)") & " from tblvousub where tblvousub.Vno=tbladdless.vno and tblvousub.VType=tbladdless.vType and tblvousub.Vdt=tbladdless.vdt ),'" & gUserName & "'" _
        & " ,ADE1FORMNO,TblMastAccount.AcCst,ADCFORMNO from tbladdLess,tblmastaccount,tblmastaccount tblBrk,tblmastaccount  tblMill,tblOutstanding,tblMastNarration where" _
        & " tblOutstanding.vNo=tbladdLess.vno and tblOutstanding.vType=tbladdLess.vType and tblOutstanding.vdt=tbladdLess.vdt" _
        & " and ADACDRCODE=tblmastaccount.Accode and ADMILLCODE=tblMill.AcCode and " _
        & " OUTBROKERCODE=tblBrk.AcCode and ADTAXCODE=NARRCODE and instr(narration,'C Form')>0 and tbladdLess.Vtype in ('ST') and tbladdless.ADTMP2 = 1 and " & CondStr
        '--- insert to setup db Table For All Firm
        GProcInsertToSetupDbTbl AllFirmFlag, X, "tmpGenTbl", "Vtype,Dt1,VbillNo,VBillDt,Nar11,Amt1,Nar2,Nar8,Nar9,Nar10,Nar1,Nar4,amt2,amt3,username,Nar6,nar7,Nar12 ", " username='" & gUserName & "'", "Nar3", "Nar5"
            
 Case "E2R" '--"E2 From Receivable"
        If gBackEndDB = gBackEndAccess Then
              CondStr = " tbladdLess.AdGpDt>= cdate('" & FromDt & "')  and tbladdLess.AdGpDt<= cdate('" & ToDt & "')"
              If OptStr = "Pending" Then '-- Pending
                 CondStr = CondStr + " and (ADE1FORMNO ='' or ADE1FORMNO is null)"
              ElseIf OptStr = "Receive" Then '-- Receive
                  CondStr = CondStr + " and ADE1FORMNO <>'' "
              End If
        Else '-- Oracle
              CondStr = " tbladdLess.AdGpDt>= to_date('" & FromDt & "','dd/MM/yyyy') and tbladdLess.AdGpDt<= to_date('" & ToDt & "','dd/MM/yyyy')"
              If OptStr = "Pending" Then '-- Pending
                  CondStr = CondStr + " and ADE1FORMNO is null "
              ElseIf OptStr = "Receive" Then '-- Receive
                  CondStr = CondStr + " and ADE1FORMNO is not null "
              End If
        End If
        If Party <> "" Then CondStr = CondStr + " and TblMastAccount.AcName='" & Party & "'" '--Party
        X = " select 1,tblAddless.Vtype,tblAddless.AdFormE1RecdDate,ADINVNO,tbladdLess.AdGpdt,ADBILLAMT,TblMastAccount.AcName,tblMill.AcName,TblMastAccount.ACCITY ,TblMastAccount.ACBST,TblMastAccount.ACCST,tblBrk.AcName,   " _
        & " (select sum(tblvousub.VSubBag) from tblvousub where tblvousub.Vno=tbladdless.vno and tblvousub.VType=tbladdless.vType and tblvousub.Vdt=tbladdless.vdt ) as TotBag, " _
        & " (select " & IIf(gBackEndDB = gBackEndAccess, "cdbl(sum(VSubWt))", "sum(VSubWt)") & " from tblvousub where tblvousub.Vno=tbladdless.vno and tblvousub.VType=tbladdless.vType and tblvousub.Vdt=tbladdless.vdt ) AS TotWt,'" & gUserName & "'" _
        & " ,ADE1FORMNO from tbladdLess,tblmastaccount,tblmastaccount  tblMill,tblMastNarration, tblmastaccount  tblBrk,tblOutStanding  where" _
        & "  ADACCRCODE=tblmastaccount.Accode and ADMILLCODE=tblMill.AcCode and OUTBROKERCODE=tblBrk.AcCode and tblOutstanding.vNo=tbladdLess.vno and tblOutstanding.vType=tbladdLess.vType and tblOutstanding.vdt=tbladdLess.vdt" _
        & " and ADTAXCODE=NARRCODE and instr(narration,'C Form')>0 and tbladdLess.Vtype in ('PT') and tbladdless.AdGpNo = 1 and " & CondStr
        
'        & " and ADTAXCODE=NARRCODE and instr(narration,'C Form')>0 and tbladdLess.Vtype in ('PT')  and " & CondStr
        '--- insert to setup db Table For All Firm
        GProcInsertToSetupDbTbl AllFirmFlag, X, "tmpGenTbl", "Vno,Vtype,Dt1,VbillNo,VBillDt,Amt1,Nar2,Nar1,Nar3,Nar4,Nar5,Nar8,Amt2,Amt3,username,Nar7", "username='" & gUserName & "'", "Nar6"
 
  
  
  Case "CRSlabWise" '-- Pending C Form Receivable Slab Wise"
        SYr = Year(gCYSDate)
        EYr = Year(gCYEDate)
        i = 1
        Do While i < 5
            Select Case i
                Case 1 '--Slab1
                    SlbFDt = Format("01/04/" & SYr)
                    SlbTDt = Format("30/06/" & SYr)
                Case 2 '--Slab2
                    SlbFDt = Format("01/07/" & SYr)
                    SlbTDt = Format("30/09/" & SYr)
                Case 3 '--Slab3
                    SlbFDt = Format("01/10/" & SYr)
                    SlbTDt = Format("31/12/" & SYr)
                Case 4 '--Slab4
                    SlbFDt = Format("01/01/" & EYr)
                    SlbTDt = Format("31/03/" & EYr)
            End Select
            If gBackEndDB = gBackEndAccess Then
                CondStr = " tbladdLess.vdt>= cdate('" & SlbFDt & "')  and tbladdLess.vdt<= cdate('" & SlbTDt & "')"
                '-- Pending
                CondStr = CondStr + " and (ADCFORMNO ='' or ADCFORMNO is null)"
'            If OptStr = "Pending" Then '-- Pending
'               CondStr = CondStr + " and (ADCFORMNO='' or ADCFORMNO is null)" ' (ADCFORMNO is null or ADCFORMNO='') " '-
'            ElseIf OptStr = "Receive" Then '-- Receive
'                CondStr = CondStr + " and ADCFORMNO <>''"
'            End If

            Else '-- Oracle
                CondStr = " tbladdLess.vdt>= to_date('" & SlbFDt & "','dd/MM/yyyy') and tbladdLess.vdt<= to_date('" & SlbTDt & "','dd/MM/yyyy')"
                '-- Pending
                CondStr = CondStr + " and ADCFORMNO is null "
            End If
            If Party <> "" Then CondStr = CondStr + " and TblMastAccount.AcName='" & Party & "'" '--Party
            If Broker <> "" Then CondStr = CondStr + " and Tblbrk.AcName='" & Broker & "'" '--Broker
            '-- Union For 4 Slabs
            X = IIf(i > 1, X & " union all ", "") + " select " & i & " ," & gCYear & " ,tbladdLess.VType,tblAddless.AdFormRecdDate,ADINVNO,tbladdLess.Vdt,ADBILLAMT,TblMastAccount.AcName,TblMastAccount.AcContno,tblBrk.AcName,tblBrk.AcContno,tblMill.AcName,'" & gUserName & "'" _
            & " from tbladdLess,tblmastaccount,tblmastaccount tblBrk,tblmastaccount  tblMill,tblOutstanding,tblMastNarration where" _
            & " tblOutstanding.vNo=tbladdLess.vno and tblOutstanding.vType=tbladdLess.vType and tblOutstanding.vdt=tbladdLess.vdt" _
            & " and ADACDRCODE=tblmastaccount.Accode and ADMILLCODE=tblMill.AcCode and " _
            & " OUTBROKERCODE=tblBrk.AcCode and ADTAXCODE=NARRCODE and instr(narration,'C Form')>0 and tbladdLess.Vtype in ('SM','SY','ST','SD','SO')  and " & CondStr
            i = i + 1
        Loop
        '--- insert to setup db Table For All Firm
            GProcInsertToSetupDbTbl AllFirmFlag, X, "tmpGenTbl", "VSrno,VYear,Vtype,Dt1,VbillNo,VBillDt,Amt1,Nar2,Nar6,Nar1,nar7,Nar4,username", " username='" & gUserName & "'", "Nar3", "Nar5"
     Case "E1RSlabWise" '-- Pending E1 Form Receivable Slab Wise"
        SYr = Year(gCYSDate)
        EYr = Year(gCYEDate)
        i = 1
        Do While i < 5
            Select Case i
                Case 1 '--Slab1
                    SlbFDt = Format("01/04/" & SYr)
                    SlbTDt = Format("30/06/" & SYr)
                Case 2 '--Slab2
                    SlbFDt = Format("01/07/" & SYr)
                    SlbTDt = Format("30/09/" & SYr)
                Case 3 '--Slab3
                    SlbFDt = Format("01/10/" & SYr)
                    SlbTDt = Format("31/12/" & SYr)
                Case 4 '--Slab4
                    SlbFDt = Format("01/01/" & EYr)
                    SlbTDt = Format("31/03/" & EYr)
            End Select
            If gBackEndDB = gBackEndAccess Then
                CondStr = " AdGpDt>= cdate('" & SlbFDt & "')  and AdGpDt<= cdate('" & SlbTDt & "')"
                '-- Pending
                CondStr = CondStr + " and (ADE1FORMNO ='' or ADE1FORMNO is null)"
            Else '-- Oracle
                CondStr = " AdGpDt>= to_date('" & SlbFDt & "','dd/MM/yyyy') and AdGpDt<= to_date('" & SlbTDt & "','dd/MM/yyyy')"
                '-- Pending
                CondStr = CondStr + " and ADE1FORMNO is null "
            End If
            If Mill <> "" Then CondStr = CondStr + " and tblmastaccount.AcName='" & Mill & "'"  '--Mill
            If Broker <> "" Then CondStr = CondStr + " and tblBrk.AcName='" & Broker & "'" '--Broker
            '-- Union For 4 Slabs
            X = IIf(i > 1, X & " union all ", "") + " select " & i & " ," & gCYear & " ,tblAddless.VType,tblAddless.AdFormE1RecdDate,ADINVNO,AdGpdt,ADBILLAMT,tblmastaccount.AcName,tblmastaccount.AcContno,tblbrk.Acname,'" & gUserName & "'" _
            & " from tbladdLess,tblmastaccount,tblmastaccount tblBrk,tblOutstanding,tblMastNarration where ADMILLCODE=tblmastaccount.AcCode and " _
            & " tblOutstanding.vNo=tbladdLess.vno and tblOutstanding.vType=tbladdLess.vType and tblOutstanding.vdt=tbladdLess.vdt" _
            & " and OUTBROKERCODE=tblBrk.AcCode and ADTAXCODE=NARRCODE and instr(narration,'C Form')>0 and tbladdLess.Vtype in ('PT')  and tbladdLess.adgpno <> 1 and " & CondStr
            i = i + 1
        Loop
        '--- insert to setup db Table For All Firm
            GProcInsertToSetupDbTbl AllFirmFlag, X, "tmpGenTbl", "VSrno,VYear,Vtype,Dt1,VbillNo,VBillDt,Amt1,Nar4,NAR6,nar7,username", " username='" & gUserName & "'", "Nar3", "Nar5"
End Select
End Sub
'---- Form F Form Statement
Public Sub GProcFFormStatement(FromDt As Date, ToDt As Date, Mill As String, AllFirmFlag As Boolean)
Dim CondStr As String
Dim X As String
        If gBackEndDB = gBackEndAccess Then
            CondStr = " tbladdLess.AdGpDt>= cdate('" & FromDt & "')  and tbladdLess.AdGpDt<= cdate('" & ToDt & "')"
        Else '-- Oracle
            CondStr = " tbladdLess.AdGpDt>= to_date('" & FromDt & "','dd/MM/yyyy') and tbladdLess.AdGpDt<= to_date('" & ToDt & "','dd/MM/yyyy')"
        End If
        X = " select Vtype,ADINVNO,tbladdLess.AdGpdt,ADBILLAMT,TblMastAccount.AcName,tblmastaccount.ACADD1 " & IIf(gBackEndDB = gBackEndAccess, " + ", " || ") & " tblmastaccount.AcAdd2,TblMastAccount.ACCITY ,TblMastAccount.ACBST,TblMastAccount.ACCST," _
        & " (select sum(VSubBag) from tblvousub where tblvousub.Vno=tbladdless.vno and tblvousub.VType=tbladdless.vType and tblvousub.Vdt=tbladdless.vdt ), " _
        & " (select sum(VSubBag) from tblvousub where tblvousub.Vno=tbladdless.vno and tblvousub.VType=tbladdless.vType and tblvousub.Vdt=tbladdless.vdt ),'" & gUserName & "'" _
        & " from tbladdLess,tblmastaccount where" _
        & "  ADACCRCODE=tblmastaccount.Accode " _
        & " and  tbladdLess.Vtype in ('PI')  and " & CondStr
        '--- insert to setup db Table For All Firm
        GProcInsertToSetupDbTbl AllFirmFlag, X, "tmpGenTbl", "Vtype,VbillNo,VBillDt,Amt1,Nar2,Nar1,Nar3,Nar4,Nar5,Amt2,Amt3,username", "username='" & gUserName & "'", "Nar6"
End Sub
'==================================================== Reports  Procedure End
'============================================== General Reports Procedure
'---- Sales Summary
Public Sub GProcSalesSummary(FromDt As Date, ToDt As Date, AllFirmFlag As Boolean, SelOptStr As String, RptTypeStr As String, Optional Mill As String, Optional Party As String, Optional Broker As String, Optional Item As String, Optional ItMill As String)
Dim CondStr As String
Dim RelationStr As String
Dim TotStr As String
Dim X As String
    If gBackEndDB = gBackEndAccess Then
        CondStr = " tbladdless.VDt>= cdate('" & FromDt & "')  and tbladdless.VDt<= cdate('" & ToDt & "')"
    Else '-- Oracle
        CondStr = " tbladdless.VDt>= to_date('" & FromDt & "','dd/MM/yyyy') and tbladdless.VDt<= to_date('" & ToDt & "','dd/MM/yyyy')"
    End If
    Select Case RptTypeStr
        Case "Total" '--Total  Sales
            '-Relation Str
            RelationStr = " from tbladdless,tblvouSub where tbladdless.vno=tblVousub.vno and tbladdless.vType=tblVousub.vType and tbladdless.VYear=tblVousub.Vyear and "
            '-- Total Str
            TotStr = IIf(SelOptStr = "NAmt", "sum(ADBILLAMT-AdTaxAmt)", IIf(SelOptStr = "GAmt", "sum(ADBILLAMT)", " Sum(VSUBBAG)"))
            '-- SY,SO,SD,ST,SM
            X = "select " & TotStr & ",0,0,0,0,'" & gUserName & "'" _
            & RelationStr & " tbladdless.Vtype = 'SY' and " & CondStr _
            & " union all select 0," & TotStr & ",0,0,0,'" & gUserName & "'" _
            & RelationStr & " tbladdless.Vtype = 'SO' and " & CondStr _
            & " union all select 0,0," & TotStr & ",0,0,'" & gUserName & "'" _
            & RelationStr & " tbladdless.Vtype = 'SD' and " & CondStr _
            & " union all select 0,0,0," & TotStr & ",0,'" & gUserName & "'" _
            & RelationStr & " tbladdless.Vtype = 'ST' and " & CondStr _
            & " union all select 0,0,0,0," & TotStr & ",'" & gUserName & "'" _
            & RelationStr & " tbladdless.Vtype = 'SM' and " & CondStr
            '--- insert to setup db Table For All Firm
            GProcInsertToSetupDbTbl AllFirmFlag, X, "tmpGenTbl", "Amt1,Amt2,Amt3,Amt4,Amt5,username", "username='" & gUserName & "'", "Nar5"
       Case "MonthWise" '--Monthwise  Sales
            If gBackEndDB = gBackEndAccess Then
                CondStr = CondStr + " group by Format(tbladdless.vdt,'MMMM'),format(tbladdless.vdt,'MM'),Format(tbladdless.vdt,'yy')"
            Else '--Oracle
                CondStr = CondStr + " group by to_char(tbladdless.vdt,'Month'),to_char(tbladdless.vdt,'MM'),to_char(tbladdless.vdt,'yy')"
            End If
            '-Relation Str
            RelationStr = IIf(gBackEndDB = gBackEndAccess, ",Format(tbladdless.vdt,'MMMM'),format(tbladdless.vdt,'MM'),Format(tbladdless.vdt,'yy')", ",to_char(tbladdless.vdt,'Month'),to_char(tbladdless.vdt,'MM'),to_char(tbladdless.vdt,'yy')") & " from tbladdless,tblvouSub" _
            & " where tbladdless.vno=tblVousub.vno and tbladdless.vType=tblVousub.vType and tbladdless.VYear=tblVousub.Vyear and "
            '-- Total Str
            TotStr = IIf(SelOptStr = "NAmt", "sum(ADBILLAMT-AdTaxAmt)", IIf(SelOptStr = "GAmt", "sum(ADBILLAMT)", " Sum(VSUBBAG)"))
            '-- SY,SO,SD,ST,SM
            X = "select " & TotStr & ",0,0,0,0,'" & gUserName & "'" _
            & RelationStr & " tbladdless.Vtype = 'SY' and " & CondStr _
            & " union all select 0," & TotStr & ",0,0,0,'" & gUserName & "'" _
            & RelationStr & " tbladdless.Vtype = 'SO' and " & CondStr _
            & " union all select 0,0," & TotStr & ",0,0,'" & gUserName & "'" _
            & RelationStr & " tbladdless.Vtype = 'SD' and " & CondStr _
            & " union all select 0,0,0," & TotStr & ",0,'" & gUserName & "'" _
            & RelationStr & " tbladdless.Vtype = 'ST' and " & CondStr _
            & " union all select 0,0,0,0," & TotStr & ",'" & gUserName & "'" _
            & RelationStr & " tbladdless.Vtype = 'SM' and " & CondStr
            '--- insert to setup db Table For All Firm
            GProcInsertToSetupDbTbl AllFirmFlag, X, "tmpGenTbl", "Amt1,Amt2,Amt3,Amt4,Amt5,username,Nar1,Vno,Vyear", "username='" & gUserName & "'", "Nar5"
       Case "CountWise" '--- Countwise Sales
            If Item <> "" Then CondStr = CondStr & " and ItName='" & Item & "'" '--count
            CondStr = CondStr + " group by ItName,AcName"
             '-Relation Str
            RelationStr = " ItName,acname,'" & gUserName & "' from tbladdless,tblmastaccount,tblmastItem,tblVousub where " _
            & " tbladdless.vno=tblVousub.vno and tbladdless.vType=tblVousub.vType and tbladdless.VYear=tblVousub.Vyear " _
            & " And VSUBITCODE = ItCode And ITMILLCODE = AcCode And "
            '- Total Str
            TotStr = IIf(SelOptStr = "NAmt", "sum(ADBILLAMT-AdTaxAmt)", IIf(SelOptStr = "GAmt", "sum(ADBILLAMT)", " Sum(VSUBBAG)"))
            '-- SY,SO,SD,ST,SM
            X = "select " & TotStr & ",0,0,0,0," _
            & RelationStr & " tbladdless.Vtype = 'SY' and " & CondStr _
            & " union all select 0," & TotStr & ",0,0,0," _
            & RelationStr & " tbladdless.Vtype = 'SO' and " & CondStr _
            & " union all select 0,0," & TotStr & ",0,0," _
            & RelationStr & " tbladdless.Vtype = 'SD' and " & CondStr _
            & " union all select 0,0,0," & TotStr & ",0," _
            & RelationStr & " tbladdless.Vtype = 'ST' and " & CondStr _
            & " union all select 0,0,0,0," & TotStr & "," _
            & RelationStr & "tbladdless.Vtype = 'SM' and " & CondStr
            '--- insert to setup db Table For All Firm
            GProcInsertToSetupDbTbl AllFirmFlag, X, "tmpGenTbl", "Amt1,Amt2,Amt3,Amt4,Amt5,Nar4,Nar3,username", "username='" & gUserName & "'", "Nar5"
       Case "MillWise" '--MillWise  Sales
            If Mill <> "" Then CondStr = CondStr & " and acname='" & Mill & "'" '--Mill
            CondStr = CondStr + " group by acname"
            '-Relation Str
            RelationStr = "acname,'" & gUserName & "' from tbladdless,tblmastaccount,tblVousub  where ADMILLCODE=accode and " _
            & " tbladdless.vno=tblVousub.vno and tbladdless.vType=tblVousub.vType and tbladdless.VYear=tblVousub.Vyear and "
            '- Total Str
            TotStr = IIf(SelOptStr = "NAmt", "sum(ADBILLAMT-AdTaxAmt)", IIf(SelOptStr = "GAmt", "sum(ADBILLAMT)", " Sum(VSUBBAG)"))
            '-- SY,SO,SD,ST,SM
            X = "select " & TotStr & ",0,0,0,0," _
            & RelationStr & "tbladdless.Vtype = 'SY' and " & CondStr _
            & " union all select 0," & TotStr & ",0,0,0," _
            & RelationStr & "tbladdless.Vtype = 'SO' and " & CondStr _
            & " union all select 0,0," & TotStr & ",0,0," _
            & RelationStr & "tbladdless.Vtype = 'SD' and " & CondStr _
            & " union all select 0,0,0," & TotStr & ",0," _
            & RelationStr & "tbladdless.Vtype = 'ST' and " & CondStr _
            & " union all select 0,0,0,0," & TotStr & "," _
            & RelationStr & "tbladdless.Vtype = 'SM' and " & CondStr
            '--- insert to setup db Table For All Firm
            GProcInsertToSetupDbTbl AllFirmFlag, X, "tmpGenTbl", "Amt1,Amt2,Amt3,Amt4,Amt5,Nar3,username", "username='" & gUserName & "'", "Nar5"
        Case "PartyWise" '--PartyWise  Sales
            If Party <> "" Then CondStr = CondStr & " and acname='" & Party & "'" '--Party
            CondStr = CondStr + " group by acname"
            '-Relation Str
            RelationStr = "acname,'" & gUserName & "' from tbladdless,tblmastaccount,tblVousub where  ADACDRCODE=accode and " _
            & " tbladdless.vno=tblVousub.vno and tbladdless.vType=tblVousub.vType and tbladdless.VYear=tblVousub.Vyear and "
            '- Total Str
            TotStr = IIf(SelOptStr = "NAmt", "sum(ADBILLAMT-AdTaxAmt)", IIf(SelOptStr = "GAmt", "sum(ADBILLAMT)", " Sum(VSUBBAG)"))
            '-- SY,SO,SD,ST,SM
            X = "select " & TotStr & ",0,0,0,0," _
            & RelationStr & "tbladdless.Vtype = 'SY' and " & CondStr _
            & " union all select 0," & TotStr & ",0,0,0," _
            & RelationStr & "tbladdless.Vtype = 'SO' and " & CondStr _
            & " union all select 0,0," & TotStr & ",0,0," _
            & RelationStr & "tbladdless.Vtype = 'SD' and " & CondStr _
            & " union all select 0,0,0," & TotStr & ",0," _
            & RelationStr & "tbladdless.Vtype = 'ST' and " & CondStr _
            & " union all select 0,0,0,0," & TotStr & "," _
            & RelationStr & "tbladdless.Vtype = 'SM' and " & CondStr
            '--- insert to setup db Table For All Firm
            GProcInsertToSetupDbTbl AllFirmFlag, X, "tmpGenTbl", "Amt1,Amt2,Amt3,Amt4,Amt5,Nar1,username", "username='" & gUserName & "'", "Nar5"
        Case "MillCountWise" '-- Mill count wise
            If Item <> "" Then CondStr = CondStr & " and ItName='" & Item & "'" '--Item
            If Mill <> "" Then CondStr = CondStr & " and AcName='" & Mill & "'" '--Mill
            CondStr = CondStr + " group by ItName,AcName"
            '--Relation Str
            RelationStr = "ItName,acname,'" & gUserName & "' from tbladdless,tblmastaccount,tblmastItem,tblVousub where " _
            & " tbladdless.vno=tblVousub.vno and tbladdless.vType=tblVousub.vType and tbladdless.VYear=tblVousub.Vyear and VSUBITCODE=Itcode and ITMILLCODE=accode and "
            '-- Total str
            TotStr = IIf(SelOptStr = "NAmt", "sum(ADBILLAMT-AdTaxAmt)", IIf(SelOptStr = "GAmt", "sum(ADBILLAMT)", " Sum(VSUBBAG)"))
            '-- SY,SO,SD,ST,SM
            X = "select " & TotStr & ",0,0,0,0," _
            & RelationStr & "tbladdless.Vtype = 'SY' and " & CondStr _
            & " union all select 0," & TotStr & ",0,0,0," _
            & RelationStr & "tbladdless.Vtype = 'SO' and " & CondStr _
            & " union all select 0,0," & TotStr & ",0,0," _
            & RelationStr & "tbladdless.Vtype = 'SD' and " & CondStr _
            & " union all select 0,0,0," & TotStr & ",0," _
            & RelationStr & "tbladdless.Vtype = 'ST' and " & CondStr _
            & " union all select 0,0,0,0," & TotStr & "," _
            & RelationStr & "tbladdless.Vtype = 'SM' and " & CondStr
            '--- insert to setup db Table For All Firm
            GProcInsertToSetupDbTbl AllFirmFlag, X, "tmpGenTbl", "Amt1,Amt2,Amt3,Amt4,Amt5,Nar4,Nar3,username", "username='" & gUserName & "'", "Nar5"
     Case "PartyCountWise" '-- Party count wise
            If Item <> "" Then CondStr = CondStr & " and ItName='" & Item & "'" '--Item
            If Party <> "" Then CondStr = CondStr & " and tblmastaccount.AcName='" & Party & "'" '--Party
            CondStr = CondStr + " group by ItName,tblmastaccount.AcName,tblmill.acname"
            '--Relation str
            RelationStr = " ItName,tblmill.acname,tblmastaccount.acname,'" & gUserName & "' from tbladdless,tblmastaccount,tblmastItem,tblVousub,tblmastaccount tblmill where " _
            & "tbladdless.vno=tblVousub.vno and tbladdless.vType=tblVousub.vType and tbladdless.VYear=tblVousub.Vyear and ADACDRCODE=tblmastaccount.AcCode  and VSUBITCODE=Itcode and ITMILLCODE=tblmill.accode and "
            '-- Total Str
            TotStr = IIf(SelOptStr = "NAmt", "sum(ADBILLAMT-AdTaxAmt)", IIf(SelOptStr = "GAmt", "sum(ADBILLAMT)", " Sum(VSUBBAG)"))
            '-- SY,SO,SD,ST,SM
            X = "select " & TotStr & ",0,0,0,0," _
            & RelationStr & "tbladdless.Vtype = 'SY' and " & CondStr _
            & " union all select 0," & TotStr & ",0,0,0," _
            & RelationStr & "tbladdless.Vtype = 'SO' and " & CondStr _
            & " union all select 0,0," & TotStr & ",0,0," _
            & RelationStr & "tbladdless.Vtype = 'SD' and " & CondStr _
            & " union all select 0,0,0," & TotStr & ",0," _
            & RelationStr & "tbladdless.Vtype = 'ST' and " & CondStr _
            & " union all select 0,0,0,0," & TotStr & "," _
            & RelationStr & "tbladdless.Vtype = 'SM' and " & CondStr
            '--- insert to setup db Table For All Firm
            GProcInsertToSetupDbTbl AllFirmFlag, X, "tmpGenTbl", "Amt1,Amt2,Amt3,Amt4,Amt5,Nar4,Nar3,Nar1,username", "username='" & gUserName & "'", "Nar5"
      Case "BrokerCountWise" '-- Broker count wise
            If Item <> "" Then CondStr = CondStr & " and ItName='" & Item & "'" '--Item
            If Broker <> "" Then CondStr = CondStr & " and AcName='" & Broker & "'"  '--Broker
            CondStr = CondStr + " group by ItName,ITTICKET,AcName"
            '-- Relation Str
            RelationStr = IIf(gBackEndDB = gBackEndAccess, " ItName & '(' & ITTICKET & ')'", "ItName || ' ( ' || ITTICKET || ')'") & ",acname,'" & gUserName & "' from " _
            & " tbladdless,tblmastaccount,tblmastItem,tblVousub,tblOutStanding where " _
            & " tbladdless.vno=tblVousub.vno and tbladdless.vType=tblVousub.vType and tbladdless.VYear=tblVousub.Vyear and VSUBITCODE=Itcode and " _
            & " tbladdless.vno=tblOutStanding.vno and tbladdless.vType=tblOutStanding.vType and tbladdless.VYear=tblOutStanding.Vyear and OUTBROKERCODE=accode and "
            '-- Total Str
            TotStr = IIf(SelOptStr = "NAmt", "sum(ADBILLAMT-AdTaxAmt)", IIf(SelOptStr = "GAmt", "sum(ADBILLAMT)", " Sum(VSUBBAG)"))
            '-- SY,SO,SD,ST,SM
            X = "select " & TotStr & ",0,0,0,0," _
            & RelationStr & "tbladdless.Vtype = 'SY' and " & CondStr _
            & " union all select 0," & TotStr & ",0,0,0," _
            & RelationStr & "tbladdless.Vtype = 'SO' and " & CondStr _
            & " union all select 0,0," & TotStr & ",0,0," _
            & RelationStr & "tbladdless.Vtype = 'SD' and " & CondStr _
            & " union all select 0,0,0," & TotStr & ",0," _
            & RelationStr & "tbladdless.Vtype = 'ST' and " & CondStr _
            & " union all select 0,0,0,0," & TotStr & "," _
            & RelationStr & "tbladdless.Vtype = 'SM' and " & CondStr
            '--- insert to setup db Table For All Firm
            GProcInsertToSetupDbTbl AllFirmFlag, X, "tmpGenTbl", "Amt1,Amt2,Amt3,Amt4,Amt5,Nar4,Nar2,username", "username='" & gUserName & "'", "Nar5"
       Case "BrokerPartyWise" '-- Broker Party wise
            If Party <> "" Then CondStr = CondStr & " and tblParty.AcName='" & Party & "'" '--Party
            If Broker <> "" Then CondStr = CondStr & " and tblmastaccount.AcName='" & Broker & "'"  '--Broker
            CondStr = CondStr + " group by tblParty.AcName,tblmastaccount.AcName"
            '--Relation Str
            RelationStr = " tblParty.AcName,tblmastaccount.AcName,'" & gUserName & "' from " _
            & " tbladdless,tblVousub,tblmastaccount,tblmastaccount tblParty,tblOutStanding where ADACDRCODE=tblParty.AcCode and " _
            & " tbladdless.vno=tblVousub.vno and tbladdless.vType=tblVousub.vType and tbladdless.VYear=tblVousub.Vyear and  " _
            & " tbladdless.vno=tblOutStanding.vno and tbladdless.vType=tblOutStanding.vType and tbladdless.VYear=tblOutStanding.Vyear and OUTBROKERCODE=tblmastaccount.accode and "
            '-- Total Str
            TotStr = IIf(SelOptStr = "NAmt", "sum(ADBILLAMT-AdTaxAmt)", IIf(SelOptStr = "GAmt", "sum(ADBILLAMT)", " Sum(VSUBBAG)"))
            '-- SY,SO,SD,ST,SM
            X = "select " & TotStr & ",0,0,0,0," _
            & RelationStr & "tbladdless.Vtype = 'SY' and " & CondStr _
            & " union all select 0," & TotStr & ",0,0,0," _
            & RelationStr & "tbladdless.Vtype = 'SO' and " & CondStr _
            & " union all select 0,0," & TotStr & ",0,0," _
            & RelationStr & "tbladdless.Vtype = 'SD' and " & CondStr _
            & " union all select 0,0,0," & TotStr & ",0," _
            & RelationStr & "tbladdless.Vtype = 'ST' and " & CondStr _
            & " union all select 0,0,0,0," & TotStr & "," _
            & RelationStr & "tbladdless.Vtype = 'SM' and " & CondStr
            '--- insert to setup db Table For All Firm
            GProcInsertToSetupDbTbl AllFirmFlag, X, "tmpGenTbl", "Amt1,Amt2,Amt3,Amt4,Amt5,Nar1,Nar2,username", "username='" & gUserName & "'", "Nar5"
    End Select
    gSetupCn.BeginTrans
    gSetupCn.Execute "delete from tmpGenTbl where Amt1+Amt2+Amt3+Amt4+Amt5=0 and username='" & gUserName & "'"
    gSetupCn.CommitTrans
End Sub
'---- Prepare TmpRptTbl For Receipt Vs Invoice Narration
Public Sub GProcPrepareTmpRptTblForRecInvNarration(FromDt As Date, ToDt As Date, PartyCode As Long, Optional Criteria As String)
Dim X As String
Dim x1 As String
Dim Narr As String
Dim Rs1  As Recordset
Dim RecVno As Long
Dim RecVType As String
Dim RecVYear As Long
If gBackEndDB = gBackEndAccess Then
    X = " RecVDt>= cdate('" & FromDt & "')  and RecVDt<= cdate('" & ToDt & "')"
Else '-- Oracle
    X = " RecVDt>= to_date('" & FromDt & "','dd/MM/yyyy') and RecVDt<= to_date('" & ToDt & "','dd/MM/yyyy')"
End If
X = X & " and PARTYCODE= " & PartyCode
If Criteria <> "" Then X = X & " and " & Criteria
gCn.Execute "delete from tmpRptTbl Where Username='" & gUserName & "'"
'-- Insert in tmpRptTbl Vno=RecVno,Vtype=RecVtype,Vyear=RecVyear
gCn.Execute "insert into tmpRptTbl (UserName,Vno,Vtype,Vyear) select distinct '" & gUserName & "',RECVNO,RECVTYPE,RECVYEAR  from tblRecVsSale where " & X
'--Set tmpRptTbl Nar1=Narration                  (Bill No : amt ,..., On A/c : Rec amt -Sum(inv Adj. Amt))
x1 = "select '" & gUserName & "',RECVNO,RECVTYPE,RECVYEAR,INVBILLNO,ADJAMT,0 as OrdNo1,RECVCTRLNO as OrdNo2 from tblRecVsSale where " & X & "" _
& " union all select '" & gUserName & "',RECVNO,RECVTYPE,RECVYEAR ,'On A/c',RECAMT- sum(ADJAMT),1,0  from tblRecVsSale where " & X _
& " group by RECVNO,RECVTYPE,RECVYEAR,RECAMT having RECAMT- sum(ADJAMT) <>0 order by RECVTYPE,RECVNO,RECVYEAR,OrdNo1,OrdNo2"
If GProcRstOpen(Rs1, x1, "R") > 0 Then
    With Rs1
        .MoveFirst
        RecVno = !RecVno
        RecVType = !RecVType
        RecVYear = !RecVYear
                Narr = "Bill Nos :"
        Do While Not .EOF
            If RecVno = !RecVno And RecVType = !RecVType And RecVYear = !RecVYear Then
                Narr = Narr & IIf(Narr <> "", ", ", "") & !InvBillno & ": " & !AdjAmt
            Else
                gCn.Execute "update tmpRptTbl set Nar1='" & Left(Narr, 250) & "' where UserName='" & gUserName & "' and Vno=" & RecVno _
                & "  and Vtype='" & RecVType & "' and VYear=" & RecVYear
                RecVno = !RecVno
                RecVType = !RecVType
                RecVYear = !RecVYear
                Narr = "Bill Nos :"
                Narr = Narr & IIf(Narr <> "", ", ", "") & !InvBillno & ": " & !AdjAmt
            End If
            Rs1.MoveNext
        Loop
        '-- 250 Characters Of Narration
        gCn.Execute "update tmpRptTbl set Nar1='" & Left(Narr, 250) & "' where UserName='" & gUserName & "' and Vno=" & RecVno _
        & "  and Vtype='" & RecVType & "' and VYear=" & RecVYear
    End With
End If
Rs1.Close
End Sub
'---- Purchase Summary
Public Sub GProcPurchaseSummary(FromDt As Date, ToDt As Date, AllFirmFlag As Boolean, SelOptStr As String, RptTypeStr As String, Optional Mill As String, Optional Item As String)
Dim CondStr As String
Dim RelationStr As String
Dim TotStr As String
Dim X As String
    If gBackEndDB = gBackEndAccess Then
        CondStr = " tbladdless.VDt>= cdate('" & FromDt & "')  and tbladdless.VDt<= cdate('" & ToDt & "')"
    Else '-- Oracle
        CondStr = " tbladdless.VDt>= to_date('" & FromDt & "','dd/MM/yyyy') and tbladdless.VDt<= to_date('" & ToDt & "','dd/MM/yyyy')"
    End If
    Select Case RptTypeStr
        Case "Total" '--Total  Purchase
            '-Relation Str
            RelationStr = " from tbladdless,tblvouSub where tbladdless.vno=tblVousub.vno and tbladdless.vType=tblVousub.vType and tbladdless.VYear=tblVousub.Vyear and " & IIf(SelOptStr = "Bag", "", " vsubctrno=1 and ")
            '-- Total Str
            TotStr = IIf(SelOptStr = "NAmt", "sum(ADBILLAMT-AdTaxAmt)", IIf(SelOptStr = "GAmt", "sum(ADBILLAMT)", " Sum(VSUBBAG)"))
            '-- PY,PT
            X = "select " & TotStr & ",0,'" & gUserName & "'" _
            & RelationStr & " tbladdless.Vtype = 'PY' and " & CondStr _
            & " union all select 0," & TotStr & ",'" & gUserName & "'" _
            & RelationStr & " tbladdless.Vtype = 'PT' and " & CondStr
            '--- insert to setup db Table For All Firm
            GProcInsertToSetupDbTbl AllFirmFlag, X, "tmpGenTbl", "Amt1,Amt2,username", "username='" & gUserName & "'", "Nar5"
       Case "MonthWise" '--Monthwise  Sales/ Purchase
            If gBackEndDB = gBackEndAccess Then
                CondStr = CondStr + " group by Format(tbladdless.vdt,'MMMM'),format(tbladdless.vdt,'MM'),Format(tbladdless.vdt,'yy')"
            Else '--Oracle
                CondStr = CondStr + " group by to_char(tbladdless.vdt,'Month'),to_char(tbladdless.vdt,'MM'),to_char(tbladdless.vdt,'yy')"
            End If
            '-Relation Str
            RelationStr = IIf(gBackEndDB = gBackEndAccess, ",Format(tbladdless.vdt,'MMMM'),format(tbladdless.vdt,'MM'),Format(tbladdless.vdt,'yy')", ",to_char(tbladdless.vdt,'Month'),to_char(tbladdless.vdt,'MM'),to_char(tbladdless.vdt,'yy')") & " from tbladdless,tblvouSub" _
            & " where tbladdless.vno=tblVousub.vno and tbladdless.vType=tblVousub.vType and tbladdless.VYear=tblVousub.Vyear and " & IIf(SelOptStr = "Bag", "", " vsubctrno=1 and ")
            '-- Total Str
            TotStr = IIf(SelOptStr = "NAmt", "sum(ADBILLAMT-AdTaxAmt)", IIf(SelOptStr = "GAmt", "sum(ADBILLAMT)", " Sum(VSUBBAG)"))
            '-- SY,ST,PY,PT
            X = "select " & TotStr & ",0,'" & gUserName & "'" _
            & RelationStr & " tbladdless.Vtype in ('SY','ST')  and " & CondStr _
            & " union all select 0," & TotStr & ",'" & gUserName & "'" _
            & RelationStr & " tbladdless.Vtype in ('PY','PT')  and " & CondStr _
            '--- insert to setup db Table For All Firm
            GProcInsertToSetupDbTbl AllFirmFlag, X, "tmpGenTbl", "Amt1,Amt2,username,Nar1,Vno,Vyear", "username='" & gUserName & "'", "Nar5"
       Case "MonthWiseP" '--Monthwise  Sales/ Purchase
            If gBackEndDB = gBackEndAccess Then
                CondStr = CondStr + " group by Format(tbladdless.vdt,'MMMM'),format(tbladdless.vdt,'MM'),Format(tbladdless.vdt,'yy')"
            Else '--Oracle
                CondStr = CondStr + " group by to_char(tbladdless.vdt,'Month'),to_char(tbladdless.vdt,'MM'),to_char(tbladdless.vdt,'yy')"
            End If
            '-Relation Str
            RelationStr = IIf(gBackEndDB = gBackEndAccess, ",Format(tbladdless.vdt,'MMMM'),format(tbladdless.vdt,'MM'),Format(tbladdless.vdt,'yy')", ",to_char(tbladdless.vdt,'Month'),to_char(tbladdless.vdt,'MM'),to_char(tbladdless.vdt,'yy')") & " from tbladdless,tblvouSub" _
            & " where tbladdless.vno=tblVousub.vno and tbladdless.vType=tblVousub.vType and tbladdless.VYear=tblVousub.Vyear and " & IIf(SelOptStr = "Bag", "", " vsubctrno=1 and ")
            
            '-- Total Str
            TotStr = IIf(SelOptStr = "NAmt", "sum(ADBILLAMT-AdTaxAmt)", IIf(SelOptStr = "GAmt", "sum(ADBILLAMT)", " Sum(VSUBBAG)"))
            '-- PY,PT,PI
            X = "select " & TotStr & ",0,0,'" & gUserName & "'" _
            & RelationStr & " tbladdless.Vtype = 'PY' and " & CondStr _
            & " union all select 0," & TotStr & ",0,'" & gUserName & "'" _
            & RelationStr & " tbladdless.Vtype = 'PT' and " & CondStr _
            & " union all select 0,0," & TotStr & ",'" & gUserName & "'" _
            & RelationStr & " tbladdless.Vtype = 'PI' and " & CondStr _
            '--- insert to setup db Table For All Firm
            GProcInsertToSetupDbTbl AllFirmFlag, X, "tmpGenTbl", "Amt1,Amt2,Amt3,username,Nar1,Vno,Vyear", "username='" & gUserName & "'", "Nar5"
       Case "CountWise" '--- Countwise Purchase
            If Item <> "" Then CondStr = CondStr & " and ItName='" & Item & "'" '--count
            CondStr = CondStr + " group by ItName,AcName"
             '-Relation Str
            RelationStr = " ItName,acname,'" & gUserName & "' from tbladdless,tblmastaccount,tblmastItem,tblVousub where " _
            & " tbladdless.vno=tblVousub.vno and tbladdless.vType=tblVousub.vType and tbladdless.VYear=tblVousub.Vyear " _
            & " And VSUBITCODE = ItCode And ITMILLCODE = AcCode And "
            '- Total Str
            TotStr = IIf(SelOptStr = "NAmt", "sum(VSUBAMT)", IIf(SelOptStr = "GAmt", "sum(VSUBAMT)", " Sum(VSUBBAG)"))
            '-- PY,PT
            X = "select " & TotStr & ",0," _
            & RelationStr & " tbladdless.Vtype = 'PY' and " & CondStr _
            & " union all select 0," & TotStr & "," _
            & RelationStr & " tbladdless.Vtype = 'PT' and " & CondStr
            '--- insert to setup db Table For All Firm
            GProcInsertToSetupDbTbl AllFirmFlag, X, "tmpGenTbl", "Amt1,Amt2,Nar4,Nar3,username", "username='" & gUserName & "'", "Nar5"
       Case "MillWise" '--MillWise  Purchase
            If Mill <> "" Then CondStr = CondStr & " and acname='" & Mill & "'" '--Mill
            CondStr = CondStr + " group by acname"
            '-Relation Str
            RelationStr = "acname,'" & gUserName & "' from tbladdless,tblmastaccount,tblVousub  where ADMILLCODE=accode and " _
            & " tbladdless.vno=tblVousub.vno and tbladdless.vType=tblVousub.vType and tbladdless.VYear=tblVousub.Vyear and " & IIf(SelOptStr = "Bag", "", " vsubctrno=1 and ")
            '- Total Str
            TotStr = IIf(SelOptStr = "NAmt", "sum(ADBILLAMT-AdTaxAmt)", IIf(SelOptStr = "GAmt", "sum(ADBILLAMT)", " Sum(VSUBBAG)"))
            '-- PY,PT
            X = "select " & TotStr & ",0," _
            & RelationStr & "tbladdless.Vtype = 'PY' and " & CondStr _
            & " union all select 0," & TotStr & "," _
            & RelationStr & "tbladdless.Vtype = 'PT' and " & CondStr
            '--- insert to setup db Table For All Firm
            GProcInsertToSetupDbTbl AllFirmFlag, X, "tmpGenTbl", "Amt1,Amt2,Nar3,username", "username='" & gUserName & "'", "Nar5"
    End Select
    gSetupCn.BeginTrans
    gSetupCn.Execute "delete from tmpGenTbl where Amt1+Amt2=0 and username='" & gUserName & "'"
    gSetupCn.CommitTrans
End Sub
'---- Inward Summary
Public Sub GProcInwardSummary(FromDt As Date, ToDt As Date, AllFirmFlag As Boolean, SelOptStr As String, RptTypeStr As String, Optional Mill As String, Optional Item As String)
Dim CondStr As String
Dim RelationStr As String
Dim TotStr As String
Dim X As String
    If gBackEndDB = gBackEndAccess Then
        CondStr = " tbladdless.VDt>= cdate('" & FromDt & "')  and tbladdless.VDt<= cdate('" & ToDt & "')"
    Else '-- Oracle
        CondStr = " tbladdless.VDt>= to_date('" & FromDt & "','dd/MM/yyyy') and tbladdless.VDt<= to_date('" & ToDt & "','dd/MM/yyyy')"
    End If
    Select Case RptTypeStr
        Case "Total" '--Total  Inward
            '-Relation Str
            RelationStr = " from tbladdless,tblvouSub where tbladdless.vno=tblVousub.vno and tbladdless.vType=tblVousub.vType and tbladdless.VYear=tblVousub.Vyear and " & IIf(SelOptStr = "Bag", "", " vsubctrno=1 and ")
            '-- Total Str
            TotStr = IIf(SelOptStr = "NAmt", "sum(ADBILLAMT-AdTaxAmt)", IIf(SelOptStr = "GAmt", "sum(ADBILLAMT)", " Sum(VSUBBAG)"))
            '-- PI
            X = "select " & TotStr & ",'" & gUserName & "'" _
            & RelationStr & " tbladdless.Vtype = 'PI' and " & CondStr
            '--- insert to setup db Table For All Firm
            GProcInsertToSetupDbTbl AllFirmFlag, X, "tmpGenTbl", "Amt1,username", "username='" & gUserName & "'", "", "Nar5"
       Case "CountWise" '--- Countwise Inward
            If Item <> "" Then CondStr = CondStr & " and ItName='" & Item & "'" '--count
            CondStr = CondStr + " group by ItName,AcName"
             '-Relation Str
            RelationStr = " ItName,acname,'" & gUserName & "' from tbladdless,tblmastaccount,tblmastItem,tblVousub where " _
            & " tbladdless.vno=tblVousub.vno and tbladdless.vType=tblVousub.vType and tbladdless.VYear=tblVousub.Vyear " _
            & " And VSUBITCODE = ItCode And ITMILLCODE = AcCode And "
            '- Total Str
            TotStr = IIf(SelOptStr = "NAmt", "sum(VSUBAMT)", IIf(SelOptStr = "GAmt", "sum(VSUBAMT)", " Sum(VSUBBAG)"))
            '-- PI
            X = "select " & TotStr & "," _
            & RelationStr & " tbladdless.Vtype = 'PI' and " & CondStr
            '--- insert to setup db Table For All Firm
            GProcInsertToSetupDbTbl AllFirmFlag, X, "tmpGenTbl", "Amt1,Nar4,Nar3,username", "username='" & gUserName & "'", "", "Nar5"
       Case "MillWise" '--MillWise  Inward
            If Mill <> "" Then CondStr = CondStr & " and acname='" & Mill & "'" '--Mill
            CondStr = CondStr + " group by acname"
            '-Relation Str
            RelationStr = "acname,'" & gUserName & "' from tbladdless,tblmastaccount,tblVousub  where ADMILLCODE=accode and " _
            & " tbladdless.vno=tblVousub.vno and tbladdless.vType=tblVousub.vType and tbladdless.VYear=tblVousub.Vyear and " & IIf(SelOptStr = "Bag", "", " vsubctrno=1 and ")
            '- Total Str
            TotStr = IIf(SelOptStr = "NAmt", "sum(ADBILLAMT-AdTaxAmt)", IIf(SelOptStr = "GAmt", "sum(ADBILLAMT)", " Sum(VSUBBAG)"))
            '-- PI
            X = "select " & TotStr & "," _
            & RelationStr & "tbladdless.Vtype = 'PI' and " & CondStr _
            '--- insert to setup db Table For All Firm
            GProcInsertToSetupDbTbl AllFirmFlag, X, "tmpGenTbl", "Amt1,Nar3,username", "username='" & gUserName & "'", "", "Nar5"
    End Select
    gSetupCn.BeginTrans
    gSetupCn.Execute "delete from tmpGenTbl where Amt1=0 and username='" & gUserName & "'"
    gSetupCn.CommitTrans
End Sub
'----Graph  Sales
Public Sub GProcGraphSales(FromDt As Date, ToDt As Date, AllFirmFlag As Boolean, SelOptStr As String, SAType As String, Optional NumDivideBy As Long, Optional MillAcCode As Long, Optional BrokAcCode As Long, Optional ItemCode As Long)
Dim CondStr As String
Dim RelationStr As String
Dim TotStr As String
Dim X As String
    '== Date
    If gBackEndDB = gBackEndAccess Then
        CondStr = " tbladdless.VDt>= cdate('" & FromDt & "')  and tbladdless.VDt<= cdate('" & ToDt & "')"
    Else '-- Oracle
        CondStr = " tbladdless.VDt>= to_date('" & FromDt & "','dd/MM/yyyy') and tbladdless.VDt<= to_date('" & ToDt & "','dd/MM/yyyy')"
    End If
    '== Mill,Broker,Item
    If MillAcCode > 0 Then
        CondStr = CondStr + " and ADMILLCODE=" & MillAcCode
    ElseIf BrokAcCode > 0 Then
        CondStr = CondStr + " and OUTBROKERCODE=" & BrokAcCode
    ElseIf ItemCode > 0 Then
        CondStr = CondStr + " and VSUBITCODE=" & ItemCode
    End If
    '=== Month,Year
    If gBackEndDB = gBackEndAccess Then
        CondStr = CondStr + " group by Format(tbladdless.vdt,'MMMM'),format(tbladdless.vdt,'MM'),Format(tbladdless.vdt,'yy') order by format(tbladdless.vdt,'MM'),Format(tbladdless.vdt,'yy')"
    Else '--Oracle
        CondStr = CondStr + " group by to_char(tbladdless.vdt,'Month'),to_char(tbladdless.vdt,'MM'),to_char(tbladdless.vdt,'yy') order by to_char(tbladdless.vdt,'MM'),to_char(tbladdless.vdt,'yy') "
    End If
    '-Relation Str
    RelationStr = IIf(gBackEndDB = gBackEndAccess, ",Format(tbladdless.vdt,'MMMM'),format(tbladdless.vdt,'MM'),Format(tbladdless.vdt,'yy')", ",to_char(tbladdless.vdt,'Month'),to_char(tbladdless.vdt,'MM'),to_char(tbladdless.vdt,'yy')") & " from tbladdless,tblvouSub,tblOutstanding" _
    & " where tbladdless.vno=tblVousub.vno and tbladdless.vType=tblVousub.vType and tbladdless.VYear=tblVousub.Vyear and " _
    & "  tbladdless.vno=tblOutstanding.vno and tbladdless.vType=tblOutstanding.vType and tbladdless.VYear=tblOutstanding.Vyear and "
    '-- Total Str
    TotStr = IIf(SelOptStr = "Amt", "sum(ADBILLAMT)", " Sum(VSUBBAG)")
    '-- SY,SO,SD,ST,SM
    X = "select " & TotStr & ",'" & gUserName & "'" _
    & RelationStr & " tbladdless.Vtype " & IIf(SAType = "", "in ('SY','SO','SD','ST','SM')", "='" + SAType & "'") & " And " & CondStr
    '--- insert to setup db Table For All Firm
    GProcInsertToSetupDbTbl AllFirmFlag, X, "tmpGenTbl", "Amt1,username,Nar1,Vno,Vyear", "username='" & gUserName & "'", "Nar5"
    '---Number Format ( Divide by)
    gSetupCn.BeginTrans
    If NumDivideBy > 0 Then gSetupCn.Execute "update tmpGenTbl set Amt1=amt1/" & NumDivideBy & " where username='" & gUserName & "'"
    gSetupCn.CommitTrans
End Sub
'----Graph Purchase
Public Sub GProcGraphPurchase(FromDt As Date, ToDt As Date, AllFirmFlag As Boolean, SelOptStr As String, PUType As String, Optional NumDivideBy As Long, Optional MillAcCode As Long, Optional ItemCode As Long)
Dim CondStr As String
Dim RelationStr As String
Dim TotStr As String
Dim X As String
    '== Date
    If gBackEndDB = gBackEndAccess Then
        CondStr = " tbladdless.VDt>= cdate('" & FromDt & "')  and tbladdless.VDt<= cdate('" & ToDt & "')"
    Else '-- Oracle
        CondStr = " tbladdless.VDt>= to_date('" & FromDt & "','dd/MM/yyyy') and tbladdless.VDt<= to_date('" & ToDt & "','dd/MM/yyyy')"
    End If
    '== Mill,Item
    If MillAcCode > 0 Then
        CondStr = CondStr + " and ADMILLCODE=" & MillAcCode
    ElseIf ItemCode > 0 Then
        CondStr = CondStr + " and VSUBITCODE=" & ItemCode
    End If
    '=== Month,Year
    If gBackEndDB = gBackEndAccess Then
        CondStr = CondStr + " group by Format(tbladdless.vdt,'MMMM'),format(tbladdless.vdt,'MM'),Format(tbladdless.vdt,'yy') order by format(tbladdless.vdt,'MM'),Format(tbladdless.vdt,'yy')"
    Else '--Oracle
        CondStr = CondStr + " group by to_char(tbladdless.vdt,'Month'),to_char(tbladdless.vdt,'MM'),to_char(tbladdless.vdt,'yy') order by to_char(tbladdless.vdt,'MM'),to_char(tbladdless.vdt,'yy') "
    End If
    '-Relation Str
    RelationStr = IIf(gBackEndDB = gBackEndAccess, ",Format(tbladdless.vdt,'MMMM'),format(tbladdless.vdt,'MM'),Format(tbladdless.vdt,'yy')", ",to_char(tbladdless.vdt,'Month'),to_char(tbladdless.vdt,'MM'),to_char(tbladdless.vdt,'yy')") & " from tbladdless,tblvouSub" _
    & " where tbladdless.vno=tblVousub.vno and tbladdless.vType=tblVousub.vType and tbladdless.VYear=tblVousub.Vyear and "
    '-- Total Str
    TotStr = IIf(SelOptStr = "Amt", IIf(gBackEndDB = gBackEndAccess, "sum(iif(VSUBCTRNO=1,ADBILLAMT,0))", "sum(decode(VSUBCTRNO,1,ADBILLAMT,0))"), " Sum(VSUBBAG)")
    '-- PY,PT,PI
    X = "select " & TotStr & ",'" & gUserName & "'" _
    & RelationStr & " tbladdless.Vtype " & IIf(PUType = "", "in ('PY','PT')", "='" + PUType & "'") & " And " & CondStr
    '--- insert to setup db Table For All Firm
    GProcInsertToSetupDbTbl AllFirmFlag, X, "tmpGenTbl", "Amt1,username,Nar1,Vno,Vyear", "username='" & gUserName & "'", "Nar5"
    '---Number Format ( Divide by)
    gSetupCn.BeginTrans
    If NumDivideBy > 0 Then gSetupCn.Execute "update tmpGenTbl set Amt1=amt1/" & NumDivideBy & " where username='" & gUserName & "'"
    gSetupCn.CommitTrans
End Sub
'----Graph Purchase Vs Sales
Public Sub GProcGraphPurchaseVsSale(FromDt As Date, ToDt As Date, AllFirmFlag As Boolean, SelOptStr As String, PUType As String, Optional NumDivideBy As Long, Optional MillAcCode As Long, Optional ItemCode As Long)
Dim CondStr As String
Dim RelationStr As String
Dim TotStr As String
Dim X As String
Dim SATypeStr As String
    '== Date
    If gBackEndDB = gBackEndAccess Then
        CondStr = " tbladdless.VDt>= cdate('" & FromDt & "')  and tbladdless.VDt<= cdate('" & ToDt & "')"
    Else '-- Oracle
        CondStr = " tbladdless.VDt>= to_date('" & FromDt & "','dd/MM/yyyy') and tbladdless.VDt<= to_date('" & ToDt & "','dd/MM/yyyy')"
    End If
    '== Mill,Item
    If MillAcCode > 0 Then
        CondStr = CondStr + " and ADMILLCODE=" & MillAcCode
    ElseIf ItemCode > 0 Then
        CondStr = CondStr + " and VSUBITCODE=" & ItemCode
    End If
    '=== Month,Year
    If gBackEndDB = gBackEndAccess Then
        CondStr = CondStr + " group by Format(tbladdless.vdt,'MMMM'),format(tbladdless.vdt,'MM'),Format(tbladdless.vdt,'yy') "
    Else '--Oracle
        CondStr = CondStr + " group by to_char(tbladdless.vdt,'Month'),to_char(tbladdless.vdt,'MM'),to_char(tbladdless.vdt,'yy') "
    End If
    '-Relation Str
    RelationStr = IIf(gBackEndDB = gBackEndAccess, ",Format(tbladdless.vdt,'MMMM'),format(tbladdless.vdt,'MM'),Format(tbladdless.vdt,'yy')", ",to_char(tbladdless.vdt,'Month'),to_char(tbladdless.vdt,'MM'),to_char(tbladdless.vdt,'yy')") & " from tbladdless,tblvouSub" _
    & " where tbladdless.vno=tblVousub.vno and tbladdless.vType=tblVousub.vType and tbladdless.VYear=tblVousub.Vyear and "
    '-- Total Str
    TotStr = IIf(SelOptStr = "Amt", IIf(gBackEndDB = gBackEndAccess, "sum(iif(VSUBCTRNO=1,ADBILLAMT,0))", "sum(decode(VSUBCTRNO,1,ADBILLAMT,0))"), " Sum(VSUBBAG)")
    '-- Purchase Vs Sale
    If PUType = "" Then '-All
        SATypeStr = " in ('SY','ST')"
    ElseIf PUType = "PY" Then '-Trade
        SATypeStr = "='SY'"
    ElseIf PUType = "PT" Then '-SIT
        SATypeStr = "='ST'"
    ElseIf PUType = "PI" Then '- Inward
        SATypeStr = "in ('SD','SO')"
    End If
    X = "select " & TotStr & ",0,'" & gUserName & "'" _
    & RelationStr & " tbladdless.Vtype " & IIf(PUType = "", "in ('PY','PT')", "='" + PUType & "'") & " And " & CondStr _
    & " union all select 0," & TotStr & ",'" & gUserName & "'" _
    & RelationStr & " tbladdless.Vtype " & SATypeStr & " And " & CondStr & " order by 4,5"
    '--- insert to setup db Table For All Firm
    GProcInsertToSetupDbTbl AllFirmFlag, X, "tmpGenTbl", "Amt1,Amt2,username,Nar1,Vno,Vyear", "username='" & gUserName & "'", "Nar5"
    '---Number Format ( Divide by)
    gSetupCn.BeginTrans
    If NumDivideBy > 0 Then gSetupCn.Execute "update tmpGenTbl set Amt1=amt1/" & NumDivideBy & ",amt2=Amt2/" & NumDivideBy & " where username='" & gUserName & "'"
    gSetupCn.CommitTrans
End Sub
'-------------- Consignment Depot Mill Statement ( used for Consignment Depot Mill statement)
Public Sub GProcConsigDepotMillStatement(FromDt As Date, ToDt As Date, CrystalRpt As CrystalReport, MillName As String, TDSP As Double, ServiceECess As Double, ServiceTax As Double, TaxBillNo As String, Surcharge As Double, TDSEcess As Double, TaxBilldt As Date)
Dim x1 As String
Dim Rs As Recordset
Dim i As Long
Dim CondStr As String
Dim CommType As String '- Commission Type From A/c Master
Dim Amt As Double
Dim MillCode As Long
Dim RptCount As Long
Dim RptEnclosure(15) As String
MillCode = GProcGetColumnValue("tblMastAccount", "AcName", MillName, "S", "AcCode", "N")
With CrystalRpt
.Formulas(4) = "wFrToDt = 'From : ' & '" & Format(FromDt, "dd/MM/yyyy") & "' & '  To ' & '" & Format(ToDt, "dd/MM/yyyy") & "'"
  '=============== Account Details
 PrepareAccountStatement FromDt, ToDt, MillName, TDSP, ServiceECess, ServiceTax, Surcharge, TDSEcess
 If gBackEndDB = gBackEndAccess Then
     .DataFiles(0) = gDbLocation
 Else '-- Oracle
    .Connect = "DSN=" + gODBCDSN + " ;UID=Scott;PWD=" + gODBCPwd + ";DBQ=" + gODBCHostStr
 End If
 x1 = "{tmpGentbl.UserName}='" & gUserName & "'"
 .SelectionFormula = x1
 .WindowTitle = "Account Details"
 .ReportFileName = gReportPath & "rptMonSTMConsgDepotAccount.rpt"
 .Action = 1
 RptEnclosure(RptCount) = "Statement Of Accounts"
 RptCount = RptCount + 1
 '=============== Receipt Details
 If gBackEndDB = gBackEndAccess Then
     CondStr = " Vdt>= cdate('" & FromDt & "')  and  Vdt<=cdate('" & ToDt & "')"
 Else '-- Oracle
     CondStr = " Vdt>= to_date('" & CDate(FromDt) & "','dd/MM/yyyy')  and Vdt<=to_date('" & CDate(ToDt) & "','dd/MM/yyyy')"
 End If
 If GProcRstOpen(Rs, "select vno from tbladdless where " & CondStr & " and vtype='PI' and ADMILLCODE=" & MillCode, "R") > 0 Then
     RptEnclosure(RptCount) = "Receipt Details"
     RptCount = RptCount + 1
     x1 = "{tblAddLess.vdt} in Date(" & Format(FromDt, "yyyy,mm,dd") & ") to Date(" & Format(ToDt, "yyyy,mm,dd") & ")"
     x1 = x1 + " and {tblAddLess.VType}='PI' and {tblMastAccount.AcName}='" & MillName & "'"
     .SelectionFormula = x1
     .WindowTitle = "Receipt Details"
     .ReportFileName = gReportPath & "rptMonSTMConsgDepotReceipt.rpt"
     .Action = 1
 End If
 '=============== Sales Details
 If gBackEndDB = gBackEndAccess Then
     CondStr = " Vdt>= cdate('" & FromDt & "')  and  Vdt<=cdate('" & ToDt & "')"
 Else '-- Oracle
     CondStr = " Vdt>= to_date('" & CDate(FromDt) & "','dd/MM/yyyy')  and Vdt<=to_date('" & CDate(ToDt) & "','dd/MM/yyyy')"
 End If
 If GProcRstOpen(Rs, "select vno from tbladdless where " & CondStr & " and (vtype='SD' or vtype='SO') and ADMILLCODE=" & MillCode, "R") > 0 Then
     RptEnclosure(RptCount) = "Sales Details"
     RptCount = RptCount + 1
     x1 = "{tblAddLess.vdt} in Date(" & Format(FromDt, "yyyy,mm,dd") & ") to Date(" & Format(ToDt, "yyyy,mm,dd") & ")"
     'x1 = x1 + " and ({tblAddLess.VType}='SY' or {tblAddLess.VType}='ST' or {tblAddLess.VType}='SO' or {tblAddLess.VType}='SD' or {tblAddLess.VType}='SM') and {tblMastAccount.AcName}='" & txtMillName & "'"
     x1 = x1 + " and ({tblAddLess.VType}='SO' or {tblAddLess.VType}='SD' ) and {tblMastAccount.AcName}='" & MillName & "'"
     .SelectionFormula = x1
     .WindowTitle = "Sales Details"
     .ReportFileName = gReportPath & "rptMonSTMConsgDepotSale.rpt"
     .Action = 1
 End If
 '=============== Stock Details
 GprocPrepareCountStock FromDt, ToDt, MillName
 .Formulas(4) = "wFrToDt = 'From : ' & '" & FromDt & "' & '  To ' & '" & ToDt & "'"
 If gBackEndDB = gBackEndAccess Then
     .DataFiles(0) = gDbLocation
 Else '-- Oracle
    .Connect = "DSN=" + gODBCDSN + " ;UID=Scott;PWD=" + gODBCPwd + ";DBQ=" + gODBCHostStr
 End If
 x1 = "{tmpGentbl.UserName}='" & gUserName & "'"
 .SelectionFormula = x1
 .WindowTitle = "Stock Details"
 .ReportFileName = gReportPath & "rptMonSTMConsgDepotStock.rpt"
 .Action = 1
  RptEnclosure(RptCount) = "Stock Details"
  RptCount = RptCount + 1
  '===============Payment Details
 If gBackEndDB = gBackEndAccess Then
     CondStr = " Vdt>= cdate('" & FromDt & "')  and  Vdt<=cdate('" & ToDt & "')"
 Else '-- Oracle
     CondStr = " Vdt>= to_date('" & CDate(FromDt) & "','dd/MM/yyyy')  and Vdt<=to_date('" & CDate(ToDt) & "','dd/MM/yyyy')"
 End If
 If GProcRstOpen(Rs, "select vno from tblVoucher,tblMastNarration where MILLEXPNARRCODE=NarrCode  and " & CondStr & " and VType in ('BR','CR','JV','CP','BP','SN','PN') and vaccode=" & MillCode, "R") > 0 Then
     RptEnclosure(RptCount) = "Payment Details"
     RptCount = RptCount + 1
     x1 = "{tblVoucher.vdt} in Date(" & Format(FromDt, "yyyy,mm,dd") & ") to Date(" & Format(ToDt, "yyyy,mm,dd") & ")"
     x1 = x1 + " and ({tblVoucher.VType}='BR' or {tblVoucher.VType}='CR' or {tblVoucher.VType}='JV' or {tblVoucher.VType}='CP' or {tblVoucher.VType}='BP' or {tblVoucher.VType}='SN' or {tblVoucher.VType}='PN') and {tblMastAccount.AcName}='" & MillName & "'"
     .SelectionFormula = x1
     .WindowTitle = "Payment Details"
     .ReportFileName = gReportPath & "rptMonSTMConsgDepotPayment.rpt"
     .Action = 1
 End If
 '=============== Debit Note Commission
 If CDbl(CommAmt) > 0 Then
     CommType = GProcGetCommissionType(GProcGetColumnValue("TblMastAccount", "AcName", MillName, "S", "AcComType", "N"))
     RptEnclosure(RptCount) = CommType & " Debit Note"
     RptCount = RptCount + 1
     x1 = "{tblMastAccount.AcName}='" & MillName & "'"
     .Formulas(11) = "WAmountInWords = '( " & GProcAmountInWords(CDbl(CommAmt)) & ")'"
     .Formulas(12) = "CommAmt = '" & CDbl(CommAmt) & "'"
     .Formulas(13) = "CommOnAmt = '" & CDbl(CommOn) & "'"
     .Formulas(14) = "CommBag = '" & CLng(CommBag) & "'"
     .Formulas(15) = "CommType = '" & CommType & "'"
     .SelectionFormula = x1
     .WindowTitle = "Debit Note ( " & CommType & " )"
     .ReportFileName = gReportPath & "rptMonSTMConsgDepotDBNTComm.rpt"
     .Action = 1
     .Formulas(11) = ""
     .Formulas(12) = ""
     .Formulas(13) = ""
     .Formulas(14) = ""
     .Formulas(15) = ""
 End If
 Amt = 0
 '----------- commission On ( Commission On Ex Mill Sales)
 'If GProcGetColumnValue("tblmastaccount", "Acname", txtMillName, "S", "ACCOMON", "N", "  ACCOMON=3") > 0 Then
 PrepareRptCommissionReg FromDt, ToDt, MillCode
 If GProcRstOpen(Rs, "select * from tmpgentbl where Username='" & gUserName & "'", "R") > 0 Then
     .SelectionFormula = "{tmpGenTbl.UserName}='" & gUserName & "'"
     .WindowTitle = "Invoicewise Ex- Mill Amount  Register"
     .ReportFileName = gReportPath & "rptMonSTMConsgDepotCommReg.rpt"
     .Action = 1
 End If
  '----------- commission On ( Commission On Ex Mill Inward)
 If gBackEndDB = gBackEndAccess Then
     CondStr = " Vdt>= cdate('" & FromDt & "')  and  Vdt<=cdate('" & ToDt & "')"
 Else '-- Oracle
     CondStr = " Vdt>= to_date('" & CDate(FromDt) & "','dd/MM/yyyy')  and Vdt<=to_date('" & CDate(ToDt) & "','dd/MM/yyyy')"
 End If
x1 = " select sum(ADSUBAMT),AcComRt  as CommAmt from tblAddLess,tblMastAccount where " _
& " tblAddLess.AdMillCode=tblMastAccount.AcCode and tblAddLess.VType in ('PI')  and " & CondStr & " and ADMILLCODE=" & MillCode _
& " and AcComRt>0 and AcComOn=1 group by AcComRt"
 If GProcRstOpen(Rs, x1, "R") > 0 Then
    If IsNumeric(Rs.Fields(0)) = True Then
        x1 = "{tblAddLess.vdt} in Date(" & Format(FromDt, "yyyy,mm,dd") & ") to Date(" & Format(ToDt, "yyyy,mm,dd") & ")"
        x1 = x1 + " and {tblAddLess.VType}='PI' and {tblMastAccount.AcComRt}>0 and {tblMastAccount.AcComOn}=1 and {tblMastAccount.AcName}='" & MillName & "'"
        .SelectionFormula = x1
        .WindowTitle = "Inwardwise Ex- Mill Amount  Register"
        .ReportFileName = gReportPath & "rptMonSTMConsgDepotCommReg_ExMillIn.rpt"
        .Action = 1
    End If
 End If
 '=============== Service tax Bill from A/c details
   PrepareAccountStatement FromDt, ToDt, MillName, TDSP, ServiceECess, ServiceTax, Surcharge, TDSEcess
  If GProcRstOpen(Rs, "select Amt1,Amt2 from tmpgentbl where Username='" & gUserName & "' and vno=3", "R") > 0 Then
     .Formulas(11) = "SaleAmt='" & Format(Rs.Fields(1), FStr) & "'"  '-- Comm On
     .Formulas(12) = "CommAmt='" & Format(Abs(Rs.Fields(0)), FStr) & "'" '-- Comm Amt
     Amt = Abs(CDbl(Rs.Fields(0)))
  End If
  If GProcRstOpen(Rs, "select Amt1,Amt2,Amt3 from tmpgentbl where Username='" & gUserName & "' and vno=7", "R") > 0 Then
     .Formulas(14) = "TotServiceTax='" & Format(Abs(Rs.Fields(0)), FStr) & "'"
     .Formulas(13) = "ServiceTaxAmt='" & Format(Abs(Rs.Fields(1)), FStr) & "'"
     .Formulas(15) = "CessAmt='" & Format(Abs(Rs.Fields(2)), FStr) & "'"
     Amt = Amt + Abs(CDbl(Rs.Fields(0)))
  End If
  If Amt > 0 Then
     .Formulas(4) = "wFrToDt = 'From : ' & '" & FromDt & "' & '  To ' & '" & ToDt & "'"
     .Formulas(16) = "CessP='" & ServiceECess & "'"
     .Formulas(17) = "TaxP='" & ServiceTax & "'"
     .Formulas(18) = "WAmountInWords='" & GProcAmountInWords(Amt) & "'"
     .Formulas(19) = "BillNo='" & TaxBillNo & "'"
     .Formulas(20) = "BillDt='" & TaxBilldt & "'"
     .Formulas(21) = "STRegNo='" & gCSTRegNo & "'"
     .Formulas(22) = "STCNo='" & gCSTCNo & "'"
     
     If gBackEndDB = gBackEndAccess Then
         .DataFiles(0) = gDbLocation
     Else '-- Oracle
        .Connect = "DSN=" + gODBCDSN + " ;UID=Scott;PWD=" + gODBCPwd + ";DBQ=" + gODBCHostStr
     End If
     x1 = "{tmpGentbl.UserName}='" & gUserName & "'"
     .SelectionFormula = x1
     .WindowTitle = "Service Tax Bill"
     .ReportFileName = gReportPath & "rptMonSTMConsgDepotTaxBill.rpt"
     RptEnclosure(RptCount) = "Service Tax Bill"
     RptCount = RptCount + 1
     .Action = 1
End If
For i = 10 To 30
     .Formulas(i) = ""
Next i
  ' =============== Debit Note Brokerage
 If gBackEndDB = gBackEndAccess Then
     CondStr = " tblAddLess.Vdt>= cdate('" & FromDt & "')  and  tblAddLess.Vdt<=cdate('" & ToDt & "')"
 Else '-- Oracle
     CondStr = " tblAddLess.Vdt>= to_date('" & CDate(FromDt) & "','dd/MM/yyyy')  and tblAddLess.Vdt<=to_date('" & CDate(ToDt) & "','dd/MM/yyyy')"
 End If
 If GProcRstOpen(Rs, "select tblAddLess.vno from tblAddLess,tblVousub,tblMastAccount,tblMastItem where AcCode=AdMillcode " _
 & " and tblAddLess.Vtype=tblVousub.Vtype and  tblAddLess.Vno=tblVousub.vno and tblAddLess.vyear=tblVousub.VYear " _
 & " and itCode= VSUBITCODE  and " & CondStr & " and tblVousub.vtype in ('SO','SD') and ITBROKRT>0  and ACISLOCALBRKBYSELF=1 and adMillCode=" & MillCode, "R") > 0 Then
     RptEnclosure(RptCount) = "Brokerage Debit Note"
     RptCount = RptCount + 1
     x1 = "{tblAddLess.vdt} in Date(" & Format(FromDt, "yyyy,mm,dd") & ") to Date(" & Format(ToDt, "yyyy,mm,dd") & ")"
     x1 = x1 + " and ({tblAddLess.VType}='SO' or {tblAddLess.VType}='SD' ) and {tblMastItem.ITBROKRT}>0 and {tblMastAccount.AcName}='" & MillName & "'"
     .SelectionFormula = x1
     .WindowTitle = "Debit Note (Brokerage)"
     .ReportFileName = gReportPath & "rptMonSTMConsgDepotDBNTBrk.rpt"
     .Action = 1
 End If
 Amt = 0
'========================= Enclosures
 .Formulas(4) = "wFrToDt = 'From : ' & '" & FromDt & "' & '  To ' & '" & ToDt & "'"
 If gBackEndDB = gBackEndAccess Then
     .DataFiles(0) = gDbLocation
 Else '-- Oracle
    .Connect = "DSN=" + gODBCDSN + " ;UID=Scott;PWD=" + gODBCPwd + ";DBQ=" + gODBCHostStr
 End If
 x1 = "{tmpGentbl.UserName}='" & gUserName & "'"
 .SelectionFormula = x1
 .WindowTitle = "Cosignment/ Depot Statement"
 .ReportFileName = gReportPath & "rptMonSTMConsgDepot.rpt"
 For i = 0 To RptCount
     If RptEnclosure(i) <> "" Then .Formulas(i + 10) = "Str" & i & " ='" & i + 1 & " ) " & RptEnclosure(i) & "'"
 Next i
 .Action = 1
 End With
End Sub
'-------------------- Prepare Mill Account Statement ( used for Consignment Depot Mill statement)
Private Sub PrepareAccountStatement(pFrmDt As Date, pToDt As Date, MillName As String, TDSP As Double, ServiceECess As Double, ServiceTax As Double, Surcharge As Double, TDSEcess As Double)
'''Dim RsTmp As Recordset
'''Dim X As String
'''Dim VATTaxCondStr As String
'''Dim MillCode As Long
'''Dim Amt As Double
'''Dim TaxAmt As Double
'''Dim BrkBag As Double
'''Dim BrkWt As Double
'''Dim BrkAmt As Double
'''Dim TotAmt As Double
'''Dim TDSAmt As Double
'''Dim SurchargeAmt As Double
'''Dim TDSCessAmt As Double
'''Dim CessAmt As Double
'''Dim ServiceTaxAmt As Double
'''Dim CondStr As String
'''MillCode = GProcGetColumnValue("tblMastAccount", "AcName", MillName, "S", "AcCode", "N")
'''gCn.BeginTrans
'''gCn.Execute "Delete from tmpGenTbl where USERNAME='" & gUserName & "'"
''''============== Opening Balance
'''Amt = GProcGetAccountBalance(MillName, pFrmDt - 1)
'''gCn.Execute "insert into tmpGenTbl (VAcCode,Nar1,UserName,Amt1,Vno,VSrno) values (" & MillCode & "," _
'''& "'Opening Balance','" & gUserName & "'," & Amt & ",1,0)"
''''============== Sales
'''Amt = 0
'''If gBackEndDB = gBackEndAccess Then
'''    CondStr = " Vdt>= cdate('" & pFrmDt & "')  and  Vdt<=cdate('" & pToDt & "')"
'''Else '-- Oracle
'''    CondStr = " Vdt>= to_date('" & CDate(pFrmDt) & "','dd/MM/yyyy')  and Vdt<=to_date('" & CDate(pToDt) & "','dd/MM/yyyy')"
'''End If
'''X = " select sum(ADBILLAMT) from tblAddLess where VType in ('SO','SD')  and " & CondStr & " and  ADMILLCODE=" & MillCode
'''If GProcRstOpen(RsTmp, X, "R") > 0 Then
'''    If IsNumeric(RsTmp.Fields(0)) = True Then
'''        Amt = CDbl(RsTmp.Fields(0))
'''    Else
'''        Amt = 0
'''    End If
'''    If Amt > 0 Then gCn.Execute "insert into tmpGenTbl (VAcCode,Nar1,UserName,Amt1,Vno,VSrno) values (" & MillCode & "," _
'''    & "'Sales','" & gUserName & "'," & Amt & ",2,0)"
'''End If
''''============== Commission,V A T Payable,TDS ,E Cess ,Service tax
'''Amt = 0
'''CommAmt = 0
'''CommOn = 0
'''CommBag = 0
'''If gBackEndDB = gBackEndAccess Then
'''    CondStr = " tblAddLess.Vdt>= cdate('" & pFrmDt & "')  and  tblAddLess.Vdt<=cdate('" & pToDt & "')"
'''Else '-- Oracle
'''    CondStr = " tblAddLess.Vdt>= to_date('" & CDate(pFrmDt) & "','dd/MM/yyyy')  and tblAddLess.Vdt<=to_date('" & CDate(pToDt) & "','dd/MM/yyyy')"
'''End If
'''If gBackEndDB = gBackEndAccess Then '-- Access
'''    If gIsNoVATVoucherSD = False And gIsNoVATVoucherSO = False Then
'''        VATTaxCondStr = "sum(AdTaxAmt)"
'''    Else
'''        If gIsNoVATVoucherSD = True Then
'''            VATTaxCondStr = "sum(iif(tblAddLess.Vtype='SD',iif(NARRATION='V A T',0,AdTaxAmt),AdTaxamt))"
'''        Else
'''            If gIsNoVATVoucherSO = True Then
'''                VATTaxCondStr = "sum(iif(tblAddLess.Vtype='SO',iif(NARRATION='V A T',0,AdTaxAmt),AdTaxamt))"
'''            Else
'''                VATTaxCondStr = "sum(iif(NARRATION='V A T',0,AdTaxAmt))"
'''            End If
'''        End If
'''    End If
'''Else '-- Oracle
'''     If gIsNoVATVoucherSD = False And gIsNoVATVoucherSO = False Then
'''        VATTaxCondStr = "sum(AdTaxAmt)"
'''     Else
'''        If gIsNoVATVoucherSD = True Then
'''            VATTaxCondStr = "sum(decode(tblAddLess.Vtype,'SD',decode(NARRATION,'V A T',0,AdTaxAmt),AdTaxamt))"
'''        Else
'''            If gIsNoVATVoucherSO = True Then
'''                VATTaxCondStr = "sum(decode(tblAddLess.Vtype,'SO',decode(NARRATION,'V A T',0,AdTaxAmt),AdTaxamt))"
'''            Else
'''                VATTaxCondStr = "sum(decode(NARRATION,'V A T',0,AdTaxAmt))"
'''            End If
'''        End If
'''    End If
'''End If
'''X = " select " & VATTaxCondStr & " as TaxAmt,sum(ADBILLAMT) as TotAmt,sum(VSubBag) as CommBag,AcComOn,AcComRt,ACCOMTYPE from " _
'''& " tblAddLess,tblMastAccount,tblVouSub,tblMastNarration where ADTAXCODE=NARRCODE and " _
'''& " tblAddLess.Vno=tblVouSub.Vno and tblAddLess.VType=tblVouSub.Vtype and tblAddLess.VYear=tblVouSub.VYear and " _
'''& " tblAddLess.AdMillCode=tblMastAccount.AcCode and tblAddLess.VType in ('SO','SD')  and " & CondStr & " and ADMILLCODE=" & MillCode _
'''& " Group by AcComOn,AcComRt,ACCOMTYPE" 'and AcComOn in (0,2)
'''If GProcRstOpen(RsTmp, X, "R") > 0 Then
'''    If IsNumeric(RsTmp!TaxAmt) = True Then
'''        TaxAmt = CDbl(RsTmp!TaxAmt)
'''    Else
'''        TaxAmt = 0
'''    End If
'''    If IsNumeric(RsTmp!TotAmt) = True Then
'''        TotAmt = CDbl(RsTmp!TotAmt)
'''    Else
'''        TotAmt = 0
'''    End If
'''    If IsNumeric(RsTmp!CommBag) = True Then
'''        CommBag = CLng(RsTmp!CommBag)
'''    Else
'''       CommBag = 0
'''    End If
'''    If RsTmp!ACCOMON = 0 Then '--Gross amt
'''        CommAmt = TotAmt * CDbl(RsTmp!AcComRt) / 100
'''        CommOn = TotAmt
'''    ElseIf RsTmp!ACCOMON = 2 Then '-- Net Amt
'''        CommAmt = (TotAmt - TaxAmt) * CDbl(RsTmp!AcComRt) / 100
'''        CommOn = CDbl(TotAmt) - CDbl(TaxAmt)
'''    ElseIf RsTmp!ACCOMON = 1 Then  '---- Commission (On Ex Mill Amt (inward))
'''        X = " select " & IIf(gBackEndDB = gBackEndAccess, "sum(iif(tblVouSub.VSUBCTRNO=1,ADSUBAMT,0))", "sum(decode(tblVouSub.VSUBCTRNO,1,ADSUBAMT,0))") & " as CommAmt,sum(VSubBag) as CommBag,ACCOMRT,ACCOMTYPE from tblAddLess,tblMastAccount,tblVouSub where " _
'''        & " tblAddLess.Vno=tblVouSub.Vno and tblAddLess.VType=tblVouSub.Vtype and tblAddLess.VYear=tblVouSub.VYear and " _
'''        & " tblAddLess.AdMillCode=tblMastAccount.AcCode and tblAddLess.VType in ('PI')  and " & CondStr & " and ADMILLCODE=" & MillCode _
'''        & " and AcComRt>0 and AcComOn=1 group by ACCOMRT,ACCOMTYPE"
'''        If GProcRstOpen(RsTmp, X, "R") > 0 Then
'''            If IsNumeric(RsTmp!CommAmt) = True Then
'''                CommAmt = GProcMakeRounding(CDbl(RsTmp!CommAmt) * CDbl(RsTmp!AcComRt) / 100)
'''                CommOn = CDbl(RsTmp!CommAmt)
'''                CommBag = CLng(RsTmp!CommBag)
'''            End If
'''        End If
'''     Else   '---- Commission (On Ex Mill Amt (sales))  sum( Sale Wt * (Purch ExMill Amt/Total Purch Wt) )
'''        X = " select sum(tblVouSub.VSubWt*tbladdLess_PU.AdSubamt/( select sum(tblVouSub_PU1.VSubWt) from tblVouSub tblVouSub_PU1 " _
'''        & " where tblVouSub_PU1.Vno=tblAddLess_PU.Vno and tblVouSub_PU1.Vtype=tblAddLess_PU.VType and tblVouSub_PU1.VYear=tblAddLess_PU.VYear)) " _
'''        & " as CommAmt,sum(tblVouSub.VSubBag) as CommBag,ACCOMRT,ACCOMTYPE " _
'''        & " from tblAddLess,tblMastAccount,tblAddLess tblAddLess_PU,tblVouSub,tblVouSub tblVouSub_PU,tblVouSub TblVouSub_GP where " _
'''        & " tblAddLess.Vno=tblVouSub.Vno and tblAddLess.VType=tblVouSub.Vtype and tblAddLess.VYear=tblVouSub.VYear and " _
'''        & " tblAddLess_PU.Vno=tblVouSub_PU.Vno and tblAddLess_PU.VType=tblVouSub_PU.Vtype and tblAddLess_PU.VYear=tblVouSub_PU.VYear  and tblAddLess_PU.Vtype='PI' and " _
'''        & " tblVouSub_GP.VSUBPITCTRLNO=tblVouSub_PU.VSUBITCTRLNO and tblVouSub_GP.Vdt=tblAddLess.AdGpDt and tblVouSub_GP.Vno=tblAddLess.AdGpNo and " _
'''        & " tblAddLess.AdMillCode=tblMastAccount.AcCode and tblAddLess.VType in ('SO','SD')  and " & CondStr & " and tblAddLess.ADMILLCODE=" & MillCode _
'''        & " and AcComRt>0 and AcComOn=3 group by ACComRt,ACCOMTYPE"
'''        If GProcRstOpen(RsTmp, X, "R") > 0 Then
'''            If IsNumeric(RsTmp!CommAmt) = True Then
'''                CommAmt = GProcMakeRounding(CDbl(RsTmp!CommAmt) * CDbl(RsTmp!AcComRt) / 100)
'''                CommOn = CDbl(RsTmp!CommAmt)
'''                CommBag = CLng(RsTmp!CommBag)
'''            End If
'''        End If
'''    End If
'''    CommOn = GProcMakeRounding(CommOn)
'''    CommAmt = GProcMakeRounding(CommAmt)
'''    Amt = CommAmt
'''    '-- Commission
'''    If Amt > 0 Then gCn.Execute "insert into tmpGenTbl (VAcCode,Nar1,UserName,Amt1,Amt2,Amt3,Vno,VSrno) values (" & MillCode & "," _
'''    & "'" & GProcGetCommissionType(RsTmp!ACCOMTYPE) & " (@ " & RsTmp!AcComRt & " On Amount " & CommOn & " )','" & gUserName & "',-" & Amt & "," & CommOn & "," & CommBag & ",3,0)"
'''    '-- V A T Payable
'''    If TaxAmt > 0 Then gCn.Execute "insert into tmpGenTbl (VAcCode,Nar1,UserName,Amt1,Vno,VSrno) values (" & MillCode & "," _
'''    & "'V A T Payable','" & gUserName & "',-" & TaxAmt & ",4,0)"
'''    '--- TDS & cess
'''    TDSAmt = CDbl((Amt * CDbl(TDSP) / 100))
'''    TDSAmt = GProcMakeRounding(TDSAmt)
'''    SurchargeAmt = CDbl(TDSAmt * CDbl(Surcharge) / 100)
'''    SurchargeAmt = GProcMakeRounding(SurchargeAmt)
'''    TDSCessAmt = CDbl(TDSAmt * CDbl(TDSEcess) / 100)
'''    TDSCessAmt = GProcMakeRounding(TDSCessAmt)
'''    If (CDbl(TDSAmt) + CDbl(SurchargeAmt) + CDbl(TDSCessAmt)) > 0 Then gCn.Execute "insert into tmpGenTbl (VAcCode,Nar1,UserName,Amt1,Vno,VSrno) values (" & MillCode & "," _
'''    & "'TDS @ " & TDSP & " , Surcharge @ " & Surcharge & " & Ed. Cess @ " & TDSEcess & "' ,'" & gUserName & "',-" & CDbl(TDSAmt) + CDbl(SurchargeAmt) + CDbl(TDSCessAmt) & ",6,0)"
'''    '--- Service Tax & E Cess
'''    ServiceTaxAmt = CDbl((Amt * CDbl(ServiceTax) / 100))
'''    ServiceTaxAmt = GProcMakeRounding(ServiceTaxAmt)
'''    CessAmt = CDbl(ServiceTaxAmt * CDbl(ServiceECess) / 100)
'''    CessAmt = GProcMakeRounding(CessAmt)
'''    If CDbl(ServiceTaxAmt + CessAmt) > 0 Then gCn.Execute "insert into tmpGenTbl (VAcCode,Nar1,UserName,Amt1,amt2,amt3,Vno,VSrno) values (" & MillCode & "," _
'''    & "'Service Tax @ " & ServiceTax & " & Ed. Cess @ " & ServiceECess & "','" & gUserName & "',-" & CDbl(ServiceTaxAmt + CessAmt) & "," & ServiceTaxAmt & "," & CessAmt & ",7,0)"
'''End If
''''==================== Brokerage
'''Amt = 0
'''If gBackEndDB = gBackEndAccess Then
'''    CondStr = " tblAddLess.Vdt>= cdate('" & pFrmDt & "')  and  tblAddLess.Vdt<=cdate('" & pToDt & "')"
'''    X = " select sum(iif(ITBROKON='Bag',VSUBBAG,iif(ITBROKON='Kg',VSUBWT/100,ADBILLAMT/100)) * ITBROKRT) as BrkAmt from tblAddLess,tblMastAccount,tblMastItem,TblVouSub where " _
'''    & " tblAddLess.Vtype=TblVouSub.Vtype and  tblAddLess.Vno=TblVouSub.Vno and  tblAddLess.VYear=TblVouSub.VYear and  " _
'''    & " AdMillCode=AcCode  and ItCode=VSUBITCODE and tblAddLess.VType in ('SO','SD')  and " & CondStr & " and ADMILLCODE=" & MillCode _
'''    & " and  ACISLOCALBRKBYSELF=1"
'''Else '-- Oracle
'''    CondStr = " tblAddLess.Vdt>= to_date('" & CDate(pFrmDt) & "','dd/MM/yyyy')  and tblAddLess.Vdt<=to_date('" & CDate(pToDt) & "','dd/MM/yyyy')"
'''    X = " select sum(decode(ITBROKON,'Bag',VSUBBAG,'Kg',VSUBWT/100,ADBILLAMT/100) * ITBROKRT) as BrkAmt from tblAddLess,tblMastAccount,tblMastItem,TblVouSub where " _
'''    & " tblAddLess.Vtype=TblVouSub.Vtype and  tblAddLess.Vno=TblVouSub.Vno and  tblAddLess.VYear=TblVouSub.VYear and  " _
'''    & " AdMillCode=AcCode  and ItCode=VSUBITCODE and tblAddLess.VType in ('SO','SD')  and " & CondStr & " and ADMILLCODE=" & MillCode _
'''    & " and  ACISLOCALBRKBYSELF=1"
'''End If
''''& ""
'''If GProcRstOpen(RsTmp, X, "R") > 0 Then
'''    If IsNull(RsTmp!BrkAmt) = False Then
'''        Amt = CDbl(RsTmp!BrkAmt)
'''    Else
'''        Amt = 0
'''    End If
'''    Amt = GProcMakeRounding(Amt)
'''End If
'''    If Amt > 0 Then gCn.Execute "insert into tmpGenTbl (VAcCode,Nar1,UserName,Amt1,Vno,VSrno) values (" & MillCode & "," _
'''    & "'Brokerage','" & gUserName & "',-" & Amt & ",5,0)"
''''==================== Other A/c
'''If gBackEndDB = gBackEndAccess Then
'''    CondStr = " Vdt>= cdate('" & pFrmDt & "')  and  Vdt<=cdate('" & pToDt & "')"
'''    gCn.Execute "insert into tmpGenTbl (VAcCode,Nar1,UserName,Amt1,Vno,VSrno)   select " & MillCode & "," _
'''    & "Narration,'" & gUserName & "',sum(iif(VACcode=" & MillCode & ",vamt,-Vamt)),8,0 from tblvoucher,tblMastNarration where VType in ('BR','CR','JV','CP','BP','SN','PN')  and " & CondStr & " and " _
'''    & " (vAcCode=" & MillCode & " or VAcOcode=" & MillCode & ")" _
'''    & " and MILLEXPNARRCODE=NarrCode and VCTRNO=1 group by Narration order by Narration"
'''Else '-- Oracle
'''    CondStr = " Vdt>= to_date('" & CDate(pFrmDt) & "','dd/MM/yyyy')  and Vdt<=to_date('" & CDate(pToDt) & "','dd/MM/yyyy')"
'''    gCn.Execute "insert into tmpGenTbl (VAcCode,Nar1,UserName,Amt1,Vno,VSrno)   select " & MillCode & "," _
'''    & "Narration,'" & gUserName & "',sum(decode(VACcode," & MillCode & ",vamt,-Vamt)),8,0 from tblvoucher,tblMastNarration where VType in ('BR','CR','JV','CP','BP','SN','PN')  and " & CondStr & " and " _
'''    & " (vAcCode=" & MillCode & " or VAcOcode=" & MillCode & ")" _
'''    & " and MILLEXPNARRCODE=NarrCode and VCTRNO=1 group by Narration order by Narration"
'''End If
'''gCn.CommitTrans
End Sub
'------------------Prepare Mill Commission Register ( used for Consignment Depot Mill statement)
Private Sub PrepareRptCommissionReg(FromDt As Date, ToDt As Date, MillCode As Long)
Dim CondStr As String
If gBackEndDB = gBackEndAccess Then
    CondStr = " tblAddLess.Vdt>= cdate('" & FromDt & "')  and  tblAddLess.Vdt<=cdate('" & ToDt & "')"
Else '-- Oracle
    CondStr = " tblAddLess.Vdt>= to_date('" & FromDt & "','dd/MM/yyyy')  and tblAddLess.Vdt<=to_date('" & CDate(ToDt) & "','dd/MM/yyyy')"
End If
gCn.Execute "delete from tmpGenTbl where Username='" & gUserName & "'"
 gCn.Execute "insert into tmpGenTbl (VAcCode,VBillNo,Vdt,Amt1,Amt2,Nar1,Nar2,amt3,Amt4,UserName) select " & MillCode & "," _
    & " tblAddLess.AdInvNo,tblAddLess.Vdt,tblVouSub.VSUBBAG,tblVouSub.VSUBWT,tblAddLess_PU.Vno,tblAddLess_PU.ADGPNO,tblAddLess_PU.ADSubAMT, " _
    & " (select sum(tblVouSub_PU1.VSubWt) from tblVouSub tblVouSub_PU1 " _
        & " where tblVouSub_PU1.Vno=tblAddLess_PU.Vno and tblVouSub_PU1.Vtype=tblAddLess_PU.VType and tblVouSub_PU1.VYear=tblAddLess_PU.VYear),'" & gUserName & "'" _
        & " from tblAddLess,tblMastAccount,tblAddLess tblAddLess_PU,tblVouSub,tblVouSub tblVouSub_PU,tblVouSub TblVouSub_GP where " _
        & " tblAddLess.Vno=tblVouSub.Vno and tblAddLess.VType=tblVouSub.Vtype and tblAddLess.VYear=tblVouSub.VYear and " _
        & " tblAddLess_PU.Vno=tblVouSub_PU.Vno and tblAddLess_PU.VType=tblVouSub_PU.Vtype and tblAddLess_PU.VYear=tblVouSub_PU.VYear  and tblAddLess_PU.Vtype='PI' and " _
        & " tblVouSub_GP.VSUBPITCTRLNO=tblVouSub_PU.VSUBITCTRLNO and tblVouSub_GP.Vdt=tblAddLess.AdGpDt and tblVouSub_GP.Vno=tblAddLess.AdGpNo and " _
        & " tblAddLess.AdMillCode=tblMastAccount.AcCode and tblAddLess.VType in ('SO','SD')  and " & CondStr & " and tblAddLess.ADMILLCODE=" & MillCode _
        & " and AcComRt>0 and AcComOn=3 "
End Sub
'=== Prepare TmpGenTbl For Party Booking Vs Sales
Public Sub GProcPrepareTmpTblForPartyBkVsSA(FromDt As Date, ToDt As Date)
Dim RsTmp As Recordset
Dim X As String
'--- Prepare TmpGenTbl For Party Booking Register
GProcPrepareTmpTblForPartyBkReg FromDt, ToDt
'--- Insert Booking Vs Deatils - GataPass No ,Inv. No,Date,Bag,Wt,Company Code
'gSetupCn.BeginTrans
gSetupCn.Execute "delete from  tmpRptTbl where UserName='" & gUserName & "'"
If GProcRstOpen(RsTmp, "Select Vno from tmpGenTbl where Username='" & gUserName & "' order by Vno", "R", gSetupCn) > 0 Then
    With RsTmp
        .MoveFirst
        Do While Not .EOF
            '--- insert Related Sale Details to setup db Table For All Firm
            X = " Select '" & gUserName & "'," & .Fields(0) & ",TblGp.Vno,TblGp.Vdt,TblAddLess.AdInvNo,TblAddLess.VDt,TblGp.VSubBag,TblGp.VSubWt from TblVouSub TblGp,TblAddLess,TblVouSub" _
            & " Where TblAddLess.Vno=TblVouSub.Vno and TblAddLess.VType=TblVouSub.VType and TblAddLess.VYear=TblVouSub.VYear and TblGp.Vtype='GP' and TblVouSub.VSUBPITCTRLNO=TblGp.VSUBITCTRLNO" _
            & " and  TblGp.VSUBBKITCTRLNO=" & .Fields(0)
            GProcInsertToSetupDbTbl True, X, "tmpRptTbl", "username,Vno,Amt4,Dt1,Vbillno,VBillDt,Amt2,Amt3", " UserName='" & gUserName & "' and Vno=" & .Fields(0), "Nar5"
            .MoveNext
        Loop
    End With
 End If
 RsTmp.Close
End Sub
'=== Prepare TmpGenTbl For Party Booking Register
Public Sub GProcPrepareTmpTblForPartyBkReg(FromDt As Date, ToDt As Date)
Dim RsTmp As Recordset
Dim BrkName As String '- Broker
Dim PartyName As String '-Party
Dim ItName As String '-Item
Dim MillCode As Long '-MillCode
Dim MillName As String '--Mill
    gSetupCn.BeginTrans
    gSetupCn.Execute "delete from tmpGenTbl where username='" & gUserName & "'"
    '-- TmpGenTbl insert          Vno=Booking ItCtrl No
    '- VAcCode=Booking Broker Code,  VAcOcode=Booking Party Code,   Amt1=Item Code
     gSetupCn.Execute "insert Into TmpGenTbl (vno,VAcCode,VAcOCode,Amt1,UserName) select " _
    & " BkSubItCtrlNo,BkAcBrkCode,BKAcCode,BkSubItCode,'" & gUserName & "'" _
    & " from tblBooking,tblBookingSub where tblBooking.Vno=tblBookingSub.Vno and tblBooking.VType=tblBookingSub.VType " _
    & "  and tblBooking.VYear=tblBookingSub.VYear  and tblBooking.Vdt>=Cdate('" & FromDt & "') and tblBooking.Vdt<=Cdate('" & ToDt & "') and BkCCode='" & gCCode & "' and BkCYear='" & gDbYr & "'"
    '-- TmpGenTbl Set         Nar1=Booking Broker ,         Nar2=Party,        Nar3=Item,      Nar4=Mill
    GProcRstOpen RsTmp, "Select * from tmpGenTbl where Username='" & gUserName & "'", "O", gSetupCn
 With RsTmp
    Do While Not .EOF
        BrkName = GProcGetColumnValue("TblMastAccount", "ACCode", !VAcCode, "N", "AcName", "S", "")   '-Broker
        PartyName = GProcGetColumnValue("TblMastAccount", "ACCode", !VAcOcode, "N", "AcName", "S", "") '-Party
        ItName = GProcGetColumnValue("TblMastItem", "ItCode", !Amt1, "N", "ItName", "S", "") '-Item
        MillCode = GProcGetColumnValue("TblMastAccount,TblMastItem", "ItCode", !Amt1, "N", "AcCode", "N", "AcCode=ITMILLCODE")   '-Mill Code
        MillName = GProcGetColumnValue("TblMastAccount", "ACCode", CStr(MillCode), "N", "AcName", "S", "")    '- mill
        !Nar1 = BrkName 'Broker
        !Nar2 = PartyName 'Party
        !Nar3 = ItName 'Item
        !Nar4 = MillName 'Mill
        .Update
        .MoveNext
    Loop
   RsTmp.Close
 End With
 gSetupCn.CommitTrans
End Sub
'========= Prepare Misssing Transanction Report
'--- I/p - FromDt,ToDate
'-- O/p - TmpGenTbl (Insert Missing Transaction) Vno - >vno,Vtype -> Vtype,Max(Vno) ->Nar1 ,Account(Sale)-> Nar2 For Current User
Public Sub GProcPrepareMisssingTransRpt(FromDt As Date, ToDt As Date, VTypeStr As String)
Dim rsVou As Recordset
Dim X As String
Dim CondStr As String '-  Condition String
Dim RCount As Long '- Record Counter
Dim MaxRNo As Long '-- Max Record No.
Dim MissInVNo As String
Select Case VTypeStr
    Case "BK" '-Booking
        CondStr = "vtype in ('OS','OP','OI')"
    Case "VU"  '- Vouchers
        CondStr = " Vtype not in ('0P','SY','ST','SO','SD','PY','PT','PO')"
    Case "SA" '-  Sales
        CondStr = " Vtype in ('SY','ST','SO','SD')"
    Case "GP" '- Gate Pass
        CondStr = "Vtype = 'GP'"
    Case "PU" '- Purchase
        CondStr = " Vtype in ('PY','PT','PO','PI')"
End Select
'--Delete from Tmp
gCn.Execute "delete from  tmpGenTbl where UserName='" & gUserName & "'"
If gBackEndDB = gBackEndAccess Then
    CondStr = CondStr & " and vdt>=cdate('" & FromDt & "') and vdt<=cdate('" & ToDt & "')"
Else
    CondStr = CondStr & " and vdt>=to_date('" & FromDt & "','dd/MM/yyyy') and vdt<=To_date('" & ToDt & "','dd/MM/yyyy')"
End If
'-- Insert To Tmp A/c Voucher
If VTypeStr = "VU" Then
    X = "select max(vno) as MaxVNo,Vtype from tblVoucher where VCTRNO=1 and " & CondStr & " Group by VType having count(*)<>Max(vno)"
    If GProcRstOpen(rsVou, X, "R") > 0 Then
        With rsVou
            .MoveFirst
            Do While Not .EOF
                RCount = 1
                MaxRNo = IIf(IsNumeric(.Fields("MaxVNo")), .Fields("MaxVNo"), 0)
                Do While RCount <= MaxRNo
                    '-- Transaction Missing
                    If GProcGetColumnValue("TblVoucher", "Vno", CStr(RCount), "N", "Vno", "N", " VCTRNO=1 and Vtype='" & .Fields("Vtype") & "' and " & CondStr) = 0 Then
                        gCn.Execute "insert into tmpGenTbl (VBillNo,Vno,Vtype,Nar1,UserName) Values (" & RCount & "," & RCount & ",'" & .Fields("VTYpe") & "'," & MaxRNo & ",'" & gUserName & "')"
                    End If
                    RCount = RCount + 1
                Loop
               .MoveNext
            Loop
        End With
    End If
End If
'-- Insert To Tmp Missing GatePass,Booking,Purchase
If VTypeStr = "GP" Or VTypeStr = "BK" Or VTypeStr = "PU" Then
    X = "select max(vno) as MaxVNo,Vtype from tblAddLess where " & CondStr & " Group by VType having count(*)<>Max(vno)"
    If GProcRstOpen(rsVou, X, "R") > 0 Then
        With rsVou
            .MoveFirst
            Do While Not .EOF
                RCount = 1
                MaxRNo = IIf(IsNumeric(.Fields("MaxVNo")), .Fields("MaxVNo"), 0)
                Do While RCount <= MaxRNo
                    '-- Transaction Missing
                    If GProcGetColumnValue("tblAddLess", "Vno", CStr(RCount), "N", "Vno", "N", " Vtype='" & .Fields("Vtype") & "' and " & CondStr) = 0 Then
                        gCn.Execute "insert into tmpGenTbl (VBillNo,Vno,Vtype,Nar1,UserName) Values (" & RCount & "," & RCount & ",'" & .Fields("VTYpe") & "'," & MaxRNo & ",'" & gUserName & "')"
                    End If
                    RCount = RCount + 1
                Loop
               .MoveNext
            Loop
        End With
    End If
End If
'-- Insert To Tmp Missing Invoices
If VTypeStr = "SA" Then
     '-- ( InvNo String union all Numeric)
    If gBackEndDB = gBackEndAccess Then 'mid(AdInVNo,5)
        X = "select max(clng(mid(AdInVNo,5))) as MaxVNo,Vtype,mid(AdInVNo,1,3) as BillSrNo,AcName from tblAddLess,tblMastAccount where AdAcCrCode =AcCode and isNumeric(adinvno)=false  " _
        & " and " & CondStr & " Group by VType,mid(AdInVNo,1,3),AcName having count(*)<>max(clng(mid(AdInVNo,5))) " _
        & " union all select max(clng(AdInVNo)),Vtype,'','' from tblAddLess where isNumeric(adinvno) " _
        & " and " & CondStr & " Group by VType having count(*)<>max(clng(AdInVNo)) order by BillSrNo" '- Min(to_number(substr(AdInVNo,5)))
    Else '-Oracle
        X = "select max(to_number(substr(AdInVNo,5))) as MaxVNo,Vtype,substr(AdInVNo,1,3) as BillSrNo,AcName from tblAddLess,tblMastAccount where AdAcCrCode =AcCode and instr(adinvno,'-')>0  " _
        & " and " & CondStr & " Group by VType,substr(AdInVNo,1,3),AcName having count(*)<>max(to_number(substr(AdInVNo,5)))" _
        & " union all select max(to_number(AdInVNo)),Vtype,'','' from tblAddLess where instr(adinvno,'-')=0" _
        & " and " & CondStr & " Group by VType having count(*)<>max(to_number(AdInVNo)) order by BillSrNo"
    End If
    If GProcRstOpen(rsVou, X, "R") > 0 Then
        With rsVou
            .MoveFirst
            Do While Not .EOF
                RCount = 1
                MaxRNo = IIf(IsNumeric(.Fields("MaxVNo")), .Fields("MaxVNo"), 0)
                Do While RCount <= MaxRNo
                     MissInVNo = CStr(IIf(IsNull(.Fields("BillSrNo")), "", .Fields("BillSrNo") & "-") & RCount)
                    '-- Transaction Missing
                    If GProcGetColumnValue("tblAddLess", "AdInvNo", MissInVNo, "S", "Vno", "N", " Vtype='" & .Fields("Vtype") & "' and " & CondStr) = 0 Then
                        gCn.Execute "insert into tmpGenTbl (VBillNo,Vno,Vtype,Nar1,Nar2,UserName) Values ('" & MissInVNo & "'," & RCount & ",'" & .Fields("VTYpe") & "','" & .Fields("BillSrNo") & "-" & MaxRNo & "','" & .Fields("AcName") & "','" & gUserName & "')"
                    End If
                    RCount = RCount + 1
                Loop
               .MoveNext
            Loop
        End With
    End If
End If
End Sub
'---- Late Pay Debit Note Register
Public Sub GProcLatePayDbNtReg(FromDt As Date, ToDt As Date, AllFirmFlag As Boolean, Broker As String, Party As String, Optional IsMillBillRec As Boolean)
Dim CondStr As String
Dim X As String
'-- Date
If gBackEndDB = gBackEndAccess Then
    CondStr = " INTDBNTDATE>= cdate('" & FromDt & "')  and INTDBNTDATE<= cdate('" & ToDt & "')"
Else '-- Oracle
    CondStr = " INTDBNTDATE>= to_date('" & FromDt & "','dd/MM/yyyy') and INTDBNTDATE<= to_date('" & ToDt & "','dd/MM/yyyy')"
End If
'-- Type
If IsMillBillRec = False Then
    CondStr = CondStr + " and (RECVTYPE='BR' or RECVTYPE='CR' ) "
Else
    CondStr = CondStr + " and RECVTYPE='MR' " '-- Mill Bill Receipt
End If
If Party <> "" Then CondStr = CondStr & " and tblMastAccount.AcName='" & Party & "'" '--Party
If Broker <> "" Then CondStr = CondStr & " and tblAcc_Brk.AcName='" & Broker & "'" '--Broker
X = " select INTDBNTDATE,INVBILLNO,INVDT,INTDBNTAMT,INTDBNTRECDAMT,ISINTDBNTCOMPLETE,tblMastAccount.AcName," _
& "tblAcc_Brk.AcName,'" & gUserName & "' from tblmastaccount,tblmastaccount tblAcc_Brk,TblOutstanding,TblRecVsSale " _
& " where tblmastaccount.accode=PartyCode and tblAcc_Brk.accode=OUTBROKERCODE and " & CondStr & " And " _
& " vno=Invno and vtype=Invtp and vYear=InvYear "
'--- insert to setup db Table For All Firm
GProcInsertToSetupDbTbl AllFirmFlag, X, "tmpGenTbl", "vdt,VBillNo,VBillDt,Amt1,amt2,Amt3,Nar2,Nar1,username", " username='" & gUserName & "'", "Nar3"
End Sub
'============================================== General Reports Procedure End
Public Function gProcPreparePartyLedgerTFormat()
Dim RsTmp_T As ADODB.Recordset, RsTmp_S As ADODB.Recordset, RsTmpParty As ADODB.Recordset
Dim RCount_T As Long
Dim RCount_S As Long

Dim X As String
Dim BalQuery As String
Dim i As Integer

BalQuery = "Select distinct AcName from qryAccountLedger where UserName='" & gUserName & "' order by AcName"
GProcRstOpen RsTmpParty, BalQuery, "R"
gCn.Execute "delete from tmpGenTbl where UserName='" & gUserName & "'"
'********* All Party Loop
RsTmpParty.MoveFirst
Do While Not RsTmpParty.EOF
        '=== Insert For Credit Side
        BalQuery = "Select AcName,Credit,AcOName,Vdt,UserName,VNar1,VNar2,VNar3,Vno from qryAccountLedger " _
        & " where UserName='" & gUserName & "' and AcName='" & RsTmpParty!AcName & "' and Credit > 0 order by Vdt,Vno,Vtype"
        If GProcRstOpen(RsTmp_S, BalQuery, "R") > 0 Then
            With RsTmp_S
                i = 0
                .MoveFirst
                Do While Not .EOF
                   i = i + 1
                   gCn.Execute "insert into tmpGenTbl (vno,Nar1,Amt1,amt2,Nar2,Vdt,username,Nar3,Nar4,Nar5,Amt3) " _
                   & " values (" & i & ",'" & .Fields(0) & "'," & .Fields(1) & ",0,'" & .Fields(2) & "','" & Format(.Fields(3), "yyyy/MM/dd") & "','" & .Fields(4) & "','" & .Fields(5) & "','" & .Fields(6) & "','" & .Fields(7) & "'," & Val(.Fields(8) & vbNullString) & ")"
'.AbsolutePosition
                   .MoveNext
                Loop
            End With
        End If
        '=== UpDate For Debit Side
        RCount_T = GProcRstOpen(RsTmp_T, "Select * from TmpGenTbl where UserName='" & gUserName & "' and Nar1='" & RsTmpParty!AcName & "' order by vno ", "O")
        
        BalQuery = "Select AcName,Debit,AcOName,Vdt,UserName,VNar1,VNar2,VNar3,Vno from qryAccountLedger " _
        & " where UserName='" & gUserName & "' and AcName='" & RsTmpParty!AcName & "' and Debit > 0 order by Vdt,Vno,Vtype"
        
        RCount_S = GProcRstOpen(RsTmp_S, BalQuery, "R")
        If RCount_S > 0 Then
            With RsTmp_S
                .MoveFirst
                i = 0
                Do While Not .EOF
                   i = i + 1
                   If i > RCount_T Then '- Right Side > Left Side
                        gCn.Execute "insert into tmpGenTbl (vno,Nar1,amt1,Amt2,Nar6,VBilldt,username,Nar7,Nar8,Nar9,Amt4) " _
                        & " values (" & i & ",'" & .Fields(0) & "',0," & .Fields(1) & ",'" & .Fields(2) & "','" & Format(.Fields(3), "yyyy/MM/dd") & "','" & .Fields(4) & "','" & .Fields(5) & "','" & .Fields(6) & "','" & .Fields(7) & "'," & Val(.Fields(8) & vbNullString) & ")"
                   Else
                        gCn.Execute "update tmpGenTbl set Amt2=" & .Fields(1) & ",Nar6='" & .Fields(2) & "',VBillDt='" & Format(.Fields(3), "yyyy/MM/dd") & "',Nar7='" & .Fields(5) & "',Nar8='" & .Fields(6) & "',Nar9='" & .Fields(7) & "',Amt4=" & Val(.Fields(8) & vbNullString) & " where " _
                        & " username='" & .Fields("UserName") & "' and Nar1='" & RsTmpParty!AcName & "' and vno=" & i
                   End If
                    If Not RsTmp_T.EOF Then RsTmp_T.MoveNext
                    .MoveNext
                Loop
            End With
        End If
        RsTmpParty.MoveNext
Loop
If RsTmp_S.State = adStateOpen Then RsTmp_S.Close
If RsTmp_T.State = adStateOpen Then RsTmp_T.Close
If RsTmpParty.State = adStateOpen Then RsTmpParty.Close
Exit Function
ErrorHandler:
   GProcErrorHandler
End Function

Public Sub procAccounts(Optional AcCode As String, Optional gfromdt, Optional gtodt, Optional CmbDebit As String, Optional Prefix As String)
For i = 0 To 40
    PrintTotal(i) = False
    HideCol(i) = False
    RepeateCol(i) = False
Next i
Set Frm1 = New frmReport
DateCaptionFromTo = gfromdt & " to " & gtodt
DateCaptionAsOn = "  As on date " & gtodt
Frm1.Label2.Caption = "" & gSelectedMenu
Frm1.Label1 = DateCaptionFromTo

DoEvents
'-----List Menu
With Frm1
Select Case gSelectedMenu


'-----Ledger Menu
    
Case "Account Ledger"
        .Label1 = DateCaptionFromTo
        Frm1.cmdBtnRetrieve.Visible = True
        Frm1.gfromdt.Visible = True
        Frm1.gtodt.Visible = True
        Frm1.OpCrDr(0).Visible = True
        Frm1.OpCrDr(1).Visible = True
        Frm1.OpCrDr(2).Visible = True
        Frm1.Label4.Visible = True
        Frm1.TxtPrefix.Visible = True
        Frm1.cmdOut(0).Visible = True
                
       Dim crdrcond As String
         If CmbDebit = "Debit" Then
            crdrcond = " and  VDrAmt > 0 "
         ElseIf CmbDebit = "Credit" Then
            crdrcond = " and  VCrAmt > 0 "
         End If
         If frmRptAccount.ChkConsolidate.Value = 0 Then crdrcond = crdrcond + " and tmpaccledger.Vfirm='" & gCCode & "'"
         If Prefix <> "" Then

           QueryStr = " select Vno,vdt,VType,upper(Vbillno),AcName,AcOName,AcCode,VDrAmt as Debit,VCrAmt as Credit,0 as BALANCE, " _
                    & " vbillno,case when vtype='0P' then 1 else 2 end as orderno " _
                    & " from TmpAccledger where  VBillchr ='" & Trim(Prefix) & "' and  username='" & gUserName & "' " & crdrcond _
                    & " order by vdt,orderno,vtype,vbillno"

         Else
            QueryStr = " select Vno,vdt,VType,upper(Vbillno),AcName,AcOName,AcCode,VDrAmt as Debit,VCrAmt as Credit,0,0 as BALANCE, " _
                     & " vbillno as billno,case when vtype='0P' then 1 else 2 end as orderno " _
                     & " from Tmpaccledger where  username='" & gUserName & "' " & crdrcond & " order by orderno,vdt,vtype,(billno)"
               
         End If
         
         If CheckForTempData = False Then
            Exit Sub
         End If
         
            If CmbDebit = "Debit" Then
               HideCol(8) = True
            ElseIf CmbDebit = "Credit" Then
               HideCol(7) = True
            End If
            HideCol(0) = True
            HideCol(4) = True
            HideCol(6) = True
            HideCol(9) = True
            PrintTotal(7) = True
            PrintTotal(8) = True
        PrepareReport 11, 1, 3, "Total : ", True, False, 4, 11, 7, 8, Array("", "DATE", "TYPE", "VBILLNO.", "", "DESCRIPTION", "", "DEBIT", "CREDIT", "", "BALANCE")
'-------------Register
Case "Trial Balance Simple"
        .Label1 = DateCaptionFromTo
         QueryStr = "select AcCode,AcType,Acname,Agname,clDR ,ClCr from tmpTblFin where username='" & gUserName & "' order by Acname"
        If CheckForTempData = False Then
            Exit Sub
        End If
        HideCol(0) = True
        HideCol(1) = True
        PrintTotal(5) = True
        PrintTotal(4) = True
        PrepareReport 6, 1, 4, "Total : ", True, False, 1, 4, 5, 5, Array("", "", "Account Name", "Group Name", "Debit", "Credit")
Case "Simple Sale Register"
        .Label1 = DateCaptionFromTo
         QueryStr = AcCode
        If CheckForTempData = False Then
            Exit Sub
        End If
        HideCol(0) = True
        HideCol(1) = True
        PrintTotal(6) = True
        PrepareReport 8, 2, 6, "Total : ", False, True, 0, , , , Array("", "", "BILLNO", "DATE", "PARTY NAME", "CITY", "BROKER", "AMOUNT")
 

Case "Detail Sale Register"
        .Label1 = DateCaptionFromTo
         QueryStr = AcCode
        If CheckForTempData = False Then
            Exit Sub
        End If
        HideCol(0) = True
        HideCol(4) = True
        PrintTotal(13) = True
        PrepareReport 14, 2, 12, "Total : ", False, True, 4, , , , Array("", "BILLNO", "TYPE", "DATE", "PARTY NAME", "CATEGORY", "PARTY", "ITEM NAME", "QTY", "RATE", "SUB AMOUNT", "ADD'N", "DED'N", "BILL AMOUNT")


Case "Simple Purchase Register"  '--"Simple Purchase Register":
        .Label1 = DateCaptionFromTo
         QueryStr = AcCode
       If CheckForTempData = False Then
            Exit Sub
        End If
        HideCol(0) = True
        PrintTotal(5) = True
        PrepareReport 6, 2, 4, "Total : ", False, True, 0, , , , Array("", "BILLNO", "TYPE", "DATE", "PARTY NAME", "CATEGORY", "PARTY", "ITEM NAME", "QTY", "RATE", "SUB AMOUNT", "ADD'N", "DED'N", "BILL AMOUNT")

Case "Detail Purchase Register"
               
        .Label1 = DateCaptionFromTo
         QueryStr = AcCode
        If CheckForTempData = False Then
            Exit Sub
        End If
        HideCol(0) = True
        PrintTotal(13) = True
        
        PrepareReport 14, 2, 12, "Total : ", False, True, 4, , , , Array("", "BILLNO", "TYPE", "DATE", "PARTY NAME", "CATEGORY", "PARTY", "ITEM NAME", "QTY", "RATE", "SUB AMOUNT", "ADD'N", "DED'N", "BILL AMOUNT")
               
 
Case "SalePartywise"
       .Label1 = DateCaptionFromTo
       QueryStr = "select SALEVNO as VNO,SALEVBILLNO AS BILLNO,SALEVTYPE AS VTYPE,SALEVDT AS VDT,ACNAME,SALEVBILLAMT AS AMOUNT,RECVDT,'' as A,ADjAmt,DISCT, RECVNO,RECVTYPE from QryRecSaleRegi where " & AcCode & " order by SALEVDT,SALEVBILLNO"

        If CheckForTempData = False Then
            Exit Sub
        End If
            
            PrintTotal(8) = True
            PrintTotal(10) = True
        HideCol(0) = True
        HideCol(4) = True
        HideCol(7) = True
        PrepareReport 12, 2, 7, "Total : ", True, True, 4
        
Case "PurchasePartywise"
        .Label1 = DateCaptionFromTo
        QueryStr = "select PURCHVNO AS VNO,PURCHVDT AS VDT,PURCHVTYPE AS VTYPE,PURCHBILLNO AS BILLNO,ACNAME,PURCHBILLAMT AS AMOUNT,RECVNO,RECVDT,ADjAmt,RECVTYPE,DISCT from QRYPAYPURCHREGI where " & AcCode & " order by PURCHBILLNO,PURCHVDT"
        If CheckForTempData = False Then
            Exit Sub
        End If
            PrintTotal(8) = True
            'PrintTotal(8) = True
            'PrintTotal(9) = True
        HideCol(4) = True
        PrepareReport 12, 2, 5, "Total : ", True, True, 4

'-----------Outstanding

Case "Outstanding Sale with Partywise"
        .Label1 = DateCaptionFromTo
        frmReport.cmdBtnRetrieve.Visible = True
    If AcCode <> "" Then
            QueryStr = "select orderNo,VDT,VTYPE,Vbillno as VBILLNo,AcName,VDRAmt as BillAmt,vbilldt as RecDate,VCrAmt as RecVAmt,decode(VDRAmt,0,(VCrAmt*-1),VBalAmt) as BALANCE,VNO,Vnar2,Vnar3 as RecVtype,Vnar4 as REcVno  from tmpAccLedger where username='" & gUserName & "' order by Acname,VNO,VDT,vbilldt"
        If CheckForTempData = False Then
            Exit Sub
        End If
        
        HideCol(4) = True
        HideCol(0) = True
        HideCol(11) = False
        HideCol(12) = False
        HideCol(9) = True
        HideCol(10) = True
        PrintTotal(7) = True
        PrintTotal(8) = True
        PrepareReport 13, 2, 3, "Total : ", True, False, 4
    Else
        QueryStr = "select orderNo,VDT,VTYPE,Vbillno as VBILLNo,AcName,VDRAmt as BillAmt,vbilldt as RecDate,VCrAmt as RecVAmt,(VDRAmt-VCrAmt) as BALANCE,VNO from tmpAccLedger where username='" & gUserName & "' order by Acname,VNO,VDT,vbilldt"
        If CheckForTempData = False Then
            Exit Sub
        End If
        
        HideCol(4) = True
        HideCol(0) = True
        HideCol(11) = False
        HideCol(12) = False
        HideCol(9) = True
        HideCol(10) = True
        PrintTotal(7) = True
        PrintTotal(8) = True
        PrepareReport 9, 2, 3, "Total : ", True, False, 4
            
    End If

Case "Outstanding Purchase with Partywise"

        .Label1 = DateCaptionFromTo
        frmReport.cmdBtnRetrieve.Visible = True
        'QueryStr = "select orderNo,VDT,VTYPE,Vbillno as VBILLNo,AcName,VDRAmt as BillAmt,vbilldt as RecDate,VCrAmt as RecVAmt,decode(VDRAmt,0,(VCrAmt*-1),VBalAmt) as BALANCE,VNO,Vnar2,Vnar3 as RecVtype,Vnar4 as REcVno  from tmpAccLedger where username='" & gUserName & "' order by VNO,VDT"
        
        QueryStr = "select orderNo,VDT,VTYPE,Vbillno as VBILLNo,AcName,VDRAmt as BillAmt,vbilldt as RecDate,VCrAmt as RecVAmt,decode(VDRAmt,0,(VCrAmt*-1),VBalAmt) as BALANCE,VNO,Vnar2,Vnar3 as RecVtype,Vnar4 as REcVno  from tmpAccLedger where username='" & gUserName & "' order by VNO,VDT"
        
        If CheckForTempData = False Then
            Exit Sub
        End If
        HideCol(4) = True
        HideCol(0) = True
        HideCol(11) = False
        HideCol(12) = False
        HideCol(4) = True
        HideCol(0) = True
        HideCol(9) = True
        PrintTotal(7) = True
        PrintTotal(8) = True
        HideCol(10) = True
        PrepareReport 13, 2, 3, "Total : ", True, False, 4

End Select
End With

If gSelectedMenu <> "" Then
   Frm1.Show
End If
End Sub


Public Sub PrepareReport(ColumnNumber As Long, AddRows As Long, TotalCol As Long, TotalText As String, GroupTotal As Boolean, GroupNumber As Boolean, GroupCol As Long, Optional RunTotColNo As Integer, Optional RunTotofColNo1 As Integer, Optional RunTotofColNo2 As Integer, Optional PCaption)
On Error Resume Next
Dim QStr As String
Dim ColNo As Long
Dim Rowno As Long
Dim GroupSrNo As Long
Dim SrNo As Long
Dim Amt As Double
Dim TickRecord As Long
Dim GTotal(40) As Double
Dim Total(40) As Double
DoEvents
If Frm1.cmdBtnRetrieve.Value = True Then
Frm1.msgReport.Clear
End If
Dim mCaption
Dim tmpRecord As Double
   mCaption = PCaption
'frm1Prep.Label1 = "Preparing Report for Display....Please Wait...."

With Frm1.msgReport
.cols = ColumnNumber
rsReport.MoveLast
.rows = Val(rsReport.RecordCount) + Val(AddRows)
' tmpRecord = Val(rsReport.RecordCount) + Val(AddRows)
If .rows < rsReport.RecordCount Then
    MsgBox "Select Limited Period ", vbInformation + vbOKOnly
    Unload Frm1
    Exit Sub
End If

TickRecord = .rows / 10

Frm1.prgBar.Visible = True
Frm1.prgBar.Value = 0
Frm1.prgBar.Min = 0
Frm1.prgBar.Max = TickRecord



.FixedRows = 1
.FixedCols = 0
.Row = 0

For i = 0 To ColumnNumber - 1
   .Col = i
   .CellFontBold = True
   .text = rsReport.Fields(i).Name
   Select Case rsReport.Fields(i).Type
        Case 200, 3, 135, 202, 5, 130
            .ColAlignment(i) = flexAlignLeftCenter
            If rsReport.Fields(i).ActualSize <= 20 Then
                .ColWidth(i) = 1000
            Else
                .ColWidth(i) = 2500
            End If

            If UCase((rsReport.Fields(i).Name)) = "VNO" Then
                .ColWidth(i) = 700
            End If
            If UCase((rsReport.Fields(i).Name)) = "VBILLNO" Then
                .ColWidth(i) = 1800
            End If
            If UCase(rsReport.Fields(i).Name) = "VDT" Then
                .ColWidth(i) = 1500
            End If
            If UCase(rsReport.Fields(i).Name) = "VTYPE" Then
                .ColWidth(i) = 700
            End If
            If UCase(rsReport.Fields(i).Name) = "CITY" Then
                .ColWidth(i) = 2000
            End If
            If UCase(rsReport.Fields(i).Name) = "ACNAME" Or UCase(rsReport.Fields(i).Name) = "PARTY" Then
                .ColWidth(i) = 3000
            End If
            If UCase(rsReport.Fields(i).Name) = "PARTICULARS" Then
                '.ColWidth(i) = 5000
                .ColWidth(i) = 3000
            End If
            If UCase(rsReport.Fields(i).Name) = "ACCOUNT" Then
                '.ColWidth(i) = 5000
                .ColWidth(i) = 3000
            End If
            If UCase(rsReport.Fields(i).Name) = "BALANCE" Then
                .ColAlignment(i) = flexAlignRightCenter
                .ColWidth(i) = 1800
            End If
            If UCase(rsReport.Fields(i).Name) = "DEBIT" Then
                .ColAlignment(i) = flexAlignRightCenter
                .ColWidth(i) = 1500
            End If
            If UCase(rsReport.Fields(i).Name) = "CREDIT" Then
                .ColAlignment(i) = flexAlignRightCenter
                .ColWidth(i) = 1500
            End If

            If Right(rsReport.Fields(i).Name, 4) = "Vdt" Then
                .ColWidth(i) = 1200
            End If
          
        Case 131, 139
            .ColAlignment(i) = flexAlignRightCenter
            .ColWidth(i) = 1500
        Case 2, 3
            .ColAlignment(i) = flexAlignRightCenter
            .ColWidth(i) = 1200
        Case 135
            .ColAlignment(i) = flexAlignRightCenter
            .ColWidth(i) = 1200
    End Select
    If HideCol(i) = True Then
        .ColWidth(i) = 0
    End If
'End If
Next i
For i = 0 To rsReport.Fields.Count - 1 'msFlex.Cols - 1
    .Row = 0
    .Col = i
    .text = mCaption(i)
Next

If Not IsNull(RunTotColNo) Then
   If RunTotColNo = i Then
   .Col = i
   .CellFontBold = True
   .text = "Balance"
   End If
End If
rsReport.MoveFirst
Rowno = 1
SrNo = 1
Amt = 0
Do While Not rsReport.EOF
    If rsReport.Fields(GroupCol) <> "" Then
        QStr = rsReport.Fields(GroupCol)
    Else
        rsReport.MoveNext
    End If
    '-----Initializing GroupTotal to 0 before starting group
    For i = 0 To 20
        GTotal(i) = 0
    Next i
    '-----
    If GroupCol >= 1 Then
        .Row = Rowno
        .MergeCells = flexMergeFree
        .MergeRow(Rowno) = True
        For i = 0 To ColumnNumber - 1
            .Col = i
            .CellFontBold = True
           ' .CellForeColor = vbActiveTitleBar
            .CellAlignment = flexAlignLeftCenter
            .CellBackColor = vbInactiveTitleBar
            .CellForeColor = vbWhite
            .text = QStr
            
           
        Next i
       Rowno = Rowno + 1
       Amt = 0
       .AddItem Rowno
       .Row = Rowno
    End If
    Do While rsReport.Fields(GroupCol) = QStr
        .Row = Rowno
        'If .Row Mod TickRecord = 0 Then
            Frm1.prgBar.Value = Frm1.prgBar.Value + 1
        'End If
        For i = 0 To ColumnNumber - 1
            .Col = i
            '-----Print Sr No if group no are not reqd
               Select Case rsReport.Fields(i).Type
                    Case 200, 202, 130
                       If rsReport.Fields(i) <> "" Then
                           .text = rsReport.Fields(i)
                       End If
                       
                   Case 131, 139, 3, 5
                       If CDbl(rsReport.Fields(i)) <> 0 Then
                           .text = Format(rsReport.Fields(i), "#####0.00####")
                       End If
                   Case 135
                       If rsReport.Fields(i) <> "" Then
                           .text = rsReport.Fields(i)
                       End If
                   Case 155 '"dbDouble"
                       If rsReport.Fields(i) <> "" Then
                           If CDbl(rsReport.Fields(i)) <> 0 Then
                               If Right(rsReport.Fields(i).Name, 3) = "Qty" Then
                                   .text = Format((rsReport.Fields(i)), LStr)
                               ElseIf rsReport.Fields(i).Name = "Cartan" Then
                                   .text = Format(rsReport.Fields(i), "#####0.0####")
                               Else
                                   .text = Format((rsReport.Fields(i)), FStr)
                               End If
                           End If
                       End If
            End Select
                DoEvents
                
                If PrintTotal(i) = True Then
                       If Trim(rsReport.Fields(2)) = "OP" And ReportRecStr = "Daywise Account Ledger" And Flage = True Then
                       
                          DoEvents
                       ElseIf rsReport.Fields(i) <> "" Then
                         Total(i) = Total(i) + rsReport.Fields(i)
                         GTotal(i) = GTotal(i) + rsReport.Fields(i)
                         
                     End If
                End If
        Next i
        If Not IsNull(RunTotColNo) Then
           If RunTotColNo = i Then
             .Col = i
                  Amt = Amt + rsReport.Fields(RunTotofColNo1) - rsReport.Fields(RunTotofColNo2)
                 .text = Format(Abs(Amt), FStr) + " " + IIf(Amt > 0, "Dr", "Cr")
           End If
        End If
        
        Rowno = Rowno + 1
        rsReport.MoveNext
        If rsReport.EOF Then
            Exit Do
        End If
    Loop
    If GroupTotal = True Then
        .AddItem Rowno
        .Row = Rowno
        .Col = 0
        .text = ""
        .Col = TotalCol
        .CellFontBold = True
        .CellAlignment = flexAlignRightCenter
        .text = "Group Total : "
        Dim mRw As Long
        Dim mCol As Long
        If ReportRecStr = "Daywise Account Ledger" And Flage = True Then
        'mRw = 1
            Do While mRw < .rows
                If Trim(.TextMatrix(mRw, 2)) = "OP" Then
                  '  mCol = 6
                    Exit Do
                End If
                mRw = mRw + 1
            Loop
            .TextMatrix(.Row, 5) = .TextMatrix(mRw, 6)
            mRw = mRw + 1
            'mCol = 0
        End If
        For i = 0 To ColumnNumber - 1
           .Col = i
           If PrintTotal(i) = True Then
               Select Case rsReport.Fields(i).Type
                    Case "dbLong"
                        .text = Format(GTotal(i), FStr)
                    Case "dbDouble"
                        If Right(rsReport.Fields(i).Name, 3) = "Qty" Then
                            .text = Format(GTotal(i), LStr)
                        Else
                            .text = Format(GTotal(i), FStr)
                        End If
                End Select
            End If
        Next i
        Rowno = Rowno + 1
        
    End If
Loop
If AddRows = 2 Then
    .Row = Rowno
    .Col = 0
    .text = ""
    .Col = TotalCol
    .CellFontBold = True
    .CellAlignment = flexAlignRightCenter
    .text = TotalText

    For i = 0 To ColumnNumber - 1
        .Col = i
        If PrintTotal(i) = True Then
            Select Case rsReport.Fields(i).Type
                Case "dbLong"
                    .text = Format(Total(i), PStr)
                Case "dbDouble"
                    If Right(rsReport.Fields(i).Name, 3) = "Qty" Then
                        .text = Format(Total(i), LStr)
                    Else
                        .text = Format(Total(i), FStr)
                    End If
            End Select
        End If
    Next i
End If
End With
Frm1.prgBar.Visible = False
End Sub
Public Function CheckForTempData() As Boolean
GProcRstOpen rsReport, QueryStr, "R", gCn
CheckForTempData = True
If rsReport.EOF And rsReport.BOF Then
    MsgBox "No records to display", vbOKOnly + vbExclamation, "Cloth_General....."
    CheckForTempData = False
End If
End Function
Public Function GProcPurpareGSTSummery(w_Fr As Date, w_To As Date, Optional MoYr As Boolean)
Dim x1 As String
             
gCn.BeginTrans
gCn.Execute "delete from tmpGenTbl where username='" & gUserName & "'"
gCn.CommitTrans
gCn.BeginTrans

If MoYr = False Then  '--For Monthly
    '---Sales Detail
    If gCIsDepotFirm = 1 Then
       x1 = " tblSale.VType ='SD'  and (tblSale.VDt)>= ('" & Format(w_Fr, "yyyy/mm/dd") & "') and (tblSale.VDt) <= ('" & Format(w_To, "yyyy/mm/dd") & "') and tblSale.SlTaxAmt+tblSale.SlTaxAmt2+tblSale.SlTaxAmt3 > 0  and tblSale.Vfirm = '" & gCCode & "'"  '
    Else
       x1 = " tblSale.VType in ('SY','SC','SB')  and (tblSale.VDt)>= ('" & Format(w_Fr, "yyyy/mm/dd") & "') and (tblSale.VDt) <= ('" & Format(w_To, "yyyy/mm/dd") & "') and tblSale.SlTaxAmt+tblSale.SlTaxAmt2+tblSale.SlTaxAmt3 > 0  and tblSale.Vfirm = '" & gCCode & "'"  '
    End If
    gCn.Execute " Insert Into TmpGentbl (Nar1,vyear,Amt1,Nar2,Nar3,Amt2,Amt3,Amt4,Amt5,Amt6,Amt7,UserName)" _
                & " select '1) Output',year(tblSale.vdt),month(tblSale.Vdt), format(tblSale.Vdt,'MMM-yyyy'),iif(tblSale.Vtype = 'SY','Trade Sales',iif(tblSale.Vtype= 'SC','Other Sales',iif(tblSale.Vtype = 'SB','Brokerage & Comm. Sales',iif(tblsale.vtype = 'SD','Depot Sales','XXX')   ))),sum(tblSale.SlTaxableAmt),sum(tblSale.SlTaxAmt3),sum(tblSale.SlTaxAmt),sum(tblSale.SlTaxAmt2),sum(tblSale.SlBillAmt),sum(tblSale.SlTcsAmt),'" & gUserName & "'" _
                & " From tblSale where  " & x1 & "" _
                & " group by year(tblSale.vdt),month(tblSale.Vdt), format(tblSale.Vdt,'MMM-yyyy'),iif(tblSale.Vtype = 'SY','Trade Sales',iif(tblSale.Vtype= 'SC','Other Sales',iif(tblSale.Vtype = 'SB','Brokerage & Comm. Sales',iif(tblsale.vtype = 'SD','Depot Sales','XXX'))))"
    '--Sales Rtn
    
    If gCIsDepotFirm = 0 Then
       x1 = " tblSale.VType = 'RY' and (tblSale.VDt)>= ('" & Format(w_Fr, "yyyy/mm/dd") & "') and (tblSale.VDt) <= ('" & Format(w_To, "yyyy/mm/dd") & "') and tblSale.SlTaxAmt+tblSale.SlTaxAmt2+tblSale.SlTaxAmt3 > 0  and tblSale.Vfirm = '" & gCCode & "'"  '
       gCn.Execute " Insert Into TmpGentbl (Nar1,vyear,Amt1,Nar2,Nar3,Amt2,Amt3,Amt4,Amt5,Amt6,UserName)" _
                   & " select '1) Output',year(tblSale.vdt),month(tblSale.Vdt), format(tblSale.Vdt,'MMM-yyyy'),'Sales Return',sum(tblSale.SlTaxableAmt)*-1,sum(tblSale.SlTaxAmt3)*-1,sum(tblSale.SlTaxAmt)*-1,sum(tblSale.SlTaxAmt2)*-1,sum(tblSale.SlBillAmt)*-1,'" & gUserName & "'" _
                   & " From tblSale where  " & x1 & "" _
                   & " group by year(tblSale.vdt),month(tblSale.Vdt), format(tblSale.Vdt,'MMM-yyyy')"
    
    End If
    
    
    '--LP Int. Sales
    x1 = " tblIntSale.VType in ('SI')  and (tblIntSale.VDt)>= ('" & Format(w_Fr, "yyyy/mm/dd") & "') and (tblIntSale.VDt) <= ('" & Format(w_To, "yyyy/mm/dd") & "') and tblIntSale.CgstAmt+tblIntSale.SgstAmt+tblIntSale.IgstAmt > 0  and tblIntSale.Vfirm = '" & gCCode & "'"  '
    gCn.Execute " Insert Into TmpGentbl (Nar1,vyear,Amt1,Nar2,Nar3,Amt2,Amt3,Amt4,Amt5,Amt6,UserName)" _
                & " select '1) Output',year(tblIntSale.vdt),month(tblIntSale.Vdt), format(tblIntSale.Vdt,'MMM-yyyy'),'L.P. Interest',sum(tblIntSale.Interest),sum(tblIntSale.IgstAmt),sum(tblIntSale.CgstAmt),sum(tblIntSale.SgstAmt),sum(tblIntSale.VAmt),'" & gUserName & "'" _
                & " From tblIntSale where  " & x1 & "" _
                & " group by year(tblIntSale.vdt),month(tblIntSale.Vdt), format(tblIntSale.Vdt,'MMM-yyyy')"
                 
    '--LP Int. Without GST Sales
    x1 = " tblIntSale.VType in ('MI')  and (tblIntSale.VDt)>= ('" & Format(w_Fr, "yyyy/mm/dd") & "') and (tblIntSale.VDt) <= ('" & Format(w_To, "yyyy/mm/dd") & "') and tblIntSale.CgstAmt+tblIntSale.SgstAmt+tblIntSale.IgstAmt = 0  and tblIntSale.Vfirm = '" & gCCode & "'"  '
    gCn.Execute " Insert Into TmpGentbl (Nar1,vyear,Amt1,Nar2,Nar3,Amt2,Amt6,UserName)" _
                & " select '2) Output-2',year(tblIntSale.vdt),month(tblIntSale.Vdt), format(tblIntSale.Vdt,'MMM-yyyy'),'L.P. Int. Without GST',sum(tblIntSale.Vamt),sum(tblIntSale.Vamt),'" & gUserName & "'" _
                & " From tblIntSale where  " & x1 & "" _
                & " group by year(tblIntSale.vdt),month(tblIntSale.Vdt), format(tblIntSale.Vdt,'MMM-yyyy')"
    '---Sales Other Without GST Detail ( Wind Mill)
    x1 = " tblSale.VType in ('SC')  and (tblSale.VDt)>= ('" & Format(w_Fr, "yyyy/mm/dd") & "') and (tblSale.VDt) <= ('" & Format(w_To, "yyyy/mm/dd") & "') and tblSale.SlTaxAmt+tblSale.SlTaxAmt2+tblSale.SlTaxAmt3 = 0  and tblSale.Vfirm = '" & gCCode & "'"  '
    gCn.Execute " Insert Into TmpGentbl (Nar1,vyear,Amt1,Nar2,Nar3,Amt2,Amt6,UserName)" _
                & " select '2) Output-2',year(tblSale.vdt),month(tblSale.Vdt), format(tblSale.Vdt,'MMM-yyyy'),'Other Sales Without GST',sum(tblSale.SlBillAmt-tblSale.SlTcsAmt),sum(tblSale.SlBillAmt),'" & gUserName & "'" _
                & " From tblSale where  " & x1 & "" _
                & " group by year(tblSale.vdt),month(tblSale.Vdt), format(tblSale.Vdt,'MMM-yyyy') "
                 
    '--Credit Note Sales
                     
    x1 = " tblIntSale.VType = 'PX'  and (tblIntSale.VDt)>= ('" & Format(w_Fr, "yyyy/mm/dd") & "') and (tblIntSale.VDt) <= ('" & Format(w_To, "yyyy/mm/dd") & "') and tblIntSale.CgstAmt+tblIntSale.SgstAmt+tblIntSale.IgstAmt > 0 and tblIntSale.Vfirm = '" & gCCode & "'"  '
    
    gCn.Execute " Insert Into TmpGentbl (Nar1,vyear,Amt1,Nar2,Nar3,Amt2,Amt3,Amt4,Amt5,Amt6,UserName)" _
                & " select '1) Output',year(tblIntSale.vdt),month(tblIntSale.Vdt), format(tblIntSale.Vdt,'MMM-yyyy'),'Credit Note Sales',sum(tblIntSale.Interest)*-1,sum(tblIntSale.IgstAmt)*-1,sum(tblIntSale.CgstAmt)*-1,sum(tblIntSale.SgstAmt)*-1,sum(tblIntSale.VAmt)*-1,'" & gUserName & "'" _
                & " From tblIntSale where  " & x1 & "" _
                & " group by year(tblIntSale.vdt),month(tblIntSale.Vdt), format(tblIntSale.Vdt,'MMM-yyyy')"
                 
    '-- Purch & Other Purch
    
    x1 = " tblPurch.VType in ('PY','PO') and tblPurch.PurBillDt>= ('" & Format(w_Fr, "yyyy/mm/dd") & "')  and tblPurch.PurBillDt<=('" & Format(w_To, "yyyy/mm/dd") & "') and tblPurch.PurTaxAmt3+tblPurch.PurTaxAmt2+tblPurch.PurTaxAmt > 0  and tblPurch.Vfirm = '" & gCCode & "' "
    gCn.Execute " Insert Into TmpGentbl (Nar1,vyear,Amt1,Nar2,Nar3,Amt2,Amt3,Amt4,Amt5,Amt6,Amt7,UserName)" _
                & " select '3) Input',year(tblPurch.PurBillDt),month(tblPurch.PurBillDt), format(tblPurch.PurBillDt,'MMM-yyyy'),iif(tblPurch.Vtype = 'PO','Other Purch','Trade Purch'),sum(tblPurch.PurTaxableAmt),sum(tblPurch.PurTaxAmt3),sum(tblPurch.PurTaxAmt2),sum(tblPurch.PurTaxAmt),sum(tblPurch.PurBillAmt),sum(tblPurch.PurTcsAmt),'" & gUserName & "'" _
                & " From tblPurch where  " & x1 & "" _
                & " group by year(tblPurch.PurBillDt),month(tblPurch.PurBillDt), format(tblPurch.PurBillDt,'MMM-yyyy'),iif(tblPurch.Vtype = 'PO','Other Purch','Trade Purch')"
    
    '--Purch Return
    x1 = " tblPurch.VType = 'VY' and tblPurch.Vdt>= ('" & Format(w_Fr, "yyyy/mm/dd") & "')  and tblPurch.Vdt<=('" & Format(w_To, "yyyy/mm/dd") & "') and tblPurch.PurTaxAmt3+tblPurch.PurTaxAmt2+tblPurch.PurTaxAmt > 0  and tblPurch.Vfirm = '" & gCCode & "' "
    gCn.Execute " Insert Into TmpGentbl (Nar1,vyear,Amt1,Nar2,Nar3,Amt2,Amt3,Amt4,Amt5,Amt6,UserName)" _
                & " select '3) Input',year(tblPurch.Vdt),month(tblPurch.Vdt), format(tblPurch.Vdt,'MMM-yyyy'),'Purchase Rtn',sum(tblPurch.PurTaxableAmt)*-1,sum(tblPurch.PurTaxAmt3)*-1,sum(tblPurch.PurTaxAmt2)*-1,sum(tblPurch.PurTaxAmt)*-1,sum(tblPurch.PurBillAmt)*-1,'" & gUserName & "'" _
                & " From tblPurch where  " & x1 & "" _
                & " group by year(tblPurch.Vdt),month(tblPurch.Vdt), format(tblPurch.Vdt,'MMM-yyyy')"
    
    
    '--Credit Note Purchase (GST)
    x1 = " tblIntSale.VType = 'PV'  and (tblIntSale.VDt)>= ('" & Format(w_Fr, "yyyy/mm/dd") & "') and (tblIntSale.VDt) <= ('" & Format(w_To, "yyyy/mm/dd") & "') and tblIntSale.CgstAmt+tblIntSale.SgstAmt+tblIntSale.IgstAmt > 0  and tblIntSale.Vfirm = '" & gCCode & "'"  '
    gCn.Execute " Insert Into TmpGentbl (Nar1,vyear,Amt1,Nar2,Nar3,Amt2,Amt3,Amt4,Amt5,Amt6,UserName)" _
                & " select '3) Input',year(tblIntSale.vdt),month(tblIntSale.Vdt), format(tblIntSale.Vdt,'MMM-yyyy'),'Credit Note',sum(tblIntSale.Interest),sum(tblIntSale.IgstAmt),sum(tblIntSale.CgstAmt),sum(tblIntSale.SgstAmt),sum(tblIntSale.VAmt),'" & gUserName & "'" _
                & " From tblIntSale where  " & x1 & "" _
                & " group by year(tblIntSale.vdt),month(tblIntSale.Vdt), format(tblIntSale.Vdt,'MMM-yyyy')"
    
    '--Debit Note Purchase (GST)
    x1 = " tblIntSale.VType = 'SV'  and (tblIntSale.VDt)>= ('" & Format(w_Fr, "yyyy/mm/dd") & "') and (tblIntSale.VDt) <= ('" & Format(w_To, "yyyy/mm/dd") & "') and tblIntSale.CgstAmt+tblIntSale.SgstAmt+tblIntSale.IgstAmt > 0  and tblIntSale.Vfirm = '" & gCCode & "'"  '
    gCn.Execute " Insert Into TmpGentbl (Nar1,vyear,Amt1,Nar2,Nar3,Amt2,Amt3,Amt4,Amt5,Amt6,UserName)" _
                & " select '3) Input',year(tblIntSale.vdt),month(tblIntSale.Vdt), format(tblIntSale.Vdt,'MMM-yyyy'),'Debit Note',sum(tblIntSale.Interest)*-1,sum(tblIntSale.IgstAmt)*-1,sum(tblIntSale.CgstAmt)*-1,sum(tblIntSale.SgstAmt)*-1,sum(tblIntSale.VAmt)*-1,'" & gUserName & "'" _
                & " From tblIntSale where  " & x1 & "" _
                & " group by year(tblIntSale.vdt),month(tblIntSale.Vdt), format(tblIntSale.Vdt,'MMM-yyyy')"
    
    
    
    
    
    '---Input RCM
    x1 = " tblPurch.VType = 'PO' and tblPurch.PurBillDt>= ('" & Format(w_Fr, "yyyy/mm/dd") & "')  and tblPurch.PurBillDt<=('" & Format(w_To, "yyyy/mm/dd") & "') and tblPurch.PurRCMCgstAmt+tblPurch.PurRCMSgstAmt+tblPurch.PurRCMIgstAmt > 0 and tblPurch.Vfirm = '" & gCCode & "'  "  '
                 
    gCn.Execute " Insert Into TmpGentbl (Nar1,vyear,Amt1,Nar2,Nar3,Amt2,Amt3,Amt4,Amt5,Amt6,UserName)" _
                             & " select '3) Input',year(tblPurch.PurBillDt),month(tblPurch.PurBillDt), format(tblPurch.PurBillDt,'MMM-yyyy'),'RCM',sum(tblPurch.PurSubAmt),sum(tblPurch.PurRCMIgstAmt),sum(tblPurch.PurRCMCgstAmt),sum(tblPurch.PurRCMSgstAmt),sum(tblPurch.PurBillAmt),'" & gUserName & "'" _
                             & " From tblPurch where " & x1 & "" _
                             & " group by year(tblPurch.PurBillDt),month(tblPurch.PurBillDt), format(tblPurch.PurBillDt,'MMM-yyyy')"
    
    
    gCn.Execute " Insert Into TmpGentbl (Nar5,Nar1,vyear,Amt1,Nar2,Nar3,Amt2,Amt3,Amt4,Amt5,Amt6,UserName)" _
                             & " select 'B','RCM Pay',year(tblPurch.PurBillDt),month(tblPurch.PurBillDt), format(tblPurch.PurBillDt,'MMM-yyyy'),'RCM',sum(tblPurch.PurSubAmt),sum(tblPurch.PurRCMIgstAmt),sum(tblPurch.PurRCMCgstAmt),sum(tblPurch.PurRCMSgstAmt),sum(tblPurch.PurBillAmt),'" & gUserName & "'" _
                             & " From tblPurch where " & x1 & "" _
                             & " group by year(tblPurch.PurBillDt),month(tblPurch.PurBillDt), format(tblPurch.PurBillDt,'MMM-yyyy')"


Else   '-- For Yearly
    '---Sales Detail
    If gCIsDepotFirm = 1 Then
       x1 = " tblSale.VType ='SD'  and (tblSale.VDt)>= ('" & Format(w_Fr, "yyyy/mm/dd") & "') and (tblSale.VDt) <= ('" & Format(w_To, "yyyy/mm/dd") & "') and tblSale.SlTaxAmt+tblSale.SlTaxAmt2+tblSale.SlTaxAmt3 > 0  and tblSale.Vfirm = '" & gCCode & "'"  '
    Else
       x1 = " tblSale.VType in ('SY','SC','SB')  and (tblSale.VDt)>= ('" & Format(w_Fr, "yyyy/mm/dd") & "') and (tblSale.VDt) <= ('" & Format(w_To, "yyyy/mm/dd") & "') and tblSale.SlTaxAmt+tblSale.SlTaxAmt2+tblSale.SlTaxAmt3 > 0  and tblSale.Vfirm = '" & gCCode & "'"  '
    End If
    gCn.Execute " Insert Into TmpGentbl (Nar1,Nar3,Amt2,Amt3,Amt4,Amt5,Amt6,Amt7,UserName)" _
                & " select '1) Output',iif(tblSale.Vtype = 'SY','Trade Sales',iif(tblSale.Vtype= 'SC','Other Sales',iif(tblSale.Vtype = 'SB','Brokerage & Comm. Sales',iif(tblsale.vtype = 'SD','Depot Sales','XXX')   ))),sum(tblSale.SlTaxableAmt),sum(tblSale.SlTaxAmt3),sum(tblSale.SlTaxAmt),sum(tblSale.SlTaxAmt2),sum(tblSale.SlBillAmt),sum(tblSale.SlTcsAmt),'" & gUserName & "'" _
                & " From tblSale where  " & x1 & "" _
                & " group by iif(tblSale.Vtype = 'SY','Trade Sales',iif(tblSale.Vtype= 'SC','Other Sales',iif(tblSale.Vtype = 'SB','Brokerage & Comm. Sales',iif(tblsale.vtype = 'SD','Depot Sales','XXX'))))"
    '--Sales Rtn
    
    If gCIsDepotFirm = 0 Then
       x1 = " tblSale.VType = 'RY' and (tblSale.VDt)>= ('" & Format(w_Fr, "yyyy/mm/dd") & "') and (tblSale.VDt) <= ('" & Format(w_To, "yyyy/mm/dd") & "') and tblSale.SlTaxAmt+tblSale.SlTaxAmt2+tblSale.SlTaxAmt3 > 0  and tblSale.Vfirm = '" & gCCode & "'"  '
       gCn.Execute " Insert Into TmpGentbl (Nar1,Nar3,Amt2,Amt3,Amt4,Amt5,Amt6,UserName)" _
                   & " select '1) Output','Sales Return',sum(tblSale.SlTaxableAmt)*-1,sum(tblSale.SlTaxAmt3)*-1,sum(tblSale.SlTaxAmt)*-1,sum(tblSale.SlTaxAmt2)*-1,sum(tblSale.SlBillAmt)*-1,'" & gUserName & "'" _
                   & " From tblSale where  " & x1 & ""
    
    End If
    
    
    '--LP Int. Sales
    x1 = " tblIntSale.VType in ('SI')  and (tblIntSale.VDt)>= ('" & Format(w_Fr, "yyyy/mm/dd") & "') and (tblIntSale.VDt) <= ('" & Format(w_To, "yyyy/mm/dd") & "') and tblIntSale.CgstAmt+tblIntSale.SgstAmt+tblIntSale.IgstAmt > 0  and tblIntSale.Vfirm = '" & gCCode & "'"  '
    gCn.Execute " Insert Into TmpGentbl (Nar1,Nar3,Amt2,Amt3,Amt4,Amt5,Amt6,UserName)" _
                & " select '1) Output','L.P. Interest',sum(tblIntSale.Interest),sum(tblIntSale.IgstAmt),sum(tblIntSale.CgstAmt),sum(tblIntSale.SgstAmt),sum(tblIntSale.VAmt),'" & gUserName & "'" _
                & " From tblIntSale where  " & x1 & ""
                 
    '--LP Int. Without GST Sales
    x1 = " tblIntSale.VType in ('MI')  and (tblIntSale.VDt)>= ('" & Format(w_Fr, "yyyy/mm/dd") & "') and (tblIntSale.VDt) <= ('" & Format(w_To, "yyyy/mm/dd") & "') and tblIntSale.CgstAmt+tblIntSale.SgstAmt+tblIntSale.IgstAmt = 0  and tblIntSale.Vfirm = '" & gCCode & "'"  '
    gCn.Execute " Insert Into TmpGentbl (Nar1,Nar3,Amt2,Amt6,UserName)" _
                & " select '2) Output-2','L.P. Int. Without GST',sum(tblIntSale.Vamt),sum(tblIntSale.Vamt),'" & gUserName & "'" _
                & " From tblIntSale where  " & x1 & ""
    '---Sales Other Without GST Detail ( Wind Mill)
    x1 = " tblSale.VType in ('SC')  and (tblSale.VDt)>= ('" & Format(w_Fr, "yyyy/mm/dd") & "') and (tblSale.VDt) <= ('" & Format(w_To, "yyyy/mm/dd") & "') and tblSale.SlTaxAmt+tblSale.SlTaxAmt2+tblSale.SlTaxAmt3 = 0  and tblSale.Vfirm = '" & gCCode & "'"  '
    gCn.Execute " Insert Into TmpGentbl (Nar1,Nar3,Amt2,Amt6,UserName)" _
                & " select '2) Output-2','Other Sales Without GST',sum(tblSale.SlBillAmt-tblSale.SlTcsAmt),sum(tblSale.SlBillAmt),'" & gUserName & "'" _
                & " From tblSale where  " & x1 & ""
                 
    '--Credit Note Sales
                     
    x1 = " tblIntSale.VType = 'PX'  and (tblIntSale.VDt)>= ('" & Format(w_Fr, "yyyy/mm/dd") & "') and (tblIntSale.VDt) <= ('" & Format(w_To, "yyyy/mm/dd") & "') and tblIntSale.CgstAmt+tblIntSale.SgstAmt+tblIntSale.IgstAmt > 0 and tblIntSale.Vfirm = '" & gCCode & "'"  '
    
    gCn.Execute " Insert Into TmpGentbl (Nar1,Nar3,Amt2,Amt3,Amt4,Amt5,Amt6,UserName)" _
                & " select '1) Output','Credit Note Sales',sum(tblIntSale.Interest)*-1,sum(tblIntSale.IgstAmt)*-1,sum(tblIntSale.CgstAmt)*-1,sum(tblIntSale.SgstAmt)*-1,sum(tblIntSale.VAmt)*-1,'" & gUserName & "'" _
                & " From tblIntSale where  " & x1 & ""
                 
    '-- Purch & Other Purch
    
    x1 = " tblPurch.VType in ('PY','PO') and tblPurch.PurBillDt>= ('" & Format(w_Fr, "yyyy/mm/dd") & "')  and tblPurch.PurBillDt<=('" & Format(w_To, "yyyy/mm/dd") & "') and tblPurch.PurTaxAmt3+tblPurch.PurTaxAmt2+tblPurch.PurTaxAmt > 0  and tblPurch.Vfirm = '" & gCCode & "' "
    gCn.Execute " Insert Into TmpGentbl (Nar1,Nar3,Amt2,Amt3,Amt4,Amt5,Amt6,Amt7,UserName)" _
                & " select '3) Input',iif(tblPurch.Vtype = 'PO','Other Purch','Trade Purch'),sum(tblPurch.PurTaxableAmt),sum(tblPurch.PurTaxAmt3),sum(tblPurch.PurTaxAmt2),sum(tblPurch.PurTaxAmt),sum(tblPurch.PurBillAmt),sum(tblPurch.PurTcsAmt),'" & gUserName & "'" _
                & " From tblPurch where  " & x1 & "" _
                & " group by iif(tblPurch.Vtype = 'PO','Other Purch','Trade Purch')"
    
    '--Purch Return
    x1 = " tblPurch.VType = 'VY' and tblPurch.Vdt>= ('" & Format(w_Fr, "yyyy/mm/dd") & "')  and tblPurch.Vdt<=('" & Format(w_To, "yyyy/mm/dd") & "') and tblPurch.PurTaxAmt3+tblPurch.PurTaxAmt2+tblPurch.PurTaxAmt > 0  and tblPurch.Vfirm = '" & gCCode & "' "
    gCn.Execute " Insert Into TmpGentbl (Nar1,Nar3,Amt2,Amt3,Amt4,Amt5,Amt6,UserName)" _
                & " select '3) Input','Purchase Rtn',sum(tblPurch.PurTaxableAmt)*-1,sum(tblPurch.PurTaxAmt3)*-1,sum(tblPurch.PurTaxAmt2)*-1,sum(tblPurch.PurTaxAmt)*-1,sum(tblPurch.PurBillAmt)*-1,'" & gUserName & "'" _
                & " From tblPurch where  " & x1 & ""
    
    
    '--Credit Note Purchase (GST)
    x1 = " tblIntSale.VType = 'PV'  and (tblIntSale.VDt)>= ('" & Format(w_Fr, "yyyy/mm/dd") & "') and (tblIntSale.VDt) <= ('" & Format(w_To, "yyyy/mm/dd") & "') and tblIntSale.CgstAmt+tblIntSale.SgstAmt+tblIntSale.IgstAmt > 0  and tblIntSale.Vfirm = '" & gCCode & "'"  '
    gCn.Execute " Insert Into TmpGentbl (Nar1,Nar3,Amt2,Amt3,Amt4,Amt5,Amt6,UserName)" _
                & " select '3) Input','Credit Note',sum(tblIntSale.Interest),sum(tblIntSale.IgstAmt),sum(tblIntSale.CgstAmt),sum(tblIntSale.SgstAmt),sum(tblIntSale.VAmt),'" & gUserName & "'" _
                & " From tblIntSale where  " & x1 & ""
    
    '--Debit Note Purchase (GST)
    x1 = " tblIntSale.VType = 'SV'  and (tblIntSale.VDt)>= ('" & Format(w_Fr, "yyyy/mm/dd") & "') and (tblIntSale.VDt) <= ('" & Format(w_To, "yyyy/mm/dd") & "') and tblIntSale.CgstAmt+tblIntSale.SgstAmt+tblIntSale.IgstAmt > 0  and tblIntSale.Vfirm = '" & gCCode & "'"  '
    gCn.Execute " Insert Into TmpGentbl (Nar1,Nar3,Amt2,Amt3,Amt4,Amt5,Amt6,UserName)" _
                & " select '3) Input','Debit Note',sum(tblIntSale.Interest)*-1,sum(tblIntSale.IgstAmt)*-1,sum(tblIntSale.CgstAmt)*-1,sum(tblIntSale.SgstAmt)*-1,sum(tblIntSale.VAmt)*-1,'" & gUserName & "'" _
                & " From tblIntSale where  " & x1 & ""
    
    
    
    '---Input RCM
    x1 = " tblPurch.VType = 'PO' and tblPurch.PurBillDt>= ('" & Format(w_Fr, "yyyy/mm/dd") & "')  and tblPurch.PurBillDt<=('" & Format(w_To, "yyyy/mm/dd") & "') and tblPurch.PurRCMCgstAmt+tblPurch.PurRCMSgstAmt+tblPurch.PurRCMIgstAmt > 0 and tblPurch.Vfirm = '" & gCCode & "'  "  '
                 
    gCn.Execute " Insert Into TmpGentbl (Nar1,Nar3,Amt2,Amt3,Amt4,Amt5,Amt6,UserName)" _
                             & " select '3) Input','RCM',sum(tblPurch.PurSubAmt),sum(tblPurch.PurRCMIgstAmt),sum(tblPurch.PurRCMCgstAmt),sum(tblPurch.PurRCMSgstAmt),sum(tblPurch.PurBillAmt),'" & gUserName & "'" _
                             & " From tblPurch where " & x1 & ""
    
    
    gCn.Execute " Insert Into TmpGentbl (Nar5,Nar1,Nar3,Amt2,Amt3,Amt4,Amt5,Amt6,UserName)" _
                             & " select 'B','RCM Pay','RCM',sum(tblPurch.PurSubAmt),sum(tblPurch.PurRCMIgstAmt),sum(tblPurch.PurRCMCgstAmt),sum(tblPurch.PurRCMSgstAmt),sum(tblPurch.PurBillAmt),'" & gUserName & "'" _
                             & " From tblPurch where " & x1 & ""



End If
gCn.CommitTrans

End Function

Public Function GProcDailyReport(w_Fr As Date, w_To As Date, FSr As Integer, TSr As Integer, w_PurBkFr As Date)
              Dim ij As Integer, Y1 As String, X As String
              Dim a12 As String
              Dim tmp1 As Recordset
              Dim tmp2 As Recordset
                gCmd.CommandText = "PrcPrepareDailyReport"
                gCmd.Parameters.Refresh
                gCmd.Parameters("@Fr_dt") = Format(w_Fr, "yyyy/mm/dd")
                gCmd.Parameters("@To_dt") = Format(w_To, "yyyy/mm/dd")
                gCmd.Parameters("@UCode") = gUserName
                gCmd.Parameters("@CYear") = gCYear
                gCmd.Parameters("@FSr") = FSr
                gCmd.Parameters("@TSr") = TSr
                gCmd.Parameters("@PurBookFrom") = Format(w_PurBkFr, "yyyy/mm/dd")
                
                gCmd.Execute
                X = "select Distinct TmpGentbl.itcode,TmpGentbl.Vfirm,'X',1, " _
                    & " tblMastAccount.AcCode,tblMastAccount.AcOurCode,tblMastAccount.Acsrno,tblMastAccount.AcName,tblMastAccount.AcRemark,tblMastItem.itname,tblMastItem.ItShort,tblMastItem.Cones,tblMastItem.BookingQty,tblMastItem.OfferQty,tblMastItem.OfferRate,tblMastItem.ItemRate,iif(dbo.tblMastItem.IsExGSTInDailyRpt = 0,'G',''),tblMastItem.ItStdpack,tblMastItem.RateTpInDailyRpt,tblMastItem.MillRate, " _
                    & " '" & gUserName & "' from TmpGentbl,tblMastItem,tblMastAccount " _
                    & " where UserName = '" & gUserName & "' and TmpGentbl.ItCode = tblMastItem.ItCode and tblMastItem.ItMillCode = tblMastAccount.AcCode " _
                    & " and tblMastAccount.Acsrno >= " & FSr & " and tblMastAccount.Acsrno <= " & TSr & " "
                
                gCn.Execute "insert into tmpGentbl2 (ItCode,Vfirm,Nar3,Amt8,PartyCode,MillCode,vsrno,Nar2,Nar10,Nar1,Nar5,Amt5,Amt13,Amt18,Amt19,Amt3,Nar12,Amt6,VTYPE,Amt20, UserName) " & X
                
                X = "select tblMastBillSerial.Vfirm,tblMastItem.ItCode,'X',1,tblMastAccount.AcCode,tblMastAccount.AcOurCode,tblMastAccount.AcSrNo,tblMastAccount.AcName,tblMastAccount.AcRemark,tblMastItem.itname,tblMastItem.ItShort,tblMastItem.Cones,tblMastItem.BookingQty,tblMastItem.OfferQty,tblMastItem.OfferRate,tblMastItem.ItemRate,iif(dbo.tblMastItem.IsExGSTInDailyRpt = 0,'G',''),tblMastItem.ItStdpack,tblMastItem.RateTpInDailyRpt,tblMastItem.MillRate,'" & gUserName & "' " _
                    & " From tblMastItem, tblMastAccount,tblMastBillSerial where tblMastItem.ItMillCode  = tblMastAccount.AcCode and tblMastItem.ItCode not in (select ItCode  from TmpGentbl where username = '" & gUserName & "') " _
                    & " and tblMastItem.InDailyReport = 0 and tblMastAccount.Acsrno > 0 and tblMastAccount.AcCode = tblMastBillSerial.MillCode " _
                    & " and tblMastAccount.Acsrno >= " & FSr & " and tblMastAccount.Acsrno <= " & TSr & " "
                
                gCn.Execute "insert into tmpGentbl2 (Vfirm,ItCode,Nar3,Amt8,PartyCode,MillCode,VsrNo,Nar2,Nar10,Nar1,Nar5,Amt5,Amt13,Amt18,Amt19,Amt3,Nar12,Amt6,VTYPE,Amt20, UserName) " & X
                
                
                '--For Trade Purchase Stock
               
                X = "Select *,MillCode+Nar1 AS a1 from TmpGentbl where  Nar3 = 'H'  and  UserName = '" & gUserName & "'  order by MillCode,Nar1,Nar6,Nar4 "
                
                If (GProcRstOpen(tmp1, X, "R") > 0) Then
                    Y1 = "select *,MillCode+Nar1 AS a2 from tmpGentbl2 where username='" & gUserName & "' order by MillCode,Nar1,Nar6,Nar4  "
                    If (GProcRstOpen(tmp2, Y1, "O") > 0) Then
                        tmp2.MoveFirst
                        Do While (Not tmp2.EOF)
                           a12 = tmp2!A2
                            tmp1.MoveFirst
                            tmp1.Find "A1='" & a12 & "'"
                            Do While (Not tmp1.EOF) And (Not tmp2.EOF)
                               If tmp1!A1 <> a12 Or tmp2!A2 <> a12 Then Exit Do
                               tmp2!Amt2 = tmp1.Fields("Amt2")
                               tmp2!Nar13 = tmp1.Fields("Nar4")
                               tmp2!Nar8 = tmp1.Fields("Nar6") & vbNullString
                               tmp2.Update
                               tmp2.MoveNext
                               tmp1.MoveNext
                            Loop
                            Do While (Not tmp1.EOF)
                               If tmp1!A1 <> a12 Then Exit Do
                               gCn.Execute "insert into tmpGentbl2 (Amt8,Vfirm,MillCode,Vsrno,Nar2,Nar10,Amt2,Nar1,Nar5,Nar13,Nar8,Username) values (2,'" & tmp1.Fields("Vfirm") & "','" & tmp1.Fields("MillCode") & "'," & tmp1.Fields("Vsrno") & ",'" & tmp1.Fields("Nar2") & "','" & tmp1.Fields("Nar10") & "'," & tmp1.Fields("Amt2") & ",'" & tmp1.Fields("Nar1") & "','" & tmp1.Fields("Nar5") & "','" & tmp1.Fields("Nar4") & "','" & IIf(IsNull(tmp1.Fields("Nar6")), "", tmp1.Fields("Nar6")) & "','" & gUserName & "')"
                               tmp1.MoveNext
                            Loop
                            Do While (Not tmp2.EOF)
                               If tmp2!A2 <> a12 Then Exit Do
                               tmp2.MoveNext
                            Loop
                        Loop
                    End If
                End If
                '--For Pending Purchase Booking
                X = "Select *,MillCode+Nar1 AS a1 from TmpGentbl where Nar3 = 'C'  and  UserName = '" & gUserName & "' order by MillCode,Nar1 "
                If (GProcRstOpen(tmp1, X, "R") > 0) Then
                    Y1 = "select *,MillCode+Nar1 AS a2 from tmpGentbl2 where username='" & gUserName & "' order by MillCode,Nar1  "
                    If (GProcRstOpen(tmp2, Y1, "O") > 0) Then
                        tmp2.MoveFirst
                        Do While (Not tmp2.EOF)
                           a12 = tmp2!A2
                            tmp1.MoveFirst
                            tmp1.Find "A1='" & a12 & "'"
                            Do While (Not tmp1.EOF) And (Not tmp2.EOF)
                               If tmp1!A1 <> a12 Or tmp2!A2 <> a12 Then Exit Do
                               tmp2!Amt10 = tmp1.Fields("Amt15")
                               tmp2.Update
                               tmp2.MoveNext
                               tmp1.MoveNext
                            Loop
                            Do While (Not tmp1.EOF)
                               If tmp1!A1 <> a12 Then Exit Do
                               gCn.Execute "insert into tmpGentbl2 (Amt8,Vfirm,MillCode,VsrNo,Nar2,Nar10,Amt10,Nar1,Nar5,Username) values (2,'" & tmp1.Fields("Vfirm") & "','" & tmp1.Fields("MillCode") & "'," & tmp1.Fields("Vsrno") & ",'" & tmp1.Fields("Nar2") & "','" & tmp1.Fields("Nar10") & "'," & tmp1.Fields("Amt15") & ",'" & tmp1.Fields("Nar1") & "','" & tmp1.Fields("Nar5") & "','" & gUserName & "')"
                               tmp1.MoveNext
                            Loop
                            Do While (Not tmp2.EOF)
                               If tmp2!A2 <> a12 Then Exit Do
                               tmp2.MoveNext
                            Loop
                        Loop
                    End If
                End If
                '-- Pending Party Booking
                X = "Select *,MillCode+VFirm+Nar1 AS a1 from TmpGentbl where Nar3 = 'J'  and  UserName = '" & gUserName & "' order by MillCode,Vfirm,Nar1 "
                
                If (GProcRstOpen(tmp1, X, "R") > 0) Then
                    Y1 = "select *,MillCode+VFirm+Nar1 AS a2 from tmpGentbl2 where  vfirm <> '' and   username='" & gUserName & "' order by MillCode,Nar1  "
                    If (GProcRstOpen(tmp2, Y1, "O") > 0) Then
                        tmp2.MoveFirst
                        Do While (Not tmp2.EOF)
                           a12 = tmp2!A2
                            tmp1.MoveFirst
                            tmp1.Find "A1='" & a12 & "'"
                            Do While (Not tmp1.EOF) And (Not tmp2.EOF)
                               If tmp1!A1 <> a12 Or tmp2!A2 <> a12 Then Exit Do
                               tmp2!Amt15 = tmp1.Fields("Amt15")
                               tmp2.Update
                               tmp2.MoveNext
                               tmp1.MoveNext
                            Loop
                            Do While (Not tmp1.EOF)
                               If tmp1!A1 <> a12 Then Exit Do
                               gCn.Execute "insert into tmpGentbl2 (Amt8,Vfirm,MillCode,VsrNo,Nar2,Nar10,Amt15,Nar1,Nar5,Username) values (2,'" & tmp1.Fields("Vfirm") & "','" & tmp1.Fields("MillCode") & "'," & tmp1.Fields("Vsrno") & ",'" & tmp1.Fields("Nar2") & "','" & tmp1.Fields("Nar10") & "'," & tmp1.Fields("Amt15") & ",'" & tmp1.Fields("Nar1") & "','" & tmp1.Fields("Nar5") & "','" & gUserName & "')"
                               tmp1.MoveNext
                            Loop
                            Do While (Not tmp2.EOF)
                               If tmp2!A2 <> a12 Then Exit Do
                               tmp2.MoveNext
                            Loop
                        Loop
                    End If
                End If
                
                
                '--For Pending Mill Booking
                X = "Select *,MillCode+VFirm+Nar1 AS a1 from TmpGentbl where Nar3 = 'F'  and  UserName = '" & gUserName & "' order by MillCode,Vfirm,Nar1 "
                
                If (GProcRstOpen(tmp1, X, "R") > 0) Then
                    Y1 = "select *,MillCode+VFirm+Nar1 AS a2 from tmpGentbl2 where vfirm <> '' and  username='" & gUserName & "' order by MillCode,Nar1  "
                    If (GProcRstOpen(tmp2, Y1, "O") > 0) Then
                        tmp2.MoveFirst
                        Do While (Not tmp2.EOF)
                           a12 = tmp2!A2
                            tmp1.MoveFirst
                            tmp1.Find "A1='" & a12 & "'"
                            Do While (Not tmp1.EOF) And (Not tmp2.EOF)
                               If tmp1!A1 <> a12 Or tmp2!A2 <> a12 Then Exit Do
                               tmp2!Amt10 = tmp1.Fields("Amt10")
                               tmp2.Update
                               tmp2.MoveNext
                               tmp1.MoveNext
                            Loop
                            Do While (Not tmp1.EOF)
                               If tmp1!A1 <> a12 Then Exit Do
                               gCn.Execute "insert into tmpGentbl2 (Amt8,Vfirm,MillCode,VsrNo,Nar2,Nar10,Amt10,Nar1,Nar5,Username) values (2,'" & tmp1.Fields("Vfirm") & "','" & tmp1.Fields("MillCode") & "'," & tmp1.Fields("Vsrno") & ",'" & tmp1.Fields("Nar2") & "','" & tmp1.Fields("Nar10") & "'," & tmp1.Fields("Amt10") & ",'" & tmp1.Fields("Nar1") & "','" & tmp1.Fields("Nar5") & "','" & gUserName & "')"
                               tmp1.MoveNext
                            Loop
                            Do While (Not tmp2.EOF)
                               If tmp2!A2 <> a12 Then Exit Do
                               tmp2.MoveNext
                            Loop
                        Loop
                    End If
                End If
                
                
                
                '--For Booking
                X = "Select *,MillCode+VFirm+Nar1 AS a1 from TmpGentbl where Nar3 = 'A'  and  UserName = '" & gUserName & "' order by MillCode,Vfirm,Nar1 "
                
                If (GProcRstOpen(tmp1, X, "R") > 0) Then
                    Y1 = "select *,MillCode+VFirm+Nar1 AS a2 from tmpGentbl2 where vfirm <> '' and  username='" & gUserName & "' order by MillCode,Vfirm,Nar1  "
                    If (GProcRstOpen(tmp2, Y1, "O") > 0) Then
                        tmp2.MoveFirst
                        Do While (Not tmp2.EOF)
                           a12 = tmp2!A2
                            tmp1.MoveFirst
                            tmp1.Find "A1='" & a12 & "'"
                            Do While (Not tmp1.EOF) And (Not tmp2.EOF)
                               If tmp1!A1 <> a12 Or tmp2!A2 <> a12 Then Exit Do
                               tmp2!Amt12 = tmp1.Fields("Amt12")
                               tmp2!Amt14 = tmp1.Fields("Amt14")
                               tmp2.Update
                               tmp2.MoveNext
                               tmp1.MoveNext
                            Loop
                            Do While (Not tmp1.EOF)
                               If tmp1!A1 <> a12 Then Exit Do
                               gCn.Execute "insert into tmpGentbl2 (Amt8,Vfirm,VFirmName,MillCode,VSrNo,Nar2,Nar10,Amt12,Nar1,Nar5,Username,Amt14) values (2,'" & tmp1.Fields("Vfirm") & "','" & tmp1.Fields("VFirmName") & "','" & tmp1.Fields("MillCode") & "'," & tmp1.Fields("Vsrno") & ",'" & tmp1.Fields("Nar2") & "','" & tmp1.Fields("Nar10") & "'," & tmp1.Fields("Amt12") & ",'" & tmp1.Fields("Nar1") & "','" & tmp1.Fields("Nar5") & "','" & gUserName & "', " & tmp1.Fields("Amt14") & ")"
                               tmp1.MoveNext
                            Loop
                            Do While (Not tmp2.EOF)
                               If tmp2!A2 <> a12 Then Exit Do
                               tmp2.MoveNext
                            Loop
                        Loop
                    End If
                End If

                '--For Transit Goods
                X = "Select *,MillCode+VFirm+Nar1 AS a1 from TmpGentbl where amt11 > 0 and  UserName = '" & gUserName & "' order by MillCode,Vfirm,Nar1 "
                If (GProcRstOpen(tmp1, X, "R") > 0) Then
                    Y1 = "select *,MillCode+VFirm+Nar1 AS a2 from tmpGentbl2 where vfirm <> '' and  username='" & gUserName & "' order by MillCode,Vfirm,Nar1"
                    If (GProcRstOpen(tmp2, Y1, "O") > 0) Then
                        tmp2.MoveFirst
                        Do While (Not tmp2.EOF)
                           a12 = tmp2!A2
                            tmp1.MoveFirst
                            tmp1.Find "A1='" & a12 & "'"
                            Do While (Not tmp1.EOF) And (Not tmp2.EOF)
                               If tmp1!A1 <> a12 Or tmp2!A2 <> a12 Then Exit Do
                               tmp2!Amt11 = tmp1.Fields("Amt11")
                               tmp2!dt1 = Format(tmp1.Fields("Dt1"), "yyyy/mm/dd")
                               tmp2.Update
                               tmp2.MoveNext
                               tmp1.MoveNext
                            Loop
                            Do While (Not tmp1.EOF)
                               If tmp1!A1 <> a12 Then Exit Do
                               gCn.Execute "insert into tmpGentbl2 (Amt8,Vfirm,VFirmName,MillCode,VSrNo,Nar2,Nar10,Amt11,Nar1,Nar5,Username,Dt1) values (2,'" & tmp1.Fields("Vfirm") & "','" & tmp1.Fields("VFirmName") & "','" & tmp1.Fields("MillCode") & "'," & tmp1.Fields("Vsrno") & ",'" & tmp1.Fields("Nar2") & "','" & tmp1.Fields("Nar10") & "'," & tmp1.Fields("Amt11") & ",'" & tmp1.Fields("Nar1") & "','" & tmp1.Fields("Nar5") & "','" & gUserName & "', '" & Format(tmp1.Fields("Dt1"), "yyyy/mm/dd") & "')"
                               tmp1.MoveNext
                            Loop
                            Do While (Not tmp2.EOF)
                               If tmp2!A2 <> a12 Then Exit Do
                               tmp2.MoveNext
                            Loop
                        Loop
                    End If
                End If
                '--For Daily Sale Rate+Qty
                X = "Select *,MillCode+VFirm+Nar1 AS a1 from TmpGentbl where amt12 > 0 and nar3 = 'B' and  UserName = '" & gUserName & "' order by MillCode,Vfirm,Nar1 "
                If (GProcRstOpen(tmp1, X, "R") > 0) Then
                    Y1 = "select *,MillCode+VFirm+Nar1 AS a2 from tmpGentbl2 where vfirm <> '' and  username='" & gUserName & "' order by MillCode,Vfirm,Nar1 "
                    If (GProcRstOpen(tmp2, Y1, "O") > 0) Then
                        tmp2.MoveFirst
                        Do While (Not tmp2.EOF)
                           a12 = tmp2!A2
                            tmp1.MoveFirst
                            tmp1.Find "A1='" & a12 & "'"
                            Do While (Not tmp1.EOF) And (Not tmp2.EOF)
                               If tmp1!A1 <> a12 Or tmp2!A2 <> a12 Then Exit Do
                               tmp2!Amt12 = tmp1.Fields("Amt12")
                               tmp2!Amt14 = tmp1.Fields("Amt14")
                               tmp2.Update
                               tmp2.MoveNext
                               tmp1.MoveNext
                            Loop
                            Do While (Not tmp1.EOF)
                               If tmp1!A1 <> a12 Then Exit Do
                               gCn.Execute "insert into tmpGentbl2 (Amt8,Vfirm,VFirmName,MillCode,VSrNo,Nar2,Nar10,Amt12,Nar1,Nar5,Username,Amt14) values (2,'" & tmp1.Fields("Vfirm") & "','" & tmp1.Fields("VFirmName") & "','" & tmp1.Fields("MillCode") & "'," & tmp1.Fields("Vsrno") & ",'" & tmp1.Fields("Nar2") & "','" & tmp1.Fields("Nar10") & "'," & tmp1.Fields("Amt12") & ",'" & tmp1.Fields("Nar1") & "','" & tmp1.Fields("Nar5") & "','" & gUserName & "', " & tmp1.Fields("Amt14") & ")"
                               tmp1.MoveNext
                            Loop
                            Do While (Not tmp2.EOF)
                               If tmp2!A2 <> a12 Then Exit Do
                               tmp2.MoveNext
                            Loop
                        Loop
                    End If
                End If
                
                
                '--For Daily Purchase Rate+Qty
                X = "Select *,MillCode+VFirm+Nar1 AS a1 from TmpGentbl where amt12 > 0 and nar3 = 'L' and  UserName = '" & gUserName & "' order by MillCode,Vfirm,Nar1 "
                If (GProcRstOpen(tmp1, X, "R") > 0) Then
                    Y1 = "select *,MillCode+VFirm+Nar1 AS a2 from tmpGentbl2 where vfirm <> '' and  username='" & gUserName & "' order by MillCode,Vfirm,Nar1 "
                    If (GProcRstOpen(tmp2, Y1, "O") > 0) Then
                        tmp2.MoveFirst
                        Do While (Not tmp2.EOF)
                           a12 = tmp2!A2
                            tmp1.MoveFirst
                            tmp1.Find "A1='" & a12 & "'"
                            Do While (Not tmp1.EOF) And (Not tmp2.EOF)
                               If tmp1!A1 <> a12 Or tmp2!A2 <> a12 Then Exit Do
                               tmp2!Amt4 = tmp1.Fields("Amt12")
                               tmp2!Amt9 = tmp1.Fields("Amt14")
                               tmp2.Update
                               tmp2.MoveNext
                               tmp1.MoveNext
                            Loop
                            Do While (Not tmp1.EOF)
                               If tmp1!A1 <> a12 Then Exit Do
                               gCn.Execute "insert into tmpGentbl2 (Amt8,Vfirm,VFirmName,MillCode,VSrNo,Nar2,Nar10,Amt4,Nar1,Nar5,Username,Amt9) values (2,'" & tmp1.Fields("Vfirm") & "','" & tmp1.Fields("VFirmName") & "','" & tmp1.Fields("MillCode") & "'," & tmp1.Fields("Vsrno") & ",'" & tmp1.Fields("Nar2") & "','" & tmp1.Fields("Nar10") & "'," & tmp1.Fields("Amt12") & ",'" & tmp1.Fields("Nar1") & "','" & tmp1.Fields("Nar5") & "','" & gUserName & "', " & tmp1.Fields("Amt14") & ")"
                               tmp1.MoveNext
                            Loop
                            Do While (Not tmp2.EOF)
                               If tmp2!A2 <> a12 Then Exit Do
                               tmp2.MoveNext
                            Loop
                        Loop
                    End If
                End If
                
                '--For Daily Party Booking Rate+Qty
                X = "Select *,MillCode+VFirm+Nar1 AS a1 from TmpGentbl where amt17 > 0 and nar3 = 'M' and  UserName = '" & gUserName & "' order by MillCode,Vfirm,Nar1 "
                If (GProcRstOpen(tmp1, X, "R") > 0) Then
                    Y1 = "select *,MillCode+VFirm+Nar1 AS a2 from tmpGentbl2 where vfirm <> '' and  username='" & gUserName & "' order by MillCode,Vfirm,Nar1 "
                    If (GProcRstOpen(tmp2, Y1, "O") > 0) Then
                        tmp2.MoveFirst
                        Do While (Not tmp2.EOF)
                           a12 = tmp2!A2
                            tmp1.MoveFirst
                            tmp1.Find "A1='" & a12 & "'"
                            Do While (Not tmp1.EOF) And (Not tmp2.EOF)
                               If tmp1!A1 <> a12 Or tmp2!A2 <> a12 Then Exit Do
                               tmp2!Amt16 = tmp1.Fields("Amt16")
                               tmp2!amt17 = tmp1.Fields("Amt17")
                               tmp2.Update
                               tmp2.MoveNext
                               tmp1.MoveNext
                            Loop
                            Do While (Not tmp1.EOF)
                               If tmp1!A1 <> a12 Then Exit Do
                               gCn.Execute "insert into tmpGentbl2 (Amt8,Vfirm,VFirmName,MillCode,VSrNo,Nar2,Nar10,Amt16,Nar1,Nar5,Username,Amt17) values (2,'" & tmp1.Fields("Vfirm") & "','" & tmp1.Fields("VFirmName") & "','" & tmp1.Fields("MillCode") & "'," & tmp1.Fields("Vsrno") & ",'" & tmp1.Fields("Nar2") & "','" & tmp1.Fields("Nar10") & "'," & tmp1.Fields("Amt16") & ",'" & tmp1.Fields("Nar1") & "','" & tmp1.Fields("Nar5") & "','" & gUserName & "', " & tmp1.Fields("Amt17") & ")"
                               tmp1.MoveNext
                            Loop
                            Do While (Not tmp2.EOF)
                               If tmp2!A2 <> a12 Then Exit Do
                               tmp2.MoveNext
                            Loop
                        Loop
                    End If
                End If
                '--For O/s
                X = "Select *,MillCode+VFirm AS a1 from TmpGentbl where nar3 = 'K' and  UserName = '" & gUserName & "' order by MillCode,Vfirm "
                If (GProcRstOpen(tmp1, X, "R") > 0) Then
                    tmp1.MoveFirst
                    Do While (Not tmp1.EOF)
                       gCn.Execute "insert into tmpGentbl2 (Amt8,Vfirm,VFirmName,MillCode,VSrNo,Nar2,Nar10,Vdt,Nar7,Nar4,Amt7,Username,nar1) values (2,'" & tmp1.Fields("Vfirm") & "','" & tmp1.Fields("VFirmName") & "','" & tmp1.Fields("MillCode") & "'," & tmp1.Fields("Vsrno") & ",'" & tmp1.Fields("Nar2") & "','" & tmp1.Fields("Nar10") & "','" & Format(tmp1.Fields("Vdt"), "yyyy/mm/dd") & "','" & tmp1.Fields("Nar7") & "','" & tmp1.Fields("Nar4") & "',  " & tmp1.Fields("Amt7") & ",  '" & gUserName & "','Z')"
                       tmp1.MoveNext
                    Loop
                End If
                '--For O/s
'                X = "Select *,MillCode+VFirm AS a1 from TmpGentbl where nar3 = 'K' and  UserName = '" & gUserName & "' order by MillCode,Vfirm "
'                If (GProcRstOpen(tmp1, X, "R") > 0) Then
'                    Y1 = "select *,MillCode+VFirm AS a2 from tmpGentbl2 where vfirm <> '' and  username='" & gUserName & "' order by MillCode,Vfirm "
'                    If (GProcRstOpen(tmp2, Y1, "O") > 0) Then
'                        tmp2.MoveFirst
'                        Do While (Not tmp2.EOF)
'                           a12 = tmp2!A2
'                            tmp1.MoveFirst
'                            tmp1.Find "A1='" & a12 & "'"
'                            Do While (Not tmp1.EOF) And (Not tmp2.EOF)
'                               If tmp1!A1 <> a12 Or tmp2!A2 <> a12 Then Exit Do
'                               tmp2!Vdt = Format(tmp1.Fields("Vdt"), "yyyy/mm/dd")
'                               tmp2!Nar4 = tmp1.Fields("Nar4")
'                               tmp2!Amt7 = tmp1.Fields("Amt7")
'                               tmp2.Update
'                               tmp2.MoveNext
'                               tmp1.MoveNext
'                            Loop
'                            Do While (Not tmp1.EOF)
'                               If tmp1!A1 <> a12 Then Exit Do
'                               gCn.Execute "insert into tmpGentbl2 (vsrno,Vfirm,VFirmName,MillCode,Nar2,Vdt,Nar4,Amt7,Username) values (2,'" & tmp1.Fields("Vfirm") & "','" & tmp1.Fields("VFirmName") & "','" & tmp1.Fields("MillCode") & "','" & tmp1.Fields("Nar2") & "','" & Format(tmp1.Fields("Vdt"), "yyyy/mm/dd") & "','" & tmp1.Fields("Nar4") & "',  " & tmp1.Fields("Amt7") & ",  '" & gUserName & "')"
'                               tmp1.MoveNext
'                            Loop
'                            Do While (Not tmp2.EOF)
'                               If tmp2!A2 <> a12 Then Exit Do
'                               tmp2.MoveNext
'                            Loop
'                        Loop
'                    End If
'                End If
'
'
'
'
'
'    gCn.Execute "update tmpGentbl2 set nar1 = 'Z' where username = '" & gUserName & "' and nar1 is null"


End Function
Public Function GProcPurpareGSTR9(w_Fr As Date, w_To As Date, OnPurBillDt As Boolean, OnPurRecDt As Boolean)
Dim X As String
Dim x1 As String
Dim i As Integer
Dim j As Long
Dim tmp1 As Recordset
Dim tmp2 As Recordset
Dim flnm As String
Dim RecYN As Long
Dim excel_app As Excel.Application
Dim workbook As Excel.workbook
Dim sheet_name As String
Dim sheet As Excel.Worksheet
Dim header_range As Excel.Range
Dim a As Long


flnm = "D:\GSTR9-" + gCCode + "-" + CStr(Year(gCYSDate)) + "-" + CStr(Year(gCYEDate)) + ".xlsx"
FileCopy gReportPath & "\GSTR9-Blank.xlsx", flnm
             
gCn.BeginTrans
gCn.Execute "delete from tmpGenTbl where username='" & gUserName & "'"
gCn.Execute "delete from tmpGenTbl2 where username='" & gUserName & "'"
If gCYear >= 2021 Then
   gCn.Execute "update tblMastItem Set ItHsn2 = 'X' where isnull(ItHsn2,'') = ''"
Else
   gCn.Execute "update tblMastItem Set ItHsn = 'X' where isnull(ItHsn,'') = ''"
End If
gCn.CommitTrans
'--For Total yarn Sale R.D.
gCn.BeginTrans

If gCIsDepotFirm = 1 Then
  x1 = " tblSale.VType in ('SD')  and tblSale.vdt>='" & Format(w_Fr, "yyyy/mm/dd") & ")  and tblSale.vdt<='" & Format(w_To, "yyyy/mm/dd") & "'  and tblSale.SlTaxAmt+tblSale.SlTaxAmt2+tblSale.SlTaxAmt3 > 0 and len(tblMastAccount_Pty.AcGSTIN) > 3 and tblSale.Vfirm = '" & gCCode & "'"
Else
  x1 = " tblSale.VType in ('SY','SC','SB','RY')  and tblSale.vdt>='" & Format(w_Fr, "yyyy/mm/dd") & "'  and tblSale.vdt<='" & Format(w_To, "yyyy/mm/dd") & "'  and tblSale.SlTaxAmt+tblSale.SlTaxAmt2+tblSale.SlTaxAmt3 > 0 and len(tblMastAccount_Pty.AcGSTIN) > 3 and tblSale.Vfirm = '" & gCCode & "'"
End If

If gCYear >= 2021 Then
    X = "select 'A','Sales',tblMastAccount.acname,tblMastItem.ItHsn2,tblSale.SlTaxRate+tblSale.SlTaxRate2+tblSale.SlTaxRate3," _
        & " Sum (tblSale.SlTaxableAmt), Sum(SlTaxAmt2), Sum(SlTaxAmt), Sum(SlTaxAmt3), Sum(SlBillAmt),'" & gUserName & "'" _
        & " from tblSale,tblSaleSub,tblMastItem,tblMastAccount,tblMastAccount tblMastAccount_Pty " _
        & " where tblSale.SlAcCrCode = tblMastAccount.AcCode and tblSaleSub.SlSubItCode = tblMastItem.ItCode and tblSale.Vno = tblSaleSub.Vno and " _
        & " tblSale.Vtype = tblSaleSub.Vtype and tblSale.Vyear = tblSaleSub.Vyear and tblSale.Vfirm = tblSaleSub.Vfirm  and " & x1 & " " _
        & " and tblSale.SlAcDrCode = tblMastAccount_Pty.AcCode  " _
        & " group by tblMastAccount.acname,tblMastItem.ItHsn2,tblSale.SlTaxRate+tblSale.SlTaxRate2+tblSale.SlTaxRate3"

Else
    X = "select 'A','Sales',tblMastAccount.acname,tblMastItem.ItHsn,tblSale.SlTaxRate+tblSale.SlTaxRate2+tblSale.SlTaxRate3," _
        & " Sum (tblSale.SlTaxableAmt), Sum(SlTaxAmt2), Sum(SlTaxAmt), Sum(SlTaxAmt3), Sum(SlBillAmt),'" & gUserName & "'" _
        & " from tblSale,tblSaleSub,tblMastItem,tblMastAccount,tblMastAccount tblMastAccount_Pty " _
        & " where tblSale.SlAcCrCode = tblMastAccount.AcCode and tblSaleSub.SlSubItCode = tblMastItem.ItCode and tblSale.Vno = tblSaleSub.Vno and " _
        & " tblSale.Vtype = tblSaleSub.Vtype and tblSale.Vyear = tblSaleSub.Vyear and tblSale.Vfirm = tblSaleSub.Vfirm  and " & x1 & " " _
        & " and tblSale.SlAcDrCode = tblMastAccount_Pty.AcCode  " _
        & " group by tblMastAccount.acname,tblMastItem.ItHsn,tblSale.SlTaxRate+tblSale.SlTaxRate2+tblSale.SlTaxRate3"
End If
gCn.Execute " Insert Into TmpGentbl (VType,Nar1,Nar2,Nar3,Amt1,Amt2,Amt3,Amt4,Amt5,Amt6,UserName)" & X
gCn.CommitTrans

Set excel_app = New Excel.Application
excel_app.Visible = False
Set workbook = excel_app.Workbooks.Open(filename:=flnm)
RecYN = GProcRstOpen(tmp1, "select * from TmpGentbl where vtype = 'A' and TmpGentbl.UserName ='" & gUserName & "'", "R", gCn)
If RecYN > 0 Then
   Set sheet = workbook.Sheets.Item("Yearly")
   sheet.cells(2, 3) = gCName
   sheet.cells(5, 4) = gCGStin
   sheet.cells(8, 4) = CStr(Year(gCYSDate)) + "-" + CStr(Year(gCYEDate))
   
   
   If (sheet Is Nothing) Then
       MsgBox "Error Yearly Sheet"
   End If
   a = 16
   With tmp1
        .MoveFirst
        Do While Not .EOF
           sheet.cells(a, 2) = "Sales"
           sheet.cells(a, 3) = !Nar2
           If Left(!Nar3, 2) = "99" Then
              sheet.cells(a, 4) = "Services"
           Else
              sheet.cells(a, 4) = "Goods"
           End If
           sheet.cells(a, 5) = "Registered"
           sheet.cells(a, 6) = "Tax Invoice"
           sheet.cells(a, 7) = !Nar3
           sheet.cells(a, 8) = !Amt1
           gCn.BeginTrans
           For i = 1 To 12
               Set sheet = workbook.Sheets.Item("" & Left(MonthName(i), 3) & "")
               If gCYear >= 2021 Then
                  X = "select 'A','1',tblMastAccount.acname,tblMastItem.ItHsn2,tblSale.SlTaxRate+tblSale.SlTaxRate2+tblSale.SlTaxRate3," _
                      & " Sum (tblSale.SlTaxableAmt), Sum(SlTaxAmt2), Sum(SlTaxAmt), Sum(SlTaxAmt3), Sum(SlBillAmt)," & i & ",'" & gUserName & "'" _
                      & " from tblSale,tblSaleSub,tblMastItem,tblMastAccount,tblMastAccount tblMastAccount_Pty " _
                      & " where tblSale.SlAcCrCode = tblMastAccount.AcCode and tblSaleSub.SlSubItCode = tblMastItem.ItCode and tblSale.Vno = tblSaleSub.Vno and " _
                      & " tblSale.Vtype = tblSaleSub.Vtype and tblSale.Vyear = tblSaleSub.Vyear and tblSale.Vfirm = tblSaleSub.Vfirm and " & x1 & " " _
                      & " and tblSale.SlAcDrCode = tblMastAccount_Pty.AcCode and month(tblSale.Vdt) = " & i & " and  tblMastItem.ItHsn2 = '" & tmp1.Fields("Nar3") & "' and tblSale.SlTaxRate+tblSale.SlTaxRate2+tblSale.SlTaxRate3 = " & tmp1.Fields("Amt1") & " " _
                      & " and tblMastAccount.AcName = '" & tmp1.Fields("Nar2") & "' " _
                      & " group by tblMastAccount.acname,tblMastItem.ItHsn2,tblSale.SlTaxRate+tblSale.SlTaxRate2+tblSale.SlTaxRate3"
               
               Else
                  X = "select 'A','1',tblMastAccount.acname,tblMastItem.ItHsn,tblSale.SlTaxRate+tblSale.SlTaxRate2+tblSale.SlTaxRate3," _
                      & " Sum (tblSale.SlTaxableAmt), Sum(SlTaxAmt2), Sum(SlTaxAmt), Sum(SlTaxAmt3), Sum(SlBillAmt)," & i & ",'" & gUserName & "'" _
                      & " from tblSale,tblSaleSub,tblMastItem,tblMastAccount,tblMastAccount tblMastAccount_Pty " _
                      & " where tblSale.SlAcCrCode = tblMastAccount.AcCode and tblSaleSub.SlSubItCode = tblMastItem.ItCode and tblSale.Vno = tblSaleSub.Vno and " _
                      & " tblSale.Vtype = tblSaleSub.Vtype and tblSale.Vyear = tblSaleSub.Vyear and tblSale.Vfirm = tblSaleSub.Vfirm and " & x1 & " " _
                      & " and tblSale.SlAcDrCode = tblMastAccount_Pty.AcCode and month(tblSale.Vdt) = " & i & " and  tblMastItem.ItHsn = '" & tmp1.Fields("Nar3") & "' and tblSale.SlTaxRate+tblSale.SlTaxRate2+tblSale.SlTaxRate3 = " & tmp1.Fields("Amt1") & " " _
                      & " and tblMastAccount.AcName = '" & tmp1.Fields("Nar2") & "' " _
                      & " group by tblMastAccount.acname,tblMastItem.ItHsn,tblSale.SlTaxRate+tblSale.SlTaxRate2+tblSale.SlTaxRate3"

               End If
               gCn.Execute " Insert Into TmpGentbl2 (VType,Nar1,Nar2,Nar3,Amt1,Amt2,Amt3,Amt4,Amt5,Amt6,Vno,UserName)" & X

               j = GProcRstOpen(tmp2, X, "R", gCn)
               If j > 0 Then
                  sheet.cells(a, 10) = IIf(IsNull(tmp2.Fields(5)), 0, tmp2.Fields(5))
                  sheet.cells(a, 11) = IIf(IsNull(tmp2.Fields(6)), 0, tmp2.Fields(6))
                  sheet.cells(a, 12) = IIf(IsNull(tmp2.Fields(7)), 0, tmp2.Fields(7))
                  sheet.cells(a, 13) = IIf(IsNull(tmp2.Fields(8)), 0, tmp2.Fields(8))
                  sheet.cells(a, 14) = 0
               Else
                  sheet.cells(a, 10) = 0
                  sheet.cells(a, 11) = 0
                  sheet.cells(a, 12) = 0
                  sheet.cells(a, 13) = 0
                  sheet.cells(a, 14) = 0
               End If

           Next i
           gCn.CommitTrans
           Set sheet = workbook.Sheets.Item("Yearly")
           a = a + 1
           If a = 35 Then a = 37
           If a = 39 Then a = 45
           If a = 60 Then a = 67
           If a = 127 Then a = 133
           If a = 171 Then a = 173
           tmp1.MoveNext
        Loop
   End With
End If

'--For Total yarn Sale U.R.D.
gCn.BeginTrans

If gCIsDepotFirm = 1 Then
  x1 = " tblSale.VType in ('SD')  and tblSale.vdt>='" & Format(w_Fr, "yyyy/mm/dd") & ")  and tblSale.vdt<='" & Format(w_To, "yyyy/mm/dd") & "'  and tblSale.SlTaxAmt+tblSale.SlTaxAmt2+tblSale.SlTaxAmt3 > 0 and len(tblMastAccount_Pty.AcGSTIN) <= 3 and tblSale.Vfirm = '" & gCCode & "'"
Else
  x1 = " tblSale.VType in ('SY','SC','SB')  and tblSale.vdt>='" & Format(w_Fr, "yyyy/mm/dd") & "'  and tblSale.vdt<='" & Format(w_To, "yyyy/mm/dd") & "'  and tblSale.SlTaxAmt+tblSale.SlTaxAmt2+tblSale.SlTaxAmt3 > 0 and len(tblMastAccount_Pty.AcGSTIN) <= 3  and tblSale.Vfirm = '" & gCCode & "'   "
End If
If gCYear >= 2021 Then
    X = "select 'B','Sales',tblMastAccount.acname,tblMastItem.ItHsn2,tblSale.SlTaxRate+tblSale.SlTaxRate2+tblSale.SlTaxRate3," _
        & " Sum (tblSale.SlTaxableAmt), Sum(SlTaxAmt2), Sum(SlTaxAmt), Sum(SlTaxAmt3), Sum(SlBillAmt),'" & gUserName & "'" _
        & " from tblSale,tblSaleSub,tblMastItem,tblMastAccount,tblMastAccount tblMastAccount_Pty " _
        & " where tblSale.SlAcCrCode = tblMastAccount.AcCode and tblSaleSub.SlSubItCode = tblMastItem.ItCode and tblSale.Vno = tblSaleSub.Vno and " _
        & " tblSale.Vtype = tblSaleSub.Vtype and tblSale.Vyear = tblSaleSub.Vyear and tblSale.Vfirm = tblSaleSub.Vfirm  and " & x1 & " " _
        & " and tblSale.SlAcDrCode = tblMastAccount_Pty.AcCode  " _
        & " group by tblMastAccount.acname,tblMastItem.ItHsn2,tblSale.SlTaxRate+tblSale.SlTaxRate2+tblSale.SlTaxRate3"

Else
    X = "select 'B','Sales',tblMastAccount.acname,tblMastItem.ItHsn,tblSale.SlTaxRate+tblSale.SlTaxRate2+tblSale.SlTaxRate3," _
        & " Sum (tblSale.SlTaxableAmt), Sum(SlTaxAmt2), Sum(SlTaxAmt), Sum(SlTaxAmt3), Sum(SlBillAmt),'" & gUserName & "'" _
        & " from tblSale,tblSaleSub,tblMastItem,tblMastAccount,tblMastAccount tblMastAccount_Pty " _
        & " where tblSale.SlAcCrCode = tblMastAccount.AcCode and tblSaleSub.SlSubItCode = tblMastItem.ItCode and tblSale.Vno = tblSaleSub.Vno and " _
        & " tblSale.Vtype = tblSaleSub.Vtype and tblSale.Vyear = tblSaleSub.Vyear and tblSale.Vfirm = tblSaleSub.Vfirm  and " & x1 & " " _
        & " and tblSale.SlAcDrCode = tblMastAccount_Pty.AcCode  " _
        & " group by tblMastAccount.acname,tblMastItem.ItHsn,tblSale.SlTaxRate+tblSale.SlTaxRate2+tblSale.SlTaxRate3"
End If
gCn.Execute " Insert Into TmpGentbl (VType,Nar1,Nar2,Nar3,Amt1,Amt2,Amt3,Amt4,Amt5,Amt6,UserName)" & X
gCn.CommitTrans

RecYN = GProcRstOpen(tmp1, "select * from TmpGentbl where vtype = 'B' and TmpGentbl.UserName ='" & gUserName & "'", "R", gCn)
If RecYN > 0 Then
   Set sheet = workbook.Sheets.Item("Yearly")
   If (sheet Is Nothing) Then
       MsgBox "Error Yearly Sheet"
   End If
   With tmp1
        .MoveFirst
        Do While Not .EOF
           sheet.cells(a, 2) = "Sales"
           sheet.cells(a, 3) = !Nar2
           If Left(!Nar3, 2) = "99" Then
              sheet.cells(a, 4) = "Services"
           Else
              sheet.cells(a, 4) = "Goods"
           End If
           sheet.cells(a, 5) = "Unregistered"
           sheet.cells(a, 6) = "Tax Invoice"
           sheet.cells(a, 7) = !Nar3
           sheet.cells(a, 8) = !Amt1
           gCn.BeginTrans
           For i = 1 To 12
               Set sheet = workbook.Sheets.Item("" & Left(MonthName(i), 3) & "")
               If gCYear >= 2021 Then
                  X = "select 'B','1',tblMastAccount.acname,tblMastItem.ItHsn2,tblSale.SlTaxRate+tblSale.SlTaxRate2+tblSale.SlTaxRate3," _
                      & " Sum (tblSale.SlTaxableAmt), Sum(SlTaxAmt2), Sum(SlTaxAmt), Sum(SlTaxAmt3), Sum(SlBillAmt)," & i & ",'" & gUserName & "'" _
                      & " from tblSale,tblSaleSub,tblMastItem,tblMastAccount,tblMastAccount tblMastAccount_Pty " _
                      & " where tblSale.SlAcCrCode = tblMastAccount.AcCode and tblSaleSub.SlSubItCode = tblMastItem.ItCode and tblSale.Vno = tblSaleSub.Vno and " _
                      & " tblSale.Vtype = tblSaleSub.Vtype and tblSale.Vyear = tblSaleSub.Vyear and tblSale.Vfirm = tblSaleSub.Vfirm and " & x1 & " " _
                        & " and tblSale.SlAcDrCode = tblMastAccount_Pty.AcCode and month(tblSale.Vdt) = " & i & " and  tblMastItem.ItHsn2 = '" & tmp1.Fields("Nar3") & "' and tblSale.SlTaxRate+tblSale.SlTaxRate2+tblSale.SlTaxRate3 = " & tmp1.Fields("Amt1") & " " _
                        & " and tblMastAccount.AcName = '" & tmp1.Fields("Nar2") & "' " _
                        & " group by tblMastAccount.acname,tblMastItem.ItHsn2,tblSale.SlTaxRate+tblSale.SlTaxRate2+tblSale.SlTaxRate3"
               
               Else
                  X = "select 'B','1',tblMastAccount.acname,tblMastItem.ItHsn,tblSale.SlTaxRate+tblSale.SlTaxRate2+tblSale.SlTaxRate3," _
                      & " Sum (tblSale.SlTaxableAmt), Sum(SlTaxAmt2), Sum(SlTaxAmt), Sum(SlTaxAmt3), Sum(SlBillAmt)," & i & ",'" & gUserName & "'" _
                      & " from tblSale,tblSaleSub,tblMastItem,tblMastAccount,tblMastAccount tblMastAccount_Pty " _
                      & " where tblSale.SlAcCrCode = tblMastAccount.AcCode and tblSaleSub.SlSubItCode = tblMastItem.ItCode and tblSale.Vno = tblSaleSub.Vno and " _
                      & " tblSale.Vtype = tblSaleSub.Vtype and tblSale.Vyear = tblSaleSub.Vyear and tblSale.Vfirm = tblSaleSub.Vfirm and " & x1 & " " _
                        & " and tblSale.SlAcDrCode = tblMastAccount_Pty.AcCode and month(tblSale.Vdt) = " & i & " and  tblMastItem.ItHsn = '" & tmp1.Fields("Nar3") & "' and tblSale.SlTaxRate+tblSale.SlTaxRate2+tblSale.SlTaxRate3 = " & tmp1.Fields("Amt1") & " " _
                        & " and tblMastAccount.AcName = '" & tmp1.Fields("Nar2") & "' " _
                        & " group by tblMastAccount.acname,tblMastItem.ItHsn,tblSale.SlTaxRate+tblSale.SlTaxRate2+tblSale.SlTaxRate3"
               End If
               gCn.Execute " Insert Into TmpGentbl2 (VType,Nar1,Nar2,Nar3,Amt1,Amt2,Amt3,Amt4,Amt5,Amt6,Vno,UserName)" & X
               j = GProcRstOpen(tmp2, X, "R", gCn)
               If j > 0 Then
                  sheet.cells(a, 10) = IIf(IsNull(tmp2.Fields(5)), 0, tmp2.Fields(5))
                  sheet.cells(a, 11) = IIf(IsNull(tmp2.Fields(6)), 0, tmp2.Fields(6))
                  sheet.cells(a, 12) = IIf(IsNull(tmp2.Fields(7)), 0, tmp2.Fields(7))
                  sheet.cells(a, 13) = IIf(IsNull(tmp2.Fields(8)), 0, tmp2.Fields(8))
                  sheet.cells(a, 14) = 0
               Else
                  sheet.cells(a, 10) = 0
                  sheet.cells(a, 11) = 0
                  sheet.cells(a, 12) = 0
                  sheet.cells(a, 13) = 0
                  sheet.cells(a, 14) = 0
               End If
           Next i
           gCn.CommitTrans
           Set sheet = workbook.Sheets.Item("Yearly")
           a = a + 1
           If a = 35 Then a = 37
           If a = 39 Then a = 45
           If a = 60 Then a = 67
           If a = 127 Then a = 133
           If a = 171 Then a = 173
           tmp1.MoveNext
        Loop
   End With
End If
'---End Yarn Sale

'--For Total Late Pay Int. R.D.
gCn.BeginTrans

x1 = " tblIntSale.VType = 'SI'  and tblIntSale.vdt>='" & Format(w_Fr, "yyyy/mm/dd") & "'  and tblIntSale.vdt<='" & Format(w_To, "yyyy/mm/dd") & "'  and tblIntSale.CgstAmt+tblIntSale.SgstAmt+tblIntSale.IgstAmt > 0 and len(tblMastAccount_Pty.AcGSTIN) > 3 and tblIntSale.Vfirm = '" & gCCode & "'  and tblIntSale.vdt < '2021/01/01'    "

If gCYear >= 2021 Then
    X = "select 'C','Sales',tblMastAccount.acname,tblMastItem.ItHsn2,tblIntSale.CgstRt+tblIntSale.SgstRt+tblIntSale.IgstRt," _
        & " Sum (tblIntSale.Interest), Sum(CgstAmt), Sum(SgstAmt), Sum(IgstAmt), Sum(Vamt),'" & gUserName & "'" _
        & " from tblIntSale,tblMastItem,tblMastAccount,tblMastAccount tblMastAccount_Pty " _
        & " where tblIntSale.VCrAcCode = tblMastAccount.AcCode and tblIntSale.ItCode = tblMastItem.ItCode " _
        & " and " & x1 & " " _
        & " and tblIntSale.VDrAcCode = tblMastAccount_Pty.AcCode  " _
        & " group by tblMastAccount.acname,tblMastItem.ItHsn2,tblIntSale.CgstRt+tblIntSale.SgstRt+tblIntSale.IgstRt"
Else
    X = "select 'C','Sales',tblMastAccount.acname,tblMastItem.ItHsn,tblIntSale.CgstRt+tblIntSale.SgstRt+tblIntSale.IgstRt," _
        & " Sum (tblIntSale.Interest), Sum(CgstAmt), Sum(SgstAmt), Sum(IgstAmt), Sum(Vamt),'" & gUserName & "'" _
        & " from tblIntSale,tblMastItem,tblMastAccount,tblMastAccount tblMastAccount_Pty " _
        & " where tblIntSale.VCrAcCode = tblMastAccount.AcCode and tblIntSale.ItCode = tblMastItem.ItCode " _
        & " and " & x1 & " " _
        & " and tblIntSale.VDrAcCode = tblMastAccount_Pty.AcCode  " _
        & " group by tblMastAccount.acname,tblMastItem.ItHsn,tblIntSale.CgstRt+tblIntSale.SgstRt+tblIntSale.IgstRt"
End If
gCn.Execute " Insert Into TmpGentbl (VType,Nar1,Nar2,Nar3,Amt1,Amt2,Amt3,Amt4,Amt5,Amt6,UserName)" & X
gCn.CommitTrans

RecYN = GProcRstOpen(tmp1, "select * from TmpGentbl where vtype = 'C' and TmpGentbl.UserName ='" & gUserName & "'", "R", gCn)
If RecYN > 0 Then
   Set sheet = workbook.Sheets.Item("Yearly")
   If (sheet Is Nothing) Then
       MsgBox "Error Yearly Sheet"
   End If
   With tmp1
        .MoveFirst
        Do While Not .EOF
           sheet.cells(a, 2) = "Sales"
           sheet.cells(a, 3) = !Nar2
           If Left(!Nar3, 2) = "99" Then
              sheet.cells(a, 4) = "Services"
           Else
              sheet.cells(a, 4) = "Goods"
           End If
           sheet.cells(a, 5) = "Registered"
           sheet.cells(a, 6) = "Tax Invoice"
           sheet.cells(a, 7) = !Nar3
           sheet.cells(a, 8) = !Amt1
           gCn.BeginTrans
           For i = 1 To 12
               Set sheet = workbook.Sheets.Item("" & Left(MonthName(i), 3) & "")
               If gCYear >= 2021 Then
                  X = "select 'C','1',tblMastAccount.acname,tblMastItem.ItHsn2,tblIntSale.CgstRt+tblIntSale.SgstRt+tblIntSale.IgstRt," _
                      & " Sum (tblIntSale.Interest), Sum(CgstAmt), Sum(SgstAmt), Sum(IgstAmt), Sum(Vamt)," & i & ",'" & gUserName & "'" _
                      & " from tblIntSale,tblMastItem,tblMastAccount,tblMastAccount tblMastAccount_Pty " _
                      & " where tblIntSale.VCrAcCode = tblMastAccount.AcCode and tblIntSale.ItCode = tblMastItem.ItCode " _
                      & " and " & x1 & " " _
                      & " and tblIntSale.VDrAcCode = tblMastAccount_Pty.AcCode and month(tblIntSale.Vdt) = " & i & " and  tblMastItem.ItHsn2 = '" & tmp1.Fields("Nar3") & "' and tblIntSale.CgstRt+tblIntSale.SgstRt+tblIntSale.IgstRt = " & tmp1.Fields("Amt1") & " " _
                      & " and tblMastAccount.AcName = '" & tmp1.Fields("Nar2") & "' " _
                      & " group by tblMastAccount.acname,tblMastItem.ItHsn2,tblIntSale.CgstRt+tblIntSale.SgstRt+tblIntSale.IgstRt"
               Else
                  X = "select 'C','1',tblMastAccount.acname,tblMastItem.ItHsn,tblIntSale.CgstRt+tblIntSale.SgstRt+tblIntSale.IgstRt," _
                      & " Sum (tblIntSale.Interest), Sum(CgstAmt), Sum(SgstAmt), Sum(IgstAmt), Sum(Vamt)," & i & ",'" & gUserName & "'" _
                      & " from tblIntSale,tblMastItem,tblMastAccount,tblMastAccount tblMastAccount_Pty " _
                      & " where tblIntSale.VCrAcCode = tblMastAccount.AcCode and tblIntSale.ItCode = tblMastItem.ItCode " _
                      & " and " & x1 & " " _
                      & " and tblIntSale.VDrAcCode = tblMastAccount_Pty.AcCode and month(tblIntSale.Vdt) = " & i & " and  tblMastItem.ItHsn = '" & tmp1.Fields("Nar3") & "' and tblIntSale.CgstRt+tblIntSale.SgstRt+tblIntSale.IgstRt = " & tmp1.Fields("Amt1") & " " _
                      & " and tblMastAccount.AcName = '" & tmp1.Fields("Nar2") & "' " _
                      & " group by tblMastAccount.acname,tblMastItem.ItHsn,tblIntSale.CgstRt+tblIntSale.SgstRt+tblIntSale.IgstRt"
               End If
               gCn.Execute " Insert Into TmpGentbl2 (VType,Nar1,Nar2,Nar3,Amt1,Amt2,Amt3,Amt4,Amt5,Amt6,Vno,UserName)" & X

               j = GProcRstOpen(tmp2, X, "R", gCn)
               If j > 0 Then
                  sheet.cells(a, 10) = IIf(IsNull(tmp2.Fields(5)), 0, tmp2.Fields(5))
                  sheet.cells(a, 11) = IIf(IsNull(tmp2.Fields(6)), 0, tmp2.Fields(6))
                  sheet.cells(a, 12) = IIf(IsNull(tmp2.Fields(7)), 0, tmp2.Fields(7))
                  sheet.cells(a, 13) = IIf(IsNull(tmp2.Fields(8)), 0, tmp2.Fields(8))
                  sheet.cells(a, 14) = 0
               Else
                  sheet.cells(a, 10) = 0
                  sheet.cells(a, 11) = 0
                  sheet.cells(a, 12) = 0
                  sheet.cells(a, 13) = 0
                  sheet.cells(a, 14) = 0
               End If

           Next i
           gCn.CommitTrans
           Set sheet = workbook.Sheets.Item("Yearly")
           a = a + 1
           If a = 35 Then a = 37
           If a = 39 Then a = 45
           If a = 60 Then a = 67
           If a = 127 Then a = 133
           If a = 171 Then a = 173
           tmp1.MoveNext
        Loop
   End With
End If

'--For Total yarn Purchase R.D.
gCn.BeginTrans
   
x1 = " tblPurch.VType = 'PY' and tblPurch.PurBillDt>= '" & Format(w_Fr, "yyyy/mm/dd") & "'  and tblPurch.PurBillDt<= '" & Format(w_To, "yyyy/mm/dd") & "' and tblPurch.PurTaxAmt3+tblPurch.PurTaxAmt2+tblPurch.PurTaxAmt > 0 and len(tblMastAccount_Pty.AcGSTIN) > 2 and tblPurch.Vfirm = '" & gCCode & "' "
If gCYear >= 2021 Then
    X = "select 'E','Purchases',tblMastAccount.acname,tblMastItem.ItHsn2,tblPurch.PurTaxRate+tblPurch.PurTaxRate2+tblPurch.PurTaxRate3," _
        & " Sum (PurTaxableAmt), Sum(PurTaxAmt2), Sum(PurTaxAmt), Sum(PurTaxAmt3), Sum(PurBillAmt),'" & gUserName & "'" _
        & " from tblPurch,tblPurchSub,tblMastItem,tblMastAccount,tblMastAccount tblMastAccount_Pty " _
        & " where tblPurch.PurAcDrCode = tblMastAccount.AcCode and tblPurchSub.PurSubItCode = tblMastItem.ItCode and tblPurch.Vno = tblPurchSub.Vno and " _
        & " tblPurch.Vtype = tblPurchSub.Vtype and tblPurch.Vyear = tblPurchSub.Vyear and tblPurch.Vfirm = tblPurchSub.Vfirm and " & x1 & " " _
        & " and tblPurch.PurAcCrCode = tblMastAccount_Pty.AcCode and PurSubItSrNo = 1 " _
        & " group by tblMastAccount.acname,tblMastItem.ItHsn2,tblPurch.PurTaxRate+tblPurch.PurTaxRate2+tblPurch.PurTaxRate3"
Else
    X = "select 'E','Purchases',tblMastAccount.acname,tblMastItem.ItHsn,tblPurch.PurTaxRate+tblPurch.PurTaxRate2+tblPurch.PurTaxRate3," _
        & " Sum (PurTaxableAmt), Sum(PurTaxAmt2), Sum(PurTaxAmt), Sum(PurTaxAmt3), Sum(PurBillAmt),'" & gUserName & "'" _
        & " from tblPurch,tblPurchSub,tblMastItem,tblMastAccount,tblMastAccount tblMastAccount_Pty " _
        & " where tblPurch.PurAcDrCode = tblMastAccount.AcCode and tblPurchSub.PurSubItCode = tblMastItem.ItCode and tblPurch.Vno = tblPurchSub.Vno and " _
        & " tblPurch.Vtype = tblPurchSub.Vtype and tblPurch.Vyear = tblPurchSub.Vyear and tblPurch.Vfirm = tblPurchSub.Vfirm and " & x1 & " " _
        & " and tblPurch.PurAcCrCode = tblMastAccount_Pty.AcCode and PurSubItSrNo = 1 " _
        & " group by tblMastAccount.acname,tblMastItem.ItHsn,tblPurch.PurTaxRate+tblPurch.PurTaxRate2+tblPurch.PurTaxRate3"
End If
gCn.Execute " Insert Into TmpGentbl (VType,Nar1,Nar2,Nar3,Amt1,Amt2,Amt3,Amt4,Amt5,Amt6,UserName)" & X
gCn.CommitTrans

RecYN = GProcRstOpen(tmp1, "select * from TmpGentbl where vtype = 'E' and TmpGentbl.UserName ='" & gUserName & "'", "R", gCn)
If RecYN > 0 Then
   Set sheet = workbook.Sheets.Item("Yearly")
   If (sheet Is Nothing) Then
       MsgBox "Error Yearly Sheet"
   End If
   With tmp1
        .MoveFirst
        Do While Not .EOF
           sheet.cells(a, 2) = "Purchases"
           sheet.cells(a, 3) = !Nar2
           If Left(!Nar3, 2) = "99" Then
              sheet.cells(a, 4) = "Services"
           Else
              sheet.cells(a, 4) = "Goods"
           End If
           sheet.cells(a, 5) = "Registered"
           sheet.cells(a, 6) = "Tax Invoice"
           sheet.cells(a, 7) = !Nar3
           sheet.cells(a, 8) = !Amt1
           gCn.BeginTrans
           For i = 1 To 12
               Set sheet = workbook.Sheets.Item("" & Left(MonthName(i), 3) & "")
               If gCYear >= 2021 Then
                  X = "select 'E','4',tblMastAccount.acname,tblMastItem.ItHsn2,tblPurch.PurTaxRate+tblPurch.PurTaxRate2+tblPurch.PurTaxRate3," _
                      & " Sum (PurTaxableAmt), Sum(PurTaxAmt2), Sum(PurTaxAmt), Sum(PurTaxAmt3), Sum(PurBillAmt)," & i & ",'" & gUserName & "'" _
                      & " from tblPurch,tblPurchSub,tblMastItem,tblMastAccount,tblMastAccount tblMastAccount_Pty " _
                      & " where tblPurch.PurAcDrCode = tblMastAccount.AcCode and tblPurchSub.PurSubItCode = tblMastItem.ItCode and tblPurch.Vno = tblPurchSub.Vno and " _
                      & " tblPurch.Vtype = tblPurchSub.Vtype and tblPurch.Vyear = tblPurchSub.Vyear and tblPurch.Vfirm = tblPurchSub.Vfirm and " & x1 & " and month(tblPurch.PurBillDt) = " & i & " and tblMastItem.ItHsn2 = '" & tmp1.Fields("Nar3") & "' and tblPurch.PurTaxRate+tblPurch.PurTaxRate2+tblPurch.PurTaxRate3 = " & tmp1.Fields("Amt1") & "" _
                      & " and tblPurch.PurAcCrCode = tblMastAccount_Pty.AcCode and PurSubItSrNo = 1 " _
                      & " and tblMastAccount.AcName = '" & tmp1.Fields("Nar2") & "' " _
                      & " group by tblMastAccount.acname,tblMastItem.ItHsn2,tblPurch.PurTaxRate+tblPurch.PurTaxRate2+tblPurch.PurTaxRate3"
               Else
                  X = "select 'E','4',tblMastAccount.acname,tblMastItem.ItHsn,tblPurch.PurTaxRate+tblPurch.PurTaxRate2+tblPurch.PurTaxRate3," _
                      & " Sum (PurTaxableAmt), Sum(PurTaxAmt2), Sum(PurTaxAmt), Sum(PurTaxAmt3), Sum(PurBillAmt)," & i & ",'" & gUserName & "'" _
                      & " from tblPurch,tblPurchSub,tblMastItem,tblMastAccount,tblMastAccount tblMastAccount_Pty " _
                      & " where tblPurch.PurAcDrCode = tblMastAccount.AcCode and tblPurchSub.PurSubItCode = tblMastItem.ItCode and tblPurch.Vno = tblPurchSub.Vno and " _
                      & " tblPurch.Vtype = tblPurchSub.Vtype and tblPurch.Vyear = tblPurchSub.Vyear and tblPurch.Vfirm = tblPurchSub.Vfirm and " & x1 & " and month(tblPurch.PurBillDt) = " & i & " and tblMastItem.ItHsn = '" & tmp1.Fields("Nar3") & "' and tblPurch.PurTaxRate+tblPurch.PurTaxRate2+tblPurch.PurTaxRate3 = " & tmp1.Fields("Amt1") & "" _
                      & " and tblPurch.PurAcCrCode = tblMastAccount_Pty.AcCode and PurSubItSrNo = 1 " _
                      & " and tblMastAccount.AcName = '" & tmp1.Fields("Nar2") & "' " _
                      & " group by tblMastAccount.acname,tblMastItem.ItHsn,tblPurch.PurTaxRate+tblPurch.PurTaxRate2+tblPurch.PurTaxRate3"
               End If
               gCn.Execute " Insert Into TmpGentbl2 (VType,Nar1,Nar2,Nar3,Amt1,Amt2,Amt3,Amt4,Amt5,Amt6,Vno,UserName)" & X
               j = GProcRstOpen(tmp2, X, "R", gCn)
               If j > 0 Then
                  sheet.cells(a, 10) = IIf(IsNull(tmp2.Fields(5)), 0, tmp2.Fields(5))
                  sheet.cells(a, 11) = IIf(IsNull(tmp2.Fields(6)), 0, tmp2.Fields(6))
                  sheet.cells(a, 12) = IIf(IsNull(tmp2.Fields(7)), 0, tmp2.Fields(7))
                  sheet.cells(a, 13) = IIf(IsNull(tmp2.Fields(8)), 0, tmp2.Fields(8))
                  sheet.cells(a, 14) = 0
               Else
                  sheet.cells(a, 10) = 0
                  sheet.cells(a, 11) = 0
                  sheet.cells(a, 12) = 0
                  sheet.cells(a, 13) = 0
                  sheet.cells(a, 14) = 0
               End If

           Next i
           gCn.CommitTrans
           Set sheet = workbook.Sheets.Item("Yearly")
           a = a + 1
           If a = 35 Then a = 37
           If a = 39 Then a = 45
           If a = 60 Then a = 67
           If a = 127 Then a = 133
           If a = 171 Then a = 173
           tmp1.MoveNext
        Loop
   End With
End If



'--For Total Other Purchase R.D.
gCn.BeginTrans
   
x1 = " tblPurch.VType = 'PO' and tblPurch.PurBillDt>= '" & Format(w_Fr, "yyyy/mm/dd") & "'  and tblPurch.PurBillDt<= '" & Format(w_To, "yyyy/mm/dd") & "' and tblPurch.PurTaxAmt3+tblPurch.PurTaxAmt2+tblPurch.PurTaxAmt > 0 and len(tblMastAccount_Pty.AcGSTIN) > 2 and tblPurch.Vfirm = '" & gCCode & "' "
If gCYear >= 2021 Then
    X = "select 'F','Indirect Expenses',tblMastAccount.acname,tblMastItem.ItHsn2,tblPurch.PurTaxRate+tblPurch.PurTaxRate2+tblPurch.PurTaxRate3," _
        & " Sum (PurTaxableAmt), Sum(PurTaxAmt2), Sum(PurTaxAmt), Sum(PurTaxAmt3), Sum(PurBillAmt),'" & gUserName & "'" _
        & " from tblPurch,tblPurchSub,tblMastItem,tblMastAccount,tblMastAccount tblMastAccount_Pty " _
        & " where tblPurch.PurAcDrCode = tblMastAccount.AcCode and tblPurchSub.PurSubItCode = tblMastItem.ItCode and tblPurch.Vno = tblPurchSub.Vno and " _
        & " tblPurch.Vtype = tblPurchSub.Vtype and tblPurch.Vyear = tblPurchSub.Vyear and tblPurch.Vfirm = tblPurchSub.Vfirm and " & x1 & " " _
        & " and tblPurch.PurAcCrCode = tblMastAccount_Pty.AcCode and PurSubItSrNo = 1 " _
        & " group by tblMastAccount.acname,tblMastItem.ItHsn2,tblPurch.PurTaxRate+tblPurch.PurTaxRate2+tblPurch.PurTaxRate3"
Else
    X = "select 'F','Indirect Expenses',tblMastAccount.acname,tblMastItem.ItHsn,tblPurch.PurTaxRate+tblPurch.PurTaxRate2+tblPurch.PurTaxRate3," _
        & " Sum (PurTaxableAmt), Sum(PurTaxAmt2), Sum(PurTaxAmt), Sum(PurTaxAmt3), Sum(PurBillAmt),'" & gUserName & "'" _
        & " from tblPurch,tblPurchSub,tblMastItem,tblMastAccount,tblMastAccount tblMastAccount_Pty " _
        & " where tblPurch.PurAcDrCode = tblMastAccount.AcCode and tblPurchSub.PurSubItCode = tblMastItem.ItCode and tblPurch.Vno = tblPurchSub.Vno and " _
        & " tblPurch.Vtype = tblPurchSub.Vtype and tblPurch.Vyear = tblPurchSub.Vyear and tblPurch.Vfirm = tblPurchSub.Vfirm and " & x1 & " " _
        & " and tblPurch.PurAcCrCode = tblMastAccount_Pty.AcCode and PurSubItSrNo = 1 " _
        & " group by tblMastAccount.acname,tblMastItem.ItHsn,tblPurch.PurTaxRate+tblPurch.PurTaxRate2+tblPurch.PurTaxRate3"
End If
gCn.Execute " Insert Into TmpGentbl (VType,Nar1,Nar2,Nar3,Amt1,Amt2,Amt3,Amt4,Amt5,Amt6,UserName)" & X
gCn.CommitTrans

RecYN = GProcRstOpen(tmp1, "select * from TmpGentbl where vtype = 'F' and TmpGentbl.UserName ='" & gUserName & "'", "R", gCn)
If RecYN > 0 Then
   Set sheet = workbook.Sheets.Item("Yearly")
   If (sheet Is Nothing) Then
       MsgBox "Error Yearly Sheet"
   End If
   With tmp1
        .MoveFirst
        Do While Not .EOF
           sheet.cells(a, 2) = "Indirect Expenses"
           sheet.cells(a, 3) = !Nar2
           If Left(!Nar3, 2) = "99" Then
              sheet.cells(a, 4) = "Services"
           Else
              sheet.cells(a, 4) = "Others"
           End If
           sheet.cells(a, 5) = "Registered"
           sheet.cells(a, 6) = "Tax Invoice"
           sheet.cells(a, 7) = !Nar3
           sheet.cells(a, 8) = !Amt1
           gCn.BeginTrans
           For i = 1 To 12
               Set sheet = workbook.Sheets.Item("" & Left(MonthName(i), 3) & "")
               If gCYear >= 2021 Then
                  X = "select 'F','4',tblMastAccount.acname,tblMastItem.ItHsn2,tblPurch.PurTaxRate+tblPurch.PurTaxRate2+tblPurch.PurTaxRate3," _
                      & " Sum (PurTaxableAmt), Sum(PurTaxAmt2), Sum(PurTaxAmt), Sum(PurTaxAmt3), Sum(PurBillAmt)," & i & ",'" & gUserName & "'" _
                      & " from tblPurch,tblPurchSub,tblMastItem,tblMastAccount,tblMastAccount tblMastAccount_Pty " _
                      & " where tblPurch.PurAcDrCode = tblMastAccount.AcCode and tblPurchSub.PurSubItCode = tblMastItem.ItCode and tblPurch.Vno = tblPurchSub.Vno and " _
                      & " tblPurch.Vtype = tblPurchSub.Vtype and tblPurch.Vyear = tblPurchSub.Vyear and tblPurch.Vfirm = tblPurchSub.Vfirm and " & x1 & " and month(tblPurch.PurBillDt) = " & i & " and tblMastItem.ItHsn2 = '" & tmp1.Fields("Nar3") & "' and tblPurch.PurTaxRate+tblPurch.PurTaxRate2+tblPurch.PurTaxRate3 = " & tmp1.Fields("Amt1") & "" _
                      & " and tblPurch.PurAcCrCode = tblMastAccount_Pty.AcCode and PurSubItSrNo = 1 " _
                      & " and tblMastAccount.AcName = '" & tmp1.Fields("Nar2") & "' " _
                      & " group by tblMastAccount.acname,tblMastItem.ItHsn2,tblPurch.PurTaxRate+tblPurch.PurTaxRate2+tblPurch.PurTaxRate3"
               Else
                  X = "select 'F','4',tblMastAccount.acname,tblMastItem.ItHsn,tblPurch.PurTaxRate+tblPurch.PurTaxRate2+tblPurch.PurTaxRate3," _
                      & " Sum (PurTaxableAmt), Sum(PurTaxAmt2), Sum(PurTaxAmt), Sum(PurTaxAmt3), Sum(PurBillAmt)," & i & ",'" & gUserName & "'" _
                      & " from tblPurch,tblPurchSub,tblMastItem,tblMastAccount,tblMastAccount tblMastAccount_Pty " _
                      & " where tblPurch.PurAcDrCode = tblMastAccount.AcCode and tblPurchSub.PurSubItCode = tblMastItem.ItCode and tblPurch.Vno = tblPurchSub.Vno and " _
                      & " tblPurch.Vtype = tblPurchSub.Vtype and tblPurch.Vyear = tblPurchSub.Vyear and tblPurch.Vfirm = tblPurchSub.Vfirm and " & x1 & " and month(tblPurch.PurBillDt) = " & i & " and tblMastItem.ItHsn = '" & tmp1.Fields("Nar3") & "' and tblPurch.PurTaxRate+tblPurch.PurTaxRate2+tblPurch.PurTaxRate3 = " & tmp1.Fields("Amt1") & "" _
                      & " and tblPurch.PurAcCrCode = tblMastAccount_Pty.AcCode and PurSubItSrNo = 1 " _
                      & " and tblMastAccount.AcName = '" & tmp1.Fields("Nar2") & "' " _
                      & " group by tblMastAccount.acname,tblMastItem.ItHsn,tblPurch.PurTaxRate+tblPurch.PurTaxRate2+tblPurch.PurTaxRate3"
               End If
               gCn.Execute " Insert Into TmpGentbl2 (VType,Nar1,Nar2,Nar3,Amt1,Amt2,Amt3,Amt4,Amt5,Amt6,Vno,UserName)" & X
               j = GProcRstOpen(tmp2, X, "R", gCn)
               If j > 0 Then
                  sheet.cells(a, 10) = IIf(IsNull(tmp2.Fields(5)), 0, tmp2.Fields(5))
                  sheet.cells(a, 11) = IIf(IsNull(tmp2.Fields(6)), 0, tmp2.Fields(6))
                  sheet.cells(a, 12) = IIf(IsNull(tmp2.Fields(7)), 0, tmp2.Fields(7))
                  sheet.cells(a, 13) = IIf(IsNull(tmp2.Fields(8)), 0, tmp2.Fields(8))
                  sheet.cells(a, 14) = 0
               Else
                  sheet.cells(a, 10) = 0
                  sheet.cells(a, 11) = 0
                  sheet.cells(a, 12) = 0
                  sheet.cells(a, 13) = 0
                  sheet.cells(a, 14) = 0
               End If

           Next i
           gCn.CommitTrans
           Set sheet = workbook.Sheets.Item("Yearly")
           a = a + 1
           If a = 35 Then a = 37
           If a = 39 Then a = 45
           If a = 60 Then a = 67
           If a = 127 Then a = 133
           If a = 171 Then a = 173
           tmp1.MoveNext
        Loop
   End With
End If



'--Debit Note
'--For Total Late Pay Int. R.D.
gCn.BeginTrans

x1 = " tblIntSale.VType = 'SI'  and tblIntSale.vdt>='" & Format(w_Fr, "yyyy/mm/dd") & "'  and tblIntSale.vdt<='" & Format(w_To, "yyyy/mm/dd") & "'  and tblIntSale.CgstAmt+tblIntSale.SgstAmt+tblIntSale.IgstAmt > 0 and len(tblMastAccount_Pty.AcGSTIN) > 3 and tblIntSale.Vfirm = '" & gCCode & "'  and tblIntSale.vdt >= '2021/01/01'    "

If gCYear >= 2021 Then
    X = "select 'G','Sales',tblMastAccount.acname,tblMastItem.ItHsn2,tblIntSale.CgstRt+tblIntSale.SgstRt+tblIntSale.IgstRt," _
        & " Sum (tblIntSale.Interest), Sum(CgstAmt), Sum(SgstAmt), Sum(IgstAmt), Sum(Vamt),'" & gUserName & "'" _
        & " from tblIntSale,tblMastItem,tblMastAccount,tblMastAccount tblMastAccount_Pty " _
        & " where tblIntSale.VCrAcCode = tblMastAccount.AcCode and tblIntSale.ItCode = tblMastItem.ItCode " _
        & " and " & x1 & " " _
        & " and tblIntSale.VDrAcCode = tblMastAccount_Pty.AcCode  " _
        & " group by tblMastAccount.acname,tblMastItem.ItHsn2,tblIntSale.CgstRt+tblIntSale.SgstRt+tblIntSale.IgstRt"
Else
    X = "select 'G','Sales',tblMastAccount.acname,tblMastItem.ItHsn,tblIntSale.CgstRt+tblIntSale.SgstRt+tblIntSale.IgstRt," _
        & " Sum (tblIntSale.Interest), Sum(CgstAmt), Sum(SgstAmt), Sum(IgstAmt), Sum(Vamt),'" & gUserName & "'" _
        & " from tblIntSale,tblMastItem,tblMastAccount,tblMastAccount tblMastAccount_Pty " _
        & " where tblIntSale.VCrAcCode = tblMastAccount.AcCode and tblIntSale.ItCode = tblMastItem.ItCode " _
        & " and " & x1 & " " _
        & " and tblIntSale.VDrAcCode = tblMastAccount_Pty.AcCode  " _
        & " group by tblMastAccount.acname,tblMastItem.ItHsn,tblIntSale.CgstRt+tblIntSale.SgstRt+tblIntSale.IgstRt"
End If
gCn.Execute " Insert Into TmpGentbl (VType,Nar1,Nar2,Nar3,Amt1,Amt2,Amt3,Amt4,Amt5,Amt6,UserName)" & X
gCn.CommitTrans

RecYN = GProcRstOpen(tmp1, "select * from TmpGentbl where vtype = 'G' and TmpGentbl.UserName ='" & gUserName & "'", "R", gCn)
If RecYN > 0 Then
   Set sheet = workbook.Sheets.Item("Yearly")
   If (sheet Is Nothing) Then
       MsgBox "Error Yearly Sheet"
   End If
   With tmp1
        .MoveFirst
        Do While Not .EOF
           sheet.cells(a, 2) = "Sales"
           sheet.cells(a, 3) = !Nar2
           If Left(!Nar3, 2) = "99" Then
              sheet.cells(a, 4) = "Services"
           Else
              sheet.cells(a, 4) = "Goods"
           End If
           sheet.cells(a, 5) = "Registered"
           sheet.cells(a, 6) = "Debit Note"
           sheet.cells(a, 7) = !Nar3
           sheet.cells(a, 8) = !Amt1
           gCn.BeginTrans
           For i = 1 To 12
               Set sheet = workbook.Sheets.Item("" & Left(MonthName(i), 3) & "")
               If gCYear >= 2021 Then
                  X = "select 'G','1',tblMastAccount.acname,tblMastItem.ItHsn2,tblIntSale.CgstRt+tblIntSale.SgstRt+tblIntSale.IgstRt," _
                      & " Sum (tblIntSale.Interest), Sum(CgstAmt), Sum(SgstAmt), Sum(IgstAmt), Sum(Vamt)," & i & ",'" & gUserName & "'" _
                      & " from tblIntSale,tblMastItem,tblMastAccount,tblMastAccount tblMastAccount_Pty " _
                      & " where tblIntSale.VCrAcCode = tblMastAccount.AcCode and tblIntSale.ItCode = tblMastItem.ItCode " _
                      & " and " & x1 & " " _
                      & " and tblIntSale.VDrAcCode = tblMastAccount_Pty.AcCode and month(tblIntSale.Vdt) = " & i & " and  tblMastItem.ItHsn2 = '" & tmp1.Fields("Nar3") & "' and tblIntSale.CgstRt+tblIntSale.SgstRt+tblIntSale.IgstRt = " & tmp1.Fields("Amt1") & " " _
                      & " and tblMastAccount.AcName = '" & tmp1.Fields("Nar2") & "' " _
                      & " group by tblMastAccount.acname,tblMastItem.ItHsn2,tblIntSale.CgstRt+tblIntSale.SgstRt+tblIntSale.IgstRt"
               Else
                  X = "select 'G','1',tblMastAccount.acname,tblMastItem.ItHsn,tblIntSale.CgstRt+tblIntSale.SgstRt+tblIntSale.IgstRt," _
                      & " Sum (tblIntSale.Interest), Sum(CgstAmt), Sum(SgstAmt), Sum(IgstAmt), Sum(Vamt)," & i & ",'" & gUserName & "'" _
                      & " from tblIntSale,tblMastItem,tblMastAccount,tblMastAccount tblMastAccount_Pty " _
                      & " where tblIntSale.VCrAcCode = tblMastAccount.AcCode and tblIntSale.ItCode = tblMastItem.ItCode " _
                      & " and " & x1 & " " _
                      & " and tblIntSale.VDrAcCode = tblMastAccount_Pty.AcCode and month(tblIntSale.Vdt) = " & i & " and  tblMastItem.ItHsn = '" & tmp1.Fields("Nar3") & "' and tblIntSale.CgstRt+tblIntSale.SgstRt+tblIntSale.IgstRt = " & tmp1.Fields("Amt1") & " " _
                      & " and tblMastAccount.AcName = '" & tmp1.Fields("Nar2") & "' " _
                      & " group by tblMastAccount.acname,tblMastItem.ItHsn,tblIntSale.CgstRt+tblIntSale.SgstRt+tblIntSale.IgstRt"
               End If
               gCn.Execute " Insert Into TmpGentbl2 (VType,Nar1,Nar2,Nar3,Amt1,Amt2,Amt3,Amt4,Amt5,Amt6,Vno,UserName)" & X

               j = GProcRstOpen(tmp2, X, "R", gCn)
               If j > 0 Then
                  sheet.cells(a, 10) = IIf(IsNull(tmp2.Fields(5)), 0, tmp2.Fields(5))
                  sheet.cells(a, 11) = IIf(IsNull(tmp2.Fields(6)), 0, tmp2.Fields(6))
                  sheet.cells(a, 12) = IIf(IsNull(tmp2.Fields(7)), 0, tmp2.Fields(7))
                  sheet.cells(a, 13) = IIf(IsNull(tmp2.Fields(8)), 0, tmp2.Fields(8))
                  sheet.cells(a, 14) = 0
               Else
                  sheet.cells(a, 10) = 0
                  sheet.cells(a, 11) = 0
                  sheet.cells(a, 12) = 0
                  sheet.cells(a, 13) = 0
                  sheet.cells(a, 14) = 0
               End If

           Next i
           gCn.CommitTrans
           Set sheet = workbook.Sheets.Item("Yearly")
           a = a + 1
           If a = 35 Then a = 37
           If a = 39 Then a = 45
           If a = 60 Then a = 67
           If a = 127 Then a = 133
           If a = 171 Then a = 173
           tmp1.MoveNext
        Loop
   End With
End If


'--Credit Note Purchase (GST)
gCn.BeginTrans

x1 = " tblIntSale.VType = 'PV'  and tblIntSale.vdt>='" & Format(w_Fr, "yyyy/mm/dd") & "'  and tblIntSale.vdt<='" & Format(w_To, "yyyy/mm/dd") & "'  and tblIntSale.CgstAmt+tblIntSale.SgstAmt+tblIntSale.IgstAmt > 0 and len(tblMastAccount_Pty.AcGSTIN) > 3 and tblIntSale.Vfirm = '" & gCCode & "'  "


X = "select 'H','Sales',tblMastAccount.acname,tblIntSale.CgstRt+tblIntSale.SgstRt+tblIntSale.IgstRt," _
    & " Sum (tblIntSale.Interest), Sum(CgstAmt), Sum(SgstAmt), Sum(IgstAmt), Sum(Vamt),'" & gUserName & "'" _
    & " from tblIntSale,tblMastAccount,tblMastAccount tblMastAccount_Pty " _
    & " where tblIntSale.VCrAcCode = tblMastAccount.AcCode " _
    & " and " & x1 & " " _
    & " and tblIntSale.VDrAcCode = tblMastAccount_Pty.AcCode  " _
    & " group by tblMastAccount.acname,tblIntSale.CgstRt+tblIntSale.SgstRt+tblIntSale.IgstRt"

gCn.Execute " Insert Into TmpGentbl (VType,Nar1,Nar2,Amt1,Amt2,Amt3,Amt4,Amt5,Amt6,UserName)" & X
gCn.CommitTrans

RecYN = GProcRstOpen(tmp1, "select * from TmpGentbl where vtype = 'H' and TmpGentbl.UserName ='" & gUserName & "'", "R", gCn)
If RecYN > 0 Then
   Set sheet = workbook.Sheets.Item("Yearly")
   If (sheet Is Nothing) Then
       MsgBox "Error Yearly Sheet"
   End If
   With tmp1
        .MoveFirst
        Do While Not .EOF
           sheet.cells(a, 2) = "Purchases"
           sheet.cells(a, 3) = !Nar2
           If Left(!Nar3, 2) = "99" Then
              sheet.cells(a, 4) = "Services"
           Else
              sheet.cells(a, 4) = "Goods"
           End If
           sheet.cells(a, 5) = "Registered"
           sheet.cells(a, 6) = "Credit Note"
           sheet.cells(a, 7) = !Nar3
           sheet.cells(a, 8) = !Amt1
           gCn.BeginTrans
           For i = 1 To 12
               Set sheet = workbook.Sheets.Item("" & Left(MonthName(i), 3) & "")
              
               X = "select 'H','1',tblMastAccount.acname,tblIntSale.CgstRt+tblIntSale.SgstRt+tblIntSale.IgstRt," _
                   & " Sum (tblIntSale.Interest), Sum(CgstAmt), Sum(SgstAmt), Sum(IgstAmt), Sum(Vamt)," & i & ",'" & gUserName & "'" _
                   & " from tblIntSale,tblMastAccount,tblMastAccount tblMastAccount_Pty " _
                   & " where tblIntSale.VCrAcCode = tblMastAccount.AcCode  " _
                   & " and " & x1 & " " _
                   & " and tblIntSale.VDrAcCode = tblMastAccount_Pty.AcCode and month(tblIntSale.Vdt) = " & i & " and tblIntSale.CgstRt+tblIntSale.SgstRt+tblIntSale.IgstRt = " & tmp1.Fields("Amt1") & " " _
                   & " and tblMastAccount.AcName = '" & tmp1.Fields("Nar2") & "' " _
                   & " group by tblMastAccount.acname,tblIntSale.CgstRt+tblIntSale.SgstRt+tblIntSale.IgstRt"

               gCn.Execute " Insert Into TmpGentbl2 (VType,Nar1,Nar2,Amt1,Amt2,Amt3,Amt4,Amt5,Amt6,Vno,UserName)" & X

               j = GProcRstOpen(tmp2, X, "R", gCn)
               If j > 0 Then
                  sheet.cells(a, 10) = IIf(IsNull(tmp2.Fields(5)), 0, tmp2.Fields(5))
                  sheet.cells(a, 11) = IIf(IsNull(tmp2.Fields(6)), 0, tmp2.Fields(6))
                  sheet.cells(a, 12) = IIf(IsNull(tmp2.Fields(7)), 0, tmp2.Fields(7))
                  sheet.cells(a, 13) = IIf(IsNull(tmp2.Fields(8)), 0, tmp2.Fields(8))
                  sheet.cells(a, 14) = 0
               Else
                  sheet.cells(a, 10) = 0
                  sheet.cells(a, 11) = 0
                  sheet.cells(a, 12) = 0
                  sheet.cells(a, 13) = 0
                  sheet.cells(a, 14) = 0
               End If

           Next i
           gCn.CommitTrans
           Set sheet = workbook.Sheets.Item("Yearly")
           a = a + 1
           If a = 35 Then a = 37
           If a = 39 Then a = 45
           If a = 60 Then a = 67
           If a = 127 Then a = 133
           If a = 171 Then a = 173
           tmp1.MoveNext
        Loop
   End With
End If

'--Debit Note Purchase (GST)

gCn.BeginTrans

x1 = " tblIntSale.VType = 'SV'  and tblIntSale.vdt>='" & Format(w_Fr, "yyyy/mm/dd") & "'  and tblIntSale.vdt<='" & Format(w_To, "yyyy/mm/dd") & "'  and tblIntSale.CgstAmt+tblIntSale.SgstAmt+tblIntSale.IgstAmt > 0 and len(tblMastAccount_Pty.AcGSTIN) > 3 and tblIntSale.Vfirm = '" & gCCode & "'  "


X = "select 'I','Sales',tblMastAccount.acname,tblIntSale.CgstRt+tblIntSale.SgstRt+tblIntSale.IgstRt," _
    & " Sum (tblIntSale.Interest), Sum(CgstAmt), Sum(SgstAmt), Sum(IgstAmt), Sum(Vamt),'" & gUserName & "'" _
    & " from tblIntSale,tblMastAccount,tblMastAccount tblMastAccount_Pty " _
    & " where tblIntSale.VCrAcCode = tblMastAccount.AcCode  " _
    & " and " & x1 & " " _
    & " and tblIntSale.VDrAcCode = tblMastAccount_Pty.AcCode  " _
    & " group by tblMastAccount.acname,tblIntSale.CgstRt+tblIntSale.SgstRt+tblIntSale.IgstRt"

gCn.Execute " Insert Into TmpGentbl (VType,Nar1,Nar2,Amt1,Amt2,Amt3,Amt4,Amt5,Amt6,UserName)" & X
gCn.CommitTrans

RecYN = GProcRstOpen(tmp1, "select * from TmpGentbl where vtype = 'I' and TmpGentbl.UserName ='" & gUserName & "'", "R", gCn)
If RecYN > 0 Then
   Set sheet = workbook.Sheets.Item("Yearly")
   If (sheet Is Nothing) Then
       MsgBox "Error Yearly Sheet"
   End If
   With tmp1
        .MoveFirst
        Do While Not .EOF
           sheet.cells(a, 2) = "Purchases"
           sheet.cells(a, 3) = !Nar2
           If Left(!Nar3, 2) = "99" Then
              sheet.cells(a, 4) = "Services"
           Else
              sheet.cells(a, 4) = "Goods"
           End If
           sheet.cells(a, 5) = "Registered"
           sheet.cells(a, 6) = "Debit Note"
           sheet.cells(a, 7) = !Nar3
           sheet.cells(a, 8) = !Amt1
           gCn.BeginTrans
           For i = 1 To 12
               Set sheet = workbook.Sheets.Item("" & Left(MonthName(i), 3) & "")
              
               X = "select 'I','1',tblMastAccount.acname,tblIntSale.CgstRt+tblIntSale.SgstRt+tblIntSale.IgstRt," _
                   & " Sum (tblIntSale.Interest), Sum(CgstAmt), Sum(SgstAmt), Sum(IgstAmt), Sum(Vamt)," & i & ",'" & gUserName & "'" _
                   & " from tblIntSale,tblMastAccount,tblMastAccount tblMastAccount_Pty " _
                   & " where tblIntSale.VCrAcCode = tblMastAccount.AcCode " _
                   & " and " & x1 & " " _
                   & " and tblIntSale.VDrAcCode = tblMastAccount_Pty.AcCode and month(tblIntSale.Vdt) = " & i & " and tblIntSale.CgstRt+tblIntSale.SgstRt+tblIntSale.IgstRt = " & tmp1.Fields("Amt1") & " " _
                   & " and tblMastAccount.AcName = '" & tmp1.Fields("Nar2") & "' " _
                   & " group by tblMastAccount.acname,tblIntSale.CgstRt+tblIntSale.SgstRt+tblIntSale.IgstRt"

               gCn.Execute " Insert Into TmpGentbl2 (VType,Nar1,Nar2,Amt1,Amt2,Amt3,Amt4,Amt5,Amt6,Vno,UserName)" & X

               j = GProcRstOpen(tmp2, X, "R", gCn)
               If j > 0 Then
                  sheet.cells(a, 10) = IIf(IsNull(tmp2.Fields(5)), 0, tmp2.Fields(5))
                  sheet.cells(a, 11) = IIf(IsNull(tmp2.Fields(6)), 0, tmp2.Fields(6))
                  sheet.cells(a, 12) = IIf(IsNull(tmp2.Fields(7)), 0, tmp2.Fields(7))
                  sheet.cells(a, 13) = IIf(IsNull(tmp2.Fields(8)), 0, tmp2.Fields(8))
                  sheet.cells(a, 14) = 0
               Else
                  sheet.cells(a, 10) = 0
                  sheet.cells(a, 11) = 0
                  sheet.cells(a, 12) = 0
                  sheet.cells(a, 13) = 0
                  sheet.cells(a, 14) = 0
               End If

           Next i
           gCn.CommitTrans
           Set sheet = workbook.Sheets.Item("Yearly")
           a = a + 1
           If a = 35 Then a = 37
           If a = 39 Then a = 45
           If a = 60 Then a = 67
           If a = 127 Then a = 133
           If a = 171 Then a = 173
           tmp1.MoveNext
        Loop
   End With
End If

'--Purchase Return
gCn.BeginTrans
   
x1 = " tblPurch.VType = 'VY' and tblPurch.vdt>= '" & Format(w_Fr, "yyyy/mm/dd") & "'  and tblPurch.vdt<= '" & Format(w_To, "yyyy/mm/dd") & "' and tblPurch.PurTaxAmt3+tblPurch.PurTaxAmt2+tblPurch.PurTaxAmt > 0 and len(tblMastAccount_Pty.AcGSTIN) > 2 and tblPurch.Vfirm = '" & gCCode & "' "
If gCYear >= 2021 Then
    X = "select 'K','Purchases',tblMastAccount.acname,tblMastItem.ItHsn2,tblPurch.PurTaxRate+tblPurch.PurTaxRate2+tblPurch.PurTaxRate3," _
        & " Sum (PurTaxableAmt), Sum(PurTaxAmt2), Sum(PurTaxAmt), Sum(PurTaxAmt3), Sum(PurBillAmt),'" & gUserName & "'" _
        & " from tblPurch,tblPurchSub,tblMastItem,tblMastAccount,tblMastAccount tblMastAccount_Pty " _
        & " where tblPurch.PurAcDrCode = tblMastAccount.AcCode and tblPurchSub.PurSubItCode = tblMastItem.ItCode and tblPurch.Vno = tblPurchSub.Vno and " _
        & " tblPurch.Vtype = tblPurchSub.Vtype and tblPurch.Vyear = tblPurchSub.Vyear and tblPurch.Vfirm = tblPurchSub.Vfirm and " & x1 & " " _
        & " and tblPurch.PurAcCrCode = tblMastAccount_Pty.AcCode and PurSubItSrNo = 1 " _
        & " group by tblMastAccount.acname,tblMastItem.ItHsn2,tblPurch.PurTaxRate+tblPurch.PurTaxRate2+tblPurch.PurTaxRate3"
Else
    X = "select 'K','Purchases',tblMastAccount.acname,tblMastItem.ItHsn,tblPurch.PurTaxRate+tblPurch.PurTaxRate2+tblPurch.PurTaxRate3," _
        & " Sum (PurTaxableAmt), Sum(PurTaxAmt2), Sum(PurTaxAmt), Sum(PurTaxAmt3), Sum(PurBillAmt),'" & gUserName & "'" _
        & " from tblPurch,tblPurchSub,tblMastItem,tblMastAccount,tblMastAccount tblMastAccount_Pty " _
        & " where tblPurch.PurAcDrCode = tblMastAccount.AcCode and tblPurchSub.PurSubItCode = tblMastItem.ItCode and tblPurch.Vno = tblPurchSub.Vno and " _
        & " tblPurch.Vtype = tblPurchSub.Vtype and tblPurch.Vyear = tblPurchSub.Vyear and tblPurch.Vfirm = tblPurchSub.Vfirm and " & x1 & " " _
        & " and tblPurch.PurAcCrCode = tblMastAccount_Pty.AcCode and PurSubItSrNo = 1 " _
        & " group by tblMastAccount.acname,tblMastItem.ItHsn,tblPurch.PurTaxRate+tblPurch.PurTaxRate2+tblPurch.PurTaxRate3"
End If
gCn.Execute " Insert Into TmpGentbl (VType,Nar1,Nar2,Nar3,Amt1,Amt2,Amt3,Amt4,Amt5,Amt6,UserName)" & X
gCn.CommitTrans

RecYN = GProcRstOpen(tmp1, "select * from TmpGentbl where vtype = 'K' and TmpGentbl.UserName ='" & gUserName & "'", "R", gCn)
If RecYN > 0 Then
   Set sheet = workbook.Sheets.Item("Yearly")
   If (sheet Is Nothing) Then
       MsgBox "Error Yearly Sheet"
   End If
   With tmp1
        .MoveFirst
        Do While Not .EOF
           sheet.cells(a, 2) = "Purchases"
           sheet.cells(a, 3) = !Nar2
           If Left(!Nar3, 2) = "99" Then
              sheet.cells(a, 4) = "Services"
           Else
              sheet.cells(a, 4) = "Goods"
           End If
           sheet.cells(a, 5) = "Registered"
           sheet.cells(a, 6) = "Debit Note"
           sheet.cells(a, 7) = !Nar3
           sheet.cells(a, 8) = !Amt1
           gCn.BeginTrans
           For i = 1 To 12
               Set sheet = workbook.Sheets.Item("" & Left(MonthName(i), 3) & "")
               If gCYear >= 2021 Then
                  X = "select 'K','4',tblMastAccount.acname,tblMastItem.ItHsn2,tblPurch.PurTaxRate+tblPurch.PurTaxRate2+tblPurch.PurTaxRate3," _
                      & " Sum (PurTaxableAmt), Sum(PurTaxAmt2), Sum(PurTaxAmt), Sum(PurTaxAmt3), Sum(PurBillAmt)," & i & ",'" & gUserName & "'" _
                      & " from tblPurch,tblPurchSub,tblMastItem,tblMastAccount,tblMastAccount tblMastAccount_Pty " _
                      & " where tblPurch.PurAcDrCode = tblMastAccount.AcCode and tblPurchSub.PurSubItCode = tblMastItem.ItCode and tblPurch.Vno = tblPurchSub.Vno and " _
                      & " tblPurch.Vtype = tblPurchSub.Vtype and tblPurch.Vyear = tblPurchSub.Vyear and tblPurch.Vfirm = tblPurchSub.Vfirm and " & x1 & " and month(tblPurch.vdt) = " & i & " and tblMastItem.ItHsn2 = '" & tmp1.Fields("Nar3") & "' and tblPurch.PurTaxRate+tblPurch.PurTaxRate2+tblPurch.PurTaxRate3 = " & tmp1.Fields("Amt1") & "" _
                      & " and tblPurch.PurAcCrCode = tblMastAccount_Pty.AcCode and PurSubItSrNo = 1 " _
                      & " and tblMastAccount.AcName = '" & tmp1.Fields("Nar2") & "' " _
                      & " group by tblMastAccount.acname,tblMastItem.ItHsn2,tblPurch.PurTaxRate+tblPurch.PurTaxRate2+tblPurch.PurTaxRate3"
               Else
                  X = "select 'K','4',tblMastAccount.acname,tblMastItem.ItHsn,tblPurch.PurTaxRate+tblPurch.PurTaxRate2+tblPurch.PurTaxRate3," _
                      & " Sum (PurTaxableAmt), Sum(PurTaxAmt2), Sum(PurTaxAmt), Sum(PurTaxAmt3), Sum(PurBillAmt)," & i & ",'" & gUserName & "'" _
                      & " from tblPurch,tblPurchSub,tblMastItem,tblMastAccount,tblMastAccount tblMastAccount_Pty " _
                      & " where tblPurch.PurAcDrCode = tblMastAccount.AcCode and tblPurchSub.PurSubItCode = tblMastItem.ItCode and tblPurch.Vno = tblPurchSub.Vno and " _
                      & " tblPurch.Vtype = tblPurchSub.Vtype and tblPurch.Vyear = tblPurchSub.Vyear and tblPurch.Vfirm = tblPurchSub.Vfirm and " & x1 & " and month(tblPurch.vdt) = " & i & " and tblMastItem.ItHsn = '" & tmp1.Fields("Nar3") & "' and tblPurch.PurTaxRate+tblPurch.PurTaxRate2+tblPurch.PurTaxRate3 = " & tmp1.Fields("Amt1") & "" _
                      & " and tblPurch.PurAcCrCode = tblMastAccount_Pty.AcCode and PurSubItSrNo = 1 " _
                      & " and tblMastAccount.AcName = '" & tmp1.Fields("Nar2") & "' " _
                      & " group by tblMastAccount.acname,tblMastItem.ItHsn,tblPurch.PurTaxRate+tblPurch.PurTaxRate2+tblPurch.PurTaxRate3"
               End If
               gCn.Execute " Insert Into TmpGentbl2 (VType,Nar1,Nar2,Nar3,Amt1,Amt2,Amt3,Amt4,Amt5,Amt6,Vno,UserName)" & X
               j = GProcRstOpen(tmp2, X, "R", gCn)
               If j > 0 Then
                  sheet.cells(a, 10) = IIf(IsNull(tmp2.Fields(5)), 0, tmp2.Fields(5))
                  sheet.cells(a, 11) = IIf(IsNull(tmp2.Fields(6)), 0, tmp2.Fields(6))
                  sheet.cells(a, 12) = IIf(IsNull(tmp2.Fields(7)), 0, tmp2.Fields(7))
                  sheet.cells(a, 13) = IIf(IsNull(tmp2.Fields(8)), 0, tmp2.Fields(8))
                  sheet.cells(a, 14) = 0
               Else
                  sheet.cells(a, 10) = 0
                  sheet.cells(a, 11) = 0
                  sheet.cells(a, 12) = 0
                  sheet.cells(a, 13) = 0
                  sheet.cells(a, 14) = 0
               End If

           Next i
           gCn.CommitTrans
           Set sheet = workbook.Sheets.Item("Yearly")
           a = a + 1
           If a = 35 Then a = 37
           If a = 39 Then a = 45
           If a = 60 Then a = 67
           If a = 127 Then a = 133
           If a = 171 Then a = 173
           tmp1.MoveNext
        Loop
   End With
End If











'--For Total RCM Purchase
gCn.BeginTrans
   
x1 = " tblPurch.VType = 'PO' and tblPurch.PurBillDt>= '" & Format(w_Fr, "yyyy/mm/dd") & "'  and tblPurch.PurBillDt<= '" & Format(w_To, "yyyy/mm/dd") & "' and tblPurch.PurRCMIgstAmt+tblPurch.PurRCMSgstAmt+tblPurch.PurRCMCgstAmt > 0  and tblPurch.Vfirm = '" & gCCode & "' "
If gCYear >= 2021 Then
    X = "select 'J','Purchases',tblMastAccount.acname,tblMastItem.ItHsn2,tblPurch.PurRCMCgstRt+tblPurch.PurRCMSgstRt+tblPurch.PurRCMIgstRt," _
        & " Sum (tblPurch.PurSubAmt), Sum(PurRCMCgstAmt), Sum(PurRCMSgstAmt), Sum(PurRCMIgstAmt), Sum(PurBillAmt),'" & gUserName & "'" _
        & " from tblPurch,tblPurchSub,tblMastItem,tblMastAccount,tblMastAccount tblMastAccount_Pty " _
        & " where tblPurch.PurAcDrCode = tblMastAccount.AcCode and tblPurchSub.PurSubItCode = tblMastItem.ItCode and tblPurch.Vno = tblPurchSub.Vno and " _
        & " tblPurch.Vtype = tblPurchSub.Vtype and tblPurch.Vyear = tblPurchSub.Vyear and tblPurch.Vfirm = tblPurchSub.Vfirm and " & x1 & " " _
        & " and tblPurch.PurAcCrCode = tblMastAccount_Pty.AcCode and PurSubItSrNo = 1 " _
        & " group by tblMastAccount.acname,tblMastItem.ItHsn2,tblPurch.PurRCMCgstRt+tblPurch.PurRCMSgstRt+tblPurch.PurRCMIgstRt"
Else
    X = "select 'J','Purchases',tblMastAccount.acname,tblMastItem.ItHsn,tblPurch.PurRCMCgstRt+tblPurch.PurRCMSgstRt+tblPurch.PurRCMIgstRt," _
        & " Sum (tblPurch.PurSubAmt), Sum(PurRCMCgstAmt), Sum(PurRCMSgstAmt), Sum(PurRCMIgstAmt), Sum(PurBillAmt),'" & gUserName & "'" _
        & " from tblPurch,tblPurchSub,tblMastItem,tblMastAccount,tblMastAccount tblMastAccount_Pty " _
        & " where tblPurch.PurAcDrCode = tblMastAccount.AcCode and tblPurchSub.PurSubItCode = tblMastItem.ItCode and tblPurch.Vno = tblPurchSub.Vno and " _
        & " tblPurch.Vtype = tblPurchSub.Vtype and tblPurch.Vyear = tblPurchSub.Vyear and tblPurch.Vfirm = tblPurchSub.Vfirm and " & x1 & " " _
        & " and tblPurch.PurAcCrCode = tblMastAccount_Pty.AcCode and PurSubItSrNo = 1 " _
        & " group by tblMastAccount.acname,tblMastItem.ItHsn,tblPurch.PurRCMCgstRt+tblPurch.PurRCMSgstRt+tblPurch.PurRCMIgstRt"
End If
gCn.Execute " Insert Into TmpGentbl (VType,Nar1,Nar2,Nar3,Amt1,Amt2,Amt3,Amt4,Amt5,Amt6,UserName)" & X
gCn.CommitTrans

RecYN = GProcRstOpen(tmp1, "select * from TmpGentbl where vtype = 'J' and TmpGentbl.UserName ='" & gUserName & "'", "R", gCn)
If RecYN > 0 Then
   Set sheet = workbook.Sheets.Item("Yearly")
   If (sheet Is Nothing) Then
       MsgBox "Error Yearly Sheet"
   End If
   With tmp1
        .MoveFirst
        Do While Not .EOF
           sheet.cells(a, 2) = "Others"
           sheet.cells(a, 3) = !Nar2
           If Left(!Nar3, 2) = "99" Then
              sheet.cells(a, 4) = "Services"
           Else
              sheet.cells(a, 4) = "Others"
           End If
           sheet.cells(a, 5) = "Unregistered"
           sheet.cells(a, 6) = "RCM Invoice"
           sheet.cells(a, 7) = !Nar3
           sheet.cells(a, 8) = !Amt1
           gCn.BeginTrans
           For i = 1 To 12
               Set sheet = workbook.Sheets.Item("" & Left(MonthName(i), 3) & "")
               If gCYear >= 2021 Then
                  X = "select 'J','4',tblMastAccount.acname,tblMastItem.ItHsn2,tblPurch.PurRCMCgstRt+tblPurch.PurRCMSgstRt+tblPurch.PurRCMIgstRt," _
                      & " Sum (tblPurch.PurSubAmt), Sum(PurRCMCgstAmt), Sum(PurRCMSgstAmt), Sum(PurRCMIgstAmt), Sum(PurBillAmt)," & i & ",'" & gUserName & "'" _
                      & " from tblPurch,tblPurchSub,tblMastItem,tblMastAccount,tblMastAccount tblMastAccount_Pty " _
                      & " where tblPurch.PurAcDrCode = tblMastAccount.AcCode and tblPurchSub.PurSubItCode = tblMastItem.ItCode and tblPurch.Vno = tblPurchSub.Vno and " _
                      & " tblPurch.Vtype = tblPurchSub.Vtype and tblPurch.Vyear = tblPurchSub.Vyear and tblPurch.Vfirm = tblPurchSub.Vfirm and " & x1 & " and month(tblPurch.PurBillDt) = " & i & " and tblMastItem.ItHsn2 = '" & tmp1.Fields("Nar3") & "' and tblPurch.PurTaxRate+tblPurch.PurTaxRate2+tblPurch.PurTaxRate3 = " & tmp1.Fields("Amt1") & "" _
                      & " and tblPurch.PurAcCrCode = tblMastAccount_Pty.AcCode and PurSubItSrNo = 1 " _
                      & " and tblMastAccount.AcName = '" & tmp1.Fields("Nar2") & "' " _
                      & " group by tblMastAccount.acname,tblMastItem.ItHsn2,tblPurch.PurRCMCgstRt+tblPurch.PurRCMSgstRt+tblPurch.PurRCMIgstRt"
               Else
                  X = "select 'J','4',tblMastAccount.acname,tblMastItem.ItHsn,tblPurch.PurRCMCgstRt+tblPurch.PurRCMSgstRt+tblPurch.PurRCMIgstRt," _
                      & " Sum (tblPurch.PurSubAmt), Sum(PurRCMCgstAmt), Sum(PurRCMSgstAmt), Sum(PurRCMIgstAmt), Sum(PurBillAmt)," & i & ",'" & gUserName & "'" _
                      & " from tblPurch,tblPurchSub,tblMastItem,tblMastAccount,tblMastAccount tblMastAccount_Pty " _
                      & " where tblPurch.PurAcDrCode = tblMastAccount.AcCode and tblPurchSub.PurSubItCode = tblMastItem.ItCode and tblPurch.Vno = tblPurchSub.Vno and " _
                      & " tblPurch.Vtype = tblPurchSub.Vtype and tblPurch.Vyear = tblPurchSub.Vyear and tblPurch.Vfirm = tblPurchSub.Vfirm and " & x1 & " and month(tblPurch.PurBillDt) = " & i & " and tblMastItem.ItHsn = '" & tmp1.Fields("Nar3") & "' and tblPurch.PurTaxRate+tblPurch.PurTaxRate2+tblPurch.PurTaxRate3 = " & tmp1.Fields("Amt1") & "" _
                      & " and tblPurch.PurAcCrCode = tblMastAccount_Pty.AcCode and PurSubItSrNo = 1 " _
                      & " and tblMastAccount.AcName = '" & tmp1.Fields("Nar2") & "' " _
                      & " group by tblMastAccount.acname,tblMastItem.ItHsn,tblPurch.PurRCMCgstRt+tblPurch.PurRCMSgstRt+tblPurch.PurRCMIgstRt"
               End If
               gCn.Execute " Insert Into TmpGentbl2 (VType,Nar1,Nar2,Nar3,Amt1,Amt2,Amt3,Amt4,Amt5,Amt6,Vno,UserName)" & X
               j = GProcRstOpen(tmp2, X, "R", gCn)
               If j > 0 Then
                  sheet.cells(a, 10) = IIf(IsNull(tmp2.Fields(5)), 0, tmp2.Fields(5))
                  sheet.cells(a, 11) = IIf(IsNull(tmp2.Fields(6)), 0, tmp2.Fields(6))
                  sheet.cells(a, 12) = IIf(IsNull(tmp2.Fields(7)), 0, tmp2.Fields(7))
                  sheet.cells(a, 13) = IIf(IsNull(tmp2.Fields(8)), 0, tmp2.Fields(8))
                  sheet.cells(a, 14) = 0
               Else
                  sheet.cells(a, 10) = 0
                  sheet.cells(a, 11) = 0
                  sheet.cells(a, 12) = 0
                  sheet.cells(a, 13) = 0
                  sheet.cells(a, 14) = 0
               End If

           Next i
           gCn.CommitTrans
           Set sheet = workbook.Sheets.Item("Yearly")
           a = a + 1
           If a = 35 Then a = 37
           If a = 39 Then a = 45
           If a = 60 Then a = 67
           If a = 127 Then a = 133
           If a = 171 Then a = 173
           tmp1.MoveNext
        Loop
   End With
End If

workbook.Close SaveChanges:=True
excel_app.Quit
MsgBox "Compileted"
Set excel_app = Nothing
End Function
Public Sub GProcWhatsAppError()
Dim X As String
Dim RsTmp As Recordset
Dim RetVal
'gReportPath = App.Path & "\HIReports\"
'gReportPath = "D:\"
'Open "" & gReportPath & "rpt1.txt" For Output As #1
'Open "d:\rpt1.txt" For Output As #1
X = "Select * from tmpWhatsAppErr where username = '" & gUserName & "' and vfirm = '" & gCCode & "'"
If GProcRstOpen(RsTmp, X, "R", gCn) > 0 Then
   Open "d:\rpt1.txt" For Output As #1
   With RsTmp
       .MoveFirst
       Print #1, ""
       Print #1, ""
       Print #1, "Mobile No 1       Mobile No 2     Party Name                      Error"
       Print #1, "-----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------"
       Print #1, ""
       Do While Not .EOF
          X = Left(.Fields(1) + Space(15), 15) + "  " + Left(IIf(IsNull(.Fields(2)), "", .Fields(2)) + Space(15), 15) + "  " + Left(.Fields(3) + Space(30), 30) + "  " + Left(.Fields(4) + Space(250), 250)
          Print #1, X
          .MoveNext
       Loop
       Print #1, ""
       Print #1, "-----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------"
   End With
   Close #1
'   RetVal = Shell(gReportPath + "WORDPAD " & "" & gReportPath & "rpt1.txt", vbMaximizedFocus)
   RetVal = Shell("D:\WORDPAD " & "d:\rpt1.txt", vbMaximizedFocus)
End If
End Sub
Public Sub GProcTDSReportInExcle(w_Fr As Date, w_To As Date, RepoName As String)
Dim x1 As String
Dim x2 As String
Dim flnm As String
Dim RecYN As Long
Dim RsTmp As Recordset
Dim RsTmp2 As Recordset
Dim excel_app As Excel.Application
Dim workbook As Excel.workbook
Dim sheet_name As String
Dim sheet As Excel.Worksheet
Dim header_range As Excel.Range
Dim a As Long
Dim xlApp As Object
Dim xlWB As Object
Dim i As Long
Dim AssVal As Double
flnm = "D:\TDS_Report_" + Format(w_To, "mmm") + "-" + gCCode + "_" & gUserName & "" + ".xlsx"
FileCopy App.Path & "\HiReports\TDS_Report_Blank.xlsx", flnm
   
Set excel_app = New Excel.Application
excel_app.Visible = False
Set workbook = excel_app.Workbooks.Open(filename:=flnm)
x1 = "Select VFirmName,PartyName,sum(Amt10),sum(iif(Amt19 >0,Amt20,0)),sum(Amt19) from TmpGentbl where username = '" & gUserName & "' and TmpGentbl.Nar6 = 'A' and  TmpGentbl.Amt9 = 1 " _
     & " group by VFirmName,PartyName"
If GProcRstOpen(RsTmp, x1, "R", gCn) > 0 Then
   Set sheet = workbook.Sheets.Item("194Q")
      If (sheet Is Nothing) Then
         MsgBox "Error 194Q Sheet"
      End If
      a = 1
      sheet.cells(1, 5) = RsTmp.Fields(0)
      sheet.cells(2, 5) = RepoName
      sheet.cells(3, 5) = "From : " & w_Fr & " To : " & w_To & ""
      a = 6
      i = 1
      With RsTmp
           RsTmp.MoveFirst
           Do While Not .EOF
              x2 = "Select AcName,AcAdd1,AcAdd2,AcCity,AcZip,AcState,AcPAN,AcGSTIN from tblMastAccount where AcName = '" & RsTmp.Fields(1) & "'"
              If GProcRstOpen(RsTmp2, x2, "R", gCn) > 0 Then
                 sheet.cells(a, 2) = i
                 i = i + 1
                 sheet.cells(a, 4) = RsTmp2!AcPAN
                 sheet.cells(a, 5) = UCase(RsTmp!PartyName)
                 sheet.cells(a, 8) = w_To
                 AssVal = Format(GProcMakeRounding(RsTmp.Fields(3)), FStr)
                 sheet.cells(a, 9) = AssVal
                 sheet.cells(a, 10) = RsTmp.Fields(4)
                 sheet.cells(a, 12) = RsTmp2!AcAdd1
                 sheet.cells(a, 13) = RsTmp2!AcAdd2
                 sheet.cells(a, 16) = RsTmp2!AcCity
                 sheet.cells(a, 17) = RsTmp2!AcZip
                 Select Case Left(RsTmp2!AcGSTIN, 2)
                 Case "35":
                     sheet.cells(a, 18) = "01. Andaman & Nicobar Islands"
                 Case "37":
                     sheet.cells(a, 18) = "02. Andhra Pradesh"
                 Case "12":
                     sheet.cells(a, 18) = "03. Arunachal Pradesh"
                 Case "18":
                     sheet.cells(a, 18) = "04. Assam"
                 Case "10":
                     sheet.cells(a, 18) = "05. Bihar"
                 Case "04":
                     sheet.cells(a, 18) = "06. Chandigarh"
                 Case "26":
                     sheet.cells(a, 18) = "07. Dadra & Nagar Haveli"
                 Case "25":
                     sheet.cells(a, 18) = "08. Daman & Diu"
                 Case "07":
                     sheet.cells(a, 18) = "09. Delhi"
                 Case "30":
                     sheet.cells(a, 18) = "10. Goa"
                 Case "24":
                     sheet.cells(a, 18) = "11. Gujarat"
                 Case "06":
                     sheet.cells(a, 18) = "12. Haryana"
                 Case "02":
                     sheet.cells(a, 18) = "13. Himachal Pradesh"
                 Case "01":
                     sheet.cells(a, 18) = "14. Jammu & Kashmir"
                 Case "29":
                     sheet.cells(a, 18) = "15. Karnataka"
                 Case "32":
                     sheet.cells(a, 18) = "16. Kerala"
                 Case "31":
                     sheet.cells(a, 18) = "17. Lakhswadeep"
                 Case "23":
                     sheet.cells(a, 18) = "18. Madhya Pradesh"
                 Case "27":
                     sheet.cells(a, 18) = "19. Maharashtra"
                 Case "14":
                     sheet.cells(a, 18) = "20. Manipur"
                 Case "17":
                     sheet.cells(a, 18) = "21. Meghalaya"
                 Case "15":
                     sheet.cells(a, 18) = "22. Mizoram"
                 Case "13":
                     sheet.cells(a, 18) = "23. Nagaland"
                 Case "21":
                     sheet.cells(a, 18) = "24. Orissa"
                 Case "34":
                     sheet.cells(a, 18) = "25. Pondicherry"
                 Case "03":
                     sheet.cells(a, 18) = "26. Punjab"
                 Case "08":
                     sheet.cells(a, 18) = "27. Rajasthan"
                 Case "11":
                     sheet.cells(a, 18) = "28. Sikkim"
                 Case "33":
                     sheet.cells(a, 18) = "29. Tamilnadu"
                 Case "16":
                     sheet.cells(a, 18) = "30. Tripura"
                 Case "09":
                     sheet.cells(a, 18) = "31. Uttar Pradesh"
                 Case "19":
                     sheet.cells(a, 18) = "32. West Bengal"
                 Case "22":
                     sheet.cells(a, 18) = "33. Chhatishgarh"
                 Case "05":
                     sheet.cells(a, 18) = "34. Uttaranchal"
                 Case "20":
                     sheet.cells(a, 18) = "35. Jharkhand"
                 End Select
                 a = a + 1
              End If
              RsTmp.MoveNext
           Loop
      End With
   End If
   workbook.Close SaveChanges:=True
    
   Set xlApp = CreateObject("Excel.Application")
   xlApp.Visible = True
    
   Set xlWB = xlApp.Workbooks.Open("" & flnm & "")
   excel_app.Quit
   Exit Sub

End Sub
Public Sub GProcTDSReportInExcleMillDepot(w_Fr As Date, w_To As Date, RepoName As String)
Dim x1 As String
Dim x2 As String
Dim flnm As String
Dim RecYN As Long
Dim RsTmp As Recordset
Dim RsTmp2 As Recordset
Dim excel_app As Excel.Application
Dim workbook As Excel.workbook
Dim sheet_name As String
Dim sheet As Excel.Worksheet
Dim header_range As Excel.Range
Dim a As Long
Dim xlApp As Object
Dim xlWB As Object
Dim i As Long
Dim AssVal As Double
flnm = "D:\TDS_Report_" + Format(w_To, "mmm") + "-" + gCCode + "_" & gUserName & "" + ".xlsx"
FileCopy App.Path & "\HiReports\TDS_Report_Blank.xlsx", flnm
Set excel_app = New Excel.Application
excel_app.Visible = False
Set workbook = excel_app.Workbooks.Open(filename:=flnm)
x1 = "Select VFirmName,MillName,PartyName,sum(Amt10),sum(Amt1),sum(Amt19) from TmpGentbl where username = '" & gUserName & "' and TmpGentbl.Nar6 = 'A' and  TmpGentbl.Amt14 = 1 " _
     & " group by VFirmName,MillName,PartyName"
If GProcRstOpen(RsTmp, x1, "R", gCn) > 0 Then
   Set sheet = workbook.Sheets.Item("194Q")
      If (sheet Is Nothing) Then
         MsgBox "Error 194Q Sheet"
      End If
      a = 1
      sheet.cells(1, 5) = RsTmp.Fields(0)
      sheet.cells(2, 5) = RepoName
      sheet.cells(3, 5) = "From : " & w_Fr & " To : " & w_To & ""
      a = 6
      i = 1
      With RsTmp
           RsTmp.MoveFirst
           Do While Not .EOF
              x2 = "Select AcName,AcAdd1,AcAdd2,AcCity,AcZip,AcState,AcPAN,AcGSTIN from tblMastAccount where AcName = '" & RsTmp.Fields(1) & "'"
              If GProcRstOpen(RsTmp2, x2, "R", gCn) > 0 Then
                 sheet.cells(a, 2) = i
                 i = i + 1
                 sheet.cells(a, 4) = RsTmp2!AcPAN
                 sheet.cells(a, 5) = RsTmp!PartyName
                 sheet.cells(a, 8) = w_To
                 AssVal = Format(GProcMakeRounding(RsTmp.Fields(4)), FStr)
                 
                 sheet.cells(a, 9) = AssVal
                 sheet.cells(a, 10) = RsTmp.Fields(5)
                 sheet.cells(a, 12) = RsTmp2!AcAdd1
                 sheet.cells(a, 13) = RsTmp2!AcAdd2
                 sheet.cells(a, 16) = RsTmp2!AcCity
                 sheet.cells(a, 17) = RsTmp2!AcZip
                 Select Case Left(RsTmp2!AcGSTIN, 2)
                 Case "35":
                     sheet.cells(a, 18) = "01. Andaman & Nicobar Islands"
                 Case "37":
                     sheet.cells(a, 18) = "02. Andhra Pradesh"
                 Case "12":
                     sheet.cells(a, 18) = "03. Arunachal Pradesh"
                 Case "18":
                     sheet.cells(a, 18) = "04. Assam"
                 Case "10":
                     sheet.cells(a, 18) = "05. Bihar"
                 Case "04":
                     sheet.cells(a, 18) = "06. Chandigarh"
                 Case "26":
                     sheet.cells(a, 18) = "07. Dadra & Nagar Haveli"
                 Case "25":
                     sheet.cells(a, 18) = "08. Daman & Diu"
                 Case "07":
                     sheet.cells(a, 18) = "09. Delhi"
                 Case "30":
                     sheet.cells(a, 18) = "10. Goa"
                 Case "24":
                     sheet.cells(a, 18) = "11. Gujarat"
                 Case "06":
                     sheet.cells(a, 18) = "12. Haryana"
                 Case "02":
                     sheet.cells(a, 18) = "13. Himachal Pradesh"
                 Case "01":
                     sheet.cells(a, 18) = "14. Jammu & Kashmir"
                 Case "29":
                     sheet.cells(a, 18) = "15. Karnataka"
                 Case "32":
                     sheet.cells(a, 18) = "16. Kerala"
                 Case "31":
                     sheet.cells(a, 18) = "17. Lakhswadeep"
                 Case "23":
                     sheet.cells(a, 18) = "18. Madhya Pradesh"
                 Case "27":
                     sheet.cells(a, 18) = "19. Maharashtra"
                 Case "14":
                     sheet.cells(a, 18) = "20. Manipur"
                 Case "17":
                     sheet.cells(a, 18) = "21. Meghalaya"
                 Case "15":
                     sheet.cells(a, 18) = "22. Mizoram"
                 Case "13":
                     sheet.cells(a, 18) = "23. Nagaland"
                 Case "21":
                     sheet.cells(a, 18) = "24. Orissa"
                 Case "34":
                     sheet.cells(a, 18) = "25. Pondicherry"
                 Case "03":
                     sheet.cells(a, 18) = "26. Punjab"
                 Case "08":
                     sheet.cells(a, 18) = "27. Rajasthan"
                 Case "11":
                     sheet.cells(a, 18) = "28. Sikkim"
                 Case "33":
                     sheet.cells(a, 18) = "29. Tamilnadu"
                 Case "16":
                     sheet.cells(a, 18) = "30. Tripura"
                 Case "09":
                     sheet.cells(a, 18) = "31. Uttar Pradesh"
                 Case "19":
                     sheet.cells(a, 18) = "32. West Bengal"
                 Case "22":
                     sheet.cells(a, 18) = "33. Chhatishgarh"
                 Case "05":
                     sheet.cells(a, 18) = "34. Uttaranchal"
                 Case "20":
                     sheet.cells(a, 18) = "35. Jharkhand"
                 End Select
                 sheet.cells(a, 21) = RsTmp.Fields(1)
                 a = a + 1
              End If
              RsTmp.MoveNext
           Loop
      End With
   End If
   workbook.Close SaveChanges:=True
    
   Set xlApp = CreateObject("Excel.Application")
   xlApp.Visible = True
    
   Set xlWB = xlApp.Workbooks.Open("" & flnm & "")
   excel_app.Quit
   Exit Sub

End Sub
Public Function GProcDailyReportMail(w_Fr As Date, w_To As Date, wMillCd As Long, wPtyGrpCd As Long, w_PurBkFr As Date)
gCn.BeginTrans
gCmd.CommandText = "PrcPrepareDailyReportMail"
gCmd.Parameters.Refresh
gCmd.Parameters("@YearStart_Dt") = Format(gCYSDate, "yyyy/mm/dd")
gCmd.Parameters("@Fr_dt") = Format(w_Fr, "yyyy/mm/dd")
gCmd.Parameters("@To_dt") = Format(w_To, "yyyy/mm/dd")
gCmd.Parameters("@BkFr_dt") = Format(w_PurBkFr, "yyyy/mm/dd")
If wMillCd <> 0 Then gCmd.Parameters("@MillCode") = wMillCd
If wPtyGrpCd <> 0 Then gCmd.Parameters("@PtyGrpCode") = wPtyGrpCd
gCmd.Parameters("@UCode") = gUserName
gCmd.Execute
gCn.CommitTrans
End Function

