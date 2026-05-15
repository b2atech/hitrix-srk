Attribute VB_Name = "Account"
Option Explicit
Dim CommAmt As Double
Dim CommOn As Double
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
Dim X1 As String
Dim ConcatOperator As String
Dim NotNullStr As String
Dim NullStr As String
If gBackEndDB = gBackEndAccess Then
    ConcatOperator = "&"
    NotNullStr = " <>'' "
    NullStr = " ='' "
Else
    ConcatOperator = "||"
    NotNullStr = " is not null "
    NullStr = " is Null "
End If
'--- Insert Into tmpRptTbl
'--- With out & with Schedule ( For 1 To 5 Group Head) ,closing Stock,Profit &Loss of Yr
gCn.Execute "delete from tmpRptTbl where UserName='" & gUserName & "'"
X1 = "insert into TmpRptTbl (Vtype,Nar1,Amt1,UserName,Nar2,Nar3,Nar4,Nar5,Nar6,Nar7,Nar8,Nar9,Nar10,Nar11,Nar12,Amt2) "

'--- Group Head & A/c H2 WithOut Schedule
gCn.Execute X1 & " select distinct 'B' as Vtype,H2 as Nar," & IIf(gBackEndDB = gBackEndAccess, "iif(H2Type='A',ClCr-ClDr,0)", "decode(H2Type,'A',ClCr-ClDr,0)") & " as Bal,username,H1,H2,'' as H3,'' as H4,'' as H5,'' as H6,H1TYPE, H2TYPE,'-' as H3Type,'-' as H4type,'-' as H5Type,0 as IsSchedule from tmpGrpTrbalDetail where IsOpStkAc=0 and H2 " & NotNullStr & " and (schedule is null or (H2Type='G' and ScheduleHno=2 ) ) and TopGroup=" & TopGrpcode & " and username='" & gUserName & "'"
'--- Group Head & A/c H3 WithOut Schedule
gCn.Execute X1 & "  select distinct 'B','  ' " & ConcatOperator & "  H3," & IIf(gBackEndDB = gBackEndAccess, "iif(H3Type='A',ClCr-ClDr,0)", "decode(H3Type,'A',ClCr-ClDr,0)") & ",username,H1,H2,H3,'','','' as H6,H1TYPE, H2TYPE,H3TYPE,'-','-' as H5Type,0 from tmpGrpTrbalDetail where IsOpStkAc=0 and  H3 " & NotNullStr & " and ( schedule is null or (H3Type='G' and ScheduleHno=3 ) ) and TopGroup=" & TopGrpcode & " and username='" & gUserName & "'"
'--- Group Head & A/c H4 WithOut Schedule
gCn.Execute X1 & " select distinct 'B','    ' " & ConcatOperator & "  H4," & IIf(gBackEndDB = gBackEndAccess, "iif(H4Type='A',ClCr-ClDr,0)", "decode(H4Type,'A',ClCr-ClDr,0)") & ",username,H1,H2,H3,H4,'','' as H6,H1TYPE, H2TYPE,H3TYPE,H4Type,'-' as H5Type,0  from tmpGrpTrbalDetail where IsOpStkAc=0 and  H4 " & NotNullStr & " and (schedule is null or (H4Type='G' and ScheduleHno=4 ) ) and TopGroup=" & TopGrpcode & "  and username='" & gUserName & "'"
'--- Group Head & A/c H5 WithOut Schedule
gCn.Execute X1 & " select distinct 'B','        ' " & ConcatOperator & "  H5," & IIf(gBackEndDB = gBackEndAccess, "iif(H5Type='A',ClCr-ClDr,0)", "decode(H5Type,'A',ClCr-ClDr,0)") & ",username,H1,H2,H3,H4,H5,'' as H6,H1TYPE, H2TYPE,H3TYPE,H4Type,tmpGrpTrbalDetail.H5Type as H5Type,0 from tmpGrpTrbalDetail where IsOpStkAc=0 and H5 " & NotNullStr & " and (schedule is null or (H5Type='G' and ScheduleHno=5 ) ) and TopGroup=" & TopGrpcode & "  and username='" & gUserName & "'"
'--- Group Head & A/c H6 WithOut Schedule
gCn.Execute X1 & " select distinct 'B','            ' " & ConcatOperator & "  H6 as Nar," & IIf(gBackEndDB = gBackEndAccess, "iif(H6Type='A',ClCr-ClDr,0)", "decode(H6Type,'A',ClCr-ClDr,0)") & " ,username,H1,H2,H3,H4,H5,H6,H1TYPE, H2TYPE,H3TYPE,H4Type,tmpGrpTrbalDetail.H5Type,0 from tmpGrpTrbalDetail where IsOpStkAc=0 and H6 " & NotNullStr & " and schedule is null and TopGroup=" & TopGrpcode & " and username='" & gUserName & "'"

'--- Group Head & A/c H2 With Schedule
gCn.Execute X1 & " select distinct 'B' ,'As Per Schedule ' " & ConcatOperator & " Schedule ,sum(ClCr-ClDr) ,username,H1,schedule ,'','','','' as H6,H1TYPE, '0','-','-','-' as H5Type,1 from tmpGrpTrbalDetail where ScheduleHNo=1 and schedule " & NotNullStr & " and " _
& " TopGroup=" & TopGrpcode & " and username='" & gUserName & "' group by Schedule,H1,H1Type,username "
'--- Group Head & A/c H3 With Schedule
gCn.Execute X1 & " select distinct 'B','  ' " & ConcatOperator & "  'As Per Schedule ' " & ConcatOperator & " Schedule ,sum(ClCr-ClDr) ,username,H1,H2,schedule ,'','','' as H6,H1TYPE, H2TYPE,'0','-','-' as H5Type,1 from tmpGrpTrbalDetail where IsOpStkAc=0 and ScheduleHNo=2 and schedule " & NotNullStr & " and TopGroup=" & TopGrpcode & " " _
& " and username='" & gUserName & "' group by Schedule,H1,H2,H1Type,H2Type,username "
'--- Group Head & A/c H4 With Schedule
gCn.Execute X1 & " select distinct 'B','    ' " & ConcatOperator & "  'As Per Schedule ' " & ConcatOperator & " Schedule,sum(ClCr-ClDr),username,H1,H2,H3,schedule ,'','' as H6,H1TYPE, H2TYPE,H3TYPE,'0' ,'-' as H5Type,1 from tmpGrpTrbalDetail where  IsOpStkAc=0 and ScheduleHNo=3 and schedule " & NotNullStr & " and TopGroup=" & TopGrpcode & "  and username='" & gUserName & "'" _
& " Group by Schedule,H1,H2,H3,H1Type,H2Type,H3Type,username "
'--- Group Head & A/c H5 With Schedule
gCn.Execute X1 & " select distinct 'B','        ' " & ConcatOperator & "  'As Per Schedule ' " & ConcatOperator & " Schedule ,sum(ClCr-ClDr),username,H1,H2,H3,H4,schedule as  H5,'' as H6,H1TYPE, H2TYPE,H3TYPE,H4Type,'0' as H5Type,1 from tmpGrpTrbalDetail where  IsOpStkAc=0 and ScheduleHNo=4 and schedule " & NotNullStr & " and TopGroup=" & TopGrpcode & "  and username='" & gUserName & "'" _
& " Group by Schedule,H1,H2,H3,H4,H1Type,H2Type,H3Type,H4Type,username "
'--- Group Head & A/c H6 With Schedule
gCn.Execute X1 & " select distinct 'B','            ' " & ConcatOperator & "  'As Per Schedule ' " & ConcatOperator & " Schedule,sum(ClCr-ClDr) ,username,H1,H2,H3,H4,H5,schedule as H6 ,H1TYPE, H2TYPE,H3TYPE,H4Type,tmpGrpTrbalDetail.H5Type,1 from tmpGrpTrbalDetail where IsOpStkAc=0 and schedule " & NotNullStr & " and ScheduleHNo=5 and TopGroup=" & TopGrpcode & " and username='" & gUserName & "'" _
& " group by Schedule,H1,H2,H3,H4,H5,H1Type,H2Type,H3Type,H4Type,H5Type,username "
'--- Closing Stock Group Head
gCn.Execute X1 & " select distinct 'B','  ' " & ConcatOperator & "  tblMastGroup.AgName,0,'" & gUserName & "',tblG1.AgName as H1,tblMastGroup.AgName as H2,'' as H3,'' as H4,'' as H5,'' as H6,'G' as H1Type,'G' as H2Type,'-' as H3Type,'-' As H4Type,'-' as H5Type,0 from TblClosingStock,tblMastAccount,tblMastGroup TblG1,TblMastGroup where " _
& " tblMastGroup.Agcode=TblMastAccount.AgCode and stAccode=AcCode and tblG1.Agcode=tblMastGroup.Gpcode and  tblMastGroup.TopGroup=" & TopGrpcode & " and StAmount>0 and StDate=" & IIf(gBackEndDB = gBackEndAccess, "cDate('" & ToDt & "')", "To_Date('" & ToDt & "','dd/MM/yyyy')")
'--- Closing Stock A/c
gCn.Execute X1 & " select 'B','    ' " & ConcatOperator & "  AcName,-StAmount,'" & gUserName & "',tblG1.AgName as H1,tblMastGroup.AgName as H2,AcName as H3,'' as H4,'' as H5,'' as H6,'G' as H1Type,'G' as H2Type,'A' as H3Type,'-' As H4Type,'-' as H5Type,0 from TblClosingStock,tblMastAccount,tblMastGroup TblG1,TblMastGroup where " _
& " tblMastGroup.Agcode=TblMastAccount.AgCode and stAccode=AcCode and tblG1.Agcode=tblMastGroup.Gpcode and  tblMastGroup.TopGroup=" & TopGrpcode & " and StDate=" & IIf(gBackEndDB = gBackEndAccess, "cDate('" & ToDt & "')", "To_Date('" & ToDt & "','dd/MM/yyyy')")

'--- Profit & Loss of Current Yr
If TopGrpcode = 90002 Then
Dim Rs1 As New ADODB.Recordset
Dim Rs2 As New ADODB.Recordset
Dim gamt As Variant
  If GProcRstOpen(Rs1, "select nar1,Amt1 from tmpRptTbl where Nar1 = '  Profit And Loss of Current Year'", "R") Then
     If GProcRstOpen(Rs2, " select 'B','  Profit And Loss of Current Year',sum(PL)as PL,username,TblG1.AgName,tblMastGroup.AgName,'Profit And Loss of Current Year','','','','G','G','A','-','-',0 from qryPLCalculation,tblMastGroup TblG1,tblMastGroup where TblG1.AgCode =tblMastGroup.GpCode  and tblMastGroup.AgCode=90004 and tblMastGroup.TopGroup=" & TopGrpcode & " and username='" & gUserName & "' Group by tblG1.AgName,tblMastGroup.AgName,UserName", "R") <> 0 Then
          If Abs(Rs1!Amt1) = Rs2!pl Then
             gCn.Execute "Delete from tmpRptTbl where Nar1 = '  Profit And Loss of Current Year' "
          Else
             gCn.Execute "update tmpRptTbl set Amt1 = " & Rs2!pl - Abs(Rs1!Amt1) & " where Nar1 = '  Profit And Loss of Current Year' "
          End If
     End If
  Else
     gCn.Execute X1 & " select 'B','  Profit And Loss of Current Year',sum(PL),username,TblG1.AgName,tblMastGroup.AgName,'Profit And Loss of Current Year','','','','G','G','A','-','-',0 from qryPLCalculation,tblMastGroup TblG1,tblMastGroup where TblG1.AgCode =tblMastGroup.GpCode  and tblMastGroup.AgCode=90004 and tblMastGroup.TopGroup=" & TopGrpcode & " and username='" & gUserName & "' Group by tblG1.AgName,tblMastGroup.AgName,UserName"
  End If
End If
End Sub
'--- Profit Loss
Public Sub GProcPrepareProfitLoss(ToDate As Date)
Dim X As String
Dim RsTmp_S As Recordset
Dim RsTmp_T As Recordset
Dim i As Long
Dim RCount_T As Long
Dim RCount_S  As Long
'---- insert To Dt in tmpSelection
gCn.Execute "delete from tmpSelection where UserName='" & gUserName & "'"
If gBackEndDB = gBackEndAccess Then
    gCn.Execute "insert into tmpSelection (dt2,UserName) values ('" & ToDate & "','" & gUserName & "')"
Else
    gCn.Execute "insert into tmpSelection (dt2,UserName) values (to_date('" & ToDate & "','dd/MM/yyyy'),'" & gUserName & "') "
End If

gCn.Execute "delete from tmpGenTbl where UserName='" & gUserName & "'"
'--- Insert Top Debit Side In TmpGenTbl
'gCn.Execute "insert into tmpGenTbl  select Gp,Sr,AgName,AcName,Debit from QryPL2 where UserName='" & gUserName & "' and GP=1 Order by Gp,Sr,AgName,AcName"
X = "select Gp,GrNo,Sr,AcName,Debit,GrDebit,UserName from QryPL2_1 where UserName='" & gUserName & "' and ( Debit>0  or GrDebit>0 ) and gp=1 Order by Gp,Sr,AgName,GrNo,AcName"
If GProcRstOpen(RsTmp_S, X, "R") > 0 Then
    With RsTmp_S
        .MoveFirst
        Do While Not .EOF
           gCn.Execute "insert into tmpGenTbl (Amt10,Amt11,vno,VSrNo,Nar1,Amt1,Amt2,username) " _
           & " values (" & .AbsolutePosition & "," & .Fields(0) & "," & .Fields(1) & "," & .Fields(2) & ",'" & .Fields(3) & "'," & .Fields(4) & "," & .Fields(5) & ",'" & .Fields(6) & "')"
            .MoveNext
        Loop
    End With
End If
'--- Insert Top Credit Side In TmpGenTbl
RCount_T = GProcRstOpen(RsTmp_T, "Select * from TmpGenTbl where UserName='" & gUserName & "' order by Amt10 ", "O")
X = "select Gp,GrNo,Sr,AcName,Credit,GrCredit,UserName from QryPL2_1 where gp=1 and UserName='" & gUserName & "' and ( Credit>0  or GrCredit>0 ) Order by Gp,Sr,AgName,GrNo,AcName"
RCount_S = GProcRstOpen(RsTmp_S, X, "R")
If RCount_S > 0 Then
    With RsTmp_S
        .MoveFirst
        Do While Not .EOF
           i = i + 1
           If i > RCount_T Then
              gCn.Execute "insert into tmpGenTbl (Amt10,Amt11,vno,VSrNo,Nar2,Amt3,Amt4,username) " _
               & " values (" & i & "," & .Fields(0) & "," & .Fields(1) & "," & .Fields(2) & ",'" & .Fields(3) & "','" & .Fields(4) & "'," & .Fields(5) & ",'" & .Fields(6) & "')"
           Else
                gCn.Execute "update tmpGenTbl set Nar2='" & .Fields("AcName") & "',Amt3=" & .Fields("Credit") & ",Amt4=" & .Fields("GrCredit") & " where " _
                 & " Amt10=" & RsTmp_T.Fields("Amt10") & " and UserName='" & gUserName & "'"
           End If
            If Not RsTmp_T.EOF Then RsTmp_T.MoveNext
            .MoveNext
        Loop
    End With
End If
'--- Insert Bottom Debit Side In TmpGenTbl
RCount_T = GProcRstOpen(RsTmp_T, "Select * from TmpGenTbl where UserName='" & gUserName & "' order by Amt10 ", "O")
X = "select Gp,GrNo,Sr,AcName,Debit,GrDebit,UserName from QryPL2_2 where gp=2 and UserName='" & gUserName & "' and ( Debit>0  or GrDebit>0 ) Order by Gp,Sr,AgName,GrNo,AcName"
RCount_S = GProcRstOpen(RsTmp_S, X, "R")
i = RCount_T
If RCount_S > 0 Then
    With RsTmp_S
        .MoveFirst
        Do While Not .EOF
           i = i + 1
           'If i > RCount_T Then
              gCn.Execute "insert into tmpGenTbl (Amt10,Amt11,vno,VSrNo,Nar1,Amt1,Amt2,username) " _
               & " values (" & i & "," & .Fields(0) & "," & .Fields(1) & "," & .Fields(2) & ",'" & .Fields(3) & "','" & .Fields(4) & "'," & .Fields(5) & ",'" & .Fields(6) & "')"
           'Else
           '     gCn.Execute "update tmpGenTbl set Nar1='" & .Fields("AcName") & "',Amt1=" & .Fields("Debit") & ",Amt2=" & .Fields("GrDebit") & " where " _
           '      & " Amt10=" & RsTmp_T.Fields("Amt10") & " and UserName='" & gUserName & "'"
           'End If
           ' If Not RsTmp_T.EOF Then RsTmp_T.MoveNext
            .MoveNext
        Loop
    End With
End If
'--- Insert Bottom Credit Side In TmpGenTbl
RCount_T = GProcRstOpen(RsTmp_T, "Select * from TmpGenTbl where UserName='" & gUserName & "' and Amt11=2 order by Amt10 ", "O")
X = "select Gp,GrNo,Sr,AcName,Credit,GrCredit,UserName from QryPL2_2 where gp=2 and UserName='" & gUserName & "' and ( Credit>0  or GrCredit>0 ) Order by Gp,Sr,AgName,GrNo,AcName"
Dim j As Long
RCount_S = GProcRstOpen(RsTmp_S, X, "R")
If RCount_S > 0 Then
    With RsTmp_S
        .MoveFirst
        Do While Not .EOF
           j = j + 1
           If j > RCount_T Then
              i = i + 1
              gCn.Execute "insert into tmpGenTbl (Amt10,Amt11,vno,VSrNo,Nar2,Amt3,Amt4,username) " _
               & " values (" & i & "," & .Fields(0) & "," & .Fields(1) & "," & .Fields(2) & ",'" & .Fields(3) & "','" & .Fields(4) & "'," & .Fields(5) & ",'" & .Fields(6) & "')"
           Else
                gCn.Execute "update tmpGenTbl set Nar2='" & .Fields("AcName") & "',Amt3=" & .Fields("Credit") & ",Amt4=" & .Fields("GrCredit") & " where " _
                 & " Amt10=" & RsTmp_T.Fields("Amt10") & " and UserName='" & gUserName & "'"
           End If
            If Not RsTmp_T.EOF Then RsTmp_T.MoveNext
            .MoveNext
        Loop
    End With
End If
End Sub
'---- Balance Sheet
Public Sub GProcPrepareBalanceSheet(FromDt As Date, ToDate As Date)
Dim RsTmp_T As Recordset, TmpTot As Recordset
Dim RsTmp_S As Recordset
Dim RCount_T As Long
Dim RCount_S As Long
Dim i As Long
Dim X As String
Dim BalQuery As String
'---- insert From dt ,To Dt in tmpGenTbl
gCn.Execute "delete from tmpSelection where UserName='" & gUserName & "'"
If gBackEndDB = gBackEndAccess Then
    gCn.Execute "insert into tmpSelection (dt1,dt2,UserName) values ('" & FromDt & "','" & ToDate & "','" & gUserName & "')"
Else
    gCn.Execute "insert into tmpSelection (dt1,dt2,UserName) values (to_date('" & FromDt & "','dd/MM/yyyy'),to_date('" & ToDate & "','dd/MM/yyyy'),'" & gUserName & "') "
End If
gCn.Execute "delete from tmpGrpTrbalDetail where UserName='" & gUserName & "'"
gCn.Execute "insert into tmpGrpTrbalDetail  select * from qryGrpTrbalDetail where UserName='" & gUserName & "'"

gCn.Execute "delete from tmpGenTbl where UserName='" & gUserName & "'"

'=== Insert For Liability Side
GetBalRecordInTmpTbl 90002, ToDate
BalQuery = "Select Vtype,Nar1,Amt1,UserName,Nar2,Nar3,Nar4,Nar5,Nar6,Nar7,Nar8,Nar9,Nar10,Nar11,Nar12,Amt2 from tmpRptTbl " _
& " where UserName='" & gUserName & "' order by Nar2,Nar9,Nar3,Nar10,Nar4,Nar11,Nar5,Nar12,Nar6,Nar7"
If GProcRstOpen(RsTmp_S, BalQuery, "R") > 0 Then
    With RsTmp_S
           .MoveFirst
        Do While Not .EOF
           gCn.Execute "insert into tmpGenTbl (vno,Vtype,Nar13,Amt1,username,Amt11) " _
           & " values (" & .AbsolutePosition & ",'" & .Fields(0) & "','" & .Fields(1) & "'," & .Fields(2) & ",'" & .Fields(3) & "'," & .Fields(15) & ")"
        
        '====================== For Group Total
           If Val(.Fields(2)) = 0 Then
              GProcRstOpen TmpTot, "SELECT sum(Total),Grp From(SELECT sum(amt1) as Total, iif(nar10='A',Nar3,nar4) as Grp from tmprpttbl where username='" & gUserName & "' group by nar4,Nar3,Nar10) A where Grp='" & Trim(.Fields(1)) & "'  Group By Grp", "R"
              If Not TmpTot.EOF Then gCn.Execute "Update TmpGenTbl Set VsrNo=1,Amt1=" & TmpTot.Fields(0) & " where nar13='" & .Fields(1) & "'"
           End If
       '============================================End

           
           .MoveNext
           
        Loop
    End With
End If
'=== UpDate For Asset Side
RCount_T = GProcRstOpen(RsTmp_T, "Select * from TmpGenTbl where UserName='" & gUserName & "' and Vtype='B' order by vno " _
& "", "O")
GetBalRecordInTmpTbl 90001, ToDate
BalQuery = "Select Vtype,Nar1,Amt1,UserName,Nar2,Nar3,Nar4,Nar5,Nar6,Nar7,Nar8,Nar9,Nar10,Nar11,Nar12,Amt2 from tmpRptTbl " _
& " where UserName='" & gUserName & "' order by Nar2,Nar9,Nar3,Nar10,Nar4,Nar11,Nar5,Nar12,Nar6,Nar7"
RCount_S = GProcRstOpen(RsTmp_S, BalQuery, "R")
If RCount_S > 0 Then
    With RsTmp_S
        .MoveFirst
        Do While Not .EOF
           i = i + 1
           If i > RCount_T Then '- Right Side > Left Side 'Nar1,Nar2,Nar3,Nar4,Nar5,Nar6,Nar8,Nar9,Nar10,Nar11,Nar12,
                gCn.Execute "insert into tmpGenTbl (vno,Vtype,Nar7,Amt2,UserName,Amt12) values (" & i & ",'" & .Fields(0) & "'," _
                & "'" & .Fields(1) & "',-1*" & .Fields(2) & ",'" & .Fields(3) & "'," & .Fields(15) & ")"
           Else
                gCn.Execute "update tmpGenTbl set Vtype='" & .Fields("Vtype") & "',Nar7='" & .Fields("Nar1") & "',Amt2=-1*" & .Fields("amt1") & ",Amt12=" & .Fields("amt2") & " where " _
                & " username='" & .Fields("UserName") & "' and vno=" & RsTmp_T.Fields("vno")
           End If
            
           '======================Group Total
             If Val(.Fields(2)) = 0 Then
             
              GProcRstOpen TmpTot, "SELECT sum(Total),Grp From(SELECT sum(amt1) as Total, iif(nar10='A',Nar3,nar4) as Grp from tmprpttbl where username='" & gUserName & "' group by nar4,Nar3,Nar10) A where Grp='" & Trim(.Fields(1)) & "'  Group By Grp", "R"
              If Not TmpTot.EOF Then gCn.Execute "Update TmpGenTbl Set VbillNo=1,Amt2=-1*" & TmpTot.Fields(0) & " where nar7='" & .Fields(1) & "'"
           '==============================End
            End If
            
            If Not RsTmp_T.EOF Then RsTmp_T.MoveNext
            .MoveNext
        Loop
    End With
End If
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
AcCode = GProcGetColumnValue("TblMastAccount", "AcName", pAname, "S", "AcCode", "N")
If IsOnlyOPBalance = False Then
   If gBackEndDB = gBackEndAccess Then
       X = "Select sum(VAmt) from tblvoucher where VYear=" & gCYear & " and vtype<>'0P' and  VAcCode=" & AcCode & " and Vdt<=cdate('" & pOnDate & "')"
   Else '-- Oracle
        X = "Select sum(VAmt) from tblvoucher where VYear=" & gCYear & " and vtype<>'0P' and VAcCode=" & AcCode & " and Vdt<=to_date('" & pOnDate & "','dd/MM/yyyy')"
   End If
   GProcRstOpen rsBal, X, "R"
    If IsNumeric(rsBal.Fields(0)) = True Then
        GProcGetAccountBalance = CDbl(rsBal.Fields(0))
    End If
    X = "Select VAmt from tblvoucher where VYear=" & gCYear & " and VAcCode=" & AcCode & " and vtype='0P'"
    GProcRstOpen rsOP, X, "R"
    If IsNumeric(rsOP.Fields(0)) = True Then
        GProcGetAccountBalance = GProcGetAccountBalance + CDbl(rsOP.Fields(0))
    End If
Else
    X = "Select VAmt from tblvoucher where VYear=" & gCYear & " and VAcCode=" & AcCode & " and vtype='0P'"
    GProcRstOpen rsBal, X, "R"
    If IsNumeric(rsBal.Fields(0)) = True Then
        GProcGetAccountBalance = CDbl(rsBal.Fields(0))
    End If
End If
End Function
'------------------------- Prepare VAT Computation Details
Public Sub GProcPrepareVATComputDet(FromDt As Date, ToDt As Date, CrystalRpt As CrystalReport)
Dim CondStr As String
Dim X1 As String
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
        X1 = "{tmpGenTbl.Username}='" & gUserName & "'"
        .SelectionFormula = X1
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
        X1 = "{tmpGenTbl.Username}='" & gUserName & "'"
        .SelectionFormula = X1
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
If Party <> "" Then CondStr = CondStr & " and AcName='" & Party & "'" '--Party
'-- Is Interest DbNt
If IsIntDbNtRpt = True Then
    CondStr = CondStr & " and IntDbNtDate is Null" '-IntDbNtDt is Null
End If
X = " select ADJNO,RecVdt,AcName,InvBillNo,InvDt,BILLAMT,AdjAmt, BILLAMT-AdjAmt,INTRECEIVED+IntDbNtRecdAmt,RecVdt- " & IIf(IntLateFrmInvDt = 1, "INVDT-", "OUTDUEDATE-") & GraseDays & ",'" & gUserName & "'," & InterestP & "," & GraseDays & ", " _
& " OUTDUEDATE from tblmastaccount,TblOutstanding,TblRecVsSale where accode=PartyCode and " & CondStr & " And " _
& " vno=Invno and vtype=Invtp and vYear=InvYear "
'--- insert to setup db Table For All Firm
GProcInsertToSetupDbTbl AllFirmFlag, X, "tmpGenTbl", "Vno,vdt,Nar1,VBillNo,VBillDt,Amt1,amt2,Amt3,Amt4,amt5,username,amt6,amt7,Dt1", " username='" & gUserName & "'", "Nar2"
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
    CondStr = " vDt<= cdate('" & OnDt & "') and Vtype in ('SY','SM','SD','SO','ST') and OUTAMOUNT-OUTRECAMT > 0 "
    DueDayStr = "cdate('" & CDate(OnDt) & "') " & " - cdate(vDt) "
Else '-- Oracle
    CondStr = " OUTDUEDATE<= to_date('" & OnDt & "','dd/MM/yyyy') and Vtype in ('SY','SM','SD','SO','ST')"
    DueDayStr = "to_date('" & CDate(OnDt) & "','dd/MM/yyyy') - outDueDate"
End If
'-- Broker
If BrkCode > 0 Then CondStr = CondStr & " and OUTBROKERCODE=" & BrkCode
gCn.Execute "delete from tmpGenTbl where UserName='" & gUserName & "'"
If DueDaysTo = 0 Then
    '-- 0-15 days
    gCn.Execute "insert into tmpGentbl (VACCODE,VACOCODE,Amt1,Amt2,Amt3,Amt4,Amt5,UserName)" _
    & " select OUTBROKERCODE,OUTACCODE,Sum(OUTAMOUNT-OUTRECAMT) as Due15,0 as Due30,0 as Due45,0 as " _
    & " DueAbove,0,'" & gUserName & "' from tbloutstanding where " & CondStr & " and " & DueDayStr & " >= 0 and " & DueDayStr & " <30 group by OUTBROKERCODE,OUTACCODE"
    '-- 15-30 days
    gCn.Execute "insert into tmpGentbl (VACCODE,VACOCODE,Amt1,Amt2,Amt3,Amt4,Amt5,UserName)" _
    & " select OUTBROKERCODE,OUTACCODE,0,Sum(OUTAMOUNT-OUTRECAMT),0,0,0 " _
    & ",'" & gUserName & "' from tbloutstanding where " & CondStr & " and " & DueDayStr & ">=30 and " & DueDayStr & "<60 group by OUTBROKERCODE,OUTACCODE"
    '-- 30 -45 days
    gCn.Execute "insert into tmpGentbl (VACCODE,VACOCODE,Amt1,Amt2,Amt3,Amt4,Amt5,UserName)" _
    & " select OUTBROKERCODE,OUTACCODE,0,0,Sum(OUTAMOUNT-OUTRECAMT),0,0" _
    & ",'" & gUserName & "' from tbloutstanding where " & CondStr & " and " & DueDayStr & ">=60 and " & DueDayStr & "<90 group by OUTBROKERCODE,OUTACCODE"
    '-- 45 above days
    gCn.Execute "insert into tmpGentbl (VACCODE,VACOCODE,Amt1,Amt2,Amt3,Amt4,Amt5,UserName)" _
    & " select OUTBROKERCODE,OUTACCODE,0,0,0,Sum(OUTAMOUNT-OUTRECAMT),0" _
    & ",'" & gUserName & "' from tbloutstanding where " & CondStr & " and " & DueDayStr & ">=90 group by OUTBROKERCODE,OUTACCODE"
    
    
    gCn.Execute "delete from tmpGenTbl where username='" & gUserName & "' and Amt1=0 and Amt2=0 and " _
    & " amt3=0 and amt4=0"
Else
    '--  FromDays to ToDays
    gCn.Execute "insert into tmpGentbl (VACCODE,VACOCODE,Amt1,Vdt,Amt2,Amt3,UserName)" _
    & " select OUTBROKERCODE,OUTACCODE,Sum(OUTAMOUNT-OUTRECAMT) as Due," & IIf(gBackEndDB = gBackEndAccess, CDate(OnDt), "to_date('" & CDate(OnDt) & "','dd/MM/yyyy')") & " as OnDt," _
    & DueDaysFrom & "," & DueDaysTo & ",'" & gUserName & "' from tbloutstanding where " & CondStr & " and (" & DueDayStr & ") >= " & DueDaysFrom & "  and (" & DueDayStr & " )<= " & DueDaysTo & " group by OUTBROKERCODE,OUTACCODE"
    
    gCn.Execute "delete from tmpGenTbl where username='" & gUserName & "' and Amt1=0 "
End If
''Dim CondStr As String
''Dim BrkCode As Long
''Dim DueDayStr As String
''BrkCode = GProcGetColumnValue("tblMastAccount", "AcName", BrkAcName, "T", "AcCode", "N")
''If gBackEndDB = gBackEndAccess Then
''    CondStr = " OUTDUEDATE<= cdate('" & OnDt & "') and Vtype in ('SY','SM','SD','SO','ST')"
''    DueDayStr = "cdate('" & CDate(OnDt) & "') " & " - cdate(outDueDate) "
''Else '-- Oracle
''    CondStr = " OUTDUEDATE<= to_date('" & OnDt & "','dd/MM/yyyy') and Vtype in ('SY','SM','SD','SO','ST')"
''    DueDayStr = "to_date('" & CDate(OnDt) & "','dd/MM/yyyy') - outDueDate"
''End If
'''-- Broker
''If BrkCode > 0 Then CondStr = CondStr & " and OUTBROKERCODE=" & BrkCode
''gCn.Execute "delete from tmpGenTbl where UserName='" & gUserName & "'"
''If DueDaysTo = 0 Then
''    '-- 0-15 days
''    gCn.Execute "insert into tmpGentbl (VACCODE,VACOCODE,Amt1,Amt2,Amt3,Amt4,UserName)" _
''    & " select OUTBROKERCODE,OUTACCODE,Sum(OUTAMOUNT-OUTRECAMT) as Due15,0 as Due30,0 as Due45,0 as " _
''    & " DueAbove,'" & gUserName & "' from tbloutstanding where " & CondStr & " and " & DueDayStr & " > 0 and " & DueDayStr & " <15 group by OUTBROKERCODE,OUTACCODE"
''    '-- 15-30 days
''    gCn.Execute "insert into tmpGentbl (VACCODE,VACOCODE,Amt1,Amt2,Amt3,Amt4,UserName)" _
''    & " select OUTBROKERCODE,OUTACCODE,0,Sum(OUTAMOUNT-OUTRECAMT),0,0 " _
''    & ",'" & gUserName & "' from tbloutstanding where " & CondStr & " and " & DueDayStr & ">=15 and " & DueDayStr & "<30 group by OUTBROKERCODE,OUTACCODE"
''    '-- 30 -45 days
''    gCn.Execute "insert into tmpGentbl (VACCODE,VACOCODE,Amt1,Amt2,Amt3,Amt4,UserName)" _
''    & " select OUTBROKERCODE,OUTACCODE,0,0,Sum(OUTAMOUNT-OUTRECAMT),0" _
''    & ",'" & gUserName & "' from tbloutstanding where " & CondStr & " and " & DueDayStr & ">=30 and " & DueDayStr & "<45 group by OUTBROKERCODE,OUTACCODE"
''    '-- 45 above days
''    gCn.Execute "insert into tmpGentbl (VACCODE,VACOCODE,Amt1,Amt2,Amt3,Amt4,UserName)" _
''    & " select OUTBROKERCODE,OUTACCODE,0,0,0,Sum(OUTAMOUNT-OUTRECAMT)" _
''    & ",'" & gUserName & "' from tbloutstanding where " & CondStr & " and " & DueDayStr & ">=45 group by OUTBROKERCODE,OUTACCODE"
''
''    gCn.Execute "delete from tmpGenTbl where username='" & gUserName & "' and Amt1=0 and Amt2=0 and " _
''    & " amt3=0 and amt4=0"
''Else
''    '--  FromDays to ToDays
''    gCn.Execute "insert into tmpGentbl (VACCODE,VACOCODE,Amt1,Vdt,Amt2,Amt3,UserName)" _
''    & " select OUTBROKERCODE,OUTACCODE,Sum(OUTAMOUNT-OUTRECAMT) as Due," & IIf(gBackEndDB = gBackEndAccess, CDate(OnDt), "to_date('" & CDate(OnDt) & "','dd/MM/yyyy')") & " as OnDt," _
''    & DueDaysFrom & "," & DueDaysTo & ",'" & gUserName & "' from tbloutstanding where " & CondStr & " and (" & DueDayStr & ") >= " & DueDaysFrom & "  and (" & DueDayStr & " )<= " & DueDaysTo & " group by OUTBROKERCODE,OUTACCODE"
''
''    gCn.Execute "delete from tmpGenTbl where username='" & gUserName & "' and Amt1=0 "
''End If
End Sub
'--------- Prepare VAT Computation
Public Sub GProcPrepareVATComputation(FromDt As Date, ToDt As Date, CrystalRpt As CrystalReport, OpBal As Double)
Dim CondStr As String
Dim X As String
Dim Rs As Recordset
Dim i As Long
If gBackEndDB = gBackEndAccess Then
    CondStr = " Vdt>= cdate('" & FromDt & "')  and  Vdt<=cdate('" & ToDt & "')"
Else '-- Oracle
    CondStr = " Vdt>= to_date('" & CDate(FromDt) & "','dd/MM/yyyy')  and Vdt<=to_date('" & CDate(ToDt) & "','dd/MM/yyyy')"
End If
gCn.Execute "delete from tmpGentbl where username='" & gUserName & "'"
With CrystalRpt
    '=======================Sale
    If gBackEndDB = gBackEndAccess Then  '--Access
'        X = "select sum(AdBillAmt) as SAAmt,sum(ADTAXAMT) as SATaxAmt, " _
'        & " sum(iif(Narration='V A T',iif(VType='SY',iif(ADEXEMPTAMT=0,ADTAXAMT,0),0),0)) as SATradeTAX , " _
'        & " sum(iif(Narration='V A T',iif(VType='SO',ADTAXAMT,0),0)) as SAConsigTAX ," _
'        & " sum(iif(Narration='V A T',iif(VType='SDD',ADTAXAMT,0),0)) as SADepotTAX ," _
'        & " sum(iif(Narration='V A T',iif(VType='SY',iif(ADEXEMPTAMT=0,0,ADTAXAMT),0),0)) as SAExmptTAX ," _
'        & " sum(iif(VType='ST',AdBillAmt,0)) as SASIT ," _
'        & " sum(iif(Narration='Tax Free',AdBillAmt,0)) as SATAXFREE, " _
'        & " sum(iif(Narration='C S T',AdBillAmt,0)) as SACST, " _
'        & " sum(ADEXEMPTAMT) as SAExmpt,sum(iif(Narration='C S T',ADTAXAMT,0)) as SaCSTTax" _
'        & " from tbladdless,tblMastNarration where ADTAXCODE=Narrcode and " & CondStr & " and vtype in ('SY','ST','SO',)"
    
'---For Tax 4 & 5 %
         X = "select sum(AdBillAmt) as SAAmt,sum(ADTAXAMT) as SATaxAmt, " _
         & " sum(iif(Narration='V A T',iif(VType='SY',iif(ADEXEMPTAMT=0,iif(AdTaxp=4,ADTAXAMT,0),0),0),0)) as SATradeTAX4 , " _
         & " sum(iif(Narration='V A T',iif(VType='SY',iif(ADEXEMPTAMT=0,iif(AdTaxp=5,ADTAXAMT,0),0),0),0)) as SATradeTAX5 , " _
         & " sum(iif(Narration='V A T',iif(VType='SY',iif(ADEXEMPTAMT=0,iif(AdTaxp=4,AdAssVal,0),0),0),0)) as SATradeAssVal4 , " _
         & " sum(iif(Narration='V A T',iif(VType='SY',iif(ADEXEMPTAMT=0,iif(AdTaxp=5,AdAssVal,0),0),0),0)) as SATradeAssVal5 , " _
         & " sum(iif(Narration='V A T',iif(VType='SO',iif(AdTaxp=4,ADTAXAMT,0),0),0)) as SAConsigTAX4 , " _
         & " sum(iif(Narration='V A T',iif(VType='SO',iif(AdTaxp=5,ADTAXAMT,0),0),0)) as SAConsigTAX5 , " _
         & " sum(iif(Narration='V A T',iif(VType='SO',iif(AdTaxp=4,AdAssVal,0),0),0)) as SAConsigAssVal4 , " _
         & " sum(iif(Narration='V A T',iif(VType='SO',iif(AdTaxp=5,AdAssVal,0),0),0)) as SAConsigAssVal5 , " _
         & " sum(iif(Narration='V A T',iif(VType='SD',iif(AdTaxp=4,ADTAXAMT,0),0),0)) as SADepotTAX4 , " _
         & " sum(iif(Narration='V A T',iif(VType='SD',iif(AdTaxp=5,ADTAXAMT,0),0),0)) as SADepotTAX5 , " _
         & " sum(iif(Narration='V A T',iif(VType='SD',iif(AdTaxp=4,AdAssVal,0),0),0)) as SADepotAssVal4 , " _
         & " sum(iif(Narration='V A T',iif(VType='SD',iif(AdTaxp=5,AdAssVal,0),0),0)) as SADepotAssVal5 , " _
         & " sum(iif(Narration='V A T',iif(VType='SY',iif(ADEXEMPTAMT>0,iif(AdTaxp=4,ADTAXAMT,0),0),0),0)) as SAExmptTAX4 , " _
         & " sum(iif(Narration='V A T',iif(VType='SY',iif(ADEXEMPTAMT>0,iif(AdTaxp=5,ADTAXAMT,0),0),0),0)) as SAExmptTAX5 , " _
         & " sum(iif(Narration='V A T',iif(VType='SY',iif(ADEXEMPTAMT>0,iif(AdTaxp=4,ADEXEMPTAMT,0),0),0),0)) as SAExmpt4, " _
         & " sum(iif(Narration='V A T',iif(VType='SY',iif(ADEXEMPTAMT>0,iif(AdTaxp=5,ADEXEMPTAMT,0),0),0),0)) as SAExmpt5, " _
         & " sum(iif(Narration='V A T',iif(VType='SY',iif(ADEXEMPTAMT>0,iif(AdTaxp=4,AdAssVal,0),0),0),0)) as SAExmptAssVal4 , " _
         & " sum(iif(Narration='V A T',iif(VType='SY',iif(ADEXEMPTAMT>0,iif(AdTaxp=5,AdAssVal,0),0),0),0)) as SAExmptAssVal5 , " _
         & " sum(iif(VType='ST',AdBillAmt,0)) as SASIT , " _
         & " sum(iif(Narration='Tax Free',AdBillAmt,0)) as SATAXFREE, " _
         & " sum(iif(Narration='C S T',AdBillAmt,0)) as SACST, " _
         & " sum(iif(Narration='C S T',ADTAXAMT,0)) as SaCSTTax, sum(IIf(VType='SY' And ADEXEMPTAMT>0,AdBillAmt,0)) as SAExemptSale , " _
         & " sum(iif(Narration='V A T',iif(VType='SY',iif(ADEXEMPTAMT=0,AdBillAmt,0),0),0)) as SATradeSale ,sum(iif(Narration='V A T',iif(VType='SO',iif(ADEXEMPTAMT=0,AdBillAmt,0),0),0)) as SAConsiSale ,sum(iif(Narration='V A T',iif(VType='SD',iif(ADEXEMPTAMT=0,AdBillAmt,0),0),0)) as SADepotSale  " _
         & " from tbladdless,tblMastNarration where ADTAXCODE=Narrcode and " & CondStr & " and vtype in ('SY','ST','SO','SD')"
'-----
    Else
        X = "select nvl(sum(AdBillAmt),0) as SAAmt,nvl(sum(ADTAXAMT),0) as SATaxAmt, " _
        & " nvl(sum(decode(Narration,'V A T',decode(VType,'SY',decode(ADEXEMPTAMT,0,ADTAXAMT,0),0),0)),0) as SATradeTAX , " _
        & " nvl(sum(decode(Narration,'V A T',decode(VType,'SO',ADTAXAMT,0),0)),0) as SAConsigTAX ," _
        & " nvl(sum(decode(Narration,'V A T',decode(VType,'SD',ADTAXAMT,0),0)),0) as SADepotTAX ," _
        & " nvl(sum(decode(Narration,'V A T',decode(VType,'SY',decode(ADEXEMPTAMT,0,0,ADTAXAMT),0),0)),0) as SAExmptTAX ," _
        & " nvl(sum(decode(VType,'ST',AdBillAmt,0)),0) as SASIT ," _
        & " nvl(sum(decode(Narration,'C S T',AdBillAmt,0)),0) as SACST, " _
        & " nvl(sum(ADEXEMPTAMT),0) as SAExmpt,nvl(sum(decode(Narration,'C S T',ADTAXAMT,0)),0) as SaCSTTax" _
        & " from tbladdless,tblMastNarration where ADTAXCODE=Narrcode and " & CondStr & " and vtype in ('SY','ST','SO','SD')"
    End If
        If GProcRstOpen(Rs, X, "R") > 0 Then
             .Formulas(4) = "SAAmt ='" & Format(IIf(IsNull(Rs!SAAmt), 0#, Rs!SAAmt), FStr) & "'" '-- GTO Sale
             .Formulas(5) = "SATaxAmt ='" & Format(IIf(IsNull(Rs!SATaxAmt), 0#, Rs!SATaxAmt), FStr) & "'" '-- SA Tax Amt
    '         .Formulas(6) = "SATradeTAX ='" & Format(IIf(IsNull(Rs!SATradeTAX), 0#, Rs!SATradeTAX), FStr) & "'"  '-- SA Trade TAX
'             .Formulas(7) = "SAConsigTAX ='" & Format(IIf(IsNull(Rs!SAConsigTAX), 0#, Rs!SAConsigTAX), FStr) & "'" '-- SA Consig TAX
 '            .Formulas(8) = "SADepotTAX ='" & Format(IIf(IsNull(Rs!SADepotTAX), 0#, Rs!SADepotTAX), FStr) & "'" '-- SA Depot TAX
  '           .Formulas(9) = "SAExmptTAX ='" & Format(IIf(IsNull(Rs!SAExmptTAX), 0#, Rs!SAExmptTAX), FStr) & "'" '-- SA Exmpt TAX
             .Formulas(10) = "SASIT ='" & Format(IIf(IsNull(Rs!SASIT), 0#, Rs!SASIT), FStr) & "'" '-- SA SIT
             .Formulas(11) = "SACST ='" & Format(IIf(IsNull(Rs!SACST), 0#, Rs!SACST), FStr) & "'" '--SA CST
'             .Formulas(12) = "SAExmpt ='" & Format(IIf(IsNull(Rs!SAExmpt), 0#, Rs!SAExmpt), FStr) & "'" '--SA Exmpt Amt
             .Formulas(28) = "SACstTax ='" & Format(IIf(IsNull(Rs!SACstTax), 0#, Rs!SACstTax), FStr) & "'" '--SA Cst Tax
             .Formulas(29) = "SATaxFree ='" & Format(IIf(IsNull(Rs!SATaxFree), 0#, Rs!SATaxFree), FStr) & "'" '--SA Tax Free
            
             .Formulas(30) = "SATradeTAX4 ='" & Format(IIf(IsNull(Rs!SATradeTAX4), 0#, Rs!SATradeTAX4), FStr) & "'" '-- Trade 4 % Tax
             .Formulas(31) = "SATradeTAX5 ='" & Format(IIf(IsNull(Rs!SATradeTAX5), 0#, Rs!SATradeTAX5), FStr) & "'" '-- Trade 5 % Tax
             .Formulas(32) = "SATradeAssVal4 ='" & Format(IIf(IsNull(Rs!SATradeAssVal4), 0#, Rs!SATradeAssVal4), FStr) & "'" '-- Trade 4 % Ass Val
             .Formulas(33) = "SATradeAssVal5 ='" & Format(IIf(IsNull(Rs!SATradeAssVal5), 0#, Rs!SATradeAssVal5), FStr) & "'" '-- Trade 4 % Ass Val
             .Formulas(34) = "SAConsigTAX4 ='" & Format(IIf(IsNull(Rs!SAConsigTAX4), 0#, Rs!SAConsigTAX4), FStr) & "'" '-- Consi 4 % Tax
             .Formulas(35) = "SAConsigTAX5 ='" & Format(IIf(IsNull(Rs!SAConsigTAX5), 0#, Rs!SAConsigTAX5), FStr) & "'" '-- Consi 5 % Tax
             .Formulas(36) = "SAConsigAssVal4 ='" & Format(IIf(IsNull(Rs!SAConsigAssVal4), 0#, Rs!SAConsigAssVal4), FStr) & "'" '-- Consi 4 % Ass Val
             .Formulas(37) = "SAConsigAssVal5 ='" & Format(IIf(IsNull(Rs!SAConsigAssVal5), 0#, Rs!SAConsigAssVal5), FStr) & "'" '-- Consi 5 % Ass Val
             .Formulas(38) = "SADepotTAX4 ='" & Format(IIf(IsNull(Rs!SADepotTAX4), 0#, Rs!SADepotTAX4), FStr) & "'" '-- Depot 4 % Tax
             .Formulas(39) = "SADepotTAX5 ='" & Format(IIf(IsNull(Rs!SADepotTAX5), 0#, Rs!SADepotTAX5), FStr) & "'" '-- Depot 5 % Tax
             .Formulas(40) = "SADepotAssVal4 ='" & Format(IIf(IsNull(Rs!SADepotAssVal4), 0#, Rs!SADepotAssVal4), FStr) & "'" '-- Depot 4 % Ass Val
             .Formulas(41) = "SADepotAssVal5 ='" & Format(IIf(IsNull(Rs!SADepotAssVal5), 0#, Rs!SADepotAssVal5), FStr) & "'" '-- Depot 5 % Ass Val
             .Formulas(42) = "SAExmptTAX4 ='" & Format(IIf(IsNull(Rs!SAExmptTAX4), 0#, Rs!SAExmptTAX4), FStr) & "'" '-- Exempt 4 % Tax
             .Formulas(43) = "SAExmptTAX5 ='" & Format(IIf(IsNull(Rs!SAExmptTAX5), 0#, Rs!SAExmptTAX5), FStr) & "'" '-- Exempt 5 % Tax
             .Formulas(44) = "SAExmptAssVal4 ='" & Format(IIf(IsNull(Rs!SAExmptAssVal4), 0#, Rs!SAExmptAssVal4), FStr) & "'" '-- Exempt 4 % Ass Val
             .Formulas(45) = "SAExmptAssVal5 ='" & Format(IIf(IsNull(Rs!SAExmptAssVal5), 0#, Rs!SAExmptAssVal5), FStr) & "'" '-- Exempt 5 % Ass Val
             .Formulas(46) = "SAExmpt4 ='" & Format(IIf(IsNull(Rs!SAExmpt4), 0#, Rs!SAExmpt4), FStr) & "'" '-- Exempt 4 % Purchase Prise
             .Formulas(47) = "SAExmpt5 ='" & Format(IIf(IsNull(Rs!SAExmpt5), 0#, Rs!SAExmpt5), FStr) & "'" '-- Exempt 5 % Purchase Prise
            
             .Formulas(48) = "SATradeSale ='" & Format(IIf(IsNull(Rs!SATradeSale), 0#, Rs!SATradeSale), FStr) & "'" '-- Total Trade Sale
             .Formulas(49) = "SAConsiSale ='" & Format(IIf(IsNull(Rs!SAConsiSale), 0#, Rs!SAConsiSale), FStr) & "'" '-- Total Consign Sale
             .Formulas(50) = "SADepotSale ='" & Format(IIf(IsNull(Rs!SADepotSale), 0#, Rs!SADepotSale), FStr) & "'" '-- Total Depot Sale
             .Formulas(52) = "SAExemptSale ='" & Format(IIf(IsNull(Rs!SAExemptSale), 0#, Rs!SAExemptSale), FStr) & "'" '-- Total Exempt Sale
            '.Action = 1
        End If
    '=======================Purchase
    If gBackEndDB = gBackEndAccess Then  '--Access
        X = "select sum(AdBillAmt) as PUAmt,sum(iif(Vtype='PI',AdBillAmt,0)) as PUPIAmt, " _
        & " sum(iif(VType='PY',ADEXEMPTAMT,0)) as PUPYExmpt , " _
        & " sum(iif(Narration<>'V A T' and Narration <>'Tax Free',iif(VType='PY',AdBillAmt,0),0)) as PUOms ," _
        & " sum(iif(Narration='Tax Free',iif(VType='PY',AdBillAmt,0),0)) as PUTaxFree ," _
        & " sum(iif(VType='PT',AdBillAmt,0)) as PUOMSSit ," _
        & " sum(iif(Narration='V A T',iif(VType='PY',AdTaxAmt,0),0)) as PUTaxPaid ," _
        & " sum(iif(Narration='V A T',iif(VType='PY',iif(ADEXEMPTAMT=0,iif(AdTaxp=4,AdTaxAmt,0),0),0),0)) as PUTradePurchTaxPaid4 ," _
        & " sum(iif(Narration='V A T',iif(VType='PY',iif(ADEXEMPTAMT=0,iif(AdTaxp=5,AdTaxAmt,0),0),0),0)) as PUTradePurchTaxPaid5 ," _
        & " sum(iif(Narration='V A T',iif(VType='PY',iif(ADEXEMPTAMT=0,iif(AdTaxp=4,AdAssVal,0),0),0),0)) as PUTradePurchAssVal4 ," _
        & " sum(iif(Narration='V A T',iif(VType='PY',iif(ADEXEMPTAMT=0,iif(AdTaxp=5,AdAssVal,0),0),0),0)) as PUTradePurchAssVal5 ," _
        & " sum(iif(Narration='V A T',iif(VType='PY',iif(ADEXEMPTAMT=0,AdBillAmt,0),0),0)) as PUTradePurch , " _
        & " sum(iif(Narration='V A T',iif(VType='PY',iif(ADEXEMPTAMT>0,iif(AdTaxp=4,AdTaxAmt,0),0),0),0)) as PUExemptPurchTaxPaid4 ," _
        & " sum(iif(Narration='V A T',iif(VType='PY',iif(ADEXEMPTAMT>0,iif(AdTaxp=5,AdTaxAmt,0),0),0),0)) as PUExemptPurchTaxPaid5 ," _
        & " sum(iif(Narration='V A T',iif(VType='PY',iif(ADEXEMPTAMT>0,iif(AdTaxp=4,AdAssVal,0),0),0),0)) as PUExemptPurchAssVal4 ," _
        & " sum(iif(Narration='V A T',iif(VType='PY',iif(ADEXEMPTAMT>0,iif(AdTaxp=5,AdAssVal,0),0),0),0)) as PUExemptPurchAssVal5 ," _
        & " sum(iif(Narration='V A T',iif(VType='PY',iif(ADEXEMPTAMT>0,AdBillAmt,0),0),0)) as PUExemptPurch , " _
        & " sum(iif(Narration='V A T',iif(VType='PY',iif(ADEXEMPTAMT>0,iif(AdTaxp=4,ADEXEMPTAMT,0),0),0),0)) as PUExemptAmt4 ," _
        & " sum(iif(Narration='V A T',iif(VType='PY',iif(ADEXEMPTAMT>0,iif(AdTaxp=5,ADEXEMPTAMT,0),0),0),0)) as PUExemptAmt5 ," _
        & " sum(ADTaxAMT) as PUTax,sum(ADAFTTAXADD) as PULess,sum(ADAFTTAXLESS) as PUAdd, " _
        & " sum(iif(Narration='V A T',iif(VType='PO',AdBillAmt,0),0)) as PUTOtherPurch , " _
        & " sum(iif(Narration='V A T',iif(VType='PO',iif(AdTaxp=4,AdAssVal,0),0),0)) as PUEOtherPurchAssVal4 ," _
        & " sum(iif(Narration='V A T',iif(VType='PO',iif(AdTaxp=5,AdAssVal,0),0),0)) as PUEOtherPurchAssVal5 ," _
        & " sum(iif(Narration='V A T',iif(VType='PO',iif(AdTaxp=12.5,AdAssVal,0),0),0)) as PUEOtherPurchAssVal125, sum(iif(Narration='V A T',iif(VType='PO',iif(AdTaxp=4,AdTaxAmt,0),0),0)) as PUEOtherPurchTaxPaid4 , sum(iif(Narration='V A T',iif(VType='PO',iif(AdTaxp=5,AdTaxAmt,0),0),0)) as PUEOtherPurchTaxPaid5 ," _
        & " sum(iif(Narration='V A T',iif(VType='PO',iif(AdTaxp=12.5,AdTaxAmt,0),0),0)) as PUEOtherPurchTaxPaid125 " _
        & " from tbladdless,tblMastNarration,tblMastAccount where  tblAddless.AdMillCode = tblMastAccount.AcCode And tblMastAccount.AcMillType <> 2 and tblAddless.ADTAXCODE=tblMastNarration.Narrcode and " & CondStr & " and tblAddless.vtype in ('PY','PT','PI','PO')"
    
'        & " from tbladdless,tblMastNarration where ADTAXCODE=Narrcode and " & CondStr & " and vtype in ('PY','PT','PI')"
    
    Else
        X = "select nvl(sum(AdBillAmt),0) as PUAmt,nvl(sum(decode(Vtype,'PI',AdBillAmt,0)),0) as PUPIAmt, " _
        & " nvl(sum(decode(VType,'PY',ADEXEMPTAMT,0)),0) as PUPYExmpt , " _
        & " nvl(sum(decode(Narration,'V A T',0,'Tax Free',0,decode(VType,'PY',AdBillAmt,0))),0) as PUOms ," _
        & " nvl(sum(decode(Narration,'Tax Free',decode(VType,'PY',AdBillAmt,0),0)),0) as PUTaxFree ," _
        & " nvl(sum(decode(VType,'PT',AdBillAmt,0)),0) as PUOMSSit ," _
        & " nvl(sum(decode(Narration,'V A T',decode(VType,'PY',AdTaxAmt,0),0)),0) as PUTaxPaid ," _
        & " nvl(sum(ADTaxAMT),0) as PUTax,nvl(sum(ADAFTTAXADD),0) as PULess,nvl(sum(ADAFTTAXLESS),0) as PUAdd" _
        & " from tbladdless,tblMastNarration where ADTAXCODE=Narrcode and " & CondStr & " and vtype in ('PY','PT','PI')"
    End If
        If GProcRstOpen(Rs, X, "R") > 0 Then '
             .Formulas(13) = "PUAmt ='" & Format(IIf(IsNull(Rs!PUAmt), 0#, Rs!PUAmt), FStr) & "'" '-- PUAmt (GTO)
             .Formulas(14) = "PUPIAmt ='" & Format(IIf(IsNull(Rs!PUPIAmt), 0#, Rs!PUPIAmt), FStr) & "'" '--PUPIAmt  (Stock Trans OMS)
             .Formulas(15) = "PUPYExmpt ='" & Format(IIf(IsNull(Rs!PUPYExmpt), 0#, Rs!PUPYExmpt), FStr) & "'" '-- PUPY Exmpt ( R.D. Exmpt)
             .Formulas(16) = "PUOms ='" & Format(IIf(IsNull(Rs!PUOms), 0#, Rs!PUOms), FStr) & "'" '-- PUOms ( OMS )
             .Formulas(17) = "PUTaxFree ='" & Format(IIf(IsNull(Rs!PUTaxFree), 0#, Rs!PUTaxFree), FStr) & "'" '-- PUTaxFree (Tax Free)
             .Formulas(18) = "PUOMSSit ='" & Format(IIf(IsNull(Rs!PUOMSSit), 0#, Rs!PUOMSSit), FStr) & "'" '-- PUOMSSit (OMS SIT)
             .Formulas(19) = "PUTaxPaid ='" & Format(IIf(IsNull(Rs!PUTaxPaid), 0#, Rs!PUTaxPaid), FStr) & "'" '-- PUTaxPaid (Tax Paid)
             .Formulas(26) = "PUOtherLess ='" & Format(IIf(IsNull(Rs!PULess), 0#, Rs!PULess), FStr) & "'" '-- PUOtherLess (Add Aft Tax)
             .Formulas(27) = "PUOtherAdd ='" & Format(IIf(IsNull(Rs!PUAdd), 0#, Rs!PUAdd), FStr) & "'" '-- PUOtherAdd (Less Aft Tax)
             
             .Formulas(53) = "PUTradePurchTaxPaid4 ='" & Format(IIf(IsNull(Rs!PUTradePurchTaxPaid4), 0#, Rs!PUTradePurchTaxPaid4), FStr) & "'" '-- Purchase tax paid 4 %
             .Formulas(54) = "PUTradePurchTaxPaid5 ='" & Format(IIf(IsNull(Rs!PUTradePurchTaxPaid5), 0#, Rs!PUTradePurchTaxPaid5), FStr) & "'" '-- Purchase tax paid 5 %
             .Formulas(55) = "PUTradePurchAssVal4 ='" & Format(IIf(IsNull(Rs!PUTradePurchAssVal4), 0#, Rs!PUTradePurchAssVal4), FStr) & "'" '-- Purchase Ass Val 4 %
             .Formulas(56) = "PUTradePurchAssVal5 ='" & Format(IIf(IsNull(Rs!PUTradePurchAssVal5), 0#, Rs!PUTradePurchAssVal5), FStr) & "'" '-- Purchase Ass Val 5 %
             .Formulas(57) = "PUTradePurch ='" & Format(IIf(IsNull(Rs!PUTradePurch), 0#, Rs!PUTradePurch), FStr) & "'" '-- Purchase Trade Amt
            
             .Formulas(58) = "PUExemptPurchTaxPaid4 ='" & Format(IIf(IsNull(Rs!PUExemptPurchTaxPaid4), 0#, Rs!PUExemptPurchTaxPaid4), FStr) & "'" '-- Purchase Exempt Tax Paid 4 %
             .Formulas(59) = "PUExemptPurchTaxPaid5 ='" & Format(IIf(IsNull(Rs!PUExemptPurchTaxPaid5), 0#, Rs!PUExemptPurchTaxPaid5), FStr) & "'" '-- Purchase Exempt Tax Paid 5 %
             .Formulas(60) = "PUExemptPurchAssVal4 ='" & Format(IIf(IsNull(Rs!PUExemptPurchAssVal4), 0#, Rs!PUExemptPurchAssVal4), FStr) & "'" '-- Purchase Exempt Ass Val 4 %
             .Formulas(61) = "PUExemptPurchAssVal5 ='" & Format(IIf(IsNull(Rs!PUExemptPurchAssVal5), 0#, Rs!PUExemptPurchAssVal5), FStr) & "'" '-- Purchase Exempt Ass Val 5 %
             .Formulas(62) = "PUExemptPurch ='" & Format(IIf(IsNull(Rs!PUExemptPurch), 0#, Rs!PUExemptPurch), FStr) & "'" '--  Exempt Purchase Amt
             .Formulas(63) = "PUExemptAmt4 ='" & Format(IIf(IsNull(Rs!PUExemptAmt4), 0#, Rs!PUExemptAmt4), FStr) & "'" '-- Purchase Exempt Amt 4 %
             .Formulas(64) = "PUExemptAmt5 ='" & Format(IIf(IsNull(Rs!PUExemptAmt5), 0#, Rs!PUExemptAmt5), FStr) & "'" '-- Purchase Exempt Amt 5 %
            
             
             .Formulas(65) = "PUTOtherPurch ='" & Format(IIf(IsNull(Rs!PUTOtherPurch), 0#, Rs!PUTOtherPurch), FStr) & "'" '-- Purchase Other Val
             .Formulas(66) = "PUEOtherPurchAssVal4 ='" & Format(IIf(IsNull(Rs!PUEOtherPurchAssVal4), 0#, Rs!PUEOtherPurchAssVal4), FStr) & "'" '-- Purchase Other Ass Val 4 %
             .Formulas(67) = "PUEOtherPurchAssVal5 ='" & Format(IIf(IsNull(Rs!PUEOtherPurchAssVal5), 0#, Rs!PUEOtherPurchAssVal5), FStr) & "'" '-- Purchase Other Ass Val 5 %
             .Formulas(68) = "PUEOtherPurchAssVal125 ='" & Format(IIf(IsNull(Rs!PUEOtherPurchAssVal125), 0#, Rs!PUEOtherPurchAssVal125), FStr) & "'" '-- Purchase Other Ass Val 125 %
            
             .Formulas(69) = "PUEOtherPurchTaxPaid4 ='" & Format(IIf(IsNull(Rs!PUEOtherPurchTaxPaid4), 0#, Rs!PUEOtherPurchTaxPaid4), FStr) & "'" '-- Purchase Other Tax Paid 4 %
             .Formulas(70) = "PUEOtherPurchTaxPaid5 ='" & Format(IIf(IsNull(Rs!PUEOtherPurchTaxPaid5), 0#, Rs!PUEOtherPurchTaxPaid5), FStr) & "'" '-- Purchase Other Tax Paid 5 %
             .Formulas(71) = "PUEOtherPurchTaxPaid125 ='" & Format(IIf(IsNull(Rs!PUEOtherPurchTaxPaid125), 0#, Rs!PUEOtherPurchTaxPaid125), FStr) & "'" '-- Purchase Other Tax Paid 125 %
            '.Action = 1
        End If
    .Formulas(20) = "SetOffOpBal ='" & Format(OpBal, FStr) & "'" '-- Set Off OP Bal
    .Formulas(21) = "wFrToDt = 'From : ' & '" & FromDt & "' & '  To ' & '" & ToDt & "'"
    .Formulas(22) = "wCSTNo = '" & gCCST & "'"
    .Formulas(23) = "wBSTNo = '" & gCBST & "'"
    .Formulas(24) = "wFileNo = '" & gCFileNo & "'"
    .Formulas(25) = "wAcSTONo = '" & gCACSTONo & "'"
    .ReportFileName = gReportPath & "rptTaxVATComput45.rpt"
    .SelectionFormula = "{tmpgenTbl.Username}='" & gUserName & "'"
    .Action = 1
    '=================================== Exempted Goods
    gCn.Execute "delete from tmpGenTbl where userName='" & gUserName & "'"
    If gBackEndDB = gBackEndAccess Then  '--Access
        gCn.Execute "insert into tmpGenTbl (Amt1,Amt2,amt3,amt4,Nar1,Nar2,Nar3,Nar4,UserName) " _
        & " select sum(AdBillAmt) as SAAmt,sum(AdTaxAmt) as TaxAmt,sum(ADEXEMPTAMT) as ExmptAmt, " _
        & "ACEXEMPTRATE,'SAAmt','TAXAmt','ExmptAmt','Exmpt%','" & gUserName & "' from tblAddLess,tblMastAccount where ADMILLCODE=AcCode and  " & CondStr & " and " _
        & " Vtype in ('SY','ST','SO','SD') and ADEXEMPTAMT>0  group by ACEXEMPTRATE"
    Else
        gCn.Execute "insert into tmpGenTbl (Amt1,Amt2,amt3,amt4,Nar1,Nar2,Nar3,Nar4,UserName) " _
        & " select nvl(sum(AdBillAmt),0) as SAAmt,nvl(sum(AdTaxAmt),0) as TaxAmt,nvl(sum(ADEXEMPTAMT),0) as ExmptAmt, " _
        & "ACEXEMPTRATE,'SAAmt','TAXAmt','ExmptAmt','Exmpt%','" & gUserName & "' from tblAddLess,tblMastAccount where ADMILLCODE=AcCode and  " & CondStr & " and " _
        & " Vtype in ('SY','ST','SO','SD') and ADEXEMPTAMT>0  group by ACEXEMPTRATE"
    End If
    For i = 4 To 35
        .Formulas(i) = ""
    Next i
    .Formulas(5) = "wFrToDt = 'From : ' & '" & FromDt & "' & '  To ' & '" & ToDt & "'"
    .ReportFileName = gReportPath & "rptTaxVATComputExmpt.rpt"
    .SelectionFormula = "{tmpgenTbl.Username}='" & gUserName & "'"
    .Formulas(6) = "wCSTNo = '" & gCCST & "'"
    .Formulas(7) = "wBSTNo = '" & gCBST & "'"
    .Formulas(8) = "wFileNo = '" & gCFileNo & "'"
    .Formulas(9) = "wAcSTONo = '" & gCACSTONo & "'"
 '   .Action = 1
End With
End Sub
Public Sub GProcPrepareVATComputation25(FromDt As Date, ToDt As Date, CrystalRpt As CrystalReport, OpBal As Double)
Dim CondStr As String
Dim X As String
Dim Rs As Recordset
Dim i As Long
If gBackEndDB = gBackEndAccess Then
    CondStr = " Vdt>= cdate('" & FromDt & "')  and  Vdt<=cdate('" & ToDt & "')"
Else '-- Oracle
    CondStr = " Vdt>= to_date('" & CDate(FromDt) & "','dd/MM/yyyy')  and Vdt<=to_date('" & CDate(ToDt) & "','dd/MM/yyyy')"
End If
gCn.Execute "delete from tmpGentbl where username='" & gUserName & "'"
With CrystalRpt
    '=======================Sale
    If gBackEndDB = gBackEndAccess Then  '--Access
'---For Tax 2 & 5 %
         X = "select sum(AdBillAmt) as SAAmt,sum(ADTAXAMT) as SATaxAmt, " _
         & " sum(iif(Narration='V A T',iif(VType='SY',iif(ADEXEMPTAMT=0,iif(AdTaxp=2,ADTAXAMT,0),0),0),0)) as SATradeTAX2 ,sum(iif(Narration='V A T',iif(VType='SY',iif(ADEXEMPTAMT=0,iif(AdTaxp=5,ADTAXAMT,0),0),0),0)) as SATradeTAX5 , " _
         & " sum(iif(Narration='V A T',iif(VType='SY',iif(ADEXEMPTAMT=0,iif(AdTaxp=2,AdAssVal,0),0),0),0)) as SATradeAssVal2 ,sum(iif(Narration='V A T',iif(VType='SY',iif(ADEXEMPTAMT=0,iif(AdTaxp=5,AdAssVal,0),0),0),0)) as SATradeAssVal5 , " _
         & " sum(iif(Narration='V A T',iif(VType='ST',iif(ADEXEMPTAMT=0,iif(AdTaxp=2,ADTAXAMT,0),0),0),0)) as SASITTAX2 ,sum(iif(Narration='V A T',iif(VType='ST',iif(ADEXEMPTAMT=0,iif(AdTaxp=5,ADTAXAMT,0),0),0),0)) as SASITTAX5 , " _
         & " sum(iif(Narration='V A T',iif(VType='ST',iif(ADEXEMPTAMT=0,iif(AdTaxp=2,AdAssVal,0),0),0),0)) as SASITAssVal2 ,sum(iif(Narration='V A T',iif(VType='ST',iif(ADEXEMPTAMT=0,iif(AdTaxp=5,AdAssVal,0),0),0),0)) as SASITAssVal5 , " _
         & " sum(iif(Narration='V A T',iif(VType='SO',iif(AdTaxp=2,ADTAXAMT,0),0),0)) as SAConsigTAX2 , " _
         & " sum(iif(Narration='V A T',iif(VType='SO',iif(AdTaxp=5,ADTAXAMT,0),0),0)) as SAConsigTAX5 , " _
         & " sum(iif(Narration='V A T',iif(VType='SO',iif(AdTaxp=2,AdAssVal,0),0),0)) as SAConsigAssVal2 , " _
         & " sum(iif(Narration='V A T',iif(VType='SO',iif(AdTaxp=5,AdAssVal,0),0),0)) as SAConsigAssVal5 , " _
         & " sum(iif(Narration='V A T',iif(VType='SD',iif(AdTaxp=2,ADTAXAMT,0),0),0)) as SADepotTAX2 , " _
         & " sum(iif(Narration='V A T',iif(VType='SD',iif(AdTaxp=5,ADTAXAMT,0),0),0)) as SADepotTAX5 ," _
         & " sum(iif(Narration='V A T',iif(VType='SD',iif(AdTaxp=2,AdAssVal,0),0),0)) as SADepotAssVal2 , " _
         & " sum(iif(Narration='V A T',iif(VType='SD',iif(AdTaxp=5,AdAssVal,0),0),0)) as SADepotAssVal5 , " _
         & " sum(iif(Narration='V A T',iif(VType='SY',iif(ADEXEMPTAMT>0,iif(AdTaxp=2,ADTAXAMT,0),0),0),0)) as SAExmptTAX2 , " _
         & " sum(iif(Narration='V A T',iif(VType='SY',iif(ADEXEMPTAMT>0,iif(AdTaxp=5,ADTAXAMT,0),0),0),0)) as SAExmptTAX5 , " _
         & " sum(iif(Narration='V A T',iif(VType='SY',iif(ADEXEMPTAMT>0,iif(AdTaxp=2,ADEXEMPTAMT,0),0),0),0)) as SAExmpt2, " _
         & " sum(iif(Narration='V A T',iif(VType='SY',iif(ADEXEMPTAMT>0,iif(AdTaxp=5,ADEXEMPTAMT,0),0),0),0)) as SAExmpt5, " _
         & " sum(iif(Narration='V A T',iif(VType='SY',iif(ADEXEMPTAMT>0,iif(AdTaxp=2,AdAssVal,0),0),0),0)) as SAExmptAssVal2 , " _
         & " sum(iif(Narration='V A T',iif(VType='SY',iif(ADEXEMPTAMT>0,iif(AdTaxp=5,AdAssVal,0),0),0),0)) as SAExmptAssVal5 , " _
         & " sum(iif(VType='ST',iif(Narration='C S T',0,AdBillAmt),0)) as SASIT , " _
         & " sum(iif(Narration='Tax Free',AdBillAmt,0)) as SATAXFREE, " _
         & " sum(iif(Narration='C S T',AdBillAmt,0)) as SACST, " _
         & " sum(iif(Narration='C S T',ADTAXAMT,0)) as SaCSTTax, sum(IIf(VType='SY' And ADEXEMPTAMT>0,AdBillAmt,0)) as SAExemptSale , " _
         & " sum(iif(Narration='V A T',iif(VType='SY',iif(ADEXEMPTAMT=0,AdBillAmt,0),0),0)) as SATradeSale ,sum(iif(Narration='V A T',iif(VType='SO',iif(ADEXEMPTAMT=0,AdBillAmt,0),0),0)) as SAConsiSale ,sum(iif(Narration='V A T',iif(VType='SD',iif(ADEXEMPTAMT=0,AdBillAmt,0),0),0)) as SADepotSale  " _
         & " from tbladdless,tblMastNarration where ADTAXCODE=Narrcode and " & CondStr & " and vtype in ('SY','ST','SO','SD')"
'-----
    End If
        If GProcRstOpen(Rs, X, "R") > 0 Then
             .Formulas(4) = "SAAmt ='" & Format(IIf(IsNull(Rs!SAAmt), 0#, Rs!SAAmt), FStr) & "'" '-- GTO Sale
             .Formulas(5) = "SATaxAmt ='" & Format(IIf(IsNull(Rs!SATaxAmt), 0#, Rs!SATaxAmt), FStr) & "'" '-- SA Tax Amt
             .Formulas(10) = "SASIT ='" & Format(IIf(IsNull(Rs!SASIT), 0#, Rs!SASIT), FStr) & "'" '-- SA SIT
             .Formulas(11) = "SACST ='" & Format(IIf(IsNull(Rs!SACST), 0#, Rs!SACST), FStr) & "'" '--SA CST
             .Formulas(28) = "SACstTax ='" & Format(IIf(IsNull(Rs!SACstTax), 0#, Rs!SACstTax), FStr) & "'" '--SA Cst Tax
             .Formulas(29) = "SATaxFree ='" & Format(IIf(IsNull(Rs!SATaxFree), 0#, Rs!SATaxFree), FStr) & "'" '--SA Tax Free
            
             .Formulas(30) = "SATradeTAX2 ='" & Format(IIf(IsNull(Rs!SATradeTAX2), 0#, Rs!SATradeTAX2), FStr) & "'" '-- Trade 2 % Tax
             .Formulas(31) = "SATradeTAX5 ='" & Format(IIf(IsNull(Rs!SATradeTAX5), 0#, Rs!SATradeTAX5), FStr) & "'" '-- Trade 5 % Tax
             .Formulas(32) = "SATradeAssVal2 ='" & Format(IIf(IsNull(Rs!SATradeAssVal2), 0#, Rs!SATradeAssVal2), FStr) & "'" '-- Trade 2 % Ass Val
             .Formulas(33) = "SATradeAssVal5 ='" & Format(IIf(IsNull(Rs!SATradeAssVal5), 0#, Rs!SATradeAssVal5), FStr) & "'" '-- Trade 5 % Ass Val
             .Formulas(34) = "SAConsigTAX2 ='" & Format(IIf(IsNull(Rs!SAConsigTAX2), 0#, Rs!SAConsigTAX2), FStr) & "'" '-- Consi 2 % Tax
             .Formulas(35) = "SAConsigTAX5 ='" & Format(IIf(IsNull(Rs!SAConsigTAX5), 0#, Rs!SAConsigTAX5), FStr) & "'" '-- Consi 5 % Tax
             .Formulas(36) = "SAConsigAssVal2 ='" & Format(IIf(IsNull(Rs!SAConsigAssVal2), 0#, Rs!SAConsigAssVal2), FStr) & "'" '-- Consi 2 % Ass Val
             .Formulas(37) = "SAConsigAssVal5 ='" & Format(IIf(IsNull(Rs!SAConsigAssVal5), 0#, Rs!SAConsigAssVal5), FStr) & "'" '-- Consi 5 % Ass Val
             .Formulas(38) = "SADepotTAX2 ='" & Format(IIf(IsNull(Rs!SADepotTAX2), 0#, Rs!SADepotTAX2), FStr) & "'" '-- Depot 2 % Tax
             .Formulas(39) = "SADepotTAX5 ='" & Format(IIf(IsNull(Rs!SADepotTAX5), 0#, Rs!SADepotTAX5), FStr) & "'" '-- Depot 5 % Tax
             .Formulas(40) = "SADepotAssVal2 ='" & Format(IIf(IsNull(Rs!SADepotAssVal2), 0#, Rs!SADepotAssVal2), FStr) & "'" '-- Depot 2 % Ass Val
             .Formulas(41) = "SADepotAssVal5 ='" & Format(IIf(IsNull(Rs!SADepotAssVal5), 0#, Rs!SADepotAssVal5), FStr) & "'" '-- Depot 5 % Ass Val
             .Formulas(42) = "SAExmptTAX2 ='" & Format(IIf(IsNull(Rs!SAExmptTAX2), 0#, Rs!SAExmptTAX2), FStr) & "'" '-- Exempt 2 % Tax
             .Formulas(43) = "SAExmptTAX5 ='" & Format(IIf(IsNull(Rs!SAExmptTAX5), 0#, Rs!SAExmptTAX5), FStr) & "'" '-- Exempt 5 % Tax
             .Formulas(44) = "SAExmptAssVal2 ='" & Format(IIf(IsNull(Rs!SAExmptAssVal2), 0#, Rs!SAExmptAssVal2), FStr) & "'" '-- Exempt 2 % Ass Val
             .Formulas(45) = "SAExmptAssVal5 ='" & Format(IIf(IsNull(Rs!SAExmptAssVal5), 0#, Rs!SAExmptAssVal5), FStr) & "'" '-- Exempt 5 % Ass Val
             .Formulas(46) = "SAExmpt2 ='" & Format(IIf(IsNull(Rs!SAExmpt2), 0#, Rs!SAExmpt2), FStr) & "'" '-- Exempt 2 % Purchase Prise
             .Formulas(47) = "SAExmpt5 ='" & Format(IIf(IsNull(Rs!SAExmpt5), 0#, Rs!SAExmpt5), FStr) & "'" '-- Exempt 5 % Purchase Prise
            
             .Formulas(48) = "SATradeSale ='" & Format(IIf(IsNull(Rs!SATradeSale), 0#, Rs!SATradeSale), FStr) & "'" '-- Total Trade Sale
             .Formulas(49) = "SAConsiSale ='" & Format(IIf(IsNull(Rs!SAConsiSale), 0#, Rs!SAConsiSale), FStr) & "'" '-- Total Consign Sale
             .Formulas(50) = "SADepotSale ='" & Format(IIf(IsNull(Rs!SADepotSale), 0#, Rs!SADepotSale), FStr) & "'" '-- Total Depot Sale
             .Formulas(52) = "SAExemptSale ='" & Format(IIf(IsNull(Rs!SAExemptSale), 0#, Rs!SAExemptSale), FStr) & "'" '-- Total Exempt Sale
             .Formulas(72) = "SASitTax2 ='" & Format(IIf(IsNull(Rs!SASITTAX2), 0#, Rs!SASITTAX2), FStr) & "'" '-- SIT Tax 2
             .Formulas(73) = "SASitTax5 ='" & Format(IIf(IsNull(Rs!SASITTAX5), 0#, Rs!SASITTAX5), FStr) & "'" '-- SIT Tax 5
             .Formulas(74) = "SASitAssVal2 ='" & Format(IIf(IsNull(Rs!SASITAssVal2), 0#, Rs!SASITAssVal2), FStr) & "'" '-- SIT AssVal 2
             .Formulas(75) = "SASitAssVal5 ='" & Format(IIf(IsNull(Rs!SASITAssVal5), 0#, Rs!SASITAssVal5), FStr) & "'" '-- SIT AssVal 5
            '.Action = 1
        End If
    '=======================Purchase
    If gBackEndDB = gBackEndAccess Then  '--Access
        X = "select sum(AdBillAmt) as PUAmt,sum(iif(Vtype='PI',AdBillAmt,0)) as PUPIAmt, " _
        & " sum(iif(VType='PY',ADEXEMPTAMT,0)) as PUPYExmpt , " _
        & " sum(iif(Narration<>'V A T' and Narration <>'Tax Free',iif(VType='PY',AdBillAmt,0),0)) as PUOms ," _
        & " sum(iif(Narration='Tax Free',iif(VType='PY',AdBillAmt,0),0)) as PUTaxFree ," _
        & " sum(iif(VType='PT',AdBillAmt,0)) as PUOMSSit ," _
        & " sum(iif(Narration='V A T',iif(VType='PY',AdTaxAmt,0),0)) as PUTaxPaid ," _
        & " sum(iif(Narration='V A T',iif(VType='PY',iif(ADEXEMPTAMT=0,iif(AdTaxp=2,AdTaxAmt,0),0),0),0)) as PUTradePurchTaxPaid2 ," _
        & " sum(iif(Narration='V A T',iif(VType='PY',iif(ADEXEMPTAMT=0,iif(AdTaxp=5,AdTaxAmt,0),0),0),0)) as PUTradePurchTaxPaid5 ," _
        & " sum(iif(Narration='V A T',iif(VType='PY',iif(ADEXEMPTAMT=0,iif(AdTaxp=2,AdAssVal,0),0),0),0)) as PUTradePurchAssVal2 ," _
        & " sum(iif(Narration='V A T',iif(VType='PY',iif(ADEXEMPTAMT=0,iif(AdTaxp=5,AdAssVal,0),0),0),0)) as PUTradePurchAssVal5 ," _
        & " sum(iif(Narration='V A T',iif(VType='PY',iif(ADEXEMPTAMT=0,AdBillAmt,0),0),0)) as PUTradePurch , " _
        & " sum(iif(Narration='V A T',iif(VType='PY',iif(ADEXEMPTAMT>0,iif(AdTaxp=2,AdTaxAmt,0),0),0),0)) as PUExemptPurchTaxPaid2 ," _
        & " sum(iif(Narration='V A T',iif(VType='PY',iif(ADEXEMPTAMT>0,iif(AdTaxp=5,AdTaxAmt,0),0),0),0)) as PUExemptPurchTaxPaid5 ," _
        & " sum(iif(Narration='V A T',iif(VType='PY',iif(ADEXEMPTAMT>0,iif(AdTaxp=2,AdAssVal,0),0),0),0)) as PUExemptPurchAssVal2 ," _
        & " sum(iif(Narration='V A T',iif(VType='PY',iif(ADEXEMPTAMT>0,iif(AdTaxp=5,AdAssVal,0),0),0),0)) as PUExemptPurchAssVal5 ," _
        & " sum(iif(Narration='V A T',iif(VType='PY',iif(ADEXEMPTAMT>0,AdBillAmt,0),0),0)) as PUExemptPurch , " _
        & " sum(iif(Narration='V A T',iif(VType='PY',iif(ADEXEMPTAMT>0,iif(AdTaxp=2,ADEXEMPTAMT,0),0),0),0)) as PUExemptAmt2 ," _
        & " sum(iif(Narration='V A T',iif(VType='PY',iif(ADEXEMPTAMT>0,iif(AdTaxp=5,ADEXEMPTAMT,0),0),0),0)) as PUExemptAmt5 ," _
        & " sum(ADTaxAMT) as PUTax,sum(ADAFTTAXADD) as PULess,sum(ADAFTTAXLESS) as PUAdd, " _
        & " sum(iif(Narration='V A T',iif(VType='PO',AdBillAmt,0),0)) as PUTOtherPurch , " _
        & " sum(iif(Narration='V A T',iif(VType='PO',iif(AdTaxp=2,AdAssVal,0),0),0)) as PUEOtherPurchAssVal2 ," _
        & " sum(iif(Narration='V A T',iif(VType='PO',iif(AdTaxp=5,AdAssVal,0),0),0)) as PUEOtherPurchAssVal5 ," _
        & " sum(iif(Narration='V A T',iif(VType='PO',iif(AdTaxp=12.5,AdAssVal,0),0),0)) as PUEOtherPurchAssVal125, sum(iif(Narration='V A T',iif(VType='PO',iif(AdTaxp=2,AdTaxAmt,0),0),0)) as PUEOtherPurchTaxPaid2 , sum(iif(Narration='V A T',iif(VType='PO',iif(AdTaxp=5,AdTaxAmt,0),0),0)) as PUEOtherPurchTaxPaid5 ," _
        & " sum(iif(Narration='V A T',iif(VType='PO',iif(AdTaxp=12.5,AdTaxAmt,0),0),0)) as PUEOtherPurchTaxPaid125 " _
        & " from tbladdless,tblMastNarration,tblMastAccount where  tblAddless.AdMillCode = tblMastAccount.AcCode And tblMastAccount.AcMillType <> 2 and tblAddless.ADTAXCODE=tblMastNarration.Narrcode and " & CondStr & " and tblAddless.vtype in ('PY','PT','PI','PO')"
    End If
        If GProcRstOpen(Rs, X, "R") > 0 Then '
             .Formulas(13) = "PUAmt ='" & Format(IIf(IsNull(Rs!PUAmt), 0#, Rs!PUAmt), FStr) & "'" '-- PUAmt (GTO)
             .Formulas(14) = "PUPIAmt ='" & Format(IIf(IsNull(Rs!PUPIAmt), 0#, Rs!PUPIAmt), FStr) & "'" '--PUPIAmt  (Stock Trans OMS)
             .Formulas(15) = "PUPYExmpt ='" & Format(IIf(IsNull(Rs!PUPYExmpt), 0#, Rs!PUPYExmpt), FStr) & "'" '-- PUPY Exmpt ( R.D. Exmpt)
             .Formulas(16) = "PUOms ='" & Format(IIf(IsNull(Rs!PUOms), 0#, Rs!PUOms), FStr) & "'" '-- PUOms ( OMS )
             .Formulas(17) = "PUTaxFree ='" & Format(IIf(IsNull(Rs!PUTaxFree), 0#, Rs!PUTaxFree), FStr) & "'" '-- PUTaxFree (Tax Free)
             .Formulas(18) = "PUOMSSit ='" & Format(IIf(IsNull(Rs!PUOMSSit), 0#, Rs!PUOMSSit), FStr) & "'" '-- PUOMSSit (OMS SIT)
             .Formulas(19) = "PUTaxPaid ='" & Format(IIf(IsNull(Rs!PUTaxPaid), 0#, Rs!PUTaxPaid), FStr) & "'" '-- PUTaxPaid (Tax Paid)
             .Formulas(26) = "PUOtherLess ='" & Format(IIf(IsNull(Rs!PULess), 0#, Rs!PULess), FStr) & "'" '-- PUOtherLess (Add Aft Tax)
             .Formulas(27) = "PUOtherAdd ='" & Format(IIf(IsNull(Rs!PUAdd), 0#, Rs!PUAdd), FStr) & "'" '-- PUOtherAdd (Less Aft Tax)
             
             .Formulas(53) = "PUTradePurchTaxPaid2 ='" & Format(IIf(IsNull(Rs!PUTradePurchTaxPaid2), 0#, Rs!PUTradePurchTaxPaid2), FStr) & "'" '-- Purchase tax paid 2 %
             .Formulas(54) = "PUTradePurchTaxPaid5 ='" & Format(IIf(IsNull(Rs!PUTradePurchTaxPaid5), 0#, Rs!PUTradePurchTaxPaid5), FStr) & "'" '-- Purchase tax paid 5 %
             .Formulas(55) = "PUTradePurchAssVal2 ='" & Format(IIf(IsNull(Rs!PUTradePurchAssVal2), 0#, Rs!PUTradePurchAssVal2), FStr) & "'" '-- Purchase Ass Val 2 %
             .Formulas(56) = "PUTradePurchAssVal5 ='" & Format(IIf(IsNull(Rs!PUTradePurchAssVal5), 0#, Rs!PUTradePurchAssVal5), FStr) & "'" '-- Purchase Ass Val 5 %
             .Formulas(57) = "PUTradePurch ='" & Format(IIf(IsNull(Rs!PUTradePurch), 0#, Rs!PUTradePurch), FStr) & "'" '-- Purchase Trade Amt
            
             .Formulas(58) = "PUExemptPurchTaxPaid2 ='" & Format(IIf(IsNull(Rs!PUExemptPurchTaxPaid2), 0#, Rs!PUExemptPurchTaxPaid2), FStr) & "'" '-- Purchase Exempt Tax Paid 2 %
             .Formulas(59) = "PUExemptPurchTaxPaid5 ='" & Format(IIf(IsNull(Rs!PUExemptPurchTaxPaid5), 0#, Rs!PUExemptPurchTaxPaid5), FStr) & "'" '-- Purchase Exempt Tax Paid 5 %
             .Formulas(60) = "PUExemptPurchAssVal2 ='" & Format(IIf(IsNull(Rs!PUExemptPurchAssVal2), 0#, Rs!PUExemptPurchAssVal2), FStr) & "'" '-- Purchase Exempt Ass Val 2 %
             .Formulas(61) = "PUExemptPurchAssVal5 ='" & Format(IIf(IsNull(Rs!PUExemptPurchAssVal5), 0#, Rs!PUExemptPurchAssVal5), FStr) & "'" '-- Purchase Exempt Ass Val 5 %
             .Formulas(62) = "PUExemptPurch ='" & Format(IIf(IsNull(Rs!PUExemptPurch), 0#, Rs!PUExemptPurch), FStr) & "'" '--  Exempt Purchase Amt
             .Formulas(63) = "PUExemptAmt2 ='" & Format(IIf(IsNull(Rs!PUExemptAmt2), 0#, Rs!PUExemptAmt2), FStr) & "'" '-- Purchase Exempt Amt 2 %
             .Formulas(64) = "PUExemptAmt5 ='" & Format(IIf(IsNull(Rs!PUExemptAmt5), 0#, Rs!PUExemptAmt5), FStr) & "'" '-- Purchase Exempt Amt 5 %
            
             
             .Formulas(65) = "PUTOtherPurch ='" & Format(IIf(IsNull(Rs!PUTOtherPurch), 0#, Rs!PUTOtherPurch), FStr) & "'" '-- Purchase Other Val
             .Formulas(66) = "PUEOtherPurchAssVal2 ='" & Format(IIf(IsNull(Rs!PUEOtherPurchAssVal2), 0#, Rs!PUEOtherPurchAssVal2), FStr) & "'" '-- Purchase Other Ass Val 2 %
             .Formulas(67) = "PUEOtherPurchAssVal5 ='" & Format(IIf(IsNull(Rs!PUEOtherPurchAssVal5), 0#, Rs!PUEOtherPurchAssVal5), FStr) & "'" '-- Purchase Other Ass Val 5 %
             .Formulas(68) = "PUEOtherPurchAssVal125 ='" & Format(IIf(IsNull(Rs!PUEOtherPurchAssVal125), 0#, Rs!PUEOtherPurchAssVal125), FStr) & "'" '-- Purchase Other Ass Val 125 %
            
             .Formulas(69) = "PUEOtherPurchTaxPaid2 ='" & Format(IIf(IsNull(Rs!PUEOtherPurchTaxPaid2), 0#, Rs!PUEOtherPurchTaxPaid2), FStr) & "'" '-- Purchase Other Tax Paid 2 %
             .Formulas(70) = "PUEOtherPurchTaxPaid5 ='" & Format(IIf(IsNull(Rs!PUEOtherPurchTaxPaid5), 0#, Rs!PUEOtherPurchTaxPaid5), FStr) & "'" '-- Purchase Other Tax Paid 5 %
             .Formulas(71) = "PUEOtherPurchTaxPaid125 ='" & Format(IIf(IsNull(Rs!PUEOtherPurchTaxPaid125), 0#, Rs!PUEOtherPurchTaxPaid125), FStr) & "'" '-- Purchase Other Tax Paid 125 %
            '.Action = 1
        End If
    .Formulas(20) = "SetOffOpBal ='" & Format(OpBal, FStr) & "'" '-- Set Off OP Bal
    .Formulas(21) = "wFrToDt = 'From : ' & '" & FromDt & "' & '  To ' & '" & ToDt & "'"
    .Formulas(22) = "wCSTNo = '" & gCCST & "'"
    .Formulas(23) = "wBSTNo = '" & gCBST & "'"
    .Formulas(24) = "wFileNo = '" & gCFileNo & "'"
    .Formulas(25) = "wAcSTONo = '" & gCACSTONo & "'"
    .ReportFileName = gReportPath & "rptTaxVATComput25.rpt"
    .SelectionFormula = "{tmpgenTbl.Username}='" & gUserName & "'"
    .Action = 1
    '=================================== Exempted Goods
    gCn.Execute "delete from tmpGenTbl where userName='" & gUserName & "'"
    If gBackEndDB = gBackEndAccess Then  '--Access
        gCn.Execute "insert into tmpGenTbl (Amt1,Amt2,amt3,amt4,Nar1,Nar2,Nar3,Nar4,UserName) " _
        & " select sum(AdBillAmt) as SAAmt,sum(AdTaxAmt) as TaxAmt,sum(ADEXEMPTAMT) as ExmptAmt, " _
        & "ACEXEMPTRATE,'SAAmt','TAXAmt','ExmptAmt','Exmpt%','" & gUserName & "' from tblAddLess,tblMastAccount where ADMILLCODE=AcCode and  " & CondStr & " and " _
        & " Vtype in ('SY','ST','SO','SD') and ADEXEMPTAMT>0  group by ACEXEMPTRATE"
    Else
        gCn.Execute "insert into tmpGenTbl (Amt1,Amt2,amt3,amt4,Nar1,Nar2,Nar3,Nar4,UserName) " _
        & " select nvl(sum(AdBillAmt),0) as SAAmt,nvl(sum(AdTaxAmt),0) as TaxAmt,nvl(sum(ADEXEMPTAMT),0) as ExmptAmt, " _
        & "ACEXEMPTRATE,'SAAmt','TAXAmt','ExmptAmt','Exmpt%','" & gUserName & "' from tblAddLess,tblMastAccount where ADMILLCODE=AcCode and  " & CondStr & " and " _
        & " Vtype in ('SY','ST','SO','SD') and ADEXEMPTAMT>0  group by ACEXEMPTRATE"
    End If
    For i = 4 To 35
        .Formulas(i) = ""
    Next i
    .Formulas(5) = "wFrToDt = 'From : ' & '" & FromDt & "' & '  To ' & '" & ToDt & "'"
    .ReportFileName = gReportPath & "rptTaxVATComputExmpt.rpt"
    .SelectionFormula = "{tmpgenTbl.Username}='" & gUserName & "'"
    .Formulas(6) = "wCSTNo = '" & gCCST & "'"
    .Formulas(7) = "wBSTNo = '" & gCBST & "'"
    .Formulas(8) = "wFileNo = '" & gCFileNo & "'"
    .Formulas(9) = "wAcSTONo = '" & gCACSTONo & "'"
 '   .Action = 1
End With

End Sub



'============= Common Outstanding
Public Function GProcCollectCommonOutstanding(w_Fr As Date, w_To As Date, Optional w_MillName As String, Optional w_BrkName As String, Optional w_PartyName As String, Optional CurrentFirmFlag As Boolean, Optional ShowUnAdjustReceipt As Boolean, Optional IsGSTOs As Boolean)
Dim mTempCn As Connection, mCurrFirmDbNm As String, mCurrFirmNm As String
Dim RstNewOutst As Recordset, RstOldOutst As Recordset, RstCompMast As Recordset, RstTmpAcMast As Recordset, RstLstRecDt As Recordset
Dim X1 As String, i As Integer, j As Long, FirmNm As String
Dim X As String
Dim XX As String
Dim RsP As Recordset, RsTmp As Recordset
'----- if ShowUnAdjustReceipt then insert UnAdjusted Receipt InTmpGenTbl
If ShowUnAdjustReceipt = True Then
    If gBackEndDB = gBackEndAccess Then
        X1 = " Vdt>=cdate('" & w_Fr & "')  and Vdt<=Cdate('" & w_To & "')"
    Else
        X1 = " Vdt>=to_Date('" & w_Fr & "','dd/MM/yyyy')  and Vdt<=to_date('" & w_To & "','dd/MM/yyyy')"
    End If
    If w_PartyName <> "" Then X1 = X1 & " and Party='" & w_PartyName & "'"
    X = "select Vno,Vdt,Vtype,Party,CrAccount,Vamt,UNADJUSTAMT,VNAR1,VNAR2,VNAR3,'" & gUserName & "' from QryRegUnadjustReceipt where UnadjustAmt>0 and " & X1
    '--- insert to setup db Table For All Firm
    GProcInsertToSetupDbTbl Not CurrentFirmFlag, X, "tmpGenTbl", "Vno,Vdt,Vtype,Nar1,Nar2,Amt1,Amt2,Nar3,Nar4,Nar5,username", "username='" & gUserName & "'", "Nar6"
End If
'---- Insert Outstanding Records In TmpTblOutstanding
gSetupCn.BeginTrans
gSetupCn.Execute "Delete from tmptbloutstanding where username = '" & gUserName & "'"
GProcRstOpen RstNewOutst, "Select * from tmptbloutstanding ", "O", gSetupCn
GProcRstOpen RstCompMast, "Select * from tblmastcompany " & IIf(CurrentFirmFlag = True, " where cname='" & gCName & "'", ""), "R", gSetupCn
'----New Temparary Connection
mCurrFirmDbNm = gDbNm
mCurrFirmNm = gCName
RstCompMast.MoveFirst
Do While Not RstCompMast.EOF
   gDbNm = RstCompMast!CCode + Right(mCurrFirmDbNm, 4)
   FirmNm = RstCompMast!Cname
   If gBackEndDB = gBackEndORA Then
        GProcConnectionOpenORA mTempCn
        X1 = "where vdt <= to_date('" & Format(w_To, "dd/MM/yyyy") & "','dd/MM/yyyy')"
        If w_Fr <> gCYSDate Then X1 = X1 + " and vdt >= to_date('" & Format(w_Fr, "dd/MM/yyyy") & "','dd/MM/yyyy')"
   Else
        GProcConnectionOpen mTempCn
        X1 = "where format(vdt,'yyyy/MM/dd') <= '" & Format(w_To, "yyyy/MM/dd") & "'"
        If w_Fr <> gCYSDate Then X1 = X1 + " and format(vdt,'yyyy/MM/dd') >= '" & Format(w_Fr, "dd/MM/yyyy") & "'"
   End If
   
'-------For Updating o/s Balance for ToDate
       '================Outstanding Vs Receipt
        mTempCn.BeginTrans
        mTempCn.Execute "Update tblOutStanding set OutRecAMT = 0, OutRecDbAMT = 0"
        If (GProcRstOpen(RsP, "select * from tblRecVsSale where format(RecVDt,'yyyy/MM/dd') <= '" & Format(w_To, "yyyy/MM/dd") & "'", "O", mTempCn) > 0) Then
            
            RsP.MoveFirst
            Do While Not RsP.EOF
                If RsP!InvTp = "SM" Then
                    mTempCn.Execute "Update tblOutStanding set OutRecAMT = OutRecAmt+" & RsP!AdjAmt & "+" & RsP!AdjDbNtAmt & " where vtype='" & RsP!InvTp & "' and vno=" & RsP!InvNo & " and VYear=" & RsP!InvYear
                Else
                    mTempCn.Execute "Update tblOutStanding set OutRecAMT = OutRecAmt+" & RsP!AdjAmt & "+" & RsP!Discount & " , OutRecDbAMT = OutRecDbAmt+" & RsP!AdjDbNtAmt & " where vtype='" & RsP!InvTp & "' and vno=" & RsP!InvNo & " and VYear=" & RsP!InvYear
                End If
                RsP.MoveNext
                Loop
            End If
            mTempCn.CommitTrans
            RsP.Close
'-------- End
   
   
   If w_MillName <> "" Then X1 = X1 + " and OutMillCode = " & GProcGetColumnValue("tblmastaccount", "Acname", w_MillName, "S", "accode", "N", , mTempCn)
   If w_BrkName <> "" Then X1 = X1 + " and OutBrokerCode = " & GProcGetColumnValue("tblmastaccount", "Acname", w_BrkName, "S", "accode", "N", , mTempCn)
   If w_PartyName <> "" Then X1 = X1 + " and OutAcCode = " & GProcGetColumnValue("tblmastaccount", "Acname", w_PartyName, "S", "accode", "N", , mTempCn)
   
   If gBackEndDB = gBackEndAccess Then
       If IsGSTOs = True Then
             X = "Select * from tbloutstanding " & X1 & " and  vtype in ('SN') and (OutAmount - OutRecAmt > 0)"
       Else
            If gSelectedMenu = "Mill Bill Outstanding Broker with Partywise" Then '"Mill Bill Outstanding Broker with Partywise"
                 X = "Select * from tbloutstanding " & X1 & " and  vtype in ('SM') and (OutAmount - OutRecAmt > 0) " 'or (OutDbNtAmt - OutRecDbAmt > 0) )
            Else
                 X = "Select * from tbloutstanding " & X1 & " and  vtype in ('SD','SY','SM','ST','SO','SN') and (OutAmount - OutRecAmt > 0) " 'or (OutDbNtAmt - OutRecDbAmt > 0) )
            End If
      End If
   Else '-- Oracle
        If gSelectedMenu = "Mill Bill Outstanding Broker with Partywise" Then '"Mill Bill Outstanding Broker with Partywise"
            X = "Select * from tbloutstanding " & X1 & " and  vtype in ('SM') and ((OutAmount - OutRecAmt > 0) or (OutDbNtAmt - OutRecDbAmt > 0) )"
        Else  'vtype<> SM
            X = "Select * from tbloutstanding " & X1 & " and  vtype in ('SD','SY','ST','SO') and ((OutAmount - OutRecAmt > 0) or (OutDbNtAmt - OutRecDbAmt > 0) )"
        End If
   End If
   If (GProcRstOpen(RstOldOutst, X, "R", mTempCn) > 0) Then
       With RstOldOutst
            .MoveFirst
Loop1:
            Do While Not .EOF
               'If (!Vtype <> "SO") And (!Vtype <> "SY") And (!Vtype <> "ST") Then
               '   .MoveNext
               '   GoTo Loop1
               'End If
               'If (!OutAmount - !OutRecAmt > 0) Or (!OutDbNtAmt - !OutRecDbAmt > 0) Then
                  RstNewOutst.AddNew
                  For i = 0 To 20
                     If (i = 5) Or (i = 6) Or (i = 11) Then
                        RstNewOutst.Fields(i) = GProcGetColumnValue("tblmastaccount", "accode", .Fields(i), "N", "acname", "S", , mTempCn) & vbNullString
                        If i = 6 Then '- Broker Phone
                            RstNewOutst.Fields("OutBrkPhone") = GProcGetColumnValue("tblmastaccount", "acCode", .Fields(i), "N", "AcContNo", "S", , mTempCn) & vbNullString
                        End If
                        If i = 5 Then '- Party Phone
'                            RstNewOutst.Fields("OutPartyPhone") = GProcGetColumnValue("tblmastaccount", "acCode", .Fields(i), "N", "'('+'Ph : '+AcContNo+')'", "S", , mTempCn) & vbNullString
                            RstNewOutst.Fields("OutPartyPhone") = GProcGetColumnValue("tblmastaccount", "acCode", .Fields(i), "N", "AcContNo", "S", , mTempCn) & vbNullString
                            RstNewOutst.Fields("OutIsShowOuts") = GProcGetColumnValue("tblmastaccount", "acCode", .Fields(i), "N", "AcIsShowOuts", "N", , mTempCn) & vbNullString
                            RstNewOutst.Fields("OutPartyTin") = GProcGetColumnValue("tblmastaccount", "acCode", .Fields(i), "N", "AcCst", "S", , mTempCn) & vbNullString
                            'RstNewOutst.Fields("OutPartyTin") = GProcGetColumnValue("tblmastaccount", "acCode", .Fields(i), "N", "AcCst", "N", , mTempCn) & vbNullString
                        End If
                     Else
                        RstNewOutst.Fields(i) = .Fields(i)
                     End If
                  Next i
                  XX = "select ADTMP8 from tblAddLess where vno =  " & .Fields(0) & "  and vtype = '" & .Fields(1) & "'  and vyear = " & .Fields(2) & ""
                  If (GProcRstOpen(RstLstRecDt, XX, "R", mTempCn) > 0) Then
                     RstNewOutst!OutGpno = RstLstRecDt.Fields(0)
                  End If
                  RstNewOutst!UserName = gUserName
                  RstNewOutst!VCompSNm = Left(gDbNm, 4)
                  RstNewOutst!VCompNm = Left(FirmNm, 40)
                  RstNewOutst.Update
               'End If
               .MoveNext
            Loop
       End With
   End If
   mTempCn.Close
   RstCompMast.MoveNext
Loop
'RstOldOutst.Close
'RstNewOutst.Close
If RstNewOutst.State = adStateOpen Then RstNewOutst.Close
If RstCompMast.State = adStateOpen Then RstCompMast.Close
If RstOldOutst.State = adStateOpen Then RstOldOutst.Close
gSetupCn.CommitTrans

gDbNm = mCurrFirmDbNm
GProcConnectionOpen gCn

End Function
'============= Purchase Outstanding
Public Function GProcPurchaseOutstanding(w_Fr As Date, w_To As Date, Optional w_BrkName As String, Optional w_PartyName As String, Optional w_MillName As String, Optional CurrentFirmFlag As Boolean)
Dim mTempCn As Connection, mCurrFirmDbNm As String
Dim RstNewOutst As Recordset, RstOldOutst As Recordset, RstCompMast As Recordset, RstTmpAcMast As Recordset
Dim X1 As String, i As Integer, j As Long
Dim X As String
'---- Insert Outstanding Records In TmpTblOutstanding
gSetupCn.BeginTrans
gSetupCn.Execute "Delete from tmptbloutstanding where username = '" & gUserName & "'"
GProcRstOpen RstNewOutst, "Select * from tmptbloutstanding ", "O", gSetupCn
GProcRstOpen RstCompMast, "Select * from tblmastcompany " & IIf(CurrentFirmFlag = True, " where cname='" & gCName & "'", ""), "R", gSetupCn
'----New Temparary Connection
mCurrFirmDbNm = gDbNm
RstCompMast.MoveFirst
Do While Not RstCompMast.EOF
   gDbNm = RstCompMast!CCode + Right(mCurrFirmDbNm, 4)
   If gBackEndDB = gBackEndORA Then
        GProcConnectionOpenORA mTempCn
        X1 = "where vdt <= to_date('" & Format(w_To, "dd/MM/yyyy") & "','dd/MM/yyyy')"
        If w_Fr <> gCYSDate Then X1 = X1 + " and vdt >= to_date('" & Format(w_Fr, "dd/MM/yyyy") & "','dd/MM/yyyy')"
   Else
        GProcConnectionOpen mTempCn
        X1 = "where format(vdt,'yyyy/MM/dd') <= '" & Format(w_To, "yyyy/MM/dd") & "'"
        If w_Fr <> gCYSDate Then X1 = X1 + " and format(vdt,'yyyy/MM/dd') >= '" & Format(w_Fr, "dd/MM/yyyy") & "'"
   End If
   If w_MillName <> "" Then X1 = X1 + " and OutMillCode = " & GProcGetColumnValue("tblmastaccount", "Acname", w_MillName, "S", "accode", "N", , mTempCn)
   If w_BrkName <> "" Then X1 = X1 + " and OutBrokerCode = " & GProcGetColumnValue("tblmastaccount", "Acname", w_BrkName, "S", "accode", "N", , mTempCn)
   If w_PartyName <> "" Then X1 = X1 + " and OutAcCode = " & GProcGetColumnValue("tblmastaccount", "Acname", w_PartyName, "S", "accode", "N", , mTempCn)
   X = "Select * from tbloutstanding " & X1 & " and  vtype in ('PY','PT','PO','PI') and (OutAmount - OutRecAmt > 0) " 'or (OutDbNtAmt - OutRecDbAmt > 0) )
   If (GProcRstOpen(RstOldOutst, X, "R", mTempCn) > 0) Then
       With RstOldOutst
            .MoveFirst
Loop1:
            Do While Not .EOF
                RstNewOutst.AddNew
                For i = 0 To 20
                   If (i = 5) Or (i = 6) Or (i = 11) Then
                      RstNewOutst.Fields(i) = GProcGetColumnValue("tblmastaccount", "accode", .Fields(i), "N", "acname", "S", , mTempCn) & vbNullString
                   Else
                      RstNewOutst.Fields(i) = .Fields(i)
                   End If
                Next i
                RstNewOutst!UserName = gUserName
                RstNewOutst!VCompSNm = Left(gDbNm, 4)
                RstNewOutst.Update
               .MoveNext
            Loop
       End With
   End If
   mTempCn.Close
   RstCompMast.MoveNext
Loop
If RstNewOutst.State = adStateOpen Then RstNewOutst.Close
If RstCompMast.State = adStateOpen Then RstCompMast.Close
If RstOldOutst.State = adStateOpen Then RstOldOutst.Close
gSetupCn.CommitTrans
gDbNm = mCurrFirmDbNm
GProcConnectionOpen gCn
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
        
        
        X = " select tbladdLess.VType,ADINVNO,tbladdLess.Vdt,ADBILLAMT,TblMastAccount.AcName,tblBrk.AcName,tblBrk.AcContno,tblMill.AcName,tblMill.Accity," _
        & " (select sum(VSubBag) from tblvousub where tblvousub.Vno=tbladdless.vno and tblvousub.VType=tbladdless.vType and tblvousub.Vdt=tbladdless.vdt ), " _
        & " (select " & IIf(gBackEndDB = gBackEndAccess, "cdbl(sum(VSubWt))", "sum(VSubWt)") & " from tblvousub where tblvousub.Vno=tbladdless.vno and tblvousub.VType=tbladdless.vType and tblvousub.Vdt=tbladdless.vdt ),'" & gUserName & "'" _
        & " ,ADCFORMNO,TblMastAccount.AcCst,TblMastAccount.AcContno,tblmastaccount.accity,tblMastItem.Itname from tbladdLess,tblmastaccount,tblmastaccount tblBrk,tblmastaccount  tblMill,tblOutstanding,tblMastNarration,tblMastItem,tblVouSub where" _
        & " tblOutstanding.vNo=tbladdLess.vno and tblOutstanding.vType=tbladdLess.vType and tblOutstanding.vdt=tbladdLess.vdt" _
        & " and tblAddLess.Vno = tblVouSub.Vno and tblAddLess.Vtype = tblVouSub.VType and tblAddLess.Vyear = tblVouSub.Vyear " _
        & " and ADACDRCODE=tblmastaccount.Accode and ADMILLCODE=tblMill.AcCode and tblVousub.VSubItCode = tblMastitem.Itcode and " _
        & " OUTBROKERCODE=tblBrk.AcCode and ADTAXCODE=NARRCODE and (instr(narration,'C Form')>0 or instr(narration,'C S T')>0) and tbladdLess.Vtype in ('SM','SY','ST','SD','SO')  and " & CondStr
        
        '--- insert to setup db Table For All Firm
        GProcInsertToSetupDbTbl AllFirmFlag, X, "tmpGenTbl", "Vtype,VbillNo,VBillDt,Amt1,Nar2,Nar1,nar10,Nar4,Nar11,Amt2,Amt3,username,Nar6,nar7,nar8,nar9,Nar12 ", " username='" & gUserName & "'", "Nar3", "Nar5"
    
'        If gBackEndDB = gBackEndAccess Then
'            CondStr = " tbladdLess.vdt>= cdate('" & FromDt & "')  and tbladdLess.vdt<= cdate('" & ToDt & "')"
'            If OptStr = "Pending" Then '-- Pending
'               CondStr = CondStr + " and (ADCFORMNO='' or ADCFORMNO is null)" ' (ADCFORMNO is null or ADCFORMNO='') " '-
'            ElseIf OptStr = "Receive" Then '-- Receive
'                CondStr = CondStr + " and ADCFORMNO <>''"
'            End If
'        Else '-- Oracle
'            CondStr = " tbladdLess.vdt>= to_date('" & FromDt & "','dd/MM/yyyy') and tbladdLess.vdt<= to_date('" & ToDt & "','dd/MM/yyyy')"
'            If OptStr = "Pending" Then '-- Pending
'                CondStr = CondStr + " and ADCFORMNO is null "
'            ElseIf OptStr = "Receive" Then '-- Receive
'                CondStr = CondStr + " and ADCFORMNO is not null "
'            End If
'        End If
'
'        If Mill <> "" Then CondStr = CondStr + " and tblmill.AcName='" & Mill & "'" '--Mill
'        If Party <> "" Then CondStr = CondStr + " and TblMastAccount.AcName='" & Party & "'" '--Party
'        If Broker <> "" Then CondStr = CondStr + " and tblBrk.AcName='" & Broker & "'" '--Broker
'        X = " select tbladdLess.VType,ADINVNO,tbladdLess.Vdt,ADBILLAMT,TblMastAccount.AcName,TblMastAccount.AcAdd1,TblMastAccount.AcAdd2,TblMastAccount.AcCity,tblBrk.AcName,tblBrk.AcContno,tblMill.AcName,'" & gUserName & "'" _
'        & " ,ADCFORMNO,TblMastAccount.AcCst,TblMastAccount.AcContno,tblmastaccount.accity from tbladdLess,tblmastaccount,tblmastaccount tblBrk,tblmastaccount  tblMill,tblOutstanding,tblMastNarration where" _
'        & " tblOutstanding.vNo=tbladdLess.vno and tblOutstanding.vType=tbladdLess.vType and tblOutstanding.vdt=tbladdLess.vdt" _
'        & " and ADACDRCODE=tblmastaccount.Accode and ADMILLCODE=tblMill.AcCode and " _
'        & " OUTBROKERCODE=tblBrk.AcCode and ADTAXCODE=NARRCODE and instr(narration,'C Form')>0 and tbladdLess.Vtype in ('SM','SY','ST','SD','SO')  and " & CondStr
'        '--- insert to setup db Table For All Firm
'        GProcInsertToSetupDbTbl AllFirmFlag, X, "tmpGenTbl", "Vtype,VbillNo,VBillDt,Amt1,Nar2,Nar11,Nar12,Nar13,Nar1,nar10,Nar4,username,Nar6,nar7,nar8,nar9 ", " username='" & gUserName & "'", "Nar3", "Nar5"
 Case "CIButNotE1"
       CondStr = " tbladdLess.AdGpDt>= cdate('" & FromDt & "')  and tbladdLess.AdGpDt<= cdate('" & ToDt & "')"
       CondStr = CondStr + " and ADCFORMNO <>'' "
       CondStr = CondStr + " and (ADE1FORMNO ='' or ADE1FORMNO is null)"
        If Broker <> "" Then CondStr = CondStr + " and tblBrk.AcName='" & Broker & "'" '--Broker
        X = " select 1,tblAddless.Vtype,ADINVNO,tbladdLess.AdGpdt,ADSUBAMT,ADTAXAMT,ADBILLAMT,TblMastAccount.AcName,tblBrk.AcName,TblMastAccount.AcContNo,tblmastaccount.ACADD1 " & IIf(gBackEndDB = gBackEndAccess, " + ", " || ") & " tblmastaccount.AcAdd2,TblMastAccount.ACCITY ,TblMastAccount.ACBST,TblMastAccount.ACCST,tblBrk.AcName," _
        & " (select sum(VSubBag) from tblvousub where tblvousub.Vno=tbladdless.vno and tblvousub.VType=tbladdless.vType and tblvousub.Vdt=tbladdless.vdt ), " _
        & " (select " & IIf(gBackEndDB = gBackEndAccess, "cdbl(sum(VSubWt))", "sum(VSubWt)") & " from tblvousub where tblvousub.Vno=tbladdless.vno and tblvousub.VType=tbladdless.vType and tblvousub.Vdt=tbladdless.vdt ),'" & gUserName & "'" _
        & " ,ADCFORMNO from tbladdLess,tblmastaccount,tblMastNarration,tblmastaccount  tblBrk,tblOutStanding  where" _
        & "  ADACCRCODE=tblmastaccount.Accode and tblAddLess.Adtmp3=tblBrk.AcCode and tblOutstanding.vNo=tbladdLess.vno and tblOutstanding.vType=tbladdLess.vType and tblOutstanding.vdt=tbladdLess.vdt  " _
        & " and ADTAXCODE=NARRCODE and instr(narration,'C Form')>0 and tbladdLess.Vtype in ('PO','PT','PY')  and " & CondStr
        '--- insert to setup db Table For All Firm
        GProcInsertToSetupDbTbl AllFirmFlag, X, "tmpGenTbl", "vno,Vtype,VbillNo,VBillDt,Amt5,Amt4,Amt1,Nar2,Nar9,Nar10,Nar1,Nar3,Nar4,Nar5,Nar8,Amt2,Amt3,username,Nar7", "username='" & gUserName & "'", "Nar6", "Nar11"
 
 
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
        
        X = " select 1,tblAddless.Vtype,ADINVNO,tbladdLess.AdGpdt,ADSUBAMT,ADTAXAMT,ADBILLAMT,TblMastAccount.AcName,TblMastAccount.AcState,TblMastAccount.AcContNo,TblMastAccount.AcEmail,tblmastaccount.ACADD1 " & IIf(gBackEndDB = gBackEndAccess, " + ", " || ") & " tblmastaccount.AcAdd2,TblMastAccount.ACCITY ,TblMastAccount.ACBST,TblMastAccount.ACCST,tblBrk.AcName," _
        & " (select sum(VSubBag) from tblvousub where tblvousub.Vno=tbladdless.vno and tblvousub.VType=tbladdless.vType and tblvousub.Vdt=tbladdless.vdt ), " _
        & " (select " & IIf(gBackEndDB = gBackEndAccess, "cdbl(sum(VSubWt))", "sum(VSubWt)") & " from tblvousub where tblvousub.Vno=tbladdless.vno and tblvousub.VType=tbladdless.vType and tblvousub.Vdt=tbladdless.vdt ),'" & gUserName & "'" _
        & " ,ADCFORMNO from tbladdLess,tblmastaccount,tblMastNarration,tblmastaccount  tblBrk,tblOutStanding  where" _
        & "  ADACCRCODE=tblmastaccount.Accode and OUTBROKERCODE=tblBrk.AcCode and tblOutstanding.vNo=tbladdLess.vno and tblOutstanding.vType=tbladdLess.vType and tblOutstanding.vdt=tbladdLess.vdt  " _
        & " and ADTAXCODE=NARRCODE and instr(narration,'C Form')>0 and tbladdLess.Vtype in ('PO','PT','PY')  and " & CondStr
        '--- insert to setup db Table For All Firm
        GProcInsertToSetupDbTbl AllFirmFlag, X, "tmpGenTbl", "vno,Vtype,VbillNo,VBillDt,Amt5,Amt4,Amt1,Nar2,Nar9,Nar10,Nar11,Nar1,Nar3,Nar4,Nar5,Nar8,Amt2,Amt3,username,Nar7", "username='" & gUserName & "'", "Nar6"
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
        X = " select 1,Vtype,ADINVNO,tbladdLess.AdGpdt,ADBILLAMT,TblMastAccount.AcName,TblMastAccount.AcState,TblMastAccount.AcContNo,TblMastAccount.AcEmail,tblmastaccount.ACADD1 " & IIf(gBackEndDB = gBackEndAccess, " + ", " || ") & " tblmastaccount.AcAdd2,TblMastAccount.ACCITY ,TblMastAccount.ACBST,TblMastAccount.ACCST," _
        & " (select sum(VSubBag) from tblvousub where tblvousub.Vno=tbladdless.vno and tblvousub.VType=tbladdless.vType and tblvousub.Vdt=tbladdless.vdt ), " _
        & " (select " & IIf(gBackEndDB = gBackEndAccess, "cdbl(sum(VSubWt))", "sum(VSubWt)") & " from tblvousub where tblvousub.Vno=tbladdless.vno and tblvousub.VType=tbladdless.vType and tblvousub.Vdt=tbladdless.vdt ),'" & gUserName & "'" _
        & " ,ADCFORMNO from tbladdLess,tblmastaccount,tblMastNarration where" _
        & "  ADACCRCODE=tblmastaccount.Accode " _
        & " and ADTAXCODE=NARRCODE and tbladdLess.Vtype in ('PI')  and " & CondStr
        '--- insert to setup db Table For All Firm
        GProcInsertToSetupDbTbl AllFirmFlag, X, "tmpGenTbl", "vno,Vtype,VbillNo,VBillDt,Amt1,Nar2,Nar1,Nar9,Nar10,Nar11,Nar3,Nar4,Nar5,Amt2,Amt3,username,Nar7", "username='" & gUserName & "'", "Nar6"
 Case "FR" '--- F form Requirement
        If gBackEndDB = gBackEndAccess Then
            CondStr = " tbladdLess.AdGpDt>= cdate('" & FromDt & "')  and tbladdLess.AdGpDt<= cdate('" & ToDt & "')"
        Else '-- Oracle
            CondStr = " tbladdLess.AdGpDt>= to_date('" & FromDt & "','dd/MM/yyyy') and tbladdLess.AdGpDt<= to_date('" & ToDt & "','dd/MM/yyyy')"
        End If
        X = " select Vtype,ADINVNO,tbladdLess.AdGpdt,ADBILLAMT,TblMastAccount.AcName,tblmastaccount.ACADD1 " & IIf(gBackEndDB = gBackEndAccess, " + ", " || ") & " tblmastaccount.AcAdd2,TblMastAccount.ACCITY ,TblMastAccount.ACBST,TblMastAccount.ACCST," _
        & " (select sum(VSubBag) from tblvousub where tblvousub.Vno=tbladdless.vno and tblvousub.VType=tbladdless.vType and tblvousub.Vdt=tbladdless.vdt ), " _
        & " (select " & IIf(gBackEndDB = gBackEndAccess, "cdbl(sum(VSubWt))", "sum(VSubWt)") & " from tblvousub where tblvousub.Vno=tbladdless.vno and tblvousub.VType=tbladdless.vType and tblvousub.Vdt=tbladdless.vdt ),'" & gUserName & "'" _
        & " from tbladdLess,tblmastaccount where" _
        & "  ADACCRCODE=tblmastaccount.Accode " _
        & " and  tbladdLess.Vtype in ('PI')  and " & CondStr
        '--- insert to setup db Table For All Firm
        GProcInsertToSetupDbTbl AllFirmFlag, X, "tmpGenTbl", "Vtype,VbillNo,VBillDt,Amt1,Nar2,Nar1,Nar3,Nar4,Nar5,Amt2,Amt3,username", "username='" & gUserName & "'", "Nar6"
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
        If Broker <> "" Then CondStr = CondStr + " and tblBrk.AcName='" & Broker & "'" '--Broker
        If Mill <> "" Then CondStr = CondStr + " and tblMill.AcName='" & Mill & "'"  '--Broker
        X = " select 1,tblAddless.Vtype,ADINVNO,tbladdLess.AdGpdt,ADBILLAMT,TblMastAccount.AcName,tblMill.AcName,TblMastAccount.ACCITY ,TblMastAccount.ACBST,TblMastAccount.ACCST,tblBrk.AcName,   " _
        & " (select sum(tblvousub.VSubBag) from tblvousub where tblvousub.Vno=tbladdless.vno and tblvousub.VType=tbladdless.vType and tblvousub.Vdt=tbladdless.vdt ) as TotBag, " _
        & " (select " & IIf(gBackEndDB = gBackEndAccess, "cdbl(sum(VSubWt))", "sum(VSubWt)") & " from tblvousub where tblvousub.Vno=tbladdless.vno and tblvousub.VType=tbladdless.vType and tblvousub.Vdt=tbladdless.vdt ) AS TotWt,'" & gUserName & "'" _
        & " ,ADE1FORMNO,AdCFormNo  from tbladdLess,tblmastaccount,tblmastaccount  tblMill,tblMastNarration, tblmastaccount  tblBrk,tblOutStanding  where" _
        & "  ADACCRCODE=tblmastaccount.Accode and ADMILLCODE=tblMill.AcCode and OUTBROKERCODE=tblBrk.AcCode and tblOutstanding.vNo=tbladdLess.vno and tblOutstanding.vType=tbladdLess.vType and tblOutstanding.vdt=tbladdLess.vdt" _
        & " and ADTAXCODE=NARRCODE and instr(narration,'C Form')>0 and tbladdLess.Vtype in ('PT')  and " & CondStr
        '--- insert to setup db Table For All Firm
        GProcInsertToSetupDbTbl AllFirmFlag, X, "tmpGenTbl", "Vno,Vtype,VbillNo,VBillDt,Amt1,Nar2,Nar1,Nar3,Nar4,Nar5,Nar8,Amt2,Amt3,username,Nar7,Nar9", "username='" & gUserName & "'", "Nar6", "Nar10"
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
        
        X = " select tbladdLess.VType,ADINVNO,tbladdLess.Vdt,ADBILLAMT,TblMastAccount.AcName,tblBrk.AcName,tblMill.AcName, " _
        & " (select sum(tblvousub.VSubBag) from tblvousub where tblvousub.Vno=tbladdless.vno and tblvousub.VType=tbladdless.vType and tblvousub.Vdt=tbladdless.vdt ) as TotBag, " _
        & " (select " & IIf(gBackEndDB = gBackEndAccess, "cdbl(sum(VSubWt))", "sum(VSubWt)") & " from tblvousub where tblvousub.Vno=tbladdless.vno and tblvousub.VType=tbladdless.vType and tblvousub.Vdt=tbladdless.vdt ) AS TotWt,'" & gUserName & "'" _
        & " ,ADE1FORMNO,TblMastAccount.AcCst,ADCFORMNO from tbladdLess,tblmastaccount,tblmastaccount tblBrk,tblmastaccount  tblMill,tblOutstanding,tblMastNarration where" _
        & " tblOutstanding.vNo=tbladdLess.vno and tblOutstanding.vType=tbladdLess.vType and tblOutstanding.vdt=tbladdLess.vdt" _
        & " and ADACDRCODE=tblmastaccount.Accode and ADMILLCODE=tblMill.AcCode and " _
        & " OUTBROKERCODE=tblBrk.AcCode and ADTAXCODE=NARRCODE and instr(narration,'C Form')>0 and tbladdLess.Vtype in ('ST') and tbladdless.ADTMP2 = 1 and " & CondStr
        '--- insert to setup db Table For All Firm
        GProcInsertToSetupDbTbl AllFirmFlag, X, "tmpGenTbl", "Vtype,VbillNo,VBillDt,Amt1,Nar2,Nar1,Nar4,Amt2,Amt3,username,Nar6,Nar7,Nar9 ", " username='" & gUserName & "'", "Nar3", "Nar5"
            
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
        X = " select 1,tblAddless.Vtype,ADINVNO,tbladdLess.AdGpdt,ADBILLAMT,TblMastAccount.AcName,tblMill.AcName,TblMastAccount.ACCITY ,TblMastAccount.ACBST,TblMastAccount.ACCST,tblBrk.AcName,   " _
        & " (select sum(tblvousub.VSubBag) from tblvousub where tblvousub.Vno=tbladdless.vno and tblvousub.VType=tbladdless.vType and tblvousub.Vdt=tbladdless.vdt ) as TotBag, " _
        & " (select " & IIf(gBackEndDB = gBackEndAccess, "cdbl(sum(VSubWt))", "sum(VSubWt)") & " from tblvousub where tblvousub.Vno=tbladdless.vno and tblvousub.VType=tbladdless.vType and tblvousub.Vdt=tbladdless.vdt ) AS TotWt,'" & gUserName & "'" _
        & " ,ADE1FORMNO from tbladdLess,tblmastaccount,tblmastaccount  tblMill,tblMastNarration, tblmastaccount  tblBrk,tblOutStanding  where" _
        & "  ADACCRCODE=tblmastaccount.Accode and ADMILLCODE=tblMill.AcCode and OUTBROKERCODE=tblBrk.AcCode and tblOutstanding.vNo=tbladdLess.vno and tblOutstanding.vType=tbladdLess.vType and tblOutstanding.vdt=tbladdLess.vdt" _
        & " and ADTAXCODE=NARRCODE and instr(narration,'C Form')>0 and tbladdLess.Vtype in ('PT') and tbladdless.AdGpNo = 1 and " & CondStr
        
'        & " and ADTAXCODE=NARRCODE and instr(narration,'C Form')>0 and tbladdLess.Vtype in ('PT')  and " & CondStr
        '--- insert to setup db Table For All Firm
        GProcInsertToSetupDbTbl AllFirmFlag, X, "tmpGenTbl", "Vno,Vtype,VbillNo,VBillDt,Amt1,Nar2,Nar1,Nar3,Nar4,Nar5,Nar8,Amt2,Amt3,username,Nar7", "username='" & gUserName & "'", "Nar6", "Nar10"
        
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
            Else '-- Oracle
                CondStr = " tbladdLess.vdt>= to_date('" & SlbFDt & "','dd/MM/yyyy') and tbladdLess.vdt<= to_date('" & SlbTDt & "','dd/MM/yyyy')"
                '-- Pending
                CondStr = CondStr + " and ADCFORMNO is null "
            End If
            If Party <> "" Then CondStr = CondStr + " and TblMastAccount.AcName='" & Party & "'" '--Party
            If Broker <> "" Then CondStr = CondStr + " and tblBrk.AcName='" & Broker & "'" '--Broker
            '-- Union For 4 Slabs
            X = IIf(i > 1, X & " union all ", "") + " select " & i & " ," & gCYear & " ,tbladdLess.VType,ADINVNO,tbladdLess.Vdt,ADBILLAMT,TblMastAccount.AcName,TblMastAccount.AcContno,tblBrk.AcName,tblBrk.AcContno,tblMill.AcName,'" & gUserName & "'" _
            & " from tbladdLess,tblmastaccount,tblmastaccount tblBrk,tblmastaccount  tblMill,tblOutstanding,tblMastNarration where" _
            & " tblOutstanding.vNo=tbladdLess.vno and tblOutstanding.vType=tbladdLess.vType and tblOutstanding.vdt=tbladdLess.vdt" _
            & " and ADACDRCODE=tblmastaccount.Accode and ADMILLCODE=tblMill.AcCode and " _
            & " OUTBROKERCODE=tblBrk.AcCode and ADTAXCODE=NARRCODE and instr(narration,'C Form')>0 and tbladdLess.Vtype in ('SM','SY','ST','SD','SO')  and " & CondStr
            i = i + 1
        Loop
        '--- insert to setup db Table For All Firm
            GProcInsertToSetupDbTbl AllFirmFlag, X, "tmpGenTbl", "VSrno,VYear,Vtype,VbillNo,VBillDt,Amt1,Nar2,Nar10,Nar1,nar9,Nar4,username", " username='" & gUserName & "'", "Nar3", "Nar5"
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
            If Mill <> "" Then CondStr = CondStr + " and AcName='" & Mill & "'"  '--Mill
            '-- Union For 4 Slabs
            X = IIf(i > 1, X & " union all ", "") + " select " & i & " ," & gCYear & " ,VType,ADINVNO,AdGpdt,ADBILLAMT,AcName,'" & gUserName & "'" _
            & " from tbladdLess,tblmastaccount,tblMastNarration where ADMILLCODE=AcCode and " _
            & " ADTAXCODE=NARRCODE and instr(narration,'C Form')>0 and tbladdLess.Vtype in ('PT')  and " & CondStr
            i = i + 1
        Loop
        '--- insert to setup db Table For All Firm
            GProcInsertToSetupDbTbl AllFirmFlag, X, "tmpGenTbl", "VSrno,VYear,Vtype,VbillNo,VBillDt,Amt1,Nar4,username", " username='" & gUserName & "'", "Nar3", "Nar5"
     Case "CsBnkStatus"
        X = "select AcName,Clbal_Dr,ClBal_cr, '" & gUserName & "' from qryTrbalDetail where qryTrbalDetail.AgCode = 90019  and username='" & gUserName & "'"
        GProcInsertToSetupDbTbl AllFirmFlag, X, "tmpGenTbl", "Nar1,Amt1,Amt2,username", " username='" & gUserName & "'", "Nar3", "Nar5"

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
Dim X1 As String
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
X1 = "select '" & gUserName & "',RECVNO,RECVTYPE,RECVYEAR,INVBILLNO,ADJAMT,0 as OrdNo1,RECVCTRLNO as OrdNo2 from tblRecVsSale where " & X & "" _
& " union all select '" & gUserName & "',RECVNO,RECVTYPE,RECVYEAR ,'On A/c',RECAMT- sum(ADJAMT),1,0  from tblRecVsSale where " & X _
& " group by RECVNO,RECVTYPE,RECVYEAR,RECAMT having RECAMT- sum(ADJAMT) <>0 order by RECVTYPE,RECVNO,RECVYEAR,OrdNo1,OrdNo2"
If GProcRstOpen(Rs1, X1, "R") > 0 Then
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
            TotStr = IIf(SelOptStr = "NAmt", "sum(ADBILLAMT-AdTaxAmt-AdTcsAmt)", IIf(SelOptStr = "GAmt", "sum(ADBILLAMT)", " Sum(VSUBBAG)"))
            '-- SY,ST,PY,PT
            X = "select " & TotStr & ",0,'" & gUserName & "'" _
            & RelationStr & " tbladdless.Vtype in ('SY','ST')  and " & CondStr _
            & " union all select 0," & TotStr & ",'" & gUserName & "'" _
            & RelationStr & " tbladdless.Vtype in ('PY','PT')  and " & CondStr _
            '--- insert to setup db Table For All Firm
            GProcInsertToSetupDbTbl AllFirmFlag, X, "tmpGenTbl", "Amt1,Amt2,username,Nar1,Vno,Vyear", "username='" & gUserName & "'", "Nar5"
       Case "MonthWisePurchase"
            If gBackEndDB = gBackEndAccess Then
                CondStr = CondStr + " group by Format(tbladdless.vdt,'MMMM'),format(tbladdless.vdt,'MM'),Format(tbladdless.vdt,'yy')"
            Else '--Oracle
                CondStr = CondStr + " group by to_char(tbladdless.vdt,'Month'),to_char(tbladdless.vdt,'MM'),to_char(tbladdless.vdt,'yy')"
            End If
            '-Relation Str
            RelationStr = IIf(gBackEndDB = gBackEndAccess, ",Format(tbladdless.vdt,'MMMM'),format(tbladdless.vdt,'MM'),Format(tbladdless.vdt,'yy')", ",to_char(tbladdless.vdt,'Month'),to_char(tbladdless.vdt,'MM'),to_char(tbladdless.vdt,'yy')") & " from tbladdless,tblvouSub" _
            & " where tbladdless.vno=tblVousub.vno and tbladdless.vType=tblVousub.vType and tbladdless.VYear=tblVousub.Vyear and vsubctrno=1 and "
            '-- Total Str
            TotStr = IIf(SelOptStr = "NAmt", "sum(ADBILLAMT-AdTaxAmt)", IIf(SelOptStr = "GAmt", "sum(ADBILLAMT)", " Sum(VSUBBAG)"))
            '-- PY,PT,SD,ST,SM
            X = "select " & TotStr & ",0,0,0,0,'" & gUserName & "'" _
            & RelationStr & " tbladdless.Vtype = 'PY' and " & CondStr _
            & " union all select 0," & TotStr & ",0,0,0,'" & gUserName & "'" _
            & RelationStr & " tbladdless.Vtype = 'PT' and " & CondStr _
            & " union all select 0,0," & TotStr & ",0,0,'" & gUserName & "'" _
            & RelationStr & " tbladdless.Vtype = 'XX' and " & CondStr _
            & " union all select 0,0,0," & TotStr & ",0,'" & gUserName & "'" _
            & RelationStr & " tbladdless.Vtype = 'YY' and " & CondStr _
            & " union all select 0,0,0,0," & TotStr & ",'" & gUserName & "'" _
            & RelationStr & " tbladdless.Vtype = 'ZZ' and " & CondStr
            '--- insert to setup db Table For All Firm
            GProcInsertToSetupDbTbl AllFirmFlag, X, "tmpGenTbl", "Amt1,Amt2,Amt3,Amt4,Amt5,username,Nar1,Vno,Vyear", "username='" & gUserName & "'", "Nar5"
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
Dim X1 As String
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
 X1 = "{tmpGentbl.UserName}='" & gUserName & "'"
 .SelectionFormula = X1
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
     X1 = "{tblAddLess.vdt} in Date(" & Format(FromDt, "yyyy,mm,dd") & ") to Date(" & Format(ToDt, "yyyy,mm,dd") & ")"
     X1 = X1 + " and {tblAddLess.VType}='PI' and {tblMastAccount.AcName}='" & MillName & "'"
     .SelectionFormula = X1
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
     X1 = "{tblAddLess.vdt} in Date(" & Format(FromDt, "yyyy,mm,dd") & ") to Date(" & Format(ToDt, "yyyy,mm,dd") & ")"
     'x1 = x1 + " and ({tblAddLess.VType}='SY' or {tblAddLess.VType}='ST' or {tblAddLess.VType}='SO' or {tblAddLess.VType}='SD' or {tblAddLess.VType}='SM') and {tblMastAccount.AcName}='" & txtMillName & "'"
     X1 = X1 + " and ({tblAddLess.VType}='SO' or {tblAddLess.VType}='SD' ) and {tblMastAccount.AcName}='" & MillName & "'"
     .SelectionFormula = X1
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
 X1 = "{tmpGentbl.UserName}='" & gUserName & "'"
 .SelectionFormula = X1
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
     X1 = "{tblVoucher.vdt} in Date(" & Format(FromDt, "yyyy,mm,dd") & ") to Date(" & Format(ToDt, "yyyy,mm,dd") & ")"
     X1 = X1 + " and ({tblVoucher.VType}='BR' or {tblVoucher.VType}='CR' or {tblVoucher.VType}='JV' or {tblVoucher.VType}='CP' or {tblVoucher.VType}='BP' or {tblVoucher.VType}='SN' or {tblVoucher.VType}='PN') and {tblMastAccount.AcName}='" & MillName & "'"
     .SelectionFormula = X1
     .WindowTitle = "Payment Details"
     .ReportFileName = gReportPath & "rptMonSTMConsgDepotPayment.rpt"
     .Action = 1
 End If
 '=============== Debit Note Commission
 If CDbl(CommAmt) > 0 Then
     CommType = GProcGetCommissionType(GProcGetColumnValue("TblMastAccount", "AcName", MillName, "S", "AcComType", "N"))
     RptEnclosure(RptCount) = CommType & " Debit Note"
     RptCount = RptCount + 1
     X1 = "{tblMastAccount.AcName}='" & MillName & "'"
     .Formulas(11) = "WAmountInWords = '( " & GProcAmountInWords(CDbl(CommAmt)) & ")'"
     .Formulas(12) = "CommAmt = '" & CDbl(CommAmt) & "'"
     .Formulas(13) = "CommOnAmt = '" & CDbl(CommOn) & "'"
     .Formulas(14) = "CommBag = '" & CLng(CommBag) & "'"
     .Formulas(15) = "CommType = '" & CommType & "'"
     .SelectionFormula = X1
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
X1 = " select sum(ADSUBAMT),AcComRt  as CommAmt from tblAddLess,tblMastAccount where " _
& " tblAddLess.AdMillCode=tblMastAccount.AcCode and tblAddLess.VType in ('PI')  and " & CondStr & " and ADMILLCODE=" & MillCode _
& " and AcComRt>0 and AcComOn=1 group by AcComRt"
 If GProcRstOpen(Rs, X1, "R") > 0 Then
    If IsNumeric(Rs.Fields(0)) = True Then
        X1 = "{tblAddLess.vdt} in Date(" & Format(FromDt, "yyyy,mm,dd") & ") to Date(" & Format(ToDt, "yyyy,mm,dd") & ")"
        X1 = X1 + " and {tblAddLess.VType}='PI' and {tblMastAccount.AcComRt}>0 and {tblMastAccount.AcComOn}=1 and {tblMastAccount.AcName}='" & MillName & "'"
        .SelectionFormula = X1
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
     X1 = "{tmpGentbl.UserName}='" & gUserName & "'"
     .SelectionFormula = X1
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
     X1 = "{tblAddLess.vdt} in Date(" & Format(FromDt, "yyyy,mm,dd") & ") to Date(" & Format(ToDt, "yyyy,mm,dd") & ")"
     X1 = X1 + " and ({tblAddLess.VType}='SO' or {tblAddLess.VType}='SD' ) and {tblMastItem.ITBROKRT}>0 and {tblMastAccount.AcName}='" & MillName & "'"
     .SelectionFormula = X1
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
 X1 = "{tmpGentbl.UserName}='" & gUserName & "'"
 .SelectionFormula = X1
 .WindowTitle = "Cosignment/ Depot Statement"
 .ReportFileName = gReportPath & "rptMonSTMConsgDepot.rpt"
 For i = 0 To RptCount
     If RptEnclosure(i) <> "" Then .Formulas(i + 10) = "Str" & i & " ='" & i + 1 & " ) " & RptEnclosure(i) & "'"
 Next i
 .Action = 1
 End With
End Sub
Public Function GProcPurpareGSTSummery(w_Fr As Date, w_To As Date, Optional IsOnPurBillDt As Boolean, Optional SplR As Boolean)
Dim X1 As String


             gCn.BeginTrans
             gCn.Execute "delete from tmpGenTbl where username='" & gUserName & "'"
             '---Sales Detail
             If SplR = True Then
                X1 = " tblAddLess.VType in ('SY')  and tblAddLess.vdt>=cdate('" & Format(w_Fr, "yyyy,mm,dd") & "')  and tblAddLess.vdt<=Cdate('" & Format(w_To, "yyyy,mm,dd") & "') and tblVouSub.VsubIGSTAmt+tblVouSub.VsubCGSTAmt+tblVouSub.VsubSGSTAmt > 0 "  '
                gCn.Execute " Insert Into TmpGentbl (Nar1,Nar3,Amt2,Amt3,Amt4,Amt5,Amt6,UserName)" _
                            & " select 'Output',iif(isnull(tblAddLess.ADTMP5 ),'Trade Sales',iif(tblAddLess.ADTMP5= 'COL','Cloth Sales',iif(tblAddLess.ADTMP5= 'OTH','Other Sales',iif(tblAddLess.ADTMP5= 'COT','Cotton','XXX')))),sum(tblVouSub.VSubAmt),sum(tblVouSub.VsubIGSTAmt),sum(tblVouSub.VsubCGSTAmt),sum(tblVouSub.VsubSGSTAmt),sum(iif(tblVouSub.VSubCtrNo = 1,tblAddLess.AdBillAmt,0)),'" & gUserName & "'" _
                            & " From tblAddLess,tblVouSub where tblAddLess.Vno = tblVouSub.Vno and tblAddLess.Vtype = tblVouSub.Vtype and tblAddLess.Vyear = tblVouSub.Vyear and " & X1 & "" _
                            & " group by iif(isnull(tblAddLess.ADTMP5 ),'Trade Sales',iif(tblAddLess.ADTMP5= 'COL','Cloth Sales',iif(tblAddLess.ADTMP5= 'OTH','Other Sales',iif(tblAddLess.ADTMP5= 'COT','Cotton','XXX'))))"
             Else
                X1 = " tblAddLess.VType in ('SY')  and tblAddLess.vdt>=cdate('" & Format(w_Fr, "yyyy,mm,dd") & "')  and tblAddLess.vdt<=Cdate('" & Format(w_To, "yyyy,mm,dd") & "') and tblVouSub.VsubIGSTAmt+tblVouSub.VsubCGSTAmt+tblVouSub.VsubSGSTAmt > 0 "  '
                gCn.Execute " Insert Into TmpGentbl (Nar1,vyear,Amt1,Nar2,Nar3,Amt2,Amt3,Amt4,Amt5,Amt6,UserName)" _
                            & " select 'Output',year(tblAddLess.vdt),month(tblAddLess.Vdt), format(tblAddLess.Vdt,'mmm-yyyy'),iif(isnull(tblAddLess.ADTMP5 ),'Trade Sales',iif(tblAddLess.ADTMP5= 'COL','Cloth Sales',iif(tblAddLess.ADTMP5= 'OTH','Other Sales',iif(tblAddLess.ADTMP5= 'COT','Cotton','XXX')))),sum(tblVouSub.VSubAmt),sum(tblVouSub.VsubIGSTAmt),sum(tblVouSub.VsubCGSTAmt),sum(tblVouSub.VsubSGSTAmt),sum(iif(tblVouSub.VSubCtrNo = 1,tblAddLess.AdBillAmt,0)),'" & gUserName & "'" _
                            & " From tblAddLess,tblVouSub where tblAddLess.Vno = tblVouSub.Vno and tblAddLess.Vtype = tblVouSub.Vtype and tblAddLess.Vyear = tblVouSub.Vyear and " & X1 & "" _
                            & " group by year(tblAddLess.vdt),month(tblAddLess.Vdt), format(tblAddLess.Vdt,'mmm-yyyy'),iif(isnull(tblAddLess.ADTMP5 ),'Trade Sales',iif(tblAddLess.ADTMP5= 'COL','Cloth Sales',iif(tblAddLess.ADTMP5= 'OTH','Other Sales',iif(tblAddLess.ADTMP5= 'COT','Cotton','XXX'))))"
             End If
             
             '--- Debit Note Sales
             X1 = " qrytblVoucherSr1.VType = 'SN' and qrytblVoucherSr1.VDRCRNTTYPE = 'B' and qrytblVoucherSr1.vdt>=cdate('" & Format(w_Fr, "yyyy,mm,dd") & "')  and qrytblVoucherSr1.vdt<=Cdate('" & Format(w_To, "yyyy,mm,dd") & "') and (qrytblVoucherSr1.VIGSTAmt+qrytblVoucherSr1.VCGSTAmt+qrytblVoucherSr1.VSGSTAmt) > 0 "  '
             If SplR = True Then
                gCn.Execute " Insert Into TmpGentbl (Nar1,Nar3,Amt2,Amt3,Amt4,Amt5,Amt6,UserName)" _
                            & " Select 'Output','Debit Note (Sales)',sum(qrytblVoucherSr1.VTaxableAmt),sum(qrytblVoucherSr1.VIGSTAmt),sum(qrytblVoucherSr1.VCGSTAmt),sum(qrytblVoucherSr1.VSGSTAmt),sum(abs(qrytblVoucherSr1.VAmt))," _
                            & " '" & gUserName & "' from qrytblVoucherSr1 where " & X1 & " " _
                            & " group by 'Debit Note (Sales)'"
             Else
                gCn.Execute " Insert Into TmpGentbl (Nar1,vyear,Amt1,Nar2,Nar3,Amt2,Amt3,Amt4,Amt5,Amt6,UserName)" _
                            & " Select 'Output',year(qrytblVoucherSr1.vdt),month(qrytblVoucherSr1.Vdt), format(qrytblVoucherSr1.Vdt,'mmm-yyyy'),'Debit Note (Sales)',sum(qrytblVoucherSr1.VTaxableAmt),sum(qrytblVoucherSr1.VIGSTAmt),sum(qrytblVoucherSr1.VCGSTAmt),sum(qrytblVoucherSr1.VSGSTAmt),sum(abs(qrytblVoucherSr1.VAmt))," _
                            & " '" & gUserName & "' from qrytblVoucherSr1 where " & X1 & " " _
                            & " group by year(qrytblVoucherSr1.vdt),month(qrytblVoucherSr1.Vdt), format(qrytblVoucherSr1.Vdt,'mmm-yyyy'),'Debit Note (Sales)'"
             End If
             
             '--- Credit Note Sales
             X1 = " qrytblVoucherSr1.VType = 'PN' and qrytblVoucherSr1.VDRCRNTTYPE = 'B' and qrytblVoucherSr1.vdt>=cdate('" & Format(w_Fr, "yyyy,mm,dd") & "')  and qrytblVoucherSr1.vdt<=Cdate('" & Format(w_To, "yyyy,mm,dd") & "') and (qrytblVoucherSr1.VIGSTAmt+qrytblVoucherSr1.VCGSTAmt+qrytblVoucherSr1.VSGSTAmt) > 0 "  '
             If SplR = True Then
                gCn.Execute " Insert Into TmpGentbl (Nar1,Nar3,Amt2,Amt3,Amt4,Amt5,Amt6,UserName)" _
                            & " Select 'Output','Credit Note (Sales)',sum(qrytblVoucherSr1.VTaxableAmt)*-1,sum(qrytblVoucherSr1.VIGSTAmt)*-1,sum(qrytblVoucherSr1.VCGSTAmt)*-1,sum(qrytblVoucherSr1.VSGSTAmt)*-1,sum(abs(qrytblVoucherSr1.VAmt))*-1," _
                            & " '" & gUserName & "' from qrytblVoucherSr1 where " & X1 & " " _
                            & " group by 'Debit Note (Sales)'"
             Else
                gCn.Execute " Insert Into TmpGentbl (Nar1,vyear,Amt1,Nar2,Nar3,Amt2,Amt3,Amt4,Amt5,Amt6,UserName)" _
                            & " Select 'Output',year(qrytblVoucherSr1.vdt),month(qrytblVoucherSr1.Vdt), format(qrytblVoucherSr1.Vdt,'mmm-yyyy'),'Credit Note (Sales)',sum(qrytblVoucherSr1.VTaxableAmt)*-1,sum(qrytblVoucherSr1.VIGSTAmt)*-1,sum(qrytblVoucherSr1.VCGSTAmt)*-1,sum(qrytblVoucherSr1.VSGSTAmt)*-1,sum(abs(qrytblVoucherSr1.VAmt))*-1," _
                            & " '" & gUserName & "' from qrytblVoucherSr1 where " & X1 & " " _
                            & " group by year(qrytblVoucherSr1.vdt),month(qrytblVoucherSr1.Vdt), format(qrytblVoucherSr1.Vdt,'mmm-yyyy'),'Debit Note (Sales)'"
             End If
             '--- Sales Return
             X1 = " tblAddLess.VType = 'RY' and tblAddLess.Vdt>=cdate('" & Format(w_Fr, "yyyy,mm,dd") & "')  and tblAddLess.Vdt<=Cdate('" & Format(w_To, "yyyy,mm,dd") & "') and tblVouSub.VsubIGSTAmt+tblVouSub.VsubCGSTAmt+tblVouSub.VsubSGSTAmt > 0 "  '
             If SplR = True Then
                gCn.Execute " Insert Into TmpGentbl (Nar1,Nar3,Amt2,Amt3,Amt4,Amt5,Amt6,UserName)" _
                            & " select 'Output','Sales Rtn',iif(isnull(sum(tblVouSub.VSubAmt)),0,sum(tblVouSub.VSubAmt))*-1,iif(isnull(sum(tblVouSub.VsubIGSTAmt)),0,sum(tblVouSub.VsubIGSTAmt))*-1,iif(isnull(sum(tblVouSub.VsubCGSTAmt)),0,sum(tblVouSub.VsubCGSTAmt))*-1,iif(isnull(sum(tblVouSub.VsubSGSTAmt)),0,sum(tblVouSub.VsubSGSTAmt))*-1,iif(isnull(sum(tblAddLess.AdBillAmt)),0,sum(iif(tblVouSub.VSubCtrNo = 1,tblAddLess.AdBillAmt,0)))*-1,'" & gUserName & "'" _
                            & " From tblAddLess,tblVouSub where tblAddLess.Vno = tblVouSub.Vno and tblAddLess.Vtype = tblVouSub.Vtype and tblAddLess.Vyear = tblVouSub.Vyear and " & X1 & ""
             Else
                gCn.Execute " Insert Into TmpGentbl (Nar1,vyear,Amt1,Nar2,Nar3,Amt2,Amt3,Amt4,Amt5,Amt6,UserName)" _
                            & " select 'Output',year(tblAddLess.Vdt),month(tblAddLess.Vdt), format(tblAddLess.Vdt,'mmm-yyyy'),'Sales Rtn',sum(tblVouSub.VSubAmt)*-1,sum(tblVouSub.VsubIGSTAmt)*-1,sum(tblVouSub.VsubCGSTAmt)*-1,sum(tblVouSub.VsubSGSTAmt)*-1,sum(iif(tblVouSub.VSubCtrNo = 1,tblAddLess.AdBillAmt,0))*-1,'" & gUserName & "'" _
                            & " From tblAddLess,tblVouSub where tblAddLess.Vno = tblVouSub.Vno and tblAddLess.Vtype = tblVouSub.Vtype and tblAddLess.Vyear = tblVouSub.Vyear and " & X1 & "" _
                            & " group by year(tblAddLess.Vdt),month(tblAddLess.Vdt), format(tblAddLess.Vdt,'mmm-yyyy')"
             End If
             
             
            '--- Purch & Other Purch
             If IsOnPurBillDt = True Then
                X1 = " tblAddLess.VType ='PY' and tblAddLess.AdGpDt>=cdate('" & Format(w_Fr, "yyyy,mm,dd") & "')  and tblAddLess.AdGpDt<=Cdate('" & Format(w_To, "yyyy,mm,dd") & "') and tblVouSub.VsubIGSTAmt+tblVouSub.VsubCGSTAmt+tblVouSub.VsubSGSTAmt > 0 "
                If SplR = True Then
                    gCn.Execute " Insert Into TmpGentbl (Nar1,Nar3,Amt2,Amt3,Amt4,Amt5,Amt6,UserName)" _
                                & " select 'Input',iif(tblAddLess.Vtype = 'PO','Other Purch',iif(isnull(tblAddLess.ADTMP5),'Trade Purch',iif(tblAddLess.ADTMP5 = 'CLO','Cloth Purch',iif(tblAddLess.ADTMP5 = 'COT','Cotton Purch','XXXX')))),sum(tblVouSub.VSubTaxableAmt),sum(tblVouSub.VsubIGSTAmt),sum(tblVouSub.VsubCGSTAmt),sum(tblVouSub.VsubSGSTAmt),sum(iif(tblVouSub.VSubCtrNo = 1,tblAddLess.AdBillAmt,0)),'" & gUserName & "'" _
                                & " From tblAddLess,tblVouSub where tblAddLess.Vno = tblVouSub.Vno and tblAddLess.Vtype = tblVouSub.Vtype and tblAddLess.Vyear = tblVouSub.Vyear and " & X1 & "" _
                                & " group by iif(tblAddLess.Vtype = 'PO','Other Purch',iif(isnull(tblAddLess.ADTMP5),'Trade Purch',iif(tblAddLess.ADTMP5 = 'CLO','Cloth Purch',iif(tblAddLess.ADTMP5 = 'COT','Cotton Purch','XXXX'))))"
                Else
                    gCn.Execute " Insert Into TmpGentbl (Nar1,vyear,Amt1,Nar2,Nar3,Amt2,Amt3,Amt4,Amt5,Amt6,UserName)" _
                                & " select 'Input',year(tblAddLess.AdGpDt),month(tblAddLess.AdGpDt), format(tblAddLess.AdGpDt,'mmm-yyyy'),iif(tblAddLess.Vtype = 'PO','Other Purch',iif(isnull(tblAddLess.ADTMP5),'Trade Purch',iif(tblAddLess.ADTMP5 = 'CLO','Cloth Purch',iif(tblAddLess.ADTMP5 = 'COT','Cotton Purch','XXXX')))),sum(tblVouSub.VSubTaxableAmt),sum(tblVouSub.VsubIGSTAmt),sum(tblVouSub.VsubCGSTAmt),sum(tblVouSub.VsubSGSTAmt),sum(iif(tblVouSub.VSubCtrNo = 1,tblAddLess.AdBillAmt,0)),'" & gUserName & "'" _
                                & " From tblAddLess,tblVouSub where tblAddLess.Vno = tblVouSub.Vno and tblAddLess.Vtype = tblVouSub.Vtype and tblAddLess.Vyear = tblVouSub.Vyear and " & X1 & "" _
                                & " group by year(tblAddLess.AdGpDt),month(tblAddLess.AdGpDt), format(tblAddLess.AdGpDt,'mmm-yyyy'),iif(tblAddLess.Vtype = 'PO','Other Purch',iif(isnull(tblAddLess.ADTMP5),'Trade Purch',iif(tblAddLess.ADTMP5 = 'CLO','Cloth Purch',iif(tblAddLess.ADTMP5 = 'COT','Cotton Purch','XXXX'))))"
                End If
                X1 = " tblAddLess.VType ='PO' and tblAddLess.AdGpDt>=cdate('" & Format(w_Fr, "yyyy,mm,dd") & "')  and tblAddLess.AdGpDt<=Cdate('" & Format(w_To, "yyyy,mm,dd") & "') and tblVouSub.VsubIGSTAmt+tblVouSub.VsubCGSTAmt+tblVouSub.VsubSGSTAmt > 0 and tblAddLess.AdLCostAdd+tblAddLess.AdLCostFreight+tblAddLess.AdLCostGRent = 0"
                If SplR = True Then
                    gCn.Execute " Insert Into TmpGentbl (Nar1,Nar3,Amt2,Amt3,Amt4,Amt5,Amt6,UserName)" _
                                & " select 'Input',iif(tblAddLess.Vtype = 'PO','Other Purch',iif(isnull(tblAddLess.ADTMP5),'Trade Purch',iif(tblAddLess.ADTMP5 = 'CLO','Cloth Purch',iif(tblAddLess.ADTMP5 = 'COT','Cotton Purch','XXXX')))),sum(tblVouSub.VSubTaxableAmt),sum(tblVouSub.VsubIGSTAmt),sum(tblVouSub.VsubCGSTAmt),sum(tblVouSub.VsubSGSTAmt),sum(iif(tblVouSub.VSubCtrNo = 1,tblAddLess.AdBillAmt,0)),'" & gUserName & "'" _
                                & " From tblAddLess,tblVouSub where tblAddLess.Vno = tblVouSub.Vno and tblAddLess.Vtype = tblVouSub.Vtype and tblAddLess.Vyear = tblVouSub.Vyear and " & X1 & "" _
                                & " group by iif(tblAddLess.Vtype = 'PO','Other Purch',iif(isnull(tblAddLess.ADTMP5),'Trade Purch',iif(tblAddLess.ADTMP5 = 'CLO','Cloth Purch',iif(tblAddLess.ADTMP5 = 'COT','Cotton Purch','XXXX'))))"
                Else
                    gCn.Execute " Insert Into TmpGentbl (Nar1,vyear,Amt1,Nar2,Nar3,Amt2,Amt3,Amt4,Amt5,Amt6,UserName)" _
                                & " select 'Input',year(tblAddLess.AdGpDt),month(tblAddLess.AdGpDt), format(tblAddLess.AdGpDt,'mmm-yyyy'),iif(tblAddLess.Vtype = 'PO','Other Purch',iif(isnull(tblAddLess.ADTMP5),'Trade Purch',iif(tblAddLess.ADTMP5 = 'CLO','Cloth Purch',iif(tblAddLess.ADTMP5 = 'COT','Cotton Purch','XXXX')))),sum(tblVouSub.VSubTaxableAmt),sum(tblVouSub.VsubIGSTAmt),sum(tblVouSub.VsubCGSTAmt),sum(tblVouSub.VsubSGSTAmt),sum(iif(tblVouSub.VSubCtrNo = 1,tblAddLess.AdBillAmt,0)),'" & gUserName & "'" _
                                & " From tblAddLess,tblVouSub where tblAddLess.Vno = tblVouSub.Vno and tblAddLess.Vtype = tblVouSub.Vtype and tblAddLess.Vyear = tblVouSub.Vyear and " & X1 & "" _
                                & " group by year(tblAddLess.AdGpDt),month(tblAddLess.AdGpDt), format(tblAddLess.AdGpDt,'mmm-yyyy'),iif(tblAddLess.Vtype = 'PO','Other Purch',iif(isnull(tblAddLess.ADTMP5),'Trade Purch',iif(tblAddLess.ADTMP5 = 'CLO','Cloth Purch',iif(tblAddLess.ADTMP5 = 'COT','Cotton Purch','XXXX'))))"
                End If
             Else
                X1 = " tblAddLess.VType ='PY' and tblAddLess.Vdt>=cdate('" & Format(w_Fr, "yyyy,mm,dd") & "')  and tblAddLess.Vdt<=Cdate('" & Format(w_To, "yyyy,mm,dd") & "') and tblVouSub.VsubIGSTAmt+tblVouSub.VsubCGSTAmt+tblVouSub.VsubSGSTAmt > 0 "  '
                If SplR = True Then
                    gCn.Execute " Insert Into TmpGentbl (Nar1,Nar3,Amt2,Amt3,Amt4,Amt5,Amt6,UserName)" _
                                & " select 'Input',iif(tblAddLess.Vtype = 'PO','Other Purch',iif(isnull(tblAddLess.ADTMP5),'Trade Purch',iif(tblAddLess.ADTMP5 = 'CLO','Cloth Purch',iif(tblAddLess.ADTMP5 = 'COT','Cotton Purch','XXXX')))),sum(tblVouSub.VSubTaxableAmt),sum(tblVouSub.VsubIGSTAmt),sum(tblVouSub.VsubCGSTAmt),sum(tblVouSub.VsubSGSTAmt),sum(iif(tblVouSub.VSubCtrNo = 1,tblAddLess.AdBillAmt,0)),'" & gUserName & "'" _
                                & " From tblAddLess,tblVouSub where tblAddLess.Vno = tblVouSub.Vno and tblAddLess.Vtype = tblVouSub.Vtype and tblAddLess.Vyear = tblVouSub.Vyear and " & X1 & "" _
                                & " group by iif(tblAddLess.Vtype = 'PO','Other Purch',iif(isnull(tblAddLess.ADTMP5),'Trade Purch',iif(tblAddLess.ADTMP5 = 'CLO','Cloth Purch',iif(tblAddLess.ADTMP5 = 'COT','Cotton Purch','XXXX'))))"
                Else
                    gCn.Execute " Insert Into TmpGentbl (Nar1,vyear,Amt1,Nar2,Nar3,Amt2,Amt3,Amt4,Amt5,Amt6,UserName)" _
                                & " select 'Input',year(tblAddLess.Vdt),month(tblAddLess.Vdt), format(tblAddLess.Vdt,'mmm-yyyy'),iif(tblAddLess.Vtype = 'PO','Other Purch',iif(isnull(tblAddLess.ADTMP5),'Trade Purch',iif(tblAddLess.ADTMP5 = 'CLO','Cloth Purch',iif(tblAddLess.ADTMP5 = 'COT','Cotton Purch','XXXX')))),sum(tblVouSub.VSubTaxableAmt),sum(tblVouSub.VsubIGSTAmt),sum(tblVouSub.VsubCGSTAmt),sum(tblVouSub.VsubSGSTAmt),sum(iif(tblVouSub.VSubCtrNo = 1,tblAddLess.AdBillAmt,0)),'" & gUserName & "'" _
                                & " From tblAddLess,tblVouSub where tblAddLess.Vno = tblVouSub.Vno and tblAddLess.Vtype = tblVouSub.Vtype and tblAddLess.Vyear = tblVouSub.Vyear and " & X1 & "" _
                                & " group by year(tblAddLess.Vdt),month(tblAddLess.Vdt), format(tblAddLess.Vdt,'mmm-yyyy'),iif(tblAddLess.Vtype = 'PO','Other Purch',iif(isnull(tblAddLess.ADTMP5),'Trade Purch',iif(tblAddLess.ADTMP5 = 'CLO','Cloth Purch',iif(tblAddLess.ADTMP5 = 'COT','Cotton Purch','XXXX'))))"
                End If
                X1 = " tblAddLess.VType ='PO' and tblAddLess.Vdt>=cdate('" & Format(w_Fr, "yyyy,mm,dd") & "')  and tblAddLess.Vdt<=Cdate('" & Format(w_To, "yyyy,mm,dd") & "') and tblVouSub.VsubIGSTAmt+tblVouSub.VsubCGSTAmt+tblVouSub.VsubSGSTAmt > 0 and tblAddLess.AdLCostAdd+tblAddLess.AdLCostFreight+tblAddLess.AdLCostGRent = 0"  '
                If SplR = True Then
                    gCn.Execute " Insert Into TmpGentbl (Nar1,Nar3,Amt2,Amt3,Amt4,Amt5,Amt6,UserName)" _
                                & " select 'Input',iif(tblAddLess.Vtype = 'PO','Other Purch',iif(isnull(tblAddLess.ADTMP5),'Trade Purch',iif(tblAddLess.ADTMP5 = 'CLO','Cloth Purch',iif(tblAddLess.ADTMP5 = 'COT','Cotton Purch','XXXX')))),sum(tblVouSub.VSubTaxableAmt),sum(tblVouSub.VsubIGSTAmt),sum(tblVouSub.VsubCGSTAmt),sum(tblVouSub.VsubSGSTAmt),sum(iif(tblVouSub.VSubCtrNo = 1,tblAddLess.AdBillAmt,0)),'" & gUserName & "'" _
                                & " From tblAddLess,tblVouSub where tblAddLess.Vno = tblVouSub.Vno and tblAddLess.Vtype = tblVouSub.Vtype and tblAddLess.Vyear = tblVouSub.Vyear and " & X1 & "" _
                                & " group by iif(tblAddLess.Vtype = 'PO','Other Purch',iif(isnull(tblAddLess.ADTMP5),'Trade Purch',iif(tblAddLess.ADTMP5 = 'CLO','Cloth Purch',iif(tblAddLess.ADTMP5 = 'COT','Cotton Purch','XXXX'))))"
                Else
                    gCn.Execute " Insert Into TmpGentbl (Nar1,vyear,Amt1,Nar2,Nar3,Amt2,Amt3,Amt4,Amt5,Amt6,UserName)" _
                                & " select 'Input',year(tblAddLess.Vdt),month(tblAddLess.Vdt), format(tblAddLess.Vdt,'mmm-yyyy'),iif(tblAddLess.Vtype = 'PO','Other Purch',iif(isnull(tblAddLess.ADTMP5),'Trade Purch',iif(tblAddLess.ADTMP5 = 'CLO','Cloth Purch',iif(tblAddLess.ADTMP5 = 'COT','Cotton Purch','XXXX')))),sum(tblVouSub.VSubTaxableAmt),sum(tblVouSub.VsubIGSTAmt),sum(tblVouSub.VsubCGSTAmt),sum(tblVouSub.VsubSGSTAmt),sum(iif(tblVouSub.VSubCtrNo = 1,tblAddLess.AdBillAmt,0)),'" & gUserName & "'" _
                                & " From tblAddLess,tblVouSub where tblAddLess.Vno = tblVouSub.Vno and tblAddLess.Vtype = tblVouSub.Vtype and tblAddLess.Vyear = tblVouSub.Vyear and " & X1 & "" _
                                & " group by year(tblAddLess.Vdt),month(tblAddLess.Vdt), format(tblAddLess.Vdt,'mmm-yyyy'),iif(tblAddLess.Vtype = 'PO','Other Purch',iif(isnull(tblAddLess.ADTMP5),'Trade Purch',iif(tblAddLess.ADTMP5 = 'CLO','Cloth Purch',iif(tblAddLess.ADTMP5 = 'COT','Cotton Purch','XXXX'))))"
                End If
             End If
             
             '--- Credit Note Purch
             X1 = " qrytblVoucherSr1.VType = 'PN' and qrytblVoucherSr1.VDRCRNTTYPE = 'S' and qrytblVoucherSr1.vdt>=cdate('" & Format(w_Fr, "yyyy,mm,dd") & "')  and qrytblVoucherSr1.vdt<=Cdate('" & Format(w_To, "yyyy,mm,dd") & "') and (qrytblVoucherSr1.VIGSTAmt+qrytblVoucherSr1.VCGSTAmt+qrytblVoucherSr1.VSGSTAmt) > 0 "  '
             If SplR = True Then
                gCn.Execute " Insert Into TmpGentbl (Nar1,Nar3,Amt2,Amt3,Amt4,Amt5,Amt6,UserName)" _
                            & " Select 'Input','Credit Note (Purch)',sum(qrytblVoucherSr1.VTaxableAmt),sum(qrytblVoucherSr1.VIGSTAmt),sum(qrytblVoucherSr1.VCGSTAmt),sum(qrytblVoucherSr1.VSGSTAmt),sum(abs(qrytblVoucherSr1.VAmt))," _
                            & " '" & gUserName & "' from qrytblVoucherSr1 where " & X1 & " " _
                            & " group by 'Debit Note (Sales)'"
             Else
                gCn.Execute " Insert Into TmpGentbl (Nar1,vyear,Amt1,Nar2,Nar3,Amt2,Amt3,Amt4,Amt5,Amt6,UserName)" _
                            & " Select 'Input',year(qrytblVoucherSr1.vdt),month(qrytblVoucherSr1.Vdt), format(qrytblVoucherSr1.Vdt,'mmm-yyyy'),'Credit Note (Purch)',sum(qrytblVoucherSr1.VTaxableAmt),sum(qrytblVoucherSr1.VIGSTAmt),sum(qrytblVoucherSr1.VCGSTAmt),sum(qrytblVoucherSr1.VSGSTAmt),sum(abs(qrytblVoucherSr1.VAmt))," _
                            & " '" & gUserName & "' from qrytblVoucherSr1 where " & X1 & " " _
                            & " group by year(qrytblVoucherSr1.vdt),month(qrytblVoucherSr1.Vdt), format(qrytblVoucherSr1.Vdt,'mmm-yyyy'),'Debit Note (Sales)'"
             End If
             
             '--- Debit Note Purch
             X1 = " qrytblVoucherSr1.VType = 'SN' and qrytblVoucherSr1.VDRCRNTTYPE = 'S' and qrytblVoucherSr1.vdt>=cdate('" & Format(w_Fr, "yyyy,mm,dd") & "')  and qrytblVoucherSr1.vdt<=Cdate('" & Format(w_To, "yyyy,mm,dd") & "') and (qrytblVoucherSr1.VIGSTAmt+qrytblVoucherSr1.VCGSTAmt+qrytblVoucherSr1.VSGSTAmt) > 0 "  '
             If SplR = True Then
                gCn.Execute " Insert Into TmpGentbl (Nar1,Nar3,Amt2,Amt3,Amt4,Amt5,Amt6,UserName)" _
                            & " Select 'Input','Debit Note (Purch)',sum(qrytblVoucherSr1.VTaxableAmt)*-1,sum(qrytblVoucherSr1.VIGSTAmt)*-1,sum(qrytblVoucherSr1.VCGSTAmt)*-1,sum(qrytblVoucherSr1.VSGSTAmt)*-1,sum(abs(qrytblVoucherSr1.VAmt))*-1," _
                            & " '" & gUserName & "' from qrytblVoucherSr1 where " & X1 & " " _
                            & " group by 'Debit Note (Sales)'"
             Else
                gCn.Execute " Insert Into TmpGentbl (Nar1,vyear,Amt1,Nar2,Nar3,Amt2,Amt3,Amt4,Amt5,Amt6,UserName)" _
                            & " Select 'Input',year(qrytblVoucherSr1.vdt),month(qrytblVoucherSr1.Vdt), format(qrytblVoucherSr1.Vdt,'mmm-yyyy'),'Debit Note (Purch)',sum(qrytblVoucherSr1.VTaxableAmt)*-1,sum(qrytblVoucherSr1.VIGSTAmt)*-1,sum(qrytblVoucherSr1.VCGSTAmt)*-1,sum(qrytblVoucherSr1.VSGSTAmt)*-1,sum(abs(qrytblVoucherSr1.VAmt))*-1," _
                            & " '" & gUserName & "' from qrytblVoucherSr1 where " & X1 & " " _
                            & " group by year(qrytblVoucherSr1.vdt),month(qrytblVoucherSr1.Vdt), format(qrytblVoucherSr1.Vdt,'mmm-yyyy'),'Debit Note (Sales)'"
             End If
             '--- Purch Return
             X1 = " tblAddLess.VType = 'VY' and tblAddLess.Vdt>=cdate('" & Format(w_Fr, "yyyy,mm,dd") & "')  and tblAddLess.Vdt<=Cdate('" & Format(w_To, "yyyy,mm,dd") & "') and tblVouSub.VsubIGSTAmt+tblVouSub.VsubCGSTAmt+tblVouSub.VsubSGSTAmt > 0 "  '
             If SplR = True Then
                gCn.Execute " Insert Into TmpGentbl (Nar1,Nar3,Amt2,Amt3,Amt4,Amt5,Amt6,UserName)" _
                            & " select 'Input','Purch Rtn',iif(isnull(sum(tblVouSub.VSubTaxableAmt)),0,sum(tblVouSub.VSubTaxableAmt))*-1,iif(isnull(sum(tblVouSub.VsubIGSTAmt)),0,sum(tblVouSub.VsubIGSTAmt))*-1,iif(isnull(sum(tblVouSub.VsubCGSTAmt)),0,sum(tblVouSub.VsubCGSTAmt))*-1,iif(isnull(sum(tblVouSub.VsubSGSTAmt)),0,sum(tblVouSub.VsubSGSTAmt))*-1,iif(isnull(sum(tblAddLess.AdBillAmt)),0,sum(iif(tblVouSub.VSubCtrNo = 1,tblAddLess.AdBillAmt,0)))*-1,'" & gUserName & "'" _
                            & " From tblAddLess,tblVouSub where tblAddLess.Vno = tblVouSub.Vno and tblAddLess.Vtype = tblVouSub.Vtype and tblAddLess.Vyear = tblVouSub.Vyear and " & X1 & ""
             Else
                gCn.Execute " Insert Into TmpGentbl (Nar1,vyear,Amt1,Nar2,Nar3,Amt2,Amt3,Amt4,Amt5,Amt6,UserName)" _
                            & " select 'Input',year(tblAddLess.Vdt),month(tblAddLess.Vdt), format(tblAddLess.Vdt,'mmm-yyyy'),'Purch Rtn',sum(tblVouSub.VSubTaxableAmt)*-1,sum(tblVouSub.VsubIGSTAmt)*-1,sum(tblVouSub.VsubCGSTAmt)*-1,sum(tblVouSub.VsubSGSTAmt)*-1,sum(iif(tblVouSub.VSubCtrNo = 1,tblAddLess.AdBillAmt,0))*-1,'" & gUserName & "'" _
                            & " From tblAddLess,tblVouSub where tblAddLess.Vno = tblVouSub.Vno and tblAddLess.Vtype = tblVouSub.Vtype and tblAddLess.Vyear = tblVouSub.Vyear and " & X1 & "" _
                            & " group by year(tblAddLess.Vdt),month(tblAddLess.Vdt), format(tblAddLess.Vdt,'mmm-yyyy')"
             End If
             
             
             
             X1 = " tblAddLess.VType = 'PO' and tblAddLess.AdBookDt>=cdate('" & Format(w_Fr, "yyyy,mm,dd") & "')  and tblAddLess.AdBookDt<=Cdate('" & Format(w_To, "yyyy,mm,dd") & "') and tblAddLess.AdLCostAdd+tblAddLess.AdLCostFreight+tblAddLess.AdLCostGRent > 0 "  '
'             If SplR = True Then
'                gCn.Execute " Insert Into TmpGentbl (Nar1,Nar3,Amt2,Amt3,Amt4,Amt5,Amt6,UserName)" _
'                            & " select 'Input','RCM',iif(isnull(sum(tblAddLess.AdBillAmt)),0,sum(tblAddLess.AdBillAmt)),iif(isnull(sum(tblAddLess.AdLCostAdd)),0,sum(tblAddLess.AdLCostAdd)),iif(isnull(sum(tblAddLess.AdLCostFreight)),0,sum(tblAddLess.AdLCostFreight)),iif(isnull(sum(tblAddLess.AdLCostGRent)),0, sum(tblAddLess.AdLCostGRent)),iif(isnull(sum(tblAddLess.AdBillAmt)),0, sum(tblAddLess.AdBillAmt)),'" & gUserName & "'" _
'                            & " From tblAddLess where " & X1 & ""
'             Else
'                gCn.Execute " Insert Into TmpGentbl (Nar1,vyear,Amt1,Nar2,Nar3,Amt2,Amt3,Amt4,Amt5,Amt6,UserName)" _
'                            & " select 'Input',year(tblAddLess.AdBookDt),month(tblAddLess.AdBookDt), format(tblAddLess.AdBookDt,'mmm-yyyy'),'RCM',sum(tblAddLess.AdBillAmt),sum(tblAddLess.AdLCostAdd),sum(tblAddLess.AdLCostFreight),sum(tblAddLess.AdLCostGRent),sum(tblAddLess.AdBillAmt),'" & gUserName & "'" _
'                            & " From tblAddLess where " & X1 & "" _
'                            & " group by year(tblAddLess.AdBookDt),month(tblAddLess.AdBookDt), format(tblAddLess.AdBookDt,'mmm-yyyy')"
'             End If
             If SplR = True Then
                gCn.Execute " Insert Into TmpGentbl (Nar5,Nar1,Nar3,Amt2,Amt3,Amt4,Amt5,Amt6,UserName)" _
                            & " select 'B','RCM Pay','RCM',iif(isnull(sum(tblAddLess.AdBillAmt)),0,sum(tblAddLess.AdBillAmt)),iif(isnull(sum(tblAddLess.AdLCostAdd)),0,sum(tblAddLess.AdLCostAdd)),iif(isnull(sum(tblAddLess.AdLCostFreight)),0,sum(tblAddLess.AdLCostFreight)),iif(isnull(sum(tblAddLess.AdLCostGRent)),0,   sum(tblAddLess.AdLCostGRent)),iif(isnull(sum(tblAddLess.AdBillAmt)),0, sum(tblAddLess.AdBillAmt)),'" & gUserName & "'" _
                            & " From tblAddLess where " & X1 & ""
             Else
                gCn.Execute " Insert Into TmpGentbl (Nar5,Nar1,vyear,Amt1,Nar2,Nar3,Amt2,Amt3,Amt4,Amt5,Amt6,UserName)" _
                            & " select 'B','RCM Pay',year(tblAddLess.AdBookDt),month(tblAddLess.AdBookDt), format(tblAddLess.AdBookDt,'mmm-yyyy'),'RCM',sum(tblAddLess.AdBillAmt),sum(tblAddLess.AdLCostAdd),sum(tblAddLess.AdLCostFreight),sum(tblAddLess.AdLCostGRent),sum(tblAddLess.AdBillAmt),'" & gUserName & "'" _
                            & " From tblAddLess where " & X1 & "" _
                            & " group by year(tblAddLess.AdBookDt),month(tblAddLess.AdBookDt), format(tblAddLess.AdBookDt,'mmm-yyyy')"
             End If
             gCn.CommitTrans

End Function

''Public Function GProcPurpareGSTSummery(w_Fr As Date, w_To As Date, Optional IsOnPurBillDt As Boolean)
''Dim X1 As String
''             gCn.BeginTrans
''             gCn.Execute "delete from tmpGenTbl where username='" & gUserName & "'"
''             '---Sales Detail
''             X1 = " tblAddLess.VType in ('SY')  and tblAddLess.vdt>=cdate('" & Format(w_Fr, "yyyy,mm,dd") & "')  and tblAddLess.vdt<=Cdate('" & Format(w_To, "yyyy,mm,dd") & "') and tblVouSub.VsubIGSTAmt+tblVouSub.VsubCGSTAmt+tblVouSub.VsubSGSTAmt > 0 "  '
''             gCn.Execute " Insert Into TmpGentbl (Nar1,vyear,Amt1,Nar2,Nar3,Amt2,Amt3,Amt4,Amt5,Amt6,UserName)" _
''                         & " select 'Output',year(tblAddLess.vdt),month(tblAddLess.Vdt), format(tblAddLess.Vdt,'mmm-yyyy'),iif(isnull(tblAddLess.ADTMP5 ),'Trade Sales',iif(tblAddLess.ADTMP5= 'COL','Cloth Sales',iif(tblAddLess.ADTMP5= 'OTH','Other Sales',iif(tblAddLess.ADTMP5= 'COT','Cotton','XXX')))),sum(tblVouSub.VSubAmt),sum(tblVouSub.VsubIGSTAmt),sum(tblVouSub.VsubCGSTAmt),sum(tblVouSub.VsubSGSTAmt),sum(iif(tblVouSub.VSubCtrNo = 1,tblAddLess.AdBillAmt,0)),'" & gUserName & "'" _
''                         & " From tblAddLess,tblVouSub where tblAddLess.Vno = tblVouSub.Vno and tblAddLess.Vtype = tblVouSub.Vtype and tblAddLess.Vyear = tblVouSub.Vyear and " & X1 & "" _
''                         & " group by year(tblAddLess.vdt),month(tblAddLess.Vdt), format(tblAddLess.Vdt,'mmm-yyyy'),iif(isnull(tblAddLess.ADTMP5 ),'Trade Sales',iif(tblAddLess.ADTMP5= 'COL','Cloth Sales',iif(tblAddLess.ADTMP5= 'OTH','Other Sales',iif(tblAddLess.ADTMP5= 'COT','Cotton','XXX'))))"
''
'''             X1 = " qrytblVoucherSr1.VType in ('SN','PN') and qrytblVoucherSr1.VDRCRNTTYPE = 'B' and qrytblVoucherSr1.vdt>=cdate('" & Format(w_Fr, "yyyy,mm,dd") & "')  and qrytblVoucherSr1.vdt<=Cdate('" & Format(w_To, "yyyy,mm,dd") & "') and (qrytblVoucherSr1.VIGSTAmt+qrytblVoucherSr1.VCGSTAmt+qrytblVoucherSr1.VSGSTAmt) > 0 "  '
''
''             '--- Debit Note Sales
''             X1 = " qrytblVoucherSr1.VType = 'SN' and qrytblVoucherSr1.VDRCRNTTYPE = 'B' and qrytblVoucherSr1.vdt>=cdate('" & Format(w_Fr, "yyyy,mm,dd") & "')  and qrytblVoucherSr1.vdt<=Cdate('" & Format(w_To, "yyyy,mm,dd") & "') and (qrytblVoucherSr1.VIGSTAmt+qrytblVoucherSr1.VCGSTAmt+qrytblVoucherSr1.VSGSTAmt) > 0 "  '
''             gCn.Execute " Insert Into TmpGentbl (Nar1,vyear,Amt1,Nar2,Nar3,Amt2,Amt3,Amt4,Amt5,Amt6,UserName)" _
''                         & " Select 'Output',year(qrytblVoucherSr1.vdt),month(qrytblVoucherSr1.Vdt), format(qrytblVoucherSr1.Vdt,'mmm-yyyy'),'Debit Note (Sales)',sum(qrytblVoucherSr1.VTaxableAmt),sum(qrytblVoucherSr1.VIGSTAmt),sum(qrytblVoucherSr1.VCGSTAmt),sum(qrytblVoucherSr1.VSGSTAmt),sum(abs(qrytblVoucherSr1.VAmt))," _
''                         & " '" & gUserName & "' from qrytblVoucherSr1 where " & X1 & " " _
''                         & " group by year(qrytblVoucherSr1.vdt),month(qrytblVoucherSr1.Vdt), format(qrytblVoucherSr1.Vdt,'mmm-yyyy'),'Debit Note (Sales)'"
''
''
''             '--- Sales Return
''             X1 = " tblAddLess.VType = 'RY' and tblAddLess.Vdt>=cdate('" & Format(w_Fr, "yyyy,mm,dd") & "')  and tblAddLess.Vdt<=Cdate('" & Format(w_To, "yyyy,mm,dd") & "') and tblVouSub.VsubIGSTAmt+tblVouSub.VsubCGSTAmt+tblVouSub.VsubSGSTAmt > 0 "  '
''
''             gCn.Execute " Insert Into TmpGentbl (Nar1,vyear,Amt1,Nar2,Nar3,Amt2,Amt3,Amt4,Amt5,Amt6,UserName)" _
''                         & " select 'Output',year(tblAddLess.Vdt),month(tblAddLess.Vdt), format(tblAddLess.Vdt,'mmm-yyyy'),'Sales Rtn',sum(tblVouSub.VSubAmt)*-1,sum(tblVouSub.VsubIGSTAmt)*-1,sum(tblVouSub.VsubCGSTAmt)*-1,sum(tblVouSub.VsubSGSTAmt)*-1,sum(iif(tblVouSub.VSubCtrNo = 1,tblAddLess.AdBillAmt,0))*-1,'" & gUserName & "'" _
''                         & " From tblAddLess,tblVouSub where tblAddLess.Vno = tblVouSub.Vno and tblAddLess.Vtype = tblVouSub.Vtype and tblAddLess.Vyear = tblVouSub.Vyear and " & X1 & "" _
''                         & " group by year(tblAddLess.Vdt),month(tblAddLess.Vdt), format(tblAddLess.Vdt,'mmm-yyyy')"
''
''             '--- Credit Note Sales
''             X1 = " qrytblVoucherSr1.VType = 'PN' and qrytblVoucherSr1.VDRCRNTTYPE = 'B' and qrytblVoucherSr1.vdt>=cdate('" & Format(w_Fr, "yyyy,mm,dd") & "')  and qrytblVoucherSr1.vdt<=Cdate('" & Format(w_To, "yyyy,mm,dd") & "') and (qrytblVoucherSr1.VIGSTAmt+qrytblVoucherSr1.VCGSTAmt+qrytblVoucherSr1.VSGSTAmt) > 0 "  '
''             gCn.Execute " Insert Into TmpGentbl (Nar1,vyear,Amt1,Nar2,Nar3,Amt2,Amt3,Amt4,Amt5,Amt6,UserName)" _
''                         & " Select 'Output',year(qrytblVoucherSr1.vdt),month(qrytblVoucherSr1.Vdt), format(qrytblVoucherSr1.Vdt,'mmm-yyyy'),'Credit Note (Sales)',sum(qrytblVoucherSr1.VTaxableAmt)*-1,sum(qrytblVoucherSr1.VIGSTAmt)*-1,sum(qrytblVoucherSr1.VCGSTAmt)*-1,sum(qrytblVoucherSr1.VSGSTAmt)*-1,sum(abs(qrytblVoucherSr1.VAmt))*-1," _
''                         & " '" & gUserName & "' from qrytblVoucherSr1 where " & X1 & " " _
''                         & " group by year(qrytblVoucherSr1.vdt),month(qrytblVoucherSr1.Vdt), format(qrytblVoucherSr1.Vdt,'mmm-yyyy'),'Debit Note (Sales)'"
''
''            '--- Purch & Other Purch
''             If IsOnPurBillDt = True Then
''                X1 = " tblAddLess.VType ='PY' and tblAddLess.AdGpDt>=cdate('" & Format(w_Fr, "yyyy,mm,dd") & "')  and tblAddLess.AdGpDt<=Cdate('" & Format(w_To, "yyyy,mm,dd") & "') and tblVouSub.VsubIGSTAmt+tblVouSub.VsubCGSTAmt+tblVouSub.VsubSGSTAmt > 0 "
''                gCn.Execute " Insert Into TmpGentbl (Nar1,vyear,Amt1,Nar2,Nar3,Amt2,Amt3,Amt4,Amt5,Amt6,UserName)" _
''                            & " select 'Input',year(tblAddLess.AdGpDt),month(tblAddLess.AdGpDt), format(tblAddLess.AdGpDt,'mmm-yyyy'),iif(tblAddLess.Vtype = 'PO','Other Purch',iif(isnull(tblAddLess.ADTMP5),'Trade Purch',iif(tblAddLess.ADTMP5 = 'CLO','Cloth Purch',iif(tblAddLess.ADTMP5 = 'COT','Cotton Purch','XXXX')))),sum(tblVouSub.VSubTaxableAmt),sum(tblVouSub.VsubIGSTAmt),sum(tblVouSub.VsubCGSTAmt),sum(tblVouSub.VsubSGSTAmt),sum(iif(tblVouSub.VSubCtrNo = 1,tblAddLess.AdBillAmt,0)),'" & gUserName & "'" _
''                            & " From tblAddLess,tblVouSub where tblAddLess.Vno = tblVouSub.Vno and tblAddLess.Vtype = tblVouSub.Vtype and tblAddLess.Vyear = tblVouSub.Vyear and " & X1 & "" _
''                            & " group by year(tblAddLess.AdGpDt),month(tblAddLess.AdGpDt), format(tblAddLess.AdGpDt,'mmm-yyyy'),iif(tblAddLess.Vtype = 'PO','Other Purch',iif(isnull(tblAddLess.ADTMP5),'Trade Purch',iif(tblAddLess.ADTMP5 = 'CLO','Cloth Purch',iif(tblAddLess.ADTMP5 = 'COT','Cotton Purch','XXXX'))))"
''
''                X1 = " tblAddLess.VType ='PO' and tblAddLess.AdGpDt>=cdate('" & Format(w_Fr, "yyyy,mm,dd") & "')  and tblAddLess.AdGpDt<=Cdate('" & Format(w_To, "yyyy,mm,dd") & "') and tblVouSub.VsubIGSTAmt+tblVouSub.VsubCGSTAmt+tblVouSub.VsubSGSTAmt > 0 and tblAddLess.AdLCostAdd+tblAddLess.AdLCostFreight+tblAddLess.AdLCostGRent = 0"
''                gCn.Execute " Insert Into TmpGentbl (Nar1,vyear,Amt1,Nar2,Nar3,Amt2,Amt3,Amt4,Amt5,Amt6,UserName)" _
''                            & " select 'Input',year(tblAddLess.AdGpDt),month(tblAddLess.AdGpDt), format(tblAddLess.AdGpDt,'mmm-yyyy'),iif(tblAddLess.Vtype = 'PO','Other Purch',iif(isnull(tblAddLess.ADTMP5),'Trade Purch',iif(tblAddLess.ADTMP5 = 'CLO','Cloth Purch',iif(tblAddLess.ADTMP5 = 'COT','Cotton Purch','XXXX')))),sum(tblVouSub.VSubTaxableAmt),sum(tblVouSub.VsubIGSTAmt),sum(tblVouSub.VsubCGSTAmt),sum(tblVouSub.VsubSGSTAmt),sum(iif(tblVouSub.VSubCtrNo = 1,tblAddLess.AdBillAmt,0)),'" & gUserName & "'" _
''                            & " From tblAddLess,tblVouSub where tblAddLess.Vno = tblVouSub.Vno and tblAddLess.Vtype = tblVouSub.Vtype and tblAddLess.Vyear = tblVouSub.Vyear and " & X1 & "" _
''                            & " group by year(tblAddLess.AdGpDt),month(tblAddLess.AdGpDt), format(tblAddLess.AdGpDt,'mmm-yyyy'),iif(tblAddLess.Vtype = 'PO','Other Purch',iif(isnull(tblAddLess.ADTMP5),'Trade Purch',iif(tblAddLess.ADTMP5 = 'CLO','Cloth Purch',iif(tblAddLess.ADTMP5 = 'COT','Cotton Purch','XXXX'))))"
''             Else
''                X1 = " tblAddLess.VType ='PY' and tblAddLess.Vdt>=cdate('" & Format(w_Fr, "yyyy,mm,dd") & "')  and tblAddLess.Vdt<=Cdate('" & Format(w_To, "yyyy,mm,dd") & "') and tblVouSub.VsubIGSTAmt+tblVouSub.VsubCGSTAmt+tblVouSub.VsubSGSTAmt > 0 "  '
''                gCn.Execute " Insert Into TmpGentbl (Nar1,vyear,Amt1,Nar2,Nar3,Amt2,Amt3,Amt4,Amt5,Amt6,UserName)" _
''                            & " select 'Input',year(tblAddLess.Vdt),month(tblAddLess.Vdt), format(tblAddLess.Vdt,'mmm-yyyy'),iif(tblAddLess.Vtype = 'PO','Other Purch',iif(isnull(tblAddLess.ADTMP5),'Trade Purch',iif(tblAddLess.ADTMP5 = 'CLO','Cloth Purch',iif(tblAddLess.ADTMP5 = 'COT','Cotton Purch','XXXX')))),sum(tblVouSub.VSubTaxableAmt),sum(tblVouSub.VsubIGSTAmt),sum(tblVouSub.VsubCGSTAmt),sum(tblVouSub.VsubSGSTAmt),sum(iif(tblVouSub.VSubCtrNo = 1,tblAddLess.AdBillAmt,0)),'" & gUserName & "'" _
''                            & " From tblAddLess,tblVouSub where tblAddLess.Vno = tblVouSub.Vno and tblAddLess.Vtype = tblVouSub.Vtype and tblAddLess.Vyear = tblVouSub.Vyear and " & X1 & "" _
''                            & " group by year(tblAddLess.Vdt),month(tblAddLess.Vdt), format(tblAddLess.Vdt,'mmm-yyyy'),iif(tblAddLess.Vtype = 'PO','Other Purch',iif(isnull(tblAddLess.ADTMP5),'Trade Purch',iif(tblAddLess.ADTMP5 = 'CLO','Cloth Purch',iif(tblAddLess.ADTMP5 = 'COT','Cotton Purch','XXXX'))))"
''
''                X1 = " tblAddLess.VType ='PO' and tblAddLess.Vdt>=cdate('" & Format(w_Fr, "yyyy,mm,dd") & "')  and tblAddLess.Vdt<=Cdate('" & Format(w_To, "yyyy,mm,dd") & "') and tblVouSub.VsubIGSTAmt+tblVouSub.VsubCGSTAmt+tblVouSub.VsubSGSTAmt > 0 and tblAddLess.AdLCostAdd+tblAddLess.AdLCostFreight+tblAddLess.AdLCostGRent = 0"  '
''                gCn.Execute " Insert Into TmpGentbl (Nar1,vyear,Amt1,Nar2,Nar3,Amt2,Amt3,Amt4,Amt5,Amt6,UserName)" _
''                            & " select 'Input',year(tblAddLess.Vdt),month(tblAddLess.Vdt), format(tblAddLess.Vdt,'mmm-yyyy'),iif(tblAddLess.Vtype = 'PO','Other Purch',iif(isnull(tblAddLess.ADTMP5),'Trade Purch',iif(tblAddLess.ADTMP5 = 'CLO','Cloth Purch',iif(tblAddLess.ADTMP5 = 'COT','Cotton Purch','XXXX')))),sum(tblVouSub.VSubTaxableAmt),sum(tblVouSub.VsubIGSTAmt),sum(tblVouSub.VsubCGSTAmt),sum(tblVouSub.VsubSGSTAmt),sum(iif(tblVouSub.VSubCtrNo = 1,tblAddLess.AdBillAmt,0)),'" & gUserName & "'" _
''                            & " From tblAddLess,tblVouSub where tblAddLess.Vno = tblVouSub.Vno and tblAddLess.Vtype = tblVouSub.Vtype and tblAddLess.Vyear = tblVouSub.Vyear and " & X1 & "" _
''                            & " group by year(tblAddLess.Vdt),month(tblAddLess.Vdt), format(tblAddLess.Vdt,'mmm-yyyy'),iif(tblAddLess.Vtype = 'PO','Other Purch',iif(isnull(tblAddLess.ADTMP5),'Trade Purch',iif(tblAddLess.ADTMP5 = 'CLO','Cloth Purch',iif(tblAddLess.ADTMP5 = 'COT','Cotton Purch','XXXX'))))"
''             End If
''
''             '--- Purch Return
''             X1 = " tblAddLess.VType = 'VY' and tblAddLess.Vdt>=cdate('" & Format(w_Fr, "yyyy,mm,dd") & "')  and tblAddLess.Vdt<=Cdate('" & Format(w_To, "yyyy,mm,dd") & "') and tblVouSub.VsubIGSTAmt+tblVouSub.VsubCGSTAmt+tblVouSub.VsubSGSTAmt > 0 "  '
''
''             gCn.Execute " Insert Into TmpGentbl (Nar1,vyear,Amt1,Nar2,Nar3,Amt2,Amt3,Amt4,Amt5,Amt6,UserName)" _
''                         & " select 'Input',year(tblAddLess.Vdt),month(tblAddLess.Vdt), format(tblAddLess.Vdt,'mmm-yyyy'),'Purch Rtn',sum(tblVouSub.VSubTaxableAmt)*-1,sum(tblVouSub.VsubIGSTAmt)*-1,sum(tblVouSub.VsubCGSTAmt)*-1,sum(tblVouSub.VsubSGSTAmt)*-1,sum(iif(tblVouSub.VSubCtrNo = 1,tblAddLess.AdBillAmt,0))*-1,'" & gUserName & "'" _
''                         & " From tblAddLess,tblVouSub where tblAddLess.Vno = tblVouSub.Vno and tblAddLess.Vtype = tblVouSub.Vtype and tblAddLess.Vyear = tblVouSub.Vyear and " & X1 & "" _
''                         & " group by year(tblAddLess.Vdt),month(tblAddLess.Vdt), format(tblAddLess.Vdt,'mmm-yyyy')"
''
''
''
''
'''             If IsOnPurBillDt = True Then
'''                X1 = " tblAddLess.VType in ('PY','PO') and tblAddLess.AdGpDt>=cdate('" & Format(w_Fr, "yyyy,mm,dd") & "')  and tblAddLess.AdGpDt<=Cdate('" & Format(w_To, "yyyy,mm,dd") & "') and tblVouSub.VsubIGSTAmt+tblVouSub.VsubCGSTAmt+tblVouSub.VsubSGSTAmt > 0 and tblAddLess.AdLCostAdd+tblAddLess.AdLCostFreight+tblAddLess.AdLCostGRent = 0 "  '
'''                gCn.Execute " Insert Into TmpGentbl (Nar1,vyear,Amt1,Nar2,Nar3,Amt2,Amt3,Amt4,Amt5,Amt6,UserName)" _
'''                            & " select 'Input',year(tblAddLess.AdGpDt),month(tblAddLess.AdGpDt), format(tblAddLess.AdGpDt,'mmm-yyyy'),iif(tblAddLess.Vtype = 'PO','Other Purch',iif(isnull(tblAddLess.ADTMP5),'Trade Purch',iif(tblAddLess.ADTMP5 = 'CLO','Cloth Purch',iif(tblAddLess.ADTMP5 = 'COT','Cotton Purch','XXXX')))),sum(tblVouSub.VSubAmt),sum(tblVouSub.VsubIGSTAmt),sum(tblVouSub.VsubCGSTAmt),sum(tblVouSub.VsubSGSTAmt),sum(iif(tblVouSub.VSubCtrNo = 1,tblAddLess.AdBillAmt,0)),'" & gUserName & "'" _
'''                            & " From tblAddLess,tblVouSub where tblAddLess.Vno = tblVouSub.Vno and tblAddLess.Vtype = tblVouSub.Vtype and tblAddLess.Vyear = tblVouSub.Vyear and " & X1 & "" _
'''                            & " group by year(tblAddLess.AdGpDt),month(tblAddLess.AdGpDt), format(tblAddLess.AdGpDt,'mmm-yyyy'),iif(tblAddLess.Vtype = 'PO','Other Purch',iif(isnull(tblAddLess.ADTMP5),'Trade Purch',iif(tblAddLess.ADTMP5 = 'CLO','Cloth Purch',iif(tblAddLess.ADTMP5 = 'COT','Cotton Purch','XXXX'))))"
'''             Else
'''                X1 = " tblAddLess.VType in ('PY','PO') and tblAddLess.Vdt>=cdate('" & Format(w_Fr, "yyyy,mm,dd") & "')  and tblAddLess.Vdt<=Cdate('" & Format(w_To, "yyyy,mm,dd") & "') and tblVouSub.VsubIGSTAmt+tblVouSub.VsubCGSTAmt+tblVouSub.VsubSGSTAmt > 0 "  '
'''                gCn.Execute " Insert Into TmpGentbl (Nar1,vyear,Amt1,Nar2,Nar3,Amt2,Amt3,Amt4,Amt5,Amt6,UserName)" _
'''                            & " select 'Input',year(tblAddLess.Vdt),month(tblAddLess.Vdt), format(tblAddLess.Vdt,'mmm-yyyy'),iif(tblAddLess.Vtype = 'PO','Other Purch',iif(isnull(tblAddLess.ADTMP5),'Trade Purch',iif(tblAddLess.ADTMP5 = 'CLO','Cloth Purch',iif(tblAddLess.ADTMP5 = 'COT','Cotton Purch','XXXX')))),sum(tblVouSub.VSubAmt),sum(tblVouSub.VsubIGSTAmt),sum(tblVouSub.VsubCGSTAmt),sum(tblVouSub.VsubSGSTAmt),sum(iif(tblVouSub.VSubCtrNo = 1,tblAddLess.AdBillAmt,0)),'" & gUserName & "'" _
'''                            & " From tblAddLess,tblVouSub where tblAddLess.Vno = tblVouSub.Vno and tblAddLess.Vtype = tblVouSub.Vtype and tblAddLess.Vyear = tblVouSub.Vyear and " & X1 & "" _
'''                            & " group by year(tblAddLess.Vdt),month(tblAddLess.Vdt), format(tblAddLess.Vdt,'mmm-yyyy'),iif(tblAddLess.Vtype = 'PO','Other Purch',iif(isnull(tblAddLess.ADTMP5),'Trade Purch',iif(tblAddLess.ADTMP5 = 'CLO','Cloth Purch',iif(tblAddLess.ADTMP5 = 'COT','Cotton Purch','XXXX'))))"
'''             End If
''
''
''             '--- Credit Note Purch
''             X1 = " qrytblVoucherSr1.VType = 'PN' and qrytblVoucherSr1.VDRCRNTTYPE = 'S' and qrytblVoucherSr1.vdt>=cdate('" & Format(w_Fr, "yyyy,mm,dd") & "')  and qrytblVoucherSr1.vdt<=Cdate('" & Format(w_To, "yyyy,mm,dd") & "') and (qrytblVoucherSr1.VIGSTAmt+qrytblVoucherSr1.VCGSTAmt+qrytblVoucherSr1.VSGSTAmt) > 0 "  '
''             gCn.Execute " Insert Into TmpGentbl (Nar1,vyear,Amt1,Nar2,Nar3,Amt2,Amt3,Amt4,Amt5,Amt6,UserName)" _
''                         & " Select 'Input',year(qrytblVoucherSr1.vdt),month(qrytblVoucherSr1.Vdt), format(qrytblVoucherSr1.Vdt,'mmm-yyyy'),'Credit Note (Purch)',sum(qrytblVoucherSr1.VTaxableAmt),sum(qrytblVoucherSr1.VIGSTAmt),sum(qrytblVoucherSr1.VCGSTAmt),sum(qrytblVoucherSr1.VSGSTAmt),sum(abs(qrytblVoucherSr1.VAmt))," _
''                         & " '" & gUserName & "' from qrytblVoucherSr1 where " & X1 & " " _
''                         & " group by year(qrytblVoucherSr1.vdt),month(qrytblVoucherSr1.Vdt), format(qrytblVoucherSr1.Vdt,'mmm-yyyy'),'Debit Note (Sales)'"
''
''             '--- Debit Note Purch
''             X1 = " qrytblVoucherSr1.VType = 'SN' and qrytblVoucherSr1.VDRCRNTTYPE = 'S' and qrytblVoucherSr1.vdt>=cdate('" & Format(w_Fr, "yyyy,mm,dd") & "')  and qrytblVoucherSr1.vdt<=Cdate('" & Format(w_To, "yyyy,mm,dd") & "') and (qrytblVoucherSr1.VIGSTAmt+qrytblVoucherSr1.VCGSTAmt+qrytblVoucherSr1.VSGSTAmt) > 0 "  '
''             gCn.Execute " Insert Into TmpGentbl (Nar1,vyear,Amt1,Nar2,Nar3,Amt2,Amt3,Amt4,Amt5,Amt6,UserName)" _
''                         & " Select 'Input',year(qrytblVoucherSr1.vdt),month(qrytblVoucherSr1.Vdt), format(qrytblVoucherSr1.Vdt,'mmm-yyyy'),'Debit Note (Purch)',sum(qrytblVoucherSr1.VTaxableAmt)*-1,sum(qrytblVoucherSr1.VIGSTAmt)*-1,sum(qrytblVoucherSr1.VCGSTAmt)*-1,sum(qrytblVoucherSr1.VSGSTAmt)*-1,sum(abs(qrytblVoucherSr1.VAmt))*-1," _
''                         & " '" & gUserName & "' from qrytblVoucherSr1 where " & X1 & " " _
''                         & " group by year(qrytblVoucherSr1.vdt),month(qrytblVoucherSr1.Vdt), format(qrytblVoucherSr1.Vdt,'mmm-yyyy'),'Debit Note (Sales)'"
''
''
''
'''
''             X1 = " tblAddLess.VType = 'PO' and tblAddLess.AdBookDt>=cdate('" & Format(w_Fr, "yyyy,mm,dd") & "')  and tblAddLess.AdBookDt<=Cdate('" & Format(w_To, "yyyy,mm,dd") & "') and tblAddLess.AdLCostAdd+tblAddLess.AdLCostFreight+tblAddLess.AdLCostGRent > 0 "  '
'''
'''             gCn.Execute " Insert Into TmpGentbl (Nar1,vyear,Amt1,Nar2,Nar3,Amt2,Amt3,Amt4,Amt5,Amt6,UserName)" _
'''                         & " select 'Input',year(tblAddLess.AdBookDt),month(tblAddLess.AdBookDt), format(tblAddLess.AdBookDt,'mmm-yyyy'),'RCM',sum(tblAddLess.AdBillAmt),sum(tblAddLess.AdLCostAdd),sum(tblAddLess.AdLCostFreight),sum(tblAddLess.AdLCostGRent),sum(tblAddLess.AdBillAmt),'" & gUserName & "'" _
'''                         & " From tblAddLess where " & X1 & "" _
'''                         & " group by year(tblAddLess.AdBookDt),month(tblAddLess.AdBookDt), format(tblAddLess.AdBookDt,'mmm-yyyy')"
''
''             gCn.Execute " Insert Into TmpGentbl (Nar5,Nar1,vyear,Amt1,Nar2,Nar3,Amt2,Amt3,Amt4,Amt5,Amt6,UserName)" _
''                         & " select 'B','RCM Pay',year(tblAddLess.AdBookDt),month(tblAddLess.AdBookDt), format(tblAddLess.AdBookDt,'mmm-yyyy'),'RCM',sum(tblAddLess.AdBillAmt),sum(tblAddLess.AdLCostAdd),sum(tblAddLess.AdLCostFreight),sum(tblAddLess.AdLCostGRent),sum(tblAddLess.AdBillAmt),'" & gUserName & "'" _
''                         & " From tblAddLess where " & X1 & "" _
''                         & " group by year(tblAddLess.AdBookDt),month(tblAddLess.AdBookDt), format(tblAddLess.AdBookDt,'mmm-yyyy')"
''
''             gCn.CommitTrans
''
''End Function
'-------------------- Prepare Mill Account Statement ( used for Consignment Depot Mill statement)
Private Sub PrepareAccountStatement(pFrmDt As Date, pToDt As Date, MillName As String, TDSP As Double, ServiceECess As Double, ServiceTax As Double, Surcharge As Double, TDSEcess As Double)
Dim RsTmp As Recordset
Dim X As String
Dim VATTaxCondStr As String
Dim MillCode As Long
Dim Amt As Double
Dim TaxAmt As Double
Dim BrkBag As Double
Dim BrkWt As Double
Dim BrkAmt As Double
Dim TotAmt As Double
Dim TDSAmt As Double
Dim SurchargeAmt As Double
Dim TDSCessAmt As Double
Dim CessAmt As Double
Dim ServiceTaxAmt As Double
Dim CondStr As String
MillCode = GProcGetColumnValue("tblMastAccount", "AcName", MillName, "S", "AcCode", "N")
gCn.BeginTrans
gCn.Execute "Delete from tmpGenTbl where USERNAME='" & gUserName & "'"
'============== Opening Balance
Amt = GProcGetAccountBalance(MillName, pFrmDt - 1)
gCn.Execute "insert into tmpGenTbl (VAcCode,Nar1,UserName,Amt1,Vno,VSrno) values (" & MillCode & "," _
& "'Opening Balance','" & gUserName & "'," & Amt & ",1,0)"
'============== Sales
Amt = 0
If gBackEndDB = gBackEndAccess Then
    CondStr = " Vdt>= cdate('" & pFrmDt & "')  and  Vdt<=cdate('" & pToDt & "')"
Else '-- Oracle
    CondStr = " Vdt>= to_date('" & CDate(pFrmDt) & "','dd/MM/yyyy')  and Vdt<=to_date('" & CDate(pToDt) & "','dd/MM/yyyy')"
End If
X = " select sum(ADBILLAMT) from tblAddLess where VType in ('SO','SD')  and " & CondStr & " and  ADMILLCODE=" & MillCode
If GProcRstOpen(RsTmp, X, "R") > 0 Then
    If IsNumeric(RsTmp.Fields(0)) = True Then
        Amt = CDbl(RsTmp.Fields(0))
    Else
        Amt = 0
    End If
    If Amt > 0 Then gCn.Execute "insert into tmpGenTbl (VAcCode,Nar1,UserName,Amt1,Vno,VSrno) values (" & MillCode & "," _
    & "'Sales','" & gUserName & "'," & Amt & ",2,0)"
End If
'============== Commission,V A T Payable,TDS ,E Cess ,Service tax
Amt = 0
CommAmt = 0
CommOn = 0
CommBag = 0
If gBackEndDB = gBackEndAccess Then
    CondStr = " tblAddLess.Vdt>= cdate('" & pFrmDt & "')  and  tblAddLess.Vdt<=cdate('" & pToDt & "')"
Else '-- Oracle
    CondStr = " tblAddLess.Vdt>= to_date('" & CDate(pFrmDt) & "','dd/MM/yyyy')  and tblAddLess.Vdt<=to_date('" & CDate(pToDt) & "','dd/MM/yyyy')"
End If
If gBackEndDB = gBackEndAccess Then '-- Access
    If gIsNoVATVoucherSD = False And gIsNoVATVoucherSO = False Then
        VATTaxCondStr = "sum(AdTaxAmt)"
    Else
        If gIsNoVATVoucherSD = True Then
            VATTaxCondStr = "sum(iif(tblAddLess.Vtype='SD',iif(NARRATION='V A T',0,AdTaxAmt),AdTaxamt))"
        Else
            If gIsNoVATVoucherSO = True Then
                VATTaxCondStr = "sum(iif(tblAddLess.Vtype='SO',iif(NARRATION='V A T',0,AdTaxAmt),AdTaxamt))"
            Else
                VATTaxCondStr = "sum(iif(NARRATION='V A T',0,AdTaxAmt))"
            End If
        End If
    End If
Else '-- Oracle
     If gIsNoVATVoucherSD = False And gIsNoVATVoucherSO = False Then
        VATTaxCondStr = "sum(AdTaxAmt)"
     Else
        If gIsNoVATVoucherSD = True Then
            VATTaxCondStr = "sum(decode(tblAddLess.Vtype,'SD',decode(NARRATION,'V A T',0,AdTaxAmt),AdTaxamt))"
        Else
            If gIsNoVATVoucherSO = True Then
                VATTaxCondStr = "sum(decode(tblAddLess.Vtype,'SO',decode(NARRATION,'V A T',0,AdTaxAmt),AdTaxamt))"
            Else
                VATTaxCondStr = "sum(decode(NARRATION,'V A T',0,AdTaxAmt))"
            End If
        End If
    End If
End If
X = " select " & VATTaxCondStr & " as TaxAmt,sum(ADBILLAMT) as TotAmt,sum(VSubBag) as CommBag,AcComOn,AcComRt,ACCOMTYPE from " _
& " tblAddLess,tblMastAccount,tblVouSub,tblMastNarration where ADTAXCODE=NARRCODE and " _
& " tblAddLess.Vno=tblVouSub.Vno and tblAddLess.VType=tblVouSub.Vtype and tblAddLess.VYear=tblVouSub.VYear and " _
& " tblAddLess.AdMillCode=tblMastAccount.AcCode and tblAddLess.VType in ('SO','SD')  and " & CondStr & " and ADMILLCODE=" & MillCode _
& " Group by AcComOn,AcComRt,ACCOMTYPE" 'and AcComOn in (0,2)
If GProcRstOpen(RsTmp, X, "R") > 0 Then
    If IsNumeric(RsTmp!TaxAmt) = True Then
        TaxAmt = CDbl(RsTmp!TaxAmt)
    Else
        TaxAmt = 0
    End If
    If IsNumeric(RsTmp!TotAmt) = True Then
        TotAmt = CDbl(RsTmp!TotAmt)
    Else
        TotAmt = 0
    End If
    If IsNumeric(RsTmp!CommBag) = True Then
        CommBag = CLng(RsTmp!CommBag)
    Else
       CommBag = 0
    End If
    If RsTmp!ACCOMON = 0 Then '--Gross amt
        CommAmt = TotAmt * CDbl(RsTmp!AcComRt) / 100
        CommOn = TotAmt
    ElseIf RsTmp!ACCOMON = 2 Then '-- Net Amt
        CommAmt = (TotAmt - TaxAmt) * CDbl(RsTmp!AcComRt) / 100
        CommOn = CDbl(TotAmt) - CDbl(TaxAmt)
    ElseIf RsTmp!ACCOMON = 1 Then  '---- Commission (On Ex Mill Amt (inward))
        X = " select " & IIf(gBackEndDB = gBackEndAccess, "sum(iif(tblVouSub.VSUBCTRNO=1,ADSUBAMT,0))", "sum(decode(tblVouSub.VSUBCTRNO,1,ADSUBAMT,0))") & " as CommAmt,sum(VSubBag) as CommBag,ACCOMRT,ACCOMTYPE from tblAddLess,tblMastAccount,tblVouSub where " _
        & " tblAddLess.Vno=tblVouSub.Vno and tblAddLess.VType=tblVouSub.Vtype and tblAddLess.VYear=tblVouSub.VYear and " _
        & " tblAddLess.AdMillCode=tblMastAccount.AcCode and tblAddLess.VType in ('PI')  and " & CondStr & " and ADMILLCODE=" & MillCode _
        & " and AcComRt>0 and AcComOn=1 group by ACCOMRT,ACCOMTYPE"
        If GProcRstOpen(RsTmp, X, "R") > 0 Then
            If IsNumeric(RsTmp!CommAmt) = True Then
                CommAmt = GProcMakeRounding(CDbl(RsTmp!CommAmt) * CDbl(RsTmp!AcComRt) / 100)
                CommOn = CDbl(RsTmp!CommAmt)
                CommBag = CLng(RsTmp!CommBag)
            End If
        End If
     Else   '---- Commission (On Ex Mill Amt (sales))  sum( Sale Wt * (Purch ExMill Amt/Total Purch Wt) )
        X = " select sum(tblVouSub.VSubWt*tbladdLess_PU.AdSubamt/( select sum(tblVouSub_PU1.VSubWt) from tblVouSub tblVouSub_PU1 " _
        & " where tblVouSub_PU1.Vno=tblAddLess_PU.Vno and tblVouSub_PU1.Vtype=tblAddLess_PU.VType and tblVouSub_PU1.VYear=tblAddLess_PU.VYear)) " _
        & " as CommAmt,sum(tblVouSub.VSubBag) as CommBag,ACCOMRT,ACCOMTYPE " _
        & " from tblAddLess,tblMastAccount,tblAddLess tblAddLess_PU,tblVouSub,tblVouSub tblVouSub_PU,tblVouSub TblVouSub_GP where " _
        & " tblAddLess.Vno=tblVouSub.Vno and tblAddLess.VType=tblVouSub.Vtype and tblAddLess.VYear=tblVouSub.VYear and " _
        & " tblAddLess_PU.Vno=tblVouSub_PU.Vno and tblAddLess_PU.VType=tblVouSub_PU.Vtype and tblAddLess_PU.VYear=tblVouSub_PU.VYear  and tblAddLess_PU.Vtype='PI' and " _
        & " tblVouSub_GP.VSUBPITCTRLNO=tblVouSub_PU.VSUBITCTRLNO and tblVouSub_GP.Vdt=tblAddLess.AdGpDt and tblVouSub_GP.Vno=tblAddLess.AdGpNo and " _
        & " tblAddLess.AdMillCode=tblMastAccount.AcCode and tblAddLess.VType in ('SO','SD')  and " & CondStr & " and tblAddLess.ADMILLCODE=" & MillCode _
        & " and AcComRt>0 and AcComOn=3 group by ACComRt,ACCOMTYPE"
        If GProcRstOpen(RsTmp, X, "R") > 0 Then
            If IsNumeric(RsTmp!CommAmt) = True Then
                CommAmt = GProcMakeRounding(CDbl(RsTmp!CommAmt) * CDbl(RsTmp!AcComRt) / 100)
                CommOn = CDbl(RsTmp!CommAmt)
                CommBag = CLng(RsTmp!CommBag)
            End If
        End If
    End If
    CommOn = GProcMakeRounding(CommOn)
    CommAmt = GProcMakeRounding(CommAmt)
    Amt = CommAmt
    '-- Commission
    If Amt > 0 Then gCn.Execute "insert into tmpGenTbl (VAcCode,Nar1,UserName,Amt1,Amt2,Amt3,Vno,VSrno) values (" & MillCode & "," _
    & "'" & GProcGetCommissionType(RsTmp!ACCOMTYPE) & " (@ " & RsTmp!AcComRt & " On Amount " & CommOn & " )','" & gUserName & "',-" & Amt & "," & CommOn & "," & CommBag & ",3,0)"
    '-- V A T Payable
    If TaxAmt > 0 Then gCn.Execute "insert into tmpGenTbl (VAcCode,Nar1,UserName,Amt1,Vno,VSrno) values (" & MillCode & "," _
    & "'V A T Payable','" & gUserName & "',-" & TaxAmt & ",4,0)"
    '--- TDS & cess
    TDSAmt = CDbl((Amt * CDbl(TDSP) / 100))
    TDSAmt = GProcMakeRounding(TDSAmt)
    SurchargeAmt = CDbl(TDSAmt * CDbl(Surcharge) / 100)
    SurchargeAmt = GProcMakeRounding(SurchargeAmt)
    TDSCessAmt = CDbl(TDSAmt * CDbl(TDSEcess) / 100)
    TDSCessAmt = GProcMakeRounding(TDSCessAmt)
    If (CDbl(TDSAmt) + CDbl(SurchargeAmt) + CDbl(TDSCessAmt)) > 0 Then gCn.Execute "insert into tmpGenTbl (VAcCode,Nar1,UserName,Amt1,Vno,VSrno) values (" & MillCode & "," _
    & "'TDS @ " & TDSP & " , Surcharge @ " & Surcharge & " & Ed. Cess @ " & TDSEcess & "' ,'" & gUserName & "',-" & CDbl(TDSAmt) + CDbl(SurchargeAmt) + CDbl(TDSCessAmt) & ",6,0)"
    '--- Service Tax & E Cess
    ServiceTaxAmt = CDbl((Amt * CDbl(ServiceTax) / 100))
    ServiceTaxAmt = GProcMakeRounding(ServiceTaxAmt)
    CessAmt = CDbl(ServiceTaxAmt * CDbl(ServiceECess) / 100)
    CessAmt = GProcMakeRounding(CessAmt)
    If CDbl(ServiceTaxAmt + CessAmt) > 0 Then gCn.Execute "insert into tmpGenTbl (VAcCode,Nar1,UserName,Amt1,amt2,amt3,Vno,VSrno) values (" & MillCode & "," _
    & "'Service Tax @ " & ServiceTax & " & Ed. Cess @ " & ServiceECess & "','" & gUserName & "',-" & CDbl(ServiceTaxAmt + CessAmt) & "," & ServiceTaxAmt & "," & CessAmt & ",7,0)"
End If
'==================== Brokerage
Amt = 0
If gBackEndDB = gBackEndAccess Then
    CondStr = " tblAddLess.Vdt>= cdate('" & pFrmDt & "')  and  tblAddLess.Vdt<=cdate('" & pToDt & "')"
    X = " select sum(iif(ITBROKON='Bag',VSUBBAG,iif(ITBROKON='Kg',VSUBWT/100,ADBILLAMT/100)) * ITBROKRT) as BrkAmt from tblAddLess,tblMastAccount,tblMastItem,TblVouSub where " _
    & " tblAddLess.Vtype=TblVouSub.Vtype and  tblAddLess.Vno=TblVouSub.Vno and  tblAddLess.VYear=TblVouSub.VYear and  " _
    & " AdMillCode=AcCode  and ItCode=VSUBITCODE and tblAddLess.VType in ('SO','SD')  and " & CondStr & " and ADMILLCODE=" & MillCode _
    & " and  ACISLOCALBRKBYSELF=1"
Else '-- Oracle
    CondStr = " tblAddLess.Vdt>= to_date('" & CDate(pFrmDt) & "','dd/MM/yyyy')  and tblAddLess.Vdt<=to_date('" & CDate(pToDt) & "','dd/MM/yyyy')"
    X = " select sum(decode(ITBROKON,'Bag',VSUBBAG,'Kg',VSUBWT/100,ADBILLAMT/100) * ITBROKRT) as BrkAmt from tblAddLess,tblMastAccount,tblMastItem,TblVouSub where " _
    & " tblAddLess.Vtype=TblVouSub.Vtype and  tblAddLess.Vno=TblVouSub.Vno and  tblAddLess.VYear=TblVouSub.VYear and  " _
    & " AdMillCode=AcCode  and ItCode=VSUBITCODE and tblAddLess.VType in ('SO','SD')  and " & CondStr & " and ADMILLCODE=" & MillCode _
    & " and  ACISLOCALBRKBYSELF=1"
End If
'& ""
If GProcRstOpen(RsTmp, X, "R") > 0 Then
    If IsNull(RsTmp!BrkAmt) = False Then
        Amt = CDbl(RsTmp!BrkAmt)
    Else
        Amt = 0
    End If
    Amt = GProcMakeRounding(Amt)
End If
    If Amt > 0 Then gCn.Execute "insert into tmpGenTbl (VAcCode,Nar1,UserName,Amt1,Vno,VSrno) values (" & MillCode & "," _
    & "'Brokerage','" & gUserName & "',-" & Amt & ",5,0)"
'==================== Other A/c
If gBackEndDB = gBackEndAccess Then
    CondStr = " Vdt>= cdate('" & pFrmDt & "')  and  Vdt<=cdate('" & pToDt & "')"
    gCn.Execute "insert into tmpGenTbl (VAcCode,Nar1,UserName,Amt1,Vno,VSrno)   select " & MillCode & "," _
    & "Narration,'" & gUserName & "',sum(iif(VACcode=" & MillCode & ",vamt,-Vamt)),8,0 from tblvoucher,tblMastNarration where VType in ('BR','CR','JV','CP','BP','SN','PN')  and " & CondStr & " and " _
    & " (vAcCode=" & MillCode & " or VAcOcode=" & MillCode & ")" _
    & " and MILLEXPNARRCODE=NarrCode and VCTRNO=1 group by Narration order by Narration"
Else '-- Oracle
    CondStr = " Vdt>= to_date('" & CDate(pFrmDt) & "','dd/MM/yyyy')  and Vdt<=to_date('" & CDate(pToDt) & "','dd/MM/yyyy')"
    gCn.Execute "insert into tmpGenTbl (VAcCode,Nar1,UserName,Amt1,Vno,VSrno)   select " & MillCode & "," _
    & "Narration,'" & gUserName & "',sum(decode(VACcode," & MillCode & ",vamt,-Vamt)),8,0 from tblvoucher,tblMastNarration where VType in ('BR','CR','JV','CP','BP','SN','PN')  and " & CondStr & " and " _
    & " (vAcCode=" & MillCode & " or VAcOcode=" & MillCode & ")" _
    & " and MILLEXPNARRCODE=NarrCode and VCTRNO=1 group by Narration order by Narration"
End If
gCn.CommitTrans
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
        BrkName = GProcGetColumnValue("TblMastAccount", "ACCode", !vacCode, "N", "AcName", "S", "")   '-Broker
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
Public Function GProcPurpareGSTR9(w_Fr As Date, w_To As Date, OnPurBillDt As Boolean, OnPurRecDt As Boolean)
Dim X As String
Dim X1 As String
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

gCn.Execute "update tblMastItem Set ITTMP2 = 'X' where isnull(ITTMP2)"
gCn.CommitTrans

'--For Total yarn Sale R.D.
gCn.BeginTrans
X1 = " tblAddLess.VType in ('SY')  and tblAddLess.vdt>=cdate('" & Format(w_Fr, "yyyy,mm,dd") & "')  and tblAddLess.vdt<=Cdate('" & Format(w_To, "yyyy,mm,dd") & "') and (tblAddLess.AdTmp5 <> 'COL' or isnull(tblAddLess.AdTmp5)) and VsubCGSTAmt+VsubSGSTAmt+VsubIGSTAmt > 0 and len(tblMastAccount_Pty.AcGSTIN) > 2"
X = "select 'A','Sales',tblMastAccount.acname,tblMastItem.ITTMP2,tblVouSub.VsubCGSTRt+tblVouSub.VsubSGSTRt+tblVouSub.VsubIGSTRt," _
    & " Sum (VSubTaxableAmt), Sum(VsubCGSTAmt), Sum(VsubSGSTAmt), Sum(VsubIGSTAmt), Sum(AdBillAmt),'" & gUserName & "'" _
    & " from tblAddLess,tblVouSub,tblMastItem,tblMastAccount,tblMastAccount tblMastAccount_Pty " _
    & " where tblAddLess.AdAcCrCode = tblMastAccount.AcCode and tblVouSub.VSubItCode = tblMastItem.ItCode and tblAddLess.Vno = tblVouSub.Vno and " _
    & " tblAddLess.Vtype = tblVouSub.Vtype and tblAddLess.Vyear = tblVouSub.Vyear and " & X1 & " " _
    & " and tblAddLess.AdAcDrCode = tblMastAccount_Pty.AcCode  " _
    & " group by tblMastAccount.acname,tblMastItem.ITTMP2,tblVouSub.VsubCGSTRt+tblVouSub.VsubSGSTRt+tblVouSub.VsubIGSTRt"

gCn.Execute " Insert Into TmpGentbl (VType,Nar1,Nar2,Nar3,Amt1,Amt2,Amt3,Amt4,Amt5,Amt6,UserName)" & X
gCn.CommitTrans

Set excel_app = New Excel.Application
excel_app.Visible = False
Set workbook = excel_app.Workbooks.Open(filename:=flnm)
RecYN = GProcRstOpen(tmp1, "select * from TmpGentbl where vtype = 'A' and TmpGentbl.UserName ='" & gUserName & "'", "R", gCn)
If RecYN > 0 Then
   Set sheet = workbook.Sheets.Item("Yearly")
   sheet.Cells(2, 3) = gCName
   sheet.Cells(5, 4) = gCGSTIN
   sheet.Cells(8, 4) = CStr(Year(gCYSDate)) + "-" + CStr(Year(gCYEDate))
   
   
   If (sheet Is Nothing) Then
       MsgBox "Error Yearly Sheet"
   End If
   a = 16
   With tmp1
        .MoveFirst
        Do While Not .EOF
           sheet.Cells(a, 2) = "Sales"
           sheet.Cells(a, 3) = !Nar2
           If Left(!Nar3, 2) = "99" Then
              sheet.Cells(a, 4) = "Services"
           Else
              sheet.Cells(a, 4) = "Goods"
           End If
           sheet.Cells(a, 5) = "Registered"
           sheet.Cells(a, 6) = "Tax Invoice"
           sheet.Cells(a, 7) = !Nar3
           sheet.Cells(a, 8) = !Amt1
           gCn.BeginTrans
           For i = 1 To 12
               Set sheet = workbook.Sheets.Item("" & Left(MonthName(i), 3) & "")
               X = "select 'A','1',tblMastAccount.acname,tblMastItem.ITTMP2,tblVouSub.VsubCGSTRt+tblVouSub.VsubSGSTRt+tblVouSub.VsubIGSTRt," _
                   & " Sum (VSubTaxableAmt), Sum(VsubCGSTAmt), Sum(VsubSGSTAmt), Sum(VsubIGSTAmt), Sum(AdBillAmt)," & i & ",'" & gUserName & "'" _
                   & " from tblAddLess,tblVouSub,tblMastItem,tblMastAccount,tblMastAccount tblMastAccount_Pty " _
                   & " where tblAddLess.AdAcCrCode = tblMastAccount.AcCode and tblVouSub.VSubItCode = tblMastItem.ItCode and tblAddLess.Vno = tblVouSub.Vno and " _
                   & " tblAddLess.Vtype = tblVouSub.Vtype and tblAddLess.Vyear = tblVouSub.Vyear and " & X1 & " " _
                   & " and tblAddLess.AdAcDrCode = tblMastAccount_Pty.AcCode and month(tblAddLess.Vdt) = " & i & " and iif(isnull(tblMastItem.ITTMP2),isnull(tblMastItem.ITTMP2), tblMastItem.ITTMP2 = '" & tmp1.Fields("Nar3") & "') and tblVouSub.VsubCGSTRt+tblVouSub.VsubSGSTRt+tblVouSub.VsubIGSTRt = " & tmp1.Fields("Amt1") & " " _
                   & " and tblMastAccount.AcName = '" & tmp1.Fields("Nar2") & "' " _
                   & " group by tblMastAccount.acname,tblMastItem.ITTMP2,tblVouSub.VsubCGSTRt+tblVouSub.VsubSGSTRt+tblVouSub.VsubIGSTRt"
               
               gCn.Execute " Insert Into TmpGentbl2 (VType,Nar1,Nar2,Nar3,Amt1,Amt2,Amt3,Amt4,Amt5,Amt6,Vno,UserName)" & X
               
               j = GProcRstOpen(tmp2, X, "R", gCn)
               If j > 0 Then
                  sheet.Cells(a, 10) = IIf(IsNull(tmp2.Fields(5)), 0, tmp2.Fields(5))
                  sheet.Cells(a, 11) = IIf(IsNull(tmp2.Fields(6)), 0, tmp2.Fields(6))
                  sheet.Cells(a, 12) = IIf(IsNull(tmp2.Fields(7)), 0, tmp2.Fields(7))
                  sheet.Cells(a, 13) = IIf(IsNull(tmp2.Fields(8)), 0, tmp2.Fields(8))
                  sheet.Cells(a, 14) = 0
               Else
                  sheet.Cells(a, 10) = 0
                  sheet.Cells(a, 11) = 0
                  sheet.Cells(a, 12) = 0
                  sheet.Cells(a, 13) = 0
                  sheet.Cells(a, 14) = 0
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

X1 = " tblAddLess.VType in ('SY')  and tblAddLess.vdt>=cdate('" & Format(w_Fr, "yyyy,mm,dd") & "')  and tblAddLess.vdt<=Cdate('" & Format(w_To, "yyyy,mm,dd") & "') and (tblAddLess.AdTmp5 <> 'COL' or isnull(tblAddLess.AdTmp5)) and VsubCGSTAmt+VsubSGSTAmt+VsubIGSTAmt > 0 and (len(tblMastAccount_Pty.AcGSTIN) <= 2 or isnull(tblMastAccount_Pty.AcGSTIN))"
X = "select 'B','Sales',tblMastAccount.acname,tblMastItem.ITTMP2,tblVouSub.VsubCGSTRt+tblVouSub.VsubSGSTRt+tblVouSub.VsubIGSTRt," _
    & " Sum (VSubTaxableAmt), Sum(VsubCGSTAmt), Sum(VsubSGSTAmt), Sum(VsubIGSTAmt), Sum(AdBillAmt),'" & gUserName & "'" _
    & " from tblAddLess,tblVouSub,tblMastItem,tblMastAccount,tblMastAccount tblMastAccount_Pty " _
    & " where tblAddLess.AdAcCrCode = tblMastAccount.AcCode and tblVouSub.VSubItCode = tblMastItem.ItCode and tblAddLess.Vno = tblVouSub.Vno and " _
    & " tblAddLess.Vtype = tblVouSub.Vtype and tblAddLess.Vyear = tblVouSub.Vyear and " & X1 & " " _
    & " and tblAddLess.AdAcDrCode = tblMastAccount_Pty.AcCode  " _
    & " group by tblMastAccount.acname,tblMastItem.ITTMP2,tblVouSub.VsubCGSTRt+tblVouSub.VsubSGSTRt+tblVouSub.VsubIGSTRt"

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
           sheet.Cells(a, 2) = "Sales"
           sheet.Cells(a, 3) = !Nar2
           If Left(!Nar3, 2) = "99" Then
              sheet.Cells(a, 4) = "Services"
           Else
              sheet.Cells(a, 4) = "Goods"
           End If
           sheet.Cells(a, 5) = "Unregistered"
           sheet.Cells(a, 6) = "Tax Invoice"
           sheet.Cells(a, 7) = !Nar3
           sheet.Cells(a, 8) = !Amt1
           gCn.BeginTrans
           For i = 1 To 12
               Set sheet = workbook.Sheets.Item("" & Left(MonthName(i), 3) & "")
               X = "select 'B','1',tblMastAccount.acname,tblMastItem.ITTMP2,tblVouSub.VsubCGSTRt+tblVouSub.VsubSGSTRt+tblVouSub.VsubIGSTRt," _
                   & " Sum (VSubTaxableAmt), Sum(VsubCGSTAmt), Sum(VsubSGSTAmt), Sum(VsubIGSTAmt), Sum(AdBillAmt)," & i & ",'" & gUserName & "'" _
                   & " from tblAddLess,tblVouSub,tblMastItem,tblMastAccount,tblMastAccount tblMastAccount_Pty " _
                   & " where tblAddLess.AdAcCrCode = tblMastAccount.AcCode and tblVouSub.VSubItCode = tblMastItem.ItCode and tblAddLess.Vno = tblVouSub.Vno and " _
                   & " tblAddLess.Vtype = tblVouSub.Vtype and tblAddLess.Vyear = tblVouSub.Vyear and " & X1 & " and month(tblAddLess.Vdt) = " & i & " and iif(isnull(tblMastItem.ITTMP2),isnull(tblMastItem.ITTMP2), tblMastItem.ITTMP2 = '" & tmp1.Fields("Nar3") & "') and tblVouSub.VsubCGSTRt+tblVouSub.VsubSGSTRt+tblVouSub.VsubIGSTRt = " & tmp1.Fields("Amt1") & " " _
                   & " and tblAddLess.AdAcDrCode = tblMastAccount_Pty.AcCode  " _
                   & " and tblMastAccount.AcName = '" & tmp1.Fields("Nar2") & "' " _
                   & " group by tblMastAccount.acname,tblMastItem.ITTMP2,tblVouSub.VsubCGSTRt+tblVouSub.VsubSGSTRt+tblVouSub.VsubIGSTRt"
               gCn.Execute " Insert Into TmpGentbl2 (VType,Nar1,Nar2,Nar3,Amt1,Amt2,Amt3,Amt4,Amt5,Amt6,Vno,UserName)" & X
               j = GProcRstOpen(tmp2, X, "R", gCn)
               If j > 0 Then
                  sheet.Cells(a, 10) = IIf(IsNull(tmp2.Fields(5)), 0, tmp2.Fields(5))
                  sheet.Cells(a, 11) = IIf(IsNull(tmp2.Fields(6)), 0, tmp2.Fields(6))
                  sheet.Cells(a, 12) = IIf(IsNull(tmp2.Fields(7)), 0, tmp2.Fields(7))
                  sheet.Cells(a, 13) = IIf(IsNull(tmp2.Fields(8)), 0, tmp2.Fields(8))
                  sheet.Cells(a, 14) = 0
               Else
                  sheet.Cells(a, 10) = 0
                  sheet.Cells(a, 11) = 0
                  sheet.Cells(a, 12) = 0
                  sheet.Cells(a, 13) = 0
                  sheet.Cells(a, 14) = 0
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
''''-- For Cloth Sales R.D.
'''gCn.BeginTrans
'''X1 = " tblAddLess.VType in ('SY')  and tblAddLess.vdt>=cdate('" & Format(w_Fr, "yyyy,mm,dd") & "')  and tblAddLess.vdt<=Cdate('" & Format(w_To, "yyyy,mm,dd") & "') and (tblAddLess.AdTmp5 = 'COL') and AdClothCgstAmt+AdClothSgstAmt+AdClothIgstAmt > 0 and len(tblMastAccount_Pty.AcGSTIN) > 2 "
'''X = "select 'C','Sales',tblMastAccount.acname,tblMastItem.ITTMP2,tblVouSub.VsubCGSTRt+tblVouSub.VsubSGSTRt+tblVouSub.VsubIGSTRt," _
'''    & " Sum (tblAddLess.AdAssVal), Sum(AdClothCgstAmt), Sum(AdClothSgstAmt), Sum(AdClothIgstAmt), Sum(AdBillAmt),'" & gUserName & "'" _
'''    & " from tblAddLess,tblVouSub,tblMastItem,tblMastAccount,tblMastAccount tblMastAccount_Pty " _
'''    & " where tblAddLess.AdAcCrCode = tblMastAccount.AcCode and tblVouSub.VSubItCode = tblMastItem.ItCode and tblAddLess.Vno = tblVouSub.Vno and " _
'''    & " tblAddLess.Vtype = tblVouSub.Vtype and tblAddLess.Vyear = tblVouSub.Vyear and " & X1 & " " _
'''    & " and tblAddLess.AdAcDrCode = tblMastAccount_Pty.AcCode  " _
'''    & " group by tblMastAccount.acname,tblMastItem.ITTMP2,tblVouSub.VsubCGSTRt+tblVouSub.VsubSGSTRt+tblVouSub.VsubIGSTRt"
'''
'''gCn.Execute " Insert Into TmpGentbl (VType,Nar1,Nar2,Nar3,Amt1,Amt2,Amt3,Amt4,Amt5,Amt6,UserName)" & X
'''gCn.CommitTrans
'''
'''RecYN = GProcRstOpen(tmp1, "select * from TmpGentbl where vtype = 'C' and TmpGentbl.UserName ='" & gUserName & "'", "R", gCn)
'''If RecYN > 0 Then
'''   Set sheet = workbook.Sheets.Item("Yearly")
'''   If (sheet Is Nothing) Then
'''       MsgBox "Error Yearly Sheet"
'''   End If
'''   With tmp1
'''        .MoveFirst
'''        Do While Not .EOF
'''           sheet.cells(a, 2) = "Sales"
'''           sheet.cells(a, 3) = !Nar2
'''           If Left(!Nar3, 2) = "99" Then
'''              sheet.cells(a, 4) = "Services"
'''           Else
'''              sheet.cells(a, 4) = "Goods"
'''           End If
'''           sheet.cells(a, 5) = "Registered"
'''           sheet.cells(a, 6) = "Tax Invoice"
'''           sheet.cells(a, 7) = !Nar3
'''           sheet.cells(a, 8) = !Amt1
'''           a = a + 1
'''           If a = 35 Then a = 37
'''           If a = 39 Then a = 45
'''           .MoveNext
'''        Loop
'''   End With
'''End If
''''-- For Cloth Sales U.R.D.
'''gCn.BeginTrans
'''X1 = " tblAddLess.VType in ('SY')  and tblAddLess.vdt>=cdate('" & Format(w_Fr, "yyyy,mm,dd") & "')  and tblAddLess.vdt<=Cdate('" & Format(w_To, "yyyy,mm,dd") & "') and (tblAddLess.AdTmp5 = 'COL') and AdClothCgstAmt+AdClothSgstAmt+AdClothIgstAmt > 0 and (len(tblMastAccount_Pty.AcGSTIN) <= 2 or isnull(tblMastAccount_Pty.AcGSTIN))"
'''X = "select 'D','Sales',tblMastAccount.acname,tblMastItem.ITTMP2,tblVouSub.VsubCGSTRt+tblVouSub.VsubSGSTRt+tblVouSub.VsubIGSTRt," _
'''    & " Sum (tblAddLess.AdAssVal), Sum(AdClothCgstAmt), Sum(AdClothSgstAmt), Sum(AdClothIgstAmt), Sum(AdBillAmt),'" & gUserName & "'" _
'''    & " from tblAddLess,tblVouSub,tblMastItem,tblMastAccount,tblMastAccount tblMastAccount_Pty " _
'''    & " where tblAddLess.AdAcCrCode = tblMastAccount.AcCode and tblVouSub.VSubItCode = tblMastItem.ItCode and tblAddLess.Vno = tblVouSub.Vno and " _
'''    & " tblAddLess.Vtype = tblVouSub.Vtype and tblAddLess.Vyear = tblVouSub.Vyear and " & X1 & " " _
'''    & " and tblAddLess.AdAcDrCode = tblMastAccount_Pty.AcCode  " _
'''    & " group by tblMastAccount.acname,tblMastItem.ITTMP2,tblVouSub.VsubCGSTRt+tblVouSub.VsubSGSTRt+tblVouSub.VsubIGSTRt"
'''
'''gCn.Execute " Insert Into TmpGentbl (VType,Nar1,Nar2,Nar3,Amt1,Amt2,Amt3,Amt4,Amt5,Amt6,UserName)" & X
'''gCn.CommitTrans
'''
'''RecYN = GProcRstOpen(tmp1, "select * from TmpGentbl where vtype = 'D' and TmpGentbl.UserName ='" & gUserName & "'", "R", gCn)
'''If RecYN > 0 Then
'''   Set sheet = workbook.Sheets.Item("Yearly")
'''   If (sheet Is Nothing) Then
'''       MsgBox "Error Yearly Sheet"
'''   End If
'''   With tmp1
'''        .MoveFirst
'''        Do While Not .EOF
'''           sheet.cells(a, 2) = "Sales"
'''           sheet.cells(a, 3) = !Nar2
'''           If Left(!Nar3, 2) = "99" Then
'''              sheet.cells(a, 4) = "Services"
'''           Else
'''              sheet.cells(a, 4) = "Goods"
'''           End If
'''           sheet.cells(a, 5) = "Unregistered"
'''           sheet.cells(a, 6) = "Tax Invoice"
'''           sheet.cells(a, 7) = !Nar3
'''           sheet.cells(a, 8) = !Amt1
'''           a = a + 1
'''           If a = 35 Then a = 37
'''           If a = 39 Then a = 45
'''           .MoveNext
'''        Loop
'''   End With
'''End If

'--For Total yarn Purchase R.D.
gCn.BeginTrans
If OnPurBillDt = True Then
   X1 = " tblAddLess.VType in ('PY')  and tblAddLess.AdGpDt>=cdate('" & Format(w_Fr, "yyyy,mm,dd") & "')  and tblAddLess.AdGpDt<=Cdate('" & Format(w_To, "yyyy,mm,dd") & "') and (tblAddLess.AdTmp5 <> 'COL' or isnull(tblAddLess.AdTmp5)) and VsubCGSTAmt+VsubSGSTAmt+VsubIGSTAmt > 0 and len(tblMastAccount_Pty.AcGSTIN) > 2"
Else
   X1 = " tblAddLess.VType in ('PY')  and tblAddLess.vdt>=cdate('" & Format(w_Fr, "yyyy,mm,dd") & "')  and tblAddLess.vdt<=Cdate('" & Format(w_To, "yyyy,mm,dd") & "') and (tblAddLess.AdTmp5 <> 'COL' or isnull(tblAddLess.AdTmp5)) and VsubCGSTAmt+VsubSGSTAmt+VsubIGSTAmt > 0 and len(tblMastAccount_Pty.AcGSTIN) > 2"
End If
X = "select 'E','Purchases',tblMastAccount.acname,tblMastItem.ITTMP2,tblVouSub.VsubCGSTRt+tblVouSub.VsubSGSTRt+tblVouSub.VsubIGSTRt," _
    & " Sum (VSubTaxableAmt), Sum(VsubCGSTAmt), Sum(VsubSGSTAmt), Sum(VsubIGSTAmt), Sum(AdBillAmt),'" & gUserName & "'" _
    & " from tblAddLess,tblVouSub,tblMastItem,tblMastAccount,tblMastAccount tblMastAccount_Pty " _
    & " where tblAddLess.AdAcDrCode = tblMastAccount.AcCode and tblVouSub.VSubItCode = tblMastItem.ItCode and tblAddLess.Vno = tblVouSub.Vno and " _
    & " tblAddLess.Vtype = tblVouSub.Vtype and tblAddLess.Vyear = tblVouSub.Vyear and " & X1 & " " _
    & " and tblAddLess.AdAcCrCode = tblMastAccount_Pty.AcCode  " _
    & " group by tblMastAccount.acname,tblMastItem.ITTMP2,tblVouSub.VsubCGSTRt+tblVouSub.VsubSGSTRt+tblVouSub.VsubIGSTRt"

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
           sheet.Cells(a, 2) = "Purchases"
           sheet.Cells(a, 3) = !Nar2
           If Left(!Nar3, 2) = "99" Then
              sheet.Cells(a, 4) = "Services"
           Else
              sheet.Cells(a, 4) = "Goods"
           End If
           sheet.Cells(a, 5) = "Registered"
           sheet.Cells(a, 6) = "Tax Invoice"
           sheet.Cells(a, 7) = !Nar3
           sheet.Cells(a, 8) = !Amt1
           gCn.BeginTrans
           For i = 1 To 12
               Set sheet = workbook.Sheets.Item("" & Left(MonthName(i), 3) & "")
               X = "select 'E','4',tblMastAccount.acname,tblMastItem.ITTMP2,tblVouSub.VsubCGSTRt+tblVouSub.VsubSGSTRt+tblVouSub.VsubIGSTRt," _
                   & " Sum (VSubTaxableAmt), Sum(VsubCGSTAmt), Sum(VsubSGSTAmt), Sum(VsubIGSTAmt), Sum(AdBillAmt)," & i & ",'" & gUserName & "'" _
                   & " from tblAddLess,tblVouSub,tblMastItem,tblMastAccount,tblMastAccount tblMastAccount_Pty " _
                   & " where tblAddLess.AdAcDrCode = tblMastAccount.AcCode and tblVouSub.VSubItCode = tblMastItem.ItCode and tblAddLess.Vno = tblVouSub.Vno and " _
                   & " tblAddLess.Vtype = tblVouSub.Vtype and tblAddLess.Vyear = tblVouSub.Vyear and " & X1 & " and month(iif(" & OnPurBillDt & " = true ,tblAddLess.AdGpDt,tblAddLess.Vdt)) = " & i & " and iif(isnull(tblMastItem.ITTMP2),isnull(tblMastItem.ITTMP2), tblMastItem.ITTMP2 = '" & tmp1.Fields("Nar3") & "') and tblVouSub.VsubCGSTRt+tblVouSub.VsubSGSTRt+tblVouSub.VsubIGSTRt = " & tmp1.Fields("Amt1") & "" _
                   & " and tblAddLess.AdAcCrCode = tblMastAccount_Pty.AcCode  " _
                   & " and tblMastAccount.AcName = '" & tmp1.Fields("Nar2") & "' " _
                   & " group by tblMastAccount.acname,tblMastItem.ITTMP2,tblVouSub.VsubCGSTRt+tblVouSub.VsubSGSTRt+tblVouSub.VsubIGSTRt"
               gCn.Execute " Insert Into TmpGentbl2 (VType,Nar1,Nar2,Nar3,Amt1,Amt2,Amt3,Amt4,Amt5,Amt6,Vno,UserName)" & X
               j = GProcRstOpen(tmp2, X, "R", gCn)
               If j > 0 Then
                  sheet.Cells(a, 10) = IIf(IsNull(tmp2.Fields(5)), 0, tmp2.Fields(5))
                  sheet.Cells(a, 11) = IIf(IsNull(tmp2.Fields(6)), 0, tmp2.Fields(6))
                  sheet.Cells(a, 12) = IIf(IsNull(tmp2.Fields(7)), 0, tmp2.Fields(7))
                  sheet.Cells(a, 13) = IIf(IsNull(tmp2.Fields(8)), 0, tmp2.Fields(8))
                  sheet.Cells(a, 14) = 0
               Else
                  sheet.Cells(a, 10) = 0
                  sheet.Cells(a, 11) = 0
                  sheet.Cells(a, 12) = 0
                  sheet.Cells(a, 13) = 0
                  sheet.Cells(a, 14) = 0
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
'--For Total yarn Purchase U.R.D.
gCn.BeginTrans
If OnPurBillDt = True Then
   X1 = " tblAddLess.VType in ('PY')  and tblAddLess.AdGpDt>=cdate('" & Format(w_Fr, "yyyy,mm,dd") & "')  and tblAddLess.AdGpDt<=Cdate('" & Format(w_To, "yyyy,mm,dd") & "') and (tblAddLess.AdTmp5 <> 'COL' or isnull(tblAddLess.AdTmp5)) and VsubCGSTAmt+VsubSGSTAmt+VsubIGSTAmt > 0 and (len(tblMastAccount_Pty.AcGSTIN) <= 2 or isnull(tblMastAccount_Pty.AcGSTIN))"
Else
   X1 = " tblAddLess.VType in ('PY')  and tblAddLess.vdt>=cdate('" & Format(w_Fr, "yyyy,mm,dd") & "')  and tblAddLess.vdt<=Cdate('" & Format(w_To, "yyyy,mm,dd") & "') and (tblAddLess.AdTmp5 <> 'COL' or isnull(tblAddLess.AdTmp5)) and VsubCGSTAmt+VsubSGSTAmt+VsubIGSTAmt > 0 and (len(tblMastAccount_Pty.AcGSTIN) <= 2 or isnull(tblMastAccount_Pty.AcGSTIN))"
End If
X = "select 'F','Purchases',tblMastAccount.acname,tblMastItem.ITTMP2,tblVouSub.VsubCGSTRt+tblVouSub.VsubSGSTRt+tblVouSub.VsubIGSTRt," _
    & " Sum (VSubTaxableAmt), Sum(VsubCGSTAmt), Sum(VsubSGSTAmt), Sum(VsubIGSTAmt), Sum(AdBillAmt),'" & gUserName & "'" _
    & " from tblAddLess,tblVouSub,tblMastItem,tblMastAccount,tblMastAccount tblMastAccount_Pty " _
    & " where tblAddLess.AdAcDrCode = tblMastAccount.AcCode and tblVouSub.VSubItCode = tblMastItem.ItCode and tblAddLess.Vno = tblVouSub.Vno and " _
    & " tblAddLess.Vtype = tblVouSub.Vtype and tblAddLess.Vyear = tblVouSub.Vyear and " & X1 & " " _
    & " and tblAddLess.AdAcCrCode = tblMastAccount_Pty.AcCode  " _
    & " group by tblMastAccount.acname,tblMastItem.ITTMP2,tblVouSub.VsubCGSTRt+tblVouSub.VsubSGSTRt+tblVouSub.VsubIGSTRt"

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
           sheet.Cells(a, 2) = "Purchases"
           sheet.Cells(a, 3) = !Nar2
           If Left(!Nar3, 2) = "99" Then
              sheet.Cells(a, 4) = "Services"
           Else
              sheet.Cells(a, 4) = "Goods"
           End If
           sheet.Cells(a, 5) = "Unregistered"
           sheet.Cells(a, 6) = "Tax Invoice"
           sheet.Cells(a, 7) = !Nar3
           sheet.Cells(a, 8) = !Amt1
           gCn.BeginTrans
           For i = 1 To 12
               Set sheet = workbook.Sheets.Item("" & Left(MonthName(i), 3) & "")
               X = "select 'F','4',tblMastAccount.acname,tblMastItem.ITTMP2,tblVouSub.VsubCGSTRt+tblVouSub.VsubSGSTRt+tblVouSub.VsubIGSTRt," _
                   & " Sum (VSubTaxableAmt), Sum(VsubCGSTAmt), Sum(VsubSGSTAmt), Sum(VsubIGSTAmt), Sum(AdBillAmt)," & i & ",'" & gUserName & "'" _
                   & " from tblAddLess,tblVouSub,tblMastItem,tblMastAccount,tblMastAccount tblMastAccount_Pty " _
                   & " where tblAddLess.AdAcDrCode = tblMastAccount.AcCode and tblVouSub.VSubItCode = tblMastItem.ItCode and tblAddLess.Vno = tblVouSub.Vno and " _
                   & " tblAddLess.Vtype = tblVouSub.Vtype and tblAddLess.Vyear = tblVouSub.Vyear and " & X1 & " and month(iif(" & OnPurBillDt & " = true ,tblAddLess.AdGpDt,tblAddLess.Vdt)) = " & i & " and iif(isnull(tblMastItem.ITTMP2),isnull(tblMastItem.ITTMP2), tblMastItem.ITTMP2 = '" & tmp1.Fields("Nar3") & "') and tblVouSub.VsubCGSTRt+tblVouSub.VsubSGSTRt+tblVouSub.VsubIGSTRt = " & tmp1.Fields("Amt1") & " " _
                   & " and tblAddLess.AdAcCrCode = tblMastAccount_Pty.AcCode  " _
                   & " and tblMastAccount.AcName = '" & tmp1.Fields("Nar2") & "' " _
                   & " group by tblMastAccount.acname,tblMastItem.ITTMP2,tblVouSub.VsubCGSTRt+tblVouSub.VsubSGSTRt+tblVouSub.VsubIGSTRt"
                
               gCn.Execute " Insert Into TmpGentbl2 (VType,Nar1,Nar2,Nar3,Amt1,Amt2,Amt3,Amt4,Amt5,Amt6,Vno,UserName)" & X
               j = GProcRstOpen(tmp2, X, "R", gCn)
               If j > 0 Then
                  sheet.Cells(a, 10) = IIf(IsNull(tmp2.Fields(5)), 0, tmp2.Fields(5))
                  sheet.Cells(a, 11) = IIf(IsNull(tmp2.Fields(6)), 0, tmp2.Fields(6))
                  sheet.Cells(a, 12) = IIf(IsNull(tmp2.Fields(7)), 0, tmp2.Fields(7))
                  sheet.Cells(a, 13) = IIf(IsNull(tmp2.Fields(8)), 0, tmp2.Fields(8))
                  sheet.Cells(a, 14) = 0
               Else
                  sheet.Cells(a, 10) = 0
                  sheet.Cells(a, 11) = 0
                  sheet.Cells(a, 12) = 0
                  sheet.Cells(a, 13) = 0
                  sheet.Cells(a, 14) = 0
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
'--For Total yarn Other Purchase R.D.
gCn.BeginTrans
If OnPurBillDt = True Then
   X1 = " tblAddLess.VType in ('PO')  and tblAddLess.AdGpDt>=cdate('" & Format(w_Fr, "yyyy,mm,dd") & "')  and tblAddLess.AdGpDt<=Cdate('" & Format(w_To, "yyyy,mm,dd") & "') and (tblAddLess.AdTmp5 <> 'COL' or isnull(tblAddLess.AdTmp5)) and VsubCGSTAmt+VsubSGSTAmt+VsubIGSTAmt > 0 and len(tblMastAccount_Pty.AcGSTIN) > 2 and tblAddLess.AdLCostAdd+tblAddLess.AdLCostFreight+tblAddLess.AdLCostGRent = 0"
Else
   X1 = " tblAddLess.VType in ('PO')  and tblAddLess.vdt>=cdate('" & Format(w_Fr, "yyyy,mm,dd") & "')  and tblAddLess.vdt<=Cdate('" & Format(w_To, "yyyy,mm,dd") & "') and (tblAddLess.AdTmp5 <> 'COL' or isnull(tblAddLess.AdTmp5)) and VsubCGSTAmt+VsubSGSTAmt+VsubIGSTAmt > 0 and len(tblMastAccount_Pty.AcGSTIN) > 2 and tblAddLess.AdLCostAdd+tblAddLess.AdLCostFreight+tblAddLess.AdLCostGRent = 0  "
End If
X = "select 'G','Purchases',tblMastAccount.acname,tblMastItem.ITTMP2,tblVouSub.VsubCGSTRt+tblVouSub.VsubSGSTRt+tblVouSub.VsubIGSTRt," _
    & " Sum (VSubTaxableAmt), Sum(VsubCGSTAmt), Sum(VsubSGSTAmt), Sum(VsubIGSTAmt), Sum(AdBillAmt),'" & gUserName & "'" _
    & " from tblAddLess,tblVouSub,tblMastItem,tblMastAccount,tblMastAccount tblMastAccount_Pty " _
    & " where tblAddLess.AdAcDrCode = tblMastAccount.AcCode and tblVouSub.VSubItCode = tblMastItem.ItCode and tblAddLess.Vno = tblVouSub.Vno and " _
    & " tblAddLess.Vtype = tblVouSub.Vtype and tblAddLess.Vyear = tblVouSub.Vyear and " & X1 & " " _
    & " and tblAddLess.AdAcCrCode = tblMastAccount_Pty.AcCode  " _
    & " group by tblMastAccount.acname,tblMastItem.ITTMP2,tblVouSub.VsubCGSTRt+tblVouSub.VsubSGSTRt+tblVouSub.VsubIGSTRt"

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
           sheet.Cells(a, 2) = "Indirect Expenses"
           sheet.Cells(a, 3) = !Nar2
           If Left(!Nar3, 2) = "99" Then
              sheet.Cells(a, 4) = "Services"
           Else
              sheet.Cells(a, 4) = "Goods"
           End If
           sheet.Cells(a, 5) = "Registered"
           sheet.Cells(a, 6) = "Tax Invoice"
           sheet.Cells(a, 7) = !Nar3
           sheet.Cells(a, 8) = !Amt1
           gCn.BeginTrans
           For i = 1 To 12
               Set sheet = workbook.Sheets.Item("" & Left(MonthName(i), 3) & "")
               X = "select 'G','4',tblMastAccount.acname,tblMastItem.ITTMP2,tblVouSub.VsubCGSTRt+tblVouSub.VsubSGSTRt+tblVouSub.VsubIGSTRt," _
                   & " Sum (VSubTaxableAmt), Sum(VsubCGSTAmt), Sum(VsubSGSTAmt), Sum(VsubIGSTAmt), Sum(AdBillAmt)," & i & ",'" & gUserName & "'" _
                   & " from tblAddLess,tblVouSub,tblMastItem,tblMastAccount,tblMastAccount tblMastAccount_Pty " _
                   & " where tblAddLess.AdAcDrCode = tblMastAccount.AcCode and tblVouSub.VSubItCode = tblMastItem.ItCode and tblAddLess.Vno = tblVouSub.Vno and " _
                   & " tblAddLess.Vtype = tblVouSub.Vtype and tblAddLess.Vyear = tblVouSub.Vyear and " & X1 & " and month(iif(" & OnPurBillDt & " = true ,tblAddLess.AdGpDt,tblAddLess.Vdt)) = " & i & " and iif(isnull(tblMastItem.ITTMP2),isnull(tblMastItem.ITTMP2), tblMastItem.ITTMP2 = '" & tmp1.Fields("Nar3") & "') and tblVouSub.VsubCGSTRt+tblVouSub.VsubSGSTRt+tblVouSub.VsubIGSTRt = " & tmp1.Fields("Amt1") & " " _
                   & " and tblAddLess.AdAcCrCode = tblMastAccount_Pty.AcCode and tblMastAccount.AcName = '" & tmp1.Fields("Nar2") & "'  " _
                   & " group by tblMastAccount.acname,tblMastItem.ITTMP2,tblVouSub.VsubCGSTRt+tblVouSub.VsubSGSTRt+tblVouSub.VsubIGSTRt"

               gCn.Execute " Insert Into TmpGentbl2 (VType,Nar1,Nar2,Nar3,Amt1,Amt2,Amt3,Amt4,Amt5,Amt6,Vno,UserName)" & X

               j = GProcRstOpen(tmp2, X, "R", gCn)
               If j > 0 Then
                  sheet.Cells(a, 10) = IIf(IsNull(tmp2.Fields(5)), 0, tmp2.Fields(5))
                  sheet.Cells(a, 11) = IIf(IsNull(tmp2.Fields(6)), 0, tmp2.Fields(6))
                  sheet.Cells(a, 12) = IIf(IsNull(tmp2.Fields(7)), 0, tmp2.Fields(7))
                  sheet.Cells(a, 13) = IIf(IsNull(tmp2.Fields(8)), 0, tmp2.Fields(8))
                  sheet.Cells(a, 14) = 0
               Else
                  sheet.Cells(a, 10) = 0
                  sheet.Cells(a, 11) = 0
                  sheet.Cells(a, 12) = 0
                  sheet.Cells(a, 13) = 0
                  sheet.Cells(a, 14) = 0
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

'--Credit/Debit Note Sales
gCn.BeginTrans

X1 = " qrytblVoucherSr1.VType in ('SN','PN') and qrytblVoucherSr1.VDRCRNTTYPE = 'B' and qrytblVoucherSr1.vdt>=cdate('" & Format(w_Fr, "yyyy,mm,dd") & "')  and qrytblVoucherSr1.vdt<=Cdate('" & Format(w_To, "yyyy,mm,dd") & "') "  '
X = "select 'H',qrytblVoucherSr1.VType,tblMastAccount.acname,qrytblVoucherSr1.VCGSTRt+qrytblVoucherSr1.VSGSTRt+qrytblVoucherSr1.VIGSTRt," _
    & " Sum (VTaxableAmt), Sum(VCGSTAmt), Sum(VSGSTAmt), Sum(VIGSTAmt), Sum(abs(VAmt)),'" & gUserName & "'" _
    & " from qrytblVoucherSr1,tblMastAccount,tblMastAccount tblMastAccount_Pty " _
    & " where qrytblVoucherSr1.VAcOCode = tblMastAccount.AcCode and " & X1 & "" _
    & " and qrytblVoucherSr1.VAcCode = tblMastAccount_Pty.AcCode  " _
    & " group by qrytblVoucherSr1.VType,tblMastAccount.acname,qrytblVoucherSr1.VCGSTRt+qrytblVoucherSr1.VSGSTRt+qrytblVoucherSr1.VIGSTRt"

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
           If !Nar1 = "SN" Then
              sheet.Cells(a, 2) = "Indirect Incomes"
           Else
               sheet.Cells(a, 2) = "Indirect Expenses"
           End If
           sheet.Cells(a, 3) = !Nar2 + " On Sales"
           sheet.Cells(a, 4) = "Materials"
           sheet.Cells(a, 5) = "Registered"
           If !Nar1 = "SN" Then
              sheet.Cells(a, 6) = "Debit Note"
           Else
              sheet.Cells(a, 6) = "Credit Note"
           End If
'           sheet.cells(a, 7) = !Nar3
           sheet.Cells(a, 8) = !Amt1
           gCn.BeginTrans
           For i = 1 To 12
               Set sheet = workbook.Sheets.Item("" & Left(MonthName(i), 3) & "")
               X = "select 'H',iif(qrytblVoucherSr1.VType = 'SN','2','3'),tblMastAccount.acname,qrytblVoucherSr1.VCGSTRt+qrytblVoucherSr1.VSGSTRt+qrytblVoucherSr1.VIGSTRt," _
                   & " Sum (VTaxableAmt), Sum(VCGSTAmt), Sum(VSGSTAmt), Sum(VIGSTAmt), Sum(abs(VAmt))," & i & ",'" & gUserName & "'" _
                   & " from qrytblVoucherSr1,tblMastAccount,tblMastAccount tblMastAccount_Pty " _
                   & " where qrytblVoucherSr1.VAcOCode = tblMastAccount.AcCode and " & X1 & " and month(qrytblVoucherSr1.Vdt) = " & i & " and qrytblVoucherSr1.VCGSTRt+qrytblVoucherSr1.VSGSTRt+qrytblVoucherSr1.VIGSTRt = " & tmp1.Fields("Amt1") & " " _
                   & " and qrytblVoucherSr1.VAcCode = tblMastAccount_Pty.AcCode  " _
                   & " and tblMastAccount.AcName = '" & tmp1.Fields("Nar2") & "' " _
                   & " and qrytblVoucherSr1.VType = '" & tmp1.Fields("Nar1") & "'" _
                   & " group by qrytblVoucherSr1.VType,tblMastAccount.acname,qrytblVoucherSr1.VCGSTRt+qrytblVoucherSr1.VSGSTRt+qrytblVoucherSr1.VIGSTRt"
               
               
               gCn.Execute " Insert Into TmpGentbl2 (VType,Nar1,Nar2,Amt1,Amt2,Amt3,Amt4,Amt5,Amt6,Vno,UserName)" & X
               j = GProcRstOpen(tmp2, X, "R", gCn)
               If j > 0 Then
                  sheet.Cells(a, 10) = IIf(IsNull(tmp2.Fields(4)), 0, tmp2.Fields(4))
                  sheet.Cells(a, 11) = IIf(IsNull(tmp2.Fields(5)), 0, tmp2.Fields(5))
                  sheet.Cells(a, 12) = IIf(IsNull(tmp2.Fields(6)), 0, tmp2.Fields(6))
                  sheet.Cells(a, 13) = IIf(IsNull(tmp2.Fields(7)), 0, tmp2.Fields(7))
                  sheet.Cells(a, 14) = 0
               Else
                  sheet.Cells(a, 10) = 0
                  sheet.Cells(a, 11) = 0
                  sheet.Cells(a, 12) = 0
                  sheet.Cells(a, 13) = 0
                  sheet.Cells(a, 14) = 0
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

'--Credit/Debit Note Purchase
gCn.BeginTrans

X1 = " qrytblVoucherSr1.VType in ('SN','PN') and qrytblVoucherSr1.VDRCRNTTYPE = 'S' and qrytblVoucherSr1.vdt>=cdate('" & Format(w_Fr, "yyyy,mm,dd") & "')  and qrytblVoucherSr1.vdt<=Cdate('" & Format(w_To, "yyyy,mm,dd") & "') "  '
X = "select 'I',qrytblVoucherSr1.VType,tblMastAccount.acname,qrytblVoucherSr1.VCGSTRt+qrytblVoucherSr1.VSGSTRt+qrytblVoucherSr1.VIGSTRt," _
    & " Sum (VTaxableAmt), Sum(VCGSTAmt), Sum(VSGSTAmt), Sum(VIGSTAmt), Sum(abs(VAmt)),'" & gUserName & "'" _
    & " from qrytblVoucherSr1,tblMastAccount,tblMastAccount tblMastAccount_Pty " _
    & " where qrytblVoucherSr1.VAcOCode = tblMastAccount.AcCode and " & X1 & "" _
    & " and qrytblVoucherSr1.VAcCode = tblMastAccount_Pty.AcCode  " _
    & " group by qrytblVoucherSr1.VType,tblMastAccount.acname,qrytblVoucherSr1.VCGSTRt+qrytblVoucherSr1.VSGSTRt+qrytblVoucherSr1.VIGSTRt"

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
           If !Nar1 = "SN" Then
              sheet.Cells(a, 2) = "Indirect Incomes"
           Else
               sheet.Cells(a, 2) = "Indirect Expenses"
           End If
           sheet.Cells(a, 3) = !Nar2 + " On Purchase"
           sheet.Cells(a, 4) = "Materials"
           sheet.Cells(a, 5) = "Registered"
           If !Nar1 = "SN" Then
              sheet.Cells(a, 6) = "Debit Note"
           Else
              sheet.Cells(a, 6) = "Credit Note"
           End If
'           sheet.cells(a, 7) = !Nar3
           sheet.Cells(a, 8) = !Amt1
           gCn.BeginTrans
           For i = 1 To 12
               Set sheet = workbook.Sheets.Item("" & Left(MonthName(i), 3) & "")
               X = "select 'I',iif(qrytblVoucherSr1.VType = 'SN','5','6'),tblMastAccount.acname,qrytblVoucherSr1.VCGSTRt+qrytblVoucherSr1.VSGSTRt+qrytblVoucherSr1.VIGSTRt," _
                   & " Sum (VTaxableAmt), Sum(VCGSTAmt), Sum(VSGSTAmt), Sum(VIGSTAmt), Sum(abs(VAmt))," & i & ",'" & gUserName & "'" _
                   & " from qrytblVoucherSr1,tblMastAccount,tblMastAccount tblMastAccount_Pty " _
                   & " where qrytblVoucherSr1.VAcOCode = tblMastAccount.AcCode and " & X1 & " and month(qrytblVoucherSr1.Vdt) = " & i & " and qrytblVoucherSr1.VCGSTRt+qrytblVoucherSr1.VSGSTRt+qrytblVoucherSr1.VIGSTRt = " & tmp1.Fields("Amt1") & " " _
                   & " and qrytblVoucherSr1.VAcCode = tblMastAccount_Pty.AcCode  " _
                   & " and tblMastAccount.AcName = '" & tmp1.Fields("Nar2") & "' " _
                   & " group by qrytblVoucherSr1.VType,tblMastAccount.acname,qrytblVoucherSr1.VCGSTRt+qrytblVoucherSr1.VSGSTRt+qrytblVoucherSr1.VIGSTRt"
                
               gCn.Execute " Insert Into TmpGentbl2 (VType,Nar1,Nar2,Amt1,Amt2,Amt3,Amt4,Amt5,Amt6,Vno,UserName)" & X
               j = GProcRstOpen(tmp2, X, "R", gCn)
               If j > 0 Then
                  sheet.Cells(a, 10) = IIf(IsNull(tmp2.Fields(4)), 0, tmp2.Fields(4))
                  sheet.Cells(a, 11) = IIf(IsNull(tmp2.Fields(5)), 0, tmp2.Fields(5))
                  sheet.Cells(a, 12) = IIf(IsNull(tmp2.Fields(6)), 0, tmp2.Fields(6))
                  sheet.Cells(a, 13) = IIf(IsNull(tmp2.Fields(7)), 0, tmp2.Fields(7))
                  sheet.Cells(a, 14) = 0
               Else
                  sheet.Cells(a, 10) = 0
                  sheet.Cells(a, 11) = 0
                  sheet.Cells(a, 12) = 0
                  sheet.Cells(a, 13) = 0
                  sheet.Cells(a, 14) = 0
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
'--For Total yarn Sale/purchase Retuen R.D.
gCn.BeginTrans

X1 = " tblAddLess.VType in ('RY','VY')  and tblAddLess.vdt>=cdate('" & Format(w_Fr, "yyyy,mm,dd") & "')  and tblAddLess.vdt<=Cdate('" & Format(w_To, "yyyy,mm,dd") & "') and (tblAddLess.AdTmp5 <> 'COL' or isnull(tblAddLess.AdTmp5)) and VsubCGSTAmt+VsubSGSTAmt+VsubIGSTAmt > 0 and (len(tblMastAccount_Pty.AcGSTIN) <= 2 or isnull(tblMastAccount_Pty.AcGSTIN))"
X = "select 'J',tblAddLess.VType,tblMastAccount.acname,tblMastItem.ITTMP2,tblVouSub.VsubCGSTRt+tblVouSub.VsubSGSTRt+tblVouSub.VsubIGSTRt," _
    & " Sum (VSubTaxableAmt), Sum(VsubCGSTAmt), Sum(VsubSGSTAmt), Sum(VsubIGSTAmt), Sum(AdBillAmt),'" & gUserName & "'" _
    & " from tblAddLess,tblVouSub,tblMastItem,tblMastAccount,tblMastAccount tblMastAccount_Pty " _
    & " where tblAddLess.AdAcCrCode = tblMastAccount.AcCode and tblVouSub.VSubItCode = tblMastItem.ItCode and tblAddLess.Vno = tblVouSub.Vno and " _
    & " tblAddLess.Vtype = tblVouSub.Vtype and tblAddLess.Vyear = tblVouSub.Vyear and " & X1 & " " _
    & " and tblAddLess.AdAcDrCode = tblMastAccount_Pty.AcCode  " _
    & " group by tblAddLess.VType,tblMastAccount.acname,tblMastItem.ITTMP2,tblVouSub.VsubCGSTRt+tblVouSub.VsubSGSTRt+tblVouSub.VsubIGSTRt"

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
           sheet.Cells(a, 2) = "Others"
           If !Nar1 = "RY" Then
              sheet.Cells(a, 3) = "Sales Rtn"
           Else
              sheet.Cells(a, 3) = "Purchase Rtn"
           End If
           If Left(!Nar3, 2) = "99" Then
              sheet.Cells(a, 4) = "Services"
           Else
              sheet.Cells(a, 4) = "Goods"
           End If
           sheet.Cells(a, 5) = "Registered"
           sheet.Cells(a, 6) = "Tax Invoice"
           sheet.Cells(a, 7) = !Nar3
           sheet.Cells(a, 8) = !Amt1
           gCn.BeginTrans
           For i = 1 To 12
               Set sheet = workbook.Sheets.Item("" & Left(MonthName(i), 3) & "")
               X = "select 'J',iif(tblAddLess.Vtype = 'RY','3','5'),tblMastAccount.acname,tblMastItem.ITTMP2,tblVouSub.VsubCGSTRt+tblVouSub.VsubSGSTRt+tblVouSub.VsubIGSTRt," _
                   & " Sum (VSubTaxableAmt), Sum(VsubCGSTAmt), Sum(VsubSGSTAmt), Sum(VsubIGSTAmt), Sum(AdBillAmt)," & i & ",'" & gUserName & "'" _
                   & " from tblAddLess,tblVouSub,tblMastItem,tblMastAccount,tblMastAccount tblMastAccount_Pty " _
                   & " where tblAddLess.AdAcCrCode = tblMastAccount.AcCode and tblVouSub.VSubItCode = tblMastItem.ItCode and tblAddLess.Vno = tblVouSub.Vno and " _
                   & " tblAddLess.Vtype = tblVouSub.Vtype and tblAddLess.Vyear = tblVouSub.Vyear and " & X1 & " and month(tblAddLess.Vdt) = " & i & " and iif(isnull(tblMastItem.ITTMP2),isnull(tblMastItem.ITTMP2), tblMastItem.ITTMP2 = '" & tmp1.Fields("Nar3") & "') and tblVouSub.VsubCGSTRt+tblVouSub.VsubSGSTRt+tblVouSub.VsubIGSTRt = " & tmp1.Fields("Amt1") & "" _
                   & " and tblAddLess.AdAcDrCode = tblMastAccount_Pty.AcCode  " _
                   & " and tblMastAccount.AcName = '" & tmp1.Fields("Nar2") & "' " _
                   & " group by tblAddLess.VType,tblMastAccount.acname,tblMastItem.ITTMP2,tblVouSub.VsubCGSTRt+tblVouSub.VsubSGSTRt+tblVouSub.VsubIGSTRt"
                
               gCn.Execute " Insert Into TmpGentbl2 (VType,Nar1,Nar2,Nar3,Amt1,Amt2,Amt3,Amt4,Amt5,Amt6,Vno,UserName)" & X
               j = GProcRstOpen(tmp2, X, "R", gCn)
               If j > 0 Then
                  sheet.Cells(a, 10) = IIf(IsNull(tmp2.Fields(5)), 0, tmp2.Fields(5))
                  sheet.Cells(a, 11) = IIf(IsNull(tmp2.Fields(6)), 0, tmp2.Fields(6))
                  sheet.Cells(a, 12) = IIf(IsNull(tmp2.Fields(7)), 0, tmp2.Fields(7))
                  sheet.Cells(a, 13) = IIf(IsNull(tmp2.Fields(8)), 0, tmp2.Fields(8))
                  sheet.Cells(a, 14) = 0
               Else
                  sheet.Cells(a, 10) = 0
                  sheet.Cells(a, 11) = 0
                  sheet.Cells(a, 12) = 0
                  sheet.Cells(a, 13) = 0
                  sheet.Cells(a, 14) = 0
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

'-sum(tblAddLess.AdBillAmt),sum(tblAddLess.AdLCostAdd),sum(tblAddLess.AdLCostFreight),sum(tblAddLess.AdLCostGRent),sum(tblAddLess.AdBillAmt),'" & gUserName & "'" _



'--For Total RCM Sales
gCn.BeginTrans
X1 = " tblAddLess.VType in ('PO')  and tblAddLess.AdBookDt>=cdate('" & Format(w_Fr, "yyyy,mm,dd") & "')  and tblAddLess.AdBookDt<=Cdate('" & Format(w_To, "yyyy,mm,dd") & "') and tblAddLess.AdLCostAdd+tblAddLess.AdLCostFreight+tblAddLess.AdLCostGRent > 0"
X = "select 'K',tblAddLess.VType,tblMastAccount.acname,tblMastItem.ITTMP2,tblVouSub.VsubCGSTRt+tblVouSub.VsubSGSTRt+tblVouSub.VsubIGSTRt," _
    & " Sum (VSubTaxableAmt),  Sum(AdLCostFreight), Sum(AdLCostGRent),Sum(AdLCostAdd), Sum(AdBillAmt),'" & gUserName & "'" _
    & " from tblAddLess,tblVouSub,tblMastItem,tblMastAccount,tblMastAccount tblMastAccount_Pty " _
    & " where tblAddLess.AdAcDrCode = tblMastAccount.AcCode and tblVouSub.VSubItCode = tblMastItem.ItCode and tblAddLess.Vno = tblVouSub.Vno and " _
    & " tblAddLess.Vtype = tblVouSub.Vtype and tblAddLess.Vyear = tblVouSub.Vyear and " & X1 & " " _
    & " and tblAddLess.AdAcCrCode = tblMastAccount_Pty.AcCode  " _
    & " group by tblAddLess.VType,tblMastAccount.acname,tblMastItem.ITTMP2,tblVouSub.VsubCGSTRt+tblVouSub.VsubSGSTRt+tblVouSub.VsubIGSTRt"

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
           sheet.Cells(a, 2) = "Purchases"
           sheet.Cells(a, 3) = !Nar2
           If Left(!Nar3, 2) = "99" Then
              sheet.Cells(a, 4) = "Services"
           Else
              sheet.Cells(a, 4) = "Goods"
           End If
           sheet.Cells(a, 5) = "Unregistered"
           sheet.Cells(a, 6) = "RCM Invoice"
           sheet.Cells(a, 7) = !Nar3 & vbNullString
           sheet.Cells(a, 8) = !Amt1
           gCn.BeginTrans
           For i = 1 To 12
               Set sheet = workbook.Sheets.Item("" & Left(MonthName(i), 3) & "")
               X = "select 'K','7',tblMastAccount.acname,tblMastItem.ITTMP2,tblVouSub.VsubCGSTRt+tblVouSub.VsubSGSTRt+tblVouSub.VsubIGSTRt," _
                   & " Sum (VSubTaxableAmt), Sum(AdLCostFreight), Sum(AdLCostGRent), Sum(AdLCostAdd), Sum(AdBillAmt)," & i & ",'" & gUserName & "'" _
                   & " from tblAddLess,tblVouSub,tblMastItem,tblMastAccount,tblMastAccount tblMastAccount_Pty " _
                   & " where tblAddLess.AdAcDrCode = tblMastAccount.AcCode and tblVouSub.VSubItCode = tblMastItem.ItCode and tblAddLess.Vno = tblVouSub.Vno and " _
                   & " tblAddLess.Vtype = tblVouSub.Vtype and tblAddLess.Vyear = tblVouSub.Vyear and " & X1 & " and month(tblAddLess.AdBookDt) = " & i & " and iif(isnull(tblMastItem.ITTMP2),isnull(tblMastItem.ITTMP2), tblMastItem.ITTMP2 = '" & tmp1.Fields("Nar3") & "') and tblVouSub.VsubCGSTRt+tblVouSub.VsubSGSTRt+tblVouSub.VsubIGSTRt = " & tmp1.Fields("Amt1") & "" _
                   & " and tblAddLess.AdAcCrCode = tblMastAccount_Pty.AcCode  " _
                   & " and tblMastAccount.AcName = '" & tmp1.Fields("Nar2") & "' " _
                   & " group by tblAddLess.VType,tblMastAccount.acname,tblMastItem.ITTMP2,tblVouSub.VsubCGSTRt+tblVouSub.VsubSGSTRt+tblVouSub.VsubIGSTRt"
                
               gCn.Execute " Insert Into TmpGentbl2 (VType,Nar1,Nar2,Nar3,Amt1,Amt2,Amt3,Amt4,Amt5,Amt6,Vno,UserName)" & X
               
               j = GProcRstOpen(tmp2, X, "R", gCn)
               If j > 0 Then
                  sheet.Cells(a, 10) = IIf(IsNull(tmp2.Fields(5)), 0, tmp2.Fields(5))
                  sheet.Cells(a, 11) = IIf(IsNull(tmp2.Fields(6)), 0, tmp2.Fields(6))
                  sheet.Cells(a, 12) = IIf(IsNull(tmp2.Fields(7)), 0, tmp2.Fields(7))
                  sheet.Cells(a, 13) = IIf(IsNull(tmp2.Fields(8)), 0, tmp2.Fields(8))
                  sheet.Cells(a, 14) = 0
               Else
                  sheet.Cells(a, 10) = 0
                  sheet.Cells(a, 11) = 0
                  sheet.Cells(a, 12) = 0
                  sheet.Cells(a, 13) = 0
                  sheet.Cells(a, 14) = 0
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

'= For Working Sheet
Set sheet = workbook.Sheets.Item("Working")
If (sheet Is Nothing) Then
    MsgBox "Error Working Sheet"
End If
Dim ColTaxable As Integer
Dim ColIgst As Integer
Dim ColCgst As Integer
Dim ColSgst As Integer

For i = 1 To 12
    If i >= 4 Then
       ColTaxable = 3 + (6 * (i - 4))
       ColIgst = 4 + (6 * (i - 4))
       ColCgst = 5 + (6 * (i - 4))
       ColSgst = 6 + (6 * (i - 4))
    Else
       ColTaxable = 57 + (6 * (i - 1))
       ColIgst = 58 + (6 * (i - 1))
       ColCgst = 59 + (6 * (i - 1))
       ColSgst = 60 + (6 * (i - 1))
    End If
    RecYN = GProcRstOpen(tmp2, "select nar1,sum(amt2),sum(amt5),sum(amt3),sum(amt4) from TmpGentbl2 where TmpGentbl2.UserName ='" & gUserName & "' and TmpGentbl2.vno = " & i & " group by Nar1   ", "R", gCn)
    If RecYN > 0 Then
       With tmp2
            .MoveFirst
            Do While Not .EOF
            
                    Select Case !Nar1
                    Case "1"  '---Sales
                       sheet.Cells(6, ColTaxable) = tmp2.Fields(1)  'Taxable
                       sheet.Cells(6, ColIgst) = tmp2.Fields(2)  'Igst
                       sheet.Cells(6, ColCgst) = tmp2.Fields(3)  'Cgst
                       sheet.Cells(6, ColSgst) = tmp2.Fields(4)  'Igst
                    
                       sheet.Cells(11, ColTaxable) = tmp2.Fields(1)  'Taxable
                       sheet.Cells(11, ColIgst) = tmp2.Fields(2)  'Igst
                       sheet.Cells(11, ColCgst) = tmp2.Fields(3)  'Cgst
                       sheet.Cells(11, ColSgst) = tmp2.Fields(4)  'Igst
                    
                       sheet.Cells(22, ColTaxable) = tmp2.Fields(1)  'Taxable
                       sheet.Cells(22, ColIgst) = tmp2.Fields(2)  'Igst
                       sheet.Cells(22, ColCgst) = tmp2.Fields(3)  'Cgst
                       sheet.Cells(22, ColSgst) = tmp2.Fields(4)  'Igst
                    Case "2"  '---Db.Note Sales
                       sheet.Cells(12, ColTaxable) = tmp2.Fields(1)  'Taxable
                       sheet.Cells(12, ColIgst) = tmp2.Fields(2)  'Igst
                       sheet.Cells(12, ColCgst) = tmp2.Fields(3)  'Cgst
                       sheet.Cells(12, ColSgst) = tmp2.Fields(4)  'Igst
                    
                       sheet.Cells(23, ColTaxable) = tmp2.Fields(1)  'Taxable
                       sheet.Cells(23, ColIgst) = tmp2.Fields(2)  'Igst
                       sheet.Cells(23, ColCgst) = tmp2.Fields(3)  'Cgst
                       sheet.Cells(23, ColSgst) = tmp2.Fields(4)  'Igst
                    Case "3"  '---Cr.Note Sales
                       sheet.Cells(13, ColTaxable) = tmp2.Fields(1)  'Taxable
                       sheet.Cells(13, ColIgst) = tmp2.Fields(2)  'Igst
                       sheet.Cells(13, ColCgst) = tmp2.Fields(3)  'Cgst
                       sheet.Cells(13, ColSgst) = tmp2.Fields(4)  'Igst
                    
                       sheet.Cells(24, ColTaxable) = tmp2.Fields(1)  'Taxable
                       sheet.Cells(24, ColIgst) = tmp2.Fields(2)  'Igst
                       sheet.Cells(24, ColCgst) = tmp2.Fields(3)  'Cgst
                       sheet.Cells(24, ColSgst) = tmp2.Fields(4)  'Igst
                    Case "4"  '---Purchase
                       sheet.Cells(7, ColTaxable) = tmp2.Fields(1)  'Taxable
                       sheet.Cells(7, ColIgst) = tmp2.Fields(2)  'Igst
                       sheet.Cells(7, ColCgst) = tmp2.Fields(3)  'Cgst
                       sheet.Cells(7, ColSgst) = tmp2.Fields(4)  'Igst
                    
                       sheet.Cells(25, ColTaxable) = tmp2.Fields(1)  'Taxable
                       sheet.Cells(25, ColIgst) = tmp2.Fields(2)  'Igst
                       sheet.Cells(25, ColCgst) = tmp2.Fields(3)  'Cgst
                       sheet.Cells(25, ColSgst) = tmp2.Fields(4)  'Igst
                    Case "5"  '---Db. Note Purchase
                       sheet.Cells(26, ColTaxable) = tmp2.Fields(1)  'Taxable
                       sheet.Cells(26, ColIgst) = tmp2.Fields(2)  'Igst
                       sheet.Cells(26, ColCgst) = tmp2.Fields(3)  'Cgst
                       sheet.Cells(26, ColSgst) = tmp2.Fields(4)  'Igst
                    Case "6"  '---Cr. Note Purchase
                       sheet.Cells(27, ColTaxable) = tmp2.Fields(1)  'Taxable
                       sheet.Cells(27, ColIgst) = tmp2.Fields(2)  'Igst
                       sheet.Cells(27, ColCgst) = tmp2.Fields(3)  'Cgst
                       sheet.Cells(27, ColSgst) = tmp2.Fields(4)  'Igst
                    Case "7"  '---RCM
                       sheet.Cells(28, ColTaxable) = tmp2.Fields(1)  'Taxable
                       sheet.Cells(28, ColIgst) = tmp2.Fields(2)  'Igst
                       sheet.Cells(28, ColCgst) = tmp2.Fields(3)  'Cgst
                       sheet.Cells(28, ColSgst) = tmp2.Fields(4)  'Igst
                    End Select
               tmp2.MoveNext
            Loop
       End With
    End If
Next i



workbook.Close SaveChanges:=True
excel_app.Quit
MsgBox "Compileted"
Set excel_app = Nothing
End Function
Public Function GProcPurpareHSNSummery(w_Fr As Date, w_To As Date, Optional IsOnPurBillDt As Boolean)
Dim X1 As String
gCn.BeginTrans
gCn.Execute "delete from tmpGenTbl where username='" & gUserName & "'"
'-- Yarn Sales
'--Ms
X1 = " tblAddLess.VType in ('SY') and tblAddLess.vdt>=cdate('" & Format(w_Fr, "yyyy,mm,dd") & "')  and tblAddLess.vdt<=Cdate('" & Format(w_To, "yyyy,mm,dd") & "')  and tblVouSub.VsubCGSTAmt+tblVouSub.VsubSGSTAmt > 0 and isnull(tblAddLess.AdTmp5) = True "
X1 = "select 'A',tblMastAccount.AcName,tblMastItem.ITTMP2,'M.S.',tblVouSub.Vno,tblVouSub.vtype,tblVouSub.vyear,tblVouSub.VSubCtrNo,tblVouSub.VsubCGSTRt+tblVouSub.VsubSGSTRt,sum(tblVouSub.VSubAmt),sum(tblVouSub.VsubIGSTAmt),sum(tblVouSub.VsubCGSTAmt),sum(tblVouSub.VsubSGSTAmt),sum(tblVouSub.VSubBag),'" & gUserName & "' from tblAddLess,tblVouSub,tblMastAccount,tblMastItem where " & X1 & " and " _
     & " tblAddLess.Vno = tblVouSub.Vno and tblAddLess.Vtype = tblVouSub.Vtype and tblAddLess.Vyear = tblVouSub.Vyear and tblVouSub.VSubItCode = tblMastItem.ItCode and tblAddLess.AdAcCrCode = tblMastAccount.Accode group by tblMastAccount.AcName,tblMastItem.ITTMP2,tblVouSub.Vno,tblVouSub.vtype,tblVouSub.vyear,tblVouSub.VSubCtrNo,tblVouSub.VsubCGSTRt+tblVouSub.VsubSGSTRt"
gCn.Execute " Insert Into TmpGentbl (VBILLNO,Nar1,Nar2,Nar3,Vno,Vtype,Vyear,VSRNO,Amt1,Amt2,Amt3,Amt4,Amt5,Amt6,UserName)" & X1
'--OMS
X1 = " tblAddLess.VType in ('SY') and tblAddLess.vdt>=cdate('" & Format(w_Fr, "yyyy,mm,dd") & "')  and tblAddLess.vdt<=Cdate('" & Format(w_To, "yyyy,mm,dd") & "')  and tblVouSub.VsubIGSTAmt > 0 and isnull(tblAddLess.AdTmp5) = True "
X1 = "select 'A',tblMastAccount.AcName,tblMastItem.ITTMP2,'O.M.S.',tblVouSub.Vno,tblVouSub.vtype,tblVouSub.vyear,tblVouSub.VSubCtrNo,tblVouSub.VsubIGSTRt,sum(tblVouSub.VSubAmt),sum(tblVouSub.VsubIGSTAmt),sum(tblVouSub.VsubCGSTAmt),sum(tblVouSub.VsubSGSTAmt),sum(tblVouSub.VSubBag),'" & gUserName & "' from tblAddLess,tblVouSub,tblMastAccount,tblMastItem where " & X1 & " and " _
     & " tblAddLess.Vno = tblVouSub.Vno and tblAddLess.Vtype = tblVouSub.Vtype and tblAddLess.Vyear = tblVouSub.Vyear and tblVouSub.VSubItCode = tblMastItem.ItCode and tblAddLess.AdAcCrCode = tblMastAccount.Accode group by tblMastAccount.AcName,tblMastItem.ITTMP2,tblVouSub.Vno,tblVouSub.vtype,tblVouSub.vyear,tblVouSub.VSubCtrNo,tblVouSub.VsubIGSTRt"
gCn.Execute " Insert Into TmpGentbl (VBILLNO,Nar1,Nar2,Nar3,Vno,Vtype,Vyear,VSRNO,Amt1,Amt2,Amt3,Amt4,Amt5,Amt6,UserName)" & X1

'--Other Sales
'--Ms
X1 = " tblAddLess.VType in ('SY') and tblAddLess.vdt>=cdate('" & Format(w_Fr, "yyyy,mm,dd") & "')  and tblAddLess.vdt<=Cdate('" & Format(w_To, "yyyy,mm,dd") & "')  and tblVouSub.VsubCGSTAmt+tblVouSub.VsubSGSTAmt > 0 and tblAddLess.AdTmp5 = 'OTH' "
X1 = "select 'B',tblMastAccount.AcName,tblMastItem.ITTMP2,'M.S.',tblVouSub.Vno,tblVouSub.vtype,tblVouSub.vyear,tblVouSub.VSubCtrNo,tblVouSub.VsubCGSTRt+tblVouSub.VsubSGSTRt,sum(tblVouSub.VSubAmt),sum(tblVouSub.VsubIGSTAmt),sum(tblVouSub.VsubCGSTAmt),sum(tblVouSub.VsubSGSTAmt),sum(tblVouSub.VSubBag),'" & gUserName & "' from tblAddLess,tblVouSub,tblMastAccount,tblMastItem where " & X1 & " and " _
     & " tblAddLess.Vno = tblVouSub.Vno and tblAddLess.Vtype = tblVouSub.Vtype and tblAddLess.Vyear = tblVouSub.Vyear and tblVouSub.VSubItCode = tblMastItem.ItCode and tblAddLess.AdAcCrCode = tblMastAccount.Accode group by tblMastAccount.AcName,tblMastItem.ITTMP2,tblVouSub.Vno,tblVouSub.vtype,tblVouSub.vyear,tblVouSub.VSubCtrNo,tblVouSub.VsubCGSTRt+tblVouSub.VsubSGSTRt"
gCn.Execute " Insert Into TmpGentbl (VBILLNO,Nar1,Nar2,Nar3,Vno,Vtype,Vyear,VSRNO,Amt1,Amt2,Amt3,Amt4,Amt5,Amt6,UserName)" & X1
'--OMS
X1 = " tblAddLess.VType in ('SY') and tblAddLess.vdt>=cdate('" & Format(w_Fr, "yyyy,mm,dd") & "')  and tblAddLess.vdt<=Cdate('" & Format(w_To, "yyyy,mm,dd") & "')  and tblVouSub.VsubIGSTAmt > 0 and tblAddLess.AdTmp5 = 'OTH' "
X1 = "select 'B',tblMastAccount.AcName,tblMastItem.ITTMP2,'O.M.S.',tblVouSub.Vno,tblVouSub.vtype,tblVouSub.vyear,tblVouSub.VSubCtrNo,tblVouSub.VsubIGSTRt,sum(tblVouSub.VSubAmt),sum(tblVouSub.VsubIGSTAmt),sum(tblVouSub.VsubCGSTAmt),sum(tblVouSub.VsubSGSTAmt),sum(tblVouSub.VSubBag),'" & gUserName & "' from tblAddLess,tblVouSub,tblMastAccount,tblMastItem where " & X1 & " and " _
     & " tblAddLess.Vno = tblVouSub.Vno and tblAddLess.Vtype = tblVouSub.Vtype and tblAddLess.Vyear = tblVouSub.Vyear and tblVouSub.VSubItCode = tblMastItem.ItCode and tblAddLess.AdAcCrCode = tblMastAccount.Accode group by tblMastAccount.AcName,tblMastItem.ITTMP2,tblVouSub.Vno,tblVouSub.vtype,tblVouSub.vyear,tblVouSub.VSubCtrNo,tblVouSub.VsubIGSTRt"
gCn.Execute " Insert Into TmpGentbl (VBILLNO,Nar1,Nar2,Nar3,Vno,Vtype,Vyear,VSRNO,Amt1,Amt2,Amt3,Amt4,Amt5,Amt6,UserName)" & X1

''--Cloth Sales
''--Ms
'X1 = " tblAddLess.VType in ('SY') and tblAddLess.vdt>=cdate('" & Format(w_Fr, "yyyy,mm,dd") & "')  and tblAddLess.vdt<=Cdate('" & Format(w_To, "yyyy,mm,dd") & "')  and tblAddLess.AdClothCgstRt+tblAddLess.AdClothSgstRt > 0 and (tblAddLess.AdTmp5 = 'COL' or tblAddLess.AdTmp5 = 'CLO') "
'X1 = "select 'C',tblMastAccount.AcName,tblMastItem.ITTMP2,'M.S.',tblVouSub.Vno,tblVouSub.vtype,tblVouSub.vyear,tblVouSub.VSubCtrNo,tblAddLess.AdClothCgstRt+tblAddLess.AdClothSgstRt,sum(iif(tblVouSub.VSubCtrNo = 1,tblAddLess.AdAssVal,0)),sum(iif(tblVouSub.VSubCtrNo = 1,AdClothIgstAmt,0)),sum(iif(tblVouSub.VSubCtrNo = 1,AdClothCgstAmt,0)),sum(iif(tblVouSub.VSubCtrNo = 1,AdClothSgstAmt,0)),sum(tblVouSub.VSubWt),'" & gUserName & "' from tblAddLess,tblVouSub,tblMastAccount,tblMastItem where " & X1 & " and " _
'     & " tblAddLess.Vno = tblVouSub.Vno and tblAddLess.Vtype = tblVouSub.Vtype and tblAddLess.Vyear = tblVouSub.Vyear and tblVouSub.VSubItCode = tblMastItem.ItCode and tblAddLess.AdAcCrCode = tblMastAccount.Accode group by tblMastAccount.AcName,tblMastItem.ITTMP2,tblVouSub.Vno,tblVouSub.vtype,tblVouSub.vyear,tblVouSub.VSubCtrNo,tblAddLess.AdClothCgstRt+tblAddLess.AdClothSgstRt"
'gCn.Execute " Insert Into TmpGentbl (VBILLNO,Nar1,Nar2,Nar3,Vno,Vtype,Vyear,VSRNO,Amt1,Amt2,Amt3,Amt4,Amt5,Amt6,UserName)" & X1
''--OMS
'X1 = " tblAddLess.VType in ('SY') and tblAddLess.vdt>=cdate('" & Format(w_Fr, "yyyy,mm,dd") & "')  and tblAddLess.vdt<=Cdate('" & Format(w_To, "yyyy,mm,dd") & "')  and tblAddLess.AdClothIgstRt > 0 and (tblAddLess.AdTmp5 = 'COL' or tblAddLess.AdTmp5 = 'CLO') "
'X1 = "select 'C',tblMastAccount.AcName,tblMastItem.ITTMP2,'O.M.S.',tblVouSub.Vno,tblVouSub.vtype,tblVouSub.vyear,tblVouSub.VSubCtrNo,tblAddLess.AdClothIgstRt,sum(iif(tblVouSub.VSubCtrNo = 1,tblAddLess.AdAssVal,0)),sum(iif(tblVouSub.VSubCtrNo = 1,AdClothIgstAmt,0)),sum(iif(tblVouSub.VSubCtrNo = 1,AdClothCgstAmt,0)),sum(iif(tblVouSub.VSubCtrNo = 1,AdClothSgstAmt,0)),sum(tblVouSub.VSubWt),'" & gUserName & "' from tblAddLess,tblVouSub,tblMastAccount,tblMastItem where " & X1 & " and " _
'     & " tblAddLess.Vno = tblVouSub.Vno and tblAddLess.Vtype = tblVouSub.Vtype and tblAddLess.Vyear = tblVouSub.Vyear and tblVouSub.VSubItCode = tblMastItem.ItCode and tblAddLess.AdAcCrCode = tblMastAccount.Accode group by tblMastAccount.AcName,tblMastItem.ITTMP2,tblVouSub.Vno,tblVouSub.vtype,tblVouSub.vyear,tblVouSub.VSubCtrNo,tblAddLess.AdClothIgstRt"
'gCn.Execute " Insert Into TmpGentbl (VBILLNO,Nar1,Nar2,Nar3,Vno,Vtype,Vyear,VSRNO,Amt1,Amt2,Amt3,Amt4,Amt5,Amt6,UserName)" & X1

'--Debit Note  Sales
'--Ms

X1 = " qrytblVoucherSr1.VType = 'SN' and qrytblVoucherSr1.VDRCRNTTYPE = 'B' and qrytblVoucherSr1.vdt>=cdate('" & Format(w_Fr, "yyyy,mm,dd") & "')  and qrytblVoucherSr1.vdt<=Cdate('" & Format(w_To, "yyyy,mm,dd") & "') and (qrytblVoucherSr1.VCGSTAmt+qrytblVoucherSr1.VSGSTAmt) > 0 "  '

X1 = "select 'D',tblMastAccount.AcName,'.','M.S.',qrytblVoucherSr1.Vno,qrytblVoucherSr1.vtype,qrytblVoucherSr1.vyear,qrytblVoucherSr1.VCtrNo,qrytblVoucherSr1.VCGSTRt+qrytblVoucherSr1.VSGSTRt,sum(qrytblVoucherSr1.VTaxableAmt),sum(qrytblVoucherSr1.VIGSTAmt),sum(qrytblVoucherSr1.VCGSTAmt),sum(qrytblVoucherSr1.VSGSTAmt),sum(qrytblVoucherSr1.VCtrNo),'" & gUserName & "' from qrytblVoucherSr1,tblMastAccount where " & X1 & " and " _
     & " qrytblVoucherSr1.VAcOCode = tblMastAccount.Accode group by tblMastAccount.AcName,qrytblVoucherSr1.Vno,qrytblVoucherSr1.vtype,qrytblVoucherSr1.vyear,qrytblVoucherSr1.VCtrNo,qrytblVoucherSr1.VCGSTRt+qrytblVoucherSr1.VSGSTRt"
gCn.Execute " Insert Into TmpGentbl (VBILLNO,Nar1,Nar2,Nar3,Vno,Vtype,Vyear,VSRNO,Amt1,Amt2,Amt3,Amt4,Amt5,Amt6,UserName)" & X1

'--OMs
X1 = " qrytblVoucherSr1.VType = 'SN' and qrytblVoucherSr1.VDRCRNTTYPE = 'B' and qrytblVoucherSr1.vdt>=cdate('" & Format(w_Fr, "yyyy,mm,dd") & "')  and qrytblVoucherSr1.vdt<=Cdate('" & Format(w_To, "yyyy,mm,dd") & "') and (qrytblVoucherSr1.VIGSTAmt) > 0 "  '

X1 = "select 'D',tblMastAccount.AcName,'.','O.M.S.',qrytblVoucherSr1.Vno,qrytblVoucherSr1.vtype,qrytblVoucherSr1.vyear,qrytblVoucherSr1.VCtrNo,qrytblVoucherSr1.VIGSTRt,sum(qrytblVoucherSr1.VTaxableAmt),sum(qrytblVoucherSr1.VIGSTAmt),sum(qrytblVoucherSr1.VCGSTAmt),sum(qrytblVoucherSr1.VSGSTAmt),sum(qrytblVoucherSr1.VCtrNo),'" & gUserName & "' from qrytblVoucherSr1,tblMastAccount where " & X1 & " and " _
     & " qrytblVoucherSr1.VAcOCode = tblMastAccount.Accode group by tblMastAccount.AcName,qrytblVoucherSr1.Vno,qrytblVoucherSr1.vtype,qrytblVoucherSr1.vyear,qrytblVoucherSr1.VCtrNo,qrytblVoucherSr1.VIGSTRt"
gCn.Execute " Insert Into TmpGentbl (VBILLNO,Nar1,Nar2,Nar3,Vno,Vtype,Vyear,VSRNO,Amt1,Amt2,Amt3,Amt4,Amt5,Amt6,UserName)" & X1


'--Credit Note  Sales
'--Ms
X1 = " qrytblVoucherSr1.VType = 'PN' and qrytblVoucherSr1.VDRCRNTTYPE = 'B' and qrytblVoucherSr1.vdt>=cdate('" & Format(w_Fr, "yyyy,mm,dd") & "')  and qrytblVoucherSr1.vdt<=Cdate('" & Format(w_To, "yyyy,mm,dd") & "') and (qrytblVoucherSr1.VCGSTAmt+qrytblVoucherSr1.VSGSTAmt) > 0 "  '
X1 = "select 'E',tblMastAccount.AcName,'.','M.S.',qrytblVoucherSr1.Vno,qrytblVoucherSr1.vtype,qrytblVoucherSr1.vyear,qrytblVoucherSr1.VCtrNo,qrytblVoucherSr1.VCGSTRt+qrytblVoucherSr1.VSGSTRt,sum(qrytblVoucherSr1.VTaxableAmt),sum(qrytblVoucherSr1.VIGSTAmt),sum(qrytblVoucherSr1.VCGSTAmt),sum(qrytblVoucherSr1.VSGSTAmt),sum(qrytblVoucherSr1.VCtrNo),'" & gUserName & "' from qrytblVoucherSr1,tblMastAccount where " & X1 & " and " _
     & " qrytblVoucherSr1.VAcOCode = tblMastAccount.Accode group by tblMastAccount.AcName,qrytblVoucherSr1.Vno,qrytblVoucherSr1.vtype,qrytblVoucherSr1.vyear,qrytblVoucherSr1.VCtrNo,qrytblVoucherSr1.VCGSTRt+qrytblVoucherSr1.VSGSTRt"
gCn.Execute " Insert Into TmpGentbl (VBILLNO,Nar1,Nar2,Nar3,Vno,Vtype,Vyear,VSRNO,Amt1,Amt2,Amt3,Amt4,Amt5,Amt6,UserName)" & X1
'--OMs
X1 = " qrytblVoucherSr1.VType = 'PN' and qrytblVoucherSr1.VDRCRNTTYPE = 'B' and qrytblVoucherSr1.vdt>=cdate('" & Format(w_Fr, "yyyy,mm,dd") & "')  and qrytblVoucherSr1.vdt<=Cdate('" & Format(w_To, "yyyy,mm,dd") & "') and (qrytblVoucherSr1.VIGSTAmt) > 0 "  '
X1 = "select 'E',tblMastAccount.AcName,'.','O.M.S.',qrytblVoucherSr1.Vno,qrytblVoucherSr1.vtype,qrytblVoucherSr1.vyear,qrytblVoucherSr1.VCtrNo,qrytblVoucherSr1.VIGSTRt,sum(qrytblVoucherSr1.VTaxableAmt),sum(qrytblVoucherSr1.VIGSTAmt),sum(qrytblVoucherSr1.VCGSTAmt),sum(qrytblVoucherSr1.VSGSTAmt),sum(qrytblVoucherSr1.VCtrNo),'" & gUserName & "' from qrytblVoucherSr1,tblMastAccount where " & X1 & " and " _
     & " qrytblVoucherSr1.VAcOCode = tblMastAccount.Accode group by tblMastAccount.AcName,qrytblVoucherSr1.Vno,qrytblVoucherSr1.vtype,qrytblVoucherSr1.vyear,qrytblVoucherSr1.VCtrNo,qrytblVoucherSr1.VIGSTRt"
gCn.Execute " Insert Into TmpGentbl (VBILLNO,Nar1,Nar2,Nar3,Vno,Vtype,Vyear,VSRNO,Amt1,Amt2,Amt3,Amt4,Amt5,Amt6,UserName)" & X1

'--Sales Return
'--Ms
X1 = " tblAddLess.VType in ('RY') and tblAddLess.vdt>=cdate('" & Format(w_Fr, "yyyy,mm,dd") & "')  and tblAddLess.vdt<=Cdate('" & Format(w_To, "yyyy,mm,dd") & "')  and tblVouSub.VsubCGSTAmt+tblVouSub.VsubSGSTAmt > 0 and isnull(tblAddLess.AdTmp5) = True "
X1 = "select 'F',tblMastAccount.AcName,tblMastItem.ITTMP2,'M.S.',tblVouSub.Vno,tblVouSub.vtype,tblVouSub.vyear,tblVouSub.VSubCtrNo,tblVouSub.VsubCGSTRt+tblVouSub.VsubSGSTRt,sum(tblVouSub.VSubAmt),sum(tblVouSub.VsubIGSTAmt),sum(tblVouSub.VsubCGSTAmt),sum(tblVouSub.VsubSGSTAmt),sum(tblVouSub.VSubBag),'" & gUserName & "' from tblAddLess,tblVouSub,tblMastAccount,tblMastItem where " & X1 & " and " _
     & " tblAddLess.Vno = tblVouSub.Vno and tblAddLess.Vtype = tblVouSub.Vtype and tblAddLess.Vyear = tblVouSub.Vyear and tblVouSub.VSubItCode = tblMastItem.ItCode and tblAddLess.AdAcCrCode = tblMastAccount.Accode group by tblMastAccount.AcName,tblMastItem.ITTMP2,tblVouSub.Vno,tblVouSub.vtype,tblVouSub.vyear,tblVouSub.VSubCtrNo,tblVouSub.VsubCGSTRt+tblVouSub.VsubSGSTRt"
gCn.Execute " Insert Into TmpGentbl (VBILLNO,Nar1,Nar2,Nar3,Vno,Vtype,Vyear,VSRNO,Amt1,Amt2,Amt3,Amt4,Amt5,Amt6,UserName)" & X1
'--OMS
X1 = " tblAddLess.VType in ('RY') and tblAddLess.vdt>=cdate('" & Format(w_Fr, "yyyy,mm,dd") & "')  and tblAddLess.vdt<=Cdate('" & Format(w_To, "yyyy,mm,dd") & "')  and tblVouSub.VsubIGSTAmt > 0 and isnull(tblAddLess.AdTmp5) = True "
X1 = "select 'F',tblMastAccount.AcName,tblMastItem.ITTMP2,'O.M.S.',tblVouSub.Vno,tblVouSub.vtype,tblVouSub.vyear,tblVouSub.VSubCtrNo,tblVouSub.VsubIGSTRt,sum(tblVouSub.VSubAmt),sum(tblVouSub.VsubIGSTAmt),sum(tblVouSub.VsubCGSTAmt),sum(tblVouSub.VsubSGSTAmt),sum(tblVouSub.VSubBag),'" & gUserName & "' from tblAddLess,tblVouSub,tblMastAccount,tblMastItem where " & X1 & " and " _
     & " tblAddLess.Vno = tblVouSub.Vno and tblAddLess.Vtype = tblVouSub.Vtype and tblAddLess.Vyear = tblVouSub.Vyear and tblVouSub.VSubItCode = tblMastItem.ItCode and tblAddLess.AdAcCrCode = tblMastAccount.Accode group by tblMastAccount.AcName,tblMastItem.ITTMP2,tblVouSub.Vno,tblVouSub.vtype,tblVouSub.vyear,tblVouSub.VSubCtrNo,tblVouSub.VsubIGSTRt"
gCn.Execute " Insert Into TmpGentbl (VBILLNO,Nar1,Nar2,Nar3,Vno,Vtype,Vyear,VSRNO,Amt1,Amt2,Amt3,Amt4,Amt5,Amt6,UserName)" & X1


'-- Yarn Purchase
If IsOnPurBillDt = True Then
    '--Ms
    X1 = " tblAddLess.VType in ('PY') and tblAddLess.AdGpDt>=cdate('" & Format(w_Fr, "yyyy,mm,dd") & "')  and tblAddLess.AdGpDt<=Cdate('" & Format(w_To, "yyyy,mm,dd") & "')  and tblVouSub.VsubCGSTAmt+tblVouSub.VsubSGSTAmt > 0 and isnull(tblAddLess.AdTmp5) = True "
    X1 = "select 'G',tblMastAccount.AcName,tblMastItem.ITTMP2,'M.S.',tblVouSub.Vno,tblVouSub.vtype,tblVouSub.vyear,tblVouSub.VSubCtrNo,tblVouSub.VsubCGSTRt+tblVouSub.VsubSGSTRt,sum(tblVouSub.VSubTaxableAmt),sum(tblVouSub.VsubIGSTAmt),sum(tblVouSub.VsubCGSTAmt),sum(tblVouSub.VsubSGSTAmt),sum(tblVouSub.VSubBag),'" & gUserName & "' from tblAddLess,tblVouSub,tblMastAccount,tblMastItem where " & X1 & " and " _
         & " tblAddLess.Vno = tblVouSub.Vno and tblAddLess.Vtype = tblVouSub.Vtype and tblAddLess.Vyear = tblVouSub.Vyear and tblVouSub.VSubItCode = tblMastItem.ItCode and tblAddLess.AdAcDrCode = tblMastAccount.Accode group by tblMastAccount.AcName,tblMastItem.ITTMP2,tblVouSub.Vno,tblVouSub.vtype,tblVouSub.vyear,tblVouSub.VSubCtrNo,tblVouSub.VsubCGSTRt+tblVouSub.VsubSGSTRt"
    gCn.Execute " Insert Into TmpGentbl (VBILLNO,Nar1,Nar2,Nar3,Vno,Vtype,Vyear,VSRNO,Amt1,Amt2,Amt3,Amt4,Amt5,Amt6,UserName)" & X1
    '--OMS
    X1 = " tblAddLess.VType in ('PY') and tblAddLess.AdGpDt>=cdate('" & Format(w_Fr, "yyyy,mm,dd") & "')  and tblAddLess.AdGpDt<=Cdate('" & Format(w_To, "yyyy,mm,dd") & "')  and tblVouSub.VsubIGSTAmt > 0 and isnull(tblAddLess.AdTmp5) = True "
    X1 = "select 'G',tblMastAccount.AcName,tblMastItem.ITTMP2,'O.M.S.',tblVouSub.Vno,tblVouSub.vtype,tblVouSub.vyear,tblVouSub.VSubCtrNo,tblVouSub.VsubIGSTRt,sum(tblVouSub.VSubTaxableAmt),sum(tblVouSub.VsubIGSTAmt),sum(tblVouSub.VsubCGSTAmt),sum(tblVouSub.VsubSGSTAmt),sum(tblVouSub.VSubBag),'" & gUserName & "' from tblAddLess,tblVouSub,tblMastAccount,tblMastItem where " & X1 & " and " _
         & " tblAddLess.Vno = tblVouSub.Vno and tblAddLess.Vtype = tblVouSub.Vtype and tblAddLess.Vyear = tblVouSub.Vyear and tblVouSub.VSubItCode = tblMastItem.ItCode and tblAddLess.AdAcDrCode = tblMastAccount.Accode group by tblMastAccount.AcName,tblMastItem.ITTMP2,tblVouSub.Vno,tblVouSub.vtype,tblVouSub.vyear,tblVouSub.VSubCtrNo,tblVouSub.VsubIGSTRt"
    gCn.Execute " Insert Into TmpGentbl (VBILLNO,Nar1,Nar2,Nar3,Vno,Vtype,Vyear,VSRNO,Amt1,Amt2,Amt3,Amt4,Amt5,Amt6,UserName)" & X1
    
    '-- Other Purchase
    '--Ms
    X1 = " tblAddLess.VType in ('PO') and tblAddLess.AdGpDt>=cdate('" & Format(w_Fr, "yyyy,mm,dd") & "')  and tblAddLess.AdGpDt<=Cdate('" & Format(w_To, "yyyy,mm,dd") & "')  and tblVouSub.VsubCGSTAmt+tblVouSub.VsubSGSTAmt > 0 and isnull(tblAddLess.AdTmp5) = True "
    X1 = "select 'H',tblMastAccount.AcName,tblMastItem.ITTMP2,'M.S.',tblVouSub.Vno,tblVouSub.vtype,tblVouSub.vyear,tblVouSub.VSubCtrNo,tblVouSub.VsubCGSTRt+tblVouSub.VsubSGSTRt,sum(tblVouSub.VSubTaxableAmt),sum(tblVouSub.VsubIGSTAmt),sum(tblVouSub.VsubCGSTAmt),sum(tblVouSub.VsubSGSTAmt),sum(tblVouSub.VSubBag),'" & gUserName & "' from tblAddLess,tblVouSub,tblMastAccount,tblMastItem where " & X1 & " and " _
         & " tblAddLess.Vno = tblVouSub.Vno and tblAddLess.Vtype = tblVouSub.Vtype and tblAddLess.Vyear = tblVouSub.Vyear and tblVouSub.VSubItCode = tblMastItem.ItCode and tblAddLess.AdAcDrCode = tblMastAccount.Accode group by tblMastAccount.AcName,tblMastItem.ITTMP2,tblVouSub.Vno,tblVouSub.vtype,tblVouSub.vyear,tblVouSub.VSubCtrNo,tblVouSub.VsubCGSTRt+tblVouSub.VsubSGSTRt"
    gCn.Execute " Insert Into TmpGentbl (VBILLNO,Nar1,Nar2,Nar3,Vno,Vtype,Vyear,VSRNO,Amt1,Amt2,Amt3,Amt4,Amt5,Amt6,UserName)" & X1
    '--OMS
    X1 = " tblAddLess.VType in ('PO') and tblAddLess.AdGpDt>=cdate('" & Format(w_Fr, "yyyy,mm,dd") & "')  and tblAddLess.AdGpDt<=Cdate('" & Format(w_To, "yyyy,mm,dd") & "')  and tblVouSub.VsubIGSTAmt > 0 and isnull(tblAddLess.AdTmp5) = True "
    X1 = "select 'H',tblMastAccount.AcName,tblMastItem.ITTMP2,'O.M.S.',tblVouSub.Vno,tblVouSub.vtype,tblVouSub.vyear,tblVouSub.VSubCtrNo,tblVouSub.VsubIGSTRt,sum(tblVouSub.VSubTaxableAmt),sum(tblVouSub.VsubIGSTAmt),sum(tblVouSub.VsubCGSTAmt),sum(tblVouSub.VsubSGSTAmt),sum(tblVouSub.VSubBag),'" & gUserName & "' from tblAddLess,tblVouSub,tblMastAccount,tblMastItem where " & X1 & " and " _
         & " tblAddLess.Vno = tblVouSub.Vno and tblAddLess.Vtype = tblVouSub.Vtype and tblAddLess.Vyear = tblVouSub.Vyear and tblVouSub.VSubItCode = tblMastItem.ItCode and tblAddLess.AdAcDrCode = tblMastAccount.Accode group by tblMastAccount.AcName,tblMastItem.ITTMP2,tblVouSub.Vno,tblVouSub.vtype,tblVouSub.vyear,tblVouSub.VSubCtrNo,tblVouSub.VsubIGSTRt"
    gCn.Execute " Insert Into TmpGentbl (VBILLNO,Nar1,Nar2,Nar3,Vno,Vtype,Vyear,VSRNO,Amt1,Amt2,Amt3,Amt4,Amt5,Amt6,UserName)" & X1
    
    
'    '--Cloth Purchase
'    '--Ms
'    X1 = " tblAddLess.VType in ('PY') and tblAddLess.AdGpDt>=cdate('" & Format(w_Fr, "yyyy,mm,dd") & "')  and tblAddLess.AdGpDt<=Cdate('" & Format(w_To, "yyyy,mm,dd") & "')  and tblVouSub.VsubCGSTAmt+tblVouSub.VsubSGSTAmt > 0 and (tblAddLess.AdTmp5 = 'COL' or tblAddLess.AdTmp5 = 'CLO')"
'    X1 = "select 'I',tblMastAccount.AcName,tblMastItem.ITTMP2,'M.S.',tblVouSub.Vno,tblVouSub.vtype,tblVouSub.vyear,tblVouSub.VSubCtrNo,tblVouSub.VsubCGSTRt+tblVouSub.VsubSGSTRt,sum(tblVouSub.VSubTaxableAmt),sum(tblVouSub.VsubIGSTAmt),sum(tblVouSub.VsubCGSTAmt),sum(tblVouSub.VsubSGSTAmt),sum(tblVouSub.VSubWt),'" & gUserName & "' from tblAddLess,tblVouSub,tblMastAccount,tblMastItem where " & X1 & " and " _
'         & " tblAddLess.Vno = tblVouSub.Vno and tblAddLess.Vtype = tblVouSub.Vtype and tblAddLess.Vyear = tblVouSub.Vyear and tblVouSub.VSubItCode = tblMastItem.ItCode and tblAddLess.AdAcDrCode = tblMastAccount.Accode group by tblMastAccount.AcName,tblMastItem.ITTMP2,tblVouSub.Vno,tblVouSub.vtype,tblVouSub.vyear,tblVouSub.VSubCtrNo,tblVouSub.VsubCGSTRt+tblVouSub.VsubSGSTRt"
'    gCn.Execute " Insert Into TmpGentbl (VBILLNO,Nar1,Nar2,Nar3,Vno,Vtype,Vyear,VSRNO,Amt1,Amt2,Amt3,Amt4,Amt5,Amt6,UserName)" & X1
'    '--OMS
'    X1 = " tblAddLess.VType in ('PY') and tblAddLess.AdGpDt>=cdate('" & Format(w_Fr, "yyyy,mm,dd") & "')  and tblAddLess.AdGpDt<=Cdate('" & Format(w_To, "yyyy,mm,dd") & "')  and tblVouSub.VsubIGSTAmt > 0 and (tblAddLess.AdTmp5 = 'COL' or tblAddLess.AdTmp5 = 'CLO')"
'    X1 = "select 'I',tblMastAccount.AcName,tblMastItem.ITTMP2,'O.M.S.',tblVouSub.Vno,tblVouSub.vtype,tblVouSub.vyear,tblVouSub.VSubCtrNo,tblVouSub.VsubIGSTRt,sum(tblVouSub.VSubTaxableAmt),sum(tblVouSub.VsubIGSTAmt),sum(tblVouSub.VsubCGSTAmt),sum(tblVouSub.VsubSGSTAmt),sum(tblVouSub.VSubWt),'" & gUserName & "' from tblAddLess,tblVouSub,tblMastAccount,tblMastItem where " & X1 & " and " _
'         & " tblAddLess.Vno = tblVouSub.Vno and tblAddLess.Vtype = tblVouSub.Vtype and tblAddLess.Vyear = tblVouSub.Vyear and tblVouSub.VSubItCode = tblMastItem.ItCode and tblAddLess.AdAcDrCode = tblMastAccount.Accode group by tblMastAccount.AcName,tblMastItem.ITTMP2,tblVouSub.Vno,tblVouSub.vtype,tblVouSub.vyear,tblVouSub.VSubCtrNo,tblVouSub.VsubIGSTRt"
'    gCn.Execute " Insert Into TmpGentbl (VBILLNO,Nar1,Nar2,Nar3,Vno,Vtype,Vyear,VSRNO,Amt1,Amt2,Amt3,Amt4,Amt5,Amt6,UserName)" & X1

Else
    
    '--Ms
    X1 = " tblAddLess.VType in ('PY') and tblAddLess.vdt>=cdate('" & Format(w_Fr, "yyyy,mm,dd") & "')  and tblAddLess.vdt<=Cdate('" & Format(w_To, "yyyy,mm,dd") & "')  and tblVouSub.VsubCGSTAmt+tblVouSub.VsubSGSTAmt > 0 and isnull(tblAddLess.AdTmp5) = True "
    X1 = "select 'G',tblMastAccount.AcName,tblMastItem.ITTMP2,'M.S.',tblVouSub.Vno,tblVouSub.vtype,tblVouSub.vyear,tblVouSub.VSubCtrNo,tblVouSub.VsubCGSTRt+tblVouSub.VsubSGSTRt,sum(tblVouSub.VSubTaxableAmt),sum(tblVouSub.VsubIGSTAmt),sum(tblVouSub.VsubCGSTAmt),sum(tblVouSub.VsubSGSTAmt),sum(tblVouSub.VSubBag),'" & gUserName & "' from tblAddLess,tblVouSub,tblMastAccount,tblMastItem where " & X1 & " and " _
         & " tblAddLess.Vno = tblVouSub.Vno and tblAddLess.Vtype = tblVouSub.Vtype and tblAddLess.Vyear = tblVouSub.Vyear and tblVouSub.VSubItCode = tblMastItem.ItCode and tblAddLess.AdAcDrCode = tblMastAccount.Accode group by tblMastAccount.AcName,tblMastItem.ITTMP2,tblVouSub.Vno,tblVouSub.vtype,tblVouSub.vyear,tblVouSub.VSubCtrNo,tblVouSub.VsubCGSTRt+tblVouSub.VsubSGSTRt"
    gCn.Execute " Insert Into TmpGentbl (VBILLNO,Nar1,Nar2,Nar3,Vno,Vtype,Vyear,VSRNO,Amt1,Amt2,Amt3,Amt4,Amt5,Amt6,UserName)" & X1
    '--OMS
    X1 = " tblAddLess.VType in ('PY') and tblAddLess.vdt>=cdate('" & Format(w_Fr, "yyyy,mm,dd") & "')  and tblAddLess.vdt<=Cdate('" & Format(w_To, "yyyy,mm,dd") & "')  and tblVouSub.VsubIGSTAmt > 0 and isnull(tblAddLess.AdTmp5) = True "
    X1 = "select 'G',tblMastAccount.AcName,tblMastItem.ITTMP2,'O.M.S.',tblVouSub.Vno,tblVouSub.vtype,tblVouSub.vyear,tblVouSub.VSubCtrNo,tblVouSub.VsubIGSTRt,sum(tblVouSub.VSubTaxableAmt),sum(tblVouSub.VsubIGSTAmt),sum(tblVouSub.VsubCGSTAmt),sum(tblVouSub.VsubSGSTAmt),sum(tblVouSub.VSubBag),'" & gUserName & "' from tblAddLess,tblVouSub,tblMastAccount,tblMastItem where " & X1 & " and " _
         & " tblAddLess.Vno = tblVouSub.Vno and tblAddLess.Vtype = tblVouSub.Vtype and tblAddLess.Vyear = tblVouSub.Vyear and tblVouSub.VSubItCode = tblMastItem.ItCode and tblAddLess.AdAcDrCode = tblMastAccount.Accode group by tblMastAccount.AcName,tblMastItem.ITTMP2,tblVouSub.Vno,tblVouSub.vtype,tblVouSub.vyear,tblVouSub.VSubCtrNo,tblVouSub.VsubIGSTRt"
    gCn.Execute " Insert Into TmpGentbl (VBILLNO,Nar1,Nar2,Nar3,Vno,Vtype,Vyear,VSRNO,Amt1,Amt2,Amt3,Amt4,Amt5,Amt6,UserName)" & X1
    
    '-- Other Purchase
    '--Ms
    X1 = " tblAddLess.VType in ('PO') and tblAddLess.vdt>=cdate('" & Format(w_Fr, "yyyy,mm,dd") & "')  and tblAddLess.vdt<=Cdate('" & Format(w_To, "yyyy,mm,dd") & "')  and tblVouSub.VsubCGSTAmt+tblVouSub.VsubSGSTAmt > 0 and isnull(tblAddLess.AdTmp5) = True "
    X1 = "select 'H',tblMastAccount.AcName,tblMastItem.ITTMP2,'M.S.',tblVouSub.Vno,tblVouSub.vtype,tblVouSub.vyear,tblVouSub.VSubCtrNo,tblVouSub.VsubCGSTRt+tblVouSub.VsubSGSTRt,sum(tblVouSub.VSubTaxableAmt),sum(tblVouSub.VsubIGSTAmt),sum(tblVouSub.VsubCGSTAmt),sum(tblVouSub.VsubSGSTAmt),sum(tblVouSub.VSubBag),'" & gUserName & "' from tblAddLess,tblVouSub,tblMastAccount,tblMastItem where " & X1 & " and " _
         & " tblAddLess.Vno = tblVouSub.Vno and tblAddLess.Vtype = tblVouSub.Vtype and tblAddLess.Vyear = tblVouSub.Vyear and tblVouSub.VSubItCode = tblMastItem.ItCode and tblAddLess.AdAcDrCode = tblMastAccount.Accode group by tblMastAccount.AcName,tblMastItem.ITTMP2,tblVouSub.Vno,tblVouSub.vtype,tblVouSub.vyear,tblVouSub.VSubCtrNo,tblVouSub.VsubCGSTRt+tblVouSub.VsubSGSTRt"
    gCn.Execute " Insert Into TmpGentbl (VBILLNO,Nar1,Nar2,Nar3,Vno,Vtype,Vyear,VSRNO,Amt1,Amt2,Amt3,Amt4,Amt5,Amt6,UserName)" & X1
    '--OMS
    X1 = " tblAddLess.VType in ('PO') and tblAddLess.vdt>=cdate('" & Format(w_Fr, "yyyy,mm,dd") & "')  and tblAddLess.vdt<=Cdate('" & Format(w_To, "yyyy,mm,dd") & "')  and tblVouSub.VsubIGSTAmt > 0 and isnull(tblAddLess.AdTmp5) = True "
    X1 = "select 'H',tblMastAccount.AcName,tblMastItem.ITTMP2,'O.M.S.',tblVouSub.Vno,tblVouSub.vtype,tblVouSub.vyear,tblVouSub.VSubCtrNo,tblVouSub.VsubIGSTRt,sum(tblVouSub.VSubTaxableAmt),sum(tblVouSub.VsubIGSTAmt),sum(tblVouSub.VsubCGSTAmt),sum(tblVouSub.VsubSGSTAmt),sum(tblVouSub.VSubBag),'" & gUserName & "' from tblAddLess,tblVouSub,tblMastAccount,tblMastItem where " & X1 & " and " _
         & " tblAddLess.Vno = tblVouSub.Vno and tblAddLess.Vtype = tblVouSub.Vtype and tblAddLess.Vyear = tblVouSub.Vyear and tblVouSub.VSubItCode = tblMastItem.ItCode and tblAddLess.AdAcDrCode = tblMastAccount.Accode group by tblMastAccount.AcName,tblMastItem.ITTMP2,tblVouSub.Vno,tblVouSub.vtype,tblVouSub.vyear,tblVouSub.VSubCtrNo,tblVouSub.VsubIGSTRt"
    gCn.Execute " Insert Into TmpGentbl (VBILLNO,Nar1,Nar2,Nar3,Vno,Vtype,Vyear,VSRNO,Amt1,Amt2,Amt3,Amt4,Amt5,Amt6,UserName)" & X1
    
    
    '--Cloth Purchase
    '--Ms
    X1 = " tblAddLess.VType in ('PY') and tblAddLess.vdt>=cdate('" & Format(w_Fr, "yyyy,mm,dd") & "')  and tblAddLess.vdt<=Cdate('" & Format(w_To, "yyyy,mm,dd") & "')  and tblVouSub.VsubCGSTAmt+tblVouSub.VsubSGSTAmt > 0 and (tblAddLess.AdTmp5 = 'COL' or tblAddLess.AdTmp5 = 'CLO')"
    X1 = "select 'I',tblMastAccount.AcName,tblMastItem.ITTMP2,'M.S.',tblVouSub.Vno,tblVouSub.vtype,tblVouSub.vyear,tblVouSub.VSubCtrNo,tblVouSub.VsubCGSTRt+tblVouSub.VsubSGSTRt,sum(tblVouSub.VSubTaxableAmt),sum(tblVouSub.VsubIGSTAmt),sum(tblVouSub.VsubCGSTAmt),sum(tblVouSub.VsubSGSTAmt),sum(tblVouSub.VSubWt),'" & gUserName & "' from tblAddLess,tblVouSub,tblMastAccount,tblMastItem where " & X1 & " and " _
         & " tblAddLess.Vno = tblVouSub.Vno and tblAddLess.Vtype = tblVouSub.Vtype and tblAddLess.Vyear = tblVouSub.Vyear and tblVouSub.VSubItCode = tblMastItem.ItCode and tblAddLess.AdAcDrCode = tblMastAccount.Accode group by tblMastAccount.AcName,tblMastItem.ITTMP2,tblVouSub.Vno,tblVouSub.vtype,tblVouSub.vyear,tblVouSub.VSubCtrNo,tblVouSub.VsubCGSTRt+tblVouSub.VsubSGSTRt"
    gCn.Execute " Insert Into TmpGentbl (VBILLNO,Nar1,Nar2,Nar3,Vno,Vtype,Vyear,VSRNO,Amt1,Amt2,Amt3,Amt4,Amt5,Amt6,UserName)" & X1
    '--OMS
    X1 = " tblAddLess.VType in ('PY') and tblAddLess.vdt>=cdate('" & Format(w_Fr, "yyyy,mm,dd") & "')  and tblAddLess.vdt<=Cdate('" & Format(w_To, "yyyy,mm,dd") & "')  and tblVouSub.VsubIGSTAmt > 0 and (tblAddLess.AdTmp5 = 'COL' or tblAddLess.AdTmp5 = 'CLO')"
    X1 = "select 'I',tblMastAccount.AcName,tblMastItem.ITTMP2,'O.M.S.',tblVouSub.Vno,tblVouSub.vtype,tblVouSub.vyear,tblVouSub.VSubCtrNo,tblVouSub.VsubIGSTRt,sum(tblVouSub.VSubTaxableAmt),sum(tblVouSub.VsubIGSTAmt),sum(tblVouSub.VsubCGSTAmt),sum(tblVouSub.VsubSGSTAmt),sum(tblVouSub.VSubWt),'" & gUserName & "' from tblAddLess,tblVouSub,tblMastAccount,tblMastItem where " & X1 & " and " _
         & " tblAddLess.Vno = tblVouSub.Vno and tblAddLess.Vtype = tblVouSub.Vtype and tblAddLess.Vyear = tblVouSub.Vyear and tblVouSub.VSubItCode = tblMastItem.ItCode and tblAddLess.AdAcDrCode = tblMastAccount.Accode group by tblMastAccount.AcName,tblMastItem.ITTMP2,tblVouSub.Vno,tblVouSub.vtype,tblVouSub.vyear,tblVouSub.VSubCtrNo,tblVouSub.VsubIGSTRt"
    gCn.Execute " Insert Into TmpGentbl (VBILLNO,Nar1,Nar2,Nar3,Vno,Vtype,Vyear,VSRNO,Amt1,Amt2,Amt3,Amt4,Amt5,Amt6,UserName)" & X1
End If

'--Debit Note  Purchase
'--Ms

X1 = " qrytblVoucherSr1.VType = 'SN' and qrytblVoucherSr1.VDRCRNTTYPE = 'S' and qrytblVoucherSr1.vdt>=cdate('" & Format(w_Fr, "yyyy,mm,dd") & "')  and qrytblVoucherSr1.vdt<=Cdate('" & Format(w_To, "yyyy,mm,dd") & "') and (qrytblVoucherSr1.VCGSTAmt+qrytblVoucherSr1.VSGSTAmt) > 0 "  '

X1 = "select 'J',tblMastAccount.AcName,'.','M.S.',qrytblVoucherSr1.Vno,qrytblVoucherSr1.vtype,qrytblVoucherSr1.vyear,qrytblVoucherSr1.VCtrNo,qrytblVoucherSr1.VCGSTRt+qrytblVoucherSr1.VSGSTRt,sum(qrytblVoucherSr1.VTaxableAmt),sum(qrytblVoucherSr1.VIGSTAmt),sum(qrytblVoucherSr1.VCGSTAmt),sum(qrytblVoucherSr1.VSGSTAmt),sum(qrytblVoucherSr1.VCtrNo),'" & gUserName & "' from qrytblVoucherSr1,tblMastAccount where " & X1 & " and " _
     & " qrytblVoucherSr1.VAcOCode = tblMastAccount.Accode group by tblMastAccount.AcName,qrytblVoucherSr1.Vno,qrytblVoucherSr1.vtype,qrytblVoucherSr1.vyear,qrytblVoucherSr1.VCtrNo,qrytblVoucherSr1.VCGSTRt+qrytblVoucherSr1.VSGSTRt"
gCn.Execute " Insert Into TmpGentbl (VBILLNO,Nar1,Nar2,Nar3,Vno,Vtype,Vyear,VSRNO,Amt1,Amt2,Amt3,Amt4,Amt5,Amt6,UserName)" & X1

'--OMs
X1 = " qrytblVoucherSr1.VType = 'SN' and qrytblVoucherSr1.VDRCRNTTYPE = 'S' and qrytblVoucherSr1.vdt>=cdate('" & Format(w_Fr, "yyyy,mm,dd") & "')  and qrytblVoucherSr1.vdt<=Cdate('" & Format(w_To, "yyyy,mm,dd") & "') and (qrytblVoucherSr1.VIGSTAmt) > 0 "  '

X1 = "select 'J',tblMastAccount.AcName,'.','O.M.S.',qrytblVoucherSr1.Vno,qrytblVoucherSr1.vtype,qrytblVoucherSr1.vyear,qrytblVoucherSr1.VCtrNo,qrytblVoucherSr1.VIGSTRt,sum(qrytblVoucherSr1.VTaxableAmt),sum(qrytblVoucherSr1.VIGSTAmt),sum(qrytblVoucherSr1.VCGSTAmt),sum(qrytblVoucherSr1.VSGSTAmt),sum(qrytblVoucherSr1.VCtrNo),'" & gUserName & "' from qrytblVoucherSr1,tblMastAccount where " & X1 & " and " _
     & " qrytblVoucherSr1.VAcOCode = tblMastAccount.Accode group by tblMastAccount.AcName,qrytblVoucherSr1.Vno,qrytblVoucherSr1.vtype,qrytblVoucherSr1.vyear,qrytblVoucherSr1.VCtrNo,qrytblVoucherSr1.VIGSTRt"
gCn.Execute " Insert Into TmpGentbl (VBILLNO,Nar1,Nar2,Nar3,Vno,Vtype,Vyear,VSRNO,Amt1,Amt2,Amt3,Amt4,Amt5,Amt6,UserName)" & X1


'--Credit Note  Purchase
'--Ms
X1 = " qrytblVoucherSr1.VType = 'PN' and qrytblVoucherSr1.VDRCRNTTYPE = 'S' and qrytblVoucherSr1.vdt>=cdate('" & Format(w_Fr, "yyyy,mm,dd") & "')  and qrytblVoucherSr1.vdt<=Cdate('" & Format(w_To, "yyyy,mm,dd") & "') and (qrytblVoucherSr1.VCGSTAmt+qrytblVoucherSr1.VSGSTAmt) > 0 "  '
X1 = "select 'K',tblMastAccount.AcName,'.','M.S.',qrytblVoucherSr1.Vno,qrytblVoucherSr1.vtype,qrytblVoucherSr1.vyear,qrytblVoucherSr1.VCtrNo,qrytblVoucherSr1.VCGSTRt+qrytblVoucherSr1.VSGSTRt,sum(qrytblVoucherSr1.VTaxableAmt),sum(qrytblVoucherSr1.VIGSTAmt),sum(qrytblVoucherSr1.VCGSTAmt),sum(qrytblVoucherSr1.VSGSTAmt),sum(qrytblVoucherSr1.VCtrNo),'" & gUserName & "' from qrytblVoucherSr1,tblMastAccount where " & X1 & " and " _
     & " qrytblVoucherSr1.VAcOCode = tblMastAccount.Accode group by tblMastAccount.AcName,qrytblVoucherSr1.Vno,qrytblVoucherSr1.vtype,qrytblVoucherSr1.vyear,qrytblVoucherSr1.VCtrNo,qrytblVoucherSr1.VCGSTRt+qrytblVoucherSr1.VSGSTRt"
gCn.Execute " Insert Into TmpGentbl (VBILLNO,Nar1,Nar2,Nar3,Vno,Vtype,Vyear,VSRNO,Amt1,Amt2,Amt3,Amt4,Amt5,Amt6,UserName)" & X1
'--OMs
X1 = " qrytblVoucherSr1.VType = 'PN' and qrytblVoucherSr1.VDRCRNTTYPE = 'S' and qrytblVoucherSr1.vdt>=cdate('" & Format(w_Fr, "yyyy,mm,dd") & "')  and qrytblVoucherSr1.vdt<=Cdate('" & Format(w_To, "yyyy,mm,dd") & "') and (qrytblVoucherSr1.VIGSTAmt) > 0 "  '
X1 = "select 'K',tblMastAccount.AcName,'.','O.M.S.',qrytblVoucherSr1.Vno,qrytblVoucherSr1.vtype,qrytblVoucherSr1.vyear,qrytblVoucherSr1.VCtrNo,qrytblVoucherSr1.VIGSTRt,sum(qrytblVoucherSr1.VTaxableAmt),sum(qrytblVoucherSr1.VIGSTAmt),sum(qrytblVoucherSr1.VCGSTAmt),sum(qrytblVoucherSr1.VSGSTAmt),sum(qrytblVoucherSr1.VCtrNo),'" & gUserName & "' from qrytblVoucherSr1,tblMastAccount where " & X1 & " and " _
     & " qrytblVoucherSr1.VAcOCode = tblMastAccount.Accode group by tblMastAccount.AcName,qrytblVoucherSr1.Vno,qrytblVoucherSr1.vtype,qrytblVoucherSr1.vyear,qrytblVoucherSr1.VCtrNo,qrytblVoucherSr1.VIGSTRt"
gCn.Execute " Insert Into TmpGentbl (VBILLNO,Nar1,Nar2,Nar3,Vno,Vtype,Vyear,VSRNO,Amt1,Amt2,Amt3,Amt4,Amt5,Amt6,UserName)" & X1

'--Purchase Return
'--Ms
X1 = " tblAddLess.VType in ('VY') and tblAddLess.vdt>=cdate('" & Format(w_Fr, "yyyy,mm,dd") & "')  and tblAddLess.vdt<=Cdate('" & Format(w_To, "yyyy,mm,dd") & "')  and tblVouSub.VsubCGSTAmt+tblVouSub.VsubSGSTAmt > 0 and isnull(tblAddLess.AdTmp5) = True "
X1 = "select 'L',tblMastAccount.AcName,tblMastItem.ITTMP2,'M.S.',tblVouSub.Vno,tblVouSub.vtype,tblVouSub.vyear,tblVouSub.VSubCtrNo,tblVouSub.VsubCGSTRt+tblVouSub.VsubSGSTRt,sum(tblVouSub.VSubTaxableAmt),sum(tblVouSub.VsubIGSTAmt),sum(tblVouSub.VsubCGSTAmt),sum(tblVouSub.VsubSGSTAmt),sum(tblVouSub.VSubBag),'" & gUserName & "' from tblAddLess,tblVouSub,tblMastAccount,tblMastItem where " & X1 & " and " _
     & " tblAddLess.Vno = tblVouSub.Vno and tblAddLess.Vtype = tblVouSub.Vtype and tblAddLess.Vyear = tblVouSub.Vyear and tblVouSub.VSubItCode = tblMastItem.ItCode and tblAddLess.AdAcCrCode = tblMastAccount.Accode group by tblMastAccount.AcName,tblMastItem.ITTMP2,tblVouSub.Vno,tblVouSub.vtype,tblVouSub.vyear,tblVouSub.VSubCtrNo,tblVouSub.VsubCGSTRt+tblVouSub.VsubSGSTRt"
gCn.Execute " Insert Into TmpGentbl (VBILLNO,Nar1,Nar2,Nar3,Vno,Vtype,Vyear,VSRNO,Amt1,Amt2,Amt3,Amt4,Amt5,Amt6,UserName)" & X1
'--OMS
X1 = " tblAddLess.VType in ('VY') and tblAddLess.vdt>=cdate('" & Format(w_Fr, "yyyy,mm,dd") & "')  and tblAddLess.vdt<=Cdate('" & Format(w_To, "yyyy,mm,dd") & "')  and tblVouSub.VsubIGSTAmt > 0 and isnull(tblAddLess.AdTmp5) = True "
X1 = "select 'L',tblMastAccount.AcName,tblMastItem.ITTMP2,'O.M.S.',tblVouSub.Vno,tblVouSub.vtype,tblVouSub.vyear,tblVouSub.VSubCtrNo,tblVouSub.VsubIGSTRt,sum(tblVouSub.VSubTaxableAmt),sum(tblVouSub.VsubIGSTAmt),sum(tblVouSub.VsubCGSTAmt),sum(tblVouSub.VsubSGSTAmt),sum(tblVouSub.VSubBag),'" & gUserName & "' from tblAddLess,tblVouSub,tblMastAccount,tblMastItem where " & X1 & " and " _
     & " tblAddLess.Vno = tblVouSub.Vno and tblAddLess.Vtype = tblVouSub.Vtype and tblAddLess.Vyear = tblVouSub.Vyear and tblVouSub.VSubItCode = tblMastItem.ItCode and tblAddLess.AdAcCrCode = tblMastAccount.Accode group by tblMastAccount.AcName,tblMastItem.ITTMP2,tblVouSub.Vno,tblVouSub.vtype,tblVouSub.vyear,tblVouSub.VSubCtrNo,tblVouSub.VsubIGSTRt"
gCn.Execute " Insert Into TmpGentbl (VBILLNO,Nar1,Nar2,Nar3,Vno,Vtype,Vyear,VSRNO,Amt1,Amt2,Amt3,Amt4,Amt5,Amt6,UserName)" & X1





gCn.CommitTrans



End Function


Public Sub GProcPrepareProfitLossMonthly(FromDt As Date, ToDate As Date, Optional OpStk As Double)
Dim X As String
Dim RsTmp_S As Recordset
Dim RsTmp_T As Recordset
Dim i As Long
Dim RCount_T As Long
Dim RCount_S  As Long
'---- insert To Dt in tmpSelection
gCn.Execute "delete from tmpSelection where UserName='" & gUserName & "'"
If gBackEndDB = gBackEndAccess Then
    gCn.Execute "insert into tmpSelection (dt1,dt2,TmpNo1,UserName) values ('" & FromDt & "','" & ToDate & "'," & OpStk & ",'" & gUserName & "')"
Else
    gCn.Execute "insert into tmpSelection (dt2,UserName) values (to_date('" & ToDate & "','dd/MM/yyyy'),'" & gUserName & "') "
End If

gCn.Execute "delete from tmpGenTbl where UserName='" & gUserName & "'"
'--- Insert Top Debit Side In TmpGenTbl
'gCn.Execute "insert into tmpGenTbl  select Gp,Sr,AgName,AcName,Debit from QryPL2 where UserName='" & gUserName & "' and GP=1 Order by Gp,Sr,AgName,AcName"
X = "select Gp,GrNo,Sr,AcName,Debit,GrDebit,UserName from QryMonPL2_1 where UserName='" & gUserName & "' and ( Debit>0  or GrDebit>0 ) and gp=1 Order by Gp,Sr,AgName,GrNo,AcName"
If GProcRstOpen(RsTmp_S, X, "R") > 0 Then
    With RsTmp_S
        .MoveFirst
        Do While Not .EOF
           gCn.Execute "insert into tmpGenTbl (Amt10,Amt11,vno,VSrNo,Nar1,Amt1,Amt2,username) " _
           & " values (" & .AbsolutePosition & "," & .Fields(0) & "," & .Fields(1) & "," & .Fields(2) & ",'" & .Fields(3) & "'," & .Fields(4) & "," & .Fields(5) & ",'" & .Fields(6) & "')"
            .MoveNext
        Loop
    End With
End If
'--- Insert Top Credit Side In TmpGenTbl
RCount_T = GProcRstOpen(RsTmp_T, "Select * from TmpGenTbl where UserName='" & gUserName & "' order by Amt10 ", "O")
X = "select Gp,GrNo,Sr,AcName,Credit,GrCredit,UserName from QryMonPL2_1 where gp=1 and UserName='" & gUserName & "' and ( Credit>0  or GrCredit>0 ) Order by Gp,Sr,AgName,GrNo,AcName"
RCount_S = GProcRstOpen(RsTmp_S, X, "R")
If RCount_S > 0 Then
    With RsTmp_S
        .MoveFirst
        Do While Not .EOF
           i = i + 1
           If i > RCount_T Then
              gCn.Execute "insert into tmpGenTbl (Amt10,Amt11,vno,VSrNo,Nar2,Amt3,Amt4,username) " _
               & " values (" & i & "," & .Fields(0) & "," & .Fields(1) & "," & .Fields(2) & ",'" & .Fields(3) & "','" & .Fields(4) & "'," & .Fields(5) & ",'" & .Fields(6) & "')"
           Else
                gCn.Execute "update tmpGenTbl set Nar2='" & .Fields("AcName") & "',Amt3=" & .Fields("Credit") & ",Amt4=" & .Fields("GrCredit") & " where " _
                 & " Amt10=" & RsTmp_T.Fields("Amt10") & " and UserName='" & gUserName & "'"
           End If
            If Not RsTmp_T.EOF Then RsTmp_T.MoveNext
            .MoveNext
        Loop
    End With
End If
'--- Insert Bottom Debit Side In TmpGenTbl
RCount_T = GProcRstOpen(RsTmp_T, "Select * from TmpGenTbl where UserName='" & gUserName & "' order by Amt10 ", "O")
X = "select Gp,GrNo,Sr,AcName,Debit,GrDebit,UserName from QryMonPL2_2 where gp=2 and UserName='" & gUserName & "' and ( Debit>0  or GrDebit>0 ) Order by Gp,Sr,AgName,GrNo,AcName"
RCount_S = GProcRstOpen(RsTmp_S, X, "R")
i = RCount_T
If RCount_S > 0 Then
    With RsTmp_S
        .MoveFirst
        Do While Not .EOF
           i = i + 1
           'If i > RCount_T Then
              gCn.Execute "insert into tmpGenTbl (Amt10,Amt11,vno,VSrNo,Nar1,Amt1,Amt2,username) " _
               & " values (" & i & "," & .Fields(0) & "," & .Fields(1) & "," & .Fields(2) & ",'" & .Fields(3) & "','" & .Fields(4) & "'," & .Fields(5) & ",'" & .Fields(6) & "')"
           'Else
           '     gCn.Execute "update tmpGenTbl set Nar1='" & .Fields("AcName") & "',Amt1=" & .Fields("Debit") & ",Amt2=" & .Fields("GrDebit") & " where " _
           '      & " Amt10=" & RsTmp_T.Fields("Amt10") & " and UserName='" & gUserName & "'"
           'End If
           ' If Not RsTmp_T.EOF Then RsTmp_T.MoveNext
            .MoveNext
        Loop
    End With
End If
'--- Insert Bottom Credit Side In TmpGenTbl
RCount_T = GProcRstOpen(RsTmp_T, "Select * from TmpGenTbl where UserName='" & gUserName & "' and Amt11=2 order by Amt10 ", "O")
X = "select Gp,GrNo,Sr,AcName,Credit,GrCredit,UserName from QryMonPL2_2 where gp=2 and UserName='" & gUserName & "' and ( Credit>0  or GrCredit>0 ) Order by Gp,Sr,AgName,GrNo,AcName"
Dim j As Long
RCount_S = GProcRstOpen(RsTmp_S, X, "R")
If RCount_S > 0 Then
    With RsTmp_S
        .MoveFirst
        Do While Not .EOF
           j = j + 1
           If j > RCount_T Then
              i = i + 1
              gCn.Execute "insert into tmpGenTbl (Amt10,Amt11,vno,VSrNo,Nar2,Amt3,Amt4,username) " _
               & " values (" & i & "," & .Fields(0) & "," & .Fields(1) & "," & .Fields(2) & ",'" & .Fields(3) & "','" & .Fields(4) & "'," & .Fields(5) & ",'" & .Fields(6) & "')"
           Else
                gCn.Execute "update tmpGenTbl set Nar2='" & .Fields("AcName") & "',Amt3=" & .Fields("Credit") & ",Amt4=" & .Fields("GrCredit") & " where " _
                 & " Amt10=" & RsTmp_T.Fields("Amt10") & " and UserName='" & gUserName & "'"
           End If
            If Not RsTmp_T.EOF Then RsTmp_T.MoveNext
            .MoveNext
        Loop
    End With
End If

'If FromDt > gCYSDate Then
'    Dim rstStkVal As Recordset
'    X = "select StAmount from tblClosingStock where Stdate =  cdate('" & FromDt - 1 & "') "
'    RCount_S = GProcRstOpen(rstStkVal, X, "R")
'    If RCount_S > 0 Then
'       gCn.Execute "update tmpGenTbl set Amt1=  " & rstStkVal.fields(0) & "  where   (trim(tmpGenTbl.Nar1) = 'Stock In Hand' )  and  UserName='" & gUserName & "' "
'       gCn.Execute "update tmpGenTbl set Amt2 = " & rstStkVal.fields(0) & "  where   (trim(tmpGenTbl.Nar1) = 'Opening Stock')  and  UserName='" & gUserName & "' "
'    End If
'End If

End Sub
Public Sub GProcWhatsAppError()
Dim X As String
Dim RsTmp As Recordset
Dim RetVal
gReportPath = App.Path & "\HIReports_Access\"
Open "" & gReportPath & "rpt1.txt" For Output As #1
X = "Select * from tmpWhatsAppErr where username = '" & gUserName & "'"
             
If GProcRstOpen(RsTmp, X, "R", gCn) > 0 Then
   With RsTmp
       .MoveFirst
       Print #1, ""
       Print #1, ""
       Print #1, "Mobile No 1       Mobile No 2     Party Name                      Error"
       Print #1, "-----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------"
       Print #1, ""
       Do While Not .EOF
          X = Left(.Fields(0) + Space(15), 15) + "  " + Left(IIf(IsNull(.Fields(1)), "", .Fields(1)) + Space(15), 15) + "  " + Left(.Fields(2) + Space(30), 30) + "  " + Left(.Fields(3) + Space(250), 250)
          Print #1, X
          .MoveNext
       Loop
       Print #1, ""
       Print #1, "-----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------"
   End With
   Close #1
   RetVal = Shell(gReportPath + "WORDPAD " & "" & gReportPath & "rpt1.txt", vbMaximizedFocus)
End If
End Sub

Public Sub GProcOutStandingSALatePayIntPurch(FromDt As Date, ToDt As Date, AllFirmFlag As Boolean, InterestP As Double, GraseDays As Long, Party As String, IntLateFrmInvDt As Long, IsIntDbNtRpt As Boolean, Optional IsMillBillRec As Boolean)
Dim CondStr As String
Dim X As String
'-- Date
If gBackEndDB = gBackEndAccess Then
    CondStr = " INVDT>= cdate('" & FromDt & "')  and INVDT<= cdate('" & ToDt & "')"
'    CondStr = " RECVDT>= cdate('" & FromDt & "')  and RECVDT<= cdate('" & ToDt & "')"
Else '-- Oracle
    CondStr = " RECVDT>= to_date('" & FromDt & "','dd/MM/yyyy') and RECVDT<= to_date('" & ToDt & "','dd/MM/yyyy')"
End If
'-- Type
If IsMillBillRec = False Then
    CondStr = CondStr + " and (RECVTYPE='BP' or RECVTYPE='CP' ) "
Else
    CondStr = CondStr + " and RECVTYPE='MR' " '-- Mill Bill Receipt
End If
If Party <> "" Then CondStr = CondStr & " and AcName='" & Party & "'" '--Party
'-- Is Interest DbNt
If IsIntDbNtRpt = True Then
    CondStr = CondStr & " and IntDbNtDate is Null" '-IntDbNtDt is Null
End If
X = " select ADJNO,RecVdt,AcName,InvBillNo,InvDt,BILLAMT,AdjAmt, BILLAMT-AdjAmt,INTRECEIVED,RecVdt- " & IIf(IntLateFrmInvDt = 1, "OutPurBillDt-", "OUTDUEDATE-") & GraseDays & ",'" & gUserName & "'," & InterestP & "," & GraseDays & ", " _
& " OUTDUEDATE,OutPurBillDt,VIrnNo from tblmastaccount,TblOutstanding,TblRecVsSale,qrytblVoucherSr1 where accode=PartyCode and " & CondStr & " And " _
& " TblOutstanding.vno=tblRecVsSale.Invno and TblOutstanding.vtype=tblRecVsSale.Invtp and TblOutstanding.vYear=tblRecVsSale.InvYear and qrytblVoucherSr1.vno=tblRecVsSale.RecVNo and qrytblVoucherSr1.vtype=tblRecVsSale.RecVType and qrytblVoucherSr1.vyear=tblRecVsSale.RecVYear"
'--- insert to setup db Table For All Firm
GProcInsertToSetupDbTbl AllFirmFlag, X, "tmpGenTbl", "Vno,vdt,Nar1,VBillNo,VBillDt,Amt1,amt2,Amt3,Amt4,amt5,username,amt6,amt7,Dt1,Dt2,Nar4", " username='" & gUserName & "'", "Nar2"
End Sub
Public Function GProcPurpareGSTR1_V21(w_Fr As Date, w_To As Date, EJ As String)
Dim RstCompMast As Recordset
Dim mTempCn As Connection, mCurrFirmDbNm As String, X1 As String, ComCd As String, Y As String
Dim RsTmp As Recordset
Dim X As String
Dim flnm As String
Dim BarCnt As Double
Dim RecYN As Long
Dim tmp1 As Recordset
Dim tmp2 As Recordset
Dim excel_app As Excel.Application
Dim workbook As Excel.workbook
Dim sheet_name As String
Dim sheet As Excel.Worksheet
Dim header_range As Excel.Range
Dim a As Long
'Dim X1 As String
Dim Y2 As String
Dim TotCount As Long
Dim aFile1 As String
Dim fm As String
Dim PGstin As String
Dim HeadPrint As Boolean
Dim LoopIn As Boolean

gSetupCn.BeginTrans
gSetupCn.Execute "delete from tmpGenTbl where username='" & gUserName & "'"
gSetupCn.CommitTrans
GProcRstOpen RstCompMast, "Select * from tblmastcompany where CGSTIN = '" & gCGSTIN & "' ", "R", gSetupCn
mCurrFirmDbNm = gDbNm
RstCompMast.MoveFirst
Do While Not RstCompMast.EOF
   gDbNm = RstCompMast!CCode + Right(mCurrFirmDbNm, 4)
   ComCd = RstCompMast!CCode
   Y = gDataPath + "" & gDbNm & "" + ".mdb"
   GProcConnectionOpen mTempCn
   
   '--For update HSN in Crn,Drn
   X = "select qrytblVoucherSr1.vno,qrytblVoucherSr1.vtype,qrytblVoucherSr1.vyear,qrytblVoucherSr1.vamt,qrytblVoucherSr1.vnar4 ,tblOutstanding.OutItemCode,tblMastItem.Itname,tblMastItem.ItTmp2 " _
       & " From [" + Y + "].qrytblVoucherSr1, [" + Y + "].tblOutstanding, [" + Y + "].tblMastItem " _
       & " where qrytblVoucherSr1.vtype in ('SN','PN') and qrytblVoucherSr1.VDRCRNTTYPE IN ('B','S') and " _
       & " qrytblVoucherSr1.VAgtVno = tblOutstanding.Vno and qrytblVoucherSr1.VAgtVtype = tblOutstanding.Vtype and qrytblVoucherSr1.VAgtVyear = tblOutstanding.Vyear and " _
       & " tblOutstanding.OutItemCode = tblMastItem.ItCode and (qrytblVoucherSr1.VHsn = '' or isnull(qrytblVoucherSr1.VHsn) )"
   
   RecYN = GProcRstOpen(tmp1, X, "R", mTempCn)
   If RecYN > 0 Then
      With tmp1
        .MoveFirst
        Do While Not .EOF
           mTempCn.Execute " Update [" + Y + "].tblVoucher set VHsn = '" & tmp1.Fields("ItTmp2") & "' where tblVoucher.vno = " & tmp1.Fields("Vno") & " and tblVoucher.vtype = '" & tmp1.Fields("Vtype") & "' and tblVoucher.vyear = " & tmp1.Fields("Vyear") & " "
           .MoveNext
        Loop
      End With
   End If
   '--B2B Yarn
   X1 = " tblAddLess.VType in ('SY') and tblAddLess.vdt>=cdate('" & Format(w_Fr, "yyyy,mm,dd") & "')  and tblAddLess.vdt<=Cdate('" & Format(w_To, "yyyy,mm,dd") & "') and len(tblmastaccount.AcGSTIN) > 3 and (isnull(tblAddLess.AdTmp5) OR tblAddLess.AdTmp5 = 'OTH')  "  '
   X = "select 'B2B',tblAddLess.vno,tblAddLess.vtype,tblAddLess.vyear,ucase(tblmastaccount.AcGSTIN),left(tblmastaccount.AcGSTIN,2),tblMastAccount.Acname,tblAddLess.AdInvNo,tblAddLess.Vdt,tblAddLess.AdBillAmt,tblMastState.StateName, 'N','Regular','', iif(tblVouSub.VsubIGSTRt>0,tblVouSub.VsubIGSTRt,tblVouSub.VsubCGSTRt+tblVouSub.VsubSGSTRt),sum(tblVouSub.VSubAmt),0,sum(tblVouSub.VsubCGSTAmt),sum(tblVouSub.VsubSGSTAmt),sum(tblVouSub.VsubIGSTAmt),'" & gUserName & "' ,val(mid(tblAddLess.AdInvNo,5,10))" _
               & " From [" + Y + "].tblmastaccount,[" + Y + "].tblAddLess,[" + Y + "].tblVouSub,[" + Y + "].tblMastState where tblAddLess.AdAcDrCode = tblMastAccount.Accode and tblAddLess.Vno = tblVouSub.Vno and tblAddLess.Vtype = tblVouSub.Vtype and tblAddLess.Vyear = tblVouSub.Vyear and left(tblMastAccount.AcGSTIN,2) = tblMastState.StateCd  and " & X1 & "" _
               & " group by tblAddLess.vno,tblAddLess.vtype,tblAddLess.vyear,tblmastaccount.AcGSTIN,tblMastAccount.Acname,AdInvNo,tblAddLess.Vdt,tblAddLess.AdBillAmt,tblMastState.StateName,iif(tblVouSub.VsubIGSTRt>0,tblVouSub.VsubIGSTRt,tblVouSub.VsubCGSTRt+tblVouSub.VsubSGSTRt)"
   
   gSetupCn.BeginTrans
   gSetupCn.Execute " Insert Into TmpGentbl (Nar10,vno,Nar6,vyear,Nar1,vtype,Nar7,VbillNo,Vdt,Amt1,Nar4,Nar5,Nar2,Nar3,Amt2,Amt3,Amt4,Amt7,Amt8,Amt9,UserName,Amt5)" & X
   gSetupCn.CommitTrans
   
'   '--B2B Cloth & Cotton
'   X1 = " tblAddLess.VType in ('SY') and tblAddLess.vdt>=cdate('" & Format(w_Fr, "yyyy,mm,dd") & "')  and tblAddLess.vdt<=Cdate('" & Format(w_To, "yyyy,mm,dd") & "') and len(tblmastaccount.AcGSTIN) > 3 and isnull(tblAddLess.AdTmp5) = false And tblAddLess.AdTmp5 <> 'OTH'"
'   X = "select 'B2B',tblAddLess.vno,tblAddLess.vtype,tblAddLess.vyear,ucase(tblmastaccount.AcGSTIN),left(tblmastaccount.AcGSTIN,2),tblMastAccount.Acname,tblAddLess.AdInvNo,tblAddLess.Vdt,tblAddLess.AdBillAmt,tblMastState.StateName, 'N','Regular','', iif(tblAddLess.IgstRt>0,tblAddLess.IgstRt,tblAddLess.CgstRt+tblAddLess.SgstRt),tblAddLess.AdAssVal,0,tblAddLess.CgstAmt,tblAddLess.SgstAmt,tblAddLess.IgstAmt,'" & gUserName & "',val(mid(tblAddLess.AdInvNo,5,10))" _
'       & " From [" + Y + "].tblmastaccount,[" + Y + "].tblAddLess,[" + Y + "].tblMastState where tblAddLess.AdAcDrCode = tblMastAccount.Accode and left(tblMastAccount.AcGSTIN,2) = tblMastState.StateCd  and " & X1 & ""
'
'   gSetupCn.BeginTrans
'   gSetupCn.Execute " Insert Into TmpGentbl (Nar10,vno,Nar6,vyear,Nar1,vtype,Nar7,VbillNo,Vdt,Amt1,Nar4,Nar5,Nar2,Nar3,Amt2,Amt3,Amt4,Amt7,Amt8,Amt9,UserName,amt5)" & X
'   gSetupCn.CommitTrans
   
   '--B2CL Yarn
   X1 = " tblAddLess.VType in ('SY')  and tblAddLess.vdt>=cdate('" & Format(w_Fr, "yyyy,mm,dd") & "')  and tblAddLess.vdt<=Cdate('" & Format(w_To, "yyyy,mm,dd") & "') and (len(tblmastaccount.AcGSTIN) <= 3 or isnull(tblmastaccount.AcGSTIN)) and isnull(tblAddLess.AdTmp5) and tblAddLess.AdBillAmt >= 100000"  '
   X = "select 'B2CL',tblAddLess.vno,tblAddLess.vtype,tblAddLess.vyear,ucase(tblmastaccount.AcGSTIN),left(tblmastaccount.AcGSTIN,2),tblAddLess.AdInvNo,tblAddLess.Vdt,tblAddLess.AdBillAmt,tblMastState.StateName, 'N','Regular','', iif(tblVouSub.VsubIGSTRt>0,tblVouSub.VsubIGSTRt,tblVouSub.VsubCGSTRt+tblVouSub.VsubSGSTRt),sum(tblVouSub.VSubAmt),0,sum(tblVouSub.VsubCGSTAmt),sum(tblVouSub.VsubSGSTAmt),sum(tblVouSub.VsubIGSTAmt),'" & gUserName & "'" _
       & " From [" + Y + "].tblmastaccount,[" + Y + "].tblAddLess,[" + Y + "].tblVouSub,[" + Y + "].tblMastState where tblAddLess.AdAcDrCode = tblMastAccount.Accode and tblAddLess.Vno = tblVouSub.Vno and tblAddLess.Vtype = tblVouSub.Vtype and tblAddLess.Vyear = tblVouSub.Vyear and left(tblMastAccount.AcGSTIN,2) = tblMastState.StateCd  and " & X1 & "" _
       & " group by tblAddLess.vno,tblAddLess.vtype,tblAddLess.vyear,tblmastaccount.AcGSTIN,AdInvNo,tblAddLess.Vdt,tblAddLess.AdBillAmt,tblMastState.StateName,iif(tblVouSub.VsubIGSTRt>0,tblVouSub.VsubIGSTRt,tblVouSub.VsubCGSTRt+tblVouSub.VsubSGSTRt)"
   
   gSetupCn.BeginTrans
   gSetupCn.Execute " Insert Into TmpGentbl (Nar10,vno,Nar6,vyear,Nar1,vtype,VbillNo,Vdt,Amt1,Nar4,Nar5,Nar2,Nar3,Amt2,Amt3,Amt4,Amt7,Amt8,Amt9,UserName)" & X
   gSetupCn.CommitTrans
  
   '--B2CL Cloth & Cotton
   
''   X1 = " tblAddLess.VType in ('SY')  and tblAddLess.vdt>=cdate('" & Format(w_Fr, "yyyy,mm,dd") & "')  and tblAddLess.vdt<=Cdate('" & Format(w_To, "yyyy,mm,dd") & "') and (len(tblmastaccount.AcGSTIN) <= 3 or isnull(tblmastaccount.AcGSTIN)) and isnull(tblAddLess.AdTmp5) = false and tblAddLess.AdBillAmt >= 100000"  '
''
''   X = "select 'B2CL',tblAddLess.vno,tblAddLess.vtype,tblAddLess.vyear,ucase(tblmastaccount.AcGSTIN),left(tblmastaccount.AcGSTIN,2),tblAddLess.AdInvNo,tblAddLess.Vdt,tblAddLess.AdBillAmt,tblMastState.StateName, 'N','Regular','', iif(tblAddLess.IgstRt>0,tblAddLess.IgstRt,tblAddLess.CgstRt+tblAddLess.SgstRt),tblAddLess.AdAssVal,0,tblAddLess.CgstAmt,tblAddLess.SgstAmt,tblAddLess.IgstAmt,'" & gUserName & "'" _
''       & " From [" + Y + "].tblmastaccount,[" + Y + "].tblAddLess,[" + Y + "].tblMastState where tblAddLess.AdAcDrCode = tblMastAccount.Accode and left(tblMastAccount.AcGSTIN,2) = tblMastState.StateCd  and " & X1 & ""
''
''   gSetupCn.BeginTrans
''   gSetupCn.Execute " Insert Into TmpGentbl (Nar10,vno,Nar6,vyear,Nar1,vtype,VbillNo,Vdt,Amt1,Nar4,Nar5,Nar2,Nar3,Amt2,Amt3,Amt4,Amt7,Amt8,Amt9,UserName)" & X
''   gSetupCn.CommitTrans
             
   '--B2CS Yarn
   X1 = " tblAddLess.VType in ('SY')  and tblAddLess.vdt>=cdate('" & Format(w_Fr, "yyyy,mm,dd") & "')  and tblAddLess.vdt<=Cdate('" & Format(w_To, "yyyy,mm,dd") & "') and (len(tblmastaccount.AcGSTIN) <= 3 or isnull(tblmastaccount.AcGSTIN)) and isnull(tblAddLess.AdTmp5) and tblAddLess.AdBillAmt < 100000"  '
             
   X = "select 'B2CS',ucase(tblmastaccount.AcGSTIN),left(tblmastaccount.AcGSTIN,2),tblMastState.StateName,iif(tblVouSub.VsubIGSTRt>0,tblVouSub.VsubIGSTRt,tblVouSub.VsubCGSTRt+tblVouSub.VsubSGSTRt),sum(tblVouSub.VSubAmt),0,'" & gUserName & "'" _
       & " From [" + Y + "].tblmastaccount,[" + Y + "].tblAddLess,[" + Y + "].tblVouSub,[" + Y + "].tblMastState where tblAddLess.AdAcDrCode = tblMastAccount.Accode and tblAddLess.Vno = tblVouSub.Vno and tblAddLess.Vtype = tblVouSub.Vtype and tblAddLess.Vyear = tblVouSub.Vyear and left(tblMastAccount.AcGSTIN,2) = tblMastState.StateCd  and " & X1 & "" _
       & " group by ucase(tblmastaccount.AcGSTIN),left(tblmastaccount.AcGSTIN,2),tblMastState.StateName,iif(tblVouSub.VsubIGSTRt>0,tblVouSub.VsubIGSTRt,tblVouSub.VsubCGSTRt+tblVouSub.VsubSGSTRt)"
             
   gSetupCn.BeginTrans
   gSetupCn.Execute " Insert Into TmpGentbl (Nar10,Nar4,Nar5,Nar2,Amt2,Amt3,Amt4,UserName)" & X
   gSetupCn.CommitTrans
   
   '--B2CS for Cloth & Cotton
             
''   X1 = " tblAddLess.VType in ('SY')  and tblAddLess.vdt>=cdate('" & Format(w_Fr, "yyyy,mm,dd") & "')  and tblAddLess.vdt<=Cdate('" & Format(w_To, "yyyy,mm,dd") & "') and (len(tblmastaccount.AcGSTIN) <= 3 or isnull(tblmastaccount.AcGSTIN)) and isnull(tblAddLess.AdTmp5)= false and tblAddLess.AdBillAmt < 100000"  '
''
''   X = "select 'B2CS',ucase(tblmastaccount.AcGSTIN),left(tblmastaccount.AcGSTIN,2),tblMastState.StateName,iif(tblAddLess.IgstRt>0,tblAddLess.IgstRt,tblAddLess.CgstRt+tblAddLess.SgstRt),tblAddLess.AdAssVal,0,'" & gUserName & "'" _
''       & " From [" + Y + "].tblmastaccount,[" + Y + "].tblAddLess,[" + Y + "].tblMastState where tblAddLess.AdAcDrCode = tblMastAccount.Accode  and left(tblMastAccount.AcGSTIN,2) = tblMastState.StateCd  and " & X1 & ""
''
''   gSetupCn.BeginTrans
''   gSetupCn.Execute " Insert Into TmpGentbl (Nar10,Nar4,Nar5,Nar2,Amt2,Amt3,Amt4,UserName)" & X
''   gSetupCn.CommitTrans
   
   '--CDNR
   
   X1 = " qrytblVoucherSr1.VType in ('SN','PN') and qrytblVoucherSr1.VDrCrNtType = 'B' and qrytblVoucherSr1.vdt>=cdate('" & Format(w_Fr, "yyyy,mm,dd") & "')  and qrytblVoucherSr1.vdt<=Cdate('" & Format(w_To, "yyyy,mm,dd") & "') and len(tblmastaccount.AcGSTIN) > 3  "  '
   

   
   X = "Select 'CDNR',tblmastaccount.AcGSTIN,left(tblmastaccount.AcGSTIN,2),tblMastAccount.Acname,qrytblVoucherSr1.Vno,qrytblVoucherSr1.VDt,iif((isnull(qrytblVoucherSr1.VDrCrNtNo) or qrytblVoucherSr1.VDrCrNtNo = ''),vno,qrytblVoucherSr1.VDrCrNtNo),qrytblVoucherSr1.Vdt,iif(qrytblVoucherSr1.vtype = 'SN','D','C'),'07-Others',tblMastState.StateName,abs(qrytblVoucherSr1.VAmt),qrytblVoucherSr1.VCGSTRt+qrytblVoucherSr1.VSGSTRt+qrytblVoucherSr1.VIGSTRt,qrytblVoucherSr1.VTaxableAmt,0,qrytblVoucherSr1.VCGSTAmt,qrytblVoucherSr1.VSGSTAmt,qrytblVoucherSr1.VIGSTAmt, " _
       & " iif(qrytblVoucherSr1.VBillDt < cdate('2017/07/01'),'Y','N'),'" & gUserName & "' from [" + Y + "].qrytblVoucherSr1,[" + Y + "].tblMastAccount,[" + Y + "].tblMastState where qrytblVoucherSr1.VacCode = tblMastAccount.AcCode and left(tblMastAccount.AcGSTIN,2) = tblMastState.StateCd and qrytblVoucherSr1.VCGSTAmt+qrytblVoucherSr1.VSGSTAmt+qrytblVoucherSr1.VIGSTAmt > 0  and " & X1 & ""
             
   gSetupCn.BeginTrans
   gSetupCn.Execute " Insert Into TmpGentbl (Nar10,Nar1,Vtype,Nar7,VBILLNO,VBILLDT,Nar6,Vdt,Nar2,Nar3,Nar5,Amt1,Amt2,Amt3,Amt4,Amt7,Amt8,Amt9,Nar4,UserName)" & X
   gSetupCn.CommitTrans
   
   '--CDNR - Sales Rtn
             
   X1 = " tblAddLess.VType in ('RY')  and tblAddLess.vdt>=cdate('" & Format(w_Fr, "yyyy,mm,dd") & "')  and tblAddLess.vdt<=Cdate('" & Format(w_To, "yyyy,mm,dd") & "') and len(tblmastaccount.AcGSTIN) > 3 "  '

   X = "select 'CDNR',ucase(tblmastaccount.AcGSTIN),tblMastAccount.Acname,left(tblmastaccount.AcGSTIN,2),tblAddLess.vno,tblAddLess.Vdt,tblAddLess.AdInvNo,tblAddLess.AdGpDt,'C','01-Sales Return', tblMastState.StateName,tblAddLess.AdBillAmt,iif(tblVouSub.VsubIGSTRt>0,tblVouSub.VsubIGSTRt,tblVouSub.VsubCGSTRt+tblVouSub.VsubSGSTRt),sum(tblVouSub.VSubAmt),0,sum(tblVouSub.VsubCGSTAmt),sum(tblVouSub.VsubSGSTAmt),sum(tblVouSub.VsubIGSTAmt),iif(tblAddLess.AdGpDt < cdate('2017/07/01'),'Y','N'),'" & gUserName & "'" _
       & " From [" + Y + "].tblmastaccount,[" + Y + "].tblAddLess,[" + Y + "].tblVouSub,[" + Y + "].tblMastState where tblAddLess.AdAcCrCode = tblMastAccount.Accode and tblAddLess.Vno = tblVouSub.Vno and tblAddLess.Vtype = tblVouSub.Vtype and tblAddLess.Vyear = tblVouSub.Vyear and left(tblMastAccount.AcGSTIN,2) = tblMastState.StateCd  and " & X1 & "" _
       & " group by ucase(tblmastaccount.AcGSTIN),tblMastAccount.Acname,left(tblmastaccount.AcGSTIN,2),tblAddLess.vno,tblAddLess.Vdt,tblAddLess.AdInvNo,tblAddLess.AdGpDt,'C','01-Sales Return', tblMastState.StateName,tblAddLess.AdBillAmt,iif(tblVouSub.VsubIGSTRt>0,tblVouSub.VsubIGSTRt,tblVouSub.VsubCGSTRt+tblVouSub.VsubSGSTRt)"
             
   gSetupCn.BeginTrans
   gSetupCn.Execute " Insert Into TmpGentbl (Nar10,Nar1,Nar7,Vtype,Nar6,Vdt,VBILLNO,VBILLDT,Nar2,Nar3,Nar5,Amt1,Amt2,Amt3,Amt4,Amt7,Amt8,Amt9,Nar4,UserName)" & X
   gSetupCn.CommitTrans
   
   '--CDNUR
             
   X1 = " qrytblVoucherSr1.VType in ('SN','PN') and qrytblVoucherSr1.VDrCrNtType = 'B' and qrytblVoucherSr1.vdt>=cdate('" & Format(w_Fr, "yyyy,mm,dd") & "')  and qrytblVoucherSr1.vdt<=Cdate('" & Format(w_To, "yyyy,mm,dd") & "') and (len(tblmastaccount.AcGSTIN) <= 3 or isnull(tblmastaccount.AcGSTIN))"  '
             
   X = "Select 'CDNUR',tblmastaccount.AcGSTIN,left(tblmastaccount.AcGSTIN,2),qrytblVoucherSr1.VBillNo,qrytblVoucherSr1.VBillDt,iif((qrytblVoucherSr1.VDrCrNtNo = '' or isnull(qrytblVoucherSr1.VDrCrNtNo)), cstr(qrytblVoucherSr1.Vno),qrytblVoucherSr1.VDrCrNtNo)   ,qrytblVoucherSr1.Vdt,iif(qrytblVoucherSr1.vtype = 'SN','D','C'),'07-Others',tblMastState.StateName,abs(qrytblVoucherSr1.VAmt),qrytblVoucherSr1.VCGSTRt+qrytblVoucherSr1.VSGSTRt+qrytblVoucherSr1.VIGSTRt,qrytblVoucherSr1.VTaxableAmt,0," _
       & " iif(qrytblVoucherSr1.VBillDt < cdate('2017/07/01'),'Y','N'),'" & gUserName & "' from [" + Y + "].qrytblVoucherSr1,[" + Y + "].tblMastAccount,[" + Y + "].tblMastState where qrytblVoucherSr1.VacCode = tblMastAccount.AcCode and left(tblMastAccount.AcGSTIN,2) = tblMastState.StateCd and qrytblVoucherSr1.VCGSTAmt+qrytblVoucherSr1.VSGSTAmt+qrytblVoucherSr1.VIGSTAmt > 0  and " & X1 & ""
             
   gSetupCn.BeginTrans
   gSetupCn.Execute " Insert Into TmpGentbl (Nar10,Nar1,Vtype,VBILLNO,VBILLDT,Nar6,Vdt,Nar2,Nar3,Nar5,Amt1,Amt2,Amt3,Amt4,Nar4,UserName)" & X
   gSetupCn.CommitTrans
   
   '--HSN Yarn (b2b)
             
   X1 = " tblAddLess.VType in ('SY')  and tblAddLess.vdt>=cdate('" & Format(w_Fr, "yyyy,mm,dd") & "')  and tblAddLess.vdt<=Cdate('" & Format(w_To, "yyyy,mm,dd") & "') and isnull(tblAddLess.AdTmp5) and len(tblmastaccount.AcGSTIN) > 3 "  '
             
   X = "select 'HSN',tblMastItem.ITTMP2,iif(tblMastItem.ItOtherItem=0,'Yarn',''),tblMastItem.ItUnit,iif(tblVouSub.VsubIGSTRt>0,tblVouSub.VsubIGSTRt,tblVouSub.VsubCGSTRt+tblVouSub.VsubSGSTRt),sum(tblVouSub.VSubBag),sum(tblAddLess.AdBillAmt),sum(tblVouSub.VSubAmt),sum(tblVouSub.VsubIGSTAmt),sum(tblVouSub.VsubCGSTAmt),sum(tblVouSub.VsubSGSTAmt),0,'" & gUserName & "'" _
       & " From [" + Y + "].tblAddLess,[" + Y + "].tblVouSub,[" + Y + "].tblMastItem,[" + Y + "].tblMastAccount where tblVouSub.VSubItCode = tblMastItem.ItCode and tblAddLess.Vno = tblVouSub.Vno and tblAddLess.Vtype = tblVouSub.Vtype and tblAddLess.Vyear = tblVouSub.Vyear " _
       & " and tblAddLess.AdAcDrCode = tblMastAccount.Accode and " & X1 & "" _
       & " group by tblMastItem.ITTMP2,iif(tblMastItem.ItOtherItem=0,'Yarn',''),tblMastItem.ItUnit,iif(tblVouSub.VsubIGSTRt>0,tblVouSub.VsubIGSTRt,tblVouSub.VsubCGSTRt+tblVouSub.VsubSGSTRt)"
             
   gSetupCn.BeginTrans
   gSetupCn.Execute " Insert Into TmpGentbl (Nar10,Nar1,Nar2,Nar3,Amt8,Amt1,Amt2,Amt3,Amt4,Amt5,Amt6,Amt7,UserName)" & X
   gSetupCn.CommitTrans

   '--HSN Other(b2b)

   X1 = " tblAddLess.VType in ('SY')  and tblAddLess.vdt>=cdate('" & Format(w_Fr, "yyyy,mm,dd") & "')  and tblAddLess.vdt<=Cdate('" & Format(w_To, "yyyy,mm,dd") & "') and tblAddLess.AdTmp5 = 'OTH' and len(tblmastaccount.AcGSTIN) > 3 "  '
             
   X = "select 'HSN',tblMastItem.ITTMP2,iif(tblMastItem.ItOtherItem=0,'','Other'),tblMastItem.ItUnit,iif(tblVouSub.VsubIGSTRt>0,tblVouSub.VsubIGSTRt,tblVouSub.VsubCGSTRt+tblVouSub.VsubSGSTRt),sum(tblVouSub.VSubBag),sum(tblAddLess.AdBillAmt),sum(tblVouSub.VSubAmt),sum(tblVouSub.VsubIGSTAmt),sum(tblVouSub.VsubCGSTAmt),sum(tblVouSub.VsubSGSTAmt),0,'" & gUserName & "'" _
       & " From [" + Y + "].tblAddLess,[" + Y + "].tblVouSub,[" + Y + "].tblMastItem,[" + Y + "].tblMastAccount where tblVouSub.VSubItCode = tblMastItem.ItCode and tblAddLess.Vno = tblVouSub.Vno and tblAddLess.Vtype = tblVouSub.Vtype and tblAddLess.Vyear = tblVouSub.Vyear " _
       & " and tblAddLess.AdAcDrCode = tblMastAccount.Accode and " & X1 & "" _
       & " group by tblMastItem.ITTMP2,iif(tblMastItem.ItOtherItem=0,'','Other'),tblMastItem.ItUnit,iif(tblVouSub.VsubIGSTRt>0,tblVouSub.VsubIGSTRt,tblVouSub.VsubCGSTRt+tblVouSub.VsubSGSTRt)"
             
   gSetupCn.BeginTrans
   gSetupCn.Execute " Insert Into TmpGentbl (Nar10,Nar1,Nar2,Nar3,Amt8,Amt1,Amt2,Amt3,Amt4,Amt5,Amt6,Amt7,UserName)" & X
   gSetupCn.CommitTrans


   '--HSN Return Yarn (b2b)
             
   X1 = " tblAddLess.VType in ('RY')  and tblAddLess.vdt>=cdate('" & Format(w_Fr, "yyyy,mm,dd") & "')  and tblAddLess.vdt<=Cdate('" & Format(w_To, "yyyy,mm,dd") & "') and len(tblmastaccount.AcGSTIN) > 3 "  '
             
   X = "select 'HSN',tblMastItem.ITTMP2,'Sales Return','OTH-OTHERS',iif(tblVouSub.VsubIGSTRt>0,tblVouSub.VsubIGSTRt,tblVouSub.VsubCGSTRt+tblVouSub.VsubSGSTRt),sum(tblVouSub.VSubBag)*-1,sum(tblAddLess.AdBillAmt)*-1,sum(tblVouSub.VSubAmt)*-1,sum(tblVouSub.VsubIGSTAmt)*-1,sum(tblVouSub.VsubCGSTAmt)*-1,sum(tblVouSub.VsubSGSTAmt)*-1,0,'" & gUserName & "'" _
       & " From [" + Y + "].tblAddLess,[" + Y + "].tblVouSub,[" + Y + "].tblMastItem,[" + Y + "].tblMastAccount where tblVouSub.VSubItCode = tblMastItem.ItCode and tblAddLess.Vno = tblVouSub.Vno and tblAddLess.Vtype = tblVouSub.Vtype and tblAddLess.Vyear = tblVouSub.Vyear " _
       & " and tblAddLess.AdAcCrCode = tblMastAccount.Accode and " & X1 & "" _
       & " group by tblMastItem.ITTMP2,'Sales Return','OTH-OTHERS',iif(tblVouSub.VsubIGSTRt>0,tblVouSub.VsubIGSTRt,tblVouSub.VsubCGSTRt+tblVouSub.VsubSGSTRt)"
             
   gSetupCn.BeginTrans
   gSetupCn.Execute " Insert Into TmpGentbl (Nar10,Nar1,Nar2,Nar3,Amt8,Amt1,Amt2,Amt3,Amt4,Amt5,Amt6,Amt7,UserName)" & X
   gSetupCn.CommitTrans


   '--HSN Cotton(b2b)
             
''   X1 = " tblAddLess.VType in ('SY')  and tblAddLess.vdt>=cdate('" & Format(w_Fr, "yyyy,mm,dd") & "')  and tblAddLess.vdt<=Cdate('" & Format(w_To, "yyyy,mm,dd") & "') and tblAddLess.AdTmp5 = 'COT' and len(tblmastaccount.AcGSTIN) > 3 "  '
''
''   X = "select 'HSN',tblMastItem.ITTMP2,'Cotton',tblMastItem.ItUnit,iif(tblAddLess.IgstRt>0,tblAddLess.IgstRt,tblAddLess.CgstRt+tblAddLess.SgstRt),sum(tblVouSub.VSubBag),sum(iif(tblVousub.VSubCtrNo = 1,tblAddLess.AdBillAmt,0)),sum(iif(tblVousub.VSubCtrNo = 1,tblAddLess.AdAssVal)),sum(iif(tblVousub.VSubCtrNo = 1,tblAddLess.IgstAmt)),sum(iif(tblVousub.VSubCtrNo = 1,tblAddLess.CgstAmt)),sum(iif(tblVousub.VSubCtrNo = 1,tblAddLess.SgstAmt)),0,'" & gUserName & "'" _
''       & " From [" + Y + "].tblAddLess,[" + Y + "].tblVouSub,[" + Y + "].tblMastItem,[" + Y + "].tblMastAccount where tblVouSub.VSubItCode = tblMastItem.ItCode and tblAddLess.Vno = tblVouSub.Vno and tblAddLess.Vtype = tblVouSub.Vtype and tblAddLess.Vyear = tblVouSub.Vyear " _
''       & " and tblAddLess.AdAcDrCode = tblMastAccount.Accode and " & X1 & "" _
''       & " group by tblMastItem.ITTMP2,'Cotton',tblMastItem.ItUnit,iif(tblAddLess.IgstRt>0,tblAddLess.IgstRt,tblAddLess.CgstRt+tblAddLess.SgstRt)"
''
''   gSetupCn.BeginTrans
''   gSetupCn.Execute " Insert Into TmpGentbl (Nar10,Nar1,Nar2,Nar3,Amt8,Amt1,Amt2,Amt3,Amt4,Amt5,Amt6,Amt7,UserName)" & X
''   gSetupCn.CommitTrans

   '--HSN Cloth(b2b)

''   X1 = " tblAddLess.VType in ('SY')  and tblAddLess.vdt>=cdate('" & Format(w_Fr, "yyyy,mm,dd") & "')  and tblAddLess.vdt<=Cdate('" & Format(w_To, "yyyy,mm,dd") & "') and tblAddLess.AdTmp5 = 'COL' and len(tblmastaccount.AcGSTIN) > 3"  '
''
''   X = "select 'HSN',tblMastItem.ITTMP2,'Cloth',tblMastItem.ItUnit,iif(tblAddLess.IgstRt>0,tblAddLess.IgstRt,tblAddLess.CgstRt+tblAddLess.SgstRt),sum(tblVouSub.VSubBag),sum(iif(tblVousub.VSubCtrNo = 1,tblAddLess.AdBillAmt)),sum(iif(tblVousub.VSubCtrNo = 1,tblAddLess.AdAssVal)),sum(iif(tblVousub.VSubCtrNo = 1,tblAddLess.IgstAmt)),sum(iif(tblVousub.VSubCtrNo = 1,tblAddLess.CgstAmt)),sum(iif(tblVousub.VSubCtrNo = 1,tblAddLess.SgstAmt)),0,'" & gUserName & "'" _
''       & " From [" + Y + "].tblAddLess,[" + Y + "].tblVouSub,[" + Y + "].tblMastItem,[" + Y + "].tblMastAccount where tblVouSub.VSubItCode = tblMastItem.ItCode and tblAddLess.Vno = tblVouSub.Vno and tblAddLess.Vtype = tblVouSub.Vtype and tblAddLess.Vyear = tblVouSub.Vyear " _
''       & " and tblAddLess.AdAcDrCode = tblMastAccount.Accode and " & X1 & "" _
''       & " group by tblMastItem.ITTMP2,'Cloth',tblMastItem.ItUnit,iif(tblAddLess.IgstRt>0,tblAddLess.IgstRt,tblAddLess.CgstRt+tblAddLess.SgstRt)"
''
''   gSetupCn.BeginTrans
''   gSetupCn.Execute " Insert Into TmpGentbl (Nar10,Nar1,Nar2,Nar3,Amt8,Amt1,Amt2,Amt3,Amt4,Amt5,Amt6,Amt7,UserName)" & X
''   gSetupCn.CommitTrans
   
   '---HSN Debit Note (b2b)
   
   X1 = " qrytblVoucherSr1.VType = ('SN') and qrytblVoucherSr1.VDRCRNTTYPE = 'B' and qrytblVoucherSr1.vdt>=cdate('" & Format(w_Fr, "yyyy,mm,dd") & "')  and qrytblVoucherSr1.vdt<=Cdate('" & Format(w_To, "yyyy,mm,dd") & "') and len(tblmastaccount.AcGSTIN) > 3"  '
             
   X = "select 'HSN',qrytblVoucherSr1.VHsn,'Debit Note','OTH-OTHERS',iif(qrytblVoucherSr1.VIGSTRt>0,qrytblVoucherSr1.VIGSTRt,qrytblVoucherSr1.VCGSTRt+qrytblVoucherSr1.VSGSTRt),0,sum(abs(qrytblVoucherSr1.Vamt)),sum(qrytblVoucherSr1.VTaxableAmt),sum(qrytblVoucherSr1.VIGSTAmt),sum(qrytblVoucherSr1.VCGSTAmt),sum(qrytblVoucherSr1.VSGSTAmt),0,'" & gUserName & "'" _
       & " From [" + Y + "].qrytblVoucherSr1,[" + Y + "].tblMastAccount where qrytblVoucherSr1.VAcCode = tblMastAccount.Accode and " & X1 & "" _
       & " group by qrytblVoucherSr1.VHsn,'Debit Note','OTH-OTHERS',iif(qrytblVoucherSr1.VIGSTRt>0,qrytblVoucherSr1.VIGSTRt,qrytblVoucherSr1.VCGSTRt+qrytblVoucherSr1.VSGSTRt)"
   
   gSetupCn.BeginTrans
   gSetupCn.Execute " Insert Into TmpGentbl (Nar10,Nar1,Nar2,Nar3,Amt8,Amt1,Amt2,Amt3,Amt4,Amt5,Amt6,Amt7,UserName)" & X
   gSetupCn.CommitTrans
   
   '---HSN Credit Note (b2b)
   
   X1 = " qrytblVoucherSr1.VType = ('PN') and qrytblVoucherSr1.VDRCRNTTYPE = 'B' and qrytblVoucherSr1.vdt>=cdate('" & Format(w_Fr, "yyyy,mm,dd") & "')  and qrytblVoucherSr1.vdt<=Cdate('" & Format(w_To, "yyyy,mm,dd") & "') and len(tblmastaccount.AcGSTIN) > 3"  '
             
   X = "select 'HSN',qrytblVoucherSr1.VHsn,'Credit Note','OTH-OTHERS',iif(qrytblVoucherSr1.VIGSTRt>0,qrytblVoucherSr1.VIGSTRt,qrytblVoucherSr1.VCGSTRt+qrytblVoucherSr1.VSGSTRt),0,sum(abs(qrytblVoucherSr1.Vamt)*-1),sum(qrytblVoucherSr1.VTaxableAmt*-1),sum(qrytblVoucherSr1.VIGSTAmt*-1),sum(qrytblVoucherSr1.VCGSTAmt*-1),sum(qrytblVoucherSr1.VSGSTAmt*-1),0,'" & gUserName & "'" _
       & " From [" + Y + "].qrytblVoucherSr1,[" + Y + "].tblMastAccount where qrytblVoucherSr1.VAcCode = tblMastAccount.Accode and " & X1 & "" _
       & " group by qrytblVoucherSr1.VHsn,'Debit Note','OTH-OTHERS',iif(qrytblVoucherSr1.VIGSTRt>0,qrytblVoucherSr1.VIGSTRt,qrytblVoucherSr1.VCGSTRt+qrytblVoucherSr1.VSGSTRt)"
   
   gSetupCn.BeginTrans
   gSetupCn.Execute " Insert Into TmpGentbl (Nar10,Nar1,Nar2,Nar3,Amt8,Amt1,Amt2,Amt3,Amt4,Amt5,Amt6,Amt7,UserName)" & X
   gSetupCn.CommitTrans
   
   
   
   
   '--HSN Yarn (b2c)
             
   X1 = " tblAddLess.VType in ('SY')  and tblAddLess.vdt>=cdate('" & Format(w_Fr, "yyyy,mm,dd") & "')  and tblAddLess.vdt<=Cdate('" & Format(w_To, "yyyy,mm,dd") & "') and isnull(tblAddLess.AdTmp5) and len(tblmastaccount.AcGSTIN) <= 3 "  '
             
   X = "select 'HSNB2C',tblMastItem.ITTMP2,iif(tblMastItem.ItOtherItem=0,'Yarn',''),tblMastItem.ItUnit,iif(tblVouSub.VsubIGSTRt>0,tblVouSub.VsubIGSTRt,tblVouSub.VsubCGSTRt+tblVouSub.VsubSGSTRt),sum(tblVouSub.VSubBag),sum(tblAddLess.AdBillAmt),sum(tblVouSub.VSubAmt),sum(tblVouSub.VsubIGSTAmt),sum(tblVouSub.VsubCGSTAmt),sum(tblVouSub.VsubSGSTAmt),0,'" & gUserName & "'" _
       & " From [" + Y + "].tblAddLess,[" + Y + "].tblVouSub,[" + Y + "].tblMastItem,[" + Y + "].tblMastAccount where tblVouSub.VSubItCode = tblMastItem.ItCode and tblAddLess.Vno = tblVouSub.Vno and tblAddLess.Vtype = tblVouSub.Vtype and tblAddLess.Vyear = tblVouSub.Vyear " _
       & " and tblAddLess.AdAcDrCode = tblMastAccount.Accode and " & X1 & "" _
       & " group by tblMastItem.ITTMP2,iif(tblMastItem.ItOtherItem=0,'Yarn',''),tblMastItem.ItUnit,iif(tblVouSub.VsubIGSTRt>0,tblVouSub.VsubIGSTRt,tblVouSub.VsubCGSTRt+tblVouSub.VsubSGSTRt)"
             
   gSetupCn.BeginTrans
   gSetupCn.Execute " Insert Into TmpGentbl (Nar10,Nar1,Nar2,Nar3,Amt8,Amt1,Amt2,Amt3,Amt4,Amt5,Amt6,Amt7,UserName)" & X
   gSetupCn.CommitTrans

   '--HSN Other(b2c)

   X1 = " tblAddLess.VType in ('SY')  and tblAddLess.vdt>=cdate('" & Format(w_Fr, "yyyy,mm,dd") & "')  and tblAddLess.vdt<=Cdate('" & Format(w_To, "yyyy,mm,dd") & "') and tblAddLess.AdTmp5 = 'OTH' and len(tblmastaccount.AcGSTIN) <= 3 "  '
             
   X = "select 'HSNB2C',tblMastItem.ITTMP2,iif(tblMastItem.ItOtherItem=0,'','Other'),tblMastItem.ItUnit,iif(tblVouSub.VsubIGSTRt>0,tblVouSub.VsubIGSTRt,tblVouSub.VsubCGSTRt+tblVouSub.VsubSGSTRt),sum(tblVouSub.VSubBag),sum(tblAddLess.AdBillAmt),sum(tblVouSub.VSubAmt),sum(tblVouSub.VsubIGSTAmt),sum(tblVouSub.VsubCGSTAmt),sum(tblVouSub.VsubSGSTAmt),0,'" & gUserName & "'" _
       & " From [" + Y + "].tblAddLess,[" + Y + "].tblVouSub,[" + Y + "].tblMastItem,[" + Y + "].tblMastAccount where tblVouSub.VSubItCode = tblMastItem.ItCode and tblAddLess.Vno = tblVouSub.Vno and tblAddLess.Vtype = tblVouSub.Vtype and tblAddLess.Vyear = tblVouSub.Vyear " _
       & " and tblAddLess.AdAcDrCode = tblMastAccount.Accode and " & X1 & "" _
       & " group by tblMastItem.ITTMP2,iif(tblMastItem.ItOtherItem=0,'','Other'),tblMastItem.ItUnit,iif(tblVouSub.VsubIGSTRt>0,tblVouSub.VsubIGSTRt,tblVouSub.VsubCGSTRt+tblVouSub.VsubSGSTRt)"
             
   gSetupCn.BeginTrans
   gSetupCn.Execute " Insert Into TmpGentbl (Nar10,Nar1,Nar2,Nar3,Amt8,Amt1,Amt2,Amt3,Amt4,Amt5,Amt6,Amt7,UserName)" & X
   gSetupCn.CommitTrans


   '--HSN Cotton(b2c)
             
'   X1 = " tblAddLess.VType in ('SY')  and tblAddLess.vdt>=cdate('" & Format(w_Fr, "yyyy,mm,dd") & "')  and tblAddLess.vdt<=Cdate('" & Format(w_To, "yyyy,mm,dd") & "') and tblAddLess.AdTmp5 = 'COT' and len(tblmastaccount.AcGSTIN) <= 3 "  '
'
'   X = "select 'HSNB2C',tblMastItem.ITTMP2,'Cotton',tblMastItem.ItUnit,iif(tblAddLess.IgstRt>0,tblAddLess.IgstRt,tblAddLess.CgstRt+tblAddLess.SgstRt),sum(tblVouSub.VSubBag),sum(iif(tblVousub.VSubCtrNo = 1,tblAddLess.AdBillAmt,0)),sum(iif(tblVousub.VSubCtrNo = 1,tblAddLess.AdAssVal)),sum(iif(tblVousub.VSubCtrNo = 1,tblAddLess.IgstAmt)),sum(iif(tblVousub.VSubCtrNo = 1,tblAddLess.CgstAmt)),sum(iif(tblVousub.VSubCtrNo = 1,tblAddLess.SgstAmt)),0,'" & gUserName & "'" _
'       & " From [" + Y + "].tblAddLess,[" + Y + "].tblVouSub,[" + Y + "].tblMastItem,[" + Y + "].tblMastAccount where tblVouSub.VSubItCode = tblMastItem.ItCode and tblAddLess.Vno = tblVouSub.Vno and tblAddLess.Vtype = tblVouSub.Vtype and tblAddLess.Vyear = tblVouSub.Vyear " _
'       & " and tblAddLess.AdAcDrCode = tblMastAccount.Accode and " & X1 & "" _
'       & " group by tblMastItem.ITTMP2,'Cotton',tblMastItem.ItUnit,iif(tblAddLess.IgstRt>0,tblAddLess.IgstRt,tblAddLess.CgstRt+tblAddLess.SgstRt)"
'
'   gSetupCn.BeginTrans
'   gSetupCn.Execute " Insert Into TmpGentbl (Nar10,Nar1,Nar2,Nar3,Amt8,Amt1,Amt2,Amt3,Amt4,Amt5,Amt6,Amt7,UserName)" & X
'   gSetupCn.CommitTrans
'
'   '--HSN Cloth(b2c)
'
'   X1 = " tblAddLess.VType in ('SY')  and tblAddLess.vdt>=cdate('" & Format(w_Fr, "yyyy,mm,dd") & "')  and tblAddLess.vdt<=Cdate('" & Format(w_To, "yyyy,mm,dd") & "') and tblAddLess.AdTmp5 = 'COL' and len(tblmastaccount.AcGSTIN) <= 3"  '
'
'   X = "select 'HSNB2C',tblMastItem.ITTMP2,'Cloth',tblMastItem.ItUnit,iif(tblAddLess.IgstRt>0,tblAddLess.IgstRt,tblAddLess.CgstRt+tblAddLess.SgstRt),sum(tblVouSub.VSubBag),sum(iif(tblVousub.VSubCtrNo = 1,tblAddLess.AdBillAmt)),sum(iif(tblVousub.VSubCtrNo = 1,tblAddLess.AdAssVal)),sum(iif(tblVousub.VSubCtrNo = 1,tblAddLess.IgstAmt)),sum(iif(tblVousub.VSubCtrNo = 1,tblAddLess.CgstAmt)),sum(iif(tblVousub.VSubCtrNo = 1,tblAddLess.SgstAmt)),0,'" & gUserName & "'" _
'       & " From [" + Y + "].tblAddLess,[" + Y + "].tblVouSub,[" + Y + "].tblMastItem,[" + Y + "].tblMastAccount where tblVouSub.VSubItCode = tblMastItem.ItCode and tblAddLess.Vno = tblVouSub.Vno and tblAddLess.Vtype = tblVouSub.Vtype and tblAddLess.Vyear = tblVouSub.Vyear " _
'       & " and tblAddLess.AdAcDrCode = tblMastAccount.Accode and " & X1 & "" _
'       & " group by tblMastItem.ITTMP2,'Cloth',tblMastItem.ItUnit,iif(tblAddLess.IgstRt>0,tblAddLess.IgstRt,tblAddLess.CgstRt+tblAddLess.SgstRt)"
'
'   gSetupCn.BeginTrans
'   gSetupCn.Execute " Insert Into TmpGentbl (Nar10,Nar1,Nar2,Nar3,Amt8,Amt1,Amt2,Amt3,Amt4,Amt5,Amt6,Amt7,UserName)" & X
'   gSetupCn.CommitTrans
   
   '--HSN Debit Note(b2c)

   X1 = " qrytblVoucherSr1.VType = ('SN') and qrytblVoucherSr1.VDRCRNTTYPE = 'B' and qrytblVoucherSr1.vdt>=cdate('" & Format(w_Fr, "yyyy,mm,dd") & "')  and qrytblVoucherSr1.vdt<=Cdate('" & Format(w_To, "yyyy,mm,dd") & "') and len(tblmastaccount.AcGSTIN) <= 3"  '
             
   X = "select 'HSNB2C',qrytblVoucherSr1.VHsn,'Debit Note','OTH-OTHERS',iif(qrytblVoucherSr1.VIGSTRt>0,qrytblVoucherSr1.VIGSTRt,qrytblVoucherSr1.VCGSTRt+qrytblVoucherSr1.VSGSTRt),0,sum(abs(qrytblVoucherSr1.Vamt)),sum(qrytblVoucherSr1.VTaxableAmt),sum(qrytblVoucherSr1.VIGSTAmt),sum(qrytblVoucherSr1.VCGSTAmt),sum(qrytblVoucherSr1.VSGSTAmt),0,'" & gUserName & "'" _
       & " From [" + Y + "].qrytblVoucherSr1,[" + Y + "].tblMastAccount where qrytblVoucherSr1.VAcCode = tblMastAccount.Accode and " & X1 & "" _
       & " group by qrytblVoucherSr1.VHsn,'Debit Note','OTH-OTHERS',iif(qrytblVoucherSr1.VIGSTRt>0,qrytblVoucherSr1.VIGSTRt,qrytblVoucherSr1.VCGSTRt+qrytblVoucherSr1.VSGSTRt)"
                         
   gSetupCn.BeginTrans
   gSetupCn.Execute " Insert Into TmpGentbl (Nar10,Nar1,Nar2,Nar3,Amt8,Amt1,Amt2,Amt3,Amt4,Amt5,Amt6,Amt7,UserName)" & X
   gSetupCn.CommitTrans
   
   '---HSN Credit Note (b2c)
   
   X1 = " qrytblVoucherSr1.VType = ('PN') and qrytblVoucherSr1.VDRCRNTTYPE = 'B' and qrytblVoucherSr1.vdt>=cdate('" & Format(w_Fr, "yyyy,mm,dd") & "')  and qrytblVoucherSr1.vdt<=Cdate('" & Format(w_To, "yyyy,mm,dd") & "') and len(tblmastaccount.AcGSTIN) <= 3"  '
             
   X = "select 'HSNB2C',qrytblVoucherSr1.VHsn,'Credit Note','OTH-OTHERS',iif(qrytblVoucherSr1.VIGSTRt>0,qrytblVoucherSr1.VIGSTRt,qrytblVoucherSr1.VCGSTRt+qrytblVoucherSr1.VSGSTRt),0,sum(abs(qrytblVoucherSr1.Vamt)*-1),sum(qrytblVoucherSr1.VTaxableAmt*-1),sum(qrytblVoucherSr1.VIGSTAmt*-1),sum(qrytblVoucherSr1.VCGSTAmt*-1),sum(qrytblVoucherSr1.VSGSTAmt*-1),0,'" & gUserName & "'" _
       & " From [" + Y + "].qrytblVoucherSr1,[" + Y + "].tblMastAccount where qrytblVoucherSr1.VAcCode = tblMastAccount.Accode and " & X1 & "" _
       & " group by qrytblVoucherSr1.VHsn,'Debit Note','OTH-OTHERS',iif(qrytblVoucherSr1.VIGSTRt>0,qrytblVoucherSr1.VIGSTRt,qrytblVoucherSr1.VCGSTRt+qrytblVoucherSr1.VSGSTRt)"
   
   gSetupCn.BeginTrans
   gSetupCn.Execute " Insert Into TmpGentbl (Nar10,Nar1,Nar2,Nar3,Amt8,Amt1,Amt2,Amt3,Amt4,Amt5,Amt6,Amt7,UserName)" & X
   gSetupCn.CommitTrans
   
   
   
   
   
   '--DOCS
   '--DOCS Yarn Sales
   X1 = " tblAddLess.VType in ('SY') and tblAddLess.vdt>=cdate('" & Format(w_Fr, "yyyy,mm,dd") & "')  and tblAddLess.vdt<=Cdate('" & Format(w_To, "yyyy,mm,dd") & "') and isnull(tblAddLess.AdTmp5) "  '
   
   Y2 = "select * from [" + Y + "].tblAddLess,[" + Y + "].tblMastAccount where tblAddLess.AdAcDrCode = tblMastAccount.Accode and " & X1 & " "
                       
   TotCount = GProcRstOpen(tmp2, Y2, "R", mTempCn)
   
   X = "Select 'DOCS-B2B-Yarn','Invoices for outward supply',left(adinvno,4),min(val(mid(adinvno,5,15))),max(val(mid(adinvno,5,15))),max(val(mid(adinvno,5,15)))-min(val(mid(adinvno,5,15)))+1," & TotCount & " ,'" & gUserName & "' from [" + Y + "].tblAddLess,[" + Y + "].tblMastAccount where tblAddLess.AdAcDrCode = tblMastAccount.Accode and " & X1 & " group by left(adinvno,4)"
   
   gSetupCn.BeginTrans
   gSetupCn.Execute " Insert Into TmpGentbl (Nar10,Nar1,Nar2,Amt1,Amt2,Amt3,Amt4,UserName)" & X
   gSetupCn.CommitTrans
   
   '--DOCS Other Sales
   
   X1 = " tblAddLess.VType in ('SY') and tblAddLess.vdt>=cdate('" & Format(w_Fr, "yyyy,mm,dd") & "')  and tblAddLess.vdt<=Cdate('" & Format(w_To, "yyyy,mm,dd") & "') and tblAddLess.AdTmp5 = 'OTH' "  '
   
   Y2 = "select * from [" + Y + "].tblAddLess,[" + Y + "].tblMastAccount where tblAddLess.AdAcDrCode = tblMastAccount.Accode and " & X1 & " "
                       
   TotCount = GProcRstOpen(tmp2, Y2, "R", mTempCn)
   
   X = "Select 'DOCS-B2B-Other','Invoices for outward supply',left(adinvno,4),min(val(mid(adinvno,5,15))),max(val(mid(adinvno,5,15))),max(val(mid(adinvno,5,15)))-min(val(mid(adinvno,5,15)))+1," & TotCount & " ,'" & gUserName & "' from [" + Y + "].tblAddLess,[" + Y + "].tblMastAccount where tblAddLess.AdAcDrCode = tblMastAccount.Accode and " & X1 & " group by left(adinvno,4)"
   
   gSetupCn.BeginTrans
   gSetupCn.Execute " Insert Into TmpGentbl (Nar10,Nar1,Nar2,Amt1,Amt2,Amt3,Amt4,UserName)" & X
   gSetupCn.CommitTrans
   
   '--DOCS Cotton Sales
   
   X1 = " tblAddLess.VType in ('SY') and tblAddLess.vdt>=cdate('" & Format(w_Fr, "yyyy,mm,dd") & "')  and tblAddLess.vdt<=Cdate('" & Format(w_To, "yyyy,mm,dd") & "')  and tblAddLess.AdTmp5 = 'COT' "  '
   
   Y2 = "select * from [" + Y + "].tblAddLess,[" + Y + "].tblMastAccount where tblAddLess.AdAcDrCode = tblMastAccount.Accode and " & X1 & " "
                       
   TotCount = GProcRstOpen(tmp2, Y2, "R", mTempCn)
   
   X = "Select 'DOCS-B2B-Cot','Invoices for outward supply',left(adinvno,4),min(val(mid(adinvno,5,15))),max(val(mid(adinvno,5,15))),max(val(mid(adinvno,5,15)))-min(val(mid(adinvno,5,15)))+1," & TotCount & " ,'" & gUserName & "' from [" + Y + "].tblAddLess,[" + Y + "].tblMastAccount where tblAddLess.AdAcDrCode = tblMastAccount.Accode and " & X1 & " group by left(adinvno,4)"
   
   gSetupCn.BeginTrans
   gSetupCn.Execute " Insert Into TmpGentbl (Nar10,Nar1,Nar2,Amt1,Amt2,Amt3,Amt4,UserName)" & X
   gSetupCn.CommitTrans
   
   '--DOCS Cloth Sales
   
   X1 = " tblAddLess.VType in ('SY') and tblAddLess.vdt>=cdate('" & Format(w_Fr, "yyyy,mm,dd") & "')  and tblAddLess.vdt<=Cdate('" & Format(w_To, "yyyy,mm,dd") & "') and tblAddLess.AdTmp5 = 'CLO' "  '
   
   Y2 = "select * from [" + Y + "].tblAddLess,[" + Y + "].tblMastAccount where tblAddLess.AdAcDrCode = tblMastAccount.Accode and " & X1 & " "
                       
   TotCount = GProcRstOpen(tmp2, Y2, "R", mTempCn)
   
   X = "Select 'DOCS-B2B-Clo','Invoices for outward supply',left(adinvno,4),min(val(mid(adinvno,5,15))),max(val(mid(adinvno,5,15))),max(val(mid(adinvno,5,15)))-min(val(mid(adinvno,5,15)))+1," & TotCount & " ,'" & gUserName & "' from [" + Y + "].tblAddLess,[" + Y + "].tblMastAccount where tblAddLess.AdAcDrCode = tblMastAccount.Accode and " & X1 & " group by left(adinvno,4)"
   
   gSetupCn.BeginTrans
   gSetupCn.Execute " Insert Into TmpGentbl (Nar10,Nar1,Nar2,Amt1,Amt2,Amt3,Amt4,UserName)" & X
   gSetupCn.CommitTrans
   
   '--DOCS Debit Note
   
   X1 = " qrytblVoucherSr1.VType = 'SN' and qrytblVoucherSr1.VDrCrNtType = 'B' and qrytblVoucherSr1.vdt>=cdate('" & Format(w_Fr, "yyyy,mm,dd") & "')  and qrytblVoucherSr1.vdt<=Cdate('" & Format(w_To, "yyyy,mm,dd") & "')  "  '
   Y2 = "select * from [" + Y + "].qrytblVoucherSr1 where " & X1 & " "
                       
   TotCount = GProcRstOpen(tmp2, Y2, "R", mTempCn)
   X = "Select 'DOCS-B2B-Drn','Debit Note',left(VDRCRNTNO,3),min(val(mid(VDRCRNTNO,4,15))),max(val(mid(VDRCRNTNO,4,15))),max(val(mid(VDRCRNTNO,4,15)))-min(val(mid(VDRCRNTNO,4,15)))+1," & TotCount & " ,'" & gUserName & "' from [" + Y + "].qrytblVoucherSr1 where  " & X1 & " group by left(VDRCRNTNO,3)"
   
   gSetupCn.BeginTrans
   gSetupCn.Execute " Insert Into TmpGentbl (Nar10,Nar1,Nar2,Amt1,Amt2,Amt3,Amt4,UserName)" & X
   gSetupCn.CommitTrans
   '--DOCS Credit Note
   X1 = " qrytblVoucherSr1.VType = 'PN' and qrytblVoucherSr1.VDrCrNtType = 'B' and qrytblVoucherSr1.vdt>=cdate('" & Format(w_Fr, "yyyy,mm,dd") & "')  and qrytblVoucherSr1.vdt<=Cdate('" & Format(w_To, "yyyy,mm,dd") & "')  "  '
   Y2 = "select * from [" + Y + "].qrytblVoucherSr1 where " & X1 & " "
                       
   TotCount = GProcRstOpen(tmp2, Y2, "R", mTempCn)
   X = "Select 'DOCS-B2B-Crn','Credit Note',left(VDRCRNTNO,3),min(val(mid(VDRCRNTNO,4,15))),max(val(mid(VDRCRNTNO,4,15))),max(val(mid(VDRCRNTNO,4,15)))-min(val(mid(VDRCRNTNO,4,15)))+1," & TotCount & " ,'" & gUserName & "' from [" + Y + "].qrytblVoucherSr1 where  " & X1 & " group by left(VDRCRNTNO,3)"
   
   gSetupCn.BeginTrans
   gSetupCn.Execute " Insert Into TmpGentbl (Nar10,Nar1,Nar2,Amt1,Amt2,Amt3,Amt4,UserName)" & X
   gSetupCn.CommitTrans
   
   '--DOCS Sales Return
   X1 = " tblAddLess.VType = 'RY' and tblAddLess.vdt>=cdate('" & Format(w_Fr, "yyyy,mm,dd") & "')  and tblAddLess.vdt<=Cdate('" & Format(w_To, "yyyy,mm,dd") & "') "  '
   
   Y2 = "select * from [" + Y + "].tblAddLess where  " & X1 & " "
                       
   TotCount = GProcRstOpen(tmp2, Y2, "R", mTempCn)
   
   X = "Select 'DOCS-B2B-SlRtn','Credit Note (Sales Rtn.)',iif(isnull(min(Vno)),0,min(Vno)),iif(isnull(max(Vno)),0,max(Vno)),iif(isnull(max(Vno)),0,max(Vno))-iif(isnull(min(Vno)),0,min(Vno)+1)," & TotCount & " ,'" & gUserName & "' from [" + Y + "].tblAddLess,[" + Y + "].tblMastAccount where tblAddLess.AdAcDrCode = tblMastAccount.Accode and " & X1 & ""
   
   gSetupCn.BeginTrans
   gSetupCn.Execute " Insert Into TmpGentbl (Nar10,Nar1,Amt1,Amt2,Amt3,Amt4,UserName)" & X
   gSetupCn.CommitTrans
   
   
   
   
   mTempCn.Close
   RstCompMast.MoveNext
Loop
If RstCompMast.State = adStateOpen Then RstCompMast.Close
gDbNm = mCurrFirmDbNm
GProcConnectionOpen gCn
If EJ = "E" Then
    flnm = "D:\GSTR1" + gCCode + "-" + Format(w_Fr, "mmm") + ".xlsx"
    FileCopy gReportPath & "\GSTR1_V21_Blank.xlsx", flnm
    Set excel_app = New Excel.Application
    excel_app.Visible = False
    Set workbook = excel_app.Workbooks.Open(filename:=flnm)
Else
    aFile1 = "D:\GSTR1" + gCCode + "-" + Format(w_Fr, "mmm") + ".json"
    If Len(Dir$(aFile1)) > 0 Then
          Kill aFile1
    End If
    flnm = "D:\GSTR1-" + gCCode + "-" + Format(w_Fr, "mmm") + ".xlsx"
    fm = Right("00" + CStr(Month(w_Fr)), 2) + CStr(Year(w_Fr))
    
    Open "D:\GSTR1" + gCCode + "-" + Format(w_Fr, "mmm") + ".json" For Output As #1

    Print #1, "{" + Chr(34) + "gstin" + Chr(34) + ":" + Chr(34) + "" & gCGSTIN & "" + Chr(34) + Chr(44)
    Print #1, Chr(34) + "fp" + Chr(34) + ":" + Chr(34) + fm + Chr(34) + Chr(44)
    Print #1, Chr(34) + "version" + Chr(34) + ":" + Chr(34) + "GST3.2" + Chr(34) + Chr(44)
    Print #1, Chr(34) + "hash" + Chr(34) + ":" + Chr(34) + "hash" + Chr(34) + Chr(44)

End If
'--B2B
RecYN = GProcRstOpen(tmp1, "select * from TmpGentbl where TmpGentbl.UserName ='" & gUserName & "' and Nar10 = 'B2B'   order by Nar1,vdt,amt5 ", "R", gSetupCn)
If RecYN > 0 Then
   If EJ = "E" Then
      Set sheet = workbook.Sheets.Item("b2b,sez,de")
      If (sheet Is Nothing) Then
          MsgBox "Error b2b,sez,de Sheet"
      End If
      a = 5
   End If
   With tmp1
        .MoveFirst
        If EJ = "J" Then
           Print #1, Chr(34) + "b2b" + Chr(34) + ":" + "["
        End If
Loop1:
        If EJ = "J" Then
           a = 1
        End If
        PGstin = !Nar1
        Do While Not .EOF
           If !Nar1 <> PGstin Then
              If EJ = "J" And a > 1 Then
                 Print #1, Chr(34) + "csamt" + Chr(34) + ":" + "0" + "}" + "}" + "]" + "}" + "]" + "}" + Chr(44)
              End If
              GoTo Loop1
           Else
             If a > 1 And EJ = "J" Then
                 Print #1, Chr(34) + "csamt" + Chr(34) + ":" + "0" + "}" + "}" + "]" + "}" + Chr(44)
             End If
           End If
           If EJ = "E" Then
              sheet.Cells(a, 1) = !Nar1
              sheet.Cells(a, 2) = !Nar7
              sheet.Cells(a, 3) = !VbillNo
              DoEvents
              sheet.Cells(a, 4) = Format(!Vdt, "dd-MMM-YYYY")
              sheet.Cells(a, 5) = !Amt1
              sheet.Cells(a, 6) = !Nar4
              sheet.Cells(a, 7) = "N"
              sheet.Cells(a, 9) = "Regular B2B"
              sheet.Cells(a, 11) = !AMT2
              sheet.Cells(a, 12) = !AMT3
              sheet.Cells(a, 13) = 0
              a = a + 1
           Else
             If a = 1 Then
                Print #1, "{" + Chr(34) + "ctin" + Chr(34) + ":" + Chr(34) + "" & !Nar1 & "" + Chr(34) + Chr(44)
                Print #1, Chr(34) + "inv" + Chr(34) + ":" + "["
             End If
             a = a + 1
             Print #1, "{" + Chr(34) + "inum" + Chr(34) + ":" + Chr(34) + "" & !VbillNo & "" + Chr(34) + Chr(44)
             Print #1, Chr(34) + "idt" + Chr(34) + ":" + Chr(34) + "" & Format(!Vdt, "dd-MM-YYYY") & "" + Chr(34) + Chr(44)
             Print #1, Chr(34) + "val" + Chr(34) + ":" + "" & !Amt1 & "" + Chr(44)
             Print #1, Chr(34) + "pos" + Chr(34) + ":" + Chr(34) + "" & Left(!Nar1, 2) & "" + Chr(34) + Chr(44)
             Print #1, Chr(34) + "rchrg" + Chr(34) + ":" + Chr(34) + "N" + Chr(34) + Chr(44)
             Print #1, Chr(34) + "inv_typ" + Chr(34) + ":" + Chr(34) + "R" + Chr(34) + Chr(44)
             Print #1, Chr(34) + "itms" + Chr(34) + ":" + "[" + "{" + Chr(34) + "num" + Chr(34) + ":" + "1" + Chr(44)
             Print #1, Chr(34) + "itm_det" + Chr(34) + ":" + "{" + Chr(34) + "txval" + Chr(34) + ":" + "" & !AMT3 & "" + Chr(44)
             Print #1, Chr(34) + "rt" + Chr(34) + ":" + "" & !AMT2 & "" + Chr(44)
             If !Amt7 + !Amt8 > 0 Then
                Print #1, Chr(34) + "camt" + Chr(34) + ":" + "" & !Amt7 & "" + Chr(44)
                Print #1, Chr(34) + "samt" + Chr(34) + ":" + "" & !Amt8 & "" + Chr(44)
             End If
             If !Amt9 > 0 Then
                Print #1, Chr(34) + "iamt" + Chr(34) + ":" + "" & !Amt9 & "" + Chr(44)
             End If
           End If
           .MoveNext
        Loop
        If EJ = "J" Then
           Print #1, Chr(34) + "csamt" + Chr(34) + ":" + "0" + "}" + "}" + "]" + "}"
           Print #1, "]" + "}" + "]" + Chr(44)
        End If
   End With
End If
           
'--B2CL
           
RecYN = GProcRstOpen(tmp1, "select * from TmpGentbl where TmpGentbl.UserName ='" & gUserName & "' and Nar10 = 'B2CL'  order by Vdt,VbillNo", "R", gSetupCn)
If RecYN > 0 Then
   If EJ = "E" Then
      Set sheet = workbook.Sheets.Item("b2cl")
      If (sheet Is Nothing) Then
         MsgBox "Error b2cl Sheet"
      End If
      a = 5
   End If
   With tmp1
        .MoveFirst
        If EJ = "J" Then
           Print #1, Chr(34) + "b2cl" + Chr(34) + ":" + "["
        End If
LoopB2CL:
        If EJ = "J" Then
           a = 1
        End If
        PGstin = !Nar1
        Do While Not .EOF
           If !Nar1 <> PGstin Then
              If EJ = "J" And a > 1 Then
                 Print #1, Chr(34) + "csamt" + Chr(34) + ":" + "0" + "}" + "}" + "]" + "}" + "]" + "}" + Chr(44)
              End If
              GoTo LoopB2CL
           Else
             If a > 1 And EJ = "J" Then
                 Print #1, Chr(34) + "csamt" + Chr(34) + ":" + "0" + "}" + "}" + "]" + "}" + Chr(44)
             End If
           End If
           If EJ = "E" Then
               sheet.Cells(a, 1) = !VbillNo
                DoEvents
               sheet.Cells(a, 2) = Format(!Vdt, "dd-MMM-YYYY")
               sheet.Cells(a, 3) = !Amt1
               sheet.Cells(a, 4) = !Nar4
               sheet.Cells(a, 6) = !AMT2
               sheet.Cells(a, 7) = !AMT3
               sheet.Cells(a, 8) = 0
              a = a + 1
           Else
             If a = 1 Then
                Print #1, "{" + Chr(34) + "pos" + Chr(34) + ":" + Chr(34) + "" & Left(!Nar1, 2) & "" + Chr(34) + Chr(44)
                Print #1, Chr(34) + "inv" + Chr(34) + ":" + "["
             End If
             a = a + 1
             Print #1, "{" + Chr(34) + "inum" + Chr(34) + ":" + Chr(34) + "" & !VbillNo & "" + Chr(34) + Chr(44)
             Print #1, Chr(34) + "idt" + Chr(34) + ":" + Chr(34) + "" & Format(!Vdt, "dd-MM-YYYY") & "" + Chr(34) + Chr(44)
             Print #1, Chr(34) + "val" + Chr(34) + ":" + "" & !Amt1 & "" + Chr(44)
             Print #1, Chr(34) + "itms" + Chr(34) + ":" + "[" + "{" + Chr(34) + "num" + Chr(34) + ":" + "1" + Chr(44)
             Print #1, Chr(34) + "itm_det" + Chr(34) + ":" + "{" + Chr(34) + "txval" + Chr(34) + ":" + "" & !AMT3 & "" + Chr(44)
             Print #1, Chr(34) + "rt" + Chr(34) + ":" + "" & !AMT2 & "" + Chr(44)
             If !Amt7 + !Amt8 > 0 Then
                Print #1, Chr(34) + "camt" + Chr(34) + ":" + "" & !Amt7 & "" + Chr(44)
                Print #1, Chr(34) + "samt" + Chr(34) + ":" + "" & !Amt8 & "" + Chr(44)
             End If
             If !Amt9 > 0 Then
                Print #1, Chr(34) + "iamt" + Chr(34) + ":" + "" & !Amt9 & "" + Chr(44)
             End If
           End If
           .MoveNext
        Loop
        If EJ = "J" Then
           Print #1, Chr(34) + "csamt" + Chr(34) + ":" + "0" + "}" + "}" + "]" + "}"
           Print #1, "]" + "}" + "]" + Chr(44)
        End If
   End With
End If
'''--B2CS
'
'If GProcRstOpen(tmp1, "select * from TmpGentbl where TmpGentbl.UserName ='" & gUserName & "'  and Nar10 = 'B2CS' ", "R", gSetupCn) > 0 Then
'   Set sheet = workbook.Sheets.Item("b2cs")
'   If (sheet Is Nothing) Then
'      MsgBox "Error b2cs Sheet"
'   End If
'   a = 5
'   With tmp1
'        .MoveFirst
'        Do While Not .EOF
'           sheet.Cells(a, 1) = "OE"
'           DoEvents
'           sheet.Cells(a, 2) = !Nar2
'           sheet.Cells(a, 4) = !AMT2
'           sheet.Cells(a, 5) = !AMT3
'           sheet.Cells(a, 6) = 0
'           a = a + 1
'           'prgBar.Value = prgBar.Value + 1
'           .MoveNext
'        Loop
'   End With
'End If

'--CDNR

If GProcRstOpen(tmp1, "select * from TmpGentbl where TmpGentbl.UserName ='" & gUserName & "' and Nar10 = 'CDNR'  order by nar1,Vbillno,Vdt", "R", gSetupCn) > 0 Then
   If EJ = "E" Then
      a = 5
      Set sheet = workbook.Sheets.Item("cdnr")
      If (sheet Is Nothing) Then
          MsgBox "Error CDNR Sheet"
      End If
   End If
   With tmp1
        .MoveFirst
        If EJ = "J" Then
           Print #1, Chr(34) + "cdnr" + Chr(34) + ":" + "["
        End If
LoopCdnr:
        If EJ = "J" Then
           a = 1
        End If
        PGstin = !Nar1
        Do While Not .EOF
           If !Nar1 <> PGstin Then
              If EJ = "J" And a > 1 Then
                 Print #1, Chr(34) + "csamt" + Chr(34) + ":" + "0" + "}" + "}" + "]" + "}" + "]" + "}" + Chr(44)
              End If
              GoTo LoopCdnr
           Else
             If a > 1 And EJ = "J" Then
                 Print #1, Chr(34) + "csamt" + Chr(34) + ":" + "0" + "}" + "}" + "]" + "}" + Chr(44)
             End If
           End If
           If EJ = "E" Then
              sheet.Cells(a, 1) = !Nar1
              DoEvents
              sheet.Cells(a, 2) = !Nar7
              sheet.Cells(a, 3) = !Nar6
              sheet.Cells(a, 4) = Format(!Vdt, "dd-MMM-YYYY")
              sheet.Cells(a, 5) = !Nar2
              sheet.Cells(a, 6) = !Nar5
              sheet.Cells(a, 7) = "N"
              If Left(!Nar1, 2) = Left(gCGSTIN, 2) Then
                 sheet.Cells(a, 8) = "Regular B2B"
              Else
                 sheet.Cells(a, 8) = "Intra-State supplies attracting IGST"
              End If
              sheet.Cells(a, 9) = !Amt1
              sheet.Cells(a, 11) = !AMT2  '--GST % rate
              sheet.Cells(a, 12) = !AMT3  '--taxable
              sheet.Cells(a, 13) = 0  'cess
              a = a + 1
            Else
              If a = 1 Then
                 Print #1, "{" + Chr(34) + "ctin" + Chr(34) + ":" + Chr(34) + "" & !Nar1 & "" + Chr(34) + Chr(44)
                 Print #1, Chr(34) + "nt" + Chr(34) + ":" + "["
              End If
              a = a + 1
              Print #1, "{" + Chr(34) + "nt_num" + Chr(34) + ":" + Chr(34) + "" & !Nar6 & "" + Chr(34) + Chr(44)
              Print #1, Chr(34) + "nt_dt" + Chr(34) + ":" + Chr(34) + "" & Format(!Vdt, "dd-MM-YYYY") & "" + Chr(34) + Chr(44)
              Print #1, Chr(34) + "ntty" + Chr(34) + ":" + Chr(34) + !Nar2 + Chr(34) + Chr(44)
              Print #1, Chr(34) + "val" + Chr(34) + ":" + "" & !Amt1 & "" + Chr(44)
              Print #1, Chr(34) + "pos" + Chr(34) + ":" + Chr(34) + "" & Left(!Nar1, 2) & "" + Chr(34) + Chr(44)
              Print #1, Chr(34) + "rchrg" + Chr(34) + ":" + Chr(34) + "N" + Chr(34) + Chr(44)
              Print #1, Chr(34) + "inv_typ" + Chr(34) + ":" + Chr(34) + "R" + Chr(34) + Chr(44)
              Print #1, Chr(34) + "itms" + Chr(34) + ":" + "[" + "{" + Chr(34) + "num" + Chr(34) + ":" + "1" + Chr(44)
              Print #1, Chr(34) + "itm_det" + Chr(34) + ":" + "{" + Chr(34) + "txval" + Chr(34) + ":" + "" & !AMT3 & "" + Chr(44)
              Print #1, Chr(34) + "rt" + Chr(34) + ":" + "" & !AMT2 & "" + Chr(44)
              If !Amt7 + !Amt8 > 0 Then
                 Print #1, Chr(34) + "camt" + Chr(34) + ":" + "" & !Amt7 & "" + Chr(44)
                 Print #1, Chr(34) + "samt" + Chr(34) + ":" + "" & !Amt8 & "" + Chr(44)
              End If
              If !Amt9 > 0 Then
                 Print #1, Chr(34) + "iamt" + Chr(34) + ":" + "" & !Amt9 & "" + Chr(44)
              End If
            End If
           .MoveNext
        Loop
        If EJ = "J" Then
           Print #1, Chr(34) + "csamt" + Chr(34) + ":" + "0" + "}" + "}" + "]" + "}"
           Print #1, "]" + "}" + "]" + Chr(44)
        End If
   End With
End If

'--CDNUR
'If GProcRstOpen(tmp1, "select * from TmpGentbl where TmpGentbl.UserName ='" & gUserName & "' and Nar10 = 'CDNUR'  ", "R", gSetupCn) > 0 Then
'   Set sheet = workbook.Sheets.Item("cdnur")
'   If (sheet Is Nothing) Then
'       MsgBox "Error cdnur Sheet"
'   End If
'   a = 5
'   With tmp1
'        .MoveFirst
'        Do While Not .EOF
'           sheet.Cells(a, 1) = "B2CL"
'           DoEvents
'           sheet.Cells(a, 2) = !Nar6
'           sheet.Cells(a, 3) = Format(!Vdt, "dd-MMM-YYYY")
'           sheet.Cells(a, 4) = !Nar2
'           sheet.Cells(a, 5) = !Nar5  'State Name
'           sheet.Cells(a, 6) = !Amt1  'Amt
'           sheet.Cells(a, 8) = !AMT2  'GST %
'           sheet.Cells(a, 9) = !AMT3
'           sheet.Cells(a, 10) = 0
'           a = a + 1
'           ' prgBar.Value = prgBar.Value + 1
'           .MoveNext
'        Loop
'    End With
'End If
'--HSN (b2b)
If GProcRstOpen(tmp1, "select * from TmpGentbl where TmpGentbl.UserName ='" & gUserName & "' and Nar10 = 'HSN' order by nar2,Nar1 ", "R", gSetupCn) > 0 Then
   If EJ = "E" Then
      Set sheet = workbook.Sheets.Item("hsn(b2b)")
      If (sheet Is Nothing) Then
         MsgBox "Error hsn(b2b) Sheet"
      End If
      a = 5
   End If
   With tmp1
        .MoveFirst
        If EJ = "J" Then
           Print #1, Chr(34) + "hsn(b2b)" + Chr(34) + ":" + "{" + Chr(34) + "data" + Chr(34) + ":" + "["
        End If
LoopHsnB2b:
        If EJ = "J" Then
           a = 1
        End If
        Do While Not .EOF
           If EJ = "E" Then
               sheet.Cells(a, 1) = !Nar1  'Hsn
               DoEvents
               sheet.Cells(a, 2) = !Nar2  'Desc
               sheet.Cells(a, 3) = !Nar3  'Unit
               sheet.Cells(a, 4) = !Amt1  'Qty
               sheet.Cells(a, 5) = !AMT2  'Amt
               sheet.Cells(a, 6) = !Amt8  'Tax Rate
               sheet.Cells(a, 7) = !AMT3  'Taxable
               sheet.Cells(a, 8) = !Amt4  'Igst
               sheet.Cells(a, 9) = !Amt5  'Cgst
               sheet.Cells(a, 10) = !Amt6  'Sgst
               sheet.Cells(a, 11) = 0 'Csee
               a = a + 1
            Else
              Print #1, "{" + Chr(34) + "num" + Chr(34) + ":" + CStr(a) + Chr(44)
              Print #1, Chr(34) + "hsn_sc" + Chr(34) + ":" + Chr(34) + !Nar1 + Chr(34) + Chr(44)
              Print #1, Chr(34) + "desc" + Chr(34) + ":" + Chr(34) + !Nar2 + Chr(34) + Chr(44)
              Print #1, Chr(34) + "uqc" + Chr(34) + ":" + Chr(34) + !Nar3 + Chr(34) + Chr(44)
              Print #1, Chr(34) + "qty" + Chr(34) + ":" + CStr(!Amt1) + Chr(44)
              Print #1, Chr(34) + "rt" + Chr(34) + ":" + CStr(!Amt8) + Chr(44)
              Print #1, Chr(34) + "txval" + Chr(34) + ":" + CStr(!AMT3) + Chr(44)
              Print #1, Chr(34) + "iamt" + Chr(34) + ":" + CStr(!Amt4) + Chr(44)
              Print #1, Chr(34) + "samt" + Chr(34) + ":" + CStr(!Amt6) + Chr(44)
              Print #1, Chr(34) + "camt" + Chr(34) + ":" + CStr(!Amt5) + Chr(44)
              Print #1, Chr(34) + "csamt" + Chr(34) + ":" + CStr("0")
              a = a + 1
            End If
           .MoveNext
           If EJ = "J" Then
              If tmp1.EOF Then
                 Print #1, "}" + "]" + "}" + Chr(44)
              Else
                 Print #1, "}" + Chr(44)
              End If
           End If
        Loop

   End With
End If

'--HSN (b2c)
If GProcRstOpen(tmp1, "select * from TmpGentbl where TmpGentbl.UserName ='" & gUserName & "' and Nar10 = 'HSNB2C' order by nar2,Nar1 ", "R", gSetupCn) > 0 Then
   If EJ = "E" Then
      Set sheet = workbook.Sheets.Item("hsn(b2c)")
      If (sheet Is Nothing) Then
          MsgBox "Error hsn(b2c) Sheet"
      End If
      a = 5
   End If
   With tmp1
        .MoveFirst
        If EJ = "J" Then
           Print #1, Chr(34) + "hsn(b2c)" + Chr(34) + ":" + "{" + Chr(34) + "data" + Chr(34) + ":" + "["
        End If
LoopHsnB2c:
        If EJ = "J" Then
           a = 1
        End If
        Do While Not .EOF
           If EJ = "E" Then
               sheet.Cells(a, 1) = !Nar1  'Hsn
               DoEvents
               sheet.Cells(a, 2) = !Nar2  'Desc
               sheet.Cells(a, 3) = !Nar3  'Unit
               sheet.Cells(a, 4) = !Amt1  'Qty
               sheet.Cells(a, 5) = !AMT2  'Amt
               sheet.Cells(a, 6) = !Amt8  'Tax Rate
               sheet.Cells(a, 7) = !AMT3  'Taxable
               sheet.Cells(a, 8) = !Amt4  'Igst
               sheet.Cells(a, 9) = !Amt5  'Cgst
               sheet.Cells(a, 10) = !Amt6  'Sgst
               sheet.Cells(a, 11) = 0 'Csee
               a = a + 1
           Else
              Print #1, "{" + Chr(34) + "num" + Chr(34) + ":" + CStr(a) + Chr(44)
              Print #1, Chr(34) + "hsn_sc" + Chr(34) + ":" + Chr(34) + !Nar1 + Chr(34) + Chr(44)
              Print #1, Chr(34) + "desc" + Chr(34) + ":" + Chr(34) + !Nar2 + Chr(34) + Chr(44)
              Print #1, Chr(34) + "uqc" + Chr(34) + ":" + Chr(34) + !Nar3 + Chr(34) + Chr(44)
              Print #1, Chr(34) + "qty" + Chr(34) + ":" + CStr(!Amt1) + Chr(44)
              Print #1, Chr(34) + "rt" + Chr(34) + ":" + CStr(!Amt8) + Chr(44)
              Print #1, Chr(34) + "txval" + Chr(34) + ":" + CStr(!AMT3) + Chr(44)
              Print #1, Chr(34) + "iamt" + Chr(34) + ":" + CStr(!Amt4) + Chr(44)
              Print #1, Chr(34) + "samt" + Chr(34) + ":" + CStr(!Amt6) + Chr(44)
              Print #1, Chr(34) + "camt" + Chr(34) + ":" + CStr(!Amt5) + Chr(44)
              Print #1, Chr(34) + "csamt" + Chr(34) + ":" + CStr("0")
              a = a + 1
           End If
           .MoveNext
           If EJ = "J" Then
              If tmp1.EOF Then
                 Print #1, "}" + "]" + "}" + Chr(44)
              Else
                 Print #1, "}" + Chr(44)
              End If
           End If
        Loop
   End With
End If

'--DOCS Yarn Sales
HeadPrint = False
LoopIn = False
If EJ = "E" Then
   a = 5
Else
   a = 1
End If
If GProcRstOpen(tmp1, "select * from TmpGentbl where TmpGentbl.UserName ='" & gUserName & "' and Nar10 = 'DOCS-B2B-Yarn'  ", "R", gSetupCn) > 0 Then
   If EJ = "E" Then
      Set sheet = workbook.Sheets.Item("docs")
      If (sheet Is Nothing) Then
          MsgBox "Error docs Sheet"
      End If
   End If
   With tmp1
        .MoveFirst
        If LoopIn = False Then
           LoopIn = True
        End If
        If EJ = "J" And HeadPrint = False Then
           Print #1, Chr(34) + "doc_issue" + Chr(34) + ":" + "{" + Chr(34) + "doc_det" + Chr(34) + ":" + "["
           Print #1, "{" + Chr(34) + "doc_num" + Chr(34) + ":" + CStr(a) + Chr(44)
           Print #1, Chr(34) + "doc_typ" + Chr(34) + ":" + Chr(34) + "Invoices for outward supply" + Chr(34) + Chr(44)
           Print #1, Chr(34) + "docs" + Chr(34) + ":" + "["
           HeadPrint = True
        End If
        Do While Not .EOF
           If EJ = "E" Then
              sheet.Cells(a, 1) = !Nar1
              DoEvents
              sheet.Cells(a, 2) = !Nar2 + Trim(CStr(!Amt1))
              sheet.Cells(a, 3) = !Nar2 + Trim(CStr(!AMT2))
              sheet.Cells(a, 4) = !AMT3
              sheet.Cells(a, 5) = !AMT3 - !Amt4
              a = a + 1
           Else
              Print #1, "{" + Chr(34) + "num" + Chr(34) + ":" + CStr(a) + Chr(44)
              Print #1, Chr(34) + "from" + Chr(34) + ":" + Chr(34) + !Nar2 + Trim(CStr(!Amt1)) + Chr(34) + Chr(44)
              Print #1, Chr(34) + "to" + Chr(34) + ":" + Chr(34) + !Nar2 + Trim(CStr(!AMT2)) + Chr(34) + Chr(44)
              Print #1, Chr(34) + "totnum" + Chr(34) + ":" + Trim(CStr(!AMT3)) + Chr(44)
              Print #1, Chr(34) + "cancel" + Chr(34) + ":" + Trim(CStr(!AMT3 - !Amt4)) + Chr(44)
              Print #1, Chr(34) + "net_issue" + Chr(34) + ":" + Trim(CStr(!AMT3 - (!AMT3 - !Amt4))) + "}"
              a = a + 1
           End If
           .MoveNext
        Loop
   End With
End If


'--DOCS Other Sales
If GProcRstOpen(tmp1, "select * from TmpGentbl where TmpGentbl.UserName ='" & gUserName & "' and Nar10 = 'DOCS-B2B-Other'  ", "R", gSetupCn) > 0 Then
   If EJ = "E" Then
      Set sheet = workbook.Sheets.Item("docs")
      If (sheet Is Nothing) Then
         MsgBox "Error docs Sheet"
      End If
   End If
   If EJ = "J" Then
      If LoopIn = True Then
         Print #1, Chr(44)
      End If
   End If
   With tmp1
        .MoveFirst
        If EJ = "J" And HeadPrint = False Then
           Print #1, Chr(34) + "doc_issue" + Chr(34) + ":" + "{" + Chr(34) + "doc_det" + Chr(34) + ":" + "["
           Print #1, "{" + Chr(34) + "doc_num" + Chr(34) + ":" + "1" + Chr(44)
           Print #1, Chr(34) + "doc_typ" + Chr(34) + ":" + Chr(34) + "Invoices for outward supply" + Chr(34) + Chr(44)
           Print #1, Chr(34) + "docs" + Chr(34) + ":" + "["
           HeadPrint = True
        End If
        Do While Not .EOF
           If EJ = "E" Then
              sheet.Cells(a, 1) = !Nar1
              DoEvents
              sheet.Cells(a, 2) = !Nar2 + Trim(CStr(!Amt1))
              sheet.Cells(a, 3) = !Nar2 + Trim(CStr(!AMT2))
              sheet.Cells(a, 4) = !AMT3
              sheet.Cells(a, 5) = !AMT3 - !Amt4
              a = a + 1
           Else
              Print #1, "{" + Chr(34) + "num" + Chr(34) + ":" + CStr(a) + Chr(44)
              Print #1, Chr(34) + "from" + Chr(34) + ":" + Chr(34) + !Nar2 + Trim(CStr(!Amt1)) + Chr(34) + Chr(44)
              Print #1, Chr(34) + "to" + Chr(34) + ":" + Chr(34) + !Nar2 + Trim(CStr(!AMT2)) + Chr(34) + Chr(44)
              Print #1, Chr(34) + "totnum" + Chr(34) + ":" + Trim(CStr(!AMT3)) + Chr(44)
              Print #1, Chr(34) + "cancel" + Chr(34) + ":" + Trim(CStr(!AMT3 - !Amt4)) + Chr(44)
              Print #1, Chr(34) + "net_issue" + Chr(34) + ":" + Trim(CStr(!AMT3 - (!AMT3 - !Amt4))) + "}"
              a = a + 1
           End If
           .MoveNext
        Loop
   End With
End If
'--DOCS Cotton Sales
If GProcRstOpen(tmp1, "select * from TmpGentbl where TmpGentbl.UserName ='" & gUserName & "' and Nar10 = 'DOCS-B2B-Cot'  ", "R", gSetupCn) > 0 Then
   If EJ = "E" Then
      Set sheet = workbook.Sheets.Item("docs")
      If (sheet Is Nothing) Then
          MsgBox "Error docs Sheet"
      End If
   End If
   If EJ = "J" Then
      If LoopIn = True Then
         Print #1, Chr(44)
      End If
   End If
   With tmp1
        .MoveFirst
        If EJ = "J" And HeadPrint = False Then
           Print #1, Chr(34) + "doc_issue" + Chr(34) + ":" + "{" + Chr(34) + "doc_det" + Chr(34) + ":" + "["
           Print #1, "{" + Chr(34) + "doc_num" + Chr(34) + ":" + "1" + Chr(44)
           Print #1, Chr(34) + "doc_typ" + Chr(34) + ":" + Chr(34) + "Invoices for outward supply" + Chr(34) + Chr(44)
           Print #1, Chr(34) + "docs" + Chr(34) + ":" + "["
           HeadPrint = True
        End If
        Do While Not .EOF
           If EJ = "E" Then
              sheet.Cells(a, 1) = !Nar1
              DoEvents
              sheet.Cells(a, 2) = !Nar2 + Trim(CStr(!Amt1))
              sheet.Cells(a, 3) = !Nar2 + Trim(CStr(!AMT2))
              sheet.Cells(a, 4) = !AMT3
              sheet.Cells(a, 5) = !AMT3 - !Amt4
              a = a + 1
           Else
              Print #1, "{" + Chr(34) + "num" + Chr(34) + ":" + CStr(a) + Chr(44)
              Print #1, Chr(34) + "from" + Chr(34) + ":" + Chr(34) + !Nar2 + Trim(CStr(!Amt1)) + Chr(34) + Chr(44)
              Print #1, Chr(34) + "to" + Chr(34) + ":" + Chr(34) + !Nar2 + Trim(CStr(!AMT2)) + Chr(34) + Chr(44)
              Print #1, Chr(34) + "totnum" + Chr(34) + ":" + Trim(CStr(!AMT3)) + Chr(44)
              Print #1, Chr(34) + "cancel" + Chr(34) + ":" + Trim(CStr(!AMT3 - !Amt4)) + Chr(44)
              Print #1, Chr(34) + "net_issue" + Chr(34) + ":" + Trim(CStr(!AMT3 - (!AMT3 - !Amt4))) + "}"
              a = a + 1
           End If
           .MoveNext
        Loop
   End With
End If
'--DOCS Cloth Sales
If GProcRstOpen(tmp1, "select * from TmpGentbl where TmpGentbl.UserName ='" & gUserName & "' and Nar10 = 'DOCS-B2B-Clo'  ", "R", gSetupCn) > 0 Then
   If EJ = "E" Then
      Set sheet = workbook.Sheets.Item("docs")
      If (sheet Is Nothing) Then
          MsgBox "Error docs Sheet"
      End If
   End If
   If EJ = "J" Then
      If LoopIn = True Then
         Print #1, Chr(44)
      End If
   End If
   With tmp1
        .MoveFirst
        If EJ = "J" And HeadPrint = False Then
           Print #1, Chr(34) + "doc_issue" + Chr(34) + ":" + "{" + Chr(34) + "doc_det" + Chr(34) + ":" + "["
           Print #1, "{" + Chr(34) + "doc_num" + Chr(34) + ":" + "1" + Chr(44)
           Print #1, Chr(34) + "doc_typ" + Chr(34) + ":" + Chr(34) + "Invoices for outward supply" + Chr(34) + Chr(44)
           Print #1, Chr(34) + "docs" + Chr(34) + ":" + "["
           HeadPrint = True
        End If
        Do While Not .EOF
           If EJ = "E" Then
              sheet.Cells(a, 1) = !Nar1
              DoEvents
              sheet.Cells(a, 2) = !Nar2 + Trim(CStr(!Amt1))
              sheet.Cells(a, 3) = !Nar2 + Trim(CStr(!AMT2))
              sheet.Cells(a, 4) = !AMT3
              sheet.Cells(a, 5) = !AMT3 - !Amt4
              a = a + 1
           Else
              Print #1, "{" + Chr(34) + "num" + Chr(34) + ":" + CStr(a) + Chr(44)
              Print #1, Chr(34) + "from" + Chr(34) + ":" + Chr(34) + !Nar2 + Trim(CStr(!Amt1)) + Chr(34) + Chr(44)
              Print #1, Chr(34) + "to" + Chr(34) + ":" + Chr(34) + !Nar2 + Trim(CStr(!AMT2)) + Chr(34) + Chr(44)
              Print #1, Chr(34) + "totnum" + Chr(34) + ":" + Trim(CStr(!AMT3)) + Chr(44)
              Print #1, Chr(34) + "cancel" + Chr(34) + ":" + Trim(CStr(!AMT3 - !Amt4)) + Chr(44)
              Print #1, Chr(34) + "net_issue" + Chr(34) + ":" + Trim(CStr(!AMT3 - (!AMT3 - !Amt4))) + "}"
              a = a + 1
           End If
           .MoveNext
        Loop
   End With
End If
If EJ = "J" Then
   If LoopIn = True Then
      Print #1, "]" + "}" + Chr(44)
   End If
End If


'--DOCS Debit Note
LoopIn = False
If EJ = "J" Then
   a = 1
End If
HeadPrint = False
If GProcRstOpen(tmp1, "select * from TmpGentbl where TmpGentbl.UserName ='" & gUserName & "' and Nar10 = 'DOCS-B2B-Drn'  ", "R", gSetupCn) > 0 Then
   If EJ = "E" Then
      Set sheet = workbook.Sheets.Item("docs")
      If (sheet Is Nothing) Then
          MsgBox "Error docs Sheet"
      End If
   End If
   If EJ = "J" Then
      If LoopIn = True Then
         Print #1, Chr(44)
      End If
   End If
   With tmp1
        .MoveFirst
        If LoopIn = False Then
           LoopIn = True
        End If
        If EJ = "J" And HeadPrint = False Then
           Print #1, "{" + Chr(34) + "doc_num" + Chr(34) + ":" + "4" + Chr(44)
           Print #1, Chr(34) + "doc_typ" + Chr(34) + ":" + Chr(34) + "Debit Note" + Chr(34) + Chr(44)
           Print #1, Chr(34) + "docs" + Chr(34) + ":" + "["
           HeadPrint = True
        End If
        Do While Not .EOF
           If EJ = "E" Then
              sheet.Cells(a, 1) = !Nar1
              DoEvents
              sheet.Cells(a, 2) = !Nar2 + Trim(CStr(!Amt1))
              sheet.Cells(a, 3) = !Nar2 + Trim(CStr(!AMT2))
              sheet.Cells(a, 4) = !AMT3
              sheet.Cells(a, 5) = !AMT3 - !Amt4
              a = a + 1
           Else
              Print #1, "{" + Chr(34) + "num" + Chr(34) + ":" + CStr(a) + Chr(44)
              Print #1, Chr(34) + "from" + Chr(34) + ":" + Chr(34) + !Nar2 + Trim(CStr(!Amt1)) + Chr(34) + Chr(44)
              Print #1, Chr(34) + "to" + Chr(34) + ":" + Chr(34) + !Nar2 + Trim(CStr(!AMT2)) + Chr(34) + Chr(44)
              Print #1, Chr(34) + "totnum" + Chr(34) + ":" + Trim(CStr(!AMT3)) + Chr(44)
              Print #1, Chr(34) + "cancel" + Chr(34) + ":" + Trim(CStr(!AMT3 - !Amt4)) + Chr(44)
              Print #1, Chr(34) + "net_issue" + Chr(34) + ":" + Trim(CStr(!AMT3 - (!AMT3 - !Amt4))) + "}"
              a = a + 1
           End If
           .MoveNext
        Loop
   End With
End If
If EJ = "J" Then
   If LoopIn = True Then
      Print #1, "]" + "}" + Chr(44)
   End If
End If
LoopIn = False

If EJ = "J" Then
   a = 1
End If
HeadPrint = False
'--DOCS Credit Note
If GProcRstOpen(tmp1, "select * from TmpGentbl where TmpGentbl.UserName ='" & gUserName & "' and Nar10 = 'DOCS-B2B-Crn'  ", "R", gSetupCn) > 0 Then
   If EJ = "E" Then
      Set sheet = workbook.Sheets.Item("docs")
      If (sheet Is Nothing) Then
          MsgBox "Error docs Sheet"
      End If
   End If
   If EJ = "J" Then
      If LoopIn = True Then
         Print #1, Chr(44)
      End If
   End If
   With tmp1
        .MoveFirst
        If LoopIn = False Then
           LoopIn = True
        End If
        If EJ = "J" And HeadPrint = False Then
           Print #1, "{" + Chr(34) + "doc_num" + Chr(34) + ":" + "5" + Chr(44)
           Print #1, Chr(34) + "doc_typ" + Chr(34) + ":" + Chr(34) + "Credit Note" + Chr(34) + Chr(44)
           Print #1, Chr(34) + "docs" + Chr(34) + ":" + "["
           HeadPrint = True
        End If
        Do While Not .EOF
           If EJ = "E" Then
              sheet.Cells(a, 1) = !Nar1
              DoEvents
              sheet.Cells(a, 2) = !Nar2 + Trim(CStr(!Amt1))
              sheet.Cells(a, 3) = !Nar2 + Trim(CStr(!AMT2))
              sheet.Cells(a, 4) = !AMT3
              sheet.Cells(a, 5) = !AMT3 - !Amt4
              a = a + 1
           Else
              Print #1, "{" + Chr(34) + "num" + Chr(34) + ":" + CStr(a) + Chr(44)
              Print #1, Chr(34) + "from" + Chr(34) + ":" + Chr(34) + !Nar2 + Trim(CStr(!Amt1)) + Chr(34) + Chr(44)
              Print #1, Chr(34) + "to" + Chr(34) + ":" + Chr(34) + !Nar2 + Trim(CStr(!AMT2)) + Chr(34) + Chr(44)
              Print #1, Chr(34) + "totnum" + Chr(34) + ":" + Trim(CStr(!AMT3)) + Chr(44)
              Print #1, Chr(34) + "cancel" + Chr(34) + ":" + Trim(CStr(!AMT3 - !Amt4)) + Chr(44)
              Print #1, Chr(34) + "net_issue" + Chr(34) + ":" + Trim(CStr(!AMT3 - (!AMT3 - !Amt4))) + "}"
              a = a + 1
           End If
           .MoveNext
        Loop
   End With
End If
If EJ = "J" Then
   If LoopIn = True Then
      Print #1, "]" + "}" + Chr(44)
   End If
End If
If EJ = "J" Then
   a = 1
End If
HeadPrint = False
LoopIn = False

'--DOCS Credit Note (Sales Rtn.)
If GProcRstOpen(tmp1, "select * from TmpGentbl where TmpGentbl.UserName ='" & gUserName & "' and Nar10 = 'DOCS-B2B-SlRtn'  ", "R", gSetupCn) > 0 Then
   If EJ = "E" Then
      Set sheet = workbook.Sheets.Item("docs")
      If (sheet Is Nothing) Then
          MsgBox "Error docs Sheet"
      End If
   End If
   If EJ = "J" Then
      If LoopIn = True Then
         Print #1, Chr(44)
      End If
   End If
   With tmp1
        .MoveFirst
        If LoopIn = False Then
           LoopIn = True
        End If
        If EJ = "J" And (!Amt1 + !AMT3 > 0) And HeadPrint = False Then
           Print #1, "{" + Chr(34) + "doc_num" + Chr(34) + ":" + "5" + Chr(44)
           Print #1, Chr(34) + "doc_typ" + Chr(34) + ":" + Chr(34) + "Credit Note" + Chr(34) + Chr(44)
           Print #1, Chr(34) + "docs" + Chr(34) + ":" + "["
           HeadPrint = True
        End If
        Do While Not .EOF
           If !Amt1 + !AMT3 > 0 Then
              If EJ = "E" Then
                 sheet.Cells(a, 1) = !Nar1
                 DoEvents
                 sheet.Cells(a, 2) = Trim(CStr(!Amt1))
                 sheet.Cells(a, 3) = Trim(CStr(!AMT2))
                 sheet.Cells(a, 4) = !AMT2 - !Amt1 + 1
                 sheet.Cells(a, 5) = (!AMT2 - !Amt1 + 1) - !AMT3
                 a = a + 1
              Else
                 Print #1, "{" + Chr(34) + "num" + Chr(34) + ":" + CStr(a) + Chr(44)
                 Print #1, Chr(34) + "from" + Chr(34) + ":" + Chr(34) + Trim(CStr(!Amt1)) + Chr(34) + Chr(44)
                 Print #1, Chr(34) + "to" + Chr(34) + ":" + Chr(34) + Trim(CStr(!AMT2)) + Chr(34) + Chr(44)
                 Print #1, Chr(34) + "totnum" + Chr(34) + ":" + Trim(CStr(!AMT2 - !Amt1 + 1)) + Chr(44)
                 Print #1, Chr(34) + "cancel" + Chr(34) + ":" + Trim(CStr((!AMT2 - !Amt1 + 1) - !AMT3)) + Chr(44)
                 Print #1, Chr(34) + "net_issue" + Chr(34) + ":" + Trim(CStr(!AMT3)) + "}"
                 a = a + 1
              End If
           End If
           .MoveNext
        Loop
   End With
End If
If EJ = "J" Then
   Print #1, "]" + "}" + "]" + "}"
   Print #1, "}"
   
   Close #1
Else
   workbook.Close SaveChanges:=True
   excel_app.Quit
End If

MsgBox "Compileted"
End Function




