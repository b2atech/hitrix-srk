Attribute VB_Name = "DataCollection"
Option Explicit
Public Sub CollectData(Index As Integer)
Dim LRst1 As Recordset, LRst2 As Recordset
gCn.Execute "delete from tmptbl1 where vusername = '" & gUserName & "'"
'GProcRstOpen LRst1, "Select * from tmptbl1", "O", gCn
Select Case Index
Case 1
End Select
End Sub
Public Sub tblVoucher_TO_tmptblFin(m_ToDate As Date)
Dim L_Rst1 As Recordset, L_Rst2 As Recordset
Dim X As String
gCn.Execute "delete from tmptblFin where username = '" & gUserName & "'"
gCn.Execute "Insert into tmptblFin(Accode,Aname,Agcode,Schedule,TopGroup,Actype,username) select Agcode,Agname,GpCode,Schedule,TopGroup,'G','" & gUserName & "' from tblmastgroup"
GProcRstOpen L_Rst1, "Select * from tmptblFin where username = '" & gUserName & "'", "O", gCn
If gBackEndDB = gBackEndAccess Then
    X = " VDT <=cdate('" & Format(m_ToDate, "dd/mm/yyyy") & "')"
Else
    X = " Vdt<=to_date('" & CDate(m_ToDate) & "','dd/MM/yyyy')"
End If

'For Liability Side Accounts
gCn.Execute "Insert into tmptblFin(Actype,username,Accode,Aname,Agcode,TopGroup,PfPer,ClBal,Clcd) " _
& " select 'A','" & gUserName & "',Accode,Acname,A.Agcode,TopGroup,AcIntPer,sum(vamt),'C' " _
& " from tblVoucher,tblmastaccount A,tblmastGroup where tblVoucher.vaccode = A.accode and " _
& " tblmastgroup.agcode = A.agcode and " & X & " and vyear = " & gCYear & " and TopGroup = 90001 " _
& " group by Accode,Acname,A.Agcode,TopGroup,AcIntPer having sum(vamt) > 0"

'For Assets Side Accounts without Stock
gCn.Execute "Insert into tmptblFin(Actype,username,Accode,Aname,Agcode,TopGroup,PfPer,ClBal,Clcd) " _
& " select 'A','" & gUserName & "',Accode,Acname,A.Agcode,TopGroup,AcIntPer,sum(vamt)*-1,'D' " _
& " from tblVoucher,tblmastaccount A,tblmastGroup where tblVoucher.vaccode = A.accode and (A.AcCode <> 90018) and (A.AgCode <> 90018) and " _
& " tblmastgroup.agcode = A.agcode and " & X & " and vyear = " & gCYear & " and  TopGroup = 90002 " _
& " group by Accode,Acname,A.Agcode,TopGroup,AcIntPer having sum(vamt) < 0 "
'For Opening Stock Only
gCn.Execute "Insert into tmptblFin(Actype,username,Accode,Aname,Agcode,TopGroup,OpBal,Opcd) " _
& " select 'A','" & gUserName & "',Accode,Acname,A.Agcode,90001,sum(vamt)*-1,'D' " _
& " from tblVoucher,tblmastaccount A,tblmastGroup where tblVoucher.vaccode = A.accode and ((A.AcCode = 90018) or (A.AgCode = 90018)) and " _
& " tblmastgroup.agcode = A.agcode and " & X & " and vyear = " & gCYear & "" _
& " group by Accode,Acname,A.Agcode,TopGroup,AcIntPer having sum(vamt) < 0  "

'For Liability Side Accounts which are Debit Balance
gCn.Execute "Insert into tmptblFin(Actype,username,Accode,Aname,Agcode,TopGroup,PfPer,ClBal,Clcd) " _
& " select 'A','" & gUserName & "',Accode,Acname,A.Agcode1,TopGroup,AcIntPer,sum(vamt)*-1,'D' " _
& " from tblVoucher,tblmastaccount A,tblmastGroup where tblVoucher.vaccode = A.accode and " _
& " tblmastgroup.agcode = A.agcode and " & X & " and vyear = " & gCYear & " and TopGroup = 90001 " _
& " group by Accode,Acname,A.Agcode1,TopGroup,AcIntPer having sum(vamt) < 0 "

'For Assets Side Accounts which are Credit Balance
gCn.Execute "Insert into tmptblFin(Actype,username,Accode,Aname,Agcode,TopGroup,PfPer,ClBal,Clcd) " _
& " select 'A','" & gUserName & "',Accode,Acname,A.Agcode1,TopGroup,AcIntPer,sum(vamt),'C' " _
& " from tblVoucher,tblmastaccount A,tblmastGroup where tblVoucher.vaccode = A.accode and " _
& " tblmastgroup.agcode = A.agcode and " & X & " and vyear = " & gCYear & " and TopGroup = 90002 " _
& " group by Accode,Acname,A.Agcode1,TopGroup,AcIntPer having sum(vamt) > 0 "

'For Credit Side of P&L
gCn.Execute "Insert into tmptblFin(Actype,username,Accode,Aname,Agcode,TopGroup,PfPer,ClBal,Clcd) " _
& " select 'A','" & gUserName & "',Accode,Acname,A.Agcode,TopGroup,AcIntPer,sum(vamt),'C' " _
& " from tblVoucher,tblmastaccount A,tblmastGroup where tblVoucher.vaccode = A.accode and " _
& " tblmastgroup.agcode = A.agcode and " & X & " and vyear = " & gCYear & " and TopGroup = 90003 " _
& " group by Accode,Acname,A.Agcode,TopGroup,AcIntPer having sum(vamt) > 0"

'For Debit Side Accounts of P&L
gCn.Execute "Insert into tmptblFin(Actype,username,Accode,Aname,Agcode,TopGroup,PfPer,ClBal,Clcd) " _
& " select 'A','" & gUserName & "',Accode,Acname,A.Agcode,TopGroup,AcIntPer,sum(vamt)*-1,'D' " _
& " from tblVoucher,tblmastaccount A,tblmastGroup where tblVoucher.vaccode = A.accode and " _
& " tblmastgroup.agcode = A.agcode and " & X & " and vyear = " & gCYear & " and  TopGroup = 90003 " _
& " group by Accode,Acname,A.Agcode,TopGroup,AcIntPer having sum(vamt) < 0 "

'If (GProcRstOpen(L_Rst2, "Select Accode,A.Agcode,Agcode1,TopGroup,Acname,AcIntPer,sum(vamt) as amt1 " _
'& " from tblVoucher,tblmastaccount A,tblmastGroup where tblVoucher.vaccode = A.accode and " _
'& " tblmastgroup.agcode = A.agcode and " & X & " and vyear = " & gCYear & " group by Accode,A.Agcode,Agcode1,TopGroup,Acname,AcIntPer order by accode", "R", gCn) > 0) Then
'With L_Rst2
'    .MoveFirst
'    Do While Not .EOF
'        L_Rst1.AddNew
'        L_Rst1!actype = "A"
'        L_Rst1!UserName = gUserName
'        L_Rst1!AcCode = !AcCode
'        L_Rst1!AgCode = !AgCode
'        L_Rst1!Aname = !AcName
'        L_Rst1!TopGroup = !TopGroup
'        If (!AcCode = 4) Then
'            L_Rst1!PfPer = !AcIntPer
'        End If
'        If (!AcCode = 90018) Or (!AgCode = 90018) Then
'            L_Rst1!OpBal = Abs(!Amt1)
'        Else
'            L_Rst1!CLBAL = Abs(!Amt1)
'        End If
'        If (!Amt1 > 0) Then
'           L_Rst1!CLCD = "C"
'           If !TopGroup = 90001 Then
'                L_Rst1!AgCode = !AgCode1
'           End If
'        Else
'           L_Rst1!CLCD = "D"
'           If !TopGroup = 90002 Then
'                L_Rst1!AgCode = !AgCode1
'           End If
'        End If
'        L_Rst1.Update
'        .MoveNext
'    Loop
'    .Close
'End With
'End If
If gBackEndDB = gBackEndAccess Then
    X = " Stdate =cdate('" & Format(m_ToDate, "dd/mm/yyyy") & "')"
Else
    X = " Stdate=to_date('" & CDate(m_ToDate) & "','dd/MM/yyyy')"
End If
If (GProcRstOpen(L_Rst2, "Select B.*,AcName from tblclosingstock B,tblmastaccount A where B.staccode = A.accode and " & X, "R", gCn) > 0) Then
With L_Rst2
    .MoveFirst
    Do While Not .EOF
        L_Rst1.AddNew
        L_Rst1!actype = "A"
        L_Rst1!UserName = gUserName
        L_Rst1!AcCode = !StAcCode
        L_Rst1!AgCode = 90018
        L_Rst1!Aname = !AcName
        L_Rst1!TopGrp = 1
        L_Rst1!CLBAL = !StAmount
        L_Rst1!CLCD = "D"
        L_Rst1.Update
        .MoveNext
    Loop
    .Close
End With
End If
L_Rst1.Close
End Sub
Public Sub CreateMofidyViews()
Dim X1 As String
Dim X2 As String
Dim RsTmp As Recordset

'-- If Temp Table Exist Then Drop It
If GProcRstOpen(RsTmp, "Select * from Tab Where TName='TMPGENTBL'", "R") > 0 Then gCn.Execute "Drop Table tmpGentbl"
If GProcRstOpen(RsTmp, "Select * from Tab Where TName='TMPRPTTBL'", "R") > 0 Then gCn.Execute "Drop Table tmpRptTbl"
If GProcRstOpen(RsTmp, "Select * from Tab Where TName='TMPTBLFIN'", "R") > 0 Then gCn.Execute "Drop Table tmptblFin"
If GProcRstOpen(RsTmp, "Select * from Tab Where TName='TMPSELECTION'", "R") > 0 Then gCn.Execute "Drop Table tmpSelection"

'tmptblFin
gCn.Execute "Create Table tmptblFin(AcType     Varchar2 (1) ," _
                                & " AcCode     Number (7,0) Default 0, " _
                                & " AgCode     Number (7,0) Default 0," _
                                & " AName      Varchar2 (40) ," _
                                & " Schedule   Varchar2 (10) ," _
                                & " ClBal      Number (15,2) Default 0," _
                                & " ClCd       Varchar2 (1)," _
                                & " OpBal      Number (15,2) Default 0," _
                                & " OpCd       Varchar2 (1)," _
                                & " TopGroup   Number (7,0) Default 0," _
                                & " PfPer      Number (6,2) Default 0," _
                                & " UserName   Varchar2 (10) )"
'tmpDtSel
gCn.Execute "Create Table tmpSelection(Dt1  Date Default Null,Dt2  Date Default Null,UserName   Varchar2 (10) Default Null)"
'tmpGenTbl
gCn.Execute "Create Table tmpGenTbl(Vno     Number (7,0) Default 0 ," _
                                & " VDt      Date Default SYSDate," _
                                & " Vtype    Varchar2 (2) Default Null ," _
                                & " VAcCode  Number (7,0) Default 0," _
                                & " VAcOCode Number (7,0) Default 0," _
                                & " VBillNo  Varchar2 (10) Default Null," _
                                & " VBillDt  Date Default Null," _
                                & " Nar1    Varchar2 (80) Default Null," _
                                & " Nar2    Varchar2 (40) Default Null," _
                                & " Nar3    Varchar2 (40) Default Null," _
                                & " Nar4    Varchar2 (40) Default Null," _
                                & " Nar5    Varchar2 (40) Default Null," _
                                & " Nar6    Varchar2 (40) Default Null," _
                                & " Amt1     Number (15,3) Default 0, Amt2     Number (15,3) Default 0 ," _
                                & " Amt3     Number (15,3) Default 0, Amt4     Number (15,3) Default 0," _
                                & " Amt5     Number (15,3) Default 0, Amt6     Number (15,3) Default 0," _
                                & " Amt7     Number (15,3) Default 0, Amt8     Number (15,3) Default 0," _
                                & " Amt9     Number (15,3) Default 0, Amt10     Number (15,3) Default 0," _
                                & " Amt11     Number (15,3) Default 0, Amt12     Number (15,3) Default 0," _
                                & " Amt13     Number (15,3) Default 0, Amt14     Number (15,3) Default 0," _
                                & " VSrNo   Number (4,0) Default 0, VYear      Number (4,0)  Default 0,UserName   Varchar2 (10) Default Null," _
                                & " Dt1  Date Default Null,Dt2  Date Default Null, NAR7  VARCHAR2(100) Default Null,NAR8 VARCHAR2(40) Default Null, " _
        & " NAR9 VARCHAR2(40) Default Null, NAR10 VARCHAR2(40) Default Null, NAR11 VARCHAR2(40) Default Null, NAR12 VARCHAR2(40) Default Null, NAR13  VARCHAR2(255) Default Null )"
        'tmpRptTbl
gCn.Execute "Create Table tmpRptTbl(Vno     Number (7,0) Default 0 ," _
                                & " VDt      Date Default SYSDate," _
                                & " Vtype    Varchar2 (2) Default Null ," _
                                & " VAcCode  Number (7,0) Default 0," _
                                & " VAcOCode Number (7,0) Default 0," _
                                & " VBillNo  Varchar2 (10) Default Null," _
                                & " VBillDt  Date Default Null," _
                                & " Nar1    Varchar2 (250) Default Null," _
                                & " Nar2    Varchar2 (40) Default Null," _
                                & " Nar3    Varchar2 (40) Default Null," _
                                & " Nar4    Varchar2 (40) Default Null," _
                                & " Nar5    Varchar2 (40) Default Null," _
                                & " Nar6    Varchar2 (40) Default Null," _
                                & " Amt1     Number (15,3) Default 0, Amt2     Number (15,3) Default 0 ," _
                                & " Amt3     Number (15,3) Default 0, Amt4     Number (15,3) Default 0," _
                                & " Amt5     Number (15,3) Default 0, Amt6     Number (15,3) Default 0," _
                                & " Amt7     Number (15,3) Default 0, Amt8     Number (15,3) Default 0," _
                                & " Amt9     Number (15,3) Default 0, Amt10     Number (15,3) Default 0," _
                                & " Amt11     Number (15,3) Default 0, Amt12     Number (15,3) Default 0," _
                                & " Amt13     Number (15,3) Default 0, Amt14     Number (15,3) Default 0," _
                                & " VSrNo   Number (4,0) Default 0, VYear      Number (4,0)  Default 0,UserName   Varchar2 (10) Default Null," _
                                & " Dt1  Date Default Null,Dt2  Date Default Null, NAR7  VARCHAR2(80) Default Null,NAR8 VARCHAR2(40) Default Null, " _
        & " NAR9 VARCHAR2(40) Default Null, NAR10 VARCHAR2(40) Default Null, NAR11 VARCHAR2(40) Default Null, NAR12 VARCHAR2(40) Default Null, NAR13  VARCHAR2(80) Default Null )"
'1> qryInwardVsGatePass
gCn.Execute ("Create or replace View qryInwardVsGatePass as SELECT tblVouSub.VNo AS GpVNo, tblVouSub.VType AS GpVType, tblVouSub.VDt AS GpVDt, " _
           & " tblVouSub.VYear AS GpVYear, tblMastAccount_1.AcName AS GpParty, tblMastAccount.AcName AS GpBroker, tblMastAccount_2.AcName AS GpMill, " _
           & " tblVouSub.VSubBag AS GpBag, tblVouSub.VSubWt AS GpWt, tblVouSub.VSubRt AS GpRate, tblVouSub.VSubPItCtrlNo AS GpPITCtrlNo, " _
           & " tblAddLess.AdGpNo AS GpDueDay, tblVouSub.VSubGpSLType AS GpSLType, tblAddLess_1.VNo AS PVNo, tblAddLess_1.VType AS PVType, " _
           & " tblAddLess_1.VDt AS PVDt, tblAddLess_1.VYear AS PVYear, tblMastItem.ItName AS PCount, tblVouSub_2.vSubLotNo AS PLotNo, " _
           & " tblVouSub_2.VSubLRNo AS PLRNo, tblVouSub_2.VSubLRDate AS PLRDt, tblAddLess_1.AdExemptAmt AS PExemptAmt, " _
           & " tblAddLess_1.AdExemptPerKg AS PExemptKg, tblMastNarration.Narration AS PGodown, tblVouSub_2.VSubItCtrlNo AS pItCtrlNo, " _
           & " tblMastNarration_1.Narration AS PTaxType, tblAddLess_1.AdTaxP AS PTaxP, tblVouSub.VSubRtPer AS GPRatePer,tblVouSub_2.VSubLRAmt AS PLRAmt, tblVouSub.VSubItCtrlNo as GpItCtrlno " _
           & " FROM tblMastItem INNER JOIN (tblMastAccount INNER JOIN (tblMastNarration tblMastNarration_1 INNER JOIN " _
           & "(tblMastNarration INNER JOIN (tblMastAccount tblMastAccount_2 INNER JOIN (tblMastAccount  tblMastAccount_1 INNER JOIN " _
           & "(tblAddLess INNER JOIN (tblAddLess  tblAddLess_1 INNER JOIN (tblVouSub  tblVouSub_2 INNER JOIN " _
           & " tblVouSub ON tblVouSub_2.VSubItCtrlNo = tblVouSub.VSubPItCtrlNo) ON (tblAddLess_1.VNo = tblVouSub_2.VNo) AND " _
           & "(tblAddLess_1.VType = tblVouSub_2.VType) AND (tblAddLess_1.VYear = tblVouSub_2.VYear)) ON (tblAddLess.VNo = tblVouSub.VNo) " _
           & "AND (tblAddLess.VYear = tblVouSub.VYear) AND (tblAddLess.VType = tblVouSub.VType)) ON tblMastAccount_1.AcCode = tblAddLess.AdAcCrCode) " _
           & "ON tblMastAccount_2.AcCode = tblAddLess.AdMillCode) ON tblMastNarration.NarrCode = tblVouSub_2.VSubGodownCode) ON " _
           & "tblMastNarration_1.NarrCode = tblAddLess_1.AdTaxCode) ON tblMastAccount.AcCode = tblAddLess.AdAcDrCode) ON tblMastItem.ItCode = tblVouSub_2.VSubItCode " _
           & " WHERE tblVouSub.VType='GP'")
'2> qryMastAccount
gCn.Execute ("Create or replace View qryMastAccount as SELECT tblMastAccount.AcName AS Account,tblMastGroup.AgName,GPCode,decode(AcMillType,0,'Trade',1,'Consignment',2,'Depot',3,'Mill Bill',4,'SIT',5,'Trade & SIT')  as AcMillTypeStr ,tblMastAccount.* " _
          & " FROM tblMastGroup INNER JOIN tblMastAccount ON tblMastGroup.AgCode = tblMastAccount.AgCode " _
          & " UNION ALL (select AcAlName ,tblMastGroup.AgName,GpCode,decode(AcMillType,0,'Trade',1,'Consignment',2,'Depot',3,'Mill Bill',4,'SIT',5,'Trade & SIT')  as AcMillTypeStr ,tblMastAccount.* " _
          & " FROM tblMastGroup INNER JOIN tblMastAccount ON tblMastGroup.AgCode = tblMastAccount.AgCode where AcAlName<>AcName and AcAlName<>'' )")

'3> qryOpBal
gCn.Execute ("Create or replace View qryOpBal as SELECT acname, vaccode, decode(sign(vamt),-1,abs(vamt),0) AS Debit, decode(sign(vamt), 1,vamt,0) AS Credit " _
          & " From tblvoucher, tblmastaccount WHERE tblvoucher.vaccode=tblmastaccount.accode And vtype='0P'")
'4> qryAccountLedger '-(dr ,Cr & 0P  for TmpGebTbl Date Range (dt1 - dt2 ) & user  )
'-- TmpRptTbl prepare TmpRptTbl For Rec Vs Sale Narration
gCn.Execute ("Create or replace View qryAccountLedger as  Select tblmastaccount.AcName, " _
            & " tblvoucher.vdt as vdt,tblvoucher.vtype as Vtype,tblvoucher.vno as vno,Vctrno, tblAccO.acname as " _
            & " AcOName,decode(sign(vamt),-1,abs(vamt),0) AS Debit,  decode(sign(vamt),1,vamt,0) " _
            & " AS Credit, tblvoucher.vbillno, tblvoucher.vbilldt, vnar1, vnar2, vnar3,VNar4,tblvoucher.vaccode, " _
            & " AgName,2 as orderno,UserName,tblMastAccount.ACCITY ,(Select Nar1 From TmpRptTbl where " _
            & " TmpRptTbl.UserName=tmpGenTbl.Username and TmpRptTbl.Vtype=TblVoucher.Vtype and " _
            & " TmpRptTbl.Vno=TblVoucher.Vno and TmpRptTbl.VYear=TblVoucher.VYear)  as RecVsSaleNarr From tblvoucher, " _
            & " tblMastAccount , TblMastGroup, tblMastAccount " _
            & " tblAccO,tmpGenTbl  Where tblMastAccount.AGCODE= " _
            & " TblMastGroup.AGCODE and tblvoucher.vacCode = tblMastAccount.AcCode " _
            & " and tblvoucher.VAcOCode = tblAccO.AcCode  and tblvoucher.Vdt>=dt1 and tblvoucher.Vdt<=dt2 " _
            & " Union all  Select AcName,null,'0P',0,0,'Opening Balance'," _
            & " decode(sign(sum(vamt)),-1,abs(sum(vamt)),0),  decode(sign(sum(vamt)),1,sum(vamt),0), " _
            & " '',null,'','','','',0,AgName,1,UserName,tblMastAccount.ACCITY,'' From tblvoucher, tblmastaccount,TblMastGroup,tmpGenTbl " _
            & " Where tblMastAccount.AGCODE= TblMastGroup.AGCODE and " _
            & " tblvoucher.vacCode = tblMastAccount.AcCode And tblvoucher.Vdt < dt1 " _
            & " group by AcName,AgName,username,tblMastAccount.ACCITY having sum(vamt)<>0" _
            & " ORDER BY AcName,orderno,vdt,vtype,vno ")
'5> qryAccountBalance '-(dr ,Cr closing for TmpGebTbl On Date (dt2)  & user  )
gCn.Execute ("Create or replace View qryAccountBalance as Select AcName," _
            & " decode(sign(sum(vamt)),-1,abs(sum(vamt)),0) as Debit,  decode(sign(sum(vamt)),1,sum(vamt),0) as Credit, " _
            & " AgName,ACCITY,AcCode,UserName,TopGroup From tblvoucher, tblmastaccount,TblMastGroup,tmpselection " _
            & " Where tblMastAccount.AGCODE= TblMastGroup.AGCODE and " _
            & " tblvoucher.vacCode = tblMastAccount.AcCode And tblvoucher.Vdt >= dt1 and tblvoucher.Vdt <= dt2 " _
            & " group by AcName,AgName,ACCITY,AcCode,username,TopGroup having sum(vamt)<>0" _
            & " ORDER BY AcName")
'6> qryAccGroupWiseLedger '-(dr ,Cr & ,Purchase,Sale,Opening for TmpGebTbl Date Range & user  )
'-(dr ,Cr & ,Purchase)
X1 = "Create or replace View qryAccGroupWiseLedger as " _
& " Select tblmastaccount.AcName, tblvoucher.vdt as vdt,tblvoucher.vtype as Vtype,tblvoucher.vno as vno,Vctrno, " _
& " tblAccO.acname as  AcOName,decode(sign(vamt),-1,abs(vamt),0) AS Debit,decode(sign(vamt),1,vamt,0)  AS Credit, " _
& " tblvoucher.vbillno, tblvoucher.vbilldt, vnar1, vnar2, vnar3, vnar4,tblvoucher.vaccode,  AgName,tblmastaccount.Agcode,2 as orderno,UserName " _
& " From tblvoucher, tblMastAccount,TblMastGroup, tblMastAccount  tblAccO,tmpGenTbl  Where " _
& " tblMastAccount.AGCODE= TblMastGroup.AGCODE and tblvoucher.vacCode = tblMastAccount.AcCode And tblvoucher.VAcOCode = tblAccO.AcCode " _
& " and tblvoucher.Vdt>=dt1 and  tblvoucher.Vdt<=dt2 and ((tblvoucher.vtype not in ('SY','ST','SO','SD','SM','PY','PT','PO','PI') " _
& " and tblMastAccount.ACISLEDGERTOT =1 ) or (tblMastAccount.ACISLEDGERTOT=0) ) " _
& " union all Select AcName,last_day(tblvoucher.vdt) as Vdt,'' as Vtype,0 as vno,0,'Purchase of Month ' || to_char(last_day(tblvoucher.Vdt),'Month') , " _
& " decode(sign(sum(vamt)),-1,abs(sum(vamt)),0),decode(sign(sum(vamt)),1,sum(vamt),0)  , '', null,'','','','',tblvoucher.vaccode,AgName,tblmastaccount.Agcode,2 as orderno, " _
& " UserName  From tblvoucher, tblMastAccount,TblMastGroup,tmpGenTbl " _
& " Where tblMastAccount.AgCode = TblMastGroup.AgCode And tblvoucher.vacCode = tblMastAccount.AcCode  and tblvoucher.Vdt>=dt1 and tblvoucher.Vdt<=dt2 and " _
 & " tblvoucher.vtype in ('PY','PT','PO','PI') and ACISLEDGERTOT =1 group by " _
& " AcName,tblvoucher.vacCode, AgName,tblmastaccount.Agcode, UserName, last_day(tblvoucher.Vdt) " _
& " union all "
'Sale,Opening
X2 = " Select AcName,last_day(tblvoucher.vdt) as Vdt,'' as Vtype,0 as vno,0,'Sales of Month ' || to_char(last_day(tblvoucher.Vdt),'Month') , " _
& " decode(sign(sum(vamt)),-1,abs(sum(vamt)),0),  decode(sign(sum(vamt)),1,sum(vamt),0)  , " _
 & " '', null,'','','','',tblvoucher.vaccode,AgName,tblmastaccount.Agcode,2 as orderno,UserName  From tblvoucher, " _
& " tblMastAccount,TblMastGroup,tmpGenTbl   Where tblMastAccount.AGCODE= TblMastGroup.AGCODE " _
& " and tblvoucher.vacCode = tblMastAccount.AcCode  and tblvoucher.Vdt>=dt1 and " _
& " tblvoucher.Vdt<=dt2 and tblvoucher.vtype in ('SY','ST','SO','SD','SM') and ACISLEDGERTOT =1 " _
& " group by AcName,tblvoucher.vaccode,AgName,tblmastaccount.Agcode,UserName,last_day(tblvoucher.vdt)  " _
 & " Union all  Select AcName,null,'0P',0,0,'Opening Balance', decode(sign(sum(vamt)),-1,abs(sum(vamt)),0), " _
& " decode(sign(sum(vamt)),1,sum(vamt),0),  '',null,'','','','',0,AgName,tblmastaccount.Agcode,1,UserName " _
& " From tblvoucher, tblMastAccount, TblMastGroup, tmpGenTbl " _
& " Where tblMastAccount.AGCODE= TblMastGroup.AGCODE and " _
 & " tblvoucher.vacCode = tblMastAccount.AcCode And tblvoucher.Vdt < dt1 " _
 & " group by AcName,AgName,tblmastaccount.Agcode,username  ORDER BY AcName,orderno,vdt,vtype,vno "
gCn.Execute (X1 & X2)
'4> qryPartyLedgerSingle
gCn.Execute ("Create or replace View qryPartyLedgerSingle as Select vaccode, vdt, vtype, vno,Vctrno, acname, decode(sign(vamt),-1,abs(vamt),0) AS Debit, " _
           & "decode(sign(vamt),1,vamt,0) AS Credit, vbillno, vbilldt, vnar1, vnar2, vnar3, vnar4 " _
           & " From tblvoucher, tblmastaccount " _
           & " Where tblvoucher.vacocode = tblmastaccount.AcCode " _
           & " ORDER BY vaccode, vdt, vtype, vno ")
'7> qryStatutoryForms
gCn.Execute ("Create or replace View qryStatutoryForms as SELECT tblMastAccount.AcName, tblMastAccount.AcBST, tblMastNarration.Narration, tblMastAccount_1.AcName as MillName, Count(tblMastAccount_1.AcName) AS CountOfAcName, Sum(tblAddLess.AdBillAmt) AS SumOfAdBillAmt " _
           & "FROM tblMastNarration INNER JOIN (tblMastAccount INNER JOIN (tblAddLess INNER JOIN tblMastAccount  tblMastAccount_1 ON tblAddLess.AdMillCode = tblMastAccount_1.AcCode) ON tblMastAccount.AcCode = tblAddLess.AdAcCrCode) ON tblMastNarration.NarrCode = tblAddLess.AdTaxCode " _
           & "GROUP BY tblMastAccount.AcName, tblMastAccount.AcBST, tblMastNarration.Narration, tblMastAccount_1.AcName " _
           & "ORDER BY tblMastAccount.AcName, tblMastAccount.AcBST, tblMastNarration.Narration, tblMastAccount_1.AcName ")
'8> qrytblVoucherSr1
gCn.Execute ("Create or replace View qrytblVoucherSr1 as SELECT * From tblvoucher WHERE VCtrNo=1")

'9> qryTrbal
gCn.Execute ("Create or replace View qryTrbal as SELECT acname, vaccode, tblmastaccount.agcode, topgroup, tmpselection.Username, decode(sign(sum(vamt)),-1,(sum(vamt)*-1),0) AS Debit," _
& " decode(sign(sum(vamt)),1,sum(vamt),0) AS Credit, tblmastgroup.AgName From tblvoucher, tblmastaccount, tblmastgroup, tmpselection " _
& " Where tblvoucher.vaccode = tblmastaccount.AcCode And tblmastgroup.AgCode = tblmastaccount.AgCode And tblvoucher.Vdt <= tmpselection.dt2 " _
& " GROUP BY acname, vaccode, tblmastaccount.agcode, topgroup, tmpselection.Username, AgName " _
& " ORDER BY acname")
'10> qryTrbalDetail1,qryTrbalDetail
gCn.Execute ("Create or replace View qryTrbalDetail1 as SELECT tblvoucher.vaccode, sum(decode(tblvoucher.vtype,'0P',vamt,0)) AS OpBal, " _
         & "abs(sum(decode(tblvoucher.vtype,'0P',0,decode(sign(vamt),-1,vamt,0)))) AS Debit, " _
         & "sum(decode(tblvoucher.vtype,'0P',0,decode(sign(vamt),1,vamt,0))) AS Credit," _
         & "UserName From tblvoucher, tmpSelection " _
         & "Where tblvoucher.Vdt<dt1 GROUP BY  tblvoucher.vaccode,Username having sum(vamt) <> 0 " _
         & " UNION ALL SELECT  tblvoucher.vaccode, sum(decode(tblvoucher.vtype,'0P',vamt,0)) as OpBal, " _
         & "abs(sum(decode(tblvoucher.vtype,'0P',0,decode(sign(vamt),-1,vamt,0)))) AS Debit, " _
         & "sum(decode(tblvoucher.vtype,'0P',0,decode(sign(vamt),1,vamt,0))) AS Credit, " _
         & "UserName From tblvoucher, tmpSelection " _
         & "Where tblvoucher.Vdt>=dt1 and tblvoucher.Vdt<=dt2 GROUP BY tblvoucher.vaccode,Username having sum(vamt) <> 0 ")
      '   (qryTrbalDetail
      '[ select 1) 2) Asset-> AgCode for Dr,Agcode1 for Cr
      ' select 3) 4) Liablility -> AgCode For Cr,Agcode1 For Dr  ]
      ' select 5) Revenue-> AgCode ]
gCn.Execute ("Create or replace View qryTrbalDetail as SELECT AcCode,AcName,tblmastgroup.AgCode, abs(sum(decode(sign(OpBal),-1,OpBal,0))) AS OpBal_Dr,sum(decode(sign(OpBal),1,OpBal,0)) AS OpBal_Cr, " _
         & "sum(Debit) as Debit,sum(Credit) as Credit, decode(sign(sum(OpBal+Credit-Debit)), -1,sum(OpBal+Credit-Debit)*-1,0) as ClBal_Dr, decode(sign(sum(OpBal+Credit-Debit)), 1,sum(OpBal+Credit-Debit),0) as ClBal_Cr, " _
         & "UserName,' ' as Schedule,'A' as AcGrpType, decode(tblmastGroup.Agcode,90018,1,0) as IsOpStkAc From qryTrbalDetail1, tblmastaccount,tblmastgroup " _
         & "Where qryTrbalDetail1.vaccode = tblmastaccount.AcCode and  tblmastaccount.AgCode=tblmastgroup.AgCode " _
         & " GROUP BY AcCode,AcName,tblmastgroup.AgCode,Username,TopGroup having TopGroup=90001 and sum(OpBal+Credit-Debit)<0 " _
         & "  union all SELECT AcCode,AcName,AgCode1, abs(sum(decode(sign(OpBal),-1,OpBal,0))),sum(decode(sign(OpBal),1,OpBal,0)) , " _
         & "sum(Debit),sum(Credit) , decode(sign(sum(OpBal+Credit-Debit)), -1,sum(OpBal+Credit-Debit)*-1,0) , decode(sign(sum(OpBal+Credit-Debit)), 1,sum(OpBal+Credit-Debit),0) , " _
         & "UserName,' ' ,'A',decode(Agcode1,90018,1,0)  From qryTrbalDetail1, tblmastaccount,tblmastgroup Where qryTrbalDetail1.vaccode = tblmastaccount.AcCode and  tblmastaccount.AgCode=tblmastgroup.AgCode " _
         & " GROUP BY AcCode,AcName,AgCode1,Username,TopGroup having TopGroup=90001 and sum(OpBal+Credit-Debit)>0 " _
         & " union all SELECT AcCode,AcName,tblmastgroup.AgCode, abs(sum(decode(sign(OpBal),-1,OpBal,0))),sum(decode(sign(OpBal),1,OpBal,0)) , " _
         & "sum(Debit),sum(Credit) , decode(sign(sum(OpBal+Credit-Debit)), -1,sum(OpBal+Credit-Debit)*-1,0) , decode(sign(sum(OpBal+Credit-Debit)), 1,sum(OpBal+Credit-Debit),0) , " _
         & "UserName,' ' ,'A',decode(tblmastGroup.Agcode,90018,1,0)  From qryTrbalDetail1, tblmastaccount,tblmastgroup  Where qryTrbalDetail1.vaccode = tblmastaccount.AcCode and  tblmastaccount.AgCode=tblmastgroup.AgCode " _
         & " GROUP BY AcCode,AcName,tblmastgroup.AgCode,Username,TopGroup having TopGroup=90002 and sum(OpBal+Credit-Debit)>0 " _
          & "  union all SELECT AcCode,AcName,AgCode1, abs(sum(decode(sign(OpBal),-1,OpBal,0))),sum(decode(sign(OpBal),1,OpBal,0)) , " _
         & "sum(Debit),sum(Credit) , decode(sign(sum(OpBal+Credit-Debit)), -1,sum(OpBal+Credit-Debit)*-1,0) , decode(sign(sum(OpBal+Credit-Debit)), 1,sum(OpBal+Credit-Debit),0) , " _
         & "UserName,' ' ,'A',decode(Agcode1,90018,1,0)  From qryTrbalDetail1, tblmastaccount,tblmastgroup Where qryTrbalDetail1.vaccode = tblmastaccount.AcCode and  tblmastaccount.AgCode=tblmastgroup.AgCode " _
         & " GROUP BY AcCode,AcName,AgCode1,Username,TopGroup having TopGroup=90002 and sum(OpBal+Credit-Debit)<0 " _
         & "union all  SELECT AcCode,AcName,tblmastgroup.AgCode, abs(sum(decode(sign(OpBal),-1,OpBal,0))) AS OpBal_Dr,sum(decode(sign(OpBal),1,OpBal,0)) AS OpBal_Cr, " _
         & "sum(Debit) as Debit,sum(Credit) as Credit, decode(sign(sum(OpBal+Credit-Debit)), -1,sum(OpBal+Credit-Debit)*-1,0) as ClBal_Dr, decode(sign(sum(OpBal+Credit-Debit)), 1,sum(OpBal+Credit-Debit),0) as ClBal_Cr, " _
         & "UserName,' ' as Schedule,'A' as AcGrpType,decode(tblmastGroup.Agcode,90018,1,0) From qryTrbalDetail1, tblmastaccount,tblmastgroup " _
         & "Where qryTrbalDetail1.vaccode = tblmastaccount.AcCode and  tblmastaccount.AgCode=tblmastgroup.AgCode " _
         & " GROUP BY AcCode,AcName,tblmastgroup.AgCode,Username,TopGroup having TopGroup=90003 " _
         & " ORDER BY acname")
         '& " UNION ALL SELECT AgCode,AgName,Gpcode, 0, 0, 0, 0,0,0, " _
         '& "UserName,Schedule,'G' From tblmastgroup, tmpSelection " _
'gCn.Execute ("Create or replace View qryTrbalDetail as SELECT AcCode,AcName,AgCode, abs(sum(decode(sign(OpBal),-1,OpBal,0))) AS OpBal_Dr,sum(decode(sign(OpBal),1,OpBal,0)) AS OpBal_Cr, " _
'         & "sum(Debit) as Debit,sum(Credit) as Credit, decode(sign(sum(OpBal+Credit-Debit)), -1,sum(OpBal+Credit-Debit)*-1,0) as ClBal_Dr, decode(sign(sum(OpBal+Credit-Debit)), 1,sum(OpBal+Credit-Debit),0) as ClBal_Cr, " _
'         & "UserName,' ' as Schedule,'A' as AcGrpType From qryTrbalDetail1, tblmastaccount " _
'         & "Where qryTrbalDetail1.vaccode = tblmastaccount.AcCode  GROUP BY AcCode,AcName,AgCode,Username " _
'         & " UNION ALL SELECT AgCode,AgName,Gpcode, 0, 0, 0, 0,0,0, " _
'         & "UserName,Schedule,'G' From tblmastgroup, tmpGentbl " _
'         & " ORDER BY acname")
'11> qryGrpTrbalDetail  (A/c Group,Sub Group 1 to 5 Level Hierarchy )
gCn.Execute "create or replace view qryGrpTrbalDetail as " _
& " SELECT AgName as H1,AcName as H2,'' as H3,'' as H4,'' as H5 ,'' as H6,OpBal_Dr as OpDr,OpBal_Cr as OpCr,Debit as Dr,Credit as Cr,ClBal_Dr as ClDr, ClBal_Cr as ClCr,'' as GrpAcctype,2 as HeadNo,UserName " _
& " ,'G' as H1Type,'A' as H2Type,'-' as H3Type,'-' as H4Type,'-' as H5Type,'-' as H6Type,TblMastGroup.Schedule as Schedule,1 as ScheduleHNo,TopGroup,IsOpStkAc  From tblMastGroup, qryTrbalDetail WHERE TblMastGroup.Agcode = qryTrbalDetail.Agcode and " _
& " TblMastGroup.AgCode=TblMastGroup.TopGroup  and AcGrpType='A' " _
& " Union All SELECT TblGrp_T.AgName AS H1,tblMastGroup.AgName AS H2, AcName as h3,'','','',OpBal_Dr,OpBal_Cr,Debit,Credit,ClBal_Dr,ClBal_Cr,'',3,username " _
& "  ,'G' as H1Type,'G' as H2Type,'A' as H3Type,'-','-','-',decode(TblGrp_T.Schedule,Null,tblMastgroup.Schedule,TblGrp_T.Schedule)  as Schedule,decode(TblGrp_T.Schedule,Null,2,1),tblMastGroup.TopGroup,IsOpStkAc FROM TblMastGroup ,TblMastGroup TblGrp_T,qryTrbalDetail where TblMastGroup.AgCode = qryTrbalDetail.AgCode And TblMastGroup.Gpcode = TblGrp_T.AgCode " _
& " AND TblGrp_T.AgCode=TblGrp_T.TopGroup  and AcGrpType='A' and TblMastGroup.AgCode<>TblMastGroup.topGroup " _
 & " Union All SELECT TblGrp_T1.Agname AS H1, TblGrp_T2.AgName AS H2,tblMastGroup.AgName as " _
& " H3,acName,'','',OpBal_Dr,OpBal_Cr,Debit,Credit,ClBal_Dr,ClBal_Cr,'',4,UserName " _
& " ,'G','G','G','A','-','-',decode(TblGrp_T1.Schedule,Null,decode(TblGrp_T2.Schedule,Null,tblMastGroup.Schedule,TblGrp_T2.Schedule),TblGrp_T1.Schedule)  as Schedule,decode(TblGrp_T1.Schedule,Null,decode(TblGrp_T2.Schedule,Null,3,2),1),tblMastGroup.TopGroup,IsOpStkAc FROM TblMastGroup,TblMastGroup TblGrp_T1 , TblMastGroup TblGrp_T2 ,qryTrbalDetail " _
& " Where TblGrp_T1.AgCode = TblGrp_T2.Gpcode And qryTrbalDetail.AgCode = TblMastGroup.AgCode " _
& " and TblGrp_T2.AgCode = tblMastGroup.GpCode and TblGrp_T1.Gpcode=TblGrp_T1.AgCode  and AcGrpType='A' and TblGrp_T2.AgCode<>TblGrp_T2.topGroup " _
& " Union All SELECT TblGrp_T1.Agname AS H1, TblGrp_T2.AgName AS H2,TblGrp_T3.AgName AS H3,tblMastGroup.AgName as " _
& " H4,AcName,'',OpBal_Dr,OpBal_Cr,Debit,Credit,ClBal_Dr,ClBal_Cr,'',5,UserName" _
& " ,'G','G','G','G','A','-',decode(TblGrp_T1.Schedule,Null,decode(TblGrp_T2.Schedule,Null,decode(TblGrp_T3.Schedule,Null,tblMastGroup.Schedule,TblGrp_T3.Schedule),TblGrp_T2.Schedule),TblGrp_T1.Schedule)  as Schedule,decode(TblGrp_T1.Schedule,Null,decode(TblGrp_T2.Schedule,Null,decode(TblGrp_T3.Schedule,Null,4,3),2),1),tblMastGroup.TopGroup,IsOpStkAc FROM TblMastGroup,TblMastGroup TblGrp_T1,TblMastGroup TblGrp_T2 ,TblMastGroup TblGrp_T3 ,qryTrbalDetail " _
& " where TblGrp_T1.AgCode= TblGrp_T2.GpCode and TblGrp_T2.AgCode= TblGrp_T3.GpCode and " _
& " qryTrbalDetail.AgCode=TblMastGroup.Agcode and TblGrp_T3.AgCode = TblMastGroup.Gpcode And TblGrp_T1.Gpcode = TblGrp_T1.AgCode " _
& " and TblGrp_T2.AgCode<>TblGrp_T2.topGroup  and AcGrpType='A'" _
& " Union All select TblGrp_T1.Agname AS H1, TblGrp_T2.AgName AS H2,TblGrp_T3.AgName AS H3, " _
& " TblGrp_T4.AgName AS H4,tblMastGroup.AgName as H5,AcName,OpBal_Dr,OpBal_Cr,Debit,Credit,ClBal_Dr,ClBal_Cr,'',6,UserName " _
& " ,'G','G','G','G','G','A',decode(TblGrp_T1.Schedule,Null,decode(TblGrp_T2.Schedule,Null,decode(TblGrp_T3.Schedule,Null,decode(TblGrp_T4.Schedule,Null,tblMastGroup.Schedule,TblGrp_T4.Schedule ),TblGrp_T3.Schedule),TblGrp_T2.Schedule),TblGrp_T1.Schedule)  as Schedule,decode(TblGrp_T1.Schedule,Null,decode(TblGrp_T2.Schedule,Null,decode(TblGrp_T3.Schedule,Null,decode(TblGrp_T4.Schedule,Null,5,4 ),3),2),1),tblMastGroup.TopGroup,IsOpStkAc FROM TblMastGroup,TblMastGroup TblGrp_T1,TblMastGroup TblGrp_T2 ,TblMastGroup TblGrp_T3 ," _
& " TblMastGroup TblGrp_T4, qryTrbalDetail Where TblGrp_T1.AgCode = TblGrp_T2.Gpcode And TblGrp_T2.AgCode = TblGrp_T3.Gpcode" _
& " and TblGrp_T3.AgCode= TblGrp_T4.GpCode and qryTrbalDetail.AgCode=TblMastGroup.Agcode and TblGrp_T4.AgCode = TblMastGroup.Gpcode And " _
& " TblGrp_T1.Gpcode = TblGrp_T1.AgCode  and TblGrp_T2.AgCode<>TblGrp_T2.topGroup  and AcGrpType='A' " _
& " ORDER BY 1, 2, 3, 4, 5 ,6"
'11> qryPLCalculation
gCn.Execute "create or replace view qryPLCalculation as " _
& " SELECT tmpselection.Username, sum(vamt) as PL FROM tblvoucher, tblmastaccount,tblmastgroup, tmpselection " _
& " Where tblvoucher.vaccode = tblmastaccount.AcCode And tblmastgroup.AgCode = tblmastaccount.AgCode And tblmastgroup.TopGroup = 90003 And tblvoucher.Vdt <= tmpselection.dt2 " _
& " GROUP BY tmpselection.Username " _
& " UNION ALL SELECT tmpselection.Username, sum(stamount) as PL FROM tblclosingstock, tmpselection " _
& " Where tblclosingstock.stdate = tmpselection.dt2 " _
& " GROUP BY tmpselection.Username " _
& " UNION ALL SELECT tmpselection.Username, sum(vamt) as PL FROM tblvoucher, tmpselection,tblmastaccount " _
& " WHERE vtype = '0P' and tblvoucher.vaccode = tblmastaccount.accode and tblmastaccount.agcode = 90018 " _
& " group by username "
'11> qryPL1
gCn.Execute "create or replace view qryPL1 as " _
& " SELECT 1 as Gp,1 as Sr,Acname,vAccode,90018 as Agcode,90003 as topgroup,tmpselection.Username,Debit,Credit,'Opening Stock' as AgName,0 as GrDebit,0 as GrCredit  FROM qrytrbal,tmpselection " _
& " Where qrytrbal.UserName = tmpselection.UserName And (qrytrbal.AgCode = 90018 Or qrytrbal.vaccode = 90018) " _
 & " Union All SELECT 1 as Gp,2 as Sr,qrytrbal.*,0,0  FROM qrytrbal,  tmpselection,tblmastgroup a " _
& " Where qrytrbal.UserName = tmpselection.UserName And qrytrbal.TopGroup = 90003 And qrytrbal.AgCode = a.AgCode And a.gpcode = 90021 " _
& " Union All SELECT  1 as Gp,2 as Sr,qrytrbal.* ,0,0 FROM qrytrbal,  tmpselection,tblmastgroup a,tblmastgroup b " _
& " Where qrytrbal.UserName = tmpselection.UserName And qrytrbal.TopGroup = 90003 And qrytrbal.AgCode = a.AgCode And a.gpcode = b.AgCode And b.gpcode = 90021 " _
& " Union All SELECT  1 as Gp,2 as Sr,qrytrbal.*,0,0  FROM qrytrbal,  tmpselection,tblmastgroup a,tblmastgroup b,tblmastgroup c " _
& " Where qrytrbal.UserName = tmpselection.UserName And qrytrbal.TopGroup = 90003 And qrytrbal.AgCode = a.AgCode And a.gpcode = b.AgCode And b.gpcode = c.AgCode And c.gpcode = 90021 " _
& " Union All SELECT 1 as Gp,3 as Sr,Acname,Accode,90018 as Agcode,90003 as topgroup,tmpselection.Username,0 as Debit,stamount as Credit,'Closing Stock'  ,0,0 FROM tblmastaccount,  tmpselection,tblclosingstock " _
& " Where tblmastaccount.AcCode = tblclosingstock.staccode And tmpselection.dt2 = tblclosingstock.stdate " _
& " Union All SELECT 2 as Gp,2 as Sr,qrytrbal.*,0,0  FROM qrytrbal,  tmpselection,tblmastgroup a " _
& " Where qrytrbal.UserName = tmpselection.UserName And qrytrbal.TopGroup = 90003 And qrytrbal.AgCode = a.AgCode And a.gpcode = 90022 " _
& " Union All SELECT  2 as Gp,2 as Sr,qrytrbal.*,0,0  FROM qrytrbal,  tmpselection,tblmastgroup a,tblmastgroup b " _
& " Where qrytrbal.UserName = tmpselection.UserName And qrytrbal.TopGroup = 90003 And qrytrbal.AgCode = a.AgCode And a.gpcode = b.AgCode And b.gpcode = 90022 " _
& " UNION ALL SELECT  2 as Gp,2 as Sr,qrytrbal.*,0,0  FROM qrytrbal,  tmpselection,tblmastgroup a,tblmastgroup b,tblmastgroup c " _
& " Where qrytrbal.UserName = tmpselection.UserName And qrytrbal.TopGroup = 90003 And qrytrbal.AgCode = a.AgCode And a.gpcode = b.AgCode And b.gpcode = c.AgCode And c.gpcode = 90022" _
& " ORDER BY gp, sr, 9, 2 "
'12> qryPL2_1
gCn.Execute "create or replace view qryPL2_1 as " _
& " select gp,sr,'     ' ||  qrypl1.Acname as AcName,vaccode,Agcode,topGroup,UserName,Debit,Credit,Agname,GrDebit,GrCredit,1 as GrNo  from qrypl1 where gp = 1 " _
& " Union All " _
& " select qrypl1.gp ,sr,AgName as AcName ,0,agcode,9003,username,0,0,AgName,sum(debit),sum(credit),0 as GrNo from qrypl1 where gp = 1 group by gp,sr,agcode,username,agname " _
& " UNION ALL select 1 as gp,4 as sr,decode(sign(sum(credit-debit)) ,1,'Gross Profit','Gross Loss'),0 ,0 ,90003 ,username,decode(sign(sum(credit-debit)),1,sum(credit-debit),0) ,decode(sign(sum(credit-debit)),-1,-sum(credit-debit),0) ,decode(sign(sum(credit-debit)),1,'Gross Profit','Gross Loss'),decode(sign(sum(credit-debit)),1,sum(credit-debit),0),decode(sign(sum(credit-debit)),-1,-sum(credit-debit),0),1 as GrNo from qrypl1 where gp = 1  group by username " _
& " ORDER BY gp, sr, AgName, GrNo, AcName "
'12> qryPL2_2
gCn.Execute "create or replace view qryPL2_2 as " _
& " select gp,sr,'     ' ||  qrypl1.Acname as AcName,vaccode,Agcode,topGroup,UserName,Debit,Credit,Agname,GrDebit,GrCredit,1 as GrNo  from qrypl1 where gp = 2 " _
& " Union All " _
& " select qrypl1.gp ,sr,AgName,0,agcode,9003,username,0,0,AgName,sum(debit),sum(credit),0 as GrNo from qrypl1 where gp = 2 group by gp,sr,agcode,username,agname " _
& " UNION ALL select 2 as gp,1 as sr,decode(sign(sum(credit-debit)),1,'Gross Profit','Gross Loss') as Acname,0 as vAccode,0 as Agcode,90003 as Topgroup,username,decode(sign(sum(credit-debit)),1,0,-sum(credit-debit)) as Dr,decode(sign(sum(credit-debit)),-1,0,sum(credit-debit)) as Cr,decode(sign(sum(credit-debit)),1,'Gross Profit','Gross Loss') as Agname,decode(sign(sum(credit-debit)),1,0,-sum(credit-debit)),decode(sign(sum(credit-debit)),-1,0,sum(credit-debit)),1 as GrNo from qrypl1 where gp = 1  group by username having nvl(sum(credit-debit),0)<>0" _
& " UNION ALL select 2 as gp,3 as sr,decode(sign(sum(credit-debit)),1,'Net Profit','Net Loss') as Acname,0 as vAccode,0  as Agcode,90003 as Topgroup,username,decode(sign(sum(credit-debit)),1,sum(credit-debit),0) as Dr,decode(sign(sum(credit-debit)),-1,-sum(credit-debit),0) as Cr,decode(sign(sum(credit-debit)),1,'Net Profit','Net Loss') as Agname,decode(sign(sum(credit-debit)),1,sum(credit-debit),0),decode(sign(sum(credit-debit)),-1,-sum(credit-debit),0),1  as GrNo from qrypl1  group by username having nvl(sum(credit-debit),0)<>0" _
& " ORDER BY gp, sr, AgName, GrNo, AcName"

'gCn.Execute ("Create or replace View qryTrbalDetail as SELECT AcName,AgCode, abs(sum(decode(tblvoucher.vtype,'0P',decode(sign(vamt),-1,Vamt,0),0))) AS OpBal_Dr, " _
'         & "sum(decode(tblvoucher.vtype,'0P',decode(sign(vamt),1,vamt,0),0)) AS OpBal_Cr, abs(sum(decode(tblvoucher.vtype,'0P',0,decode(sign(vamt),-1,vamt,0)))) AS Debit, " _
'         & "sum(decode(tblvoucher.vtype,'0P',0,decode(sign(vamt),1,vamt,0))) AS Credit, decode(sign(sum(vamt)),-1,abs(sum(vamt)),0) AS ClBal_Dr, " _
'         & "decode(sign(sum(vamt)),1, sum(vamt), 0) As ClBal_Cr,UserName From tblvoucher, tblmastaccount,tmpGentbl " _
'         & "Where tblvoucher.vaccode = tblmastaccount.AcCode and tblvoucher.Vdt<=dt2 GROUP BY acname,AgCode,tblvoucher.vaccode,Username having sum(vamt)<>0 " _
'         & " UNION ALL SELECT AcName,AgCode, abs(sum(decode(tblvoucher.vtype,'0P',decode(sign(vamt),-1,Vamt,0),0))) AS OpBal_Dr, " _
'         & "sum(decode(tblvoucher.vtype,'0P',decode(sign(vamt),1,vamt,0),0)) AS OpBal_Cr, abs(sum(decode(tblvoucher.vtype,'0P',0,decode(sign(vamt),-1,vamt,0)))) AS Debit, " _
'         & "sum(decode(tblvoucher.vtype,'0P',0,decode(sign(vamt),1,vamt,0))) AS Credit, decode(sign(sum(vamt)),-1,abs(sum(vamt)),0) AS ClBal_Dr, " _
'         & "decode(sign(sum(vamt)),1, sum(vamt), 0) As ClBal_Cr,UserName From tblvoucher, tblmastaccount,tmpGentbl " _
'         & "Where tblvoucher.vaccode = tblmastaccount.AcCode and tblvoucher.Vdt<=dt2 GROUP BY acname,AgCode,tblvoucher.vaccode,Username having sum(vamt)<>0 " _
'         & " ORDER BY acname")
'14> QryRegUnadjustReceipt
gCn.Execute ("Create or replace View QryRegUnadjustReceipt as " _
& " select vno,vdt,vtype,decode(Vtype,'CR',tblmastAccount.Acname,'BR',tblmastAccount.Acname,tblMAccount_1.Acname) as " _
& " party,decode(Vtype,'CR',tblMAccount_1.Acname,'BR',tblMAccount_1.Acname,tblmastAccount.Acname) as CrAccount," _
& " VAMT,vamt as UnadjustAmt ,VNAR1,VNAR2,VNAR3," _
& " VNAR4,VBILLNO, VBILLDT, VRECONDT from " _
& " tblVoucher , tblMastAccount, tblMastAccount " _
& " tblMAccount_1 Where " _
& " tblVoucher.VAcCode = tblMastAccount.AcCode" _
& " and  tblVoucher.VAcOCode=tblMAccount_1.AcCode and " _
& " Vtype in ('CR','BR','BP','CP') and VCTRNO=1 and (tblMastAccount.AgCode in (90017,90015,90029) " _
& " or  tblMAccount_1.AgCode in (90017,90015,90029) ) and vno not in (select RecVno from tblRecVsSale where " _
& " RecVtype=Vtype and recVno=Vno and recVyear=Vyear) " _
& " Union All " _
& " select vno,vdt,vtype,decode(Vtype,'CR',tblmastAccount.Acname,'BR',tblmastAccount.Acname,tblMAccount_1.Acname) as " _
& " party,decode(Vtype,'CR',tblMAccount_1.Acname,'BR',tblMAccount_1.Acname,tblmastAccount.Acname) as CrAccount, " _
& " VAMT,vamt - (select Sum(AdjAmt+AdjDBNtAmt+Intreceived) from tblRecVsSale " _
& " where RecVtype=Vtype and recVno=Vno and recVyear=Vyear) " _
 & " as UnadjustAmt ,VNAR1,VNAR2,VNAR3, " _
& " VNAR4,VBILLNO, VBILLDT, VRECONDT from " _
 & " tblVoucher , tblMastAccount, tblMastAccount  tblMAccount_1 Where " _
 & " tblVoucher.VAcCode = tblMastAccount.AcCode " _
& " and  tblVoucher.VAcOCode=tblMAccount_1.AcCode and " _
& " Vtype in ('CR','BR','BP','CP') and VCTRNO=1 and (tblMastAccount.AgCode in (90017,90015,90029) " _
& " or tblMAccount_1.AgCode in (90017,90015,90029) ) and vno in (select RecVno from tblRecVsSale where" _
 & " RecVtype=Vtype and recVno=Vno and recVyear=Vyear)")
 '15> QryStkItemLedger
 gCn.Execute ("Create or replace view QryStkItemLedger as " _
& "  SELECT tblAddLess.VNo, to_char(tblAddLess.Vno) as BillNo, tblAddLess.VType,tblAddLess.Vdt as BillDt,ItName as Item, " _
& " tblMastAccount_Mill.AcName as Mill,tblMastAccount.AcName  as Particulars,VSubBag as InBag,VSubWt as InWt, " _
& " 0 as OutBag,0 as OutWt,0 as orderno,ItCode,ItName || ' - ' ||  tblMastAccount_Mill.AcName as ItemMill " _
 & " FROM tblAddLess,tblVouSub,tblMastItem,tblMastAccount,tblMastAccount tblMastAccount_Mill where " _
 & " tblAddLess.VYear = tblVouSub.VYear And tblAddLess.Vtype = tblVouSub.Vtype " _
 & " AND tblAddLess.VNo = tblVouSub.VNo and VSubItCode = ItCode and  AdAcCrCode=tblMastAccount.AcCode and ItMillCode=tblMastAccount_Mill.AcCode" _
 & " and tblAddLess.vtype in ('PY','PI','PT','RY','RO')" _
 & " Union All SELECT tblAddLess.VNo,decode(AdInvNo,Null,to_char(tblAddLess.vno),adinvno), tblAddLess.VType,tblAddLess.VDt,ItName,tblMastAccount_Mill.AcName,tblMastAccount.AcName ,0,0,VSubBag, " _
 & " VSubWt,1 ,ItCode,ItName || ' - ' ||  tblMastAccount_Mill.AcName FROM tblAddLess,tblVouSub,tblMastItem,tblMastAccount,tblMastAccount tblMastAccount_Mill where " _
 & " tblAddLess.VYear = tblVouSub.VYear And tblAddLess.Vtype = tblVouSub.Vtype AND tblAddLess.VNo = tblVouSub.VNo and VSubItCode = ItCode and AdAcDrCode = tblMastAccount.AcCode and ItMillCode = tblMastAccount_Mill.AcCode " _
 & " and tblAddLess.vtype in ('SY','SD','ST','SO','VY','VI')  ORDER BY billdt, orderno, vno ")
 '16> qryAccLedgerMillBill '-(dr ,Cr)
gCn.Execute ("Create or replace View qryAccLedgerMillBill as  Select tblmastaccount.AcName, " _
            & " vdt,vtype,vno,tblAccO.acname as " _
            & " AcOName,decode(sign(MRecAmt),1,0,abs(MRecAmt)) AS Debit,decode(sign(MRecAmt),1,MRecAmt,0)" _
            & " AS Credit,to_char(vno) as BillNo,MRecNar1 as Nar1,MRecNar2 as Nar2, MRecNar3 as Nar3,MRecNar4 as Nar4,MRecAcCode as AcCode, " _
            & " AgName,2 as orderno,tblMastAccount.ACCITY from tblMillRecPay, " _
            & " tblMastAccount , TblMastGroup, tblMastAccount tblAccO Where Vtype='MR' and tblMastAccount.AGCODE= " _
            & " TblMastGroup.AGCODE and MRecAcCode = tblMastAccount.AcCode  and MRecAcOCode = tblAccO.AcCode" _
            & " union all Select tblmastaccount.AcName, " _
            & " vdt,vtype,vno,tblAccO.acname as AcOName,AdBillAmt AS Debit,0 ,AdInvNo" _
            & " ,'','','','',ADACDRCODE, AgName,1 as orderno,tblMastAccount.ACCITY from tblAddLess, " _
            & " tblMastAccount , TblMastGroup, tblMastAccount tblAccO Where Vtype='SM' and tblMastAccount.AGCODE= " _
            & " TblMastGroup.AGCODE and ADACDRCODE = tblMastAccount.AcCode and ADACCRCODE = tblAccO.AcCode" _
            & " union all Select tblmastaccount.AcName, " _
            & " vdt,vtype,vno,tblAccO.acname as AcOName,0,AdBillAmt,AdInvNo" _
            & " ,'','','','',ADACCRCODE, AgName,1 as orderno,tblMastAccount.ACCITY from tblAddLess, " _
            & " tblMastAccount , TblMastGroup, tblMastAccount tblAccO Where Vtype='SM' and tblMastAccount.AGCODE= " _
            & " TblMastGroup.AGCODE and ADACCRCODE = tblMastAccount.AcCode and ADACDRCODE = tblAccO.AcCode" _
            & " ORDER BY AcName,vdt,Orderno")
 '17> qryInterestLedger '-(dr ,Cr & 0P  for TmpGebTbl Date Range (dt1 - dt2 ),Sale Side Days Amt1 & user  )
 '--- Consider Amt1( sale side Days) for Sales Vouchers only
'-- TmpRptTbl prepare TmpRptTbl For Rec Vs Sale Narration
gCn.Execute ("Create or replace View qryInterestLedger as  Select tblmastaccount.AcName, " _
            & " tblvoucher.vdt as vdt,tblvoucher.vtype as Vtype,tblvoucher.vno as vno,Vctrno, tblAccO.acname as " _
            & " AcOName,decode(sign(vamt),-1,abs(vamt),0) AS Debit,  decode(sign(vamt),1,vamt,0) " _
            & " AS Credit, tblvoucher.vbillno, tblvoucher.vbilldt, vnar1, vnar2, vnar3,VNar4,tblvoucher.vaccode, " _
            & " AgName,2 as orderno,UserName,tblMastAccount.ACCITY ,'' as RecVsSaleNarr,0 as ClDebit,0 as ClCredit From tblvoucher, tblMastAccount , TblMastGroup, tblMastAccount " _
            & " tblAccO,tmpGenTbl  Where tblMastAccount.AGCODE= TblMastGroup.AGCODE and tblvoucher.vacCode = tblMastAccount.AcCode  " _
            & " and tblvoucher.VAcOCode = tblAccO.AcCode  and ((tblvoucher.Vdt>=dt1 and tblvoucher.Vdt<=dt2 and tblvoucher.Vtype not in ('SY','ST','SO','SD') ) " _
            & " or ( (tblvoucher.Vdt + Amt1)>=dt1 and (tblvoucher.Vdt+amt1)<=dt2 and tblvoucher.Vtype in ('SY','ST','SO','SD') ) ) " _
            & " Union all  Select AcName,dt1 as vdt ,'0P' as Vtype,0 as vno ,0,'Opening Balance',decode(sign(sum(vamt)),-1,abs(sum(vamt)),0),  decode(sign(sum(vamt)),1,sum(vamt),0), " _
            & " '',null,'','','','',0,AgName,1 as orderno,UserName,tblMastAccount.ACCITY,'',0,0 From tblvoucher, tblmastaccount,TblMastGroup,tmpGenTbl " _
            & " Where tblMastAccount.AGCODE= TblMastGroup.AGCODE and " _
            & " tblvoucher.vacCode = tblMastAccount.AcCode And ( ( tblvoucher.Vdt < dt1 and tblvoucher.Vtype not in ('SY','ST','SO','SD')) or  ( (tblvoucher.Vdt+amt1) < dt1 and tblvoucher.Vtype in  ('SY','ST','SO','SD')  ) )" _
            & " group by AcName,dt1 ,AgName,username,tblMastAccount.ACCITY having sum(vamt)<>0" _
            & " union all select ACNAME,Null as vdt ,'CL' as Vtype,0 as vno,0,'A/c Closing Balance',0,0,'',null,'','','','',AcCode,AgName,3 as orderno,UserName,AcCity,'',Debit,Credit from qryAccountBalance " _
            & " ORDER BY AcName,orderno,vdt,vtype,vno ")
            '& "union all select acname,null,null,0,0,'A/c Closing Balance',0,0,0,null,'','','','',accode,agname,3 as orderno ,UserName,accity,'',debit,credit from qryAccountBalance " _
 '18> qryBankReconcilLedger '-(dr ,Cr & 0P  for TmpGebTbl Date Range (dt1 - dt2 ) & user  )
gCn.Execute ("Create or replace View qryBankReconcilLedger as  Select tblmastaccount.AcName, " _
            & " tblvoucher.VRECONDT as vdt,tblvoucher.vtype as Vtype,tblvoucher.vno as vno,Vctrno, tblAccO.acname as " _
            & " AcOName,decode(sign(vamt),-1,abs(vamt),0) AS Debit,  decode(sign(vamt),1,vamt,0) " _
            & " AS Credit, tblvoucher.vbillno, tblvoucher.vbilldt, vnar1, vnar2, vnar3,VNar4,tblvoucher.vaccode, " _
            & " AgName,2 as orderno,UserName,tblMastAccount.ACCITY  From tblvoucher, " _
            & " tblMastAccount , TblMastGroup, tblMastAccount " _
            & " tblAccO,tmpGenTbl  Where tblvoucher.vtype in ('BP','BR')  and tblMastAccount.AGCODE= " _
            & " TblMastGroup.AGCODE and tblvoucher.vacCode = tblMastAccount.AcCode " _
            & " and tblvoucher.VAcOCode = tblAccO.AcCode  and VRECONDT>=dt1 and VRECONDT<=dt2 " _
            & " Union all  Select AcName,null,'0P',0,0,'Opening Balance'," _
            & " decode(sign(sum(vamt)),-1,abs(sum(vamt)),0),  decode(sign(sum(vamt)),1,sum(vamt),0), " _
            & " '',null,'','','','',0,AgName,1,UserName,tblMastAccount.ACCITY From tblvoucher, tblmastaccount,TblMastGroup,tmpGenTbl " _
            & " Where tblvoucher.vtype in ('BP','BR') and tblMastAccount.AGCODE= TblMastGroup.AGCODE and " _
            & " tblvoucher.vacCode = tblMastAccount.AcCode And VRECONDT< dt1 " _
            & " group by AcName,AgName,username,tblMastAccount.ACCITY having sum(vamt)<>0" _
            & " ORDER BY AcName,orderno,vdt,vtype,vno ")
 '-------------------------------
 MsgBox "Database Modification is Completed Successfully.", vbOKOnly + vbInformation, App.Title
End Sub
'================  Modify Tables For All Company Current Year
Public Sub ModifyTables()
'On Error GoTo ErrorHandler
Dim AllFirmFlag As Boolean
Dim TmpCn As Connection
Dim RstCompMast As Recordset
If MsgBox("Database Changes will get applied to all Company.Are you Sure?", vbQuestion + vbYesNo + vbDefaultButton1, "Confirm Delete") = vbYes Then
    If gBackEndDB = gBackEndORA Then '-- oracle
        AllFirmFlag = True
        '--- company Recordset
        GProcRstOpen RstCompMast, "Select * from tblCompanyYear,tblMastCompany  where compCode=cCode " _
        & IIf(AllFirmFlag = True, "", " and cName='" & gCName & "'"), "R", gSetupCn
        If RstCompMast.EOF = False Then RstCompMast.MoveFirst
        Do While Not RstCompMast.EOF
            '----Open Temp Connection
            GProcConnectionOpenORA TmpCn, RstCompMast!CompCode + RstCompMast!CompYear
'            '--- Account
'            TmpCn.Execute "alter Table tblMastAccount add (AcFaxNo Varchar2 (40) Default Null)"
'            '-- Voucher
'            TmpCn.Execute "alter Table tblvoucher modify (VNAR1 VARCHAR2(60))"
'            '--- tblConfirmation
'            TmpCn.Execute "alter Table tblConfirmation add (CfRemark2  Varchar2 (50) Default Null)"
'            TmpCn.Execute "alter Table tblConfirmation add (CfRateType Varchar2 (10) Default Null)"
'             TmpCn.Execute "update tblVoucher set vnar1='Bill No ' || vbillno || ' Dtd. ' || VBillDt where " _
'            & " (vnar1 ='' or vnar1 is null ) and vtype in ('PY','PT') "
'            '- Sale Confirmation Qty (Double)
'             TmpCn.Execute " alter table tblConfirmation modify (CFQTY  NUMBER(10,2) )"
'             TmpCn.Execute " alter table tblConfirmation modify ( CfPackingNarr2 Varchar2(60) )"
'             '- Setting (TDSPaidAcCode)
'             TmpCn.Execute " alter table tblMastSetting add ( TDSPaidAcCode  Number (7,0) Default 0)"
'             '--tblRecVsSale (LessTDS,JvNoLessTDS)
'             TmpCn.Execute " alter table tblRecVsSale add ( LessTDS    Number (12,2) Default 0)"
'             TmpCn.Execute " alter table tblRecVsSale add ( JvNoLessTDS     Number (7,0) Default 0)"
            RstCompMast.MoveNext
        Loop
        If RstCompMast.State = adStateOpen Then RstCompMast.Close
        If TmpCn.State = adStateOpen Then TmpCn.Close
        MsgBox "Tables are modified Successfully", vbOKOnly + vbInformation, "Utility"
    Else

        '--- company Recordset
        GProcRstOpen RstCompMast, "Select * from tblCompanyYear,tblMastCompany  where compCode=cCode " _
        & IIf(AllFirmFlag = False, "", " and cName='" & gCName & "'"), "R", gSetupCn
        If RstCompMast.EOF = False Then RstCompMast.MoveFirst

        Do While Not RstCompMast.EOF
            '----Open Temp Connection
            GProcConnectionOpen TmpCn, RstCompMast!CompCode + RstCompMast!CompYear
''            '--------------tblVoucher
''            TmpCn.Execute "alter table tblMastItem add Ittmp3 Long Integer"
''            TmpCn.Execute "update tblMastItem set Ittmp3=0"
''
''            DoEvents
''            DoEvents
            
''            TmpCn.Execute "alter table tblMastNarration add ConDist decimal(7,0) Default 0"
''            TmpCn.Execute "update tblMastNarration set ConDist=0"
''
''            DoEvents
''            DoEvents
''
''            TmpCn.Execute "alter table tblMastAccount add AcDistance decimal(7,0) Default 0"
''            TmpCn.Execute "update tblMastAccount set AcDistance=0"
               
             TmpCn.Execute "alter table tblVoucher add VHsn text(12)"
             DoEvents
             DoEvents
            
            RstCompMast.MoveNext
        Loop

        MsgBox "Tables are modified Successfully", vbOKOnly + vbInformation, "Utility"
        If RstCompMast.State = adStateOpen Then RstCompMast.Close
        If TmpCn.State = adStateOpen Then TmpCn.Close
    End If
End If
Exit Sub
ErrorHandler:
    Select Case Err.Number
        Case -2147467259, -2147217900 '--Column Already Exists,Not Exists
                If MsgBox(Err.Description, vbAbortRetryIgnore + vbQuestion, "Utility") = vbIgnore Then
                    Resume Next
                Else
                    MsgBox "Tables modification is Incomplete.", vbOKOnly + vbInformation, "Utility"
                    Exit Sub
                End If
        Case Else
            GProcErrorHandler
    End Select
'GProcErrorHandler
End Sub
'-------------- Create Tables
Public Sub CreateTables()
Dim X1 As String, X2 As String
If MsgBox("All Tables will get Dropped for Creation.Are you Sure?", vbInformation + vbYesNo + vbDefaultButton2, "Confirm Creation") = vbYes Then
   '===  Drop Tables
    gCn.Execute "Drop Table tblTDS" 'tblTDS
    gCn.Execute "Drop Table tblRecVsSale" 'tblRecVsSale
    gCn.Execute "Drop Table tblMillRecPay" 'tblMillRecPay
    gCn.Execute "Drop Table tblOutstanding" 'tblOutstanding
    gCn.Execute "Drop Table tblMastSetting" 'tblMastSetting
    gCn.Execute "Drop Table tblClosingStock" 'tblClosingStock
    gCn.Execute "Drop Table tblVouSub" 'tblVouSub
    gCn.Execute "Drop Table tblVoucher" 'tblVoucher
    gCn.Execute "Drop Table tblAddLess" 'tblAddLess
    gCn.Execute "Drop Table tblBags" 'tblBags
    gCn.Execute "Drop Table tblMastItem" 'tblMastItem
    gCn.Execute "Drop Table tblMastNarration"  'tblMastNarration
    gCn.Execute "Drop Table tblMastAccount" 'tblMastAccount
    gCn.Execute "Drop Table tblMastCompGroup" 'tblMastCompGroup
    gCn.Execute "Drop Table tblMastGroup" 'tblMastGroup
    gCn.Execute "Drop Table tblConfirmation" 'tblConfirmation
    


'1> tblMastCompGroup
gCn.Execute "Create Table tblMastCompGroup(CGCode     Number (7,0) Constraint PKMastCmpCGCode Primary Key ," _
                                        & "CGName    Varchar2 (40) not null Constraint UnMastCmpCGName Unique," _
                                        & "CGAdd1    Varchar2 (40)," _
                                        & "CGAdd2    Varchar2 (40) ," _
                                        & "CGAdd3    Varchar2 (40)  ," _
                                        & "CGCity    Varchar2 (20)  ," _
                                        & "CGContPer Varchar2 (40)  ," _
                                        & "CGContNo  Varchar2 (40)  ," _
                                        & "CGTmp1    Varchar2 (20)  ," _
                                        & "CGTmp2    Varchar2 (10)  ," _
                                        & "CGTmp3    Number (12,2) Default 0 )"

gCn.Execute "Insert into tblMastCompGroup (CgCode,CgName) Values (1,'-')"
'2> tblMastGroup

gCn.Execute "Create Table tblMastGroup(AgCode     Number (7,0) Constraint PKMastGrpAgCode Primary Key ," _
                                    & "AgName     Varchar2 (40) not null Constraint UnMastGrpAGName Unique," _
                                    & "GpCode     Number (7,0)   Constraint RfMastGrpGpCode References tblMastGroup(AgCode) ," _
                                    & "Schedule   Varchar2 (10) ," _
                                    & "TopGroup   Number (7,0) Default 90001 Constraint ChkMastGrpTopGrp   Check(TopGroup in(90001,90002,90003))," _
                                    & "IsFixGroup Number (2,0) Default 0," _
                                    & "GpTmp1     Number (10,2)  Default 0," _
                                    & "GpTmp2     Varchar2 (20)  )"
                                    
gCn.Execute "Insert into tblMastGroup (AgCode,AgName,GpCode,TopGroup,IsFixGroup) Values (90001,'Asset',90001,90001,1)"
gCn.Execute "Insert into tblMastGroup (AgCode,AgName,GpCode,TopGroup,IsFixGroup) Values (90002,'Liability',90002,90002,1)"
gCn.Execute "Insert into tblMastGroup (AgCode,AgName,GpCode,TopGroup,IsFixGroup) Values (90003,'Revenue',90003,90003,1)"
gCn.Execute "Insert into tblMastGroup (AgCode,AgName,GpCode,TopGroup,IsFixGroup) Values (90004,'Capital Account',90002,90002,1)"
gCn.Execute "Insert into tblMastGroup (AgCode,AgName,GpCode,TopGroup,IsFixGroup) Values (90005,'Reserve And Surplus',90002,90002,1)"
gCn.Execute "Insert into tblMastGroup (AgCode,AgName,GpCode,TopGroup,IsFixGroup) Values (90006,'Loans Liabilities',90002,90002,1)"
gCn.Execute "Insert into tblMastGroup (AgCode,AgName,GpCode,TopGroup,IsFixGroup) Values (90007,'Current Liabilities',90002,90002,1)"
gCn.Execute "Insert into tblMastGroup (AgCode,AgName,GpCode,TopGroup,IsFixGroup) Values (90008,'Fixed Assets',90001,90001,1)"
gCn.Execute "Insert into tblMastGroup (AgCode,AgName,GpCode,TopGroup,IsFixGroup) Values (90009,'Investments',90001,90001,1)"
gCn.Execute "Insert into tblMastGroup (AgCode,AgName,GpCode,TopGroup,IsFixGroup) Values (90010,'Deposit Assets',90001,90001,1)"
gCn.Execute "Insert into tblMastGroup (AgCode,AgName,GpCode,TopGroup,IsFixGroup) Values (90011,'Current Assets',90001,90001,1)"
gCn.Execute "Insert into tblMastGroup (AgCode,AgName,GpCode,TopGroup,IsFixGroup) Values (90013,'Secured Loans',90006,90002,1)"
gCn.Execute "Insert into tblMastGroup (AgCode,AgName,GpCode,TopGroup,IsFixGroup) Values (90014,'Unsecured Loans',90006,90002,1)"
gCn.Execute "Insert into tblMastGroup (AgCode,AgName,GpCode,TopGroup,IsFixGroup) Values (90015,'Sundry Creditors',90007,90002,1)"
gCn.Execute "Insert into tblMastGroup (AgCode,AgName,GpCode,TopGroup,IsFixGroup) Values (90016,'Brokers (Sale Commission Agents)',90007,90002,1)"
gCn.Execute "Insert into tblMastGroup (AgCode,AgName,GpCode,TopGroup,IsFixGroup) Values (90017,'Sundey Debtors',90011,90001,1)"
gCn.Execute "Insert into tblMastGroup (AgCode,AgName,GpCode,TopGroup,IsFixGroup) Values (90018,'Stock In Hand',90011,90001,1)"
gCn.Execute "Insert into tblMastGroup (AgCode,AgName,GpCode,TopGroup,IsFixGroup) Values (90019,'Cash & Bank Balances',90011,90001,1)"
gCn.Execute "Insert into tblMastGroup (AgCode,AgName,GpCode,TopGroup,IsFixGroup) Values (90021,'Trading & Manufacturing A/c',90003,90003,1)"
gCn.Execute "Insert into tblMastGroup (AgCode,AgName,GpCode,TopGroup,IsFixGroup) Values (90022,'Profit & Loss A/c',90003,90003,1)"
gCn.Execute "Insert into tblMastGroup (AgCode,AgName,GpCode,TopGroup,IsFixGroup) Values (90023,'Sales',90021,90003,1)"
gCn.Execute "Insert into tblMastGroup (AgCode,AgName,GpCode,TopGroup,IsFixGroup) Values (90024,'Income',90021,90003,1)"
gCn.Execute "Insert into tblMastGroup (AgCode,AgName,GpCode,TopGroup,IsFixGroup) Values (90025,'Purchase',90021,90003,1)"
gCn.Execute "Insert into tblMastGroup (AgCode,AgName,GpCode,TopGroup,IsFixGroup) Values (90026,'Expenses Direct (T&M)',90021,90003,1)"
gCn.Execute "Insert into tblMastGroup (AgCode,AgName,GpCode,TopGroup,IsFixGroup) Values (90027,'Other Income',90022,90003,1)"
gCn.Execute "Insert into tblMastGroup (AgCode,AgName,GpCode,TopGroup,IsFixGroup) Values (90028,'Expenses Indirect (P&L)',90022,90003,1)"
gCn.Execute "Insert into tblMastGroup (AgCode,AgName,GpCode,TopGroup,IsFixGroup) Values (90029,'Sundry Creditors For Mills',90007,90002,1)"
gCn.Execute "Insert into tblMastGroup (AgCode,AgName,GpCode,TopGroup,IsFixGroup) Values (90037,'Creditors For Expenses',90015,90002,1)"
gCn.Execute "Insert into tblMastGroup (AgCode,AgName,GpCode,TopGroup,IsFixGroup) Values (90038,'Duties And Taxes',90007,90002,1)"
                                  
'3> tblMastAccount
X1 = " AcCode     Number (7,0) Constraint PKAcMastAcCode Primary Key ," _
    & "AcName     Varchar2 (40) not null Constraint UnAcMastAcName Unique," _
    & "AgCode     Number (7,0)  Constraint RfAcMastGpCode References tblMastGroup(AgCode)," _
    & "AgCode1    Number (7,0)  Constraint RfAcMastGpCode1 References tblMastGroup(AgCode)," _
    & "AcAdd1     Varchar2 (40) ," _
    & "AcAdd2     Varchar2 (40)  ," _
    & "AcAlName   Varchar2 (40) ," _
    & "AcCity     Varchar2 (20)  ," _
    & "AcState    Varchar2 (20)  ," _
    & "AcDist     Varchar2 (20)  ," _
    & "AcZip      Varchar2 (10)  ," _
    & "AcContPer  Varchar2 (40)  ," _
    & "AcContNo   Varchar2 (40)  ," _
    & "AcBST      Varchar2 (30)  ," _
    & "AcCST      Varchar2 (30)  ," _
    & "AcPAN      Varchar2 (10)  ," _
    & "AcCmpCode  Number (7,0)  Constraint RfAcMastCmpCode References tblMastCompGroup(CgCode)," _
    & "AcBrkCode  Number (7,0) Constraint RfAcMastBrkCode References tblMastAccount(AcCode)," _
    & "AcBillSrNo  Varchar2 (3)  ," _
    & "AcIntPer   Number (5,2)  Default 0," _
    & "AcIsNonWork  Number (2,0) Default 0 Constraint ChkAcMastNonWork   Check(AcIsNonWork in(1,0)) ," _
    & "AcIsShowOuts Number (2,0) Default 0 Constraint ChkAcMastShowOuts  Check(AcIsShowOuts in(1,0)) ," _
    & "AcTDSNature  Varchar2 (15)  ,"
X2 = " AcEmail      Varchar2 (25)  ," _
    & "AcComOn      Number (2,0)  Default 0," _
    & "AcComRt      Number (5,2)  Default 0," _
    & "AcIsLocalBrkBySelf Number (2,0) Default 0," _
    & "AcIsTDSPay     Number (2,0)  Default 0," _
    & "AcIsSrvTaxPay  Number (2,0)  Default 0," _
    & "AcIsExemptMill Number (2,0) Default 0 ," _
    & "AcExemptRate   Number (6,2)  Default 0," _
    & "AcOsLimit      Number (12,2)  Default 0," _
    & "AcIsLedgerTot  Number (2,0) Default 0," _
    & "AcMillType     Number (2,0)  Default 0," _
    & "AcIsDbNtAc Number (1)  Default 0,AcComType  Number (1,0)  Default 0,AcFaxNo   Varchar2 (40) Default Null"
gCn.Execute "Create Table tblMastAccount(" & X1 & X2 & ")"
gCn.Execute "Insert into tblMastAccount (AcCode,AgCode,AgCode1,AcName,AcCmpCode,AcBrkCode) Values (1,90016,90016,'-',1,1)"


'4> tblMastNarration
gCn.Execute "Create Table tblMastNarration(NarrType  Varchar2 (2) Constraint NLMastNarrNarrType not null ," _
                                        & "NarrCode  Number (7,0) Constraint PKMastNarrNarrCode Primary Key ," _
                                        & "Narration Varchar2 (40) not null Constraint UnMastNarrNarration Unique," _
                                        & "GodAdd1    Varchar2 (40) ," _
                                        & "GodAdd2    Varchar2 (40)  ," _
                                        & "GodCity    Varchar2 (20)  ," _
                                        & "GodPhone  Varchar2 (40)  ," _
                                        & "MastTaxRate Number (5,2) Default 0 ," _
                                        & "MastTaxPurAcCode Number (7,0)  Constraint RfMastNarrPurTaxAc References tblMastAccount(AcCode)," _
                                        & "MastTaxSAAcCode Number (7,0)  Constraint RfMastNarrSaleTaxAc References tblMastAccount(AcCode)," _
                                        & "NrTmp1    Varchar2 (20)  ," _
                                        & "NrTmp2    Varchar2 (10)  ," _
                                        & "NrTmp3    Number (12,2) Default 0 )"


gCn.Execute "Insert into tblMastNarration (NarrType,NarrCode,Narration,MastTaxPurAcCode,MastTaxSAAcCode) Values ('N',1,'-',1,1)"
'5> tblMastItem
gCn.Execute "Create Table tblMastItem(ItCode        Number (7,0) Constraint PKMastItItCode Primary Key," _
                                    & "ItName       Varchar2 (40) Constraint NLMastItItName not null ," _
                                    & "ItTicket       Varchar2 (20)  Constraint NLMastItTicket not null ," _
                                    & "ItShort      Varchar2 (10) Constraint NLMastItItShort not null," _
                                    & "ItStdPack    Number (8,3)  Default 0 ," _
                                    & "ItStdRatePer Number (5,2) Default 0 ," _
                                    & "ItUnit       Varchar2 (7)  ," _
                                    & "ItBrokOn     Varchar2 (3) ," _
                                    & "ItBrokRt     Number (5,2)  Default 0," _
                                    & "ItMillCode   Number (7,0) Constraint RfMastItMillAcCode References tblMastAccount(AcCode)," _
                                    & "ItOutBrkRt   Number (5,2)  Default 0," _
                                    & "ItTmp1   Number (10,2) Default 0 ," _
                                    & "ItTmp2   Varchar2 (20)  ," _
                                    & " ITMAXRATE NUMBER(10,2) Default 0," _
                                    & " ITMINRATE NUMBER(10,2) Default 0," _
                                    & " Constraint UnMastItItNameTicket Unique(ItName,ItTicket))"

'6> tblAddLess
X1 = " Vno          Number (7,0) Default 0," _
    & "VType        Varchar2 (2) ," _
    & "Vdt          Date Default SYSDate," _
    & "AdAcCrCode   Number (7,0)  Constraint RfTblALCrCode References tblMastAccount(AcCode), " _
    & "AdAcDrCode   Number (7,0) Constraint RfTblALDrCode References tblMastAccount(AcCode) ," _
    & "AdMillCode   Number (7,0)  Constraint RfTblALMlCode References tblMastAccount(AcCode)," _
    & "AdGpNo       Number (7,0)  Default 0 ," _
    & "AdGpDt       Date Default SYSDate," _
    & "AdBookNo     Number (7,0)  Default 0 ," _
    & "AdBookDt     Date Default SYSDate ," _
    & "AdSubAmt     Number (12,2) Default 0 ," _
    & "AdBillAmt    Number (12,2) Default 0 ," _
    & "AdTaxCode    Number (7,0) Constraint RfTblALTaxCode References tblMastNarration(NarrCode)," _
    & "AdTaxP       Number (5,2) Default 0 ," _
    & "AdTaxAmt     Number (12,2) Default 0 ," _
    & "AdExemptAmt  Number (12,2) Default 0 ," _
    & "AdExemptPerKg Number (7,2) Default 0 ," _
    & "AdRoff       Number (6,2) Default 0 ," _
    & "AdOtherAdd   Number (12,2) Default 0 ," _
    & "AdOtherLess  Number (12,2) Default 0 ," _
    & "AdDbNtNo     Number (7,0)  Default 0," _
    & "AdDbNtAmt    Number (12,2) Default 0 ," _
    & "AdDbNtRt     Number (12,2) Default 0  ,"
X2 = " VYear        Number (4,0) Default 0 ," _
    & "AdAssVal     Number (12,2) Default 0 ," _
    & "AdLCostFreight Number (12,2) Default 0 ," _
    & "AdLCostGRent Number (12,2) Default 0 ," _
    & "AdLCostAdd   Number (12,2) Default 0 ," _
    & "AdLCostLess  Number (12,2) Default 0  ," _
    & "AdLCostTot   Number (12,2) Default 0 ," _
    & "AdLCostPerBag Number (12,2) Default 0 ," _
    & "AdTotBag      Number (12,2)  Default 0," _
    & "AdLrNo        Varchar2 (10) ," _
    & "AdLrDt        Date Default Null," _
    & "AdDbNtDt      Date Default SYSDate," _
    & "AdTransport   Varchar2 (20) ," _
    & "AdInvNo       Varchar2 (10)," _
    & "AdCFormNo     Varchar2 (10) ," _
    & "AdE1FormNo    Varchar2 (10) ," _
    & "AdTmp1         Number (10,2)  Default 0," _
    & "AdTmp2         Number (10,2)  Default 0," _
    & "AdTmp3         Number (10,2)  Default 0," _
    & "AdTmp4         Number (10,2)  Default 0," _
    & "AdTmp5         Varchar2 (40) , AdTmp6         Varchar2 (20) , AdTmp7         Varchar2 (20) , " _
    & "AdTmp8         Varchar2 (20) , ADBookRtType number(1) default 0 ,AdAcDbNtCode Number (7,0)  Default 0,AdAftTaxAdd   Number (12,2) Default 0,AdAftTaxLess Number (12,2) Default 0," _
    & " Constraint PKTblALVnoVtypeVYr Primary Key(Vno,Vtype,vYEAR)"
   'AdTmp5 sizer  for gatePass
gCn.Execute "Create Table tblAddLess(" & X1 & X2 & ")"


'7> tblVoucher
gCn.Execute "Create Table tblVoucher(Vno     Number (7,0)  Default 0," _
                                & " VCtrNo   Number (3,0) Default 1," _
                                & " VDt      Date Default SYSDate," _
                                & " Vtype    Varchar2 (2)  ," _
                                & " VAcCode  Number (7,0) Constraint RfTblVouAcCode References tblMastAccount(AcCode)," _
                                & " VAcOCode Number (7,0) Constraint RfTblVouAcOCode References tblMastAccount(AcCode)," _
                                & " VAmt     Number (12,2) Default 0," _
                                & " VNar1    Varchar2 (60)  ," _
                                & " VNar2    Varchar2 (40)  ," _
                                & " VNar3    Varchar2 (40)  ," _
                                & " VNar4    Varchar2 (40)  ," _
                                & " VBillNo  Varchar2 (10)  ," _
                                & " VBillDt  Date Default SYSDate," _
                                & " VReconDt Date Default SYSDate," _
                                & " VRecAmt  Number (15,2)  Default 0," _
                                & " VUserName Varchar2 (10)     Default 'Admin'," _
                                & " VEntStatus  Varchar2 (1)      Default 'A'," _
                                & " VYear       Number   (13,0)   Default 0," _
                                & " VIsAudited  Number   (13,0)   Default 0 ," _
                                & " VIsRefEntType  Number   (1,0)   Default 0 ," _
                                & " MillExpNarrCode NUMBER(7) default 0," _
                                & " Constraint PKTblVouVnoVtypeVyrCtrNo Primary Key(Vno,Vtype,vcTRnO,vYEAR))"
'8> tblVouSub
X1 = "VNo        Number (7,0)  Default 0, " _
  & "VType      Varchar2 (2) , " _
  & "VDt        Date Default SYSDate, " _
  & "VSubCtrNo  Number (2,0)  Default 1," _
  & "VSubItCode Number (7,0) Constraint RfTblVouSubItCode References tblMastItem(ItCode)," _
  & "VSubBag    Number (5,0)  Default 0," _
  & "VSubWt     Number (10,3)  Default 0," _
  & "VSubRt     Number (11,5)  Default 0," _
  & "VSubRtPer  Number (5,2)  Default 1," _
  & "VSubAmt    Number (12,2) Default 0," _
  & "VSubLotNo  Varchar2 (10) ," _
  & "VSubItCtrlNo  Number (7,0) Default 0 Constraint UnTblVouSubItCtrlNo Unique," _
  & "VSubPItCtrlNo  Number (7,0) Default 0," _
  & "VSubIssBag   Number (5,0) Default 0," _
  & "VSubIssWt   Number (10,3) Default 0," _
  & "VSubTmpBag   Number (5,0) Default 0," _
  & "VSubTmpWt   Number (10,3) Default 0," _
  & "VSubGodownCode Number (7,0) Constraint RfTblVouSubGodCode References tblMastNarration(NarrCode),"
X2 = "VYear         Number (4,0) Default 0," _
  & "VSubLrNo       Varchar2(20) ," _
  & "VSubLrDate     Date Default SYSDate, " _
  & "VSubLrAmt     Number (12,2) Default 0," _
  & "VSubNetRate   Number (9,3) Default 0," _
  & "VSubGPSLType   Number (1,0) Default 0," _
  & "VSubTmp1       Number (12,2) Default 0," _
  & "VSubTmp2       Number (13,2) Default 0," _
  & "VSubTmp3       Number (13,2) Default 0," _
  & "VSubTmp4       Number (13,0) Default 0," _
  & "VSubTmp5       Varchar2 (25), " _
  & "VSubTmp6       Varchar2 (25), " _
  & "VSubTmp7       Varchar2 (25), " _
  & "VSubBookAcCode Number (7) Default 1," _
  & " VSubBagPerLR1 NUMBER(7) default 0," _
  & " VSubBagPerLR2 NUMBER(7) default 0," _
  & " VSubBagPerLR3 NUMBER(7) default 0," _
  & " VSubBagPerLR4 NUMBER(7) default 0," _
  & " VSubNoOfLR1 NUMBER(7) default 0," _
  & " VSubNoOfLR2 NUMBER(7) default 0,VSubNoOfLR3 NUMBER(7) default 0,VSubNoOfLR4 NUMBER(7) default 0," _
  & " VSubBookBrkAcCode Number (7) Default 0,VSubBookIsCompleted Number (1) Default 0," _
  & " VSubBookNo Number(7) Default 0,VSubBookDt Date Default Null,VSubBkItCtrlNo  Number (7,0) Default 0,VSubItStdPack    Number (8,3)  Default 0," _
  & " Constraint PKTblVouSubVnoVtypeVyrCtrNo Primary Key(Vno,Vtype,VSubCtrNo,vYEAR)," _
  & " Constraint RfTblVouSubVnoVtypeVyr foreign key (vno,vtype,vyear) References tblAddLess(Vno,Vtype,VYear)"
gCn.Execute " Create Table tblVouSub(" & X1 & X2 & " )"

'9> tblClosing Stock
gCn.Execute "Create Table tblClosingStock(StAcCode  Number (7,0)  Constraint RfTblClosStkAcCode References tblMastAccount(AcCode)," _
                                        & "StDate   Date Default SYSDate," _
                                        & "StAmount Number (12,2) Default 0," _
                                        & " Constraint PKTblClosStkAcCodeDt Primary Key(StAcCode,StDate))"
'10> tblMastSetting
gCn.Execute "Create Table tblMastSetting(TaxCodeSY   Number (7,0) Default 0 ," _
                                    & "TaxCodeSO   Number (7,0) Default 0 ," _
                                    & "TaxCodeSD   Number (7,0) Default 0 ," _
                                    & "TaxCodeSM   Number (7,0) Default 0 ," _
                                    & "TaxCodeST   Number (7,0) Default 0 ," _
                                    & "TaxCodePY   Number (7,0) Default 0 ," _
                                    & "TaxCodePT   Number (7,0) Default 0 ," _
                                    & "LatePayIntAcCode   Number (7,0) Default 0 ," _
                                    & "BillDbNtAcCode     Number (7,0) Default 0 ," _
                                    & "ServiceTaxAcCode    Number (7,0) Default 0 ," _
                                    & "TDSAcCode          Number (7,0) Default 0 ," _
                                    & "BankCommAcCode     Number (7,0) Default 0 ," _
                                    & "BrokerageAcCode    Number (7,0) Default 0 ," _
                                    & "AcCodePY      Number (7,0) Default 0 ," _
                                    & "AcCodePT      Number (7,0) Default 0 ," _
                                    & "AcCodeSY      Number (7,0) Default 0 ," _
                                    & "AcCodeST      Number (7,0)  Default 0," _
                                    & "ExpensesAcCode Number (7,0) Default 0," _
                                    & "MILLBILLDBNTACCODE Number (7,0) Default 0," _
                                    & "IsNoVATVoucherSY NUMBER(1) default 0, " _
                                    & "IsNoVATVoucherSO NUMBER(1) default 0, " _
                                    & "IsNoVATVoucherSD NUMBER(1) default 0,IsMillBillAmtWithDbNt NUMBER(1) default 0,TDSPaidAcCode  Number (7,0) Default 0 " _
                                    & ")"
'11> tblOutstanding
gCn.Execute "Create Table tblOutstanding(Vno     Number (7,0) Default 0," _
                                & " Vtype        Varchar2 (2)  ," _
                                & " VYear        Number   (4,0)   Default 0," _
                                & " VDt          Date Default SYSDate," _
                                & " OutBillNo    Varchar2 (10)  ," _
                                & " OutAcCode    Number (7,0) Constraint RfTblOutStdAcCode References tblMastAccount(AcCode)," _
                                & " OutBrokerCode Number (7,0)  Constraint RfTblOutStdBrkCode References tblMastAccount(AcCode)," _
                                & " OutAmount     Number (12,2) Default 0," _
                                & " OutDbNtAmt    Number (12,2) Default 0," _
                                & " OutDueDays    Number (4,0) Default 0," _
                                & " OutDueDate    Date Default SYSDate," _
                                & " OutMillCode   Number (7,0) Constraint RfTblOutStdMillCode References tblMastAccount(AcCode)," _
                                & " OutItShort  Varchar2 (10)," _
                                & " OutBag  Number (5,0) Default 0," _
                                & " OutTicket    Varchar2 (20)  ," _
                                & " OutFreight  Number (12,2) Default 0 ," _
                                & " OutRecAmt  Number (12,2) Default 0," _
                                & " OutRecTmp  Number (12,2) Default 0," _
                                & " OutRecDbAmt  Number (12,2)Default 0 ," _
                                & " OutRecDbTmp  Number (12,2) Default 0," _
                                & " OutRecFreight  Number (12,2) Default 0 ," _
                                & " OutRecFreightTmp  Number (12,2) Default 0," _
                                & " Constraint PKTblOutStdVnoVtypeVYr Primary Key(Vno,Vtype,vYEAR))"

'12> tblRecVsSale
gCn.Execute "Create Table tblRecVsSale(AdjNo     Number (7,0) Constraint PKRecVsSAAdjNo Primary Key," _
                                & " InvNo        Number (7,0) Default 0, " _
                                & " InvDt        Date Default SYSDate," _
                                & " InvTp        Varchar2 (2)  ," _
                                & " InvYear      Number   (4,0) Default 0 ," _
                                & " InvBillno    Varchar2(10) ," _
                                & " BillAmt      Number (12,2) Default 0 ," _
                                & " DbNtAmt      Number (12,2) Default 0 ," _
                                & " PartyCode     Number (7,0)  Constraint RfRecVsSAPartyCode References tblMastAccount(AcCode)," _
                                & " AdjAmt      Number (12,2) Default 0," _
                                & " AdjDbNtAmt      Number (12,2) Default 0," _
                                & " Discount    Number (12,2) Default 0," _
                                & " LateDays    Number (4,0) Default 0," _
                                & " IntReceived Number (12,2) Default 0," _
                                & " Interest    Number (12,2) Default 0," _
                                & " RecVNo      Number (7,0) Default 0  ," _
                                & " RecVType    Varchar2 (2) ," _
                                & " RecVYear    Number (4,0) Default 0," _
                                & " RecVCtrlNo  Number (2,0) Default 0," _
                                & " RecVDt      Date Default SYSDate," _
                                & " RecAmt      Number (12,2) Default 0," _
                                & " JvNoDisc    Number (7,0) Default 0," _
                                & " JvNoInt     Number (7,0) Default 0, Nar1        Varchar2 (40) ," _
& "IntDbNtDate Date Default Null ,IntDbNtRate Number (6,2) Default 0,IntDbNtGraseDay Number (4) Default 0, " _
& " IntDbNtLateDay Number (7) Default 0,IsIntDbNtLateFromInvDt Number(1,0) Default 0 ,IntDbNtAmt Number(13,2) Default 0,IntDbNtRecdAmt Number(13,2) Default 0,IsIntDbNtComplete Number(1) Default 0,AddAmt Number (7,2) Default 0,LessAmt Number (7,2) Default 0,LessTDS    Number (12,2) Default 0,JvNoLessTDS     Number (7,0) Default 0)"
            
'13> tblMillRecPay
gCn.Execute "Create Table tblMillRecPay(Vno     Number (7,0)  Default 0," _
                                & " VCtrNo   Number (3,0) Default 1," _
                                & " VDt      Date Default SYSDate," _
                                & " Vtype    Varchar2 (2)  Not Null," _
                                & " MRecAcCode  Number (7,0) Constraint RfTblMillRecPayAcCode References tblMastAccount(AcCode)," _
                                & " MRecAcOCode Number (7,0) Constraint RfTblMillRecPayAcOCode References tblMastAccount(AcCode)," _
                                & " MRecAmt     Number (12,2) Default 0," _
                                & " MRecNar1    Varchar2 (40)  ," _
                                & " MRecNar2    Varchar2 (40)  ," _
                                & " MRecNar3    Varchar2 (40)  ," _
                                & " MRecNar4    Varchar2 (40)  ," _
                                & " MRecBillNo  Varchar2 (10)  ," _
                                & " MRecBillDt  Date Default SYSDate," _
                                & " MRecReconDt Date Default SYSDate," _
                                & " VUserName Varchar2 (10)     Default 'Admin'," _
                                & " VEntStatus  Varchar2 (1)      Default 'A'," _
                                & " VYear       Number   (13,0)   Default 0," _
                                & " VIsAudited  Number   (13,0)   Default 0 ," _
                                & " MRecMillExpNarrCode NUMBER(7) default 0," _
                                & " Constraint PKTblMillRecVnoVtpVyrCtrNo Primary Key(Vno,Vtype,vcTRnO,vYEAR))"
'14> tblBags
gCn.Execute "Create Table tblBags(SNo Number (13)  Default 0 Constraint PKBagsSNo Primary Key , " _
                                & "BagNo Varchar2 (15) not null," _
                                & "Cone Number (5,0)  Default 0," _
                                & "Wt Number (10,3)  Default 0," _
                                & "InwNo        Number (7,0)  Default 0, " _
                                & "InwType      Varchar2 (2) default Null, " _
                                & "GPNo        Number (7,0)  Default 0, " _
                                & "GPType      Varchar2 (2) default Null) "
'15> tblTDS '-- TDS Entry
gCn.Execute "Create Table tblTDS(Vno          Number (7) Default 0," _
    & "VCtrlNo  Number (2) Default 0," _
    & "VYear        Number (4) Default 0 ," _
    & "Vdt          Date Default SYSDate," _
    & "TDSAcCode Number (7)  Constraint RfTblTDSACCode References tblMastAccount(AcCode), " _
    & "TDSChNo    Varchar2 (20) default Null ," _
    & "TDSChDt     Date Default Null," _
    & "TDSFromDt   Date Default Null," _
    & "TDSToDt      Date Default Null," _
    & "TDSPayType  Number (1)  Default 0 /* 0-Interest,1-Brokerage,2-Commission,3-Tech.Fee,4-Labour/Contract */," _
    & "TDSBillNo  Varchar2 (10) Default Null," _
    & "TDSCrDt     Date Default Null," _
    & "TDSCrAmt     Number (12,2) Default 0 ," _
    & "TDSTDSAmt     Number (12,2) Default 0 ," _
    & "TDSSurchgAmt  Number (12,2) Default 0 ," _
    & "TDSEdCessAmt  Number (12,2) Default 0 ," _
    & "TDSTaxAmt  Number (12,2) Default 0 ," _
    & "TDSChqNo  Varchar2 (20) Default null," _
    & "TDSBSRNo  Varchar2 (15) Default null ," _
    & "TDSPlace  Varchar2 (20) Default null ," _
    & "TDSDate   Date Default Null," _
    & "TDSDeductTaxAcCode   Number (7,0) Constraint RfTblTDSDeductTaxACCode References tblMastAccount(AcCode), " _
    & "TDSTotTaxAmt  Number (12,2) Default 0, " _
    & " Constraint PKTblTDSVnoVYrVCtrlNo Primary Key(Vno,vYEAR,VCtrlNo))" ',TDSDepositDt Date Default Null,
'16> tblConfirmation '-- Confirmation Entry
gCn.Execute "Create Table tblConfirmation(Vno          Number (13) Default 0," _
    & "VYear        Number (4) Default 0 ," _
    & "Vdt          Date Default SYSDate," _
    & "Vtype    Varchar2 (2) default Null ," _
    & "CfBillNo    Varchar2 (10) default Null ," _
    & "CfCrAcCode Number (7)  Constraint RfTblConfirmCrACCode References tblMastAccount(AcCode), " _
    & "CfDrAcCode Number (7)  Constraint RfTblConfirmDrACCode References tblMastAccount(AcCode), " _
    & "CfItemNarr  Varchar2 (40) Default Null," _
    & "CfQty        Number (10,2) Default 0 ," _
    & "CfItUnit      Varchar2 (10) Default Null /* Bags,Cartons,Kgs,Load,Ton,Truck*/," _
    & "CfRate       Number (8,2) Default 0 ," _
    & "CfDeliveryNarr  Varchar2 (40) Default Null," _
    & "CfPackingNarr1  Varchar2 (40) Default Null," _
    & "CfPackingNarr2  Varchar2 (60) Default Null," _
    & "CfPaymentNarr  Varchar2 (30) Default Null," _
    & "CfItRemark  Varchar2 (20) Default Null," _
    & "CfDestinationNarr1  Varchar2 (40) Default Null," _
    & "CfDestinationNarr2  Varchar2 (50) Default Null," _
    & "CfDestinationNarr3  Varchar2 (40) Default Null," _
    & "CfCashDiscountP   Number (6,2) Default 0, " _
    & "CfCommissionP   Number (6,2) Default 0, " _
    & "CfRemark  Varchar2 (40) Default Null,CfRemark2  Varchar2 (50) Default Null,CfRateType Varchar2 (10) Default Null" _
    & " Constraint PKTblConfirmVnoVYrVType Primary Key(Vno,vYEAR,Vtype))"
  End If
End Sub
Public Sub AccountTransferToOtherFirm()
Dim mTempCn As Connection
Dim rst1 As Recordset
Dim rst2 As Recordset
Dim X As String, X1 As String, i As Integer, j As Long
Dim wPFBal As Double
wPFBal = 0
If ((GProcRstOpen(rst1, "Select * from tblmastaccount order by AcCode", "R", gCn)) = 0) Then
   Exit Sub
End If
    '----New Temparary Connection
X = gDbNm
gDbNm = InputBox("Give Short Name to Transfer Account List ")
GProcConnectionOpenORA mTempCn
gDbNm = X
Set rst2 = New Recordset
GProcRstOpen rst2, "Select * from tblmastaccount order by AcName", "O", mTempCn
rst1.MoveLast
j = rst1!AcCode + 1
rst2.MoveLast
If (rst2!AcCode >= j) Then j = rst2!AcCode + 1
rst1.MoveFirst
Do While Not rst1.EOF
        rst2.MoveFirst
        rst2.Find "AcName='" & rst1!AcName & "'"
        If rst2.EOF = False Then
           rst2.MoveFirst
           rst2.Find "AcCode=" & rst1!AcCode
           If rst2.EOF = True Then
              rst2.Fields(0) = j
              j = j + 1
           Else
              rst2.Fields(0) = rst1.Fields(0)
           End If
           For i = 1 To 35
              rst2.Fields(i) = rst1.Fields(i)
           Next i
           rst2.Update
        End If
    rst1.MoveNext
Loop
rst2.Close
rst1.Close
End Sub
Public Sub ItemMastTransferToOtherFirm()
Dim mTempCn As Connection
Dim rst1 As Recordset, mRstNewMastAccount As Recordset
Dim rst2 As Recordset
Dim X As String, X1 As String, i As Integer, j As Long
Dim wPFBal As Double
wPFBal = 0
If ((GProcRstOpen(rst1, "Select * from tblmastItem order by ItCode", "R", gCn)) = 0) Then
   Exit Sub
End If
'----New Temparary Connection
X = gDbNm
gDbNm = InputBox("Give Short Name to Transfer Count List ")
GProcConnectionOpenORA mTempCn
gDbNm = X
Dim mItcode As Long, mNewMillCode As Long
Set rst2 = New Recordset
GProcRstOpen rst2, "Select * from tblmastItem order by ItName", "O", mTempCn
rst1.MoveLast
j = rst1!ItCode + 1
If rst2.RecordCount > 0 Then
rst2.MoveLast
If (rst2!ItCode >= j) Then j = rst2!ItCode + 1
End If
rst1.MoveFirst
Do While Not rst1.EOF
        mItcode = rst1!ItCode
        X = GProcGetColumnValue("tblmastaccount", "Accode", rst1!ItMillCode, "N", "acname", "S", , gCn)
        mNewMillCode = GProcGetColumnValue("tblmastaccount", "Acname", X, "S", "accode", "N", , mTempCn)
        If mNewMillCode <= 0 Then
           GProcRstOpen mRstNewMastAccount, "Select max(accode) from tblMastAccount ", "R", mTempCn
           mNewMillCode = mRstNewMastAccount.Fields(0) + 1
           mRstNewMastAccount.Close
           mTempCn.Execute "Insert into tblMastAccount (AcCode,AgCode,AgCode1,AcName,AcCmpCode,AcBrkCode) Values (" & mNewMillCode & ",90016,90016,'" & X & "',1,1)"
        End If
        If rst2.RecordCount > 0 Then
        rst2.MoveFirst
        rst2.Find "ItName='" & rst1!ItName & "'"
        End If
        If rst2.EOF = True Then
           If rst2.RecordCount > 0 Then
           rst2.MoveFirst
           rst2.Find "ItCode=" & mItcode
           End If
           If rst2.EOF = True Then
              mItcode = j
              j = j + 1
           End If
           rst2.AddNew
           For i = 0 To 12
              rst2.Fields(i) = rst1.Fields(i)
           Next i
           rst2.Fields(0) = mItcode
           rst2!ItMillCode = mNewMillCode
           rst2.Update
        End If
    rst1.MoveNext
Loop
rst2.Close
rst1.Close
End Sub

Public Sub PurchaseStockTransferToNextYear()
    Dim mTempCn As Connection
    Dim X As String, X1 As String, i As Long, j As Long
    '--- TblAddLess
    Dim mRstOld As Recordset
    Dim mRstNew As Recordset
    '-- TblVouSub
    Dim mRstOldSub As Recordset
    Dim mRstNewSub As Recordset
    '-- Mast Item
    Dim mRstOldMastItem As Recordset
    Dim mRstNewMastItem As Recordset
    '-- Mast Account
    Dim mRstOldMastAccount As Recordset
    Dim mRstNewMastAccount As Recordset
    '-- New Reference Code
    Dim mNewAcCrCode As Long, mNewAcDrCode  As Long, mNewMillCode  As Long, mNewGodCode As Long
    Dim mNewItemNo
    
    If ((GProcRstOpen(mRstOld, "Select * from tblAddLess  where VType in ( 'PI','PY','PT') order By Vno", "R", gCn)) = 0) Then
        Exit Sub
    End If
    GProcRstOpen mRstOldSub, "Select tblVouSub.* from tblVouSub where VType in ( 'PI','PY','PT') order By Vno,Vtype", "R", gCn
    'GProcRstOpen mRstOldMastAccount, "Select * from tblMastAccount order By AcName,AcCode", "R", gCn
    'GProcRstOpen mRstOldMastItem, "Select * from tblMastItem order By ItName , ItCode", "R", gCn
    
    '----New Temparary Connection
    X = gDbNm
    gDbNm = InputBox("Give Short Name to Next Year Account  ")
    GProcConnectionOpenORA mTempCn
    gDbNm = X
    
    mTempCn.Execute "delete from tblvousub where vno < 0 "
    mTempCn.Execute "delete from tbladdless where vno < 0 "
    Set mRstNew = New Recordset
    GProcRstOpen mRstNew, "Select * from tblAddless ", "O", mTempCn
    GProcRstOpen mRstNewSub, "Select * from tblVouSub order by Vno,Vtype", "O", mTempCn
   ' GProcRstOpen mRstNewMastAccount, "Select * from tblMastAccount order By AcName,AcCode", "O", mTempCn
   ' GProcRstOpen mRstNewMastItem, "Select * from tblMastItem order By ItName , ItCode", "O", mTempCn
   
    mRstOld.MoveLast
    mRstOld.MoveFirst
    Dim mCnt As Integer
    Do While Not mRstOld.EOF
        mCnt = mRstNew.Fields.Count - 1
        'Chekc Old Ac & if create new
             X = GProcGetColumnValue("tblmastaccount", "Accode", mRstOld!AdAcCrCode, "N", "acname", "S", , gCn)
             mNewAcCrCode = GProcGetColumnValue("tblmastaccount", "Acname", X, "S", "accode", "N", , mTempCn)
             If mNewAcCrCode <= 0 Then
                GProcRstOpen mRstNewMastAccount, "Select max(accode) from tblMastAccount ", "R", mTempCn
                mNewAcCrCode = mRstNewMastAccount.Fields(0) + 1
                mRstNewMastAccount.Close    '-- insert AcMillType
                mTempCn.Execute "Insert into tblMastAccount (AcCode,AgCode,AgCode1,AcName,AcCmpCode,AcBrkCode) Values (" & mNewAcCrCode & ",90016,90016,'" & X & "',1,1)"
             End If
             X = GProcGetColumnValue("tblmastaccount", "Accode", mRstOld!AdAcDrCode, "N", "acname", "S", , gCn)
             mNewAcDrCode = GProcGetColumnValue("tblmastaccount", "Acname", X, "S", "accode", "N", , mTempCn)
             If mNewAcDrCode <= 0 Then
                GProcRstOpen mRstNewMastAccount, "Select max(accode) from tblMastAccount ", "R", mTempCn
                mNewAcDrCode = mRstNewMastAccount.Fields(0) + 1
                mRstNewMastAccount.Close
                mTempCn.Execute "Insert into tblMastAccount (AcCode,AgCode,AgCode1,AcName,AcCmpCode,AcBrkCode) Values (" & mNewAcDrCode & ",90016,90016,'" & X & "',1,1)"
             End If
             X = GProcGetColumnValue("tblmastaccount", "Accode", mRstOld!AdMillCode, "N", "acname", "S", , gCn)
             mNewMillCode = GProcGetColumnValue("tblmastaccount", "Acname", X, "S", "accode", "N", , mTempCn)
             If mNewMillCode <= 0 Then
                GProcRstOpen mRstNewMastAccount, "Select max(accode) from tblMastAccount ", "R", mTempCn
                mNewMillCode = mRstNewMastAccount.Fields(0) + 1
                mRstNewMastAccount.Close
                mTempCn.Execute "Insert into tblMastAccount (AcCode,AgCode,AgCode1,AcName,AcCmpCode,AcBrkCode) Values (" & mNewMillCode & ",90016,90016,'" & X & "',1,1)"
             End If
           mRstNew.AddNew
           For i = 0 To mCnt - 1
               mRstNew.Fields(i) = mRstOld.Fields(i)
           Next i
           mRstNew!AdAcCrCode = mNewAcCrCode
           mRstNew!AdAcDrCode = mNewAcDrCode
           mRstNew!AdMillCode = mNewMillCode
           mRstNew!Vno = mRstNew!Vno * (-1)
           mRstNew.Update '-- Find
           '---------------------------- Update TblVouSub
           mRstOldSub.MoveFirst
           mRstOldSub.Find "vno = " & mRstOld!Vno
'           mRstOldSub.Find ("VtypeNo ='" & mRstOld!Vtype & Trim(Str(mRstOld!Vno)) & "'")
           Do While mRstOldSub.EOF = False
               If mRstOldSub!Vno <> mRstOld!Vno Then Exit Do
               If (mRstOldSub!Vtype = mRstOld!Vtype) Then
              
               'Chekc itemno & add if not
               mCnt = mRstNewSub.Fields.Count - 1
               X = GProcGetColumnValue("tblMastItem", "Itcode", mRstOldSub!vsubITCODE, "N", "ItName", "S", , gCn)
               mNewItemNo = GProcGetColumnValue("tblMastItem", "ItName", X, "S", "ItCode", "N", , mTempCn)
               If mNewItemNo <= 0 Then
                   GProcRstOpen mRstNewMastAccount, "Select max(Itcode) from tblMastItem ", "R", mTempCn
                   mNewItemNo = mRstNewMastAccount.Fields(0) + 1
                   mRstNewMastAccount.Close
                   mTempCn.Execute "Insert into tblMastItem (ItCode,Itname,ItTicket,ItShort,ItMillCode) Values (" & mNewItemNo & ",'" & X & "',' ',' '," & mNewMillCode & ")"
               End If
               X = GProcGetColumnValue("tblMastNarration", "Narrcode", mRstOldSub!VSubGodownCode, "N", "Narration", "S", , gCn)
               mNewGodCode = GProcGetColumnValue("tblMastnarration", "Narration", X, "S", "NarrCode", "N", , mTempCn)
               If mNewGodCode <= 0 Then
                   GProcRstOpen mRstNewMastAccount, "Select max(narrcode) from tblMastNarration ", "R", mTempCn
                   mNewGodCode = mRstNewMastAccount.Fields(0) + 1
                   mRstNewMastAccount.Close
                   mTempCn.Execute "Insert into tblMastNarration (NarrCode,NarrType,Narration) Values (" & mNewGodCode & ",'G','" & X & "')"
               End If
               mRstNewSub.AddNew
               For i = 0 To mCnt
                  mRstNewSub.Fields(i) = mRstOldSub.Fields(i)
               Next i
               mRstNewSub!vsubITCODE = mNewItemNo
               mRstNewSub!VSubGodownCode = mNewGodCode
               mRstNewSub!Vno = mRstNewSub!Vno * (-1)
               mRstNewSub!VSubItCtrlNo = mRstNewSub!VSubItCtrlNo * (-1)
               mRstNewSub.Update
               End If
               mRstOldSub.MoveNext
           Loop
           '-----------------------------------------------------------
           mRstOld.MoveNext
    Loop
    mRstOld.Close
    mRstNew.Close
End Sub
 '-- Access Db insert field
 'Dim TmpS_Cn As Connection '-- Source Connection
'Dim NameOfQuery As String
'Dim NameOfTable As String
'Dim Flag As Long
'Dim i As Long
'Dim f As Long
'Dim TblAr(200) As String
'Dim TblS As TableDef
'Dim tblT As TableDef
'        AllFirmFlag = True
'        '--- company Recordset
'        '----- Open Source  Db Connection
'        GProcConnectionOpen TmpS_Cn, "HiTrix"
'        GProcRstOpen RstCompMast, "Select * from tblCompanyYear,tblMastCompany  where compCode=cCode " _
'        & " and CompFdt=cdate('" & gCYSDate & "')" & IIf(AllFirmFlag = True, "", " and cName='" & gCName & "'"), "R", gSetupCn
'        If RstCompMast.EOF = False Then RstCompMast.MoveFirst
'        Do While Not RstCompMast.EOF
'            '----Open Temp Connection
'            GProcConnectionOpen TmpCn, RstCompMast!CompCode + RstCompMast!CompYear
'             '-----Delete Queries from table (Target ) from current path
'             Do While Not TmpCn.QueryDefs.Count = 0
'                 TmpCn.QueryDefs.Refresh
'                 NameOfQuery = TmpCn.QueryDefs(0).Name
'                 TmpCn.QueryDefs.Delete NameOfQuery
'             Loop
'             '-----Copy Query from source Db in c:\
'            TmpS_Cn.QueryDefs.Refresh
'            For i = 0 To TmpS_Cn.QueryDefs.Count - 1
'                 TmpCn.CreateQueryDef TmpS_Cn.QueryDefs(i).Name, TmpS_Cn.QueryDefs(i).SQL
'            Next
'            MsgBox i & " queries are copied in " & RstCompMast!cName & " in Company Year " & RstCompMast!CompFdt
'            Flag = 0
'            '-----Deleting tables from Target path
'            For i = 0 To TmpCn.TableDefs.Count - 1
'                NameOfTable = TmpCn.TableDefs(i).Name
'                If UCase(Left(NameOfTable, 3)) = UCase("tbl") Then
'                Else
'                    TblAr(Flag) = NameOfTable
'                    Flag = Flag + 1
'                End If
'            Next i
'            For i = 0 To Flag - 1
'                TmpCn.TableDefs.Delete TblAr(i)
'            Next i
'            MsgBox i & " Tables are deleted from " & RstCompMast!cName & " in Company Year " & RstCompMast!CompFdt
'            Flag = 0
'            '-----Copy tables from source Db
'         TmpS_Cn.TableDefs.Refresh
'         TmpCn.TableDefs.Refresh
'         For i = 0 To TmpS_Cn.TableDefs.Count - 1
'             NameOfTable = TmpS_Cn.TableDefs(i).Name
'             If UCase(Left(NameOfTable, 3)) = UCase("tbl") Then
'             Else
'                 Set TblS = TmpS_Cn.TableDefs(NameOfTable)
'                 Set tblT = TmpCn.CreateTableDef(NameOfTable)
'                 On Error Resume Next
'                 With tblT
'                     For f = 0 To TblS.Fields.Count - 1
'                         tblT.Fields.Append .CreateField(TblS.Fields(f).Name, TblS.Fields(f).Type, TblS.Fields(f).Size)
'                             If tblT.Fields(f).Type = dbDouble Or tblT.Fields(f).Type = dbLong Then
'                                 tblT.Fields(f).Required = True
'                                 tblT.Fields(f).DefaultValue = 0
'                             ElseIf tblT.Fields(f).Type = dbText Then
'                                 tblT.Fields(f).AllowZeroLength = True
'                             End If
'                     Next f
'                 End With
'                 On Error GoTo 0
'                 TmpCn.TableDefs.Append tblT
'                 Flag = Flag + 1
'                 TmpS_Cn.TableDefs.Refresh
'             End If
'         Next i
'            MsgBox i & " Tables are checked and " & Flag & " tables are copied  in " & rsCompany.Fields("Cname") & " in Company Year " & rsCompany.Fields("CYear")
'            RstCompMast.MoveNext
       ' Loop
       ' If RstCompMast.State = adStateOpen Then RstCompMast.Close
       ' If TmpCn.State = adStateOpen Then TmpCn.Close

