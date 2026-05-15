Attribute VB_Name = "CommonPro"
Option Explicit
'====  Get Account Commission Type From Given Numeric Value
Public Function GProcGetCommissionType(CommTypeIndex As Integer) As String
Select Case CommTypeIndex
    Case 0
        GProcGetCommissionType = "Commission"
    Case 1
        GProcGetCommissionType = "Cash Discount"
    Case 2
        GProcGetCommissionType = "Sales Incentive"
End Select
End Function
'======== Procedure Update AcCode
'-- Inputs pAcCodeFieldName ( Accode Field Name),pFromAcCode -> (FromAcCode Value ) ,pToAcCode -> (ToAcCode Value )
'-- pCondStr
Public Sub GProcUpdateAcCode(pTblName As String, pAcCodeFieldName As String, pFromAcCode As Long, pToAcCode As Long, Optional pCondStr As String = "", Optional pDbCn As Connection)
    If pDbCn Is Nothing Then
        Set pDbCn = gCn
    End If
    pDbCn.Execute "update " & pTblName & " set " & pAcCodeFieldName & "=" & pToAcCode & " where " & pAcCodeFieldName & "=" & pFromAcCode & IIf(pCondStr = "", "", " and " & pCondStr)
End Sub
'--- Function Return Formatted String From Given RecodSet Field No As Per Data Type
Private Function GetFormatStrFromRSField(rstbl As Recordset, FldNo As Long)
Dim Str
Select Case rstbl.Fields(FldNo).Type
    Case adInteger, 2, 3
         Str = CLng(rstbl.Fields(FldNo))
    Case adNumeric
         Str = Format(rstbl.Fields(FldNo), GProcNumberFormat(rstbl.Fields(FldNo).Precision, rstbl.Fields(FldNo).NumericScale))
    Case adDecimal, 131
         Str = Format(rstbl.Fields(FldNo), GProcNumberFormat(rstbl.Fields(FldNo).Precision, rstbl.Fields(FldNo).NumericScale))
    Case adDouble
         Str = Format(rstbl.Fields(FldNo), GProcNumberFormat(rstbl.Fields(FldNo).Precision, 2))
    Case adChar, adVarWChar, 200
         Str = rstbl.Fields(FldNo) & vbNullString
    Case adDate, 135
        If IsDate(rstbl.Fields(FldNo)) = True Then
            Str = Format(rstbl.Fields(FldNo), "dd/MM/yyyy")
        Else
            Str = "__/__/____"
        End If
End Select
GetFormatStrFromRSField = Str
End Function
'----- Show Record
Public Sub GProcShowRecord(pFrm As Form, pRstbl As Recordset, pMaxno As Long, Optional pOtherTxtFromNo As Long, Optional pOtherTxtToNo As Long)
'On Error GoTo ErrorHandler
Dim i As Long
If pMaxno = 0 Then pMaxno = pRstbl.Fields.Count - 1
    '--- 0 to MaxNo
    For i = 0 To pMaxno 'pRstbl.Fields.Count - 1
        pFrm.txt(i).text = GetFormatStrFromRSField(pRstbl, i)
    Next i
    '----- (Show Other Text Controls) pOtherTxtFromNo To pOtherTxtToNo
If pOtherTxtFromNo > pMaxno And pOtherTxtFromNo <= pOtherTxtToNo Then
    For i = pOtherTxtFromNo To pOtherTxtToNo
        pFrm.txt(i).text = GetFormatStrFromRSField(pRstbl, i)
    Next i
End If
Exit Sub
ErrorHandler:
If Err.Number = 340 Or Err.Number = 3265 Then '-- Control Not Exist
    Resume Next
Else
    MsgBox Err.Description
End If
End Sub
'Public Sub GProcShowRecord(pFrm As Form, pRstbl As Recordset, pMaxno As Long, Optional pOtherTxtFromNo As Long, Optional pOtherTxtToNo As Long)
'Dim i As Long
'If pMaxno = 0 Then pMaxno = pRstbl.Fields.Count - 1
'    For i = 0 To pMaxno 'pRstbl.Fields.Count - 1
'        With pFrm.txt(i)
'            Select Case pRstbl.Fields(i).Type
'                Case adInteger, 2, 3
'                    .Text = CLng(pRstbl.Fields(i))
'                Case adNumeric
'                    .Text = Format(pRstbl.Fields(i), GProcNumberFormat(pRstbl.Fields(i).Precision, pRstbl.Fields(i).NumericScale))
'                Case adDecimal, 131
'                    .Text = Format(pRstbl.Fields(i), GProcNumberFormat(pRstbl.Fields(i).Precision, pRstbl.Fields(i).NumericScale))
'                Case adDouble
'                    .Text = Format(pRstbl.Fields(i), GProcNumberFormat(pRstbl.Fields(i).Precision, 2))
'                Case adChar, adVarWChar, 200
'                    .Text = pRstbl.Fields(i) & vbNullString
'                Case adDate, 135
'                    If IsDate(pRstbl.Fields(i)) = True Then
'                        .Text = Format(pRstbl.Fields(i), "dd/MM/yyyy")
'                    Else
'                        .Text = "__/__/____"
'                    End If
'            End Select
'        End With
'    Next i
'    '----- Show Other Text Controls
'If pOtherTxtFromNo > pMaxno And pOtherTxtToNo > pMaxno And pOtherTxtFromNo < pOtherTxtToNo Then
'    For i = pOtherTxtFromNo To pOtherTxtToNo
'        With pFrm.txt(i)
'            Select Case pRstbl.Fields(i).Type
'                Case adInteger, 2, 3
'                    .Text = CLng(pRstbl.Fields(i))
'                Case adNumeric
'                    .Text = Format(pRstbl.Fields(i), GProcNumberFormat(pRstbl.Fields(i).Precision, pRstbl.Fields(i).NumericScale))
'                Case adDecimal, 131
'                    .Text = Format(pRstbl.Fields(i), GProcNumberFormat(pRstbl.Fields(i).Precision, pRstbl.Fields(i).NumericScale))
'                Case adDouble
'                    .Text = Format(pRstbl.Fields(i), GProcNumberFormat(pRstbl.Fields(i).Precision, 2))
'                Case adChar, adVarWChar, 200
'                    .Text = pRstbl.Fields(i) & vbNullString
'                Case adDate, 135
'                    If IsDate(pRstbl.Fields(i)) = True Then
'                        .Text = Format(pRstbl.Fields(i), "dd/MM/yyyy")
'                    Else
'                        .Text = "__/__/____"
'                    End If
'            End Select
'        End With
'    Next i
'End If
'End Sub
'----- Fill ComboBox
Public Sub GProcFillCombo(pCombo As ComboBox, pSql As String, Optional pCon As Connection)
    Dim mRst As New ADODB.Recordset
    Dim i As Integer
    pCombo.Clear
    If pCon Is Nothing Then
        Set pCon = gCn
    End If
    i = GProcRstOpen(mRst, pSql, "R", pCon)
    If i = 0 Then
        mRst.Close
        Exit Sub
    End If
    
    With mRst
        pCombo.Clear
        Do While Not .EOF
            pCombo.AddItem (.Fields(0)) & vbNullString
            .MoveNext
        Loop
    End With
    mRst.Close
    If pCombo.ListCount <> 0 Then pCombo.ListIndex = 0
End Sub
'===== Get Company Details
Public Sub GProcGetCompDetails(Rscomp As Recordset)
Dim X As String
Dim rsC As Recordset
'frmMain.Caption = gCName & " " & Year(Rscomp!CompFdt) & "-" & (Year(Rscomp!CompFdt) + 1) & " GSTIN : " & gCGSTIN
X = "Select * from tblMastCompany where ccode='" & Rscomp!CompCode & "' "
If GProcRstOpen(rsC, X, "O", gSetupCn) > 0 Then
    gCAdd1 = rsC!CAdd1 & vbNullString
    gCAdd2 = rsC!CAdd2 & vbNullString
    gCAdd3 = rsC!CAdd3 & vbNullString
    gCPhNo = rsC!cPh & vbNullString
    gCCST = rsC!cCst & vbNullString
    gCBST = rsC!cBSt & vbNullString
    gCTAN = rsC!cTAN & vbNullString
    gCPAN = rsC!CPan & vbNullString
    gCSTRegNo = rsC!cSTRegNo & vbNullString
    gCSTCNo = rsC!cSTCNo & vbNullString
    gCFileNo = rsC!CFileNo & vbNullString
    gCACSTONo = rsC!CACSTONo & vbNullString
    gCTDS = rsC!CTDS & vbNullString '-- TDS No
    gCTDSCircle = rsC!CTDSCir & vbNullString '-- TDS Circle
    gCDesignation = Choose(CSng(rsC!CDesignation + 1), "Proprietor", "Partner", "Director") '- Designation
    gCBankName = rsC!CBankName & vbNullString
    gCBankAcNo = rsC!CBankAcNo & vbNullString
    gCBankRtgsCode = rsC!CBankRtgsCode & vbNullString
    gCGSTIN = rsC!CGSTIN & vbNullString
    gCIsTcsFirm = rsC!CGIsTcsFirm
    gCTDSFirm = rsC!CTDSFirm & vbNullString
    gCCity = rsC!Ccity & vbNullString
    gCPin = rsC!CPin & vbNullString
End If
rsC.Close
gCSAcCode = Rscomp!CompSAcCode '- Start AcCode
gCEAcCode = Rscomp!CompEAcCode '- End AcCode
gCSAgCode = Val(Rscomp!CompSAgCode) '- Start AgCode
gCEAgCode = Val(Rscomp!CompEAgCode) '- End AgCode
gCSCGCode = Val(Rscomp!CompSCGCode) '- Start CGCode
gCECGCode = Val(Rscomp!CompECGCode) '- End CGCode
gCSItCode = Val(Rscomp!CompSItCode) '- Start ItCode
gCEItCode = Val(Rscomp!CompEItCode) '- End ItCode
gCSNarrCode = Val(Rscomp!CompSNarrCode) '- Start NarrCode
gCENarrCode = Val(Rscomp!CompENarrCode) '- End NarrCode

gCCode = Rscomp!CompCode
gDbYr = Rscomp!CompYear
gDbNm = gCCode + gDbYr
gCYSDate = Rscomp!CompFdt
gCYEDate = Rscomp!CompTDt
gCYear = Year(Rscomp!CompFdt)
frmMain.Caption = gCName & " " & Year(Rscomp!CompFdt) & "-" & (Year(Rscomp!CompFdt) + 1) & "     GSTIN : " & gCGSTIN

If gBackEndDB = gBackEndAccess Then
    Call GProcConnectionOpen(gCn)
Else '-- Oracle
    Call GProcConnectionOpenORA(gCn)
End If
'--- Create Mast Default Record
GProcCreateDefaultRecord
'-- Get Setting Details
GProcGetSettingDetail
End Sub
'---- Get Master Setting Details
Public Sub GProcGetSettingDetail()
On Error GoTo ErrorRoutine
Dim X As String
Dim RsS As Recordset
Dim DefTaxCode As Long
Dim DefAcCode As Long
X = "Select * from tblMastSetting "
gCashAcCode = GProcGetColumnValue("tblMastAccount", "AcName", "Cash In Hand", "S", "AcCode", "N")
    If GProcRstOpen(RsS, X, "O", gCn) > 0 Then
        gTaxCodeSY = RsS!TaxCodeSY
        gTaxCodeSO = RsS!TaxCodeSO
        gTaxCodeSD = RsS!TaxCodeSD
        gTaxCodeSM = RsS!TaxCodeSM
        gTaxCodeST = RsS!TaxCodeST
        gTaxCodePY = RsS!TaxCodePY
        gTaxCodePT = RsS!TaxCodePT
        gLatePayIntAcCode = RsS!LatePayIntAcCode
        gBillDbNtAcCode = RsS!BillDbNtAcCode
        gServiceTaxAcCode = RsS!ServiceTaxAcCode
        gTDSAcCode = RsS!TDSAcCode
        gBankCommAcCode = RsS!BankCommAcCode
        gBrokerageAcCode = RsS!BrokerageAcCode
        gAcCodePY = RsS!AcCodePY
        gAcCodePT = RsS!AcCodePT
        gAcCodeSY = RsS!AcCodeSY
        gAcCodeST = RsS!AcCodeST
        gAcCodePYHank = RsS!AcCodePYHunk
        gAcCodeSYHank = RsS!AcCodeSYHunk
        gTaxCodePYHunk = RsS!TaxCodePYHunk
        gTaxCodeSYHunk = RsS!TaxCodeSYHunk
        
        gExpensesAcCode = RsS!ExpensesAcCode
        gMillBillDbNtAcCode = RsS!MILLBILLDBNTACCODE
        gTDSPaidAcCode = RsS!TDSPaidAcCode '- TDS Paid AcCode
        gIsNoVATVoucherSY = RsS!IsNoVATVoucherSY
        gIsNoVATVoucherSO = RsS!IsNoVATVoucherSO
        gIsNoVATVoucherSD = RsS!IsNoVATVoucherSD
        gIsMillBillAmtWithDbNt = RsS!IsMillBillAmtWithDbNt '- Is Mill Bill Amt With Db Nt
        '---TCS
        gAcCodeTcsPay = RsS!AcCodeTcsPay
        gAcCodeTcsRec = RsS!AcCodeTcsRec
    Else
        DefTaxCode = GProcGetColumnValue("tblMastNarration", "Narration", "-", "S", "NarrCode", "N")
        DefAcCode = GProcGetColumnValue("tblMastAccount ", "AcName", "-", "S", "AcCode", "N")
        gTaxCodeSY = DefTaxCode
        gTaxCodeSO = DefTaxCode
        gTaxCodeSD = DefTaxCode
        gTaxCodeSM = DefTaxCode
        gTaxCodeST = DefTaxCode
        gTaxCodePY = DefTaxCode
        gTaxCodePT = DefTaxCode
        gLatePayIntAcCode = DefAcCode
        gBillDbNtAcCode = DefAcCode
        gServiceTaxAcCode = DefAcCode
        gTDSAcCode = DefAcCode
        gBankCommAcCode = DefAcCode
        gBrokerageAcCode = DefAcCode
        gAcCodePY = DefAcCode
        gAcCodePT = DefAcCode
        gAcCodeSY = DefAcCode
        gAcCodeST = DefAcCode
        gExpensesAcCode = DefAcCode
        gMillBillDbNtAcCode = DefAcCode
        gTDSPaidAcCode = DefAcCode
        '---TCS
        gAcCodeTcsPay = DefAcCode
        gAcCodeTcsRec = DefAcCode
    End If
RsS.Close
Exit Sub
ErrorRoutine:
    If Err.Number = 3265 Then '- Setting No field
        MsgBox Err.Description, vbInformation + vbOKOnly, App.Title
        Resume Next
    Else
        GProcErrorHandler
    End If
End Sub
'--- Function Return Check for Number Formatted String From Given RecodSet Field No As Per Data Type Of Given Txt Value
'-- UCaseCtrlNo - UCase TxtCtrl No Array
'--- NCaseCtrlNo - No Case TxtCtrl No Array
Private Function GetChkFormatStrFromRSField(rstbl As Recordset, FldNo As Long, TxtValue As Variant, Optional UCaseCtrlNo, Optional NCaseCtrlNo)
Dim IsProperCase  As Boolean
Dim LIndex As Long
Dim UIndex As Long
Dim CLIndex As Long
Dim CUIndex As Long
Dim i As Long
Dim Str
UIndex = -1
CUIndex = -1
If IsMissing(UCaseCtrlNo) = False Then
    LIndex = LBound(UCaseCtrlNo)
    UIndex = UBound(UCaseCtrlNo)
End If
If IsMissing(NCaseCtrlNo) = False Then
    CLIndex = LBound(NCaseCtrlNo)
    CUIndex = UBound(NCaseCtrlNo)
End If
Select Case rstbl.Fields(FldNo).Type
    Case adInteger, 2, 3
        If IsNumeric(TxtValue) = True Then
            Str = CLng(TxtValue)
        Else
            Str = "0"
        End If
    Case adDouble, 131, adDecimal
        If IsNumeric(TxtValue) = True Then
            Str = Format(TxtValue, GProcNumberFormat(rstbl.Fields(FldNo).Precision, rstbl.Fields(FldNo).NumericScale))
        Else
            Str = Format(0, GProcNumberFormat(rstbl.Fields(FldNo).Precision, rstbl.Fields(FldNo).NumericScale))
        End If
    Case adChar, adVarWChar, 200
        IsProperCase = True
        '-- case Sensitive
        If CUIndex >= 0 Then
            For i = CLIndex To CUIndex
                If NCaseCtrlNo(i) = FldNo Then
                    '.Text = Trim(.Text)
                    IsProperCase = False
                    Exit For
                End If
            Next i
        End If
        '--- Proper
        If IsProperCase = True Then
            Str = GProcProperCase(Trim(TxtValue))
        Else
            Str = TxtValue
        End If
        '--- Upper Case
        If UIndex >= 0 Then
            For i = LIndex To UIndex
                If UCaseCtrlNo(i) = FldNo Then
                    Str = GProcUpperCase(Trim(TxtValue))
                    Exit For
                End If
            Next i
        End If
    Case adDate
        If IsDate(TxtValue) = True Then
            Str = TxtValue
        Else
            Str = ""
        End If
End Select
GetChkFormatStrFromRSField = Str
End Function
'---- GProcCheckForNumber
'-- pUCaseCtrlNo - UCase TxtCtrl No Array
'--- pNCaseCtrlNo - No Case TxtCtrl No Array
Public Sub GProcCheckForNumber(pFrm As Form, pRstbl As Recordset, pMaxno As Long, Optional pUCaseCtrlNo, Optional pNCaseCtrlNo, Optional pOtherTxtFromNo As Long, Optional pOtherTxtToNo As Long)
'On Error GoTo ErrorHandler
Dim pTxtIndex As Long
'-- 0 to MaxNo
For pTxtIndex = 0 To pMaxno 'pRstbl.Fields.Count - 1
     With pFrm.txt(pTxtIndex)
        .text = GetChkFormatStrFromRSField(pRstbl, pTxtIndex, .text, pUCaseCtrlNo, pNCaseCtrlNo)
     End With
Next
'----- (Check for number Other Text Controls) pOtherTxtFromNo To pOtherTxtToNo
If pOtherTxtFromNo > pMaxno And pOtherTxtFromNo <= pOtherTxtToNo Then
    For pTxtIndex = pOtherTxtFromNo To pOtherTxtToNo
       With pFrm.txt(pTxtIndex)
            .text = GetChkFormatStrFromRSField(pRstbl, pTxtIndex, .text, pUCaseCtrlNo, pNCaseCtrlNo)
        End With
    Next pTxtIndex
End If
Exit Sub
ErrorHandler:
If Err.Number = 340 Or Err.Number = 3265 Then  '-- Control Not Exist
    Resume Next
Else
    MsgBox Err.Description
End If
End Sub
'--- Convert To Proper Case
Public Function GProcProperCase(Str As String) As String
GProcProperCase = StrConv(Trim(Str), vbProperCase)
End Function
'--- Convert To Upper Case
Public Function GProcUpperCase(Str As String) As String
GProcUpperCase = StrConv(Trim(Str), vbUpperCase)
End Function
'--- Function Return Save Formatted String From Given Txt RecodSet Field No As Per Data Type
Private Function GetSaveFormatStrFromTxtRSField(rstbl As Recordset, FldNo As Long, TxtValue As Variant)
Dim Str
Select Case rstbl.Fields(FldNo).Type
Case adInteger, 2, 3
    If IsNumeric(TxtValue) = True Then
        Str = CLng(TxtValue)
    Else
        Str = "0"
    End If
Case adNumeric
    If IsNumeric(TxtValue) = True Then
        Str = Format(TxtValue, GProcNumberFormat(rstbl.Fields(FldNo).Precision, rstbl.Fields(FldNo).NumericScale))
    Else
        Str = "0.00"
    End If
Case adDecimal, 131
    If IsNumeric(TxtValue) = True Then
        Str = Format(TxtValue, GProcNumberFormat(rstbl.Fields(FldNo).Precision, rstbl.Fields(FldNo).NumericScale))
    Else
        Str = "0"
    End If
Case adDouble
    If IsNumeric(TxtValue) = True Then
        Str = Format(TxtValue, GProcNumberFormat(rstbl.Fields(FldNo).Precision, 2))
    Else
        Str = "0"
    End If
Case adChar, adVarWChar, 200
    If Trim(TxtValue) = "" Then
        Str = Null
    Else
        Str = Trim(TxtValue)
    End If
Case adDate, 135
    If IsDate(TxtValue) = True Then
        Str = CDate(TxtValue)
    Else
        Str = Null
    End If
Case Else
    MsgBox "Field Type " & rstbl.Fields(FldNo).Type & " Not Found"
End Select
GetSaveFormatStrFromTxtRSField = Str
End Function
'----- Save Record
Public Sub GProcSaveRecord(pFrm As Form, pRstbl As Recordset, PDataAction As Integer, pMaxno As Long, Optional pOtherTxtFromNo As Long, Optional pOtherTxtToNo As Long)
On Error GoTo ErrorHandler
Dim i As Long
'If IsMissing(pFrmTxtCtrlName) Then Set pFrmTxtCtrlName = txt
'gCn.BeginTrans
If PDataAction = vbDataActionAddNew Then pRstbl.AddNew
    '--- Save For 0 To MaxNo
    For i = 0 To pMaxno 'pRstbl.Fields.Count - 1
        pRstbl.Fields(i) = GetSaveFormatStrFromTxtRSField(pRstbl, i, pFrm.txt(i).text)
    Next i
    '----- (Save Other Text Controls) pOtherTxtFromNo To pOtherTxtToNo
    If pOtherTxtFromNo > pMaxno And pOtherTxtFromNo <= pOtherTxtToNo Then
        For i = pOtherTxtFromNo To pOtherTxtToNo
            pRstbl.Fields(i) = GetSaveFormatStrFromTxtRSField(pRstbl, i, pFrm.txt(i).text)
        Next i
    End If
    pRstbl.Update
'gCn.CommitTrans
Exit Sub
ErrorHandler:
If Err.Number = 340 Or Err.Number = 3265 Then '-- Control Not Exist
    Resume Next
Else
    MsgBox Err.Description
End If
End Sub
'----- Character Only
Public Function GProcCharacterOnly(pKeyCode As Integer) As Integer
 If (pKeyCode >= 97 And pKeyCode <= 122) Or (pKeyCode >= 65 And pKeyCode <= 90) Or pKeyCode = 8 Then
    GProcCharacterOnly = pKeyCode
 Else
    GProcCharacterOnly = 0
 End If
End Function

'----- Validate Key
Public Function GProcValidateKey(pRstbl As Recordset, pIndex As Integer, pKeyAscii As Integer, pTxtText As String, Optional pISOnlyChar As Boolean) As Integer
Select Case pRstbl.Fields(pIndex).Type
    Case adInteger
        GProcValidateKey = GProcNumberOnly(pKeyAscii)
    Case adDouble, 2, 3, 131
        GProcValidateKey = GProcNumberWithDecimal(pKeyAscii, pTxtText)
    Case Else
        If pISOnlyChar = True Then
            GProcValidateKey = GProcNumberOnly(pKeyAscii)
        Else
            GProcValidateKey = pKeyAscii
        End If
End Select

End Function
'----- Set Main From Manu Variable
Public Sub GProcfrmMainSelectedMenu()

End Sub
'---- Set given control Text Selected
Public Sub GProcSelectBox(pControl As Control)
On Error Resume Next
pControl.SelStart = 0
pControl.SelLength = Len(pControl)
End Sub
'--- Number Only
Public Function GProcNumberOnly(pKeyCode As Integer) As Integer
 If pKeyCode >= 48 And pKeyCode <= 58 Or pKeyCode = 46 Or pKeyCode = 8 Then
    GProcNumberOnly = pKeyCode
 Else
    GProcNumberOnly = 0
 End If
End Function
'--- Number With Decimal
Public Function GProcNumberWithDecimal(pKey As Integer, PStr As String)
If PStr <> "0.00" Then
       If (pKey < 48 Or pKey > 57) Then
              If pKey = 8 Or pKey = 46 Or InStr(PStr, ".") <> 0 Then 'backspace
              Else
                 pKey = 0
              End If
        End If
End If
GProcNumberWithDecimal = pKey
End Function
'--- Function Return MaxLength From Given RecodSet Field No As Per Data Type
Private Function GetMaxLengthFromRSField(rstbl As Recordset, FldNo As Long) As Long
Dim MaxLen As Long
Select Case rstbl.Fields(FldNo).Type
    Case adInteger, 2, 3
        MaxLen = rstbl.Fields(FldNo).Precision
    Case adDouble, 131
        MaxLen = rstbl.Fields(FldNo).Precision
    Case adNumeric
        MaxLen = rstbl.Fields(FldNo).Precision
    Case adChar, adVarWChar, 200
        MaxLen = rstbl.Fields(FldNo).DefinedSize
    Case adDate
        MaxLen = 10
End Select
GetMaxLengthFromRSField = MaxLen
End Function
'--- Function Return Clear Formatted String From Given RecodSet Field No As Per Data Type
Private Function GetClearFormatStrFromRSField(rstbl As Recordset, FldNo As Long)
Dim Str
Select Case rstbl.Fields(FldNo).Type
    Case adInteger, 2, 3
        Str = "0"
    Case adDouble, 131
        Str = Format(0, GProcNumberFormat(rstbl.Fields(FldNo).Precision, rstbl.Fields(FldNo).NumericScale))
    Case adNumeric
        Str = Format(0, GProcNumberFormat(rstbl.Fields(FldNo).Precision, rstbl.Fields(FldNo).NumericScale))
    Case adChar, adVarWChar, 200
        Str = ""
    Case adDate, 135
        Str = "__/__/____"
End Select
GetClearFormatStrFromRSField = Str
End Function
'====  To Clear Contains Of all Controls
Public Sub GProcClearForm(pFrm As Form, pRstbl As Recordset, pMaxno As Long, Optional pIsFrmTxtCtrl As Boolean, Optional pOtherTxtFromNo As Long, Optional pOtherTxtToNo As Long)
'To Clear Contains Of Text Controls
On Error GoTo ErrorHandler
Dim i As Long
Dim Ctr
If pIsFrmTxtCtrl = True Then
    For i = 0 To pMaxno 'pRstbl.Fields.Count - 1
        With pFrm.txt(i)
            .text = GetClearFormatStrFromRSField(pRstbl, i)
            .MaxLength = GetMaxLengthFromRSField(pRstbl, i)
        End With
    Next i
    '----- (Clear Other Text Controls) pOtherTxtFromNo To pOtherTxtToNo
    If pOtherTxtFromNo > pMaxno And pOtherTxtFromNo <= pOtherTxtToNo Then
        For i = pOtherTxtFromNo To pOtherTxtToNo
           With pFrm.txt(i)
                .text = GetClearFormatStrFromRSField(pRstbl, i)
                .MaxLength = GetMaxLengthFromRSField(pRstbl, i)
            End With
        Next i
    End If
End If
'To Clear Contains Of Temp Controls
For Each Ctr In pFrm
    If TypeOf Ctr Is TextBox Then
        If Ctr.Name <> "txt" Then Ctr.text = ""
    ElseIf TypeOf Ctr Is ComboBox Then
        If Ctr.Style = 2 Then
            Ctr.ListIndex = 0
        Else
            Ctr.text = ""
        End If
        If Ctr.ListCount > 0 Then
            Ctr.ListIndex = 0
           ' Ctr.SelText = ""
        End If
    ElseIf TypeOf Ctr Is MaskEdBox Then
        Ctr.text = "__/__/____"
    'ElseIf TypeOf Ctr Is MSFlexGrid Then
    ' Ctr.Clear
    ElseIf TypeOf Ctr Is CheckBox Then
        Ctr.Value = 0
    ElseIf TypeOf Ctr Is DTPicker Then
        Ctr.Value = gCYSDate
    ElseIf TypeOf Ctr Is TreeView Then
        Ctr.Nodes.Clear
    End If
Next Ctr
Exit Sub
ErrorHandler:
    If Err.Number = 340 Or Err.Number = 3265 Then '-- Control Not Exist
        Resume Next
    Else
        MsgBox Err.Description
    End If
End Sub
Public Function GProcNumberFormat(pPrecision As Long, pScale As Long) As String
Dim FStr As String
Dim j As Integer
'If Prec = 13 Then
    'GProcNumberFormat = "#############"
 '   Exit Function
'End If
FStr = ""
    For j = 1 To pPrecision - 4
        FStr = FStr & "#"
    Next j
    FStr = FStr & "0"
    If pScale = 0 Then
    Else
        FStr = FStr & "."
        For j = 1 To pScale
        FStr = FStr & "0"
        Next j
    End If
    GProcNumberFormat = FStr
End Function
'------ Error Handler
Public Sub GProcErrorHandler()
    MsgBox Err.Description, vbInformation + vbOKOnly, App.Title
End Sub
'-------- Generate Invoice No
Public Function GProcGenerateInvNo(pVtype As String, pAcCode As Long, Optional pCriteria As String) As String
    Dim Y As String
    Dim G_Rs As Recordset
    Dim BillSrNo As String
    BillSrNo = GProcGetColumnValue("TblMastAccount", "AcCode", CStr(pAcCode), "N", "AcBillSrNo", "S")
    If gBackEndDB = gBackEndAccess Then
        If Trim(BillSrNo) = "" Or Len(BillSrNo) <> 3 Then
            Y = "Select Max(clng(AdInVNo)) from tblAddLess where VYear=" & gCYear & " and Vtype='" & pVtype & "' and isNumeric(AdInVNo) and AdAcCrCode= " & pAcCode & IIf(Trim(pCriteria) = "", "", " and " & pCriteria)
        Else
            Y = "Select Max(clng(mid(AdInVNo,5))) from tblAddLess where VYear=" & gCYear & " and Vtype='" & pVtype & "' and isNumeric(AdInVNo)=false and  AdAcCrCode= " & pAcCode & IIf(Trim(pCriteria) = "", "", " and " & pCriteria)
        End If
    Else '-- Oracle
        If Trim(BillSrNo) = "" Or Len(BillSrNo) <> 3 Then
            Y = "Select Max(to_number(AdInVNo)) from tblAddLess where VYear=" & gCYear & " and Vtype='" & pVtype & "'  and  AdAcCrCode= " & pAcCode & IIf(Trim(pCriteria) = "", "", " and " & pCriteria)
        Else
            Y = "Select Max(to_number(substr(AdInVNo,5))) from tblAddLess where VYear=" & gCYear & " and Vtype='" & pVtype & "' and AdAcCrCode= " & pAcCode & IIf(Trim(pCriteria) = "", "", " and " & pCriteria)
        End If
    End If
    GProcRstOpen G_Rs, Y, "R", gCn
    If G_Rs.Fields(0) = Null Or IsNumeric(G_Rs.Fields(0)) = False Then
        GProcGenerateInvNo = 1
    Else
        GProcGenerateInvNo = CLng(G_Rs.Fields(0)) + 1
    End If
    If Trim(BillSrNo) <> "" And Len(BillSrNo) = 3 Then GProcGenerateInvNo = BillSrNo & "-" & GProcGenerateInvNo
    G_Rs.Close
End Function
'-------- Generate Id
Public Function GProcGenerateId(pTblName As String, pFieldname As String, Optional pCriteria As String, Optional pConnection As Connection, Optional pStartCode As Long) As Long
    Dim Y As String
    Dim G_Rs As Recordset
    If pConnection Is Nothing Then
        Set pConnection = gCn
    End If
    Y = "Select Max( " & pFieldname & ") from " & pTblName & IIf(Trim(pCriteria) = "", "", " where " & pCriteria)
    GProcRstOpen G_Rs, Y, "R", pConnection
    If G_Rs.Fields(0) = Null Or IsNumeric(G_Rs.Fields(0)) = False Then
        GProcGenerateId = 1
    Else
        GProcGenerateId = CLng(G_Rs.Fields(0)) + 1
    End If
    If pStartCode > 0 And GProcGenerateId < pStartCode Then GProcGenerateId = pStartCode
    G_Rs.Close
End Function
'-------- Generate Date
Public Function GProcGenerateVDt(pTblName As String, pFieldname As String, Optional Criteria As String, Optional pConnection As Connection) As Date
    Dim Y As String
    Dim G_Rs As Recordset
    If pConnection Is Nothing Then
        Set pConnection = gCn
    End If
    Y = "Select Max( " & pFieldname & ") from " & pTblName & IIf(Trim(Criteria) = "", "", " where " & Criteria)
    GProcRstOpen G_Rs, Y, "R", pConnection
    If G_Rs.Fields(0) = Null Or IsDate(G_Rs.Fields(0)) = False Then
        GProcGenerateVDt = gCYSDate
    Else
        GProcGenerateVDt = Format(G_Rs.Fields(0), "dd/MM/yyyy")
    End If
    G_Rs.Close
End Function
'---- To Set InActive controls Back Color
Public Sub GProcSetInActivateCtrlBackClr(pForm As Form)
On Error Resume Next
    Dim lIntControl As Control
    For Each lIntControl In pForm
        If TypeOf lIntControl Is TextBox And lIntControl.Enabled = False Then lIntControl.BackColor = vbButtonFace
    Next lIntControl
End Sub
'---- To  Disable And Enable controls as Per Actions
Public Sub GProcActivateControls(pForm As Form, pCtrlEnable As Boolean)
On Error Resume Next
    Dim lIntControl As Control
    For Each lIntControl In pForm
        If TypeOf lIntControl Is TextBox Then
            lIntControl.Enabled = pCtrlEnable
        ElseIf TypeOf lIntControl Is ComboBox Then
            lIntControl.Enabled = pCtrlEnable
        ElseIf TypeOf lIntControl Is ListBox Then
           lIntControl.Enabled = pCtrlEnable
        ElseIf TypeOf lIntControl Is DTPicker Then
           lIntControl.Enabled = pCtrlEnable
        ElseIf TypeOf lIntControl Is MaskEdBox Then
           lIntControl.Enabled = pCtrlEnable
        ElseIf TypeOf lIntControl Is MSFlexGrid Then
            lIntControl.Enabled = pCtrlEnable
        ElseIf TypeOf lIntControl Is OptionButton Then
            lIntControl.Enabled = pCtrlEnable
        ElseIf TypeOf lIntControl Is CheckBox Then
            lIntControl.Enabled = pCtrlEnable
        ElseIf TypeOf lIntControl Is TreeView Then
            lIntControl.Enabled = pCtrlEnable
        End If
    Next lIntControl
End Sub
'-------------------- As Per User Previlage Set Button Visiblity
Public Sub GProcSetButtonVisiblity(pForm As Form)
Dim AddStatus As Long
Dim ModifyStatus As Long
Dim DeleteStatus As Long
Dim ViewStatus As Long
Dim i As Long
Dim TempMenuStr As String
AddStatus = 1
ModifyStatus = 1
DeleteStatus = 1
ViewStatus = 1
If gfrmMainMenu Is Nothing Then Exit Sub
    
TempMenuStr = Left(LCase(gfrmMainMenu.Name), 6)
If gIsAdminUser = True Then Exit Sub
If (TempMenuStr = "mnumst" Or TempMenuStr = "mnutrn" Or TempMenuStr = "mnuutl") And Right(LCase(gfrmMainMenu.Name), 3) = "frm" Then
    '--- Get Status from MenuName,MenuIndex,UserName
    AddStatus = GProcGetColumnValue("tblUserPrevilage", "MenuName", gfrmMainMenu.Name, "S", "Add", "N", " Add=1 and UserName='" & gUserName & "' and menuIndex=" & gfrmMainMenu.Index, gSetupCn)
    ModifyStatus = GProcGetColumnValue("tblUserPrevilage", "MenuName", gfrmMainMenu.Name, "S", "Modify", "N", " Modify=1 and UserName='" & gUserName & "' and menuIndex=" & gfrmMainMenu.Index, gSetupCn)
    DeleteStatus = GProcGetColumnValue("tblUserPrevilage", "MenuName", gfrmMainMenu.Name, "S", "Delete", "N", " Delete=1 and UserName='" & gUserName & "' and menuIndex=" & gfrmMainMenu.Index, gSetupCn)
    ViewStatus = GProcGetColumnValue("tblUserPrevilage", "MenuName", gfrmMainMenu.Name, "S", "View", "N", " View=1 and UserName='" & gUserName & "' and menuIndex=" & gfrmMainMenu.Index, gSetupCn)
    
    If ModifyStatus = 1 Or DeleteStatus = 1 Then ViewStatus = 1
    
    '-- Add
    If AddStatus = 0 Then pForm.cmdBtn(6).Visible = False
    '-- Modify
    If ModifyStatus = 0 Then pForm.cmdBtn(7).Visible = False
    '-- Delete
    If DeleteStatus = 0 Then pForm.cmdBtn(8).Visible = False
    '-- View
    If ViewStatus = 0 Then
        For i = 0 To 4
            pForm.cmdBtn(i).Visible = False  '--Navigation,Find
        Next
    End If
    '-- Save
    If AddStatus = 0 And ModifyStatus = 0 Then pForm.cmdBtn(13).Visible = False
    '-- Cancel
   ' If AddStatus = 0 And ModifyStatus = 0 Then pForm.cmdBtn(14).Visible = False
End If
End Sub
Public Sub GProcChkPass(pType As String, PDataAction As Integer, PVno As String, pVdt As String, pAcName As String, pUser As String, pHank As String)
    Dim Line1 As String
    Dim Line2 As String
    Dim Line3 As String
    Dim Line4 As String
    Dim Line5 As String
    Line1 = ""   '---- Record Deleted Or Modified
    Line2 = ""   '---- Record Type
    Line3 = ""   '---- Record Ref. No & Date
    Line4 = ""   '---- Ref. A/c
    Line5 = ""   '---- Firm
    
    pType = pType + pHank
    
'    gChkPassTrueFalse = False
    frmUtlCheckPassword.Show 1
    
'   If gChkPassTrueFalse = True Then
        
        '--- Line 1
        If PDataAction = 7 Then
           Line1 = "Record Deleted"
        Else
           Line1 = "Record Modified"
        End If
        
        '--- Line 2
        Select Case pType
            Case "SY"
                Line2 = "Trade Sale Bill"
            Case "SO"
                Line2 = "Consignment Sale Bill"
            Case "SD"
                Line2 = "Depot Sale Bill"
            Case "ST"
                Line2 = "SIT Sale Bill"
            Case "SM"
                Line2 = "MILL Bill"
            Case "SYHNK"
                Line2 = "Hank Sale Bill"
            Case "RY"
                Line2 = "Trade Sale Return"
            Case "RO"
                Line2 = "Consign/Depot Sale Return"
            Case "PY"
                Line2 = "Trade Purch."
            Case "PT"
                Line2 = "SIT Purch."
            Case "PO"
                Line2 = "Other Purch."
            Case "PI"
                Line2 = "Googds Inward"
            Case "PYHNK"
                Line2 = "Hank Purch."
            Case "VY"
                Line2 = "Trade Purch. Return"
            Case "VI"
                Line2 = "Consign./Depot Inw. Return"
            Case "GP"
                Line2 = "Gate Pass"
            Case "OM"
                Line2 = "Mill Bill Booking"
            Case "OS"
                Line2 = "SIT Purch. Booking"
            Case "OI"
                Line2 = "Consign./Depot Purch. Booking"
            Case "OP"
                Line2 = "Party Booking"
            Case "CP"
                Line2 = "Cash Payment"
            Case "BP"
                Line2 = "Bank Payment"
            Case "CR"
                Line2 = "Cash Receipt"
            Case "BR"
                Line2 = "Bank Receipt"
            Case "MR"
                Line2 = "Mill Bill Receipt"
            Case "JV"
                Line2 = "J V"
            Case "PN"
                Line2 = "Cr. Note"
            Case "SN"
                Line2 = "DCr. Note"
        End Select
        '--- Line 3
        Line3 = "No-" + PVno + " Dt-" + pVdt
        '--- Line 4
        If pAcName <> "" Then
           Line4 = "Of-" + pAcName
        End If
        '--- Line 5
        Line5 = "Firm-" + Left(gCName, 10)

'-----------For SMS
            Dim HttpReq As New WinHttpRequest
            Dim URL As String
            Dim params As String
            Dim params2 As String
            Dim SmsMblNo As String
            Dim SmsString As String

            SmsMblNo = "9665748022"
            SmsString = Line1 & vbLf
            SmsString = SmsString + "By User-" + pUser & vbLf
            SmsString = SmsString & vbLf
            SmsString = SmsString + Line2 & vbLf
            SmsString = SmsString + Line3 & vbLf
            If pAcName <> "" Then
               SmsString = SmsString + Line4 & vbLf
            End If
            SmsString = SmsString & vbLf
            SmsString = SmsString + Line5 & vbLf
            If SmsMblNo = "" Then
               MsgBox ("Party Mobile No Not Found....")
            Else
                        URL = "http://alerts.sinfini.com/api/web2sms.php"
                        params = "workingkey= Ad583e28d45e23d78611667b060ed2e40&sender=MCROCO&to=" & SmsMblNo & " &message=" & SmsString & ""
                        HttpReq.Open "POST", URL, False
                        HttpReq.SetRequestHeader "Content-Type", "application/x-www-form-urlencoded"
                        HttpReq.Send params
            End If
            DoEvents
'------End SMS
 '   End If
End Sub






Public Sub GProcSetButtons(pForm As Form, PDataAction As Integer)
    Dim i As Long
    DoEvents
    Select Case PDataAction
        Case vbDataActionAddNew, vbDataActionUpdate
            If PDataAction = vbDataActionAddNew Then
              pForm.lblActionStatus.Caption = "New Record..."
            Else
              pForm.lblActionStatus.Caption = "Modify Record..."
            End If
            For i = 0 To 12
               pForm.cmdBtn(i).Enabled = False
            Next i
            pForm.cmdBtn(10).Enabled = True
            If PDataAction = vbDataActionUpdate Then
                pForm.cmdBtn(13).Enabled = True
                pForm.cmdBtn(13).BackColor = &HFFFF80
            Else
                pForm.cmdBtn(13).Enabled = False 'True
                pForm.cmdBtn(13).BackColor = vbButtonFace
            End If
            pForm.cmdBtn(14).Enabled = True
            pForm.cmdBtn(14).BackColor = &HFFFF80
        Case vbDataActionDelete
            pForm.lblActionStatus.Caption = "Delete Record..."
            For i = 0 To 13
               pForm.cmdBtn(i).Enabled = False
            Next i
            pForm.cmdBtn(14).Enabled = True
            pForm.cmdBtn(10).Enabled = True
            pForm.cmdBtn(13).BackColor = vbButtonFace
            pForm.cmdBtn(14).BackColor = vbButtonFace
         Case vbDataActionCancel
            For i = 0 To 14
               pForm.cmdBtn(i).Enabled = False
            Next i
            pForm.cmdBtn(6).Enabled = True
            pForm.cmdBtn(10).Enabled = True
            pForm.cmdBtn(13).BackColor = vbButtonFace
            pForm.cmdBtn(14).BackColor = vbButtonFace
         Case Else
            pForm.lblActionStatus.Caption = "View Record..."
            For i = 0 To 12
               pForm.cmdBtn(i).Enabled = True
            Next i
            pForm.cmdBtn(14).Enabled = False
            pForm.cmdBtn(13).Enabled = False
            pForm.cmdBtn(13).BackColor = vbButtonFace
            pForm.cmdBtn(14).BackColor = vbButtonFace
    End Select
    DoEvents
End Sub
Public Function GProcNoToStr(xnum As Double, ylen As Integer, zdes As Integer) As String
Dim X As String, Y As String, z As String
If (InStr(1, Str(xnum), ".") > 0) Then
    If (Len(Str(xnum)) - InStr(1, Str(xnum), ".") > zdes + 2) Then
        xnum = Val(Left(Str(xnum), InStr(1, Str(xnum), ".") + zdes + 2))
    End If
End If
If (xnum >= 0) Then
   X = Str(Int(xnum))
Else
   X = "-" + LTrim(Str(Int(Abs(xnum))))
End If
    If (zdes > 0) Then
        Y = LTrim(Str(Round((Abs(xnum) - Int(Abs(xnum))), zdes + 1)))
        X = Right(Space(ylen - zdes - 1) + X, ylen - zdes - 1)
        If Not (Left(Y, 1) = ".") Then
               Y = "." + Y
        End If
        Y = Left(Y + "00000000000", zdes + 1)
    Else
        Y = ""
        X = Right(Space(ylen) + X, ylen)
    End If
'If (zdelimtr <> "") Then
'   ylen = Len(X)
'   z = Right(X, 3)
'   X = Left(X, ylen - 3)
'   Do While Len(Trim(X)) > 0
'      z = Right(X, 2) + zdelimtr + z
'      X = Left(X, Len(X) - 2)
'   Loop
'   X = Right(Space(ylen) + z, ylen)
'End If
GProcNoToStr = X + Y
End Function
Public Function GProcNumToWords(pAmount As Double)
Dim Rs As Long
Dim wNum(1 To 5) As Integer
Dim wNumInWord(1 To 5) As String
Dim onetonine, TenToHund, EleToNinteen
Dim div As Integer, Ps As Integer
Dim i As Integer
Dim Rsword As String, PsWord As String, AmtInWords As String
onetonine = Array("", "One ", "Two ", "Three ", "Four ", "Five ", "Six ", "Seven ", "Eight ", "Nine ")
TenToHund = Array("", "Ten ", "Twenty ", "Thirty ", "Fourty ", "Fifty ", "Sixty ", "Seventy ", "Eighty ", "Ninety ")
EleToNinteen = Array("", "", "", "", "", "", "", "", "", "", "", "Eleven ", "Twelve ", "Thirteen ", "Fourteen ", "Fifteen ", "Sixteen ", "Seventeen ", "Eighteen ", "Nineteen ")

Rs = Int(pAmount)
Ps = (pAmount - Rs) * 100

For i = 1 To 5
   If i = 2 Then
        div = 10
   Else
        div = 100
   End If
   wNum(i) = Rs Mod div

   If wNum(i) = 0 Then
      wNumInWord(i) = ""
   ElseIf wNum(i) > 0 And wNum(i) < 10 Then
      wNumInWord(i) = onetonine(wNum(i))
   ElseIf wNum(i) Mod 10 = 0 Then
      wNumInWord(i) = TenToHund(wNum(i) / 10)
   ElseIf wNum(i) > 10 And wNum(i) < 20 Then
      wNumInWord(i) = EleToNinteen(wNum(i))
   Else
      wNumInWord(i) = TenToHund(Int(wNum(i) / 10)) & onetonine(wNum(i) Mod 10)
   End If

   Select Case i
    Case 1
        wNumInWord(i) = wNumInWord(i) & ""
    Case 2
        If wNumInWord(i) <> "" Then wNumInWord(i) = wNumInWord(i) & "Hundred "
    Case 3
         If wNumInWord(i) <> "" Then wNumInWord(i) = wNumInWord(i) & "Thousand "
    Case 4
         If wNumInWord(i) <> "" Then wNumInWord(i) = wNumInWord(i) & "Lakh "
    Case 5
         If wNumInWord(i) <> "" Then wNumInWord(i) = wNumInWord(i) & "Crore "
   End Select
   Rs = Int(Rs / div)
   Rsword = wNumInWord(i) & Rsword
Next

If Ps = 0 Then
    PsWord = ""
ElseIf Ps > 0 And Ps < 10 Then
    PsWord = onetonine(Ps)
ElseIf Ps Mod 10 = 0 Then
    PsWord = TenToHund(Ps / 10)
ElseIf Ps > 10 And Ps < 20 Then
    PsWord = EleToNinteen(Ps)
Else
    PsWord = TenToHund(Int(Ps / 10)) & onetonine(Ps Mod 10)
End If

AmtInWords = Rsword & IIf(PsWord = "", "", "and " & PsWord & " Ps ")
If AmtInWords <> "" Then AmtInWords = AmtInWords & "Only."
GProcNumToWords = AmtInWords
End Function
'---- Show Form on F2
Public Sub GProcShowForm(ShowFormCtrl As Form, ShowFrmMenuName As Menu, Optional frmCtrls, Optional frmCtrlsValue)
Dim meCtrl() As Control
Dim X As String
Dim RsMenu As Recordset
Dim i As Long
Set gfrmMainMenu = ShowFrmMenuName
DoEvents
If gIsAdminUser = False Then
    X = "Select * from tblUserPrevilage where UserName='" & gUserName & "' and " _
    & " menuName='" & gfrmMainMenu.Name & "' and MenuIndex=" & CLng(gfrmMainMenu.Index) & " and Add=1"
    If GProcRstOpen(RsMenu, X, "O", gSetupCn) = 0 Then
        Unload ShowFormCtrl
        Exit Sub
    End If
End If
'Unload ShowFormCtrl
'DoEvents
ShowFormCtrl.Show
ShowFormCtrl.ZOrder 0
If IsArray(frmCtrls) = True Then
    For i = LBound(frmCtrls) To UBound(frmCtrls)
        ReDim Preserve meCtrl(i)
        Set meCtrl(i) = frmCtrls(i)
    Next i
    For i = 0 To UBound(meCtrl)
        meCtrl(i) = frmCtrlsValue(i)
    Next i
End If
DoEvents
End Sub

'--- Make Rounding
Public Function GProcMakeRounding(Amt As Double) As Double
Dim Fraction As Double
Fraction = Amt - Int(Amt)
If Fraction >= 0.5 Then
    GProcMakeRounding = Amt + (1 - Fraction)
Else
    GProcMakeRounding = Amt - Fraction
End If
End Function
'---- Is Date in Working Year
Public Function GProcIsDateValid(d As String, Optional Year As Long) As Boolean
GProcIsDateValid = False
d = Format(d, "dd/MM/yyyy")
If Year = 0 Then
    If IsDate(d) = True Then
        If CDate(d) >= gCYSDate And CDate(d) <= gCYEDate Then
            GProcIsDateValid = True
        End If
    End If
Else
    If IsDate(d) = True Then
        If CDate(d) >= CDate(Format("01/04/" & Year, "dd/mm/yyyy")) And CDate(d) <= CDate(Format("31/03/" & (Year + 1), "dd/mm/yyyy")) Then
            GProcIsDateValid = True
        End If
    End If
End If
End Function
'--- Get Main Group Code
Public Function GProcGetMainGroupCode(AGroupName As String) As Long
Dim X As String
Dim G_Rs As Recordset
X = "Select AgCode from tblMastGroup where AgName='" & AGroupName & "' and IsFixGroup=1 union all" _
& " Select GpCode from tblMastGroup where AgName='" & AGroupName & "' and IsFixGroup=0"
If GProcRstOpen(G_Rs, X, "R") > 0 Then
    GProcGetMainGroupCode = CLng(G_Rs.Fields(0))
End If
G_Rs.Close
End Function
'--- Create Default Master Record
Public Sub GProcCreateDefaultRecord()
Dim X As String
Dim Aid As Long
Dim Rs1 As Recordset
Set Rs1 = New Recordset
'-- Company Group
X = "select * from tblMastCompGroup where cgName='-'"
Set Rs1 = New Recordset
If GProcRstOpen(Rs1, X, "O") = 0 Then
    Rs1.AddNew
    Rs1.Fields("cgCode") = GProcGenerateId("tblMastCompGroup", "CgCode")
    Rs1.Fields("cgName") = "-"
    Rs1.Update
    Rs1.Close
End If
'-- Account
X = "select * from tblmastaccount where AcName='-'"
If GProcRstOpen(Rs1, X, "O") = 0 Then
    Rs1.AddNew
    Aid = GProcGenerateId("tblmastaccount", "AcCode")
    Rs1.Fields("AcCode") = Aid
    Rs1.Fields("AcName") = "-"
    Rs1.Fields("AgCode") = 90001
    Rs1.Fields("AgCode1") = 90001
    Rs1.Fields("AcCmpCode") = GProcGetColumnValue("tblmastcompgroup", "CGName", "-", "S", "CGCode", "N")
    Rs1.Fields("AcBrkCode") = Aid
    Rs1.Update
    Rs1.Close
End If
'--- Narration
X = "select * from tblMastNarration where Narration='-'"
Set Rs1 = New Recordset
If GProcRstOpen(Rs1, X, "O") = 0 Then
  Rs1.AddNew
  Rs1.Fields("NarrCode") = GProcGenerateId("tblMastNarration", "NarrCode")
  Rs1.Fields("NarrType") = "N"
  Rs1.Fields("Narration") = "-"
  Rs1.Fields("MastTaxPurAcCode") = GProcGetColumnValue("tblMastAccount", "AcName", "-", "s", "AcCode", "N")
  Rs1.Fields("MastTaxSAAcCode") = GProcGetColumnValue("tblMastAccount", "AcName", "-", "s", "AcCode", "N")
  Rs1.Update
  Rs1.Close
End If
End Sub
'---- Is Voucher Audited
Public Function GProcISAudited(PVno As Long, pVtype As String, pVYear As Long) As Boolean
Dim X As String
Dim Rs As Recordset
X = "Select * from tblVoucher where VIsAudited=1 and Vno=" & PVno & " and Vtype='" & pVtype & "' and Vyear=" & pVYear
If GProcRstOpen(Rs, X, "R") > 0 Then
    GProcISAudited = True
End If
Rs.Close
End Function
'--- Get Target field from Source Field
Public Function GProcGetColumnValue(pTblName As String, pSFieldName As String, pSValue As String, pSFieldType As String, pTFieldName As String, pTFieldType As String, Optional pCriteria As String, Optional pConnection As Connection)
Dim X As String
Dim G_Rs As Recordset

If pSFieldType = "N" Then
    X = "Select " & pTFieldName & " from " & pTblName & " where " & pSFieldName & "=" & pSValue & IIf(Trim(pCriteria) <> "", " and " & pCriteria, "")
Else
    X = "Select " & pTFieldName & " from " & pTblName & " where " & pSFieldName & "='" & pSValue & "' " & IIf(Trim(pCriteria) <> "", " and " & pCriteria, "")
End If
If pConnection Is Nothing Then
    Set pConnection = gCn
End If
If pTFieldType = "N" Then
    GProcGetColumnValue = 0
Else
    GProcGetColumnValue = ""
End If
If GProcRstOpen(G_Rs, X, "R", pConnection) > 0 Then
    If Not IsNull(G_Rs.Fields(0)) Then GProcGetColumnValue = G_Rs.Fields(0)
End If
G_Rs.Close
End Function
'--- Create OutStanding Record
Public Sub GProcCreateOutStanding(Vno As Long, Vdt As Date, Vtype As String, BillNo As String, AcCode As Long, BrkCode As Long, Amt As Double, Optional DueDays As Long, Optional DueDt, Optional MillCode As Long, Optional DbNtAmt As Double, Optional Bag As Long, Optional ItemCode As Long, Optional PurBlDt As Date)
Dim X As String
Dim rsOS As Recordset
X = "select * from tblOutStanding where vno=0"
If AcCode = 0 Then AcCode = GProcGetColumnValue("tblMastAccount", "AcName", "-", "S", "AcCode", "N")
If BrkCode = 0 Then BrkCode = GProcGetColumnValue("tblMastAccount", "AcName", "-", "S", "AcCode", "N")
If MillCode = 0 Then MillCode = GProcGetColumnValue("tblMastAccount", "AcName", "-", "S", "AcCode", "N")

GProcRstOpen rsOS, X, "O"
With rsOS
    .AddNew
    .Fields("Vno") = Vno
    .Fields("VDt") = Vdt
    .Fields("VType") = Vtype
    .Fields("vYear") = gCYear
    .Fields("Outbillno") = IIf(BillNo = "", "-", BillNo)
    .Fields("OutAcCode") = AcCode
    .Fields("OutBrokerCode") = BrkCode
    .Fields("OutMillCode") = MillCode
    .Fields("OutAmount") = Amt
    .Fields("OutDuedays") = DueDays
    .Fields("OutDuedate") = IIf(IsDate(DueDt), DueDt, "")
    .Fields("OutDbNtAmt") = CDbl(DbNtAmt)
    .Fields("OutBag") = Bag
    If ItemCode > 0 Then
        .Fields("OutItShort") = GProcGetColumnValue("TblMastItem", "ItCode", CStr(ItemCode), "N", "ITSHORT", "S") & vbNullString
        .Fields("OutTicket") = GProcGetColumnValue("TblMastItem", "ItCode", CStr(ItemCode), "N", "ITTICKET", "S") & vbNullString
        .Fields("OutItemCode") = ItemCode
    End If
'    .Fields("OutCrt") = CRt
'    .Fields("OutSrt") = SRt
'    .Fields("OutIrt") = IRt
    If Year(PurBlDt) > 1899 Then
      .Fields("OutPurBillDt") = PurBlDt
    End If
    .Update
End With
rsOS.Close
End Sub

'--- Create Voucher
Public Sub GProcCreateVoucher(frmAction As Integer, Vno As Long, VSrno As Long, Vdt As Date, Vtype As String, AcCode As Long, AcOCode As Long, Amt As Double, Optional BillNo As String, Optional BillDt, Optional RconDt, Optional Nar1 As String, Optional Nar2 As String, Optional Nar3 As String, Optional Nar4 As String, Optional IsRefEntryType As Boolean, Optional MillExpNarrCode As Long, Optional Taxable As Double, Optional Crt As Double, Optional Camt As Double, Optional Srt As Double, Optional Samt As Double, Optional Irt As Double, Optional Iamt As Double, Optional AgtVno As Long, Optional AgtVtype As String, Optional AgtVyear As Long, Optional GSTEnt As Long, Optional DrCrNtNo As String, Optional DrCrNtType As String, Optional Ramt As Double, Optional wBroker As Long, Optional wIrnNo As String, Optional HsnCd As String)
Dim X As String
Dim rsVou As Recordset
X = "select * from tblVoucher where vno=0"
GProcRstOpen rsVou, X, "O"
With rsVou
    .AddNew
    .Fields("Vno") = Vno
    .Fields("VDt") = Vdt
    .Fields("VType") = Vtype
    .Fields("vctrno") = VSrno
    .Fields("vAcCode") = AcCode
    .Fields("VAcOCode") = AcOCode
    .Fields("Vamt") = Amt
    .Fields("vbillno") = BillNo
    If IsDate(BillDt) = True Then
      .Fields("VBillDt") = CDate(BillDt)
    End If
    If IsDate(RconDt) = True Then
      .Fields("VReconDt") = CDate(RconDt)
    End If
    .Fields("VNar1") = Nar1 & vbNullString
    .Fields("VNar2") = Nar2 & vbNullString
    .Fields("VNar3") = Nar3 & vbNullString
    .Fields("VNar4") = Nar4 & vbNullString
    .Fields("vUserName") = gUserName
    .Fields("vEntStatus") = IIf(frmAction = 5, 0, 1) '-- Add 0,Modify 1
    .Fields("vYear") = gCYear
    .Fields("MillExpNarrCode") = MillExpNarrCode
    If IsRefEntryType = True Then .Fields("VIsRefEntType") = 1
    .Fields("VTaxableAmt") = Taxable
    .Fields("VCGSTRt") = Crt
    .Fields("VCGSTAmt") = Camt
    .Fields("VSGSTRt") = Srt
    .Fields("VSGSTAmt") = Samt
    .Fields("VIGSTRt") = Irt
    .Fields("VIGSTAmt") = Iamt
    .Fields("VAgtVno") = AgtVno
    .Fields("VAgtVtype") = UCase(AgtVtype)
    .Fields("VAgtVyear") = AgtVyear
    .Fields("VIsGSTEntry") = GSTEnt
    .Fields("VDrCrNtNo") = DrCrNtNo & vbNullString
    .Fields("VDrCrNtType") = DrCrNtType & vbNullString
    If Not IsNull(Ramt) Then
      .Fields("VRecAmt") = Ramt
    End If
    .Fields("Vbroker") = wBroker
    .Fields("VIrnNo") = wIrnNo & vbNullString
    If gCYear >= 2025 Then
       .Fields("VHsn") = HsnCd & vbNullString
    End If
    
    .Update
End With
rsVou.Close
End Sub
'--- Create Mill Rec Pay
Public Sub GProcCreateMillRecPay(frmAction As Integer, Vno As Long, VSrno As Long, Vdt As Date, Vtype As String, AcCode As Long, AcOCode As Long, Amt As Double, Optional BillNo As String, Optional BillDt, Optional RconDt, Optional Nar1 As String, Optional Nar2 As String, Optional Nar3 As String, Optional Nar4 As String, Optional MillExpNarrCode As Long)
Dim X As String
Dim rsVou As Recordset
X = "select * from TblMillRecPay where vno=0"
GProcRstOpen rsVou, X, "O"
With rsVou
    .AddNew
    .Fields("Vno") = Vno
    .Fields("VDt") = Vdt
    .Fields("VType") = Vtype
    .Fields("vctrno") = VSrno
    .Fields("MRecAcCode") = AcCode
    .Fields("MRecAcOCode") = AcOCode
    .Fields("MRecAmt") = Amt
    .Fields("MRecbillno") = BillNo
    If IsDate(BillDt) = True Then
      .Fields("MRecBillDt") = CDate(BillDt)
    End If
    If IsDate(RconDt) = True Then
      .Fields("MRecReconDt") = CDate(RconDt)
    End If
    .Fields("MRecNar1") = Nar1 & vbNullString
    .Fields("MRecNar2") = Nar2 & vbNullString
    .Fields("MRecNar3") = Nar3 & vbNullString
    .Fields("MRecNar4") = Nar4 & vbNullString
    .Fields("vUserName") = gUserName
    .Fields("vEntStatus") = IIf(frmAction = 5, 0, 1) '-- Add 0,Modify 1
    .Fields("vYear") = gCYear
    .Fields("MRecMillExpNarrCode") = MillExpNarrCode
    .Update
End With
rsVou.Close
End Sub
'----Set Crystal Properties For Report Preparation
Public Sub GProcCrystalRptPreparation(CryRptCtrl As CrystalReport, wDestination As String)
Dim i As Long
With CryRptCtrl
    .Reset
    .WindowLeft = 0
    .WindowTop = 0
    .MarginTop = 360
    .WindowState = crptMaximized
    .WindowTitle = gSelectedMenu
    .WindowAllowDrillDown = True
    .WindowShowPrintSetupBtn = True
    .WindowShowProgressCtls = True
    .WindowShowZoomCtl = True
    .WindowControls = True
    .WindowShowSearchBtn = True
    .WindowShowZoomCtl = True
    .WindowShowNavigationCtls = True
    .WindowShowRefreshBtn = True
    .ProgressDialog = True
    .DiscardSavedData = True
    .PrinterDriver = Printer.DriverName
    If gBackEndDB = gBackEndAccess Then
        .DataFiles(0) = gDbLocation
        gReportPath = App.Path & "\HiReports_Access\"
    Else '-- Oracle
       .Connect = "DSN=" + gODBCDSN + " ;UID=Scott;PWD=" + gODBCPwd + ";DBQ=" + gODBCHostStr
        gReportPath = App.Path & "\HiReports_ORA\"
    End If
    If wDestination = "Window" Then
        .Destination = crptToWindow
    Else
        .Destination = crptToPrinter
    End If
    For i = 0 To 35
        .Formulas(i) = ""
    Next i
    .Formulas(0) = "FirmNm = '" & UCase(gCName) & "'"
    .Formulas(1) = "wAdd1 = '" & gCAdd1 & " " & gCAdd2 & "'"
    .Formulas(2) = "wAdd2 = '" & gCAdd3 & "'"
    .Formulas(3) = "wPhNo = '" & gCPhNo & "'"
    .SelectionFormula = ""
End With
End Sub
'--------- Amount In Words
Public Function GProcAmountInWords(Amount)
Dim Rs As Long
Dim Paise As Integer
Dim digit(1 To 6) As Integer
Dim digitword(1 To 6) As String
Dim onetonine(1 To 9) As String
Dim tentohundred(1 To 9) As String
Dim eleventoninteen(11 To 19) As String
Dim Unit(1 To 5) As String
Dim div As Integer
Dim n As Integer
Dim Rsword, Paiseword, answer As String
If Amount < 0 Then
    GProcAmountInWords = "NIL"
    Exit Function
End If
onetonine(1) = " One "
onetonine(2) = " Two "
onetonine(3) = " Three "
onetonine(4) = " Four "
onetonine(5) = " Five "
onetonine(6) = " Six "
onetonine(7) = " Seven "
onetonine(8) = " Eight "
onetonine(9) = " Nine "

tentohundred(1) = " Ten "
tentohundred(2) = " Twenty "
tentohundred(3) = " Thirty "
tentohundred(4) = " Forty "
tentohundred(5) = " Fifty "
tentohundred(6) = " Sixty "
tentohundred(7) = " Seventy "
tentohundred(8) = " Eighty "
tentohundred(9) = " Ninety "

eleventoninteen(11) = " Eleven "
eleventoninteen(12) = " Twelve "
eleventoninteen(13) = " Thirteen "
eleventoninteen(14) = " Fourteen "
eleventoninteen(15) = " Fifteen "
eleventoninteen(16) = " Sixteen "
eleventoninteen(17) = " Seventeen "
eleventoninteen(18) = " Eighteen "
eleventoninteen(19) = " Nineteen "

Rs = Int(Amount)
Paise = (Amount - Rs) * 100

For n = 1 To 6
   If n = 2 Or n = 6 Then
        div = 10
   Else
        div = 100
   End If
   digit(n) = Rs Mod div
   If digit(n) = 0 Then
      digitword(n) = ""
   ElseIf digit(n) > 0 And digit(n) < 10 Then
      digitword(n) = onetonine(digit(n))
   ElseIf digit(n) Mod 10 = 0 Then
      digitword(n) = tentohundred(digit(n) / 10)
   ElseIf digit(n) > 10 And digit(n) < 20 Then
      digitword(n) = eleventoninteen(digit(n))
   Else
      digitword(n) = tentohundred(Int(digit(n) / 10)) & onetonine(digit(n) Mod 10)
   End If
   
   Select Case n
    Case 1
            digitword(n) = digitword(n) & ""
    Case 2
           If digitword(n) <> "" Then digitword(n) = digitword(n) & " Hundred "
    Case 3
         If digitword(n) <> "" Then digitword(n) = digitword(n) & " Thousand "
    Case 4
         If digitword(n) <> "" Then digitword(n) = digitword(n) & " Lac "
    Case 5
         If digitword(n) <> "" Then digitword(n) = digitword(n) & " Crore "
    Case 6
         If digitword(n) <> "" Then digitword(n) = digitword(n) & " Billion "
   End Select
      
   Rs = Int(Rs / div)
 
Next

For n = 6 To 1 Step -1
    Rsword = Rsword & digitword(n)
Next
If Paise = 0 Then
      Paiseword = ""
   ElseIf Paise > 0 And Paise < 10 Then
      Paiseword = onetonine(Paise)
   ElseIf Paise Mod 10 = 0 Then
      Paiseword = tentohundred(Paise / 10)
   ElseIf Paise > 10 And Paise < 20 Then
      Paiseword = eleventoninteen(Paise)
   Else
      Paiseword = tentohundred(Int(Paise / 10)) & onetonine(Paise Mod 10)
   End If
   
 answer = IIf(Rsword = "", "", " " & Rsword) & IIf(Paiseword = "", "", IIf(Rsword = "", "", "and ") & Paiseword & "Paise ")
If answer = "" Then
    answer = "NIL"
Else
    answer = answer & "only."
End If
GProcAmountInWords = "Rs. " & answer
End Function
'------------  (For All Firm) Insert into Setup Database  Table From Given Source Query
Public Sub GProcInsertToSetupDbTbl(AllFirmFlag As Boolean, SQry As String, TTblName As String, TTblFieldStr As String, TTblSelCriteria As String, TTblCompCodeField As String, Optional TTblCompNameField As String)
Dim TmpCn As Connection
Dim RstCompMast As Recordset
Dim RsS As Recordset
Dim RsT As Recordset
Dim i As Long
'-- Delete as per Criteria
gSetupCn.BeginTrans
gSetupCn.Execute "delete from " & TTblName & IIf(TTblSelCriteria <> "", " where " & TTblSelCriteria, "")
gSetupCn.CommitTrans
gSetupCn.BeginTrans
GProcRstOpen RsT, "Select " & TTblFieldStr & IIf(TTblCompCodeField <> "", " ," & TTblCompCodeField, "") & IIf(TTblCompNameField <> "", " ," & TTblCompNameField, "") & " from " & TTblName & IIf(TTblSelCriteria <> "", " where " & TTblSelCriteria, ""), "O", gSetupCn
'--- For Firm
    GProcRstOpen RstCompMast, "Select * from tblCompanyYear,tblMastCompany  where compCode=cCode " _
    & " and CompFdt=cdate('" & gCYSDate & "')" & IIf(AllFirmFlag = True, "", " and cName='" & gCName & "'"), "R", gSetupCn
    If RstCompMast.EOF = False Then RstCompMast.MoveFirst
    Do While Not RstCompMast.EOF  '----Temparary Connection
        If gBackEndDB = gBackEndAccess Then '-- Access
            GProcConnectionOpen TmpCn, RstCompMast!CompCode + RstCompMast!CompYear
        Else
            GProcConnectionOpenORA TmpCn, RstCompMast!CompCode + RstCompMast!CompYear
        End If
        If GProcRstOpen(RsS, SQry, "R", TmpCn) > 0 Then
            RsS.MoveFirst
            Do While Not RsS.EOF
                RsT.AddNew
                For i = 0 To RsS.Fields.Count - 1
                    Select Case RsS.Fields(i).Type
                        Case adDecimal, 131, adNumeric, adInteger, 2, 3, adDouble, 5
                            If IsNumeric(RsS.Fields(i)) = False Then
                                RsT.Fields(i) = 0
                            Else
                                RsT.Fields(i) = RsS.Fields(i)
                            End If
                        Case Else
                            RsT.Fields(i) = RsS.Fields(i) '& vbNullString
                    End Select
                Next i
                If TTblCompCodeField <> "" Then RsT.Fields(i) = RstCompMast!CCode
                If TTblCompNameField <> "" Then
                    If TTblCompCodeField = "" Then
                        RsT.Fields(i) = RstCompMast!Cname
                    Else
                        RsT.Fields(i + 1) = RstCompMast!Cname
                    End If
                End If
                RsS.MoveNext
                RsT.Update
            Loop
            '-- Update Company Name
           ' If TTblCompNameField <> "" Then
           '     gSetupCn.Execute "update " & TTblName & " set " & TTblCompNameField & "='" & _
           '     RstCompMast!Ccode & "' where " & TTblCompNameField & " is null " & IIf(TTblSelCriteria <> "", " and " & TTblSelCriteria, "")
           ' End If
        End If
        RstCompMast.MoveNext
    Loop
If RstCompMast.State = adStateOpen Then RstCompMast.Close
If TmpCn.State = adStateOpen Then TmpCn.Close
If RsS.State = adStateOpen Then RsS.Close
If RsT.State = adStateOpen Then RsT.Close
gSetupCn.CommitTrans
End Sub
Public Sub Sendkeys(text$, Optional wait As Boolean = False)
    Dim WshShell As Object
    Set WshShell = CreateObject("wscript.shell")
    WshShell.Sendkeys text, wait
    Set WshShell = Nothing
End Sub
Public Function GProcGenerateCrDrNtNo(pVtype As String, pBS As String, pBillSr As String) As String
    Dim Y As String
    Dim G_Rs As Recordset
    Dim BillSrNo As String
    BillSrNo = pBillSr
    Y = "Select Max(clng(mid(VDrCrNtNo,4))) from tblVoucher where VYear=" & gCYear & " and VDrCrNtType = '" & pBS & "'   and Vtype='" & pVtype & "' and isNumeric(VDrCrNtNo)=false and isnull(VDrCrNtNo) = false and VDrCrNtNo <> ''"
    GProcRstOpen G_Rs, Y, "R", gCn
    If G_Rs.Fields(0) = Null Or IsNumeric(G_Rs.Fields(0)) = False Then
        GProcGenerateCrDrNtNo = 1
    Else
        GProcGenerateCrDrNtNo = CLng(G_Rs.Fields(0)) + 1
    End If
    If Trim(BillSrNo) <> "" And Len(BillSrNo) = 2 Then GProcGenerateCrDrNtNo = BillSrNo & "-" & GProcGenerateCrDrNtNo
    G_Rs.Close
End Function


Public Function GProcGenerateOtherInvNo(pVtype As String, pAcCode As Long, Optional pCriteria As String) As String
    Dim Y As String
    Dim G_Rs As Recordset
    Dim BillSrNo As String
    BillSrNo = "OTH"
    Y = "Select Max(clng(mid(AdInVNo,5))) from tblAddLess where VYear=" & gCYear & " and Vtype='" & pVtype & "' and isNumeric(AdInVNo)=false and  ADTMP5 = '" & gfrmHunkStr & "' "
    
    GProcRstOpen G_Rs, Y, "R", gCn
    If G_Rs.Fields(0) = Null Or IsNumeric(G_Rs.Fields(0)) = False Then
        GProcGenerateOtherInvNo = 1
    Else
        GProcGenerateOtherInvNo = CLng(G_Rs.Fields(0)) + 1
    End If
    If Trim(BillSrNo) <> "" And Len(BillSrNo) = 3 Then GProcGenerateOtherInvNo = BillSrNo & "-" & GProcGenerateOtherInvNo
    G_Rs.Close
End Function

Public Sub Report_ExportToPDF(mCryCtrl As CrystalReport, mReportPath As String)
        Dim objCrystal As CRAXDRT.Application
        Dim objReport As CRAXDRT.Report
        Dim Tattach As String
        Set objCrystal = New CRAXDRT.Application
        Set objReport = objCrystal.OpenReport(mReportPath, 1)
        ExportReportToPDF objReport, "D:\Report.pdf", "foo"

End Sub
Public Sub ExportReportToPDF(ReportObject As CRAXDRT.Report, ByVal filename As String, ByVal ReportTitle As String)
'On Error GoTo ExportErr
   Dim objExportOptions As CRAXDRT.ExportOptions
   ReportObject.ReportTitle = ReportTitle
    With ReportObject
        .EnableParameterPrompting = False
        .MorePrintEngineErrorMessages = True
    End With

   Set objExportOptions = ReportObject.ExportOptions
    With objExportOptions
        .DestinationType = crEDTDiskFile
        .DiskFileName = filename
        .FormatType = crEFTPortableDocFormat
        .PDFExportAllPages = True
        
    End With

    ReportObject.Export False
'ExportErr:
'  MsgBox Err.Number & ", " & Err.Description
End Sub

Public Function SentWhatsApp(WAMsgType As String, WAFile As String, WAMsg As String, ToMblNo As String, Optional wPartyName)
Dim WfFileWithPath As String
Dim oExec As Object
Dim oShell As Object
Dim URL As String
Dim params As String
Dim HttpReq As New WinHttpRequest
Dim YourToken As String
Dim ProviderKey
Dim ToMblNo1 As String
Dim ToMblNo2 As String
Dim HttpResponse As String

ToMblNo1 = Left(ToMblNo, 10)
ToMblNo2 = ""   '-- Mid(ToMblNo, 12, 10)

ProviderKey = "4c934105f595a7d434c25db540594c987d779f02931bd5483fee580e710c31bd"
YourToken = "5f7402416bee018640b5f363"
URL = "http://aronertech.com/wsa/sendmessage.php"

If WAMsgType = "PDF" Then
    WfFileWithPath = gReportPath & WAFile
    Set oShell = CreateObject("WScript.Shell")
    Set oExec = oShell.Exec(App.Path & "\VB6UploadFile.exe " & YourToken & " """ & WfFileWithPath & """")
    Gdelay (3)
End If

'--for text
If ToMblNo1 <> "" Then
    If WAMsgType = "Txt" Then
       params = "key=" & ProviderKey & "&token=" & YourToken & "&number=" & ToMblNo1 & "&type=text&message=" & WAMsg
    Else
       params = "key=" & ProviderKey & "&token=" & YourToken & "&number=" & ToMblNo1 & "&type=file&filepath=http://aronertech.com/wsa/files/" & WAFile & ""
    End If
    HttpReq.Open "POST", URL, False
    HttpReq.SetRequestHeader "Content-Type", "application/x-www-form-urlencoded"
    HttpReq.Send params
    HttpResponse = HttpReq.ResponseText
'    If Mid(HttpReq.ResponseText, 12, 7) <> "success" Then
'        gCn.Execute "insert into tmpWhatsAppErr (Mbl1,PartyName,Error,USERNAME) values " _
'        & " ('" & ToMblNo1 & "','" & wPartyName & "','" & Left(HttpReq.ResponseText, 250) & "','" & gUserName & "')"
'    End If
'--{"status":"error","message":"10-Download file failed","responsetime":0.545756562}
    
    If Mid(HttpReq.ResponseText, 30, 23) = "10-Download file failed" Then
       HttpReq.Send params
       HttpResponse = HttpReq.ResponseText
    End If
    
    If Mid(HttpReq.ResponseText, 30, 23) = "10-Download file failed" Then
       HttpReq.Send params
       HttpResponse = HttpReq.ResponseText
    End If
    
    
    
    If Mid(HttpReq.ResponseText, 13, 7) <> "success" Then
        gCn.Execute "insert into tmpWhatsAppErr (Mbl1,PartyName,Error,USERNAME) values " _
        & " ('" & ToMblNo1 & "','" & wPartyName & "','" & Left(HttpReq.ResponseText, 250) & "','" & gUserName & "')"
    End If

End If
If ToMblNo2 <> "" Then
    If WAMsgType = "Txt" Then
       params = "key=" & ProviderKey & "&token=" & YourToken & "&number=" & ToMblNo2 & "&type=text&message=" & WAMsg
    Else
       params = "key=" & ProviderKey & "&token=" & YourToken & "&number=" & ToMblNo2 & "&type=file&filepath=http://aronertech.com/wsa/files/" & WAFile & ""
       'params = "key=" & ProviderKey & "&token=" & YourToken & "&number=" & ToMblNo2 & "&type=file&filepath=http://universal.aronertech.com/whatsapp/" & WAFile & ""
    End If
    HttpReq.Open "POST", URL, False
    HttpReq.SetRequestHeader "Content-Type", "application/x-www-form-urlencoded"
    HttpReq.Send params
    HttpResponse = HttpReq.ResponseText
    If Mid(HttpReq.ResponseText, 30, 23) = "10-Download file failed" Then
       HttpReq.Send params
       HttpResponse = HttpReq.ResponseText
    End If
    
    If Mid(HttpReq.ResponseText, 30, 23) = "10-Download file failed" Then
       HttpReq.Send params
       HttpResponse = HttpReq.ResponseText
    End If
    
    
    If Mid(HttpReq.ResponseText, 12, 7) <> "success" Then
        gCn.Execute "insert into tmpWhatsAppErr (Mbl1,PartyName,Error,USERNAME) values " _
        & " ('" & ToMblNo2 & "','" & wPartyName & "','" & Left(HttpReq.ResponseText, 250) & "','" & gUserName & "')"
   
    End If

End If



'--Trial On
'params = "key=" & ProviderKey & "&token=" & YourToken & "&number=" & ToMblNo & "&type=file&filepath=https://aronertech.com/wsa/files/" & WAFile & ""

End Function


Public Sub Gdelay(PauseTime As Integer)
    Dim start As Single
    start = Timer
   Do While Timer < start + PauseTime
    If (Timer < start) Then 'midnight crossover
        start = start - (86400 + 1)
    End If
    DoEvents   ' Yield to other processes.
    Loop
End Sub

