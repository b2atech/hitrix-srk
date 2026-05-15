Attribute VB_Name = "Bascivar"
Option Explicit
'Public i As Integer
Public Const FStr As String = "#########0.00"
Public Const PStr As String = "########0"
Public Const WtStr As String = "########0.000"
Public Const LStr As String = "########0.00000"
'---- Form Variables
'Public gFormAction As Integer
Public gfrmTypeStr As String
Public gfrmHunkStr As String

'----Form Bag Details Variables
Public gfrmBagInwVType As String
Public gfrmBagInwVNo As Long
'---- Back End Data Base
Public gBackEndDB As String '-- Back End
Public Const gBackEndORA As String = "ORA"
Public Const gBackEndAccess As String = "ACCESS"
Public Const gBackEndSQL As String = "SQL"
'---- Connection declaration
Public gCn As Connection
Public gSetupCn As Connection
Public gMyConn As String
Public gDataPath As String
Public gDbLocation As String
Public gODBCDSN As String '= "ABC1" '--DSN
Public gODBCPwd As String '= "HITRIX9" '--Password
Public gODBCHostStr As String '= "ABCD0506" '--HostStr
'---- Logon Variables
Public gLogOnSucc As Boolean
Public gUserName As String
Public gPassword As String
Public gChkPassword As String
Public gChkPassTrueFalse As Boolean

Public gIsAdminUser As Boolean
'---- Company Variables
Public gDbNm As String, gDbYr As String, gCCode As String
Public gCName As String, gCAdd1 As String, gCAdd2 As String, gCAdd3 As String, gCPhNo As String
Public gCCST As String, gCBST As String, gCTAN As String, gCPAN As String, gCSTRegNo As String, gCSTCNo As String
Public gCFileNo As String, gCACSTONo As String
Public gCDesignation As String '-Designation
Public gCTDSCircle As String  '- TDS Circle
Public gCTDS As String '- TDS No
Public gFinYr As String
Public gCYSDate As Date
Public gCYEDate As Date
Public gCYear As Long
Public gCSAcCode As Long '- Start AcCode   (Start & End AcCode)
Public gCEAcCode As Long '- End AcCode
Public gCSAgCode As Long '- Start AgCode
Public gCEAgCode As Long '- End AgGCode
Public gCSCGCode As Long '- Start CGCode
Public gCECGCode As Long '- End CGCode
Public gCSItCode As Long '- Start ITCode
Public gCEItCode As Long '- End ITCode
Public gCSNarrCode As Long '- Start NarrCode
Public gCENarrCode As Long '- End NarrCode

Public gCBankName As String '- Bank Name
Public gCBankAcNo As String '- Bank A/c No.
Public gCBankRtgsCode As String '- Bank Rtgs
Public gCGSTIN As String

Public gCCity As String
Public gCPin As String

Public gCIsTcsFirm As Long
Public gCTDSFirm As Integer

'--TCS
Public gAcCodeTcsPay As Long
Public gAcCodeTcsRec As Long

'----Mast Setting form Variables
Public gTaxCodeSY As Long
Public gTaxCodeSO As Long
Public gTaxCodeSD As Long
Public gTaxCodeSM As Long
Public gTaxCodeST As Long
Public gTaxCodePY As Long
Public gTaxCodePT As Long
Public gLatePayIntAcCode As Long
Public gBillDbNtAcCode As Long
Public gServiceTaxAcCode As Long
Public gTDSAcCode As Long
Public gBankCommAcCode As Long
Public gBrokerageAcCode As Long
Public gAcCodePY As Long
Public gAcCodePT As Long
Public gAcCodeSY As Long
Public gAcCodeST As Long
Public gAcCodePYHank As Long
Public gAcCodeSYHank As Long
Public gTaxCodePYHunk As Long
Public gTaxCodeSYHunk As Long
Public gExpensesAcCode As Long
Public gCashAcCode As Long
Public gMillBillDbNtAcCode As Long
Public gTDSPaidAcCode As Long
Public gIsNoVATVoucherSY As Boolean '--- No Separate VAT voucher Sale Trade
Public gIsNoVATVoucherSO As Boolean '--- No Separate VAT voucher Sale Cosignment
Public gIsNoVATVoucherSD As Boolean '--- No Separate VAT voucher Sale Depot
Public gIsMillBillAmtWithDbNt As Boolean '--Is Mill Bill Amt Include DbNt (used in Mill Bill Entry For DB Nt Calculation)
'------variables for Crystal Reporting
Public gPrintFromdt As Date
Public gPrintTodt As Date
Public gReportPath As String
'------variables for Reporting
Public gSearchStr As String
Public gReportPrint As String
Public gSelectedMenu As String
'----------- Report Selection Variables
Public gRptMillFlag As Boolean
Public gRptBrokerFlag As Boolean
Public gRptPartyFlag As Boolean
Public gRptOnDateFlag As Boolean
'---------- Main Form Variables
Public gfrmMainMenu As Menu
'---- Search Class
Public gClsSearch As New ClsSearch
''....FOR GRAPH
'Public mGrfHd(6) As String, mGrfLb1(20) As String, mGrfDt1(20) As Double, mGrfLb2(20) As String, mGrfDt2(20) As Double, mGrfTp As Integer
'----Dos Printing
Public gPc As String, gP10 As String, gP12 As String
Public gScreen As Boolean
Public Sub Main()
   ' gBackEndDB = gBackEndORA ''-- Oracle
    gfrmHunkStr = ""
    gBackEndDB = gBackEndAccess ''--Access
    gReportPrint = "c:\windows\wordpad c:\rpt1.txt"
'    If Dir("c:\Windows\system32\HtxOcx.Txt", vbDirectory) = "" Then Exit Sub

    gDataPath = CStr(App.Path) & "\HiData\"
    gDbNm = "Setup"
    gMyConn = "gSetupCn"
    Call GProcConnectionOpen(gSetupCn)
    gSelectedMenu = "E"
    frmLogOn.Show 1
    If gLogOnSucc = True Then frmMain.Show
    frmSelCompany.Show 1
    frmMain.Show
   ' frmMain.ProcAllowFrmMainMenu
    DoEvents
End Sub
'---- Open Connection
Public Function GProcConnectionOpen(pCn As Connection, Optional pDbNm As String)
    Set pCn = New Connection
    If pDbNm = "" Then
        If pCn Is Nothing Then
             gDbLocation = gDataPath & gDbNm & ".mdb"
        Else
            If pCn = gSetupCn Then
                gDbLocation = gDataPath & "setup.mdb"
            Else
                gDbLocation = gDataPath & gDbNm & ".mdb"
            End If
        End If
        pCn.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source= " & gDbLocation & ";Mode=ReadWrite|Share Deny None;Persist Security Info=False"
    Else
        pCn.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source= " & gDataPath & pDbNm & ".mdb;Mode=ReadWrite|Share Deny None;Persist Security Info=False"
    End If
    pCn.CursorLocation = adUseClient
End Function
'---- Oracle Connection
Public Function GProcConnectionOpenORA(pCn As Connection, Optional pDbNm As String)
    Set pCn = New Connection
    gDbLocation = ""
    If pDbNm = "" Then
        gODBCDSN = gDbNm
        gODBCPwd = gDbNm '--Password
        gODBCHostStr = gDbNm '--HostStr
        pCn.Open "Provider=MSDASQL.1;Password=" & gODBCPwd & " ;Persist Security Info=True;User ID=Scott;Data Source=" + gODBCDSN
    Else
        pCn.Open "Provider=MSDASQL.1;Password=" & pDbNm & " ;Persist Security Info=True;User ID=Scott;Data Source=" + pDbNm
    End If
    pCn.CursorLocation = adUseClient 'adUseServer
End Function
'---- Open Recordset
Public Function GProcRstOpen(pRstName As Recordset, pSqlQuery As String, pLockType As String, Optional pCon As Connection) As Long
Set pRstName = New Recordset
If pCon Is Nothing Then Set pCon = gCn
With pRstName
    Select Case pLockType
        Case "R"
            .Open pSqlQuery, pCon, adOpenStatic, adLockReadOnly
        Case "O"
            .Open pSqlQuery, pCon, adOpenStatic, adLockOptimistic
        End Select
        GProcRstOpen = .RecordCount
End With
End Function

