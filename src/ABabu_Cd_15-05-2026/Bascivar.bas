Attribute VB_Name = "Bascivar"
Option Explicit
Public i As Integer
Private Declare Function SQLConfigDataSource Lib "ODBCCP32.DLL" _
        (ByVal hwndParent As Long, ByVal fRequest As Long, _
          ByVal lpszDriver As String, ByVal lpszAttributes As String) _
          As Long
Public Const SQL_NULL_HANDLE = 0
Private Const ODBC_ADD_DSN = 1 ' Add data source
Public Const ODBC_CONFIG_DSN = 2 ' Configure (edit) data source
Public Const ODBC_REMOVE_DSN = 3 ' Remove data source
Public Const ODBC_ADD_SYS_DSN = 4 ' Add data source
Public Const ODBC_CONFIG_SYS_DSN = 5 ' Configure (edit) system data source
Public Const ODBC_REMOVE_SYS_DSN = 6 ' Remove system data source
    
'*****************************************************************************************************************
Public Objshell 'For Windows os sENDKEYS
Public gCmd As New ADODB.Command              '---- Connection Variable for Adodb Command bject Declared
Public gCmd1 As New ADODB.Command              '---- Connection Variable for Adodb Command bject Declared
Public pPWD As String, pUID As String, pServ As String
Public gDSN As String        '---- Used To Store DSN Name
Public gDbNm As String       '---- Used To Store Database Name
Public gDUsrId As String     '---- Used To Store Database UserName
Public gDPswd As String      '---- Used To Store Database Password
Public Const FStr As String = "#########0.00"
Public Const PStr As String = "########0"
Public Const WtStr As String = "########0.000"
Public Const LStr As String = "########0.00000"
Public gOldBillYear As Long      '---- Used To Store Database Name

'---- Form Variables
'Public gFormAction As Integer
Public gfrmTypeStr As String
Public gfrmHunkStr As String
Public mShowRec As String
Public mShowRecSaleAc As String
Public gDepotConsiMillHelp As Boolean
Public gMailsendToMill As Boolean

'----Form Bag Details Variables
Public gfrmBagInwVType As String
Public gfrmBagInwVNo As Long
Public gfrmBagLrDt As Date
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
Public gRstAccMast As Recordset


'---- Logon Variables
Public gLogOnSucc As Boolean
Public gUserName As String
Public gPassword As String
Public gIsAdminUser As Boolean
'---- Company Variables
Public gDbYr As String, gCCode As String
Public gCName As String, gCAdd1 As String, gCAdd2 As String, gCAdd3 As String, gCPhNo As String
Public gCCST As String, gCBST As String, gCTAN As String, gCPAN As String, gCSTRegNo As String, gCSTCNo As String, gCCapitalName As String
Public gCFileNo As String, gCACSTONo As String
Public gCDesignation As String '-Designation
Public gCTDSCircle As String  '- TDS Circle
Public gCTDS As String '- TDS No
Public gFinYr As String
Public gCYSDate As Date
Public gCYEDate As Date
Public gCYear As Long
Public gCompIsYrEnd As Integer
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
'Public gCFirm As String '- Firm Short
Public gCisGst As Long '- End NarrCode
Public gIsMillBillFirm As Long
Public gCHamaliRt As Double
Public gCLevyRt As Double
Public gCSrvTaxRt As Double
Public gCSBCessRt As Double
Public gCKKCessRt As Double
Public gCIsDepotFirm As Long
Public gCDepotMainFirm As Long
Public gCDepotMainFirmCompCd As String
Public gCGStin As String
Public gCIBOffice As String
Public gCMailId As String
Public gCMailPass As String
Public gCBankName As String
Public gCBankAcNo As String
Public gCBankRTGSCode As String
Public gCIsTcsFirm As Integer
Public gCIsLog As Integer
'----Mast Setting form Variables
Public gCashAcCode As Long
Public gTaxCodeSY As Long
Public gTaxCodeSO As Long
Public gTaxCodeSD As Long
Public gTaxCodeSM As Long
Public gTaxCodeST As Long
Public gTaxCodePY As Long
Public gTaxCodePT As Long
Public gTaxCodePI As Long
Public gTaxCodePYHunk As Long
Public gTaxCodeSYHunk As Long
Public gTaxCodeSYExempt As Long
Public gTaxCodePYExempt As Long

Public gTranCdSY As String
Public gTranCdSO As String
Public gTranCdSD As String
Public gTranCd As String
Public gTranCdST As String
Public gTranCdPY As String
Public gTranCdPT As String
Public gTranCdPI As String
Public gTranCdPYHunk As String
Public gTranCdSYHunk As String
Public gTranCdSYExempt As String
Public gTranCdPYExempt As String

Public gTranCdRY As String
Public gTranCdRO As String
Public gTranCdVY As String
Public gTranCdVI As String
Public gLatePayIntAcCodeRec As Long
Public gLatePayIntAcCodeRecGST As Long
Public gBillDbNtAcCode As Long
Public gServiceTaxAcCode As Long
Public gTDSAcCodeRec As Long

Public gBankCommAcCode As Long
Public gBrokerageAcCode As Long
Public gAcCodePY As Long
Public gAcCodePT As Long
Public gAcCodeSY As Long
Public gAcCodeST As Long
Public gAcCodePYHank As Long
Public gAcCodeSYHank As Long
Public gAcCodePYExempt As Long
Public gAcCodeSYExempt As Long
Public gBrokerageSaleAc As Long
Public gCommissionSaleAc As Long
Public gBrokerageItem As Long
Public gCommissionItem As Long
Public gUnloadingChg As Long
Public gRoundOffAc As Long
Public gTcsRec As Long
Public gTcsPay As Long
Public gAcCodeRY As Long
Public gAcCodeVY As Long

'Public gExpensesAcCode As Long
'Public gMillBillDbNtAcCode As Long
'Public gTDSPaidAcCode As Long
'Public gIsNoVATVoucherSY As Boolean '--- No Separate VAT voucher Sale Trade
'Public gIsNoVATVoucherSO As Boolean '--- No Separate VAT voucher Sale Cosignment
'Public gIsNoVATVoucherSD As Boolean '--- No Separate VAT voucher Sale Depot
'Public gIsMillBillAmtWithDbNt As Boolean '--Is Mill Bill Amt Include DbNt (used in Mill Bill Entry For DB Nt Calculation)

Public gDiscAcCodeRec As Long
Public gLatePayIntAcCodePay As Long
Public gTDSAcCodePay As Long
Public gDiscAcCodePay As Long
Public gSgstTaxRate As Double
Public gCgstTaxRate As Double
Public gIgstTaxRate As Double
Public gSgstAcCode As Long
Public gCgstAcCode As Long
Public gIgstAcCode As Long
Public gSgstRCMAcCode As Long
Public gCgstRCMAcCode As Long
Public gIgstRCMAcCode As Long
Public gLPGrase As Long
Public gLPIntRt As Double
Public gTDSRate As Double
Public gLPExtraGrase As Long
Public gLpGraceCrNote As Long
Public gLpIntRtCrNote As Double
Public gSgstRCMRecCode As Long
Public gCgstRCMRecCode As Long
Public gIgstRCMRecCode As Long
Public gLPIntAcCodeWithoutGST As Long
Public gTdsOnPurchCode As Long
Public gTdsOnSalesCode As Long

'------variables for Crystal Reporting
Public gPrintFromdt As Date
Public gPrintTodt As Date
Public gReportPath As String
'------variables for Reporting
Public gSearchStr As String
Public gReportPrint As String
Public gSelectedMenu As String
Public gSelectedMenuFormEnt As String
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
'Report on Grid Variables
Public QueryStr As String
Public LastVouDate As Date
Public PrintTotal(40) As Boolean
Public HideCol(40) As Boolean
Public RepeateCol(40) As Boolean
Public DateRange As String
Public DateCaption As String
Public Colwith(20) As Long
Public DateCaptionFromTo As String
Public DateCaptionAsOn As String
Public rsReport As New Recordset
Public ReportOpt(7) As Boolean
Public PrintToDateStatus As Boolean
Public PrintReportFlag As Boolean
Public PrintFromDate As Date
Public PrintToDate As Date
Public ReportRecStr As String
Public MenuRecStr As String
Public Flage As Boolean


'Public gCGSTPayAc As Long
'Public gSGSTPayAc As Long
'Public gIGSTPayAc As Long
'----Dos Printing
Public gPc As String, gP10 As String, gP12 As String
Public gScreen As Boolean
Public Sub Main()
   ' gBackEndDB = gBackEndORA ''-- Oracle
    gfrmHunkStr = ""
    gBackEndDB = gBackEndAccess ''--Access
    gReportPrint = "c:\windows\wordpad c:\rpt1.txt"
    'If Dir("c:\Windows\system32\HtxOcx.Txt", vbDirectory) = "" Then Exit Sub
    gDataPath = CStr(App.Path) & "\HiData\"
    gReportPath = CStr(App.Path) & "\HiReports\"

     gDSN = "HiTrix"     '----- DSN Name
     pUID = "sa"
''----------------ICh Office
'     pServ = "srknew" '-------- Hitrix
     pServ = "VSERVER\SQLEXPRESS" '------- VTC Ich Office
'     pServ = "VISHNU3-PC\SQLEXPRESS" '------- VTC Mum Office
'----------------Mum Office
'     pServ = "DESKTOP-U7SJKS8\SQLEXPRESS"  'LAPTOP no use
     
     
     
     pPWD = "123456"
'     gDbNm = "Vishnu"
     gDbNm = "Vishnu"
'-----------------
    GProcConnectionOpen gCn, pServ, gDbNm, pUID, pPWD
    gSelectedMenu = "E"
    frmLogOn.Show 1
    If gLogOnSucc = True Then frmMain.Show
    frmSelCompany.Show 1
    frmMain.Show
    DoEvents
End Sub
'---- Open Connection
Public Function GProcConnectionOpenOld(pCn As Connection, Optional Pdbnm As String)
    Set pCn = New Connection
    If Pdbnm = "" Then
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
        pCn.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source= " & gDataPath & Pdbnm & ".mdb;Mode=ReadWrite|Share Deny None;Persist Security Info=False"
    End If
    pCn.CursorLocation = adUseClient
End Function
'---- Oracle Connection
Public Function GProcConnectionOpenORA(pCn As Connection, Optional Pdbnm As String)
    Set pCn = New Connection
    gDbLocation = ""
    If Pdbnm = "" Then
        gODBCDSN = gDbNm
        gODBCPwd = gDbNm '--Password
        gODBCHostStr = gDbNm '--HostStr
        pCn.Open "Provider=MSDASQL.1;Password=" & gODBCPwd & " ;Persist Security Info=True;User ID=Scott;Data Source=" + gODBCDSN
    Else
        pCn.Open "Provider=MSDASQL.1;Password=" & Pdbnm & " ;Persist Security Info=True;User ID=Scott;Data Source=" + Pdbnm
    End If
    pCn.CursorLocation = adUseClient 'adUseServer
End Function
Function cmdODBC_Create(Pdbnm As String, PServerNm As String, Optional sUID As String, Optional sPWD As String)
Dim params As Integer
Dim lpszDriver As String
Dim lpszAttributes As String
Dim RetVal As Long
lpszDriver = "SQL Server"

Dim DSN_NAME
Dim SERVER_NAME
Dim DATABASE_NAME

DSN_NAME = "Hitrix"
SERVER_NAME = PServerNm
DATABASE_NAME = Pdbnm

lpszAttributes = "DSN=" & DSN_NAME & Chr$(0)
lpszAttributes = lpszAttributes & "Description=DSN" & Chr$(0)
lpszAttributes = lpszAttributes & "Server=" & SERVER_NAME & Chr$(0)
lpszAttributes = lpszAttributes & "Database=" & DATABASE_NAME & Chr$(0)

RetVal = SQLConfigDataSource(0, ODBC_REMOVE_DSN, lpszDriver, "DSN=" & DSN_NAME & Chr$(0))
RetVal = SQLConfigDataSource(0, ODBC_ADD_DSN, lpszDriver, lpszAttributes)

lpszAttributes = "DSN=" & DSN_NAME & Chr$(0)

RetVal = SQLConfigDataSource(0, ODBC_CONFIG_DSN, lpszDriver, lpszAttributes)

If params > 0 Then End
End Function
Public Function GProcConnectionOpen(pCn As ADODB.Connection, PservNm As String, Optional Pdbnm As String, Optional pUID As String, Optional pPSWD As String) As Boolean
 On Error GoTo errroutine
           GProcConnectionOpen = False
'           If pCn.State = 1 Then pCn.Close
           If TypeName(pCn) = "Connection" Then Set pCn = Nothing
           Set pCn = New ADODB.Connection
          Call cmdODBC_Create(Pdbnm, PservNm, pUID, pPSWD)
          pCn.ConnectionString = "Dsn=" & gDSN & ";DATABASE=" & Pdbnm & ";Uid=" & pUID & ";PWD=" & pPSWD
          pCn.Open
           gCmd.CommandType = adCmdStoredProc
           gCmd.ActiveConnection = pCn
           gCmd.CommandTimeout = 0
          pCn.CursorLocation = adUseClient
          GProcConnectionOpen = True
        Exit Function
errroutine:
  GProcErrorHandler
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

