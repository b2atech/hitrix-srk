VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form frmRptSelection 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Selection"
   ClientHeight    =   5835
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10800
   Icon            =   "frmRptSelection.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5835
   ScaleWidth      =   10800
   Begin VB.TextBox txt 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   4
      Left            =   1860
      TabIndex        =   4
      Top             =   1140
      Visible         =   0   'False
      Width           =   7410
   End
   Begin VB.TextBox txt 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   3
      Left            =   8685
      TabIndex        =   7
      Text            =   "3"
      Top             =   2745
      Width           =   1575
   End
   Begin VB.CommandButton cmdDbNt 
      Caption         =   "&Cancel Interest Debit Note"
      Height          =   480
      Index           =   1
      Left            =   11385
      TabIndex        =   19
      Top             =   2655
      Visible         =   0   'False
      Width           =   1290
   End
   Begin VB.CommandButton cmdCrystal 
      Height          =   480
      Index           =   1
      Left            =   4260
      Picture         =   "frmRptSelection.frx":058A
      Style           =   1  'Graphical
      TabIndex        =   17
      ToolTipText     =   "Print"
      Top             =   5085
      Width           =   1185
   End
   Begin VB.CommandButton cmdCrystal 
      Height          =   480
      Index           =   0
      Left            =   2415
      Picture         =   "frmRptSelection.frx":0BF4
      Style           =   1  'Graphical
      TabIndex        =   16
      Top             =   5085
      Width           =   1185
   End
   Begin VB.CommandButton cmdDbNt 
      Caption         =   "&Apply Interest Debit Note"
      Height          =   480
      Index           =   0
      Left            =   11430
      TabIndex        =   14
      Top             =   3240
      Visible         =   0   'False
      Width           =   1305
   End
   Begin VB.CheckBox ChkConsolidate 
      BackColor       =   &H00EDFDFE&
      Caption         =   "Is All Firm  Report ?"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   330
      Left            =   1845
      TabIndex        =   8
      Top             =   3870
      Width           =   3360
   End
   Begin VB.TextBox txt 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   2
      Left            =   5130
      TabIndex        =   6
      Text            =   "2"
      Top             =   2745
      Width           =   1575
   End
   Begin VB.TextBox txt 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   1
      Left            =   1860
      TabIndex        =   5
      Text            =   "1"
      Top             =   2760
      Width           =   1575
   End
   Begin VB.TextBox txt 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   0
      Left            =   1860
      TabIndex        =   3
      Top             =   690
      Width           =   7410
   End
   Begin VB.CommandButton cmdExit 
      Cancel          =   -1  'True
      Caption         =   "E&xit"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   480
      Left            =   6180
      TabIndex        =   18
      Top             =   5085
      Width           =   1095
   End
   Begin Crystal.CrystalReport CrystalReport1 
      Left            =   210
      Top             =   3555
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
   Begin MSComctlLib.ProgressBar prgBar 
      Height          =   255
      Left            =   165
      TabIndex        =   15
      Top             =   4470
      Visible         =   0   'False
      Width           =   6375
      _ExtentX        =   11245
      _ExtentY        =   450
      _Version        =   393216
      Appearance      =   1
   End
   Begin MSComCtl2.DTPicker dtpToDt 
      Height          =   360
      Left            =   5130
      TabIndex        =   2
      Top             =   225
      Width           =   1800
      _ExtentX        =   3175
      _ExtentY        =   635
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      CustomFormat    =   "dd/MM/yyyy"
      Format          =   161153027
      CurrentDate     =   38050
   End
   Begin MSComCtl2.DTPicker dtpFrDt 
      Height          =   360
      Left            =   1860
      TabIndex        =   1
      Top             =   225
      Width           =   1800
      _ExtentX        =   3175
      _ExtentY        =   635
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      CustomFormat    =   "dd/MM/yyyy"
      Format          =   161153027
      CurrentDate     =   38050
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Broker           :"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   240
      Left            =   360
      TabIndex        =   21
      Top             =   1140
      Visible         =   0   'False
      Width           =   1290
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "&Extra Grase Days :"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   240
      Index           =   1
      Left            =   6885
      TabIndex        =   20
      Top             =   2790
      Width           =   1680
   End
   Begin VB.Shape Shape2 
      FillColor       =   &H00BDDFD1&
      FillStyle       =   0  'Solid
      Height          =   765
      Left            =   1560
      Shape           =   4  'Rounded Rectangle
      Top             =   4950
      Width           =   6780
   End
   Begin VB.Line Line2 
      BorderColor     =   &H000040C0&
      X1              =   210
      X2              =   10575
      Y1              =   3240
      Y2              =   3240
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "&Late Pay Interest Report"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   330
      Left            =   360
      TabIndex        =   11
      Top             =   2100
      Width           =   2595
   End
   Begin VB.Line Line1 
      BorderColor     =   &H000040C0&
      X1              =   240
      X2              =   10530
      Y1              =   1590
      Y2              =   1590
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "&Grase Days :"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   240
      Index           =   4
      Left            =   3615
      TabIndex        =   13
      Top             =   2745
      Width           =   1155
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "&Interest %    :"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   240
      Index           =   2
      Left            =   360
      TabIndex        =   12
      Top             =   2760
      Width           =   1185
   End
   Begin VB.Label lblName 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Party             :"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   240
      Left            =   360
      TabIndex        =   10
      Top             =   690
      Width           =   1305
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "&To   :"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   240
      Left            =   4260
      TabIndex        =   9
      Top             =   225
      Width           =   480
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "&From Date     :"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   240
      Index           =   0
      Left            =   360
      TabIndex        =   0
      Top             =   225
      Width           =   1350
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00000080&
      FillColor       =   &H00EDFDFE&
      FillStyle       =   0  'Solid
      Height          =   4770
      Left            =   225
      Top             =   60
      Width           =   10335
   End
End
Attribute VB_Name = "frmRptSelection"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim i As Long
'--- Validate data
Private Function ValidataData(Optional CryIndex As Integer) As Boolean
    Select Case gSelectedMenu
      '-- Broker, Party & Due Dayswise Aging Report
        Case "Broker, Party & Due Dayswise Aging Report" '-"Broker, Party & Due Dayswise Aging Report"
            If IsNumeric(txt(3)) = False Then txt(3) = 0
            If IsNumeric(txt(4)) = False Then txt(4) = 0
            If CLng(txt(3)) > CLng(txt(4)) Then txt(3) = txt(4)
            If txt(4) = 0 Then
                MsgBox "Check Due Days To.", vbOKOnly + vbInformation, Me.Caption
                Exit Function
            End If
    End Select
ValidataData = True
End Function
Private Sub cmdCrystal_Click(Index As Integer)
Dim X1 As String
Dim FromDt As Date
Dim ToDt As Date
Dim ShowRepo As Boolean
ShowRepo = False
Dim wAcCode As Long
Dim wBrkCode As Long

FromDt = Format(dtpFrDt.Value, "dd/MM/yyyy")
ToDt = Format(dtpToDt.Value, "dd/MM/yyyy")
'--- Validata Data
If ValidataData(Index) = False Then Exit Sub
GProcCrystalRptPreparation CrystalReport1, "Window"
With CrystalReport1
    If Index = 1 Then .Destination = crptToPrinter
    Select Case gSelectedMenu
        '- "Broker & Partywise Aging Report"
        Case "Broker & Partywise Aging Report"
            GProcPrepareAging FromDt, txt(0)
            If gBackEndDB = gBackEndAccess Then
                .DataFiles(0) = gDbLocation
            Else '-- Oracle
                .Connect = "DSN=" + gODBCDSN + " ;UID=Scott;PWD=" + gODBCPwd + ";DBQ=" + gODBCHostStr
            End If
            .Formulas(4) = "wFrToDt = ' On Date: ' & '" & FromDt & "'"
            .ReportFileName = gReportPath & "rptOutStandingAging.rpt"
            .SelectionFormula = "{tmpGentbl.UserName}='" & gUserName & "'"
         '-- Broker, Party & Due Dayswise Aging Report
        Case "Broker, Party & Due Dayswise Aging Report"
            GProcPrepareAging FromDt, txt(0), txt(3), txt(4)
            If gBackEndDB = gBackEndAccess Then
                .DataFiles(0) = gDbLocation
            Else '-- Oracle
                .Connect = "DSN=" + gODBCDSN + " ;UID=Scott;PWD=" + gODBCPwd + ";DBQ=" + gODBCHostStr
            End If
            .Formulas(4) = "wFrToDt = ' On Date: ' & '" & FromDt & "'"
            .ReportFileName = gReportPath & "rptOutStandingAgingDayswise.rpt"
            .SelectionFormula = "{tmpGentbl.UserName}='" & gUserName & "'"
            '.SubreportToChange = "rptSubOutStandingAging.rpt"
            If gBackEndDB = gBackEndAccess Then
                .DataFiles(0) = gDbLocation
            Else '-- Oracle
                .Connect = "DSN=" + gODBCDSN + " ;UID=Scott;PWD=" + gODBCPwd + ";DBQ=" + gODBCHostStr
            End If
         Case "Late Payment Interest List", "Early Payment Interest List"
              If txt(0) <> "" Then
                 wAcCode = GProcGetColumnValue("tblMastAccount", "AcName", txt(0), "S", "AcCode", "N")
              End If
              If txt(4) <> "" Then
                 wBrkCode = GProcGetColumnValue("tblMastAccount", "AcName", txt(4), "S", "AcCode", "N")
              End If
              
              ShowRepo = True
         
                gCmd.CommandText = "PrcPrepareIntDebitNote"
                gCmd.Parameters.Refresh
                gCmd.Parameters("@From_dt") = Format(FromDt, "yyyy/mm/dd")
                gCmd.Parameters("@To_dt") = Format(ToDt, "yyyy/mm/dd")
                gCmd.Parameters("@UCode") = gUserName
                gCmd.Parameters("@Grace") = Val(txt(2)) + Val(txt(3))
                If ChkConsolidate.Value = 1 Then
                   gCmd.Parameters("@VFirm") = ""
                Else
                   gCmd.Parameters("@VFirm") = Trim(gCCode)
                End If
                If wAcCode <> 0 Then gCmd.Parameters("@Accode") = wAcCode
                If wBrkCode <> 0 Then gCmd.Parameters("@Accode_Br") = wBrkCode
                If gSelectedMenu = "Early Payment Interest List" Then
                   gCmd.Parameters("@IntTP") = "E"
                Else
                   gCmd.Parameters("@IntTP") = "L"
                End If
                gCmd.Execute
                X1 = "{TmpGentbl.UserName}='" & gUserName & "'"
              .SelectionFormula = X1
              .Formulas(4) = "wFrToDt = 'From : ' & '" & Format(dtpFrDt.Value, "dd/MM/yyyy") & "' & '  To ' & '" & Format(dtpToDt.Value, "dd/MM/yyyy") & "'"
              .Formulas(5) = "wIntRt =  '" & txt(1) & "'"
              .Formulas(6) = "wGrase =  '" & Val(txt(2)) + Val(txt(3)) & "'"
              If gSelectedMenu = "Early Payment Interest List" Then
                 .Formulas(7) = "wRptTp =  'E'"
              Else
                 .Formulas(7) = "wRptTp =  'L'"
              End If
              .ReportFileName = gReportPath & "rptOutStandingSALatePayInt.rpt"
              
'        Case "Late Payment Interest Register", "Late Payment Interest Register (Mill Bill)"
'             If gSelectedMenu = "Late Payment Interest Register" Then
'                GProcOutStandingSALatePayInt FromDt, ToDt, ChkConsolidate, txt(1), txt(2), txt(0), ChkIntLateFrmInvDt, IIf(Index = 2, True, False)
'             Else
'                GProcOutStandingSALatePayInt FromDt, ToDt, ChkConsolidate, txt(1), txt(2), txt(0), ChkIntLateFrmInvDt, IIf(Index = 2, True, False), True
'             End If
'             .Formulas(4) = "wFrToDt = 'From : ' & '" & Format(dtpFrDt.Value, "dd/MM/yyyy") & "' & '  To ' & '" & Format(dtpToDt.Value, "dd/MM/yyyy") & "'"
'             If ChkIntLateFrmInvDt = 1 Then '-- Late Days From
'                .Formulas(5) = "LateDaysFromStr = 'Late Days From Invoice Date'"
'             Else
'                .Formulas(5) = "LateDaysFromStr = 'Late Days From Due Date'"
'             End If
'             '-- Is Interest Rate Monthly
'             If ChkIntRateMonthly = 1 Then
'                .Formulas(8) = "IsIntRateMonthly ='1'"
'             Else
'                .Formulas(8) = "IsIntRateMonthly ='0'"
'             End If
'             .DataFiles(0) = gDataPath & "\Setup.mdb"
'             gReportPath = App.Path & "\HiReports\"
'             .ReportFileName = gReportPath & "rptOutStandingSALatePayInt.rpt"
'             .SelectionFormula = "{tmpGentbl.UserName}='" & gUserName & "' and {@Balinterest}>0"
        '-Interest Demand Note
        Case "Interest Demand Note"
            .Formulas(4) = "wFrToDt = 'From : ' & '" & Format(dtpFrDt.Value, "dd/MM/yyyy") & "' & '  To ' & '" & Format(dtpToDt.Value, "dd/MM/yyyy") & "'"
            '- Vtype ,Vdt
            X1 = " ({tblRecVsSale.RECVTYPE}='BR' or {tblRecVsSale.RECVTYPE}='CR' ) and {tblRecVsSale.RECVDT} in Date(" & Format(dtpFrDt.Value, "yyyy,mm,dd") & ") to Date(" & Format(dtpToDt.Value, "yyyy,mm,dd") & ")  " '
            If txt(0) <> "" Then X1 = X1 + " and {tblMastAccount.AcName} = '" & txt(0) & "'" '-Party
            X1 = X1 + " and {@Balinterest}> 0 "
            .SelectionFormula = X1
            If ChkConsolidate = 1 Then
               .ReportFileName = gReportPath & "rptAccInterestDbNtNoteSummery.rpt"
            Else
               .ReportFileName = gReportPath & "rptAccInterestDbNtNote.rpt"
            End If
        '-- Stock Statement
        Case "Stock Statement Datewise", "Stock Statement"
            If gSelectedMenu = "Stock Statement" Then '"Stock Statement" inward Type wise
                GprocCollectPurStock dtpFrDt, dtpToDt, True
            Else
                GprocCollectPurStock dtpFrDt, dtpToDt
            End If
            If gBackEndDB = gBackEndAccess Then
                .DataFiles(0) = gDbLocation
            Else '-- Oracle
                .Connect = "DSN=" + gODBCDSN + " ;UID=Scott;PWD=" + gODBCPwd + ";DBQ=" + gODBCHostStr
            End If
            .Formulas(4) = "wFrToDt = 'From : ' & '" & Format(dtpFrDt.Value, "dd/MM/yyyy") & "' & '  To ' & '" & Format(dtpToDt.Value, "dd/MM/yyyy") & "'"
            If gSelectedMenu = "Stock Statement" Then '-- Stock Statement Inward Type wise
                .SelectionFormula = "{tmpGentbl.UserName}='" & gUserName & "'"
                .ReportFileName = gReportPath & "rptStkRptInTypeWise.rpt"
            Else
                .SelectionFormula = "{tmpGentbl.UserName}='" & gUserName & "'  and {@balBag}<>0 and {@BalWt}<>0"
                .ReportFileName = gReportPath & "rptStkRptDatewise.rpt"
            End If
            .Formulas(1) = ""
            .Formulas(2) = ""
            .Formulas(3) = ""
        '---Pending Purchase SIT Register
        Case "Pending Purchase SIT Register"
            If gBackEndDB = gBackEndAccess Then
                .DataFiles(0) = gDbLocation
            Else '-- Oracle
                .Connect = "DSN=" + gODBCDSN + " ;UID=Scott;PWD=" + gODBCPwd + ";DBQ=" + gODBCHostStr
            End If
            .Formulas(4) = "wFrToDt = 'From : ' & '" & Format(dtpFrDt.Value, "dd/MM/yyyy") & "' & '  To ' & '" & Format(dtpToDt.Value, "dd/MM/yyyy") & "'"
            X1 = "{tblAddLess.VDt} in Date(" & Format(dtpFrDt.Value, "yyyy,mm,dd") & ") to Date(" & Format(dtpToDt.Value, "yyyy,mm,dd") & ")"
            X1 = X1 + " and {tblAddLess.VType}='PT' and ({tblVouSub.VSubBag}-{tblVouSub.VSubIssBag}>0)"
            If txt(0) <> "" Then X1 = X1 + " and {tblMastAccount.AcName}='" & txt(0) & "'"
            .SelectionFormula = X1
            .ReportFileName = gReportPath & "rptStkRptPendingInward.rpt"
'         Case "Daily Massage"
'            Dim RsTmp As Recordset
'            Dim X2 As String
'            Dim CondStr As String
'            CondStr = "tblAddless.vdt>=cdate('" & Format(dtpFrDt.Value, "dd/MM/yyyy") & "') and   tblAddless.vdt<=cdate('" & Format(dtpToDt.Value, "dd/MM/yyyy") & "')"
'
'            X1 = "SELECT tblMastAccount.AcFaxNo,AdInvNo, tblAddLess.Vdt, AdbillAmt,ItShort+'-'+ItTicket,tblVouSub.VSubBag,tblVouSub.VSubNetRate,tblVouSub.VSubRtPer, tblBrk.AcName" _
'                 & " FROM tblAddless,tblMastAccount,tblMastItem,tblVouSub,tblMastAccount tblBrk,tblOutStanding" _
'                 & " WHERE tblAddLess.vtype in ('SY','ST','SM','SD','SO') and tblAddless.AdAcDrCode = tblMastAccount.AcCode  and tblAddless.Vno = tblVousub.Vno" _
'                 & " and tblAddless.Vtype = tblVousub.Vtype and tblAddless.VYear = tblVousub.Vyear and tblVousub.VSubItCode = tblMastItem.Itcode" _
'                 & " and tblAddLess.Vno = tblOutStanding.Vno and tblAddLess.Vyear = tblOutStanding.Vyear and tblAddLess.Vtype = tblOutStanding.Vtype" _
'                 & " and tblOutStanding.OutBrokerCode = tblBrk.AcCode and tblMastAccount.AcFaxNo <> '' and " & CondStr & " order by tblAddLess.Vdt,AdInvNo"
'
'             Dim objExcel As Object
'             Dim objWorkbook As Object
''             Dim oSheet As Object
'             Set objExcel = CreateObject("Excel.Application")
'             Set objWorkbook = objExcel.Workbooks.Open("C:\test.xls")
'
''              objExcel.Columns("A:A").Select
'              objExcel.cells.ClearContents
'
''             objExcel.Col(1).Delete
''             objExcel.Col(2).Delete
'
'             objExcel.Application.Visible = False
''            objExcel.Workbooks.Add
'             If GProcRstOpen(RsTmp, X1, "O") > 0 Then
'               With RsTmp
'                   .MoveFirst
'                   i = 1
'                   Do While Not .EOF
'                      X2 = Left("Bl.Dt-" + CStr(.Fields(2)) + "-No-" + .Fields(1) + "-Amt-" + CStr(.Fields(3)) + "-Count-" + .Fields(4) + "-Bg-" + CStr(.Fields(5)) + "-Rate-" + CStr(.Fields(6)) + "-Brk-" + .Fields(8) + "-Firm-" + gCName + Space(160), 160)
'                      objExcel.cells(i, 1).Value = RsTmp.Fields(0)
'                      objExcel.cells(i, 2).Value = X2
'                      DoEvents
'                      i = i + 1
'                     .MoveNext
'                   Loop
'               End With
'               RsTmp.Close
'             End If
'             objExcel.ActiveWorkbook.Save
'             '"C:\test.xls"
'
'             objExcel.ActiveWorkbook.Close
'             objExcel.Application.Quit
'
'             MsgBox ("Massage Genrated...")
             
        Case "Receipt Against Invoice Detail"
'             GProcOutStandingSALatePayInt FromDt, ToDt, ChkConsolidate, txt(1), txt(2), txt(0), ChkIntLateFrmInvDt, False
'             '-- Is Interest Rate Monthly
'             If ChkIntRateMonthly = 1 Then
'                .Formulas(8) = "IsIntRateMonthly ='1'"
'             Else
'                .Formulas(8) = "IsIntRateMonthly ='0'"
'             End If
'             .Formulas(4) = "wFrToDt = 'Receipt Date From : ' & '" & Format(dtpFrDt.Value, "dd/MM/yyyy") & "' & '  To ' & '" & Format(dtpToDt.Value, "dd/MM/yyyy") & "'"
'             .DataFiles(0) = gDataPath & "\Setup.mdb"
'             gReportPath = App.Path & "\HiReports\"
'             .ReportFileName = gReportPath & "rptReceiptInvoiceDetail.rpt"
'             .SelectionFormula = "{tmpGentbl.UserName}='" & gUserName & "' and {tmpGentbl.amt2}>0    "
    End Select
    If gSelectedMenu <> "Daily Massage" Then
      .WindowState = crptMaximized
      If ShowRepo = True Then
        .Action = 1
        .PageZoom (120)
      End If
    End If
End With
End Sub
Private Sub cmdExit_Click()
    Unload Me
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
    Case 13:
       Sendkeys "{TAB}"
    Case 27:
        Unload Me
    End Select
End Sub
Private Sub Form_Unload(Cancel As Integer)
    gRptPartyFlag = False
    gRptBrokerFlag = False
    gRptMillFlag = False
End Sub
Private Sub Form_Load()
    Me.Top = 700
    Me.Left = Screen.Width / 2 - Me.Width / 2
    If gRptOnDateFlag = True Then
        dtpFrDt.Value = Format(Date, "dd/MM/yyyy")
    Else
        dtpFrDt.Value = Format(gCYSDate, "dd/MM/yyyy")
    End If
    dtpToDt.Value = Now() 'Format(gCYEDate, "dd/MM/yyyy")
    Me.Caption = gSelectedMenu
    If gRptBrokerFlag = True Then
        lblName.Caption = "Broker"
    End If
    If gRptPartyFlag = True Then
        lblName.Caption = "Party"
    End If
    For i = 1 To 3
        txt(i) = 0
    Next i
    '-- Activate Controls
    GProcActivateControls Me, False
    dtpFrDt.Enabled = True
    dtpToDt.Enabled = True
    txt(0).Enabled = True
    Select Case gSelectedMenu
        Case "Stock Statement" '-"Stock Statement"
             dtpFrDt.Enabled = False
             txt(0).Enabled = False '-Mill
        Case "Late Payment Interest List", "Early Payment Interest List"
             txt(1).Enabled = True
             txt(2).Enabled = True
             txt(3).Enabled = True
             If gSelectedMenu = "Early Payment Interest List" Then
                txt(1) = "15.00"
             Else
                txt(1) = "18.00"
             End If
             txt(2) = "5"
             txt(4).Visible = True
             txt(4).Enabled = True
             Label4.Visible = True
             ChkConsolidate.Enabled = True
'        Case "Late Payment Interest Register", "Late Payment Interest Register (Mill Bill)" ''===  Late Payment Interest Register
'            txt(0).Enabled = True '-party
'            ChkConsolidate.Enabled = True
'            txt(1).Enabled = True '-Int %
'            txt(2).Enabled = True '- Grase Days
'            DtpDbNtDt.Enabled = True '--DB Nt Dt
'            DtpDbNtDt.Value = Date
'            ChkIntLateFrmInvDt.Enabled = True '--Interest Late From Inv Dt
'            ChkIntRateMonthly.Enabled = True '--Interest Rate Monthly
'            cmdDbNt(0).Visible = True '- Apply Int Debit Note
'            cmdDbNt(1).Visible = True '- Cancel Int Debit Note
        Case "Broker, Party & Due Dayswise Aging Report" '=== "Broker, Party & Due Dayswise Aging Report"
            txt(3).Enabled = True
            txt(4).Enabled = True
        Case "Interest Demand Note" '== Interest Demand Note
            txt(0).Enabled = True '-party
            ChkConsolidate.Enabled = True
            ChkConsolidate.Caption = "Summery"
'        Case "Daily Massage"
'            lblName.Visible = False
'            txt(0).Visible = False
'            ChkConsolidate.Visible = False
'            ChkIntDbNtShort.Visible = False
'            Label3.Visible = False
'            Label1(4).Visible = False
'            txt(2).Visible = False
'            Label1(2).Visible = False
'            txt(1).Visible = False
'            ChkIntLateFrmInvDt.Visible = False
'            Label5.Visible = False
'            DtpDbNtDt.Visible = False
'            ChkIntRateMonthly.Visible = False
'            Label4.Visible = False
'            Label1(3).Caption = "From Inv. No :"
'            Label1(1).Caption = "To Inv. No :"
'            txt(3).Enabled = True
'            txt(4).Enabled = True
        Case "Receipt Against Invoice Detail"
            txt(0).Enabled = True '-party
            ChkConsolidate.Enabled = True
            txt(1).Enabled = True '-Int %
            txt(2).Enabled = True '- Grase Days
    End Select
'-- Set InActivate Ctrl Back Color
GProcSetInActivateCtrlBackClr Me
End Sub

Private Sub txt_GotFocus(Index As Integer)
GProcSelectBox Me.ActiveControl
End Sub

Private Sub txt_KeyPress(Index As Integer, KeyAscii As Integer)
If KeyAscii = 13 Then Exit Sub '-- Enter
Select Case Index
    Case 0
        If gRptBrokerFlag = True Then '-- Broker
'            gClsSearch.SearchMultiField "QryMastaccount", "Account,AcName,ACCode", Array("Broker", "AcName", ""), Array(txt(0).Width, 0, 0), " (AgCode=90016  or GpCode=90016)", Chr(KeyAscii), "AcName", txt(Index).Left + Me.Left, txt(Index).Top + Me.Top + 1000, False
            gClsSearch.SearchMultiFieldAccMaster Chr(KeyAscii), txt(0).Left + Me.Left, txt(0).Top + Me.Top + 1000
        Else
            If gRptMillFlag = True Then '-mill
'                gClsSearch.SearchMultiField "QryMastaccount", "Account,AcName,ACCode", Array("Account Name", "AcName", ""), Array(txt(0).Width, 0, 0), " (AgCode=90029 or GpCode=90029)", Chr(KeyAscii), "AcName", txt(Index).Left + Me.Left, txt(Index).Top + Me.Top + 1000, False
                gClsSearch.SearchMultiFieldAccMaster Chr(KeyAscii), txt(0).Left + Me.Left, txt(0).Top + Me.Top + 1000
            Else '-Party
'                gClsSearch.SearchMultiField "QryMastaccount", "Account,AcName,ACCode", Array("Account Name", "AcName", ""), Array(txt(0).Width, 0, 0), "", Chr(KeyAscii), "AcName", txt(Index).Left + Me.Left, txt(Index).Top + Me.Top + 1000, False
                gClsSearch.SearchMultiFieldAccMaster Chr(KeyAscii), txt(0).Left + Me.Left, txt(0).Top + Me.Top + 1000
            End If
        End If
        If gClsSearch.SearchMultiRetColAccMaster(0) <> "" Then
'            txt(Index) = gClsSearch.SearchMultiRetCol(1)
            txt(Index) = gClsSearch.SearchMultiRetColAccMaster(0)
            KeyAscii = 0
        End If
     Case 1, 2
         KeyAscii = GProcNumberOnly(KeyAscii)
     Case 4
        gClsSearch.SearchMultiField "tblMastaccount", "AcName,AcOurCode,AcCity,ACCode", Array("Broker Name", "Code", "City", "Master Code"), Array(8000, 1500, 2000, 0), "(AgCode=90016 ) ", Chr(KeyAscii), "AcName", txt(Index).Left + Me.Left, txt(Index).Top + Me.Top + 1000
        KeyAscii = 0
        If gClsSearch.SearchMultiRetCol(0) <> "" Then
            txt(Index) = gClsSearch.SearchMultiRetCol(0)
            KeyAscii = 0
        End If
End Select
End Sub

