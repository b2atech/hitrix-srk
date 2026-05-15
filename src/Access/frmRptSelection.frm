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
   ClientWidth     =   6945
   Icon            =   "frmRptSelection.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5835
   ScaleWidth      =   6945
   Begin VB.CheckBox ChkIntDbNtShort 
      BackColor       =   &H80000018&
      Caption         =   "Is Interest Demand Note Short ?"
      ForeColor       =   &H00000080&
      Height          =   330
      Left            =   3915
      TabIndex        =   27
      Top             =   1170
      Width           =   2595
   End
   Begin VB.CheckBox ChkIntRateMonthly 
      BackColor       =   &H80000018&
      Caption         =   "Is Interest Rate Monthly ?"
      ForeColor       =   &H00000080&
      Height          =   330
      Left            =   3435
      TabIndex        =   13
      Top             =   2625
      Width           =   3015
   End
   Begin VB.CommandButton cmdDbNt 
      Caption         =   "&Cancel Interest Debit Note"
      Height          =   480
      Index           =   1
      Left            =   1605
      TabIndex        =   26
      Top             =   5085
      Width           =   1290
   End
   Begin VB.CommandButton cmdCrystal 
      Height          =   480
      Index           =   1
      Left            =   4245
      Picture         =   "frmRptSelection.frx":058A
      Style           =   1  'Graphical
      TabIndex        =   25
      ToolTipText     =   "Print"
      Top             =   5085
      Width           =   1185
   End
   Begin VB.CommandButton cmdCrystal 
      Height          =   480
      Index           =   0
      Left            =   3000
      Picture         =   "frmRptSelection.frx":0BF4
      Style           =   1  'Graphical
      TabIndex        =   24
      Top             =   5085
      Width           =   1185
   End
   Begin VB.CheckBox ChkIntLateFrmInvDt 
      BackColor       =   &H80000018&
      Caption         =   "Is Late Days From Invoice Date ?"
      ForeColor       =   &H00000080&
      Height          =   330
      Left            =   3435
      TabIndex        =   12
      Top             =   2130
      Width           =   3015
   End
   Begin VB.CommandButton cmdDbNt 
      Caption         =   "&Apply Interest Debit Note"
      Height          =   480
      Index           =   0
      Left            =   180
      TabIndex        =   21
      Top             =   5085
      Width           =   1305
   End
   Begin VB.TextBox txt 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   3
      Left            =   1680
      TabIndex        =   18
      Text            =   "3"
      Top             =   3825
      Width           =   1575
   End
   Begin VB.TextBox txt 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   4
      Left            =   4950
      TabIndex        =   20
      Text            =   "4"
      Top             =   3825
      Width           =   1575
   End
   Begin VB.CheckBox ChkConsolidate 
      BackColor       =   &H80000018&
      Caption         =   "Is &Consolidate Report ?"
      ForeColor       =   &H00000080&
      Height          =   330
      Left            =   1680
      TabIndex        =   6
      Top             =   1155
      Width           =   2100
   End
   Begin VB.TextBox txt 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   2
      Left            =   4950
      TabIndex        =   11
      Text            =   "2"
      Top             =   1710
      Width           =   1575
   End
   Begin VB.TextBox txt 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   1
      Left            =   1680
      TabIndex        =   9
      Text            =   "1"
      Top             =   2130
      Width           =   1575
   End
   Begin VB.TextBox txt 
      Height          =   375
      Index           =   0
      Left            =   1680
      TabIndex        =   5
      Top             =   690
      Width           =   4845
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
      Left            =   5505
      TabIndex        =   22
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
      TabIndex        =   23
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
      Left            =   4950
      TabIndex        =   3
      Top             =   225
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   635
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      CustomFormat    =   "dd/MM/yyyy"
      Format          =   122748931
      CurrentDate     =   38050
   End
   Begin MSComCtl2.DTPicker dtpFrDt 
      Height          =   360
      Left            =   1680
      TabIndex        =   1
      Top             =   225
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   635
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      CustomFormat    =   "dd/MM/yyyy"
      Format          =   122748931
      CurrentDate     =   38050
   End
   Begin MSComCtl2.DTPicker DtpDbNtDt 
      Height          =   360
      Left            =   1680
      TabIndex        =   15
      Top             =   2625
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   635
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      CustomFormat    =   "dd/MM/yyyy"
      Format          =   122748931
      CurrentDate     =   38050
   End
   Begin VB.Shape Shape2 
      FillColor       =   &H00BDDFD1&
      FillStyle       =   0  'Solid
      Height          =   765
      Left            =   75
      Shape           =   4  'Rounded Rectangle
      Top             =   4950
      Width           =   6780
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Db&Nt Date   :"
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
      Left            =   180
      TabIndex        =   14
      Top             =   2625
      Width           =   1140
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "&O/S Aging Report"
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
      Left            =   180
      TabIndex        =   16
      Top             =   3375
      Width           =   2220
   End
   Begin VB.Line Line2 
      BorderColor     =   &H000040C0&
      X1              =   75
      X2              =   6810
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
      Left            =   180
      TabIndex        =   7
      Top             =   1785
      Width           =   2595
   End
   Begin VB.Line Line1 
      BorderColor     =   &H000040C0&
      X1              =   60
      X2              =   6765
      Y1              =   1590
      Y2              =   1590
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Due Days From :"
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
      Index           =   3
      Left            =   180
      TabIndex        =   17
      Top             =   3825
      Width           =   1485
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Due Days To :"
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
      Left            =   3435
      TabIndex        =   19
      Top             =   3825
      Width           =   1245
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
      Left            =   3435
      TabIndex        =   10
      Top             =   1710
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
      Left            =   180
      TabIndex        =   8
      Top             =   2130
      Width           =   1185
   End
   Begin VB.Label lblName 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "&Mill Name       :"
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
      Left            =   180
      TabIndex        =   4
      Top             =   690
      Width           =   1350
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
      TabIndex        =   2
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
      Left            =   180
      TabIndex        =   0
      Top             =   225
      Width           =   1350
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00000080&
      FillColor       =   &H00EDFDFE&
      FillStyle       =   0  'Solid
      Height          =   4770
      Left            =   45
      Top             =   60
      Width           =   6780
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
         Case "Late Payment Interest Register" '- "Late Payment Interest Register"
             If CryIndex = 2 And ChkConsolidate = 1 Then
                MsgBox "Select Single Firm Selection For Interest Debit Note.", vbOKOnly + vbInformation, Me.Caption
                ChkConsolidate.SetFocus
                Exit Function
            End If
    End Select
ValidataData = True
End Function
Private Sub cmdCrystal_Click(Index As Integer)
Dim X1 As String
Dim FromDt As Date
Dim ToDt As Date
FromDt = Format(dtpFrDt.Value, "dd/MM/yyyy")
ToDt = Format(dtpToDt.Value, "dd/MM/yyyy")
Dim RsP As Recordset
Dim RsTmp As Recordset
Dim Ramt As Double

'--- Validata Data
If ValidataData(Index) = False Then Exit Sub
GProcCrystalRptPreparation CrystalReport1, "Window"
With CrystalReport1
    If Index = 1 Then .Destination = crptToPrinter
    Select Case gSelectedMenu
        '- "Broker & Partywise Aging Report"
        Case "Broker & Partywise Aging Report"
            GProcCollectCommonOutstanding FromDt, FromDt, "", txt(0)
            GProcPrepareAging FromDt, txt(0)
            If gBackEndDB = gBackEndAccess Then
                .DataFiles(0) = gDbLocation
            Else '-- Oracle
                .Connect = "DSN=" + gODBCDSN + " ;UID=Scott;PWD=" + gODBCPwd + ";DBQ=" + gODBCHostStr
            End If
            .Formulas(4) = "wFrToDt = ' On Date: ' & '" & FromDt & "'"
            If ChkConsolidate.Value = 1 Then
                .ReportFileName = gReportPath & "rptOutStandingAgingBank.rpt"
            Else
                .ReportFileName = gReportPath & "rptOutStandingAging.rpt"
            End If
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
            .SubreportToChange = "rptSubOutStandingAging.rpt"
            If gBackEndDB = gBackEndAccess Then
                .DataFiles(0) = gDbLocation
            Else '-- Oracle
                .Connect = "DSN=" + gODBCDSN + " ;UID=Scott;PWD=" + gODBCPwd + ";DBQ=" + gODBCHostStr
            End If
        Case "Late Payment Interest Purchase List"
             GProcOutStandingSALatePayIntPurch FromDt, ToDt, ChkConsolidate, txt(1), txt(2), txt(0), ChkIntLateFrmInvDt, IIf(Index = 2, True, False)
             .Formulas(4) = "wFrToDt = 'From : ' & '" & Format(dtpFrDt.Value, "dd/MM/yyyy") & "' & '  To ' & '" & Format(dtpToDt.Value, "dd/MM/yyyy") & "'"
             If ChkIntLateFrmInvDt = 1 Then '-- Late Days From
                .Formulas(5) = "LateDaysFromStr = 'Late Days From Invoice Date'"
             Else
                .Formulas(5) = "LateDaysFromStr = 'Late Days From Due Date'"
             End If
             '-- Is Interest Rate Monthly
             If ChkIntRateMonthly = 1 Then
                .Formulas(8) = "IsIntRateMonthly ='1'"
             Else
                .Formulas(8) = "IsIntRateMonthly ='0'"
             End If
             .DataFiles(0) = gDataPath & "\Setup.mdb"
             gReportPath = App.Path & "\HiReports\"
             .ReportFileName = gReportPath & "rptOutStandingSALatePayIntPurch.rpt"
             .SelectionFormula = "{tmpGentbl.UserName}='" & gUserName & "' and {@Balinterest}>0"
        '-- Late Payment Interest Register
        Case "Late Payment Interest Register", "Late Payment Interest Register (Mill Bill)"
             If gSelectedMenu = "Late Payment Interest Register" Then
                GProcOutStandingSALatePayInt FromDt, ToDt, ChkConsolidate, txt(1), txt(2), txt(0), ChkIntLateFrmInvDt, IIf(Index = 2, True, False)
             Else
                GProcOutStandingSALatePayInt FromDt, ToDt, ChkConsolidate, txt(1), txt(2), txt(0), ChkIntLateFrmInvDt, IIf(Index = 2, True, False), True
             End If
             .Formulas(4) = "wFrToDt = 'From : ' & '" & Format(dtpFrDt.Value, "dd/MM/yyyy") & "' & '  To ' & '" & Format(dtpToDt.Value, "dd/MM/yyyy") & "'"
             If ChkIntLateFrmInvDt = 1 Then '-- Late Days From
                .Formulas(5) = "LateDaysFromStr = 'Late Days From Invoice Date'"
             Else
                .Formulas(5) = "LateDaysFromStr = 'Late Days From Due Date'"
             End If
             '-- Is Interest Rate Monthly
             If ChkIntRateMonthly = 1 Then
                .Formulas(8) = "IsIntRateMonthly ='1'"
             Else
                .Formulas(8) = "IsIntRateMonthly ='0'"
             End If
             .DataFiles(0) = gDataPath & "\Setup.mdb"
             gReportPath = App.Path & "\HiReports\"
             .ReportFileName = gReportPath & "rptOutStandingSALatePayInt.rpt"
             .SelectionFormula = "{tmpGentbl.UserName}='" & gUserName & "' and {@Balinterest}>0"
        '-Interest Demand Note
        Case "Interest Demand Note"
            .Formulas(4) = "wFrToDt = 'From : ' & '" & Format(dtpFrDt.Value, "dd/MM/yyyy") & "' & '  To ' & '" & Format(dtpToDt.Value, "dd/MM/yyyy") & "'"
            '- Vtype ,Vdt
            X1 = " ({tblRecVsSale.RECVTYPE}='BR' or {tblRecVsSale.RECVTYPE}='CR' ) and {tblRecVsSale.RECVDT} in Date(" & Format(dtpFrDt.Value, "yyyy,mm,dd") & ") to Date(" & Format(dtpToDt.Value, "yyyy,mm,dd") & ")  " '
            If txt(0) <> "" Then X1 = X1 + " and {tblMastAccount.AcName} = '" & txt(0) & "'" '-Party
            X1 = X1 + " and {@Balinterest}> 0 "
            .SelectionFormula = X1
            If ChkIntDbNtShort = 1 Then '-- Short
                .ReportFileName = gReportPath & "rptAccInterestDbNtNoteShort.rpt"
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
'                .SelectionFormula = "{tmpGentbl.UserName}='" & gUserName & "'  and {@balBag}<>0 and {@BalWt}<>0"
                .SelectionFormula = "{tmpGentbl.UserName}='" & gUserName & "' "
                If txt(0) <> "" Then
                     .SelectionFormula = .SelectionFormula & " and {tmpGentbl.nar2}='" & txt(0) & "'  and {@balBag}<>0 and {@BalWt}<>0"
                End If
                If ChkConsolidate.Value = 1 Then
                   .ReportFileName = gReportPath & "rptStkRptDateCountwise.rpt"
                Else
                   .ReportFileName = gReportPath & "rptStkRptDatewise.rpt"
                End If
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
    Case "O/s In Date Range"
          X1 = "select Vno,Vtype,Vyear,OutBillNo,Vdt,OutDueDate,OutAmount,tblMastAccount.Acname as Party,tblMastAccount_Br.Acname as Brok,'" & gCCode & "' As MillFirm ,OutItShort,OutBag,tblMastAccount_Br.AcFaxNo as BrkPh " _
               & " from tblOutstanding,tblMastAccount,tblMastAccount tblMastAccount_Br " _
               & " where vtype = 'SY' and Vdt >=cdate('" & Format(dtpFrDt.Value, "dd/MM/yyyy") & "') and vdt <= cdate('" & Format(dtpToDt.Value, "dd/MM/yyyy") & "') " _
               & " and OutAcCode = tblMastAccount.AcCode and OutBrokerCode = tblMastAccount_Br.AcCode"
          gCn.Execute "delete from tmpGenTbl2 where username='" & gUserName & "'"
          If (GProcRstOpen(RsP, X1, "R") > 0) Then
                 RsP.MoveFirst
                 Do While Not RsP.EOF '-issue bag,Wt
                    Ramt = 0
                    If (GProcRstOpen(RsTmp, "select sum(AdjAmt+Discount) from tblRecVsSale where RecVDt >=cdate('" & Format(dtpFrDt.Value, "dd/MM/yyyy") & "') and RecVDt <= cdate('" & Format(dtpToDt.Value, "dd/MM/yyyy") & "')  and InvTp='" & RsP!Vtype & "' and InvYear= " & RsP!VYear & " and InvNo= " & RsP!Vno & "", "R") > 0) Then
                       Ramt = IIf(IsNumeric(RsTmp.Fields(0)), RsTmp.Fields(0), 0)
                    End If
                    If RsP.Fields("OutAmount") - Ramt > 0 Then
                       gCn.Execute "insert into tmpgentbl2 (Username,VBILLNO,VBILLDT,Vdt,Amt1,Nar1,Nar2,Amt2,Nar3,Nar4,Amt3,Nar5)  select '" & gUserName & "','" & RsP.Fields("OutBillNo") & "','" & RsP.Fields("Vdt") & "','" & RsP.Fields("OutDueDate") & "'," & RsP.Fields("OutAmount") & ",'" & RsP.Fields("Party") & "','" & RsP.Fields("Brok") & "'," & Ramt & ", " _
                                   & " '" & RsP.Fields("MillFirm") & "','" & RsP.Fields("OutItShort") & "'," & RsP.Fields("OutBag") & ",'" & RsP.Fields("BrkPh") & "'"
                    End If
                    RsP.MoveNext
                Loop
          
          End If
          .Formulas(4) = "wDateFrom =  '" & Format(dtpFrDt.Value, "dd/MM/yyyy") & "'"
          .Formulas(5) = "wDateTo =  '" & Format(dtpToDt.Value, "dd/MM/yyyy") & "'"
          X1 = " {tmpGenTbl2.Username}='" & gUserName & "'"
          .SelectionFormula = X1
          .ReportFileName = gReportPath & "rptOsDateRangerpt.rpt"
          
          
        Case "Countwise Profit and Loss"
'             Dim Rsp As Recordset
'             Dim RsTmp As Recordset
             Dim X2 As String
             Dim X3 As String
             gCn.Execute "Delete from tmpGenTbl where USERNAME='" & gUserName & "'"
             '---Op.Stock
             X1 = " Vdt<cdate('" & dtpFrDt.Value & "')"
             X2 = " Vdt<cdate('" & dtpFrDt.Value & "')"
             X3 = " Vdt<=cdate('" & dtpToDt.Value & "')"
             X1 = X1 + " and vtype in ('PI','PY','PT','OM','OS','OP','OI') "
             '--- Update to p It Ctrl No Stock
             gCn.Execute "Update tblVousub set VSubTmpBag = 0, VSubTmpWt = 0 where " & X1
             If (GProcRstOpen(RsP, "select * from tblVousub where " & X1, "O") > 0) Then
                 RsP.MoveFirst
                 Do While Not RsP.EOF '-issue bag,Wt
                 If (GProcRstOpen(RsTmp, "select sum(VSUBBAG),sum(VSUBWt) from tblVousub where " & X2 & "  and VSUBPITCTRLNO=" & RsP.Fields("VSUBITCTRLNO"), "O") > 0) Then
                    RsP.Fields("vsubissbag") = IIf(IsNumeric(RsTmp.Fields(0)), RsTmp.Fields(0), 0)
                    RsP.Fields("vsubisswt") = IIf(IsNumeric(RsTmp.Fields(1)), RsTmp.Fields(1), 0)
                    RsP.Update
                End If
                RsP.MoveNext
                Loop
             End If
             
             If (GProcRstOpen(RsP, "select * from tblVousub where " & X1, "O") > 0) Then
                 RsP.MoveFirst
                 Do While Not RsP.EOF '-issue bag,Wt
                 If (GProcRstOpen(RsTmp, "select sum(VSUBBAG),sum(VSUBWt) from tblVousub where " & X3 & "  and VSUBPITCTRLNO=" & RsP.Fields("VSUBITCTRLNO"), "O") > 0) Then
                    RsP.Fields("VSubTmpWt") = IIf(IsNumeric(RsTmp.Fields(1)), RsTmp.Fields(1), 0)
                    RsP.Update
                End If
                RsP.MoveNext
                Loop
             End If
             RsP.Close
             '-- update to 0 if null
             gCn.Execute "update tblVousub set vsubissbag=0 where vsubissbag is null and " & X1
             gCn.Execute "update tblVousub set vsubisswt=0 where vsubisswt is null and " & X1
             
             X1 = " tblVousub.Vdt<cdate('" & dtpFrDt.Value & "')"
                
'---add Less On Bag
             gCn.Execute "insert into tmpGenTbl (vtype,VNO,VAcCode,Nar1,Nar3,Nar2,UserName,Amt1,Amt2,Amt7,Amt10,Amt12,Amt13,Nar6) select tblAddLess.vtype,VSubItCtrlNo,VsubItCode,ItName,ItTicket,tblMastNarration.Narration,'" & gUserName & "'," _
             & "sum(vsubbag-TBLVOUSUB.VSubIssBag),sum(vsubwt-VSubIssWt)," _
             & "((iif(tblMastNarration.Narration = 'V A T',sum(VsubAmt+((((AdBillAmt-AdTaxAmt-AdSubAmt))/AdTotBag)* (vsubBag-VSubIssBag))),sum(VsubAmt+((((AdBillAmt-AdSubAmt))/AdTotBag)* (vsubBag-VSubIssBag)))))/(sum(vsubwt)))*sum(vsubwt-VSubIssWt)" _
             & ",sum((AdLCostFreight/AdTotBag)*(vsubbag-TBLVOUSUB.VSubIssBag)), " _
             & "(iif(tblMastNarration.Narration = 'V A T', sum(VsubAmt+((((AdBillAmt-AdTaxAmt-AdSubAmt))/AdTotBag)* (vsubBag))),sum(VsubAmt+((((AdBillAmt-AdSubAmt))/AdTotBag)* (vsubBag)))))/(sum(vsubwt)),sum(vsubwt-VSubTmpWt) " _
             & ", iif(tblAddLess.Vtype = 'PY',' Trade','Sale In Transit') " _
             & " from tblVousub,tblMastItem,tblAddLess,tblMastNarration,tblMastAccount " _
             & " where tblVouSub.vtype in ('PY','PT') and tblVouSub.VsubItCode = tblMastItem.ItCode and" _
             & " tblAddLess.Vno = tblVousub.Vno and tblAddLess.Vtype = tblVousub.Vtype and tblAddLess.Vyear = tblVousub.Vyear" _
             & " and tblAddLess.adTaxcode = tblMastNarration.NarrCode and (TBLVOUSUB.VSubBag-TBLVOUSUB.VSubIssBag)>0 " _
             & " and tblMastItem.ItMillCode = tblMastAccount.AcCode and tblAddLess.Vdt < CDate('01/07/2017') and " _
             & " " & X1 & " group by tblAddless.vtype,VsubItCode,ItName,ItTicket,VSubItCtrlNo,tblMastNarration.Narration"
                
             '--gst
'+sum((AdLCostFreight/AdTotBag)*(vsubBag-VSubIssBag))
             
             gCn.Execute "insert into tmpGenTbl (vtype,VNO,VAcCode,Nar1,Nar3,Nar2,UserName,Amt1,Amt2,Amt7,Amt10,Amt12,Amt13,Nar6) select tblAddLess.vtype,VSubItCtrlNo,VsubItCode,ItName,ItTicket,tblMastNarration.Narration,'" & gUserName & "'," _
             & "sum(vsubbag-TBLVOUSUB.VSubIssBag),sum(vsubwt-VSubIssWt)," _
             & "sum((VSubTaxableAmt/vsubBag)*(vsubBag-VSubIssBag))+sum((AdLCostFreight/AdTotBag)*(vsubBag-VSubIssBag))" _
             & ",0, " _
             & "sum(VSubTaxableAmt)/sum(VSubWt),sum(vsubwt-VSubTmpWt) " _
             & ", iif(tblAddLess.Vtype = 'PY',' Trade','Sale In Transit') " _
             & " from tblVousub,tblMastItem,tblAddLess,tblMastNarration,tblMastAccount " _
             & " where tblVouSub.vtype in ('PY','PT') and tblVouSub.VsubItCode = tblMastItem.ItCode and" _
             & " tblAddLess.Vno = tblVousub.Vno and tblAddLess.Vtype = tblVousub.Vtype and tblAddLess.Vyear = tblVousub.Vyear" _
             & " and tblAddLess.adTaxcode = tblMastNarration.NarrCode and (TBLVOUSUB.VSubBag-TBLVOUSUB.VSubIssBag)>0 " _
             & " and tblMastItem.ItMillCode = tblMastAccount.AcCode and tblAddLess.Vdt >= CDate('01/07/2017') and tblMastItem.ItOtherItem = 0 and " _
             & " " & X1 & " group by tblAddless.vtype,VsubItCode,ItName,ItTicket,VSubItCtrlNo,tblMastNarration.Narration"
             
             
'             gCn.Execute "insert into tmpGenTbl (vtype,VNO,VAcCode,Nar1,Nar3,Nar2,UserName,Amt1,Amt2,Amt7,Amt10,Amt12,Amt13,Nar6) select tblAddLess.vtype,VSubItCtrlNo,VsubItCode,ItName,ItTicket,tblMastNarration.Narration,'" & gUserName & "'," _
'             & "sum(vsubbag-TBLVOUSUB.VSubIssBag),sum(vsubwt-VSubIssWt)," _
'             & "sum((VSubTaxableAmt/vsubBag)*(vsubBag-VSubIssBag))" _
'             & ",sum((ADTMP2/AdTotBag)*(vsubbag-TBLVOUSUB.VSubIssBag)), " _
'             & "sum(VSubTaxableAmt)/sum(VSubWt),sum(vsubwt-VSubTmpWt) " _
'             & ", iif(tblAddLess.Vtype = 'PY',' Trade','Sale In Transit') " _
'             & " from tblVousub,tblMastItem,tblAddLess,tblMastNarration,tblMastAccount " _
'             & " where tblVouSub.vtype in ('PY','PT') and tblVouSub.VsubItCode = tblMastItem.ItCode and" _
'             & " tblAddLess.Vno = tblVousub.Vno and tblAddLess.Vtype = tblVousub.Vtype and tblAddLess.Vyear = tblVousub.Vyear" _
'             & " and tblAddLess.adTaxcode = tblMastNarration.NarrCode and (TBLVOUSUB.VSubBag-TBLVOUSUB.VSubIssBag)>0 " _
'             & " and tblMastItem.ItMillCode = tblMastAccount.AcCode and tblAddLess.Vdt >= CDate('01/07/2017') and tblMastItem.ItOtherItem = 0 and " _
'             & " " & X1 & " group by tblAddless.vtype,VsubItCode,ItName,ItTicket,VSubItCtrlNo,tblMastNarration.Narration"
             
             
             '-- Inward
             
             X1 = " tblVousub.Vdt>= cdate('" & dtpFrDt.Value & "')  and  tblVousub.Vdt<=cdate('" & dtpToDt.Value & "')"
             X1 = X1 + " and vtype in ('PI','PY','PT','OM','OS','OP','OI') "
             X2 = " Vdt<=cdate('" & dtpToDt.Value & "')"

             gCn.Execute "Update tblVousub set VSubTmpBag = 0, VSubTmpWt = 0 where " & X1
             If (GProcRstOpen(RsP, "select * from tblVousub where " & X1, "O") > 0) Then
                 RsP.MoveFirst
                 Do While Not RsP.EOF '-issue bag,Wt
                 If (GProcRstOpen(RsTmp, "select sum(VSUBBAG),sum(VSUBWt) from tblVousub where " & X2 & "  and VSUBPITCTRLNO=" & RsP.Fields("VSUBITCTRLNO"), "O") > 0) Then
                    RsP.Fields("vsubissbag") = IIf(IsNumeric(RsTmp.Fields(0)), RsTmp.Fields(0), 0)
                    RsP.Fields("vsubisswt") = IIf(IsNumeric(RsTmp.Fields(1)), RsTmp.Fields(1), 0)
                    RsP.Update
                End If
                RsP.MoveNext
                Loop
             End If
             RsP.Close
             '-- update to 0 if null
             gCn.Execute "update tblVousub set vsubissbag=0 where vsubissbag is null and " & X1
             gCn.Execute "update tblVousub set vsubisswt=0 where vsubisswt is null and " & X1

             DoEvents
            
             X1 = " tblVousub.Vdt>= cdate('" & dtpFrDt.Value & "')  and  tblVousub.Vdt<=cdate('" & dtpToDt.Value & "')"
'---add Less On Bag
             gCn.Execute "insert into tmpGenTbl (vtype,VNO,VAcCode,Nar1,Nar3,Nar2,UserName,Amt3,Amt4,Amt9,Amt10,Amt12,Amt13,Nar6) select tblAddLess.vtype,VSubItCtrlNo,VsubItCode,ItName,ItTicket,tblMastNarration.Narration,'" & gUserName & "'," _
             & "sum(vsubbag),sum(vsubwt),iif(tblMastNarration.Narration = 'V A T',sum(VsubAmt+((((AdBillAmt-AdTaxAmt-AdSubAmt))/AdTotBag)* (vsubBag))),sum(VsubAmt+((((AdBillAmt-AdSubAmt))/AdTotBag)* (vsubBag)))),sum((AdLCostFreight/AdTotBag)*(vsubbag))," _
             & "(iif(tblMastNarration.Narration = 'V A T',sum(VsubAmt+((((AdBillAmt-AdTaxAmt-AdSubAmt))/AdTotBag)* (vsubBag))),sum(VsubAmt+((((AdBillAmt-AdSubAmt))/AdTotBag)* (vsubBag)))))/(sum(vsubwt)),sum(vsubwt-VSubIssWt)" _
             & ", iif(tblAddLess.Vtype = 'PY',' Trade','Sale In Transit') " _
             & " from tblVousub,tblMastItem,tblAddLess,tblMastNarration,tblMastAccount " _
             & " where tblVouSub.vtype in ('PY','PT') and tblVouSub.VsubItCode = tblMastItem.ItCode and" _
             & " tblAddLess.Vno = tblVousub.Vno and tblAddLess.Vtype = tblVousub.Vtype and tblAddLess.Vyear = tblVousub.Vyear" _
             & " and tblAddLess.adTaxcode = tblMastNarration.NarrCode " _
             & " and tblMastItem.ItMillCode = tblMastAccount.AcCode and tblAddLess.Vdt < CDate('01/07/2017') and" _
             & " " & X1 & " group by tblAddless.Vtype,VsubItCode,ItName,ItTicket,VSubItCtrlNo,tblMastNarration.Narration"

             '--GST
    
             gCn.Execute "insert into tmpGenTbl (vtype,VNO,VAcCode,Nar1,Nar3,Nar2,UserName,Amt3,Amt4,Amt9,Amt10,Amt12,Amt13,Nar6) select tblAddLess.vtype,VSubItCtrlNo,VsubItCode,ItName,ItTicket,tblMastNarration.Narration,'" & gUserName & "'," _
             & "sum(vsubbag),sum(vsubwt),sum(VSubTaxableAmt+((AdLCostFreight/AdTotBag)*vsubbag)),sum((ADTMP2/AdTotBag)*(vsubbag))," _
             & "sum(VSubTaxableAmt/vsubwt),sum(vsubwt-VSubIssWt)" _
             & ", iif(tblAddLess.Vtype = 'PY',' Trade','Sale In Transit') " _
             & " from tblVousub,tblMastItem,tblAddLess,tblMastNarration,tblMastAccount " _
             & " where tblVouSub.vtype in ('PY','PT') and tblVouSub.VsubItCode = tblMastItem.ItCode and" _
             & " tblAddLess.Vno = tblVousub.Vno and tblAddLess.Vtype = tblVousub.Vtype and tblAddLess.Vyear = tblVousub.Vyear" _
             & " and tblAddLess.adTaxcode = tblMastNarration.NarrCode " _
             & " and tblMastItem.ItMillCode = tblMastAccount.AcCode and tblAddLess.Vdt >= CDate('01/07/2017') and tblMastItem.ItOtherItem = 0  and" _
             & " " & X1 & " group by tblAddless.Vtype,VsubItCode,ItName,ItTicket,VSubItCtrlNo,tblMastNarration.Narration"

             '--Freight & Insu

             gCn.Execute "insert into tmpGenTbl (vtype,VNO,VAcCode,Nar1,Nar3,Nar2,UserName,Amt3,Amt4,Amt9,Amt10,Amt12,Amt13,Nar6) select tblAddLess.vtype,VSubItCtrlNo,VsubItCode,ItName,ItTicket,tblMastNarration.Narration,'" & gUserName & "'," _
             & "0,0,0,sum(VSubTaxableAmt)," _
             & "0,0" _
             & ", iif(tblAddLess.Vtype = 'PY',' Trade','Sale In Transit') " _
             & " from tblVousub,tblMastItem,tblAddLess,tblMastNarration,tblMastAccount " _
             & " where tblVouSub.vtype in ('PY','PT') and tblVouSub.VsubItCode = tblMastItem.ItCode and" _
             & " tblAddLess.Vno = tblVousub.Vno and tblAddLess.Vtype = tblVousub.Vtype and tblAddLess.Vyear = tblVousub.Vyear" _
             & " and tblAddLess.adTaxcode = tblMastNarration.NarrCode " _
             & " and tblMastItem.ItMillCode = tblMastAccount.AcCode and tblAddLess.Vdt >= CDate('01/07/2017') and tblMastItem.ItOtherItem = 1  and" _
             & " " & X1 & " group by tblAddless.Vtype,VsubItCode,ItName,ItTicket,VSubItCtrlNo,tblMastNarration.Narration"




             '--- Outward
             
             X1 = " tblAddless.Vdt>= cdate('" & dtpFrDt.Value & "')  and  tblAddless.Vdt<=cdate('" & dtpToDt.Value & "')"
              gCn.Execute "insert into tmpGenTbl (vtype,VAcCode,Nar1,Nar3,Nar4,UserName,Amt5,Amt6,Amt11,Nar6) select tblAddLess.vtype,VsubItCode,ItName,ItTicket,tblMastNarration.Narration,'" & gUserName & "'," _
                & "sum(vsubbag),sum(vsubwt),iif(tblMastNarration.Narration = 'V A T',sum(AdBillAmt-AdTaxAmt),iif(tblMastNarration.Narration = 'C S T',sum(AdBillAmt-AdTaxAmt),sum(AdBillAmt))) " _
                & ",iif(tblAddLess.Vtype = 'SY',' Trade','Sale In Transit')   from tblVousub,tblMastItem,tblAddLess,tblMastNarration,tblMastAccount " _
                & " where tblAddLess.vtype in ('SY','ST') and " _
                & " tblAddLess.Vno = tblVousub.Vno and tblAddLess.Vtype = tblVousub.Vtype and tblAddLess.Vyear = tblVousub.Vyear" _
                & " and tblVouSub.VsubItCode = tblMastItem.ItCode and tblAddLess.adTaxcode = tblMastNarration.NarrCode  " _
                & " and tblMastItem.ItMillCode = tblMastAccount.AcCode and tblAddLess.Vdt < CDate('01/07/2017')  " _
                & "  and " & X1 & " group by tblAddLess.Vtype,VsubItCode,ItName,ItTicket,tblMastNarration.Narration"

             '-Gst

             X1 = " tblAddless.Vdt>= cdate('" & dtpFrDt.Value & "')  and  tblAddless.Vdt<=cdate('" & dtpToDt.Value & "')"
              gCn.Execute "insert into tmpGenTbl (vtype,VAcCode,Nar1,Nar3,Nar4,UserName,Amt5,Amt6,Amt11,Nar6) select tblAddLess.vtype,VsubItCode,ItName,ItTicket,tblMastNarration.Narration,'" & gUserName & "'," _
                & "sum(vsubbag),sum(vsubwt),sum(AdBillAmt-(VsubCGSTAmt+VsubSGSTAmt+VsubIGSTAmt)) " _
                & ",iif(tblAddLess.Vtype = 'SY',' Trade','Sale In Transit')   from tblVousub,tblMastItem,tblAddLess,tblMastNarration,tblMastAccount " _
                & " where tblAddLess.vtype in ('SY','ST') and " _
                & " tblAddLess.Vno = tblVousub.Vno and tblAddLess.Vtype = tblVousub.Vtype and tblAddLess.Vyear = tblVousub.Vyear" _
                & " and tblVouSub.VsubItCode = tblMastItem.ItCode and tblAddLess.adTaxcode = tblMastNarration.NarrCode  " _
                & " and tblMastItem.ItMillCode = tblMastAccount.AcCode and tblAddLess.Vdt >= CDate('01/07/2017') " _
                & "  and " & X1 & " group by tblAddLess.Vtype,VsubItCode,ItName,ItTicket,tblMastNarration.Narration"
                
                
                
                
                If gBackEndDB = gBackEndAccess Then
                   .DataFiles(0) = gDbLocation
                Else '-- Oracle
                    .Connect = "DSN=" + gODBCDSN + " ;UID=Scott;PWD=" + gODBCPwd + ";DBQ=" + gODBCHostStr
                End If
'                .SelectionFormula = "{tmpGentbl.UserName}='" & gUserName & "'  and {@balBag}<>0 and {@BalWt}<>0"
                .SelectionFormula = "{tmpGentbl.UserName}='" & gUserName & "'  "
                
'                .SelectionFormula = "{tmpGentbl.UserName}='" & gUserName & "' and ({tmpGentbl.Vtype}= 'PT' or  {tmpGentbl.Vtype}= 'ST' ) "
                .ReportFileName = gReportPath & "rptCountwiseProfitNew.rpt"
                .Formulas(4) = "wDateFrom =  '" & Format(dtpFrDt.Value, "dd/MM/yyyy") & "'"
                .Formulas(5) = "wDateTo =  '" & Format(dtpToDt.Value, "dd/MM/yyyy") & "'"
    End Select
    .WindowState = crptMaximized
    .Action = 1
    If gSelectedMenu = "Countwise Profit and Loss" Or gSelectedMenu = "Stock Statement Inwardwise Datewise" Then
'             Dim Rsp As Recordset
'             Dim RsTmp As Recordset
             gCn.Execute "Delete from tmpGenTbl where USERNAME='" & gUserName & "'"
             '---Op.Stock
             X1 = " Vdt<=cdate('" & dtpToDt.Value & "')"
             X1 = X1 + " and vtype in ('PI','PY','PT','OM','OS','OP','OI') "
             '--- Update to p It Ctrl No Stock
             gCn.Execute "Update tblVousub set VSubTmpBag = 0, VSubTmpWt = 0 where " & X1
             If (GProcRstOpen(RsP, "select * from tblVousub where " & X1, "O") > 0) Then
                 RsP.MoveFirst
                 Do While Not RsP.EOF '-issue bag,Wt
                 If (GProcRstOpen(RsTmp, "select sum(VSUBBAG),sum(VSUBWt) from tblVousub where VSUBPITCTRLNO=" & RsP.Fields("VSUBITCTRLNO"), "O") > 0) Then
                    RsP.Fields("vsubissbag") = IIf(IsNumeric(RsTmp.Fields(0)), RsTmp.Fields(0), 0)
                    RsP.Fields("vsubisswt") = IIf(IsNumeric(RsTmp.Fields(1)), RsTmp.Fields(1), 0)
                    RsP.Update
                End If
                RsP.MoveNext
                Loop
             End If
             RsP.Close
             '-- update to 0 if null
             gCn.Execute "update tblVousub set vsubissbag=0 where vsubissbag is null and " & X1
             gCn.Execute "update tblVousub set vsubisswt=0 where vsubisswt is null and " & X1
    End If
End With
End Sub
'=============  Apply , Cancel Interest Debit Note
Private Sub cmdDbNt_Click(Index As Integer)
Dim FromDt As Date
Dim ToDt As Date
'-- Validate
If ChkConsolidate = 1 Then
    MsgBox "Select Single Firm Selection For Interest Debit Note.", vbOKOnly + vbInformation, Me.Caption
    ChkConsolidate.SetFocus
    Exit Sub
End If
FromDt = Format(dtpFrDt.Value, "dd/MM/yyyy")
ToDt = Format(dtpToDt.Value, "dd/MM/yyyy")
If Index = 0 Then '=== Apply Interest Debit Note
    If gSelectedMenu = "Late Payment Interest Register" Then '- "Late Payment Interest Register"
        GProcOutStandingSALatePayInt FromDt, ToDt, ChkConsolidate, txt(1), txt(2), txt(0), ChkIntLateFrmInvDt, True
    Else '- Mill Late Payment Interest Register"
        GProcOutStandingSALatePayInt FromDt, ToDt, ChkConsolidate, txt(1), txt(2), txt(0), ChkIntLateFrmInvDt, True, True
    End If
    ProcPrepareIntDbNt DtpDbNtDt.Value, ChkIntLateFrmInvDt, CLng(txt(2)), CDbl(txt(1)), ChkIntRateMonthly
Else '=== Cancel Interest Debit Note
    If gSelectedMenu = "Late Payment Interest Register" Then '- "Late Payment Interest Register" then
        ProcPrepareCancelIntDbNt FromDt, ToDt, txt(0)
    Else '-Mill Bill
        ProcPrepareCancelIntDbNt FromDt, ToDt, txt(0), True
    End If
End If
End Sub
Private Sub cmdExit_Click()
    Unload Me
End Sub
'--- Procedure Prepare Cancel Interest Debit Note
Private Sub ProcPrepareCancelIntDbNt(FromDt As Date, ToDt As Date, Optional Party As String, Optional IsMillBillRec As Boolean)
Dim CondStr As String
Dim X As String
Dim SmsString As String
Dim PtyMbl As String
Dim RecAdjNo As Long
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
    CondStr = CondStr & " and IntDbNtDate is Not Null" '-IntDbNtDt is Not Null
    X = " select RecVdt,ADJNO,AcName,InvBillNo,InvDt,BILLAMT,INTDBNTAMT,INTDBNTDATE,INTDBNTRATE,INTDBNTGRASEDAY,INTDBNTLATEDAY, AdjAmt,BILLAMT-AdjAmt" _
    & " from tblmastaccount,TblRecVsSale where accode=PartyCode and " & CondStr
    '& " And vno=Invno and vtype=Invtp and vYear=InvYear "
   
    gCn.Execute "Delete from tmpGenTbl2 where  USERNAME='" & gUserName & "'"
   
   
   gClsSearch.SearchMultiSelField X, "", Array("Recd .Dt", "AdjNo", "Party", "Bill No", "Bill Dt", "Bill Amt", "DbNt Amt", "DbNt Dt.", "DbNt Rate", "Grase Day", "Late Day", "Recd. Amt", "Bal Recd"), Array(1400, 0, 2500, 1000, 1400, 1200, 1200, 1400, 1200, 1000, 1100, 1200, 1300), OrderCol:="AcName,RecVdt,InvDt,InvBillNo", IstblNameAsQry:=True, IsReportFlag:=True, ReportCaption:="Cancel Late Pay Interest Debit Note"
    With gClsSearch
        If .SearchMultiSelRetCol(0, 0) <> "" Then
            If MsgBox("Interest Debit Note will Get Cancelled for Selected Receipt. Are You Sure ?", vbYesNo + vbInformation, "Cancel Late Pay Interest Debit Note") = vbYes Then
                For i = 0 To UBound(.SearchMultiSelRetCol)
                    RecAdjNo = Trim(.SearchMultiSelRetCol(i, 1)) '-Rec Adj No
                    
                     SmsString = "Firm   :" & vbLf
                     SmsString = SmsString + "*" + gCName + "*" & vbLf & vbLf
                     SmsString = SmsString + "DELETED INT. DB. NOTE" & vbLf & vbLf
                     SmsString = SmsString + "Party     : *" + .SearchMultiSelRetCol(i, 2) + "*" & vbLf
                     SmsString = SmsString + "Bill No   : " + .SearchMultiSelRetCol(i, 3) & vbLf
                     SmsString = SmsString + "Bill Dt.  : " + .SearchMultiSelRetCol(i, 4) & vbLf
                     SmsString = SmsString + "Rec. Dt.  : " + .SearchMultiSelRetCol(i, 0) & vbLf
                     SmsString = SmsString + "Rec. Amt. : " + Format(.SearchMultiSelRetCol(i, 11), "#########0.00") & vbLf
                     SmsString = SmsString + "Db.Nt.Amt.: *" + Format(.SearchMultiSelRetCol(i, 6), "#########0.00") + "*" & vbLf
                     SmsString = SmsString + "Db.Nt.Dt. : " + .SearchMultiSelRetCol(i, 7) & vbLf
                     SmsString = SmsString + "Int. Rate : " + Format(.SearchMultiSelRetCol(i, 8), "#######0.00") + " %" & vbLf
                     SmsString = SmsString + "Late Days : " + Format(.SearchMultiSelRetCol(i, 10), "########0") & vbLf
                     SmsString = SmsString + "Grase     : " + Format(.SearchMultiSelRetCol(i, 9), "######0") & vbLf
                    
                     PtyMbl = "9665748022"
                     SentWhatsApp "Txt", "", SmsString, PtyMbl, ""
                    
                    
                    
                    '-- Update TblRecVsSale Interest Fields to Null & Zero
                    gCn.Execute "update tblRecVsSale set INTDBNTDATE=Null,INTDBNTRATE =0,INTDBNTGRASEDAY=0," _
                    & " INTDBNTLATEDAY=0,ISINTDBNTLATEFROMINVDT=0, INTDBNTAMT=0 ,IntDbNtRecdAmt=0,IsIntDbNtComplete=0  where AdjNo=" & RecAdjNo
                Next i
                MsgBox "Interest Debit Note Get Cancelled for Selected Receipt Successfully.", vbOKOnly, "Cancel Late Pay Interest Debit Note"
            End If
        End If
    End With
End Sub
'--- Procedure Prepare Interest Debit Note
Private Sub ProcPrepareIntDbNt(DbNtDt As Date, IsIntLateFrmInvDt As Long, IntGraseDays As Long, IntRate As Double, IntRateMonthly As Integer)
Dim X As String
Dim RecAdjNo As Long '-Rec Adj No
Dim IntDbNtAmt As Double '-Interest DbNt Amount
Dim IntLateDay As Long '-- Interest Late Days
Dim IntDivideBy As Long
If IntRateMonthly = 1 Then
    IntDivideBy = 3000
Else
    IntDivideBy = 36500
End If
'-- Bal Interest >0 and Recd. dt <= Db Nt Date
X = "select Vdt,Vno,Nar1,Nar2,VBillDt,VBillNo,amt1,Amt5,clng((AMT2*AMT6/" & IntDivideBy & ")*AMT5),Amt4,Amt2,amt3,Dt1" _
& " from tmpGentbl where UserName='" & gUserName & "' and vdt<=cdate('" & CDate(DbNtDt) & "') and (clng((AMT2*AMT6/" & IntDivideBy & ")*AMT5)-amt4>0)"
gClsSearch.SearchMultiSelField X, "", Array("Recd .Dt", "AdjNo", "Party", "Firm", "Bill Dt", "Bill No", "Bill Amt", "Late", "Interest Amt", "Int Recd", "Recd. Amt", "Bal Amt", "Due Dt"), Array(1400, 0, 2500, 0, 1400, 1000, 1400, 1100, 1100, 1400, 1100, 1100, 1100), OrderCol:="Nar1,Vdt,VbillDt,VbillNo", Cn:=gSetupCn, IstblNameAsQry:=True, IsReportFlag:=True, ReportCaption:="Late Pay Interest Debit Note On Date " & CDate(DtpDbNtDt.Value) & "  (Interest: " & txt(1) & " ,Grase Days: " & txt(2) & ")"
    With gClsSearch
        If .SearchMultiSelRetCol(0, 0) <> "" Then
            If MsgBox("Interest Debit Note Will Get Applied to Selected Receipt. Are You Sure ?", vbYesNo + vbInformation, "Late Pay Interest Debit Note") = vbYes Then
                For i = 0 To UBound(.SearchMultiSelRetCol)
                    RecAdjNo = Trim(.SearchMultiSelRetCol(i, 1)) '-Rec Adj No
                    IntLateDay = .SearchMultiSelRetCol(i, 7) '- Interest Late Days
                    IntDbNtAmt = .SearchMultiSelRetCol(i, 8) '-Interest Amount
                    '-- Update TblRecVsSale Interest Fields
                    If IntDbNtAmt > 0 Then
                        gCn.Execute "update tblRecVsSale set INTDBNTDATE=" & IIf(gBackEndDB = gBackEndAccess, "'" & CDate(DbNtDt) & "'", "to_date('" & DbNtDt & "','dd/MM/yyyy')") & ",INTDBNTRATE =" & IntRate & ",INTDBNTGRASEDAY=" & IntGraseDays & "," _
                        & " INTDBNTLATEDAY=" & IntLateDay & ",ISINTDBNTLATEFROMINVDT=" & IsIntLateFrmInvDt & " , INTDBNTAMT=" & IntDbNtAmt & " where AdjNo=" & RecAdjNo
                    End If
                Next i
                MsgBox "Interest Debit Note Get Applied to Selected Receipt Successfully.", vbOKOnly, "Late Pay Interest Debit Note"
            End If
        End If
    End With
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
    gRptOnDateFlag = False
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
    dtpToDt.Value = Format(gCYEDate, "dd/MM/yyyy")
    Me.Caption = gSelectedMenu
    If gRptBrokerFlag = True Then
        lblName.Caption = "Broker"
    End If
    If gRptPartyFlag = True Then
        lblName.Caption = "Party"
    End If
    For i = 1 To 4
        txt(i) = 0
    Next i
    '-- Activate Controls
    GProcActivateControls Me, False
    dtpFrDt.Enabled = True
    dtpToDt.Enabled = True
    txt(0).Enabled = True
    If gRptOnDateFlag = True Then dtpToDt.Enabled = False
    cmdDbNt(0).Visible = False '- Int Debit Note
    cmdDbNt(1).Visible = False
    Select Case gSelectedMenu
        Case "Stock Statement" '-"Stock Statement"
             dtpFrDt.Enabled = False
             txt(0).Enabled = False '-Mill
        Case "Late Payment Interest Register", "Late Payment Interest Register (Mill Bill)" ''===  Late Payment Interest Register
            txt(0).Enabled = True '-party
            ChkConsolidate.Enabled = True
            txt(1).Enabled = True '-Int %
            txt(2).Enabled = True '- Grase Days
            DtpDbNtDt.Enabled = True '--DB Nt Dt
            DtpDbNtDt.Value = Date
            ChkIntLateFrmInvDt.Enabled = True '--Interest Late From Inv Dt
            ChkIntRateMonthly.Enabled = True '--Interest Rate Monthly
            cmdDbNt(0).Visible = True '- Apply Int Debit Note
            cmdDbNt(1).Visible = True '- Cancel Int Debit Note
            If gIsAdminUser = False Then
               cmdDbNt(1).Visible = False
               cmdDbNt(1).Enabled = False
            End If
        Case "Broker, Party & Due Dayswise Aging Report" '=== "Broker, Party & Due Dayswise Aging Report"
            txt(3).Enabled = True
            txt(4).Enabled = True
        Case "Interest Demand Note" '== Interest Demand Note
            txt(0).Enabled = True '-party
            ChkIntDbNtShort.Enabled = True
       Case "Stock Statement Datewise"
           ChkConsolidate.Enabled = True
           ChkConsolidate.Caption = "Is Countwise Report"
        Case "Late Payment Interest Purchase List"
            txt(0).Enabled = True '-party
            ChkConsolidate.Visible = False
            ChkConsolidate.Enabled = False
            txt(1).Enabled = True '-Int %
            txt(2).Enabled = True '- Grase Days
            DtpDbNtDt.Enabled = True '--DB Nt Dt
            DtpDbNtDt.Value = Date
            ChkIntLateFrmInvDt.Enabled = True '--Interest Late From Inv Dt
            ChkIntRateMonthly.Enabled = True '--Interest Rate Monthly
            cmdDbNt(0).Visible = False '- Apply Int Debit Note
            cmdDbNt(1).Visible = False '- Cancel Int Debit Note
        Case "O/s In Date Range"
            txt(0).Enabled = False
        Case "Broker & Partywise Aging Report"
           ChkConsolidate.Enabled = True
           ChkConsolidate.Caption = "Is Bank Ageing Report ?"
        Case "Countwise Profit and Loss"
             txt(0).Enabled = False '-Mill
             txt(0).Visible = False
             dtpFrDt.Enabled = True
             ChkConsolidate.Visible = False
             ChkIntDbNtShort.Visible = False
             Label1(4).Visible = False
             txt(2).Visible = False
         '    Label3(1).Visible = False
             Label1(2).Visible = False
             txt(1).Visible = False
             ChkIntLateFrmInvDt.Visible = False
             Label5.Visible = False
             DtpDbNtDt.Visible = False
             ChkIntRateMonthly.Visible = False
             Label4.Visible = False
             Label1(3).Visible = False
             txt(3).Visible = False
             Label1(1).Visible = False
             txt(4).Visible = False
             lblName.Visible = False
End Select
'-- Set InActivate Ctrl Back Color
GProcSetInActivateCtrlBackClr Me
End Sub

Private Sub txt_GotFocus(Index As Integer)
GProcSelectBox Me.ActiveControl
End Sub

Private Sub txt_KeyPress(Index As Integer, KeyAscii As Integer)
Dim tEMPACCCODE As Integer
If KeyAscii = 13 Then Exit Sub '-- Enter
Select Case Index
    Case 0
        If gRptBrokerFlag = True Then '-- Broker
            gClsSearch.SearchMultiField "QryMastaccount", "Account,AcName,ACCode", Array("Broker", "AcName", ""), Array(txt(0).Width, 0, 0), " (AgCode=90016  or GpCode=90016)", Chr(KeyAscii), "AcName", txt(Index).Left + Me.Left, txt(Index).Top + Me.Top + 1000, False
        Else
            If gRptMillFlag = True Then '-mill
                gClsSearch.SearchMultiField "QryMastaccount", "Account,AcName,ACCode", Array("Account Name", "AcName", ""), Array(txt(0).Width, 0, 0), " (AgCode=90029 or GpCode=90029)", Chr(KeyAscii), "AcName", txt(Index).Left + Me.Left, txt(Index).Top + Me.Top + 1000, False
            Else '-Party
                gClsSearch.SearchMultiField "QryMastaccount", "Account,AcName,ACCode", Array("Account Name", "AcName", ""), Array(txt(0).Width, 0, 0), "", Chr(KeyAscii), "AcName", txt(Index).Left + Me.Left, txt(Index).Top + Me.Top + 1000, False
            End If
        End If
        If gClsSearch.SearchMultiRetCol(1) <> "" Then
            txt(Index) = gClsSearch.SearchMultiRetCol(1)
            tEMPACCCODE = gClsSearch.SearchMultiRetCol(2)
            KeyAscii = 0
        End If
     Case 1, 2, 3, 4
         KeyAscii = GProcNumberOnly(KeyAscii)
End Select
End Sub
'--- Interest Debit Note
Private Sub cmdIntDbNt_Click()

End Sub


