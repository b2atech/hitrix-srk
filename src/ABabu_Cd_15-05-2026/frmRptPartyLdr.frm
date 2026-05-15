VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form frmRptPartyLdr 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "."
   ClientHeight    =   4365
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9435
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4365
   ScaleWidth      =   9435
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Txt 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   2
      Left            =   1605
      TabIndex        =   9
      Top             =   1635
      Width           =   4335
   End
   Begin VB.CheckBox ChkIsCashBankGroup 
      Caption         =   "&Cash and Bank Book"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   6345
      TabIndex        =   21
      Top             =   2625
      Value           =   1  'Checked
      Width           =   3135
   End
   Begin VB.CheckBox ChkIsRefNo 
      Caption         =   "&Ref. No Required"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   6345
      TabIndex        =   20
      Top             =   2220
      Value           =   1  'Checked
      Width           =   3135
   End
   Begin VB.CommandButton cmdSPE 
      Height          =   495
      Index           =   0
      Left            =   1065
      Picture         =   "frmRptPartyLdr.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   22
      Top             =   3600
      Width           =   1335
   End
   Begin VB.CommandButton cmdSPE 
      Height          =   495
      Index           =   1
      Left            =   2640
      Picture         =   "frmRptPartyLdr.frx":0442
      Style           =   1  'Graphical
      TabIndex        =   23
      ToolTipText     =   "Print"
      Top             =   3600
      Width           =   1335
   End
   Begin VB.TextBox Txt 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   0
      Left            =   1605
      TabIndex        =   5
      Top             =   675
      Width           =   4335
   End
   Begin VB.Frame Frame1 
      Caption         =   "Report Format"
      ForeColor       =   &H000000C0&
      Height          =   855
      Left            =   270
      TabIndex        =   26
      Top             =   2205
      Width           =   5580
      Begin VB.OptionButton optReportOn 
         Caption         =   "&Grid"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   2
         Left            =   3885
         TabIndex        =   12
         Top             =   360
         Width           =   1425
      End
      Begin VB.OptionButton optReportOn 
         Caption         =   "&Windows"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   1
         Left            =   2085
         TabIndex        =   11
         Top             =   360
         Value           =   -1  'True
         Width           =   1425
      End
      Begin VB.OptionButton optReportOn 
         Caption         =   "MS &DOS"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   0
         Left            =   210
         TabIndex        =   10
         Top             =   360
         Width           =   1425
      End
   End
   Begin VB.OptionButton optSelType 
      Caption         =   "&T Format"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Index           =   3
      Left            =   7755
      TabIndex        =   17
      Top             =   930
      Width           =   1425
   End
   Begin VB.OptionButton optSelType 
      Caption         =   "&Monthly"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Index           =   2
      Left            =   6315
      TabIndex        =   16
      Top             =   915
      Width           =   1425
   End
   Begin VB.OptionButton optSelType 
      Caption         =   "&Partywise"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Index           =   1
      Left            =   7755
      TabIndex        =   15
      Top             =   555
      Width           =   1425
   End
   Begin VB.OptionButton optSelType 
      Caption         =   "&Simple"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Index           =   0
      Left            =   6315
      TabIndex        =   14
      Top             =   540
      Value           =   -1  'True
      Width           =   1425
   End
   Begin VB.CheckBox chkDaySum 
      Caption         =   "Daywise &Summary"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   6345
      TabIndex        =   18
      Top             =   1380
      Width           =   3135
   End
   Begin Crystal.CrystalReport CrystalReport1 
      Left            =   5520
      Top             =   4050
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
   Begin VB.CommandButton cmdSPE 
      Caption         =   "E&xit"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   2
      Left            =   4215
      TabIndex        =   24
      Top             =   3600
      Width           =   1335
   End
   Begin MSComCtl2.DTPicker DTPicker1 
      Height          =   375
      Left            =   1605
      TabIndex        =   1
      Top             =   195
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   661
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
      Format          =   52363267
      CurrentDate     =   38050
   End
   Begin MSComCtl2.DTPicker DTPicker2 
      Height          =   375
      Left            =   4380
      TabIndex        =   3
      Top             =   195
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   661
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
      Format          =   52363267
      CurrentDate     =   38050
   End
   Begin VB.TextBox Txt 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   1
      Left            =   1605
      TabIndex        =   7
      Top             =   1140
      Width           =   4335
   End
   Begin VB.ListBox lstParty 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Left            =   1605
      Sorted          =   -1  'True
      TabIndex        =   25
      Top             =   1140
      Width           =   4335
   End
   Begin MSComctlLib.ProgressBar prgBar 
      Height          =   255
      Left            =   240
      TabIndex        =   27
      Top             =   3150
      Visible         =   0   'False
      Width           =   5595
      _ExtentX        =   9869
      _ExtentY        =   450
      _Version        =   393216
      Appearance      =   1
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "&Mill             :"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   210
      TabIndex        =   8
      Top             =   1635
      Width           =   1275
   End
   Begin VB.Line Line2 
      BorderColor     =   &H000040C0&
      X1              =   8010
      X2              =   9390
      Y1              =   1950
      Y2              =   1950
   End
   Begin VB.Label Label6 
      Caption         =   "A/c G&roupwise Ledger"
      ForeColor       =   &H000040C0&
      Height          =   330
      Left            =   6255
      TabIndex        =   19
      Top             =   1845
      Width           =   1770
   End
   Begin VB.Line Line3 
      BorderColor     =   &H000040C0&
      X1              =   6120
      X2              =   6120
      Y1              =   195
      Y2              =   3390
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "A/C &Group   :"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   210
      TabIndex        =   4
      Top             =   675
      Width           =   1335
   End
   Begin VB.Label Label4 
      Caption         =   "Accoutwise &Ledger"
      ForeColor       =   &H000040C0&
      Height          =   330
      Left            =   6270
      TabIndex        =   13
      Top             =   180
      Width           =   1545
   End
   Begin VB.Line Line1 
      BorderColor     =   &H000040C0&
      X1              =   7875
      X2              =   9360
      Y1              =   285
      Y2              =   285
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "To Date       :"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3285
      TabIndex        =   2
      Top             =   195
      Width           =   990
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "&Account       :"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   210
      TabIndex        =   6
      Top             =   1140
      Width           =   1305
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "From &Date   :"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   210
      TabIndex        =   0
      Top             =   195
      Width           =   1335
   End
End
Attribute VB_Name = "frmRptPartyLdr"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim i As Integer, LnCnt As Integer, PgCnt As Integer, wDotLine As String
Dim wAcCode As Long, mScreen As Boolean, mP12 As String, mPc As String
Dim wPartyCd(20) As String
Private Sub cmdSPE_Click(Index As Integer)
Dim sTp As String
Dim FromDt As Date
Dim ToDt As Date
Dim X As String
Select Case Index
Case 0
    mScreen = True
Case 1
    mScreen = False
Case 2
   Unload Me
   Exit Sub
End Select
FromDt = Format(DTPicker1.Value, "dd/MM/yyyy") '-From Dt
ToDt = Format(DTPicker2.Value, "dd/MM/yyyy") '-To Dt
'--- Ledger Type
If optSelType(0) = True Then '-Simple
     sTp = "D"
ElseIf optSelType(1) = True Then '-Partywise
     sTp = "P"
ElseIf optSelType(2) = True Then '-Monthly
     sTp = "M"
ElseIf optSelType(3) = True Then '-T Format
      sTp = "T"
End If
'== Dos
If optReportOn(0) = True Then
   Pty_Ldr1 (sTp)
'== Window
ElseIf optReportOn(1) = True Then
    If mScreen Then
        GProcCrystalRptPreparation CrystalReport1, "Window"
    Else
        GProcCrystalRptPreparation CrystalReport1, "Printer"
    End If
    With CrystalReport1
        Select Case gSelectedMenu
            Case "Account Ledger"  '====="Account Ledger"
                If sTp = "T" Then '-- T Format Ledger
                Else
                     '-insert Date Range,UserName
                     gCn.Execute "delete from tmpGenTbl where UserName='" & gUserName & "'"
                     If gBackEndDB = gBackEndAccess Then
                         gCn.Execute "insert into tmpGenTbl (dt1,Dt2,UserName) values ('" & FromDt & "','" & ToDt & "','" & gUserName & "')"
                     Else
                         gCn.Execute "insert into tmpGenTbl (dt1,dt2,UserName) values (to_date('" & FromDt & "','dd/MM/yyyy'),to_Date('" & ToDt & "','dd/MM/yyyy'),'" & gUserName & "') "
                     End If
                     '--Set Loacation
                     If gBackEndDB = gBackEndAccess Then
                         .DataFiles(0) = gDbLocation
                         gReportPath = App.Path & "\HiReports_Access\"
                     Else '-- Oracle
                         .Connect = "DSN=" + gODBCDSN + " ;UID=Scott;PWD=" + gODBCPwd + ";DBQ=" + gODBCHostStr
                         gReportPath = App.Path & "\HiReports_ORA\"
                     End If
                     .Formulas(4) = "wFrToDt = 'From : ' & '" & FromDt & "' & '  To ' & '" & ToDt & "'"
                     '-- Selection Formula AcName
                     X = "{QryAccountLedger.UserName}='" & gUserName & "'"
                     If Txt(0) <> "" Then X = X + " and {QryAccountLedger.AGNAME}='" & Txt(0) & "'"
                     If Txt(1) <> "" Then X = X + " and {QryAccountLedger.ACNAME}='" & Txt(1) & "'"
                     If gSelectedMenu = "Bank Book" Then X = X + " and {QryAccountLedger.ACNAME}<>'Cash In Hand'"
                     .SelectionFormula = X
                    '-- Report Name
                    Select Case sTp
                      Case "D"  '-Detail
                        If chkDaySum = 1 Then
                            .ReportFileName = gReportPath & "rptAccLedgerVdtwise.rpt" '-Dt wise
                        Else
                            .Formulas(5) = "wRptName = 'ACCOUNT LEDGER'"
                            .ReportFileName = gReportPath & "rptAccountLedger.rpt"
                        End If
                      Case "P"  '-Party
                         .ReportFileName = gReportPath & "rptAccLedgerpartywise.rpt"
                      Case "M"  '-Monthly
                        .ReportFileName = gReportPath & "rptAccLedgermonthly.rpt"
                    End Select
              End If
             Case "Cash Book", "Bank Book", "Credit Cash Report"  '---"Cash Book", "Bank Book","Credit Cash Report"
                 '-insert Date Range,UserName
                 gCn.Execute "delete from tmpGenTbl where UserName='" & gUserName & "'"
                 If gBackEndDB = gBackEndAccess Then
                     gCn.Execute "insert into tmpGenTbl (dt1,Dt2,UserName) values ('" & FromDt & "','" & ToDt & "','" & gUserName & "')"
                 Else
                     gCn.Execute "insert into tmpGenTbl (dt1,dt2,UserName) values (to_date('" & FromDt & "','dd/MM/yyyy'),to_Date('" & ToDt & "','dd/MM/yyyy'),'" & gUserName & "') "
                 End If
                 '--Set Loacation
                 If gBackEndDB = gBackEndAccess Then
                     .DataFiles(0) = gDbLocation
                     gReportPath = App.Path & "\HiReports_Access\"
                 Else '-- Oracle
                     .Connect = "DSN=" + gODBCDSN + " ;UID=Scott;PWD=" + gODBCPwd + ";DBQ=" + gODBCHostStr
                     gReportPath = App.Path & "\HiReports_ORA\"
                 End If
                 .Formulas(4) = "wFrToDt = 'From : ' & '" & FromDt & "' & '  To ' & '" & ToDt & "'"
                 '-- Selection Formula AcName
                 X = "{QryAccountLedger.UserName}='" & gUserName & "'"
                 If Txt(0) <> "" Then X = X + " and {QryAccountLedger.AGNAME}='" & Txt(0) & "'"
                 If Txt(1) <> "" Then X = X + " and {QryAccountLedger.ACNAME}='" & Txt(1) & "'"
                 If gSelectedMenu = "Bank Book" Then X = X + " and {QryAccountLedger.ACNAME}<>'Cash In Hand'"
                .SelectionFormula = X
                '-- Report Name
                If gSelectedMenu = "Credit Cash Report" Then '-- Credit Cash Report
                    .Formulas(5) = "wRptName = 'Credit Cash Report'"
                    .Formulas(6) = "wIsCrCashRpt = 'Y'"
                Else
                    If gSelectedMenu = "Cash Book" Then '-- Cash Book
                        .Formulas(5) = "wRptName = 'Cash Book'"
                    Else '-- Bank Book
                        .Formulas(5) = "wRptName = 'Bank Book'"
                    End If
                 End If
                 .ReportFileName = gReportPath & "rptAccountLedger.rpt"
            Case "Groupwise Account Ledger" '===="Groupwise Account Ledger"
                '-insert Date Range,UserName
                     gCn.Execute "delete from tmpGenTbl where UserName='" & gUserName & "'"
                     If gBackEndDB = gBackEndAccess Then
                         gCn.Execute "insert into tmpGenTbl (dt1,Dt2,UserName) values ('" & FromDt & "','" & ToDt & "','" & gUserName & "')"
                     Else
                         gCn.Execute "insert into tmpGenTbl (dt1,dt2,UserName) values (to_date('" & FromDt & "','dd/MM/yyyy'),to_Date('" & ToDt & "','dd/MM/yyyy'),'" & gUserName & "') "
                     End If
                     '--Set Loacation
                     If gBackEndDB = gBackEndAccess Then
                         .DataFiles(0) = gDbLocation
                         gReportPath = App.Path & "\HiReports_Access\"
                     Else '-- Oracle
                         .Connect = "DSN=" + gODBCDSN + " ;UID=Scott;PWD=" + gODBCPwd + ";DBQ=" + gODBCHostStr
                         gReportPath = App.Path & "\HiReports_ORA\"
                     End If
                     .Formulas(4) = "wFrToDt = 'From : ' & '" & FromDt & "' & '  To ' & '" & ToDt & "'"
                     '-- Selection Formula AcName ,AgName,Bank Book ,cash Book
                     X = "{qryAccGroupWiseLedger.UserName}='" & gUserName & "'"
                     If Txt(1) <> "" Then X = X + " and {qryAccGroupWiseLedger.ACNAME}='" & Txt(1) & "'" 'A/c
                     If Txt(0) <> "" Then X = X + " and {qryAccGroupWiseLedger.AGNAME}='" & Txt(0) & "'" 'Group
                     If ChkIsCashBankGroup = 0 Then X = X + " and {qryAccGroupWiseLedger.AGCode}<>90019"  'No Cash Bank Group
                     .SelectionFormula = X
                     If ChkIsRefNo = 1 Then '-Ref No in Ledger
                        .ReportFileName = gReportPath & "rptAccGroupLedger.rpt"
                     Else
                        .ReportFileName = gReportPath & "rptAccGroupLedgerNoRefNo.rpt"
                     End If
            Case "Account Confirmation" '-"Account Confirmation"
                  '-insert Date Range,UserName
                     gCn.Execute "delete from tmpGenTbl where UserName='" & gUserName & "'"
                     If gBackEndDB = gBackEndAccess Then
                         gCn.Execute "insert into tmpGenTbl (dt1,Dt2,UserName) values ('" & FromDt & "','" & ToDt & "','" & gUserName & "')"
                     Else
                         gCn.Execute "insert into tmpGenTbl (dt1,dt2,UserName) values (to_date('" & FromDt & "','dd/MM/yyyy'),to_Date('" & ToDt & "','dd/MM/yyyy'),'" & gUserName & "') "
                     End If
                     '--Set Loacation
                     If gBackEndDB = gBackEndAccess Then
                         .DataFiles(0) = gDbLocation
                         gReportPath = App.Path & "\HiReports_Access\"
                     Else '-- Oracle
                         .Connect = "DSN=" + gODBCDSN + " ;UID=Scott;PWD=" + gODBCPwd + ";DBQ=" + gODBCHostStr
                         gReportPath = App.Path & "\HiReports_ORA\"
                     End If
                     .Formulas(4) = "wFrToDt = 'From ' & '" & FromDt & "' & '  To ' & '" & ToDt & "'"
                     .Formulas(5) = "PANNo= '" & gCPAN & "'"
                     '-- Selection Formula AcName
                     X = "{QryAccountLedger.UserName}='" & gUserName & "'"
                     If Txt(1) <> "" Then X = X + " and {QryAccountLedger.ACNAME}='" & Txt(1) & "'" 'A/c
                     If Txt(0) <> "" Then X = X + " and {QryAccountLedger.AGNAME}='" & Txt(0) & "'" 'Group
                     .SelectionFormula = X
                     '-- Report Name
                    .ReportFileName = gReportPath & "rptAccConfirmation.rpt"
            Case "R25":
                .ReportFileName = App.Path & "\Reports\rptAcCon.rpt"
            Case "R2M":
                .Formulas(7) = "wRepoHead='MULTIPLE ACCOUNT LEDGER'"
                .ReportFileName = App.Path & "\Reports\rptSingleAcLed.rpt"
        End Select
        .Action = 1
    End With
    Screen.MousePointer = vbDefault
ElseIf optReportOn(2) = True Then '-Grid
    Pty_Ldr1 ("F")
End If
End Sub
Private Sub DTPicker2_LostFocus()
If (gSelectedMenu = "R2M") Then
    lstParty_GotFocus
    Exit Sub
End If
End Sub
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
Select Case KeyCode
Case 13
    SendKeys "{TAB}"
Case 27
    Unload Me
End Select
End Sub
Private Sub Form_Load()
Dim i As Long
Me.Top = 700
Me.Left = Screen.Width / 2 - Me.Width / 2
 '-- Activate Controls
GProcActivateControls Me, False
DTPicker1.Enabled = True
DTPicker2.Enabled = True
lstParty.Visible = False
For i = 0 To 2 '-- Report On
    optReportOn(i).Enabled = True
Next i
'-- Set Values
DTPicker1.Value = Format(gCYSDate, "dd/MM/yyyy")
DTPicker2.Value = Format(gCYEDate, "dd/MM/yyyy")
mScreen = True
Me.Caption = gSelectedMenu
Select Case gSelectedMenu
    Case "Account Ledger" '-"Account Ledger"
        Txt(1).Enabled = True '-A/c
       ' lstParty.Enabled = True
        chkDaySum.Enabled = True '-Summ
        For i = 0 To 3 '-- Report format
            optSelType(i).Enabled = True
        Next i
    Case "Groupwise Account Ledger" '-"Groupwise Account Ledger"
        Txt(0).Enabled = True '-A/c Group
        Txt(1).Enabled = True '-A/c
        ChkIsCashBankGroup.Enabled = True '-CAsh ,Bank Group
        ChkIsRefNo.Enabled = True '-Ref No
    Case "Account Confirmation" '-"Account Confirmation"
        Txt(1).Enabled = True '-A/c
        Txt(0).Enabled = True '-A/c Group
    Case "Cash Book" '"Cash Book"
        Txt(1) = "Cash In Hand" '-Cash A/c
    Case "Bank Book" 'Bank Book"
        Txt(1).Enabled = True '-A/c
        Txt(0) = GProcGetColumnValue("tblMastGroup", "AgCode", 90019, "N", "AgName", "S") '-Bank Group
    Case "Credit Cash Report" '"Credit Cash Report"
        Txt(1) = "Cash In Hand" '-Cash A/c
    'Case "R22"
    '    Label3.Caption = "Group Name :"
    '    Me.Caption = "Single Group Ledger"
    '    txt(0).Enabled = True '-Group
    'Case "R23"
    '    Me.Caption = "All Account Ledger"
    '    txt(1).Visible = False
    '    Label3.Visible = False
    Case "R25"
        Me.Caption = "Account Confirmation"
    'Case "R26"
    '    Me.Caption = "Account Ledger (Monthly)"
    'Case "R27"
    '    Me.Caption = "Account Ledger (Partywise)"
    Case "R2M"
        Me.Caption = "Multi Account Ledger"
        lstParty.Enabled = True
        Txt(1).Visible = False
End Select
End Sub

Private Sub Label3_Click()

End Sub

Private Sub lstParty_GotFocus()
'lstParty.Clear
'Dim i As Integer
'Call iisGPrActivateControls(Me, False)
'TopPos = Me.Top
''+ lstParty.Top
'LeftPos = Me.Left
''+ lstParty.Left
'    '----Account Name
'    mSelHead = "Select Multiple Account Name "
'    mySele = lstParty.Text
'    mySql = "select ANAME from ACCT where ACTYPE='A'"
'    myField = "ANAME"
'    myCnd = " and "
'    DoEvents
'    frmSearchMultiSel.Show 1
'    Do While mRet(i) <> ""
'        lstParty.AddItem (mRet(i))
'        DoEvents
'        wPartyCd(i) = iisGFnGetColumnValue(gCn, "select ACCODE from ACCT where ANAME = '" & mRet(i) & "'")
'        i = i + 1
'    Loop
'Call iisGPrActivateControls(Me, True)
'Command1.SetFocus
End Sub

Private Sub txt_GotFocus(Index As Integer)
'If index = 1 And lstParty.Visible = True Then
 '   lstParty.SetFocus
'Else
    GProcSelectBox Me.ActiveControl
'End If
End Sub
Private Sub Pty_Ldr1(wRpTp As String)
Dim tdt As Date, currdt As Date, CrSide(2000) As String, DrSide(2000) As String, CrCnt As Integer, DrCnt As Integer
Dim t1 As Double, t2 As Double, t3 As Double, wnp As Double, T11 As Double
Dim wTp As String, X As String, X1 As String, ttp As String, wAcCode1 As Long, wwaname As String, wwagcode As Long
Dim rst1 As Recordset, r_s As String, wTTrno As String, wTContP As Boolean
Dim rst2 As Recordset, rstTran33 As Recordset, j As Integer
Dim wHdPrn As Boolean, wNamePrn As Boolean
Dim wRecLen As Long, z1 As String, rstIt As Recordset, rstItP As Recordset
Dim wDPay As Double, wDRec As Double, mbkMk As Variant
wDotLine = String(96, "-")
PgCnt = 1
CrCnt = 0
DrCnt = 0
If (gSelectedMenu <> "R23") And (gSelectedMenu <> "R2M") Then
    If (Len(wAcCode) = 0) Then
        MsgBox "Select Party Name", vbCritical + vbOKOnly, "Select Party Name"
        txt_GotFocus (1)
        Exit Sub
    End If
End If

If mScreen = True Then
    mP12 = ""
    mPc = ""
Else
    mP12 = Chr(27) + Chr(33) + Chr(1)
    mPc = Chr(27) + Chr(15)
End If

Select Case gSelectedMenu
Case "Account Ledger", "R25", "R26", "R27"
    r_s = "and vACCODE = " & wAcCode
Case "R22"
    r_s = "and AGCODE = " & wAcCode
Case "R23"
    r_s = ""
   ' gCn.Execute "Update ACCT set PGPR =0 where ACCODE ='" & wAcCode & "' "
Case "R2M":
    z1 = ""
    For i = 0 To (lstParty.ListCount - 1)
        If i < (lstParty.ListCount - 1) Then
            z1 = z1 + " (vACCODE = " & wPartyCd(i) & ") or "
        ElseIf i = (lstParty.ListCount - 1) Then
            z1 = z1 + " (vACCODE = " & wPartyCd(i) & ") "
        End If
    Next i
    r_s = "and (" & z1 & ")"
End Select
DoEvents


'======================================= PRINTING START ========================
wHdPrn = True
If gBackEndDB = gBackEndORA Then
T11 = GProcRstOpen(rstTran33, "Select * from qryPartyLedgerSingle where vdt <= to_date('" & DTPicker2 & "','dd/MM/yyyy') " & r_s & " order by vaccode,vdt,vtype,vno,Vctrno", "R", gCn)
Else
T11 = GProcRstOpen(rstTran33, "Select * from qryPartyLedgerSingle where format(vdt,'yyyy/MM/dd') <= '" & Format(DTPicker2, "yyyy/MM/dd") & "' " & r_s & " order by vaccode,vdt,vtype,vno,Vctrno", "R", gCn)
End If
If (T11 = 0) Then
   rstTran33.Close
   MsgBox "No Report available", vbCritical + vbOKOnly, "No Report available"
   Exit Sub
End If
Dim mOpBal As Double, TmpVno As Long, TmpVType As String, TmpVamt As Double
mOpBal = 0
prgBar.Value = 0
prgBar.Max = rstTran33.RecordCount
prgBar.Visible = True
If (wRpTp = "D") Or (wRpTp = "T") Or (wRpTp = "M") Or (wRpTp = "P") Then
    Dim wChgDate As Date
    wDPay = 0
    wDRec = 0
    wChgDate = DateAdd("d", -1, gCYSDate)
    Open "c:\rpt1.txt" For Output As #1
       With rstTran33
            .MoveFirst
            Do While Not .EOF
               wAcCode = !vacCode
                If (wRpTp = "T") Or (wRpTp = "M") Or (wRpTp = "P") Then
                    For i = 1 To 2000
                        CrSide(i) = ""
                        DrSide(i) = ""
                    Next i
                End If
               t1 = 0
               t2 = 0
               wNamePrn = True
Loop2:
               Do While Not .EOF
                  If (wAcCode <> !vacCode) Then Exit Do
                  If (!Vdt < DTPicker1) Then
                     mOpBal = mOpBal + !Credit - !Debit
                     .MoveNext
                     GoTo Loop2
                  End If
                  Select Case wRpTp
                    Case "M":
                        If (!Debit > 0) Then
                           DrSide(Month(!Vdt)) = GProcNoToStr(!Debit + Val(DrSide(Month(!Vdt))), 12, 2)
                        Else
                           t2 = t2 + T11
                           CrSide(Month(!Vdt)) = GProcNoToStr(!Credit + Val(CrSide(Month(!Vdt))), 12, 2)
                        End If
                    Case "P":
                        For j = 1 To DrCnt
                            If Mid(DrSide(j), 25) = !AcName Then Exit For
                        Next j
                        If j > DrCnt Then
                            DrCnt = DrCnt + 1
                            If (!Debit > 0) Then
                               DrSide(j) = GProcNoToStr(!Debit, 12, 2) + Space(12) + !AcName
                            Else
                               DrSide(j) = Space(12) + GProcNoToStr(!Credit, 12, 2) + !AcName
                            End If
                        Else
                            If (!Debit > 0) Then
                               DrSide(j) = GProcNoToStr(!Debit + Val(Left(DrSide(j), 12)), 12, 2) + Mid(DrSide(j), 13)
                            Else
                               DrSide(j) = Left(DrSide(j), 12) + GProcNoToStr(!Credit + Val(Mid(DrSide(j), 13, 12)), 12, 2) + Mid(DrSide(j), 25)
                            End If
                        End If
                    Case "T":
                        Select Case i
                        Case (i < 3):
                           X1 = Left(!AcName + Space(19), 19) + " No." + Left(!VBILLNO, 7)
                        Case (i < 5):
                           X1 = Left(!AcName + Space(19), 19) + " No." + Left(!VBILLNO, 7)
                        Case Else
                           X1 = Left(!AcName + Space(30), 30)
                        End Select
                        If (!Debit > 0) Then
                           DrCnt = DrCnt + 1
                           DrSide(DrCnt) = GProcNoToStr(!Debit, 12, 2) + "  " + Format(!Vdt, "dd/MM/yyyy") + "  " + Left(!Vtype, 2) + "  " + X1
                        Else
                           CrCnt = CrCnt + 1
                           CrSide(CrCnt) = GProcNoToStr(!Credit, 12, 2) + "  " + Format(!Vdt, "dd/MM/yyyy") + "  " + Left(!Vtype, 2) + "  " + X1
                        End If
                           If (Len(Trim(!VNar1)) > 0) Then
                              If (!Debit > 0) Then
                                 DrCnt = DrCnt + 1
                                 DrSide(DrCnt) = Space(30) + Left(!VNar1, 30)
                              Else
                                 CrCnt = CrCnt + 1
                                 CrSide(CrCnt) = Space(30) + Left(!VNar1, 30)
                              End If
                           End If
                           If (Len(Trim(!VNar2)) > 0) Then
                              If (!Debit > 0) Then
                                 DrCnt = DrCnt + 1
                                 DrSide(DrCnt) = Space(30) + Left(!VNar2, 30)
                              Else
                                 CrCnt = CrCnt + 1
                                 CrSide(CrCnt) = Space(30) + Left(!VNar2, 30)
                              End If
                           End If
                           If (Len(Trim(!VNar3)) > 0) Then
                              If (!Debit > 0) Then
                                 DrCnt = DrCnt + 1
                                 DrSide(DrCnt) = Space(30) + Left(!VNar3, 30)
                              Else
                                 CrCnt = CrCnt + 1
                                 CrSide(CrCnt) = Space(30) + Left(!VNar3, 30)
                              End If
                           End If
                           If (Len(Trim(!VNAR4)) > 0) Then
                              If (!Debit > 0) Then
                                 DrCnt = DrCnt + 1
                                 DrSide(DrCnt) = Space(30) + Left(!VNAR4, 30)
                              Else
                                 CrCnt = CrCnt + 1
                                 CrSide(CrCnt) = Space(30) + Left(!VNAR4, 30)
                              End If
                           End If
                  Case "D":
                        If wHdPrn Then
                           wHdPrn = False
                           DosHdg
                        End If
                        DoEvents
                        If wNamePrn Then
                           wNamePrn = False
                           wwaname = GProcGetColumnValue("tblmastaccount", "accode", CStr(wAcCode), "N", "acname", "S")
                           Print #1, Space(40 - (Len(wwaname) / 2)) + wwaname
                           Print #1, Space(40 - (Len(wwaname) / 2)) + String(Len(wwaname), "-") + mP12
                          ' If gSelectedmenu = "R23" Then Cn.Execute "Update ACCT set PGPR =" & (PgCnt - 1) & " where ACCODE ='" & wAcCode & "' and PGPR = 0 "
                           LnCnt = LnCnt + 2
                        End If
                        If (!Vdt <> wChgDate) Then
                           If ((wDPay > 0) Or (wDRec > 0)) And (chkDaySum.Value = 1) Then
                                If (LnCnt > 65) Then
                                    If mScreen = True Then
                                        Print #1, wDotLine
                                    Else
                                        Print #1, wDotLine + Chr(27) + Chr(33) + Chr(0)
                                    End If
                                    wHdPrn = True
                                    LnCnt = 0
                                End If
                                Print #1, Space(58) + String(21, "-")
                                If mScreen = True Then
                                    Print #1, Space(10) + "Day Totals : " + Space(35) + GProcNoToStr(wDRec, 11, 2) + GProcNoToStr(wDPay, 10, 2)
                                Else
                                    Print #1, Chr(27) + Chr(69) & Space(10) + "Day Totals : " + Space(35) + GProcNoToStr(wDRec, 11, 2) + GProcNoToStr(wDPay, 10, 2) + Chr(27) + Chr(70)
                                End If
                                Print #1, ""
                                LnCnt = LnCnt + 3
                                wDPay = 0
                                wDRec = 0
                            End If
                           X = Format(!Vdt, "dd/MM/yy") + "  "
                        Else
                           X = Space(10)
                        End If
                        wChgDate = !Vdt
                        If (mOpBal <> 0) Then
                            X1 = Space(10) + "             " + Left("Opening Balance" + Space(32), 32) + "  "
                            If (mOpBal < 0) Then
                               t1 = t1 + Abs(mOpBal)
                               wDRec = wDRec + Abs(mOpBal)
                               X1 = X1 + GProcNoToStr(Abs(mOpBal), 11, 2) + Space(10)
                            Else
                               t2 = t2 + Abs(mOpBal)
                               wDPay = wDPay + Abs(mOpBal)
                               X1 = X1 + Space(10) + GProcNoToStr(Abs(mOpBal), 11, 2)
                            End If
                            X1 = X1 + "  " + GProcNoToStr(Abs(Round((t1 - t2), 2)), 12, 2)
                            If (t1 > t2) Then
                               X1 = X1 + " Dr"
                            Else
                               X1 = X1 + " Cr"
                            End If
                            Print #1, mP12 + X1
                            LnCnt = LnCnt + 1
                            mOpBal = 0
                        End If
                        If mScreen = True Then
                           X = X + Left(Trim(!Vtype) + Space(2), 2) + "  " + Left(Trim(!Vno & vbNullString) + Space(7), 7) + "  " + Left(!AcName + Space(32), 32) + "  "
                        Else
                           X = X + Left(Trim(!Vtype) + Space(2), 2) + "  " + Left(Trim(!VBILLNO & vbNullString) + Space(7), 7) + "  " + Left(!AcName + Space(32), 32) + "  "
                        End If
                        TmpVType = !Vtype
                        TmpVno = !Vno
                        TmpVamt = 0
                        Do While Not .EOF
                            If (TmpVType <> !Vtype) Or (TmpVno <> !Vno) Then Exit Do
                            If (!Debit > 0) Then
                              TmpVamt = TmpVamt - !Debit
                            Else
                              TmpVamt = TmpVamt + !Credit
                            End If
                            .MoveNext
                        Loop
                        .MovePrevious
                            If (TmpVamt < 0) Then
                               t1 = t1 + Abs(TmpVamt)
                               wDRec = wDRec + Abs(TmpVamt)
                               X = X + GProcNoToStr(Abs(TmpVamt), 11, 2) + Space(10)
                            Else
                               t2 = t2 + TmpVamt
                               wDPay = wDPay + TmpVamt
                               X = X + Space(10) + GProcNoToStr(TmpVamt, 11, 2)
                            End If
                        X1 = GProcNoToStr(Abs(Round((t1 - t2), 2)), 12, 2)
                        If (t1 > t2) Then
                           X1 = X + "  " + X1 + " Dr"
                        Else
                           X1 = X + "  " + X1 + " Cr"
                        End If
                        Print #1, mP12 + X1
                        LnCnt = LnCnt + 1
                        If (Len(Trim(!VNar1)) > 0) Then
                           Print #1, Space(25) & mPc & Trim(!VNar1) & mP12
                           LnCnt = LnCnt + 1
                        End If
                        If (Len(Trim(!VNar2)) > 0) Then
                           Print #1, Space(25) & mPc & Trim(!VNar2) & mP12
                           LnCnt = LnCnt + 1
                        End If
                        If (Len(Trim(!VNar3)) > 0) Then
                           Print #1, Space(25) & mPc & Trim(!VNar3) & mP12
                           LnCnt = LnCnt + 1
                        End If
                        If (Len(Trim(!VNAR4)) > 0) Then
                           Print #1, Space(25) & mPc & Trim(!VNAR4) & mP12
                           LnCnt = LnCnt + 1
                        End If
                        If (LnCnt > 60) Then
                            If mScreen = True Then
                                Print #1, wDotLine
                            Else
                                Print #1, wDotLine + Chr(27) + Chr(33) + Chr(0) + Chr(12)
                            End If
                            wHdPrn = True
                            LnCnt = 0
                        End If
                   End Select
                  .MoveNext
               Loop
                If (wRpTp = "D") And ((wDPay > 0) Or (wDRec > 0)) And (chkDaySum.Value = 1) Then
                     Print #1, Space(58) + String(21, "-")
                     If mScreen = True Then
                         Print #1, Space(10) + "Day Totals : " + Space(35) + GProcNoToStr(wDRec, 11, 2) + GProcNoToStr(wDPay, 10, 2)
                     Else
                         Print #1, Chr(27) + Chr(69) & Space(10) + "Day Totals : " + Space(35) + GProcNoToStr(wDRec, 11, 2) + GProcNoToStr(wDPay, 10, 2) + Chr(27) + Chr(70)
                     End If
                     Print #1, ""
                     LnCnt = LnCnt + 3
                        If (LnCnt > 60) Then
                            If mScreen = True Then
                                Print #1, wDotLine
                            Else
                                Print #1, wDotLine + Chr(27) + Chr(33) + Chr(0) + Chr(12)
                            End If
                            wHdPrn = True
                            LnCnt = 0
                        End If
                 End If
                Select Case wRpTp
                Case "D":
                    If ((t1 > 0) Or (t2 > 0)) Then
                        If wHdPrn Then
                            wHdPrn = False
                            DosHdg
                        End If
                        If wNamePrn Then
                            wNamePrn = False
                           wwaname = GProcGetColumnValue("tblmastaccount", "accode", CStr(wAcCode), "N", "acname", "S", , gCn)
                            Print #1, Space(40 - (Len(wwaname) / 2)) + wwaname
                            Print #1, Space(40 - (Len(wwaname) / 2)) + String(Len(wwaname), "-") + mP12
                            LnCnt = LnCnt + 2
                        End If
                        X1 = GProcNoToStr(Abs(Round((t1 - t2), 2)), 12, 2)
                        If (t1 > t2) Then
                           X1 = X1 + " Dr"
                        Else
                           X1 = X1 + " Cr"
                        End If
                        Print #1, Space(54) + String(42, "-")
                        Print #1, Space(54) + GProcNoToStr(t1, 12, 2) + " " + GProcNoToStr(t2, 12, 2) + " " + X1
                        Print #1, Space(54) + String(42, "-") + Chr(27) + Chr(33) + Chr(0)
                        LnCnt = LnCnt + 4
                        If (LnCnt > 60) Then
                            If mScreen = True Then
                                Print #1, wDotLine
                            Else
                                Print #1, wDotLine + Chr(27) + Chr(33) + Chr(0) + Chr(12)
                            End If
                            wHdPrn = True
                            LnCnt = 0
                        End If
                    End If
                Case "T":      'T-Format Party Ledger
                    wDotLine = String(126, "-")
                    If (CrCnt > DrCnt) Then
                       wnp = CrCnt
                    Else
                       wnp = DrCnt
                    End If
                    t1 = 0
                    t2 = 0
                    For t3 = 1 To wnp
                        If wHdPrn Then
                           wHdPrn = False
                           If mScreen = True Then
                               Print #1, gCName
                               Print #1, "Party Ledger For Period From " & DTPicker1.Value & "  To  " & DTPicker2.Value & "           Page No. " + Str(PgCnt)
                               Print #1, wDotLine
                           Else
                               Print #1, Chr(18) + Chr(27) + Chr(14) & gCName & Chr(27) + Chr(18)
                               Print #1, Chr(27) + Chr(69) & "Party Ledger " & Chr(27) + Chr(70); " From " & DTPicker1.Value & "  To  " & DTPicker2.Value & "           Page No. " + Str(PgCnt)
                               Print #1, mP12 & wDotLine
                           End If
                           Print #1, "Credit Amt.   Date     Tp   Description                           Debit Amt.       Date     Tp   Description"
                           Print #1, wDotLine
                           LnCnt = 5
                           PgCnt = PgCnt + 1
                        End If
                        If wNamePrn Then
                           wNamePrn = False
                           wwaname = GProcGetColumnValue("tblmastaccount", "accode", CStr(wAcCode), "N", "acname", "S", , gCn)
                           Print #1, Space(60 - (Len(wwaname) / 2)) + wwaname
                           Print #1, Space(60 - (Len(wwaname) / 2)) + String(Len(wwaname), "-")
                           LnCnt = LnCnt + 2
                        End If
                        If (t3 <= CrCnt) Then
                           X = Left(CrSide(t3) + Space(63), 63)
                           t1 = t1 + Val(Left(CrSide(t3), 12))
                        Else
                           X = Space(63)
                        End If
                        If (t3 <= DrCnt) Then
                           X = X + DrSide(t3)
                           t2 = t2 + Val(Left(DrSide(t3), 12))
                        End If
                        Print #1, X
                        LnCnt = LnCnt + 1
                        If (LnCnt > 65) And (wRpTp = "D") Then
                            If mScreen = True Then
                                Print #1, wDotLine
                            Else
                                Print #1, wDotLine & Chr(12)
                            End If
                            wHdPrn = True
                            LnCnt = 0
                        End If
                    Next t3
                    If (t1 <> t2) Then
                       X = GProcNoToStr(Abs(t2 - t1), 12, 2)
                       If (t2 > t1) Then
                          t1 = t2
                          X = X + "  " + Format(DTPicker2.Value, "dd/MM/yyyy") + Space(6) + "Closing Balance"
                       Else
                          t2 = t1
                          X = Space(63) + X + "  " + Format(DTPicker2.Value, "dd/MM/yyyy") + Space(6) + "Closing Balance"
                       End If
                       Print #1, X
                    End If
                    Print #1, wDotLine
                    Print #1, GProcNoToStr(t1, 12, 2) + Space(51) + GProcNoToStr(t2, 12, 2)
                    Print #1, wDotLine + Chr(27) + Chr(18) + Chr(12)
                Case "M":      'Monthly Party Ledger
                    wDotLine = String(66, "-")
                    T11 = 0
                    t1 = 0
                    t2 = 0
                    For wnp = 0 To 12
                        If wHdPrn Then
                           wHdPrn = False
                           If mScreen = True Then
                               Print #1, gCName
                               Print #1, "Monthly Party Ledger For Period From " & DTPicker1.Value & "  To  " & DTPicker2.Value & "           Page No. " + Str(PgCnt)
                           Else
                               Print #1, Chr(18) + Chr(27) + Chr(14) & gCName & Chr(27) + Chr(18)
                               Print #1, Chr(27) + Chr(69) & "Monthly Party Ledger " & Chr(27) + Chr(70); " From " & DTPicker1.Value & "  To  " & DTPicker2.Value & "           Page No. " + Str(PgCnt)
                           End If
                           wwaname = GProcGetColumnValue("tblmastaccount", "accode", CStr(wAcCode), "N", "acname", "S", , gCn)
                           Print #1, Space(25 - (Len(wwaname) / 2)) + wwaname
                           Print #1, Space(25 - (Len(wwaname) / 2)) + String(Len(wwaname), "-")
                           Print #1, mP12 & wDotLine
                           Print #1, "Month                     Debit        Credit       Balance"
                           Print #1, wDotLine
                           LnCnt = 7
                           PgCnt = PgCnt + 1
                        End If
                        t3 = (wnp - Int(wnp / 10) * 10) + 3 - (Int(wnp / 10) * 2)
                        If (Val(CrSide(t3)) > 0) Or (Val(DrSide(t3)) > 0) Then
                        If (Val(DrSide(t3)) > 0) Then
                           X = DrSide(t3)
                           t1 = t1 + Val(DrSide(t3))
                           T11 = T11 - Val(DrSide(t3))
                        Else
                           X = Space(12)
                        End If
                        If (Val(CrSide(t3)) > 0) Then
                           X = X + "  " + CrSide(t3)
                           T11 = T11 + Val(CrSide(t3))
                           t2 = t2 + Val(CrSide(t3))
                        Else
                           X = X + "  " + Space(12)
                        End If
                        X = X + "  " + GProcNoToStr(Abs(T11), 12, 2)
                        If (T11 < 0) Then
                           X = X + "  Dr"
                        Else
                           X = X + "  Cr"
                        End If
                        Print #1, MonthName(Int(t3)) + Space(11) + X
                        LnCnt = LnCnt + 1
                        End If
                    Next wnp
                    Print #1, wDotLine
                    Print #1, Space(20) + GProcNoToStr(t1, 12, 2) + "  " + GProcNoToStr(t2, 12, 2)
                    Print #1, wDotLine + Chr(12)
                Case "P":      'Partywise Party Ledger
                    wDotLine = String(70, "-")
                    T11 = 0
                    t1 = 0
                    t2 = 0
                    For t3 = 0 To DrCnt
                        If wHdPrn Then
                           wHdPrn = False
                           If mScreen = True Then
                               Print #1, gCName
                               Print #1, "Accountwise Party Ledger For Period From " & DTPicker1.Value & "  To  " & DTPicker2.Value & "           Page No. " + Str(PgCnt)
                           Else
                               Print #1, Chr(18) + Chr(27) + Chr(14) & gCName & Chr(27) + Chr(18)
                               Print #1, Chr(27) + Chr(69) & "Accountwise Party Ledger " & Chr(27) + Chr(70); " From " & DTPicker1.Value & "  To  " & DTPicker2.Value & "           Page No. " + Str(PgCnt)
                           End If
                           wwaname = GProcGetColumnValue("tblmastaccount", "accode", CStr(wAcCode), "N", "acname", "S", , gCn)
                           Print #1, Tab(25 - (Len(wwaname) / 2)), wwaname
                           Print #1, Tab(25 - (Len(wwaname) / 2)), String(Len(wwaname), "-")
                           Print #1, mP12 & wDotLine
                           Print #1, "Account Name                                   Debit        Credit "
                           Print #1, wDotLine
                           LnCnt = 7
                           PgCnt = PgCnt + 1
                        End If
                        If (Val(DrSide(t3)) > 0) Then
                            X = Left(Mid(DrSide(t3), 25) + Space(40), 40)
                            X = X + "  " + Left(DrSide(t3), 12)
                            X = X + "  " + Mid(DrSide(t3), 13, 12)
                            t1 = t1 + Val(Left(DrSide(t3), 12))
                            t2 = t2 + Val(Mid(DrSide(t3), 13, 12))
                            Print #1, X
                            LnCnt = LnCnt + 1
                        End If
                    Next t3
                    Print #1, wDotLine
                    Print #1, Space(42) + GProcNoToStr(t1, 12, 2) + "  " + GProcNoToStr(t2, 12, 2)
                    Print #1, wDotLine + Chr(12)
                End Select
            Loop
        End With
        If wRpTp = "D" Then
            If mScreen = True Then
                Print #1, wDotLine + Chr(27) + Chr(33) + Chr(0) + Chr(27) + Chr(12)
            Else
                Print #1, wDotLine
            End If
        End If
'        If gSelectedmenu = "R23" Then
'            Print #1, wDotLine + Chr(12)
'            '----Display INDEX
'            Dim rstAc As Recordset
'            If GRstOpen(Cn, rstAc, "select ANAME, PGPR from ACCT where PGPR > 0 order by ANAME", "R") > 0 Then
'                PgCnt = 1
'                wHdPrn = True
'                LnCnt = 0
'                With rstAc
'                    .MoveFirst
'                    Do While Not .EOF
'                       If wHdPrn Then
'                            wHdPrn = False
'                            Print #1, "INDEX REPORT                               Page No. " + Str(PgCnt)
'                            Print #1, String(55, "-")
'                            Print #1, "ACCOUNT NAME                                    PAGE NO."
'                            Print #1, String(55, "-")
'                            LnCnt = 4
'                            PgCnt = PgCnt + 1
'                       End If
'                       Print #1, Left(!ANAME + Space(40), 40) + "  " + gprocnotostr(!PGPR, 13, 0)
'                       LnCnt = LnCnt + 1
'                       If (LnCnt > 65) Then
'                          If mScreen = True Then
'                            Print #1, String(55, "-")
'                          Else
'                            Print #1, String(55, "-") & Chr(12)
'                          End If
'                          wHdPrn = True
'                          LnCnt = 0
'                       End If
'                       .MoveNext
'                    Loop
'                    .Close
'                End With
'                Print #1, String(55, "-")
'            End If
'        End If
        
    Close #1
End If

Select Case wRpTp
Case "D", "T", "P", "M":
    Dim RetVal
    If mScreen = False Then
       RetVal = Shell(App.Path & "\txtprint.exe", vbMaximizedFocus)
    ElseIf mScreen = True Then
       RetVal = Shell(App.Path & "\WordPad c:\rpt1.txt", vbMaximizedFocus)
    End If
Case "F":
'    Dim wTotDr As Double, wTotCr As Double, q As Integer, wBal As Double, wCrDr As String
'    Dim wX As String
'    wTotCr = 0
'    wTotDr = 0
'    wBal = 0
'    frmRptFlexLed.msfLed.Clear
'    DoEvents
'    frmRptFlexLed.msfLed.FormatString = "<Date             |<Type  |>No          |<Account Name                                      |>Debit                        |>Credit                           |>Amount                  |       "
'    frmRptFlexLed.lblHead.Caption = (iisGFnGetColumnValue(Cn, "Select ANAME from acct where accode='" & wAcCode1 & "'")) & "  From  " & DTPicker1.Value & "  To  " & DTPicker2.Value
'    With frmRptFlexLed.msfLed
'        rstTran33.MoveFirst
'        Do While Not rstTran33.EOF
'            .TextMatrix(.Row, 0) = rstTran33!Vdt
'            .TextMatrix(.Row, 1) = Left(rstTran33!Nar1, 2)
'            .TextMatrix(.Row, 2) = rstTran33!DocNo & vbNullString
'            .TextMatrix(.Row, 3) = rstTran33!AcName & vbNullString
'            .TextMatrix(.Row, 4) = Format(rstTran33!Debit, "#######0.00")
'            .TextMatrix(.Row, 5) = Format(rstTran33!Credit, "#######0.00")
'            wTotDr = wTotDr + rstTran33!Debit
'            wTotCr = wTotCr + rstTran33!Credit
'            .TextMatrix(.Row, 6) = Format(Abs(wTotDr - wTotCr), "#######0.00")
'            If wTotDr > wTotCr Then
'                wCrDr = "Dr"
'            Else
'                wCrDr = "Cr"
'            End If
'            .TextMatrix(.Row, 7) = wCrDr
'            .Rows = .Rows + 1
'            .Row = .Row + 1
'            For j = 3 To 6
'                Select Case j
'                Case 3:
'                    wX = rstTran33!Nar3 & vbNullString
'                Case 4:
'                    wX = rstTran33!Nar4 & vbNullString
'                Case 5:
'                    wX = rstTran33!NAR5 & vbNullString
'                Case 6:
'                    wX = rstTran33!NAR6 & vbNullString
'                End Select
'                If Len(Trim(wX)) > 0 Then
'                    .TextMatrix(.Row, 3) = Space(5) + wX & vbNullString
'                    .Rows = .Rows + 1
'                    .Row = .Row + 1
'                End If
'           Next j
'           rstTran33.MoveNext
'        Loop
'        .Rows = .Rows + 1
'        .Row = .Row + 1
'        For q = 3 To 6
'            .Col = q
'            .CellFontBold = True
'            .CellForeColor = vbHighlight
'        Next q
'        .TextMatrix(.Row, 3) = "Total"
'        .TextMatrix(.Row, 4) = Format(Round(wTotDr, 2), "#############0.00")
'        .TextMatrix(.Row, 5) = Format(Round(wTotCr, 2), "#############0.00")
'    End With
'    frmRptFlexLed.Show 1
End Select
rstTran33.Close
End Sub
Private Sub DosHdg()
    Dim wAgCode As String, wReHead As String
    wAgCode = GProcGetColumnValue("tblmastaccount", "accode", CStr(wAcCode), "N", "agcode", "N", , gCn)
    wReHead = "Party Ledger"
    If wAgCode = 90019 Then
        wReHead = "Bank Book"
    End If
    
    If mScreen = True Then
       Print #1, gCName
       Print #1, wReHead & " For Period From " & DTPicker1.Value & "  To  " & DTPicker2.Value & "        Page No. " + Str(PgCnt)
       Print #1, wDotLine
    Else
       Print #1, Chr(18) + Chr(27) + Chr(14) & gCName & Chr(27) + Chr(18)
       Print #1, Chr(27) + Chr(69) & wReHead & " For Period " & Chr(27) + Chr(70) & " From " & DTPicker1.Value & "  To  " & DTPicker2.Value & "        Page No. " + Str(PgCnt)
       Print #1, mP12 & wDotLine
    End If
    If (wAgCode = 90019) Then
        Print #1, "   Date   Tp   No.     Account Name                         Receipt    Payment      Balance"
    Else
        Print #1, "   Date   Tp   No.     Account Name                           Debit     Credit      Balance"
    End If
    Print #1, wDotLine
    LnCnt = 5
    PgCnt = PgCnt + 1
End Sub
Private Sub txt_KeyPress(Index As Integer, KeyAscii As Integer)
Dim LeftPos As Long
Dim TopPos As Long
TopPos = Txt(Index).Top + Me.Top + 650
LeftPos = Txt(Index).Left + Me.Left
If KeyAscii = 13 Then Exit Sub
Select Case Index
    Case 0 '-Group
        gClsSearch.SearchMultiField "tblMastGroup", "AGName", Array("Account Group"), Array(Txt(Index).Width), "", Chr(KeyAscii), "AgName", LeftPos, TopPos
        KeyAscii = 0
        If gClsSearch.SearchMultiRetCol(0) <> "" Then Txt(Index) = gClsSearch.SearchMultiRetCol(0)
    Case 1 '-Account
        '-Selected Group
        If Txt(0) = "" Then
            gClsSearch.SearchMultiField "QryMastaccount", "Account,AcName,Accode,AcCity,AgName", Array("Account", "AcCode", "AcName", "City", "Group"), Array(3000, 0, 0, 2000, 3000), "", Chr(KeyAscii), "AcName", LeftPos, TopPos
        Else
            gClsSearch.SearchMultiField "QryMastaccount", "Account,AcName,Accode,AcCity,AgName", Array("Account", "AcCode", "AcName", "City", "Group"), Array(3000, 0, 0, 2000, 3000), "AgName='" & Txt(0) & "'", Chr(KeyAscii), "AcName", LeftPos, TopPos
        End If
        KeyAscii = 0
        If gClsSearch.SearchMultiRetCol(0) <> "" And gClsSearch.SearchMultiRetCol(1) = "" Then
            Txt(Index) = gClsSearch.SearchMultiRetCol(0)
            wAcCode = gClsSearch.SearchMultiRetCol(2)
        ElseIf gClsSearch.SearchMultiRetCol(0) <> "" And gClsSearch.SearchMultiRetCol(1) <> "" Then
            Txt(Index) = gClsSearch.SearchMultiRetCol(1)
            wAcCode = gClsSearch.SearchMultiRetCol(2)
        End If
    Case 2 '-Mill
        gClsSearch.SearchMultiField "QryMastaccount", "Account,AcName,Accode,AcCity,AgName", Array("Account", "AcCode", "AcName", "City", "Group"), Array(3000, 0, 0, 2000, 3000), " AgCode=90029", Chr(KeyAscii), "AcName", LeftPos, TopPos
        KeyAscii = 0
        If gClsSearch.SearchMultiRetCol(0) <> "" And gClsSearch.SearchMultiRetCol(1) = "" Then
            Txt(Index) = gClsSearch.SearchMultiRetCol(0)
            wAcCode = gClsSearch.SearchMultiRetCol(2)
        ElseIf gClsSearch.SearchMultiRetCol(0) <> "" And gClsSearch.SearchMultiRetCol(1) <> "" Then
            Txt(Index) = gClsSearch.SearchMultiRetCol(1)
            wAcCode = gClsSearch.SearchMultiRetCol(2)
        End If
End Select
End Sub
