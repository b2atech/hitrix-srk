VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form frmRptSelection 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Reports"
   ClientHeight    =   3240
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6420
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3240
   ScaleWidth      =   6420
   Begin VB.TextBox txtBrokerName 
      Appearance      =   0  'Flat
      Height          =   375
      Left            =   1680
      TabIndex        =   7
      Top             =   1440
      Width           =   4575
   End
   Begin VB.TextBox txtMillName 
      Appearance      =   0  'Flat
      Height          =   375
      Left            =   1680
      TabIndex        =   5
      Top             =   960
      Width           =   4575
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
      Height          =   495
      Left            =   3480
      TabIndex        =   9
      Top             =   2565
      Width           =   1335
   End
   Begin VB.CommandButton cmdCrystal 
      Caption         =   "&Genrate"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1605
      TabIndex        =   8
      Top             =   2565
      Width           =   1335
   End
   Begin Crystal.CrystalReport CrystalReport1 
      Left            =   240
      Top             =   2520
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
   Begin MSComctlLib.ProgressBar prgBar 
      Height          =   255
      Left            =   120
      TabIndex        =   10
      Top             =   2085
      Visible         =   0   'False
      Width           =   6135
      _ExtentX        =   10821
      _ExtentY        =   450
      _Version        =   393216
      Appearance      =   1
   End
   Begin MSComCtl2.DTPicker dtpToDt 
      Height          =   360
      Left            =   4680
      TabIndex        =   3
      Top             =   480
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
      Format          =   51970051
      CurrentDate     =   38050
   End
   Begin MSComCtl2.DTPicker dtpFrDt 
      Height          =   360
      Left            =   1680
      TabIndex        =   1
      Top             =   480
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
      Format          =   51970051
      CurrentDate     =   38050
   End
   Begin VB.Label lblBrokerName 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Broker Name  :"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   120
      TabIndex        =   6
      Top             =   1440
      Width           =   1425
   End
   Begin VB.Label lblMillName 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Mill Name       :"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   120
      TabIndex        =   4
      Top             =   960
      Width           =   1440
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "To"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   3960
      TabIndex        =   2
      Top             =   495
      Width           =   360
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "From Date       :"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   120
      TabIndex        =   0
      Top             =   495
      Width           =   1470
   End
End
Attribute VB_Name = "frmRptSelection"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim wBrokerCode As Long, wMillCode As Long
Private Sub cmdCrystal_Click()
Dim x1 As String
x1 = ""
Screen.MousePointer = vbHourglass
With CrystalReport1
    Select Case gSelectedMenu
        Case "Stock Statement Datewise"
            GprocCollectPurStock dtpFrDt, dtpToDt
            GProcCrystalRptPreparation CrystalReport1, "Window"
            .Formulas(4) = "wFrToDt = 'From : ' & '" & Format(dtpFrDt.Value, "dd/MM/yyyy") & "' & '  To ' & '" & Format(dtpToDt.Value, "dd/MM/yyyy") & "'"
            .ReportFileName = gReportPath & "rptStkRptDatewise.rpt"
            .SelectionFormula = "{tmpGentbl.UserName}='" & gUserName & "'"
            .Formulas(1) = ""
            .Formulas(2) = ""
            .Formulas(3) = ""
        Case "Pending Purchase SIT Register" '---Pending Purchase SIT Register
            GProcCrystalRptPreparation CrystalReport1, "Window"
            .Formulas(4) = "wFrToDt = 'From : ' & '" & Format(dtpFrDt.Value, "dd/MM/yyyy") & "' & '  To ' & '" & Format(dtpToDt.Value, "dd/MM/yyyy") & "'"
            x1 = "{tblAddLess.VDt} in Date(" & Format(dtpFrDt.Value, "yyyy,mm,dd") & ") to Date(" & Format(dtpToDt.Value, "yyyy,mm,dd") & ")"
            x1 = x1 + " and {tblAddLess.VType}='PT' and ({tblVouSub.VSubBag}-{tblVouSub.VSubIssBag}>0)"
            If txtMillName = "" Then
                .SelectionFormula = x1
                .ReportFileName = gReportPath & "rptStkRptPendingInward.rpt"
           Else '-- Mill Wise
                x1 = x1 + " and {tblMastAccount.AcName}='" & txtMillName & "'"
                .Formulas(7) = "WMillName='Mill Name : " & txtMillName & "'"
                 .SelectionFormula = x1
                .ReportFileName = gReportPath & "rptStkRptPendingInwardMill.rpt"
           End If
    End Select
    .WindowState = crptMaximized
    .Action = 1
End With
Screen.MousePointer = vbDefault
End Sub
Private Sub cmdExit_Click()
    DoEvents
    Unload Me
End Sub
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
    Case 13:
       SendKeys "{TAB}"
    Case 27:
        Unload Me
    End Select
End Sub
Private Sub Form_Load()
Me.Top = 700
Me.Left = Screen.Width / 2 - Me.Width / 2
dtpFrDt.Value = Format(gCYSDate, "dd/MM/yyyy")
dtpToDt.Value = Format(gCYEDate, "dd/MM/yyyy")
Select Case gSelectedMenu
Case "RCREC":
    Me.Caption = "C Form Receivable"
Case "RCISS":
    Me.Caption = "C Form Issuable"
Case "RE1REC":
    Me.Caption = "E1 Form Receivable"
Case "RCFORM":
    Me.Caption = "Requirement of C Form"
Case "REFORM":
    Me.Caption = "Requirement of E Form"
Case "RFSTATE":
    Me.Caption = "F Form Statement"
Case "Stock Statement Datewise"
    Me.Caption = gfrmTypeStr
    txtBrokerName.Visible = False
    lblBrokerName.Visible = False
Case "Pending Purchase SIT Register" '-- "Pending Purchase SIT Register"
    Me.Caption = gSelectedMenu
    txtBrokerName.Visible = False
    lblBrokerName.Visible = False
End Select
End Sub

Private Sub txtMillName_GotFocus()
GProcSelectBox Me.ActiveControl
End Sub

Private Sub txtMillName_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then Exit Sub '-- Enter
gClsSearch.SearchMultiField "QryMastaccount", "Account,AcName,ACCode", Array("Mill", "AcName", ""), Array(txtMillName.Width, 0, 0), " (AgCode=90029 or GpCode=90029)", Chr(KeyAscii), "AcName", txtMillName.Left + Me.Left, txtMillName.Top + Me.Top + 1000, False
If gClsSearch.SearchMultiRetCol(1) <> "" Then
    txtMillName = gClsSearch.SearchMultiRetCol(1)
    KeyAscii = 0
End If
End Sub
