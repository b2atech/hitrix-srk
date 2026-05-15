VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form frmRptFinal 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Report"
   ClientHeight    =   4530
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   13020
   Icon            =   "frmRptFinal.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4530
   ScaleWidth      =   13020
   Begin VB.Frame Frame1 
      Height          =   3240
      Left            =   5700
      TabIndex        =   14
      Top             =   60
      Visible         =   0   'False
      Width           =   6270
      Begin VB.TextBox txtStAmount 
         Alignment       =   1  'Right Justify
         Height          =   375
         Index           =   0
         Left            =   4200
         TabIndex        =   17
         Top             =   720
         Width           =   1695
      End
      Begin VB.CommandButton cmdOk 
         Caption         =   "&OK"
         Height          =   375
         Left            =   2160
         TabIndex        =   16
         Top             =   2640
         Width           =   1455
      End
      Begin VB.TextBox txtstAccode 
         Height          =   315
         Index           =   0
         Left            =   2880
         TabIndex        =   15
         Text            =   "stAcCode"
         Top             =   720
         Visible         =   0   'False
         Width           =   735
      End
      Begin VB.Shape Shape3 
         FillColor       =   &H00BDDFD1&
         FillStyle       =   0  'Solid
         Height          =   585
         Left            =   2040
         Shape           =   4  'Rounded Rectangle
         Top             =   2520
         Width           =   1785
      End
      Begin VB.Label lblActualStkValue 
         BackStyle       =   0  'Transparent
         Caption         =   "LblActualStkValue"
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
         Height          =   495
         Left            =   120
         TabIndex        =   22
         Top             =   2520
         Width           =   3540
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "Account Name"
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
         Height          =   255
         Left            =   360
         TabIndex        =   20
         Top             =   240
         Width           =   1815
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00000080&
         X1              =   0
         X2              =   6000
         Y1              =   600
         Y2              =   600
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "Closing Stock"
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
         Height          =   255
         Left            =   4440
         TabIndex        =   19
         Top             =   240
         Width           =   1095
      End
      Begin VB.Label lblStAccount 
         BackStyle       =   0  'Transparent
         Caption         =   "Label6"
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
         Height          =   375
         Index           =   0
         Left            =   360
         TabIndex        =   18
         Top             =   720
         Width           =   3615
      End
      Begin VB.Shape Shape1 
         BackColor       =   &H00808080&
         BorderColor     =   &H004D177D&
         FillColor       =   &H00ECFFFF&
         FillStyle       =   0  'Solid
         Height          =   3135
         Left            =   0
         Top             =   240
         Width           =   6240
      End
   End
   Begin VB.CommandButton cmdPrint 
      Height          =   480
      Index           =   1
      Left            =   1740
      Picture         =   "frmRptFinal.frx":058A
      Style           =   1  'Graphical
      TabIndex        =   11
      ToolTipText     =   "Print"
      Top             =   3795
      Width           =   1185
   End
   Begin VB.CommandButton cmdPrint 
      Height          =   480
      Index           =   0
      Left            =   390
      Picture         =   "frmRptFinal.frx":0BF4
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   3795
      Width           =   1185
   End
   Begin VB.CommandButton cmdClosing 
      Caption         =   "Closing Stock"
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
      Left            =   4275
      TabIndex        =   13
      Top             =   3795
      Width           =   1710
   End
   Begin Crystal.CrystalReport cryRepo 
      Left            =   7020
      Top             =   1920
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
   Begin VB.TextBox txtTemp 
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
      Left            =   1800
      TabIndex        =   5
      Top             =   870
      Width           =   5940
   End
   Begin VB.CommandButton cmdExit 
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
      Left            =   3030
      TabIndex        =   12
      Top             =   3795
      Width           =   1185
   End
   Begin MSComCtl2.DTPicker dtpFrDt 
      Height          =   375
      Left            =   1800
      TabIndex        =   1
      Top             =   330
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   661
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
      Format          =   97058819
      CurrentDate     =   38050
   End
   Begin MSComCtl2.DTPicker dtpToDt 
      Height          =   375
      Left            =   5925
      TabIndex        =   3
      Top             =   330
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   661
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
      Format          =   97058819
      CurrentDate     =   38050
   End
   Begin MSComctlLib.ProgressBar prgBar1 
      Height          =   315
      Left            =   195
      TabIndex        =   21
      Top             =   3000
      Visible         =   0   'False
      Width           =   5535
      _ExtentX        =   9763
      _ExtentY        =   556
      _Version        =   393216
      Appearance      =   1
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00ECFFFF&
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
      Height          =   1095
      Left            =   195
      TabIndex        =   6
      Top             =   1290
      Width           =   3060
      Begin VB.TextBox TxtAmount 
         Height          =   375
         Left            =   1680
         TabIndex        =   30
         Top             =   480
         Visible         =   0   'False
         Width           =   1335
      End
      Begin VB.OptionButton optReportOn 
         Caption         =   "&Windows"
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
         Height          =   315
         Index           =   3
         Left            =   0
         TabIndex        =   24
         Top             =   0
         Value           =   -1  'True
         Visible         =   0   'False
         Width           =   210
      End
      Begin VB.OptionButton optReportOn 
         Caption         =   "MS &DOS"
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
         Height          =   315
         Index           =   0
         Left            =   0
         TabIndex        =   7
         Top             =   0
         Visible         =   0   'False
         Width           =   225
      End
      Begin VB.OptionButton optReportOn 
         Caption         =   "&Windows"
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
         Height          =   315
         Index           =   1
         Left            =   0
         TabIndex        =   8
         Top             =   0
         Visible         =   0   'False
         Width           =   210
      End
      Begin VB.OptionButton optReportOn 
         Caption         =   "&Grid"
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
         Height          =   315
         Index           =   2
         Left            =   5280
         TabIndex        =   9
         Top             =   720
         Visible         =   0   'False
         Width           =   210
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Balance Above :"
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
         Left            =   120
         TabIndex        =   23
         Top             =   480
         Visible         =   0   'False
         Width           =   1410
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00ECFFFF&
      ForeColor       =   &H00FFFFFF&
      Height          =   1335
      Left            =   3360
      TabIndex        =   25
      Top             =   1200
      Width           =   2760
      Begin VB.CheckBox chkIsWithCapital 
         BackColor       =   &H00ECFFFF&
         Caption         =   " With Capital Name"
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
         Height          =   195
         Left            =   120
         TabIndex        =   31
         Top             =   1020
         Visible         =   0   'False
         Width           =   2475
      End
      Begin VB.OptionButton OptListconf 
         BackColor       =   &H00ECFFFF&
         Caption         =   "Option1"
         Height          =   255
         Index           =   0
         Left            =   120
         MaskColor       =   &H00ECFFFF&
         TabIndex        =   27
         Top             =   240
         Visible         =   0   'False
         Width           =   255
      End
      Begin VB.OptionButton OptListconf 
         BackColor       =   &H00ECFFFF&
         Caption         =   "Option1"
         Height          =   315
         Index           =   1
         Left            =   120
         MaskColor       =   &H00ECFFFF&
         TabIndex        =   26
         Top             =   600
         Visible         =   0   'False
         Width           =   255
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "List "
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
         Height          =   360
         Left            =   480
         TabIndex        =   29
         Top             =   240
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Confirmation Letter"
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
         Left            =   480
         TabIndex        =   28
         Top             =   600
         Visible         =   0   'False
         Width           =   1635
      End
   End
   Begin VB.Shape Shape5 
      FillColor       =   &H00BDDFD1&
      FillStyle       =   0  'Solid
      Height          =   705
      Left            =   255
      Shape           =   4  'Rounded Rectangle
      Top             =   3720
      Width           =   6000
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "&Group Name :"
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
      Left            =   210
      TabIndex        =   4
      Top             =   855
      Width           =   1215
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
      ForeColor       =   &H00000080&
      Height          =   240
      Left            =   5505
      TabIndex        =   2
      Top             =   375
      Width           =   240
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "From &Date       :"
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
      Left            =   210
      TabIndex        =   0
      Top             =   330
      Width           =   1470
   End
   Begin VB.Shape Shape2 
      BackColor       =   &H00C0FFFF&
      BorderColor     =   &H004D177D&
      FillColor       =   &H00ECFFFF&
      FillStyle       =   0  'Solid
      Height          =   3315
      Left            =   120
      Top             =   120
      Width           =   8475
   End
End
Attribute VB_Name = "frmRptFinal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim srcnt As Integer, LnCnt As Integer, PgCnt As Integer, mScreen As Boolean
Dim wHdPrn As Boolean, wCntPrn As Boolean, wDotLine As String, mP12 As String
Dim wAgCode As Long, RstClosingStock As Recordset, Maxi As Integer, rsttmpGenTbl As Recordset
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
Select Case KeyCode
    Case 13:
       Sendkeys "{TAB}"
    Case 27:
        Unload Me
End Select
End Sub
Private Sub Form_Load()
Maxi = 0
Me.Top = 700
Me.Left = Screen.Width / 2 - Me.Width / 2
mScreen = True
 '-- Activate Controls to False
GProcActivateControls Me, False
dtpFrDt.Enabled = True
dtpToDt.Enabled = True
dtpFrDt.Value = Format(gCYSDate, "dd/MM/yyyy")
dtpToDt.Value = Format(gCYEDate, "dd/MM/yyyy")
Me.Caption = gSelectedMenu
optReportOn(0).Enabled = True
optReportOn(1).Enabled = True
OptListconf(0).Enabled = True
OptListconf(1).Enabled = True
TxtAmount.Enabled = True
lblActualStkValue = ""
Select Case gSelectedMenu
    Case "Schedule Printing", "Expences Shedule", "Opposite A/c wise Summery", "Loan A/c Details", "Asset A/c Details", "Asset A/c Details" '-- "Schedule Printing"
        txtTemp(0).Enabled = True '-Group
        If gSelectedMenu = "Schedule Printing" Then
          OptListconf(0).Value = True
          OptListconf(0).Visible = True
          OptListconf(0).Enabled = True
          Label6.Visible = True
          Label6.Caption = "Without Capital Name"
          
          OptListconf(1).Visible = True
          OptListconf(1).Enabled = True
          Label7.Visible = True
          Label7.Caption = "With Capital Name"
        End If
        If gSelectedMenu = "Expences Shedule" Then
           OptListconf(0).Visible = True
           OptListconf(0).Enabled = True
           OptListconf(1).Visible = True
           OptListconf(1).Enabled = True
           Label6.Visible = True
           Label7.Visible = True
           Label6.Caption = "All Accounts"
           Label7.Caption = "For Audit   "
           OptListconf(0).Value = True
           chkIsWithCapital.Visible = True
           chkIsWithCapital.Enabled = True
           chkIsWithCapital.Value = 0
        End If
    Case "Trial Balance Simple", "Trial Balance Detail", "Trial Balance Group Simple", "Trial Balance Group Detail", "T-Format", "Balance Sheet"
       txtTemp(0).Visible = False
       Label3.Visible = False
       If gSelectedMenu = "Balance Sheet" Then
          OptListconf(0).Value = True
          OptListconf(0).Visible = True
          OptListconf(0).Enabled = True
          Label6.Visible = True
          Label6.Caption = "Without Capital Name"
          
          OptListconf(1).Visible = True
          OptListconf(1).Enabled = True
          Label7.Visible = True
          Label7.Caption = "With Capital Name"
       End If
       'Check2.Visible = True
       'Label3.Left = 550
       'Label3.Caption = "Debtors and Creditors Total   "
    Case "Debitor/Creditor List Beyound Limit" '-- "Debitor/Creditor List Beyound Limit"
        txtTemp(0).Enabled = True '-Group
        Label6.Visible = True
        Label7.Visible = True
        Label8.Visible = True
        optReportOn(0).Visible = True
        optReportOn(1).Visible = True
        optReportOn(1).Value = True
        OptListconf(0).Visible = True
        OptListconf(1).Visible = True
        OptListconf(0).Value = True
        TxtAmount.Visible = True
    Case "Loan Reporting"
        txtTemp(0).Enabled = True '-Group
        Label6.Visible = True
    Case "Profit/Loss Account"
          OptListconf(0).Value = True
          OptListconf(0).Visible = True
          OptListconf(0).Enabled = True
          Label6.Visible = True
          Label6.Caption = "Without Capital Name"
          
          OptListconf(1).Visible = True
          OptListconf(1).Enabled = True
          Label7.Visible = True
          Label7.Caption = "With Capital Name"
End Select
End Sub
Private Sub dtpFrDt_GotFocus()
GProcSelectBox Me.ActiveControl
End Sub
Private Sub dtpToDt_GotFocus()
GProcSelectBox Me.ActiveControl
End Sub
Private Sub cmdExit_Click()
    Unload Me
End Sub
Private Sub Form_Unload(Cancel As Integer)
Screen.MousePointer = vbDefault
End Sub
Private Sub TxtAmount_KeyPress(KeyAscii As Integer)
  KeyAscii = GProcNumberOnly(KeyAscii)
End Sub

Private Sub txtStAmount_LostFocus(Index As Integer)
    txtStAmount(Index) = Format(txtStAmount(Index), "###0.00")
End Sub
Private Sub txttemp_GotFocus(Index As Integer)
GProcSelectBox Me.ActiveControl
End Sub
Private Sub txtTemp_KeyPress(Index As Integer, KeyAscii As Integer)
Dim X As String
Select Case Index
    Case 0 '- Group
        If gSelectedMenu = "Expences Shedule" Then
           X = "select AcOurCode,acname,AcCode from tblMastAccount where AcCode in (select Distinct SubGpCode from tblMastSubGroup where SubName <> '-') "
           gClsSearch.SearchMultiField X, "", Array("Code", "A/c Name", ""), Array(1200, 5000, 0), "", "", "AcOurCode,acname", 0, 0, False, , True
           KeyAscii = 0
           If gClsSearch.SearchMultiRetCol(0) <> "" Then
              txtTemp(0) = gClsSearch.SearchMultiRetCol(1)
              wAgCode = gClsSearch.SearchMultiRetCol(2)
           End If
        Else
           gClsSearch.SearchMultiField "tblMastGroup", "AGName,Agcode", Array("Account Group", "Agcode"), Array(txtTemp(Index).Width, 0), " Not(Agcode = TopGroup)  ", Trim(Chr(KeyAscii)), "", txtTemp(Index).Left + Me.Left, txtTemp(Index).Top + Me.Top + 1000, False, , False
           KeyAscii = 0
           If gClsSearch.SearchMultiRetCol(0) <> "" Then
              txtTemp(0) = gClsSearch.SearchMultiRetCol(0)
              wAgCode = gClsSearch.SearchMultiRetCol(1)
           End If
        End If
    Case 5 '- Group
        gClsSearch.SearchMultiField "tblMastGroup", "AGName,AgCode", Array("Account Group", ""), Array(txtTemp(0).Width, 0), "", Chr(KeyAscii), "AgName", txtTemp(Index).Left + Me.Left, txtTemp(Index).Top + Me.Top + 1000
        KeyAscii = 0
        If gClsSearch.SearchMultiRetCol(0) <> "" Then
            txtTemp(0) = gClsSearch.SearchMultiRetCol(0)
            wAgCode = gClsSearch.SearchMultiRetCol(1)
        End If

End Select
End Sub
Private Sub cmdPrint_Click(Index As Integer)
Dim x1 As String
Dim x2 As String
Dim x3 As String
Dim RetVal
Dim tmpRst As Recordset
Dim tmpRst2 As Recordset
Dim tmpRst3 As Recordset
Dim PfAmt As Double
Dim PfCurrYrCd As Long
Screen.MousePointer = vbHourglass
'=== Window Report
'If optReportOn(1) = True Then

'If gSelectedMenu = "Loan Reporting" And txtTemp(0) = "" Then
'    MsgBox "Please Select Group Name"
'    txtTemp(0).SetFocus
''    Screen.MousePointer = vbDefault
'    Exit Function
'End If

With cryRepo
    If txtTemp(0) = "" Then wAgCode = 0 '-- Group
    GProcCrystalRptPreparation cryRepo, IIf(Index = 0, "Window", "Print")
    Select Case gSelectedMenu
        Case "Balance Sheet" '--Balance Sheet
           ' cryBalSht ("C")
             GProcPrepareBalanceSheet dtpFrDt.Value, dtpToDt.Value
            .Formulas(4) = "wFrToDt='From : '& '" & dtpFrDt.Value & "' & ' To ' & '" & dtpToDt.Value & "'"
            .Formulas(5) = "LcHead2= '" & dtpFrDt.Value & "'"
            .Formulas(6) = "wCPAN = 'PAN No  :  ' & '" & gCPAN & "'"
            .Formulas(7) = "wCapitalNm = '" & gCCapitalName & "'"
            .Formulas(8) = "wDesi = '(" & gCDesignation & ")'"
            If OptListconf(0).Value = True Then
               .Formulas(9) = "wPrintCapital = 'No'"
            Else
               .Formulas(9) = "wPrintCapital = 'Yes'"
            End If
            
'            X1 = "select * from TmpGentbl  where username = '" & gUserName & "' and vfirm = '" & gCCode & "' and left(ltrim(nar13),15) = 'As Per Schedule'"
'            gCn.BeginTrans
'            If (GProcRstOpen(tmpRst, X1, "O") > 0) Then
'               tmpRst.MoveFirst
'               PfAmt = 65
'               Do While (Not tmpRst.EOF)
'                  tmpRst.Fields("Nar13") = Left(tmpRst.Fields("Nar13"), InStr(tmpRst.Fields("Nar13"), "Schedule") + 8) + " " + Chr(PfAmt)
'                  PfAmt = PfAmt + 1
'                  tmpRst.Update
'                  tmpRst.MoveNext
'                Loop
'            End If
'            X1 = "select * from TmpGentbl  where username = '" & gUserName & "' and vfirm = '" & gCCode & "' and left(ltrim(nar7),15) = 'As Per Schedule'"
'            If (GProcRstOpen(tmpRst, X1, "O") > 0) Then
'               tmpRst.MoveFirst
'               Do While (Not tmpRst.EOF)
'                  tmpRst.Fields("Nar7") = Left(tmpRst.Fields("Nar7"), InStr(tmpRst.Fields("Nar7"), "Schedule") + 8) + " " + Chr(PfAmt)
'                  PfAmt = PfAmt + 1
'                  tmpRst.Update
'                  tmpRst.MoveNext
'                Loop
'             End If
'             gCn.CommitTrans
            .SelectionFormula = "{tmpgenTbl.Username}='" & gUserName & "' and {tmpgenTbl.VFirm}='" & gCCode & "'   and {tmpGenTbl.Vtype}='B'"
            .ReportFileName = gReportPath & "rptBalanceSheet.rpt"
        Case "Profit/Loss Account" '--"Profit/Loss Account"
            gCn.Execute "delete from tmpSelection where Vfirm = '" & Trim(gCCode) & "' and  UserName='" & gUserName & "'"
            gCn.Execute "insert into tmpSelection (dt1,dt2,cfafr,cfato,UserName,Vfirm) values ('" & Format(dtpFrDt, "yyyy/mm/dd") & "','" & Format(dtpToDt, "yyyy/mm/dd") & "','" & Format(gCYSDate, "yyyy/mm/dd") & "','" & Format(gCYEDate, "yyyy/mm/dd") & "','" & gUserName & "','" & gCCode & "')"
        
            GProcPrepareProfitLoss dtpFrDt.Value, dtpToDt.Value
           ' Prof_Loss_Ac ("C")
            .Formulas(4) = "wFrToDt='From : '& '" & dtpFrDt.Value & "' & ' To ' & '" & dtpToDt.Value & "'"
            .Formulas(5) = "wCPAN = 'PAN No : ' & '" & gCPAN & "'"
            .Formulas(6) = "wCapitalNm = '" & gCCapitalName & "'"
            .Formulas(7) = "wDesi = '(" & gCDesignation & ")'"
            If OptListconf(0).Value = True Then
               .Formulas(8) = "wPrintCapital = 'No'"
            Else
               .Formulas(8) = "wPrintCapital = 'Yes'"
            End If
            .ReportFileName = gReportPath & "RptPandL.rpt"
            .SelectionFormula = "{tmpgenTbl.Username}='" & gUserName & "' and {tmpgenTbl.VFirm}= '" & gCCode & "'  "
        Case "Expences Shedule"
                 gCmd.CommandText = "PrcPrepareSubAcShedule"
                 gCmd.Parameters.Refresh
                 gCmd.Parameters("@VFirm") = Trim(gCCode)
                 gCmd.Parameters("@From_dt") = Format(dtpFrDt.Value, "yyyy/mm/dd")
                 gCmd.Parameters("@To_dt") = Format(dtpToDt.Value, "yyyy/mm/dd")
                 gCmd.Parameters("@UCode") = gUserName
                 If txtTemp(0) = "" Then
                    wAgCode = 0
                 End If
                 If wAgCode <> 0 Then gCmd.Parameters("@Accode") = wAgCode
                 If OptListconf(1).Value = True Then
                    gCmd.Parameters("@ForAudit") = 1
'                 Else
 '                   gCmd.Parameters("@ForAudit") = 0
                 End If
                 gCmd.Execute
                 .Formulas(4) = "wFrToDt='From : '& '" & dtpFrDt.Value & "' & ' To ' & '" & dtpToDt.Value & "'"
                 .Formulas(5) = "wCPAN = 'PAN No : ' & '" & gCPAN & "'"
                 .Formulas(6) = "wCapitalNm = '" & gCCapitalName & "'"
                 .Formulas(7) = "wDesi = '(" & gCDesignation & ")'"
                 If chkIsWithCapital.Value = 0 Then
                    .Formulas(8) = "wPrintCapital = 'No'"
                 Else
                    .Formulas(8) = "wPrintCapital = 'Yes'"
                 End If
                 x1 = "{tmpAccLedger.username}='" & gUserName & "' "
                 .SelectionFormula = x1
                 .ReportFileName = gReportPath & "rptAccSubAcSchedule.rpt"
        
        Case "Trial Balance Simple", "Trial Balance Detail", "Trial Balance Group Simple", "Trial Balance Group Detail", "Schedule Printing", "All Schedule Printing", "Debitor/Creditor List Beyound Limit" '== Trial Balance ,Schedule,Debitor/Creditor List Beyound Limit
            '---- insert From dt ,To Dt in tmpGenTbl
                 gCmd.CommandText = "PrcPrepareClBalance"
                 gCmd.Parameters.Refresh
'                 gCmd.Parameters("@From_dt") = "'" & Format(FromDt, "yyyy/mm/dd") & "'"
'                 If ChkConsolidate.Value = 1 Then
'                    gCmd.Parameters("@VFirm") = ""
'                 Else
                    gCmd.Parameters("@VFirm") = Trim(gCCode)
'                 End If
                 gCmd.Parameters("@Fr_dt") = Format(dtpFrDt.Value, "yyyy/mm/dd")
                 gCmd.Parameters("@To_dt") = Format(dtpToDt.Value, "yyyy/mm/dd")
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
                 .Formulas(4) = "wFrToDt='From : '& '" & dtpFrDt.Value & "' & ' To ' & '" & dtpToDt.Value & "'"
                Select Case gSelectedMenu
                Case "Trial Balance Simple" '-simple
                    .SelectionFormula = "{qryTrbalDetail.username}='" & gUserName & "' and {qryTrbalDetail.ACGRPTYPE}='A'"
                    .ReportFileName = gReportPath & "RptTrialBalanceSim.rpt"
                Case "Trial Balance Detail"  '-Detail
                    .SelectionFormula = "{qryTrbalDetail.username}='" & gUserName & "' and {qryTrbalDetail.ACGRPTYPE}='A'"
                    .ReportFileName = gReportPath & "rptTrialBalanceDet.rpt"
                Case "Schedule Printing" '-- Schedule Printing
                    x1 = "{TmpClosingBalance.username}='" & gUserName & "' and {@Sh} > 0    and {TmpClosingBalance.Vfirm} = '" & Trim(gCCode) & "'"
                    If wAgCode > 0 Then '-- Group
                        x1 = x1 + " and {TmpClosingBalance.AgCode}=" & wAgCode
                    End If
                    .Formulas(5) = "wCPAN = 'PAN No : ' & '" & gCPAN & "'"
                    .Formulas(6) = "wCapitalNm = '" & gCCapitalName & "'"
                    .Formulas(7) = "wDesi = '(" & gCDesignation & ")'"
                    If OptListconf(0).Value = True Then
                       .Formulas(8) = "wPrintCapital = 'No'"
                    Else
                       .Formulas(8) = "wPrintCapital = 'Yes'"
                    End If
                    .SelectionFormula = x1
                    .ReportFileName = gReportPath & "rptAccSchedule.rpt"
                Case "All Schedule Printing" '-- "All Schedule Printing"
                    x1 = "{qryTrbalDetail.username}='" & gUserName & "' and {qryTrbalDetail.ACGRPTYPE}='A' and isnull({tblMastGroup.Schedule})=False"
                    .SelectionFormula = x1
                    .ReportFileName = gReportPath & "rptAccSchedule.rpt"
                Case "Trial Balance Group Simple", "Trial Balance Group Detail"
                
                     PfCurrYrCd = GProcGetColumnValue("tblMastAccount", "AcName", "Profit And Loss Of Current Year", "S", "accode", "N", "", gCn)
                     x1 = "select * from tblVoucher where (VCrAcCode = " & PfCurrYrCd & " or VDrAcCode = " & PfCurrYrCd & ") and vfirm = '" & gCCode & "' and vyear = " & gCYear & " "
                     If (GProcRstOpen(tmpRst, x1, "R") > 0) Then
                        PfAmt = 0
                        x1 = "select StAcCode,StAmount   from tblMastAccount,tblClosingStock  where tblMastAccount.AcCode = tblClosingStock.StAcCode " _
                              & " and tblMastAccount.AcName = 'Yarn Stock' and tblClosingStock.stDate = ('" & Format(dtpToDt.Value, "yyyy/mm/dd") & "')  and tblClosingStock.Vfirm = '" & gCCode & "'"
                        If GProcRstOpen(tmpRst, x1, "R", gCn) > 0 Then
                           gCn.BeginTrans
                           x3 = "Select * from TmpClosingBalance where UserName='" & gUserName & "' and Vfirm = '" & gCCode & "' and TmpClosingBalance.AcCode = " & tmpRst.Fields(0) & "  "
                           If GProcRstOpen(tmpRst3, x3, "R", gCn) = 0 Then
                              x2 = "Select tblMastAccount.AcCode,tblMastAccount.AcName,tblMastAccount.AgCode,tblMastGroup.AgName from tblMastAccount,tblMastGroup where tblMastAccount.AcCode = " & tmpRst.Fields(0) & " and " _
                                   & " tblMastAccount.AgCode = tblMastGroup.AgCode "
                              If GProcRstOpen(tmpRst2, x2, "R", gCn) > 0 Then
                                 gCn.Execute "insert into TmpClosingBalance (Vsrno,Vfirm,VFirmName,AcCode,AcName,AgCode,AgName,UserName) values (2,'" & gCCode & "','" & gCName & "'," & tmpRst.Fields(0) & ",'" & tmpRst2.Fields(1) & "'," & tmpRst2.Fields(2) & " ,'" & tmpRst2.Fields(3) & "',  '" & gUserName & "')"
                              End If
                              tmpRst2.Close
                           End If
                           tmpRst3.Close
                           
                           gCn.Execute "update TmpClosingBalance set ClBal_Dr = " & tmpRst.Fields(1) & " where UserName='" & gUserName & "' and Vfirm = '" & gCCode & "' and AcCode = " & tmpRst.Fields(0) & " "
                           gCn.CommitTrans
                        End If
                        x1 = "select sum(ClBal_Dr)-sum(OpBal_Dr) from TmpClosingBalance where TmpClosingBalance.Vfirm = '" & Trim(gCCode) & "' and TmpClosingBalance.username = '" & gUserName & "' and TmpClosingBalance.Agcode = 90018 "
                        If (GProcRstOpen(tmpRst, x1, "R") > 0) Then
                           PfAmt = PfAmt + IIf(IsNull(tmpRst.Fields(0)), 0, tmpRst.Fields(0))
                        End If
                        x1 = "select sum(ClBal_Cr)-sum(ClBal_Dr) from TmpClosingBalance where TmpClosingBalance.Vfirm = '" & Trim(gCCode) & "' and TmpClosingBalance.username = '" & gUserName & "' and TmpClosingBalance.TopGrp = 90003 "
                        If (GProcRstOpen(tmpRst, x1, "R") > 0) Then
                           PfAmt = PfAmt + tmpRst.Fields(0)
                        End If
                        If PfAmt > 0 Then
                           gCn.Execute "update TmpClosingBalance set Credit = Credit + " & PfAmt & "  where UserName='" & gUserName & "' and Vfirm = '" & gCCode & "' and AcName = 'Profit And Loss Of Current Year'"
                           gCn.Execute "update TmpClosingBalance set ClBal_Dr = Debit-Credit,ClBal_Cr = Credit-Debit  where UserName='" & gUserName & "' and Vfirm = '" & gCCode & "' and AcName = 'Profit And Loss Of Current Year'"

                        Else
                           gCn.Execute "update TmpClosingBalance set Debit = Debit + " & Abs(PfAmt) & "  where UserName='" & gUserName & "' and Vfirm = '" & gCCode & "' and AcName = 'Profit And Loss Of Current Year'"
                           gCn.Execute "update TmpClosingBalance set ClBal_Dr = Debit-Credit,ClBal_Cr = Credit-Debit  where UserName='" & gUserName & "' and Vfirm = '" & gCCode & "' and AcName = 'Profit And Loss Of Current Year'"
                        
                        End If
'                         gCn.Execute "Delete TmpClosingBalance  where UserName='" & gUserName & "' and Vfirm = '" & gCCode & "' and AcName = 'Profit And Loss Of Current Year'"
                        
                        gCn.Execute "delete TmpClosingBalance  where UserName='" & gUserName & "' and Vfirm = '" & gCCode & "' and TmpClosingBalance.TopGrp = 90003 "
                    End If
                    .Formulas(5) = "wCPAN = 'PAN No : ' & '" & gCPAN & "'"
                    .Formulas(6) = "wIsDepotFirm = " & gCIsDepotFirm & ""
                    If gSelectedMenu = "Trial Balance Group Simple" Then
                       .SelectionFormula = "{TmpClosingBalance.Vfirm} = '" & Trim(gCCode) & "' and {TmpClosingBalance.username}='" & gUserName & "' and {TmpClosingBalance.ClBal_Dr}+{TmpClosingBalance.ClBal_Cr} > 0 and  {qryAccountGrouping.Agcode} = {TmpClosingBalance.Agcode}"
                       .ReportFileName = gReportPath & "rptTrialBalanceGrpSim.rpt"
                    Else
                       .SelectionFormula = "{TmpClosingBalance.Vfirm} = '" & Trim(gCCode) & "' and {TmpClosingBalance.username}='" & gUserName & "'  and {TmpClosingBalance.OpBal_Dr}+{TmpClosingBalance.OpBal_Cr}+{TmpClosingBalance.Debit}+{TmpClosingBalance.Credit}+{TmpClosingBalance.ClBal_Dr}+{TmpClosingBalance.ClBal_Cr} > 0    and  {qryAccountGrouping.Agcode} = {TmpClosingBalance.Agcode} "
                       .ReportFileName = gReportPath & "rptTrialBalanceGrpDetail.rpt"
                    End If
                    
'                    .ReportFileName = gReportPath & "rptTrialBalanceGrpSim-Mum.rpt"
'                Case "Trial Balance Group Detail" '-Group Simple
' '                    gCn.Execute "delete from TmpClosingBalance where ClBal_Dr + ClBal_Cr = 0 and Vfirm = '" & Trim(gCCode) & "' and UserName='" & gUserName & "'"
'                    .SelectionFormula = "{TmpClosingBalance.Vfirm} = '" & Trim(gCCode) & "' and {TmpClosingBalance.username}='" & gUserName & "' and  {qryAccountGrouping.Agcode} = {TmpClosingBalance.Agcode} "
'                    .Formulas(5) = "wCPAN = 'PAN No : ' & '" & gCPAN & "'"
'                    .Formulas(6) = "wIsDepotFirm = " & gCIsDepotFirm & ""
'                    .ReportFileName = gReportPath & "rptTrialBalanceGrpDetail.rpt"
                Case "Debitor/Creditor List Beyound Limit"
                    .Formulas(5) = "BalLimit = '" & CDbl(Val(TxtAmount)) & "'"
                    x1 = "{qryTrbalDetail.username}='" & gUserName & "' and {qryTrbalDetail.ACGRPTYPE}='A' "
                   .Formulas(6) = "Wpanno  = 'Pan No : ' & '" & gCPAN & "'"
                    If wAgCode > 0 Then '-- Group
                        x1 = x1 + " and {qryTrbalDetail.AgCode}=" & wAgCode
                    End If
                    x1 = x1 + " and ({qryTrbalDetail.clBal_Dr}+ {qryTrbalDetail.clBal_Cr}) >= " & Val(TxtAmount)
                    .SelectionFormula = x1
                    
                    If OptListconf(0).Value = True Then
                       .ReportFileName = gReportPath & "rptDrCrLimitList.rpt"
                    Else
                       .ReportFileName = gReportPath & "rptDrCrLimitConfLtr.rpt"
                    End If
           End Select
        Case "Loan Reporting"
             gCn.Execute "delete from tmpselection where UserName='" & gUserName & "'"
             If gBackEndDB = gBackEndAccess Then
                gCn.Execute "insert into tmpselection (dt1,Dt2,UserName) values ('" & dtpFrDt & "','" & dtpToDt & "','" & gUserName & "')"
             Else
                gCn.Execute "insert into tmpselection (dt1,dt2,UserName) values (to_date('" & dtpFrDt & "','dd/MM/yyyy'),to_Date('" & dtpToDt & "','dd/MM/yyyy'),'" & gUserName & "') "
             End If
             
             gCn.Execute "Delete from tmpGenTbl where USERNAME='" & gUserName & "'"
             Dim RsP As Recordset
             Dim Party As String
             Dim Bal As Double
             Dim Bal2 As Double
             Dim BalDt As Date
             x1 = "Select Acname,Vdt,Vtype,Vno,AcOname,Debit,Credit,Username,Accity,AcPan from qryAccountLedger where Agname = '" & txtTemp(0) & "' and UserName='" & gUserName & "' "
             
             gCn.Execute "insert into tmpGenTbl (Nar1,Vdt,Vtype,Vno,Nar2,Amt1,Amt2,UserName,Nar3,Nar4) Select Acname,Vdt,Vtype,Vno,AcOname,Debit,Credit,Username,Accity,AcPan from qryAccountLedger where Agname = '" & txtTemp(0) & "' and UserName='" & gUserName & "' order by Acname,Vdt,Vtype,Vno "
             
             If (GProcRstOpen(RsP, "select * from tmpGenTbl where UserName='" & gUserName & "' order by Nar1,Vdt,Vtype,Vno", "O") > 0) Then
                 RsP.MoveFirst
                 Do While Not RsP.EOF
                    Party = RsP.Fields("Nar1")
                    Bal = 0
                    Bal2 = 0
                    BalDt = RsP.Fields("Vdt")
                    Do While Not RsP.EOF
                       If RsP.Fields("Nar1") <> Party Then Exit Do
                       Bal = Bal + (RsP.Fields("Amt2") - RsP.Fields("Amt1"))
                       If RsP.Fields("Vtype") <> "JV" Then
                          Bal2 = Bal2 + (RsP.Fields("Amt2") - RsP.Fields("Amt1"))
                       End If
                       RsP.Fields("Amt3") = Bal    ' Ladger Bal
                       RsP.Fields("Amt4") = Bal2   ' Bal without JV
                       RsP.Update
                       RsP.MoveNext
'                       If Not RsP.EOF And RsP.Fields("Nar1") = Party Then
'                          If RsP.Fields("Vdt") <> BalDt Then
'                             RsP.MovePrevious
'                             RsP.Fields("Amt3") = Bal    ' Ladger Bal
'                             RsP.Fields("Amt4") = Bal2   ' Bal without JV
'                             RsP.Update
'                             RsP.MoveNext
'                          End If
'                       End If
'                       If RsP.EOF Or RsP.Fields("Nar1") <> Party Then
'                          If RsP.Fields("Vdt") <> BalDt Then
'                             RsP.Fields("Amt3") = Bal    ' Ladger Bal
'                             RsP.Fields("Amt4") = Bal2   ' Bal without JV
'                             RsP.Update
'                          End If
'                       End If
                    Loop
                 Loop
             End If
             RsP.Close
             
             '--Set Loacation
             If gBackEndDB = gBackEndAccess Then
                .DataFiles(0) = gDbLocation
                 gReportPath = App.Path & "\HiReports_Access\"
             Else '-- Oracle
                 .Connect = "DSN=" + gODBCDSN + " ;UID=Scott;PWD=" + gODBCPwd + ";DBQ=" + gODBCHostStr
                  gReportPath = App.Path & "\HiReports_ORA\"
             End If
             .Formulas(4) = "wFrToDt = 'From ' & '" & dtpFrDt & "' & '  To ' & '" & dtpToDt & "'"
             .Formulas(5) = "PANNo= '" & gCPAN & "'"
             .Formulas(6) = "PropNm= '" & gCCapitalName & "'"
             '-- Selection Formula
             .SelectionFormula = "{tmpGentbl.UserName}='" & gUserName & "'"
             
             '-- Report Name
             .ReportFileName = gReportPath & "rptLoanReporting.rpt"
        Case "T-Format":
            gCn.BeginTrans
            gCn.Execute "delete from tmpSelection where UserName='" & gUserName & "'"
             If gBackEndDB = gBackEndAccess Then
                gCn.Execute "insert into tmpSelection (dt1,dt2,UserName) values ('" & dtpFrDt.Value & "','" & dtpToDt.Value & "','" & gUserName & "')"
            Else
                gCn.Execute "insert into tmpSelection (dt1,dt2,UserName) values (to_date('" & dtpFrDt.Value & "','dd/MM/yyyy'),to_date('" & dtpToDt.Value & "','dd/MM/yyyy'),'" & gUserName & "') "
            End If
           .Formulas(4) = "wFrToDt='From : '& '" & dtpFrDt.Value & "' & ' To ' & '" & dtpToDt.Value & "'"
            gCn.CommitTrans
            CryTFormat
            .ReportFileName = gReportPath & "rptTFormat.rpt"
            x1 = "{tmpgenTbl.Username}='" & gUserName & "'"
            .SelectionFormula = x1
'            CryTFormat
'            .ReportFileName = gReportPath & "rptTFormat.rpt"
'        Case "Monthly Expences"
'            .ReportFileName = gReportPath & "rptMonExp.rpt"
    End Select
    .Action = 1
    .PageZoom (120)
End With
'==================================== Dos Report
'ElseIf optReportOn(0) = True Then
'     If Index = 1 Then mScreen = False
'     Select Case gSelectedMenu
'        Case "Balance Sheet":
'          '  cryBalSht ("G")
'        Case "Trial Balance Simple":
'          '  Trial_Bal_SA "D"
'        Case "Trial Balance Detail":
'          '  Trial_Bal_DA Format(dtpToDt, "dd/mm/yyyy")
'        Case "Trial Balance Group Simple":
'          '  Trial_Bal_SG
'        Case "Trial Balance Group Detail":
'          '  Trial_Bal_DG
'        Case "Profit/Loss Account":
'           'Prof_Loss_Ac "D"
'        Case "Schedule Printing":
''            If Trim(txtTemp(0).Text) = "" Then
''                MsgBox "Group Selection is Compulsary", vbCritical + vbOKOnly, "Select Group"
''                txtTemp(0).SetFocus
''                Exit Sub
''            End If
''            Trial_Bal_SA "D"
'        Case "All Schedules"
'          '  All_Schedule
'    End Select
'    DoEvents
'    If (Len(gSelectedMenu) > 0) Then
'        If mScreen = False Then
'           RetVal = Shell(App.Path & "\txtprint.exe", vbMaximizedFocus)
'        ElseIf mScreen = True Then
'           RetVal = Shell(App.Path & "\WordPad.exe c:\rpt1.txt", vbMaximizedFocus)
'        End If
'    Else
'       MsgBox "No Report", vbCritical + vbOKOnly, "No Report"
'    End If
'    Unload Me
'End If
Screen.MousePointer = vbDefault
End Sub

Private Sub cmdClosing_Click()
Dim RstAcc As Recordset, i As Integer

If GProcRstOpen(RstAcc, "Select AcCode,Acname from tblmastaccount where agcode = 90018", "R", gCn) > 0 Then
   GProcRstOpen RstClosingStock, "Select * from tblclosingstock where VFirm = '" & gCCode & "' and  format(Stdate,'yyyy/MM/dd') = '" & Format(dtpToDt.Value, "yyyy/MM/dd") & "'  order by staccode", "O", gCn
   Frame1.Visible = True
   With RstAcc
        .MoveFirst
        i = 0
        Do While Not .EOF
           If RstClosingStock.RecordCount > 0 Then
                RstClosingStock.MoveFirst
                RstClosingStock.Find "Staccode = " & !AcCode
           End If
           If i > Maxi Then
              Load lblStAccount(i)
              Load txtStAmount(i)
              Load txtstAccode(i)
              lblStAccount(i).Top = 720 + (i * 700)
              txtStAmount(i).Top = 720 + (i * 700)
              Maxi = i
           End If
           lblStAccount(i) = !AcName
           txtstAccode(i) = !AcCode
            lblStAccount(i).Visible = True
            txtStAmount(i).Visible = True
            txtStAmount(i).Enabled = True
           If RstClosingStock.EOF = False Then
              txtStAmount(i) = Format(RstClosingStock!StAmount, "###0.00")
           Else
              txtStAmount(i) = Format(0, "###0.00")
           End If
           .MoveNext
           i = i + 1
        Loop
   End With
   txtStAmount(0).SetFocus
Else
   RstAcc.Close
   MsgBox "Stock Accounts are not Available..."
End If
End Sub
Private Sub cmdOK_Click()
   Dim i As Integer
   gCn.BeginTrans
   For i = 0 To Maxi
'      If Val(txtStAmount(i)) > 0 Then
         With RstClosingStock
         '   .MoveLast
            If .RecordCount > 0 Then
               .MoveFirst
               .Find "Staccode = " & CLng(txtstAccode(i))
            End If
            If .EOF = True Then .AddNew
            !StAcCode = CLng(txtstAccode(i))
            !stdate = dtpToDt
            !StAmount = Val(txtStAmount(i))
            !VFirm = gCCode
            .Update
         End With
 '     End If
   Next i
   gCn.CommitTrans
   RstClosingStock.Close
   Frame1.Visible = False
End Sub



'Private Function FillDataPLA(m_Rec As Recordset, L_R As String, mCno As Double, m_Tp As String, m_Name As String, m_Clbal As Double)
''gCn.BeginTrans
''m_Rec.Requery
'With m_Rec
'    If .RecordCount < mCno Then
'       '.AddNew
'       '!Vtype = m_Tp
'       '!Vno = mCno
'       '!UserName = gUserName
'       If (L_R = "L") Then
'            '!Nar1 = m_Name
'            '!Amt1 = m_Clbal
'            gCn.Execute "insert into tmpGentbl (Vtype,Vno,Username,Nar1,Amt1) values ('" & m_Tp & "'," & mCno & ",'" & gUserName & "','" & m_Name & "'," & m_Clbal & ")"
'       Else
'            '!Nar2 = m_Name
'            '!AMT2 = m_Clbal
'            gCn.Execute "insert into tmpGentbl (Vtype,Vno,Username,Nar2,Amt2) values ('" & m_Tp & "'," & mCno & ",'" & gUserName & "','" & m_Name & "'," & m_Clbal & ")"
'       End If
'      .Requery
'    Else
'      ' .MoveFirst
'      ' .Find "Vno = " & mCno
'      If (L_R = "L") Then
'            gCn.Execute "update  tmpGentbl Set Nar1='" & m_Name & "',Amt1=" & m_Clbal & "  where vno = " & mCno & " and username='" & gUserName & "'"
'      Else
'            gCn.Execute "update  tmpGentbl Set Nar2='" & m_Name & "',Amt2=" & m_Clbal & "  where vno = " & mCno & " and username='" & gUserName & "'"
'      End If
'    End If
'End With
'DoEvents
''gCn.CommitTrans
'End Function
'Private Sub Prof_Loss_Ac(RptTp As String)
'Dim rst1(2) As Recordset, mDocL As Double, MDocR As Double
'Dim mAgCode(5) As Long, mbkMk(5) As Variant
'Dim i As Integer, j As Integer
'Dim mPc As String
'Dim t1 As Double, t2 As Double, wGp As Double, mFirstTime As Boolean, mGpName As String
'Dim wTp As String, X As String, X1 As String, wttp(7) As String, wprline As String
'tblVoucher_TO_tmptblFin dtpToDt
'gCn.Execute "Delete from tmpgentbl where Username='" & gUserName & "' "
'
'GProcRstOpen rsttmpGenTbl, "Select * from tmpgentbl where Username='" & gUserName & "'", "O", gCn 'order by docno
'X = "Select Accode,Agcode,ANAME, CLBAL, CLCD,OpBal,Opcd from tmptblFin where Username='" & gUserName & "' and AgCode = 90018 order by Aname"
'GProcRstOpen rst1(1), X, "R", gCn
'X = "Select Accode,Agcode,ANAME, CLBAL, CLCD from tmptblFin where Username='" & gUserName & "' and TopGroup = 90003 order by AgCode,accode"
'If GProcRstOpen(rst1(0), X, "R", gCn) > 0 Then
'mDocL = 0
'MDocR = 0
'' M - Manufacturung & Trading A/c
'wGp = 0
'    If rst1(1).RecordCount > 0 Then
'        With rst1(1)
'            .MoveFirst
'            mFirstTime = True
'            Do While Not .EOF
'               If (!OpBal > 0) Then
'                  If mFirstTime = True Then
'                     mDocL = mDocL + 1
'                     FillDataPLA rsttmpGenTbl, "L", mDocL, "M", "Opening Stock", 0
'                     mFirstTime = False
'                  End If
'                  mDocL = mDocL + 1
'                  FillDataPLA rsttmpGenTbl, "L", mDocL, "M", "   " + !Aname, !OpBal
'                  wGp = wGp - !OpBal
'               End If
'               .MoveNext
'            Loop
'        End With
'    End If
'For i = 1 To 2
'    If (i = 1) Then
'       wTp = "M"
'       mAgCode(0) = 90021
'    Else
'       If (mDocL > MDocR) Then
'          MDocR = mDocL
'       Else
'          mDocL = MDocR
'       End If
'       wTp = "P"
'       mAgCode(0) = 90022
'        If wGp <> 0 Then
'            If (wGp < 0) Then
'               mDocL = mDocL + 1
'               FillDataPLA rsttmpGenTbl, "L", mDocL, "P", "Gross Loss", Abs(wGp)
'            Else
'               MDocR = MDocR + 1
'               FillDataPLA rsttmpGenTbl, "R", MDocR, "P", "Gross Profit", wGp
'            End If
'        End If
'    End If
'    With rst1(0)
'        .MoveFirst
'        .Find "AgCode = " & mAgCode(0)
'        Do While Not .EOF
'            If !AgCode <> mAgCode(0) Then Exit Do
'            mAgCode(1) = !AcCode
'            mbkMk(1) = .Bookmark
'            mFirstTime = False
'            If (!CLBAL > 0) Then
'               If (!CLCD = "D") Then
'                  mDocL = mDocL + 1
'                  FillDataPLA rsttmpGenTbl, "L", mDocL, wTp, !Aname, !CLBAL
'                  wGp = wGp - !CLBAL
'               Else
'                  MDocR = MDocR + 1
'                  FillDataPLA rsttmpGenTbl, "R", MDocR, wTp, !Aname, !CLBAL
'                  wGp = wGp + !CLBAL
'               End If
'            Else
'               mGpName = !Aname
'               mFirstTime = True
'            End If
'            .Find "AgCode = " & mAgCode(1)
'            Do While Not .EOF
'                If !AgCode <> mAgCode(1) Then Exit Do
'                mbkMk(2) = .Bookmark
'                mAgCode(2) = !AcCode
'                If (!CLBAL > 0) Then
'                   If (!CLCD = "D") Then
'                      If mFirstTime = True Then
'                         mDocL = mDocL + 1
'                         FillDataPLA rsttmpGenTbl, "L", mDocL, wTp, mGpName, 0
'                         mFirstTime = False
'                      End If
'                      mDocL = mDocL + 1
'                      FillDataPLA rsttmpGenTbl, "L", mDocL, wTp, "   " + !Aname, !CLBAL
'                      wGp = wGp - !CLBAL
'                   Else
'                      If mFirstTime = True Then
'                         MDocR = MDocR + 1
'                         FillDataPLA rsttmpGenTbl, "R", MDocR, wTp, mGpName, 0
'                         mFirstTime = False
'                      End If
'                      MDocR = MDocR + 1
'                      FillDataPLA rsttmpGenTbl, "R", MDocR, wTp, "   " + !Aname, !CLBAL
'                      wGp = wGp + !CLBAL
'                   End If
'                End If
'                .Find "AgCode = " & mAgCode(2)
'                Do While Not .EOF
'                    If !AgCode <> mAgCode(2) Then Exit Do
'                    mbkMk(3) = .Bookmark
'                    mAgCode(3) = !AcCode
'                    If (!CLBAL > 0) Then
'                       If (!CLCD = "D") Then
'                          mDocL = mDocL + 1
'                          FillDataPLA rsttmpGenTbl, "L", mDocL, wTp, "   " + !Aname, !CLBAL
'                          wGp = wGp - !CLBAL
'                       Else
'                          MDocR = MDocR + 1
'                          FillDataPLA rsttmpGenTbl, "R", MDocR, wTp, "   " + !Aname, !CLBAL
'                          wGp = wGp + !CLBAL
'                       End If
'                    End If
'                    .Find "AgCode = " & mAgCode(3)
'                    Do While Not .EOF
'                        If !AgCode <> mAgCode(3) Then Exit Do
'                        mbkMk(4) = .Bookmark
'                        mAgCode(4) = !AcCode
'                        If (!CLBAL > 0) Then
'                           If (!CLCD = "D") Then
'                              mDocL = mDocL + 1
'                              FillDataPLA rsttmpGenTbl, "L", mDocL, wTp, "   " + !Aname, !CLBAL
'                              wGp = wGp - !CLBAL
'                           Else
'                              MDocR = MDocR + 1
'                              FillDataPLA rsttmpGenTbl, "R", MDocR, wTp, "   " + !Aname, !CLBAL
'                              wGp = wGp + !CLBAL
'                           End If
'                        End If
'                        .Find "AgCode = " & mAgCode(4)
'                        Do While Not .EOF
'                            If !AgCode <> mAgCode(4) Then Exit Do
'                            mbkMk(5) = .Bookmark
'                            mAgCode(5) = !AcCode
'                            If (!CLBAL > 0) Then
'                               If (!CLCD = "D") Then
'                                  mDocL = mDocL + 1
'                                  FillDataPLA rsttmpGenTbl, "L", mDocL, wTp, "   " + !Aname, !CLBAL
'                                  wGp = wGp - !CLBAL
'                               Else
'                                  MDocR = MDocR + 1
'                                  FillDataPLA rsttmpGenTbl, "R", MDocR, wTp, "   " + !Aname, !CLBAL
'                                  wGp = wGp + !CLBAL
'                               End If
'                            End If
'                            .Bookmark = mbkMk(5)
'                            .MoveNext
'                        Loop
'                        .Bookmark = mbkMk(4)
'                        .MoveNext
'                    Loop
'                    .Bookmark = mbkMk(3)
'                    .MoveNext
'                Loop
'                .Bookmark = mbkMk(2)
'                .MoveNext
'            Loop
'            .Bookmark = mbkMk(1)
'            .MoveNext
'        Loop
'    End With
'    If i = 1 Then
'        If rst1(1).RecordCount > 0 Then
'        With rst1(1)
'            .MoveFirst
'            mFirstTime = True
'            Do While Not .EOF
'               If (!CLBAL > 0) Then
'                    If mFirstTime = True Then
'                       MDocR = MDocR + 1
'                       FillDataPLA rsttmpGenTbl, "R", MDocR, "M", "Closing Stock", 0
'                       mFirstTime = False
'                    End If
'                  MDocR = MDocR + 1
'                  FillDataPLA rsttmpGenTbl, "R", MDocR, "M", "   " + !Aname, !CLBAL
'                  wGp = wGp + !CLBAL
'               End If
'               .MoveNext
'            Loop
'        End With
'        End If
'        If wGp <> 0 Then
'            If (wGp > 0) Then
'               mDocL = mDocL + 1
'               FillDataPLA rsttmpGenTbl, "L", mDocL, "M", "Gross Profit", wGp
'            Else
'               MDocR = MDocR + 1
'               FillDataPLA rsttmpGenTbl, "R", MDocR, "M", "Gross Loss", wGp
'            End If
'        End If
'    Else
'        If wGp <> 0 Then
'            If (wGp > 0) Then
'               mDocL = mDocL + 1
'               FillDataPLA rsttmpGenTbl, "L", mDocL, "P", "Net Profit", wGp
'            Else
'               MDocR = MDocR + 1
'               FillDataPLA rsttmpGenTbl, "R", MDocR, "P", "Net Loss", Abs(wGp)
'            End If
'        End If
'    End If
'Next i
'End If
'rst1(0).Close
'rst1(1).Close
'rsttmpGenTbl.Close
'If RptTp = "D" Then
'    If GProcRstOpen(rsttmpGenTbl, "select * from tmpgentbl where username='" & gUserName & "' order by vtype,vno", "R", gCn) > 0 Then
'        wDotLine = String(92, "-")
'        If mScreen = True Then
'           mPc = ""
'        Else
'           mPc = Chr(27) + Chr(33) + Chr(1)
'        End If
'        Open "c:\rpt1.txt" For Output As #1
'        wHdPrn = True
'        LnCnt = 0
'        PgCnt = 1
'        srcnt = 1
'        t1 = 0
'        t2 = 0
'        Print #1, "" & gCName & "                   From " & dtpFrDt & " To " & dtpToDt & "            Page No. " + Str(PgCnt)
'        PgCnt = PgCnt + 1
'        Print #1, "Manufacturing & Trading Account"
'        Print #1, mPc + wDotLine
'        Print #1, "Account Name                           Amount  Account Name                           Amount"
'        Print #1, wDotLine
'        With rsttmpGenTbl
'            t1 = 0
'            t2 = 0
'            .MoveFirst
'            Do While Not .EOF
'               If !Vtype <> "M" Then Exit Do
'               '************** Left Side
'               wprline = ""
'                wprline = Left(!Nar1 & vbNullString + Space(34), 34)
'                If (!Amt1 > 0) Then
'                   wprline = wprline + GProcNoToStr(!Amt1, 12, 2) + "  "
'                   t1 = t1 + !Amt1
'                Else
'                   wprline = wprline + Space(14)
'                End If
'                wprline = wprline + Left(!Nar2 & vbNullString + Space(34), 34)
'                If (!AMT2 > 0) Then
'                   wprline = wprline + GProcNoToStr(!AMT2, 12, 2)
'                   t2 = t2 + !AMT2
'                End If
'               Print #1, wprline
'               .MoveNext
'            Loop
'            Print #1, wDotLine
'            Print #1, Space(33) + GProcNoToStr(t1, 13, 2) + "  " + Space(33) + GProcNoToStr(t2, 13, 2)
'            Print #1, wDotLine + Chr(12)
'            '************************* Profit Loss A/c.
'            Print #1, Chr(18) + "" & gCName & "                   From " & dtpFrDt & " To " & dtpToDt & "            Page No. " + Str(PgCnt)
'            Print #1, "Profit / Loss Account"
'            Print #1, mPc + wDotLine
'            Print #1, "Account Name                           Amount  Account Name                           Amount"
'            Print #1, wDotLine
'            t1 = 0
'            t2 = 0
'            Do While Not .EOF
'               '************** Left Side
'               wprline = ""
'                wprline = Left(!Nar1 & vbNullString + Space(34), 34)
'                If (!Amt1 > 0) Then
'                   wprline = wprline + GProcNoToStr(!Amt1, 12, 2) + "  "
'                   t1 = t1 + !Amt1
'                Else
'                   wprline = wprline + Space(14)
'                End If
'                wprline = wprline + Left(!Nar2 & vbNullString + Space(34), 34)
'                If (!AMT2 > 0) Then
'                   wprline = wprline + GProcNoToStr(!AMT2, 12, 2)
'                   t2 = t2 + !AMT2
'                End If
'               Print #1, wprline
'               .MoveNext
'            Loop
'        End With
'        Print #1, wDotLine
'        Print #1, Space(33) + GProcNoToStr(t1, 13, 2) + "  " + Space(33) + GProcNoToStr(t2, 13, 2)
'        Print #1, wDotLine + Chr(27) + Chr(33) + Chr(0) + Chr(12)
'        Close #1
'    End If
'    rsttmpGenTbl.Close
'End If
'End Sub
'---- Trial Balance Simple
'Private Sub Trial_Bal_SA(RptType As String)
'Dim rst1 As Recordset
'Dim i As Integer, OpDiff As Double
'Dim tdt As Date
'Dim currdt As Date
'Dim t1 As Double, t2 As Double, wSndDrTot As Double, wSndCrTot As Double
'Dim wTp As String, X As String, X1 As String
'wDotLine = String(80, "-")
'tblVoucher_TO_tmptblFin dtpToDt
'If RptType = "C" Then '-Crystal Report
'    Exit Sub
'End If
'OpDiff = 0
'If gSelectedMenu = "Schedule Printing" Then
'   X = "Select ANAME, CLBAL, CLCD, OPBAL, OPCD, ACCODE, AGCODE from tmptblFin where CLBAL > 0 and AGCODE = " & wAgCode & " order by ANAME"
'Else
'   X = "Select ANAME, CLBAL, CLCD, OPBAL, OPCD, ACCODE, AGCODE from tmptblFin where (CLBAL > 0 or OPBAL > 0) order by ANAME"
'End If
'i = GProcRstOpen(rst1, X, "R", gCn)
'If (i = 0) Then
'   gSelectedMenu = ""
'   Exit Sub
'End If
'Open "c:\rpt1.txt" For Output As #1
'wHdPrn = True
'
'wSndDrTot = 0
'wSndCrTot = 0
'LnCnt = 0
'PgCnt = 1
'srcnt = 1
't1 = 0
't2 = 0
'With rst1
'    .MoveFirst
'    Do While .EOF = False
'       If (gSelectedMenu = "Trial Balance 1") And (.Fields(3) > 0) Then
'           If (.Fields(4) = "C") Then
'              OpDiff = OpDiff + .Fields(3)
'           Else
'              OpDiff = OpDiff - .Fields(3)
'           End If
'       End If
'       X1 = ""
'       If (.Fields(5) = 90018) Or (.Fields(6) = 90018) Then
'            If (.Fields(3) > 0) Then
'               X1 = GProcNoToStr(.Fields(3), 15, 2)
'               If (.Fields(4) = "C") Then
'                  X1 = Space(15) + X1
'                  t1 = t1 + .Fields(3)
'               Else
'                  t2 = t2 + .Fields(3)
'               End If
'            End If
'       Else
'            If (Check2.Value = 1) And ((.Fields(6) = 90015) Or (.Fields(6) = 90017)) Then
'               If (.Fields(6) = 90015) Then
'                  wSndCrTot = wSndCrTot + .Fields(1)
'               Else
'                  wSndDrTot = wSndDrTot + .Fields(1)
'               End If
'            Else
'            If (.Fields(1) > 0) Then
'               X1 = GProcNoToStr(.Fields(1), 15, 2)
'               If (.Fields(2) = "C") Then
'                  X1 = Space(15) + X1
'                  t1 = t1 + .Fields(1)
'               Else
'                  t2 = t2 + .Fields(1)
'               End If
'            End If
'            End If
'       End If
'       If (X1 <> "") Then
'          If wHdPrn Then
'             wHdPrn = False
'             If mScreen = True Then
'                Print #1, gCName
'             Else
'                Print #1, Chr(18) + Chr(27) + Chr(14) & gCName & Chr(27) + Chr(18)
'             End If
'             Print #1, "For Period From " & dtpFrDt.Value & "  To  " & dtpToDt.Value & "                  Page No. " + Str(PgCnt)
'             If mScreen = True Then
'                If (gSelectedMenu = "Schedule Printing") Then
'                   Print #1, "Schedule Printing of " + txtTemp(0).Text
'                Else
'                   Print #1, "Trial Balnce List "
'                End If
'             Else
'                If (gSelectedMenu = "Schedule Printing") Then
'                   Print #1, Chr(27) + Chr(69) & "Schedule Printing " & Chr(27) + Chr(70) + txtTemp(0).Text
'                Else
'                   Print #1, Chr(27) + Chr(69) & "Trial Balnce List " & Chr(27) + Chr(70)
'                End If
'             End If
'             Print #1, wDotLine
'             Print #1, "Sr.    Account Name                                   Debit          Credit "
'             Print #1, wDotLine
'             LnCnt = 6
'             PgCnt = PgCnt + 1
'          End If
'          Print #1, Right("    " + Str(srcnt), 4) + "  " + Left(.Fields(0) + Space(40), 40) & "  " & X1
'          LnCnt = LnCnt + 1
'          srcnt = srcnt + 1
'          If (LnCnt > 65) Then
'             If mScreen = True Then
'                Print #1, wDotLine
'             Else
'                Print #1, wDotLine + Chr(12)
'             End If
'             wHdPrn = True
'             LnCnt = 0
'          End If
'       End If
'       .MoveNext
'    Loop
'    .Close
'End With
'If (wSndDrTot <> 0) Then
'    X1 = GProcNoToStr(wSndDrTot, 15, 2)
'    If (wSndDrTot < 0) Then
'       X1 = Space(15) + X1
'       t1 = t1 + Abs(wSndDrTot)
'    Else
'       t2 = t2 + wSndDrTot
'    End If
'    Print #1, Right("    " + Str(srcnt), 4) + "  " + Left("Total of SUNDRY DEBTORS  " + Space(40), 40) & "  " & X1
'    LnCnt = LnCnt + 1
'    srcnt = srcnt + 1
'End If
'If (wSndCrTot <> 0) Then
'    X1 = GProcNoToStr(wSndCrTot, 15, 2)
'    If (wSndCrTot > 0) Then
'       X1 = Space(15) + X1
'       t1 = t1 + wSndDrTot
'    Else
'       t2 = t2 + Abs(wSndDrTot)
'    End If
'    Print #1, Right("    " + Str(srcnt), 4) + "  " + Left("Total of SUNDRY CREDITORS  " + Space(40), 40) & "  " & X1
'    LnCnt = LnCnt + 1
'    srcnt = srcnt + 1
'End If
'If (gSelectedMenu = "Trial Balance 1") And (OpDiff <> 0) Then
'    X1 = GProcNoToStr(OpDiff, 15, 2)
'    If (OpDiff < 0) Then
'       X1 = Space(15) + X1
'       t1 = t1 + Abs(OpDiff)
'    Else
'       t2 = t2 + OpDiff
'    End If
'    Print #1, Right("    " + Str(srcnt), 4) + "  " + Left("Difference in Opening Balance " + Space(40), 40) & "  " & X1
'    srcnt = srcnt + 1
'    LnCnt = LnCnt + 1
'End If
'Print #1, wDotLine
'Print #1, Space(46) + "  " + GProcNoToStr(t1, 15, 2) + " " + GProcNoToStr(t2, 15, 2)
'If mScreen = True Then
'   Print #1, wDotLine
'Else
'   Print #1, wDotLine + Chr(12)
'End If
'Close #1
'End Sub
'--- Trial Balance Detail
'Private Sub Trial_Bal_DA(ToDt As Date)
'Dim rst1 As Recordset, rst2 As Recordset
'Dim i As Integer, OpDiff As Double
'Dim tdt As Date
'Dim currdt As Date, t1d(6) As Double, t1c(6) As Double
'Dim t1 As Double, t2 As Double, t3 As Double, t4 As Double, t5 As Double, t6 As Double
'Dim X As String, X1 As String
'wDotLine = String(130, "-")
''i = GProcRstOpen(rst2, "Select * from tmpgentbl", "O", gCn)
''tblVoucher_TO_tmptblFin dtpToDt
'OpDiff = 0
''If (gSelectedmenu = "Schedule Printing") Then
''   If (Len(wAcCode) = 0) Then Unload Me
''   X = "Select ANAME, ACCODE, AGCODE,opbal,clbal,amt1,amt2,amt3,amt4,amt5,amt6 from tmptblFin,tmpgentbl where tmptblFin.accode = tmpgentbl.pty and AGCODE = " & wAcCode & " order by ANAME"
''Else
''   wAcCode = 0
''   X = "Select ANAME, ACCODE, AGCODE ,amt1,amt2,amt3,amt4,amt5,amt6 from tmptblFin,tmpgentbl where tmptblFin.accode = tmpgentbl.VAccode  order by ANAME"
''End If
''---- insert From dt ,To Dt in tmpGenTbl
' gCn.Execute "delete from tmpGenTbl where UserName='" & gUserName & "'"
'If gBackEndDB = gBackEndAccess Then
'    gCn.Execute "insert into tmpGenTbl (dt1,UserName) values ('" & ToDt & "','" & gUserName & "')"
'Else
'    gCn.Execute "insert into tmpGenTbl (dt1,UserName) values (to_Date('" & ToDt & "','dd/MM/yyyy'),'" & gUserName & "') "
'End If
'i = GProcRstOpen(rst1, "Select * from qryTrbalDetail  order by acname", "R", gCn)
'If (i = 0) Then
'   gSelectedMenu = ""
'   Exit Sub
'End If
'Open "c:\rpt1.txt" For Output As #1
'wHdPrn = True
'LnCnt = 0
'PgCnt = 1
'srcnt = 1
't1 = 0
't2 = 0
'With rst1
'    .MoveFirst
'    Do While .EOF = False
'       X1 = ""
'       If (!AgCode = 90018) Then
'            If (!OpBal > 0) Then
'                X1 = GProcNoToStr(!OpBal, 12, 2) + "  " + Space(12) + "  " + Space(12) + "  " + Space(12) + "  " + GProcNoToStr(!OpBal, 12, 2)
'                t1 = t1 + !OpBal
'                t5 = t5 + !CLBAL
'            End If
'       Else
'            If (Check2.Value = 1) And ((!AgCode = 90015) Or (!AgCode = 90017)) Then
'               If (!AgCode = 90015) Then
'                t1c(1) = t1c(1) + !OpBal_Dr
'                t1c(2) = t1c(2) + !OpBal_Cr
'                t1c(3) = t1c(3) + !Debit
'                t1c(4) = t1c(4) + !Credit
'                t1c(5) = t1c(5) + !ClBal_Dr
'                t1c(6) = t1c(6) + !ClBal_Dr
'               Else
'                t1d(1) = t1d(1) + !OpBal_Dr
'                t1d(2) = t1d(2) + !OpBal_Cr
'                t1d(3) = t1d(3) + !Debit
'                t1d(4) = t1d(4) + !Credit
'                t1d(5) = t1d(5) + !ClBal_Dr
'                t1d(6) = t1d(6) + !ClBal_Cr
'               End If
'            Else
'               If (!OpBal_Dr + !OpBal_Cr + !Debit + !Credit + !ClBal_Dr + !ClBal_Cr > 0) Then
'                X1 = GProcNoToStr(!OpBal_Dr, 12, 2) + "  " + GProcNoToStr(!OpBal_Cr, 12, 2) + "  " + GProcNoToStr(!Debit, 12, 2) + "  " + GProcNoToStr(!Credit, 12, 2) + "  " + GProcNoToStr(!ClBal_Dr, 12, 2) + "  " + GProcNoToStr(!ClBal_Cr, 12, 2)
'                t1 = t1 + !OpBal_Dr
'                t2 = t2 + !OpBal_Cr
'                t3 = t3 + !Debit
'                t4 = t4 + !Credit
'                t5 = t5 + !ClBal_Dr
'                t6 = t6 + !ClBal_Cr
'              End If
'          End If
'       End If
'       If (X1 <> "") Then
'          If wHdPrn Then
'             wHdPrn = False
'             If mScreen = True Then
'                Print #1, gCName
'             Else
'                Print #1, Chr(18) + Chr(27) + Chr(14) & gCName & Chr(27) + Chr(18)
'             End If
'             Print #1, "For Period From " & dtpFrDt.Value & "  To  " & dtpToDt.Value & "                  Page No. " + Str(PgCnt)
'             If (gSelectedMenu = "Schedule Printing") Then
'                Print #1, "Schedule Printing of " + txtTemp(0).Text
'             Else
'                Print #1, "Trial Balnce List "
'             End If
'             If mScreen Then
'                Print #1, wDotLine
'             Else
'                Print #1, Chr(27) + Chr(15) + wDotLine
'             End If
'             Print #1, "Sr.    Account Name                                      Opening Balance               Transaction               Closing Balance"
'             Print #1, "No.                                                    Debit        Credit         Debit        Credit         Debit        Credit"
'             Print #1, wDotLine
'             LnCnt = 5
'             PgCnt = PgCnt + 1
'          End If
'          Print #1, Right("    " + Str(srcnt), 4) + "  " + Left(.Fields(0) + Space(40), 40) & "  " & X1
'          LnCnt = LnCnt + 1
'          srcnt = srcnt + 1
'          If (LnCnt > 65) Then
'             If mScreen = True Then
'                Print #1, wDotLine
'             Else
'                Print #1, wDotLine + Chr(12)
'             End If
'             wHdPrn = True
'             LnCnt = 0
'          End If
'       End If
'       .MoveNext
'    Loop
'    .Close
'End With
'If (t1d(1) + t1d(2) + t1d(3) + t1d(4) > 0) Then
'    X1 = GProcNoToStr(t1d(1), 12, 2) + "  " + GProcNoToStr(t1d(2), 12, 2) + "  " + GProcNoToStr(t1d(3), 12, 2) + "  " + GProcNoToStr(t1d(4), 12, 2) + "  " + GProcNoToStr(t1d(5), 12, 2) + "  " + GProcNoToStr(t1d(6), 12, 2)
'    Print #1, Right("    " + Str(srcnt), 4) + "  " + Left("Total of SUNDRY DEBTORS" + Space(40), 40) & "  " & X1
'    LnCnt = LnCnt + 1
'    srcnt = srcnt + 1
'    t1 = t1 + t1d(1)
'    t2 = t2 + t1d(2)
'    t3 = t3 + t1d(3)
'    t4 = t4 + t1d(4)
'    t5 = t5 + t1d(5)
'    t6 = t6 + t1d(6)
'End If
'If (t1c(1) + t1c(2) + t1c(3) + t1c(4) > 0) Then
'    X1 = GProcNoToStr(t1c(1), 12, 2) + "  " + GProcNoToStr(t1c(2), 12, 2) + "  " + GProcNoToStr(t1c(3), 12, 2) + "  " + GProcNoToStr(t1c(4), 12, 2) + "  " + GProcNoToStr(t1c(5), 12, 2) + "  " + GProcNoToStr(t1c(6), 12, 2)
'    Print #1, Right("    " + Str(srcnt), 4) + "  " + Left("Total of SUNDRY CREDITORS" + Space(40), 40) & "  " & X1
'    LnCnt = LnCnt + 1
'    srcnt = srcnt + 1
'    t1 = t1 + t1c(1)
'    t2 = t2 + t1c(2)
'    t3 = t3 + t1c(3)
'    t4 = t4 + t1c(4)
'    t5 = t5 + t1c(5)
'    t6 = t6 + t1c(6)
'End If
'If (t1 <> t2) Then
'   If (t1 > t2) Then
'      X1 = Space(12) + "  " + GProcNoToStr(t1 - t2, 12, 2) + "  " + Space(12) + "  " + Space(12) + "  " + Space(12) + "  " + GProcNoToStr(t1 - t2, 12, 2)
'      t6 = t6 + t1 - t2
'      t2 = t1
'   Else
'      X1 = GProcNoToStr(t2 - t1, 12, 2) + "  " + Space(12) + "  " + Space(12) + "  " + Space(12) + "  " + GProcNoToStr(t2 - t1, 12, 2) + "  " + Space(12)
'      t5 = t5 + t2 - t1
'      t1 = t2
'   End If
'   Print #1, Right("    " + Str(srcnt), 4) + "  " + Left("Difference in Opening Balance" + Space(40), 40) & "  " & X1
'   LnCnt = LnCnt + 1
'   srcnt = srcnt + 1
'End If
'Print #1, wDotLine
'Print #1, Space(46) + " " + GProcNoToStr(t1, 13, 2) + " " + GProcNoToStr(t2, 13, 2) + " " + GProcNoToStr(t3, 13, 2) + " " + GProcNoToStr(t4, 13, 2) + " " + GProcNoToStr(t5, 13, 2) + " " + GProcNoToStr(t6, 13, 2)
'If mScreen = True Then
'   Print #1, wDotLine
'Else
'   Print #1, wDotLine + Chr(27) + Chr(18) + Chr(12)
'End If
'Close #1
''rst2.Close
''rst1.Close
'End Sub
'Private Sub Trial_Bal_SG()
'Dim rstAccount As Recordset
'Dim Ac1 As Long, Ac2 As Long, Ac3 As Long, Ac4 As Long, Ac5 As Long, Ac6 As Long
'Dim BK1, Bk2, Bk3, Bk4, Bk5, Bk6, Bk7
'Dim t2 As Double, t1 As Double
'Dim X As String, X1 As String, i As Integer
'tblVoucher_TO_tmptblFin dtpToDt
'X = "Select ACCODE, AGCODE, ANAME, CLBAL, CLCD,OPBAL,OPCD From tmptblFin where (opbal > 0 or clbal > 0 or actype = 'G') Order By AGCODE, ANAME"
'i = GProcRstOpen(rstAccount, X, "R", gCn)
'Open "c:\rpt1.txt" For Output As #1
'wHdPrn = True
'LnCnt = 0
'PgCnt = 1
'srcnt = 1
'wDotLine = String(80, "-")
't1 = 0
't2 = 0
'prgBar1.Max = i
'prgBar1.Value = 0
'prgBar1.Visible = True
'With rstAccount
'    .MoveFirst
'    '---------------- GROUP SPACE(5) A/C - ASSET LIABI
'    Do While .EOF = False
'        If IsNull(!AgCode) Then Exit Do
'        Ac1 = !AcCode
'        BK1 = .Bookmark
'        X1 = !Aname
'        Pr_Line (X1)
'        .MoveFirst
'        .Find "AGCODE=" & Ac1
'        '---------------- GROUP ASSET A/C - CURR ASSET
'        Do While .EOF = False
'            If Not !AgCode = Ac1 Then Exit Do
'            Ac2 = !AcCode
'            Bk2 = .Bookmark
'            X1 = " " + Left(!Aname + Space(40), 40) + Space(7)
'            If (.Fields(0) = 90018) Or (.Fields(1) = 90018) Then
'                 If (!OpBal > 0) Then
'                    X1 = X1 + GProcNoToStr(!OpBal, 12, 2)
'                    t1 = t1 + .Fields(3)
'                 End If
'            Else
'                 If (!CLBAL > 0) Then
'                    If (!CLCD = "C") Then
'                       X1 = X1 + Space(14) + GProcNoToStr(!CLBAL, 12, 2)
'                       t2 = t2 + !CLBAL
'                    Else
'                       X1 = X1 + GProcNoToStr(!CLBAL, 12, 2)
'                       t1 = t1 + !CLBAL
'                    End If
'                 End If
'            End If
'            Pr_Line (X1)
'            .MoveFirst
'            .Find "AgCode=" & Ac2
'            '---------------- GROUP Curr ASSET A/C - Bank
'            Do While .EOF = False
'                If Not !AgCode = Ac2 Then Exit Do
'                Ac3 = !AcCode
'                Bk3 = .Bookmark
'                X1 = Space(3) + Left(!Aname + Space(40), 40) + Space(5)
'                If (.Fields(0) = 90018) Or (.Fields(1) = 90018) Then
'                     If (!OpBal > 0) Then
'                        X1 = X1 + GProcNoToStr(!OpBal, 12, 2)
'                        t1 = t1 + .Fields(3)
'                     End If
'                Else
'                     If (!CLBAL > 0) Then
'                        If (!CLCD = "C") Then
'                           X1 = X1 + Space(14) + GProcNoToStr(!CLBAL, 12, 2)
'                           t2 = t2 + !CLBAL
'                        Else
'                           X1 = X1 + GProcNoToStr(!CLBAL, 12, 2)
'                           t1 = t1 + !CLBAL
'                        End If
'                     End If
'                End If
'                Pr_Line (X1)
'                .MoveFirst
'                .Find "AgCode=" & Ac3
'                '---------------- GROUP Bank A/C - The Ich Bank
'                Do While .EOF = False
'                    If Not !AgCode = Ac3 Then Exit Do
'                    Ac4 = !AcCode
'                    Bk4 = .Bookmark
'                    X1 = Space(5) + Left(!Aname + Space(40), 40) + Space(3)
'                    If (.Fields(0) = 90018) Or (.Fields(1) = 90018) Then
'                         If (!OpBal > 0) Then
'                            X1 = X1 + GProcNoToStr(!OpBal, 12, 2)
'                            t1 = t1 + .Fields(3)
'                         End If
'                    Else
'                         If (!CLBAL > 0) Then
'                            If (!CLCD = "C") Then
'                               X1 = X1 + Space(14) + GProcNoToStr(!CLBAL, 12, 2)
'                               t2 = t2 + !CLBAL
'                            Else
'                               X1 = X1 + GProcNoToStr(!CLBAL, 12, 2)
'                               t1 = t1 + !CLBAL
'                            End If
'                         End If
'                    End If
'                    Pr_Line (X1)
'                    .MoveFirst
'                    .Find "AgCode=" & Ac4
'                    '---------------- GROUP The Ich Bank A/c - The ich bank branch - ich
'                    Do While .EOF = False
'                        If Not !AgCode = Ac4 Then Exit Do
'                        Ac5 = !AcCode
'                        Bk5 = .Bookmark
'                        If (!CLBAL > 0) Then
'                           X1 = Space(7) + Left(!Aname + Space(40), 40)
'                            If (!CLCD = "C") Then
'                               X1 = X1 + Space(14) + GProcNoToStr(!CLBAL, 12, 2)
'                               t2 = t2 + !CLBAL
'                            Else
'                               X1 = X1 + GProcNoToStr(!CLBAL, 12, 2)
'                               t1 = t1 + !CLBAL
'                            End If
'                            Pr_Line (X1)
'                        End If
'                        .MoveFirst
'                        .Find "AgCode=" & Ac5
'                        '---------------- GROUP The ich bank branch - ich  A/c - The ich bank (SUB)branch - ich
'                        Do While .EOF = False
'                            If Not !AgCode = Ac5 Then Exit Do
'                            Ac6 = !AcCode
'                            Bk6 = .Bookmark
'                            If (!CLBAL > 0) Then
'                               X1 = Space(9) + Left(!Aname + Space(37), 37)
'                                 If (!CLCD = "C") Then
'                                    X1 = X1 + Space(14) + GProcNoToStr(!CLBAL, 12, 2)
'                                    t2 = t2 + !CLBAL
'                                 Else
'                                    X1 = X1 + GProcNoToStr(!CLBAL, 12, 2)
'                                    t1 = t1 + !CLBAL
'                                 End If
'                                 Pr_Line (X1)
'                            End If
'                            .Bookmark = Bk6
'                            .MoveNext
'                           ' prgBar1.Value = prgBar1.Value + 1
'                        Loop
'                        '''''''''''''''''
'                        .Bookmark = Bk5
'                        .MoveNext
'                       ' prgBar1.Value = prgBar1.Value + 1
'                    Loop
'                    .Bookmark = Bk4
'                    .MoveNext
'                   ' prgBar1.Value = prgBar1.Value + 1
'                Loop
'                rstAccount.Bookmark = Bk3
'                rstAccount.MoveNext
'                'prgBar1.Value = prgBar1.Value + 1
'            Loop
'            rstAccount.Bookmark = Bk2
'            rstAccount.MoveNext
'          '  prgBar1.Value = prgBar1.Value + 1
'        Loop
'        rstAccount.Bookmark = BK1
'        rstAccount.MoveNext
'       ' prgBar1.Value = prgBar1.Value + 1
'    Loop
'End With
'prgBar1.Visible = False
'Print #1, wDotLine
'Print #1, Space(51) + GProcNoToStr(t1, 14, 2) + " " + GProcNoToStr(t2, 14, 2)
'If mScreen = True Then
'   Print #1, wDotLine
'Else
'   Print #1, wDotLine + Chr(12)
'End If
'Close #1
'rstAccount.Close
'End Sub
'Private Sub Trial_Bal_DG()
'Dim rstAccount As Recordset, rst2 As Recordset
'Dim Ac1 As Long, Ac2 As Long, Ac3 As Long, Ac4 As Long, Ac5 As Long, Ac6 As Long
'Dim BK1, Bk2, Bk3, Bk4, Bk5, Bk6, Bk7
'Dim t2 As Double, t1 As Double, t3 As Double, t4 As Double, t5 As Double, t6 As Double
'Dim X As String, X1 As String, i As Integer
'i = GProcRstOpen(rst2, "Select * from tmpgentbl", "O", gCn)
'tblVoucher_TO_tmptblFin dtpToDt
'X = "Select ACCODE, AGCODE, ANAME, OPBAL, CLBAL, AMT1, AMT2, AMT3, AMT4, AMT5, AMT6,ACTYPE from tmptblFin, tmpgentbl where (tmptblFin.ACCODE = tmpgentbl.PTY) order by AGCODE, ANAME"
'i = GProcRstOpen(rstAccount, X, "R", gCn)
'MsgBox i
'Open "c:\rpt1.txt" For Output As #1
'wHdPrn = True
'LnCnt = 0
'PgCnt = 1
'srcnt = 1
'wDotLine = String(145, "-")
't1 = 0
't2 = 0
'prgBar1.Max = i
'prgBar1.Value = 0
'prgBar1.Visible = True
'With rstAccount
'    .MoveFirst
'    '---------------- GROUP SPACE(5) A/C - ASSET LIABI
'    Do While .EOF = False
'        If Not IsNull(!AgCode) Then Exit Do
'        Ac1 = !AcCode
'        BK1 = .Bookmark
'        X1 = !Aname
'        Pr_Line (X1)
'        .MoveFirst
'        .Find "AGCODE=" & Ac1
'        '---------------- GROUP ASSET A/C - CURR ASSET
'        Do While .EOF = False
'            If Not !AgCode = Ac1 Then Exit Do
'            Ac2 = !AcCode
'            Bk2 = .Bookmark
'            X1 = "   " + Left(!Aname + Space(40), 40) + Space(12)
'            If (.Fields(0) = 90018) Or (.Fields(1) = 90018) Then
'                 If (!OpBal > 0) Then
'                    X1 = X1 + GProcNoToStr(!OpBal, 12, 2) + "  " + Space(12) + "  " + Space(12) + "  " + Space(12) + "  " + GProcNoToStr(!OpBal, 12, 2)
'                    t1 = t1 + !OpBal
'                    t5 = t5 + !OpBal
'                 End If
'            Else
'                If (!Amt1 + !AMT2 + !AMT3 + !Amt4 + !AMT5 + !AMT6 > 0) Then
'                      X1 = X1 + GProcNoToStr(!Amt1, 12, 2) + "  " + GProcNoToStr(!AMT2, 12, 2) + "  " + GProcNoToStr(!AMT3, 12, 2) + "  " + GProcNoToStr(!Amt4, 12, 2) + "  " + GProcNoToStr(!AMT5, 12, 2) + "  " + GProcNoToStr(!AMT6, 12, 2)
'                      t1 = t1 + !Amt1
'                      t2 = t2 + !AMT2
'                      t3 = t3 + !AMT3
'                      t4 = t4 + !Amt4
'                      t5 = t5 + !AMT5
'                      t6 = t6 + !AMT6
'                 End If
'            End If
'            If (!Amt1 + !AMT2 + !AMT3 + !Amt4 > 0) Or !actype = "G" Then
'               Pr_Line (X1)
'            End If
'            .MoveFirst
'            .Find "AgCode=" & Ac2
'            '---------------- GROUP Curr ASSET A/C - Bank
'            Do While .EOF = False
'                If Not !AgCode = Ac2 Then Exit Do
'                Ac3 = !AcCode
'                Bk3 = .Bookmark
'                X1 = Space(6) + Left(!Aname + Space(40), 40) + Space(9)
'                If (.Fields(0) = 90018) Or (.Fields(1) = 90018) Then
'                 If (!OpBal > 0) Then
'                    X1 = X1 + GProcNoToStr(!OpBal, 12, 2) + "  " + Space(12) + "  " + Space(12) + "  " + Space(12) + "  " + GProcNoToStr(!OpBal, 12, 2)
'                    t1 = t1 + !OpBal
'                    t5 = t5 + !OpBal
'                     End If
'                Else
'                If (!Amt1 + !AMT2 + !AMT3 + !Amt4 + !AMT5 + !AMT6 > 0) Then
'                      X1 = X1 + GProcNoToStr(!Amt1, 12, 2) + "  " + GProcNoToStr(!AMT2, 12, 2) + "  " + GProcNoToStr(!AMT3, 12, 2) + "  " + GProcNoToStr(!Amt4, 12, 2) + "  " + GProcNoToStr(!AMT5, 12, 2) + "  " + GProcNoToStr(!AMT6, 12, 2)
'                      t1 = t1 + !Amt1
'                      t2 = t2 + !AMT2
'                      t3 = t3 + !AMT3
'                      t4 = t4 + !Amt4
'                      t5 = t5 + !AMT5
'                      t6 = t6 + !AMT6
'                 End If
'                End If
'                If (!Amt1 + !AMT2 + !AMT3 + !Amt4 > 0) Or !actype = "G" Then
'                   Pr_Line (X1)
'                End If
'                .MoveFirst
'                .Find "AgCode=" & Ac3
'                '---------------- GROUP Bank A/C - The Ich Bank
'                Do While .EOF = False
'                    If Not !AgCode = Ac3 Then Exit Do
'                    Ac4 = !AcCode
'                    Bk4 = .Bookmark
'                    X1 = Space(9) + Left(!Aname + Space(40), 40) + Space(6)
'                    If (.Fields(0) = 90018) Or (.Fields(1) = 90018) Then
'                        If (!OpBal > 0) Then
'                           X1 = X1 + GProcNoToStr(!OpBal, 12, 2) + "  " + Space(12) + "  " + Space(12) + "  " + Space(12) + "  " + GProcNoToStr(!OpBal, 12, 2)
'                           t1 = t1 + !OpBal
'                           t5 = t5 + !OpBal
'                         End If
'                    Else
'                        If (!Amt1 + !AMT2 + !AMT3 + !Amt4 + !AMT5 + !AMT6 > 0) Then
'                             X1 = X1 + GProcNoToStr(!Amt1, 12, 2) + "  " + GProcNoToStr(!AMT2, 12, 2) + "  " + GProcNoToStr(!AMT3, 12, 2) + "  " + GProcNoToStr(!Amt4, 12, 2) + "  " + GProcNoToStr(!AMT5, 12, 2) + "  " + GProcNoToStr(!AMT6, 12, 2)
'                             t1 = t1 + !Amt1
'                             t2 = t2 + !AMT2
'                             t3 = t3 + !AMT3
'                             t4 = t4 + !Amt4
'                             t5 = t5 + !AMT5
'                             t6 = t6 + !AMT6
'                        End If
'                    End If
'                    If (!Amt1 + !AMT2 + !AMT3 + !Amt4 > 0) Or !actype = "G" Then
'                       Pr_Line (X1)
'                    End If
'                    .MoveFirst
'                    .Find "AgCode=" & Ac4
'                    '---------------- GROUP The Ich Bank A/c - The ich bank branch - ich
'                    Do While .EOF = False
'                        If Not !AgCode = Ac4 Then Exit Do
'                        Ac5 = !AcCode
'                        Bk5 = .Bookmark
'                        If (!Amt1 + !AMT2 + !AMT3 + !Amt4 + !AMT5 + !AMT6 > 0) Then
'                           X1 = Space(12) + Left(!Aname + Space(40), 40) + Space(3)
'                             X1 = X1 + GProcNoToStr(!Amt1, 12, 2) + "  " + GProcNoToStr(!AMT2, 12, 2) + "  " + GProcNoToStr(!AMT3, 12, 2) + "  " + GProcNoToStr(!Amt4, 12, 2) + "  " + GProcNoToStr(!AMT5, 12, 2) + "  " + GProcNoToStr(!AMT6, 12, 2)
'                             t1 = t1 + !Amt1
'                             t2 = t2 + !AMT2
'                             t3 = t3 + !AMT3
'                             t4 = t4 + !Amt4
'                             t5 = t5 + !AMT5
'                             t6 = t6 + !AMT6
'                             Pr_Line (X1)
'                        End If
'                        .MoveFirst
'                        .Find "AgCode=" & Ac5
'                        '---------------- GROUP The ich bank branch - ich  A/c - The ich bank (SUB)branch - ich
'                        Do While .EOF = False
'                            If Not !AgCode = Ac5 Then Exit Do
'                            Ac6 = !AcCode
'                            Bk6 = .Bookmark
'                            If (!Amt1 + !AMT2 + !AMT3 + !Amt4 + !AMT5 + !AMT6 > 0) Then
'                               X1 = Space(15) + Left(!Aname + Space(40), 40)
'                                  X1 = X1 + GProcNoToStr(!Amt1, 12, 2) + "  " + GProcNoToStr(!AMT2, 12, 2) + "  " + GProcNoToStr(!AMT3, 12, 2) + "  " + GProcNoToStr(!Amt4, 12, 2) + "  " + GProcNoToStr(!AMT5, 12, 2) + "  " + GProcNoToStr(!AMT6, 12, 2)
'                                  t1 = t1 + !Amt1
'                                  t2 = t2 + !AMT2
'                                  t3 = t3 + !AMT3
'                                  t4 = t4 + !Amt4
'                                  t5 = t5 + !AMT5
'                                  t6 = t6 + !AMT6
'                                  Pr_Line (X1)
'                            End If
'                            .Bookmark = Bk6
'                            .MoveNext
'                            prgBar1.Value = prgBar1.Value + 1
'                        Loop
'                        '''''''''''''''''
'                        .Bookmark = Bk5
'                        .MoveNext
'                        prgBar1.Value = prgBar1.Value + 1
'                    Loop
'                    .Bookmark = Bk4
'                    .MoveNext
'                    prgBar1.Value = prgBar1.Value + 1
'                Loop
'                rstAccount.Bookmark = Bk3
'                rstAccount.MoveNext
'                prgBar1.Value = prgBar1.Value + 1
'            Loop
'            rstAccount.Bookmark = Bk2
'            rstAccount.MoveNext
'            prgBar1.Value = prgBar1.Value + 1
'        Loop
'        rstAccount.Bookmark = BK1
'        rstAccount.MoveNext
'        prgBar1.Value = prgBar1.Value + 1
'    Loop
'End With
'prgBar1.Visible = False
'If (t1 <> t2) Then
'   If (t1 > t2) Then
'      X1 = Space(12) + "  " + GProcNoToStr(t1 - t2, 12, 2) + "  " + Space(12) + "  " + Space(12) + "  " + Space(12) + "  " + GProcNoToStr(t1 - t2, 12, 2)
'      t6 = t6 + t1 - t2
'      t2 = t1
'   Else
'      X1 = GProcNoToStr(t2 - t1, 12, 2) + "  " + Space(12) + "  " + Space(12) + "  " + Space(12) + "  " + GProcNoToStr(t2 - t1, 12, 2) + "  " + Space(12)
'      t5 = t5 + t2 - t1
'      t1 = t2
'   End If
'   Print #1, Right("    " + Str(srcnt), 4) + "  " + Left("Difference in Opening Balance" + Space(53), 53) & "  " & X1
'   LnCnt = LnCnt + 1
'   srcnt = srcnt + 1
'End If
'Print #1, wDotLine
'Print #1, Space(59) + " " + GProcNoToStr(t1, 13, 2) + " " + GProcNoToStr(t2, 13, 2) + " " + GProcNoToStr(t3, 13, 2) + " " + GProcNoToStr(t4, 13, 2) + " " + GProcNoToStr(t5, 13, 2) + " " + GProcNoToStr(t6, 13, 2)
'If mScreen = True Then
'   Print #1, wDotLine
'Else
'   Print #1, wDotLine + Chr(27) + Chr(18) + Chr(12)
'End If
'Close #1
'rstAccount.Close
'End Sub
'Private Sub Pr_Line(x_x As String)
'          If wHdPrn Then
'             wHdPrn = False
'             If mScreen = True Then
'                Print #1, gCName
'             Else
'                Print #1, Chr(18) + Chr(27) + Chr(14) & gCName & Chr(27) + Chr(18)
'             End If
'             Print #1, "For Period From " & dtpFrDt.Value & "  To  " & dtpToDt.Value & "                  Page No. " + Str(PgCnt)
'             Print #1, "Trial Balnce List "
'             If (gSelectedMenu = "Trial Balnce 3") Then
'                Print #1, wDotLine
'                Print #1, "Sr.    Account Name                                         Debit        Credit "
'             Else
'                If mScreen = True Then
'                   Print #1, wDotLine
'                Else
'                   Print #1, Chr(27) + Chr(15) + wDotLine
'                End If
'                Print #1, "Sr.    Account Name                                                   Opening Balance               Transaction               Closing Balance"
'                Print #1, "No.                                                                 Debit        Credit         Debit        Credit         Debit        Credit"
'             End If
'             Print #1, wDotLine
'             LnCnt = 6
'             PgCnt = PgCnt + 1
'          End If
'          Print #1, Right("    " + Str(srcnt), 4) + "  " + x_x
'          LnCnt = LnCnt + 1
'          srcnt = srcnt + 1
'          If (LnCnt > 65) Then
'            If mScreen = True Then
'               Print #1, wDotLine
'            Else
'               Print #1, wDotLine + Chr(12)
'            End If
'             wHdPrn = True
'             LnCnt = 0
'          End If
'End Sub
'Public Sub FlexPL()
'Screen.MousePointer = vbHourglass
'frmRptFlexFin.msfPandL.Clear
'frmRptFlexFin.msfPandL.Row = 1
'frmRptFlexFin.msfPandL.Rows = 2
'frmRptFlexFin.lblHead.Caption = "From  " & dtpFrDt.Value & "  To  " & dtpToDt.Value
'frmRptFlexFin.msfPandL.FormatString = "<Account Name                                         |>Amount                             |Account Name                                         |>Amount                             "
''----Proceduure to Write Data in FlexGrid
'Dim Rst1(7) As Recordset, rsttmpgentbl As Recordset
'Dim i As Integer, j As Integer
'Dim tdt As Date
'Dim wRecFind(7) As Boolean
'Dim currdt As Date
'Dim t1 As Double, t2 As Double, wgp As Double, wgp1 As Double, wnp As Double
'Dim wTp As String, x As String, x1 As String, wttp(7) As String, wprline As String
'wDotLine = String(90, "-")
'If Check1.Value = 1 Then
'   i = RunClosBal1(gCn, dtpFrDt.Value, dtpToDt.Value, PgrBar1, "S", Rst1(1), Space(5))
'End If
'i = GProcRstOpen(rsttmpgentbl, "Select * from tmpgentbl Order by DOCTP", "O", gCn)
'x = "Select ANAME, CLBAL, CLCD from tmptblFin where CLBAL > 0 and (AGCODE = 'SALE ' or ACCODE = 'SALE ') order by ANAME"
'i = GProcRstOpen(Rst1(0), x, "R", gCn)
'x = "Select ANAME, CLBAL, CLCD from tmptblFin where CLBAL > 0 and (AGCODE = 'STOCK' or ACCODE = 'STOCK') order by ANAME"
'i = GProcRstOpen(Rst1(1), x, "R", gCn)
'x = "Select ANAME, OPBAL, OPCD from tmptblFin where OPBAL > 0 and (AGCODE = 'STOCK' or ACCODE = 'STOCK') order by ANAME"
'i = GProcRstOpen(Rst1(2), x, "R", gCn)
'x = "Select ANAME, CLBAL, CLCD from tmptblFin where CLBAL > 0 and (AGCODE = 'PURCH' or ACCODE = 'PURCH') order by ANAME"
'i = GProcRstOpen(Rst1(3), x, "R", gCn)
'x = "Select ANAME, CLBAL, CLCD from tmptblFin where CLBAL > 0 and CLCD = 'D' and (AGCODE = 'EXPND') order by ANAME"
'i = GProcRstOpen(Rst1(4), x, "R", gCn)
'x = "Select ANAME, CLBAL, CLCD from tmptblFin where CLBAL > 0 and CLCD = 'D' and (AGCODE = 'EXPNI') order by ANAME"
'i = GProcRstOpen(Rst1(5), x, "R", gCn)
'x = "Select ANAME, CLBAL, CLCD from tmptblFin where CLBAL > 0 and ((CLCD = 'C' and (AGCODE = 'OTHIN' or ACCODE = 'OTHIN')) or (CLCD = 'C' and (AGCODE = 'EXPNI' or ACCODE = 'EXPND'))) order by ANAME"
'i = GProcRstOpen(Rst1(6), x, "R", gCn)
'wttp(0) = ""
'wttp(1) = "Closing "
'wttp(2) = "Opening "
'wttp(3) = ""
'wttp(4) = ""
'wttp(5) = ""
'For i = 0 To 6
'   With Rst1(i)
'    If (.RecordCount > 0) Then
'       wRecFind(i) = True
'       .MoveFirst
'       Do While Not .EOF
'          Select Case i
'            Case 0, 1:
'               wgp = wgp + .Fields(1)
'            Case 2, 3, 4:
'               wgp = wgp - .Fields(1)
'            Case 5:
'               wnp = wnp - .Fields(1)
'            Case 6:
'               wnp = wnp + .Fields(1)
'          End Select
'          .MoveNext
'       Loop
'    End If
'   End With
'Next i
'wgp1 = wgp
'wnp = wnp + wgp
't1 = 0
't2 = 0
'For i = 0 To 6
'    If wRecFind(i) Then
'       Rst1(i).MoveFirst
'    End If
'Next i
'
'Do While True
'   '************** Left Side
'   wprline = ""
'   For j = 2 To 4
'      If wRecFind(j) Then
'         With Rst1(j)
'            t1 = t1 + .Fields(1)
'            wprline = Left(wttp(j) + .Fields(0) + Space(30), 30) + GProcNoToStr(.Fields(1), 15, 2) + "  "
'            .MoveNext
'            If .EOF Then wRecFind(j) = False
'            j = 10
'         End With
'      End If
'   Next j
'   If (j < 10) Then
'       If (wgp > 0) Then
'          t1 = t1 + wgp
'          wprline = Left("Gross Profit" + Space(30), 30) + GProcNoToStr(wgp, 15, 2) + "  "
'          wgp = 0
'       Else
'          wprline = Space(45)
'       End If
'   End If
'   '************** Right Side
'   For j = 0 To 1
'      If wRecFind(j) Then
'         With Rst1(j)
'         t2 = t2 + .Fields(1)
'         wprline = wprline + Left(wttp(j) + .Fields(0) + Space(30), 30) + GProcNoToStr(.Fields(1), 15, 2) + "  "
'         .MoveNext
'         If .EOF Then wRecFind(j) = False
'         j = 10
'         End With
'      End If
'    Next j
'    If (j < 10) Then
'         If (wgp < 0) Then
'            t2 = t2 + Abs(wgp)
'            wprline = wprline + Left("Gross Loss" + Space(30), 30) + GProcNoToStr(Abs(wgp), 15, 2) + "  "
'            wgp = 0
'         Else
'            wprline = wprline + Space(45)
'         End If
'    End If
'    If (Len(Trim(wprline)) = 0) Then Exit Do
'    '----To Write Data in FLexGrid
'    With frmRptFlexFin.msfPandL
'        If Trim(Left(wprline, 30)) = "Gross Profit" Then
'            .Col = 0
'                .CellForeColor = vbHighlight
'            .Col = 1
'                .CellForeColor = vbHighlight
'        End If
'        .TextMatrix(.Row, 0) = Left(wprline, 30)
'        .TextMatrix(.Row, 1) = Format(Val(Mid(wprline, 31, 15)), "#############0.00")
'        If Trim(Mid(wprline, 48, 30)) = "Gross Loss" Then
'            .Col = 0
'                .CellForeColor = vbHighlight
'            .Col = 1
'                .CellForeColor = vbHighlight
'        End If
'        .TextMatrix(.Row, 2) = Mid(wprline, 48, 30)
'        .TextMatrix(.Row, 3) = Format(Val(Mid(wprline, 76, 15)), "############0.00")
'        .Rows = .Rows + 1
'        .Row = .Row + 1
'    End With
'Loop
'With frmRptFlexFin.msfPandL
'     .Col = 1
'        .CellFontBold = True
'        .CellForeColor = vbHighlight
'        .Text = Format(t1, "############0.00")
'     .Col = 3
'        .CellFontBold = True
'        .CellForeColor = vbHighlight
'        .Text = Format(t2, "############0.00")
'        .Rows = .Rows + 1
'        .Row = .Row + 1
'End With
''************************* Profit Loss A/c.
't1 = 0
't2 = 0
'Do While True
'   '************** Left Side
'   wprline = ""
'   If (wgp1 < 0) Then
'      t1 = t1 + Abs(wgp1)
'      wprline = Left("Gross Loss" + Space(30), 30) + GProcNoToStr(Abs(wgp1), 15, 2) + "  "
'      wgp1 = 0
'   Else
'      If wRecFind(5) Then
'         With Rst1(5)
'            t1 = t1 + .Fields(1)
'            wprline = Left(wttp(5) + .Fields(0) + Space(30), 30) + GProcNoToStr(.Fields(1), 15, 2) + "  "
'            .MoveNext
'            If .EOF Then wRecFind(5) = False
'         End With
'      Else
'       If (wnp > 0) Then
'          t1 = t1 + wnp
'          wprline = Left("Net Profit" + Space(30), 30) + GProcNoToStr(wnp, 15, 2) + "  "
'          wnp = 0
'       Else
'          wprline = Space(45)
'       End If
'    End If
'   End If
'   '************** Right Side
'   If (wgp1 > 0) Then
'          t2 = t2 + wgp1
'          wprline = wprline + Left("Gross Profit" + Space(30), 30) + GProcNoToStr(wgp1, 15, 2) + "  "
'          wgp1 = 0
'   Else
'      If wRecFind(6) Then
'         With Rst1(6)
'         t2 = t2 + .Fields(1)
'         wprline = wprline + Left(wttp(6) + .Fields(0) + Space(30), 30) + GProcNoToStr(.Fields(1), 15, 2) + "  "
'         .MoveNext
'         If .EOF Then wRecFind(6) = False
'         End With
'      Else
'         If (wnp < 0) Then
'            t2 = t2 + Abs(wnp)
'            wprline = wprline + Left("Net Loss" + Space(30), 30) + GProcNoToStr(Abs(wnp), 15, 2) + "  "
'            wnp = 0
'         Else
'            wprline = wprline + Space(45)
'         End If
'      End If
'    End If
'    If (Len(Trim(wprline)) = 0) Then Exit Do
'    With frmRptFlexFin.msfPandL
'        .TextMatrix(.Row, 0) = Left(wprline, 30)
'        .TextMatrix(.Row, 1) = Format(Val(Mid(wprline, 31, 15)), "############0.00")
'        .TextMatrix(.Row, 2) = Mid(wprline, 48, 30)
'        .TextMatrix(.Row, 3) = Format(Val(Mid(wprline, 76, 15)), "############0.00")
'    End With
'Loop
'With frmRptFlexFin.msfPandL
'    .Rows = .Rows + 1
'    .Row = .Row + 1
'    .Col = 1
'        .CellFontBold = True
'        .CellForeColor = vbHighlight
'        .Text = Format(t1, "############0.00")
'    .Col = 3
'        .CellFontBold = True
'        .CellForeColor = vbHighlight
'        .Text = Format(t2, "############0.00")
'End With
'For i = 0 To 5
'   Rst1(i).Close
'Next i
'rsttmpgentbl.Close
'Screen.MousePointer = vbDefault
'frmRptFlexFin.Show 1
'End Sub
'Public Sub FlexSchedule()
'Screen.MousePointer = vbHourglass
'frmRptFlexSch.msfSch.Clear
'frmRptFlexSch.msfSch.Row = 1
'frmRptFlexSch.msfSch.Rows = 2
'DoEvents
'frmRptFlexSch.msfSch.FormatString = "<Sr No       |<Account Name                                                                |>Debit                                              |>Credit                                          "
'Dim Rst1 As Recordset
'Dim i As Integer, srcnt As Integer, OpDiff As Double, q As Integer
'Dim tdt As Date
'Dim currdt As Date
'Dim t1 As Double, t2 As Double
'Dim wDr As Double, wCr As Double
'Dim wTp As String, x As String, x1 As String, ttp(15) As String
'If Check1.Value = 1 Then
'   i = RunClosBal1(gCn, dtpFrDt.Value, dtpToDt.Value, PgrBar1, "S", Rst1, Space(5))
'End If
'OpDiff = 0
'wAcCode = iisGFnGetColumnValue(gCn, "Select ACCODE from tmptblFin where ANAME='" & Trim(txtTemp(0).Text) & "'")
'DoEvents
'If (gSelectedmenu = "R14") Then
'    If wAcCode = "" Then
'        MsgBox "Select Group First", vbCritical + vbOKOnly, "Select Group"
'        txtTemp(0)_GotFocus
'        Exit Sub
'    End If
'    frmRptFlexSch.lblHead.Caption = "For Account  " & txtTemp(0).Text & "  From  " & dtpFrDt.Value & "  To  " & dtpToDt.Value
'    x = "Select ANAME, CLBAL, CLCD, OPBAL, OPCD, ACCODE, AGCODE from tmptblFin where CLBAL > 0 and AGCODE = '" & wAcCode & "' order by ANAME"
'ElseIf (gSelectedmenu = "R121") Then
'    frmRptFlexSch.Caption = "Trial Balance"
'    frmRptFlexSch.lblHead.Caption = "From  " & dtpFrDt.Value & "  To  " & dtpToDt.Value
'    x = "Select ANAME, CLBAL, CLCD, OPBAL, OPCD, ACCODE, AGCODE from tmptblFin where (CLBAL > 0) or (OPBAL > 0) order by ANAME"
'End If
'i = GProcRstOpen(Rst1, x, "R", gCn)
'If (i = 0) Then
'   gSelectedmenu = ""
'   Exit Sub
'End If
'srcnt = 1
't1 = 0
't2 = 0
'frmRptFlexSch.prgBarSch.Max = Rst1.RecordCount
'frmRptFlexSch.prgBarSch.Value = 0
'frmRptFlexSch.prgBarSch.Visible = True
'With Rst1
'    .MoveFirst
'    Do While .EOF = False
'
'        If (gSelectedmenu = "R121") And (.Fields(3) > 0) Then
'           If (.Fields(4) = "C") Then
'              OpDiff = OpDiff + .Fields(3)
'           Else
'              OpDiff = OpDiff - .Fields(3)
'           End If
'       End If
'       x1 = ""
'       If (.Fields(5) = "STOCK") Or (.Fields(6) = "STOCK") Then
'            If (.Fields(3) > 0) Then
'                With frmRptFlexSch.msfSch
'                    .TextMatrix(.Row, 0) = srcnt
'                    .TextMatrix(.Row, 1) = Rst1.Fields(0)
'                    If (Rst1.Fields(4) = "C") Then
'                        t1 = t1 + Rst1.Fields(3)
'                        .TextMatrix(.Row, 3) = Format(Rst1.Fields(3), "#############0.00")
'                   Else
'                        t2 = t2 + Rst1.Fields(3)
'                        .TextMatrix(.Row, 2) = Format(Rst1.Fields(3), "#############0.00")
'                   End If
'                    srcnt = srcnt + 1
'                    .Rows = .Rows + 1
'                    .Row = .Row + 1
'                End With
'            End If
'       Else
'            If (.Fields(1) > 0) Then
'                With frmRptFlexSch.msfSch
'                    .TextMatrix(.Row, 0) = srcnt
'                    .TextMatrix(.Row, 1) = Rst1.Fields(0)
'                    If (Rst1.Fields(2) = "C") Then
'                        t1 = t1 + Rst1.Fields(1)
'                        .TextMatrix(.Row, 3) = Format(Rst1.Fields(1), "#############0.00")
'                    Else
'                        t2 = t2 + Rst1.Fields(1)
'                        .TextMatrix(.Row, 2) = Format(Rst1.Fields(1), "#############0.00")
'                    End If
'                    srcnt = srcnt + 1
'                    .Rows = .Rows + 1
'                    .Row = .Row + 1
'                End With
'            End If
'       End If
'       frmRptFlexSch.prgBarSch.Value = frmRptFlexSch.prgBarSch.Value + 1
'       .MoveNext
'    Loop
'    .Close
'End With
'If (gSelectedmenu = "R121") And (OpDiff <> 0) Then
'    With frmRptFlexSch.msfSch
'        .TextMatrix(.Row, 0) = srcnt
'        .TextMatrix(.Row, 1) = "Difference in Opening Balance"
'        If (OpDiff < 0) Then
'           .TextMatrix(.Row, 3) = Abs(OpDiff)
'           t1 = t1 + Abs(OpDiff)
'        Else
'           .TextMatrix(.Row, 2) = OpDiff
'           t2 = t2 + OpDiff
'        End If
'        .Rows = .Rows + 1
'        .Row = .Row + 1
'    End With
'End If
''----To Calculate Grand Total
'With frmRptFlexSch.msfSch
'    .Rows = .Rows + 1
'    .Row = .Row + 1
'    For q = 1 To 3
'        .Col = q
'            .CellFontBold = True
'            .CellForeColor = vbHighlight
'    Next q
'    .TextMatrix(.Row, 1) = "Total"
'    .TextMatrix(.Row, 2) = Format(t2, "#############0.00")
'    .TextMatrix(.Row, 3) = Format(t1, "#############0.00")
'End With
'frmRptFlexSch.prgBarSch.Visible = False
'Screen.MousePointer = vbDefault
'End Sub
'Private Sub CryTFormat()
'Dim rstqryTrBal As Recordset, rsttmpGenTbl As Recordset
'Dim wtmpGenTbl As Long, wCntCr As Long, wCntDr As Long
'Dim wBkMark
'gCn.Execute "Delete from tmpGenTbl where vUsername='" & gUserName & "' "
'
'If GProcRstOpen(rstqryTrBal, "select * from qryTrBal order by ACNAME", "R") = 0 Then
'    Exit Sub
'End If
'wtmpGenTbl = GProcRstOpen(rsttmpGenTbl, "select * from tmpGenTbl", "O")
'
'wCntCr = 1
'wCntDr = 1
'
'With rstqryTrBal
'    .MoveFirst
'    Do While Not .EOF
'        gCn.BeginTrans
'        If !Debit > 0 Then '---Debit
'            If wCntDr >= wCntCr Then
'                rsttmpGenTbl.AddNew
'                rsttmpGenTbl!docno = wCntDr
'                rsttmpGenTbl!Nar1 = !AcName
'                rsttmpGenTbl!Amt1 = !Debit
'                rsttmpGenTbl.Update
'            Else
'                wBkMark = rsttmpGenTbl.Bookmark
'                rsttmpGenTbl.MoveFirst
'                rsttmpGenTbl.Find "DOCNO= " & wCntDr & ""
'                If rsttmpGenTbl.EOF = False Then
'                    rsttmpGenTbl!Nar1 = !AcName
'                    rsttmpGenTbl!Amt1 = !Debit
'                    rsttmpGenTbl.Update
'                End If
'                rsttmpGenTbl.Bookmark = wBkMark
'            End If
'            wCntDr = wCntDr + 1
'        Else '----Credit
'            If wCntDr <= wCntDr Then
'                rsttmpGenTbl.AddNew
'                rsttmpGenTbl!docno = wCntCr
'                rsttmpGenTbl!Nar2 = !AcName
'                rsttmpGenTbl!AMT2 = !Credit
'                rsttmpGenTbl.Update
'            Else
'                wBkMark = rsttmpGenTbl.Bookmark
'                rsttmpGenTbl.MoveFirst
'                rsttmpGenTbl.Find "DOCNO= " & wCntCr & ""
'                If rsttmpGenTbl.EOF = False Then
'                rsttmpGenTbl!Nar2 = !AcName
'                rsttmpGenTbl!AMT2 = !Credit
'                rsttmpGenTbl.Update
'                End If
'                rsttmpGenTbl.Bookmark = wBkMark
'            End If
'            wCntCr = wCntCr + 1
'        End If
'        gCn.CommitTrans
'        .MoveNext
'    Loop
'End With
'rstqryTrBal.Close
'rsttmpGenTbl.Close
'End Sub
'Private Sub cryBalSht(RptTp As String)
'Dim rstAccount As Recordset, rsttmpGenTbl As Recordset
'Dim Ac1 As Long, Ac2 As Long, Ac3 As Long, Ac4 As Long, Ac5 As Long, Ac6 As Long
'Dim BK1, Bk2, Bk3, Bk4, Bk5, Bk6, Bk7
'Dim t2 As Double, t3 As Double, t4 As Double, t5 As Double, t6 As Double, t7 As Double
'Dim a As Integer, b As Integer, cnt1 As Double, cnt2 As Double
'Dim X As String, wTD2 As String, wTD3 As String, wTD4 As String, wTD5 As String, wTD6 As String, wTD7 As String
'Dim wTopGrp As String, i As Integer, wOpDiff As Double
'Dim wnp As Double, r_s(7) As String, rst1(7) As Recordset
'cnt1 = 0
'cnt2 = 0
'gCn.Execute "Delete from tmpGenTbl where Username='" & gUserName & "' "
'DoEvents
'tblVoucher_TO_tmptblFin dtpToDt
'    r_s(1) = "Select sum(OPBAL),'D' from tmptblFin where  username = '" & gUserName & "' and  OPBAL > 0 and (agcode = 90018 or accode = 90018) " _
' & "Union All Select sum(CLBAL), 'C' from tmptblFin where  username = '" & gUserName & "' and  CLBAL > 0 and (agcode = 90018 or accode = 90018) " _
' & "Union All Select sum(CLBAL), 'D' from tmptblFin where  username = '" & gUserName & "' and  CLBAL > 0 and (topgroup = 90003) and Clcd = 'D' " _
' & "Union All Select sum(CLBAL), 'C' from tmptblFin where  username = '" & gUserName & "' and  CLBAL > 0 and (topgroup = 90003) and Clcd = 'C' "
'    wnp = 0
'    If (GProcRstOpen(rst1(1), r_s(1), "R", gCn) > 0) Then
'       With rst1(1)
'          .MoveFirst
'          Do While Not .EOF
'              If .Fields(1) = "D" Then
'                 wnp = wnp - .Fields(0)
'              Else
'                 wnp = wnp + .Fields(0)
'              End If
'              .MoveNext
'          Loop
'          .Close
'        End With
'    End If
'b = GProcRstOpen(rsttmpGenTbl, "Select Vno, NAR1, AMT1, NAR2, AMT2, Vtype,UserName from tmpgentbl  where username = '" & gUserName & "' order by VNO", "O", gCn)
'wOpDiff = 0
'If GProcRstOpen(rstAccount, "Select sum(vamt) from tblvoucher where vtype = '0P'", "R", gCn) > 0 Then
'   wOpDiff = rstAccount.Fields(0)
'End If
'X = "Select ACCODE, ANAME, AGCODE, CLBAL, CLCD, TOPGroup, SCHEDULE,PFPER,ACTYPE,opbal,opcd From tmptblFin where username = '" & gUserName & "'  and topGroup < 90003 Order By AGCODE, SCHEDULE, ANAME"
'a = GProcRstOpen(rstAccount, X, "R", gCn)
'If a <= 0 Then Exit Sub
'
'Dim wTmpBal As Double, wTmpcd As String
'prgBar1.Max = a
'prgBar1.Value = 0
'prgBar1.Visible = True
'With rstAccount
'    .MoveFirst
'    '---------------- GROUP SPACE(5) A/C - ASSET LIABI
'0:
'    Do While .EOF = False
'        If !AgCode <> !AcCode Then
'            .MoveNext
'            GoTo 0
'        End If
'        Ac1 = !AcCode
'        BK1 = .Bookmark
'        .MoveFirst
'        .Find "AGCODE=" & Ac1
'        '---------------- GROUP ASSET A/C - CURR ASSET
'1:
'        Do While .EOF = False
'            If Not !AgCode = Ac1 Then Exit Do
'            If ((!CLBAL = 0) And (!actype = "A")) Or (!AgCode = !AcCode) Then
'               .MoveNext
'               GoTo 1
'            End If
'            t2 = 0
'            Ac2 = !AcCode
'            wTD2 = !SCHEDULE & vbNullString
'            wTopGrp = !TopGroup & vbNullString
'            Bk2 = .Bookmark
'            Select Case !TopGroup
'            Case 90002:  '2 - Liability
'                If wTD2 = "" Then
'                    cnt1 = cnt1 + 1
'                    FillDataPLA rsttmpGenTbl, "L", cnt1, "", !Aname, IIf(!CLCD = "D", !CLBAL * (-1), !CLBAL)
'                Else
'                    t2 = IIf(!CLCD = "C", t2 + !CLBAL, t2 - !CLBAL)
'                End If
'            Case 90001: '1-Assets
'                If wTD2 = "" Then
'                    cnt2 = cnt2 + 1
'                    FillDataPLA rsttmpGenTbl, "R", cnt2, "", !Aname, IIf(!CLCD = "C", !CLBAL * (-1), !CLBAL)
'                Else
'                    t2 = IIf(!CLCD = "C", t2 + !CLBAL, t2 - !CLBAL)
'                End If
'            End Select
'            .MoveFirst
'            .Find "AgCode=" & Ac2
'            '---------------- GROUP Curr ASSET A/C - Bank
'2:
'            Do While .EOF = False
'                If Not !AgCode = Ac2 Then Exit Do
'                If (Ac2 <> 90004) And (!CLBAL = 0) And (!actype = "A") Then    '4 - "CAPIT"
'                   .MoveNext
'                   GoTo 2
'                End If
'                t3 = 0
'                Ac3 = !AcCode
'                wTD3 = !SCHEDULE & vbNullString
'                wTopGrp = !TopGroup
'                wTmpBal = !CLBAL
'                wTmpcd = !CLCD & vbNullString
'                If (Ac2 = 90004) And (!PfPer > 0) Then  '4 - "CAPIT"
'                    Select Case !CLCD
'                        Case "C":
'                             wTmpBal = wTmpBal + (wnp * !PfPer / 100)
'                             If (wTmpBal < 0) Then
'                                wTmpBal = Abs(wTmpBal)
'                                wTmpcd = "D"
'                             End If
'                        Case "D":
'                            If (wnp < 0) Then
'                               wTmpBal = Abs((wnp * !PfPer / 100)) + wTmpBal
'                            Else
'                            If (wTmpBal > (wnp * !PfPer / 100)) Then
'                               wTmpBal = wTmpBal - (wnp * !PfPer / 100)
'                            Else
'                               wTmpBal = (wnp * !PfPer / 100) - wTmpBal
'                                wTmpcd = "C"
'                            End If
'                            End If
'                    End Select
'                End If
'                Bk3 = .Bookmark
'                Select Case !TopGroup
'                Case 90002: '2 - Liability
'                    If wTD2 <> "" Then
'                        t2 = IIf(wTmpcd = "C", t2 + wTmpBal, t2 - wTmpBal)
'                    Else
'                        If wTD3 = "" Then
'                            cnt1 = cnt1 + 1
'                            FillDataPLA rsttmpGenTbl, "L", cnt1, "", Space(3) & !Aname, IIf(!CLCD = "C", wTmpBal, wTmpBal * (-1))
'                        Else
'                            t3 = IIf(wTmpcd = "C", t3 + wTmpBal, t3 - wTmpBal)
'                        End If
'                    End If
'                 Case 90001: '1-Assets
'                    If wTD2 <> "" Then
'                        t2 = IIf(wTmpcd = "C", t2 + wTmpBal, t2 - wTmpBal)
'                    Else
'                        If wTD3 = "" Then
'                            cnt2 = cnt2 + 1
'                            FillDataPLA rsttmpGenTbl, "R", cnt2, "", Space(3) & !Aname, IIf(!CLCD = "D", wTmpBal, wTmpBal * (-1))
'                        Else
'                            t3 = IIf(wTmpcd = "C", t3 + wTmpBal, t3 - wTmpBal)
'                        End If
'                    End If
'                End Select
'                .MoveFirst
'                .Find "AgCode=" & Ac3
'                '---------------- GROUP Bank A/C - The Ich Bank
'3:
'                Do While .EOF = False
'                    If Not !AgCode = Ac3 Then Exit Do
'                    If (!CLBAL = 0) And (!actype = "A") Then
'                       .MoveNext
'                       GoTo 3
'                    End If
'                    t4 = 0
'                    Ac4 = !AcCode
'                    wTD4 = !SCHEDULE & vbNullString
'                    wTopGrp = !TopGroup & vbNullString
'                    Bk4 = .Bookmark
'                    Select Case !TopGroup
'                    Case 90002: '2 - Liability
'                        If wTD2 = "" Then
'                            If wTD3 = "" Then
'                                cnt1 = cnt1 + 1
'                                FillDataPLA rsttmpGenTbl, "L", cnt1, "", Space(6) & !Aname, IIf(!CLCD = "C", !CLBAL, !CLBAL * (-1))
'                            Else
'                                t3 = IIf(!CLCD = "C", t3 + !CLCD, t3 - !CLCD)
'                            End If
'                        Else
'                            t2 = IIf(!CLCD = "C", t2 + !CLCD, t2 - !CLCD)
'                        End If
'                     Case 90001: '1-Assets
'                        If wTD2 = "" Then
'                            If wTD3 = "" Then
'                                cnt2 = cnt2 + 1
'                                FillDataPLA rsttmpGenTbl, "R", cnt2, "", Space(6) & !Aname, IIf(!CLCD = "C", !CLBAL * (-1), !CLBAL)
'                            Else
'                                t3 = IIf(!CLCD = "C", t3 + !CLCD, t3 - !CLCD)
'                            End If
'                        Else
'                            t2 = IIf(!CLCD = "C", t2 + !CLCD, t2 - !CLCD)
'                        End If
'                    End Select
'                    .MoveFirst
'                    .Find "AgCode=" & Ac4
'                    '---------------- GROUP The Ich Bank A/c - The ich bank branch - ich
'4:
'                    Do While .EOF = False
'                        If Not !AgCode = Ac4 Then Exit Do
'                        If (!CLBAL = 0) And (!actype = "A") Then
'                           .MoveNext
'                           GoTo 4
'                        End If
'                        t5 = 0
'                        Ac5 = !AcCode
'                        wTD5 = !SCHEDULE & vbNullString
'                        wTopGrp = !TopGroup
'                        Bk5 = .Bookmark
'                        Select Case !TopGroup
'                        Case 90002: '2 - Liability
'                            If wTD2 <> "" Then
'                                t2 = IIf(!CLCD = "C", t2 + !CLCD, t2 - !CLCD)
'                            ElseIf wTD3 <> "" Then
'                                t3 = IIf(!CLCD = "C", t3 + !CLCD, t3 - !CLCD)
'                            Else
'                                cnt1 = cnt1 + 1
'                                FillDataPLA rsttmpGenTbl, "L", cnt1, "", Space(9) & !Aname, IIf(!CLCD = "D", !CLBAL * (-1), !CLBAL)
'                            End If
'                         Case 90001:    '1-Assets:
'                            If wTD2 <> "" Then
'                                t2 = IIf(!CLCD = "C", t2 + !CLCD, t2 - !CLCD)
'                            ElseIf wTD3 <> "" Then
'                                t3 = IIf(!CLCD = "C", t3 + !CLCD, t3 - !CLCD)
'                            Else
'                                cnt2 = cnt2 + 1
'                                FillDataPLA rsttmpGenTbl, "R", cnt2, "", Space(9) & !Aname, IIf(!CLCD = "C", !CLBAL * (-1), !CLBAL)
'                            End If
'                        End Select
'                        ''''''''''''''''''
'                        .MoveFirst
'                        .Find "AgCode=" & Ac5
'                        '---------------- GROUP The ich bank branch - ich  A/c - The ich bank (SUB)branch - ich
'                        Do While .EOF = False
'                            If Not !AgCode = Ac5 Then Exit Do
'                            t6 = 0
'                            Ac6 = !AcCode
'                            wTD6 = !SCHEDULE & vbNullString
'                            wTopGrp = !TopGroup
'                            Bk6 = .Bookmark
'                            Select Case !TopGroup
'                            Case 90002: '2 - Liability
'                                If wTD2 <> "" Then
'                                    t2 = IIf(!CLCD = "C", t2 + !CLCD, t2 - !CLCD)
'                                ElseIf wTD3 <> "" Then
'                                    t3 = IIf(!CLCD = "C", t3 + !CLCD, t3 - !CLCD)
'                                ElseIf wTD4 <> "" Then
'                                    t4 = IIf(!CLCD = "C", t4 + !CLCD, t4 - !CLCD)
'                                Else
'                                    cnt1 = cnt1 + 1
'                                    FillDataPLA rsttmpGenTbl, "L", cnt1, "", Space(12) & !Aname, IIf(!CLCD = "D", !CLBAL * (-1), !CLBAL)
'                                End If
'                             Case 90001:    '1-Assets:
'                                If wTD2 <> "" Then
'                                    t2 = IIf(!CLCD = "C", t2 + !CLCD, t2 - !CLCD)
'                                ElseIf wTD3 <> "" Then
'                                    t3 = IIf(!CLCD = "C", t3 + !CLCD, t3 - !CLCD)
'                                ElseIf wTD4 <> "" Then
'                                    t4 = IIf(!CLCD = "C", t4 + !CLCD, t4 - !CLCD)
'                                Else
'                                    cnt2 = cnt2 + 1
'                                    FillDataPLA rsttmpGenTbl, "R", cnt2, "", Space(12) & !Aname, IIf(!CLCD = "C", !CLBAL * (-1), !CLBAL)
'                                End If
'                            End Select
'                            .Bookmark = Bk6
'                            .MoveNext
''                            prgBar1.Value = prgBar1.Value + 1
'                        Loop
'                        If t5 <> 0 Then
'                            If wTopGrp = 90002 Then '2 - Liability
'                                cnt1 = cnt1 + 1
'                                FillDataPLA rsttmpGenTbl, "L", cnt1, "", Space(12) & "As Per Schedule " & wTD5, t5
'                            ElseIf wTopGrp = 90001 Then: '1-Assets
'                                cnt2 = cnt2 + 1
'                                FillDataPLA rsttmpGenTbl, "R", cnt2, "", Space(12) & "As Per Schedule " & wTD5, t5 * (-1)
'                            End If
'                        End If
'                        '''''''''''''''''
'                        .Bookmark = Bk5
'                        .MoveNext
''                        prgBar1.Value = prgBar1.Value + 1
'                    Loop
'                    If t4 <> 0 Then
'                        If wTopGrp = 90002 Then '2 - Liability
'                            cnt1 = cnt1 + 1
'                            FillDataPLA rsttmpGenTbl, "L", cnt1, "", Space(9) & "As Per Schedule " & wTD4, t4
'                        ElseIf wTopGrp = 90001 Then '1-Assets
'                            cnt2 = cnt2 + 1
'                            FillDataPLA rsttmpGenTbl, "R", cnt2, "", Space(9) & "As Per Schedule " & wTD4, t4 * (-1)
'                        End If
'                    End If
'                    .Bookmark = Bk4
'                    .MoveNext
''                    prgBar1.Value = prgBar1.Value + 1
'                Loop
'                If t3 <> 0 Then
'                    If wTopGrp = 90002 Then '2 - Liability
'                        cnt1 = cnt1 + 1
'                        FillDataPLA rsttmpGenTbl, "L", cnt1, "", Space(6) & "As Per Schedule " & wTD3, t3
'                    ElseIf wTopGrp = 90001 Then '1-Assets Then
'                        cnt2 = cnt2 + 1
'                        FillDataPLA rsttmpGenTbl, "R", cnt2, "", Space(6) & "As Per Schedule " & wTD3, t3 * (-1)
'                    End If
'                End If
'                rstAccount.Bookmark = Bk3
'                rstAccount.MoveNext
''                prgBar1.Value = prgBar1.Value + 1
'            Loop
'            If t2 <> 0 Then
'                If wTopGrp = 90002 Then '2 - Liability
'                    cnt1 = cnt1 + 1
'                    FillDataPLA rsttmpGenTbl, "L", cnt1, "", Space(3) & "As Per Schedule " & wTD2, t2
'                ElseIf wTopGrp = 90001 Then '1-Assets Then
'                    cnt2 = cnt2 + 1
'                    FillDataPLA rsttmpGenTbl, "R", cnt2, "", Space(3) & "As Per Schedule " & wTD2, t2 * (-1)
'                End If
'            End If
'            rstAccount.Bookmark = Bk2
'            rstAccount.MoveNext
'            'prgBar1.Value = prgBar1.Value + 1
'        Loop
'        rstAccount.Bookmark = BK1
'        rstAccount.MoveNext
''        prgBar1.Value = prgBar1.Value + 1
'    Loop
'End With
'If wOpDiff <> 0 Then
'    If wOpDiff < 0 Then
'        cnt1 = cnt1 + 1
'        FillDataPLA rsttmpGenTbl, "L", cnt1, "", "Opening Balance Difference ", Abs(wOpDiff)
'    Else
'        cnt2 = cnt2 + 1
'        FillDataPLA rsttmpGenTbl, "R", cnt2, "", "Opening Balance Difference ", wOpDiff
'    End If
'End If
'
'prgBar1.Visible = False
''rsttmpGenTbl.Close
'rstAccount.Close
'DoEvents
'If RptTp = "G" Then
'    Dim LnCnt As Integer, PgCnt As Integer
'    Dim wHdPrn As Boolean, X1 As String
'    Dim wTotLia As Double, wTotAss As Double
'    Dim wDotLine As String
'    wTotLia = 0
'    wTotAss = 0
'    wDotLine = String(136, "-")
'    If GProcRstOpen(rsttmpGenTbl, "Select NAR1, AMT1, NAR2, AMT2 from tmpgentbl ", "R", gCn) = 0 Then
'       gSelectedMenu = ""
'       Exit Sub
'    End If
'    If mScreen = True Then
'        mP12 = ""
'    Else
'        mP12 = Chr(27) + Chr(15)
'    End If
'    Open "c:\rpt1.txt" For Output As #1
'    wHdPrn = True
'    LnCnt = 0
'    PgCnt = 1
'    With rsttmpGenTbl
'        .MoveFirst
'        Do While Not .EOF
'            If wHdPrn Then
'                wHdPrn = False
'                LnCnt = 5
'                If mScreen = True Then
'                    Print #1, gCName
'                    Print #1, "BALANCE SHEET From " + Format(dtpFrDt, "dd/MM/yyyy") + " To " + Format(dtpToDt, "dd/MM/yyyy") & "          Page No. " + Str(PgCnt)
'                    Print #1, wDotLine
'                Else
'                    Print #1, Chr(18) + Chr(27) + Chr(14) & gCName & Chr(27) + Chr(18)
'                    Print #1, Chr(27) + Chr(69) & "BALANCE SHEET " & Chr(27) + Chr(70) & " From " + Format(dtpFrDt, "dd/MM/yyyy") + " To " + Format(dtpToDt, "dd/MM/yyyy") & "          Page No. " + Str(PgCnt)
'                    Print #1, mP12 & wDotLine
'                End If
'                Print #1, "LIABILITY                                                   AMOUNT | ASSET                                                       AMOUNT"
'                Print #1, wDotLine
'                PgCnt = PgCnt + 1
'            End If
'            X1 = ""
'            If Not IsNull(!Nar1) Then
'                X1 = Left(!Nar1 + Space(48), 48) + "  "
'                If !Amt1 <> 0 Then
'                    X1 = X1 + GProcNoToStr(!Amt1, 16, 2)
'                    wTotLia = wTotLia + !Amt1
'                Else
'                    X1 = X1 + Space(16)
'                End If
'            Else
'                X1 = Space(64)
'            End If
'            X1 = X1 + " | "
'            If Not IsNull(!Nar2) Then
'                 X1 = X1 + Left(!Nar2 + Space(48), 48) + "  "
'                 If !AMT2 <> 0 Then
'                     X1 = X1 + GProcNoToStr(!AMT2, 16, 2)
'                     wTotAss = wTotAss + !AMT2
'                 Else
'                     X1 = X1 + Space(16)
'                 End If
'            End If
'            Print #1, X1
'            LnCnt = LnCnt + 1
'            If (LnCnt > 65) Then
'                 If mScreen = True Then
'                     Print #1, wDotLine
'                 Else
'                     Print #1, wDotLine & Chr(12)
'                 End If
'               wHdPrn = True
'               LnCnt = 0
'            End If
'            .MoveNext
'        Loop
'        .Close
'    End With
'    Print #1, wDotLine
'    Print #1, "Total" + Space(43) + GProcNoToStr(Round(wTotLia, 2), 18, 2) + " |" + Space(49) + GProcNoToStr(Round(wTotAss, 2), 18, 2)
'    If mScreen = True Then
'        Print #1, wDotLine
'    Else
'        Print #1, wDotLine & Chr(27) + Chr(15) + Chr(12)
'    End If
'    Close #1
'End If
'End Sub
'Private Sub All_Schedule()
'Dim rst1 As Recordset
'Dim i As Integer
'Dim wSndDrTot As Double, wSndCrTot As Double
'Dim wTp As String, X As String, X1 As String, X2 As String, wGrp As String
'Dim wSndDrGrp As Double, wSndCrGrp As Double
'wDotLine = String(80, "-")
'tblVoucher_TO_tmptblFin dtpToDt
'wGrp = "XZXSDSDFFSDV"
'X = "Select A.ANAME, A.CLBAL, A.CLCD, A.OPBAL, A.OPCD, A.ACCODE, A.AGCODE, B.ANAME from tmptblFin A, tmptblFin B where A.CLBAL > 0 and A.AGCODE = B.ACCODE  order by B.ANAME, A.ANAME"
'i = GProcRstOpen(rst1, X, "R", gCn)
'If (i = 0) Then
'   gSelectedMenu = ""
'   Exit Sub
'End If
'Open "c:\rpt1.txt" For Output As #1
'wHdPrn = True
'wSndDrTot = 0
'wSndCrTot = 0
'wSndDrGrp = 0
'wSndCrGrp = 0
'LnCnt = 0
'PgCnt = 1
'srcnt = 1
'With rst1
'    .MoveFirst
'    Do While .EOF = False
'          If wHdPrn Then
'             wHdPrn = False
'             If mScreen = True Then
'                Print #1, gCName
'             Else
'                Print #1, Chr(18) + Chr(27) + Chr(14) & gCName & Chr(27) + Chr(18)
'             End If
'             Print #1, "For Period From " & dtpFrDt.Value & "  To  " & dtpToDt.Value & "                  Page No. " + Str(PgCnt)
'             If mScreen = True Then
'                Print #1, "All Schedule Printing "
'             Else
'                Print #1, Chr(27) + Chr(69) & "All Schedule Printing " & Chr(27) + Chr(70)
'             End If
'             Print #1, wDotLine
'             Print #1, "SrNo  Account Name                                   Debit                Credit"
'             Print #1, wDotLine
'             LnCnt = 6
'             PgCnt = PgCnt + 1
'          End If
'          If wGrp <> .Fields(7) Then
'            If (wSndDrGrp + wSndCrGrp <> 0) Then
'                Print #1, Space(38) + String(42, "-")
'                X2 = ""
'                If wSndDrGrp <> 0 Then
'                    X2 = X2 + Space(38) + GProcNoToStr(wSndDrGrp, 20, 2) + "  "
'                Else
'                    X2 = X2 + Space(60)
'                End If
'                If wSndCrGrp <> 0 Then
'                    X2 = X2 + GProcNoToStr(wSndCrGrp, 20, 2)
'                Else
'                    X2 = X2 + Space(20)
'                End If
'                Print #1, X2
'                Print #1, Space(38) + String(42, "-")
'                wSndDrGrp = 0
'                wSndCrGrp = 0
'                LnCnt = LnCnt + 3
'            End If
'            wGrp = .Fields(7)
'            Print #1, wGrp
'            LnCnt = LnCnt + 1
'          End If
'          X1 = ""
'          If (.Fields(6) = 17) Then
'               X1 = X1 + GProcNoToStr(.Fields(1), 20, 2)
'               wSndDrGrp = wSndDrGrp + .Fields(1)
'               wSndDrTot = wSndDrTot + .Fields(1)
'          Else
'               X1 = X1 + Space(22) + GProcNoToStr(.Fields(1), 20, 2)
'               wSndCrGrp = wSndCrGrp + .Fields(1)
'               wSndCrTot = wSndCrTot + .Fields(1)
'          End If
'          Print #1, Right("    " + Str(srcnt), 4) + "  " + Left(.Fields(0) + Space(30), 30) & "  " & X1
'          LnCnt = LnCnt + 1
'          srcnt = srcnt + 1
'          If (LnCnt > 65) Then
'             If mScreen = True Then
'                Print #1, wDotLine
'             Else
'                Print #1, wDotLine + Chr(12)
'             End If
'             wHdPrn = True
'             LnCnt = 0
'          End If
'       .MoveNext
'    Loop
'    .Close
'End With
'If (wSndDrGrp + wSndCrGrp <> 0) Then
'    Print #1, Space(38) + String(42, "-")
'    X2 = ""
'    If wSndDrGrp <> 0 Then
'        X2 = X2 + Space(38) + GProcNoToStr(wSndDrGrp, 20, 2) + "  "
'    Else
'        X2 = X2 + Space(60)
'    End If
'    If wSndCrGrp <> 0 Then
'        X2 = X2 + GProcNoToStr(wSndCrGrp, 20, 2)
'    Else
'        X2 = X2 + Space(20)
'    End If
'    Print #1, X2
'    Print #1, Space(38) + String(42, "-")
'    LnCnt = LnCnt + 3
'End If
'Print #1, wDotLine
'Print #1, Space(37) + "  " + GProcNoToStr(wSndDrTot, 20, 2) + " " + GProcNoToStr(wSndCrTot, 20, 2)
'If mScreen = True Then
'   Print #1, wDotLine
'Else
'   Print #1, wDotLine + Chr(12)
'End If
'Close #1
'End Sub

Private Sub CryTFormat()
Dim rstqryTrBal As Recordset, rsttmpGenTbl As Recordset
Dim wtmpGenTbl As Long, wCntCr As Long, wCntDr As Long
Dim wBkMark

gCn.Execute "Delete from tmpGenTbl where Username='" & gUserName & "' "

If GProcRstOpen(rstqryTrBal, "select * from qryTrbalDetail where ClBal_Dr+ClBal_Cr > 0 and Username='" & gUserName & "' order by ACNAME", "R") = 0 Then
    Exit Sub
End If

wtmpGenTbl = GProcRstOpen(rsttmpGenTbl, "select * from tmpGenTbl where userName='" & gUserName & "' order by vno", "O")

wCntCr = 1
wCntDr = 1

With rstqryTrBal
    .MoveFirst
    Do While Not .EOF
        gCn.BeginTrans
        If !ClBal_Dr > 0 Then '---Debit
            If wCntDr >= wCntCr Then
                rsttmpGenTbl.AddNew
                rsttmpGenTbl!Vno = wCntDr
                rsttmpGenTbl!Nar1 = !AcName
                rsttmpGenTbl!Amt1 = !ClBal_Dr
                rsttmpGenTbl!UserName = gUserName
                rsttmpGenTbl.Update
            Else
                wBkMark = rsttmpGenTbl.Bookmark
                rsttmpGenTbl.MoveFirst
                rsttmpGenTbl.Find "vno= " & wCntDr & ""
                If rsttmpGenTbl.EOF = False Then
                   gCn.Execute "update tmpGenTbl set Amt1=" & .Fields("ClBal_Dr") & " , Nar1= '" & .Fields("AcName") & "'  where vno = " & wCntDr & " and Username='" & gUserName & "'  "
                End If
                rsttmpGenTbl.Bookmark = wBkMark
            End If
            wCntDr = wCntDr + 1
        Else '----Credit
'            If wCntDr <= wCntCr Then
            If wCntDr <= wCntCr Then
                rsttmpGenTbl.AddNew
                rsttmpGenTbl!Vno = wCntCr
                rsttmpGenTbl!Nar2 = !AcName
                rsttmpGenTbl!Amt2 = !ClBal_Cr
                rsttmpGenTbl!UserName = gUserName
                rsttmpGenTbl.Update
            Else
                wBkMark = rsttmpGenTbl.Bookmark
                rsttmpGenTbl.MoveFirst
                rsttmpGenTbl.Find "vno= " & wCntCr & ""
                If rsttmpGenTbl.EOF = False Then
                   gCn.Execute "update tmpGenTbl set Amt2=" & .Fields("ClBal_Cr") & " , Nar2= '" & .Fields("AcName") & "'  where vno = " & wCntCr & " and Username='" & gUserName & "'  "
                End If
                rsttmpGenTbl.Bookmark = wBkMark
            End If
            wCntCr = wCntCr + 1
        End If
        gCn.CommitTrans
        .MoveNext
    Loop
End With
rstqryTrBal.Close
rsttmpGenTbl.Close
End Sub

