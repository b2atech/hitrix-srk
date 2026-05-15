VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "comctl32.ocx"
Begin VB.Form frmUtlAccountTransfer 
   BackColor       =   &H00C2E3FC&
   BorderStyle     =   0  'None
   Caption         =   " "
   ClientHeight    =   2910
   ClientLeft      =   0
   ClientTop       =   60
   ClientWidth     =   9435
   FillColor       =   &H000040C0&
   FillStyle       =   0  'Solid
   Icon            =   "frmUtlAccountTransfer.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   2910
   ScaleWidth      =   9435
   ShowInTaskbar   =   0   'False
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   0
      Top             =   315
   End
   Begin VB.CommandButton cmdExit 
      Cancel          =   -1  'True
      Caption         =   "E&xit"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   400
      Left            =   7515
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   1425
      Width           =   1290
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "&OK"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   400
      Left            =   7515
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   840
      Width           =   1290
   End
   Begin VB.TextBox txtTo 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   2070
      TabIndex        =   3
      Text            =   "txtTo"
      Top             =   1365
      Width           =   4695
   End
   Begin VB.TextBox txtFrom 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   2070
      TabIndex        =   1
      Text            =   "txtFrom"
      Top             =   855
      Width           =   4695
   End
   Begin ComctlLib.ProgressBar ProgressBar1 
      Height          =   315
      Left            =   420
      TabIndex        =   7
      Top             =   2085
      Visible         =   0   'False
      Width           =   6345
      _ExtentX        =   11192
      _ExtentY        =   556
      _Version        =   327682
      Appearance      =   1
   End
   Begin VB.Shape Shape1 
      FillColor       =   &H00BDDFD1&
      FillStyle       =   0  'Solid
      Height          =   1545
      Left            =   7245
      Shape           =   4  'Rounded Rectangle
      Top             =   615
      Width           =   1830
   End
   Begin VB.Label lbl1 
      BackStyle       =   0  'Transparent
      Caption         =   "&To Account"
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
      Height          =   360
      Index           =   1
      Left            =   420
      TabIndex        =   2
      Top             =   1365
      Width           =   1320
   End
   Begin VB.Shape Shape3 
      BorderColor     =   &H00C00000&
      BorderWidth     =   5
      FillColor       =   &H00C00000&
      Height          =   4335
      Left            =   0
      Top             =   -240
      Width           =   9435
   End
   Begin VB.Shape Shape5 
      BorderColor     =   &H00C00000&
      FillColor       =   &H00C00000&
      FillStyle       =   0  'Solid
      Height          =   1785
      Left            =   0
      Top             =   2850
      Width           =   9420
   End
   Begin VB.Label lbl1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "&From Account"
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
      Left            =   420
      TabIndex        =   0
      Top             =   855
      Width           =   1320
   End
   Begin VB.Label Label6 
      BackColor       =   &H00C00000&
      Caption         =   "  Utility Account Transfer"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   390
      Left            =   -15
      TabIndex        =   6
      Top             =   0
      Width           =   9405
   End
   Begin VB.Shape Shape2 
      BackColor       =   &H00808080&
      BorderColor     =   &H004D177D&
      FillColor       =   &H00ECFFFF&
      FillStyle       =   0  'Solid
      Height          =   2100
      Left            =   240
      Top             =   615
      Width           =   6840
   End
End
Attribute VB_Name = "frmUtlAccountTransfer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub cmdExit_Click()
Unload Me
End Sub
Private Sub cmdOK_Click()
Dim FromAcCode As Long
Dim ToAcCode As Long
'-- Validate
FromAcCode = GProcGetColumnValue("tblMastAccount", "AcName", txtFrom, "S", "AcCode", "N")
ToAcCode = GProcGetColumnValue("tblMastAccount", "AcName", txtTo, "S", "AcCode", "N")
'-- From A/c
If FromAcCode = 0 Then
    MsgBox "Check From Account.", vbOKOnly + vbInformation, App.Title
    txtFrom.SetFocus
    Exit Sub
End If
'-- To A/c
If ToAcCode = 0 Then
    MsgBox "Check To Account.", vbOKOnly + vbInformation, App.Title
    txtTo.SetFocus
    Exit Sub
End If
If FromAcCode = ToAcCode Then
    MsgBox "From Account and To Account Should not be same.", vbOKOnly + vbInformation, App.Title
    txtTo.SetFocus
    Exit Sub
End If
If MsgBox("Account " & txtFrom & " will be transfer to account " & txtTo & " . Are You Sure ?.You can not Redo affected Transfer.", vbOKCancel + vbInformation + vbDefaultButton2) = vbOK Then
    Timer1.Enabled = True
    ProgressBar1.Visible = True
    gCn.BeginTrans
    gSetupCn.BeginTrans
    '========= Set Up Booking
    '-- tblBooking
    GProcUpdateAcCode "tblBooking", "BkAcBrkCode", FromAcCode, ToAcCode, "BkCCode='" & gCCode & "' and BkCYear='" & gDbYr & "'", gSetupCn
    GProcUpdateAcCode "tblBooking", "BKAcCode", FromAcCode, ToAcCode, "BkCCode='" & gCCode & "' and BkCYear='" & gDbYr & "'", gSetupCn
    '-- Mast Account
    GProcUpdateAcCode "tblmastAccount", "ACBRKCODE", FromAcCode, ToAcCode
    '--Mast Item
    GProcUpdateAcCode "tblmastItem", "ITMILLCODE", FromAcCode, ToAcCode
    '--Setting
    GProcUpdateAcCode "tblmastSetting", "LATEPAYINTACCODE", FromAcCode, ToAcCode
    GProcUpdateAcCode "tblmastSetting", "BILLDBNTACCODE", FromAcCode, ToAcCode
    GProcUpdateAcCode "tblmastSetting", "SERVICETAXACCODE", FromAcCode, ToAcCode
    GProcUpdateAcCode "tblmastSetting", "TDSACCODE", FromAcCode, ToAcCode
    GProcUpdateAcCode "tblmastSetting", "BANKCOMMACCODE", FromAcCode, ToAcCode
    GProcUpdateAcCode "tblmastSetting", "BROKERAGEACCODE", FromAcCode, ToAcCode
    GProcUpdateAcCode "tblmastSetting", "ACCODEPY", FromAcCode, ToAcCode
    GProcUpdateAcCode "tblmastSetting", "ACCODEPT", FromAcCode, ToAcCode
    GProcUpdateAcCode "tblmastSetting", "ACCODESY", FromAcCode, ToAcCode
    GProcUpdateAcCode "tblmastSetting", "ACCODEST", FromAcCode, ToAcCode
    GProcUpdateAcCode "tblmastSetting", "EXPENSESACCODE", FromAcCode, ToAcCode
    GProcUpdateAcCode "tblmastSetting", "MILLBILLDBNTACCODE", FromAcCode, ToAcCode
    GProcUpdateAcCode "tblmastSetting", "TDSPaidAcCode", FromAcCode, ToAcCode
    '-- Narration
    GProcUpdateAcCode "tblmastNarration", "MASTTAXPURACCODE", FromAcCode, ToAcCode
    GProcUpdateAcCode "tblmastNarration", "MASTTAXSAACCODE", FromAcCode, ToAcCode
    '-- tblAddless
    GProcUpdateAcCode "tblAddLess", "ADACCRCODE", FromAcCode, ToAcCode, "Vyear=" & gCYear
    GProcUpdateAcCode "tblAddLess", "ADACDRCODE", FromAcCode, ToAcCode, "Vyear=" & gCYear
    GProcUpdateAcCode "tblAddLess", "ADMILLCODE", FromAcCode, ToAcCode, "Vyear=" & gCYear
    GProcUpdateAcCode "tblAddLess", "AdAcDbNtCode", FromAcCode, ToAcCode, "Vyear=" & gCYear
    '-- Tblvousub
    GProcUpdateAcCode "tblVousub", "VSUBBOOKACCODE", FromAcCode, ToAcCode, "Vyear=" & gCYear
    GProcUpdateAcCode "tblVousub", "VSUBBOOKBRKACCODE", FromAcCode, ToAcCode, "Vyear=" & gCYear
    '-- TblVoucher
    GProcUpdateAcCode "TblVoucher", "VACCODE", FromAcCode, ToAcCode, "Vyear=" & gCYear
    GProcUpdateAcCode "TblVoucher", "VACOCODE", FromAcCode, ToAcCode, "Vyear=" & gCYear
    '-- TblOutStanding
    GProcUpdateAcCode "TblOutStanding", "OUTACCODE", FromAcCode, ToAcCode, "Vyear=" & gCYear
    GProcUpdateAcCode "TblOutStanding", "OUTBROKERCODE", FromAcCode, ToAcCode, "Vyear=" & gCYear
    GProcUpdateAcCode "TblOutStanding", "OUTMILLCODE", FromAcCode, ToAcCode, "Vyear=" & gCYear
    '-- TblRecVsSale
    GProcUpdateAcCode "TBLRECVSSALE", "PARTYCODE", FromAcCode, ToAcCode, "INVYEAR=" & gCYear
    '-- TblMillRecPay
    GProcUpdateAcCode "TblMillRecPay", "MRecAcCode", FromAcCode, ToAcCode, "Vyear=" & gCYear
    GProcUpdateAcCode "TblMillRecPay", "MRecAcOCode", FromAcCode, ToAcCode, "Vyear=" & gCYear
    '-- tblClosingStock
    GProcUpdateAcCode "tblClosingStock", "StAcCode", FromAcCode, ToAcCode
    '-- tblConfirmation
    GProcUpdateAcCode "tblConfirmation", "CfCrAcCode", FromAcCode, ToAcCode, "Vyear=" & gCYear
    GProcUpdateAcCode "tblConfirmation", "CfDrAcCode", FromAcCode, ToAcCode, "Vyear=" & gCYear
    '-- tblTDS
    GProcUpdateAcCode "tblTDS", "TDSAcCode", FromAcCode, ToAcCode, "Vyear=" & gCYear
    GProcUpdateAcCode "tblTDS", "TDSDeductTaxAcCode", FromAcCode, ToAcCode, "Vyear=" & gCYear
    gCn.CommitTrans
    gSetupCn.CommitTrans
    ProgressBar1.Visible = False
    Timer1.Enabled = False
    Unload Me
    End If
End Sub
Private Sub Form_Load()
Me.Top = 300
Me.Left = Screen.Width / 2 - Me.Width / 2
txtFrom = ""
txtTo = ""
End Sub
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case 13:
           SendKeys "{TAB}"
    End Select
End Sub

Private Sub txtFrom_GotFocus()
GProcSelectBox Me.ActiveControl
End Sub

Private Sub txtFrom_KeyPress(KeyAscii As Integer)
Dim LeftPos As Long
Dim TopPos As Long
TopPos = txtFrom.Top + Me.Top + 650
LeftPos = txtFrom.Left + Me.Left
If KeyAscii = 13 Then Exit Sub
gClsSearch.SearchMultiField "QryMastaccount", "Account,AcName,AcCity,AgName", Array("Party", "AcName", "City", "Group"), Array(3000, 0, 2000, 3000), "", Chr(KeyAscii), "acName", LeftPos, TopPos
KeyAscii = 0
If gClsSearch.SearchMultiRetCol(1) <> "" Then
    txtFrom = gClsSearch.SearchMultiRetCol(1)
End If
End Sub
Private Sub txtTo_GotFocus()
GProcSelectBox Me.ActiveControl
End Sub
Private Sub txtTo_KeyPress(KeyAscii As Integer)
Dim LeftPos As Long
Dim TopPos As Long
TopPos = txtTo.Top + Me.Top + 650
LeftPos = txtTo.Left + Me.Left
If KeyAscii = 13 Then Exit Sub
gClsSearch.SearchMultiField "QryMastaccount", "Account,AcName,AcCity,AgName", Array("Party", "AcName", "City", "Group"), Array(3000, 0, 2000, 3000), "", Chr(KeyAscii), "acName", LeftPos, TopPos
KeyAscii = 0
If gClsSearch.SearchMultiRetCol(1) <> "" Then
    txtTo = gClsSearch.SearchMultiRetCol(1)
End If
End Sub

Private Sub txtFrom_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyF3 Then GProcShowForm frmMastAccount, frmMain.mnuMstfrm(0)
End Sub
Private Sub txtTo_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyF3 Then GProcShowForm frmMastAccount, frmMain.mnuMstfrm(0)
End Sub
Private Sub Timer1_Timer()
If ProgressBar1.Value < 90 Then
    ProgressBar1.Value = ProgressBar1.Value + 10
Else
    ProgressBar1.Value = 0
End If
DoEvents
End Sub

