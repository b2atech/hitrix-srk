VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form frmRptStatement 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Reports"
   ClientHeight    =   5535
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6615
   Icon            =   "frmRptStatement.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5535
   ScaleWidth      =   6615
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
      Index           =   6
      Left            =   1905
      TabIndex        =   24
      Text            =   "6"
      Top             =   3240
      Width           =   1395
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
      Index           =   5
      Left            =   1905
      TabIndex        =   11
      Text            =   "5"
      Top             =   1710
      Width           =   1395
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
      Left            =   4740
      TabIndex        =   9
      Text            =   "4"
      Top             =   1245
      Width           =   1530
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
      Left            =   1905
      TabIndex        =   17
      Text            =   "3"
      Top             =   2715
      Width           =   1395
   End
   Begin VB.CommandButton cmdCrystal 
      Height          =   480
      Index           =   0
      Left            =   1410
      Picture         =   "frmRptStatement.frx":058A
      Style           =   1  'Graphical
      TabIndex        =   20
      Top             =   4605
      Width           =   1185
   End
   Begin VB.CommandButton cmdCrystal 
      Height          =   480
      Index           =   1
      Left            =   2715
      Picture         =   "frmRptStatement.frx":09CC
      Style           =   1  'Graphical
      TabIndex        =   21
      ToolTipText     =   "Print"
      Top             =   4605
      Width           =   1185
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
      Index           =   0
      Left            =   1905
      TabIndex        =   7
      Text            =   "0"
      Top             =   1245
      Width           =   1395
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
      Left            =   4740
      TabIndex        =   15
      Text            =   "1"
      Top             =   2220
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
      Index           =   2
      Left            =   1905
      TabIndex        =   13
      Text            =   "2"
      Top             =   2205
      Width           =   1395
   End
   Begin VB.TextBox txtMillName 
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
      Left            =   1905
      TabIndex        =   5
      Top             =   765
      Width           =   4410
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
      Left            =   4020
      TabIndex        =   22
      Top             =   4605
      Width           =   1185
   End
   Begin Crystal.CrystalReport CrystalReport1 
      Left            =   240
      Top             =   4200
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
   Begin MSComctlLib.ProgressBar prgBar 
      Height          =   255
      Left            =   120
      TabIndex        =   23
      Top             =   3765
      Visible         =   0   'False
      Width           =   6135
      _ExtentX        =   10821
      _ExtentY        =   450
      _Version        =   393216
      Appearance      =   1
   End
   Begin MSComCtl2.DTPicker dtpToDt 
      Height          =   360
      Left            =   4755
      TabIndex        =   3
      Top             =   315
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
      Format          =   52297731
      CurrentDate     =   38050
   End
   Begin MSComCtl2.DTPicker dtpFrDt 
      Height          =   360
      Left            =   1905
      TabIndex        =   1
      Top             =   315
      Width           =   1395
      _ExtentX        =   2461
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
      Format          =   52297731
      CurrentDate     =   38050
   End
   Begin MSComCtl2.DTPicker dtpBillDt 
      Height          =   360
      Left            =   4740
      TabIndex        =   19
      Top             =   2715
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
      Format          =   52297731
      CurrentDate     =   38050
   End
   Begin VB.Shape Shape2 
      FillColor       =   &H00BDDFD1&
      FillStyle       =   0  'Solid
      Height          =   780
      Left            =   1185
      Shape           =   4  'Rounded Rectangle
      Top             =   4440
      Width           =   4320
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "B&rokerage %  :"
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
      Index           =   7
      Left            =   180
      TabIndex        =   25
      Top             =   3240
      Width           =   1305
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "TDS E .&Cess %  :"
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
      Index           =   6
      Left            =   180
      TabIndex        =   10
      Top             =   1710
      Width           =   1560
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "&Surcharge % :"
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
      Index           =   5
      Left            =   3405
      TabIndex        =   8
      Top             =   1245
      Width           =   1245
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Bill Date      :"
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
      Left            =   3405
      TabIndex        =   18
      Top             =   2715
      Width           =   1155
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Service Tax &Bill No."
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
      Left            =   180
      TabIndex        =   16
      Top             =   2715
      Width           =   1695
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "&TDS %    :"
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
      TabIndex        =   6
      Top             =   1245
      Width           =   915
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Service Tax % :"
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
      TabIndex        =   12
      Top             =   2205
      Width           =   1365
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "E. Cess  %  :"
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
      Left            =   3405
      TabIndex        =   14
      Top             =   2220
      Width           =   1185
   End
   Begin VB.Label lblMillName 
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
      Top             =   750
      Width           =   1350
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "&To    :"
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
      Left            =   4020
      TabIndex        =   2
      Top             =   315
      Width           =   540
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "From &Date     :"
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
      Top             =   315
      Width           =   1350
   End
   Begin VB.Shape Shape6 
      BorderColor     =   &H00000080&
      FillColor       =   &H00EDFDFE&
      FillStyle       =   0  'Solid
      Height          =   4155
      Left            =   75
      Top             =   90
      Width           =   6345
   End
End
Attribute VB_Name = "frmRptStatement"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub cmdCrystal_Click(Index As Integer)
Dim FromDt As Date
Dim ToDt As Date
Dim X As String
FromDt = Format(dtpFrDt.Value, "dd/MM/yyyy")
ToDt = Format(dtpToDt.Value, "dd/MM/yyyy")
GProcCrystalRptPreparation CrystalReport1, IIf(Index = 0, "Window", "Printer")
With CrystalReport1
    Select Case gSelectedMenu
        Case "Consignment/ Depot Statement"  '--- "Consignment/ Depot Statement"
            If ValidateData = False Then Exit Sub
            GProcConsigDepotMillStatement FromDt, ToDt, CrystalReport1, txtMillName, txt(0), txt(1), txt(2), txt(3), txt(4), txt(5), Format(dtpBillDt.Value, "dd/MM/yyyy")
            Exit Sub
        Case "Sales In Transit Statement" '"Sales In Transit Statement"
              If ValidateData = False Then Exit Sub
             .Formulas(4) = "wFrToDt = 'From : ' & '" & Format(FromDt, "dd/MM/yyyy") & "' & '  To ' & '" & Format(ToDt, "dd/MM/yyyy") & "'"
             '-- SIT detail
             X = "{tblAddLess.vdt} in Date(" & Format(FromDt, "yyyy,mm,dd") & ") to Date(" & Format(ToDt, "yyyy,mm,dd") & ")"
             X = X + " and {tblAddLess.VType}='PT' and {tblMastAccount.AcName}='" & txtMillName & "'"
             .WindowTitle = "SIT Purchase Details"
            .SelectionFormula = X
            .ReportFileName = gReportPath & "rptMonSTMSITPurchase.rpt"
            .Action = 1
            '-- Brokerage DbNT
            .Formulas(5) = "BrkP = '" & txt(6) & "'"
            .Formulas(6) = "ServiceTaxP = '" & txt(2) & "'"
             X = "{tblAddLess.vdt} in Date(" & Format(FromDt, "yyyy,mm,dd") & ") to Date(" & Format(ToDt, "yyyy,mm,dd") & ")"
             X = X + " and {tblAddLess.VType}='PT' and {tblMastAccount.AcName}='" & txtMillName & "'"
             .WindowTitle = "Brokerage Debit Note On SIT Purchase"
            .SelectionFormula = X
            .ReportFileName = gReportPath & "rptMonSTMSITBrkDbNt.rpt"
        Case "Mill Billing Statement" '== "Mill Billing Statement"
             If ValidateData = False Then Exit Sub
             .Formulas(4) = "wFrToDt = 'From : ' & '" & Format(FromDt, "dd/MM/yyyy") & "' & '  To ' & '" & Format(ToDt, "dd/MM/yyyy") & "'"
             '-- MIll Bill detail
             X = "{tblAddLess.vdt} in Date(" & Format(FromDt, "yyyy,mm,dd") & ") to Date(" & Format(ToDt, "yyyy,mm,dd") & ")"
             X = X + " and {tblAddLess.VType}='SM' and {tblMastAccount.AcName}='" & txtMillName & "'"
             .WindowTitle = "CST Bills Details"
            .SelectionFormula = X
            .ReportFileName = gReportPath & "rptMonSTMMillBill.rpt"
            .Action = 1
            '-- Brokerage DbNT
            .Formulas(5) = "BrkP = '" & txt(6) & "'"
             X = "{tblAddLess.vdt} in Date(" & Format(FromDt, "yyyy,mm,dd") & ") to Date(" & Format(ToDt, "yyyy,mm,dd") & ")"
             X = X + " and {tblAddLess.VType}='SM' and {tblMastAccount.AcName}='" & txtMillName & "'"
             .WindowTitle = "Brokerage Debit Note On CST"
            .SelectionFormula = X
            .ReportFileName = gReportPath & "rptMonSTMMillBillTBrkDbNt.rpt"
    End Select
    .Action = 1
End With
End Sub
Private Sub cmdExit_Click()
    DoEvents
    Unload Me
End Sub

Private Sub dtpFrDt_GotFocus()
GProcSelectBox Me.ActiveControl
End Sub

Private Sub dtpToDt_GotFocus()
GProcSelectBox Me.ActiveControl
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
Dim i As Long
Me.Top = 700
Me.Left = Screen.Width / 2 - Me.Width / 2
dtpFrDt.Value = Format(gCYSDate, "dd/MM/yyyy")
dtpToDt.Value = Format(gCYEDate, "dd/MM/yyyy")
dtpBillDt.Value = Format(gCYSDate, "dd/MM/yyyy")
Me.Caption = gSelectedMenu
For i = 0 To 6
    txt(i) = "0.00"
Next i
txt(3) = 0
Select Case gSelectedMenu
    Case "Consignment/ Depot Statement" '--"Consignment/ Depot Statement"
        txt(6).Enabled = False '- Brk %
    Case "Sales In Transit Statement" '== "Sales In Transit Statement"
        dtpBillDt.Enabled = False '-service tax Bill
        For i = 0 To 5 '- %
            txt(i).Enabled = False
        Next i
        txt(2).Enabled = True
    Case "Mill Billing Statement" '== "Mill Billing Statement"
        dtpBillDt.Enabled = False '-service tax Bill
        For i = 0 To 5 '- %
            txt(i).Enabled = False
        Next i
End Select
 '-- Set InActivate Ctrl Back Color
GProcSetInActivateCtrlBackClr Me
End Sub

Private Sub txt_GotFocus(Index As Integer)
GProcSelectBox Me.ActiveControl
End Sub

Private Sub txt_KeyPress(Index As Integer, KeyAscii As Integer)
If Index <= 6 Then
    KeyAscii = GProcNumberOnly(KeyAscii)
End If
End Sub

Private Sub txt_LostFocus(Index As Integer)
ValidateData
End Sub

Private Sub txtMillName_GotFocus()
GProcSelectBox Me.ActiveControl
End Sub

Private Sub txtMillName_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then Exit Sub '-- Enter
Select Case gSelectedMenu
    Case "Consignment/ Depot Statement" '"Consignment/ Depot Statement"
        gClsSearch.SearchMultiField "QryMastaccount", "Account,AcName,ACCode", Array("Mill", "AcName", ""), Array(txtMillName.Width, 0, 0), " (AgCode=90029 or GpCode=90029) and ( ACMILLTYPE =1 or ACMILLTYPE =2)", Chr(KeyAscii), "AcName", txtMillName.Left + Me.Left, txtMillName.Top + Me.Top + 1000, False
    Case "Mill Billing Statement" '"Mill Billing Statement"
        gClsSearch.SearchMultiField "QryMastaccount", "Account,AcName,ACCode", Array("Mill", "AcName", ""), Array(txtMillName.Width, 0, 0), " (AgCode=90029 or GpCode=90029) and ( ACMILLTYPE =3)", Chr(KeyAscii), "AcName", txtMillName.Left + Me.Left, txtMillName.Top + Me.Top + 1000, False
    Case "Sales In Transit Statement" '"Sales In Transit Statement"
        gClsSearch.SearchMultiField "QryMastaccount", "Account,AcName,ACCode", Array("Mill", "AcName", ""), Array(txtMillName.Width, 0, 0), " (AgCode=90029 or GpCode=90029) and ( ACMILLTYPE =4 or ACMILLTYPE =5)", Chr(KeyAscii), "AcName", txtMillName.Left + Me.Left, txtMillName.Top + Me.Top + 1000, False
End Select
If gClsSearch.SearchMultiRetCol(1) <> "" Then
    txtMillName = gClsSearch.SearchMultiRetCol(1)
    KeyAscii = 0
End If
End Sub
'------------ Validate Selection
Private Function ValidateData() As Boolean
Dim i As Long
'-- Mill Name
If txtMillName = "" Then
    MsgBox "Check Mill Name", vbOKOnly + vbInformation, Me.Caption
    txtMillName.SetFocus
    Exit Function
End If
'-- %
For i = 0 To 6
    If IsNumeric(txt(i)) = False Then
        txt(i) = "0.00"
    Else
        txt(i) = Format(txt(i), FStr)
    End If
Next i
txt(3) = Format(txt(3), PStr)
'--- TDS Pay
If GProcGetColumnValue("tblMastAccount", "AcName", txtMillName, "S", " ACISTDSPAY", "N") = 1 Then
     txt(0).Enabled = True
     txt(4).Enabled = True
     txt(5).Enabled = True
Else
     txt(0).Enabled = False
     txt(4).Enabled = False
     txt(5).Enabled = False
     txt(0) = "0.00"
     txt(4) = "0.00"
     txt(5) = "0.00"
End If
ValidateData = True
End Function
