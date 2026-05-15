VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form frmFormEntry 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00C2E3FC&
   BorderStyle     =   0  'None
   Caption         =   "Form Entry"
   ClientHeight    =   7245
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   10515
   Icon            =   "frmFormEntry.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   7245
   ScaleWidth      =   10515
   ShowInTaskbar   =   0   'False
   Begin VB.CheckBox ChkPrint 
      BackColor       =   &H00E7FAFA&
      Caption         =   "Print Report"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   720
      TabIndex        =   36
      Top             =   6600
      Width           =   1935
   End
   Begin VB.TextBox txttemp 
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
      Left            =   1785
      MaxLength       =   40
      TabIndex        =   1
      Text            =   "txttemp(0)"
      ToolTipText     =   "Press F3 for Creation"
      Top             =   525
      Width           =   5040
   End
   Begin VB.ComboBox cboType 
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
      ItemData        =   "frmFormEntry.frx":058A
      Left            =   1785
      List            =   "frmFormEntry.frx":058C
      TabIndex        =   3
      Text            =   "cboType"
      Top             =   975
      Width           =   1560
   End
   Begin VB.TextBox txttemp 
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
      Left            =   1785
      MaxLength       =   40
      TabIndex        =   9
      Text            =   "txttemp(1)"
      ToolTipText     =   "Press F3 for Creation"
      Top             =   1815
      Width           =   5040
   End
   Begin VB.TextBox txttemp 
      Alignment       =   1  'Right Justify
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
      Left            =   6090
      MaxLength       =   40
      TabIndex        =   30
      Text            =   "txttemp(2)"
      Top             =   6615
      Width           =   2040
   End
   Begin VB.TextBox txtGrid 
      BackColor       =   &H8000000A&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   540
      TabIndex        =   11
      Top             =   3675
      Visible         =   0   'False
      Width           =   795
   End
   Begin VB.TextBox txt 
      Height          =   345
      Index           =   17
      Left            =   11190
      TabIndex        =   28
      Text            =   "17"
      Top             =   5940
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.TextBox txt 
      Height          =   345
      Index           =   16
      Left            =   11190
      TabIndex        =   27
      Text            =   "16"
      Top             =   5460
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.TextBox txt 
      Height          =   345
      Index           =   15
      Left            =   11190
      TabIndex        =   26
      Text            =   "15"
      Top             =   4980
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.TextBox txt 
      Height          =   375
      Index           =   12
      Left            =   11100
      TabIndex        =   25
      Text            =   "12"
      Top             =   2865
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.TextBox txt 
      Height          =   325
      Index           =   10
      Left            =   11070
      TabIndex        =   24
      Text            =   "10"
      Top             =   2580
      Visible         =   0   'False
      Width           =   360
   End
   Begin VB.TextBox txt 
      Height          =   375
      Index           =   14
      Left            =   11190
      TabIndex        =   18
      Text            =   "14"
      Top             =   3660
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.TextBox txt 
      Alignment       =   1  'Right Justify
      Height          =   360
      Index           =   13
      Left            =   11070
      TabIndex        =   17
      Text            =   "13"
      Top             =   3300
      Visible         =   0   'False
      Width           =   345
   End
   Begin VB.CommandButton cmdBtn 
      Caption         =   "&Save"
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
      Index           =   13
      Left            =   7395
      Style           =   1  'Graphical
      TabIndex        =   12
      ToolTipText     =   "To Save Record"
      Top             =   825
      Width           =   1095
   End
   Begin VB.CommandButton cmdBtn 
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
      Index           =   10
      Left            =   7395
      Style           =   1  'Graphical
      TabIndex        =   13
      ToolTipText     =   "Exit"
      Top             =   1335
      Width           =   1095
   End
   Begin VB.CommandButton cmdBtn 
      Caption         =   "&Label"
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
      Index           =   11
      Left            =   11070
      Style           =   1  'Graphical
      TabIndex        =   19
      ToolTipText     =   "To Add New Record"
      Top             =   4110
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.CommandButton cmdBtn 
      Caption         =   "Optional"
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
      Index           =   12
      Left            =   11130
      Style           =   1  'Graphical
      TabIndex        =   20
      ToolTipText     =   "To Add New Record"
      Top             =   4545
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.TextBox txt 
      Height          =   345
      Index           =   3
      Left            =   11190
      TabIndex        =   23
      Text            =   "3"
      Top             =   1260
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.TextBox txt 
      Alignment       =   1  'Right Justify
      Height          =   360
      Index           =   5
      Left            =   11190
      TabIndex        =   16
      Text            =   "5"
      Top             =   1980
      Visible         =   0   'False
      Width           =   345
   End
   Begin VB.TextBox txt 
      Height          =   325
      Index           =   7
      Left            =   11070
      TabIndex        =   14
      Text            =   "7"
      Top             =   2340
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.TextBox txt 
      Alignment       =   1  'Right Justify
      Height          =   360
      Index           =   4
      Left            =   11190
      TabIndex        =   15
      Text            =   "4"
      Top             =   1620
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.TextBox txt 
      Height          =   345
      Index           =   2
      Left            =   11190
      TabIndex        =   22
      Text            =   "2"
      Top             =   900
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.TextBox txt 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Index           =   1
      Left            =   11070
      TabIndex        =   21
      Text            =   "1"
      Top             =   540
      Visible         =   0   'False
      Width           =   555
   End
   Begin MSMask.MaskEdBox mskDt 
      Height          =   360
      Index           =   1
      Left            =   5265
      TabIndex        =   7
      Top             =   1395
      Width           =   1560
      _ExtentX        =   2752
      _ExtentY        =   635
      _Version        =   393216
      MaxLength       =   10
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   "dd/mm/yyyy"
      Mask            =   "##/##/####"
      PromptChar      =   "_"
   End
   Begin MSFlexGridLib.MSFlexGrid msGrid 
      Height          =   3795
      Left            =   315
      TabIndex        =   10
      Top             =   2685
      Width           =   9750
      _ExtentX        =   17198
      _ExtentY        =   6694
      _Version        =   393216
      Rows            =   1
      FixedCols       =   0
      RowHeightMin    =   275
      BackColorFixed  =   -2147483647
      ForeColorFixed  =   -2147483634
      BackColorBkg    =   12632256
      Appearance      =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSMask.MaskEdBox mskDt 
      Height          =   360
      Index           =   0
      Left            =   1785
      TabIndex        =   5
      Top             =   1395
      Width           =   1560
      _ExtentX        =   2752
      _ExtentY        =   635
      _Version        =   393216
      MaxLength       =   10
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   "dd/mm/yyyy"
      Mask            =   "##/##/####"
      PromptChar      =   "_"
   End
   Begin VB.Label lbl1 
      BackStyle       =   0  'Transparent
      Caption         =   "TIN  :"
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
      Height          =   375
      Index           =   1
      Left            =   3825
      TabIndex        =   35
      Top             =   960
      Width           =   600
   End
   Begin VB.Label lblPartyTin 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Party's CST TIN"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   330
      Left            =   4560
      TabIndex        =   34
      Top             =   960
      Width           =   2220
   End
   Begin VB.Shape Shape4 
      BorderColor     =   &H00C00000&
      BorderWidth     =   5
      Height          =   7245
      Left            =   -15
      Top             =   -15
      Width           =   10485
   End
   Begin VB.Label lblActionStatus 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "View Record.."
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FFFF&
      Height          =   495
      Left            =   6180
      TabIndex        =   33
      Top             =   15
      Width           =   2475
   End
   Begin VB.Label Lblform 
      BackColor       =   &H00C00000&
      Caption         =   "  Form Entry"
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
      Height          =   345
      Left            =   0
      TabIndex        =   32
      Top             =   30
      Width           =   10455
   End
   Begin VB.Label lbl1 
      BackStyle       =   0  'Transparent
      Caption         =   "Total Bill Amout  :"
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
      Height          =   375
      Index           =   0
      Left            =   4320
      TabIndex        =   31
      Top             =   6615
      Width           =   1740
   End
   Begin VB.Label lblType 
      BackStyle       =   0  'Transparent
      Caption         =   "Sales &Type   : "
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
      Left            =   360
      TabIndex        =   2
      Top             =   975
      Width           =   1320
   End
   Begin VB.Label lbl1 
      BackStyle       =   0  'Transparent
      Caption         =   "Bill Details  :"
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
      Index           =   2
      Left            =   360
      TabIndex        =   29
      Top             =   2415
      Width           =   2175
   End
   Begin VB.Shape Shape3 
      BackColor       =   &H00808080&
      BorderColor     =   &H004D177D&
      FillColor       =   &H00ECFFFF&
      FillStyle       =   0  'Solid
      Height          =   4740
      Left            =   150
      Top             =   2340
      Width           =   10095
   End
   Begin VB.Label lbl1 
      BackStyle       =   0  'Transparent
      Caption         =   "&From Date    :"
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
      Index           =   10
      Left            =   360
      TabIndex        =   4
      Top             =   1395
      Width           =   1455
   End
   Begin VB.Label lbl1 
      BackStyle       =   0  'Transparent
      Caption         =   "To Date :"
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
      Index           =   8
      Left            =   3825
      TabIndex        =   6
      Top             =   1395
      Width           =   1095
   End
   Begin VB.Label lbl1 
      BackStyle       =   0  'Transparent
      Caption         =   "&Mill               :"
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
      Height          =   375
      Index           =   4
      Left            =   360
      TabIndex        =   8
      Top             =   1815
      Width           =   1440
   End
   Begin VB.Label lbl1 
      BackStyle       =   0  'Transparent
      Caption         =   "&Party             :"
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
      Height          =   375
      Index           =   3
      Left            =   360
      TabIndex        =   0
      Top             =   525
      Width           =   1530
   End
   Begin VB.Shape Shape1 
      FillColor       =   &H00BDDFD1&
      FillStyle       =   0  'Solid
      Height          =   1695
      Left            =   7200
      Shape           =   4  'Rounded Rectangle
      Top             =   450
      Width           =   1515
   End
   Begin VB.Shape Shape2 
      BackColor       =   &H00808080&
      BorderColor     =   &H004D177D&
      FillColor       =   &H00ECFFFF&
      FillStyle       =   0  'Solid
      Height          =   1860
      Left            =   150
      Top             =   420
      Width           =   6855
   End
End
Attribute VB_Name = "frmFormEntry"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim ShowFrmFlag As Boolean, rstbl As Recordset
Dim FirstTimeFlag As Boolean
Dim frmTypeStr As String
Dim oldFrmTypeStr As String
Dim tmpFormNo As String
Dim tmpNote As String
Private Sub cboType_Change()
SetCtrlEd
End Sub

Private Sub cboType_KeyPress(KeyAscii As Integer)
KeyAscii = 0
End Sub

Private Sub cboType_LostFocus()
SaveBtnEd
SetCtrlEd
End Sub

Private Sub cmdBtn_Click(Index As Integer)
'On Error GoTo ErrorRoutine
Dim X As String
Select Case Index
Case 10 'Exit
    Unload Me
Case 13 'Save
    If ValidateData = True Then
        If Trim(txtGrid) <> "" Then
            msGrid.Text = txtGrid
            DoEvents
            txtGrid = ""
        End If
        txtGrid.Visible = False
        SaveRelatedRecord
        cmdBtn(13).Enabled = False 'True
        cmdBtn(13).BackColor = vbButtonFace
        
        'Unload Me
    End If
End Select
Exit Sub
ErrorRoutine:
    'rstbl.CancelUpdate
    GProcErrorHandler
End Sub
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case 13:
            SendKeys "{TAB}"
        Case 27:
            If (cmdBtn(13).Enabled = False) Then
                Unload Me
            End If
    End Select
End Sub
Private Sub Form_Load()
    Me.Top = 300
    Me.Left = Screen.Width / 2 - Me.Width / 2
    lblType.Visible = False '-- Type
    cboType.Visible = False
    Select Case gSelectedMenuFormEnt
        Case "CR" '--  c Recd
            Lblform.Caption = "  Form Entry - Form C  Received"
            cboType.AddItem "Other Sale"
            cboType.AddItem "Mill Bill"
            lblType.Visible = True '-- Type
            cboType.Visible = True
        Case "CI" '-- c issue
            Lblform.Caption = "  Form Entry - Form C  Issued"
        Case "E1R" '--  E Recd
            Lblform.Caption = "  Form Entry - Form E1  Received"
        Case "FI" '--  F Issue
            Lblform.Caption = "  Form Entry - Form F  Issued"
        Case "E2R" '--  E2 Recd
            Lblform.Caption = "  Form Entry - Form E2  Received"
        Case "E2I" '--  E2 Issued
            Lblform.Caption = "  Form Entry - Form E2  Issued"
    End Select
    SetCtrlEd
    txttemp(2).Locked = True
    '---Grid
    FirstTimeFlag = True
    SetGrid
    Call GProcClearForm(Me, rstbl, 0, False)
    If FirstTimeFlag = True Then
        FirstTimeFlag = False
    Else
        txttemp(0).SetFocus
    End If
    ChkPrint.Value = 1
End Sub
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
'If gFormAction = vbDataActionUpdate Then
'    MsgBox "You are about to modify the record. Save the record and then exit.", vbInformation + vbOKOnly, Me.Caption
'    Cancel = True
'    Exit Sub
'End If
If rstbl.State = adStateOpen Then rstbl.Close
End Sub

Private Sub mskDt_LostFocus(Index As Integer)
SaveBtnEd
If gSelectedMenuFormEnt = "CI" Or gSelectedMenuFormEnt = "FI" Then
    If IsDate(mskDt(Index)) = False Then mskDt(Index) = gCYSDate
Else
    If GProcIsDateValid(mskDt(Index)) = False Then mskDt(Index) = gCYSDate
End If
If Index = 1 Then FillRelatedRecord
End Sub
Private Sub mskDt_GotFocus(Index As Integer)
GProcSelectBox Me.ActiveControl
End Sub
Private Sub SaveBtnEd()
'--Party,Frmdt,To dt
If txttemp(0) = "" Or IsDate(mskDt(0)) = False Or IsDate(mskDt(1)) = False Then
    cmdBtn(13).Enabled = False
    cmdBtn(13).BackColor = &H80000004
Else
    cmdBtn(13).Enabled = True
    cmdBtn(13).BackColor = &HFFFF80
End If
End Sub

Private Function ValidateData() As Boolean
Dim X As String
Dim GName As String
Dim rstAddCheck As Recordset
Dim MillAgCode As Long
'---- Required Data
'--- Validate Grid
'If ValidateGrid = False Then Exit Function
If gSelectedMenuFormEnt = "FI" Or gSelectedMenuFormEnt = "CI" Then
    '--- From dt
    If IsDate(mskDt(0)) = False Then
        MsgBox "Check From Date.", vbInformation + vbOKOnly, Me.Caption
        mskDt(0).SetFocus
        Exit Function
    End If
    '---To dt
    If IsDate(mskDt(1)) = False Then
        MsgBox "Check to Date.", vbInformation + vbOKOnly, Me.Caption
        mskDt(1).SetFocus
        Exit Function
    End If
Else
    '--- From dt
    If GProcIsDateValid(mskDt(0)) = False Then
        MsgBox "Check From Date.", vbInformation + vbOKOnly, Me.Caption
        mskDt(0).SetFocus
        Exit Function
    End If
    '---To dt
    If GProcIsDateValid(mskDt(1)) = False Then
        MsgBox "Check to Date.", vbInformation + vbOKOnly, Me.Caption
        mskDt(1).SetFocus
        Exit Function
    End If
End If

'---- Party
If txttemp(0) = "" Then
    MsgBox "Check Account.", vbInformation + vbOKOnly, Me.Caption
    txttemp(0).SetFocus
    Exit Function
End If
'--- Mill
'If txttemp(1) = "" Then
'    MsgBox "Check Mill.", vbInformation + vbOKOnly, Me.Caption
'    txttemp(1).SetFocus
'    Exit Function
'End If
ValidateData = True
End Function

Private Sub txtGrid_LostFocus()
If LCase(Me.ActiveControl.Name) <> "msgrid" Then
    txtGrid.Visible = False
    msGrid = txtGrid
    CalGridAmt
End If
End Sub

Private Sub txttemp_GotFocus(Index As Integer)
GProcSelectBox Me.ActiveControl
End Sub
Private Sub txttemp_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
Dim GName As String
If KeyCode = vbKeyF3 Then
    Select Case Index
        Case 0 '-- Party
            GName = GProcGetColumnValue("TblMastGroup", "AgCode", 90017, "N", "AgName", "S")
            GProcShowForm frmMastAccount, frmMain.mnuMstfrm(0), Array(frmMastAccount.txttemp(2)), Array(GName)
        Case 1 '-- Mill
            GName = GProcGetColumnValue("TblMastGroup", "AgCode", 90029, "N", "AgName", "S")
            GProcShowForm frmMastAccount, frmMain.mnuMstfrm(0), Array(frmMastAccount.txttemp(2)), Array(GName)
    End Select
End If
End Sub
Private Sub txtTemp_KeyPress(Index As Integer, KeyAscii As Integer)
Dim LeftPos As Long
Dim TopPos As Long
TopPos = txttemp(Index).Top + Me.Top + 650
LeftPos = txttemp(Index).Left + Me.Left
If KeyAscii = 13 Then Exit Sub
Select Case Index
    Case 0 '--Party
        gClsSearch.SearchMultiField "QryMastaccount", "Account,AcName,AcCity,AgName", Array("Account", "AcName", "City", "Group"), Array(3000, 0, 2000, 3000), "(AgCode=90017 or AgCode=90015 or AgCode=90029)", Chr(KeyAscii), "AcName", LeftPos, TopPos
        KeyAscii = 0
        If gClsSearch.SearchMultiRetCol(1) <> "" Then txttemp(Index) = gClsSearch.SearchMultiRetCol(1)
    Case 1 '--Mill
        gClsSearch.SearchMultiField "QryMastaccount", "Account,AcName,AcCity,AgName", Array("Account", "AcName", "City", "Group"), Array(3000, 0, 2000, 3000), "AgCode=90029 ", Chr(KeyAscii), "AcName", LeftPos, TopPos
        KeyAscii = 0
        If gClsSearch.SearchMultiRetCol(1) <> "" Then txttemp(Index) = gClsSearch.SearchMultiRetCol(1)
End Select
End Sub

Private Sub txtTemp_LostFocus(Index As Integer)
lblPartyTin.Caption = GProcGetColumnValue("tblMastAccount", "AcName", txttemp(0), "S", "AcCST", "S")
SaveBtnEd
If Index = 1 Then FillRelatedRecord
End Sub
'--- Related Record
Private Function IsRelatedRecord() As Boolean
'-- Audited
'If GProcISAudited(txt(0), txt(3), txt(17)) = True Then
'    MsgBox "Record is Audited.You can not Modify/Delete this Record.", vbCritical + vbOKOnly, App.Title
'    IsRelatedRecord = True
'    Exit Function
'End If
End Function
Private Sub SetCtrlEd()
'-- Mill Name
If cboType = "Other Sale" Then
    txttemp(1).Enabled = False
    txttemp(1) = ""
Else
    txttemp(1).Enabled = True
End If
End Sub
Private Sub SetGrid()
Dim X As String
Dim MillCode As Long
Dim PartyCode As Long
Dim CondStr As String
Dim DtStr As String
Dim FormType As String
tmpFormNo = ""
tmpNote = ""
Dim tmpFormRecdDt As String
tmpFormRecdDt = "__/__/____"
PartyCode = GProcGetColumnValue("tblmastAccount", "AcName", txttemp(0), "S", "AcCode", "N")
Select Case gSelectedMenuFormEnt
    Case "CR" '----- C Form Rec
        If cboType = "Mill Bill" Then
            CondStr = " vtype ='SM' and instr(narration,'C Form')>0"
            MillCode = GProcGetColumnValue("tblmastAccount", "AcName", txttemp(1), "S", "AcCode", "N")
            CondStr = CondStr + " and  AdMillCode=" & MillCode & " and ADACDRCODE=" & PartyCode
        Else
            CondStr = " Vtype <> 'SM' and (instr(narration,'C Form')>0 or instr(narration,'C S T')>0 )  And ADACDRCODE = " & PartyCode
        End If
        If IsDate(mskDt(0)) = True And IsDate(mskDt(1)) = True Then
            If gBackEndDB = gBackEndAccess Then
                DtStr = " vdt>=cdate('" & mskDt(0) & "') and Vdt<=cdate('" & mskDt(1) & "')"
            Else
                DtStr = " vdt>=to_date('" & CDate(mskDt(0)) & "','dd/MM/yyyy') and Vdt<=To_date('" & CDate(mskDt(1)) & "','dd/MM/yyyy')"
            End If
        Else
            DtStr = " vno=0"
        End If
        '--Vtype,AdInvNo,Vdt,AdBillAmt, ADCFORMNO,Vno,VYear
        X = "select Vtype,AdInvNo,Vdt,AdBillAmt, ADCFORMNO,AdFormRecdDate,AdTmp8,Vno,VYear from tblAddLess,tblmastNarration where NarrCode=ADTAXCODE and " _
        & DtStr & " and " & CondStr & " order by Vdt,AdInvNo"
        
'        X = "select Vtype,AdInvNo,Vdt,AdBillAmt, ADCFORMNO,AdFormRecdDate,Vno,VYear from tblAddLess,tblmastNarration where NarrCode=ADTAXCODE and " _
'        & DtStr & " and " & CondStr & " order by Vdt,AdInvNo"
    Case "E2I" '----- C Form Rec
         CondStr = " Vtype = 'ST' and instr(narration,'C Form')>0  And ADACDRCODE = " & PartyCode
        If IsDate(mskDt(0)) = True And IsDate(mskDt(1)) = True Then
            If gBackEndDB = gBackEndAccess Then
                DtStr = " vdt>=cdate('" & mskDt(0) & "') and Vdt<=cdate('" & mskDt(1) & "')"
            Else
                DtStr = " vdt>=to_date('" & CDate(mskDt(0)) & "','dd/MM/yyyy') and Vdt<=To_date('" & CDate(mskDt(1)) & "','dd/MM/yyyy')"
            End If
        Else
            DtStr = " vno=0"
        End If
        '--Vtype,AdInvNo,Vdt,AdBillAmt, ADCFORMNO,Vno,VYear
        X = "select Vtype,AdInvNo,Vdt,AdBillAmt, ADE1FORMNO,AdFormE1RecdDate,Adtmp8,Vno,VYear from tblAddLess,tblmastNarration where NarrCode=ADTAXCODE and " _
        & DtStr & " and tblAddless.ADTMP2 = 1 and " & CondStr & " order by Vdt,AdInvNo"
'        X = "select Vtype,AdInvNo,Vdt,AdBillAmt, ADE1FORMNO,AdFormE1RecdDate,Vno,VYear from tblAddLess,tblmastNarration where NarrCode=ADTAXCODE and " _
'        & DtStr & " and tblAddless.ADTMP2 = 1 and " & CondStr & " order by Vdt,AdInvNo"
    Case "CI", "E1R", "E2R" '----- C Form Issue ,E1 Receive
        If gSelectedMenuFormEnt = "CI" Then '----- C Form Issue
            CondStr = " vtype in ('PY','PO','PT') and instr(narration,'C Form')>0"
        Else '-- E1 Receive
            CondStr = " vtype in ('PT') and instr(narration,'C Form')>0"
        End If
        MillCode = GProcGetColumnValue("tblmastAccount", "AcName", txttemp(1), "S", "AcCode", "N")
        CondStr = CondStr + " and  AdMillCode=" & MillCode & " and ADACCRCODE=" & PartyCode
        If IsDate(mskDt(0)) = True And IsDate(mskDt(1)) = True Then
            If gBackEndDB = gBackEndAccess Then
                DtStr = " adGpDt>=cdate('" & mskDt(0) & "') and adGpDt<=cdate('" & mskDt(1) & "')"
            Else
                DtStr = " adGpDt>=to_date('" & CDate(mskDt(0)) & "','dd/MM/yyyy') and adGpDt<=To_date('" & CDate(mskDt(1)) & "','dd/MM/yyyy')"
            End If
        Else
            DtStr = " vno=0"
        End If
        '--Vtype,AdInvNo,Vdt,AdBillAmt, ADCFORMNO,Vno,VYear
        Select Case gSelectedMenuFormEnt
        Case "CI"
            X = "select Vtype,AdInvNo,adGpDt,AdBillAmt, ADCFORMNO,AdFormRecdDate,Adtmp8,Vno,VYear from tblAddLess,tblmastNarration where NarrCode=ADTAXCODE and " _
            & DtStr & " and " & CondStr & " order by AdGpdt,AdInvNo"
'            X = "select Vtype,AdInvNo,adGpDt,AdBillAmt, ADCFORMNO,AdFormRecdDate,Vno,VYear from tblAddLess,tblmastNarration where NarrCode=ADTAXCODE and " _
'            & DtStr & " and " & CondStr & " order by AdGpdt,AdInvNo"
        Case "E1R"
            X = "select Vtype,AdInvNo,adGpDt,AdBillAmt, ADE1FORMNO,AdFormE1RecdDate,Adtmp8,Vno,VYear from tblAddLess,tblmastNarration where NarrCode=ADTAXCODE and " _
            & DtStr & " and tblAddless.AdgpNo <> 1 and " & CondStr & " order by AdGpdt,AdInvNo"
'            X = "select Vtype,AdInvNo,adGpDt,AdBillAmt, ADE1FORMNO,AdFormE1RecdDate,Vno,VYear from tblAddLess,tblmastNarration where NarrCode=ADTAXCODE and " _
'            & DtStr & " and " & CondStr & " order by AdGpdt,AdInvNo"
        Case "E2R"
            X = "select Vtype,AdInvNo,adGpDt,AdBillAmt, ADE1FORMNO,AdFormE1RecdDate,Adtmp8,Vno,VYear from tblAddLess,tblmastNarration where NarrCode=ADTAXCODE and " _
            & DtStr & " and tblAddless.AdgpNo = 1 and " & CondStr & " order by AdGpdt,AdInvNo"
        End Select
    Case "FI" '-- F Form Issue
        CondStr = " vtype in ('PI') "
        MillCode = GProcGetColumnValue("tblmastAccount", "AcName", txttemp(1), "S", "AcCode", "N")
        CondStr = CondStr + " and  AdMillCode=" & MillCode & " and ADACCRCODE=" & PartyCode
        If IsDate(mskDt(0)) = True And IsDate(mskDt(1)) = True Then
            If gBackEndDB = gBackEndAccess Then
                DtStr = " adGpDt>=cdate('" & mskDt(0) & "') and adGpDt<=cdate('" & mskDt(1) & "')"
            Else
                DtStr = " adGpDt>=to_date('" & CDate(mskDt(0)) & "','dd/MM/yyyy') and adGpDt<=To_date('" & CDate(mskDt(1)) & "','dd/MM/yyyy')"
            End If
        Else
            DtStr = " vno=0"
        End If
        '--Vtype,AdInvNo,Vdt,AdBillAmt, ADCFORMNO,Vno,VYear
        X = "select Vtype,AdGPNo,AdGpdt,AdBillAmt, ADCFORMNO,AdFormRecdDate,Adtmp8,Vno,VYear from tblAddLess where " _
        & DtStr & " and " & CondStr & " order by AdGpdt,AdGpNo"
End Select
Set rstbl = New Recordset
GProcRstOpen rstbl, X, "R"
With msGrid
    .Clear
    .FormatString = "<Type       |>            Bill No|< Date                 |>             Bill Amount|<Form No                      |<Date              |<Form Sr.No                                                                 |>Vno|>Vyear"
    .ColWidth(5) = 0
    .ColWidth(7) = 0
    .ColWidth(8) = 0
    .rows = 2
    'FillRelatedRecord
End With
End Sub
Private Sub msGrid_Click()
'On Error Resume Next
    msGrid_EnterCell
End Sub
Private Sub msGrid_GotFocus()
If LCase(Me.ActiveControl) <> "txtgrid" Then msGrid.Col = 4
   ' cmdBtn(10).Cancel = False
End Sub
Private Sub MSGrid_LostFocus()
'CalAmount
End Sub
Private Sub msGrid_EnterCell()
'If FormAction <> vbDataActionClose Then
    CheckForNumberTxtGrid
    DoEvents
    txtGrid = ""
    DoEvents
    If msGrid.TextMatrix(msGrid.Row, 4) = "" Then msGrid.TextMatrix(msGrid.Row, 4) = tmpFormNo  ''-Form No
    If msGrid.TextMatrix(msGrid.Row, 6) = "" Then msGrid.TextMatrix(msGrid.Row, 6) = tmpNote  ''-Form No
    If msGrid.Text <> "" Then txtGrid.Text = msGrid.Text
    '--- Lock TextBox
    With msGrid
        Select Case .Col
            Case 0, 1, 2, 3 '--Type,Bill No,Date,Amt
                txtGrid.Visible = True
                txtGrid.Locked = True
            Case Else
                txtGrid.Locked = False
        End Select
    End With
    msGrid_KeyPress (0)
'End If
End Sub
Private Sub msGrid_KeyPress(KeyAscii As Integer)
With msGrid
    If KeyAscii = 27 Then '----Esc Key
        KeyAscii = 0
        txtGrid.Visible = False
        Exit Sub
    End If
    '---- Navigate Row ,Col
    If KeyAscii = 13 Then
        KeyAscii = 0
        If .Col = 6 Then
            If .Row = (.rows - 2) Then '-- Add New Row
            Else '--Next Row
                .Row = .Row + 1
            End If
            .Col = 5
        Else    '-- Next col
            .Col = .Col + 1
        End If
    End If
'    If KeyAscii <> 13 And KeyAscii <> 27 And .Col = 0 Then
'        DoEvents
'        KeyAscii = 0
'        Pending_In
'    End If
End With
'ProcSetAlignment rstbl, msGrid.Col
'---- Set TxtGrid
If msGrid.Text <> "" Then txtGrid.Text = msGrid.Text
If msGrid.Col < 7 Then
        txtGrid.Move msGrid.CellLeft + msGrid.Left, msGrid.CellTop + msGrid.Top, msGrid.CellWidth, msGrid.CellHeight
End If
txtGrid.Visible = True
txtGrid.Enabled = True
DoEvents
If KeyAscii <> 0 Then
    Select Case rstbl.Fields(msGrid.Col).Type
        Case adChar, adVarWChar
            txtGrid.Text = txtGrid.Text & Chr(KeyAscii)
        Case Else
            txtGrid.Text = Chr(GProcValidateKey(rstbl, msGrid.Col, KeyAscii, txtGrid.Text)) 'Chr(GProcNumberWithDecimal(KeyAscii, txtGrid))
    End Select
End If
DoEvents
If msGrid.Text <> "" Then
    txtGrid.SelStart = Len(txtGrid.Text) + 1
Else
    txtGrid.SelStart = 1
End If
DoEvents
If txtGrid.Visible = True Then txtGrid.SetFocus
End Sub
Private Sub msGrid_LeaveCell()
If txtGrid.Visible Then
    msGrid.Text = txtGrid.Text
    DoEvents
    txtGrid.Text = ""
    txtGrid.Visible = False
    tmpFormNo = msGrid.TextMatrix(msGrid.Row, 4)
    tmpNote = msGrid.TextMatrix(msGrid.Row, 6)
End If
If msGrid.Col = 4 Then CalGridAmt
End Sub
Private Sub CalGridAmt()
Dim RowIndex As Long
Dim Amt As Double
RowIndex = 1
With msGrid
    If txtGrid.Visible = True Then .TextMatrix(.Row, .Col) = txtGrid
    Do While RowIndex < .rows
        '-- Bill amt with Form No
        
'---For Fill C Form nos
        If .TextMatrix(RowIndex, 4) = "" And RowIndex > 1 Then
           .TextMatrix(RowIndex, 4) = .TextMatrix(RowIndex - 1, 4)
        End If

'---For Fill C Form Rec Date
'        If .TextMatrix(RowIndex, 5) = "" And RowIndex > 1 Then
'           .TextMatrix(RowIndex, 5) = .TextMatrix(RowIndex - 1, 5)
'        End If
'---
'---For Fill C Form Note
        If .TextMatrix(RowIndex, 6) = "" And RowIndex > 1 Then
           .TextMatrix(RowIndex, 6) = .TextMatrix(RowIndex - 1, 6)
        End If
'---
        
        If .TextMatrix(RowIndex, 3) <> "" Then
            Amt = Amt + CDbl(.TextMatrix(RowIndex, 3))
        End If
        RowIndex = RowIndex + 1
    Loop
End With
txttemp(2) = Format(Amt, FStr)
End Sub
Private Sub txtGrid_GotFocus()
    txtGrid.Alignment = ProcSetAlignment(rstbl, msGrid.Col)
    DoEvents
    Call GProcSelectBox(txtGrid)
End Sub
Private Sub txtGrid_KeyPress(KeyAscii As Integer)
Dim LeftPos As Long
Dim TopPos As Long
Dim PartyCode As Long
TopPos = txtGrid.Top + Me.Top + 650
LeftPos = txtGrid.Left + Me.Left
DoEvents
With msGrid
    If KeyAscii <> 13 And KeyAscii <> 27 Then
        '--- Search
        Select Case .Col
            Case 0, 1, 2, 3
                KeyAscii = 0
            Case Else
                KeyAscii = GProcValidateKey(rstbl, msGrid.Col, KeyAscii, txtGrid)
        End Select
   End If
    If KeyAscii = 13 Then
        CheckForNumberTxtGrid
    End If
    '--- Navigate Row ,Col
    If KeyAscii = 13 And .Col < 6 Then
        .Col = .Col + 1
        If .Col = 4 Then tmpFormNo = txtGrid
    
    ElseIf KeyAscii = 13 And .Col = 6 Then
        If .Row = (.rows - 1) Then
           ChkPrint.SetFocus
        Else
            .Row = .Row + 1
        End If
        If .Col = 6 Then tmpNote = txtGrid

'        tmpFormNo = txtGrid
        .Col = 6
  '      .Text = tmpFormNo
    End If
    '-- On ESC
    With msGrid
        If KeyAscii = 27 Then '----Esc Key
            KeyAscii = 0
            txtGrid.Visible = False
        End If
    End With
End With
End Sub
Private Sub CheckForNumberGrid()
Dim RowIndex As Long
Dim i As Long
With rstbl
For RowIndex = 1 To msGrid.rows - 1
    For i = 1 To .Fields.Count - 1
        Select Case .Fields(i).Type
            Case adInteger, 2, 3 '-- Integer
                If IsNumeric(msGrid.TextMatrix(RowIndex, i)) = False Then
                    msGrid.TextMatrix(RowIndex, i) = 0
                Else
                    msGrid.TextMatrix(RowIndex, i) = CLng(msGrid.TextMatrix(RowIndex, i))
                End If
            Case adDecimal, 131, adNumeric, adDouble '-- Double
                If IsNumeric(msGrid.TextMatrix(RowIndex, i)) = False Then
                    msGrid.TextMatrix(RowIndex, i) = Format(0, GProcNumberFormat(.Fields(i).Precision, .Fields(i).NumericScale))
                Else
                    msGrid.TextMatrix(RowIndex, i) = Format(msGrid.TextMatrix(RowIndex, i), GProcNumberFormat(.Fields(i).Precision, .Fields(i).NumericScale))
                End If
            Case adChar, adVarWChar
            Case adDate '-- Date
                If IsDate(msGrid.TextMatrix(RowIndex, i)) = False Then
                    msGrid.TextMatrix(RowIndex, i) = Null & vbNullString
                End If
        End Select
    Next i
  Next RowIndex
End With
End Sub
'---  Check For Number Grid
Private Sub CheckForNumberTxtGrid()
With rstbl
    Select Case .Fields(msGrid.Col).Type
        Case adInteger, 2, 3 '-- Integer
            If IsNumeric(txtGrid) = False Then
                txtGrid = 0
            Else
                txtGrid = CLng(txtGrid)
            End If
            txtGrid.MaxLength = .Fields(msGrid.Col).Precision
        Case adDecimal, 131, adNumeric, adDouble '-- Double
            If IsNumeric(txtGrid) = False Then
                txtGrid = Format(0, GProcNumberFormat(.Fields(msGrid.Col).Precision, .Fields(msGrid.Col).NumericScale))
            Else
                txtGrid = Format(txtGrid, GProcNumberFormat(.Fields(msGrid.Col).Precision, .Fields(msGrid.Col).NumericScale))
            End If
            txtGrid.MaxLength = .Fields(msGrid.Col).Precision
        Case adDate '-- Date
            If IsDate(txtGrid) = False Then
                txtGrid = Null & vbNullString
            Else
                txtGrid = CDate(txtGrid)
            End If
            txtGrid.MaxLength = 10
        Case Else
            txtGrid.MaxLength = .Fields(msGrid.Col).DefinedSize
    End Select
End With
End Sub
Private Function ProcSetAlignment(pRstbl As Recordset, pColIndex As Long)
Select Case pRstbl.Fields(pColIndex).Type
    Case adInteger, adDouble, 2, 3, 131, adDecimal
        ProcSetAlignment = 1
    Case Else
        ProcSetAlignment = 0
End Select
End Function
'-- Grid Required
Private Function ValidateGrid() As Boolean
End Function
Private Sub SaveRelatedRecord()
Dim i As Long, X As String
Dim Vno As Long
Dim Vtype As String
Dim VYear As String
Dim FormNo As String
Dim FormDate As String
Dim Note As String
'Dim Rstbl1 As Recordset
'--- Save Grid Records
CheckForNumberGrid
With msGrid
    i = 1
    Do While i < .rows
        Vno = CLng(.TextMatrix(i, 7)) '--Vno
        Vtype = .TextMatrix(i, 0) '--VType
        VYear = CLng(.TextMatrix(i, 8)) '--VYear
        FormNo = .TextMatrix(i, 4) '--FromNo
        If IsDate(.TextMatrix(i, 5)) = True Then
           FormDate = CDate(.TextMatrix(i, 5)) '--From date
        Else
          ' FormDate = "__/__/____"  '--From date
           FormDate = ""
        End If
        Note = .TextMatrix(i, 6) '--Note
        
        If gSelectedMenuFormEnt = "E1R" Or gSelectedMenuFormEnt = "E2R" Or gSelectedMenuFormEnt = "E2I" Then '-- E 1 From Receive
            gCn.Execute "update tbladdless set ADE1FORMNO='" & FormNo & "' where vno=" & Vno & " and Vtype='" & Vtype & "' and Vyear=" & VYear
            If FormDate <> "" Then gCn.Execute "update tbladdless set AdFormE1RecdDate= '" & Format(FormDate, "yyyy/mm/dd") & " 'where vno=" & Vno & " and Vtype='" & Vtype & "' and Vyear=" & VYear
            gCn.Execute "update tbladdless set AdTmp8='" & Note & "' where vno=" & Vno & " and Vtype='" & Vtype & "' and Vyear=" & VYear
        Else
            gCn.Execute "update tbladdless set ADCFORMNO='" & FormNo & "' where vno=" & Vno & " and Vtype='" & Vtype & "' and Vyear=" & VYear
            'gCn.Execute "update tbladdless set AdFormRecdDate= " & FormDate & " where vno=" & Vno & " and Vtype='" & Vtype & "' and Vyear=" & VYear
            If FormDate <> "" Then gCn.Execute "update tbladdless set AdFormRecdDate= '" & Format(FormDate, "yyyy/mm/dd") & " 'where vno=" & Vno & " and Vtype='" & Vtype & "' and Vyear=" & VYear
            gCn.Execute "update tbladdless set AdTmp8='" & Note & "' where vno=" & Vno & " and Vtype='" & Vtype & "' and Vyear=" & VYear
        End If
        i = i + 1
     Loop
End With
If ChkPrint.Value = 1 Then
    Select Case gSelectedMenuFormEnt
        Case "CR"  '--"C Form Receivable"
            gSelectedMenu = "C Form Receivable"
        Case "CI" '-- "C Form Issuable"
            gSelectedMenu = "C Form Issuable"
        Case "E1R"  '-- "E1 From Receivable"
            gSelectedMenu = "E1 Form Receivable"
        Case "FI"  '-- "F Form Issuable"
            gSelectedMenu = "F Form Issuable"
        Case "E2R"  '-- "E2 From Receivable"
            gSelectedMenu = "E2 Form Receivable"
        Case "E2I"  '-- "E2 From Issuable"
            gSelectedMenu = "E2 From Issuable"
    End Select
    If gSelectedMenu <> "" Then
       frmRptTaxReports.Opt(2).Value = True
       frmRptTaxReports.txtParty = txttemp(0)
       frmRptTaxReports.ChkConsolidate.Value = 0
       frmRptTaxReports.dtpFrDt.Value = Format(mskDt(0), "dd/MM/yyyy")
       frmRptTaxReports.dtpToDt.Value = Format(mskDt(1), "dd/MM/yyyy")
       If gSelectedMenu = "E1 Form Receivable" Or gSelectedMenu = "C Form Issuable" Then
          frmRptTaxReports.OptParty.Value = 0
       Else
          frmRptTaxReports.OptParty.Value = 1
       End If
       frmRptTaxReports.Show
       frmRptTaxReports.SetFocus
'       frmRptTaxReports.cmdCrystal(1).Value = True
 '      Unload frmRptTaxReports
       Exit Sub
    End If
    frmFormEntry.Visible = True
End If
ChkPrint.Value = 0
End Sub
Private Sub FillRelatedRecord()
Dim RowIndex As Long
Dim ColIndex As Long
tmpFormNo = ""
tmpNote = ""
txtGrid.Visible = False
'--Fill Grid
SetGrid
With rstbl
    If Not .EOF Then
        .MoveFirst
        RowIndex = 1
        msGrid.rows = 2
        Do While Not .EOF
            For ColIndex = 0 To .Fields.Count - 1
                Select Case .Fields(ColIndex).Type
                    Case adInteger, 2, 3 '-- Integer
                        msGrid.TextMatrix(RowIndex, ColIndex) = CLng(.Fields(ColIndex))
                    Case adDecimal, 131, adNumeric, adDouble  '-- Double
                        msGrid.TextMatrix(RowIndex, ColIndex) = Format(.Fields(ColIndex), GProcNumberFormat(.Fields(ColIndex).Precision, .Fields(ColIndex).NumericScale))
                    Case Else
                        msGrid.TextMatrix(RowIndex, ColIndex) = .Fields(ColIndex) & vbNullString
                End Select
            Next ColIndex
            RowIndex = RowIndex + 1
            .MoveNext
            If Not .EOF Then msGrid.rows = msGrid.rows + 1
        Loop
    End If
End With
CalGridAmt
End Sub
