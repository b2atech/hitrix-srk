VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form frmFormEntry 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00C2E3FC&
   BorderStyle     =   0  'None
   Caption         =   "Form Entry"
   ClientHeight    =   8370
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   9390
   Icon            =   "frmFormEntry.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   8370
   ScaleWidth      =   9390
   ShowInTaskbar   =   0   'False
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
      Index           =   3
      Left            =   3180
      MaxLength       =   40
      TabIndex        =   37
      Text            =   "txttemp(3)"
      ToolTipText     =   "Press F3 for Creation"
      Top             =   7380
      Visible         =   0   'False
      Width           =   5040
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
      Left            =   4770
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
      Left            =   12690
      TabIndex        =   28
      Text            =   "17"
      Top             =   6000
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.TextBox txt 
      Height          =   345
      Index           =   16
      Left            =   12690
      TabIndex        =   27
      Text            =   "16"
      Top             =   5520
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.TextBox txt 
      Height          =   345
      Index           =   15
      Left            =   12690
      TabIndex        =   26
      Text            =   "15"
      Top             =   5040
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.TextBox txt 
      Height          =   375
      Index           =   12
      Left            =   12600
      TabIndex        =   25
      Text            =   "12"
      Top             =   2925
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.TextBox txt 
      Height          =   325
      Index           =   10
      Left            =   12570
      TabIndex        =   24
      Text            =   "10"
      Top             =   2640
      Visible         =   0   'False
      Width           =   360
   End
   Begin VB.TextBox txt 
      Height          =   375
      Index           =   14
      Left            =   12690
      TabIndex        =   18
      Text            =   "14"
      Top             =   3720
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.TextBox txt 
      Alignment       =   1  'Right Justify
      Height          =   360
      Index           =   13
      Left            =   12570
      TabIndex        =   17
      Text            =   "13"
      Top             =   3360
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
      Left            =   7635
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
      Height          =   405
      Index           =   10
      Left            =   7635
      Style           =   1  'Graphical
      TabIndex        =   13
      ToolTipText     =   "Exit"
      Top             =   1455
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
      Left            =   12570
      Style           =   1  'Graphical
      TabIndex        =   19
      ToolTipText     =   "To Add New Record"
      Top             =   4170
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
      Left            =   12630
      Style           =   1  'Graphical
      TabIndex        =   20
      ToolTipText     =   "To Add New Record"
      Top             =   4605
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.TextBox txt 
      Height          =   345
      Index           =   3
      Left            =   12690
      TabIndex        =   23
      Text            =   "3"
      Top             =   1320
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.TextBox txt 
      Alignment       =   1  'Right Justify
      Height          =   360
      Index           =   5
      Left            =   12690
      TabIndex        =   16
      Text            =   "5"
      Top             =   2040
      Visible         =   0   'False
      Width           =   345
   End
   Begin VB.TextBox txt 
      Height          =   325
      Index           =   7
      Left            =   12570
      TabIndex        =   14
      Text            =   "7"
      Top             =   2400
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.TextBox txt 
      Alignment       =   1  'Right Justify
      Height          =   360
      Index           =   4
      Left            =   12690
      TabIndex        =   15
      Text            =   "4"
      Top             =   1680
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.TextBox txt 
      Height          =   345
      Index           =   2
      Left            =   12690
      TabIndex        =   22
      Text            =   "2"
      Top             =   960
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
      Left            =   12570
      TabIndex        =   21
      Text            =   "1"
      Top             =   600
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
      Width           =   8670
      _ExtentX        =   15293
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
      Caption         =   "C Form Issued To Broker   :"
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
      Index           =   5
      Left            =   420
      TabIndex        =   36
      Top             =   7380
      Visible         =   0   'False
      Width           =   2820
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
      Left            =   4335
      TabIndex        =   35
      Top             =   960
      Width           =   2220
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
      Left            =   3600
      TabIndex        =   34
      Top             =   960
      Width           =   600
   End
   Begin VB.Shape Shape4 
      BorderColor     =   &H00C00000&
      BorderWidth     =   5
      Height          =   8205
      Left            =   -15
      Top             =   -15
      Width           =   9345
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
      Left            =   6720
      TabIndex        =   33
      Top             =   0
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
      Width           =   9255
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
      Left            =   3000
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
      Height          =   5700
      Left            =   150
      Top             =   2340
      Width           =   9015
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
      Height          =   1815
      Left            =   7440
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
Dim ShowFrmFlag As Boolean, Rstbl As Recordset
Dim FirstTimeFlag As Boolean
Dim frmTypeStr As String
Dim oldFrmTypeStr As String
Dim tmpFormNo As String
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
On Error GoTo ErrorRoutine
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
    Select Case gSelectedMenu
        Case "CR" '--  c Recd
            Lblform.Caption = "  Form Entry - Form C  Received"
            cboType.AddItem "Other Sale"
            cboType.AddItem "Mill Bill"
            lblType.Visible = True '-- Type
            cboType.Visible = True
        Case "CI" '-- c issue
            Lblform.Caption = "  Form Entry - Form C  Issued"
            lbl1(5).Visible = True
            txtTemp(3).Visible = True
        Case "E1R" '--  E Recd
            Lblform.Caption = "  Form Entry - Form E1  Received"
        Case "FI" '--  F Issue
            Lblform.Caption = "  Form Entry - Form F  Issued"
    End Select
    SetCtrlEd
    txtTemp(2).Locked = True
    '---Grid
    FirstTimeFlag = True
    SetGrid
    Call GProcClearForm(Me, Rstbl, 0, False)
    If FirstTimeFlag = True Then
        FirstTimeFlag = False
    Else
        txtTemp(0).SetFocus
    End If
End Sub
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
'If gFormAction = vbDataActionUpdate Then
'    MsgBox "You are about to modify the record. Save the record and then exit.", vbInformation + vbOKOnly, Me.Caption
'    Cancel = True
'    Exit Sub
'End If
If Rstbl.State = adStateOpen Then Rstbl.Close
End Sub

Private Sub mskDt_LostFocus(Index As Integer)
SaveBtnEd
If gSelectedMenu = "CI" Or gSelectedMenu = "FI" Then
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
If txtTemp(0) = "" Or IsDate(mskDt(0)) = False Or IsDate(mskDt(1)) = False Then
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
If gSelectedMenu = "FI" Or gSelectedMenu = "CI" Then
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
If txtTemp(0) = "" Then
    MsgBox "Check Account.", vbInformation + vbOKOnly, Me.Caption
    txtTemp(0).SetFocus
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
            GProcShowForm frmMastAccount, frmMain.mnuMstfrm(0), Array(frmMastAccount.txtTemp(2)), Array(GName)
        Case 1 '-- Mill
            GName = GProcGetColumnValue("TblMastGroup", "AgCode", 90029, "N", "AgName", "S")
            GProcShowForm frmMastAccount, frmMain.mnuMstfrm(0), Array(frmMastAccount.txtTemp(2)), Array(GName)
        Case 3 '-- Broker
            GName = GProcGetColumnValue("TblMastGroup", "AgCode", 90016, "N", "AgName", "S")
            GProcShowForm frmMastAccount, frmMain.mnuMstfrm(0), Array(frmMastAccount.txtTemp(3)), Array(GName)
    End Select
End If
End Sub
Private Sub txtTemp_KeyPress(Index As Integer, KeyAscii As Integer)
Dim LeftPos As Long
Dim TopPos As Long
TopPos = txtTemp(Index).Top + Me.Top + 650
LeftPos = txtTemp(Index).Left + Me.Left
If KeyAscii = 13 Then Exit Sub
Select Case Index
    Case 0 '--Party
        gClsSearch.SearchMultiField "QryMastaccount", "Account,AcName,AcCity,AgName", Array("Account", "AcName", "City", "Group"), Array(3000, 0, 2000, 3000), "(AgCode=90017 or AgCode=90015 or AgCode=90029)", Chr(KeyAscii), "AcName", LeftPos, TopPos
        KeyAscii = 0
        If gClsSearch.SearchMultiRetCol(1) <> "" Then txtTemp(Index) = gClsSearch.SearchMultiRetCol(1)
    Case 1 '--Mill
        gClsSearch.SearchMultiField "QryMastaccount", "Account,AcName,AcCity,AgName", Array("Account", "AcName", "City", "Group"), Array(3000, 0, 2000, 3000), "AgCode=90029 ", Chr(KeyAscii), "AcName", LeftPos, TopPos
        KeyAscii = 0
        If gClsSearch.SearchMultiRetCol(1) <> "" Then txtTemp(Index) = gClsSearch.SearchMultiRetCol(1)
    Case 3 '--Broker
        gClsSearch.SearchMultiField "QryMastaccount", "Account,AcName,AcCity,AgName", Array("Account", "AcName", "City", "Group"), Array(3000, 0, 2000, 3000), "AgCode=90016", Chr(KeyAscii), "AcName", LeftPos, TopPos
        KeyAscii = 0
        If gClsSearch.SearchMultiRetCol(1) <> "" Then txtTemp(Index) = gClsSearch.SearchMultiRetCol(1)
End Select
End Sub

Private Sub txtTemp_LostFocus(Index As Integer)
lblPartyTin.Caption = GProcGetColumnValue("tblMastAccount", "AcName", txtTemp(0), "S", "AcCST", "S")
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
    txtTemp(1).Enabled = False
    txtTemp(1) = ""
Else
    txtTemp(1).Enabled = True
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
PartyCode = GProcGetColumnValue("tblmastAccount", "AcName", txtTemp(0), "S", "AcCode", "N")
Select Case gSelectedMenu
    Case "CR" '----- C Form Rec
        If cboType = "Mill Bill" Then
            CondStr = " tblAddLess.vtype ='SM' and instr(narration,'C Form')>0"
            MillCode = GProcGetColumnValue("tblmastAccount", "AcName", txtTemp(1), "S", "AcCode", "N")
            CondStr = CondStr + " and  tblAddLess.AdMillCode=" & MillCode & " and tblAddLess.ADACDRCODE=" & PartyCode
        Else
            CondStr = " tblAddLess.Vtype <> 'SM' and instr(narration,'C Form')>0  And tblAddLess.ADACDRCODE = " & PartyCode
        End If
        If IsDate(mskDt(0)) = True And IsDate(mskDt(1)) = True Then
            If gBackEndDB = gBackEndAccess Then
                DtStr = " tblAddLess.vdt>=cdate('" & mskDt(0) & "') and tblAddLess.Vdt<=cdate('" & mskDt(1) & "')"
            Else
                DtStr = " tblAddLess.vdt>=to_date('" & CDate(mskDt(0)) & "','dd/MM/yyyy') and tblAddLess.Vdt<=To_date('" & CDate(mskDt(1)) & "','dd/MM/yyyy')"
            End If
        Else
            DtStr = " tblAddLess.vno=0"
        End If
        '--Vtype,AdInvNo,Vdt,AdBillAmt, ADCFORMNO,Vno,VYear
        X = "select tblAddLess.Vtype,tblAddLess.AdInvNo,tblAddLess.Vdt,tblMastAccount_Br.Acname,tblAddLess.AdBillAmt, tblAddLess.ADCFORMNO,tblAddLess.Vno,tblAddLess.VYear from tblAddLess,tblmastNarration,tblMastAccount tblMastAccount_Br,tblOutStanding where NarrCode=ADTAXCODE and tblAddless.Vno = tblOutStanding.Vno and tblAddless.Vtype = tblOutStanding.Vtype and tblAddless.Vyear = tblOutStanding.Vyear and tblOutStanding.OutBrokerCode = tblMastAccount_Br.AcCode and " _
        & DtStr & " and " & CondStr & " order by tblAddLess.Vdt,tblAddLess.AdInvNo"
    Case "E2I" '----- C Form Rec
         CondStr = " tblAddLess.Vtype = 'ST' and instr(narration,'C Form')>0  And tblAddLess.ADACDRCODE = " & PartyCode
        If IsDate(mskDt(0)) = True And IsDate(mskDt(1)) = True Then
            If gBackEndDB = gBackEndAccess Then
                DtStr = " tblAddLess.vdt>=cdate('" & mskDt(0) & "') and tblAddLess.Vdt<=cdate('" & mskDt(1) & "')"
            Else
                DtStr = " tblAddLess.vdt>=to_date('" & CDate(mskDt(0)) & "','dd/MM/yyyy') and tblAddLess.Vdt<=To_date('" & CDate(mskDt(1)) & "','dd/MM/yyyy')"
            End If
        Else
            DtStr = " tblAddLess.vno=0"
        End If
        '--Vtype,AdInvNo,Vdt,AdBillAmt, ADCFORMNO,Vno,VYear
        X = "select tblAddLess.Vtype,tblAddLess.AdInvNo,tblAddLess.Vdt,tblMastAccount_Br.Acname,tblAddLess.AdBillAmt, tblAddLess.ADE1FORMNO,tblAddLess.Vno,tblAddLess.VYear from tblAddLess,tblmastNarration,tblMastAccount tblMastAccount_Br,tblOutStanding where NarrCode=ADTAXCODE and tblAddless.Vno = tblOutStanding.Vno and tblAddless.Vtype = tblOutStanding.Vtype and tblAddless.Vyear = tblOutStanding.Vyear and tblOutStanding.OutBrokerCode = tblMastAccount_Br.AcCode and  " _
        & DtStr & " and tblAddless.ADTMP2 = 1 and " & CondStr & " order by tblAddLess.Vdt,tblAddLess.AdInvNo"
    
    Case "CI", "E1R", "E2R" '----- C Form Issue ,E1 Receive
        If gSelectedMenu = "CI" Then '----- C Form Issue
            CondStr = " tblAddLess.vtype in ('PY','PO','PT') and instr(narration,'C Form')>0"
        Else '-- E1 Receive
            CondStr = " tblAddLess.vtype in ('PT') and instr(narration,'C Form')>0"
        End If
        MillCode = GProcGetColumnValue("tblmastAccount", "AcName", txtTemp(1), "S", "AcCode", "N")
        If MillCode <> 0 Then
           CondStr = CondStr + " and  tblAddLess.AdMillCode=" & MillCode & " and tblAddLess.ADACCRCODE=" & PartyCode
        Else
           CondStr = CondStr + " and tblAddLess.ADACCRCODE=" & PartyCode
        End If
'        CondStr = CondStr + " and  AdMillCode=" & MillCode & " and ADACCRCODE=" & PartyCode
        If IsDate(mskDt(0)) = True And IsDate(mskDt(1)) = True Then
            If gBackEndDB = gBackEndAccess Then
                DtStr = " tblAddLess.adGpDt>=cdate('" & mskDt(0) & "') and tblAddLess.adGpDt<=cdate('" & mskDt(1) & "')"
            Else
                DtStr = " tblAddLess.adGpDt>=to_date('" & CDate(mskDt(0)) & "','dd/MM/yyyy') and tblAddLess.adGpDt<=To_date('" & CDate(mskDt(1)) & "','dd/MM/yyyy')"
            End If
        Else
            DtStr = " tblAddLess.vno=0"
        End If
        '--Vtype,AdInvNo,Vdt,AdBillAmt, ADCFORMNO,Vno,VYear
         Select Case gSelectedMenu
        Case "CI"
            X = "select tblAddLess.Vtype,tblAddLess.AdInvNo,tblAddLess.adGpDt,tblMastAccount_Br.Acname,tblAddLess.AdBillAmt, tblAddLess.ADCFORMNO,tblAddLess.Vno,tblAddLess.VYear from tblAddLess,tblmastNarration ,tblMastAccount tblMastAccount_Br,tblOutStanding where NarrCode=ADTAXCODE and tblAddless.Vno = tblOutStanding.Vno and tblAddless.Vtype = tblOutStanding.Vtype and tblAddless.Vyear = tblOutStanding.Vyear and tblOutStanding.OutBrokerCode = tblMastAccount_Br.AcCode and " _
            & DtStr & " and " & CondStr & " order by tblAddLess.AdGpdt,tblAddLess.AdInvNo"
        Case "E1R"
            X = "select tblAddLess.Vtype,tblAddLess.AdInvNo,tblAddLess.adGpDt,tblMastAccount_Br.Acname,tblAddLess.AdBillAmt,tblAddLess.ADE1FORMNO,tblAddLess.Vno,tblAddLess.VYear from tblAddLess,tblmastNarration ,tblMastAccount tblMastAccount_Br,tblOutStanding where NarrCode=ADTAXCODE and tblAddless.Vno = tblOutStanding.Vno and tblAddless.Vtype = tblOutStanding.Vtype and tblAddless.Vyear = tblOutStanding.Vyear and tblOutStanding.OutBrokerCode = tblMastAccount_Br.AcCode and " _
            & DtStr & " and tblAddless.AdgpNo <> 1 and " & CondStr & " order by tblAddLess.AdGpdt,tblAddLess.AdInvNo"
        Case "E2R"
            X = "select tblAddLess.Vtype,tblAddLess.AdInvNo,tblAddLess.adGpDt,tblMastAccount_Br.Acname,tblAddLess.AdBillAmt, tblAddLess.ADE1FORMNO,tblAddLess.Vno,tblAddLess.VYear from tblAddLess,tblmastNarration ,tblMastAccount tblMastAccount_Br,tblOutStanding where NarrCode=ADTAXCODE and tblAddless.Vno = tblOutStanding.Vno and tblAddless.Vtype = tblOutStanding.Vtype and tblAddless.Vyear = tblOutStanding.Vyear and tblOutStanding.OutBrokerCode = tblMastAccount_Br.AcCode and " _
            & DtStr & " and tblAddless.AdgpNo = 1 and " & CondStr & " order by tblAddLess.AdGpdt,tblAddLess.AdInvNo"
        End Select
        
''        If gSelectedMenu = "CI" Then '----- C Form Issue
''            X = "select Vtype,AdInvNo,adGpDt,AdBillAmt, ADCFORMNO,Vno,VYear from tblAddLess,tblmastNarration where NarrCode=ADTAXCODE and " _
''            & DtStr & " and " & CondStr & " order by AdGpdt,AdInvNo"
''        Else
''            X = "select Vtype,AdInvNo,adGpDt,AdBillAmt, ADE1FORMNO,Vno,VYear from tblAddLess,tblmastNarration where NarrCode=ADTAXCODE and " _
''            & DtStr & " and " & CondStr & " order by AdGpdt,AdInvNo"
''        End If
    Case "FI" '-- F Form Issue
        CondStr = " tblAddLess.vtype in ('PI') "
        MillCode = GProcGetColumnValue("tblmastAccount", "AcName", txtTemp(1), "S", "AcCode", "N")
        CondStr = CondStr + " and  tblAddLess.AdMillCode=" & MillCode & " and tblAddLess.ADACCRCODE=" & PartyCode
        If IsDate(mskDt(0)) = True And IsDate(mskDt(1)) = True Then
            If gBackEndDB = gBackEndAccess Then
                DtStr = " tblAddLess.adGpDt>=cdate('" & mskDt(0) & "') and tblAddLess.adGpDt<=cdate('" & mskDt(1) & "')"
            Else
                DtStr = " tblAddLess.adGpDt>=to_date('" & CDate(mskDt(0)) & "','dd/MM/yyyy') and tblAddLess.adGpDt<=To_date('" & CDate(mskDt(1)) & "','dd/MM/yyyy')"
            End If
        Else
            DtStr = " tblAddLess.vno=0"
        End If
        '--Vtype,AdInvNo,Vdt,AdBillAmt, ADCFORMNO,Vno,VYear
        X = "select tblAddLess.Vtype,tblAddLess.AdGPNo,tblAddLess.AdGpdt,tblMastAccount_Br.Acname,tblAddLess.AdBillAmt, tblAddLess.ADCFORMNO,tblAddLess.Vno,tblAddLess.VYear from tblAddLess ,tblMastAccount tblMastAccount_Br where  tblAddLess.AdAcDrCode = tblMastAccount_Br.AcCode and  " _
        & DtStr & " and " & CondStr & " order by tblAddLess.AdGpdt,tblAddLess.AdGpNo"
End Select
Set Rstbl = New Recordset
GProcRstOpen Rstbl, X, "R"
With msGrid
    .Clear
    .FormatString = "<Type       |>            Bill No|< Date                 |<Broker                            |>             Bill Amount|<Form No              |>Vno|>Vyear"
    .ColWidth(6) = 0
    .ColWidth(7) = 0
    .Rows = 2
    'FillRelatedRecord
End With
If gSelectedMenu = "CI" Then
    Dim X2 As String
    Dim i As Integer
    Dim rstCheck As Recordset
    Dim BrkCd As Integer
    X2 = "select tblAddless.Adtmp3 from tblAddLess,tblmastNarration  where   " _
            & DtStr & " and " & CondStr & " order by tblAddLess.AdGpdt,tblAddLess.AdGpNo"
    
    i = GProcRstOpen(rstCheck, X2, "R")
       If i > 0 Then
        BrkCd = rstCheck.Fields(0)
        DoEvents
        txtTemp(3).Text = GProcGetColumnValue("tblMastAccount", "AcCode", rstCheck.Fields(0), "N", "AcName", "S")
        If txtTemp(3).Text = "" Then
           txtTemp(3).Text = "-"
        End If
        
    End If
End If


End Sub
Private Sub msGrid_Click()
On Error Resume Next
    msGrid_EnterCell
End Sub
Private Sub msGrid_GotFocus()
If LCase(Me.ActiveControl) <> "txtgrid" Then msGrid.Col = 5
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
    If msGrid.TextMatrix(msGrid.Row, 5) = "" Then msGrid.TextMatrix(msGrid.Row, 5) = tmpFormNo  ''-Form No
    If msGrid.Text <> "" Then txtGrid.Text = msGrid.Text
    '--- Lock TextBox
    With msGrid
        Select Case .Col
            Case 0, 1, 2, 3, 4 '--Type,Bill No,Date,Amt
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
        If .Col = 5 Then
            If .Row = (.Rows - 1) Then '-- Add New Row
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
If msGrid.Col < 6 Then
    txtGrid.Move msGrid.CellLeft + msGrid.Left, msGrid.CellTop + msGrid.Top, msGrid.CellWidth, msGrid.CellHeight
End If
txtGrid.Visible = True
txtGrid.Enabled = True
DoEvents
If KeyAscii <> 0 Then
    Select Case Rstbl.Fields(msGrid.Col).Type
        Case adChar, adVarWChar
            txtGrid.Text = txtGrid.Text & Chr(KeyAscii)
        Case Else
            txtGrid.Text = Chr(GProcValidateKey(Rstbl, msGrid.Col, KeyAscii, txtGrid.Text)) 'Chr(GProcNumberWithDecimal(KeyAscii, txtGrid))
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
    tmpFormNo = msGrid.TextMatrix(msGrid.Row, 5)
End If
If msGrid.Col = 5 Then CalGridAmt
End Sub
Private Sub CalGridAmt()
Dim RowIndex As Long
Dim Amt As Double
RowIndex = 1
With msGrid
    If txtGrid.Visible = True Then .TextMatrix(.Row, .Col) = txtGrid
    Do While RowIndex < .Rows
        '-- Bill amt with Form No
        If .TextMatrix(RowIndex, 5) = "" And RowIndex > 1 Then
           .TextMatrix(RowIndex, 5) = .TextMatrix(RowIndex - 1, 5)
        End If
 
        If .TextMatrix(RowIndex, 5) <> "" Then
            Amt = Amt + CDbl(.TextMatrix(RowIndex, 4))
        End If
        RowIndex = RowIndex + 1
    Loop
End With
txtTemp(2) = Format(Amt, FStr)
End Sub
Private Sub txtGrid_GotFocus()
    txtGrid.Alignment = ProcSetAlignment(Rstbl, msGrid.Col)
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
            Case 0, 1, 2, 3, 4
                KeyAscii = 0
            Case Else
                KeyAscii = GProcValidateKey(Rstbl, msGrid.Col, KeyAscii, txtGrid)
        End Select
   End If
    If KeyAscii = 13 Then
        CheckForNumberTxtGrid
    End If
    '--- Navigate Row ,Col
    If KeyAscii = 13 And .Col < 5 Then
        .Col = .Col + 1
        If .Col = 5 Then tmpFormNo = txtGrid
    ElseIf KeyAscii = 13 And .Col = 5 Then
        If .Row = (.Rows - 1) Then
        Else
            .Row = .Row + 1
        End If
        tmpFormNo = txtGrid
        .Col = 5
     '   .Text = tmpFormNo
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
With Rstbl
For RowIndex = 1 To msGrid.Rows - 1
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
                    msGrid.TextMatrix(RowIndex, i) = Null
                End If
        End Select
    Next i
  Next RowIndex
End With
End Sub
'---  Check For Number Grid
Private Sub CheckForNumberTxtGrid()
With Rstbl
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
                txtGrid = Null
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
Dim Brk As Integer
'Dim Rstbl1 As Recordset
'--- Save Grid Records
CheckForNumberGrid
If txtTemp(3).Text = "" Then
   txtTemp(3).Text = "-"
End If
Brk = GProcGetColumnValue("tblMastAccount", "AcName", txtTemp(3), "S", "AcCode", "N")
With msGrid
    i = 1
    Do While i < .Rows
        Vno = CLng(.TextMatrix(i, 6)) '--Vno
        Vtype = .TextMatrix(i, 0) '--VType
        VYear = CLng(.TextMatrix(i, 7)) '--VYear
        FormNo = .TextMatrix(i, 5) '--FromNo
        If gSelectedMenu = "E1R" Or gSelectedMenu = "E2R" Or gSelectedMenu = "E2I" Then   '-- E 1 From Receive
            gCn.Execute "update tbladdless set ADE1FORMNO='" & FormNo & "' where vno=" & Vno & " and Vtype='" & Vtype & "' and Vyear=" & VYear
        Else
            gCn.Execute "update tbladdless set ADCFORMNO='" & FormNo & "' where vno=" & Vno & " and Vtype='" & Vtype & "' and Vyear=" & VYear
            gCn.Execute "update tbladdless set ADTMP3=" & Brk & " where vno=" & Vno & " and Vtype='" & Vtype & "' and Vyear=" & VYear
        End If
        i = i + 1
     Loop
End With
End Sub
Private Sub FillRelatedRecord()
Dim RowIndex As Long
Dim ColIndex As Long
tmpFormNo = ""
txtGrid.Visible = False
'--Fill Grid
SetGrid
With Rstbl
    If Not .EOF Then
        .MoveFirst
        RowIndex = 1
        msGrid.Rows = 2
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
            If Not .EOF Then msGrid.Rows = msGrid.Rows + 1
        Loop
    End If
End With
CalGridAmt
End Sub
