VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form frmGPDeliveryEntryNew 
   BackColor       =   &H00C2E3FC&
   BorderStyle     =   0  'None
   Caption         =   "Goods Inward"
   ClientHeight    =   9045
   ClientLeft      =   0
   ClientTop       =   15
   ClientWidth     =   10905
   FillColor       =   &H000040C0&
   FillStyle       =   0  'Solid
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   9045
   ScaleWidth      =   10905
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox txttempGodown1 
      Enabled         =   0   'False
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
      Left            =   2475
      MaxLength       =   40
      TabIndex        =   34
      Text            =   "txttempGodown1"
      Top             =   3645
      Width           =   4635
   End
   Begin VB.TextBox txttempItem 
      Enabled         =   0   'False
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
      Left            =   2475
      MaxLength       =   40
      TabIndex        =   32
      Text            =   "txttempItem"
      Top             =   3150
      Width           =   4635
   End
   Begin VB.TextBox txttempMill 
      Enabled         =   0   'False
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
      Left            =   2475
      MaxLength       =   40
      TabIndex        =   31
      Text            =   "txttempMill"
      Top             =   2610
      Width           =   4635
   End
   Begin VB.TextBox txttempParty 
      Enabled         =   0   'False
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
      Left            =   2475
      MaxLength       =   40
      TabIndex        =   26
      Text            =   "txttempParty"
      Top             =   1605
      Width           =   4635
   End
   Begin VB.TextBox txttempGodown 
      Enabled         =   0   'False
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
      Left            =   2475
      MaxLength       =   40
      TabIndex        =   25
      Text            =   "txttempGodown"
      Top             =   1125
      Width           =   4635
   End
   Begin VB.TextBox txt 
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
      Index           =   0
      Left            =   2475
      TabIndex        =   1
      Text            =   "0"
      ToolTipText     =   "Press F3 for Creation"
      Top             =   615
      Width           =   1230
   End
   Begin VB.TextBox txttmpBag 
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
      Height          =   360
      Left            =   2475
      MaxLength       =   40
      TabIndex        =   6
      Text            =   "txttmpBag"
      Top             =   2085
      Width           =   1395
   End
   Begin VB.TextBox txt 
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
      Left            =   11220
      TabIndex        =   4
      Text            =   "2"
      ToolTipText     =   "Press F3 for Creation"
      Top             =   3180
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.TextBox txt 
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
      Left            =   11280
      TabIndex        =   3
      Text            =   "1"
      ToolTipText     =   "Press F3 for Creation"
      Top             =   2340
      Visible         =   0   'False
      Width           =   435
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
      Left            =   8175
      Style           =   1  'Graphical
      TabIndex        =   10
      ToolTipText     =   "Exit"
      Top             =   2745
      Width           =   1095
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
      Left            =   8175
      Style           =   1  'Graphical
      TabIndex        =   9
      ToolTipText     =   "To Save Record"
      Top             =   2235
      Width           =   1095
   End
   Begin VB.TextBox txtGrid 
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   2340
      TabIndex        =   8
      Top             =   6840
      Visible         =   0   'False
      Width           =   1035
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
      Left            =   11580
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   6240
      Visible         =   0   'False
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
      Left            =   11400
      Style           =   1  'Graphical
      TabIndex        =   11
      ToolTipText     =   "To Add New Record"
      Top             =   7020
      Visible         =   0   'False
      Width           =   1095
   End
   Begin MSFlexGridLib.MSFlexGrid msGrid 
      Height          =   3975
      Left            =   480
      TabIndex        =   7
      Top             =   4590
      Width           =   10020
      _ExtentX        =   17674
      _ExtentY        =   7011
      _Version        =   393216
      FixedCols       =   0
      RowHeightMin    =   275
      BackColorFixed  =   -2147483647
      ForeColorFixed  =   -2147483634
      BackColorBkg    =   12632256
      Appearance      =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSMask.MaskEdBox mskDt 
      Height          =   375
      Index           =   0
      Left            =   6030
      TabIndex        =   2
      Top             =   600
      Width           =   1230
      _ExtentX        =   2170
      _ExtentY        =   661
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
   Begin VB.Label Label8 
      BackStyle       =   0  'Transparent
      Caption         =   "Godown Name   :  "
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
      Left            =   495
      TabIndex        =   33
      Top             =   3690
      Width           =   1635
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "Count Name      :  "
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
      Left            =   495
      TabIndex        =   30
      Top             =   3150
      Width           =   1635
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Mill Name         :  "
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
      Left            =   540
      TabIndex        =   29
      Top             =   2520
      Width           =   1635
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "G.P.  Date        :  "
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
      Left            =   4095
      TabIndex        =   28
      Top             =   660
      Width           =   1635
   End
   Begin VB.Shape Shape3 
      BorderColor     =   &H00000080&
      FillColor       =   &H00EDFDFE&
      FillStyle       =   0  'Solid
      Height          =   4530
      Left            =   240
      Top             =   4350
      Width           =   10440
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "Party Name      :  "
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
      Left            =   540
      TabIndex        =   27
      Top             =   1605
      Width           =   1635
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Broker Name    :  "
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
      Left            =   540
      TabIndex        =   24
      Top             =   1125
      Width           =   1635
   End
   Begin VB.Image Image1 
      Height          =   1140
      Left            =   7905
      Stretch         =   -1  'True
      Top             =   525
      Width           =   1335
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "&Gate Pass No   :  "
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
      Left            =   540
      TabIndex        =   0
      Top             =   615
      Width           =   1635
   End
   Begin VB.Label lbl1 
      BackStyle       =   0  'Transparent
      Caption         =   "&Bags                :"
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
      Index           =   0
      Left            =   540
      TabIndex        =   5
      Top             =   2085
      Width           =   1650
   End
   Begin VB.Shape Shape1 
      FillColor       =   &H00BDDFD1&
      FillStyle       =   0  'Solid
      Height          =   1620
      Left            =   8040
      Shape           =   4  'Rounded Rectangle
      Top             =   1845
      Width           =   1350
   End
   Begin VB.Shape Shape5 
      BorderColor     =   &H00C00000&
      BorderWidth     =   5
      FillColor       =   &H00C00000&
      Height          =   10140
      Left            =   0
      Top             =   -1155
      Width           =   10860
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
      ForeColor       =   &H0000FFFF&
      Height          =   495
      Left            =   7290
      TabIndex        =   23
      Top             =   15
      Width           =   2475
   End
   Begin VB.Label Label27 
      BackStyle       =   0  'Transparent
      Caption         =   "Landing Cost Per Bag"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   735
      Left            =   12255
      TabIndex        =   22
      Top             =   7800
      Width           =   975
   End
   Begin VB.Label Label26 
      BackStyle       =   0  'Transparent
      Caption         =   "Landing Cost"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   255
      Left            =   12375
      TabIndex        =   21
      Top             =   8160
      Width           =   1455
   End
   Begin VB.Label Label25 
      BackStyle       =   0  'Transparent
      Caption         =   "Other Less"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   255
      Left            =   11415
      TabIndex        =   20
      Top             =   8280
      Width           =   1215
   End
   Begin VB.Label Label24 
      BackStyle       =   0  'Transparent
      Caption         =   "Other Add"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   255
      Left            =   11295
      TabIndex        =   19
      Top             =   7920
      Width           =   1095
   End
   Begin VB.Label Label22 
      BackStyle       =   0  'Transparent
      Caption         =   "Godown Rent"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   255
      Left            =   11175
      TabIndex        =   18
      Top             =   8160
      Width           =   1575
   End
   Begin VB.Label Label21 
      BackStyle       =   0  'Transparent
      Caption         =   "Freight"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   255
      Left            =   11175
      TabIndex        =   17
      Top             =   8040
      Width           =   1335
   End
   Begin VB.Label Label20 
      BackStyle       =   0  'Transparent
      Caption         =   "Total Amount"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   255
      Left            =   11415
      TabIndex        =   16
      Top             =   7860
      Width           =   1455
   End
   Begin VB.Label Label19 
      BackStyle       =   0  'Transparent
      Caption         =   "Landing Cost Details :-"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   375
      Left            =   11175
      TabIndex        =   15
      Top             =   7590
      Width           =   3135
   End
   Begin VB.Label Label13 
      BackStyle       =   0  'Transparent
      Caption         =   "Ass.Value"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   255
      Left            =   11055
      TabIndex        =   14
      Top             =   8160
      Width           =   1095
   End
   Begin VB.Label Label6 
      BackColor       =   &H00C00000&
      Caption         =   " Gate Pass Delivery Entry"
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
      Left            =   45
      TabIndex        =   13
      Top             =   0
      Width           =   10815
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H00000080&
      FillColor       =   &H00EDFDFE&
      FillStyle       =   0  'Solid
      Height          =   3840
      Left            =   270
      Top             =   465
      Width           =   7260
   End
End
Attribute VB_Name = "frmGPDeliveryEntryNew"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'Dim rstbl As Recordset
Dim rstblSub As Recordset
Dim FirstTimeFlag As Boolean
Dim oldFrmTypeStr As String
Dim tmpBagNo As String
Dim i As Long
Private Sub cmdBtn_Click(Index As Integer)
'On Error GoTo ErrorRoutine
Dim X As String
Select Case Index
Case 10 'Exit
    Unload Me
Case 13 'Save
    If txtGrid.Visible = True Then msGrid.text = txtGrid
    txtGrid = ""
    If ValidateData = True Then
        txtGrid.Visible = False
        SaveRelatedRecord
        DoEvents
        DoEvents
        DoEvents
        txt(0) = 0
        Call GProcClearForm(Me, rstblSub, 0, False)
        SetGrid
        cmdBtn(13).Enabled = False
        txt(0).SetFocus
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
           If Me.ActiveControl.Name <> "txtGrid" Then Sendkeys "{TAB}"
        Case 27:
            If (cmdBtn(13).Enabled = False) Then
                Unload Me
            End If
    End Select
End Sub
Private Sub Form_Load()
    Me.Top = 300
    Me.Left = Screen.Width / 2 - Me.Width / 2
    FirstTimeFlag = True
    txttmpBag.Enabled = False '-Bags Total
    txttempGodown.Enabled = False
    txttempGodown1.Enabled = False
    txttempParty.Enabled = False
    txttempMill.Enabled = False
    txttempItem.Enabled = False
    mskDt(0).Enabled = False
    '---Grid
    SetGrid
    Call GProcClearForm(Me, rstblSub, 0, False)
    If FirstTimeFlag = True Then
        FirstTimeFlag = False
    Else
        txt(0).SetFocus
    End If
End Sub
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
'If rstbl.State = adStateOpen Then rstbl.Close
If rstblSub.State = adStateOpen Then rstblSub.Close
End Sub

Private Sub mskDt_LostFocus(Index As Integer)
SaveBtnEd
If IsDate(mskDt(Index)) = False Then mskDt(Index) = "__/__/____"
End Sub
Private Sub mskDt_GotFocus(Index As Integer)
GProcSelectBox Me.ActiveControl
End Sub
Private Sub SaveBtnEd()
'--GP No
If txt(0) = 0 Then
    cmdBtn(13).Enabled = False
    cmdBtn(13).BackColor = &H80000004
Else
    cmdBtn(13).Enabled = True
    cmdBtn(13).BackColor = &HFFFF80
End If
End Sub

Private Function ValidateData() As Boolean
'---- Required Data
'--- Validate Grid
If ValidateGrid = False Then Exit Function
'--G P Vno
If GProcGetColumnValue("tblAddLess", "Vno", txt(0), "N", "Vno", "N", "vtype='GP' and Vyear=" & gCYear) = 0 Then
    MsgBox "Check Gate Pass No.", vbInformation + vbOKOnly, Me.Caption
    txt(0).SetFocus
    Exit Function
End If
''--- Delivery dt
'If IsDate(mskDt(0)) = False Then
'    MsgBox "Check Delivery Date.", vbInformation + vbOKOnly, Me.Caption
'    mskDt(0).SetFocus
'    Exit Function
'End If
ValidateData = True
End Function
Private Sub txt_GotFocus(Index As Integer)
GProcSelectBox Me.ActiveControl
End Sub
Private Sub txt_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyF3 Then
    Select Case Index
        Case 2 '-- Sizer
            oldFrmTypeStr = gfrmTypeStr
            gfrmTypeStr = "S"
            GProcShowForm frmMastNarration, frmMain.mnuMstfrm(9), "S"
            gfrmTypeStr = oldFrmTypeStr
    End Select
End If
End Sub
Private Sub txt_KeyPress(Index As Integer, KeyAscii As Integer)
Dim LeftPos As Long
Dim TopPos As Long
TopPos = txt(Index).Top + Me.Top + 650
LeftPos = txt(Index).Left + Me.Left
If KeyAscii = 13 Then Exit Sub
Select Case Index
    Case 2 '--Sizer
        gClsSearch.SearchMultiField "tblMastNarration", "Narration", Array("Sizer"), Array(txt(Index).Width), " NarrType='S' ", Chr(KeyAscii), "Narration", LeftPos, TopPos
        KeyAscii = 0
        If gClsSearch.SearchMultiRetCol(0) <> "" Then txt(Index) = gClsSearch.SearchMultiRetCol(0)
End Select
End Sub
Private Sub txt_LostFocus(Index As Integer)
SaveBtnEd
If Index = 0 Then FillRelatedRecord
End Sub
Private Sub SetGrid()
Dim X As String
Dim RCount As Integer
Dim RsGp As Recordset
Dim RowIndex As Integer
Dim ColIndex As Integer
txttmpBag = "0"
tmpBagNo = "" ''-Auto Bag No
'--Fill Gp Details
X = "select VSubBag,tblMastAccount.AcName As Party,tblMastAccountMill.AcName As Mill,tblMastAccount_Br.AcName As Broker,tblMastItem.ItName As Item,tblMastNarration.Narration As Godown,tblVouSub.Vdt from tblvousub,TblAddLess,tblMastAccount,tblMastItem,tblMastNarration,tblMastAccount tblMastAccountMill,tblMastAccount tblMastAccount_Br  where " _
& " tblvousub.Vno=TblAddLess.Vno and tblvousub.Vtype=TblAddLess.Vtype and tblvousub.VYear=TblAddLess.Vyear and tblvousub.VSubGodownCode = tblMastNarration.Narrcode and " _
& " tblvousub.Vno=" & txt(0) & " and tblvousub.vtype='GP' and tblvousub.VYear= " & gCYear & " and " _
& " tblVouSub.VsubItcode = tblMastItem.ItCode and  tblAddLess.AdAcCrCode = tblMastAccount.AcCode and tblAddLess.AdMillCode = tblMastAccountMill.AcCode and tblAddLess.AdAcDrCode = tblMastAccount_br.AcCode"


Set RsGp = New Recordset
    If GProcRstOpen(RsGp, X, "R") > 0 Then
        With RsGp
            txttmpBag = .Fields("VSubBag") '-G P Bags
            txttempGodown = .Fields("Broker")
            txttempParty = .Fields("Party")
            txttempMill = .Fields("Mill")
            txttempItem = .Fields("Item")
            txttempGodown1 = .Fields("Godown")
            If IsDate(.Fields("VDt")) = True Then  '-Delivery Dt
                mskDt(0) = CDate(.Fields("Vdt"))
            Else
                mskDt(0) = "__/__/____"
            End If
'            txt(1) = .fields("AdTransport") & vbNullString '-Vehicle No
'            txt(2) = .fields("AdTmp8") & vbNullString  '-Sizer
'           txt(2) = .Fields("AdTmp5") & vbNullString  '-Sizer
        End With
    End If
txttmpBag = Format(txttmpBag, PStr)
'--BagNo,cone,Wt,InwNo,InwType,GpNo,GpType
X = "select VSubCtrNo,VSubBag,VSubWt,VDeliveryDt,VehicleNo,tblMastNarration.Narration from tblGpSub,tblMastNarration where " _
& " tblGpSub.Vno=" & txt(0) & " and tblGpSub.vtype='GP' and  tblGpSub.VSizer = tblMastNarration.NarrCode  and tblGpSub.VYear=" & gCYear
Set rstblSub = New Recordset
RCount = GProcRstOpen(rstblSub, X, "R")
With msGrid
    .Clear
    .FormatString = ">Sr.No.|>Bag     |>Weight      |>Delevery On   |<Vehicle No         |<Sizer                                       "
    .Rows = RCount + 1
End With
'-- Fill Record
With rstblSub
    If Not .EOF Then
        .MoveFirst
        RowIndex = 1
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
           ' If Not .EOF Then msGrid.Rows = msGrid.Rows + 1
        Loop
    End If
End With
End Sub

Private Sub msGrid_Click()
On Error Resume Next
    msGrid_EnterCell
End Sub
Private Sub msGrid_GotFocus()
' cmdBtn(10).Cancel = False
End Sub
Private Sub msGrid_EnterCell()
'If FormAction <> vbDataActionClose Then
    CheckForNumberTxtGrid
    DoEvents
    txtGrid = ""
    DoEvents
    If Trim(msGrid.text) <> "" Then txtGrid.text = msGrid.text
    '--- Lock TextBox
    With msGrid
        Select Case .Col
            Case 0:
                 txtGrid.Locked = True
                 .Col = .Col + 1
            Case 1:
                 txtGrid.Locked = True
                 .Col = .Col + 1
            Case 2:
                 txtGrid.Locked = True
                 .Col = .Col + 1
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
End With
    '---- Navigate Row ,Col
'''    If KeyAscii = 13 Then
'''        KeyAscii = 0
'''      With msGrid
'''        If .Col = 6 Then
'''            If .Row = (.rows - 1) Then '-- Add New Row
'''            Else '--Next Row
'''                .Row = .Row + 1
'''            End If
'''            .Col = 0
'''        Else    '-- Next col
'''            .Col = .Col + 1
'''        End If
'''        End With
'''    End If


If KeyAscii = 13 Then
    KeyAscii = 0
    With msGrid
        If .Col = 5 Then
            If .Row = (.Rows - 1) Then '-- Add New Row
'                .rows = .rows + 1
'                .Row = .Row + 1
'                GridAddNew
            Else '--Next Row
                .Row = .Row + 1
            End If
            .Col = 1
        Else    '-- Next col
'            If .Col = 6 Or .Col = 7 Or .Col = 8 Then
'               .Col = 9
'            Else
               .Col = .Col + 1
'            End If
            
        End If
    End With
End If
'---- Set TxtGrid

If msGrid.Col = 6 Or msGrid.Col = 7 Or msGrid.Col = 8 Then
   msGrid.Col = 9
End If






'    If KeyAscii <> 13 And KeyAscii <> 27 And .Col = 0 Then
'        DoEvents
'        KeyAscii = 0
'        Pending_In
'    End If
'ProcSetAlignment rstblSub, msGrid.Col
'---- Set TxtGrid
If msGrid.text <> "" Then txtGrid.text = msGrid.text
If msGrid.Col < 6 Then
    txtGrid.Move msGrid.CellLeft + msGrid.Left, msGrid.CellTop + msGrid.Top, msGrid.CellWidth, msGrid.CellHeight
End If
txtGrid.Visible = True
txtGrid.Enabled = True
DoEvents
If KeyAscii <> 0 Then
    Select Case rstblSub.Fields(msGrid.Col).Type
        Case adChar, adVarWChar
            txtGrid.text = txtGrid.text & Chr(KeyAscii)
        Case Else
            txtGrid.text = Chr(GProcValidateKey(rstblSub, msGrid.Col, KeyAscii, txtGrid.text)) 'Chr(GProcNumberWithDecimal(KeyAscii, txtGrid))
    End Select
End If
DoEvents
If msGrid.text <> "" Then
    txtGrid.SelStart = Len(txtGrid.text) + 1
Else
    txtGrid.SelStart = 1
End If
DoEvents
If txtGrid.Visible = True Then txtGrid.SetFocus
End Sub
Private Sub msGrid_LeaveCell()
If txtGrid.Visible Then
    msGrid.text = txtGrid.text
    DoEvents
    txtGrid.text = ""
    txtGrid.Visible = False
'    If IsNumeric(msGrid.TextMatrix(msGrid.Row, 0)) = True Then ''-Auto Bag No
'        tmpBagNo = CLng(msGrid.TextMatrix(msGrid.Row, 0)) + 1
'    Else
'        tmpBagNo = tmpBagNo
'    End If
End If
End Sub
Private Sub txtGrid_GotFocus()
    txtGrid.Alignment = ProcSetAlignment(rstblSub, msGrid.Col)
    DoEvents
    Call GProcSelectBox(txtGrid)
End Sub

Private Sub txtGrid_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyF3 Then
Select Case msGrid.Col
    Case 5 '-- Godown
        gfrmTypeStr = "S"
        GProcShowForm frmMastNarration, frmMain.mnuMstfrm(8), "S"
End Select
End If
End Sub

Private Sub txtGrid_KeyPress(KeyAscii As Integer)
Dim LeftPos As Long
Dim TopPos As Long
TopPos = txtGrid.Top + Me.Top + 650
LeftPos = txtGrid.Left + Me.Left
DoEvents
With msGrid
    If KeyAscii <> 13 And KeyAscii <> 27 Then
        '--- Search
        Select Case .Col
            Case 5  '-- Transport
                gClsSearch.SearchMultiField "tblMastNarration", "Narration", Array("Sizer Name"), Array(txtGrid.Width), " NarrType='S' ", Chr(KeyAscii), "Narration", LeftPos, TopPos, True
                txtGrid = ""
                KeyAscii = 0
                If gClsSearch.SearchMultiRetCol(0) <> "" Then
                    txtGrid = gClsSearch.SearchMultiRetCol(0)
                    txtGrid_KeyPress 13
                End If
            Case Else
                KeyAscii = GProcValidateKey(rstblSub, msGrid.Col, KeyAscii, txtGrid)
        End Select
   End If
    If KeyAscii = 13 Then
  '      CheckForNumberTxtGrid
    End If
    '--- Navigate Row ,Col
    If KeyAscii = 13 And .Col < 5 Then
        .Col = .Col + 1
    ElseIf KeyAscii = 13 And .Col = 5 Then
        If .Row = (.Rows - 1) Then
        Else
            .Row = .Row + 1
        End If
        .Col = 0
    End If
    '-- On ESC
    With msGrid
        If KeyAscii = 27 Then '----Esc Key
            KeyAscii = 0
            txtGrid.Visible = False
            Exit Sub
        End If
    End With
End With
End Sub
Private Sub CheckForNumberGrid()
Dim RowIndex As Long
With rstblSub
For RowIndex = 1 To msGrid.Rows - 1
    For i = 0 To .Fields.Count - 1
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
                    msGrid.TextMatrix(RowIndex, i) = ""  ' Null
                End If
        End Select
    Next i
  Next RowIndex
End With
End Sub
'---  Check For Number Grid
Private Sub CheckForNumberTxtGrid()
With rstblSub
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
    '    Case adChar, adVarWChar '--String
   '         txtGrid.MaxLength = .Fields(msGrid.Col).DefinedSize
        Case adDate '-- Date
            If IsDate(txtGrid) = False Then
                txtGrid = ""
            Else
                txtGrid = CDate(txtGrid)
            End If
            txtGrid.MaxLength = 10
         Case Else
            txtGrid.MaxLength = .Fields(msGrid.Col).DefinedSize
    End Select
End With
End Sub
Private Function ProcSetAlignment(prstblSub As Recordset, pColIndex As Long)
Select Case prstblSub.Fields(pColIndex).Type
    Case adInteger, adDouble, 2, 3, 131, adDecimal
        ProcSetAlignment = 1
    Case Else
        ProcSetAlignment = 0
End Select
End Function
'-- Grid Required
Private Function ValidateGrid() As Boolean
Dim i As Long
Dim RowIndex As Long
If Trim(txtGrid.text) <> "" Then msGrid.text = txtGrid.text
txtGrid.Visible = False
txtGrid.text = ""
'CheckForNumberGrid
With msGrid
Validate:
    For RowIndex = 1 To .Rows - 1
        For i = 0 To rstblSub.Fields.Count - 1
            Select Case i
                Case 3
                    If Trim(.TextMatrix(RowIndex, 3)) <> "" Then
                       If IsDate(.TextMatrix(RowIndex, 3)) = False Then
                          MsgBox "Check Deleviry Date", vbCritical + vbOKOnly, Me.Caption
                          msGrid.SetFocus
                          Exit Function
                       End If
                    End If
            End Select
        Next i
      Next RowIndex
'    '--Bag Row More than GP Bags
'    If CLng(.rows - 1) > CLng(txttmpBag) Then
'        MsgBox "Check Bag Detail.Bags " & CLng(.rows - 1) & " more than Gate Pass Bags " & CLng(txttmpBag), vbCritical + vbOKOnly, Me.Caption
'        msGrid.SetFocus
'        Exit Function
'    End If
End With
ValidateGrid = True
End Function
Private Sub SaveRelatedRecord()
Dim i As Long, X As String
Dim Ddt As Date
Dim wSizer As Long
'--- Save Grid Records
With msGrid
    i = 1
    Do While i < .Rows
        If IsDate(.TextMatrix(i, 3)) Then '- BagNo
           If .TextMatrix(i, 5) = "" Or IsNull(.TextMatrix(i, 5)) Then
              wSizer = 1
           Else
              wSizer = GProcGetColumnValue("tblMastnarration", "Narration", .TextMatrix(i, 5), "S", "NarrCode", "N")
           End If
           gCn.Execute "update tblGpSub set VDeliveryDt = cdate('" & .TextMatrix(i, 3) & "'),VehicleNo ='" & .TextMatrix(i, 4) & "',VSizer = " & wSizer & " where Vno=" & CLng(txt(0)) & " and Vtype='GP' and VSubCtrNo = " & .TextMatrix(i, 0) & "  and Vyear=" & gCYear
        Else
           gCn.Execute "update tblGpSub set VDeliveryDt = Null where Vno=" & CLng(txt(0)) & " and Vtype='GP' and VSubCtrNo = " & .TextMatrix(i, 0) & "  and Vyear=" & gCYear
           If .TextMatrix(i, 5) = "" Or IsNull(.TextMatrix(i, 5)) Then
              wSizer = 1
           Else
              wSizer = GProcGetColumnValue("tblMastnarration", "Narration", .TextMatrix(i, 5), "S", "NarrCode", "N")
           End If
           gCn.Execute "update tblGpSub set VehicleNo ='" & .TextMatrix(i, 4) & "',VSizer = " & wSizer & " where Vno=" & CLng(txt(0)) & " and Vtype='GP' and VSubCtrNo = " & .TextMatrix(i, 0) & "  and Vyear=" & gCYear
        End If
        i = i + 1
     Loop
 End With
 rstblSub.Close
 End Sub
Private Sub FillRelatedRecord()
Dim RowIndex As Long
Dim ColIndex As Long
'txtGrid.Visible = False
'--Fill Grid
SetGrid
With rstblSub
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
                    Case adDate
                        msGrid.TextMatrix(RowIndex, ColIndex) = Format(.Fields(ColIndex), "'dd/mm/yyyy")
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
End Sub
'Private Sub SaveRelatedRecord()
'Dim i As Long, X As String
'Dim Vno As Long
'Dim VYear As String
'Dim Sizer As String
'Dim VehicalNo As String
''--- Save Grid Records
'CheckForNumberGrid
'With msGrid
'    i = 1
'    Do While i < .Rows
'        Vno = CLng(.TextMatrix(i, 0)) '--Vno
'        VYear = CLng(.TextMatrix(i, 6)) '--VYear
'        VehicalNo = .TextMatrix(i, 4) '- Vehical No
'        Sizer = .TextMatrix(i, 5) '--Sizer
'        If VehicalNo = "" Then
'            gCn.Execute "update tbladdless set AdTmp5='" & Sizer & "',ADTRANSPORT='" & VehicalNo & "',ADLRDT=Null where vno=" & Vno & " and Vtype='GP' and Vyear=" & VYear
'        Else
'            If gBackEndDB = gBackEndAccess Then
'                gCn.Execute "update tbladdless set AdTmp5='" & Sizer & "',ADTRANSPORT='" & VehicalNo & "',ADLRDT='" & CDate(mskDt(0)) & "' where vno=" & Vno & " and Vtype='GP' and Vyear=" & VYear
'            Else
'                gCn.Execute "update tbladdless set AdTmp5='" & Sizer & "',ADTRANSPORT='" & VehicalNo & "',ADLRDT=to_date('" & CDate(mskDt(0)) & "','dd/MM/yyyy') where vno=" & Vno & " and Vtype='GP' and Vyear=" & VYear
'            End If
'        End If
'        i = i + 1
'     Loop
'End With
'End Sub


