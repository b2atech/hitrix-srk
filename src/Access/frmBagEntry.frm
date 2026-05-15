VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form frmBagEntry 
   BackColor       =   &H00C2E3FC&
   BorderStyle     =   0  'None
   Caption         =   "Goods Inward"
   ClientHeight    =   6075
   ClientLeft      =   0
   ClientTop       =   15
   ClientWidth     =   8550
   FillColor       =   &H000040C0&
   FillStyle       =   0  'Solid
   Icon            =   "frmBagEntry.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6075
   ScaleWidth      =   8550
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox txtBalWt 
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
      Left            =   5220
      MaxLength       =   40
      TabIndex        =   35
      Text            =   "txtBalWt"
      Top             =   5265
      Width           =   1560
   End
   Begin VB.TextBox txtTotWt 
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
      Left            =   5220
      MaxLength       =   40
      TabIndex        =   5
      Text            =   "txtTotWt"
      Top             =   4845
      Width           =   1560
   End
   Begin VB.TextBox txttmpWt 
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
      Left            =   2325
      MaxLength       =   40
      TabIndex        =   7
      Text            =   "txttmpWt"
      Top             =   5265
      Width           =   1500
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
      Height          =   375
      Left            =   2325
      MaxLength       =   40
      TabIndex        =   3
      Text            =   "txttmpBag"
      Top             =   4845
      Width           =   1500
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
      Left            =   7185
      Style           =   1  'Graphical
      TabIndex        =   9
      ToolTipText     =   "Exit"
      Top             =   3270
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
      Left            =   7185
      Style           =   1  'Graphical
      TabIndex        =   8
      ToolTipText     =   "To Save Record"
      Top             =   2760
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
      Height          =   330
      Left            =   630
      TabIndex        =   1
      Top             =   1485
      Visible         =   0   'False
      Width           =   1035
   End
   Begin VB.TextBox txt 
      Alignment       =   1  'Right Justify
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
      Height          =   375
      Index           =   1
      Left            =   10875
      TabIndex        =   33
      Text            =   "1"
      Top             =   2430
      Width           =   420
   End
   Begin VB.TextBox txt 
      Alignment       =   1  'Right Justify
      Height          =   315
      Index           =   2
      Left            =   11040
      TabIndex        =   22
      Text            =   "2"
      Top             =   1440
      Visible         =   0   'False
      Width           =   525
   End
   Begin VB.TextBox txt 
      Alignment       =   1  'Right Justify
      Height          =   315
      Index           =   5
      Left            =   11040
      TabIndex        =   21
      Text            =   "5"
      Top             =   2640
      Width           =   525
   End
   Begin VB.TextBox txt 
      Alignment       =   1  'Right Justify
      Height          =   315
      Index           =   3
      Left            =   11040
      TabIndex        =   20
      Text            =   "3"
      Top             =   1800
      Width           =   525
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
      Left            =   8550
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   8055
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
      Left            =   9315
      Style           =   1  'Graphical
      TabIndex        =   10
      ToolTipText     =   "To Add New Record"
      Top             =   7980
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.TextBox txt 
      Alignment       =   1  'Right Justify
      Height          =   315
      Index           =   4
      Left            =   11040
      TabIndex        =   19
      Text            =   "4"
      Top             =   2160
      Width           =   525
   End
   Begin VB.TextBox txt 
      Alignment       =   1  'Right Justify
      Height          =   315
      Index           =   7
      Left            =   11040
      TabIndex        =   18
      Text            =   "7"
      Top             =   3360
      Visible         =   0   'False
      Width           =   525
   End
   Begin VB.TextBox txt 
      Alignment       =   1  'Right Justify
      Height          =   315
      Index           =   9
      Left            =   11040
      TabIndex        =   17
      Text            =   "9"
      Top             =   3720
      Width           =   525
   End
   Begin VB.TextBox txt 
      Alignment       =   1  'Right Justify
      Height          =   315
      Index           =   12
      Left            =   11040
      TabIndex        =   16
      Text            =   "12"
      Top             =   4080
      Width           =   525
   End
   Begin VB.TextBox txt 
      Alignment       =   1  'Right Justify
      Height          =   315
      Index           =   20
      Left            =   11040
      TabIndex        =   15
      Text            =   "20"
      Top             =   4440
      Visible         =   0   'False
      Width           =   525
   End
   Begin VB.TextBox txt 
      Alignment       =   1  'Right Justify
      Height          =   315
      Index           =   21
      Left            =   11040
      TabIndex        =   14
      Text            =   "21"
      Top             =   4800
      Visible         =   0   'False
      Width           =   525
   End
   Begin VB.TextBox txt 
      Alignment       =   1  'Right Justify
      Height          =   315
      Index           =   22
      Left            =   11040
      TabIndex        =   13
      Text            =   "22"
      Top             =   5160
      Visible         =   0   'False
      Width           =   525
   End
   Begin VB.TextBox txt 
      Alignment       =   1  'Right Justify
      Height          =   315
      Index           =   23
      Left            =   11040
      TabIndex        =   12
      Text            =   "23"
      Top             =   5400
      Visible         =   0   'False
      Width           =   525
   End
   Begin Crystal.CrystalReport cryReport 
      Bindings        =   "frmBagEntry.frx":058A
      Left            =   0
      Top             =   600
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      WindowControlBox=   -1  'True
      WindowMaxButton =   -1  'True
      WindowMinButton =   -1  'True
      PrintFileType   =   15
      PrintFileLinesPerPage=   60
   End
   Begin MSFlexGridLib.MSFlexGrid msGrid 
      Height          =   4065
      Left            =   255
      TabIndex        =   0
      Top             =   615
      Width           =   6525
      _ExtentX        =   11509
      _ExtentY        =   7170
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
   Begin VB.Image Image1 
      Height          =   1545
      Left            =   7020
      Picture         =   "frmBagEntry.frx":059B
      Stretch         =   -1  'True
      Top             =   810
      Width           =   1395
   End
   Begin VB.Label lbl1 
      BackStyle       =   0  'Transparent
      Caption         =   "Balance Wt :"
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
      Index           =   3
      Left            =   3990
      TabIndex        =   36
      Top             =   5265
      Width           =   1275
   End
   Begin VB.Label lbl1 
      BackStyle       =   0  'Transparent
      Caption         =   "Total Wt :"
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
      Index           =   2
      Left            =   3990
      TabIndex        =   4
      Top             =   4845
      Width           =   1200
   End
   Begin VB.Label lbl1 
      BackStyle       =   0  'Transparent
      Caption         =   "Total Inward Wt :"
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
      Index           =   1
      Left            =   390
      TabIndex        =   6
      Top             =   5265
      Width           =   1740
   End
   Begin VB.Label lbl1 
      BackStyle       =   0  'Transparent
      Caption         =   "Total Inward Bags :"
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
      Left            =   390
      TabIndex        =   2
      Top             =   4845
      Width           =   2070
   End
   Begin VB.Shape Shape1 
      FillColor       =   &H00BDDFD1&
      FillStyle       =   0  'Solid
      Height          =   1410
      Left            =   7050
      Shape           =   4  'Rounded Rectangle
      Top             =   2505
      Width           =   1350
   End
   Begin VB.Shape Shape5 
      BorderColor     =   &H00C00000&
      BorderWidth     =   5
      FillColor       =   &H00000080&
      Height          =   6030
      Left            =   0
      Top             =   30
      Width           =   8550
   End
   Begin VB.Shape Shape3 
      BorderColor     =   &H00004080&
      FillColor       =   &H00000080&
      FillStyle       =   0  'Solid
      Height          =   1605
      Left            =   0
      Top             =   6135
      Width           =   10935
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
      Left            =   5505
      TabIndex        =   34
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
      TabIndex        =   32
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
      TabIndex        =   31
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
      TabIndex        =   30
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
      TabIndex        =   29
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
      TabIndex        =   28
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
      TabIndex        =   27
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
      TabIndex        =   26
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
      TabIndex        =   25
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
      TabIndex        =   24
      Top             =   8160
      Width           =   1095
   End
   Begin VB.Label Label6 
      BackColor       =   &H00C00000&
      Caption         =   " Cartoon Bag Entry"
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
      Left            =   30
      TabIndex        =   23
      Top             =   15
      Width           =   10815
   End
   Begin VB.Shape Shape2 
      BackColor       =   &H00808080&
      BorderColor     =   &H004D177D&
      FillColor       =   &H00EDFDFE&
      FillStyle       =   0  'Solid
      Height          =   5355
      Left            =   150
      Top             =   465
      Width           =   6795
   End
End
Attribute VB_Name = "frmBagEntry"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rstblSub As Recordset '-- TblBags
Dim FirstTimeFlag As Boolean
Dim tmpBagNo As String
Dim i As Long

Private Sub cmdBtn_Click(Index As Integer)
On Error GoTo ErrorRoutine
Dim x As String
Select Case Index
Case 10 'Exit
    Unload Me
Case 13 'Save
    If msGrid.Rows = 1 Then Exit Sub
    If txtGrid <> "" And txtGrid.Visible = True Then msGrid.Text = txtGrid
    CalAmount
    txtGrid = ""
    If ValidateData = True Then
         txtGrid.Visible = False
         SaveRelatedRecord
         Unload Me
    End If
End Select
Exit Sub
ErrorRoutine:
    'rstblSub.CancelUpdate
    GProcErrorHandler
End Sub
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case 13:
           If Me.ActiveControl.Name <> "txtGrid" Then SendKeys "{TAB}"
        Case 27:
            If (cmdBtn(13).Enabled = False) Then
                Unload Me
            End If
    End Select
End Sub
Private Sub Form_Load()
    Me.Top = 300
    Me.Left = Screen.Width / 2 - Me.Width / 2
    txttmpBag.Enabled = False 'In Bags
    txttmpWt.Enabled = False 'In Wt
    txtTotWt.Enabled = False 'Tot Wt
    txtBalWt.Enabled = False 'Bal Wt
    FirstTimeFlag = True
    Call GProcClearForm(Me, rstblSub, 0, False)
     '---Grid
    FillRelatedRecord
    If FirstTimeFlag = True Then
        FirstTimeFlag = False
    Else
        msGrid.SetFocus
    End If
End Sub
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
If rstblSub.State = adStateOpen Then rstblSub.Close
End Sub
Private Function ValidateData() As Boolean
'---- Required Data
'--- Validate Grid
If ValidateGrid = False Then Exit Function
If CDbl(txtTotWt) < (txttmpWt) Then
    MsgBox "Total  Wt is not equal to Inward /Purchase Wt " & txttmpWt, vbInformation + vbOKOnly, Me.Caption
    msGrid.SetFocus
    Exit Function
End If
ValidateData = True
End Function

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
    If msGrid.TextMatrix(msGrid.Row, 0) = "" And IsNumeric(tmpBagNo) = True Then msGrid.TextMatrix(msGrid.Row, 0) = tmpBagNo  ''-Auto Bag No
    If Trim(msGrid.Text) <> "" Then txtGrid.Text = msGrid.Text
    '--- Lock TextBox
    With msGrid
        Select Case .Col
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
    If KeyAscii = 13 Then
        KeyAscii = 0
      With msGrid
        If .Col = 2 Then
            If .Row = (.Rows - 1) Then '-- Add New Row
            Else '--Next Row
                .Row = .Row + 1
            End If
            .Col = 0
        Else    '-- Next col
            .Col = .Col + 1
        End If
        End With
    End If
'    If KeyAscii <> 13 And KeyAscii <> 27 And .Col = 0 Then
'        DoEvents
'        KeyAscii = 0
'        Pending_In
'    End If
'ProcSetAlignment rstblSub, msGrid.Col
'---- Set TxtGrid
If msGrid.Text <> "" Then txtGrid.Text = msGrid.Text
If msGrid.Col < 3 Then
    txtGrid.Move msGrid.CellLeft + msGrid.Left, msGrid.CellTop + msGrid.Top, msGrid.CellWidth, msGrid.CellHeight
End If
txtGrid.Visible = True
txtGrid.Enabled = True
DoEvents
If KeyAscii <> 0 Then
    Select Case rstblSub.Fields(msGrid.Col).Type
        Case adChar, adVarWChar
            txtGrid.Text = txtGrid.Text & Chr(KeyAscii)
        Case Else
            txtGrid.Text = Chr(GProcValidateKey(rstblSub, msGrid.Col, KeyAscii, txtGrid.Text)) 'Chr(GProcNumberWithDecimal(KeyAscii, txtGrid))
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
    CalAmount
    If IsNumeric(msGrid.TextMatrix(msGrid.Row, 0)) = True Then ''-Auto Bag No
        tmpBagNo = CLng(msGrid.TextMatrix(msGrid.Row, 0)) + 1
    Else
        tmpBagNo = tmpBagNo
    End If
End If
End Sub

Private Sub MSGrid_LostFocus()
CalAmount
End Sub

Private Sub txtGrid_GotFocus()
    txtGrid.Alignment = ProcSetAlignment(rstblSub, msGrid.Col)
    DoEvents
    Call GProcSelectBox(txtGrid)
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
            Case Else
                KeyAscii = GProcValidateKey(rstblSub, msGrid.Col, KeyAscii, txtGrid)
        End Select
   End If
    If KeyAscii = 13 Then
  '      CheckForNumberTxtGrid
    End If
    '--- Navigate Row ,Col
    If KeyAscii = 13 And .Col < 2 Then
        .Col = .Col + 1
        If .Col = 0 Then
            If IsNumeric(txtGrid) = True Then '-Auto Bag No
                tmpBagNo = CLng(txtGrid)
            Else
                tmpBagNo = ""
            End If
        End If
    ElseIf KeyAscii = 13 And .Col = 2 Then
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
                    msGrid.TextMatrix(RowIndex, i) = Null
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
Dim BagNo As String
If Trim(txtGrid.Text) <> "" Then msGrid.Text = txtGrid.Text
txtGrid.Visible = False
txtGrid.Text = ""
CheckForNumberGrid
With msGrid
Validate:
    For RowIndex = 1 To .Rows - 1
        For i = 0 To rstblSub.Fields.Count - 1
            Select Case i
                Case 0 '-- Bag  No
                   
                    '-Bag No="" ,wt=0  and Bags row >In Bag delete row
                    If RowIndex > 0 And Trim(.TextMatrix(RowIndex, i)) = "" And CDbl(.TextMatrix(RowIndex, 2)) = 0 And CLng(.Rows - 1) > CLng(txttmpBag) Then
                        .RemoveItem (RowIndex)
                        GoTo Validate
                    End If
                    If Trim(.TextMatrix(RowIndex, i)) = "" Then
                        MsgBox "Check Bag No.", vbCritical + vbOKOnly, Me.Caption
                        .Col = i
                        msGrid.SetFocus
                        Exit Function
                    End If
              '  Case 1 '-- Cone
              '      If CDbl(.TextMatrix(RowIndex, i)) = 0 Then
'                        MsgBox "Check Cone.", vbCritical + vbOKOnly, Me.Caption
'                        .Col = i
'                        msGrid.SetFocus
'                        Exit Function
              '      End If
                Case 2 '-- Wt
                    If CDbl(.TextMatrix(RowIndex, i)) = 0 Then
                        MsgBox "Check Weight.", vbCritical + vbOKOnly, Me.Caption
                        .Col = i
                        msGrid.SetFocus
                        Exit Function
                    End If
           End Select
        Next i
         '-- duplicate Bag No
        BagNo = .TextMatrix(RowIndex, 0) '-Bag No
        For i = 1 To .Rows - 1
            If i <> RowIndex And .TextMatrix(i, 0) = BagNo Then
                MsgBox "Duplicate BagNo.", vbCritical + vbOKOnly, Me.Caption
                .Col = 0
                msGrid.SetFocus
                Exit Function
            End If
        Next i
      Next RowIndex
    '--Bag Row More than In Bags
    If CLng(.Rows - 1) > CLng(txttmpBag) Then
        MsgBox "Check Bag Detail.Bags " & CLng(.Rows - 1) & " more than Inward /Purchase Bags " & CLng(txttmpBag), vbCritical + vbOKOnly, Me.Caption
        msGrid.SetFocus
        Exit Function
    End If
End With
ValidateGrid = True
End Function
Private Sub SaveRelatedRecord()
Dim i As Long, x As String
'--- Save Grid Records
CheckForNumberGrid
'--- Save Grid Records To tblBags
gCn.Execute "delete from tblBags where InwType='" & gfrmBagInwVType & "' and InwNo=" & gfrmBagInwVNo
x = "select * from tblBags where InwType='" & gfrmBagInwVType & "' and InwNo=" & gfrmBagInwVNo
GProcRstOpen rstblSub, x, "O"
With msGrid
    i = 1
    Do While i < .Rows
       If .TextMatrix(i, 0) <> "" Then '- BagNo
            rstblSub.AddNew
            '-- SNo
            If CLng(.TextMatrix(i, 7)) > 0 Then
                rstblSub!SNo = .TextMatrix(i, 7)
            Else
                rstblSub!SNo = GProcGenerateId("tblBags", "SNo")
            End If
            rstblSub!BagNo = .TextMatrix(i, 0) '--BagNo
            rstblSub!Cone = CLng(.TextMatrix(i, 1)) '-- Cone
            rstblSub!Wt = CDbl(.TextMatrix(i, 2)) '--Wt
            rstblSub!InwNo = gfrmBagInwVNo    '--Inw No
            rstblSub!InwType = gfrmBagInwVType    '--Inw Type
            rstblSub!GpNo = CLng(.TextMatrix(i, 5))    '--GP No
            rstblSub!GPTYPE = .TextMatrix(i, 6) '-- GP Type
            rstblSub.Update
        End If
        i = i + 1
     Loop
 End With
 rstblSub.Close
 End Sub
Private Sub CalAmount()
Dim RowIndex  As Long
Dim TotWt As Double
CheckForNumberGrid
RowIndex = 1
With msGrid
    Do While RowIndex < .Rows
        '-- Wt
        If Val(.TextMatrix(RowIndex, 2)) > 0 Then
            TotWt = TotWt + Val(.TextMatrix(RowIndex, 2))
        End If
        RowIndex = RowIndex + 1
    Loop
End With
TotWt = Format(TotWt, WtStr) '-Total Wt
txtTotWt = TotWt
txtBalWt = CDbl(txttmpWt) - CDbl(txtTotWt) '- Balance Wt
txtBalWt = Format(txtBalWt, WtStr)
End Sub
Private Sub FillRelatedRecord()
Dim x As String
Dim RCount As Integer
Dim RsIn As Recordset
Dim RowIndex As Integer
Dim ColIndex As Integer
txttmpBag = "0"
txttmpWt = "0.000"
tmpBagNo = "" ''-Auto Bag No
'--Cartoon Inward Total Bags,Wt
x = "select sum(VSubBag),sum(VSubWt) from tblvousub,TblMastItem " _
& " where VSUBITCODE=Itcode and ITUNIT='Cartoon' and Vno=" & gfrmBagInwVNo & " and vtype='" & gfrmBagInwVType & "' and VYear=" & gCYear

Set RsIn = New Recordset
If GProcRstOpen(RsIn, x, "R") > 0 Then
    txttmpBag = IIf(IsNumeric(RsIn.Fields(0)), RsIn.Fields(0), 0)
    txttmpWt = IIf(IsNumeric(RsIn.Fields(1)), RsIn.Fields(1), 0)
End If
txttmpBag = Format(txttmpBag, PStr)
txttmpWt = Format(txttmpWt, WtStr)
'--BagNo,cone,Wt,InwNo,InwType,GpNo,GpType
x = "select BagNo,cone,Wt,InwNo,InwType,GpNo,GpType,Sno from tblBags where " _
& " InwType='" & gfrmBagInwVType & "' and InwNo=" & gfrmBagInwVNo & " order by SNo"
Set rstblSub = New Recordset
RCount = GProcRstOpen(rstblSub, x, "R")
With msGrid
    .Clear
    .FormatString = "<Bag No.                 |>Cone       |>Weight       |InwNo |InwType|GpNo |GpType|SNo"
    .ColWidth(3) = 0 '-Inw No
    .ColWidth(4) = 0 '-Inw Type
    .ColWidth(5) = 0 '-Gp No
    .ColWidth(6) = 0 '-Gp Type
    .ColWidth(7) = 0 '-SNo
    If CLng(RCount) > CLng(txttmpBag) Then '--No Of Rows = No. Of Bags
        .Rows = RCount + 1
    Else
        .Rows = txttmpBag + 1
    End If
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
CalAmount
End Sub
