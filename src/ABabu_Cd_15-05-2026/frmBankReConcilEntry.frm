VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form frmBankReConcilEntry 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00C2E3FC&
   BorderStyle     =   0  'None
   Caption         =   "Form Entry"
   ClientHeight    =   7245
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   8970
   Icon            =   "frmBankReConcilEntry.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   7245
   ScaleWidth      =   8970
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
      Index           =   0
      Left            =   1785
      MaxLength       =   40
      TabIndex        =   1
      Text            =   "txttemp(0)"
      ToolTipText     =   "Press F3 for Creation"
      Top             =   525
      Width           =   5040
   End
   Begin VB.TextBox txttemp 
      Alignment       =   1  'Right Justify
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
      Index           =   2
      Left            =   4770
      MaxLength       =   40
      TabIndex        =   26
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
      Left            =   375
      TabIndex        =   8
      Top             =   2145
      Visible         =   0   'False
      Width           =   795
   End
   Begin VB.TextBox txt 
      Height          =   345
      Index           =   17
      Left            =   9690
      TabIndex        =   25
      Text            =   "17"
      Top             =   6000
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.TextBox txt 
      Height          =   345
      Index           =   16
      Left            =   9690
      TabIndex        =   24
      Text            =   "16"
      Top             =   5520
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.TextBox txt 
      Height          =   345
      Index           =   15
      Left            =   9690
      TabIndex        =   23
      Text            =   "15"
      Top             =   5040
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.TextBox txt 
      Height          =   375
      Index           =   12
      Left            =   9600
      TabIndex        =   22
      Text            =   "12"
      Top             =   2925
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.TextBox txt 
      Height          =   325
      Index           =   10
      Left            =   9570
      TabIndex        =   21
      Text            =   "10"
      Top             =   2640
      Visible         =   0   'False
      Width           =   360
   End
   Begin VB.TextBox txt 
      Height          =   375
      Index           =   14
      Left            =   9690
      TabIndex        =   15
      Text            =   "14"
      Top             =   3720
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.TextBox txt 
      Alignment       =   1  'Right Justify
      Height          =   360
      Index           =   13
      Left            =   9570
      TabIndex        =   14
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
      Left            =   7380
      Style           =   1  'Graphical
      TabIndex        =   9
      ToolTipText     =   "To Save Record"
      Top             =   495
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
      TabIndex        =   10
      ToolTipText     =   "Exit"
      Top             =   945
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
      Left            =   9570
      Style           =   1  'Graphical
      TabIndex        =   16
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
      Left            =   9630
      Style           =   1  'Graphical
      TabIndex        =   17
      ToolTipText     =   "To Add New Record"
      Top             =   4605
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.TextBox txt 
      Height          =   345
      Index           =   3
      Left            =   9690
      TabIndex        =   20
      Text            =   "3"
      Top             =   1320
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.TextBox txt 
      Alignment       =   1  'Right Justify
      Height          =   360
      Index           =   5
      Left            =   9690
      TabIndex        =   13
      Text            =   "5"
      Top             =   2040
      Visible         =   0   'False
      Width           =   345
   End
   Begin VB.TextBox txt 
      Height          =   325
      Index           =   7
      Left            =   9570
      TabIndex        =   11
      Text            =   "7"
      Top             =   2400
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.TextBox txt 
      Alignment       =   1  'Right Justify
      Height          =   360
      Index           =   4
      Left            =   9690
      TabIndex        =   12
      Text            =   "4"
      Top             =   1680
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.TextBox txt 
      Height          =   345
      Index           =   2
      Left            =   9690
      TabIndex        =   19
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
      Left            =   9570
      TabIndex        =   18
      Text            =   "1"
      Top             =   600
      Visible         =   0   'False
      Width           =   555
   End
   Begin MSMask.MaskEdBox mskDt 
      Height          =   360
      Index           =   1
      Left            =   5265
      TabIndex        =   5
      Top             =   975
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
      Height          =   4680
      Left            =   315
      TabIndex        =   7
      Top             =   1845
      Width           =   8385
      _ExtentX        =   14790
      _ExtentY        =   8255
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
      TabIndex        =   3
      Top             =   975
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
   Begin VB.Shape Shape4 
      BorderColor     =   &H00C00000&
      BorderWidth     =   5
      Height          =   7245
      Left            =   -15
      Top             =   -15
      Width           =   8985
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
      TabIndex        =   29
      Top             =   15
      Width           =   2475
   End
   Begin VB.Label Lblform 
      BackColor       =   &H00C00000&
      Caption         =   "  Bank Reconciliation Entry"
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
      TabIndex        =   28
      Top             =   30
      Width           =   10455
   End
   Begin VB.Label lbl1 
      BackStyle       =   0  'Transparent
      Caption         =   "Reconciliation Total Amout  :"
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
      Left            =   1815
      TabIndex        =   27
      Top             =   6615
      Width           =   3075
   End
   Begin VB.Label lbl1 
      BackStyle       =   0  'Transparent
      Caption         =   "Voucher Details  :"
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
      TabIndex        =   6
      Top             =   1575
      Width           =   2175
   End
   Begin VB.Shape Shape3 
      BackColor       =   &H00808080&
      BorderColor     =   &H004D177D&
      FillColor       =   &H00ECFFFF&
      FillStyle       =   0  'Solid
      Height          =   5595
      Left            =   150
      Top             =   1485
      Width           =   8655
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
      TabIndex        =   2
      Top             =   975
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
      TabIndex        =   4
      Top             =   975
      Width           =   1095
   End
   Begin VB.Label lbl1 
      BackStyle       =   0  'Transparent
      Caption         =   "&Bank             :"
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
      Height          =   1035
      Left            =   7125
      Shape           =   4  'Rounded Rectangle
      Top             =   390
      Width           =   1650
   End
   Begin VB.Shape Shape2 
      BackColor       =   &H00808080&
      BorderColor     =   &H004D177D&
      FillColor       =   &H00ECFFFF&
      FillStyle       =   0  'Solid
      Height          =   1020
      Left            =   150
      Top             =   405
      Width           =   6855
   End
End
Attribute VB_Name = "frmBankReConcilEntry"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Rstbl As Recordset
Dim FirstTimeFlag As Boolean
Dim tmpReConDt  As String
Dim TypeStr As String
Private Sub cmdBtn_Click(Index As Integer)
On Error GoTo ErrorRoutine
Dim X As String
Select Case Index
Case 10 'Exit
    Unload Me
Case 13 'Save
    If ValidateData = True Then
       ' If Trim(txtGrid) <> "" Then
            msGrid.Text = txtGrid
            DoEvents
            txtGrid = ""
       ' End If
        txtGrid.Visible = False
        SaveRelatedRecord
        Unload Me
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
    TypeStr = gSelectedMenu
    Select Case TypeStr
        Case "RE" '--  Receipt
            Lblform.Caption = Lblform.Caption & " - Receipt"
        Case Else
            Lblform.Caption = Lblform.Caption & " - Payment"
    End Select
    '---Grid
    FirstTimeFlag = True
    SetGrid
    Call GProcClearForm(Me, Rstbl, 0, False)
    mskDt(0) = gCYSDate '-From dt
    mskDt(1) = gCYEDate '-to dt
    If FirstTimeFlag = True Then
        FirstTimeFlag = False
    Else
        txttemp(0).SetFocus
    End If
End Sub
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
If Rstbl.State = adStateOpen Then Rstbl.Close
End Sub
Private Sub mskDt_LostFocus(Index As Integer)
SaveBtnEd
If GProcIsDateValid(mskDt(Index)) = False Then mskDt(Index) = gCYSDate
If Index = 1 Then FillRelatedRecord
End Sub
Private Sub mskDt_GotFocus(Index As Integer)
GProcSelectBox Me.ActiveControl
End Sub
Private Sub SaveBtnEd()
'--Bank ,Frmdt,To dt
If txttemp(0) = "" Or IsDate(mskDt(0)) = False Or IsDate(mskDt(1)) = False Then
    cmdBtn(13).Enabled = False
    cmdBtn(13).BackColor = &H80000004
Else
    cmdBtn(13).Enabled = True
    cmdBtn(13).BackColor = &HFFFF80
End If
End Sub
'-- Grid Required
Private Function ValidateGrid() As Boolean
Dim RowIndex As Long
Dim i As Long
txtGrid.Visible = False
txtGrid.Text = ""
CheckForNumberGrid
With msGrid
    For RowIndex = 1 To .Rows - 1
        For i = 0 To Rstbl.Fields.Count - 1
            Select Case i
                Case 5 '-- Recon Date
                    If .TextMatrix(RowIndex, i) <> "" And GProcIsDateValid(.TextMatrix(RowIndex, i)) = False Then
                        MsgBox "Check Date.", vbCritical + vbOKOnly, Me.Caption
                        .Col = i
                        msGrid.SetFocus
                        Exit Function
                    End If
           End Select
        Next i
      Next RowIndex
End With
ValidateGrid = True
End Function
Private Function ValidateData() As Boolean
Dim X As String
Dim GName As String
Dim rstAddCheck As Recordset
Dim MillAgCode As Long
'---- Required Data
'--- Validate Grid
If ValidateGrid = False Then Exit Function
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
'---- Bank
If txttemp(0) = "" Then
    MsgBox "Check Bank Account.", vbInformation + vbOKOnly, Me.Caption
    txttemp(0).SetFocus
    Exit Function
End If
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
        Case 0 '-- Bank
            GName = GProcGetColumnValue("TblMastGroup", "AgCode", 90019, "N", "AgName", "S")
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
    Case 0 '--Bank
        gClsSearch.SearchMultiField "QryMastaccount", "Account,AcName,AcCity,AgName", Array("Bank", "AcName", "City", "Group"), Array(3000, 0, 2000, 3000), "(AgCode=90019)", Chr(KeyAscii), "AcName", LeftPos, TopPos
        KeyAscii = 0
        If gClsSearch.SearchMultiRetCol(1) <> "" Then txttemp(Index) = gClsSearch.SearchMultiRetCol(1)
End Select
End Sub

Private Sub txtTemp_LostFocus(Index As Integer)
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
Private Sub SetGrid()
Dim X As String
Dim BankAcCode As Long
Dim CondStr As String
Dim DtStr As String
tmpReConDt = ""
BankAcCode = GProcGetColumnValue("tblmastAccount", "AcName", txttemp(0), "S", "AcCode", "N")
CondStr = " Vtype = " & IIf(TypeStr = "RE", "'BR'", "'BP'") & " and VACOCODE = " & BankAcCode
If IsDate(mskDt(0)) = True And IsDate(mskDt(1)) = True Then
    If gBackEndDB = gBackEndAccess Then
        DtStr = " vdt>=cdate('" & mskDt(0) & "') and Vdt<=cdate('" & mskDt(1) & "')"
    Else
        DtStr = " vdt>=to_date('" & CDate(mskDt(0)) & "','dd/MM/yyyy') and Vdt<=To_date('" & CDate(mskDt(1)) & "','dd/MM/yyyy')"
    End If
Else
    DtStr = " vno=0"
End If
'--Vno ,Date,Party , Amount ,Cheque No,Recon. Date,Vtype"
If gBackEndDB = gBackEndAccess Then
    X = "select VNo,ACName,abs(VAmt),mid(VNar1,10),VDt,VReconDt,Vtype from tblVoucher,tblMastAccount where VACCODE=AcCode and " _
    & DtStr & " and " & CondStr & " order by Vdt,VReConDt,Vno"
Else
    X = "select VNo,ACName,abs(VAmt) ,Substr(VNar1,10),VDt,VReconDt,Vtype from tblVoucher,tblMastAccount where VACCODE=AcCode and " _
    & DtStr & " and " & CondStr & " order by Vdt,VReConDt,Vno"
End If
Set Rstbl = New Recordset
GProcRstOpen Rstbl, X, "R"
With msGrid
    .Clear
    .FormatString = "<Vno       |<Party                                         |>            Amount|<Cheque No.    |< Date           |<Recon. Date  |<Vtype"
    .ColWidth(6) = 0
    .Rows = 2
    'FillRelatedRecord
End With
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
    If msGrid.TextMatrix(msGrid.Row, 5) = "" Then msGrid.TextMatrix(msGrid.Row, 5) = tmpReConDt  '- ReCon Date
    If msGrid.Text <> "" Then txtGrid.Text = msGrid.Text
    '--- Lock TextBox
    With msGrid
        Select Case .Col
            Case 5 '-- Recon. Dt
                txtGrid.Locked = False
            Case Else
                txtGrid.Visible = True
                txtGrid.Locked = True
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
End With
'---- Set TxtGrid
If msGrid.Text <> "" Then txtGrid.Text = msGrid.Text
If msGrid.Col < 6 Then
    txtGrid.Move msGrid.CellLeft + msGrid.Left, msGrid.CellTop + msGrid.Top, msGrid.CellWidth, msGrid.CellHeight
End If
If msGrid.Col = 5 Then
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
Else
    txtGrid.Visible = False
    txtGrid.Enabled = False
End If
End Sub
Private Sub msGrid_LeaveCell()
If txtGrid.Visible Then
    msGrid.Text = txtGrid.Text
    DoEvents
    txtGrid.Text = ""
    txtGrid.Visible = False
    tmpReConDt = msGrid.TextMatrix(msGrid.Row, 5) '-ReCon Dt
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
        '-- Total Amt with ReCon. Date
        If IsDate(.TextMatrix(RowIndex, 5)) = True Then
            Amt = Amt + CDbl(.TextMatrix(RowIndex, 2))
        End If
        RowIndex = RowIndex + 1
    Loop
End With
txttemp(2) = Format(Amt, FStr)
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
        If .Col = 4 Then tmpReConDt = txtGrid
    ElseIf KeyAscii = 13 And .Col = 5 Then
        If .Row = (.Rows - 1) Then
        Else
            .Row = .Row + 1
        End If
        tmpReConDt = txtGrid
        .Col = 5
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
                    msGrid.TextMatrix(RowIndex, i) = Null & vbNullString
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
Private Sub SaveRelatedRecord()
Dim i As Long
Dim Vno As Long
Dim Vtype As String
Dim VYear As String
'Dim Rstbl1 As Recordset
'--- Save Grid Records
CheckForNumberGrid
With msGrid
    i = 1
    Do While i < .Rows
        Vno = CLng(.TextMatrix(i, 0)) '--Vno
        Vtype = .TextMatrix(i, 6) '--VType
        VYear = gCYear  '--VYear
        If gBackEndDB = gBackEndAccess Then
            If IsDate(.TextMatrix(i, 5)) = True Then '-- ReCon. Date
                gCn.Execute "update tblVoucher set VRECONDT=cdate('" & .TextMatrix(i, 5) & "') where vno=" & Vno & " and Vtype='" & Vtype & "' and Vyear=" & VYear
            Else
                gCn.Execute "update tblVoucher set VRECONDT=Null where vno=" & Vno & " and Vtype='" & Vtype & "' and Vyear=" & VYear
            End If
        Else '- Oracle
            If IsDate(.TextMatrix(i, 5)) = True Then '-- ReCon. Date
                gCn.Execute "update tblVoucher set VRECONDT=to_date('" & .TextMatrix(i, 5) & "','dd/MM/yyyy')  where vno=" & Vno & " and Vtype='" & Vtype & "' and Vyear=" & VYear
            Else
                gCn.Execute "update tblVoucher set VRECONDT=Null where vno=" & Vno & " and Vtype='" & Vtype & "' and Vyear=" & VYear
            End If
        End If
        i = i + 1
     Loop
End With
End Sub
Private Sub FillRelatedRecord()
Dim RowIndex As Long
Dim ColIndex As Long
tmpReConDt = ""
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
                        If .Fields(ColIndex).NumericScale > 6 Then
                            msGrid.TextMatrix(RowIndex, ColIndex) = Format(.Fields(ColIndex), GProcNumberFormat(.Fields(ColIndex).Precision, 2))
                        Else
                            msGrid.TextMatrix(RowIndex, ColIndex) = Format(.Fields(ColIndex), GProcNumberFormat(.Fields(ColIndex).Precision, .Fields(ColIndex).NumericScale))
                        End If
                    Case adDate, 135
                        If IsDate(.Fields(ColIndex)) = True Then
                            msGrid.TextMatrix(RowIndex, ColIndex) = Format(.Fields(ColIndex), "dd/MM/yyyy")
                        Else
                            msGrid.TextMatrix(RowIndex, ColIndex) = ""
                        End If
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
