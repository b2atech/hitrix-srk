VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form frmLatePayIntDbNt 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00C2E3FC&
   BorderStyle     =   0  'None
   Caption         =   "Late Pay Interest Debit Note Entry"
   ClientHeight    =   7725
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   15240
   Icon            =   "frmLatePayIntDbNt.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   7725
   ScaleWidth      =   15240
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox txt 
      Height          =   345
      Index           =   23
      Left            =   1260
      TabIndex        =   29
      Text            =   "23"
      Top             =   7065
      Visible         =   0   'False
      Width           =   9765
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
      Left            =   1890
      MaxLength       =   40
      TabIndex        =   1
      Text            =   "txttemp(0)"
      ToolTipText     =   "Press F3 for Creation"
      Top             =   585
      Width           =   5025
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
      Left            =   315
      TabIndex        =   8
      Top             =   3810
      Visible         =   0   'False
      Width           =   795
   End
   Begin VB.TextBox txt 
      Height          =   345
      Index           =   17
      Left            =   14400
      TabIndex        =   25
      Text            =   "17"
      Top             =   7065
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.TextBox txt 
      Height          =   345
      Index           =   16
      Left            =   17700
      TabIndex        =   24
      Text            =   "16"
      Top             =   5835
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.TextBox txt 
      Height          =   345
      Index           =   15
      Left            =   17700
      TabIndex        =   23
      Text            =   "15"
      Top             =   5355
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.TextBox txt 
      Height          =   375
      Index           =   12
      Left            =   17685
      TabIndex        =   22
      Text            =   "12"
      Top             =   3240
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.TextBox txt 
      Height          =   325
      Index           =   10
      Left            =   17655
      TabIndex        =   21
      Text            =   "10"
      Top             =   2955
      Visible         =   0   'False
      Width           =   360
   End
   Begin VB.TextBox txt 
      Height          =   375
      Index           =   14
      Left            =   17700
      TabIndex        =   15
      Text            =   "14"
      Top             =   4035
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.TextBox txt 
      Alignment       =   1  'Right Justify
      Height          =   360
      Index           =   13
      Left            =   17655
      TabIndex        =   14
      Text            =   "13"
      Top             =   3675
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
      Left            =   9765
      Style           =   1  'Graphical
      TabIndex        =   9
      ToolTipText     =   "To Save Record"
      Top             =   720
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
      Left            =   11475
      Style           =   1  'Graphical
      TabIndex        =   11
      ToolTipText     =   "Exit"
      Top             =   720
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
      Left            =   17655
      Style           =   1  'Graphical
      TabIndex        =   16
      ToolTipText     =   "To Add New Record"
      Top             =   4485
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
      Left            =   17640
      Style           =   1  'Graphical
      TabIndex        =   17
      ToolTipText     =   "To Add New Record"
      Top             =   4920
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.TextBox txt 
      Height          =   345
      Index           =   3
      Left            =   17700
      TabIndex        =   20
      Text            =   "3"
      Top             =   1635
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.TextBox txt 
      Alignment       =   1  'Right Justify
      Height          =   360
      Index           =   5
      Left            =   17700
      TabIndex        =   13
      Text            =   "5"
      Top             =   2355
      Visible         =   0   'False
      Width           =   345
   End
   Begin VB.TextBox txt 
      Height          =   325
      Index           =   7
      Left            =   18000
      TabIndex        =   10
      Text            =   "7"
      Top             =   2655
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.TextBox txt 
      Alignment       =   1  'Right Justify
      Height          =   360
      Index           =   4
      Left            =   17700
      TabIndex        =   12
      Text            =   "4"
      Top             =   1995
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.TextBox txt 
      Height          =   345
      Index           =   2
      Left            =   17700
      TabIndex        =   19
      Text            =   "2"
      Top             =   1275
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
      Left            =   17655
      TabIndex        =   18
      Text            =   "1"
      Top             =   915
      Visible         =   0   'False
      Width           =   555
   End
   Begin MSFlexGridLib.MSFlexGrid msGrid 
      Height          =   4665
      Left            =   180
      TabIndex        =   7
      Top             =   1980
      Width           =   14955
      _ExtentX        =   26379
      _ExtentY        =   8229
      _Version        =   393216
      Rows            =   1
      FixedCols       =   0
      RowHeightMin    =   275
      BackColorFixed  =   -2147483647
      ForeColorFixed  =   -2147483634
      BackColorBkg    =   12632256
      AllowBigSelection=   0   'False
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
      Left            =   1890
      TabIndex        =   3
      Top             =   1020
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
   Begin MSMask.MaskEdBox mskDt 
      Height          =   360
      Index           =   1
      Left            =   5340
      TabIndex        =   5
      Top             =   1020
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
      Caption         =   "Remark  :"
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
      Left            =   225
      TabIndex        =   28
      Top             =   7110
      Visible         =   0   'False
      Width           =   1080
   End
   Begin VB.Shape Shape4 
      BorderColor     =   &H00C00000&
      BorderWidth     =   5
      FillColor       =   &H00C00000&
      Height          =   7650
      Left            =   0
      Top             =   30
      Width           =   15210
   End
   Begin VB.Label Label6 
      BackColor       =   &H00C00000&
      Caption         =   "  Late Pay Interest Debit Note"
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
      Height          =   360
      Left            =   0
      TabIndex        =   27
      Top             =   15
      Width           =   17340
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
      Left            =   5460
      TabIndex        =   26
      Top             =   0
      Width           =   2475
   End
   Begin VB.Label lbl1 
      BackStyle       =   0  'Transparent
      Caption         =   "&To Date :"
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
      Index           =   0
      Left            =   4260
      TabIndex        =   4
      Top             =   1020
      Width           =   1455
   End
   Begin VB.Label lbl1 
      BackStyle       =   0  'Transparent
      Caption         =   "Interest Debit Note &Details  :"
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
      Left            =   330
      TabIndex        =   6
      Top             =   1680
      Width           =   3630
   End
   Begin VB.Shape Shape3 
      BackColor       =   &H00808080&
      BorderColor     =   &H004D177D&
      FillColor       =   &H00ECFFFF&
      FillStyle       =   0  'Solid
      Height          =   5310
      Left            =   105
      Top             =   1575
      Width           =   15030
   End
   Begin VB.Label lbl1 
      BackStyle       =   0  'Transparent
      Caption         =   "DbNt &From Date :"
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
      Height          =   510
      Index           =   10
      Left            =   240
      TabIndex        =   2
      Top             =   1020
      Width           =   1635
   End
   Begin VB.Label lbl1 
      BackStyle       =   0  'Transparent
      Caption         =   "&Party     :"
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
      Left            =   240
      TabIndex        =   0
      Top             =   585
      Width           =   1440
   End
   Begin VB.Shape Shape1 
      FillColor       =   &H00BDDFD1&
      FillStyle       =   0  'Solid
      Height          =   885
      Left            =   9090
      Shape           =   4  'Rounded Rectangle
      Top             =   495
      Width           =   4005
   End
   Begin VB.Shape Shape2 
      BackColor       =   &H00808080&
      BorderColor     =   &H004D177D&
      FillColor       =   &H00ECFFFF&
      FillStyle       =   0  'Solid
      Height          =   1065
      Left            =   105
      Top             =   435
      Width           =   7125
   End
End
Attribute VB_Name = "frmLatePayIntDbNt"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Rstbl As Recordset
Dim FirstTimeFlag As Boolean
Private Sub cmdBtn_Click(Index As Integer)
'On Error GoTo ErrorRoutine
Dim X As String
Select Case Index
Case 10 'Exit
    Unload Me
Case 13 'Save
    If ValidateData = True Then
       ' If Trim(txtGrid) <> "" Then
            msGrid.text = txtGrid
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
            Sendkeys "{TAB}"
        Case 27:
            If (cmdBtn(13).Enabled = False) Then
                Unload Me
            End If
    End Select
End Sub
Private Sub Form_Load()
    Me.Top = 300
    Me.Left = Screen.Width / 2 - Me.Width / 2
    '---Grid
    FirstTimeFlag = True
    SetGrid
    Call GProcClearForm(Me, Rstbl, 0, False)
    If FirstTimeFlag = True Then
        FirstTimeFlag = False
    Else
        txtTemp(0).SetFocus
    End If
    mskDt(0) = gCYSDate
    mskDt(1) = gCYEDate
End Sub
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
If Rstbl.State = adStateOpen Then Rstbl.Close
End Sub

Private Sub mskDt_LostFocus(Index As Integer)
SaveBtnEd
'If IsDate(mskDt(Index)) = False Then mskDt(Index) = gCYSDate
FillRelatedRecord
End Sub
Private Sub mskDt_GotFocus(Index As Integer)
GProcSelectBox Me.ActiveControl
End Sub
Private Sub SaveBtnEd()
'--Party
If txtTemp(0) = "" Then 'Or IsDate(mskDt(0)) = False Or IsDate(mskDt(1)) = False
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
'If ValidateGrid = False Then Exit Function
'--- From dt
If IsDate(mskDt(0)) = False Then
    MsgBox "Check From Date.", vbInformation + vbOKOnly, Me.Caption
    mskDt(0).SetFocus
    Exit Function
End If
'--- To dt
If IsDate(mskDt(1)) = False Then
    MsgBox "Check To Date.", vbInformation + vbOKOnly, Me.Caption
    mskDt(1).SetFocus
    Exit Function
End If
'---- Party
If txtTemp(0) = "" Then
    MsgBox "Check Party.", vbInformation + vbOKOnly, Me.Caption
    txtTemp(0).SetFocus
    Exit Function
End If
ValidateData = True
End Function
Private Sub txttemp_GotFocus(Index As Integer)
GProcSelectBox Me.ActiveControl
End Sub
Private Sub txttemp_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyF3 Then
    Select Case Index
        Case 0 '-- Account
            GProcShowForm frmMastAccount, frmMain.mnuMstfrm(0)
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
        gClsSearch.SearchMultiField "QryMastaccount", "Account,AcName,AcCity,AgName", Array("Party", "AcName", "City", "Group"), Array(3000, 0, 2000, 3000), " (Agcode=90017 or Agcode=90015 or Agcode=90029)", Chr(KeyAscii), "acName", LeftPos, TopPos
        KeyAscii = 0
        If gClsSearch.SearchMultiRetCol(1) <> "" Then txtTemp(Index) = gClsSearch.SearchMultiRetCol(1)
End Select
End Sub
Private Sub txtTemp_LostFocus(Index As Integer)
SaveBtnEd
If Index = 0 Then FillRelatedRecord
End Sub
Private Sub SetGrid()
Dim X As String
Dim AcCode As Long
Dim DtStr As String
'-- Date Str
If IsDate(mskDt(0)) = True And IsDate(mskDt(1)) = True Then
    If gBackEndDB = gBackEndAccess Then
        DtStr = " and INTDBNTDATE>=cdate('" & mskDt(0) & "') and INTDBNTDATE<=cdate('" & mskDt(1) & "') "
    Else
        DtStr = " and INTDBNTDATE>=to_date('" & CDate(mskDt(0)) & "','dd/MM/yyyy') and INTDBNTDATE<=to_date('" & CDate(mskDt(1)) & "','dd/MM/yyyy') "
    End If
End If
'--Party
AcCode = GProcGetColumnValue("tblmastAccount", "AcName", txtTemp(0), "S", "AcCode", "N")
'-- Int DbNt Date,Int DbNt Amt,IntDb Nt Recd,Int Db Nt Complete (Y/N),AdjNo
X = "select  RecVno,format(RecVDt,'dd/mm/yyyy'),InvBillNo,format(InvDt,'dd/mm/yyyy'),BillAmt,AdjAmt,IntDbNtDate,IntDbNtAmt,IntReceived,IntDbNtRecdAmt,IntDbNtAmt-IntReceived-IntDbNtRecdAmt," & IIf(gBackEndDB = gBackEndAccess, "iif(IsIntDbNtComplete=1,'Y','')", "decode(IsIntDbNtComplete,1,'Y','')") & " ,AdjNo from tblRecVsSale where " _
& " PARTYCODE =" & AcCode & DtStr & " and RECVTYPE in ('BR','CR','MR')  order by IntDbNtDate,RECVDT,INVDT"
Set Rstbl = New Recordset
GProcRstOpen Rstbl, X, "R"
With msGrid
    .Clear
    .FormatString = ">Rec.No  |<Rec. Date      |>Invoice No    |< Inv.Date       |> Inv. Amount      |> Recd. Amount         |< Db.Nt.Date  |>    Db.Nt.Amt.|>Int.Recd. In Rec. |>Received Amt.|>Bal.Int.Db.Nt|<Is Comp.(Y)    |<AdjNo"
    .ColWidth(12) = 0
    .Rows = 2
    'FillRelatedRecord
End With
End Sub
Private Sub msGrid_Click()
On Error Resume Next
    msGrid_EnterCell
End Sub
Private Sub msGrid_GotFocus()
If LCase(Me.ActiveControl) <> "txtgrid" Then msGrid.Col = 9
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
    If Trim(msGrid.text) <> "" Then txtGrid.text = msGrid.text
    '--- Lock TextBox
    With msGrid
        Select Case .Col
            Case 0, 1, 2, 3, 4, 5, 6, 7, 8, 10 '--Date,Amount
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
        If .Col = 10 Then
            If .Row = (.Rows - 1) Then '-- Add New Row
            Else '--Next Row
                .Row = .Row + 1
            End If
            .Col = 9
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
If msGrid.text <> "" Then txtGrid.text = msGrid.text
If msGrid.Col < 12 Then
    txtGrid.Move msGrid.CellLeft + msGrid.Left, msGrid.CellTop + msGrid.Top, msGrid.CellWidth, msGrid.CellHeight
End If
txtGrid.Visible = True
txtGrid.Enabled = True
DoEvents
If KeyAscii <> 0 Then
    Select Case Rstbl.Fields(msGrid.Col).Type
        Case adChar, adVarWChar
            txtGrid.text = txtGrid.text & Chr(KeyAscii)
        Case Else
            txtGrid.text = Chr(GProcValidateKey(Rstbl, msGrid.Col, KeyAscii, txtGrid.text)) 'Chr(GProcNumberWithDecimal(KeyAscii, txtGrid))
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
    msGrid.text = Trim(txtGrid.text)
    DoEvents
    txtGrid.text = ""
    txtGrid.Visible = False
    'tmpFormNo = msGrid.TextMatrix(msGrid.Row, 4)
End If
DoEvents
'If msGrid.Col = 4 Then CalGridAmt
End Sub
Private Sub txtGrid_GotFocus()
    txtGrid.Alignment = ProcSetAlignment(Rstbl, msGrid.Col)
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
        Select Case .Col
            Case 0, 1, 2, 3, 4, 5, 6, 7, 8, 10
                KeyAscii = 0
            Case 11 '- Y/N
                If UCase(Chr(KeyAscii)) = "Y" Then 'Y
                    txtGrid = ""
                    KeyAscii = Asc("Y")
                Else
                    txtGrid = ""
                    KeyAscii = 0
                End If
             Case Else
                KeyAscii = GProcValidateKey(Rstbl, msGrid.Col, KeyAscii, txtGrid)
        End Select
   End If
    If KeyAscii = 13 Then
        CheckForNumberTxtGrid
    End If
    '--- Navigate Row ,Col
    If KeyAscii = 13 And .Col < 12 Then
        .Col = .Col + 1
    ElseIf KeyAscii = 13 And .Col = 12 Then
        If .Row = (.Rows - 1) Then
        Else
            .Row = .Row + 1
        End If
        .Col = 9
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
                If .Fields(msGrid.Col).NumericScale > 6 Then
                    txtGrid = Format(0, GProcNumberFormat(.Fields(msGrid.Col).Precision, 2))
                Else
                    txtGrid = Format(0, GProcNumberFormat(.Fields(msGrid.Col).Precision, .Fields(msGrid.Col).NumericScale))
                End If
            Else
                If .Fields(msGrid.Col).NumericScale > 6 Then
                    txtGrid = Format(txtGrid, GProcNumberFormat(.Fields(msGrid.Col).Precision, 2))
                Else
                    txtGrid = Format(txtGrid, GProcNumberFormat(.Fields(msGrid.Col).Precision, .Fields(msGrid.Col).NumericScale))
                End If
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
Dim AdjNo As Long
Dim AmtRecd As Double
Dim IsComplete As Long
'--- Save Grid Records
CheckForNumberGrid
With msGrid
    i = 1
    Do While i < .Rows
        AdjNo = CLng(.TextMatrix(i, 12)) '--AdjNo
        AmtRecd = CDbl(.TextMatrix(i, 9)) '--Recd Amt
        IsComplete = IIf(Trim(UCase(.TextMatrix(i, 11))) = "Y", 1, 0) '- Is Complete
        '-- Update Recd Amt ,Is Complete
        gCn.Execute "update tblRecVsSale set IntDbNtRecdAmt=" & AmtRecd & ", " _
        & " IsIntDbNtComplete=" & IsComplete & " where Adjno=" & AdjNo
        i = i + 1
     Loop
End With
End Sub
Private Sub FillRelatedRecord()
Dim RowIndex As Long
Dim ColIndex As Long
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
