VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form frmMultiAccount 
   BackColor       =   &H00C2E3FC&
   BorderStyle     =   0  'None
   Caption         =   "t"
   ClientHeight    =   6870
   ClientLeft      =   0
   ClientTop       =   15
   ClientWidth     =   6405
   DrawMode        =   10  'Mask Pen
   FillColor       =   &H000040C0&
   FillStyle       =   0  'Solid
   Icon            =   "frmMultiAccount.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6870
   ScaleWidth      =   6405
   ShowInTaskbar   =   0   'False
   Begin VB.CheckBox ChkAcConf 
      BackColor       =   &H00C0FFC0&
      Caption         =   "Is Account Confirmation ?"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C000C0&
      Height          =   330
      Left            =   1620
      TabIndex        =   14
      Top             =   5535
      Width           =   2940
   End
   Begin VB.TextBox txt 
      Height          =   315
      Index           =   0
      Left            =   10500
      TabIndex        =   12
      Text            =   "0"
      Top             =   1380
      Width           =   315
   End
   Begin VB.CommandButton cmdBtn 
      Caption         =   "&Print"
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
      Index           =   2
      Left            =   3300
      Style           =   1  'Graphical
      TabIndex        =   3
      ToolTipText     =   "To Save Record"
      Top             =   6150
      Width           =   1215
   End
   Begin VB.TextBox txtGrid 
      BackColor       =   &H8000000A&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   2640
      TabIndex        =   7
      Top             =   1320
      Visible         =   0   'False
      Width           =   1035
   End
   Begin MSFlexGridLib.MSFlexGrid msGrid 
      Height          =   4215
      Left            =   375
      TabIndex        =   2
      Top             =   1140
      Width           =   5685
      _ExtentX        =   10028
      _ExtentY        =   7435
      _Version        =   393216
      FixedCols       =   0
      RowHeightMin    =   275
      BackColorFixed  =   -2147483647
      ForeColorFixed  =   -2147483634
      BackColorBkg    =   12632256
      AllowUserResizing=   1
      Appearance      =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
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
      Index           =   4
      Left            =   4620
      Style           =   1  'Graphical
      TabIndex        =   6
      ToolTipText     =   "Exit"
      Top             =   6150
      Width           =   1215
   End
   Begin VB.CommandButton cmdBtn 
      Caption         =   "&Screen"
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
      Index           =   1
      Left            =   1920
      Style           =   1  'Graphical
      TabIndex        =   5
      ToolTipText     =   "To Modify Record"
      Top             =   6150
      Width           =   1215
   End
   Begin VB.CommandButton cmdBtn 
      Caption         =   "&Add"
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
      Index           =   0
      Left            =   540
      Style           =   1  'Graphical
      TabIndex        =   4
      ToolTipText     =   "To Add New Record"
      Top             =   6150
      Width           =   1215
   End
   Begin Crystal.CrystalReport cryReport 
      Bindings        =   "frmMultiAccount.frx":058A
      Left            =   2820
      Top             =   30
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      WindowControlBox=   -1  'True
      WindowMaxButton =   -1  'True
      WindowMinButton =   -1  'True
      PrintFileType   =   15
      PrintFileLinesPerPage=   60
   End
   Begin MSComCtl2.DTPicker dtpToDt 
      Height          =   360
      Left            =   4455
      TabIndex        =   1
      Top             =   600
      Width           =   1635
      _ExtentX        =   2884
      _ExtentY        =   635
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
      Format          =   101056515
      CurrentDate     =   38050
   End
   Begin MSComCtl2.DTPicker dtpFrDt 
      Height          =   360
      Left            =   1245
      TabIndex        =   0
      Top             =   600
      Width           =   1635
      _ExtentX        =   2884
      _ExtentY        =   635
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
      Format          =   101056515
      CurrentDate     =   38050
   End
   Begin Crystal.CrystalReport CrystalReport1 
      Left            =   7440
      Top             =   300
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
   Begin VB.Label Label3 
      BackColor       =   &H00C00000&
      Caption         =   "   Multi Account Ledger"
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
      Height          =   375
      Left            =   0
      TabIndex        =   13
      Top             =   0
      Width           =   11835
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "From   :"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   240
      Index           =   0
      Left            =   300
      TabIndex        =   11
      Top             =   600
      Width           =   735
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
      ForeColor       =   &H00800000&
      Height          =   240
      Left            =   3720
      TabIndex        =   10
      Top             =   660
      Width           =   540
   End
   Begin VB.Shape Shape8 
      BorderColor     =   &H00C00000&
      BorderWidth     =   5
      Height          =   6825
      Left            =   0
      Top             =   0
      Width           =   6345
   End
   Begin VB.Shape Shape3 
      BorderColor     =   &H00C00000&
      FillColor       =   &H00C00000&
      FillStyle       =   0  'Solid
      Height          =   0
      Left            =   -15
      Top             =   6570
      Width           =   10875
   End
   Begin VB.Label lblActionStatus 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
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
      Left            =   8340
      TabIndex        =   9
      Top             =   45
      Width           =   2475
   End
   Begin VB.Label Label6 
      BackColor       =   &H00C00000&
      Caption         =   "    "
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
      Height          =   375
      Left            =   0
      TabIndex        =   8
      Top             =   0
      Width           =   10815
   End
   Begin VB.Shape Shape1 
      FillColor       =   &H00BDDFD1&
      FillStyle       =   0  'Solid
      Height          =   675
      Left            =   360
      Shape           =   4  'Rounded Rectangle
      Top             =   6015
      Width           =   5700
   End
End
Attribute VB_Name = "frmMultiAccount"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Const tblName As String = "tmpRptTbl"
Const IdField As String = "USERNAME"
Const OrderField As String = "USERNAME,Nar1"
Const MaxNo As Long = 0
Dim rstbl As Recordset '-- TblAddLess
Dim rstblSub As Recordset '-- TblVouSub
Dim FormAction As Integer
Dim frmTypeStr As String
Dim i As Long
Private Sub cmdBtn_Click(Index As Integer)
On Error GoTo ErrorRoutine
Dim X As String
Dim cdt As Date
Select Case Index
Case 0 'Add
    FormAction = vbDataActionAddNew
    Call GProcActivateControls(Me, True)
    Call GProcClearForm(Me, rstbl, MaxNo, True)
    cmdBtn(1).Enabled = False
    cmdBtn(2).Enabled = False
    cmdBtn(1).BackColor = &H80000004
    cmdBtn(2).BackColor = &H80000004
    dtpFrDt.Value = Format(gCYSDate, "dd/MM/yyyy")
    cdt = Now
    If CDate(cdt) > gCYEDate Then
       dtpToDt.Value = Format(gCYEDate, "dd/MM/yyyy")
    Else
       dtpToDt.Value = Format(Now, "dd/MM/yyyy")
    End If
    
    SetGrid
Case 1, 2
    gCn.BeginTrans
    gCn.Execute "Delete from tmpGenTbl2 where USERNAME='" & gUserName & "'"
    gCn.CommitTrans
    X = "Select tblmastaccount.AcName,  tblvoucher.vdt as vdt,tblvoucher.vtype as Vtype,tblvoucher.vno as vno,Vctrno, tblAccO.acname as  AcOName,iif(vamt<0,abs(vamt),0) AS Debit,  iif(vamt>0,vamt,0)  AS Credit, tblvoucher.vbillno, tblvoucher.vbilldt, vnar1, vnar2, vnar3, vnar4,tblvoucher.vaccode,  AgName,2 as orderno,'" & gUserName & "',tblMastAccount.ACCITY,tblMastAccount.AcPan,tblMastAccount.ACAdd1,tblMastAccount.ACAdd2 " _
        & "From tblvoucher, tblMastAccount , TblMastGroup, tblMastAccount as  tblAccO,tmpRptTbl  Where tblMastAccount.AGCODE=  TblMastGroup.AGCODE and tblvoucher.vacCode = tblMastAccount.AcCode  and tblvoucher.VAcOCode = tblAccO.AcCode  and tblvoucher.Vdt>= cdate('" & Format(dtpFrDt.Value, "yyyy,mm,dd") & "') and tblvoucher.Vdt<=cdate('" & Format(dtpToDt.Value, "yyyy,mm,dd") & "') " _
        & " and tmpRptTbl.VACCODE = tblMastAccount.AcCode and tmpRptTbl.UserName = '" & gUserName & "'  "
    
    gCn.BeginTrans
    gCn.Execute "insert into tmpGenTbl2 (Nar1,Vdt,Vtype,Vno,VsrNo,Nar2,Amt1,Amt2,VBillNo,VBillDt,Nar3,Nar4,Nar5,Nar6,VACCODE,Nar7,Amt3,UserName,Nar8,Nar9,Nar10,Nar11) " & X & ""
    gCn.CommitTrans
             
    X = " Select AcName,null,'0P',0,0,'Opening Balance', iif(sum(vamt)<0,abs(sum(vamt)),0),  iif(sum(vamt)>0,sum(vamt),0),  '',null,'','','','',0,AgName,1,'" & gUserName & "',tblMastAccount.ACCITY,tblMastAccount.AcPan,tblMastAccount.ACAdd1,tblMastAccount.ACAdd2 " _
         & " From tblvoucher, tblmastaccount,TblMastGroup,tmpRptTbl  Where tblMastAccount.AGCODE= TblMastGroup.AGCODE and  tblvoucher.vacCode = tblMastAccount.AcCode And tblvoucher.Vdt < cdate('" & Format(dtpFrDt.Value, "yyyy,mm,dd") & "') " _
         & " and tmpRptTbl.VACCODE = tblMastAccount.AcCode and tmpRptTbl.UserName = '" & gUserName & "'  " _
         & " group by AcName,AgName,'" & gUserName & "',tblMastAccount.ACCITY,tblMastAccount.AcPan,tblMastAccount.ACAdd1,tblMastAccount.ACAdd2"

    gCn.BeginTrans
    gCn.Execute "insert into tmpGenTbl2 (Nar1,Vdt,Vtype,Vno,VsrNo,Nar2,Amt1,Amt2,VBillNo,VBillDt,Nar3,Nar4,Nar5,Nar6,VACCODE,Nar7,Amt3,UserName,Nar8,Nar9,Nar10,Nar11) " & X & ""
    gCn.CommitTrans
    CryReport.WindowTitle = "ACCOUNT LEDGER"
    If Index = 1 Then
        GProcCrystalRptPreparation CrystalReport1, "Window"
    Else
        GProcCrystalRptPreparation CrystalReport1, "Printer"
    End If
     
     With CrystalReport1
          .Formulas(5) = "wRptName = 'ACCOUNT LEDGER'"
          .Formulas(6) = "PANNo= '" & gCPAN & "'"
          
          X = "{tmpGenTbl2.UserName}='" & gUserName & "'"
          .SelectionFormula = X
           If ChkAcConf.Value = 1 Then
              .ReportFileName = gReportPath & "rptAcLedgerConfMulti.rpt"
           Else
              .ReportFileName = gReportPath & "rptAcLedgerMulti.rpt"
           End If
          .Action = 1
          .PageZoom (130)
     End With
Case 4 'Exit
    Unload Me
    Exit Sub
End Select
Exit Sub
ErrorRoutine:
    rstbl.CancelUpdate
    GProcErrorHandler
End Sub
Private Sub SetControlEd()
End Sub
Private Sub DeleteRelatedRecord()
'--- Voucher Sub
gCn.Execute "delete from tmpRptTbl where UserName='" & gUserName & "'"
End Sub
Private Sub SaveRelatedRecord()
Dim i As Long
Dim rsVouSub As Recordset
Dim MillCode As Long
Dim X As String
'-- Delete Records
DeleteRelatedRecord
'--- Save Grid Records
X = "select * from tmpRptTbl where UserName='" & gUserName & "'"
GProcRstOpen rsVouSub, X, "O"
With msGrid
    i = 1
    Do While i < .Rows
        rsVouSub.AddNew
        rsVouSub!Nar1 = .TextMatrix(i, 0) & vbNullString
        rsVouSub!vacCode = GProcGetColumnValue("tblMastAccount", "AcName", .TextMatrix(i, 0), "S", "AcCode", "N")
        rsVouSub!UserName = gUserName
        rsVouSub.Update
        i = i + 1
     Loop
 End With
 rsVouSub.Close
 cmdBtn(1).Enabled = True
 cmdBtn(1).BackColor = &HFFFF80
 cmdBtn(2).Enabled = True
 cmdBtn(2).BackColor = &HFFFF80
 
 cmdBtn(1).SetFocus
 
End Sub
Private Sub dtpToDt_LostFocus()
    msGrid.Col = 0
    msGrid_EnterCell
    txtGrid.SetFocus
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case 13:
            If Me.ActiveControl.Name <> "txtGrid" Then Sendkeys "{TAB}"
            'SendKeys "{TAB}"
'        Case 27:
'            If (cmdBtn(13).Enabled = False) Then
'                Unload Me
'            End If
    End Select
End Sub
Private Sub Form_Load()
    Dim X As String
    Me.Top = 300
    Me.Left = Screen.Width / 2 - Me.Width / 2
    '--- TblAddLess
    X = "Select * from " & tblName & " where username='" & gUserName & "'"
    If Trim(OrderField) <> "" Then X = X & " order by " & OrderField
    GProcRstOpen rstbl, X, "O"
    '---Grid
    
    SetGrid
    SetRelatedRS
    '---- Button Visiblity Set As Per User Previlage
    GProcSetButtonVisiblity Me
    cmdBtn_Click (0)
End Sub
Private Sub SetRelatedRS()
End Sub
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
If FormAction = vbDataActionUpdate Then
    MsgBox "You are about to modify the record. Save the record and then exit.", vbInformation + vbOKOnly, Me.Caption
    Cancel = True
    Exit Sub
End If
rstbl.Requery
rstbl.Close
End Sub
'---------------- Grid
Private Sub SetGrid()
Dim X As String
Dim i As Long
Set rstblSub = New Recordset
X = "select Nar1 from tmpRptTbl where username = '" & gUserName & "' order by Nar1 "
GProcRstOpen rstblSub, X, "R"
With msGrid
    .Clear
    .FormatString = "<Account Name                                                        "
    .ColWidth(1) = 0
    
    If FormAction = vbDataActionAddNew Then
        .Rows = 2
        GridAddNew
    Else
        .Rows = 1
    End If
End With
End Sub
Private Sub GridAddNew()
Dim i As Long
With msGrid
    For i = 0 To rstblSub.Fields.Count - 1
        Select Case rstblSub.Fields(i).Type
            Case adInteger, 2, 3 '-- Integer
                .TextMatrix(.Row, i) = 0
            Case adDecimal, 131, adNumeric, adDouble '-- Double
                .TextMatrix(.Row, i) = Format(0, GProcNumberFormat(rstblSub.Fields(i).Precision, rstblSub.Fields(i).NumericScale))
            Case adDate '-- Date
                 .TextMatrix(.Row, i) = Null
            Case Else
                 .TextMatrix(.Row, i) = ""
        End Select
    Next i
.Refresh
End With
End Sub
Private Sub msGrid_Click()
On Error Resume Next
Dim i As Long
msGrid_EnterCell
End Sub
Private Sub msGrid_GotFocus()
If LCase(Me.ActiveControl.Name) <> "txtgrid" And LCase(Me.ActiveControl.Name) <> "msgrid" Then msGrid.Col = 1
    cmdBtn(4).Cancel = False
End Sub
Private Sub msGrid_EnterCell()
If FormAction <> vbDataActionCancel Then
    DoEvents
    txtGrid = ""
    DoEvents
    If msGrid.text <> "" Then txtGrid.text = msGrid.text
    '--- Lock TextBox
        With msGrid
            Select Case .Col
'                Case 0: '---- Sr No
'                    txtGrid.Visible = True
'                    txtGrid.Locked = True
'                    DoEvents
'                    .Col = .Col + 1
                Case Else
                    txtGrid.Locked = False
            End Select
        End With
    DoEvents
    msGrid_KeyPress (0)
    DoEvents
End If
msGrid_KeyPress (0)
End Sub
Private Sub msGrid_KeyPress(KeyAscii As Integer)
With msGrid
    If KeyAscii = 27 Then '----Esc Key
        KeyAscii = 0
        If .TextMatrix(.Row, 0) = "" And .Rows > 2 Then
            .RemoveItem (.Row)
            txtGrid.text = ""
        End If
        txtGrid.Visible = False
        Exit Sub
    End If
End With
'---- Navigate Row ,Col
If KeyAscii = 13 Then
    KeyAscii = 0
    With msGrid
        If .Col = 0 Then
            If .Row = (.Rows - 1) Then '-- Add New Row
                .Rows = .Rows + 1
                .Row = .Row + 1
                GridAddNew
            Else '--Next Row
                .Row = .Row + 1
            End If
            .Col = 0
        Else    '-- Next col
           .Col = .Col + 1
        End If
    End With
End If
'---- Set TxtGrid
If msGrid.text <> "" Then txtGrid.text = msGrid.text
If msGrid.Col <= rstblSub.Fields.Count - 1 Then
    txtGrid.Move msGrid.CellLeft + msGrid.Left, msGrid.CellTop + msGrid.Top, msGrid.CellWidth, msGrid.CellHeight
End If
txtGrid.Visible = True
txtGrid.Enabled = True
DoEvents
If KeyAscii <> 0 Then
    Select Case rstblSub.Fields(msGrid.Col).Type
        Case adChar, adVarWChar
            If msGrid.Col = 1 Then  '-- Count
                txtGrid_KeyPress (KeyAscii)
            Else
                txtGrid.text = txtGrid.text & Chr(KeyAscii)
            End If
        Case Else
            txtGrid.text = Chr(GProcValidateKey(rstblSub, msGrid.Col, KeyAscii, txtGrid.text))
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
End If
DoEvents
DoEvents
End Sub
Private Sub txtGrid_GotFocus()
    txtGrid.Alignment = ProcSetAlignment(rstblSub, msGrid.Col)
    DoEvents
    Call GProcSelectBox(txtGrid)
End Sub
Private Sub txtGrid_KeyPress(KeyAscii As Integer)
Dim LeftPos As Long
Dim TopPos As Long
Dim MillCode As Long
TopPos = txtGrid.Top + Me.Top + 650
LeftPos = txtGrid.Left + Me.Left
DoEvents
With msGrid
    If KeyAscii = 13 Then
       If .Row = (.Rows - 1) Then
          .Rows = .Rows + 1
          .Row = .Row + 1
          GridAddNew
          cmdBtn(1).Enabled = False
          cmdBtn(2).Enabled = False
          cmdBtn(1).BackColor = &H80000004
          cmdBtn(2).BackColor = &H80000004
       End If
    ElseIf KeyAscii <> 13 And KeyAscii <> 27 Then
      '--- Search
      Select Case .Col
      Case 0  '-- A/c Name
           gClsSearch.SearchMultiField "tblMastaccount", "AcName,AcCity", Array("Account Name", "City"), Array(5000, 2000), "", Chr(KeyAscii), "AcName", LeftPos, TopPos, True
           txtGrid = ""
           KeyAscii = 0
           If gClsSearch.SearchMultiRetCol(0) <> "" Then
              txtGrid = gClsSearch.SearchMultiRetCol(0)
              txtGrid_KeyPress 13
'              If .Row = (.Rows - 1) Then
'                 .Rows = .Rows + 1
'                 .Row = .Row + 1
'                 GridAddNew
'              Else
'                 .Row = .Row + 1
'              End If
           End If
       Case Else
           KeyAscii = GProcValidateKey(rstblSub, msGrid.Col, KeyAscii, txtGrid)
       End Select
    End If
    .Col = 0
    '-- On ESC
    With msGrid
        If KeyAscii = 27 Then '----Esc Key
            KeyAscii = 0
            If (.TextMatrix(.Row, 0) = "") And .Rows > 2 Then
                .RemoveItem (.Row)
                txtGrid.text = ""
            End If
            txtGrid.Visible = False
            SaveRelatedRecord
            
            
           ' txt(18).SetFocus
            Exit Sub
        End If
    End With
End With
End Sub
Private Sub txtGrid_KeyDown(KeyCode As Integer, Shift As Integer)
Dim GName As String
'--- Delete Row On Ctrl+Y
If KeyCode = vbKeyY And Shift = 2 And msGrid.Rows > 2 Then
    msGrid = 0
    msGrid.RemoveItem (msGrid.Row)
    txtGrid.Visible = False
    txtGrid.text = ""
    Exit Sub
End If
'--- Show Form
End Sub
Private Function ProcSetAlignment(pRstbl As Recordset, pColIndex As Long)
Select Case pRstbl.Fields(pColIndex).Type
    Case adInteger, adDouble, 2, 3, 131, adDecimal
        ProcSetAlignment = 1
    Case Else
        ProcSetAlignment = 0
End Select
End Function
Private Sub ReadFields()
FormAction = vbDataActionClose
GProcShowRecord Me, rstbl, MaxNo
End Sub


    
'    If KeyAscii = 13 And .Row = (.Rows - 1) Then
'       .Rows = .Rows + 1
'       .Row = .Row + 1
'       GridAddNew
'    End If



