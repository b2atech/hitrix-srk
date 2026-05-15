VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmAdmUserPrevilage 
   BackColor       =   &H00C2E3FC&
   BorderStyle     =   0  'None
   Caption         =   "User Previlage"
   ClientHeight    =   6915
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   10095
   ControlBox      =   0   'False
   Icon            =   "frmAdmUserPrevilage.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6915
   ScaleWidth      =   10095
   ShowInTaskbar   =   0   'False
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
      Left            =   8160
      Style           =   1  'Graphical
      TabIndex        =   24
      Top             =   7080
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
      Left            =   10080
      Style           =   1  'Graphical
      TabIndex        =   23
      ToolTipText     =   "To Add New Record"
      Top             =   7080
      Visible         =   0   'False
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
      Left            =   8520
      Style           =   1  'Graphical
      TabIndex        =   22
      ToolTipText     =   "Exit"
      Top             =   5100
      Width           =   1095
   End
   Begin VB.CommandButton cmdBtn 
      Caption         =   "&Cancel"
      CausesValidation=   0   'False
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
      Index           =   14
      Left            =   8520
      Style           =   1  'Graphical
      TabIndex        =   21
      ToolTipText     =   "To Cancel Record"
      Top             =   4470
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
      Left            =   8520
      Style           =   1  'Graphical
      TabIndex        =   20
      ToolTipText     =   "To Save Record"
      Top             =   3990
      Width           =   1095
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
      Index           =   9
      Left            =   9120
      Style           =   1  'Graphical
      TabIndex        =   19
      Top             =   7080
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.CommandButton cmdBtn 
      Caption         =   "&Delete"
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
      Index           =   8
      Left            =   8520
      Style           =   1  'Graphical
      TabIndex        =   18
      ToolTipText     =   "To Delete Record"
      Top             =   3330
      Width           =   1095
   End
   Begin VB.CommandButton cmdBtn 
      Caption         =   "&Modify"
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
      Index           =   7
      Left            =   8520
      Style           =   1  'Graphical
      TabIndex        =   17
      ToolTipText     =   "To Modify Record"
      Top             =   2850
      Width           =   1080
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
      Index           =   6
      Left            =   8520
      Style           =   1  'Graphical
      TabIndex        =   16
      ToolTipText     =   "To Add New Record"
      Top             =   2370
      Width           =   1095
   End
   Begin VB.CommandButton cmdBtn 
      Caption         =   "&List"
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
      Index           =   5
      Left            =   8115
      Style           =   1  'Graphical
      TabIndex        =   15
      Top             =   6975
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.CommandButton cmdBtn 
      Caption         =   "&Find"
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
      Left            =   8520
      Style           =   1  'Graphical
      TabIndex        =   14
      ToolTipText     =   "To Find Record"
      Top             =   1680
      Width           =   1095
   End
   Begin VB.CommandButton cmdBtn 
      Caption         =   ">>"
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
      Index           =   3
      Left            =   9120
      Style           =   1  'Graphical
      TabIndex        =   13
      ToolTipText     =   "Last"
      Top             =   1200
      Width           =   495
   End
   Begin VB.CommandButton cmdBtn 
      Caption         =   "<<"
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
      Left            =   8520
      Style           =   1  'Graphical
      TabIndex        =   12
      ToolTipText     =   "First"
      Top             =   1200
      Width           =   495
   End
   Begin VB.CommandButton cmdBtn 
      Caption         =   ">"
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
      Left            =   9120
      Style           =   1  'Graphical
      TabIndex        =   11
      ToolTipText     =   "Next"
      Top             =   720
      Width           =   495
   End
   Begin VB.CommandButton cmdBtn 
      Caption         =   "<"
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
      Left            =   8520
      Style           =   1  'Graphical
      TabIndex        =   10
      ToolTipText     =   "Previous"
      Top             =   720
      Width           =   495
   End
   Begin VB.TextBox txt 
      Appearance      =   0  'Flat
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
      Index           =   7
      Left            =   11160
      MaxLength       =   3
      TabIndex        =   9
      Text            =   "7"
      Top             =   360
      Visible         =   0   'False
      Width           =   315
   End
   Begin VB.TextBox txt 
      Appearance      =   0  'Flat
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
      Index           =   6
      Left            =   10440
      MaxLength       =   3
      TabIndex        =   8
      Text            =   "6"
      Top             =   360
      Visible         =   0   'False
      Width           =   315
   End
   Begin VB.TextBox txt 
      Appearance      =   0  'Flat
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
      Index           =   5
      Left            =   9960
      MaxLength       =   3
      TabIndex        =   7
      Text            =   "5"
      Top             =   360
      Visible         =   0   'False
      Width           =   315
   End
   Begin VB.TextBox txt 
      Appearance      =   0  'Flat
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
      Index           =   4
      Left            =   11040
      MaxLength       =   3
      TabIndex        =   6
      Text            =   "4"
      Top             =   0
      Visible         =   0   'False
      Width           =   315
   End
   Begin VB.TextBox txt 
      Appearance      =   0  'Flat
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
      Index           =   3
      Left            =   10560
      MaxLength       =   3
      TabIndex        =   5
      Text            =   "3"
      Top             =   0
      Visible         =   0   'False
      Width           =   315
   End
   Begin VB.TextBox txt 
      Appearance      =   0  'Flat
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
      Index           =   2
      Left            =   10200
      MaxLength       =   3
      TabIndex        =   4
      Text            =   "2"
      Top             =   0
      Visible         =   0   'False
      Width           =   315
   End
   Begin VB.TextBox txt 
      Appearance      =   0  'Flat
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
      Index           =   1
      Left            =   9720
      MaxLength       =   3
      TabIndex        =   3
      Text            =   "1"
      Top             =   0
      Visible         =   0   'False
      Width           =   315
   End
   Begin VB.TextBox txt 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Index           =   0
      Left            =   1515
      MaxLength       =   3
      TabIndex        =   1
      Text            =   "0"
      Top             =   705
      Width           =   6540
   End
   Begin MSComctlLib.TreeView TreeView1 
      Height          =   5220
      Left            =   225
      TabIndex        =   2
      Top             =   1320
      Width           =   7860
      _ExtentX        =   13864
      _ExtentY        =   9208
      _Version        =   393217
      LineStyle       =   1
      Style           =   6
      Checkboxes      =   -1  'True
      Appearance      =   1
   End
   Begin VB.Line Line3 
      BorderColor     =   &H000000C0&
      X1              =   150
      X2              =   8175
      Y1              =   1170
      Y2              =   1170
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
      Left            =   7260
      TabIndex        =   26
      Top             =   30
      Width           =   2475
   End
   Begin VB.Label Label6 
      BackColor       =   &H00C00000&
      Caption         =   "   User Previlage"
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
      TabIndex        =   25
      Top             =   0
      Width           =   10455
   End
   Begin VB.Shape Shape4 
      BorderColor     =   &H00C00000&
      BorderWidth     =   5
      FillColor       =   &H00C00000&
      Height          =   6555
      Left            =   0
      Top             =   345
      Width           =   10095
   End
   Begin VB.Shape Shape1 
      FillColor       =   &H00BDDFD1&
      FillStyle       =   0  'Solid
      Height          =   5280
      Left            =   8355
      Shape           =   4  'Rounded Rectangle
      Top             =   600
      Width           =   1470
   End
   Begin VB.Label Label14 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "User Name  :"
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
      Left            =   225
      TabIndex        =   0
      Top             =   705
      Width           =   1215
   End
   Begin VB.Shape Shape2 
      BackColor       =   &H00808080&
      BorderColor     =   &H004D177D&
      FillColor       =   &H00EDFDFE&
      FillStyle       =   0  'Solid
      Height          =   6225
      Left            =   120
      Top             =   525
      Width           =   8055
   End
End
Attribute VB_Name = "frmAdmUserPrevilage"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Const tblName As String = "tblUserPrevilage"
Const OrderField As String = "Username,srno"
Const CriteriaStr As String = ""
Const MaxNo As Long = 7
Dim Rstbl As Recordset
Dim varBookMark As Variant
Dim FirstTimeFlag As Boolean
Dim SrNo As Long
Dim FormAction As Integer

Private Sub cmdBtn_Click(Index As Integer)
On Error GoTo ErrorRoutine
Dim NodeCn As Long
Select Case Index
Case 0 'Previous
    Do While Not Rstbl.BOF
       If Rstbl.Fields(0) <> txt(0) Then Exit Do
        Rstbl.MovePrevious
    Loop
    If Not Rstbl.BOF Then
       txt(0) = Rstbl.Fields(0)
    Else
        Rstbl.MoveFirst
        MsgBox ("This is First Record"), vbInformation + vbOKOnly, "First Record"
    End If
    ReadFields
Case 1 'Next
    Do While Not Rstbl.EOF
       If Rstbl.Fields(0) <> txt(0) Then Exit Do
       Rstbl.MoveNext
    Loop
    If Not Rstbl.EOF Then
       txt(0) = Rstbl.Fields(0)
    Else
        Rstbl.MoveLast
        MsgBox ("This is Last Record"), vbInformation + vbOKOnly, "Last Record"
    End If
    ReadFields
Case 2 'First
    Rstbl.MoveFirst
    txt(0) = Rstbl.Fields(0)
    ReadFields
Case 3 'Last
    Rstbl.MoveLast
    txt(0) = Rstbl.Fields(0)
    ReadFields
Case 4 'Find
    gClsSearch.SearchMultiField tblName, "UserName", Array("User Name"), Array(txt(0).Width), " srno=1", "", "UserName", txt(0).Left + Me.Left, txt(0).Top + Me.Top + 650, False, gSetupCn
    If gClsSearch.SearchMultiRetCol(0) <> "" Then
        Rstbl.MoveFirst
        Rstbl.Find "UserName = '" & gClsSearch.SearchMultiRetCol(0) & "'"
        gClsSearch.SearchMultiRetCol(0) = ""
        ReadFields
    End If
Case 5 'List
    'GenMastList ("LAC")
Case 6 'Add
    FormAction = vbDataActionAddNew
    Call GProcActivateControls(Me, True)
    Call GProcSetButtons(Me, vbDataActionAddNew)
    Call GProcClearForm(Me, Rstbl, MaxNo, True)
    FillTreeViewFromMenu
    If FirstTimeFlag = True Then
        FirstTimeFlag = False
    Else
        txt(0).SetFocus
    End If
Case 7 'Modify
    'If mURecModi = True Then
        FormAction = vbDataActionUpdate
        Call GProcActivateControls(Me, True)
        Call GProcSetButtons(Me, vbDataActionUpdate)
        txt(0).Enabled = False '-- User Name
        TreeView1.SetFocus
    'End If
Case 8 'Delete
    'If mURecDel = True Then
        If MsgBox("Do You Really Delete This Record?", vbQuestion + vbYesNo + vbDefaultButton1, "Confirm Delete") = vbYes Then
            If Rstbl.RecordCount > 1 Then
                '-- Multi Records
                Do While Not Rstbl.EOF
                    If Rstbl.Fields(0) <> txt(0) Then
                        txt(0) = Rstbl.Fields(0)
                        Exit Do
                    End If
                    Rstbl.Delete
                    Rstbl.MoveNext
                Loop
                cmdBtn_Click (1)
            ElseIf Rstbl.RecordCount = 1 Then
                '-- Multi Records
                Do While Not Rstbl.EOF
                    If Rstbl.Fields(0) <> txt(0) Then
                        txt(0) = Rstbl.Fields(0)
                        Exit Do
                    End If
                    Rstbl.Delete
                    Rstbl.MoveNext
                Loop
                Call GProcActivateControls(Me, False)
                Call GProcSetButtons(Me, vbDataActionCancel)
                Form_Load
            End If
        End If
    'End If
Case 9 'Print
Case 10 'Exit
    DoEvents
    Unload Me
Case 11 'Label
    'frmLabel.Show 1
Case 12 'Optional
    
Case 13 'Save
    If ValidateData = True Then
        '------------------ Save Multi Record -------------
        If FormAction = vbDataActionUpdate Then
            Rstbl.MoveFirst
            Rstbl.Find "username='" & txt(0) & "'"
            Do While Not Rstbl.EOF
                If Rstbl.Fields(0) <> txt(0) Then
                    'txt(0) = rstbl.Fields(0)
                    Exit Do
                End If
                Rstbl.Delete
                Rstbl.MoveNext
            Loop
        End If
        FormAction = vbDataActionAddNew
        SrNo = 1
        For NodeCn = 1 To TreeView1.Nodes.Count
            If TreeView1.Nodes(NodeCn).Checked = True Then
                FillTxtFromTemp NodeCn
                'If LCase(Right(txt(1), 3)) <> "frm" Or (LCase(Right(txt(1), 3)) = "frm" And (txt(3) = 1 Or txt(4) = 1 Or txt(5) = 1 Or txt(6) = 1)) Then
                    SrNo = SrNo + 1
                    GProcSaveRecord Me, Rstbl, FormAction, MaxNo
                'End If
            End If
        Next NodeCn
        '-------------------------------------------------
        Call GProcActivateControls(Me, False)
        Call GProcSetButtons(Me, True)
        If FormAction = vbDataActionAddNew Then
            If cmdBtn(6).Visible = True Then cmdBtn(6).SetFocus
        Else
            If cmdBtn(7).Visible = True Then cmdBtn(7).SetFocus
        End If
        FormAction = vbDataActionClose
    End If
Case 14 'Cancel
    Call GProcClearForm(Me, Rstbl, MaxNo, True)
    If Rstbl.RecordCount <= 0 Then
        Call GProcActivateControls(Me, False)
        Call GProcSetButtons(Me, vbDataActionCancel)
    Else
        If FormAction = vbDataActionAddNew Then
            Rstbl.MoveLast
        Else
            'Rstbl.CancelUpdate
        End If
        ReadFields
        Call GProcActivateControls(Me, False)
        Call GProcSetButtons(Me, True)
    End If
    If FormAction = vbDataActionAddNew Then
        If cmdBtn(6).Visible = True Then cmdBtn(6).SetFocus
    Else
        If cmdBtn(7).Visible = True Then cmdBtn(7).SetFocus
    End If
    FormAction = vbDataActionClose
End Select
Exit Sub
ErrorRoutine:
    Rstbl.CancelUpdate
    GProcErrorHandler
End Sub

Private Sub Form_Load()
    Dim x As String
    Me.Top = 300
    Me.Left = Screen.Width / 2 - Me.Width / 2
    x = "Select * from " & tblName
    If Trim(CriteriaStr) <> "" Then x = x & " where " & CriteriaStr
    If Trim(OrderField) <> "" Then x = x & " order by " & OrderField
    GProcRstOpen Rstbl, x, "O", gSetupCn
    FirstTimeFlag = True
    cmdBtn_Click (6)
End Sub
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
If FormAction = vbDataActionUpdate Then
    MsgBox "You are about to modify the record. Save the record and then exit.", vbInformation + vbOKOnly, Me.Caption
    Cancel = True
    Exit Sub
End If
Rstbl.Close
End Sub
Private Sub txt_GotFocus(Index As Integer)
GProcSelectBox Me.ActiveControl
End Sub

Private Sub txt_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
'If Index = 0 Then GProcShowForm frmAdmUser, KeyCode
End Sub

Private Sub txt_LostFocus(Index As Integer)
'-- User Name
If Index <> 0 Then GProcCheckForNumber Me, Rstbl, MaxNo
SaveBtnEd
End Sub
Private Sub txt_KeyPress(Index As Integer, KeyAscii As Integer)
Dim LeftPos As Long
Dim TopPos As Long
If KeyAscii = 13 Then Exit Sub
TopPos = txt(Index).Top + Me.Top + 650
LeftPos = txt(Index).Left + Me.Left
If Index = 0 Then '-- User Name
    gClsSearch.SearchMultiField "tblMastUser", "UserName", Array("User Name"), Array(txt(Index).Width), " isAdminUser=0 ", Chr(KeyAscii), "UserName", LeftPos, TopPos, True, gSetupCn
    'gClsSearch.SearchMultiSelField "select vno,vdt,AdBillAmt,Vtype from tblAddLess", "", Array("Vno", "Vdt", "Amt", "VType"), Array(3000, 2000, 1000, 1000), " ", Chr(KeyAscii), "", LeftPos, TopPos, , True
    KeyAscii = 0
    If gClsSearch.SearchMultiRetCol(0) <> "" Then txt(Index) = gClsSearch.SearchMultiRetCol(0)
    'If gClsSearch.SearchMultiSelRetCol(0, 0) <> "" Then txt(Index) = gClsSearch.SearchMultiSelRetCol(0, 0)
Else
    KeyAscii = GProcValidateKey(Rstbl, Index, KeyAscii, txt(Index).Text)
End If
End Sub

Private Sub txt_Validate(Index As Integer, Cancel As Boolean)
'If Index = 0 Then '-- User Name
'        If Trim(txt(Index)) = "" Then
'            MsgBox "Enter User Name.", vbCritical + vbOKOnly, Me.Caption
'            Cancel = True
'            Exit Sub
'        End If
'End If
End Sub

Private Sub SaveBtnEd()
'--- User Name
If txt(0) = "" Then
    cmdBtn(13).Enabled = False
    cmdBtn(13).BackColor = &H80000004
Else
    cmdBtn(13).Enabled = True
    cmdBtn(13).BackColor = &HFFFF80
End If
End Sub
Private Function ValidateData() As Boolean
Dim i As Long
Dim x As String
Dim Y As String
Dim rstAddCheck As Recordset
Dim rstModiCheck As Recordset
    '---- Required Fields
        '--- User Name
        If Trim(txt(0)) = "" Then
            MsgBox "Enter User Name.", vbCritical + vbOKOnly, Me.Caption
            txt(0).SetFocus
            Exit Function
        End If
'        '--- Create User Name
'        If GProcGetColumnValue("tblMastUser", "UserName", txt(0), "S", "UserName", "S", "", gSetupCn) = "" Then
'            GProcShowForm frmAdmUser, vbKeyF2, "", Array(frmAdmUser.txt(0)), Array(txt(0))
'            Exit Function
'        End If
    '----Duplicate Records
    '----Addition Case
    If FormAction = vbDataActionAddNew Then
        x = "Select username From " & tblName & " Where Username='" & Trim(txt(0).Text) & "'"
        i = GProcRstOpen(rstAddCheck, x, "R", gSetupCn)
        If i > 0 Then
            MsgBox "Record is Available of User Name." & Trim(txt(0).Text), vbCritical + vbOKOnly, "Duplicate Record"
            txt(0).SetFocus
            Exit Function
        End If
        rstAddCheck.Close
    End If
    ValidateData = True
End Function
Private Sub FillTxtFromTemp(Optional NIndex As Long)
'-------------- Fill Txt From Tree view
Dim MenuName As String
Dim TempStr As String
Dim MenuIndex As Long
Dim AddStatus As Long
Dim ModStatus As Long
Dim DeleteStatus As Long
Dim ViewStatus As Long
With TreeView1.Nodes(NIndex)
        MenuName = Left(.Key, InStr(1, .Key, "-") - 1)
        TempStr = Right(.Key, Len(.Key) - InStr(1, .Key, "-"))
        AddStatus = 0
        ModStatus = 0
        DeleteStatus = 0
        ViewStatus = 0
        Select Case Left(TempStr, 1)
            Case "A"
                AddStatus = 1
            Case "M"
                ModStatus = 1
            Case "D"
                DeleteStatus = 1
            Case "V"
                ViewStatus = 1
        End Select
        If IsNumeric(TempStr) Then
            MenuIndex = CLng(TempStr)
        Else
            MenuIndex = CLng(Mid(TempStr, 2))
        End If
        txt(1) = MenuName
        txt(2) = MenuIndex
        txt(3) = AddStatus
        txt(4) = ModStatus
        txt(5) = DeleteStatus
        txt(6) = ViewStatus
        txt(7) = SrNo
End With
'---------------- end
End Sub
Private Sub FillTempFromTxt()
FillTreeViewFromMenu
End Sub
Private Sub ReadFields()
Dim MNname As String
Dim MNIndex As Long
Dim MNKey As String
'-- Fill Txt Contrl
GProcShowRecord Me, Rstbl, MaxNo
'-- Fill Temp Control
FillTempFromTxt
'--- Fill Tree View
Rstbl.MoveFirst
Rstbl.Find "UserName = '" & txt(0) & "'"
If Not Rstbl.EOF Then
   varBookMark = Rstbl.Bookmark
Do While Not Rstbl.EOF
    If (txt(0) <> Rstbl.Fields(0)) Then Exit Do
    MNname = Rstbl.Fields("menuName")
    MNIndex = Rstbl.Fields("Menuindex")
    MNKey = MNname & "-" & MNIndex
    Dim n As Node
    
    If IsNodeExist(MNKey) = True Then
        TreeView1.Nodes(MNKey).Checked = True
        '--- Add,Modify,Delete,View Nodes
        If Rstbl.Fields("Add") = 1 Then
            MNKey = MNname & "-A" & MNIndex
            TreeView1.Nodes(MNKey).Checked = True
        End If
        If Rstbl.Fields("Modify") = 1 Then
            MNKey = MNname & "-M" & MNIndex
            TreeView1.Nodes(MNKey).Checked = True
        End If
        If Rstbl.Fields("delete") = 1 Then
            MNKey = MNname & "-D" & MNIndex
            TreeView1.Nodes(MNKey).Checked = True
        End If
        If Rstbl.Fields("View") = 1 Then
            MNKey = MNname & "-V" & MNIndex
            TreeView1.Nodes(MNKey).Checked = True
        End If
    End If
    Rstbl.MoveNext
Loop
    Rstbl.Bookmark = varBookMark
End If
End Sub
Private Function IsNodeExist(NodeKey As String) As Boolean
Dim i As Long
For i = 1 To TreeView1.Nodes.Count
    If TreeView1.Nodes(i).Key = NodeKey Then
        IsNodeExist = True
        Exit For
    End If
Next
End Function
'----- Fill Menu In Treeview
Private Sub FillTreeViewFromMenu()
'On Error GoTo ErrorHandler
Dim i As Long
Dim MnuCtrl As Control
Dim RootStr(0 To 30) As String, TempStr As String
Dim RootCount As Long
Dim MnuIndex As String
Dim MnuCaption As String
Dim j As Long

RootCount = 0
TreeView1.Nodes.Clear
For Each MnuCtrl In frmMain
    If TypeOf MnuCtrl Is Menu Then
    Select Case LCase(Left(MnuCtrl.Name, 6))
'============= Previlage for menu "mnutl", "mnutrn", "mnurpt", "mnumst" Only
    Case "mnuutl", "mnutrn", "mnurpt", "mnumst", "mnugst"
        If Trim(MnuCtrl.Caption) <> "-" Then
            '--- Child
            For i = RootCount To 1 Step -1
               TempStr = Left(RootStr(i - 1), InStr(1, RootStr(i - 1), "-") - 1)
               MnuIndex = MnuCtrl.Index
                If InStr(1, LCase(MnuCtrl.Name), LCase(TempStr)) > 0 And (LCase(MnuCtrl.Name) <> LCase(TempStr)) Then
                    '--------------------- Remove '&' Character
                    MnuCaption = MnuCtrl.Caption
                    MnuCaption = Replace(MnuCaption, "&&", " and ")
                    MnuCaption = Replace(MnuCaption, "&", "")
                    MnuCaption = Replace(MnuCaption, " and ", "&")
                    '----------------------
                    TreeView1.Nodes.Add RootStr(i - 1), tvwChild, MnuCtrl.Name & "-" & MnuIndex, MnuCaption
                    'TreeView1.Nodes(MnuCtrl.Name & "-" & MnuIndex).Expanded = False
                    TreeView1.Nodes(MnuCtrl.Name & "-" & MnuIndex).EnsureVisible
                    RootStr(i) = MnuCtrl.Name & "-" & MnuIndex
                    RootCount = i + 1 ' - 1
                    '---- Add Transaction Nodes [Add,Modify,Delete,View]
                    If LCase(Right(Trim(MnuCtrl.Name), 3)) = "frm" Then AddTransactionNodes MnuCtrl, RootStr(i), MnuIndex
                    Exit For
                End If
            Next i
            '--- Root
            If i = 0 Then
                '--------------------- Remove '&' Character
                MnuCaption = MnuCtrl.Caption
                MnuCaption = Replace(MnuCaption, "&&", " and ")
                MnuCaption = Replace(MnuCaption, "&", "")
                MnuCaption = Replace(MnuCaption, " and ", "&")
                '----------------------
                MnuIndex = MnuCtrl.Index
                TreeView1.Nodes.Add , , MnuCtrl.Name & "-" & MnuIndex, MnuCaption
                TreeView1.Nodes(MnuCtrl.Name & "-" & MnuIndex).EnsureVisible
                RootStr(i) = MnuCtrl.Name & "-" & MnuIndex
                '---- Add Transaction Nodes [Add,Modify,Delete,View]
                If LCase(Right(Trim(MnuCtrl.Name), 3)) = "frm" Then AddTransactionNodes MnuCtrl, RootStr(i), MnuIndex
                RootCount = 1
            Else
            End If
         End If
    End Select
    End If
Next
For i = 1 To TreeView1.Nodes.Count
    ' Expand all nodes.
      TreeView1.Nodes(i).Expanded = False
Next i
TreeView1.Nodes(1).Selected = True
Exit Sub
ErrorHandler:
    MsgBox Err.Description, vbOKOnly, Me.Caption
End Sub
'---- Add Transaction Nodes [Add,Modify,Delete,View]
Private Sub AddTransactionNodes(mnu As Menu, RtStr As String, NodeCnt As String)
    Dim keyStr As String
    '-- Add
    keyStr = mnu.Name & "-A" & NodeCnt
    TreeView1.Nodes.Add RtStr, tvwChild, keyStr, "Add"
    TreeView1.Nodes(keyStr).EnsureVisible
    '-- Modify
    keyStr = mnu.Name & "-M" & NodeCnt
    TreeView1.Nodes.Add RtStr, tvwChild, keyStr, "Modify"
    TreeView1.Nodes(keyStr).EnsureVisible
    '-- Delete
    keyStr = mnu.Name & "-D" & NodeCnt
    TreeView1.Nodes.Add RtStr, tvwChild, keyStr, "Delete"
    TreeView1.Nodes(keyStr).EnsureVisible
    '-- View
    keyStr = mnu.Name & "-V" & NodeCnt
    TreeView1.Nodes.Add RtStr, tvwChild, keyStr, "View"
    TreeView1.Nodes(keyStr).EnsureVisible
End Sub

Private Sub TreeView1_NodeCheck(ByVal Node As MSComctlLib.Node)
Dim i As Long
Dim IsSelected As Boolean
Dim ParentKey As String
Dim ParentNode As Node
Dim ChkNode As Node
'================== if node checked Set All Child Nodes Checked
ProcSetChildNode Node
'================== if One of Child Nodes Check then set All Parent Nodes checked
 Set ChkNode = Node
 Set ParentNode = ChkNode.Parent
 While Not ParentNode Is Nothing
     ParentKey = ChkNode.Parent.Key
     IsSelected = ChkNode.Checked
     i = ChkNode.FirstSibling.Index
     Do While i <= TreeView1.Nodes.Count
         If Not TreeView1.Nodes(i).Parent Is Nothing Then
             If TreeView1.Nodes(i).Parent.Key = ParentKey Then IsSelected = IsSelected Or TreeView1.Nodes(i).Checked
         End If
         i = i + 1
     Loop
     ChkNode.Parent.Checked = IsSelected
     Set ChkNode = ChkNode.Parent
     Set ParentNode = ChkNode.Parent
Wend
End Sub
'---- Set All Child Nodes Checked
Private Sub ProcSetChildNode(pNode As Node)
Dim ChildNode As Node
Dim IsChecked As Boolean
Dim ChildCount As Long
Dim i As Long
Dim ChkNode As Node
Set ChkNode = pNode

IsChecked = ChkNode.Checked
If ChkNode.Children > 0 Then
    i = ChkNode.Child.FirstSibling.Index
    Set ChildNode = TreeView1.Nodes(i)
    For ChildCount = 1 To ChkNode.Children
        ChildNode.Checked = IsChecked
        ProcSetChildNode ChildNode
        Set ChildNode = ChildNode.Next
        If ChildNode Is Nothing Then Exit For
     Next
End If
End Sub


