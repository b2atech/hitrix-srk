VERSION 5.00
Begin VB.Form frmAdmUser 
   BackColor       =   &H00C0E0FF&
   BorderStyle     =   0  'None
   Caption         =   "User Creation"
   ClientHeight    =   3990
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   9645
   ControlBox      =   0   'False
   Icon            =   "frmAdmUser.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3990
   ScaleWidth      =   9645
   ShowInTaskbar   =   0   'False
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
      Left            =   285
      MaxLength       =   3
      TabIndex        =   22
      Text            =   "2"
      Top             =   3150
      Visible         =   0   'False
      Width           =   435
   End
   Begin VB.CheckBox Chk 
      BackColor       =   &H80000018&
      Caption         =   "Is &Admin User ?"
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
      Height          =   360
      Index           =   2
      Left            =   2430
      TabIndex        =   6
      Top             =   2460
      Width           =   1815
   End
   Begin VB.TextBox txtPassword 
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
      Height          =   360
      IMEMode         =   3  'DISABLE
      Left            =   2400
      MaxLength       =   40
      PasswordChar    =   "*"
      TabIndex        =   5
      Text            =   "txtPassword"
      Top             =   1950
      Width           =   1935
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
      Left            =   6885
      Style           =   1  'Graphical
      TabIndex        =   12
      ToolTipText     =   "Previous"
      Top             =   570
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
      Left            =   7485
      Style           =   1  'Graphical
      TabIndex        =   13
      ToolTipText     =   "Next"
      Top             =   570
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
      Left            =   6885
      Style           =   1  'Graphical
      TabIndex        =   14
      ToolTipText     =   "First"
      Top             =   1020
      Width           =   495
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
      Left            =   7485
      Style           =   1  'Graphical
      TabIndex        =   15
      ToolTipText     =   "Last"
      Top             =   1020
      Width           =   495
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
      Left            =   8130
      Style           =   1  'Graphical
      TabIndex        =   16
      ToolTipText     =   "To Find Record"
      Top             =   570
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
      Left            =   825
      Style           =   1  'Graphical
      TabIndex        =   17
      Top             =   3060
      Visible         =   0   'False
      Width           =   1095
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
      Left            =   6885
      Style           =   1  'Graphical
      TabIndex        =   7
      ToolTipText     =   "To Add New Record"
      Top             =   1545
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
      Left            =   8130
      Style           =   1  'Graphical
      TabIndex        =   8
      ToolTipText     =   "To Modify Record"
      Top             =   2010
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
      Left            =   6885
      Style           =   1  'Graphical
      TabIndex        =   9
      ToolTipText     =   "To Delete Record"
      Top             =   2010
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
      Left            =   1245
      Style           =   1  'Graphical
      TabIndex        =   18
      Top             =   3270
      Visible         =   0   'False
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
      Left            =   6885
      Style           =   1  'Graphical
      TabIndex        =   10
      ToolTipText     =   "To Save Record"
      Top             =   2475
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
      Left            =   8130
      Style           =   1  'Graphical
      TabIndex        =   11
      ToolTipText     =   "To Cancel Record"
      Top             =   2475
      Width           =   1095
   End
   Begin VB.CommandButton cmdBtn 
      Cancel          =   -1  'True
      Caption         =   "E&xit"
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
      Index           =   10
      Left            =   6885
      Style           =   1  'Graphical
      TabIndex        =   21
      ToolTipText     =   "Exit"
      Top             =   2970
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
      Left            =   2460
      Style           =   1  'Graphical
      TabIndex        =   19
      Top             =   3300
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.CommandButton cmdBtn 
      Caption         =   "User &Previlage"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   795
      Index           =   12
      Left            =   8130
      Style           =   1  'Graphical
      TabIndex        =   20
      ToolTipText     =   "User Previlage"
      Top             =   1080
      Width           =   1095
   End
   Begin VB.TextBox txt 
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
      Height          =   360
      IMEMode         =   3  'DISABLE
      Index           =   1
      Left            =   2400
      MaxLength       =   3
      PasswordChar    =   "*"
      TabIndex        =   3
      Text            =   "1"
      Top             =   1350
      Width           =   1950
   End
   Begin VB.TextBox txt 
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
      Left            =   2400
      MaxLength       =   3
      TabIndex        =   1
      Text            =   "0"
      Top             =   735
      Width           =   1950
   End
   Begin VB.Image Image1 
      Height          =   1845
      Left            =   4605
      Picture         =   "frmAdmUser.frx":058A
      Stretch         =   -1  'True
      Top             =   870
      Width           =   1575
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
      Left            =   5475
      TabIndex        =   24
      Top             =   15
      Width           =   2475
   End
   Begin VB.Label Label6 
      BackColor       =   &H00C00000&
      Caption         =   "   User Creation"
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
      Left            =   15
      TabIndex        =   23
      Top             =   0
      Width           =   10455
   End
   Begin VB.Shape Shape4 
      BorderColor     =   &H00C00000&
      BorderWidth     =   5
      FillColor       =   &H00C00000&
      Height          =   3915
      Left            =   0
      Top             =   45
      Width           =   9645
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Confirm Password :"
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
      Index           =   1
      Left            =   360
      TabIndex        =   4
      Top             =   1950
      Width           =   1845
   End
   Begin VB.Shape Shape1 
      FillColor       =   &H00BDDFD1&
      FillStyle       =   0  'Solid
      Height          =   3090
      Left            =   6645
      Shape           =   4  'Rounded Rectangle
      Top             =   465
      Width           =   2775
   End
   Begin VB.Label Label14 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "User Name            :"
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
      Left            =   360
      TabIndex        =   0
      Top             =   735
      Width           =   1815
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Password              :"
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
      Left            =   360
      TabIndex        =   2
      Top             =   1350
      Width           =   1815
   End
   Begin VB.Shape Shape2 
      BackColor       =   &H00808080&
      BorderColor     =   &H004D177D&
      FillColor       =   &H00EDFDFE&
      FillStyle       =   0  'Solid
      Height          =   3180
      Left            =   195
      Top             =   495
      Width           =   6285
   End
End
Attribute VB_Name = "frmAdmUser"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Const tblName As String = "tblMastUser"
Const IdField As String = "Username"
Const OrderField As String = "Username"
Const MaxNo As Long = 2
Dim rstbl As Recordset
Dim varBookMark As Variant
Dim FirstTimeFlag As Boolean
Dim FormAction As Integer

Private Sub cmdBtn_Click(Index As Integer)
On Error GoTo ErrorRoutine
Select Case Index
Case 0 'Previous
    rstbl.MovePrevious
    If rstbl.BOF = True Then
        rstbl.MoveFirst
        MsgBox ("This is First Record"), vbInformation + vbOKOnly, "First Record"
    End If
    ReadFields
Case 1 'Next
    rstbl.MoveNext
    If rstbl.EOF = True Then
        rstbl.MoveLast
        MsgBox ("This is Last Record"), vbInformation + vbOKOnly, "Last Record"
    End If
    ReadFields
Case 2 'First
    rstbl.MoveFirst
    ReadFields
Case 3 'Last
    rstbl.MoveLast
    ReadFields
Case 4 'Find
    gClsSearch.SearchMultiField tblName, "UserName", Array("User Name"), Array(txt(0).Width), "", "", "UserName", txt(0).Left + Me.Left, txt(0).Top + Me.Top + 650, False, gSetupCn
    If gClsSearch.SearchMultiRetCol(0) <> "" Then
        rstbl.MoveFirst
        rstbl.Find "UserName = '" & gClsSearch.SearchMultiRetCol(0) & "'"
        gClsSearch.SearchMultiRetCol(0) = ""
        ReadFields
    End If
Case 5 'List
    'GenMastList ("LAC")
Case 6 'Add
    FormAction = vbDataActionAddNew
    Call GProcActivateControls(Me, True)
    Call GProcSetButtons(Me, vbDataActionAddNew)
    Call GProcClearForm(Me, rstbl, MaxNo, True)
    If FirstTimeFlag = True Then
        FirstTimeFlag = False
    Else
        txt(0).SetFocus
    End If
Case 7 'Modify
    'If mURecModi = True Then
    '    SearchHelp
    '    DoEvents
        FormAction = vbDataActionUpdate
        Call GProcActivateControls(Me, True)
        Call GProcSetButtons(Me, vbDataActionUpdate)
        txt(0).Enabled = False '-- User Name
        txt(1).SetFocus
    'End If
Case 8 'Delete
    'If mURecDel = True Then
        If MsgBox("Do You Really Delete This Record?", vbQuestion + vbYesNo + vbDefaultButton1, "Confirm Delete") = vbYes Then
            If rstbl.RecordCount > 1 Then
                rstbl.Delete
                cmdBtn_Click (1)
            ElseIf rstbl.RecordCount = 1 Then
                rstbl.Delete
                Call GProcActivateControls(Me, False)
                Call GProcSetButtons(Me, vbDataActionCancel)
                Form_Load
            End If
        End If
    'End If
Case 9 'Print
Case 10 'Exit
    Unload Me
Case 11 'Label
    'frmLabel.Show 1
Case 12 'Previlage
    frmAdmUserPrevilage.Show
Case 13 'Save
    If ValidateData = True Then
        'If FormAction = vbDataActionAddNew Then
        FillTxtFromTemp
        GProcSaveRecord Me, rstbl, FormAction, MaxNo
        Call GProcActivateControls(Me, False)
        Call GProcSetButtons(Me, True)
        FormAction = vbDataActionClose
        If FormAction = vbDataActionAddNew Then
            If cmdBtn(6).Visible = True Then cmdBtn(6).SetFocus
        Else
            If cmdBtn(7).Visible = True Then cmdBtn(7).SetFocus
        End If
    End If
Case 14 'Cancel
    Call GProcClearForm(Me, rstbl, MaxNo, True)
    If rstbl.RecordCount <= 0 Then
        Call GProcActivateControls(Me, False)
        Call GProcSetButtons(Me, vbDataActionCancel)
    Else
        If FormAction = vbDataActionAddNew Then
            rstbl.MoveLast
        Else
            rstbl.CancelUpdate
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
    rstbl.CancelUpdate
    GProcErrorHandler
End Sub

Private Sub Form_Load()
    Dim X As String
    Dim Y As String
    Me.Top = 300
    Me.Left = Screen.Width / 2 - Me.Width / 2
    Y = "Select * from " & tblName & " order by " & OrderField
    GProcRstOpen rstbl, Y, "O", gSetupCn
    FirstTimeFlag = True
    '---- Button Visiblity Set As Per User Previlage
    GProcSetButtonVisiblity Me
    cmdBtn_Click (6)
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
If FormAction = vbDataActionUpdate Then
    MsgBox "You are about to modify the record. Save the record and then exit.", vbInformation + vbOKOnly, Me.Caption
    Cancel = True
    Exit Sub
End If
'mShowRec = Space(5)
rstbl.Close
End Sub

Private Sub txt_GotFocus(Index As Integer)
GProcSelectBox Me.ActiveControl
End Sub
Private Sub txt_LostFocus(Index As Integer)
'GProcCheckForNumber Me, rstbl, Index
SaveBtnEd
End Sub
Private Sub txt_KeyPress(Index As Integer, KeyAscii As Integer)
If KeyAscii = 13 Then Exit Sub
KeyAscii = GProcValidateKey(rstbl, Index, KeyAscii, txt(Index).text)
End Sub

Private Sub txt_Validate(Index As Integer, Cancel As Boolean)
If Index = 0 Then '-- User Name
        If Trim(txt(Index)) = "" Then
            MsgBox "Enter User Name.", vbCritical + vbOKOnly, Me.Caption
            Cancel = True
            Exit Sub
        End If
ElseIf Index = 1 Then '-- Password
        If Trim(txt(Index)) = "" Then
            MsgBox "Enter Password.", vbCritical + vbOKOnly, Me.Caption
            Cancel = True
            Exit Sub
        End If
End If
End Sub

Private Sub txtPassword_GotFocus()
GProcSelectBox Me.ActiveControl
End Sub
Private Sub txtPassword_LostFocus()
SaveBtnEd
End Sub

Private Sub txtPassword_Validate(Cancel As Boolean)
'--- Password Confirmation
If txtPassword <> txt(1) Then
    MsgBox "Password and Confirm Password must be equal.", vbCritical + vbOKOnly, Me.Caption
    Cancel = True
End If
End Sub

Private Sub SaveBtnEd()
'--- User Name,Password
If (txt(0) = "" Or txt(1) = "" Or txtPassword <> txt(1)) Then
    cmdBtn(13).Enabled = False
    cmdBtn(13).BackColor = &H80000004
Else
    cmdBtn(13).Enabled = True
    cmdBtn(13).BackColor = &HFFFF80
End If
End Sub
Private Function ValidateData() As Boolean
Dim i As Long
Dim X As String
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
        '--- Password
        If Trim(txt(1)) = "" Then
            MsgBox "Enter Password.", vbCritical + vbOKOnly, Me.Caption
            txt(1).SetFocus
            Exit Function
        End If
        '--- Confirm Password
        If txtPassword <> txt(1) Then
            MsgBox "Password and Confirm Password must be equal.", vbCritical + vbOKOnly, Me.Caption
            txtPassword.SetFocus
            Exit Function
        End If
        '--- Password & User
        If txtPassword = txt(0) Then
            MsgBox "Password should not be same like User Name.", vbCritical + vbOKOnly, Me.Caption
            txtPassword.SetFocus
            Exit Function
        End If
    '----Duplicate Records
    '----Addition Case
    If FormAction = vbDataActionAddNew Then
        X = "Select username From " & tblName & " Where Username='" & Trim(txt(0).text) & "'"
        i = GProcRstOpen(rstAddCheck, X, "R", gSetupCn)
        If i > 0 Then
            MsgBox "Record is Available of User Name." & Trim(txt(0).text), vbCritical + vbOKOnly, "Duplicate Record"
            txt(0).SetFocus
            Exit Function
        End If
        rstAddCheck.Close
    End If
    ValidateData = True
End Function
Private Sub FillTxtFromTemp()
txt(2) = Chk(2)
End Sub
Private Sub FillTempFromTxt()
Chk(2) = txt(2)
txtPassword = txt(1)
End Sub
Private Sub ReadFields()
GProcShowRecord Me, rstbl, MaxNo
FillTempFromTxt
End Sub
