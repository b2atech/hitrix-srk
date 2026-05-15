VERSION 5.00
Begin VB.Form frmMastUser 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "User Creation"
   ClientHeight    =   3630
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9660
   Icon            =   "frmMastUser.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3630
   ScaleWidth      =   9660
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
      Left            =   4680
      MaxLength       =   3
      TabIndex        =   22
      Text            =   "2"
      Top             =   2280
      Visible         =   0   'False
      Width           =   435
   End
   Begin VB.CheckBox Chk 
      Caption         =   "Is &Admin User ?"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Index           =   2
      Left            =   1800
      TabIndex        =   6
      Top             =   2280
      Width           =   2175
   End
   Begin VB.TextBox txtPassword 
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
      IMEMode         =   3  'DISABLE
      Left            =   1800
      MaxLength       =   40
      PasswordChar    =   "*"
      TabIndex        =   5
      Text            =   "txtPassword"
      Top             =   1800
      Width           =   4515
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
      Left            =   360
      Style           =   1  'Graphical
      TabIndex        =   12
      ToolTipText     =   "To Add New Record"
      Top             =   240
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
      Left            =   975
      Style           =   1  'Graphical
      TabIndex        =   13
      ToolTipText     =   "To Add New Record"
      Top             =   240
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
      Left            =   1590
      Style           =   1  'Graphical
      TabIndex        =   14
      ToolTipText     =   "To Add New Record"
      Top             =   240
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
      Left            =   2205
      Style           =   1  'Graphical
      TabIndex        =   15
      ToolTipText     =   "To Add New Record"
      Top             =   240
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
      Left            =   2820
      Style           =   1  'Graphical
      TabIndex        =   16
      ToolTipText     =   "To Add New Record"
      Top             =   240
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
      Left            =   360
      Style           =   1  'Graphical
      TabIndex        =   17
      ToolTipText     =   "To Add New Record"
      Top             =   2880
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
      Left            =   1560
      Style           =   1  'Graphical
      TabIndex        =   7
      ToolTipText     =   "To Add New Record"
      Top             =   2880
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
      Left            =   2760
      Style           =   1  'Graphical
      TabIndex        =   8
      ToolTipText     =   "To Add New Record"
      Top             =   2880
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
      Left            =   3960
      Style           =   1  'Graphical
      TabIndex        =   9
      ToolTipText     =   "To Add New Record"
      Top             =   2880
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
      Left            =   4035
      Style           =   1  'Graphical
      TabIndex        =   18
      ToolTipText     =   "To Add New Record"
      Top             =   240
      Width           =   1095
   End
   Begin VB.CommandButton cmdBtn 
      Caption         =   "&Save"
      Default         =   -1  'True
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
      Left            =   5160
      Style           =   1  'Graphical
      TabIndex        =   10
      ToolTipText     =   "To Add New Record"
      Top             =   2880
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
      Left            =   6480
      Style           =   1  'Graphical
      TabIndex        =   11
      ToolTipText     =   "To Add New Record"
      Top             =   1680
      Width           =   1095
   End
   Begin VB.CommandButton cmdBtn 
      Cancel          =   -1  'True
      Caption         =   "&Exit"
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
      Left            =   5250
      Style           =   1  'Graphical
      TabIndex        =   21
      ToolTipText     =   "To Add New Record"
      Top             =   240
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
      Left            =   6480
      Style           =   1  'Graphical
      TabIndex        =   19
      ToolTipText     =   "To Add New Record"
      Top             =   240
      Width           =   1095
   End
   Begin VB.CommandButton cmdBtn 
      Caption         =   "Previlage"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   645
      Index           =   12
      Left            =   6480
      Style           =   1  'Graphical
      TabIndex        =   20
      ToolTipText     =   "To Add New Record"
      Top             =   840
      Width           =   1335
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
      IMEMode         =   3  'DISABLE
      Index           =   1
      Left            =   1800
      MaxLength       =   3
      PasswordChar    =   "*"
      TabIndex        =   3
      Text            =   "1"
      Top             =   1320
      Width           =   4515
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
      Index           =   0
      Left            =   1800
      MaxLength       =   3
      TabIndex        =   1
      Text            =   "0"
      Top             =   840
      Width           =   4515
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Confirm Password"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000001&
      Height          =   285
      Index           =   1
      Left            =   360
      TabIndex        =   4
      Top             =   1800
      Width           =   1815
   End
   Begin VB.Shape Shape1 
      FillColor       =   &H00BDDFD1&
      FillStyle       =   0  'Solid
      Height          =   615
      Left            =   7920
      Shape           =   4  'Rounded Rectangle
      Top             =   1680
      Width           =   1575
   End
   Begin VB.Label Label14 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "User Name"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000001&
      Height          =   285
      Left            =   360
      TabIndex        =   0
      Top             =   840
      Width           =   1140
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Password"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000001&
      Height          =   285
      Index           =   0
      Left            =   360
      TabIndex        =   2
      Top             =   1320
      Width           =   945
   End
End
Attribute VB_Name = "frmMastUser"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Const tblName As String = "tblMastUser"
Const IdField As String = "Username"
Const OrderField As String = "Username"
Dim rstbl As Recordset
Dim varBookMark As Variant
Dim FindFlag As Boolean
Dim FirstTimeFlag As Boolean

Private Sub cmdBtn_Click(Index As Integer)
'On Error GoTo ErrorRoutine
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
Case 5 'List
    'GenMastList ("LAC")
Case 6 'Add
    gFormAction = vbDataActionAddNew
    Call GProcActivateControls(Me, True)
    Call GProcSetButtons(Me, vbDataActionAddNew)
    Call GProcClearForm(Me, rstbl)
    If FirstTimeFlag = True Then
        FirstTimeFlag = False
    Else
        txt(0).SetFocus
    End If
Case 7 'Modify
    'If mURecModi = True Then
    '    SearchHelp
    '    DoEvents
        gFormAction = vbDataActionUpdate
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
        DoEvents
    'End If
Case 9 'Print
Case 10 'Exit
    DoEvents
    Unload Me
Case 11 'Label
    'frmLabel.Show 1
Case 12 'Previlage
    frmAdmUserPrevilage.Show
Case 13 'Save
    If ValidateData = True Then
        'If gFormAction = vbDataActionAddNew Then
        FillTxtFromTemp
        GProcSaveRecord Me, rstbl, gFormAction
        Call GProcActivateControls(Me, False)
        Call GProcSetButtons(Me, True)
        gFormAction = vbDataActionClose
    End If
Case 14 'Cancel
    Call GProcClearForm(Me, rstbl)
    If rstbl.RecordCount <= 0 Then
        Call GProcActivateControls(Me, False)
        Call GProcSetButtons(Me, vbDataActionCancel)
    Else
        If gFormAction = vbDataActionAddNew Then
            rstbl.MoveLast
        Else
            rstbl.CancelUpdate
        End If
        ReadFields
        Call GProcActivateControls(Me, False)
        Call GProcSetButtons(Me, True)
    End If
    gFormAction = vbDataActionClose
End Select
Exit Sub
ErrorRoutine:
    GProcErrorHandler
End Sub
Private Sub Form_Load()
    Dim X As String
    Dim y As String
    Me.Top = 300
    Me.Left = Screen.Width / 2 - Me.Width / 2
    y = "Select * from " & tblName & " order by " & OrderField
    GProcRstOpen rstbl, y, "O", gSetupCn
    FirstTimeFlag = True
    '---- Button Visiblity Set As Per User Previlage
    GProcSetButtonVisiblity Me
    cmdBtn_Click (6)
End Sub
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
If gFormAction = vbDataActionUpdate Then
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
KeyAscii = GProcValidateKey(rstbl, Index, KeyAscii, txt(Index).Text)
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

Private Sub txtpassword_GotFocus()
GProcSelectBox Me.ActiveControl
End Sub
Private Sub txtPassword_LostFocus()
'txtPassword.Text = (StrConv(Trim(txtPassword.Text), vbProperCase))
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
If txt(0) = "" Or txt(1) = "" Or txtPassword <> txt(1) Or txtPassword = txt(0) Then
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
Dim y As String
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
    If gFormAction = vbDataActionAddNew Then
        X = "Select username From " & tblName & " Where Username='" & Trim(txt(0).Text) & "'"
        i = GProcRstOpen(rstAddCheck, X, "R", gSetupCn)
        If i > 0 Then
            MsgBox "Record is Available of User Name." & Trim(txt(0).Text), vbCritical + vbOKOnly, "Duplicate Record"
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
GProcShowRecord Me, rstbl
FillTempFromTxt
End Sub
