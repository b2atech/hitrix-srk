VERSION 5.00
Begin VB.Form frmSecurity 
   Caption         =   "User Security"
   ClientHeight    =   2376
   ClientLeft      =   60
   ClientTop       =   348
   ClientWidth     =   6060
   Icon            =   "frmSecurity.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2376
   ScaleWidth      =   6060
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdClose 
      Cancel          =   -1  'True
      Caption         =   "&Close"
      Height          =   400
      Left            =   4320
      TabIndex        =   9
      Top             =   1620
      Width           =   1200
   End
   Begin VB.CommandButton cmdOk 
      Caption         =   "&Ok"
      Default         =   -1  'True
      Height          =   400
      Left            =   4320
      TabIndex        =   8
      Top             =   1140
      Width           =   1200
   End
   Begin VB.TextBox txt 
      Height          =   300
      IMEMode         =   3  'DISABLE
      Index           =   3
      Left            =   2580
      PasswordChar    =   "*"
      TabIndex        =   7
      Text            =   "3"
      Top             =   1680
      Width           =   1515
   End
   Begin VB.TextBox txt 
      Height          =   300
      IMEMode         =   3  'DISABLE
      Index           =   2
      Left            =   2580
      PasswordChar    =   "*"
      TabIndex        =   5
      Text            =   "2"
      Top             =   1260
      Width           =   1515
   End
   Begin VB.TextBox txt 
      Height          =   300
      IMEMode         =   3  'DISABLE
      Index           =   1
      Left            =   2580
      PasswordChar    =   "*"
      TabIndex        =   3
      Text            =   "1"
      Top             =   840
      Width           =   1515
   End
   Begin VB.TextBox txt 
      Enabled         =   0   'False
      Height          =   300
      Index           =   0
      Left            =   2580
      TabIndex        =   1
      Text            =   "0"
      Top             =   420
      Width           =   1515
   End
   Begin VB.Label lblVerifyPassword 
      Caption         =   "Verify New Password"
      Height          =   300
      Index           =   3
      Left            =   600
      TabIndex        =   6
      Top             =   1680
      Width           =   1800
   End
   Begin VB.Label lblEnterNewPassword 
      Caption         =   "Enter New Password"
      Height          =   300
      Index           =   2
      Left            =   600
      TabIndex        =   4
      Top             =   1260
      Width           =   1800
   End
   Begin VB.Label lblEnterPassword 
      Caption         =   "Enter Current Password"
      Height          =   300
      Index           =   1
      Left            =   600
      TabIndex        =   2
      Top             =   840
      Width           =   1800
   End
   Begin VB.Label lblUserName 
      Caption         =   "User Name"
      Height          =   300
      Index           =   0
      Left            =   600
      TabIndex        =   0
      Top             =   420
      Width           =   1800
   End
End
Attribute VB_Name = "frmSecurity"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdClose_Click()
Unload Me
End Sub

Private Sub cmdOK_Click()
If ValidateData = True Then
    SaveProc
    db.Execute "update tbladmuser set other = 0 WHERE USERNAME ='" & UserName & "'"
    End
Else
   ' txt(1).SetFocus
End If
End Sub

Private Sub Form_Load()
'FirstTimeFlag = True
Set rsTbl = New Recordset
rsTbl.Open "select * from tblAdmUser order by UserName", db, adOpenStatic, adLockOptimistic

txt(0) = UserName
For i = 1 To 3
    txt(i) = ""
Next i
End Sub

Private Sub SaveProc()
    db.Execute "update tbladmuser set passward ='" & Trim(txt(2)) & "' where username='" & txt(0) & "' and passward ='" & txt(1) & "'"
End Sub

Private Function ValidateData() As Boolean
Dim rs1 As Recordset
Dim rsp As Recordset
'rsTbl.Find "username = '" & UserName & "'"
ValidateData = True
'If rsTbl.Fields("passward") = txt(1) Then
'Else
'    MsgBox "Current password is not correct.", vbOKOnly, Me.Caption
'    txt(1).SetFocus
'    ValidateData = False
'    Exit Function
'End If
Set rsp = New Recordset
rsp.Open "select * from tbladmuser where passward='" & txt(1) & "' and username='" & UserName & "'", db, adOpenDynamic, adLockOptimistic
rsp.MoveFirst
If rsp.EOF = True And rsp.BOF = True Then
    MsgBox "Current password is not correct.", vbOKOnly, Me.Caption
    txt(1).SetFocus
    ValidateData = False
   Exit Function
End If

If txt(2) = txt(3) Then
Else
    MsgBox "Entered new password and varified password are mismatched.", vbOKOnly, Me.Caption
    txt(3).SetFocus
    ValidateData = False
   Exit Function
End If

Set rs1 = New Recordset
rs1.Open "select * from tbladmuser where passward='" & txt(2) & "'", db, adOpenDynamic, adLockOptimistic
If rs1.EOF = True And rs1.BOF = True Then
Else
    MsgBox "This Password Already Present.", vbOKOnly, Me.Caption
    txt(2).SetFocus
    ValidateData = False
   Exit Function
End If
End Function

Private Sub txt_GotFocus(Index As Integer)
On Error Resume Next
selectbox Me.ActiveControl
End Sub
