VERSION 5.00
Begin VB.Form frmUtlChangePassword 
   BackColor       =   &H00C0E0FF&
   BorderStyle     =   0  'None
   Caption         =   "User Security"
   ClientHeight    =   3225
   ClientLeft      =   0
   ClientTop       =   -45
   ClientWidth     =   8190
   FillColor       =   &H00C0E0FF&
   Icon            =   "frmUtlCheckPassword.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3225
   ScaleWidth      =   8190
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdClose 
      Cancel          =   -1  'True
      Caption         =   "&Close"
      Height          =   400
      Left            =   6645
      TabIndex        =   9
      Top             =   1815
      Width           =   1200
   End
   Begin VB.CommandButton cmdOk 
      Caption         =   "&Ok"
      Default         =   -1  'True
      Height          =   400
      Left            =   6645
      TabIndex        =   8
      Top             =   1335
      Width           =   1200
   End
   Begin VB.TextBox txt 
      Height          =   300
      IMEMode         =   3  'DISABLE
      Index           =   3
      Left            =   2985
      PasswordChar    =   "*"
      TabIndex        =   7
      Text            =   "3"
      Top             =   2250
      Width           =   1425
   End
   Begin VB.TextBox txt 
      Height          =   300
      IMEMode         =   3  'DISABLE
      Index           =   2
      Left            =   2985
      PasswordChar    =   "*"
      TabIndex        =   5
      Text            =   "2"
      Top             =   1830
      Width           =   1410
   End
   Begin VB.TextBox txt 
      Height          =   300
      IMEMode         =   3  'DISABLE
      Index           =   1
      Left            =   2985
      PasswordChar    =   "*"
      TabIndex        =   3
      Text            =   "1"
      Top             =   1410
      Width           =   1410
   End
   Begin VB.TextBox txt 
      Enabled         =   0   'False
      Height          =   300
      Index           =   0
      Left            =   2985
      TabIndex        =   1
      Text            =   "0"
      Top             =   990
      Width           =   1410
   End
   Begin VB.Image Image2 
      Height          =   1440
      Left            =   4665
      Picture         =   "frmUtlCheckPassword.frx":058A
      Stretch         =   -1  'True
      Top             =   1035
      Width           =   1365
   End
   Begin VB.Shape Shape1 
      FillColor       =   &H00BDDFD1&
      FillStyle       =   0  'Solid
      Height          =   1455
      Left            =   6420
      Shape           =   4  'Rounded Rectangle
      Top             =   1020
      Width           =   1575
   End
   Begin VB.Label Label6 
      BackColor       =   &H00C00000&
      Caption         =   "  Utility User Security"
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
      TabIndex        =   10
      Top             =   0
      Width           =   8790
   End
   Begin VB.Shape Shape3 
      BorderColor     =   &H00C00000&
      BorderWidth     =   5
      FillColor       =   &H000000C0&
      Height          =   3180
      Left            =   0
      Top             =   45
      Width           =   8190
   End
   Begin VB.Label lblVerifyPassword 
      AutoSize        =   -1  'True
      BackColor       =   &H80000018&
      BackStyle       =   0  'Transparent
      Caption         =   "&Verify New Password  :"
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
      Index           =   3
      Left            =   465
      TabIndex        =   6
      Top             =   2250
      Width           =   2190
   End
   Begin VB.Label lblEnterNewPassword 
      AutoSize        =   -1  'True
      BackColor       =   &H80000018&
      BackStyle       =   0  'Transparent
      Caption         =   "Enter &New Password  :"
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
      Left            =   465
      TabIndex        =   4
      Top             =   1830
      Width           =   2130
   End
   Begin VB.Label lblEnterPassword 
      AutoSize        =   -1  'True
      BackColor       =   &H80000018&
      BackStyle       =   0  'Transparent
      Caption         =   "Enter Current &Password  :"
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
      Left            =   465
      TabIndex        =   2
      Top             =   1410
      Width           =   2415
   End
   Begin VB.Label lblUserName 
      AutoSize        =   -1  'True
      BackColor       =   &H80000018&
      BackStyle       =   0  'Transparent
      Caption         =   "User &Name   :"
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
      Left            =   465
      TabIndex        =   0
      Top             =   990
      Width           =   1275
   End
   Begin VB.Shape Shape2 
      BackColor       =   &H00808080&
      BorderColor     =   &H004D177D&
      FillColor       =   &H00EDFDFE&
      FillStyle       =   0  'Solid
      Height          =   2175
      Left            =   210
      Top             =   675
      Width           =   6060
   End
End
Attribute VB_Name = "frmUtlChangePassword"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim i As Long
Private Sub Form_Load()
Me.Top = 300
Me.Left = Screen.Width / 2 - Me.Width / 2
'-- Set Text Box Values
txt(0) = gUserName
For i = 1 To 3
    txt(i) = ""
Next i
End Sub

Private Sub txt_GotFocus(Index As Integer)
GProcSelectBox Me.ActiveControl
End Sub
Private Sub CmdClose_Click()
Unload Me
End Sub
Private Sub cmdOK_Click()
On Error GoTo ErrorRoutine
If ValidateData = True Then
    SaveProc
    '- Set IsLogOn=0
    gSetupCn.Execute "update tblMastUser set IsLogOn= 0 WHERE USERNAME ='" & gUserName & "'"
    End
Else
    txt(1).SetFocus
End If
Exit Sub
ErrorRoutine:
GProcErrorHandler
End Sub
'=== Procedure SaveProc (Change Password )
Private Sub SaveProc()
    '- Set Password
    gSetupCn.Execute "update tblMastUser set tblMastUser.Password= '" & Trim(txt(2)) & "' WHERE tblMastUser.USERNAME ='" & gUserName & "'"
End Sub
'=== Validate Date
Private Function ValidateData() As Boolean
Dim Rs1 As Recordset
ValidateData = True
'-- Current password
If gPassword <> Trim(txt(1)) Then
    MsgBox "Current password is not correct.", vbOKOnly, Me.Caption
    txt(1).SetFocus
    ValidateData = False
    Exit Function
End If
'-- New Password
If txt(2) = "" Then
    MsgBox "Enter new password.", vbOKOnly, Me.Caption
    txt(2).SetFocus
    ValidateData = False
    Exit Function
End If
'-- new password and varified password
If txt(2) <> txt(3) Then
    MsgBox "Entered new password and varified password are mismatched.", vbOKOnly, Me.Caption
    txt(3).SetFocus
    ValidateData = False
    Exit Function
End If
End Function

