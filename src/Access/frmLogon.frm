VERSION 5.00
Begin VB.Form frmLogOn 
   BackColor       =   &H80000000&
   BorderStyle     =   0  'None
   Caption         =   "Log On"
   ClientHeight    =   5250
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   9330
   Icon            =   "frmLogon.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   Picture         =   "frmLogon.frx":058A
   ScaleHeight     =   5250
   ScaleWidth      =   9330
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton CmdLogOff 
      BackColor       =   &H00FFFFFF&
      Caption         =   "&Log Off"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   7410
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   2640
      Width           =   1335
   End
   Begin VB.TextBox txtLoginNm 
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      IMEMode         =   3  'DISABLE
      Left            =   6150
      MaxLength       =   10
      TabIndex        =   1
      Top             =   1050
      Width           =   2535
   End
   Begin VB.TextBox txtPassword 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      IMEMode         =   3  'DISABLE
      Left            =   6150
      MaxLength       =   15
      PasswordChar    =   "*"
      TabIndex        =   3
      Top             =   1770
      Width           =   2535
   End
   Begin VB.CommandButton cmdOk 
      BackColor       =   &H00FFFFFF&
      Caption         =   "&Ok"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   4230
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   2640
      Width           =   1335
   End
   Begin VB.CommandButton cmdCancel 
      BackColor       =   &H00FFFFFF&
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   5865
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   2640
      Width           =   1335
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "Hi-Trix Infotech Pvt Ltd"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000040C0&
      Height          =   315
      Left            =   3750
      TabIndex        =   9
      Top             =   4365
      Width           =   4245
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "Ichalkaranji - 416115,  +91 0230 2424933"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   195
      Left            =   3780
      TabIndex        =   8
      Top             =   4620
      Width           =   4425
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Hi-Trix Infotech Pvt Ltd"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   315
      Left            =   3750
      TabIndex        =   7
      Top             =   4365
      Width           =   4245
   End
   Begin VB.Label Label3 
      BackColor       =   &H80000009&
      BackStyle       =   0  'Transparent
      Caption         =   "Software By :"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   315
      Left            =   3720
      TabIndex        =   6
      Top             =   4110
      Width           =   3735
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00C00000&
      BorderWidth     =   5
      Height          =   5205
      Left            =   0
      Top             =   30
      Width           =   9330
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Login &Name       :"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   285
      Left            =   3540
      TabIndex        =   0
      Top             =   1050
      Width           =   1725
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Enter &Password  :"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   285
      Left            =   3510
      TabIndex        =   2
      Top             =   1770
      Width           =   1755
   End
End
Attribute VB_Name = "frmLogOn"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rstPass As Recordset
Dim cnt As Integer
Private Sub cmdCancel_Click()
    gLogOnSucc = False
    Unload Me
    End
End Sub
'-- Log Off
Private Sub CmdLogOff_Click()
If GProcRstOpen(rstPass, "Select * from tblMastUser where UserName='" & txtLoginNm.Text & "' and Password='" & txtPassword.Text & "'", "O", gSetupCn) > 0 Then
        If rstPass!IsLogOn = 1 Then
            '-- Set LogOn Status
            gSetupCn.Execute "update tblMastUser set IsLogOn=0 where UserName='" & txtLoginNm.Text & "'"
            MsgBox "User Successfully Logged Off.", vbOKOnly + vbInformation, "Log Off"
            cmdCancel_Click
        Else
            MsgBox "User already Logged Off.", vbOKOnly + vbInformation, "Log Off"
            cmdOk.SetFocus
        End If
Else
    cnt = cnt + 1
    If cnt = 3 Then
        cmdCancel_Click
        Exit Sub
    End If
    txtLoginNm.SetFocus
End If
rstPass.Close
End Sub
Private Sub cmdOK_Click()
If gSelectedMenu = "P" Then
    If GProcRstOpen(rstPass, "Select * from PASS", "O", gSetupCn) > 0 Then
        If txtPassword.Text = rstPass!Password Then
       '     frmLoginFaci.Show 1
        Else
            cnt = cnt + 1
            If cnt = 3 Then
                cmdCancel_Click
                Exit Sub
            End If
            MsgBox "You Have Entered Invalid Password", vbOKOnly + vbCritical, "Invalid Password"
            txtPassword.Text = ""
            txtPassword.SetFocus
        End If
    End If
    rstPass.Close
ElseIf Left(gSelectedMenu, 1) = "E" Then
    If GProcRstOpen(rstPass, "Select * from tblMastUser where tblMastUser.UserName='" & txtLoginNm.Text & "' and tblMastUser.Password='" & txtPassword.Text & "'", "O", gSetupCn) > 0 Then
        If rstPass!IsLogOn = 0 Then
            gUserName = rstPass!UserName & vbNullString
            gPassword = rstPass!Password & vbNullString
            gChkPassword = rstPass!ChkPassword & vbNullString
            gIsAdminUser = IIf(rstPass!isAdminUser = 1, True, False)
            'mUAccEntry = IIf(rstPass!ACCENTRY = True, 1, 0)
            gLogOnSucc = True
            '-- Set LogOn Status
            gSetupCn.Execute "update tblMastUser set IsLogOn=1 where UserName='" & gUserName & "'"
            rstPass.Close
            Unload Me
         Else
            MsgBox "User already Logged On.", vbOKOnly + vbInformation, "Log On"
            txtLoginNm.SetFocus
         End If
    Else
        cnt = cnt + 1
        If cnt = 3 Then
            cmdCancel_Click
            Exit Sub
        End If
        txtLoginNm.SetFocus
    End If
    If rstPass.State = adStateOpen Then rstPass.Close
End If
End Sub
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
Select Case KeyCode
Case 13:
    SendKeys "{TAB}"
Case 27:
    gLogOnSucc = False
    Unload Me
End Select
End Sub
Private Sub Form_Load()
If gSelectedMenu = "P" Then
    Label2.Visible = False
    txtLoginNm.Visible = False
End If
End Sub

Private Sub txtLoginNm_GotFocus()
If TypeOf Me.ActiveControl Is TextBox Then GProcSelectBox Me.ActiveControl
End Sub
Private Sub txtLoginNm_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then Exit Sub
End Sub

Private Sub txtPassword_GotFocus()
If TypeOf Me.ActiveControl Is TextBox Then GProcSelectBox Me.ActiveControl
End Sub

Private Sub txtPassword_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then Exit Sub
End Sub
