VERSION 5.00
Begin VB.Form frmUtlBackUpData 
   BackColor       =   &H00C0E0FF&
   BorderStyle     =   0  'None
   Caption         =   "Utility Back Up Data"
   ClientHeight    =   4635
   ClientLeft      =   0
   ClientTop       =   -45
   ClientWidth     =   6615
   ControlBox      =   0   'False
   Icon            =   "frmUtlBackupData.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   4635
   ScaleWidth      =   6615
   ShowInTaskbar   =   0   'False
   Begin VB.DirListBox DirBack 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1440
      Left            =   540
      TabIndex        =   5
      Top             =   1635
      Width           =   5385
   End
   Begin VB.DriveListBox drlBack 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Left            =   540
      TabIndex        =   4
      ToolTipText     =   "Select Drive To Backup Data"
      Top             =   1155
      Width           =   5385
   End
   Begin VB.CommandButton CmdAll 
      Caption         =   "&All Firms"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2760
      TabIndex        =   3
      ToolTipText     =   "To Save Entered Information"
      Top             =   3660
      Width           =   1605
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "&Current Firm"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1035
      TabIndex        =   2
      ToolTipText     =   "To Save Entered Information"
      Top             =   3660
      Width           =   1605
   End
   Begin VB.CommandButton CmdClose 
      Cancel          =   -1  'True
      Caption         =   "E&xit"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   4515
      Picture         =   "frmUtlBackupData.frx":058A
      TabIndex        =   1
      ToolTipText     =   "To Come Out "
      Top             =   3660
      Width           =   915
   End
   Begin VB.Label Label9 
      BackStyle       =   0  'Transparent
      Caption         =   "Select Drive  :"
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
      Left            =   465
      TabIndex        =   6
      Top             =   780
      Width           =   1470
   End
   Begin VB.Shape Shape1 
      FillColor       =   &H00BDDFD1&
      FillStyle       =   0  'Solid
      Height          =   840
      Left            =   675
      Shape           =   4  'Rounded Rectangle
      Top             =   3480
      Width           =   5265
   End
   Begin VB.Label Label6 
      BackColor       =   &H00C00000&
      Caption         =   "  Utility Back Up Data"
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
      TabIndex        =   0
      Top             =   15
      Width           =   8790
   End
   Begin VB.Shape Shape3 
      BorderColor     =   &H00C00000&
      BorderWidth     =   5
      FillColor       =   &H00C00000&
      Height          =   4620
      Left            =   0
      Top             =   15
      Width           =   6630
   End
   Begin VB.Shape Shape4 
      BorderColor     =   &H00000080&
      FillColor       =   &H00EDFDFE&
      FillStyle       =   0  'Solid
      Height          =   2745
      Left            =   225
      Top             =   615
      Width           =   6180
   End
End
Attribute VB_Name = "frmUtlBackUpData"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub Form_Load()
Dim fs, d
On Error GoTo ErrorHandler
    Me.Top = 50
    Me.Left = 6000 - (Me.Width / 2)
    Set fs = CreateObject("Scripting.FileSystemObject")
    Set d = fs.GetDrive(fs.GetDriveName(App.Path & "\HiBackUpData"))
    drlBack.Drive = d
    DirBack.Path = App.Path & "\HiBackUpData"
Exit Sub
ErrorHandler:
 If Err.Number = 68 Then
    drlBack.Drive = "C:"
    DirBack.Path = drlBack.Drive
 Else
    MsgBox Err.Description, vbInformation + vbOKOnly, Me.Caption
 End If
End Sub

Private Sub CmdClose_Click()
    Unload Me
End Sub
Private Sub drlBack_Change()
DirBack.Path = drlBack.Drive
End Sub
Private Sub cmdSave_Click()
    ProcBackUp
End Sub
Private Sub CmdAll_Click()
    ProcBackUp True
End Sub
Private Sub ProcBackUp(Optional AllFirmFlag As Boolean)
'''On Error GoTo ErrorHandler
'''Dim SourcePath As String
'''Dim TargetPath As String
'''Dim SetUpSPath As String
'''Dim SetUpTPath As String
'''Dim Rscomp As Recordset
'''Dim X As String
'''If DirBack.Path & "\" = gDataPath Then
'''    MsgBox "Check Path.Selected Folder is Data Folder.", vbInformation + vbOKOnly, "Utility Backup Data"
'''    Exit Sub
'''End If
'''With DirBack
'''    If MsgBox("Data will Backup to " & DirBack.Path & " folder.Are you Sure ?", vbOKCancel + vbInformation, Me.Caption) = vbOK Then
'''    '--- Back Up Setup Database
'''    SetUpSPath = gDataPath & "SetUp.mdb"
'''    SetUpTPath = DirBack.Path & "\Setup.mdb"
'''    If gSetupCn.State = adStateOpen Then
'''        gSetupCn.Close
'''    End If
'''    If Dir(SetUpTPath, vbDirectory) <> "" Then
'''        If MsgBox("Overwrite Data File Setup.mdb.", vbInformation + vbYesNo, "Utility Backup Data") = vbYes Then
'''            Kill SetUpTPath
'''            FileCopy SetUpSPath, SetUpTPath
'''        End If
'''    Else
'''        FileCopy SetUpSPath, SetUpTPath
'''    End If
'''    DoEvents
'''    '-- Backup Company Database
'''    gCn.Close
'''    If AllFirmFlag = False Then
'''        SourcePath = gDataPath & gDbNm & ".mdb"
'''        TargetPath = DirBack.Path & "\" & gDbNm & ".mdb"
'''        If Dir(TargetPath, vbDirectory) <> "" Then
'''            If MsgBox("Overwrite Data File " & gDbNm & ".mdb.", vbInformation + vbYesNo, "Utility Backup Data") = vbYes Then
'''                Kill TargetPath
'''                FileCopy SourcePath, TargetPath
'''            End If
'''        Else
'''            FileCopy SourcePath, TargetPath
'''        End If
'''        DoEvents
'''    Else '--- All Firm
'''        Call GProcConnectionOpen(gSetupCn)
'''        If GProcRstOpen(Rscomp, "select * from tblCompanyYear", "R", gSetupCn) > 0 Then
'''            Rscomp.MoveFirst
'''            Do While Not Rscomp.EOF
'''                SourcePath = gDataPath & Rscomp!CompCode & Rscomp!CompYear & ".mdb"
'''
'''                TargetPath = DirBack.Path & "\" & Rscomp!CompCode & Rscomp!CompYear & ".mdb"
'''                If Dir(SourcePath, vbDirectory) <> "" Then
'''                    If Dir(TargetPath, vbDirectory) <> "" Then
'''                        If MsgBox("Overwrite Data File " & Rscomp!CompCode & Rscomp!CompYear & ".mdb.", vbInformation + vbYesNo, "Utility Backup Data") = vbYes Then
'''                            Kill TargetPath
'''                            FileCopy SourcePath, TargetPath
'''                        End If
'''                    Else
'''                        FileCopy SourcePath, TargetPath
'''                    End If
'''                End If
'''                DoEvents
'''                Rscomp.MoveNext
'''            Loop
'''        End If
'''    End If
'''    'DBEngine.CompactDatabase mStrPath, msou
'''    '--- Set Connection
'''    Call GProcConnectionOpen(gCn)
'''    MsgBox "Data Backup is completed successfully.", vbOKOnly, Me.Caption
'''    Unload Me
'''    End If
''' End With
''' Exit Sub
'''ErrorHandler:
'''MsgBox Err.Description, vbInformation + vbOKOnly, Me.Caption
''''MsgBox "Exit and take backup by restarting the program again", vbInformation + vbOKOnly, Me.Caption
'''Call GProcConnectionOpen(gCn)
End Sub

