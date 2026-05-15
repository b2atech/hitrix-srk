VERSION 5.00
Begin VB.Form frmUtlRestoreData 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "."
   ClientHeight    =   2895
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6210
   ControlBox      =   0   'False
   Icon            =   "frmUtlRestoreData.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   2895
   ScaleWidth      =   6210
   Begin VB.Frame fraButtons 
      ForeColor       =   &H00C0C0C0&
      Height          =   870
      Left            =   30
      TabIndex        =   6
      Top             =   1995
      Width           =   6105
      Begin VB.CommandButton cmdSave 
         Caption         =   "Current Firm"
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
         Left            =   465
         Picture         =   "frmUtlRestoreData.frx":0442
         TabIndex        =   2
         ToolTipText     =   "To Save Entered Information"
         Top             =   225
         Width           =   1605
      End
      Begin VB.CommandButton cmdBack 
         CausesValidation=   0   'False
         Height          =   495
         Left            =   4980
         Picture         =   "frmUtlRestoreData.frx":074C
         Style           =   1  'Graphical
         TabIndex        =   4
         ToolTipText     =   "To Come Out "
         Top             =   225
         Width           =   915
      End
   End
   Begin VB.Frame fraBackUp 
      Height          =   1440
      Left            =   30
      TabIndex        =   1
      Top             =   525
      Width           =   6105
      Begin VB.DriveListBox drlBack 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   15.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   480
         Left            =   2940
         TabIndex        =   0
         Top             =   555
         Width           =   2910
      End
      Begin VB.Label lblDrive 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Restore From Drive"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   135
         TabIndex        =   5
         Top             =   615
         Width           =   2700
      End
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Restore Data"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   540
      Left            =   1710
      TabIndex        =   3
      Top             =   0
      Width           =   2670
   End
End
Attribute VB_Name = "frmUtlRestoreData"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub cmdBack_Click()
    Unload Me
End Sub

Private Sub cmdSave_Click()
 Dim mStrPath As String
 Dim mStrBack As String
If MsgBox("Restore Data For Firm " & gCompName & " For Year " & gDbYr, vbOKCancel + vbInformation, Me.Caption) = vbOK Then
    With drlBack
        If Dir(Left(.Drive, 2) & "\Back" & gDbNm, vbDirectory) = "" Then
            mErr.ErrorMsg "Directory " & Left(.Drive, 2) & "\Back" & gDbNm & " Not Present , Create New Directory", gCnm, 1, Array("CANCLE")
            Exit Sub
        End If
        mStrPath = mAppPath & "\" & gCcode & gDnm & ".mdb"
        mStrBack = Left(App.Path, 2) & "\BackVer\" & gCcode & "Year" & gDnm & Format(Now, "dd_MM_yyyy_HH_mm_SS") & ".mdb.BackVer"
        If Dir(Left(App.Path, 2) & "\BackVer", vbDirectory) = "" Then
            FileSystem.MkDir Left(App.Path, 2) & "\BackVer"
        End If
        gDB.Close
        
        'Copy C Ver To O Ver
        On Error Resume Next
        DBEngine.CompactDatabase mStrPath, mStrBack
        
        'To Kill Old Backup
        On Error Resume Next
        Kill mAppPath & "\" & gCcode & gDnm & ".mdb"
        
        On Error GoTo LocalErr
        'iN bETWWEN OPTIONS ARE       dbLangGeneral , dbEncrypt
        Dim msouFrom
        msouFrom = Left(.Drive, 2) & "\Back" & gDnm & "\" & gCcode & gDnm & "Back.mdb"
        'DBEngine.CompactDatabase mStrPath, msou, dbLangGeneral, , ";pwd=" & mPsIn
        DBEngine.CompactDatabase msouFrom, mStrPath
        DoEvents
        End
    End With
 End If
    Exit Sub
LocalErr:
      
      MsgBox Err.Number & "  -  " & Err.Description, vbCritical + vbOKOnly, "Data Back Up Error ...."
      End
End Sub
Private Sub Form_Load()
   Me.Top = 50
   Me.Left = 6000 - (Me.Width / 2)
End Sub
