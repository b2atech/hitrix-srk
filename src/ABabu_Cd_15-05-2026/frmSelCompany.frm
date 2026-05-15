VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmSelCompany 
   BackColor       =   &H80000000&
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   6660
   ClientLeft      =   255
   ClientTop       =   1410
   ClientWidth     =   8175
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "frmSelCompany.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmSelCompany.frx":058A
   ScaleHeight     =   6660
   ScaleWidth      =   8175
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdOk 
      Caption         =   "&Ok"
      Default         =   -1  'True
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
      Left            =   2520
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   5985
      Width           =   1125
   End
   Begin VB.CommandButton cmdExit 
      Cancel          =   -1  'True
      Caption         =   "E&xit"
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
      Left            =   4560
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   5985
      Width           =   1125
   End
   Begin MSComctlLib.ListView ListView1 
      Height          =   4425
      Left            =   840
      TabIndex        =   0
      Top             =   1080
      Width           =   6375
      _ExtentX        =   11245
      _ExtentY        =   7805
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   0
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Select Company . . ."
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   21.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00400000&
      Height          =   540
      Left            =   2280
      TabIndex        =   3
      Top             =   360
      Width           =   3240
   End
End
Attribute VB_Name = "frmSelCompany"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim mItem
Dim rstData As Recordset
Private Sub cmdExit_Click()
    'Unload frmMain
    'Unload Me
    'Call Main
    gSetupCn.Execute "update tblMastUser set IsLogOn=0 where UserName='" & gUserName & "'"
    If gSetupCn.State = adStateOpen Then gSetupCn.Close
   ' If gCn.State = adStateOpen Then gCn.Close
    End
End Sub
Private Sub cmdOK_Click()
'On Error GoTo ErrorHandler
'    gCName = ListView1.SelectedItem
'    mTempYr = ListView1.SelectedItem.SubItems(1)
    gCName = ListView1.SelectedItem.SubItems(1)
    gCCode = ListView1.SelectedItem.SubItems(2)
    gMyConn = "gCn"
        '& " " & mTempYr
    Dim rst1 As Recordset
    If GProcRstOpen(rst1, "Select * from tblCompanyYear order by compcode,compyear desc", "R", gCn) > 0 Then
        rst1.MoveFirst
        rst1.Find "CompCode='" & gCCode & "'"
        If Not rst1.EOF Then
            GProcGetCompDetails rst1
            Unload Me
        End If
    End If
Exit Sub
ErrorHandler:
    GProcErrorHandler
End Sub
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
Select Case KeyCode
Case 13:
    cmdOK_Click
Case 27:
    If Len(gDbNm) = 0 Then End
    Unload Me
End Select
End Sub
Private Sub Form_Load()
Dim k As Integer
k = GProcRstOpen(rstData, "Select CCompNo,CName,CCode From tblMastCompany Order by CCompNo ", "R", gCn)
'Call GProcClearForm(Me)
MakeColumns
GetTitles
End Sub
Private Sub MakeColumns()
   ListView1.ColumnHeaders.Clear
   ' Add ColumnHeaders.
   ListView1.ColumnHeaders.Add , , "No", 800
   ListView1.ColumnHeaders.Add , , "Company Name", 4000
'   ListView1.ColumnHeaders.Add , , "Financial Year", 1500
   ListView1.ColumnHeaders.Add , , "SUBDIR Name", 0
End Sub
Private Sub GetTitles()
ListView1.ListItems.Clear
Do Until rstData.EOF
    Set mItem = ListView1.ListItems.Add()
    mItem.text = rstData!CCompNo
    mItem.SubItems(1) = rstData!Cname & vbNullString
    mItem.SubItems(2) = rstData!CCode & vbNullString
    rstData.MoveNext
Loop
End Sub
Private Sub ListView1_DblClick()
    cmdOK_Click
    DoEvents
End Sub
