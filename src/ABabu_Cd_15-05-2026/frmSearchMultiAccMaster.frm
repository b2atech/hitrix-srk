VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form frmSearchMultiAccMaster 
   BackColor       =   &H00FCE2E3&
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   5460
   ClientLeft      =   45
   ClientTop       =   45
   ClientWidth     =   13680
   ControlBox      =   0   'False
   FillColor       =   &H80000005&
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   12
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H80000007&
   Icon            =   "frmSearchMultiAccMaster.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5460
   ScaleWidth      =   13680
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox txt 
      Height          =   375
      Index           =   0
      Left            =   270
      TabIndex        =   0
      Top             =   0
      Width           =   6660
   End
   Begin VB.CommandButton cmdOk 
      Caption         =   "&OK"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   450
      Left            =   5415
      TabIndex        =   2
      Top             =   4815
      Width           =   1215
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "E&XIT"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   450
      Left            =   6915
      TabIndex        =   3
      Top             =   4800
      Width           =   1215
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid msFlex 
      Height          =   4320
      Left            =   240
      TabIndex        =   1
      Top             =   360
      Width           =   13260
      _ExtentX        =   23389
      _ExtentY        =   7620
      _Version        =   393216
      BackColor       =   16777215
      Rows            =   3
      Cols            =   4
      FixedCols       =   0
      BackColorFixed  =   16576
      ForeColorFixed  =   15597054
      BackColorSel    =   16711680
      ForeColorSel    =   15597054
      BackColorBkg    =   15597054
      GridColor       =   15597054
      GridColorFixed  =   12648447
      GridColorUnpopulated=   12648447
      AllowBigSelection=   0   'False
      HighLight       =   2
      GridLines       =   0
      SelectionMode   =   1
      AllowUserResizing=   1
      FormatString    =   $"frmSearchMultiAccMaster.frx":0442
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _NumberOfBands  =   1
      _Band(0).Cols   =   4
   End
   Begin VB.Label lbl 
      Height          =   375
      Index           =   0
      Left            =   10350
      TabIndex        =   5
      Top             =   0
      Width           =   1035
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Search...."
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   15.75
         Charset         =   0
         Weight          =   900
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   510
      Left            =   450
      TabIndex        =   4
      Top             =   4800
      Width           =   3630
   End
End
Attribute VB_Name = "frmSearchMultiAccMaster"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'---------- Search Multi Form Variables
Dim mCaption
Dim mColWidth
Dim mQuery As String
'Dim mIsUnload As Boolean
Dim mCn As Connection
Dim mRetCol(4) As String
Dim mUpperHeading As Integer, mLowerHeading As Integer ''Upper & Lower bound of columns
Dim mIsTNameAsQry As Boolean
Dim mOrderCol As String
Dim mCriteria As String
Dim mIsAllowAdd As Boolean
Dim mIsOnlySearchMatch As Boolean
Dim mTName As String
Dim mRsSearch As New ADODB.Recordset
Dim mFieldname As String
Dim mMatchChar As String
Dim mLeft As Long
Dim mTop As Long
Dim mUnloadFlag As Boolean
Dim mGroupCol As String
'---------- Search Multi Field
Public Property Get RetCol() As Variant
    RetCol = mRetCol
End Property
'====== Name : SearchMultiField
'--- Variables : pTName (Table Name),pFieldname ( Array Fieldname),pCaptions ( Array colume Caption),pColWidths (Array Col width)
'-- pCriteria (condition), pMatchChar (Match Character),pOrderCol (order by cols),pfrmLeft ( Form Left Pos),pfrmTop  ( Form Top Pos)
'-- pIsAllowAdd (Allow to write New Value),pCon (connection Name),pIsTNameAsQry (Yes If 1 st parameter is Query)

Public Sub SearchMultiFieldAccMaster(Optional pMatchChar As String, Optional pfrmLeft As Long, Optional pfrmTop As Long)
    mMatchChar = pMatchChar
    mLeft = pfrmLeft
    mTop = pfrmTop
    msFlex.ColWidth(3) = 0
     
     mUnloadFlag = False
     frmSearchMultiAccMaster.txt(0) = pMatchChar
     frmSearchMultiAccMaster.txt(0).SelStart = Len(pMatchChar)
     If mUnloadFlag = False Then
     frmSearchMultiAccMaster.Show vbModal
        
     End If
End Sub

Private Sub cmdExit_Click()
mUnloadFlag = True
Unload Me
End Sub


Private Sub cmdOK_Click()
If msFlex.Row <> 0 Then
   mRetCol(0) = msFlex.TextMatrix(msFlex.Row, 0)
   mRetCol(1) = msFlex.TextMatrix(msFlex.Row, 1)
   mRetCol(2) = msFlex.TextMatrix(msFlex.Row, 2)
   mRetCol(3) = msFlex.TextMatrix(msFlex.Row, 3)
End If
mUnloadFlag = True
Unload Me
End Sub
Private Sub Form_KeyPress(KeyAscii As Integer)
Dim i As Long
Select Case KeyAscii
    Case 13 '-- Enter
        cmdOK_Click
    Case 27 '-- Esc
        cmdExit_Click
End Select
End Sub

Private Sub Form_Load()
Dim mRetCol(4) As String

mRetCol(0) = ""
mRetCol(1) = ""
mRetCol(2) = ""
mRetCol(3) = ""

If mLeft = 0 Then
    Me.Top = Screen.Height / 2 - Me.Height / 2
    Me.Left = Screen.Width / 2 - Me.Width / 2
Else
    Me.Left = mLeft
    Me.Top = mTop
End If
If mMatchChar <> "" Then txt(0) = mMatchChar Else gRstAccMast.MoveFirst: GridFill
If gCIBOffice = "I" Then
   msFlex.FormatString = "Account Name                                               |<Address                                                                            |<Contact Person                       |Code"
Else
   msFlex.FormatString = "Account Name                                               |<Address                                                                                                    |<A/c Code     |Code"
End If
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
   'Set frmSearchMulti = Nothing
End Sub

Private Sub msFlex_DblClick()
cmdOK_Click
End Sub
Private Sub SetSelRowColor()

End Sub

Private Sub msFlex_EnterCell()
Dim i As Long
If msFlex.rows > 1 Then
    For i = 0 To 1
        msFlex.Col = i
        msFlex.CellBackColor = vbBlue
        msFlex.CellForeColor = vbWhite
    Next i
End If
msFlex.Col = 0
msFlex.Refresh
End Sub

Private Sub msFlex_GotFocus()
msFlex_EnterCell
End Sub

Private Sub msFlex_LeaveCell()
Dim i As Long
If msFlex.rows > 1 Then
    For i = 0 To 1
        msFlex.Col = i
        msFlex.CellBackColor = vbWhite
        msFlex.CellForeColor = vbBlack
    Next i
End If
msFlex.Col = 0
msFlex.Refresh
End Sub

Private Sub txt_Change(Index As Integer)
gRstAccMast.MoveFirst
If txt(0) <> "" Then gRstAccMast.Find "AcNAme >='" & txt(0) & "'"
GridFill
End Sub
Private Sub txt_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
If KeyCode = 40 Then
    msFlex.SetFocus
End If
End Sub
Private Sub GridFill()
Dim RowIndex As Long
Dim i As Long
GridAddNew
  With gRstAccMast
    If Not .EOF Then
        RowIndex = 1
        msFlex.rows = 2
        Do While Not .EOF
           If txt(0) <> "" And UCase(Left(!AcName, Len(txt(0)))) <> UCase(txt(0)) Then Exit Sub
           i = i + 1
           If i > 40 Then Exit Sub
           msFlex.TextMatrix(RowIndex, 0) = .Fields(0) & vbNullString
           msFlex.TextMatrix(RowIndex, 1) = .Fields(1) & vbNullString
           If gCIBOffice = "I" Then
              msFlex.TextMatrix(RowIndex, 2) = .Fields(4) & vbNullString
           Else
              msFlex.TextMatrix(RowIndex, 2) = .Fields(3) & vbNullString
           End If
           msFlex.TextMatrix(RowIndex, 3) = .Fields(2) & vbNullString
           RowIndex = RowIndex + 1
           .MoveNext
           If Not .EOF Then msFlex.rows = msFlex.rows + 1
        Loop
    End If
 End With
End Sub
Private Sub GridAddNew()
With msFlex
    .TextMatrix(.Row, 0) = ""
    .TextMatrix(.Row, 1) = ""
    .TextMatrix(.Row, 2) = ""
    .TextMatrix(.Row, 3) = 0
End With
End Sub

