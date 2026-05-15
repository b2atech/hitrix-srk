VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form frmSearchMulti 
   BackColor       =   &H00C0E0FF&
   BorderStyle     =   0  'None
   Caption         =   "Search"
   ClientHeight    =   5430
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   11700
   ControlBox      =   0   'False
   FillColor       =   &H80000005&
   ForeColor       =   &H80000007&
   Icon            =   "frmSearchMulti.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5430
   ScaleWidth      =   11700
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox txt 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   0
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   1035
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
      Left            =   3660
      TabIndex        =   2
      Top             =   4530
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
      Left            =   5040
      TabIndex        =   3
      Top             =   4530
      Width           =   1215
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   735
      Left            =   10260
      Top             =   0
      Visible         =   0   'False
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   1296
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   8
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   ""
      OLEDBString     =   ""
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   ""
      Caption         =   "Adodc1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid msFlex 
      Bindings        =   "frmSearchMulti.frx":058A
      Height          =   3930
      Left            =   15
      TabIndex        =   1
      Top             =   390
      Width           =   11415
      _ExtentX        =   20135
      _ExtentY        =   6932
      _Version        =   393216
      Rows            =   5
      Cols            =   5
      FixedCols       =   0
      BackColorFixed  =   16576
      ForeColorFixed  =   -2147483634
      BackColorSel    =   16711680
      ForeColorSel    =   -2147483643
      GridColor       =   -2147483634
      AllowBigSelection=   0   'False
      HighLight       =   2
      GridLines       =   0
      SelectionMode   =   1
      AllowUserResizing=   1
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
      _Band(0).Cols   =   5
   End
   Begin VB.Shape Shape1 
      FillColor       =   &H00BDDFD1&
      FillStyle       =   0  'Solid
      Height          =   750
      Left            =   3420
      Shape           =   4  'Rounded Rectangle
      Top             =   4380
      Width           =   3105
   End
   Begin VB.Shape Shape4 
      BorderColor     =   &H00C00000&
      BorderWidth     =   5
      Height          =   5400
      Left            =   0
      Top             =   0
      Width           =   11640
   End
   Begin VB.Label lbl 
      BackColor       =   &H00C0E0FF&
      BackStyle       =   0  'Transparent
      Height          =   375
      Index           =   0
      Left            =   4335
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
      ForeColor       =   &H00000080&
      Height          =   510
      Left            =   405
      TabIndex        =   4
      Top             =   4290
      Width           =   3630
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H00000080&
      FillColor       =   &H00EDFDFE&
      FillStyle       =   0  'Solid
      Height          =   5265
      Left            =   0
      Top             =   0
      Width           =   11565
   End
End
Attribute VB_Name = "frmSearchMulti"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'---------- Search Multi Form Variables
Dim mRsSearch As Recordset
Dim mCaption
Dim mColWidth
Dim mQuery As String
'Dim mIsUnload As Boolean
Dim mCn As Connection
Dim mRetCol() As String
Dim mUpperHeading As Integer, mLowerHeading As Integer ''Upper & Lower bound of columns
Dim mIsTNameAsQry As Boolean
Dim mOrderCol As String
Dim mCriteria As String
Dim mIsAllowAdd As Boolean
Dim mIsOnlySearchMatch As Boolean
Dim mTName As String
Dim mFieldname As String
Dim mMatchChar As String
Dim mleft As Long
Dim mTop As Long
Dim mUnloadFlag As Boolean
'---------- Search Multi Field
Public Property Get RetCol() As Variant
    RetCol = mRetCol
End Property
'====== Name : SearchMultiField
'--- Variables : pTName (Table Name),pFieldname ( Array Fieldname),pCaptions ( Array colume Caption),pColWidths (Array Col width)
'-- pCriteria (condition), pMatchChar (Match Character),pOrderCol (order by cols),pfrmLeft ( Form Left Pos),pfrmTop  ( Form Top Pos)
'-- pIsAllowAdd (Allow to write New Value),pCon (connection Name),pIsTNameAsQry (Yes If 1 st parameter is Query)
Public Sub SearchMultiField(pTName As String, pFieldname As String, pCaptions, Optional pColWidths, Optional pCriteria As String, Optional pMatchChar As String, Optional pOrderCol As String, Optional pfrmLeft As Long, Optional pfrmTop As Long, Optional pIsAllowAdd As Boolean, Optional pCon As Connection, Optional pIsTNameAsQry As Boolean, Optional pIsOnlyMatch As Boolean)
    mTName = pTName
    mFieldname = pFieldname
    mOrderCol = pOrderCol
    mCriteria = pCriteria
    mIsAllowAdd = pIsAllowAdd
    mIsTNameAsQry = pIsTNameAsQry
    mCaption = pCaptions
    mUpperHeading = UBound(pCaptions)
    mLowerHeading = LBound(pCaptions)
    mIsOnlySearchMatch = pIsOnlyMatch
    mMatchChar = pMatchChar
    mleft = pfrmLeft
    mTop = pfrmTop
    '-- Set Query
    If mIsTNameAsQry = True Then
        mQuery = pTName & IIf(Trim(mCriteria) = "", "", " where " & mCriteria) '& IIf(Trim(mOrderCol) = "", "", " order by " & mOrderCol)
    Else
        mQuery = "select " & pFieldname & " from " & pTName & " " & IIf(Trim(mCriteria) = "", "", " where " & pCriteria) '& IIf(Trim(mOrderCol) = "", "", " order by " & mOrderCol)
    End If
    '-- Set connection
    If pCon Is Nothing Then
        Set mCn = gCn
    Else
        Set mCn = pCon
    End If
'    '-- Width
    If Not IsMissing(pColWidths) Then mColWidth = pColWidths
'    frmSearchMulti.txt(0) = pMatchChar
'    frmSearchMulti.txt(0).SelStart = Len(pMatchChar)
'    If pfrmLeft = 0 Then
'        frmSearchMulti.Top = Screen.Height / 2 - frmSearchMulti.Height / 2
'        frmSearchMulti.Left = Screen.Width / 2 - frmSearchMulti.Width / 2
'    Else
'        frmSearchMulti.Left = pfrmLeft
'        frmSearchMulti.Top = pfrmTop
'    End If
     mUnloadFlag = False
     frmSearchMulti.txt(0) = pMatchChar
     'If mUnloadFlag = True Then Exit Sub
     frmSearchMulti.txt(0).SelStart = Len(pMatchChar)
     If mUnloadFlag = False Then
        frmSearchMulti.Show vbModal
     End If
'     If pMatchChar <> "" And mUnloadFlag = False Then
'        frmSearchMulti.txt(0) = pMatchChar
'        frmSearchMulti.txt(0).SelStart = Len(pMatchChar)
'     End If
End Sub
Private Sub cmdExit_Click()
mUnloadFlag = True
Unload Me
End Sub

Private Sub cmdOK_Click()
Dim i As Long
ReDim mRetCol(mUpperHeading)
If msFlex.Row <> 0 Then
    For i = 0 To mRsSearch.Fields.Count - 1
        'If Len(txt(i)) > 0 Then
        msFlex.Col = i
        mRetCol(i) = msFlex.Text
        'End If
    Next i
ElseIf mIsAllowAdd = True Then '-- Allow Add Record
    If mRsSearch.EOF And mRsSearch.BOF Then
        mRetCol(0) = GProcProperCase(txt(0))
    Else
        mRsSearch.MoveFirst
        mRsSearch.Find mRsSearch.Fields(0).Name & "='" & txt(0) & "'"
        If Not mRsSearch.EOF Then
            For i = 0 To mRsSearch.Fields.Count - 1
                'If Len(txt(i)) > 0 Then
                msFlex.Col = i
                mRetCol(i) = msFlex.Text
                'End If
            Next i
        Else
            mRetCol(0) = GProcProperCase(txt(0))
        End If
    End If
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
On Error GoTo ErrorHandler
Dim x As String
Static MaxIdx               '-----Largest index in array.
ReDim mRetCol(1)
'If mUnloadFlag = True Then
'    Unload Me
'    Exit Sub
'End If
mRetCol(0) = ""
mRetCol(1) = ""
MaxIdx = 1
If mleft = 0 Then
    Me.Top = Screen.Height / 2 - Me.Height / 2
    Me.Left = Screen.Width / 2 - Me.Width / 2
Else
    Me.Left = mleft
    Me.Top = mTop
End If
x = mQuery & IIf(Trim(mOrderCol) = "", "", " order by " & mOrderCol)
GProcRstOpen mRsSearch, x, "R", mCn
Adodc1.ConnectionString = mCn
Adodc1.RecordSource = x
If mIsOnlySearchMatch = False Then
    Adodc1.Refresh
Else
    msFlex.Rows = 2
    msFlex.Cols = mRsSearch.Fields.Count
End If
'If mIsOnlyMatch = False Then
'Adodc1.Refresh
Loop1:
    If MaxIdx > mRsSearch.Fields.Count - 1 Then
        SetSearchCtrlWidth
        Exit Sub
    Else
        Load txt(MaxIdx)    '-----Create new item in array.
        txt(MaxIdx).Top = txt(MaxIdx - 1).Top
        'txt(MaxIdx).Left = txt(MaxIdx - 1).Left + txt(MaxIdx - 1).Width + 20
        txt(MaxIdx).Visible = True
        If MaxIdx = 1 Then
           ' If txt(0).Visible = False Then txt(1).TabIndex = txt(0).TabIndex
          '  txt(MaxIdx).SetFocus
        End If
        MaxIdx = MaxIdx + 1
    End If
GoTo Loop1
Exit Sub
ErrorHandler:
 GProcErrorHandler
End Sub

Private Sub msFlex_DblClick()
'SetSelRowColor
cmdOK_Click
End Sub
Private Sub SetSelRowColor()
'Dim i As Long
'If msFlex.Rows > 1 Then
'    For i = 0 To (mRsSearch.Fields.Count - 1)
'        msFlex.Col = i
'        msFlex.CellBackColor = vbBlue
'        msFlex.CellForeColor = vbWhite
'    Next i
'End If
'msFlex.Col = 0
End Sub

Private Sub msFlex_EnterCell()
Dim i As Long
If msFlex.Rows > 1 Then
    For i = 0 To (mRsSearch.Fields.Count - 1)
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
If msFlex.Rows > 1 Then
    For i = 0 To (mRsSearch.Fields.Count - 1)
        msFlex.Col = i
        msFlex.CellBackColor = vbWhite
        msFlex.CellForeColor = vbBlack
    Next i
End If
msFlex.Col = 0
msFlex.Refresh
End Sub

Private Sub txt_Change(Index As Integer)
On Error GoTo ErrorRoutine
Dim QStr As String
Dim i As Long
Dim QueryStr As String
Dim tblName As String
Dim strLen As Long
Dim ColString As String
Dim OrderByStr  As String
'If gIsSearchMultiUnload = True Then Exit Sub
ColString = ""
OrderByStr = Trim(mOrderCol)
For i = 0 To mRsSearch.Fields.Count - 1
    If txt(i) <> "" Then
        Select Case mRsSearch.Fields(i).Type
            Case adDouble, adInteger, 2, 3, adNumeric, adDecimal 'dbLong, dbDouble
                If IsNumeric(txt(i)) = True Then
                    If QStr = "" Then
                        QStr = mRsSearch.Fields(i).Name & "= " & txt(i)
                    Else
                        QStr = QStr & " AND " & mRsSearch.Fields(i).Name & "= " & txt(i)
                    End If
                End If
            Case Else
                If QStr = "" Then
                    If gBackEndDB = gBackEndAccess Or mCn = gSetupCn Then
                       QStr = "left([" & mRsSearch.Fields(i).Name & "]," & Len(txt(i)) & ")='" & txt(i) & "'"
                    Else '-- Oracle
                       QStr = "LPad(upper(" & mRsSearch.Fields(i).Name & ")," & Len(txt(i)) & ")=UPPER('" & txt(i) & "')"
                    End If
                Else
                    If gBackEndDB = gBackEndAccess Or mCn = gSetupCn Then
                       QStr = QStr & " AND left([" & mRsSearch.Fields(i).Name & "]," & Len(txt(i)) & ")='" & txt(i) & "'"
                    Else '-- Oracle
                       QStr = QStr & " AND LPad(upper(" & mRsSearch.Fields(i).Name & ")," & Len(txt(i)) & ")=upper('" & txt(i) & "')"
                    End If
                End If
        End Select
        If OrderByStr = "" Then '--Order by
            OrderByStr = mRsSearch.Fields(i).Name
        Else
            OrderByStr = OrderByStr & " ," & mRsSearch.Fields(i).Name
        End If
    End If
'    If txt(i) <> "" Then
'        Select Case mRsSearch.Fields(i).Type
'            Case adDouble, adInteger, 2, 3, adNumeric, adDecimal 'dbLong, dbDouble
'                If IsNumeric(txt(i)) = True Then
'                    If QStr = "" Then
'                        QStr = "[" & mRsSearch.Fields(i).Name & "]= " & txt(i)
'                    Else
'                        QStr = QStr & " AND [" & mRsSearch.Fields(i).Name & "]= " & txt(i)
'                    End If
'                End If
'            Case Else
'                If QStr = "" Then
'                    QStr = StrForLeft & "([" & mRsSearch.Fields(i).Name & "]," & Len(txt(i)) & ")='" & txt(i) & "'"
'                Else
'                    QStr = QStr & " AND " & StrForLeft & "([" & mRsSearch.Fields(i).Name & "]," & Len(txt(i)) & ")='" & txt(i) & "'"
'                End If
'        End Select
'        If OrderByStr = "" Then '--Order by
'            OrderByStr = "[" & mRsSearch.Fields(i).Name & "]"
'        Else
'            OrderByStr = OrderByStr & " ," & "[" & mRsSearch.Fields(i).Name & "]"
'        End If
'    End If
Next i
If QStr = "" Then
    If mIsOnlySearchMatch = False Then
        QueryStr = mQuery & IIf(Trim(OrderByStr) = "", "", " order by " & OrderByStr)
    Else
        QueryStr = ""
    End If
Else
    If mIsTNameAsQry = True Then '--Query
        strLen = InStr(LCase(mQuery), "where")
        If strLen = 0 Then
            QueryStr = mQuery & " where " & QStr & IIf(Trim(OrderByStr) = "", "", " order by " & OrderByStr)
        Else
            QueryStr = mQuery & " and  " & QStr & IIf(Trim(OrderByStr) = "", "", " order by " & OrderByStr)
        End If
   Else '--Tables
'        For i = 0 To Adodc1.Recordset.Fields.Count - 1
'            If ColString = "" Then
'                ColString = "[" & Adodc1.Recordset.Fields(i).Name & "]"
'            Else
'                ColString = ColString & ",[" & Adodc1.Recordset.Fields(i).Name & "]"
'            End If
'        Next i
        ColString = mFieldname
        tblName = mTName
        'tblName = FindTblName(Adodc1.RecordSource)
        strLen = InStr(LCase(mQuery), "where")
        If strLen = 0 Then
            QueryStr = "select " & ColString & " from " & tblName & " where " & QStr & IIf(Trim(OrderByStr) = "", "", " order by " & OrderByStr)
        Else
            QueryStr = "select " & ColString & " from " & tblName & IIf(Trim(mCriteria) = "", "", " where " & mCriteria) & " and  " & QStr & IIf(Trim(OrderByStr) = "", "", " order by " & OrderByStr)
        End If
   End If
End If
If QueryStr <> "" Then
    Adodc1.RecordSource = QueryStr
    Adodc1.Refresh
Else
    msFlex.Clear
    msFlex.Rows = 2
End If
SetSearchCtrlWidth
'If mUnloadFlag = True Then Unload Me
Exit Sub
ErrorRoutine:
    GProcErrorHandler
End Sub
'------
Private Sub SetSearchCtrlWidth()
Dim i As Long
For i = 0 To mRsSearch.Fields.Count - 1 'msFlex.Cols - 1
    msFlex.Row = 0
    msFlex.Col = i
    msFlex.Text = mCaption(i)
    msFlex.ColWidth(i) = mColWidth(i)
    txt(i).Width = mColWidth(i)
    If mColWidth(i) = 0 Then txt(i).Visible = False
    If i > 0 Then
        If mColWidth(i - 1) = 0 Then
            txt(i).Left = txt(i - 1).Left
        Else
            txt(i).Left = txt(i - 1).Left + txt(i - 1).Width + 20
        End If
    End If
    msFlex.Col = 0
Next i
If msFlex.Rows > 1 Then
    msFlex.Row = 1
'    'Dim i As Long
    If msFlex.Rows > 1 Then
        For i = 0 To (mRsSearch.Fields.Count - 1)
            msFlex.Col = i
            msFlex.CellBackColor = vbBlue
            msFlex.CellForeColor = vbWhite
        Next i
    End If
    msFlex.Col = 0
    msFlex.Refresh
'    'SetSelRowColor
End If
End Sub
Private Sub txt_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
If KeyCode = 40 Then
    msFlex.SetFocus
End If
End Sub

Private Sub txt_KeyPress(Index As Integer, KeyAscii As Integer)
Select Case mRsSearch.Fields(Index).Type
    Case adDouble, adInteger, 2, 3, adNumeric, adDecimal
        KeyAscii = GProcNumberOnly(KeyAscii)
End Select
End Sub
Private Function FindTblName(QStr As String)
Dim x As Long
Dim Y As Long
Dim z As Long
Dim tempqstr As String
If InStr(1, UCase(QStr), "FROM") = 0 Then
    FindTblName = QStr
Else
    x = InStr(1, UCase(QStr), " FROM ")
    Y = InStr(x, UCase(QStr), "WHERE")
    If Y > 0 Then
        FindTblName = Mid(QStr, x + 6, Y - x - 6)
    Else
        FindTblName = Right(QStr, Len(QStr) - (x + 5))
    End If
    z = InStr(1, UCase(FindTblName), "ORDER BY")
    If z > 0 Then
        FindTblName = Left(FindTblName, z - 1)
    End If
End If
End Function


