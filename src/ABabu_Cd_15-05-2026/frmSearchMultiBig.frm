VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form frmSearchMultiBig 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   5460
   ClientLeft      =   45
   ClientTop       =   45
   ClientWidth     =   14685
   ControlBox      =   0   'False
   FillColor       =   &H80000005&
   ForeColor       =   &H80000007&
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5460
   ScaleWidth      =   14685
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox txt 
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
      Left            =   5415
      TabIndex        =   2
      Top             =   4800
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
      Left            =   15
      TabIndex        =   1
      Top             =   390
      Width           =   14610
      _ExtentX        =   25770
      _ExtentY        =   7620
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
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _NumberOfBands  =   1
      _Band(0).Cols   =   5
   End
   Begin VB.Label lbl 
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
      ForeColor       =   &H000000C0&
      Height          =   510
      Left            =   90
      TabIndex        =   4
      Top             =   4710
      Width           =   3630
   End
End
Attribute VB_Name = "frmSearchMultiBig"
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
Dim mRetCol() As String
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
Public Sub SearchMultiField(pTName As String, pFieldname As String, pCaptions, Optional pColWidths, Optional pCriteria As String, Optional pMatchChar As String, Optional pOrderCol As String, Optional pfrmLeft As Long, Optional pfrmTop As Long, Optional pIsAllowAdd As Boolean, Optional pCon As Connection, Optional pIsTNameAsQry As Boolean, Optional pIsOnlyMatch As Boolean, Optional pGroupCol As String)
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
    mLeft = pfrmLeft
    mTop = pfrmTop
    mColWidth = pColWidths
    mGroupCol = pGroupCol
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
    
    msFlex.FormatString = Mid(Left(mQuery, InStr(UCase(mQuery), " FROM")), 8)
    ' msFlex.FormatString = mQuery

     SetSearchCtrlWidth
     mUnloadFlag = False
     frmSearchMultiBig.txt(0) = pMatchChar
     frmSearchMultiBig.txt(0).SelStart = Len(pMatchChar)
     If mUnloadFlag = False Then
     frmSearchMultiBig.Show vbModal
        
     End If
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
        mRetCol(i) = Trim(msFlex.Text)
    Next i
    If mRetCol(0) = "" And mIsAllowAdd = True Then mRetCol(0) = GProcProperCase(txt(0))
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
'On Error GoTo ErrorHandler
Dim X As String
'Dim mRsSearch As New ADODB.New ADODB.recordset
Static MaxIdx               '-----Largest index in array.
ReDim mRetCol(1)
'If mUnloadFlag = True Then
'    Unload Me
'    Exit Sub
'End If
mRetCol(0) = ""
mRetCol(1) = ""
MaxIdx = 1
If mLeft = 0 Then
    Me.Top = Screen.Height / 2 - Me.Height / 2
    Me.Left = Screen.Width / 2 - Me.Width / 2
Else
    Me.Left = mLeft
    Me.Top = mTop
End If
X = mQuery
If Trim(mGroupCol) <> "" Then X = X + " Group by " & mGroupCol & ""
If Trim(mOrderCol) <> "" Then X = X + " Order by " & mOrderCol
'x = x + mQuery & IIf(Trim(mOrderCol) = "", "", " order by " & mOrderCol)
GProcRstOpen mRsSearch, X, "R", gCn
msFlex.FormatString = mFieldname
If mMatchChar = "" Then
    msFlex.cols = mRsSearch.Fields.Count
    GridFill
End If

If mIsOnlySearchMatch = False Then
    
Else
    msFlex.rows = 2
    msFlex.cols = mRsSearch.Fields.Count
    End If

Loop1:
    If MaxIdx > mRsSearch.Fields.Count - 1 Then
        SetSearchCtrlWidth
        Exit Sub
    Else
    
        Load txt(MaxIdx)    '-----Create new item in array.
        txt(MaxIdx).Top = txt(MaxIdx - 1).Top

        txt(MaxIdx).Visible = True
        If MaxIdx = 1 Then
        End If
        MaxIdx = MaxIdx + 1
    End If
GoTo Loop1
Exit Sub
ErrorHandler:
 GProcErrorHandler
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
If msFlex.rows > 1 Then
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
'On Error GoTo ErrorRoutine
Dim QStr As String
Dim i As Long, j As Long
Dim QueryStr As String
Dim tblName As String
Dim strLen As Long
Dim ColString As String, Mystr As String
Dim OrderByStr  As String, GroupByStr  As String
'If gIsSearchMultiUnload = True Then Exit Sub
ColString = ""
GroupByStr = Trim(mGroupCol)
OrderByStr = Trim(mOrderCol)
For i = 0 To mRsSearch.Fields.Count - 1
    If txt(i) <> "" Then
        Select Case mRsSearch.Fields(i).Type
            Case adDouble, adInteger, 2, 3, adNumeric, adDecimal 'dbLong, dbDouble
                If IsNumeric(txt(i)) = True Then
                    Mystr = mRsSearch.Fields(i).Name
                    QStr = IIf(QStr <> "", QStr & " AND ", "") + "" & Mystr & " like '" & txt(i) & "%'"
                    'If Trim(QStr) = "" Then
                    '    QStr = "left(" & mRsSearch.Fields(i).Name & "," & Len(txt(i)) & ")='" & txt(i) & "'"
                    'Else
                    '    QStr = QStr & " AND left(" & mRsSearch.Fields(i).Name & "," & Len(txt(i)) & ")='" & txt(i) & "'"
                    'End If
                End If
            
            Case Else
                If mRsSearch.Fields(i).Name <> "" Then
                   Mystr = mRsSearch.Fields(i).Name
                Else
                   Mystr = Mid(mQuery, 7)
                   For j = 1 To i
                      Mystr = Mid(Mystr, InStr(Mystr, " ,") + 2)  ', InStr(mystr, " ,"))
                   Next j
                   If InStr(Mystr, " ,") > 0 Then
                        Mystr = Left(Mystr, InStr(Mystr, " ,") - 1)
                   Else
                        Mystr = Left(Mystr, InStr(Mystr, " from") - 1)
                   End If
                End If

                'QStr = IIf(QStr <> "", " AND ", "") + "left(" & mystr & "," & Len(txt(i)) & ")='" & txt(i) & "'"
                QStr = IIf(QStr <> "", QStr & " AND ", "") + " charindex('" & txt(i) & "'," & Mystr & ") > 0 "

''                If Trim(QStr) = "" Then
''                    QStr = "left(" & mRsSearch.Fields(i).Name & "," & Len(txt(i)) & ")='" & txt(i) & "'"
''                Else
''                    QStr = QStr & " AND left(" & mRsSearch.Fields(i).Name & "," & Len(txt(i)) & ")='" & txt(i) & "'"
''                End If
        End Select
        If OrderByStr = "" Then '--Order by
           OrderByStr = mRsSearch.Fields(i).Name
        Else
            If (InStr(1, OrderByStr, mRsSearch.Fields(i).Name)) = 0 Then
                OrderByStr = OrderByStr & " ," & mRsSearch.Fields(i).Name
            End If
        End If
    End If

Next i
If QStr = "" Then
    If mIsOnlySearchMatch = False Then
        QueryStr = mQuery & IIf(Trim(GroupByStr) = "", "", " Group by " & GroupByStr) & IIf(Trim(OrderByStr) = "", "", " order by " & OrderByStr)
    Else
        QueryStr = ""
    End If
Else
    If mIsTNameAsQry = True Then '--Query
        strLen = InStr(LCase(mQuery), "where")
        If strLen = 0 Then
            QueryStr = mQuery & " where " & QStr & IIf(Trim(GroupByStr) = "", "", " Group by " & GroupByStr) & IIf(Trim(OrderByStr) = "", "", " order by " & OrderByStr)
        Else
            QueryStr = mQuery & " and  " & QStr & IIf(Trim(GroupByStr) = "", "", " Group by " & GroupByStr) & IIf(Trim(OrderByStr) = "", "", " order by " & OrderByStr)
        End If
   Else '--Tables

        ColString = mFieldname
        tblName = mTName

        strLen = InStr(LCase(mQuery), "where")
        If strLen = 0 Then
            QueryStr = "select " & ColString & " from " & tblName & " where " & QStr & IIf(Trim(GroupByStr) = "", "", " Group by " & GroupByStr) & IIf(Trim(OrderByStr) = "", "", " order by " & OrderByStr)
        Else
            QueryStr = "select " & ColString & " from " & tblName & IIf(Trim(mCriteria) = "", "", " where " & mCriteria) & " and  " & QStr & IIf(Trim(GroupByStr) = "", "", " Group by " & GroupByStr) & IIf(Trim(OrderByStr) = "", "", " order by " & OrderByStr)
        End If
   End If
End If
If QueryStr <> "" Then
    msFlex.Clear
    msFlex.rows = 2
   If GProcRstOpen(mRsSearch, QueryStr, "R", gCn) > 0 Then GridFill
Else
    msFlex.Clear
    msFlex.rows = 2
End If
SetSearchCtrlWidth
'Exit Sub
'ErrorRoutine:
'    GProcErrorHandler
End Sub
'------
Private Sub SetSearchCtrlWidth()
Dim i As Long
For i = 0 To mRsSearch.Fields.Count - 1 'msFlex.Cols - 1
    msFlex.cols = mRsSearch.Fields.Count
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


If msFlex.rows > 1 Then
    msFlex.Row = 1
'    'Dim i As Long
    If msFlex.rows > 1 Then
        For i = 0 To (mRsSearch.Fields.Count - 1)
            msFlex.Col = i
            msFlex.CellBackColor = vbBlue
            msFlex.CellForeColor = vbWhite
        Next i
    End If
    msFlex.Col = 0
   ' msFlex.Refresh
'    'SetSelRowColor
End If
End Sub
Private Sub txt_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
If KeyCode = 40 Then
    msFlex.SetFocus
End If
End Sub

Private Sub txt_KeyPress(Index As Integer, KeyAscii As Integer)
If KeyAscii = 39 Then KeyAscii = vbEmpty
Select Case mRsSearch.Fields(Index).Type
    Case adDouble, adInteger, 2, 3, adNumeric, adDecimal
        KeyAscii = GProcNumberOnly(KeyAscii)
        
End Select
End Sub
Private Function FindTblName(QStr As String)
Dim X As Long
Dim Y As Long
Dim z As Long
Dim tempqstr As String
If InStr(1, UCase(QStr), "FROM") = 0 Then
    FindTblName = QStr
Else
    X = InStr(1, UCase(QStr), " FROM ")
    Y = InStr(X, UCase(QStr), "WHERE")
    If Y > 0 Then
        FindTblName = Mid(QStr, X + 6, Y - X - 6)
    Else
        FindTblName = Right(QStr, Len(QStr) - (X + 5))
    End If
    z = InStr(1, UCase(FindTblName), "ORDER BY")
    If z > 0 Then
        FindTblName = Left(FindTblName, z - 1)
    End If
End If
End Function
Private Sub GridFill()
Dim RowIndex As Long
Dim ColIndex As Long
Dim i As Long
GridAddNew
  With mRsSearch
    If Not .EOF Then
        .MoveFirst
        RowIndex = 1
        msFlex.rows = 2
        Do While Not .EOF
            i = i + 1
          If i > 40 Then
            Exit Sub
          End If
             For ColIndex = 0 To .Fields.Count - 1
                Select Case .Fields(ColIndex).Type
                    Case adInteger, 2, 3 '-- Integer
                        msFlex.TextMatrix(RowIndex, ColIndex) = IIf(IsNull(.Fields(ColIndex)), 0, .Fields(ColIndex))
                    Case adDecimal, 131, adNumeric, adDouble  '-- Double
                        msFlex.TextMatrix(RowIndex, ColIndex) = Format(.Fields(ColIndex), GProcNumberFormat(.Fields(ColIndex).Precision, .Fields(ColIndex).NumericScale))
                    Case Else
                        'msFlex.ColAlignment(ColIndex) = vbLeftJustify
                        msFlex.TextMatrix(RowIndex, ColIndex) = "  " & .Fields(ColIndex) & vbNullString
                End Select
            Next ColIndex
            
            RowIndex = RowIndex + 1
            .MoveNext
            If Not .EOF Then msFlex.rows = msFlex.rows + 1
        Loop
    End If
 End With
End Sub

Private Sub GridAddNew()
Dim i As Long
With msFlex
    For i = 0 To mRsSearch.Fields.Count - 1
            Select Case mRsSearch.Fields(i).Type
            Case adInteger, 2, 3 '-- Integer
                .TextMatrix(.Row, i) = 0
            Case adDecimal, 131, adNumeric, adDouble '-- Double
                .TextMatrix(.Row, i) = Format(0, GProcNumberFormat(mRsSearch.Fields(i).Precision, mRsSearch.Fields(i).NumericScale))
            Case adDate '-- Date
                 .TextMatrix(.Row, i) = ""
            Case Else
             '.ColAlignment(I) = flexAlignLeftCenter
             .TextMatrix(.Row, i) = ""
        End Select
    Next i
End With
End Sub
