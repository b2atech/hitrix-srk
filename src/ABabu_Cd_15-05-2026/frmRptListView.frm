VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form frmRptListView 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "List View"
   ClientHeight    =   6795
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   11505
   ControlBox      =   0   'False
   Icon            =   "frmRptListView.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6795
   ScaleWidth      =   11505
   Begin VB.CommandButton cmdAll 
      Caption         =   "&Select All"
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
      Left            =   10125
      TabIndex        =   2
      Top             =   75
      Width           =   1260
   End
   Begin VB.CommandButton cmdOk 
      Caption         =   "&OK"
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
      Height          =   450
      Left            =   7455
      TabIndex        =   1
      Top             =   75
      Width           =   1215
   End
   Begin VB.CommandButton cmdExit 
      Cancel          =   -1  'True
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
      Left            =   8790
      TabIndex        =   3
      Top             =   75
      Width           =   1215
   End
   Begin MSComctlLib.ListView ListView 
      Height          =   6135
      Left            =   30
      TabIndex        =   0
      Top             =   615
      Width           =   11400
      _ExtentX        =   20108
      _ExtentY        =   10821
      View            =   3
      MultiSelect     =   -1  'True
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      Checkboxes      =   -1  'True
      FullRowSelect   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   0
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
         Weight          =   400
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
Attribute VB_Name = "frmRptListView"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Option Explicit
''---------- Search Multi Form Variables
'Dim mRsSearch As Recordset
'Dim mCaption
'Dim mColWidth
'Dim mQuery As String
'Dim mCn As Connection
'Dim mRetCol() As String
'Dim mUpperHeading As Integer, mLowerHeading As Integer ''Upper & Lower bound of columns
'Dim mIsTNameAsQry As Boolean
'Dim mOrderCol As String
'Dim mCriteria As String
'Dim mIsOnlySearchMatch As Boolean
'Dim mTName As String
'Dim mFieldname As String
'Dim mUnloadFlag As Boolean
''---------- Search Multi Field
'Public Property Get RetCol() As Variant
'    RetCol = mRetCol
'End Property
'Public Sub SearchMultiField(pTName As String, pFieldname As String, pCaptions, Optional pColWidths, Optional pCriteria As String, Optional pMatchChar As String, Optional pOrderCol As String, Optional pfrmLeft As Long, Optional pfrmTop As Long, Optional pCon As Connection, Optional pIsTNameAsQry As Boolean, Optional pIsOnlyMatch As Boolean, Optional pListCtrlWidth As Long)
'    mTName = pTName
'    mFieldname = pFieldname
'    mOrderCol = pOrderCol
'    mCriteria = pCriteria
'    mIsTNameAsQry = pIsTNameAsQry
'    mCaption = pCaptions
'    mUpperHeading = UBound(pCaptions)
'    mLowerHeading = LBound(pCaptions)
'    mIsOnlySearchMatch = pIsOnlyMatch
'    '-- Set Query
'    If mIsTNameAsQry = True Then
'        mQuery = pTName & IIf(Trim(mCriteria) = "", "", " where " & mCriteria) '& IIf(Trim(mOrderCol) = "", "", " order by " & mOrderCol)
'    Else
'        mQuery = "select " & pFieldname & " from " & pTName & " " & IIf(Trim(mCriteria) = "", "", " where " & pCriteria) '& IIf(Trim(mOrderCol) = "", "", " order by " & mOrderCol)
'    End If
'    '-- Set connection
'    If pCon Is Nothing Then
'        Set mCn = gCn
'    Else
'        Set mCn = pCon
'    End If
'    '-- Width
'    If Not IsMissing(pColWidths) Then mColWidth = pColWidths
'    mUnloadFlag = False
'    frmSearchMultiSelect.txt(0) = pMatchChar
'    frmSearchMultiSelect.txt(0).SelStart = Len(pMatchChar)
'    If pListCtrlWidth > 0 Then
'        frmSearchMultiSelect.ListView.Width = pListCtrlWidth
'        frmSearchMultiSelect.Width = frmSearchMultiSelect.ListView.Width + 500
'    End If
'    If pfrmLeft = 0 Then
'        frmSearchMultiSelect.Top = Screen.Height / 2 - frmSearchMultiSelect.Height / 2
'        frmSearchMultiSelect.Left = Screen.Width / 2 - frmSearchMultiSelect.Width / 2
'    Else
'        frmSearchMultiSelect.Left = pfrmLeft
'        frmSearchMultiSelect.Top = pfrmTop
'    End If
'    If mUnloadFlag = False Then
'        frmSearchMultiSelect.Show vbModal
'    End If
'End Sub
'Private Sub CmdAll_Click()
'Dim i As Long
'For i = 1 To ListView.ListItems.Count
'    ListView.ListItems(i).Checked = True
'Next
'cmdOK_Click
'End Sub
'
'Private Sub cmdExit_Click()
'mUnloadFlag = True
'Unload Me
'End Sub
'Private Sub cmdOK_Click()
'Dim i As Long, j As Long
'Dim RowIndex As Integer
'Dim itmX As ListItem
'For i = 1 To ListView.ListItems.Count
'     If ListView.ListItems(i).Checked Then
'        RowIndex = RowIndex + 1
'     End If
'Next
'If RowIndex = 0 Then Exit Sub
'ReDim mRetCol(RowIndex - 1, mUpperHeading)
'RowIndex = 0
'For i = 1 To ListView.ListItems.Count
'    If ListView.ListItems(i).Checked Then
'        Set itmX = ListView.ListItems.Item(i)
'        mRetCol(RowIndex, 0) = itmX
'        For j = 1 To mUpperHeading
'            mRetCol(RowIndex, j) = itmX.SubItems(j)
'        Next j
'        RowIndex = RowIndex + 1
'    End If
'Next
'mUnloadFlag = True
'Unload Me
'End Sub
'Private Sub Form_KeyPress(KeyAscii As Integer)
'Select Case KeyAscii
'    Case vbKeyReturn '-- Enter
'        cmdOK_Click
'    Case 27 '-- Esc
'        cmdExit_Click
'End Select
'End Sub
'Private Sub PrepareListView()
'Dim X As String
'Dim i As Long
'Dim objItem As ListItem
'    '--- Add Records
'    Do While Not mRsSearch.EOF
'        Set objItem = ListView.ListItems.Add(, , mRsSearch.Fields(0) & vbNullString)
'        For i = 1 To mRsSearch.Fields.Count - 1
'           Select Case mRsSearch.Fields(i).Type 'right(space(20)+str(vno),20)
'                Case adVarChar, adWChar, adChar, adVarWChar '-- Text
'                    objItem.SubItems(i) = mRsSearch.Fields(i) & vbNullString
'                Case adDate, 7, 135 '--Date
'                    objItem.SubItems(i) = mRsSearch.Fields(i) & vbNullString
'                Case Else
'                    objItem.SubItems(i) = Right(Space(7) + Str(mRsSearch.Fields(i)), 7) & vbNullString
'           End Select
'        Next i
'        mRsSearch.MoveNext
'    Loop
'End Sub
'
'Private Sub Form_Load()
'Dim X As String
'Static MaxIdx               '-----Largest index in array.
'ReDim mRetCol(0, 0)
'mRetCol(0, 0) = ""
'MaxIdx = 1
'X = mQuery & IIf(Trim(mOrderCol) = "", "", " order by " & mOrderCol)
'GProcRstOpen mRsSearch, X, "R", mCn
'Loop1:
'    If MaxIdx > mRsSearch.Fields.Count - 1 Then
'        SetSearchCtrl
'         If mIsOnlySearchMatch = False Then PrepareListView
'        Exit Sub
'    Else
'        Load txt(MaxIdx)    '-----Create new item in array.
'        txt(MaxIdx).Top = txt(MaxIdx - 1).Top
'        txt(MaxIdx).Visible = True
'        MaxIdx = MaxIdx + 1
'    End If
'GoTo Loop1
'End Sub
'
'Private Sub ListView_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
'ListView.Sorted = True
'ListView.SortKey = ColumnHeader.Index - 1
''ListView.SelectedItem.ForeColor = vbWhite
'End Sub
'
'Private Sub ListView_ItemCheck(ByVal Item As MSComctlLib.ListItem)
'Item.Ghosted = True
''Dim i As Long
''Dim LstItem As ListItem
''If Item.Checked = True Then
''    Item.ForeColor = vbBlue
''    Do While i <= mUpperHeading
''        Set LstItem = ListView.ListItems(Item).SubItems(i)                  'Item.SubItems(1)
''        LstItem.ForeColor = vbBlue
''    Loop
''Else
''    Item.ForeColor = vbBlack
''    'Item.Bold = False
''End If
'End Sub
'
'Private Sub txt_Change(Index As Integer)
'Dim QStr As String
'Dim i As Long
'Dim QueryStr As String
'Dim tblName As String
'Dim strLen As Long
'Dim ColString As String
'Dim OrderByStr  As String
'ColString = ""
'OrderByStr = Trim(mOrderCol)
'For i = 0 To mRsSearch.Fields.Count - 1
'    If txt(i) <> "" Then
'        Select Case mRsSearch.Fields(i).Type
'            Case adDouble, adInteger, 2, 3, adNumeric, adDecimal 'dbLong, dbDouble
'                If IsNumeric(txt(i)) = True Then
'                    If QStr = "" Then
'                        QStr = mRsSearch.Fields(i).Name & " = " & txt(i)
'                    Else
'                        QStr = QStr & " AND " & mRsSearch.Fields(i).Name & "= " & txt(i)
'                    End If
'                End If
'            Case Else
'                If QStr = "" Then
'                    If gBackEndDB = gBackEndAccess Then
'                       QStr = "left(" & mRsSearch.Fields(i).Name & "," & Len(txt(i)) & ")='" & txt(i) & "'"
'                    Else '-- Oracle
'                       QStr = "LPad(Upper(" & mRsSearch.Fields(i).Name & ")," & Len(txt(i)) & ")=Upper('" & txt(i) & "')"
'                    End If
'                Else
'                    If gBackEndDB = gBackEndAccess Then
'                       QStr = QStr & " AND left(" & mRsSearch.Fields(i).Name & "," & Len(txt(i)) & ")='" & txt(i) & "'"
'                    Else '-- Oracle
'                       QStr = QStr & " AND LPad(Upper(" & mRsSearch.Fields(i).Name & ")," & Len(txt(i)) & ")=Upper('" & txt(i) & "')"
'                    End If
'                End If
'           End Select
'        If OrderByStr = "" Then '--Order by
'            OrderByStr = mRsSearch.Fields(i).Name
'        Else
'            OrderByStr = OrderByStr & " ," & mRsSearch.Fields(i).Name
'        End If
'    End If
'Next i
'If QStr = "" Then
'    If mIsOnlySearchMatch = False Then
'        QueryStr = mQuery & IIf(Trim(OrderByStr) = "", "", " order by " & OrderByStr)
'    Else
'        QueryStr = ""
'    End If
'Else
'    If mIsTNameAsQry = True Then '--Query
'        strLen = InStr(LCase(mQuery), "where")
'        If strLen = 0 Then
'            QueryStr = mQuery & " where " & QStr & IIf(Trim(OrderByStr) = "", "", " order by " & OrderByStr)
'        Else
'            QueryStr = mQuery & " and  " & QStr & IIf(Trim(OrderByStr) = "", "", " order by " & OrderByStr)
'        End If
'   Else '--Tables
'        ColString = mFieldname
'        tblName = mTName
'        strLen = InStr(LCase(mQuery), "where")
'        If strLen = 0 Then
'            QueryStr = "select " & ColString & " from " & tblName & " where " & QStr & IIf(Trim(OrderByStr) = "", "", " order by " & OrderByStr)
'        Else
'            QueryStr = "select " & ColString & " from " & tblName & IIf(Trim(mCriteria) = "", "", " where " & mCriteria) & " and  " & QStr & IIf(Trim(OrderByStr) = "", "", " order by " & OrderByStr)
'        End If
'   End If
'End If
'Set mRsSearch = New Recordset
'GProcRstOpen mRsSearch, QueryStr, "R", mCn
'SetSearchCtrl
'If QueryStr <> "" Then
'    PrepareListView
'End If
'End Sub
''------
'Private Sub SetSearchCtrl()
'Dim i As Long
'With ListView
'    .ColumnHeaders.Clear
'    .ListItems.Clear
'    .Sorted = False
'    For i = 0 To mRsSearch.Fields.Count - 1
'    '--- Clear List Ctrl
'        '--- Add Column Headers
'        Select Case mRsSearch.Fields(i).Type
'            Case adVarChar, adWChar, adChar, adVarWChar '-- Text
'                .ColumnHeaders.Add , , mRsSearch.Fields(i).Name, , lvwColumnLeft
'            Case adDate '-- Date
'                If i = 0 Then
'                    .ColumnHeaders.Add , , mRsSearch.Fields(i).Name, , lvwColumnLeft
'                Else
'                    .ColumnHeaders.Add , , mRsSearch.Fields(i).Name, , lvwColumnCenter
'                End If
'            Case Else '--Number
'                If i = 0 Then
'                    .ColumnHeaders.Add , , mRsSearch.Fields(i).Name, , lvwColumnLeft
'                Else
'                    .ColumnHeaders.Add , , mRsSearch.Fields(i).Name, , lvwColumnRight
'                End If
'        End Select
'       ' .ColumnHeaders.Add , , mRsSearch.Fields(I).Name, , lvwColumnLeft
'        txt(i).Width = mColWidth(i)
'        If mColWidth(i) = 0 Then txt(i).Visible = False
'        If i > 0 Then txt(i).Left = txt(i - 1).Left + txt(i - 1).Width + 20
'    Next i
'    '--- Change Column Width
'    For i = 0 To .ColumnHeaders.Count - 1
'        .ColumnHeaders(i + 1).Text = mCaption(i)
'        .ColumnHeaders(i + 1).Width = mColWidth(i)
'    Next i
'End With
'End Sub
'Private Sub txt_GotFocus(Index As Integer)
'If TypeOf Me.ActiveControl Is TextBox Then GProcSelectBox Me.ActiveControl
'End Sub
'
'Private Sub txt_KeyPress(Index As Integer, KeyAscii As Integer)
'Select Case mRsSearch.Fields(Index).Type
'    Case adDouble, adInteger, 2, 3, adNumeric, adDecimal
'        KeyAscii = GProcNumberOnly(KeyAscii)
'End Select
'End Sub
'
'Private Sub txt_LostFocus(Index As Integer)
'txt(Index).Text = GProcProperCase(txt(Index).Text)
'End Sub
'
''Select Case mRsSearch.Fields(i).Type 'right(space(20)+str(vno),20)
''   Case adVarChar, adWChar, adChar, adVarWChar '-- Text
''       If i = 0 Then
''           Set ObjItem = .ListItems.Add(, , mRsSearch.Fields(i))
''       Else
''           ObjItem.SubItems(i) = mRsSearch.Fields(i) & vbNullString ' CStr(IIf(IsNull(mRsSearch.Fields(i)), "", mRsSearch.Fields(i)))
''       End If
''   Case adDate
''       If i = 0 Then
''
''       Else
''           ObjItem.SubItems(i) = CStr(IIf(IsNull(mRsSearch.Fields(i)), "", mRsSearch.Fields(i)))
''       End If
''   Case Else
''       If i = 0 Then
''           Set ObjItem = .ListItems.Add(, , mRsSearch.Fields(i))
''       Else
''           ObjItem.SubItems(i) = IIf(IsNull(mRsSearch.Fields(i)), "", Right(Space(7) + Str(mRsSearch.Fields(i)), 7)) 'Right(Space(7) + Str(R_Rs.Fields(i)), Len(R_Rs.Fields(i))))
''       End If
''End Select
''ObjItem.SubItems(i) = CStr(mRsSearch.Fields(i)) & vbNullString
