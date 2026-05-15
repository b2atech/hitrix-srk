VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmSearchMultiSelect 
   BackColor       =   &H00C0E0FF&
   BorderStyle     =   0  'None
   Caption         =   "Search"
   ClientHeight    =   5985
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   8130
   ControlBox      =   0   'False
   Icon            =   "frmSearchMultiSelect.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5985
   ScaleWidth      =   8130
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox txtBal 
      Alignment       =   1  'Right Justify
      BackColor       =   &H80000018&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   375
      Index           =   0
      Left            =   1005
      Locked          =   -1  'True
      TabIndex        =   9
      Top             =   4620
      Width           =   1035
   End
   Begin VB.TextBox txtTot 
      Alignment       =   1  'Right Justify
      BackColor       =   &H80000018&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   375
      Index           =   0
      Left            =   1005
      Locked          =   -1  'True
      TabIndex        =   7
      Top             =   4245
      Width           =   1035
   End
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
      Height          =   465
      Left            =   5040
      TabIndex        =   3
      Top             =   5235
      Width           =   1260
   End
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
      Height          =   465
      Left            =   3660
      TabIndex        =   2
      Top             =   5235
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
      Height          =   465
      Left            =   6450
      TabIndex        =   4
      Top             =   5235
      Width           =   1215
   End
   Begin MSComctlLib.ListView ListView 
      Height          =   3795
      Left            =   15
      TabIndex        =   1
      Top             =   405
      Width           =   7860
      _ExtentX        =   13864
      _ExtentY        =   6694
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
   Begin VB.Label LblBalOfValue 
      BackStyle       =   0  'Transparent
      Caption         =   "Balance Of :"
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
      Height          =   300
      Left            =   90
      TabIndex        =   12
      Top             =   5010
      Width           =   3060
   End
   Begin VB.Label LblBalance 
      BackStyle       =   0  'Transparent
      Caption         =   "Balance :"
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
      Height          =   495
      Left            =   90
      TabIndex        =   11
      Top             =   4620
      Width           =   1365
   End
   Begin VB.Label LblTotal 
      BackStyle       =   0  'Transparent
      Caption         =   "Total :"
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
      Height          =   495
      Left            =   90
      TabIndex        =   10
      Top             =   4245
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "Label1"
      Height          =   495
      Left            =   3465
      TabIndex        =   8
      Top             =   2745
      Width           =   1215
   End
   Begin VB.Shape Shape1 
      FillColor       =   &H00BDDFD1&
      FillStyle       =   0  'Solid
      Height          =   765
      Left            =   3420
      Shape           =   4  'Rounded Rectangle
      Top             =   5085
      Width           =   4455
   End
   Begin VB.Shape Shape4 
      BorderColor     =   &H00C00000&
      BorderWidth     =   5
      Height          =   5940
      Left            =   0
      Top             =   30
      Width           =   8130
   End
   Begin VB.Label lbl 
      BackColor       =   &H00C0E0FF&
      BackStyle       =   0  'Transparent
      Height          =   375
      Index           =   0
      Left            =   4335
      TabIndex        =   6
      Top             =   0
      Width           =   1035
   End
   Begin VB.Label lblSearch 
      BackStyle       =   0  'Transparent
      Caption         =   "Search...."
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   510
      Left            =   165
      TabIndex        =   5
      Top             =   5355
      Width           =   3630
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H00000080&
      FillColor       =   &H00EDFDFE&
      FillStyle       =   0  'Solid
      Height          =   5850
      Left            =   0
      Top             =   75
      Width           =   8040
   End
End
Attribute VB_Name = "frmSearchMultiSelect"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'---------- Search Multi Form Variables
Dim mRsSearch As Recordset
Dim mCaption '-- Array Col Caption
Dim mColWidth '-- Array Col Width
Dim mTName As String '-- Table Name
Dim mFieldname As String '-- Array Field Name
Dim mQuery As String
Dim mCn As Connection
Dim mRetCol() As String '-- Array Selected Row Values
Dim mUpperHeading As Integer, mLowerHeading As Integer ''Upper & Lower bound of columns
Dim mIsTNameAsQry As Boolean '-- Yes If mTName is Query
Dim mOrderCol As String '-- Order by Col String
Dim mCriteria As String '-- Criteria String
Dim mIsOnlySearchMatch As Boolean
Dim mUnloadFlag As Boolean
Dim mReportFlag As Boolean
Dim mTotOfSelCols '-- Array Col no which Selected col Total u want
Dim mLBoundTotOfSelCols As Long '-''Upper & Lower bound of mTotOfSelCols
Dim mUBoundTotOfSelCols As Long
'---------- Search Multi Field
Public Property Get RetCol() As Variant
    RetCol = mRetCol
End Property
'====== Name : SearchMultiField
'--- Variables : pTName (Table Name),pFieldname ( Array Fieldname),pCaptions ( Array colume Caption),pColWidths (Array Col width)
'-- pCriteria (condition), pMatchChar (Match Character),pOrderCol (order by cols),pfrmLeft ( Form Left Pos),pfrmTop  ( Form Top Pos)
'-- pIsAllowAdd (Allow to write New Value),pCon (connection Name),pIsTNameAsQry (Yes If 1 st parameter is Query)
'-- pIsReportFlag -  (if Yes format form show As Selection Report) ,pRptCaption - (Set Report Form Caption)
'--- pTotOfSelCols (Array of cols -show total of Selected, (Balance Value))
Public Sub SearchMultiField(pTName As String, pFieldname As String, pCaptions, Optional pColWidths, Optional pCriteria As String, Optional pMatchChar As String, Optional pOrderCol As String, Optional pfrmLeft As Long, Optional pfrmTop As Long, Optional pCon As Connection, Optional pIsTNameAsQry As Boolean, Optional pIsOnlyMatch As Boolean, Optional pIsReportFlag As Boolean, Optional pRptCaption As String, Optional pTotOfSelCols)
Dim cmdTop As Long
    mTName = pTName
    mFieldname = pFieldname
    mOrderCol = pOrderCol
    mCriteria = pCriteria
    mIsTNameAsQry = pIsTNameAsQry
    mCaption = pCaptions
    mUpperHeading = UBound(pCaptions)
    mLowerHeading = LBound(pCaptions)
    mIsOnlySearchMatch = pIsOnlyMatch
    mReportFlag = pIsReportFlag
    mTotOfSelCols = pTotOfSelCols
    If IsMissing(mTotOfSelCols) = False Then
        mLBoundTotOfSelCols = LBound(mTotOfSelCols)
        mUBoundTotOfSelCols = UBound(mTotOfSelCols)
    End If
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
    '-- Width
    If Not IsMissing(pColWidths) Then mColWidth = pColWidths
    mUnloadFlag = False
    With frmSearchMultiSelect
        '--- (Report Flag True) Give Report Format
        If mReportFlag = True Then
            '-- Set Height,Top,Width,Left
            .Caption = pRptCaption
           .Top = 50
           .ListView.Top = 70
           .ListView.Width = 15000
           .Width = .ListView.Width + 500
           .ListView.Height = 6600
           .Height = .ListView.Height + 1000
           cmdTop = ListView.Top + ListView.Height + 80
           cmdOk.Top = cmdTop
           cmdExit.Top = cmdTop
           cmdAll.Top = cmdTop
           .Left = Screen.Width / 2 - Me.Width / 2
           lblSearch.Visible = False
           txt(0).Visible = False
        End If
        .txt(0) = pMatchChar
        .txt(0).SelStart = Len(pMatchChar)
        If pfrmLeft = 0 Then
            .Top = Screen.Height / 2 - .Height / 2
            .Left = Screen.Width / 2 - .Width / 2
        Else
            .Left = pfrmLeft
            .Top = pfrmTop
        End If
        If mUnloadFlag = False Then
            .Show vbModal
        End If
     End With
End Sub
Private Sub CmdAll_Click()
Dim i As Long
For i = 1 To ListView.ListItems.Count
    ListView.ListItems(i).Checked = True
Next
cmdOK_Click
End Sub

Private Sub cmdExit_Click()
mUnloadFlag = True
Unload Me
End Sub
Private Sub cmdOK_Click()
Dim i As Long, j As Long
Dim RowIndex As Integer
Dim itmX As ListItem
For i = 1 To ListView.ListItems.Count
     If ListView.ListItems(i).Checked Then
        RowIndex = RowIndex + 1
     End If
Next
If RowIndex = 0 Then Exit Sub
ReDim mRetCol(RowIndex - 1, mUpperHeading)
RowIndex = 0
For i = 1 To ListView.ListItems.Count
    If ListView.ListItems(i).Checked Then
        Set itmX = ListView.ListItems.Item(i)
        mRetCol(RowIndex, 0) = itmX
        For j = 1 To mUpperHeading
            mRetCol(RowIndex, j) = itmX.SubItems(j)
        Next j
        RowIndex = RowIndex + 1
    End If
Next
mUnloadFlag = True
Unload Me
End Sub
Private Sub Form_KeyPress(KeyAscii As Integer)
Select Case KeyAscii
    Case vbKeyReturn '-- Enter
        cmdOK_Click
    Case 27 '-- Esc
        cmdExit_Click
End Select
End Sub
Private Sub PrepareListView()
Dim X As String
Dim i As Long
Dim objItem As ListItem
    '--- Add Records
    Do While Not mRsSearch.EOF
        Set objItem = ListView.ListItems.Add(, , mRsSearch.Fields(0) & vbNullString)
        For i = 1 To mRsSearch.Fields.Count - 1
           Select Case mRsSearch.Fields(i).Type 'right(space(20)+str(vno),20)
                Case adVarChar, adWChar, adChar, adVarWChar '-- Text
                    objItem.SubItems(i) = mRsSearch.Fields(i) & vbNullString
                Case adDate, 7, 135 '--Date
                    objItem.SubItems(i) = mRsSearch.Fields(i) & vbNullString
                Case Else
                    objItem.SubItems(i) = Right(Space(12) + Str(mRsSearch.Fields(i)), 12) & vbNullString
           End Select
        Next i
        mRsSearch.MoveNext
    Loop
End Sub

Private Sub Form_Load()
Dim X As String
Static MaxIdx               '-----Largest index in array.
ReDim mRetCol(0, 0)
mRetCol(0, 0) = ""
MaxIdx = 1
X = mQuery & IIf(Trim(mOrderCol) = "", "", " order by " & mOrderCol)
GProcRstOpen mRsSearch, X, "R", mCn
Loop1:
    If MaxIdx > mRsSearch.Fields.Count - 1 Then
        SetSearchCtrl
        SetTotBalCtrl
        If mIsOnlySearchMatch = False Then PrepareListView
        Exit Sub
    Else
        If mReportFlag = False Then
            Load txt(MaxIdx)    '-----Create new item in array.
            txt(MaxIdx).Top = txt(MaxIdx - 1).Top
            txt(MaxIdx).Visible = True
        End If
        MaxIdx = MaxIdx + 1
    End If
GoTo Loop1
End Sub

Private Sub ListView_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
ListView.Sorted = True
ListView.SortKey = ColumnHeader.Index - 1
'ListView.SelectedItem.ForeColor = vbWhite
End Sub

Private Sub ListView_ItemCheck(ByVal Item As MSComctlLib.ListItem)
Item.Ghosted = True
ShowTotBalOfSelCol
End Sub
'-- Show Total & Balance Of Selected Columns
Private Sub ShowTotBalOfSelCol()
Dim i As Long
Dim j As Long
Dim ColNo As Long
Dim Value As Double
Dim ShowBal As Boolean
Dim itmX As ListItem
If IsMissing(mTotOfSelCols) = False Then
   For i = mLBoundTotOfSelCols To mUBoundTotOfSelCols
        ColNo = mTotOfSelCols(i, 0)
        txtTot(ColNo) = "0.00"
        If mTotOfSelCols(i, 1) <> 0 Then txtBal(ColNo) = mTotOfSelCols(i, 1)
   Next i
    For j = mLBoundTotOfSelCols To mUBoundTotOfSelCols
        ColNo = mTotOfSelCols(j, 0)
        If mTotOfSelCols(j, 1) > 0 Then ShowBal = True
        For i = 1 To ListView.ListItems.Count
            If ListView.ListItems(i).Checked Then
                Set itmX = ListView.ListItems.Item(i)
                If ColNo = 0 Then
                    Value = CDbl(ListView.ListItems.Item(ColNo))
                Else
                    Value = CDbl(itmX.SubItems(ColNo))
                End If
                txtTot(ColNo) = CDbl(txtTot(ColNo)) + CDbl(Value)
            End If
        Next i
        If ShowBal = True Then
           ' If CDbl(txtTot(ColNo)) > txtBal(ColNo) Then
           '     txtBal(ColNo) = 0
           ' Else
                txtBal(ColNo) = CDbl(txtBal(ColNo)) - CDbl(txtTot(ColNo))
           ' End If
        End If
    Next j
End If
End Sub
Private Sub txt_Change(Index As Integer)
Dim QStr As String
Dim i As Long
Dim QueryStr As String
Dim tblName As String
Dim strLen As Long
Dim ColString As String
Dim OrderByStr  As String
ColString = ""
OrderByStr = Trim(mOrderCol)
For i = 0 To mRsSearch.Fields.Count - 1
    If txt(i) <> "" Then
        Select Case mRsSearch.Fields(i).Type
            Case adDouble, adInteger, 2, 3, adNumeric, adDecimal 'dbLong, dbDouble
                If IsNumeric(txt(i)) = True Then
                    If QStr = "" Then
                        QStr = mRsSearch.Fields(i).Name & " = " & txt(i)
                    Else
                        QStr = QStr & " AND " & mRsSearch.Fields(i).Name & "= " & txt(i)
                    End If
                End If
            Case Else
                If QStr = "" Then
                    If gBackEndDB = gBackEndAccess Then
                       QStr = "left(" & mRsSearch.Fields(i).Name & "," & Len(txt(i)) & ")='" & txt(i) & "'"
                    Else '-- Oracle
                       QStr = "LPad(Upper(" & mRsSearch.Fields(i).Name & ")," & Len(txt(i)) & ")=Upper('" & txt(i) & "')"
                    End If
                Else
                    If gBackEndDB = gBackEndAccess Then
                       QStr = QStr & " AND left(" & mRsSearch.Fields(i).Name & "," & Len(txt(i)) & ")='" & txt(i) & "'"
                    Else '-- Oracle
                       QStr = QStr & " AND LPad(Upper(" & mRsSearch.Fields(i).Name & ")," & Len(txt(i)) & ")=Upper('" & txt(i) & "')"
                    End If
                End If
           End Select
        If OrderByStr = "" Then '--Order by
            OrderByStr = mRsSearch.Fields(i).Name
        Else
            OrderByStr = OrderByStr & " ," & mRsSearch.Fields(i).Name
        End If
    End If
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
        ColString = mFieldname
        tblName = mTName
        strLen = InStr(LCase(mQuery), "where")
        If strLen = 0 Then
            QueryStr = "select " & ColString & " from " & tblName & " where " & QStr & IIf(Trim(OrderByStr) = "", "", " order by " & OrderByStr)
        Else
            QueryStr = "select " & ColString & " from " & tblName & IIf(Trim(mCriteria) = "", "", " where " & mCriteria) & " and  " & QStr & IIf(Trim(OrderByStr) = "", "", " order by " & OrderByStr)
        End If
   End If
End If
Set mRsSearch = New Recordset
GProcRstOpen mRsSearch, QueryStr, "R", mCn
SetSearchCtrl
If QueryStr <> "" Then
    PrepareListView
End If
End Sub
'--- Set Tot & Bal Col Controls
Private Sub SetTotBalCtrl()
Dim i As Long
Dim ColNo As Long
txtTot(0).Visible = False
txtBal(0).Visible = False
LblBalOfValue = ""
If IsMissing(mTotOfSelCols) = True Then
    lblTotal.Visible = False
    lblBalance.Visible = False
Else
    lblTotal.Visible = True
      '-- Load txtTot,txtBal Text Boxes
      For i = mLBoundTotOfSelCols To mUBoundTotOfSelCols
        ColNo = mTotOfSelCols(i, 0)
        If ColNo <> 0 Then Load txtTot(ColNo)
        txtTot(ColNo).Visible = True
        If mTotOfSelCols(i, 1) <> 0 Then
            LblBalOfValue = "Balance Of :  " & mTotOfSelCols(i, 1)
            lblBalance.Visible = True
            If ColNo <> 0 Then Load txtBal(ColNo)
            txtBal(ColNo).Visible = True
            txtBal(ColNo).Left = txt(ColNo).Left
            txtBal(ColNo).Width = txt(ColNo).Width
        End If
        txtTot(ColNo).Left = txt(ColNo).Left
        txtTot(ColNo).Width = txt(ColNo).Width
      Next i
End If
End Sub
Private Sub SetSearchCtrl()
Dim i As Long
With ListView
    .ColumnHeaders.Clear
    .ListItems.Clear
    .Sorted = False
    For i = 0 To mRsSearch.Fields.Count - 1
    '--- Clear List Ctrl
        '--- Add Column Headers
        Select Case mRsSearch.Fields(i).Type
            Case adVarChar, adWChar, adChar, adVarWChar '-- Text
                .ColumnHeaders.Add , , mRsSearch.Fields(i).Name, , lvwColumnLeft
            Case adDate '-- Date
                If i = 0 Then
                    .ColumnHeaders.Add , , mRsSearch.Fields(i).Name, , lvwColumnLeft
                Else
                    .ColumnHeaders.Add , , mRsSearch.Fields(i).Name, , lvwColumnCenter
                End If
            Case Else '--Number
                If i = 0 Then
                    .ColumnHeaders.Add , , mRsSearch.Fields(i).Name, , lvwColumnLeft
                Else
                    .ColumnHeaders.Add , , mRsSearch.Fields(i).Name, , lvwColumnRight
                End If
        End Select
       ' .ColumnHeaders.Add , , mRsSearch.Fields(I).Name, , lvwColumnLeft
        If mReportFlag = False Then
            txt(i).Width = mColWidth(i)
            If mColWidth(i) = 0 Then txt(i).Visible = False
            If i > 0 Then txt(i).Left = txt(i - 1).Left + txt(i - 1).Width + 20
        End If
    Next i
    '--- Change Column Width
    For i = 0 To .ColumnHeaders.Count - 1
        .ColumnHeaders(i + 1).text = mCaption(i)
        .ColumnHeaders(i + 1).Width = mColWidth(i)
    Next i
End With
End Sub
Private Sub txt_GotFocus(Index As Integer)
If TypeOf Me.ActiveControl Is TextBox Then GProcSelectBox Me.ActiveControl
End Sub

Private Sub txt_KeyPress(Index As Integer, KeyAscii As Integer)
Select Case mRsSearch.Fields(Index).Type
    Case adDouble, adInteger, 2, 3, adNumeric, adDecimal
        KeyAscii = GProcNumberOnly(KeyAscii)
End Select
End Sub

Private Sub txt_LostFocus(Index As Integer)
txt(Index).text = GProcProperCase(txt(Index).text)
End Sub

'Select Case mRsSearch.Fields(i).Type 'right(space(20)+str(vno),20)
'   Case adVarChar, adWChar, adChar, adVarWChar '-- Text
'       If i = 0 Then
'           Set ObjItem = .ListItems.Add(, , mRsSearch.Fields(i))
'       Else
'           ObjItem.SubItems(i) = mRsSearch.Fields(i) & vbNullString ' CStr(IIf(IsNull(mRsSearch.Fields(i)), "", mRsSearch.Fields(i)))
'       End If
'   Case adDate
'       If i = 0 Then
'
'       Else
'           ObjItem.SubItems(i) = CStr(IIf(IsNull(mRsSearch.Fields(i)), "", mRsSearch.Fields(i)))
'       End If
'   Case Else
'       If i = 0 Then
'           Set ObjItem = .ListItems.Add(, , mRsSearch.Fields(i))
'       Else
'           ObjItem.SubItems(i) = IIf(IsNull(mRsSearch.Fields(i)), "", Right(Space(7) + Str(mRsSearch.Fields(i)), 7)) 'Right(Space(7) + Str(R_Rs.Fields(i)), Len(R_Rs.Fields(i))))
'       End If
'End Select
'ObjItem.SubItems(i) = CStr(mRsSearch.Fields(i)) & vbNullString
