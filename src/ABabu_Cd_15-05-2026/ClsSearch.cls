VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClsSearch"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'--- Search Multi Field
Public Sub SearchMultiField(tblName As String, ColName As String, Headings, Optional Widths, Optional CondStr As String, Optional MatchChar As String, Optional OrderCol As String, Optional frmLeft As Long, Optional frmTop As Long, Optional IsAllowAdd As Boolean, Optional Cn As Connection, Optional IstblNameAsQry As Boolean, Optional IsShowOnlyMatch As Boolean, Optional GroupCol As String)
    frmSearchMulti.SearchMultiField tblName, ColName, Headings, Widths, CondStr, MatchChar, OrderCol, frmLeft, frmTop, IsAllowAdd, Cn, IstblNameAsQry, IsShowOnlyMatch, GroupCol
End Sub
Public Property Get SearchMultiRetCol() As Variant
    SearchMultiRetCol = frmSearchMulti.RetCol
End Property
'--- Search Multi Select Field
Public Sub SearchMultiSelField(tblName As String, ColName As String, Headings, Optional Widths, Optional CondStr As String, Optional MatchChar As String, Optional OrderCol As String, Optional frmLeft As Long, Optional frmTop As Long, Optional Cn As Connection, Optional IstblNameAsQry As Boolean, Optional IsShowOnlyMatch As Boolean, Optional IsReportFlag As Boolean, Optional ReportCaption As String, Optional TotOfSelCols)
    frmSearchMultiSelect.SearchMultiField tblName, ColName, Headings, Widths, CondStr, MatchChar, OrderCol, frmLeft, frmTop, Cn, IstblNameAsQry, IsShowOnlyMatch
'    frmSearchMultiSelect.SearchMultiField tblName, ColName, Headings, Widths, CondStr, MatchChar, OrderCol, frmLeft, frmTop, Cn, IstblNameAsQry, IsShowOnlyMatch, IsReportFlag, ReportCaption, TotOfSelCols
End Sub
Public Property Get SearchMultiSelRetCol() As Variant
    SearchMultiSelRetCol = frmSearchMultiSelect.RetCol
End Property
'--- Search Multi Field AccMaster
Public Sub SearchMultiFieldAccMaster(Optional MatchChar As String, Optional frmLeft As Long, Optional frmTop As Long)
    frmSearchMultiAccMaster.SearchMultiFieldAccMaster MatchChar, frmLeft, frmTop
End Sub
Public Property Get SearchMultiRetColAccMaster() As Variant
    SearchMultiRetColAccMaster = frmSearchMultiAccMaster.RetCol
End Property

