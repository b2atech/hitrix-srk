VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmUserPrevilage 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "User Previlage Entry"
   ClientHeight    =   7350
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11460
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7350
   ScaleWidth      =   11460
   Begin VB.CommandButton cmdBtn 
      Caption         =   "Optional"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   400
      Index           =   12
      Left            =   9960
      Style           =   1  'Graphical
      TabIndex        =   24
      ToolTipText     =   "To Add New Record"
      Top             =   6480
      Width           =   1095
   End
   Begin VB.CommandButton cmdBtn 
      Caption         =   "&Label"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   400
      Index           =   11
      Left            =   9960
      Style           =   1  'Graphical
      TabIndex        =   23
      ToolTipText     =   "To Add New Record"
      Top             =   6000
      Width           =   1095
   End
   Begin VB.CommandButton cmdBtn 
      Caption         =   "&Exit"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   400
      Index           =   10
      Left            =   9960
      Style           =   1  'Graphical
      TabIndex        =   22
      ToolTipText     =   "To Add New Record"
      Top             =   5520
      Width           =   1095
   End
   Begin VB.CommandButton cmdBtn 
      Caption         =   "&Cancel"
      CausesValidation=   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   400
      Index           =   14
      Left            =   9960
      Style           =   1  'Graphical
      TabIndex        =   21
      ToolTipText     =   "To Add New Record"
      Top             =   5040
      Width           =   1095
   End
   Begin VB.CommandButton cmdBtn 
      Caption         =   "&Save"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   400
      Index           =   13
      Left            =   9960
      Style           =   1  'Graphical
      TabIndex        =   20
      ToolTipText     =   "To Add New Record"
      Top             =   4560
      Width           =   1095
   End
   Begin VB.CommandButton cmdBtn 
      Caption         =   "&Print"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   400
      Index           =   9
      Left            =   9960
      Style           =   1  'Graphical
      TabIndex        =   19
      ToolTipText     =   "To Add New Record"
      Top             =   4080
      Width           =   1095
   End
   Begin VB.CommandButton cmdBtn 
      Caption         =   "&Delete"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   400
      Index           =   8
      Left            =   9960
      Style           =   1  'Graphical
      TabIndex        =   18
      ToolTipText     =   "To Add New Record"
      Top             =   3600
      Width           =   1095
   End
   Begin VB.CommandButton cmdBtn 
      Caption         =   "&Modify"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Index           =   7
      Left            =   9960
      Style           =   1  'Graphical
      TabIndex        =   17
      ToolTipText     =   "To Add New Record"
      Top             =   3120
      Width           =   1095
   End
   Begin VB.CommandButton cmdBtn 
      Caption         =   "&Add"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   400
      Index           =   6
      Left            =   9960
      Style           =   1  'Graphical
      TabIndex        =   16
      ToolTipText     =   "To Add New Record"
      Top             =   2640
      Width           =   1095
   End
   Begin VB.CommandButton cmdBtn 
      Caption         =   "&List"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   400
      Index           =   5
      Left            =   9960
      Style           =   1  'Graphical
      TabIndex        =   15
      ToolTipText     =   "To Add New Record"
      Top             =   2160
      Width           =   1095
   End
   Begin VB.CommandButton cmdBtn 
      Caption         =   "&Find"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   400
      Index           =   4
      Left            =   9960
      Style           =   1  'Graphical
      TabIndex        =   14
      ToolTipText     =   "To Add New Record"
      Top             =   1680
      Width           =   1095
   End
   Begin VB.CommandButton cmdBtn 
      Caption         =   ">>"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   400
      Index           =   3
      Left            =   10560
      Style           =   1  'Graphical
      TabIndex        =   13
      ToolTipText     =   "To Add New Record"
      Top             =   1200
      Width           =   495
   End
   Begin VB.CommandButton cmdBtn 
      Caption         =   "<<"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   400
      Index           =   2
      Left            =   9960
      Style           =   1  'Graphical
      TabIndex        =   12
      ToolTipText     =   "To Add New Record"
      Top             =   1200
      Width           =   495
   End
   Begin VB.CommandButton cmdBtn 
      Caption         =   ">"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   400
      Index           =   1
      Left            =   10560
      Style           =   1  'Graphical
      TabIndex        =   11
      ToolTipText     =   "To Add New Record"
      Top             =   720
      Width           =   495
   End
   Begin VB.CommandButton cmdBtn 
      Caption         =   "<"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   400
      Index           =   0
      Left            =   9960
      Style           =   1  'Graphical
      TabIndex        =   10
      ToolTipText     =   "To Add New Record"
      Top             =   720
      Width           =   495
   End
   Begin VB.TextBox txt 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Index           =   7
      Left            =   11160
      MaxLength       =   3
      TabIndex        =   9
      Text            =   "7"
      Top             =   360
      Visible         =   0   'False
      Width           =   315
   End
   Begin VB.TextBox txt 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Index           =   6
      Left            =   10440
      MaxLength       =   3
      TabIndex        =   8
      Text            =   "6"
      Top             =   360
      Visible         =   0   'False
      Width           =   315
   End
   Begin VB.TextBox txt 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Index           =   5
      Left            =   9960
      MaxLength       =   3
      TabIndex        =   7
      Text            =   "5"
      Top             =   360
      Visible         =   0   'False
      Width           =   315
   End
   Begin VB.TextBox txt 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Index           =   4
      Left            =   11040
      MaxLength       =   3
      TabIndex        =   6
      Text            =   "4"
      Top             =   0
      Visible         =   0   'False
      Width           =   315
   End
   Begin VB.TextBox txt 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Index           =   3
      Left            =   10560
      MaxLength       =   3
      TabIndex        =   5
      Text            =   "3"
      Top             =   0
      Visible         =   0   'False
      Width           =   315
   End
   Begin VB.TextBox txt 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Index           =   2
      Left            =   10200
      MaxLength       =   3
      TabIndex        =   4
      Text            =   "2"
      Top             =   0
      Visible         =   0   'False
      Width           =   315
   End
   Begin VB.TextBox txt 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Index           =   1
      Left            =   9720
      MaxLength       =   3
      TabIndex        =   3
      Text            =   "1"
      Top             =   0
      Visible         =   0   'False
      Width           =   315
   End
   Begin VB.TextBox txt 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Index           =   0
      Left            =   1440
      MaxLength       =   3
      TabIndex        =   1
      Text            =   "0"
      Top             =   240
      Width           =   8115
   End
   Begin MSComctlLib.TreeView TreeView1 
      Height          =   6615
      Left            =   120
      TabIndex        =   2
      Top             =   720
      Width           =   9495
      _ExtentX        =   16748
      _ExtentY        =   11668
      _Version        =   393217
      Style           =   7
      Checkboxes      =   -1  'True
      Appearance      =   1
   End
   Begin VB.Shape Shape1 
      FillColor       =   &H00BDDFD1&
      FillStyle       =   0  'Solid
      Height          =   6495
      Left            =   9720
      Shape           =   4  'Rounded Rectangle
      Top             =   600
      Width           =   1575
   End
   Begin VB.Label Label14 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "User Name"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   1140
   End
End
Attribute VB_Name = "frmUserPrevilage"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Const tblName As String = "tblUserPrevilage"
Const OrderField As String = "Username"
Const CriteriaStr As String = " srno=1"
Dim rstbl As Recordset
Dim varBookMark As Variant
Dim FindFlag As Boolean
Dim FirstTimeFlag As Boolean
Dim SrNo As Long

Private Sub cmdBtn_Click(Index As Integer)
On Error GoTo ErrorRoutine
Dim NodeCn As Long
Select Case Index
Case 0 'Previous
    rstbl.MovePrevious
    If rstbl.BOF = True Then
        rstbl.MoveFirst
        MsgBox ("This is First Record"), vbInformation + vbOKOnly, "First Record"
    End If
    ReadFields
Case 1 'Next
    rstbl.MoveNext
    If rstbl.EOF = True Then
        rstbl.MoveLast
        MsgBox ("This is Last Record"), vbInformation + vbOKOnly, "Last Record"
    End If
    ReadFields
Case 2 'First
    rstbl.MoveFirst
    ReadFields
Case 3 'Last
    rstbl.MoveLast
    ReadFields
Case 4 'Find
Case 5 'List
    'GenMastList ("LAC")
Case 6 'Add
    gFormAction = vbDataActionAddNew
    Call GProActivateControls(Me, True)
    Call GProSetButtons(Me, vbDataActionAddNew)
    Call GProClearForm(Me, rstbl)
    FillTreeViewFromMenu
    If FirstTimeFlag = True Then
        FirstTimeFlag = False
    Else
        txt(0).SetFocus
    End If
Case 7 'Modify
    'If mURecModi = True Then
    '    SearchHelp
    '    DoEvents
        gFormAction = vbDataActionUpdate
        Call GProActivateControls(Me, True)
        Call GProSetButtons(Me, vbDataActionUpdate)
        txt(0).Enabled = False '-- User Name
        TreeView1.SetFocus
    'End If
Case 8 'Delete
    'If mURecDel = True Then
        If MsgBox("Do You Really Delete This Record?", vbQuestion + vbYesNo + vbDefaultButton1, "Confirm Delete") = vbYes Then
            If rstbl.RecordCount > 1 Then
                '-- Multi Records
'                varBookMark = rstbl.Bookmark
'                rstbl.Find " Username='" & txt(0) & "'"
'                While Not rstbl.EOF
'                    rstbl.Delete
'                    rstbl.MoveNext
'                Wend
'                rstbl.Bookmark = varBookMark
                gSetupCn.Execute "delete from tbluserPrevilage where Username='" & txt(0) & "'"
                cmdBtn_Click (1)
            ElseIf rstbl.RecordCount = 1 Then
                '-- Multi Records
'                varBookMark = rstbl.Bookmark
'                rstbl.Find " Username='" & txt(0) & "'"
'                While Not rstbl.EOF
'                    rstbl.Delete
'                    rstbl.MoveNext
'                Wend
                'rstbl.Bookmark = varBookMark
                gSetupCn.Execute "delete from tbluserPrevilage where Username='" & txt(0) & "'"
                Call GProActivateControls(Me, False)
                Call GProSetButtons(Me, vbDataActionCancel)
                Form_Load
            End If
        End If
        DoEvents
    'End If
Case 9 'Print
Case 10 'Exit
    DoEvents
    Unload Me
Case 11 'Label
    'frmLabel.Show 1
Case 12 'Optional
    
Case 13 'Save
    If ValidateData = True Then
        '------------------ Save Multi Record -------------
        gSetupCn.Execute "delete from tblUserPrevilage where username='" & txt(0) & "'"
        gFormAction = vbDataActionAddNew
        SrNo = 1
        For NodeCn = 1 To TreeView1.Nodes.Count
            If TreeView1.Nodes(NodeCn).Checked = True Then
                FillTxtFromTemp NodeCn
                If LCase(Right(txt(1), 3)) <> "frm" Or (LCase(Right(txt(1), 3)) = "frm" And (txt(3) = 1 Or txt(4) = 1 Or txt(5) = 1 Or txt(6) = 1)) Then
                    SrNo = SrNo + 1
                    GProSaveRecord Me, rstbl, gFormAction
                End If
            End If
        Next NodeCn
        '-------------------------------------------------
        Call GProActivateControls(Me, False)
        Call GProSetButtons(Me, True)
        gFormAction = vbDataActionClose
    End If
Case 14 'Cancel
    Call GProClearForm(Me, rstbl)
    If rstbl.RecordCount <= 0 Then
        Call GProActivateControls(Me, False)
        Call GProSetButtons(Me, vbDataActionCancel)
    Else
        If gFormAction = vbDataActionAddNew Then
            rstbl.MoveLast
        Else
            rstbl.CancelUpdate
        End If
        ReadFields
        Call GProActivateControls(Me, False)
        Call GProSetButtons(Me, True)
    End If
    gFormAction = vbDataActionClose
End Select
Exit Sub
ErrorRoutine:
    GProcErrorHandler
End Sub
Private Sub Form_Load()
    Dim x As String
    Me.Top = 300
    Me.Left = Screen.Width / 2 - Me.Width / 2
    x = "Select * from " & tblName
    If Trim(CriteriaStr) <> "" Then x = x & " where " & CriteriaStr
    If Trim(OrderField) <> "" Then x = x & " order by " & OrderField
    GProcRstOpen rstbl, x, "O", gSetupCn
    FirstTimeFlag = True
    cmdBtn_Click (6)
End Sub
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
If gFormAction = vbDataActionUpdate Then
    MsgBox "You are about to modify the record. Save the record and then exit.", vbInformation + vbOKOnly, Me.Caption
    Cancel = True
    Exit Sub
End If
rstbl.Close
End Sub
'---- Set All Child Nodes Checked
Private Sub ProcSetChildNode(pNode As Node)
Dim ChildNode As Node
Dim IsChecked As Boolean
Dim ChildCount As Long
Dim i As Long
Dim ChkNode As Node
Set ChkNode = pNode

IsChecked = ChkNode.Checked
If ChkNode.Children > 0 Then
    i = ChkNode.Child.FirstSibling.Index
    Set ChildNode = TreeView1.Nodes(i)
    For ChildCount = 1 To ChkNode.Children
        ChildNode.Checked = IsChecked
        ProcSetChildNode ChildNode
        Set ChildNode = ChildNode.Next
        If ChildNode Is Nothing Then Exit For
     Next
End If
End Sub
Private Sub TreeView1_NodeCheck(ByVal Node As MSComctlLib.Node)
Dim i As Long
Dim IsSelected As Boolean
Dim ParentKey As String
Dim ParentNode As Node
Dim ChkNode As Node
'================== if node checked Set All Child Nodes Checked
ProcSetChildNode Node
'================== if One of Child Nodes Check then set All Parent Nodes checked
 Set ChkNode = Node
 Set ParentNode = ChkNode.Parent
 While Not ParentNode Is Nothing
     ParentKey = ChkNode.Parent.Key
     IsSelected = ChkNode.Checked
     i = ChkNode.FirstSibling.Index
     Do While i <= TreeView1.Nodes.Count
         If Not TreeView1.Nodes(i).Parent Is Nothing Then
             If TreeView1.Nodes(i).Parent.Key = ParentKey Then IsSelected = IsSelected Or TreeView1.Nodes(i).Checked
         End If
         i = i + 1
     Loop
     ChkNode.Parent.Checked = IsSelected
     Set ChkNode = ChkNode.Parent
     Set ParentNode = ChkNode.Parent
Wend
End Sub

Private Sub txt_GotFocus(Index As Integer)
If TypeOf Me.ActiveControl Is TextBox Then GProcSelectBox Me.ActiveControl
End Sub
Private Sub txt_LostFocus(Index As Integer)
GProcCheckForNumber Me, rstbl, Index
SaveBtnEd
End Sub
Private Sub txt_KeyPress(Index As Integer, KeyAscii As Integer)
If Index = 0 Then '-- User Name
    GProcSearchField Me, "tblMastUser", "UserName", " where isAdminUser=0 ", False, Chr(KeyAscii), gSetupCn
    KeyAscii = 0
    txt(Index) = gSearchRetcol1
Else
    KeyAscii = GProcValidateKey(rstbl, Index, KeyAscii, txt(Index).Text)
End If
End Sub

Private Sub txt_Validate(Index As Integer, Cancel As Boolean)
If Index = 0 Then '-- User Name
        If Trim(txt(Index)) = "" Then
            MsgBox "Enter User Name.", vbCritical + vbOKOnly, Me.Caption
            Cancel = True
            Exit Sub
        End If
End If
End Sub

Private Sub SaveBtnEd()
'--- User Name
If txt(0) = "" Then
    cmdBtn(13).Enabled = False
    cmdBtn(13).BackColor = &H80000004
Else
    cmdBtn(13).Enabled = True
    cmdBtn(13).BackColor = &HFFFF80
End If
End Sub
Private Function ValidateData() As Boolean
Dim i As Long
Dim x As String
Dim Y As String
Dim rstAddCheck As Recordset
Dim rstModiCheck As Recordset
    '---- Required Fields
        '--- User Name
        If Trim(txt(0)) = "" Then
            MsgBox "Enter User Name.", vbCritical + vbOKOnly, Me.Caption
            txt(0).SetFocus
            Exit Function
        End If
    '----Duplicate Records
    '----Addition Case
    If gFormAction = vbDataActionAddNew Then
        x = "Select username From " & tblName & " Where Username='" & Trim(txt(0).Text) & "'"
        i = GProcRstOpen(rstAddCheck, x, "R", gSetupCn)
        If i > 0 Then
            MsgBox "Record is Available of User Name." & Trim(txt(0).Text), vbCritical + vbOKOnly, "Duplicate Record"
            txt(0).SetFocus
            Exit Function
        End If
        rstAddCheck.Close
    End If
    ValidateData = True
End Function
Private Sub FillTxtFromTemp(Optional NIndex As Long)
'-------------- Fill Txt From Tree view
Dim MenuName As String
Dim TempStr As String
Dim MenuIndex As Long
Dim AddStatus As Long
Dim ModStatus As Long
Dim DeleteStatus As Long
Dim ViewStatus As Long
With TreeView1.Nodes(NIndex)
        MenuName = Left(.Key, InStr(1, .Key, "-") - 1)
        TempStr = Right(.Key, Len(.Key) - InStr(1, .Key, "-"))
        AddStatus = 0
        ModStatus = 0
        DeleteStatus = 0
        ViewStatus = 0
        Select Case Left(TempStr, 1)
            Case "A"
                AddStatus = 1
            Case "M"
                ModStatus = 1
            Case "D"
                DeleteStatus = 1
            Case "V"
                ViewStatus = 1
        End Select
        If IsNumeric(TempStr) Then
            MenuIndex = CLng(TempStr)
        Else
            MenuIndex = CLng(Mid(TempStr, 2))
        End If
        txt(1) = MenuName
        txt(2) = MenuIndex
        txt(3) = AddStatus
        txt(4) = ModStatus
        txt(5) = DeleteStatus
        txt(6) = ViewStatus
        txt(7) = SrNo
End With
'---------------- end
End Sub
Private Sub FillTempFromTxt()
FillTreeViewFromMenu
End Sub
Private Sub ReadFields()
Dim MNname As String
Dim MNIndex As Long
Dim MNKey As String
Dim Rs As Recordset
Dim x As String
'-- Fill Txt Contrl
GProShowRecord Me, rstbl
'-- Fill Temp Control
FillTempFromTxt
'--- Fill Tree View
x = "select * from " & tblName & " where username='" & txt(0) & "' order by srno"
GProcRstOpen Rs, x, "O", gSetupCn
If Rs.EOF And Rs.BOF Then
Else
    Rs.MoveFirst
End If
While Not Rs.EOF
    MNname = Rs.Fields("menuName")
    MNIndex = Rs.Fields("Menuindex")
    MNKey = MNname & "-" & MNIndex
    Dim n As Node
    
    If IsNodeExist(MNKey) = True Then
    'IsObject (TreeView1.Nodes(MNKey))=True then
    TreeView1.Nodes(MNKey).Checked = True
    '--- Add,Modify,Delete,View Nodes
    If Rs.Fields("Add") = 1 Then
        MNKey = MNname & "-A" & MNIndex
        TreeView1.Nodes(MNKey).Checked = True
    End If
    If Rs.Fields("Modify") = 1 Then
        MNKey = MNname & "-M" & MNIndex
        TreeView1.Nodes(MNKey).Checked = True
    End If
    If Rs.Fields("delete") = 1 Then
        MNKey = MNname & "-D" & MNIndex
        TreeView1.Nodes(MNKey).Checked = True
    End If
    If Rs.Fields("View") = 1 Then
        MNKey = MNname & "-V" & MNIndex
        TreeView1.Nodes(MNKey).Checked = True
    End If
    End If
    Rs.MoveNext
Wend
End Sub
Private Function IsNodeExist(NodeKey As String) As Boolean
Dim i As Long
For i = 1 To TreeView1.Nodes.Count
    If TreeView1.Nodes(i).Key = NodeKey Then
        IsNodeExist = True
        Exit For
    End If
Next
End Function
'----- Fill Menu In Treeview
Private Sub FillTreeViewFromMenu()
'On Error GoTo ErrorHandler
Dim i As Long
Dim MnuCtrl As Menu
Dim RootStr(0 To 30) As String, TempStr As String
Dim RootCount As Long
Dim MnuIndex As String
Dim MnuCaption As String
Dim j As Long

RootCount = 0
TreeView1.Nodes.Clear
For Each MnuCtrl In HiMain
Select Case LCase(Left(MnuCtrl.Name, 6))
'============= Previlage for menu "mnutl", "mnutrs", "mnurpt", "mnumst" Only
    Case "mnuutl", "mnutrs", "mnurpt", "mnumst"
        If Trim(MnuCtrl.Caption) <> "-" Then
            '--- Child
            For i = RootCount To 1 Step -1
               TempStr = Left(RootStr(i - 1), InStr(1, RootStr(i - 1), "-") - 1)
               MnuIndex = MnuCtrl.Index
                If InStr(1, MnuCtrl.Name, TempStr) > 0 And (MnuCtrl.Name <> TempStr) Then
                    '--------------------- Remove '&' Character
                    MnuCaption = MnuCtrl.Caption
                    MnuCaption = Replace(MnuCaption, "&&", " and ")
                    MnuCaption = Replace(MnuCaption, "&", "")
                    MnuCaption = Replace(MnuCaption, " and ", "&")
                    '----------------------
                    TreeView1.Nodes.Add RootStr(i - 1), tvwChild, MnuCtrl.Name & "-" & MnuIndex, MnuCaption
                    TreeView1.Nodes(MnuCtrl.Name & "-" & MnuIndex).EnsureVisible
                    RootStr(i) = MnuCtrl.Name & "-" & MnuIndex
                    RootCount = i + 1 ' - 1
                    '---- Add Transaction Nodes [Add,Modify,Delete,View]
                    If LCase(Right(Trim(MnuCtrl.Name), 3)) = "frm" Then AddTransactionNodes MnuCtrl, RootStr(i), MnuIndex
                    Exit For
                End If
            Next i
            '--- Root
            If i = 0 Then
                '--------------------- Remove '&' Character
                MnuCaption = MnuCtrl.Caption
                MnuCaption = Replace(MnuCaption, "&&", " and ")
                MnuCaption = Replace(MnuCaption, "&", "")
                MnuCaption = Replace(MnuCaption, " and ", "&")
                '----------------------
                MnuIndex = MnuCtrl.Index
                TreeView1.Nodes.Add , , MnuCtrl.Name & "-" & MnuIndex, MnuCaption
                TreeView1.Nodes(MnuCtrl.Name & "-" & MnuIndex).EnsureVisible
                RootStr(i) = MnuCtrl.Name & "-" & MnuIndex
                '---- Add Transaction Nodes [Add,Modify,Delete,View]
                If LCase(Right(Trim(MnuCtrl.Name), 3)) = "frm" Then AddTransactionNodes MnuCtrl, RootStr(i), MnuIndex
                RootCount = 1
            Else
            End If
         End If
    End Select
Next
Exit Sub
ErrorHandler:
    MsgBox Err.Description, vbOKOnly, Me.Caption
End Sub
'---- Add Transaction Nodes [Add,Modify,Delete,View]
Private Sub AddTransactionNodes(mnu As Menu, RtStr As String, NodeCnt As String)
    Dim keyStr As String
    '-- Add
    keyStr = mnu.Name & "-A" & NodeCnt
    TreeView1.Nodes.Add RtStr, tvwChild, keyStr, "Add"
    TreeView1.Nodes(keyStr).EnsureVisible
    '-- Modify
    keyStr = mnu.Name & "-M" & NodeCnt
    TreeView1.Nodes.Add RtStr, tvwChild, keyStr, "Modify"
    TreeView1.Nodes(keyStr).EnsureVisible
    '-- Delete
    keyStr = mnu.Name & "-D" & NodeCnt
    TreeView1.Nodes.Add RtStr, tvwChild, keyStr, "Delete"
    TreeView1.Nodes(keyStr).EnsureVisible
    '-- View
    keyStr = mnu.Name & "-V" & NodeCnt
    TreeView1.Nodes.Add RtStr, tvwChild, keyStr, "View"
    TreeView1.Nodes(keyStr).EnsureVisible
End Sub

