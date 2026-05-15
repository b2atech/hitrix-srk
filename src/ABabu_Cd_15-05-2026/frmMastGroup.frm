VERSION 5.00
Begin VB.Form frmMastGroup 
   BackColor       =   &H00FBF2E1&
   BorderStyle     =   0  'None
   Caption         =   " Account Group Master"
   ClientHeight    =   4410
   ClientLeft      =   0
   ClientTop       =   -105
   ClientWidth     =   9345
   FillColor       =   &H00FBF2E1&
   Icon            =   "frmMastGroup.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4410
   ScaleWidth      =   9345
   ShowInTaskbar   =   0   'False
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
      Left            =   7650
      Style           =   1  'Graphical
      TabIndex        =   13
      ToolTipText     =   "Previous"
      Top             =   795
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
      Left            =   8250
      Style           =   1  'Graphical
      TabIndex        =   14
      ToolTipText     =   "Next"
      Top             =   795
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
      Left            =   7650
      Style           =   1  'Graphical
      TabIndex        =   15
      ToolTipText     =   "First"
      Top             =   1275
      Width           =   495
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
      Left            =   8250
      Style           =   1  'Graphical
      TabIndex        =   16
      ToolTipText     =   "Last"
      Top             =   1275
      Width           =   495
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
      Left            =   7650
      Style           =   1  'Graphical
      TabIndex        =   12
      ToolTipText     =   "To Find Record"
      Top             =   1755
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
      Left            =   7650
      Style           =   1  'Graphical
      TabIndex        =   17
      Top             =   2235
      Visible         =   0   'False
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
      Left            =   465
      Style           =   1  'Graphical
      TabIndex        =   7
      ToolTipText     =   "To Add New Record"
      Top             =   3390
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
      Height          =   400
      Index           =   7
      Left            =   1680
      Style           =   1  'Graphical
      TabIndex        =   8
      ToolTipText     =   "To Modify Record"
      Top             =   3390
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
      Left            =   2895
      Style           =   1  'Graphical
      TabIndex        =   9
      ToolTipText     =   "To Delete Record"
      Top             =   3390
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
      Left            =   12150
      Style           =   1  'Graphical
      TabIndex        =   19
      Top             =   3705
      Visible         =   0   'False
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
      Left            =   4110
      Style           =   1  'Graphical
      TabIndex        =   10
      ToolTipText     =   "To Save Record"
      Top             =   3390
      Width           =   1095
   End
   Begin VB.CommandButton cmdBtn 
      Caption         =   "&Cancel"
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
      Left            =   5325
      Style           =   1  'Graphical
      TabIndex        =   11
      ToolTipText     =   "To Cancel Record"
      Top             =   3390
      Width           =   1095
   End
   Begin VB.CommandButton cmdBtn 
      Cancel          =   -1  'True
      Caption         =   "E&xit"
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
      Left            =   7650
      Style           =   1  'Graphical
      TabIndex        =   18
      ToolTipText     =   "Exit"
      Top             =   2730
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
      Left            =   12045
      Style           =   1  'Graphical
      TabIndex        =   20
      Top             =   2190
      Visible         =   0   'False
      Width           =   1095
   End
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
      Left            =   12045
      Style           =   1  'Graphical
      TabIndex        =   21
      Top             =   2670
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.CheckBox Chk 
      BackColor       =   &H00FBF2E1&
      Caption         =   "Is Main Group ?"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   375
      Index           =   5
      Left            =   4020
      TabIndex        =   6
      Top             =   2355
      Width           =   1920
   End
   Begin VB.TextBox Txt 
      Height          =   285
      Index           =   5
      Left            =   12480
      TabIndex        =   25
      Text            =   "5"
      Top             =   870
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.TextBox Txt 
      Height          =   285
      Index           =   4
      Left            =   12360
      TabIndex        =   24
      Text            =   "4"
      Top             =   1170
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.TextBox Txt 
      Height          =   285
      Index           =   2
      Left            =   12600
      TabIndex        =   23
      Text            =   "2"
      Top             =   1755
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.TextBox Txt 
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
      Left            =   2355
      MaxLength       =   1
      TabIndex        =   5
      Text            =   "3"
      Top             =   2340
      Width           =   1035
   End
   Begin VB.TextBox Txt 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Index           =   1
      Left            =   2355
      MaxLength       =   40
      TabIndex        =   1
      Text            =   "1"
      Top             =   1140
      Width           =   4695
   End
   Begin VB.TextBox txtTemp 
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
      Left            =   2355
      Locked          =   -1  'True
      MaxLength       =   40
      TabIndex        =   3
      Text            =   "txttemp(2)"
      Top             =   1725
      Width           =   4695
   End
   Begin VB.TextBox Txt 
      Height          =   285
      Index           =   0
      Left            =   12420
      TabIndex        =   22
      Text            =   "0"
      Top             =   870
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Shape Shape5 
      FillColor       =   &H00FED8FC&
      FillStyle       =   0  'Solid
      Height          =   780
      Left            =   270
      Shape           =   4  'Rounded Rectangle
      Top             =   3210
      Width           =   6315
   End
   Begin VB.Shape Shape3 
      BorderColor     =   &H00EFB17A&
      BorderWidth     =   5
      FillColor       =   &H00EFB17A&
      Height          =   4350
      Left            =   0
      Top             =   0
      Width           =   9315
   End
   Begin VB.Label lblActionStatus 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "View Record.."
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   375
      Left            =   6075
      TabIndex        =   26
      Top             =   30
      Width           =   2475
   End
   Begin VB.Label lbl1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Sub&group Name  :"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   240
      Index           =   0
      Left            =   480
      TabIndex        =   0
      Top             =   1140
      Width           =   1725
   End
   Begin VB.Label lbl1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Mai&n Group         :"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   240
      Index           =   1
      Left            =   480
      TabIndex        =   2
      Top             =   1725
      Width           =   1710
   End
   Begin VB.Label lbl1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "S&chedule Name     :"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   240
      Index           =   2
      Left            =   480
      TabIndex        =   4
      Top             =   2340
      Width           =   1740
   End
   Begin VB.Label Label1 
      BackColor       =   &H00EFB17A&
      Caption         =   "   Account Group Master"
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
      TabIndex        =   27
      Top             =   -15
      Width           =   9315
   End
   Begin VB.Shape Shape4 
      FillColor       =   &H00FED8FC&
      FillStyle       =   0  'Solid
      Height          =   2760
      Left            =   7425
      Shape           =   4  'Rounded Rectangle
      Top             =   600
      Width           =   1575
   End
End
Attribute VB_Name = "frmMastGroup"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Const tblName As String = "tblMastGroup"
Const IdField As String = "AgCode"
Const OrderField As String = "AgName"
Public CriteriaStr As String
Const MaxNo As Long = 5
Dim rstbl As Recordset
Dim varBookMark As Variant
Dim FirstTimeFlag As Boolean
Dim FormAction As Integer
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case 13:
            Sendkeys "{TAB}"
        Case 27:
            If (cmdBtn(13).Enabled = False) Then
                Unload Me
            End If
    End Select
End Sub

Private Sub cmdBtn_Click(Index As Integer)
'On Error GoTo ErrorRoutine
Select Case Index
'Case 0 'Previous
'    rstbl.MovePrevious
'    If rstbl.BOF = True Then
'        rstbl.MoveFirst
'        MsgBox ("This is First Record"), vbInformation + vbOKOnly, "First Record"
'    End If
'    ReadFields
'Case 1 'Next
'    rstbl.MoveNext
'    If rstbl.EOF = True Then
'        rstbl.MoveLast
'        MsgBox ("This is Last Record"), vbInformation + vbOKOnly, "Last Record"
'    End If
'    ReadFields
'Case 2 'First
'    rstbl.MoveFirst
'    ReadFields
'Case 3 'Last
'    rstbl.MoveLast
'    ReadFields
Case 0 '---- Previous
     If GProcRstOpen(rstbl, "Select * from " & tblName & " where " & OrderField & "  = (Select max(" & OrderField & " ) from " & tblName & " where  " & OrderField & " < '" & txt(1) & "' and " & CriteriaStr & " )and " & CriteriaStr, "R") > 0 Then ReadFields Else MsgBox "This is First Record:", vbOKOnly + vbInformation
Case 1 '---- Next
      If GProcRstOpen(rstbl, "Select * from " & tblName & " where " & OrderField & " = (Select min(" & OrderField & ") from " & tblName & " where " & OrderField & " > '" & txt(1) & "' and " & CriteriaStr & ")and " & CriteriaStr, "R") > 0 Then ReadFields Else MsgBox "This is last Record:", vbOKOnly + vbInformation
Case 2 '---- First
         If GProcRstOpen(rstbl, "Select * from " & tblName & " where " & OrderField & " = (Select min(" & OrderField & ") from " & tblName & " where " & OrderField & " <> '' and " & CriteriaStr & ")and " & CriteriaStr, "R") > 0 Then ReadFields
Case 3 '---- Last
        If GProcRstOpen(rstbl, "Select * from " & tblName & " where " & OrderField & " = (Select max(" & OrderField & ") from " & tblName & " where " & OrderField & " <> '' and " & CriteriaStr & ")and " & CriteriaStr, "R") > 0 Then ReadFields Else GProcClearForm Me, rstbl, MaxNo
Case 4 'Find
    gClsSearch.SearchMultiField tblName, "AgName,Schedule,iif(isFixGroup=0,'No','Yes')", Array("Account Group", "Schedule", "Fix Group"), Array(txt(1).Width, 1000, 1000), "", "", "AgName", txt(1).Left + Me.Left, txt(1).Top + Me.Top + 650
     If gClsSearch.SearchMultiRetCol(0) <> "" Then
        GProcRstOpen rstbl, "Select * from " & tblName & " where " & OrderField & " = '" & gClsSearch.SearchMultiRetCol(0) & "' and " & CriteriaStr, "R"
        gClsSearch.SearchMultiRetCol(1) = ""
        ReadFields
    End If
    
    
    
    
'    If gBackEndDB = gBackEndAccess Then
'        gClsSearch.SearchMultiField tblName, "AgName,Schedule,iif(isFixGroup=0,'No','Yes')", Array("Account Group", "Schedule", "Fix Group"), Array(txt(1).Width, 1000, 1000), "", "", "AgName", txt(1).Left + Me.Left, txt(1).Top + Me.Top + 650
'    Else '-- Oracle
'        gClsSearch.SearchMultiField tblName, "AgName,Schedule,decode(isFixGroup,0,'No','Yes')", Array("Account Group", "Schedule", "Fix Group"), Array(txt(1).Width, 1000, 1000), "", "", "AgName", txt(1).Left + Me.Left, txt(1).Top + Me.Top + 650
'    End If
'    If gClsSearch.SearchMultiRetCol(0) <> "" Then
'        rstbl.MoveFirst
'        rstbl.Find "AgName = '" & gClsSearch.SearchMultiRetCol(0) & "'"
'        gClsSearch.SearchMultiRetCol(0) = ""
'        ReadFields
'    End If
Case 5 'List
    'GenMastList ("LAC")
Case 6 'Add
    FormAction = vbDataActionAddNew
    Call GProcActivateControls(Me, True)
    Call GProcSetButtons(Me, vbDataActionAddNew)
    Call GProcClearForm(Me, rstbl, MaxNo, True)
    If FirstTimeFlag = True Then
        FirstTimeFlag = False
    Else
'        txt(1).SetFocus
    End If
    FillCombo
Case 7 'Modify
    'If mURecModi = True Then
    '    SearchHelp
    '    DoEvents
        cmdBtn_Click 4 '-- Find
        If gClsSearch.SearchMultiRetCol(0) = "" Then Exit Sub
'        If Chk(5) = 1 Then
'            MsgBox "This is default Group.You can not Modify/Delete this Record.", vbCritical + vbOKOnly, Me.Caption
'            Exit Sub
'        End If
        FormAction = vbDataActionUpdate
        Call GProcActivateControls(Me, True)
        Call GProcSetButtons(Me, vbDataActionUpdate)
        txt(1).SetFocus
        If Chk(5) = 1 Then
            Chk(5).Enabled = False
            txttemp(2).Enabled = False
            txt(1).Enabled = False
        End If
    'End If
Case 8 'Delete
    lblActionStatus.Caption = "Delete Record..."
    'If mURecDel = True Then
        If Chk(5) = 1 Then
            MsgBox "This is default Group.You can not Modify/Delete this Record.", vbCritical + vbOKOnly, Me.Caption
            Exit Sub
        End If
        If MsgBox("Do You Really Delete This Record?", vbQuestion + vbYesNo + vbDefaultButton1, "Confirm Delete") = vbYes Then
            gCn.Execute "Delete from " & tblName & " where " & IdField & " =  " & txt(0) & " and " & CriteriaStr, "O"
        End If
        cmdBtn_Click (3)
    'End If
Case 9 'Print
Case 10 'Exit
    Unload Me
Case 11 'Label
    'frmLabel.Show 1
Case 12 'Optional
Case 13 'Save
    If ValidateData = True Then
        FillTxtFromTemp
        GProcRstOpen rstbl, "Select * from " & tblName & " where " & IdField & " =  " & txt(0) & " and " & CriteriaStr, "O"
        GProcSaveRecord Me, rstbl, FormAction, MaxNo
        Call GProcActivateControls(Me, False)
        Call GProcSetButtons(Me, True)
        If FormAction = vbDataActionAddNew Then
            If cmdBtn(6).Visible = True Then cmdBtn(6).SetFocus
        Else
            If cmdBtn(7).Visible = True Then cmdBtn(7).SetFocus
        End If
        FormAction = vbDataActionClose
    End If
Case 14 'Cancel
    cmdBtn_Click (3)
    If rstbl.RecordCount <= 0 Then
        Call GProcActivateControls(Me, False)
        Call GProcSetButtons(Me, vbDataActionCancel)
    Else
        Call GProcActivateControls(Me, False)
        Call GProcSetButtons(Me, True)
    End If
    If FormAction = vbDataActionAddNew Then
        If cmdBtn(6).Visible = True Then cmdBtn(6).SetFocus
    Else
        If cmdBtn(7).Visible = True Then cmdBtn(7).SetFocus
    End If
    FormAction = vbDataActionClose
End Select
Exit Sub
ErrorRoutine:
    rstbl.CancelUpdate
    GProcErrorHandler
End Sub
Private Sub Form_Load()
    Dim x As String
    Me.Top = 300
    Me.Left = Screen.Width / 2 - Me.Width / 2
    x = "Select * from " & tblName
    CriteriaStr = " AgCode > 1"
    
    If Trim(CriteriaStr) <> "" Then x = x & " where " & CriteriaStr
    If Trim(OrderField) <> "" Then x = x & " order by " & OrderField
'    GProcRstOpen rstbl, x, "O"
    If Trim(mShowRec) <> "" Then
       GProcRstOpen rstbl, "Select * from " & tblName & " where " & IdField & " = " & CLng(Trim(mShowRec)), "R"
       ReadFields
       Call GProcSetButtons(Me, vbDataActionCancel)
       mShowRec = ""
    Else
       GProcRstOpen rstbl, "Select * from " & tblName & " where " & IdField & " = 0", "R"
       cmdBtn_Click (6)
    End If
    
    'MaxNo = rstbl.Fields.Count - 1
    FirstTimeFlag = True
    '---- Button Visiblity Set As Per User Previlage
    GProcSetButtonVisiblity Me
    cmdBtn_Click (6)
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
If FormAction = vbDataActionUpdate Then
    MsgBox "You are about to modify the record. Save the record and then exit.", vbInformation + vbOKOnly, Me.Caption
    Cancel = True
    Exit Sub
End If
rstbl.Requery
rstbl.Close
End Sub

Private Sub TabStrip1_Click()

End Sub

Private Sub txt_GotFocus(Index As Integer)
GProcSelectBox Me.ActiveControl
End Sub
Private Sub txt_LostFocus(Index As Integer)
If Index = 1 Then txt(1) = GProcProperCase(txt(1))
GProcCheckForNumber Me, rstbl, MaxNo
SaveBtnEd
End Sub
Private Sub txt_KeyPress(Index As Integer, KeyAscii As Integer)
If KeyAscii = 13 Then Exit Sub
KeyAscii = GProcValidateKey(rstbl, Index, KeyAscii, txt(Index).text)
End Sub
Private Sub txttemp_GotFocus(Index As Integer)
GProcSelectBox Me.ActiveControl
End Sub
Private Sub txtTemp_KeyPress(Index As Integer, KeyAscii As Integer)
Dim LeftPos As Long
Dim TopPos As Long
TopPos = txttemp(Index).Top + Me.Top + 650
LeftPos = txttemp(Index).Left + Me.Left
If KeyAscii = 13 Then Exit Sub
If Index = 2 Then '-- Group Name
    gClsSearch.SearchMultiField "tblMastGroup", "AGName", Array("Group Name"), Array(txttemp(Index).Width), " isFixGroup=1 ", Chr(KeyAscii), "", LeftPos, TopPos
    If gClsSearch.SearchMultiRetCol(0) <> "" Then txttemp(Index) = gClsSearch.SearchMultiRetCol(0)
    KeyAscii = 0
End If
End Sub
Private Sub txtTemp_LostFocus(Index As Integer)
SaveBtnEd
End Sub
Private Sub FillCombo()
'Call GProcFillCombo(cbo(7), "select Distinct AcCity from " & tblName & " where AcCity <> '' order by AcCity")
End Sub
Private Sub SaveBtnEd()
'--- AgName,Group Name
If txt(1) = "" Or txttemp(2) = "" Then
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
    GProcCheckForNumber Me, rstbl, MaxNo
    '--- Required fields
    '-- AgCode
    If FormAction = vbDataActionAddNew Then txt(0) = GProcGenerateId(tblName, IdField, pStartCode:=gCSAgCode)
    If gCEAgCode > 0 And CLng(txt(0)) > gCEAgCode Then
        MsgBox "Year End is completed. You can not create new group.", vbCritical + vbOKOnly, Me.Caption
        txt(1).SetFocus
        Exit Function
    End If
    '--- Group Name
    If Trim(txt(1)) = "" Then
        MsgBox "Check Group Name.", vbCritical + vbOKOnly, Me.Caption
        txt(1).SetFocus
        Exit Function
    End If
    '--- Main Group Name
    If Trim(txttemp(2)) = "" Then
        MsgBox "Check Main Group Name.", vbCritical + vbOKOnly, Me.Caption
        txttemp(2).SetFocus
        Exit Function
    End If
    '----Duplicate Records
    '---- Group Name
    '----Addition Case
    If FormAction = vbDataActionAddNew Then
        x = "Select AgCode From " & tblName & " Where AgName='" & Trim(txt(1).text) & "'"
        i = GProcRstOpen(rstAddCheck, x, "R")
        If i > 0 Then
            MsgBox "Record is Available of Group Name " & Trim(txt(1).text), vbCritical + vbOKOnly, "Duplicate Record"
            txt(1).SetFocus
            Exit Function
        End If
        rstAddCheck.Close
    End If
    '----Modification case
    If FormAction = vbDataActionUpdate Then
        x = "Select AgCode From " & tblName & " Where ((AgName='" & Trim(txt(1).text) & "') and (AgCode<>" & txt(0).text & "))"
        i = GProcRstOpen(rstModiCheck, x, "O")
        If i > 0 Then
            MsgBox "Record is Available of Group Name  " & Trim(txt(1).text), vbCritical + vbOKOnly, "Duplicate Record"
            txt(1).SetFocus
            Exit Function
        End If
        rstModiCheck.Close
    End If
    '---- Schedule
    '----Addition Case
    If FormAction = vbDataActionAddNew And txt(3) <> "" Then
        x = "Select Schedule From " & tblName & " Where Schedule='" & Trim(txt(3).text) & "'"
        i = GProcRstOpen(rstAddCheck, x, "R")
        If i > 0 Then
            MsgBox "Record is Available of Schedule " & Trim(txt(3).text), vbCritical + vbOKOnly, "Duplicate Record"
            txt(3).SetFocus
            Exit Function
        End If
        rstAddCheck.Close
    End If
    '----Modification case
    If FormAction = vbDataActionUpdate And txt(3) <> "" Then
        x = "Select AgCode From " & tblName & " Where ((Schedule='" & Trim(txt(3).text) & "') and (AgCode<>" & txt(0).text & "))"
        i = GProcRstOpen(rstModiCheck, x, "O")
        If i > 0 Then
            MsgBox "Record is Available of Schedule " & Trim(txt(3).text), vbCritical + vbOKOnly, "Duplicate Record"
            txt(3).SetFocus
            Exit Function
        End If
        rstModiCheck.Close
    End If
    ValidateData = True
End Function
Private Sub FillTxtFromTemp()
txt(2) = GProcGetColumnValue("tblmastGroup", "AgName", txttemp(2), "S", "agcode", "N")
txt(5) = Chk(5)
'Dim mTmprst1 As Recordset, mTmpGp As Long
txt(4) = GProcGetColumnValue("tblmastgroup", "AgName", txttemp(2), "S", "TopGroup", "N")
'mTmprst1, "select agcode,gpcode from tblmastgroup order by agcode", "R", gCn
'With mTmprst1
'    .MoveFirst
'    .Find "agcode = " & txt(2)
'    Do While Not .EOF
'        If !GpCode < 90004 Then Exit Do
'        mTmpGp = !GpCode
'        .MoveFirst
'        .Find "agcode = " & mTmpGp
'    Loop
'    If mTmpGp > 90000 And mTmpGp < 90004 Then txt(4) = mTmpGp
'End With
End Sub
Private Sub FillTempFromTxt()
txttemp(2) = GProcGetColumnValue("tblmastGroup", "AgCode", txt(2), "N", "agName", "N")
Chk(5) = txt(5)
End Sub
Private Sub ReadFields()
GProcShowRecord Me, rstbl, MaxNo
FillTempFromTxt
End Sub

