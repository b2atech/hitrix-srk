VERSION 5.00
Begin VB.Form frmMastNarrationTax 
   BackColor       =   &H00C2E3FC&
   BorderStyle     =   0  'None
   Caption         =   " Tax Master"
   ClientHeight    =   4245
   ClientLeft      =   0
   ClientTop       =   60
   ClientWidth     =   10695
   FillColor       =   &H000040C0&
   FillStyle       =   0  'Solid
   Icon            =   "frmMastNarrationTax.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4245
   ScaleWidth      =   10695
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdBtn 
      Caption         =   "Creation &To Firm"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   795
      Index           =   12
      Left            =   9165
      Style           =   1  'Graphical
      TabIndex        =   32
      Top             =   2445
      Width           =   1095
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
      Index           =   9
      Left            =   2400
      MaxLength       =   40
      TabIndex        =   7
      Text            =   "txttemp(9)"
      ToolTipText     =   "Press F3 For Creation"
      Top             =   2595
      Width           =   4875
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
      Index           =   8
      Left            =   2400
      MaxLength       =   40
      TabIndex        =   5
      Text            =   "txttemp(8)"
      ToolTipText     =   "Press F3 For Creation"
      Top             =   2100
      Width           =   4875
   End
   Begin VB.TextBox txt 
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
      Index           =   9
      Left            =   2820
      MaxLength       =   40
      TabIndex        =   31
      Text            =   "9"
      Top             =   3930
      Visible         =   0   'False
      Width           =   915
   End
   Begin VB.TextBox txt 
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
      Index           =   8
      Left            =   1830
      MaxLength       =   40
      TabIndex        =   30
      Text            =   "8"
      Top             =   3960
      Visible         =   0   'False
      Width           =   915
   End
   Begin VB.TextBox txt 
      Alignment       =   1  'Right Justify
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
      Left            =   2400
      MaxLength       =   40
      TabIndex        =   3
      Text            =   "7"
      Top             =   1620
      Width           =   2430
   End
   Begin VB.TextBox txt 
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
      Left            =   1170
      MaxLength       =   40
      TabIndex        =   28
      Text            =   "1"
      Top             =   3885
      Visible         =   0   'False
      Width           =   915
   End
   Begin VB.TextBox txt 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Index           =   0
      Left            =   3765
      TabIndex        =   27
      Text            =   "0"
      Top             =   3915
      Visible         =   0   'False
      Width           =   795
   End
   Begin VB.TextBox txt 
      Height          =   360
      Index           =   6
      Left            =   10545
      TabIndex        =   23
      Text            =   "6"
      Top             =   2850
      Visible         =   0   'False
      Width           =   690
   End
   Begin VB.TextBox txt 
      Height          =   360
      Index           =   5
      Left            =   10545
      TabIndex        =   22
      Text            =   "5"
      Top             =   2370
      Visible         =   0   'False
      Width           =   690
   End
   Begin VB.TextBox txt 
      Height          =   360
      Index           =   4
      Left            =   10545
      TabIndex        =   21
      Text            =   "4"
      Top             =   1890
      Visible         =   0   'False
      Width           =   690
   End
   Begin VB.TextBox txt 
      Height          =   360
      Index           =   3
      Left            =   10545
      TabIndex        =   20
      Text            =   "3"
      Top             =   1410
      Visible         =   0   'False
      Width           =   690
   End
   Begin VB.TextBox txt 
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
      Index           =   2
      Left            =   2400
      TabIndex        =   1
      Text            =   "2"
      Top             =   1140
      Width           =   4875
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
      Left            =   5880
      Style           =   1  'Graphical
      TabIndex        =   25
      Top             =   3720
      Visible         =   0   'False
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
      Left            =   9165
      Style           =   1  'Graphical
      TabIndex        =   19
      ToolTipText     =   "Exit"
      Top             =   3405
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
      Left            =   7965
      Style           =   1  'Graphical
      TabIndex        =   12
      ToolTipText     =   "To Cancel Record"
      Top             =   3405
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
      Left            =   7965
      Style           =   1  'Graphical
      TabIndex        =   11
      ToolTipText     =   "To Save Record"
      Top             =   2940
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
      Left            =   4800
      Style           =   1  'Graphical
      TabIndex        =   24
      Top             =   3720
      Visible         =   0   'False
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
      Left            =   7965
      Style           =   1  'Graphical
      TabIndex        =   10
      ToolTipText     =   "To Delete Record"
      Top             =   2340
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
      Left            =   9165
      Style           =   1  'Graphical
      TabIndex        =   9
      ToolTipText     =   "To Modify Record"
      Top             =   1860
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
      Left            =   7965
      Style           =   1  'Graphical
      TabIndex        =   8
      ToolTipText     =   "To Add New Record"
      Top             =   1860
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
      Left            =   9165
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   1260
      Visible         =   0   'False
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
      Left            =   7965
      Style           =   1  'Graphical
      TabIndex        =   13
      ToolTipText     =   "To Find Record"
      Top             =   1260
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
      Left            =   9765
      Style           =   1  'Graphical
      TabIndex        =   18
      ToolTipText     =   "Last"
      Top             =   780
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
      Left            =   9165
      Style           =   1  'Graphical
      TabIndex        =   17
      ToolTipText     =   "First"
      Top             =   780
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
      Left            =   8565
      Style           =   1  'Graphical
      TabIndex        =   16
      ToolTipText     =   "Next"
      Top             =   780
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
      Left            =   7965
      Style           =   1  'Graphical
      TabIndex        =   15
      ToolTipText     =   "Previous"
      Top             =   780
      Width           =   495
   End
   Begin VB.Shape Shape4 
      BorderColor     =   &H00C00000&
      BorderWidth     =   5
      FillColor       =   &H00C00000&
      Height          =   4140
      Left            =   0
      Top             =   330
      Width           =   10695
   End
   Begin VB.Shape Shape3 
      BorderColor     =   &H00C00000&
      FillColor       =   &H00C00000&
      FillStyle       =   0  'Solid
      Height          =   135
      Left            =   0
      Top             =   4200
      Width           =   10695
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
      Height          =   495
      Left            =   7935
      TabIndex        =   29
      Top             =   60
      Width           =   2475
   End
   Begin VB.Label lbl1 
      BackStyle       =   0  'Transparent
      Caption         =   "Tax D&escription   :"
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
      Height          =   375
      Index           =   0
      Left            =   480
      TabIndex        =   0
      Top             =   1140
      Width           =   2055
   End
   Begin VB.Label lbl1 
      BackStyle       =   0  'Transparent
      Caption         =   "Sale Tax A/c        :"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   375
      Index           =   6
      Left            =   480
      TabIndex        =   6
      Top             =   2595
      Width           =   1815
   End
   Begin VB.Label lbl1 
      BackStyle       =   0  'Transparent
      Caption         =   "&Purchase Tax A/c :"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   495
      Index           =   5
      Left            =   480
      TabIndex        =   4
      Top             =   2100
      Width           =   1695
   End
   Begin VB.Label lbl1 
      BackStyle       =   0  'Transparent
      Caption         =   "Tax %                  :"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   375
      Index           =   3
      Left            =   480
      TabIndex        =   2
      Top             =   1620
      Width           =   1935
   End
   Begin VB.Label Label6 
      BackColor       =   &H00C00000&
      Caption         =   "   Tax Master"
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
      Height          =   405
      Left            =   0
      TabIndex        =   26
      Top             =   0
      Width           =   10695
   End
   Begin VB.Shape Shape1 
      FillColor       =   &H00BDDFD1&
      FillStyle       =   0  'Solid
      Height          =   3420
      Left            =   7710
      Shape           =   4  'Rounded Rectangle
      Top             =   570
      Width           =   2775
   End
   Begin VB.Shape Shape2 
      BackColor       =   &H00808080&
      BorderColor     =   &H004D177D&
      FillColor       =   &H00ECFFFF&
      FillStyle       =   0  'Solid
      Height          =   3315
      Left            =   240
      Top             =   660
      Width           =   7275
   End
End
Attribute VB_Name = "frmMastNarrationTax"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Const tblName As String = "tblMastNarration"
Const IdField As String = "narrcode"
Const OrderField As String = "Narration"
Const MaxNo As Long = 9
'Const CriteriaStr As String = " narrtp = '" & gfrmTypeStr & "'"
Dim Rstbl As Recordset
'Dim varBookMark As Variant
Dim FirstTimeFlag As Boolean
Dim FormAction As Integer

Private Sub cmdBtn_Click(Index As Integer)
Dim FindQStr As String
Dim X As String
On Error GoTo ErrorRoutine
Select Case Index
Case 0 'Previous
    Rstbl.MovePrevious
    If Rstbl.BOF = True Then
        Rstbl.MoveFirst
        MsgBox ("This is First Record"), vbInformation + vbOKOnly, "First Record"
    End If
    ReadFields
Case 1 'Next
    Rstbl.MoveNext
    If Rstbl.EOF = True Then
        Rstbl.MoveLast
        MsgBox ("This is Last Record"), vbInformation + vbOKOnly, "Last Record"
    End If
    ReadFields
Case 2 'First
    Rstbl.MoveFirst
    ReadFields
Case 3 'Last
    Rstbl.MoveLast
    ReadFields
Case 4 'Find
    If gBackEndDB = gBackEndAccess Then
        X = " TblMastAccount as TblMastac1 "
    Else '-- Oracle
        X = " TblMastAccount TblMastac1 "
    End If
    FindQStr = "Select Narration,MastTaxRate,tblMastAccount.AcName,tblMastAc1.AcName from " _
    & " tblMastNarration,TblMastaccount," & X & " where tblMastNarration.MastTaxPurAcCode=TblMastaccount.AcCode" _
    & " and tblMastNarration.MastTaxSAAcCode=TblMastac1.AcCode and NarrType='T'"
    gClsSearch.SearchMultiField FindQStr, "", Array("Tax", "Rate", "Purchase Account", "Sales Account"), Array(2000, 1000, 2000, 2000), "", "", "Narration", txt(2).Left + Me.Left, txt(2).Top + Me.Top + 650, False, , True
    If gClsSearch.SearchMultiRetCol(0) <> "" Then
        Rstbl.MoveFirst
        Rstbl.Find "Narration= '" & gClsSearch.SearchMultiRetCol(0) & "'"
        gClsSearch.SearchMultiRetCol(0) = ""
        ReadFields
    End If
Case 5 'List
    'GenMastList ("LAC")
Case 6 'Add
    FormAction = vbDataActionAddNew
    Call GProcActivateControls(Me, True)
    Call GProcSetButtons(Me, vbDataActionAddNew)
    Call GProcClearForm(Me, Rstbl, MaxNo, True)
    If FirstTimeFlag = True Then
        FirstTimeFlag = False
    Else
    txt(2).SetFocus
    End If
    FillCombo

Case 7 'Modify
    'If mURecModi = True Then
    '    SearchHelp
    '    DoEvents
        cmdBtn_Click 4 '-- Find
        If gClsSearch.SearchMultiRetCol(0) = "" Then Exit Sub
        FormAction = vbDataActionUpdate
        Call GProcActivateControls(Me, True)
        Call GProcSetButtons(Me, vbDataActionUpdate)
        txt(2).SetFocus
    'End If
Case 8 'Delete
    'If mURecDel = True Then
        If MsgBox("Do You Really Delete This Record?", vbQuestion + vbYesNo + vbDefaultButton1, "Confirm Delete") = vbYes Then
            If Rstbl.RecordCount > 1 Then
                Rstbl.Delete
                cmdBtn_Click (1)
            ElseIf Rstbl.RecordCount = 1 Then
                Rstbl.Delete
                Call GProcActivateControls(Me, False)
                Call GProcSetButtons(Me, vbDataActionCancel)
                Form_Load
            End If
        End If
    'End If
Case 9 'Print
Case 10 'Exit
    Unload Me
Case 11 'Label
    'frmLabel.Show 1
Case 12 'Creation To Firm
     gfrmTypeStr = "Tax"
     frmUtlCreationToFirm.Show 1
     Rstbl.Requery
Case 13 'Save
    If ValidateData = True Then
        FillTxtFromTemp
        GProcSaveRecord Me, Rstbl, FormAction, MaxNo
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
    Call GProcClearForm(Me, Rstbl, MaxNo, True)
    If Rstbl.RecordCount <= 0 Then
        Call GProcActivateControls(Me, False)
        Call GProcSetButtons(Me, vbDataActionCancel)
    Else
        If FormAction = vbDataActionAddNew Then
            Rstbl.MoveLast
        Else
            Rstbl.CancelUpdate
        End If
        ReadFields
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
    Rstbl.CancelUpdate
    GProcErrorHandler
End Sub
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case 13:
            SendKeys "{TAB}"
        Case 27:
            If (cmdBtn(13).Enabled = False) Then
                Unload Me
            End If
    End Select
End Sub
Private Sub Form_Load()
    Dim X As String
    Me.Top = 300
    Me.Left = Screen.Width / 2 - Me.Width / 2
    X = "Select * from " & tblName & " where NarrType='T'"
    If Trim(OrderField) <> "" Then X = X & " order by " & OrderField
    '---- Button Visiblity Set As Per User Previlage
    GProcSetButtonVisiblity Me
    GProcRstOpen Rstbl, X, "O"
    FirstTimeFlag = True
    cmdBtn_Click (6)
End Sub
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
If FormAction = vbDataActionUpdate Then
    MsgBox "You are about to modify the record. Save the record and then exit.", vbInformation + vbOKOnly, Me.Caption
    Cancel = True
    Exit Sub
End If
'mShowRec = Space(5)
Rstbl.Requery
Rstbl.Close
End Sub

Private Sub txt_GotFocus(Index As Integer)
GProcSelectBox Me.ActiveControl
End Sub
Private Sub txt_LostFocus(Index As Integer)
GProcCheckForNumber Me, Rstbl, MaxNo
SaveBtnEd
End Sub
Private Sub txt_KeyPress(Index As Integer, KeyAscii As Integer)
If KeyAscii = 13 Then Exit Sub
KeyAscii = GProcValidateKey(Rstbl, Index, KeyAscii, txt(Index).Text)
End Sub

Private Sub FillCombo()
'Call GProFillCombo(cbo(7), "select Distinct AcCity from " & tblName & " where AcCity <> '' order by AcCity")
End Sub
Private Sub SaveBtnEd()
'--- Narration
If txt(2) = "" Then
    cmdBtn(13).Enabled = False
    cmdBtn(13).BackColor = &H80000004
Else
    cmdBtn(13).Enabled = True
    cmdBtn(13).BackColor = &HFFFF80
End If
End Sub
Private Function ValidateData() As Boolean
Dim i As Long
Dim X As String
Dim Y As String
Dim rstAddCheck As Recordset
Dim rstModiCheck As Recordset
    GProcCheckForNumber Me, Rstbl, MaxNo
    '---- Required
    '-- NarrCode
    If FormAction = vbDataActionAddNew Then txt(1) = GProcGenerateId(tblName, IdField, pStartCode:=gCSNarrCode)
    If gCENarrCode > 0 And CLng(txt(1)) > gCENarrCode Then
        MsgBox "Year End is completed. You can not create new narration.", vbCritical + vbOKOnly, Me.Caption
        txt(2).SetFocus
        Exit Function
    End If
    '--- Tax Account Purchase
    If txtTemp(8) = "" Then txtTemp(8) = "-"
    If GProcGetColumnValue("tblMastAccount", "AcName", txtTemp(8), "S", "AcCode", "N") = 0 Then
        'GName = GProcGetColumnValue("TblMastGroup", "AgCode", 25, "N", "AgName", "S")
        GProcShowForm frmMastAccount, frmMain.mnuMstfrm(0), Array(frmMastAccount.txt(1)), Array(txtTemp(8))
        Exit Function
    End If
    '--- Tax Account Sale
    If txtTemp(9) = "" Then txtTemp(9) = "-"
    If GProcGetColumnValue("tblMastAccount", "AcName", txtTemp(9), "S", "AcCode", "N") = 0 Then
        'GName = GProcGetColumnValue("TblMastGroup", "AgCode", 25, "N", "AgName", "S")
        GProcShowForm frmMastAccount, frmMain.mnuMstfrm(0), Array(frmMastAccount.txt(1)), Array(txtTemp(9))
        Exit Function
    End If
    '--- Narration
    If Trim(txt(2)) = "" Then
        MsgBox "Check Narration.", vbCritical + vbOKOnly, Me.Caption
        txt(2).SetFocus
        Exit Function
    End If
    '----Duplicate Records
    '----Addition Case
    If FormAction = vbDataActionAddNew Then
         X = "Select NArration From " & tblName & " Where Narration='" & Trim(txt(2).Text) & "'"
        i = GProcRstOpen(rstAddCheck, X, "R")
        If i > 0 Then
            MsgBox "Record is Available of Same Narration  " & Trim(txt(2).Text), vbCritical + vbOKOnly, "Duplicate Record"
            txt(2).SetFocus
            Exit Function
        End If
        rstAddCheck.Close
    End If
    '----Modification case
    If FormAction = vbDataActionUpdate Then
        X = "Select narration From " & tblName & " Where ((narration='" & Trim(txt(2).Text) & "') and (narrCode<>" & txt(1).Text & "))"
        i = GProcRstOpen(rstModiCheck, X, "O")
        If i > 0 Then
            MsgBox "Record is Available of Same Narration  " & Trim(txt(2).Text), vbCritical + vbOKOnly, "Duplicate Record"
            txt(2).SetFocus
            Exit Function
        End If
        rstModiCheck.Close
    End If
    ValidateData = True
End Function
Private Sub FillTxtFromTemp()
txt(0) = "T"
'-- Tax AcCode
txt(8) = GProcGetColumnValue("TblMastAccount", "ACName", txtTemp(8), "S", "AcCode", "N")
txt(9) = GProcGetColumnValue("TblMastAccount", "ACName", txtTemp(9), "S", "AcCode", "N")
End Sub
Private Sub FillTempFromTxt()
'-- Tax AcName
txtTemp(8) = GProcGetColumnValue("TblMastAccount", "AcCode", txt(8), "N", "ACName", "S")
txtTemp(9) = GProcGetColumnValue("TblMastAccount", "AcCode", txt(9), "N", "ACName", "S")
End Sub
Private Sub ReadFields()
GProcShowRecord Me, Rstbl, MaxNo
FillTempFromTxt
End Sub

Private Sub txttemp_GotFocus(Index As Integer)
GProcSelectBox Me.ActiveControl
End Sub

Private Sub txttemp_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
Dim GName As String
If KeyCode = vbKeyF3 Then
    Select Case Index
        Case 8, 9 '-- Party
            'GName = GProcGetColumnValue("TblMastAccount", "AgCode", 15, "N", "AgName", "S")
            GProcShowForm frmMastAccount, frmMain.mnuMstfrm(0)
    End Select
End If
End Sub

Private Sub txtTemp_KeyPress(Index As Integer, KeyAscii As Integer)
Dim LeftPos As Long
Dim TopPos As Long
TopPos = txtTemp(Index).Top + Me.Top + 650
LeftPos = txtTemp(Index).Left + Me.Left
If KeyAscii = 13 Then Exit Sub
Select Case Index
    Case 8, 9 '-- Tax Account ,Alias Name
        gClsSearch.SearchMultiField "QryMastaccount", "Account,AcName,AgName", Array("Party", "AcName", "Group"), Array(txtTemp(Index).Width, 0, 3000), " (Agcode=90044 )", Chr(KeyAscii), "acName", LeftPos, TopPos, True
        KeyAscii = 0
        If gClsSearch.SearchMultiRetCol(0) <> "" And gClsSearch.SearchMultiRetCol(1) = "" Then
            txtTemp(Index) = gClsSearch.SearchMultiRetCol(0)
        Else
           If gClsSearch.SearchMultiRetCol(1) <> "" Then txtTemp(Index) = gClsSearch.SearchMultiRetCol(1)
        End If
End Select
End Sub
