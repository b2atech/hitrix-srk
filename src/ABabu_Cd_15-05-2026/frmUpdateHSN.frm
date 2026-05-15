VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Begin VB.Form frmUpdateHSN 
   BackColor       =   &H00C1E3FD&
   BorderStyle     =   0  'None
   Caption         =   " "
   ClientHeight    =   9510
   ClientLeft      =   0
   ClientTop       =   60
   ClientWidth     =   15360
   DrawMode        =   10  'Mask Pen
   FillColor       =   &H000040C0&
   FillStyle       =   0  'Solid
   Icon            =   "frmUpdateHSN.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   9510
   ScaleWidth      =   15360
   ShowInTaskbar   =   0   'False
   Begin VB.CheckBox Check1 
      BackColor       =   &H00BDDFCA&
      Caption         =   "Not Used"
      Height          =   375
      Left            =   17775
      TabIndex        =   29
      Top             =   5850
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.CommandButton cmdBtn 
      Caption         =   "&Find Item"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Index           =   12
      Left            =   17460
      Style           =   1  'Graphical
      TabIndex        =   28
      ToolTipText     =   "To Come out"
      Top             =   1665
      Width           =   1125
   End
   Begin VB.CommandButton cmdBtn 
      Caption         =   "&Print"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Index           =   11
      Left            =   14055
      Style           =   1  'Graphical
      TabIndex        =   27
      ToolTipText     =   "To Modify Current  Record"
      Top             =   2610
      Width           =   945
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
      Index           =   13
      Left            =   6735
      TabIndex        =   26
      Text            =   "13"
      Top             =   10410
      Visible         =   0   'False
      Width           =   465
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
      Index           =   12
      Left            =   7290
      TabIndex        =   25
      Text            =   "12"
      Top             =   10410
      Visible         =   0   'False
      Width           =   465
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
      Index           =   11
      Left            =   6165
      TabIndex        =   24
      Text            =   "11"
      Top             =   10410
      Visible         =   0   'False
      Width           =   465
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
      Index           =   10
      Left            =   6000
      TabIndex        =   23
      Text            =   "10"
      Top             =   10560
      Visible         =   0   'False
      Width           =   465
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
      Index           =   4
      Left            =   4485
      TabIndex        =   22
      Text            =   "4"
      Top             =   10410
      Visible         =   0   'False
      Width           =   465
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
      Index           =   3
      Left            =   4965
      TabIndex        =   21
      Text            =   "3"
      Top             =   10410
      Visible         =   0   'False
      Width           =   465
   End
   Begin VB.CommandButton cmdBtn 
      Caption         =   "&Print"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Index           =   9
      Left            =   17460
      Style           =   1  'Graphical
      TabIndex        =   20
      ToolTipText     =   "To Add New Record"
      Top             =   2745
      Visible         =   0   'False
      Width           =   1125
   End
   Begin VB.CommandButton cmdBtn 
      Caption         =   "&List"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Index           =   5
      Left            =   17550
      Style           =   1  'Graphical
      TabIndex        =   19
      ToolTipText     =   "To List of  Records"
      Top             =   2295
      Width           =   1125
   End
   Begin VB.CommandButton cmdBtn 
      Caption         =   "<"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Index           =   0
      Left            =   17730
      Style           =   1  'Graphical
      TabIndex        =   18
      ToolTipText     =   "To Move Previous  Record"
      Top             =   4815
      Visible         =   0   'False
      Width           =   450
   End
   Begin VB.CommandButton cmdBtn 
      Caption         =   ">"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Index           =   1
      Left            =   18180
      Style           =   1  'Graphical
      TabIndex        =   17
      ToolTipText     =   "To Move Next  Record"
      Top             =   4770
      Visible         =   0   'False
      Width           =   450
   End
   Begin VB.CommandButton cmdBtn 
      Caption         =   "<<"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Index           =   2
      Left            =   17595
      Style           =   1  'Graphical
      TabIndex        =   16
      ToolTipText     =   "To Move First  Record"
      Top             =   5310
      Visible         =   0   'False
      Width           =   450
   End
   Begin VB.CommandButton cmdBtn 
      Caption         =   ">>"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Index           =   3
      Left            =   18225
      Style           =   1  'Graphical
      TabIndex        =   15
      ToolTipText     =   "To Move Last  Record"
      Top             =   5310
      Visible         =   0   'False
      Width           =   450
   End
   Begin VB.CommandButton cmdBtn 
      Caption         =   "&Mill Sr."
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Index           =   4
      Left            =   14055
      Style           =   1  'Graphical
      TabIndex        =   14
      ToolTipText     =   "To Find Record"
      Top             =   5400
      Width           =   945
   End
   Begin VB.CommandButton cmdBtn 
      Caption         =   "&Add"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Index           =   6
      Left            =   17505
      Style           =   1  'Graphical
      TabIndex        =   13
      ToolTipText     =   "To Add New Record"
      Top             =   3780
      Visible         =   0   'False
      Width           =   1125
   End
   Begin VB.CommandButton cmdBtn 
      Caption         =   "&Modify"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Index           =   7
      Left            =   14055
      Style           =   1  'Graphical
      TabIndex        =   12
      ToolTipText     =   "To Modify Current  Record"
      Top             =   1800
      Width           =   945
   End
   Begin VB.CommandButton cmdBtn 
      Caption         =   "&Delete"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Index           =   8
      Left            =   17550
      Style           =   1  'Graphical
      TabIndex        =   11
      ToolTipText     =   "To Delete Current Record"
      Top             =   3285
      Visible         =   0   'False
      Width           =   1125
   End
   Begin VB.CommandButton cmdBtn 
      Caption         =   "&Save"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Index           =   13
      Left            =   14055
      Style           =   1  'Graphical
      TabIndex        =   10
      ToolTipText     =   "To Save Current  Record"
      Top             =   3510
      Width           =   945
   End
   Begin VB.CommandButton cmdBtn 
      Caption         =   "&Cancel"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Index           =   14
      Left            =   14055
      Style           =   1  'Graphical
      TabIndex        =   9
      ToolTipText     =   "To undo Changes"
      Top             =   4545
      Width           =   945
   End
   Begin VB.CommandButton cmdBtn 
      Caption         =   "E&xit"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Index           =   10
      Left            =   14055
      Style           =   1  'Graphical
      TabIndex        =   8
      ToolTipText     =   "To Come out"
      Top             =   6255
      Width           =   945
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
      Index           =   1
      Left            =   2340
      TabIndex        =   2
      Text            =   "1"
      Top             =   10410
      Visible         =   0   'False
      Width           =   615
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
      Left            =   3960
      TabIndex        =   6
      Text            =   "7"
      Top             =   10410
      Visible         =   0   'False
      Width           =   465
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
      Index           =   5
      Left            =   2880
      TabIndex        =   5
      Text            =   "5"
      Top             =   10410
      Visible         =   0   'False
      Width           =   465
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
      Left            =   2805
      TabIndex        =   1
      Text            =   "2"
      Top             =   10410
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.TextBox txtGrid 
      Height          =   405
      Left            =   15705
      TabIndex        =   7
      Text            =   " "
      Top             =   9000
      Visible         =   0   'False
      Width           =   750
   End
   Begin MSFlexGridLib.MSFlexGrid msGrid 
      Height          =   8100
      Left            =   540
      TabIndex        =   0
      Top             =   945
      Width           =   13200
      _ExtentX        =   23283
      _ExtentY        =   14288
      _Version        =   393216
      FixedCols       =   0
      RowHeightMin    =   275
      BackColorFixed  =   -2147483647
      ForeColorFixed  =   -2147483634
      BackColorBkg    =   12632256
      AllowUserResizing=   1
      Appearance      =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Shape Shape3 
      BorderColor     =   &H00EFB17A&
      BorderWidth     =   5
      FillColor       =   &H00CC004F&
      Height          =   9435
      Left            =   0
      Top             =   45
      Width           =   15315
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
      Left            =   10005
      TabIndex        =   4
      Top             =   90
      Width           =   2475
   End
   Begin VB.Label Label6 
      BackColor       =   &H00EFB17A&
      Caption         =   "              Price List"
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
      Height          =   345
      Left            =   0
      TabIndex        =   3
      Top             =   75
      Width           =   15270
   End
   Begin VB.Shape Shape1 
      FillColor       =   &H00BDDFCA&
      FillStyle       =   0  'Solid
      Height          =   6105
      Left            =   13905
      Shape           =   4  'Rounded Rectangle
      Top             =   1140
      Width           =   1215
   End
   Begin VB.Shape Shape2 
      BackColor       =   &H00808080&
      BorderColor     =   &H00000000&
      FillColor       =   &H00ECFFFF&
      FillStyle       =   0  'Solid
      Height          =   8805
      Left            =   165
      Top             =   495
      Width           =   15150
   End
End
Attribute VB_Name = "frmUpdateHSN"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Const tblName As String = "tblMastItem"
Const IdField As String = "ItCode"
Const OrderField As String = "ItName"
Const MaxNo As Long = 11
Dim rstbl As ADODB.Recordset
Dim rstblSub As ADODB.Recordset
Dim X As String
Dim FormAction As Integer
Dim FirstTimeFlag As Boolean
Private Sub SetGrid()
    X = "select AcOurCode,ItName,ItHsn2,ItCode" _
         & " FROM tblMastItem,tblMastAccount where tblMastItem.ItMillCode = tblMastAccount.AcCode  and itname <> '-'  order by AcOurCode,ItName,ItHsn2"
    GProcRstOpen rstblSub, X, "O"
     With msGrid
        .Clear          ' 0                                              1                                  2                         3
       .FormatString = "<Mill Code         |<Item Name                                                 |<HSN Codes               |ItCode"
        .ColWidth(3) = 0
        .rows = 2
     End With
    GridStop True
End Sub
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
Case 7 'Modify
        FormAction = vbDataActionUpdate
        Call GProcActivateControls(Me, True)
        Call GProcSetButtons(Me, vbDataActionUpdate)
        FillRelatedRecord
        msGrid_EnterCell
Case 10 'Exit
    Unload Me
Case 11 'Print
'    With frmMain.cryReport
'      GProcCrystalRptPreparation frmMain.cryReport, "Window"
'        .Formulas(5) = "wRptName = 'Price List Report'"
'        X = "ucase({TblpriCeList.PrYn})='Y' and {TblpriCeList.Entno}=" & Val(txt(0))
'        .SelectionFormula = X
'        .ReportFileName = gReportPath & "rptPriceList.rpt"
'        .Action = 1
'     End With
Case 13 'Save
'    If ValidateData = True Then
'        FillTxtFromTemp
        WriteFields
        Call GProcActivateControls(Me, False)
        Call GProcSetButtons(Me, True)
        If FormAction = vbDataActionAddNew Then
            If cmdBtn(6).Visible = True Then cmdBtn(6).SetFocus
        Else
            If cmdBtn(7).Visible = True Then cmdBtn(7).SetFocus
        End If
        FormAction = vbDataActionClose
 '   End If
Case 14 'Cancel
    Call GProcClearForm(Me, rstbl, MaxNo, True)
    Call GProcActivateControls(Me, False)
    Call GProcSetButtons(Me, True)
    If FormAction = vbDataActionAddNew Then
        If cmdBtn(6).Visible = True Then cmdBtn(6).SetFocus
    Else
        If cmdBtn(7).Visible = True Then cmdBtn(7).SetFocus
    End If
    FormAction = vbDataActionClose
End Select
Exit Sub
ErrorRoutine:
    GProcErrorHandler
End Sub
Private Sub Form_Load()
    Me.Top = 300
    Me.Left = Screen.Width / 2 - Me.Width / 2
    X = "Select * from " & tblName & " where ItType = 0 and itname <> '-' "
    If Trim(OrderField) <> "" Then X = X & " order by " & OrderField
    GProcRstOpen rstbl, X, "O"
    GProcSetButtonVisiblity Me
    FirstTimeFlag = True
   cmdBtn_Click (7)
End Sub
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
If FormAction = vbDataActionUpdate Then
    MsgBox "You are about to modify the record. Save the record and then exit.", vbInformation + vbOKOnly, Me.Caption
    Cancel = True
    Exit Sub
End If
rstbl.Close
rstblSub.Close
Set frmPriceList = Nothing
End Sub
Private Sub msGrid_EnterCell()
On Error Resume Next
    With msGrid
If .Row <> 0 Then
        Select Case .Col
                Case 0
                   .Col = 2
                Case 1, 2, 3, 4, 5, 6, 7
                    .CellForeColor = vbWhite
                    .CellBackColor = vbBlue
        End Select
End If
    End With
End Sub
Private Sub msGrid_GotFocus()
  cmdBtn(10).Cancel = False
  GridStop False
   msGrid_EnterCell
End Sub
Private Sub msGrid_KeyDown(KeyCode As Integer, Shift As Integer)
 With msGrid
    If KeyCode = vbKeyDelete And .Col > 1 Then msGrid.text = ""
    
    If KeyCode = vbKeyF9 Then
'        If (Trim(.TextMatrix(.Row, 1)) = "") And .rows > 2 Then
 '               .RemoveItem (.Row)
  '               DoEvents
   '    End If
       GridStop False
       cmdBtn(13).SetFocus
       Exit Sub
    End If
 End With
' If (KeyCode = vbKeyY Or KeyCode = 121) And Shift = 2 And msGrid.Col = 0 Then
'    If msGrid.rows = 2 Then
'      Dim Ir As Integer
'      For Ir = 0 To msGrid.cols - 1
'            msGrid.TextMatrix(1, Ir) = ""
'      Next
'    Else
'      msGrid.RemoveItem msGrid.Row
'   End If
'End If
'If KeyCode = vbKeyDelete Then msGrid.text = ""
    With msGrid
                If KeyCode = vbKeyReturn Then
                        If .Col < 2 Then
                              .Col = .Col + 1
                        ElseIf .Col = 2 Then
                                If .rows = .Row + 1 Then
                                    .rows = .rows + 1
                                   .Row = .rows - 1
                                Else
                                    .Row = .Row + 1
                                End If
                                If .Row - .TopRow > 16 Then .TopRow = .TopRow + 1
                                .Col = 2
                        End If
                   .SetFocus
                End If
    End With
End Sub
Private Sub msGrid_KeyPress(KeyAscii As Integer)
Dim LeftPos As Long, aa As String
Dim TopPos As Long
TopPos = msGrid.Top + Me.Top + 650
LeftPos = msGrid.Left + Me.Left
    DoEvents
    With msGrid
        If KeyAscii <> 13 And KeyAscii <> 27 And KeyAscii <> 9 And KeyAscii <> 8 And KeyAscii > 31 Then
            '--- Search
            Select Case .Col
                Case 1 '-- Quality
'                        gClsSearch.SearchMultiField "tblMastTrade", "TrName,TrCode", Array("Trade Name", "code"), Array(5000, 0), "Folder=" & Val(txt(10)), Trim(Chr(KeyAscii)), "TrName", LeftPos, TopPos
'                        KeyAscii = 0
'                        If gClsSearch.SearchMultiRetCol(0) <> "" Then
'                            aa = gClsSearch.SearchMultiRetCol(0)
'                            For i = 1 To .rows - 1
'                              If .TextMatrix(i, 1) = aa And i <> .Row Then
'                                 MsgBox "Duplicate Item "
'                                 Exit Sub
'                              End If
'                            Next i
'                            .text = aa
'                            .TextMatrix(.Row, 4) = gClsSearch.SearchMultiRetCol(1)
'                            If Val(.TextMatrix(.Row, 2)) = 0 Then .TextMatrix(.Row, 2) = GProcGetColumnValue("TblPriceList", "Qcode", Val(.TextMatrix(.Row, 4)), "N", "Rate", "N", "frdate < '" & Format(DTPicker(0), "yyyy/MM/dd") & "'")
'                            If .TextMatrix(.Row, 3) = "" Then .TextMatrix(.Row, 3) = "Y"
'                            .Col = 2
'                        End If
                Case 3
                Case 5
                Case Else
                    KeyAscii = GProcValidateKey(rstbl, msGrid.Col, KeyAscii, msGrid.text)
            End Select
       End If
        If KeyAscii = 13 Then
            CheckForNumbermsGrid
'            Select Case .Col
'            Case 0, 1
'            Case Else
'                If Val(.text) = "0" Then .text = ""
'            End Select
        End If
        If KeyAscii = 8 Then
           If Len(Trim(.text)) > 0 Then .text = Mid(.text, 1, Len(Trim(.text)) - 1)
        End If
        If KeyAscii <> 13 And KeyAscii <> 27 And KeyAscii <> 9 And KeyAscii <> 8 And KeyAscii <> 25 Then
            msGrid.text = msGrid.text & Chr(KeyAscii)
        End If
            If KeyAscii = 27 Then '----Esc Key
              DoEvents
              DoEvents
              KeyAscii = 0
'             If (Trim(.TextMatrix(.Row, 1)) = "") And .rows > 2 Then
'                    .RemoveItem (.Row)
'                  DoEvents
'              End If
                 GridStop False
                 cmdBtn(13).SetFocus
                 Exit Sub
            End If
        End With
End Sub
Private Sub msGrid_LeaveCell()
With msGrid
   If .Row <> 0 Then
        CheckForNumbermsGrid
        .CellForeColor = vbBlack
        .CellBackColor = vbWhite
   End If
  End With
End Sub
Private Sub MSGrid_LostFocus()
    GridStop True
End Sub
Private Sub txt_GotFocus(Index As Integer)
GProcSelectBox Me.ActiveControl
End Sub
'---  Check For Number Grid
Private Sub CheckForNumbermsGrid()
With rstblSub
    Select Case .Fields(msGrid.Col).Type
        Case adInteger, 2, 3, 20 '-- Integer
            If IsNumeric(msGrid.text) = False Then
                msGrid = 0
            Else
                msGrid = CLng(msGrid.text)
            End If
        Case adDecimal, 131, adNumeric, adDouble '-- Double
            If IsNumeric(msGrid) = False Then
                'msGrid = Format(0, GProcNumberFormat(.Fields(msGrid.Col).Precision, .Fields(msGrid.Col).NumericScale))
            Else
                msGrid = Format(msGrid, GProcNumberFormat(.Fields(msGrid.Col).Precision, .Fields(msGrid.Col).NumericScale))
            End If
        Case adChar, adVarWChar, 200, 130
            If IsMissing(msGrid) = False Then
                msGrid = Trim(msGrid)
            Else
                msGrid = ""
            End If
        Case adDate '-- Date
            If IsDate(msGrid) = False Then
                msGrid = ""
            Else
                msGrid = CDate(msGrid)
            End If
         Case Else
    End Select
End With
End Sub

Private Sub txt_KeyPress(Index As Integer, KeyAscii As Integer)
If KeyAscii = 13 Then Exit Sub
KeyAscii = GProcValidateKey(rstbl, Index, KeyAscii, txt(Index).text)
End Sub

Private Sub txt_LostFocus(Index As Integer)
GProcCheckForNumber Me, rstbl, MaxNo
SaveBtnEd
End Sub
Private Sub SaveBtnEd()
'--- Narration
'If txttemp(5) = "" Then
'    cmdBtn(13).Enabled = False
'    cmdBtn(13).BackColor = &H80000004
'Else
'    cmdBtn(13).Enabled = True
'    cmdBtn(13).BackColor = &HFFFF80
'End If
End Sub
Private Function ValidateData() As Boolean
End Function
Private Function isDeleteValidate() As Boolean
  isDeleteValidate = True
End Function
Private Sub FillTempFromTxt()
'Check1 = Val(txt(11))
'DTPicker(0) = Format(txt(1), "yyyy/mm/dd")
'txttemp(5) = GProcGetColumnValue("TblMastAccount", "AcCode", txt(5), "N", "AcName+', '+Accity", "S")
'txttemp(10) = GProcGetColumnValue("TblMastNarration", "NarrCode", txt(10), "N", "Narration", "S")
FillRelatedRecord
End Sub
Private Sub FillTxtFromTemp()
'txt(6) = Trim(gUserName)
'txt(11) = Check1.Value
'txt(1) = Format(DTPicker(0), "YYYY/MM/dd")
End Sub
Private Sub FillRelatedRecord()
Dim RowIndex As Long
Dim ColIndex As Long
'--Fill Grid
SetGrid
With rstblSub
    If Not .EOF Then
        .MoveFirst
        RowIndex = 1
        msGrid.rows = 2
        Do While Not .EOF
            For ColIndex = 0 To .Fields.Count - 1
                Select Case .Fields(ColIndex).Type
                    Case adInteger, 2, 3, 20 '-- Integer
                        msGrid.TextMatrix(RowIndex, ColIndex) = IIf(IsNull(Val(.Fields(ColIndex))), 0, Val(.Fields(ColIndex)))
                    Case adDecimal, 131, adNumeric, adDouble  '-- Double
                        msGrid.TextMatrix(RowIndex, ColIndex) = Format(.Fields(ColIndex), GProcNumberFormat(.Fields(ColIndex).Precision, .Fields(ColIndex).NumericScale))
                    Case Else
                        msGrid.TextMatrix(RowIndex, ColIndex) = .Fields(ColIndex) & vbNullString
                End Select
            Next ColIndex
        '    If GProcRstOpen(Tmp, "Select  Rate from TblPriceList where Qcode=" & Val(msGrid.TextMatrix(RowIndex, 2)) & " and  frdate < '" & Format(DTPicker(0), "yyyy/MM/dd") & "'  order by frdate desc", "R") > 0 Then
        '       msGrid.TextMatrix(RowIndex, 6) = Tmp.Fields(0)
        '    End If
                        
            RowIndex = RowIndex + 1
            .MoveNext
            If Not .EOF Then msGrid.rows = msGrid.rows + 1
        Loop
    End If
End With
End Sub
Private Sub ReadFields()
    GProcShowRecord Me, rstbl, MaxNo
    FillTempFromTxt
End Sub
Private Sub GridStop(mTrFl As Boolean)
'For i = 0 To 9
'txt(i).TabStop = mTrFl
'Next
For i = 0 To 14
 cmdBtn(i).TabStop = mTrFl
Next
'txttemp(5).TabStop = mTrFl
'txttemp(10).TabStop = mTrFl
'DTPicker(0).TabStop = mTrFl
'Check1.TabStop = mTrFl
End Sub
Private Sub WriteFields()
Dim mRow As Integer
'--- Save Grid Records
'CheckForNumbermsGrid
'gCn.Execute "update tblMastItem Set ItemRate = 0 , BookingQty = 0  where ItType = 0 "
'If FormAction = vbDataActionAddNew Then txt(0) = GProcGenerateId("tblpricelist", "entno")
With msGrid
    mRow = 1
    Do While mRow < .rows
'        If Val(.TextMatrix(mRow, 3)) + Val(.TextMatrix(mRow, 4)) > 0 Then
        DoEvents
            gCn.Execute "update tblMastItem Set ItHsn2 = '" & (.TextMatrix(mRow, 2)) & "'  where ItCode = " & Val(.TextMatrix(mRow, 3)) & " "
        DoEvents
 '       End If
        mRow = mRow + 1
     Loop
End With
End Sub

