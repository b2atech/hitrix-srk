VERSION 5.00
Begin VB.Form frmItemMast 
   BackColor       =   &H00C2E3FC&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "&H00FF80FF&"
   ClientHeight    =   7590
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   10455
   FillColor       =   &H000040C0&
   FillStyle       =   0  'Solid
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7590
   ScaleWidth      =   10455
   Begin VB.TextBox txt 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Index           =   7
      Left            =   1080
      TabIndex        =   38
      Text            =   "7"
      Top             =   6000
      Visible         =   0   'False
      Width           =   315
   End
   Begin VB.TextBox txt 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Index           =   6
      Left            =   1560
      TabIndex        =   37
      Text            =   "6"
      Top             =   6000
      Visible         =   0   'False
      Width           =   315
   End
   Begin VB.TextBox temptxt 
      Height          =   375
      Index           =   9
      Left            =   2400
      TabIndex        =   2
      Text            =   "9"
      Top             =   2280
      Width           =   5415
   End
   Begin VB.ComboBox Cbo 
      Height          =   315
      Index           =   7
      Left            =   2400
      TabIndex        =   8
      Text            =   "Cbo(7)"
      Top             =   5040
      Width           =   1455
   End
   Begin VB.ComboBox Cbo 
      Height          =   315
      Index           =   6
      Left            =   6360
      TabIndex        =   7
      Text            =   "Cbo(6)"
      Top             =   4200
      Width           =   1575
   End
   Begin VB.TextBox txt 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Index           =   8
      Left            =   6360
      TabIndex        =   9
      Text            =   "8"
      Top             =   5040
      Width           =   1635
   End
   Begin VB.TextBox txt 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Index           =   9
      Left            =   600
      TabIndex        =   34
      Text            =   "9"
      Top             =   6000
      Visible         =   0   'False
      Width           =   315
   End
   Begin VB.TextBox txt 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Index           =   5
      Left            =   2400
      TabIndex        =   6
      Text            =   "5"
      Top             =   4200
      Width           =   1515
   End
   Begin VB.TextBox txt 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Index           =   4
      Left            =   2400
      TabIndex        =   5
      Text            =   "4"
      Top             =   3480
      Width           =   1515
   End
   Begin VB.TextBox txt 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Index           =   3
      Left            =   6360
      TabIndex        =   4
      Text            =   "3"
      Top             =   2880
      Width           =   1515
   End
   Begin VB.TextBox txt 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Index           =   2
      Left            =   2400
      TabIndex        =   3
      Text            =   "2"
      Top             =   2880
      Width           =   1515
   End
   Begin VB.TextBox txt 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Index           =   1
      Left            =   2400
      TabIndex        =   1
      Text            =   "1"
      Top             =   1560
      Width           =   5475
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
      Left            =   8880
      Style           =   1  'Graphical
      TabIndex        =   24
      ToolTipText     =   "Optional"
      Top             =   6720
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
      Left            =   8880
      Style           =   1  'Graphical
      TabIndex        =   23
      ToolTipText     =   "Label"
      Top             =   6240
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
      Left            =   8880
      Style           =   1  'Graphical
      TabIndex        =   22
      ToolTipText     =   "Exit from form"
      Top             =   5760
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
      Left            =   8880
      Style           =   1  'Graphical
      TabIndex        =   21
      ToolTipText     =   "cancel"
      Top             =   5280
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
      Left            =   8880
      Style           =   1  'Graphical
      TabIndex        =   20
      ToolTipText     =   "Save Record"
      Top             =   4800
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
      Left            =   8880
      Style           =   1  'Graphical
      TabIndex        =   19
      ToolTipText     =   "Print Record"
      Top             =   4320
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
      Left            =   8880
      Style           =   1  'Graphical
      TabIndex        =   18
      ToolTipText     =   "Delete Record"
      Top             =   3840
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
      Left            =   8880
      Style           =   1  'Graphical
      TabIndex        =   17
      ToolTipText     =   " Modify Record"
      Top             =   3360
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
      Left            =   8880
      Style           =   1  'Graphical
      TabIndex        =   16
      ToolTipText     =   " Add New Record"
      Top             =   2880
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
      Left            =   8880
      Style           =   1  'Graphical
      TabIndex        =   15
      ToolTipText     =   "List Of Record"
      Top             =   2400
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
      Left            =   8880
      Style           =   1  'Graphical
      TabIndex        =   14
      ToolTipText     =   "Find Record"
      Top             =   1920
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
      Left            =   9480
      Style           =   1  'Graphical
      TabIndex        =   13
      ToolTipText     =   "Last Record"
      Top             =   1440
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
      Left            =   8880
      Style           =   1  'Graphical
      TabIndex        =   12
      ToolTipText     =   "First  Record"
      Top             =   1440
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
      Left            =   9480
      Style           =   1  'Graphical
      TabIndex        =   11
      ToolTipText     =   "Previos  Record"
      Top             =   960
      Width           =   495
   End
   Begin VB.TextBox txt 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Index           =   0
      Left            =   2400
      TabIndex        =   0
      Top             =   960
      Visible         =   0   'False
      Width           =   1155
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
      Left            =   8880
      Style           =   1  'Graphical
      TabIndex        =   10
      ToolTipText     =   "Next Record"
      Top             =   960
      Width           =   495
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   " Brokerage Rate"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H004D177D&
      Height          =   240
      Index           =   10
      Left            =   4560
      TabIndex        =   36
      Top             =   5040
      Width           =   1545
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Mill Name"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H004D177D&
      Height          =   240
      Index           =   9
      Left            =   600
      TabIndex        =   35
      Top             =   2280
      Width           =   960
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Brokerage On"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H004D177D&
      Height          =   240
      Index           =   8
      Left            =   600
      TabIndex        =   33
      Top             =   5040
      Width           =   1320
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Item Unit"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H004D177D&
      Height          =   240
      Index           =   7
      Left            =   4680
      TabIndex        =   32
      Top             =   4200
      Width           =   855
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Rate Scale"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H004D177D&
      Height          =   240
      Index           =   6
      Left            =   600
      TabIndex        =   31
      Top             =   4200
      Width           =   1035
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Std Packing"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H004D177D&
      Height          =   240
      Index           =   5
      Left            =   600
      TabIndex        =   30
      Top             =   3480
      Width           =   1140
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Short Name"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H004D177D&
      Height          =   240
      Index           =   4
      Left            =   4680
      TabIndex        =   29
      Top             =   2880
      Width           =   1125
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Ticket Name"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H004D177D&
      Height          =   240
      Index           =   3
      Left            =   600
      TabIndex        =   28
      Top             =   2880
      Width           =   1185
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Item Name"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H004D177D&
      Height          =   240
      Index           =   2
      Left            =   600
      TabIndex        =   27
      Top             =   1560
      Width           =   1035
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Item Code"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H004D177D&
      Height          =   240
      Index           =   1
      Left            =   0
      TabIndex        =   26
      Top             =   0
      Width           =   975
   End
   Begin VB.Label Label6 
      BackColor       =   &H004D177D&
      Caption         =   "    Item Master"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   615
      Left            =   0
      TabIndex        =   25
      Top             =   0
      Width           =   10455
   End
   Begin VB.Shape Shape1 
      FillColor       =   &H00BDDFD1&
      FillStyle       =   0  'Solid
      Height          =   6495
      Left            =   8640
      Shape           =   4  'Rounded Rectangle
      Top             =   840
      Width           =   1575
   End
   Begin VB.Shape Shape2 
      BackColor       =   &H00808080&
      BorderColor     =   &H004D177D&
      FillColor       =   &H00ECFFFF&
      FillStyle       =   0  'Solid
      Height          =   5655
      Left            =   360
      Top             =   960
      Width           =   8055
   End
End
Attribute VB_Name = "frmItemMast"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Const tblName As String = "tblItemMast"
Const IdField As String = "ItCode"
Const OrderField As String = "ItName"
Dim rstbl As Recordset
Dim varBookMark As Variant
Dim FindFlag As Boolean
Dim FirstTimeFlag As Boolean
Private Sub cmdBtn_Click(Index As Integer)
On Error GoTo ErrorRoutine
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
    If FirstTimeFlag = True Then
        FirstTimeFlag = False
    Else
        txt(1).SetFocus
    End If
    FillCombo
    
Case 7 'Modify
    'If mURecModi = True Then
    '    SearchHelp
    '    DoEvents
        gFormAction = vbDataActionUpdate
        Call GProActivateControls(Me, True)
        Call GProSetButtons(Me, vbDataActionUpdate)
        txt(1).SetFocus
    'End If
Case 8 'Delete
    'If mURecDel = True Then
        If MsgBox("Do You Really Want to Delete This Record?", vbQuestion + vbYesNo + vbDefaultButton1, "Confirm Delete") = vbYes Then
            If rstbl.RecordCount > 1 Then
                rstbl.Delete
                cmdBtn_Click (1)
            ElseIf rstbl.RecordCount = 1 Then
                rstbl.Delete
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
        gCn.BeginTrans
        With rstbl
            If gFormAction = vbDataActionAddNew Then
                .AddNew
                txt(0) = GProcGenerateId(tblName, IdField)
            End If
            FillTxtFromTemp
            GProSaveRecord Me, rstbl
            .Update
        End With
        gCn.CommitTrans
        
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
    Dim y As String
    Me.Top = 300
    Me.Left = Screen.Width / 2 - Me.Width / 2
    Me.Caption = "ITEM MASTER for " & gCompName & " " & gFinYr
    y = "Select * from " & tblName & " order by " & OrderField
    GProcRstOpen rstbl, y, "O"
    'MaxNo = rstbl.Fields.Count - 1
    FirstTimeFlag = True
    cmdBtn_Click (6)
End Sub


Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
If gFormAction = vbDataActionUpdate Then
    MsgBox "You are about to modify the record. Save the record and then exit.", vbInformation + vbOKOnly, Me.Caption
    Cancel = True
    Exit Sub
End If
'mShowRec = Space(5)
rstbl.Requery
rstbl.Close
End Sub

Private Sub tempTxt_KeyPress(Index As Integer, KeyAscii As Integer)
If Index = 9 Then '-- Mill Name
    GProcSearchField Me, "tblMastaccount", "AcName", "", False, Chr(KeyAscii)
    KeyAscii = 0
    temptxt(Index) = gSearchRetcol1
End If
End Sub

Private Sub txt_GotFocus(Index As Integer)
GProcSelectBox Me.ActiveControl
End Sub
Private Sub txt_LostFocus(Index As Integer)
GProcCheckForNumber Me, rstbl, Index
SaveBtnEd
End Sub
Private Sub txt_KeyPress(Index As Integer, KeyAscii As Integer)
KeyAscii = GProcValidateKey(rstbl, Index, KeyAscii, txt(Index).Text)
End Sub
Private Sub txtTemp_GotFocus(Index As Integer)
GProcSelectBox Me.ActiveControl
End Sub
Private Sub txtTemp_LostFocus(Index As Integer)
SaveBtnEd
End Sub

Private Sub FillCombo()
Call GProFillCombo(cbo(6), "select Distinct ItUnit from " & tblName & " where ItUnit <> '' order by ItUnit")
Call GProFillCombo(cbo(7), "select Distinct ItBrokon from " & tblName & " where ItBrokon <> '' order by ItBrokon")
'Call GProFillCombo(cbo(9), "select Distinct AcDist from " & tblName & " where AcDist <> '' order by AcDist")
End Sub
Private Sub SaveBtnEd()
'--- AcName,AgCode,Agcode1
If txt(1) = "" Or temptxt(9) = "" Then
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
Dim y As String
Dim rstAddCheck As Recordset
Dim rstModiCheck As Recordset
    '----Duplicate Records
    '----Addition Case
    If gFormAction = vbDataActionAddNew Then
        x = "Select ItCode,ItName From " & tblName & " Where ItName='" & Trim(txt(1).Text) & "'"
        i = GProcRstOpen(rstAddCheck, x, "R")
        If i > 0 Then
            MsgBox "Record already exists  " & Trim(txt(1).Text), vbCritical + vbOKOnly, "Duplicate Record"
            txt(1).SetFocus
            Exit Function
        End If
        rstAddCheck.Close
    End If
    '----Modification case
    If gFormAction = vbDataActionUpdate Then
        x = "Select ItCode,ItName From " & tblName & " Where ((ItName='" & Trim(txt(1).Text) & "') and (ItCode<>" & txt(0).Text & "))"
        i = GProcRstOpen(rstModiCheck, x, "O")
        If i > 0 Then
            MsgBox "Record already exists  " & Trim(txt(1).Text), vbCritical + vbOKOnly, "Duplicate Record"
            txt(1).SetFocus
            Exit Function
        End If
        rstModiCheck.Close
    End If
    ValidateData = True
End Function
Private Sub FillTxtFromTemp()
txt(6) = cbo(6)
txt(7) = cbo(7)
txt(9) = GProcGetColumnValue("TBLMASTaccount", "acname", temptxt(9), "S", "accode", "", gCn)
End Sub
Private Sub FillTempFromTxt()
cbo(6) = txt(6)
cbo(7) = txt(7)
temptxt(9) = GProcGetColumnValue("TBLMASTaccount", "accode", txt(9), "N", "acname", "", gCn)
End Sub
Private Sub ReadFields()
GProShowRecord Me, rstbl
FillTempFromTxt
End Sub
