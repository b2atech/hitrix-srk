VERSION 5.00
Begin VB.Form frmMastNarration 
   BackColor       =   &H00C2E3FC&
   BorderStyle     =   0  'None
   Caption         =   " Narration Master"
   ClientHeight    =   4350
   ClientLeft      =   0
   ClientTop       =   60
   ClientWidth     =   10725
   FillColor       =   &H000040C0&
   FillStyle       =   0  'Solid
   Icon            =   "frmMastNarration.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4350
   ScaleWidth      =   10725
   ShowInTaskbar   =   0   'False
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
      Index           =   24
      Left            =   6030
      TabIndex        =   8
      Text            =   "24"
      Top             =   2880
      Width           =   930
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
      Index           =   23
      Left            =   7605
      TabIndex        =   50
      Text            =   "23"
      Top             =   4095
      Visible         =   0   'False
      Width           =   525
   End
   Begin VB.TextBox txt 
      Height          =   360
      Index           =   11
      Left            =   11280
      TabIndex        =   47
      Text            =   "11"
      Top             =   1200
      Visible         =   0   'False
      Width           =   465
   End
   Begin VB.TextBox txt 
      Height          =   360
      Index           =   12
      Left            =   11370
      TabIndex        =   46
      Text            =   "12"
      Top             =   1605
      Visible         =   0   'False
      Width           =   465
   End
   Begin VB.TextBox txt 
      Height          =   360
      Index           =   13
      Left            =   11415
      TabIndex        =   45
      Text            =   "13"
      Top             =   2010
      Visible         =   0   'False
      Width           =   465
   End
   Begin VB.TextBox txt 
      Height          =   360
      Index           =   14
      Left            =   11415
      TabIndex        =   44
      Text            =   "14"
      Top             =   2415
      Visible         =   0   'False
      Width           =   465
   End
   Begin VB.TextBox txt 
      Height          =   360
      Index           =   15
      Left            =   11460
      TabIndex        =   43
      Text            =   "15"
      Top             =   2865
      Visible         =   0   'False
      Width           =   465
   End
   Begin VB.TextBox txt 
      Height          =   360
      Index           =   16
      Left            =   11415
      TabIndex        =   42
      Text            =   "16"
      Top             =   3270
      Visible         =   0   'False
      Width           =   465
   End
   Begin VB.TextBox txt 
      Height          =   360
      Index           =   17
      Left            =   11460
      TabIndex        =   41
      Text            =   "17"
      Top             =   3675
      Visible         =   0   'False
      Width           =   465
   End
   Begin VB.TextBox txt 
      Height          =   360
      Index           =   18
      Left            =   11460
      TabIndex        =   40
      Text            =   "18"
      Top             =   4035
      Visible         =   0   'False
      Width           =   465
   End
   Begin VB.TextBox txt 
      Height          =   360
      Index           =   19
      Left            =   11415
      TabIndex        =   39
      Text            =   "19"
      Top             =   4440
      Visible         =   0   'False
      Width           =   465
   End
   Begin VB.TextBox txt 
      Height          =   360
      Index           =   20
      Left            =   11460
      TabIndex        =   38
      Text            =   "20"
      Top             =   4800
      Visible         =   0   'False
      Width           =   465
   End
   Begin VB.TextBox txt 
      Height          =   360
      Index           =   21
      Left            =   11415
      TabIndex        =   37
      Text            =   "21"
      Top             =   5160
      Visible         =   0   'False
      Width           =   465
   End
   Begin VB.TextBox txt 
      Height          =   360
      Index           =   10
      Left            =   12000
      TabIndex        =   36
      Text            =   "10"
      Top             =   1200
      Visible         =   0   'False
      Width           =   465
   End
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
      Height          =   390
      Index           =   22
      Left            =   2280
      TabIndex        =   9
      Text            =   "22"
      Top             =   3360
      Width           =   3285
   End
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
      Left            =   9120
      Style           =   1  'Graphical
      TabIndex        =   34
      Top             =   2340
      Width           =   1095
   End
   Begin VB.ComboBox cbo 
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
      Left            =   2280
      TabIndex        =   4
      Text            =   "cbo(5)"
      Top             =   2400
      Width           =   2595
   End
   Begin VB.TextBox txt 
      Height          =   360
      Index           =   9
      Left            =   3105
      TabIndex        =   33
      Text            =   "9"
      Top             =   3975
      Visible         =   0   'False
      Width           =   465
   End
   Begin VB.TextBox txt 
      Height          =   360
      Index           =   8
      Left            =   2490
      TabIndex        =   32
      Text            =   "8"
      Top             =   4005
      Visible         =   0   'False
      Width           =   390
   End
   Begin VB.TextBox txt 
      Height          =   360
      Index           =   7
      Left            =   5850
      TabIndex        =   5
      Text            =   "7"
      Top             =   2400
      Width           =   1440
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
      Left            =   5550
      MaxLength       =   40
      TabIndex        =   29
      Text            =   "1"
      Top             =   3810
      Visible         =   0   'False
      Width           =   915
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
      Left            =   6630
      TabIndex        =   28
      Text            =   "0"
      Top             =   3930
      Visible         =   0   'False
      Width           =   795
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
      Index           =   6
      Left            =   2280
      TabIndex        =   7
      Text            =   "6"
      Top             =   2880
      Width           =   2595
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
      Index           =   5
      Left            =   3660
      TabIndex        =   6
      Text            =   "5"
      Top             =   3960
      Visible         =   0   'False
      Width           =   525
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
      Index           =   4
      Left            =   2280
      TabIndex        =   3
      Text            =   "4"
      Top             =   1920
      Width           =   4995
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
      Index           =   3
      Left            =   2280
      TabIndex        =   2
      Text            =   "3"
      Top             =   1440
      Width           =   4995
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
      Left            =   2280
      TabIndex        =   1
      Text            =   "2"
      Top             =   960
      Width           =   4995
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
      TabIndex        =   22
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
      Left            =   9105
      Style           =   1  'Graphical
      TabIndex        =   21
      ToolTipText     =   "Exit"
      Top             =   3345
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
      Left            =   7920
      Style           =   1  'Graphical
      TabIndex        =   11
      ToolTipText     =   "To Cancel Record"
      Top             =   3360
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
      Left            =   7920
      Style           =   1  'Graphical
      TabIndex        =   10
      ToolTipText     =   "To Save Record"
      Top             =   2895
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
      TabIndex        =   20
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
      Left            =   7920
      Style           =   1  'Graphical
      TabIndex        =   19
      ToolTipText     =   "To Delete Record"
      Top             =   2295
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
      Left            =   9120
      Style           =   1  'Graphical
      TabIndex        =   18
      ToolTipText     =   "To Modify Record"
      Top             =   1815
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
      Left            =   7920
      Style           =   1  'Graphical
      TabIndex        =   17
      ToolTipText     =   "To Add New Record"
      Top             =   1815
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
      Left            =   9120
      Style           =   1  'Graphical
      TabIndex        =   16
      Top             =   1215
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
      Left            =   7920
      Style           =   1  'Graphical
      TabIndex        =   15
      ToolTipText     =   "To Find Record"
      Top             =   1215
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
      Left            =   9720
      Style           =   1  'Graphical
      TabIndex        =   14
      ToolTipText     =   "Last"
      Top             =   735
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
      Left            =   9120
      Style           =   1  'Graphical
      TabIndex        =   13
      ToolTipText     =   "First"
      Top             =   735
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
      Left            =   8520
      Style           =   1  'Graphical
      TabIndex        =   12
      ToolTipText     =   "Next"
      Top             =   735
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
      Left            =   7920
      Style           =   1  'Graphical
      TabIndex        =   0
      ToolTipText     =   "Previous"
      Top             =   735
      Width           =   495
   End
   Begin VB.Label lbl1 
      BackStyle       =   0  'Transparent
      Caption         =   "Distance :"
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
      Height          =   315
      Index           =   7
      Left            =   5085
      TabIndex        =   49
      Top             =   2880
      Width           =   975
   End
   Begin VB.Label lbl1 
      BackStyle       =   0  'Transparent
      Caption         =   "Pin  :"
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
      Height          =   315
      Index           =   2
      Left            =   5085
      TabIndex        =   48
      Top             =   2400
      Width           =   480
   End
   Begin VB.Label lbl1 
      BackStyle       =   0  'Transparent
      Caption         =   "GSTIN No          :"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   240
      Index           =   1
      Left            =   480
      TabIndex        =   35
      Top             =   3405
      Width           =   1650
   End
   Begin VB.Shape Shape3 
      BorderColor     =   &H00C00000&
      BorderWidth     =   5
      FillColor       =   &H00C00000&
      Height          =   4320
      Left            =   0
      Top             =   0
      Width           =   10680
   End
   Begin VB.Shape Shape5 
      BorderColor     =   &H00000080&
      FillColor       =   &H00000080&
      FillStyle       =   0  'Solid
      Height          =   15
      Left            =   0
      Top             =   4410
      Width           =   10695
   End
   Begin VB.Label lbl1 
      BackStyle       =   0  'Transparent
      Caption         =   "    :"
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
      Height          =   255
      Index           =   4
      Left            =   1800
      TabIndex        =   31
      Top             =   2160
      Visible         =   0   'False
      Width           =   495
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
      TabIndex        =   30
      Top             =   60
      Width           =   2475
   End
   Begin VB.Label lbl1 
      BackStyle       =   0  'Transparent
      Caption         =   "Godown &Name   :"
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
      Height          =   240
      Index           =   0
      Left            =   480
      TabIndex        =   27
      Top             =   960
      Width           =   1650
   End
   Begin VB.Label lbl1 
      BackStyle       =   0  'Transparent
      Caption         =   "&Phone No.          :"
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
      TabIndex        =   26
      Top             =   2880
      Width           =   1575
   End
   Begin VB.Label lbl1 
      BackStyle       =   0  'Transparent
      Caption         =   "City                    :"
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
      TabIndex        =   25
      Top             =   2400
      Width           =   1695
   End
   Begin VB.Label lbl1 
      BackStyle       =   0  'Transparent
      Caption         =   "Add&ress              :"
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
      TabIndex        =   24
      Top             =   1440
      Width           =   1815
   End
   Begin VB.Label Label6 
      BackColor       =   &H00C00000&
      Caption         =   "   Godown Master"
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
      TabIndex        =   23
      Top             =   0
      Width           =   10695
   End
   Begin VB.Shape Shape1 
      FillColor       =   &H00BDDFD1&
      FillStyle       =   0  'Solid
      Height          =   3390
      Left            =   7680
      Shape           =   4  'Rounded Rectangle
      Top             =   570
      Width           =   2775
   End
   Begin VB.Shape Shape2 
      BackColor       =   &H00808080&
      BorderColor     =   &H004D177D&
      FillColor       =   &H00ECFFFF&
      FillStyle       =   0  'Solid
      Height          =   3120
      Left            =   240
      Top             =   645
      Width           =   7245
   End
End
Attribute VB_Name = "frmMastNarration"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Const tblName As String = "tblMastNarration"
Const IdField As String = "narrcode"
Const OrderField As String = "Narration"
Const MaxNo As Long = 24
'Const CriteriaStr As String = " narrtp = '" & gfrmTypeStr & "'"
Dim rstbl As Recordset
Dim mNarrTp As String
Dim i As Long
'Dim varBookMark As Variant
Dim FirstTimeFlag As Boolean
Dim FormAction As Integer
Private Sub Cbo_KeyPress(Index As Integer, KeyAscii As Integer)
If KeyAscii = 13 Then Exit Sub
End Sub
Private Sub CheckForNumber()
GProcCheckForNumber Me, rstbl, MaxNo, Array(22), Array(23, 36)
End Sub
Private Sub cbo_LostFocus(Index As Integer)
If Index = 5 Then '--city
    cbo(Index).text = GProcProperCase(Trim(cbo(Index).text))
End If
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
    gClsSearch.SearchMultiField tblName, "Narration", Array("Narration"), Array(txt(2).Width), " NarrType='" & mNarrTp & "'", "", "Narration", txt(2).Left + Me.Left, txt(2).Top + Me.Top + 650
    If gClsSearch.SearchMultiRetCol(0) <> "" Then
        rstbl.MoveFirst
        rstbl.Find "Narration= '" & gClsSearch.SearchMultiRetCol(0) & "'"
        gClsSearch.SearchMultiRetCol(0) = ""
        ReadFields
    End If
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
            If rstbl.RecordCount > 1 Then
                rstbl.Delete
                cmdBtn_Click (1)
            ElseIf rstbl.RecordCount = 1 Then
                rstbl.Delete
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
    Select Case mNarrTp
        Case "G" '-Godown
            gfrmTypeStr = "Godown"
        Case "N" '-Narration
            gfrmTypeStr = "Narration"
        Case "M" '-Mill Expense Des.
            gfrmTypeStr = "MillExp"
        Case "R" '-Transport
            gfrmTypeStr = "Transport"
        Case "S" '-Sizer
            gfrmTypeStr = "Sizer"
    End Select
    frmUtlCreationToFirm.Show 1
    rstbl.Requery
Case 13 'Save
    If ValidateData = True Then
        FillTxtFromTemp
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
    Call GProcClearForm(Me, rstbl, MaxNo, True)
    If rstbl.RecordCount <= 0 Then
        Call GProcActivateControls(Me, False)
        Call GProcSetButtons(Me, vbDataActionCancel)
    Else
        If FormAction = vbDataActionAddNew Then
            rstbl.MoveLast
        Else
            rstbl.CancelUpdate
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
    rstbl.CancelUpdate
    GProcErrorHandler
End Sub

Private Sub Form_Load()
    Dim X As String
    mNarrTp = gfrmTypeStr
    Me.Top = 300
    Me.Left = Screen.Width / 2 - Me.Width / 2
    X = "Select * from " & tblName & " where NarrType='" & mNarrTp & "'"
    If Trim(OrderField) <> "" Then X = X & " order by " & OrderField
    GProcSetButtonVisiblity Me
    GProcRstOpen rstbl, X, "O"
    FirstTimeFlag = True
    For i = 3 To 6
        If i = 5 Then cbo(i).Visible = False
        lbl1(i).Visible = False
        txt(i).Visible = False
    Next i
    Select Case mNarrTp
    Case "N":
        Label6.Caption = "    Narration Master"
        lbl1(0).Caption = "Narration      :"
    Case "G":
        Label6.Caption = "    Godown Master"
        lbl1(0).Caption = "Godown Name  :"
        For i = 3 To 6
            If i = 5 Then
                cbo(i).Visible = True
            Else
                txt(i).Visible = True
            End If
            lbl1(i).Visible = True
        Next i
    Case "R":
         Label6.Caption = "    Transport Master"
         lbl1(0).Caption = "Transport Name :"
         For i = 3 To 6
            If i = 5 Then
                cbo(i).Visible = True
            Else
                txt(i).Visible = True
            End If
            lbl1(i).Visible = True
         Next i
    Case "M":
           Label6.Caption = "    Mills Expenses Master"
           lbl1(0).Caption = "Mill Expenses  :"
    Case "S": '--Sizer
        Label6.Caption = "    Sizer Master"
        lbl1(0).Caption = "Sizer Name  :"
        For i = 3 To 6
            If i = 5 Then
                cbo(i).Visible = True
            Else
                txt(i).Visible = True
            End If
            lbl1(i).Visible = True
        Next i
   End Select
   cmdBtn_Click (6)
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
If FormAction = vbDataActionUpdate Then
    MsgBox "You are about to modify the record. Save the record and then exit.", vbInformation + vbOKOnly, Me.Caption
    Cancel = True
    Exit Sub
End If
'mShowRec = Space(5)
rstbl.Close
End Sub

Private Sub txt_GotFocus(Index As Integer)
GProcSelectBox Me.ActiveControl
End Sub

Private Sub txt_KeyPress(Index As Integer, KeyAscii As Integer)
If KeyAscii = 13 Then Exit Sub
KeyAscii = GProcValidateKey(rstbl, Index, KeyAscii, txt(Index).text)
End Sub

Private Sub txt_LostFocus(Index As Integer)
CheckForNumber
'GProcCheckForNumber Me, Rstbl, MaxNo
SaveBtnEd
End Sub
Private Sub FillCombo()
Call GProcFillCombo(cbo(5), "select Distinct AcCity from " & "tblMastAccount" & " where AcCity <> '' order by AcCity")
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
    GProcCheckForNumber Me, rstbl, MaxNo
    '---- Required
    '-- NarrCode
    If FormAction = vbDataActionAddNew Then txt(1) = GProcGenerateId(tblName, IdField, pStartCode:=gCSNarrCode)
    If gCENarrCode > 0 And CLng(txt(1)) > gCENarrCode Then
        MsgBox "Year End is completed. You can not create new narration.", vbCritical + vbOKOnly, Me.Caption
        txt(2).SetFocus
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
        X = "Select NArration From " & tblName & " Where Narration='" & Trim(txt(2).text) & "'"
        i = GProcRstOpen(rstAddCheck, X, "R")
        If i > 0 Then
            MsgBox "Record is Available of Same Narration  " & Trim(txt(2).text), vbCritical + vbOKOnly, "Duplicate Record"
            txt(2).SetFocus
            Exit Function
        End If
        rstAddCheck.Close
    End If
    '----Modification case
    If FormAction = vbDataActionUpdate Then
        X = "Select narration From " & tblName & " Where ((narration='" & Trim(txt(2).text) & "') and (narrCode<>" & txt(1).text & "))"
        i = GProcRstOpen(rstModiCheck, X, "O")
        If i > 0 Then
            MsgBox "Record is Available of Same Narration  " & Trim(txt(2).text), vbCritical + vbOKOnly, "Duplicate Record"
            txt(2).SetFocus
            Exit Function
        End If
        rstModiCheck.Close
    End If
    ValidateData = True
End Function
Private Sub FillTxtFromTemp()
txt(0) = mNarrTp
'-- Tax Purchase,Sale Accode
txt(8) = GProcGetColumnValue("TblMastAccount", "AcName", "-", "S", "AcCode", "N")
txt(9) = GProcGetColumnValue("TblMastAccount", "AcName", "-", "S", "AcCode", "N")
txt(5) = cbo(5)
End Sub
Private Sub FillTempFromTxt()
cbo(5) = txt(5)
End Sub
Private Sub ReadFields()
GProcShowRecord Me, rstbl, MaxNo
FillTempFromTxt
End Sub

