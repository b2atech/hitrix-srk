VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form frmMastItem 
   BackColor       =   &H00C0E0FF&
   BorderStyle     =   0  'None
   Caption         =   " Count Master"
   ClientHeight    =   6915
   ClientLeft      =   0
   ClientTop       =   60
   ClientWidth     =   9720
   FillColor       =   &H000040C0&
   FillStyle       =   0  'Solid
   Icon            =   "frmMastItem.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6915
   ScaleWidth      =   9720
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox txt 
      Alignment       =   1  'Right Justify
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
      Height          =   360
      Index           =   17
      Left            =   4920
      TabIndex        =   54
      Text            =   "17"
      Top             =   6360
      Visible         =   0   'False
      Width           =   1350
   End
   Begin VB.ComboBox Cbo 
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
      ItemData        =   "frmMastItem.frx":058A
      Left            =   6120
      List            =   "frmMastItem.frx":0597
      TabIndex        =   52
      Text            =   "Cbo(8)"
      Top             =   4995
      Width           =   1515
   End
   Begin VB.TextBox txt 
      Alignment       =   1  'Right Justify
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
      Height          =   360
      Index           =   15
      Left            =   2280
      TabIndex        =   48
      Text            =   "15"
      Top             =   5910
      Width           =   1515
   End
   Begin VB.TextBox txt 
      Alignment       =   1  'Right Justify
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
      Height          =   360
      Index           =   16
      Left            =   6360
      TabIndex        =   47
      Text            =   "16"
      Top             =   6360
      Visible         =   0   'False
      Width           =   1350
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
      Height          =   615
      Index           =   12
      Left            =   8280
      Style           =   1  'Graphical
      TabIndex        =   46
      Top             =   2550
      Width           =   1095
   End
   Begin VB.TextBox txt 
      Alignment       =   1  'Right Justify
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
      Height          =   360
      Index           =   14
      Left            =   2280
      TabIndex        =   15
      Text            =   "14"
      Top             =   4215
      Width           =   2115
   End
   Begin VB.TextBox txt 
      Alignment       =   1  'Right Justify
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
      Height          =   360
      Index           =   13
      Left            =   5865
      TabIndex        =   17
      Text            =   "13"
      Top             =   4215
      Width           =   1350
   End
   Begin VB.TextBox txt 
      Alignment       =   1  'Right Justify
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
      Height          =   360
      Index           =   12
      Left            =   6090
      TabIndex        =   45
      Text            =   "12"
      Top             =   3630
      Visible         =   0   'False
      Width           =   420
   End
   Begin VB.TextBox txt 
      Alignment       =   1  'Right Justify
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
      Height          =   360
      Index           =   11
      Left            =   5475
      TabIndex        =   44
      Text            =   "11"
      Top             =   3630
      Visible         =   0   'False
      Width           =   540
   End
   Begin VB.TextBox txt 
      Alignment       =   1  'Right Justify
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
      Height          =   360
      Index           =   10
      Left            =   6120
      TabIndex        =   23
      Text            =   "10"
      Top             =   5422
      Width           =   1155
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
      Index           =   7
      Left            =   3240
      TabIndex        =   42
      Text            =   "7"
      Top             =   6600
      Visible         =   0   'False
      Width           =   315
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
      Index           =   6
      Left            =   3720
      TabIndex        =   41
      Text            =   "6"
      Top             =   6600
      Visible         =   0   'False
      Width           =   315
   End
   Begin VB.TextBox txttemp 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   9
      Left            =   2280
      TabIndex        =   5
      Text            =   "txttemp(9)"
      ToolTipText     =   "Press F3 For Creation"
      Top             =   1920
      Width           =   3630
   End
   Begin VB.ComboBox Cbo 
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
      ItemData        =   "frmMastItem.frx":05A9
      Left            =   2280
      List            =   "frmMastItem.frx":05B6
      TabIndex        =   19
      Text            =   "Cbo(7)"
      Top             =   4995
      Width           =   1515
   End
   Begin VB.ComboBox Cbo 
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
      ItemData        =   "frmMastItem.frx":05C8
      Left            =   5865
      List            =   "frmMastItem.frx":05D2
      TabIndex        =   13
      Text            =   "Cbo"
      Top             =   3180
      Width           =   1350
   End
   Begin VB.TextBox txt 
      Alignment       =   1  'Right Justify
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
      Height          =   360
      Index           =   8
      Left            =   2280
      TabIndex        =   21
      Text            =   "8"
      Top             =   5422
      Width           =   1515
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
      Index           =   9
      Left            =   2760
      TabIndex        =   40
      Text            =   "9"
      Top             =   6600
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
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Index           =   5
      Left            =   2280
      TabIndex        =   11
      Text            =   "5"
      Top             =   3660
      Width           =   2115
   End
   Begin VB.TextBox txt 
      Alignment       =   1  'Right Justify
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
      Height          =   360
      Index           =   4
      Left            =   2280
      TabIndex        =   9
      Text            =   "4"
      Top             =   3180
      Width           =   2115
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
      Height          =   375
      Index           =   3
      Left            =   2280
      TabIndex        =   3
      Text            =   "3"
      Top             =   1440
      Width           =   3630
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
      Height          =   375
      Index           =   2
      Left            =   2280
      TabIndex        =   7
      Text            =   "2"
      Top             =   2400
      Width           =   3630
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
      Height          =   375
      Index           =   1
      Left            =   2280
      TabIndex        =   1
      Text            =   "1"
      Top             =   960
      Width           =   3630
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
      Left            =   4800
      Style           =   1  'Graphical
      TabIndex        =   37
      ToolTipText     =   "Label"
      Top             =   7200
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
      Left            =   8280
      Style           =   1  'Graphical
      TabIndex        =   35
      ToolTipText     =   "Exit from form"
      Top             =   5640
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
      Left            =   8280
      Style           =   1  'Graphical
      TabIndex        =   28
      ToolTipText     =   "cancel"
      Top             =   5160
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
      Left            =   8280
      Style           =   1  'Graphical
      TabIndex        =   27
      ToolTipText     =   "Save Record"
      Top             =   4680
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
      Left            =   7200
      Style           =   1  'Graphical
      TabIndex        =   36
      ToolTipText     =   "Print Record"
      Top             =   7320
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
      Left            =   8280
      Style           =   1  'Graphical
      TabIndex        =   26
      ToolTipText     =   "Delete Record"
      Top             =   4185
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
      Left            =   8280
      Style           =   1  'Graphical
      TabIndex        =   25
      ToolTipText     =   " Modify Record"
      Top             =   3705
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
      Left            =   8280
      Style           =   1  'Graphical
      TabIndex        =   24
      ToolTipText     =   " Add New Record"
      Top             =   3225
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
      Left            =   8280
      Style           =   1  'Graphical
      TabIndex        =   29
      ToolTipText     =   "List Of Record"
      Top             =   2085
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
      Left            =   8280
      Style           =   1  'Graphical
      TabIndex        =   30
      ToolTipText     =   "Find Record"
      Top             =   1620
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
      Left            =   8880
      Style           =   1  'Graphical
      TabIndex        =   34
      ToolTipText     =   "Last"
      Top             =   1155
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
      Left            =   8280
      Style           =   1  'Graphical
      TabIndex        =   33
      ToolTipText     =   "First"
      Top             =   1155
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
      Left            =   8880
      Style           =   1  'Graphical
      TabIndex        =   32
      ToolTipText     =   "Next"
      Top             =   675
      Width           =   495
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
      Left            =   2760
      TabIndex        =   38
      Text            =   "0"
      Top             =   600
      Visible         =   0   'False
      Width           =   315
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
      Left            =   8280
      Style           =   1  'Graphical
      TabIndex        =   31
      ToolTipText     =   "Previous"
      Top             =   675
      Width           =   495
   End
   Begin MSMask.MaskEdBox mskDt 
      Height          =   360
      Left            =   2280
      TabIndex        =   49
      Top             =   6330
      Width           =   1515
      _ExtentX        =   2672
      _ExtentY        =   635
      _Version        =   393216
      MaxLength       =   10
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   "dd/mm/yyyy"
      Mask            =   "##/##/####"
      PromptChar      =   "_"
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00000080&
      X1              =   4080
      X2              =   4080
      Y1              =   4800
      Y2              =   6720
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "&Brokerage On     :"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H004D177D&
      Height          =   240
      Index           =   14
      Left            =   4320
      TabIndex        =   53
      Top             =   5055
      Width           =   1560
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Before Date          :"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H004D177D&
      Height          =   240
      Index           =   13
      Left            =   480
      TabIndex        =   51
      Top             =   6390
      Width           =   1680
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Old Brokerage Rt. :"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H004D177D&
      Height          =   240
      Index           =   12
      Left            =   480
      TabIndex        =   50
      Top             =   5970
      Width           =   1680
   End
   Begin VB.Image Image1 
      Height          =   1485
      Left            =   6105
      Picture         =   "frmMastItem.frx":05E4
      Stretch         =   -1  'True
      Top             =   1080
      Width           =   1365
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Min Rate           :"
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
      Height          =   240
      Index           =   11
      Left            =   480
      TabIndex        =   14
      Top             =   4215
      Width           =   1500
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Max Rate  :"
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
      Height          =   360
      Index           =   1
      Left            =   4680
      TabIndex        =   16
      Top             =   4215
      Width           =   1020
   End
   Begin VB.Shape Shape6 
      BorderColor     =   &H00C00000&
      BorderWidth     =   5
      FillColor       =   &H00C00000&
      Height          =   7080
      Left            =   0
      Top             =   -225
      Width           =   9720
   End
   Begin VB.Shape Shape5 
      BorderColor     =   &H00C00000&
      FillColor       =   &H00C00000&
      FillStyle       =   0  'Solid
      Height          =   75
      Left            =   0
      Top             =   6840
      Width           =   9735
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
      Left            =   6975
      TabIndex        =   43
      Top             =   15
      Width           =   2475
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "&Purch Brok Rt.    :"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H004D177D&
      Height          =   240
      Index           =   0
      Left            =   4320
      TabIndex        =   22
      Top             =   5482
      Width           =   1590
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Brokerage Rate :"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H004D177D&
      Height          =   375
      Index           =   10
      Left            =   480
      TabIndex        =   20
      Top             =   5415
      Width           =   1470
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "M&ill Name         :"
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
      Index           =   9
      Left            =   480
      TabIndex        =   4
      Top             =   1920
      Width           =   1560
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "&Brokerage On     :"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H004D177D&
      Height          =   240
      Index           =   8
      Left            =   480
      TabIndex        =   18
      Top             =   4995
      Width           =   1560
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Item Unit  :"
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
      Index           =   7
      Left            =   4680
      TabIndex        =   12
      Top             =   3180
      Width           =   1035
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "&Rate Scale        :"
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
      Height          =   240
      Index           =   6
      Left            =   480
      TabIndex        =   10
      Top             =   3660
      Width           =   1500
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Std Packing      :"
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
      Height          =   240
      Index           =   5
      Left            =   480
      TabIndex        =   8
      Top             =   3180
      Width           =   1485
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Short Name      :"
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
      Index           =   4
      Left            =   480
      TabIndex        =   2
      Top             =   1440
      Width           =   1545
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Ticket Name     :"
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
      Index           =   3
      Left            =   480
      TabIndex        =   6
      Top             =   2400
      Width           =   1545
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Count &Name     :"
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
      Index           =   2
      Left            =   480
      TabIndex        =   0
      Top             =   960
      Width           =   1530
   End
   Begin VB.Label Label6 
      BackColor       =   &H00C00000&
      Caption         =   "   Count Master"
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
      TabIndex        =   39
      Top             =   0
      Width           =   9735
   End
   Begin VB.Shape Shape1 
      FillColor       =   &H00BDDFD1&
      FillStyle       =   0  'Solid
      Height          =   5730
      Left            =   8085
      Shape           =   4  'Rounded Rectangle
      Top             =   495
      Width           =   1485
   End
   Begin VB.Shape Shape2 
      BackColor       =   &H00808080&
      BorderColor     =   &H004D177D&
      FillColor       =   &H00ECFFFF&
      FillStyle       =   0  'Solid
      Height          =   2175
      Left            =   240
      Top             =   720
      Width           =   7455
   End
   Begin VB.Shape Shape3 
      BorderColor     =   &H00000080&
      FillColor       =   &H00EDFDFE&
      FillStyle       =   0  'Solid
      Height          =   1740
      Left            =   240
      Top             =   2985
      Width           =   7455
   End
   Begin VB.Shape Shape4 
      BorderColor     =   &H00000080&
      FillColor       =   &H00EDFDFE&
      FillStyle       =   0  'Solid
      Height          =   1935
      Left            =   240
      Top             =   4800
      Width           =   7455
   End
End
Attribute VB_Name = "frmMastItem"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Const tblName As String = "tblMastItem"
Const IdField As String = "ItCode"
Const OrderField As String = "ItName"
Const MaxNo As Long = 17
Dim Rstbl As Recordset
Dim varBookMark As Variant
Dim FirstTimeFlag As Boolean
Dim FormAction As Integer
Private Sub Cbo_KeyPress(Index As Integer, KeyAscii As Integer)
If KeyAscii = 13 Then Exit Sub
If Index = 7 Or Index = 6 Then KeyAscii = 0
End Sub
Private Sub cmdBtn_Click(Index As Integer)
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
    gClsSearch.SearchMultiField "tblMastItem,tblMastAccount", "ItName,ItCode,Itticket,Acname,ItStdpack,ItStdrateper,ITUnit", Array("Item", "Item", "Ticket", "Mill", "Std.Packing", "Rate Per", "Unit"), Array(2000, 0, 1000, 2000, 1000, 1000, 1000), " tblMastItem.ItMillcode=tblMastAccount.AcCode", "", "ItName", txt(1).Left + Me.Left, txt(1).Top + Me.Top + 650
    If gClsSearch.SearchMultiRetCol(0) <> "" Then
        Rstbl.MoveFirst
        Rstbl.Find "ItCode= " & gClsSearch.SearchMultiRetCol(1)
        gClsSearch.SearchMultiRetCol(0) = ""
        ReadFields
    End If
Case 5 'List
    frmMain.mnuRptGenrSubLst_Click (5)
Case 6 'Add
    FormAction = vbDataActionAddNew
    Call GProcActivateControls(Me, True)
    Call GProcSetButtons(Me, vbDataActionAddNew)
    Call GProcClearForm(Me, Rstbl, MaxNo, True)
    If FirstTimeFlag = True Then
        FirstTimeFlag = False
    Else
        txt(1).SetFocus
    End If
    
Case 7 'Modify
    'If mURecModi = True Then
    '    SearchHelp
    '    DoEvents
    'If IsRelatedRecord = False Then
        cmdBtn_Click 4 '-- Find
        If gClsSearch.SearchMultiRetCol(0) = "" Then Exit Sub
        FormAction = vbDataActionUpdate
        Call GProcActivateControls(Me, True)
        Call GProcSetButtons(Me, vbDataActionUpdate)
        '--Related Mill & Item Record
        If IsRelatedRecord = True Then txttemp(9).Enabled = False
        txt(1).SetFocus
    'End If
Case 8 'Delete
    'If mURecDel = True Then
     'If IsRelatedRecord = False Then
        If MsgBox("Do You Really Want to Delete This Record?", vbQuestion + vbYesNo + vbDefaultButton1, "Confirm Delete") = vbYes Then
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
    gfrmTypeStr = "Item"
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
    X = "Select * from " & tblName
    If Trim(OrderField) <> "" Then X = X & " order by " & OrderField
    GProcRstOpen Rstbl, X, "O"
    '---- Button Visiblity Set As Per User Previlage
    GProcSetButtonVisiblity Me
    FirstTimeFlag = True
    cmdBtn_Click (6)
    Cbo(6).ListIndex = 0
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
If Index = 1 Then
    Dim m As Integer
    For m = 1 To Len(txt(Index))
    If (Mid(txt(Index), m, 1)) = "'" Then
    txt(1) = Left(txt(Index), m - 1) + Mid(txt(Index), m + 1)
    End If
Next m
End If
GProcCheckForNumber Me, Rstbl, MaxNo
SaveBtnEd
End Sub
Private Sub txt_KeyPress(Index As Integer, KeyAscii As Integer)
If KeyAscii = 13 Then Exit Sub
KeyAscii = GProcValidateKey(Rstbl, Index, KeyAscii, txt(Index).Text)
End Sub
Private Sub txttemp_GotFocus(Index As Integer)
GProcSelectBox Me.ActiveControl
End Sub
Private Sub txtTemp_LostFocus(Index As Integer)
SaveBtnEd
End Sub
Private Sub txtTemp_KeyPress(Index As Integer, KeyAscii As Integer)
Dim LeftPos As Long
Dim TopPos As Long
Dim MillCode As Long
TopPos = txttemp(Index).Top + Me.Top + 650
LeftPos = txttemp(Index).Left + Me.Left
If KeyAscii = 13 Then Exit Sub
If Index = 9 Then '-- Mill Name
    gClsSearch.SearchMultiField "QryMastaccount", "Account,AcName,AcCity,AgName", Array("Mill", "AcName", "City", "Group"), Array(3000, 0, 2000, 3000), " (AgCode=90029 or GpCode=90029)", Chr(KeyAscii), "AcName", LeftPos, TopPos, True
    KeyAscii = 0
    If gClsSearch.SearchMultiRetCol(0) <> "" And gClsSearch.SearchMultiRetCol(1) = "" Then
        txttemp(Index) = gClsSearch.SearchMultiRetCol(0)
    Else
        If gClsSearch.SearchMultiRetCol(1) <> "" Then txttemp(Index) = gClsSearch.SearchMultiRetCol(1)
    End If
    MillCode = GProcGetColumnValue("tblMastAccount", "AcName", txttemp(9), "S", "AcCode", "N")
    '-- Ticket
    If txt(2) = "" Then txt(2) = GProcGetColumnValue("tblMastItem", "ItMillCode", CStr(MillCode), "N", "ItTicket", "S")
End If
End Sub
Private Sub txttemp_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
Dim GName As String
If KeyCode = vbKeyF3 Then
    Select Case Index
        Case 9 '-- mill
            GName = GProcGetColumnValue("TblMastGroup", "AgCode", 90029, "N", "AgName", "S")
            GProcShowForm frmMastAccount, frmMain.mnuMstfrm(0), Array(frmMastAccount.txttemp(2)), Array(GName)
    End Select
End If
End Sub
Private Sub SaveBtnEd()
'--- Name,Ticket,Short,Unit,Mill
If Trim(txt(1)) = "" Or Trim(txt(2)) = "" Or Trim(txt(3)) = "" Or Trim(Cbo(6)) = "" Or Trim(txttemp(9)) = "" Then
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
Dim GName As String
Dim MillCode As Long
Dim rstAddCheck As Recordset
Dim rstModiCheck As Recordset
    GProcCheckForNumber Me, Rstbl, MaxNo
    '--- Required
    '-- ITCode
    If FormAction = vbDataActionAddNew Then txt(0) = GProcGenerateId(tblName, IdField, pStartCode:=gCSItCode)
    If gCEItCode > 0 And CLng(txt(0)) > gCEItCode Then
        MsgBox "Year End is completed. You can not create new count.", vbCritical + vbOKOnly, Me.Caption
        txt(1).SetFocus
        Exit Function
    End If
    '--- Name
    If Trim(txt(1)) = "" Then
        MsgBox "Check Item Name.", vbCritical + vbOKOnly, Me.Caption
        txt(1).SetFocus
        Exit Function
    End If
    '--- Ticket
    If Trim(txt(2)) = "" Then
        MsgBox "Check Item Ticket.", vbCritical + vbOKOnly, Me.Caption
        txt(2).SetFocus
        Exit Function
    End If
    '---Short
    If Trim(txt(3)) = "" Then
        MsgBox "Check Item Short.", vbCritical + vbOKOnly, Me.Caption
        txt(3).SetFocus
        Exit Function
    End If
    '---Unit
    If Trim(Cbo(6)) = "" Then
        MsgBox "Check Unit.", vbCritical + vbOKOnly, Me.Caption
        Cbo(6).SetFocus
        Exit Function
    End If
    '--- Mill
    If Trim(txttemp(9)) = "" Then
        MsgBox "Check Mill.", vbCritical + vbOKOnly, Me.Caption
        txttemp(9).SetFocus
        Exit Function
    End If
    '-- Create Mill
    If GProcGetColumnValue("tblMastAccount", "AcName", txttemp(9), "S", "AcCode", "N") = 0 Then
        GName = GProcGetColumnValue("TblMastGroup", "AgCode", 90029, "N", "AgName", "S")
        GProcShowForm frmMastAccount, frmMain.mnuMstfrm(0), Array(frmMastAccount.txt(1), frmMastAccount.txttemp(2)), Array(txttemp(9), GName)
        Exit Function
    End If
    '-- Max Rate
    If CDbl(txt(13)) < CDbl(txt(14)) Then
        MsgBox "Max Rate is less than Min Rate.", vbCritical + vbOKOnly, App.Title
        txt(13).SetFocus
        Exit Function
    End If
    '----Duplicate Records
    '----Addition Case
    MillCode = GProcGetColumnValue("TblMastAccount", "AcName", txttemp(9), "S", "Accode", "N")
    If FormAction = vbDataActionAddNew Then
        X = "Select ItCode,ItName From " & tblName & " Where ItMillCode=" & MillCode & " and ItName='" & Trim(txt(1).Text) & "'"
        i = GProcRstOpen(rstAddCheck, X, "R")
        If i > 0 Then
            MsgBox "Record already exists  " & Trim(txt(1).Text), vbCritical + vbOKOnly, "Duplicate Record"
            txt(1).SetFocus
            Exit Function
        End If
        rstAddCheck.Close
    End If
    '----Modification case
    If FormAction = vbDataActionUpdate Then
        X = "Select ItCode,ItName From " & tblName & " Where (ItMillCode=" & MillCode & " and (ItName='" & Trim(txt(1).Text) & "') and (ItCode<>" & txt(0).Text & "))"
        i = GProcRstOpen(rstModiCheck, X, "O")
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
txt(6) = Cbo(6)
txt(7) = Cbo(7)
txt(17) = Cbo(8)
txt(9) = GProcGetColumnValue("TBLMASTaccount", "acname", txttemp(9), "S", "accode", "N", "", gCn)
If IsDate(mskDt) Then
    txt(16) = CDate(mskDt)
Else
    txt(16) = ""
End If
End Sub
Private Sub FillTempFromTxt()
Cbo(6) = txt(6)
Cbo(7) = txt(7)
Cbo(8) = txt(17)
txttemp(9) = GProcGetColumnValue("TBLMASTaccount", "accode", txt(9), "N", "acname", "N", "", gCn)
mskDt = txt(16)
End Sub
Private Sub ReadFields()
GProcShowRecord Me, Rstbl, MaxNo
FillTempFromTxt
End Sub
'--- Is Releted Mill & Item Record
Private Function IsRelatedRecord() As Boolean
Dim X As String
Dim MillCode As Long
Dim Rs1 As Recordset
MillCode = GProcGetColumnValue("tblMastAccount", "AcName", txttemp(9), "S", "AcCode", "N")
X = "select tblAddLess.Vno,tblAddLess.Vtype,tblAddLess.Vdt from tblAddLess,tblVouSub where " _
& " tblAddLess.Vno=tblVouSub.Vno and tblAddLess.Vtype=tblVouSub.Vtype and tblAddLess.VYear=tblVouSub.VYear and " _
& " tblAddLess.AdMillCode = " & MillCode & " And tblVouSub.vSubItCode=" & txt(0)
GProcRstOpen Rs1, X, "R"
If Rs1.EOF And Rs1.BOF Then
Else
    IsRelatedRecord = True
End If
End Function
