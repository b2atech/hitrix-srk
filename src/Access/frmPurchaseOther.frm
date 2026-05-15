VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form frmPurchaseOther 
   BackColor       =   &H00C2E3FC&
   BorderStyle     =   0  'None
   Caption         =   "Purchase Entry"
   ClientHeight    =   7110
   ClientLeft      =   0
   ClientTop       =   15
   ClientWidth     =   10935
   DrawMode        =   10  'Mask Pen
   FillColor       =   &H000040C0&
   FillStyle       =   0  'Solid
   Icon            =   "frmPurchaseOther.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7110
   ScaleWidth      =   10935
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
      Height          =   375
      Index           =   36
      Left            =   315
      TabIndex        =   93
      Text            =   "36"
      Top             =   750
      Visible         =   0   'False
      Width           =   450
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
      Height          =   375
      Index           =   35
      Left            =   0
      TabIndex        =   92
      Text            =   "35"
      Top             =   435
      Visible         =   0   'False
      Width           =   555
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
      Height          =   375
      Index           =   34
      Left            =   30
      TabIndex        =   91
      Text            =   "34"
      Top             =   0
      Visible         =   0   'False
      Width           =   570
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
      Index           =   12
      ItemData        =   "frmPurchaseOther.frx":058A
      Left            =   3255
      List            =   "frmPurchaseOther.frx":0594
      TabIndex        =   33
      Text            =   "Cbo(12)"
      Top             =   5655
      Width           =   1260
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
      Height          =   375
      Index           =   10
      Left            =   7545
      TabIndex        =   17
      Text            =   "10"
      Top             =   4260
      Width           =   1365
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
      Left            =   9750
      Style           =   1  'Graphical
      TabIndex        =   44
      ToolTipText     =   "Previous"
      Top             =   2130
      Width           =   375
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
      Left            =   10230
      Style           =   1  'Graphical
      TabIndex        =   45
      ToolTipText     =   "Next"
      Top             =   2130
      Width           =   375
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
      Left            =   9750
      Style           =   1  'Graphical
      TabIndex        =   46
      ToolTipText     =   "First"
      Top             =   2610
      Width           =   375
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
      Left            =   10230
      Style           =   1  'Graphical
      TabIndex        =   47
      ToolTipText     =   "Last"
      Top             =   2610
      Width           =   375
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
      Left            =   9750
      Style           =   1  'Graphical
      TabIndex        =   43
      ToolTipText     =   "To Find Record"
      Top             =   3090
      Width           =   855
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
      Left            =   9480
      Style           =   1  'Graphical
      TabIndex        =   48
      Top             =   345
      Visible         =   0   'False
      Width           =   855
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
      Left            =   9750
      Style           =   1  'Graphical
      TabIndex        =   38
      ToolTipText     =   "To Add New Record"
      Top             =   3720
      Width           =   855
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
      Left            =   9750
      Style           =   1  'Graphical
      TabIndex        =   39
      ToolTipText     =   "To Modify Record"
      Top             =   4200
      Width           =   855
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
      Left            =   9750
      Style           =   1  'Graphical
      TabIndex        =   40
      ToolTipText     =   "To Delete Record"
      Top             =   4680
      Width           =   855
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
      Left            =   9525
      Style           =   1  'Graphical
      TabIndex        =   49
      Top             =   465
      Visible         =   0   'False
      Width           =   855
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
      Left            =   9750
      Style           =   1  'Graphical
      TabIndex        =   41
      ToolTipText     =   "To Save Record"
      Top             =   5325
      Width           =   855
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
      Left            =   9750
      Style           =   1  'Graphical
      TabIndex        =   42
      Top             =   5805
      Width           =   855
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
      Left            =   9750
      Style           =   1  'Graphical
      TabIndex        =   50
      ToolTipText     =   "Exit"
      Top             =   6360
      Width           =   855
   End
   Begin VB.TextBox txtVou 
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
      Index           =   7
      Left            =   3120
      TabIndex        =   13
      Text            =   "txtVou(7)"
      Top             =   2775
      Width           =   5760
   End
   Begin VB.TextBox txtVou 
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
      Index           =   8
      Left            =   3120
      TabIndex        =   14
      Text            =   "txtVou(8)"
      Top             =   3210
      Width           =   5760
   End
   Begin VB.TextBox txtVou 
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
      Left            =   3120
      TabIndex        =   15
      Text            =   "txtVou(9)"
      Top             =   3615
      Width           =   5760
   End
   Begin VB.TextBox txt 
      Alignment       =   1  'Right Justify
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   33
      Left            =   345
      TabIndex        =   90
      Text            =   "33"
      Top             =   5805
      Visible         =   0   'False
      Width           =   420
   End
   Begin VB.TextBox txt 
      Alignment       =   1  'Right Justify
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   32
      Left            =   330
      TabIndex        =   89
      Text            =   "32"
      Top             =   5400
      Visible         =   0   'False
      Width           =   420
   End
   Begin VB.TextBox txtVou 
      Height          =   375
      Index           =   10
      Left            =   540
      TabIndex        =   88
      Text            =   "txtVou(10)"
      Top             =   3780
      Visible         =   0   'False
      Width           =   375
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
      Index           =   3
      Left            =   1950
      TabIndex        =   11
      Text            =   "txttemp(3)"
      ToolTipText     =   "Press F3 for Creation"
      Top             =   2055
      Width           =   4245
   End
   Begin VB.TextBox txtCrBalance 
      Enabled         =   0   'False
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
      Left            =   6285
      TabIndex        =   87
      Text            =   "txtCrBalance"
      Top             =   2055
      Width           =   1740
   End
   Begin VB.TextBox txtDrBalance 
      Enabled         =   0   'False
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
      Left            =   6285
      TabIndex        =   86
      Text            =   "txtDrBalance"
      Top             =   1605
      Width           =   1740
   End
   Begin VB.CommandButton cmdLedger 
      Caption         =   "Ledger"
      Height          =   375
      Left            =   8085
      TabIndex        =   85
      Top             =   1605
      Width           =   1065
   End
   Begin VB.TextBox txt 
      Alignment       =   1  'Right Justify
      Height          =   375
      Index           =   31
      Left            =   360
      TabIndex        =   84
      Text            =   "31"
      Top             =   5025
      Visible         =   0   'False
      Width           =   420
   End
   Begin VB.TextBox txt 
      Alignment       =   1  'Right Justify
      Height          =   375
      Index           =   30
      Left            =   60
      TabIndex        =   83
      Text            =   "30"
      Top             =   3750
      Visible         =   0   'False
      Width           =   435
   End
   Begin VB.TextBox txt 
      Alignment       =   1  'Right Justify
      Height          =   375
      Index           =   29
      Left            =   45
      TabIndex        =   82
      Text            =   "29"
      Top             =   3660
      Visible         =   0   'False
      Width           =   435
   End
   Begin VB.TextBox txt 
      Alignment       =   1  'Right Justify
      Height          =   375
      Index           =   28
      Left            =   60
      TabIndex        =   81
      Text            =   "28"
      Top             =   4140
      Visible         =   0   'False
      Width           =   435
   End
   Begin VB.TextBox txt 
      Alignment       =   1  'Right Justify
      Height          =   375
      Index           =   27
      Left            =   -60
      TabIndex        =   80
      Text            =   "27"
      Top             =   4020
      Visible         =   0   'False
      Width           =   435
   End
   Begin VB.TextBox txt 
      Alignment       =   1  'Right Justify
      Height          =   375
      Index           =   26
      Left            =   30
      TabIndex        =   79
      Text            =   "26"
      Top             =   3930
      Visible         =   0   'False
      Width           =   435
   End
   Begin VB.TextBox txt 
      Alignment       =   1  'Right Justify
      Height          =   375
      Index           =   25
      Left            =   60
      TabIndex        =   78
      Text            =   "25"
      Top             =   3870
      Visible         =   0   'False
      Width           =   435
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
      Height          =   375
      Index           =   24
      Left            =   3255
      TabIndex        =   31
      Text            =   "24"
      Top             =   5175
      Width           =   1320
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
      Height          =   375
      Index           =   1
      Left            =   10290
      TabIndex        =   75
      Text            =   "1"
      Top             =   2430
      Visible         =   0   'False
      Width           =   420
   End
   Begin VB.TextBox txt 
      Alignment       =   1  'Right Justify
      Height          =   315
      Index           =   2
      Left            =   10455
      TabIndex        =   66
      Text            =   "2"
      Top             =   1440
      Visible         =   0   'False
      Width           =   525
   End
   Begin VB.TextBox txt 
      Alignment       =   1  'Right Justify
      Height          =   315
      Index           =   5
      Left            =   10455
      TabIndex        =   65
      Text            =   "5"
      Top             =   2640
      Visible         =   0   'False
      Width           =   525
   End
   Begin VB.TextBox txt 
      Alignment       =   1  'Right Justify
      Height          =   315
      Index           =   3
      Left            =   10455
      TabIndex        =   64
      Text            =   "3"
      Top             =   1800
      Visible         =   0   'False
      Width           =   525
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
      Left            =   990
      Style           =   1  'Graphical
      TabIndex        =   54
      Top             =   7110
      Visible         =   0   'False
      Width           =   1455
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
      Left            =   2475
      Style           =   1  'Graphical
      TabIndex        =   53
      ToolTipText     =   "To Add New Record"
      Top             =   6990
      Visible         =   0   'False
      Width           =   1095
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
      Height          =   375
      Index           =   0
      Left            =   1950
      TabIndex        =   1
      Text            =   "0"
      Top             =   540
      Width           =   1350
   End
   Begin VB.TextBox txtVou 
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
      Height          =   375
      Index           =   11
      Left            =   7860
      TabIndex        =   5
      Text            =   "txtVou(11)"
      Top             =   540
      Width           =   1275
   End
   Begin VB.TextBox txt 
      Alignment       =   1  'Right Justify
      Height          =   315
      Index           =   4
      Left            =   10455
      TabIndex        =   63
      Text            =   "4"
      Top             =   2160
      Visible         =   0   'False
      Width           =   525
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
      Index           =   4
      Left            =   1950
      TabIndex        =   9
      Text            =   "txttemp(4)"
      ToolTipText     =   "Press F3 for Creation"
      Top             =   1605
      Width           =   4245
   End
   Begin VB.TextBox txt 
      Height          =   315
      Index           =   6
      Left            =   10455
      TabIndex        =   62
      Text            =   "6"
      Top             =   3000
      Visible         =   0   'False
      Width           =   435
   End
   Begin VB.TextBox txt 
      Alignment       =   1  'Right Justify
      Height          =   315
      Index           =   7
      Left            =   10455
      TabIndex        =   61
      Text            =   "7"
      Top             =   3360
      Visible         =   0   'False
      Width           =   525
   End
   Begin VB.TextBox txt 
      Alignment       =   1  'Right Justify
      Height          =   375
      Index           =   8
      Left            =   10740
      TabIndex        =   52
      Text            =   "8"
      Top             =   540
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.TextBox txt 
      Alignment       =   1  'Right Justify
      Height          =   315
      Index           =   9
      Left            =   10455
      TabIndex        =   60
      Text            =   "9"
      Top             =   3720
      Visible         =   0   'False
      Width           =   525
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
      Height          =   375
      Index           =   11
      Left            =   7545
      TabIndex        =   27
      Text            =   "11"
      Top             =   6525
      Width           =   1365
   End
   Begin VB.TextBox txt 
      Alignment       =   1  'Right Justify
      Height          =   315
      Index           =   12
      Left            =   10455
      TabIndex        =   59
      Text            =   "12"
      Top             =   4080
      Visible         =   0   'False
      Width           =   525
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
      Height          =   375
      Index           =   13
      Left            =   4590
      TabIndex        =   34
      Text            =   "13"
      Top             =   5640
      Width           =   555
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
      Height          =   375
      Index           =   14
      Left            =   7545
      TabIndex        =   23
      Text            =   "14"
      Top             =   5640
      Width           =   1365
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
      Height          =   375
      Index           =   15
      Left            =   3255
      TabIndex        =   29
      Text            =   "15"
      Top             =   4695
      Width           =   1320
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
      Height          =   375
      Index           =   16
      Left            =   3255
      TabIndex        =   37
      Text            =   "16"
      Top             =   6090
      Width           =   1320
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
      Height          =   375
      Index           =   17
      Left            =   7545
      TabIndex        =   25
      Text            =   "17"
      Top             =   6090
      Width           =   1365
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
      Height          =   405
      Index           =   18
      Left            =   7545
      TabIndex        =   19
      Text            =   "18"
      Top             =   4695
      Width           =   1365
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
      Height          =   405
      Index           =   19
      Left            =   7545
      TabIndex        =   21
      Text            =   "19"
      Top             =   5175
      Width           =   1365
   End
   Begin VB.TextBox txt 
      Alignment       =   1  'Right Justify
      Height          =   315
      Index           =   20
      Left            =   10455
      TabIndex        =   58
      Text            =   "20"
      Top             =   4440
      Visible         =   0   'False
      Width           =   525
   End
   Begin VB.TextBox txt 
      Alignment       =   1  'Right Justify
      Height          =   315
      Index           =   21
      Left            =   10455
      TabIndex        =   57
      Text            =   "21"
      Top             =   4800
      Visible         =   0   'False
      Width           =   525
   End
   Begin VB.TextBox txt 
      Alignment       =   1  'Right Justify
      Height          =   315
      Index           =   22
      Left            =   10455
      TabIndex        =   56
      Text            =   "22"
      Top             =   5160
      Visible         =   0   'False
      Width           =   525
   End
   Begin VB.TextBox txt 
      Alignment       =   1  'Right Justify
      Height          =   315
      Index           =   23
      Left            =   10455
      TabIndex        =   55
      Text            =   "23"
      Top             =   5400
      Visible         =   0   'False
      Width           =   525
   End
   Begin MSMask.MaskEdBox mskDt 
      Height          =   375
      Index           =   2
      Left            =   1950
      TabIndex        =   3
      Top             =   1020
      Width           =   1350
      _ExtentX        =   2381
      _ExtentY        =   661
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
   Begin MSMask.MaskEdBox mskVou 
      Height          =   375
      Index           =   12
      Left            =   7860
      TabIndex        =   7
      Top             =   1020
      Width           =   1275
      _ExtentX        =   2249
      _ExtentY        =   661
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
   Begin Crystal.CrystalReport cryReport 
      Bindings        =   "frmPurchaseOther.frx":05A6
      Left            =   0
      Top             =   600
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      WindowControlBox=   -1  'True
      WindowMaxButton =   -1  'True
      WindowMinButton =   -1  'True
      PrintFileType   =   15
      PrintFileLinesPerPage=   60
   End
   Begin VB.Image Image2 
      Height          =   1290
      Left            =   9585
      Picture         =   "frmPurchaseOther.frx":05B7
      Stretch         =   -1  'True
      Top             =   510
      Width           =   1125
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00000080&
      X1              =   5730
      X2              =   5730
      Y1              =   4140
      Y2              =   7020
   End
   Begin VB.Shape Shape1 
      FillColor       =   &H00BDDFD1&
      FillStyle       =   0  'Solid
      Height          =   5115
      Left            =   9600
      Shape           =   4  'Rounded Rectangle
      Top             =   1860
      Width           =   1155
   End
   Begin VB.Shape Shape5 
      BorderColor     =   &H00C00000&
      BorderWidth     =   5
      FillColor       =   &H00C00000&
      Height          =   7095
      Left            =   0
      Top             =   0
      Width           =   10935
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Pa&rty  Name    :"
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
      Left            =   330
      TabIndex        =   10
      Top             =   2055
      Width           =   1665
   End
   Begin VB.Label Label9 
      BackStyle       =   0  'Transparent
      Caption         =   "P&urchase  A/C :"
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
      Left            =   315
      TabIndex        =   8
      Top             =   1605
      Width           =   1620
   End
   Begin VB.Shape Shape3 
      BorderColor     =   &H000000C0&
      FillColor       =   &H000000C0&
      FillStyle       =   0  'Solid
      Height          =   210
      Left            =   0
      Top             =   7695
      Width           =   11415
   End
   Begin VB.Label Label33 
      BackStyle       =   0  'Transparent
      Caption         =   " Purchase &Goods Details    :"
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
      Left            =   330
      TabIndex        =   12
      Top             =   2775
      Width           =   2985
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Vou. &No     :"
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
      Left            =   690
      TabIndex        =   0
      Top             =   540
      Width           =   1110
   End
   Begin VB.Label Label31 
      BackStyle       =   0  'Transparent
      Caption         =   "Exem/Kg.      :"
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
      Left            =   1830
      TabIndex        =   36
      Top             =   6090
      Width           =   1335
   End
   Begin VB.Label Label32 
      BackStyle       =   0  'Transparent
      Caption         =   "Tax Amt    : "
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
      Left            =   6450
      TabIndex        =   22
      Top             =   5640
      Width           =   1095
   End
   Begin VB.Label Label10 
      BackStyle       =   0  'Transparent
      Caption         =   "Sub Amt  :"
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
      Left            =   6450
      TabIndex        =   16
      Top             =   4260
      Width           =   1095
   End
   Begin VB.Label Label11 
      BackStyle       =   0  'Transparent
      Caption         =   "Add          :"
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
      Left            =   6450
      TabIndex        =   18
      Top             =   4695
      Width           =   1095
   End
   Begin VB.Label Label12 
      BackStyle       =   0  'Transparent
      Caption         =   "Less         :"
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
      Left            =   6450
      TabIndex        =   20
      Top             =   5175
      Width           =   1095
   End
   Begin VB.Label Label30 
      BackStyle       =   0  'Transparent
      Caption         =   "Amount    :"
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
      Left            =   6450
      TabIndex        =   26
      Top             =   6525
      Width           =   1095
   End
   Begin VB.Label Label23 
      BackStyle       =   0  'Transparent
      Caption         =   "Booking Dt.      :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   375
      Left            =   10605
      TabIndex        =   77
      Top             =   555
      Visible         =   0   'False
      Width           =   1545
   End
   Begin VB.Label Label17 
      BackStyle       =   0  'Transparent
      Caption         =   "Ass. &Value     :"
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
      Height          =   300
      Left            =   1830
      TabIndex        =   30
      Top             =   5175
      Width           =   1530
   End
   Begin VB.Label lblActionStatus 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "View Record.."
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FFFF&
      Height          =   495
      Left            =   8340
      TabIndex        =   76
      Top             =   0
      Width           =   2475
   End
   Begin VB.Line Line3 
      BorderColor     =   &H00000080&
      X1              =   240
      X2              =   9300
      Y1              =   1485
      Y2              =   1485
   End
   Begin VB.Label Label28 
      BackStyle       =   0  'Transparent
      Caption         =   "ROff          :"
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
      Left            =   6450
      TabIndex        =   24
      Top             =   6090
      Width           =   1095
   End
   Begin VB.Label Label27 
      BackStyle       =   0  'Transparent
      Caption         =   "Landing Cost Per Bag"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   735
      Left            =   12255
      TabIndex        =   74
      Top             =   7800
      Width           =   975
   End
   Begin VB.Label Label26 
      BackStyle       =   0  'Transparent
      Caption         =   "Landing Cost"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   255
      Left            =   12375
      TabIndex        =   73
      Top             =   8160
      Width           =   1455
   End
   Begin VB.Label Label25 
      BackStyle       =   0  'Transparent
      Caption         =   "Other Less"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   255
      Left            =   11415
      TabIndex        =   72
      Top             =   8280
      Width           =   1215
   End
   Begin VB.Label Label22 
      BackStyle       =   0  'Transparent
      Caption         =   "Godown Rent"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   255
      Left            =   11175
      TabIndex        =   71
      Top             =   8160
      Width           =   1575
   End
   Begin VB.Label Label21 
      BackStyle       =   0  'Transparent
      Caption         =   "Freight"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   255
      Left            =   11175
      TabIndex        =   70
      Top             =   8040
      Width           =   1335
   End
   Begin VB.Label Label20 
      BackStyle       =   0  'Transparent
      Caption         =   "Total Amount"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   255
      Left            =   11415
      TabIndex        =   69
      Top             =   7860
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.Label Label18 
      BackStyle       =   0  'Transparent
      Caption         =   "&Exempt. Amt  :"
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
      Height          =   300
      Left            =   1830
      TabIndex        =   28
      Top             =   4695
      Width           =   1440
   End
   Begin VB.Label Label16 
      BackStyle       =   0  'Transparent
      Caption         =   "&TaxType         : "
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
      Left            =   1830
      TabIndex        =   32
      Top             =   5640
      Width           =   1335
   End
   Begin VB.Label Label14 
      BackStyle       =   0  'Transparent
      Caption         =   "Booking No.    :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   375
      Left            =   10500
      TabIndex        =   51
      Top             =   600
      Visible         =   0   'False
      Width           =   1665
   End
   Begin VB.Label Label13 
      BackStyle       =   0  'Transparent
      Caption         =   "Ass.Value"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   255
      Left            =   11055
      TabIndex        =   68
      Top             =   8160
      Width           =   1095
   End
   Begin VB.Label Label8 
      BackStyle       =   0  'Transparent
      Caption         =   "Bill Date  :"
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
      Left            =   6675
      TabIndex        =   6
      Top             =   1020
      Width           =   1215
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "&Bill No     :    "
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
      Left            =   6675
      TabIndex        =   4
      Top             =   540
      Width           =   1095
   End
   Begin VB.Label Label6 
      BackColor       =   &H00C00000&
      Caption         =   "  Other  Purchase"
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
      Height          =   360
      Left            =   0
      TabIndex        =   67
      Top             =   0
      Width           =   11415
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Vou. Da&te  :"
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
      Left            =   675
      TabIndex        =   2
      Top             =   1020
      Width           =   1110
   End
   Begin VB.Label Label29 
      BackStyle       =   0  'Transparent
      Caption         =   "%"
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
      Left            =   5250
      TabIndex        =   35
      Top             =   5640
      Width           =   255
   End
   Begin VB.Shape Shape4 
      BorderColor     =   &H00000080&
      FillColor       =   &H00EDFDFE&
      FillStyle       =   0  'Solid
      Height          =   2895
      Left            =   240
      Top             =   4110
      Width           =   9075
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H00000080&
      FillColor       =   &H00EDFDFE&
      FillStyle       =   0  'Solid
      Height          =   2130
      Left            =   240
      Top             =   420
      Width           =   9075
   End
   Begin VB.Shape Shape6 
      BorderColor     =   &H00000080&
      FillColor       =   &H00EDFDFE&
      FillStyle       =   0  'Solid
      Height          =   1425
      Left            =   240
      Top             =   2655
      Width           =   9075
   End
End
Attribute VB_Name = "frmPurchaseOther"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Const tblName As String = "tblAddLess"
Const IdField As String = "Vno"
Const OrderField As String = "Vdt,Vno"
Const MaxNo As Long = 36
Dim rstbl As Recordset '-- TblAddLess
Dim varBookMark As Variant
Dim FirstTimeFlag As Boolean
Dim rsVou As Recordset '--TblVoucher
Dim OldgFrmTypeStr As String
Dim FormAction As Integer
Dim frmTypeStr As String

Private Sub cbo_LostFocus(Index As Integer)
If Index = 12 And FormAction = vbDataActionAddNew Then '--Tax %
    txt(13) = GProcGetColumnValue("tblMastNarration", "Narration", Cbo(12), "S", "MastTaxRate", "N")
    txt(13) = Format(txt(13), FStr)
End If
End Sub

Private Sub cmdBtn_Click(Index As Integer)
On Error GoTo ErrorRoutine
Dim Hank As String
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
    gClsSearch.SearchMultiField "tblAddLess,TblMastAccount", "Vno,AcName,Vdt", Array("Vno", "Party", "Vdt"), Array(txt(0).Width, 4000, 1080), " tblAddLess.AdAcCrCode=TblMastAccount.AcCode and vtype='" & frmTypeStr & "' and VYear=" & gCYear, "", "Vno,Vdt", 0, 0
    If gClsSearch.SearchMultiRetCol(0) <> "" Then
        rstbl.MoveFirst
        rstbl.Find "Vno = " & gClsSearch.SearchMultiRetCol(0)
        gClsSearch.SearchMultiRetCol(0) = ""
        ReadFields
    End If
Case 6 'Add
    FormAction = vbDataActionAddNew
    Call GProcActivateControls(Me, True)
    SetControlEd
    Call GProcSetButtons(Me, vbDataActionAddNew)
    Call GProcClearForm(Me, rstbl, MaxNo, True)
    If FirstTimeFlag = True Then
        FirstTimeFlag = False
    Else
        txt(0).SetFocus
    End If
    txt(0) = GProcGenerateId("TblAddLess", "Vno", "Vtype='" & frmTypeStr & "' and VYear=" & gCYear)
    mskDt(2) = GProcGenerateVDt("tblAddLess", "Vdt", "Vtype='" & frmTypeStr & "' and VYear=" & gCYear)
    FillCombo
Case 7 'Modify
    FormAction = vbDataActionUpdate
    Hank = ""
    'Call GProcChkPass(frmTypeStr, FormAction, txt(36), txt(2), txtTemp(4), gUserName, Hank)
    'If gChkPassTrueFalse = True Then
        'If mURecModi = True Then
        '    SearchHelp
        '    DoEvents
            If IsRelatedRecord = False Then
                FormAction = vbDataActionUpdate
                Call GProcActivateControls(Me, True)
                SetControlEd
                Call GProcSetButtons(Me, vbDataActionUpdate)
                txt(0).Enabled = False
                mskDt(2).SetFocus
            End If
        'End If
    'Else
    '   FormAction = vbDataActionClose
    'End If
Case 8 'Delete
    FormAction = vbDataActionDelete
    Hank = ""
'    Call GProcChkPass(frmTypeStr, FormAction, txt(36), txt(2), txtTemp(4), gUserName, Hank)
'    If gChkPassTrueFalse = True Then
        'If mURecDel = True Then
          If IsRelatedRecord = False Then
            If MsgBox("Do You Really Delete This Record?", vbQuestion + vbYesNo + vbDefaultButton1, "Confirm Delete") = vbYes Then
                If rstbl.RecordCount > 1 Then
                    rstbl.Delete
                    '-- Delete Related Record
                    DeleteRelatedRecord
                    cmdBtn_Click (1)
                ElseIf rstbl.RecordCount = 1 Then
                    rstbl.Delete
                    '-- Delete Related Record
                    DeleteRelatedRecord
                    Call GProcActivateControls(Me, False)
                    Call GProcSetButtons(Me, vbDataActionCancel)
                    Form_Load
                End If
            End If
        End If
 '   End If
Case 9 'Print
Case 10 'Exit
'    DoEvents
    Unload Me
    Exit Sub
Case 11 'Label
    'frmLabel.Show 1
Case 12 'Optional
Case 13 'Save
    If ValidateData = True Then
       ' If FormAction = vbDataActionAddNew Then txt(0) = GProcGenerateId(tblName, IdField, "vtype='" & FrmTypeStr & "' and VYear=" & gCYear)
        CalAmount
        FillTxtFromTemp
        '-- Save Record
        GProcSaveRecord Me, rstbl, FormAction, MaxNo
        SaveRelatedRecord
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
Private Sub SetControlEd()
'txt(10).Enabled = False '--Sub Amt
txt(14).Enabled = False '--TaxAmt
'txt(17).Enabled = False '--Roff
txt(11).Enabled = False '-- Amt
txt(24).Enabled = False '-- Ass Val
txtCrBalance.Enabled = False
txtDrBalance.Enabled = False
End Sub
Private Sub DeleteRelatedRecord()
'-- Voucher
gCn.Execute "delete from tblVoucher where vtype='" & frmTypeStr & "' and vno=" & txt(0) & " and " _
& " VYear=" & gCYear
End Sub
Private Sub SaveRelatedRecord()
Dim i As Long
Dim CrCode As Long
Dim DrCode As Long
Dim TaxAcCode As Long
Dim X As String
Dim BillDt
'--- Account Code
CrCode = GProcGetColumnValue("tblMastAccount", "AcName", txtTemp(3), "S", "AcCode", "N")
DrCode = GProcGetColumnValue("tblMastAccount", "AcName", txtTemp(4), "S", "AcCode", "N")
TaxAcCode = GProcGetColumnValue("tblMastNarration", "Narration", Cbo(12), "S", "MastTaxPurAcCode", "N")
'-- Delete Records
DeleteRelatedRecord
'======================Voucher Creation
If IsDate(mskVou(12)) Then
    BillDt = mskVou(12)
Else
    BillDt = ""
End If
'--
If GProcGetColumnValue("tblMastAccount", "AcCode", CStr(TaxAcCode), "N", "AcName", "S") = "-" Then
    '-- Bill Amt
    GProcCreateVoucher FormAction, txt(0), 1, txt(2), txt(1), CrCode, DrCode, txt(11), txtVou(11), BillDt, , txtVou(7), txtVou(8), txtVou(9), txtVou(10)
    GProcCreateVoucher FormAction, txt(0), 2, txt(2), txt(1), DrCode, CrCode, -txt(11), txtVou(11), BillDt, , txtVou(7), txtVou(8), txtVou(9), txtVou(10)
Else '-- If Tax A/c
    '-- Bill -Tax Amt
    GProcCreateVoucher FormAction, txt(0), 1, txt(2), txt(1), CrCode, DrCode, txt(11) - txt(14), txtVou(11), BillDt, , txtVou(7), txtVou(8), txtVou(9), txtVou(10)
    GProcCreateVoucher FormAction, txt(0), 2, txt(2), txt(1), DrCode, CrCode, -(txt(11) - txt(14)), txtVou(11), BillDt, , txtVou(7), txtVou(8), txtVou(9), txtVou(10)
    '-- Tax Amt
    If CDbl(txt(14)) > 0 Then
        GProcCreateVoucher FormAction, txt(0), 3, txt(2), txt(1), CrCode, TaxAcCode, txt(14), txtVou(11), BillDt, , txtVou(7), txtVou(8), txtVou(9), txtVou(10)
        GProcCreateVoucher FormAction, txt(0), 4, txt(2), txt(1), TaxAcCode, CrCode, -txt(14), txtVou(11), BillDt, , txtVou(7), txtVou(8), txtVou(9), txtVou(10)
    End If
End If
'--Voucher Sub,Outstanding
End Sub
Private Sub cmdLedger_Click()
On Error GoTo ErrorHandler
frmRptAccount.txt(1) = txtTemp(4)
frmMain.mnuRptRSubAccLeg_Click (0)
Exit Sub
ErrorHandler:
    GProcErrorHandler
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
Private Sub Form_Load()
    Dim X As String
    Me.Top = 300
    Me.Left = Screen.Width / 2 - Me.Width / 2
    '--- TblAddLess
    frmTypeStr = gfrmTypeStr
    X = "Select * from " & tblName & " where Vtype='" & frmTypeStr & "' and VYear=" & gCYear
    If Trim(OrderField) <> "" Then X = X & " order by " & OrderField
    GProcRstOpen rstbl, X, "O"
    FirstTimeFlag = True
    '---Voucher
    SetRelatedRS
    '---- Button Visiblity Set As Per User Previlage
    GProcSetButtonVisiblity Me
    cmdBtn_Click (6)
End Sub
Private Sub SetRelatedRS()
Dim X As String
Set rsVou = New Recordset
'---TblVoucher
X = "Select * from tblVoucher where Vtype='" & frmTypeStr & "' and vno=" & txt(0) & " and VYear=" & gCYear & " order by VCtrNo"
GProcRstOpen rsVou, X, "O"
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
Private Sub mskDt_GotFocus(Index As Integer)
GProcSelectBox Me.ActiveControl
End Sub
Private Sub mskVou_GotFocus(Index As Integer)
GProcSelectBox Me.ActiveControl
End Sub
Private Sub mskVou_LostFocus(Index As Integer)
If Index = 2 Then
    If GProcIsDateValid(mskDt(Index)) = False Then
        mskDt(Index).SetFocus
        Exit Sub
    End If
End If
If IsDate(mskVou(Index)) = False Then mskVou(Index) = "__/__/____"
End Sub
Private Sub ShowAccountBalance()
Dim DrAcBal  As Double
Dim CrAcBal As Double
If IsDate(mskDt(2)) Then
    DrAcBal = GProcGetAccountBalance(txtTemp(4), mskDt(2))
    CrAcBal = GProcGetAccountBalance(txtTemp(3), mskDt(2))
    txtDrBalance = IIf(DrAcBal > 0, DrAcBal & " Cr", Abs(DrAcBal) & " Dr")
    txtCrBalance = IIf(CrAcBal > 0, CrAcBal & " Cr", Abs(CrAcBal) & " Dr")
End If
End Sub
Private Sub mskDt_LostFocus(Index As Integer)
If Index = 2 Then
    If GProcIsDateValid(mskDt(Index)) = False Then
        mskDt(Index).SetFocus
        Exit Sub
    End If
End If
If Index = 2 Then ShowAccountBalance
End Sub
Private Sub txt_GotFocus(Index As Integer)
GProcSelectBox Me.ActiveControl
End Sub

Private Sub txtVou_GotFocus(Index As Integer)
GProcSelectBox Me.ActiveControl
End Sub
Private Sub txtVou_LostFocus(Index As Integer)
SaveBtnEd
End Sub
Private Sub txtVou_KeyPress(Index As Integer, KeyAscii As Integer)
If KeyAscii = 13 Then Exit Sub
KeyAscii = GProcValidateKey(rsVou, Index, KeyAscii, txtVou(Index).text)
End Sub

Private Sub txt_LostFocus(Index As Integer)
GProcCheckForNumber Me, rstbl, MaxNo
SaveBtnEd
CalAmount
End Sub
Private Sub txt_KeyPress(Index As Integer, KeyAscii As Integer)
If KeyAscii = 13 Then Exit Sub
KeyAscii = GProcValidateKey(rstbl, Index, KeyAscii, txt(Index).text)
End Sub
Private Sub txttemp_GotFocus(Index As Integer)
GProcSelectBox Me.ActiveControl
End Sub
Private Sub txttemp_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
Dim GName As String
If KeyCode = vbKeyF3 Then
    Select Case Index
        Case 3 '-- Party
            GName = GProcGetColumnValue("TblMastGroup", "AgCode", 90015, "N", "AgName", "S")
            GProcShowForm frmMastAccount, frmMain.mnuMstfrm(0), Array(frmMastAccount.txtTemp(2)), Array(GName)
        Case 4 '-- Purchase A/c
            GName = GProcGetColumnValue("TblMastGroup", "AgCode", 90025, "N", "AgName", "S")
            GProcShowForm frmMastAccount, frmMain.mnuMstfrm(0), Array(frmMastAccount.txtTemp(2)), Array(GName)
        Case 12 '--TaxType
            'OldgFrmTypeStr = gfrmTypeStr
            GProcShowForm frmMastNarrationTax, frmMain.mnuMstfrm(6)
            'gfrmTypeStr = OldgFrmTypeStr
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
    Case 3  '-- Party   Alias Account
        gClsSearch.SearchMultiField "QryMastaccount", "Account,AcName,AcCity,AgName", Array("Party", "AcName", "City", "Group"), Array(3000, 0, 2000, 3000), "(Agcode=90015 or GpCode=90015 or Agcode=90017 or GpCode=90017 or Agcode=90029 or GpCode=90029 )", Chr(KeyAscii), "acName", LeftPos, TopPos, True
        KeyAscii = 0
        If gClsSearch.SearchMultiRetCol(0) <> "" And gClsSearch.SearchMultiRetCol(1) = "" Then
            txtTemp(Index) = gClsSearch.SearchMultiRetCol(0)
        Else
            If gClsSearch.SearchMultiRetCol(1) <> "" Then txtTemp(Index) = gClsSearch.SearchMultiRetCol(1)
        End If
        'gClsSearch.SearchMultiField "tblMastaccount,tblMastGroup", "AcName", Array("Party"), Array(txtTemp(Index).Width), " tblMastaccount.AgCode=tblMastGroup.AgCode and " _
        '& " (tblMastGroup.Agcode=15 or GpCode=15 or tblMastGroup.Agcode=17 or GpCode=17 or tblMastGroup.Agcode=29 or GpCode=29 )", Chr(KeyAscii), "acName", LeftPos, TopPos, True
    Case 4 '-- Purchase A/c  Alias Account
        gClsSearch.SearchMultiField "QryMastaccount", "Account,AcName,AcCity,AgName", Array("Purchase Account", "AcName", "City", "Group"), Array(3000, 0, 2000, 3000), "(AgCode=90025 or GpCode=90025)", Chr(KeyAscii), "AcName", LeftPos, TopPos, True
        KeyAscii = 0
        If gClsSearch.SearchMultiRetCol(0) <> "" And gClsSearch.SearchMultiRetCol(1) = "" Then
            txtTemp(Index) = gClsSearch.SearchMultiRetCol(0)
        Else
            If gClsSearch.SearchMultiRetCol(1) <> "" Then txtTemp(Index) = gClsSearch.SearchMultiRetCol(1)
        End If
        'gClsSearch.SearchMultiField "tblMastaccount,tblMastGroup", "AcName", Array("Purchase Account"), Array(txtTemp(Index).Width), " tblMastaccount.AgCode=tblMastGroup.AgCode and (tblMastGroup.AgCode=25 or GpCode=25)", Chr(KeyAscii), "AcName", LeftPos, TopPos, True
    Case 12 '-- Tax Type
        gClsSearch.SearchMultiField "tblMastNarration", "Narration,MastTaxRate", Array("Tax Type", "TaxRate"), Array(txtTemp(Index).Width, 0), " NarrType='T'", Chr(KeyAscii), "Narration", 0, 0, True
        KeyAscii = 0
        If gClsSearch.SearchMultiRetCol(0) <> "" Then
            txtTemp(Index) = gClsSearch.SearchMultiRetCol(0) '--Tax
            If FormAction = vbDataActionAddNew Then txt(13) = gClsSearch.SearchMultiRetCol(1)  '--Tax %
        End If
End Select
End Sub
Private Sub txtTemp_LostFocus(Index As Integer)
SaveBtnEd
ShowAccountBalance
End Sub
Private Sub FillCombo()
Call GProcFillCombo(Cbo(12), "select Narration from tblMastNarration where NarrType ='T' order by Narration")
'Cbo(12) = GProcGetColumnValue("tblMastNarration", "NarrCode", CStr(gTaxCodeP), "N", "Narration", "S")
End Sub
Private Sub SaveBtnEd()
'--- Vno,vdt,Goods,Party,A/c,Amount
If CDbl(txt(0)) = 0 Or GProcIsDateValid(mskDt(2)) = False Or txtVou(7) = "" Or txtTemp(3) = "" Or txtTemp(4) = "" Or CDbl(txt(11)) = 0 Then
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
Dim rstAddCheck As Recordset
Dim rstModiCheck As Recordset
    '---- Required Data
    '--- Vno
    If CLng(txt(0)) = 0 Then
        MsgBox "Check Voucher No.", vbInformation + vbOKOnly, Me.Caption
        txt(0).SetFocus
        Exit Function
    End If
    '--- Vdt
    If GProcIsDateValid(mskDt(2)) = False Then
        MsgBox "Check Voucher Date.", vbInformation + vbOKOnly, Me.Caption
        mskDt(2).SetFocus
        Exit Function
    End If
    '---Tax Type
    'If txttemp(12) = "" Then txttemp(12) = "-"
    If Cbo(12) = "" Then
        MsgBox "Check Tax Type.", vbInformation + vbOKOnly, Me.Caption
        Cbo(12).SetFocus
        Exit Function
    End If
    '-- Tax %
    If CDbl(txt(13)) > 0 And Cbo(12) = "Tax Free" Then
        MsgBox "Check Tax %.", vbInformation + vbOKOnly, Me.Caption
        txt(13).SetFocus
        Exit Function
    End If
    '---- Party Account
    If txtTemp(3) = "" Then
        MsgBox "Check Party.", vbInformation + vbOKOnly, Me.Caption
        txtTemp(3).SetFocus
        Exit Function
    End If
        '--Create Party
        If GProcGetColumnValue("tblMastAccount", "AcName", txtTemp(3), "S", "AcCode", "N") = 0 Then
           GName = GProcGetColumnValue("TblMastGroup", "AgCode", 90015, "N", "AgName", "S")
           GProcShowForm frmMastAccount, frmMain.mnuMstfrm(0), Array(frmMastAccount.txt(1), frmMastAccount.txtTemp(2)), Array(txtTemp(3), GName)
           Exit Function
        End If
    '---- Purchase Account
    If txtTemp(4) = "" Then
        MsgBox "Check Purchase Account.", vbInformation + vbOKOnly, Me.Caption
        txtTemp(4).SetFocus
        Exit Function
    End If
        '--Create Party
        If GProcGetColumnValue("tblMastAccount", "AcName", txtTemp(4), "S", "AcCode", "N") = 0 Then
           GName = GProcGetColumnValue("TblMastGroup", "AgCode", 90025, "N", "AgName", "S")
           GProcShowForm frmMastAccount, frmMain.mnuMstfrm(0), Array(frmMastAccount.txt(1), frmMastAccount.txtTemp(2)), Array(txtTemp(4), GName)
           Exit Function
        End If
    '--- Good Details
    If txtVou(7) = "" Then
        MsgBox "Check Goods Details.", vbInformation + vbOKOnly, Me.Caption
        txtVou(7).SetFocus
        Exit Function
    End If
    '---- Amount
    If CDbl(txt(11)) <= 0 Then
        MsgBox "Check Voucher Amount.", vbInformation + vbOKOnly, Me.Caption
        Exit Function
    End If
    '----Duplicate Records
    '----Addition Case
    If FormAction = vbDataActionAddNew Then
        X = "Select vno From " & tblName & " Where vno=" & CLng(txt(0)) & " and Vtype='" & frmTypeStr & "' and VYear=" & gCYear
        i = GProcRstOpen(rstAddCheck, X, "R")
        If i > 0 Then
            MsgBox "Record is Available of Voucher No. " & CLng(txt(0)), vbCritical + vbOKOnly, "Duplicate Record"
            txt(0).SetFocus
            Exit Function
        End If
        rstAddCheck.Close
    End If
    ValidateData = True
End Function
Private Sub FillTxtFromTemp()
txt(1) = frmTypeStr '--VType
If IsDate(mskVou(12)) = False Then mskVou(12) = "__/__/____" '--BillDt
If IsDate(mskDt(2)) Then '--Vdt
    txt(2) = CDate(mskDt(2))
Else
    txt(2) = ""
End If
'--- Bill No
txt(36) = txtVou(11)
'-- Cr Code
txt(3) = GProcGetColumnValue("tblMastAccount", "AcName", txtTemp(3), "S", "AcCode", "N")
'-- Dr Code
txt(4) = GProcGetColumnValue("tblMastAccount", "AcName", txtTemp(4), "S", "AcCode", "N")
txt(5) = GProcGetColumnValue("tblMastAccount", "AcName", "-", "S", "AcCode", "N")
'-- TaxType
'If txttemp(12) = "" Then txttemp(12) = "-"
txt(12) = GProcGetColumnValue("tblMastNarration", "Narration", Cbo(12), "S", "NarrCode", "N")
txt(23) = gCYear
ShowAccountBalance
End Sub
Private Sub FillTempFromTxt()
Dim X As String
FillCombo
If IsDate(txt(2)) Then '--VDt
    mskDt(2) = txt(2)
Else
    mskDt(2) = "__/__/____"
End If
'-- Cr Code
txtTemp(3) = GProcGetColumnValue("tblMastAccount", "AcCode", txt(3), "N", "AcName", "S")
'-- Dr Code
txtTemp(4) = GProcGetColumnValue("tblMastAccount", "AcCode", txt(4), "N", "AcName", "S")
'-- TaxType
Cbo(12) = GProcGetColumnValue("tblMastNarration", "NarrCode", txt(12), "N", "Narration", "S")
'-- Fill Related Record
FillRelatedRecord
ShowAccountBalance
End Sub
Private Sub FillRelatedRecord()
Dim RowIndex As Long
Dim ColIndex As Long
'--- Fill Related Control
SetRelatedRS
txtVou(11) = rsVou.Fields("vBillNo") & vbNullString     '--Bill No
txtVou(7) = rsVou.Fields("VNar1") & vbNullString '--Narration
txtVou(8) = rsVou.Fields("VNar2") & vbNullString '--
txtVou(9) = rsVou.Fields("VNar3") & vbNullString '--
txtVou(10) = rsVou.Fields("VNar4") & vbNullString '--
If IsDate(rsVou.Fields("vBillDt")) Then '--Bill Date
    mskVou(12) = Format(rsVou.Fields("vBilldt"), "dd/mm/yyyy")
Else
    mskVou(12) = "__/__/____"
End If
End Sub
Private Sub ReadFields()
FormAction = vbDataActionClose
GProcShowRecord Me, rstbl, MaxNo
FillTempFromTxt
End Sub
Private Sub CalAmount()
Dim TaxAmt As Double
Dim SubAmt As Double
Dim Amt As Double
Dim ROff As Double
Dim AssValue As Double
Dim RowIndex As Integer
Dim TotBag As Long
Dim LCost As Double
SubAmt = txt(10)
'--- Subamt+Add -Less-Exempt.Amt
AssValue = SubAmt + CDbl(txt(18)) - CDbl(txt(19)) - CDbl(txt(15))
TaxAmt = AssValue * CDbl(txt(13)) / 100
'--Roff
ROff = txt(17) '= 'CDbl(GProcMakeRounding(Amt) - Amt)
'-- SubAmt+Tax Amt +Add -Less+roff
Amt = SubAmt + TaxAmt + CDbl(txt(18)) - CDbl(txt(19)) + ROff
'--- Landing Cost =Bill+Freight+Rent +add-Less
LCost = Amt + CDbl(txt(25)) + CDbl(txt(26)) + CDbl(txt(27)) - CDbl(txt(28))
txt(10) = SubAmt
txt(14) = TaxAmt
txt(24) = AssValue
'txt(17) = Roff
txt(11) = Amt
txt(29) = LCost
If TotBag > 0 Then txt(30) = txt(29) / TotBag
GProcCheckForNumber Me, rstbl, MaxNo
End Sub
'--- Is Related Record
Private Function IsRelatedRecord() As Boolean
Dim RsG As Recordset
Dim X As String
'-- Audited
If GProcISAudited(txt(0), txt(1), txt(23)) = True Then
    MsgBox "Record is Audited.You can not Modify/Delete this Record.", vbCritical + vbOKOnly, App.Title
    IsRelatedRecord = True
    Exit Function
End If
'-- Is Gate Pass
X = "Select Vno,Vdt from tblVouSub where Vtype='GP' and " _
& " VSubPItCtrlNo in (select VSubItCtrlNo from tblVouSub where vtype='" & frmTypeStr & "' and vno=" & txt(0) & " and " _
& " VYear=" & gCYear & ")"
GProcRstOpen RsG, X, "R"
If RsG.EOF Then
Else
    IsRelatedRecord = True
    MsgBox "You can not Modify/Delete this record.Related Gate Pass No. " & RsG.Fields(0) & " is exists.", vbCritical + vbOKOnly, App.Title
    Exit Function
End If
RsG.Close
End Function



