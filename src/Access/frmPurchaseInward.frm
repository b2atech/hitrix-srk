VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form frmPurchaseInward 
   BackColor       =   &H00C2E3FC&
   BorderStyle     =   0  'None
   Caption         =   "Goods Inward"
   ClientHeight    =   7230
   ClientLeft      =   0
   ClientTop       =   15
   ClientWidth     =   10845
   FillColor       =   &H000040C0&
   FillStyle       =   0  'Solid
   Icon            =   "frmPurchaseInward.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7230
   ScaleWidth      =   10845
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton CmdBagDetails 
      Caption         =   "Cartoon &Bag Details"
      Height          =   750
      Left            =   9690
      TabIndex        =   100
      Top             =   2370
      Width           =   855
   End
   Begin VB.CheckBox Chk 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00ECFFFF&
      Caption         =   "Is &Taxable Inward ?"
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
      Index           =   40
      Left            =   3150
      TabIndex        =   22
      Top             =   5880
      Width           =   2280
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
      Index           =   40
      Left            =   30
      TabIndex        =   99
      Text            =   "40"
      Top             =   1770
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
      Index           =   39
      Left            =   15
      TabIndex        =   98
      Text            =   "39"
      Top             =   1335
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
      Index           =   38
      Left            =   15
      TabIndex        =   97
      Text            =   "38"
      Top             =   915
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
      Index           =   37
      Left            =   75
      TabIndex        =   96
      Text            =   "37"
      Top             =   435
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
      Index           =   36
      Left            =   285
      TabIndex        =   95
      Text            =   "36"
      Top             =   735
      Visible         =   0   'False
      Width           =   450
   End
   Begin VB.TextBox txttmpTotBag 
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
      Left            =   7785
      TabIndex        =   94
      Text            =   "txttmpTotBag"
      Top             =   6210
      Width           =   1365
   End
   Begin VB.TextBox txttmpTotWt 
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
      Left            =   7785
      TabIndex        =   91
      Text            =   "txttmpTotWt"
      Top             =   6675
      Width           =   1365
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
      Height          =   375
      Index           =   35
      Left            =   1635
      TabIndex        =   24
      Text            =   "35"
      ToolTipText     =   "Press F3 for Creation"
      Top             =   4860
      Width           =   3810
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
      Left            =   3330
      TabIndex        =   90
      Text            =   "34"
      Top             =   6825
      Visible         =   0   'False
      Width           =   330
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
      Left            =   4935
      TabIndex        =   89
      Text            =   "33"
      Top             =   7110
      Visible         =   0   'False
      Width           =   420
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
      Height          =   375
      Index           =   32
      Left            =   1635
      TabIndex        =   26
      Text            =   "32"
      Top             =   5280
      Width           =   1485
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
      Index           =   31
      Left            =   4230
      TabIndex        =   88
      Text            =   "31"
      Top             =   6945
      Visible         =   0   'False
      Width           =   300
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
      Index           =   30
      Left            =   1665
      TabIndex        =   87
      Text            =   "30"
      Top             =   6975
      Visible         =   0   'False
      Width           =   330
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
      Index           =   29
      Left            =   1110
      TabIndex        =   86
      Text            =   "29"
      Top             =   7095
      Visible         =   0   'False
      Width           =   330
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
      Index           =   28
      Left            =   1920
      TabIndex        =   85
      Text            =   "28"
      Top             =   6960
      Visible         =   0   'False
      Width           =   330
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
      Index           =   27
      Left            =   1635
      TabIndex        =   34
      Text            =   "27"
      Top             =   6585
      Width           =   1485
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
      Index           =   26
      Left            =   1635
      TabIndex        =   32
      Text            =   "26"
      Top             =   6135
      Width           =   1485
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
      Index           =   25
      Left            =   1635
      TabIndex        =   30
      Text            =   "25"
      Top             =   5715
      Width           =   1485
   End
   Begin VB.CommandButton cmdLedger 
      Caption         =   "Ledger"
      Height          =   375
      Left            =   6450
      TabIndex        =   84
      Top             =   1680
      Width           =   1035
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
      Left            =   3675
      TabIndex        =   83
      Text            =   "txtDrBalance"
      Top             =   6945
      Visible         =   0   'False
      Width           =   660
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
      Left            =   5130
      TabIndex        =   82
      Text            =   "txtCrBalance"
      Top             =   1695
      Width           =   1275
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
      Left            =   2220
      TabIndex        =   81
      Text            =   "txttemp(3)"
      ToolTipText     =   "Press F3 for Creation"
      Top             =   6885
      Visible         =   0   'False
      Width           =   600
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
      Left            =   4005
      TabIndex        =   80
      Text            =   "24"
      Top             =   6930
      Visible         =   0   'False
      Width           =   330
   End
   Begin VB.TextBox txtGrid 
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   600
      TabIndex        =   17
      Top             =   3480
      Visible         =   0   'False
      Width           =   1035
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
      Left            =   10875
      TabIndex        =   78
      Text            =   "1"
      Top             =   2430
      Width           =   420
   End
   Begin VB.TextBox txt 
      Alignment       =   1  'Right Justify
      Height          =   315
      Index           =   2
      Left            =   11040
      TabIndex        =   67
      Text            =   "2"
      Top             =   1440
      Visible         =   0   'False
      Width           =   525
   End
   Begin VB.TextBox txt 
      Alignment       =   1  'Right Justify
      Height          =   315
      Index           =   5
      Left            =   11040
      TabIndex        =   66
      Text            =   "5"
      Top             =   2640
      Width           =   525
   End
   Begin VB.TextBox txt 
      Alignment       =   1  'Right Justify
      Height          =   315
      Index           =   3
      Left            =   11040
      TabIndex        =   65
      Text            =   "3"
      Top             =   1800
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
      Left            =   8550
      Style           =   1  'Graphical
      TabIndex        =   56
      Top             =   8055
      Visible         =   0   'False
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
      Left            =   9315
      Style           =   1  'Graphical
      TabIndex        =   55
      ToolTipText     =   "To Add New Record"
      Top             =   7980
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.CommandButton cmdBtn 
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
      Left            =   9690
      Style           =   1  'Graphical
      TabIndex        =   47
      ToolTipText     =   "Exit"
      Top             =   6105
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
      Left            =   9690
      Style           =   1  'Graphical
      TabIndex        =   39
      Top             =   5460
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
      Left            =   9690
      Style           =   1  'Graphical
      TabIndex        =   38
      ToolTipText     =   "To Save Record"
      Top             =   4980
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
      Left            =   8490
      Style           =   1  'Graphical
      TabIndex        =   40
      Top             =   6900
      Visible         =   0   'False
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
      Left            =   9690
      Style           =   1  'Graphical
      TabIndex        =   37
      ToolTipText     =   "To Delete Record"
      Top             =   4320
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
      Left            =   9690
      Style           =   1  'Graphical
      TabIndex        =   36
      ToolTipText     =   "To Modify Record"
      Top             =   3840
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
      Left            =   9690
      Style           =   1  'Graphical
      TabIndex        =   35
      ToolTipText     =   "To Add New Record"
      Top             =   3360
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
      Left            =   7680
      Style           =   1  'Graphical
      TabIndex        =   41
      Top             =   6945
      Visible         =   0   'False
      Width           =   855
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
      Left            =   9690
      Style           =   1  'Graphical
      TabIndex        =   42
      ToolTipText     =   "To Find Record"
      Top             =   1755
      Width           =   855
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
      Left            =   10170
      Style           =   1  'Graphical
      TabIndex        =   46
      ToolTipText     =   "Last"
      Top             =   1275
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
      Left            =   9690
      Style           =   1  'Graphical
      TabIndex        =   45
      ToolTipText     =   "First"
      Top             =   1275
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
      Left            =   10170
      Style           =   1  'Graphical
      TabIndex        =   44
      ToolTipText     =   "Next"
      Top             =   795
      Width           =   375
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
      Left            =   2955
      TabIndex        =   1
      Text            =   "0"
      Top             =   645
      Width           =   1245
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
      Left            =   9690
      Style           =   1  'Graphical
      TabIndex        =   43
      ToolTipText     =   "Previous"
      Top             =   795
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
      Index           =   5
      Left            =   1305
      TabIndex        =   9
      Text            =   "txttemp(5)"
      ToolTipText     =   "Press F3 for Creation"
      Top             =   1695
      Width           =   3780
   End
   Begin VB.TextBox txt 
      Alignment       =   1  'Right Justify
      Height          =   315
      Index           =   4
      Left            =   11040
      TabIndex        =   64
      Text            =   "4"
      Top             =   2160
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
      Left            =   1305
      TabIndex        =   15
      Text            =   "txttemp(4)"
      ToolTipText     =   "Press F3 for Creation"
      Top             =   2145
      Width           =   5220
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
      Height          =   375
      Index           =   6
      Left            =   5880
      TabIndex        =   5
      Text            =   "6"
      Top             =   645
      Width           =   1305
   End
   Begin VB.TextBox txt 
      Alignment       =   1  'Right Justify
      Height          =   315
      Index           =   7
      Left            =   11040
      TabIndex        =   63
      Text            =   "7"
      Top             =   3360
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
      Height          =   375
      Index           =   8
      Left            =   7320
      TabIndex        =   11
      Text            =   "8"
      Top             =   195
      Visible         =   0   'False
      Width           =   525
   End
   Begin VB.TextBox txt 
      Alignment       =   1  'Right Justify
      Height          =   315
      Index           =   9
      Left            =   11040
      TabIndex        =   62
      Text            =   "9"
      Top             =   3720
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
      Index           =   10
      Left            =   7785
      TabIndex        =   19
      Text            =   "10"
      Top             =   5070
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
      Index           =   11
      Left            =   7785
      TabIndex        =   21
      Text            =   "11"
      Top             =   5610
      Width           =   1365
   End
   Begin VB.TextBox txt 
      Alignment       =   1  'Right Justify
      Height          =   315
      Index           =   12
      Left            =   11040
      TabIndex        =   61
      Text            =   "12"
      Top             =   4080
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
      Left            =   4455
      TabIndex        =   48
      Text            =   "13"
      Top             =   7005
      Visible         =   0   'False
      Width           =   405
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
      Left            =   4335
      TabIndex        =   51
      Text            =   "14"
      Top             =   6915
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
      Index           =   15
      Left            =   2160
      TabIndex        =   50
      Text            =   "15"
      Top             =   7065
      Visible         =   0   'False
      Width           =   570
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
      Left            =   2775
      TabIndex        =   49
      Text            =   "16"
      Top             =   6990
      Visible         =   0   'False
      Width           =   330
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
      Left            =   3120
      TabIndex        =   54
      Text            =   "17"
      Top             =   6990
      Visible         =   0   'False
      Width           =   330
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
      Left            =   3525
      TabIndex        =   52
      Text            =   "18"
      Top             =   6915
      Visible         =   0   'False
      Width           =   285
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
      Left            =   3885
      TabIndex        =   53
      Text            =   "19"
      Top             =   6915
      Visible         =   0   'False
      Width           =   285
   End
   Begin VB.TextBox txt 
      Alignment       =   1  'Right Justify
      Height          =   315
      Index           =   20
      Left            =   11040
      TabIndex        =   60
      Text            =   "20"
      Top             =   4440
      Visible         =   0   'False
      Width           =   525
   End
   Begin VB.TextBox txt 
      Alignment       =   1  'Right Justify
      Height          =   315
      Index           =   21
      Left            =   11040
      TabIndex        =   59
      Text            =   "21"
      Top             =   4800
      Visible         =   0   'False
      Width           =   525
   End
   Begin VB.TextBox txt 
      Alignment       =   1  'Right Justify
      Height          =   315
      Index           =   22
      Left            =   11040
      TabIndex        =   58
      Text            =   "22"
      Top             =   5160
      Visible         =   0   'False
      Width           =   525
   End
   Begin VB.TextBox txt 
      Alignment       =   1  'Right Justify
      Height          =   315
      Index           =   23
      Left            =   11040
      TabIndex        =   57
      Text            =   "23"
      Top             =   5400
      Visible         =   0   'False
      Width           =   525
   End
   Begin MSMask.MaskEdBox mskDt 
      Height          =   375
      Index           =   2
      Left            =   2955
      TabIndex        =   3
      Top             =   1065
      Width           =   1245
      _ExtentX        =   2196
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
   Begin MSMask.MaskEdBox mskDt 
      Height          =   375
      Index           =   9
      Left            =   7875
      TabIndex        =   13
      Top             =   195
      Visible         =   0   'False
      Width           =   525
      _ExtentX        =   926
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
      Bindings        =   "frmPurchaseInward.frx":058A
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
   Begin MSMask.MaskEdBox mskDt 
      Height          =   375
      Index           =   7
      Left            =   5880
      TabIndex        =   7
      Top             =   1065
      Width           =   1305
      _ExtentX        =   2302
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
   Begin MSMask.MaskEdBox mskDt 
      Height          =   375
      Index           =   33
      Left            =   4140
      TabIndex        =   28
      Top             =   5280
      Width           =   1305
      _ExtentX        =   2302
      _ExtentY        =   661
      _Version        =   393216
      MaxLength       =   10
      Format          =   "dd/mm/yyyy"
      Mask            =   "##/##/####"
      PromptChar      =   "_"
   End
   Begin MSFlexGridLib.MSFlexGrid msGrid 
      Height          =   2010
      Left            =   120
      TabIndex        =   16
      Top             =   2655
      Width           =   9330
      _ExtentX        =   16457
      _ExtentY        =   3545
      _Version        =   393216
      FixedCols       =   0
      RowHeightMin    =   275
      BackColorFixed  =   -2147483647
      ForeColorFixed  =   -2147483634
      BackColorBkg    =   12632256
      Appearance      =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Shape Shape5 
      BorderColor     =   &H00C00000&
      BorderWidth     =   5
      FillColor       =   &H00C00000&
      Height          =   7215
      Left            =   0
      Top             =   -30
      Width           =   10845
   End
   Begin VB.Image Image2 
      Height          =   1650
      Left            =   7680
      Picture         =   "frmPurchaseInward.frx":059B
      Stretch         =   -1  'True
      Top             =   690
      Width           =   1515
   End
   Begin VB.Label Label17 
      BackStyle       =   0  'Transparent
      Caption         =   "Total Bag               :"
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
      Left            =   5820
      TabIndex        =   93
      Top             =   6210
      Width           =   2055
   End
   Begin VB.Label Label31 
      BackStyle       =   0  'Transparent
      Caption         =   "Total Weight           :"
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
      Left            =   5820
      TabIndex        =   92
      Top             =   6675
      Width           =   2055
   End
   Begin VB.Label Label16 
      BackStyle       =   0  'Transparent
      Caption         =   "Booking No. :"
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
      Left            =   7710
      TabIndex        =   10
      Top             =   255
      Visible         =   0   'False
      Width           =   1665
   End
   Begin VB.Label Label23 
      BackStyle       =   0  'Transparent
      Caption         =   "Booking Dt.  :"
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
      Left            =   7140
      TabIndex        =   12
      Top             =   285
      Visible         =   0   'False
      Width           =   1545
   End
   Begin VB.Label Label15 
      BackStyle       =   0  'Transparent
      Caption         =   "&Unloading  :"
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
      Left            =   390
      TabIndex        =   33
      Top             =   6585
      Width           =   1635
   End
   Begin VB.Label Label14 
      BackStyle       =   0  'Transparent
      Caption         =   "Octroi       :"
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
      Left            =   390
      TabIndex        =   31
      Top             =   6135
      Width           =   1275
   End
   Begin VB.Label Label12 
      BackStyle       =   0  'Transparent
      Caption         =   "&Freight      :"
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
      Left            =   390
      TabIndex        =   29
      Top             =   5715
      Width           =   1635
   End
   Begin VB.Label Label11 
      BackStyle       =   0  'Transparent
      Caption         =   "&Transport   :"
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
      Left            =   390
      TabIndex        =   23
      Top             =   4860
      Width           =   1275
   End
   Begin VB.Label Label9 
      BackStyle       =   0  'Transparent
      Caption         =   "L.R. Date  :"
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
      Height          =   300
      Left            =   3150
      TabIndex        =   27
      Top             =   5280
      Width           =   1755
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "L.R. No.    :"
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
      Left            =   390
      TabIndex        =   25
      Top             =   5280
      Width           =   1635
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Goods Inward &No        :"
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
      Left            =   540
      TabIndex        =   0
      Top             =   645
      Width           =   2175
   End
   Begin VB.Label Label10 
      BackStyle       =   0  'Transparent
      Caption         =   "&Ex-Mill Amount    :"
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
      Left            =   5820
      TabIndex        =   18
      Top             =   5070
      Width           =   2055
   End
   Begin VB.Label Label30 
      BackStyle       =   0  'Transparent
      Caption         =   "Bill Am&ount          :"
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
      Left            =   5820
      TabIndex        =   20
      Top             =   5610
      Width           =   2055
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
      ForeColor       =   &H0000FFFF&
      Height          =   495
      Left            =   8325
      TabIndex        =   79
      Top             =   -15
      Width           =   2475
   End
   Begin VB.Line Line3 
      BorderColor     =   &H000000C0&
      X1              =   270
      X2              =   7440
      Y1              =   1575
      Y2              =   1575
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
      TabIndex        =   77
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
      TabIndex        =   76
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
      TabIndex        =   75
      Top             =   8280
      Width           =   1215
   End
   Begin VB.Label Label24 
      BackStyle       =   0  'Transparent
      Caption         =   "Other Add"
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
      Left            =   11295
      TabIndex        =   74
      Top             =   7920
      Width           =   1095
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
      TabIndex        =   73
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
      TabIndex        =   72
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
      TabIndex        =   71
      Top             =   7860
      Width           =   1455
   End
   Begin VB.Label Label19 
      BackStyle       =   0  'Transparent
      Caption         =   "Landing Cost Details :-"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   375
      Left            =   11175
      TabIndex        =   70
      Top             =   7590
      Width           =   3135
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
      TabIndex        =   69
      Top             =   8160
      Width           =   1095
   End
   Begin VB.Label Label8 
      BackStyle       =   0  'Transparent
      Caption         =   "Mill Bill Date  :"
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
      Left            =   4395
      TabIndex        =   6
      Top             =   1065
      Width           =   1695
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "Mill Bill  No.   :   "
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
      Left            =   4395
      TabIndex        =   4
      Top             =   645
      Width           =   1935
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "B&roker  :"
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
      Left            =   405
      TabIndex        =   14
      Top             =   2145
      Width           =   1470
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "M&ill       :  "
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
      Left            =   405
      TabIndex        =   8
      Top             =   1695
      Width           =   1635
   End
   Begin VB.Label Label6 
      BackColor       =   &H00C00000&
      Caption         =   "  Goods Inward"
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
      TabIndex        =   68
      Top             =   0
      Width           =   10815
   End
   Begin VB.Shape Shape1 
      FillColor       =   &H00BDDFD1&
      FillStyle       =   0  'Solid
      Height          =   6225
      Left            =   9570
      Shape           =   4  'Rounded Rectangle
      Top             =   555
      Width           =   1125
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Goods Received Da&te :"
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
      Left            =   540
      TabIndex        =   2
      Top             =   1065
      Width           =   2160
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H00000080&
      FillColor       =   &H00EDFDFE&
      FillStyle       =   0  'Solid
      Height          =   2175
      Left            =   270
      Top             =   420
      Width           =   9165
   End
   Begin VB.Shape Shape4 
      BorderColor     =   &H00000080&
      FillColor       =   &H00EDFDFE&
      FillStyle       =   0  'Solid
      Height          =   2355
      Left            =   5685
      Top             =   4740
      Width           =   3765
   End
   Begin VB.Shape Shape6 
      BorderColor     =   &H00000080&
      FillColor       =   &H00EDFDFE&
      FillStyle       =   0  'Solid
      Height          =   2355
      Left            =   270
      Top             =   4740
      Width           =   5325
   End
End
Attribute VB_Name = "frmPurchaseInward"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Const tblName As String = "tblAddLess"
Const IdField As String = "Vno"
Const OrderField As String = "Vdt,Vno"
Const MaxNo As Long = 40
Dim Rstbl As Recordset '-- TblAddLess
Dim rstblSub As Recordset '-- TblVouSub
Dim varBookMark As Variant
Dim FirstTimeFlag As Boolean
'Dim rsVou As Recordset '--TblVoucher
'Dim rsOS As Recordset '--TblOutStanding
Dim OldgFrmTypeStr As String
Dim FormAction As Integer
Dim frmTypeStr As String
Dim i As Long
'-- Cartoon Bag Detail Entry
Private Sub CmdBagDetails_Click()
gfrmBagInwVNo = txt(0)
gfrmBagInwVType = frmTypeStr
frmBagEntry.Show
End Sub
'Dim GridFlag As Boolean
Private Sub cmdBtn_Click(Index As Integer)
On Error GoTo ErrorRoutine
Dim Hank As String
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
    gClsSearch.SearchMultiField "tblAddLess,TblMastAccount", "Vno,AcName,Vdt", Array("Vno", "Mill", "Vdt"), Array(txt(0).Width, 4000, 1080), " tblAddLess.AdMillCode=TblMastAccount.AcCode and vtype='" & frmTypeStr & "' and VYear=" & gCYear, "", "Vno,Vdt", 0, 0
    If gClsSearch.SearchMultiRetCol(0) <> "" Then
        Rstbl.MoveFirst
        Rstbl.Find "Vno = " & gClsSearch.SearchMultiRetCol(0)
        gClsSearch.SearchMultiRetCol(0) = ""
        ReadFields
    End If
Case 6 'Add
    FormAction = vbDataActionAddNew
    Call GProcActivateControls(Me, True)
    SetControlEd
    Call GProcSetButtons(Me, vbDataActionAddNew)
    Call GProcClearForm(Me, Rstbl, MaxNo, True)
    If FirstTimeFlag = True Then
        FirstTimeFlag = False
    Else
        txt(0).SetFocus
    End If
    txt(0) = GProcGenerateId("TblAddLess", "Vno", "Vtype='" & frmTypeStr & "' and VYear=" & gCYear)
    mskDt(2) = GProcGenerateVDt("tblAddLess", "Vdt", "Vtype='" & frmTypeStr & "' and VYear=" & gCYear)
    FillCombo
    SetGrid
Case 7 'Modify
    FormAction = vbDataActionUpdate
    Hank = ""
    Call GProcChkPass(frmTypeStr, FormAction, txt(0), txt(2), txtTemp(5), gUserName, Hank)
    If gChkPassTrueFalse = True Then
        'If mURecModi = True Then
        '    SearchHelp
        '    DoEvents
            If IsRelatedRecord = False Then
                FormAction = vbDataActionUpdate
                Call GProcActivateControls(Me, True)
                SetControlEd
                '--Update Stock
                GProcUpdateVSubStock txt(0), gCYear, frmTypeStr, True
                Call GProcSetButtons(Me, vbDataActionUpdate)
                txt(0).Enabled = False
                mskDt(2).SetFocus
                 '--- If Gate Pass Set Ctrl Enable
                If IsRelatedGatePass = True Then
                     Call GProcActivateControls(Me, False)
                     txt(6).Enabled = True '--Bill No
                     mskDt(7).Enabled = True '--Bill Dt
                     txt(35).Enabled = True '-Transport
                     txt(25).Enabled = True '-Freight
                     txt(26).Enabled = True '-Octroi
                     txt(27).Enabled = True '-Unloading
                     txt(32).Enabled = True '-L R No
                     mskDt(33).Enabled = True '--L R Dt
                     Chk(40).Enabled = True '- Is Taxable IN
                     txt(11).Enabled = True '-Bill amt
                     txt(10).Enabled = True '-Ex Mill amt
                     SetControlEd
                End If
            End If
        'End If
    Else
       FormAction = vbDataActionClose
    End If
Case 8 'Delete
    FormAction = vbDataActionDelete
    Hank = ""
    Call GProcChkPass(frmTypeStr, FormAction, txt(0), txt(2), txtTemp(5), gUserName, Hank)
    If gChkPassTrueFalse = True Then
        'If mURecDel = True Then
           If IsRelatedGatePass = True Then
                MsgBox "You can not Modify/Delete this record.Related Gate Pass is exists.", vbCritical + vbOKOnly, App.Title
                Exit Sub
          End If
          If IsRelatedRecord = False Then
            If MsgBox("Do You Really Delete This Record?", vbQuestion + vbYesNo + vbDefaultButton1, "Confirm Delete") = vbYes Then
                If Rstbl.RecordCount > 1 Then
                    '--Update Stock
                    GProcUpdateVSubStock txt(0), gCYear, frmTypeStr, False, True
                     '-- Delete Related Record
                    DeleteRelatedRecord
                    Rstbl.Delete
                    '--- Bag Details
                    gCn.Execute "delete From tblBags where InwNo=" & txt(0) & " and InwType='" & frmTypeStr & "'"
                    cmdBtn_Click (1)
                ElseIf Rstbl.RecordCount = 1 Then
                    '--Update Stock
                    GProcUpdateVSubStock txt(0), gCYear, frmTypeStr, False, True
                    '-- Delete Related Record
                    DeleteRelatedRecord
                    Rstbl.Delete
                    '--- Bag Details
                    gCn.Execute "delete From tblBags where InwNo=" & txt(0) & " and InwType='" & frmTypeStr & "'"
                    Call GProcActivateControls(Me, False)
                    Call GProcSetButtons(Me, vbDataActionCancel)
                    Form_Load
                End If
            End If
        End If
     End If
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
        If Trim(txtGrid) <> "" Then
            msGrid.Text = txtGrid
            DoEvents
            txtGrid = ""
        End If
        txtGrid.Visible = False
        'If FormAction = vbDataActionAddNew Then txt(0) = GProcGenerateId(tblName, IdField, "vtype='" & FrmTypeStr & "' and VYear=" & gCYear)
        CalAmount
        FillTxtFromTemp
        '-- Save Record
        GProcSaveRecord Me, Rstbl, FormAction, MaxNo
        SaveRelatedRecord
        '--Update Stock
        GProcUpdateVSubStock txt(0), gCYear, frmTypeStr
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
    txtGrid.Visible = False
    If Rstbl.RecordCount <= 0 Then
        Call GProcClearForm(Me, Rstbl, MaxNo, True)
        Call GProcActivateControls(Me, False)
        Call GProcSetButtons(Me, vbDataActionCancel)
    Else
        If FormAction = vbDataActionAddNew Then
            Rstbl.MoveLast
        Else
            '--Update Stock
            GProcUpdateVSubStock txt(0), gCYear, frmTypeStr
            Rstbl.CancelUpdate
        End If
        Call GProcClearForm(Me, Rstbl, MaxNo, True)
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
    msGrid.Col = 0
    msGrid.Row = 1
    txtGrid.Move msGrid.CellLeft + msGrid.Left, msGrid.CellTop + msGrid.Top, msGrid.CellWidth, msGrid.CellHeight
End Select
Exit Sub
ErrorRoutine:
    Rstbl.CancelUpdate
    GProcErrorHandler
End Sub
Private Sub SetControlEd()
Dim ExmptP As Double
'txt(10).Enabled = False '--Sub Amt
'txt(11).Enabled = False '-- Amt
txtCrBalance.Enabled = False
txttmpTotBag.Enabled = False
txttmpTotWt.Enabled = False
ExmptP = GProcGetColumnValue("tblMastAccount", "AcName", txtTemp(5), "S", "ACEXEMPTRATE", "N")
'---- Is Taxable Inward
Chk(40).Enabled = False
mskDt(9).Enabled = False '--Book vDt
If ExmptP > 0 Then
    If FormAction = vbDataActionAddNew Or FormAction = vbDataActionUpdate Then
        Chk(40).Enabled = True
    End If
End If
End Sub
Private Sub DeleteRelatedRecord()
'--- Voucher Sub
gCn.Execute "delete from tblVouSub where vtype='" & frmTypeStr & "' and vno=" & txt(0) & " and " _
& " VYear=" & gCYear
End Sub
Private Sub SaveRelatedRecord()
Dim i As Long
Dim MillCode As Long
Dim rsVouSub As Recordset
Dim x As String
'--- Account Code
MillCode = GProcGetColumnValue("tblMastAccount", "AcName", txtTemp(5), "S", "AcCode", "N")
'-- Delete Records
DeleteRelatedRecord
'--- Save Grid Records
'--- TblVouSub
x = "select * from tblVouSub where vtype='" & frmTypeStr & "' and vno=" & txt(0) & " and " _
& " VYear=" & gCYear
GProcRstOpen rsVouSub, x, "O"
With msGrid
    i = 1
    Do While i < .Rows
        rsVouSub.AddNew
        rsVouSub!Vno = CLng(txt(0)) '--Vno(0)
        rsVouSub!Vtype = txt(1) & vbNullString '--Vtype(1)
        rsVouSub!Vdt = CDate(txt(2)) '--Vdt(2)
        rsVouSub!VSubCtrNo = i '--Sr No(3)
        rsVouSub!vsubBookNo = CLng(.TextMatrix(i, 1)) '--Book No
        '--Book Dt
        If CLng(.TextMatrix(i, 1)) = 0 Then  'Book No =0
            .TextMatrix(i, 2) = "" 'Book Dt =Null
            .TextMatrix(i, 16) = 0 ' Booking It ctrl No =0
        End If
        If IsDate(.TextMatrix(i, 2)) Then
            rsVouSub!vsubBookDt = .TextMatrix(i, 2)
        Else
            rsVouSub!vsubBookDt = Null
        End If
        '--Itcode(4)
        rsVouSub!vsubITCODE = GProcGetColumnValue("tblMastItem", "ItName", .TextMatrix(i, 3), "S", "ItCode", "N", " ItMillCode=" & MillCode) '--Item Code
        rsVouSub!vsubBAG = CLng(.TextMatrix(i, 6)) '--Bag(5)
        rsVouSub!vsubwt = CDbl(.TextMatrix(i, 7)) '-- Wt(6)
        rsVouSub!vsubRt = CDbl(.TextMatrix(i, 8)) '--Rate(7)
        rsVouSub!vsubRTPER = CDbl(.TextMatrix(i, 9)) '--Rate Per(8)
        rsVouSub!vsubAMT = CDbl(.TextMatrix(i, 10)) '--Amt(9)
        rsVouSub!vsubLOTNO = .TextMatrix(i, 5) '--Lot No(10)
        '--Godown Code(17)
        rsVouSub!VSubGodownCode = GProcGetColumnValue("tblMastNarration", "Narration", .TextMatrix(i, 4), "S", "NarrCode", "N", "")
        '--- It Ctrl No(11)
        If CLng(.TextMatrix(i, 11)) = 0 Then 'FormAction = vbDataActionAddNew Then
            rsVouSub!VSubItCtrlNo = GProcGenerateId("tblVouSub", "VSubItCtrlNo", "VYear=" & gCYear)   '--It Ctrl No
        Else
            rsVouSub!VSubItCtrlNo = CLng(.TextMatrix(i, 11))
        End If
        rsVouSub!VSubIssBag = CLng(.TextMatrix(i, 12)) '--Issue Bag(13)
        rsVouSub!VSubIsswt = CDbl(.TextMatrix(i, 13)) '-- Issue Wt(14)
        rsVouSub!VSubTmpBag = CLng(.TextMatrix(i, 14)) '--Temp Bag(15)
        rsVouSub!VSubTmpWt = CDbl(.TextMatrix(i, 15)) '-- Temp Wt(16)
        '-- P It Ctrl No =Booking It ctrl No
        rsVouSub!VSubPItCtrlNo = CLng(.TextMatrix(i, 16))
        rsVouSub!VYear = txt(23) '--VYear(18)
        rsVouSub.Update
        i = i + 1
     Loop
 End With
 rsVouSub.Close
End Sub
Private Sub cmdLedger_Click()
On Error GoTo ErrorHandler
frmRptAccount.txt(1) = txtTemp(5)
frmMain.mnuRptRSubAccLeg_Click (0)
Exit Sub
ErrorHandler:
    GProcErrorHandler
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case 13:
            If Me.ActiveControl.Name <> "txtGrid" Then SendKeys "{TAB}"
        Case 27:
            If (cmdBtn(13).Enabled = False) Then
                Unload Me
            End If
    End Select
End Sub
Private Sub Form_Load()
    Dim x As String
    Me.Top = 300
    Me.Left = Screen.Width / 2 - Me.Width / 2
    '--- TblAddLess
    frmTypeStr = gfrmTypeStr
    x = "Select * from " & tblName & " where Vtype='" & frmTypeStr & "' and VYear=" & gCYear
    If Trim(OrderField) <> "" Then x = x & " order by " & OrderField
    GProcRstOpen Rstbl, x, "O"
    FirstTimeFlag = True
    '---Grid
    SetGrid
    '---- Button Visiblity Set As Per User Previlage
    GProcSetButtonVisiblity Me
    cmdBtn_Click (6)
End Sub
Private Sub SetGrid()
Dim x As String
Dim i As Long
Set rstblSub = New Recordset
x = "select VSubCtrNo,VSubBookNo,VsubBookDt,ItName,Narration,vSubLotNo,VSubBag,VSubWt" _
& " ,VSubRt,VSubRtPer,VSubAmt,VSubItCtrlNo,VSubIssBag,VSubIssWt,VSubTmpBag,VSubTmpWt,VSubPItCtrlNo from tblvouSub,TblMastNarration,TblMastItem where " _
& " tblvouSub.VsubItCode=TblMastItem.ItCode and tblvouSub.VSubGodownCode=TblMastNarration.NarrCode " _
& " and tblvouSub.vtype='" & frmTypeStr & "' and tblvouSub.vno=" & txt(0) & " and tblvouSub.VYear=" & gCYear & " order by VSubCtrNo"
GProcRstOpen rstblSub, x, "R"
With msGrid
    .FormatString = ">Sr|<Book No|<Book Date   |<Count                           |<Godown           |<Lot No  |>Bag  |>Weight      |>Rate          |>Rt Per|>Amount        ||||||"
    If FormAction = vbDataActionAddNew Then
        .Rows = 2
        GridAddNew
    Else
        .Rows = 1
    End If
    For i = 11 To 16
        .ColWidth(i) = 0
    Next i
End With
End Sub
Private Sub GridAddNew()
Dim i As Long
With msGrid
    For i = 0 To rstblSub.Fields.Count - 1
        Select Case rstblSub.Fields(i).Type
            Case adInteger, 2, 3 '-- Integer
                .TextMatrix(.Row, i) = 0
            Case adDecimal, 131, adNumeric, adDouble '-- Double
                .TextMatrix(.Row, i) = Format(0, GProcNumberFormat(rstblSub.Fields(i).Precision, rstblSub.Fields(i).NumericScale))
            Case adDate '-- Date
                 .TextMatrix(.Row, i) = ""
            Case Else
                 .TextMatrix(.Row, i) = ""
        End Select
    Next i
ReNumberGridSrno
End With
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

Private Sub msGrid_GotFocus()
    cmdBtn(10).Cancel = False
End Sub
Private Sub MSGrid_LostFocus()
CalAmount
End Sub
Private Sub mskDt_GotFocus(Index As Integer)
GProcSelectBox Me.ActiveControl
End Sub
Private Sub ShowAccountBalance()
'Dim DrAcBal  As Double
Dim CrAcBal As Double
If IsDate(mskDt(2)) Then
    'DrAcBal = GProcGetAccountBalance(txttemp(4), mskDt(2))
    CrAcBal = GProcGetAccountBalance(txtTemp(5), mskDt(2))
    'txtDrBalance = IIf(DrAcBal > 0, DrAcBal & " Cr", Abs(DrAcBal) & " Dr")
    txtCrBalance = IIf(CrAcBal > 0, CrAcBal & " Cr", Abs(CrAcBal) & " Dr")
End If
End Sub
Private Sub mskDt_LostFocus(Index As Integer)
If Index = 7 Then
    If GProcIsDateValid(mskDt(Index)) = False Then
        mskDt(Index).SetFocus
        Exit Sub
    End If
End If
If Index = 2 Then ShowAccountBalance
End Sub
Private Sub txt_GotFocus(Index As Integer)
If Index = 10 Then
    cmdBtn(10).Cancel = True
End If
GProcSelectBox Me.ActiveControl
End Sub
Private Sub txt_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyF3 Then
    Select Case Index
        Case 35 '-- Transport
            OldgFrmTypeStr = gfrmTypeStr
            gfrmTypeStr = "R"
            GProcShowForm frmMastNarration, frmMain.mnuMstfrm(8), "R"
            gfrmTypeStr = OldgFrmTypeStr
    End Select
End If
End Sub
Private Sub txt_LostFocus(Index As Integer)
GProcCheckForNumber Me, Rstbl, MaxNo
SaveBtnEd
CalAmount
End Sub
Private Sub txt_KeyPress(Index As Integer, KeyAscii As Integer)
If KeyAscii = 13 Then Exit Sub
KeyAscii = GProcValidateKey(Rstbl, Index, KeyAscii, txt(Index).Text)
If Index = 35 Then '-- Transport
    gClsSearch.SearchMultiField "tblMastNarration", "Narration", Array("Transport"), Array(txt(Index).Width), " NarrType='R'", Chr(KeyAscii), "Narration", 0, 0, True
    KeyAscii = 0
    If gClsSearch.SearchMultiRetCol(0) <> "" Then
        txt(Index) = gClsSearch.SearchMultiRetCol(0) '--Transport
    End If
End If
'If Index = 8 Then '--Booking Consign/Depot
'    PrepareBooking
'    KeyAscii = 0
'End If
End Sub

Private Sub txttemp_GotFocus(Index As Integer)
GProcSelectBox Me.ActiveControl
End Sub
Private Sub txttemp_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
Dim GName As String
If KeyCode = vbKeyF3 Then
    Select Case Index
        Case 4 '-- Broker
            GName = GProcGetColumnValue("TblMastGroup", "AgCode", 90016, "N", "AgName", "S")
            GProcShowForm frmMastAccount, frmMain.mnuMstfrm(0), Array(frmMastAccount.txtTemp(2)), Array(GName)
        Case 5 '-- millcode
            GName = GProcGetColumnValue("TblMastGroup", "AgCode", 90029, "N", "AgName", "S")
            GProcShowForm frmMastAccount, frmMain.mnuMstfrm(0), Array(frmMastAccount.txtTemp(2)), Array(GName)
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
    Case 4 '-- Broker A/c  Alias Account
        gClsSearch.SearchMultiField "QryMastaccount", "Account,AcName,AcCity,AgName", Array("Broker", "AcName", "City", "Group"), Array(3000, 0, 2000, 3000), "(AgCode=90016 or GpCode=90016)", Chr(KeyAscii), "AcName", LeftPos, TopPos, True
        KeyAscii = 0
        If gClsSearch.SearchMultiRetCol(0) <> "" And gClsSearch.SearchMultiRetCol(1) = "" Then
            txtTemp(Index) = gClsSearch.SearchMultiRetCol(0)
        Else
           If gClsSearch.SearchMultiRetCol(1) <> "" Then txtTemp(Index) = gClsSearch.SearchMultiRetCol(1)
        End If
        'gClsSearch.SearchMultiField "tblMastaccount,tblMastGroup", "AcName", Array("Purchase Account"), Array(txtTemp(Index).Width), " tblMastaccount.AgCode=tblMastGroup.AgCode and (tblMastGroup.AgCode=25 or GpCode=25)", Chr(KeyAscii), "AcName", LeftPos, TopPos, True
    Case 5  '-- Mill  Alias Account
        gClsSearch.SearchMultiField "QryMastaccount", "Account,AcName,AcCity,AgName", Array("Mill", "AcName", "City", "Group"), Array(3000, 0, 2000, 3000), " ((AgCode=90029 or GpCode=90029) and (AcMillType=1 or AcMillType=2) )", Chr(KeyAscii), "AcName", LeftPos, TopPos, True
        KeyAscii = 0
        If gClsSearch.SearchMultiRetCol(0) <> "" And gClsSearch.SearchMultiRetCol(1) = "" Then
            txtTemp(Index) = gClsSearch.SearchMultiRetCol(0)
        Else
            If gClsSearch.SearchMultiRetCol(1) <> "" Then txtTemp(Index) = gClsSearch.SearchMultiRetCol(1)
        End If
        'gClsSearch.SearchMultiField "tblMastaccount,tblMastGroup", "AcName", Array("Mill"), Array(txtTemp(Index).Width), " tblMastaccount.AgCode=tblMastGroup.AgCode and (tblMastGroup.AgCode=29 or GpCode=29)", Chr(KeyAscii), "AcName", LeftPos, TopPos, True
End Select
End Sub
Private Sub txtTemp_LostFocus(Index As Integer)
SaveBtnEd
ShowAccountBalance
SetControlEd
End Sub
'---------------- Grid
Private Sub msGrid_Click()
On Error Resume Next
    msGrid_EnterCell
End Sub
Private Sub msGrid_EnterCell()
If FormAction <> vbDataActionClose Then
    CheckForNumberTxtGrid
    DoEvents
    txtGrid = ""
    DoEvents
    If msGrid.Text <> "" Then txtGrid.Text = msGrid.Text
    '--- Lock TextBox
    With msGrid
        Select Case .Col
            Case 0: '---- Sr No
                txtGrid.Locked = True
                ReNumberGridSrno
                txtGrid = ""
                .Col = .Col + 1
            Case 2 '-- Booking Date
                txtGrid.Locked = True
            Case 3 '-- Count
                If .TextMatrix(.Row, 1) > 0 Then   '-- Book No
                    txtGrid.Locked = True
                Else
                    txtGrid.Locked = False
                End If
            Case Else
                txtGrid.Locked = False
        End Select
    End With
    msGrid_KeyPress (0)
End If
End Sub
Private Sub msGrid_KeyPress(KeyAscii As Integer)
With msGrid
    If KeyAscii = 27 Then '----Esc Key
        KeyAscii = 0
        If (Val(.TextMatrix(.Row, 8)) = 0) And .Rows > 2 Then
            .RemoveItem (.Row)
            txtGrid.Text = ""
        End If
        txtGrid.Visible = False
        txt(18).SetFocus
        Exit Sub
    End If
End With

'If FormAction <> vbDataActionAddNew And FormAction <> vbDataActionUpdate Then Exit Sub
'---- Navigate Row ,Col
If KeyAscii = 13 Then
    KeyAscii = 0
    With msGrid
        If .Col = 10 Then 'rstblSub.Fields.Count - 2 Then
            If .Row = (.Rows - 1) Then '-- Add New Row
                .Rows = .Rows + 1
                .Row = .Row + 1
                GridAddNew
            Else '--Next Row
                .Row = .Row + 1
            End If
            .Col = 1
        Else    '-- Next col
            .Col = .Col + 1
        End If
    End With
End If
'---- Set TxtGrid
If msGrid.Text <> "" Then txtGrid.Text = msGrid.Text
'If MSGrid.Col <> rstblSub.Fields.Count - 1 Then txtGrid.Move MSGrid.CellLeft + MSGrid.Left, MSGrid.CellTop + MSGrid.Top, MSGrid.CellWidth, MSGrid.CellHeight
If msGrid.Col < rstblSub.Fields.Count - 1 Then
    txtGrid.Move msGrid.CellLeft + msGrid.Left, msGrid.CellTop + msGrid.Top, msGrid.CellWidth, msGrid.CellHeight
End If
txtGrid.Visible = True
txtGrid.Enabled = True
DoEvents
If KeyAscii <> 0 Then
    Select Case rstblSub.Fields(msGrid.Col).Type
        Case adChar, adVarWChar
            If msGrid.Col = 1 Or msGrid.Col = 3 Or msGrid.Col = 4 Then '-- Count ,Godown,Book No
                txtGrid_KeyPress (KeyAscii)
            Else
                txtGrid.Text = txtGrid.Text & Chr(KeyAscii)
            End If
        Case Else
            txtGrid.Text = Chr(GProcValidateKey(rstblSub, msGrid.Col, KeyAscii, txtGrid.Text)) 'Chr(GProcNumberWithDecimal(KeyAscii, txtGrid))
    End Select
End If
DoEvents
If msGrid.Text <> "" Then
    txtGrid.SelStart = Len(txtGrid.Text) + 1
Else
    txtGrid.SelStart = 1
End If
DoEvents
If txtGrid.Visible = True Then txtGrid.SetFocus
End Sub
Private Sub msGrid_LeaveCell()
If txtGrid.Visible Then
    msGrid.Text = txtGrid.Text
    DoEvents
    txtGrid.Text = ""
    txtGrid.Visible = False
End If
DoEvents
CalGridAmt
CalAmount
End Sub
Private Sub CalGridAmt()
Dim MillCode As Long
Dim StdPack As Double
CheckForNumberGrid
If msGrid.TextMatrix(msGrid.Row, 1) = 0 Then msGrid.TextMatrix(msGrid.Row, 2) = "" '-- Book date
MillCode = GProcGetColumnValue("TblMastAccount", "AcName", txtTemp(5), "S", "Accode", "N")
StdPack = GProcGetColumnValue("TblMastItem", "ITName", msGrid.TextMatrix(msGrid.Row, 3), "S", "ItStdpack", "N", " ItMillCode=" & MillCode)
'-- Wt=Bag * Std Pack
If msGrid.TextMatrix(msGrid.Row, 7) = 0 Then
    msGrid.TextMatrix(msGrid.Row, 7) = msGrid.TextMatrix(msGrid.Row, 6) * StdPack
End If
'-- Amt
'If msGrid.TextMatrix(msGrid.Row, 8) = 0 Then
'    '--Wt*Rate/Rate Per
'    If CDbl(msGrid.TextMatrix(msGrid.Row, 7)) > 0 Then
'        msGrid.TextMatrix(msGrid.Row, 8) = CDbl(msGrid.TextMatrix(msGrid.Row, 5)) * CDbl(msGrid.TextMatrix(msGrid.Row, 6)) / CDbl(msGrid.TextMatrix(msGrid.Row, 7))
'    Else
'        msGrid.TextMatrix(msGrid.Row, 8) = CDbl(msGrid.TextMatrix(msGrid.Row, 5)) * CDbl(msGrid.TextMatrix(msGrid.Row, 6))
'    End If
'    msGrid.TextMatrix(msGrid.Row, 8) = Format(msGrid.TextMatrix(msGrid.Row, 8), FStr)
'Else '--Rate =Amt*Rate Per/Wt
    If CDbl(msGrid.TextMatrix(msGrid.Row, 7)) > 0 Then
        msGrid.TextMatrix(msGrid.Row, 8) = CDbl(msGrid.TextMatrix(msGrid.Row, 10)) * CDbl(msGrid.TextMatrix(msGrid.Row, 9)) / CDbl(msGrid.TextMatrix(msGrid.Row, 7))
    Else
        msGrid.TextMatrix(msGrid.Row, 8) = 0
    End If
    msGrid.TextMatrix(msGrid.Row, 8) = Format(msGrid.TextMatrix(msGrid.Row, 8), LStr)
'End If
End Sub
Private Sub CheckForNumberGrid()
Dim RowIndex As Long
With rstblSub
For RowIndex = 1 To msGrid.Rows - 1
    For i = 0 To .Fields.Count - 1
        Select Case .Fields(i).Type
            Case adInteger, 2, 3 '-- Integer
                If IsNumeric(msGrid.TextMatrix(RowIndex, i)) = False Then
                    msGrid.TextMatrix(RowIndex, i) = 0
                Else
                    msGrid.TextMatrix(RowIndex, i) = CLng(msGrid.TextMatrix(RowIndex, i))
                End If
            Case adDecimal, 131, adNumeric, adDouble '-- Double
                If IsNumeric(msGrid.TextMatrix(RowIndex, i)) = False Then
                    msGrid.TextMatrix(RowIndex, i) = Format(0, GProcNumberFormat(.Fields(i).Precision, .Fields(i).NumericScale))
                Else
                    msGrid.TextMatrix(RowIndex, i) = Format(msGrid.TextMatrix(RowIndex, i), GProcNumberFormat(.Fields(i).Precision, .Fields(i).NumericScale))
                End If
            Case adChar, adVarWChar
            Case adDate '-- Date
                If IsDate(msGrid.TextMatrix(RowIndex, i)) = False Then
                    msGrid.TextMatrix(RowIndex, i) = ""
                End If
        End Select
    Next i
  Next RowIndex
End With
End Sub
'---  Check For Number Grid
Private Sub CheckForNumberTxtGrid()
With rstblSub
    Select Case .Fields(msGrid.Col).Type
        Case adInteger, 2, 3 '-- Integer
            If IsNumeric(txtGrid) = False Then
                txtGrid = 0
            Else
                txtGrid = CLng(txtGrid)
            End If
            txtGrid.MaxLength = .Fields(msGrid.Col).Precision
        Case adDecimal, 131, adNumeric, adDouble '-- Double
            If IsNumeric(txtGrid) = False Then
                txtGrid = Format(0, GProcNumberFormat(.Fields(msGrid.Col).Precision, .Fields(msGrid.Col).NumericScale))
            Else
                txtGrid = Format(txtGrid, GProcNumberFormat(.Fields(msGrid.Col).Precision, .Fields(msGrid.Col).NumericScale))
            End If
            txtGrid.MaxLength = .Fields(msGrid.Col).Precision
    '    Case adChar, adVarWChar '--String
   '         txtGrid.MaxLength = .Fields(msGrid.Col).DefinedSize
        Case adDate '-- Date
            If IsDate(txtGrid) = False Then
                txtGrid = ""
            Else
                txtGrid = CDate(txtGrid)
            End If
            txtGrid.MaxLength = 10
         Case Else
            txtGrid.MaxLength = .Fields(msGrid.Col).DefinedSize
    End Select
End With
End Sub
Private Sub txtGrid_GotFocus()
   txtGrid.Alignment = ProcSetAlignment(rstblSub, msGrid.Col)
    DoEvents
    Call GProcSelectBox(txtGrid)
End Sub
Private Function ProcSetAlignment(pRstbl As Recordset, pColIndex As Long)
Select Case pRstbl.Fields(pColIndex).Type
    Case adInteger, adDouble, 2, 3, 131, adDecimal
        ProcSetAlignment = 1
    Case Else
        ProcSetAlignment = 0
End Select
End Function
Private Sub txtGrid_KeyDown(KeyCode As Integer, Shift As Integer)
'--- Delete Row On Ctrl+Y
If KeyCode = vbKeyY And Shift = 2 And msGrid.Rows > 2 Then
    msGrid = 4
    msGrid.RemoveItem (msGrid.Row)
    txtGrid.Visible = False
    txtGrid.Text = ""
    ReNumberGridSrno
    CalAmount
    Exit Sub
End If

'--- Show Form
If KeyCode = vbKeyF3 Then
Select Case msGrid.Col
    Case 3 '-- Item
        GProcShowForm frmMastItem, frmMain.mnuMstfrm(3), Array(frmMastItem.txtTemp(9)), Array(txtTemp(5))
        Exit Sub
    Case 4 '-- Godown
        OldgFrmTypeStr = gfrmTypeStr
        gfrmTypeStr = "G"
        GProcShowForm frmMastNarration, frmMain.mnuMstfrm(4), "G"
        gfrmTypeStr = OldgFrmTypeStr
End Select
End If
End Sub
'-- ReNumber Grid Sr No
Private Sub ReNumberGridSrno()
Dim i As Long
With msGrid
    For i = 1 To .Rows - 1
        .TextMatrix(i, 0) = i
    Next i
End With
End Sub
Private Sub txtGrid_KeyPress(KeyAscii As Integer)
Dim LeftPos As Long
Dim TopPos As Long
Dim MillCode As Long
TopPos = txtGrid.Top + Me.Top + 650
LeftPos = txtGrid.Left + Me.Left
DoEvents
With msGrid
    If KeyAscii <> 13 And KeyAscii <> 27 Then
        '--- Search
        Select Case .Col
            Case 1 '--Booking Consign/Depot
                PrepareBooking Chr(KeyAscii)
                KeyAscii = 0
            Case 3 '-- Count
                If msGrid.TextMatrix(msGrid.Row, 1) > 0 Then
                    KeyAscii = 0
                Else
                    MillCode = GProcGetColumnValue("TblMastAccount", "AcName", txtTemp(5), "S", "Accode", "N")
                    gClsSearch.SearchMultiField "tblMastItem,tblMastaccount", "ItName,AcName,ItStdpack,ItStdrateper", Array("Count Name", "Mill", "Std. Pack", "Rate Per"), Array(txtGrid.Width, 2000, 1000, 1000), "tblMastItem.ITMillCode=tblMastaccount.acCode and ITMillCode=" & MillCode, Chr(KeyAscii), "ItName", LeftPos, TopPos, True
                    txtGrid = ""
                    KeyAscii = 0
                    If gClsSearch.SearchMultiRetCol(0) <> "" Then '-- Item Wt ,Rate From Master
                        txtGrid = gClsSearch.SearchMultiRetCol(0)
                        If IsNumeric(.TextMatrix(.Row, 7)) = False Then .TextMatrix(.Row, 7) = 0
                        If IsNumeric(.TextMatrix(.Row, 9)) = False Then .TextMatrix(.Row, 9) = 0
                        'If .TextMatrix(.Row, 5) = 0 Then .TextMatrix(.Row, 5) = gClsSearch.SearchMultiRetCol(2)
                        If .TextMatrix(.Row, 9) = 0 Then .TextMatrix(.Row, 9) = gClsSearch.SearchMultiRetCol(3)
                        txtGrid_KeyPress 13
                    End If
                End If
            Case 4  '-- Godown
                gClsSearch.SearchMultiField "tblMastNarration", "Narration", Array("Godown"), Array(txtGrid.Width), " NarrType='G' ", Chr(KeyAscii), "Narration", LeftPos, TopPos, True
                txtGrid = ""
                KeyAscii = 0
                If gClsSearch.SearchMultiRetCol(0) <> "" Then
                    txtGrid = gClsSearch.SearchMultiRetCol(0)
                     txtGrid_KeyPress 13
                    'msGrid_LeaveCell
                    'DoEvents
                    'MSGrid.Col = MSGrid.Col + 1
                    'txtGrid_KeyPress 13
                    ' Exit Sub
                End If
            Case Else
                KeyAscii = GProcValidateKey(rstblSub, msGrid.Col, KeyAscii, txtGrid)
        End Select
   End If
    If KeyAscii = 13 Then
        CheckForNumberTxtGrid
    End If
    '--- Navigate Row ,Col
    If KeyAscii = 13 And .Col < 10 Then 'rstblSub.Fields.Count - 2 Then
        .Col = .Col + 1
    ElseIf KeyAscii = 13 And .Col = 10 Then '= rstblSub.Fields.Count - 2 Then
        If .Row = (.Rows - 1) Then
            .Rows = .Rows + 1
            .Row = .Row + 1
            GridAddNew
        Else
            .Row = .Row + 1
        End If
        .Col = 1
    End If
    '-- On ESC
    With msGrid
        If KeyAscii = 27 Then '----Esc Key
            KeyAscii = 0
            If (Val(.TextMatrix(.Row, 8)) = 0) And .Rows > 2 Then
                .RemoveItem (.Row)
                txtGrid.Text = ""
            End If
            txtGrid.Visible = False
            txt(18).SetFocus
            Exit Sub
        End If
    End With
End With
End Sub

'-- Grid Required
Private Function ValidateGrid() As Boolean
Dim i As Long
Dim MillCode As Long
Dim RowIndex As Long
Dim oldFrmTypeStr As String
Dim pItCtrlNo As Long
Dim ItCode As Long
Dim x As String
Dim Rs1 As Recordset
If Trim(txtGrid.Text) <> "" Then
    msGrid.Text = txtGrid.Text
    CalGridAmt
End If
txtGrid.Visible = False
txtGrid.Text = ""
CheckForNumberGrid
With msGrid
Validate:
    For RowIndex = 1 To .Rows - 1
        For i = 0 To rstblSub.Fields.Count - 1
            Select Case i
                Case 3 '-- Count
                    '-Rate=0 ,count,wt=0 then delete row
                    If RowIndex > 1 And Trim(.TextMatrix(RowIndex, i)) = "" And CDbl(.TextMatrix(RowIndex, 7)) = 0 And CDbl(.TextMatrix(RowIndex, 8)) = 0 Then
                        .RemoveItem (RowIndex)
                        GoTo Validate
                    End If
                    If Trim(.TextMatrix(RowIndex, i)) = "" Then
                        MsgBox "Check Count.", vbCritical + vbOKOnly, Me.Caption
                        .Col = i
                        msGrid.SetFocus
                        Exit Function
                    End If
                    MillCode = GProcGetColumnValue("TblMastAccount", "AcName", txtTemp(5), "S", "Accode", "N")
                    If GProcGetColumnValue("TblMastItem", "ItName", .TextMatrix(RowIndex, i), "S", "Itcode", "N", "ITMillCode=" & MillCode) = 0 Then
                        MsgBox "Check Count.Mill and Count are Mismatch.", vbCritical + vbOKOnly, Me.Caption
                        .Col = i
                        msGrid.SetFocus
                        Exit Function
                    End If
                Case 4 '-- Godown
                    If Trim(.TextMatrix(RowIndex, i)) = "" Then
                        MsgBox "Check Godown.", vbCritical + vbOKOnly, Me.Caption
                        .Col = i
                        msGrid.SetFocus
                        Exit Function
                    End If
                    If GProcGetColumnValue("TblMastNarration", "Narration", .TextMatrix(RowIndex, i), "S", "Narrcode", "N") = 0 Then
                        oldFrmTypeStr = gfrmTypeStr
                        gfrmTypeStr = "G"
                        GProcShowForm frmMastNarration, frmMain.mnuMstfrm(4), Array(frmMastNarration.txt(2)), Array(.TextMatrix(RowIndex, i))
                        gfrmTypeStr = oldFrmTypeStr
                        Exit Function
                    End If
                Case 5 '-- Lot No
'                    If Trim(.TextMatrix(RowIndex, i)) = "" Then
'                        MsgBox "Check Lot No.", vbCritical + vbOKOnly, Me.Caption
'                        .Col = i
'                        MSGrid.SetFocus
'                        Exit Function
'                    End If
                Case 6 '-- Bag
                    If CDbl(.TextMatrix(RowIndex, i)) = 0 Then
                        MsgBox "Check Bag.", vbCritical + vbOKOnly, Me.Caption
                        .Col = i
                        msGrid.SetFocus
                        Exit Function
                    End If
'                    '-- If Bag < Issue Bag
'                    If CDbl(.TextMatrix(RowIndex, 4)) < CDbl(.TextMatrix(RowIndex, 10)) Then
'                        MsgBox "Check Bag.Bag Less than issued.", vbCritical + vbOKOnly, Me.Caption
'                        .Col = i
'                        msGrid.SetFocus
'                        Exit Function
'                    End If
                Case 7 '-- Wt
                    If CDbl(.TextMatrix(RowIndex, i)) = 0 Then
                        MsgBox "Check Weight.", vbCritical + vbOKOnly, Me.Caption
                        .Col = i
                        msGrid.SetFocus
                        Exit Function
                    End If
                    '-- If Wt < Issue Wt
'                    If CDbl(.TextMatrix(RowIndex, 5)) < CDbl(.TextMatrix(RowIndex, 11)) Then
'                        MsgBox "Check Wt.Weight Less than issued.", vbCritical + vbOKOnly, Me.Caption
'                        .Col = i
'                        msGrid.SetFocus
'                        Exit Function
'                    End If
                Case 8 '-- Rate
                    If CDbl(.TextMatrix(RowIndex, i)) = 0 Then
                        MsgBox "Check Rate.", vbCritical + vbOKOnly, Me.Caption
                        .Col = i
                        msGrid.SetFocus
                        Exit Function
                    End If
           End Select
        Next i
        '---It Ctrl No (Booking)
        pItCtrlNo = .TextMatrix(RowIndex, 16)
         For i = 1 To .Rows - 1
            '-- Duplicate Booking Count Selection
            If i <> RowIndex And .TextMatrix(i, 16) = pItCtrlNo And pItCtrlNo > 0 Then
                MsgBox "Duplicate Booking Count Selection.", vbCritical + vbOKOnly, Me.Caption
                .Col = 1
                msGrid.SetFocus
                Exit Function
            End If
        Next i
        '-- Invalid Booking
        If CLng(.TextMatrix(RowIndex, 1)) > 0 Then
            MillCode = GProcGetColumnValue("TblMastAccount", "AcName", txtTemp(5), "S", "Accode", "N")
            ItCode = GProcGetColumnValue("tblMastItem", "ItName", .TextMatrix(RowIndex, 3), "S", "ItCode", "N", " ItMillCode=" & MillCode)
            x = "Select * from tblvouSub where Vno=" & CLng(.TextMatrix(RowIndex, 1)) & " and  " _
            & " VSUBITCODE=" & ItCode & " and  vsubitctrlno=" & .TextMatrix(RowIndex, 16) & " and Vtype='OI' and "
            If gBackEndDB = gBackEndAccess Then
                x = x + "  vdt=cdate('" & CDate(.TextMatrix(RowIndex, 2)) & "')"
            Else '-- Oracle
                x = x + "  vdt= to_date('" & CDate(.TextMatrix(RowIndex, 2)) & "','dd/MM/yyyy')"
            End If
            GProcRstOpen Rs1, x, "R"
            If Rs1.EOF And Rs1.BOF Then
                MsgBox "Check Booking No.", vbInformation + vbOKOnly, Me.Caption
                .Col = 1
                msGrid.SetFocus
                Exit Function
            End If
            Rs1.Close
        End If
      Next RowIndex
End With
ValidateGrid = True
ReNumberGridSrno
End Function
Private Sub FillCombo()
'Call GProcFillCombo(cbo(7), "select Distinct AcCity from " & tblName & " where AcCity <> '' order by AcCity")
End Sub
Private Sub SaveBtnEd()
'--- Vno,vdt,Mill,Amount
If CDbl(txt(0)) = 0 Or GProcIsDateValid(mskDt(2)) = False Or txtTemp(5) = "" Or CDbl(txt(11)) = 0 Then
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
Dim GName As String
Dim rstAddCheck As Recordset
Dim rstModiCheck As Recordset
    '--- Validate Grid
    If ValidateGrid = False Then
        'MSGrid.SetFocus
        Exit Function
    End If
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
    '---- Mill
    If txtTemp(5) = "" Then
        MsgBox "Check Mill.", vbInformation + vbOKOnly, Me.Caption
        txtTemp(5).SetFocus
        Exit Function
    End If
        '--Create Mill
        If GProcGetColumnValue("tblMastAccount", "AcName", txtTemp(5), "S", "AcCode", "N") = 0 Then
           GName = GProcGetColumnValue("TblMastGroup", "AgCode", 90029, "N", "AgName", "S")
           GProcShowForm frmMastAccount, frmMain.mnuMstfrm(0), Array(frmMastAccount.txt(1), frmMastAccount.txtTemp(2)), Array(txtTemp(5), GName)
           Exit Function
        End If
    '--- Broker
    If txtTemp(4) = "" Then txtTemp(4) = "-"
'    If txtTemp(4) = "" Then
'        MsgBox "Check Broker.", vbInformation + vbOKOnly, Me.Caption
'        txtTemp(4).SetFocus
'        Exit Function
'    End If
        '--Create Broker
        If GProcGetColumnValue("tblMastAccount", "AcName", txtTemp(4), "S", "AcCode", "N") = 0 Then
           GName = GProcGetColumnValue("TblMastGroup", "AgCode", 90016, "N", "AgName", "S")
           GProcShowForm frmMastAccount, frmMain.mnuMstfrm(0), Array(frmMastAccount.txt(1), frmMastAccount.txtTemp(2)), Array(txtTemp(4), GName)
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
        x = "Select vno From " & tblName & " Where vno=" & CLng(txt(0)) & " and Vtype='" & frmTypeStr & "' and VYear=" & gCYear
        i = GProcRstOpen(rstAddCheck, x, "R")
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
If IsDate(mskDt(2)) Then '--Vdt
    txt(2) = CDate(mskDt(2))
Else
    txt(2) = ""
End If
If IsDate(mskDt(9)) Then '-- Book dt
    txt(9) = CDate(mskDt(9))
Else
    txt(9) = ""
End If
If IsDate(mskDt(7)) Then
    txt(7) = CDate(mskDt(7))
Else
    txt(7) = ""
End If
If IsDate(mskDt(33)) Then
    txt(33) = CDate(mskDt(33))
Else
    txt(33) = ""
End If
'--- Bill No
txt(36) = txt(6)
txt(3) = GProcGetColumnValue("tblMastAccount", "AcName", txtTemp(5), "S", "AcCode", "N")
'-- Broker
txt(4) = GProcGetColumnValue("tblMastAccount", "AcName", txtTemp(4), "S", "AcCode", "N")
'--Mill
txt(5) = GProcGetColumnValue("tblMastAccount", "AcName", txtTemp(5), "S", "AcCode", "N")
'--Tax Code
txt(12) = GProcGetColumnValue("tblMastNarration", "Narration", "-", "S", "NarrCode", "N")
txt(23) = gCYear
'-- Taxble Inward
 txt(40) = Chk(40)
ShowAccountBalance
End Sub
Private Sub FillTempFromTxt()
Dim x As String
If IsDate(txt(2)) Then '--VDt
    mskDt(2) = txt(2)
Else
    mskDt(2) = "__/__/____"
End If
If IsDate(txt(7)) Then
    mskDt(7) = txt(7)
Else
    mskDt(7) = "__/__/____"
End If
If IsDate(txt(9)) Then '--Book Dt
    mskDt(9) = txt(9)
Else
    mskDt(9) = "__/__/____"
End If
If IsDate(txt(33)) Then
    mskDt(33) = txt(33)
Else
    mskDt(33) = "__/__/____"
End If
'-- Broker
txtTemp(4) = GProcGetColumnValue("tblMastAccount", "AcCode", txt(4), "N", "AcName", "S")
'--Mill
txtTemp(5) = GProcGetColumnValue("tblMastAccount", "AcCode", txt(5), "N", "AcName", "S")
'-- Taxble Inward
Chk(40) = txt(40)
'-- Fill Related Record
FillRelatedRecord
ShowAccountBalance
SetControlEd
End Sub
Private Sub FillRelatedRecord()
Dim RowIndex As Long
Dim ColIndex As Long
Dim TotBag As Long
Dim TotWt As Double
'--Fill Grid
SetGrid
With rstblSub
    If Not .EOF Then
        .MoveFirst
        RowIndex = 1
        msGrid.Rows = 2
        Do While Not .EOF
            For ColIndex = 0 To .Fields.Count - 1
                Select Case .Fields(ColIndex).Type
                    Case adInteger, 2, 3 '-- Integer
                        msGrid.TextMatrix(RowIndex, ColIndex) = CLng(.Fields(ColIndex))
                    Case adDecimal, 131, adNumeric, adDouble  '-- Double
                        msGrid.TextMatrix(RowIndex, ColIndex) = Format(.Fields(ColIndex), GProcNumberFormat(.Fields(ColIndex).Precision, .Fields(ColIndex).NumericScale))
                    Case Else
                        msGrid.TextMatrix(RowIndex, ColIndex) = .Fields(ColIndex) & vbNullString
                End Select
                 If ColIndex = 6 Then TotBag = TotBag + CLng(msGrid.TextMatrix(RowIndex, ColIndex)) '-- Tot Bag
                 If ColIndex = 7 Then TotWt = TotWt + CDbl(msGrid.TextMatrix(RowIndex, ColIndex)) '-- Tot Wt
            Next ColIndex
            RowIndex = RowIndex + 1
            .MoveNext
            If Not .EOF Then msGrid.Rows = msGrid.Rows + 1
        Loop
    End If
End With
txttmpTotBag = TotBag '-- Tot Bag
txttmpTotWt = Format(TotWt, WtStr) '-- Tot Wt
End Sub
Private Sub ReadFields()
FormAction = vbDataActionClose
GProcShowRecord Me, Rstbl, MaxNo
FillTempFromTxt
End Sub
Private Sub CalAmount()
Dim SubAmt As Double
Dim RowIndex As Long
Dim TotBag As Long
Dim TotWt As Double
'---- Sub amt
CheckForNumberGrid
RowIndex = 1
With msGrid
    Do While RowIndex < .Rows
        '-- sub amt
        If Val(.TextMatrix(RowIndex, 10)) > 0 Then
            SubAmt = SubAmt + Val(.TextMatrix(RowIndex, 10))
        End If
        '--Tot Bag
        If Val(.TextMatrix(RowIndex, 6)) > 0 Then
            TotBag = TotBag + Val(.TextMatrix(RowIndex, 6))
        End If
        '--Tot Wt
        If Val(.TextMatrix(RowIndex, 7)) > 0 Then
            TotWt = TotWt + Val(.TextMatrix(RowIndex, 7))
        End If
        RowIndex = RowIndex + 1
    Loop
End With
txttmpTotBag = TotBag
txttmpTotWt = Format(TotWt, WtStr)
If txt(10) = 0 Then
    txt(10) = SubAmt '-- Ex Mill Amt
Else
    If FormAction = vbDataActionAddNew And txt(10) < SubAmt Then txt(10) = SubAmt
End If
'If txt(11) = 0 Then txt(11) = SubAmt '-- Bill Amt
GProcCheckForNumber Me, Rstbl, MaxNo
End Sub
'--- Is Related Record
Private Function IsRelatedRecord() As Boolean
Dim RsG As Recordset
Dim x As String
Dim RsR As Recordset
'-- Audited
If GProcISAudited(txt(0), txt(1), txt(23)) = True Then
    MsgBox "Record is Audited.You can not Modify/Delete this Record.", vbCritical + vbOKOnly, App.Title
    IsRelatedRecord = True
    Exit Function
End If
'-- Purchase Return
x = "Select Vtype,Vdt,Vno from tblVouSub where Vtype in ('VY','VI') and VSubPItCtrlNo in" _
& " (select VSubItCtrlNo from tblVouSub where vtype='" & frmTypeStr & "' and vno=" & txt(0) & " and " _
& " VYear=" & gCYear & ")"
GProcRstOpen RsR, x, "R"
If RsR.EOF And RsR.BOF Then
Else
    MsgBox "Related Purchase Return Record " & RsR.Fields("Vno") & " is exists.You can not Modify/Delete this Record.", vbCritical + vbOKOnly, App.Title
    IsRelatedRecord = True
    Exit Function
End If
''-- Is Gate Pass
'X = "Select Vno,Vdt from tblVouSub where Vtype='GP' and " _
'& " VSubPItCtrlNo in (select VSubItCtrlNo from tblVouSub where vtype='" & frmTypeStr & "' and vno=" & txt(0) & " and " _
'& " VYear=" & gCYear & ")"
'GProcRstOpen RsG, X, "R"
'If RsG.EOF Then
'Else
'    IsRelatedRecord = True
'    MsgBox "You can not Modify/Delete this record.Related Gate Pass No. " & RsG.Fields(0) & " is exists.", vbCritical + vbOKOnly, App.Title
'    Exit Function
'End If
'RsG.Close
End Function
'-- Is Related GatePass
Private Function IsRelatedGatePass() As Boolean
Dim RsG As Recordset
Dim x As String
'-- Is Gate Pass
x = "Select Vno,Vdt from tblVouSub where Vtype='GP' and " _
& " VSubPItCtrlNo in (select VSubItCtrlNo from tblVouSub where vtype='" & frmTypeStr & "' and vno=" & txt(0) & " and " _
& " VYear=" & gCYear & ")"
GProcRstOpen RsG, x, "R"
If RsG.EOF Then
Else
    IsRelatedGatePass = True
   ' MsgBox "You can not Modify/Delete this record.Related Gate Pass No. " & RsG.Fields(0) & " is exists.", vbCritical + vbOKOnly, App.Title
    Exit Function
End If
RsG.Close
End Function

'---------- Booking Vs Purchase Inward
Private Sub PrepareBooking(Optional Chr As String)
Dim MillCode As Long
Dim QStr As String
Dim RsBk As Recordset
If IsDate(mskDt(2)) = True Then
    MillCode = GProcGetColumnValue("TblMastAccount", "AcName", txtTemp(5), "S", "Accode", "N")
   '--- Add
    QStr = " select tblVousub.VNo,tblVousub.vdt,AcName,ItName,vsubbag-(vsubissbag),vsubwt-(vsubisswt),vsubbag,vsubwt,vsubitctrlno from " _
    & " tblVousub,tblMastAccount,tblmastItem,tbladdless Where " _
    & " tblVousub.vno=tbladdless.vno and tblVousub.vType=tbladdless.vType and tblVousub.vYear=tbladdless.vYear and " _
    & " Accode=ADMILLCODE and Itcode=VSUBITCODE and ADMILLCODE=" & MillCode & " and " _
    & " ((vsubbag-(vsubissbag))> 0) and tblVousub.vtype ='OI'  and VSUBBOOKISCOMPLETED=0 and "
    If gBackEndDB = gBackEndAccess Then
        QStr = QStr + " tblVousub.vdt<=cdate('" & CDate(mskDt(2)) & "')"
    Else '-- Oracle
        QStr = QStr + " tblVousub.vdt <= to_date('" & CDate(mskDt(2)) & "','dd/MM/yyyy')"
    End If
    With gClsSearch
        .SearchMultiField QStr, "", Array("Booking No.", "Date", "Mill", "Count", "Bal Bag", "Bal Wt.", "Bag", "Wt", "itctrlno"), Array(1000, 1500, 0, 1500, 1000, 1000, 1000, 1000, 0), "", Chr, "tblVousub.Vdt,tblVousub.vno,VSUBCTRNO ", 0, 0, False, , True
        If .SearchMultiRetCol(0) <> "" Then
            QStr = "select tblvousub.*,tblmastItem.ItName from tblvousub,tblMastItem where " _
            & " ItCode=VSUBITCODE and VSUBITCTRLNO= " & CLng(.SearchMultiRetCol(8))
            If GProcRstOpen(RsBk, QStr, "R") > 0 Then
                txtGrid = ""
                txtGrid = CLng(.SearchMultiRetCol(0)) '--Book Vno
                msGrid.TextMatrix(msGrid.Row, 1) = CLng(.SearchMultiRetCol(0)) '--Book Vno
                msGrid.TextMatrix(msGrid.Row, 2) = Format(.SearchMultiRetCol(1), "dd/mm/yyyy") '-- Book Vdt
                msGrid.TextMatrix(msGrid.Row, 3) = RsBk!ItName  '-- Count
                msGrid.TextMatrix(msGrid.Row, 6) = .SearchMultiRetCol(4)  '-- Bal Bag
                msGrid.TextMatrix(msGrid.Row, 7) = .SearchMultiRetCol(5) '-- Bal Wt
                msGrid.TextMatrix(msGrid.Row, 8) = RsBk!vsubRt  '-- Rate
                msGrid.TextMatrix(msGrid.Row, 9) = RsBk!vsubRTPER  '--Rate Per
                msGrid.TextMatrix(msGrid.Row, 16) = RsBk!VSubItCtrlNo  '--  P ItCtrl No =Book ItCtrl No
                CalGridAmt
                CalAmount
            End If
        End If
    End With
End If
End Sub



