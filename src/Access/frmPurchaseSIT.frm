VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Begin VB.Form frmPurchaseSIT 
   BackColor       =   &H00C2E3FC&
   BorderStyle     =   0  'None
   Caption         =   "Purchase Entry"
   ClientHeight    =   9000
   ClientLeft      =   0
   ClientTop       =   15
   ClientWidth     =   11415
   FillColor       =   &H000040C0&
   FillStyle       =   0  'Solid
   Icon            =   "frmPurchaseSIT.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   9000
   ScaleWidth      =   11415
   ShowInTaskbar   =   0   'False
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
      Index           =   39
      Left            =   1560
      TabIndex        =   45
      Top             =   8400
      Visible         =   0   'False
      Width           =   5595
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
      Index           =   51
      Left            =   120
      TabIndex        =   132
      Text            =   "51"
      Top             =   1440
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
      Index           =   52
      Left            =   0
      TabIndex        =   131
      Text            =   "52"
      Top             =   0
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
      Index           =   53
      Left            =   0
      TabIndex        =   130
      Text            =   "53"
      Top             =   2520
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
      Index           =   54
      Left            =   120
      TabIndex        =   129
      Text            =   "54"
      Top             =   2040
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
      Index           =   55
      Left            =   0
      TabIndex        =   128
      Text            =   "55"
      Top             =   960
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
      Left            =   10500
      TabIndex        =   122
      Text            =   "36"
      Top             =   9120
      Visible         =   0   'False
      Width           =   450
   End
   Begin VB.CheckBox Chk 
      BackColor       =   &H00ECFFFF&
      Caption         =   "Is E2 Purchase ?"
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
      Index           =   1
      Left            =   7440
      TabIndex        =   121
      Top             =   2160
      Width           =   2055
   End
   Begin VB.TextBox txttmpAmt 
      Alignment       =   1  'Right Justify
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   4500
      TabIndex        =   23
      Text            =   "txttmpAmt"
      Top             =   6600
      Width           =   990
   End
   Begin VB.TextBox txttmpSubTotDiff 
      Alignment       =   1  'Right Justify
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   6510
      TabIndex        =   119
      Text            =   "txttmpSubTotDiff"
      Top             =   6600
      Width           =   1005
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
      Height          =   330
      Index           =   49
      Left            =   8775
      TabIndex        =   38
      Text            =   "49"
      Top             =   5670
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
      Height          =   330
      Index           =   50
      Left            =   8775
      TabIndex        =   40
      Text            =   "50"
      Top             =   6045
      Width           =   1365
   End
   Begin VB.CommandButton CmdBagDetails 
      Caption         =   "Cartoon &Bag Details"
      Height          =   720
      Left            =   10605
      TabIndex        =   118
      Top             =   1845
      Width           =   690
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
      Index           =   8
      Left            =   8100
      TabIndex        =   16
      Text            =   "8"
      Top             =   60
      Visible         =   0   'False
      Width           =   345
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
      Left            =   9780
      TabIndex        =   117
      Text            =   "40"
      Top             =   9300
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
      TabIndex        =   116
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
      TabIndex        =   115
      Text            =   "34"
      Top             =   0
      Visible         =   0   'False
      Width           =   570
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
      Height          =   330
      Left            =   6450
      TabIndex        =   114
      Text            =   "txttmpTotWt"
      Top             =   4230
      Width           =   1065
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
      Height          =   330
      Left            =   4815
      TabIndex        =   113
      Text            =   "txttmpTotBag"
      Top             =   4245
      Width           =   765
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
      Height          =   315
      Index           =   15
      Left            =   5235
      TabIndex        =   112
      Text            =   "15"
      Top             =   4590
      Width           =   1290
   End
   Begin VB.TextBox txtExpensesAmt 
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
      Left            =   9960
      TabIndex        =   70
      Text            =   "txtExpensesAmt"
      Top             =   9180
      Visible         =   0   'False
      Width           =   1365
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
      ItemData        =   "frmPurchaseSIT.frx":058A
      Left            =   5235
      List            =   "frmPurchaseSIT.frx":0594
      TabIndex        =   32
      Text            =   "Cbo(12)"
      Top             =   5295
      Width           =   1290
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
      Index           =   33
      Left            =   10980
      TabIndex        =   109
      Text            =   "33"
      Top             =   9360
      Visible         =   0   'False
      Width           =   420
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
      Index           =   32
      Left            =   8880
      TabIndex        =   108
      Text            =   "32"
      Top             =   9300
      Visible         =   0   'False
      Width           =   420
   End
   Begin VB.TextBox txttempOS 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Index           =   6
      Left            =   1380
      TabIndex        =   12
      Text            =   "txttempOS(6)"
      ToolTipText     =   "Press F3 for Creation"
      Top             =   1845
      Width           =   4275
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
      Height          =   330
      Index           =   5
      Left            =   1380
      TabIndex        =   14
      Text            =   "txttemp(5)"
      ToolTipText     =   "Press F3 for Creation"
      Top             =   2220
      Width           =   4275
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
      Height          =   330
      Index           =   3
      Left            =   1380
      TabIndex        =   9
      Text            =   "txttemp(3)"
      ToolTipText     =   "Press F3 for Creation"
      Top             =   1095
      Width           =   4275
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
      Height          =   330
      Left            =   1380
      TabIndex        =   10
      Text            =   "txtCrBalance"
      Top             =   1470
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
      Left            =   7035
      TabIndex        =   107
      Text            =   "txtDrBalance"
      Top             =   525
      Width           =   1740
   End
   Begin VB.CommandButton cmdLedger 
      Caption         =   "Ledger"
      Height          =   375
      Left            =   9015
      TabIndex        =   106
      Top             =   525
      Width           =   1095
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
      Index           =   31
      Left            =   8460
      TabIndex        =   105
      Text            =   "31"
      Top             =   9240
      Visible         =   0   'False
      Width           =   420
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
      Left            =   1860
      TabIndex        =   66
      Text            =   "30"
      Top             =   6510
      Width           =   1290
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
      Left            =   1860
      TabIndex        =   64
      Text            =   "29"
      Top             =   6105
      Width           =   1290
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
      Left            =   1860
      TabIndex        =   62
      Text            =   "28"
      Top             =   5685
      Width           =   1290
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
      Left            =   1860
      TabIndex        =   60
      Text            =   "27"
      Top             =   5280
      Width           =   1290
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
      Left            =   1860
      TabIndex        =   58
      Text            =   "26"
      Top             =   4860
      Width           =   1290
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
      Left            =   1860
      TabIndex        =   56
      Text            =   "25"
      Top             =   4440
      Width           =   1290
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
      Height          =   315
      Index           =   24
      Left            =   5235
      TabIndex        =   30
      Text            =   "24"
      Top             =   4950
      Width           =   1290
   End
   Begin VB.TextBox txtGrid 
      BackColor       =   &H80000000&
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
      Left            =   435
      TabIndex        =   19
      Top             =   3360
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
      Left            =   11400
      TabIndex        =   103
      Text            =   "1"
      Top             =   2430
      Visible         =   0   'False
      Width           =   420
   End
   Begin VB.TextBox txt 
      Alignment       =   1  'Right Justify
      Height          =   315
      Index           =   2
      Left            =   11565
      TabIndex        =   94
      Text            =   "2"
      Top             =   1440
      Visible         =   0   'False
      Width           =   525
   End
   Begin VB.TextBox txt 
      Alignment       =   1  'Right Justify
      Height          =   315
      Index           =   5
      Left            =   11565
      TabIndex        =   93
      Text            =   "5"
      Top             =   2640
      Visible         =   0   'False
      Width           =   525
   End
   Begin VB.TextBox txt 
      Alignment       =   1  'Right Justify
      Height          =   315
      Index           =   3
      Left            =   11565
      TabIndex        =   92
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
      Left            =   6180
      Style           =   1  'Graphical
      TabIndex        =   84
      Top             =   9300
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
      Left            =   5160
      Style           =   1  'Graphical
      TabIndex        =   83
      ToolTipText     =   "To Add New Record"
      Top             =   9300
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
      Left            =   10605
      Style           =   1  'Graphical
      TabIndex        =   81
      ToolTipText     =   "Exit"
      Top             =   6570
      Width           =   690
   End
   Begin VB.CommandButton cmdBtn 
      Cancel          =   -1  'True
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
      Left            =   10605
      Style           =   1  'Graphical
      TabIndex        =   74
      Top             =   6075
      Width           =   690
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
      Left            =   10605
      Style           =   1  'Graphical
      TabIndex        =   46
      ToolTipText     =   "To Save Record"
      Top             =   5640
      Width           =   690
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
      Left            =   7920
      Style           =   1  'Graphical
      TabIndex        =   82
      Top             =   9240
      Visible         =   0   'False
      Width           =   780
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
      Left            =   10605
      Style           =   1  'Graphical
      TabIndex        =   73
      ToolTipText     =   "To Delete Record"
      Top             =   5130
      Width           =   690
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
      Left            =   10605
      Style           =   1  'Graphical
      TabIndex        =   72
      ToolTipText     =   "To Modify Record"
      Top             =   4680
      Width           =   690
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
      Left            =   10605
      Style           =   1  'Graphical
      TabIndex        =   71
      ToolTipText     =   "To Add New Record"
      Top             =   4230
      Width           =   690
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
      Left            =   9360
      Style           =   1  'Graphical
      TabIndex        =   75
      Top             =   9120
      Visible         =   0   'False
      Width           =   780
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
      Left            =   10605
      Style           =   1  'Graphical
      TabIndex        =   76
      ToolTipText     =   "To Find Record"
      Top             =   1395
      Width           =   690
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
      Left            =   10965
      Style           =   1  'Graphical
      TabIndex        =   80
      ToolTipText     =   "Last"
      Top             =   960
      Width           =   330
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
      Left            =   10605
      Style           =   1  'Graphical
      TabIndex        =   79
      ToolTipText     =   "First"
      Top             =   960
      Width           =   330
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
      Left            =   10965
      Style           =   1  'Graphical
      TabIndex        =   78
      ToolTipText     =   "Next"
      Top             =   540
      Width           =   330
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
      Left            =   7440
      TabIndex        =   3
      Text            =   "0"
      Top             =   1073
      Width           =   1380
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
      Left            =   10605
      Style           =   1  'Graphical
      TabIndex        =   77
      ToolTipText     =   "Previous"
      Top             =   540
      Width           =   330
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
      Height          =   480
      Index           =   11
      Left            =   7440
      TabIndex        =   6
      Text            =   "txtVou(11)"
      Top             =   1590
      Width           =   1380
   End
   Begin VB.TextBox txt 
      Alignment       =   1  'Right Justify
      Height          =   315
      Index           =   4
      Left            =   11565
      TabIndex        =   91
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
      Locked          =   -1  'True
      TabIndex        =   1
      Text            =   "txttemp(4)"
      ToolTipText     =   "Press F3 for Creation"
      Top             =   525
      Width           =   4875
   End
   Begin VB.TextBox txt 
      Height          =   315
      Index           =   6
      Left            =   11565
      TabIndex        =   90
      Text            =   "6"
      Top             =   3000
      Visible         =   0   'False
      Width           =   435
   End
   Begin VB.TextBox txt 
      Alignment       =   1  'Right Justify
      Height          =   315
      Index           =   7
      Left            =   11565
      TabIndex        =   89
      Text            =   "7"
      Top             =   3360
      Visible         =   0   'False
      Width           =   525
   End
   Begin VB.TextBox txt 
      Alignment       =   1  'Right Justify
      Height          =   315
      Index           =   9
      Left            =   11565
      TabIndex        =   88
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
      Height          =   315
      Index           =   10
      Left            =   8775
      TabIndex        =   21
      Text            =   "10"
      Top             =   4215
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
      Height          =   315
      Index           =   11
      Left            =   8775
      TabIndex        =   44
      Text            =   "11"
      Top             =   6750
      Width           =   1365
   End
   Begin VB.TextBox txt 
      Alignment       =   1  'Right Justify
      Height          =   315
      Index           =   12
      Left            =   11565
      TabIndex        =   87
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
      Left            =   6600
      TabIndex        =   33
      Text            =   "13"
      Top             =   5295
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
      Height          =   330
      Index           =   14
      Left            =   8775
      TabIndex        =   36
      Text            =   "14"
      Top             =   5295
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
      Height          =   330
      Index           =   16
      Left            =   5235
      TabIndex        =   50
      Text            =   "16"
      Top             =   5700
      Width           =   1290
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
      Height          =   315
      Index           =   17
      Left            =   8775
      TabIndex        =   42
      Text            =   "17"
      Top             =   6405
      Width           =   1365
   End
   Begin VB.TextBox txtOS 
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
      Index           =   9
      Left            =   5235
      TabIndex        =   52
      Text            =   "txtOS(9)"
      Top             =   6135
      Width           =   645
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
      Height          =   315
      Index           =   18
      Left            =   8775
      TabIndex        =   25
      Text            =   "18"
      Top             =   4560
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
      Height          =   345
      Index           =   19
      Left            =   8775
      TabIndex        =   27
      Text            =   "19"
      Top             =   4905
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
      Index           =   20
      Left            =   10635
      TabIndex        =   68
      Text            =   "20"
      Top             =   7140
      Visible         =   0   'False
      Width           =   1290
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
      Index           =   21
      Left            =   11985
      TabIndex        =   48
      Text            =   "21"
      Top             =   6465
      Visible         =   0   'False
      Width           =   1425
   End
   Begin VB.TextBox txt 
      Alignment       =   1  'Right Justify
      Height          =   315
      Index           =   22
      Left            =   11565
      TabIndex        =   86
      Text            =   "22"
      Top             =   5160
      Visible         =   0   'False
      Width           =   525
   End
   Begin VB.TextBox txt 
      Alignment       =   1  'Right Justify
      Height          =   315
      Index           =   23
      Left            =   11565
      TabIndex        =   85
      Text            =   "23"
      Top             =   5400
      Visible         =   0   'False
      Width           =   525
   End
   Begin MSMask.MaskEdBox mskDt 
      Height          =   375
      Index           =   2
      Left            =   8970
      TabIndex        =   4
      Top             =   1073
      Width           =   1200
      _ExtentX        =   2117
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
      Left            =   8970
      TabIndex        =   7
      Top             =   1590
      Width           =   1200
      _ExtentX        =   2117
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
   Begin MSMask.MaskEdBox mskOS 
      Height          =   375
      Index           =   10
      Left            =   6000
      TabIndex        =   53
      Top             =   6120
      Width           =   1155
      _ExtentX        =   2037
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
      Bindings        =   "frmPurchaseSIT.frx":05A6
      Left            =   4110
      Top             =   0
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      WindowControlBox=   -1  'True
      WindowMaxButton =   -1  'True
      WindowMinButton =   -1  'True
      PrintFileType   =   15
      PrintFileLinesPerPage=   60
   End
   Begin MSFlexGridLib.MSFlexGrid msGrid 
      Height          =   1455
      Left            =   75
      TabIndex        =   18
      Top             =   2670
      Width           =   11250
      _ExtentX        =   19844
      _ExtentY        =   2566
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
   Begin MSMask.MaskEdBox mskDt 
      Height          =   375
      Index           =   9
      Left            =   8505
      TabIndex        =   17
      Top             =   45
      Visible         =   0   'False
      Width           =   240
      _ExtentX        =   423
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
   Begin VB.Label Label42 
      BackStyle       =   0  'Transparent
      Caption         =   "To  :"
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
      Left            =   1080
      TabIndex        =   136
      Top             =   8400
      Visible         =   0   'False
      Width           =   450
   End
   Begin VB.Label LblPan 
      BackColor       =   &H00FFFFFF&
      Caption         =   "LblPan"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   330
      Left            =   6720
      TabIndex        =   135
      Top             =   7860
      Width           =   3360
   End
   Begin VB.Label CPan 
      BackColor       =   &H00FFFFFF&
      Caption         =   "PAN :"
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
      Left            =   6000
      TabIndex        =   134
      Top             =   7860
      Width           =   795
   End
   Begin VB.Label Ctin 
      BackColor       =   &H00FFFFFF&
      Caption         =   "C.S.T. Tin"
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
      Left            =   300
      TabIndex        =   133
      Top             =   7860
      Width           =   1155
   End
   Begin VB.Label LblCst 
      BackColor       =   &H00FFFFFF&
      Caption         =   "CstTin"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   330
      Left            =   1620
      TabIndex        =   127
      Top             =   7800
      Width           =   3720
   End
   Begin VB.Label LblEmail 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Email"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   330
      Left            =   4200
      TabIndex        =   126
      Top             =   7320
      Width           =   5940
   End
   Begin VB.Label lbl1 
      BackStyle       =   0  'Transparent
      Caption         =   "Email :"
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
      Left            =   3360
      TabIndex        =   125
      Top             =   7320
      Width           =   720
   End
   Begin VB.Label lblCell 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Cell No"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   330
      Left            =   1095
      TabIndex        =   124
      Top             =   7320
      Width           =   2220
   End
   Begin VB.Label lbl1 
      BackStyle       =   0  'Transparent
      Caption         =   "Cell :"
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
      Index           =   1
      Left            =   360
      TabIndex        =   123
      Top             =   7320
      Width           =   600
   End
   Begin VB.Shape Shape9 
      BorderColor     =   &H00000080&
      FillColor       =   &H00EDFDFE&
      FillStyle       =   0  'Solid
      Height          =   1035
      Left            =   240
      Top             =   7200
      Width           =   10005
   End
   Begin VB.Label Label41 
      BackStyle       =   0  'Transparent
      Caption         =   "Total1 Amt(For Calculation)  :"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   510
      Left            =   3375
      TabIndex        =   22
      Top             =   6600
      Width           =   1110
   End
   Begin VB.Shape Shape3 
      FillColor       =   &H00BDDFD1&
      FillStyle       =   0  'Solid
      Height          =   2175
      Left            =   10530
      Shape           =   4  'Rounded Rectangle
      Top             =   450
      Width           =   825
   End
   Begin VB.Label Label40 
      BackStyle       =   0  'Transparent
      Caption         =   "Total && Sub Amt Diff        :"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   510
      Left            =   5520
      TabIndex        =   120
      Top             =   6600
      Width           =   1005
   End
   Begin VB.Label Label31 
      BackStyle       =   0  'Transparent
      Caption         =   "Aft Tax Add:"
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
      Left            =   7635
      TabIndex        =   37
      Top             =   5670
      Width           =   1155
   End
   Begin VB.Label Label24 
      BackStyle       =   0  'Transparent
      Caption         =   "Aft Tax Less:"
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
      Left            =   7635
      TabIndex        =   39
      Top             =   6060
      Width           =   1155
   End
   Begin VB.Label Label14 
      BackStyle       =   0  'Transparent
      Caption         =   "Book. No. :"
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
      Left            =   7560
      TabIndex        =   15
      Top             =   75
      Visible         =   0   'False
      Width           =   465
   End
   Begin VB.Label Label23 
      BackStyle       =   0  'Transparent
      Caption         =   "Total Bag :"
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
      Left            =   3780
      TabIndex        =   111
      Top             =   4245
      Width           =   1440
   End
   Begin VB.Label Label19 
      BackStyle       =   0  'Transparent
      Caption         =   "Weight :"
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
      Left            =   5670
      TabIndex        =   110
      Top             =   4245
      Width           =   750
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Total Amount  :"
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
      Left            =   10215
      TabIndex        =   47
      Top             =   6930
      Visible         =   0   'False
      Width           =   1635
   End
   Begin VB.Label Label8 
      BackStyle       =   0  'Transparent
      Caption         =   "&JV VNo:"
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
      Left            =   10410
      TabIndex        =   67
      Top             =   6930
      Visible         =   0   'False
      Width           =   1245
   End
   Begin VB.Label Label18 
      BackStyle       =   0  'Transparent
      Caption         =   "Total  Amt:"
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
      Left            =   10245
      TabIndex        =   69
      Top             =   7005
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.Shape Shape5 
      BorderColor     =   &H00C00000&
      BorderWidth     =   5
      FillColor       =   &H00C00000&
      Height          =   9030
      Left            =   0
      Top             =   -90
      Width           =   11430
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "&Party   :"
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
      Left            =   450
      TabIndex        =   8
      Top             =   1065
      Width           =   1665
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "M&ill      :"
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
      Left            =   450
      TabIndex        =   13
      Top             =   2220
      Width           =   1635
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "&Broker  : "
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
      Left            =   450
      TabIndex        =   11
      Top             =   1845
      Width           =   870
   End
   Begin VB.Label Label9 
      BackStyle       =   0  'Transparent
      Caption         =   "Purch.  A/C  :"
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
      Left            =   450
      TabIndex        =   0
      Top             =   525
      Width           =   1620
   End
   Begin VB.Label Label39 
      BackStyle       =   0  'Transparent
      Caption         =   "Cost/Bag    :"
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
      Left            =   480
      TabIndex        =   65
      Top             =   6540
      Width           =   1530
   End
   Begin VB.Label Label38 
      BackStyle       =   0  'Transparent
      Caption         =   "L. Cost       :"
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
      Left            =   480
      TabIndex        =   63
      Top             =   6135
      Width           =   1530
   End
   Begin VB.Label Label37 
      BackStyle       =   0  'Transparent
      Caption         =   "Oth. Less    :"
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
      Left            =   480
      TabIndex        =   61
      Top             =   5715
      Width           =   1530
   End
   Begin VB.Label Label36 
      BackStyle       =   0  'Transparent
      Caption         =   "&Oth. Add     :"
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
      Left            =   480
      TabIndex        =   59
      Top             =   5310
      Width           =   1530
   End
   Begin VB.Label Label35 
      BackStyle       =   0  'Transparent
      Caption         =   "&God.Rent    :"
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
      Left            =   480
      TabIndex        =   57
      Top             =   4905
      Width           =   1530
   End
   Begin VB.Label Label34 
      BackStyle       =   0  'Transparent
      Caption         =   "&Freight        :"
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
      Left            =   480
      TabIndex        =   55
      Top             =   4500
      Width           =   1530
   End
   Begin VB.Label Label33 
      BackStyle       =   0  'Transparent
      Caption         =   "  &Landing Cost Details"
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
      Left            =   315
      TabIndex        =   54
      Top             =   4185
      Width           =   2640
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Vou. &No   :"
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
      Left            =   6225
      TabIndex        =   2
      Top             =   1140
      Width           =   990
   End
   Begin VB.Label LblExmptKg 
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
      Left            =   3780
      TabIndex        =   49
      Top             =   5700
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
      Left            =   7635
      TabIndex        =   35
      Top             =   5295
      Width           =   1155
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
      Left            =   7635
      TabIndex        =   20
      Top             =   4215
      Width           =   1155
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
      Left            =   7635
      TabIndex        =   24
      Top             =   4575
      Width           =   1155
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
      Left            =   7635
      TabIndex        =   26
      Top             =   4920
      Width           =   1155
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
      Left            =   7680
      TabIndex        =   43
      Top             =   6750
      Width           =   1095
   End
   Begin VB.Label Label17 
      BackStyle       =   0  'Transparent
      Caption         =   "Ass. Value    :"
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
      Left            =   3780
      TabIndex        =   29
      Top             =   4950
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
      Left            =   8370
      TabIndex        =   104
      Top             =   45
      Width           =   2475
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
      Left            =   7635
      TabIndex        =   41
      Top             =   6405
      Width           =   1155
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
      TabIndex        =   102
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
      TabIndex        =   101
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
      TabIndex        =   100
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
      Left            =   9420
      TabIndex        =   99
      Top             =   9300
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
      Left            =   7500
      TabIndex        =   98
      Top             =   9300
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
      TabIndex        =   97
      Top             =   7860
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.Label LblExmptAmt 
      BackStyle       =   0  'Transparent
      Caption         =   "&Exempt. Amt :"
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
      Left            =   3780
      TabIndex        =   28
      Top             =   4590
      Width           =   1440
   End
   Begin VB.Label Label16 
      BackStyle       =   0  'Transparent
      Caption         =   "&TaxType        : "
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
      Left            =   3780
      TabIndex        =   31
      Top             =   5295
      Width           =   1335
   End
   Begin VB.Label Label15 
      BackStyle       =   0  'Transparent
      Caption         =   "D&ue Days       :"
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
      Height          =   255
      Left            =   3780
      TabIndex        =   51
      Top             =   6120
      Width           =   1335
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
      Left            =   9600
      TabIndex        =   96
      Top             =   9300
      Width           =   1095
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "Bill No        :    "
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
      Height          =   255
      Left            =   6225
      TabIndex        =   5
      Top             =   1650
      Width           =   1095
   End
   Begin VB.Label Label6 
      BackColor       =   &H00C00000&
      Caption         =   "   Sale In Transite  Purchase"
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
      TabIndex        =   95
      Top             =   0
      Width           =   11835
   End
   Begin VB.Shape Shape1 
      FillColor       =   &H00BDDFD1&
      FillStyle       =   0  'Solid
      Height          =   2880
      Left            =   10530
      Shape           =   4  'Rounded Rectangle
      Top             =   4170
      Width           =   825
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
      Left            =   7230
      TabIndex        =   34
      Top             =   5295
      Width           =   255
   End
   Begin VB.Shape Shape4 
      BorderColor     =   &H00000080&
      FillColor       =   &H00EDFDFE&
      FillStyle       =   0  'Solid
      Height          =   2955
      Left            =   3315
      Top             =   4170
      Width           =   7020
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H00000080&
      FillColor       =   &H00EDFDFE&
      FillStyle       =   0  'Solid
      Height          =   1620
      Left            =   330
      Top             =   1020
      Width           =   5505
   End
   Begin VB.Shape Shape6 
      BorderColor     =   &H00000080&
      FillColor       =   &H00EDFDFE&
      FillStyle       =   0  'Solid
      Height          =   555
      Left            =   330
      Top             =   420
      Width           =   10005
   End
   Begin VB.Shape Shape7 
      BorderColor     =   &H00000080&
      FillColor       =   &H00EDFDFE&
      FillStyle       =   0  'Solid
      Height          =   1620
      Left            =   6030
      Top             =   1020
      Width           =   4305
   End
   Begin VB.Shape Shape8 
      BorderColor     =   &H00000080&
      FillColor       =   &H00EDFDFE&
      FillStyle       =   0  'Solid
      Height          =   2955
      Left            =   300
      Top             =   4170
      Width           =   2970
   End
End
Attribute VB_Name = "frmPurchaseSIT"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Const tblName As String = "tblAddLess"
Const IdField As String = "Vno"
Const OrderField As String = "Vdt,Vno"
Const MaxNo As Long = 36
Const OtherTxtFromNo As Long = 49 '-- Other Txt Ctrl From No
Const OtherTxtToNo As Long = 50 '-- Other Txt Ctrl Less No
Dim Rstbl As Recordset '-- TblAddLess
Dim rstblSub As Recordset '-- TblVouSub
Dim varBookMark As Variant
Dim FirstTimeFlag As Boolean
Dim rsVou As Recordset '--TblVoucher
Dim rsOS As Recordset '--TblOutStanding
Dim OldgFrmTypeStr As String
Dim FormAction As Integer
Dim frmTypeStr As String
Dim tmpAutoLRNo As String '- Auto LR No
'Dim tmpAutoLRNo As Long '- Auto LR No
Dim i As Long

Private Sub Cbo_KeyPress(Index As Integer, KeyAscii As Integer)
If Index = 12 Then KeyAscii = 0 '--Tax
End Sub
Private Sub cbo_LostFocus(Index As Integer)
If Index = 12 And FormAction = vbDataActionAddNew Then '--Tax %
    txt(13) = GProcGetColumnValue("tblMastNarration", "Narration", Cbo(12), "S", "MastTaxRate", "N")
    txt(13) = Format(txt(13), FStr)
End If
End Sub

'-- Cartoon Bag Detail Entry
Private Sub CmdBagDetails_Click()
gfrmBagInwVNo = txt(0)
gfrmBagInwVType = frmTypeStr
frmBagEntry.Show
End Sub
'Dim GridFlag As Boolean
Private Sub cmdBtn_Click(Index As Integer)
'On Error GoTo ErrorRoutine
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
    gClsSearch.SearchMultiField "tblAddLess,TblMastAccount", "Vno,AcName,Vdt,AdInvNo,AdGpDt", Array("Vno", "Party", "Vdt", "Bill No", "Bill Dt"), Array(txt(0).Width, 4000, 1080, 1000, 1200), " tblAddLess.AdAcCrCode=TblMastAccount.AcCode and vtype='" & frmTypeStr & "' and VYear=" & gCYear, "", "Vno,Vdt", 0, 0
    If gClsSearch.SearchMultiRetCol(0) <> "" Then
        Rstbl.MoveFirst
        Rstbl.Find "Vno = " & gClsSearch.SearchMultiRetCol(0)
        gClsSearch.SearchMultiRetCol(0) = ""
        ReadFields
    End If
Case 6 'Add
    FormAction = vbDataActionAddNew
    txt(55) = "A"
    Call GProcActivateControls(Me, True)
    SetControlEd
    Call GProcSetButtons(Me, vbDataActionAddNew)
    Call GProcClearForm(Me, Rstbl, MaxNo, True, OtherTxtFromNo, OtherTxtToNo)
    If FirstTimeFlag = True Then
        FirstTimeFlag = False
    Else
        txt(0).SetFocus
    End If
    txt(0) = GProcGenerateId("TblAddLess", "Vno", "Vtype='" & frmTypeStr & "' and VYear=" & gCYear)
    mskDt(2) = GProcGenerateVDt("tblAddLess", "Vdt", "Vtype='" & frmTypeStr & "' and VYear=" & gCYear)
    '--A/C
    txtTemp(4) = GProcGetColumnValue("tblMastAccount", "AcCode", CStr(gAcCodePT), "N", "AcName", "S")
    FillCombo
    SetGrid
Case 7 'Modify
    FormAction = vbDataActionUpdate
    Hank = ""
    'Call GProcChkPass(frmTypeStr, FormAction, txt(0), txt(2), txtTemp(3), gUserName, Hank)
    'If gChkPassTrueFalse = True Then
        'If mURecModi = True Then
        '    SearchHelp
        '    DoEvents
        txt(55) = "M"
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
                     msGrid.Enabled = True '-Grid
                     txtVou(11).Enabled = True '--Bill No
                     mskVou(12).Enabled = True '--Bill Dt
                     txttempOS(6).Enabled = True
                     For i = 25 To 30 '-Landing Cost
                        txt(i).Enabled = True
                     Next i
                     txt(15).Enabled = True '-Ex amt
                     txt(24).Enabled = True '-Ass value
                     Cbo(12).Enabled = True '-Tax Type
                     txt(13).Enabled = True '-Tax %
                     txt(16).Enabled = True '-Ex /Kg
                     txtOS(9).Enabled = True '-Due days
                    ' txt(21).Enabled = True '-Tot amt
                     txt(18).Enabled = True 'add
                     txt(19).Enabled = True 'Less
                     txt(14).Enabled = True 'Tax amt
                     txt(17).Enabled = True '-Roff
                     Chk(1).Enabled = True
                     SetControlEd
                End If
            End If
        'End If
    'Else
    '   FormAction = vbDataActionClose
    'End If
Case 8 'Delete
    FormAction = vbDataActionDelete
    Hank = ""
    Call GProcChkPass(frmTypeStr, FormAction, txt(0), txt(2), txtTemp(3), gUserName, Hank)
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
    txt(52) = CDate(Date)
    If ValidateData = True Then
        If Trim(txtGrid) <> "" Then
            msGrid.Text = txtGrid
            DoEvents
            txtGrid = ""
        End If
        txtGrid.Visible = False
       ' If FormAction = vbDataActionAddNew Then txt(0) = GProcGenerateId(tblName, IdField, "vtype='" & FrmTypeStr & "' and VYear=" & gCYear)
        CalAmount
        FillTxtFromTemp
        '-- Save Record
        GProcSaveRecord Me, Rstbl, FormAction, MaxNo, OtherTxtFromNo, OtherTxtToNo
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
        msGrid.Col = 0 '-----
        msGrid.Row = 1 '-----
        txtGrid.Move msGrid.CellLeft + msGrid.Left, msGrid.CellTop + msGrid.Top, msGrid.CellWidth, msGrid.CellHeight '-----
    End If
Case 14 'Cancel
    txtGrid.Visible = False
    If Rstbl.RecordCount <= 0 Then
        Call GProcClearForm(Me, Rstbl, MaxNo, True, OtherTxtFromNo, OtherTxtToNo)
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
        Call GProcClearForm(Me, Rstbl, MaxNo, True, OtherTxtFromNo, OtherTxtToNo)
        ReadFields
        Call GProcActivateControls(Me, False)
        Call GProcSetButtons(Me, True)
    End If
    If FormAction = vbDataActionAddNew Then
        If cmdBtn(6).Visible = True Then cmdBtn(6).SetFocus
    Else
        If cmdBtn(7).Visible = True Then cmdBtn(7).SetFocus
    End If
    FormAction = vbDataActionCancel
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
txtTemp(4).Enabled = False '----Purch A/c
'txt(14).Enabled = False '--TaxAmt
txt(14).Enabled = True '--TaxAmt
'txt(17).Enabled = False '--Roff
txt(21).Enabled = False
txt(11).Enabled = False '-- Amt
txt(24).Enabled = False '-- Ass Val
mskOS(10).Enabled = False '--Due Dt
txt(29).Enabled = False '-- LCost
txt(30).Enabled = False '-- Lcost/Bag
'txt(20).Enabled = False '-- JV No
'txtExpensesAmt.Enabled = False '-- Expenses
mskDt(9).Enabled = False '--Book vDt
txttmpSubTotDiff.Enabled = False
txtCrBalance.Enabled = False
txtDrBalance.Enabled = False
txttmpTotBag.Enabled = False
txttmpTotWt.Enabled = False
End Sub
Private Sub DeleteRelatedRecord()
'--- Voucher Sub
gCn.Execute "delete from tblVouSub where vtype='" & frmTypeStr & "' and vno=" & txt(0) & " and " _
& " VYear=" & gCYear
'-- OutStanding
gCn.Execute "delete from tblOutStanding where vtype='" & frmTypeStr & "' and vno=" & txt(0) & " and " _
& " VYear=" & gCYear
'-- Voucher
gCn.Execute "delete from tblVoucher where vtype='" & frmTypeStr & "' and vno=" & txt(0) & " and " _
& " VYear=" & gCYear
'--- JV
'gCn.Execute "delete from tblVoucher where vtype='JV' and vno=" & txt(20) & " and " _
'& " VYear=" & gCYear
End Sub
Private Sub SaveRelatedRecord()
Dim i As Long
Dim MillCode As Long
Dim rsVouSub As Recordset
Dim CrCode As Long
Dim DrCode As Long
Dim BrkCode As Long
Dim TaxAcCode As Long
Dim X As String
Dim BillDt
Dim DueDt
'--- Account Code
MillCode = GProcGetColumnValue("tblMastAccount", "AcName", txtTemp(5), "S", "AcCode", "N")
CrCode = GProcGetColumnValue("tblMastAccount", "AcName", txtTemp(3), "S", "AcCode", "N")
DrCode = GProcGetColumnValue("tblMastAccount", "AcName", txtTemp(4), "S", "AcCode", "N")
BrkCode = GProcGetColumnValue("tblMastAccount", "AcName", txttempOS(6), "S", "AcCode", "N")
TaxAcCode = GProcGetColumnValue("tblMastNarration", "Narration", Cbo(12), "S", "MastTaxPurAcCode", "N")
'-- Delete Records
DeleteRelatedRecord
'--- Save Grid Records
'--- TblVouSub
X = "select * from tblVouSub where vtype='" & frmTypeStr & "' and vno=" & txt(0) & " and " _
& " VYear=" & gCYear
GProcRstOpen rsVouSub, X, "O"
With msGrid
    i = 1
    Do While i < .Rows
        rsVouSub.AddNew
        rsVouSub!Vno = CLng(txt(0)) '--Vno(0)
        rsVouSub!Vtype = txt(1) & vbNullString  '--Vtype(1)
        rsVouSub!Vdt = CDate(txt(2)) '--Vdt(2)
        rsVouSub!VSubCtrNo = i '--Sr No(3)
        rsVouSub!vsubBookNo = CLng(.TextMatrix(i, 1)) '--Book No
        '--Book Dt
        If CLng(.TextMatrix(i, 1)) = 0 Then '-- if Book No =0
            .TextMatrix(i, 2) = "" '--Book Dt =Null
            .TextMatrix(i, 19) = 0 ' Booking It ctrl No =0
        End If
        If IsDate(.TextMatrix(i, 2)) Then
            rsVouSub!vsubBookDt = .TextMatrix(i, 2)
        Else
            rsVouSub!vsubBookDt = Null
        End If
        '--Itcode(4)
        rsVouSub!vsubITCODE = GProcGetColumnValue("tblMastItem", "ItName", .TextMatrix(i, 3), "S", "ItCode", "N", " ItMillCode=" & MillCode) '--Item Code
        rsVouSub!vsubBAG = CLng(.TextMatrix(i, 9)) '--Bag(5)
        rsVouSub!vsubwt = CDbl(.TextMatrix(i, 10)) '-- Wt(6)
        rsVouSub!vsubRt = CDbl(.TextMatrix(i, 11)) '--Rate(7)
        rsVouSub!vsubRTPER = CDbl(.TextMatrix(i, 12)) '--Rate Per(8)
        rsVouSub!vsubAMT = CDbl(.TextMatrix(i, 13)) '--Amt(9)
        rsVouSub!vsubLOTNO = .TextMatrix(i, 5) '--Lot No(10)
        '--Godown Code(17)
        rsVouSub!VSubGodownCode = GProcGetColumnValue("tblMastNarration", "Narration", .TextMatrix(i, 4), "S", "NarrCode", "N", "")
        '--- It Ctrl No(11)
        If CLng(.TextMatrix(i, 14)) = 0 Then 'FormAction = vbDataActionAddNew Then
            rsVouSub!VSubItCtrlNo = GProcGenerateId("tblVouSub", "VSubItCtrlNo", "VYear=" & gCYear)  '--It Ctrl No
        Else
            rsVouSub!VSubItCtrlNo = CLng(.TextMatrix(i, 14))
        End If
        '-- P It Ctrl No =Booking It ctrl No
        rsVouSub!VSubPItCtrlNo = CLng(.TextMatrix(i, 19))
        rsVouSub!VSubIssBag = CLng(.TextMatrix(i, 15)) '--Issue Bag(13)
        rsVouSub!VSubIsswt = CDbl(.TextMatrix(i, 16)) '-- Issue Wt(14)
        rsVouSub!VSubTmpBag = CLng(.TextMatrix(i, 17)) '--Temp Bag(15)
        rsVouSub!VSubTmpWt = CDbl(.TextMatrix(i, 18)) '-- Temp Wt(16)
        rsVouSub!VYear = txt(23) '--VYear(18)
        rsVouSub!vsubLRNO = .TextMatrix(i, 6) & vbNullString  '-- LR No
        If IsDate(.TextMatrix(i, 7)) Then
            rsVouSub!vsubLRDate = CDate(.TextMatrix(i, 7)) '-- LR Date
        Else
            rsVouSub!vsubLRDate = "" '-- LR Date
        End If
        rsVouSub!VSubLRAmt = CDbl(.TextMatrix(i, 8)) '-- LR Amt
        rsVouSub.Update
        i = i + 1
     Loop
 End With
 rsVouSub.Close
'======================Voucher Creation
If IsDate(mskVou(12)) Then
    BillDt = mskVou(12)
Else
    BillDt = ""
End If
'--
If GProcGetColumnValue("tblMastAccount", "AcCode", CStr(TaxAcCode), "N", "AcName", "S") = "-" Then
    '-- Bill Amt
    GProcCreateVoucher FormAction, txt(0), 1, txt(2), txt(1), CrCode, DrCode, txt(11), txtVou(11), BillDt, "", "Bill No " & txtVou(11) & " Dtd. " & BillDt
    GProcCreateVoucher FormAction, txt(0), 2, txt(2), txt(1), DrCode, CrCode, -txt(11), txtVou(11), BillDt, "", "Bill No " & txtVou(11) & " Dtd. " & BillDt
Else '-- If Tax A/c
    '-- Bill -Tax Amt
    GProcCreateVoucher FormAction, txt(0), 1, txt(2), txt(1), CrCode, DrCode, txt(11) - txt(14), txtVou(11), BillDt, "", "Bill No " & txtVou(11) & " Dtd. " & BillDt
    GProcCreateVoucher FormAction, txt(0), 2, txt(2), txt(1), DrCode, CrCode, -(txt(11) - txt(14)), txtVou(11), BillDt, "", "Bill No " & txtVou(11) & " Dtd. " & BillDt
    '-- Tax Amt
    If CDbl(txt(14)) > 0 Then
        GProcCreateVoucher FormAction, txt(0), 3, txt(2), txt(1), CrCode, TaxAcCode, txt(14), txtVou(11), BillDt, "", "Bill No " & txtVou(11) & " Dtd. " & BillDt
        GProcCreateVoucher FormAction, txt(0), 4, txt(2), txt(1), TaxAcCode, CrCode, -txt(14), txtVou(11), BillDt, "", "Bill No " & txtVou(11) & " Dtd. " & BillDt
    End If
End If
'--
'----------- JV
'gCn.Execute "delete from tblVoucher where Vtype='JV' and Vyear=" & gCYear & " and Vno=" & txt(20)
'If txtExpensesAmt > 0 Then
'    If (txt(20) = 0 Or FormAction = vbDataActionAddNew) Then txt(20) = GProcGenerateId("TblVoucher", "Vno", " Vtype='JV' and Vyear=" & gCYear)
'    GProcCreateVoucher FormAction, txt(20), 1, txt(2), "JV", CrCode, gExpensesAcCode, CDbl(txtExpensesAmt), txtVou(11), BillDt, "", "", "", "", "", True
'    GProcCreateVoucher FormAction, txt(20), 2, txt(2), "JV", gExpensesAcCode, CrCode, -CDbl(txtExpensesAmt), txtVou(11), BillDt, "", "", "", "", "", True
'End If
'======================Outstanding Creation
If IsDate(mskOS(10)) Then
    DueDt = mskOS(10)
Else
    DueDt = ""
End If
GProcCreateOutStanding txt(0), CDate(txt(2)), txt(1), txtVou(11), CrCode, BrkCode, CDbl(txt(11)), CLng(txtOS(9)), DueDt
End Sub
Private Sub cmdLedger_Click()
'On Error GoTo ErrorHandler
frmRptAccount.txt(1) = txtTemp(4)
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
    Dim X As String
    LblEmail.Caption = ""
    lblCell.Caption = ""
    LblCst.Caption = ""
    LblPan.Caption = ""

    Me.Top = 300
    Me.Left = Screen.Width / 2 - Me.Width / 2
    '--- TblAddLess
    frmTypeStr = gfrmTypeStr
    X = "Select * from " & tblName & " where Vtype='" & frmTypeStr & "' and VYear=" & gCYear
    If Trim(OrderField) <> "" Then X = X & " order by " & OrderField
    GProcRstOpen Rstbl, X, "O"
    FirstTimeFlag = True
    '---Grid
    SetGrid
    SetRelatedRS
    '---- Button Visiblity Set As Per User Previlage
    GProcSetButtonVisiblity Me
    cmdBtn_Click (6)
End Sub
Private Sub SetRelatedRS()
Dim X As String
Set rsOS = New Recordset
Set rsVou = New Recordset
'---TblVoucher
X = "Select * from tblVoucher where Vtype='" & frmTypeStr & "' and vno=" & txt(0) & " and VYear=" & gCYear & " order by VCtrNo"
GProcRstOpen rsVou, X, "O"
'---TblOutStanding
X = "Select * from tblOutStanding where Vtype='" & frmTypeStr & "' and vno=" & txt(0) & " and VYear=" & gCYear
GProcRstOpen rsOS, X, "O"
End Sub
Private Sub SetGrid()
Dim X As String
Set rstblSub = New Recordset
X = "select VSubCtrNo,VSubBookNo,VSubBookDt,ItName,Narration,vSubLotNo,VSubLRNo,VSubLRDate,VSubLRAmt,VSubBag,VSubWt" _
& " ,VSubRt,VSubRtPer,VSubAmt,VSubItCtrlNo,VSubIssBag,VSubIssWt,VSubTmpBag,VSubTmpWt,VSubPItCtrlNo from tblvouSub,TblMastNarration,TblMastItem where " _
& " tblvouSub.VsubItCode=TblMastItem.ItCode and tblvouSub.VSubGodownCode=TblMastNarration.NarrCode " _
& " and tblvouSub.vtype='" & frmTypeStr & "' and tblvouSub.vno=" & txt(0) & " and tblvouSub.VYear=" & gCYear & " order by VSubCtrNo"
GProcRstOpen rstblSub, X, "R"
With msGrid
    .FormatString = ">Sr|<Bk.No|<Book Date  |<Count                         |<Transport            |<Lot No |<LR No     |<LR Date     |>LR Amt  |>Bag   |>Weight     |>Rate           |>Per  |>Amount     ||||||"
    For i = 14 To 19
        .ColWidth(i) = 0
    Next i
    If FormAction = vbDataActionAddNew Then
        .Rows = 2
        GridAddNew
    Else
        .Rows = 1
    End If

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
'-- LR Date =billDate
If .Row > 1 Then
    For i = 3 To 13
       If i = 6 Then '- LR No
       Else
        ' .TextMatrix(.Row, 7) = IIf(IsDate(mskVou(12)), mskVou(12), "")
            .TextMatrix(.Row, i) = .TextMatrix(.Row - 1, i)
       End If
    Next i
    txtGrid = .TextMatrix(.Row, 13) '--Last Column Amt
End If
''.TextMatrix(.Row,6 ) = tmpau
End With
End Sub
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
If FormAction = vbDataActionUpdate Then
    MsgBox "You are about to modify the record. Save the record and then exit.", vbInformation + vbOKOnly, Me.Caption
    Cancel = True
    Exit Sub
End If
'mShowRec = Space(5)
Rstbl.Close
End Sub
Private Sub MSGrid_LostFocus()
CalAmount
End Sub
Private Sub mskDt_GotFocus(Index As Integer)
GProcSelectBox Me.ActiveControl
End Sub

Private Sub mskVou_GotFocus(Index As Integer)
GProcSelectBox Me.ActiveControl
End Sub
Private Sub mskVou_LostFocus(Index As Integer)
If Index = 12 Then
    If GProcIsDateValid(mskVou(Index)) = False Then
        mskVou(Index).SetFocus
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
lblCell.Caption = GProcGetColumnValue("tblMastAccount", "AcName", txtTemp(3), "S", "AcContno", "S")
LblEmail.Caption = GProcGetColumnValue("tblMastAccount", "AcName", txtTemp(3), "S", "AcEmail", "S")
LblCst.Caption = GProcGetColumnValue("tblMastAccount", "AcName", txtTemp(3), "S", "AcCST", "S")
LblPan.Caption = GProcGetColumnValue("tblMastAccount", "AcName", txtTemp(3), "S", "AcPan", "S")
End Sub
Private Sub mskOS_GotFocus(Index As Integer)
GProcSelectBox Me.ActiveControl
End Sub
Private Sub mskOS_LostFocus(Index As Integer)
If IsDate(mskOS(Index)) = False Then mskOS(Index) = "__/__/____"
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
Private Sub txttmpAmt_GotFocus()
GProcSelectBox Me.ActiveControl
End Sub
Private Sub txttmpAmt_KeyPress(KeyAscii As Integer)
KeyAscii = GProcNumberWithDecimal(KeyAscii, txttmpAmt)
End Sub
Private Sub txttmpAmt_LostFocus()
CalAmount
End Sub
Private Sub txtOS_GotFocus(Index As Integer)
GProcSelectBox Me.ActiveControl
End Sub
Private Sub txtOS_LostFocus(Index As Integer)
SaveBtnEd
If Index = 9 Then
    txtOS(Index) = Format(txtOS(Index), PStr) '--Due Days
End If
CalAmount
End Sub
Private Sub txtOS_KeyPress(Index As Integer, KeyAscii As Integer)
If KeyAscii = 13 Then Exit Sub
KeyAscii = GProcValidateKey(rsOS, Index, KeyAscii, txtOS(Index).Text)
End Sub
Private Sub txtVou_GotFocus(Index As Integer)
GProcSelectBox Me.ActiveControl
End Sub
Private Sub txtVou_LostFocus(Index As Integer)
SaveBtnEd
End Sub
Private Sub txtVou_KeyPress(Index As Integer, KeyAscii As Integer)
If KeyAscii = 13 Then Exit Sub
KeyAscii = GProcValidateKey(rsVou, Index, KeyAscii, txtVou(Index).Text)
End Sub
Private Sub CheckForNumber()
GProcCheckForNumber Me, Rstbl, MaxNo, pOtherTxtFromNo:=OtherTxtFromNo, pOtherTxtToNo:=OtherTxtToNo
End Sub
Private Sub txt_LostFocus(Index As Integer)
CheckForNumber
SaveBtnEd
CalAmount
End Sub
Private Sub txt_KeyPress(Index As Integer, KeyAscii As Integer)
If KeyAscii = 13 Then Exit Sub
KeyAscii = GProcValidateKey(Rstbl, Index, KeyAscii, txt(Index).Text)
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
'        Case 4 '-- Purchase A/c
'            GName = GProcGetColumnValue("TblMastGroup", "AgCode", 90025, "N", "AgName", "S")
'            GProcShowForm frmMastAccount, frmMain.mnuMstfrm(0), Array(frmMastAccount.txtTemp(2)), Array(GName)
        Case 5 '-- millcode
            GName = GProcGetColumnValue("TblMastGroup", "AgCode", 90029, "N", "AgName", "S")
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
Dim AgCode As Long
Dim BrkAcCode As Long
TopPos = txtTemp(Index).Top + Me.Top + 650
LeftPos = txtTemp(Index).Left + Me.Left
If KeyAscii = 13 Then Exit Sub
Select Case Index
    Case 3  '-- Party   Alias Account
        gClsSearch.SearchMultiField "QryMastaccount", "Account,AcName,AcCity,AgName", Array("Party", "AcName", "City", "Group"), Array(3000, 0, 2000, 3000), "(Agcode=90015 or GpCode=90015 or Agcode=90017 or GpCode=90017 or ( (Agcode=90029 or GpCode=90029) and (AcMillType=4 or AcMillType=5) ) )", Chr(KeyAscii), "acName", LeftPos, TopPos, True
        KeyAscii = 0
        If gClsSearch.SearchMultiRetCol(0) <> "" And gClsSearch.SearchMultiRetCol(1) = "" Then
            txtTemp(Index) = gClsSearch.SearchMultiRetCol(0)
        Else
            If gClsSearch.SearchMultiRetCol(1) <> "" Then txtTemp(Index) = gClsSearch.SearchMultiRetCol(1)
        End If
        AgCode = GProcGetColumnValue("tblMastAccount", "AcName", txtTemp(3), "S", "AgCode", "N")
        If AgCode = 90029 Then
            txtTemp(5) = txtTemp(3) '-- Mill =Party
            txtTemp(5).Enabled = False
        Else
            txtTemp(5).Enabled = True
        End If
        '-- Broker from master
        If txttempOS(6) = "" Then
            BrkAcCode = GProcGetColumnValue("TblMastAccount", "ACName", txtTemp(Index), "S", "AcBrkCode", "N")
            txttempOS(6) = GProcGetColumnValue("TblMastAccount", "ACCode", CStr(BrkAcCode), "N", "ACname", "S")
        End If
        
        
        
        'gClsSearch.SearchMultiField "tblMastaccount,tblMastGroup", "AcName", Array("Party"), Array(txtTemp(Index).Width), " tblMastaccount.AgCode=tblMastGroup.AgCode and " _
        '& " (tblMastGroup.Agcode=15 or GpCode=15 or tblMastGroup.Agcode=17 or GpCode=17 or tblMastGroup.Agcode=29 or GpCode=29 )", Chr(KeyAscii), "acName", LeftPos, TopPos, True
'    Case 4 '-- Purchase A/c  Alias Account
'        gClsSearch.SearchMultiField "QryMastaccount", "Account,AcName,AcCity,AgName", Array("Purchase Account", "AcName", "City", "Group"), Array(3000, 0, 2000, 3000), "(AgCode=90025 or GpCode=90025)", Chr(KeyAscii), "AcName", LeftPos, TopPos, True
'        KeyAscii = 0
'        If gClsSearch.SearchMultiRetCol(0) <> "" And gClsSearch.SearchMultiRetCol(1) = "" Then
'            txtTemp(Index) = gClsSearch.SearchMultiRetCol(0)
'        Else
'            If gClsSearch.SearchMultiRetCol(1) <> "" Then txtTemp(Index) = gClsSearch.SearchMultiRetCol(1)
'        End If
'        'gClsSearch.SearchMultiField "tblMastaccount,tblMastGroup", "AcName", Array("Purchase Account"), Array(txtTemp(Index).Width), " tblMastaccount.AgCode=tblMastGroup.AgCode and (tblMastGroup.AgCode=25 or GpCode=25)", Chr(KeyAscii), "AcName", LeftPos, TopPos, True
    Case 5  '-- Mill  Alias Account
        gClsSearch.SearchMultiField "QryMastaccount", "Account,AcName,AcCity,AgName", Array("Mill", "AcName", "City", "Group"), Array(3000, 0, 2000, 3000), "((AgCode=90029 or GpCode=90029) and (AcMillType=4 or AcMillType=5))", Chr(KeyAscii), "AcName", LeftPos, TopPos, True
        KeyAscii = 0
        If gClsSearch.SearchMultiRetCol(0) <> "" And gClsSearch.SearchMultiRetCol(1) = "" Then
            txtTemp(Index) = gClsSearch.SearchMultiRetCol(0)
        Else
            If gClsSearch.SearchMultiRetCol(1) <> "" Then txtTemp(Index) = gClsSearch.SearchMultiRetCol(1)
        End If
        MillCtrlEd
        'gClsSearch.SearchMultiField "tblMastaccount,tblMastGroup", "AcName", Array("Mill"), Array(txtTemp(Index).Width), " tblMastaccount.AgCode=tblMastGroup.AgCode and (tblMastGroup.AgCode=29 or GpCode=29)", Chr(KeyAscii), "AcName", LeftPos, TopPos, True
'    Case 12 '-- Tax Type
'        gClsSearch.SearchMultiField "tblMastNarration", "Narration,MastTaxRate", Array("Tax Type", "TaxRate"), Array(txttemp(Index).Width, 0), " NarrType='T'", Chr(KeyAscii), "Narration", 0, 0, True
'        KeyAscii = 0
'        If gClsSearch.SearchMultiRetCol(0) <> "" Then
'            txttemp(Index) = gClsSearch.SearchMultiRetCol(0) '--Tax
'            If FormAction = vbDataActionAddNew Then txt(13) = gClsSearch.SearchMultiRetCol(1)  '--Tax %
'        End If
End Select
End Sub
Private Sub MillCtrlEd()
Dim ExemptMill As Long
'-- Is Exempt Mill
ExemptMill = GProcGetColumnValue("TblMastAccount", "AcName", txtTemp(5), "S", "ACIsExemptMill", "N")
If ExemptMill = 1 Then
    If FormAction = vbDataActionAddNew Or FormAction = vbDataActionUpdate Then
        txt(15).Enabled = True
        txt(16).Enabled = True
        LblExmptAmt.Visible = True
        LblExmptKg.Visible = True
        txt(15).Visible = True
        txt(16).Visible = True
    End If
Else
    txt(15).Enabled = False
    txt(16).Enabled = False
    LblExmptAmt.Visible = False
    LblExmptKg.Visible = False
    txt(15).Visible = False
    txt(16).Visible = False
    txt(15) = 0
    txt(16) = 0
End If
End Sub
Private Sub txtTemp_LostFocus(Index As Integer)
lblCell.Caption = GProcGetColumnValue("tblMastAccount", "AcName", txtTemp(3), "S", "AcContno", "S")
LblEmail.Caption = GProcGetColumnValue("tblMastAccount", "AcName", txtTemp(3), "S", "AcEmail", "S")
LblCst.Caption = GProcGetColumnValue("tblMastAccount", "AcName", txtTemp(3), "S", "AcCST", "S")
LblPan.Caption = GProcGetColumnValue("tblMastAccount", "AcName", txtTemp(3), "S", "AcPan", "S")
SaveBtnEd
MillCtrlEd
ShowAccountBalance
End Sub
Private Sub txtTempOS_LostFocus(Index As Integer)
SaveBtnEd
End Sub
Private Sub txttempOS_KeyPress(Index As Integer, KeyAscii As Integer)
Dim LeftPos As Long
Dim TopPos As Long
TopPos = txttempOS(Index).Top + Me.Top + 650
LeftPos = txttempOS(Index).Left + Me.Left
If KeyAscii = 13 Then Exit Sub
If Index = 6 Then '-- Broker
    gClsSearch.SearchMultiField "QryMastaccount", "Account,AcName,AcCity,AgName", Array("Broker", "AcName", "City", "Group"), Array(3000, 0, 2000, 3000), "(AgCode=90016 or GpCode=90016)", Chr(KeyAscii), "AcName", LeftPos, TopPos, True
    KeyAscii = 0
    If gClsSearch.SearchMultiRetCol(0) <> "" And gClsSearch.SearchMultiRetCol(1) = "" Then
        txttempOS(Index) = gClsSearch.SearchMultiRetCol(0)
    Else
        If gClsSearch.SearchMultiRetCol(1) <> "" Then txttempOS(Index) = gClsSearch.SearchMultiRetCol(1)
    End If
End If
End Sub
Private Sub txttempOS_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
Dim GName As String
If KeyCode = vbKeyF3 Then
    Select Case Index
        Case 6 '-- Broker
           GName = GProcGetColumnValue("TblMastGroup", "AgCode", 90016, "N", "AgName", "S")
           GProcShowForm frmMastAccount, frmMain.mnuMstfrm(0), Array(frmMastAccount.txtTemp(2)), Array(GName)
    End Select
End If
End Sub
Private Sub txttempOS_GotFocus(Index As Integer)
GProcSelectBox Me.ActiveControl
End Sub
'---------------- Grid
Private Sub msGrid_Click()
'On Error Resume Next
    msGrid_EnterCell
End Sub
Private Sub msGrid_EnterCell()
If FormAction <> vbDataActionClose Then
    CheckForNumberTxtGrid
    DoEvents
    txtGrid = ""
    DoEvents
    If msGrid.TextMatrix(msGrid.Row, 6) = "" And IsNumeric(tmpAutoLRNo) = True Then msGrid.TextMatrix(msGrid.Row, 6) = tmpAutoLRNo  ''-Auto LR No
    If msGrid.Text <> "" Then txtGrid.Text = msGrid.Text
    '--- Lock TextBox
    '--- Check is Related Gate Pass
    If IsRelatedItCtrlGatePass(msGrid.TextMatrix(msGrid.Row, 14)) = False Then
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
                Case 11 '-- rate
                    txtGrid.Locked = False
                Case Else
                    txtGrid.Locked = False
            End Select
        End With
    Else '-Related Gate Pass
          txtGrid.Locked = True
    End If
End If
msGrid_KeyPress (0)
End Sub
Private Sub msGrid_KeyPress(KeyAscii As Integer)
'--- Check is Related Gate Pass



'If IsRelatedItCtrlGatePass(msGrid.TextMatrix(msGrid.Row, 14)) = True Then Exit Sub


With msGrid
    If KeyAscii = 27 Then '----Esc Key
        KeyAscii = 0
        If (Val(.TextMatrix(.Row, 9)) = 0) And .Rows > 2 Then
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
        If .Col = 13 Then 'rstblSub.Fields.Count - 2 Then
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
If msGrid.Col < rstblSub.Fields.Count - 1 Then
    txtGrid.Move msGrid.CellLeft + msGrid.Left, msGrid.CellTop + msGrid.Top, msGrid.CellWidth, msGrid.CellHeight
End If
txtGrid.Visible = True
txtGrid.Enabled = True
DoEvents
If KeyAscii <> 0 Then
    Select Case rstblSub.Fields(msGrid.Col).Type
        Case adChar, adVarWChar
            If msGrid.Col = 1 Or msGrid.Col = 3 Or msGrid.Col = 4 Then '-- Count ,godown
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
    txtGrid.Text = ""
    txtGrid.Visible = False
    If IsNumeric(msGrid.TextMatrix(msGrid.Row, 6)) = True Then ''-Auto LR No
        tmpAutoLRNo = CDbl(msGrid.TextMatrix(msGrid.Row, 6)) + 1
    Else
        tmpAutoLRNo = tmpAutoLRNo
    End If
End If
DoEvents
'If msGrid.Col >= 13 Then
    CalGridAmt
    CalAmount
'End If
End Sub
Private Sub CalGridAmt()
Dim MillCode As Long
Dim StdPack As Double
CheckForNumberGrid
If msGrid.TextMatrix(msGrid.Row, 1) = 0 Then msGrid.TextMatrix(msGrid.Row, 2) = "" '-- Book date
MillCode = GProcGetColumnValue("TblMastAccount", "AcName", txtTemp(5), "S", "Accode", "N")
StdPack = GProcGetColumnValue("TblMastItem", "ITName", msGrid.TextMatrix(msGrid.Row, 3), "S", "ItStdpack", "N", " ItMillCode=" & MillCode)
If msGrid.Row = 0 Then Exit Sub
'-- Wt=Bag * Std Pack
If msGrid.TextMatrix(msGrid.Row, 10) = 0 Then
    msGrid.TextMatrix(msGrid.Row, 10) = msGrid.TextMatrix(msGrid.Row, 9) * StdPack
End If
'-- Rate=Amt*Rate Per/Wt
If CDbl(msGrid.TextMatrix(msGrid.Row, 10)) > 0 Then
    msGrid.TextMatrix(msGrid.Row, 11) = CDbl(msGrid.TextMatrix(msGrid.Row, 13)) * CDbl(msGrid.TextMatrix(msGrid.Row, 12)) / CDbl(msGrid.TextMatrix(msGrid.Row, 10))
Else
    msGrid.TextMatrix(msGrid.Row, 11) = 0
End If

If CDbl(msGrid.TextMatrix(msGrid.Row, 12)) > 0 Then
   msGrid.TextMatrix(msGrid.Row, 13) = CDbl(msGrid.TextMatrix(msGrid.Row, 11)) * CDbl(msGrid.TextMatrix(msGrid.Row, 10)) / CDbl(msGrid.TextMatrix(msGrid.Row, 12))
Else
   msGrid.TextMatrix(msGrid.Row, 13) = CDbl(msGrid.TextMatrix(msGrid.Row, 11)) * CDbl(msGrid.TextMatrix(msGrid.Row, 10))
End If
''''If CDbl(msGrid.TextMatrix(msGrid.Row, 10)) > 0 Then
''''    'msGrid.TextMatrix(msGrid.Row, 11) = CDbl(msGrid.TextMatrix(msGrid.Row, 8)) * CDbl(msGrid.TextMatrix(msGrid.Row, 9)) / CDbl(msGrid.TextMatrix(msGrid.Row, 10))
''''    msGrid.TextMatrix(msGrid.Row, 11) = CDbl(msGrid.TextMatrix(msGrid.Row, 13)) * CDbl(msGrid.TextMatrix(msGrid.Row, 12)) / CDbl(msGrid.TextMatrix(msGrid.Row, 10))
''''Else
''''    'msGrid.TextMatrix(msGrid.Row, 11) = CDbl(msGrid.TextMatrix(msGrid.Row, 8)) * CDbl(msGrid.TextMatrix(msGrid.Row, 9))
''''    'msGrid.TextMatrix(msGrid.Row, 9) = CDbl(msGrid.TextMatrix(msGrid.Row, 8)) * CDbl(msGrid.TextMatrix(msGrid.Row, 9))
''''    msGrid.TextMatrix(msGrid.Row, 11) = 0
''''End If
'''msGrid.TextMatrix(msGrid.Row, 11) = Format(msGrid.TextMatrix(msGrid.Row, 11), LStr)
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
                '-- LR Date =billDate
                If i = 7 And IsDate(msGrid.TextMatrix(RowIndex, i)) = False Then
                    msGrid.TextMatrix(RowIndex, i) = IIf(IsDate(mskVou(12)), mskVou(12), "")
                End If
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
       ' Case adChar, adVarWChar '--String
       '     txtGrid.MaxLength = .Fields(msGrid.Col).DefinedSize
        Case adDate '-- Date
            If IsDate(txtGrid) = False Then
                txtGrid = vbNullString
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
'--- Check is Related Gate Pass
If IsRelatedItCtrlGatePass(msGrid.TextMatrix(msGrid.Row, 14)) = True Then Exit Sub
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
    Case 4 '-- Transport
        OldgFrmTypeStr = gfrmTypeStr
        gfrmTypeStr = "R"
        GProcShowForm frmMastNarration, frmMain.mnuMstfrm(8), "R"
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
'--- Check is Related Gate Pass
'If IsRelatedItCtrlGatePass(msGrid.TextMatrix(msGrid.Row, 14)) = True Then Exit Sub
With msGrid
    If KeyAscii <> 13 And KeyAscii <> 27 And KeyAscii <> 25 Then
        '--- Search
        Select Case .Col
            Case 1 '--Booking
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
                        If IsNumeric(.TextMatrix(.Row, 10)) = False Then .TextMatrix(.Row, 10) = 0
                        If IsNumeric(.TextMatrix(.Row, 12)) = False Then .TextMatrix(.Row, 12) = 0
                        If .TextMatrix(.Row, 12) = 0 Then .TextMatrix(.Row, 12) = gClsSearch.SearchMultiRetCol(3)
                        txtGrid_KeyPress 13
                    End If
                  End If
            Case 4  '-- Transport
                gClsSearch.SearchMultiField "tblMastNarration", "Narration", Array("Transport"), Array(3000), " NarrType='R' ", Chr(KeyAscii), "Narration", LeftPos, TopPos, True
                txtGrid = ""
                KeyAscii = 0
                If gClsSearch.SearchMultiRetCol(0) <> "" Then
                    txtGrid = gClsSearch.SearchMultiRetCol(0)
                    txtGrid_KeyPress 13
                End If
            Case Else
                KeyAscii = GProcValidateKey(rstblSub, msGrid.Col, KeyAscii, txtGrid)
        End Select
   End If
    If KeyAscii = 13 Then
        CheckForNumberTxtGrid
    End If
    '--- Navigate Row ,Col
    If KeyAscii = 13 And .Col < 13 Then 'rstblSub.Fields.Count - 2 Then
        .Col = .Col + 1
        If .Col = 6 Then
            If IsNumeric(txtGrid) = True Then '-Auto LR No
                tmpAutoLRNo = CDbl(txtGrid)
            Else
                tmpAutoLRNo = ""
            End If
        End If
    ElseIf KeyAscii = 13 And .Col = 13 Then '= rstblSub.Fields.Count - 2 Then
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
    If KeyAscii = 27 Then
        KeyAscii = 0
        If (Val(.TextMatrix(.Row, 11)) = 0) And .Rows > 2 Then
            .RemoveItem (.Row)
            txtGrid.Text = ""
        End If
        txtGrid.Visible = False
        txt(18).SetFocus
        Exit Sub
    End If
End With
End Sub

'-- Grid Required
Private Function ValidateGrid() As Boolean
Dim i As Long
Dim MillCode As Long
Dim RowIndex As Long
Dim ItCode As Long
Dim X As String
Dim Rs1 As Recordset
Dim oldFrmTypeStr As String
Dim LRNo As String
Dim pItCtrlNo As Long
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
                    If RowIndex > 1 And Trim(.TextMatrix(RowIndex, 3)) = "" And CDbl(.TextMatrix(RowIndex, 10)) = 0 And CDbl(.TextMatrix(RowIndex, 11)) = 0 Then
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
                        MsgBox "Check Transport.", vbCritical + vbOKOnly, Me.Caption
                        .Col = i
                        msGrid.SetFocus
                        Exit Function
                    End If
                    If GProcGetColumnValue("TblMastNarration", "Narration", .TextMatrix(RowIndex, i), "S", "Narrcode", "N") = 0 Then
                        'MSGrid.SetFocus
                        oldFrmTypeStr = gfrmTypeStr
                        gfrmTypeStr = "R"
                        GProcShowForm frmMastNarration, frmMain.mnuMstfrm(8), Array(frmMastNarration.txt(2)), Array(.TextMatrix(RowIndex, i))
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
                Case 6 '-- LR No
                    If Trim(.TextMatrix(RowIndex, i)) = "" Then
                        MsgBox "Check LR No.", vbCritical + vbOKOnly, Me.Caption
                        .Col = i
                        msGrid.SetFocus
                        Exit Function
                    End If
                Case 7 '-- LR Date
                    If IsDate(.TextMatrix(RowIndex, i)) = False Then
                        MsgBox "Check LR Date.", vbCritical + vbOKOnly, Me.Caption
                        .Col = i
                        msGrid.SetFocus
                        Exit Function
                    End If
                Case 8 '-- LR Amt
'                    If CDbl(.TextMatrix(RowIndex, i)) = 0 Then
'                        MsgBox "Check LR Amount.", vbCritical + vbOKOnly, Me.Caption
'                        .Col = i
'                        MSGrid.SetFocus
'                        Exit Function
'                    End If
                Case 9 '-- Bag
                    If CDbl(.TextMatrix(RowIndex, i)) = 0 Then
                        MsgBox "Check Bag.", vbCritical + vbOKOnly, Me.Caption
                        .Col = i
                        msGrid.SetFocus
                        Exit Function
                    End If
'                    '-- If Bag < Issue Bag
'                    If CDbl(.TextMatrix(RowIndex, 7)) < CDbl(.TextMatrix(RowIndex, 13)) Then
'                        MsgBox "Check Bag.Bag Less than issued.", vbCritical + vbOKOnly, Me.Caption
'                        .Col = i
'                        msGrid.SetFocus
'                        Exit Function
'                    End If
                Case 10 '-- Wt
                    If CDbl(.TextMatrix(RowIndex, i)) = 0 Then
                        MsgBox "Check Weight.", vbCritical + vbOKOnly, Me.Caption
                        .Col = i
                        msGrid.SetFocus
                        Exit Function
                    End If
'                    '-- If Wt < Issue wt
'                    If CDbl(.TextMatrix(RowIndex, 8)) < CDbl(.TextMatrix(RowIndex, 14)) Then
'                        MsgBox "Check Wt.Weight Less than issued.", vbCritical + vbOKOnly, Me.Caption
'                        .Col = i
'                        msGrid.SetFocus
'                        Exit Function
'                    End If
                Case 13 '-- amount
                    If CDbl(.TextMatrix(RowIndex, i)) = 0 Then
                        MsgBox "Check amount.", vbCritical + vbOKOnly, Me.Caption
                        .Col = i
                        msGrid.SetFocus
                        Exit Function
                    End If
           End Select
        Next i
        '--- Duplicate LR No ,It Ctrl  No (Booking)
        LRNo = .TextMatrix(RowIndex, 6)
        pItCtrlNo = .TextMatrix(RowIndex, 19)
        For i = 1 To .Rows - 1
            '--- Duplicate LR No
            If i <> RowIndex And .TextMatrix(i, 6) = LRNo Then
                MsgBox "Duplicate LR No.", vbCritical + vbOKOnly, Me.Caption
                .Col = 6
                msGrid.SetFocus
                Exit Function
            End If
            '-- Duplicate Booking Count Selection
 '           If i <> RowIndex And .TextMatrix(i, 19) = pItCtrlNo And pItCtrlNo > 0 Then
  '              MsgBox "Duplicate Booking Count Selection.", vbCritical + vbOKOnly, Me.Caption
   '             .Col = 1
    '            msGrid.SetFocus
     '           Exit Function
      '      End If
        Next i
        '-- Invalid Booking
        If CLng(.TextMatrix(RowIndex, 1)) > 0 Then
            MillCode = GProcGetColumnValue("TblMastAccount", "AcName", txtTemp(5), "S", "Accode", "N")
            ItCode = GProcGetColumnValue("tblMastItem", "ItName", .TextMatrix(RowIndex, 3), "S", "ItCode", "N", " ItMillCode=" & MillCode)
            X = "Select * from tblvouSub where Vno=" & CLng(.TextMatrix(RowIndex, 1)) & " and  " _
            & " VSUBITCODE=" & ItCode & " and  vsubitctrlno=" & .TextMatrix(RowIndex, 19) & " and Vtype='OS' and "
            If gBackEndDB = gBackEndAccess Then
                X = X + "  vdt=cdate('" & CDate(.TextMatrix(RowIndex, 2)) & "')"
'                X = X + "  vdt=cdate('" & Format(CDate(.TextMatrix(RowIndex, 2)), "dd/mm/yyyy") & "')"
            Else '-- Oracle
                X = X + "  vdt= to_date('" & CDate(.TextMatrix(RowIndex, 2)) & "','dd/MM/yyyy')"
            End If
            GProcRstOpen Rs1, X, "R"
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
Call GProcFillCombo(Cbo(12), "select Narration from tblMastNarration where NarrType ='T' order by Narration")
Cbo(12) = GProcGetColumnValue("tblMastNarration", "NarrCode", CStr(gTaxCodePT), "N", "Narration", "S")
If FormAction = vbDataActionAddNew Then
    txt(13) = GProcGetColumnValue("tblMastNarration", "Narration", Cbo(12), "S", "MastTaxRate", "N")
    txt(13) = Format(txt(13), FStr)
End If
End Sub
Private Sub SaveBtnEd()
'--- Vno,vdt,Party,A/c,Mill,Amount
If CDbl(txt(0)) = 0 Or GProcIsDateValid(mskDt(2)) = False Or txtTemp(3) = "" Or txtTemp(5) = "" Or txtTemp(4) = "" Or CDbl(txt(11)) = 0 Then
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
Dim ExpAcCode As Long
    MillCtrlEd
''    If mskDt(2) >= CDate("25/03/2011") Then
''        MsgBox "Demo Experied Please Contact on MBL - 9922997663", vbInformation + vbOKOnly, Me.Caption
''        txt(0).SetFocus
''        Exit Function
''    End If

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
        '--Create Broker
        If txttempOS(6) = "" Then txttempOS(6) = "-"
        If GProcGetColumnValue("tblMastAccount", "AcName", txttempOS(6), "S", "AcCode", "N") = 0 Then
           GName = GProcGetColumnValue("TblMastGroup", "AgCode", 90016, "N", "AgName", "S")
           GProcShowForm frmMastAccount, frmMain.mnuMstfrm(0), Array(frmMastAccount.txt(1), frmMastAccount.txtTemp(2)), Array(txttempOS(6), GName)
           Exit Function
        End If
    '---- Amount
    If CDbl(txt(11)) <= 0 Then
        MsgBox "Check Voucher Amount.", vbInformation + vbOKOnly, Me.Caption
        Exit Function
    End If
     '-- Expenses A/c
'    If gExpensesAcCode = GProcGetColumnValue("tblmastAccount", "AcName", "-", "S", "AcCode", "N") Then
'        ExpAcCode = 0
'    Else
'        ExpAcCode = gExpensesAcCode
'    End If
'    If txtExpensesAmt > 0 And ExpAcCode = 0 Then
'        MsgBox "Check Purchase Expenses Account.", vbInformation + vbOKOnly, Me.Caption
'        GProcShowForm frmSetting, frmMain.mnuMstSetting(0)
'        Exit Function
'    End If
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
If IsDate(mskOS(10)) = False Then mskOS(10) = "__/__/____" '--Due Dt
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
'--- Bill No ,dt
'txt(6) = txtVou(11) '--gpno
txt(7) = mskVou(12) '--gpDt
txt(36) = txtVou(11)
'-- Cr Code
txt(3) = GProcGetColumnValue("tblMastAccount", "AcName", txtTemp(3), "S", "AcCode", "N")
'-- Dr Code
txt(4) = GProcGetColumnValue("tblMastAccount", "AcName", txtTemp(4), "S", "AcCode", "N")
'--Mill
txt(5) = GProcGetColumnValue("tblMastAccount", "AcName", txtTemp(5), "S", "AcCode", "N")
'-- TaxType
txt(12) = GProcGetColumnValue("tblMastNarration", "Narration", Cbo(12), "S", "NarrCode", "N")
'If txttemp(12) = "" Then txttemp(12) = "-"
'txt(12) = GProcGetColumnValue("tblMastNarration", "Narration", txttemp(12), "S", "NarrCode", "N")
txt(23) = gCYear
' '--Total Amt
'If CDbl(txt(21)) = 0 Then txt(21) = txt(11)
''-- JV
'gCn.Execute "delete from tblVoucher where vtype='JV' and vno=" & txt(20) & " and " _
'& " VYear=" & gCYear
'If CDbl(txt(21)) = CDbl(txt(11)) Then
'    txt(20) = 0
'End If
txt(6) = Chk(1)

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
If IsDate(txt(9)) Then '--Book Dt
    mskDt(9) = txt(9)
Else
    mskDt(9) = "__/__/____"
End If
'-- Cr Code
txtTemp(3) = GProcGetColumnValue("tblMastAccount", "AcCode", txt(3), "N", "AcName", "S")
'-- Dr Code
txtTemp(4) = GProcGetColumnValue("tblMastAccount", "AcCode", txt(4), "N", "AcName", "S")
'--Mill
txtTemp(5) = GProcGetColumnValue("tblMastAccount", "AcCode", txt(5), "N", "AcName", "S")
'-- TaxType
Cbo(12) = GProcGetColumnValue("tblMastNarration", "NarrCode", txt(12), "N", "Narration", "S")
'txttemp(12) = GProcGetColumnValue("tblMastNarration", "NarrCode", txt(12), "N", "Narration", "S")
'-- Fill Related Record
FillRelatedRecord
'-- Expenses
'txtExpensesAmt = CDbl(txt(21)) - CDbl(txt(11))
ShowAccountBalance
MillCtrlEd
'--Total & Sub Diff. Amt   (Total - Sub) amt
If IsNumeric(txttmpAmt) = False Then txttmpAmt = 0
txttmpAmt = txt(11)
txttmpSubTotDiff = Format(CDbl(txttmpAmt) - CDbl(txt(10)), FStr)

Chk(1) = txt(6)
End Sub
Private Sub FillRelatedRecord()
Dim RowIndex As Long
Dim ColIndex As Long
Dim TotBag As Long
Dim TotWt As Double
tmpAutoLRNo = "" ''-Auto LR No
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
                If ColIndex = 9 Then TotBag = TotBag + CLng(msGrid.TextMatrix(RowIndex, 9)) '-- Tot Bag
                If ColIndex = 10 Then TotWt = TotWt + CDbl(msGrid.TextMatrix(RowIndex, 10)) '-- Tot Wt
            Next ColIndex
            RowIndex = RowIndex + 1
            .MoveNext
            If Not .EOF Then msGrid.Rows = msGrid.Rows + 1
        Loop
    End If
End With
txttmpTotBag = TotBag '-- Tot Bag
txttmpTotWt = Format(TotWt, "###0.000") '-- Tot Wt
'--- Fill Related Control
SetRelatedRS
txtVou(11) = rsVou.Fields("vBillNo") & vbNullString     '--Bill No
If IsDate(rsVou.Fields("vBillDt")) Then '--Bill Date
    mskVou(12) = Format(rsVou.Fields("vBilldt"), "dd/mm/yyyy")
Else
    mskVou(12) = "__/__/____"
End If
'--Broker
If rsOS.Fields("OutBrokerCode") > 0 Then txttempOS(6) = GProcGetColumnValue("tblMastAccount", "AcCode", rsOS.Fields("OutBrokerCode"), "N", "AcName", "S")
txtOS(9) = rsOS.Fields("OutDueDays") '--Due Days
If IsDate(rsOS.Fields("OutDueDate")) = True Then '--Due Dt
    mskOS(10) = Format(rsOS.Fields("OutDueDate"), "dd/mm/yyyy")
Else
    mskOS(10) = "__/__/____"
End If
End Sub
Private Sub ReadFields()
FormAction = vbDataActionClose
GProcShowRecord Me, Rstbl, MaxNo, OtherTxtFromNo, OtherTxtToNo
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
Dim TotWt As Double
Dim LCost As Double
Dim MillAcCode As Long
'-- Due Date
If IsNumeric(txtOS(9)) = False Then txtOS(9) = "0"
txtOS(9) = Format(txtOS(9), PStr)
If IsDate(mskDt(2)) Then mskOS(10) = DateAdd("d", txtOS(9), mskDt(2))
'---- Sub amt
CheckForNumberGrid
CheckForNumber
RowIndex = 1
With msGrid
    Do While RowIndex < .Rows
        '-- sub amt
        If Val(.TextMatrix(RowIndex, 13)) > 0 Then
            SubAmt = SubAmt + Val(.TextMatrix(RowIndex, 13))
        End If
        '--Tot Bag
        If Val(.TextMatrix(RowIndex, 9)) > 0 Then
            TotBag = TotBag + Val(.TextMatrix(RowIndex, 9))
        End If
        '--Tot Wt
        If Val(.TextMatrix(RowIndex, 10)) > 0 Then
            TotWt = TotWt + Val(.TextMatrix(RowIndex, 10))
        End If
        RowIndex = RowIndex + 1
    Loop
End With
txttmpTotBag = TotBag
txttmpTotWt = Format(TotWt, "###0.000")
'--- Subamt+Add -Less-Exempt.Amt
AssValue = SubAmt + CDbl(txt(18)) - CDbl(txt(19)) - CDbl(txt(15))
If txt(14) = 0 Then
   TaxAmt = AssValue * CDbl(txt(13)) / 100
Else
   TaxAmt = txt(14)
End If
'--Roff
ROff = txt(17) '= 'CDbl(GProcMakeRounding(Amt) - Amt)
'-- SubAmt+Tax Amt +Add -Less+roff + Aft Tax Add - Aft Tax Less
Amt = SubAmt + TaxAmt + CDbl(txt(18)) - CDbl(txt(19)) + ROff + CDbl(txt(49)) - CDbl(txt(50))
'--G.Rent
'If (Val(txt(26)) = 0) And (TotBag > 0) Then txt(26) = Format(TotBag * 5, FStr)
If (Val(txt(26)) = 0) And (TotBag > 0) And IsDate(mskDt(2)) Then
    If CDate(mskDt(2)) < CDate("05/10/2006") Then '-Date <  CDate("06/10/2006")  Rs 5
        txt(26) = Format(TotBag * 5, FStr)
    Else '- Rs 6
        txt(26) = Format(TotBag * 6, FStr)
    End If
End If
'--Other Add
MillAcCode = GProcGetColumnValue("tblmastAccount", "AcName", txtTemp(5), "S", "AcCode", "N")
If (Val(txt(27)) = 0) And (TotBag > 0) Then txt(27) = Format(GProcGetColumnValue("tblmastitem", "itname", msGrid.TextMatrix(1, 3), "S", "Itbrokrt", "N", " ItMillCode=" & MillAcCode) * TotBag, FStr)
'--- Landing Cost =Bill+Freight+Rent +add-Less
LCost = Amt + CDbl(txt(25)) + CDbl(txt(26)) + CDbl(txt(27)) - CDbl(txt(28))
txt(10) = SubAmt
txt(14) = TaxAmt
txt(24) = AssValue
'txt(17) = Roff
txt(11) = Amt
txt(29) = LCost
If TotBag > 0 Then txt(30) = txt(29) / TotBag
txt(31) = CLng(TotBag) '- Total Bag
'--Total & Sub Diff. Amt   (Total - Sub) amt
If IsNumeric(txttmpAmt) = False Then txttmpAmt = 0
txttmpSubTotDiff = Format(CDbl(txttmpAmt) - CDbl(SubAmt), FStr)
'============= Expenses JV
'txtExpensesAmt = 0
'If CDbl(txt(21)) < Amt Then txt(21) = Amt '--Total Amt
''If CDbl(txt(21)) = 0 Then txt(21) = Amt '--Total Amt
'txtExpensesAmt = CDbl(txt(21)) - Amt '-- Expenses
'If txtExpensesAmt > 0 And txt(20) = 0 Then
'    txt(20) = GProcGenerateId("TblVoucher", "Vno", " Vtype='JV' and Vyear=" & gCYear)
'End If
'txtExpensesAmt = Format(txtExpensesAmt, FStr)
CheckForNumber
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
Dim X As String
'-- Is Gate Pass
X = "Select Vno,Vdt from tblVouSub where Vtype='GP' and " _
& " VSubPItCtrlNo in (select VSubItCtrlNo from tblVouSub where vtype='" & frmTypeStr & "' and vno=" & txt(0) & " and " _
& " VYear=" & gCYear & ")"
GProcRstOpen RsG, X, "R"
If RsG.EOF Then
Else
    IsRelatedGatePass = True
   ' MsgBox "You can not Modify/Delete this record.Related Gate Pass No. " & RsG.Fields(0) & " is exists.", vbCritical + vbOKOnly, App.Title
    Exit Function
End If
RsG.Close
End Function
'-- Is Related It Ctrl GatePass
Private Function IsRelatedItCtrlGatePass(ItCtrlNo As Long) As Boolean
Dim RsG As Recordset
Dim X As String
'-- Is Gate Pass
X = "Select Vno,Vdt from tblVouSub where Vtype='GP' and " _
& " VSubPItCtrlNo = " & ItCtrlNo
GProcRstOpen RsG, X, "R"
If RsG.EOF Then
Else
    IsRelatedItCtrlGatePass = True
    Exit Function
End If
RsG.Close
End Function
'---------- Booking Vs Purchase SIT
Private Sub PrepareBooking(Optional Chr As String)
Dim MillCode As Long
Dim QStr As String
Dim RsBk As Recordset
If IsDate(mskDt(2)) = True Then
    MillCode = GProcGetColumnValue("TblMastAccount", "AcName", txtTemp(5), "S", "Accode", "N")
    QStr = " select tblVousub.VNo,format(tblVousub.vdt,'dd/mm/yyyy'),ItName,vsubbag-(vsubissbag),vsubwt-(vsubisswt),vsubbag,vsubwt,vsubitctrlno,VsubRt from " _
    & " tblVousub,tblmastItem Where " _
    & " Itcode=VSUBITCODE and ItMillCode=" & MillCode & " and " _
    & " ((vsubbag-(vsubissbag))> 0) and tblVousub.vtype ='OS'  and VSUBBOOKISCOMPLETED=0 and "
    If gBackEndDB = gBackEndAccess Then
        QStr = QStr + " tblVousub.vdt<=cdate('" & CDate(mskDt(2)) & "')"
    Else '-- Oracle
        QStr = QStr + " tblVousub.vdt <= to_date('" & CDate(mskDt(2)) & "','dd/MM/yyyy')"
    End If
    
    With gClsSearch
        .SearchMultiField QStr, "", Array("Booking No.", "Date", "Count", "Bal Bag", "Bal Wt.", "Bag", "Wt", "itctrlno", "Rate"), Array(600, 1500, 1500, 1000, 1000, 500, 1000, 0, 1000), "", Chr, "tblVousub.Vdt,tblVousub.vno,VSUBCTRNO", 0, 0, False, , True
        If .SearchMultiRetCol(0) <> "" Then
            QStr = "select tblvousub.*,tblmastItem.ItName from tblvousub,tblMastItem where " _
            & " ItCode=VSUBITCODE and VSUBITCTRLNO= " & CLng(.SearchMultiRetCol(7))
              If GProcRstOpen(RsBk, QStr, "R") > 0 Then
                    txtGrid = ""
                    txtGrid = CLng(.SearchMultiRetCol(0)) '--Book Vno
                    msGrid.TextMatrix(msGrid.Row, 1) = CLng(.SearchMultiRetCol(0)) '--Book Vno
                    msGrid.TextMatrix(msGrid.Row, 2) = Format(.SearchMultiRetCol(1), "dd/mm/yyyy") '-- Book Vdt
                    msGrid.TextMatrix(msGrid.Row, 3) = RsBk!ItName  '-- Count
                    msGrid.TextMatrix(msGrid.Row, 9) = .SearchMultiRetCol(3)  '-- Bal Bag
                    msGrid.TextMatrix(msGrid.Row, 10) = .SearchMultiRetCol(4) '-- Bal Wt
                    msGrid.TextMatrix(msGrid.Row, 19) = RsBk!VSubItCtrlNo  '--  P ItCtrl No =Book ItCtrl No
                    msGrid.TextMatrix(msGrid.Row, 11) = RsBk!vsubRt  '-- Rate
                    msGrid.TextMatrix(msGrid.Row, 12) = RsBk!vsubRTPER  '--Rate Per
                    CalGridAmt
                    CalAmount
              End If
        End If
    End With
    
    
    
    
'    With gClsSearch
'        .SearchMultiField QStr, "", Array("Booking No.", "Date", "Count", "Bal Bag", "Bal Wt.", "Rate", "Wt", "itctrlno"), Array(1000, 1500, 1500, 1000, 1000, 1000, 0, 0), "", Chr, "tblVousub.Vdt,tblVousub.vno,VSUBCTRNO ", 0, 0, False, , True
'        If .SearchMultiRetCol(0) <> "" Then
'            QStr = "select tblvousub.*,tblmastItem.ItName from tblvousub,tblMastItem where " _
'            & " ItCode=VSUBITCODE and VSUBITCTRLNO= " & CLng(.SearchMultiRetCol(7))
'              If GProcRstOpen(RsBk, QStr, "R") > 0 Then
'                    txtGrid = ""
'                    txtGrid = CLng(.SearchMultiRetCol(0)) '--Book Vno
'                    msGrid.TextMatrix(msGrid.Row, 1) = CLng(.SearchMultiRetCol(0)) '--Book Vno
'                    msGrid.TextMatrix(msGrid.Row, 2) = Format(.SearchMultiRetCol(1), "dd/mm/yyyy") '-- Book Vdt
'                    msGrid.TextMatrix(msGrid.Row, 3) = RsBk!ItName  '-- Count
'              '      msGrid.TextMatrix(msGrid.Row, 9) = .SearchMultiRetCol(3)  '-- Bal Bag
'               '     msGrid.TextMatrix(msGrid.Row, 10) = .SearchMultiRetCol(4) '-- Bal Wt
'                    msGrid.TextMatrix(msGrid.Row, 4) = RsBk!vsubRt  '-- Rate
'
'                    msGrid.TextMatrix(msGrid.Row, 19) = RsBk!VSubItCtrlNo  '--  P ItCtrl No =Book ItCtrl No
''                    msGrid.TextMatrix(msGrid.Row, 11) = RsBk!vsubRt  '-- Rate
'                    msGrid.TextMatrix(msGrid.Row, 12) = RsBk!vsubRTPER  '--Rate Per
'                    CalGridAmt
'                    CalAmount
'              End If
'        End If
'    End With
End If
End Sub

