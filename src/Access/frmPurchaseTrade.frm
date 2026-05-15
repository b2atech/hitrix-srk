VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form frmPurchaseTrade 
   BackColor       =   &H00C2E3FC&
   BorderStyle     =   0  'None
   Caption         =   "Purchase Entry"
   ClientHeight    =   8760
   ClientLeft      =   0
   ClientTop       =   15
   ClientWidth     =   10830
   DrawMode        =   10  'Mask Pen
   FillColor       =   &H000040C0&
   FillStyle       =   0  'Solid
   Icon            =   "frmPurchaseTrade.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   8760
   ScaleWidth      =   10830
   ShowInTaskbar   =   0   'False
   Visible         =   0   'False
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
      TabIndex        =   132
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
      Index           =   54
      Left            =   120
      TabIndex        =   131
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
      Index           =   52
      Left            =   0
      TabIndex        =   129
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
      Index           =   51
      Left            =   120
      TabIndex        =   128
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
      Index           =   43
      Left            =   3000
      TabIndex        =   127
      Text            =   "43"
      Top             =   8760
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
      Index           =   42
      Left            =   2280
      TabIndex        =   126
      Text            =   "42"
      Top             =   9720
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
      Index           =   41
      Left            =   2280
      TabIndex        =   125
      Text            =   "41"
      Top             =   9240
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
      Index           =   40
      Left            =   2280
      TabIndex        =   124
      Text            =   "40"
      Top             =   8760
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
      Left            =   1680
      TabIndex        =   123
      Text            =   "39"
      Top             =   9600
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
      Left            =   1680
      TabIndex        =   122
      Text            =   "38"
      Top             =   9240
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
      Left            =   1680
      TabIndex        =   121
      Text            =   "37"
      Top             =   8760
      Visible         =   0   'False
      Width           =   450
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
      Left            =   3840
      TabIndex        =   22
      Text            =   "txttmpAmt"
      Top             =   6615
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
      Left            =   5805
      TabIndex        =   118
      Text            =   "txttmpSubTotDiff"
      Top             =   6615
      Width           =   900
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
      Index           =   16
      Left            =   4485
      TabIndex        =   116
      Text            =   "16"
      Top             =   5880
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
      Height          =   330
      Index           =   49
      Left            =   7920
      TabIndex        =   115
      Text            =   "49"
      Top             =   5730
      Width           =   1335
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
      Index           =   25
      Left            =   1530
      TabIndex        =   112
      Text            =   "25"
      Top             =   4815
      Width           =   1110
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
      Left            =   4485
      TabIndex        =   111
      Text            =   "15"
      Top             =   4785
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
      Index           =   18
      Left            =   7920
      TabIndex        =   24
      Text            =   "18"
      Top             =   4725
      Width           =   1335
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
      Index           =   50
      Left            =   7920
      TabIndex        =   36
      Text            =   "50"
      Top             =   6075
      Width           =   1335
   End
   Begin VB.CommandButton CmdBagDetails 
      Caption         =   "Cartoon &Bag Details"
      Height          =   750
      Left            =   9690
      TabIndex        =   109
      Top             =   2295
      Width           =   855
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
      TabIndex        =   108
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
      Left            =   -30
      TabIndex        =   107
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
      Left            =   -45
      TabIndex        =   106
      Text            =   "34"
      Top             =   645
      Visible         =   0   'False
      Width           =   570
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
      Height          =   315
      Left            =   4140
      TabIndex        =   103
      Text            =   "txttmpTotBag"
      Top             =   4395
      Width           =   675
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
      Height          =   315
      Left            =   5640
      TabIndex        =   102
      Text            =   "txttmpTotWt"
      Top             =   4395
      Width           =   1050
   End
   Begin VB.TextBox txtGrid 
      BackColor       =   &H8000000A&
      BorderStyle     =   0  'None
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
      Left            =   495
      TabIndex        =   18
      Top             =   3210
      Visible         =   0   'False
      Width           =   1035
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
      Left            =   8940
      TabIndex        =   56
      Text            =   "txtExpensesAmt"
      Top             =   8865
      Visible         =   0   'False
      Width           =   2310
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
      ItemData        =   "frmPurchaseTrade.frx":058A
      Left            =   4500
      List            =   "frmPurchaseTrade.frx":0594
      TabIndex        =   30
      Text            =   "Cbo(12)"
      Top             =   5460
      Width           =   1260
   End
   Begin VB.CommandButton cmdLedger 
      Caption         =   "Ledger"
      Height          =   375
      Index           =   1
      Left            =   3630
      TabIndex        =   101
      Top             =   1515
      Width           =   1635
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
      Height          =   360
      Index           =   0
      Left            =   6750
      TabIndex        =   3
      Text            =   "0"
      Top             =   1245
      Width           =   1200
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
      Height          =   360
      Index           =   3
      Left            =   1320
      TabIndex        =   12
      Text            =   "txttemp(3)"
      ToolTipText     =   "Press F3 for Creation"
      Top             =   1110
      Width           =   3945
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
      Left            =   10560
      TabIndex        =   100
      Text            =   "33"
      Top             =   6975
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
      Left            =   10020
      TabIndex        =   99
      Text            =   "32"
      Top             =   6870
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
      Index           =   31
      Left            =   7080
      TabIndex        =   98
      Text            =   "31"
      Top             =   9090
      Visible         =   0   'False
      Width           =   1785
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
      Index           =   30
      Left            =   1530
      TabIndex        =   54
      Text            =   "30"
      Top             =   6690
      Width           =   1110
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
      Index           =   29
      Left            =   1530
      TabIndex        =   52
      Text            =   "29"
      Top             =   6300
      Width           =   1110
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
      Index           =   28
      Left            =   1530
      TabIndex        =   50
      Text            =   "28"
      Top             =   5940
      Width           =   1110
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
      Index           =   27
      Left            =   1530
      TabIndex        =   48
      Text            =   "27"
      Top             =   5565
      Width           =   1110
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
      Index           =   26
      Left            =   1530
      TabIndex        =   46
      Text            =   "26"
      Top             =   5190
      Width           =   1110
   End
   Begin VB.CommandButton cmdLedger 
      Caption         =   "Ledger"
      Height          =   375
      Index           =   0
      Left            =   8085
      TabIndex        =   97
      Top             =   525
      Width           =   1065
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
      Height          =   360
      Left            =   6285
      TabIndex        =   96
      Text            =   "txtDrBalance"
      Top             =   525
      Width           =   1740
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
      Height          =   360
      Left            =   1320
      TabIndex        =   95
      Text            =   "txtCrBalance"
      Top             =   1515
      Width           =   2070
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
      Left            =   4485
      TabIndex        =   28
      Text            =   "24"
      Top             =   5115
      Width           =   1290
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
      TabIndex        =   93
      Text            =   "1"
      Top             =   2430
      Width           =   420
   End
   Begin VB.TextBox txt 
      Alignment       =   1  'Right Justify
      Height          =   315
      Index           =   2
      Left            =   11040
      TabIndex        =   82
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
      TabIndex        =   81
      Text            =   "5"
      Top             =   2640
      Width           =   525
   End
   Begin VB.TextBox txt 
      Alignment       =   1  'Right Justify
      Height          =   315
      Index           =   3
      Left            =   11040
      TabIndex        =   80
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
      Left            =   9960
      Style           =   1  'Graphical
      TabIndex        =   72
      Top             =   8520
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
      Left            =   9840
      Style           =   1  'Graphical
      TabIndex        =   71
      Top             =   7980
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
      Left            =   9690
      Style           =   1  'Graphical
      TabIndex        =   70
      ToolTipText     =   "Exit"
      Top             =   6000
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
      TabIndex        =   62
      Top             =   5400
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
      TabIndex        =   61
      ToolTipText     =   "To Save Record"
      Top             =   4920
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
      Left            =   9960
      Style           =   1  'Graphical
      TabIndex        =   69
      Top             =   7200
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
      TabIndex        =   60
      ToolTipText     =   "To Delete Record"
      Top             =   4245
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
      TabIndex        =   59
      ToolTipText     =   "To Modify Record"
      Top             =   3765
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
      Height          =   420
      Index           =   6
      Left            =   9690
      Style           =   1  'Graphical
      TabIndex        =   58
      ToolTipText     =   "To Add New Record"
      Top             =   3285
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
      Left            =   10080
      Style           =   1  'Graphical
      TabIndex        =   63
      Top             =   7560
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
      TabIndex        =   64
      ToolTipText     =   "To Find Record"
      Top             =   1725
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
      TabIndex        =   66
      ToolTipText     =   "Last"
      Top             =   1245
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
      TabIndex        =   65
      ToolTipText     =   "First"
      Top             =   1245
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
      TabIndex        =   68
      ToolTipText     =   "Next"
      Top             =   765
      Width           =   375
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
      TabIndex        =   67
      ToolTipText     =   "Previous"
      Top             =   765
      Width           =   375
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
      Height          =   360
      Index           =   11
      Left            =   6750
      TabIndex        =   6
      Text            =   "txtVou(11)"
      Top             =   1785
      Width           =   1200
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
      Height          =   360
      Index           =   5
      Left            =   1320
      TabIndex        =   16
      Text            =   "txttemp(5)"
      ToolTipText     =   "Press F3 for Creation"
      Top             =   2340
      Width           =   3945
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
      Height          =   360
      Index           =   6
      Left            =   1320
      TabIndex        =   14
      Text            =   "txttempOS(6)"
      ToolTipText     =   "Press F3 for Creation"
      Top             =   1935
      Width           =   3945
   End
   Begin VB.TextBox txt 
      Alignment       =   1  'Right Justify
      Height          =   315
      Index           =   4
      Left            =   11040
      TabIndex        =   79
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
      Height          =   360
      Index           =   4
      Left            =   1950
      TabIndex        =   1
      Text            =   "txttemp(4)"
      ToolTipText     =   "Press F3 for Creation"
      Top             =   525
      Width           =   4245
   End
   Begin VB.TextBox txt 
      Height          =   315
      Index           =   6
      Left            =   11040
      TabIndex        =   78
      Text            =   "6"
      Top             =   3000
      Visible         =   0   'False
      Width           =   435
   End
   Begin VB.TextBox txt 
      Alignment       =   1  'Right Justify
      Height          =   315
      Index           =   7
      Left            =   11040
      TabIndex        =   77
      Text            =   "7"
      Top             =   3360
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
      Height          =   360
      Index           =   8
      Left            =   6750
      TabIndex        =   9
      Text            =   "8"
      Top             =   2340
      Width           =   1200
   End
   Begin VB.TextBox txt 
      Alignment       =   1  'Right Justify
      Height          =   315
      Index           =   9
      Left            =   11040
      TabIndex        =   76
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
      Height          =   315
      Index           =   10
      Left            =   7920
      TabIndex        =   20
      Text            =   "10"
      Top             =   4395
      Width           =   1335
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
      Left            =   7920
      TabIndex        =   40
      Text            =   "11"
      Top             =   6735
      Width           =   1335
   End
   Begin VB.TextBox txt 
      Alignment       =   1  'Right Justify
      Height          =   315
      Index           =   12
      Left            =   11040
      TabIndex        =   75
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
      Height          =   360
      Index           =   13
      Left            =   5865
      TabIndex        =   31
      Text            =   "13"
      Top             =   5460
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
      Height          =   315
      Index           =   14
      Left            =   7920
      TabIndex        =   34
      Text            =   "14"
      Top             =   5400
      Width           =   1335
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
      Left            =   7920
      TabIndex        =   38
      Text            =   "17"
      Top             =   6405
      Width           =   1335
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
      Height          =   360
      Index           =   9
      Left            =   4485
      TabIndex        =   43
      Text            =   "txtOS(9)"
      Top             =   6225
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
      Height          =   315
      Index           =   19
      Left            =   7920
      TabIndex        =   26
      Text            =   "19"
      Top             =   5055
      Width           =   1335
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
      Left            =   9735
      TabIndex        =   55
      Text            =   "20"
      Top             =   6630
      Visible         =   0   'False
      Width           =   810
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
      Index           =   21
      Left            =   11940
      TabIndex        =   41
      Text            =   "21"
      Top             =   6525
      Width           =   1290
   End
   Begin VB.TextBox txt 
      Alignment       =   1  'Right Justify
      Height          =   315
      Index           =   22
      Left            =   11040
      TabIndex        =   74
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
      TabIndex        =   73
      Text            =   "23"
      Top             =   5400
      Visible         =   0   'False
      Width           =   525
   End
   Begin MSMask.MaskEdBox mskDt 
      Height          =   360
      Index           =   2
      Left            =   8070
      TabIndex        =   4
      Top             =   1245
      Width           =   1140
      _ExtentX        =   2011
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
   Begin MSMask.MaskEdBox mskVou 
      Height          =   360
      Index           =   12
      Left            =   8070
      TabIndex        =   7
      Top             =   1785
      Width           =   1140
      _ExtentX        =   2011
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
   Begin MSMask.MaskEdBox mskDt 
      Height          =   360
      Index           =   9
      Left            =   8070
      TabIndex        =   10
      Top             =   2340
      Width           =   1140
      _ExtentX        =   2011
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
   Begin MSMask.MaskEdBox mskOS 
      Height          =   315
      Index           =   10
      Left            =   5190
      TabIndex        =   44
      Top             =   6225
      Width           =   1155
      _ExtentX        =   2037
      _ExtentY        =   556
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
      Bindings        =   "frmPurchaseTrade.frx":05A6
      Left            =   2820
      Top             =   30
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
      Left            =   240
      TabIndex        =   17
      Top             =   2805
      Width           =   9135
      _ExtentX        =   16113
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
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
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
      Height          =   255
      Left            =   1680
      TabIndex        =   140
      Top             =   7980
      Width           =   3360
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
      Left            =   360
      TabIndex        =   139
      Top             =   7980
      Width           =   1155
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
      Left            =   5100
      TabIndex        =   138
      Top             =   7980
      Width           =   795
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
      Height          =   255
      Left            =   5880
      TabIndex        =   137
      Top             =   7980
      Width           =   3360
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
      TabIndex        =   136
      Top             =   7320
      Width           =   600
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
      Height          =   375
      Left            =   1095
      TabIndex        =   135
      Top             =   7320
      Width           =   2220
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
      TabIndex        =   134
      Top             =   7320
      Width           =   720
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
      Height          =   375
      Left            =   4200
      TabIndex        =   133
      Top             =   7320
      Width           =   5100
   End
   Begin VB.Shape Shape9 
      BorderColor     =   &H00000080&
      FillColor       =   &H00EDFDFE&
      FillStyle       =   0  'Solid
      Height          =   1155
      Left            =   240
      Top             =   7200
      Width           =   9165
   End
   Begin VB.Label Label33 
      BackStyle       =   0  'Transparent
      Caption         =   " &Landing Cost Details"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00004080&
      Height          =   300
      Left            =   315
      TabIndex        =   120
      Top             =   4440
      Width           =   2640
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
      Left            =   2775
      TabIndex        =   21
      Top             =   6615
      Width           =   1110
   End
   Begin VB.Label Label40 
      BackStyle       =   0  'Transparent
      Caption         =   "Total1 && Sub Amt Diff     :"
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
      Left            =   4860
      TabIndex        =   119
      Top             =   6615
      Width           =   1020
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Aft Tax Add :"
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
      Left            =   6735
      TabIndex        =   117
      Top             =   5730
      Width           =   1170
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
      Left            =   6735
      TabIndex        =   23
      Top             =   4725
      Width           =   1095
   End
   Begin VB.Label Label34 
      BackStyle       =   0  'Transparent
      Caption         =   "&Freight       :"
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
      Left            =   390
      TabIndex        =   114
      Top             =   4815
      Width           =   1170
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
      Left            =   3180
      TabIndex        =   113
      Top             =   4785
      Width           =   1440
   End
   Begin VB.Label Label18 
      BackStyle       =   0  'Transparent
      Caption         =   "Total Amt  :"
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
      Left            =   10710
      TabIndex        =   110
      Top             =   6015
      Width           =   1095
   End
   Begin VB.Label Label8 
      BackStyle       =   0  'Transparent
      Caption         =   "Aft Tax Less :"
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
      Left            =   6735
      TabIndex        =   35
      Top             =   6075
      Width           =   1230
   End
   Begin VB.Label Label31 
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
      ForeColor       =   &H00404040&
      Height          =   375
      Left            =   4875
      TabIndex        =   105
      Top             =   4395
      Width           =   750
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
      ForeColor       =   &H00404040&
      Height          =   375
      Left            =   3090
      TabIndex        =   104
      Top             =   4395
      Width           =   1440
   End
   Begin VB.Shape Shape8 
      BorderColor     =   &H00C00000&
      BorderWidth     =   5
      FillColor       =   &H00C00000&
      Height          =   8880
      Left            =   0
      Top             =   -120
      Width           =   10830
   End
   Begin VB.Label Label39 
      BackStyle       =   0  'Transparent
      Caption         =   "Cost/Bag   :"
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
      Left            =   390
      TabIndex        =   53
      Top             =   6690
      Width           =   1545
   End
   Begin VB.Label Label38 
      BackStyle       =   0  'Transparent
      Caption         =   "L. Cost      :"
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
      Left            =   390
      TabIndex        =   51
      Top             =   6300
      Width           =   1530
   End
   Begin VB.Label Label37 
      BackStyle       =   0  'Transparent
      Caption         =   "Oth. Less  :"
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
      Left            =   390
      TabIndex        =   49
      Top             =   5940
      Width           =   1260
   End
   Begin VB.Label Label36 
      BackStyle       =   0  'Transparent
      Caption         =   "&Oth. Add    :"
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
      Left            =   390
      TabIndex        =   47
      Top             =   5565
      Width           =   1530
   End
   Begin VB.Label Label35 
      BackStyle       =   0  'Transparent
      Caption         =   "&God.Rent   :"
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
      Left            =   390
      TabIndex        =   45
      Top             =   5190
      Width           =   1530
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "&Vou. No   :"
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
      Left            =   5640
      TabIndex        =   2
      Top             =   1245
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
      Left            =   3180
      TabIndex        =   57
      Top             =   5880
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
      Left            =   6735
      TabIndex        =   33
      Top             =   5400
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
      Left            =   6735
      TabIndex        =   19
      Top             =   4395
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
      Left            =   6735
      TabIndex        =   25
      Top             =   5055
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
      Left            =   6735
      TabIndex        =   39
      Top             =   6735
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
      Left            =   3180
      TabIndex        =   27
      Top             =   5115
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
      ForeColor       =   &H0000FFFF&
      Height          =   495
      Left            =   8340
      TabIndex        =   94
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
      Left            =   6735
      TabIndex        =   37
      Top             =   6405
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
      TabIndex        =   92
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
      TabIndex        =   91
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
      TabIndex        =   90
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
      TabIndex        =   89
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
      TabIndex        =   88
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
      TabIndex        =   87
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
      TabIndex        =   86
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
      TabIndex        =   85
      Top             =   7590
      Width           =   3135
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
      Left            =   3180
      TabIndex        =   29
      Top             =   5460
      Width           =   1335
   End
   Begin VB.Label Label15 
      BackStyle       =   0  'Transparent
      Caption         =   "Due Days      :"
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
      Left            =   3180
      TabIndex        =   42
      Top             =   6225
      Width           =   1335
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
      Left            =   5640
      TabIndex        =   8
      Top             =   2340
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
      TabIndex        =   84
      Top             =   8160
      Width           =   1095
   End
   Begin VB.Label Label9 
      BackStyle       =   0  'Transparent
      Caption         =   "Purchase  A/C :"
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
      Left            =   420
      TabIndex        =   0
      Top             =   525
      Width           =   1620
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "&Bill No      :    "
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
      Left            =   5640
      TabIndex        =   5
      Top             =   1785
      Width           =   1095
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "&Broker  :"
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
      Left            =   420
      TabIndex        =   13
      Top             =   1935
      Width           =   870
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
      Left            =   420
      TabIndex        =   15
      Top             =   2340
      Width           =   1140
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
      Height          =   375
      Left            =   420
      TabIndex        =   11
      Top             =   1110
      Width           =   990
   End
   Begin VB.Label Label6 
      BackColor       =   &H00C00000&
      Caption         =   "    Trade Purchase"
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
      TabIndex        =   83
      Top             =   0
      Width           =   10815
   End
   Begin VB.Shape Shape1 
      FillColor       =   &H00BDDFD1&
      FillStyle       =   0  'Solid
      Height          =   6105
      Left            =   9540
      Shape           =   4  'Rounded Rectangle
      Top             =   525
      Width           =   1155
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
      Left            =   6420
      TabIndex        =   32
      Top             =   5460
      Width           =   255
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H00000080&
      FillColor       =   &H00EDFDFE&
      FillStyle       =   0  'Solid
      Height          =   1785
      Left            =   210
      Top             =   990
      Width           =   5145
   End
   Begin VB.Shape Shape5 
      BorderColor     =   &H00000080&
      FillColor       =   &H00EDFDFE&
      FillStyle       =   0  'Solid
      Height          =   540
      Left            =   210
      Top             =   420
      Width           =   9135
   End
   Begin VB.Shape Shape6 
      BorderColor     =   &H00000080&
      FillColor       =   &H00EDFDFE&
      FillStyle       =   0  'Solid
      Height          =   1785
      Left            =   5445
      Top             =   990
      Width           =   3900
   End
   Begin VB.Shape Shape7 
      BorderColor     =   &H00000080&
      FillColor       =   &H00EDFDFE&
      FillStyle       =   0  'Solid
      Height          =   2805
      Left            =   270
      Top             =   4305
      Width           =   2445
   End
   Begin VB.Shape Shape4 
      BorderColor     =   &H00000080&
      FillColor       =   &H00EDFDFE&
      FillStyle       =   0  'Solid
      Height          =   2805
      Left            =   2745
      Top             =   4305
      Width           =   6630
   End
End
Attribute VB_Name = "frmPurchaseTrade"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Const tblName As String = "tblAddLess"
Const IdField As String = "Vno"
Const OrderField As String = "Vdt,Vno"
Const MaxNo As Long = 43
Const OtherTxtFromNo As Long = 49 '-- Other Txt Ctrl From No
Const OtherTxtToNo As Long = 55 '-- Other Txt Ctrl Less No
Dim Rstbl As Recordset '-- TblAddLess
Dim rstblSub As Recordset '-- TblVouSub
Dim rstblBagSub As Recordset '-- TblBags
Dim varBookMark As Variant
Dim FirstTimeFlag As Boolean
Dim rsVou As Recordset '--TblVoucher
Dim rsOS As Recordset '--TblOutStanding
Dim OldgFrmTypeStr As String
Dim FormAction As Integer
Dim frmTypeStr As String
Dim i As Long

Private Sub Cbo_KeyPress(Index As Integer, KeyAscii As Integer)
If Index = 12 Then KeyAscii = 0 '--Tax
End Sub

Private Sub cbo_LostFocus(Index As Integer)
If Index = 12 And FormAction = vbDataActionAddNew Then '--Tax %
'    txt(13) = GProcGetColumnValue("tblMastNarration", "Narration", Cbo(12), "S", "MastTaxRate", "N")
'    txt(13) = Format(txt(13), FStr)
   Dim wMillCode As String
   Dim VATRtFromItem As Double
   wMillCode = GProcGetColumnValue("tblMastAccount", "AcName", txtTemp(5), "S", "AcCode", "N")
   VATRtFromItem = GProcGetColumnValue("tblMastItem", "ItName", msGrid.TextMatrix(1, 1), "S", "ITTMP1", "N", "ItMillcode=" & wMillCode) '--Item Code
   If VATRtFromItem > 0 Then
      txt(13) = VATRtFromItem
   Else
       txt(13) = GProcGetColumnValue("tblMastNarration", "Narration", Cbo(12), "S", "MastTaxRate", "N")
   End If
    txt(13) = Format(txt(13), FStr)
End If
End Sub
'-- Cartoon Bag Detail Entry
Private Sub CmdBagDetails_Click()
gfrmBagInwVNo = txt(0)
gfrmBagInwVType = frmTypeStr
frmBagEntry.Show
End Sub

Private Sub cmdBtn_Click(Index As Integer)
On Error GoTo ErrorRoutine
Dim i As Long
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
'    gClsSearch.SearchMultiField "tblAddLess,TblMastAccount", "Vno,AcName,Vdt,AdBillAmt,adInvno", Array("Vno", "Party", "Vdt", "Amount", "Bill No"), Array(txt(0).Width, 4000, 1080, 1500, 1000), " tblAddLess.AdAcCrCode=TblMastAccount.AcCode and vtype='" & frmTypeStr & "' and VYear=" & gCYear, "", "Vno,Vdt", 0, 0
    If gfrmHunkStr = "HNK" Then
       gClsSearch.SearchMultiField "tblAddLess,TblMastAccount", "Vno,AcName,Vdt,AdBillAmt,adInvno", Array("Vno", "Party", "Vdt", "Amount", "Bill No"), Array(txt(0).Width, 4000, 1080, 1500, 1000), " tblAddLess.AdAcCrCode=TblMastAccount.AcCode and vtype='" & frmTypeStr & "'  and left(adTmp5,3)= 'HNK'  and VYear=" & gCYear, "", "Vno,Vdt", 0, 0
    Else
       gClsSearch.SearchMultiField "tblAddLess,TblMastAccount", "Vno,AcName,Vdt,AdBillAmt,adInvno", Array("Vno", "Party", "Vdt", "Amount", "Bill No"), Array(txt(0).Width, 4000, 1080, 1500, 1000), " tblAddLess.AdAcCrCode=TblMastAccount.AcCode and vtype='" & frmTypeStr & "' and  (adtmp5 is null or adtmp5 = ' ')   and VYear=" & gCYear, "", "Vno,Vdt", 0, 0
    End If
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
    If gfrmHunkStr = "HNK" Then
       txtTemp(4) = GProcGetColumnValue("tblMastAccount", "AcCode", CStr(gAcCodePYHank), "N", "AcName", "S")
    Else
       txtTemp(4) = GProcGetColumnValue("tblMastAccount", "AcCode", CStr(gAcCodePY), "N", "AcName", "S")
    End If
'    txttemp(4) = GProcGetColumnValue("tblMastAccount", "AcCode", CStr(gAcCodePY), "N", "AcName", "S")
    FillCombo
    SetControlEd
    SetGrid
Case 7 'Modify
    FormAction = vbDataActionUpdate
    If gfrmHunkStr = "HNK" Then
      Hank = "HNK"
    Else
      Hank = ""
    End If
    'Call GProcChkPass(frmTypeStr, FormAction, txt(0), txt(2), txtTemp(3), gUserName, Hank)
    'If gChkPassTrueFalse = True Then
        'If mURecModi = True Then
        '    SearchHelp
        '    DoEvents
            txt(55) = "M"
            If IsRelatedRecord = False Then
                FormAction = vbDataActionUpdate
                Call GProcActivateControls(Me, True)
                Call GProcSetButtons(Me, vbDataActionUpdate)
                txt(0).Enabled = False
                mskDt(2).SetFocus
                SetControlEd
                '--- If Gate Pass Set Ctrl Enable
                If IsRelatedGatePass = True Then
                     Call GProcActivateControls(Me, False)
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
                     'txt(21).Enabled = True '-Tot amt
                     txt(18).Enabled = True 'add
                     txt(19).Enabled = True 'Less
                     txt(14).Enabled = True 'Tax amt
                     txt(17).Enabled = True '-Roff
                     SetControlEd
                End If
            End If
                     
    '--- Temp Chenge For 07/07/2008
                     Call GProcActivateControls(Me, True)
    '-----ebd
        'End If
    'Else
    '   FormAction = vbDataActionClose
    'End If
Case 8 'Delete
    FormAction = vbDataActionDelete
    If gfrmHunkStr = "HNK" Then
      Hank = "HNK"
    Else
      Hank = ""
    End If
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
                    '-- Delete Related Record
                    DeleteRelatedRecord
                    Rstbl.Delete
                    '--- Bag Details
                    gCn.Execute "delete From tblBags where InwNo=" & txt(0) & " and InwType='" & frmTypeStr & "'"
                    cmdBtn_Click (1)
                ElseIf Rstbl.RecordCount = 1 Then
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
         If gfrmHunkStr = "HNK" Then
           txt(43) = "HNK"
        Else
           txt(43) = ""
        End If
        GProcSaveRecord Me, Rstbl, FormAction, MaxNo, OtherTxtFromNo, OtherTxtToNo
        SaveRelatedRecord
        Call GProcActivateControls(Me, False)
        Call GProcSetButtons(Me, True)
        If FormAction = vbDataActionAddNew Then
            If cmdBtn(6).Visible = True Then cmdBtn(6).SetFocus
        Else
            If cmdBtn(7).Visible = True Then cmdBtn(7).SetFocus
        End If
        FormAction = vbDataActionClose
        msGrid.Col = 0
        msGrid.Row = 1
        txtGrid.Move msGrid.CellLeft + msGrid.Left, msGrid.CellTop + msGrid.Top, msGrid.CellWidth, msGrid.CellHeight
    End If
Case 14 'Cancel
    txtGrid.Visible = False
    DoEvents
    Call GProcClearForm(Me, Rstbl, MaxNo, True, OtherTxtFromNo, OtherTxtToNo)
    If Rstbl.RecordCount <= 0 Then
        Call GProcActivateControls(Me, False)
        Call GProcSetButtons(Me, vbDataActionCancel)
    Else
        If FormAction = vbDataActionAddNew Then
            Rstbl.CancelUpdate
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
    FormAction = vbDataActionCancel
    DoEvents
    msGrid.Col = 0
    If msGrid.Rows <= 1 Then msGrid.Rows = msGrid.Rows + 1
    msGrid.Row = 1
    txtGrid.Move msGrid.CellLeft + msGrid.Left, msGrid.CellTop + msGrid.Top, msGrid.CellWidth, msGrid.CellHeight
End Select
Exit Sub
ErrorRoutine:
    Rstbl.CancelUpdate
    GProcErrorHandler
End Sub
Private Sub SetControlEd()
txt(10).Enabled = False '--Sub Amt
'txt(14).Enabled = False '--TaxAmt
'txt(17).Enabled = False '--Roff
txt(21).Enabled = False
txt(11).Enabled = False '-- Amt
txt(24).Enabled = False '-- Ass Val
mskOS(10).Enabled = False '--Due Dt
txt(29).Enabled = False '-- LCost
txt(30).Enabled = False '-- Lcost/Bag
'txt(20).Enabled = False '-- JV No
txtTemp(4).Enabled = False '--Purchase A/c
'txtExpensesAmt.Enabled = False '-- Expenses
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
'Dim JVAcCode As Long
Dim X As String
Dim BillDt
Dim DueDt
'--- Account Code
MillCode = GProcGetColumnValue("tblMastAccount", "AcName", txtTemp(5), "S", "AcCode", "N")
CrCode = GProcGetColumnValue("tblMastAccount", "AcName", txtTemp(3), "S", "AcCode", "N")
DrCode = GProcGetColumnValue("tblMastAccount", "AcName", txtTemp(4), "S", "AcCode", "N")
BrkCode = GProcGetColumnValue("tblMastAccount", "AcName", txttempOS(6), "S", "AcCode", "N")
'TaxAcCode = GProcGetColumnValue("tblMastNarration", "Narration", txttemp(12), "S", "MastTaxPurAcCode", "N")
TaxAcCode = GProcGetColumnValue("tblMastNarration", "Narration", Cbo(12), "S", "MastTaxPurAcCode", "N")
'JVAcCode = gExpensesAcCode
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
        rsVouSub!Vtype = txt(1) & vbNullString '--Vtype(1)
        rsVouSub!Vdt = CDate(txt(2)) '--Vdt(2)
        rsVouSub!VSubCtrNo = i '--Sr No(3)
        '--Itcode(4)
        rsVouSub!vsubITCODE = GProcGetColumnValue("tblMastItem", "ItName", .TextMatrix(i, 1), "S", "ItCode", "N", " ItMillCode=" & MillCode) '--Item Code
        rsVouSub!vsubBAG = CLng(.TextMatrix(i, 4)) '--Bag(5)
        rsVouSub!vsubwt = CDbl(.TextMatrix(i, 5)) '-- Wt(6)
        rsVouSub!vsubRt = CDbl(.TextMatrix(i, 6)) '--Rate(7)
        rsVouSub!vsubRTPER = CDbl(.TextMatrix(i, 7)) '--Rate Per(8)
        rsVouSub!vsubAMT = CDbl(.TextMatrix(i, 8)) '--Amt(9)
        rsVouSub!vsubLOTNO = .TextMatrix(i, 3) '--Lot No(10)
        '--Godown Code(17)
        rsVouSub!VSubGodownCode = GProcGetColumnValue("tblMastNarration", "Narration", .TextMatrix(i, 2), "S", "NarrCode", "N", "")
        '--- It Ctrl No(11)
        If CLng(.TextMatrix(i, 9)) = 0 Then 'FormAction = vbDataActionAddNew Then
            rsVouSub!VSubItCtrlNo = GProcGenerateId("tblVouSub", "VSubItCtrlNo", "Vyear=" & gCYear) '--It Ctrl No
        Else
           rsVouSub!VSubItCtrlNo = CLng(.TextMatrix(i, 9))
        End If
        rsVouSub!VSubIssBag = CLng(.TextMatrix(i, 10)) '--Issue Bag(13)
        rsVouSub!VSubIsswt = CDbl(.TextMatrix(i, 11)) '-- Issue Wt(14)
        rsVouSub!VSubTmpBag = CLng(.TextMatrix(i, 12)) '--Temp Bag(15)
        rsVouSub!VSubTmpWt = CDbl(.TextMatrix(i, 13)) '-- Temp Wt(16)
        rsVouSub!VYear = txt(23) '--VYear(18)
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
'----------- JV
'gCn.Execute "delete from tblVoucher where Vtype='JV' and Vyear=" & gCYear & " and Vno=" & txt(20)
'If txtExpensesAmt > 0 Then
'    If (txt(20) = 0 Or FormAction = vbDataActionAddNew) Then txt(20) = GProcGenerateId("TblVoucher", "Vno", " Vtype='JV' and Vyear=" & gCYear)
'    GProcCreateVoucher FormAction, txt(20), 1, txt(2), "JV", CrCode, gExpensesAcCode, CDbl(txtExpensesAmt), txtVou(11), BillDt, "", "", "", "", "", True
'    GProcCreateVoucher FormAction, txt(20), 2, txt(2), "JV", gExpensesAcCode, CrCode, -CDbl(txtExpensesAmt), txtVou(11), BillDt, "", "", "", "", "", True
'End If
'--
'======================Outstanding Creation
If IsDate(mskOS(10)) Then
    DueDt = mskOS(10)
Else
    DueDt = ""
End If
GProcCreateOutStanding txt(0), CDate(txt(2)), txt(1), txtVou(11), CrCode, BrkCode, CDbl(txt(11)), CLng(txtOS(9)), DueDt
'GProcCreateOutStanding CLng(txt(0)), CDate(txt(2)), CStr(txt(1)), 12, 3, BrkCode, CDbl(txt(11)), CLng(txtOS(9)), DueDt
End Sub
Private Sub cmdLedger_Click(Index As Integer)
On Error GoTo ErrorHandler
If Index = 0 Then
    frmRptAccount.txt(1) = txtTemp(4)
Else
    frmRptAccount.txt(1) = txtTemp(3)
End If
frmMain.mnuRptRSubAccLeg_Click (0)
Exit Sub
ErrorHandler:
    GProcErrorHandler
End Sub
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case 13:
            If Me.ActiveControl.Name <> "txtGrid" Then SendKeys "{TAB}"
            'SendKeys "{TAB}"
'        Case 27:
'            If (cmdBtn(13).Enabled = False) Then
'                Unload Me
'            End If
    End Select
End Sub
Private Sub Form_Load()
    LblEmail.Caption = ""
    lblCell.Caption = ""
    LblCst.Caption = ""
    LblPan.Caption = ""
    Dim X As String
    Me.Top = 300
    Me.Left = Screen.Width / 2 - Me.Width / 2
    '--- TblAddLess
    frmTypeStr = gfrmTypeStr
    If gfrmHunkStr = "HNK" Then
       X = "Select * from " & tblName & " where Vtype='" & frmTypeStr & "'  and left(adTmp5,3)= 'HNK'   and VYear=" & gCYear
    Else
       X = "Select * from " & tblName & " where Vtype='" & frmTypeStr & "'  and  (adtmp5 is null or adtmp5 = ' ')   and VYear=" & gCYear
    End If
'    X = "Select * from " & tblName & " where Vtype='" & frmTypeStr & "' and VYear=" & gCYear
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
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
If FormAction = vbDataActionUpdate Then
    MsgBox "You are about to modify the record. Save the record and then exit.", vbInformation + vbOKOnly, Me.Caption
    Cancel = True
    Exit Sub
End If
'mShowRec = Space(5)
If Rstbl.State = adOpenStatic Then
    Rstbl.Requery
    Rstbl.Close
End If
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
mskVou(12) = mskDt(2)
If Index = 2 Then ShowAccountBalance
End Sub
Private Sub txt_GotFocus(Index As Integer)
cmdBtn(10).Cancel = True
GProcSelectBox Me.ActiveControl
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
Private Sub txttmpAmt_GotFocus()
GProcSelectBox Me.ActiveControl
End Sub
Private Sub txttmpAmt_KeyPress(KeyAscii As Integer)
KeyAscii = GProcNumberWithDecimal(KeyAscii, txttmpAmt)
End Sub
Private Sub txttmpAmt_LostFocus()
CalAmount
End Sub
Private Sub txtVou_GotFocus(Index As Integer)
GProcSelectBox Me.ActiveControl
End Sub
Private Sub txtVou_KeyPress(Index As Integer, KeyAscii As Integer)
If KeyAscii = 13 Then Exit Sub
KeyAscii = GProcValidateKey(rsVou, Index, KeyAscii, txtVou(Index).Text)
End Sub
Private Sub txtVou_LostFocus(Index As Integer)
SaveBtnEd
End Sub
Private Sub txt_LostFocus(Index As Integer)
CheckForNumber
SaveBtnEd
'If FormAction = vbDataActionAddNew And Index >= 10 And Index < 21 Then txt(21) = txt(11)
CalAmount
End Sub
Private Sub txt_KeyPress(Index As Integer, KeyAscii As Integer)
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
        Case 4 '-- Purchase A/c
            GName = GProcGetColumnValue("TblMastGroup", "AgCode", 90025, "N", "AgName", "S")
            GProcShowForm frmMastAccount, frmMain.mnuMstfrm(0), Array(frmMastAccount.txtTemp(2)), Array(GName)
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
TopPos = txtTemp(Index).Top + Me.Top + 650
LeftPos = txtTemp(Index).Left + Me.Left
If KeyAscii = 13 Then Exit Sub
Select Case Index
    Case 3  '-- Party   Alias Account
        gClsSearch.SearchMultiField "QryMastaccount", "Account,AcName,AcCity,AgName", Array("Party", "AcName", "City", "Group"), Array(3000, 0, 2000, 3000), "(Agcode=90015 or GpCode=90015 or Agcode=90017 or GpCode=90017 or ((Agcode=90029 or GpCode=90029) and (AcMillType=0 or AcMillType=5) ) )", Chr(KeyAscii), "acName", LeftPos, TopPos, True
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
    Case 4 '-- Purchase A/c  Alias Account
        gClsSearch.SearchMultiField "QryMastaccount", "Account,AcName,AcCity,AgName", Array("Purchase Account", "AcName", "City", "Group"), Array(3000, 0, 2000, 3000), "(AgCode=90025 or GpCode=90025)", Chr(KeyAscii), "AcName", LeftPos, TopPos, True
        KeyAscii = 0
        If gClsSearch.SearchMultiRetCol(0) <> "" And gClsSearch.SearchMultiRetCol(1) = "" Then
            txtTemp(Index) = gClsSearch.SearchMultiRetCol(0)
        Else
            If gClsSearch.SearchMultiRetCol(1) <> "" Then txtTemp(Index) = gClsSearch.SearchMultiRetCol(1)
        End If
    Case 5  '-- Mill  Alias Account
        gClsSearch.SearchMultiField "QryMastaccount", "Account,AcName,AcCity,AgName", Array("Mill", "AcName", "City", "Group"), Array(3000, 0, 2000, 3000), "((AgCode=90029 or GpCode=90029) and (AcMillType=0 or AcMillType=5))", Chr(KeyAscii), "AcName", LeftPos, TopPos, True
        KeyAscii = 0
        If gClsSearch.SearchMultiRetCol(0) <> "" And gClsSearch.SearchMultiRetCol(1) = "" Then
            txtTemp(Index) = gClsSearch.SearchMultiRetCol(0)
        Else
            If gClsSearch.SearchMultiRetCol(1) <> "" Then txtTemp(Index) = gClsSearch.SearchMultiRetCol(1)
        End If
        MillCtrlEd
    Case 12 '-- Tax Type
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
   ' If FormAction = vbDataActionAddNew Or FormAction = vbDataActionUpdate Then
        txt(15).Enabled = False
        txt(16).Enabled = False
        LblExmptAmt.Visible = True
        LblExmptKg.Visible = True
        txt(15).Visible = True
        txt(16).Visible = True
    'End If
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
DoEvents
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
Private Sub SetGrid()
Dim X As String
'--- Set Grid
Set rstblSub = New Recordset
X = "select VSubCtrNo,ItName,Narration,vSubLotNo,VSubBag,VSubWt" _
& " ,VSubRt,VSubRtPer,VSubAmt,VSubItCtrlNo,VSubIssBag,VSubIssWt,VSubTmpBag,VSubTmpWt from tblvouSub,TblMastNarration,TblMastItem where " _
& " tblvouSub.VsubItCode=TblMastItem.ItCode and tblvouSub.VSubGodownCode=TblMastNarration.NarrCode " _
& " and tblvouSub.vtype='" & frmTypeStr & "' and tblvouSub.vno=" & txt(0) & " and tblvouSub.VYear=" & gCYear & " order by VSubCtrNo"
GProcRstOpen rstblSub, X, "R"
With msGrid
    .FormatString = ">SR|<Count                      |<Godown        |<Lot No  |>Bag  |>Weight      |>Rate          |>Rate Per|>Amount        |||||"
    .ColWidth(9) = 0
    .ColWidth(10) = 0
    .ColWidth(11) = 0
    .ColWidth(12) = 0
    .ColWidth(13) = 0
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
                 .TextMatrix(.Row, i) = Null
            Case Else
                 .TextMatrix(.Row, i) = ""
        End Select
    Next i
ReNumberGridSrno
.Refresh
End With
End Sub
Private Sub msGrid_Click()
On Error Resume Next
    msGrid_EnterCell
End Sub
Private Sub msGrid_GotFocus()
cmdBtn(10).Cancel = False
End Sub
Private Sub MSGrid_LostFocus()
CalAmount
End Sub
Private Sub msGrid_EnterCell()
If FormAction <> vbDataActionCancel Then
    CheckForNumberTxtGrid
    DoEvents
    txtGrid = ""
    DoEvents
    If msGrid.Text <> "" Then txtGrid.Text = msGrid.Text
    '--- Lock TextBox
    With msGrid
        Select Case .Col
            Case 0: '---- Sr No
                txtGrid.Visible = True
                txtGrid.Locked = True
                ReNumberGridSrno
                DoEvents
                .Col = .Col + 1
            'Case 8 '-- Amount
            '    txtGrid.Locked = True
            Case Else
                txtGrid.Locked = False
        End Select
    End With
    DoEvents
    msGrid_KeyPress (0)
    DoEvents
End If
End Sub
Private Sub msGrid_KeyPress(KeyAscii As Integer)
With msGrid
    If KeyAscii = 27 Then '----Esc Key
        KeyAscii = 0
        If (Val(.TextMatrix(.Row, 7)) = 0) And .Rows > 2 Then
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
        If .Col = 8 Then 'rstblSub.Fields.Count - 2 Then
            If .Row = (.Rows - 1) Then '-- Add New Row
                .Rows = .Rows + 1
                .Row = .Row + 1
                GridAddNew
            Else '--Next Row
                .Row = .Row + 1
            End If
            .Col = 0
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
            If msGrid.Col = 1 Or msGrid.Col = 2 Then '-- Count ,godown
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
If msGrid.Col >= 4 Then
    CalGridAmt
    CalAmount
End If
DoEvents
End Sub
Private Sub txtGrid_GotFocus()
    txtGrid.Alignment = ProcSetAlignment(rstblSub, msGrid.Col)
    DoEvents
    Call GProcSelectBox(txtGrid)
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
            Case 1 '-- Count
                MillCode = GProcGetColumnValue("TblMastAccount", "AcName", txtTemp(5), "S", "Accode", "N")
                gClsSearch.SearchMultiField "tblMastItem,tblMastaccount", "ItName,AcName,ItStdpack,ItStdrateper", Array("Count Name", "Mill", "Std. Pack", "Rate Per"), Array(txtGrid.Width, 2000, 1000, 1000), "tblMastItem.ITMillCode=tblMastaccount.acCode and ITMillCode=" & MillCode, Chr(KeyAscii), "ItName", LeftPos, TopPos, True
                txtGrid = ""
                KeyAscii = 0
                If gClsSearch.SearchMultiRetCol(0) <> "" Then '-- Item Wt ,Rate From Master
                    txtGrid = gClsSearch.SearchMultiRetCol(0)
                    If IsNumeric(.TextMatrix(.Row, 5)) = False Then .TextMatrix(.Row, 5) = 0
                    If IsNumeric(.TextMatrix(.Row, 7)) = False Then .TextMatrix(.Row, 7) = 0
                    'If .TextMatrix(.Row, 5) = 0 Then .TextMatrix(.Row, 5) = gClsSearch.SearchMultiRetCol(2)
                    If .TextMatrix(.Row, 7) = 0 Then .TextMatrix(.Row, 7) = gClsSearch.SearchMultiRetCol(3)
                    txtGrid_KeyPress 13
                End If
            Case 2  '-- Godown
                gClsSearch.SearchMultiField "tblMastNarration", "Narration", Array("Godown"), Array(txtGrid.Width), " NarrType='G' ", Chr(KeyAscii), "Narration", LeftPos, TopPos, True
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
    If KeyAscii = 13 Then CheckForNumberTxtGrid
    '--- Navigate Row ,Col
    If KeyAscii = 13 And .Col < 8 Then
        .Col = .Col + 1
    ElseIf KeyAscii = 13 And .Col = 8 Then
        If .Row = (.Rows - 1) Then
            .Rows = .Rows + 1
            .Row = .Row + 1
            GridAddNew
        Else
            .Row = .Row + 1
        End If
        .Col = 0
    End If
    '-- On ESC
    With msGrid
        If KeyAscii = 27 Then '----Esc Key
            KeyAscii = 0
            If (Val(.TextMatrix(.Row, 7)) = 0) And .Rows > 2 Then
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
    Case 1 '-- Item
        GProcShowForm frmMastItem, frmMain.mnuMstfrm(3), Array(frmMastItem.txtTemp(9)), Array(txtTemp(5))
        Exit Sub
    Case 2 '-- Godown
        OldgFrmTypeStr = gfrmTypeStr
        gfrmTypeStr = "G"
        GProcShowForm frmMastNarration, frmMain.mnuMstfrm(4), "G"
        gfrmTypeStr = OldgFrmTypeStr
End Select
End If
End Sub
Private Sub CalGridAmt()
Dim MillCode As Long
Dim StdPack As Double
CheckForNumberGrid
MillCode = GProcGetColumnValue("TblMastAccount", "AcName", txtTemp(5), "S", "Accode", "N")
StdPack = GProcGetColumnValue("TblMastItem", "ITName", msGrid.TextMatrix(msGrid.Row, 1), "S", "ItStdpack", "N", " ItMillCode=" & MillCode)
'-- Wt=Bag * Std Pack
If msGrid.TextMatrix(msGrid.Row, 5) = 0 Then
    msGrid.TextMatrix(msGrid.Row, 5) = msGrid.TextMatrix(msGrid.Row, 4) * StdPack
End If
'-- Amt
'If msGrid.TextMatrix(msGrid.Row, 8) = 0 Then
'    '-- Amt =Wt*Rate/Rate Per
'    If CDbl(msGrid.TextMatrix(msGrid.Row, 7)) > 0 Then
'        msGrid.TextMatrix(msGrid.Row, 8) = CDbl(msGrid.TextMatrix(msGrid.Row, 5)) * CDbl(msGrid.TextMatrix(msGrid.Row, 6)) / CDbl(msGrid.TextMatrix(msGrid.Row, 7))
'    Else
'        msGrid.TextMatrix(msGrid.Row, 8) = CDbl(msGrid.TextMatrix(msGrid.Row, 5)) * CDbl(msGrid.TextMatrix(msGrid.Row, 6))
'    End If
'    msGrid.TextMatrix(msGrid.Row, 8) = Format(msGrid.TextMatrix(msGrid.Row, 8), FStr)
'Else '--Rate =Amt*Rate Per/Wt
    If CDbl(msGrid.TextMatrix(msGrid.Row, 5)) > 0 Then
        msGrid.TextMatrix(msGrid.Row, 6) = CDbl(msGrid.TextMatrix(msGrid.Row, 8)) * CDbl(msGrid.TextMatrix(msGrid.Row, 7)) / CDbl(msGrid.TextMatrix(msGrid.Row, 5))
    Else
        msGrid.TextMatrix(msGrid.Row, 6) = 0
    End If
    msGrid.TextMatrix(msGrid.Row, 6) = Format(msGrid.TextMatrix(msGrid.Row, 6), LStr)
'End If
End Sub
Private Sub CheckForNumber()
GProcCheckForNumber Me, Rstbl, MaxNo, pOtherTxtFromNo:=OtherTxtFromNo, pOtherTxtToNo:=OtherTxtToNo
End Sub
Private Sub CheckForNumberGrid()
Dim RowIndex As Long
With rstblSub
For RowIndex = 1 To msGrid.Rows - 1
    For i = 1 To .Fields.Count - 1
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
                    msGrid.TextMatrix(RowIndex, i) = Null
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
        Case adDate '-- Date
            If IsDate(txtGrid) = False Then
                txtGrid = Null
            Else
                txtGrid = CDate(txtGrid)
            End If
            txtGrid.MaxLength = 10
        Case Else
            txtGrid.MaxLength = .Fields(msGrid.Col).DefinedSize
    End Select
End With
End Sub
Private Function ProcSetAlignment(pRstbl As Recordset, pColIndex As Long)
Select Case pRstbl.Fields(pColIndex).Type
    Case adInteger, adDouble, 2, 3, 131, adDecimal
        ProcSetAlignment = 1
    Case Else
        ProcSetAlignment = 0
End Select
End Function
'-- ReNumber Grid Sr No
Private Sub ReNumberGridSrno()
Dim i As Long
With msGrid
    For i = 1 To .Rows - 1
        .TextMatrix(i, 0) = i
    Next i
End With
End Sub
'-- Grid Required
Private Function ValidateGrid() As Boolean
Dim i As Long
Dim MillCode As Long
Dim RowIndex As Long
Dim oldFrmTypeStr As String
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
                Case 1 '-- Count
                    '-Rate=0 ,count,wt=0 then delete row
                    If RowIndex > 1 And Trim(.TextMatrix(RowIndex, i)) = "" And CDbl(.TextMatrix(RowIndex, 5)) = 0 And CDbl(.TextMatrix(RowIndex, 6)) = 0 Then
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
                        'MSGrid.SetFocus
                        'GProcShowForm frmMastItem, "", Array(frmMastItem.txt(1), frmMastItem.txttemp(9)), Array(.TextMatrix(RowIndex, i), txttemp(5))
                        MsgBox "Check Count.Mill and Count are Mismatch.", vbCritical + vbOKOnly, Me.Caption
                        .Col = i
                        msGrid.SetFocus
                        Exit Function
                    End If
                Case 2 '-- Godown
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
                  
                Case 3 '-- Lot No
'                    If Trim(.TextMatrix(RowIndex, i)) = "" Then
'                        MsgBox "Check Lot No.", vbCritical + vbOKOnly, Me.Caption
'                        .Col = i
'                        MSGrid.SetFocus
'                        Exit Function
'                    End If
                Case 4 '-- Bag
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
                Case 5 '-- Wt
                    If CDbl(.TextMatrix(RowIndex, i)) = 0 Then
                        MsgBox "Check Weight.", vbCritical + vbOKOnly, Me.Caption
                        .Col = i
                        msGrid.SetFocus
                        Exit Function
                    End If
'                    '-- If Wt < Issue Wt
'                    If CDbl(.TextMatrix(RowIndex, 5)) < CDbl(.TextMatrix(RowIndex, 11)) Then
'                        MsgBox "Check Wt.Weight Less than issued.", vbCritical + vbOKOnly, Me.Caption
'                        .Col = i
'                        msGrid.SetFocus
'                        Exit Function
'                    End If
                Case 6 '-- Rate
                    If CDbl(.TextMatrix(RowIndex, i)) = 0 Then
                        MsgBox "Check Rate.", vbCritical + vbOKOnly, Me.Caption
                        .Col = i
                        msGrid.SetFocus
                        Exit Function
                    End If
           End Select
        Next i
      Next RowIndex
End With
ValidateGrid = True
ReNumberGridSrno
End Function
Private Sub FillCombo()
Call GProcFillCombo(Cbo(12), "select Narration from tblMastNarration where NarrType ='T' order by Narration")
If gfrmHunkStr = "HNK" Then
   Cbo(12) = GProcGetColumnValue("tblMastNarration", "NarrCode", CStr(gTaxCodePYHunk), "N", "Narration", "S")
Else
   Cbo(12) = GProcGetColumnValue("tblMastNarration", "NarrCode", CStr(gTaxCodePY), "N", "Narration", "S")
End If
'Cbo(12) = GProcGetColumnValue("tblMastNarration", "NarrCode", CStr(gTaxCodePY), "N", "Narration", "S")
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
Dim ExpAcCode As Long
Dim rstAddCheck As Recordset
Dim rstModiCheck As Recordset
    MillCtrlEd
    '---Validate Vou Date = Bill Date
''    If mskDt(2) >= CDate("25/03/2011") Then
''        MsgBox "Demo Experied Please Contact on MBL - 9922997663", vbInformation + vbOKOnly, Me.Caption
''        txt(0).SetFocus
''        Exit Function
''    End If

    
    
    
    If mskVou(12) <> mskDt(2) Then
        MsgBox "Voucher Date Must Be Same As Bill Date", vbInformation + vbOKOnly, Me.Caption
        mskDt(2).SetFocus
        Exit Function
    End If

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
    '-- Entry Tax - Calculated Tax (on Ass.Value )  Diff >1
    If CDbl(txt(14)) > 0 Then
        If Abs(CDbl(txt(14)) - (CDbl(txt(24)) * CDbl(txt(13)) / 100)) > 1 Then
            MsgBox "Check Tax Amt.", vbInformation + vbOKOnly, Me.Caption
            Exit Function
        End If
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
'--- Bill No
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
 '--Total Amt
'If CDbl(txt(21)) = 0 Then txt(21) = txt(11)
'-- JV
'gCn.Execute "delete from tblVoucher where vtype='JV' and vno=" & txt(20) & " and " _
'& " VYear=" & gCYear
'If CDbl(txt(21)) = CDbl(txt(11)) Then
'    txt(20) = 0
'End If
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
                If ColIndex = 4 Then TotBag = TotBag + CLng(msGrid.TextMatrix(RowIndex, 4)) '-- Tot Bag
                If ColIndex = 5 Then TotWt = TotWt + CDbl(msGrid.TextMatrix(RowIndex, 5)) '-- Tot Wt
            Next ColIndex
            RowIndex = RowIndex + 1
            .MoveNext
            If Not .EOF Then msGrid.Rows = msGrid.Rows + 1
        Loop
    End If
End With
txttmpTotBag = TotBag '-- Tot Bag
txttmpTotWt = Format(TotWt, WtStr) '-- Tot Wt
'--- Fill Related Control
SetRelatedRS
If rsVou.EOF = True Or rsVou.BOF Then
Else
txtVou(11) = rsVou.Fields("vBillNo") & vbNullString   '--Bill No
If IsDate(rsVou.Fields("vBillDt")) Then '--Bill Date
    mskVou(12) = Format(rsVou.Fields("vBilldt"), "dd/mm/yyyy")
Else
    mskVou(12) = "__/__/____"
End If
End If
'--Broker
If rsOS.EOF = True Or rsOS.BOF = True Then
Else
If rsOS.Fields("OutBrokerCode") > 0 Then txttempOS(6) = GProcGetColumnValue("tblMastAccount", "AcCode", rsOS.Fields("OutBrokerCode"), "N", "AcName", "S")
txtOS(9) = rsOS.Fields("OutDueDays") '--Due Days
If IsDate(rsOS.Fields("OutDueDate")) = True Then '--Due Dt
    mskOS(10) = Format(rsOS.Fields("OutDueDate"), "dd/mm/yyyy")
Else
    mskOS(10) = "__/__/____"
End If
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
Dim ExmptAmt As Double
Dim ExmptKg As Double
Dim AssValue As Double
Dim RowIndex As Integer
Dim TotBag As Long
Dim TotWt As Double
Dim LCost As Double
Dim IsExemptMill As Long
Dim MillAcCode  As Long
CalGridAmt
'-- Is Exempt Mill
IsExemptMill = GProcGetColumnValue("TblMastAccount", "AcName", txtTemp(5), "S", "ACIsExemptMill", "N")
'-- Due Date
If IsNumeric(txtOS(9)) = False Then txtOS(9) = "0"
txtOS(9) = Format(txtOS(9), PStr)
If IsDate(mskDt(2)) Then mskOS(10) = DateAdd("d", txtOS(9), mskDt(2))
'---- Sub amt
CheckForNumber
CheckForNumberGrid
RowIndex = 1
With msGrid
    Do While RowIndex < .Rows
        '-- sub amt
        If Val(.TextMatrix(RowIndex, 8)) > 0 Then
            SubAmt = SubAmt + Val(.TextMatrix(RowIndex, 8))
        End If
        '--Tot Bag
        If Val(.TextMatrix(RowIndex, 4)) > 0 Then
            TotBag = TotBag + Val(.TextMatrix(RowIndex, 4))
        End If
         '--Tot Wt
        If Val(.TextMatrix(RowIndex, 5)) > 0 Then
            TotWt = TotWt + Val(.TextMatrix(RowIndex, 5))
        End If
If RowIndex = 1 Then
   Dim wMillCode As String
   Dim VATRtFromItem As Double
   wMillCode = GProcGetColumnValue("tblMastAccount", "AcName", txtTemp(5), "S", "AcCode", "N")
   VATRtFromItem = GProcGetColumnValue("tblMastItem", "ItName", msGrid.TextMatrix(1, 1), "S", "ITTMP1", "N", "ItMillcode=" & wMillCode) '--Item Code
   If VATRtFromItem > 0 Then
      txt(13) = VATRtFromItem
   Else
       txt(13) = GProcGetColumnValue("tblMastNarration", "Narration", Cbo(12), "S", "MastTaxRate", "N")
   End If
    txt(13) = Format(txt(13), FStr)
End If
        
        RowIndex = RowIndex + 1
    Loop
End With
txttmpTotBag = TotBag
txttmpTotWt = Format(TotWt, WtStr)
SubAmt = Format(SubAmt, FStr)
'-- Exmpt Amt For Exmpt Mill
If IsExemptMill = 1 Then
    TaxAmt = CDbl(txt(14))
    '-- Ass Value
    If CDbl(txt(13)) > 0 Then
        AssValue = TaxAmt * 100 / CDbl(txt(13))
    Else
        AssValue = 0
    End If
    AssValue = Format(AssValue, FStr)
    '-- Exmpt amt=Sub +add-Less-Ass Value
    ExmptAmt = SubAmt + CDbl(txt(18)) - CDbl(txt(19)) - AssValue
    If ExmptAmt < 0 Then ExmptAmt = 0
    '--Exmpt/Kg
    If TotWt > 0 Then ExmptKg = ExmptAmt / TotWt
    ExmptKg = Format(ExmptKg, LStr)
Else
    '--- Subamt+Add -Less-Exempt.Amt
    AssValue = SubAmt + CDbl(txt(18)) - CDbl(txt(19))
    If FormAction = vbDataActionAddNew Then
        If txt(14) = 0 Then
            TaxAmt = AssValue * CDbl(txt(13)) / 100
        Else
            TaxAmt = CDbl(txt(14))
        End If
    Else
'--- Temp Change 07/04/2008
        If txt(14) = 0 Then
            TaxAmt = AssValue * CDbl(txt(13)) / 100
        Else
            TaxAmt = CDbl(txt(14))
        End If
'------
'       TaxAmt = CDbl(txt(14))
    End If
    TaxAmt = Format(TaxAmt, FStr)
End If
'--Roff
ROff = txt(17) '= 'CDbl(GProcMakeRounding(Amt) - Amt)
'-- SubAmt+Tax Amt +Add -Less+roff + Aft Tax Add - Aft Tax Less
Amt = SubAmt + TaxAmt + CDbl(txt(18)) - CDbl(txt(19)) + ROff + CDbl(txt(49)) - CDbl(txt(50))
'--G.Rent
If (Val(txt(26)) = 0) And (TotBag > 0) And IsDate(mskDt(2)) Then
    If CDate(mskDt(2)) < CDate("05/10/2006") Then '-Date <  CDate("06/10/2006")  Rs 5
        txt(26) = Format(TotBag * 5, FStr)
    Else '- Rs 6
        txt(26) = Format(TotBag * 6, FStr)
    End If
End If
'--Other Add
MillAcCode = GProcGetColumnValue("tblmastAccount", "AcName", txtTemp(5), "S", "AcCode", "N")
If (Val(txt(27)) = 0) And (TotBag > 0) Then
    If Left(GProcGetColumnValue("tblmastitem", "itname", msGrid.TextMatrix(1, 1), "S", "ItbrokOn", "S", " ItMillCode=" & MillAcCode), 1) = "B" Then
       txt(27) = Format(GProcGetColumnValue("tblmastitem", "itname", msGrid.TextMatrix(1, 1), "S", "Itbrokrt", "N", " ItMillCode=" & MillAcCode) * TotBag, FStr)
    Else
       txt(27) = Format(GProcGetColumnValue("tblmastitem", "itname", msGrid.TextMatrix(1, 1), "S", "Itbrokrt", "N", " ItMillCode=" & MillAcCode) * Val(txt(11)) / 100, FStr)
    End If
End If

'--- Landing Cost =Bill+Freight+Rent + L Cost add- L Cost Less
LCost = Amt + CDbl(txt(25)) + CDbl(txt(26)) + CDbl(txt(27)) - CDbl(txt(28))
txt(10) = SubAmt
txt(14) = TaxAmt
txt(24) = AssValue
'txt(17) = Roff
txt(11) = Amt
txt(29) = LCost
If Abs(SubAmt + CDbl(txt(18)) - CDbl(txt(19)) - AssValue) < 2 Then
    ExmptAmt = 0
    ExmptKg = 0
End If
txt(15) = ExmptAmt
txt(16) = ExmptKg
'-- Cost/Bag
txt(30) = 0
txt(31) = CLng(TotBag) '- Total Bag
If TotBag > 0 Then txt(30) = txt(29) / TotBag
'--Total & Sub Diff. Amt   (Total - Sub) amt
If IsNumeric(txttmpAmt) = False Then txttmpAmt = 0
txttmpSubTotDiff = Format(CDbl(txttmpAmt) - CDbl(SubAmt), FStr)
'============= Expenses JV
'txtExpensesAmt = 0
'If CDbl(txt(21)) < Amt Then txt(21) = Amt '--Total Amt
'txtExpensesAmt = CDbl(txt(21)) - Amt '-- Expenses
'If txtExpensesAmt > 0 And txt(20) = 0 Then
'    txt(20) = GProcGenerateId("TblVoucher", "Vno", " Vtype='JV' and Vyear=" & gCYear)
'End If
'txtExpensesAmt = Format(txtExpensesAmt, FStr)
CheckForNumber
End Sub
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
'--- Is Related Record
Private Function IsRelatedRecord() As Boolean
Dim RsR As Recordset
Dim RsG As Recordset
Dim X As String
'-- Audited
If GProcISAudited(txt(0), txt(1), txt(23)) = True Then
    MsgBox "Record is Audited.You can not Modify/Delete this Record.", vbCritical + vbOKOnly, App.Title
    IsRelatedRecord = True
    Exit Function
End If
'-- Purchase Return
X = "Select Vtype,Vdt,Vno from tblVouSub where Vtype in ('VY','VI') and VSubPItCtrlNo in" _
& " (select VSubItCtrlNo from tblVouSub where vtype='" & frmTypeStr & "' and vno=" & txt(0) & " and " _
& " VYear=" & gCYear & ")"
GProcRstOpen RsR, X, "R"
If RsR.EOF And RsR.BOF Then
Else
    MsgBox "Related Purchase Return Record " & RsR.Fields("Vno") & " is exists.You can not Modify/Delete this Record.", vbCritical + vbOKOnly, App.Title
    IsRelatedRecord = True
    Exit Function
End If
'-- Is Gate Pass
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
RsR.Close
End Function
