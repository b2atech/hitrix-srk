VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form frmPurchaseTradeGST 
   BackColor       =   &H00C2E3FC&
   BorderStyle     =   0  'None
   Caption         =   "Purchase Entry"
   ClientHeight    =   9420
   ClientLeft      =   0
   ClientTop       =   15
   ClientWidth     =   15570
   FillColor       =   &H000040C0&
   FillStyle       =   0  'Solid
   Icon            =   "frmPurchaseTradeGST.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   9420
   ScaleWidth      =   15570
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox txtVsub 
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
      Index           =   30
      Left            =   1305
      MaxLength       =   25
      TabIndex        =   34
      Text            =   "txtVsub(30)"
      Top             =   7335
      Width           =   2580
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
      Height          =   375
      Index           =   39
      Left            =   6435
      TabIndex        =   35
      Text            =   "39"
      Top             =   9090
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
      Height          =   375
      Index           =   38
      Left            =   15615
      TabIndex        =   133
      Text            =   "38"
      Top             =   7290
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
      Index           =   37
      Left            =   15840
      TabIndex        =   132
      Text            =   "37"
      Top             =   8370
      Visible         =   0   'False
      Width           =   555
   End
   Begin VB.TextBox txttemp 
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
      Index           =   28
      Left            =   1440
      TabIndex        =   38
      Text            =   "txttemp(28)"
      ToolTipText     =   "Press F3 for Creation"
      Top             =   8730
      Width           =   4125
   End
   Begin VB.TextBox txt 
      Alignment       =   1  'Right Justify
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
      Index           =   58
      Left            =   11655
      TabIndex        =   25
      Text            =   "58"
      Top             =   7785
      Width           =   1860
   End
   Begin VB.TextBox txt 
      Alignment       =   1  'Right Justify
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
      Index           =   57
      Left            =   10620
      TabIndex        =   24
      Text            =   "57"
      Top             =   7785
      Width           =   870
   End
   Begin VB.TextBox txt 
      Alignment       =   1  'Right Justify
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
      Index           =   56
      Left            =   9855
      TabIndex        =   23
      Text            =   "56"
      Top             =   7245
      Width           =   1680
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
      Left            =   4500
      TabIndex        =   123
      Text            =   "55"
      Top             =   9585
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
      Index           =   54
      Left            =   3645
      TabIndex        =   122
      Text            =   "54"
      Top             =   9720
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
      Index           =   53
      Left            =   2970
      TabIndex        =   121
      Text            =   "53"
      Top             =   9450
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
      Index           =   52
      Left            =   1935
      TabIndex        =   120
      Text            =   "52"
      Top             =   9675
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
      Index           =   51
      Left            =   540
      TabIndex        =   119
      Text            =   "51"
      Top             =   9765
      Visible         =   0   'False
      Width           =   555
   End
   Begin VB.TextBox txttemp 
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
      Index           =   35
      Left            =   9900
      TabIndex        =   9
      Text            =   "txttemp(35)"
      ToolTipText     =   "Press F3 for Creation"
      Top             =   2115
      Width           =   3420
   End
   Begin VB.TextBox txtVsub 
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
      Index           =   19
      Left            =   1305
      MaxLength       =   25
      TabIndex        =   30
      Text            =   "txtVsub(19)"
      Top             =   5310
      Width           =   2580
   End
   Begin VB.TextBox txtGodown 
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
      Left            =   1305
      TabIndex        =   33
      Text            =   "txtGodown"
      ToolTipText     =   "Press F3 for Creation"
      Top             =   6885
      Width           =   2970
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
      Index           =   48
      Left            =   9855
      TabIndex        =   111
      Text            =   "48"
      Top             =   9630
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
      Index           =   47
      Left            =   16605
      TabIndex        =   110
      Text            =   "47"
      Top             =   8235
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
      Index           =   46
      Left            =   16110
      TabIndex        =   109
      Text            =   "46"
      Top             =   8055
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
      Index           =   45
      Left            =   16290
      TabIndex        =   108
      Text            =   "45"
      Top             =   7425
      Visible         =   0   'False
      Width           =   555
   End
   Begin VB.TextBox txt 
      Alignment       =   1  'Right Justify
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
      Index           =   44
      Left            =   9885
      TabIndex        =   10
      Text            =   "44"
      Top             =   2565
      Width           =   1785
   End
   Begin VB.TextBox txtIGSTAmt 
      Alignment       =   1  'Right Justify
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
      Left            =   6300
      TabIndex        =   19
      Text            =   "txtIGSTAmt"
      Top             =   6480
      Width           =   2085
   End
   Begin VB.TextBox txtSGSTAmt 
      Alignment       =   1  'Right Justify
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
      Left            =   6300
      TabIndex        =   18
      Text            =   "txtSGSTAmt"
      Top             =   6030
      Width           =   2085
   End
   Begin VB.TextBox txtCGSTAmt 
      Alignment       =   1  'Right Justify
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
      Left            =   6300
      TabIndex        =   17
      Text            =   "txtCGSTAmt"
      Top             =   5580
      Width           =   2085
   End
   Begin VB.TextBox txt 
      Alignment       =   1  'Right Justify
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
      Index           =   49
      Left            =   11655
      TabIndex        =   21
      Text            =   "49"
      Top             =   6390
      Width           =   1860
   End
   Begin VB.TextBox txt 
      Alignment       =   1  'Right Justify
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
      Index           =   50
      Left            =   11655
      TabIndex        =   22
      Text            =   "50"
      Top             =   6795
      Width           =   1860
   End
   Begin VB.CommandButton CmdBagDetails 
      Caption         =   "Cartoon &Bag Details"
      Height          =   720
      Left            =   13845
      TabIndex        =   98
      Top             =   2085
      Width           =   870
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
      Left            =   17955
      TabIndex        =   42
      Text            =   "8"
      Top             =   3105
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
      Index           =   36
      Left            =   17370
      TabIndex        =   97
      Text            =   "36"
      Top             =   1260
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
      Left            =   17235
      TabIndex        =   96
      Text            =   "35"
      Top             =   765
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
      Left            =   17280
      TabIndex        =   95
      Text            =   "34"
      Top             =   315
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
      Left            =   15795
      TabIndex        =   94
      Text            =   "txttmpTotWt"
      Top             =   6030
      Visible         =   0   'False
      Width           =   555
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
      Left            =   15795
      TabIndex        =   93
      Text            =   "txttmpTotBag"
      Top             =   5625
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
      Height          =   375
      Index           =   15
      Left            =   16470
      TabIndex        =   92
      Text            =   "15"
      Top             =   4275
      Visible         =   0   'False
      Width           =   555
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
      Left            =   18000
      TabIndex        =   58
      Text            =   "txtExpensesAmt"
      Top             =   6210
      Visible         =   0   'False
      Width           =   270
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
      Left            =   17955
      TabIndex        =   91
      Text            =   "33"
      Top             =   7335
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
      Left            =   18135
      TabIndex        =   90
      Text            =   "32"
      Top             =   7830
      Visible         =   0   'False
      Width           =   420
   End
   Begin VB.TextBox txttempOS 
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
      Index           =   6
      Left            =   1245
      TabIndex        =   3
      Text            =   "txttempOS(6)"
      ToolTipText     =   "Press F3 for Creation"
      Top             =   2115
      Width           =   5445
   End
   Begin VB.TextBox txttemp 
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
      Index           =   5
      Left            =   1245
      TabIndex        =   4
      Text            =   "txttemp(5)"
      ToolTipText     =   "Press F3 for Creation"
      Top             =   2520
      Width           =   5445
   End
   Begin VB.TextBox txttemp 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Index           =   3
      Left            =   1245
      TabIndex        =   2
      Text            =   "txttemp(3)"
      ToolTipText     =   "Press F3 for Creation"
      Top             =   1185
      Width           =   5445
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
      Left            =   6795
      TabIndex        =   28
      Text            =   "txtCrBalance"
      Top             =   1200
      Width           =   1560
   End
   Begin VB.TextBox txtDrBalance 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   9270
      TabIndex        =   89
      Text            =   "txtDrBalance"
      Top             =   525
      Width           =   2100
   End
   Begin VB.CommandButton cmdLedger 
      Caption         =   "Ledger"
      Height          =   375
      Left            =   11475
      TabIndex        =   88
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
      Left            =   17865
      TabIndex        =   87
      Text            =   "31"
      Top             =   6840
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
      Left            =   17340
      TabIndex        =   56
      Text            =   "30"
      Top             =   3690
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
      Index           =   29
      Left            =   17340
      TabIndex        =   55
      Text            =   "29"
      Top             =   3285
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
      Index           =   28
      Left            =   17340
      TabIndex        =   54
      Text            =   "28"
      Top             =   2865
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
      Index           =   27
      Left            =   17340
      TabIndex        =   53
      Text            =   "27"
      Top             =   2460
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
      Index           =   26
      Left            =   17340
      TabIndex        =   52
      Text            =   "26"
      Top             =   2040
      Visible         =   0   'False
      Width           =   345
   End
   Begin VB.TextBox txt 
      Alignment       =   1  'Right Justify
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   25
      Left            =   1260
      TabIndex        =   32
      Text            =   "25"
      Top             =   6300
      Width           =   1545
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
      Left            =   16245
      TabIndex        =   46
      Text            =   "24"
      Top             =   6210
      Visible         =   0   'False
      Width           =   555
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
      TabIndex        =   15
      Top             =   3945
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
      Left            =   17895
      TabIndex        =   85
      Text            =   "1"
      Top             =   1080
      Visible         =   0   'False
      Width           =   420
   End
   Begin VB.TextBox txt 
      Alignment       =   1  'Right Justify
      Height          =   315
      Index           =   2
      Left            =   17895
      TabIndex        =   83
      Text            =   "2"
      Top             =   90
      Visible         =   0   'False
      Width           =   525
   End
   Begin VB.TextBox txt 
      Alignment       =   1  'Right Justify
      Height          =   315
      Index           =   5
      Left            =   17895
      TabIndex        =   82
      Text            =   "5"
      Top             =   1290
      Visible         =   0   'False
      Width           =   525
   End
   Begin VB.TextBox txt 
      Alignment       =   1  'Right Justify
      Height          =   315
      Index           =   3
      Left            =   17895
      TabIndex        =   81
      Text            =   "3"
      Top             =   450
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
      Left            =   1665
      Style           =   1  'Graphical
      TabIndex        =   73
      Top             =   9540
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
      Left            =   2520
      Style           =   1  'Graphical
      TabIndex        =   72
      ToolTipText     =   "To Add New Record"
      Top             =   9945
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
      Left            =   13845
      Style           =   1  'Graphical
      TabIndex        =   70
      ToolTipText     =   "Exit"
      Top             =   7830
      Width           =   870
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
      Left            =   13845
      Style           =   1  'Graphical
      TabIndex        =   63
      Top             =   7320
      Width           =   870
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
      Left            =   13845
      Style           =   1  'Graphical
      TabIndex        =   62
      ToolTipText     =   "To Save Record"
      Top             =   6825
      Width           =   870
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
      Left            =   14085
      Style           =   1  'Graphical
      TabIndex        =   71
      Top             =   8550
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
      Left            =   13845
      Style           =   1  'Graphical
      TabIndex        =   61
      ToolTipText     =   "To Delete Record"
      Top             =   6330
      Width           =   870
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
      Left            =   13845
      Style           =   1  'Graphical
      TabIndex        =   60
      ToolTipText     =   "To Modify Record"
      Top             =   5865
      Width           =   870
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
      Left            =   13845
      Style           =   1  'Graphical
      TabIndex        =   59
      ToolTipText     =   "To Add New Record"
      Top             =   5415
      Width           =   870
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
      Left            =   18360
      Style           =   1  'Graphical
      TabIndex        =   64
      Top             =   6615
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
      Left            =   13845
      Style           =   1  'Graphical
      TabIndex        =   65
      ToolTipText     =   "To Find Record"
      Top             =   1575
      Width           =   870
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
      Left            =   14340
      Style           =   1  'Graphical
      TabIndex        =   69
      ToolTipText     =   "Last"
      Top             =   1080
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
      Left            =   13845
      Style           =   1  'Graphical
      TabIndex        =   68
      ToolTipText     =   "First"
      Top             =   1080
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
      Left            =   14340
      Style           =   1  'Graphical
      TabIndex        =   67
      ToolTipText     =   "Next"
      Top             =   585
      Width           =   375
   End
   Begin VB.TextBox txt 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Index           =   0
      Left            =   9885
      TabIndex        =   5
      Text            =   "0"
      Top             =   1140
      Width           =   1785
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
      Left            =   13845
      Style           =   1  'Graphical
      TabIndex        =   66
      ToolTipText     =   "Previous"
      Top             =   585
      Width           =   375
   End
   Begin VB.TextBox txtVou 
      Alignment       =   1  'Right Justify
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Index           =   11
      Left            =   9885
      TabIndex        =   7
      Text            =   "txtVou(11)"
      Top             =   1620
      Width           =   1785
   End
   Begin VB.TextBox txt 
      Alignment       =   1  'Right Justify
      Height          =   315
      Index           =   4
      Left            =   17895
      TabIndex        =   80
      Text            =   "4"
      Top             =   810
      Visible         =   0   'False
      Width           =   525
   End
   Begin VB.TextBox txttemp 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   4
      Left            =   1725
      Locked          =   -1  'True
      TabIndex        =   1
      Text            =   "txttemp(4)"
      ToolTipText     =   "Press F3 for Creation"
      Top             =   525
      Width           =   6765
   End
   Begin VB.TextBox txt 
      Height          =   315
      Index           =   6
      Left            =   17895
      TabIndex        =   79
      Text            =   "6"
      Top             =   1650
      Visible         =   0   'False
      Width           =   435
   End
   Begin VB.TextBox txt 
      Alignment       =   1  'Right Justify
      Height          =   315
      Index           =   7
      Left            =   17895
      TabIndex        =   78
      Text            =   "7"
      Top             =   2010
      Visible         =   0   'False
      Width           =   525
   End
   Begin VB.TextBox txt 
      Alignment       =   1  'Right Justify
      Height          =   315
      Index           =   9
      Left            =   17895
      TabIndex        =   77
      Text            =   "9"
      Top             =   2370
      Visible         =   0   'False
      Width           =   525
   End
   Begin VB.TextBox txt 
      Alignment       =   1  'Right Justify
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
      Index           =   10
      Left            =   11655
      TabIndex        =   16
      Text            =   "10"
      Top             =   5445
      Width           =   1860
   End
   Begin VB.TextBox txt 
      Alignment       =   1  'Right Justify
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   11
      Left            =   10530
      TabIndex        =   29
      Text            =   "11"
      Top             =   8700
      Width           =   2985
   End
   Begin VB.TextBox txt 
      Alignment       =   1  'Right Justify
      Height          =   315
      Index           =   12
      Left            =   17895
      TabIndex        =   76
      Text            =   "12"
      Top             =   2730
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
      Left            =   16695
      TabIndex        =   47
      Text            =   "13"
      Top             =   3330
      Visible         =   0   'False
      Width           =   555
   End
   Begin VB.TextBox txt 
      Alignment       =   1  'Right Justify
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
      Index           =   14
      Left            =   11655
      TabIndex        =   20
      Text            =   "14"
      Top             =   5940
      Width           =   1860
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
      Left            =   16110
      TabIndex        =   50
      Text            =   "16"
      Top             =   4815
      Visible         =   0   'False
      Width           =   555
   End
   Begin VB.TextBox txt 
      Alignment       =   1  'Right Justify
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
      Index           =   17
      Left            =   11655
      TabIndex        =   27
      Text            =   "17"
      Top             =   8265
      Width           =   1860
   End
   Begin VB.TextBox txtOS 
      Alignment       =   1  'Right Justify
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   9
      Left            =   5760
      TabIndex        =   41
      Text            =   "txtOS(9)"
      Top             =   7290
      Width           =   1095
   End
   Begin VB.TextBox txt 
      Alignment       =   1  'Right Justify
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   18
      Left            =   1440
      TabIndex        =   36
      Text            =   "18"
      Top             =   8145
      Width           =   1095
   End
   Begin VB.TextBox txt 
      Alignment       =   1  'Right Justify
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   19
      Left            =   4005
      TabIndex        =   37
      Text            =   "19"
      Top             =   8100
      Width           =   1500
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
      Left            =   17895
      TabIndex        =   57
      Text            =   "20"
      Top             =   5790
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
      Index           =   21
      Left            =   17895
      TabIndex        =   49
      Text            =   "21"
      Top             =   5115
      Visible         =   0   'False
      Width           =   435
   End
   Begin VB.TextBox txt 
      Alignment       =   1  'Right Justify
      Height          =   315
      Index           =   22
      Left            =   17895
      TabIndex        =   75
      Text            =   "22"
      Top             =   3810
      Visible         =   0   'False
      Width           =   525
   End
   Begin VB.TextBox txt 
      Alignment       =   1  'Right Justify
      Height          =   315
      Index           =   23
      Left            =   17895
      TabIndex        =   74
      Text            =   "23"
      Top             =   4050
      Visible         =   0   'False
      Width           =   525
   End
   Begin MSMask.MaskEdBox mskDt 
      Height          =   375
      Index           =   2
      Left            =   11820
      TabIndex        =   6
      Top             =   1140
      Width           =   1515
      _ExtentX        =   2672
      _ExtentY        =   661
      _Version        =   393216
      MaxLength       =   10
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   12
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
      Left            =   11820
      TabIndex        =   8
      Top             =   1620
      Width           =   1515
      _ExtentX        =   2672
      _ExtentY        =   661
      _Version        =   393216
      MaxLength       =   10
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   12
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
      Left            =   6975
      TabIndex        =   43
      Top             =   7290
      Width           =   1425
      _ExtentX        =   2514
      _ExtentY        =   661
      _Version        =   393216
      MaxLength       =   10
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   12
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
      Bindings        =   "frmPurchaseTradeGST.frx":058A
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
      Height          =   1905
      Left            =   225
      TabIndex        =   12
      Top             =   3195
      Width           =   15135
      _ExtentX        =   26696
      _ExtentY        =   3360
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
      Left            =   17595
      TabIndex        =   44
      Top             =   -45
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
   Begin MSMask.MaskEdBox mskDt 
      Height          =   375
      Index           =   34
      Left            =   11820
      TabIndex        =   11
      Top             =   2595
      Width           =   1515
      _ExtentX        =   2672
      _ExtentY        =   661
      _Version        =   393216
      MaxLength       =   10
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   12
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
   Begin MSMask.MaskEdBox mskVouSub 
      Height          =   375
      Index           =   20
      Left            =   1305
      TabIndex        =   31
      Top             =   5805
      Width           =   1515
      _ExtentX        =   2672
      _ExtentY        =   661
      _Version        =   393216
      MaxLength       =   10
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   12
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
   Begin VB.Label Label32 
      BackStyle       =   0  'Transparent
      Caption         =   "Freight  :"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   375
      Left            =   300
      TabIndex        =   134
      Top             =   6360
      Width           =   990
   End
   Begin VB.Label Label31 
      BackStyle       =   0  'Transparent
      Caption         =   "Challan No: "
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   375
      Left            =   270
      TabIndex        =   131
      Top             =   7425
      Width           =   1050
   End
   Begin VB.Label Label29 
      BackStyle       =   0  'Transparent
      Caption         =   "TDS Rate   : "
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   375
      Left            =   270
      TabIndex        =   130
      Top             =   8145
      Width           =   1365
   End
   Begin VB.Label Label28 
      BackStyle       =   0  'Transparent
      Caption         =   "TDS Amount  : "
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   375
      Left            =   2580
      TabIndex        =   129
      Top             =   8145
      Width           =   1365
   End
   Begin VB.Label Label27 
      BackStyle       =   0  'Transparent
      Caption         =   "TDS A/c     : "
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   375
      Left            =   270
      TabIndex        =   128
      Top             =   8760
      Width           =   1365
   End
   Begin VB.Label lblTotPur 
      BackStyle       =   0  'Transparent
      Caption         =   "Amount Rs."
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800080&
      Height          =   375
      Left            =   6255
      TabIndex        =   127
      Top             =   8415
      Width           =   1995
   End
   Begin VB.Label Label26 
      BackStyle       =   0  'Transparent
      Caption         =   "Total Purchase Amount"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   375
      Left            =   5760
      TabIndex        =   126
      Top             =   7920
      Width           =   2865
   End
   Begin VB.Shape Shape10 
      BorderColor     =   &H00000080&
      FillColor       =   &H00EDFDFE&
      FillStyle       =   0  'Solid
      Height          =   1320
      Left            =   180
      Top             =   7920
      Width           =   5520
   End
   Begin VB.Label Label25 
      BackStyle       =   0  'Transparent
      Caption         =   "TCS @ :                      :"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   375
      Left            =   8895
      TabIndex        =   125
      Top             =   7740
      Width           =   945
   End
   Begin VB.Label Label24 
      BackStyle       =   0  'Transparent
      Caption         =   "TCS ON :                      :"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   375
      Left            =   8895
      TabIndex        =   124
      Top             =   7245
      Width           =   945
   End
   Begin VB.Label Label23 
      BackStyle       =   0  'Transparent
      Caption         =   "Transport  : "
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   375
      Left            =   8715
      TabIndex        =   118
      Top             =   2115
      Width           =   1050
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00000080&
      X1              =   180
      X2              =   4275
      Y1              =   6735
      Y2              =   6735
   End
   Begin VB.Label Label21 
      BackStyle       =   0  'Transparent
      Caption         =   "L.R. Dt   :"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   375
      Left            =   270
      TabIndex        =   117
      Top             =   5850
      Width           =   990
   End
   Begin VB.Label Label16 
      BackStyle       =   0  'Transparent
      Caption         =   "L.R. No   :"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   375
      Left            =   270
      TabIndex        =   116
      Top             =   5355
      Width           =   990
   End
   Begin VB.Label Label20 
      BackStyle       =   0  'Transparent
      Caption         =   "After Tax Less           :"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   375
      Left            =   8895
      TabIndex        =   115
      Top             =   6750
      Width           =   2235
   End
   Begin VB.Label Label19 
      BackStyle       =   0  'Transparent
      Caption         =   "After Tax Add            :"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   375
      Left            =   8895
      TabIndex        =   114
      Top             =   6345
      Width           =   2235
   End
   Begin VB.Label Label18 
      BackStyle       =   0  'Transparent
      Caption         =   "Total G S T Amount   :"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   375
      Left            =   8895
      TabIndex        =   113
      Top             =   5985
      Width           =   2235
   End
   Begin VB.Label Label17 
      BackStyle       =   0  'Transparent
      Caption         =   "Godown : "
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   375
      Left            =   270
      TabIndex        =   112
      Top             =   6885
      Width           =   915
   End
   Begin VB.Shape Shape3 
      FillColor       =   &H00BDDFD1&
      FillStyle       =   0  'Solid
      Height          =   2610
      Left            =   13725
      Shape           =   4  'Rounded Rectangle
      Top             =   450
      Width           =   1125
   End
   Begin VB.Label Label14 
      BackStyle       =   0  'Transparent
      Caption         =   "ERN. No   :"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   375
      Left            =   8715
      TabIndex        =   107
      Top             =   2610
      Width           =   1125
   End
   Begin VB.Label Label12 
      BackStyle       =   0  'Transparent
      Caption         =   "GSTIN  :"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   375
      Left            =   315
      TabIndex        =   106
      Top             =   1665
      Width           =   945
   End
   Begin VB.Label lblGSTINNo 
      BackStyle       =   0  'Transparent
      Caption         =   "GSTIN No "
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   375
      Left            =   1215
      TabIndex        =   105
      Top             =   1665
      Width           =   3780
   End
   Begin VB.Label Label22 
      BackStyle       =   0  'Transparent
      Caption         =   "State Code :"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   375
      Left            =   5580
      TabIndex        =   104
      Top             =   1665
      Width           =   1215
   End
   Begin VB.Label lblStateCd 
      BackStyle       =   0  'Transparent
      Caption         =   "State Cd"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   375
      Left            =   6795
      TabIndex        =   103
      Top             =   1665
      Width           =   990
   End
   Begin VB.Label Label11 
      BackStyle       =   0  'Transparent
      Caption         =   "R. Off                        :"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   375
      Left            =   8895
      TabIndex        =   102
      Top             =   8280
      Width           =   2160
   End
   Begin VB.Label Label13 
      BackStyle       =   0  'Transparent
      Caption         =   "C GST Amount   :"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   375
      Left            =   4590
      TabIndex        =   101
      Top             =   5580
      Width           =   1710
   End
   Begin VB.Label Label8 
      BackStyle       =   0  'Transparent
      Caption         =   "S GST Amount   :"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   375
      Left            =   4590
      TabIndex        =   100
      Top             =   6075
      Width           =   1605
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "I  GST  Amount   :"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   375
      Left            =   4590
      TabIndex        =   99
      Top             =   6585
      Width           =   1665
   End
   Begin VB.Shape Shape5 
      BorderColor     =   &H00C00000&
      BorderWidth     =   5
      FillColor       =   &H00C00000&
      Height          =   9420
      Left            =   60
      Top             =   -45
      Width           =   15495
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "&Party  :"
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
      Left            =   360
      TabIndex        =   26
      Top             =   1155
      Width           =   855
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "M&ill     :"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   375
      Left            =   360
      TabIndex        =   40
      Top             =   2535
      Width           =   1635
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "&Broker : "
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   375
      Left            =   360
      TabIndex        =   39
      Top             =   2115
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
      ForeColor       =   &H000000FF&
      Height          =   375
      Left            =   360
      TabIndex        =   0
      Top             =   585
      Width           =   1620
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
      ForeColor       =   &H00800000&
      Height          =   240
      Left            =   8715
      TabIndex        =   13
      Top             =   1140
      Width           =   990
   End
   Begin VB.Label Label10 
      BackStyle       =   0  'Transparent
      Caption         =   "Taxable Amount       :"
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
      Height          =   375
      Left            =   8895
      TabIndex        =   45
      Top             =   5490
      Width           =   2055
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
      ForeColor       =   &H000000FF&
      Height          =   375
      Left            =   8895
      TabIndex        =   48
      Top             =   8790
      Width           =   1095
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
      Left            =   12285
      TabIndex        =   86
      Top             =   0
      Width           =   2475
   End
   Begin VB.Label Label15 
      BackStyle       =   0  'Transparent
      Caption         =   "D&ue Days    :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   255
      Left            =   4500
      TabIndex        =   51
      Top             =   7335
      Width           =   1200
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
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   8715
      TabIndex        =   14
      Top             =   1710
      Width           =   1095
   End
   Begin VB.Label Label6 
      BackColor       =   &H00C00000&
      Caption         =   "   Trade  Purchase"
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
      TabIndex        =   84
      Top             =   0
      Width           =   15435
   End
   Begin VB.Shape Shape1 
      FillColor       =   &H00BDDFD1&
      FillStyle       =   0  'Solid
      Height          =   3150
      Left            =   13725
      Shape           =   4  'Rounded Rectangle
      Top             =   5265
      Width           =   1125
   End
   Begin VB.Shape Shape4 
      BorderColor     =   &H00000080&
      FillColor       =   &H00F5F5DE&
      FillStyle       =   0  'Solid
      Height          =   3990
      Left            =   8640
      Top             =   5265
      Width           =   4995
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H00000080&
      FillColor       =   &H00F5F5DE&
      FillStyle       =   0  'Solid
      Height          =   2025
      Left            =   195
      Top             =   1020
      Width           =   8295
   End
   Begin VB.Shape Shape6 
      BorderColor     =   &H00000080&
      FillColor       =   &H00EDFDFE&
      FillStyle       =   0  'Solid
      Height          =   510
      Left            =   180
      Top             =   450
      Width           =   13470
   End
   Begin VB.Shape Shape7 
      BorderColor     =   &H00000080&
      FillColor       =   &H00EDFDFE&
      FillStyle       =   0  'Solid
      Height          =   2025
      Left            =   8595
      Top             =   1020
      Width           =   5025
   End
   Begin VB.Shape Shape8 
      BorderColor     =   &H00000080&
      FillColor       =   &H00EDFDFE&
      FillStyle       =   0  'Solid
      Height          =   2670
      Left            =   210
      Top             =   5175
      Width           =   4140
   End
   Begin VB.Shape Shape9 
      BorderColor     =   &H00000080&
      FillColor       =   &H00F5F5DE&
      FillStyle       =   0  'Solid
      Height          =   1875
      Left            =   4455
      Top             =   5265
      Width           =   4050
   End
End
Attribute VB_Name = "frmPurchaseTradeGST"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Const tblName As String = "tblAddLess"
Const IdField As String = "Vno"
Const OrderField As String = "Vdt,Vno"
Const MaxNo As Long = 36
Const OtherTxtFromNo As Long = 44 '-- Other Txt Ctrl From No
Const OtherTxtToNo As Long = 58 '-- Other Txt Ctrl Less No
Dim rstbl As Recordset '-- TblAddLess
Dim rstblSub As Recordset '-- TblVouSub
Dim varBookMark As Variant
Dim FirstTimeFlag As Boolean
Dim rsVou As Recordset '--TblVoucher
Dim rsOS As Recordset '--TblOutStanding
Dim rsVouSubTmp As Recordset
Dim OldgFrmTypeStr As String
Dim FormAction As Integer
Dim frmTypeStr As String
Dim CriteriaStr As String
Dim i As Long

'-- Cartoon Bag Detail Entry
Private Sub CmdBagDetails_Click()
gfrmBagInwVNo = txt(0)
gfrmBagInwVType = frmTypeStr
frmBagEntry.Show
End Sub
'Dim GridFlag As Boolean
Private Sub cmdBtn_Click(Index As Integer)
'On Error GoTo ErrorRoutine
Select Case Index
Case 0 'Previous
    If GProcRstOpen(rstbl, "Select * from " & tblName & " where vno = (Select max(Vno) from " & tblName & " where Vno <> 0  and Vno < " & txt(0) & " and  " & CriteriaStr & "  ) and  " & CriteriaStr & " order by " & OrderField & " ", "O") > 0 Then
       ReadFields
    Else
      MsgBox "This is First Record ", vbOKOnly + vbInformation
    End If

'    rstbl.MovePrevious
'    If rstbl.BOF = True Then
'        rstbl.MoveFirst
'        MsgBox ("This is First Record"), vbInformation + vbOKOnly, "First Record"
'    End If
'    ReadFields
Case 1 'Next
    If GProcRstOpen(rstbl, "Select * from " & tblName & " where Vno = (Select min(Vno) from " & tblName & " where Vno <> 0  and Vno > " & txt(0) & "  and " & CriteriaStr & ") and  " & CriteriaStr & " order by " & OrderField & "  ", "O") > 0 Then
       ReadFields
    Else
       MsgBox "This is Last Record ", vbOKOnly + vbInformation
    End If
'    rstbl.MoveNext
'    If rstbl.EOF = True Then
'        rstbl.MoveLast
'        MsgBox ("This is Last Record"), vbInformation + vbOKOnly, "Last Record"
'    End If
'    ReadFields
Case 2 'First
    If GProcRstOpen(rstbl, "Select * from " & tblName & " where Vno = (Select min(Vno) from " & tblName & " where Vno <> 0 and " & CriteriaStr & ") and  " & CriteriaStr & " order by " & OrderField & "  ", "O") > 0 Then ReadFields
'    rstbl.MoveFirst
'    ReadFields
Case 3 'Last
    If GProcRstOpen(rstbl, "Select * from " & tblName & " where " & CriteriaStr & "  and  " & IdField & " = (Select max(" & IdField & ") from " & tblName & " where vdt = (SELECT MAX(vdt) from " & tblName & " where " & CriteriaStr & " )  and  " & IdField & " <> 0 and  " & CriteriaStr & " ) order by " & OrderField & " ", "R") > 0 Then ReadFields Else GProcClearForm Me, rstbl, MaxNo
'    rstbl.MoveLast
'    ReadFields
Case 4 'Find
    gClsSearch.SearchMultiField "tblAddLess,TblMastAccount", "Vno,AcName,Vdt,AdInvNo,AdGpDt", Array("Vno", "Party", "Vdt", "Bill No", "Bill Dt"), Array(txt(0).Width, 4000, 1080, 1000, 1200), " tblAddLess.AdAcCrCode=TblMastAccount.AcCode and vtype='" & frmTypeStr & "' and VYear=" & gCYear, "", "Vno,Vdt", 0, 0
    If gClsSearch.SearchMultiRetCol(0) <> "" Then
        If GProcRstOpen(rstbl, "Select * from " & tblName & " where Vno = (Select max(" & gClsSearch.SearchMultiRetCol(0) & ") from " & tblName & " where Vno <> 0 and " & CriteriaStr & " )and  " & CriteriaStr & " order by " & OrderField & "  ", "O") > 0 Then ReadFields
'        rstbl.MoveFirst
'        rstbl.Find "Vno = " & gClsSearch.SearchMultiRetCol(0)
        gClsSearch.SearchMultiRetCol(0) = ""
'        ReadFields
    End If
Case 6 'Add
    FormAction = vbDataActionAddNew
    Call GProcActivateControls(Me, True)
    SetControlEd
    Call GProcSetButtons(Me, vbDataActionAddNew)
    Call GProcClearForm(Me, rstbl, MaxNo, True, OtherTxtFromNo, OtherTxtToNo)
    If FirstTimeFlag = True Then
        FirstTimeFlag = False
    Else
        txttemp(3).SetFocus
    End If
    txt(0) = GProcGenerateId("TblAddLess", "Vno", "Vtype='" & frmTypeStr & "' and VYear=" & gCYear)
    mskDt(2) = GProcGenerateVDt("tblAddLess", "Vdt", "Vtype='" & frmTypeStr & "' and VYear=" & gCYear)
    '--A/C
'    If frmTypeStr <> "PY" Then
       txttemp(4) = GProcGetColumnValue("tblMastAccount", "AcCode", CStr(gAcCodePY), "N", "AcName", "S")
'       txttemp(4).SetFocus
'    End If
    SetGrid
Case 7 'Modify
    'If mURecModi = True Then
    '    SearchHelp
    '    DoEvents
        If IsRelatedRecord = False Then
            GProcRstOpen rstbl, "Select * from " & tblName & " where " & IdField & " = " & txt(0) & " and " & CriteriaStr, "R"
            ReadFields
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
                 txtOS(9).Enabled = True '-Due days
                 txttemp(35).Enabled = True
                 txttempOS(6).Enabled = True
                 txtVsub(19).Enabled = True
                 mskVouSub(20).Enabled = True
                 txt(44).Enabled = True
                 mskDt(34).Enabled = True
                 txt(17).Enabled = True '-Roff
                 txt(58).Enabled = True
                 SetControlEd
            End If
        End If
'        txt(18).Enabled = True
'        txt(19).Enabled = True
'        txttemp(28).Enabled = True
    'End If
Case 8 'Delete
    'If mURecDel = True Then
      If IsRelatedGatePass = True Then
            MsgBox "You can not Modify/Delete this record.Related Gate Pass is exists.", vbCritical + vbOKOnly, App.Title
            Exit Sub
      End If
      
      
      If IsRelatedRecord = False Then
        If MsgBox("Do You Really Delete This Record?", vbQuestion + vbYesNo + vbDefaultButton1, "Confirm Delete") = vbYes Then
           
           GProcRstOpen rstbl, "Select * from " & tblName & " where " & IdField & " = " & txt(0) & " and " & CriteriaStr, "O"
           ReadFields
            
            If rstbl.RecordCount > 1 Then
                '--Update Stock
                GProcUpdateVSubStock txt(0), gCYear, frmTypeStr, False, True
                 '-- Delete Related Record
                DeleteRelatedRecord
                rstbl.Delete
                 '--- Bag Details
                gCn.Execute "delete From tblBags where InwNo=" & txt(0) & " and InwType='" & frmTypeStr & "'"
                cmdBtn_Click (1)
            ElseIf rstbl.RecordCount = 1 Then
                '--Update Stock
                GProcUpdateVSubStock txt(0), gCYear, frmTypeStr, False, True
                '-- Delete Related Record
                DeleteRelatedRecord
                rstbl.Delete
                cmdBtn_Click (3)
                '--- Bag Details
                gCn.Execute "delete From tblBags where InwNo=" & txt(0) & " and InwType='" & frmTypeStr & "'"
                Call GProcActivateControls(Me, False)
                Call GProcSetButtons(Me, vbDataActionCancel)
                Form_Load
                cmdBtn_Click (3)
                cmdBtn_Click (14)
                
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
       GProcRstOpen rstbl, "Select * from " & tblName & " where " & IdField & " = " & txt(0) & " and " & CriteriaStr, "O"
        If Trim(txtGrid) <> "" Then
            msGrid.text = txtGrid
            DoEvents
            txtGrid = ""
        End If
        txtGrid.Visible = False
       ' If FormAction = vbDataActionAddNew Then txt(0) = GProcGenerateId(tblName, IdField, "vtype='" & FrmTypeStr & "' and VYear=" & gCYear)
        CalAmount
        FillTxtFromTemp
        '-- Save Record
        txt(7) = mskVou(12) '--gpDt
        txt(12) = GProcGetColumnValue("tblMastNarration", "Narration", "-", "S", "NarrCode", "N")
        GProcSaveRecord Me, rstbl, FormAction, MaxNo, OtherTxtFromNo, OtherTxtToNo
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
    cmdBtn_Click (3)
    txtGrid.Visible = False
    If rstbl.RecordCount <= 0 Then
        Call GProcActivateControls(Me, False)
        Call GProcSetButtons(Me, vbDataActionCancel)
    Else
        Call GProcActivateControls(Me, False)
        Call GProcSetButtons(Me, True)
    End If
    
    
    
'    If rstbl.RecordCount <= 0 Then
'        Call GProcClearForm(Me, rstbl, MaxNo, True, OtherTxtFromNo, OtherTxtToNo)
'        Call GProcActivateControls(Me, False)
'        Call GProcSetButtons(Me, vbDataActionCancel)
'    Else
'        If FormAction = vbDataActionAddNew Then
'            rstbl.MoveLast
'        Else
'            '--Update Stock
'            GProcUpdateVSubStock txt(0), gCYear, frmTypeStr
'            rstbl.CancelUpdate
'        End If
'        Call GProcClearForm(Me, rstbl, MaxNo, True, OtherTxtFromNo, OtherTxtToNo)
'        ReadFields
'        Call GProcActivateControls(Me, False)
'        Call GProcSetButtons(Me, True)
'    End If
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
    rstbl.CancelUpdate
    GProcErrorHandler
End Sub
Private Sub SetControlEd()
If frmTypeStr = "PY" Then
   txttemp(4).Enabled = False '----Purch A/c
Else
   txttemp(4).Enabled = True '----Purch A/c
End If
txt(11).Enabled = False '-- Amt
mskOS(10).Enabled = False '--Due Dt
txtCrBalance.Enabled = False
txtDrBalance.Enabled = False
txttmpTotBag.Enabled = False
txttmpTotWt.Enabled = False
txtCGSTAmt.Enabled = False
txtSGSTAmt.Enabled = False
txtIGSTAmt.Enabled = False
txt(14).Enabled = False
txt(10).Enabled = False
txt(57).Enabled = False
txt(58).Enabled = False
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
Dim TaxAcCodeCGST As Long
Dim TaxAcCodeSGST As Long
Dim TaxAcCodeIGST As Long
Dim X As String
Dim BillDt
Dim DueDt
Dim TaxAmt As Double
'--- Account Code
MillCode = GProcGetColumnValue("tblMastAccount", "AcName", txttemp(5), "S", "AcCode", "N")
CrCode = GProcGetColumnValue("tblMastAccount", "AcName", txttemp(3), "S", "AcCode", "N")
DrCode = GProcGetColumnValue("tblMastAccount", "AcName", txttemp(4), "S", "AcCode", "N")
BrkCode = GProcGetColumnValue("tblMastAccount", "AcName", txttempOS(6), "S", "AcCode", "N")
TaxAcCodeCGST = GProcGetColumnValue("tblMastNarration", "Narration", "G S T", "S", "CGSTInPutAc", "N")
TaxAcCodeSGST = GProcGetColumnValue("tblMastNarration", "Narration", "G S T", "S", "SGSTInPutAc", "N")
TaxAcCodeIGST = GProcGetColumnValue("tblMastNarration", "Narration", "G S T", "S", "IGSTInPutAc", "N")
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
            .TextMatrix(i, 25) = 0 ' Booking It ctrl No =0
        End If
        If IsDate(.TextMatrix(i, 2)) Then
            rsVouSub!vsubBookDt = .TextMatrix(i, 2)
        Else
            rsVouSub!vsubBookDt = Null
        End If
        '--Itcode(4)
        
        rsVouSub!vsubITCODE = GProcGetColumnValue("tblMastItem", "ItName", .TextMatrix(i, 3), "S", "ItCode", "N", "(ItGstAcInPY = 1 or ItMillCode=" & MillCode & ")") '--Item Code
        rsVouSub!vsubBAG = CLng(.TextMatrix(i, 6)) '--Bag(5)
        rsVouSub!vsubwt = CDbl(.TextMatrix(i, 7)) '-- Wt(6)
        rsVouSub!vsubRt = CDbl(.TextMatrix(i, 8)) '--Rate(7)
        rsVouSub!vsubRTPER = CDbl(.TextMatrix(i, 9)) '--Rate Per(8)
        rsVouSub!vsubAMT = CDbl(.TextMatrix(i, 10)) '--Amt(9)
        rsVouSub!vsubLOTNO = .TextMatrix(i, 5) '--Lot No(10)
        '--Godown Code(17)
        If txtGodown = "" Then
           txtGodown = "-"
        End If
        rsVouSub!VSubGodownCode = GProcGetColumnValue("tblMastNarration", "Narration", txtGodown, "S", "NarrCode", "N", "")
        rsVouSub!vsubLRNO = txtVsub(19) & vbNullString
        rsVouSub!VSUBTMP7 = txtVsub(30) & vbNullString
        If IsDate(mskVouSub(20)) Then '--Vdt
           rsVouSub!vsubLRDate = CDate(mskVouSub(20))
        Else
           rsVouSub!vsubLRDate = Null
        End If
        
        '--- It Ctrl No(11)
        If CLng(.TextMatrix(i, 20)) = 0 Then 'FormAction = vbDataActionAddNew Then
            rsVouSub!VSubItCtrlNo = GProcGenerateId("tblVouSub", "VSubItCtrlNo", "VYear=" & gCYear)  '--It Ctrl No
        Else
            rsVouSub!VSubItCtrlNo = CLng(.TextMatrix(i, 20))
        End If
        '-- P It Ctrl No =Booking It ctrl No
        rsVouSub!VSubPItCtrlNo = CLng(.TextMatrix(i, 25))
        rsVouSub!VSubIssBag = CLng(.TextMatrix(i, 21)) '--Issue Bag(13)
        rsVouSub!VSubIsswt = CDbl(.TextMatrix(i, 22)) '-- Issue Wt(14)
        rsVouSub!VSubTmpBag = CLng(.TextMatrix(i, 23)) '--Temp Bag(15)
        rsVouSub!VSubTmpWt = CDbl(.TextMatrix(i, 24)) '-- Temp Wt(16)
        rsVouSub!VYear = txt(23) '--VYear(18)
        rsVouSub!VSUBTMP1 = Val(.TextMatrix(i, 11)) '- Add
        rsVouSub!VsubLess = Val(.TextMatrix(i, 12)) '- Less
        rsVouSub!VSubTaxableAmt = Val(.TextMatrix(i, 13))
        rsVouSub!VsubCGSTRt = Val(.TextMatrix(i, 14))
        rsVouSub!VsubCGSTAmt = Val(.TextMatrix(i, 15))
        rsVouSub!VsubSGSTRt = Val(.TextMatrix(i, 16))
        rsVouSub!VsubSGSTAmt = Val(.TextMatrix(i, 17))
        rsVouSub!VsubIGSTRt = Val(.TextMatrix(i, 18))
        rsVouSub!VsubIGSTAmt = Val(.TextMatrix(i, 19))
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
'-- Voucher Entry
If frmTypeStr <> "PI" Then
    TaxAmt = Val(txtCGSTAmt) + Val(txtSGSTAmt) + Val(txtIGSTAmt) + Val(txt(58))
    GProcCreateVoucher FormAction, txt(0), 1, txt(2), txt(1), CrCode, DrCode, (txt(11)), txtVou(11), BillDt, "", "Bill No " & txtVou(11) & " Dtd. " & BillDt
    GProcCreateVoucher FormAction, txt(0), 2, txt(2), txt(1), DrCode, CrCode, -(txt(11) - TaxAmt), txtVou(11), BillDt, "", "Bill No " & txtVou(11) & " Dtd. " & BillDt
    i = 3
    If Val(txtCGSTAmt) > 0 Then
       GProcCreateVoucher FormAction, txt(0), i, txt(2), txt(1), TaxAcCodeCGST, CrCode, Val(txtCGSTAmt) * -1, txtVou(11), BillDt, "", "Bill No " & txtVou(11) & " Dtd. " & BillDt
       i = i + 1
       DoEvents
    End If
    If Val(txtSGSTAmt) > 0 Then
       GProcCreateVoucher FormAction, txt(0), i, txt(2), txt(1), TaxAcCodeSGST, CrCode, Val(txtSGSTAmt) * -1, txtVou(11), BillDt, "", "Bill No " & txtVou(11) & " Dtd. " & BillDt
       i = i + 1
       DoEvents
    End If
    If Val(txtIGSTAmt) > 0 Then
       GProcCreateVoucher FormAction, txt(0), i, txt(2), txt(1), TaxAcCodeIGST, CrCode, Val(txtIGSTAmt) * -1, txtVou(11), BillDt, "", "Bill No " & txtVou(11) & " Dtd. " & BillDt
       i = i + 1
       DoEvents
    End If
    
    If Val(txt(58)) > 0 Then
       GProcCreateVoucher FormAction, txt(0), i, txt(2), txt(1), gAcCodeTcsRec, CrCode, Val(txt(58)) * -1, txtVou(11), BillDt, "", "Bill No " & txtVou(11) & " Dtd. " & BillDt
       i = i + 1
       DoEvents
    End If
    
    '--TDS
    If Val(txt(19)) > 0 Then
       Dim TDSAc As Long
       TDSAc = GProcGetColumnValue("tblMastAccount", "AcName", txttemp(28), "S", "AcCode", "N")
       
       GProcCreateVoucher FormAction, txt(0), i, txt(2), txt(1), txt(3), TDSAc, Val(txt(19)) * -1, txtVou(11), BillDt, "", "T.D.S. On Rs. " + txt(11) + " @ " + txt(18) + " %", "Against Bill No " + txtVou(11) + " Dtd. " & BillDt
       i = i + 1
       GProcCreateVoucher FormAction, txt(0), i, txt(2), txt(1), TDSAc, txt(3), Val(txt(19)), txtVou(11), BillDt, "", "T.D.S. On Rs. " + txt(11) + " @ " + txt(18) + " %", "Against Bill No " + txtVou(11) + " Dtd. " & BillDt
       i = i + 1
    End If
    
    
    
    '======================Outstanding Creation
    If IsDate(mskOS(10)) Then
        DueDt = mskOS(10)
    Else
        DueDt = ""
    End If
    GProcCreateOutStanding txt(0), CDate(txt(2)), txt(1), txtVou(11), CrCode, BrkCode, CDbl(txt(11)) - Val(txt(19)), CLng(txtOS(9)), DueDt, MillCode, CDbl(txt(58)), PurBlDt:=mskVou(12)
End If
End Sub
Private Sub cmdLedger_Click()
'On Error GoTo ErrorHandler
frmRptAccount.txt(1) = txttemp(4)
frmMain.mnuRptRSubAccLeg_Click (0)
Exit Sub
ErrorHandler:
    GProcErrorHandler
End Sub
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case 13:
            If Me.ActiveControl.Name <> "txtGrid" Then Sendkeys "{TAB}"
        Case 27:
            If (cmdBtn(13).Enabled = False) And txtGrid.Enabled = False Then
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
    X = "Select * from " & tblName & " where Vtype='" & frmTypeStr & "' and  vno = 0  and VYear=" & gCYear
    If Trim(OrderField) <> "" Then X = X & " order by " & OrderField
    CriteriaStr = " VType = '" & frmTypeStr & "' and VYear=" & gCYear
    GProcRstOpen rstbl, X, "O"
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
Set rsVouSubTmp = New Recordset
'---TblVoucher
X = "Select * from tblVoucher where Vtype='" & frmTypeStr & "' and vno=" & txt(0) & " and VYear=" & gCYear & " order by VCtrNo"
GProcRstOpen rsVou, X, "O"
'---TblOutStanding
X = "Select * from tblOutStanding where Vtype='" & frmTypeStr & "' and vno=" & txt(0) & " and VYear=" & gCYear
GProcRstOpen rsOS, X, "O"

X = "Select * from tblVouSub where Vtype='" & frmTypeStr & "' and vno=" & txt(0) & " and VYear=" & gCYear & " order by VSubCtrNo"
GProcRstOpen rsVouSubTmp, X, "O"
End Sub
Private Sub SetGrid()
Dim X As String
Set rstblSub = New Recordset
X = "select VSubCtrNo,VSubBookNo,VSubBookDt,ItName,Narration,vSubLotNo,VSubBag,VSubWt" _
& " ,VSubRt,VSubRtPer,VSubAmt,VSUBTMP1,VsubLess,VSubTaxableAmt,VsubCGSTRt,VsubCGSTAmt,VsubSGSTRt,VsubSGSTAmt,VsubIGSTRt,VsubIGSTAmt,VSubItCtrlNo,VSubIssBag,VSubIssWt,VSubTmpBag,VSubTmpWt,VSubPItCtrlNo from tblvouSub,TblMastNarration,TblMastItem where " _
& " tblvouSub.VsubItCode=TblMastItem.ItCode and tblvouSub.VSubGodownCode=TblMastNarration.NarrCode " _
& " and tblvouSub.vtype='" & frmTypeStr & "' and tblvouSub.vno=" & txt(0) & " and tblvouSub.VYear=" & gCYear & " order by VSubCtrNo"
GProcRstOpen rstblSub, X, "R"
With msGrid
   .FormatString = ">Sr|<Bk.No|<Bk Date |<Count                               |<Godown                |<Lot No                 |>Bag   |>Weight     |>Rate           |>Per  |>Amount              |>Add      |>Less      |>Taxable Amt. |>C Rt  |>CGST Amt.|>S Rt  |>SGST Amt.|>I Rt    |>IGST Amt.||||||"
   .ColWidth(4) = 0
    For i = 20 To 25
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
End With
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
'    If GProcIsDateValid(mskVou(Index)) = False Then
'        mskVou(Index).SetFocus
'        Exit Sub
'    End If
     mskVouSub(20) = mskVou(12)
End If
If IsDate(mskVou(Index)) = False Then mskVou(Index) = "__/__/____"
End Sub
Private Sub ShowAccountBalance()
Dim DrAcBal  As Double
Dim CrAcBal As Double
'If IsDate(mskDt(2)) Then
'    DrAcBal = GProcGetAccountBalance(txttemp(4), mskDt(2))
'    CrAcBal = GProcGetAccountBalance(txttemp(3), mskDt(2))
'    txtDrBalance = IIf(DrAcBal > 0, DrAcBal & " Cr", Abs(DrAcBal) & " Dr")
'    txtCrBalance = IIf(CrAcBal > 0, CrAcBal & " Cr", Abs(CrAcBal) & " Dr")
'End If
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
If Index = 2 Then 'ShowAccountBalance
If Index = 34 Then
   msGrid.Col = 1
   msGrid_EnterCell
End If
End If
End Sub
Private Sub txt_GotFocus(Index As Integer)
GProcSelectBox Me.ActiveControl
End Sub

Private Sub txtGodown_GotFocus()
GProcSelectBox Me.ActiveControl
End Sub

Private Sub txtGodown_KeyDown(KeyCode As Integer, Shift As Integer)
Dim OldFrmType As String
If KeyCode = vbKeyF3 Then
   OldFrmType = gfrmTypeStr
   gfrmTypeStr = "G"
   GProcShowForm frmMastNarration, frmMain.mnuMstfrm(5), gfrmTypeStr
   gfrmTypeStr = OldFrmType
End If
End Sub

Private Sub txtGodown_KeyPress(KeyAscii As Integer)
gClsSearch.SearchMultiField "tblMastNarration", "Narration", Array("Narration"), Array(txtGodown.Width), "NarrType='G'", Chr(KeyAscii), "Narration", txtGodown.Left + Me.Left, txtGodown.Top + Me.Top + 650, False
KeyAscii = 0
If gClsSearch.SearchMultiRetCol(0) <> "" Then
   txtGodown.text = gClsSearch.SearchMultiRetCol(0)
End If
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
KeyAscii = GProcValidateKey(rsOS, Index, KeyAscii, txtOS(Index).text)
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
Private Sub CheckForNumber()
GProcCheckForNumber Me, rstbl, MaxNo, pOtherTxtFromNo:=OtherTxtFromNo, pOtherTxtToNo:=OtherTxtToNo
End Sub
Private Sub txt_LostFocus(Index As Integer)
CheckForNumber
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
            GProcShowForm frmMastAccount, frmMain.mnuMstfrm(0), Array(frmMastAccount.txttemp(2)), Array(GName)
'        Case 4 '-- Purchase A/c
'            GName = GProcGetColumnValue("TblMastGroup", "AgCode", 90025, "N", "AgName", "S")
'            GProcShowForm frmMastAccount, frmMain.mnuMstfrm(0), Array(frmMastAccount.txtTemp(2)), Array(GName)
        Case 5 '-- millcode
            GName = GProcGetColumnValue("TblMastGroup", "AgCode", 90029, "N", "AgName", "S")
            GProcShowForm frmMastAccount, frmMain.mnuMstfrm(0), Array(frmMastAccount.txttemp(2)), Array(GName)
        Case 35 '-- transport
            OldgFrmTypeStr = gfrmTypeStr
            gfrmTypeStr = "G"
            GProcShowForm frmMastNarration, frmMain.mnuMstfrm(8), "R"
            gfrmTypeStr = OldgFrmTypeStr
    End Select
End If
End Sub
Private Sub txtTemp_KeyPress(Index As Integer, KeyAscii As Integer)
Dim LeftPos As Long
Dim TopPos As Long
Dim AgCode As Long
Dim BrkAcCode As Long
TopPos = txttemp(Index).Top + Me.Top + 650
LeftPos = txttemp(Index).Left + Me.Left
If KeyAscii = 13 Then Exit Sub
Select Case Index
    Case 3  '-- Party   Alias Account
        gClsSearch.SearchMultiField "QryMastaccount", "Account,AcName,AcCity,AgName", Array("Party", "AcName", "City", "Group"), Array(6000, 0, 2000, 3000), "(Agcode=90015 or GpCode=90015 or Agcode=90017 or GpCode=90017 or ((Agcode=90029 or GpCode=90029) and (AcMillType=0  or AcMillType = 5)) )", Chr(KeyAscii), "acName", LeftPos, TopPos, True
        KeyAscii = 0
        If gClsSearch.SearchMultiRetCol(0) <> "" And gClsSearch.SearchMultiRetCol(1) = "" Then
            txttemp(Index) = gClsSearch.SearchMultiRetCol(0)
        Else
            If gClsSearch.SearchMultiRetCol(1) <> "" Then txttemp(Index) = gClsSearch.SearchMultiRetCol(1)
        End If
        lblGSTINNo.Caption = GProcGetColumnValue("tblMastAccount", "AcName", txttemp(3), "S", "AcGSTIN", "S")
        lblStateCd = Left(lblGSTINNo.Caption, 2)
        AgCode = GProcGetColumnValue("tblMastAccount", "AcName", txttemp(3), "S", "AgCode", "N")
        If AgCode = 90029 Then
            txttemp(5) = txttemp(3) '-- Mill =Party
            txttemp(5).Enabled = False
        Else
            txttemp(5).Enabled = True
        End If
    Case 4 '-- Purchase A/c  Alias Account
        If frmTypeStr = "PI" Then
           gClsSearch.SearchMultiField "QryMastaccount", "Account,AcName,AcCode,AcCity,AcMillTypeStr", Array("Sales Account", "AcName", "acCode", "City", "Mill Type"), Array(3000, 0, 0, 2000, 2000), " ((AgCode=90029 or GpCode=90029) and (AcMillType=2))", , "AcName"
        Else
           gClsSearch.SearchMultiField "QryMastaccount", "Account,AcName,AcCity,AgName", Array("Purchase Account", "AcName", "City", "Group"), Array(3000, 0, 2000, 3000), "(AgCode=90025 or GpCode=90025)", Chr(KeyAscii), "AcName", LeftPos, TopPos, True
        End If
        KeyAscii = 0
        If gClsSearch.SearchMultiRetCol(0) <> "" And gClsSearch.SearchMultiRetCol(1) = "" Then
            txttemp(Index) = gClsSearch.SearchMultiRetCol(0)
        Else
            If gClsSearch.SearchMultiRetCol(1) <> "" Then txttemp(Index) = gClsSearch.SearchMultiRetCol(1)
        End If
    Case 5  '-- Mill  Alias Account
        gClsSearch.SearchMultiField "QryMastaccount", "Account,AcName,AcCity,AgName", Array("Mill", "AcName", "City", "Group"), Array(6000, 0, 2000, 3000), "((AgCode=90029 or GpCode=90029) and (AcMillType=0 or AcMillType=5))", Chr(KeyAscii), "AcName", LeftPos, TopPos, True
        KeyAscii = 0
        If gClsSearch.SearchMultiRetCol(0) <> "" And gClsSearch.SearchMultiRetCol(1) = "" Then
            txttemp(Index) = gClsSearch.SearchMultiRetCol(0)
        Else
            If gClsSearch.SearchMultiRetCol(1) <> "" Then txttemp(Index) = gClsSearch.SearchMultiRetCol(1)
        End If
        MillCtrlEd
    Case 35  '-- Mill  Alias Account
         gClsSearch.SearchMultiField "tblMastNarration", "Narration", Array("Narration"), Array(txttemp(35).Width), "NarrType='R'", Chr(KeyAscii), "Narration", txttemp(35).Left + Me.Left, txttemp(35).Top + Me.Top + 650, False
         KeyAscii = 0
         If gClsSearch.SearchMultiRetCol(0) <> "" Then
            txttemp(Index) = gClsSearch.SearchMultiRetCol(0)
         End If
    Case 28 '-- TDS A/c  Alias Account
        gClsSearch.SearchMultiField "QryMastaccount", "Account,AcName,AcCity,AgName", Array("Account Name", "AcName", "City", "Group"), Array(3000, 0, 2000, 3000), "", Chr(KeyAscii), "AcName", 0, 0, True
        KeyAscii = 0
        If gClsSearch.SearchMultiRetCol(0) <> "" And gClsSearch.SearchMultiRetCol(1) = "" Then
            txttemp(Index) = gClsSearch.SearchMultiRetCol(0)
        Else
            If gClsSearch.SearchMultiRetCol(1) <> "" Then txttemp(Index) = gClsSearch.SearchMultiRetCol(1)
        End If
End Select
End Sub
Private Sub MillCtrlEd()
'Dim ExemptMill As Long
''-- Is Exempt Mill
'ExemptMill = GProcGetColumnValue("TblMastAccount", "AcName", txttemp(5), "S", "ACIsExemptMill", "N")
'If ExemptMill = 1 Then
'    If FormAction = vbDataActionAddNew Or FormAction = vbDataActionUpdate Then
'        txt(15).Enabled = True
'        txt(16).Enabled = True
'        txt(15).Visible = True
'        txt(16).Visible = True
'    End If
'Else
'    txt(15).Enabled = False
'    txt(16).Enabled = False
'    txt(15).Visible = False
'    txt(16).Visible = False
'    txt(15) = 0
'    txt(16) = 0
'End If
End Sub
Private Sub txtTemp_LostFocus(Index As Integer)
SaveBtnEd
MillCtrlEd
'ShowAccountBalance
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
    gClsSearch.SearchMultiField "QryMastaccount", "Account,AcName,AcCity,AgName", Array("Broker", "AcName", "City", "Group"), Array(6000, 0, 2000, 3000), "(AgCode=90016 or GpCode=90016)", Chr(KeyAscii), "AcName", LeftPos, TopPos, True
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
           GProcShowForm frmMastAccount, frmMain.mnuMstfrm(0), Array(frmMastAccount.txttemp(2)), Array(GName)
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
    If msGrid.text <> "" Then txtGrid.text = msGrid.text
    '--- Lock TextBox
    '--- Check is Related Gate Pass
    If IsRelatedItCtrlGatePass(msGrid.TextMatrix(msGrid.Row, 20)) = False Then
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
        If (Val(.TextMatrix(.Row, 10)) = 0) And .Rows > 2 Then
            .RemoveItem (.Row)
            txtGrid.text = ""
        End If
        txtGrid.Visible = False
        txt(49).SetFocus
        Exit Sub
    End If
End With
'If FormAction <> vbDataActionAddNew And FormAction <> vbDataActionUpdate Then Exit Sub
'---- Navigate Row ,Col
If KeyAscii = 13 Then
    KeyAscii = 0
    With msGrid
        If .Col = 20 Then 'rstblSub.Fields.Count - 2 Then
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
If msGrid.Col = 4 Then
   msGrid.Col = 5
   If msGrid.Enabled = True Then
      msGrid.SetFocus
   End If
End If

If msGrid.Col = 8 Then
   msGrid.Col = 10
   If msGrid.Enabled = True Then
      msGrid.SetFocus
   End If
End If

If msGrid.text <> "" Then txtGrid.text = msGrid.text
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
                txtGrid.text = txtGrid.text & Chr(KeyAscii)
            End If
        Case Else
            txtGrid.text = Chr(GProcValidateKey(rstblSub, msGrid.Col, KeyAscii, txtGrid.text)) 'Chr(GProcNumberWithDecimal(KeyAscii, txtGrid))
    End Select
End If
DoEvents
If msGrid.text <> "" Then
    txtGrid.SelStart = Len(txtGrid.text) + 1
Else
    txtGrid.SelStart = 1
End If
DoEvents
If txtGrid.Visible = True Then txtGrid.SetFocus
End Sub
Private Sub msGrid_LeaveCell()
If txtGrid.Visible Then
    msGrid.text = txtGrid.text
    txtGrid.text = ""
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
MillCode = GProcGetColumnValue("TblMastAccount", "AcName", txttemp(5), "S", "Accode", "N")
StdPack = GProcGetColumnValue("TblMastItem", "ITName", msGrid.TextMatrix(msGrid.Row, 3), "S", "ItStdpack", "N", " ItMillCode=" & MillCode)
If msGrid.Row = 0 Then Exit Sub
'-- Wt=Bag * Std Pack
If msGrid.TextMatrix(msGrid.Row, 7) = 0 Then
    msGrid.TextMatrix(msGrid.Row, 7) = msGrid.TextMatrix(msGrid.Row, 6) * StdPack
End If
' Rate = (Amt/Wt)* RtPer
If CDbl(msGrid.TextMatrix(msGrid.Row, 7)) > 0 Then
    msGrid.TextMatrix(msGrid.Row, 8) = (Val(msGrid.TextMatrix(msGrid.Row, 10)) / Val(msGrid.TextMatrix(msGrid.Row, 7))) * Val(msGrid.TextMatrix(msGrid.Row, 9))
Else
   msGrid.TextMatrix(msGrid.Row, 8) = 0
End If
msGrid.TextMatrix(msGrid.Row, 10) = Format(msGrid.TextMatrix(msGrid.Row, 10), FStr)

msGrid.TextMatrix(msGrid.Row, 13) = Val(msGrid.TextMatrix(msGrid.Row, 10)) + Val(msGrid.TextMatrix(msGrid.Row, 11)) - Val(msGrid.TextMatrix(msGrid.Row, 12))
msGrid.TextMatrix(msGrid.Row, 13) = Format(msGrid.TextMatrix(msGrid.Row, 13), FStr)
' - cgst Rt = 0 Cgstamt = 0
If msGrid.TextMatrix(msGrid.Row, 14) = 0 Then
   msGrid.TextMatrix(msGrid.Row, 15) = 0
End If
' - sgst Rt = 0 Sgstamt = 0
If msGrid.TextMatrix(msGrid.Row, 16) = 0 Then
   msGrid.TextMatrix(msGrid.Row, 17) = 0
End If
' - Igst Rt = 0 Igstamt = 0
If msGrid.TextMatrix(msGrid.Row, 18) = 0 Then
   msGrid.TextMatrix(msGrid.Row, 19) = 0
End If


'--CGST
If msGrid.TextMatrix(msGrid.Row, 15) = 0 Then
    msGrid.TextMatrix(msGrid.Row, 15) = (Val(msGrid.TextMatrix(msGrid.Row, 13)) * Val(msGrid.TextMatrix(msGrid.Row, 14))) / 100
    msGrid.TextMatrix(msGrid.Row, 15) = GProcMakeRounding(msGrid.TextMatrix(msGrid.Row, 15))
    msGrid.TextMatrix(msGrid.Row, 15) = Format(msGrid.TextMatrix(msGrid.Row, 15), FStr)
End If
'--SGST
If msGrid.TextMatrix(msGrid.Row, 17) = 0 Then
    msGrid.TextMatrix(msGrid.Row, 17) = (Val(msGrid.TextMatrix(msGrid.Row, 13)) * Val(msGrid.TextMatrix(msGrid.Row, 16))) / 100
    msGrid.TextMatrix(msGrid.Row, 17) = GProcMakeRounding(msGrid.TextMatrix(msGrid.Row, 17))
    msGrid.TextMatrix(msGrid.Row, 17) = Format(msGrid.TextMatrix(msGrid.Row, 17), FStr)
End If
'--IGST
If msGrid.TextMatrix(msGrid.Row, 19) = 0 Then
    msGrid.TextMatrix(msGrid.Row, 19) = (Val(msGrid.TextMatrix(msGrid.Row, 13)) * Val(msGrid.TextMatrix(msGrid.Row, 18))) / 100
    msGrid.TextMatrix(msGrid.Row, 19) = GProcMakeRounding(msGrid.TextMatrix(msGrid.Row, 19))
    msGrid.TextMatrix(msGrid.Row, 19) = Format(msGrid.TextMatrix(msGrid.Row, 19), FStr)
End If

'------------
msGrid.TextMatrix(msGrid.Row, 11) = Format(msGrid.TextMatrix(msGrid.Row, 11), LStr)
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
If IsRelatedItCtrlGatePass(msGrid.TextMatrix(msGrid.Row, 20)) = True Then Exit Sub
'--- Delete Row On Ctrl+Y
If KeyCode = vbKeyY And Shift = 2 And msGrid.Rows > 2 Then
    msGrid = 4
    msGrid.RemoveItem (msGrid.Row)
    txtGrid.Visible = False
    txtGrid.text = ""
    ReNumberGridSrno
    CalAmount
    Exit Sub
End If
'--- Show Form
If KeyCode = vbKeyF3 Then
Select Case msGrid.Col
    Case 3 '-- Item
        GProcShowForm frmMastItem, frmMain.mnuMstfrm(3), Array(frmMastItem.txttemp(9)), Array(txttemp(5))
        Exit Sub
    Case 4 '-- Godown
        OldgFrmTypeStr = gfrmTypeStr
        gfrmTypeStr = "G"
        GProcShowForm frmMastNarration, frmMain.mnuMstfrm(8), "G"
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
If IsRelatedItCtrlGatePass(msGrid.TextMatrix(msGrid.Row, 20)) = True Then Exit Sub
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
                    MillCode = GProcGetColumnValue("TblMastAccount", "AcName", txttemp(5), "S", "Accode", "N")
                   
                    '-- For Ittype = ITTMP1
'                    gClsSearch.SearchMultiField "tblMastItem,tblMastaccount", "ItName,AcName,ItStdpack,ItStdrateper,ITTMP2,ItGstAcInPY", Array("Count Name", "Mill", "Std. Pack", "Rate Per", "", ""), Array(txtGrid.Width, 2000, 1000, 1000, 0, 0), "tblMastItem.ITMillCode=tblMastaccount.acCode and  (ITMillCode= " & MillCode & " or ItGstAcInPY = 1)", Chr(KeyAscii), "ItName", LeftPos, TopPos, True
                    
                    '-- For Ittype = ItType
                    gClsSearch.SearchMultiField "tblMastItem,tblMastaccount", "ItName,AcName,ItStdpack,ItStdrateper,ItType,ItGstAcInPY,ItTicket,Ittmp2", Array("Count Name", "Mill", "Std. Pack", "Rate Per", "", "", "Ticket", "HSN"), Array(txtGrid.Width, 2000, 1000, 1000, 0, 0, 1800, 1200), "tblMastItem.ITMillCode=tblMastaccount.acCode and  (ITMillCode= " & MillCode & " or ItGstAcInPY = 1)", Chr(KeyAscii), "ItName", LeftPos, TopPos, True
                    txtGrid = ""
                    KeyAscii = 0
                    If gClsSearch.SearchMultiRetCol(0) <> "" Then '-- Item Wt ,Rate From Master
                        txtGrid = gClsSearch.SearchMultiRetCol(0)
                        If IsNumeric(.TextMatrix(.Row, 7)) = False Then .TextMatrix(.Row, 7) = 0
                        If IsNumeric(.TextMatrix(.Row, 9)) = False Then .TextMatrix(.Row, 9) = 0
                        If .TextMatrix(.Row, 9) = 0 Then .TextMatrix(.Row, 9) = gClsSearch.SearchMultiRetCol(3)
                        If gClsSearch.SearchMultiRetCol(5) = 1 Then '-- Insureace , Freight
                               If Left(lblGSTINNo.Caption, 2) = "27" Or Left(lblGSTINNo.Caption, 1) = "U" Or Trim(lblGSTINNo.Caption) = "" Then
                                  .TextMatrix(.Row, 14) = GProcGetColumnValue("tblMastItem", "Itname", "" & gClsSearch.SearchMultiRetCol(0) & "", "S", "ItCGSTRt", "N")
                                  .TextMatrix(.Row, 16) = GProcGetColumnValue("tblMastItem", "Itname", "" & gClsSearch.SearchMultiRetCol(0) & "", "S", "ItSGSTRt", "N")
                                  .TextMatrix(.Row, 18) = 0
                               Else
                                  .TextMatrix(.Row, 14) = 0
                                  .TextMatrix(.Row, 16) = 0
                                  .TextMatrix(.Row, 18) = GProcGetColumnValue("tblMastItem", "Itname", "" & gClsSearch.SearchMultiRetCol(0) & "", "S", "ItIGSTRt", "N")
                               End If
                               .TextMatrix(.Row, 6) = 1
                               .TextMatrix(.Row, 7) = 1
                               .TextMatrix(.Row, 8) = 1
                               .TextMatrix(.Row, 9) = 1
                               .Col = 9
                               msGrid.SetFocus
                        Else
                            If gClsSearch.SearchMultiRetCol(4) = 0 Then
                               If Left(lblGSTINNo.Caption, 2) = "27" Or Left(lblGSTINNo.Caption, 1) = "U" Or Trim(lblGSTINNo.Caption) = "" Then
                                  .TextMatrix(.Row, 14) = GProcGetColumnValue("tblMastNarration", "Narration", "G S T", "S", "CotCGSTRt", "N")
                                  .TextMatrix(.Row, 16) = GProcGetColumnValue("tblMastNarration", "Narration", "G S T", "S", "CotSGSTRt", "N")
                                  .TextMatrix(.Row, 18) = 0
                               Else
                                  .TextMatrix(.Row, 14) = 0
                                  .TextMatrix(.Row, 16) = 0
                                  .TextMatrix(.Row, 18) = GProcGetColumnValue("tblMastNarration", "Narration", "G S T", "S", "CotIGSTRt", "N")
                               End If
                            Else
                               If Left(lblGSTINNo.Caption, 2) = "27" Or Left(lblGSTINNo.Caption, 1) = "U" Or Trim(lblGSTINNo.Caption) = "" Then
                                  .TextMatrix(.Row, 14) = GProcGetColumnValue("tblMastNarration", "Narration", "G S T", "S", "PolCGSTRt", "N")
                                  .TextMatrix(.Row, 16) = GProcGetColumnValue("tblMastNarration", "Narration", "G S T", "S", "PolSGSTRt", "N")
                                  .TextMatrix(.Row, 18) = 0
                               Else
                                  .TextMatrix(.Row, 14) = 0
                                  .TextMatrix(.Row, 16) = 0
                                  .TextMatrix(.Row, 18) = GProcGetColumnValue("tblMastNarration", "Narration", "G S T", "S", "PolIGSTRt", "N")
                               End If
                            End If
                        End If
                        txtGrid_KeyPress 13
                    End If
                  End If
            Case 4  '-- Godown
                gClsSearch.SearchMultiField "tblMastNarration", "Narration", Array("Godown"), Array(3000), " NarrType='G' ", Chr(KeyAscii), "Narration", LeftPos, TopPos, True
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
    If KeyAscii = 13 And .Col < 19 Then 'rstblSub.Fields.Count - 2 Then
        .Col = .Col + 1
    ElseIf KeyAscii = 13 And .Col = 19 Then '= rstblSub.Fields.Count - 2 Then
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
        If (Val(.TextMatrix(.Row, 10)) = 0) And .Rows > 2 Then
            .RemoveItem (.Row)
            txtGrid.text = ""
        End If
        txtGrid.Visible = False
        txt(49).Enabled = True
        txt(49).SetFocus
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
Dim IsInsuItem As Long
Dim X As String
Dim Rs1 As Recordset
Dim oldFrmTypeStr As String
Dim LRNo As String
Dim pItCtrlNo As Long
If Trim(txtGrid.text) <> "" Then
    msGrid.text = txtGrid.text
    CalGridAmt
End If
txtGrid.Visible = False
txtGrid.text = ""
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
                    MillCode = GProcGetColumnValue("TblMastAccount", "AcName", txttemp(5), "S", "Accode", "N")
                    IsInsuItem = GProcGetColumnValue("tblMastItem", "ItName", .TextMatrix(RowIndex, 3), "S", "ItGstAcInPY", "N", "(ItGstAcInPY = 1 or ItMillCode=" & MillCode & ")") '--Item Code
                    
                    If GProcGetColumnValue("TblMastItem", "ItName", .TextMatrix(RowIndex, i), "S", "Itcode", "N", "ITMillCode=" & MillCode) = 0 And IsInsuItem = 0 Then
                        MsgBox "Check Count.Mill and Count are Mismatch.", vbCritical + vbOKOnly, Me.Caption
                        .Col = i
                        msGrid.SetFocus
                        Exit Function
                    End If
'                Case 4 '-- Godown
'                    If Trim(.TextMatrix(RowIndex, i)) = "" Then
'                        MsgBox "Check Godown.", vbCritical + vbOKOnly, Me.Caption
'                        .Col = i
'                        msGrid.SetFocus
'                        Exit Function
'                    End If
'                    If GProcGetColumnValue("TblMastNarration", "Narration", .TextMatrix(RowIndex, i), "S", "Narrcode", "N") = 0 Then
'                        'MSGrid.SetFocus
'                        oldFrmTypeStr = gfrmTypeStr
'                        gfrmTypeStr = "G"
'                        GProcShowForm frmMastNarration, frmMain.mnuMstfrm(8), Array(frmMastNarration.txt(2)), Array(.TextMatrix(RowIndex, i))
'                        gfrmTypeStr = oldFrmTypeStr
'                        Exit Function
'                    End If
                    
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
'                    If CDbl(.TextMatrix(RowIndex, 7)) < CDbl(.TextMatrix(RowIndex, 13)) Then
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
'                    '-- If Wt < Issue wt
'                    If CDbl(.TextMatrix(RowIndex, 8)) < CDbl(.TextMatrix(RowIndex, 14)) Then
'                        MsgBox "Check Wt.Weight Less than issued.", vbCritical + vbOKOnly, Me.Caption
'                        .Col = i
'                        msGrid.SetFocus
'                        Exit Function
'                    End If
                Case 10 '-- amount
                    If CDbl(.TextMatrix(RowIndex, i)) = 0 Then
                        MsgBox "Check amount.", vbCritical + vbOKOnly, Me.Caption
                        .Col = i
                        msGrid.SetFocus
                        Exit Function
                    End If
                Case 15 '-- CGST Amt
                     If Val(.TextMatrix(RowIndex, 15)) > 0 Then
                        If Abs(Round((Val(.TextMatrix(RowIndex, 13)) * Val(.TextMatrix(RowIndex, 14))) / 100, 2) - Val(.TextMatrix(RowIndex, 15))) > 1 Then
                           MsgBox "Check CGST Amount", vbCritical + vbOKOnly, Me.Caption
                           .Col = 15
                           msGrid.SetFocus
                           Exit Function
                        End If
                     End If
                Case 17 '-- SGST Amt
                     If Val(.TextMatrix(RowIndex, 17)) > 0 Then
                        If Abs(Round((Val(.TextMatrix(RowIndex, 13)) * Val(.TextMatrix(RowIndex, 16))) / 100, 2) - Val(.TextMatrix(RowIndex, 17))) > 1 Then
                           MsgBox "Check SGST Amount", vbCritical + vbOKOnly, Me.Caption
                           .Col = 17
                           msGrid.SetFocus
                           Exit Function
                        End If
                     End If
                Case 19 '-- IGST Amt
                     If Val(.TextMatrix(RowIndex, 19)) > 0 Then
                        If Abs(Round((Val(.TextMatrix(RowIndex, 13)) * Val(.TextMatrix(RowIndex, 18))) / 100, 2) - Val(.TextMatrix(RowIndex, 19))) > 1 Then
                           MsgBox "Check SGST Amount", vbCritical + vbOKOnly, Me.Caption
                           .Col = 19
                           msGrid.SetFocus
                           Exit Function
                        End If
                     End If
           End Select
        Next i
        '--- Duplicate LR No ,It Ctrl  No (Booking)
'        LRNo = .TextMatrix(RowIndex, 6)
'        pItCtrlNo = .TextMatrix(RowIndex, 19)
'        For i = 1 To .Rows - 1
'            '--- Duplicate LR No
'            If i <> RowIndex And .TextMatrix(i, 6) = LRNo Then
'                MsgBox "Duplicate LR No.", vbCritical + vbOKOnly, Me.Caption
'                .Col = 6
'                msGrid.SetFocus
'                Exit Function
'            End If
'            '-- Duplicate Booking Count Selection
' '           If i <> RowIndex And .TextMatrix(i, 19) = pItCtrlNo And pItCtrlNo > 0 Then
'  '              MsgBox "Duplicate Booking Count Selection.", vbCritical + vbOKOnly, Me.Caption
'   '             .Col = 1
'    '            msGrid.SetFocus
'     '           Exit Function
'      '      End If
'        Next i
        '-- Invalid Booking
        If CLng(.TextMatrix(RowIndex, 1)) = 0 Then
                MsgBox "Purchase Save Without Booking............", vbInformation + vbOKOnly, Me.Caption
                .Col = 1
                msGrid.SetFocus
        End If
        
        
        If CLng(.TextMatrix(RowIndex, 1)) > 0 Then
            MillCode = GProcGetColumnValue("TblMastAccount", "AcName", txttemp(5), "S", "Accode", "N")
            ItCode = GProcGetColumnValue("tblMastItem", "ItName", .TextMatrix(RowIndex, 3), "S", "ItCode", "N", " ItMillCode=" & MillCode)
            X = "Select * from tblvouSub where Vno=" & CLng(.TextMatrix(RowIndex, 1)) & " and  " _
            & " VSUBITCODE=" & ItCode & " and  vsubitctrlno=" & .TextMatrix(RowIndex, 25) & " and Vtype='OT' and "
            If gBackEndDB = gBackEndAccess Then
                X = X + "  vdt=cdate('" & CDate(.TextMatrix(RowIndex, 2)) & "')"
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
Private Sub SaveBtnEd()
'--- Vno,vdt,Party,A/c,Mill,Amount
If CDbl(txt(0)) = 0 Or GProcIsDateValid(mskDt(2)) = False Or txttemp(3) = "" Or txttemp(5) = "" Or txttemp(4) = "" Or CDbl(txt(11)) = 0 Then
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
    '--TDS
    If Val(txt(19)) > 0 And (txttemp(28) = "" Or txttemp(28) = "-") Then
       MsgBox "TDS A/c Not Found...", vbInformation + vbOKOnly, Me.Caption
       txttemp(28).SetFocus
       Exit Function
    End If
''    If mskDt(2) >= CDate("20/04/2012") Then
''        MsgBox "Demo Experied Please Contact on MBL - 9922997663", vbInformation + vbOKOnly, Me.Caption
''        txt(0).SetFocus
''        Exit Function
''    End If
    '---Validate Vou.Date = Bill Date
'    If mskVou(12) <> mskDt(2) Then
'        MsgBox "Vou.Date Muse Be Same As Bill Date", vbInformation + vbOKOnly, Me.Caption
'        mskDt(2).SetFocus
'        Exit Function
'    End If
    If mskVou(12) < gCYSDate Or mskVou(12) > gCYEDate Then
        MsgBox "Check Bill Date.", vbInformation + vbOKOnly, Me.Caption
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
    '---- Party Account
    If txttemp(3) = "" Then
        MsgBox "Check Party.", vbInformation + vbOKOnly, Me.Caption
        txttemp(3).SetFocus
        Exit Function
    End If
        '--Create Party
        If GProcGetColumnValue("tblMastAccount", "AcName", txttemp(3), "S", "AcCode", "N") = 0 Then
           GName = GProcGetColumnValue("TblMastGroup", "AgCode", 90015, "N", "AgName", "S")
           GProcShowForm frmMastAccount, frmMain.mnuMstfrm(0), Array(frmMastAccount.txt(1), frmMastAccount.txttemp(2)), Array(txttemp(3), GName)
           Exit Function
        End If
    '---- Purchase Account
    If txttemp(4) = "" Then
        MsgBox "Check Purchase Account.", vbInformation + vbOKOnly, Me.Caption
        txttemp(4).SetFocus
        Exit Function
    End If
        '--Create Party
        If GProcGetColumnValue("tblMastAccount", "AcName", txttemp(4), "S", "AcCode", "N") = 0 Then
           GName = GProcGetColumnValue("TblMastGroup", "AgCode", 90025, "N", "AgName", "S")
           GProcShowForm frmMastAccount, frmMain.mnuMstfrm(0), Array(frmMastAccount.txt(1), frmMastAccount.txttemp(2)), Array(txttemp(4), GName)
           Exit Function
        End If
    '---- Mill
    If txttemp(5) = "" Then
        MsgBox "Check Mill.", vbInformation + vbOKOnly, Me.Caption
        txttemp(5).SetFocus
        Exit Function
    End If
        '--Create Mill
        If GProcGetColumnValue("tblMastAccount", "AcName", txttemp(5), "S", "AcCode", "N") = 0 Then
           GName = GProcGetColumnValue("TblMastGroup", "AgCode", 90029, "N", "AgName", "S")
           GProcShowForm frmMastAccount, frmMain.mnuMstfrm(0), Array(frmMastAccount.txt(1), frmMastAccount.txttemp(2)), Array(txttemp(5), GName)
           Exit Function
        End If
    '--- Broker
        '--Create Broker
        If txttempOS(6) = "" Then txttempOS(6) = "-"
        If GProcGetColumnValue("tblMastAccount", "AcName", txttempOS(6), "S", "AcCode", "N") = 0 Then
           GName = GProcGetColumnValue("TblMastGroup", "AgCode", 90016, "N", "AgName", "S")
           GProcShowForm frmMastAccount, frmMain.mnuMstfrm(0), Array(frmMastAccount.txt(1), frmMastAccount.txttemp(2)), Array(txttempOS(6), GName)
           Exit Function
        End If
    '---- Amount
    If CDbl(txt(11)) <= 0 Then
        MsgBox "Check Voucher Amount.", vbInformation + vbOKOnly, Me.Caption
        Exit Function
    End If
    
    If Val(txtIGSTAmt) > 0 And ((Val(txtCGSTAmt) + Val(txtSGSTAmt)) > 0) Then
        MsgBox "Check Tax  Amount.", vbInformation + vbOKOnly, Me.Caption
        Exit Function
    End If
    
    If Val(txtCGSTAmt) > 0 And Val(txtSGSTAmt) = 0 Then
        MsgBox "Check Tax  Amount.", vbInformation + vbOKOnly, Me.Caption
        Exit Function
    End If
    If Val(txtSGSTAmt) > 0 And Val(txtCGSTAmt) = 0 Then
        MsgBox "Check Tax  Amount.", vbInformation + vbOKOnly, Me.Caption
        Exit Function
    End If
    If (Val(txtCGSTAmt) + Val(txtSGSTAmt) + Val(txtIGSTAmt)) > 0 And lblGSTINNo = "" Then
        MsgBox "Check GSTIN of Party", vbInformation + vbOKOnly, Me.Caption
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
If IsDate(mskOS(10)) = False Then mskOS(10) = "__/__/____" '--Due Dt
If IsDate(mskDt(2)) Then '--Vdt
    txt(2) = CDate(mskDt(2))
Else
    txt(2) = ""
End If
If IsDate(mskDt(34)) Then '--Vdt
    txt(34) = CDate(mskDt(34))
Else
    txt(34) = ""
End If
'--- Bill No ,dt
'txt(6) = txtVou(11) '--gpno
txt(7) = mskVou(12) '--gpDt
txt(36) = txtVou(11)
'-- Cr Code
txt(3) = GProcGetColumnValue("tblMastAccount", "AcName", txttemp(3), "S", "AcCode", "N")
'-- Dr Code
txt(4) = GProcGetColumnValue("tblMastAccount", "AcName", txttemp(4), "S", "AcCode", "N")
'--Mill
txt(5) = GProcGetColumnValue("tblMastAccount", "AcName", txttemp(5), "S", "AcCode", "N")
txt(35) = GProcGetColumnValue("tblMastNarration", "Narration", txttemp(35), "S", "NarrCode", "N")
txt(23) = gCYear
'--For TDS
If txttemp(28) = "" Then txttemp(28) = "-"
txt(28) = GProcGetColumnValue("tblMastAccount", "AcName", txttemp(28), "S", "AcCode", "N")
'--
'ShowAccountBalance
End Sub
Private Sub FillTempFromTxt()
Dim X As String
If IsDate(txt(2)) Then '--VDt
    mskDt(2) = txt(2)
Else
    mskDt(2) = "__/__/____"
End If
If IsDate(txt(34)) Then '--VDt
    mskDt(34) = txt(34)
Else
    mskDt(34) = "__/__/____"
End If
If IsDate(txt(9)) Then '--Book Dt
    mskDt(9) = txt(9)
Else
    mskDt(9) = "__/__/____"
End If
txttemp(3) = GProcGetColumnValue("tblMastAccount", "AcCode", txt(3), "N", "AcName", "S")
'--Trans
If txt(35) = "" Then
   txt(35) = GProcGetColumnValue("tblMastNarration", "Narration", "-", "S", "NarrCode", "N")
End If
txttemp(35) = GProcGetColumnValue("tblMastNarration", "NarrCode", txt(35), "N", "Narration", "S")

'-- Cr Code

txttemp(3) = GProcGetColumnValue("tblMastAccount", "AcCode", txt(3), "N", "AcName", "S")
'-- Dr Code
txttemp(4) = GProcGetColumnValue("tblMastAccount", "AcCode", txt(4), "N", "AcName", "S")
'--Mill
txttemp(5) = GProcGetColumnValue("tblMastAccount", "AcCode", txt(5), "N", "AcName", "S")
lblGSTINNo.Caption = GProcGetColumnValue("tblMastAccount", "AcName", txttemp(3), "S", "AcGSTIN", "S")
lblStateCd = Left(lblGSTINNo.Caption, 2)

'--TDS A/c
txttemp(28) = GProcGetColumnValue("tblMastAccount", "AcCode", txt(28), "N", "AcName", "S")

'-- Fill Related Record
FillRelatedRecord
'-- Expenses
'ShowAccountBalance
MillCtrlEd
End Sub
Private Sub FillRelatedRecord()
Dim RowIndex As Long
Dim ColIndex As Long
Dim TotBag As Long
Dim TotWt As Double
Dim Camt As Double
Dim Samt As Double
Dim Iamt As Double
Camt = 0
Samt = 0
Iamt = 0
'--Fill Grid
SetGrid
With rstblSub
    If Not .EOF Then
        .MoveFirst
        txtGodown = rstblSub.Fields("Narration") & vbNullString

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
                If ColIndex = 6 Then TotBag = TotBag + CLng(msGrid.TextMatrix(RowIndex, 6)) '-- Tot Bag
                If ColIndex = 7 Then TotWt = TotWt + CDbl(msGrid.TextMatrix(RowIndex, 7)) '-- Tot Wt
                If ColIndex = 15 Then Camt = Camt + Val(msGrid.TextMatrix(RowIndex, 15))
                If ColIndex = 17 Then Samt = Samt + Val(msGrid.TextMatrix(RowIndex, 17))
                If ColIndex = 19 Then Iamt = Iamt + Val(msGrid.TextMatrix(RowIndex, 19))
                txtGodown = msGrid.TextMatrix(RowIndex, 4)
            Next ColIndex
            RowIndex = RowIndex + 1
            .MoveNext
            If Not .EOF Then msGrid.Rows = msGrid.Rows + 1
        Loop
    End If
End With
txttmpTotBag = TotBag '-- Tot Bag
txttmpTotWt = Format(TotWt, "###0.000") '-- Tot Wt
txtCGSTAmt = Format(Camt, "#########0.00")
txtSGSTAmt = Format(Samt, "#########0.00")
txtIGSTAmt = Format(Iamt, "#########0.00")
'--- Fill Related Control
SetRelatedRS
txtVou(11) = rsVou.Fields("vBillNo") & vbNullString     '--Bill No
If IsDate(rsVou.Fields("vBillDt")) Then '--Bill Date
    mskVou(12) = Format(rsVou.Fields("vBilldt"), "dd/mm/yyyy")
Else
    mskVou(12) = "__/__/____"
End If

If IsDate(rsVouSubTmp.Fields("VSubLRDate")) Then '--Bill Date
   mskVouSub(20) = Format(rsVouSubTmp.Fields("VSubLRDate"), "dd/mm/yyyy")
Else
   mskVouSub(20) = "__/__/____"
End If
txtVsub(19) = rsVouSubTmp.Fields("VSubLRNo") & vbNullString
txtVsub(30) = rsVouSubTmp.Fields("VSUBTMP7") & vbNullString
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
GProcShowRecord Me, rstbl, MaxNo, OtherTxtFromNo, OtherTxtToNo
FillTempFromTxt
End Sub
Private Sub CalAmount()
Dim TaxAmt As Double
Dim SubAmt As Double
Dim Amt As Double
Dim ROff As Double
Dim RowIndex As Integer
Dim TotBag As Long
Dim TotWt As Double
Dim MillAcCode As Long
Dim IsExemptMill As Long
Dim Camt As Double
Dim Samt As Double
Dim Iamt As Double

'-- Is Exempt Mill
'IsExemptMill = GProcGetColumnValue("TblMastAccount", "AcName", txttemp(5), "S", "ACIsExemptMill", "N")
'-- Due Date
If IsNumeric(txtOS(9)) = False Then txtOS(9) = "0"
txtOS(9) = Format(txtOS(9), PStr)
If IsDate(mskDt(2)) Then mskOS(10) = DateAdd("d", txtOS(9), mskDt(2))
'---- Sub amt
CheckForNumberGrid
CheckForNumber
Camt = 0
Samt = 0
Iamt = 0
RowIndex = 1
With msGrid
    Do While RowIndex < .Rows
        '-- sub amt
        If Val(.TextMatrix(RowIndex, 13)) > 0 Then
            SubAmt = SubAmt + Val(.TextMatrix(RowIndex, 13))
        End If
        '--Tot Bag
        If Val(.TextMatrix(RowIndex, 6)) > 0 Then
            TotBag = TotBag + Val(.TextMatrix(RowIndex, 6))
        End If
        '--Tot Wt
        If Val(.TextMatrix(RowIndex, 7)) > 0 Then
            TotWt = TotWt + Val(.TextMatrix(RowIndex, 7))
        End If
        Camt = Camt + Val(.TextMatrix(RowIndex, 15))
        Samt = Samt + Val(.TextMatrix(RowIndex, 17))
        Iamt = Iamt + Val(.TextMatrix(RowIndex, 19))
        RowIndex = RowIndex + 1
    Loop
End With
txttmpTotBag = TotBag
txttmpTotWt = Format(TotWt, "###0.000")

txtCGSTAmt = Format(Camt, "#########0.00")
txtSGSTAmt = Format(Samt, "#########0.00")
txtIGSTAmt = Format(Iamt, "#########0.00")


'--TCS
If Val(txt(56)) > 0 And Val(txt(58)) = 0 Then
   Dim TcsAmt As Double
   txt(57) = 0.1
   TcsAmt = (Val(txt(56)) * Val(txt(57))) / 100
   
   TcsAmt = GProcMakeRounding(TcsAmt)

   txt(58) = TcsAmt

End If


txt(56) = Format(Val(txt(56)), FStr)
txt(57) = Format(Val(txt(57)), WtStr)
txt(58) = Format(Val(txt(58)), FStr)
'-- Exmpt Amt For Exmpt Mill
'--Roff
ROff = txt(17) '= 'CDbl(GProcMakeRounding(Amt) - Amt)
'-- SubAmt+Tax Amt +Add -Less+roff + Aft Tax Add - Aft Tax Less
Amt = SubAmt + Camt + Samt + Iamt + Val(txt(49)) - Val(txt(50)) + ROff + Val(txt(58)) '-- TCS


'---end TCS

'--TDS
Dim X As String
Dim wPurPtyCode As Long
Dim tmpRst As Recordset
Dim TDSLimitAmt As Double
Dim TDSAmt As Double
Dim PtyPan As String

wPurPtyCode = GProcGetColumnValue("tblMastAccount", "AcName", txttemp(3), "S", "AcCode", "N")
PtyPan = GProcGetColumnValue("tblMastAccount", "AcName", txttemp(5), "S", "AcPAN", "S")
If PtyPan = "" Then PtyPan = "XXXX"

TDSAmt = 0
'txt(18) = 0 '--Tds Rate
lblTotPur.Caption = ""
If Format(mskDt(2), "YYYY/MM/dd") >= Format("2021/07/01", "YYYY/MM/dd") And gCTDSFirm = 1 Then
''    X = "Select sum(AdSubAmt) from tblAddLess,tblMastAccount where tblAddLess.AdAcCrCode = tblMastAccount.AcCode and  (AdAcCrCode = " & wPurPtyCode & " or tblMastAccount.AcPan = '" & PtyPan & "'  ) and Format(vdt,'YYYY/MM/dd') <= '" & Format(mskDt(2), "YYYY/MM/dd") & "'"
''    X = X + " and  ltrim(str(vno))+vtype+ltrim(str(vyear)) <> '" & Trim(Str(txt(0))) & "" & gfrmTypeStr & "" & LTrim(Str(gCYear)) & "' and VYear=" & gCYear & " and vno < " & txt(0) & " "
       
    X = "Select sum(AdSubAmt) from tblAddLess,tblMastAccount where tblAddLess.AdAcCrCode = tblMastAccount.AcCode and (AdAcCrCode = " & wPurPtyCode & " or tblMastAccount.AcPan = '" & PtyPan & "'  ) and Format(vdt,'YYYY/MM/dd') <= '" & Format(mskDt(2), "YYYY/MM/dd") & "'"
    X = X + " and  ltrim(str(vno))+vtype+ltrim(str(vyear)) <> '" & Trim(Str(txt(0))) & "" & gfrmTypeStr & "" & LTrim(Str(gCYear)) & "' and VYear=" & gCYear & " and vno < " & txt(0) & " "
       
    i = GProcRstOpen(tmpRst, X, "R")
       
    TDSLimitAmt = IIf(IsNull(tmpRst.Fields(0)), 0, tmpRst.Fields(0))
    TDSLimitAmt = TDSLimitAmt + Val(txt(10))
    lblTotPur.Caption = TDSLimitAmt
    '--------------------------------TCs Amt
    If TDSLimitAmt >= 5000000 And Val(txt(58)) = 0 Then
       txt(18) = 0.1
       txttemp(28) = GProcGetColumnValue("tblMastAccount", "AcCode", CStr(gServiceTaxAcCode), "N", "AcName", "S")
       TDSAmt = (Val(txt(10)) * Val(txt(18))) / 100
       TDSAmt = GProcMakeRounding(TDSAmt)
       TDSAmt = Format(TDSAmt, FStr)
    End If
End If
txt(19) = TDSAmt
'--End TDS

'-- Exmpt Amt For Exmpt Mill
'--Roff
'ROff = txt(17) '= 'CDbl(GProcMakeRounding(Amt) - Amt)
'-- SubAmt+Tax Amt +Add -Less+roff + Aft Tax Add - Aft Tax Less
'Amt = SubAmt + Camt + Samt + Iamt + Val(txt(49)) - Val(txt(50)) + ROff
'--Other Add
MillAcCode = GProcGetColumnValue("tblmastAccount", "AcName", txttemp(5), "S", "AcCode", "N")
If (Val(txt(27)) = 0) And (TotBag > 0) Then txt(27) = Format(GProcGetColumnValue("tblmastitem", "itname", msGrid.TextMatrix(1, 3), "S", "Itbrokrt", "N", " ItMillCode=" & MillAcCode) * TotBag, FStr)
'--- Landing Cost =Bill+Freight+Rent +add-Less
txt(10) = SubAmt
txt(14) = Format(Camt + Samt + Iamt, "#########0.00")
'txt(24) = AssValue
'txt(17) = Roff
txt(11) = Amt
txt(31) = CLng(TotBag) '- Total Bag
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
'---------- Booking Vs Purchase
Private Sub PrepareBooking(Optional Chr As String)
Dim MillCode As Long
Dim QStr As String
Dim RsBk As Recordset
Dim Chr2 As String
If IsDate(mskDt(2)) = True Then
    MillCode = GProcGetColumnValue("TblMastAccount", "AcName", txttemp(5), "S", "Accode", "N")
    QStr = " select tblAddless.Adgpno, tblVousub.VNo,format(tblVousub.vdt,'dd/mm/yyyy'),ItName,vsubbag-(vsubissbag),vsubwt-(vsubisswt),vsubbag,vsubwt,vsubitctrlno,VsubRt from " _
    & " tblVousub,tblmastItem,tblAddless Where " _
    & " tblVousub.vno = tblAddless.vno and  tblVousub.vtype  = tblAddless.vtype and  tblVousub.vyear = tblAddless.vyear and " _
    & " Itcode=VSUBITCODE and ItMillCode=" & MillCode & " and " _
    & " ((vsubbag-(vsubissbag))> 0) and tblVousub.vtype ='OT'  and VSUBBOOKISCOMPLETED=0 and "
    If gBackEndDB = gBackEndAccess Then
        QStr = QStr + " tblVousub.vdt<=cdate('" & CDate(mskDt(2)) & "')"
    Else '-- Oracle
        QStr = QStr + " tblVousub.vdt <= to_date('" & CDate(mskDt(2)) & "','dd/MM/yyyy')"
    End If
    With gClsSearch
        .SearchMultiField QStr, "", Array("Contract No.", " Booking No.", "Date", "Count", "Bal Bag", "Bal Wt.", "Bag", "Wt", "itctrlno", "Rate"), Array(1200, 1200, 1500, 1500, 1000, 1000, 500, 1000, 0, 1000), "", Chr, "tblVousub.Vdt,tblAddless.Adgpno,tblVousub.vno,VSUBCTRNO", 0, 0, False, , True
        If .SearchMultiRetCol(1) <> "" Then
            QStr = "select tblvousub.*,tblmastItem.ItName,tblmastItem.ItType from tblvousub,tblMastItem where " _
            & " ItCode=VSUBITCODE and VSUBITCTRLNO= " & CLng(.SearchMultiRetCol(8))
              If GProcRstOpen(RsBk, QStr, "R") > 0 Then
                    txtGrid = ""
                    txtGrid = CLng(.SearchMultiRetCol(1)) '--Book Vno
                    msGrid.TextMatrix(msGrid.Row, 1) = CLng(.SearchMultiRetCol(1)) '--Book Vno
                    msGrid.TextMatrix(msGrid.Row, 2) = Format(.SearchMultiRetCol(2), "dd/mm/yyyy") '-- Book Vdt
                    msGrid.TextMatrix(msGrid.Row, 3) = RsBk!ItName  '-- Count
                    msGrid.TextMatrix(msGrid.Row, 6) = .SearchMultiRetCol(4)  '-- Bal Bag
                    msGrid.TextMatrix(msGrid.Row, 7) = .SearchMultiRetCol(5) '-- Bal Wt
                    msGrid.TextMatrix(msGrid.Row, 25) = RsBk!VSubItCtrlNo  '--  P ItCtrl No =Book ItCtrl No
                    msGrid.TextMatrix(msGrid.Row, 8) = RsBk!vsubRt  '-- Rate
                    msGrid.TextMatrix(msGrid.Row, 9) = RsBk!vsubRTPER  '--Rate Per
                    If RsBk!ItType = 0 Then
                       If Left(lblGSTINNo.Caption, 2) = "27" Or Left(lblGSTINNo.Caption, 1) = "U" Or Trim(lblGSTINNo.Caption) = "" Then
                          msGrid.TextMatrix(msGrid.Row, 14) = GProcGetColumnValue("tblMastNarration", "Narration", "G S T", "S", "CotCGSTRt", "N")
                          msGrid.TextMatrix(msGrid.Row, 16) = GProcGetColumnValue("tblMastNarration", "Narration", "G S T", "S", "CotSGSTRt", "N")
                          msGrid.TextMatrix(msGrid.Row, 18) = 0
                       Else
                          msGrid.TextMatrix(msGrid.Row, 14) = 0
                          msGrid.TextMatrix(msGrid.Row, 16) = 0
                          msGrid.TextMatrix(msGrid.Row, 18) = GProcGetColumnValue("tblMastNarration", "Narration", "G S T", "S", "CotIGSTRt", "N")
                       End If
                    Else
                       If Left(lblGSTINNo.Caption, 2) = "27" Or Left(lblGSTINNo.Caption, 1) = "U" Or Trim(lblGSTINNo.Caption) = "" Then
                          msGrid.TextMatrix(msGrid.Row, 14) = GProcGetColumnValue("tblMastNarration", "Narration", "G S T", "S", "PolCGSTRt", "N")
                          msGrid.TextMatrix(msGrid.Row, 16) = GProcGetColumnValue("tblMastNarration", "Narration", "G S T", "S", "PolSGSTRt", "N")
                          msGrid.TextMatrix(msGrid.Row, 18) = 0
                       Else
                          msGrid.TextMatrix(msGrid.Row, 14) = 0
                          msGrid.TextMatrix(msGrid.Row, 16) = 0
                          msGrid.TextMatrix(msGrid.Row, 18) = GProcGetColumnValue("tblMastNarration", "Narration", "G S T", "S", "PolIGSTRt", "N")
                       End If
                    End If
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

