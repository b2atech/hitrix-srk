VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form frmPurchaseOtherGST 
   BackColor       =   &H00C2E3FC&
   BorderStyle     =   0  'None
   Caption         =   "Purchase Entry"
   ClientHeight    =   9600
   ClientLeft      =   0
   ClientTop       =   15
   ClientWidth     =   15105
   FillColor       =   &H000040C0&
   FillStyle       =   0  'Solid
   Icon            =   "frmPurchaseOtherGST.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   9600
   ScaleWidth      =   15105
   ShowInTaskbar   =   0   'False
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
      TabIndex        =   22
      Text            =   "txttemp(28)"
      ToolTipText     =   "Press F3 for Creation"
      Top             =   7440
      Width           =   3420
   End
   Begin VB.TextBox txtVou 
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
      Index           =   9
      Left            =   990
      MaxLength       =   40
      TabIndex        =   30
      Text            =   "txtVou(9)"
      Top             =   8820
      Width           =   7560
   End
   Begin VB.TextBox txtVou 
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
      Index           =   8
      Left            =   990
      MaxLength       =   40
      TabIndex        =   29
      Text            =   "txtVou(8)"
      Top             =   8325
      Width           =   7560
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFC0C0&
      Caption         =   "R C M Bill Printing"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2970
      Left            =   4275
      TabIndex        =   118
      Top             =   2835
      Visible         =   0   'False
      Width           =   5895
      Begin VB.TextBox txtPInvno 
         Appearance      =   0  'Flat
         Height          =   360
         Index           =   0
         Left            =   3240
         MaxLength       =   7
         TabIndex        =   124
         Top             =   630
         Width           =   1695
      End
      Begin VB.CommandButton cmdPrintBill 
         Caption         =   "&Print "
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   0
         Left            =   705
         TabIndex        =   123
         Top             =   1905
         Width           =   1335
      End
      Begin VB.TextBox txtPInvno 
         Appearance      =   0  'Flat
         Height          =   360
         Index           =   1
         Left            =   3240
         MaxLength       =   7
         TabIndex        =   122
         Top             =   1155
         Width           =   1695
      End
      Begin VB.CommandButton cmdPrintBill 
         Caption         =   "&Screen"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   1
         Left            =   2205
         TabIndex        =   121
         Top             =   1890
         Width           =   1335
      End
      Begin VB.CommandButton cmdPrintBill 
         Caption         =   "&Close"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   2
         Left            =   3660
         TabIndex        =   120
         Top             =   1860
         Width           =   1335
      End
      Begin MSMask.MaskEdBox mskDate 
         Height          =   375
         Left            =   3240
         TabIndex        =   119
         Top             =   630
         Visible         =   0   'False
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   661
         _Version        =   393216
         MaxLength       =   10
         Format          =   "dd/mm/yyyy"
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin VB.Label Label31 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "From Invoice No :"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   1200
         TabIndex        =   126
         Top             =   630
         Width           =   1680
      End
      Begin VB.Label Label21 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "To Invoice No     :"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   1200
         TabIndex        =   125
         Top             =   1155
         Width           =   1665
      End
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
      Index           =   43
      Left            =   2250
      TabIndex        =   24
      Text            =   "43"
      Top             =   5400
      Width           =   1695
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
      Left            =   16650
      TabIndex        =   113
      Text            =   "40"
      Top             =   1350
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
      Left            =   16650
      TabIndex        =   112
      Text            =   "39"
      Top             =   945
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
      Left            =   16650
      TabIndex        =   111
      Text            =   "38"
      Top             =   540
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
      Left            =   16560
      TabIndex        =   110
      Text            =   "37"
      Top             =   180
      Visible         =   0   'False
      Width           =   450
   End
   Begin VB.TextBox txtCGSTRcm 
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
      Left            =   6345
      TabIndex        =   26
      Text            =   "txtCGSTRcm"
      Top             =   6435
      Width           =   2085
   End
   Begin VB.TextBox txtSGSTRcm 
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
      Left            =   6345
      TabIndex        =   27
      Text            =   "txtSGSTRcm"
      Top             =   6885
      Width           =   2085
   End
   Begin VB.TextBox txtIGSTRcm 
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
      Left            =   6345
      TabIndex        =   28
      Text            =   "txtIGSTRcm"
      Top             =   7335
      Width           =   2085
   End
   Begin VB.CheckBox chkIsRCM 
      BackColor       =   &H00F5F5DE&
      Caption         =   "Is Reverse Charge Mechanism Applicable ?"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF00FF&
      Height          =   420
      Left            =   2250
      TabIndex        =   3
      Top             =   2430
      Width           =   5415
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
      Left            =   16335
      TabIndex        =   102
      Text            =   "48"
      Top             =   8550
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
      Left            =   17100
      TabIndex        =   101
      Text            =   "47"
      Top             =   8145
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
      Left            =   17055
      TabIndex        =   100
      Text            =   "46"
      Top             =   7740
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
      Left            =   17055
      TabIndex        =   99
      Text            =   "45"
      Top             =   7335
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
      Left            =   9795
      TabIndex        =   8
      Text            =   "44"
      Top             =   2430
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
      Left            =   11160
      TabIndex        =   17
      Text            =   "txtIGSTAmt"
      Top             =   6795
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
      Left            =   11160
      TabIndex        =   16
      Text            =   "txtSGSTAmt"
      Top             =   6345
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
      Left            =   11160
      TabIndex        =   15
      Text            =   "txtCGSTAmt"
      Top             =   5895
      Width           =   2085
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
      Index           =   49
      Left            =   16965
      TabIndex        =   40
      Text            =   "49"
      Top             =   6390
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
      Index           =   50
      Left            =   16965
      TabIndex        =   41
      Text            =   "50"
      Top             =   6885
      Visible         =   0   'False
      Width           =   555
   End
   Begin VB.CommandButton CmdBagDetails 
      Caption         =   "Cartoon &Bag Details"
      Height          =   720
      Left            =   19170
      TabIndex        =   89
      Top             =   1575
      Visible         =   0   'False
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
      TabIndex        =   33
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
      TabIndex        =   88
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
      TabIndex        =   87
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
      TabIndex        =   86
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
      Left            =   19800
      TabIndex        =   85
      Text            =   "txttmpTotWt"
      Top             =   6030
      Visible         =   0   'False
      Width           =   255
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
      Left            =   19755
      TabIndex        =   84
      Text            =   "txttmpTotBag"
      Top             =   5535
      Visible         =   0   'False
      Width           =   225
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
      Left            =   17010
      TabIndex        =   83
      Text            =   "15"
      Top             =   4230
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
      TabIndex        =   50
      Text            =   "txtExpensesAmt"
      Top             =   6210
      Visible         =   0   'False
      Width           =   975
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
      TabIndex        =   82
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
      TabIndex        =   81
      Text            =   "32"
      Top             =   7830
      Visible         =   0   'False
      Width           =   420
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
      Left            =   2235
      TabIndex        =   2
      Text            =   "txttemp(3)"
      ToolTipText     =   "Press F3 for Creation"
      Top             =   1185
      Width           =   5940
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
      Left            =   18675
      TabIndex        =   32
      Text            =   "txtCrBalance"
      Top             =   225
      Visible         =   0   'False
      Width           =   510
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
      Left            =   8460
      TabIndex        =   80
      Text            =   "txtDrBalance"
      Top             =   525
      Width           =   2100
   End
   Begin VB.CommandButton cmdLedger 
      Caption         =   "Ledger"
      Height          =   375
      Left            =   11745
      TabIndex        =   79
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
      TabIndex        =   78
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
      TabIndex        =   49
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
      TabIndex        =   48
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
      TabIndex        =   47
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
      TabIndex        =   46
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
      TabIndex        =   45
      Text            =   "26"
      Top             =   2040
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
      Index           =   25
      Left            =   17340
      TabIndex        =   44
      Text            =   "25"
      Top             =   1650
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
      Index           =   24
      Left            =   16920
      TabIndex        =   37
      Text            =   "24"
      Top             =   5985
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
      TabIndex        =   13
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
      TabIndex        =   76
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
      TabIndex        =   74
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
      TabIndex        =   73
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
      TabIndex        =   72
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
      Left            =   1800
      Style           =   1  'Graphical
      TabIndex        =   65
      Top             =   9450
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
      Left            =   2610
      Style           =   1  'Graphical
      TabIndex        =   64
      ToolTipText     =   "To Add New Record"
      Top             =   9495
      Visible         =   0   'False
      Width           =   1140
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
      Left            =   13890
      Style           =   1  'Graphical
      TabIndex        =   62
      ToolTipText     =   "Exit"
      Top             =   6300
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
      Left            =   13890
      Style           =   1  'Graphical
      TabIndex        =   55
      Top             =   5790
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
      Left            =   13890
      Style           =   1  'Graphical
      TabIndex        =   54
      ToolTipText     =   "To Save Record"
      Top             =   5295
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
      Left            =   13890
      Style           =   1  'Graphical
      TabIndex        =   63
      Top             =   4680
      Width           =   870
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
      Left            =   13890
      Style           =   1  'Graphical
      TabIndex        =   53
      ToolTipText     =   "To Delete Record"
      Top             =   4125
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
      Left            =   13890
      Style           =   1  'Graphical
      TabIndex        =   52
      ToolTipText     =   "To Modify Record"
      Top             =   3660
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
      Left            =   13890
      Style           =   1  'Graphical
      TabIndex        =   51
      ToolTipText     =   "To Add New Record"
      Top             =   3210
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
      TabIndex        =   56
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
      Left            =   13890
      Style           =   1  'Graphical
      TabIndex        =   57
      ToolTipText     =   "To Find Record"
      Top             =   2475
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
      Left            =   14385
      Style           =   1  'Graphical
      TabIndex        =   61
      ToolTipText     =   "Last"
      Top             =   1935
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
      Left            =   13890
      Style           =   1  'Graphical
      TabIndex        =   60
      ToolTipText     =   "First"
      Top             =   1935
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
      Left            =   14385
      Style           =   1  'Graphical
      TabIndex        =   59
      ToolTipText     =   "Next"
      Top             =   1395
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
      Left            =   9795
      TabIndex        =   4
      Text            =   "0"
      Top             =   1230
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
      Left            =   13890
      Style           =   1  'Graphical
      TabIndex        =   58
      ToolTipText     =   "Previous"
      Top             =   1395
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
      Left            =   9795
      TabIndex        =   6
      Text            =   "txtVou(11)"
      Top             =   1800
      Width           =   1785
   End
   Begin VB.TextBox txt 
      Alignment       =   1  'Right Justify
      Height          =   315
      Index           =   4
      Left            =   17895
      TabIndex        =   71
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
      Left            =   2265
      Locked          =   -1  'True
      TabIndex        =   1
      Text            =   "txttemp(4)"
      ToolTipText     =   "Press F3 for Creation"
      Top             =   525
      Width           =   5775
   End
   Begin VB.TextBox txt 
      Height          =   315
      Index           =   6
      Left            =   17895
      TabIndex        =   70
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
      TabIndex        =   69
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
      TabIndex        =   68
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
      Left            =   11160
      TabIndex        =   14
      Text            =   "10"
      Top             =   5445
      Width           =   2085
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
      Left            =   10215
      TabIndex        =   19
      Text            =   "11"
      Top             =   7680
      Width           =   3030
   End
   Begin VB.TextBox txt 
      Alignment       =   1  'Right Justify
      Height          =   315
      Index           =   12
      Left            =   17895
      TabIndex        =   67
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
      TabIndex        =   38
      Text            =   "13"
      Top             =   3330
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
      Index           =   14
      Left            =   16695
      TabIndex        =   39
      Text            =   "14"
      Top             =   3735
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
      Index           =   16
      Left            =   16965
      TabIndex        =   43
      Text            =   "16"
      Top             =   4680
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
      Left            =   11160
      TabIndex        =   18
      Text            =   "17"
      Top             =   7230
      Width           =   2085
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
      Index           =   18
      Left            =   17010
      TabIndex        =   35
      Text            =   "18"
      Top             =   5040
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
      Index           =   19
      Left            =   17010
      TabIndex        =   36
      Text            =   "19"
      Top             =   5445
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
      Index           =   20
      Left            =   2520
      TabIndex        =   23
      Text            =   "20"
      Top             =   6945
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
      Index           =   21
      Left            =   2520
      TabIndex        =   21
      Text            =   "21"
      Top             =   6585
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
      Index           =   22
      Left            =   1170
      TabIndex        =   20
      Text            =   "22"
      Top             =   6585
      Width           =   840
   End
   Begin VB.TextBox txt 
      Alignment       =   1  'Right Justify
      Height          =   315
      Index           =   23
      Left            =   17895
      TabIndex        =   66
      Text            =   "23"
      Top             =   4050
      Visible         =   0   'False
      Width           =   525
   End
   Begin MSMask.MaskEdBox mskDt 
      Height          =   375
      Index           =   2
      Left            =   11730
      TabIndex        =   5
      Top             =   1230
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
      Left            =   11730
      TabIndex        =   7
      Top             =   1800
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
   Begin Crystal.CrystalReport cryReport 
      Bindings        =   "frmPurchaseOtherGST.frx":058A
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
      Left            =   180
      TabIndex        =   12
      Top             =   3195
      Width           =   13335
      _ExtentX        =   23521
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
      Index           =   34
      Left            =   11730
      TabIndex        =   9
      Top             =   2460
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
   Begin MSMask.MaskEdBox mskDt 
      Height          =   375
      Index           =   9
      Left            =   5940
      TabIndex        =   25
      Top             =   5400
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
   Begin VB.Label LblPan 
      BackStyle       =   0  'Transparent
      Caption         =   "PANNO"
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
      Left            =   2250
      TabIndex        =   133
      Top             =   2070
      Width           =   3180
   End
   Begin VB.Label Label28 
      BackStyle       =   0  'Transparent
      Caption         =   "PAN     :"
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
      Left            =   1305
      TabIndex        =   132
      Top             =   2070
      Width           =   945
   End
   Begin VB.Label Label27 
      BackStyle       =   0  'Transparent
      Caption         =   "TDS A/c  : "
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
      Left            =   330
      TabIndex        =   131
      Top             =   7440
      Width           =   1365
   End
   Begin VB.Label Label26 
      BackStyle       =   0  'Transparent
      Caption         =   "TDS J.V. No              :"
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
      Height          =   330
      Left            =   330
      TabIndex        =   130
      Top             =   6990
      Width           =   2070
   End
   Begin VB.Label Label25 
      BackStyle       =   0  'Transparent
      Caption         =   "% :"
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
      Height          =   330
      Left            =   2040
      TabIndex        =   129
      Top             =   6585
      Width           =   405
   End
   Begin VB.Label Label24 
      BackStyle       =   0  'Transparent
      Caption         =   "TDS @ :"
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
      Height          =   330
      Left            =   330
      TabIndex        =   128
      Top             =   6585
      Width           =   810
   End
   Begin VB.Shape Shape3 
      BorderColor     =   &H00000080&
      FillColor       =   &H00F5F5DE&
      FillStyle       =   0  'Solid
      Height          =   990
      Left            =   8775
      Top             =   8325
      Width           =   4740
   End
   Begin VB.Label Label23 
      BackStyle       =   0  'Transparent
      Caption         =   "Note :"
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
      Height          =   285
      Left            =   240
      TabIndex        =   127
      Top             =   8415
      Width           =   1155
   End
   Begin VB.Label lblRcmSubAmt 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "RcmSubAmount"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000040C0&
      Height          =   375
      Left            =   6300
      TabIndex        =   117
      Top             =   6030
      Width           =   2070
   End
   Begin VB.Label Label20 
      BackStyle       =   0  'Transparent
      Caption         =   "Sales A/c"
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
      Height          =   285
      Left            =   330
      TabIndex        =   116
      Top             =   6120
      Width           =   1155
   End
   Begin VB.Label lblRcmSaleAc 
      BackStyle       =   0  'Transparent
      Caption         =   "RcmSalesAc"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000040C0&
      Height          =   375
      Left            =   1530
      TabIndex        =   115
      Top             =   6120
      Width           =   3375
   End
   Begin VB.Label Label15 
      BackStyle       =   0  'Transparent
      Caption         =   "Sub Amt   :"
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
      Left            =   5130
      TabIndex        =   114
      Top             =   6030
      Width           =   1110
   End
   Begin VB.Label lblRCMTotal 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "RCM Total"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF00FF&
      Height          =   375
      Left            =   6435
      TabIndex        =   109
      Top             =   7785
      Width           =   1875
   End
   Begin VB.Label Label19 
      BackStyle       =   0  'Transparent
      Caption         =   "Total         :"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000040C0&
      Height          =   375
      Left            =   5130
      TabIndex        =   108
      Top             =   7785
      Width           =   1170
   End
   Begin VB.Label Label18 
      BackStyle       =   0  'Transparent
      Caption         =   "I  GST        :"
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
      Left            =   5130
      TabIndex        =   107
      Top             =   7425
      Width           =   1215
   End
   Begin VB.Label Label17 
      BackStyle       =   0  'Transparent
      Caption         =   "S GST       :"
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
      Left            =   5130
      TabIndex        =   106
      Top             =   6930
      Width           =   1245
   End
   Begin VB.Label Label16 
      BackStyle       =   0  'Transparent
      Caption         =   "C GST       :"
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
      Left            =   5130
      TabIndex        =   105
      Top             =   6435
      Width           =   1125
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00000080&
      X1              =   180
      X2              =   8595
      Y1              =   5940
      Y2              =   5940
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "RCM Bill Date  :"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000040C0&
      Height          =   375
      Left            =   4365
      TabIndex        =   104
      Top             =   5430
      Width           =   1440
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "RCM Bill No       :"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000040C0&
      Height          =   375
      Left            =   450
      TabIndex        =   103
      Top             =   5445
      Width           =   1665
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
      Left            =   8625
      TabIndex        =   98
      Top             =   2475
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
      Left            =   1305
      TabIndex        =   97
      Top             =   1710
      Width           =   945
   End
   Begin VB.Label lblGSTINNo 
      BackStyle       =   0  'Transparent
      Caption         =   "GSTIN No "
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
      Left            =   2250
      TabIndex        =   96
      Top             =   1710
      Width           =   3180
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
      Left            =   5535
      TabIndex        =   95
      Top             =   1710
      Width           =   1215
   End
   Begin VB.Label lblStateCd 
      BackStyle       =   0  'Transparent
      Caption         =   "State Cd"
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
      Left            =   6750
      TabIndex        =   94
      Top             =   1710
      Width           =   990
   End
   Begin VB.Label Label11 
      BackStyle       =   0  'Transparent
      Caption         =   "R. Off    :"
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
      Left            =   10260
      TabIndex        =   93
      Top             =   7290
      Width           =   900
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
      Left            =   9450
      TabIndex        =   92
      Top             =   5895
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
      Left            =   9450
      TabIndex        =   91
      Top             =   6390
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
      Left            =   9450
      TabIndex        =   90
      Top             =   6900
      Width           =   1665
   End
   Begin VB.Shape Shape5 
      BorderColor     =   &H00C00000&
      BorderWidth     =   5
      FillColor       =   &H00C00000&
      Height          =   9600
      Left            =   0
      Top             =   -90
      Width           =   15105
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "&Party / Credit A/c :"
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
      TabIndex        =   31
      Top             =   1200
      Width           =   1800
   End
   Begin VB.Label Label9 
      BackStyle       =   0  'Transparent
      Caption         =   "Purch/Debit  A/c   :"
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
      Width           =   2430
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
      Left            =   8625
      TabIndex        =   10
      Top             =   1230
      Width           =   990
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
      ForeColor       =   &H00800000&
      Height          =   375
      Left            =   10020
      TabIndex        =   34
      Top             =   5490
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
      ForeColor       =   &H000000FF&
      Height          =   375
      Left            =   9030
      TabIndex        =   42
      Top             =   7755
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
      Left            =   12375
      TabIndex        =   77
      Top             =   0
      Width           =   2475
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
      Left            =   8625
      TabIndex        =   11
      Top             =   1890
      Width           =   1095
   End
   Begin VB.Label Label6 
      BackColor       =   &H00C00000&
      Caption         =   "  Other  Purchase / Expenses"
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
      Left            =   -120
      TabIndex        =   75
      Top             =   0
      Width           =   15285
   End
   Begin VB.Shape Shape1 
      FillColor       =   &H00BDDFD1&
      FillStyle       =   0  'Solid
      Height          =   6435
      Left            =   13680
      Shape           =   4  'Rounded Rectangle
      Top             =   1035
      Width           =   1260
   End
   Begin VB.Shape Shape4 
      BorderColor     =   &H00000080&
      FillColor       =   &H00F5F5DE&
      FillStyle       =   0  'Solid
      Height          =   3000
      Left            =   8775
      Top             =   5265
      Width           =   4725
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H00000080&
      FillColor       =   &H00F5F5DE&
      FillStyle       =   0  'Solid
      Height          =   2025
      Left            =   225
      Top             =   1020
      Width           =   8115
   End
   Begin VB.Shape Shape6 
      BorderColor     =   &H00000080&
      FillColor       =   &H00EDFDFE&
      FillStyle       =   0  'Solid
      Height          =   510
      Left            =   180
      Top             =   450
      Width           =   13335
   End
   Begin VB.Shape Shape7 
      BorderColor     =   &H00000080&
      FillColor       =   &H00EDFDFE&
      FillStyle       =   0  'Solid
      Height          =   2025
      Left            =   8460
      Top             =   1020
      Width           =   5070
   End
   Begin VB.Shape Shape8 
      BorderColor     =   &H00000080&
      FillColor       =   &H00EDFDFE&
      FillStyle       =   0  'Solid
      Height          =   3000
      Left            =   165
      Top             =   5265
      Width           =   8415
   End
End
Attribute VB_Name = "frmPurchaseOtherGST"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Const tblName As String = "tblAddLess"
Const IdField As String = "Vno"
Const OrderField As String = "Vdt,Vno"
Const MaxNo As Long = 40
Const OtherTxtFromNo As Long = 43 '-- Other Txt Ctrl From No
Const OtherTxtToNo As Long = 50 '-- Other Txt Ctrl Less No
Dim rstbl As Recordset '-- TblAddLess
Dim rstblSub As Recordset '-- TblVouSub
Dim varBookMark As Variant
Dim FirstTimeFlag As Boolean
Dim rsVou As Recordset '--TblVoucher
Dim rsOS As Recordset '--TblOutStanding
Dim OldgFrmTypeStr As String
Dim FormAction As Integer
Dim frmTypeStr As String
Dim wTaxable As Double
Dim i As Long

Private Sub chkIsRCM_LostFocus()
GenRCMBillNo
End Sub

'Dim GridFlag As Boolean
Private Sub cmdBtn_Click(Index As Integer)
'On Error GoTo ErrorRoutine
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
    gClsSearch.SearchMultiField "tblAddLess,TblMastAccount", "Vno,AcName,Vdt,AdInvNo,AdGpDt", Array("Vno", "Party", "Vdt", "Bill No", "Bill Dt"), Array(txt(0).Width, 4000, 1080, 1000, 1200), " tblAddLess.AdAcCrCode=TblMastAccount.AcCode and vtype='" & frmTypeStr & "' and VYear=" & gCYear, "", "Vno,Vdt", 0, 0
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
    Call GProcClearForm(Me, rstbl, MaxNo, True, OtherTxtFromNo, OtherTxtToNo)
    If FirstTimeFlag = True Then
        FirstTimeFlag = False
    Else
        txttemp(4).SetFocus
    End If
    txt(0) = GProcGenerateId("TblAddLess", "Vno", "Vtype='" & frmTypeStr & "' and VYear=" & gCYear)
    mskDt(2) = GProcGenerateVDt("tblAddLess", "Vdt", "Vtype='" & frmTypeStr & "' and VYear=" & gCYear)
    lblRCMTotal.Caption = ""
    lblRcmSubAmt.Caption = ""
    
    SetGrid
Case 7 'Modify
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
        End If
    'End If
Case 8 'Delete
    'If mURecDel = True Then
      If IsRelatedRecord = False Then
        If MsgBox("Do You Really Delete This Record?", vbQuestion + vbYesNo + vbDefaultButton1, "Confirm Delete") = vbYes Then
            If rstbl.RecordCount > 1 Then
               '-- Delete Related Record
                DeleteRelatedRecord
                rstbl.Delete
                cmdBtn_Click (1)
            ElseIf rstbl.RecordCount = 1 Then
                '-- Delete Related Record
                DeleteRelatedRecord
                rstbl.Delete
                Call GProcActivateControls(Me, False)
                Call GProcSetButtons(Me, vbDataActionCancel)
                Form_Load
            End If
        End If
    End If
Case 9 'Print
        '-- Inv No
    If IsNumeric(txt(43)) = False Then
        txtPInvno(0) = Mid(txt(43).text, 5)
        txtPInvno(1) = Mid(txt(43).text, 5)
    Else
        txtPInvno(0) = txt(43).text
        txtPInvno(1) = txt(43).text
    End If
    Frame1.Visible = True
    DoEvents
    txtPInvno(0).Enabled = True
    txtPInvno(1).Enabled = True
    txtPInvno(0).SetFocus
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
        If chkIsRCM.Value = 0 And txt(43) <> "" Then
           txt(43) = ""
        End If
        GProcSaveRecord Me, rstbl, FormAction, MaxNo, OtherTxtFromNo, OtherTxtToNo
        SaveRelatedRecord
        ReadFields
        '--Update Stock
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
    If rstbl.RecordCount <= 0 Then
        Call GProcClearForm(Me, rstbl, MaxNo, True, OtherTxtFromNo, OtherTxtToNo)
        Call GProcActivateControls(Me, False)
        Call GProcSetButtons(Me, vbDataActionCancel)
    Else
        If FormAction = vbDataActionAddNew Then
            rstbl.MoveLast
        Else
            rstbl.CancelUpdate
        End If
        Call GProcClearForm(Me, rstbl, MaxNo, True, OtherTxtFromNo, OtherTxtToNo)
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
    rstbl.CancelUpdate
    GProcErrorHandler
End Sub
Private Sub SetControlEd()
txttemp(4).Enabled = True '----Purch A/c
txt(11).Enabled = False '-- Amt
txtCrBalance.Enabled = False
txtDrBalance.Enabled = False
txtCGSTAmt.Enabled = False
txtSGSTAmt.Enabled = False
txtIGSTAmt.Enabled = False
txtCGSTRcm.Enabled = False
txtSGSTRcm.Enabled = False
txtIGSTRcm.Enabled = False
mskDt(9).Enabled = False
txt(20).Enabled = False
txt(21).Enabled = False
txt(10).Enabled = False
End Sub
Private Sub DeleteRelatedRecord()
'--- Voucher Sub
gCn.Execute "delete from tblVouSub where vtype='" & frmTypeStr & "' and vno=" & txt(0) & " and " _
& " VYear=" & gCYear
'-- Voucher
gCn.Execute "delete from tblVoucher where vtype='" & frmTypeStr & "' and vno=" & txt(0) & " and " _
& " VYear=" & gCYear

'-- Voucher
gCn.Execute "delete from tblVoucher where vtype='SZ' and vno=" & txt(0) & " and " _
& " VYear=" & gCYear

'--- JV
gCn.Execute "delete from tblVoucher where vtype='JV' and vno=" & txt(20) & " and " _
& " VYear=" & gCYear
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
Dim TaxAcCodeCGSTRcm As Long
Dim TaxAcCodeSGSTRcm As Long
Dim TaxAcCodeIGSTRcm As Long
Dim RcmSaleAc As Long

Dim X As String
Dim BillDt
Dim DueDt
Dim TaxAmt As Double
Dim TDSAc As Long
'--- Account Code
MillCode = GProcGetColumnValue("tblMastAccount", "AcName", "-", "S", "AcCode", "N")
CrCode = GProcGetColumnValue("tblMastAccount", "AcName", txttemp(3), "S", "AcCode", "N")
DrCode = GProcGetColumnValue("tblMastAccount", "AcName", txttemp(4), "S", "AcCode", "N")

If txttemp(28) = "" Then txttemp(28) = "-"
TDSAc = GProcGetColumnValue("tblMastAccount", "AcName", txttemp(28), "S", "AcCode", "N")


BrkCode = GProcGetColumnValue("tblMastAccount", "AcName", "-", "S", "AcCode", "N")
TaxAcCodeCGST = GProcGetColumnValue("tblMastNarration", "Narration", "G S T", "S", "CGSTInPutAc", "N")
TaxAcCodeSGST = GProcGetColumnValue("tblMastNarration", "Narration", "G S T", "S", "SGSTInPutAc", "N")
TaxAcCodeIGST = GProcGetColumnValue("tblMastNarration", "Narration", "G S T", "S", "IGSTInPutAc", "N")

TaxAcCodeCGSTRcm = GProcGetColumnValue("tblMastNarration", "Narration", "G S T", "S", "CGSTPayAc", "N")
TaxAcCodeSGSTRcm = GProcGetColumnValue("tblMastNarration", "Narration", "G S T", "S", "SGSTPayAc", "N")
TaxAcCodeIGSTRcm = GProcGetColumnValue("tblMastNarration", "Narration", "G S T", "S", "IGSTPayAc", "N")
RcmSaleAc = GProcGetColumnValue("tblMastNarration", "Narration", "G S T", "S", "RCMPayAc", "N")


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
        '--Itcode(4)
        rsVouSub!vsubITCODE = GProcGetColumnValue("tblMastItem", "ItName", .TextMatrix(i, 1), "S", "ItCode", "N", " ItMillCode=" & MillCode) '--Item Code
        rsVouSub!vsubwt = CDbl(.TextMatrix(i, 2)) '-- Wt(6)
        rsVouSub!vsubRt = CDbl(.TextMatrix(i, 3)) '--Rate(7)
        rsVouSub!vsubRTPER = 1 '--Rate Per(8)
        rsVouSub!vsubAMT = CDbl(.TextMatrix(i, 4)) '--Amt(9)
        '--Godown Code(17)
        rsVouSub!VSubGodownCode = GProcGetColumnValue("tblMastNarration", "Narration", "-", "S", "NarrCode", "N", "")
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
        rsVouSub!VSUBTMP1 = Val(.TextMatrix(i, 5)) '- Add
        rsVouSub!VsubLess = Val(.TextMatrix(i, 6)) '- Less
        rsVouSub!VSubTaxableAmt = Val(.TextMatrix(i, 7))
        rsVouSub!VsubCGSTRt = Val(.TextMatrix(i, 8))
        rsVouSub!VsubCGSTAmt = Val(.TextMatrix(i, 9))
        rsVouSub!VsubSGSTRt = Val(.TextMatrix(i, 10))
        rsVouSub!VsubSGSTAmt = Val(.TextMatrix(i, 11))
        rsVouSub!VsubIGSTRt = Val(.TextMatrix(i, 12))
        rsVouSub!VsubIGSTAmt = Val(.TextMatrix(i, 13))
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
TaxAmt = Val(txtCGSTAmt) + Val(txtSGSTAmt) + Val(txtIGSTAmt)
GProcCreateVoucher FormAction, txt(0), 1, txt(2), txt(1), CrCode, DrCode, (txt(11)), txtVou(11), BillDt, "", "Bill No " & txtVou(11) & " Dtd. " & BillDt, txtVou(8), txtVou(9)
GProcCreateVoucher FormAction, txt(0), 2, txt(2), txt(1), DrCode, CrCode, -(txt(11) - TaxAmt), txtVou(11), BillDt, "", "Bill No " & txtVou(11) & " Dtd. " & BillDt, txtVou(8), txtVou(9)
i = 3
If Val(txtCGSTAmt) > 0 Then
   GProcCreateVoucher FormAction, txt(0), i, txt(2), txt(1), TaxAcCodeCGST, DrCode, Val(txtCGSTAmt) * -1, txtVou(11), BillDt, "", "Bill No " & txtVou(11) & " Dtd. " & BillDt
   i = i + 1
   DoEvents
End If
If Val(txtSGSTAmt) > 0 Then
   GProcCreateVoucher FormAction, txt(0), i, txt(2), txt(1), TaxAcCodeSGST, DrCode, Val(txtSGSTAmt) * -1, txtVou(11), BillDt, "", "Bill No " & txtVou(11) & " Dtd. " & BillDt
   i = i + 1
   DoEvents
End If
If Val(txtIGSTAmt) > 0 Then
   GProcCreateVoucher FormAction, txt(0), i, txt(2), txt(1), TaxAcCodeIGST, DrCode, Val(txtIGSTAmt) * -1, txtVou(11), BillDt, "", "Bill No " & txtVou(11) & " Dtd. " & BillDt
   i = i + 1
   DoEvents
End If
'---- For RCM

If Val(txtCGSTRcm) + Val(txtSGSTRcm) + Val(txtIGSTRcm) > 0 Then
    TaxAmt = Val(txtCGSTRcm) + Val(txtSGSTRcm) + Val(txtIGSTRcm)
    i = 3
    If Val(txtCGSTRcm) > 0 Then
       i = i + 1
       GProcCreateVoucher FormAction, txt(0), i, txt(2), txt(1), TaxAcCodeCGST, DrCode, Val(txtCGSTRcm) * -1, txtVou(11), BillDt, "", txt(43) + " Dtd.-" & mskDt(9), "Bill No " & txtVou(11) & " Dtd. " & BillDt
       i = i + 1
       GProcCreateVoucher FormAction, txt(0), i, txt(2), txt(1), TaxAcCodeCGSTRcm, DrCode, Val(txtCGSTRcm), txtVou(11), BillDt, "", txt(43) + " Dtd.-" & mskDt(9), "Bill No " & txtVou(11) & " Dtd. " & BillDt
    End If
    If Val(txtSGSTRcm) > 0 Then
       i = i + 1
       GProcCreateVoucher FormAction, txt(0), i, txt(2), txt(1), TaxAcCodeSGST, DrCode, Val(txtSGSTRcm) * -1, txtVou(11), BillDt, "", txt(43) + " Dtd.-" & mskDt(9), "Bill No " & txtVou(11) & " Dtd. " & BillDt
       i = i + 1
       GProcCreateVoucher FormAction, txt(0), i, txt(2), txt(1), TaxAcCodeSGSTRcm, DrCode, Val(txtSGSTRcm), txtVou(11), BillDt, "", txt(43) + " Dtd.-" & mskDt(9), "Bill No " & txtVou(11) & " Dtd. " & BillDt
    End If
    If Val(txtIGSTRcm) > 0 Then
       i = i + 1
       GProcCreateVoucher FormAction, txt(0), i, txt(2), txt(1), TaxAcCodeIGST, DrCode, Val(txtIGSTRcm) * -1, txtVou(11), BillDt, "", txt(43) + " Dtd.-" & mskDt(9), "Bill No " & txtVou(11) & " Dtd. " & BillDt
       i = i + 1
       GProcCreateVoucher FormAction, txt(0), i, txt(2), txt(1), TaxAcCodeIGSTRcm, DrCode, Val(txtIGSTRcm), txtVou(11), BillDt, "", txt(43) + " Dtd.-" & mskDt(9), "Bill No " & txtVou(11) & " Dtd. " & BillDt
    End If
End If

'-- Debit Amt
gCn.Execute "delete from tblVoucher where Vtype='JV' and Vyear=" & gCYear & " and vno=" & txt(20)
If CDbl(txt(21)) > 0 Then
    If txt(20) = 0 Then txt(20) = GProcGenerateId("tblVoucher", "Vno", "Vtype='JV' and Vyear=" & gCYear)
    GProcCreateVoucher FormAction, txt(20), 1, mskDt(2), "JV", TDSAc, CrCode, txt(21), txtVou(11), mskVou(12), "", "T D S on Amount " + CStr(wTaxable), "@ " + txt(22) + " %", "Agt. Bill " + txtVou(11) + " Dtd-" + CStr(mskVou(12)), "", True
    GProcCreateVoucher FormAction, txt(20), 2, mskDt(2), "JV", CrCode, TDSAc, -txt(21), txtVou(11), mskVou(12), "", "T D S on Amount " + CStr(wTaxable), "@ " + txt(22) + " %", "Agt. Bill " + txtVou(11) + " Dtd-" + CStr(mskVou(12)), "", True
    gCn.Execute "update tblAddLess set AdDbNtNo= " & Val(txt(20)) & " where Vno=" & CLng(txt(0)) & " and Vtype='" & frmTypeStr & "' and Vyear=" & gCYear
    rstbl.Requery
    rstbl.Find "Vno = " & Val(txt(0))
Else
    txt(20) = 0
    gCn.Execute "update tblAddLess set AdDbNtNo=0 where Vno=" & CLng(txt(0)) & " and Vtype='" & frmTypeStr & "' and Vyear=" & gCYear
    rstbl.Requery
    rstbl.Find "Vno = " & Val(txt(0))
End If

'-- End Rcm
End Sub
Private Sub cmdLedger_Click()
'On Error GoTo ErrorHandler
frmRptAccount.txt(1) = txttemp(4)
frmMain.mnuRptRSubAccLeg_Click (0)
Exit Sub
ErrorHandler:
    GProcErrorHandler
End Sub

Private Sub cmdPrintBill_Click(Index As Integer)
Dim RsSA As Recordset
Dim X As String
Dim BillSrNo As String
Dim InvNo As Variant
Dim DbNtAcName As String
Dim Y As String
Dim Lrd As Recordset
Dim LrString As String
Dim LrStringDt As String
Dim BilNo As String
LrString = ""
Dim RsBl As Recordset
Select Case Index
Case 0, 1:
    '------ Validate
    If IsNumeric(txtPInvno(0)) = False Then txtPInvno(0) = 0
    If IsNumeric(txtPInvno(1)) = False Then txtPInvno(1) = txtPInvno(0)
    If CLng(txtPInvno(1)) < CLng(txtPInvno(0)) Then txtPInvno(1) = txtPInvno(0)
    '-------
    If Index = 0 Then GProcCrystalRptPreparation cryReport, "Printer"
    If Index = 1 Then GProcCrystalRptPreparation cryReport, "Window"
    cryReport.WindowTitle = "Invoice"
    '-- A/C Bill SrNo
    InvNo = txtPInvno(0)
    BillSrNo = "RCM"
    X = "("
    Do While CDbl(InvNo) <= CDbl(txtPInvno(1))
       If BillSrNo = "" Then
          X = X + "'" & InvNo & "',"
       Else
          X = X + "'" & BillSrNo & "-" & InvNo & "',"
       End If
       InvNo = InvNo + 1
    Loop
    X = Left(X, Len(X) - 1) + ")"
    X = "Select * From tblAddLess where  ADTMP5 in " & X & " and VType='" & frmTypeStr & "' and VYear=" & gCYear
    X = X & " order by vno"
     If GProcRstOpen(RsSA, X, "R") <= 0 Then
       MsgBox "No Records..."
       Exit Sub
     End If
     RsSA.MoveFirst
     Do While Not RsSA.EOF
        With cryReport
             .Formulas(10) = "wCSTNo = '" & gCCST & "'"
             .Formulas(11) = "wBSTNo = '" & gCBST & "'"
             .Formulas(12) = "wAmountInWords = '" & GProcAmountInWords(RsSA!AdBillAmt) & "'"
             .Formulas(13) = "wType = '" & RsSA!Vtype & "'"
             .Formulas(14) = "wpan = '" & gCPAN & "'"
             .Formulas(15) = "wCompGstin = '" & gCGSTIN & "'"
             .SelectionFormula = "{tblAddLess.ADTMP5} = '" & RsSA!ADTMP5 & "' and {tblAddLess.VType}='" & RsSA!Vtype & "' and {tblAddLess.VYear}=" & RsSA!VYear
             .ReportFileName = gReportPath & "rptSaleBillPrintRCM.rpt"
             .Action = 1
        End With
        RsSA.MoveNext
    Loop
Case 2:
    DoEvents
    Frame1.Visible = False
End Select

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
    Dim RcmTmp As Long
    Me.Top = 300
    Me.Left = Screen.Width / 2 - Me.Width / 2
    '--- TblAddLess
    frmTypeStr = gfrmTypeStr
    X = "Select * from " & tblName & " where Vtype='" & frmTypeStr & "' and VYear=" & gCYear
    If Trim(OrderField) <> "" Then X = X & " order by " & OrderField
    GProcRstOpen rstbl, X, "O"
    FirstTimeFlag = True

    RcmTmp = GProcGetColumnValue("tblMastNarration", "Narration", "G S T", "S", "RCMPayAc", "N")
    
    lblRcmSaleAc.Caption = GProcGetColumnValue("tblMastAccount", "AcCode", "" & RcmTmp & "", "N", "AcName", "S")
    
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
End Sub
Private Sub SetGrid()
Dim X As String
Set rstblSub = New Recordset
X = "select VSubCtrNo,ItName,VSubWt" _
& " ,VSubRt,VSubAmt,VSUBTMP1,VsubLess,VSubTaxableAmt,VsubCGSTRt,VsubCGSTAmt,VsubSGSTRt,VsubSGSTAmt,VsubIGSTRt,VsubIGSTAmt,VSubItCtrlNo,VSubIssBag,VSubIssWt,VSubTmpBag,VSubTmpWt,VSubPItCtrlNo from tblvouSub,TblMastItem where " _
& " tblvouSub.VsubItCode=TblMastItem.ItCode " _
& " and tblvouSub.vtype='" & frmTypeStr & "' and tblvouSub.vno=" & txt(0) & " and tblvouSub.VYear=" & gCYear & " order by VSubCtrNo"
GProcRstOpen rstblSub, X, "R"
With msGrid
   .FormatString = ">Sr|<Item Description                        |>Nos        |>Rate           |>Amount                |>Add      |>Less      |>Taxable Amount |>C Rt    |>CGST Amount|>S Rt    |>SGST Amount|>I Rt    |>IGST Amount||||||"
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
    DrAcBal = GProcGetAccountBalance(txttemp(4), mskDt(2))
    CrAcBal = GProcGetAccountBalance(txttemp(3), mskDt(2))
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
mskVou(12) = mskDt(2)
If Index = 2 Then ShowAccountBalance
If Index = 34 Then
   msGrid.Col = 1
   msGrid_EnterCell
End If
End Sub
Private Sub txt_GotFocus(Index As Integer)
GProcSelectBox Me.ActiveControl
End Sub
Private Sub txtVou_GotFocus(Index As Integer)
GProcSelectBox Me.ActiveControl
End Sub
Private Sub txtVou_LostFocus(Index As Integer)
If Index = 8 Or Index = 9 Then
   txtVou(Index) = GProcProperCase(txtVou(Index))
End If

SaveBtnEd
End Sub
Private Sub txtVou_KeyPress(Index As Integer, KeyAscii As Integer)
If KeyAscii = 13 Then Exit Sub
KeyAscii = GProcValidateKey(rsVou, Index, KeyAscii, txtVou(Index).text)
End Sub
Private Sub CheckForNumber()
GProcCheckForNumber Me, rstbl, MaxNo, Array(43), pOtherTxtFromNo:=OtherTxtFromNo, pOtherTxtToNo:=OtherTxtToNo
End Sub
Private Sub txt_LostFocus(Index As Integer)
GenRCMBillNo
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
        Case 4 '-- Purchase A/c
            GName = GProcGetColumnValue("TblMastGroup", "AgCode", 90025, "N", "AgName", "S")
            GProcShowForm frmMastAccount, frmMain.mnuMstfrm(0), Array(frmMastAccount.txttemp(2)), Array(GName)
        Case 5 '-- millcode
            GName = GProcGetColumnValue("TblMastGroup", "AgCode", 90029, "N", "AgName", "S")
            GProcShowForm frmMastAccount, frmMain.mnuMstfrm(0), Array(frmMastAccount.txttemp(2)), Array(GName)
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
        gClsSearch.SearchMultiField "QryMastaccount", "Account,AcName,AcCity,AgName", Array("Party", "AcName", "City", "Group"), Array(6000, 0, 2000, 3000), "", Chr(KeyAscii), "acName", LeftPos, TopPos, True
        KeyAscii = 0
        If gClsSearch.SearchMultiRetCol(0) <> "" And gClsSearch.SearchMultiRetCol(1) = "" Then
            txttemp(Index) = gClsSearch.SearchMultiRetCol(0)
        Else
            If gClsSearch.SearchMultiRetCol(1) <> "" Then txttemp(Index) = gClsSearch.SearchMultiRetCol(1)
        End If
        lblGSTINNo.Caption = GProcGetColumnValue("tblMastAccount", "AcName", txttemp(3), "S", "AcGSTIN", "S")
        lblStateCd = GProcGetColumnValue("tblMastAccount", "AcName", txttemp(3), "S", "AcStateCode", "S")
        LblPan.Caption = GProcGetColumnValue("tblMastAccount", "AcName", txttemp(3), "S", "AcPan", "S")
        AgCode = GProcGetColumnValue("tblMastAccount", "AcName", txttemp(3), "S", "AgCode", "N")
        txt(22) = GProcGetColumnValue("tblMastAccount", "AcName", txttemp(3), "S", "AcIntPer", "N")
    Case 4 '-- Purchase A/c  Alias Account
        gClsSearch.SearchMultiField "QryMastaccount", "Account,AcName,AcCity,AgName", Array("Account Name", "AcName", "City", "Group"), Array(3000, 0, 2000, 3000), "", Chr(KeyAscii), "AcName", LeftPos, TopPos, True
        KeyAscii = 0
        If gClsSearch.SearchMultiRetCol(0) <> "" And gClsSearch.SearchMultiRetCol(1) = "" Then
            txttemp(Index) = gClsSearch.SearchMultiRetCol(0)
        Else
            If gClsSearch.SearchMultiRetCol(1) <> "" Then txttemp(Index) = gClsSearch.SearchMultiRetCol(1)
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
Private Sub txtTemp_LostFocus(Index As Integer)
If Left(txttemp(4), 7) = "Freight" Then
   chkIsRCM.Value = 1
Else
   chkIsRCM.Value = 0
End If

'If txttemp(4).text = "Freight" Then chkIsRCM.Value = 1
  
SaveBtnEd
ShowAccountBalance
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
        With msGrid
            Select Case .Col
                Case 0: '---- Sr No
                    txtGrid.Locked = True
                    ReNumberGridSrno
                    txtGrid = ""
                    .Col = .Col + 1
                Case Else
                    txtGrid.Locked = False
            End Select
        End With
End If
msGrid_KeyPress (0)
End Sub
Private Sub msGrid_KeyPress(KeyAscii As Integer)
With msGrid
    If KeyAscii = 27 Then '----Esc Key
        KeyAscii = 0
        If (Val(.TextMatrix(.Row, 4)) = 0) And .Rows > 2 Then
            .RemoveItem (.Row)
            txtGrid.text = ""
        End If
        txtGrid.Visible = False
        txt(17).SetFocus
        Exit Sub
    End If
End With
'If FormAction <> vbDataActionAddNew And FormAction <> vbDataActionUpdate Then Exit Sub
'---- Navigate Row ,Col
If KeyAscii = 13 Then
    KeyAscii = 0
    With msGrid
        If .Col = 14 Then 'rstblSub.Fields.Count - 2 Then
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
            If msGrid.Col = 1 Then   '-- Count ,godown
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
MillCode = GProcGetColumnValue("TblMastAccount", "AcName", "-", "S", "Accode", "N")
If msGrid.Row = 0 Then Exit Sub
'-- Wt=Bag * Std Pack
msGrid.TextMatrix(msGrid.Row, 4) = CDbl(msGrid.TextMatrix(msGrid.Row, 2)) * CDbl(msGrid.TextMatrix(msGrid.Row, 3))
msGrid.TextMatrix(msGrid.Row, 4) = Format(msGrid.TextMatrix(msGrid.Row, 4), FStr)

msGrid.TextMatrix(msGrid.Row, 7) = Val(msGrid.TextMatrix(msGrid.Row, 4)) + Val(msGrid.TextMatrix(msGrid.Row, 5)) - Val(msGrid.TextMatrix(msGrid.Row, 6))
msGrid.TextMatrix(msGrid.Row, 7) = Format(msGrid.TextMatrix(msGrid.Row, 7), FStr)

'--CGST
If msGrid.TextMatrix(msGrid.Row, 9) = 0 Then
    msGrid.TextMatrix(msGrid.Row, 9) = (Val(msGrid.TextMatrix(msGrid.Row, 7)) * Val(msGrid.TextMatrix(msGrid.Row, 8))) / 100
    msGrid.TextMatrix(msGrid.Row, 9) = GProcMakeRounding(msGrid.TextMatrix(msGrid.Row, 9))
    msGrid.TextMatrix(msGrid.Row, 9) = Format(msGrid.TextMatrix(msGrid.Row, 9), FStr)
End If
'--SGST
If msGrid.TextMatrix(msGrid.Row, 11) = 0 Then
    msGrid.TextMatrix(msGrid.Row, 11) = (Val(msGrid.TextMatrix(msGrid.Row, 7)) * Val(msGrid.TextMatrix(msGrid.Row, 10))) / 100
    msGrid.TextMatrix(msGrid.Row, 11) = GProcMakeRounding(msGrid.TextMatrix(msGrid.Row, 11))
    msGrid.TextMatrix(msGrid.Row, 11) = Format(msGrid.TextMatrix(msGrid.Row, 11), FStr)
End If
'--IGST
If msGrid.TextMatrix(msGrid.Row, 13) = 0 Then
    msGrid.TextMatrix(msGrid.Row, 13) = (Val(msGrid.TextMatrix(msGrid.Row, 7)) * Val(msGrid.TextMatrix(msGrid.Row, 12))) / 100
    msGrid.TextMatrix(msGrid.Row, 13) = GProcMakeRounding(msGrid.TextMatrix(msGrid.Row, 13))
    msGrid.TextMatrix(msGrid.Row, 13) = Format(msGrid.TextMatrix(msGrid.Row, 13), FStr)
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
'--- Delete Row On Ctrl+Y
If KeyCode = vbKeyY And Shift = 2 And msGrid.Rows > 2 Then
    msGrid = 1
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
    Case 1 '-- Item
        GProcShowForm frmMastItem, frmMain.mnuMstfrm(10), Array(frmMastItemOther.txttemp(9)), Array("-")
        Exit Sub
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
    If KeyAscii <> 13 And KeyAscii <> 27 And KeyAscii <> 25 Then
        '--- Search
        Select Case .Col
            Case 1 '-- Count
                 MillCode = GProcGetColumnValue("TblMastAccount", "AcName", "-", "S", "Accode", "N")
                 gClsSearch.SearchMultiField "tblMastItem,tblMastaccount", "ItName", Array("Item Name"), Array(5000), "tblMastItem.ITMillCode=tblMastaccount.acCode and ITMillCode=" & MillCode, Chr(KeyAscii), "ItName", LeftPos, TopPos, True
                    txtGrid = ""
                    KeyAscii = 0
                    
                    If gClsSearch.SearchMultiRetCol(0) <> "" Then '-- Item Wt ,Rate From Master
                        txtGrid = gClsSearch.SearchMultiRetCol(0)
                        If IsNumeric(.TextMatrix(.Row, 7)) = False Then .TextMatrix(.Row, 7) = 0
                        If IsNumeric(.TextMatrix(.Row, 9)) = False Then .TextMatrix(.Row, 9) = 0
                        If Left(lblGSTINNo.Caption, 2) = "27" Or Left(lblGSTINNo.Caption, 1) = "U" Or Trim(lblGSTINNo.Caption) = "" Then
                           .TextMatrix(.Row, 8) = GProcGetColumnValue("tblMastItem", "Itname", "" & gClsSearch.SearchMultiRetCol(0) & "", "S", "ItCGSTRt", "N")
                           .TextMatrix(.Row, 10) = GProcGetColumnValue("tblMastItem", "Itname", "" & gClsSearch.SearchMultiRetCol(0) & "", "S", "ItSGSTRt", "N")
                           .TextMatrix(.Row, 12) = 0
                        Else
                           .TextMatrix(.Row, 8) = 0
                           .TextMatrix(.Row, 10) = 0
                           .TextMatrix(.Row, 12) = GProcGetColumnValue("tblMastItem", "Itname", "" & gClsSearch.SearchMultiRetCol(0) & "", "S", "ItIGSTRt", "N")
                        End If
                         msGrid.SetFocus
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
        If (Val(.TextMatrix(.Row, 4)) = 0) And .Rows > 2 Then
            .RemoveItem (.Row)
            txtGrid.text = ""
        End If
        txtGrid.Visible = False
        txt(17).Enabled = True
        txt(17).SetFocus
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
                Case 1 '-- Count
                 '-Rate=0 ,count,wt=0 then delete row
                    If RowIndex > 1 And Trim(.TextMatrix(RowIndex, 1)) = "" And CDbl(.TextMatrix(RowIndex, 4)) = 0 Then
                        .RemoveItem (RowIndex)
                        GoTo Validate
                    End If
                    If Trim(.TextMatrix(RowIndex, i)) = "" Then
                        MsgBox "Check Item Name.", vbCritical + vbOKOnly, Me.Caption
                        .Col = i
                        msGrid.SetFocus
                        Exit Function
                    End If
                    MillCode = GProcGetColumnValue("TblMastAccount", "AcName", "-", "S", "Accode", "N")
                    If GProcGetColumnValue("TblMastItem", "ItName", .TextMatrix(RowIndex, i), "S", "Itcode", "N", "ITMillCode=" & MillCode) = 0 Then
                        MsgBox "Check Count.Mill and Count are Mismatch.", vbCritical + vbOKOnly, Me.Caption
                        .Col = i
                        msGrid.SetFocus
                        Exit Function
                    End If
                Case 2 '-- Wt
                    If CDbl(.TextMatrix(RowIndex, i)) = 0 Then
                        MsgBox "Check Qty.", vbCritical + vbOKOnly, Me.Caption
                        .Col = i
                        msGrid.SetFocus
                        Exit Function
                    End If
                Case 4 '-- amount
                    If CDbl(.TextMatrix(RowIndex, i)) = 0 Then
                        MsgBox "Check amount.", vbCritical + vbOKOnly, Me.Caption
                        .Col = i
                        msGrid.SetFocus
                        Exit Function
                    End If
                Case 9 '-- CGST Amt
                     If Val(.TextMatrix(RowIndex, 9)) > 0 Then
                        If Abs(Round((Val(.TextMatrix(RowIndex, 7)) * Val(.TextMatrix(RowIndex, 8))) / 100, 2) - Val(.TextMatrix(RowIndex, 9))) > 1 Then
                           MsgBox "Check CGST Amount", vbCritical + vbOKOnly, Me.Caption
                           .Col = 9
                           msGrid.SetFocus
                           Exit Function
                        End If
                     End If
                Case 11 '-- SGST Amt
                     If Val(.TextMatrix(RowIndex, 11)) > 0 Then
                        If Abs(Round((Val(.TextMatrix(RowIndex, 7)) * Val(.TextMatrix(RowIndex, 10))) / 100, 2) - Val(.TextMatrix(RowIndex, 11))) > 1 Then
                           MsgBox "Check SGST Amount", vbCritical + vbOKOnly, Me.Caption
                           .Col = 11
                           msGrid.SetFocus
                           Exit Function
                        End If
                     End If
                Case 13 '-- IGST Amt
                     If Val(.TextMatrix(RowIndex, 13)) > 0 Then
                        If Abs(Round((Val(.TextMatrix(RowIndex, 7)) * Val(.TextMatrix(RowIndex, 12))) / 100, 2) - Val(.TextMatrix(RowIndex, 13))) > 1 Then
                           MsgBox "Check SGST Amount", vbCritical + vbOKOnly, Me.Caption
                           .Col = 13
                           msGrid.SetFocus
                           Exit Function
                        End If
                     End If
           End Select
        Next i
      Next RowIndex
End With
ValidateGrid = True
ReNumberGridSrno
End Function
Private Sub SaveBtnEd()
'--- Vno,vdt,Party,A/c,Mill,Amount
If CDbl(txt(0)) = 0 Or GProcIsDateValid(mskDt(2)) = False Or txttemp(3) = "" Or txttemp(4) = "" Or CDbl(txt(11)) = 0 Then
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
'    If mskVou(12) < gCYSDate Or mskVou(12) > gCYEDate Then
'        MsgBox "Check Bill Date.", vbInformation + vbOKOnly, Me.Caption
'        mskDt(2).SetFocus
'        Exit Function
'    End If
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
    
    ExpAcCode = GProcGetColumnValue("tblMastNarration", "Narration", "G S T", "S", "RCMPayAc", "N")
    If ExpAcCode = 0 And (Val(txtCGSTRcm) + Val(txtSGSTRcm) + Val(txtIGSTRcm)) > 0 Then
        MsgBox "RCM Sales A/c Not Found.", vbInformation + vbOKOnly, Me.Caption
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
    
    If Val(txt(21)) > 0 And txttemp(28) = "" Then
        MsgBox "Check TDS A/c.", vbInformation + vbOKOnly, Me.Caption
        txttemp(28).SetFocus
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
If IsDate(mskDt(9)) Then '--RCM Dt
    txt(9) = CDate(mskDt(9))
Else
    txt(9) = ""
End If
If IsDate(mskDt(34)) Then '--ern Dt
    txt(34) = CDate(mskDt(34))
Else
    txt(34) = ""
End If
'--- Bill No ,dt
txt(7) = mskVou(12) '--gpDt

txt(36) = txtVou(11)
txt(40) = chkIsRCM.Value
'-- Cr Code
txt(3) = GProcGetColumnValue("tblMastAccount", "AcName", txttemp(3), "S", "AcCode", "N")
'-- Dr Code
txt(4) = GProcGetColumnValue("tblMastAccount", "AcName", txttemp(4), "S", "AcCode", "N")
txt(5) = GProcGetColumnValue("tblMastAccount", "AcName", "-", "S", "AcCode", "N")
txt(23) = gCYear
If txttemp(28) = "" Then txttemp(28) = "-"
txt(28) = GProcGetColumnValue("tblMastAccount", "AcName", txttemp(28), "S", "AcCode", "N")

ShowAccountBalance
End Sub
Private Sub FillTempFromTxt()
Dim X As String
If IsDate(txt(2)) Then '--VDt
    mskDt(2) = txt(2)
Else
    mskDt(2) = "__/__/____"
End If
If IsDate(txt(34)) Then '--ERN Dt
    mskDt(34) = txt(34)
Else
    mskDt(34) = "__/__/____"
End If
If IsDate(txt(9)) Then '--RCM Dt
    mskDt(9) = txt(9)
Else
    mskDt(9) = "__/__/____"
End If

'-- Cr Code
txttemp(3) = GProcGetColumnValue("tblMastAccount", "AcCode", txt(3), "N", "AcName", "S")
'-- Dr Code
txttemp(4) = GProcGetColumnValue("tblMastAccount", "AcCode", txt(4), "N", "AcName", "S")
lblGSTINNo.Caption = GProcGetColumnValue("tblMastAccount", "AcName", txttemp(3), "S", "AcGSTIN", "S")
lblStateCd = GProcGetColumnValue("tblMastAccount", "AcName", txttemp(3), "S", "AcStateCode", "S")
LblPan.Caption = GProcGetColumnValue("tblMastAccount", "AcName", txttemp(3), "S", "AcPan", "S")
chkIsRCM.Value = Val(txt(40))

txttemp(28) = GProcGetColumnValue("tblMastAccount", "AcCode", txt(28), "N", "AcName", "S")
'-- Fill Related Record
FillRelatedRecord
'-- Expenses
ShowAccountBalance
End Sub
Private Sub FillRelatedRecord()
Dim RowIndex As Long
Dim ColIndex As Long
Dim Camt As Double
Dim Samt As Double
Dim Iamt As Double
Camt = 0
Samt = 0
Iamt = 0
txtCGSTRcm = Format(Camt, "#########0.00")
txtSGSTRcm = Format(Samt, "#########0.00")
txtIGSTRcm = Format(Iamt, "#########0.00")
txtCGSTAmt = Format(Camt, "#########0.00")
txtSGSTAmt = Format(Samt, "#########0.00")
txtIGSTAmt = Format(Iamt, "#########0.00")
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
                If ColIndex = 9 Then Camt = Camt + Val(msGrid.TextMatrix(RowIndex, 9))
                If ColIndex = 11 Then Samt = Samt + Val(msGrid.TextMatrix(RowIndex, 11))
                If ColIndex = 13 Then Iamt = Iamt + Val(msGrid.TextMatrix(RowIndex, 13))
            Next ColIndex
            RowIndex = RowIndex + 1
            .MoveNext
            If Not .EOF Then msGrid.Rows = msGrid.Rows + 1
        Loop
    End If
End With
If chkIsRCM.Value = 1 Then
    txtCGSTRcm = Format(Camt, "#########0.00")
    txtSGSTRcm = Format(Samt, "#########0.00")
    txtIGSTRcm = Format(Iamt, "#########0.00")
Else
    txtCGSTAmt = Format(Camt, "#########0.00")
    txtSGSTAmt = Format(Samt, "#########0.00")
    txtIGSTAmt = Format(Iamt, "#########0.00")
End If
lblRcmSubAmt.Caption = Format((Val(txt(10))), FStr)
lblRCMTotal.Caption = Format((Val(txt(10)) + (Val(txtCGSTRcm) + Val(txtSGSTRcm) + Val(txtIGSTRcm))), FStr)

'--- Fill Related Control
SetRelatedRS
txtVou(11) = rsVou.Fields("vBillNo") & vbNullString     '--Bill No
If IsDate(rsVou.Fields("vBillDt")) Then '--Bill Date
    mskVou(12) = Format(rsVou.Fields("vBilldt"), "dd/mm/yyyy")
Else
    mskVou(12) = "__/__/____"
End If
txtVou(8) = rsVou.Fields("VNar2") & vbNullString     '--Bill No
txtVou(9) = rsVou.Fields("VNar3") & vbNullString     '--Bill No
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
Dim MillAcCode As Long
Dim Camt As Double
Dim Samt As Double
Dim Iamt As Double
Dim RcmBlTot As Double
Dim Tds As Double
'---- Sub amt
CheckForNumberGrid
CheckForNumber
Camt = 0
Samt = 0
Iamt = 0
wTaxable = 0
txtCGSTRcm = Format(Camt, "#########0.00")
txtSGSTRcm = Format(Samt, "#########0.00")
txtIGSTRcm = Format(Iamt, "#########0.00")
txtCGSTAmt = Format(Camt, "#########0.00")
txtSGSTAmt = Format(Samt, "#########0.00")
txtIGSTAmt = Format(Iamt, "#########0.00")
RowIndex = 1
With msGrid
    Do While RowIndex < .Rows
        '-- sub amt
        If Val(.TextMatrix(RowIndex, 7)) > 0 Then
            SubAmt = SubAmt + Val(.TextMatrix(RowIndex, 7))
        End If
        If Val(.TextMatrix(RowIndex, 7)) > 0 Then
            wTaxable = wTaxable + Val(.TextMatrix(RowIndex, 7))
        End If
        Camt = Camt + Val(.TextMatrix(RowIndex, 9))
        Samt = Samt + Val(.TextMatrix(RowIndex, 11))
        Iamt = Iamt + Val(.TextMatrix(RowIndex, 13))
        RowIndex = RowIndex + 1
    Loop
End With
If chkIsRCM.Value = 1 Then
    txtCGSTRcm = Format(Camt, "#########0.00")
    txtSGSTRcm = Format(Samt, "#########0.00")
    txtIGSTRcm = Format(Iamt, "#########0.00")
    mskDt(9) = mskDt(2)
Else
    txtCGSTAmt = Format(Camt, "#########0.00")
    txtSGSTAmt = Format(Samt, "#########0.00")
    txtIGSTAmt = Format(Iamt, "#########0.00")
    mskDt(9) = "__/__/____"
End If

'--Roff
ROff = txt(17) '= 'CDbl(GProcMakeRounding(Amt) - Amt)
'-- SubAmt+Tax Amt +Add -Less+roff + Aft Tax Add - Aft Tax Less
Amt = SubAmt + Val(txtCGSTAmt) + Val(txtSGSTAmt) + Val(txtIGSTAmt) + ROff
  
Tds = wTaxable * Val(txt(22)) / 100
Tds = GProcMakeRounding(Tds)
txt(21) = Tds
  
  
'--Other Add
'--- Landing Cost =Bill+Freight+Rent +add-Less
txt(10) = SubAmt
txt(11) = Amt
txt(25) = txtCGSTRcm
txt(26) = txtSGSTRcm
txt(27) = txtIGSTRcm
RcmBlTot = Round((Val(txt(10)) + (Val(txtCGSTRcm) + Val(txtSGSTRcm) + Val(txtIGSTRcm))), 2)
RcmBlTot = GProcMakeRounding(RcmBlTot)
DoEvents
txt(29) = RcmBlTot
CheckForNumber
lblRcmSubAmt.Caption = Format((Val(txt(10))), FStr)
lblRCMTotal.Caption = Format((Val(txt(10)) + (Val(txtCGSTRcm) + Val(txtSGSTRcm) + Val(txtIGSTRcm))), FStr)

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
End Function

Private Sub GenRCMBillNo()
If chkIsRCM.Value = 1 And txt(43) = "" Then
    Dim Y As String
    Dim G_Rs As Recordset
    Dim GInvNo As Long
    Y = "Select Max(clng(mid(ADTMP5,5))) from tblAddLess where VYear=" & gCYear & " and Vtype='PO' and ADTMP5 <> ''"
    GProcRstOpen G_Rs, Y, "R", gCn
    If G_Rs.Fields(0) = Null Or IsNumeric(G_Rs.Fields(0)) = False Then
        GInvNo = 1
    Else
        GInvNo = CLng(G_Rs.Fields(0)) + 1
    End If
    txt(43) = "RCM" & "-" & GInvNo
    If IsDate(mskDt(9)) = False Then
       mskDt(9) = mskDt(2)
    End If
    G_Rs.Close
End If
End Sub



