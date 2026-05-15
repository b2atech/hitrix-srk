VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Begin VB.Form frmSales 
   BackColor       =   &H00C2E3FC&
   BorderStyle     =   0  'None
   Caption         =   "Sales Entry"
   ClientHeight    =   7500
   ClientLeft      =   0
   ClientTop       =   15
   ClientWidth     =   10845
   DrawMode        =   10  'Mask Pen
   FillColor       =   &H000040C0&
   FillStyle       =   0  'Solid
   Icon            =   "frmSales.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7500
   ScaleWidth      =   10845
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
      Height          =   405
      Index           =   54
      Left            =   0
      TabIndex        =   162
      Text            =   "54"
      Top             =   5280
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
      Height          =   405
      Index           =   53
      Left            =   9240
      TabIndex        =   161
      Text            =   "53"
      Top             =   6960
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
      Height          =   405
      Index           =   55
      Left            =   10080
      TabIndex        =   160
      Text            =   "55"
      Top             =   6960
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
      Height          =   405
      Index           =   52
      Left            =   9600
      TabIndex        =   159
      Text            =   "52"
      Top             =   6960
      Visible         =   0   'False
      Width           =   420
   End
   Begin VB.Frame Frame1 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFC0C0&
      Caption         =   "Bill Printing"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   3015
      Left            =   2400
      TabIndex        =   127
      Top             =   2100
      Visible         =   0   'False
      Width           =   5175
      Begin MSMask.MaskEdBox mskDate 
         Height          =   375
         Left            =   2640
         TabIndex        =   137
         Top             =   1080
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
      Begin VB.OptionButton optInvDate 
         BackColor       =   &H00FFC0C0&
         Caption         =   "Datewise"
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
         Index           =   1
         Left            =   2400
         TabIndex        =   136
         Top             =   480
         Width           =   1815
      End
      Begin VB.OptionButton optInvDate 
         BackColor       =   &H00FFC0C0&
         Caption         =   "Invoicewise"
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
         Index           =   0
         Left            =   240
         TabIndex        =   135
         Top             =   480
         Width           =   1815
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
         Left            =   3360
         TabIndex        =   134
         Top             =   2400
         Width           =   1335
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
         Left            =   1800
         TabIndex        =   133
         Top             =   2400
         Width           =   1335
      End
      Begin VB.TextBox txtPInvno 
         Appearance      =   0  'Flat
         Height          =   360
         Index           =   1
         Left            =   2640
         MaxLength       =   7
         TabIndex        =   129
         Top             =   1560
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
         Left            =   360
         TabIndex        =   130
         Top             =   2400
         Width           =   1335
      End
      Begin VB.TextBox txtPInvno 
         Appearance      =   0  'Flat
         Height          =   360
         Index           =   0
         Left            =   2640
         MaxLength       =   7
         TabIndex        =   128
         Top             =   1080
         Width           =   1695
      End
      Begin VB.Label Label18 
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
         Left            =   600
         TabIndex        =   132
         Top             =   1560
         Width           =   1665
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
         Left            =   600
         TabIndex        =   131
         Top             =   1080
         Width           =   1680
      End
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
      Left            =   3600
      TabIndex        =   156
      Text            =   "51"
      Top             =   8040
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
      Index           =   50
      Left            =   840
      TabIndex        =   155
      Text            =   "50"
      Top             =   8400
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
      Index           =   49
      Left            =   360
      TabIndex        =   154
      Text            =   "49"
      Top             =   8400
      Visible         =   0   'False
      Width           =   420
   End
   Begin VB.CheckBox Chk 
      BackColor       =   &H00ECFFFF&
      Caption         =   "Is E2 Sales ?"
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
      Left            =   3645
      TabIndex        =   153
      Top             =   5903
      Visible         =   0   'False
      Width           =   2055
   End
   Begin VB.CommandButton CmdOutstandingRpt 
      Caption         =   "&Outstand Report"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   810
      Left            =   9615
      TabIndex        =   152
      Top             =   3060
      Width           =   945
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
      Index           =   48
      Left            =   1470
      TabIndex        =   70
      Text            =   "txttemp(48)"
      ToolTipText     =   "Press F3 for Creation"
      Top             =   6300
      Width           =   1575
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
      Index           =   48
      Left            =   45
      TabIndex        =   151
      Text            =   "48"
      Top             =   1065
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
      Height          =   405
      Index           =   47
      Left            =   30
      TabIndex        =   150
      Text            =   "47"
      Top             =   1440
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
      Height          =   405
      Index           =   46
      Left            =   75
      TabIndex        =   149
      Text            =   "46"
      Top             =   1620
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
      Height          =   405
      Index           =   45
      Left            =   45
      TabIndex        =   148
      Text            =   "45"
      Top             =   705
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
      Height          =   405
      Index           =   44
      Left            =   0
      TabIndex        =   147
      Text            =   "44"
      Top             =   1935
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
      Height          =   405
      Index           =   43
      Left            =   0
      TabIndex        =   146
      Text            =   "43"
      Top             =   2340
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
      Height          =   405
      Index           =   42
      Left            =   -15
      TabIndex        =   145
      Text            =   "42"
      Top             =   2655
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
      Height          =   405
      Index           =   41
      Left            =   -90
      TabIndex        =   144
      Text            =   "41"
      Top             =   3045
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
      Height          =   405
      Index           =   40
      Left            =   15
      TabIndex        =   143
      Text            =   "40"
      Top             =   3435
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
      Height          =   405
      Index           =   39
      Left            =   -60
      TabIndex        =   142
      Text            =   "39"
      Top             =   3930
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
      Height          =   405
      Index           =   38
      Left            =   -15
      TabIndex        =   141
      Text            =   "38"
      Top             =   4410
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
      Height          =   405
      Index           =   37
      Left            =   -105
      TabIndex        =   140
      Text            =   "37"
      Top             =   4935
      Visible         =   0   'False
      Width           =   420
   End
   Begin VB.CommandButton CmdModifyAll 
      Caption         =   "Modify All Inv"
      Height          =   495
      Left            =   2385
      TabIndex        =   139
      Top             =   0
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.TextBox txtVSub 
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
      Left            =   7275
      TabIndex        =   138
      Text            =   "txtVSub(21)"
      Top             =   3795
      Width           =   960
   End
   Begin VB.TextBox txt 
      Alignment       =   1  'Right Justify
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
      Index           =   36
      Left            =   6810
      TabIndex        =   6
      Text            =   "36"
      Top             =   1635
      Width           =   1185
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
      ItemData        =   "frmSales.frx":058A
      Left            =   4980
      List            =   "frmSales.frx":0594
      TabIndex        =   41
      Text            =   "Cbo(12)"
      Top             =   5115
      Width           =   1230
   End
   Begin VB.CommandButton cmdLedger 
      Caption         =   "Ledger"
      Height          =   375
      Index           =   1
      Left            =   3405
      TabIndex        =   125
      Top             =   1620
      Width           =   1710
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
      Left            =   6810
      TabIndex        =   15
      Text            =   "35"
      ToolTipText     =   "Press F3 for Creation"
      Top             =   2535
      Width           =   2430
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
      Left            =   4920
      TabIndex        =   124
      Text            =   "34"
      Top             =   8040
      Visible         =   0   'False
      Width           =   330
   End
   Begin VB.TextBox txttmpAmt 
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
      Left            =   5760
      TabIndex        =   123
      Text            =   "txttmpAmt"
      Top             =   8040
      Visible         =   0   'False
      Width           =   180
   End
   Begin VB.TextBox txttmpInNo 
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
      Left            =   1470
      TabIndex        =   59
      Text            =   "txttmpInNo"
      Top             =   4320
      Width           =   1575
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
      Index           =   22
      Left            =   1470
      TabIndex        =   68
      Text            =   "22"
      Top             =   5895
      Width           =   1575
   End
   Begin VB.TextBox txtVSub 
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
      Index           =   12
      Left            =   7080
      TabIndex        =   60
      Text            =   "txtVSub(12)"
      Top             =   8520
      Visible         =   0   'False
      Width           =   825
   End
   Begin VB.TextBox txtVSub 
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
      Left            =   1470
      TabIndex        =   62
      Text            =   "txtVSub(10)"
      Top             =   4710
      Width           =   1575
   End
   Begin VB.TextBox txttmpVSub 
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
      Left            =   3015
      TabIndex        =   25
      Text            =   "txttmpVSub(17)"
      ToolTipText     =   "Press F3 for Creation"
      Top             =   3375
      Width           =   1245
   End
   Begin VB.TextBox txtVSub 
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
      Index           =   22
      Left            =   1635
      TabIndex        =   23
      Text            =   "txtVSub(22)"
      Top             =   3780
      Width           =   1320
   End
   Begin VB.TextBox txtVSub 
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
      Left            =   7935
      TabIndex        =   35
      Text            =   "txtVSub(9)"
      Top             =   3375
      Width           =   1380
   End
   Begin VB.TextBox txtVSub 
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
      Left            =   7275
      TabIndex        =   33
      Text            =   "txtVSub(8)"
      Top             =   3375
      Width           =   585
   End
   Begin VB.TextBox txtVSub 
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
      Left            =   6225
      TabIndex        =   31
      Text            =   "txtVSub(7)"
      Top             =   3375
      Width           =   960
   End
   Begin VB.TextBox txtVSub 
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
      Index           =   6
      Left            =   5025
      TabIndex        =   29
      Text            =   "txtVSub(6)"
      Top             =   3375
      Width           =   1125
   End
   Begin VB.TextBox txttmpVSub 
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
      Left            =   360
      Locked          =   -1  'True
      TabIndex        =   21
      Text            =   "txttmpVSub(4)"
      ToolTipText     =   "Press F3 for Creation"
      Top             =   3375
      Width           =   2595
   End
   Begin VB.TextBox txtVSub 
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
      Left            =   4335
      TabIndex        =   27
      Text            =   "txtVSub(5)"
      Top             =   3375
      Width           =   630
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
      Left            =   5760
      TabIndex        =   122
      Text            =   "33"
      Top             =   3780
      Width           =   1410
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
      Left            =   4335
      TabIndex        =   86
      Text            =   "32"
      Top             =   3765
      Width           =   1380
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
      Left            =   9255
      TabIndex        =   121
      Text            =   "31"
      Top             =   7995
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
      Left            =   9315
      TabIndex        =   93
      Text            =   "30"
      Top             =   8250
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
      Left            =   8205
      TabIndex        =   92
      Text            =   "29"
      Top             =   8190
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
      Index           =   28
      Left            =   6405
      TabIndex        =   91
      Text            =   "28"
      Top             =   8130
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
      Index           =   27
      Left            =   8760
      TabIndex        =   90
      Text            =   "27"
      Top             =   7965
      Visible         =   0   'False
      Width           =   495
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
      Left            =   7245
      TabIndex        =   89
      Text            =   "26"
      Top             =   8055
      Visible         =   0   'False
      Width           =   390
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
      Index           =   25
      Left            =   9810
      TabIndex        =   88
      Text            =   "25"
      Top             =   8295
      Visible         =   0   'False
      Width           =   510
   End
   Begin VB.CommandButton cmdLedger 
      Caption         =   "Ledger"
      Height          =   375
      Index           =   0
      Left            =   7110
      TabIndex        =   120
      Top             =   525
      Width           =   1815
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
      Left            =   1395
      TabIndex        =   119
      Text            =   "txtDrBalance"
      Top             =   1620
      Width           =   1890
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
      Left            =   5295
      TabIndex        =   118
      Text            =   "txtCrBalance"
      Top             =   555
      Width           =   1755
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
      Left            =   1395
      TabIndex        =   1
      Text            =   "txttemp(3)"
      ToolTipText     =   "Press F3 for Creation"
      Top             =   570
      Width           =   3765
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
      Index           =   24
      Left            =   4980
      TabIndex        =   39
      Text            =   "24"
      Top             =   4725
      Width           =   1230
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
      TabIndex        =   116
      Text            =   "1"
      Top             =   2430
      Width           =   420
   End
   Begin VB.TextBox txt 
      Alignment       =   1  'Right Justify
      Height          =   315
      Index           =   2
      Left            =   11040
      TabIndex        =   105
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
      TabIndex        =   104
      Text            =   "5"
      Top             =   2640
      Width           =   525
   End
   Begin VB.TextBox txt 
      Alignment       =   1  'Right Justify
      Height          =   315
      Index           =   3
      Left            =   11040
      TabIndex        =   103
      Text            =   "3"
      Top             =   1800
      Width           =   525
   End
   Begin VB.CommandButton cmdBtn 
      Caption         =   "Db&Nt Print"
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
      Left            =   9615
      Style           =   1  'Graphical
      TabIndex        =   97
      Top             =   2430
      Width           =   945
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
      Left            =   8685
      Style           =   1  'Graphical
      TabIndex        =   96
      Top             =   8220
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
      Left            =   9615
      Style           =   1  'Graphical
      TabIndex        =   85
      ToolTipText     =   "Exit"
      Top             =   6330
      Width           =   945
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
      Left            =   9615
      Style           =   1  'Graphical
      TabIndex        =   77
      Top             =   5820
      Width           =   945
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
      Left            =   9615
      Style           =   1  'Graphical
      TabIndex        =   76
      ToolTipText     =   "To Save Record"
      Top             =   5385
      Width           =   945
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
      Left            =   9615
      Style           =   1  'Graphical
      TabIndex        =   84
      Top             =   1980
      Width           =   945
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
      Left            =   9615
      Style           =   1  'Graphical
      TabIndex        =   75
      ToolTipText     =   "To Delete Record"
      Top             =   4845
      Width           =   945
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
      Left            =   9615
      Style           =   1  'Graphical
      TabIndex        =   74
      ToolTipText     =   "To Modify Record"
      Top             =   4410
      Width           =   945
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
      Left            =   9615
      Style           =   1  'Graphical
      TabIndex        =   73
      ToolTipText     =   "To Add New Record"
      Top             =   3975
      Width           =   945
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
      Left            =   8640
      Style           =   1  'Graphical
      TabIndex        =   83
      Top             =   255
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
      Left            =   9615
      Style           =   1  'Graphical
      TabIndex        =   78
      ToolTipText     =   "To Find Record"
      Top             =   1485
      Width           =   945
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
      Left            =   10140
      Style           =   1  'Graphical
      TabIndex        =   82
      ToolTipText     =   "Last"
      Top             =   1050
      Width           =   420
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
      Left            =   9615
      Style           =   1  'Graphical
      TabIndex        =   81
      ToolTipText     =   "First"
      Top             =   1050
      Width           =   420
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
      Left            =   10140
      Style           =   1  'Graphical
      TabIndex        =   80
      ToolTipText     =   "Next"
      Top             =   615
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
      Height          =   375
      Index           =   0
      Left            =   10560
      TabIndex        =   10
      Text            =   "0"
      Top             =   510
      Visible         =   0   'False
      Width           =   195
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
      Left            =   9615
      Style           =   1  'Graphical
      TabIndex        =   79
      ToolTipText     =   "Previous"
      Top             =   615
      Width           =   420
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
      Left            =   1395
      TabIndex        =   19
      Text            =   "txttemp(5)"
      ToolTipText     =   "Press F3 for Creation"
      Top             =   2550
      Width           =   3765
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
      Height          =   375
      Index           =   6
      Left            =   1395
      TabIndex        =   17
      Text            =   "txttempOS(6)"
      ToolTipText     =   "Press F3 for Creation"
      Top             =   2085
      Width           =   3765
   End
   Begin VB.TextBox txt 
      Alignment       =   1  'Right Justify
      Height          =   315
      Index           =   4
      Left            =   11040
      TabIndex        =   102
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
      Left            =   1395
      TabIndex        =   9
      Text            =   "txttemp(4)"
      ToolTipText     =   "Press F3 for Creation"
      Top             =   1155
      Width           =   3765
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
      Left            =   6810
      TabIndex        =   3
      Text            =   "6"
      Top             =   1185
      Width           =   1185
   End
   Begin VB.TextBox txt 
      Alignment       =   1  'Right Justify
      Height          =   315
      Index           =   7
      Left            =   11040
      TabIndex        =   101
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
      Height          =   375
      Index           =   8
      Left            =   6810
      TabIndex        =   12
      Text            =   "8"
      Top             =   2085
      Width           =   1185
   End
   Begin VB.TextBox txt 
      Alignment       =   1  'Right Justify
      Height          =   315
      Index           =   9
      Left            =   11040
      TabIndex        =   100
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
      Height          =   360
      Index           =   10
      Left            =   7935
      TabIndex        =   51
      Text            =   "10"
      Top             =   4335
      Width           =   1380
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
      Left            =   7935
      TabIndex        =   57
      Text            =   "11"
      Top             =   5910
      Width           =   1380
   End
   Begin VB.TextBox txt 
      Alignment       =   1  'Right Justify
      Height          =   315
      Index           =   12
      Left            =   11040
      TabIndex        =   99
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
      Left            =   6225
      TabIndex        =   42
      Text            =   "13"
      Top             =   5115
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
      Index           =   14
      Left            =   7935
      TabIndex        =   53
      Text            =   "14"
      Top             =   5115
      Width           =   1380
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
      Index           =   15
      Left            =   4980
      TabIndex        =   37
      Text            =   "15"
      Top             =   4335
      Width           =   1230
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
      Index           =   16
      Left            =   4980
      TabIndex        =   45
      Text            =   "16"
      Top             =   5505
      Width           =   1230
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
      Index           =   17
      Left            =   7935
      TabIndex        =   55
      Text            =   "17"
      Top             =   5505
      Width           =   1380
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
      Left            =   4980
      TabIndex        =   47
      Text            =   "txtOS(9)"
      Top             =   6457
      Width           =   1230
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
      Left            =   10560
      TabIndex        =   94
      Text            =   "18"
      Top             =   8145
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
      Height          =   405
      Index           =   19
      Left            =   6045
      TabIndex        =   95
      Text            =   "19"
      Top             =   7995
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
      Height          =   360
      Index           =   20
      Left            =   1470
      TabIndex        =   64
      Text            =   "20"
      Top             =   5190
      Width           =   1575
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
      Left            =   1470
      TabIndex        =   72
      Text            =   "21"
      Top             =   6705
      Width           =   1575
   End
   Begin VB.TextBox txt 
      Alignment       =   1  'Right Justify
      Height          =   315
      Index           =   23
      Left            =   11040
      TabIndex        =   98
      Text            =   "23"
      Top             =   5400
      Visible         =   0   'False
      Width           =   525
   End
   Begin MSMask.MaskEdBox mskDt 
      Height          =   375
      Index           =   2
      Left            =   8070
      TabIndex        =   7
      Top             =   1635
      Width           =   1170
      _ExtentX        =   2064
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
      Index           =   7
      Left            =   8070
      TabIndex        =   4
      Top             =   1185
      Width           =   1170
      _ExtentX        =   2064
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
      Height          =   360
      Index           =   10
      Left            =   4980
      TabIndex        =   49
      Top             =   6810
      Width           =   1230
      _ExtentX        =   2170
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
   Begin Crystal.CrystalReport cryReport 
      Bindings        =   "frmSales.frx":05A6
      Left            =   30
      Top             =   240
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
      Index           =   9
      Left            =   8070
      TabIndex        =   13
      Top             =   2085
      Width           =   1170
      _ExtentX        =   2064
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
      Left            =   7635
      TabIndex        =   87
      Top             =   7995
      Visible         =   0   'False
      Width           =   210
      _ExtentX        =   370
      _ExtentY        =   661
      _Version        =   393216
      MaxLength       =   10
      Format          =   "dd/mm/yyyy"
      Mask            =   "##/##/####"
      PromptChar      =   "_"
   End
   Begin MSMask.MaskEdBox mskDt 
      Height          =   315
      Index           =   34
      Left            =   1470
      TabIndex        =   66
      Top             =   5550
      Width           =   1575
      _ExtentX        =   2778
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
   Begin MSMask.MaskEdBox mskDt 
      Height          =   360
      Index           =   0
      Left            =   7920
      TabIndex        =   158
      Top             =   6810
      Visible         =   0   'False
      Width           =   1380
      _ExtentX        =   2434
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
   Begin VB.Label Label33 
      BackStyle       =   0  'Transparent
      Caption         =   "Goods Desp.Dt :"
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
      Left            =   6360
      TabIndex        =   157
      Top             =   6863
      Visible         =   0   'False
      Width           =   1605
   End
   Begin VB.Label Label23 
      BackStyle       =   0  'Transparent
      Caption         =   "Account   :"
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
      Left            =   360
      TabIndex        =   69
      Top             =   6300
      Width           =   1110
   End
   Begin VB.Label LblLRNo 
      BackStyle       =   0  'Transparent
      Caption         =   "LR No /Date :"
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
      Left            =   3015
      TabIndex        =   126
      Top             =   3795
      Width           =   1485
   End
   Begin VB.Label Label17 
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
      Left            =   6780
      TabIndex        =   50
      Top             =   4335
      Width           =   1095
   End
   Begin VB.Label Label4 
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
      Height          =   360
      Left            =   6765
      TabIndex        =   43
      Top             =   5115
      Visible         =   0   'False
      Width           =   360
   End
   Begin VB.Line Line3 
      BorderColor     =   &H000000C0&
      X1              =   255
      X2              =   3345
      Y1              =   5115
      Y2              =   5115
   End
   Begin VB.Label Label15 
      BackStyle       =   0  'Transparent
      Caption         =   "Due Date   :"
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
      Left            =   3630
      TabIndex        =   48
      Top             =   6803
      Width           =   1365
   End
   Begin VB.Label Label8 
      BackStyle       =   0  'Transparent
      Caption         =   "Ass. &Value    :"
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
      Left            =   3630
      TabIndex        =   38
      Top             =   4725
      Width           =   1365
   End
   Begin VB.Line Line1 
      BorderColor     =   &H000000C0&
      X1              =   3480
      X2              =   9360
      Y1              =   6330
      Y2              =   6330
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "Date         :"
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
      Left            =   360
      TabIndex        =   65
      Top             =   5550
      Width           =   1155
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Transport    :"
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
      Left            =   5580
      TabIndex        =   14
      Top             =   2535
      Width           =   1440
   End
   Begin VB.Label Label10 
      BackStyle       =   0  'Transparent
      Caption         =   "DbNt Amt :"
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
      Left            =   360
      TabIndex        =   71
      Top             =   6705
      Width           =   1530
   End
   Begin VB.Label Label45 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Amount"
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
      Height          =   360
      Left            =   7935
      TabIndex        =   34
      Top             =   3105
      Width           =   1380
   End
   Begin VB.Label Label44 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   " Rt.Per"
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
      Height          =   360
      Left            =   7275
      TabIndex        =   32
      Top             =   3105
      Width           =   630
   End
   Begin VB.Label Label43 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Rate"
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
      Height          =   360
      Left            =   6225
      TabIndex        =   30
      Top             =   3105
      Width           =   960
   End
   Begin VB.Label Label42 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Weight"
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
      Height          =   360
      Left            =   5025
      TabIndex        =   28
      Top             =   3105
      Width           =   1125
   End
   Begin VB.Label Label41 
      BackStyle       =   0  'Transparent
      Caption         =   "Count Name"
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
      Height          =   360
      Left            =   360
      TabIndex        =   20
      Top             =   3105
      Width           =   2595
   End
   Begin VB.Label Label40 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Bag"
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
      Height          =   360
      Left            =   4335
      TabIndex        =   26
      Top             =   3105
      Width           =   630
   End
   Begin VB.Shape Shape5 
      BorderColor     =   &H00C00000&
      BorderWidth     =   5
      FillColor       =   &H00C00000&
      Height          =   7815
      Left            =   0
      Top             =   -375
      Width           =   10845
   End
   Begin VB.Label Label39 
      BackStyle       =   0  'Transparent
      Caption         =   "Purch No  :"
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
      Left            =   360
      TabIndex        =   58
      Top             =   4320
      Width           =   1200
   End
   Begin VB.Label Label37 
      BackStyle       =   0  'Transparent
      Caption         =   "&Rate         :"
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
      Left            =   360
      TabIndex        =   67
      Top             =   5895
      Width           =   1110
   End
   Begin VB.Label Label36 
      BackStyle       =   0  'Transparent
      Caption         =   "Debit Nt.  :"
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
      Left            =   360
      TabIndex        =   63
      Top             =   5190
      Width           =   1245
   End
   Begin VB.Label Label35 
      BackStyle       =   0  'Transparent
      Caption         =   "Net Rate   :"
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
      TabIndex        =   22
      Top             =   3780
      Width           =   1530
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Invoice &No :"
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
      Left            =   5595
      TabIndex        =   5
      Top             =   1635
      Width           =   1125
   End
   Begin VB.Label LblExmptKg 
      BackStyle       =   0  'Transparent
      Caption         =   "Exempt / Kg. :"
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
      Left            =   3630
      TabIndex        =   44
      Top             =   5505
      Width           =   1395
   End
   Begin VB.Label Label32 
      BackStyle       =   0  'Transparent
      Caption         =   "Lot No      : "
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
      Left            =   360
      TabIndex        =   61
      Top             =   4710
      Width           =   1095
   End
   Begin VB.Label Label11 
      BackStyle       =   0  'Transparent
      Caption         =   "Godown"
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
      Left            =   3015
      TabIndex        =   24
      Top             =   3105
      Width           =   1245
   End
   Begin VB.Label Label12 
      BackStyle       =   0  'Transparent
      Caption         =   "D&ue Days  :"
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
      Left            =   3630
      TabIndex        =   46
      Top             =   6450
      Width           =   1125
   End
   Begin VB.Label Label30 
      BackStyle       =   0  'Transparent
      Caption         =   "Total :"
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
      Height          =   330
      Left            =   7170
      TabIndex        =   56
      Top             =   5925
      Width           =   855
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
      TabIndex        =   117
      Top             =   45
      Width           =   2475
   End
   Begin VB.Label Label28 
      BackStyle       =   0  'Transparent
      Caption         =   "ROff  : "
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
      Height          =   330
      Left            =   7200
      TabIndex        =   54
      Top             =   5520
      Width           =   675
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
      TabIndex        =   115
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
      TabIndex        =   114
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
      TabIndex        =   113
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
      TabIndex        =   112
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
      TabIndex        =   111
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
      TabIndex        =   110
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
      TabIndex        =   109
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
      TabIndex        =   108
      Top             =   7590
      Width           =   3135
   End
   Begin VB.Label LblExmptAmt 
      BackStyle       =   0  'Transparent
      Caption         =   "Exempt. Amt :"
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
      Left            =   3630
      TabIndex        =   36
      Top             =   4328
      Width           =   1485
   End
   Begin VB.Label Label16 
      BackStyle       =   0  'Transparent
      Caption         =   "TaxType        :"
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
      Left            =   3630
      TabIndex        =   40
      Top             =   5115
      Width           =   1335
   End
   Begin VB.Label Label14 
      BackStyle       =   0  'Transparent
      Caption         =   "Boo&king No :"
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
      Left            =   5580
      TabIndex        =   11
      Top             =   2085
      Width           =   1440
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
      TabIndex        =   107
      Top             =   8160
      Width           =   1095
   End
   Begin VB.Label Label9 
      BackStyle       =   0  'Transparent
      Caption         =   "Sale A/C  :"
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
      Left            =   360
      TabIndex        =   0
      Top             =   570
      Width           =   1470
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "&GP No        : "
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
      Left            =   5595
      TabIndex        =   2
      Top             =   1185
      Width           =   1500
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "&Broker   :"
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
      Left            =   360
      TabIndex        =   16
      Top             =   2085
      Width           =   1395
   End
   Begin VB.Label LblMill 
      BackStyle       =   0  'Transparent
      Caption         =   "M&ill       :"
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
      Left            =   360
      TabIndex        =   18
      Top             =   2550
      Width           =   1500
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Par&ty    :"
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
      Left            =   360
      TabIndex        =   8
      Top             =   1155
      Width           =   1440
   End
   Begin VB.Label LblSaleType 
      BackColor       =   &H00C00000&
      Caption         =   "  Trade Sale"
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
      TabIndex        =   106
      Top             =   0
      Width           =   10815
   End
   Begin VB.Shape Shape1 
      FillColor       =   &H00BDDFD1&
      FillStyle       =   0  'Solid
      Height          =   6465
      Left            =   9510
      Shape           =   4  'Rounded Rectangle
      Top             =   480
      Width           =   1185
   End
   Begin VB.Label Label29 
      BackStyle       =   0  'Transparent
      Caption         =   "Tax   :"
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
      Left            =   7200
      TabIndex        =   52
      Top             =   5115
      Width           =   945
   End
   Begin VB.Shape Shape4 
      BorderColor     =   &H000000C0&
      FillColor       =   &H00EDFDFE&
      FillStyle       =   0  'Solid
      Height          =   1950
      Left            =   240
      Top             =   1050
      Width           =   5025
   End
   Begin VB.Shape Shape9 
      BorderColor     =   &H000000C0&
      FillColor       =   &H00EDFDFE&
      FillStyle       =   0  'Solid
      Height          =   3000
      Left            =   240
      Top             =   4245
      Width           =   3135
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H00000080&
      FillColor       =   &H00EDFDFE&
      FillStyle       =   0  'Solid
      Height          =   585
      Left            =   240
      Top             =   435
      Width           =   9150
   End
   Begin VB.Shape Shape10 
      BorderColor     =   &H000000C0&
      FillColor       =   &H00EDFDFE&
      FillStyle       =   0  'Solid
      Height          =   1950
      Left            =   5430
      Top             =   1050
      Width           =   3960
   End
   Begin VB.Shape Shape6 
      BorderColor     =   &H000000C0&
      FillColor       =   &H00EDFDFE&
      FillStyle       =   0  'Solid
      Height          =   1200
      Left            =   240
      Top             =   3015
      Width           =   9165
   End
   Begin VB.Shape Shape7 
      BorderColor     =   &H00000080&
      FillColor       =   &H00EDFDFE&
      FillStyle       =   0  'Solid
      Height          =   3000
      Left            =   3480
      Top             =   4245
      Width           =   5955
   End
End
Attribute VB_Name = "frmSales"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Const tblName As String = "tblAddLess"
Const IdField As String = "Vno"
Const OrderField As String = "Vdt,Vno"
Const MaxNo As Long = 55
Dim SalesAccode As Long
Dim Rstbl As Recordset '-- TblAddLess
Dim rstblSub As Recordset '-- TblVouSub
Dim varBookMark As Variant
Dim FirstTimeFlag As Boolean
Dim rsVou As Recordset '--TblVoucher
Dim rsOS As Recordset '--TblOutStanding
Dim OldgFrmTypeStr As String
Dim FormAction As Integer
Dim frmTypeStr As String
Dim modflag As Boolean
Private Sub Cbo_KeyPress(Index As Integer, KeyAscii As Integer)
If Index = 12 Then KeyAscii = 0 '--Tax
End Sub
Private Sub cbo_LostFocus(Index As Integer)
If Index = 12 And FormAction = vbDataActionAddNew Then '--Tax %
'    txt(13) = GProcGetColumnValue("tblMastNarration", "Narration", Cbo(12), "S", "MastTaxRate", "N")
'    txt(13) = Format(txt(13), FStr)
'--By SRK FOr VAT From Item Master %
   Dim wMillCode As String
   Dim VATRtFromItem As Double
   wMillCode = GProcGetColumnValue("tblMastAccount", "AcName", txtTemp(5), "S", "AcCode", "N")
   VATRtFromItem = GProcGetColumnValue("tblMastItem", "ItName", txttmpVSub(4), "S", "ITTMP1", "N", "ItMillcode=" & wMillCode) '--Item Code
   If VATRtFromItem > 0 Then
      txt(13) = GProcGetColumnValue("tblMastItem", "ItName", txttmpVSub(4), "S", "ITTMP1", "N", "ItMillcode=" & wMillCode) '--Item Code
   Else
      txt(13) = GProcGetColumnValue("tblMastNarration", "Narration", Cbo(12), "S", "MastTaxRate", "N")
   End If
'----
   If gfrmHunkStr = "HNK" Then
      txt(13) = 0
   End If
   txt(13) = Format(txt(13), FStr)
End If
End Sub

'Dim GridFlag As Boolean
Private Sub cmdBtn_Click(Index As Integer)
On Error GoTo ErrorRoutine
Dim X As String
Dim OrderByStr As String
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
    If gBackEndDB = gBackEndAccess Then
        OrderByStr = "IIf(IsNumeric(Trim(adinvno))=-1,Clng(adinvno),clng(mid(adinvno,5))) ,vdt"
    Else
        OrderByStr = "decode(InStr(adinvno,'-',4),0,to_number(adinvno),to_number(substr(AdInVNo,5))),vdt"
    End If
    gClsSearch.SearchMultiField "tbladdLess,TblMastAccount,TblMastAccount TblMastAc1", "AdInvNo,Vno,Vdt,AdGpNo,TblMastAccount.AcName,TblMastAc1.AcName,AdbillAmt,AdGpDt", Array("Invoice No", "Vno", "Date", "Gp No", "Party", "Sales A/c", "Amount", "Gp Date"), Array(1000, 0, 1200, 1000, 2500, 2500, 1500, 1500), " AdAcDrCode=TblMastAccount.AcCode and AdAcCrcode=TblMastAc1.AcCode and " & IIf(SalesAccode = 0, "", " AdAcCrCode=" & SalesAccode & " and ") & " vtype='" & frmTypeStr & "' and VYear=" & gCYear, "", OrderByStr, 0, 0
    If gClsSearch.SearchMultiRetCol(0) <> "" Then
        Rstbl.MoveFirst
        Rstbl.Find "Vno = " & gClsSearch.SearchMultiRetCol(1)
        gClsSearch.SearchMultiRetCol(0) = ""
        ReadFields
    End If
Case 6 'Add
    Dim mTmpTxtSaleAc As String
    mTmpTxtSaleAc = txtTemp(3)
    txt(55) = "A"
    FormAction = vbDataActionAddNew
    Call GProcActivateControls(Me, True)
    SetControlEd
    Call GProcSetButtons(Me, vbDataActionAddNew)
    Call GProcClearForm(Me, Rstbl, MaxNo, True)
    If FirstTimeFlag = True Then
        FirstTimeFlag = False
    Else
       txt(6).SetFocus
    End If
    FillCombo
    txt(0) = GProcGenerateId("TblAddLess", "Vno", "Vtype='" & frmTypeStr & "' and VYear=" & gCYear)
    mskDt(2) = GProcGenerateVDt("tblAddLess", "Vdt", "Vtype='" & frmTypeStr & "' and VYear=" & gCYear)
    '-- Sales A/C
    txtTemp(3) = GProcGetColumnValue("tblMastAccount", "AcCode", CStr(SalesAccode), "N", "AcName", "S")
    txtTemp_LostFocus (3)
    
    'If (Left(mTmpTxtSaleAc, 3) <> "txt") Then txtTemp(3) = mTmpTxtSaleAc
Case 7 'Modify
    FormAction = vbDataActionUpdate
    If gfrmHunkStr = "HNK" Then
      Hank = "HNK"
    Else
      Hank = ""
    End If
'    Call GProcChkPass(frmTypeStr, FormAction, txt(36), txt(2), txtTemp(4), gUserName, Hank)
'    If gChkPassTrueFalse = True Then
'        txt(55) = "M"
        If IsRelatedRecord = False Then   '--- Or modflag = True
            Call GProcActivateControls(Me, True)
            txt(6).SetFocus
            SetControlEd
        Else
            Call GProcActivateControls(Me, False)
            SetControlEd
            mskDt(34).Enabled = True '- Db Dt
            txt(22).Enabled = True '- Db Amt
            txttempOS(6).Enabled = True '-- Broker
            mskDt(34).SetFocus
        End If
        Call GProcSetButtons(Me, vbDataActionUpdate)
        txt(36).Enabled = False
'    Else
'       FormAction = vbDataActionClose
'    End If
Case 8 'Delete
    FormAction = vbDataActionDelete
    If gfrmHunkStr = "HNK" Then
      Hank = "HNK"
    Else
      Hank = ""
    End If
    Call GProcChkPass(frmTypeStr, FormAction, txt(36), txt(2), txtTemp(4), gUserName, Hank)
    If gChkPassTrueFalse = True Then
        'If mURecDel = True Then
          If IsRelatedRecord = False Then
            If MsgBox("Do You Really Delete This Record?", vbQuestion + vbYesNo + vbDefaultButton1, "Confirm Delete") = vbYes Then
                If Rstbl.RecordCount > 1 Then
                     '-- Delete Related Record
                    DeleteRelatedRecord
                    Rstbl.Delete
                    cmdBtn_Click (1)
                ElseIf Rstbl.RecordCount = 1 Then
                    '-- Delete Related Record
                    DeleteRelatedRecord
                    Rstbl.Delete
                    Call GProcActivateControls(Me, False)
                    Call GProcSetButtons(Me, vbDataActionCancel)
                    Form_Load
                End If
            End If
         End If
    End If
Case 9, 12 'Print
    If Index = 9 Then
       Frame1.Caption = "Invoice Printing " & LblSaleType
    Else
       Frame1.Caption = "Debit Note Printing " & LblSaleType
    End If
    optInvDate(0).Value = True
        '-- Inv No
    If IsNumeric(txt(36)) = False Then
        txtPInvno(0) = Mid(txt(36).Text, 5)
        txtPInvno(1) = Mid(txt(36).Text, 5)
    Else
        txtPInvno(0) = txt(36).Text
        txtPInvno(1) = txt(36).Text
    End If
    Frame1.Visible = True
    DoEvents
    txtPInvno(0).Enabled = True
    txtPInvno(1).Enabled = True
    optInvDate(0).Enabled = True
    optInvDate(1).Enabled = True
    mskDate.Enabled = True
   txtPInvno(0).SetFocus
Case 10 'Exit
'    DoEvents
    Unload Me
    Exit Sub
Case 13 'Save
    txt(52) = CDate(Date)
    If ValidateData = True Then
        If FormAction = vbDataActionAddNew Then txt(0) = GProcGenerateId(tblName, IdField, "vtype='" & frmTypeStr & "' and VYear=" & gCYear)
        CalAmount
        FillTxtFromTemp
        '-- Save Record
        GProcSaveRecord Me, Rstbl, FormAction, MaxNo
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
Private Sub SetControlEd()
Dim ExemptMill As Long
txtCrBalance.Enabled = False
txtDrBalance.Enabled = False
txt(14).Enabled = False '--Tax
txt(11).Enabled = False '--Amt
txt(10).Enabled = False '--Sub Amt
txt(15).Enabled = False '--Exmpt Amt
txt(16).Enabled = False '--Exmpt Kg
txt(32).Enabled = False '--LR No
txt(33).Enabled = False '--LR dt
txtVSub(21).Enabled = False '--LR AMt
mskDt(7).Enabled = False '--GatePass Date
If frmTypeStr = "ST" Then '--SIT
    lblLRNo.Visible = True
    txt(32).Visible = True '--LR No
    txt(33).Visible = True '--LR dt
    txtVSub(21).Visible = True '--LR AMt
    Chk(1).Visible = True
    Chk(1).Enabled = True
    Label33.Visible = True
    mskDt(0).Visible = True
    mskDt(0).Enabled = True
Else
    lblLRNo.Visible = False
    txt(32).Visible = False '--LR No
    txt(33).Visible = False '--LR dt
    txtVSub(21).Visible = False '--LR AMt
    Label33.Visible = False
    mskDt(0).Visible = False
End If
txt(20).Enabled = False '--Debit No
txt(21).Enabled = False '--Debit Amt
'mskDt(34).Enabled = False '-- Db Date
txt(24).Enabled = False '--Ass Value
mskDt(33).Enabled = False '--LR Dt
txtTemp(4).Enabled = False '--Party
txtTemp(5).Enabled = False '--Mill
txtTemp(3).Enabled = False '--Sales A/c
txttmpInNo.Enabled = False '--Purchase No
txttmpAmt.Enabled = False
'--- V Sub
txttmpVSub(4).Enabled = False '--count
txtVSub(5).Enabled = False '--Bag
txtVSub(6).Enabled = False '--Wt
txtVSub(7).Enabled = False '--Rate
txtVSub(9).Enabled = False '--Sub Amt
txtVSub(12).Enabled = False '--Inward no
txtVSub(10).Enabled = False '--Lot no
txttmpVSub(17).Enabled = False '--Godown
txtVSub(21).Enabled = False '--LR amt
'-- Out Standing
txtOS(9).Enabled = False '--due Days
mskOS(10).Enabled = False '--Due Dt
txttempOS(6).Enabled = False '--Broker
'--- Mill
If frmTypeStr = "SY" Or frmTypeStr = "ST" Then
    txtTemp(5).Visible = True
    lblMill.Visible = True
Else
    txtTemp(5).Visible = False
    lblMill.Visible = False
End If
'-- Is Exempt Mill
ExemptMill = GProcGetColumnValue("TblMastAccount", "AcName", txtTemp(5), "S", "ACIsExemptMill", "N")
If ExemptMill = 1 Then
    'If FormAction = vbDataActionAddNew Or FormAction = vbDataActionUpdate Then
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
'-- Debit Note
gCn.Execute "delete from tblVoucher where Vtype='SN' and Vyear=" & gCYear & " and vno=" & txt(20)
End Sub
Private Sub SaveRelatedRecord()
Dim i As Long
Dim MillCode As Long
Dim rsVouSub As Recordset
Dim CrCode As Long
Dim DrCode As Long
Dim BrkCode As Long
Dim TaxAcCode As Long
Dim DebitNtAcCode As Long
Dim ItCode As Long
Dim IsNoVATVoucher As Boolean
Dim X As String
Dim BillDt
Dim DueDt
'--- Account Code
MillCode = GProcGetColumnValue("tblMastAccount", "AcName", txtTemp(5), "S", "AcCode", "N")
CrCode = GProcGetColumnValue("tblMastAccount", "AcName", txtTemp(3), "S", "AcCode", "N")
DrCode = GProcGetColumnValue("tblMastAccount", "AcName", txtTemp(4), "S", "AcCode", "N")
BrkCode = GProcGetColumnValue("tblMastAccount", "AcName", txttempOS(6), "S", "AcCode", "N")
TaxAcCode = GProcGetColumnValue("tblMastNarration", "Narration", Cbo(12), "S", "MastTaxSAAcCode", "N")
DebitNtAcCode = GProcGetColumnValue("tblMastAccount", "AcName", txtTemp(48), "S", "AcCode", "N")
'-- Delete Records
DeleteRelatedRecord
'--- TblVouSub
X = "select * from tblVouSub where vtype='" & frmTypeStr & "' and vno=" & txt(0) & " and " _
& " VYear=" & gCYear
GProcRstOpen rsVouSub, X, "O"
rsVouSub.AddNew
rsVouSub!Vno = CLng(txt(0)) '--Vno(0)
rsVouSub!Vtype = txt(1) & vbNullString '--Vtype(1)
rsVouSub!Vdt = CDate(txt(2)) '--Vdt(2)
rsVouSub!VSubCtrNo = 1 '--Sr No(3)
'--Itcode(4)
ItCode = GProcGetColumnValue("tblMastItem", "ItName", txttmpVSub(4), "S", "ItCode", "N", "ItMillcode=" & MillCode) '--Item Code
rsVouSub!vsubITCODE = ItCode
rsVouSub!vsubBAG = CLng(txtVSub(5))  '--Bag(5)
rsVouSub!vsubwt = CDbl(txtVSub(6)) '-- Wt(6)
rsVouSub!vsubRt = CDbl(txtVSub(7)) '--Rate(7)
rsVouSub!vsubRTPER = CDbl(txtVSub(8)) '--Rate Per(8)
rsVouSub!vsubAMT = CDbl(txtVSub(9)) '--Amt(9)
rsVouSub!vsubLOTNO = txtVSub(10) & vbNullString '--Lot No(10)
rsVouSub!vsubLRNO = txt(32) & vbNullString '--LR No(19)
rsVouSub!VSubLRAmt = CDbl(txtVSub(21)) '--LR Amt(21)
If IsDate(txt(33)) Then
    rsVouSub!vsubLRDate = txt(33) '--LR Dt(20)
Else
    rsVouSub!vsubLRDate = Null '--LR Dt(20)
End If
rsVouSub!VSubPItCtrlNo = txtVSub(12) '--PIT Ctrl No(12)
'--Godown Code(17)
rsVouSub!VSubGodownCode = GProcGetColumnValue("tblMastNarration", "Narration", txttmpVSub(17), "S", "NarrCode", "N", "")
rsVouSub!VSubItCtrlNo = GProcGenerateId("tblVouSub", "VSubItCtrlNo", "VYear=" & gCYear) '--It Ctrl No
rsVouSub!vsubNetRate = txtVSub(22) & vbNullString '--Net Rate
rsVouSub!VYear = txt(23) '--VYear(18)
rsVouSub.Update
rsVouSub.Close
'======================Voucher Creation
If GProcGetColumnValue("tblMastAccount", "AcCode", CStr(TaxAcCode), "N", "AcName", "S") = "-" Then
    '-- Bill Amt
    GProcCreateVoucher FormAction, txt(0), 1, txt(2), txt(1), CrCode, DrCode, txt(11), txt(36), mskDt(2)
    GProcCreateVoucher FormAction, txt(0), 2, txt(2), txt(1), DrCode, CrCode, -txt(11), txt(36), mskDt(2)
Else '-- If Tax A/c
        If frmTypeStr = "SY" And Cbo(12) = "V A T" Then IsNoVATVoucher = gIsNoVATVoucherSY
        If frmTypeStr = "SO" And Cbo(12) = "V A T" Then IsNoVATVoucher = gIsNoVATVoucherSO
        If frmTypeStr = "SD" And Cbo(12) = "V A T" Then IsNoVATVoucher = gIsNoVATVoucherSD
        '--- Is No Separate  VAT Voucher
        If Cbo(12) = "V A T" And IsNoVATVoucher = True Then
             '-- Bill Amt
            GProcCreateVoucher FormAction, txt(0), 1, txt(2), txt(1), CrCode, DrCode, txt(11), txt(36), mskDt(2)
            GProcCreateVoucher FormAction, txt(0), 2, txt(2), txt(1), DrCode, CrCode, -txt(11), txt(36), mskDt(2)
        Else
            '-- Bill -Tax Amt
            GProcCreateVoucher FormAction, txt(0), 1, txt(2), txt(1), CrCode, DrCode, txt(11) - txt(14), txt(36), mskDt(2)
            GProcCreateVoucher FormAction, txt(0), 2, txt(2), txt(1), DrCode, CrCode, -txt(11), txt(36), mskDt(2)
            '-- Tax Amt
            If CDbl(txt(14)) > 0 Then
                GProcCreateVoucher FormAction, txt(0), 3, txt(2), txt(1), TaxAcCode, DrCode, txt(14), txt(36), mskDt(2)
                'GProcCreateVoucher FormAction, txt(0), 4, txt(2), txt(1), DrCode, TaxAcCode, -txt(14), txt(36), mskDt(2)
            End If
        End If
End If
'-- Debit Amt
gCn.Execute "delete from tblVoucher where Vtype='SN' and Vyear=" & gCYear & " and vno=" & txt(20)
If CDbl(txt(21)) > 0 Then
    If txt(20) = 0 Then txt(20) = GProcGenerateId("tblVoucher", "Vno", "Vtype='SN' and Vyear=" & gCYear)
    GProcCreateVoucher FormAction, txt(20), 1, mskDt(34), "SN", DebitNtAcCode, DrCode, txt(21), txt(36), mskDt(2), "", "", "", "", "", True
    GProcCreateVoucher FormAction, txt(20), 2, mskDt(34), "SN", DrCode, DebitNtAcCode, -txt(21), txt(36), mskDt(2), "", "", "", "", "", True
End If
'======================Outstanding Creation
If IsDate(mskOS(10)) Then
    DueDt = mskOS(10)
Else
    DueDt = ""
End If
GProcCreateOutStanding txt(0), CDate(txt(2)), txt(1), txt(36), DrCode, BrkCode, CDbl(txt(11)), CLng(txtOS(9)), DueDt, MillCode, CDbl(txt(21)), txtVSub(5), ItCode
End Sub
'-- show Ledger
Private Sub cmdLedger_Click(Index As Integer)
On Error GoTo ErrorHandler
If Index = 0 Then
   frmRptAccount.txt(1) = txtTemp(3)
Else
   frmRptAccount.txt(1) = txtTemp(4)
End If
frmMain.mnuRptRSubAccLeg_Click (0)
Exit Sub
ErrorHandler:
    GProcErrorHandler
End Sub
'--- Outstanding Report
Private Sub CmdOutstandingRpt_Click()
On Error GoTo ErrorHandler
frmMain.mnurptRSubOutStaSale_Click (1)
frmRptRegi.txtTemp(2) = txtTemp(4) '- Party
Exit Sub
ErrorHandler:
    GProcErrorHandler
End Sub
Private Sub CmdModifyAll_Click()
    modflag = True
    cmdBtn_Click (2) '--First
    Do While Not Rstbl.EOF
        cmdBtn_Click (7) '-- Modify
        cmdBtn_Click (13) '-- Save
        cmdBtn_Click (1) '-- Next
       'MsgBox rstbl.Fields("vno")
       If Rstbl.EOF = True Then
            modflag = False
            Exit Sub
       End If
    Loop
    modflag = False
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
    Me.Top = 300
    Me.Left = Screen.Width / 2 - Me.Width / 2
    '--- TblAddLess
    frmTypeStr = gfrmTypeStr
    FirstTimeFlag = True
    If frmTypeStr = "SO" Then  ' --- Consignment Sale
        LblSaleType.Caption = "   Consignment Sale"
        SalesAccode = 0
    ElseIf frmTypeStr = "SD" Then ' --- Depot Sale
        LblSaleType.Caption = "   Depot Sale"
        SalesAccode = gClsSearch.SearchMultiRetCol(2)
    ElseIf frmTypeStr = "ST" Then ' --- SIT
        LblSaleType.Caption = "    Sale In Transite Sale"
        SalesAccode = gAcCodeST
    ElseIf frmTypeStr = "SY" Then
        If gfrmHunkStr = "HNK" Then
           SalesAccode = gAcCodeSYHank
        Else
           SalesAccode = gAcCodeSY
        End If
    End If
    If SalesAccode = 0 Then
        X = "Select * from " & tblName & " where  Vtype='" & frmTypeStr & "' and VYear=" & gCYear
    Else
        X = "Select * from " & tblName & " where AdAcCrCode=" & SalesAccode & " and Vtype='" & frmTypeStr & "' and VYear=" & gCYear
    End If
    If Trim(OrderField) <> "" Then X = X & " order by " & OrderField
    GProcRstOpen Rstbl, X, "O"
    '---Grid
    SetRelatedRS
    '---- Button Visiblity Set As Per User Previlage
    GProcSetButtonVisiblity Me
    cmdBtn_Click (6)
End Sub
Private Sub SetRelatedRS()
Dim X As String
Set rsOS = New Recordset
Set rsVou = New Recordset
Set rstblSub = New Recordset
'---TblVoucher
X = "Select * from tblVoucher where Vtype='" & frmTypeStr & "' and vno=" & txt(0) & " and VYear=" & gCYear & " order by VCtrNo"
GProcRstOpen rsVou, X, "O"
'---TblOutStanding
X = "Select * from tblOutStanding where Vtype='" & frmTypeStr & "' and vno=" & txt(0) & " and VYear=" & gCYear
GProcRstOpen rsOS, X, "O"
'-- tblVouSub
X = "select * from tblvouSub where vtype='" & frmTypeStr & "' and vno=" & txt(0) & " and VYear=" & gCYear & " order by VSubCtrNo"
GProcRstOpen rstblSub, X, "R"
End Sub
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
If FormAction = vbDataActionUpdate Then
    MsgBox "You are about to modify the record. Save the record and then exit.", vbInformation + vbOKOnly, Me.Caption
    Cancel = True
    Exit Sub
End If
'mShowRec = Space(5)
If Rstbl.State = adStateOpen Then Rstbl.Close
End Sub
Private Sub mskDate_GotFocus()
GProcSelectBox Me.ActiveControl
End Sub
Private Sub mskDt_GotFocus(Index As Integer)
GProcSelectBox Me.ActiveControl
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
Private Sub mskOS_GotFocus(Index As Integer)
GProcSelectBox Me.ActiveControl
End Sub
Private Sub mskOS_LostFocus(Index As Integer)
If IsDate(mskOS(Index)) = False Then mskOS(Index) = "__/__/____"
End Sub
Private Sub mskDt_LostFocus(Index As Integer)
'If Index = 2 Then
'    If GProcIsDateValid(mskDt(Index)) = False Then
'        mskDt(Index).SetFocus
'        Exit Sub
'    End If
'End If
If Index = 2 Then ShowAccountBalance
End Sub

Private Sub optInvDate_Click(Index As Integer)
If Index = 0 Then
   Label31.Caption = "From Invoice No :"
   mskDate.Visible = False
   txtPInvno(0).Visible = True
   txtPInvno(1).Visible = True
   Label18.Visible = True
Else
   Label31.Caption = "Print Date     :"
   mskDate.Visible = True
   txtPInvno(0).Visible = False
   txtPInvno(1).Visible = False
   Label18.Visible = False
End If
End Sub
Private Sub txt_GotFocus(Index As Integer)
GProcSelectBox Me.ActiveControl
End Sub
Private Sub txt_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyF3 Then
    Select Case Index
        Case 6 '--GP No
            GProcShowForm frmGatepass, frmMain.mnutrnSubGPFrm(0)
        Case 35 '-- Transport
            OldgFrmTypeStr = gfrmTypeStr
            gfrmTypeStr = "R"
            GProcShowForm frmMastNarration, frmMain.mnuMstfrm(8), "R"
            gfrmTypeStr = OldgFrmTypeStr
    End Select
End If
End Sub

Private Sub txt_KeyPress(Index As Integer, KeyAscii As Integer)
If KeyAscii = vbKeyReturn Then Exit Sub
KeyAscii = GProcValidateKey(Rstbl, Index, KeyAscii, txt(Index).Text)
If Index = 6 Then '--Gate Pass No
    PrepareGatePass
    KeyAscii = 0
End If
If Index = 35 Then '-- Transport
    gClsSearch.SearchMultiField "tblMastNarration", "Narration", Array("Transport"), Array(txt(Index).Width), " NarrType='R'", Chr(KeyAscii), "Narration", 0, 0, True
    KeyAscii = 0
    If gClsSearch.SearchMultiRetCol(0) <> "" Then
        txt(Index) = gClsSearch.SearchMultiRetCol(0) '--Transport
    End If
End If
End Sub
Private Sub PrepareGatePass()
Dim RsG As Recordset
Dim X As String
Dim GPSLType As Long
Dim MillCondStr As String
Dim wMillCode As String
Dim VATRtFromItem As Double

Set RsG = New Recordset
'If IsDate(mskDt(2)) = True Then
    If frmTypeStr = "SO" Then '--- Consignment
        GPSLType = 1
       ' MillCondStr = " and GPmill='" & txttemp(3) & "'"
    ElseIf frmTypeStr = "SD" Then '--Depot Sale
        MillCondStr = " and GPmill='" & txtTemp(3) & "'"
        GPSLType = 2
    ElseIf frmTypeStr = "ST" Then '--SIT
        GPSLType = 3
    Else
        '-- Trade
        If gfrmHunkStr = "HNK" Then
           GPSLType = 4
        Else
           GPSLType = 0
        End If
        MillCondStr = ""
    End If
    CheckForNumberTmp
    If gBackEndDB = gBackEndAccess Then
        X = "Select GpVno,format(GpVdt,'dd/mm/yyyy'),GPMill,GpParty,PCount,GPBag,GpWt,GpRate,Pvno,PVdt,PLRNo,PLRDt,PLRAmt,E2,DespDate from QryInwardVsGatePass where  GPSLType=" & GPSLType & " " & MillCondStr & " and (SAPItCtrlNo is null or ( isnumeric(SAPItCtrlNo) and SAPItCtrlNo=" & txtVSub(12) & " ) )"
    Else '-- Oracle
        X = "Select GpVno,GpVdt,GPMill,GpParty,PCount,GPBag,GpWt,GpRate,Pvno,PVdt,PLRNo,PLRDt,PLRAmt from QryInwardVsGatePass where  GPSLType=" & GPSLType & " " & MillCondStr & " and  "
        X = X + " GpVno not in (select AdGpNo from tblAddLess where vtype='" & frmTypeStr & "' and  vyear=" & gCYear & " and Vno<>" & txt(0) & ")"
    End If
    gClsSearch.SearchMultiField X, "", Array("Gp No", "Gp Date", "Mill", "Party", "Count", "Bag", "Wt", "Rate", "Purchase Vno", "Pur. Vdt", "LR No", "LR Date", "LR Amt", "E2", "Desp Dt"), Array(800, 1200, IIf(MillCondStr = "", 2000, 0), 2200, 2000, 1000, 1000, 1000, 1500, 1200, IIf(frmTypeStr = "ST", 1000, 0), IIf(frmTypeStr = "ST", 1200, 0), IIf(frmTypeStr = "ST", 1500, 0), 0, 0), "", "", "GpVdt,GpVNo", 0, 0, False, , True
    If gClsSearch.SearchMultiRetCol(0) <> "" Then
        txt(6) = gClsSearch.SearchMultiRetCol(0)
        mskDt(7) = Format(gClsSearch.SearchMultiRetCol(1), "dd/mm/yyyy")
        If frmTypeStr = "SO" Then
            txtTemp(3) = gClsSearch.SearchMultiRetCol(2) '--Mill
            txtTemp_LostFocus (3)
        End If
        If FormAction = vbDataActionAddNew Then mskDt(2) = CDate(mskDt(7))   '-- Inv Dt
        If FormAction = vbDataActionAddNew Then mskDt(34) = CDate(mskDt(7))  '-- DBNT Dt
        '--- Fill Gate Pass Details
        If gBackEndDB = gBackEndAccess Then
           X = "Select * from QryInwardVsGatePass where GpVNo=" & txt(6) & " and GpVdt=cdate('" & mskDt(7) & "')"
        Else '-- Oracle
           X = "Select * from QryInwardVsGatePass where GpVNo=" & txt(6) & " and GpVdt=to_date('" & mskDt(7) & "','dd/MM/yyyy')"
        End If
        GProcRstOpen RsG, X, "R"
        With RsG
            If RsG.EOF Then
            Else
                txtTemp(4) = .Fields("GpParty") '--Party
                txttempOS(6) = .Fields("GpBroker") '--Broker
                txtTemp(5) = .Fields("GpMill") '--Mill
                txttmpVSub(4) = .Fields("PCount") '--Count
                txtVSub(5) = .Fields("GpBag") '--Bag
                txtVSub(6) = .Fields("GpWt") '--Wt
                txtVSub(22) = .Fields("GpRate") '--Net Rate
                txtVSub(8) = .Fields("GpRatePer") '--Rate
                'txtVSub(12) = .Fields("GpPITCtrlNo") '--Gate Pass Pur .No
                txtVSub(12) = .Fields("GpITCtrlNo") '--Gate Pass It Ctrl .No
                '--Purchase No
                txttmpInNo = GProcGetColumnValue("TblvouSub", "VSubItCtrlNo", .Fields("GpPITCtrlNo"), "N", "VNo", "N")
                'txttmpInNo = GProcGetColumnValue("TblvouSub", "VSubItCtrlNo", txtVSub(12), "N", "VNo", "N")
                txtVSub(10) = .Fields("PLotNo") & vbNullString '--Lot No
                txttmpVSub(17) = .Fields("PGodown") '--Godown
                txt(15) = .Fields("PExemptAmt") '--Exmpt Amt
                txt(16) = .Fields("PExemptKg") '--Exmpt Kg
                txt(32) = .Fields("PLRNo") & vbNullString '--LR No
                txtVSub(21) = CDbl(.Fields("PLRAmt"))  '--LR Amt
                If IsDate(.Fields("PLRDt")) = True Then
                    mskDt(33) = Format(.Fields("PLRDt"), "dd/mm/yyyy") '--LR dt
                    txt(33) = Format(.Fields("PLRDt"), "dd/mm/yyyy") '--LR dt
                Else
                    mskDt(33) = "__/__/____"
                    txt(33) = ""  '--LR dt
                End If
                txtOS(9) = .Fields("GpDueDay") '--due days
                'cbo(12) = .Fields("PTaxType") '--Tax Type
                'txt(13) = .Fields("PTaxP") '--Tax %
                 Chk(1) = .Fields("E2") '--E2
                If IsDate(.Fields("DespDate")) = True Then
                    mskDt(0) = Format(.Fields("DespDate"), "dd/mm/yyyy") '--Desp Date
                    txt(51) = Format(.Fields("DespDate"), "dd/mm/yyyy") '--Desp Date
                Else
                    mskDt(0) = "__/__/____"
                    txt(51) = ""  '--Desp Date
                End If
                wMillCode = GProcGetColumnValue("tblMastAccount", "AcName", txtTemp(5), "S", "AcCode", "N")
                VATRtFromItem = GProcGetColumnValue("tblMastItem", "ItName", txttmpVSub(4), "S", "ITTMP1", "N", "ItMillcode=" & wMillCode) '--Item Code
                If VATRtFromItem > 0 Then
                   txt(13) = GProcGetColumnValue("tblMastItem", "ItName", txttmpVSub(4), "S", "ITTMP1", "N", "ItMillcode=" & wMillCode) '--Item Code
                End If
                If gfrmHunkStr = "HNK" Then
                   txt(13) = 0
                End If
                txt(46) = .Fields("GpNote") & vbNullString '--LR No
            
            
            
            End If
          RsG.Close
        End With
    End If
'End If
SetControlEd
CalAmount
ShowAccountBalance
End Sub
Private Sub txtVSub_GotFocus(Index As Integer)
GProcSelectBox Me.ActiveControl
End Sub
Private Sub txtVSub_KeyPress(Index As Integer, KeyAscii As Integer)
If KeyAscii = vbKeyReturn Then Exit Sub
KeyAscii = GProcValidateKey(rstblSub, Index, KeyAscii, txtVSub(Index).Text)
End Sub
Private Sub txtVSub_LostFocus(Index As Integer)
CheckForNumberTmp
SaveBtnEd
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
Private Sub txt_LostFocus(Index As Integer)
CkeckForNumber
SaveBtnEd
CalAmount
End Sub
Private Sub CkeckForNumber()
GProcCheckForNumber Me, Rstbl, MaxNo, Array(36)
End Sub
Private Sub txttemp_GotFocus(Index As Integer)
GProcSelectBox Me.ActiveControl
End Sub
Private Sub txttemp_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
Dim GName As String
If KeyCode = vbKeyF3 Then
    Select Case Index
        Case 3 '-- Sales A/c
            '--- Consignment & Depot Sale
            If frmTypeStr = "SO" Or frmTypeStr = "SD" Then '-- Mill
                GName = GProcGetColumnValue("TblMastGroup", "AgCode", 90029, "N", "AgName", "S")
                GProcShowForm frmMastAccount, frmMain.mnuMstfrm(0), Array(frmMastAccount.txtTemp(2)), Array(GName)
            Else
                GName = GProcGetColumnValue("TblMastGroup", "AgCode", 90023, "N", "AgName", "S")
                GProcShowForm frmMastAccount, frmMain.mnuMstfrm(0), Array(frmMastAccount.txtTemp(2)), Array(GName)
            End If
        Case 4 '-- Party
            GName = GProcGetColumnValue("TblMastGroup", "AgCode", 90017, "N", "AgName", "S")
            GProcShowForm frmMastAccount, frmMain.mnuMstfrm(0), Array(frmMastAccount.txtTemp(2)), Array(GName)
        Case 48 '-- Db Nt A/c
            GName = GProcGetColumnValue("TblMastGroup", "AgCode", 90027, "N", "AgName", "S")
            GProcShowForm frmMastAccount, frmMain.mnuMstfrm(0), Array(frmMastAccount.txtTemp(2)), Array(GName)
    End Select
End If
End Sub
Private Sub txtTemp_KeyPress(Index As Integer, KeyAscii As Integer)
Dim LeftPos As Long
Dim TopPos As Long
Dim AcCode As Long
Dim X As String
TopPos = txtTemp(Index).Top + Me.Top + 650
LeftPos = txtTemp(Index).Left + Me.Left
If KeyAscii = 13 Then Exit Sub
Select Case Index
    Case 4, 48  '-- Party ,Db Nt A/c  Alias Account
        '-- Search Condition
        If Index = 4 Then '-Party
            X = "(Agcode=90017 or GpCode=90017)"
        Else '-Db Nt A/c
            X = "(AcIsDbNtAc=1)"
            TopPos = 0
            LeftPos = 0
        End If
        gClsSearch.SearchMultiField "QryMastaccount", "Account,AcName,AcCity,AgName", Array("Account", "AcName", "City", "Group"), Array(3000, 0, 2000, 3000), X, Chr(KeyAscii), "acName", LeftPos, TopPos, True
        KeyAscii = 0
        If gClsSearch.SearchMultiRetCol(0) <> "" And gClsSearch.SearchMultiRetCol(1) = "" Then
            txtTemp(Index) = gClsSearch.SearchMultiRetCol(0)
        Else
            If gClsSearch.SearchMultiRetCol(1) <> "" Then txtTemp(Index) = gClsSearch.SearchMultiRetCol(1)
        End If
    Case 3 '-- Sales A/c  Alias Account
         '--- Consignment & Depot Sale
        If frmTypeStr = "SO" Or frmTypeStr = "SD" Then '-- Mill
            X = "(AgCode=90029 or GpCode=90029)"
        Else
            X = "(AgCode=90023 or GpCode=90023)"
        End If
        gClsSearch.SearchMultiField "QryMastaccount", "Account,AcName,AcCity,AgName", Array("Sales Account", "AcName", "City", "Group"), Array(3000, 0, 2000, 3000), X, Chr(KeyAscii), "AcName", LeftPos, TopPos, True
        KeyAscii = 0
        If gClsSearch.SearchMultiRetCol(0) <> "" And gClsSearch.SearchMultiRetCol(1) = "" Then
            txtTemp(Index) = gClsSearch.SearchMultiRetCol(0)
        Else
            If gClsSearch.SearchMultiRetCol(1) <> "" Then txtTemp(Index) = gClsSearch.SearchMultiRetCol(1)
        End If
        '--- Invoice No
        If (Index = 3 And FormAction = vbDataActionAddNew) Or txt(36) = "" Then
            AcCode = GProcGetColumnValue("TblMastAccount", "AcName", txtTemp(3), "S", "AcCode", "N")
            txt(36) = GProcGenerateInvNo(frmTypeStr, AcCode) '--Inv No
        End If
    Case 5  '-- Mill  Alias Account
        If frmTypeStr = "SY" Then
            gClsSearch.SearchMultiField "QryMastaccount", "Account,AcName,AcCity,AgName", Array("Mill", "AcName", "City", "Group"), Array(3000, 0, 2000, 3000), " ((AgCode=90029 or GpCode=90029) and  (AcMillType=0 or AcMillType=5))", Chr(KeyAscii), "AcName", LeftPos, TopPos, True
        ElseIf frmTypeStr = "ST" Then
            gClsSearch.SearchMultiField "QryMastaccount", "Account,AcName,AcCity,AgName", Array("Mill", "AcName", "City", "Group"), Array(3000, 0, 2000, 3000), " ((AgCode=90029 or GpCode=90029) and  (AcMillType=4 or AcMillType=5))", Chr(KeyAscii), "AcName", LeftPos, TopPos, True
        ElseIf frmTypeStr = "SO" Then
            gClsSearch.SearchMultiField "QryMastaccount", "Account,AcName,AcCity,AgName", Array("Mill", "AcName", "City", "Group"), Array(3000, 0, 2000, 3000), " ((AgCode=90029 or GpCode=90029) and  (AcMillType=1))", Chr(KeyAscii), "AcName", LeftPos, TopPos, True
        ElseIf frmTypeStr = "SD" Then
            gClsSearch.SearchMultiField "QryMastaccount", "Account,AcName,AcCity,AgName", Array("Mill", "AcName", "City", "Group"), Array(3000, 0, 2000, 3000), " ((AgCode=90029 or GpCode=90029) and  (AcMillType=2))", Chr(KeyAscii), "AcName", LeftPos, TopPos, True
        Else
            gClsSearch.SearchMultiField "QryMastaccount", "Account,AcName,AcCity,AgName", Array("Mill", "AcName", "City", "Group"), Array(3000, 0, 2000, 3000), " (AgCode=90029 or GpCode=90029)", Chr(KeyAscii), "AcName", LeftPos, TopPos, True
        End If
        KeyAscii = 0
        If gClsSearch.SearchMultiRetCol(0) <> "" And gClsSearch.SearchMultiRetCol(1) = "" Then
            txtTemp(Index) = gClsSearch.SearchMultiRetCol(0)
        Else
            If gClsSearch.SearchMultiRetCol(1) <> "" Then txtTemp(Index) = gClsSearch.SearchMultiRetCol(1)
        End If
'    Case 12 '-- Tax Type
'        gClsSearch.SearchMultiField "tblMastNarration", "Narration,MastTaxRate", Array("Tax Type", "TaxRate"), Array(txttemp(index).Width, 0), " NarrType='T'", Chr(KeyAscii), "Narration", 0, 0, True
'        KeyAscii = 0
'        If gClsSearch.SearchMultiRetCol(0) <> "" Then
'            txttemp(index) = gClsSearch.SearchMultiRetCol(0) '--Tax
'            If FormAction = vbDataActionAddNew Then txt(13) = gClsSearch.SearchMultiRetCol(1)  '--Tax %
'        End If
End Select
End Sub
Private Sub txtTemp_LostFocus(Index As Integer)
'If txtTemp(Index) <> "" And txtTemp(Index) <> "-" Then
Dim AcCode As Long
SaveBtnEd
ShowAccountBalance
If (Index = 3 And FormAction = vbDataActionAddNew) Or txt(36) = "" Then
     AcCode = GProcGetColumnValue("TblMastAccount", "AcName", txtTemp(3), "S", "AcCode", "N")
     txt(36) = GProcGenerateInvNo(frmTypeStr, AcCode) '--Inv No
End If
'End If
End Sub
Private Sub txtTempOS_LostFocus(Index As Integer)
SaveBtnEd
End Sub
Private Sub txttempOS_KeyPress(Index As Integer, KeyAscii As Integer)
Dim LeftPos As Long
Dim TopPos As Long
TopPos = txttempOS(Index).Top + Me.Top + 650
LeftPos = txttempOS(Index).Left + Me.Left
If KeyAscii <> 13 Then
If Index = 6 Then '-- Broker
    gClsSearch.SearchMultiField "QryMastaccount", "Account,AcName", Array("Broker", "AcName"), Array(txttempOS(Index).Width, 0), "(AgCode=90016 or GpCode=90016)", Chr(KeyAscii), "AcName", LeftPos, TopPos, True
    KeyAscii = 0
    If gClsSearch.SearchMultiRetCol(0) <> "" And gClsSearch.SearchMultiRetCol(1) = "" Then
        txttempOS(Index) = gClsSearch.SearchMultiRetCol(0)
    Else
        If gClsSearch.SearchMultiRetCol(1) <> "" Then txttempOS(Index) = gClsSearch.SearchMultiRetCol(1)
    End If
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
Private Sub SaveBtnEd()
'--- Vno,Inv No,vdt,Party,A/c,Mill,Amount
If CDbl(txt(0)) = 0 Or txt(36) = "" Or GProcIsDateValid(mskDt(2)) = False Or txtTemp(3) = "" Or txtTemp(5) = "" Or txtTemp(4) = "" Or CDbl(txt(11)) = 0 Then
    cmdBtn(13).Enabled = False
    cmdBtn(13).BackColor = &H80000004
Else
    cmdBtn(13).Enabled = True
    cmdBtn(13).BackColor = &HFFFF80
End If
End Sub
Private Sub FillCombo()
Call GProcFillCombo(Cbo(12), "select Narration from tblMastNarration where NarrType ='T' order by Narration")
Select Case frmTypeStr    '-- Tax Type
    Case "SY"
        If gfrmHunkStr = "HNK" Then
           Cbo(12) = GProcGetColumnValue("tblMastNarration", "NarrCode", CStr(gTaxCodeSYHunk), "N", "Narration", "S")
        Else
           Cbo(12) = GProcGetColumnValue("tblMastNarration", "NarrCode", CStr(gTaxCodeSY), "N", "Narration", "S")
        End If
'        Cbo(12) = GProcGetColumnValue("tblMastNarration", "NarrCode", CStr(gTaxCodeSY), "N", "Narration", "S")
    Case "ST"
        Cbo(12) = GProcGetColumnValue("tblMastNarration", "NarrCode", CStr(gTaxCodeST), "N", "Narration", "S")
    Case "SO"
        Cbo(12) = GProcGetColumnValue("tblMastNarration", "NarrCode", CStr(gTaxCodeSO), "N", "Narration", "S")
    Case "SD"
        Cbo(12) = GProcGetColumnValue("tblMastNarration", "NarrCode", CStr(gTaxCodeSD), "N", "Narration", "S")
End Select
'--Tax %
If FormAction = vbDataActionAddNew Then
    txt(13) = GProcGetColumnValue("tblMastNarration", "Narration", Cbo(12), "S", "MastTaxRate", "N")
    txt(13) = Format(txt(13), FStr)
End If
End Sub

Private Function ValidateData() As Boolean
Dim i As Long
Dim X As String
Dim Y As String
Dim GName As String
Dim rstAddCheck As Recordset
Dim rstModiCheck As Recordset
Dim ExmptRate As Double
Dim DBNTAcCode As Long
    '---- Required Data
    '--- Inv No.
    If txt(36) = "" Then
        MsgBox "Check Invoice No.", vbInformation + vbOKOnly, Me.Caption
        txt(36).SetFocus
        Exit Function
    End If
    '--- Vdt
    If GProcIsDateValid(mskDt(2)) = False Then
        MsgBox "Check Invoice Date.", vbInformation + vbOKOnly, Me.Caption
        mskDt(2).SetFocus
        Exit Function
    End If
    If mskDt(2) < mskDt(7) Then
        MsgBox "Check Invoice Date with Gatepass Date.", vbInformation + vbOKOnly, Me.Caption
        mskDt(2).SetFocus
        Exit Function
    End If
    '-- Gate Pass
    If txt(6) = 0 Then
        MsgBox "Check Gate Pass No.", vbInformation + vbOKOnly, Me.Caption
        txt(6).SetFocus
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
    '---- Sale Account
    If txtTemp(3) = "" Then
        MsgBox "Check Sale Account.", vbInformation + vbOKOnly, Me.Caption
        txtTemp(3).SetFocus
        Exit Function
    End If
        '--Create Sale A/c
        If GProcGetColumnValue("tblMastAccount", "AcName", txtTemp(3), "S", "AcCode", "N") = 0 Then
           If frmTypeStr = "SO" Or frmTypeStr = "SD" Then '-- Mill
                GName = GProcGetColumnValue("TblMastGroup", "AgCode", 90029, "N", "AgName", "S")
           Else
                GName = GProcGetColumnValue("TblMastGroup", "AgCode", 90023, "N", "AgName", "S")
           End If
           GProcShowForm frmMastAccount, frmMain.mnuMstfrm(0), Array(frmMastAccount.txt(1), frmMastAccount.txtTemp(2)), Array(txtTemp(3), GName)
           Exit Function
        End If
    '---- Party
    If txtTemp(4) = "" Then
        MsgBox "Check Party.", vbInformation + vbOKOnly, Me.Caption
        txtTemp(4).SetFocus
        Exit Function
    End If
        '--Create Party
        If GProcGetColumnValue("tblMastAccount", "AcName", txtTemp(4), "S", "AcCode", "N") = 0 Then
           GName = GProcGetColumnValue("TblMastGroup", "AgCode", 90017, "N", "AgName", "S")
           GProcShowForm frmMastAccount, frmMain.mnuMstfrm(0), Array(frmMastAccount.txt(1), frmMastAccount.txtTemp(2)), Array(txtTemp(4), GName)
           Exit Function
        End If
    If frmTypeStr = "SY" Or frmTypeStr = "ST" Then
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
    Else
        txtTemp(5) = txtTemp(3)
    End If
    '--- Broker
        '--Create Broker
        If GProcGetColumnValue("tblMastAccount", "AcName", txttempOS(6), "S", "AcCode", "N") = 0 Then
           GName = GProcGetColumnValue("TblMastGroup", "AgCode", 90016, "N", "AgName", "S")
           GProcShowForm frmMastAccount, frmMain.mnuMstfrm(0), Array(frmMastAccount.txt(1), frmMastAccount.txtTemp(2)), Array(txttempOS(6), GName)
           Exit Function
        End If
    '--- Db Nt A/c
        If CDbl(txt(21)) > 0 Then
            If txtTemp(48) = "" Then
                MsgBox "Check Debit Note Account.", vbInformation + vbOKOnly, Me.Caption
                txtTemp(48).SetFocus
                Exit Function
            End If
           '-- Create Db Nt A/c
           If GProcGetColumnValue("tblMastAccount", "AcName", txtTemp(48), "S", "AcCode", "N") = 0 Then
               GName = GProcGetColumnValue("TblMastGroup", "AgCode", 90027, "N", "AgName", "S")
               GProcShowForm frmMastAccount, frmMain.mnuMstfrm(0), Array(frmMastAccount.txt(1), frmMastAccount.txtTemp(2)), Array(txtTemp(48), GName)
               Exit Function
           End If
        End If
    '---- Amount
    If CDbl(txt(11)) <= 0 Then
        MsgBox "Check Voucher Amount.", vbInformation + vbOKOnly, Me.Caption
        Exit Function
    End If
     '-- Debit Amt
    If CDbl(txt(21)) > 0 And GProcIsDateValid(mskDt(34)) = False Then
        MsgBox "Check Debit Date.", vbInformation + vbOKOnly, App.Title
        mskDt(34).SetFocus
        Exit Function
    End If
    '-- Debit A/c
'    If gBillDbNtAcCode = GProcGetColumnValue("tblmastAccount", "AcName", "-", "S", "AcCode", "N") Then
'        DBNTAcCode = 0
'    Else
'        DBNTAcCode = gBillDbNtAcCode
'    End If
    If CDbl(txt(21)) > 0 And (txtTemp(48) = "" Or txtTemp(48) = "-") Then
        'And DBNTAcCode = 0 Then
        MsgBox "Check Debit Note Account.", vbInformation + vbOKOnly, Me.Caption
        GProcShowForm frmSetting, frmMain.mnuMstSetting(0)
        Exit Function
    End If
    '----Duplicate Records
    '----Addition Case '--Invoice No
    If FormAction = vbDataActionAddNew Then
        X = "Select vBillNo From tblVoucher Where VbillNo='" & Trim(txt(36)) & "' and Vtype='" & frmTypeStr & "' and VYear=" & gCYear
        i = GProcRstOpen(rstAddCheck, X, "R")
        If i > 0 Then
            MsgBox "Record is Available of Invoice No. " & CLng(txt(36)), vbCritical + vbOKOnly, "Duplicate Record"
            txt(36).SetFocus
            Exit Function
        End If
        rstAddCheck.Close
    End If
    '------------ Exmpt %
    ExmptRate = CDbl(GProcGetColumnValue("tblMastAccount", "AcName", txtTemp(5), "S", "AcExemptRate", "N"))
    If ExmptRate < 100 And ExmptRate > 0 And (frmTypeStr = "SO" Or frmTypeStr = "SD") Then
       MsgBox "Check Mill Exemption %", vbOKOnly + vbInformation, Me.Caption
       Exit Function
    End If
'    If (frmTypeStr = "SO" Or frmTypeStr = "SD") Then
'        MsgBox "     ", vbOKOnly + vbInformation, Me.Caption
'    End If
    ValidateData = True
End Function
Private Sub FillTxtFromTemp()
txt(1) = frmTypeStr '--VType
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
If IsDate(mskDt(7)) Then '-- GP dt
    txt(7) = CDate(mskDt(7))
Else
    txt(7) = ""
End If
If IsDate(mskDt(33)) Then '-- LR dt
    txt(33) = CDate(mskDt(33))
Else
    txt(33) = ""
End If
If IsDate(mskDt(34)) Then '-- Debit dt
    txt(34) = CDate(mskDt(34))
Else
    txt(34) = ""
End If
'-- Cr Code
txt(3) = GProcGetColumnValue("tblMastAccount", "AcName", txtTemp(3), "S", "AcCode", "N")
'-- Dr Code
txt(4) = GProcGetColumnValue("tblMastAccount", "AcName", txtTemp(4), "S", "AcCode", "N")
'--Mill
txt(5) = GProcGetColumnValue("tblMastAccount", "AcName", txtTemp(5), "S", "AcCode", "N")
'--DB Nt A/c
txt(48) = GProcGetColumnValue("tblMastAccount", "AcName", txtTemp(48), "S", "AcCode", "N")
'-- TaxType
'If txttemp(12) = "" Then txttemp(12) = "-"
'txt(12) = GProcGetColumnValue("tblMastNarration", "Narration", txttemp(12), "S", "NarrCode", "N")
txt(12) = GProcGetColumnValue("tblMastNarration", "Narration", Cbo(12), "S", "NarrCode", "N")
txt(23) = gCYear
'-- Debit
'-- Debit Amt
gCn.Execute "delete from tblVoucher where Vtype='SN' and Vyear=" & gCYear & " and vno=" & txt(20)
If CDbl(txt(21)) = 0 Then
    txt(20) = 0
    txt(34) = ""
    txt(48) = 0
End If
txt(40) = Chk(1)
If IsDate(mskDt(0)) Then '-- Desp dt
    txt(51) = CDate(mskDt(0))
Else
    txt(51) = ""
End If
ShowAccountBalance
End Sub
Private Sub FillTempFromTxt()
Dim X As String
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
If IsDate(txt(7)) Then '--GP Dt
    mskDt(7) = txt(7)
Else
    mskDt(7) = "__/__/____"
End If
If IsDate(txt(33)) Then '--LR Dt
    mskDt(33) = txt(33)
Else
    mskDt(33) = "__/__/____"
End If
If IsDate(txt(34)) Then '--Debit Dt
    mskDt(34) = txt(34)
Else
    mskDt(34) = "__/__/____"
End If
FillCombo
'-- Cr Code
txtTemp(3) = GProcGetColumnValue("tblMastAccount", "AcCode", txt(3), "N", "AcName", "S")
'-- Dr Code
txtTemp(4) = GProcGetColumnValue("tblMastAccount", "AcCode", txt(4), "N", "AcName", "S")
'--Mill
txtTemp(5) = GProcGetColumnValue("tblMastAccount", "AcCode", txt(5), "N", "AcName", "S")
'--DB Nt A/c
txtTemp(48) = GProcGetColumnValue("tblMastAccount", "AcCode", txt(48), "N", "AcName", "S")
'-- TaxType
'txttemp(12) = GProcGetColumnValue("tblMastNarration", "NarrCode", txt(12), "N", "Narration", "S")
Cbo(12) = GProcGetColumnValue("tblMastNarration", "NarrCode", txt(12), "N", "Narration", "S")
'-- Fill Related Record
FillRelatedRecord
'--- Total+ Debit Amt
txttmpAmt = CDbl(txt(11)) + CDbl(txt(21))
Chk(1) = txt(40)
If IsDate(txt(51)) Then '--Desp Dt
    mskDt(0) = txt(51)
Else
    mskDt(0) = "__/__/____"
End If

SetControlEd
ShowAccountBalance
End Sub
Private Sub FillRelatedRecord()
Dim Pu_ItCtrlNo As Long
'--- Fill Related Control
SetRelatedRS
'==== Vou Sub
txtVSub(5) = rstblSub.Fields("VSubBag") '--Bag
txtVSub(6) = rstblSub.Fields("VSubWt") '--Wt
txtVSub(7) = rstblSub.Fields("VSubRt") '--Rate
txtVSub(8) = rstblSub.Fields("VSubRtPer") '--Rate Per
txtVSub(9) = rstblSub.Fields("VSubAmt") '--Amt
txtVSub(22) = rstblSub.Fields("VSubNetRate") '--Net Rate
txtVSub(10) = rstblSub.Fields("VSubLotNo") & vbNullString  '--Lot No
txtVSub(21) = rstblSub.Fields("VSubLRAmt") '--LR Amt
txtVSub(12) = rstblSub.Fields("VSubPItCtrlNo") '--PIt Ctrl No
Pu_ItCtrlNo = GProcGetColumnValue("TblvouSub", "VSubItCtrlNo", txtVSub(12), "N", "VSubPItCtrlNo", "N", " vtype='GP'")
txttmpInNo = GProcGetColumnValue("TblvouSub", "VSubItCtrlNo", CStr(Pu_ItCtrlNo), "N", "VNo", "N")
'--Item
txttmpVSub(4) = GProcGetColumnValue("tblMastItem", "ItCode", rstblSub.Fields("VSubITCode"), "N", "ItName", "S")
'--Godown
txttmpVSub(17) = GProcGetColumnValue("tblMastNarration", "NarrCode", rstblSub.Fields("VSubGodownCode"), "N", "Narration", "S")
'==== OutStanding
'--Broker
If rsOS.Fields("OutBrokerCode") > 0 Then txttempOS(6) = GProcGetColumnValue("tblMastAccount", "AcCode", rsOS.Fields("OutBrokerCode"), "N", "AcName", "S")
txtOS(9) = rsOS.Fields("OutDueDays") '--Due Days
If IsDate(rsOS.Fields("OutDueDate")) = True Then '--Due Dt
    mskOS(10) = rsOS.Fields("OutDueDate")
Else
    mskOS(10) = "__/__/____"
End If
End Sub
Private Sub ReadFields()
FormAction = vbDataActionClose
GProcShowRecord Me, Rstbl, MaxNo
FillTempFromTxt
End Sub
Private Sub CheckForNumberTmp()
Dim i As Long
For i = 6 To 9
If IsNumeric(txtVSub(i)) = False Then txtVSub(i) = "0.00"
Next i
If IsNumeric(txtVSub(5)) = False Then txtVSub(5) = "0"
If IsNumeric(txtVSub(22)) = False Then txtVSub(22) = "0.00"
If IsNumeric(txtVSub(12)) = False Then txtVSub(12) = "0"
If IsNumeric(txtVSub(21)) = False Then txtVSub(21) = "0.00"
End Sub
Private Sub CalAmount()
Dim TaxAmt As Double
Dim TaxExmpt As Double
Dim TaxWithOutExmpt As Double
Dim SubAmt As Double
Dim SubExmpt As Double
Dim SubWithOutExmpt As Double
Dim Amt As Double
Dim AmtExmpt As Double
Dim AmtWithOutExmpt As Double
Dim Rate As Double
Dim ROff As Double
Dim MastExmptP As Double
Dim ExmptAmt As Double
Dim VAddWithTax As Double
Dim AssVal As Double
Dim AssValExmpt As Double
Dim DbAmt As Double
Dim IsTaxableInward As Long
Dim X As String
'Dim AssValWithOutExmpt As Double
CheckForNumberTmp
CkeckForNumber
If IsDate(mskDt(7)) = True Then
    If gBackEndDB = gBackEndAccess Then
        X = "tblVouSub_GP.Vdt=Cdate('" & mskDt(7) & "')"
    Else
        X = "tblVouSub_GP.Vdt=to_date('" & mskDt(7) & "','dd/MM/yyyy')"
    End If
    X = X + " and tblVouSub_GP.VSUBPITCTRLNO = tblVouSub.VSUBITCTRLNO and tblAddLess.Vno=tblVouSub.Vno" _
        & " and tblAddLess.VType=tblVouSub.VType and tblAddLess.VYear=tblVouSub.VYear"
    IsTaxableInward = GProcGetColumnValue("tblAddLess,tblVouSub,tblVouSub tblVouSub_GP", "tblVouSub_GP.VNo", txt(6), "N", "tblAddLess.AdTmp2", "N", X)
End If
'-- Debit Amt=Rate*Weight/Rate Per
If txtVSub(8) > 0 Then DbAmt = CDbl(txt(22)) * CDbl(txtVSub(6)) / CDbl(txtVSub(8))
'DbAmt = Format(DbAmt, FStr)
DbAmt = GProcMakeRounding(DbAmt)
txt(21) = DbAmt
'-- Debit No & A/c
If IsDate(mskDt(34)) = False Then mskDt(34) = mskDt(2)
If CDbl(txt(21)) > 0 Then
    If txt(20) = 0 Then txt(20) = GProcGenerateId("TblVoucher", "Vno", "Vtype='SN' and VYear=" & gCYear)
    If txtTemp(48) = "" Then txtTemp(48) = GProcGetColumnValue("tblmastAccount", "AcCode", CStr(gBillDbNtAcCode), "N", "AcName", "S")
End If
'-- Due Date
If IsNumeric(txtOS(9)) = False Then txtOS(9) = "0"
txtOS(9) = Format(txtOS(9), PStr)
If IsDate(mskDt(0)) Then
    mskOS(10) = DateAdd("d", txtOS(9), mskDt(0))
Else
   If IsDate(mskDt(2)) Then mskOS(10) = DateAdd("d", txtOS(9), mskDt(2))
End If
'-------------
MastExmptP = CDbl(GProcGetColumnValue("tblMastAccount", "AcName", txtTemp(5), "S", "AcExemptRate", "N"))
'-- Rate Per
If txtVSub(8) = 0 Then txtVSub(8) = 1
'=========== Mast Exmpt % =0
If MastExmptP = 0 Then
    '---- Amt =[Wt* Net Rate/Rate Per]
    Amt = (CDbl(txtVSub(6)) * CDbl(txtVSub(22)) / CDbl(txtVSub(8)))
    Amt = GProcMakeRounding(Amt)
    '--------- ( Universal ) ----------------------------
    If frmTypeStr = "ST" Then
        '-- Amt -LR Amt
        Amt = Amt - txtVSub(21)
        Amt = GProcMakeRounding(Amt)
    End If
    '------------------( Universal End)
    '--Sub Amt=(Amt)*100/(100+Tax %)
    SubAmt = Amt * 100 / (100 + CDbl(txt(13)))
    SubAmt = Format(SubAmt, FStr)
    '---Tax Amt= Sub amt*Taxp/100
    TaxAmt = SubAmt * CDbl(txt(13)) / 100
    TaxAmt = Format(TaxAmt, FStr)
    '--ROff = Amt - SubAmt - TaxAmt
    ROff = Amt - SubAmt - TaxAmt
    ROff = Format(ROff, FStr)
    '--Rate=SubAmt * Rate Per /Wt
    If CDbl(txtVSub(6)) > 0 Then
        Rate = SubAmt * CDbl(txtVSub(8)) / CDbl(txtVSub(6))
        Rate = Format(Rate, LStr)
    End If
    AssVal = SubAmt '--Ass .Value
ElseIf MastExmptP = 100 Then '-- Mast Exmpt % =100
    '---- Amt =[Wt* Net Rate/Rate Per]
    'If txtVSub(22) = 0 Then txtVSub(22) = txtVSub(7) '--New Rate
    Amt = (CDbl(txtVSub(6)) * CDbl(txtVSub(22)) / CDbl(txtVSub(8)))
    Amt = GProcMakeRounding(Amt)
    '--ExmptAmt=Wt* Exmpt /Kg
    If frmTypeStr = "SY" Then
        ExmptAmt = txtVSub(6) * CDbl(txt(16))
        ExmptAmt = Format(ExmptAmt, FStr)
    Else
        If IsTaxableInward = 1 Then
            ExmptAmt = 0
            txt(16) = 0
        Else
            ExmptAmt = Amt
            ExmptAmt = Format(ExmptAmt, FStr)
            txt(16) = ExmptAmt / CDbl(txtVSub(6)) '-- Exmpt/Kg
        End If
    End If
    '-- VAddWithTax
    VAddWithTax = Amt - ExmptAmt
    If VAddWithTax < 0 Then
        VAddWithTax = 0
        ExmptAmt = Amt
    End If
    '-- ASS. Value =(VAddWithTax)*100/(100+Tax %)
    AssVal = VAddWithTax * 100 / (100 + CDbl(txt(13)))
    AssVal = Format(AssVal, FStr)
    '-- Tax =Ass Val * % /100
    TaxAmt = AssVal * CDbl(txt(13)) / 100
    TaxAmt = Format(TaxAmt, FStr)
    '--Sub Amt=Amt-Tax
    SubAmt = Amt - TaxAmt
    SubAmt = Format(SubAmt, FStr)
    '--Rate=SubAmt * RatePer /Wt
    If CDbl(txtVSub(6)) > 0 Then
        Rate = SubAmt * CDbl(txtVSub(8)) / CDbl(txtVSub(6))
        Rate = Format(Rate, LStr)
    End If
    'txt(15) = ExmptAmt '--Exmpt Amt
    'txt(24) = AssVal '--Ass. Amt
Else
    '---- Amt =[Wt* Net Rate/Rate Per]
    'If txtVSub(22) = 0 Then txtVSub(22) = txtVSub(7) '--New Rate
    Amt = (CDbl(txtVSub(6)) * CDbl(txtVSub(22)) / CDbl(txtVSub(8)))
    Amt = GProcMakeRounding(Amt)
    '============================= Exmpted
    '--- Amt Exmpted =amt * Exmpt %
    AmtExmpt = Amt * MastExmptP / 100
    AmtExmpt = Format(AmtExmpt, FStr)
    '--ExmptAmt=Wt* Exmpt /Kg
    ExmptAmt = txtVSub(6) * CDbl(txt(16))
    ExmptAmt = Format(ExmptAmt, FStr)
    '-- VAddWithTax
    VAddWithTax = AmtExmpt - ExmptAmt
    If VAddWithTax < 0 Then
        VAddWithTax = 0
        ExmptAmt = AmtExmpt
    End If
    '-- ASS. Value Empt =(VAddWithTax)*100/(100+Tax %)
    AssValExmpt = VAddWithTax * 100 / (100 + CDbl(txt(13)))
    AssValExmpt = Format(AssValExmpt, FStr)
    '-- Tax Exmpt =Ass Val Exmpt * % /100
    TaxExmpt = AssValExmpt * CDbl(txt(13)) / 100
    TaxExmpt = Format(TaxExmpt, FStr)
    '--Sub Amt exmpt=Amt exmpt -Tax exmpt
    SubExmpt = AmtExmpt - TaxExmpt
    SubExmpt = Format(SubExmpt, FStr)
    '============================= With Out Exmpted
    '--- Amt With out Exmpted =Amt - Amt Exmpt
    AmtWithOutExmpt = Amt - AmtExmpt
    '--Sub Amt with out exmpt=(Amt)*100/(100+Tax %)
    SubWithOutExmpt = AmtWithOutExmpt * 100 / (100 + CDbl(txt(13)))
    SubWithOutExmpt = Format(SubWithOutExmpt, FStr)
    '---Tax Amt with out= Sub amt * tax P/100
    TaxWithOutExmpt = SubWithOutExmpt * CDbl(txt(13)) / 100
    TaxWithOutExmpt = Format(TaxWithOutExmpt, FStr)
    '-------------------------------------
    SubAmt = SubExmpt + SubWithOutExmpt
    TaxAmt = TaxExmpt + TaxWithOutExmpt
    AssVal = AssValExmpt + SubWithOutExmpt
    '--Rate=SubAmt * Rate Per/Wt
    If CDbl(txtVSub(6)) > 0 Then
        Rate = SubAmt * CDbl(txtVSub(8)) / CDbl(txtVSub(6))
        Rate = Format(Rate, LStr)
    End If
    'txt(15) = ExmptAmt '--Exmpt Amt
    'txt(24) = AssVal '--Ass. Amt
End If
ROff = Amt - SubAmt - TaxAmt  'Round(Amt, 0) - Amt
ROff = Format(ROff, FStr)
txtVSub(7) = Rate '--Rate
'If txtVSub(8) > 1 Then txtVSub(7) = Rate * txtVSub(8)
txtVSub(9) = Format(SubAmt, FStr) '--SubAmt
txt(14) = TaxAmt '--Tax
txt(11) = Amt '--Amt
txt(17) = ROff '--round off
txt(10) = SubAmt '--Sub amt
txt(24) = AssVal '-- Ass. Value
txt(15) = ExmptAmt '--Exmpt Amt
'--- Total+ Debit Amt
txttmpAmt = CDbl(txt(11)) + CDbl(txt(21))
CkeckForNumber
End Sub
'-- Related Record
Private Function IsRelatedRecord() As Boolean
Dim X As String
Dim Rs1 As Recordset
If modflag = True Then Exit Function '---
'-- Audited
If GProcISAudited(txt(0), txt(1), txt(23)) = True Then
    MsgBox "Record is Audited.You can not Modify/Delete this Record.", vbCritical + vbOKOnly, App.Title
    IsRelatedRecord = True
    Exit Function
End If
'-- Sales Return
SetRelatedRS
X = "Select Vtype,Vdt,Vno from tblVouSub where VSubPItCtrlNo=" & rstblSub.Fields("VSubITCtrlNo")
GProcRstOpen Rs1, X, "R"
If Rs1.EOF And Rs1.BOF Then
Else
    MsgBox "Related Sales Return Record " & Rs1.Fields("Vno") & " is exists.You can not Modify/Delete this Record.", vbCritical + vbOKOnly, App.Title
    IsRelatedRecord = True
    Exit Function
End If
'--- Receipt
X = "Select RecVno from tblRecVsSale where InvNo=" & txt(0) & " and InvTp='" & txt(1) & "' and InvYear=" & txt(23)
GProcRstOpen Rs1, X, "R"
If Rs1.EOF And Rs1.BOF Then
Else
    If FormAction = vbDataActionUpdate Then
        MsgBox "Related Receipt Record " & Rs1.Fields("RecVno") & " is exists.You can Modify Broker,Debit Note only.", vbCritical + vbOKOnly, App.Title
    Else
        MsgBox "Related Receipt Record " & Rs1.Fields("RecVno") & " is exists.You can not Delete this Record.", vbCritical + vbOKOnly, App.Title
    End If
    IsRelatedRecord = True
    Exit Function
End If
Rs1.Close
End Function
'------ Bill Print
Private Sub txtPInvno_GotFocus(Index As Integer)
GProcSelectBox Me.ActiveControl
End Sub

Private Sub txtPInvno_KeyPress(Index As Integer, KeyAscii As Integer)
KeyAscii = GProcNumberOnly(KeyAscii)
End Sub

Private Sub cmdPrintBill_Click(Index As Integer)
Dim RsSA As Recordset
Dim X As String
Dim BillSrNo As String
Dim InvNo As Long
Dim DbNtAcName As String
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
    If optInvDate(1).Value = True Then
       Dim mRrst As Recordset
        If gBackEndDB = gBackEndORA Then
            X = "Select * From tblAddLess where  VDt = to_date('" & Format(mskDate, "dd/MM/yyyy") & "','dd/MM/yyyy') and VType ='" & frmTypeStr & "'"
        Else
            X = "Select * From tblAddLess where  format(Vdt,'dd/MM/yyyy') = '" & Format(mskDate, "dd/MM/yyyy") & "' and VType ='" & frmTypeStr & "'"
        End If
    Else
       InvNo = CLng(txtPInvno(0))
       BillSrNo = GProcGetColumnValue("tblMastAccount", "AcName", txtTemp(3), "S", "AcBillSrNo", "S")
       X = "("
       Do While InvNo <= txtPInvno(1)
           If BillSrNo = "" Then
              X = X + "'" & InvNo & "',"
           Else
              X = X + "'" & BillSrNo & "-" & InvNo & "',"
           End If
           InvNo = InvNo + 1
       Loop
       X = Left(X, Len(X) - 1) + ")"
       X = "Select * From tblAddLess where  AdInvNo in " & X & " and VType='" & frmTypeStr & "' and VYear=" & gCYear
     End If
     If GProcRstOpen(RsSA, X, "R") <= 0 Then
       MsgBox "No Records..."
       Exit Sub
     End If
     RsSA.MoveFirst
     Do While Not RsSA.EOF
         With cryReport
                If Left(Frame1.Caption, 7) = "Invoice" Then
                    If Trim(RsSA!Vtype) = "SD" Then
                        .Formulas(0) = "FirmNm = '" & UCase(GProcGetColumnValue("tblMastAccount", "AcCode", RsSA!AdAcCrCode, "N", "AcName", "S")) & "'"
                        .Formulas(10) = "wCSTNo = '" & GProcGetColumnValue("tblMastAccount", "AcCode", RsSA!AdAcCrCode, "N", "AcCst", "S") & "'"
                        .Formulas(11) = "wBSTNo = '" & GProcGetColumnValue("tblMastAccount", "AcCode", RsSA!AdAcCrCode, "N", "ACBST", "S") & "'"
                        .Formulas(14) = "Payto = '" & gCName & "'"
                    Else
                        .Formulas(10) = "wCSTNo = '" & gCCST & "'"
                        .Formulas(11) = "wBSTNo = '" & gCBST & "'"
                    End If
                    
                    .Formulas(12) = "wAmountInWords = '" & GProcAmountInWords(RsSA!AdBillAmt) & "'"
                    .Formulas(13) = "wType = '" & RsSA!Vtype & "'"
                    .Formulas(16) = "wcBnknm = '" & gCBankName & "'"
                    .Formulas(17) = "wcBnkac= '" & gCBankAcNo & "'"
                    .Formulas(18) = "wcrtgs = '" & gCBankRtgsCode & "'"
                    .Formulas(20) = "wcPan = '" & gCPAN & "'"

                    .SelectionFormula = "{tblAddLess.AdInvNo} = '" & RsSA!AdInVNo & "' and {tblAddLess.VType}='" & RsSA!Vtype & "' and {tblAddLess.VYear}=" & RsSA!VYear
                    '--Report File At  D:\HiReports if exist
                    If Dir("D:\HiReports\rptSaleBillPrint.rpt") <> "" Then
                        .ReportFileName = "D:\HiReports\rptSaleBillPrint.rpt"
                    Else
                        .ReportFileName = gReportPath & "rptSaleBillPrint.rpt"
                    End If
                Else
                     DbNtAcName = GProcGetColumnValue("tblMastAccount,tblVoucher", "VNo", RsSA!AdDbNtNo, "N", "AcName", "S", " tblvoucher.Vtype='SN' and " _
                     & " TblVoucher.VYear=" & gCYear & " and tblMastAccount.AcCode=tblVoucher.VAcCode")
                    .Formulas(11) = "DbNtAccount= '" & DbNtAcName & "'"
                    .Formulas(12) = "wAmountInWords = '" & GProcAmountInWords(RsSA.Fields("AdDbNtAmt")) & "'"
                    .SelectionFormula = "{tblAddLess.AdInvNo} = '" & RsSA!AdInVNo & "' and {tblAddLess.VType}='" & RsSA!Vtype & "' and {tblAddLess.VYear}=" & RsSA!VYear
                    '--Report File At  D:\HiReports if exist
                    If Dir("D:\HiReports\rptSaleBillDbNtPrint.rpt") <> "" Then
                        .ReportFileName = "D:\HiReports\rptSaleBillDbNtPrint.rpt"
                    Else
                        .ReportFileName = gReportPath & "rptSaleBillDbNtPrint.rpt"
                    End If
                End If
                'Printer.PaperSize = vbPRPSUser
                .Action = 1
            End With
           RsSA.MoveNext
    Loop
Case 2:
    DoEvents
    Frame1.Visible = False
End Select
End Sub
