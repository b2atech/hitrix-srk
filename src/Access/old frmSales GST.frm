VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form frmSalesGST 
   BackColor       =   &H00C2E3FC&
   BorderStyle     =   0  'None
   Caption         =   "Sales Entry"
   ClientHeight    =   10260
   ClientLeft      =   0
   ClientTop       =   15
   ClientWidth     =   14100
   DrawMode        =   10  'Mask Pen
   FillColor       =   &H000040C0&
   FillStyle       =   0  'Solid
   Icon            =   "frmSales GST.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   10260
   ScaleWidth      =   14100
   ShowInTaskbar   =   0   'False
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
      Left            =   9915
      TabIndex        =   31
      Text            =   "58"
      Top             =   8280
      Width           =   2100
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
      Left            =   7380
      TabIndex        =   30
      Text            =   "57"
      Top             =   8325
      Width           =   750
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
      Index           =   56
      Left            =   3465
      TabIndex        =   170
      Text            =   "56"
      Top             =   10350
      Visible         =   0   'False
      Width           =   330
   End
   Begin VB.Frame Frame1 
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
      Height          =   4695
      Left            =   2385
      TabIndex        =   116
      Top             =   3510
      Visible         =   0   'False
      Width           =   6015
      Begin VB.CommandButton cmdPrintBill 
         Caption         =   "&Whatsapp - Text"
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
         Index           =   6
         Left            =   900
         TabIndex        =   173
         Top             =   3900
         Width           =   2835
      End
      Begin VB.CommandButton cmdPrintBill 
         Caption         =   "&S M S"
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
         Index           =   4
         Left            =   3780
         TabIndex        =   168
         Top             =   3360
         Width           =   1215
      End
      Begin VB.CheckBox ChkDc 
         BackColor       =   &H00FFC0C0&
         Caption         =   "Delivery Challan"
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
         Left            =   3240
         MaskColor       =   &H00FFC0C0&
         TabIndex        =   159
         Top             =   2160
         Width           =   2175
      End
      Begin VB.CheckBox Check1 
         BackColor       =   &H00FFC0C0&
         Caption         =   "Print Ticket Name"
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
         Height          =   375
         Left            =   600
         TabIndex        =   158
         Top             =   2160
         Width           =   2295
      End
      Begin VB.OptionButton OptDup 
         BackColor       =   &H00FFC0C0&
         Caption         =   "Duplicate Copy"
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
         Height          =   375
         Left            =   3240
         TabIndex        =   157
         Top             =   2760
         Width           =   2055
      End
      Begin VB.OptionButton OptOrg 
         BackColor       =   &H00FFC0C0&
         Caption         =   "Original Copy"
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
         Height          =   315
         Left            =   600
         TabIndex        =   156
         Top             =   2760
         Width           =   1815
      End
      Begin MSMask.MaskEdBox mskDate 
         Height          =   375
         Left            =   3240
         TabIndex        =   126
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
         Left            =   3000
         TabIndex        =   125
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
         Left            =   840
         TabIndex        =   124
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
         Left            =   3780
         TabIndex        =   123
         Top             =   3900
         Width           =   1200
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
         Left            =   2340
         TabIndex        =   122
         Top             =   3360
         Width           =   1335
      End
      Begin VB.TextBox txtPInvno 
         Appearance      =   0  'Flat
         Height          =   360
         Index           =   1
         Left            =   3240
         MaxLength       =   7
         TabIndex        =   118
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
         Left            =   900
         TabIndex        =   119
         Top             =   3360
         Width           =   1335
      End
      Begin VB.TextBox txtPInvno 
         Appearance      =   0  'Flat
         Height          =   360
         Index           =   0
         Left            =   3240
         MaxLength       =   7
         TabIndex        =   117
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
         Left            =   1200
         TabIndex        =   121
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
         Left            =   1200
         TabIndex        =   120
         Top             =   1080
         Width           =   1680
      End
   End
   Begin VB.CheckBox chkIsIncludingGst 
      BackColor       =   &H00EDFDFE&
      Caption         =   "Is Rate Excluding GST ?"
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
      Height          =   405
      Left            =   4140
      TabIndex        =   167
      Top             =   5460
      Width           =   4095
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
      Left            =   14280
      TabIndex        =   166
      Text            =   "54"
      Top             =   6360
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
      Index           =   54
      Left            =   14280
      TabIndex        =   165
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
      Index           =   53
      Left            =   14280
      TabIndex        =   164
      Text            =   "53"
      Top             =   5880
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
      Left            =   14280
      TabIndex        =   163
      Text            =   "52"
      Top             =   5400
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
      Index           =   51
      Left            =   14520
      TabIndex        =   162
      Text            =   "51"
      Top             =   7800
      Visible         =   0   'False
      Width           =   420
   End
   Begin VB.CheckBox Chk 
      BackColor       =   &H00F5F5DE&
      Caption         =   "Ticket Name?"
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
      Left            =   5595
      TabIndex        =   161
      Top             =   8985
      Width           =   2055
   End
   Begin VB.CheckBox Chk 
      BackColor       =   &H00F5F5DE&
      Caption         =   "Is Export?"
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
      Left            =   5595
      TabIndex        =   160
      Top             =   9465
      Width           =   2055
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
      Index           =   12
      Left            =   1485
      TabIndex        =   14
      Text            =   "txttemp(12)"
      ToolTipText     =   "Press F3 for Creation"
      Top             =   3555
      Width           =   5970
   End
   Begin VB.TextBox txtVSub 
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
      Index           =   52
      Left            =   9915
      TabIndex        =   29
      Text            =   "txtVSub(52)"
      Top             =   7740
      Width           =   2100
   End
   Begin VB.TextBox txtVSub 
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
      Height          =   360
      Index           =   51
      Left            =   7380
      TabIndex        =   28
      Text            =   "txtVSub(51)"
      Top             =   7830
      Width           =   750
   End
   Begin VB.TextBox txtVSub 
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
      Left            =   9915
      TabIndex        =   27
      Text            =   "txtVSub(50)"
      Top             =   7245
      Width           =   2100
   End
   Begin VB.TextBox txtVSub 
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
      Left            =   7380
      TabIndex        =   26
      Text            =   "txtVSub(49)"
      Top             =   7290
      Width           =   750
   End
   Begin VB.TextBox txtVSub 
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
      Index           =   48
      Left            =   9915
      TabIndex        =   25
      Text            =   "txtVSub(48)"
      Top             =   6750
      Width           =   2100
   End
   Begin VB.TextBox txtVSub 
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
      Index           =   47
      Left            =   7380
      TabIndex        =   24
      Text            =   "txtVSub(47)"
      Top             =   6795
      Width           =   750
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
      Index           =   6
      Left            =   9015
      TabIndex        =   2
      Text            =   "6"
      Top             =   1425
      Width           =   1635
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
      Index           =   45
      Left            =   9015
      TabIndex        =   4
      Text            =   "45"
      Top             =   1890
      Width           =   1635
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
      Index           =   50
      Left            =   17460
      TabIndex        =   48
      Text            =   "50"
      Top             =   7335
      Visible         =   0   'False
      Width           =   480
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
      Left            =   4605
      TabIndex        =   140
      Text            =   "49"
      Top             =   10605
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
      Index           =   33
      Left            =   17460
      TabIndex        =   42
      Text            =   "33"
      Top             =   8190
      Visible         =   0   'False
      Width           =   345
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
      Index           =   28
      Left            =   6345
      TabIndex        =   139
      Text            =   "txtVSub(28)"
      Top             =   10245
      Visible         =   0   'False
      Width           =   360
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
      Left            =   12675
      TabIndex        =   138
      Top             =   4230
      Width           =   945
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
      Left            =   765
      TabIndex        =   137
      Text            =   "48"
      Top             =   10785
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
      Left            =   3945
      TabIndex        =   136
      Text            =   "47"
      Top             =   10425
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
      Left            =   1785
      TabIndex        =   135
      Text            =   "46"
      Top             =   10845
      Visible         =   0   'False
      Width           =   420
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
      Index           =   44
      Left            =   9015
      TabIndex        =   9
      Text            =   "44"
      Top             =   3555
      Width           =   1635
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
      Left            =   2985
      TabIndex        =   134
      Text            =   "43"
      Top             =   10485
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
      Height          =   300
      Index           =   42
      Left            =   2625
      TabIndex        =   133
      Text            =   "42"
      Top             =   10470
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
      Left            =   2085
      TabIndex        =   132
      Text            =   "41"
      Top             =   10365
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
      Left            =   1725
      TabIndex        =   131
      Text            =   "40"
      Top             =   10365
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
      Left            =   1245
      TabIndex        =   130
      Text            =   "39"
      Top             =   10425
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
      Left            =   825
      TabIndex        =   129
      Text            =   "38"
      Top             =   10365
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
      Left            =   300
      TabIndex        =   128
      Text            =   "37"
      Top             =   10305
      Visible         =   0   'False
      Width           =   420
   End
   Begin VB.CommandButton CmdModifyAll 
      Caption         =   "Modify All Inv"
      Height          =   495
      Left            =   2385
      TabIndex        =   127
      Top             =   0
      Visible         =   0   'False
      Width           =   1215
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
      Index           =   36
      Left            =   9015
      TabIndex        =   5
      Text            =   "36"
      Top             =   2550
      Width           =   1635
   End
   Begin VB.CommandButton cmdLedger 
      Caption         =   "Ledger"
      Height          =   375
      Index           =   1
      Left            =   16380
      TabIndex        =   115
      Top             =   540
      Visible         =   0   'False
      Width           =   675
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
      Left            =   17460
      TabIndex        =   40
      Text            =   "35"
      ToolTipText     =   "Press F3 for Creation"
      Top             =   6930
      Visible         =   0   'False
      Width           =   480
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
      Left            =   5145
      TabIndex        =   114
      Text            =   "34"
      Top             =   10605
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
      Left            =   8145
      TabIndex        =   113
      Text            =   "txttmpAmt"
      Top             =   10245
      Visible         =   0   'False
      Width           =   180
   End
   Begin VB.TextBox txttmpInNo 
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
      Left            =   1920
      TabIndex        =   34
      Text            =   "txttmpInNo"
      Top             =   7095
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
      Height          =   420
      Index           =   22
      Left            =   17460
      TabIndex        =   51
      Text            =   "22"
      Top             =   5805
      Visible         =   0   'False
      Width           =   315
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
      Left            =   6825
      TabIndex        =   73
      Text            =   "txtVSub(12)"
      Top             =   10305
      Visible         =   0   'False
      Width           =   825
   End
   Begin VB.TextBox txtVSub 
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
      Index           =   10
      Left            =   1920
      TabIndex        =   35
      Text            =   "txtVSub(10)"
      Top             =   7545
      Width           =   1575
   End
   Begin VB.TextBox txttmpVSub 
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
      Left            =   3915
      TabIndex        =   16
      Text            =   "txttmpVSub(17)"
      ToolTipText     =   "Press F3 for Creation"
      Top             =   4950
      Width           =   1470
   End
   Begin VB.TextBox txtVSub 
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
      Left            =   2535
      TabIndex        =   19
      Text            =   "txtVSub(22)"
      Top             =   5445
      Width           =   1320
   End
   Begin VB.TextBox txtVSub 
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
      Index           =   9
      Left            =   9915
      TabIndex        =   22
      Text            =   "txtVSub(9)"
      Top             =   4950
      Width           =   2100
   End
   Begin VB.TextBox txtVSub 
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
      Index           =   8
      Left            =   9165
      TabIndex        =   21
      Text            =   "txtVSub(8)"
      Top             =   4950
      Width           =   585
   End
   Begin VB.TextBox txtVSub 
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
      Index           =   7
      Left            =   7710
      TabIndex        =   20
      Text            =   "txtVSub(7)"
      Top             =   4950
      Width           =   1365
   End
   Begin VB.TextBox txtVSub 
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
      Index           =   6
      Left            =   6150
      TabIndex        =   18
      Text            =   "txtVSub(6)"
      Top             =   4950
      Width           =   1485
   End
   Begin VB.TextBox txttmpVSub 
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
      Index           =   4
      Left            =   360
      Locked          =   -1  'True
      TabIndex        =   15
      Text            =   "txttmpVSub(4)"
      ToolTipText     =   "Press F3 for Creation"
      Top             =   4950
      Width           =   3495
   End
   Begin VB.TextBox txtVSub 
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
      Index           =   5
      Left            =   5415
      TabIndex        =   17
      Text            =   "txtVSub(5)"
      Top             =   4950
      Width           =   675
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
      Left            =   17505
      TabIndex        =   41
      Text            =   "32"
      Top             =   7740
      Visible         =   0   'False
      Width           =   390
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
      Left            =   10845
      TabIndex        =   112
      Text            =   "31"
      Top             =   10575
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
      Left            =   10185
      TabIndex        =   94
      Text            =   "30"
      Top             =   10365
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
      Left            =   8685
      TabIndex        =   93
      Text            =   "29"
      Top             =   10365
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
      Left            =   5505
      TabIndex        =   92
      Text            =   "28"
      Top             =   10725
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
      Left            =   10305
      TabIndex        =   91
      Text            =   "27"
      Top             =   10620
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
      Left            =   7665
      TabIndex        =   90
      Text            =   "26"
      Top             =   10245
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
      Left            =   10425
      TabIndex        =   89
      Text            =   "25"
      Top             =   10425
      Visible         =   0   'False
      Width           =   510
   End
   Begin VB.CommandButton cmdLedger 
      Caption         =   "Ledger"
      Height          =   375
      Index           =   0
      Left            =   10170
      TabIndex        =   111
      Top             =   660
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
      Left            =   16380
      TabIndex        =   39
      Text            =   "txtDrBalance"
      Top             =   90
      Width           =   765
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
      Left            =   7830
      TabIndex        =   110
      Text            =   "txtCrBalance"
      Top             =   660
      Width           =   1755
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
      Index           =   3
      Left            =   1485
      TabIndex        =   1
      Text            =   "txttemp(3)"
      ToolTipText     =   "Press F3 for Creation"
      Top             =   660
      Width           =   5970
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
      Left            =   17460
      TabIndex        =   45
      Text            =   "24"
      Top             =   6570
      Visible         =   0   'False
      Width           =   285
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
      Left            =   17460
      TabIndex        =   108
      Text            =   "1"
      Top             =   540
      Width           =   420
   End
   Begin VB.TextBox txt 
      Alignment       =   1  'Right Justify
      Height          =   315
      Index           =   2
      Left            =   17460
      TabIndex        =   106
      Text            =   "2"
      Top             =   945
      Visible         =   0   'False
      Width           =   525
   End
   Begin VB.TextBox txt 
      Alignment       =   1  'Right Justify
      Height          =   315
      Index           =   5
      Left            =   17460
      TabIndex        =   105
      Text            =   "5"
      Top             =   1935
      Width           =   525
   End
   Begin VB.TextBox txt 
      Alignment       =   1  'Right Justify
      Height          =   315
      Index           =   3
      Left            =   17460
      TabIndex        =   104
      Text            =   "3"
      Top             =   1260
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
      Left            =   12675
      Style           =   1  'Graphical
      TabIndex        =   98
      Top             =   3450
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
      Height          =   405
      Index           =   11
      Left            =   9285
      Style           =   1  'Graphical
      TabIndex        =   97
      Top             =   10485
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
      Left            =   12675
      Style           =   1  'Graphical
      TabIndex        =   87
      ToolTipText     =   "Exit"
      Top             =   8190
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
      Left            =   12675
      Style           =   1  'Graphical
      TabIndex        =   79
      Top             =   7650
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
      Left            =   12675
      Style           =   1  'Graphical
      TabIndex        =   78
      ToolTipText     =   "To Save Record"
      Top             =   7050
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
      Left            =   12675
      Style           =   1  'Graphical
      TabIndex        =   86
      Top             =   2850
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
      Left            =   12675
      Style           =   1  'Graphical
      TabIndex        =   77
      ToolTipText     =   "To Delete Record"
      Top             =   6450
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
      Left            =   12675
      Style           =   1  'Graphical
      TabIndex        =   76
      ToolTipText     =   "To Modify Record"
      Top             =   5850
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
      Left            =   12675
      Style           =   1  'Graphical
      TabIndex        =   75
      ToolTipText     =   "To Add New Record"
      Top             =   5250
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
      TabIndex        =   85
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
      Left            =   12675
      Style           =   1  'Graphical
      TabIndex        =   80
      ToolTipText     =   "To Find Record"
      Top             =   2295
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
      Left            =   13200
      Style           =   1  'Graphical
      TabIndex        =   84
      ToolTipText     =   "Last"
      Top             =   1740
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
      Left            =   12675
      Style           =   1  'Graphical
      TabIndex        =   83
      ToolTipText     =   "First"
      Top             =   1740
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
      Left            =   13200
      Style           =   1  'Graphical
      TabIndex        =   82
      ToolTipText     =   "Next"
      Top             =   1185
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
      Left            =   17460
      TabIndex        =   55
      Text            =   "0"
      Top             =   45
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
      Left            =   12675
      Style           =   1  'Graphical
      TabIndex        =   81
      ToolTipText     =   "Previous"
      Top             =   1185
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
      Height          =   390
      Index           =   5
      Left            =   1485
      TabIndex        =   13
      Text            =   "txttemp(5)"
      ToolTipText     =   "Press F3 for Creation"
      Top             =   2910
      Width           =   5970
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
      Left            =   1485
      TabIndex        =   12
      Text            =   "txttempOS(6)"
      ToolTipText     =   "Press F3 for Creation"
      Top             =   2355
      Width           =   3990
   End
   Begin VB.TextBox txt 
      Alignment       =   1  'Right Justify
      Height          =   315
      Index           =   4
      Left            =   17460
      TabIndex        =   103
      Text            =   "4"
      Top             =   1575
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
      Left            =   1485
      TabIndex        =   11
      Text            =   "txttemp(4)"
      ToolTipText     =   "Press F3 for Creation"
      Top             =   1380
      Width           =   5970
   End
   Begin VB.TextBox txt 
      Alignment       =   1  'Right Justify
      Height          =   315
      Index           =   7
      Left            =   17460
      TabIndex        =   102
      Text            =   "7"
      Top             =   2295
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
      Index           =   8
      Left            =   9015
      TabIndex        =   7
      Text            =   "8"
      Top             =   3060
      Width           =   1635
   End
   Begin VB.TextBox txt 
      Alignment       =   1  'Right Justify
      Height          =   315
      Index           =   9
      Left            =   17460
      TabIndex        =   101
      Text            =   "9"
      Top             =   2655
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
      Height          =   360
      Index           =   10
      Left            =   9915
      TabIndex        =   23
      Text            =   "10"
      Top             =   6300
      Width           =   2100
   End
   Begin VB.TextBox txt 
      Alignment       =   1  'Right Justify
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   480
      Index           =   11
      Left            =   9015
      TabIndex        =   33
      Text            =   "11"
      Top             =   9360
      Width           =   3000
   End
   Begin VB.TextBox txt 
      Alignment       =   1  'Right Justify
      Height          =   315
      Index           =   12
      Left            =   17460
      TabIndex        =   100
      Text            =   "12"
      Top             =   3015
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
      Left            =   17460
      TabIndex        =   46
      Text            =   "13"
      Top             =   3420
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
      Index           =   14
      Left            =   17460
      TabIndex        =   47
      Text            =   "14"
      Top             =   3825
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
      Height          =   360
      Index           =   15
      Left            =   17460
      TabIndex        =   44
      Text            =   "15"
      Top             =   4275
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
      Height          =   360
      Index           =   16
      Left            =   17460
      TabIndex        =   49
      Text            =   "16"
      Top             =   4590
      Visible         =   0   'False
      Width           =   375
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
      Left            =   9915
      TabIndex        =   32
      Text            =   "17"
      Top             =   8820
      Width           =   2100
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
      Height          =   360
      Index           =   9
      Left            =   1575
      TabIndex        =   36
      Text            =   "txtOS(9)"
      Top             =   8235
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
      Left            =   10965
      TabIndex        =   95
      Text            =   "18"
      Top             =   10365
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
      Left            =   5925
      TabIndex        =   96
      Text            =   "19"
      Top             =   10545
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
      Height          =   420
      Index           =   20
      Left            =   17460
      TabIndex        =   50
      Text            =   "20"
      Top             =   4950
      Visible         =   0   'False
      Width           =   360
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
      Height          =   420
      Index           =   21
      Left            =   17460
      TabIndex        =   52
      Text            =   "21"
      Top             =   5310
      Visible         =   0   'False
      Width           =   405
   End
   Begin VB.TextBox txt 
      Alignment       =   1  'Right Justify
      Height          =   315
      Index           =   23
      Left            =   17460
      TabIndex        =   99
      Text            =   "23"
      Top             =   6210
      Visible         =   0   'False
      Width           =   525
   End
   Begin MSMask.MaskEdBox mskDt 
      Height          =   375
      Index           =   2
      Left            =   10725
      TabIndex        =   6
      Top             =   2550
      Width           =   1395
      _ExtentX        =   2461
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
      Index           =   7
      Left            =   10725
      TabIndex        =   3
      Top             =   1433
      Width           =   1395
      _ExtentX        =   2461
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
      Left            =   1575
      TabIndex        =   37
      Top             =   8775
      Width           =   1470
      _ExtentX        =   2593
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
      Bindings        =   "frmSales GST.frx":058A
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
      Left            =   10725
      TabIndex        =   8
      Top             =   3060
      Width           =   1395
      _ExtentX        =   2461
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
      Index           =   33
      Left            =   8445
      TabIndex        =   88
      Top             =   10245
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
      Height          =   375
      Index           =   0
      Left            =   16740
      TabIndex        =   38
      Top             =   5535
      Width           =   270
      _ExtentX        =   476
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
      Left            =   10710
      TabIndex        =   10
      Top             =   3555
      Width           =   1395
      _ExtentX        =   2461
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
   Begin VB.Label Label26 
      BackStyle       =   0  'Transparent
      Caption         =   " TCS  Amount    :"
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
      Left            =   8280
      TabIndex        =   172
      Top             =   8325
      Width           =   1530
   End
   Begin VB.Label Label25 
      BackStyle       =   0  'Transparent
      Caption         =   "TCS Rate      :"
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
      Left            =   5985
      TabIndex        =   171
      Top             =   8325
      Width           =   1260
   End
   Begin VB.Label LblBrkMobile 
      BackStyle       =   0  'Transparent
      Caption         =   "BrkMobile"
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
      Left            =   5670
      TabIndex        =   169
      Top             =   2385
      Width           =   1800
   End
   Begin VB.Label lblHsnCode 
      BackStyle       =   0  'Transparent
      Caption         =   "HSN"
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
      Left            =   1935
      TabIndex        =   155
      Top             =   6435
      Width           =   2745
   End
   Begin VB.Label lblStateCd 
      BackStyle       =   0  'Transparent
      Caption         =   "Mobile"
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
      Left            =   5670
      TabIndex        =   154
      Top             =   1890
      Width           =   1800
   End
   Begin VB.Label Label22 
      BackStyle       =   0  'Transparent
      Caption         =   "Mobile :"
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
      Left            =   4905
      TabIndex        =   153
      Top             =   1890
      Width           =   810
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
      Left            =   1485
      TabIndex        =   152
      Top             =   1890
      Width           =   3780
   End
   Begin VB.Label Label21 
      BackStyle       =   0  'Transparent
      Caption         =   "H S N Code    :"
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
      Height          =   330
      Left            =   405
      TabIndex        =   151
      Top             =   6435
      Width           =   1500
   End
   Begin VB.Label Label20 
      BackStyle       =   0  'Transparent
      Caption         =   "Consi.     :"
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
      TabIndex        =   150
      Top             =   3600
      Width           =   1080
   End
   Begin VB.Line Line1 
      BorderColor     =   &H000000C0&
      X1              =   225
      X2              =   7560
      Y1              =   3420
      Y2              =   3420
   End
   Begin VB.Label Label19 
      BackStyle       =   0  'Transparent
      Caption         =   "GSTIN No :"
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
      TabIndex        =   149
      Top             =   1890
      Width           =   1080
   End
   Begin VB.Label Label16 
      BackStyle       =   0  'Transparent
      Caption         =   "I  GST  Amount  :"
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
      Left            =   8235
      TabIndex        =   148
      Top             =   7845
      Width           =   1530
   End
   Begin VB.Label Label10 
      BackStyle       =   0  'Transparent
      Caption         =   "I  GST Rate   :"
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
      Left            =   5985
      TabIndex        =   147
      Top             =   7875
      Width           =   1260
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
      Left            =   8235
      TabIndex        =   146
      Top             =   7335
      Width           =   1605
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "S GST Rate  :"
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
      Left            =   5985
      TabIndex        =   145
      Top             =   7335
      Width           =   1260
   End
   Begin VB.Label Label4 
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
      Left            =   8235
      TabIndex        =   144
      Top             =   6840
      Width           =   1710
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "C GST Rate  :"
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
      Left            =   5985
      TabIndex        =   143
      Top             =   6840
      Width           =   1305
   End
   Begin VB.Label Label13 
      BackStyle       =   0  'Transparent
      Caption         =   "ERN. No     :"
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
      Left            =   7800
      TabIndex        =   142
      Top             =   3600
      Width           =   1440
   End
   Begin VB.Line Line2 
      BorderColor     =   &H000000C0&
      X1              =   7695
      X2              =   12240
      Y1              =   2385
      Y2              =   2385
   End
   Begin VB.Label Label38 
      BackStyle       =   0  'Transparent
      Caption         =   "M.GP No    : "
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
      Left            =   7800
      TabIndex        =   141
      Top             =   1935
      Width           =   1260
   End
   Begin VB.Label Label17 
      BackStyle       =   0  'Transparent
      Caption         =   "Taxable Amount    :"
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
      Left            =   7905
      TabIndex        =   69
      Top             =   6300
      Width           =   1950
   End
   Begin VB.Line Line3 
      BorderColor     =   &H000000C0&
      X1              =   255
      X2              =   4905
      Y1              =   8010
      Y2              =   8010
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
      ForeColor       =   &H00800000&
      Height          =   375
      Left            =   405
      TabIndex        =   68
      Top             =   8865
      Width           =   1125
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
      Left            =   10575
      TabIndex        =   66
      Top             =   4500
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
      Left            =   9120
      TabIndex        =   65
      Top             =   4500
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
      Left            =   8115
      TabIndex        =   64
      Top             =   4500
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
      Left            =   6510
      TabIndex        =   63
      Top             =   4500
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
      TabIndex        =   59
      Top             =   4500
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
      Left            =   5460
      TabIndex        =   62
      Top             =   4500
      Width           =   630
   End
   Begin VB.Shape Shape5 
      BorderColor     =   &H00C00000&
      BorderWidth     =   5
      FillColor       =   &H00C00000&
      Height          =   10470
      Left            =   0
      Top             =   -270
      Width           =   14070
   End
   Begin VB.Label Label39 
      BackStyle       =   0  'Transparent
      Caption         =   "Purch No         :"
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
      Left            =   405
      TabIndex        =   72
      Top             =   7065
      Width           =   1515
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
      TabIndex        =   60
      Top             =   5445
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
      ForeColor       =   &H000000FF&
      Height          =   240
      Left            =   7800
      TabIndex        =   53
      Top             =   2610
      Width           =   1125
   End
   Begin VB.Label Label32 
      BackStyle       =   0  'Transparent
      Caption         =   "Lot No             : "
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
      Left            =   405
      TabIndex        =   74
      Top             =   7545
      Width           =   1680
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
      Left            =   3915
      TabIndex        =   61
      Top             =   4493
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
      ForeColor       =   &H00800000&
      Height          =   375
      Left            =   405
      TabIndex        =   67
      Top             =   8190
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
      ForeColor       =   &H000000FF&
      Height          =   330
      Left            =   8235
      TabIndex        =   71
      Top             =   9450
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
      Left            =   11205
      TabIndex        =   109
      Top             =   0
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
      ForeColor       =   &H00800000&
      Height          =   330
      Left            =   9225
      TabIndex        =   70
      Top             =   8865
      Width           =   630
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
      ForeColor       =   &H00800000&
      Height          =   375
      Left            =   7800
      TabIndex        =   56
      Top             =   3105
      Width           =   1440
   End
   Begin VB.Label Label9 
      BackStyle       =   0  'Transparent
      Caption         =   "Sale A/C :"
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
      Top             =   675
      Width           =   1470
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "&GP No       : "
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
      Left            =   7800
      TabIndex        =   43
      Top             =   1433
      Width           =   1320
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "&Broker      :"
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
      TabIndex        =   57
      Top             =   2355
      Width           =   1125
   End
   Begin VB.Label LblMill 
      BackStyle       =   0  'Transparent
      Caption         =   "M&ill          :"
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
      TabIndex        =   58
      Top             =   2910
      Width           =   1050
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Par&ty       :"
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
      TabIndex        =   54
      Top             =   1380
      Width           =   1080
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
      TabIndex        =   107
      Top             =   0
      Width           =   14040
   End
   Begin VB.Shape Shape1 
      FillColor       =   &H00BDDFD1&
      FillStyle       =   0  'Solid
      Height          =   7905
      Left            =   12450
      Shape           =   4  'Rounded Rectangle
      Top             =   930
      Width           =   1365
   End
   Begin VB.Shape Shape4 
      BorderColor     =   &H000000C0&
      FillColor       =   &H00F5F5DE&
      FillStyle       =   0  'Solid
      Height          =   2910
      Left            =   240
      Top             =   1215
      Width           =   7365
   End
   Begin VB.Shape Shape9 
      BorderColor     =   &H000000C0&
      FillColor       =   &H00EDFDFE&
      FillStyle       =   0  'Solid
      Height          =   3225
      Left            =   240
      Top             =   6120
      Width           =   4665
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H00000080&
      FillColor       =   &H00EDFDFE&
      FillStyle       =   0  'Solid
      Height          =   585
      Left            =   240
      Top             =   525
      Width           =   11940
   End
   Begin VB.Shape Shape10 
      BorderColor     =   &H000000C0&
      FillColor       =   &H00EDFDFE&
      FillStyle       =   0  'Solid
      Height          =   2910
      Left            =   7695
      Top             =   1215
      Width           =   4530
   End
   Begin VB.Shape Shape6 
      BorderColor     =   &H000000C0&
      FillColor       =   &H00EDFDFE&
      FillStyle       =   0  'Solid
      Height          =   1620
      Left            =   240
      Top             =   4305
      Width           =   12045
   End
   Begin VB.Shape Shape7 
      BorderColor     =   &H00000080&
      FillColor       =   &H00F5F5DE&
      FillStyle       =   0  'Solid
      Height          =   3795
      Left            =   5445
      Top             =   6120
      Width           =   6840
   End
End
Attribute VB_Name = "frmSalesGST"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Const tblName As String = "tblAddLess"
Const IdField As String = "Vno"
Const OrderField As String = "Vdt,Vno"
Const MaxNo As Long = 58
Dim SalesAccode As Long
Dim rstbl As Recordset '-- TblAddLess
Dim rstblSub As Recordset '-- TblVouSub
Dim varBookMark As Variant
Dim FirstTimeFlag As Boolean
Dim rsVou As Recordset '--TblVoucher
Dim rsOS As Recordset '--TblOutStanding
Dim OldgFrmTypeStr As String
Dim FormAction As Integer
Dim frmTypeStr As String
Dim modflag As Boolean
Dim ItemCat As Integer
Dim Stat As String
Dim MilCd As Long
Dim SITSalesMillAccode As Long

Private Sub chkIsIncludingGst_Click()
CalAmount
End Sub

'Dim GridFlag As Boolean
Private Sub cmdBtn_Click(Index As Integer)
'On Error GoTo ErrorRoutine
Dim X As String
Dim OrderByStr As String
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
    If gBackEndDB = gBackEndAccess Then
        OrderByStr = "IIf(IsNumeric(Trim(adinvno))=-1,Clng(adinvno),clng(mid(adinvno,5))) ,vdt"
    Else
        OrderByStr = "decode(InStr(adinvno,'-',4),0,to_number(adinvno),to_number(substr(AdInVNo,5))),vdt"
    End If
    If frmTypeStr = "ST" Then
        gClsSearch.SearchMultiField "tbladdLess,TblMastAccount,TblMastAccount TblMastAc1", "AdInvNo,Vno,Vdt,TblMastAccount.AcName,TblMastAc1.AcName,ADTMP7,AdLRNo,AdbillAmt,AdGpNo,AdGpDt", Array("Invoice No", "Vno", "Date", "Party", "Sales A/c", "M.G.P.No.", "L.R. No", "Amount", "Gp No", "Gp Date"), Array(1000, 0, 1200, 2500, 1500, 1000, 2000, 1500, 1000, 1500), " AdAcDrCode=TblMastAccount.AcCode and AdAcCrcode=TblMastAc1.AcCode and " & IIf(SalesAccode = 0, "", " TblAddless.AdAcCrCode=" & SalesAccode & " and ") & " vtype='" & frmTypeStr & "' and AdMillCode=" & SITSalesMillAccode & " and VYear=" & gCYear, "", OrderByStr, 0, 0
    
    Else
        gClsSearch.SearchMultiField "tbladdLess,TblMastAccount,TblMastAccount TblMastAc1", "AdInvNo,Vno,Vdt,TblMastAccount.AcName,TblMastAc1.AcName,AdbillAmt,AdGpNo,AdGpDt", Array("Invoice No", "Vno", "Date", "Party", "Sales A/c", "Amount", "Gp No", "Gp Date"), Array(1000, 0, 1200, 2500, 2500, 1500, 1000, 1500), " AdAcDrCode=TblMastAccount.AcCode and AdAcCrcode=TblMastAc1.AcCode and " & IIf(SalesAccode = 0, "", " AdAcCrCode=" & SalesAccode & " and ") & " vtype='" & frmTypeStr & "' and VYear=" & gCYear, "", OrderByStr, 0, 0
    End If
    If gClsSearch.SearchMultiRetCol(0) <> "" Then
        rstbl.MoveFirst
        rstbl.Find "Vno = " & gClsSearch.SearchMultiRetCol(1)
        gClsSearch.SearchMultiRetCol(0) = ""
        ReadFields
    End If
Case 6 'Add
    txt(55) = "A"
    Dim mTmpTxtSaleAc As String
    mTmpTxtSaleAc = txtTemp(3)
    FormAction = vbDataActionAddNew
    Call GProcActivateControls(Me, True)
    SetControlEd
    Call GProcSetButtons(Me, vbDataActionAddNew)
    Call GProcClearForm(Me, rstbl, MaxNo, True)
    If FirstTimeFlag = True Then
        FirstTimeFlag = False
    Else
       txt(6).SetFocus
    End If
    txt(0) = GProcGenerateId("TblAddLess", "Vno", "Vtype='" & frmTypeStr & "' and VYear=" & gCYear)
    mskDt(2) = GProcGenerateVDt("tblAddLess", "Vdt", "Vtype='" & frmTypeStr & "' and VYear=" & gCYear)
    '-- Sales A/C
    txtTemp(3) = GProcGetColumnValue("tblMastAccount", "AcCode", CStr(SalesAccode), "N", "AcName", "S")
    txtTemp_LostFocus (3)
    'If (Left(mTmpTxtSaleAc, 3) <> "txt") Then txtTemp(3) = mTmpTxtSaleAc
Case 7 'Modify
    FormAction = vbDataActionUpdate
    Call GProcChkPass(frmTypeStr, FormAction, txt(36), txt(2), txtTemp(4), gUserName, "")
    If gChkPassTrueFalse = True Then
        txt(55) = "M"
        If IsRelatedRecord = False Then   '--- Or modflag = True
            Call GProcActivateControls(Me, True)
            txt(6).SetFocus
            SetControlEd
            txttempOS(6).Enabled = True '-- Broker
            txtTemp(4).Enabled = True '--Party
        Else
            Call GProcActivateControls(Me, False)
            SetControlEd
            txttempOS(6).Enabled = True '-- Broker
        End If
        Call GProcSetButtons(Me, vbDataActionUpdate)
        txt(36).Enabled = False
    Else
       FormAction = vbDataActionClose
    End If
Case 8 'Delete
    'If mURecDel = True Then
    Call GProcChkPass(frmTypeStr, FormAction, txt(36), txt(2), txtTemp(4), gUserName, "")
    If gChkPassTrueFalse = True Then
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
        txtPInvno(0) = Mid(txt(36).text, 5)
        txtPInvno(1) = Mid(txt(36).text, 5)
    Else
        txtPInvno(0) = txt(36).text
        txtPInvno(1) = txt(36).text
    End If
    Frame1.Visible = True
    DoEvents
    txtPInvno(0).Enabled = True
    txtPInvno(1).Enabled = True
    optInvDate(0).Enabled = True
    optInvDate(1).Enabled = True
    mskDate.Enabled = True
    txtPInvno(0).SetFocus
    OptOrg.Enabled = True
    OptDup.Enabled = True
    Check1.Enabled = True
    ChkDc.Enabled = True
Case 10 'Exit
'    DoEvents
    Unload Me
    Exit Sub
Case 13 'Save
    If ValidateData = True Then
        If FormAction = vbDataActionAddNew Then txt(0) = GProcGenerateId(tblName, IdField, "vtype='" & frmTypeStr & "' and VYear=" & gCYear)
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
Dim ExemptMill As Long
txtCrBalance.Enabled = False
txtDrBalance.Enabled = False
mskDt(7).Enabled = False '--GatePass Date
txt(10).Enabled = False '--Ass Value
txt(17).Enabled = False '--RofF
txt(57).Enabled = False
txt(58).Enabled = False


txtTemp(4).Enabled = False '--Party
txtTemp(5).Enabled = False '--Mill
txtTemp(3).Enabled = False '--Sales A/c
txtTemp(12).Enabled = True '--Consinee

txttmpInNo.Enabled = False '--Purchase No
'--- V Sub
txttmpVSub(4).Enabled = False '--count
txtVsub(5).Enabled = False '--Bag
txtVsub(6).Enabled = False '--Wt
txtVsub(7).Enabled = False '--Rate
txtVsub(9).Enabled = False '--Sub Amt
txtVsub(12).Enabled = False '--Inward no
txtVsub(10).Enabled = False '--Lot No
txtVsub(48).Enabled = False '--CGST Amt
txtVsub(50).Enabled = False '--SGST Amt
txtVsub(52).Enabled = False '--IGST Amt
txttmpVSub(17).Enabled = False '--Godown
'-- Out Standing
txtOS(9).Enabled = True '--due Days
mskOS(10).Enabled = True '--Due Dt
txttempOS(6).Enabled = False '--Broker
'--- Mill
txtTemp(5).Visible = True
LblMill.Visible = True
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
End Sub
Private Sub SaveRelatedRecord()
Dim i As Long
Dim MillCode As Long
Dim rsVouSub As Recordset
Dim rsAddLessGP As Recordset
Dim rsVouSubGp As Recordset
Dim CrCode As Long
Dim DrCode As Long
Dim BrkCode As Long
Dim TaxAcCodeCGST As Long
Dim TaxAcCodeSGST As Long
Dim TaxAcCodeIGST As Long
Dim ItCode As Long
Dim X As String
Dim TaxAmt As Double
Dim BillDt
Dim DueDt
'--- Account Code
MillCode = GProcGetColumnValue("tblMastAccount", "AcName", txtTemp(5), "S", "AcCode", "N")
CrCode = GProcGetColumnValue("tblMastAccount", "AcName", txtTemp(3), "S", "AcCode", "N")
DrCode = GProcGetColumnValue("tblMastAccount", "AcName", txtTemp(4), "S", "AcCode", "N")
BrkCode = GProcGetColumnValue("tblMastAccount", "AcName", txttempOS(6), "S", "AcCode", "N")
TaxAcCodeCGST = GProcGetColumnValue("tblMastNarration", "Narration", "G S T", "S", "CGSTInPutAc", "N")
TaxAcCodeSGST = GProcGetColumnValue("tblMastNarration", "Narration", "G S T", "S", "SGSTInPutAc", "N")
TaxAcCodeIGST = GProcGetColumnValue("tblMastNarration", "Narration", "G S T", "S", "IGSTInPutAc", "N")

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
rsVouSub!vsubBAG = CLng(txtVsub(5))  '--Bag(5)
rsVouSub!vsubwt = CDbl(txtVsub(6)) '-- Wt(6)
rsVouSub!vsubRt = CDbl(txtVsub(7)) '--Rate(7)
rsVouSub!vsubRTPER = CDbl(txtVsub(8)) '--Rate Per(8)
rsVouSub!vsubAMT = CDbl(txtVsub(9)) '--Amt(9)
rsVouSub!vsubLOTNO = txtVsub(10) & vbNullString '--Lot No(10)
rsVouSub!VSubPItCtrlNo = txtVsub(12) '--PIT Ctrl No(12)
'--Godown Code(17)
rsVouSub!VSubGodownCode = GProcGetColumnValue("tblMastNarration", "Narration", txttmpVSub(17), "S", "NarrCode", "N", "")
rsVouSub!VSubItCtrlNo = GProcGenerateId("tblVouSub", "VSubItCtrlNo", "VYear=" & gCYear) '--It Ctrl No
rsVouSub!vsubNetRate = txtVsub(22) & vbNullString '--Net Rate
rsVouSub!VYear = txt(23) '--VYear(18)
rsVouSub!VsubCGSTRt = CDbl(txtVsub(47)) '--CGST Rt
rsVouSub!VsubCGSTAmt = CDbl(txtVsub(48)) '--CGST Amt
rsVouSub!VsubSGSTRt = CDbl(txtVsub(49)) '--SGST Rt
rsVouSub!VsubSGSTAmt = CDbl(txtVsub(50)) '--SGST Amt
rsVouSub!VsubIGSTRt = CDbl(txtVsub(51)) '--IGST Rt
rsVouSub!VsubIGSTAmt = CDbl(txtVsub(52)) '--IGST Amt
rsVouSub!VSubTaxableAmt = CDbl(txt(10)) '--Taxable Amt
rsVouSub.Update
rsVouSub.Close
'======================Voucher Creation
    '-- Bill Amt
TaxAmt = Val(txtVsub(48)) + Val(txtVsub(50)) + Val(txtVsub(52)) + Val(txt(58))
    
GProcCreateVoucher FormAction, txt(0), 1, txt(2), txt(1), CrCode, DrCode, txt(11) - TaxAmt, txt(36), mskDt(2)
GProcCreateVoucher FormAction, txt(0), 2, txt(2), txt(1), DrCode, CrCode, -txt(11), txt(36), mskDt(2)
i = 3
If Val(txtVsub(48)) > 0 Then
   GProcCreateVoucher FormAction, txt(0), i, txt(2), txt(1), TaxAcCodeCGST, DrCode, Val(txtVsub(48)), txt(36), mskDt(2)
   i = i + 1
   DoEvents
End If
If Val(txtVsub(50)) > 0 Then
   GProcCreateVoucher FormAction, txt(0), i, txt(2), txt(1), TaxAcCodeSGST, DrCode, Val(txtVsub(50)), txt(36), mskDt(2)
   i = i + 1
   DoEvents
End If
If Val(txtVsub(52)) > 0 Then
   GProcCreateVoucher FormAction, txt(0), i, txt(2), txt(1), TaxAcCodeIGST, DrCode, Val(txtVsub(52)), txt(36), mskDt(2)
   i = i + 1
   DoEvents
End If

'--- For TCS
If Val(txt(58)) > 0 Then
   GProcCreateVoucher FormAction, txt(0), i, txt(2), txt(1), gAcCodeTcsPay, DrCode, Val(txt(58)), txt(36), mskDt(2)
   i = i + 1
   DoEvents
End If


'======================Outstanding Creation
If IsDate(mskOS(10)) Then
    DueDt = mskOS(10)
Else
    DueDt = ""
End If
GProcCreateOutStanding txt(0), CDate(txt(2)), txt(1), txt(36), DrCode, BrkCode, CDbl(txt(11)), CLng(txtOS(9)), DueDt, MillCode, CDbl(txt(58)), txtVsub(5), ItCode

'---Update Party,Broker,Rate in Gp
X = "select * from tblAddless where vtype= 'GP' and vno=" & txt(6) & " and " _
& " VYear=" & gCYear
GProcRstOpen rsAddLessGP, X, "O"
rsAddLessGP!AdAcDrCode = BrkCode '--Broker in GP
rsAddLessGP!AdAcCrCode = DrCode & vbNullString '--Party in GP
rsAddLessGP.Update
rsAddLessGP.Close

X = "select * from tblVouSub where vtype= 'GP' and vno=" & txt(6) & " and " _
& " VYear=" & gCYear
GProcRstOpen rsVouSubGp, X, "O"
rsVouSubGp!vsubRt = CDbl(txtVsub(22)) '--Rate in GP
rsVouSubGp.Update
rsVouSubGp.Close

'-----

End Sub
'-- show Ledger
Private Sub cmdLedger_Click(Index As Integer)
'On Error GoTo ErrorHandler
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
'On Error GoTo ErrorHandler
frmMain.mnurptRSubOutStaSale_Click (1)
frmRptRegi.txtTemp(2) = txtTemp(4) '- Party
Exit Sub
ErrorHandler:
    GProcErrorHandler
End Sub
Private Sub CmdModifyAll_Click()
    modflag = True
    cmdBtn_Click (2) '--First
    Do While Not rstbl.EOF
        cmdBtn_Click (7) '-- Modify
        cmdBtn_Click (13) '-- Save
        cmdBtn_Click (1) '-- Next
       'MsgBox rstbl.Fields("vno")
       If rstbl.EOF = True Then
            modflag = False
            Exit Sub
       End If
    Loop
    modflag = False
End Sub
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case 13:
            If Me.ActiveControl.Name <> "txtGrid" Then Sendkeys "{TAB}"
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
        SalesAccode = gClsSearch.SearchMultiRetCol(2)
    ElseIf frmTypeStr = "SD" Then ' --- Depot Sale
        LblSaleType.Caption = "   Depot Sale"
        SalesAccode = gClsSearch.SearchMultiRetCol(2)
    ElseIf frmTypeStr = "SY" Then
        SalesAccode = gAcCodeSY
    End If
    If SalesAccode = 0 Then
       X = "Select * from " & tblName & " where  Vtype='" & frmTypeStr & "' and VYear=" & gCYear
    Else
       X = "Select * from " & tblName & " where AdAcCrCode=" & SalesAccode & " and Vtype='" & frmTypeStr & "' and VYear=" & gCYear
    End If
    If Trim(OrderField) <> "" Then X = X & " order by " & OrderField
    GProcRstOpen rstbl, X, "O"
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
If rstbl.State = adStateOpen Then rstbl.Close
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
    End Select
End If
End Sub
Private Sub txt_KeyPress(Index As Integer, KeyAscii As Integer)
If KeyAscii = vbKeyReturn Then Exit Sub
KeyAscii = GProcValidateKey(rstbl, Index, KeyAscii, txt(Index).text)
If Index = 6 Then '--Gate Pass No
    PrepareGatePass
    KeyAscii = 0
End If
End Sub
Private Sub PrepareGatePass()
Dim RsG As Recordset
Dim X As String
Dim GPSLType As Long
Dim MillCondStr As String
Dim SitMillName As String
Dim MilCd As Long
Dim ItCd As Long
Dim i As Integer
Set RsG = New Recordset
'If IsDate(mskDt(2)) = True Then
    If frmTypeStr = "SO" Then '--- Consignment
        GPSLType = 1
    ElseIf frmTypeStr = "SD" Then '--Depot Sale
        MillCondStr = " and GPmill='" & txtTemp(3) & "'"
        GPSLType = 2
    Else
        '-- Trade
        GPSLType = 0
        MillCondStr = ""
    End If
    CheckForNumberTmp
    X = "Select GpVno,format(GpVdt,'dd/mm/yyyy'),GPMill,GpParty,PCount,GPBag,GpWt,GpRate,Pvno,PVdt,PLRNo,PLRDt,PLRAmt,HSN from QryInwardVsGatePass where  GPSLType=" & GPSLType & " " & MillCondStr & " and (SAPItCtrlNo is null or ( isnumeric(SAPItCtrlNo) and SAPItCtrlNo=" & txtVsub(12) & " ) )"
    gClsSearch.SearchMultiField X, "", Array("Gp No", "Gp Date", "Mill", "Party", "Count", "Bag", "Wt", "Rate", "Purchase Vno", "Pur. Vdt", "LR No", "LR Date", "LR Amt", "HSN"), Array(800, 800, 1200, 1000, IIf(MillCondStr = "", 2000, 0), 2200, 2000, 1000, 1000, 1000, 1500, 1200, IIf(frmTypeStr = "ST", 1000, 0), IIf(frmTypeStr = "ST", 1200, 0), IIf(frmTypeStr = "ST", 1500, 0), 0), "", "", "GpVdt,GpVNo", 0, 0, False, , True
    If gClsSearch.SearchMultiRetCol(0) <> "" Then
        txt(6) = gClsSearch.SearchMultiRetCol(0)
        mskDt(7) = Format(gClsSearch.SearchMultiRetCol(1), "dd/mm/yyyy")
        If frmTypeStr = "SO" Then
            txtTemp(3) = gClsSearch.SearchMultiRetCol(4) '--Mill
            txtTemp_LostFocus (3)
        End If
        If FormAction = vbDataActionAddNew Then mskDt(2) = CDate(mskDt(7))   '-- Inv Dt
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
                txtVsub(5) = .Fields("GpBag") '--Bag
                txtVsub(6) = .Fields("GpWt") '--Wt
                txtVsub(22) = .Fields("GpRate") '--Net Rate
                txtVsub(8) = .Fields("GpRatePer") '--Rate
                txtVsub(12) = .Fields("GpITCtrlNo") '--Gate Pass It Ctrl .No
                '--Purchase No
                txttmpInNo = GProcGetColumnValue("TblvouSub", "VSubItCtrlNo", .Fields("GpPITCtrlNo"), "N", "VNo", "N")
                txtVsub(10) = .Fields("PLotNo") & vbNullString '--Lot No
                txttmpVSub(17) = .Fields("PGodown") '--Godown
                txtOS(9) = .Fields("GpDueDay") '--due days
                Chk(1) = .Fields("E2") & vbNullString '--E2
                Chk(0) = .Fields("IsTicket") & vbNullString '--E2
'                txtVSub(28) = .Fields("CheRecd") & vbNullString
 '               txt(43) = .Fields("Note1") & vbNullString
                'txt(45) = .Fields("MGpno") & vbNullString '--M.Gp.No.
                lblHsnCode.Caption = .Fields("HSN") & vbNullString '--hsn Code
                chkIsIncludingGst.Value = .Fields("ExGst")
                lblStateCd.Caption = GProcGetColumnValue("tblMastAccount", "Acname", txtTemp(4), "S", "AcFaxNo", "S")
                LblBrkMobile.Caption = GProcGetColumnValue("tblMastAccount", "Acname", txttempOS(6), "S", "AcFaxNo", "S")
                lblGSTINNo.Caption = GProcGetColumnValue("tblMastAccount", "Acname", txtTemp(4), "S", "AcGSTIN", "S")
            
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
KeyAscii = GProcValidateKey(rstblSub, Index, KeyAscii, txtVsub(Index).text)
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
KeyAscii = GProcValidateKey(rsOS, Index, KeyAscii, txtOS(Index).text)
End Sub
Private Sub txt_LostFocus(Index As Integer)
If Val(txtVsub(47)) + Val(txtVsub(49)) + Val(txtVsub(51)) = 0 And txtTemp(4) <> "" And txtTemp(5) <> "" And txttmpVSub(4) <> "" Then
   MilCd = GProcGetColumnValue("tblMastAccount", "AcName", txtTemp(5), "S", "AcCode", "N")
   ItemCat = GProcGetColumnValue("tblMastItem", "ItName", txttmpVSub(4), "S", "ITtype", "N", "ItMillcode=" & MilCd)
   Stat = GProcGetColumnValue("tblMastAccount", "AcName", txtTemp(4), "S", "AcGSTIN", "S")
   If ItemCat = 0 Then
      If Left(Stat, 2) = "" Or Left(Stat, 1) = "U" Or Left(Stat, 2) = "27" Then
         txtVsub(47) = GProcGetColumnValue("tblMastNarration", "Narration", "G S T", "S", "CotCGSTRt", "N")
         txtVsub(49) = GProcGetColumnValue("tblMastNarration", "Narration", "G S T", "S", "CotSGSTRt", "N")
      Else
         txtVsub(51) = GProcGetColumnValue("tblMastNarration", "Narration", "G S T", "S", "CotIGSTRt", "N")
      End If
   Else
      If Left(Stat, 2) = "" Or Left(Stat, 1) = "U" Or Left(Stat, 2) = "27" Then
         txtVsub(47) = GProcGetColumnValue("tblMastNarration", "Narration", "G S T", "S", "PolCGSTRt", "N")
         txtVsub(49) = GProcGetColumnValue("tblMastNarration", "Narration", "G S T", "S", "PolSGSTRt", "N")
      Else
         txtVsub(51) = GProcGetColumnValue("tblMastNarration", "Narration", "G S T", "S", "PolIGSTRt", "N")
      End If
   End If
   End If

CkeckForNumber
SaveBtnEd
CalAmount
End Sub
Private Sub CkeckForNumber()
GProcCheckForNumber Me, rstbl, MaxNo, Array(36)
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
        Case 12 '-- Consinee
            OldgFrmTypeStr = gfrmTypeStr
            gfrmTypeStr = "S"
            GProcShowForm frmMastNarration, frmMain.mnuMstfrm(8), "S"
            gfrmTypeStr = OldgFrmTypeStr
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
    Case 4 '-- Party
        '-- Search Condition
        X = "(Agcode=90017 or GpCode=90017)"
        gClsSearch.SearchMultiField "QryMastaccount", "Account,AcName,AcCity,AgName,AcGSTIN,AcFaxno", Array("Account", "AcName", "City", "Group", "GSTIN", "State Cd"), Array(3000, 0, 2000, 3000, 0, 0), X, Chr(KeyAscii), "acName", LeftPos, TopPos, True
        KeyAscii = 0
        If gClsSearch.SearchMultiRetCol(0) <> "" And gClsSearch.SearchMultiRetCol(1) = "" Then
            txtTemp(Index) = gClsSearch.SearchMultiRetCol(0)
            lblGSTINNo.Caption = gClsSearch.SearchMultiRetCol(4)
            lblStateCd.Caption = gClsSearch.SearchMultiRetCol(5)
        Else
            If gClsSearch.SearchMultiRetCol(1) <> "" Then
                txtTemp(Index) = gClsSearch.SearchMultiRetCol(1)
                lblGSTINNo.Caption = gClsSearch.SearchMultiRetCol(4)
                lblStateCd.Caption = gClsSearch.SearchMultiRetCol(5)
             End If
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
    Case 12 '-- Consinee
        gClsSearch.SearchMultiField "tblMastNarration", "Narration,GodAdd1,GodAdd2,GodCity", Array("Consinee Name", "Add 1", "Add 2", "City"), Array(3000, 2000, 2000, 1500), " NarrType='S'", Chr(KeyAscii), "Narration", 0, 0, True
        KeyAscii = 0
        If gClsSearch.SearchMultiRetCol(0) <> "" Then
            txtTemp(Index) = gClsSearch.SearchMultiRetCol(0)
        End If
End Select
End Sub
Private Sub txtTemp_LostFocus(Index As Integer)
Dim AcCode As Long
SaveBtnEd
ShowAccountBalance
If (Index = 3 And FormAction = vbDataActionAddNew) Or txt(36) = "" Then
     AcCode = GProcGetColumnValue("TblMastAccount", "AcName", txtTemp(3), "S", "AcCode", "N")
     If frmTypeStr = "ST" Then
        txt(36) = GProcGenerateInvNo(frmTypeStr, SITSalesMillAccode) '--Inv No
     Else
        txt(36) = GProcGenerateInvNo(frmTypeStr, AcCode) '--Inv No
     End If
End If
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
    gClsSearch.SearchMultiField "QryMastaccount", "Account,AcName,AcFaxNo", Array("Broker", "AcName", "FaxNo"), Array(txttempOS(Index).Width, 0, 0), "(AgCode=90016 or GpCode=90016)", Chr(KeyAscii), "AcName", LeftPos, TopPos, True
    KeyAscii = 0
    If gClsSearch.SearchMultiRetCol(0) <> "" And gClsSearch.SearchMultiRetCol(1) = "" Then
        txttempOS(Index) = gClsSearch.SearchMultiRetCol(0)
        LblBrkMobile.Caption = gClsSearch.SearchMultiRetCol(2)
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
    If txtTemp(12) = "" Then txtTemp(12) = "-"
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
    If frmTypeStr = "SY" Then
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
    '---- Amount
    If CDbl(txt(11)) <= 0 Then
        MsgBox "Check Voucher Amount.", vbInformation + vbOKOnly, Me.Caption
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
If IsDate(mskDt(34)) Then '-- ERN dt
    txt(34) = CDate(mskDt(34))
Else
    txt(34) = ""
End If
txt(40) = Chk(1)
txt(41) = Chk(0)
'-- Cr Code
txt(3) = GProcGetColumnValue("tblMastAccount", "AcName", txtTemp(3), "S", "AcCode", "N")
'-- Dr Code
txt(4) = GProcGetColumnValue("tblMastAccount", "AcName", txtTemp(4), "S", "AcCode", "N")
'--Mill
txt(5) = GProcGetColumnValue("tblMastAccount", "AcName", txtTemp(5), "S", "AcCode", "N")
'--Consinee
txt(12) = GProcGetColumnValue("tblMastNarration", "Narration", txtTemp(12), "S", "NarrCode", "N")
txt(23) = gCYear
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
If IsDate(txt(34)) Then '--ERN Dt
    mskDt(34) = txt(34)
Else
    mskDt(34) = "__/__/____"
End If
Chk(1) = txt(40)
Chk(0) = txt(41)

'-- Cr Code
txtTemp(3) = GProcGetColumnValue("tblMastAccount", "AcCode", txt(3), "N", "AcName", "S")
'-- Dr Code
txtTemp(4) = GProcGetColumnValue("tblMastAccount", "AcCode", txt(4), "N", "AcName", "S")
lblGSTINNo.Caption = GProcGetColumnValue("tblMastAccount", "AcCode", txt(4), "N", "AcGSTIN", "S")
lblStateCd.Caption = GProcGetColumnValue("tblMastAccount", "AcCode", txt(4), "N", "AcFaxNo", "S")
'LblBrkMobile.Caption = GProcGetColumnValue("tblMastAccount", "Acname", txttempOS(6), "S", "AcFaxNo", "S")

'--Mill
txtTemp(5) = GProcGetColumnValue("tblMastAccount", "AcCode", txt(5), "N", "AcName", "S")
'--Consinee
txtTemp(12) = GProcGetColumnValue("tblMastNarration", "NarrCode", txt(12), "N", "Narration", "S")
'-- Fill Related Record
FillRelatedRecord
'--- Total+ Debit Amt
SetControlEd
ShowAccountBalance
LblBrkMobile.Caption = ""
LblBrkMobile.Caption = GProcGetColumnValue("tblMastAccount", "Acname", txttempOS(6), "S", "AcFaxNo", "S")

End Sub
Private Sub FillRelatedRecord()
Dim Pu_ItCtrlNo As Long
'--- Fill Related Control
SetRelatedRS
'==== Vou Sub
txtVsub(5) = Format(rstblSub.Fields("VSubBag"), PStr) '--Bag
txtVsub(6) = Format(rstblSub.Fields("VSubWt"), WtStr) '--Wt
txtVsub(7) = Format(rstblSub.Fields("VSubRt"), LStr) '--Rate
txtVsub(8) = Format(rstblSub.Fields("VSubRtPer"), FStr) '--Rate Per
txtVsub(9) = Format(rstblSub.Fields("VSubAmt"), FStr) '--Amt
txtVsub(22) = Format(rstblSub.Fields("VSubNetRate"), FStr) '--Net Rate
txtVsub(10) = rstblSub.Fields("VSubLotNo") & vbNullString  '--Lot No
txtVsub(12) = rstblSub.Fields("VSubPItCtrlNo") '--PIt Ctrl No
txtVsub(47) = Format(rstblSub.Fields("VsubCGSTRt"), FStr) '--CGST Rt
txtVsub(48) = Format(rstblSub.Fields("VsubCGSTAmt"), FStr) '--CGST Amt
txtVsub(49) = Format(rstblSub.Fields("VsubSGSTRt"), FStr) '--SGST Rt
txtVsub(50) = Format(rstblSub.Fields("VsubSGSTAmt"), FStr) '--SGST Amt
txtVsub(51) = Format(rstblSub.Fields("VsubIGSTRt"), FStr) '--IGST Rt
txtVsub(52) = Format(rstblSub.Fields("VsubIGSTAmt"), FStr) '--IGST Amt


Pu_ItCtrlNo = GProcGetColumnValue("TblvouSub", "VSubItCtrlNo", txtVsub(12), "N", "VSubPItCtrlNo", "N", " vtype='GP'")
txttmpInNo = GProcGetColumnValue("TblvouSub", "VSubItCtrlNo", CStr(Pu_ItCtrlNo), "N", "VNo", "N")
'--Item
txttmpVSub(4) = GProcGetColumnValue("tblMastItem", "ItCode", rstblSub.Fields("VSubITCode"), "N", "ItName", "S")
lblHsnCode.Caption = GProcGetColumnValue("tblMastItem", "ItCode", rstblSub.Fields("VSubITCode"), "N", "ITTMP2", "S")

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
GProcShowRecord Me, rstbl, MaxNo
FillTempFromTxt
End Sub
Private Sub CheckForNumberTmp()
Dim i As Long
For i = 6 To 9
If IsNumeric(txtVsub(i)) = False Then txtVsub(i) = "0.00"
Next i
If IsNumeric(txtVsub(5)) = False Then txtVsub(5) = "0"
If IsNumeric(txtVsub(22)) = False Then txtVsub(22) = "0.00"
If IsNumeric(txtVsub(12)) = False Then txtVsub(12) = "0"

If IsNumeric(txtVsub(47)) = False Then txtVsub(47) = "0.00"
If IsNumeric(txtVsub(48)) = False Then txtVsub(48) = "0.00"
If IsNumeric(txtVsub(49)) = False Then txtVsub(49) = "0.00"
If IsNumeric(txtVsub(50)) = False Then txtVsub(50) = "0.00"
If IsNumeric(txtVsub(51)) = False Then txtVsub(51) = "0.00"
If IsNumeric(txtVsub(52)) = False Then txtVsub(52) = "0.00"

End Sub
Private Sub CalAmount()
Dim CTaxAmt As Double
Dim STaxAmt As Double
Dim ITaxAmt As Double
Dim SubAmt As Double
Dim Amt As Double
Dim Rate As Double
Dim ROff As Double
Dim AssVal As Double
Dim X As String
CheckForNumberTmp
CkeckForNumber
'-- Due Date
If IsNumeric(txtOS(9)) = False Then txtOS(9) = "0"
txtOS(9) = Format(txtOS(9), PStr)
If IsDate(mskDt(2)) Then mskOS(10) = DateAdd("d", txtOS(9), mskDt(2))
'-------------
'-- Rate Per
If txtVsub(8) = 0 Then txtVsub(8) = 1
If chkIsIncludingGst.Value = 1 Then
    SubAmt = (CDbl(txtVsub(6)) * CDbl(txtVsub(22)) / CDbl(txtVsub(8)))
    SubAmt = Format(SubAmt, FStr)

Else
    '---- Amt =[Wt* Net Rate/Rate Per]
    Amt = (CDbl(txtVsub(6)) * CDbl(txtVsub(22)) / CDbl(txtVsub(8)))
    Amt = GProcMakeRounding(Amt)
    '--Sub Amt=(Amt)*100/(100+Tax %)
    SubAmt = Amt * 100 / (100 + CDbl(txtVsub(47)) + CDbl(txtVsub(49)) + CDbl(txtVsub(51)))
    SubAmt = Format(SubAmt, FStr)
    '---Tax Amt= Sub amt*Taxp/100
End If
    '-- CGST
             
    CTaxAmt = Round(SubAmt * CDbl(txtVsub(47)) / 100, 2)
    'CTaxAmt = GProcMakeRounding(CTaxAmt)
    CTaxAmt = Format(CTaxAmt, FStr)
        
    '-- SGST
             
    STaxAmt = Round(SubAmt * CDbl(txtVsub(49)) / 100, 2)
    'STaxAmt = GProcMakeRounding(STaxAmt)
    STaxAmt = Format(STaxAmt, FStr)
        
    '-- IGST
             
    ITaxAmt = Round(SubAmt * CDbl(txtVsub(51)) / 100, 2)
    'ITaxAmt = GProcMakeRounding(ITaxAmt)
    ITaxAmt = Format(ITaxAmt, FStr)

'---TCS
Dim wPtyCd As Long
Dim TcsLimitAmt As Double
Dim TcsOn As Double
Dim TcsRt As Double
Dim TcsAmt As Double
Dim tmpRst As Recordset
Dim i As Integer
If gCIsTcsFirm <> 1 Then
If Format(mskDt(2), "YYYY/MM/dd") >= Format("2020/10/01", "YYYY/MM/dd") Then
   wPtyCd = GProcGetColumnValue("tblMastAccount", "AcName", txtTemp(4), "S", "AcCode", "N")
   X = "Select sum(AdBillAmt) from tblAddLess where AdAcDrCode = " & wPtyCd & "  and vtype in ('SY','SD') and vyear = " & gCYear & " and Format(vdt,'YYYY/MM/dd') <= '" & Format(mskDt(2), "YYYY/MM/dd") & "'  "
   X = X + " and  trim(str(vno))+vtype+str(vyear) <> '" & Trim(Str(txt(0))) & "" & gfrmTypeStr & "" & Str(gCYear) & "'"
   i = GProcRstOpen(tmpRst, X, "R")
   TcsLimitAmt = IIf(IsNull(tmpRst.Fields(0)), 0, tmpRst.Fields(0))
   TcsLimitAmt = TcsLimitAmt + SubAmt + CTaxAmt + STaxAmt + ITaxAmt
   If TcsLimitAmt >= 5000000 Then
      TcsOn = (SubAmt + CTaxAmt + STaxAmt + ITaxAmt)
      TcsRt = 0.075
      TcsAmt = (TcsOn * TcsRt) / 100
      TcsAmt = Format(GProcMakeRounding(TcsAmt), FStr)
   Else
      TcsOn = 0
      TcsRt = 0
      TcsAmt = 0
   End If
Else
   TcsOn = 0
   TcsRt = 0
   TcsAmt = 0
End If

txt(56) = Format(TcsOn, FStr)
txt(57) = Format(TcsRt, WtStr)
txt(58) = Format(TcsAmt, FStr)
    
    
'--ROff = Amt - SubAmt - TaxAmt
Amt = Amt + TcsAmt

ROff = Amt - SubAmt - (CTaxAmt + STaxAmt + ITaxAmt + TcsAmt)
ROff = Format(ROff, FStr)

'--- End TCS
End If



If chkIsIncludingGst.Value = 1 Then
   Amt = SubAmt + (CTaxAmt + STaxAmt + ITaxAmt + TcsAmt)
   Amt = GProcMakeRounding(Amt)
End If



'--ROff = Amt - SubAmt - TaxAmt
'ROff = Amt - SubAmt - (CTaxAmt + STaxAmt + ITaxAmt)
'ROff = Format(ROff, FStr)
'--Rate=SubAmt * Rate Per /Wt
If CDbl(txtVsub(6)) > 0 Then
   Rate = SubAmt * CDbl(txtVsub(8)) / CDbl(txtVsub(6))
   Rate = Format(Rate, LStr)
End If
txtVsub(7) = Rate '--Rate
txtVsub(9) = Format(SubAmt, FStr) '--SubAmt
txtVsub(48) = Format(CTaxAmt, FStr) '--CGST
txtVsub(50) = Format(STaxAmt, FStr) '--SGST
txtVsub(52) = Format(ITaxAmt, FStr) '--IGST
txt(11) = Amt '--Amt
txt(17) = ROff '--round off
txt(10) = SubAmt '--Sub amt
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
Dim InvNo As Variant
Dim DbNtAcName As String
Dim Y As String
Dim Lrd As Recordset
Dim LrString As String
Dim LrStringDt As String
Dim BilNo As String
LrString = ""
Dim RsBl As Recordset
            Dim HttpReq As New WinHttpRequest
            Dim URL As String
            Dim params As String
            Dim params2 As String
            Dim SmsMblNo As String
            Dim SmsString As String

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
'       InvNo = CLng(txtPInvno(0))
        InvNo = txtPInvno(0)
       If frmTypeStr = "ST" Then
          BillSrNo = GProcGetColumnValue("tblMastAccount", "AcName", txtTemp(5), "S", "AcBillSrNo", "S")
       Else
          BillSrNo = GProcGetColumnValue("tblMastAccount", "AcName", txtTemp(3), "S", "AcBillSrNo", "S")
       End If
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
       X = "Select * From tblAddLess where  AdInvNo in " & X & " and VType='" & frmTypeStr & "' and VYear=" & gCYear
     End If

    X = X & " order by vno"
     
     
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
                    .Formulas(16) = "wpan = '" & gCPAN & "'"
                    .Formulas(20) = "wcBnknm = '" & gCBankName & "'"
                    .Formulas(21) = "wcBnkac= '" & gCBankAcNo & "'"
                    .Formulas(22) = "wcrtgs = '" & gCBankRtgsCode & "'"
                    .Formulas(23) = "wCompGstin = '" & gCGSTIN & "'"
                     If OptOrg.Value = True Then
                       .Formulas(24) = "wCopy = 'Original Copy'"
                     ElseIf OptDup.Value = True Then
                        .Formulas(24) = "wCopy = 'Duplicate Copy'"
                     End If
                    If Check1.Value = 1 Then
                        .Formulas(25) = "wMnm = 'Y'"
                    Else
                       .Formulas(25) = "wMnm = 'N'"
                    End If

                    .SelectionFormula = "{tblAddLess.AdInvNo} = '" & RsSA!AdInVNo & "' and {tblAddLess.VType}='" & RsSA!Vtype & "' and {tblAddLess.VYear}=" & RsSA!VYear
                    If ChkDc.Value = 1 Then
                         If Dir("D:\HiReports\rptDeliveryChln.rpt") <> "" Then
                             .ReportFileName = "D:\HiReports\rptDeliveryChln.rpt"
                         Else
                            .ReportFileName = gReportPath & "rptDeliveryChln.rpt"
                         End If
                    Else
                         If Dir("D:\HiReports\rptSaleBillPrintGST.rpt") <> "" Then
                             .ReportFileName = "D:\HiReports\rptSaleBillPrintGST.rpt"
                         Else
                            .ReportFileName = gReportPath & "rptSaleBillPrintGST.rpt"
                         End If
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
Case 4, 6:
       Dim PSide As String
       InvNo = txtPInvno(0)
       If frmTypeStr = "ST" Then
          BillSrNo = GProcGetColumnValue("tblMastAccount", "AcName", txtTemp(5), "S", "AcBillSrNo", "S")
       Else
          BillSrNo = GProcGetColumnValue("tblMastAccount", "AcName", txtTemp(3), "S", "AcBillSrNo", "S")
       End If
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
       LrString = "Select tblAddLess.AdInvNo,tblAddLess.AdGpDt,tblMastAccount.AcName,tblMastItem.Itname,tblMastItem.ItTicket,VSubBag,VSubWt,VSubNetRate,AdBillAmt,tblOutStanding.OutDueDate,tblMastAccount_Br.AcName,tblMastAccount.AcFaxNo,tblMastAccount_Br.AcFaxNo,tblOutStanding.OutDueDays,VSubRtPer " _
                   & "From tblAddLess,tblMastAccount,tblVouSub,tblMastItem,tblOutStanding,tblMastAccount tblMastAccount_Br where tblAddLess.AdAcDrCode = tblMastAccount.AcCode and tblAddLess.Vno = tblVouSub.Vno and tblAddLess.Vtype = tblVouSub.Vtype and tblAddLess.Vyear = tblVouSub.Vyear and " _
                   & " tblAddLess.Vno = tblOutStanding.Vno and tblAddLess.Vtype = tblOutStanding.Vtype and tblAddLess.Vyear = tblOutStanding.Vyear and tblOutStanding.OutBrokerCode = tblMastAccount_Br.AcCode  and  tblVousub.VSubItCode = tblMastItem.Itcode and  AdInvNo in " & X & " and tblAddLess.VType='" & frmTypeStr & "' and tblAddLess.VYear=" & gCYear


     If GProcRstOpen(RsSA, LrString, "R") <= 0 Then
       MsgBox "No Records..."
       Exit Sub
     End If
     RsSA.MoveFirst
     Do While Not RsSA.EOF

        SmsMblNo = RsSA.Fields(11) & vbNullString
        SmsString = "B.No-" + RsSA.Fields(0) + " Dt-" + Trim(CStr(Format(RsSA.Fields(1), "dd/mm/yy"))) & vbLf
        SmsString = SmsString + "Pty-" + Left(RsSA.Fields(2), 30) & vbLf
        SmsString = SmsString + "Cnt-" + Left(RsSA.Fields(3), 30) + "-" + Left(RsSA.Fields(4), 10) & vbLf
        SmsString = SmsString + "Bag-" + CStr(RsSA.Fields(5)) + " WT-" + CStr(RsSA.Fields(6)) + " Kg" & vbLf
        SmsString = SmsString + "Rate-" + CStr(RsSA.Fields(7)) + " Per " + CStr(Int(RsSA.Fields(14))) + " Kg" & vbLf
        If RsSA.Fields(13) = 25 Then
           PSide = "Regular"
        ElseIf RsSA.Fields(13) = 4 Then
           PSide = "Second Day"
        ElseIf RsSA.Fields(13) = 1 Then
           PSide = "Next Day"
        End If
        SmsString = SmsString + "Payment Condition-" + PSide & vbLf
        SmsString = SmsString + "Amt-" + CStr(Int(RsSA.Fields(8))) & vbLf
        SmsString = SmsString + "Due-" + CStr(Format(RsSA.Fields(9), "dd/mm/yy")) & vbLf
        SmsString = SmsString + "Brk-" + Left(RsSA.Fields(10), 25) & vbLf
        SmsString = SmsString + "Firm-" + gCName
        'SmsString = Left(SmsString, 180)
        If SmsMblNo = "" Then
           MsgBox ("Party Mobile No Not Found....")
        Else
        DoEvents
        DoEvents
        DoEvents
        DoEvents
        DoEvents
        DoEvents
        DoEvents
        DoEvents
        DoEvents
          If Index = 6 Then
             SentWhatsApp "Txt", "", SmsString, SmsMblNo
          Else
             URL = "http://alerts.sinfini.com/api/web2sms.php"
             params = "workingkey= Ad583e28d45e23d78611667b060ed2e40&sender=MCROCO&to=" & SmsMblNo & " &message=" & SmsString & ""
             HttpReq.Open "POST", URL, False
             HttpReq.SetRequestHeader "Content-Type", "application/x-www-form-urlencoded"
             HttpReq.Send params
          End If
        End If
        DoEvents
        DoEvents
        DoEvents
        
        SmsMblNo = RsSA.Fields(12)
        If SmsMblNo = "" Then
           MsgBox ("Broker Mobile No Not Found....")
        Else
        DoEvents
        DoEvents
        DoEvents
        DoEvents
        DoEvents
        DoEvents
          If Index = 6 Then
             SentWhatsApp "Txt", "", SmsString, SmsMblNo
          Else
             URL = "http://alerts.sinfini.com/api/web2sms.php"
             params = "workingkey= Ad583e28d45e23d78611667b060ed2e40&sender=MCROCO&to=" & SmsMblNo & " &message=" & SmsString & ""
             HttpReq.Open "POST", URL, False
             HttpReq.SetRequestHeader "Content-Type", "application/x-www-form-urlencoded"
             HttpReq.Send params
          End If
        End If
        DoEvents
        DoEvents
        DoEvents
        RsSA.MoveNext
     Loop
     MsgBox ("All SMS Sent....")
Case 2:
    DoEvents
    Frame1.Visible = False
End Select
End Sub
