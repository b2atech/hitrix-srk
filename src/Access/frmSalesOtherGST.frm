VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form frmSalesOtherGST 
   BackColor       =   &H00C2E3FC&
   BorderStyle     =   0  'None
   Caption         =   "Sales Entry"
   ClientHeight    =   9615
   ClientLeft      =   0
   ClientTop       =   15
   ClientWidth     =   14070
   DrawMode        =   10  'Mask Pen
   FillColor       =   &H000040C0&
   FillStyle       =   0  'Solid
   Icon            =   "frmSalesOtherGST.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   9615
   ScaleWidth      =   14070
   ShowInTaskbar   =   0   'False
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
      Height          =   4245
      Left            =   3465
      TabIndex        =   109
      Top             =   2790
      Visible         =   0   'False
      Width           =   5175
      Begin VB.CommandButton cmdPrintBill 
         Caption         =   "Update IRN"
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
         Index           =   5
         Left            =   3435
         TabIndex        =   157
         Top             =   3510
         Width           =   1455
      End
      Begin VB.CommandButton cmdPrintBill 
         Caption         =   "e-Inv Json"
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
         Index           =   3
         Left            =   225
         TabIndex        =   156
         Top             =   3510
         Width           =   1410
      End
      Begin VB.CommandButton cmdPrintBill 
         Caption         =   "e-Inv. Report"
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
         Left            =   1785
         TabIndex        =   155
         Top             =   3510
         Width           =   1500
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
         Left            =   360
         TabIndex        =   153
         Top             =   2040
         Width           =   1815
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
         Left            =   2640
         TabIndex        =   152
         Top             =   2040
         Width           =   2055
      End
      Begin MSMask.MaskEdBox mskDate 
         Height          =   375
         Left            =   2640
         TabIndex        =   119
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
         TabIndex        =   118
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
         TabIndex        =   117
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
         Left            =   3300
         TabIndex        =   116
         Top             =   2640
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
         Left            =   1845
         TabIndex        =   115
         Top             =   2670
         Width           =   1335
      End
      Begin VB.TextBox txtPInvno 
         Appearance      =   0  'Flat
         Height          =   360
         Index           =   1
         Left            =   2640
         MaxLength       =   7
         TabIndex        =   111
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
         Left            =   345
         TabIndex        =   112
         Top             =   2685
         Width           =   1335
      End
      Begin VB.TextBox txtPInvno 
         Appearance      =   0  'Flat
         Height          =   360
         Index           =   0
         Left            =   2640
         MaxLength       =   7
         TabIndex        =   110
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
         TabIndex        =   114
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
         TabIndex        =   113
         Top             =   1080
         Width           =   1680
      End
   End
   Begin VB.TextBox txtVSub 
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
      Index           =   30
      Left            =   1980
      MaxLength       =   25
      TabIndex        =   8
      Text            =   "txtVSub(30)"
      Top             =   5445
      Width           =   6435
   End
   Begin VB.TextBox txtVSub 
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
      Index           =   29
      Left            =   1980
      MaxLength       =   25
      TabIndex        =   7
      Text            =   "txtVSub(29)"
      Top             =   4950
      Width           =   6435
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
      TabIndex        =   3
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
      TabIndex        =   33
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
      TabIndex        =   13
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
      TabIndex        =   32
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
      TabIndex        =   12
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
      TabIndex        =   31
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
      TabIndex        =   11
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
      TabIndex        =   16
      Text            =   "6"
      Top             =   1425
      Visible         =   0   'False
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
      TabIndex        =   18
      Text            =   "45"
      Top             =   1890
      Visible         =   0   'False
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
      TabIndex        =   133
      Text            =   "49"
      Top             =   10605
      Visible         =   0   'False
      Width           =   330
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
      Index           =   33
      Left            =   2160
      TabIndex        =   42
      Text            =   "33"
      Top             =   11070
      Visible         =   0   'False
      Width           =   1560
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
      TabIndex        =   132
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
      TabIndex        =   131
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
      TabIndex        =   130
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
      TabIndex        =   129
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
      TabIndex        =   128
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
      TabIndex        =   21
      Text            =   "44"
      Top             =   3555
      Visible         =   0   'False
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
      TabIndex        =   127
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
      TabIndex        =   126
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
      TabIndex        =   125
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
      TabIndex        =   124
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
      TabIndex        =   123
      Text            =   "39"
      Top             =   10425
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
      Height          =   405
      Index           =   38
      Left            =   990
      TabIndex        =   122
      Text            =   "38"
      Top             =   8865
      Width           =   3840
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
      Height          =   765
      Index           =   37
      Left            =   990
      TabIndex        =   121
      Text            =   "37"
      Top             =   8100
      Width           =   3840
   End
   Begin VB.CommandButton CmdModifyAll 
      Caption         =   "Modify All Inv"
      Height          =   495
      Left            =   2385
      TabIndex        =   120
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
      TabIndex        =   4
      Text            =   "36"
      Top             =   2550
      Width           =   1635
   End
   Begin VB.CommandButton cmdLedger 
      Caption         =   "Ledger"
      Height          =   375
      Index           =   1
      Left            =   16380
      TabIndex        =   108
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
      TabIndex        =   107
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
      TabIndex        =   106
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
      Top             =   6645
      Visible         =   0   'False
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
      TabIndex        =   66
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
      Top             =   7095
      Visible         =   0   'False
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
      Left            =   14355
      TabIndex        =   25
      Text            =   "txttmpVSub(17)"
      ToolTipText     =   "Press F3 for Creation"
      Top             =   4050
      Visible         =   0   'False
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
      Left            =   14445
      TabIndex        =   28
      Text            =   "txtVSub(22)"
      Top             =   6255
      Visible         =   0   'False
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
      TabIndex        =   9
      Text            =   "txtVSub(9)"
      Top             =   4470
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
      Left            =   14400
      TabIndex        =   30
      Text            =   "txtVSub(8)"
      Top             =   5805
      Visible         =   0   'False
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
      Left            =   14355
      TabIndex        =   29
      Text            =   "txtVSub(7)"
      Top             =   5355
      Visible         =   0   'False
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
      Left            =   14355
      TabIndex        =   27
      Text            =   "txtVSub(6)"
      Top             =   4905
      Visible         =   0   'False
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
      Height          =   420
      Index           =   4
      Left            =   1980
      Locked          =   -1  'True
      TabIndex        =   6
      Text            =   "txttmpVSub(4)"
      ToolTipText     =   "Press F3 for Creation"
      Top             =   4455
      Width           =   6435
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
      Left            =   14355
      TabIndex        =   26
      Text            =   "txtVSub(5)"
      Top             =   4455
      Visible         =   0   'False
      Width           =   675
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
      Index           =   32
      Left            =   2160
      TabIndex        =   41
      Text            =   "32"
      Top             =   10485
      Visible         =   0   'False
      Width           =   2865
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
      Left            =   10800
      TabIndex        =   105
      Text            =   "31"
      Top             =   9990
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
      TabIndex        =   87
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
      TabIndex        =   86
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
      TabIndex        =   85
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
      Left            =   10170
      TabIndex        =   84
      Text            =   "27"
      Top             =   9900
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
      TabIndex        =   83
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
      TabIndex        =   82
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
      TabIndex        =   104
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
      TabIndex        =   103
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
      TabIndex        =   101
      Text            =   "1"
      Top             =   540
      Width           =   420
   End
   Begin VB.TextBox txt 
      Alignment       =   1  'Right Justify
      Height          =   315
      Index           =   2
      Left            =   17460
      TabIndex        =   99
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
      TabIndex        =   98
      Text            =   "5"
      Top             =   1935
      Width           =   525
   End
   Begin VB.TextBox txt 
      Alignment       =   1  'Right Justify
      Height          =   315
      Index           =   3
      Left            =   17460
      TabIndex        =   97
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
      TabIndex        =   91
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
      TabIndex        =   90
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
      TabIndex        =   80
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
      TabIndex        =   72
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
      TabIndex        =   71
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
      TabIndex        =   79
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
      TabIndex        =   70
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
      TabIndex        =   69
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
      TabIndex        =   68
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
      TabIndex        =   78
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
      TabIndex        =   73
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
      TabIndex        =   77
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
      TabIndex        =   76
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
      TabIndex        =   75
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
      TabIndex        =   74
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
      TabIndex        =   24
      Text            =   "txttemp(5)"
      ToolTipText     =   "Press F3 for Creation"
      Top             =   2910
      Visible         =   0   'False
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
      TabIndex        =   23
      Text            =   "txttempOS(6)"
      ToolTipText     =   "Press F3 for Creation"
      Top             =   2355
      Visible         =   0   'False
      Width           =   5970
   End
   Begin VB.TextBox txt 
      Alignment       =   1  'Right Justify
      Height          =   315
      Index           =   4
      Left            =   17460
      TabIndex        =   96
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
      TabIndex        =   2
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
      TabIndex        =   95
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
      TabIndex        =   19
      Text            =   "8"
      Top             =   3060
      Visible         =   0   'False
      Width           =   1635
   End
   Begin VB.TextBox txt 
      Alignment       =   1  'Right Justify
      Height          =   315
      Index           =   9
      Left            =   17460
      TabIndex        =   94
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
      TabIndex        =   10
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
      TabIndex        =   15
      Text            =   "11"
      Top             =   8775
      Width           =   3000
   End
   Begin VB.TextBox txt 
      Alignment       =   1  'Right Justify
      Height          =   315
      Index           =   12
      Left            =   17460
      TabIndex        =   93
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
      TabIndex        =   14
      Text            =   "17"
      Top             =   8235
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
      Left            =   1935
      TabIndex        =   36
      Text            =   "txtOS(9)"
      Top             =   7605
      Visible         =   0   'False
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
      TabIndex        =   88
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
      TabIndex        =   89
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
      TabIndex        =   92
      Text            =   "23"
      Top             =   6210
      Visible         =   0   'False
      Width           =   525
   End
   Begin MSMask.MaskEdBox mskDt 
      Height          =   375
      Index           =   2
      Left            =   10725
      TabIndex        =   5
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
      TabIndex        =   17
      Top             =   1433
      Visible         =   0   'False
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
      Left            =   3330
      TabIndex        =   37
      Top             =   7605
      Visible         =   0   'False
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
      Bindings        =   "frmSalesOtherGST.frx":058A
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
      TabIndex        =   20
      Top             =   3060
      Visible         =   0   'False
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
      TabIndex        =   81
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
      Visible         =   0   'False
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
      TabIndex        =   22
      Top             =   3555
      Visible         =   0   'False
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
      Index           =   1
      Left            =   15705
      TabIndex        =   151
      Top             =   2250
      Visible         =   0   'False
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
   Begin VB.Label Label25 
      BackStyle       =   0  'Transparent
      Caption         =   "ACK :"
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
      Left            =   405
      TabIndex        =   159
      Top             =   8955
      Width           =   540
   End
   Begin VB.Label Label24 
      BackStyle       =   0  'Transparent
      Caption         =   "IRN  :"
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
      Left            =   405
      TabIndex        =   158
      Top             =   8100
      Width           =   540
   End
   Begin VB.Label Label11 
      BackStyle       =   0  'Transparent
      Caption         =   "Note             :"
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
      Height          =   315
      Left            =   360
      TabIndex        =   154
      Top             =   5040
      Width           =   1425
   End
   Begin VB.Label Label23 
      BackStyle       =   0  'Transparent
      Caption         =   "L. R. Date       :"
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
      Left            =   630
      TabIndex        =   150
      Top             =   11070
      Visible         =   0   'False
      Width           =   1515
   End
   Begin VB.Label Label15 
      BackStyle       =   0  'Transparent
      Caption         =   "L. R. No          :"
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
      Left            =   630
      TabIndex        =   149
      Top             =   10530
      Visible         =   0   'False
      Width           =   1515
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
      TabIndex        =   148
      Top             =   6255
      Width           =   2745
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
      Left            =   6480
      TabIndex        =   147
      Top             =   1890
      Width           =   990
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
      Left            =   5310
      TabIndex        =   146
      Top             =   1890
      Width           =   1215
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
      TabIndex        =   145
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
      TabIndex        =   144
      Top             =   6255
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
      TabIndex        =   143
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
      TabIndex        =   142
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
      TabIndex        =   141
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
      TabIndex        =   140
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
      TabIndex        =   139
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
      TabIndex        =   138
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
      TabIndex        =   137
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
      TabIndex        =   136
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
      TabIndex        =   135
      Top             =   3600
      Visible         =   0   'False
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
      TabIndex        =   134
      Top             =   1935
      Visible         =   0   'False
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
      TabIndex        =   62
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
   Begin VB.Label Label45 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Amount   :"
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
      Height          =   360
      Left            =   8370
      TabIndex        =   60
      Top             =   4485
      Width           =   1110
   End
   Begin VB.Label Label41 
      BackStyle       =   0  'Transparent
      Caption         =   "Description   :"
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
      Height          =   315
      Left            =   360
      TabIndex        =   59
      Top             =   4500
      Width           =   1425
   End
   Begin VB.Shape Shape5 
      BorderColor     =   &H00C00000&
      BorderWidth     =   5
      FillColor       =   &H00C00000&
      Height          =   9930
      Left            =   0
      Top             =   -375
      Width           =   14025
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
      TabIndex        =   65
      Top             =   6615
      Visible         =   0   'False
      Width           =   1515
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
      TabIndex        =   67
      Top             =   7095
      Visible         =   0   'False
      Width           =   1680
   End
   Begin VB.Label Label12 
      BackStyle       =   0  'Transparent
      Caption         =   "D&ue Days        :"
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
      TabIndex        =   61
      Top             =   7560
      Visible         =   0   'False
      Width           =   1485
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
      TabIndex        =   64
      Top             =   8865
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
      TabIndex        =   102
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
      TabIndex        =   63
      Top             =   8280
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
      Visible         =   0   'False
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
      Visible         =   0   'False
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
      Visible         =   0   'False
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
      Visible         =   0   'False
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
      Caption         =   "  Other Sale"
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
      TabIndex        =   100
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
      Left            =   225
      Top             =   4275
      Width           =   12000
   End
   Begin VB.Shape Shape7 
      BorderColor     =   &H00000080&
      FillColor       =   &H00F5F5DE&
      FillStyle       =   0  'Solid
      Height          =   3255
      Left            =   5520
      Top             =   6120
      Width           =   6840
   End
End
Attribute VB_Name = "frmSalesOtherGST"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Const tblName As String = "tblAddLess"
Const IdField As String = "Vno"
Const OrderField As String = "Vdt,Vno"
Const MaxNo As Long = 50
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
        gClsSearch.SearchMultiField "tbladdLess,TblMastAccount,TblMastAccount TblMastAc1", "AdInvNo,Vno,Vdt,TblMastAccount.AcName,TblMastAc1.AcName,AdbillAmt,AdGpNo,AdGpDt", Array("Invoice No", "Vno", "Date", "Party", "Sales A/c", "Amount", "Gp No", "Gp Date"), Array(1000, 0, 1200, 2500, 2500, 1500, 1000, 1500), " AdAcDrCode=TblMastAccount.AcCode and AdAcCrcode=TblMastAc1.AcCode and " & IIf(SalesAccode = 0, "", " AdAcCrCode=" & SalesAccode & " and ") & " vtype='" & frmTypeStr & "' and left(adTmp5,3)= 'OTH' and VYear=" & gCYear, "", OrderByStr, 0, 0
    End If
    
    If gClsSearch.SearchMultiRetCol(0) <> "" Then
        rstbl.MoveFirst
        rstbl.Find "Vno = " & gClsSearch.SearchMultiRetCol(1)
        gClsSearch.SearchMultiRetCol(0) = ""
        ReadFields
    End If
Case 6 'Add
    Dim mTmpTxtSaleAc As String
    mTmpTxtSaleAc = txttemp(3)
    FormAction = vbDataActionAddNew
    Call GProcActivateControls(Me, True)
    SetControlEd
    Call GProcSetButtons(Me, vbDataActionAddNew)
    Call GProcClearForm(Me, rstbl, MaxNo, True)
    If FirstTimeFlag = True Then
        FirstTimeFlag = False
    Else
       txttemp(3).SetFocus
    End If
    txt(0) = GProcGenerateId("TblAddLess", "Vno", "Vtype='" & frmTypeStr & "' and VYear=" & gCYear)
    mskDt(2) = GProcGenerateVDt("tblAddLess", "Vdt", "Vtype='" & frmTypeStr & "' and VYear=" & gCYear)
    '-- Sales A/C
'    txttemp(3) = GProcGetColumnValue("tblMastAccount", "AcCode", CStr(SalesAccod), "N", "AcName", "S")
    txtTemp_LostFocus (3)
    'If (Left(mTmpTxtSaleAc, 3) <> "txt") Then txtTemp(3) = mTmpTxtSaleAc
Case 7 'Modify
    FormAction = vbDataActionUpdate
    If IsRelatedRecord = False Then   '--- Or modflag = True
        Call GProcActivateControls(Me, True)
'        txt(6).SetFocus
        SetControlEd
'        txttempOS(6).Enabled = True '-- Broker
        txttemp(4).Enabled = True '--Party
        txttemp(4).SetFocus
    Else
        Call GProcActivateControls(Me, False)
        SetControlEd
        txttempOS(6).Enabled = True '-- Broker
    End If
    Call GProcSetButtons(Me, vbDataActionUpdate)
    txt(36).Enabled = False
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
        txt(43) = gfrmHunkStr
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
txt(17).Enabled = False '--Roff
txt(11).Enabled = False '--Bill amt



'txttemp(4).Enabled = False '--Party
txttemp(5).Enabled = False '--Mill
'txttemp(3).Enabled = False '--Sales A/c
txttemp(12).Enabled = True '--Consinee

txttmpInNo.Enabled = False '--Purchase No
'--- V Sub
'txttmpVSub(4).Enabled = False '--count
txtVSub(5).Enabled = False '--Bag
txtVSub(6).Enabled = False '--Wt
txtVSub(7).Enabled = False '--Rate
'txtVSub(9).Enabled = False '--Sub Amt
txtVSub(12).Enabled = False '--Inward no
txtVSub(10).Enabled = False '--Lot No
txtVSub(48).Enabled = False '--CGST Amt
txtVSub(50).Enabled = False '--SGST Amt
txtVSub(52).Enabled = False '--IGST Amt
txttmpVSub(17).Enabled = False '--Godown
'-- Out Standing
'txtOS(9).Enabled = True '--due Days
'mskOS(10).Enabled = True '--Due Dt
'txttempOS(6).Enabled = False '--Broker
'--- Mill
'txttemp(5).Visible = True
'LblMill.Visible = True
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
MillCode = GProcGetColumnValue("tblMastAccount", "AcName", "-", "S", "AcCode", "N")
CrCode = GProcGetColumnValue("tblMastAccount", "AcName", txttemp(3), "S", "AcCode", "N")
DrCode = GProcGetColumnValue("tblMastAccount", "AcName", txttemp(4), "S", "AcCode", "N")
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
rsVouSub!vsubBAG = CLng(txtVSub(5))  '--Bag(5)
rsVouSub!vsubwt = CDbl(txtVSub(6)) '-- Wt(6)
rsVouSub!vsubRt = CDbl(txtVSub(7)) '--Rate(7)
rsVouSub!vsubRTPER = CDbl(txtVSub(8)) '--Rate Per(8)
rsVouSub!vsubAMT = CDbl(txtVSub(9)) '--Amt(9)
rsVouSub!vsubLOTNO = txtVSub(10) & vbNullString '--Lot No(10)
rsVouSub!VSubPItCtrlNo = txtVSub(12) '--PIT Ctrl No(12)
'--Godown Code(17)
rsVouSub!VSubGodownCode = GProcGetColumnValue("tblMastNarration", "Narration", "-", "S", "NarrCode", "N", "")
rsVouSub!VSubItCtrlNo = GProcGenerateId("tblVouSub", "VSubItCtrlNo", "VYear=" & gCYear) '--It Ctrl No
rsVouSub!vsubNetRate = txtVSub(22) & vbNullString '--Net Rate
rsVouSub!VYear = txt(23) '--VYear(18)
rsVouSub!VsubCGSTRt = CDbl(txtVSub(47)) '--CGST Rt
rsVouSub!VsubCGSTAmt = CDbl(txtVSub(48)) '--CGST Amt
rsVouSub!VsubSGSTRt = CDbl(txtVSub(49)) '--SGST Rt
rsVouSub!VsubSGSTAmt = CDbl(txtVSub(50)) '--SGST Amt
rsVouSub!VsubIGSTRt = CDbl(txtVSub(51)) '--IGST Rt
rsVouSub!VsubIGSTAmt = CDbl(txtVSub(52)) '--IGST Amt
rsVouSub!VSubTaxableAmt = CDbl(txt(10)) '--Taxable Amt
rsVouSub!VSUBTMP6 = txtVSub(29) & vbNullString '--Note 1
rsVouSub!VSUBTMP7 = txtVSub(30) & vbNullString '--Note 2
rsVouSub!VSUBTMP5 = gfrmHunkStr & vbNullString '--Note 2
rsVouSub.Update
rsVouSub.Close
'======================Voucher Creation
    '-- Bill Amt
TaxAmt = Val(txtVSub(48)) + Val(txtVSub(50)) + Val(txtVSub(52))
    
GProcCreateVoucher FormAction, txt(0), 1, txt(2), txt(1), CrCode, DrCode, txt(11) - TaxAmt, txt(36), mskDt(2)
GProcCreateVoucher FormAction, txt(0), 2, txt(2), txt(1), DrCode, CrCode, -txt(11), txt(36), mskDt(2)
i = 3
If Val(txtVSub(48)) > 0 Then
   GProcCreateVoucher FormAction, txt(0), i, txt(2), txt(1), TaxAcCodeCGST, DrCode, Val(txtVSub(48)), txt(36), mskDt(2)
   i = i + 1
   DoEvents
End If
If Val(txtVSub(50)) > 0 Then
   GProcCreateVoucher FormAction, txt(0), i, txt(2), txt(1), TaxAcCodeSGST, DrCode, Val(txtVSub(50)), txt(36), mskDt(2)
   i = i + 1
   DoEvents
End If
If Val(txtVSub(52)) > 0 Then
   GProcCreateVoucher FormAction, txt(0), i, txt(2), txt(1), TaxAcCodeIGST, DrCode, Val(txtVSub(52)), txt(36), mskDt(2)
   i = i + 1
   DoEvents
End If

'======================Outstanding Creation
If IsDate(mskOS(10)) Then
    DueDt = mskOS(10)
Else
    DueDt = ""
End If
DoEvents
DoEvents
DoEvents
DoEvents
DoEvents
'GProcCreateOutStanding txt(0), CDate(txt(2)), txt(1), txt(36), DrCode, BrkCode, CDbl(txt(11)), CLng(txtOS(9)), DueDt, MillCode, CDbl(txt(21)), txtVsub(5), ItCode
'---Update Party,Broker,Rate in Gp
'X = "select * from tblAddless where vtype= 'GP' and vno=" & txt(6) & " and " _
'& " VYear=" & gCYear
'GProcRstOpen rsAddLessGP, X, "O"
'rsAddLessGP!AdAcDrCode = BrkCode '--Broker in GP
'rsAddLessGP!AdAcCrCode = DrCode & vbNullString '--Party in GP
'rsAddLessGP.Update
'rsAddLessGP.Close

'X = "select * from tblVouSub where vtype= 'GP' and vno=" & txt(6) & " and " _
'& " VYear=" & gCYear
'GProcRstOpen rsVouSubGp, X, "O"
'rsVouSubGp!vsubRt = CDbl(txtVsub(22)) '--Rate in GP
'rsVouSubGp.Update
'rsVouSubGp.Close

'-----

End Sub
'-- show Ledger
Private Sub cmdLedger_Click(Index As Integer)
'On Error GoTo ErrorHandler
If Index = 0 Then
   frmRptAccount.txt(1) = txttemp(3)
Else
   frmRptAccount.txt(1) = txttemp(4)
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
frmRptRegi.txttemp(2) = txttemp(4) '- Party
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
'    If frmTypeStr = "SO" Then  ' --- Consignment Sale
'        LblSaleType.Caption = "   Consignment Sale"
'        SalesAccode = gClsSearch.SearchMultiRetCol(2)
'    ElseIf frmTypeStr = "SD" Then ' --- Depot Sale
'        LblSaleType.Caption = "   Depot Sale"
'        SalesAccode = gClsSearch.SearchMultiRetCol(2)
'    ElseIf frmTypeStr = "SY" Then
'        SalesAccode = gAcCodeSY
'    End If
    X = "Select * from " & tblName & " where ADTMP5='" & gfrmHunkStr & "' and Vtype='" & frmTypeStr & "' and VYear=" & gCYear
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
    DrAcBal = GProcGetAccountBalance(txttemp(4), mskDt(2))
    CrAcBal = GProcGetAccountBalance(txttemp(3), mskDt(2))
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
'Private Sub txt_KeyPress(Index As Integer, KeyAscii As Integer)
'If KeyAscii = vbKeyReturn Then Exit Sub
'KeyAscii = GProcValidateKey(rstbl, Index, KeyAscii, txt(Index).text)
'If Index = 6 Then '--Gate Pass No
'    PrepareGatePass
'    KeyAscii = 0
'End If
'End Sub
Private Sub PrepareGatePass()
Dim RsG As Recordset
Dim X As String
Dim GPSLType As Long
Dim MillCondStr As String
Dim SitMillName As String
Dim i As Integer
Set RsG = New Recordset
'If IsDate(mskDt(2)) = True Then
    If frmTypeStr = "SO" Then '--- Consignment
        GPSLType = 1
    ElseIf frmTypeStr = "SD" Then '--Depot Sale
        MillCondStr = " and GPmill='" & txttemp(3) & "'"
        GPSLType = 2
    Else
        '-- Trade
        GPSLType = 0
        MillCondStr = ""
    End If
    CheckForNumberTmp
    X = "Select GpVno,format(GpVdt,'dd/mm/yyyy'),GPMill,GpParty,PCount,GPBag,GpWt,GpRate,Pvno,PVdt,PLRNo,PLRDt,PLRAmt,HSN from QryInwardVsGatePass where  GPSLType=" & GPSLType & " " & MillCondStr & " and (SAPItCtrlNo is null or ( isnumeric(SAPItCtrlNo) and SAPItCtrlNo=" & txtVSub(12) & " ) )"
    gClsSearch.SearchMultiField X, "", Array("Gp No", "Gp Date", "Mill", "Party", "Count", "Bag", "Wt", "Rate", "Purchase Vno", "Pur. Vdt", "LR No", "LR Date", "LR Amt", "HSN"), Array(800, 800, 1200, 1000, IIf(MillCondStr = "", 2000, 0), 2200, 2000, 1000, 1000, 1000, 1500, 1200, IIf(frmTypeStr = "ST", 1000, 0), IIf(frmTypeStr = "ST", 1200, 0), IIf(frmTypeStr = "ST", 1500, 0), 0), "", "", "GpVdt,GpVNo", 0, 0, False, , True
    If gClsSearch.SearchMultiRetCol(1) <> "" Then
        txt(6) = gClsSearch.SearchMultiRetCol(0)
        mskDt(7) = Format(gClsSearch.SearchMultiRetCol(1), "dd/mm/yyyy")
        If frmTypeStr = "SO" Then
            txttemp(3) = gClsSearch.SearchMultiRetCol(4) '--Mill
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
                txttemp(4) = .Fields("GpParty") '--Party
                txttempOS(6) = .Fields("GpBroker") '--Broker
                txttemp(5) = .Fields("GpMill") '--Mill
                txttmpVSub(4) = .Fields("PCount") '--Count
                txtVSub(5) = .Fields("GpBag") '--Bag
                txtVSub(6) = .Fields("GpWt") '--Wt
                txtVSub(22) = .Fields("GpRate") '--Net Rate
                txtVSub(8) = .Fields("GpRatePer") '--Rate
                txtVSub(12) = .Fields("GpITCtrlNo") '--Gate Pass It Ctrl .No
                '--Purchase No
                txttmpInNo = GProcGetColumnValue("TblvouSub", "VSubItCtrlNo", .Fields("GpPITCtrlNo"), "N", "VNo", "N")
                txtVSub(10) = .Fields("PLotNo") & vbNullString '--Lot No
                txttmpVSub(17) = .Fields("PGodown") '--Godown
                txtOS(9) = .Fields("GpDueDay") '--due days
    '            txtVSub(28) = .Fields("CheRecd") & vbNullString
    '            txt(43) = .Fields("Note1") & vbNullString
    '            txt(45) = .Fields("MGpno") & vbNullString '--M.Gp.No.
                lblHsnCode.Caption = .Fields("HSN") & vbNullString '--hsn Code
                txt(32) = .Fields("PLRNo") & vbNullString '--LR No
                If IsDate(.Fields("PLRDt")) = True Then
                    mskDt(33) = Format(.Fields("PLRDt"), "dd/mm/yyyy") '--LR dt
                    txt(33) = Format(.Fields("PLRDt"), "dd/mm/yyyy") '--LR dt
                Else
                    mskDt(33) = "__/__/____"
                    txt(33) = ""  '--LR dt
                End If
            End If
          RsG.Close
        End With
    End If
'End If
SetControlEd
CalAmount
ShowAccountBalance
lblGSTINNo.Caption = GProcGetColumnValue("tblMastAccount", "AcName", txttemp(4), "S", "AcGSTIN", "S")

End Sub
Private Sub txttmpVSub_KeyPress(Index As Integer, KeyAscii As Integer)
gClsSearch.SearchMultiField "tblMastItem", "ItName,ItCode,ITTMP2", Array("Item", "Item", ""), Array(2000, 0, 0), " tblMastItem.ItOtherItem=1", "", "ItName", txttmpVSub(4).Left + Me.Left, txttmpVSub(4).Top + Me.Top + 650
KeyAscii = 0
If gClsSearch.SearchMultiRetCol(0) <> "" Then
   txttmpVSub(Index) = gClsSearch.SearchMultiRetCol(0)
   lblHsnCode.Caption = gClsSearch.SearchMultiRetCol(2)
   txtVSub(29).SetFocus
End If
End Sub

Private Sub txtVSub_GotFocus(Index As Integer)
GProcSelectBox Me.ActiveControl
End Sub
Private Sub txtVSub_KeyPress(Index As Integer, KeyAscii As Integer)
If KeyAscii = vbKeyReturn Then Exit Sub
KeyAscii = GProcValidateKey(rstblSub, Index, KeyAscii, txtVSub(Index).text)
End Sub
Private Sub txtVSub_LostFocus(Index As Integer)
CheckForNumberTmp
If Index = 29 Or Index = 30 Then
   txtVSub(Index) = GProcProperCase(txtVSub(Index))
End If
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
If Val(txtVSub(47)) + Val(txtVSub(49)) + Val(txtVSub(51)) = 0 And txttemp(4) <> "" And txttemp(5) <> "" And txttmpVSub(4) <> "" Then
   MilCd = GProcGetColumnValue("tblMastAccount", "AcName", txttemp(5), "S", "AcCode", "N")
   ItemCat = GProcGetColumnValue("tblMastItem", "ItName", txttmpVSub(4), "S", "Ittype", "N", "ItMillcode=" & MilCd)
   Stat = GProcGetColumnValue("tblMastAccount", "AcName", txttemp(4), "S", "AcGSTIN", "S")
   If ItemCat = 0 Then
      If Left(Stat, 2) = "" Or Left(Stat, 1) = "U" Or Left(Stat, 2) = "27" Then
         txtVSub(47) = GProcGetColumnValue("tblMastNarration", "Narration", "G S T", "S", "CotCGSTRt", "N")
         txtVSub(49) = GProcGetColumnValue("tblMastNarration", "Narration", "G S T", "S", "CotSGSTRt", "N")
      Else
         txtVSub(51) = GProcGetColumnValue("tblMastNarration", "Narration", "G S T", "S", "CotIGSTRt", "N")
      End If
   Else
      If Left(Stat, 2) = "" Or Left(Stat, 1) = "U" Or Left(Stat, 2) = "27" Then
         txtVSub(47) = GProcGetColumnValue("tblMastNarration", "Narration", "G S T", "S", "PolCGSTRt", "N")
         txtVSub(49) = GProcGetColumnValue("tblMastNarration", "Narration", "G S T", "S", "PolSGSTRt", "N")
      Else
         txtVSub(51) = GProcGetColumnValue("tblMastNarration", "Narration", "G S T", "S", "PolIGSTRt", "N")
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
                GProcShowForm frmMastAccount, frmMain.mnuMstfrm(0), Array(frmMastAccount.txttemp(2)), Array(GName)
            Else
                GName = GProcGetColumnValue("TblMastGroup", "AgCode", 90023, "N", "AgName", "S")
                GProcShowForm frmMastAccount, frmMain.mnuMstfrm(0), Array(frmMastAccount.txttemp(2)), Array(GName)
            End If
        Case 4 '-- Party
            GName = GProcGetColumnValue("TblMastGroup", "AgCode", 90017, "N", "AgName", "S")
            GProcShowForm frmMastAccount, frmMain.mnuMstfrm(0), Array(frmMastAccount.txttemp(2)), Array(GName)
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
TopPos = txttemp(Index).Top + Me.Top + 650
LeftPos = txttemp(Index).Left + Me.Left
If KeyAscii = 13 Then Exit Sub
Select Case Index
    Case 4 '-- Party
        '-- Search Condition
        X = "(Agcode=90017 or GpCode=90017)"
        gClsSearch.SearchMultiField "QryMastaccount", "Account,AcName,AcCity,AgName,AcGSTIN,AcStateCode", Array("Account", "AcName", "City", "Group", "GSTIN", "State Cd"), Array(3000, 0, 2000, 3000, 0, 0), X, Chr(KeyAscii), "acName", LeftPos, TopPos, True
        KeyAscii = 0
        If gClsSearch.SearchMultiRetCol(0) <> "" And gClsSearch.SearchMultiRetCol(1) = "" Then
            txttemp(Index) = gClsSearch.SearchMultiRetCol(0)
            lblGSTINNo.Caption = gClsSearch.SearchMultiRetCol(4)
            lblStateCd.Caption = gClsSearch.SearchMultiRetCol(5)
        Else
            If gClsSearch.SearchMultiRetCol(1) <> "" Then txttemp(Index) = gClsSearch.SearchMultiRetCol(1)
            lblGSTINNo.Caption = gClsSearch.SearchMultiRetCol(4)
            lblStateCd.Caption = gClsSearch.SearchMultiRetCol(5)
        End If
    Case 3 '-- Sales A/c  Alias Account
         '--- Consignment & Depot Sale
        If frmTypeStr = "SO" Or frmTypeStr = "SD" Then '-- Mill
            X = "(AgCode=90029 or GpCode=90029)"
        Else
'            X = "(AgCode=90023 or GpCode=90023 or AgCode = 90026 or AgCode = 90028)"
        End If
        gClsSearch.SearchMultiField "QryMastaccount", "Account,AcName,AcCity,AgName", Array("Sales Account", "AcName", "City", "Group"), Array(3000, 0, 2000, 3000), X, Chr(KeyAscii), "AcName", LeftPos, TopPos, True
        KeyAscii = 0
        If gClsSearch.SearchMultiRetCol(0) <> "" And gClsSearch.SearchMultiRetCol(1) = "" Then
            txttemp(Index) = gClsSearch.SearchMultiRetCol(0)
        Else
            If gClsSearch.SearchMultiRetCol(1) <> "" Then txttemp(Index) = gClsSearch.SearchMultiRetCol(1)
        End If
        '--- Invoice No
        If (Index = 3 And FormAction = vbDataActionAddNew) Or txt(36) = "" Then
            AcCode = GProcGetColumnValue("TblMastAccount", "AcName", txttemp(3), "S", "AcCode", "N")
            txt(36) = GProcGenerateOtherInvNo(frmTypeStr, AcCode) '--Inv No
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
            txttemp(Index) = gClsSearch.SearchMultiRetCol(0)
        Else
            If gClsSearch.SearchMultiRetCol(1) <> "" Then txttemp(Index) = gClsSearch.SearchMultiRetCol(1)
        End If
    Case 12 '-- Consinee
        gClsSearch.SearchMultiField "tblMastNarration", "Narration,GodAdd1,GodAdd2,GodCity", Array("Consinee Name", "Add 1", "Add 2", "City"), Array(3000, 2000, 2000, 1500), " NarrType='S'", Chr(KeyAscii), "Narration", 0, 0, True
        KeyAscii = 0
        If gClsSearch.SearchMultiRetCol(0) <> "" Then
            txttemp(Index) = gClsSearch.SearchMultiRetCol(0)
        End If
End Select
End Sub
Private Sub txtTemp_LostFocus(Index As Integer)
Dim AcCode As Long
SaveBtnEd
ShowAccountBalance
If (Index = 3 And FormAction = vbDataActionAddNew) Or txt(36) = "" Then
     AcCode = GProcGetColumnValue("TblMastAccount", "AcName", txttemp(3), "S", "AcCode", "N")
     If frmTypeStr = "ST" Then
        txt(36) = GProcGenerateOtherInvNo(frmTypeStr, SITSalesMillAccode) '--Inv No
     Else
        txt(36) = GProcGenerateOtherInvNo(frmTypeStr, AcCode) '--Inv No
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
           GProcShowForm frmMastAccount, frmMain.mnuMstfrm(0), Array(frmMastAccount.txttemp(2)), Array(GName)
    End Select
End If
End Sub
Private Sub txttempOS_GotFocus(Index As Integer)
GProcSelectBox Me.ActiveControl
End Sub
Private Sub SaveBtnEd()
'--- Vno,Inv No,vdt,Party,A/c,Mill,Amount
If CDbl(txt(0)) = 0 Or txt(36) = "" Or GProcIsDateValid(mskDt(2)) = False Or txttemp(3) = "" Or txttemp(4) = "" Or CDbl(txt(11)) = 0 Then
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
    '---Tax Type
    If txttemp(12) = "" Then txttemp(12) = "-"
    '---- Sale Account
    If txttemp(3) = "" Then
        MsgBox "Check Sale Account.", vbInformation + vbOKOnly, Me.Caption
        txttemp(3).SetFocus
        Exit Function
    End If
        '--Create Sale A/c
        If GProcGetColumnValue("tblMastAccount", "AcName", txttemp(3), "S", "AcCode", "N") = 0 Then
           If frmTypeStr = "SO" Or frmTypeStr = "SD" Then '-- Mill
                GName = GProcGetColumnValue("TblMastGroup", "AgCode", 90029, "N", "AgName", "S")
           Else
                GName = GProcGetColumnValue("TblMastGroup", "AgCode", 90023, "N", "AgName", "S")
           End If
           GProcShowForm frmMastAccount, frmMain.mnuMstfrm(0), Array(frmMastAccount.txt(1), frmMastAccount.txttemp(2)), Array(txttemp(3), GName)
           Exit Function
        End If
    '---- Party
    If txttemp(4) = "" Then
        MsgBox "Check Party.", vbInformation + vbOKOnly, Me.Caption
        txttemp(4).SetFocus
        Exit Function
    End If
        '--Create Party
        If GProcGetColumnValue("tblMastAccount", "AcName", txttemp(4), "S", "AcCode", "N") = 0 Then
           GName = GProcGetColumnValue("TblMastGroup", "AgCode", 90017, "N", "AgName", "S")
           GProcShowForm frmMastAccount, frmMain.mnuMstfrm(0), Array(frmMastAccount.txt(1), frmMastAccount.txttemp(2)), Array(txttemp(4), GName)
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
            MsgBox "Record is Available of Invoice No. " & txt(36), vbCritical + vbOKOnly, "Duplicate Record"
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
'-- Cr Code
txt(3) = GProcGetColumnValue("tblMastAccount", "AcName", txttemp(3), "S", "AcCode", "N")
'-- Dr Code
txt(4) = GProcGetColumnValue("tblMastAccount", "AcName", txttemp(4), "S", "AcCode", "N")
'--Mill
txt(5) = GProcGetColumnValue("tblMastAccount", "AcName", "-", "S", "AcCode", "N")
'--Consinee
If txttemp(12) = "" Then
   txttemp(12) = "-"
End If
txt(12) = GProcGetColumnValue("tblMastNarration", "Narration", txttemp(12), "S", "NarrCode", "N")
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
'-- Cr Code
txttemp(3) = GProcGetColumnValue("tblMastAccount", "AcCode", txt(3), "N", "AcName", "S")
'-- Dr Code
txttemp(4) = GProcGetColumnValue("tblMastAccount", "AcCode", txt(4), "N", "AcName", "S")
lblGSTINNo.Caption = GProcGetColumnValue("tblMastAccount", "AcCode", txt(4), "N", "AcGSTIN", "S")
lblStateCd.Caption = GProcGetColumnValue("tblMastAccount", "AcCode", txt(4), "N", "AcStateCode", "S")

'--Mill
txttemp(5) = "-"
'--Consinee
txttemp(12) = GProcGetColumnValue("tblMastNarration", "NarrCode", txt(12), "N", "Narration", "S")
'-- Fill Related Record
FillRelatedRecord
'--- Total+ Debit Amt
SetControlEd
ShowAccountBalance
End Sub
Private Sub FillRelatedRecord()
Dim Pu_ItCtrlNo As Long
'--- Fill Related Control
SetRelatedRS
'==== Vou Sub
txtVSub(5) = Format(rstblSub.Fields("VSubBag"), PStr) '--Bag
txtVSub(6) = Format(rstblSub.Fields("VSubWt"), WtStr) '--Wt
txtVSub(7) = Format(rstblSub.Fields("VSubRt"), LStr) '--Rate
txtVSub(8) = Format(rstblSub.Fields("VSubRtPer"), FStr) '--Rate Per
txtVSub(9) = Format(rstblSub.Fields("VSubAmt"), FStr) '--Amt
txtVSub(22) = Format(rstblSub.Fields("VSubNetRate"), FStr) '--Net Rate
txtVSub(10) = rstblSub.Fields("VSubLotNo") & vbNullString  '--Lot No
txtVSub(12) = rstblSub.Fields("VSubPItCtrlNo") '--PIt Ctrl No
txtVSub(47) = Format(rstblSub.Fields("VsubCGSTRt"), FStr) '--CGST Rt
txtVSub(48) = Format(rstblSub.Fields("VsubCGSTAmt"), FStr) '--CGST Amt
txtVSub(49) = Format(rstblSub.Fields("VsubSGSTRt"), FStr) '--SGST Rt
txtVSub(50) = Format(rstblSub.Fields("VsubSGSTAmt"), FStr) '--SGST Amt
txtVSub(51) = Format(rstblSub.Fields("VsubIGSTRt"), FStr) '--IGST Rt
txtVSub(52) = Format(rstblSub.Fields("VsubIGSTAmt"), FStr) '--IGST Amt
txtVSub(29) = rstblSub.Fields("VSUBTMP6") & vbNullString
txtVSub(30) = rstblSub.Fields("VSUBTMP7") & vbNullString


Pu_ItCtrlNo = GProcGetColumnValue("TblvouSub", "VSubItCtrlNo", txtVSub(12), "N", "VSubPItCtrlNo", "N", " vtype='GP'")
txttmpInNo = GProcGetColumnValue("TblvouSub", "VSubItCtrlNo", CStr(Pu_ItCtrlNo), "N", "VNo", "N")
'--Item
txttmpVSub(4) = GProcGetColumnValue("tblMastItem", "ItCode", rstblSub.Fields("VSubITCode"), "N", "ItName", "S")
lblHsnCode.Caption = GProcGetColumnValue("tblMastItem", "ItCode", rstblSub.Fields("VSubITCode"), "N", "ITTMP2", "S")

'--Godown
txttmpVSub(17) = GProcGetColumnValue("tblMastNarration", "NarrCode", rstblSub.Fields("VSubGodownCode"), "N", "Narration", "S")
'==== OutStanding
'--Broker
'If rsOS.Fields("OutBrokerCode") > 0 Then txttempOS(6) = GProcGetColumnValue("tblMastAccount", "AcCode", rsOS.Fields("OutBrokerCode"), "N", "AcName", "S")
'txtOS(9) = rsOS.Fields("OutDueDays") '--Due Days
'If IsDate(rsOS.Fields("OutDueDate")) = True Then '--Due Dt
'    mskOS(10) = rsOS.Fields("OutDueDate")
'Else
'    mskOS(10) = "__/__/____"
'End If
End Sub
Private Sub ReadFields()
FormAction = vbDataActionClose
GProcShowRecord Me, rstbl, MaxNo
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

If IsNumeric(txtVSub(47)) = False Then txtVSub(47) = "0.00"
If IsNumeric(txtVSub(48)) = False Then txtVSub(48) = "0.00"
If IsNumeric(txtVSub(49)) = False Then txtVSub(49) = "0.00"
If IsNumeric(txtVSub(50)) = False Then txtVSub(50) = "0.00"
If IsNumeric(txtVSub(51)) = False Then txtVSub(51) = "0.00"
If IsNumeric(txtVSub(52)) = False Then txtVSub(52) = "0.00"

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
If txtVSub(8) = 0 Then txtVSub(8) = 1
    
SubAmt = Val(txtVSub(9))
    
'-- CGST
         
         
         
         
CTaxAmt = SubAmt * CDbl(txtVSub(47)) / 100
CTaxAmt = GProcMakeRounding(CTaxAmt)
CTaxAmt = Format(CTaxAmt, FStr)
    
'-- SGST
         
STaxAmt = SubAmt * CDbl(txtVSub(49)) / 100
STaxAmt = GProcMakeRounding(STaxAmt)
STaxAmt = Format(STaxAmt, FStr)
    
'-- IGST
         
ITaxAmt = SubAmt * CDbl(txtVSub(51)) / 100
ITaxAmt = GProcMakeRounding(ITaxAmt)
ITaxAmt = Format(ITaxAmt, FStr)
    
Amt = SubAmt + (CTaxAmt + STaxAmt + ITaxAmt)
Amt = GProcMakeRounding(Amt)
    
    
'--ROff = Amt - SubAmt - TaxAmt
ROff = Amt - (SubAmt + CTaxAmt + STaxAmt + ITaxAmt)
ROff = Format(ROff, FStr)


txtVSub(7) = Rate '--Rate
txtVSub(9) = Format(SubAmt, FStr) '--SubAmt
txtVSub(48) = Format(CTaxAmt, FStr) '--CGST
txtVSub(50) = Format(STaxAmt, FStr) '--SGST
txtVSub(52) = Format(ITaxAmt, FStr) '--IGST
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
'       If frmTypeStr = "ST" Then
'          BillSrNo = GProcGetColumnValue("tblMastAccount", "AcName", txttemp(5), "S", "AcBillSrNo", "S")
'       Else
'          BillSrNo = GProcGetColumnValue("tblMastAccount", "AcName", txttemp(3), "S", "AcBillSrNo", "S")
'       End If
       BillSrNo = "OTH"
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
                       .Formulas(25) = "wCopy = 'Original Copy'"
                     ElseIf OptDup.Value = True Then
                        .Formulas(25) = "wCopy = 'Duplicate Copy'"
                     End If

                    .SelectionFormula = "{tblAddLess.AdInvNo} = '" & RsSA!AdInVNo & "' and {tblAddLess.VType}='" & RsSA!Vtype & "' and {tblAddLess.VYear}=" & RsSA!VYear
                    '--Report File At  D:\HiReports if exist
                    If Dir("D:\HiReports\rptSaleBillPrint.rpt") <> "" Then
                        .ReportFileName = "D:\HiReports\rptSaleBillPrint.rpt"
                    Else
                        .ReportFileName = gReportPath & "rptSaleBillPrintGSTOther.rpt"
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
Case 3, 4:
    If IsNumeric(txtPInvno(0)) = False Then txtPInvno(0) = 0
    If IsNumeric(txtPInvno(1)) = False Then txtPInvno(1) = txtPInvno(0)
    If CLng(txtPInvno(1)) < CLng(txtPInvno(0)) Then txtPInvno(1) = txtPInvno(0)
    
    InvNo = txtPInvno(0)
    BillSrNo = "OTH"
       
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

    Y = "Select a.Vno,a.Vtype,a.Vdt,a.Vyear,a.AdAcDrCode,a.AdTaxCode,a.AdInvNo,iif('" & BillSrNo & "' = '',val(a.AdInvNo),val(mid(a.AdInvNo,5,10))),a.AdSubAmt,a.AdBillAmt,b.VsubCGSTRt,b.VsubCGSTAmt,b.VsubSGSTRt,b.VsubSGSTAmt, " _
        & " b.VsubIGSTRt,b.VsubIGSTAmt,a.AdRoff,a.AdTcsAmt,b.VSubItCode,b.VSubBag,b.VSubWt,b.VSubRtPer,b.VSubAmt,'" & gUserName & "' " _
        & " from tblAddLess A ,tblVouSub B" _
        & " where a.vno = b.vno and a.vtype = b.vtype and a.vyear = b.vyear and b.VSubCtrNo = 1 and a.vtype='" & frmTypeStr & "' and a.AdInvNo in " & X & "   "
    gCn.BeginTrans
    gCn.Execute "Delete from tmpGenTbl2 where  USERNAME='" & gUserName & "'"

    gCn.Execute " insert into tmpGenTbl2 (Vno,Vtype,Vdt,Vyear,VACCODE,Amt15,VBILLNO,VsrNo,Amt1,Amt2,Amt3,Amt4,Amt5,Amt6,Amt7,Amt8,Amt9,Amt10," _
                & "VACOCODE,Amt11,Amt12,Amt13,Amt14,USERNAME  ) " & Y
    gCn.CommitTrans
    X = "Select * from tmpGenTbl2 where  USERNAME='" & gUserName & "' order by VsrNo"
    If GProcRstOpen(RsSA, X, "R") <= 0 Then
       MsgBox "No Records..."
       Exit Sub
    End If
     
    If Index = 4 Then
        GProcCrystalRptPreparation cryReport, "Window"
        With cryReport
             .Formulas(4) = "wCompGstin = '" & gCGSTIN & "'"
             .SelectionFormula = "{tmpGenTbl2.USERNAME} = '" & gUserName & "' "
             .ReportFileName = gReportPath & "rpte-Invoice.rpt"
             .Action = 1
        End With
        Exit Sub
    End If

    Dim LnCnt As Integer
    Dim FirmGstin As String
    Dim PartyName As String
    Dim PartyGstin As String
    Dim PartyAdd1 As String
    Dim PartyAdd2 As String
    Dim PartyCity As String
    Dim PartyPin As String
    Dim ItemHsn As String
    Dim ItemName As String
    Dim aFile As String
    
    Dim ConsiName As String
    Dim ConsiGstin As String
    Dim ConsiAdd1 As String
    Dim ConsiAdd2 As String
    Dim ConsiCity As String
    Dim ConsiPin As String
    Dim ii As Integer
    
    
    
    
    LnCnt = 1

    aFile = "D:\e-Inv-" + gCCode + "*.json"
    If Len(Dir$(aFile)) > 0 Then
          Kill aFile
    End If

    Open "D:\e-Inv-" + gCCode + ".json" For Output As #1

    Print #1, "["
    RsSA.MoveFirst
    Do While Not RsSA.EOF
       FirmGstin = gCGSTIN
       PartyName = GProcGetColumnValue("tblMastAccount", "AcCode", RsSA.Fields("VACCODE"), "N", "AcName", "S")
       
       PartyGstin = GProcGetColumnValue("tblMastAccount", "AcCode", RsSA.Fields("VACCODE"), "N", "AcGSTIN", "S")
       PartyAdd1 = GProcGetColumnValue("tblMastAccount", "AcCode", RsSA.Fields("VACCODE"), "N", "AcAdd1", "S")
       PartyAdd2 = GProcGetColumnValue("tblMastAccount", "AcCode", RsSA.Fields("VACCODE"), "N", "AcAdd2", "S")
       PartyCity = GProcGetColumnValue("tblMastAccount", "AcCode", RsSA.Fields("VACCODE"), "N", "AcCity", "S")
       PartyPin = GProcGetColumnValue("tblMastAccount", "AcCode", RsSA.Fields("VACCODE"), "N", "AcZip", "S")
       ItemHsn = GProcGetColumnValue("tblMastItem", "ItCode", RsSA.Fields("VACOCODE"), "N", "ITTMP2", "S")
       ItemName = GProcGetColumnValue("tblMastItem", "ItCode", RsSA.Fields("VACOCODE"), "N", "ItName", "S")
       
       If Len(PartyPin) > 6 Then
          PartyPin = Left(PartyPin, 3) + Right(PartyPin, 3)
       End If
       If Trim(PartyAdd1) = "" Then PartyAdd1 = "null"
       If Trim(PartyAdd2) = "" Then PartyAdd2 = "null"
       
       '---Checks
       ii = InStr(1, PartyName, Chr(34))
       If ii > 0 Then
          MsgBox ("Name Contens " + """ ""  ..For '" & PartyName & "'   ")
          Close #1
          Exit Sub
       End If
       ii = InStr(1, PartyAdd1, Chr(34))
       If ii > 0 Then
          MsgBox ("Address Line 1 Contens " + """ ""  ..For '" & PartyName & "'   ")
          Close #1
          Exit Sub
       End If
       ii = InStr(1, PartyAdd2, Chr(34))
       If ii > 0 Then
          MsgBox ("Address Line 2 Contens " + """ ""  ..For '" & PartyName & "'   ")
          Close #1
          Exit Sub
       End If
       If Trim(PartyCity) = "" Then
          MsgBox ("City Not Found..For +'" & PartyName & "'   ")
          Close #1
          Exit Sub
       End If
       If Trim(PartyPin) = "" Then
          MsgBox ("Pin Not Found..For +'" & PartyName & "'   ")
          Close #1
          Exit Sub
       End If
       
       ConsiName = GProcGetColumnValue("tblMastNarration", "NarrCode", RsSA.Fields("Amt15"), "N", "Narration", "S")
       If ConsiName <> "-" Then
          ConsiGstin = GProcGetColumnValue("tblMastNarration", "NarrCode", RsSA.Fields("Amt15"), "N", "GSTIN", "S")
          ConsiGstin = UCase(ConsiGstin)
          ConsiAdd1 = GProcGetColumnValue("tblMastNarration", "NarrCode", RsSA.Fields("Amt15"), "N", "GodAdd1", "S")
          ConsiAdd2 = GProcGetColumnValue("tblMastNarration", "NarrCode", RsSA.Fields("Amt15"), "N", "GodAdd2", "S")
          ConsiCity = GProcGetColumnValue("tblMastNarration", "NarrCode", RsSA.Fields("Amt15"), "N", "GodCity", "S")
          ConsiPin = GProcGetColumnValue("tblMastNarration", "NarrCode", RsSA.Fields("Amt15"), "N", "MastTaxRate", "N")
          ConsiPin = Left(ConsiPin, 6)
          If ConsiAdd1 = "" Then ConsiAdd1 = "null"
          If ConsiAdd2 = "" Then ConsiAdd2 = "null"
          If ConsiPin = 0 Then
             MsgBox ("Consignee Pin Code Not Found..For +'" & ConsiName & "'   ")
             Close #1
             Exit Sub
          End If
          If ConsiGstin = "" Then
             MsgBox ("Consignee GSTIN Not Found..For +'" & ConsiName & "'   ")
             Close #1
             Exit Sub
          End If
          If ConsiCity = "" Then
             MsgBox ("Consignee City Not Found..For +'" & ConsiName & "'   ")
             Close #1
             Exit Sub
          End If
       
       End If


       Print #1, "{"
       Print #1, Chr(34) + "Version" + Chr(34) + ":" + Chr(34) + "1.1" + Chr(34) + Chr(44)
       Print #1, Chr(34) + "TranDtls" + Chr(34) + ":{"
       Print #1, Space(15) + Chr(34) + "TaxSch" + Chr(34) + ":" + Chr(34) + "GST" + Chr(34) + Chr(44)
       Print #1, Space(15) + Chr(34) + "SupTyp" + Chr(34) + ":" + Chr(34) + "B2B" + Chr(34) + Chr(44)
       Print #1, Space(15) + Chr(34) + "IgstOnIntra" + Chr(34) + ":" + Chr(34) + "N" + Chr(34) + Chr(44)
       Print #1, Space(15) + Chr(34) + "RegRev" + Chr(34) + ":" + "null" + Chr(44)
       Print #1, Space(15) + Chr(34) + "EcmGstin" + Chr(34) + ":" + "null"
       Print #1, "}" + Chr(44)

       Print #1, Chr(34) + "DocDtls" + Chr(34) + ":{"
       Print #1, Space(15) + Chr(34) + "Typ" + Chr(34) + ":" + Chr(34) + "INV" + Chr(34) + Chr(44)
       Print #1, Space(15) + Chr(34) + "No" + Chr(34) + ":" + Chr(34) + RsSA.Fields("VBILLNO") + Chr(34) + Chr(44)
       Print #1, Space(15) + Chr(34) + "Dt" + Chr(34) + ":" + Chr(34) + CStr(RsSA.Fields("Vdt")) + Chr(34)
       Print #1, "}" + Chr(44)

       Print #1, Chr(34) + "SellerDtls" + Chr(34) + ":{"
       Print #1, Space(15) + Chr(34) + "Gstin" + Chr(34) + ":" + Chr(34) + "" & FirmGstin & "" + Chr(34) + Chr(44)
       Print #1, Space(15) + Chr(34) + "LglNm" + Chr(34) + ":" + Chr(34) + "" & UCase(gCName) & "" + Chr(34) + Chr(44)
       Print #1, Space(15) + Chr(34) + "TrdNm" + Chr(34) + ":" + "null" + Chr(44)
       Print #1, Space(15) + Chr(34) + "Addr1" + Chr(34) + ":" + Chr(34) + "" & gCAdd1 & "" + Chr(34) + Chr(44)
       Print #1, Space(15) + Chr(34) + "Addr2" + Chr(34) + ":" + Chr(34) + "" & gCAdd2 & "" + Chr(34) + Chr(44)
       Print #1, Space(15) + Chr(34) + "Addr3" + Chr(34) + ":" + Chr(34) + "" & gCAdd3 & "" + Chr(34) + Chr(44)
       Print #1, Space(15) + Chr(34) + "Loc" + Chr(34) + ":" + Chr(34) + "Mumbai" + Chr(34) + Chr(44)
       Print #1, Space(15) + Chr(34) + "Pin" + Chr(34) + ":" + "400002" + Chr(44)
       Print #1, Space(15) + Chr(34) + "Stcd" + Chr(34) + ":" + Chr(34) + "" & Left(FirmGstin, 2) & "" + Chr(34) + Chr(44)
       Print #1, Space(15) + Chr(34) + "Ph" + Chr(34) + ":" + "null" + Chr(44)
       Print #1, Space(15) + Chr(34) + "Em" + Chr(34) + ":" + "null"
       Print #1, "}" + Chr(44)

       Print #1, Chr(34) + "BuyerDtls" + Chr(34) + ":{"
       Print #1, Space(15) + Chr(34) + "Gstin" + Chr(34) + ":" + Chr(34) + "" & PartyGstin & "" + Chr(34) + Chr(44)
       Print #1, Space(15) + Chr(34) + "LglNm" + Chr(34) + ":" + Chr(34) + "" & PartyName & "" + Chr(34) + Chr(44)
       Print #1, Space(15) + Chr(34) + "TrdNm" + Chr(34) + ":" + "null" + Chr(44)
       Print #1, Space(15) + Chr(34) + "Pos" + Chr(34) + ":" + Chr(34) + "" & Left(PartyGstin, 2) & "" + Chr(34) + Chr(44)
       Print #1, Space(15) + Chr(34) + "Addr1" + Chr(34) + ":" + Chr(34) + "" & PartyAdd1 & "" + Chr(34) + Chr(44)
       Print #1, Space(15) + Chr(34) + "Addr2" + Chr(34) + ":" + Chr(34) + "" & PartyAdd2 & "" + Chr(34) + Chr(44)
       Print #1, Space(15) + Chr(34) + "Loc" + Chr(34) + ":" + Chr(34) + "" & PartyCity & "" + Chr(34) + Chr(44)
       Print #1, Space(15) + Chr(34) + "Pin" + Chr(34) + ":" + "" & PartyPin & "" + Chr(44)
       Print #1, Space(15) + Chr(34) + "Stcd" + Chr(34) + ":" + Chr(34) + "" & Left(PartyGstin, 2) & "" + Chr(34) + Chr(44)
       Print #1, Space(15) + Chr(34) + "Ph" + Chr(34) + ":" + "null" + Chr(44)
       Print #1, Space(15) + Chr(34) + "Em" + Chr(34) + ":" + "null"
       Print #1, "}" + Chr(44)

       Print #1, Chr(34) + "DispDtls" + Chr(34) + ":" + "null" + Chr(44)

       If ConsiName <> "-" Then
          Print #1, Chr(34) + "ShipDtls" + Chr(34) + ":{"
          Print #1, Space(15) + Chr(34) + "Gstin" + Chr(34) + ":" + Chr(34) + "" & ConsiGstin & "" + Chr(34) + Chr(44)
          Print #1, Space(15) + Chr(34) + "LglNm" + Chr(34) + ":" + Chr(34) + "" & ConsiName & "" + Chr(34) + Chr(44)
          Print #1, Space(15) + Chr(34) + "TrdNm" + Chr(34) + ":" + "null" + Chr(44)
          Print #1, Space(15) + Chr(34) + "Addr1" + Chr(34) + ":" + Chr(34) + "" & ConsiAdd1 & "" + Chr(34) + Chr(44)
          Print #1, Space(15) + Chr(34) + "Addr2" + Chr(34) + ":" + Chr(34) + "" & ConsiAdd2 & "" + Chr(34) + Chr(44)
          Print #1, Space(15) + Chr(34) + "Loc" + Chr(34) + ":" + Chr(34) + "" & ConsiCity & "" + Chr(34) + Chr(44)
          Print #1, Space(15) + Chr(34) + "Pin" + Chr(34) + ":" + "" & ConsiPin & "" + Chr(44)
          Print #1, Space(15) + Chr(34) + "Stcd" + Chr(34) + ":" + Chr(34) + "" & Left(ConsiGstin, 2) & "" + Chr(34)
          Print #1, "}" + Chr(44)
       
       Else
          Print #1, Chr(34) + "ShipDtls" + Chr(34) + ":" + "null" + Chr(44)
       End If

       Print #1, Chr(34) + "ValDtls" + Chr(34) + ":{"
       Print #1, Space(15) + Chr(34) + "AssVal" + Chr(34) + ":" + Trim(CStr(Format(RsSA.Fields("Amt1"), "############0.00"))) + Chr(44)
       Print #1, Space(15) + Chr(34) + "IgstVal" + Chr(34) + ":" + Trim(CStr(Format(RsSA.Fields("Amt8"), "############0.00"))) + Chr(44)
       Print #1, Space(15) + Chr(34) + "CgstVal" + Chr(34) + ":" + Trim(CStr(Format(RsSA.Fields("Amt4"), "############0.00"))) + Chr(44)
       Print #1, Space(15) + Chr(34) + "SgstVal" + Chr(34) + ":" + Trim(CStr(Format(RsSA.Fields("Amt6"), "############0.00"))) + Chr(44)
       Print #1, Space(15) + Chr(34) + "CesVal" + Chr(34) + ":" + Trim(CStr(Format(0, "############0.00"))) + Chr(44)
       Print #1, Space(15) + Chr(34) + "StCesVal" + Chr(34) + ":" + Trim(CStr(Format(0, "############0.00"))) + Chr(44)
       Print #1, Space(15) + Chr(34) + "Discount" + Chr(34) + ":" + Trim(CStr(Format(0, "############0.00"))) + Chr(44)
       Print #1, Space(15) + Chr(34) + "OthChrg" + Chr(34) + ":" + Trim(CStr(Format(RsSA.Fields("Amt10"), "############0.00"))) + Chr(44)
       Print #1, Space(15) + Chr(34) + "RndOffAmt" + Chr(34) + ":" + Trim(CStr(Format(RsSA.Fields("Amt9"), "############0.00"))) + Chr(44)
       Print #1, Space(15) + Chr(34) + "TotInvVal" + Chr(34) + ":" + Trim(CStr(Format(RsSA.Fields("Amt2"), "############0.00")))
       Print #1, "}" + Chr(44)
       
       Print #1, Space(25) + Chr(34) + "ItemList" + Chr(34) + ":[{"
       Print #1, Space(32) + Chr(34) + "SlNo" + Chr(34) + ":" + Chr(34) + "1" + Chr(34) + Chr(44)
       Print #1, Space(32) + Chr(34) + "PrdDesc" + Chr(34) + ":" + Chr(34) + Trim(ItemName) + Chr(34) + Chr(44)
       Print #1, Space(32) + Chr(34) + "IsServc" + Chr(34) + ":" + Chr(34) + "Y" + Chr(34) + Chr(44)
       Print #1, Space(32) + Chr(34) + "HsnCd" + Chr(34) + ":" + Chr(34) + Trim(ItemHsn) + Chr(34) + Chr(44)
       Print #1, Space(32) + Chr(34) + "Qty" + Chr(34) + ":" + Trim(CStr(Format(1, "############0.000"))) + Chr(44)
       Print #1, Space(32) + Chr(34) + "Unit" + Chr(34) + ":" + Chr(34) + "OTH" + Chr(34) + Chr(44)
       Print #1, Space(32) + Chr(34) + "UnitPrice" + Chr(34) + ":" + Trim(CStr(Format(RsSA.Fields("Amt14"), "############0.000"))) + Chr(44)
       Print #1, Space(32) + Chr(34) + "TotAmt" + Chr(34) + ":" + Trim(CStr(Format(RsSA.Fields("Amt14"), "############0.00"))) + Chr(44)
       Print #1, Space(32) + Chr(34) + "Discount" + Chr(34) + ":" + Trim(CStr(Format(0, "############0.00"))) + Chr(44)
       Print #1, Space(32) + Chr(34) + "PreTaxVal" + Chr(34) + ":" + Trim(CStr(Format(0, "############0.00"))) + Chr(44)
       Print #1, Space(32) + Chr(34) + "AssAmt" + Chr(34) + ":" + Trim(CStr(Format(RsSA.Fields("Amt1"), "############0.00"))) + Chr(44)
       Print #1, Space(32) + Chr(34) + "GstRt" + Chr(34) + ":" + Trim(CStr(Format(RsSA.Fields("Amt3") + RsSA.Fields("Amt5") + RsSA.Fields("Amt7"), "############0"))) + Chr(44)
       Print #1, Space(32) + Chr(34) + "IgstAmt" + Chr(34) + ":" + Trim(CStr(Format(RsSA.Fields("Amt8"), "############0.00"))) + Chr(44)
       Print #1, Space(32) + Chr(34) + "CgstAmt" + Chr(34) + ":" + Trim(CStr(Format(RsSA.Fields("Amt4"), "############0.00"))) + Chr(44)
       Print #1, Space(32) + Chr(34) + "SgstAmt" + Chr(34) + ":" + Trim(CStr(Format(RsSA.Fields("Amt6"), "############0.00"))) + Chr(44)
       Print #1, Space(32) + Chr(34) + "CesRt" + Chr(34) + ":" + Trim(CStr(Format(0, "############0.00"))) + Chr(44)
       Print #1, Space(32) + Chr(34) + "CesAmt" + Chr(34) + ":" + Trim(CStr(Format(0, "############0.00"))) + Chr(44)
       Print #1, Space(32) + Chr(34) + "CesNonAdvlAmt" + Chr(34) + ":" + Trim(CStr(Format(0, "############0.00"))) + Chr(44)
       Print #1, Space(32) + Chr(34) + "StateCesRt" + Chr(34) + ":" + Trim(CStr(Format(0, "############0.00"))) + Chr(44)
       Print #1, Space(32) + Chr(34) + "StateCesAmt" + Chr(34) + ":" + Trim(CStr(Format(0, "############0.00"))) + Chr(44)
       Print #1, Space(32) + Chr(34) + "StateCesNonAdvlAmt" + Chr(34) + ":" + Trim(CStr(Format(0, "############0.00"))) + Chr(44)
       Print #1, Space(32) + Chr(34) + "OthChrg" + Chr(34) + ":" + Trim(CStr(Format(0, "############0.00"))) + Chr(44)
       Print #1, Space(32) + Chr(34) + "TotItemVal" + Chr(34) + ":" + Trim(CStr(Format(RsSA.Fields("Amt1") + RsSA.Fields("Amt8") + RsSA.Fields("Amt4") + RsSA.Fields("Amt6"), "############0.00")))
       Print #1, Space(30) + "}"
       Print #1, Space(30) + "]"
       RsSA.MoveNext
       If RsSA.EOF Then
          Print #1, Space(15) + "}"
       Else
          Print #1, "}" + Chr(44)
       End If
    Loop
    Print #1, "]"
    Close #1
    MsgBox ("e-Invoice Jason File Created....")
    Dim RetVal
'--Hitrix
'   RetVal = Shell("C:\Program Files (x86)\Google\Chrome\Application\chrome.exe https://einvoice1.gst.gov.in", vbMaximizedFocus)
'--"C:\Program Files (x86)\Google\Chrome\Application"
'    RetVal = Shell("""" & CStr(App.Path) & "" & "\chrome.exe https://einvoice1.gst.gov.in" & """", vbMaximizedFocus)
Case 5:
    Dim excel_app As Excel.Application
    Dim workbook As Excel.workbook
    Dim sheet_name As String
    Dim sheet As Excel.Worksheet
    Dim header_range As Excel.Range

    Dim flnm As String
    Dim a As Integer
    Dim PtyCd As Long
    Dim i As Integer
    
    flnm = "d:\UploadedInvoiceDetails"
    Set excel_app = New Excel.Application
    excel_app.Visible = False
    Set workbook = excel_app.Workbooks.Open(filename:=flnm)
    Set sheet = workbook.Sheets.Item("UploadedInvoiceDetails")
    If (sheet Is Nothing) Then
        MsgBox "Error Sales Sheet"
    End If
    a = 2
    Do While sheet.Cells(a, 1) <> ""
       X = "select * from tblMastAccount where AcGSTIN = '" & sheet.Cells(a, 9) & "'"
       i = GProcRstOpen(RsSA, X, "R")
       PtyCd = GProcGetColumnValue("tblMastAccount", "AcGSTIN", sheet.Cells(a, 9), "S", "AcCode", "N")
       If i > 1 Then
          gCn.Execute "Update tblAddLess set AdCFormNo  = '" & sheet.Cells(a, 2) & "',AdE1FormNo = '" & sheet.Cells(a, 3) & "'  where vtype='" & frmTypeStr & "' and AdInvNo= '" & sheet.Cells(a, 5) & "' and " _
          & " VYear=" & gCYear
       Else
          gCn.Execute "Update tblAddLess set AdCFormNo = '" & sheet.Cells(a, 2) & "',AdE1FormNo = '" & sheet.Cells(a, 3) & "'  where AdAcDrCode= " & PtyCd & " and vtype='" & frmTypeStr & "' and AdInvNo= '" & sheet.Cells(a, 5) & "' and " _
          & " VYear=" & gCYear
       End If
       a = a + 1
    Loop
    workbook.Close SaveChanges:=True
    excel_app.Quit
    rstbl.Requery
    MsgBox ("e-Invoice IRN No Updated....")
Case 2:
    DoEvents
    Frame1.Visible = False
End Select
End Sub
