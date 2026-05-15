VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form frmSalesMillBill 
   BackColor       =   &H00C2E3FC&
   BorderStyle     =   0  'None
   Caption         =   "Sales Entry"
   ClientHeight    =   7245
   ClientLeft      =   0
   ClientTop       =   15
   ClientWidth     =   10830
   DrawMode        =   10  'Mask Pen
   FillColor       =   &H000040C0&
   FillStyle       =   0  'Solid
   Icon            =   "frmSalesMillBill.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7245
   ScaleWidth      =   10830
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
      Index           =   55
      Left            =   0
      TabIndex        =   149
      Text            =   "55"
      Top             =   0
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
      Left            =   0
      TabIndex        =   148
      Text            =   "54"
      Top             =   4680
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
      Left            =   0
      TabIndex        =   147
      Text            =   "53"
      Top             =   0
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
      Left            =   8520
      TabIndex        =   146
      Text            =   "52"
      Top             =   6720
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
      Left            =   7920
      TabIndex        =   145
      Text            =   "51"
      Top             =   6720
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
      Index           =   50
      Left            =   7440
      TabIndex        =   144
      Text            =   "50"
      Top             =   6720
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
      Index           =   49
      Left            =   6960
      TabIndex        =   143
      Text            =   "49"
      Top             =   6720
      Visible         =   0   'False
      Width           =   420
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
      Height          =   2295
      Left            =   2760
      TabIndex        =   120
      Top             =   3240
      Visible         =   0   'False
      Width           =   5775
      Begin VB.CheckBox Check1 
         BackColor       =   &H00FFC0C0&
         Caption         =   "Payment Letter"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   4560
         TabIndex        =   142
         Top             =   1560
         Width           =   1095
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
         Left            =   3120
         TabIndex        =   125
         Top             =   1560
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
         Left            =   1560
         TabIndex        =   124
         Top             =   1560
         Width           =   1335
      End
      Begin VB.TextBox txtPInvno 
         Appearance      =   0  'Flat
         Height          =   360
         Index           =   1
         Left            =   2640
         MaxLength       =   7
         TabIndex        =   123
         Top             =   1080
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
         Left            =   120
         TabIndex        =   122
         Top             =   1560
         Width           =   1335
      End
      Begin VB.TextBox txtPInvno 
         Appearance      =   0  'Flat
         Height          =   360
         Index           =   0
         Left            =   2640
         MaxLength       =   7
         TabIndex        =   121
         Top             =   600
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
         TabIndex        =   127
         Top             =   1080
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
         TabIndex        =   126
         Top             =   600
         Width           =   1680
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00FFC0C0&
      Caption         =   "Debit Note Printing"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2295
      Left            =   2760
      TabIndex        =   111
      Top             =   4320
      Visible         =   0   'False
      Width           =   5175
      Begin VB.CommandButton CmdPrintDbNt 
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
         Left            =   3480
         TabIndex        =   118
         Top             =   1560
         Width           =   1335
      End
      Begin VB.CommandButton CmdPrintDbNt 
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
         Left            =   1920
         TabIndex        =   117
         Top             =   1560
         Width           =   1335
      End
      Begin VB.TextBox txtDbNtToNo 
         Appearance      =   0  'Flat
         Height          =   360
         Left            =   2640
         MaxLength       =   7
         TabIndex        =   115
         Top             =   1080
         Width           =   1695
      End
      Begin VB.CommandButton CmdPrintDbNt 
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
         Left            =   480
         TabIndex        =   116
         Top             =   1560
         Width           =   1335
      End
      Begin VB.TextBox txtDbNtFromNo 
         Appearance      =   0  'Flat
         Height          =   360
         Left            =   2640
         MaxLength       =   7
         TabIndex        =   113
         Top             =   600
         Width           =   1695
      End
      Begin VB.Label Label23 
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
         Top             =   1080
         Width           =   1665
      End
      Begin VB.Label Label33 
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
         TabIndex        =   112
         Top             =   600
         Width           =   1680
      End
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
      Left            =   3675
      TabIndex        =   49
      Text            =   "txttemp(48)"
      ToolTipText     =   "Press F3 for Creation"
      Top             =   6165
      Width           =   2385
   End
   Begin VB.TextBox txt 
      Alignment       =   1  'Right Justify
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
      Left            =   -120
      TabIndex        =   140
      Text            =   "37"
      Top             =   4230
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
      Left            =   -135
      TabIndex        =   139
      Text            =   "38"
      Top             =   3705
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
      Left            =   -105
      TabIndex        =   138
      Text            =   "39"
      Top             =   3210
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
      Left            =   -135
      TabIndex        =   137
      Text            =   "40"
      Top             =   2730
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
      Left            =   -120
      TabIndex        =   136
      Text            =   "41"
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
      Left            =   -135
      TabIndex        =   135
      Text            =   "42"
      Top             =   1950
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
      Left            =   -120
      TabIndex        =   134
      Text            =   "43"
      Top             =   1635
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
      Left            =   -120
      TabIndex        =   133
      Text            =   "44"
      Top             =   1230
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
      Left            =   -75
      TabIndex        =   132
      Text            =   "45"
      Top             =   0
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
      Index           =   46
      Left            =   8160
      TabIndex        =   131
      Text            =   "46"
      Top             =   2325
      Width           =   1140
   End
   Begin VB.TextBox txt 
      Alignment       =   1  'Right Justify
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
      Left            =   -90
      TabIndex        =   130
      Text            =   "47"
      Top             =   735
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
      Index           =   48
      Left            =   -75
      TabIndex        =   129
      Text            =   "48"
      Top             =   360
      Visible         =   0   'False
      Width           =   420
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
      Index           =   12
      Left            =   0
      TabIndex        =   128
      Text            =   "txtVSub(12)"
      Top             =   0
      Visible         =   0   'False
      Width           =   1050
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
      Left            =   9690
      Style           =   1  'Graphical
      TabIndex        =   119
      Top             =   4395
      Width           =   855
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
      Left            =   6630
      TabIndex        =   3
      Text            =   "36"
      Top             =   1395
      Width           =   1185
   End
   Begin VB.TextBox txt 
      Alignment       =   1  'Right Justify
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
      TabIndex        =   110
      Text            =   "35"
      Top             =   0
      Visible         =   0   'False
      Width           =   420
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
      ItemData        =   "frmSalesMillBill.frx":058A
      Left            =   5655
      List            =   "frmSalesMillBill.frx":0594
      TabIndex        =   31
      Text            =   "Cbo(12)"
      Top             =   4350
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
      Index           =   34
      Left            =   5310
      TabIndex        =   109
      Text            =   "34"
      Top             =   6615
      Visible         =   0   'False
      Width           =   510
   End
   Begin VB.CommandButton cmdLedger 
      Caption         =   "Ledger"
      Height          =   375
      Index           =   1
      Left            =   4035
      TabIndex        =   108
      Top             =   1875
      Width           =   1530
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
      Left            =   1530
      TabIndex        =   6
      Text            =   "txttemp(4)"
      ToolTipText     =   "Press F3 for Creation"
      Top             =   1395
      Width           =   4065
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
      Left            =   1530
      TabIndex        =   17
      Text            =   "txttempOS(6)"
      ToolTipText     =   "Press F3 for Creation"
      Top             =   2340
      Width           =   4065
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
      Left            =   1530
      TabIndex        =   106
      Text            =   "txtDrBalance"
      Top             =   1875
      Width           =   2415
   End
   Begin VB.TextBox txt 
      Alignment       =   1  'Right Justify
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
      Left            =   6630
      TabIndex        =   8
      Text            =   "8"
      Top             =   1875
      Width           =   1185
   End
   Begin VB.TextBox txt 
      Alignment       =   1  'Right Justify
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
      Left            =   6870
      TabIndex        =   14
      Text            =   "32"
      Top             =   210
      Visible         =   0   'False
      Width           =   1185
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
      Left            =   7605
      TabIndex        =   53
      Text            =   "txttmpAmt"
      Top             =   6165
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
      Index           =   22
      Left            =   2475
      TabIndex        =   47
      Text            =   "22"
      Top             =   6165
      Width           =   1155
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
      Left            =   7755
      TabIndex        =   29
      Text            =   "txtVSub(9)"
      Top             =   3525
      Width           =   1500
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
      Left            =   6630
      TabIndex        =   27
      Text            =   "txtVSub(8)"
      Top             =   3525
      Width           =   1035
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
      Left            =   5295
      TabIndex        =   25
      Text            =   "txtVSub(7)"
      Top             =   3525
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
      Index           =   6
      Left            =   4155
      TabIndex        =   23
      Text            =   "txtVSub(6)"
      Top             =   3525
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
      Left            =   420
      Locked          =   -1  'True
      TabIndex        =   19
      Text            =   "txttmpVSub(4)"
      ToolTipText     =   "Press F3 for Creation"
      Top             =   3525
      Width           =   2685
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
      Left            =   3195
      TabIndex        =   21
      Text            =   "txtVSub(5)"
      Top             =   3525
      Width           =   915
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
      Left            =   1830
      TabIndex        =   105
      Text            =   "33"
      Top             =   6720
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
      Index           =   31
      Left            =   2070
      TabIndex        =   104
      Text            =   "31"
      Top             =   6690
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
      Left            =   6420
      TabIndex        =   76
      Text            =   "30"
      Top             =   6705
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
      Left            =   2535
      TabIndex        =   75
      Text            =   "29"
      Top             =   6690
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
      Left            =   180
      TabIndex        =   74
      Text            =   "28"
      Top             =   6705
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
      Left            =   1305
      TabIndex        =   73
      Text            =   "27"
      Top             =   6705
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
      Left            =   5550
      TabIndex        =   72
      Text            =   "26"
      Top             =   6675
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
      Left            =   4005
      TabIndex        =   71
      Text            =   "25"
      Top             =   6645
      Visible         =   0   'False
      Width           =   510
   End
   Begin VB.CommandButton cmdLedger 
      Caption         =   "Ledger"
      Height          =   375
      Index           =   0
      Left            =   7680
      TabIndex        =   103
      Top             =   735
      Visible         =   0   'False
      Width           =   1545
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
      Left            =   5745
      TabIndex        =   102
      Text            =   "txtCrBalance"
      Top             =   735
      Width           =   1875
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
      Left            =   1530
      TabIndex        =   1
      Text            =   "txttemp(3)"
      ToolTipText     =   "Press F3 for Creation"
      Top             =   735
      Width           =   4065
   End
   Begin VB.TextBox txt 
      Alignment       =   1  'Right Justify
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
      Left            =   900
      TabIndex        =   77
      Text            =   "24"
      Top             =   6660
      Visible         =   0   'False
      Width           =   510
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
      TabIndex        =   100
      Text            =   "1"
      Top             =   2430
      Width           =   420
   End
   Begin VB.TextBox txt 
      Alignment       =   1  'Right Justify
      Height          =   315
      Index           =   2
      Left            =   11040
      TabIndex        =   89
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
      TabIndex        =   88
      Text            =   "5"
      Top             =   2640
      Width           =   525
   End
   Begin VB.TextBox txt 
      Alignment       =   1  'Right Justify
      Height          =   315
      Index           =   3
      Left            =   11040
      TabIndex        =   87
      Text            =   "3"
      Top             =   1800
      Width           =   525
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
      Left            =   8625
      Style           =   1  'Graphical
      TabIndex        =   81
      Top             =   6885
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
      TabIndex        =   66
      ToolTipText     =   "Exit"
      Top             =   6270
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
      TabIndex        =   58
      Top             =   5670
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
      TabIndex        =   57
      ToolTipText     =   "To Save Record"
      Top             =   5175
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
      Left            =   9690
      Style           =   1  'Graphical
      TabIndex        =   60
      Top             =   3960
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
      TabIndex        =   56
      ToolTipText     =   "To Delete Record"
      Top             =   3360
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
      TabIndex        =   55
      ToolTipText     =   "To Modify Record"
      Top             =   2880
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
      TabIndex        =   54
      ToolTipText     =   "To Add New Record"
      Top             =   2400
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
      Left            =   7995
      Style           =   1  'Graphical
      TabIndex        =   59
      Top             =   6825
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
      TabIndex        =   61
      ToolTipText     =   "To Find Record"
      Top             =   1815
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
      TabIndex        =   65
      ToolTipText     =   "Last"
      Top             =   1335
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
      TabIndex        =   64
      ToolTipText     =   "First"
      Top             =   1335
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
      TabIndex        =   63
      ToolTipText     =   "Next"
      Top             =   855
      Width           =   375
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
      Left            =   2160
      TabIndex        =   67
      Text            =   "0"
      Top             =   8085
      Visible         =   0   'False
      Width           =   255
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
      TabIndex        =   62
      ToolTipText     =   "Previous"
      Top             =   855
      Width           =   375
   End
   Begin VB.TextBox txt 
      Alignment       =   1  'Right Justify
      Height          =   315
      Index           =   4
      Left            =   11040
      TabIndex        =   86
      Text            =   "4"
      Top             =   2160
      Width           =   525
   End
   Begin VB.TextBox txt 
      Alignment       =   1  'Right Justify
      Height          =   375
      Index           =   6
      Left            =   8280
      TabIndex        =   11
      Text            =   "6"
      Top             =   600
      Width           =   1095
   End
   Begin VB.TextBox txt 
      Alignment       =   1  'Right Justify
      Height          =   315
      Index           =   7
      Left            =   11040
      TabIndex        =   85
      Text            =   "7"
      Top             =   3360
      Visible         =   0   'False
      Width           =   525
   End
   Begin VB.TextBox txt 
      Alignment       =   1  'Right Justify
      Height          =   315
      Index           =   9
      Left            =   11040
      TabIndex        =   84
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
      Left            =   450
      TabIndex        =   78
      Text            =   "10"
      Top             =   6690
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
      Height          =   360
      Index           =   11
      Left            =   7605
      TabIndex        =   37
      Text            =   "11"
      Top             =   5205
      Width           =   1650
   End
   Begin VB.TextBox txt 
      Alignment       =   1  'Right Justify
      Height          =   315
      Index           =   12
      Left            =   11040
      TabIndex        =   83
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
      Left            =   6945
      TabIndex        =   32
      Text            =   "13"
      Top             =   4335
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
      Left            =   7590
      TabIndex        =   33
      Text            =   "14"
      Top             =   4335
      Width           =   1665
   End
   Begin VB.TextBox txt 
      Alignment       =   1  'Right Justify
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
      Left            =   3450
      TabIndex        =   69
      Text            =   "15"
      Top             =   6690
      Visible         =   0   'False
      Width           =   510
   End
   Begin VB.TextBox txt 
      Alignment       =   1  'Right Justify
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
      Left            =   2985
      TabIndex        =   70
      Text            =   "16"
      Top             =   6690
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
      Height          =   360
      Index           =   17
      Left            =   7605
      TabIndex        =   35
      Text            =   "17"
      Top             =   4755
      Width           =   1650
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
      Left            =   1875
      TabIndex        =   39
      Text            =   "txtOS(9)"
      Top             =   4335
      Width           =   1245
   End
   Begin VB.TextBox txt 
      Alignment       =   1  'Right Justify
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
      Left            =   4500
      TabIndex        =   79
      Text            =   "18"
      Top             =   6675
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
      Left            =   4965
      TabIndex        =   80
      Text            =   "19"
      Top             =   6690
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
      Left            =   315
      TabIndex        =   43
      Text            =   "20"
      Top             =   6165
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
      Height          =   360
      Index           =   21
      Left            =   6135
      TabIndex        =   51
      Text            =   "21"
      Top             =   6165
      Width           =   1410
   End
   Begin VB.TextBox txt 
      Alignment       =   1  'Right Justify
      Height          =   315
      Index           =   23
      Left            =   11040
      TabIndex        =   82
      Text            =   "23"
      Top             =   5400
      Visible         =   0   'False
      Width           =   525
   End
   Begin MSMask.MaskEdBox mskOS 
      Height          =   360
      Index           =   10
      Left            =   1875
      TabIndex        =   41
      Top             =   4755
      Width           =   1260
      _ExtentX        =   2223
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
      Bindings        =   "frmSalesMillBill.frx":05A6
      Left            =   30
      Top             =   870
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
      Left            =   8100
      TabIndex        =   9
      Top             =   1875
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
      Left            =   8100
      TabIndex        =   15
      Top             =   210
      Visible         =   0   'False
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
      Index           =   2
      Left            =   8100
      TabIndex        =   4
      Top             =   1410
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
      Height          =   360
      Index           =   34
      Left            =   1260
      TabIndex        =   45
      Top             =   6165
      Width           =   1170
      _ExtentX        =   2064
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
      Height          =   375
      Index           =   7
      Left            =   6480
      TabIndex        =   10
      Top             =   2325
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
   Begin VB.Label Label8 
      BackStyle       =   0  'Transparent
      Caption         =   "G.P. No    :"
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
      Height          =   510
      Left            =   7680
      TabIndex        =   141
      Top             =   2340
      Width           =   420
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "Account"
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
      Left            =   3675
      TabIndex        =   48
      Top             =   5820
      Width           =   1110
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "Delivery Date       :"
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
      Height          =   510
      Left            =   5670
      TabIndex        =   12
      Top             =   2340
      Width           =   660
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   " Date"
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
      Left            =   1275
      TabIndex        =   44
      Top             =   5820
      Width           =   1170
   End
   Begin VB.Line Line5 
      BorderColor     =   &H00000080&
      X1              =   4500
      X2              =   4500
      Y1              =   4155
      Y2              =   5715
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
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
      Left            =   7200
      TabIndex        =   107
      Top             =   4365
      Width           =   255
   End
   Begin VB.Line Line1 
      BorderColor     =   &H000000C0&
      X1              =   255
      X2              =   9375
      Y1              =   3405
      Y2              =   3405
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Inv. &No  :"
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
      Left            =   5670
      TabIndex        =   2
      Top             =   1410
      Width           =   840
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Par&ty        :"
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
      TabIndex        =   5
      Top             =   1410
      Width           =   1440
   End
   Begin VB.Label Label14 
      BackStyle       =   0  'Transparent
      Caption         =   "B. No.    :"
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
      Left            =   5670
      TabIndex        =   7
      Top             =   1875
      Width           =   1350
   End
   Begin VB.Label Label34 
      BackStyle       =   0  'Transparent
      Caption         =   "LR  No.  :"
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
      Left            =   5910
      TabIndex        =   13
      Top             =   210
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Line Line2 
      BorderColor     =   &H000000C0&
      X1              =   5640
      X2              =   5640
      Y1              =   1260
      Y2              =   2865
   End
   Begin VB.Line Line3 
      BorderColor     =   &H000000C0&
      X1              =   240
      X2              =   9330
      Y1              =   1245
      Y2              =   1245
   End
   Begin VB.Label Label10 
      BackStyle       =   0  'Transparent
      Caption         =   "DbNt Amount"
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
      Left            =   6135
      TabIndex        =   50
      Top             =   5820
      Width           =   1410
   End
   Begin VB.Line Line4 
      BorderColor     =   &H000000C0&
      X1              =   240
      X2              =   9330
      Y1              =   5715
      Y2              =   5715
   End
   Begin VB.Label Label45 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Amount"
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
      Height          =   360
      Left            =   7755
      TabIndex        =   28
      Top             =   3105
      Width           =   1380
   End
   Begin VB.Label Label44 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Rate Per"
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
      Left            =   6630
      TabIndex        =   26
      Top             =   3105
      Width           =   1035
   End
   Begin VB.Label Label43 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Rate"
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
      Height          =   360
      Left            =   5295
      TabIndex        =   24
      Top             =   3105
      Width           =   1320
   End
   Begin VB.Label Label42 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Weight"
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
      Height          =   360
      Left            =   4155
      TabIndex        =   22
      Top             =   3105
      Width           =   1125
   End
   Begin VB.Label Label41 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Count Name"
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
      Height          =   360
      Left            =   420
      TabIndex        =   18
      Top             =   3105
      Width           =   2685
   End
   Begin VB.Label Label40 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Bag"
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
      Height          =   360
      Left            =   3195
      TabIndex        =   20
      Top             =   3105
      Width           =   915
   End
   Begin VB.Shape Shape5 
      BorderColor     =   &H00C00000&
      BorderWidth     =   5
      FillColor       =   &H00C00000&
      Height          =   7545
      Left            =   0
      Top             =   -345
      Width           =   10830
   End
   Begin VB.Shape Shape3 
      BorderColor     =   &H00000080&
      FillColor       =   &H00000080&
      FillStyle       =   0  'Solid
      Height          =   120
      Left            =   -105
      Top             =   7230
      Width           =   10875
   End
   Begin VB.Label Label38 
      BackStyle       =   0  'Transparent
      Caption         =   "Total Amount"
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
      Left            =   7605
      TabIndex        =   52
      Top             =   5820
      Width           =   1650
   End
   Begin VB.Label Label37 
      BackStyle       =   0  'Transparent
      Caption         =   "&Rate"
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
      Left            =   2475
      TabIndex        =   46
      Top             =   5820
      Width           =   1065
   End
   Begin VB.Label Label36 
      BackStyle       =   0  'Transparent
      Caption         =   "Db Nt No."
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
      Left            =   330
      TabIndex        =   42
      Top             =   5820
      Width           =   900
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
      ForeColor       =   &H00000080&
      Height          =   360
      Left            =   6780
      TabIndex        =   68
      Top             =   6735
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.Label Label12 
      BackStyle       =   0  'Transparent
      Caption         =   "D&ue Days     :"
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
      Left            =   540
      TabIndex        =   38
      Top             =   4335
      Width           =   1605
   End
   Begin VB.Label Label30 
      BackStyle       =   0  'Transparent
      Caption         =   "Total    :"
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
      Height          =   360
      Left            =   6540
      TabIndex        =   36
      Top             =   5205
      Width           =   1605
   End
   Begin VB.Label Label17 
      BackStyle       =   0  'Transparent
      Caption         =   "Due Date      :"
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
      Left            =   540
      TabIndex        =   40
      Top             =   4755
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
      TabIndex        =   101
      Top             =   45
      Width           =   2475
   End
   Begin VB.Label Label28 
      BackStyle       =   0  'Transparent
      Caption         =   "ROff      :"
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
      Left            =   6540
      TabIndex        =   34
      Top             =   4755
      Width           =   1605
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
      TabIndex        =   99
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
      TabIndex        =   98
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
      TabIndex        =   97
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
      TabIndex        =   96
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
      TabIndex        =   95
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
      TabIndex        =   94
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
      TabIndex        =   93
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
      TabIndex        =   92
      Top             =   7590
      Width           =   3135
   End
   Begin VB.Label Label16 
      BackStyle       =   0  'Transparent
      Caption         =   "TaxT&ype  :"
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
      Left            =   4650
      TabIndex        =   30
      Top             =   4335
      Width           =   1020
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
      TabIndex        =   91
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
      Top             =   735
      Width           =   1470
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "&Broker       :"
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
      Top             =   2340
      Width           =   1395
   End
   Begin VB.Label LblSaleType 
      BackColor       =   &H00C00000&
      Caption         =   "   Mill Bill   ( Direct Sales From Mill )"
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
      Height          =   390
      Left            =   0
      TabIndex        =   90
      Top             =   0
      Width           =   10815
   End
   Begin VB.Shape Shape1 
      FillColor       =   &H00BDDFD1&
      FillStyle       =   0  'Solid
      Height          =   6225
      Left            =   9540
      Shape           =   4  'Rounded Rectangle
      Top             =   615
      Width           =   1155
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H00000080&
      FillColor       =   &H00EDFDFE&
      FillStyle       =   0  'Solid
      Height          =   2280
      Left            =   225
      Top             =   600
      Width           =   9135
   End
   Begin VB.Shape Shape4 
      BorderColor     =   &H000000C0&
      FillColor       =   &H00EDFDFE&
      FillStyle       =   0  'Solid
      Height          =   1095
      Left            =   240
      Top             =   2955
      Width           =   9135
   End
   Begin VB.Shape Shape6 
      BorderColor     =   &H000000C0&
      FillColor       =   &H00EDFDFE&
      FillStyle       =   0  'Solid
      Height          =   2625
      Left            =   225
      Top             =   4140
      Width           =   9150
   End
End
Attribute VB_Name = "frmSalesMillBill"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Const tblName As String = "tblAddLess"
Const IdField As String = "Vno"
Const OrderField As String = "Vdt,Vno"
Const MaxNo As Long = 55
Dim Rstbl As Recordset '-- TblAddLess
Dim rstblSub As Recordset '-- TblVouSub
Dim varBookMark As Variant
Dim FirstTimeFlag As Boolean
Dim rsVou As Recordset '--TblVoucher
Dim rsOS As Recordset '--TblOutStanding
Dim OldgFrmTypeStr As String
Dim FormAction As Integer
Dim frmTypeStr As String
Dim SalesAccode As Long

Private Sub Cbo_KeyPress(Index As Integer, KeyAscii As Integer)
If Index = 12 Then KeyAscii = 0 '--Tax
End Sub
Private Sub cbo_LostFocus(Index As Integer)
If Index = 12 And FormAction = vbDataActionAddNew Then '--Tax %
    txt(13) = GProcGetColumnValue("tblMastNarration", "Narration", Cbo(12), "S", "MastTaxRate", "N")
    txt(13) = Format(txt(13), FStr)
End If
End Sub

'Dim GridFlag As Boolean
Private Sub cmdBtn_Click(Index As Integer)
On Error GoTo ErrorRoutine
Dim X As String
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
        X = " TblMastAccount as TblMastAc1 "
    Else '-- Oracle
        X = " TblMastAccount TblMastAc1 "
    End If
    gClsSearch.SearchMultiField "tbladdLess,TblMastAccount, " & X, "AdInvNo,Vno,Vdt,TblMastAccount.AcName,TblMastAc1.AcName,AdbillAmt", Array("Invoice No", "Vno", "Date", "Party", "Sales A/c", "Amount"), Array(1000, 0, 1200, 2500, 2500, 1500), " AdAcDrCode=TblMastAccount.AcCode and AdAcCrcode=TblMastAc1.AcCode and AdAcCrCode=" & SalesAccode & " and vtype='" & frmTypeStr & "' and VYear=" & gCYear, "", "AdInvNo,Vdt", 0, 0
    If gClsSearch.SearchMultiRetCol(0) <> "" Then
        Rstbl.CancelUpdate
        Rstbl.MoveFirst
        Rstbl.Find "Vno = " & gClsSearch.SearchMultiRetCol(1)
        gClsSearch.SearchMultiRetCol(0) = ""
        ReadFields
    End If
Case 6 'Add
    FormAction = vbDataActionAddNew
    txt(55) = "A"
    Call GProcActivateControls(Me, True)
    SetControlEd
    Call GProcSetButtons(Me, vbDataActionAddNew)
    Call GProcClearForm(Me, Rstbl, MaxNo, True)
    If FirstTimeFlag = True Then
        FirstTimeFlag = False
    Else
       txt(36).SetFocus
    End If
    FillCombo
    txt(0) = GProcGenerateId("TblAddLess", "Vno", "Vtype='" & frmTypeStr & "' and VYear=" & gCYear)
    mskDt(2) = GProcGenerateVDt("tblAddLess", "Vdt", "Vtype='" & frmTypeStr & "' and VYear=" & gCYear)
    mskDt(34) = mskDt(2)
    '-- Sales A/C
    txtTemp(3) = GProcGetColumnValue("tblMastAccount", "AcCode", CStr(SalesAccode), "N", "AcName", "S")
    txtTemp_LostFocus (3)
Case 7 'Modify
    FormAction = vbDataActionUpdate
    Hank = ""
    'Call GProcChkPass(frmTypeStr, FormAction, txt(36), txt(2), txtTemp(4), gUserName, Hank)
    'If gChkPassTrueFalse = True Then
        txt(55) = "M"
        FormAction = vbDataActionUpdate
        If IsRelatedRecord = False Then
            Call GProcActivateControls(Me, True)
            mskDt(2).SetFocus
            SetControlEd
        Else '--- Allow Only DB Note Modify
            Call GProcActivateControls(Me, False)
            SetControlEd
            txttempOS(6).Enabled = True '-- Broker
            mskDt(34).Enabled = True '- Db Dt
            txt(22).Enabled = True '- Db Amt
            mskDt(34).SetFocus
        End If
        '--Update Stock
        GProcUpdateItCtrlStock txtVSub(12), CLng(txtVSub(5)), CDbl(txtVSub(6)), True
        Call GProcSetButtons(Me, vbDataActionUpdate)
        txt(36).Enabled = False
    'Else
    '   FormAction = vbDataActionClose
    'End If
Case 8 'Delete
    FormAction = vbDataActionDelete
    Hank = ""
    Call GProcChkPass(frmTypeStr, FormAction, txt(36), txt(2), txtTemp(4), gUserName, Hank)
    If gChkPassTrueFalse = True Then
        'If mURecDel = True Then
          If IsRelatedRecord = False Then
            If MsgBox("Do You Really Delete This Record?", vbQuestion + vbYesNo + vbDefaultButton1, "Confirm Delete") = vbYes Then
                If Rstbl.RecordCount > 1 Then
                     '-- Delete Related Record
                    DeleteRelatedRecord
                    Rstbl.Delete
                   '--Update Stock
                    GProcUpdateItCtrlStock txtVSub(12), CLng(txtVSub(5)), CDbl(txtVSub(6)), False, True
                    cmdBtn_Click (1)
                ElseIf Rstbl.RecordCount = 1 Then
                    '-- Delete Related Record
                    DeleteRelatedRecord
                    Rstbl.Delete
                    '--Update Stock
                    GProcUpdateItCtrlStock txtVSub(12), CLng(txtVSub(5)), CDbl(txtVSub(6)), False, True
                    Call GProcActivateControls(Me, False)
                    Call GProcSetButtons(Me, vbDataActionCancel)
                    Form_Load
                End If
            End If
        End If
    End If
Case 9 'Print
    Frame1.Caption = LblSaleType & " Bill Printing"
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
    txtPInvno(0).SetFocus
    Check1.Enabled = True
    Check1.Value = 1
Case 10 'Exit
'    DoEvents
    Unload Me
    Exit Sub
Case 11 'Label
    'frmLabel.Show 1
Case 12 'Optional  Debit Note Print
    Frame2.Caption = LblSaleType & " Debit Not Printing"
    '-- Inv No
    If IsNumeric(txt(36)) = False Then
        txtDbNtFromNo.Text = Mid(txt(36).Text, 5)
        txtDbNtToNo.Text = Mid(txt(36).Text, 5)
    Else
        txtDbNtFromNo.Text = txt(36).Text
        txtDbNtToNo.Text = txt(36).Text
    End If
    Frame2.Visible = True
    DoEvents
    txtDbNtFromNo.Enabled = True
    txtDbNtToNo.Enabled = True
    txtDbNtFromNo.SetFocus
Case 13 'Save
    txt(52) = CDate(Date)
    If ValidateData = True Then
        If FormAction = vbDataActionAddNew Then txt(0) = GProcGenerateId(tblName, IdField, "vtype='" & frmTypeStr & "' and VYear=" & gCYear)
        CalAmount
        FillTxtFromTemp
        '-- Save Record
        GProcSaveRecord Me, Rstbl, FormAction, MaxNo
        SaveRelatedRecord
        '--Update Stock
        GProcUpdateItCtrlStock txtVSub(12), CLng(txtVSub(5)), CDbl(txtVSub(6))
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
txt(14).Enabled = False '--Tax
txt(11).Enabled = False '--Amt
txt(20).Enabled = False '--Debit No
txt(21).Enabled = False '--Debit Amt
'mskDt(34).Enabled = False '--Db Date
txttmpAmt.Enabled = False
txtTemp(3).Enabled = False '--Sales A/c
'--- V Sub
'txtVSub(9).Enabled = False '--Sub Amt
'-- Out Standing
mskOS(10).Enabled = False '--Due Dt
'-- count
If txt(8) <> 0 Then
    txttmpVSub(4).Enabled = False
Else
    mskDt(9) = "__/__/____"
    txttmpVSub(4).Enabled = True
End If
mskDt(9).Enabled = False '--Book vDt
txtCrBalance.Enabled = False
txtDrBalance.Enabled = False
CheckForNumberTmp
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
Dim X As String
Dim BillDt
Dim DueDt
'--- Account Code
MillCode = GProcGetColumnValue("tblMastAccount", "AcName", txtTemp(3), "S", "AcCode", "N") '-- -
CrCode = GProcGetColumnValue("tblMastAccount", "AcName", txtTemp(3), "S", "AcCode", "N")
DrCode = GProcGetColumnValue("tblMastAccount", "AcName", txtTemp(4), "S", "AcCode", "N")
BrkCode = GProcGetColumnValue("tblMastAccount", "AcName", txttempOS(6), "S", "AcCode", "N")
TaxAcCode = GProcGetColumnValue("tblMastNarration", "Narration", Cbo(12), "S", "MastTaxSAAcCode", "N")
DebitNtAcCode = GProcGetColumnValue("tblMastAccount", "AcName", txtTemp(48), "S", "AcCode", "N") 'gMillBillDbNtAcCode
'-- Delete Records
DeleteRelatedRecord
'--- TblVouSub
X = "select * from tblVouSub where vtype='" & frmTypeStr & "' and vno=" & txt(0) & " and " _
& " VYear=" & gCYear
GProcRstOpen rsVouSub, X, "O"
rsVouSub.AddNew
rsVouSub!Vno = CLng(txt(0)) '--Vno(0)
rsVouSub!Vtype = txt(1) '--Vtype(1)
rsVouSub!Vdt = CDate(txt(2)) '--Vdt(2)
rsVouSub!VSubCtrNo = 1 '--Sr No(3)
'--Itcode(4)
ItCode = GProcGetColumnValue("tblMastItem", "ItName", txttmpVSub(4), "S", "ItCode", "N", "ItMillcode=" & MillCode) '--Item Code
rsVouSub!vsubITCODE = ItCode
rsVouSub!vsubBAG = CLng(txtVSub(5))  '--Bag(5)
rsVouSub!vsubwt = CDbl(txtVSub(6)) '-- Wt(6)
rsVouSub!vsubRt = CDbl(txtVSub(7)) '--Rate(7)
rsVouSub!vsubRTPER = CDbl(txtVSub(8)) '--Rate Per(8)
rsVouSub!VSubPItCtrlNo = CLng(txtVSub(12)) '-- PIt Ctrl No(12)
rsVouSub!vsubAMT = CDbl(txtVSub(9)) '--Amt(9)
'--Godown Code(17)
rsVouSub!VSubGodownCode = GProcGetColumnValue("tblMastNarration", "Narration", "-", "S", "NarrCode", "N", "")
rsVouSub!VSubItCtrlNo = GProcGenerateId("tblVouSub", "VSubItCtrlNo", "VYear=" & gCYear)  '--It Ctrl No
rsVouSub!VYear = txt(23) '--VYear(18)
rsVouSub.Update
rsVouSub.Close
'======================Voucher Creation
If GProcGetColumnValue("tblMastAccount", "AcCode", CStr(TaxAcCode), "N", "AcName", "S") = "-" Then
    '-- Bill Amt
    GProcCreateVoucher FormAction, txt(0), 1, txt(2), txt(1), CrCode, DrCode, txt(11), txt(36), mskDt(2)
    GProcCreateVoucher FormAction, txt(0), 2, txt(2), txt(1), DrCode, CrCode, -txt(11), txt(36), mskDt(2)
Else '-- If Tax A/c
    '-- Bill -Tax Amt
    GProcCreateVoucher FormAction, txt(0), 1, txt(2), txt(1), CrCode, DrCode, txt(11) - txt(14), txt(36), mskDt(2)
    GProcCreateVoucher FormAction, txt(0), 2, txt(2), txt(1), DrCode, CrCode, -(txt(11) - txt(14)), txt(36), mskDt(2)
    '-- Tax Amt
    If CDbl(txt(14)) > 0 Then
        GProcCreateVoucher FormAction, txt(0), 3, txt(2), txt(1), TaxAcCode, DrCode, txt(14), txt(36), mskDt(2)
        GProcCreateVoucher FormAction, txt(0), 4, txt(2), txt(1), DrCode, TaxAcCode, -txt(14), txt(36), mskDt(2)
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
'-- Show Ledger
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
    SalesAccode = gClsSearch.SearchMultiRetCol(2)
    X = "Select * from " & tblName & " where AdAcCrCode=" & SalesAccode & " and Vtype='" & frmTypeStr & "' and VYear=" & gCYear
    If Trim(OrderField) <> "" Then X = X & " order by " & OrderField
    GProcRstOpen Rstbl, X, "O"
    FirstTimeFlag = True
    '---- Button Visiblity Set As Per User Previlage
    GProcSetButtonVisiblity Me
    cmdBtn_Click (6)
    '-- Set Related Rs
    SetRelatedRS
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
Rstbl.Close
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

Private Sub mskDt_KeyPress(Index As Integer, KeyAscii As Integer)
If KeyAscii = vbKeyReturn Then Exit Sub
End Sub

Private Sub mskOS_GotFocus(Index As Integer)
GProcSelectBox Me.ActiveControl
End Sub

Private Sub mskOS_KeyPress(Index As Integer, KeyAscii As Integer)
If KeyAscii = vbKeyReturn Then Exit Sub
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
    mskDt(34) = mskDt(2)
End If
If Index = 2 Then ShowAccountBalance
End Sub
Private Sub txt_GotFocus(Index As Integer)
GProcSelectBox Me.ActiveControl
End Sub

Private Sub txt_KeyPress(Index As Integer, KeyAscii As Integer)
If KeyAscii = vbKeyReturn Then Exit Sub
KeyAscii = GProcValidateKey(Rstbl, Index, KeyAscii, txt(Index).Text)
If Index = 8 Then '--Booking Mill bill
    PrepareBooking
    KeyAscii = 0
End If
End Sub

Private Sub txttmpVSub_GotFocus(Index As Integer)
GProcSelectBox Me.ActiveControl
End Sub
Private Sub txttmpVSub_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyF3 Then
Select Case Index
    Case 4 '--count
        GProcShowForm frmMastItem, frmMain.mnuMstfrm(3), Array(frmMastItem.txtTemp(9)), Array(txtTemp(3))
End Select
End If
End Sub
Private Sub txttmpVSub_KeyPress(Index As Integer, KeyAscii As Integer)
Dim LeftPos As Long
Dim TopPos As Long
Dim MillCode As Long
If KeyAscii = 13 Then Exit Sub
TopPos = txttmpVSub(Index).Top + Me.Top + 650
LeftPos = txttmpVSub(Index).Left + Me.Left
Select Case Index
    Case 4 '--count
        MillCode = GProcGetColumnValue("TblMastAccount", "AcName", txtTemp(3), "S", "Accode", "N")
        gClsSearch.SearchMultiField "tblMastItem", "ItName,ItStdRatePer,ItStdPack", Array("Count Name", "Rate Per", "Std. Packing"), Array(txttmpVSub(Index).Width, 1000, 1200), "ITMillCode=" & MillCode, Chr(KeyAscii), "ItName", LeftPos, TopPos, True
        KeyAscii = 0
        If gClsSearch.SearchMultiRetCol(0) <> "" Then
            txttmpVSub(Index) = gClsSearch.SearchMultiRetCol(0)
            txtVSub(8) = gClsSearch.SearchMultiRetCol(1) '-- RatePer
        End If
End Select
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
If (Index = 9) Then
   If txtVSub(6) > 0 Then
        txtVSub(7) = txtVSub(9) * txtVSub(8) / txtVSub(6)
   Else
        txtVSub(7) = 0
   End If
End If
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
CheckForNumber
SaveBtnEd
CalAmount
SetControlEd
End Sub
Private Sub CheckForNumber()
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
            GName = GProcGetColumnValue("TblMastGroup", "AgCode", 90029, "N", "AgName", "S")
            GProcShowForm frmMastAccount, frmMain.mnuMstfrm(0), Array(frmMastAccount.txtTemp(2)), Array(GName)
        Case 4 '-- Party
            GName = GProcGetColumnValue("TblMastGroup", "AgCode", 90017, "N", "AgName", "S")
            GProcShowForm frmMastAccount, frmMain.mnuMstfrm(0), Array(frmMastAccount.txtTemp(2)), Array(GName)
        Case 12 '--TaxType
            'OldgFrmTypeStr = gfrmTypeStr
            GProcShowForm frmMastNarrationTax, frmMain.mnuMstfrm(6)
            'gfrmTypeStr = OldgFrmTypeStr
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
    Case 4, 48 '-- Party   Alias Account
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
        gClsSearch.SearchMultiField "QryMastaccount", "Account,AcName,AcCity,AgName", Array("Sales Account", "AcName", "City", "Group"), Array(3000, 0, 2000, 3000), "(AgCode=90029 or GpCode=90029)", Chr(KeyAscii), "AcName", LeftPos, TopPos, True
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
        gClsSearch.SearchMultiField "QryMastaccount", "Account,AcName,AcCity,AgName", Array("Mill", "AcName", "City", "Group"), Array(3000, 0, 2000, 3000), "(AgCode=90029 or GpCode=90029)", Chr(KeyAscii), "AcName", LeftPos, TopPos, True
        KeyAscii = 0
        If gClsSearch.SearchMultiRetCol(0) <> "" And gClsSearch.SearchMultiRetCol(1) = "" Then
            txtTemp(Index) = gClsSearch.SearchMultiRetCol(0)
        Else
            If gClsSearch.SearchMultiRetCol(1) <> "" Then txtTemp(Index) = gClsSearch.SearchMultiRetCol(1)
        End If
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
Dim AcCode As Long
SaveBtnEd
ShowAccountBalance
If (Index = 3 And FormAction = vbDataActionAddNew) Or txt(36) = "" Then
     AcCode = GProcGetColumnValue("TblMastAccount", "AcName", txtTemp(3), "S", "AcCode", "N")
     txt(36) = GProcGenerateInvNo(frmTypeStr, AcCode) '--Inv No
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
    gClsSearch.SearchMultiField "QryMastaccount", "Account,AcName,AcCity,AgName", Array("Broker", "AcName", "City", "Group"), Array(3000, 0, 2000, 3000), "(AgCode=90016 or GpCode=90016)", Chr(KeyAscii), "AcName", LeftPos, TopPos, True
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
If CDbl(txt(0)) = 0 Or txt(36) = "" Or GProcIsDateValid(mskDt(2)) = False Or txtTemp(3) = "" Or txtTemp(4) = "" Or CDbl(txt(11)) = 0 Then
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
Dim MillCode As Long
Dim ItCode As Long
Dim AcCode As Long
Dim GName As String
Dim Rs1 As Recordset
Dim rstAddCheck As Recordset
Dim rstModiCheck As Recordset
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
        MsgBox "Check Voucher Date.", vbInformation + vbOKOnly, Me.Caption
        mskDt(2).SetFocus
        Exit Function
    End If
    '--- Delivery dt
    If IsDate(mskDt(7)) = False Then
        MsgBox "Check Delivery Date.", vbInformation + vbOKOnly, Me.Caption
        mskDt(7).SetFocus
        Exit Function
    End If
    If CDate(mskDt(7)) < CDate(mskDt(2)) Then
        MsgBox "Check Delivery Date.", vbInformation + vbOKOnly, Me.Caption
        mskDt(7).SetFocus
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
           GName = GProcGetColumnValue("TblMastGroup", "AgCode", 90029, "N", "AgName", "S")
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
    '--- Broker
        If txttempOS(6) = "" Then txttempOS(6) = "-"
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
    '-- Item
    If Trim(txttmpVSub(4)) = "" Then
        MsgBox "Check Count.", vbInformation + vbOKOnly, Me.Caption
        Exit Function
    End If
    MillCode = GProcGetColumnValue("TblMastAccount", "AcName", txtTemp(3), "S", "Accode", "N")
    '--Create Item
    If GProcGetColumnValue("tblMastItem", "ItName", txttmpVSub(4), "S", "ITCode", "N", "ITMillCode=" & MillCode) = 0 Then
       GProcShowForm frmMastItem, frmMain.mnuMstfrm(3), Array(frmMastItem.txt(1), frmMastItem.txtTemp(9)), Array(txttmpVSub(4), txtTemp(3))
       Exit Function
    End If
    '---- Bag
    If CDbl(txtVSub(5)) = 0 Then
        MsgBox "Check Bag.", vbInformation + vbOKOnly, Me.Caption
        txtVSub(5).SetFocus
        Exit Function
    End If
    '---- Wt
    If CDbl(txtVSub(6)) = 0 Then
        MsgBox "Check Weight.", vbInformation + vbOKOnly, Me.Caption
        txtVSub(6).SetFocus
        Exit Function
    End If
    '---- Rate
    If CDbl(txtVSub(7)) = 0 Then
        MsgBox "Check Rate.", vbInformation + vbOKOnly, Me.Caption
        txtVSub(7).SetFocus
        Exit Function
    End If
    '---- Amount
    If CDbl(txt(11)) <= 0 Then
        MsgBox "Check Voucher Amount.", vbInformation + vbOKOnly, Me.Caption
        Exit Function
    End If
    '-- Debit Amt
   ' If CDbl(txt(21)) > 0 Then mskDt(34) = mskDt(2)
    If CDbl(txt(21)) > 0 And GProcIsDateValid(mskDt(34)) = False Then
        MsgBox "Check Debit Date.", vbInformation + vbOKOnly, App.Title
        mskDt(34).SetFocus
        Exit Function
    End If
    '-- Debit A/c
'    If gMillBillDbNtAcCode = GProcGetColumnValue("tblmastAccount", "AcName", "-", "S", "AcCode", "N") Then
'        DBNTAcCode = 0
'    Else
'        DBNTAcCode = gMillBillDbNtAcCode
'    End If
    If CDbl(txt(21)) > 0 And (txtTemp(48) = "" Or txtTemp(48) = "-") Then
        MsgBox "Check Mill Bill Debit Note Account.", vbInformation + vbOKOnly, Me.Caption
        GProcShowForm frmSetting, frmMain.mnuMstSetting(0)
        Exit Function
    End If
    '-- Booking No
    '-- Validate Wt,Bag
    If CLng(txt(8)) > 0 Then
        ItCode = GProcGetColumnValue("tblMastItem", "ItName", txttmpVSub(4), "S", "ItCode", "N", " ItMillCode=" & MillCode)
        AcCode = GProcGetColumnValue("TblMastAccount", "AcName", txtTemp(4), "S", "Accode", "N")
        If FormAction = vbDataActionAddNew Then '--Add
            X = "Select (vsubbag-(vsubissbag)) as BalBag," _
            & " (vsubwt-(vsubisswt)) as BalWt from tblvouSub where VSUBBOOKACCODE=" & AcCode & " and  vsubitctrlno=" & CLng(txtVSub(12)) & " and vsubitcode=" & ItCode & " and "
        Else '--modify
            X = "Select (vsubbag-(vsubissbag-vSubTmpBag)) as BalBag," _
            & " (vsubwt-(vsubisswt-vSubTmpWt)) as BalWt from tblvouSub where VSUBBOOKACCODE=" & AcCode & " and vsubitctrlno=" & CLng(txtVSub(12)) & " and vsubitcode=" & ItCode & " and "
        End If
        If gBackEndDB = gBackEndAccess Then
            X = X + "  vdt<=cdate('" & CDate(mskDt(2)) & "')"
        Else '-- Oracle
            X = X + "  vdt<= to_date('" & CDate(mskDt(2)) & "','dd/MM/yyyy')"
        End If
        GProcRstOpen Rs1, X, "R"
        If Rs1.EOF And Rs1.BOF Then
            MsgBox "Check Booking No.", vbInformation + vbOKOnly, Me.Caption
            Exit Function
        Else
            '-- Balance Qty
'            If CDbl(txtVSub(5)) > Rs1.Fields("BalBag") Then
'                MsgBox "Check Bag.Balance Bags are " & Rs1.Fields("BalBag"), vbInformation + vbOKOnly, Me.Caption
'                txtVSub(5).SetFocus
'                Exit Function
'            End If
        End If
    End If
    '----Duplicate Records
    '----Addition Case '--Invoice No
    AcCode = GProcGetColumnValue("TblMastAccount", "AcName", txtTemp(3), "S", "Accode", "N")
    If FormAction = vbDataActionAddNew Then
        X = "Select AdInvNo From tblAddLess Where AdAcCrCode=" & AcCode & " and AdInvNo='" & Trim(txt(36)) & "' and Vtype='" & frmTypeStr & "' and VYear=" & gCYear
        i = GProcRstOpen(rstAddCheck, X, "R")
        If i > 0 Then
            MsgBox "Record is Available for Sales account of Invoice No. " & CLng(txt(36)), vbCritical + vbOKOnly, "Duplicate Record"
           ' txt(36).SetFocus
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
If IsDate(mskDt(7)) Then '-- Delivery dt
    txt(7) = CDate(mskDt(7))
Else
    txt(7) = ""
End If
If IsDate(mskDt(9)) Then '-- Book dt
    txt(9) = CDate(mskDt(9))
Else
    txt(9) = ""
End If
If IsDate(mskDt(33)) Then '-- LR dt
    txt(33) = CDate(mskDt(33))
Else
    txt(33) = ""
End If
If IsDate(mskDt(34)) Then '-- Debit Note dt
    txt(34) = CDate(mskDt(34))
Else
    txt(34) = ""
End If
'-- Cr Code
txt(3) = GProcGetColumnValue("tblMastAccount", "AcName", txtTemp(3), "S", "AcCode", "N")
'-- Dr Code
txt(4) = GProcGetColumnValue("tblMastAccount", "AcName", txtTemp(4), "S", "AcCode", "N")
'-- TaxType
'If txttemp(12) = "" Then txttemp(12) = "-"
'txt(12) = GProcGetColumnValue("tblMastNarration", "Narration", txttemp(12), "S", "NarrCode", "N")
txt(12) = GProcGetColumnValue("tblMastNarration", "Narration", Cbo(12), "S", "NarrCode", "N")
txt(23) = gCYear
'--Mill Code
txt(5) = txt(3) 'GProcGetColumnValue("tblMastAccount", "AcName", "-", "S", "AcCode", "N")
'--DB Nt A/c
txt(48) = GProcGetColumnValue("tblMastAccount", "AcName", txtTemp(48), "S", "AcCode", "N")
'-- Debit
'-- Debit Amt
gCn.Execute "delete from tblVoucher where Vtype='SN' and Vyear=" & gCYear & " and vno=" & txt(20)
If CDbl(txt(21)) = 0 Then
    txt(20) = 0
    txt(34) = ""
    txt(48) = 0
End If
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
If IsDate(txt(7)) Then '--GpDt =Delivery dt
    mskDt(7) = txt(7)
Else
    mskDt(7) = "__/__/____"
End If
If IsDate(txt(9)) Then '--Book Dt
    mskDt(9) = txt(9)
Else
    mskDt(9) = "__/__/____"
End If
If IsDate(txt(33)) Then '--LR Dt
    mskDt(33) = txt(33)
Else
    mskDt(33) = "__/__/____"
End If
If IsDate(txt(34)) Then '--Debit Note Dt
    mskDt(34) = txt(34)
Else
    mskDt(34) = "__/__/____"
End If
'-- Cr Code
txtTemp(3) = GProcGetColumnValue("tblMastAccount", "AcCode", txt(3), "N", "AcName", "S")
'-- Dr Code
txtTemp(4) = GProcGetColumnValue("tblMastAccount", "AcCode", txt(4), "N", "AcName", "S")
'--DB Nt A/c
txtTemp(48) = GProcGetColumnValue("tblMastAccount", "AcCode", txt(48), "N", "AcName", "S")
'-- TaxType
'txttemp(12) = GProcGetColumnValue("tblMastNarration", "NarrCode", txt(12), "N", "Narration", "S")
Cbo(12) = GProcGetColumnValue("tblMastNarration", "NarrCode", txt(12), "N", "Narration", "S")
'-- Fill Related Record
FillRelatedRecord
'--- Total+ Debit Amt
txttmpAmt = CDbl(txt(11)) + CDbl(txt(21))
ShowAccountBalance
End Sub
Private Sub FillRelatedRecord()
'--- Fill Related Control
SetRelatedRS
'==== Voucher
'txtVou(11) = rsVou.Fields("VbillNo") '--Invoice No
'==== Vou Sub
txtVSub(5) = rstblSub.Fields("VSubBag") '--Bag
txtVSub(6) = rstblSub.Fields("VSubWt") '--Wt
txtVSub(7) = rstblSub.Fields("VSubRt") '--Rate
txtVSub(8) = rstblSub.Fields("VSubRtPer") '--Rate Per
txtVSub(9) = rstblSub.Fields("VSubAmt") '--Amt
txtVSub(12) = rstblSub.Fields("VSUBPITCTRLNO") '--P It Ctrl no
'--Item
txttmpVSub(4) = GProcGetColumnValue("tblMastItem", "ItCode", rstblSub.Fields("VSubITCode"), "N", "ItName", "S")
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
If IsNumeric(txtVSub(12)) = False Then txtVSub(12) = "0"
End Sub
Private Sub CalAmount()
Dim TaxAmt As Double
Dim SubAmt As Double
Dim Amt As Double
Dim ROff As Double
Dim MillCode As Long
Dim StdPack As Double
Dim DbAmt1 As Double
Dim DbAmt As Double
CheckForNumberTmp
CheckForNumber
MillCode = GProcGetColumnValue("TblMastAccount", "AcName", txtTemp(3), "S", "Accode", "N")
StdPack = GProcGetColumnValue("TblMastItem", "ITName", txttmpVSub(4), "S", "ItStdpack", "N", " ItMillCode=" & MillCode)
'-- Wt=Bag * Std Pack
If txtVSub(6) = 0 Then
    txtVSub(6) = txtVSub(5) * StdPack
End If
'-- Due Date
If IsNumeric(txtOS(9)) = False Then txtOS(9) = "0"
txtOS(9) = Format(txtOS(9), PStr)
If IsDate(mskDt(7)) Then
    mskOS(10) = DateAdd("d", txtOS(9), mskDt(7)) '-- Delivery Dt + Due days
Else
    mskOS(10) = "__/__/____"
End If
ROff = CDbl(txt(17))
'---- SubAmt =[Wt*  Rate/Rate Per]
If txtVSub(8) = 0 Then txtVSub(8) = 1
SubAmt = CDbl(txtVSub(6)) * CDbl(txtVSub(7)) / CDbl(txtVSub(8))
SubAmt = Format(SubAmt, FStr)
'---Tax Amt= Sub amt* Tax %/100
TaxAmt = SubAmt * CDbl(txt(13)) / 100
'-- Amt=Sub +Tax +Roff
Amt = SubAmt + TaxAmt + ROff
Amt = GProcMakeRounding(Amt)
txtVSub(9) = SubAmt
txt(10) = SubAmt
txt(14) = TaxAmt
txt(11) = Amt
txt(24) = SubAmt
'-- Debit Amt=Rate*Weight/Rate Per
If txtVSub(8) > 0 Then
    DbAmt1 = CDbl(txt(22)) * CDbl(txtVSub(6)) / CDbl(txtVSub(8))
    DbAmt1 = Format(DbAmt1, FStr)
    If gIsMillBillAmtWithDbNt = True Then '-- Is Mill Bill Amt Include Db Nt  Amt
        DbAmt = DbAmt1 - Amt
    Else
        DbAmt = DbAmt1
    End If
End If
If DbAmt < 0 Then DbAmt = 0
txt(21) = DbAmt
'-- Debit No
If IsDate(mskDt(34)) = False Then mskDt(34) = mskDt(2)
If CDbl(txt(21)) > 0 Then
    If txt(20) = 0 Then txt(20) = GProcGenerateId("TblVoucher", "Vno", "Vtype='SN' and VYear=" & gCYear)
    If txtTemp(48) = "" Then txtTemp(48) = GProcGetColumnValue("tblmastAccount", "AcCode", CStr(gMillBillDbNtAcCode), "N", "AcName", "S")
End If
'--- Total+ Debit Amt
txttmpAmt = CDbl(txt(11)) + CDbl(txt(21))
CheckForNumber
End Sub
Private Sub FillCombo()
Call GProcFillCombo(Cbo(12), "select Narration from tblMastNarration where NarrType ='T' order by Narration")
Cbo(12) = GProcGetColumnValue("tblMastNarration", "NarrCode", CStr(gTaxCodeSM), "N", "Narration", "S")
If FormAction = vbDataActionAddNew Then
    txt(13) = GProcGetColumnValue("tblMastNarration", "Narration", Cbo(12), "S", "MastTaxRate", "N")
    txt(13) = Format(txt(13), FStr)
End If
End Sub

Private Function IsRelatedRecord() As Boolean
Dim Rs1 As Recordset
Dim X As String
'-- Audited
If GProcISAudited(txt(0), txt(1), txt(23)) = True Then
    MsgBox "Record is Audited.You can not Modify/Delete this Record.", vbCritical + vbOKOnly, App.Title
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
        MsgBox "Related Receipt Record " & Rs1.Fields("RecVno") & " is exists.You can not Modify/Delete this Record.", vbCritical + vbOKOnly, App.Title
    End If
    IsRelatedRecord = True
    Exit Function
End If
Rs1.Close
End Function
'---------- Booking Vs Sale
Private Sub PrepareBooking()
Dim PartyAcCode As Long
Dim MillCode As Long
Dim QStr As String
Dim RsBk As Recordset
If IsDate(mskDt(2)) = True Then
    PartyAcCode = GProcGetColumnValue("TblMastAccount", "AcName", txtTemp(4), "S", "Accode", "N")
    MillCode = GProcGetColumnValue("TblMastAccount", "AcName", txtTemp(3), "S", "Accode", "N")
   '--- Add
    If FormAction = vbDataActionAddNew Then
        QStr = " select tblVousub.VNo,tblVousub.vdt,AcName,ItName,vsubbag-(vsubissbag),vsubwt-(vsubisswt),vsubbag,vsubwt,vsubitctrlno from " _
        & " tblVousub,tblMastAccount,tblmastItem,tbladdless Where " _
        & " tblVousub.vno=tbladdless.vno and tblVousub.vType=tbladdless.vType and tblVousub.vYear=tbladdless.vYear and " _
        & " Accode=ADMILLCODE and Itcode=VSUBITCODE and ADMILLCODE=" & MillCode & " and " _
        & " VSUBBOOKACCODE = " & PartyAcCode & " and ((vsubbag-(vsubissbag))> 0) and tblVousub.vtype ='OM'  and VSUBBOOKISCOMPLETED=0 and "
        If gBackEndDB = gBackEndAccess Then
           QStr = QStr + " tblVousub.vdt<=cdate('" & CDate(mskDt(2)) & "')"
        Else '-- Oracle
           QStr = QStr + " tblVousub.vdt <= to_date('" & CDate(mskDt(2)) & "','dd/MM/yyyy')"
        End If
    Else
       QStr = " select tblVousub.VNo,tblVousub.vdt,AcName,ItName,vsubbag-(vsubissbag-VSubTmpBag),vsubwt-(vsubisswt-VSubTmpWt),vsubbag,vsubwt,vsubitctrlno from " _
        & " tblVousub,tblMastAccount,tblmastItem,tbladdless Where " _
        & " tblVousub.vno=tbladdless.vno and tblVousub.vType=tbladdless.vType and tblVousub.vYear=tbladdless.vYear and " _
        & " Accode=ADMILLCODE and Itcode=VSUBITCODE and ADMILLCODE=" & MillCode & " and  " _
        & " VSUBBOOKACCODE = " & PartyAcCode & " and ((vsubbag-(vsubissbag-VSubTmpBag))> 0) and tblVousub.vtype ='OM' and VSUBBOOKISCOMPLETED=0 and "
        If gBackEndDB = gBackEndAccess Then
           QStr = QStr + " tblVousub.vdt<=cdate('" & CDate(mskDt(2)) & "')"
        Else '-- Oracle
           QStr = QStr + " tblVousub.vdt <= to_date('" & CDate(mskDt(2)) & "','dd/MM/yyyy')"
        End If
    End If
    With gClsSearch
        .SearchMultiField QStr, "", Array("Booking No.", "Date", "Mill", "Count", "Bal Bag", "Bal Wt.", "Bag", "Wt", "itctrlno"), Array(1200, 1200, 0, 1500, 1000, 1000, 1000, 1000, 0), "", "", "tblVousub.Vdt,tblVousub.vno", 0, 0, False, , True
        If .SearchMultiRetCol(0) <> "" Then
            QStr = "select tblvousub.*,tblmastItem.ItName from tblvousub,tblMastItem where " _
            & " ItCode=VSUBITCODE and VSUBITCTRLNO= " & CLng(.SearchMultiRetCol(8))
            If GProcRstOpen(RsBk, QStr, "R") > 0 Then
                txt(8) = CLng(.SearchMultiRetCol(0)) '--Book Vno
                mskDt(9) = Format(.SearchMultiRetCol(1), "dd/mm/yyyy") '-- Book Vdt
                txttmpVSub(4) = .SearchMultiRetCol(3) '--count
                txtVSub(5) = .SearchMultiRetCol(4) '-- Bal Bag
                txtVSub(6) = .SearchMultiRetCol(5) '-- Bal Wt
                txtVSub(12) = .SearchMultiRetCol(8) '-- It Ctrl No
                txtVSub(7) = RsBk!vsubRt '-- Rate
                txtVSub(8) = RsBk!vsubRTPER '--Rate Per
                txtVSub(9) = RsBk!vsubAMT '--Amt
                CalAmount
            End If
        End If
        SetControlEd
    End With
End If
End Sub
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
    BillSrNo = GProcGetColumnValue("tblMastAccount", "AcName", txtTemp(3), "S", "AcBillSrNo", "S")
    InvNo = CLng(txtPInvno(0))
    Do While InvNo <= txtPInvno(1)
        If BillSrNo = "" Then
            X = "Select * From tblAddLess where  AdInvNo='" & InvNo & "' and VType='" & frmTypeStr & "' and VYear=" & gCYear
        Else
            X = "Select * From tblAddLess where  AdInvNo='" & BillSrNo & "-" & InvNo & "' and VType='" & frmTypeStr & "' and VYear=" & gCYear
        End If
        If GProcRstOpen(RsSA, X, "R") > 0 Then
            With cryReport
                If Trim(RsSA!Vtype) = "SD" Then
                    .Formulas(0) = "FirmNm = '" & UCase(GProcGetColumnValue("tblMastAccount", "AcCode", RsSA!AdAcCrCode, "N", "AcName", "S")) & "'"
                End If
                .Formulas(10) = "wCSTNo = '" & gCCST & "'"
                .Formulas(11) = "wBSTNo = '" & gCBST & "'"
                .Formulas(12) = "wAmountInWords = '" & GProcAmountInWords(RsSA!AdBillAmt) & "'"
                .Formulas(13) = "wType = '" & RsSA!Vtype & "'"
                .SelectionFormula = "{tblAddLess.AdInvNo} = '" & RsSA!AdInVNo & "' and {tblAddLess.VType}='" & RsSA!Vtype & "' and {tblAddLess.VYear}=" & RsSA!VYear
                '--Report File At  D:\HiReports if exist
                If Dir("D:\HiReports\rptSaleBillPrint.rpt") <> "" Then
                     .ReportFileName = "D:\HiReports\rptSaleBillPrint.rpt"
                Else
                     If Check1.Value = 1 Then
                        .ReportFileName = gReportPath & "rptMillBlPayLtr.rpt"
                     Else
                        .ReportFileName = gReportPath & "rptSaleBillPrint.rpt"
                     End If
                End If
                'Printer.PaperSize = vbPRPSUser
                .Action = 1
            End With
        End If
        InvNo = InvNo + 1
    Loop
Case 2:
    DoEvents
    Frame1.Visible = False
End Select
End Sub
'-- Debit Note Print
Private Sub CmdPrintDbNt_Click(Index As Integer)
Dim RsSA As Recordset
Dim X As String
Dim BillSrNo As String
Dim InvNo As Long
Dim DbNtAcName As String
Select Case Index
Case 0, 1:
    '-- Validate
    If IsNumeric(txtDbNtFromNo) = False Then txtDbNtFromNo = 0
    If IsNumeric(txtDbNtToNo) = False Then txtDbNtToNo = txtDbNtFromNo
    If CLng(txtDbNtToNo) < CLng(txtDbNtFromNo) Then txtDbNtToNo = txtDbNtFromNo
    '-----------------
    If Index = 0 Then GProcCrystalRptPreparation cryReport, "Printer"
    If Index = 1 Then GProcCrystalRptPreparation cryReport, "Window"
    cryReport.WindowTitle = "Debit Note"
    '-- A/C Bill SrNo
    BillSrNo = GProcGetColumnValue("tblMastAccount", "AcName", txtTemp(3), "S", "AcBillSrNo", "S")
    InvNo = CLng(txtDbNtFromNo)
    Do While InvNo <= txtDbNtToNo
        If BillSrNo = "" Then
            X = "Select * From tblAddLess where AdDbNtAmt>0 and AdInvNo='" & InvNo & "' and VType='" & frmTypeStr & "' and VYear=" & gCYear
        Else
            X = "Select * From tblAddLess where AdDbNtAmt>0 and AdInvNo='" & BillSrNo & "-" & InvNo & "' and VType='" & frmTypeStr & "' and VYear=" & gCYear
        End If
        If GProcRstOpen(RsSA, X, "R") > 0 Then
            With cryReport
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
                .Action = 1
            End With
        End If
        InvNo = InvNo + 1
    Loop
Case 2:
    DoEvents
    Frame2.Visible = False
End Select
End Sub
Private Sub txtDbNtFromNo_GotFocus()
GProcSelectBox Me.ActiveControl
End Sub
Private Sub txtDbNtToNo_GotFocus()
GProcSelectBox Me.ActiveControl
End Sub
Private Sub txtDbNtToNo_KeyPress(KeyAscii As Integer)
KeyAscii = GProcNumberOnly(KeyAscii)
End Sub
Private Sub txtDbNtFromNo_KeyPress(KeyAscii As Integer)
KeyAscii = GProcNumberOnly(KeyAscii)
End Sub






