VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Begin VB.Form frmBookingPartyNew 
   BackColor       =   &H00C2E3FC&
   BorderStyle     =   0  'None
   Caption         =   "t"
   ClientHeight    =   8730
   ClientLeft      =   0
   ClientTop       =   15
   ClientWidth     =   10875
   DrawMode        =   10  'Mask Pen
   FillColor       =   &H000040C0&
   FillStyle       =   0  'Solid
   Icon            =   "frmBookingParty New.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   8730
   ScaleWidth      =   10875
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFC0C0&
      Caption         =   "Party Book Printing"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3510
      Left            =   2295
      TabIndex        =   94
      Top             =   1980
      Visible         =   0   'False
      Width           =   5175
      Begin VB.CommandButton cmdPrint 
         Caption         =   "Whatsapp - &Party"
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
         Left            =   180
         TabIndex        =   114
         Top             =   2835
         Width           =   2280
      End
      Begin VB.CommandButton cmdPrint 
         Caption         =   "Whatsapp - &Broker"
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
         Left            =   2655
         TabIndex        =   113
         Top             =   2835
         Width           =   2280
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
         Left            =   1395
         TabIndex        =   112
         Top             =   1530
         Width           =   2535
      End
      Begin VB.CommandButton cmdPrint 
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
         TabIndex        =   99
         Top             =   2100
         Width           =   1335
      End
      Begin VB.CommandButton cmdPrint 
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
         Left            =   1935
         TabIndex        =   98
         Top             =   2100
         Width           =   1335
      End
      Begin VB.TextBox txtno 
         Appearance      =   0  'Flat
         Height          =   360
         Index           =   1
         Left            =   1950
         MaxLength       =   7
         TabIndex        =   97
         Top             =   1080
         Width           =   1575
      End
      Begin VB.CommandButton cmdPrint 
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
         TabIndex        =   96
         Top             =   2100
         Width           =   1335
      End
      Begin VB.TextBox txtno 
         Appearance      =   0  'Flat
         Height          =   360
         Index           =   0
         Left            =   1950
         MaxLength       =   7
         TabIndex        =   95
         Top             =   600
         Width           =   1575
      End
      Begin VB.Label Label18 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "To VNo     :"
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
         TabIndex        =   101
         Top             =   1080
         Width           =   1050
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "From VNo :"
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
         TabIndex        =   100
         Top             =   600
         Width           =   1065
      End
   End
   Begin VB.CheckBox Chk 
      BackColor       =   &H00C0E0FF&
      Caption         =   "Print Ticket Name ?"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H004D177D&
      Height          =   375
      Index           =   0
      Left            =   450
      TabIndex        =   110
      Top             =   5430
      Width           =   2325
   End
   Begin VB.CheckBox Chk 
      BackColor       =   &H00C0E0FF&
      Caption         =   "Print Export ?"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H004D177D&
      Height          =   375
      Index           =   1
      Left            =   3180
      TabIndex        =   109
      Top             =   5430
      Width           =   1785
   End
   Begin VB.CheckBox ChkExGST 
      BackColor       =   &H00C0E0FF&
      Caption         =   "Is Rate Excluding GST ?"
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
      Height          =   405
      Left            =   5565
      TabIndex        =   108
      Top             =   5430
      Width           =   2820
   End
   Begin VB.TextBox txtPartyGrp 
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
      Left            =   1800
      TabIndex        =   6
      Text            =   "txtPartyGrp"
      ToolTipText     =   "Press F3 for Creation"
      Top             =   660
      Width           =   3465
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
      Left            =   1320
      TabIndex        =   7
      Text            =   "txttemp(4)"
      ToolTipText     =   "Press F3 for Creation"
      Top             =   1140
      Width           =   3945
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
      Left            =   1320
      TabIndex        =   8
      Text            =   "txttemp(3)"
      ToolTipText     =   "Press F3 for Creation"
      Top             =   1665
      Width           =   3945
   End
   Begin VB.TextBox txtGridSub 
      Alignment       =   1  'Right Justify
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
      Left            =   3195
      TabIndex        =   93
      Top             =   3060
      Visible         =   0   'False
      Width           =   1035
   End
   Begin MSFlexGridLib.MSFlexGrid msGridSub 
      Height          =   2070
      Left            =   3060
      TabIndex        =   24
      Top             =   3015
      Width           =   2415
      _ExtentX        =   4260
      _ExtentY        =   3651
      _Version        =   393216
      FixedCols       =   0
      RowHeightMin    =   275
      BackColorFixed  =   -2147483647
      ForeColorFixed  =   -2147483634
      BackColorBkg    =   12632256
      AllowUserResizing=   1
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
   Begin VB.TextBox txt 
      Alignment       =   1  'Right Justify
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
      Left            =   10095
      TabIndex        =   92
      Text            =   "47"
      Top             =   9150
      Visible         =   0   'False
      Width           =   360
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
      Height          =   450
      Index           =   46
      Left            =   2025
      TabIndex        =   17
      Text            =   "46"
      Top             =   7740
      Width           =   7185
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
      Height          =   450
      Index           =   45
      Left            =   2025
      TabIndex        =   16
      Text            =   "45"
      Top             =   7260
      Width           =   7185
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
      Index           =   44
      Left            =   2025
      TabIndex        =   15
      Text            =   "44"
      ToolTipText     =   "Press F2 for Selection"
      Top             =   6840
      Width           =   7185
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
      Index           =   43
      Left            =   2025
      TabIndex        =   14
      Text            =   "43"
      ToolTipText     =   "Press F2 for Selection"
      Top             =   6420
      Width           =   7185
   End
   Begin VB.TextBox txt 
      Alignment       =   1  'Right Justify
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
      Left            =   10050
      TabIndex        =   91
      Text            =   "42"
      Top             =   10155
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
      Height          =   405
      Index           =   41
      Left            =   10005
      TabIndex        =   90
      Text            =   "41"
      Top             =   9720
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
      Height          =   405
      Index           =   40
      Left            =   10125
      TabIndex        =   89
      Text            =   "40"
      Top             =   9435
      Visible         =   0   'False
      Width           =   360
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
      Index           =   39
      Left            =   9000
      TabIndex        =   21
      Text            =   "39"
      Top             =   9720
      Visible         =   0   'False
      Width           =   720
   End
   Begin VB.TextBox txt 
      Alignment       =   1  'Right Justify
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
      Left            =   10050
      TabIndex        =   88
      Text            =   "38"
      Top             =   8895
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
      Height          =   405
      Index           =   37
      Left            =   10455
      TabIndex        =   87
      Text            =   "37"
      Top             =   6885
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
      Height          =   405
      Index           =   17
      Left            =   10380
      TabIndex        =   86
      Text            =   "17"
      Top             =   6705
      Visible         =   0   'False
      Width           =   360
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
      TabIndex        =   22
      ToolTipText     =   "To Save Record"
      Top             =   4170
      Width           =   915
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
      TabIndex        =   23
      Top             =   4650
      Width           =   915
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
      Height          =   360
      Index           =   14
      Left            =   9675
      TabIndex        =   20
      Text            =   "14"
      Top             =   8145
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.TextBox txt 
      Alignment       =   1  'Right Justify
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
      Left            =   10095
      TabIndex        =   85
      Text            =   "36"
      Top             =   9675
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
      Left            =   9270
      TabIndex        =   84
      Text            =   "35"
      Top             =   10035
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
      Left            =   7290
      TabIndex        =   83
      Text            =   "34"
      Top             =   9450
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
      Height          =   420
      Left            =   2010
      TabIndex        =   80
      Text            =   "txttmpTotBag"
      Top             =   6000
      Width           =   1230
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
      Height          =   420
      Left            =   4770
      TabIndex        =   79
      Text            =   "txttmpTotWt"
      Top             =   6000
      Width           =   1380
   End
   Begin VB.TextBox txtGrid 
      BackColor       =   &H8000000A&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   495
      TabIndex        =   37
      Top             =   3390
      Visible         =   0   'False
      Width           =   1035
   End
   Begin MSFlexGridLib.MSFlexGrid msGrid 
      Height          =   2535
      Left            =   225
      TabIndex        =   13
      Top             =   2835
      Width           =   9165
      _ExtentX        =   16166
      _ExtentY        =   4471
      _Version        =   393216
      FixedCols       =   0
      RowHeightMin    =   275
      BackColorFixed  =   -2147483647
      ForeColorFixed  =   -2147483634
      BackColorBkg    =   12632256
      AllowUserResizing=   1
      Appearance      =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
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
      Index           =   47
      ItemData        =   "frmBookingParty New.frx":058A
      Left            =   6840
      List            =   "frmBookingParty New.frx":0594
      TabIndex        =   3
      Text            =   "Cbo(47)"
      Top             =   1125
      Width           =   2370
   End
   Begin VB.TextBox txt 
      Alignment       =   1  'Right Justify
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
      Left            =   10125
      TabIndex        =   39
      Text            =   "18"
      Top             =   9720
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
      Index           =   15
      Left            =   10200
      TabIndex        =   40
      Text            =   "15"
      Top             =   9120
      Visible         =   0   'False
      Width           =   315
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
      Height          =   480
      Index           =   0
      Left            =   6840
      TabIndex        =   1
      Text            =   "0"
      Top             =   600
      Width           =   1200
   End
   Begin VB.TextBox txt 
      Alignment       =   1  'Right Justify
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
      Left            =   1800
      TabIndex        =   78
      Text            =   "33"
      Top             =   9450
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
      Left            =   9000
      TabIndex        =   77
      Text            =   "32"
      Top             =   9585
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
      Left            =   2640
      TabIndex        =   76
      Text            =   "31"
      Top             =   9000
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
      Left            =   1845
      TabIndex        =   49
      Text            =   "30"
      Top             =   9675
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
      Index           =   29
      Left            =   900
      TabIndex        =   48
      Text            =   "29"
      Top             =   9495
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
      Index           =   28
      Left            =   1170
      TabIndex        =   47
      Text            =   "28"
      Top             =   9225
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
      Index           =   27
      Left            =   495
      TabIndex        =   46
      Text            =   "27"
      Top             =   9270
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
      Index           =   26
      Left            =   1200
      TabIndex        =   45
      Text            =   "26"
      Top             =   9120
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
      Left            =   10125
      TabIndex        =   44
      Text            =   "25"
      Top             =   9375
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
      Height          =   375
      Index           =   24
      Left            =   8775
      TabIndex        =   41
      Text            =   "24"
      Top             =   9630
      Visible         =   0   'False
      Width           =   390
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
      TabIndex        =   74
      Text            =   "1"
      Top             =   2430
      Width           =   420
   End
   Begin VB.TextBox txt 
      Alignment       =   1  'Right Justify
      Height          =   315
      Index           =   2
      Left            =   11040
      TabIndex        =   63
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
      TabIndex        =   62
      Text            =   "5"
      Top             =   2640
      Width           =   525
   End
   Begin VB.TextBox txt 
      Alignment       =   1  'Right Justify
      Height          =   315
      Index           =   3
      Left            =   11040
      TabIndex        =   61
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
      Left            =   8325
      Style           =   1  'Graphical
      TabIndex        =   54
      Top             =   9810
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
      Left            =   9360
      Style           =   1  'Graphical
      TabIndex        =   53
      ToolTipText     =   "To Add New Record"
      Top             =   9555
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
      TabIndex        =   35
      ToolTipText     =   "Exit"
      Top             =   5235
      Width           =   915
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
      TabIndex        =   52
      Top             =   2115
      Width           =   915
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
      TabIndex        =   29
      ToolTipText     =   "To Delete Record"
      Top             =   3585
      Width           =   915
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
      TabIndex        =   28
      ToolTipText     =   "To Modify Record"
      Top             =   3105
      Width           =   915
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
      TabIndex        =   27
      ToolTipText     =   "To Add New Record"
      Top             =   2640
      Width           =   915
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
      TabIndex        =   51
      Top             =   6120
      Visible         =   0   'False
      Width           =   915
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
      TabIndex        =   30
      ToolTipText     =   "Find Record"
      Top             =   1590
      Width           =   915
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
      TabIndex        =   34
      ToolTipText     =   "Last"
      Top             =   1110
      Width           =   435
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
      TabIndex        =   33
      ToolTipText     =   "First"
      Top             =   1110
      Width           =   435
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
      TabIndex        =   32
      ToolTipText     =   "Next"
      Top             =   630
      Width           =   435
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
      TabIndex        =   31
      ToolTipText     =   "Previous"
      Top             =   630
      Width           =   435
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
      Left            =   1320
      TabIndex        =   9
      Text            =   "txttemp(5)"
      ToolTipText     =   "Press F3 for Creation"
      Top             =   2160
      Width           =   3945
   End
   Begin VB.TextBox txt 
      Alignment       =   1  'Right Justify
      Height          =   315
      Index           =   4
      Left            =   11040
      TabIndex        =   60
      Text            =   "4"
      Top             =   2160
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
      Height          =   480
      Index           =   6
      Left            =   6840
      TabIndex        =   4
      Text            =   "6"
      Top             =   1680
      Width           =   1200
   End
   Begin VB.TextBox txt 
      Alignment       =   1  'Right Justify
      Height          =   315
      Index           =   7
      Left            =   4860
      TabIndex        =   59
      Text            =   "7"
      Top             =   9585
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
      Left            =   6795
      TabIndex        =   36
      Text            =   "8"
      Top             =   9360
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.TextBox txt 
      Alignment       =   1  'Right Justify
      Height          =   315
      Index           =   9
      Left            =   11040
      TabIndex        =   58
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
      Left            =   8055
      TabIndex        =   38
      Text            =   "10"
      Top             =   9555
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
      Height          =   420
      Index           =   11
      Left            =   7620
      TabIndex        =   26
      Text            =   "11"
      Top             =   6000
      Width           =   1590
   End
   Begin VB.TextBox txt 
      Alignment       =   1  'Right Justify
      Height          =   315
      Index           =   12
      Left            =   11040
      TabIndex        =   57
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
      Left            =   8100
      TabIndex        =   42
      Text            =   "13"
      Top             =   9855
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
      Index           =   16
      Left            =   10125
      TabIndex        =   50
      Text            =   "16"
      Top             =   9375
      Visible         =   0   'False
      Width           =   345
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
      Index           =   19
      Left            =   9945
      TabIndex        =   19
      Text            =   "19"
      Top             =   7560
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
      Height          =   360
      Index           =   20
      Left            =   2025
      TabIndex        =   18
      Text            =   "20"
      Top             =   8190
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
      Index           =   21
      Left            =   2070
      TabIndex        =   43
      Text            =   "21"
      Top             =   9270
      Width           =   330
   End
   Begin VB.TextBox txt 
      Alignment       =   1  'Right Justify
      Height          =   315
      Index           =   22
      Left            =   11040
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
      Left            =   11040
      TabIndex        =   55
      Text            =   "23"
      Top             =   5400
      Visible         =   0   'False
      Width           =   525
   End
   Begin MSMask.MaskEdBox mskDt 
      Height          =   375
      Index           =   2
      Left            =   8070
      TabIndex        =   2
      Top             =   600
      Width           =   1140
      _ExtentX        =   2011
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
      Bindings        =   "frmBookingParty New.frx":05B6
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
   Begin MSMask.MaskEdBox mskDt 
      Height          =   375
      Index           =   0
      Left            =   8070
      TabIndex        =   5
      Top             =   1680
      Width           =   1140
      _ExtentX        =   2011
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
   Begin VB.Label Label14 
      BackStyle       =   0  'Transparent
      Caption         =   "Side   :"
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
      Height          =   420
      Left            =   1260
      TabIndex        =   111
      Top             =   8235
      Width           =   720
   End
   Begin VB.Label Label12 
      BackStyle       =   0  'Transparent
      Caption         =   "Party Group :"
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
      TabIndex        =   107
      Top             =   660
      Width           =   1320
   End
   Begin VB.Label Label10 
      BackStyle       =   0  'Transparent
      Caption         =   "Party    :"
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
      TabIndex        =   106
      Top             =   1140
      Width           =   1140
   End
   Begin VB.Label Label11 
      BackStyle       =   0  'Transparent
      Caption         =   "Payment Terms :"
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
      Height          =   420
      Left            =   450
      TabIndex        =   105
      Top             =   6840
      Width           =   1530
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "Remark         :"
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
      Height          =   420
      Left            =   450
      TabIndex        =   104
      Top             =   7260
      Width           =   1350
   End
   Begin VB.Label Label9 
      BackStyle       =   0  'Transparent
      Caption         =   "Delivery Place  :"
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
      Height          =   420
      Left            =   450
      TabIndex        =   103
      Top             =   6420
      Width           =   1485
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Ref. No        :"
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
      Left            =   5520
      TabIndex        =   102
      Top             =   1740
      Width           =   1230
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "&Broker :"
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
      TabIndex        =   12
      Top             =   1665
      Width           =   1140
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "&Rate Per     :"
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
      Left            =   5550
      TabIndex        =   10
      Top             =   1125
      Width           =   1305
   End
   Begin VB.Label Label31 
      BackStyle       =   0  'Transparent
      Caption         =   "Total Weight :"
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
      Height          =   420
      Left            =   3420
      TabIndex        =   82
      Top             =   6000
      Width           =   1230
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
      Height          =   420
      Left            =   660
      TabIndex        =   81
      Top             =   6000
      Width           =   1620
   End
   Begin VB.Shape Shape8 
      BorderColor     =   &H00C00000&
      BorderWidth     =   5
      Height          =   8340
      Left            =   0
      Top             =   360
      Width           =   10845
   End
   Begin VB.Shape Shape3 
      BorderColor     =   &H00C00000&
      FillColor       =   &H00C00000&
      FillStyle       =   0  'Solid
      Height          =   0
      Left            =   -15
      Top             =   6570
      Width           =   10875
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Booking &No :"
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
      Left            =   5550
      TabIndex        =   0
      Top             =   600
      Width           =   1215
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
      Height          =   420
      Left            =   6375
      TabIndex        =   25
      Top             =   6000
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
      ForeColor       =   &H0000FFFF&
      Height          =   495
      Left            =   8340
      TabIndex        =   75
      Top             =   45
      Width           =   2475
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
      TabIndex        =   73
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
      TabIndex        =   72
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
      TabIndex        =   71
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
      TabIndex        =   70
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
      TabIndex        =   69
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
      TabIndex        =   68
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
      TabIndex        =   67
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
      TabIndex        =   66
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
      TabIndex        =   65
      Top             =   8160
      Width           =   1095
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
      TabIndex        =   11
      Top             =   2160
      Width           =   1140
   End
   Begin VB.Label Label6 
      BackColor       =   &H00C00000&
      Caption         =   "     Party Booking"
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
      TabIndex        =   64
      Top             =   0
      Width           =   10815
   End
   Begin VB.Shape Shape1 
      FillColor       =   &H00BDDFD1&
      FillStyle       =   0  'Solid
      Height          =   5280
      Left            =   9540
      Shape           =   4  'Rounded Rectangle
      Top             =   525
      Width           =   1200
   End
   Begin VB.Shape Shape4 
      BorderColor     =   &H00000080&
      FillColor       =   &H00EDFDFE&
      FillStyle       =   0  'Solid
      Height          =   2700
      Left            =   210
      Top             =   5910
      Width           =   9135
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H00000080&
      FillColor       =   &H00EDFDFE&
      FillStyle       =   0  'Solid
      Height          =   2235
      Left            =   210
      Top             =   510
      Width           =   5145
   End
   Begin VB.Shape Shape6 
      BorderColor     =   &H00000080&
      FillColor       =   &H00EDFDFE&
      FillStyle       =   0  'Solid
      Height          =   2190
      Left            =   5415
      Top             =   510
      Width           =   3930
   End
End
Attribute VB_Name = "frmBookingPartyNew"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Const tblName As String = "tblAddLess_Book"
Const IdField As String = "Vno"
Const OrderField As String = "Vdt,Vno"
Const MaxNo As Long = 47
Dim rstbl As Recordset '-- TblAddLess
Dim rstblSub As Recordset '-- TblVouSub
Dim varBookMark As Variant
Dim FirstTimeFlag As Boolean
Dim OldgFrmTypeStr As String
Dim FormAction As Integer
Dim frmTypeStr As String
Dim GridSubFlag As Boolean
Dim IsLockCtrl As Boolean '-- Set True If Related SIT & Allow IsComplete grid entry only
Dim PtyGrp As String
Dim PtyGrpCd As Long
Dim i As Long

Private Sub Cbo_KeyPress(Index As Integer, KeyAscii As Integer)
If Index = 47 Then KeyAscii = 0 '--Rate Type
End Sub
Private Sub cmdBtn_Click(Index As Integer)
'On Error GoTo ErrorRoutine
Dim X As String
Dim RSAc As Recordset
Dim RsTmp As Recordset
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
    gClsSearch.SearchMultiField "tblAddLess_Book,TblMastAccount", "Vno,AcName,Vdt,AdBillAmt", Array("Vno", "Mill", "Vdt", "Amount"), Array(txt(0).Width, 4000, 1080, 1500), " tblAddLess_Book.AdAcDrCode=TblMastAccount.AcCode and vtype='" & frmTypeStr & "' and VYear=" & gCYear, "", "Vno,Vdt", 0, 0
    If gClsSearch.SearchMultiRetCol(0) <> "" Then
        rstbl.MoveFirst
        rstbl.Find "Vno = " & gClsSearch.SearchMultiRetCol(0)
        gClsSearch.SearchMultiRetCol(0) = ""
        ReadFields
    End If
Case 6 'Add
    FormAction = vbDataActionAddNew
    Call GProcActivateControls(Me, True)
    Call GProcSetButtons(Me, vbDataActionAddNew)
    Call GProcClearForm(Me, rstbl, MaxNo, True)
    
    If FirstTimeFlag = True Then
        FirstTimeFlag = False
    Else
        txt(0).SetFocus
    End If
    txt(0) = GProcGenerateId("tblAddLess_Book", "Vno", "Vtype='" & frmTypeStr & "' and VYear=" & gCYear)
    mskDt(2) = GProcGenerateVDt("tblAddLess_Book", "Vdt", "Vtype='" & frmTypeStr & "' and VYear=" & gCYear)
    mskDt(0) = GProcGenerateVDt("tblAddLess_Book", "Vdt", "Vtype='" & frmTypeStr & "' and VYear=" & gCYear)
    IsLockCtrl = False
    SetControlEd
    SetGrid
Case 7 'Modify
    'If IsRelatedRecord = False Then
        IsLockCtrl = IsRelatedRecord
        FormAction = vbDataActionUpdate
        Call GProcActivateControls(Me, True)
        SetControlEd
        Call GProcSetButtons(Me, vbDataActionUpdate)
        txt(0).Enabled = False
        mskDt(2).SetFocus
        If IsLockCtrl = True Then '--Allow  Booking Complete
            Call GProcActivateControls(Me, False)
            msGrid.Enabled = True
        End If
        txt(45).Enabled = True
        txt(46).Enabled = True
        
    'End If
Case 8 'Delete
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
    Frame1.Visible = True
    txtno(0).Enabled = True
    txtno(1).Enabled = True
    txtno(0) = txt(0)
    txtno(1) = txt(0)
    txtno(0).SetFocus
    Check1.Enabled = True
Case 10 'Exit
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

''-----------For SMS
        Dim aa As Boolean
        aa = False
        If MsgBox("Do You Want To Send Whatsapp SMS ? ", vbYesNo) = vbYes Then aa = True

        If aa Then

            Dim HttpReq As New WinHttpRequest
            Dim URL As String
            Dim params As String
            Dim params2 As String
            Dim SmsMblNo As String
            Dim SmsString As String
            Dim Pside As String
            Dim RtNar As String
            Dim PtyName As String
            Dim MillCd As Long
            Dim TktNm As String
            gCn.BeginTrans
            gCn.Execute "Delete from tmpWhatsAppErr where username = '" & gUserName & "'"
            gCn.CommitTrans
      
            PtyName = txttemp(4)
            MillCd = GProcGetColumnValue("TblMastAccount", "AcName", txttemp(5), "S", "AcCode", "N")
            TktNm = GProcGetColumnValue("tblMastItem", "ItName", msGrid.TextMatrix(1, 1), "S", "ItTicket", "S", " ItMillcode=" & MillCd) '--ItCode

            SmsMblNo = GProcGetColumnValue("TblMastAccount", "AcName", txttemp(4), "S", "AcFaxNo", "S")
            
            SmsString = "*" + "Party Booking" + "*" & vbLf
            SmsString = SmsString + "" & vbLf
            SmsString = SmsString + "Firm : " + "*" + gCName + "*" & vbLf
            SmsString = SmsString + "" & vbLf
            SmsString = SmsString + "No     : " + Trim(txt(0)) & vbLf
            SmsString = SmsString + "Date   : " + Trim(CStr(Format(mskDt(2), "dd/mm/yyyy"))) & vbLf
            SmsString = SmsString + "" & vbLf
            SmsString = SmsString + "Party  : " + "*" + Left(txttemp(4), 40) + "*" & vbLf
            SmsString = SmsString + "" & vbLf
            SmsString = SmsString + "Broker : " + "*" + Left(txttemp(3), 40) + "*" & vbLf
            SmsString = SmsString + "Mill   : " + "*" + Left(TktNm, 40) + "*" & vbLf
            SmsString = SmsString + "" & vbLf
            SmsString = SmsString + "Count  : " + "*" + msGrid.TextMatrix(1, 1) + "*" & vbLf
            SmsString = SmsString + "" & vbLf
            SmsString = SmsString + "Bags   : " + msGrid.TextMatrix(1, 10) & vbLf
            SmsString = SmsString + "Kgs.   : " + msGrid.TextMatrix(1, 12) & vbLf
            SmsString = SmsString + "Rate   : " + msGrid.TextMatrix(1, 13) + "/" + msGrid.TextMatrix(1, 14) + " Kg" & vbLf
            If txt(47) = 1 Then
                RtNar = "Rate Excluding GST"
            Else
                RtNar = "Rate Including GST"
            End If
            SmsString = SmsString + "*" + RtNar + "*" & vbLf
            SmsString = SmsString + "" & vbLf
            If txt(20).text = 25 Then
                Pside = "Regular"
            ElseIf txt(20).text = 4 Then
                Pside = "Second Day"
            ElseIf txt(20).text = 1 Then
               Pside = "Next Day"
            End If
            SmsString = SmsString + "Payment Condition-" + Left(txt(44), 30) & vbLf
            SmsString = SmsString + "Delivey Period -" + Left(txt(43), 30) & vbLf
            SmsString = SmsString + "Remark -" + Left(txt(45), 40) & vbLf
            If SmsMblNo = "" Then
               MsgBox ("Party Mobile No Not Found....")
            Else
               SentWhatsApp "Txt", "", SmsString, SmsMblNo, PtyName
            End If
            '----For Broker SMS
            DoEvents
            SmsMblNo = GProcGetColumnValue("TblMastAccount", "AcName", txttemp(3), "S", "AcFaxNo", "S")
            PtyName = txttemp(3)
            If SmsMblNo = "" Then
               MsgBox ("Broker Mobile No Not Found....")
            Else
               SentWhatsApp "Txt", "", SmsString, SmsMblNo, PtyName
            End If
            X = "Select * from tmpWhatsAppErr where username = '" & gUserName & "'"
            If GProcRstOpen(RsTmp, X, "R", gCn) > 0 Then
               GProcWhatsAppError
            End If
           
            MsgBox ("Message send succesfully....")
            DoEvents
        End If
''------End SMS
    End If

Case 14 'Cancel
    txtGrid.Visible = False
    DoEvents
    Call GProcClearForm(Me, rstbl, MaxNo, True)
    If rstbl.RecordCount <= 0 Then
        Call GProcActivateControls(Me, False)
        Call GProcSetButtons(Me, vbDataActionCancel)
    Else
        If FormAction = vbDataActionAddNew Then
            rstbl.CancelUpdate
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
    FormAction = vbDataActionCancel
    DoEvents
    msGrid.Col = 0
    If msGrid.Rows <= 1 Then msGrid.Rows = msGrid.Rows + 1
    msGrid.Row = 1
    txtGrid.Move msGrid.CellLeft + msGrid.Left, msGrid.CellTop + msGrid.Top, msGrid.CellWidth, msGrid.CellHeight
End Select
Exit Sub
ErrorRoutine:
    rstbl.CancelUpdate
    GProcErrorHandler
End Sub
Private Sub SetControlEd()
txt(11).Enabled = False '--Bill Amt
txttmpTotBag.Enabled = False
txttmpTotWt.Enabled = False
End Sub
Private Sub DeleteRelatedRecord()
'--- Voucher Sub
gCn.Execute "delete from tblVouSub_Book where vtype='" & frmTypeStr & "' and vno=" & txt(0) & " and " _
& " VYear=" & gCYear
End Sub
Private Sub SaveRelatedRecord()
Dim i As Long
Dim rsVouSub As Recordset
Dim MillCode As Long
Dim X As String
'--- Mill Code
MillCode = GProcGetColumnValue("tblMastAccount", "AcName", txttemp(5), "S", "AcCode", "N")
'-- Delete Records
DeleteRelatedRecord
'--- Save Grid Records
'--- TblVouSub
X = "select * from tblVouSub_Book where vtype='" & frmTypeStr & "' and vno=" & txt(0) & " and " _
& " VYear=" & gCYear
GProcRstOpen rsVouSub, X, "O"
With msGrid
    i = 1
    Do While i < .Rows
        rsVouSub.AddNew
        rsVouSub!Vno = CLng(txt(0)) '--Vno(0)
        rsVouSub!Vtype = txt(1) '--Vtype(1)
        rsVouSub!Vdt = CDate(txt(2)) '--Vdt(2)
        rsVouSub!VSubCtrNo = i '--Sr No(3)
        '--- Book Ac code
        rsVouSub!VSUBBOOKACCODE = txt(4)    'GProcGetColumnValue("tblMastAccount", "AcName", "-", "S", "AcCode", "N")
        rsVouSub!VSUBBOOKBRKACCODE = txt(3) ' GProcGetColumnValue("tblMastAccount", "AcName", "-", "S", "AcCode", "N")
        rsVouSub!VSubBookMillCode = txt(5) ' GProcGetColumnValue("tblMastAccount", "AcName", "-", "S", "AcCode", "N")
        '--Itcode(4)
        rsVouSub!vsubITCODE = GProcGetColumnValue("tblMastItem", "ItName", .TextMatrix(i, 1), "S", "ItCode", "N", " ItMillCode=" & MillCode)  '--Item Code
        rsVouSub!vsubBAG = CLng(.TextMatrix(i, 10)) '--Bag(5)
        rsVouSub!VSubItStdpack = CDbl(.TextMatrix(i, 11)) '--Std Packing
        rsVouSub!vsubwt = CDbl(.TextMatrix(i, 12)) '-- Wt(6)
        rsVouSub!vsubRt = CDbl(.TextMatrix(i, 13)) '--Rate(7)
        rsVouSub!vsubRTPER = CDbl(.TextMatrix(i, 14)) '--Rate Per(8)
        rsVouSub!vsubAMT = CDbl(.TextMatrix(i, 15)) '--Amt(9)
        '--Godown Code(17)
        rsVouSub!VSubGodownCode = GProcGetColumnValue("tblMastNarration", "Narration", "-", "S", "NarrCode", "N", "")
        rsVouSub!VSUBBOOKISCOMPLETED = IIf(Trim(UCase(.TextMatrix(i, 16))) = "Y", 1, 0) '-Is complete
        '--- It Ctrl No(11)
        If CLng(.TextMatrix(i, 17)) = 0 Then
            rsVouSub!VSubItCtrlNo = GProcGenerateId("tblVouSub_Book", "VSubItCtrlNo", "VYear=" & gCYear)   '--It Ctrl No
        Else
            rsVouSub!VSubItCtrlNo = CLng(.TextMatrix(i, 17))
        End If
        rsVouSub!VSubIssBag = CLng(.TextMatrix(i, 18)) '--Issue Bag(13)
        rsVouSub!VSubIsswt = CDbl(.TextMatrix(i, 19)) '-- Issue Wt(14)
        rsVouSub!VSUBBAGPERLR1 = CLng(.TextMatrix(i, 2)) '--Bag per Lr
        rsVouSub!VSUBBAGPERLR2 = CLng(.TextMatrix(i, 4)) '--Bag per Lr
        rsVouSub!VSUBBAGPERLR3 = CLng(.TextMatrix(i, 6)) '--Bag per Lr
        rsVouSub!VSUBBAGPERLR4 = CLng(.TextMatrix(i, 8)) '--Bag per Lr
        rsVouSub!VSUBNOOFLR1 = CLng(.TextMatrix(i, 3)) '--NO Of LR
        rsVouSub!VSUBNOOFLR2 = CLng(.TextMatrix(i, 5)) '--NO Of LR
        rsVouSub!VSUBNOOFLR3 = CLng(.TextMatrix(i, 7)) '--NO Of LR
        rsVouSub!VSUBNOOFLR4 = CLng(.TextMatrix(i, 9)) '--NO Of LR
        rsVouSub!VYear = txt(23) '--VYear(18)
        rsVouSub.Update
        i = i + 1
     Loop
 End With
 rsVouSub.Close
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case 13:
            If Me.ActiveControl.Name <> "txtGrid" Then Sendkeys "{TAB}"
            'SendKeys "{TAB}"
'        Case 27:
'            If (cmdBtn(13).Enabled = False) Then
'                Unload Me
'            End If
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
    '---Grid
    SetGrid
    SetRelatedRS
    '---- Button Visiblity Set As Per User Previlage
    GProcSetButtonVisiblity Me
    cmdBtn_Click (6)
End Sub
Private Sub SetRelatedRS()
End Sub
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
If FormAction = vbDataActionUpdate Then
    MsgBox "You are about to modify the record. Save the record and then exit.", vbInformation + vbOKOnly, Me.Caption
    Cancel = True
    Exit Sub
End If
rstbl.Requery
rstbl.Close
End Sub

Private Sub mskDt_GotFocus(Index As Integer)
GProcSelectBox Me.ActiveControl
End Sub
Private Sub mskDt_LostFocus(Index As Integer)
If Index = 2 Then
    If GProcIsDateValid(mskDt(Index)) = False Then
        mskDt(Index).SetFocus
        Exit Sub
    End If
End If
SaveBtnEd
End Sub
Private Sub txt_GotFocus(Index As Integer)
GProcSelectBox Me.ActiveControl
End Sub

Private Sub txt_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
Dim LeftPos As Long
Dim TopPos As Long
If KeyCode = vbKeyF2 Then '-- Select
    TopPos = txt(Index).Top + Me.Top + 650
    LeftPos = txt(Index).Left + Me.Left
    Select Case Index
        Case 43, 44 '--Narration
            gClsSearch.SearchMultiField "tblMastNarration", "Narration", Array("Narration"), Array(txt(Index).Width), "NarrType='N'", "", "Narration", LeftPos, TopPos, False
            If gClsSearch.SearchMultiRetCol(0) <> "" Then txt(Index) = gClsSearch.SearchMultiRetCol(0)
    End Select
End If
End Sub

Private Sub txt_LostFocus(Index As Integer)
GProcCheckForNumber Me, rstbl, MaxNo
CalAmount
SaveBtnEd
End Sub
Private Sub txt_KeyPress(Index As Integer, KeyAscii As Integer)
KeyAscii = GProcValidateKey(rstbl, Index, KeyAscii, txt(Index).text)
End Sub

Private Sub txtPartyGrp_KeyPress(KeyAscii As Integer)
Dim LeftPos As Long
Dim TopPos As Long
TopPos = txtPartyGrp.Top + Me.Top + 650
LeftPos = txtPartyGrp.Left + Me.Left
If KeyAscii = 13 Then Exit Sub
    gClsSearch.SearchMultiField "tblMastCompGroup", "CGName,CGCode", Array("Company Group", ""), Array(txtPartyGrp.Width, 0), " ", Chr(KeyAscii), "CGName", LeftPos, TopPos, True
    KeyAscii = 0
    If gClsSearch.SearchMultiRetCol(0) <> "" Then
       txtPartyGrp = gClsSearch.SearchMultiRetCol(0)
       PtyGrpCd = gClsSearch.SearchMultiRetCol(1)
    Else
       txtPartyGrp = "-"
    End If
End Sub

Private Sub txttemp_GotFocus(Index As Integer)
GProcSelectBox Me.ActiveControl
End Sub
Private Sub txttemp_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
Dim GName As String
If KeyCode = vbKeyF3 Then
    Select Case Index
        Case 3 '-- Broker code
            GName = GProcGetColumnValue("TblMastGroup", "AgCode", 90016, "N", "AgName", "S")
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
    Case 3  '-- Broker Alias Account
        gClsSearch.SearchMultiField "QryMastaccount", "Account,AcName,AcCity,AgName", Array("Broker", "AcName", "City", "Group"), Array(3000, 0, 2000, 3000), "(AgCode=90016 or GpCode=90016)", Chr(KeyAscii), "AcName", LeftPos, TopPos, True
        KeyAscii = 0
        If gClsSearch.SearchMultiRetCol(0) <> "" And gClsSearch.SearchMultiRetCol(1) = "" Then
            txttemp(Index) = gClsSearch.SearchMultiRetCol(0)
        Else
            If gClsSearch.SearchMultiRetCol(1) <> "" Then txttemp(Index) = gClsSearch.SearchMultiRetCol(1)
        End If
    Case 4  '-- Party Alias Account
        If PtyGrpCd > 1 And txtPartyGrp <> "" Then
           gClsSearch.SearchMultiField "QryMastaccount", "Account,AcName,AcCity,AgName,AcCmpCode", Array("Party", "AcName", "City", "Group", "Party Group"), Array(3000, 0, 2000, 3000, 1000), "(AcCmpCode= " & PtyGrpCd & ")", Chr(KeyAscii), "AcName", LeftPos, TopPos, True
        Else
           gClsSearch.SearchMultiField "QryMastaccount", "Account,AcName,AcCity,AgName,AcCmpCode", Array("Party", "AcName", "City", "Group", "Party Group"), Array(3000, 0, 2000, 3000, 1000), "(AgCode=90017 or AgCode=90015)", Chr(KeyAscii), "AcName", LeftPos, TopPos, True
        End If
        KeyAscii = 0
        If gClsSearch.SearchMultiRetCol(0) <> "" And gClsSearch.SearchMultiRetCol(1) = "" Then
            txttemp(Index) = gClsSearch.SearchMultiRetCol(0)
        Else
            If gClsSearch.SearchMultiRetCol(1) <> "" Then txttemp(Index) = gClsSearch.SearchMultiRetCol(1)
        End If
        PtyGrp = GProcGetColumnValue("TblMastAccount", "AcName", txttemp(4), "C", "AcCmpCode", "N")   '-Party Group Code
        txtPartyGrp = GProcGetColumnValue("tblmastcompGroup", "CgCode", PtyGrp, "N", "Cgname", "S")
    
'        gClsSearch.SearchMultiField "QryMastaccount", "Account,AcName,AcCity,AgName", Array("Broker", "AcName", "City", "Group"), Array(3000, 0, 2000, 3000), "(AgCode=90017 or GpCode=90017)", Chr(KeyAscii), "AcName", LeftPos, TopPos, True
'        KeyAscii = 0
'        If gClsSearch.SearchMultiRetCol(0) <> "" And gClsSearch.SearchMultiRetCol(1) = "" Then
'            txttemp(Index) = gClsSearch.SearchMultiRetCol(0)
'        Else
'            If gClsSearch.SearchMultiRetCol(1) <> "" Then txttemp(Index) = gClsSearch.SearchMultiRetCol(1)
'        End If
    Case 5  '-- Mill  Alias Account
        gClsSearch.SearchMultiField "QryMastaccount", "Account,AcName,AcCity,AgName", Array("Mill", "AcName", "City", "Group"), Array(3000, 0, 2000, 3000), "((AgCode=90029 or GpCode=90029) and (AcMillType=0 or AcMillType=5))", Chr(KeyAscii), "AcName", LeftPos, TopPos, True
        KeyAscii = 0
        If gClsSearch.SearchMultiRetCol(0) <> "" And gClsSearch.SearchMultiRetCol(1) = "" Then
            txttemp(Index) = gClsSearch.SearchMultiRetCol(0)
        Else
            If gClsSearch.SearchMultiRetCol(1) <> "" Then txttemp(Index) = gClsSearch.SearchMultiRetCol(1)
        End If
        If txttemp(3) = "" Then
            BrkAcCode = GProcGetColumnValue("TblMastAccount", "ACName", txttemp(Index), "S", "AcBrkCode", "N")
            txttemp(3) = GProcGetColumnValue("TblMastAccount", "ACCode", CStr(BrkAcCode), "N", "ACname", "S")
        End If
End Select
End Sub
Private Sub txtTemp_LostFocus(Index As Integer)
SaveBtnEd
End Sub
'---------------- Grid
Private Sub SetGrid()
Dim X As String
Dim i As Long
Set rstblSub = New Recordset
X = "select VSubCtrNo,ItName,VSUBBAGPERLR1,VSUBNOOFLR1,VSUBBAGPERLR2,VSUBNOOFLR2,VSUBBAGPERLR3," _
& "VSUBNOOFLR3,VSUBBAGPERLR4,VSUBNOOFLR4,VSubBag,VSubItStdpack,VSubWt" _
& " ,VSubRt,VSubRtPer,VSubAmt," & IIf(gBackEndDB = gBackEndAccess, "iif(VSUBBOOKISCOMPLETED=1,'Y','')", "decode(VSUBBOOKISCOMPLETED,1,'Y','')") & ",VSUBITCTRLNO,VSUBISSBAG,VSUBISSWt from tblvouSub_Book,TblMastItem where " _
& " VSUBItCODE=ItCode and tblvouSub_Book.vtype='" & frmTypeStr & "' and tblvouSub_Book.vno=" & txt(0) & " and tblvouSub_Book.VYear=" & gCYear & " order by VSubCtrNo"
GProcRstOpen rstblSub, X, "R"
With msGrid
    .FormatString = ">SR|<Count                                   |>Bag/LR|>LR|>Bag/LR|>LR|>Bag/LR|>LR|>Bag/LR|>LR|>Bag |>Std Packing|>Weight      |>Rate            |>Rt Per  |>Amount       |<Is Complete(Y)|||"
    .ColWidth(17) = 0 '-It Ctrl
    .ColWidth(18) = 0 '-Issue Bag
    .ColWidth(19) = 0 '-Issue Wt
    For i = 2 To 9
        .ColWidth(i) = 0
    Next i
    If FormAction = vbDataActionAddNew Then
        .Rows = 2
        GridAddNew
    Else
        .Rows = 1
    End If
End With
'---Format Sub Grid
With msGridSub
    .FormatString = ">Bag / LR    |>No Of LR  "
    .Rows = 5
    .Row = 1
    .Col = 0
    For i = 1 To 4
        .TextMatrix(.Row, 0) = 0 '--Bag/LR
        .TextMatrix(.Row, 1) = 0 '--No of LRS
    Next i
    .Visible = False
    txtGridSub.Visible = False
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
Dim i As Long
    msGrid_EnterCell
'    If msGrid.Col = 10 Then
'        For i = 2 To 9
'
'        Next i
'    End If
End Sub
Private Sub msGrid_GotFocus()
If GridSubFlag = True Then
'-- Fill Grid From Grid Sub
    FillGridFromGridSub
    Exit Sub
End If 'And msGrid.Col = 0
If LCase(Me.ActiveControl.Name) <> "txtgrid" And LCase(Me.ActiveControl.Name) <> "msgrid" Then msGrid.Col = 1
    cmdBtn(10).Cancel = False
End Sub
Private Sub MSGrid_LostFocus()
CalAmount
End Sub
Private Sub msGrid_EnterCell()
If GridSubFlag = True Then Exit Sub
If FormAction <> vbDataActionCancel Then
    CheckForNumberTxtGrid
    DoEvents
    txtGrid = ""
    DoEvents
    If msGrid.text <> "" Then txtGrid.text = msGrid.text
    '--- Lock TextBox
    If IsLockCtrl = False Then
        With msGrid
            Select Case .Col
                Case 0: '---- Sr No
                    txtGrid.Visible = True
                    txtGrid.Locked = True
                    ReNumberGridSrno
                    DoEvents
                    .Col = .Col + 1
'                Case 10:
'                    '-- Fill Grid Sub From Grid
'                    FillGridSubFromGrid
'                    Exit Sub
                Case Else
                    txtGrid.Locked = False
            End Select
        End With
    Else
        If msGrid.Col = 16 Then '-Is complete
            txtGrid.Locked = False
        Else
            txtGrid.Locked = True
        End If
    End If
    DoEvents
    msGrid_KeyPress (0)
    DoEvents
End If
End Sub
Private Sub msGrid_KeyPress(KeyAscii As Integer)
With msGrid
    If KeyAscii = 27 Then '----Esc Key
        KeyAscii = 0
        If (Val(.TextMatrix(.Row, 15)) = 0) And .Rows > 2 Then '-- amt =0
            .RemoveItem (.Row)
            txtGrid.text = ""
        End If
        txtGrid.Visible = False
        Exit Sub
    End If
End With

'If FormAction <> vbDataActionAddNew And FormAction <> vbDataActionUpdate Then Exit Sub
'---- Navigate Row ,Col
If KeyAscii = 13 Then
    KeyAscii = 0
    With msGrid
        If .Col = 16 Then
            If .Row = (.Rows - 1) Then '-- Add New Row
                .Rows = .Rows + 1
                .Row = .Row + 1
                GridAddNew
            Else '--Next Row
                .Row = .Row + 1
            End If
            .Col = 0
        Else    '-- Next col
            If .Col = 1 Then
                .Col = 10
            Else
                .Col = .Col + 1
            End If
        End If
    End With
End If
'---- Set TxtGrid
If msGrid.text <> "" Then txtGrid.text = msGrid.text
If msGrid.Col <= rstblSub.Fields.Count - 1 Then
    txtGrid.Move msGrid.CellLeft + msGrid.Left, msGrid.CellTop + msGrid.Top, msGrid.CellWidth, msGrid.CellHeight
End If
txtGrid.Visible = True
txtGrid.Enabled = True
DoEvents
If KeyAscii <> 0 Then
    Select Case rstblSub.Fields(msGrid.Col).Type
        Case adChar, adVarWChar
            If msGrid.Col = 1 Then  '-- Count
                txtGrid_KeyPress (KeyAscii)
            Else
                txtGrid.text = txtGrid.text & Chr(KeyAscii)
            End If
        Case Else
            txtGrid.text = Chr(GProcValidateKey(rstblSub, msGrid.Col, KeyAscii, txtGrid.text))
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
    DoEvents
    txtGrid.text = ""
    txtGrid.Visible = False
End If
DoEvents
'If msGrid.Col = 1 Then
'    '-- Fill Grid Sub From Grid
'    FillGridSubFromGrid
'    Exit Sub
'End If
If msGrid.Col >= 1 Then
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
        If .Col <> 16 And IsLockCtrl = True Then
            KeyAscii = 0
        Else
            '--- Search
            Select Case .Col
                Case 1  '-- Count
                    MillCode = GProcGetColumnValue("TblMastAccount", "AcName", txttemp(5), "S", "Accode", "N")
                    gClsSearch.SearchMultiField "tblMastItem,tblMastaccount", "ItName,AcName,ItStdpack,ItStdrateper", Array("Count Name", "Mill", "Std. Pack", "Rate Per"), Array(txtGrid.Width, 2000, 1000, 1000), "tblMastItem.ITMillCode=tblMastaccount.acCode and ITMillCode=" & MillCode, Chr(KeyAscii), "ItName", LeftPos, TopPos, True
                    txtGrid = ""
                    KeyAscii = 0
                    If gClsSearch.SearchMultiRetCol(0) <> "" Then '-- Rate Per From Master
                        txtGrid = gClsSearch.SearchMultiRetCol(0)
                        If .TextMatrix(.Row, 11) = 0 Then .TextMatrix(.Row, 11) = gClsSearch.SearchMultiRetCol(2) '-Std .Pack
                        If .TextMatrix(.Row, 14) = 0 Then .TextMatrix(.Row, 14) = gClsSearch.SearchMultiRetCol(3) '-Rt /per
                        txtGrid_KeyPress 13
                    End If
    '                 msGrid.Col = 10
    '                msGrid.SetFocus
                Case 16 '- Is Complete Y/N
                    If UCase(Chr(KeyAscii)) = "Y" Then 'Y
                        txtGrid = ""
                        KeyAscii = Asc("Y")
                    Else
                        txtGrid = ""
                        KeyAscii = 0
                    End If
                Case Else
                    KeyAscii = GProcValidateKey(rstblSub, msGrid.Col, KeyAscii, txtGrid)
            End Select
        End If
   End If
    If KeyAscii = 13 Then CheckForNumberTxtGrid
    '--- Navigate Row ,Col
    If KeyAscii = 13 And .Col < 16 Then
        If .Col = 1 Then
            .Col = 10
        Else
            .Col = .Col + 1
        End If
    ElseIf KeyAscii = 13 And .Col = 16 Then
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
            If (Val(.TextMatrix(.Row, 15)) = 0) And .Rows > 2 Then '-- Amt=0
                .RemoveItem (.Row)
                txtGrid.text = ""
            End If
            txtGrid.Visible = False
           ' txt(18).SetFocus
            Exit Sub
        End If
    End With
End With
End Sub
Private Sub txtGrid_KeyDown(KeyCode As Integer, Shift As Integer)
Dim GName As String
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
    Case 1 '-- count
        GProcShowForm frmMastItem, frmMain.mnuMstfrm(3), Array(frmMastItem.txttemp(9)), Array(txttemp(5))
        Exit Sub
End Select
End If
End Sub
Private Sub CalGridAmt()
Dim MillCode As Long
Dim StdPack As Double
Dim Bags As Long
CheckForNumberGrid
MillCode = GProcGetColumnValue("TblMastAccount", "AcName", txttemp(5), "S", "Accode", "N")
StdPack = GProcGetColumnValue("TblMastItem", "ITName", msGrid.TextMatrix(msGrid.Row, 1), "S", "ItStdpack", "N", " ItMillCode=" & MillCode)
With msGrid
    
    If .TextMatrix(.Row, 12) = 0 Then
        .TextMatrix(.Row, 12) = .TextMatrix(.Row, 10) * StdPack
    End If
    '--Amt=Rate * Wt/Rate Per
    If CDbl(.TextMatrix(.Row, 14)) > 0 Then
        .TextMatrix(.Row, 15) = CDbl(.TextMatrix(.Row, 13)) * CDbl(.TextMatrix(.Row, 12)) / CDbl(.TextMatrix(.Row, 14))
    Else
        .TextMatrix(.Row, 15) = 0
    End If
    .TextMatrix(.Row, 15) = Format(.TextMatrix(.Row, 15), FStr)
    
    
    
    
'    If .Row = 0 Then Exit Sub
'    '--- Bags = Bags/LR * No of LR
'    Bags = (.TextMatrix(.Row, 2) * .TextMatrix(.Row, 3)) + (.TextMatrix(.Row, 4) * .TextMatrix(.Row, 5)) + (.TextMatrix(.Row, 6) * .TextMatrix(.Row, 7)) + (.TextMatrix(.Row, 8) * .TextMatrix(.Row, 9))
'    .TextMatrix(.Row, 10) = Format(Bags, PStr)
'    If (.TextMatrix(.Row, 11)) = 0 Then .TextMatrix(.Row, 11) = StdPack
'     '-- Wt=Bag * Std Pack
'   ' If .TextMatrix(.Row, 12) = 0 Then
'        .TextMatrix(.Row, 12) = CDbl(.TextMatrix(.Row, 10)) * CDbl(.TextMatrix(.Row, 11))
'   ' End If
'    '--Amt=Rate * Wt/Rate Per
'    If CDbl(.TextMatrix(.Row, 14)) > 0 Then
'        .TextMatrix(.Row, 15) = CDbl(.TextMatrix(.Row, 12)) * CDbl(.TextMatrix(.Row, 13)) / CDbl(.TextMatrix(.Row, 14))
'    Else
'        .TextMatrix(.Row, 15) = 0
'    End If
'    .TextMatrix(.Row, 15) = Format(.TextMatrix(.Row, 15), FStr)
''    '--Rate =Amt*Rate Per/Wt
''    If CDbl(.TextMatrix(.Row, 11)) > 0 Then
''        .TextMatrix(.Row, 12) = CDbl(.TextMatrix(.Row, 14)) * CDbl(.TextMatrix(.Row, 13)) / CDbl(.TextMatrix(.Row, 11))
''    Else
''        .TextMatrix(.Row, 12) = 0
''    End If
''    .TextMatrix(.Row, 12) = Format(.TextMatrix(.Row, 12), LStr)
End With
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
Dim MillCode  As Long
Dim RowIndex As Long
Dim ItName  As String
'-- Fill Grid From Grid Sub
If GridSubFlag = True Then
    FillGridFromGridSub
End If
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
                    If RowIndex > 1 And Trim(.TextMatrix(RowIndex, i)) = "" And CDbl(.TextMatrix(RowIndex, 3)) = 0 Then
                        .RemoveItem (RowIndex)
                        GoTo Validate
                    End If
                    '--Create Count
                    If Trim(.TextMatrix(RowIndex, i)) = "" Then
                        MsgBox "Check Count.", vbCritical + vbOKOnly, Me.Caption
                        .Col = i
                        msGrid.SetFocus
                        Exit Function
                    End If
                    MillCode = GProcGetColumnValue("TblMastAccount", "AcName", txttemp(5), "S", "Accode", "N")
                    If GProcGetColumnValue("TblMastItem", "ItName", .TextMatrix(RowIndex, i), "S", "Itcode", "N", "ITMillCode=" & MillCode) = 0 Then
                        MsgBox "Check Count.Mill and Count are Mismatch.", vbCritical + vbOKOnly, Me.Caption
                        .Col = i
                        msGrid.SetFocus
                        Exit Function
                    End If
                Case 10 '-- Bag
                    If CDbl(.TextMatrix(RowIndex, i)) = 0 Then
                        MsgBox "Check Bag.", vbCritical + vbOKOnly, Me.Caption
                        .Col = i
                        msGrid.SetFocus
                        Exit Function
                    End If
                Case 12 '-- Wt
                    If CDbl(.TextMatrix(RowIndex, i)) = 0 Then
                        MsgBox "Check Weight.", vbCritical + vbOKOnly, Me.Caption
                        .Col = i
                        msGrid.SetFocus
                        Exit Function
                    End If
                Case 13 '-- Rate
                    If CDbl(.TextMatrix(RowIndex, i)) = 0 Then
                        MsgBox "Check Rate.", vbCritical + vbOKOnly, Me.Caption
                        .Col = i
                        msGrid.SetFocus
                        Exit Function
                    End If
           End Select
           '--- Bag /Lr and No Of LR
           If (CLng(.TextMatrix(RowIndex, 2)) > 0 And CLng(.TextMatrix(RowIndex, 3)) = 0) Or (CLng(.TextMatrix(RowIndex, 2)) = 0 And CLng(.TextMatrix(RowIndex, 3)) > 0) Then
                MsgBox "Check Bag LR Details.", vbCritical + vbOKOnly, Me.Caption
                msGrid.SetFocus
                Exit Function
           End If
           If (CLng(.TextMatrix(RowIndex, 4)) > 0 And CLng(.TextMatrix(RowIndex, 5)) = 0) Or (CLng(.TextMatrix(RowIndex, 4)) = 0 And CLng(.TextMatrix(RowIndex, 5)) > 0) Then
                MsgBox "Check Bag LR Details.", vbCritical + vbOKOnly, Me.Caption
                msGrid.SetFocus
                Exit Function
           End If
           If (CLng(.TextMatrix(RowIndex, 6)) > 0 And CLng(.TextMatrix(RowIndex, 7)) = 0) Or (CLng(.TextMatrix(RowIndex, 6)) = 0 And CLng(.TextMatrix(RowIndex, 7)) > 0) Then
                MsgBox "Check Bag LR Details.", vbCritical + vbOKOnly, Me.Caption
                msGrid.SetFocus
                Exit Function
           End If
           If (CLng(.TextMatrix(RowIndex, 8)) > 0 And CLng(.TextMatrix(RowIndex, 9)) = 0) Or (CLng(.TextMatrix(RowIndex, 8)) = 0 And CLng(.TextMatrix(RowIndex, 9)) > 0) Then
                MsgBox "Check Bag LR Details.", vbCritical + vbOKOnly, Me.Caption
                msGrid.SetFocus
                Exit Function
           End If
        Next i
        '--- Duplicate Count
        ItName = .TextMatrix(RowIndex, 1)
        For i = 1 To .Rows - 1
            If i <> RowIndex And .TextMatrix(i, 1) = ItName Then
                MsgBox "Duplicate count.", vbCritical + vbOKOnly, Me.Caption
                .Col = 1
                msGrid.SetFocus
                Exit Function
            End If
        Next i
      Next RowIndex
End With
ValidateGrid = True
ReNumberGridSrno
End Function
Private Sub SaveBtnEd()
'--- Vno,vdt,Mill,Amount
If CDbl(txt(0)) = 0 Or GProcIsDateValid(mskDt(2)) = False Or txttemp(5) = "" Or CDbl(txt(11)) = 0 Then
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
Dim MillCode As Long
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
    '---- Broker
    If txttemp(3) = "" Then txttemp(3) = "-"
    '--Create Broker
    If GProcGetColumnValue("tblMastAccount", "AcName", txttemp(3), "S", "AcCode", "N") = 0 Then
        GName = GProcGetColumnValue("TblMastGroup", "AgCode", 90016, "N", "AgName", "S")
        GProcShowForm frmMastAccount, frmMain.mnuMstfrm(0), Array(frmMastAccount.txt(1), frmMastAccount.txttemp(2)), Array(txttemp(3), GName)
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
If IsDate(mskDt(2)) Then '--Vdt
    txt(2) = CDate(mskDt(2))
Else
    txt(2) = ""
End If

If IsDate(mskDt(0)) Then '--Vdt
    txt(7) = CDate(mskDt(2))
Else
    txt(7) = ""
End If


'--- Book No
txt(8) = txt(0)
'--- Book Dt
txt(9) = txt(2)
If txttemp(4) = "" Or IsNull(txttemp(4)) Then
   txttemp(4) = "-"
End If
If IsNull(txttemp(3)) Then
   txttemp(3) = "-"
End If
'-- Broker Code
txt(3) = GProcGetColumnValue("tblMastAccount", "AcName", txttemp(3), "S", "AcCode", "N")
'-- Dr Code
txt(4) = GProcGetColumnValue("tblMastAccount", "AcName", txttemp(4), "S", "AcCode", "N")
'--Mill
txt(5) = GProcGetColumnValue("tblMastAccount", "AcName", txttemp(5), "S", "AcCode", "N")
'-- TaxType
txt(12) = GProcGetColumnValue("tblMastNarration", "Narration", "-", "S", "NarrCode", "N")
txt(23) = gCYear
'-- Sub Amt
txt(10) = txt(11)
'-- Book Rt Type
txt(47) = cbo(47).ListIndex
txt(40) = Chk(0).Value
txt(41) = Chk(1).Value
txt(42) = ChkExGST.Value

'PtyGrp = GProcGetColumnValue("TblMastAccount", "AcName", txttemp(4), "C", "AcCmpCode", "N")   '-Party Group Code
'lblPartyGrp.Caption = GProcGetColumnValue("tblmastcompGroup", "CgCode", PtyGrp, "N", "Cgname", "S")
End Sub
Private Sub FillTempFromTxt()
Dim X As String
If IsNull(txt(40)) Then txt(40) = 0
Chk(0).Value = Val(txt(40))
If IsNull(txt(41)) Then txt(41) = 0
Chk(1).Value = Val(txt(41))
If IsNull(txt(42)) Then txt(42) = 0
ChkExGST.Value = Val(txt(42))

If IsDate(txt(2)) Then '--VDt
    mskDt(2) = txt(2)
Else
    mskDt(2) = "__/__/____"
End If
If IsDate(txt(7)) Then '--VDt
    mskDt(0) = txt(7)
Else
    mskDt(0) = "__/__/____"
End If
cbo(47).ListIndex = txt(47) '-- Rt Type
'--Broker
txttemp(3) = GProcGetColumnValue("tblMastAccount", "AcCode", txt(3), "N", "AcName", "S")
'--Party
txttemp(4) = GProcGetColumnValue("tblMastAccount", "AcCode", txt(4), "N", "AcName", "S")
'--Mill
txttemp(5) = GProcGetColumnValue("tblMastAccount", "AcCode", txt(5), "N", "AcName", "S")
'--Party Group
PtyGrpCd = GProcGetColumnValue("tblMastAccount", "AcCode", txt(4), "N", "AcCmpCode", "N")
txtPartyGrp = GProcGetColumnValue("tblMastCompGroup", "CGCode", "" & PtyGrpCd & "", "N", "CGName", "S")
'-- Fill Related Record
FillRelatedRecord
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
                If ColIndex = 10 Then TotBag = TotBag + CLng(msGrid.TextMatrix(RowIndex, ColIndex)) '-- Tot Bag
                If ColIndex = 12 Then TotWt = TotWt + CDbl(msGrid.TextMatrix(RowIndex, ColIndex)) '-- Tot Wt
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
End Sub
Private Sub ReadFields()
FormAction = vbDataActionClose
GProcShowRecord Me, rstbl, MaxNo
FillTempFromTxt
End Sub
Private Sub CalAmount()
Dim Amt As Double
Dim RowIndex As Integer
Dim TotBag As Long
Dim TotWt As Double
CalGridAmt
'---- Sub amt
GProcCheckForNumber Me, rstbl, MaxNo
CheckForNumberGrid
RowIndex = 1
With msGrid
    Do While RowIndex < .Rows
        '-- sub amt
        If Val(.TextMatrix(RowIndex, 15)) > 0 Then
            Amt = Amt + Val(.TextMatrix(RowIndex, 15))
        End If
        '--Tot Bag
        If Val(.TextMatrix(RowIndex, 10)) > 0 Then
            TotBag = TotBag + Val(.TextMatrix(RowIndex, 10))
        End If
         '--Tot Wt
        If Val(.TextMatrix(RowIndex, 12)) > 0 Then
            TotWt = TotWt + Val(.TextMatrix(RowIndex, 12))
        End If
        RowIndex = RowIndex + 1
    Loop
End With
txttmpTotBag = TotBag
txttmpTotWt = Format(TotWt, WtStr)
Amt = Format(Amt, FStr)
txt(11) = Amt
GProcCheckForNumber Me, rstbl, MaxNo
End Sub
'--- Is Related Record
Private Function IsRelatedRecord() As Boolean
Dim Rs1 As Recordset
Dim X As String
'-- Audited
If GProcISAudited(txt(0), txt(1), txt(23)) = True Then
    MsgBox "Record is Audited.You can not Modify/Delete this Record.", vbCritical + vbOKOnly, App.Title
    IsRelatedRecord = True
    Exit Function
End If
'-- Is SIT Purchase
X = "Select Vno,Vdt from tblVouSub where Vtype='GP' and VSubBOOKNO =" & txt(0) & " and VSubBOOKDT=cdate('" & mskDt(2) & "')"
GProcRstOpen Rs1, X, "R"
If Rs1.EOF Then
Else
    IsRelatedRecord = True
    MsgBox "You can not Modify/Delete this record.Related GP VNo. " & Rs1.Fields("Vno") & " is exists.", vbCritical + vbOKOnly, App.Title
    Exit Function
End If
Rs1.Close
End Function
'================= Sub Grid code
Private Sub msGridsub_EnterCell()
If FormAction <> vbDataActionCancel Then
    txtGridSub = ""
    If msGridSub.text <> "" Then txtGridSub.text = msGridSub.text
    '--- Lock TextBox
    With msGridSub
        Select Case .Col
            Case Else
                txtGridSub.Visible = True
                txtGridSub.Locked = False
        End Select
    End With
    msGridsub_KeyPress (0)
End If
End Sub
Private Sub msGridsub_KeyPress(KeyAscii As Integer)
With msGridSub
    If KeyAscii = 27 Then '----Esc Key
        KeyAscii = 0
'       txtGridSub.Visible = False
'       msGridSub.Visible = False
'-- Fill Grid From Grid Sub
        FillGridFromGridSub
        Exit Sub
    End If
End With
'---- Navigate Row ,Col
If KeyAscii = 13 Then
    KeyAscii = 0
    With msGridSub
        If .Col = 1 Then
            If .Row = (.Rows - 1) Then
            Else '--Next Row
                .Row = .Row + 1
            End If
            .Col = 0
        Else    '-- Next col
            .Col = .Col + 1
        End If
    End With
End If
'---- Set TxtGridSub
If msGridSub.text <> "" Then txtGridSub.text = msGridSub.text
If msGridSub.Col <= 1 Then
    txtGridSub.Move msGridSub.CellLeft + msGridSub.Left, msGridSub.CellTop + msGridSub.Top, msGridSub.CellWidth, msGridSub.CellHeight
End If
txtGridSub.Visible = True
txtGridSub.Enabled = True
If KeyAscii <> 0 Then
    Select Case msGridSub.Col
        Case Else
            txtGridSub.text = Chr(GProcNumberOnly(KeyAscii))
    End Select
End If
If msGridSub.text <> "" Then
    txtGridSub.SelStart = Len(txtGridSub.text) + 1
Else
    txtGridSub.SelStart = 1
End If
If txtGridSub.Visible = True Then txtGridSub.SetFocus
End Sub
Private Sub msGridsub_LeaveCell()
If txtGridSub.Visible Then
    msGridSub.text = txtGridSub.text
    DoEvents
    txtGridSub.text = ""
    txtGridSub.Visible = False
End If
DoEvents
If msGridSub.Col >= 0 Then
    CalGridAmt
    CalAmount
End If
DoEvents
End Sub
Private Sub txtGridSub_GotFocus()
    Call GProcSelectBox(txtGridSub)
End Sub
Private Sub txtGridsub_KeyPress(KeyAscii As Integer)
With msGridSub
    If KeyAscii <> 13 And KeyAscii <> 27 Then
        Select Case .Col
            Case Else
                KeyAscii = GProcNumberOnly(KeyAscii)
        End Select
   End If
   If KeyAscii = 13 Then
     If IsNumeric(txtGridSub) = False Then
        txtGridSub = 0
     Else
        txtGridSub = CLng(txtGridSub)
     End If
   End If
    '--- Navigate Row ,Col
    If KeyAscii = 13 And .Col < 1 Then
        .Col = .Col + 1
    ElseIf KeyAscii = 13 And .Col = 1 Then
        If .Row = (.Rows - 1) Then
        Else
            .Row = .Row + 1
        End If
        .Col = 0
    End If
    '-- On ESC
    With msGridSub
        If KeyAscii = 27 Then '----Esc Key
            KeyAscii = 0
            '--- Fill Grid From Grid Sub
            FillGridFromGridSub
            Exit Sub
        End If
    End With
End With
End Sub
'--- Fill Grid From GridSub
Private Sub FillGridFromGridSub()
GridSubFlag = False
'--- Bags /LR & LR From sub Grid
With msGridSub
    If txtGridSub.Visible = True Then .TextMatrix(.Row, .Col) = txtGridSub
    If IsNumeric(.TextMatrix(1, 0)) = True Then msGrid.TextMatrix(msGrid.Row, 2) = .TextMatrix(1, 0)
    If IsNumeric(.TextMatrix(2, 0)) = True Then msGrid.TextMatrix(msGrid.Row, 4) = .TextMatrix(2, 0)
    If IsNumeric(.TextMatrix(3, 0)) = True Then msGrid.TextMatrix(msGrid.Row, 6) = .TextMatrix(3, 0)
    If IsNumeric(.TextMatrix(4, 0)) = True Then msGrid.TextMatrix(msGrid.Row, 8) = .TextMatrix(4, 0)

    If IsNumeric(.TextMatrix(1, 1)) = True Then msGrid.TextMatrix(msGrid.Row, 3) = .TextMatrix(1, 1)
    If IsNumeric(.TextMatrix(2, 1)) = True Then msGrid.TextMatrix(msGrid.Row, 5) = .TextMatrix(2, 1)
    If IsNumeric(.TextMatrix(3, 1)) = True Then msGrid.TextMatrix(msGrid.Row, 7) = .TextMatrix(3, 1)
    If IsNumeric(.TextMatrix(4, 1)) = True Then msGrid.TextMatrix(msGrid.Row, 9) = .TextMatrix(4, 1)
    .Visible = False
    txtGridSub = ""
    txtGridSub.Visible = False
    CalGridAmt
    CalAmount
    msGrid.Col = 11
    msGrid.SetFocus
End With
End Sub
'--- Fill Grid Sub From Grid
Private Sub FillGridSubFromGrid()
GridSubFlag = True
'--- Bags /LR & LR From Grid
With msGrid
    msGridSub.Visible = True
    txtGridSub.Visible = True
'    msGridSub.SetFocus
    msGridSub.Col = 0
    msGridSub.Row = 1
    msGridSub.TextMatrix(1, 0) = .TextMatrix(.Row, 2)
    txtGridSub = msGridSub.TextMatrix(1, 0)
    txtGridSub.Move msGridSub.CellLeft + msGridSub.Left, msGridSub.CellTop + msGridSub.Top, msGridSub.CellWidth, msGridSub.CellHeight
    msGridSub.TextMatrix(2, 0) = .TextMatrix(.Row, 4)
    msGridSub.TextMatrix(3, 0) = .TextMatrix(.Row, 6)
    msGridSub.TextMatrix(4, 0) = .TextMatrix(.Row, 8)
    msGridSub.TextMatrix(1, 1) = .TextMatrix(.Row, 3)
    msGridSub.TextMatrix(2, 1) = .TextMatrix(.Row, 5)
    msGridSub.TextMatrix(3, 1) = .TextMatrix(.Row, 7)
    msGridSub.TextMatrix(4, 1) = .TextMatrix(.Row, 9)
    DoEvents
    msGridSub.SetFocus
    txtGridSub.SetFocus
End With
End Sub
'------  Print
Private Sub txtno_GotFocus(Index As Integer)
GProcSelectBox Me.ActiveControl
End Sub
Private Sub txtno_KeyPress(Index As Integer, KeyAscii As Integer)
KeyAscii = GProcNumberOnly(KeyAscii)
End Sub
Private Sub cmdPrint_Click(Index As Integer)
Dim X As String
Dim RsGp As Recordset
Dim RSAc As Recordset
Dim RSAc1 As Recordset
Dim GpNo As Long
Dim XX As String
Dim RsTmp As Recordset

Select Case Index
    Case 0, 1, 3, 4:
        '------ Validate
        If IsNumeric(txtno(0)) = False Then txtno(0) = 0
        If IsNumeric(txtno(1)) = False Then txtno(1) = txtno(0)
        If CLng(txtno(1)) < CLng(txtno(0)) Then txtno(1) = txtno(0)
        '-------
        If Index = 0 Then GProcCrystalRptPreparation CryReport, "Printer"
        If Index = 1 Then GProcCrystalRptPreparation CryReport, "Window"
        If Index = 3 Then GProcCrystalRptPreparation CryReport, "Window"
        If Index = 4 Then GProcCrystalRptPreparation CryReport, "Window"
        CryReport.WindowTitle = "Party Booking"
        GpNo = CLng(txtno(0))
        Do While GpNo <= txtno(1)
           XX = "Select * From tblAddLess_Book where  vNo=" & GpNo & " and tblAddLess_Book.VYear=" & gCYear & "  and tblAddLess_Book.Vtype='" & txt(1) & "'"
           If GProcRstOpen(RsGp, XX, "R", gCn) > 0 Then
               With CryReport
               X = "{tblAddLess_Book.Vno}>=" & txtno(0) & " and {tblAddLess_Book.Vno}<=" & txtno(1) & " and {tblAddLess_Book.VType}='" & frmTypeStr & "' and {tblAddLess_Book.VYear}=" & gCYear
               CryReport.SelectionFormula = X
               CryReport.Formulas(8) = "wCompGstin = '" & gCGSTIN & "'"
               If Check1.Value = 1 Then
                   CryReport.Formulas(15) = "wMnm = 'Y'"
               Else
                   CryReport.Formulas(15) = "wMnm = 'N'"
               End If
               CryReport.ReportFileName = gReportPath & "rptBookingPartyNew.rpt"
               If Index <> 3 And Index <> 4 Then
                 CryReport.Action = 1
               End If

'-------------------Mail
                 If (Index = 3 Or Index = 4) Then
                    Dim objCrystal As CRAXDRT.Application
                    Dim objReport As CRAXDRT.Report
                    Dim Tattach As String
                    Dim Mfile As String
                    Dim Mcp As String
                    Dim PtyName As String
                    Dim PtyMailId As String
                    Dim PtyMailIdCc As String
                    Dim PtyMailIdBcc As String
                    Dim MailSub As String
                    Dim MailStr As String
                    Set objCrystal = New CRAXDRT.Application
                    Tattach = gReportPath & "rptBookingPartyNew.rpt"
                    Set objReport = objCrystal.OpenReport(Tattach, 1)
                    
                    Dim crTab As CRAXDRT.DatabaseTable
                    For Each crTab In objReport.Database.Tables
                        crTab.Location = gDbLocation
                    Next
                    
                    
                    objReport.RecordSelectionFormula = "{tblAddLess_Book.Vno}>=" & txtno(0) & " and {tblAddLess_Book.Vno}<=" & txtno(1) & " and {tblAddLess_Book.VType}='" & frmTypeStr & "' and {tblAddLess_Book.VYear}=" & gCYear
                    
                    Dim j As Integer
                    For j = 1 To objReport.FormulaFields.Count
                        Select Case objReport.FormulaFields(j).Name
                        Case "{@FirmNm}"
                            objReport.FormulaFields(j).text = "Trim(" & Chr(39) & UCase(gCName) & Chr(39) & ")"
                        Case "{@wAdd1}"
                            objReport.FormulaFields(j).text = "Trim(" & Chr(39) & gCAdd1 & Chr(39) & ")"
                        Case "{@wAdd2}"
                            objReport.FormulaFields(j).text = "Trim(" & Chr(39) & gCAdd2 & Chr(39) & ")"
                        Case "{@wAdd3}"
                            objReport.FormulaFields(j).text = "Trim(" & Chr(39) & gCAdd3 & Chr(39) & ")"
                        Case "{@wPhNo}"
                            objReport.FormulaFields(j).text = "Trim(" & Chr(39) & gCPhNo & Chr(39) & ")"
                        Case "{@wCompGstin}"
                            objReport.FormulaFields(j).text = "Trim(" & Chr(39) & gCGSTIN & Chr(39) & ")"
                        Case "{@Wcopy}"
                            If Index = 3 And Index = 4 Then
                               objReport.FormulaFields(j).text = "Trim(" & Chr(39) & "WhatsApp Copy" & Chr(39) & ")"
                            Else
                               objReport.FormulaFields(j).text = "Trim(" & Chr(39) & "Mail Copy" & Chr(39) & ")"
                            End If
                        Case "{@BrkName}"
                            'objReport.FormulaFields(j).text = "Trim(" & Chr(39) & UCase(Bnm) & Chr(39) & ")"
                        End Select
                   
                    Next j
                    Dim PtyMbl As String
                    Dim BrkName As String
                    Dim BrkMbl As String
                    Dim BrkCd As Long
                      
                    Dim FileNMWithPath As String
                    Dim FileNM As String
                    Dim aFile As String
                    FileNM = "Booking" & Format(Now, "ddMMyyyyhhmmss") & "" & ".pdf"
                    FileNMWithPath = gReportPath & "" & FileNM & ""
                    PtyMbl = GProcGetColumnValue("TblMastAccount", "AcCode", RsGp!AdAcDrCode, "N", "AcFaxNo", "S")
                    PtyName = GProcGetColumnValue("TblMastAccount", "AcCode", RsGp!AdAcDrCode, "N", "AcName", "S")
                    If Index = 3 Then
                       ExportReportToPDF objReport, FileNMWithPath, "foo"
                       Gdelay (5)
                       SentWhatsApp "PDF", FileNM, "", PtyMbl, PtyName
                    End If

                    If Index = 4 Then
                      '-- To Broker
                          BrkCd = RsGp!AdAcCrCode
                          BrkName = GProcGetColumnValue("TblMastAccount", "AcCode", "" & BrkCd & "", "N", "AcName", "S")
                          BrkMbl = GProcGetColumnValue("TblMastAccount", "AcCode", "" & BrkCd & "", "N", "AcFaxNo", "S")
                          ExportReportToPDF objReport, FileNMWithPath, "foo"
                          Gdelay (5)
                          SentWhatsApp "PDF", FileNM, "", BrkMbl, BrkName
                     End If

                     gReportPath = App.Path & "\HIReports_Access\"
                     aFile = gReportPath & "*.pdf"
                     If Len(Dir$(aFile)) > 0 Then
                        Kill aFile
                     End If
                 End If
            End With
        End If
        GpNo = GpNo + 1
    Loop
    X = "Select * from tmpWhatsAppErr where username = '" & gUserName & "'"
    If GProcRstOpen(RsTmp, X, "R", gCn) > 0 Then
       GProcWhatsAppError
    End If
        
        
Case 2:
    DoEvents
    Frame1.Visible = False
End Select
End Sub
