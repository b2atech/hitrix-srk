VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form frmBookingSIT 
   BackColor       =   &H00C2E3FC&
   BorderStyle     =   0  'None
   Caption         =   "SIT Booking Entry"
   ClientHeight    =   8655
   ClientLeft      =   0
   ClientTop       =   15
   ClientWidth     =   10845
   DrawMode        =   10  'Mask Pen
   FillColor       =   &H000040C0&
   FillStyle       =   0  'Solid
   Icon            =   "frmBookingSIT.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   8655
   ScaleWidth      =   10845
   ShowInTaskbar   =   0   'False
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
      Top             =   1245
      Width           =   3945
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFC0C0&
      Caption         =   "SIT Book Printing"
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
      Left            =   2625
      TabIndex        =   92
      Top             =   2790
      Visible         =   0   'False
      Width           =   5175
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
         TabIndex        =   97
         Top             =   1560
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
         TabIndex        =   96
         Top             =   1560
         Width           =   1335
      End
      Begin VB.TextBox txtno 
         Appearance      =   0  'Flat
         Height          =   360
         Index           =   1
         Left            =   1950
         MaxLength       =   7
         TabIndex        =   95
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
         TabIndex        =   94
         Top             =   1560
         Width           =   1335
      End
      Begin VB.TextBox txtno 
         Appearance      =   0  'Flat
         Height          =   360
         Index           =   0
         Left            =   1950
         MaxLength       =   7
         TabIndex        =   93
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
         TabIndex        =   99
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
         TabIndex        =   98
         Top             =   600
         Width           =   1065
      End
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
      Left            =   3435
      TabIndex        =   91
      Top             =   2580
      Visible         =   0   'False
      Width           =   1035
   End
   Begin MSFlexGridLib.MSFlexGrid msGridSub 
      Height          =   2070
      Left            =   3045
      TabIndex        =   10
      Top             =   1890
      Visible         =   0   'False
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
      Left            =   0
      TabIndex        =   90
      Text            =   "47"
      Top             =   1200
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
      Index           =   46
      Left            =   2025
      TabIndex        =   16
      Text            =   "46"
      Top             =   6525
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
      Index           =   45
      Left            =   2025
      TabIndex        =   15
      Text            =   "45"
      Top             =   6165
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
      Height          =   360
      Index           =   44
      Left            =   2025
      TabIndex        =   18
      Text            =   "44"
      ToolTipText     =   "Press F2 for Selection"
      Top             =   7275
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
      Height          =   360
      Index           =   43
      Left            =   2025
      TabIndex        =   14
      Text            =   "43"
      ToolTipText     =   "Press F2 for Selection"
      Top             =   5775
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
      Left            =   60
      TabIndex        =   89
      Text            =   "42"
      Top             =   1335
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
      Left            =   105
      TabIndex        =   88
      Text            =   "41"
      Top             =   990
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
      Left            =   120
      TabIndex        =   87
      Text            =   "40"
      Top             =   1455
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
      Left            =   2025
      TabIndex        =   17
      Text            =   "39"
      Top             =   6930
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
      Index           =   38
      Left            =   45
      TabIndex        =   86
      Text            =   "38"
      Top             =   915
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
      Left            =   135
      TabIndex        =   85
      Text            =   "37"
      Top             =   765
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
      Left            =   60
      TabIndex        =   84
      Text            =   "17"
      Top             =   585
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
      TabIndex        =   20
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
      TabIndex        =   21
      Top             =   4650
      Width           =   915
   End
   Begin VB.TextBox txt 
      Alignment       =   1  'Right Justify
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
      Index           =   14
      Left            =   9990
      TabIndex        =   83
      Text            =   "14"
      Top             =   5985
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
      Left            =   10170
      TabIndex        =   82
      Text            =   "36"
      Top             =   5820
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
      Left            =   9630
      TabIndex        =   81
      Text            =   "35"
      Top             =   5940
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
      Left            =   9750
      TabIndex        =   80
      Text            =   "34"
      Top             =   5955
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
      Height          =   375
      Left            =   2010
      TabIndex        =   77
      Text            =   "txttmpTotBag"
      Top             =   5100
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
      Height          =   375
      Left            =   4770
      TabIndex        =   76
      Text            =   "txttmpTotWt"
      Top             =   5100
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
      TabIndex        =   32
      Top             =   3390
      Visible         =   0   'False
      Width           =   1035
   End
   Begin MSFlexGridLib.MSFlexGrid msGrid 
      Height          =   3135
      Left            =   195
      TabIndex        =   9
      Top             =   1800
      Width           =   9165
      _ExtentX        =   16166
      _ExtentY        =   5530
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
      ItemData        =   "frmBookingSIT.frx":058A
      Left            =   6840
      List            =   "frmBookingSIT.frx":0594
      TabIndex        =   4
      Text            =   "Cbo(47)"
      Top             =   1245
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
      Left            =   9780
      TabIndex        =   34
      Text            =   "18"
      Top             =   5850
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
      Left            =   9765
      TabIndex        =   36
      Text            =   "15"
      Top             =   5910
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
      Height          =   375
      Index           =   0
      Left            =   6840
      TabIndex        =   1
      Text            =   "0"
      Top             =   720
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
      Left            =   2370
      TabIndex        =   75
      Text            =   "33"
      Top             =   9165
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
      Left            =   2025
      TabIndex        =   19
      Text            =   "32"
      Top             =   7680
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
      Height          =   375
      Index           =   31
      Left            =   2280
      TabIndex        =   74
      Text            =   "31"
      Top             =   8760
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
      Left            =   1650
      TabIndex        =   45
      Text            =   "30"
      Top             =   9390
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
      Left            =   1305
      TabIndex        =   44
      Text            =   "29"
      Top             =   9315
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
      Left            =   1230
      TabIndex        =   43
      Text            =   "28"
      Top             =   8910
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
      Left            =   810
      TabIndex        =   42
      Text            =   "27"
      Top             =   9075
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
      Left            =   10560
      TabIndex        =   41
      Text            =   "26"
      Top             =   7440
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
      Left            =   10560
      TabIndex        =   40
      Text            =   "25"
      Top             =   5835
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
      Left            =   9585
      TabIndex        =   37
      Text            =   "24"
      Top             =   5820
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
      TabIndex        =   72
      Text            =   "1"
      Top             =   2430
      Width           =   420
   End
   Begin VB.TextBox txt 
      Alignment       =   1  'Right Justify
      Height          =   315
      Index           =   2
      Left            =   11040
      TabIndex        =   61
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
      TabIndex        =   60
      Text            =   "5"
      Top             =   2640
      Width           =   525
   End
   Begin VB.TextBox txt 
      Alignment       =   1  'Right Justify
      Height          =   315
      Index           =   3
      Left            =   11040
      TabIndex        =   59
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
      Left            =   9480
      Style           =   1  'Graphical
      TabIndex        =   51
      Top             =   7080
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
      Left            =   9600
      Style           =   1  'Graphical
      TabIndex        =   50
      ToolTipText     =   "To Add New Record"
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
      TabIndex        =   30
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
      TabIndex        =   49
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
      TabIndex        =   24
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
      TabIndex        =   23
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
      TabIndex        =   22
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
      Left            =   9690
      Style           =   1  'Graphical
      TabIndex        =   48
      Top             =   6105
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
      TabIndex        =   25
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
      TabIndex        =   29
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
      TabIndex        =   28
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
      TabIndex        =   27
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
      TabIndex        =   26
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
      TabIndex        =   6
      Text            =   "txttemp(5)"
      ToolTipText     =   "Press F3 for Creation"
      Top             =   720
      Width           =   3945
   End
   Begin VB.TextBox txt 
      Alignment       =   1  'Right Justify
      Height          =   315
      Index           =   4
      Left            =   11040
      TabIndex        =   58
      Text            =   "4"
      Top             =   2160
      Width           =   525
   End
   Begin VB.TextBox txt 
      Height          =   315
      Index           =   6
      Left            =   11040
      TabIndex        =   57
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
      TabIndex        =   56
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
      Left            =   9945
      TabIndex        =   31
      Text            =   "8"
      Top             =   5685
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.TextBox txt 
      Alignment       =   1  'Right Justify
      Height          =   315
      Index           =   9
      Left            =   11040
      TabIndex        =   55
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
      Left            =   9405
      TabIndex        =   33
      Text            =   "10"
      Top             =   5970
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
      Height          =   375
      Index           =   11
      Left            =   7620
      TabIndex        =   12
      Text            =   "11"
      Top             =   5100
      Width           =   1590
   End
   Begin VB.TextBox txt 
      Alignment       =   1  'Right Justify
      Height          =   315
      Index           =   12
      Left            =   11040
      TabIndex        =   54
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
      Left            =   9705
      TabIndex        =   38
      Text            =   "13"
      Top             =   5880
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
      Left            =   10110
      TabIndex        =   47
      Text            =   "16"
      Top             =   5970
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
      Index           =   19
      Left            =   9390
      TabIndex        =   35
      Text            =   "19"
      Top             =   5850
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
      Index           =   20
      Left            =   9630
      TabIndex        =   46
      Text            =   "20"
      Top             =   5835
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
      Index           =   21
      Left            =   1905
      TabIndex        =   39
      Text            =   "21"
      Top             =   8850
      Width           =   330
   End
   Begin VB.TextBox txt 
      Alignment       =   1  'Right Justify
      Height          =   315
      Index           =   22
      Left            =   11040
      TabIndex        =   53
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
      TabIndex        =   52
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
      Top             =   720
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
      Bindings        =   "frmBookingSIT.frx":05B1
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
   Begin VB.Label Label12 
      BackStyle       =   0  'Transparent
      Caption         =   "Remark             :  "
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
      Height          =   270
      Left            =   270
      TabIndex        =   104
      Top             =   7725
      Width           =   1665
   End
   Begin VB.Label Label11 
      BackStyle       =   0  'Transparent
      Caption         =   "Transport            :"
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
      Height          =   270
      Left            =   270
      TabIndex        =   103
      Top             =   6990
      Width           =   1665
   End
   Begin VB.Label Label10 
      BackStyle       =   0  'Transparent
      Caption         =   "Commssion        :"
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
      Height          =   270
      Left            =   270
      TabIndex        =   102
      Top             =   6585
      Width           =   1665
   End
   Begin VB.Label Label9 
      BackStyle       =   0  'Transparent
      Caption         =   "Payment Terms   :"
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
      Height          =   270
      Left            =   270
      TabIndex        =   101
      Top             =   6225
      Width           =   1665
   End
   Begin VB.Label Label8 
      BackStyle       =   0  'Transparent
      Caption         =   "Insurance           :  "
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
      Height          =   270
      Left            =   270
      TabIndex        =   100
      Top             =   7320
      Width           =   1665
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "Delivery Schedule :"
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
      Height          =   270
      Left            =   270
      TabIndex        =   13
      Top             =   5820
      Width           =   1665
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
      TabIndex        =   7
      Top             =   1245
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
      TabIndex        =   3
      Top             =   1245
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
      Height          =   375
      Left            =   3420
      TabIndex        =   79
      Top             =   5100
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
      Height          =   375
      Left            =   915
      TabIndex        =   78
      Top             =   5100
      Width           =   1440
   End
   Begin VB.Shape Shape8 
      BorderColor     =   &H00C00000&
      BorderWidth     =   5
      Height          =   8655
      Left            =   15
      Top             =   -30
      Width           =   10845
   End
   Begin VB.Shape Shape3 
      BorderColor     =   &H00C00000&
      FillColor       =   &H00C00000&
      FillStyle       =   0  'Solid
      Height          =   45
      Left            =   -15
      Top             =   8520
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
      Top             =   720
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
      Height          =   375
      Left            =   6375
      TabIndex        =   11
      Top             =   5100
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
      TabIndex        =   73
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
      TabIndex        =   71
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
      TabIndex        =   70
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
      TabIndex        =   69
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
      TabIndex        =   68
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
      TabIndex        =   67
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
      TabIndex        =   66
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
      TabIndex        =   65
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
      TabIndex        =   64
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
      TabIndex        =   63
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
      TabIndex        =   5
      Top             =   720
      Width           =   1140
   End
   Begin VB.Label Label6 
      BackColor       =   &H00C00000&
      Caption         =   "    SIT Booking"
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
      TabIndex        =   62
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
      Height          =   2625
      Left            =   240
      Top             =   5640
      Width           =   9135
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H00000080&
      FillColor       =   &H00EDFDFE&
      FillStyle       =   0  'Solid
      Height          =   1230
      Left            =   210
      Top             =   510
      Width           =   5145
   End
   Begin VB.Shape Shape6 
      BorderColor     =   &H00000080&
      FillColor       =   &H00EDFDFE&
      FillStyle       =   0  'Solid
      Height          =   1230
      Left            =   5415
      Top             =   510
      Width           =   3930
   End
End
Attribute VB_Name = "frmBookingSIT"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Const tblName As String = "tblAddLess"
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
Dim i As Long

Private Sub Cbo_KeyPress(Index As Integer, KeyAscii As Integer)
If Index = 47 Then KeyAscii = 0 '--Rate Type
End Sub
Private Sub cmdBtn_Click(Index As Integer)
On Error GoTo ErrorRoutine
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
    gClsSearch.SearchMultiField "tblAddLess,TblMastAccount", "Vno,AcName,Vdt,AdBillAmt", Array("Vno", "Mill", "Vdt", "Amount"), Array(txt(0).Width, 4000, 1080, 1500), " tblAddLess.AdMillCode=TblMastAccount.AcCode and vtype='" & frmTypeStr & "' and VYear=" & gCYear, "", "Vno,Vdt", 0, 0
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
    txt(0) = GProcGenerateId("TblAddLess", "Vno", "Vtype='" & frmTypeStr & "' and VYear=" & gCYear)
    mskDt(2) = GProcGenerateVDt("tblAddLess", "Vdt", "Vtype='" & frmTypeStr & "' and VYear=" & gCYear)
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
Case 10 'Exit
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
'        msGrid.Col = 0
'        msGrid.Row = 1
'        txtGrid.Move msGrid.CellLeft + msGrid.Left, msGrid.CellTop + msGrid.Top, msGrid.CellWidth, msGrid.CellHeight
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
    If msGrid.rows <= 1 Then msGrid.rows = msGrid.rows + 1
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
gCn.Execute "delete from tblVouSub where vtype='" & frmTypeStr & "' and vno=" & txt(0) & " and " _
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
X = "select * from tblVouSub where vtype='" & frmTypeStr & "' and vno=" & txt(0) & " and " _
& " VYear=" & gCYear
GProcRstOpen rsVouSub, X, "O"
With msGrid
    i = 1
    Do While i < .rows
        rsVouSub.AddNew
        rsVouSub!Vno = CLng(txt(0)) '--Vno(0)
        rsVouSub!Vtype = txt(1) '--Vtype(1)
        rsVouSub!Vdt = CDate(txt(2)) '--Vdt(2)
        rsVouSub!VSubCtrNo = i '--Sr No(3)
        '--- Book Ac code
        rsVouSub!VSUBBOOKACCODE = GProcGetColumnValue("tblMastAccount", "AcName", "-", "S", "AcCode", "N")
        '--Itcode(4)
        rsVouSub!vsubITCODE = GProcGetColumnValue("tblMastItem", "ItName", .TextMatrix(i, 1), "S", "ItCode", "N", " ItMillCode=" & MillCode)  '--Item Code
        rsVouSub!vsubBAG = CLng(.TextMatrix(i, 2)) '--Bag(5)
        rsVouSub!VSubItStdpack = CDbl(.TextMatrix(i, 3)) '--Std Packing
        rsVouSub!vsubwt = CDbl(.TextMatrix(i, 4)) '-- Wt(6)
        rsVouSub!vsubRt = CDbl(.TextMatrix(i, 5)) '--Rate(7)
        rsVouSub!vsubRTPER = CDbl(.TextMatrix(i, 6)) '--Rate Per(8)
        rsVouSub!vsubAMT = CDbl(.TextMatrix(i, 7)) '--Amt(9)
        '--Godown Code(17)
        rsVouSub!VSubGodownCode = GProcGetColumnValue("tblMastNarration", "Narration", "-", "S", "NarrCode", "N", "")
        rsVouSub!VSUBBOOKISCOMPLETED = IIf(Trim(UCase(.TextMatrix(i, 8))) = "Y", 1, 0) '-Is complete
        '--- It Ctrl No(11)
        If CLng(.TextMatrix(i, 9)) = 0 Then
            rsVouSub!VSubItCtrlNo = GProcGenerateId("tblVouSub", "VSubItCtrlNo", "VYear=" & gCYear)   '--It Ctrl No
        Else
            rsVouSub!VSubItCtrlNo = CLng(.TextMatrix(i, 9))
        End If
        rsVouSub!VSubIssBag = CLng(.TextMatrix(i, 10)) '--Issue Bag(13)
        rsVouSub!VSubIsswt = CDbl(.TextMatrix(i, 11)) '-- Issue Wt(14)
'        rsVouSub!VSUBBAGPERLR1 = CLng(.TextMatrix(i, 2)) '--Bag per Lr
'        rsVouSub!VSUBBAGPERLR2 = CLng(.TextMatrix(i, 4)) '--Bag per Lr
'        rsVouSub!VSUBBAGPERLR3 = CLng(.TextMatrix(i, 6)) '--Bag per Lr
'        rsVouSub!VSUBBAGPERLR4 = CLng(.TextMatrix(i, 8)) '--Bag per Lr
'        rsVouSub!VSUBNOOFLR1 = CLng(.TextMatrix(i, 3)) '--NO Of LR
'        rsVouSub!VSUBNOOFLR2 = CLng(.TextMatrix(i, 5)) '--NO Of LR
'        rsVouSub!VSUBNOOFLR3 = CLng(.TextMatrix(i, 7)) '--NO Of LR
'        rsVouSub!VSUBNOOFLR4 = CLng(.TextMatrix(i, 9)) '--NO Of LR
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
            If Me.ActiveControl.Name <> "txtGrid" Then SendKeys "{TAB}"
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
KeyAscii = GProcValidateKey(rstbl, Index, KeyAscii, txt(Index).Text)
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
    Case 5  '-- Mill  Alias Account
        gClsSearch.SearchMultiField "QryMastaccount", "Account,AcName,AcCity,AgName", Array("Mill", "AcName", "City", "Group"), Array(3000, 0, 2000, 3000), "((AgCode=90029 or GpCode=90029) and (AcMillType=4 or AcMillType=5))", Chr(KeyAscii), "AcName", LeftPos, TopPos, True
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
X = "select VSubCtrNo,ItName" _
& ",VSubBag,VSubItStdpack,VSubWt" _
& " ,VSubRt,VSubRtPer,VSubAmt," & IIf(gBackEndDB = gBackEndAccess, "iif(VSUBBOOKISCOMPLETED=1,'Y','')", "decode(VSUBBOOKISCOMPLETED,1,'Y','')") & ",VSUBITCTRLNO,VSUBISSBAG,VSUBISSWt from tblvouSub,TblMastItem where " _
& " VSUBItCODE=ItCode and tblvouSub.vtype='" & frmTypeStr & "' and tblvouSub.vno=" & txt(0) & " and tblvouSub.VYear=" & gCYear & " order by VSubCtrNo"

GProcRstOpen rstblSub, X, "R"
With msGrid
    .FormatString = ">SR|<Count                                   |>Bag |>Std Packing|>Weight      |>Rate            |>Rt Per  |>Amount       |<Is Complete(Y)|||"
    .ColWidth(9) = 0 '-It Ctrl
    .ColWidth(10) = 0 '-Issue Bag
    .ColWidth(11) = 0 '-Issue Wt
    If FormAction = vbDataActionAddNew Then
        .rows = 2
        GridAddNew
    Else
        .rows = 1
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
Dim i As Long
    msGrid_EnterCell
End Sub
Private Sub msGrid_GotFocus()
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
    If msGrid.Text <> "" Then txtGrid.Text = msGrid.Text
    '--- Lock TextBox
'    If IsLockCtrl = False Then
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
'    Else
'        If msGrid.Col = 16 Then '-Is complete
'            txtGrid.Locked = False
'        Else
'            txtGrid.Locked = True
'        End If
'    End If
    DoEvents
    msGrid_KeyPress (0)
    DoEvents
End If
End Sub
Private Sub msGrid_KeyPress(KeyAscii As Integer)
With msGrid
    If KeyAscii = 27 Then '----Esc Key
        KeyAscii = 0
        If (Val(.TextMatrix(.Row, 7)) = 0) And .rows > 2 Then '-- amt =0
            .RemoveItem (.Row)
            txtGrid.Text = ""
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
        If .Col = 8 Then
            If .Row = (.rows - 1) Then '-- Add New Row
                .rows = .rows + 1
                .Row = .Row + 1
                GridAddNew
            Else '--Next Row
                .Row = .Row + 1
            End If
            .Col = 0
        Else    '-- Next col
'            If .Col = 1 Then
'                .Col = 10
'            Else
                .Col = .Col + 1
'            End If
        End If
    End With
End If
'---- Set TxtGrid
If msGrid.Text <> "" Then txtGrid.Text = msGrid.Text
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
                txtGrid.Text = txtGrid.Text & Chr(KeyAscii)
            End If
        Case Else
            txtGrid.Text = Chr(GProcValidateKey(rstblSub, msGrid.Col, KeyAscii, txtGrid.Text))
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
                        If .TextMatrix(.Row, 3) = 0 Then .TextMatrix(.Row, 3) = gClsSearch.SearchMultiRetCol(2) '-Std .Pack
                        If .TextMatrix(.Row, 6) = 0 Then .TextMatrix(.Row, 6) = gClsSearch.SearchMultiRetCol(3) '-Rt /per
                        txtGrid_KeyPress 13
                    End If
    '                 msGrid.Col = 10
    '                msGrid.SetFocus
                Case 8 '- Is Complete Y/N
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
    If KeyAscii = 13 And .Col < 8 Then
'        If .Col = 1 Then
'            .Col = 10
'        Else
            .Col = .Col + 1
'        End If
    ElseIf KeyAscii = 13 And .Col = 8 Then
        If .Row = (.rows - 1) Then
            .rows = .rows + 1
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
            If (Val(.TextMatrix(.Row, 7)) = 0) And .rows > 2 Then '-- Amt=0
                .RemoveItem (.Row)
                txtGrid.Text = ""
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
If KeyCode = vbKeyY And Shift = 2 And msGrid.rows > 2 Then
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
    If .Row = 0 Then Exit Sub
    '--- Bags = Bags/LR * No of LR
'    Bags = (.TextMatrix(.Row, 2) * .TextMatrix(.Row, 3)) + (.TextMatrix(.Row, 4) * .TextMatrix(.Row, 5)) + (.TextMatrix(.Row, 6) * .TextMatrix(.Row, 7)) + (.TextMatrix(.Row, 8) * .TextMatrix(.Row, 9))
'    .TextMatrix(.Row, 10) = Format(Bags, PStr)
    If (.TextMatrix(.Row, 3)) = 0 Then .TextMatrix(.Row, 3) = StdPack
     '-- Wt=Bag * Std Pack
   ' If .TextMatrix(.Row, 12) = 0 Then
        .TextMatrix(.Row, 4) = CDbl(.TextMatrix(.Row, 2)) * CDbl(.TextMatrix(.Row, 3))
   ' End If
    '--Amt=Rate * Wt/Rate Per
    If CDbl(.TextMatrix(.Row, 6)) > 0 Then
        .TextMatrix(.Row, 7) = CDbl(.TextMatrix(.Row, 4)) * CDbl(.TextMatrix(.Row, 5)) / CDbl(.TextMatrix(.Row, 6))
    Else
        .TextMatrix(.Row, 7) = 0
    End If
    .TextMatrix(.Row, 7) = Format(.TextMatrix(.Row, 7), FStr)
End With
End Sub
Private Sub CheckForNumberGrid()
Dim RowIndex As Long
With rstblSub
For RowIndex = 1 To msGrid.rows - 1
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
    For i = 1 To .rows - 1
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
'If GridSubFlag = True Then
'    FillGridFromGridSub
'End If
If Trim(txtGrid.Text) <> "" Then
    msGrid.Text = txtGrid.Text
    CalGridAmt
End If
txtGrid.Visible = False
txtGrid.Text = ""
CheckForNumberGrid
With msGrid
Validate:
    For RowIndex = 1 To .rows - 1
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
                Case 2 '-- Bag
                    If CDbl(.TextMatrix(RowIndex, i)) = 0 Then
                        MsgBox "Check Bag.", vbCritical + vbOKOnly, Me.Caption
                        .Col = i
                        msGrid.SetFocus
                        Exit Function
                    End If
                Case 4 '-- Wt
                    If CDbl(.TextMatrix(RowIndex, i)) = 0 Then
                        MsgBox "Check Weight.", vbCritical + vbOKOnly, Me.Caption
                        .Col = i
                        msGrid.SetFocus
                        Exit Function
                    End If
                Case 5 '-- Rate
                    If CDbl(.TextMatrix(RowIndex, i)) = 0 Then
                        MsgBox "Check Rate.", vbCritical + vbOKOnly, Me.Caption
                        .Col = i
                        msGrid.SetFocus
                        Exit Function
                    End If
           End Select
           '--- Bag /Lr and No Of LR
'           If (CLng(.TextMatrix(RowIndex, 2)) > 0 And CLng(.TextMatrix(RowIndex, 3)) = 0) Or (CLng(.TextMatrix(RowIndex, 2)) = 0 And CLng(.TextMatrix(RowIndex, 3)) > 0) Then
'                MsgBox "Check Bag LR Details.", vbCritical + vbOKOnly, Me.Caption
'                msGrid.SetFocus
'                Exit Function
'           End If
'           If (CLng(.TextMatrix(RowIndex, 4)) > 0 And CLng(.TextMatrix(RowIndex, 5)) = 0) Or (CLng(.TextMatrix(RowIndex, 4)) = 0 And CLng(.TextMatrix(RowIndex, 5)) > 0) Then
'                MsgBox "Check Bag LR Details.", vbCritical + vbOKOnly, Me.Caption
'                msGrid.SetFocus
'                Exit Function
'           End If
'           If (CLng(.TextMatrix(RowIndex, 6)) > 0 And CLng(.TextMatrix(RowIndex, 7)) = 0) Or (CLng(.TextMatrix(RowIndex, 6)) = 0 And CLng(.TextMatrix(RowIndex, 7)) > 0) Then
'                MsgBox "Check Bag LR Details.", vbCritical + vbOKOnly, Me.Caption
'                msGrid.SetFocus
'                Exit Function
'           End If
'           If (CLng(.TextMatrix(RowIndex, 8)) > 0 And CLng(.TextMatrix(RowIndex, 9)) = 0) Or (CLng(.TextMatrix(RowIndex, 8)) = 0 And CLng(.TextMatrix(RowIndex, 9)) > 0) Then
'                MsgBox "Check Bag LR Details.", vbCritical + vbOKOnly, Me.Caption
'                msGrid.SetFocus
'                Exit Function
'           End If
        Next i
        '--- Duplicate Count
        ItName = .TextMatrix(RowIndex, 1)
        For i = 1 To .rows - 1
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
'--- Book No
txt(8) = txt(0)
'--- Book Dt
txt(9) = txt(2)
'-- Broker Code
txt(3) = GProcGetColumnValue("tblMastAccount", "AcName", txttemp(3), "S", "AcCode", "N")
'-- Dr Code
txt(4) = GProcGetColumnValue("tblMastAccount", "AcName", "-", "S", "AcCode", "N")
'--Mill
txt(5) = GProcGetColumnValue("tblMastAccount", "AcName", txttemp(5), "S", "AcCode", "N")
'-- TaxType
txt(12) = GProcGetColumnValue("tblMastNarration", "Narration", "-", "S", "NarrCode", "N")
txt(23) = gCYear
'-- Sub Amt
txt(10) = txt(11)
'-- Book Rt Type
txt(47) = Cbo(47).ListIndex
End Sub
Private Sub FillTempFromTxt()
Dim X As String
If IsDate(txt(2)) Then '--VDt
    mskDt(2) = txt(2)
Else
    mskDt(2) = "__/__/____"
End If
Cbo(47).ListIndex = txt(47) '-- Rt Type
'--Broker
txttemp(3) = GProcGetColumnValue("tblMastAccount", "AcCode", txt(3), "N", "AcName", "S")
'--Mill
txttemp(5) = GProcGetColumnValue("tblMastAccount", "AcCode", txt(5), "N", "AcName", "S")
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
        msGrid.rows = 2
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
            If Not .EOF Then msGrid.rows = msGrid.rows + 1
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
    Do While RowIndex < .rows
        '-- sub amt
        If Val(.TextMatrix(RowIndex, 7)) > 0 Then
            Amt = Amt + Val(.TextMatrix(RowIndex, 7))
        End If
        '--Tot Bag
        If Val(.TextMatrix(RowIndex, 2)) > 0 Then
            TotBag = TotBag + Val(.TextMatrix(RowIndex, 2))
        End If
         '--Tot Wt
        If Val(.TextMatrix(RowIndex, 4)) > 0 Then
            TotWt = TotWt + Val(.TextMatrix(RowIndex, 4))
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
If gBackEndDB = gBackEndAccess Then
    X = "Select Vno,Vdt from tblVouSub where Vtype='PT' and VSubBOOKNO =" & txt(0) & " and VSubBOOKDT=cdate('" & mskDt(2) & "')"
Else
    X = "Select Vno,Vdt from tblVouSub where Vtype='PT' and VSubBOOKNO =" & txt(0) & " and VSubBOOKDT=to_date('" & mskDt(2) & "','dd/MM/yyyy')"
End If
GProcRstOpen Rs1, X, "R"
If Rs1.EOF Then
Else
    IsRelatedRecord = True
    MsgBox "You can not Modify/Delete this record.Related SIT Purchase VNo. " & Rs1.Fields("Vno") & " is exists.", vbCritical + vbOKOnly, App.Title
    Exit Function
End If
Rs1.Close
End Function
'================= Sub Grid code
Private Sub msGridsub_EnterCell()
'If FormAction <> vbDataActionCancel Then
'    txtGridSub = ""
'    If msGridSub.Text <> "" Then txtGridSub.Text = msGridSub.Text
'    '--- Lock TextBox
'    With msGridSub
'        Select Case .Col
'            Case Else
'                txtGridSub.Visible = True
'                txtGridSub.Locked = False
'        End Select
'    End With
'    msGridsub_KeyPress (0)
'End If
End Sub
Private Sub msGridsub_KeyPress(KeyAscii As Integer)
'With msGridSub
'    If KeyAscii = 27 Then '----Esc Key
'        KeyAscii = 0
''       txtGridSub.Visible = False
''       msGridSub.Visible = False
''-- Fill Grid From Grid Sub
'        FillGridFromGridSub
'        Exit Sub
'    End If
'End With
''---- Navigate Row ,Col
'If KeyAscii = 13 Then
'    KeyAscii = 0
'    With msGridSub
'        If .Col = 1 Then
'            If .Row = (.rows - 1) Then
'            Else '--Next Row
'                .Row = .Row + 1
'            End If
'            .Col = 0
'        Else    '-- Next col
'            .Col = .Col + 1
'        End If
'    End With
'End If
''---- Set TxtGridSub
'If msGridSub.Text <> "" Then txtGridSub.Text = msGridSub.Text
'If msGridSub.Col <= 1 Then
'    txtGridSub.Move msGridSub.CellLeft + msGridSub.Left, msGridSub.CellTop + msGridSub.Top, msGridSub.CellWidth, msGridSub.CellHeight
'End If
'txtGridSub.Visible = True
'txtGridSub.Enabled = True
'If KeyAscii <> 0 Then
'    Select Case msGridSub.Col
'        Case Else
'            txtGridSub.Text = Chr(GProcNumberOnly(KeyAscii))
'    End Select
'End If
'If msGridSub.Text <> "" Then
'    txtGridSub.SelStart = Len(txtGridSub.Text) + 1
'Else
'    txtGridSub.SelStart = 1
'End If
'If txtGridSub.Visible = True Then txtGridSub.SetFocus
End Sub
Private Sub msGridsub_LeaveCell()
'If txtGridSub.Visible Then
'    msGridSub.Text = txtGridSub.Text
'    DoEvents
'    txtGridSub.Text = ""
'    txtGridSub.Visible = False
'End If
'DoEvents
'If msGridSub.Col >= 0 Then
'    CalGridAmt
'    CalAmount
'End If
'DoEvents
End Sub
Private Sub txtGridSub_GotFocus()
'    Call GProcSelectBox(txtGridSub)
End Sub
Private Sub txtGridsub_KeyPress(KeyAscii As Integer)
'With msGridSub
'    If KeyAscii <> 13 And KeyAscii <> 27 Then
'        Select Case .Col
'            Case Else
'                KeyAscii = GProcNumberOnly(KeyAscii)
'        End Select
'   End If
'   If KeyAscii = 13 Then
'     If IsNumeric(txtGridSub) = False Then
'        txtGridSub = 0
'     Else
'        txtGridSub = CLng(txtGridSub)
'     End If
'   End If
'    '--- Navigate Row ,Col
'    If KeyAscii = 13 And .Col < 1 Then
'        .Col = .Col + 1
'    ElseIf KeyAscii = 13 And .Col = 1 Then
'        If .Row = (.rows - 1) Then
'        Else
'            .Row = .Row + 1
'        End If
'        .Col = 0
'    End If
'    '-- On ESC
'    With msGridSub
'        If KeyAscii = 27 Then '----Esc Key
'            KeyAscii = 0
'            '--- Fill Grid From Grid Sub
'            FillGridFromGridSub
'            Exit Sub
'        End If
'    End With
'End With
End Sub
'--- Fill Grid From GridSub
Private Sub FillGridFromGridSub()
'GridSubFlag = False
''--- Bags /LR & LR From sub Grid
'With msGridSub
'    If txtGridSub.Visible = True Then .TextMatrix(.Row, .Col) = txtGridSub
'    If IsNumeric(.TextMatrix(1, 0)) = True Then msGrid.TextMatrix(msGrid.Row, 2) = .TextMatrix(1, 0)
'    If IsNumeric(.TextMatrix(2, 0)) = True Then msGrid.TextMatrix(msGrid.Row, 4) = .TextMatrix(2, 0)
'    If IsNumeric(.TextMatrix(3, 0)) = True Then msGrid.TextMatrix(msGrid.Row, 6) = .TextMatrix(3, 0)
'    If IsNumeric(.TextMatrix(4, 0)) = True Then msGrid.TextMatrix(msGrid.Row, 8) = .TextMatrix(4, 0)
'
'    If IsNumeric(.TextMatrix(1, 1)) = True Then msGrid.TextMatrix(msGrid.Row, 3) = .TextMatrix(1, 1)
'    If IsNumeric(.TextMatrix(2, 1)) = True Then msGrid.TextMatrix(msGrid.Row, 5) = .TextMatrix(2, 1)
'    If IsNumeric(.TextMatrix(3, 1)) = True Then msGrid.TextMatrix(msGrid.Row, 7) = .TextMatrix(3, 1)
'    If IsNumeric(.TextMatrix(4, 1)) = True Then msGrid.TextMatrix(msGrid.Row, 9) = .TextMatrix(4, 1)
'    .Visible = False
'    txtGridSub = ""
'    txtGridSub.Visible = False
'    CalGridAmt
'    CalAmount
'    msGrid.Col = 11
'    msGrid.SetFocus
'End With
End Sub
'--- Fill Grid Sub From Grid
Private Sub FillGridSubFromGrid()
'GridSubFlag = True
''--- Bags /LR & LR From Grid
'With msGrid
'    msGridSub.Visible = True
'    txtGridSub.Visible = True
''    msGridSub.SetFocus
'    msGridSub.Col = 0
'    msGridSub.Row = 1
'    msGridSub.TextMatrix(1, 0) = .TextMatrix(.Row, 2)
'    txtGridSub = msGridSub.TextMatrix(1, 0)
'    txtGridSub.Move msGridSub.CellLeft + msGridSub.Left, msGridSub.CellTop + msGridSub.Top, msGridSub.CellWidth, msGridSub.CellHeight
'    msGridSub.TextMatrix(2, 0) = .TextMatrix(.Row, 4)
'    msGridSub.TextMatrix(3, 0) = .TextMatrix(.Row, 6)
'    msGridSub.TextMatrix(4, 0) = .TextMatrix(.Row, 8)
'    msGridSub.TextMatrix(1, 1) = .TextMatrix(.Row, 3)
'    msGridSub.TextMatrix(2, 1) = .TextMatrix(.Row, 5)
'    msGridSub.TextMatrix(3, 1) = .TextMatrix(.Row, 7)
'    msGridSub.TextMatrix(4, 1) = .TextMatrix(.Row, 9)
'    DoEvents
'    msGridSub.SetFocus
'    txtGridSub.SetFocus
'End With
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
Select Case Index
    Case 0, 1:
        '------ Validate
        If IsNumeric(txtno(0)) = False Then txtno(0) = 0
        If IsNumeric(txtno(1)) = False Then txtno(1) = txtno(0)
        If CLng(txtno(1)) < CLng(txtno(0)) Then txtno(1) = txtno(0)
        '-------
        If Index = 0 Then GProcCrystalRptPreparation cryReport, "Printer"
        If Index = 1 Then GProcCrystalRptPreparation cryReport, "Window"
        cryReport.WindowTitle = "SIT Booking"
        X = "{tblAddLess.Vno}>=" & txtno(0) & " and {tblAddLess.Vno}<=" & txtno(1) & " and {tblAddLess.VType}='" & frmTypeStr & "' and {tblAddLess.VYear}=" & gCYear
        cryReport.SelectionFormula = X
        cryReport.Formulas(5) = "WCSTNo = '" & gCCST & "'"
        cryReport.Formulas(6) = "WVATNo = '" & gCBST & "'"
        cryReport.ReportFileName = gReportPath & "rptBookingSITPrint.rpt"
        cryReport.Action = 1
    Case 2:
        DoEvents
        Frame1.Visible = False
    End Select
End Sub
