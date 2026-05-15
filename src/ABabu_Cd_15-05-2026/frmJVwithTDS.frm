VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form frmJVwithTDS 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00FBF2E1&
   BorderStyle     =   0  'None
   Caption         =   "Receipt Entry"
   ClientHeight    =   9150
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   16545
   Icon            =   "frmJVwithTDS.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   9150
   ScaleWidth      =   16545
   ShowInTaskbar   =   0   'False
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
      ForeColor       =   &H00000080&
      Height          =   390
      Index           =   60
      Left            =   18240
      MaxLength       =   50
      TabIndex        =   117
      Text            =   "60"
      ToolTipText     =   "Press F3 for Creation & Press F2 for Selection"
      Top             =   6720
      Width           =   420
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
      ForeColor       =   &H00000080&
      Height          =   390
      Index           =   59
      Left            =   18180
      MaxLength       =   50
      TabIndex        =   116
      Text            =   "59"
      ToolTipText     =   "Press F3 for Creation & Press F2 for Selection"
      Top             =   6360
      Width           =   420
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
      ForeColor       =   &H00000080&
      Height          =   390
      Index           =   58
      Left            =   18060
      MaxLength       =   50
      TabIndex        =   115
      Text            =   "58"
      ToolTipText     =   "Press F3 for Creation & Press F2 for Selection"
      Top             =   5940
      Width           =   420
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
      ForeColor       =   &H00000080&
      Height          =   390
      Index           =   57
      Left            =   18180
      MaxLength       =   50
      TabIndex        =   114
      Text            =   "57"
      ToolTipText     =   "Press F3 for Creation & Press F2 for Selection"
      Top             =   5700
      Width           =   420
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
      ForeColor       =   &H00000080&
      Height          =   390
      Index           =   56
      Left            =   18060
      MaxLength       =   50
      TabIndex        =   113
      Text            =   "56"
      ToolTipText     =   "Press F3 for Creation & Press F2 for Selection"
      Top             =   5280
      Width           =   420
   End
   Begin VB.TextBox txtCodeTDSAc 
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
      Left            =   10620
      TabIndex        =   10
      Text            =   "txtCodeTDSAc"
      ToolTipText     =   "Press F3 for Creation"
      Top             =   1860
      Width           =   1410
   End
   Begin VB.TextBox txttempTDSAc 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   12120
      MaxLength       =   60
      TabIndex        =   11
      Text            =   "txttempTDSAc"
      ToolTipText     =   "Press F3 for Creation"
      Top             =   1860
      Width           =   3855
   End
   Begin VB.ComboBox cboNature 
      Appearance      =   0  'Flat
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
      ItemData        =   "frmJVwithTDS.frx":058A
      Left            =   10620
      List            =   "frmJVwithTDS.frx":05A0
      Style           =   2  'Dropdown List
      TabIndex        =   8
      Top             =   960
      Width           =   5310
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
      ForeColor       =   &H00000080&
      Height          =   390
      Index           =   55
      Left            =   16680
      MaxLength       =   50
      TabIndex        =   2
      Text            =   "55"
      ToolTipText     =   "Press F3 for Creation & Press F2 for Selection"
      Top             =   1800
      Width           =   420
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
      ForeColor       =   &H00000080&
      Height          =   390
      Index           =   54
      Left            =   16740
      MaxLength       =   50
      TabIndex        =   107
      Text            =   "54"
      ToolTipText     =   "Press F3 for Creation & Press F2 for Selection"
      Top             =   7740
      Visible         =   0   'False
      Width           =   360
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
      ForeColor       =   &H00000080&
      Height          =   390
      Index           =   53
      Left            =   16740
      MaxLength       =   50
      TabIndex        =   106
      Text            =   "53"
      ToolTipText     =   "Press F3 for Creation & Press F2 for Selection"
      Top             =   7335
      Visible         =   0   'False
      Width           =   360
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
      ForeColor       =   &H00000080&
      Height          =   390
      Index           =   52
      Left            =   16740
      MaxLength       =   50
      TabIndex        =   105
      Text            =   "52"
      ToolTipText     =   "Press F3 for Creation & Press F2 for Selection"
      Top             =   6930
      Visible         =   0   'False
      Width           =   360
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
      ForeColor       =   &H00000080&
      Height          =   390
      Index           =   51
      Left            =   16740
      MaxLength       =   50
      TabIndex        =   104
      Text            =   "51"
      ToolTipText     =   "Press F3 for Creation & Press F2 for Selection"
      Top             =   6525
      Visible         =   0   'False
      Width           =   360
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
      ForeColor       =   &H00000080&
      Height          =   390
      Index           =   50
      Left            =   16740
      MaxLength       =   50
      TabIndex        =   103
      Text            =   "50"
      ToolTipText     =   "Press F3 for Creation & Press F2 for Selection"
      Top             =   6120
      Visible         =   0   'False
      Width           =   360
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
      ForeColor       =   &H00000080&
      Height          =   390
      Index           =   49
      Left            =   16740
      MaxLength       =   50
      TabIndex        =   102
      Text            =   "49"
      ToolTipText     =   "Press F3 for Creation & Press F2 for Selection"
      Top             =   5670
      Visible         =   0   'False
      Width           =   360
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
      ForeColor       =   &H00000080&
      Height          =   390
      Index           =   48
      Left            =   16740
      MaxLength       =   50
      TabIndex        =   101
      Text            =   "48"
      ToolTipText     =   "Press F3 for Creation & Press F2 for Selection"
      Top             =   5220
      Visible         =   0   'False
      Width           =   360
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
      ForeColor       =   &H00000080&
      Height          =   390
      Index           =   47
      Left            =   16740
      MaxLength       =   50
      TabIndex        =   100
      Text            =   "47"
      ToolTipText     =   "Press F3 for Creation & Press F2 for Selection"
      Top             =   4860
      Visible         =   0   'False
      Width           =   360
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
      ForeColor       =   &H00000080&
      Height          =   390
      Index           =   46
      Left            =   16740
      MaxLength       =   50
      TabIndex        =   99
      Text            =   "46"
      ToolTipText     =   "Press F3 for Creation & Press F2 for Selection"
      Top             =   4500
      Visible         =   0   'False
      Width           =   360
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
      ForeColor       =   &H00000080&
      Height          =   390
      Index           =   45
      Left            =   16740
      MaxLength       =   50
      TabIndex        =   98
      Text            =   "45"
      ToolTipText     =   "Press F3 for Creation & Press F2 for Selection"
      Top             =   4095
      Visible         =   0   'False
      Width           =   360
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
      ForeColor       =   &H00000080&
      Height          =   390
      Index           =   44
      Left            =   16740
      MaxLength       =   50
      TabIndex        =   97
      Text            =   "44"
      ToolTipText     =   "Press F3 for Creation & Press F2 for Selection"
      Top             =   3690
      Visible         =   0   'False
      Width           =   360
   End
   Begin VB.TextBox txtCode 
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
      Left            =   1515
      TabIndex        =   4
      Text            =   "txtCode(7)"
      ToolTipText     =   "Press F3 for Creation"
      Top             =   1260
      Width           =   1410
   End
   Begin VB.TextBox txt 
      Height          =   325
      Index           =   43
      Left            =   17550
      TabIndex        =   95
      Text            =   "43"
      Top             =   7695
      Visible         =   0   'False
      Width           =   360
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
      Height          =   450
      Index           =   42
      Left            =   10620
      TabIndex        =   9
      Text            =   "42"
      Top             =   1380
      Width           =   1410
   End
   Begin VB.TextBox txt 
      Height          =   325
      Index           =   41
      Left            =   17550
      TabIndex        =   94
      Text            =   "41"
      Top             =   6975
      Visible         =   0   'False
      Width           =   360
   End
   Begin VB.TextBox txt 
      Height          =   325
      Index           =   40
      Left            =   17550
      TabIndex        =   93
      Text            =   "40"
      Top             =   6615
      Visible         =   0   'False
      Width           =   360
   End
   Begin VB.TextBox txt 
      Height          =   325
      Index           =   39
      Left            =   17595
      TabIndex        =   92
      Text            =   "39"
      Top             =   6255
      Visible         =   0   'False
      Width           =   360
   End
   Begin VB.TextBox txt 
      Height          =   325
      Index           =   38
      Left            =   17595
      TabIndex        =   91
      Text            =   "38"
      Top             =   5940
      Visible         =   0   'False
      Width           =   360
   End
   Begin VB.TextBox txt 
      Height          =   325
      Index           =   37
      Left            =   17595
      TabIndex        =   90
      Text            =   "37"
      Top             =   5580
      Visible         =   0   'False
      Width           =   360
   End
   Begin VB.TextBox txt 
      Height          =   325
      Index           =   36
      Left            =   17595
      TabIndex        =   89
      Text            =   "36"
      Top             =   5265
      Visible         =   0   'False
      Width           =   360
   End
   Begin VB.TextBox txt 
      Height          =   325
      Index           =   35
      Left            =   17595
      TabIndex        =   88
      Text            =   "35"
      Top             =   4950
      Visible         =   0   'False
      Width           =   360
   End
   Begin VB.TextBox txt 
      Height          =   325
      Index           =   34
      Left            =   17595
      TabIndex        =   87
      Text            =   "34"
      Top             =   4590
      Visible         =   0   'False
      Width           =   360
   End
   Begin VB.TextBox txt 
      Height          =   325
      Index           =   33
      Left            =   17595
      TabIndex        =   86
      Text            =   "33"
      Top             =   4275
      Visible         =   0   'False
      Width           =   360
   End
   Begin VB.TextBox txt 
      Height          =   325
      Index           =   32
      Left            =   17595
      TabIndex        =   85
      Text            =   "32"
      Top             =   3960
      Visible         =   0   'False
      Width           =   360
   End
   Begin VB.TextBox txt 
      Height          =   325
      Index           =   31
      Left            =   17595
      TabIndex        =   84
      Text            =   "31"
      Top             =   3645
      Visible         =   0   'False
      Width           =   360
   End
   Begin VB.TextBox txt 
      Height          =   325
      Index           =   30
      Left            =   17595
      TabIndex        =   83
      Text            =   "30"
      Top             =   3330
      Visible         =   0   'False
      Width           =   360
   End
   Begin VB.TextBox txt 
      Height          =   325
      Index           =   29
      Left            =   17595
      TabIndex        =   82
      Text            =   "29"
      Top             =   2970
      Visible         =   0   'False
      Width           =   360
   End
   Begin VB.TextBox txt 
      Height          =   325
      Index           =   28
      Left            =   17595
      TabIndex        =   81
      Text            =   "28"
      Top             =   2610
      Visible         =   0   'False
      Width           =   360
   End
   Begin VB.TextBox txt 
      Height          =   325
      Index           =   27
      Left            =   17595
      TabIndex        =   80
      Text            =   "27"
      Top             =   2250
      Visible         =   0   'False
      Width           =   360
   End
   Begin VB.TextBox txt 
      Height          =   325
      Index           =   26
      Left            =   17595
      TabIndex        =   77
      Text            =   "26"
      Top             =   1890
      Visible         =   0   'False
      Width           =   360
   End
   Begin VB.TextBox txt 
      Height          =   325
      Index           =   25
      Left            =   17595
      TabIndex        =   76
      Text            =   "25"
      Top             =   1575
      Visible         =   0   'False
      Width           =   360
   End
   Begin VB.TextBox txt 
      Height          =   325
      Index           =   24
      Left            =   17550
      TabIndex        =   75
      Text            =   "24"
      Top             =   1215
      Visible         =   0   'False
      Width           =   360
   End
   Begin VB.TextBox txt 
      Height          =   325
      Index           =   23
      Left            =   17550
      TabIndex        =   74
      Text            =   "23"
      Top             =   900
      Visible         =   0   'False
      Width           =   360
   End
   Begin VB.TextBox txt 
      Height          =   325
      Index           =   22
      Left            =   17550
      TabIndex        =   73
      Text            =   "22"
      Top             =   540
      Visible         =   0   'False
      Width           =   360
   End
   Begin VB.TextBox txt 
      Height          =   325
      Index           =   21
      Left            =   17595
      TabIndex        =   72
      Text            =   "21"
      Top             =   180
      Visible         =   0   'False
      Width           =   360
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
      Left            =   1515
      TabIndex        =   18
      Text            =   "8"
      Top             =   1800
      Width           =   2535
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
      Index           =   7
      Left            =   3045
      MaxLength       =   60
      TabIndex        =   6
      Text            =   "txttemp(7)"
      ToolTipText     =   "Press F3 for Creation"
      Top             =   1260
      Width           =   5055
   End
   Begin VB.ComboBox cboCrDr 
      Appearance      =   0  'Flat
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
      ItemData        =   "frmJVwithTDS.frx":0636
      Left            =   8190
      List            =   "frmJVwithTDS.frx":0640
      Style           =   2  'Dropdown List
      TabIndex        =   7
      Top             =   1260
      Width           =   1050
   End
   Begin VB.TextBox txt 
      Height          =   325
      Index           =   20
      Left            =   18270
      TabIndex        =   71
      Text            =   "20"
      Top             =   4770
      Visible         =   0   'False
      Width           =   360
   End
   Begin VB.TextBox txt 
      Height          =   325
      Index           =   19
      Left            =   18225
      TabIndex        =   70
      Text            =   "19"
      Top             =   4410
      Visible         =   0   'False
      Width           =   360
   End
   Begin VB.TextBox txt 
      Height          =   325
      Index           =   18
      Left            =   18225
      TabIndex        =   69
      Text            =   "18"
      Top             =   4050
      Visible         =   0   'False
      Width           =   360
   End
   Begin VB.TextBox txt 
      Height          =   325
      Index           =   17
      Left            =   18180
      TabIndex        =   68
      Text            =   "17"
      Top             =   3735
      Visible         =   0   'False
      Width           =   360
   End
   Begin VB.TextBox txt 
      Height          =   325
      Index           =   12
      Left            =   18135
      TabIndex        =   67
      Text            =   "12"
      Top             =   3375
      Visible         =   0   'False
      Width           =   360
   End
   Begin VB.TextBox txt 
      Height          =   325
      Index           =   11
      Left            =   18135
      TabIndex        =   66
      Text            =   "11"
      Top             =   3015
      Visible         =   0   'False
      Width           =   360
   End
   Begin VB.TextBox txt 
      Height          =   325
      Index           =   9
      Left            =   18180
      TabIndex        =   65
      Text            =   "9"
      Top             =   2655
      Visible         =   0   'False
      Width           =   330
   End
   Begin VB.TextBox txt 
      Alignment       =   1  'Right Justify
      Height          =   360
      Index           =   6
      Left            =   18180
      TabIndex        =   64
      Text            =   "6"
      Top             =   1935
      Visible         =   0   'False
      Width           =   345
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
      ForeColor       =   &H00000080&
      Height          =   390
      Index           =   16
      Left            =   16740
      MaxLength       =   50
      TabIndex        =   17
      Text            =   "16"
      ToolTipText     =   "Press F3 for Creation & Press F2 for Selection"
      Top             =   3330
      Visible         =   0   'False
      Width           =   360
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
      ForeColor       =   &H00000080&
      Height          =   390
      Index           =   15
      Left            =   16740
      MaxLength       =   50
      TabIndex        =   16
      Text            =   "15"
      ToolTipText     =   "Press F3 for Creation & Press F2 for Selection"
      Top             =   2925
      Visible         =   0   'False
      Width           =   405
   End
   Begin VB.CheckBox chkDontShowInUnadj 
      BackColor       =   &H00FFC0FF&
      Caption         =   "Dont Show In Unadjusted Receipt"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   6345
      TabIndex        =   63
      Top             =   9540
      Width           =   4515
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFC0C0&
      Caption         =   "Voucher Printing"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2835
      Left            =   5130
      TabIndex        =   54
      Top             =   3540
      Visible         =   0   'False
      Width           =   5175
      Begin VB.CommandButton cmdPrint 
         Caption         =   "S&hort Note"
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
         Left            =   1950
         TabIndex        =   62
         Top             =   2160
         Width           =   1335
      End
      Begin VB.TextBox txtno 
         Appearance      =   0  'Flat
         Height          =   360
         Index           =   0
         Left            =   1950
         MaxLength       =   7
         TabIndex        =   59
         Top             =   600
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
         TabIndex        =   58
         Top             =   1560
         Width           =   1335
      End
      Begin VB.TextBox txtno 
         Appearance      =   0  'Flat
         Height          =   360
         Index           =   1
         Left            =   1950
         MaxLength       =   7
         TabIndex        =   57
         Top             =   1080
         Width           =   1575
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
         TabIndex        =   56
         Top             =   1560
         Width           =   1335
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
         TabIndex        =   55
         Top             =   1560
         Width           =   1335
      End
      Begin VB.Label Label31 
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
         TabIndex        =   61
         Top             =   600
         Width           =   1065
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
         TabIndex        =   60
         Top             =   1080
         Width           =   1050
      End
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
      Left            =   14535
      Style           =   1  'Graphical
      TabIndex        =   53
      ToolTipText     =   "Exit"
      Top             =   8385
      Width           =   1335
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
      Left            =   13110
      Style           =   1  'Graphical
      TabIndex        =   52
      ToolTipText     =   "To Cancel Record"
      Top             =   8385
      Width           =   1335
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
      Left            =   11700
      Style           =   1  'Graphical
      TabIndex        =   51
      ToolTipText     =   "To Save Record"
      Top             =   8385
      Width           =   1335
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
      Left            =   10095
      Style           =   1  'Graphical
      TabIndex        =   50
      Top             =   8385
      Width           =   1335
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
      Left            =   3870
      Style           =   1  'Graphical
      TabIndex        =   49
      ToolTipText     =   "To Delete Record"
      Top             =   8370
      Width           =   1335
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
      Left            =   2475
      Style           =   1  'Graphical
      TabIndex        =   48
      ToolTipText     =   "To Modify Record"
      Top             =   8370
      Width           =   1335
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
      Left            =   1035
      Style           =   1  'Graphical
      TabIndex        =   47
      ToolTipText     =   "To Add New Record"
      Top             =   8370
      Width           =   1335
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
      Left            =   5400
      Style           =   1  'Graphical
      TabIndex        =   46
      ToolTipText     =   "To Find Record"
      Top             =   8370
      Width           =   1335
   End
   Begin VB.CommandButton cmdIntDbNt 
      Caption         =   "&Interest Debit Note"
      Height          =   945
      Left            =   19530
      TabIndex        =   45
      Top             =   7065
      Visible         =   0   'False
      Width           =   795
   End
   Begin VB.TextBox txt 
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Index           =   10
      Left            =   18450
      TabIndex        =   20
      Text            =   "10"
      Top             =   2790
      Visible         =   0   'False
      Width           =   360
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
      Left            =   18630
      TabIndex        =   42
      Text            =   "txtCrBalance"
      Top             =   180
      Visible         =   0   'False
      Width           =   615
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
      Left            =   18630
      TabIndex        =   41
      Text            =   "txtDrBalance"
      Top             =   585
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.CommandButton cmdLedger 
      Caption         =   "Ledger"
      Height          =   375
      Index           =   0
      Left            =   19215
      TabIndex        =   40
      Top             =   5085
      Visible         =   0   'False
      Width           =   1200
   End
   Begin VB.CommandButton cmdLedger 
      Caption         =   "Ledger"
      Height          =   375
      Index           =   1
      Left            =   19305
      TabIndex        =   39
      Top             =   5490
      Visible         =   0   'False
      Width           =   1185
   End
   Begin VB.TextBox txtGrid 
      BackColor       =   &H8000000A&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   480
      TabIndex        =   13
      Top             =   4815
      Visible         =   0   'False
      Width           =   885
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
      ForeColor       =   &H00000080&
      Height          =   390
      Index           =   13
      Left            =   16740
      MaxLength       =   50
      TabIndex        =   14
      Text            =   "13"
      ToolTipText     =   "Press F3 for Creation & Press F2 for Selection"
      Top             =   2115
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.TextBox txt 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFFFF&
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
      Height          =   360
      Index           =   0
      Left            =   1515
      TabIndex        =   1
      Text            =   "0"
      Top             =   585
      Width           =   2175
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
      Index           =   6
      Left            =   16125
      MaxLength       =   40
      TabIndex        =   19
      Text            =   "txttemp(6)"
      ToolTipText     =   "Press F3 for Creation"
      Top             =   630
      Visible         =   0   'False
      Width           =   1620
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
      Left            =   6930
      Style           =   1  'Graphical
      TabIndex        =   24
      ToolTipText     =   "Previous"
      Top             =   8370
      Width           =   540
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
      Left            =   7785
      Style           =   1  'Graphical
      TabIndex        =   25
      ToolTipText     =   "Next"
      Top             =   8385
      Width           =   585
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
      Left            =   8550
      Style           =   1  'Graphical
      TabIndex        =   26
      ToolTipText     =   "First"
      Top             =   8385
      Width           =   585
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
      Left            =   9315
      Style           =   1  'Graphical
      TabIndex        =   27
      ToolTipText     =   "Last"
      Top             =   8385
      Width           =   585
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
      Left            =   19350
      Style           =   1  'Graphical
      TabIndex        =   28
      Top             =   5895
      Visible         =   0   'False
      Width           =   975
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
      Left            =   19305
      Style           =   1  'Graphical
      TabIndex        =   32
      Top             =   6615
      Visible         =   0   'False
      Width           =   1095
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
      Left            =   19305
      Style           =   1  'Graphical
      TabIndex        =   33
      Top             =   6300
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.TextBox txt 
      Height          =   345
      Index           =   3
      Left            =   18135
      TabIndex        =   36
      Text            =   "3"
      Top             =   900
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.TextBox txt 
      Alignment       =   1  'Right Justify
      Height          =   360
      Index           =   5
      Left            =   18135
      TabIndex        =   31
      Text            =   "5"
      Top             =   1575
      Visible         =   0   'False
      Width           =   345
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
      ForeColor       =   &H00000080&
      Height          =   390
      Index           =   14
      Left            =   16740
      MaxLength       =   50
      TabIndex        =   15
      Text            =   "14"
      ToolTipText     =   "Press F3 for Creation & Press F2 for Selection"
      Top             =   2520
      Visible         =   0   'False
      Width           =   450
   End
   Begin VB.TextBox txt 
      Height          =   325
      Index           =   7
      Left            =   18180
      TabIndex        =   29
      Text            =   "7"
      Top             =   2340
      Visible         =   0   'False
      Width           =   330
   End
   Begin VB.TextBox txt 
      Alignment       =   1  'Right Justify
      Height          =   360
      Index           =   4
      Left            =   18135
      TabIndex        =   30
      Text            =   "4"
      Top             =   1260
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.TextBox txt 
      Height          =   345
      Index           =   2
      Left            =   18135
      TabIndex        =   35
      Text            =   "2"
      Top             =   540
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.TextBox txt 
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
      Height          =   345
      Index           =   1
      Left            =   18135
      TabIndex        =   34
      Text            =   "1"
      Top             =   180
      Visible         =   0   'False
      Width           =   375
   End
   Begin MSMask.MaskEdBox mskDt 
      Height          =   360
      Index           =   2
      Left            =   16305
      TabIndex        =   5
      Top             =   1170
      Visible         =   0   'False
      Width           =   1620
      _ExtentX        =   2858
      _ExtentY        =   635
      _Version        =   393216
      ClipMode        =   1
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
   Begin MSFlexGridLib.MSFlexGrid msGrid 
      Height          =   4680
      Left            =   210
      TabIndex        =   12
      Top             =   2835
      Width           =   15945
      _ExtentX        =   28125
      _ExtentY        =   8255
      _Version        =   393216
      FixedCols       =   0
      RowHeightMin    =   275
      BackColorFixed  =   -2147483647
      ForeColorFixed  =   -2147483634
      BackColorBkg    =   12632256
      Appearance      =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin Crystal.CrystalReport cryReport 
      Bindings        =   "frmJVwithTDS.frx":064C
      Left            =   -45
      Top             =   660
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      WindowControlBox=   -1  'True
      WindowMaxButton =   -1  'True
      WindowMinButton =   -1  'True
      PrintFileType   =   15
      PrintFileLinesPerPage=   60
   End
   Begin MSComCtl2.DTPicker dtpDt 
      Height          =   360
      Index           =   2
      Left            =   7515
      TabIndex        =   3
      Top             =   585
      Width           =   1710
      _ExtentX        =   3016
      _ExtentY        =   635
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      CustomFormat    =   "dd/MM/yyyy"
      Format          =   99811331
      CurrentDate     =   -273
   End
   Begin VB.Label lbl1 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "TotAmt"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C000C0&
      Height          =   315
      Index           =   12
      Left            =   5715
      TabIndex        =   120
      Top             =   7650
      Width           =   1635
   End
   Begin VB.Label lbl1 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "TdsAmt"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C000C0&
      Height          =   315
      Index           =   11
      Left            =   9180
      TabIndex        =   119
      Top             =   7650
      Width           =   1410
   End
   Begin VB.Label lbl1 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "TdsOn"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C000C0&
      Height          =   315
      Index           =   10
      Left            =   7560
      TabIndex        =   118
      Top             =   7650
      Width           =   1410
   End
   Begin VB.Label lbl1 
      BackStyle       =   0  'Transparent
      Caption         =   "%"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   375
      Index           =   8
      Left            =   12120
      TabIndex        =   112
      Top             =   1440
      Width           =   915
   End
   Begin VB.Label lbl1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "TDS Details"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C000C0&
      Height          =   315
      Index           =   7
      Left            =   9780
      TabIndex        =   111
      Top             =   600
      Width           =   2895
   End
   Begin VB.Label lbl1 
      BackStyle       =   0  'Transparent
      Caption         =   "A/c  Cd :"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   375
      Index           =   6
      Left            =   9660
      TabIndex        =   110
      Top             =   1920
      Width           =   915
   End
   Begin VB.Label lbl1 
      BackStyle       =   0  'Transparent
      Caption         =   "Rate     :"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   375
      Index           =   5
      Left            =   9660
      TabIndex        =   109
      Top             =   1440
      Width           =   915
   End
   Begin VB.Label lbl1 
      BackStyle       =   0  'Transparent
      Caption         =   "Nature  :"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   375
      Index           =   2
      Left            =   9660
      TabIndex        =   108
      Top             =   1020
      Width           =   915
   End
   Begin VB.Shape Shape4 
      BackColor       =   &H00808080&
      BorderColor     =   &H00000080&
      FillColor       =   &H00ECFFFF&
      FillStyle       =   0  'Solid
      Height          =   1905
      Left            =   9480
      Top             =   480
      Width           =   6645
   End
   Begin VB.Label lblCrDrNtNo 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Entry &No.  :"
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
      Left            =   16680
      TabIndex        =   96
      Top             =   180
      Width           =   1215
   End
   Begin VB.Label lblSecondAc 
      BackStyle       =   0  'Transparent
      Caption         =   "Credit / Debit To  A/c   :"
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
      Left            =   225
      TabIndex        =   79
      Top             =   2475
      Width           =   3330
   End
   Begin VB.Label lblCrDrAc 
      BackStyle       =   0  'Transparent
      Caption         =   "A/c   :"
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
      Index           =   2
      Left            =   15075
      TabIndex        =   78
      Top             =   7650
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Label lbl1 
      BackStyle       =   0  'Transparent
      Caption         =   "&Party        :"
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
      Index           =   3
      Left            =   16260
      TabIndex        =   44
      Top             =   315
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Label lblBalance 
      BackColor       =   &H00ECFFFE&
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
      Left            =   15300
      TabIndex        =   43
      Top             =   7200
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.Shape Shape6 
      BorderColor     =   &H00EFB17A&
      BorderWidth     =   5
      FillColor       =   &H00EFB17A&
      Height          =   9075
      Left            =   0
      Top             =   0
      Width           =   16305
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
      Left            =   12495
      TabIndex        =   38
      Top             =   0
      Width           =   3180
   End
   Begin VB.Label lbl1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Entry &No.  :"
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
      Index           =   0
      Left            =   360
      TabIndex        =   0
      Top             =   585
      Width           =   1035
   End
   Begin VB.Label lbl1 
      BackStyle       =   0  'Transparent
      Caption         =   "Am&ount    :"
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
      Height          =   255
      Index           =   9
      Left            =   360
      TabIndex        =   23
      Top             =   1845
      Width           =   1215
   End
   Begin VB.Label lbl1 
      BackStyle       =   0  'Transparent
      Caption         =   "A/c           :"
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
      Index           =   4
      Left            =   360
      TabIndex        =   22
      Top             =   1305
      Width           =   1215
   End
   Begin VB.Label lbl1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Da&te   :"
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
      Index           =   1
      Left            =   6705
      TabIndex        =   21
      Top             =   630
      Width           =   780
   End
   Begin VB.Shape Shape1 
      FillColor       =   &H00BDDFD1&
      FillStyle       =   0  'Solid
      Height          =   825
      Left            =   315
      Shape           =   4  'Rounded Rectangle
      Top             =   8130
      Width           =   15840
   End
   Begin VB.Label Label6 
      BackColor       =   &H00EFB17A&
      Caption         =   "   J.V. with TDS"
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
      TabIndex        =   37
      Top             =   30
      Width           =   16275
   End
   Begin VB.Shape Shape2 
      BackColor       =   &H00808080&
      BorderColor     =   &H00000080&
      FillColor       =   &H00ECFFFF&
      FillStyle       =   0  'Solid
      Height          =   585
      Left            =   240
      Top             =   465
      Width           =   9165
   End
   Begin VB.Shape Shape3 
      BackColor       =   &H00ECFFFE&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00000080&
      FillColor       =   &H00ECFFFF&
      Height          =   1245
      Left            =   225
      Top             =   1125
      Width           =   9120
   End
End
Attribute VB_Name = "frmJVwithTDS"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Const tblName As String = "tblVoucher"
Const IdField As String = "Vno"
Const MaxNo As Long = 60
Dim CriteriaStr As String
Dim rstbl As Recordset, ShowFrmFlag As Boolean
Dim rsRecVsSale As Recordset
Dim varBookMark As Variant
Dim FirstTimeFlag As Boolean, mLastEntryBank As String, mLastEntryDate As Date
Dim FormAction As Integer
Dim frmTypeStr As String
Dim i As Long
Dim Mon As Integer
Dim OldgFrmTypeStr As String
Const OrderField As String = "year(vdt),Vno"
Private Sub cboCrDr_LostFocus()
'   msGrid.Col = 1
'   msGrid_EnterCell
   If cboCrDr.ListIndex = 0 Then
      lblSecondAc.Caption = "Debit To A/c :"
   Else
      lblSecondAc.Caption = "Credit To A/c :"
   End If
End Sub
Private Sub cmdBtn_Click(Index As Integer)
Dim FindQStr As String
Dim X As String
'On Error GoTo ErrorRoutine
Select Case Index
Case 0 'Previous
    If GProcRstOpen(rstbl, "Select * from " & tblName & " where vno = (Select max(Vno) from " & tblName & " where Vno <> 0  and Vno < " & txt(0) & " and  " & CriteriaStr & "  ) and  " & CriteriaStr & "  ", "O") > 0 Then
       ReadFields
    Else
      MsgBox "This is First Record ", vbOKOnly + vbInformation
    End If
Case 1 'Next
    If GProcRstOpen(rstbl, "Select * from " & tblName & " where Vno = (Select min(Vno) from " & tblName & " where Vno <> 0  and Vno > " & txt(0) & "  and " & CriteriaStr & ") and  " & CriteriaStr & "   ", "O") > 0 Then
       ReadFields
    Else
       MsgBox "This is Last Record ", vbOKOnly + vbInformation
    End If
Case 2 'First
         If GProcRstOpen(rstbl, "Select * from " & tblName & " where " & CriteriaStr & "  and  " & IdField & " = (Select min(" & IdField & ") from " & tblName & " where " & IdField & " <> '' and " & CriteriaStr & ")and " & CriteriaStr, "R") > 0 Then ReadFields
Case 3 'Last
'        If GProcRstOpen(rstbl, "Select * from " & tblName & " where " & CriteriaStr & "  and  " & IdField & " = (Select max(" & IdField & ") from " & tblName & " where " & IdField & " <> '' and " & CriteriaStr & ")and " & CriteriaStr, "R") > 0 Then ReadFields Else GProcClearForm Me, rstbl, MaxNo
        If GProcRstOpen(rstbl, "Select * from " & tblName & " where " & CriteriaStr & "  and  " & IdField & " = (Select max(" & IdField & ") from " & tblName & " where vdt = (SELECT MAX(vdt) from " & tblName & " where " & CriteriaStr & " )  and  " & IdField & " <> 0 and  " & CriteriaStr & ")", "R") > 0 Then ReadFields Else GProcClearForm Me, rstbl, MaxNo

Case 4 'Find
    gClsSearch.SearchMultiField "qryVoucherHelp", "Vno,Vdt,CreditAc,DebitAc,VAmt,VCtrNo", Array("Vno", "Vdt", "Credit A/c", "Debit A/c", "Amount", "VCtrNo"), Array(1700, 1500, 4000, 4000, 2000, 0), CriteriaStr, , OrderField, , , True
    If gClsSearch.SearchMultiRetCol(0) <> "" Then
        If GProcRstOpen(rstbl, "Select * from " & tblName & " where Vno = (Select max(" & gClsSearch.SearchMultiRetCol(0) & ") from " & tblName & " where Vno <> 0 and " & CriteriaStr & " )and  " & CriteriaStr & "   ", "O") > 0 Then ReadFields
    End If
Case 5 'List
    'GenMastList ("LAC")
Case 6 'Add
    FormAction = vbDataActionAddNew
    Call GProcActivateControls(Me, True)
    Call GProcSetButtons(Me, vbDataActionAddNew)
    Call GProcClearForm(Me, rstbl, MaxNo, True)
    If Year(mLastEntryDate) = 1899 Then
        dtpDt(2) = Format(Now, "dd/MM/yyyy")
    Else
        dtpDt(2) = mLastEntryDate
    End If
    Select Case frmTypeStr
    Case "SN"
          txt(55).text = GProcGenerateNoWithIBCrDrNt("SN", gCCode)
          cboCrDr.ListIndex = 1
          lblSecondAc.Caption = "Credit To A/c :"
    Case "PN"
          txt(55).text = GProcGenerateNoWithIBCrDrNt("PN", gCCode)
          cboCrDr.ListIndex = 0
          lblSecondAc.Caption = "Debit To A/c :"
    Case "SQ"
          txt(55).text = GProcGenerateNoWithIBCrDrNt("SQ", gCCode)
          cboCrDr.ListIndex = 1
          lblSecondAc.Caption = "Credit To A/c :"
    Case "PQ"
          txt(55).text = GProcGenerateNoWithIBCrDrNt("PQ", gCCode)
          cboCrDr.ListIndex = 0
          lblSecondAc.Caption = "Debit To A/c :"
    End Select

'    txt(0).text = GProcGenerateId(tblName, "VNO", CriteriaStr, gCn)
'    dtpdt(2) = GProcGenerateVDt(tblName, "VDt", CriteriaStr)
    If FirstTimeFlag = True Then
        FirstTimeFlag = False
    Else
        dtpDt(2).SetFocus
'        txt(0).SetFocus
    End If
    SetControlEd
    SetGrid
Case 7 'Modify
    'If mURecModi = True Then
    '    SearchHelp
    '    DoEvents
    If IsRelatedRecord = False Then
        FormAction = vbDataActionUpdate
        Call GProcActivateControls(Me, True)
        Call GProcSetButtons(Me, vbDataActionUpdate)
        txt(0).Enabled = False
        txt(55).Enabled = False
        dtpDt(2).SetFocus
    End If
Case 8 'Delete
    Dim x1Rec As Double
   ' If mURecDel = True Then
     If IsRelatedRecord = False Then
        If MsgBox("Do You Really Delete This Record?", vbQuestion + vbYesNo + vbDefaultButton1, "Confirm Delete") = vbYes Then
          X = "select VCtrNo from tblVoucher where VType='" & frmTypeStr & "' and VNo=" & txt(0) & " and VFirm  = '" & gCCode & "' and  VYear=" & gCYear & "  order by VCtrNo"
          x1Rec = GProcRstOpen(rsRecVsSale, X, "R")
           With rsRecVsSale
                If Not .EOF Then
                  .MoveFirst
                  Do While Not .EOF
                     gCn.Execute "Delete from tblvoucher where " & CriteriaStr & " and VCtrNo = " & rsRecVsSale.Fields(0) & "   and  vno = " & txt(0)
                    .MoveNext
                  Loop
                End If
           End With
            '-- Delete Related Record
            DeleteRelatedRecord
            If GProcRstOpen(rstbl, "Select * from " & tblName & " where Vno = (Select max(Vno) from " & tblName & " where Vno <> 0 and " & CriteriaStr & " )and  " & CriteriaStr & "   ", "O") > 0 Then ReadFields
            If rstbl.RecordCount > 1 Then
                rstbl.Requery
                cmdBtn_Click (1)
            ElseIf rstbl.RecordCount = 1 Then
                Call GProcActivateControls(Me, False)
                Call GProcSetButtons(Me, vbDataActionCancel)
                Form_Load
            End If
        End If
        DoEvents
    End If
   ' End If
Case 9 'Print
    Frame1.Visible = True
    DoEvents
    txtno(0).Enabled = True
    txtno(1).Enabled = True
    txtno(0) = txt(0)
    txtno(1) = txt(0)
    txtno(0).SetFocus
Case 10 'Exit
    DoEvents
    Unload Me
Case 13 'Save
    If ValidateData = True Then
       If Trim(txtGrid) <> "" Then
            msGrid.text = txtGrid
            DoEvents
            txtGrid = ""
        End If
        txtGrid.Visible = False
        FillTxtFromTemp
        '--- Related Rcord Rec Vs Sale
        If FormAction = vbDataActionAddNew Then
           Mon = Month(dtpDt(2))
           txt(0) = GProcGenerateIdMonthwise("TblVoucher", "Vno", Mon, "Vdt", "Vtype='" & frmTypeStr & "' and VFirm = '" & gCCode & "'   and VYear=" & gCYear)
        End If
        txt(55) = UCase(txt(55))
        GProcRstOpen rstbl, "Select * from " & tblName & " where " & IdField & " = " & txt(0) & " and " & CriteriaStr, "O"
        SaveRelatedRecord
        rstbl.Requery
        Call GProcActivateControls(Me, False)
        Call GProcSetButtons(Me, True)
        mLastEntryDate = dtpDt(2)
        FormAction = vbDataActionClose
        msGrid.Col = 0
        msGrid.Row = 1
        txtGrid.Move msGrid.CellLeft + msGrid.Left, msGrid.CellTop + msGrid.Top, msGrid.CellWidth, msGrid.CellHeight
        lblBalance.Caption = ""
'        If FormAction = vbDataActionAddNew Then
            If cmdBtn(6).Visible = True Then cmdBtn(6).SetFocus
 '       Else
  '          If cmdBtn(7).Visible = True Then cmdBtn(7).SetFocus
   '     End If
    End If
Case 14 'Cancel
    txtGrid.Visible = False
    
    If GProcRstOpen(rstbl, "Select * from " & tblName & " where Vno = (Select max(Vno) from " & tblName & " where Vno <> 0 and " & CriteriaStr & " )and  " & CriteriaStr & "   ", "O") > 0 Then ReadFields
    
    
    If rstbl.RecordCount <= 0 Then
        Call GProcClearForm(Me, rstbl, MaxNo, True)
        Call GProcActivateControls(Me, False)
        Call GProcSetButtons(Me, vbDataActionCancel)
    Else
        If FormAction = vbDataActionAddNew Then
            rstbl.MoveLast
        Else
            rstbl.CancelUpdate
        End If
        Call GProcClearForm(Me, rstbl, MaxNo, True)
       
        ReadFields
        Call GProcActivateControls(Me, False)
        Call GProcSetButtons(Me, True)
    End If
    FormAction = vbDataActionClose
    msGrid.Col = 0
    If msGrid.Row >= 1 Then msGrid.Row = 1
    txtGrid.Move msGrid.CellLeft + msGrid.Left, msGrid.CellTop + msGrid.Top, msGrid.CellWidth, msGrid.CellHeight
    lblBalance.Caption = ""
    If FormAction = vbDataActionAddNew Then
        If cmdBtn(6).Visible = True Then cmdBtn(6).SetFocus
    Else
        If cmdBtn(6).Visible = True And cmdBtn(6).Enabled = True Then
           cmdBtn(6).SetFocus
        Else
           cmdBtn(10).SetFocus
        End If
    End If
End Select
Exit Sub
ErrorRoutine:
    rstbl.CancelUpdate
    GProcErrorHandler
End Sub
Private Sub txtCodeTDSAc_LostFocus()
If Trim(txtCodeTDSAc) = "" Then Exit Sub
         
         txttempTDSAc = GProcGetColumnValue("tblMastAccount", "AcOurCode", UCase(txtCodeTDSAc), "S", "AcName", "S")
         If Trim(txttempTDSAc) = "" Then
            MsgBox ("Code Not Found ...")
            txttempTDSAc = ""
            txttempTDSAc.Enabled = True
            txttempTDSAc.SetFocus
         End If
End Sub

'------  Print Voucher
Private Sub txtno_GotFocus(Index As Integer)
GProcSelectBox Me.ActiveControl
End Sub
Private Sub txtno_KeyPress(Index As Integer, KeyAscii As Integer)
KeyAscii = GProcNumberOnly(KeyAscii)
End Sub
Private Sub cmdPrint_Click(Index As Integer)
Dim Rs1 As Recordset
Dim X As String
Dim Vno As Long
Select Case Index
Case 0, 1:
    '------ Validate
    If IsNumeric(txtno(0)) = False Then txtno(0) = 0
    If IsNumeric(txtno(1)) = False Then txtno(1) = txtno(0)
    If CLng(txtno(1)) < CLng(txtno(0)) Then txtno(1) = txtno(0)
    '-------
    If Index = 0 Then GProcCrystalRptPreparation frmMain.cryReport, "Printer"
    If Index = 1 Then GProcCrystalRptPreparation frmMain.cryReport, "Window"
    frmMain.cryReport.WindowTitle = "Receipt Voucher"
    Vno = CLng(txtno(0))
    Do While Vno <= txtno(1)
        X = "Select * From tblVoucher where  vNo=" & Vno & " and VCTRNO =1 and VType='" & frmTypeStr & "' and VYear=" & gCYear
        If GProcRstOpen(Rs1, X, "R") > 0 Then
            With frmMain.cryReport
                .SelectionFormula = "{tblVoucher.vNo} = " & Rs1!Vno & " and {tblVoucher.VCTRNO}=1 and {tblVoucher.VType}='" & Rs1!Vtype & "' and {tblVoucher.VYear}=" & Rs1!VYear
                .Formulas(5) = "wAmountInWords = '" & GProcAmountInWords(Abs(Rs1!vamt)) & "'"
                .ReportFileName = gReportPath & "rptVouReceiptPrint.rpt"
                .Action = 1
            End With
        End If
        Vno = Vno + 1
    Loop
Case 2:
    DoEvents
    Frame1.Visible = False
Case 3:
    '------ Validate
    If IsNumeric(txtno(0)) = False Then txtno(0) = 0
    If IsNumeric(txtno(1)) = False Then txtno(1) = txtno(0)
    If CLng(txtno(1)) < CLng(txtno(0)) Then txtno(1) = txtno(0)
    '-------
    If Index = 3 Then GProcCrystalRptPreparation frmMain.cryReport, "Printer"
    If Index = 3 Then GProcCrystalRptPreparation frmMain.cryReport, "Window"
    frmMain.cryReport.WindowTitle = "Short Payment Voucher"
    Vno = CLng(txtno(0))
    Do While Vno <= txtno(1)
        X = "Select * From qryshortrecvoucher where  RecVNo=" & Vno & " and RecVType='" & frmTypeStr & "' and RecVYear=" & gCYear
        If GProcRstOpen(Rs1, X, "R") > 0 Then
            With frmMain.cryReport
                .SelectionFormula = "{qryshortrecvoucher.RecVno} = " & Rs1!RecVno & "  and {qryshortrecvoucher.RecVType}='" & Rs1!RecVType & "' and {qryshortrecvoucher.RecVYear}=" & Rs1!RecVYear
                .ReportFileName = gReportPath & "rptVouReceiptShort.rpt"
                .Action = 1
            End With
        End If
        Vno = Vno + 1
    Loop
End Select
End Sub
Private Sub cmdLedger_Click(Index As Integer)
On Error GoTo ErrorHandler
If Index = 0 Then
    frmRptAccount.txt(1) = txttemp(7)
Else
    frmRptAccount.txt(1) = txttemp(6)
End If
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
'            If (cmdBtn(13).Enabled = False) Then
'                Unload Me
'            End If
    End Select
End Sub
Private Sub Form_Load()
    Dim X As String
    Dim Y As String
    mLastEntryBank = ""
    Me.Top = 300
    Me.Left = Screen.Width / 2 - Me.Width / 2
    frmTypeStr = gfrmTypeStr
    Select Case frmTypeStr
    Case "JV"
        Label6.Caption = "   Journal Voucher"
    Case "SN"
        lblCrDrNtNo.Caption = "Debit Note No  :"
        Label6.Caption = "   Debit Note ( Sales )"
    Case "PN"
        lblCrDrNtNo.Caption = "Credit Note No :"
        Label6.Caption = "   Credit Note ( Sales )"
    Case "SQ"
        lblCrDrNtNo.Caption = "Debit Note No  :"
        Label6.Caption = "   Debit Note ( Purchase )"
    Case "PQ"
        lblCrDrNtNo.Caption = "Credit Note No :"
        Label6.Caption = "   Credit Note ( Purchase )"
    End Select
    CriteriaStr = " VType = '" & frmTypeStr & "' and VFirm  = '" & gCCode & "' and VYear=" & gCYear
    Y = "Select * from " & tblName & " where " & CriteriaStr & " "
    If Trim(OrderField) <> "" Then Y = Y & " order by " & OrderField
'    GProcRstOpen rstbl, Y, "O"
    FirstTimeFlag = True
    cboCrDr.ListIndex = 0
    '---Grid
    If Trim(mShowRec) <> "" Then
       GProcRstOpen rstbl, "Select * from " & tblName & " where  " & CriteriaStr & " and  " & IdField & " = " & CLng(Trim(mShowRec)), "R"
       ReadFields
       Call GProcSetButtons(Me, vbDataActionCancel)
       mShowRec = ""
    Else
       GProcRstOpen rstbl, "Select * from " & tblName & " where " & CriteriaStr & " and  " & IdField & " = 0", "R"
       SetGrid
       cmdBtn_Click (6)
    End If
    '---- Button Visiblity Set As Per User Previlage
    GProcSetButtonVisiblity Me
'    cmdBtn_Click (6)
End Sub
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
'If gFormAction = vbDataActionUpdate Then
'    MsgBox "You are about to modify the record. Save the record and then exit.", vbInformation + vbOKOnly, Me.Caption
'    Cancel = True
'    Exit Sub
'End If
rstbl.Close
End Sub
Private Sub dtpDt_LostFocus(Index As Integer)
SaveBtnEd
'If Index = 2 Then ShowAccountBalance
End Sub
Private Sub dtpDt_GotFocus(Index As Integer)
GProcSelectBox Me.ActiveControl
End Sub

Private Sub txt_GotFocus(Index As Integer)
GProcSelectBox Me.ActiveControl
End Sub
Private Sub txt_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
Dim OldFrmType As String
Dim LeftPos As Long
Dim TopPos As Long
If KeyCode = vbKeyF3 Then '--F3 Create
    Select Case Index
        Case 13, 14, 15, 16 '--Narration
            OldFrmType = gfrmTypeStr
            gfrmTypeStr = "N"
            GProcShowForm frmMastNarration, frmMain.mnuMstfrm(5), , gfrmTypeStr
            gfrmTypeStr = OldFrmType
    End Select
End If
If KeyCode = vbKeyF2 Then '-- Select
    TopPos = txt(Index).Top + Me.Top + 650
    LeftPos = txt(Index).Left + Me.Left
    Select Case Index
        Case 13, 14, 15, 16 '--Narration
            gClsSearch.SearchMultiField "tblMastNarration", "Narration", Array("Narration"), Array(txt(Index).Width), "NarrType='N'", "", "Narration", LeftPos, TopPos, False
            If gClsSearch.SearchMultiRetCol(0) <> "" Then txt(Index) = gClsSearch.SearchMultiRetCol(0)
    End Select
End If
End Sub
Private Sub txt_LostFocus(Index As Integer)
GProcCheckForNumber Me, rstbl, MaxNo
FillTxtFromTemp
SaveBtnEd
If Index = 16 And cmdBtn(13).Enabled = True Then
   cmdBtn(13).SetFocus
End If
If cboCrDr.ListIndex = 0 Then
  lblSecondAc.Caption = "Debit To A/c :"
Else
  lblSecondAc.Caption = "Credit To A/c :"
End If
End Sub
Private Sub txt_KeyPress(Index As Integer, KeyAscii As Integer)
    KeyAscii = GProcValidateKey(rstbl, Index, KeyAscii, txt(Index).text)
End Sub
Private Sub SaveBtnEd()
'--Vno,Vdt,Amount,Cr A/c,Dr A/c
If txt(0) = "" Or GProcIsDateValid(dtpDt(2)) = False Or txttemp(7) = "" Then
    cmdBtn(13).Enabled = False
    cmdBtn(13).BackColor = &H80000004
Else
    cmdBtn(13).Enabled = True
    cmdBtn(13).BackColor = &HFFFF80
End If
End Sub
Private Sub FillTxtFromTemp()
txt(1) = frmTypeStr
If IsDate(dtpDt(2)) Then '--Vdt
    txt(2) = CDate(dtpDt(2))
Else
    txt(2) = ""
End If
txt(3) = gCYear
txt(4) = gCCode
txt(26) = cboCrDr.ListIndex
txt(45) = cboNature.ListIndex

If cboCrDr.ListIndex = 0 Then
   txt(6) = GProcGetColumnValue("tblMastAccount", "AcName", txttemp(7), "S", "Accode", "N")
Else
   txt(7) = GProcGetColumnValue("tblMastAccount", "AcName", txttemp(7), "S", "Accode", "N")
End If
If txttempTDSAc = "" Then txttempTDSAc = "-"
If Val(txt(26)) = 0 Then
   txt(60) = GProcGetColumnValue("tblMastAccount", "AcName", txttempTDSAc, "S", "Accode", "N")
Else
   txt(59) = GProcGetColumnValue("tblMastAccount", "AcName", txttempTDSAc, "S", "Accode", "N")
End If
ShowAccountBalance
End Sub
Private Sub FillTempFromTxt()
dtpDt(2) = IIf(IsDate(txt(2)), CDate(txt(2)), "__/__/____") '--Vdt
If txt(26) = 0 Then ' -- Cr
   txttemp(7) = GProcGetColumnValue("tblMastAccount", "AcCode", txt(6), "N", "AcName", "S")
Else
   txttemp(7) = GProcGetColumnValue("tblMastAccount", "AcCode", txt(7), "N", "AcName", "S")
End If
txtCode(7) = GProcGetColumnValue("tblMastAccount", "AcName", txttemp(7), "S", "AcOurCode", "S")

If Val(txt(26)) = 0 Then
   txttempTDSAc = GProcGetColumnValue("tblMastAccount", "AcCode", txt(60), "N", "AcName", "S")
   txtCodeTDSAc = GProcGetColumnValue("tblMastAccount", "AcName", txttempTDSAc, "S", "AcOurCode", "S")
Else
   txttempTDSAc = GProcGetColumnValue("tblMastAccount", "AcCode", txt(59), "N", "AcName", "S")
   txtCodeTDSAc = GProcGetColumnValue("tblMastAccount", "AcName", txttempTDSAc, "S", "AcOurCode", "S")
End If

cboCrDr.ListIndex = Val(txt(26))
cboNature.ListIndex = Val(txt(45))

If cboCrDr.ListIndex = 0 Then
   lblSecondAc.Caption = "Debit To A/c :"
Else
   lblSecondAc.Caption = "Credit To A/c :"
End If
'-- Fill Related Record Rec Vs Sale
FillRelatedRecord
'ShowAccountBalance
End Sub
Private Function ValidateGrid() As Boolean
Dim TotAmt As Double
Dim i As Long
Dim RCount As Long
Dim j As Long
Dim ChkAcName As String
If Trim(txtGrid.text) <> "" Then
   msGrid.text = txtGrid.text
End If
txtGrid.Visible = False
txtGrid.text = ""
ChkAcName = ""
CheckForNumberGrid
'With msGrid
'    '-- Duplicate invoice selection
'    For i = 1 To .rows - 1
'        ChkAcName = .TextMatrix(i, 1)
'        For J = 1 To .rows - 1
'            If J <> i And ChkAcName = .TextMatrix(J, 1) Then
'                MsgBox "Duplicate A/c Selection.", vbOKOnly, Me.Caption
'                msGrid.SetFocus
'                Exit Function
'            End If
'        Next J
'    Next i
'End With
'ValidateGrid = True
ReNumberGridSrno
End Function
Private Function ValidateData() As Boolean
Dim X As String
Dim GName As String
Dim BnkTot As Double
Dim RowIndex As Integer
Dim rstAddCheck As Recordset
Dim Dub As Boolean
If FormAction <> vbDataActionAddNew Then
   If Len(txt(0)) = 5 Then
      i = Val(Left(txt(0), 1))
   Else
      i = Val(Left(txt(0), 2))
   End If
   If i <> Month(dtpDt(2)) Then
      MsgBox "Cant Change Month Of Voucher Date", vbInformation + vbOKOnly, Me.Caption
      dtpDt(2).SetFocus
      Exit Function
   End If
End If

GProcCheckForNumber Me, rstbl, MaxNo
RowIndex = 1
Dub = False
With msGrid
     Do While RowIndex < .rows
        If .TextMatrix(RowIndex, 1) = txttemp(7).text Then
           Dub = True
        End If
        RowIndex = RowIndex + 1
     Loop
End With
If Dub = True Then
   MsgBox "Credit A/c Name And Debit A/c are Same Please Check", vbInformation + vbOKOnly, Me.Caption
   txttemp(7).SetFocus
   Exit Function
End If
'---- Required Data
'If ValidateGrid = False Then Exit Function
'--- Vno
If CLng(txt(0)) = 0 And FormAction <> vbDataActionAddNew Then
    MsgBox "Check Voucher No.", vbInformation + vbOKOnly, Me.Caption
    txt(0).SetFocus
    Exit Function
End If
'If txt(55) = "" And FormAction <> vbDataActionAddNew And gfrmTypeStr <> "JV" Then
'    MsgBox "Check  No.", vbInformation + vbOKOnly, Me.Caption
'    txt(55).SetFocus
'    Exit Function
'End If

'--- Vdt
If GProcIsDateValid(dtpDt(2)) = False Then
    MsgBox "Check Voucher Date.", vbInformation + vbOKOnly, Me.Caption
    dtpDt(2).SetFocus
    Exit Function
End If
'--- Cr A/c
If txttemp(7) = "" Then
    MsgBox "Check Account.", vbInformation + vbOKOnly, Me.Caption
    txttemp(7).SetFocus
    Exit Function
End If
'--Create Cr A/c
If GProcGetColumnValue("tblMastAccount", "AcName", txttemp(7), "S", "AcCode", "N") = 0 Then
   GName = GProcGetColumnValue("TblMastGroup", "AgCode", 90019, "N", "AgName", "S")
   GProcShowForm frmMastAccount, frmMain.mnuMstfrm(0), Array(frmMastAccount.txt(1), frmMastAccount.txttemp(2)), Array(txttemp(7), GName)
   Exit Function
End If

'---- Amount
If CDbl(txt(8)) <= 0 Then
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
Private Sub ReadFields()
GProcShowRecord Me, rstbl, MaxNo
FillTempFromTxt
End Sub
Private Sub txttemp_GotFocus(Index As Integer)
GProcSelectBox Me.ActiveControl
End Sub
Private Sub txttemp_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
Dim GName As String
If KeyCode = vbKeyF3 Then
    Select Case Index
        Case 7 '-- cr A/c
         If frmTypeStr <> "CR" Then '-- Bank Receipt
            GName = GProcGetColumnValue("TblMastGroup", "AgCode", 90019, "N", "AgName", "S")
            GProcShowForm frmMastAccount, frmMain.mnuMstfrm(0), Array(frmMastAccount.txttemp(2)), Array(GName)
         End If
        Case 6 '-- dr A/c
            GName = GProcGetColumnValue("TblMastGroup", "AgCode", 90017, "N", "AgName", "S")
            GProcShowForm frmMastAccount, frmMain.mnuMstfrm(0), Array(frmMastAccount.txttemp(2)), Array(GName)
    End Select
End If
End Sub
Private Sub txtTemp_KeyPress(Index As Integer, KeyAscii As Integer)
Dim LeftPos As Long
Dim TopPos As Long
TopPos = txttemp(Index).Top + Me.Top + 650
LeftPos = txttemp(Index).Left + Me.Left
If KeyAscii = 13 Then Exit Sub
'--Cash Receipt
'If (frmTypeStr = "CR") And (Index = 5) Then
'   KeyAscii = 0
'   txt(Index) = 2
'   txttemp(Index) = GProcGetColumnValue("tblmastaccount", "Accode", 2, "N", "Acname", "S")
'   Exit Sub
'End If
Select Case Index
    Case 7: '--A/c Name
        gClsSearch.SearchMultiFieldAccMaster Chr(KeyAscii), LeftPos, TopPos
        KeyAscii = 0
        If gClsSearch.SearchMultiRetColAccMaster(0) <> "" Then
            txttemp(Index) = gClsSearch.SearchMultiRetColAccMaster(0)
            txtCode(7) = GProcGetColumnValue("tblMastAccount", "AcName", txttemp(7), "S", "AcOurCode", "S")
        End If
    Case 6: '--Account
'        gClsSearch.SearchMultiField "tblMastAccount", "AcName,AcCity", Array("AcName", "City"), Array(4000, 2000), "", Chr(KeyAscii), "AcName", LeftPos, TopPos, True
        gClsSearch.SearchMultiFieldAccMaster Chr(KeyAscii), LeftPos, TopPos
        KeyAscii = 0
        If gClsSearch.SearchMultiRetColAccMaster(0) <> "" Then
            txttemp(Index) = gClsSearch.SearchMultiRetColAccMaster(0)
        End If
End Select
End Sub
Private Sub txtTemp_LostFocus(Index As Integer)
SaveBtnEd
'SetControlEd
'ShowAccountBalance
If Index = 7 Then
   mLastEntryBank = txttemp(7).text
   mLastEntryDate = dtpDt(2)
End If
'If Index = 41 Then
'   msGrid.Col = 1
'   msGrid_EnterCell
'End If
DoEvents
End Sub
'--- Related Record
Private Function IsRelatedRecord() As Boolean
'-- Audited
'If GProcISAudited(txt(0), txt(3), gCYear) = True Then
'    MsgBox "Record is Audited.You can not Modify/Delete this Record.", vbCritical + vbOKOnly, App.Title
'    IsRelatedRecord = True
'    Exit Function
'End If
If rstbl.Fields("VBillNo") <> "" Then
    MsgBox "J.V. Against Invoice You can not Modify/Delete this Record.", vbCritical + vbOKOnly, App.Title
    IsRelatedRecord = True
    Exit Function
End If
'--- Related Record Receipt
If rstbl.EOF And rstbl.BOF Then
Else
    If rstbl.Fields("VIsRefEntType") = 1 Then
        MsgBox "Related Record is Present.You can not Modify/Delete this Record.", vbCritical + vbOKOnly, App.Title
        IsRelatedRecord = True
        Exit Function
    End If
End If
End Function
Private Sub ShowAccountBalance()
Dim DrAcBal  As Double
Dim CrAcBal As Double
If IsDate(dtpDt(2)) Then
    DrAcBal = GProcGetAccountBalance(txttemp(6), dtpDt(2))
    CrAcBal = GProcGetAccountBalance(txttemp(7), dtpDt(2))
    txtDrBalance = IIf(DrAcBal > 0, DrAcBal & " Cr", Abs(DrAcBal) & " Dr")
    txtCrBalance = IIf(CrAcBal > 0, CrAcBal & " Cr", Abs(CrAcBal) & " Dr")
End If
End Sub
'---------------- Grid
Private Sub SetGrid()
Dim X As String, x1Rec As Double
Dim TopCd As String
Set rsRecVsSale = New Recordset
If Val(txt(26)) = 0 Then
   TopCd = "VDrAcCode"
Else
   TopCd = "VCrAcCode"
End If
'X = "select VCtrNo,AcName,VAmt,Vnar1,Vnar2,VBillNo,VBillVno,VBillType,VBillDate,VBillVYear,VBillAmt,LateDays,Interest,InterestRecd,JvNoInt,LessTDS,JvNoLessTDS,VnoList,DueOnDt," _
'& " Grace,IntRt,GraceExtra,IntFromDate,IntParty,TdsRt,TDSOn,DBNtInt,DbNtMonthly,VBillFirm " _

X = "select VCtrNo,AcOurCode,AcName,VAmt,TDSOn,LessTDS,Vnar1,Vnar2" _
& " from tblVoucher,tblMastAccount where " _
& " VType='" & frmTypeStr & "' and VNo=" & txt(0) & " and VFirm  = '" & gCCode & "' and  VYear=" & gCYear & " and " & TopCd & "  = tblMastAccount.AcCode  order by VCtrNo"
x1Rec = GProcRstOpen(rsRecVsSale, X, "R")
With msGrid
    .Clear
    .FormatString = ">Sr |<A/c Code |<Account Name                                   |>Amount                |>TDS On Amt.   |>Less TDS       |<Narration 1                      |<Narration 2                       "
    If FormAction = vbDataActionAddNew Then
        .rows = 2
        GridAddNew
    Else
        .rows = 2
    End If
End With
End Sub
Private Sub GridAddNew()
Dim i As Long
With msGrid
    For i = 0 To rsRecVsSale.Fields.Count - 1
        Select Case rsRecVsSale.Fields(i).Type
            Case adInteger, 2, 3 '-- Integer
                .TextMatrix(.Row, i) = 0
            Case adDecimal, 131, adNumeric, adDouble '-- Double
                .TextMatrix(.Row, i) = Format(0, GProcNumberFormat(rsRecVsSale.Fields(i).Precision, rsRecVsSale.Fields(i).NumericScale))
            Case adDate '-- Date
                 .TextMatrix(.Row, i) = ""
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
With msGrid
     If ((.TextMatrix(.Row, 2)) = "") And .Row = 1 And txtGrid.Visible = True Then
        msGrid.Col = 1
        msGrid.SetFocus
        DoEvents
        If txtGrid.Visible = True Then
           txtGrid.SetFocus
        End If
     End If
End With
End Sub
Private Sub MSGrid_LostFocus()
CalAmount
End Sub
Private Sub msGrid_EnterCell()
If FormAction <> vbDataActionClose Then
    CheckForNumberTxtGrid
    DoEvents
    txtGrid = ""
    DoEvents
    If msGrid.text <> "" Then txtGrid.text = msGrid.text
    '--- Lock TextBox
'0 - VCtrNo
'1 - Ac Our Coude
'2 - AcName
'3 - VAmt
'4 - VNar1
'5 - VNar2
    With msGrid
        Select Case .Col
            Case 0, 2
                If .Col = 0 Then
                   ReNumberGridSrno
                End If
                txtGrid.Locked = True
            Case 6, 7
                txtGrid.MaxLength = 50
                txtGrid.Locked = False
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
        If (Val(.TextMatrix(.Row, 3)) = 0) And .rows > 2 Then
            .RemoveItem (.Row)
            txtGrid.text = ""
            If cmdBtn(13).Enabled = True Then
                cmdBtn(13).SetFocus
            Else
                cmdBtn(14).SetFocus
            End If
        End If
        txtGrid.Visible = False
        Exit Sub
    End If
    '---- Navigate Row ,Col
    If KeyAscii = 13 Then
        KeyAscii = 0
        If .Col = 7 Then
            If .Row = (.rows - 1) Then '-- Add New Row
                .rows = .rows + 1
                .Row = .Row + 1
            Else '--Next Row
                .Row = .Row + 1
            End If
            ReNumberGridSrno
            .Col = 1
        Else    '-- Next col
            .Col = .Col + 1
        End If
    End If
End With
ReNumberGridSrno
CalAmount
'---- Set TxtGrid
If Trim(msGrid.text) <> "" Then
    txtGrid = Trim(msGrid.text)
End If
If msGrid.Col < (rsRecVsSale.Fields.Count) Then  '- -3
    txtGrid.Move msGrid.CellLeft + msGrid.Left, msGrid.CellTop + msGrid.Top, msGrid.CellWidth, msGrid.CellHeight
End If
txtGrid.Visible = True
txtGrid.Enabled = True
DoEvents
If KeyAscii <> 0 Then
    Select Case rsRecVsSale.Fields(msGrid.Col).Type
        Case adChar, adVarWChar
            txtGrid.text = txtGrid.text & Chr(KeyAscii)
        Case Else
            txtGrid.text = Chr(GProcValidateKey(rsRecVsSale, msGrid.Col, KeyAscii, txtGrid.text)) 'Chr(GProcNumberWithDecimal(KeyAscii, txtGrid))
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
Select Case msGrid.Col
Case 6, 7:
    With msGrid
         .TextMatrix(.Row, msGrid.Col) = GProcProperCase(.TextMatrix(.Row, msGrid.Col))
    End With
Case 1:
    With msGrid
       .TextMatrix(.Row, 2) = GProcGetColumnValue("tblMastAccount", "AcOurCode", "" & .TextMatrix(.Row, 1) & "", "S", "AcName", "S")
    End With
End Select
CalGridAmt
'----Late Payment Interest Calcualtion
'---Interest Rate 1.5
'Select Case msGrid.Col
'Case 4:
'    With msGrid
'        .TextMatrix(.Row, 8) = Format(Round((((Val(.TextMatrix(.Row, 4)) * 1.5) / 3000) * Val(.TextMatrix(.Row, 7))), 2), "###0.00")
'    End With
'   ' CalBalance
'Case 6, 9:
'   ' CalBalance
'End Select
End Sub
Private Sub CalBalance()
Dim k As Integer
Dim TAmt As Double
With msGrid
    k = 1
    DoEvents
    Do While k < .rows
       TAmt = TAmt + Val(.TextMatrix(k, 3))
       k = k + 1
    Loop
    txt(8) = Format(TAmt, FStr)
'
'    DoEvents
'    If Val(lblBalance) > Val(txt(6)) Then MsgBox "Amount Should be less than or equal to " & txt(6), vbCritical + vbOKOnly, "Amount not Tally"
End With
End Sub
Private Sub txtGrid_GotFocus()
    txtGrid.Alignment = ProcSetAlignment(rsRecVsSale, msGrid.Col)
    DoEvents
    Call GProcSelectBox(txtGrid)
'    If msGrid.Col = 1 Then
'       Pending_Invoice
'    End If
End Sub
Private Sub txtGrid_KeyPress(KeyAscii As Integer)
Dim LeftPos As Long
Dim TopPos As Long
Dim PartyCode As Long
TopPos = txtGrid.Top + Me.Top + 650
LeftPos = txtGrid.Left + Me.Left
DoEvents
With msGrid
    If KeyAscii <> 13 And KeyAscii <> 27 Then
        '--- Search
        Select Case .Col
            Case 2 '-- A/c Name
'                 gClsSearch.SearchMultiField "QryMastAccount", "Account,AcOurCode,CGName,AcContPer,AcCity,AgName", Array("Account Name", "Code", "Party Group", "Contact Person", "City", "Group"), Array(4000, 1000, 2000, 2000, 2000, 2500), "", "", "Account"
                 gClsSearch.SearchMultiFieldAccMaster Chr(KeyAscii), LeftPos, TopPos
                 txtGrid = ""
                 KeyAscii = 0
                 If gClsSearch.SearchMultiRetColAccMaster(0) <> "" Then '-- Item Wt ,Rate From Master
                    txtGrid = gClsSearch.SearchMultiRetColAccMaster(0)
                    .TextMatrix(.Row, 1) = GProcGetColumnValue("tblMastAccount", "AcName", "" & gClsSearch.SearchMultiRetColAccMaster(0) & "", "S", "AcOurCode", "S")
                    txtGrid_KeyPress 13
                 End If
                DoEvents
                KeyAscii = 0
            Case Else
                KeyAscii = GProcValidateKey(rsRecVsSale, msGrid.Col, KeyAscii, txtGrid)
                If txtGrid.MaxLength > 0 And Len(txtGrid) > txtGrid.MaxLength - 1 Then KeyAscii = 0
                
        End Select
   End If
    If KeyAscii = 13 Then
        CheckForNumberTxtGrid
    End If
    '--- Navigate Row ,Col
    If KeyAscii = 13 And .Col < 7 Then
       .Col = .Col + 1
    ElseIf KeyAscii = 13 And .Col = 7 Then
        If .Row = (.rows - 1) Then
            .rows = .rows + 1
            .Row = .Row + 1
        Else
            .Row = .Row + 1
        End If
        .Col = 1
    End If
    '-- On ESC
    With msGrid
        If KeyAscii = 27 Then '----Esc Key
            KeyAscii = 0
            If (Val(.TextMatrix(.Row, 3)) = 0) And .rows > 2 Then
                .RemoveItem (.Row)
                txtGrid.text = ""
            End If
            txtGrid.Visible = False
            If cmdBtn(13).Enabled = True Then
                cmdBtn(13).SetFocus
            Else
                cmdBtn(14).SetFocus
            End If
            Exit Sub
        End If
    End With
End With
End Sub
Private Sub txtGrid_KeyDown(KeyCode As Integer, Shift As Integer)
'--- Delete Row On Ctrl+Y
If KeyCode = vbKeyY And Shift = 2 And msGrid.rows > 2 Then
    msGrid = 1
    msGrid.RemoveItem (msGrid.Row)
    txtGrid.Visible = False
    txtGrid.text = ""
    CalAmount
    ReNumberGridSrno
    Exit Sub
End If
If KeyCode = vbKeyF3 Then
Select Case msGrid.Col
    Case 6, 7 '-- Party Bank
        OldgFrmTypeStr = gfrmTypeStr
        gfrmTypeStr = "N"
        GProcShowForm frmMastNarration, frmMain.mnuMstfrm(10), "N"
        gfrmTypeStr = OldgFrmTypeStr
End Select
End If


With msGrid
     If (.Col = 6 Or .Col = 7) And KeyCode = vbKeyF1 Then
        gClsSearch.SearchMultiField "tblMastNarration", "Narration", Array("Narration"), Array(6000), "NarrType='N'", "", "Narration", 3500, 4500, False
        If gClsSearch.SearchMultiRetCol(0) <> "" Then txtGrid = gClsSearch.SearchMultiRetCol(0)
     End If
End With



End Sub
Private Sub CheckForNumberGrid()
Dim RowIndex As Long
With rsRecVsSale
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
                    msGrid.TextMatrix(RowIndex, i) = ""
                End If
        End Select
    Next i
  Next RowIndex
End With
End Sub
'---  Check For Number Grid
Private Sub CheckForNumberTxtGrid()
With rsRecVsSale
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
                txtGrid = ""
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
Private Sub DeleteRelatedRecord()
Dim X As String
Dim x1Rec As Double
Dim rstTmp As Recordset
X = "select VCtrNo from tblVoucher where VType='" & frmTypeStr & "' and VNo=" & txt(0) & " and VFirm  = '" & gCCode & "' and  VYear=" & gCYear & "  order by VCtrNo"
x1Rec = GProcRstOpen(rstTmp, X, "R")
With rstTmp
     If Not .EOF Then
       .MoveFirst
       Do While Not .EOF
          gCn.Execute "Delete from tblvoucher where " & CriteriaStr & " and VCtrNo = " & rstTmp.Fields(0) & "   and  vno = " & txt(0)
          .MoveNext
       Loop
     End If
End With
End Sub
Private Sub SaveRelatedRecord()
Dim i As Long, X As String
Dim MacCode As Long
Dim RowIndex As Integer
RowIndex = 1
With msGrid
    Do While RowIndex < .rows
       If (Val(.TextMatrix(.Row, 3)) = 0) And .rows > 2 Then
          .RemoveItem (.Row)
       End If
       RowIndex = RowIndex + 1
    Loop
End With
ReNumberGridSrno
'-- Delete Records
DeleteRelatedRecord
'--- Save Grid Records
'0 - VCtrNo
'1 - AcName
'2 - Vamt
MacCode = GProcGetColumnValue("tblMastAccount", "AcName", txttemp(7), "S", "AcCode", "N")
With msGrid
    i = 1
    Do While i < .rows
        DoEvents
        If CDec(Val(.TextMatrix(i, 3))) > 0 Then
            rstbl.AddNew
            rstbl!Vno = CLng(txt(0)) '--VNo
            rstbl!Vtype = frmTypeStr  '--Vtype
            rstbl!Vdt = CDate(txt(2)) '--Vdt(2)
            rstbl!VYear = txt(3) '--VYear
            rstbl!VFirm = gCCode '--VFirm
            rstbl!VCtrNo = i '--Sr No
            If cboCrDr.ListIndex = 0 Then
               rstbl!VCrAcCode = MacCode
               rstbl!VDrAcCode = GProcGetColumnValue("tblMastAccount", "AcName", .TextMatrix(i, 2), "S", "AcCode", "N")
            
               rstbl!VCrAcCodeTDS = GProcGetColumnValue("tblMastAccount", "AcName", .TextMatrix(i, 2), "S", "AcCode", "N")
               rstbl!VDrAcCodeTDS = GProcGetColumnValue("tblMastAccount", "AcName", txttempTDSAc, "S", "AcCode", "N")
            
            Else
               rstbl!VDrAcCode = MacCode
               rstbl!VCrAcCode = GProcGetColumnValue("tblMastAccount", "AcName", .TextMatrix(i, 2), "S", "AcCode", "N")
            
               rstbl!VDrAcCodeTDS = GProcGetColumnValue("tblMastAccount", "AcName", .TextMatrix(i, 2), "S", "AcCode", "N")
               rstbl!VCrAcCodeTDS = GProcGetColumnValue("tblMastAccount", "AcName", txttempTDSAc, "S", "AcCode", "N")
            
            End If
            rstbl!vamt = CDec(.TextMatrix(i, 3))
            rstbl!TDSOn = CDec(.TextMatrix(i, 4))
            rstbl!LessTDS = CDec(.TextMatrix(i, 5))
            
            rstbl!VRefTp = txt(9) & vbNullString
            rstbl!VRefNo = txt(10) & vbNullString
            If IsDate(txt(11)) = True Then
                rstbl!VRefDate = CDate(txt(11))
            Else
                rstbl!VRefDate = Null
            End If
            rstbl!VRefBank = txt(12) & vbNullString
            rstbl!VNar1 = .TextMatrix(i, 6) & vbNullString
            rstbl!VNar2 = .TextMatrix(i, 7) & vbNullString
            rstbl!VNar3 = txt(15) & vbNullString
            rstbl!VNAR4 = txt(16) & vbNullString
            If IsDate(txt(17)) = True Then
                rstbl!VReconDt = CDate(txt(17))
            Else
                rstbl!VReconDt = Null
            End If
            rstbl!VIsRefEntType = Val(txt(18))
            rstbl!VIsAudited = Val(txt(19))
            rstbl!VBillNo = txt(20) & vbNullString
            rstbl!VBillVno = Val(txt(21))
            rstbl!VBillType = UCase(txt(22)) & vbNullString
            If IsDate(txt(23)) = True Then
               rstbl!VBillDate = CDate(txt(23))
            Else
               rstbl!VBillDate = Null
            End If
            rstbl!VBillVYear = Val(txt(24))
            rstbl!VBillAmt = Val(txt(25))
            rstbl!VTopCrDr = cboCrDr.ListIndex
            rstbl!Discount = Val(txt(27))
            rstbl!LateDays = Val(txt(28))
            rstbl!Interest = Val(txt(29))
            rstbl!InterestCredit = Val(txt(30))
            rstbl!JvNoDisc = Val(txt(31))
'            rstbl!JvNoInt = Val(txt(32))
'            rstbl!LessTds = Val(txt(33))
            rstbl!JvNoLessTDS = Val(txt(34))
            rstbl!VnoList = Val(txt(35))
            If IsDate(txt(36)) = True Then
                rstbl!DueOnDt = CDate(txt(36))
            Else
                rstbl!DueOnDt = Null
            End If
            rstbl!Grace = Val(txt(37))
            rstbl!IntRt = Val(txt(38))
            rstbl!GraceExtra = Val(txt(39))
            If IsDate(txt(40)) = True Then
               rstbl!IntFromDate = CDate(txt(40))
            Else
               rstbl!IntFromDate = Null
            End If
            rstbl!IntParty = Val(txt(41))
            rstbl!TdsRt = Val(txt(42))
            rstbl!DBNtInt = txt(44) & vbNullString
            rstbl!DbNtMonthly = Val(txt(45))
            rstbl!VBillFirm = txt(46) & vbNullString
            rstbl!DbNtInBillNo = txt(47) & vbNullString
            rstbl!JvNoIntInDepot = Val(txt(48))
            rstbl!CgstRt = Val(txt(49))
            rstbl!CgstAmt = Val(txt(50))
            rstbl!SgstRt = Val(txt(51))
            rstbl!SgstAmt = Val(txt(52))
            rstbl!IgstRt = Val(txt(53))
            rstbl!IgstAmt = Val(txt(54))
            rstbl!CrDrNoteNo = txt(55) & vbNullString
            rstbl!VNoListVYear = 0
            rstbl.Update
            DoEvents
            i = i + 1
        End If
     Loop
End With
End Sub
Private Sub FillRelatedRecord()
Dim RowIndex As Long
Dim ColIndex As Long
Dim BnkTot As Double
'--Fill Grid
SetGrid
BnkTot = 0
With rsRecVsSale
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
                If ColIndex = 3 Then BnkTot = BnkTot + CLng(msGrid.TextMatrix(RowIndex, 3)) '-- Vamt
            Next ColIndex
            RowIndex = RowIndex + 1
            .MoveNext
            If Not .EOF Then msGrid.rows = msGrid.rows + 1
        Loop
    End If
End With
txt(8) = Format(BnkTot, "###0.00")
CalAmount
End Sub
Private Sub CalAmount()
Dim BnkTot As Double
Dim BnkTot1 As Double
Dim BnkTot2 As Double
Dim RowIndex As Integer
RowIndex = 1
With msGrid
    Do While RowIndex < .rows
        '-- sub amt
        If Val(.TextMatrix(RowIndex, 3)) > 0 Then
            BnkTot = BnkTot + Val(.TextMatrix(RowIndex, 3))
        End If
        If Val(.TextMatrix(RowIndex, 4)) > 0 Then
            BnkTot1 = BnkTot1 + Val(.TextMatrix(RowIndex, 4))
        End If
        If Val(.TextMatrix(RowIndex, 5)) > 0 Then
            BnkTot2 = BnkTot2 + Val(.TextMatrix(RowIndex, 5))
        End If
        RowIndex = RowIndex + 1
    Loop
End With
txt(8) = Format(BnkTot, "###0.00")
lbl1(12).Caption = Format(BnkTot, "###0.00")
lbl1(10).Caption = Format(BnkTot1, "###0.00")
lbl1(11).Caption = Format(BnkTot2, "###0.00")
End Sub
Private Sub CalGridAmt()
Dim Amt As Double
Dim Amt2 As Double

If msGrid.Row = 0 Then Exit Sub
'If Val(msGrid.TextMatrix(msGrid.Row, 3)) > 0 And Val(msGrid.TextMatrix(msGrid.Row, 5)) = 0 Then
'    msGrid.TextMatrix(msGrid.Row, 4) = msGrid.TextMatrix(msGrid.Row, 3)
'End If
'If Val(msGrid.TextMatrix(msGrid.Row, 4)) > 0 Then
'   msGrid.TextMatrix(msGrid.Row, 5) = Round((Val((msGrid.TextMatrix(msGrid.Row, 4))) * Val(txt(42))) / 100, 0)
'End If
If Val(msGrid.TextMatrix(msGrid.Row, 5)) = 0 Or Abs((Val(msGrid.TextMatrix(msGrid.Row, 5)) - Round((Val((msGrid.TextMatrix(msGrid.Row, 4))) * Val(txt(42))) / 100, 0))) > 1 Then
   Amt2 = (Val(msGrid.TextMatrix(msGrid.Row, 4)) * Val(txt(42))) / 100
   Amt = GProcMakeRounding(Amt2)
   msGrid.TextMatrix(msGrid.Row, 5) = Amt
'Else
'   msGrid.TextMatrix(msGrid.Row, 5) = 0
End If
End Sub
Private Sub SetControlEd()
txt(8).Enabled = False
lblCrDrNtNo.Visible = False
txt(55).Visible = False
End Sub
Private Sub ReNumberGridSrno()
Dim i As Long
With msGrid
    For i = 1 To .rows - 1
        .TextMatrix(i, 0) = i
    Next i
End With
End Sub

Private Sub txtCode_GotFocus(Index As Integer)
'If Index = 7 And txtCode(7) = "" Then
'   If gCIsDepotFirm = 1 Then
'      txtCode(7) = GProcGetColumnValue("tblMastAccount", "AcCode", "'" & gCDepotMainFirm & "' ", "N", "AcOurCode", "S")
'   End If
''   txtCode(7) = "ACAB009"
'   txtCode(7).SelStart = 8
'End If
End Sub
Private Sub txtCode_LostFocus(Index As Integer)
If Trim(txtCode(Index)) = "" Then Exit Sub
SetControlEd
Select Case Index
    Case 7  '-- Bank  Code
         txttemp(7) = GProcGetColumnValue("tblMastAccount", "AcOurCode", UCase(txtCode(7)), "S", "AcName", "S")
         If Trim(txttemp(7)) = "" Then
            MsgBox ("Code Not Found ...")
            txttemp(7) = ""
            txttemp(7).Enabled = True
            txttemp(7).SetFocus
'         Else
'            msGrid.Col = 1
'            msGrid_EnterCell
         End If
End Select
End Sub
'Private Sub CheckForNumber()
'GProcCheckForNumber Me, rstbl, MaxNo, Array(55)
'End Sub

Private Sub txttempTDSAc_KeyPress(KeyAscii As Integer)
Dim LeftPos As Long
Dim TopPos As Long
TopPos = txttempTDSAc.Top + Me.Top + 650
LeftPos = txttempTDSAc.Left + Me.Left
If KeyAscii = 13 Then Exit Sub
        
        gClsSearch.SearchMultiFieldAccMaster Chr(KeyAscii), LeftPos, TopPos
        KeyAscii = 0
        If gClsSearch.SearchMultiRetColAccMaster(0) <> "" Then
            txttempTDSAc = gClsSearch.SearchMultiRetColAccMaster(0)
            txtCodeTDSAc = GProcGetColumnValue("tblMastAccount", "AcName", txttempTDSAc, "S", "AcOurCode", "S")
        End If

End Sub
