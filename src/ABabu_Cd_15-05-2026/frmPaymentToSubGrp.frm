VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form frmPaymentToSubGroup 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00FBF2E1&
   BorderStyle     =   0  'None
   Caption         =   "Receipt Entry"
   ClientHeight    =   9405
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   17955
   Icon            =   "frmPaymentToSubGrp.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   9405
   ScaleWidth      =   17955
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox txt 
      Height          =   325
      Index           =   62
      Left            =   15420
      TabIndex        =   119
      Text            =   "62"
      Top             =   3660
      Visible         =   0   'False
      Width           =   360
   End
   Begin VB.TextBox txt 
      Height          =   325
      Index           =   61
      Left            =   15480
      TabIndex        =   118
      Text            =   "61"
      Top             =   3360
      Visible         =   0   'False
      Width           =   360
   End
   Begin VB.TextBox txt 
      Height          =   325
      Index           =   60
      Left            =   15540
      TabIndex        =   117
      Text            =   "60"
      Top             =   3000
      Visible         =   0   'False
      Width           =   360
   End
   Begin VB.TextBox txt 
      Height          =   325
      Index           =   59
      Left            =   15420
      TabIndex        =   116
      Text            =   "59"
      Top             =   2640
      Visible         =   0   'False
      Width           =   360
   End
   Begin VB.TextBox txt 
      Height          =   325
      Index           =   58
      Left            =   15540
      TabIndex        =   115
      Text            =   "58"
      Top             =   2340
      Visible         =   0   'False
      Width           =   360
   End
   Begin VB.TextBox txt 
      Height          =   325
      Index           =   57
      Left            =   14940
      TabIndex        =   114
      Text            =   "57"
      Top             =   4080
      Visible         =   0   'False
      Width           =   360
   End
   Begin VB.TextBox txt 
      Height          =   325
      Index           =   56
      Left            =   14880
      TabIndex        =   113
      Text            =   "56"
      Top             =   3600
      Visible         =   0   'False
      Width           =   360
   End
   Begin VB.TextBox txt 
      Height          =   325
      Index           =   55
      Left            =   14880
      TabIndex        =   112
      Text            =   "55"
      Top             =   3300
      Visible         =   0   'False
      Width           =   360
   End
   Begin VB.TextBox txt 
      Height          =   325
      Index           =   54
      Left            =   14940
      TabIndex        =   111
      Text            =   "54"
      Top             =   3060
      Visible         =   0   'False
      Width           =   360
   End
   Begin VB.TextBox txt 
      Height          =   325
      Index           =   53
      Left            =   14760
      TabIndex        =   110
      Text            =   "53"
      Top             =   2760
      Visible         =   0   'False
      Width           =   360
   End
   Begin VB.TextBox txt 
      Height          =   325
      Index           =   52
      Left            =   14880
      TabIndex        =   109
      Text            =   "52"
      Top             =   2460
      Visible         =   0   'False
      Width           =   360
   End
   Begin VB.TextBox txt 
      Height          =   325
      Index           =   51
      Left            =   14760
      TabIndex        =   108
      Text            =   "51"
      Top             =   1980
      Visible         =   0   'False
      Width           =   360
   End
   Begin VB.TextBox txt 
      Height          =   325
      Index           =   50
      Left            =   14820
      TabIndex        =   107
      Text            =   "50"
      Top             =   1620
      Visible         =   0   'False
      Width           =   360
   End
   Begin VB.TextBox txt 
      Height          =   325
      Index           =   49
      Left            =   14820
      TabIndex        =   106
      Text            =   "49"
      Top             =   1260
      Visible         =   0   'False
      Width           =   360
   End
   Begin VB.TextBox txt 
      Height          =   325
      Index           =   48
      Left            =   14880
      TabIndex        =   105
      Text            =   "48"
      Top             =   960
      Visible         =   0   'False
      Width           =   360
   End
   Begin VB.TextBox txt 
      Height          =   325
      Index           =   47
      Left            =   14760
      TabIndex        =   104
      Text            =   "47"
      Top             =   600
      Visible         =   0   'False
      Width           =   360
   End
   Begin VB.TextBox txtDebit 
      Alignment       =   1  'Right Justify
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
      Height          =   390
      Left            =   8640
      TabIndex        =   99
      Text            =   "txtDebit"
      Top             =   960
      Width           =   2280
   End
   Begin VB.TextBox txtOpBalance 
      Alignment       =   1  'Right Justify
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
      Left            =   8640
      TabIndex        =   98
      Text            =   "txtOpBalance"
      Top             =   585
      Width           =   2280
   End
   Begin VB.TextBox txtCredit 
      Alignment       =   1  'Right Justify
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
      Height          =   390
      Left            =   8640
      TabIndex        =   97
      Text            =   "txtCredit"
      Top             =   1365
      Width           =   2280
   End
   Begin VB.TextBox txtClBalance 
      Alignment       =   1  'Right Justify
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
      Height          =   390
      Left            =   8640
      TabIndex        =   96
      Text            =   "txtClBalance"
      Top             =   1755
      Width           =   2280
   End
   Begin VB.TextBox txt 
      Height          =   325
      Index           =   46
      Left            =   16065
      TabIndex        =   95
      Text            =   "46"
      Top             =   7875
      Visible         =   0   'False
      Width           =   360
   End
   Begin VB.TextBox txt 
      Height          =   325
      Index           =   45
      Left            =   16110
      TabIndex        =   94
      Text            =   "45"
      Top             =   7470
      Visible         =   0   'False
      Width           =   360
   End
   Begin VB.TextBox txt 
      Height          =   325
      Index           =   44
      Left            =   16065
      TabIndex        =   93
      Text            =   "44"
      Top             =   6975
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
      Left            =   15570
      TabIndex        =   8
      Text            =   "txtCode(7)"
      ToolTipText     =   "Press F3 for Creation"
      Top             =   1260
      Visible         =   0   'False
      Width           =   420
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
      Index           =   6
      Left            =   1530
      TabIndex        =   4
      Text            =   "txtCode(6)"
      ToolTipText     =   "Press F3 for Creation"
      Top             =   1222
      Width           =   1410
   End
   Begin VB.TextBox txt 
      Height          =   325
      Index           =   43
      Left            =   16920
      TabIndex        =   92
      Text            =   "43"
      Top             =   8145
      Visible         =   0   'False
      Width           =   360
   End
   Begin VB.TextBox txt 
      Height          =   325
      Index           =   42
      Left            =   16965
      TabIndex        =   91
      Text            =   "42"
      Top             =   7740
      Visible         =   0   'False
      Width           =   360
   End
   Begin VB.TextBox txt 
      Height          =   325
      Index           =   41
      Left            =   17010
      TabIndex        =   90
      Text            =   "41"
      Top             =   7380
      Visible         =   0   'False
      Width           =   360
   End
   Begin VB.TextBox txt 
      Height          =   325
      Index           =   40
      Left            =   16965
      TabIndex        =   89
      Text            =   "40"
      Top             =   7065
      Visible         =   0   'False
      Width           =   360
   End
   Begin VB.TextBox txt 
      Height          =   325
      Index           =   39
      Left            =   16965
      TabIndex        =   88
      Text            =   "39"
      Top             =   6750
      Visible         =   0   'False
      Width           =   360
   End
   Begin VB.TextBox txt 
      Height          =   325
      Index           =   38
      Left            =   16965
      TabIndex        =   87
      Text            =   "38"
      Top             =   6390
      Visible         =   0   'False
      Width           =   360
   End
   Begin VB.TextBox txt 
      Height          =   325
      Index           =   37
      Left            =   16965
      TabIndex        =   86
      Text            =   "37"
      Top             =   6030
      Visible         =   0   'False
      Width           =   360
   End
   Begin VB.TextBox txt 
      Height          =   325
      Index           =   35
      Left            =   16965
      TabIndex        =   85
      Text            =   "35"
      Top             =   5265
      Visible         =   0   'False
      Width           =   360
   End
   Begin VB.TextBox txt 
      Height          =   325
      Index           =   21
      Left            =   16965
      TabIndex        =   84
      Text            =   "21"
      Top             =   315
      Visible         =   0   'False
      Width           =   360
   End
   Begin VB.TextBox txt 
      Height          =   325
      Index           =   22
      Left            =   16965
      TabIndex        =   83
      Text            =   "22"
      Top             =   675
      Visible         =   0   'False
      Width           =   360
   End
   Begin VB.TextBox txt 
      Height          =   325
      Index           =   23
      Left            =   16965
      TabIndex        =   82
      Text            =   "23"
      Top             =   1035
      Visible         =   0   'False
      Width           =   360
   End
   Begin VB.TextBox txt 
      Height          =   325
      Index           =   24
      Left            =   16965
      TabIndex        =   81
      Text            =   "24"
      Top             =   1395
      Visible         =   0   'False
      Width           =   360
   End
   Begin VB.TextBox txt 
      Height          =   325
      Index           =   25
      Left            =   16965
      TabIndex        =   80
      Text            =   "25"
      Top             =   1755
      Visible         =   0   'False
      Width           =   360
   End
   Begin VB.TextBox txt 
      Height          =   325
      Index           =   26
      Left            =   16965
      TabIndex        =   79
      Text            =   "26"
      Top             =   2070
      Visible         =   0   'False
      Width           =   360
   End
   Begin VB.TextBox txt 
      Height          =   325
      Index           =   27
      Left            =   17010
      TabIndex        =   78
      Text            =   "27"
      Top             =   2430
      Visible         =   0   'False
      Width           =   360
   End
   Begin VB.TextBox txt 
      Height          =   325
      Index           =   28
      Left            =   17010
      TabIndex        =   77
      Text            =   "28"
      Top             =   2790
      Visible         =   0   'False
      Width           =   360
   End
   Begin VB.TextBox txt 
      Height          =   325
      Index           =   29
      Left            =   16920
      TabIndex        =   76
      Text            =   "29"
      Top             =   3105
      Visible         =   0   'False
      Width           =   360
   End
   Begin VB.TextBox txt 
      Height          =   325
      Index           =   30
      Left            =   16965
      TabIndex        =   75
      Text            =   "30"
      Top             =   3465
      Visible         =   0   'False
      Width           =   360
   End
   Begin VB.TextBox txt 
      Height          =   325
      Index           =   31
      Left            =   17010
      TabIndex        =   74
      Text            =   "31"
      Top             =   3825
      Visible         =   0   'False
      Width           =   360
   End
   Begin VB.TextBox txt 
      Height          =   325
      Index           =   32
      Left            =   17010
      TabIndex        =   73
      Text            =   "32"
      Top             =   4185
      Visible         =   0   'False
      Width           =   360
   End
   Begin VB.TextBox txt 
      Height          =   325
      Index           =   33
      Left            =   17010
      TabIndex        =   72
      Text            =   "33"
      Top             =   4545
      Visible         =   0   'False
      Width           =   360
   End
   Begin VB.TextBox txt 
      Height          =   325
      Index           =   34
      Left            =   17010
      TabIndex        =   71
      Text            =   "34"
      Top             =   4905
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
      Index           =   36
      Left            =   16965
      TabIndex        =   70
      Text            =   "36"
      Top             =   5625
      Visible         =   0   'False
      Width           =   405
   End
   Begin VB.TextBox txt 
      Height          =   325
      Index           =   20
      Left            =   18270
      TabIndex        =   69
      Text            =   "20"
      Top             =   4770
      Visible         =   0   'False
      Width           =   360
   End
   Begin VB.TextBox txt 
      Height          =   325
      Index           =   19
      Left            =   18225
      TabIndex        =   68
      Text            =   "19"
      Top             =   4410
      Visible         =   0   'False
      Width           =   360
   End
   Begin VB.TextBox txt 
      Height          =   325
      Index           =   18
      Left            =   18225
      TabIndex        =   67
      Text            =   "18"
      Top             =   4050
      Visible         =   0   'False
      Width           =   360
   End
   Begin VB.TextBox txt 
      Height          =   325
      Index           =   17
      Left            =   18180
      TabIndex        =   66
      Text            =   "17"
      Top             =   3735
      Visible         =   0   'False
      Width           =   360
   End
   Begin VB.TextBox txt 
      Height          =   325
      Index           =   12
      Left            =   18135
      TabIndex        =   65
      Text            =   "12"
      Top             =   3375
      Visible         =   0   'False
      Width           =   360
   End
   Begin VB.TextBox txt 
      Height          =   325
      Index           =   11
      Left            =   18135
      TabIndex        =   64
      Text            =   "11"
      Top             =   3015
      Visible         =   0   'False
      Width           =   360
   End
   Begin VB.TextBox txt 
      Height          =   325
      Index           =   9
      Left            =   18180
      TabIndex        =   63
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
      TabIndex        =   62
      Text            =   "6"
      Top             =   1935
      Visible         =   0   'False
      Width           =   345
   End
   Begin VB.ListBox List1 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Left            =   16425
      TabIndex        =   13
      Top             =   3825
      Visible         =   0   'False
      Width           =   435
   End
   Begin VB.TextBox txt 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   375
      Index           =   16
      Left            =   16335
      MaxLength       =   50
      TabIndex        =   19
      Text            =   "16"
      ToolTipText     =   "Press F3 for Creation & Press F2 for Selection"
      Top             =   3375
      Visible         =   0   'False
      Width           =   540
   End
   Begin VB.TextBox txt 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   375
      Index           =   15
      Left            =   16335
      MaxLength       =   50
      TabIndex        =   18
      Text            =   "15"
      ToolTipText     =   "Press F3 for Creation & Press F2 for Selection"
      Top             =   2970
      Visible         =   0   'False
      Width           =   495
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
      TabIndex        =   61
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
      Left            =   3825
      TabIndex        =   52
      Top             =   3555
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
         TabIndex        =   60
         Top             =   2160
         Width           =   1335
      End
      Begin VB.TextBox txtno 
         Appearance      =   0  'Flat
         Height          =   360
         Index           =   0
         Left            =   1950
         MaxLength       =   7
         TabIndex        =   57
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
         TabIndex        =   56
         Top             =   1560
         Width           =   1335
      End
      Begin VB.TextBox txtno 
         Appearance      =   0  'Flat
         Height          =   360
         Index           =   1
         Left            =   1950
         MaxLength       =   7
         TabIndex        =   55
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
         TabIndex        =   54
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
         TabIndex        =   53
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
         TabIndex        =   59
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
         TabIndex        =   58
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
      Left            =   13320
      Style           =   1  'Graphical
      TabIndex        =   51
      ToolTipText     =   "Exit"
      Top             =   1665
      Width           =   795
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
      Left            =   12375
      Style           =   1  'Graphical
      TabIndex        =   50
      ToolTipText     =   "To Cancel Record"
      Top             =   1665
      Width           =   795
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
      Left            =   12375
      Style           =   1  'Graphical
      TabIndex        =   49
      ToolTipText     =   "To Save Record"
      Top             =   1170
      Width           =   795
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
      Left            =   17505
      Style           =   1  'Graphical
      TabIndex        =   48
      Top             =   6795
      Visible         =   0   'False
      Width           =   795
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
      Left            =   11385
      Style           =   1  'Graphical
      TabIndex        =   47
      ToolTipText     =   "To Delete Record"
      Top             =   1665
      Width           =   795
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
      Left            =   11385
      Style           =   1  'Graphical
      TabIndex        =   46
      ToolTipText     =   "To Modify Record"
      Top             =   1170
      Width           =   795
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
      Left            =   11385
      Style           =   1  'Graphical
      TabIndex        =   45
      ToolTipText     =   "To Add New Record"
      Top             =   630
      Width           =   795
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
      Left            =   12375
      Style           =   1  'Graphical
      TabIndex        =   44
      ToolTipText     =   "To Find Record"
      Top             =   630
      Width           =   795
   End
   Begin VB.CommandButton cmdIntDbNt 
      Caption         =   "&Interest Debit Note"
      Height          =   945
      Left            =   17595
      TabIndex        =   43
      Top             =   5805
      Visible         =   0   'False
      Width           =   795
   End
   Begin VB.TextBox txt 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Index           =   10
      Left            =   16110
      TabIndex        =   12
      Text            =   "10"
      Top             =   5265
      Visible         =   0   'False
      Width           =   315
   End
   Begin VB.TextBox txttemp1 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Index           =   9
      Left            =   15705
      TabIndex        =   11
      Text            =   "txttemp1(9)"
      Top             =   5265
      Visible         =   0   'False
      Width           =   330
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
      TabIndex        =   41
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
      TabIndex        =   40
      Text            =   "txtDrBalance"
      Top             =   585
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.CommandButton cmdLedger 
      Caption         =   "Ledger"
      Height          =   375
      Index           =   0
      Left            =   18225
      TabIndex        =   39
      Top             =   5175
      Visible         =   0   'False
      Width           =   1200
   End
   Begin VB.CommandButton cmdLedger 
      Caption         =   "Ledger"
      Height          =   375
      Index           =   1
      Left            =   18225
      TabIndex        =   38
      Top             =   5535
      Visible         =   0   'False
      Width           =   1185
   End
   Begin VB.TextBox txtGrid 
      BackColor       =   &H8000000A&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   1755
      TabIndex        =   7
      Top             =   3240
      Visible         =   0   'False
      Width           =   885
   End
   Begin VB.TextBox txttemp1 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   375
      Index           =   12
      Left            =   15570
      MaxLength       =   40
      TabIndex        =   15
      Text            =   "txttemp1(12)"
      ToolTipText     =   "Press F3 for Creation & Press F2 for Selection"
      Top             =   5715
      Visible         =   0   'False
      Width           =   990
   End
   Begin VB.TextBox txt 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   375
      Index           =   13
      Left            =   16335
      MaxLength       =   50
      TabIndex        =   16
      Text            =   "13"
      ToolTipText     =   "Press F3 for Creation & Press F2 for Selection"
      Top             =   2160
      Visible         =   0   'False
      Width           =   540
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
      Width           =   1320
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
      Index           =   6
      Left            =   360
      MaxLength       =   60
      TabIndex        =   5
      Text            =   "txttemp(6)"
      ToolTipText     =   "Press F3 for Creation"
      Top             =   1665
      Width           =   6705
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
      Left            =   15615
      MaxLength       =   60
      TabIndex        =   9
      Text            =   "txttemp(7)"
      ToolTipText     =   "Press F3 for Creation"
      Top             =   1710
      Visible         =   0   'False
      Width           =   1350
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
      Left            =   13320
      Style           =   1  'Graphical
      TabIndex        =   23
      ToolTipText     =   "Previous"
      Top             =   630
      Width           =   360
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
      Left            =   13755
      Style           =   1  'Graphical
      TabIndex        =   24
      ToolTipText     =   "Next"
      Top             =   630
      Width           =   360
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
      Left            =   13320
      Style           =   1  'Graphical
      TabIndex        =   25
      ToolTipText     =   "First"
      Top             =   1125
      Width           =   360
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
      Left            =   13755
      Style           =   1  'Graphical
      TabIndex        =   26
      ToolTipText     =   "Last"
      Top             =   1125
      Width           =   360
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
      TabIndex        =   27
      Top             =   6480
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
      Left            =   18585
      Style           =   1  'Graphical
      TabIndex        =   31
      Top             =   8325
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
      Left            =   18585
      Style           =   1  'Graphical
      TabIndex        =   32
      Top             =   7920
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.TextBox txt 
      Height          =   345
      Index           =   3
      Left            =   18135
      TabIndex        =   35
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
      TabIndex        =   30
      Text            =   "5"
      Top             =   1575
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
      Height          =   360
      Index           =   8
      Left            =   6285
      TabIndex        =   10
      Text            =   "8"
      Top             =   8880
      Width           =   2805
   End
   Begin VB.TextBox txt 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   375
      Index           =   14
      Left            =   16335
      MaxLength       =   50
      TabIndex        =   17
      Text            =   "14"
      ToolTipText     =   "Press F3 for Creation & Press F2 for Selection"
      Top             =   2565
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.TextBox txt 
      Height          =   325
      Index           =   7
      Left            =   18180
      TabIndex        =   28
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
      TabIndex        =   29
      Text            =   "4"
      Top             =   1260
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.TextBox txt 
      Height          =   345
      Index           =   2
      Left            =   18135
      TabIndex        =   34
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
      TabIndex        =   33
      Text            =   "1"
      Top             =   180
      Visible         =   0   'False
      Width           =   375
   End
   Begin MSMask.MaskEdBox mskDt 
      Height          =   360
      Index           =   2
      Left            =   15030
      TabIndex        =   3
      Top             =   90
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
   Begin MSMask.MaskEdBox mskDt 
      Height          =   360
      Index           =   11
      Left            =   16020
      TabIndex        =   14
      Top             =   6165
      Visible         =   0   'False
      Width           =   375
      _ExtentX        =   661
      _ExtentY        =   635
      _Version        =   393216
      MaxLength       =   10
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   11.25
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
      Height          =   6375
      Left            =   210
      TabIndex        =   6
      Top             =   2385
      Width           =   14145
      _ExtentX        =   24950
      _ExtentY        =   11245
      _Version        =   393216
      FixedCols       =   0
      RowHeightMin    =   275
      BackColorFixed  =   -2147483647
      ForeColorFixed  =   -2147483634
      BackColorBkg    =   12632256
      Appearance      =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin Crystal.CrystalReport cryReport 
      Bindings        =   "frmPaymentToSubGrp.frx":058A
      Left            =   12375
      Top             =   180
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
      Left            =   5400
      TabIndex        =   2
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
      Format          =   59506691
      CurrentDate     =   -273
   End
   Begin VB.Label lbl1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Op.         :"
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
      Index           =   2
      Left            =   7560
      TabIndex        =   103
      Top             =   630
      Width           =   930
   End
   Begin VB.Label lbl1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Credit     :"
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
      Index           =   5
      Left            =   7560
      TabIndex        =   102
      Top             =   1440
      Width           =   930
   End
   Begin VB.Label lbl1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Debit      :"
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
      Index           =   6
      Left            =   7560
      TabIndex        =   101
      Top             =   1035
      Width           =   915
   End
   Begin VB.Label lbl1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Balance  :"
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
      Index           =   7
      Left            =   7560
      TabIndex        =   100
      Top             =   1845
      Width           =   960
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
      Left            =   15480
      TabIndex        =   42
      Top             =   495
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.Shape Shape6 
      BorderColor     =   &H00EFB17A&
      BorderWidth     =   5
      FillColor       =   &H00EFB17A&
      Height          =   9330
      Left            =   0
      Top             =   0
      Width           =   14505
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
      Left            =   10575
      TabIndex        =   37
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
      Left            =   5130
      TabIndex        =   22
      Top             =   8880
      Width           =   1215
   End
   Begin VB.Label lbl1 
      BackStyle       =   0  'Transparent
      Caption         =   "&Bank Cd.  :"
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
      TabIndex        =   21
      Top             =   1230
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
      Left            =   4545
      TabIndex        =   20
      Top             =   630
      Width           =   690
   End
   Begin VB.Shape Shape1 
      FillColor       =   &H00BDDFD1&
      FillStyle       =   0  'Solid
      Height          =   1695
      Left            =   11160
      Shape           =   4  'Rounded Rectangle
      Top             =   495
      Width           =   3195
   End
   Begin VB.Label Label6 
      BackColor       =   &H00EFB17A&
      Caption         =   "   Bank Payments"
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
      Left            =   45
      TabIndex        =   36
      Top             =   45
      Width           =   14475
   End
   Begin VB.Shape Shape2 
      BackColor       =   &H00808080&
      BorderColor     =   &H00000080&
      FillColor       =   &H00ECFFFF&
      FillStyle       =   0  'Solid
      Height          =   1755
      Left            =   7350
      Top             =   465
      Width           =   3705
   End
   Begin VB.Shape Shape3 
      BackColor       =   &H00ECFFFE&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00000080&
      FillColor       =   &H00ECFFFF&
      Height          =   1695
      Left            =   225
      Top             =   480
      Width           =   6960
   End
End
Attribute VB_Name = "frmPaymentToSubGroup"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Const tblName As String = "tblVoucher"
Const IdField As String = "Vno"
Const MaxNo As Long = 46
Dim CriteriaStr As String
Dim rstbl As Recordset, ShowFrmFlag As Boolean
Dim rsRecVsSale As Recordset
Dim varBookMark As Variant
Dim FirstTimeFlag As Boolean
Dim FormAction As Integer
Dim frmTypeStr As String
Dim i As Long
Dim rstblBlChk As Recordset
Dim Mon As Integer
Dim PartyCode As Long
Dim OldgFrmTypeStr As String
Dim mLastEntryDate As Date
Dim mLastEntryBank As String
Const OrderField As String = "year(Vdt),Vno"
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
    If GProcRstOpen(rstbl, "Select * from " & tblName & " where Vno = (Select min(Vno) from " & tblName & " where Vno <> 0 and " & CriteriaStr & ") and  " & CriteriaStr & "   ", "O") > 0 Then ReadFields
Case 3 'Last
'     If GProcRstOpen(rstbl, "Select * from " & tblName & " where Vno = (Select max(Vno) from " & tblName & " where Vno <> 0 and " & CriteriaStr & " )and  " & CriteriaStr & "   ", "O") > 0 Then ReadFields
     If GProcRstOpen(rstbl, "Select * from " & tblName & " where " & CriteriaStr & "  and  " & IdField & " = (Select max(" & IdField & ") from " & tblName & " where vdt = (SELECT MAX(vdt) from " & tblName & " where " & CriteriaStr & " )  and  " & IdField & " <> 0 and  " & CriteriaStr & ")", "R") > 0 Then ReadFields Else GProcClearForm Me, rstbl, MaxNo

Case 4 'Find
    gClsSearch.SearchMultiField "qryVoucherHelp", "Vno,Vdt,DebitAc,CreditAc,VAmt,VBillNo,VCtrNo", Array("Vno", "Vdt", "Debit A/c", "Credit A/c", "Amount", "Bill No", "VCtrNo"), Array(1400, 1300, 3500, 3500, 1800, 1600, 0), CriteriaStr, , OrderField, , , True
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
        txtCode(7) = mLastEntryBank
    End If
    
    If FirstTimeFlag = True Then
        FirstTimeFlag = False
    Else
        dtpDt(2).SetFocus
    End If
'    dtpDt(2) = Format(Now, "dd/MM/yyyy")
    txt(0).Enabled = False
    FillCombo
    SetGrid
    SetControlEd
Case 7 'Modify
    'If mURecModi = True Then
    '    SearchHelp
    '    DoEvents
    If IsRelatedRecord = False Then
        FormAction = vbDataActionUpdate
        Call GProcActivateControls(Me, True)
        Call GProcSetButtons(Me, vbDataActionUpdate)
        txt(0).Enabled = False
        SetControlEd
        dtpDt(2).SetFocus
    End If
Case 8 'Delete
   ' If mURecDel = True Then
     If IsRelatedRecord = False Then
        If MsgBox("Do You Really Delete This Record?", vbQuestion + vbYesNo + vbDefaultButton1, "Confirm Delete") = vbYes Then
'           gCn.Execute "Delete from tblvoucher where " & CriteriaStr & " and vno = " & txt(0)
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
   '     GProcSaveRecord Me, rstbl, gFormAction
'        gCn.Execute "Delete from tblvoucher where " & CriteriaStr & " and vno = " & txt(0)
'        GProcCreateVoucher FormAction, txt(0), 1, txt(2), frmTypeStr, txt(4), txt(5), Val(txt(6)), txt(11), txt(12), txt(13), txt(7), txt(8), txt(9), txt(10), False, txt(20), 0, Val(txt(21))
'        GProcCreateVoucher FormAction, txt(0), 2, txt(2), frmTypeStr, txt(5), txt(4), Val(txt(6)) * (-1), txt(11), txt(12), txt(13), txt(7), txt(8), txt(9), txt(10), False, txt(20), 0, Val(txt(21))
        '--- Related Rcord Rec Vs Sale
        If FormAction = vbDataActionAddNew Then
           Mon = Month(dtpDt(2))
           txt(0) = GProcGenerateIdMonthwise("TblVoucher", "Vno", Mon, "Vdt", "Vtype='" & frmTypeStr & "' and VFirm = '" & gCCode & "'   and VYear=" & gCYear)
        End If
        GProcRstOpen rstbl, "Select * from " & tblName & " where " & IdField & " = " & txt(0) & " and " & CriteriaStr, "O"
        SaveRelatedRecord
        mLastEntryDate = dtpDt(2)
        mLastEntryBank = txtCode(7)
        '--Update Outstanding Balance
'        GProcUpdateOutstandingBal CLng(txt(0)), UCase(txt(3)), gCYear
        rstbl.Requery
        ShowAccountBalance
        FormAction = vbDataActionClose
        Call GProcActivateControls(Me, False)
        Call GProcSetButtons(Me, True)
        msGrid.Col = 0
        msGrid.Row = 1
        txtGrid.Move msGrid.CellLeft + msGrid.Left, msGrid.CellTop + msGrid.Top, msGrid.CellWidth, msGrid.CellHeight
        lblBalance.Caption = ""
        If FormAction = vbDataActionAddNew Then
            If cmdBtn(6).Visible = True Then cmdBtn(6).SetFocus
        Else
            If cmdBtn(6).Visible = True Then cmdBtn(6).SetFocus
        End If
    End If
Case 14 'Cancel
    txtGrid.Visible = False
     If GProcRstOpen(rstbl, "Select * from " & tblName & " where " & CriteriaStr & "  and  " & IdField & " = (Select max(" & IdField & ") from " & tblName & " where vdt = (SELECT MAX(vdt) from " & tblName & " where " & CriteriaStr & " )  and  " & IdField & " <> 0 and   " & CriteriaStr & ")", "R") > 0 Then ReadFields Else GProcClearForm Me, rstbl, MaxNo
    
    
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
        If cmdBtn(7).Visible = True And cmdBtn(7).Enabled = True Then
           cmdBtn(7).SetFocus
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
Private Sub dtpDt_LostFocus(Index As Integer)
'If (frmTypeStr = "CP") Then
'    msGrid.Col = 1
'    msGrid_EnterCell
'End If
End Sub

Private Sub txtCode_GotFocus(Index As Integer)
If Index = 6 And txtCode(6) = "" Then
   If gCIsDepotFirm = 1 Then
      txtCode(6) = GProcGetColumnValue("tblMastAccount", "AcCode", "'" & gCDepotMainFirm & "' ", "N", "AcOurCode", "S")
   End If


'   If gCIsDepotFirm = 1 Then
'      txtCode(6) = GProcGetColumnValue("tblMastAccount", "AcCode", "'" & gCDepotMainFirm & "' ", "N", "AcOurCode", "S")
'   End If
'   txtCode(7) = "ACAB009"
'   txtCode(7).SelStart = 8
End If
End Sub
Private Sub txtCode_LostFocus(Index As Integer)
If Trim(txtCode(Index)) = "" Then Exit Sub
SetControlEd
Select Case Index
    Case 7  '-- Party  Code
         txtTemp(7) = GProcGetColumnValue("tblMastAccount", "AcOurCode", UCase(txtCode(7)), "S", "AcName", "S")
         If Trim(txtTemp(7)) = "" Then
            MsgBox ("Code Not Found ...")
            txtTemp(7) = ""
            txtTemp(7).Enabled = True
            txtTemp(7).SetFocus
         End If
    Case 6  '-- Bank  Code
         txtTemp(6) = GProcGetColumnValue("tblMastAccount", "AcOurCode", UCase(txtCode(6)), "S", "AcName", "S")
         If Trim(txtTemp(6)) = "" Then
            MsgBox ("Code Not Found ...")
            txtTemp(6) = ""
            txtTemp(6).Enabled = True
            txtTemp(6).SetFocus
         Else
            msGrid.Col = 1
            msGrid_EnterCell
         End If
End Select
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
    frmRptAccount.txt(1) = txtTemp(7)
Else
    frmRptAccount.txt(1) = txtTemp(6)
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
    mLastEntryDate = vbNull
    Me.Top = 300
    List1.Clear
    List1.AddItem "Cheque"
    List1.AddItem "RTGS"
    List1.AddItem "Transfer"
    List1.AddItem "D.D."
    List1.AddItem "Other"
    lblBalance.Caption = ""
    Me.Left = Screen.Width / 2 - Me.Width / 2
    frmTypeStr = gfrmTypeStr
    CriteriaStr = " VType = '" & frmTypeStr & "' and VFirm  = '" & gCCode & "' and VYear=" & gCYear
    Y = "Select * from " & tblName & " where " & CriteriaStr & " "
    If Trim(OrderField) <> "" Then Y = Y & " order by " & OrderField
'    GProcRstOpen rstbl, Y, "O"
    Select Case frmTypeStr
    Case "BP" '--Bank
        Label6.Caption = "Bank Payments"
    Case "CP" '--Cash
        Label6.Caption = "Cash Payments"
        lbl1(4).Caption = "Code        :"
'        lbl1(3).Caption = "A/c Name :"
    
    End Select
    FirstTimeFlag = True
    If Trim(mShowRec) <> "" Then
       GProcRstOpen rstbl, "Select * from " & tblName & " where  " & CriteriaStr & " and  " & IdField & " = " & CLng(Trim(mShowRec)), "R"
       ReadFields
       Call GProcSetButtons(Me, vbDataActionCancel)
       mShowRec = ""
    Else
       GProcRstOpen rstbl, "Select * from " & tblName & " where " & CriteriaStr & "    and  " & IdField & " = 0", "R"
       '---Grid
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

Private Sub mskDt_LostFocus(Index As Integer)
SaveBtnEd
If Index = 2 Then ShowAccountBalance
End Sub
Private Sub mskDt_GotFocus(Index As Integer)
GProcSelectBox Me.ActiveControl
End Sub

Private Sub txt_GotFocus(Index As Integer)
GProcSelectBox Me.ActiveControl
If Index = 8 Then
   SetControlEd
End If
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
SaveBtnEd
If Index = 8 Then
   SetControlEd
End If
End Sub
Private Sub txt_KeyPress(Index As Integer, KeyAscii As Integer)
    KeyAscii = GProcValidateKey(rstbl, Index, KeyAscii, txt(Index).text)
End Sub
Private Sub FillCombo()
If (frmTypeStr = "CP") Then
   txtTemp(6) = "Cash In Hand"
   txtCode(6) = GProcGetColumnValue("tblMastAccount", "AcName", txtTemp(6), "S", "AcOurCode", "N")
End If
End Sub
Private Sub SaveBtnEd()
'--Vno,Vdt,Amount,Cr A/c,Dr A/c
If txt(0) = "" Or GProcIsDateValid(dtpDt(2)) = False Or txtTemp(6) = "" Then
    cmdBtn(13).Enabled = False
    cmdBtn(13).BackColor = &H80000004
Else
    cmdBtn(13).Enabled = True
    cmdBtn(13).BackColor = &HFFFF80
End If
End Sub
Private Sub FillTxtFromTemp()
txt(1) = frmTypeStr
txt(2) = IIf(IsDate(dtpDt(2)), dtpDt(2), "") '--Vdt
txt(3) = gCYear
txt(4) = gCCode
txt(9) = txttemp1(9) & vbNullString
txt(12) = txttemp1(12) & vbNullString
txt(6) = GProcGetColumnValue("tblMastAccount", "AcName", txtTemp(6), "S", "Accode", "N")
txt(7) = GProcGetColumnValue("tblMastAccount", "AcName", txtTemp(7), "S", "AcCode", "N")
'txt(21) = chkDontShowInUnadj.Value
End Sub
Private Sub FillTempFromTxt()
'mskDt(2) = IIf(IsDate(txt(2)), CDate(txt(2)), "__/__/____") '--Vdt
dtpDt(2) = CDate(txt(2))

If IsDate(txt(11)) Then
   mskDt(11) = CDate(txt(11))
Else
   mskDt(11) = "__/__/____"
End If
txttemp1(9) = txt(9)
txttemp1(12) = txt(12)
txtTemp(6) = GProcGetColumnValue("tblMastAccount", "AcCode", txt(6), "N", "AcName", "S")
txtCode(6) = GProcGetColumnValue("tblMastAccount", "AcCode", txt(6), "N", "AcOurCode", "S")

txtTemp(7) = GProcGetColumnValue("tblMastAccount", "AcCode", txt(7), "N", "AcName", "S")
txtCode(7) = GProcGetColumnValue("tblMastAccount", "AcCode", txt(7), "N", "AcOurCode", "S")
'-- Fill Related Record Rec Vs Sale
FillRelatedRecord
ShowAccountBalance
End Sub
Private Function ValidateGrid() As Boolean
'--- Pending
Dim TotAmt As Double
Dim i As Long
Dim RCount As Long
Dim j As Long
Dim InvNo As Long
Dim InVtype As String
Dim InvYear As Long
Dim IntAcCode As Long
Dim TDSPaidAcCode As Long
Dim BillDet As String
If Trim(txtGrid.text) <> "" Then
        msGrid.text = txtGrid.text
End If
txtGrid.Visible = False
txtGrid.text = ""
BillDet = ""
CheckForNumberGrid
With msGrid
     For i = 1 To .rows - 1
         If CDbl(.TextMatrix(i, 3)) <= 0 Then
            MsgBox "Check Received Amount", vbInformation + vbOKOnly, Me.Caption
            .Col = 3
            msGrid.SetFocus
            Exit Function
         End If
     Next i
End With



'    '-- Duplicate invoice selection
'    For i = 1 To .rows - 1
'        If Trim(.TextMatrix(i, 1)) <> "" Then
'           BillDet = BillDet + .TextMatrix(i, 1) + "-" + CStr(Val(.TextMatrix(i, 4)) + Val(.TextMatrix(i, 6)) + Val(.TextMatrix(i, 9))) + ":"
'        End If
'
'        '-- Int A/c
'           If CDbl(.TextMatrix(i, 9)) > 0 Then
'                If gLatePayIntAcCode = GProcGetColumnValue("tblmastAccount", "AcName", "-", "S", "AcCode", "N") Then
'                    IntAcCode = 0
'                Else
'                    IntAcCode = gLatePayIntAcCode
'                End If
'                If IntAcCode = 0 Then
'                    MsgBox "Check Late Pay Interest Account.", vbInformation + vbOKOnly, Me.Caption
'                    GProcShowForm frmSetting, frmMain.mnuMstSetting(0)
'                    Exit Function
'                End If
'           End If
'        '---
'        '-- TDS Paid A/c ( For Less TDS Amt JV )
'           If CDbl(.TextMatrix(i, 13)) > 0 Then
'                If gTDSPaidAcCode = GProcGetColumnValue("tblmastAccount", "AcName", "-", "S", "AcCode", "N") Then
'                    TDSPaidAcCode = 0
'                Else
'                    TDSPaidAcCode = gTDSPaidAcCode
'                End If
'                If TDSPaidAcCode = 0 Then
'                    MsgBox "Check TDS Paid Account.", vbInformation + vbOKOnly, Me.Caption
'                    GProcShowForm frmSetting, frmMain.mnuMstSetting(0)
'                    Exit Function
'                End If
'           End If
'        '---
'        InvNo = .TextMatrix(i, 11)
'        InVtype = .TextMatrix(i, 0)
'        InvYear = .TextMatrix(i, 10)
'        For j = 1 To .rows - 1
'            If j <> i And InvNo = .TextMatrix(j, 11) And InVtype = .TextMatrix(j, 0) And InvYear = .TextMatrix(j, 10) And InvNo <> 0 Then
'                MsgBox "Duplicate Invoice Selection.", vbOKOnly, Me.Caption
'                msGrid.SetFocus
'                Exit Function
'            End If
'        Next j
'    Next i
'    If BillDet <> "" Then
'        BillDet = Left(Trim(BillDet) + Space(120), 120)
'        txt(8) = Left(BillDet, 40)
'        txt(9) = Mid(BillDet, 41, 40)
'        txt(10) = Mid(BillDet, 81, 40)
'    End If
'    '-- Adjusted Amt
'    If .rows > 1 Then
'       TotAmt = 0
'       For i = 1 To .rows - 1
'           '-- Adj amt > Bill Amt
'           If CDbl(.TextMatrix(i, 4)) - CDbl(.TextMatrix(i, 5)) > CDbl(.TextMatrix(i, 3)) Then
'                MsgBox "Adjusted Amount is more than Balance Bill Amount. ", vbOKOnly, Me.Caption
'                msGrid.SetFocus
'                Exit Function
'           End If
'            '-- Adj DB amt > Db Amt
''           If CDbl(.TextMatrix(i, 6)) > CDbl(.TextMatrix(i, 5)) Then
''                MsgBox "Adjusted DbNt Amount is more than Balance DbNt Amount. ", vbOKOnly, Me.Caption
''                msGrid.SetFocus
''                Exit Function
''           End If
'           TotAmt = TotAmt + Val(.TextMatrix(i, 4)) - Val(.TextMatrix(i, 5)) + Val(.TextMatrix(i, 9)) - Val(.TextMatrix(i, 13))
'        Next i
'        If (TotAmt > 0) And (TotAmt <> Val(txt(6))) Then
'           MsgBox "Total of Adjusted Amount " & TotAmt & " & Received Amount not tally.", vbOKOnly, Me.Caption
'     '      msGrid.SetFocus
'     '      Exit Function
'        End If
'    End If
'End With
ValidateGrid = True
End Function
Private Function ValidateData() As Boolean
Dim X As String
Dim GName As String
Dim BnkTot As Double
Dim RowIndex As Integer
Dim rstAddCheck As Recordset
GProcCheckForNumber Me, rstbl, MaxNo

RowIndex = 1
If txtTemp(7).text <> "" Then
    With msGrid
        Do While RowIndex < .rows
            '-- sub amt
            If Val(.TextMatrix(RowIndex, 3)) > 0 Then
                BnkTot = BnkTot + Val(.TextMatrix(RowIndex, 3))
            End If
            RowIndex = RowIndex + 1
        Loop
    End With
    If BnkTot > 0 Then
        MsgBox "For Bill Detail Party Name Should Be Blank.", vbInformation + vbOKOnly, Me.Caption
        txtTemp(6).SetFocus
        Exit Function
    End If
End If



'---- Required Data
If ValidateGrid = False Then Exit Function
'--- Vno
If CLng(txt(0)) = 0 And FormAction <> vbDataActionAddNew Then
    MsgBox "Check Voucher No.", vbInformation + vbOKOnly, Me.Caption
    txt(0).SetFocus
    Exit Function
End If
'--- Vdt
If GProcIsDateValid(dtpDt(2)) = False Then
    MsgBox "Check Voucher Date.", vbInformation + vbOKOnly, Me.Caption
    dtpDt(2).SetFocus
    Exit Function
End If

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
'---- party
'If txttemp(6) = "" Then
'    MsgBox "Check Account.", vbInformation + vbOKOnly, Me.Caption
'    txttemp(6).SetFocus
'    Exit Function
'End If
'--Bank A/c
'If GProcGetColumnValue("tblMastAccount", "AcName", txttemp(6), "S", "AcCode", "N") = 0 Then
'   GName = GProcGetColumnValue("TblMastGroup", "AgCode", 90017, "N", "AgName", "S")
'   GProcShowForm frmMastAccount, frmMain.mnuMstfrm(0), Array(frmMastAccount.txt(1), frmMastAccount.txttemp(2)), Array(txttemp(6), GName)
'   Exit Function
'End If
'--- Cr A/c
If txtTemp(6) = "" Then
    MsgBox "Check Account.", vbInformation + vbOKOnly, Me.Caption
    txtTemp(6).SetFocus
    Exit Function
End If
'--Create Cr A/c
If GProcGetColumnValue("tblMastAccount", "AcName", txtTemp(6), "S", "AcCode", "N") = 0 Then
   GName = GProcGetColumnValue("TblMastGroup", "AgCode", 90019, "N", "AgName", "S")
   GProcShowForm frmMastAccount, frmMain.mnuMstfrm(0), Array(frmMastAccount.txt(1), frmMastAccount.txtTemp(2)), Array(txtTemp(6), GName)
   Exit Function
End If
'--- Cr A/c ,Dr A/c
If txtTemp(6) = txtTemp(7) Then
    MsgBox "Check Account.Debit and Credit Accounts are Same.", vbInformation + vbOKOnly, Me.Caption
    txtTemp(6).SetFocus
    Exit Function
End If

'--- Pending

'---- Amount
'If CDbl(txt(6)) <= 0 Then
'    MsgBox "Check Voucher Amount.", vbInformation + vbOKOnly, Me.Caption
'    Exit Function
'End If
'----Duplicate Records
'----Addition Case


'If FormAction = vbDataActionAddNew Then
'    X = "Select vno From " & tblName & " Where vno=" & CLng(txt(0)) & " and Vtype='" & frmTypeStr & "' and VYear=" & gCYear
'    i = GProcRstOpen(rstAddCheck, X, "R")
'    If i > 0 Then
'        MsgBox "Record is Available of Voucher No. " & CLng(txt(0)), vbCritical + vbOKOnly, "Duplicate Record"
'        txt(0).SetFocus
'        Exit Function
'    End If
'    rstAddCheck.Close
'End If
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
        Case 6 '-- Dr A/c
         If frmTypeStr <> "CP" Then '-- Bank Receipt
            GName = GProcGetColumnValue("TblMastGroup", "AgCode", 90019, "N", "AgName", "S")
            GProcShowForm frmMastAccount, frmMain.mnuMstfrm(0), Array(frmMastAccount.txtTemp(2)), Array(GName)
         End If
        Case 7 '-- Cr A/c
            GName = GProcGetColumnValue("TblMastGroup", "AgCode", 90017, "N", "AgName", "S")
            GProcShowForm frmMastAccount, frmMain.mnuMstfrm(0), Array(frmMastAccount.txtTemp(2)), Array(GName)
    End Select
End If
End Sub
Private Sub txtTemp_KeyPress(Index As Integer, KeyAscii As Integer)
Dim LeftPos As Long
Dim TopPos As Long
TopPos = txtTemp(Index).Top + Me.Top + 650
LeftPos = txtTemp(Index).Left + Me.Left
If KeyAscii = 13 Then Exit Sub
'--Cash Receipt
'If (frmTypeStr = "CR") And (Index = 5) Then
'   KeyAscii = 0
'   txt(Index) = 2
'   txttemp(Index) = GProcGetColumnValue("tblmastaccount", "Accode", 2, "N", "Acname", "S")
'   Exit Sub
'End If
Select Case Index
    Case 6: '--Bank Payment
        gClsSearch.SearchMultiField "tblMastAccount", "AcName,AcCity,AcOurCode", Array("AcName", "City", "Code"), Array(4000, 2000, 1500), "AgCode=90019", Chr(KeyAscii), "AcName", LeftPos, TopPos, True
        KeyAscii = 0
        If gClsSearch.SearchMultiRetCol(0) <> "" Then
            txtTemp(Index) = gClsSearch.SearchMultiRetCol(0)
            txtCode(Index) = gClsSearch.SearchMultiRetCol(2)
            msGrid.Col = 1
            msGrid_EnterCell
        End If
    Case 7: '--Account
        gClsSearch.SearchMultiField "tblMastAccount", "AcName,AcCity", Array("AcName", "City"), Array(4000, 2000), "", Chr(KeyAscii), "AcName", LeftPos, TopPos, True
        KeyAscii = 0
        If gClsSearch.SearchMultiRetCol(0) <> "" Then
            txtTemp(Index) = gClsSearch.SearchMultiRetCol(0)
        End If
End Select
End Sub
Private Sub txtTemp_LostFocus(Index As Integer)
SaveBtnEd
If Index = 6 Then
   ShowAccountBalance
End If
DoEvents
End Sub

Private Sub txttemp1_GotFocus(Index As Integer)
If Index = 9 And msGrid.Enabled = False Then
   List1.Visible = True
   DoEvents
   List1.SetFocus
End If

End Sub
Private Sub list1_DblClick()
    Call list1_KeyPress(13)
    DoEvents
End Sub
Private Sub list1_GotFocus()
If Trim(txttemp1(9).text) <> "" Then
    List1.text = txttemp1(9).text
Else
    List1.ListIndex = 0
End If
End Sub
Private Sub list1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 And txt(10).Enabled = True Then
    txttemp1(9).text = Trim(List1.text)
    List1.Visible = False
    DoEvents
    txt(10).SetFocus
End If
End Sub
Private Sub list1_LostFocus()
    Call list1_KeyPress(13)
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
Dim X As String
Dim wAcCode As Long
Dim rstOpCal As Recordset
If txtTemp(6) <> "" And IsDate(dtpDt(2)) Then
   wAcCode = GProcGetColumnValue("tblMastAccount", "AcName", txtTemp(6), "S", "AcCode", "N")
   gCmd.CommandText = "PrcPrepareClBalance"
   gCmd.Parameters.Refresh
   gCmd.Parameters("@VFirm") = Trim(gCCode)
   gCmd.Parameters("@Fr_dt") = Format(dtpDt(2), "yyyy/mm/dd")
   gCmd.Parameters("@To_dt") = Format(dtpDt(2), "yyyy/mm/dd")
   gCmd.Parameters("@FyStart_dt") = Format(gCYSDate, "yyyy/mm/dd")
   gCmd.Parameters("@UCode") = gUserName
   gCmd.Parameters("@VFirm") = Trim(gCCode)
   gCmd.Parameters("@CGSTAc") = gCgstAcCode
   gCmd.Parameters("@SGSTAc") = gSgstAcCode
   gCmd.Parameters("@IGSTAc") = gIgstAcCode
   gCmd.Parameters("@TcsRec") = gTcsRec
   gCmd.Parameters("@TcsPay") = gTcsPay
   gCmd.Parameters("@CGSTRcmPayAc") = gCgstRCMAcCode
   gCmd.Parameters("@SGSTRcmPayAc") = gSgstRCMAcCode
   gCmd.Parameters("@IGSTRcmPayAc") = gIgstRCMAcCode
   If wAcCode <> 0 Then gCmd.Parameters("@Accode") = wAcCode
   gCmd.Execute
   X = "select OpBal_Dr,OpBal_Cr,Debit,Credit,ClBal_Dr,ClBal_Cr from TmpClosingBalance where accode = " & wAcCode & " and vfirm = '" & Trim(gCCode) & "' and UserName = '" & gUserName & "' "
   i = GProcRstOpen(rstOpCal, X, "R")
   If i > 0 Then
      txtOpBalance = IIf(rstOpCal.Fields("OpBal_Dr") > 0, Format(rstOpCal.Fields("OpBal_Dr"), "##########0.00") & " Dr", Format(rstOpCal.Fields("OpBal_Cr"), "##########0.00") & " Cr")
      txtDebit = Format(rstOpCal.Fields("Debit"), "##########0.00") & " Dr"
      txtCredit = Format(rstOpCal.Fields("Credit"), "##########0.00") & " Cr"
      txtClBalance = IIf(rstOpCal.Fields("ClBal_Dr") > 0, Format(rstOpCal.Fields("ClBal_Dr"), "########0.00") & " Dr", Format(rstOpCal.Fields("ClBal_Cr"), "########0.00") & " Cr")
   End If
End If
End Sub
'---------------- Grid
Private Sub SetGrid()
Dim X As String, x1Rec As Double
Set rsRecVsSale = New Recordset
X = "select VCtrNo,AcOurCode,AcName, VAmt,left(VRefTp,1),VRefTp,VRefNo,VRefBank,VBillNo,VBillType,VBillDate,VBillAmt,VBillVYear,VBillVno,VBillFirm" _
& " from tblVoucher,tblMastAccount where " _
& " VType='" & frmTypeStr & "' and VNo=" & txt(0) & " and VFirm  = '" & gCCode & "' and  VYear=" & gCYear & " and VDrAcCode = tblMastAccount.AcCode  order by VCtrNo"
x1Rec = GProcRstOpen(rsRecVsSale, X, "R")
With msGrid
    .Clear
    .FormatString = ">Sr|<Party Code|<Party Name                                      |>Paid. Amt     |   |<Ref  |<Ref.No.     |<Narration                   |<Bill No           |<Tp  ||>Bill Amt          |||"
    .ColWidth(10) = 0
    .ColWidth(12) = 0
    .ColWidth(13) = 0
    .ColWidth(14) = 0
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
     If ((.TextMatrix(.Row, 1)) = "") And .Row = 1 And txtGrid.Visible = True Then
        msGrid.Col = 1
        msGrid.SetFocus
        DoEvents
        If txtGrid.Visible = True Then
           txtGrid.SetFocus
           DoEvents
        End If
     End If
End With
   ' cmdBtn(10).Cancel = False
End Sub
Private Sub MSGrid_LostFocus()
CalAmount
SaveBtnEd
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
'1 - Party OurCode
'2 - AcName
'3 - VAmt
'5 - VRefTp
'6 - VRefNo
'7 - VRefBank
'8 - VBillNo
'9 - VBillType
'10 - VbillDt
'11 - VBillAmt
'12 -VBillVYear
'13 - VBillVno
'14 - VBillFirm
    With msGrid
        Select Case .Col
            Case 0
                txtGrid.Locked = True
                ReNumberGridSrno
                txtGrid = ""
                .Col = .Col + 1
            Case 9, 10, 5, 11, 12, 13, 14
                txtGrid.Locked = True
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
        If .Col = rsRecVsSale.Fields.Count - 7 Then '-4
            If .Row = (.rows - 1) Then '-- Add New Row
                .rows = .rows + 1
                .Row = .Row + 1
'                GridAddNew
            Else '--Next Row
                .Row = .Row + 1
            End If
            .Col = 1
        Else    '-- Next col
            .Col = .Col + 1
        End If
    End If
'-- Pending
'    If KeyAscii <> 13 And KeyAscii <> 27 And .Col = 1 Then
'        DoEvents
'        KeyAscii = 0
'        Pending_Invoice
'        If .TextMatrix(.Row, 1) = "" Then
'           msGrid.Col = 1
'        Else
'           msGrid.Col = 6
'           txtGrid.SetFocus
'        End If
'        Exit Sub
'        KeyAscii = 0
'    End If
End With

'---- Set TxtGrid
If Trim(msGrid.text) <> "" Then
    txtGrid = Trim(msGrid.text)
End If
If msGrid.Col < (rsRecVsSale.Fields.Count - 1) Then '- -3
   txtGrid.Move msGrid.CellLeft + msGrid.Left, msGrid.CellTop + msGrid.Top, msGrid.CellWidth, msGrid.CellHeight
End If
txtGrid.Visible = True
txtGrid.Enabled = True
DoEvents
If KeyAscii <> 0 Then
    Select Case rsRecVsSale.Fields(msGrid.Col).Type
        Case adChar, adVarWChar
             '-- Pending
'            If msGrid.Col = 0 Then  '-- Invoice
'                txtGrid_KeyPress (KeyAscii)
'            Else
                txtGrid.text = txtGrid.text & Chr(KeyAscii)
'            End If
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
DoEvents
End Sub
Private Sub msGrid_LeaveCell()
Dim X As String
Dim j As Integer
If txtGrid.Visible Then
    msGrid.text = txtGrid.text
    DoEvents
    txtGrid.text = ""
    txtGrid.Visible = False
End If
If msGrid.Col = 4 Then
    With msGrid
        Select Case UCase(msGrid.TextMatrix(.Row, 4))
        Case "C":
             msGrid.TextMatrix(.Row, 5) = "Chq."
        Case "R":
             msGrid.TextMatrix(.Row, 5) = "RTGS."
        Case "N":
             msGrid.TextMatrix(.Row, 5) = "NEFT."
        Case "D":
             msGrid.TextMatrix(.Row, 5) = "D.D."
        Case "O":
             msGrid.TextMatrix(.Row, 5) = "Other."
        End Select
        If frmTypeStr = "BP" Then
            If .Row = 1 And msGrid.TextMatrix(.Row, 6) = "" Then
               PartyCode = GProcGetColumnValue("tblMastAccount", "AcName", "" & txtTemp(6).text & "", "S", "AcCode", "N")
               msGrid.TextMatrix(.Row, 6) = Val((GProcGetColumnValue("tblVoucher", "VCrAcCode", "" & PartyCode & "", "N", "VRefNo", "S", "Vtype = '" & frmTypeStr & "' and VRefNo <> ''  and Vfirm = '" & gCCode & "' order by vdt desc,vno desc, VCtrNo desc "))) + 1
            End If
            If .Row > 1 And msGrid.TextMatrix(.Row, 6) = "" Then
               msGrid.TextMatrix(.Row, 6) = Val(msGrid.TextMatrix(.Row - 1, 6)) + 1
            End If
            If .Row > 1 And FormAction = vbDataActionAddNew Then
               If msGrid.TextMatrix(.Row, 2) = msGrid.TextMatrix(.Row - 1, 2) Then
                  msGrid.TextMatrix(.Row, 7) = msGrid.TextMatrix(.Row - 1, 7)
               End If
            End If
        End If
    End With
End If
'----Late Payment Interest Calcualtion
'---Interest Rate 1.5
Select Case msGrid.Col
Case 1, 2:
    With msGrid
        If .TextMatrix(.Row, 1) <> "" Then
           .TextMatrix(.Row, 2) = GProcGetColumnValue("tblMastAccount", "AcOurCode", UCase(.TextMatrix(.Row, 1)), "S", "AcName", "S")
        End If
        If .TextMatrix(.Row, 4) = "" Then
           If frmTypeStr = "CP" Then
              .TextMatrix(.Row, 4) = "O"
           Else
              .TextMatrix(.Row, 4) = "C"
           End If
        End If
        If .TextMatrix(.Row, 7) = "" And .TextMatrix(.Row, 1) <> "" Then
           j = GProcGetColumnValue("tblMastAccount", "AcOurCode", "" & .TextMatrix(.Row, 1) & "", "S", "AcCode", "N")
           DoEvents
           .TextMatrix(.Row, 7) = GProcGetColumnValue("tblVoucher", "VDrAcCode", "" & j & "", "N", "VRefBank", "S", "Vtype = '" & frmTypeStr & "' and Vfirm = '" & gCCode & "' order by vdt desc,vno desc, VCtrNo desc  ")
        End If
    End With
Case 7:
    With msGrid
         .TextMatrix(.Row, 7) = GProcProperCase(.TextMatrix(.Row, 7))
    End With
End Select
CalBalance
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
        DoEvents
    If txtGrid <> "" Then
        
        txtGrid.Alignment = ProcSetAlignment(rsRecVsSale, msGrid.Col)
        DoEvents
        Call GProcSelectBox(txtGrid)
     End If
        DoEvents
'    If msGrid.Col = 1 Then
'       Pending_Invoice
'    End If
End Sub
Private Sub txtGrid_KeyPress(KeyAscii As Integer)
Dim LeftPos As Long
Dim TopPos As Long
Dim FindStr As String
TopPos = txtGrid.Top + Me.Top + 650
LeftPos = txtGrid.Left + Me.Left
DoEvents
With msGrid
    If KeyAscii <> 13 And KeyAscii <> 27 Then
        '--- Search
        Select Case .Col
            Case 2 '-- party
                    gClsSearch.SearchMultiFieldAccMaster Chr(KeyAscii), 0, 0
                    txtGrid = ""
                    KeyAscii = 0
                    If gClsSearch.SearchMultiRetColAccMaster(0) <> "" Then
                        txtGrid = gClsSearch.SearchMultiRetColAccMaster(0)
                        .TextMatrix(.Row, 1) = GProcGetColumnValue("tblMastAccount", "AcName", "" & gClsSearch.SearchMultiRetColAccMaster(0) & "", "S", "AcOurCode", "S")
'                        PartyCode = GProcGetColumnValue("tblMastAccount", "AcOurCode", "" & .TextMatrix(.Row, 1) & "", "S", "AcCode", "N")
                        DoEvents
                        If .TextMatrix(.Row, 4) = "" Then
                            If frmTypeStr = "CP" Then
                               .TextMatrix(.Row, 4) = "O"
                            Else
                               .TextMatrix(.Row, 4) = "C"
                            End If
                        End If
                        txtGrid_KeyPress 13
                    End If
            Case Else
                KeyAscii = GProcValidateKey(rsRecVsSale, msGrid.Col, KeyAscii, txtGrid)
        End Select
   End If
   If KeyAscii = 13 Then
        CheckForNumberTxtGrid
        If (.Col = 8 And txtGrid <> "") Then
           Dim j As Integer, X As String
               j = .Row
               txtGrid = UCase(txtGrid)
               If .Col = 8 Then
                  .TextMatrix(j, 8) = UCase(txtGrid)
                  X = "select tblMastAccount.AcOurCode,tblMastAccount.AcName,SlBillNo,VType,format(Vdt,'dd/MM/yyyy'),SlBillAmt,SlBillAmt-SlBillRecdAmt," _
                      & " VYear,Vno,rtrim(vfirm),ltrim(str(Vno))+'-'+Vtype+'-'+ltrim(str(Vyear))+'-'+rtrim(vfirm) as VnoVtypeVyearVfirm,tblMastAccount.AcCode from tblSale, tblMastAccount where tblSale.SlAcDrCode = tblMastAccount.AcCode and " _
                      & " SlBillNo = '" & .TextMatrix(j, 8) & "' and vfirm = '" & gCCode & "'"
               End If
'               If .Col = 4 Then
'                  .TextMatrix(J, 4) = UCase(txtGrid)
'                  PartyCode = GProcGetColumnValue("tblMastAccount", "AcOurCode", "" & .TextMatrix(J, 4) & "", "S", "AcCode", "N")
'                  X = "select tblMastAccount.AcOurCode,tblMastAccount.AcName,SlBillNo,VType,format(Vdt,'dd/MM/yyyy'),SlBillAmt,SlBillAmt-SlBillRecdAmt," _
'                      & " VYear,Vno,ltrim(str(Vno))+'-'+Vtype+'-'+ltrim(str(Vyear))+'-'+rtrim(vfirm) as VnoVtypeVyearVfirm,tblMastAccount.AcCode from tblSale, tblMastAccount where tblSale.SlAcDrCode = tblMastAccount.AcCode and " _
'                      & " tblSale.SlAcDrCode = " & PartyCode & " and vfirm = '" & gCCode & "'"
'               End If
'               If .Col = 5 Then
'                  .TextMatrix(J, 5) = txtGrid
'                  X = "select tblMastAccount.AcOurCode,tblMastAccount.AcName,SlBillNo,VType,format(Vdt,'dd/MM/yyyy'),SlBillAmt,SlBillAmt-SlBillRecdAmt," _
'                      & " VYear,Vno,ltrim(str(Vno))+'-'+Vtype+'-'+ltrim(str(Vyear))+'-'+rtrim(vfirm) as VnoVtypeVyearVfirm,tblMastAccount.AcCode from tblSale, tblMastAccount where tblSale.SlAcDrCode = tblMastAccount.AcCode and " _
'                      & " SlBillAmt-SlBillRecdAmt > 0 and vfirm = '" & gCCode & "'"
'               End If
                i = GProcRstOpen(rstblBlChk, X, "R")
                If i = 0 Then
                  X = "select tblMastAccount.AcOurCode,tblMastAccount.AcName,SlBillNo,VType,format(Vdt,'dd/MM/yyyy'),SlBillAmt,SlBillAmt-SlBillRecdAmt," _
                      & " VYear,Vno,rtrim(vfirm),ltrim(str(Vno))+'-'+Vtype+'-'+ltrim(str(Vyear))+'-'+rtrim(vfirm) as VnoVtypeVyearVfirm,tblMastAccount.AcCode from tblSale, tblMastAccount where tblSale.SlAcDrCode = tblMastAccount.AcCode and " _
                      & " SlBillNo = '" & .TextMatrix(j, 8) & "' "
                      i = GProcRstOpen(rstblBlChk, X, "R")
                End If
                If i = 0 Then
                   MsgBox "No Data"
                   .TextMatrix(.Row, 8) = ""
                   txtGrid = ""
                   Exit Sub
                ElseIf i > 1 Then
                    gClsSearch.SearchMultiField X, "", Array("Party Code", "Party Name", "Bill No", "Type", "Date", "Bill Amt", "Balance Amt.", "VYear", "INV Vno", "Vfirm", "FindStr", ""), Array(1200, 5000, 1500, 600, 1500, 1700, 1700, 0, 0, 0, 0, 0), "", "", "Vdt,SlBillNo", 0, 0, False, , True
                    If gClsSearch.SearchMultiRetCol(1) = "" Then
                        MsgBox "No Data"
                        .TextMatrix(.Row, 8) = ""
                        txtGrid = ""
                        Exit Sub
                    End If
                    If .Col = 2 Then
                        txtGrid = gClsSearch.SearchMultiRetCol(1)
                    End If
                    rstblBlChk.MoveFirst
                    FindStr = gClsSearch.SearchMultiRetCol(8) & "-" & gClsSearch.SearchMultiRetCol(3) & "-" & gClsSearch.SearchMultiRetCol(7) & "-" & gClsSearch.SearchMultiRetCol(9)
                    rstblBlChk.Find "VnoVtypeVyearVfirm='" & FindStr & "'"
                End If
                If i > 0 Then
                   .TextMatrix(j, 8) = rstblBlChk.Fields(2)  '--Bill No
                   .TextMatrix(j, 9) = rstblBlChk.Fields(3)  '--Type
                   .TextMatrix(j, 10) = Format(rstblBlChk.Fields(4), "dd/MM/yyyy") '--VDt
                   .TextMatrix(j, 1) = rstblBlChk.Fields(0) & vbNullString '--Party Code
                   .TextMatrix(j, 2) = rstblBlChk.Fields(1) '--Party
                   .TextMatrix(j, 11) = Format(rstblBlChk.Fields(5), "#######0.00") '--Bill amt
'                   If FormAction = vbDataActionAddNew Then
 '                     .TextMatrix(J, 3) = Format(rstblBlChk.Fields(6), "#######0.00") '--Balance Amt
  '                 End If
                   .TextMatrix(j, 12) = CLng(rstblBlChk.Fields(7)) '--Year
                   .TextMatrix(j, 13) = rstblBlChk.Fields(8) '--Vno
                   .TextMatrix(j, 14) = rstblBlChk.Fields(9) '--vBillFirm
                   If .TextMatrix(j, 4) = "" Then
                       If frmTypeStr = "CP" Then
                          .TextMatrix(.Row, 4) = "O"
                       Else
                          .TextMatrix(j, 4) = "C"
                       End If
                   End If
'                   If .TextMatrix(J, 7) = "" Then
'                       .TextMatrix(J, 7) = GProcGetColumnValue("tblVoucher", "VDrAcCode", rstblBlChk.Fields(10), "N", "VRefBank", "S", "Vtype = 'BR' and Vfirm = '" & gCCode & "' order by vdt desc ")
'                   End If
'                 .Col = 7
              End If
        End If
        '----Adj Amt <=Inv Amt & Adj Dr Note Amt <  Dr Note Amt
'        Select Case .Col
'        Case 11:
'
'            If Val(txtGrid) > Val(.TextMatrix(.Row, 3)) Then
'                MsgBox "Adjustment Amount is less than Invoice Amount", vbCritical + vbOKOnly, "Adjustment Amount is less than Invoice Amount"
'                DoEvents
'                txtGrid.SetFocus
'                Exit Sub
'            End If
'        Case 5:
'            If Val(txtGrid) > Val(.TextMatrix(.Row, 3)) Then
'                MsgBox "Discount Amount must less than Invoice Amount", vbCritical + vbOKOnly, "Adjustment Dr Note Amount is less than Dr Note Amount"
'                DoEvents
'                txtGrid.SetFocus
'                Exit Sub
'            End If
'        End Select
    End If
    '--- Navigate Row ,Col
    If KeyAscii = 13 And .Col < rsRecVsSale.Fields.Count - 7 Then '- -4
        If .Col = 4 Then
            .Col = .Col + 2
        ElseIf .Col = 7 Then
            PartyCode = GProcGetColumnValue("tblMastAccount", "AcName", "" & .TextMatrix(.Row, 2) & "", "S", "AgCode", "N")
            If PartyCode <> 90017 Then
               If .Row = (.rows - 1) Then
                  .rows = .rows + 1
                  .Row = .Row + 1
                  .Col = 1
                  ReNumberGridSrno
                  
              Else
                  .Row = .Row + 1
                  .Col = 1
               End If
            Else
               .Col = .Col + 1
            End If
        Else
            .Col = .Col + 1
        End If
    ElseIf KeyAscii = 13 And .Col = rsRecVsSale.Fields.Count - 7 Then '- -4
        If .Row = (.rows - 1) Then
            .rows = .rows + 1
            .Row = .Row + 1
        Else
            .Row = .Row + 1
        End If
        .Col = 1
        ReNumberGridSrno
    End If
    '-- On ESC
        If KeyAscii = 27 Then '----Esc Key
            KeyAscii = 0
            If (Val(.TextMatrix(.Row, 3)) = 0) And .rows > 2 Then
                .RemoveItem (.Row)
                txtGrid.text = ""
                ReNumberGridSrno
            
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
End Sub
Private Sub txtGrid_KeyDown(KeyCode As Integer, Shift As Integer)
Dim TopPos As Long
Dim LeftPos As Long
Dim X As String
TopPos = txtGrid.Top + Me.Top + 650
LeftPos = txtGrid.Left + Me.Left

'--- Delete Row On Ctrl+Y
With msGrid
     If .Col = 7 And KeyCode = vbKeyF1 Then
        gClsSearch.SearchMultiField "tblMastNarration", "Narration", Array("Party Bank Name"), Array(6000), "NarrType='N'", "", "Narration", 3500, 4500, False
        If gClsSearch.SearchMultiRetCol(0) <> "" Then txtGrid = gClsSearch.SearchMultiRetCol(0)
     End If
     If .Col = 8 And KeyCode = vbKeyF1 Then
         X = "select tblMastAccount.AcOurCode,tblMastAccount.AcName,SlBillNo,VType,format(Vdt,'dd/MM/yyyy'),SlBillAmt,SlBillAmt-SlBillRecdAmt," _
             & " VYear,Vno,ltrim(str(Vno))+'-'+Vtype+'-'+ltrim(str(Vyear))+'-'+rtrim(vfirm) as VnoVtypeVyearVfirm,tblMastAccount.AcCode from tblSale, tblMastAccount where tblSale.SlAcDrCode = tblMastAccount.AcCode and " _
             & " SlBillAmt-SlBillRecdAmt > 0 and vfirm = '" & gCCode & "'"
         i = GProcRstOpen(rstblBlChk, X, "R")
         If i = 0 Then
            X = "select tblMastAccount.AcOurCode,tblMastAccount.AcName,SlBillNo,VType,format(Vdt,'dd/MM/yyyy'),SlBillAmt,SlBillAmt-SlBillRecdAmt," _
                & " VYear,Vno,ltrim(str(Vno))+'-'+Vtype+'-'+ltrim(str(Vyear))+'-'+rtrim(vfirm) as VnoVtypeVyearVfirm,tblMastAccount.AcCode from tblSale, tblMastAccount where tblSale.SlAcDrCode = tblMastAccount.AcCode and " _
                & " SlBillAmt-SlBillRecdAmt > 0 "
               i = GProcRstOpen(rstblBlChk, X, "R")
         End If
         If i > 1 Then
            gClsSearch.SearchMultiField X, "", Array("Party Code", "Party Name", "Bill No", "Type", "Date", "Bill Amt", "Balance Amt.", "VYear", "INV Vno", "VFirm", "FindStr", ""), Array(1200, 5000, 1500, 600, 1500, 1700, 1700, 0, 0, 0, 0, 0), "", "", "Vdt,SlBillNo", 0, 0, False, , True
            If gClsSearch.SearchMultiRetCol(0) <> "" Then txtGrid = gClsSearch.SearchMultiRetCol(2)
         End If
     End If
End With
If KeyCode = vbKeyY And Shift = 2 And msGrid.rows > 2 Then
    msGrid = 1
    msGrid.RemoveItem (msGrid.Row)
    txtGrid.Visible = False
    txtGrid.text = ""
    ReNumberGridSrno
'    CalAmount
    Exit Sub
End If
'--- Show Form
If KeyCode = vbKeyF3 Then
Select Case msGrid.Col
    Case 7 '-- Party Bank
        OldgFrmTypeStr = gfrmTypeStr
        gfrmTypeStr = "N"
        GProcShowForm frmMastNarration, frmMain.mnuMstfrm(10), "N"
        gfrmTypeStr = OldgFrmTypeStr
End Select
End If
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
If FormAction <> vbDataActionAddNew Then
Dim Maxi As Integer
Maxi = GProcGetColumnValue("tblVoucher", "VNo", txt(0), "N", "max(VCtrNo)", "N", "" & CriteriaStr & "", gCn)
For i = 1 To Maxi
    gCn.Execute "delete from tblVoucher where VNo=" & txt(0) & " and " & CriteriaStr & " and VCtrNo = " & i & " "
 Next
End If
'gCn.Execute "delete from tblVoucher where VNo=" & txt(0) & " and " & CriteriaStr & " "

'--Pending
'gCn.Execute "delete from tblVoucher where Vtype='JV' and VNo in (select JVNOINT from tblRecVsSale where " _
'& " INTRECEIVED>0 and RecVtype='" & frmTypeStr & "' and RecVNo=" & txt(0) & " and  RecVyear= " & gCYear & ")  and " _
' & " Vyear= " & gCYear
' '-- JV Less TDS
' gCn.Execute "delete from tblVoucher where Vtype='JV' and VNo in (select JVNOLessTDS from tblRecVsSale where " _
'& " LessTDS>0 and RecVtype='" & frmTypeStr & "' and RecVNo=" & txt(0) & " and  RecVyear= " & gCYear & ")  and " _
' & " Vyear= " & gCYear
'
''---Mill Bill Recd. in our bank
' '-- JV Less TDS
' gCn.Execute "delete from tblVoucher where Vtype='JV' and VNo in (select JvNoMillBill from tblRecVsSale where " _
'& "  RecVtype='" & frmTypeStr & "' and RecVNo=" & txt(0) & " and  RecVyear= " & gCYear & ")  and " _
' & " Vyear= " & gCYear
'
''---JV Discount
' gCn.Execute "delete from tblVoucher where Vtype='PN' and VNo in (select JvNoDisc from tblRecVsSale where " _
'& " Discount>0 and RecVtype='" & frmTypeStr & "' and RecVNo=" & txt(0) & " and  RecVyear= " & gCYear & ")  and " _
' & " Vyear= " & gCYear
'
'
''--- Rec Vs Sale
'gCn.Execute "delete from tblRecVsSale where RecVtype='" & frmTypeStr & "' and RecVNo=" & txt(0) & " and " _
' & " RecVyear= " & gCYear
End Sub
Private Sub SaveRelatedRecord()
Dim i As Long, X As String
Dim RowIndex As Integer
RowIndex = 1
With msGrid
    Do While RowIndex < .rows
       If (Val(.TextMatrix(RowIndex, 3)) <= 0) And .rows > 2 Then
          .RemoveItem (RowIndex)
       End If
       RowIndex = RowIndex + 1
    Loop
End With
'-- Delete Records
DeleteRelatedRecord
'--- Save Grid Records
'--- tblRecSale
'X = "select * from tblRecVsSale where Recvtype='" & frmTypeStr & "' and Recvno=" & txt(0) & " and " _
'& " RecVYear=" & gCYear
'GProcRstOpen rsRecVsSale, X, "O"
'0 - VCtrNo
'1 - Party OurCode
'2 - AcName
'3 - VAmt
'5 - VRefTp
'6 - VRefNo
'7 - VRefBank
'8 - VBillNo
'9 - VBillType
'10 - VbillDt
'11 - VBillAmt
'12 -VBillVYear
'13 - VBillVno
'14 - VBillFirm
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
            rstbl!VDrAcCode = GProcGetColumnValue("tblMastAccount", "AcName", .TextMatrix(i, 2), "S", "AcCode", "N")
            rstbl!VCrAcCode = txt(6)
            rstbl!VBillNo = .TextMatrix(i, 8) & vbNullString
            rstbl!VBillType = .TextMatrix(i, 9) & vbNullString
            If IsDate(.TextMatrix(i, 10)) Then
               rstbl!VBillDate = CDate(.TextMatrix(i, 10))
            End If
            rstbl!VBillAmt = Val(.TextMatrix(i, 11))
            rstbl!vamt = CDec(.TextMatrix(i, 3))
            rstbl!VRefTp = .TextMatrix(i, 5) & vbNullString
            rstbl!VRefNo = .TextMatrix(i, 6) & vbNullString
            rstbl!VRefBank = .TextMatrix(i, 7) & vbNullString
            rstbl!VBillVYear = .TextMatrix(i, 12)
            rstbl!VBillVno = .TextMatrix(i, 13)
            rstbl!VBillFirm = .TextMatrix(i, 14)
            rstbl!VNar1 = ""
            rstbl!VNar2 = ""
            rstbl!VNar3 = ""
            rstbl!VNAR4 = ""
            
'            If IsDate(txt(11)) = True Then
 '               rstbl!VRefDate = CDate(txt(11))
  '          Else
   '             rstbl!VRefDate = Null
    '        End If
     '       If IsDate(txt(17)) = True Then
      '          rstbl!VReconDt = CDate(txt(17))
       '     Else
        '        rstbl!VReconDt = Null
         '   End If
            rstbl!VIsRefEntType = Val(txt(18))
            rstbl!VIsAudited = Val(txt(19))
            rstbl!VTopCrDr = txt(26) & vbNullString
            rstbl!Discount = Val(txt(27))
            rstbl!LateDays = Val(txt(28))
            rstbl!Interest = Val(txt(29))
            rstbl!InterestCredit = Val(txt(30))
            rstbl!JvNoDisc = Val(txt(31))
            rstbl!JvNoInt = Val(txt(32))
            rstbl!LessTDS = Val(txt(33))
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
            rstbl!TDSOn = Val(txt(43))
            rstbl!DBNtInt = Val(txt(44))
            rstbl!DbNtMonthly = Val(txt(45))
            rstbl.Update
            DoEvents
        End If
        i = i + 1
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
txtTemp(7) = ""
txt(8) = Format(BnkTot, "###0.00")
txttemp1(9) = ""
txt(10) = ""
txttemp1(12) = ""
txt(13) = ""
txt(14) = ""
txt(15) = ""
txt(16) = ""

End Sub
Private Sub txttemp1_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
Dim OldFrmType As String
Dim LeftPos As Long
Dim TopPos As Long
If KeyCode = vbKeyF3 Then '--F3 Create
    Select Case Index
        Case 12 '--Narration
            OldFrmType = gfrmTypeStr
            gfrmTypeStr = "N"
            GProcShowForm frmMastNarration, frmMain.mnuMstfrm(5), , gfrmTypeStr
            gfrmTypeStr = OldFrmType
    End Select
End If
If KeyCode = vbKeyF2 Then '-- Select
    TopPos = txttemp1(Index).Top + Me.Top + 650
    LeftPos = txttemp1(Index).Left + Me.Left
    Select Case Index
        Case 12 '--Narration
            gClsSearch.SearchMultiField "tblMastNarration", "Narration", Array("Narration"), Array(txttemp1(Index).Width), "NarrType='N'", "", "Narration", LeftPos, TopPos, False
            If gClsSearch.SearchMultiRetCol(0) <> "" Then txttemp1(Index) = gClsSearch.SearchMultiRetCol(0)
    End Select
End If
End Sub
Private Sub CalAmount()
Dim BnkTot As Double
Dim RowIndex As Integer
RowIndex = 1
With msGrid
    Do While RowIndex < .rows
        '-- sub amt
        If Val(.TextMatrix(RowIndex, 3)) > 0 Then
            BnkTot = BnkTot + Val(.TextMatrix(RowIndex, 3))
        End If
        RowIndex = RowIndex + 1
    Loop
End With
txt(8) = Format(BnkTot, "###0.00")
End Sub
Private Sub SetControlEd()
txt(8).Enabled = False
If (frmTypeStr = "CP") Then
   txtCode(6).Enabled = False
   txtTemp(6).Enabled = False
End If
txtOpBalance.Enabled = False
txtDebit.Enabled = False
txtCredit.Enabled = False
txtClBalance.Enabled = False
End Sub

Private Sub txttemp1_LostFocus(Index As Integer)
If Index = 12 Then
   txttemp1(12) = StrConv(Trim(txttemp1(12)), vbProperCase)
End If
End Sub

Private Sub ReNumberGridSrno()
Dim i As Long
With msGrid
    For i = 1 To .rows - 1
        .TextMatrix(i, 0) = i
    Next i
End With
End Sub


