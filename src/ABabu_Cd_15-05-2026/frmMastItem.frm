VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form frmMastItem 
   BackColor       =   &H00FBF2E1&
   BorderStyle     =   0  'None
   Caption         =   " Count Master"
   ClientHeight    =   8700
   ClientLeft      =   0
   ClientTop       =   60
   ClientWidth     =   11460
   FillColor       =   &H000040C0&
   FillStyle       =   0  'Solid
   Icon            =   "frmMastItem.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   8700
   ScaleWidth      =   11460
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFF80&
      Height          =   2895
      Left            =   3360
      TabIndex        =   80
      Top             =   1440
      Visible         =   0   'False
      Width           =   5115
      Begin VB.CommandButton cmdUpdate 
         Caption         =   "Update"
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
         Left            =   420
         TabIndex        =   84
         Top             =   2070
         Visible         =   0   'False
         Width           =   1935
      End
      Begin VB.TextBox txtOldHSN 
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
         Height          =   360
         Left            =   2415
         MaxLength       =   20
         TabIndex        =   83
         Top             =   720
         Visible         =   0   'False
         Width           =   2055
      End
      Begin VB.TextBox txtNewHSN 
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
         Height          =   360
         Left            =   2400
         MaxLength       =   20
         TabIndex        =   82
         Top             =   1200
         Visible         =   0   'False
         Width           =   2055
      End
      Begin VB.CommandButton cmdExit 
         Caption         =   "Exit"
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
         Left            =   2580
         TabIndex        =   81
         Top             =   2040
         Width           =   1935
      End
      Begin VB.Label Label2 
         BackColor       =   &H00FFFF80&
         Caption         =   "Old HSN No     :"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   600
         TabIndex        =   86
         Top             =   720
         Visible         =   0   'False
         Width           =   1635
      End
      Begin VB.Label Label3 
         BackColor       =   &H00FFFF80&
         Caption         =   "New HSN No   :"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   600
         TabIndex        =   85
         Top             =   1200
         Visible         =   0   'False
         Width           =   1635
      End
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
      Height          =   360
      Index           =   35
      Left            =   13080
      TabIndex        =   93
      Text            =   "35"
      Top             =   1320
      Visible         =   0   'False
      Width           =   345
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
      Height          =   360
      Index           =   34
      Left            =   12960
      TabIndex        =   92
      Text            =   "34"
      Top             =   960
      Visible         =   0   'False
      Width           =   345
   End
   Begin VB.TextBox txt 
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
      Index           =   33
      Left            =   2550
      MaxLength       =   20
      TabIndex        =   3
      Text            =   "33"
      Top             =   1380
      Width           =   3360
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
      Height          =   360
      Index           =   32
      Left            =   13020
      TabIndex        =   88
      Text            =   "32"
      Top             =   600
      Visible         =   0   'False
      Width           =   345
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
      Height          =   360
      Index           =   31
      Left            =   13080
      TabIndex        =   87
      Text            =   "31"
      Top             =   180
      Visible         =   0   'False
      Width           =   345
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
      Height          =   360
      Index           =   30
      Left            =   11580
      TabIndex        =   79
      Text            =   "30"
      Top             =   5940
      Visible         =   0   'False
      Width           =   345
   End
   Begin VB.CheckBox chkIsInInsu 
      BackColor       =   &H00EDFDFE&
      Caption         =   "Booking Rate Excluding Insurance"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   330
      Left            =   3300
      TabIndex        =   23
      Top             =   8040
      Width           =   5520
   End
   Begin VB.CheckBox chkInDailyReport 
      BackColor       =   &H00EDFDFE&
      Caption         =   "Count Not Required in Daily Rport ?"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   330
      Left            =   4260
      TabIndex        =   20
      Top             =   6780
      Width           =   4920
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
      Height          =   360
      Index           =   29
      Left            =   11520
      TabIndex        =   78
      Text            =   "29"
      Top             =   5460
      Visible         =   0   'False
      Width           =   345
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
      Height          =   360
      Index           =   28
      Left            =   11700
      TabIndex        =   77
      Text            =   "28"
      Top             =   4440
      Visible         =   0   'False
      Width           =   345
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
      Height          =   360
      Index           =   27
      Left            =   11460
      TabIndex        =   76
      Text            =   "27"
      Top             =   4980
      Visible         =   0   'False
      Width           =   345
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
      Height          =   360
      Index           =   26
      Left            =   11640
      TabIndex        =   75
      Text            =   "26"
      Top             =   3900
      Visible         =   0   'False
      Width           =   345
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
      Height          =   360
      Index           =   25
      Left            =   12375
      TabIndex        =   74
      Text            =   "25"
      Top             =   5850
      Visible         =   0   'False
      Width           =   300
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
      Height          =   360
      Index           =   24
      Left            =   12375
      TabIndex        =   73
      Text            =   "24"
      Top             =   5490
      Visible         =   0   'False
      Width           =   300
   End
   Begin VB.CheckBox chkIsFrghtInsu 
      BackColor       =   &H00EDFDFE&
      Caption         =   "Booking Rate Including Freight"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   330
      Left            =   3300
      TabIndex        =   22
      Top             =   7680
      Width           =   5520
   End
   Begin VB.CheckBox chkIsExGST 
      BackColor       =   &H00EDFDFE&
      Caption         =   "Booking Rate Excluding GST"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   330
      Left            =   3300
      TabIndex        =   21
      Top             =   7320
      Visible         =   0   'False
      Width           =   3900
   End
   Begin VB.TextBox txt 
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
      Index           =   23
      Left            =   2550
      MaxLength       =   20
      TabIndex        =   4
      Text            =   "23"
      Top             =   1770
      Width           =   3360
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
      Index           =   21
      ItemData        =   "frmMastItem.frx":058A
      Left            =   2550
      List            =   "frmMastItem.frx":0594
      TabIndex        =   12
      Text            =   "Cbo(21)"
      Top             =   4710
      Width           =   1515
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
      Height          =   360
      Index           =   22
      Left            =   7380
      TabIndex        =   13
      Text            =   "22"
      Top             =   4320
      Width           =   1515
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
      Height          =   360
      Index           =   21
      Left            =   13050
      TabIndex        =   69
      Text            =   "21"
      Top             =   5985
      Visible         =   0   'False
      Width           =   345
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
      Height          =   360
      Index           =   20
      Left            =   2550
      TabIndex        =   11
      Text            =   "20"
      Top             =   4320
      Width           =   1515
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
      Height          =   360
      Index           =   19
      Left            =   12960
      TabIndex        =   68
      Text            =   "19"
      Top             =   5670
      Visible         =   0   'False
      Width           =   345
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
      Height          =   360
      Index           =   18
      Left            =   12375
      TabIndex        =   66
      Text            =   "18"
      Top             =   4770
      Visible         =   0   'False
      Width           =   345
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
      Index           =   18
      ItemData        =   "frmMastItem.frx":05A1
      Left            =   7365
      List            =   "frmMastItem.frx":05AE
      TabIndex        =   10
      Text            =   "Cbo(18)"
      Top             =   3720
      Width           =   1515
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
      Height          =   360
      Index           =   17
      Left            =   9270
      TabIndex        =   64
      Text            =   "17"
      Top             =   8730
      Visible         =   0   'False
      Width           =   1350
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
      Index           =   8
      ItemData        =   "frmMastItem.frx":05CE
      Left            =   2550
      List            =   "frmMastItem.frx":05DB
      TabIndex        =   16
      Text            =   "Cbo(8)"
      Top             =   6390
      Width           =   1515
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
      Height          =   360
      Index           =   15
      Left            =   2550
      TabIndex        =   17
      Text            =   "15"
      Top             =   6795
      Width           =   1515
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
      Height          =   360
      Index           =   16
      Left            =   10845
      TabIndex        =   60
      Text            =   "16"
      Top             =   8775
      Visible         =   0   'False
      Width           =   1350
   End
   Begin VB.CommandButton cmdBtn 
      Caption         =   "Creation &To Firm"
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
      Left            =   12375
      Style           =   1  'Graphical
      TabIndex        =   59
      Top             =   2700
      Visible         =   0   'False
      Width           =   1095
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
      Height          =   360
      Index           =   14
      Left            =   12375
      TabIndex        =   24
      Text            =   "14"
      Top             =   4005
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
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Index           =   13
      Left            =   12375
      TabIndex        =   25
      Text            =   "13"
      Top             =   5175
      Visible         =   0   'False
      Width           =   300
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
      Height          =   360
      Index           =   12
      Left            =   7830
      TabIndex        =   58
      Text            =   "12"
      Top             =   9000
      Visible         =   0   'False
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
      Height          =   360
      Index           =   11
      Left            =   12375
      TabIndex        =   26
      Text            =   "11"
      Top             =   4410
      Visible         =   0   'False
      Width           =   345
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
      Height          =   360
      Index           =   10
      Left            =   7380
      TabIndex        =   19
      Text            =   "10"
      Top             =   6270
      Visible         =   0   'False
      Width           =   1515
   End
   Begin VB.TextBox txt 
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
      Height          =   345
      Index           =   7
      Left            =   3420
      TabIndex        =   56
      Text            =   "7"
      Top             =   9090
      Visible         =   0   'False
      Width           =   315
   End
   Begin VB.TextBox txt 
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
      Height          =   345
      Index           =   6
      Left            =   3960
      TabIndex        =   55
      Text            =   "6"
      Top             =   9180
      Visible         =   0   'False
      Width           =   315
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
      Index           =   9
      Left            =   2550
      TabIndex        =   5
      Text            =   "txttemp(9)"
      ToolTipText     =   "Press F3 For Creation"
      Top             =   2205
      Width           =   6330
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
      Index           =   7
      ItemData        =   "frmMastItem.frx":05ED
      Left            =   2550
      List            =   "frmMastItem.frx":05FA
      TabIndex        =   14
      Text            =   "Cbo(7)"
      Top             =   5340
      Width           =   1515
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
      Index           =   6
      ItemData        =   "frmMastItem.frx":0612
      Left            =   7380
      List            =   "frmMastItem.frx":0622
      TabIndex        =   9
      Text            =   "Cbo(6)"
      Top             =   3270
      Width           =   1515
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
      Height          =   360
      Index           =   8
      Left            =   2550
      TabIndex        =   15
      Text            =   "8"
      Top             =   5700
      Width           =   1515
   End
   Begin VB.TextBox txt 
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
      Height          =   345
      Index           =   9
      Left            =   2790
      TabIndex        =   54
      Text            =   "9"
      Top             =   9090
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
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Index           =   5
      Left            =   2550
      TabIndex        =   8
      Text            =   "5"
      Top             =   3720
      Width           =   1515
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
      Height          =   360
      Index           =   4
      Left            =   2550
      TabIndex        =   7
      Text            =   "4"
      Top             =   3270
      Width           =   1515
   End
   Begin VB.TextBox txt 
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
      Index           =   3
      Left            =   2550
      TabIndex        =   2
      Text            =   "3"
      Top             =   945
      Width           =   3360
   End
   Begin VB.TextBox txt 
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
      Index           =   2
      Left            =   2550
      TabIndex        =   6
      Text            =   "2"
      Top             =   2640
      Width           =   3360
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
      Height          =   375
      Index           =   1
      Left            =   2550
      TabIndex        =   1
      Text            =   "1"
      Top             =   495
      Width           =   6330
   End
   Begin VB.CommandButton cmdBtn 
      Caption         =   "Update HSN"
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
      Left            =   9480
      Style           =   1  'Graphical
      TabIndex        =   51
      ToolTipText     =   "Label"
      Top             =   7800
      Width           =   1755
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
      Height          =   405
      Index           =   10
      Left            =   9900
      Style           =   1  'Graphical
      TabIndex        =   49
      ToolTipText     =   "Exit from form"
      Top             =   6315
      Width           =   960
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
      Height          =   405
      Index           =   14
      Left            =   9900
      Style           =   1  'Graphical
      TabIndex        =   42
      ToolTipText     =   "cancel"
      Top             =   5835
      Width           =   960
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
      Height          =   405
      Index           =   13
      Left            =   9900
      Style           =   1  'Graphical
      TabIndex        =   27
      ToolTipText     =   "Save Record"
      Top             =   5355
      Width           =   960
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
      Left            =   6705
      Style           =   1  'Graphical
      TabIndex        =   50
      ToolTipText     =   "Print Record"
      Top             =   8865
      Visible         =   0   'False
      Width           =   1095
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
      Height          =   405
      Index           =   8
      Left            =   9900
      Style           =   1  'Graphical
      TabIndex        =   41
      ToolTipText     =   "Delete Record"
      Top             =   4500
      Width           =   960
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
      Left            =   9900
      Style           =   1  'Graphical
      TabIndex        =   40
      ToolTipText     =   " Modify Record"
      Top             =   4020
      Width           =   960
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
      Height          =   405
      Index           =   6
      Left            =   9900
      Style           =   1  'Graphical
      TabIndex        =   39
      ToolTipText     =   " Add New Record"
      Top             =   3540
      Width           =   960
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
      Height          =   405
      Index           =   5
      Left            =   9900
      MaskColor       =   &H00000000&
      Style           =   1  'Graphical
      TabIndex        =   43
      ToolTipText     =   "List Of Record"
      Top             =   2670
      Width           =   960
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
      Height          =   405
      Index           =   4
      Left            =   9900
      Style           =   1  'Graphical
      TabIndex        =   44
      ToolTipText     =   "Find Record"
      Top             =   2205
      Width           =   960
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
      Left            =   10455
      Style           =   1  'Graphical
      TabIndex        =   48
      ToolTipText     =   "Last"
      Top             =   1425
      Width           =   405
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
      Left            =   9900
      Style           =   1  'Graphical
      TabIndex        =   47
      ToolTipText     =   "First"
      Top             =   1425
      Width           =   405
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
      Left            =   10455
      Style           =   1  'Graphical
      TabIndex        =   46
      ToolTipText     =   "Next"
      Top             =   945
      Width           =   405
   End
   Begin VB.TextBox txt 
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
      Height          =   345
      Index           =   0
      Left            =   12375
      TabIndex        =   52
      Text            =   "0"
      Top             =   585
      Visible         =   0   'False
      Width           =   315
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
      Left            =   9900
      Style           =   1  'Graphical
      TabIndex        =   45
      ToolTipText     =   "Previous"
      Top             =   945
      Width           =   405
   End
   Begin MSMask.MaskEdBox mskDt 
      Height          =   360
      Left            =   10980
      TabIndex        =   18
      Top             =   7320
      Visible         =   0   'False
      Width           =   375
      _ExtentX        =   661
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
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Brokerage Before  - 01/11/2025"
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
      Index           =   23
      Left            =   720
      TabIndex        =   94
      Top             =   6060
      Width           =   2880
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "From 01/04/2021"
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
      Index           =   22
      Left            =   6060
      TabIndex        =   91
      Top             =   1440
      Width           =   1515
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Up To 31/03/2021"
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
      Index           =   21
      Left            =   6060
      TabIndex        =   90
      Top             =   1860
      Width           =   1575
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "HSN Code        :"
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
      Index           =   20
      Left            =   720
      TabIndex        =   89
      Top             =   1380
      Width           =   1500
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "For Count Booking "
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
      Index           =   18
      Left            =   705
      TabIndex        =   72
      Top             =   7380
      Width           =   1830
   End
   Begin VB.Shape Shape10 
      BackColor       =   &H00ECFFFE&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00000080&
      FillColor       =   &H00ECFFFF&
      Height          =   1155
      Left            =   450
      Top             =   7290
      Width           =   8760
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "HSN Code        :"
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
      Index           =   19
      Left            =   750
      TabIndex        =   71
      Top             =   1815
      Width           =   1500
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Purchase Brokerage"
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
      Index           =   17
      Left            =   5475
      TabIndex        =   70
      Top             =   5445
      Visible         =   0   'False
      Width           =   1935
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Charity On        :"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   240
      Index           =   16
      Left            =   720
      TabIndex        =   67
      Top             =   4755
      Width           =   1470
   End
   Begin VB.Shape Shape5 
      BackColor       =   &H00C00000&
      BackStyle       =   1  'Opaque
      Height          =   150
      Left            =   5760
      Shape           =   3  'Circle
      Top             =   3810
      Width           =   150
   End
   Begin VB.Shape Shape3 
      BackColor       =   &H00C00000&
      BackStyle       =   1  'Opaque
      Height          =   150
      Left            =   5760
      Shape           =   3  'Circle
      Top             =   3390
      Width           =   150
   End
   Begin VB.Shape Shape2 
      BackColor       =   &H00C00000&
      BackStyle       =   1  'Opaque
      Height          =   150
      Left            =   360
      Shape           =   3  'Circle
      Top             =   1845
      Width           =   150
   End
   Begin VB.Shape Shape4 
      BackColor       =   &H00C00000&
      BackStyle       =   1  'Opaque
      Height          =   150
      Left            =   360
      Shape           =   3  'Circle
      Top             =   540
      Width           =   150
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Item Type  : "
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   240
      Index           =   15
      Left            =   6075
      TabIndex        =   65
      Top             =   3780
      Width           =   1185
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "&Brokerage On      :"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   240
      Index           =   14
      Left            =   5475
      TabIndex        =   63
      Top             =   5910
      Visible         =   0   'False
      Width           =   1620
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Old Brokerage On :"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   240
      Index           =   13
      Left            =   750
      TabIndex        =   62
      Top             =   6480
      Width           =   1680
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Old Brokerage Rt. :"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   240
      Index           =   12
      Left            =   750
      TabIndex        =   61
      Top             =   6795
      Width           =   1680
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Charity (GF) @  :"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   240
      Index           =   11
      Left            =   705
      TabIndex        =   34
      Top             =   4380
      Width           =   1500
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Exemption %  :"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   240
      Index           =   1
      Left            =   5850
      TabIndex        =   35
      Top             =   4380
      Width           =   1485
   End
   Begin VB.Shape Shape6 
      BorderColor     =   &H00EFB17A&
      BorderWidth     =   5
      FillColor       =   &H00EFB17A&
      Height          =   8745
      Left            =   0
      Top             =   -90
      Width           =   11415
   End
   Begin VB.Label lblActionStatus 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "View Record.."
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   375
      Left            =   8010
      TabIndex        =   57
      Top             =   15
      Width           =   2475
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "&Purch Brok Rt.     :"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   240
      Index           =   0
      Left            =   5475
      TabIndex        =   38
      Top             =   6330
      Visible         =   0   'False
      Width           =   1650
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Brokerage Rate    :"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   240
      Index           =   10
      Left            =   750
      TabIndex        =   37
      Top             =   5700
      Width           =   1650
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "M&ill Name        :"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   240
      Index           =   9
      Left            =   765
      TabIndex        =   29
      Top             =   2205
      Width           =   1500
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "&Brokerage On      :"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   240
      Index           =   8
      Left            =   750
      TabIndex        =   36
      Top             =   5340
      Width           =   1620
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Item Unit   :"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   240
      Index           =   7
      Left            =   6075
      TabIndex        =   33
      Top             =   3330
      Width           =   1095
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "&Rate Scale        :"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   240
      Index           =   6
      Left            =   750
      TabIndex        =   32
      Top             =   3780
      Width           =   1500
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Std Packing      :"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   240
      Index           =   5
      Left            =   750
      TabIndex        =   31
      Top             =   3285
      Width           =   1485
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Short Name       :"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   240
      Index           =   4
      Left            =   750
      TabIndex        =   28
      Top             =   945
      Width           =   1515
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Ticket Name      :"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   240
      Index           =   3
      Left            =   750
      TabIndex        =   30
      Top             =   2640
      Width           =   1515
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Count &Name     :"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   240
      Index           =   2
      Left            =   750
      TabIndex        =   0
      Top             =   510
      Width           =   1530
   End
   Begin VB.Label Label6 
      BackColor       =   &H00EFB17A&
      Caption         =   "   Count Master"
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
      TabIndex        =   53
      Top             =   0
      Width           =   11445
   End
   Begin VB.Shape Shape1 
      FillColor       =   &H00FED8FC&
      FillStyle       =   0  'Solid
      Height          =   6765
      Left            =   9645
      Shape           =   4  'Rounded Rectangle
      Top             =   495
      Width           =   1440
   End
   Begin VB.Shape Shape7 
      BackColor       =   &H00ECFFFE&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00000080&
      FillColor       =   &H00ECFFFF&
      Height          =   1110
      Left            =   450
      Top             =   3060
      Width           =   8760
   End
   Begin VB.Shape Shape8 
      BackColor       =   &H00ECFFFE&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00000080&
      FillColor       =   &H00ECFFFF&
      Height          =   960
      Left            =   450
      Top             =   4215
      Width           =   8760
   End
   Begin VB.Shape Shape9 
      BackColor       =   &H00ECFFFE&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00000080&
      FillColor       =   &H00ECFFFF&
      Height          =   1965
      Left            =   450
      Top             =   5220
      Width           =   8760
   End
End
Attribute VB_Name = "frmMastItem"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Const tblName As String = "tblMastItem"
Const IdField As String = "ItCode"
Const OrderField As String = "ItName"
Const MaxNo As Long = 35
Dim rstbl As Recordset
Dim CriteriaStr As String
Dim varBookMark As Variant
Dim FirstTimeFlag As Boolean
Dim FormAction As Integer
Dim ExemptMill As Integer
Private Sub Cbo_KeyPress(Index As Integer, KeyAscii As Integer)
If KeyAscii = 13 Then Exit Sub
If Index = 7 Or Index = 6 Or Index = 8 Or Index = 18 Then KeyAscii = 0
End Sub
Private Sub cmdBtn_Click(Index As Integer)
'On Error GoTo ErrorRoutine
Select Case Index
Case 0 '---- Previous
     If GProcRstOpen(rstbl, "Select * from " & tblName & " where " & OrderField & "  = (Select max(" & OrderField & " ) from " & tblName & " where  " & OrderField & " < '" & txt(1) & "' and " & CriteriaStr & " )and " & CriteriaStr, "R") > 0 Then ReadFields Else MsgBox "This is First Record:", vbOKOnly + vbInformation
Case 1 '---- Next
      If GProcRstOpen(rstbl, "Select * from " & tblName & " where " & OrderField & " = (Select min(" & OrderField & ") from " & tblName & " where " & OrderField & " > '" & txt(1) & "' and " & CriteriaStr & ")and " & CriteriaStr, "R") > 0 Then ReadFields Else MsgBox "This is last Record:", vbOKOnly + vbInformation
Case 2 '---- First
         If GProcRstOpen(rstbl, "Select * from " & tblName & " where " & OrderField & " = (Select min(" & OrderField & ") from " & tblName & " where " & OrderField & " <> '' and " & CriteriaStr & ")and " & CriteriaStr, "R") > 0 Then ReadFields
Case 3 '---- Last
        If GProcRstOpen(rstbl, "Select * from " & tblName & " where " & OrderField & " = (Select max(" & OrderField & ") from " & tblName & " where " & OrderField & " <> '' and " & CriteriaStr & ")and " & CriteriaStr, "R") > 0 Then ReadFields Else GProcClearForm Me, rstbl, MaxNo
Case 4 'Find
'    gClsSearch.SearchMultiField "tblMastItem,tblMastAccount", "ItName,ItCode,Itticket,Acname,ItStdpack,ItStdrateper,ITUnit,ItBrokRt,ItHsn", Array("Item", "Item", "Ticket", "Mill", "Std.Packing", "Rate Per", "Unit", "Brk.Rt.", "HSN"), Array(5000, 0, 1000, 2000, 1000, 1000, 1000, 1000, 1000), " tblMastItem.ItMillcode=tblMastAccount.AcCode and tblMastItem.IsOtherItem = 0", "", "ItName", txt(1).Left + Me.Left, txt(1).Top + Me.Top + 650
    gClsSearch.SearchMultiField "tblMastItem,tblMastAccount", "ItName,ItCode,iif(InDailyReport= 0,'N','Y'),AcOurCode,ItStdpack,ItStdrateper,ITUnit,ItBrokRt,ItHsn2", Array("Item", "Item", "Daily Rpt", "Mill Cd", "Std.Packing", "Rate Per", "Unit", "Brk.Rt.", "HSN"), Array(5000, 0, 1000, 2000, 1000, 1000, 1000, 1000, 1000), " tblMastItem.ItMillcode=tblMastAccount.AcCode and tblMastItem.IsOtherItem = 0", "", "ItName"
     If gClsSearch.SearchMultiRetCol(0) <> "" Then
        GProcRstOpen rstbl, "Select * from " & tblName & " where " & IdField & " = '" & gClsSearch.SearchMultiRetCol(1) & "' and " & CriteriaStr, "R"
        gClsSearch.SearchMultiRetCol(1) = ""
        ReadFields
    End If
Case 5 'List
    frmMain.mnuRptGenrSubLst_Click (5)
Case 6 'Add
    FormAction = vbDataActionAddNew
    Call GProcActivateControls(Me, True)
    Call GProcSetButtons(Me, vbDataActionAddNew)
    Call GProcClearForm(Me, rstbl, MaxNo, True)
    If FirstTimeFlag = True Then
        FirstTimeFlag = False
    Else
'        txt(1).SetFocus
    End If
    
Case 7 'Modify
    'If mURecModi = True Then
    '    SearchHelp
    '    DoEvents
    'If IsRelatedRecord = False Then
        cmdBtn_Click 4 '-- Find
        If gClsSearch.SearchMultiRetCol(0) = "" Then Exit Sub
        FormAction = vbDataActionUpdate
        Call GProcActivateControls(Me, True)
        Call GProcSetButtons(Me, vbDataActionUpdate)
        '--Related Mill & Item Record
        If IsRelatedRecord = True Then txttemp(9).Enabled = False
        txt(1).SetFocus
    'End If
Case 8 'Delete
    'If mURecDel = True Then
     'If IsRelatedRecord = False Then
        If MsgBox("Do You Really Want to Delete This Record?", vbQuestion + vbYesNo + vbDefaultButton1, "Confirm Delete") = vbYes Then
            gCn.Execute "Delete from " & tblName & " where " & IdField & " =  " & txt(0) & " and " & CriteriaStr, "O"
        End If
        cmdBtn_Click (3)
    'End If
Case 9 'Print
Case 10 'Exit
    Unload Me
Case 11 'Label
'---HSN Update
    Frame1.Visible = True
    txtOldHSN.text = ""
    txtNewHSN.text = ""
    txtOldHSN.Visible = True
    txtNewHSN.Visible = True
    txtOldHSN.Enabled = True
    txtNewHSN.Enabled = True
    cmdUpdate.Visible = True
    Label3.Visible = True
    Label2.Visible = True
    txtOldHSN.SetFocus

    'frmLabel.Show 1
Case 12 'Creation To Firm
    gfrmTypeStr = "Item"
    frmUtlCreationToFirm.Show 1
    rstbl.Requery
Case 13 'Save
    If ValidateData = True Then
        FillTxtFromTemp
        GProcRstOpen rstbl, "Select * from " & tblName & " where " & IdField & " =  " & txt(0) & " and " & CriteriaStr, "O"
        GProcSaveRecord Me, rstbl, FormAction, MaxNo
        FormAction = vbDataActionClose
        Call GProcActivateControls(Me, False)
        Call GProcSetButtons(Me, True)
        If FormAction = vbDataActionAddNew Then
            If cmdBtn(6).Visible = True Then cmdBtn(6).SetFocus
        Else
            If cmdBtn(7).Visible = True Then cmdBtn(7).SetFocus
        End If
    End If
Case 14 'Cancel
'    Call GProcClearForm(Me, rstbl, MaxNo, True)
    cmdBtn_Click (3)
    If rstbl.RecordCount <= 0 Then
        Call GProcActivateControls(Me, False)
        Call GProcSetButtons(Me, vbDataActionCancel)
    Else
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

Private Sub cmdExit_Click()
Frame1.Visible = False
cmdBtn(10).SetFocus
End Sub

Private Sub cmdUpdate_Click()
'---HSN Update
'If txtNewHSN <> "" And txtOldHSN <> "" Then
'   gCn.Execute "Update tblMastItem set ItHsn = '" & txtNewHSN & "' where ItHsn= '" & txtOldHSN & "' "
'   rstbl.Requery
'   ReadFields
'   MsgBox ("HSN Updated...")
'End If
Frame1.Visible = False

End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case 13:
            Sendkeys "{TAB}"
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
    X = "Select * from " & tblName & " where IsOtherItem=0"
    CriteriaStr = " IsOtherItem=0 "
    If Trim(OrderField) <> "" Then X = X & " order by " & OrderField
'    GProcRstOpen rstbl, X, "O"
    If Trim(mShowRec) <> "" Then
       GProcRstOpen rstbl, "Select * from " & tblName & " where IsOtherItem = 0  and  " & IdField & " = " & CLng(Trim(mShowRec)), "R"
       ReadFields
       Call GProcSetButtons(Me, vbDataActionCancel)
       mShowRec = ""
    Else
       GProcRstOpen rstbl, "Select * from " & tblName & " where IsOtherItem=0  and  " & IdField & " = 0  ", "R"
       cmdBtn_Click (6)
    End If



    '---- Button Visiblity Set As Per User Previlage
    GProcSetButtonVisiblity Me
    FirstTimeFlag = True
    cmdBtn_Click (6)
    Cbo(6).ListIndex = 0
    Cbo(7).ListIndex = 0
    Cbo(8).ListIndex = 0
    Cbo(18).ListIndex = 0
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
If FormAction = vbDataActionUpdate Then
    MsgBox "You are about to modify the record. Save the record and then exit.", vbInformation + vbOKOnly, Me.Caption
    Cancel = True
    Exit Sub
End If
'mShowRec = Space(5)
rstbl.Requery
rstbl.Close
End Sub
Private Sub txt_GotFocus(Index As Integer)
GProcSelectBox Me.ActiveControl
End Sub
Private Sub txt_LostFocus(Index As Integer)
Dim MillCode As Long
If Index = 1 Then
    Dim m As Integer
    For m = 1 To Len(txt(Index))
    If (Mid(txt(Index), m, 1)) = "'" Then
    txt(1) = Left(txt(Index), m - 1) + Mid(txt(Index), m + 1)
    End If
Next m
End If
If txt(2) = "" And txttemp(9) <> "" Then
    MillCode = GProcGetColumnValue("tblMastAccount", "AcName", txttemp(9), "S", "AcCode", "N")
    txt(2) = GProcGetColumnValue("tblMastItem", "ItMillCode", CStr(MillCode), "N", "ItTicket", "S")
End If


GProcCheckForNumber Me, rstbl, MaxNo
SaveBtnEd
End Sub
Private Sub txt_KeyPress(Index As Integer, KeyAscii As Integer)
If KeyAscii = 13 Then Exit Sub
KeyAscii = GProcValidateKey(rstbl, Index, KeyAscii, txt(Index).text)
End Sub
Private Sub txtOldHSN_KeyPress(KeyAscii As Integer)
    gClsSearch.SearchMultiField "tblMastItem", "Distinct ItHsn", Array("HSN Code"), Array(3000), " ", "", "ItHsn", txtOldHSN.Left + Me.Left, txtOldHSN.Top + Me.Top + 650
    KeyAscii = 0
    If gClsSearch.SearchMultiRetCol(0) <> "" Then
       txtOldHSN.text = gClsSearch.SearchMultiRetCol(0)
       txtNewHSN.SetFocus
    End If
End Sub

Private Sub txttemp_GotFocus(Index As Integer)
GProcSelectBox Me.ActiveControl
End Sub
Private Sub txtTemp_LostFocus(Index As Integer)
SaveBtnEd
End Sub
Private Sub txtTemp_KeyPress(Index As Integer, KeyAscii As Integer)
Dim LeftPos As Long
Dim TopPos As Long
Dim MillCode As Long
TopPos = txttemp(Index).Top + Me.Top + 650
LeftPos = txttemp(Index).Left + Me.Left
If KeyAscii = 13 Then Exit Sub
If Index = 9 Then '-- Mill Name
    gClsSearch.SearchMultiField "TblMastaccount", "AcName,AcCity,AcAlName,Accode", Array("AcName", "City", "Alias Name", "Accode"), Array(4000, 1000, 2000, 0), " (not AcName = '-') and agcode = 90029   ", "", " ", txt(1).Left + Me.Left, txt(1).Top + Me.Top + 650
    KeyAscii = 0
    If gClsSearch.SearchMultiRetCol(0) <> "" Then
        txttemp(Index) = gClsSearch.SearchMultiRetCol(0)
    End If
    MillCode = GProcGetColumnValue("tblMastAccount", "AcName", txttemp(9), "S", "AcCode", "N")
    '-- Ticket
    If txt(2) = "" Then txt(2) = GProcGetColumnValue("tblMastItem", "ItMillCode", CStr(MillCode), "N", "ItTicket", "S")
    
    ExemptMill = GProcGetColumnValue("tblMastAccount", "AcName", txttemp(9), "S", "AcIsExemptMill", "N")
    If ExemptMill = 1 Then
       txt(22).Visible = True
       Label1(1).Visible = True
    Else
       txt(22).Visible = False
       Label1(1).Visible = False
    End If
End If
End Sub
Private Sub txttemp_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
Dim GName As String
If KeyCode = vbKeyF3 Then
    Select Case Index
        Case 9 '-- mill
            GName = GProcGetColumnValue("TblMastGroup", "AgCode", 90029, "N", "AgName", "S")
            GProcShowForm frmMastAccount, frmMain.mnuMstfrm(0), Array(frmMastAccount.txttemp(2)), Array(GName)
    End Select
End If
End Sub
Private Sub SaveBtnEd()
'--- Name,Ticket,Short,Unit,Mill
If Trim(txt(1)) = "" Or Trim(txt(3)) = "" Or Trim(Cbo(6)) = "" Or Trim(txttemp(9)) = "" Then
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
    GProcCheckForNumber Me, rstbl, MaxNo
    '--- Required
    '-- ITCode
    If FormAction = vbDataActionAddNew Then txt(0) = GProcGenerateId(tblName, IdField, pStartCode:=gCSItCode)
    If gCEItCode > 0 And CLng(txt(0)) > gCEItCode Then
        MsgBox "Year End is completed. You can not create new count.", vbCritical + vbOKOnly, Me.Caption
        txt(1).SetFocus
        Exit Function
    End If
    '--- Name
    If Trim(txt(1)) = "" Then
        MsgBox "Check Item Name.", vbCritical + vbOKOnly, Me.Caption
        txt(1).SetFocus
        Exit Function
    End If
    '--- Ticket
'    If Trim(txt(2)) = "" Then
'        MsgBox "Check Item Ticket.", vbCritical + vbOKOnly, Me.Caption
'        txt(2).SetFocus
'        Exit Function
'    End If
    '---Short
    If Trim(txt(3)) = "" Then
        MsgBox "Check Item Short.", vbCritical + vbOKOnly, Me.Caption
        txt(3).SetFocus
        Exit Function
    End If
    If Val(txt(5)) = 0 Then
        MsgBox "Check Scale", vbCritical + vbOKOnly, Me.Caption
        txt(5).SetFocus
        Exit Function
    End If
    '---Unit
    If Trim(Cbo(6)) = "" Then
        MsgBox "Check Unit.", vbCritical + vbOKOnly, Me.Caption
        Cbo(6).SetFocus
        Exit Function
    End If
    If Trim(Cbo(18)) = "" Then
        MsgBox "Check Item Type", vbCritical + vbOKOnly, Me.Caption
        Cbo(18).SetFocus
        Exit Function
    End If
    '--- Mill
    If Trim(txttemp(9)) = "" Then
        MsgBox "Check Mill.", vbCritical + vbOKOnly, Me.Caption
        txttemp(9).SetFocus
        Exit Function
    End If
    '-- Create Mill
    If GProcGetColumnValue("tblMastAccount", "AcName", txttemp(9), "S", "AcCode", "N") = 0 Then
        GName = GProcGetColumnValue("TblMastGroup", "AgCode", 90029, "N", "AgName", "S")
        GProcShowForm frmMastAccount, frmMain.mnuMstfrm(0), Array(frmMastAccount.txt(1), frmMastAccount.txttemp(2)), Array(txttemp(9), GName)
        Exit Function
    End If
    '-- Max Rate
    If CDbl(txt(13)) < CDbl(txt(14)) Then
        MsgBox "Max Rate is less than Min Rate.", vbCritical + vbOKOnly, App.Title
        txt(13).SetFocus
        Exit Function
    End If
    '----Duplicate Records
    '----Addition Case
    MillCode = GProcGetColumnValue("TblMastAccount", "AcName", txttemp(9), "S", "Accode", "N")
    If FormAction = vbDataActionAddNew Then
        X = "Select ItCode,ItName From " & tblName & " Where ItMillCode=" & MillCode & " and ItName='" & Trim(txt(1).text) & "'"
        i = GProcRstOpen(rstAddCheck, X, "R")
        If i > 0 Then
            MsgBox "Record already exists  " & Trim(txt(1).text), vbCritical + vbOKOnly, "Duplicate Record"
            txt(1).SetFocus
            Exit Function
        End If
        rstAddCheck.Close
    End If
    '----Modification case
    If FormAction = vbDataActionUpdate Then
        X = "Select ItCode,ItName From " & tblName & " Where (ItMillCode=" & MillCode & " and (ItName='" & Trim(txt(1).text) & "') and (ItCode<>" & txt(0).text & "))"
        i = GProcRstOpen(rstModiCheck, X, "O")
        If i > 0 Then
            MsgBox "Record already exists  " & Trim(txt(1).text), vbCritical + vbOKOnly, "Duplicate Record"
            txt(1).SetFocus
            Exit Function
        End If
        rstModiCheck.Close
    End If
    ValidateData = True
End Function
Private Sub FillTxtFromTemp()
txt(6) = Cbo(6)
txt(7) = Cbo(7)
txt(17) = Cbo(8)
txt(18) = Cbo(18).ListIndex
txt(21) = Cbo(21).ListIndex
txt(9) = GProcGetColumnValue("TBLMASTaccount", "acname", txttemp(9), "S", "accode", "N", "", gCn)

txt(24) = chkIsExGST.Value
txt(25) = chkIsFrghtInsu.Value
txt(29) = chkInDailyReport.Value
txt(30) = chkIsInInsu.Value
If IsDate(mskDt) Then
    txt(16) = CDate(mskDt)
Else
    txt(16) = ""
End If
End Sub
Private Sub FillTempFromTxt()
Cbo(6) = txt(6)
Cbo(7) = txt(7)
Cbo(8) = txt(17)
Cbo(18).ListIndex = Val(txt(18))
Cbo(21).ListIndex = Val(txt(21))
chkIsExGST.Value = Val(txt(24))
chkIsFrghtInsu.Value = Val(txt(25))
chkInDailyReport.Value = Val(txt(29))
chkIsInInsu.Value = Val(txt(30))
txttemp(9) = GProcGetColumnValue("TBLMASTaccount", "accode", txt(9), "N", "acname", "S", "", gCn)
ExemptMill = GProcGetColumnValue("TBLMASTaccount", "accode", txt(9), "N", "AcIsExemptMill", "N", "", gCn)
If ExemptMill = 1 Then
   txt(22).Visible = True
   Label1(1).Visible = True
Else
   txt(22).Visible = False
   Label1(1).Visible = False
End If


'mskDt = txt(16) & vbNullString
If IsDate(txt(16)) Then '--Delivery Date
    mskDt = txt(26)
Else
    mskDt = "__/__/____"
End If

End Sub
Private Sub ReadFields()
GProcShowRecord Me, rstbl, MaxNo
FillTempFromTxt



End Sub
'--- Is Releted Mill & Item Record
Private Function IsRelatedRecord() As Boolean
'--- Pending
'Dim X As String
'Dim MillCode As Long
'Dim Rs1 As Recordset
'MillCode = GProcGetColumnValue("tblMastAccount", "AcName", txttemp(9), "S", "AcCode", "N")
'X = "select tblAddLess.Vno,tblAddLess.Vtype,tblAddLess.Vdt from tblAddLess,tblVouSub where " _
'& " tblAddLess.Vno=tblVouSub.Vno and tblAddLess.Vtype=tblVouSub.Vtype and tblAddLess.VYear=tblVouSub.VYear and " _
'& " tblAddLess.AdMillCode = " & MillCode & " And tblVouSub.vSubItCode=" & txt(0)
'GProcRstOpen Rs1, X, "R"
'If Rs1.EOF And Rs1.BOF Then
'Else
'    IsRelatedRecord = True
'End If
End Function


