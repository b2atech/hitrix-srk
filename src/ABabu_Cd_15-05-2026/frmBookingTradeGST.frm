VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form frmBookingTradeGST 
   BackColor       =   &H00FBF2E1&
   BorderStyle     =   0  'None
   Caption         =   "Purchase Entry"
   ClientHeight    =   9570
   ClientLeft      =   0
   ClientTop       =   15
   ClientWidth     =   13950
   FillColor       =   &H000040C0&
   FillStyle       =   0  'Solid
   Icon            =   "frmBookingTradeGST.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   9570
   ScaleWidth      =   13950
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
      Height          =   390
      Index           =   38
      Left            =   6255
      MaxLength       =   60
      TabIndex        =   23
      Text            =   "38"
      Top             =   8820
      Width           =   7230
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
      Index           =   37
      Left            =   6255
      MaxLength       =   60
      TabIndex        =   22
      Text            =   "37"
      Top             =   8370
      Width           =   7230
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
      Index           =   36
      Left            =   14985
      TabIndex        =   134
      Text            =   "36"
      Top             =   5220
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
      Height          =   360
      Index           =   35
      Left            =   14940
      TabIndex        =   133
      Text            =   "35"
      Top             =   4905
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
      Height          =   360
      Index           =   34
      Left            =   14940
      TabIndex        =   132
      Text            =   "34"
      Top             =   4545
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
      Height          =   360
      Index           =   33
      Left            =   14940
      TabIndex        =   131
      Text            =   "33"
      Top             =   4185
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
      Height          =   360
      Index           =   32
      Left            =   14805
      TabIndex        =   130
      Text            =   "32"
      Top             =   3825
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
      Height          =   360
      Index           =   31
      Left            =   14850
      TabIndex        =   129
      Text            =   "31"
      Top             =   3465
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
      Height          =   360
      Index           =   30
      Left            =   14895
      TabIndex        =   128
      Text            =   "30"
      Top             =   3105
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.TextBox txtStdPack 
      Alignment       =   1  'Right Justify
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Left            =   7245
      TabIndex        =   126
      Text            =   "txtStdPack"
      Top             =   3870
      Width           =   1245
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
      Left            =   1485
      TabIndex        =   7
      Text            =   "txtCode(7)"
      ToolTipText     =   "Press F3 for Creation"
      Top             =   2655
      Width           =   1095
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
      Index           =   5
      Left            =   1485
      TabIndex        =   5
      Text            =   "txtCode(5)"
      ToolTipText     =   "Press F3 for Creation"
      Top             =   1800
      Width           =   1095
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
      Left            =   1485
      TabIndex        =   3
      Text            =   "txtCode(6)"
      ToolTipText     =   "Press F3 for Creation"
      Top             =   1350
      Width           =   1095
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
      Left            =   12510
      Style           =   1  'Graphical
      TabIndex        =   125
      ToolTipText     =   "To Find Record"
      Top             =   2025
      Width           =   1005
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
      Left            =   13050
      Style           =   1  'Graphical
      TabIndex        =   124
      ToolTipText     =   "Next"
      Top             =   975
      Width           =   465
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
      Left            =   13050
      Style           =   1  'Graphical
      TabIndex        =   123
      ToolTipText     =   "Last"
      Top             =   1500
      Width           =   465
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
      Left            =   12510
      Style           =   1  'Graphical
      TabIndex        =   122
      ToolTipText     =   "To Modify Record"
      Top             =   3120
      Width           =   1005
   End
   Begin VB.Frame Frame1 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFC0C0&
      Caption         =   "Booking Printing"
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
      Height          =   2865
      Left            =   3195
      TabIndex        =   113
      Top             =   5580
      Visible         =   0   'False
      Width           =   6165
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
         Left            =   4605
         TabIndex        =   119
         Top             =   2010
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
         Left            =   1710
         TabIndex        =   118
         Top             =   2025
         Width           =   1335
      End
      Begin VB.TextBox txtPInvno 
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
         Height          =   405
         Index           =   1
         Left            =   2820
         MaxLength       =   7
         TabIndex        =   117
         Top             =   975
         Width           =   2055
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
         Left            =   225
         TabIndex        =   116
         Top             =   2010
         Width           =   1335
      End
      Begin VB.TextBox txtPInvno 
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
         Height          =   405
         Index           =   0
         Left            =   2820
         MaxLength       =   7
         TabIndex        =   115
         Top             =   405
         Width           =   2055
      End
      Begin VB.CommandButton cmdPrintBill 
         Caption         =   "&Mail"
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
         Left            =   3105
         TabIndex        =   114
         Top             =   2025
         Width           =   1335
      End
      Begin VB.Label Label39 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "To Booking No         :"
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
         TabIndex        =   121
         Top             =   975
         Width           =   1995
      End
      Begin VB.Label Label38 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "From Booking  No    :"
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
         TabIndex        =   120
         Top             =   495
         Width           =   2010
      End
      Begin VB.Line Line3 
         X1              =   0
         X2              =   6120
         Y1              =   1665
         Y2              =   1665
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
      Height          =   360
      Index           =   29
      Left            =   18540
      TabIndex        =   105
      Text            =   "29"
      Top             =   7515
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
      Index           =   28
      Left            =   10350
      TabIndex        =   27
      Text            =   "28"
      Top             =   2700
      Width           =   1560
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
      Index           =   27
      Left            =   9495
      TabIndex        =   26
      Text            =   "27"
      Top             =   2700
      Width           =   795
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
      Index           =   26
      Left            =   2280
      MaxLength       =   50
      TabIndex        =   17
      Text            =   "26"
      Top             =   5175
      Width           =   8760
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
      Index           =   25
      Left            =   2295
      MaxLength       =   50
      TabIndex        =   16
      Text            =   "25"
      Top             =   4725
      Width           =   8760
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
      Left            =   2700
      TabIndex        =   4
      Text            =   "txttemp(6)"
      ToolTipText     =   "Press F3 for Creation"
      Top             =   1350
      Width           =   6315
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
      Left            =   2700
      TabIndex        =   6
      Text            =   "txttemp(5)"
      ToolTipText     =   "Press F3 for Creation"
      Top             =   1800
      Width           =   6315
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
      Left            =   2700
      TabIndex        =   9
      Text            =   "txttemp(7)"
      ToolTipText     =   "Press F3 for Creation"
      Top             =   2655
      Width           =   6315
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
      Index           =   24
      Left            =   16380
      TabIndex        =   8
      Text            =   "txttemp(24)"
      ToolTipText     =   "Press F3 for Creation"
      Top             =   2160
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
      Height          =   360
      Index           =   24
      Left            =   18585
      TabIndex        =   87
      Text            =   "24"
      Top             =   7155
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.ComboBox Cbo 
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
      ItemData        =   "frmBookingTradeGST.frx":058A
      Left            =   2295
      List            =   "frmBookingTradeGST.frx":0594
      TabIndex        =   15
      Text            =   "Cbo(22)"
      Top             =   4320
      Width           =   2715
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
      Index           =   23
      Left            =   5760
      TabIndex        =   21
      Text            =   "txttemp(23)"
      ToolTipText     =   "Press F3 for Creation"
      Top             =   6075
      Width           =   5730
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
      Index           =   23
      Left            =   18540
      TabIndex        =   77
      Text            =   "23"
      Top             =   6795
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
      Height          =   360
      Index           =   16
      Left            =   18540
      TabIndex        =   74
      Text            =   "16"
      Top             =   5625
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
      Height          =   360
      Index           =   15
      Left            =   18540
      TabIndex        =   73
      Text            =   "15"
      Top             =   5265
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
      Height          =   360
      Index           =   13
      Left            =   18540
      TabIndex        =   72
      Text            =   "13"
      Top             =   4455
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
      Height          =   360
      Index           =   12
      Left            =   18540
      TabIndex        =   71
      Text            =   "12"
      Top             =   4005
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
      Index           =   9
      Left            =   18570
      TabIndex        =   70
      Text            =   "9"
      Top             =   3195
      Visible         =   0   'False
      Width           =   345
   End
   Begin VB.TextBox txt 
      Alignment       =   1  'Right Justify
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Index           =   21
      Left            =   11340
      TabIndex        =   14
      Text            =   "21"
      Top             =   3870
      Width           =   645
   End
   Begin VB.TextBox txt 
      Alignment       =   1  'Right Justify
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Index           =   20
      Left            =   9990
      TabIndex        =   13
      Text            =   "20"
      Top             =   3870
      Width           =   1275
   End
   Begin VB.TextBox txttemp 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Index           =   17
      Left            =   585
      TabIndex        =   10
      Text            =   "txttemp(17)"
      ToolTipText     =   "Press F3 for Creation"
      Top             =   3870
      Width           =   5640
   End
   Begin VB.TextBox txt 
      Alignment       =   1  'Right Justify
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Index           =   19
      Left            =   8550
      TabIndex        =   12
      Text            =   "19"
      Top             =   3870
      Width           =   1380
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
      Left            =   12510
      Style           =   1  'Graphical
      TabIndex        =   63
      Top             =   5055
      Width           =   1005
   End
   Begin VB.TextBox txt 
      Alignment       =   1  'Right Justify
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Index           =   18
      Left            =   6390
      TabIndex        =   11
      Text            =   "18"
      Top             =   3870
      Width           =   795
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
      Left            =   18540
      TabIndex        =   60
      Text            =   "22"
      Top             =   6390
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.TextBox txt 
      Alignment       =   1  'Right Justify
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Index           =   10
      Left            =   18495
      TabIndex        =   19
      Text            =   "10"
      Top             =   7830
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
      Height          =   360
      Index           =   14
      Left            =   18540
      TabIndex        =   58
      Text            =   "14"
      Top             =   4905
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.CommandButton CmdBagDetails 
      Caption         =   "Cartoon &Bag Details"
      Height          =   720
      Left            =   17415
      TabIndex        =   56
      Top             =   3285
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
      Left            =   18570
      TabIndex        =   33
      Text            =   "8"
      Top             =   2790
      Visible         =   0   'False
      Width           =   345
   End
   Begin VB.TextBox txtCrBalance 
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
      Left            =   16425
      TabIndex        =   28
      Text            =   "txtCrBalance"
      Top             =   1035
      Visible         =   0   'False
      Width           =   390
   End
   Begin VB.CommandButton cmdLedger 
      Caption         =   "Ledger"
      Height          =   375
      Left            =   16740
      TabIndex        =   55
      Top             =   5895
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.TextBox txtGrid 
      BackColor       =   &H80000000&
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
      Left            =   585
      TabIndex        =   20
      Top             =   7155
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
      Left            =   18495
      TabIndex        =   53
      Text            =   "1"
      Top             =   450
      Visible         =   0   'False
      Width           =   420
   End
   Begin VB.TextBox txt 
      Alignment       =   1  'Right Justify
      Height          =   315
      Index           =   2
      Left            =   18390
      TabIndex        =   51
      Text            =   "2"
      Top             =   900
      Visible         =   0   'False
      Width           =   525
   End
   Begin VB.TextBox txt 
      Alignment       =   1  'Right Justify
      Height          =   315
      Index           =   5
      Left            =   18390
      TabIndex        =   50
      Text            =   "5"
      Top             =   1845
      Visible         =   0   'False
      Width           =   525
   End
   Begin VB.TextBox txt 
      Alignment       =   1  'Right Justify
      Height          =   315
      Index           =   3
      Left            =   18390
      TabIndex        =   49
      Text            =   "3"
      Top             =   1260
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
      Left            =   16560
      Style           =   1  'Graphical
      TabIndex        =   45
      Top             =   6615
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
      Left            =   15255
      Style           =   1  'Graphical
      TabIndex        =   44
      ToolTipText     =   "To Add New Record"
      Top             =   5850
      Visible         =   0   'False
      Width           =   1095
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
      Left            =   12510
      Style           =   1  'Graphical
      TabIndex        =   42
      ToolTipText     =   "Exit"
      Top             =   7110
      Width           =   1005
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
      Left            =   12510
      Style           =   1  'Graphical
      TabIndex        =   38
      ToolTipText     =   "To Save Record"
      Top             =   4455
      Width           =   1005
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
      Left            =   12510
      Style           =   1  'Graphical
      TabIndex        =   43
      Top             =   5625
      Width           =   1005
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
      Left            =   12510
      Style           =   1  'Graphical
      TabIndex        =   37
      ToolTipText     =   "To Delete Record"
      Top             =   3645
      Width           =   1005
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
      Left            =   12510
      Style           =   1  'Graphical
      TabIndex        =   36
      ToolTipText     =   "To Add New Record"
      Top             =   2595
      Width           =   1005
   End
   Begin VB.CommandButton cmdBtn 
      Caption         =   "Cancel Bag"
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
      Index           =   5
      Left            =   12510
      Style           =   1  'Graphical
      TabIndex        =   39
      Top             =   6165
      Width           =   1005
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
      Left            =   12510
      Style           =   1  'Graphical
      TabIndex        =   41
      ToolTipText     =   "First"
      Top             =   1500
      Width           =   465
   End
   Begin VB.TextBox txt 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFFFF&
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
      Index           =   0
      Left            =   1890
      TabIndex        =   0
      Text            =   "0"
      Top             =   585
      Width           =   1830
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
      Left            =   12510
      Style           =   1  'Graphical
      TabIndex        =   40
      ToolTipText     =   "Previous"
      Top             =   975
      Width           =   465
   End
   Begin VB.TextBox txt 
      Alignment       =   1  'Right Justify
      Height          =   315
      Index           =   4
      Left            =   18390
      TabIndex        =   48
      Text            =   "4"
      Top             =   1530
      Visible         =   0   'False
      Width           =   525
   End
   Begin VB.TextBox txt 
      Height          =   315
      Index           =   6
      Left            =   18480
      TabIndex        =   47
      Text            =   "6"
      Top             =   2160
      Visible         =   0   'False
      Width           =   435
   End
   Begin VB.TextBox txt 
      Alignment       =   1  'Right Justify
      Height          =   315
      Index           =   7
      Left            =   18390
      TabIndex        =   46
      Text            =   "7"
      Top             =   2520
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
      Index           =   11
      Left            =   18540
      TabIndex        =   35
      Text            =   "11"
      Top             =   3645
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
      Height          =   360
      Index           =   17
      Left            =   18540
      TabIndex        =   29
      Text            =   "17"
      Top             =   5940
      Visible         =   0   'False
      Width           =   375
   End
   Begin MSMask.MaskEdBox mskDt 
      Height          =   375
      Index           =   2
      Left            =   14265
      TabIndex        =   2
      Top             =   495
      Visible         =   0   'False
      Width           =   1560
      _ExtentX        =   2752
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
      Bindings        =   "frmBookingTradeGST.frx":05A7
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
      Height          =   2985
      Left            =   360
      TabIndex        =   18
      Top             =   6075
      Width           =   4965
      _ExtentX        =   8758
      _ExtentY        =   5265
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
   Begin MSMask.MaskEdBox mskDt 
      Height          =   375
      Index           =   29
      Left            =   14265
      TabIndex        =   25
      Top             =   990
      Visible         =   0   'False
      Width           =   1560
      _ExtentX        =   2752
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
   Begin MSComCtl2.DTPicker dtpDt 
      Height          =   360
      Index           =   2
      Left            =   4680
      TabIndex        =   1
      Top             =   630
      Width           =   1665
      _ExtentX        =   2937
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
      Format          =   112984067
      CurrentDate     =   -273
   End
   Begin MSComCtl2.DTPicker dtpDt 
      Height          =   360
      Index           =   29
      Left            =   10170
      TabIndex        =   24
      Top             =   1980
      Width           =   1665
      _ExtentX        =   2937
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
      Format          =   112984067
      CurrentDate     =   -273
   End
   Begin VB.Label Label35 
      BackStyle       =   0  'Transparent
      Caption         =   "Note  :"
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
      Left            =   5580
      TabIndex        =   135
      Top             =   8415
      Width           =   735
   End
   Begin VB.Shape Shape6 
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00000080&
      FillColor       =   &H00EDFDFE&
      Height          =   1140
      Left            =   5490
      Top             =   8235
      Width           =   8205
   End
   Begin VB.Label Label22 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Std.Pack"
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
      Left            =   7650
      TabIndex        =   127
      Top             =   3375
      Width           =   825
   End
   Begin VB.Shape Shape1 
      FillColor       =   &H00FED8FC&
      FillStyle       =   0  'Solid
      Height          =   7410
      Left            =   12330
      Shape           =   4  'Rounded Rectangle
      Top             =   540
      Width           =   1350
   End
   Begin VB.Label lblContMbl 
      BackStyle       =   0  'Transparent
      Caption         =   "Mbl"
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
      Height          =   285
      Left            =   9990
      TabIndex        =   112
      Top             =   7830
      Width           =   2220
   End
   Begin VB.Label Label36 
      BackStyle       =   0  'Transparent
      Caption         =   "Mbl   :"
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
      Left            =   9225
      TabIndex        =   111
      Top             =   7785
      Width           =   690
   End
   Begin VB.Label lblContPer 
      BackStyle       =   0  'Transparent
      Caption         =   "Contact To"
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
      Height          =   285
      Left            =   6795
      TabIndex        =   110
      Top             =   6525
      Width           =   4695
   End
   Begin VB.Label Label34 
      BackStyle       =   0  'Transparent
      Caption         =   "Address  :"
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
      Left            =   5760
      TabIndex        =   109
      Top             =   6795
      Width           =   1005
   End
   Begin VB.Label Label33 
      BackStyle       =   0  'Transparent
      Caption         =   "Contact   :"
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
      Left            =   5760
      TabIndex        =   108
      Top             =   6525
      Width           =   1050
   End
   Begin VB.Label lblMblNo 
      BackStyle       =   0  'Transparent
      Caption         =   "Mbl No"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF00FF&
      Height          =   240
      Left            =   9945
      TabIndex        =   107
      Top             =   1215
      Width           =   2325
   End
   Begin VB.Label lblPan 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Pan No"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF00FF&
      Height          =   240
      Left            =   9945
      TabIndex        =   106
      Top             =   900
      Width           =   2220
   End
   Begin VB.Label Label32 
      BackStyle       =   0  'Transparent
      Caption         =   "Date  :"
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
      Left            =   9495
      TabIndex        =   104
      Top             =   2025
      Width           =   960
   End
   Begin VB.Label Label31 
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
      Height          =   375
      Left            =   11250
      TabIndex        =   103
      Top             =   2385
      Width           =   780
   End
   Begin VB.Label Label30 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Bag "
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
      Left            =   9630
      TabIndex        =   102
      Top             =   2385
      Width           =   735
   End
   Begin VB.Label Label29 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Cancel Bags Details  :"
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
      Left            =   9450
      TabIndex        =   101
      Top             =   1575
      Width           =   2070
   End
   Begin VB.Label Label28 
      BackStyle       =   0  'Transparent
      Caption         =   "Delivery Period  :"
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
      Left            =   585
      TabIndex        =   100
      Top             =   5175
      Width           =   1635
   End
   Begin VB.Label Label27 
      BackStyle       =   0  'Transparent
      Caption         =   "Extra Count      :"
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
      Left            =   585
      TabIndex        =   99
      Top             =   4770
      Width           =   1500
   End
   Begin VB.Shape Shape3 
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00000080&
      FillColor       =   &H00EDFDFE&
      Height          =   1680
      Left            =   9270
      Top             =   1530
      Width           =   2805
   End
   Begin VB.Label Label26 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "MBL  :  "
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF00FF&
      Height          =   240
      Left            =   9225
      TabIndex        =   98
      Top             =   1215
      Width           =   720
   End
   Begin VB.Label Label25 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "PAN  :  "
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF00FF&
      Height          =   240
      Left            =   9225
      TabIndex        =   97
      Top             =   900
      Width           =   705
   End
   Begin VB.Label Label24 
      BackStyle       =   0  'Transparent
      Caption         =   "&Party's :"
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
      Left            =   8100
      TabIndex        =   96
      Top             =   540
      Width           =   765
   End
   Begin VB.Label Label23 
      BackStyle       =   0  'Transparent
      Caption         =   "Dist :"
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
      Left            =   14715
      TabIndex        =   95
      Top             =   8505
      Visible         =   0   'False
      Width           =   690
   End
   Begin VB.Label Label19 
      BackStyle       =   0  'Transparent
      Caption         =   "State :"
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
      Left            =   14805
      TabIndex        =   94
      Top             =   7830
      Visible         =   0   'False
      Width           =   690
   End
   Begin VB.Label Label18 
      BackStyle       =   0  'Transparent
      Caption         =   "Tal.   :"
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
      Left            =   14850
      TabIndex        =   93
      Top             =   7425
      Visible         =   0   'False
      Width           =   600
   End
   Begin VB.Label Label15 
      BackStyle       =   0  'Transparent
      Caption         =   "Pin  :"
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
      Left            =   14805
      TabIndex        =   92
      Top             =   8145
      Visible         =   0   'False
      Width           =   555
   End
   Begin VB.Label Label14 
      BackStyle       =   0  'Transparent
      Caption         =   "City   :"
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
      Left            =   6030
      TabIndex        =   91
      Top             =   7785
      Width           =   600
   End
   Begin VB.Label lblWeight 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Weight"
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
      Left            =   2880
      TabIndex        =   90
      Top             =   9135
      Width           =   1500
   End
   Begin VB.Label lblBags 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Bags"
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
      Left            =   1350
      TabIndex        =   89
      Top             =   9135
      Width           =   1365
   End
   Begin VB.Label Label13 
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
      ForeColor       =   &H00800000&
      Height          =   375
      Left            =   315
      TabIndex        =   88
      Top             =   9135
      Width           =   1365
   End
   Begin VB.Label Label8 
      BackStyle       =   0  'Transparent
      Caption         =   "Party Gp.   :"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   330
      Left            =   14625
      TabIndex        =   86
      Top             =   2160
      Visible         =   0   'False
      Width           =   1665
   End
   Begin VB.Label lblState 
      BackStyle       =   0  'Transparent
      Caption         =   "State"
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
      Left            =   15525
      TabIndex        =   85
      Top             =   7830
      Visible         =   0   'False
      Width           =   2490
   End
   Begin VB.Label lblDist 
      BackStyle       =   0  'Transparent
      Caption         =   "Dist."
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
      Left            =   15300
      TabIndex        =   84
      Top             =   8505
      Visible         =   0   'False
      Width           =   2040
   End
   Begin VB.Label lblTal 
      BackStyle       =   0  'Transparent
      Caption         =   "Taluka"
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
      Left            =   15570
      TabIndex        =   83
      Top             =   7425
      Visible         =   0   'False
      Width           =   2535
   End
   Begin VB.Label lblAddPin 
      BackStyle       =   0  'Transparent
      Caption         =   "Pin"
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
      Left            =   15390
      TabIndex        =   82
      Top             =   8145
      Visible         =   0   'False
      Width           =   1950
   End
   Begin VB.Label lblAddCity 
      BackStyle       =   0  'Transparent
      Caption         =   "City"
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
      Left            =   6795
      TabIndex        =   81
      Top             =   7830
      Width           =   2175
   End
   Begin VB.Label lblAdd3 
      BackStyle       =   0  'Transparent
      Caption         =   "Add Line 3"
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
      Height          =   285
      Left            =   6795
      TabIndex        =   80
      Top             =   7515
      Width           =   4695
   End
   Begin VB.Label lblAdd2 
      BackStyle       =   0  'Transparent
      Caption         =   "Add Line 2"
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
      Height          =   285
      Left            =   6795
      TabIndex        =   79
      Top             =   7155
      Width           =   4695
   End
   Begin VB.Label lblAdd1 
      BackStyle       =   0  'Transparent
      Caption         =   "Add Line 1"
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
      Height          =   285
      Left            =   6795
      TabIndex        =   78
      Top             =   6795
      Width           =   4695
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "Delivery Address"
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
      Left            =   5760
      TabIndex        =   76
      Top             =   5760
      Width           =   2985
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Delivery Details"
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
      Left            =   315
      TabIndex        =   75
      Top             =   5760
      Width           =   2985
   End
   Begin VB.Line Line2 
      BorderColor     =   &H000000C0&
      X1              =   315
      X2              =   12060
      Y1              =   3735
      Y2              =   3735
   End
   Begin VB.Label Label17 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Per"
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
      Left            =   11520
      TabIndex        =   69
      Top             =   3375
      Width           =   465
   End
   Begin VB.Label Label16 
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
      Height          =   375
      Left            =   10440
      TabIndex        =   68
      Top             =   3375
      Width           =   825
   End
   Begin VB.Label Label12 
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
      Height          =   375
      Left            =   9105
      TabIndex        =   67
      Top             =   3375
      Width           =   825
   End
   Begin VB.Label Label11 
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
      Height          =   375
      Left            =   6360
      TabIndex        =   66
      Top             =   3375
      Width           =   825
   End
   Begin VB.Label Label9 
      BackStyle       =   0  'Transparent
      Caption         =   "Count "
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
      Left            =   585
      TabIndex        =   65
      Top             =   3375
      Width           =   870
   End
   Begin VB.Shape Shape5 
      BorderColor     =   &H00EFB17A&
      BorderWidth     =   5
      FillColor       =   &H00EFB17A&
      Height          =   9540
      Left            =   0
      Top             =   0
      Width           =   13905
   End
   Begin VB.Label lblMsOms 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Ms/Oms"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000080FF&
      Height          =   240
      Left            =   15435
      TabIndex        =   64
      Top             =   1710
      Width           =   735
   End
   Begin VB.Label lblTinNo 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Tin No"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF00FF&
      Height          =   240
      Left            =   9945
      TabIndex        =   62
      Top             =   555
      Width           =   2190
   End
   Begin VB.Label Label21 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "GSTIN  :  "
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF00FF&
      Height          =   240
      Left            =   9000
      TabIndex        =   61
      Top             =   555
      Width           =   900
   End
   Begin VB.Label Label20 
      BackStyle       =   0  'Transparent
      Caption         =   "Date   :    "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   255
      Left            =   4005
      TabIndex        =   59
      Top             =   675
      Width           =   870
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Booking No    :"
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
      Left            =   360
      TabIndex        =   57
      Top             =   630
      Width           =   1395
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "&Party     :"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   330
      Left            =   540
      TabIndex        =   30
      Top             =   1860
      Width           =   1035
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "M&ill       :"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   375
      Left            =   540
      TabIndex        =   32
      Top             =   1395
      Width           =   960
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "&Broker  : "
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   375
      Left            =   540
      TabIndex        =   31
      Top             =   2700
      Width           =   960
   End
   Begin VB.Label Label10 
      BackStyle       =   0  'Transparent
      Caption         =   "Tax Type         :"
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
      Left            =   585
      TabIndex        =   34
      Top             =   4365
      Width           =   1605
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
      Left            =   11070
      TabIndex        =   54
      Top             =   0
      Width           =   2475
   End
   Begin VB.Label Label6 
      BackColor       =   &H00EFB17A&
      Caption         =   "     Booking ( Trade Purchase )"
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
      TabIndex        =   52
      Top             =   45
      Width           =   13860
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H00000080&
      FillColor       =   &H00EDFDFE&
      FillStyle       =   0  'Solid
      Height          =   2400
      Left            =   315
      Top             =   3330
      Width           =   11760
   End
   Begin VB.Shape Shape4 
      BorderColor     =   &H00000080&
      FillColor       =   &H00EDFDFE&
      FillStyle       =   0  'Solid
      Height          =   1995
      Left            =   315
      Top             =   1215
      Width           =   8835
   End
End
Attribute VB_Name = "frmBookingTradeGST"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Const tblName As String = "tblBooking"
Const IdField As String = "Vno"
Const OrderField As String = "Vdt,Vno"
Const MaxNo As Long = 38
Const OtherTxtFromNo As Long = 0 '-- Other Txt Ctrl From No
Const OtherTxtToNo As Long = 0 '-- Other Txt Ctrl Less No
Dim rstbl As Recordset '-- TblPurch
Dim rstblSub As Recordset '-- TblPurchSub
Dim varBookMark As Variant
Dim FirstTimeFlag As Boolean, Criteria As String
Dim OldgFrmTypeStr As String
Dim FormAction As Integer
Dim frmTypeStr As String
Dim DelNote As Boolean
Dim i As Long
Dim TotBag As Long
Dim TotWt As Double
Dim OldBag As Integer
Dim OldBag2 As Integer
Dim mLastEntryDate As Date
Dim mLastEntryMill As String
Dim mLastEntryMillOurCd As String
Dim mLastItem As String
Dim mLastRate As Double
Dim mLastRatePer As Double
Dim mStdPack As Double
Dim Mon As Integer
Private Sub Cbo_KeyPress(Index As Integer, KeyAscii As Integer)
If Index = 23 Then KeyAscii = 0 '--Rate Type
End Sub
'Dim GridFlag As Boolean
Private Sub cmdBtn_Click(Index As Integer)
'On Error GoTo ErrorRoutine
Select Case Index
Case 0 '---- Previous
     If GProcRstOpen(rstbl, "Select * from " & tblName & " where " & Criteria & "  and " & IdField & " = (Select max(" & IdField & ") from " & tblName & " where " & Criteria & " And " & IdField & "  < " & Val(txt(0)) & ")", "R") > 0 Then ReadFields Else MsgBox "This is First Record:", vbOKOnly + vbInformation
Case 1 '---- Next
      If GProcRstOpen(rstbl, "Select * from " & tblName & " where " & Criteria & "  and   " & IdField & " = (Select min(" & IdField & ") from " & tblName & " where " & IdField & " > '" & txt(0) & "' and " & Criteria & ") and " & Criteria, "R") > 0 Then ReadFields Else MsgBox "This is last Record:", vbOKOnly + vbInformation
Case 2 '---- First
         If GProcRstOpen(rstbl, "Select * from " & tblName & " where " & Criteria & "  and  " & IdField & " = (Select min(" & IdField & ") from " & tblName & " where " & IdField & " <> '' and " & Criteria & ")and " & Criteria, "R") > 0 Then ReadFields
Case 3 '---- Last
        If GProcRstOpen(rstbl, "Select * from " & tblName & " where " & Criteria & "  and  " & IdField & " = (Select max(" & IdField & ") from " & tblName & " where " & IdField & " <> '' and " & Criteria & ")and " & Criteria, "R") > 0 Then ReadFields Else GProcClearForm Me, rstbl, MaxNo
Case 4 'Find ' QryMastAccount
     gClsSearch.SearchMultiField "qryBookingMillHelp", "Vno,PartyCode,Party,Vdt,MillCode,Mill", Array("Vno", "Party Code", "Party Name", "Vdt", "Mill Code", "Mill Name"), Array(txt(0).Width, 1500, 4000, 1500, 1500, 4000), " vtype='" & frmTypeStr & "'  and VFirm = '" & gCCode & "'  and VYear=" & gCYear, "", "Vno,Vdt", 0, 0
     If gClsSearch.SearchMultiRetCol(0) <> "" Then
        GProcRstOpen rstbl, "Select * from " & tblName & " where " & IdField & " = " & gClsSearch.SearchMultiRetCol(0) & " and " & Criteria, "R"
        gClsSearch.SearchMultiRetCol(1) = ""
        ReadFields
    End If
Case 5 'Cancel Bags
    gClsSearch.SearchMultiField "tblBooking,TblMastAccount", "Vno,AcName,Vdt", Array("Vno", "Party", "Vdt"), Array(txt(0).Width, 4000, 1080), " tblBooking.BkParty=TblMastAccount.AcCode and vtype='" & frmTypeStr & "'  and VFirm = '" & gCCode & "' and BkSrNo = 1  and VYear=" & gCYear, "", "Vno,Vdt", 0, 0
    If gClsSearch.SearchMultiRetCol(0) <> "" Then
        GProcRstOpen rstbl, "Select * from " & tblName & " where " & IdField & " = " & gClsSearch.SearchMultiRetCol(0) & " and " & Criteria, "R"
        gClsSearch.SearchMultiRetCol(1) = ""
        gClsSearch.SearchMultiRetCol(0) = ""
        ReadFields
    End If
    FormAction = vbDataActionUpdate
    Call GProcActivateControls(Me, True)
    Call GProcSetButtons(Me, vbDataActionUpdate)
    dtpDt(29).Enabled = True
    txt(27).Enabled = True
    txt(28).Enabled = True
    dtpDt(29).SetFocus
Case 6 'Add
    lblTinNo.Caption = ""
    lblPan.Caption = ""
    lblMblNo.Caption = ""
    FormAction = vbDataActionAddNew
    Call GProcActivateControls(Me, True)
    Call GProcSetButtons(Me, vbDataActionAddNew)
    Call GProcClearForm(Me, rstbl, MaxNo, True, OtherTxtFromNo, OtherTxtToNo, True)
    SetControlEd
    If FirstTimeFlag = True Then
        FirstTimeFlag = False
    Else
'        dtpDt(2).SetFocus
    End If
    txt(0) = GProcGenerateId("TblBooking", "Vno", "Vtype='" & frmTypeStr & "' and VFirm = '" & gCCode & "'   and VYear=" & gCYear)
'    dtpDt(2) = GProcGenerateVDt("tblBooking", "Vdt", "Vtype='" & frmTypeStr & "' and VFirm = '" & gCCode & "'   and VYear=" & gCYear)
'    dtpDt(2) = Format(Now, "dd/MM/yyyy")
    If Year(mLastEntryDate) = 1899 Then
        dtpDt(2) = Format(Now, "dd/MM/yyyy")
    Else
        dtpDt(2) = mLastEntryDate
    End If
    
    FillCombo
    SetGrid
    If txt(0).Visible = True Then txt(0).SetFocus
Case 7 'Modify
    'If mURecModi = True Then
    '    SearchHelp
    '    DoEvents
        If IsRelatedRecord = False Then
            GProcRstOpen rstbl, "Select * from " & tblName & " where " & IdField & " = " & txt(0) & " and " & Criteria, "R"
            ReadFields
            FormAction = vbDataActionUpdate
            Call GProcActivateControls(Me, True)
            SetControlEd
            Call GProcSetButtons(Me, vbDataActionUpdate)
            txt(0).Enabled = False
            dtpDt(2).SetFocus
        End If
    'End If
Case 8 'Delete
      If IsRelatedRecord = False Then
        If MsgBox("Do You Really Delete This Record?", vbQuestion + vbYesNo + vbDefaultButton1, "Confirm Delete") = vbYes Then
           DeleteRelatedRecord
           cmdBtn_Click (3)
        End If
    End If
Case 9 'Print
    Frame1.Caption = "Booking Confirmation Printing "
        '-- Inv No
    txtPInvno(0) = txt(0).text
    txtPInvno(1) = txt(0).text
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
        CalAmount
        FillTxtFromTemp
        SaveRelatedRecord
        mLastEntryDate = dtpDt(2)
        mLastEntryMill = txttemp(6).text
        mLastEntryMillOurCd = txtCode(6).text
        mLastItem = txttemp(17).text
        mLastRate = txt(20)
        mLastRatePer = txt(21)
        mStdPack = txtStdPack
        FormAction = vbDataActionClose
        Call GProcActivateControls(Me, False)
        Call GProcSetButtons(Me, True)
        msGrid.Col = 0 '-----
        msGrid.Row = 1 '-----
        txtGrid.Move msGrid.CellLeft + msGrid.Left, msGrid.CellTop + msGrid.Top, msGrid.CellWidth, msGrid.CellHeight '-----
'        If FormAction = vbDataActionAddNew Then
            If cmdBtn(6).Visible = True Then cmdBtn(6).SetFocus
 '       Else
  '          If cmdBtn(7).Visible = True Then cmdBtn(7).SetFocus
   '     End If
    End If
Case 14 'Cancel
    cmdBtn_Click (3)
    DoEvents
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
    FormAction = vbDataActionCancel
    If msGrid.Enabled = True Then
       msGrid.Col = 0
       msGrid.Row = 1
       txtGrid.Move msGrid.CellLeft + msGrid.Left, msGrid.CellTop + msGrid.Top, msGrid.CellWidth, msGrid.CellHeight
    End If
End Select
Exit Sub
ErrorRoutine:
    rstbl.CancelUpdate
    GProcErrorHandler
End Sub
Private Sub SetControlEd()
'If FormAction = vbDataActionAddNew Then
'   txt(0).Enabled = False
'End If
txtCrBalance.Enabled = False
txt(10).Enabled = False
dtpDt(29).Enabled = False
txt(27).Enabled = False
txt(28).Enabled = False
txt(21).Enabled = False
'If Trim(txtCode(5)) <> "" Then
'   txttemp(5).Enabled = False
'Else
'   txttemp(5).Enabled = True
'End If
'If Trim(txtCode(6)) <> "" Then
'   txttemp(6).Enabled = False
'Else
'   txttemp(6).Enabled = True
'End If
'If Trim(txtCode(7)) <> "" Then
'   txttemp(7).Enabled = False
'Else
'   txttemp(7).Enabled = True
'End If
End Sub
Private Sub DeleteRelatedRecord()
''--- Voucher Sub
gCn.Execute "delete from tblBooking where vtype='" & frmTypeStr & "' and vno=" & txt(0) & " and Vfirm = '" & gCCode & "'  and " _
& " VYear=" & gCYear
End Sub
Private Sub SaveRelatedRecord()
Dim i As Long
Dim MillCode As Long
Dim PartyCode As Long
Dim BrokerCode As Long
Dim rstSub As Recordset
Dim ItemCode As Long
Dim X As String, wit As Long, wgd As Long, Y As String, maxy As Long
Dim BillDt
Dim DueDt
Dim RowIndex As Integer
RowIndex = 1
With msGrid
    Do While RowIndex < .rows
       If (Val(.TextMatrix(RowIndex, 2)) = 0) And .rows > 2 Then
          .RemoveItem (RowIndex)
       End If
       RowIndex = RowIndex + 1
    Loop
End With
'--- Account Code
PartyCode = GProcGetColumnValue("tblMastAccount", "AcName", txttemp(5), "S", "AcCode", "N")
BrokerCode = GProcGetColumnValue("tblMastAccount", "AcName", txttemp(7), "S", "AcCode", "N")

MillCode = GProcGetColumnValue("tblMastAccount", "AcName", txttemp(6), "S", "AcCode", "N")
ItemCode = GProcGetColumnValue("tblMastItem", "ItName", txttemp(17), "S", "Itcode", "N", "ItMillCode = " & MillCode & "")
'-- Delete Records
DeleteRelatedRecord
'--- Save Grid Records
'--- TblVouSub
X = "select * from tblBooking where vtype='" & frmTypeStr & "' and vno=" & txt(0) & " and Vfirm = '" & gCCode & "' and " _
& " VYear=" & gCYear
GProcRstOpen rstSub, X, "O"
'maxy = GProcGetColumnValue("tblPurchSub", "Vno", txt(0), "N", "max(PurSubItCtrlNo)", "N", Criteria)
With msGrid
    i = 1
    Do While i < .rows
 '       If Val(.TextMatrix(i, 2)) > 0 Then
            rstSub.AddNew
            rstSub!Vno = CLng(txt(0)) '--Vno(0)
            rstSub!Vtype = txt(1) & vbNullString  '--Vtype(1)
            rstSub!Vdt = CDate(txt(2)) '--Vdt(2)
            rstSub!VYear = txt(3) '--VYear
            rstSub!VFirm = gCCode '--VFirm
            rstSub!BkParty = PartyCode
            rstSub!BkMillCode = MillCode
            rstSub!BkBroker = BrokerCode
            rstSub!BkBroker = BrokerCode
            rstSub!BkRefNo = txt(8) & vbNullString
            rstSub!BkItCode = ItemCode
            rstSub!BkBag = Val(txt(18))
            rstSub!BkWt = Val(txt(19))
            rstSub!BkRt = Val(txt(20))
            rstSub!BkRtPer = Val(txt(21))
            rstSub!BkRateType = Val(txt(22))
            rstSub!BkDeleAdd = Val(txt(23))
            rstSub!BkPartyGroup = Val(txt(24))
            rstSub!BkExtraCount = txt(25) & vbNullString
            rstSub!BkNote = txt(26) & vbNullString
            rstSub!BkCbag = Val(txt(27))
            rstSub!BkCWt = Val(txt(28))
            If IsDate(dtpDt(29)) = True And Year(dtpDt(29)) > 1899 Then
               rstSub!BkCDate = CDate(txt(29))
            End If
            rstSub!BkSrNo = i '--Sr No
            If CLng(.TextMatrix(i, 2)) > 0 Then
               rstSub!BkShDate = CDate(.TextMatrix(i, 1)) '--SH.Date
            Else
               rstSub!BkShDate = Null  '--Des.Date
            End If
            rstSub!BkShBag = CLng(.TextMatrix(i, 2)) '--Bag(2)
            rstSub!BkShWt = CDbl(.TextMatrix(i, 3)) '-- Wt(3)
            rstSub!BkTmpChr1 = txt(37) & vbNullString
            rstSub!BkTmpChr2 = txt(38) & vbNullString
            rstSub.Update
'        End If
        i = i + 1
     Loop
 End With
 rstSub.Close
End Sub

Private Sub cmdPrintBill_Click(Index As Integer)
Dim RsSA As Recordset
Dim X As String
Dim XX As String
Dim BillSrNo As String
Dim InvNo As Long
Dim Y As String
Dim BilNo As String
Dim YY As String
Dim tmpVno As Long
Dim tmpVtype As String
Dim tmpVyear As Long
Dim tmpVFirm As String

Select Case Index
Case 0, 1, 3:
    '------ Validate
    If IsNumeric(txtPInvno(0)) = False Then txtPInvno(0) = 0
    If IsNumeric(txtPInvno(1)) = False Then txtPInvno(1) = txtPInvno(0)
    If CLng(txtPInvno(1)) < CLng(txtPInvno(0)) Then txtPInvno(1) = txtPInvno(0)
    '-------
    If Index = 0 Then GProcCrystalRptPreparation cryReport, "Printer"
    If Index = 1 Then GProcCrystalRptPreparation cryReport, "Window"
    If Index = 3 Then GProcCrystalRptPreparation cryReport, "Window"
    
    cryReport.WindowTitle = "Booking"
    '-- A/C Bill SrNo
    InvNo = CLng(txtPInvno(0))
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
    Y = X
    X = "Select a.Vno,a.Vtype,a.Vdt,a.Vyear,a.VFirm,BkParty,BkMillCode,BkBroker,BkRefNo,BkRefDt,BkSubAmt,BkBillAmt,BkRoff,DueDays,DueDate,BkEntDt,'" & gUserName & "',BkItCode,BkBag,BkWt,BkRt,BkRtPer,BkRateType,BkDeleAdd,BkPartyGroup,BkExtraCount,BkNote,BkCbag,BkCWt," _
        & " BkCDate,BkTmpNo1,BkTmpNo2,BkTmpNo3,BkTmpChr1,BkTmpChr2,BkTmpChr3,BkShBag,BkShWt,BkShDate " _
        & " From tblBooking A  where a.vno in " & X & " and a.Vtype='" & frmTypeStr & "' and a.VFirm = '" & gCCode & "' And a.VYear = " & gCYear
     
   gCn.Execute "Delete from TmpBooking where  BkUser='" & gUserName & "'"
     
   gCn.Execute " insert into TmpBooking (Vno,Vtype,Vdt,Vyear,VFirm,BkParty,BkMillCode,BkBroker,BkRefNo,BkRefDt,BkSubAmt,BkBillAmt,BkRoff,DueDays,DueDate,BkEntDt,BkUser,BkItCode,BkBag,BkWt,BkRt,BkRtPer,BkRateType,BkDeleAdd,BkPartyGroup,BkExtraCount,BkNote,BkCbag,BkCWt," _
              & " BkCDate,BkTmpNo1,BkTmpNo2,BkTmpNo3,BkTmpChr1,BkTmpChr2,BkTmpChr3,BkSubBag,BkSubWt,BkSubDDate ) " & X
     
     
     
     X = "select * from TmpBooking where BkUser='" & gUserName & "'"
     With cryReport
        .Formulas(10) = "wCSTNo = '" & gCCST & "'"
        .Formulas(11) = "wBSTNo = '" & gCBST & "'"
        X = "{TmpBooking.BkUser}='" & gUserName & "' "
       .SelectionFormula = X
    '    .SelectionFormula = "{TmpBooking.Vno} = " & RsSA!Vno & " and {TmpBooking.VType}='" & RsSA!Vtype & "' and {TmpBooking.VFirm} = '" & RsSA!VFirm & "' and  {TmpBooking.VYear}=" & RsSA!VYear
        .ReportFileName = gReportPath & "rptBookingPrint_Trade.rpt"
         If Index = 0 Or Index = 1 Then
           .Action = 1
         End If
     End With
     
     
     
     
     
     
'''     If GProcRstOpen(RsSA, X, "R") <= 0 Then
'''        MsgBox "No Records..."
'''        Exit Sub
'''     Else
'''        RsSA.MoveFirst
'''        Do While Not RsSA.EOF
'''           With cryReport
'''              .Formulas(10) = "wCSTNo = '" & gCCST & "'"
'''              .Formulas(11) = "wBSTNo = '" & gCBST & "'"
'''              .SelectionFormula = "{TmpBooking.Vno} = " & RsSA!Vno & " and {TmpBooking.VType}='" & RsSA!Vtype & "' and {TmpBooking.VFirm} = '" & RsSA!VFirm & "' and  {TmpBooking.VYear}=" & RsSA!VYear
'''              .ReportFileName = gReportPath & "rptBookingPrint_MillBill.rpt"
''''              .ReportFileName = gReportPath & "xx.rpt"
'''              If Index = 0 Or Index = 1 Then
'''                 .Action = 1
''''              Else
''''                 '-- For Mail
''''                 Dim Tattach As String
''''                 Dim EmailId As String
''''                 Dim MailToParty As String
''''                 Dim Subject As String
''''                 Dim Massage As String
''''
'''''                 EmailId = "sanjivtavate@gmail.com"
'''''                 MailToParty = GProcGetColumnValue("tblMastAccount", "AcCode", RsSA!BkSCode, "N", "AcName", "S")
''''                 Subject = "Order Confirmation From " + gCName
''''                 Massage = "Dear Sir, <br> <br>"
''''                 Massage = Massage + "<font color=red>Please see the attachment <br>"
''''                 Massage = Massage + "<font size=3> <font color=blue> Line 2 -------------- "
''''
''''
''''                 '------ Export To PDF
''''                .SelectionFormula = "{tblBookingTmp.BkInvNo} = '" & RsSA!BkInvNo & "' and {tblBookingTmp.VType}='" & RsSA!Vtype & "' and {tblBookingTmp.VBranch} = '" & RsSA!VBranch & "' and {tblBookingTmp.BkFirm} = '" & RsSA!BkFirm & "' and  {tblBookingTmp.VYear}=" & RsSA!VYear
''''
''''                 Tattach = gReportPath & "rptBookingPrint.rpt"
'''''                 Tattach = gReportPath & "xx.rpt"
''''
''''                 Report_ExportToPDF frmMain.cryReport, gReportPath & "rptBookingPrint.rpt"
''''
''''                 If chkMailSupplier.Value = 1 Then
''''                    MailToParty = GProcGetColumnValue("tblMastAccount", "AcCode", RsSA!BkSCode, "N", "AcName", "S")
''''                    EmailId = GProcGetColumnValue("tblMastAccount", "AcCode", RsSA!BkSCode, "N", "AcEmail", "S")
''''                    SendEmail frmMain.cryReport, Tattach, EmailId, MailToParty, Subject, Massage
''''                 End If
''''                 If chkMailBuyer.Value = 1 Then
''''                    MailToParty = GProcGetColumnValue("tblMastAccount", "AcCode", RsSA!BkBCode, "N", "AcName", "S")
''''                    EmailId = GProcGetColumnValue("tblMastAccount", "AcCode", RsSA!BkBCode, "N", "AcEmail", "S")
''''                    SendEmail frmMain.cryReport, Tattach, EmailId, MailToParty, Subject, Massage
''''                 End If
''''                 Kill Mid(Tattach, 1, Len(Tattach) - 4) & ".pdf"
'''              End If
''''              Exit Sub
'''              tmpVno = RsSA!Vno
'''              tmpVtype = RsSA!Vtype
'''              tmpVyear = RsSA!VYear
'''              tmpVFirm = RsSA!VFirm
'''              Do While Not RsSA.EOF And RsSA!Vno = tmpVno And RsSA!Vtype = tmpVtype And RsSA!VYear = tmpVyear And RsSA!VFirm = tmpVFirm
''''                 If RsSA!Vno = tmpVno And RsSA!Vtype = tmpVtype And RsSA!VYear = tmpVyear And RsSA!VFirm = tmpVFirm Then
'''                    RsSA.MoveNext
'''                    If RsSA.EOF Then
'''                       Exit Sub
'''                    End If
''' '                End If
'''              Loop
'''           End With
'''        Loop
'''     End If
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
    Me.Top = 300
    Me.Left = Screen.Width / 2 - Me.Width / 2
    '--- TblAddLess
    mLastEntryDate = vbNull
    frmTypeStr = gfrmTypeStr
    X = "Select * from " & tblName & " where Vtype='" & frmTypeStr & "' and VFirm  = '" & gCCode & "' and VYear=" & gCYear
    If Trim(OrderField) <> "" Then X = X & " order by " & OrderField
    Criteria = "vtype='" & frmTypeStr & "' and Vfirm = '" & gCCode & "'  and   VYear=" & gCYear
    'GProcRstOpen rstbl, X, "O"
    If Trim(mShowRec) <> "" Then
       GProcRstOpen rstbl, "Select * from " & tblName & " where  " & Criteria & " and  " & IdField & " = " & CLng(Trim(mShowRec)), "R"
       ReadFields
       Call GProcSetButtons(Me, vbDataActionCancel)
       mShowRec = ""
    Else
       GProcRstOpen rstbl, "Select * from " & tblName & " where " & Criteria & " and  " & IdField & " = 0", "R"
       cmdBtn_Click (6)
    End If
    
    
    
    
    FirstTimeFlag = True
    '---Grid
    SetGrid
    SetRelatedRS
    '---- Button Visiblity Set As Per User Previlage
    GProcSetButtonVisiblity Me
    cmdBtn_Click (6)
End Sub
Private Sub SetRelatedRS()
'Dim X As String
'Set rsOS = New Recordset
'Set rsVou = New Recordset
''---TblVoucher
'X = "Select * from tblVoucher where Vtype='" & frmTypeStr & "' and vno=" & txt(0) & " and VYear=" & gCYear & " order by VCtrNo"
'GProcRstOpen rsVou, X, "O"
''---TblOutStanding
'X = "Select * from tblOutStanding where Vtype='" & frmTypeStr & "' and vno=" & txt(0) & " and VYear=" & gCYear
'GProcRstOpen rsOS, X, "O"
End Sub
Private Sub SetGrid()
Dim X As String
Set rstblSub = New Recordset
X = "select BkSrNo,BkShDate,BkShBag,BkShWt" _
& "  from tblBooking where " _
& "  tblBooking.vtype='" & frmTypeStr & "' and tblBooking.vno=" & txt(0) & " and tblBooking.VFirm = '" & gCCode & "' and " _
& "  tblBooking.VYear=" & gCYear & "  order by BkSrNo"
GProcRstOpen rstblSub, X, "R"
With msGrid
   .FormatString = ">Sr|<Date            |>Bag        |>Weight            "
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
rstbl.Close
End Sub
Private Sub MSGrid_LostFocus()
CalAmount
End Sub
Private Sub dtpDt_GotFocus(Index As Integer)
GProcSelectBox Me.ActiveControl
End Sub
Private Sub dtpDt_LostFocus(Index As Integer)
If Index = 2 Then
    If GProcIsDateValid(dtpDt(Index)) = False Then
        dtpDt(Index).SetFocus
        Exit Sub
    End If
End If
'If IsDate(mskDt(Index)) = False Then mskDt(Index) = "__/__/____"
End Sub
Private Sub txt_GotFocus(Index As Integer)
If Index = 18 Then
   OldBag2 = Val(txt(18))
End If

GProcSelectBox Me.ActiveControl
End Sub
Private Sub CheckForNumber()
GProcCheckForNumber Me, rstbl, MaxNo, pOtherTxtFromNo:=OtherTxtFromNo, pOtherTxtToNo:=OtherTxtToNo
End Sub

Private Sub txt_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
Dim OldFrmType As String
Dim LeftPos As Long
Dim TopPos As Long
If KeyCode = vbKeyF3 Then '--F3 Create
    Select Case Index
        Case 37, 38 '--Narration
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
        Case 37, 38 '--Narration
'            gClsSearch.SearchMultiField "tblMastNarration", "Narration", Array("Narration"), Array(txt(Index).Width), "NarrType='N'", "", "Narration", LeftPos, TopPos, False
            gClsSearch.SearchMultiField "tblMastNarration", "Narration", Array("Narration"), Array(txt(Index).Width), "NarrType='N' and Narration <> '-' ", "", "Narration", , , False
            If gClsSearch.SearchMultiRetCol(0) <> "" Then txt(Index) = gClsSearch.SearchMultiRetCol(0)
    End Select
End If
End Sub

Private Sub txt_LostFocus(Index As Integer)
Dim tmpVno As String
If Index = 0 Then
   If GProcRstOpen(rstbl, "Select * from " & tblName & " where " & IdField & " = " & txt(0) & " and " & Criteria, "R") > 0 Then
      ReadFields
      FormAction = vbDataActionUpdate
      Call GProcActivateControls(Me, True)
      SetControlEd
      Call GProcSetButtons(Me, vbDataActionUpdate)
   Else
      lblTinNo.Caption = ""
      lblPan.Caption = ""
      lblMblNo.Caption = ""
      tmpVno = txt(0)
      Call GProcActivateControls(Me, True)
      Call GProcSetButtons(Me, vbDataActionAddNew)
      Call GProcClearForm(Me, rstbl, MaxNo, True, OtherTxtFromNo, OtherTxtToNo, True)
      txt(0) = tmpVno
      SetControlEd
      If mLastEntryMill <> "" And FormAction = vbDataActionAddNew Then
         txttemp(6) = mLastEntryMill
         txtCode(6) = mLastEntryMillOurCd
         txttemp(17).text = mLastItem
         txt(20) = mLastRate
         txt(21) = mLastRatePer
         txtStdPack = mStdPack
      End If
   End If
End If
CheckForNumber
CalAmount
SaveBtnEd
If Index = 26 Then
   msGrid.Col = 1
   msGrid_EnterCell
End If
End Sub
Private Sub txt_KeyPress(Index As Integer, KeyAscii As Integer)
If KeyAscii = 13 Then Exit Sub
KeyAscii = GProcValidateKey(rstbl, Index, KeyAscii, txt(Index).text)
End Sub
Private Sub txtCode_GotFocus(Index As Integer)
If Index = 5 And txtCode(5) = "" Then
   txtCode(5) = "D"
   txtCode(5).SelStart = 2
End If

End Sub

Private Sub txtCode_LostFocus(Index As Integer)
If Trim(txtCode(Index)) = "" Then Exit Sub
Dim BrkAcCode As Long
SetControlEd
Select Case Index
    Case 5  '-- Party  Code
         txttemp(5) = GProcGetColumnValue("tblMastAccount", "AcOurCode", UCase(txtCode(5)), "S", "AcName", "S")
         '--- Broker From Master
         If txttemp(7) = "" Then
            BrkAcCode = GProcGetColumnValue("TblMastAccount", "ACName", txttemp(Index), "S", "AcBrkCode", "N")
            If BrkAcCode = 1 Then
               txttemp(7) = ""
               txtCode(7) = ""
            Else
               txttemp(7) = GProcGetColumnValue("TblMastAccount", "ACCode", CStr(BrkAcCode), "N", "ACname", "S")
               txtCode(7) = GProcGetColumnValue("TblMastAccount", "ACCode", CStr(BrkAcCode), "N", "AcOurCode", "S")
            End If
         End If
         If Trim(txttemp(5)) = "" Then
            MsgBox ("Code Not Found ...")
            txttemp(5) = ""
            txttemp(5).Enabled = True
            txttemp(5).SetFocus
         Else
            txtCode(7).SetFocus
         End If
    Case 6  '-- Mill  Code
        Dim MilTpCd  As Long
        Dim X As String
        Dim tmpRst As Recordset
         txttemp(6) = GProcGetColumnValue("tblMastAccount", "AcOurCode", UCase(txtCode(6)), "S", "AcName", "S")
         If Trim(txttemp(6)) = "" Then
            MsgBox ("Code Not Found ...")
            txttemp(6) = ""
            txttemp(6).Enabled = True
            txttemp(6).SetFocus
         Else
            txtCode(5).SetFocus
         End If
         If mLastEntryMill <> txttemp(6) And FormAction = vbDataActionAddNew Then
            txttemp(17) = ""
            txtStdPack = 0
            txt(20) = 0
         End If
          MilTpCd = GProcGetColumnValue("tblMastNarration", "Narration", "Trade", "S", "Narrcode", "N")
          X = "Select Acname from tblMastAccount,tblMastBillSerial where tblMastaccount.Accode = tblMastBillSerial.MillCode and  tblMastBillSerial.SaleType = " & MilTpCd & "  and tblMastBillSerial.Vfirm = '" & gCCode & "'   and  tblMastaccount.AcOurCode = '" & txtCode(6) & "'   "
          i = GProcRstOpen(tmpRst, X, "R")
          If i = 0 Then
              MsgBox ("Mill Not in Firm")
              txttemp(6).text = ""
              txtCode(6).SetFocus
              Exit Sub
          End If
    Case 7  '-- Broker  Code
         txttemp(7) = GProcGetColumnValue("tblMastAccount", "AcOurCode", UCase(txtCode(7)), "S", "AcName", "S")
         If Trim(txttemp(7)) = "" Then
            MsgBox ("Code Not Found ...")
            txttemp(7) = ""
            txttemp(7).Enabled = True
            txttemp(7).SetFocus
         Else
            If txttemp(17).Enabled = True Then txttemp(17).SetFocus
         End If
End Select
SaveBtnEd
End Sub
Private Sub txttemp_GotFocus(Index As Integer)
GProcSelectBox Me.ActiveControl
'Call txtTemp_KeyPress(Index, 32)
End Sub
Private Sub txttemp_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
Dim GName As String
Dim OldFrmType As String
If KeyCode = vbKeyF3 Then
    Select Case Index
        Case 5 '-- Party
            GName = GProcGetColumnValue("TblMastGroup", "AgCode", 90015, "N", "AgName", "S")
            GProcShowForm frmMastAccount, frmMain.mnuMstfrm(0), Array(frmMastAccount.txttemp(2)), Array(GName)
        Case 6 '-- millcode
            GName = GProcGetColumnValue("TblMastGroup", "AgCode", 90029, "N", "AgName", "S")
            GProcShowForm frmMastAccount, frmMain.mnuMstfrm(0), Array(frmMastAccount.txttemp(2)), Array(GName)
        Case 7 '-- Broker
           GName = GProcGetColumnValue("TblMastGroup", "AgCode", 90016, "N", "AgName", "S")
           GProcShowForm frmMastAccount, frmMain.mnuMstfrm(0), Array(frmMastAccount.txttemp(2)), Array(GName)
        Case 17 '-- item
           GProcShowForm frmMastItem, frmMain.mnuMstfrm(3), Array(frmMastItem.txttemp(9)), Array(txttemp(6))
        Case 24 '-- Party Group
            GProcShowForm frmMastCompGroup, frmMain.mnuMstfrm(2)
        Case 23 '-- Dele address
            GProcShowForm frmMastDelAdd, frmMain.mnuMstfrm(12)
    End Select
End If
End Sub
Private Sub txtTemp_KeyPress(Index As Integer, KeyAscii As Integer)
Dim LeftPos As Long
Dim TopPos As Long
Dim AgCode As Long
Dim wPartyCode As Long
Dim BrkAcCode As Long
Dim MillCode As Long
Dim ShowMill As String
Dim MilTpCd As Long

TopPos = txttemp(Index).Top + Me.Top + 650
LeftPos = txttemp(Index).Left + Me.Left
If KeyAscii = 13 Then Exit Sub
Select Case Index
    Case 5  '-- Party   Alias Account
'        gClsSearch.SearchMultiField "tblMastaccount", "AcOurCode,AcName,AcContPer,AcCity,AcBST,AcPAN,AcContNo", Array("Code", "Party Name", "Contact Person", "City", "", "", ""), Array(1500, 5000, 3000, 2000, 0, 0, 0), "(tblMastaccount.Agcode in (90017) ) ", Chr(KeyAscii), "", LeftPos, TopPos, True
        gClsSearch.SearchMultiFieldAccMaster Chr(KeyAscii), LeftPos, TopPos
        KeyAscii = 0
        lblTinNo.Caption = ""
        lblMsOms.Caption = ""
        lblPan.Caption = ""
        lblMblNo.Caption = ""
        txtCode(5) = ""
        If gClsSearch.SearchMultiRetCol(0) <> "" Then
            txttemp(Index) = gClsSearch.SearchMultiRetColAccMaster(0)
            txtCode(5) = GProcGetColumnValue("TblMastAccount", "ACName", txttemp(Index), "S", "AcOurCode", "S")
            lblTinNo.Caption = GProcGetColumnValue("TblMastAccount", "ACName", txttemp(Index), "S", "AcGSTIN", "S")
            lblPan.Caption = GProcGetColumnValue("TblMastAccount", "ACName", txttemp(Index), "S", "AcPAN", "N")
            lblMblNo.Caption = GProcGetColumnValue("TblMastAccount", "ACName", txttemp(Index), "S", "AcContNo", "N")
'            If gClsSearch.SearchMultiRetCol(7) = 0 Then
'               lblMsOms.Caption = "M.S."
'            Else
'               lblMsOms.Caption = "O.M.S."
'            End If
        End If
        AgCode = GProcGetColumnValue("tblMastAccount", "AcName", txttemp(5), "S", "AgCode", "N")
        '-- Broker from master
        If txttemp(7) = "" Then
            BrkAcCode = GProcGetColumnValue("TblMastAccount", "ACName", txttemp(Index), "S", "AcBrkCode", "N")
            txttemp(7) = GProcGetColumnValue("TblMastAccount", "ACCode", CStr(BrkAcCode), "N", "ACname", "S")
            txtCode(7) = GProcGetColumnValue("TblMastAccount", "ACCode", CStr(BrkAcCode), "N", "AcOurCode", "S")
            If Val(txtCode(7)) = 0 Then
               txtCode(7) = ""
            End If
        End If
    Case 6  '-- Mill  Alias Account
         MilTpCd = GProcGetColumnValue("tblMastNarration", "Narration", "Trade", "S", "Narrcode", "N")
         ShowMill = "AgCode=90029 and   tblMastaccount.Accode = tblMastBillSerial.MillCode and  tblMastBillSerial.SaleType = " & MilTpCd & "  and tblMastBillSerial.Vfirm = '" & gCCode & "'  "
         gClsSearch.SearchMultiField "tblMastaccount,tblMastBillSerial", "BillSr,AcName,AcOurCode", Array("Mill Ini.", "Mill Name", "Code"), Array(1000, 8000, 1000), ShowMill, Trim(Chr(KeyAscii)), "AcName", LeftPos, TopPos, True
'        gClsSearch.SearchMultiField "tblMastaccount", "AcName,AcCity,AcOurCode", Array("Mill Name", "City", ""), Array(7000, 2000, 0), "(AgCode=90029) ", Trim(Chr(KeyAscii)), "AcName", LeftPos, TopPos, True
        KeyAscii = 0
        txtCode(6) = ""
        If gClsSearch.SearchMultiRetCol(1) <> "" Then
            txttemp(Index) = gClsSearch.SearchMultiRetCol(1)
            txtCode(6) = gClsSearch.SearchMultiRetCol(2)
        End If
        If mLastEntryMill <> txttemp(6) And FormAction = vbDataActionAddNew Then
           txttemp(17) = ""
           txtStdPack = 0
           txt(20) = 0
        End If
'        MillCtrlEd
    Case 7 '-- Broker
        gClsSearch.SearchMultiField "tblMastaccount", "AcOurCode,AcName,AcCity", Array("Code", "Broker", "City"), Array(1000, 5000, 3000), "(AgCode=90016 )", Trim(Chr(KeyAscii)), "", LeftPos, TopPos, True
        KeyAscii = 0
        txtCode(7) = ""
        If gClsSearch.SearchMultiRetCol(1) <> "" Then
            txttemp(Index) = gClsSearch.SearchMultiRetCol(1)
            txtCode(7) = gClsSearch.SearchMultiRetCol(0)
        End If
    Case 17  '-- Item
         MillCode = GProcGetColumnValue("TblMastAccount", "AcName", txttemp(6), "S", "Accode", "N")
         gClsSearch.SearchMultiField "tblMastItem,tblMastaccount", "ItName,AcName,ItStdpack,ItStdrateper", Array("Count Name", "Mill", "Std. Pack", "Rate Per"), Array(5000, 5000, 1000, 1000), "tblMastItem.ITMillCode=tblMastaccount.acCode and ITMillCode=" & MillCode, Chr(KeyAscii), "ItName", LeftPos, TopPos, True
         KeyAscii = 0
         If gClsSearch.SearchMultiRetCol(0) <> "" Then
             txttemp(Index) = gClsSearch.SearchMultiRetCol(0)
             txtStdPack = gClsSearch.SearchMultiRetCol(2)
'             If Val(txt(21)) = 0 Then
                txt(21) = gClsSearch.SearchMultiRetCol(3)
                txt(20) = 0
 
 '            End If
         End If
    Case 24 '-- Party Group
        gClsSearch.SearchMultiField "tblMastCompGroup", "CGName", Array("Company Group"), Array(txttemp(Index).Width), " ", Chr(KeyAscii), "CGName", LeftPos, TopPos, True
        KeyAscii = 0
        If gClsSearch.SearchMultiRetCol(0) <> "" Then
            txttemp(Index) = gClsSearch.SearchMultiRetCol(0)
        End If
    Case 23 '-- Delevery Add
        wPartyCode = GProcGetColumnValue("tblMastAccount", "AcName", txttemp(5), "S", "AcCode", "N")
'        gClsSearch.SearchMultiField "tblMastDeleAdd", "DCity,ContactPer,DAdd1,DAdd2,DAdd3,DZip,DTal,DDist,DState,DPhone", Array("City", "Contanc To", "Add1", "Add2", "Add3", "", "", "", "", ""), Array(txttemp(Index).Width, 2000, 2000, 2000, 2000, 0, 0, 0, 0, 0), "PartyCode = " & wPartyCode & "  ", Chr(KeyAscii), "DCity", LeftPos, TopPos, True
        gClsSearch.SearchMultiField "tblMastDeleAdd", "DCity,ContactPer,DAdd1,DAdd2,DAdd3,DZip,DTal,DDist,DState,DPhone", Array("City", "Contanc To", "Add1", "Add2", "Add3", "", "", "", "", ""), Array(2500, 3000, 4000, 4000, 3000, 0, 0, 0, 0, 0), "PartyCode = " & wPartyCode & "  ", Chr(KeyAscii), "DCity", 0, 0, True
        KeyAscii = 0
        lblContPer.Caption = ""
        lblAdd1.Caption = ""
        lblAdd2.Caption = ""
        lblAdd3.Caption = ""
        lblAddCity.Caption = ""
        lblAddPin.Caption = ""
        lblTal.Caption = ""
        lblDist.Caption = ""
        lblState.Caption = ""
        lblContMbl.Caption = ""
        If gClsSearch.SearchMultiRetCol(0) <> "" Then
            txttemp(Index) = gClsSearch.SearchMultiRetCol(0)
            lblContPer.Caption = gClsSearch.SearchMultiRetCol(1)
            lblAdd1.Caption = gClsSearch.SearchMultiRetCol(2)
            lblAdd2.Caption = gClsSearch.SearchMultiRetCol(3)
            lblAdd3.Caption = gClsSearch.SearchMultiRetCol(4)
            lblAddCity.Caption = gClsSearch.SearchMultiRetCol(0)
            lblAddPin.Caption = gClsSearch.SearchMultiRetCol(5)
            lblTal.Caption = gClsSearch.SearchMultiRetCol(6)
            lblDist.Caption = gClsSearch.SearchMultiRetCol(7)
            lblState.Caption = gClsSearch.SearchMultiRetCol(8)
            lblContMbl.Caption = gClsSearch.SearchMultiRetCol(9)
        End If
End Select
End Sub
Private Sub txtTemp_LostFocus(Index As Integer)
SaveBtnEd
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
    '--- Check is Related Gate Pass
'    If IsRelatedItCtrlGatePass(msGrid.TextMatrix(msGrid.Row, 14)) = False Then
        With msGrid
            Select Case .Col
                Case 0: '---- Sr No
                    txtGrid.Locked = True
                    ReNumberGridSrno
                    txtGrid = ""
                    .Col = .Col + 1
                Case Else
                    txtGrid.Locked = False
                    If .Col = 1 And txtGrid = "" Then
                       txtGrid = dtpDt(2)
                    End If
                    If .Col = 2 Then
                       OldBag = .TextMatrix(.Row, 2)
                    End If
            End Select
        End With
'    Else '-Related Gate Pass
 '         txtGrid.Locked = True
  '  End If
End If
msGrid_KeyPress (0)
End Sub
Private Sub msGrid_KeyPress(KeyAscii As Integer)
'--- Check is Related Gate Pass
'If IsRelatedItCtrlGatePass(msGrid.TextMatrix(msGrid.Row, 14)) = True Then Exit Sub
With msGrid
    If KeyAscii = 27 Then '----Esc Key
        KeyAscii = 0
        If ((Val(.TextMatrix(.Row, 2)) = 0)) And .rows > 2 Then
            .RemoveItem (.Row)
            txtGrid.text = ""
        End If
        txtGrid.Visible = False
        txttemp(23).SetFocus
        Exit Sub
    End If
End With
'If FormAction <> vbDataActionAddNew And FormAction <> vbDataActionUpdate Then Exit Sub
'---- Navigate Row ,Col
If KeyAscii = 13 Then
    KeyAscii = 0
    With msGrid
        If .Col = 3 Then 'rstblSub.Fields.Count - 2 Then
            If .Row = (.rows - 1) Then '-- Add New Row
                .rows = .rows + 1
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
If msGrid.Col < rstblSub.Fields.Count Then
    txtGrid.Move msGrid.CellLeft + msGrid.Left, msGrid.CellTop + msGrid.Top, msGrid.CellWidth, msGrid.CellHeight
End If
txtGrid.Visible = True
txtGrid.Enabled = True
DoEvents
If KeyAscii <> 0 Then
    Select Case rstblSub.Fields(msGrid.Col).Type
        Case adChar, adVarWChar
'            If msGrid.Col = 1 Or msGrid.Col = 3 Or msGrid.Col = 4 Then '-- Count ,godown
 '               txtGrid_KeyPress (KeyAscii)
  '          Else
                txtGrid.text = txtGrid.text & Chr(KeyAscii)
   '         End If
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
CalGridAmt
If TotBag > Val(txt(18)) Then
'   MsgBox "Check Sh.Bags Total.", vbCritical + vbOKOnly, Me.Caption
   msGrid.Col = 2
End If
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
End Sub
Private Sub CalGridAmt()
Dim MillCode As Long
Dim StdPack As Double
'Dim TotBag As Long
'Dim TotWt As Double
Dim RowIndex As Integer
CheckForNumberGrid
DelNote = False
RowIndex = 1
TotBag = 0
TotWt = 0

With msGrid
    Do While RowIndex < .rows
        '--Tot Bag
        If Val(.TextMatrix(RowIndex, 2)) > 0 Then
            TotBag = TotBag + Val(.TextMatrix(RowIndex, 2))
        End If
        '--Tot Wt
        If Val(.TextMatrix(RowIndex, 3)) > 0 Then
            TotWt = TotWt + Val(.TextMatrix(RowIndex, 3))
        End If
        If Val(.TextMatrix(RowIndex, 2)) = Val(txt(18)) And IsDate(.TextMatrix(RowIndex, 1)) = True Then
           If CDate(.TextMatrix(RowIndex, 1)) = dtpDt(2) Then
              DelNote = True
           End If
        End If
        
        RowIndex = RowIndex + 1
    Loop
End With
lblBags.Caption = TotBag
lblWeight.Caption = TotWt
MillCode = GProcGetColumnValue("TblMastAccount", "AcName", txttemp(6), "S", "Accode", "N")
StdPack = GProcGetColumnValue("TblMastItem", "ITName", txttemp(17), "S", "ItStdpack", "N", " ItMillCode=" & MillCode)
If msGrid.Row = 0 Then Exit Sub
'-- Wt=Bag * Std Pack
If OldBag <> msGrid.TextMatrix(msGrid.Row, 2) And StdPack > 0 Then
    msGrid.TextMatrix(msGrid.Row, 3) = msGrid.TextMatrix(msGrid.Row, 2) * StdPack
    OldBag = msGrid.TextMatrix(msGrid.Row, 2)
End If
If msGrid.TextMatrix(msGrid.Row, 2) = 0 Then
    msGrid.TextMatrix(msGrid.Row, 3) = 0
End If

'If IsDate(msGrid.TextMatrix(msGrid.Row, 1)) = False Then
'    msGrid.TextMatrix(msGrid.Row, 1) = dtpDt(2).text
'End If
'If Val(txt(18)) - TotBag > 0 Then
'   msGrid.TextMatrix(msGrid.Row, 2) = Val(txt(18)) - TotBag
'End If
'If Val(txt(19)) - TotWt > 0 Then
'   msGrid.TextMatrix(msGrid.Row, 3) = Val(txt(19)) - TotWt
'End If

End Sub
Private Sub CheckForNumberGrid()
Dim RowIndex As Long
With rstblSub
For RowIndex = 1 To msGrid.rows - 1
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
                If IsDate(msGrid.TextMatrix(RowIndex, i)) = False Then
'                    msGrid.TextMatrix(RowIndex, i) = ""
                    msGrid.TextMatrix(RowIndex, i) = dtpDt(2)
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
'--- Check is Related Gate Pass
'If IsRelatedItCtrlGatePass(msGrid.TextMatrix(msGrid.Row, 14)) = True Then Exit Sub
'--- Delete Row On Ctrl+Y
If KeyCode = vbKeyY And Shift = 2 And msGrid.rows > 2 Then
    msGrid = 1
    msGrid.RemoveItem (msGrid.Row)
    txtGrid.Visible = False
    txtGrid.text = ""
    ReNumberGridSrno
    CalAmount
    Exit Sub
End If
'--- Show Form
'If KeyCode = vbKeyF3 Then
'Select Case msGrid.Col
'    Case 3 '-- Item
'        GProcShowForm frmMastItem, frmMain.mnuMstfrm(3), Array(frmMastItem.txttemp(9)), Array(txttemp(7))
'        Exit Sub
'    Case 4 '-- Godown
'        OldgFrmTypeStr = gfrmTypeStr
'        gfrmTypeStr = "G"
'        GProcShowForm frmMastNarration, frmMain.mnuMstfrm(8), "G"
'        gfrmTypeStr = OldgFrmTypeStr
'End Select
'End If
End Sub
'-- ReNumber Grid Sr No
Private Sub ReNumberGridSrno()
Dim i As Long
With msGrid
    For i = 1 To .rows - 1
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
'--- Check is Related Gate Pass
'If IsRelatedItCtrlGatePass(msGrid.TextMatrix(msGrid.Row, 14)) = True Then Exit Sub
With msGrid
    If KeyAscii <> 13 And KeyAscii <> 27 And KeyAscii <> 25 Then
        '--- Search
        Select Case .Col
            Case 1 '--Booking
'                PrepareBooking Chr(KeyAscii)
'                KeyAscii = 0
                KeyAscii = GProcValidateKey(rstblSub, msGrid.Col, KeyAscii, txtGrid)
            Case Else
                KeyAscii = GProcValidateKey(rstblSub, msGrid.Col, KeyAscii, txtGrid)
        End Select
   End If
    If KeyAscii = 13 Then
        CheckForNumberTxtGrid
        If txtGrid = "" And .Col = 1 Then
           KeyAscii = 27
        End If
    End If
    '--- Navigate Row ,Col
    If KeyAscii = 13 And .Col < 3 Then 'rstblSub.Fields.Count - 2 Then
        .Col = .Col + 1
    ElseIf KeyAscii = 13 And .Col = 3 Then '= rstblSub.Fields.Count - 2 Then
        If .Row = (.rows - 1) Then
            .rows = .rows + 1
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
        If (Val(.TextMatrix(.Row, 2)) = 0) And .rows > 2 Then
            .RemoveItem (.Row)
            txtGrid.text = ""
        End If
        If TotBag > Val(txt(18)) Then
           MsgBox "Check Sh.Bags Total.", vbCritical + vbOKOnly, Me.Caption
           msGrid.Col = 2
           txtGrid.SetFocus
           Exit Sub
        End If
        txtGrid.Visible = False
        txttemp(23).SetFocus
        Exit Sub
    End If
End With
End Sub

'-- Grid Required
Private Function ValidateGrid() As Boolean
Dim TotBag As Long
Dim TotWt As Double
Dim i As Long
Dim RowIndex As Long
Dim X As String
Dim Rs1 As Recordset
Dim oldFrmTypeStr As String
If Trim(txtGrid.text) <> "" Then
    msGrid.text = txtGrid.text
    CalGridAmt
End If
txtGrid.Visible = False
txtGrid.text = ""
CheckForNumberGrid
With msGrid
Validate:
    For RowIndex = 1 To .rows - 1
        If Val(.TextMatrix(RowIndex, 2)) > 0 Then
            TotBag = TotBag + Val(.TextMatrix(RowIndex, 2))
        End If
        '--Tot Wt
        If Val(.TextMatrix(RowIndex, 3)) > 0 Then
            TotWt = TotWt + Val(.TextMatrix(RowIndex, 3))
        End If
        For i = 0 To rstblSub.Fields.Count - 1
            Select Case i
                Case 1 '-- Date
                    If CDbl(.TextMatrix(RowIndex, 2)) > 0 And IsDate(.TextMatrix(RowIndex, i)) = False Then
                        MsgBox "Check Date.", vbCritical + vbOKOnly, Me.Caption
                        .Col = i
                        msGrid.SetFocus
                        Exit Function
                    End If
                    If CDbl(.TextMatrix(RowIndex, 3)) > 0 And IsDate(.TextMatrix(RowIndex, i)) = False Then
                        MsgBox "Check Date.", vbCritical + vbOKOnly, Me.Caption
                        .Col = i
                        msGrid.SetFocus
                        Exit Function
                    End If
                    If IsDate(.TextMatrix(RowIndex, i)) = True Then
                       If CDate(.TextMatrix(RowIndex, i)) < dtpDt(2) Then
                           MsgBox "Delivery Date Must Be Less Than Booking Date.", vbCritical + vbOKOnly, Me.Caption
                           .Col = i
                           msGrid.SetFocus
                        Exit Function
                       End If
                    End If
            
'                Case 2 '-- Bag
'                    If CDbl(.TextMatrix(RowIndex, i)) = 0 Then
'                        MsgBox "Check Bag.", vbCritical + vbOKOnly, Me.Caption
'                        .Col = i
'                        msGrid.SetFocus
'                        Exit Function
'                    End If
'                Case 3 '-- Wt
'                    If CDbl(.TextMatrix(RowIndex, i)) = 0 Then
'                        MsgBox "Check Weight.", vbCritical + vbOKOnly, Me.Caption
'                        .Col = i
'                        msGrid.SetFocus
'                        Exit Function
'                    End If
           End Select
        Next i
      Next RowIndex
End With
If TotBag > Val(txt(18)) Then
   MsgBox "Check Despatch Detail Bag Total.", vbCritical + vbOKOnly, Me.Caption
   Exit Function
End If
If TotWt > Val(txt(19)) Then
   MsgBox "Check Despatch Detail Weight Total.", vbCritical + vbOKOnly, Me.Caption
   Exit Function
End If
ValidateGrid = True
ReNumberGridSrno
End Function
Private Sub FillCombo()
'Call GProcFillCombo(Cbo(22), "select Narration from tblMastNarration where NarrType ='T' order by Narration")
'Cbo(22) = GProcGetColumnValue("tblMastNarration", "NarrCode", CStr(gTaxCodePY), "N", "Narration", "S")
End Sub
Private Sub SaveBtnEd()
'--- Vno,vdt,Party,A/c,Mill,Amount
If (CDbl(txt(0)) = 0 And FormAction <> vbDataActionAddNew) Or GProcIsDateValid(dtpDt(2)) = False Or txttemp(6) = "" Or CDbl(txt(20)) = 0 Then
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
'    MillCtrlEd
    '--- Validate Grid
    If ValidateGrid = False Then
        'MSGrid.SetFocus
        Exit Function
    End If
    '--- Vno
    If CLng(txt(0)) = 0 And FormAction <> vbDataActionAddNew Then
        MsgBox "Check Booking No.", vbInformation + vbOKOnly, Me.Caption
        txt(0).SetFocus
        Exit Function
    End If
    '--- Vdt
    If GProcIsDateValid(dtpDt(2)) = False Then
        MsgBox "Check Booking Date.", vbInformation + vbOKOnly, Me.Caption
        dtpDt(2).SetFocus
        Exit Function
    End If
    '---Rate Type
    If Cbo(22) = "" Then
        MsgBox "Check Rate Type.", vbInformation + vbOKOnly, Me.Caption
        Cbo(22).SetFocus
        Exit Function
    End If
    '-- Bags
    If CDbl(txt(18)) <= 0 Then
        MsgBox "Check Bags", vbInformation + vbOKOnly, Me.Caption
        txt(18).SetFocus
        Exit Function
    End If
    '-- Weight
    If CDbl(txt(19)) <= 0 Then
        MsgBox "Check Weight", vbInformation + vbOKOnly, Me.Caption
        txt(19).SetFocus
        Exit Function
    End If
    '-- Rate
    If CDbl(txt(20)) <= 0 Then
        MsgBox "Check Rate", vbInformation + vbOKOnly, Me.Caption
        txt(20).SetFocus
        Exit Function
    End If
    '-- Per
    If CDbl(txt(21)) <= 0 Then
        MsgBox "Check Rate Per", vbInformation + vbOKOnly, Me.Caption
'        txt(21).SetFocus
        Exit Function
    End If
    '---- Mill Account
    If txttemp(6) = "" Then
        MsgBox "Check Mill Name", vbInformation + vbOKOnly, Me.Caption
        txttemp(6).SetFocus
        Exit Function
    End If
    '--- Broker
    If txttemp(7) = "" Then
        MsgBox "Check Broker Name", vbInformation + vbOKOnly, Me.Caption
        txttemp(7).SetFocus
        Exit Function
    End If
    '-- Mill & Count Check
    ExpAcCode = GProcGetColumnValue("TblMastAccount", "AcName", txttemp(6), "S", "Accode", "N")
    If GProcGetColumnValue("TblMastItem", "ItName", txttemp(17), "S", "Itcode", "N", "ITMillCode=" & ExpAcCode) = 0 Then
       MsgBox "Check Count.Mill and Count arr Not Mismatch.", vbCritical + vbOKOnly, Me.Caption
       txttemp(17).SetFocus
       Exit Function
    End If
    '----Duplicate Records
    '----Addition Case
    If FormAction = vbDataActionAddNew Then
        X = "Select vno From " & tblName & " Where vno=" & CLng(txt(0)) & " and Vtype='" & frmTypeStr & "'  and Vfirm = '" & gCCode & "'  and VYear=" & gCYear
        i = GProcRstOpen(rstAddCheck, X, "R")
        If i > 0 Then
            MsgBox "Record is Booking No. " & CLng(txt(0)), vbCritical + vbOKOnly, "Duplicate Record"
            txt(0).SetFocus
            Exit Function
        End If
        rstAddCheck.Close
    End If
    ValidateData = True
End Function
Private Sub FillTxtFromTemp()
txt(1) = frmTypeStr '--VType
If IsDate(dtpDt(2)) Then '--Vdt
    txt(2) = CDate(dtpDt(2))
Else
    txt(2) = ""
End If

'-- Party Code
txt(5) = GProcGetColumnValue("tblMastAccount", "AcName", txttemp(5), "S", "AcCode", "N")
'-- Mill Code
txt(6) = GProcGetColumnValue("tblMastAccount", "AcName", txttemp(6), "S", "AcCode", "N")
'--broker Code
txt(7) = GProcGetColumnValue("tblMastAccount", "AcName", txttemp(7), "S", "AcCode", "N")

'-- Party Group
txt(24) = GProcGetColumnValue("tblMastCompGroup", "CGName", txttemp(24), "S", "CGCode", "N")
'--Del Add
If txttemp(23) = "" Then
   txt(23) = 1
Else
   txt(23) = GProcGetColumnValue("tblMastDeleAdd", "isnull(ContactPer,'')+isnull(Dadd1,'')+isnull(Dadd2,'')+Isnull(Dadd3,'')+isnull(Dcity,'')", lblContPer + lblAdd1 + lblAdd2 + lblAdd3 + txttemp(23), "S", "DelCode", "N", "PartyCode = " & txt(5) & "")
End If
If TotBag > 0 And FormAction = vbDataActionAddNew Then
   If txt(26) = "" Then
      txt(26) = "As Per Schedule"
'   Else
'      txt(26) = Trim(txt(26)) + " As Per Schedule"
   End If
End If
If TotBag = 0 And FormAction = vbDataActionAddNew Then
   If txt(26) = "" Then
      txt(26) = "Earliest"
'   Else
'      txt(26) = Trim(txt(26)) + " Earliest"
   End If
End If
'--Count
txt(17) = GProcGetColumnValue("tblMastItem", "ItName", txttemp(17), "S", "Itcode", "N", "ItMillCode = " & txt(6) & "")
txt(3) = gCYear
txt(4) = gCCode
txt(22) = Cbo(22).ListIndex
End Sub
Private Sub FillTempFromTxt()
Dim X As String
Dim MsOms As Integer
FillCombo
If IsDate(txt(2)) Then '--VDt
    dtpDt(2) = txt(2)
End If
If IsDate(txt(29)) Then '--VDt
    dtpDt(29) = txt(29)
End If

'-- Party Code
txttemp(5) = GProcGetColumnValue("tblMastAccount", "AcCode", txt(5), "N", "AcName", "S")
txtCode(5) = GProcGetColumnValue("tblMastAccount", "AcCode", txt(5), "N", "AcOurCode", "S")

lblTinNo.Caption = GProcGetColumnValue("tblMastAccount", "AcCode", txt(5), "N", "AcGSTIN", "S")
lblPan.Caption = GProcGetColumnValue("tblMastAccount", "AcCode", txt(5), "N", "AcPAN", "S")
lblMblNo.Caption = GProcGetColumnValue("tblMastAccount", "AcCode", txt(5), "N", "AcContNo", "S")

'MsOms = GProcGetColumnValue("tblMastAccount", "AcCode", txt(5), "N", "AcMsOmsParty", "N")
'If MsOms = 0 Then
'   lblMsOms.Caption = "M.S."
'Else
'   lblMsOms.Caption = "O.M.S."
'End If

'-- Mill Code
txttemp(6) = GProcGetColumnValue("tblMastAccount", "AcCode", txt(6), "N", "AcName", "S")
txtCode(6) = GProcGetColumnValue("tblMastAccount", "AcCode", txt(6), "N", "AcOurCode", "S")

'--Broker
txttemp(7) = GProcGetColumnValue("tblMastAccount", "AcCode", txt(7), "N", "AcName", "S")
txtCode(7) = GProcGetColumnValue("tblMastAccount", "AcCode", txt(7), "N", "AcOurCode", "S")

'--Party Group
txttemp(24) = GProcGetColumnValue("tblMastCompGroup", "CGCode", txt(24), "N", "CGName", "S")

'--Item
txttemp(17) = GProcGetColumnValue("tblMastItem", "Itcode", txt(17), "N", "ItName", "S", "ItMillCode = " & txt(6) & "")
txtStdPack = GProcGetColumnValue("tblMastItem", "Itcode", txt(17), "N", "ItStdpack", "S", "ItMillCode = " & txt(6) & "")

'--Del Add
If txt(23) = "" Or txt(23) = 1 Then
    txttemp(23) = ""
    lblContPer.Caption = GProcGetColumnValue("tblMastAccount", "AcCode", txt(5), "N", "AcContPer", "S")
    lblAdd1.Caption = GProcGetColumnValue("tblMastAccount", "AcCode", txt(5), "N", "AcAdd1", "S")
    lblAdd2.Caption = GProcGetColumnValue("tblMastAccount", "AcCode", txt(5), "N", "AcAdd2", "S")
    lblAdd3.Caption = GProcGetColumnValue("tblMastAccount", "AcCode", txt(5), "N", "AcAdd3", "S")
    lblAddCity.Caption = ""
    lblAddPin.Caption = ""
    lblTal.Caption = ""
    lblDist.Caption = ""
    lblState.Caption = ""
    lblContMbl.Caption = GProcGetColumnValue("tblMastAccount", "AcCode", txt(5), "N", "AcContNo", "S")
Else
    txttemp(23) = GProcGetColumnValue("tblMastDeleAdd", "DelCode", txt(23), "N", "Dcity", "S", "PartyCode = " & txt(5) & "")
    lblContPer.Caption = GProcGetColumnValue("tblMastDeleAdd", "DelCode", txt(23), "N", "ContactPer", "S", "PartyCode = " & txt(5) & "")
    lblAdd1.Caption = GProcGetColumnValue("tblMastDeleAdd", "DelCode", txt(23), "N", "DAdd1", "S", "PartyCode = " & txt(5) & "")
    lblAdd2.Caption = GProcGetColumnValue("tblMastDeleAdd", "DelCode", txt(23), "N", "DAdd2", "S", "PartyCode = " & txt(5) & "")
    lblAdd3.Caption = GProcGetColumnValue("tblMastDeleAdd", "DelCode", txt(23), "N", "DAdd3", "S", "PartyCode = " & txt(5) & "")
    lblAddCity.Caption = GProcGetColumnValue("tblMastDeleAdd", "DelCode", txt(23), "N", "DCity", "S", "PartyCode = " & txt(5) & "")
    lblAddPin.Caption = GProcGetColumnValue("tblMastDeleAdd", "DelCode", txt(23), "N", "DZip", "S", "PartyCode = " & txt(5) & "")
    lblTal.Caption = GProcGetColumnValue("tblMastDeleAdd", "DelCode", txt(23), "N", "DTal", "S", "PartyCode = " & txt(5) & "")
    lblDist.Caption = GProcGetColumnValue("tblMastDeleAdd", "DelCode", txt(23), "N", "DDist", "S", "PartyCode = " & txt(5) & "")
    lblState.Caption = GProcGetColumnValue("tblMastDeleAdd", "DelCode", txt(23), "N", "DState", "S", "PartyCode = " & txt(5) & "")
    lblContMbl.Caption = GProcGetColumnValue("tblMastDeleAdd", "DelCode", txt(23), "N", "DPhone", "S", "PartyCode = " & txt(5) & "")
End If
'-- TaxType
Cbo(22).ListIndex = Val(txt(22))
'-- Fill Related Record
FillRelatedRecord
CalGridAmt
End Sub
Private Sub FillRelatedRecord()
Dim RowIndex As Long
Dim ColIndex As Long
'Dim TotBag As Long
'Dim TotWt As Double
'tmpAutoLRNo = "" ''-Auto LR No
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
'                If ColIndex = 9 Then TotBag = TotBag + CLng(msGrid.TextMatrix(RowIndex, 9)) '-- Tot Bag
'                If ColIndex = 10 Then TotWt = TotWt + CDbl(msGrid.TextMatrix(RowIndex, 10)) '-- Tot Wt
            Next ColIndex
            RowIndex = RowIndex + 1
            .MoveNext
            If Not .EOF Then msGrid.rows = msGrid.rows + 1
        Loop
    Else
        msGrid.rows = msGrid.rows + 1
        ReNumberGridSrno
    End If
End With
'--- Fill Related Control
SetRelatedRS
End Sub
Private Sub ReadFields()
FormAction = vbDataActionClose
GProcShowRecord Me, rstbl, MaxNo, OtherTxtFromNo, OtherTxtToNo
FillTempFromTxt
End Sub
Private Sub CalAmount()
Dim Amt As Double
'Dim RowIndex As Integer
'Dim TotBag As Long
'Dim TotWt As Double
Dim MillCode As Long
Dim StdPack As Double

'---- Sub amt
CheckForNumberGrid
CheckForNumber
'RowIndex = 1
'With msGrid
'    Do While RowIndex < .rows
'        '--Tot Bag
'        If Val(.TextMatrix(RowIndex, 2)) > 0 Then
'            TotBag = TotBag + Val(.TextMatrix(RowIndex, 2))
'        End If
'        '--Tot Wt
'        If Val(.TextMatrix(RowIndex, 3)) > 0 Then
'            TotWt = TotWt + Val(.TextMatrix(RowIndex, 3))
'        End If
'        RowIndex = RowIndex + 1
'    Loop
'End With
'lblBags.Caption = TotBag
'lblWeight.Caption = TotWt
MillCode = GProcGetColumnValue("TblMastAccount", "AcName", txttemp(6), "S", "Accode", "N")
StdPack = GProcGetColumnValue("TblMastItem", "ITName", txttemp(17), "S", "ItStdpack", "N", " ItMillCode=" & MillCode)
'-- Wt=Bag * Std Pack
If OldBag2 <> Val(txt(18)) And StdPack > 0 Then
    txt(19) = Val(txt(18)) * StdPack
    OldBag2 = Val(txt(18))
End If
'If Val(txt(19)) = 0 Then
'    txt(19) = Format((Val(txt(18)) * StdPack), WtStr)
'End If

If Val(txt(28)) = 0 Then
    txt(28) = Format((Val(txt(27)) * StdPack), WtStr)
End If
If Val(txt(21)) > 0 Then
   Amt = (CDbl(txt(19)) * CDbl(txt(20))) / CDbl(txt(21))
Else
   Amt = 0
End If
txt(10) = Amt
CheckForNumber
End Sub
'--- Is Related Record
Private Function IsRelatedRecord() As Boolean
Dim RsG As Recordset
Dim X As String
'-- Audited
'If GProcISAudited(txt(0), txt(1), txt(23)) = True Then
'    MsgBox "Record is Audited.You can not Modify/Delete this Record.", vbCritical + vbOKOnly, App.Title
'    IsRelatedRecord = True
'    Exit Function
'End If
'-- Is Related Despatch Detail

X = "Select Vno,Vdt from tblBooKingDesp where tblBooKingDesp.vtype='" & frmTypeStr & "' and tblBooKingDesp.vno=" & txt(0) & " and tblBooKingDesp.VFirm = '" & gCCode & "' and " _
    & "  tblBooKingDesp.VYear=" & gCYear
GProcRstOpen RsG, X, "R"
If RsG.EOF Then
Else
    IsRelatedRecord = True
    MsgBox "You can not Delete or Modify this record.Related Despatch Details are exists.", vbCritical + vbOKOnly, App.Title
    Exit Function
End If
RsG.Close
End Function
'-- Is Related GatePass
Private Function IsRelatedGatePass() As Boolean
Dim RsG As Recordset
Dim X As String
'-- Is Gate Pass
'X = "Select Vno,Vdt from tblVouSub where Vtype='GP' and " _
'& " VSubPItCtrlNo in (select VSubItCtrlNo from tblVouSub where vtype='" & frmTypeStr & "' and vno=" & txt(0) & " and " _
'& " VYear=" & gCYear & ")"
'GProcRstOpen RsG, X, "R"
'If RsG.EOF Then
'Else
'    IsRelatedGatePass = True
'    MsgBox "You can not Modify/Delete this record.Related Gate Pass No. " & RsG.Fields(0) & " is exists.", vbCritical + vbOKOnly, App.Title
'    Exit Function
'End If
'RsG.Close
End Function
'-- Is Related It Ctrl GatePass
Private Function IsRelatedItCtrlGatePass(ItCtrlNo As Long) As Boolean
Dim RsG As Recordset
Dim X As String
'-- Is Gate Pass
'X = "Select Vno,Vdt from tblVouSub where Vtype='GP' and " _
'& " VSubPItCtrlNo = " & ItCtrlNo
'GProcRstOpen RsG, X, "R"
'If RsG.EOF Then
'Else
'    IsRelatedItCtrlGatePass = False
'    Exit Function
'End If
'RsG.Close
End Function

