VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form frmBookingParty 
   BackColor       =   &H00FBF2E1&
   BorderStyle     =   0  'None
   Caption         =   "Purchase Entry"
   ClientHeight    =   9300
   ClientLeft      =   0
   ClientTop       =   15
   ClientWidth     =   13875
   FillColor       =   &H000040C0&
   FillStyle       =   0  'Solid
   Icon            =   "frmBookingParty.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   9300
   ScaleWidth      =   13875
   ShowInTaskbar   =   0   'False
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
      Height          =   3705
      Left            =   3000
      TabIndex        =   96
      Top             =   2400
      Visible         =   0   'False
      Width           =   6165
      Begin VB.CheckBox chkOff 
         BackColor       =   &H00FFC0C0&
         Caption         =   "Office Copy"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   195
         Left            =   1440
         TabIndex        =   124
         Top             =   2040
         Width           =   1575
      End
      Begin VB.CheckBox ChkCus 
         BackColor       =   &H00FFC0C0&
         Caption         =   "Customer Copy"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   195
         Left            =   3360
         TabIndex        =   123
         Top             =   2040
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
         Left            =   4605
         TabIndex        =   102
         Top             =   2730
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
         TabIndex        =   101
         Top             =   2745
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
         TabIndex        =   100
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
         TabIndex        =   99
         Top             =   2730
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
         TabIndex        =   98
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
         TabIndex        =   97
         Top             =   2745
         Visible         =   0   'False
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
         TabIndex        =   104
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
         TabIndex        =   103
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
   Begin VB.Frame Frame2 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
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
      Height          =   2115
      Left            =   3060
      TabIndex        =   127
      Top             =   2640
      Visible         =   0   'False
      Width           =   5865
      Begin VB.CommandButton cmdClose 
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
         Left            =   2220
         TabIndex        =   128
         Top             =   1320
         Width           =   1335
      End
      Begin VB.Label Label7 
         BackColor       =   &H00C0FFFF&
         Caption         =   "This Is No Work A/c"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   24
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   735
         Left            =   360
         TabIndex        =   129
         Top             =   420
         Width           =   6495
      End
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
      ItemData        =   "frmBookingParty.frx":058A
      Left            =   2295
      List            =   "frmBookingParty.frx":0594
      TabIndex        =   15
      Text            =   "Cbo(22)"
      Top             =   6285
      Width           =   2520
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
      Index           =   43
      Left            =   15480
      TabIndex        =   122
      Text            =   "43"
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
      Index           =   42
      Left            =   15480
      TabIndex        =   121
      Text            =   "42"
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
      Index           =   41
      Left            =   15435
      TabIndex        =   120
      Text            =   "41"
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
      Index           =   40
      Left            =   15435
      TabIndex        =   119
      Text            =   "40"
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
      Index           =   39
      Left            =   15345
      TabIndex        =   118
      Text            =   "39"
      Top             =   3105
      Visible         =   0   'False
      Width           =   375
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
      Index           =   38
      Left            =   2160
      MaxLength       =   60
      TabIndex        =   18
      Text            =   "38"
      Top             =   8550
      Width           =   6420
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
      Left            =   2160
      MaxLength       =   60
      TabIndex        =   17
      Text            =   "37"
      Top             =   8100
      Width           =   6420
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
      TabIndex        =   116
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
      TabIndex        =   115
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
      TabIndex        =   114
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
      TabIndex        =   113
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
      TabIndex        =   112
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
      TabIndex        =   111
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
      TabIndex        =   110
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
      Left            =   8280
      TabIndex        =   11
      Text            =   "txtStdPack"
      Top             =   5490
      Width           =   1110
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
      Left            =   1395
      TabIndex        =   4
      Text            =   "txtCode(7)"
      ToolTipText     =   "Press F3 for Creation"
      Top             =   2925
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
      Left            =   1395
      TabIndex        =   2
      Text            =   "txtCode(5)"
      ToolTipText     =   "Press F3 for Creation"
      Top             =   1980
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
      Left            =   1395
      TabIndex        =   6
      Text            =   "txtCode(6)"
      ToolTipText     =   "Press F3 for Creation"
      Top             =   4140
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
      TabIndex        =   108
      ToolTipText     =   "To Find Record"
      Top             =   1665
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
      TabIndex        =   107
      ToolTipText     =   "Next"
      Top             =   705
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
      TabIndex        =   106
      ToolTipText     =   "Last"
      Top             =   1185
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
      TabIndex        =   105
      ToolTipText     =   "To Modify Record"
      Top             =   2580
      Width           =   1005
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
      TabIndex        =   93
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
      Left            =   10215
      TabIndex        =   27
      Text            =   "28"
      Top             =   3150
      Width           =   1650
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
      Left            =   10215
      TabIndex        =   26
      Text            =   "27"
      Top             =   2655
      Width           =   1650
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
      Left            =   2160
      MaxLength       =   50
      TabIndex        =   16
      Text            =   "26"
      Top             =   7515
      Width           =   6420
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
      Left            =   15480
      MaxLength       =   50
      TabIndex        =   22
      Text            =   "25"
      Top             =   4905
      Visible         =   0   'False
      Width           =   390
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
      Left            =   2610
      TabIndex        =   7
      Text            =   "txttemp(6)"
      ToolTipText     =   "Press F3 for Creation"
      Top             =   4140
      Width           =   6045
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
      Left            =   2520
      TabIndex        =   3
      Text            =   "txttemp(5)"
      ToolTipText     =   "Press F3 for Creation"
      Top             =   1980
      Width           =   6135
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
      Left            =   2520
      TabIndex        =   5
      Text            =   "txttemp(7)"
      ToolTipText     =   "Press F3 for Creation"
      Top             =   2925
      Width           =   6135
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
      TabIndex        =   21
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
      TabIndex        =   81
      Text            =   "24"
      Top             =   7155
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
      Index           =   23
      Left            =   18540
      TabIndex        =   75
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
      TabIndex        =   73
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
      TabIndex        =   72
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
      TabIndex        =   71
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
      TabIndex        =   70
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
      TabIndex        =   69
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
      Left            =   11070
      TabIndex        =   14
      Text            =   "21"
      Top             =   5490
      Width           =   825
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
      Left            =   9450
      TabIndex        =   13
      Text            =   "20"
      Top             =   5490
      Width           =   1500
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
      Left            =   540
      TabIndex        =   8
      Text            =   "txttemp(17)"
      ToolTipText     =   "Press F3 for Creation"
      Top             =   5520
      Width           =   5370
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
      Left            =   6930
      TabIndex        =   10
      Text            =   "19"
      Top             =   5490
      Width           =   1275
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
      TabIndex        =   62
      Top             =   4020
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
      Left            =   6075
      TabIndex        =   9
      Text            =   "18"
      Top             =   5490
      Width           =   735
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
      TabIndex        =   59
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
      TabIndex        =   23
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
      TabIndex        =   57
      Text            =   "14"
      Top             =   4905
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.CommandButton CmdBagDetails 
      Caption         =   "Cartoon &Bag Details"
      Height          =   720
      Left            =   17415
      TabIndex        =   55
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
      TabIndex        =   54
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
      Left            =   10890
      TabIndex        =   20
      Top             =   7875
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
      TabIndex        =   52
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
      TabIndex        =   50
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
      TabIndex        =   49
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
      TabIndex        =   48
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
      TabIndex        =   44
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
      TabIndex        =   43
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
      TabIndex        =   41
      ToolTipText     =   "Exit"
      Top             =   5895
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
      TabIndex        =   37
      ToolTipText     =   "To Save Record"
      Top             =   3555
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
      TabIndex        =   42
      Top             =   4500
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
      TabIndex        =   36
      ToolTipText     =   "To Delete Record"
      Top             =   3060
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
      TabIndex        =   35
      ToolTipText     =   "To Add New Record"
      Top             =   2100
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
      TabIndex        =   38
      Top             =   4995
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
      TabIndex        =   40
      ToolTipText     =   "First"
      Top             =   1185
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
      Left            =   2520
      TabIndex        =   0
      Text            =   "0"
      Top             =   810
      Width           =   1650
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
      TabIndex        =   39
      ToolTipText     =   "Previous"
      Top             =   705
      Width           =   465
   End
   Begin VB.TextBox txt 
      Alignment       =   1  'Right Justify
      Height          =   315
      Index           =   4
      Left            =   18390
      TabIndex        =   47
      Text            =   "4"
      Top             =   1530
      Visible         =   0   'False
      Width           =   525
   End
   Begin VB.TextBox txt 
      Height          =   315
      Index           =   6
      Left            =   18480
      TabIndex        =   46
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
      TabIndex        =   45
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
      TabIndex        =   34
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
      TabIndex        =   12
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
      Bindings        =   "frmBookingParty.frx":05B6
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
      Height          =   1950
      Left            =   8955
      TabIndex        =   19
      Top             =   7110
      Width           =   4695
      _ExtentX        =   8281
      _ExtentY        =   3440
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
      Height          =   390
      Index           =   2
      Left            =   7005
      TabIndex        =   1
      Top             =   810
      Width           =   1650
      _ExtentX        =   2910
      _ExtentY        =   688
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
      Format          =   113377283
      CurrentDate     =   -273
   End
   Begin MSComCtl2.DTPicker dtpDt 
      Height          =   390
      Index           =   29
      Left            =   10215
      TabIndex        =   24
      Top             =   2160
      Width           =   1650
      _ExtentX        =   2910
      _ExtentY        =   688
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
      Format          =   113377283
      CurrentDate     =   -273
   End
   Begin VB.Label Label10 
      BackStyle       =   0  'Transparent
      Caption         =   "In./Ex. GST         :"
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
      Height          =   255
      Left            =   540
      TabIndex        =   126
      Top             =   6300
      Width           =   1995
   End
   Begin VB.Label Label37 
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
      ForeColor       =   &H000000C0&
      Height          =   255
      Left            =   540
      TabIndex        =   125
      Top             =   6060
      Width           =   1995
   End
   Begin VB.Line Line2 
      BorderColor     =   &H000000C0&
      X1              =   315
      X2              =   12060
      Y1              =   5235
      Y2              =   5235
   End
   Begin VB.Line Line1 
      BorderColor     =   &H000000C0&
      X1              =   315
      X2              =   12060
      Y1              =   4695
      Y2              =   4695
   End
   Begin VB.Label Label35 
      BackStyle       =   0  'Transparent
      Caption         =   "Note                :"
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
      Left            =   630
      TabIndex        =   117
      Top             =   8145
      Width           =   1500
   End
   Begin VB.Label Label22 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Std.Pkg."
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
      Left            =   8565
      TabIndex        =   109
      Top             =   4815
      Width           =   825
   End
   Begin VB.Shape Shape1 
      FillColor       =   &H00FED8FC&
      FillStyle       =   0  'Solid
      Height          =   5880
      Left            =   12330
      Shape           =   4  'Rounded Rectangle
      Top             =   585
      Width           =   1350
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
      Left            =   6435
      TabIndex        =   95
      Top             =   2475
      Width           =   2145
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
      Left            =   14760
      TabIndex        =   94
      Top             =   90
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
      Left            =   9180
      TabIndex        =   92
      Top             =   2205
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
      Left            =   9135
      TabIndex        =   91
      Top             =   3150
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
      Left            =   8865
      TabIndex        =   90
      Top             =   2745
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
      TabIndex        =   89
      Top             =   1800
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
      TabIndex        =   88
      Top             =   7575
      Width           =   1635
   End
   Begin VB.Shape Shape3 
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00000080&
      FillColor       =   &H00EDFDFE&
      FillStyle       =   0  'Solid
      Height          =   2040
      Left            =   8955
      Top             =   1620
      Width           =   3120
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
      Left            =   5580
      TabIndex        =   87
      Top             =   2475
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
      Left            =   14040
      TabIndex        =   86
      Top             =   90
      Width           =   705
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
      TabIndex        =   85
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
      TabIndex        =   84
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
      TabIndex        =   83
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
      TabIndex        =   82
      Top             =   8145
      Visible         =   0   'False
      Width           =   555
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
      TabIndex        =   80
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
      TabIndex        =   79
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
      TabIndex        =   78
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
      TabIndex        =   77
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
      TabIndex        =   76
      Top             =   8145
      Visible         =   0   'False
      Width           =   1950
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
      Left            =   8940
      TabIndex        =   74
      Top             =   6825
      Width           =   1695
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
      Left            =   11430
      TabIndex        =   68
      Top             =   4815
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
      Left            =   10125
      TabIndex        =   67
      Top             =   4815
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
      Left            =   7380
      TabIndex        =   66
      Top             =   4815
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
      Left            =   5985
      TabIndex        =   65
      Top             =   4815
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
      Left            =   540
      TabIndex        =   64
      Top             =   4815
      Width           =   870
   End
   Begin VB.Shape Shape5 
      BorderColor     =   &H00EFB17A&
      BorderWidth     =   5
      FillColor       =   &H00EFB17A&
      Height          =   9315
      Left            =   0
      Top             =   -45
      Width           =   13860
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
      TabIndex        =   63
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
      Left            =   2520
      TabIndex        =   61
      Top             =   2475
      Width           =   2190
   End
   Begin VB.Label Label21 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "GSTN    :   "
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
      Left            =   1395
      TabIndex        =   60
      Top             =   2475
      Width           =   1020
   End
   Begin VB.Label Label20 
      BackStyle       =   0  'Transparent
      Caption         =   "Date    :    "
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
      Left            =   6075
      TabIndex        =   58
      Top             =   855
      Width           =   825
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Booking No     :"
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
      Left            =   585
      TabIndex        =   56
      Top             =   855
      Width           =   1455
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "&Party  :"
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
      Top             =   2025
      Width           =   1035
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "M&ill     :"
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
      Top             =   4140
      Width           =   960
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "&Brok.  : "
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
      Top             =   2925
      Width           =   960
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
      TabIndex        =   53
      Top             =   0
      Width           =   2475
   End
   Begin VB.Label Label6 
      BackColor       =   &H00EFB17A&
      Caption         =   "     Booking ( Party  )"
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
      TabIndex        =   51
      Top             =   0
      Width           =   13860
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H00000080&
      FillColor       =   &H00EDFDFE&
      FillStyle       =   0  'Solid
      Height          =   2970
      Left            =   315
      Top             =   3870
      Width           =   11760
   End
   Begin VB.Shape Shape4 
      BorderColor     =   &H00000080&
      FillColor       =   &H00EDFDFE&
      FillStyle       =   0  'Solid
      Height          =   2040
      Left            =   315
      Top             =   1620
      Width           =   8430
   End
   Begin VB.Shape Shape6 
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00000080&
      FillColor       =   &H00EDFDFE&
      FillStyle       =   0  'Solid
      Height          =   1950
      Left            =   315
      Top             =   7110
      Width           =   8430
   End
   Begin VB.Shape Shape7 
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00000080&
      FillColor       =   &H00EDFDFE&
      FillStyle       =   0  'Solid
      Height          =   825
      Left            =   315
      Top             =   585
      Width           =   8430
   End
End
Attribute VB_Name = "frmBookingParty"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Const tblName As String = "tblBooking"
Const IdField As String = "Vno"
Const OrderField As String = "Vdt,Vno"
Const MaxNo As Long = 43
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
'    If Year(mLastEntryDate) = 1899 Then
        dtpDt(2) = Format(Now, "dd/MM/yyyy")
 '   Else
  '      dtpDt(2) = mLastEntryDate
   ' End If
    
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
    chkOff.Enabled = True
    chkOff.Value = 1
    ChkCus.Enabled = True
    ChkCus.Value = 1
    DoEvents
    txtPInvno(0).Enabled = True
    txtPInvno(1).Enabled = True
    txtPInvno(0).SetFocus
    If gCIsDepotFirm = 1 Then
       cmdPrintBill(3).Visible = True
    End If
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
        FillTempFromTxt
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

'-- For WhatsApp SMS PDF
'''        Dim RsSA As Recordset
'''        Dim aa As Boolean
'''        Dim X As String
'''        aa = False
'''        If MsgBox("Do You Want To Send SMS ? ", vbYesNo) = vbYes Then aa = True
'''        If aa Then
'''           GProcCrystalRptPreparation cryReport, "Window"
'''
'''           X = "Select a.Vno,a.Vtype,a.Vdt,a.Vyear,a.VFirm,BkParty,BkMillCode,BkBroker,BkRefNo,BkRefDt,BkSubAmt,BkBillAmt,BkRoff,DueDays,DueDate,BkEntDt,'" & gUserName & "',BkItCode,BkBag,BkWt,BkRt,BkRtPer,BkRateType,BkDeleAdd,BkPartyGroup,BkExtraCount,BkNote,BkCbag,BkCWt," _
'''                & " BkCDate,BkTmpNo1,BkTmpNo2,BkTmpNo3,BkTmpChr1,BkTmpChr2,BkTmpChr3,BkShBag,BkShWt,BkShDate,a.BkIsDirectPayment,a.BkIsFrghtInsu,a.BkIsInsu " _
'''                & " From tblBooking A  where a.vno = " & txt(0) & " and a.Vtype='" & frmTypeStr & "' and a.VFirm = '" & gCCode & "' And a.VYear = " & gCYear
'''
'''           gCn.Execute "Delete from TmpBooking where  BkUser='" & gUserName & "'"
'''
'''           gCn.Execute " insert into TmpBooking (Vno,Vtype,Vdt,Vyear,VFirm,BkParty,BkMillCode,BkBroker,BkRefNo,BkRefDt,BkSubAmt,BkBillAmt,BkRoff,DueDays,DueDate,BkEntDt,BkUser,BkItCode,BkBag,BkWt,BkRt,BkRtPer,BkRateType,BkDeleAdd,BkPartyGroup,BkExtraCount,BkNote,BkCbag,BkCWt," _
'''                       & " BkCDate,BkTmpNo1,BkTmpNo2,BkTmpNo3,BkTmpChr1,BkTmpChr2,BkTmpChr3,BkSubBag,BkSubWt,BkSubDDate,BkIsDirectPayment,BkIsFrghtInsu,BkIsInsu ) " & X
'''
'''           X = "select * from TmpBooking where BkUser='" & gUserName & "'"
'''
'''           If GProcRstOpen(RsSA, X, "R") <= 0 Then
'''              MsgBox "No Records..."
'''              Exit Sub
'''           End If
'''           With cryReport
'''                Dim objCrystal As CRAXDRT.Application
'''                Dim objReport As CRAXDRT.Report
'''                Dim Tattach As String
'''                Dim Mfile As String
'''                Dim GstinDepot As String
'''                Dim PanDepot As String
'''                Dim PayTo As String
'''
'''                Set objCrystal = New CRAXDRT.Application
'''                Tattach = gReportPath & "rptBookingPrint_Party.rpt"
'''                Set objReport = objCrystal.OpenReport(Tattach, 1)
'''                objReport.Database.LogOnServer "pdsodbc.dll", "Hitrix", "Vishnu", "sa", "123456"
'''                objReport.RecordSelectionFormula = "{TmpBooking.BkUser}='" & gUserName & "' "
'''                Dim j As Integer
'''                For j = 1 To objReport.FormulaFields.Count
'''                    Select Case objReport.FormulaFields(j).Name
'''                    Case "{@FirmNm}"
'''                          objReport.FormulaFields(j).text = "Trim(" & Chr(39) & UCase(gCName) & Chr(39) & ")"
'''                    Case "{@wAdd1}"
'''                          objReport.FormulaFields(j).text = "Trim(" & Chr(39) & gCAdd1 & Chr(39) & ")"
'''                    Case "{@wAdd2}"
'''                          objReport.FormulaFields(j).text = "Trim(" & Chr(39) & gCAdd2 & Chr(39) & ")"
'''                    Case "{@wAdd3}"
'''                          objReport.FormulaFields(j).text = "Trim(" & Chr(39) & gCAdd3 & Chr(39) & ")"
'''                    Case "{@wPhNo}"
'''                          objReport.FormulaFields(j).text = "Trim(" & Chr(39) & gCPhNo & Chr(39) & ")"
'''                    Case "{@wcpan}"
'''                          If frmTypeStr = "SD" Then
'''                             PanDepot = GProcGetColumnValue("tblMastAccount", "AcCode", txt(7), "N", "AcPAN", "S")
'''                             objReport.FormulaFields(j).text = "Trim(" & Chr(39) & UCase(PanDepot) & Chr(39) & ")"
'''                          Else
'''                             objReport.FormulaFields(j).text = "Trim(" & Chr(39) & UCase(gCPAN) & Chr(39) & ")"
'''                          End If
'''                    Case "{@wCSTNo}"
'''                          objReport.FormulaFields(j).text = "Trim(" & Chr(39) & UCase(gCCST) & Chr(39) & ")"
'''                    Case "{@wBSTNo}"
'''                          objReport.FormulaFields(j).text = "Trim(" & Chr(39) & UCase(gCBST) & Chr(39) & ")"
'''                    Case "{@wGSTNo}"
'''                          If frmTypeStr = "SD" Then
'''                             GstinDepot = GProcGetColumnValue("tblMastAccount", "AcCode", txt(7), "N", "AcGSTIN", "S")
'''                             objReport.FormulaFields(j).text = "Trim(" & Chr(39) & UCase(GstinDepot) & Chr(39) & ")"
'''                          Else
'''                             objReport.FormulaFields(j).text = "Trim(" & Chr(39) & UCase(gCGStin) & Chr(39) & ")"
'''                          End If
'''                    Case "{@wPayTo}"
'''                          If frmTypeStr = "SD" Then
'''                             PayTo = GProcGetColumnValue("tblMastAccount", "AcCode", "'" & gCDepotMainFirm & "' ", "N", "AcName", "S")
'''                          Else
'''                             PayTo = gCName
'''                          End If
'''                          objReport.FormulaFields(j).text = "Trim(" & Chr(39) & (PayTo) & Chr(39) & ")"
'''                    Case "{@wBank}"
'''                          objReport.FormulaFields(j).text = "Trim(" & Chr(39) & gCBankName & Chr(39) & ")"
'''                    Case "{@wBankAcNo}"
'''                          objReport.FormulaFields(j).text = "Trim(" & Chr(39) & gCBankAcNo & Chr(39) & ")"
'''                    Case "{@wBankIfsc}"
'''                          objReport.FormulaFields(j).text = "Trim(" & Chr(39) & gCBankRTGSCode & Chr(39) & ")"
'''                    Case "{@wCopy}"
'''                          objReport.FormulaFields(j).text = "Trim(" & Chr(39) + "Customer Copy" + Chr(39) & ")"
'''                    End Select
'''                Next j
'''           End With
'''           Dim PtyMbl As String
'''           Dim BrkMbl As String
'''           Dim FileNMWithPath As String
'''           Dim FileNM As String
'''           Dim PtyName As String
'''           Dim BrkName As String
'''           Dim aFile As String
'''           Dim RsTmp As Recordset
'''
'''                   '--Error
'''           gCn.BeginTrans
'''           gCn.Execute "Delete from tmpWhatsAppErr where username = '" & gUserName & "' and Vfirm = '" & gCCode & "' "
'''           gCn.CommitTrans
'''                   '--
'''           FileNM = "" & Format(Now, "ddMMyyyyhhmmss") & "" & ".pdf"
'''           FileNMWithPath = gReportPath & "" & FileNM & ""
'''           PtyMbl = GProcGetColumnValue("TblMastAccount", "AcCode", RsSA!BkParty, "N", "AcContNo", "S")
'''           PtyName = GProcGetColumnValue("TblMastAccount", "AcCode", RsSA!BkParty, "N", "AcName", "S")
'''
'''           BrkMbl = GProcGetColumnValue("TblMastAccount", "AcCode", RsSA!BkBroker, "N", "AcContNo", "S")
'''           BrkName = GProcGetColumnValue("TblMastAccount", "AcCode", RsSA!BkBroker, "N", "AcName", "S")
'''
'''           If PtyMbl = "" Or IsNull(PtyMbl) Then
'''              MsgBox "Party Phone Not found..."
'''           Else
'''              ExportReportToPDF objReport, FileNMWithPath, "foo"
'''              SentWhatsApp "PDF", FileNM, "", PtyMbl, PtyName
'''           End If
'''
'''
'''           If BrkMbl = "" Or IsNull(BrkMbl) Then
'''              MsgBox "Broker Phone Not found..."
'''           Else
'''              ExportReportToPDF objReport, FileNMWithPath, "foo"
'''              SentWhatsApp "PDF", FileNM, "", BrkMbl, BrkName
'''           End If
'''
'''
'''
'''           X = "Select * from tmpWhatsAppErr where username = '" & gUserName & "' and vfirm = '" & gCCode & "'"
'''           If GProcRstOpen(RsTmp, X, "R", gCn) > 0 Then
'''              GProcWhatsAppError
'''           End If
'''           gReportPath = App.Path & "\HIReports\"
'''           aFile = gReportPath & "*.pdf"
'''           If Len(Dir$(aFile)) > 0 Then
'''              Kill aFile
'''           End If
'''        End If
'''     End If



        Dim aa As Boolean
        aa = False
        If MsgBox("Do You Want To Send SMS ? ", vbYesNo) = vbYes Then aa = True
        If aa Then
            Dim HttpReq As New WinHttpRequest
            Dim url As String
            Dim params As String
            Dim params2 As String
            Dim SmsMblNo As String
            Dim SmsMblNoBrk As String
            Dim SmsString As String
            Dim MillShort As String
            Dim MillCode As Long
            Dim TktNm As String
            Dim PtyName As String

            MillCode = GProcGetColumnValue("tblMastAccount", "AcName", txttemp(6), "S", "AcCode", "N")

            TktNm = GProcGetColumnValue("tblMastItem", "ItName", txttemp(17), "S", "ItTicket", "S", " ItMillCode=" & MillCode) '--Item Code




            SmsMblNo = lblMblNo.Caption
            
'---For WhatsApp Text SMS
            
            If Val(Right(Trim(txttemp(5)), 2)) > 0 Then
               PtyName = Trim(Left(Trim(txttemp(5)), Len(Trim(txttemp(5))) - 2))
            Else
               PtyName = Trim(txttemp(5))
            End If
            
            If Val(Right(Trim(txttemp(6)), 2)) > 0 Then
               TktNm = Trim(Left(Trim(txttemp(6)), Len(Trim(txttemp(6))) - 2))
            Else
               TktNm = Trim(txttemp(6))
            End If

            
            
            
            SmsString = "Firm   :" & vbLf
            SmsString = SmsString + "*" + gCName + "*" & vbLf & vbLf
            SmsString = SmsString + "*Local Booking*" & vbLf & vbLf
            SmsString = SmsString + "No     : *" + Trim(txt(0)) + "*" & vbLf
            SmsString = SmsString + "Date  : " + Trim(CStr(Format(dtpDt(2), "dd/mm/yy"))) & vbLf & vbLf
            
            SmsString = SmsString + "Party  : *" + PtyName + "*" & vbLf & vbLf
            
            SmsString = SmsString + "Broker : " + txttemp(7) & vbLf
            SmsString = SmsString + "Mill   : " + TktNm & vbLf & vbLf
            SmsString = SmsString + "Count  : " & vbLf
            SmsString = SmsString + "*" + Trim(txttemp(17)) + "*" & vbLf & vbLf
            
            SmsString = SmsString + "Qty.    : " + txt(18) & vbLf
            SmsString = SmsString + "Kgs.   : " + txt(19) & vbLf
            
            SmsString = SmsString + "Rate   : *" + txt(20) + "*" & vbLf
            SmsString = SmsString + "( Plus GST and TCS if applicable)" & vbLf
            
            
            
            
'            SmsString = "BO-" + Trim(txt(0)) + " Dt-" + Trim(CStr(Format(dtpDt(2), "dd/mm/yy"))) & vbLf
'            SmsString = SmsString + "Pty-" + Left(txttemp(5), 20) & vbLf
'            SmsString = SmsString + "Cnt-" + Left(txttemp(17), 15) & vbLf
'            If TktNm = "" Or IsNull(TktNm) Then
'               SmsString = SmsString + "Pkg-" + txt(18) + "-" + Left(txttemp(6), 10) & vbLf
'            Else
'               SmsString = SmsString + "Pkg-" + txt(18) + "-" + Left(TktNm, 10) & vbLf
'            End If
'            SmsString = SmsString + "Rate-" + txt(20) & vbLf
'            SmsString = SmsString + "Brk-" + Left(txttemp(7), 15) & vbLf
'            SmsString = SmsString + "Firm-" + Left(gCName, 15)
            
            
            
'            Dim PtyName As String
'            Dim RsSA As Recordset
            Dim RsTmp As Recordset
            Dim X As String
            
            gCn.BeginTrans
            gCn.Execute "Delete from tmpWhatsAppErr where username = '" & gUserName & "' and Vfirm = '" & gCCode & "' "
            gCn.CommitTrans
            If SmsMblNo = "" Then
               MsgBox ("Party Mobile No Not Found....")
            Else
              PtyName = txttemp(5)
              SentWhatsApp "Txt", "", SmsString, SmsMblNo, PtyName
            End If
            '-- SMS To Broker
            
            SmsMblNoBrk = GProcGetColumnValue("TblMastAccount", "ACName", txttemp(7), "S", "AcContNo", "S")
            
            If SmsMblNoBrk = "" Then
               MsgBox ("Broker Mobile No Not Found....")
            Else
              PtyName = txttemp(7)
              SentWhatsApp "Txt", "", SmsString, SmsMblNoBrk, PtyName
            End If
           
            '--Error
            X = "Select * from tmpWhatsAppErr where username = '" & gUserName & "' and Vfirm = '" & gCCode & "'"
            If GProcRstOpen(RsTmp, X, "R", gCn) > 0 Then
               GProcWhatsAppError
            End If

'---End For WhatsApp Text SMS
'            SmsString = "BO-" + Trim(txt(0)) + " Dt-" + Trim(CStr(Format(dtpDt(2), "dd/mm/yy"))) & vbLf
'            SmsString = SmsString + "Pty-" + Left(txttemp(5), 20) & vbLf
'            SmsString = SmsString + "Cnt-" + Left(txttemp(17), 15) & vbLf
'            If TktNm = "" Or IsNull(TktNm) Then
'               SmsString = SmsString + "Pkg-" + txt(18) + "-" + Left(txttemp(6), 10) & vbLf
'            Else
'               SmsString = SmsString + "Pkg-" + txt(18) + "-" + Left(TktNm, 10) & vbLf
'            End If
'            SmsString = SmsString + "Rate-" + txt(20) & vbLf
'            SmsString = SmsString + "Brk-" + Left(txttemp(7), 15) & vbLf
'            SmsString = SmsString + "Firm-" + Left(gCName, 15)

'''            If SmsMblNo = "" Then
'''               MsgBox ("Party Mobile No Not Found....")
'''            Else
'''                url = "https://api-alerts.kaleyra.com/v4/?"
'''                params = "api_key=A967d646cc4202f5c44523b468617f117&method=sms&message=" & SmsString & "&to=" & SmsMblNo & "&sender=VISHTX"
'''                HttpReq.Open "POST", url, False
'''                HttpReq.setRequestHeader "Content-Type", "application/x-www-form-urlencoded"
'''                HttpReq.send params
'''            End If
'''            '-- SMS To Broker
'''            SmsMblNoBrk = GProcGetColumnValue("TblMastAccount", "ACName", txttemp(7), "S", "AcContNo", "S")
'''            If SmsMblNoBrk = "" Then
'''               MsgBox ("Broker Mobile No Not Found....")
'''            Else
'''                url = "https://api-alerts.kaleyra.com/v4/?"
'''                params = "api_key=A967d646cc4202f5c44523b468617f117&method=sms&message=" & SmsString & "&to=" & SmsMblNoBrk & "&sender=VISHTX"
'''                HttpReq.Open "POST", url, False
'''                HttpReq.setRequestHeader "Content-Type", "application/x-www-form-urlencoded"
'''                HttpReq.send params
'''            End If
        End If
   
   
   
   
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
    'If FormAction = vbDataActionAddNew Then
        If cmdBtn(6).Visible = True Then cmdBtn(6).SetFocus
    'Else
    '    If cmdBtn(7).Visible = True Then cmdBtn(7).SetFocus
    'End If
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
If FormAction = vbDataActionAddNew Then
    txt(10).Enabled = False
    dtpDt(29).Enabled = False
    txt(27).Enabled = False
    txt(28).Enabled = False
    txt(21).Enabled = False
End If
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
       If (Val(.TextMatrix(RowIndex, 0)) = 0) And .rows > 2 Then
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
               rstSub!BkCDate = CDate(dtpDt(29))
            End If
            rstSub!BkSrNo = i '--Sr No
            If Val(.TextMatrix(i, 2)) = 0 Then
               rstSub!BkShDate = CDate(txt(2)) '--SH.Date
               rstSub!BkShBag = Val(txt(18)) '--Bag(2)
               rstSub!BkShWt = Val(txt(19)) '-- Wt(3)
            Else
                If CLng(.TextMatrix(i, 2)) > 0 Then
                   rstSub!BkShDate = CDate(.TextMatrix(i, 1)) '--SH.Date
                Else
                   rstSub!BkShDate = Null  '--Des.Date
                End If
                rstSub!BkShBag = CLng(.TextMatrix(i, 2)) '--Bag(2)
                rstSub!BkShWt = CDbl(.TextMatrix(i, 3)) '-- Wt(3)
            End If
            rstSub!BkTmpChr1 = txt(37) & vbNullString
            rstSub!BkTmpChr2 = txt(38) & vbNullString
            rstSub!BkIsDirectPayment = Val(txt(41))
            rstSub!BkIsFrghtInsu = Val(txt(42))
            rstSub!BkIsInsu = Val(txt(43))
            rstSub.Update
'        End If
        i = i + 1
     Loop
 End With
 rstSub.Close
End Sub

Private Sub CmdClose_Click()
    DoEvents
    Frame2.Visible = False
    txtCode(7).SetFocus
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
        & " BkCDate,BkTmpNo1,BkTmpNo2,BkTmpNo3,BkTmpChr1,BkTmpChr2,BkTmpChr3,BkShBag,BkShWt,BkShDate,a.BkIsDirectPayment,a.BkIsFrghtInsu,a.BkIsInsu " _
        & " From tblBooking A  where a.vno in " & X & " and a.Vtype='" & frmTypeStr & "' and a.VFirm = '" & gCCode & "' And a.VYear = " & gCYear
     
   gCn.Execute "Delete from TmpBooking where  BkUser='" & gUserName & "'"
     
   gCn.Execute " insert into TmpBooking (Vno,Vtype,Vdt,Vyear,VFirm,BkParty,BkMillCode,BkBroker,BkRefNo,BkRefDt,BkSubAmt,BkBillAmt,BkRoff,DueDays,DueDate,BkEntDt,BkUser,BkItCode,BkBag,BkWt,BkRt,BkRtPer,BkRateType,BkDeleAdd,BkPartyGroup,BkExtraCount,BkNote,BkCbag,BkCWt," _
              & " BkCDate,BkTmpNo1,BkTmpNo2,BkTmpNo3,BkTmpChr1,BkTmpChr2,BkTmpChr3,BkSubBag,BkSubWt,BkSubDDate,BkIsDirectPayment,BkIsFrghtInsu,BkIsInsu ) " & X
     
     
     
     X = "select * from TmpBooking where BkUser='" & gUserName & "'"
     With cryReport
        .Formulas(10) = "wCSTNo = '" & gCCST & "'"
        .Formulas(11) = "wBSTNo = '" & gCBST & "'"
        X = "{TmpBooking.BkUser}='" & gUserName & "' "
       .SelectionFormula = X
    '    .SelectionFormula = "{TmpBooking.Vno} = " & RsSA!Vno & " and {TmpBooking.VType}='" & RsSA!Vtype & "' and {TmpBooking.VFirm} = '" & RsSA!VFirm & "' and  {TmpBooking.VYear}=" & RsSA!VYear
        .ReportFileName = gReportPath & "rptBookingPrint_Party.rpt"
         If Index = 0 Or Index = 1 Then
            If ChkCus.Value = 1 Then
               .Formulas(12) = "wCopy = 'Customer Copy'"
               .Action = 1
            End If
            If chkOff.Value = 1 Then
               .Formulas(12) = "wCopy = 'Office Copy'"
               .Action = 1
            End If
         End If
         If Index = 3 Then
            Dim objCrystal As CRAXDRT.Application
            Dim objReport As CRAXDRT.Report
            Dim Tattach As String
            Dim Mfile As String
            Dim Mcp As String
            Dim PtyMailId As String
            Dim PtyMailIdCc As String
            Dim PtyMailIdBcc As String
            Dim MailSub As String
            Dim MailStr As String
            Set objCrystal = New CRAXDRT.Application
            Tattach = gReportPath & "rptBookingPrint_Party.rpt"
            Set objReport = objCrystal.OpenReport(Tattach, 1)
            objReport.RecordSelectionFormula = "{TmpBooking.BkUser}='" & gUserName & "'"
            objReport.Database.LogOnServer "pdsodbc.dll", "Hitrix", "Vishnu", "sa", "123456"
            
            Dim j As Integer
                For j = 1 To objReport.FormulaFields.Count
                    Select Case objReport.FormulaFields(j).Name
                    Case "{@FirmNm}"
                        objReport.FormulaFields(j).text = "Trim(" & Chr(39) & UCase(gCName) & Chr(39) & ")"
                    Case "{@wAdd1}"
                        objReport.FormulaFields(j).text = "Trim(" & Chr(39) & gCAdd1 & Chr(39) & ")"
                    Case "{@wAdd2}"
                        objReport.FormulaFields(j).text = "Trim(" & Chr(39) & gCAdd2 & Chr(39) & ")"
'                    Case "{@wAdd3}"
'                        objReport.FormulaFields(j).text = "Trim(" & Chr(39) & gCAdd3 & Chr(39) & ")"
                    Case "{@wPhNo}"
                        objReport.FormulaFields(j).text = "Trim(" & Chr(39) & gCPhNo & Chr(39) & ")"
                   End Select
                   Next j
                   If Index = 3 Then
                       X = "Select  AcEmail from tblBooking,tblMastAccount where tblBooking.BkMillCode =  tblMastAccount.AcCode and  tblBooking.Vno = " & CLng(txtPInvno(0)) & " and tblBooking.Vtype='" & frmTypeStr & "' and tblBooking.VFirm = '" & gCCode & "' And tblBooking.VYear = " & gCYear
                   Else
                       X = "Select  AcEmail from tblBooking,tblMastAccount where tblBooking.BkParty =  tblMastAccount.AcCode and  tblBooking.Vno = " & CLng(txtPInvno(0)) & " and tblBooking.Vtype='" & frmTypeStr & "' and tblBooking.VFirm = '" & gCCode & "' And tblBooking.VYear = " & gCYear
                   End If
                   i = GProcRstOpen(RsSA, X, "R")
                   If i = 0 Or RsSA.Fields(0) = "" Or IsNull(RsSA.Fields(0)) Then
                      MsgBox "Mail ID Not Found"
                      Exit Sub
                   End If
                   PtyMailId = RsSA.Fields(0)
                   Mfile = CStr(App.Path) + "\MailPDF\" + gCCode + "-" + gUserName + "-" + "Party_Booking"
                   ExportReportToPDF objReport, Mfile & ".pdf", "foo"
                   
                   Mfile = CStr(App.Path) + "\MailPDF\" + gCCode + "-" + gUserName + "-" + "Party_Booking.pdf"
                   
                   MailSub = "Party Booking Order No  " & YY
                   MailStr = "Dear Sir, " + Chr(13) + Chr(10)
                   MailStr = MailStr + "Please find attached herewith Party Booking Order" + Chr(13) + Chr(10)
                   MailStr = MailStr + "Note : This is system generated email, Please do not reply."

'                   MailStr = MailStr + " " & txtMailDesc & " " & vbLf

                   gMailsendToMill = False
                   Call SendEmail1("" & Trim(gCMailId) & "", "" & PtyMailId & "", MailSub, MailStr, "" & PtyMailIdCc & "", "" & PtyMailIdBcc & "", Mfile)
                   If gMailsendToMill = True And Index = 3 Then
                       gCn.Execute "update tblBooking set  BkMailSendDt = getdate(),    BkUser = '" & gUserName & "'  " _
                       & " where   vno in " & Y & " and    Vtype='" & frmTypeStr & "' and VFirm = '" & gCCode & "' And VYear = " & gCYear
                      GProcRstOpen rstbl, "Select * from " & tblName & " where " & IdField & " = " & txt(0) & " and " & Criteria, "R"
                       ReadFields
                   End If
                   If gMailsendToMill = True And Index = 4 Then
                       gCn.Execute "update tblBooking set  BkmailSendPartryDt = getdate(),    BkUser = '" & gUserName & "'  " _
                       & " where   vno in " & Y & " and    Vtype='" & frmTypeStr & "' and VFirm = '" & gCCode & "' And VYear = " & gCYear
                      GProcRstOpen rstbl, "Select * from " & tblName & " where " & IdField & " = " & txt(0) & " and " & Criteria, "R"
                       ReadFields
                   End If
         End If
     End With
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
If FormAction = vbDataActionAddNew And txt(26) = "" Then
   txt(26) = "Earliest"
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
      Cbo(22).ListIndex = 1
   End If
End If
CheckForNumber
CalAmount
SaveBtnEd
If Index = 19 And FormAction = vbDataActionAddNew Then
    msGrid.TextMatrix(1, 2) = txt(18)
    msGrid.TextMatrix(1, 3) = txt(19)
End If
If Index = 38 Then
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
Dim tmpRst As Recordset
Dim X As String
Dim wIsNoWork As Integer
SetControlEd
Select Case Index
    Case 5  '-- Party  Code
         txttemp(5) = GProcGetColumnValue("tblMastAccount", "AcOurCode", UCase(txtCode(5)), "S", "AcName", "S")
         lblTinNo.Caption = ""
         lblMsOms.Caption = ""
         lblPan.Caption = ""
         lblMblNo.Caption = ""
         If txttemp(5) <> "" Then
            lblTinNo.Caption = GProcGetColumnValue("TblMastAccount", "ACName", txttemp(Index), "S", "AcGSTIN", "S")
            lblPan.Caption = GProcGetColumnValue("TblMastAccount", "ACName", txttemp(Index), "S", "AcPAN", "N")
            lblMblNo.Caption = GProcGetColumnValue("TblMastAccount", "ACName", txttemp(Index), "S", "AcContNo", "N")
            wIsNoWork = GProcGetColumnValue("tblMastAccount", "AcOurCode", UCase(txtCode(5)), "S", "AcIsNonWork", "N")
            If wIsNoWork Then
               Frame2.Visible = True
               cmdClose.SetFocus
            Else
                txt(5) = GProcGetColumnValue("tblMastAccount", "AcName", txttemp(5), "S", "AcCode", "N")
                gCmd.CommandText = "PrcPrepareOutStangingSale"
                gCmd.Parameters.Refresh
                gCmd.Parameters("@To_dt") = Format(dtpDt(2), "yyyy/mm/dd")
                gCmd.Parameters("@UCode") = gUserName
                gCmd.Parameters("@IntDbAc") = gLatePayIntAcCodeRecGST
                gCmd.Parameters("@IntDbAcWoGST") = gLPIntAcCodeWithoutGST
                gCmd.Parameters("@IntTDSAc") = gTDSAcCodeRec
                gCmd.Parameters("@IntCrAc") = gLatePayIntAcCodePay
                gCmd.Parameters("@VFirm") = ""
                gCmd.Parameters("@DueOn") = "D"
                gCmd.Parameters("@Accode") = txt(5)
                gCmd.Parameters("@X_dt1") = Format(dtpDt(2), "yyyy/mm/dd")
                gCmd.Parameters("@X_dt2") = Format(dtpDt(2), "yyyy/mm/dd")
                gCmd.Execute
                X = "Select VFirmName,VBillNo,VBillDt,Amt1,Amt3,Amt2,amt1+amt3-Amt2 as Os,Amt5 from TmpGentbl where UserName = '" & gUserName & "' and amt5 > 60 and  amt1+amt3-Amt2 > amt3 "
                i = GProcRstOpen(tmpRst, X, "R")
                If i > 0 Then
                   If MsgBox("Total O/s Bills Over 60 days are " + Str(i) + " Do You Want To See O/s List ? ", vbYesNo) = vbYes Then
                      gClsSearch.SearchMultiField X, "VFirmName,VBillNo,VBillDt,Amt1,Amt3,Amt2,os,Amt5", Array("Firm Name", "Bill No", "Date", "Inv.Amt", "Debit Amt", "Recd.Amt", "O/s Amt", "L.Days"), Array(2000, 2000, 1800, 1500, 1500, 1500, 1500, 1000), "", "", "VFirmName,VBillDt,VBillNo ", 0, 0, False, , True
                   End If
                End If
            End If
         End If
         '--- Broker From Master
         If txttemp(7) = "" Then
            BrkAcCode = GProcGetColumnValue("TblMastAccount", "ACName", txttemp(Index), "S", "AcBrkCode", "N")
            If BrkAcCode = 1 Then
               txttemp(7) = ""
               txtCode(7) = ""
            Else
               txttemp(7) = GProcGetColumnValue("TblMastAccount", "ACCode", CStr(BrkAcCode), "N", "ACname", "S")
               txtCode(7) = GProcGetColumnValue("TblMastAccount", "ACCode", CStr(BrkAcCode), "N", "AcOurCode", "S")
               wIsNoWork = GProcGetColumnValue("tblMastAccount", "AcOurCode", UCase(txtCode(7)), "S", "AcIsNonWork", "N")
               If wIsNoWork Then
                  Frame2.Visible = True
                  cmdClose.SetFocus
               End If
            End If
         End If
         If Trim(txttemp(5)) = "" Then
'            MsgBox ("Code Not Found ...")
            txttemp(5) = ""
            txttemp(5).Enabled = True
            txttemp(5).SetFocus
         Else
            txtCode(7).SetFocus
         End If
    Case 6  '-- Mill  Code
         txttemp(6) = GProcGetColumnValue("tblMastAccount", "AcOurCode", UCase(txtCode(6)), "S", "AcName", "S")
         If Trim(txttemp(6)) = "" Then
            MsgBox ("Code Not Found ...")
            txttemp(6) = ""
            txttemp(6).Enabled = True
            txttemp(6).SetFocus
         Else
            txttemp(17).SetFocus
         End If
         If mLastEntryMill <> txttemp(6) And FormAction = vbDataActionAddNew Then
            txttemp(17) = ""
            txtStdPack = 0
            txt(20) = 0
         End If
         X = "Select Acname from tblMastAccount,tblMastBillSerial where tblMastaccount.Accode = tblMastBillSerial.MillCode  and tblMastBillSerial.Vfirm = '" & gCCode & "'   and  tblMastaccount.AcOurCode = '" & txtCode(6) & "'   "
          i = GProcRstOpen(tmpRst, X, "R")
          If i = 0 Then
             txttemp(6) = ""
              MsgBox ("Mill Not in Firm")
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
               wIsNoWork = GProcGetColumnValue("tblMastAccount", "AcOurCode", UCase(txtCode(7)), "S", "AcIsNonWork", "N")
               If wIsNoWork Then
                  Frame2.Visible = True
                  cmdClose.SetFocus
               End If
            
            If txtCode(6).Enabled = True Then txtCode(6).SetFocus
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
Dim wIsNoWork As Integer
Dim X As String
Dim tmpRst As Recordset
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
        If gClsSearch.SearchMultiRetColAccMaster(0) <> "" Then
            txttemp(Index) = gClsSearch.SearchMultiRetColAccMaster(0)
            txtCode(5) = GProcGetColumnValue("TblMastAccount", "ACName", txttemp(Index), "S", "AcOurCode", "S")
            lblTinNo.Caption = GProcGetColumnValue("TblMastAccount", "ACName", txttemp(Index), "S", "AcGSTIN", "S")
            lblPan.Caption = GProcGetColumnValue("TblMastAccount", "ACName", txttemp(Index), "S", "AcPAN", "N")
            lblMblNo.Caption = GProcGetColumnValue("TblMastAccount", "ACName", txttemp(Index), "S", "AcContNo", "N")
            wIsNoWork = GProcGetColumnValue("tblMastAccount", "AcOurCode", UCase(txtCode(5)), "S", "AcIsNonWork", "N")
            If wIsNoWork Then
               Frame2.Visible = True
               cmdClose.SetFocus
            Else
                txt(5) = GProcGetColumnValue("tblMastAccount", "AcName", txttemp(5), "S", "AcCode", "N")
                gCmd.CommandText = "PrcPrepareOutStangingSale"
                gCmd.Parameters.Refresh
                gCmd.Parameters("@To_dt") = Format(dtpDt(2), "yyyy/mm/dd")
                gCmd.Parameters("@UCode") = gUserName
                gCmd.Parameters("@IntDbAc") = gLatePayIntAcCodeRecGST
                gCmd.Parameters("@IntDbAcWoGST") = gLPIntAcCodeWithoutGST
                gCmd.Parameters("@IntTDSAc") = gTDSAcCodeRec
                gCmd.Parameters("@IntCrAc") = gLatePayIntAcCodePay
                gCmd.Parameters("@VFirm") = ""
                gCmd.Parameters("@DueOn") = "D"
                gCmd.Parameters("@Accode") = txt(5)
                gCmd.Parameters("@X_dt1") = Format(dtpDt(2), "yyyy/mm/dd")
                gCmd.Parameters("@X_dt2") = Format(dtpDt(2), "yyyy/mm/dd")
                gCmd.Execute
                X = "Select VFirmName,VBillNo,VBillDt,Amt1,Amt3,Amt2,amt1+amt3-Amt2 as Os,Amt5 from TmpGentbl where UserName = '" & gUserName & "' and amt5 > 60 and  amt1+amt3-Amt2 > amt3 "
                i = GProcRstOpen(tmpRst, X, "R")
                If i > 0 Then
                   If MsgBox("Total O/s Bills Over 60 days are " + Str(i) + " Do You Want To See O/s List ? ", vbYesNo) = vbYes Then
                      gClsSearch.SearchMultiField X, "VFirmName,VBillNo,VBillDt,Amt1,Amt3,Amt2,os,Amt5", Array("Firm Name", "Bill No", "Date", "Inv.Amt", "Debit Amt", "Recd.Amt", "O/s Amt", "L.Days"), Array(2000, 2000, 1800, 1500, 1500, 1500, 1500, 1000), "", "", "VFirmName,VBillDt,VBillNo ", 0, 0, False, , True
                   End If
                End If
            End If
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
            ShowMill = "AgCode=90029 and   tblMastaccount.Accode = tblMastBillSerial.MillCode  and tblMastBillSerial.Vfirm = '" & gCCode & "'  "
            gClsSearch.SearchMultiField "tblMastaccount,tblMastBillSerial", "BillSr,AcName,AcOurCode", Array("Mill Ini.", "Mill Name", "Code"), Array(1000, 8000, 1000), ShowMill, Trim(Chr(KeyAscii)), "AcName", LeftPos, TopPos, True
'        gClsSearch.SearchMultiField "tblMastaccount", "AcName,AcCity,AcOurCode", Array("Mill Name", "City", ""), Array(7000, 2000, 0), "(AgCode=90029  and AcMillBillFirm = '" & gCCode & "') ", Trim(Chr(KeyAscii)), "AcName", LeftPos, TopPos, True
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
        gClsSearch.SearchMultiField "tblMastaccount", "AcName,AcOurCode,AcCity", Array("Code", "Broker", "City"), Array(5000, 1000, 3000), "(AgCode=90016 )", Trim(Chr(KeyAscii)), "", LeftPos, TopPos, True
        KeyAscii = 0
        txtCode(7) = ""
        If gClsSearch.SearchMultiRetCol(0) <> "" Then
            txttemp(Index) = gClsSearch.SearchMultiRetCol(0)
            txtCode(7) = gClsSearch.SearchMultiRetCol(1)
            wIsNoWork = GProcGetColumnValue("tblMastAccount", "AcOurCode", UCase(txtCode(7)), "S", "AcIsNonWork", "N")
            If wIsNoWork Then
               Frame2.Visible = True
               cmdClose.SetFocus
            End If
        End If
    Case 17  '-- Item
         MillCode = GProcGetColumnValue("TblMastAccount", "AcName", txttemp(6), "S", "Accode", "N")
         gClsSearch.SearchMultiField "tblMastItem,tblMastaccount", "ItName,AcName,ItStdpack,ItStdrateper,InDailyReport", Array("Count Name", "Mill", "Std. Pack", "Rate Per", ""), Array(5000, 5000, 1000, 1000, 0), "tblMastItem.ITMillCode=tblMastaccount.acCode and ITMillCode=" & MillCode, Chr(KeyAscii), "ItName", LeftPos, TopPos, True
         If gClsSearch.SearchMultiRetCol(4) = 1 Then
            MsgBox "This is No Use Count "
         End If
         KeyAscii = 0
         If gClsSearch.SearchMultiRetCol(0) <> "" And gClsSearch.SearchMultiRetCol(4) = 0 Then
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
        gClsSearch.SearchMultiField "tblMastDeleAdd", "DCity,ContactPer,DAdd1,DAdd2,DAdd3,DZip,DTal,DDist,DState,DPhone", Array("City", "Contanc To", "Add1", "Add2", "Add3", "", "", "", "", ""), Array(2500, 3000, 4000, 4000, 3000, 0, 0, 0, 0, 0), "PartyCode = " & wPartyCode & "  ", Chr(KeyAscii), "DCity", 0, 0, True
        KeyAscii = 0
End Select
End Sub
Private Sub txtTemp_LostFocus(Index As Integer)
'If Index = 7 Then txtCode(6).SetFocus
Dim MillCode As Long
MillCode = GProcGetColumnValue("TblMastAccount", "AcName", txttemp(6), "S", "Accode", "N")
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
If TotBag > Val(txt(18)) And TotBag > 0 Then
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
        If (TotBag > Val(txt(18))) Or (TotBag > 0 And TotBag < Val(txt(18))) Then
           MsgBox "Check Sh.Bags Total.", vbCritical + vbOKOnly, Me.Caption
           msGrid.Col = 2
           txtGrid.SetFocus
           Exit Sub
        End If
        txtGrid.Visible = False
        If cmdBtn(13).Enabled = True Then
           cmdBtn(13).SetFocus
        ElseIf cmdBtn(14).Enabled = True Then
           cmdBtn(14).SetFocus
        End If
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
If TotBag <> Val(txt(18)) And TotBag > 0 Then
   MsgBox "Check Despatch Detail Bag Total.", vbCritical + vbOKOnly, Me.Caption
   Exit Function
End If

If Round(TotWt, 3) <> Round(Val(txt(19)), 3) And TotWt > 0 Then
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
    If Val(txt(0)) < 0 Then
        cmdBtn(13).Enabled = True
        cmdBtn(13).BackColor = &HFFFF80
    Else
        cmdBtn(13).Enabled = False
        cmdBtn(13).BackColor = &H80000004
    End If
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

If Len(Trim(lblTinNo.Caption)) = 0 Then
      MsgBox "Party GSTIN Not Found", vbInformation + vbOKOnly, Me.Caption
      txttemp(5).SetFocus
      Exit Function
End If


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
    If GProcIsDateValid(dtpDt(2)) = False And Val(txt(0)) > 0 Then
        MsgBox "Check Booking Date.", vbInformation + vbOKOnly, Me.Caption
        dtpDt(2).SetFocus
        Exit Function
    End If
    '---Rate Type
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
    If GProcGetColumnValue("TblMastItem", "ItName", txttemp(17), "S", "InDailyReport", "N", "ITMillCode=" & ExpAcCode) = 1 Then
       MsgBox "This is No Use Count "
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
'--Count
txt(17) = GProcGetColumnValue("tblMastItem", "ItName", txttemp(17), "S", "Itcode", "N", "ItMillCode = " & txt(6) & "")
txt(3) = gCYear
txt(4) = gCCode
txt(22) = Cbo(22).ListIndex
txt(23) = 1
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
If OldBag2 <> Val(txt(18)) And StdPack > 0 Then  'And FormAction = vbDataActionAddNew
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

