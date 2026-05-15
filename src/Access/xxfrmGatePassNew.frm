VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form frmGatepassNew 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00C2E3FC&
   BorderStyle     =   0  'None
   Caption         =   "Gate Pass Entry"
   ClientHeight    =   8370
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   10830
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   8370
   ScaleWidth      =   10830
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox txtGrid 
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
      Left            =   7620
      TabIndex        =   126
      Top             =   4260
      Visible         =   0   'False
      Width           =   1035
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFC0C0&
      Caption         =   "Gate Pass Printing"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3945
      Left            =   1710
      TabIndex        =   120
      Top             =   1560
      Visible         =   0   'False
      Width           =   5175
      Begin VB.CommandButton cmdPrint 
         Caption         =   "&Whatsapp - PDF"
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
         Left            =   2700
         TabIndex        =   124
         Top             =   3330
         Width           =   2250
      End
      Begin VB.CheckBox ChkRt 
         BackColor       =   &H00FFC0C0&
         Caption         =   "Print Rate ?"
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
         Left            =   2835
         TabIndex        =   123
         Top             =   1575
         Width           =   2175
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
         ForeColor       =   &H000000FF&
         Height          =   315
         Left            =   540
         TabIndex        =   4
         Top             =   2025
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
         ForeColor       =   &H000000FF&
         Height          =   375
         Left            =   2730
         TabIndex        =   5
         Top             =   2025
         Width           =   2055
      End
      Begin VB.TextBox txtno 
         Appearance      =   0  'Flat
         Height          =   360
         Index           =   0
         Left            =   2640
         MaxLength       =   7
         TabIndex        =   1
         Top             =   600
         Width           =   1695
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
         Left            =   540
         TabIndex        =   6
         Top             =   2625
         Width           =   1335
      End
      Begin VB.TextBox txtno 
         Appearance      =   0  'Flat
         Height          =   360
         Index           =   1
         Left            =   2640
         MaxLength       =   7
         TabIndex        =   2
         Top             =   1080
         Width           =   1695
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
         Left            =   2085
         TabIndex        =   7
         Top             =   2625
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
         Left            =   3630
         TabIndex        =   8
         Top             =   2625
         Width           =   1335
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
         Left            =   540
         TabIndex        =   3
         Top             =   1560
         Width           =   2220
      End
      Begin VB.CommandButton cmdPrint 
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
         Index           =   3
         Left            =   315
         TabIndex        =   9
         Top             =   3330
         Width           =   2250
      End
      Begin VB.Label Label31 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "From G P No :"
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
         Left            =   480
         TabIndex        =   122
         Top             =   600
         Width           =   1335
      End
      Begin VB.Label Label18 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "To G P No     :"
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
         Top             =   1080
         Width           =   1320
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
      Left            =   3960
      TabIndex        =   29
      Top             =   6180
      Width           =   3135
   End
   Begin VB.CheckBox Chk 
      BackColor       =   &H00ECFFFF&
      Caption         =   "Print Ticket Name ?"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H004D177D&
      Height          =   375
      Index           =   0
      Left            =   3960
      TabIndex        =   32
      Top             =   6960
      Width           =   2055
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
      Index           =   45
      Left            =   16020
      TabIndex        =   116
      Text            =   "45"
      Top             =   4020
      Visible         =   0   'False
      Width           =   390
   End
   Begin VB.TextBox txtAL 
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
      Index           =   0
      Left            =   16020
      TabIndex        =   115
      Text            =   "txtAl(0)"
      Top             =   4800
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
      Height          =   360
      Index           =   19
      Left            =   16020
      TabIndex        =   44
      Text            =   "19"
      Top             =   4380
      Visible         =   0   'False
      Width           =   390
   End
   Begin VB.TextBox txtE2 
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
      Left            =   15600
      TabIndex        =   113
      Text            =   "E2"
      Top             =   6480
      Visible         =   0   'False
      Width           =   390
   End
   Begin VB.CheckBox Chk 
      BackColor       =   &H00ECFFFF&
      Caption         =   "Is Export?"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H004D177D&
      Height          =   375
      Index           =   1
      Left            =   1200
      TabIndex        =   31
      Top             =   6960
      Width           =   2055
   End
   Begin VB.CommandButton CmdOutstandingRpt 
      Caption         =   "&Outstanding Report"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   660
      Left            =   5910
      TabIndex        =   112
      Top             =   1995
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
      Index           =   44
      Left            =   15600
      TabIndex        =   111
      Text            =   "44"
      Top             =   1230
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
      Height          =   360
      Index           =   42
      Left            =   2190
      TabIndex        =   20
      Text            =   "42"
      Top             =   3360
      Width           =   1365
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
      Left            =   16005
      TabIndex        =   110
      Text            =   "43"
      Top             =   375
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
      Height          =   360
      Index           =   41
      Left            =   16005
      TabIndex        =   109
      Text            =   "41"
      Top             =   720
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
      Height          =   360
      Index           =   40
      Left            =   16005
      TabIndex        =   108
      Text            =   "40"
      Top             =   1140
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
      Height          =   360
      Index           =   39
      Left            =   15600
      TabIndex        =   107
      Text            =   "39"
      Top             =   4875
      Visible         =   0   'False
      Width           =   435
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
      Index           =   38
      Left            =   15600
      TabIndex        =   106
      Text            =   "38"
      Top             =   5310
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
      Height          =   360
      Index           =   37
      Left            =   15600
      TabIndex        =   105
      Text            =   "37"
      Top             =   4470
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
      Index           =   36
      Left            =   15600
      TabIndex        =   104
      Text            =   "36"
      Top             =   5655
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
      Index           =   35
      Left            =   15600
      TabIndex        =   103
      Text            =   "35"
      Top             =   6075
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
      Index           =   34
      Left            =   15600
      TabIndex        =   102
      Text            =   "34"
      Top             =   3015
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
      Height          =   360
      Index           =   33
      Left            =   15600
      TabIndex        =   101
      Text            =   "33"
      Top             =   3375
      Visible         =   0   'False
      Width           =   270
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
      Left            =   15600
      TabIndex        =   100
      Text            =   "32"
      Top             =   2610
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
      Index           =   31
      Left            =   15600
      TabIndex        =   99
      Text            =   "31"
      Top             =   3720
      Visible         =   0   'False
      Width           =   270
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
      Index           =   30
      Left            =   2160
      MaxLength       =   25
      TabIndex        =   30
      Text            =   "30"
      Top             =   6600
      Width           =   4680
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
      Left            =   5445
      TabIndex        =   98
      Text            =   "29"
      Top             =   4185
      Width           =   1695
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
      Index           =   28
      Left            =   15600
      TabIndex        =   97
      Text            =   "28"
      Top             =   1755
      Visible         =   0   'False
      Width           =   315
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
      Index           =   27
      Left            =   15600
      TabIndex        =   96
      Text            =   "27"
      Top             =   795
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
      Index           =   26
      Left            =   15600
      TabIndex        =   95
      Text            =   "26"
      Top             =   4140
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
      Index           =   25
      Left            =   15600
      TabIndex        =   94
      Text            =   "25"
      Top             =   2175
      Visible         =   0   'False
      Width           =   270
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
      Left            =   15600
      TabIndex        =   93
      Text            =   "24"
      Top             =   360
      Visible         =   0   'False
      Width           =   360
   End
   Begin VB.CommandButton CmdGPDelivery 
      Caption         =   "&GP Delivery"
      Height          =   400
      Left            =   9120
      TabIndex        =   92
      Top             =   1935
      Width           =   1095
   End
   Begin VB.TextBox txttmpLRAmt 
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
      Left            =   5445
      TabIndex        =   39
      Text            =   "txttmpLRAmt"
      Top             =   5010
      Width           =   1695
   End
   Begin VB.TextBox txttmpLRDt 
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
      Left            =   5445
      TabIndex        =   38
      Text            =   "txttmpLRDt"
      Top             =   4605
      Width           =   1695
   End
   Begin VB.TextBox txttmpLRNo 
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
      Left            =   16005
      TabIndex        =   37
      Text            =   "txttmpLRNo"
      Top             =   3630
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
      Height          =   360
      Index           =   8
      Left            =   2190
      TabIndex        =   28
      Text            =   "8"
      Top             =   6210
      Width           =   1695
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
      Index           =   7
      Left            =   2190
      TabIndex        =   26
      Text            =   "7"
      Top             =   5805
      Width           =   1695
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
      Index           =   17
      Left            =   2190
      MaxLength       =   40
      TabIndex        =   22
      Text            =   "txttemp(17)"
      Top             =   3765
      Width           =   4950
   End
   Begin Crystal.CrystalReport CryReport 
      Left            =   3975
      Top             =   2985
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrintFileLinesPerPage=   60
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
      Left            =   16005
      TabIndex        =   90
      Text            =   "23"
      Top             =   1635
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
      Height          =   360
      Index           =   22
      Left            =   16005
      TabIndex        =   89
      Text            =   "22"
      Top             =   2085
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
      Height          =   360
      Index           =   21
      Left            =   16005
      TabIndex        =   88
      Text            =   "21"
      Top             =   2475
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
      Height          =   360
      Index           =   20
      Left            =   16005
      TabIndex        =   87
      Text            =   "20"
      Top             =   2790
      Visible         =   0   'False
      Width           =   390
   End
   Begin VB.CommandButton cmdLedger 
      Caption         =   "&Ledger"
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
      Left            =   5910
      TabIndex        =   86
      Top             =   1575
      Width           =   1230
   End
   Begin VB.TextBox txttempAL 
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
      Index           =   3
      Left            =   1560
      MaxLength       =   40
      TabIndex        =   16
      Text            =   "txttempAL(3)"
      ToolTipText     =   "Press F3 for Creation"
      Top             =   1575
      Width           =   4140
   End
   Begin VB.TextBox txttmpInNo 
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
      Left            =   2190
      TabIndex        =   23
      Text            =   "txttmpInNo"
      Top             =   4185
      Width           =   1695
   End
   Begin VB.TextBox txtAL 
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
      Index           =   6
      Left            =   5445
      TabIndex        =   27
      Text            =   "txtAl(6)"
      Top             =   5805
      Width           =   1695
   End
   Begin VB.TextBox txt 
      Height          =   345
      Index           =   18
      Left            =   16560
      TabIndex        =   85
      Text            =   "18"
      Top             =   750
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.TextBox txt 
      Height          =   345
      Index           =   17
      Left            =   16560
      TabIndex        =   84
      Text            =   "17"
      Top             =   1110
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.TextBox txt 
      Height          =   345
      Index           =   16
      Left            =   16560
      TabIndex        =   83
      Text            =   "16"
      Top             =   360
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.TextBox txt 
      Height          =   345
      Index           =   15
      Left            =   16545
      TabIndex        =   82
      Text            =   "15"
      Top             =   2310
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
      Height          =   360
      Index           =   12
      Left            =   16005
      TabIndex        =   69
      Text            =   "12"
      Top             =   3210
      Visible         =   0   'False
      Width           =   390
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
      Index           =   10
      Left            =   5445
      TabIndex        =   40
      Text            =   "10"
      Top             =   5415
      Width           =   1695
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
      Height          =   360
      Index           =   0
      Left            =   1560
      TabIndex        =   10
      Text            =   "0"
      Top             =   540
      Width           =   1695
   End
   Begin VB.ComboBox cbo 
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
      ItemData        =   "frmGatePassNew.frx":0000
      Left            =   5520
      List            =   "frmGatePassNew.frx":0007
      Style           =   2  'Dropdown List
      TabIndex        =   14
      Top             =   810
      Width           =   1620
   End
   Begin VB.TextBox txttempAL 
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
      Left            =   1560
      MaxLength       =   40
      TabIndex        =   18
      Text            =   "txttempAL(5)"
      ToolTipText     =   "Press F3 for Creation"
      Top             =   2400
      Width           =   4140
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
      Index           =   4
      Left            =   2190
      MaxLength       =   40
      TabIndex        =   19
      Text            =   "txttemp(4)"
      ToolTipText     =   "Press F3 for Creation"
      Top             =   2955
      Width           =   4950
   End
   Begin VB.TextBox txttempAL 
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
      Left            =   1560
      MaxLength       =   40
      TabIndex        =   17
      Text            =   "txtempAL(4)"
      ToolTipText     =   "Press F3 for Creation"
      Top             =   1995
      Width           =   4140
   End
   Begin VB.TextBox txt 
      Height          =   375
      Index           =   14
      Left            =   16560
      TabIndex        =   72
      Text            =   "14"
      Top             =   1500
      Width           =   465
   End
   Begin VB.TextBox txt 
      Alignment       =   1  'Right Justify
      Height          =   360
      Index           =   13
      Left            =   16560
      TabIndex        =   71
      Text            =   "13"
      Top             =   1890
      Width           =   450
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
      Left            =   9150
      Style           =   1  'Graphical
      TabIndex        =   64
      ToolTipText     =   "Previous"
      Top             =   540
      Width           =   495
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
      Left            =   9750
      Style           =   1  'Graphical
      TabIndex        =   65
      ToolTipText     =   "Next"
      Top             =   540
      Width           =   495
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
      Left            =   9150
      Style           =   1  'Graphical
      TabIndex        =   66
      ToolTipText     =   "First"
      Top             =   990
      Width           =   495
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
      Left            =   9750
      Style           =   1  'Graphical
      TabIndex        =   67
      ToolTipText     =   "Last"
      Top             =   990
      Width           =   495
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
      Left            =   9150
      Style           =   1  'Graphical
      TabIndex        =   63
      ToolTipText     =   "To Find Record"
      Top             =   1440
      Width           =   1095
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
      Left            =   7185
      Style           =   1  'Graphical
      TabIndex        =   62
      Top             =   8055
      Visible         =   0   'False
      Width           =   1095
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
      Left            =   7830
      Style           =   1  'Graphical
      TabIndex        =   58
      ToolTipText     =   "To Add New Record"
      Top             =   540
      Width           =   1095
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
      Left            =   7830
      Style           =   1  'Graphical
      TabIndex        =   59
      ToolTipText     =   "To Modify Record"
      Top             =   990
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
      Height          =   400
      Index           =   8
      Left            =   7830
      Style           =   1  'Graphical
      TabIndex        =   60
      ToolTipText     =   "To Delete Record"
      Top             =   1440
      Width           =   1095
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
      Left            =   9150
      Style           =   1  'Graphical
      TabIndex        =   61
      Top             =   2385
      Width           =   1095
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
      Left            =   7830
      Style           =   1  'Graphical
      TabIndex        =   42
      ToolTipText     =   "To Save Record"
      Top             =   1935
      Width           =   1095
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
      Left            =   7830
      Style           =   1  'Graphical
      TabIndex        =   43
      ToolTipText     =   "To Cancel Record"
      Top             =   2385
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
      Left            =   8430
      Style           =   1  'Graphical
      TabIndex        =   68
      ToolTipText     =   "Exit"
      Top             =   2910
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
      Left            =   16140
      Style           =   1  'Graphical
      TabIndex        =   73
      ToolTipText     =   "To Add New Record"
      Top             =   5220
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
      Left            =   16260
      Style           =   1  'Graphical
      TabIndex        =   74
      ToolTipText     =   "To Add New Record"
      Top             =   5700
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.TextBox txt 
      Height          =   345
      Index           =   3
      Left            =   16575
      TabIndex        =   79
      Text            =   "3"
      Top             =   3915
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
      Index           =   5
      Left            =   2190
      TabIndex        =   24
      Text            =   "5"
      Top             =   5010
      Width           =   1695
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
      Index           =   6
      Left            =   2190
      TabIndex        =   25
      Text            =   "6"
      Top             =   5415
      Width           =   1695
   End
   Begin VB.TextBox txt 
      Height          =   345
      Index           =   9
      Left            =   16515
      TabIndex        =   78
      Text            =   "9"
      Top             =   4785
      Visible         =   0   'False
      Width           =   420
   End
   Begin VB.TextBox txt 
      Height          =   345
      Index           =   11
      Left            =   16545
      TabIndex        =   75
      Text            =   "11"
      Top             =   2655
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.TextBox txt 
      Alignment       =   1  'Right Justify
      Height          =   360
      Index           =   4
      Left            =   16560
      TabIndex        =   70
      Text            =   "4"
      Top             =   3450
      Visible         =   0   'False
      Width           =   390
   End
   Begin VB.TextBox txt 
      Height          =   345
      Index           =   2
      Left            =   16530
      TabIndex        =   77
      Text            =   "2"
      Top             =   4320
      Visible         =   0   'False
      Width           =   465
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
      Left            =   16530
      TabIndex        =   76
      Text            =   "1"
      Top             =   3030
      Visible         =   0   'False
      Width           =   555
   End
   Begin MSMask.MaskEdBox mskDt 
      Height          =   360
      Index           =   2
      Left            =   1560
      TabIndex        =   12
      Top             =   960
      Width           =   1695
      _ExtentX        =   2990
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
   Begin VB.CommandButton CmdBagDetails 
      Caption         =   "Cartoon &Bag Details"
      Height          =   375
      Left            =   465
      TabIndex        =   48
      Top             =   4530
      Width           =   1695
   End
   Begin MSMask.MaskEdBox mskDt 
      Height          =   375
      Index           =   43
      Left            =   3630
      TabIndex        =   21
      Top             =   3360
      Width           =   1320
      _ExtentX        =   2328
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
      Index           =   0
      Left            =   12000
      TabIndex        =   41
      Top             =   7980
      Visible         =   0   'False
      Width           =   1695
      _ExtentX        =   2990
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
   Begin MSFlexGridLib.MSFlexGrid msGrid 
      Height          =   4425
      Left            =   7500
      TabIndex        =   33
      Top             =   3660
      Width           =   3105
      _ExtentX        =   5477
      _ExtentY        =   7805
      _Version        =   393216
      FixedCols       =   0
      RowHeightMin    =   275
      BackColorFixed  =   -2147483647
      ForeColorFixed  =   -2147483634
      BackColorBkg    =   12632256
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
   Begin VB.Label LblTicket 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Ticket"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   330
      Left            =   5760
      TabIndex        =   125
      Top             =   7740
      Width           =   1455
   End
   Begin VB.Label LblHsn 
      BackColor       =   &H00FFFFFF&
      Caption         =   "HSNCODE"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   330
      Left            =   5040
      TabIndex        =   119
      Top             =   3375
      Width           =   1455
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "GSTIN :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   450
      Left            =   960
      TabIndex        =   118
      Top             =   7680
      Width           =   1260
   End
   Begin VB.Label lblGST 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Party's GST"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   450
      Left            =   2280
      TabIndex        =   117
      Top             =   7680
      Width           =   3375
   End
   Begin VB.Label lbl1 
      BackStyle       =   0  'Transparent
      Caption         =   "Goods Desp.Dt. :"
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
      Index           =   16
      Left            =   10320
      TabIndex        =   114
      Top             =   7980
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.Label lbl1 
      BackStyle       =   0  'Transparent
      Caption         =   "Not&e                 :"
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
      Index           =   10
      Left            =   480
      TabIndex        =   53
      Top             =   6600
      Width           =   1695
   End
   Begin VB.Image Image2 
      Height          =   1095
      Left            =   15900
      Stretch         =   -1  'True
      Top             =   7440
      Width           =   1440
   End
   Begin VB.Label Label14 
      BackStyle       =   0  'Transparent
      Caption         =   "Booking No.     :"
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
      Height          =   330
      Left            =   480
      TabIndex        =   45
      Top             =   3360
      Width           =   1575
   End
   Begin VB.Label lblLRAmt 
      BackStyle       =   0  'Transparent
      Caption         =   "LR Amount  :"
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
      Height          =   345
      Left            =   4125
      TabIndex        =   91
      Top             =   5010
      Width           =   1650
   End
   Begin VB.Label lblLRDate 
      BackStyle       =   0  'Transparent
      Caption         =   "LR Date       :"
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
      Height          =   345
      Left            =   4125
      TabIndex        =   55
      Top             =   4605
      Width           =   1650
   End
   Begin VB.Label lblLRNo 
      BackStyle       =   0  'Transparent
      Caption         =   "Challan No   :"
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
      Height          =   345
      Left            =   4125
      TabIndex        =   54
      Top             =   4185
      Width           =   1650
   End
   Begin VB.Label lbl1 
      BackStyle       =   0  'Transparent
      Caption         =   "Rate Per            :"
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
      Index           =   15
      Left            =   480
      TabIndex        =   52
      Top             =   6210
      Width           =   1695
   End
   Begin VB.Label lbl1 
      BackStyle       =   0  'Transparent
      Caption         =   "Rate                  :"
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
      Index           =   9
      Left            =   480
      TabIndex        =   51
      Top             =   5805
      Width           =   1695
   End
   Begin VB.Label lbl1 
      BackStyle       =   0  'Transparent
      Caption         =   "Godo&wn           :"
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
      Index           =   14
      Left            =   480
      TabIndex        =   46
      Top             =   3765
      Width           =   1695
   End
   Begin VB.Shape Shape4 
      BorderColor     =   &H00C00000&
      BorderWidth     =   5
      FillColor       =   &H00C00000&
      Height          =   8385
      Left            =   15
      Top             =   -90
      Width           =   10800
   End
   Begin VB.Shape Shape3 
      BorderColor     =   &H000000C0&
      FillColor       =   &H000000C0&
      FillStyle       =   0  'Solid
      Height          =   0
      Left            =   0
      Top             =   7515
      Width           =   9375
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
      Left            =   6435
      TabIndex        =   81
      Top             =   0
      Width           =   2475
   End
   Begin VB.Label lbl1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "G P &No.  :"
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
      Left            =   480
      TabIndex        =   0
      Top             =   540
      Width           =   900
   End
   Begin VB.Label lbl1 
      BackStyle       =   0  'Transparent
      Caption         =   "P&urch/Inw No.  :"
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
      Height          =   405
      Index           =   7
      Left            =   480
      TabIndex        =   47
      Top             =   4185
      Width           =   1605
   End
   Begin VB.Label lbl1 
      BackStyle       =   0  'Transparent
      Caption         =   "Bags                 :"
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
      Index           =   8
      Left            =   480
      TabIndex        =   49
      Top             =   5010
      Width           =   1695
   End
   Begin VB.Label lbl1 
      BackStyle       =   0  'Transparent
      Caption         =   "M&ill         :"
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
      Index           =   5
      Left            =   480
      TabIndex        =   35
      Top             =   2415
      Width           =   1695
   End
   Begin VB.Label lbl1 
      BackStyle       =   0  'Transparent
      Caption         =   "C&ount Name     :"
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
      Index           =   6
      Left            =   480
      TabIndex        =   36
      Top             =   2955
      Width           =   1695
   End
   Begin VB.Label lbl1 
      BackStyle       =   0  'Transparent
      Caption         =   "&Broker    :"
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
      Left            =   480
      TabIndex        =   34
      Top             =   1995
      Width           =   1695
   End
   Begin VB.Label lbl1 
      BackStyle       =   0  'Transparent
      Caption         =   "Pa&rty      :"
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
      Left            =   480
      TabIndex        =   15
      Top             =   1575
      Width           =   1695
   End
   Begin VB.Label lbl1 
      BackStyle       =   0  'Transparent
      Caption         =   "Sa&les Type   : "
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
      Index           =   2
      Left            =   3510
      TabIndex        =   13
      Top             =   810
      Width           =   1890
   End
   Begin VB.Label lbl1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Date       :"
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
      Left            =   480
      TabIndex        =   11
      Top             =   960
      Width           =   915
   End
   Begin VB.Shape Shape1 
      FillColor       =   &H00BDDFD1&
      FillStyle       =   0  'Solid
      Height          =   3240
      Left            =   7560
      Shape           =   4  'Rounded Rectangle
      Top             =   360
      Width           =   2955
   End
   Begin VB.Label Label6 
      BackColor       =   &H00C00000&
      Caption         =   "    Gate Pass"
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
      Height          =   360
      Left            =   -15
      TabIndex        =   80
      Top             =   15
      Width           =   10800
   End
   Begin VB.Label lbl1 
      BackStyle       =   0  'Transparent
      Caption         =   "Weig&ht             :"
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
      Index           =   12
      Left            =   480
      TabIndex        =   50
      Top             =   5415
      Width           =   1680
   End
   Begin VB.Label lbl1 
      BackStyle       =   0  'Transparent
      Caption         =   "Lot No.        :"
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
      Height          =   345
      Index           =   11
      Left            =   4125
      TabIndex        =   56
      Top             =   5415
      Width           =   1650
   End
   Begin VB.Label lbl1 
      BackStyle       =   0  'Transparent
      Caption         =   "Due Days    :"
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
      Index           =   13
      Left            =   4125
      TabIndex        =   57
      Top             =   5805
      Width           =   1650
   End
   Begin VB.Shape Shape5 
      BorderColor     =   &H00000080&
      FillColor       =   &H00EDFDFE&
      FillStyle       =   0  'Solid
      Height          =   975
      Left            =   240
      Top             =   435
      Width           =   7065
   End
   Begin VB.Shape Shape2 
      BackColor       =   &H00808080&
      BorderColor     =   &H004D177D&
      FillColor       =   &H00ECFFFF&
      FillStyle       =   0  'Solid
      Height          =   1380
      Left            =   240
      Top             =   1455
      Width           =   7065
   End
   Begin VB.Shape Shape6 
      BorderColor     =   &H00000080&
      FillColor       =   &H00EDFDFE&
      FillStyle       =   0  'Solid
      Height          =   4545
      Left            =   240
      Top             =   2865
      Width           =   7065
   End
End
Attribute VB_Name = "frmGatepassNew"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Const tblName As String = "tblVouSub"
Const IdField As String = "vno"
Const MaxNo As Long = 45
Const OrderField As String = "Vdt,Vno"
Dim rstbl As Recordset, ShowFrmFlag As Boolean
Dim RsP As Recordset, RsTmp As Recordset
Dim rstblSub As Recordset '-- TblBags
Dim varBookMark As Variant
Dim FirstTimeFlag As Boolean
Dim FormAction As Integer
Dim rsAddLess As Recordset '-- TblAddLess
Dim oldQty As Long
Dim WA_OVno As String
Dim WA_OParty As String
Dim WA_OBrok As String
Dim WA_OMill As String
Dim WA_OCount As String
Dim WA_OBkNo As Long
Dim WA_OBag As Integer
Dim WA_OWt As Double
Dim WA_ORate As Double
Dim WA_ORtPer As Double
Dim WA_ODdays As Integer

Dim WA_NVno As String
Dim WA_NParty As String
Dim WA_NBrok As String
Dim WA_NMill As String
Dim WA_NCount As String
Dim WA_NBkNo As Long
Dim WA_NBag As Integer
Dim WA_NWt As Double
Dim WA_NRate As Double
Dim WA_NRtPer As Double
Dim WA_NDdays As Integer



Dim i As Long
Private Sub cbo_LostFocus(Index As Integer)
SetControlEd
End Sub
Private Sub cmdBtn_Click(Index As Integer)
'On Error GoTo ErrorRoutine
Dim FindQStr As String
Dim X As String
Dim Hank As String
Select Case Index
Case 0 'Previous
    rstbl.MovePrevious
    If rstbl.BOF = True Then
        rstbl.MoveFirst
        MsgBox ("This is First Record"), vbInformation + vbOKOnly, "First Record"
    End If
    ReadFields
    If cbo(23).ListIndex = 3 Then '-- SIT
        lbl1(16).Visible = True
        mskDt(0).Visible = True '-goods desp Dt
        Chk(1).Visible = True
        Chk(0).Visible = True
    Else
        lbl1(16).Visible = False
        mskDt(0).Visible = False '-goods desp Dt
        Chk(1).Visible = True
        Chk(0).Visible = True
    End If
Case 1 'Next
    rstbl.MoveNext
    If rstbl.EOF = True Then
        rstbl.MoveLast
        MsgBox ("This is Last Record"), vbInformation + vbOKOnly, "Last Record"
    End If
    ReadFields
    If cbo(23).ListIndex = 3 Then '-- SIT
        lbl1(16).Visible = True
        mskDt(0).Visible = True '-goods desp Dt
        Chk(1).Visible = True
        Chk(0).Visible = True
    Else
        lbl1(16).Visible = False
        mskDt(0).Visible = False '-goods desp Dt
        Chk(1).Visible = True
        Chk(0).Visible = True
    End If
Case 2 'First
    rstbl.MoveFirst
    ReadFields
    If cbo(23).ListIndex = 3 Then '-- SIT
        lbl1(16).Visible = True
        mskDt(0).Visible = True '-goods desp Dt
        Chk(1).Visible = True
        Chk(0).Visible = True
    Else
        lbl1(16).Visible = False
        mskDt(0).Visible = False '-goods desp Dt
        Chk(1).Visible = True
        Chk(0).Visible = True
    End If
Case 3 'Last
    rstbl.MoveLast
    ReadFields
    If cbo(23).ListIndex = 3 Then '-- SIT
        lbl1(16).Visible = True
        mskDt(0).Visible = True '-goods desp Dt
        Chk(1).Visible = True
        Chk(0).Visible = True
    Else
        lbl1(16).Visible = False
        mskDt(0).Visible = False '-goods desp Dt
        Chk(1).Visible = True
        Chk(0).Visible = True
    End If
Case 4 'Find
    If gBackEndDB = gBackEndAccess Then
        X = "(select iif(VSubGPSLType=0,'Trade',iif(VSubGPSLType=1,'Cosignment',iif(VSubGPSLType=2,'Depot','SIT') )) from tblVouSub where" _
        & " tblVouSub.Vno=tblAddLess.Vno and tblVouSub.VType=tblAddLess.VType and tblVouSub.VYear=tblAddLess.VYear)"
    Else '-- Oracle
        X = "(select decode(VSubGPSLType,0,'Trade',1,'Cosignment',2,'Depot','SIT') from tblVouSub where" _
        & " tblVouSub.Vno=tblAddLess.Vno and tblVouSub.VType=tblAddLess.VType and tblVouSub.VYear=tblAddLess.VYear)"
    End If
    FindQStr = "select tblAddLess.vno,AcName,tblAddLess.Vdt," & X & " from " _
    & " tblAddLess,TblMastAccount where tblAddLess.AdAcCrCode=TblMastAccount.AcCode and " _
    & "  tblAddLess.vtype='GP' and tblAddLess.VYear=" & gCYear
    gClsSearch.SearchMultiField FindQStr, "", Array("Vno", "Party", "Vdt", "Sales Type"), Array(800, 4000, 1080, 2000), "", " ", " tblAddLess.vno,tblAddLess.vdt ", 0, 0, False, , True
    If gClsSearch.SearchMultiRetCol(0) <> "" Then
        rstbl.MoveFirst
        rstbl.Find "Vno = " & gClsSearch.SearchMultiRetCol(0)
        gClsSearch.SearchMultiRetCol(0) = ""
        ReadFields
        If cbo(23).ListIndex = 3 Then '-- SIT
            lbl1(16).Visible = True
            mskDt(0).Visible = True '-goods desp Dt
            Chk(1).Visible = True
        Else
            lbl1(16).Visible = False
            mskDt(0).Visible = False '-goods desp Dt
            Chk(1).Visible = True
        End If
    End If
Case 5 'List
    'GenMastList ("LAC")
Case 6 'Add
    FormAction = vbDataActionAddNew
    txt(28) = "A"
    Call GProcActivateControls(Me, True)
    Call GProcSetButtons(Me, vbDataActionAddNew)
    Call GProcClearForm(Me, rstbl, MaxNo, True)
    SetControlEd
    If cbo(23).ListIndex < 0 Then cbo(23).ListIndex = 0
    txt(0).text = GProcGenerateId(tblName, "vno", "vtype='GP' and VYear=" & gCYear)
    mskDt(2) = GProcGenerateVDt("tblVouSub", "Vdt", "Vtype='GP' and VYear=" & gCYear)
    If FirstTimeFlag = True Then
        FirstTimeFlag = False
    Else
        txt(0).SetFocus
    End If
    SetGrid
Case 7 'Modify
    WA_OVno = txt(0)
    WA_OParty = txttempAL(3)
    WA_OBrok = txttempAL(4)
    WA_OMill = txttempAL(5)
    WA_OCount = txttemp(4)
    WA_OBkNo = txt(42)
    WA_OBag = txt(5)
    WA_OWt = txt(6)
    WA_ORate = txt(7)
    WA_ORtPer = txt(8)
    WA_ODdays = txtAL(6)
    
    FormAction = vbDataActionUpdate
    Hank = ""
    'Call GProcChkPass("GP", FormAction, txt(0), txt(2), txttempAL(3), gUserName, Hank)
    'If gChkPassTrueFalse = True Then
        'If mURecModi = True Then
        '    SearchHelp
        '    DoEvents
    '    If IsRelatedRecord = False Then
            FormAction = vbDataActionUpdate
            txt(28) = "M"
            Call GProcActivateControls(Me, True)
            Call GProcSetButtons(Me, vbDataActionUpdate)
            SetControlEd
             '--Update In Stock
            GProcUpdateItCtrlStock txt(12), CLng(txt(5)), CDbl(txt(6)), True
             '--Update Booking Stock
'            GProcUpdateSetupBkItCtrlStock txt(44), CLng(txt(5)), CDbl(txt(6)), True
        If Val(txt(42)) > 0 Then
        If (GProcRstOpen(RsP, "select * from tblVousub_book where vtype = 'OP' and vno = " & txt(42) & " and vdt = Cdate('" & txt(43) & "')   ", "O") > 0) Then
            RsP.MoveFirst
            Do While Not RsP.EOF '-issue bag,Wt
               If (GProcRstOpen(RsTmp, "select sum(VSUBBAG),sum(VSUBWt) from tblVousub where vtype = 'GP'  and  VSubBookNo = " & RsP.Fields("VNO") & " and  VSubBookDt = Cdate('" & RsP.Fields("VDt") & "') and VSubBkItCtrlNo = " & RsP.Fields("VSubItCtrlNo"), "O") > 0) Then
                   RsP.Fields("vsubissbag") = IIf(IsNumeric(RsTmp.Fields(0)), RsTmp.Fields(0), 0)
                   RsP.Fields("vsubisswt") = IIf(IsNumeric(RsTmp.Fields(1)), RsTmp.Fields(1), 0)
                   RsP.Update
               End If
               RsP.MoveNext
            Loop
        End If
        End If
            
            
            txt(0).Enabled = False
            mskDt(2).SetFocus
            msGrid.Enabled = True '-Grid
            
            '--- if Ralated Invoice
    '        If IsRelatedRecord = True Then
    '            Call GProcActivateControls(Me, False)
                 txttempAL(3).Enabled = True '-Party
                 txttempAL(4).Enabled = True '-Broker
    '             Chk(1).Enabled = True
    '        End If
       ' End If
    'Else
    '   FormAction = vbDataActionClose
    'End If
Case 8 'Delete
    FormAction = vbDataActionDelete
    Hank = ""
    'Call GProcChkPass("GP", FormAction, txt(0), txt(2), txttempAL(3), gUserName, Hank)
    'If gChkPassTrueFalse = True Then
        ' If mURecDel = True Then
             If IsRelatedRecord = True Then
                 MsgBox "Related Invoice is exists.You can not Modify/Delete this Record.", vbCritical + vbOKOnly, App.Title
                 Exit Sub
             End If
             If IsRelatedRecord = False Then
                 If MsgBox("Do You Really Delete This Record?", vbQuestion + vbYesNo + vbDefaultButton1, "Confirm Delete") = vbYes Then
                     If rstbl.RecordCount > 1 Then
                         '-- Delete Related Record
                         rstbl.Delete
                        ' X = "select BagNo,Sno from tblBags where InwType='' and InwNo=0 and GPNo=" & txt(0) & " and GpType='GP' order by BagNo"
                         '--- Bag Details
'                         gCn.Execute "delete From tblBags where InwType is null and InwNo=0 and GPNo=" & txt(0) & " and GpType='GP'"
                         gCn.Execute "delete from tblGpSub where vtype='GP' and vno=" & txt(0) & " and " _
                         & " VYear=" & gCYear
                         
                         '--- Cartoon Bag Details
'                         gCn.Execute "update tblBags set GPNo=0 ,GpType=Null where InwNo>0 and  GpNo=" & txt(0) & " and GPType='GP'"
                         '--Update In Stock
                         GProcUpdateItCtrlStock txt(12), CLng(txt(5)), CDbl(txt(6)), False, True
                          '--Update Booking Stock
                         GProcUpdateSetupBkItCtrlStock txt(44), CLng(txt(5)), CDbl(txt(6)), False, True
                         DeleteRelatedRecord
                         cmdBtn_Click (1)
                     ElseIf rstbl.RecordCount = 1 Then
                         '-- Delete Related Record
                         rstbl.Delete
                         '--- Bag Details
                         gCn.Execute "delete from tblGpSub where vtype='GP' and vno=" & txt(0) & " and " _
                         & " VYear=" & gCYear
'                         gCn.Execute "delete From tblBags where InwType is null and InwNo=0 and GPNo=" & txt(0) & " and GpType='GP'"
                         '--- Cartoon Bag Details
'                         gCn.Execute "update tblBags set GPNo=0 ,GpType=Null where InwNo>0 and  GpNo=" & txt(0) & " and GPType='GP'"
                         '--Update In Stock
                         GProcUpdateItCtrlStock txt(12), CLng(txt(5)), CDbl(txt(6)), False, True
                         '--Update Booking Stock
                         GProcUpdateSetupBkItCtrlStock txt(44), CLng(txt(5)), CDbl(txt(6)), False, True
                         DeleteRelatedRecord
                         Call GProcActivateControls(Me, False)
                         Call GProcSetButtons(Me, vbDataActionCancel)
                         SetControlEd
                         Form_Load
                     End If
                 End If
             End If
     ' End If
Case 9 'Print
    Frame1.Visible = True
    DoEvents
    txtno(0).Enabled = True
    txtno(1).Enabled = True
    txtno(0) = txt(0)
    txtno(1) = txt(0)
    OptOrg.Enabled = True
    OptDup.Enabled = True
    OptOrg.Value = True
    txtno(0).SetFocus
    Check1.Enabled = True
    ChkRt.Enabled = True
    ChkRt.Value = 1
Case 10 'Exit
    Unload Me
Case 13 'Save
    If ValidateData = True Then
    If FormAction = vbDataActionUpdate Then
       '-- For WhatsApp
        WA_NVno = txt(0)
        WA_NParty = txttempAL(3)
        WA_NBrok = txttempAL(4)
        WA_NMill = txttempAL(5)
        WA_NCount = txttemp(4)
        WA_NBkNo = txt(42)
        WA_NBag = txt(5)
        WA_NWt = txt(6)
        WA_NRate = txt(7)
        WA_NRtPer = txt(8)
        WA_NDdays = txtAL(6)
        If WA_NParty <> WA_OParty Or WA_NBrok <> WA_OBrok Or WA_NMill <> WA_OMill Or WA_NCount <> WA_OCount Or WA_NBkNo <> WA_OBkNo Or WA_NBag <> WA_OBag Or WA_NWt <> WA_OWt Or WA_NRtPer <> WA_ORtPer Or WA_NDdays <> WA_ODdays Or WA_NRate <> WA_ORate Then
           GProcCrystalRptPreparation CryReport, "Window"
           CryReport.WindowTitle = "Account List"
           CryReport.Formulas(4) = "wGpNo = '" & WA_NVno & "'"
           CryReport.Formulas(5) = "wParty = '" & WA_NParty & "'"
           CryReport.Formulas(6) = "wBrok = '" & WA_NBrok & "'"
           CryReport.Formulas(7) = "wMill = '" & WA_NMill & "'"
           CryReport.Formulas(8) = "wCount = '" & WA_NCount & "'"
           CryReport.Formulas(9) = "wBkNo = " & WA_NBkNo & ""
           CryReport.Formulas(10) = "wBags = " & WA_NBag & ""
           CryReport.Formulas(11) = "wWeight = " & WA_NWt & ""
           CryReport.Formulas(12) = "wRate = " & WA_NRate & ""
           CryReport.Formulas(13) = "wRatePer = " & WA_NRtPer & ""
           CryReport.Formulas(14) = "wDueDays = " & WA_NDdays & ""
           
           CryReport.Formulas(15) = "oParty = '" & WA_OParty & "'"
           CryReport.Formulas(16) = "oBrok = '" & WA_OBrok & "'"
           CryReport.Formulas(17) = "oMill = '" & WA_OMill & "'"
           CryReport.Formulas(18) = "oCount = '" & WA_OCount & "'"
           CryReport.Formulas(19) = "oBkNo = " & WA_OBkNo & ""
           CryReport.Formulas(20) = "oBags = " & WA_OBag & ""
           CryReport.Formulas(21) = "oWeight = " & WA_OWt & ""
           CryReport.Formulas(22) = "oRate = " & WA_ORate & ""
           CryReport.Formulas(23) = "oRatePer = " & WA_ORtPer & ""
           CryReport.Formulas(24) = "oDueDays = " & WA_ODdays & ""
           
           CryReport.Formulas(25) = "WGpDt = '" & txt(2) & "'"
           CryReport.Formulas(26) = "WModiBy = '" & gUserName & "'"
           CryReport.Formulas(27) = "WModidtTm = '" & Format(Now, "dd/mm/yyyy     hh:mm:ss") & "'"
           
           Dim objCrystal As CRAXDRT.Application
           Dim objReport As CRAXDRT.Report
           Dim Tattach As String
           Dim Mfile As String
                    
           Set objCrystal = New CRAXDRT.Application
           Tattach = gReportPath & "RptGpChange.rpt"
           Set objReport = objCrystal.OpenReport(Tattach, 1)
                    
           Dim crTab As CRAXDRT.DatabaseTable
           For Each crTab In objReport.Database.Tables
               crTab.Location = gDbLocation
           Next
                    
           objReport.RecordSelectionFormula = ""
                    
                    
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
               Case "{@wGpNo}"
                     objReport.FormulaFields(j).text = "Trim(" & Chr(39) & WA_NVno & Chr(39) & ")"
               Case "{@wParty}"
                    objReport.FormulaFields(j).text = "Trim(" & Chr(39) & WA_NParty & Chr(39) & ")"
               Case "{@wBrok}"
                   objReport.FormulaFields(j).text = "Trim(" & Chr(39) & WA_NBrok & Chr(39) & ")"
               Case "{@wMill}"
                    objReport.FormulaFields(j).text = "Trim(" & Chr(39) & WA_NMill & Chr(39) & ")"
               Case "{@wCount}"
                    objReport.FormulaFields(j).text = "Trim(" & Chr(39) & WA_NCount & Chr(39) & ")"
               Case "{@wBkNo}"
                   objReport.FormulaFields(j).text = "" & WA_NBkNo & ""
               Case "{@wBags}"
                   objReport.FormulaFields(j).text = "" & WA_NBag & ""
               Case "{@wWeight}"
                   objReport.FormulaFields(j).text = "" & WA_NWt & ""
               Case "{@wRate}"
                   objReport.FormulaFields(j).text = "" & WA_NRate & ""
               Case "{@wRatePer}"
                   objReport.FormulaFields(j).text = "" & WA_NRtPer & ""
               Case "{@wDueDays}"
                   objReport.FormulaFields(j).text = "" & WA_NDdays & ""
               
               Case "{@oParty}"
                    objReport.FormulaFields(j).text = "Trim(" & Chr(39) & WA_OParty & Chr(39) & ")"
               Case "{@oBrok}"
                   objReport.FormulaFields(j).text = "Trim(" & Chr(39) & WA_OBrok & Chr(39) & ")"
               Case "{@oMill}"
                    objReport.FormulaFields(j).text = "Trim(" & Chr(39) & WA_OMill & Chr(39) & ")"
               Case "{@oCount}"
                    objReport.FormulaFields(j).text = "Trim(" & Chr(39) & WA_OCount & Chr(39) & ")"
               Case "{@oBkNo}"
                   objReport.FormulaFields(j).text = "" & WA_OBkNo & ""
               Case "{@oBags}"
                   objReport.FormulaFields(j).text = "" & WA_OBag & ""
               Case "{@oWeight}"
                   objReport.FormulaFields(j).text = "" & WA_OWt & ""
               Case "{@oRate}"
                   objReport.FormulaFields(j).text = "" & WA_ORate & ""
               Case "{@oRatePer}"
                   objReport.FormulaFields(j).text = "" & WA_ORtPer & ""
               Case "{@oDueDays}"
                   objReport.FormulaFields(j).text = "" & WA_ODdays & ""
               
               Case "{@WGpDt}"
                    objReport.FormulaFields(j).text = "Trim(" & Chr(39) & txt(2) & Chr(39) & ")"
               Case "{@WModiBy}"
                    objReport.FormulaFields(j).text = "Trim(" & Chr(39) & gUserName & Chr(39) & ")"
               Case "{@WModidtTm}"
                    objReport.FormulaFields(j).text = "Trim(" & Chr(39) & Format(Now, "dd/mm/yyyy     hh:mm:ss") & Chr(39) & ")"
               
               End Select
                   
           Next j
           Dim PtyMbl As String
           Dim FileNMWithPath As String
           Dim FileNM As String
           Dim aFile As String
           FileNM = "" & Format(Now, "ddMMyyyyhhmmss") & "" & ".pdf"
           FileNMWithPath = gReportPath & "" & FileNM & ""
           PtyMbl = "9665748022"
                     
           ExportReportToPDF objReport, FileNMWithPath, "foo"
           SentWhatsApp "PDF", FileNM, "", PtyMbl, ""

           gReportPath = App.Path & "\HIReports_Access\"
           aFile = gReportPath & "*.pdf"
           If Len(Dir$(aFile)) > 0 Then
              Kill aFile
           End If
        End If
        
       '---End
       End If
        
        FillTxtFromTemp
        If cbo(23).ListIndex = 4 Then '-- Hank Sales
           txt(45) = "HNK"
        Else
           txt(45) = ""
        End If
        SaveRelatedRecord '-- Add Less
        GProcSaveRecord Me, rstbl, FormAction, MaxNo
        '--Update In Stock
        GProcUpdateItCtrlStock txt(12), CLng(txt(5)), CDbl(txt(6))
        '--Update Booking Stock
'        GProcUpdateSetupBkItCtrlStock txt(44), CLng(txt(5)), CDbl(txt(6))
        If Val(txt(42)) > 0 Then
        
        If (GProcRstOpen(RsP, "select * from tblVousub_book where vtype = 'OP' and vno = " & txt(42) & " and vdt = Cdate('" & txt(43) & "')   ", "O") > 0) Then
            RsP.MoveFirst
            Do While Not RsP.EOF '-issue bag,Wt
               If (GProcRstOpen(RsTmp, "select sum(VSUBBAG),sum(VSUBWt) from tblVousub where vtype = 'GP'  and  VSubBookNo = " & RsP.Fields("VNO") & " and  VSubBookDt = Cdate('" & RsP.Fields("VDt") & "') and VSubBkItCtrlNo = " & RsP.Fields("VSubItCtrlNo"), "O") > 0) Then
                   RsP.Fields("vsubissbag") = IIf(IsNumeric(RsTmp.Fields(0)), RsTmp.Fields(0), 0)
                   RsP.Fields("vsubisswt") = IIf(IsNumeric(RsTmp.Fields(1)), RsTmp.Fields(1), 0)
                   RsP.Update
               End If
               RsP.MoveNext
            Loop
        End If
        End If
        Call GProcActivateControls(Me, False)
        Call GProcSetButtons(Me, True)
        If FormAction = vbDataActionAddNew Then
            If cmdBtn(6).Visible = True Then cmdBtn(6).SetFocus
        Else
            If cmdBtn(7).Visible = True Then cmdBtn(7).SetFocus
        End If
        FormAction = vbDataActionClose
        '-- For Whatsapp SMS Text
        Dim Pside As String
        Dim GpAmt As Double
        Dim GpRate As Double
        
        GpAmt = 0
        Dim Due_dt As Date
        Dim ItTicket As String
        Dim MillCode As Integer
        Dim ItType As Integer
        Dim GstRate As Double
        Dim ExGstRate As Double
        Dim ItTicketPrint As Integer
        Dim Crt As Double
        Dim Srt As Double
        Dim Irt As Double
        Dim BasicAmt As Double
        Dim GSTAmt As Double
        Dim SmsString As String
        Dim SmsMblNo As String
        Dim SmsMblNo2 As String
            gCn.BeginTrans
            gCn.Execute "Delete from tmpWhatsAppErr where username = '" & gUserName & "'  "
            gCn.CommitTrans
            
            
            SmsString = "Firm   :" & vbLf
            SmsString = SmsString + "*" + gCName + "*" & vbLf & vbLf
            SmsString = SmsString + "*GATE PASS*" & vbLf & vbLf
            SmsString = SmsString + "No     : *" + Trim(txt(0)) + "*" & vbLf
            SmsString = SmsString + "Date  : " + Trim(CStr(Format(mskDt(2), "dd/mm/yy"))) & vbLf & vbLf
            
            SmsString = SmsString + "Party  : *" + txttempAL(3) + "*" & vbLf & vbLf
            
            SmsString = SmsString + "Broker : " + txttempAL(4) & vbLf
            MillCode = GProcGetColumnValue("tblMastAccount", "AcName", txttempAL(5), "S", "AcCode", "N")
            ItTicket = GProcGetColumnValue("tblMastItem", "ItName", txttemp(4), "S", "ItTicket", "S", " ItMillcode=" & MillCode)
            ItType = GProcGetColumnValue("tblMastItem", "ItName", txttemp(4), "S", "ItType", "N", " ItMillcode=" & MillCode)
            ItTicketPrint = GProcGetColumnValue("tblMastItem", "ItName", txttemp(4), "S", "ItTicketPrint", "N", " ItMillcode=" & MillCode)
'SmsMblNo = GProcGetColumnValue("tblMastAccount", "AcName", txttempAL(3), "S", "AcFaxNo", "S")
            If Chk(0).Value = 1 Or ItTicketPrint = 1 Then
               SmsString = SmsString + "Mill   : " + ItTicket & vbLf & vbLf
            ElseIf Chk(1).Value = 1 Then
               SmsString = SmsString + "Mill   :  Export" & vbLf & vbLf
            Else
               SmsString = SmsString + "Mill   : " + txttempAL(5) & vbLf & vbLf
            End If
            SmsString = SmsString + "Count  : " & vbLf
            SmsString = SmsString + "*" + Trim(txttemp(4)) + "*" & vbLf & vbLf
            
            SmsString = SmsString + "Qty.      : " + Format(txt(5), "#####0") & vbLf
            SmsString = SmsString + "Weight : " + Format(txt(6), "#####0.000") + " Kgs " & vbLf
            
            If ItType = 0 Then
              Crt = GProcGetColumnValue("tblMastNarration", "Narration", "G S T", "S", "CotCGSTRt", "N")
              Srt = GProcGetColumnValue("tblMastNarration", "Narration", "G S T", "S", "CotSGSTRt", "N")
              Irt = GProcGetColumnValue("tblMastNarration", "Narration", "G S T", "S", "CotIGSTRt", "N")
            Else
              Crt = GProcGetColumnValue("tblMastNarration", "Narration", "G S T", "S", "PolCGSTRt", "N")
              Srt = GProcGetColumnValue("tblMastNarration", "Narration", "G S T", "S", "PolSGSTRt", "N")
              Irt = GProcGetColumnValue("tblMastNarration", "Narration", "G S T", "S", "PolIGSTRt", "N")
            End If
            If Val(txt(26)) = 1 Then '---Rate Excluding GST
               ExGstRate = Val(txt(7))
               BasicAmt = Round(Val(txt(6)) * Val(txt(7)) / Val(txt(8)), 2)
               GSTAmt = Round((BasicAmt * Crt) / 100, 2) + Round((BasicAmt * Srt) / 100, 2)
               GpAmt = Round(BasicAmt + GSTAmt, 0)
               GstRate = (GpAmt / Val(txt(6))) * Val(txt(8))
            Else
               GstRate = Val(txt(7))
               GpAmt = Round(Val(txt(6)) * Val(txt(7)) / Val(txt(8)), 0)
               BasicAmt = (GpAmt * 100) / (100 + Crt + Srt)
               ExGstRate = (BasicAmt / Val(txt(6))) * Val(txt(8))
            End If
            
            SmsString = SmsString + "Without GST Rate : *" + Format(ExGstRate, "#####0.000") + "*" + " / " + Format(txt(8), "#####0.0") + " Kgs " & vbLf
            
            SmsString = SmsString + "With GST Rate    : *" + Format(GstRate, "#####0.000") + "*" + " / " + Format(txt(8), "#####0.0") + " Kgs " & vbLf
            SmsString = SmsString + "Amount   : " + Format(GpAmt, "#####0.000") & vbLf
            If Val(txtAL(6)) = 30 Then
               Pside = "Regular"
            ElseIf Val(txtAL(6)) = 7 Then
               Pside = "Second Day"
            ElseIf Val(txtAL(6)) = 1 Then
               Pside = "Next Day"
            End If
            
            
            
            SmsString = SmsString + "Payment Condition : " + Pside & vbLf
            Due_dt = DateAdd("d", txtAL(6), mskDt(2))
            SmsString = SmsString + "*Due Date  : " + Trim(CStr(Format(Due_dt, "dd/mm/yy"))) + "*" & vbLf
            SmsMblNo = GProcGetColumnValue("tblMastAccount", "AcName", txttempAL(3), "S", "AcFaxNo", "S")
            If SmsMblNo = "" Then
               MsgBox ("Party Mobile No Not Found....")
            Else
              SentWhatsApp "Txt", "", SmsString, SmsMblNo, txttempAL(3)
            End If
            SmsMblNo2 = GProcGetColumnValue("tblMastAccount", "AcName", txttempAL(4), "S", "AcFaxNo", "S")
            If SmsMblNo2 = "" Then
               MsgBox ("Broker Mobile No Not Found....")
            Else
              SentWhatsApp "Txt", "", SmsString, SmsMblNo2, txttempAL(4)
            End If

            X = "Select * from tmpWhatsAppErr where username = '" & gUserName & "'"
            If GProcRstOpen(RsTmp, X, "R", gCn) > 0 Then
               GProcWhatsAppError
            End If
        '-- End For Whatsapp SMS Text

    
    End If
Case 14 'Cancel
    txtGrid.Visible = False
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
    If msGrid.Rows > 1 Then
        msGrid.Col = 0
        msGrid.Row = 1
        txtGrid.Move msGrid.CellLeft + msGrid.Left, msGrid.CellTop + msGrid.Top, msGrid.CellWidth, msGrid.CellHeight
    End If
    If cbo(23).ListIndex = 3 Then '-- SIT
        lbl1(16).Visible = True
        mskDt(0).Visible = True '-goods desp Dt
        Chk(1).Visible = True
       Chk(0).Visible = True
    Else
        lbl1(16).Visible = False
        mskDt(0).Visible = False '-goods desp Dt
        Chk(1).Visible = True
        Chk(0).Visible = True
    End If
End Select
Exit Sub
ErrorRoutine:
    rstbl.CancelUpdate
    GProcErrorHandler
End Sub
Private Sub Cbo_KeyPress(Index As Integer, KeyAscii As Integer)
If KeyAscii = vbKeyReturn Then Exit Sub
KeyAscii = 0
End Sub


Private Sub CmdGPDelivery_Click()
frmGPDeliveryEntryNew.Show
End Sub
'---Show  Ledger
Private Sub cmdLedger_Click()
'On Error GoTo ErrorHandler
frmRptAccount.txt(1) = txttempAL(3)
frmMain.mnuRptRSubAccLeg_Click (0)
Exit Sub
ErrorHandler:
    GProcErrorHandler
End Sub
'--- Outstanding Report
Private Sub CmdOutstandingRpt_Click()
'On Error GoTo ErrorHandler
frmMain.mnurptRSubOutStaSale_Click (1)
frmRptRegi.txttemp(2) = txttempAL(3) '- Party
Exit Sub
ErrorHandler:
    GProcErrorHandler
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
    Dim Y As String
    Me.Top = 300
    Me.Left = Screen.Width / 2 - Me.Width / 2
    Y = "Select * from " & tblName & " where vtype='GP' and VYear=" & gCYear
    If Trim(OrderField) <> "" Then Y = Y & " order by " & OrderField
    GProcRstOpen rstbl, Y, "O"
    FirstTimeFlag = True
    SetGrid
    SetRelatedRS '--Set Add Less
    '---- Button Visiblity Set As Per User Previlage
    GProcSetButtonVisiblity Me
   ' FrameGrid.Visible = False
    cmdBtn_Click (6)
    cbo(23).ListIndex = 0
    SetGrid
End Sub
Private Sub SetRelatedRS()
Dim X As String
Set rsAddLess = New Recordset
X = "select * from tblAddLess where vtype='GP' and vno=" & txt(0) & " and VYear=" & gCYear
GProcRstOpen rsAddLess, X, "O"
End Sub
Private Sub SaveRelatedRecord()
Dim MillCode As Long
Dim CrCode As Long
Dim DrCode As Long
Dim TaxCode As Long
Dim CountCd As Long
Dim X As String
Dim i As Long
Dim SNo As Long
'======================== TblBags Cartoon GpNo,GpType  ================
'--- Save Grid Records
CheckForNumberGrid
'--- Update tblGpSub
X = "select * from tblGpSub where vtype='GP' and vno=" & txt(0) & " and " _
& " VYear=" & gCYear
GProcRstOpen rstblSub, X, "O"

gCn.Execute "delete from tblGpSub where vtype='GP' and vno=" & txt(0) & " and " _
& " VYear=" & gCYear

MillCode = GProcGetColumnValue("tblMastAccount", "AcName", txttempAL(5), "S", "AcCode", "N")
CountCd = GProcGetColumnValue("tblMastItem", "ItName", txttemp(4), "S", "ItCode", "N", " ItMillcode=" & MillCode) '--ItCode



SNo = 0
With msGrid
    i = 1
    Do While i < .Rows
       If Val(.TextMatrix(i, 1)) > 0 Then
          rstblSub.AddNew
          rstblSub!Vno = CLng(txt(0)) '--Vno(0)
          rstblSub!Vtype = "GP" '--Vtype(1)
          rstblSub!Vdt = CDate(txt(2)) '--Vdt(2)
          rstblSub!VYear = gCYear  '--VYear(18)
          rstblSub!VSubCtrNo = Val(.TextMatrix(i, 0))
          rstblSub!vsubITCODE = CountCd
          rstblSub!vsubBAG = Val(.TextMatrix(i, 1))
          rstblSub!vsubwt = Val(.TextMatrix(i, 2))
          rstblSub!VSubGodownCode = Val(txt(17))
          If rstblSub!Vsizer = 0 Or IsNull(rstblSub!Vsizer) Then
             rstblSub!Vsizer = 1
          Else
          End If
          SNo = SNo + 1
          rstblSub.Update
        End If
        i = i + 1
     Loop
End With
rstblSub.Close
 
 
 
 '=====================================
'--- Account Code
MillCode = GProcGetColumnValue("tblMastAccount", "AcName", txttempAL(5), "S", "AcCode", "N")
CrCode = GProcGetColumnValue("tblMastAccount", "AcName", txttempAL(3), "S", "AcCode", "N")
DrCode = GProcGetColumnValue("tblMastAccount", "AcName", txttempAL(4), "S", "AcCode", "N")
TaxCode = GProcGetColumnValue("tblMastNarration", "Narration", "-", "S", "NarrCode", "N")
'-- Delete Records
'DeleteRelatedRecord
'--- TblAddLess
Set rsAddLess = New Recordset
X = "select * from tblAddLess where vtype='GP' and vno=" & txt(0) & " and " _
& " VYear=" & gCYear
GProcRstOpen rsAddLess, X, "O"
If rsAddLess.EOF Then rsAddLess.AddNew
rsAddLess!Vno = CLng(txt(0)) '--Vno(0)
rsAddLess!Vtype = "GP" '--Vtype(1)
rsAddLess!Vdt = CDate(txt(2)) '--Vdt(2)
rsAddLess!AdAcCrCode = CrCode '--Party
rsAddLess!AdAcDrCode = DrCode '--Broker
rsAddLess!AdMillCode = MillCode '--Mill
rsAddLess!ADGPNO = txtAL(6) '--Due Days
rsAddLess!AdTmp2 = CLng(txtE2) '--E2
rsAddLess!AdTmp3 = CLng(txt(41))
If IsDate(txtAL(0)) = True Then
  rsAddLess!AdDespDt = CDate(txtAL(0)) '--Goods Desp Date
End If

'If IsDate(rsAddLess!DespDate) = True Then
'    rsAddLess!AdDespDt = CDate(txtAL(0)) '--Goods Desp DateElse
'Else
'    txtAL(0) = "__/__/____"
'End If

'--Tax Code
rsAddLess!AdTaxCode = TaxCode
rsAddLess!VYear = gCYear  '--VYear(18)
If cbo(23).ListIndex = 4 Then '-- Hank Sales
   rsAddLess!ADTMP7 = "HNK"
Else
   rsAddLess!ADTMP7 = ""
End If
rsAddLess.Update
rsAddLess.Close
End Sub
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
If FormAction = vbDataActionUpdate Then
    MsgBox "You are about to modify the record. Save the record and then exit.", vbInformation + vbOKOnly, Me.Caption
    Cancel = True
    Exit Sub
End If
rstbl.Close
End Sub
Private Sub SetControlEd()
'--  if Cartoon Inward Details
If GProcGetColumnValue("TblBags", "INWNO", IIf(IsNumeric(txttmpInNo), txttmpInNo, 0), "N", "InwNo", "N", "InwType='" & GetInwardType & "'") > 0 Then
    CmdBagDetails.Enabled = True
Else
    CmdBagDetails.Enabled = False
End If
mskDt(43).Enabled = False '-Book Dt
lbl1(16).Visible = False
mskDt(0).Visible = False '-goods desp Dt
If IsRelatedRecord = False Then
    txt(10).Enabled = False '--Lot No
    txttemp(17).Enabled = False '--Godown
    txttmpLRDt.Enabled = False '--LRDt
    txttmpLRNo.Enabled = False '--LR No
    txttmpLRAmt.Enabled = False  '--LR Amt
    txt(8).Enabled = False
    If cbo(23).ListIndex = 3 Then '-- SIT
        lbl1(16).Visible = True
        mskDt(0).Visible = True '-goods desp Dt
        mskDt(0).Enabled = True '-goods desp Dt
        txt(5).Enabled = True
        txt(6).Enabled = True
        lblLRDate.Visible = True
        lblLRNo.Visible = True
        lblLRAmt.Visible = True
        txttmpLRDt.Visible = True '--LRDt
        txttmpLRNo.Visible = True '--LR No
        txttmpLRAmt.Visible = True '--LR amt
        Chk(1).Visible = True
        Chk(1).Enabled = True
        Chk(0).Enabled = True
        Chk(0).Visible = True
    Else
        If FormAction = vbDataActionAddNew Or FormAction = vbDataActionUpdate Then
            txt(5).Enabled = True
            txt(6).Enabled = True
        End If
        lblLRDate.Visible = False
        lblLRNo.Visible = True
        lblLRAmt.Visible = False
        txttmpLRDt.Visible = False '--LRDt
        txttmpLRNo.Visible = False '--LR No
        txttmpLRAmt.Visible = False '--LR Amt
        Chk(0).Enabled = True
        Chk(1).Visible = True
    End If
'---TEMP CHANGE FOR MIX LR SALE
    If cbo(23).ListIndex = 3 Then '-- SIT
        txt(5).Enabled = True
        txt(6).Enabled = True
        lblLRDate.Visible = True
        lblLRNo.Visible = True
        lblLRAmt.Visible = True
        txttmpLRDt.Visible = True '--LRDt
        txttmpLRNo.Visible = True '--LR No
        txttmpLRAmt.Visible = True '--LR amt
        txttmpLRDt.Enabled = True  '--LRDt
        txttmpLRNo.Enabled = True  '--LR No
        txttmpLRAmt.Enabled = True  '--LR amt
    End If

'-----------
    
End If
End Sub

Private Sub txtAl_GotFocus(Index As Integer)
If Index = 6 And Val(txtAL(6)) = 0 Then txtAL(6) = 30  '-- 7 For Universal = 15
GProcSelectBox Me.ActiveControl
End Sub
Private Sub txtAl_LostFocus(Index As Integer)
SaveBtnEd
If Index = 6 And IsNumeric(txtAL(6)) = False Then txtAL(6) = 0
End Sub
Private Sub txtAl_KeyPress(Index As Integer, KeyAscii As Integer)
If KeyAscii = 13 Then Exit Sub
KeyAscii = GProcValidateKey(rsAddLess, Index, KeyAscii, txtAL(Index).text)
End Sub
Private Sub txtTemp_LostFocus(Index As Integer)
SaveBtnEd
End Sub
Private Sub txtTempAL_LostFocus(Index As Integer)
If txttempAL(3) = "" Then
   lblGST.Caption = "Party's GSTIN :"
End If
lblGST.Caption = GProcGetColumnValue("tblMastAccount", "AcName", txttempAL(3), "S", "AcGSTIN", "S")
SaveBtnEd
SetControlEd
End Sub
Private Sub txttempAL_KeyPress(Index As Integer, KeyAscii As Integer)
If KeyAscii = 13 Then Exit Sub
Dim LeftPos As Long
Dim TopPos As Long
Dim BrkAcCode As Long
TopPos = txttempAL(Index).Top + Me.Top + 650
LeftPos = txttempAL(Index).Left + Me.Left
Select Case Index
    Case 3 '-- Party
        gClsSearch.SearchMultiField "QryMastaccount", "Account,AcName,AcCity,AgName", Array("Party Name", "AcName", "City", "Group"), Array(3000, 0, 2000, 3000), "(Agcode=90015 or GpCode=90015 or Agcode=90017 or GpCode=90017  )", Chr(KeyAscii), "acName", LeftPos, TopPos, True
        KeyAscii = 0
        If gClsSearch.SearchMultiRetCol(0) <> "" And gClsSearch.SearchMultiRetCol(1) = "" Then
            txttempAL(Index) = gClsSearch.SearchMultiRetCol(0)
        Else
            txttempAL(Index) = gClsSearch.SearchMultiRetCol(1)
        End If
        '-- Broker from master
        If txttempAL(4) = "" Then
            BrkAcCode = GProcGetColumnValue("TblMastAccount", "ACName", txttempAL(Index), "S", "AcBrkCode", "N")
            txttempAL(4) = GProcGetColumnValue("TblMastAccount", "ACCode", CStr(BrkAcCode), "N", "ACname", "S")
        End If
        ChkOutStLimit
    Case 4 '-- Broker
        gClsSearch.SearchMultiField "QryMastaccount", "Account,AcName,AcCity,AgName", Array("Broker", "AcName", "City", "Group"), Array(3000, 0, 2000, 3000), "(AgCode=90016 or GpCode=90016)", Chr(KeyAscii), "AcName", LeftPos, TopPos, True
        KeyAscii = 0
        If gClsSearch.SearchMultiRetCol(0) <> "" And gClsSearch.SearchMultiRetCol(1) = "" Then
             txttempAL(Index) = gClsSearch.SearchMultiRetCol(0)
        Else
             txttempAL(Index) = gClsSearch.SearchMultiRetCol(1)
        End If
    Case 5 '-- Mill
        If cbo(23).ListIndex = 0 Or cbo(23).ListIndex = 4 Then  '--Trade
            gClsSearch.SearchMultiField "QryMastaccount", "Account,AcName,AcCity,AgName", Array("Mill", "AcName", "City", "Group"), Array(3000, 0, 2000, 3000), " (AgCode=90029 and (AcMillType =0 or AcMillType=5) )", Chr(KeyAscii), "AcName", LeftPos, TopPos, True
        ElseIf cbo(23).ListIndex = 1 Then '--Cons
            gClsSearch.SearchMultiField "QryMastaccount", "Account,AcName,AcCity,AgName", Array("Mill", "AcName", "City", "Group"), Array(3000, 0, 2000, 3000), " (AgCode=90029 and AcMillType =1  )", Chr(KeyAscii), "AcName", LeftPos, TopPos, True
        ElseIf cbo(23).ListIndex = 2 Then '--Depot
            gClsSearch.SearchMultiField "QryMastaccount", "Account,AcName,AcCity,AgName", Array("Mill", "AcName", "City", "Group"), Array(3000, 0, 2000, 3000), " (AgCode=90029 and AcMillType =2  )", Chr(KeyAscii), "AcName", LeftPos, TopPos, True
        ElseIf cbo(23).ListIndex = 3 Then '--SIT
            gClsSearch.SearchMultiField "QryMastaccount", "Account,AcName,AcCity,AgName", Array("Mill", "AcName", "City", "Group"), Array(3000, 0, 2000, 3000), " (AgCode=90029 and (AcMillType =4 or AcMillType=5) )", Chr(KeyAscii), "AcName", LeftPos, TopPos, True
        End If
        KeyAscii = 0
        If gClsSearch.SearchMultiRetCol(0) <> "" And gClsSearch.SearchMultiRetCol(1) = "" Then
            txttempAL(Index) = gClsSearch.SearchMultiRetCol(0)
        Else
            txttempAL(Index) = gClsSearch.SearchMultiRetCol(1)
        End If
End Select
End Sub
Private Sub txttempAL_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
Dim GName As String
If KeyCode = vbKeyF3 Then
    Select Case Index
        Case 3 '-- Party
            GName = GProcGetColumnValue("TblMastGroup", "AgCode", 90015, "N", "AgName", "S")
            GProcShowForm frmMastAccount, frmMain.mnuMstfrm(0), Array(frmMastAccount.txttemp(2)), Array(GName)
        Case 4 '-- Broker
           GName = GProcGetColumnValue("TblMastGroup", "AgCode", 90016, "N", "AgName", "S")
           GProcShowForm frmMastAccount, frmMain.mnuMstfrm(0), Array(frmMastAccount.txttemp(2)), Array(GName)
        Case 5 '-- millcode
            GName = GProcGetColumnValue("TblMastGroup", "AgCode", 90029, "N", "AgName", "S")
            GProcShowForm frmMastAccount, frmMain.mnuMstfrm(0), Array(frmMastAccount.txttemp(2)), Array(GName)
    End Select
End If
End Sub
Private Sub txtTempAL_GotFocus(Index As Integer)
GProcSelectBox Me.ActiveControl
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
If Index = 5 Then oldQty = txt(5) '-- OldQty
End Sub
Private Sub txt_LostFocus(Index As Integer)
GProcCheckForNumber Me, rstbl, MaxNo
CalAmount
SaveBtnEd
If Index = 30 Then
   txtGrid.Visible = True
End If
If Index = 30 And txtGrid.Visible = True Then
   msGrid.Col = 1
   txtGrid.SetFocus
   msGrid_EnterCell
End If
End Sub
Private Sub txt_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyF3 Then
Select Case Index
'    Case 12 '- Purchase
'        gfrmTypeStr = GetInwardType
'        If gfrmTypeStr = "PY" Then
'            GProcShowForm frmPurchaseTrade, gfrmTypeStr
'        ElseIf gfrmTypeStr = "PT" Then
'            GProcShowForm frmPurchaseSIT, gfrmTypeStr
'        Else
'            GProcShowForm frmPurchaseInward, gfrmTypeStr
'        End If
End Select
End If
End Sub

Private Sub txt_KeyPress(Index As Integer, KeyAscii As Integer)
If KeyAscii = 13 Then Exit Sub
KeyAscii = GProcValidateKey(rstbl, Index, KeyAscii, txt(Index).text)
If Index = 42 Then '-- Party Booking
    PrepareBooking
    KeyAscii = 0
End If
End Sub
'=== Check Is Party Outstanding more than O/S Limit
Private Sub ChkOutStLimit()
Dim RsO As Recordset
Dim X As String
Dim DtStr As String
Dim AcCode As Long
Dim OSLimit As Double
Dim OSAmt As Double
AcCode = GProcGetColumnValue("tblMastAccount", "AcName", txttempAL(3), "S", "AcCode", "N")
'-------------------- Os Amt > Os Limt
'- Os Limit
OSLimit = GProcGetColumnValue("tblMastAccount", "AcName", txttempAL(3), "S", "ACOSLIMIT", "N")
If OSLimit > 0 And IsDate(mskDt(2)) = True Then
    If gBackEndDB = gBackEndAccess Then
        DtStr = " vdt<= cdate('" & CDate(mskDt(2)) & "')"
    Else '- Oracle
        DtStr = " vdt<= to_date('" & CDate(mskDt(2)) & "','dd/MM/yyyy')"
    End If
    '-- Os Amt
    X = "Select sum(OutAmount-OutRecAmt) from tblOutStanding where " _
    & " OutAcCode=" & AcCode & " and " & DtStr & "  and vtype in ('SY','ST','SO','SD')"
    If GProcRstOpen(RsO, X, "R") > 0 Then
        If IsNumeric(RsO.Fields(0)) = True Then OSAmt = RsO.Fields(0)
        If OSAmt > OSLimit Then MsgBox "Party Outstanding amount " & OSAmt & "  is more than its O/S limit " & OSLimit, vbInformation + vbOKOnly, Me.Caption
    End If
End If
'-------------------- (GatePass Dt - Os Bill Dt)  > 25
If IsDate(mskDt(2)) = True Then
    If gBackEndDB = gBackEndAccess Then
        DtStr = " (cdate('" & CDate(mskDt(2)) & "') - vdt) >25"
    Else '- Oracle
        DtStr = " (to_date('" & CDate(mskDt(2)) & "','dd/MM/yyyy') - vdt) >25"
    End If
    '-- Os Bills
    X = "Select count(*) from tblOutStanding where OutAcCode=" & AcCode & " and " _
    & " (OutAmount-OutRecAmt) >0  and vtype in ('SY','ST','SO','SD') and " & DtStr
    If GProcRstOpen(RsO, X, "R") > 0 Then
        If RsO.Fields(0) > 0 Then MsgBox "Party having " & RsO.Fields(0) & " outstanding bill with age more than 25 days.", vbInformation + vbOKOnly, Me.Caption
    End If
End If
'if tmptblOutStanding.OutDbNtAmt}-{tmptblOutStanding.OutRecDbAmt}
End Sub
Private Sub SaveBtnEd()
'-- Vno,Vdt,Sale Type,Party,Broker,Mill,Wt,Inward No
If txt(0) = 0 Or GProcIsDateValid(mskDt(2)) = False Or cbo(23) = "" Or txttempAL(3) = "" Or txttempAL(4) = "" Or txttempAL(5) = "" Or txttemp(4) = "" Or txt(12) = 0 Or txt(6) = 0 Then
    cmdBtn(13).Enabled = False
    cmdBtn(13).BackColor = &H80000004
Else
    cmdBtn(13).Enabled = True
    cmdBtn(13).BackColor = &HFFFF80
End If
End Sub
Private Sub FillTxtFromTemp()
Dim MillCode As Long
Dim ItTicket As Long
txt(1) = "GP" '--VtYpe
txt(18) = gCYear
txt(3) = 1 '--Ctrl No
txt(2) = IIf(IsDate(mskDt(2)), mskDt(2), "") '--Vdt
txt(43) = IIf(IsDate(mskDt(43)), mskDt(43), "") '--Book Vdt
txtAL(0) = IIf(IsDate(mskDt(0)), mskDt(0), "") '--Goods Desp Date
If cbo(23).ListIndex = -1 Then
    txt(23) = 0
Else
    txt(23) = cbo(23).ListIndex '-- Sl Type
End If
MillCode = GProcGetColumnValue("tblMastAccount", "AcName", txttempAL(5), "S", "AcCode", "N")
txt(4) = GProcGetColumnValue("tblMastItem", "ItName", txttemp(4), "S", "ItCode", "N", " ItMillcode=" & MillCode) '--ItCode
txt(17) = GProcGetColumnValue("tblMastnarration", "Narration", txttemp(17), "S", "NarrCode", "N")
If FormAction = vbDataActionAddNew Then txt(11) = GProcGenerateId("tblVouSub", "VSubItCtrlNo", "Vyear=" & gCYear)
If IsNumeric(txtAL(6)) = False Then txtAL(6) = 0
txtE2 = Chk(1)
txt(41) = Chk(0)
txt(20) = txttmpLRDt
txt(19) = txttmpLRNo
txt(21) = txttmpLRAmt
txt(26) = chkIsIncludingGst.Value
LblHsn.Caption = GProcGetColumnValue("tblMastItem", "ItName", txttemp(4), "S", "Ittmp2", "S", "ITMillCode=" & MillCode)
ItTicket = GProcGetColumnValue("tblMastItem", "ItName", txttemp(4), "S", "ItTicketPrint", "N", "ITMillCode=" & MillCode)
If ItTicket = 1 Then
   LblTicket.Caption = "Print Ticket"
Else
   LblTicket.Caption = ""
End If
End Sub
Private Sub DeleteRelatedRecord()
'--- Voucher Sub
gCn.Execute "delete from tblAddLess where vtype='GP' and vno=" & txt(0) & " and " _
& " VYear=" & gCYear
End Sub
Private Sub FillTempFromTxt()
Dim ItTicket As Integer
'--Tmp In No
txttmpInNo = GProcGetColumnValue("tblVouSub", "VSubItCtrlNo", txt(12), "N", "Vno", "N")
'--Tmp LR No
txttmpLRNo = GProcGetColumnValue("tblVouSub", "VSubItCtrlNo", txt(11), "N", "VSubLRno", "S")
'--Tmp LR Dt
txttmpLRDt = GProcGetColumnValue("tblVouSub", "VSubItCtrlNo", txt(11), "N", "VSubLRDate", "N")
If IsDate(txttmpLRDt) = True Then txttmpLRDt = Format(txttmpLRDt, "dd/mm/yyyy")
'--Tmp LR Amt
txttmpLRAmt = GProcGetColumnValue("tblVouSub", "VSubItCtrlNo", txt(11), "N", "VSubLRAmt", "N")


mskDt(2) = IIf(IsDate(txt(2)), CDate(txt(2)), "__/__/____") '--Vdt
If IsDate(txt(43)) = True Then '--Book dt
    mskDt(43) = CDate(txt(43))
Else
    mskDt(43) = "__/__/____"
End If
cbo(23).ListIndex = txt(23) '--SL Type
If cbo(23).ListIndex < 0 Then cbo(23).ListIndex = 0
txttemp(4) = GProcGetColumnValue("tblMastItem", "ItCode", txt(4), "N", "ItName", "S") '--Item
txttemp(17) = GProcGetColumnValue("tblMastNarration", "NarrCode", txt(17), "N", "Narration", "S") '--Godown
'-- Fill Related Record
SetRelatedRS
FillRelatedRecord '-TblBags
'--Due Days
txtAL(6) = rsAddLess.Fields("AdGpNo")
txtAL(0) = rsAddLess.Fields("AdDespDt") & vbNullString
If IsDate(txtAL(0)) = True Then '--Goods Desp Date
    mskDt(0) = CDate(txtAL(0))
Else
    mskDt(0) = "__/__/____"
End If
'--Party
txttempAL(3) = GProcGetColumnValue("tblMastAccount", "AcCode", rsAddLess.Fields("AdAcCrCode"), "N", "AcName", "T")
'-- Broker
txttempAL(4) = GProcGetColumnValue("tblMastAccount", "AcCode", rsAddLess.Fields("AdAcDrCode"), "N", "AcName", "T")
'--Mill
txttempAL(5) = GProcGetColumnValue("tblMastAccount", "AcCode", rsAddLess.Fields("AdMillCode"), "N", "AcName", "T")
oldQty = txt(5)
Chk(1) = rsAddLess.Fields("AdTmp2")
Chk(0) = rsAddLess.Fields("AdTmp3")
chkIsIncludingGst.Value = Val(txt(26))
If txttempAL(3) = "" Then
   lblGST.Caption = "Party's GSTIN"
End If
lblGST.Caption = GProcGetColumnValue("tblMastAccount", "AcName", txttempAL(3), "S", "AcGSTIN", "S")
LblHsn.Caption = GProcGetColumnValue("tblMastItem", "ItName", txttemp(4), "S", "Ittmp2", "S", "ITMillCode=" & rsAddLess.Fields("AdMillCode"))

ItTicket = GProcGetColumnValue("tblMastItem", "ItName", txttemp(4), "S", "ItTicketPrint", "N", "ITMillCode=" & rsAddLess.Fields("AdMillCode"))
If ItTicket = 1 Then
   LblTicket.Caption = "Print Ticket"
Else
   LblTicket.Caption = ""
End If

SetControlEd
End Sub
Private Function ValidateData() As Boolean
Dim GName As String
Dim MillCode As Long
Dim MstMinRate As Double
Dim MstMaxRate As Double
Dim Rs1 As Recordset
Dim rstAddCheck As Recordset
Dim X As String
Dim ItCode As Long
Dim InVtype As String
GProcCheckForNumber Me, rstbl, MaxNo
CalAmount
'-Grid Validate
If ValidateGrid = False Then Exit Function
'if CmdBagDetails
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
    '---Validate Despatch Date
    If GProcIsDateValid(mskDt(0)) = True Then
        If (CDate(mskDt(0)) < CDate(mskDt(2))) Then
           MsgBox "Check Despatch Date Date.", vbInformation + vbOKOnly, Me.Caption
           mskDt(0).SetFocus
           Exit Function
        End If
    End If
    '---- Party Account
    If txttempAL(3) = "" Then
        MsgBox "Check Party.", vbInformation + vbOKOnly, Me.Caption
        txttempAL(3).SetFocus
        Exit Function
    End If
    '--Create Party
    If GProcGetColumnValue("tblMastAccount", "AcName", txttempAL(3), "S", "AcCode", "N") = 0 Then
       GName = GProcGetColumnValue("TblMastGroup", "AgCode", 90015, "N", "AgName", "S")
       GProcShowForm frmMastAccount, frmMain.mnuMstfrm(0), Array(frmMastAccount.txt(1), frmMastAccount.txttemp(2)), Array(txttempAL(3), GName)
       Exit Function
    End If
    '---- Broker
    If txttempAL(4) = "" Or txttempAL(4) = "-" Then
        MsgBox "Check Broker.", vbInformation + vbOKOnly, Me.Caption
        txttempAL(4).SetFocus
        Exit Function
    End If
    '--Create Broker
    If GProcGetColumnValue("tblMastAccount", "AcName", txttempAL(4), "S", "AcCode", "N") = 0 Then
       GName = GProcGetColumnValue("TblMastGroup", "AgCode", 90016, "N", "AgName", "S")
       GProcShowForm frmMastAccount, frmMain.mnuMstfrm(0), Array(frmMastAccount.txt(1), frmMastAccount.txttemp(2)), Array(txttempAL(4), GName)
       Exit Function
    End If
    '---- Mill
    If txttempAL(5) = "" Then
        MsgBox "Check Broker.", vbInformation + vbOKOnly, Me.Caption
        txttempAL(5).SetFocus
        Exit Function
    End If
    '--Create Mill
    If GProcGetColumnValue("tblMastAccount", "AcName", txttempAL(5), "S", "AcCode", "N") = 0 Then
       GName = GProcGetColumnValue("TblMastGroup", "AgCode", 90029, "N", "AgName", "S")
       GProcShowForm frmMastAccount, frmMain.mnuMstfrm(0), Array(frmMastAccount.txt(1), frmMastAccount.txttemp(2)), Array(txttempAL(5), GName)
       Exit Function
    End If
    '---Count
    If txttemp(4) = "" Then
        MsgBox "Check Count.", vbInformation + vbOKOnly, Me.Caption
        txttemp(4).SetFocus
        Exit Function
    End If
    MillCode = GProcGetColumnValue("TblMastAccount", "AcName", txttempAL(5), "S", "Accode", "N")
    If GProcGetColumnValue("TblMastItem", "ItName", txttemp(4), "S", "Itcode", "N", "ITMillCode=" & MillCode) = 0 Then
        'MSGrid.SetFocus
        GProcShowForm frmMastItem, frmMain.mnuMstfrm(3), Array(txttemp(4), frmMastItem.txttemp(9)), Array(txttemp(4), txttempAL(5))
        Exit Function
    End If
    '-- Inward No
    If txt(12) = 0 Then
        MsgBox "Check Inward No.", vbInformation + vbOKOnly, Me.Caption
        txt(12).SetFocus
        Exit Function
    End If
    '-- Wt
    If CDbl(txt(6)) = 0 Then
        MsgBox "Check Wt.", vbInformation + vbOKOnly, Me.Caption
        txt(6).SetFocus
        Exit Function
    End If
    '-- Rate
    If CDbl(txt(7)) = 0 Then
        MsgBox "Check Rate.", vbInformation + vbOKOnly, Me.Caption
        txt(7).SetFocus
        Exit Function
    End If
    MstMinRate = GProcGetColumnValue("TblMastItem", "ItName", txttemp(4), "S", "ITMINRATE", "N", "ITMillCode=" & MillCode)
    MstMaxRate = GProcGetColumnValue("TblMastItem", "ItName", txttemp(4), "S", "ITMaxRATE", "N", "ITMillCode=" & MillCode)
    If CDbl(txt(7)) < MstMinRate Then
        MsgBox "Check Rate.Rate is less than Min Rate. " & MstMinRate, vbInformation + vbOKOnly, Me.Caption
        txt(7).SetFocus
        Exit Function
    End If
    If CDbl(txt(7)) > MstMaxRate Then
        MsgBox "Check Rate.Rate is more than Max Rate. " & MstMaxRate, vbInformation + vbOKOnly, Me.Caption
        txt(7).SetFocus
        Exit Function
    End If
    '-- Validate Wt,Bag
    InVtype = GetInwardType
    ItCode = GProcGetColumnValue("tblMastItem", "ItName", txttemp(4), "S", "ItCode", "N", " ItMillCode=" & MillCode)
    If FormAction = vbDataActionAddNew Then '--Add
        X = "Select vsublotno,(vsubbag-(vsubissbag)) as BalBag," _
        & " (vsubwt-(vsubisswt)) as BalWt from tblvouSub where vsubitctrlno=" & CLng(txt(12)) & " and vsubitcode=" & ItCode & " and "
    Else '--modify
        X = "Select vsublotno,(vsubbag-(vsubissbag-vSubTmpBag)) as BalBag," _
        & " (vsubwt-(vsubisswt-vSubTmpWt)) as BalWt from tblvouSub where vsubitctrlno=" & CLng(txt(12)) & " and vsubitcode=" & ItCode & " and "
    End If
    If gBackEndDB = gBackEndAccess Then
        X = X + "  vdt<=cdate('" & CDate(mskDt(2)) & "') and vtype ='" & InVtype & "'"
    Else '-- Oracle
        X = X + "  vdt<= to_date('" & CDate(mskDt(2)) & "','dd/MM/yyyy') and vtype ='" & InVtype & "'"
    End If
    GProcRstOpen Rs1, X, "R"
    If Rs1.EOF And Rs1.BOF Then
        MsgBox "Check Inward No.", vbInformation + vbOKOnly, Me.Caption
       ' txt(12).SetFocus
        Exit Function
    Else
'        '-- Balance Wt
'        If CDbl(txt(6)) > Rs1.Fields("BalWt") Then
'            MsgBox "Check Wt.Balance Wt is " & Rs1.Fields("BalWt"), vbInformation + vbOKOnly, Me.Caption
'            'txt(6).SetFocus
'            'Exit Function
'        End If
        '-- Balance Qty
        If CDbl(txt(5)) > Rs1.Fields("BalBag") Then
            MsgBox "Check Bag.Balance Bags are " & Rs1.Fields("BalBag"), vbInformation + vbOKOnly, Me.Caption
            txt(5).SetFocus
            Exit Function
        End If
    End If
    '--Validate Booking No
'    If txt(42) > 0 Then
'        ProcInsertCodeNameInTmpTbl '--- Insert Code & Name In Setup TmpTbl As Per  Comp code,Year
'        X = " select tblBookingSub.VNo,tblBookingSub.vdt from " _
'        & " tblBookingSub,tblBooking,TmpGenTbl Where TmpGenTbl.UserName='" & gUserName & "' and " _
'        & " tblBookingSub.vno=tblBooking.vno and tblBookingSub.vType=tblBooking.vType and tblBookingSub.vYear=tblBooking.vYear and " _
'        & " BkAcBrkCode=Amt1 and BKAcCode=Amt2 and BkSubItCode=amt3 and tblBooking.vtype ='OP' and BkSubIsCompleted=0 and Nar1='" & txttempAL(4) & "'" _
'        & "  and (Nar2='" & txttempAL(3) & "' or Nar2 ='') and Nar3='" & txttemp(4) & "' and Nar4='" & txttempAL(5) & "' and tblBooking.Vno=" & txt(42)
'        '-Date Condition
'        X = X + " and tblBooking.vdt<=cdate('" & CDate(mskDt(2)) & "')"
'        If GProcRstOpen(Rs1, X, "R", gSetupCn) = 0 Then
'            MsgBox "Check Booking No.", vbInformation + vbOKOnly, Me.Caption
'            txt(42).SetFocus
'            Exit Function
'        End If
'     End If
    '----Duplicate Records
    '----Addition Case
    If FormAction = vbDataActionAddNew Then
        X = "Select vno From " & tblName & " Where vno=" & CLng(txt(0)) & " and Vtype='GP' and VYear=" & gCYear
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
If KeyCode = vbKeyF3 Then
Select Case Index
    Case 4 '--count
        GProcShowForm frmMastItem, frmMain.mnuMstfrm(3), Array(frmMastItem.txttemp(9)), Array(txttempAL(5))
End Select
End If
End Sub
Private Sub txtTemp_KeyPress(Index As Integer, KeyAscii As Integer)
Dim LeftPos As Long
Dim TopPos As Long
Dim MillCode As Long
Dim ItTicket As Long
If KeyAscii = 13 Then Exit Sub
TopPos = txttemp(Index).Top + Me.Top + 650
LeftPos = txttemp(Index).Left + Me.Left
Select Case Index
    Case 4 '--count
        MillCode = GProcGetColumnValue("TblMastAccount", "AcName", txttempAL(5), "S", "Accode", "N")
        gClsSearch.SearchMultiField "tblMastItem", "ItName,ItStdrateper,ITTMP1", Array("Count Name", "Rate Per", ""), Array(txttemp(Index).Width, 0, 0), "ITMillCode=" & MillCode, Chr(KeyAscii), "ItName", LeftPos, TopPos, True
        KeyAscii = 0
        If gClsSearch.SearchMultiRetCol(0) <> "" Then
            txttemp(Index) = gClsSearch.SearchMultiRetCol(0)
            txt(8) = gClsSearch.SearchMultiRetCol(1) '--Rate Per
            chkIsIncludingGst = Val(gClsSearch.SearchMultiRetCol(2))
            LblHsn.Caption = GProcGetColumnValue("tblMastItem", "ItName", txttemp(4), "S", "Ittmp2", "S", " ItMillcode=" & MillCode)
            ItTicket = GProcGetColumnValue("tblMastItem", "ItName", txttemp(4), "S", "ItTicketPrint", "N", "ITMillCode=" & MillCode)
            If ItTicket = 1 Then
               LblTicket.Caption = "Print Ticket"
            Else
               LblTicket.Caption = ""
            End If
            'If CDbl(txt(8)) = 0 Then txt(8) = gClsSearch.SearchMultiRetCol(1) '--Rate Per
        End If
End Select
End Sub
Private Sub CollectStock(KeyChar As String)
Dim InVtype As String
Dim ItCode As Long
Dim MillCode As Long
Dim QStr As String
Dim QStrFields As String
InVtype = GetInwardType
If IsDate(mskDt(2)) = True Then
    MillCode = GProcGetColumnValue("TblMastAccount", "AcName", txttempAL(5), "S", "Accode", "N")
    ItCode = GProcGetColumnValue("tblMastItem", "ItName", txttemp(4), "S", "ItCode", "N", " ItMillCode=" & MillCode)
   '--- Add
    If FormAction = vbDataActionAddNew Then
        If cbo(23).ListIndex = 3 Then '-- SIT LR Deatils
            QStrFields = "VSubLRNo,format(VSubLRDate,'dd/mm/yyyy'),VNo,vdt,vsubitctrlno,VType,Narration," _
            & " VSubLRAmt ,vsublotno,vsubbag-vsubissbag," & IIf(gBackEndDB = gBackEndAccess, "vsubwt-vsubisswt", "to_char(vsubwt-vsubisswt,999999999.999)")
        Else
            QStrFields = "VNo,vdt,vsubitctrlno,VType,Narration," _
            & " vsublotno,vsubbag-vsubissbag," & IIf(gBackEndDB = gBackEndAccess, "vsubwt-vsubisswt", "to_char(vsubwt-vsubisswt,999999999.999)") & " ,vsubbag,vsubwt,VSubLRNo,VSubLRDate,VSUBTMP7"
        End If
         If cbo(23).ListIndex = 4 Then '-- Hank Sales
            QStrFields = "tblVouSub.VNo,tblVouSub.vdt,vsubitctrlno,tblVouSub.VType,Narration," _
            & " vsublotno,vsubbag-vsubissbag," & IIf(gBackEndDB = gBackEndAccess, "vsubwt-vsubisswt", "to_char(vsubwt-vsubisswt,999999999.999)") & " ,vsubbag,vsubwt,VSubLRNo,VSubLRDate"
            
            QStr = " select " & QStrFields & " from tblVousub,tblMastNarration,tblAddless Where " _
            & " vsubitcode=" & ItCode & " and tblAddless.Vno = tblVousub.Vno and tblAddless.Vtype = tblVouSub.Vtype and tblAddless.vyear = tblVouSub.vyear and tblVousub.VSubGodownCode=tblMastNarration.Narrcode and " _
            & " ((vsubbag-(vsubissbag))> 0) and left(adTmp5,3)= 'HNK'  and tblVouSub.vtype ='" & InVtype & "' and"
        Else
            If cbo(23).ListIndex = 0 Then '-- Trade Sales
                QStrFields = "tblVouSub.VNo,tblVouSub.vdt,vsubitctrlno,tblVouSub.VType,Narration," _
                & " vsublotno,vsubbag-vsubissbag," & IIf(gBackEndDB = gBackEndAccess, "vsubwt-vsubisswt", "to_char(vsubwt-vsubisswt,999999999.999)") & " ,vsubbag,vsubwt,VSubLRNo,VSubLRDate,VSUBTMP7"
                
                QStr = " select " & QStrFields & " from tblVousub,tblMastNarration,tblAddless Where " _
                & " vsubitcode=" & ItCode & " and tblAddless.Vno = tblVousub.Vno and tblAddless.Vtype = tblVouSub.Vtype and tblAddless.vyear = tblVouSub.vyear and tblVousub.VSubGodownCode=tblMastNarration.Narrcode and " _
                & " ((vsubbag-(vsubissbag))> 0) and  (adtmp5 is null or adtmp5 = ' ')  and tblVouSub.vtype ='" & InVtype & "' and"
            Else
                QStr = " select " & QStrFields & " from tblVousub,tblMastNarration Where " _
                & " vsubitcode=" & ItCode & " and tblVousub.VSubGodownCode=tblMastNarration.Narrcode and " _
                & " ((vsubbag-(vsubissbag))> 0) and vtype ='" & InVtype & "' and"
            End If
        End If
'        QStr = " select " & QStrFields & " from tblVousub,tblMastNarration Where " _
'        & " vsubitcode=" & ItCode & " and tblVousub.VSubGodownCode=tblMastNarration.Narrcode and " _
'        & " ((vsubbag-(vsubissbag))> 0) and vtype ='" & InVtype & "' and"
    Else
        If cbo(23).ListIndex = 3 Then '--SIT LR Deatils
            QStrFields = "VSubLRNo,format(VSubLRDate,'dd/mm/yyyy'),VNo,vdt,vsubitctrlno,VType,Narration," _
            & " VSubLRAmt,vsublotno,vsubbag-(vsubissbag-VSubTmpBag)," & IIf(gBackEndDB = gBackEndAccess, "vsubwt-vsubisswt+VSubTmpWt", "to_char(vsubwt-vsubisswt+VSubTmpWt,999999999.999)") & ""
        Else
            QStrFields = "VNo,vdt,vsubitctrlno,VType,Narration," _
            & " vsublotno,vsubbag-(vsubissbag-VSubTmpBag)," & IIf(gBackEndDB = gBackEndAccess, "vsubwt-vsubisswt+VSubTmpWt", "to_char(vsubwt-vsubisswt+VSubTmpWt,999999999.999)") & ",vsubbag,vsubwt,VSubLRNo,VSubLRDate"
        End If
        QStr = " select " & QStrFields & " from tblVousub,tblMastNarration Where " _
        & " vsubitcode=" & ItCode & " and tblVousub.VSubGodownCode=tblMastNarration.Narrcode and " _
        & " ((vsubbag-(vsubissbag-VSubTmpBag))> 0) and vtype ='" & InVtype & "' and "
    End If
    If gBackEndDB = gBackEndAccess Then
       QStr = QStr + " tblVousub.vdt<=cdate('" & CDate(mskDt(2)) & "')"
    Else '-- Oracle
        QStr = QStr + " vdt<= to_date('" & CDate(mskDt(2)) & "','dd/MM/yyyy')"
    End If
    If cbo(23).ListIndex = 3 Then '-- SIT Sales
        gClsSearch.SearchMultiField QStr, "", Array("LR No", "LR Date", "No.", "Date", "CntNo.", "Type", "Godown", "LR Amt", "Lotno", "Bal Bag", "Bal Wt."), Array(1800, 1200, 800, 1200, 0, 0, 1300, 0, 1000, 800, 1000), "", "", "Vdt,vno", 0, 0, False, , True
    Else
        gClsSearch.SearchMultiField QStr, "", Array("No.", "Date", "CntNo.", "Type", "Godown", "Lotno", "Bal Bag", "Bal Wt.", "Bag", "Weight", "", "", "Challan"), Array(800, 1200, 0, 0, 1500, 1000, 800, 1000, 800, 1500, 0, 0, 1500), "", "", "tblVouSub.Vdt,tblVouSub.vno", 0, 0, False, , True
    End If
    DoEvents
    If gClsSearch.SearchMultiRetCol(0) <> "" Then
        If cbo(23).ListIndex = 3 Then    '--SIT Sales
            txt(12) = gClsSearch.SearchMultiRetCol(4)
            txttmpInNo = gClsSearch.SearchMultiRetCol(2)
            txt(10) = gClsSearch.SearchMultiRetCol(8) '--Lot No
            txttemp(17) = gClsSearch.SearchMultiRetCol(6) '--Godown
            txt(6) = 0
            If Val(txt(5)) <= 0 Then txt(5) = CLng(gClsSearch.SearchMultiRetCol(9)) '--Bal Bag
            oldQty = txt(5)
            If Val(txt(6)) <= 0 Then txt(6) = CDbl(gClsSearch.SearchMultiRetCol(10)) '--Bal Wt
            '-- LR No
            txttmpLRNo = gClsSearch.SearchMultiRetCol(0)
            '-- LR Date
            txttmpLRDt = gClsSearch.SearchMultiRetCol(1)
            If IsDate(txttmpLRDt) = True Then txttmpLRDt = Format(txttmpLRDt, "dd/mm/yyyy")
            '-- LR amt
            txttmpLRAmt = gClsSearch.SearchMultiRetCol(7)
       Else
            txt(12) = gClsSearch.SearchMultiRetCol(2)
            txt(29) = gClsSearch.SearchMultiRetCol(12)
            txttmpInNo = gClsSearch.SearchMultiRetCol(0)
            txt(10) = gClsSearch.SearchMultiRetCol(5) '--Lot No
            txttemp(17) = gClsSearch.SearchMultiRetCol(4) '--Godown
            txt(6) = 0
            If Val(txt(5)) <= 0 Then txt(5) = CLng(gClsSearch.SearchMultiRetCol(6)) '--Bal Bag
            oldQty = txt(5)
            If Val(txt(6)) <= 0 Then txt(6) = CDbl(gClsSearch.SearchMultiRetCol(7)) '--Bal Wt
       End If
            ' txt(6) = txt(5) * txt(6)
            CalAmount
    End If
End If
End Sub
'Private Sub CollectStock()
'Dim InVtype As String
'Dim ItCode As Long
'Dim MillCode As Long
'Dim QStr As String
'Dim QStrFields As String
'InVtype = GetInwardType
'If IsDate(mskDt(2)) = True Then
'    MillCode = GProcGetColumnValue("TblMastAccount", "AcName", txttempAL(5), "S", "Accode", "N")
'    ItCode = GProcGetColumnValue("tblMastItem", "ItName", txttemp(4), "S", "ItCode", "N", " ItMillCode=" & MillCode)
'   '--- Add
'    If FormAction = vbDataActionAddNew Then
'        QStrFields = "vsubitctrlno,VSubLRNo,VSubLRDate,VNo,vdt,VType,Narration," _
'        & " VSubLRAmt ,vsublotno,vsubbag-vsubissbag," & IIf(gBackEndDB = gBackEndAccess, "vsubwt-vsubisswt", "to_char(vsubwt-vsubisswt,999999999.999)") & ",vsubbag,vsubwt"
'        QStr = " select " & QStrFields & " from tblVousub,tblMastNarration Where " _
'        & " vsubitcode=" & ItCode & " and tblVousub.VSubGodownCode=tblMastNarration.Narrcode and " _
'        & " ((vsubbag-(vsubissbag))> 0) and vtype ='" & InVtype & "' and"
'    Else
'        QStrFields = "vsubitctrlno,VSubLRNo,VSubLRDate,VNo,vdt,VType,Narration," _
'        & " VSubLRAmt,vsublotno,vsubbag-(vsubissbag-VSubTmpBag)," & IIf(gBackEndDB = gBackEndAccess, "vsubwt-vsubisswt+VSubTmpWt", "to_char(vsubwt-vsubisswt+VSubTmpWt,999999999.999)") & ",vsubbag,vsubwt"
'        QStr = " select " & QStrFields & " from tblVousub,tblMastNarration Where " _
'        & " vsubitcode=" & ItCode & " and tblVousub.VSubGodownCode=tblMastNarration.Narrcode and " _
'        & " ((vsubbag-(vsubissbag-VSubTmpBag))> 0) and vtype ='" & InVtype & "' and "
'    End If
'    If gBackEndDB = gBackEndAccess Then
'        QStr = QStr + " vdt<=cdate('" & CDate(mskDt(2)) & "')"
'    Else '-- Oracle
'        QStr = QStr + " vdt<= to_date('" & CDate(mskDt(2)) & "','dd/MM/yyyy')"
'    End If
'    If Cbo(23).ListIndex = 3 Then '--For SIT Show LR Detail
'        gClsSearch.SearchMultiField QStr, "", Array("CntNo.", "LR No", "LR Date", "No.", "Date", "Type", "Godown", "LR Amt", "Lotno", "Bal Bag", "Bal Wt.", "Bag", "Wt."), Array(0, 800, 1200, 800, 1200, 0, 1500, 0, 1000, 800, 1000, 800, 1000), "", "", "Vdt,vno", 0, 0, False, , True
'    Else
'        gClsSearch.SearchMultiField QStr, "", Array("CntNo.", "LR No", "LR Date", "No.", "Date", "Type", "Godown", "LR Amt", "Lotno", "Bal Bag", "Bal Wt.", "Bag", "Wt."), Array(0, 0, 0, 800, 1200, 0, 1500, 0, 1000, 800, 1000, 800, 1000), "", "", "Vdt,vno", 0, 0, False, , True
'    End If
'    DoEvents
'    If gClsSearch.SearchMultiRetCol(0) <> "" Then
'        txt(12) = gClsSearch.SearchMultiRetCol(0)
'        txttmpInNo = gClsSearch.SearchMultiRetCol(3)
'        txt(10) = gClsSearch.SearchMultiRetCol(8) '--Lot No
'        txttemp(17) = gClsSearch.SearchMultiRetCol(6) '--Godown
'        txt(6) = 0
'        If Val(txt(5)) <= 0 Then txt(5) = CLng(gClsSearch.SearchMultiRetCol(9)) '--Bal Bag
'        oldQty = txt(5)
'        If Val(txt(6)) <= 0 Then txt(6) = CDbl(gClsSearch.SearchMultiRetCol(10)) '--Bal Wt
'        '-- LR No
'        txttmpLRNo = gClsSearch.SearchMultiRetCol(1)
'        '-- LR Date
'        txttmpLRDt = gClsSearch.SearchMultiRetCol(2)
'        If IsDate(txttmpLRDt) = True Then txttmpLRDt = Format(txttmpLRDt, "dd/mm/yyyy")
'        '-- LR amt
'        txttmpLRAmt = gClsSearch.SearchMultiRetCol(7)
'        CalAmount
'    End If
'End If
'End Sub
Private Sub CalAmount()
Dim Wt As Double
Dim Qty As Long
Dim MillCode As Long
Dim StdP As Double
'-- Booking No=0
If txt(42) = 0 Then
    mskDt(43) = "__/__/____" '- Bk Dt
    txt(44) = 0 '- Bk ItCtrl No
End If
Qty = CLng(txt(5))
If Qty = 0 Then txt(6) = 0


If (CDbl(txt(6)) = 0) Or txt(5) <> oldQty Then
    MillCode = GProcGetColumnValue("TblMastAccount", "AcName", txttempAL(5), "S", "AcCode", "N")
    StdP = CDbl(GProcGetColumnValue("tblMastItem", "ItName", txttemp(4), "S", "ItStdpack", "N", " ItMillCode= " & MillCode))
    If StdP > 0 Then
       Wt = Qty * StdP
    End If
    If Wt > 0 Then
       txt(6) = Wt
    End If
End If

End Sub
Private Function GetInwardType() As String
    If cbo(23).ListIndex = 0 Or cbo(23).ListIndex = -1 Or cbo(23).ListIndex = 4 Then
        GetInwardType = "PY"
    ElseIf cbo(23).ListIndex = 1 Or cbo(23).ListIndex = 2 Then
        GetInwardType = "PI"
    ElseIf cbo(23).ListIndex = 3 Then
        GetInwardType = "PT"
    End If
End Function

Private Sub txttmpInNo_GotFocus()
GProcSelectBox Me.ActiveControl
End Sub
Private Sub txttmpInNo_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyReturn Then Exit Sub
CollectStock Chr(KeyAscii)
KeyAscii = 0
End Sub

Private Sub txttmpInNo_KeyDown(KeyCode As Integer, Shift As Integer)
Dim oldFrmTypeStr As String
If KeyCode = vbKeyF3 Then
    '- Purchase
    oldFrmTypeStr = gfrmTypeStr
    gfrmTypeStr = GetInwardType
    If gfrmTypeStr = "PY" Then
        GProcShowForm frmPurchaseTrade, frmMain.mnutrnSubPurfrm(0), gfrmTypeStr
    ElseIf gfrmTypeStr = "PT" Then
        GProcShowForm frmPurchaseSIT, frmMain.mnutrnSubPurfrm(1), gfrmTypeStr
    Else
        GProcShowForm frmPurchaseInward, frmMain.mnutrnSubPurfrm(3), gfrmTypeStr
    End If
    gfrmTypeStr = oldFrmTypeStr
End If
End Sub
'--- Related Sale
Private Function IsRelatedRecord() As Boolean
Dim X As String
Dim Rs1 As Recordset
If IsDate(mskDt(2)) = True Then
    If gBackEndDB = gBackEndAccess Then
        X = "Select Vtype,Vdt,Vno from tblAddLess where AdGPno=" & txt(0) & " and AdGPDt=cdate('" & mskDt(2) & "')"
    Else '-- Oracle
        X = "Select Vtype,Vdt,Vno from tblAddLess where AdGPno=" & txt(0) & " and AdGPDt= to_date('" & mskDt(2) & "','dd/MM/yyyy')"
    End If
    GProcRstOpen Rs1, X, "R"
    If Rs1.EOF And Rs1.BOF Then
    Else
 '       MsgBox "Related Invoice is exists.You can not Modify/Delete this Record.", vbCritical + vbOKOnly, App.Title
        IsRelatedRecord = True
        Exit Function
    End If
End If
End Function
'===  Print
Private Sub txtno_GotFocus(Index As Integer)
GProcSelectBox Me.ActiveControl
End Sub

Private Sub txtno_KeyPress(Index As Integer, KeyAscii As Integer)
KeyAscii = GProcNumberOnly(KeyAscii)
End Sub

Private Sub cmdPrint_Click(Index As Integer)
Dim RsGp As Recordset
Dim X As String
Dim GpNo As Long
Select Case Index
Case 0, 1, 4:
    '------ Validate
    If IsNumeric(txtno(0)) = False Then txtno(0) = 0
    If IsNumeric(txtno(1)) = False Then txtno(1) = txtno(0)
    If CLng(txtno(1)) < CLng(txtno(0)) Then txtno(1) = txtno(0)
    '-------
    If Index = 0 Then GProcCrystalRptPreparation CryReport, "Printer"
    If Index = 1 Then GProcCrystalRptPreparation CryReport, "Window"
    If Index = 4 Then GProcCrystalRptPreparation CryReport, "Window"
    CryReport.WindowTitle = "Gate Pass"
    GpNo = CLng(txtno(0))
    Do While GpNo <= txtno(1)
       X = "Select tblAddLess.Vno,tblAddLess.Vdt,tblAddLess.ADGPNO,tblMastAccount.AcName,tblMastAccount.AcAdd1,tblMastAccount.AcAdd2,tblMastAccount.AcCity,tblMastAccount.AcGstin,tblMastAccount.AcFaxNo," _
            & "tblMastAccount_Br.AcName,tblMastItem.ItName,tblMastItem.ItTicket,tblMastItem.ItTmp2,tblMastAccount_Mill.AcName,tblGpSub.VSubCtrNo,tblGpSub.VSubBag,tblGpSub.VSubWt,tblVouSub.VSubRt,tblGpSub.VSubGodownCode,tblVouSub.VSubLotno,tblVouSub.VSUBTMP3,tblVouSub.VSubTmp7,tblMastItem.ITType,tblMastItem.ItStdrateper,tblMastItem.ITTMP2,tblMastItem.ItTicketPrint,tblVouSub.VSubBookNo,tblVouSub.VSubBookDt,'" & gUserName & "' " _
            & " from tblAddLess,tblMastAccount,tblMastAccount tblMastAccount_Br,tblMastItem,tblGpSub,tblMastAccount tblMastAccount_Mill,tblVouSub " _
            & "where tblAddLess.AdAcCrCode = tblMastAccount.AcCode and tblAddLess.AdAcDrCode = tblMastAccount_Br.AcCode and tblGpSub.VSubItCode = tblMastItem.ItCode" _
            & " and tblAddLess.Vno = tblVouSub.Vno and tblAddLess.Vtype = tblVouSub.Vtype and tblAddLess.Vyear = tblVouSub.Vyear and tblAddLess.Vtype = 'GP' and tblAddLess.Vyear = " & gCYear & "" _
            & " and tblAddLess.Vno = tblGpSub.Vno and tblAddLess.Vtype = tblGpSub.Vtype and tblAddLess.Vyear = tblGpSub.Vyear and tblAddLess.Vtype = 'GP' and tblAddLess.Vyear = " & gCYear & " and " _
            & " tblAddLess.AdMillCode = tblMastAccount_Mill.AcCode and tblAddLess.Vno = " & GpNo & "  "
    If GProcRstOpen(RsGp, X, "R", gCn) > 0 Then
       
       gCn.BeginTrans
       gCn.Execute "Delete from tmpGpPrint where  UserName='" & gUserName & "'"
       gCn.Execute " insert into tmpGpPrint (Vno,Vdt,Side,Party,Add1,Add2,City,Gstin,Phone,Broker,Item,Ticket,ItHsn,Mill,GpSr,Bag,Wt,Rate,VSubGodownCode,Lotno,RateType,Note1,ItType,RatrPerKg,HsnCd,BookNo,BookingNo,BookingDt,UserName) " & X
       gCn.CommitTrans
        
       CryReport.Formulas(5) = "FirmGst = '" & gCGSTIN & "'"
       X = "{tmpGpPrint.UserName}='" & gUserName & "' "
       CryReport.SelectionFormula = X
       CryReport.ReportFileName = gReportPath & "rptGatePassPrint_Multi.rpt"
'       CryReport.Formulas(7) = "wpan = '" & gCPAN & "'"
'       CryReport.Formulas(8) = "wCompGstin = '" & gCGSTIN & "'"
       If Check1.Value = 1 Then
          CryReport.Formulas(9) = "wMnm = 'Y'"
       Else
           CryReport.Formulas(9) = "wMnm = 'N'"
       End If
       If ChkRt.Value = 1 Then
          CryReport.Formulas(11) = "wPrt = 'Y'"
       Else
          CryReport.Formulas(11) = "wPrt = 'N'"
       End If
'       If OptOrg.Value = True Then
'            CryReport.Formulas(10) = "wCopy = 'Original Copy'"
'       ElseIf OptDup.Value = True Then
'             CryReport.Formulas(10) = "wCopy = 'Duplicate Copy'"
'       End If
        CryReport.ReportFileName = gReportPath & "rptGatePassPrint_Multi.rpt"
        If Index <> 4 Then
           CryReport.Action = 1
        End If
        If Index = 4 Then
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
           Tattach = gReportPath & "rptGatePassPrint_Multi.rpt"
           Set objReport = objCrystal.OpenReport(Tattach, 1)
                           
           Dim crTab As CRAXDRT.DatabaseTable
           For Each crTab In objReport.Database.Tables
               crTab.Location = gDbLocation
           Next
                            
             
           objReport.RecordSelectionFormula = "{tmpGpPrint.UserName}='" & gUserName & "' "
                            
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
                            If OptOrg.Value = True Then
                               objReport.FormulaFields(j).text = "Trim(" & Chr(39) & "Original Copy" & Chr(39) & ")"
                            ElseIf OptDup.Value = True Then
                               objReport.FormulaFields(j).text = "Trim(" & Chr(39) & "Duplicate Copy" & Chr(39) & ")"
                            End If
               End Select
                           
           Next j
           Dim PtyMbl As String
           Dim BrkName As String
           Dim BrkMbl As String
           Dim BrkCd As Long
                             
           Dim FileNMWithPath As String
           Dim FileNM As String
           Dim aFile As String
           FileNM = "GatePass" & Format(Now, "ddMMyyyyhhmmss") & "" & ".pdf"
           FileNMWithPath = gReportPath & "" & FileNM & ""
           PtyMbl = RsGp.Fields(8) & vbNullString  ' GProcGetColumnValue("TblMastAccount", "AcName", RsGp.Fields(3), "N", "AcFaxNo", "S")
           PtyName = RsGp.Fields(3) ' GProcGetColumnValue("TblMastAccount", "AcCode", RsGp!AdAcCrCode, "N", "AcName", "S")
           If Index = 4 Then
              ExportReportToPDF objReport, FileNMWithPath, "foo"
              Gdelay (5)
              SentWhatsApp "PDF", FileNM, "", PtyMbl, PtyName
           End If
        
           gReportPath = App.Path & "\HIReports_Access\"
           aFile = gReportPath & "*.pdf"
           If Len(Dir$(aFile)) > 0 Then
              Kill aFile
           End If
        End If
     End If
        GpNo = GpNo + 1
    Loop
    X = "Select * from tmpWhatsAppErr where username = '" & gUserName & "'"
    If GProcRstOpen(RsTmp, X, "R", gCn) > 0 Then
       GProcWhatsAppError
    End If
        
        
        
        
        
        
        
        
        
        
'''        X = "Select * From tblAddLess where  vNo=" & GpNo & " and VType='GP' and VYear=" & gCYear & " order by vno"
'''        If GProcRstOpen(RsGp, X, "R") > 0 Then
''''            GpNo = RsGp!Vno
'''            With CryReport
'''                With RsGp
'''                     .MoveFirst
'''                     Do While Not .EOF
'''                        X = "Select tblAddLess.Vno,tblAddLess.Vdt,tblAddLess.ADGPNO,tblMastAccount.AcName,tblMastAccount.AcAdd1,tblMastAccount.AcAdd2,tblMastAccount.AcCity,tblMastAccount.AcGstin,tblMastAccount.AcFaxNo," _
'''                            & "tblMastAccount_Br.AcName,tblMastItem.ItName,tblMastItem.ItTicket,tblMastItem.ItTmp2,tblMastAccount_Mill.AcName,tblGpSub.VSubCtrNo,tblGpSub.VSubBag,tblGpSub.VSubWt,tblVouSub.VSubRt,tblGpSub.VSubGodownCode,tblVouSub.VSubLotno,tblVouSub.VSubTmp7,tblMastItem.ITType,tblMastItem.ItStdrateper,tblMastItem.ITTMP2,tblVouSub.VSubBookNo,'" & gUserName & "' " _
'''                            & " from tblAddLess,tblMastAccount,tblMastAccount tblMastAccount_Br,tblMastItem,tblGpSub,tblMastAccount tblMastAccount_Mill,tblVouSub " _
'''                            & "where tblAddLess.AdAcCrCode = tblMastAccount.AcCode and tblAddLess.AdAcDrCode = tblMastAccount_Br.AcCode and tblGpSub.VSubItCode = tblMastItem.ItCode" _
'''                            & " and tblAddLess.Vno = tblVouSub.Vno and tblAddLess.Vtype = tblVouSub.Vtype and tblAddLess.Vyear = tblVouSub.Vyear and tblAddLess.Vtype = 'GP' and tblAddLess.Vyear = " & gCYear & "" _
'''                            & " and tblAddLess.Vno = tblGpSub.Vno and tblAddLess.Vtype = tblGpSub.Vtype and tblAddLess.Vyear = tblGpSub.Vyear and tblAddLess.Vtype = 'GP' and tblAddLess.Vyear = " & gCYear & " and " _
'''                            & " tblAddLess.AdMillCode = tblMastAccount_Mill.AcCode and tblAddLess.Vno = " & RsGp!Vno & "  "
'''                        gCn.BeginTrans
'''                        gCn.Execute "Delete from tmpGpPrint where  UserName='" & gUserName & "'"
'''                        gCn.Execute " insert into tmpGpPrint (Vno,Vdt,Side,Party,Add1,Add2,City,Gstin,Phone,Broker,Item,Ticket,ItHsn,Mill,GpSr,Bag,Wt,Rate,VSubGodownCode,Lotno,Note1,ItType,RatrPerKg,HsnCd,BookNo,UserName) " & X
'''                        gCn.CommitTrans
'''                        CryReport.Formulas(5) = "FirmGst = '" & gCGSTIN & "'"
'''                        X = "{tmpGpPrint.UserName}='" & gUserName & "' "
'''                        CryReport.SelectionFormula = X
'''                        CryReport.ReportFileName = gReportPath & "rptGatePassPrint_Multi.rpt"
''''                        .SelectionFormula = "{tblAddLess.vNo} = " & RsGp!Vno & " and {tblAddLess.VType}='" & RsGp!Vtype & "' and {tblAddLess.VYear}=" & RsGp!VYear
'''       '                CryReport.Formulas(7) = "wpan = '" & gCPAN & "'"
'''       '                CryReport.Formulas(8) = "wCompGstin = '" & gCGSTIN & "'"
'''       '                If Check1.Value = 1 Then
'''        '                  CryReport.Formulas(9) = "wMnm = 'Y'"
'''        '               Else
'''        '                  CryReport.Formulas(9) = "wMnm = 'N'"
'''        '               End If
'''        '               If ChkRt.Value = 1 Then
'''        '                  CryReport.Formulas(11) = "wPrt = 'Y'"
'''        '               Else
'''        '                  CryReport.Formulas(11) = "wPrt = 'N'"
'''        '               End If
'''        '               If OptOrg.Value = True Then
'''        '                  CryReport.Formulas(10) = "wCopy = 'Original Copy'"
'''        '               ElseIf OptDup.Value = True Then
'''        '                  CryReport.Formulas(10) = "wCopy = 'Duplicate Copy'"
'''        '               End If
'''                       CryReport.ReportFileName = gReportPath & "rptGatePassPrint_Multi.rpt"
'''                       If Index <> 4 Then
'''                          CryReport.Action = 1
'''                       End If
'''                       If Index = 4 Then
'''                            Dim objCrystal As CRAXDRT.Application
'''                            Dim objReport As CRAXDRT.Report
'''                            Dim Tattach As String
'''                            Dim Mfile As String
'''                            Dim Mcp As String
'''                            Dim PtyName As String
'''                            Dim PtyMailId As String
'''                            Dim PtyMailIdCc As String
'''                            Dim PtyMailIdBcc As String
'''                            Dim MailSub As String
'''                            Dim MailStr As String
'''                            Set objCrystal = New CRAXDRT.Application
'''                            Tattach = gReportPath & "rptGatePassPrint.rpt"
'''                            Set objReport = objCrystal.OpenReport(Tattach, 1)
'''
'''                            Dim crTab As CRAXDRT.DatabaseTable
'''                            For Each crTab In objReport.Database.Tables
'''                                crTab.Location = gDbLocation
'''                            Next
'''
'''
'''                            objReport.RecordSelectionFormula = "{tblAddLess.vNo} = " & RsGp!Vno & " and {tblAddLess.VType}='" & RsGp!Vtype & "' and {tblAddLess.VYear}=" & RsGp!VYear
'''
'''                            Dim j As Integer
'''                            For j = 1 To objReport.FormulaFields.Count
'''                                Select Case objReport.FormulaFields(j).Name
'''                                Case "{@FirmNm}"
'''                                    objReport.FormulaFields(j).text = "Trim(" & Chr(39) & UCase(gCName) & Chr(39) & ")"
'''                                Case "{@wAdd1}"
'''                                    objReport.FormulaFields(j).text = "Trim(" & Chr(39) & gCAdd1 & Chr(39) & ")"
'''                                Case "{@wAdd2}"
'''                                    objReport.FormulaFields(j).text = "Trim(" & Chr(39) & gCAdd2 & Chr(39) & ")"
'''                                Case "{@wAdd3}"
'''                                    objReport.FormulaFields(j).text = "Trim(" & Chr(39) & gCAdd3 & Chr(39) & ")"
'''                                Case "{@wPhNo}"
'''                                    objReport.FormulaFields(j).text = "Trim(" & Chr(39) & gCPhNo & Chr(39) & ")"
'''                                Case "{@wCompGstin}"
'''                                    objReport.FormulaFields(j).text = "Trim(" & Chr(39) & gCGSTIN & Chr(39) & ")"
'''                                Case "{@Wcopy}"
'''                                    If OptOrg.Value = True Then
'''                                       objReport.FormulaFields(j).text = "Trim(" & Chr(39) & "Original Copy" & Chr(39) & ")"
'''                                    ElseIf OptDup.Value = True Then
'''                                       objReport.FormulaFields(j).text = "Trim(" & Chr(39) & "Duplicate Copy" & Chr(39) & ")"
'''                                    End If
'''                                End Select
'''
'''                            Next j
'''                            Dim PtyMbl As String
'''                            Dim BrkName As String
'''                            Dim BrkMbl As String
'''                            Dim BrkCd As Long
'''
'''                            Dim FileNMWithPath As String
'''                            Dim FileNM As String
'''                            Dim aFile As String
'''                            FileNM = "GatePass" & Format(Now, "ddMMyyyyhhmmss") & "" & ".pdf"
'''                            FileNMWithPath = gReportPath & "" & FileNM & ""
'''                            PtyMbl = GProcGetColumnValue("TblMastAccount", "AcCode", RsGp!AdAcCrCode, "N", "AcFaxNo", "S")
'''                            PtyName = GProcGetColumnValue("TblMastAccount", "AcCode", RsGp!AdAcCrCode, "N", "AcName", "S")
'''                            If Index = 4 Then
'''                               ExportReportToPDF objReport, FileNMWithPath, "foo"
'''                               Gdelay (5)
'''                               SentWhatsApp "PDF", FileNM, "", PtyMbl, PtyName
'''                            End If
'''
'''                             gReportPath = App.Path & "\HIReports_Access\"
'''                             aFile = gReportPath & "*.pdf"
'''                             If Len(Dir$(aFile)) > 0 Then
'''                                Kill aFile
'''                             End If
'''                       End If
'''                       RsGp.MoveNext
'''                     Loop
'''                End With
'''            End With
'''        End If
''''        GpNo = GpNo + 1
'''    Loop
'''    X = "Select * from tmpWhatsAppErr where username = '" & gUserName & "'"
'''    If GProcRstOpen(RsTmp, X, "R", gCn) > 0 Then
'''       GProcWhatsAppError
'''    End If
        
Case 3:
       Dim InvNo As String
       Dim LrString As String
       Dim SmsMblNo As String
       Dim SmsString As String
       Dim Pside As String
       Dim Rate As Integer
       Dim RsSA As Recordset
'       Dim PtyName As String
'       Dim RsTmp As Recordset
       gCn.BeginTrans
       gCn.Execute "Delete from tmpWhatsAppErr where username = '" & gUserName & "'"
       gCn.CommitTrans
       InvNo = txtno(0)
       X = "("
       Do While CDbl(InvNo) <= CDbl(txtno(1))
              X = X + "" & InvNo & ","
           InvNo = InvNo + 1
       Loop
       X = Left(X, Len(X) - 1) + ")"
       LrString = "Select tblAddLess.vNo,tblAddLess.Vdt,tblMastAccount.AcName,tblMastItem.Itname,tblMastItem.ItTicket,tblGpSub.VSubBag,tblGpSub.VSubWt,tblVouSub.VSubRt,VSubRtPer,DateAdd('d', tblAddLess.AdGpNo, tblAddLess.Vdt),tblMastAccount_Br.AcName,tblMastAccount.AcFaxNo,tblMastAccount_Br.AcFaxNo,tblAddLess.AdGpNo,tblMastItem.ItType,tblGpSub.VSubCtrNo " _
                   & "From tblAddLess,tblMastAccount,tblVouSub,tblMastItem,tblMastAccount tblMastAccount_Br,tblGpsub where tblAddLess.AdAcCrCode = tblMastAccount.AcCode and tblAddLess.Vno = tblVouSub.Vno and tblAddLess.Vtype = tblVouSub.Vtype and tblAddLess.Vyear = tblVouSub.Vyear and " _
                   & " tblAddLess.Vno = tblGpSub.Vno and tblAddLess.Vtype = tblGPSub.Vtype and tblAddLess.Vyear = tblGpSub.Vyear and " _
                   & " tblAddLess.AdAcDrCode = tblMastAccount_Br.AcCode  and  tblVousub.VSubItCode = tblMastItem.Itcode and  tblAddLess.VNo in " & X & " and tblAddLess.VType='GP' and tblAddLess.VYear=" & gCYear
     If GProcRstOpen(RsSA, LrString, "R") <= 0 Then
       MsgBox "No Records..."
       Exit Sub
     End If
     RsSA.MoveFirst
     Do While Not RsSA.EOF
        PtyName = RsSA.Fields(2)

        SmsMblNo = RsSA.Fields(11) & vbNullString
        SmsString = "GP.No-" + CStr(RsSA.Fields(0)) + "/" + CStr(RsSA.Fields(15)) + " Dt-" + Trim(CStr(Format(RsSA.Fields(1), "dd/mm/yy"))) & vbLf
        SmsString = SmsString + "Pty-" + Left(RsSA.Fields(2), 30) & vbLf
        SmsString = SmsString + "Cnt-" + Left(RsSA.Fields(3), 25) + "-" + Left(RsSA.Fields(4), 10) & vbLf
        SmsString = SmsString + "Bag-" + CStr(RsSA.Fields(5)) + " WT-" + CStr(RsSA.Fields(6)) + " Kg" & vbLf
        If RsSA.Fields(13) = 30 Then
           Pside = "Regular"
        ElseIf RsSA.Fields(13) = 7 Then
           Pside = "Second Day"
        ElseIf RsSA.Fields(13) = 1 Then
           Pside = "Next Day"
        End If
        If RsSA.Fields(14) = 1 Then
           Rate = RsSA.Fields(7) * 1.12
        Else
           Rate = RsSA.Fields(7) * 1.05
        End If
        SmsString = SmsString + "Without GST Rate-" + CStr(RsSA.Fields(7)) + " Per " + CStr(Int(RsSA.Fields(8))) + " Kg" & vbLf
        SmsString = SmsString + "With GST Rate-" + CStr(Rate) + " Per " + CStr(Int(RsSA.Fields(8))) + " Kg" & vbLf
        SmsString = SmsString + "Payment Condition-" + Pside & vbLf
        SmsString = SmsString + "Due-" + CStr(Format(RsSA.Fields(9), "dd/mm/yy")) & vbLf
        SmsString = SmsString + "Brk-" + Left(RsSA.Fields(10), 25) & vbLf
        SmsString = SmsString + "Firm-" + gCName
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
            SentWhatsApp "Txt", "", SmsString, SmsMblNo, PtyName
        End If
        DoEvents
        DoEvents
        DoEvents
        PtyName = RsSA.Fields(10)
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
            SentWhatsApp "Txt", "", SmsString, SmsMblNo, PtyName
        End If
        DoEvents
        DoEvents
        DoEvents
        RsSA.MoveNext
     Loop
     X = "Select * from tmpWhatsAppErr where username = '" & gUserName & "'"
     If GProcRstOpen(RsTmp, X, "R", gCn) > 0 Then
        GProcWhatsAppError
     End If
     MsgBox ("All SMS Sent....")
Case 2:
    DoEvents
    Frame1.Visible = False
End Select
End Sub
'------- Cartoon Bag Details
'Private Sub CmdBagDetails_Click()
'FrameGrid.Visible = True
'msGrid.Enabled = False
'If FormAction = vbDataActionAddNew Or FormAction = vbDataActionUpdate Then
'    msGrid.Enabled = True
'    msGrid.SetFocus
'End If
'FillRelatedRecord
'End Sub
'-- Exit from Grid Frame
'Private Sub cmdExitFrameGrid_Click(Index As Integer)
'FrameGrid.Visible = False
'End Sub
Private Sub SetGrid()
Dim X As String
Dim RCount As Long
If IsNumeric(txttmpInNo) = False Then
    txttmpInNo = 0
End If
'--BagNo,cone,Wt,InwNo,InwType,GpNo,GpType
X = "select VSubCtrNo,VSubBag,VSubWt from tblGpSub where " _
& " tblGpSub.vtype='GP' and tblGpSub.vno=" & txt(0) & " and tblGpSub.VYear=" & gCYear & " order by VSubCtrNo"
Set rstblSub = New Recordset
RCount = GProcRstOpen(rstblSub, X, "R")
GProcRstOpen rstblSub, X, "R"
With msGrid
    .Clear
    .FormatString = ">Sr. No.|>Bags       |>Weight       "
    If FormAction = vbDataActionAddNew Or RCount = 0 Then
        .Rows = 2
        GridAddNew
    Else
        .Rows = 1
    End If
End With
End Sub
'======================== Grid code
'-- Fill Grid
Private Sub FillRelatedRecord()
Dim RowIndex As Integer
Dim ColIndex As Integer
SetGrid
'-- Fill Record
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
            Next ColIndex
            RowIndex = RowIndex + 1
            .MoveNext
            If Not .EOF Then msGrid.Rows = msGrid.Rows + 1
        Loop
    End If
End With
End Sub
Private Sub msGrid_Click()
'On Error Resume Next
    msGrid_EnterCell
End Sub
Private Sub msGrid_GotFocus()
 cmdBtn(10).Cancel = False
End Sub
Private Sub msGrid_EnterCell()
'If FormAction <> vbDataActionClose Then
    CheckForNumberTxtGrid
    DoEvents
    txtGrid = ""
    CalGridAmt
    If Trim(msGrid.text) <> "" Then txtGrid.text = msGrid.text
    '--- Lock TextBox
    With msGrid
        Select Case .Col
            Case 0
                txtGrid.Locked = True
                ReNumberGridSrno
                txtGrid = ""
                .Col = .Col + 1
            Case Else
                txtGrid.Locked = False
        End Select
    End With
    msGrid_KeyPress (0)
    CalGridAmt
'End If
End Sub
Private Sub msGrid_KeyPress(KeyAscii As Integer)
With msGrid
    If KeyAscii = 27 Then '----Esc Key
        KeyAscii = 0
        txtGrid.Visible = False
        msGrid.Visible = False '-Grid Visible -False
        Exit Sub
    End If
End With
    '---- Navigate Row ,Col
    If KeyAscii = 13 Then
        KeyAscii = 0
      With msGrid
        If .Col = 3 Then
            If .Row = (.Rows - 1) Then '-- Add New Row
            Else '--Next Row
                .Row = .Row + 1
            End If
            .Col = 1
        Else    '-- Next col
            .Col = .Col + 1
        End If
        End With
    End If
'    If KeyAscii <> 13 And KeyAscii <> 27 And .Col = 0 Then
'        DoEvents
'        KeyAscii = 0
'        Pending_In
'    End If
'ProcSetAlignment rstblSub, msGrid.Col
'---- Set TxtGrid
If msGrid.text <> "" Then txtGrid.text = msGrid.text
If msGrid.Col < 3 Then
    txtGrid.Move msGrid.CellLeft + msGrid.Left, msGrid.CellTop + msGrid.Top, msGrid.CellWidth, msGrid.CellHeight
End If
txtGrid.Visible = True
txtGrid.Enabled = True
DoEvents
If KeyAscii <> 0 Then
    Select Case rstblSub.Fields(msGrid.Col).Type
        Case adChar, adVarWChar
            If msGrid.Col = 0 Then '-Inward Cartton Bag
                txtGrid_KeyPress (KeyAscii)
            Else
                txtGrid.text = txtGrid.text & Chr(KeyAscii)
            End If
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
If txtGrid.Visible = True Then txtGrid.SetFocus
End Sub
Private Sub msGrid_LeaveCell()
If txtGrid.Visible Then
    msGrid.text = txtGrid.text
    DoEvents
    txtGrid.text = ""
    txtGrid.Visible = False
    CalGridAmt
    CalAmount
End If
End Sub

Private Sub MSGrid_LostFocus()
CalAmount
End Sub
Private Sub txtGrid_GotFocus()
    txtGrid.Alignment = ProcSetAlignment(rstblSub, msGrid.Col)
    DoEvents
    Call GProcSelectBox(txtGrid)
End Sub
Private Sub txtGrid_KeyPress(KeyAscii As Integer)
Dim LeftPos As Long
Dim TopPos As Long
TopPos = txtGrid.Top + Me.Top + 650
LeftPos = txtGrid.Left + Me.Left
DoEvents
With msGrid
    If KeyAscii = 27 Then '----Esc Key
        KeyAscii = 0
        If (Val(.TextMatrix(.Row, 1)) = 0) And .Rows > 2 Then
            .RemoveItem (.Row)
            txtGrid.text = ""
        End If
        txtGrid.Visible = False
        txt(30).SetFocus
        Exit Sub
    End If
End With
With msGrid
    If KeyAscii <> 13 And KeyAscii <> 27 Then
        '--- Search
        Select Case .Col
            Case 0 '-BagNo
                KeyAscii = 0
'                GetInwardCartoonBags
            Case Else
                KeyAscii = GProcValidateKey(rstblSub, msGrid.Col, KeyAscii, txtGrid)
        End Select
   End If
    If KeyAscii = 13 Then
  '      CheckForNumberTxtGrid
    End If
    '--- Navigate Row ,Col
    If KeyAscii = 13 And .Col < 2 Then
        .Col = .Col + 1
    ElseIf KeyAscii = 13 And .Col = 2 Then
        If .Row = (.Rows - 1) Then
            .Rows = .Rows + 1
            .Row = .Row + 1
            GridAddNew
        Else
            .Row = .Row + 1
        End If
        .Col = 1
    End If
    '-- On ESC
    With msGrid
        If KeyAscii = 27 Then '----Esc Key
            If (Val(.TextMatrix(.Row, 1)) = 0) And .Rows > 2 Then
                .RemoveItem (.Row)
                txtGrid.text = ""
            End If
            txtGrid.Visible = False
            txt(30).SetFocus
            KeyAscii = 0
            msGrid.Visible = False '-Grid Visible -False
            Exit Sub
        End If
    End With
End With
End Sub
Private Sub CheckForNumberGrid()
Dim RowIndex As Long
'On Error Resume Next
With rstblSub
For RowIndex = 1 To msGrid.Rows - 1
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
    '    Case adChar, adVarWChar '--String
   '         txtGrid.MaxLength = .Fields(msGrid.Col).DefinedSize
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
Private Function ProcSetAlignment(prstblSub As Recordset, pColIndex As Long)
Select Case prstblSub.Fields(pColIndex).Type
    Case adInteger, adDouble, 2, 3, 131, adDecimal
        ProcSetAlignment = 1
    Case Else
        ProcSetAlignment = 0
End Select
End Function
'-- Grid Required
Private Function ValidateGrid() As Boolean
Dim i As Long
Dim RowIndex As Long
Dim TotBag As Double
Dim TotWt As Double

If Trim(txtGrid.text) <> "" Then
    msGrid.text = txtGrid.text
    CalGridAmt
End If
txtGrid.Visible = False
txtGrid.text = ""
CheckForNumberGrid




'If txtGrid.Visible = True Then msGrid.text = txtGrid.text
'txtGrid.Visible = False
'txtGrid.text = ""
'CheckForNumberGrid
With msGrid
Validate:
'On Error GoTo nobagno
    For RowIndex = 1 To .Rows - 1
        For i = 0 To rstblSub.Fields.Count - 1
            Select Case i
                Case 1 '-- Bag
                    If RowIndex > 0 And Trim(.TextMatrix(RowIndex, i)) = "" Then
                        .RemoveItem (RowIndex)
                        GoTo Validate
                    End If
                    If Val(.TextMatrix(RowIndex, i)) > 0 Then
                       TotBag = TotBag + Val(.TextMatrix(RowIndex, i))
                    End If
                Case 2 '-- wt
                    If Val(.TextMatrix(RowIndex, i)) > 0 Then
                       TotWt = TotWt + Val(.TextMatrix(RowIndex, i))
                    End If
           
           End Select
        Next i
    Next RowIndex
    '--Bag Row More than GP Bags
    If TotBag <> Val(txt(5)) Then
       MsgBox "Check Total Bags "
'       FrameGrid.Visible = True
       msGrid.SetFocus
       Exit Function
    End If
    If Round(TotWt, 3) <> Round(Val(txt(6)), 3) Then
       MsgBox "Check total Weight "
'       FrameGrid.Visible = True
       msGrid.SetFocus
       Exit Function
    End If

End With
ValidateGrid = True
ReNumberGridSrno
End Function
Private Sub txttmpInNo_LostFocus()
SetControlEd
End Sub
'---------- Party Booking Vs Sale
Private Sub PrepareBooking()
Dim X1 As String
Dim X2 As String
Dim QStr As String
Dim SelePtyGrp As Long
Dim RsTmp As Recordset
Dim MillCode As Long
Dim ItCd As Long
Dim PtyCd As Long
Dim BrkCd As Long
Dim ItNm As String
SelePtyGrp = GProcGetColumnValue("TblMastAccount", "AcName", txttempAL(3), "C", "AcCmpCode", "N")  '-Party Group Code
MillCode = GProcGetColumnValue("tblMastAccount", "AcName", txttempAL(5), "S", "AcCode", "N")
ItCd = GProcGetColumnValue("tblMastItem", "ItName", txttemp(4), "S", "ItCode", "N", " ItMillcode=" & MillCode) '--ItCode
PtyCd = GProcGetColumnValue("tblMastAccount", "AcName", txttempAL(3), "S", "AcCode", "N")
BrkCd = GProcGetColumnValue("tblMastAccount", "AcName", txttempAL(4), "S", "AcCode", "N")
If SelePtyGrp > 1 Then
   GProcRstOpen RsTmp, "select accode from tblMastAccount where AcCmpCode = " & SelePtyGrp & "", "O", gCn
   With RsTmp
        RsTmp.MoveFirst
        X2 = "("
        Do While Not .EOF  '----Temparary Connection
           X2 = X2 + "" & RsTmp.Fields(0) & ","
           .MoveNext
        Loop
        X2 = Left(X2, Len(X2) - 1) + ")"
   End With
Else
   X2 = "(" + "" & PtyCd & "" + ")"
End If

If IsDate(mskDt(2)) = True Then
'   X1 = "select VNo,format(vdt,'dd/mm/yyyy'),Itname,VSubItCtrlNo,VSubBag-VSubIssBag,format(VSubWt-VSubIssWt,'########0.000'),VSubBag,VSubWt,VSubRt " _
'        & " From tblVouSub_Book,tblMastItem where vtype ='OP' and " _
'        & " VSubBookMillCode = " & MillCode & " and tblVouSub_book.vdt<=cdate('" & CDate(mskDt(2)) & "')  and   ((VSubBag-VSubIssBag)> 0) and " _
'        & " ( VSubBookAcCode in " & X2 & " or  VSubBookBrkAcCode = " & BrkCd & ") and VSubBookIsCompleted = 0 and tblVouSub_Book.VSubItCode = tblMastItem.Itcode"

   X1 = "select VNo,format(vdt,'dd/mm/yyyy'),tblMastAccount.AcName,Itname,tblMastAccount_Ml.AcName,VSubItCtrlNo,VSubBag-VSubIssBag,format(VSubWt-VSubIssWt,'########0.000'),VSubBag,VSubWt,VSubRt " _
        & " From tblVouSub_Book,tblMastItem,tblMastAccount,tblMastAccount tblMastAccount_Ml where vtype ='OP' and " _
        & " tblVouSub_book.vdt<=cdate('" & CDate(mskDt(2)) & "')  and   ((VSubBag-VSubIssBag)> 0) and " _
        & " ( VSubBookAcCode in " & X2 & " or  VSubBookBrkAcCode = " & BrkCd & ") and VSubBookIsCompleted = 0 and tblVouSub_Book.VSubItCode = tblMastItem.Itcode " _
        & " and VSubBookAcCode = tblMastAccount.AcCode and VSubBookMillCode = tblMastAccount_Ml.accode"


      With gClsSearch
           .SearchMultiField X1, "", Array("Bk.No.", "Date", "Party", "Count", "Mill", "ItCtrlNo", "Bal Bag", "Bal Wt.", "Bag", "Wt", "Bk.Rate"), Array(800, 1500, 2500, 2500, 2500, 0, 1000, 2000, 1000, 2000, 2000), "", "", "Vdt,vno", 0, 0, False, gCn, True
            If .SearchMultiRetCol(0) <> "" Then
               txt(42) = CLng(.SearchMultiRetCol(0)) '--Book Vno
               mskDt(43) = Format(.SearchMultiRetCol(1), "dd/mm/yyyy") '-- Book Vdt
               txt(44) = CLng(.SearchMultiRetCol(5)) '--Book ItCtrlNo
               txt(7) = CDbl(.SearchMultiRetCol(10)) '--Book Rate
'               txtAL(6) = CDbl(.SearchMultiRetCol(9)) '--Side
            End If
        End With
End If
txttmpInNo.SetFocus
'Dim X1 As String
'Dim X2 As String
'Dim QStr As String
'If IsDate(mskDt(2)) = True Then
'     If FormAction = vbDataActionAddNew Then
'        X1 = "Bksubbag-(Bksubissbag),Bksubwt-(Bksubisswt)"
'        X2 = "((Bksubbag-(Bksubissbag))> 0)"
'    Else
'        X1 = "Bksubbag-(Bksubissbag)+BkSubTmpBag,Bksubwt-(Bksubisswt) + BkSubTmpWt"
'        X2 = "((Bksubbag-(Bksubissbag)+BkSubTmpBag)> 0)"
'    End If
'    ProcInsertCodeNameInTmpTbl '--- Insert Code & Name In Setup TmpTbl As Per  Comp code,Year
'    QStr = " select tblBookingSub.VNo,tblBookingSub.vdt,Bksubitctrlno,Bksubbag-(Bksubissbag),Bksubwt-(Bksubisswt),Bksubbag,Bksubwt,BkCCode from " _
'    & " tblBookingSub,tblBooking,TmpGenTbl Where TmpGenTbl.UserName='" & gUserName & "' and " _
'    & " tblBookingSub.vno=tblBooking.vno and tblBookingSub.vType=tblBooking.vType and tblBookingSub.vYear=tblBooking.vYear and " _
'    & " BkAcBrkCode=Amt1 and BKAcCode=Amt2 and BkSubItCode=amt3 and " & X2 & " and tblBooking.vtype ='OP' " _
'    & " and BkSubIsCompleted=0 and Nar1='" & txttempAL(4) & "' and ( Nar2='" & txttempAL(3) & "' or Nar2 ='') and Nar3='" & txttemp(4) & "' and Nar4='" & txttempAL(5) & "'"
'    'If gBackEndDB = gBackEndAccess Then
'        QStr = QStr + " and tblBooking.vdt<=cdate('" & CDate(mskDt(2)) & "')"
'    'Else '-- Oracle
'    '    QStr = QStr + " and tblBooking.vdt <= to_date('" & CDate(mskDt(2)) & "','dd/MM/yyyy')"
'    'End If
'    With gClsSearch
'        .SearchMultiField QStr, "", Array("Booking No.", "Date", "ItCtrlNo", "Bal Bag", "Bal Wt.", "Bag", "Wt", "C. Code"), Array(1200, 1200, 0, 1000, 1000, 1000, 1000, 1000), "", "", "tblBookingsub.Vdt,tblBookingSub.vno", 0, 0, False, gSetupCn, True
'        If .SearchMultiRetCol(0) <> "" Then
'                txt(42) = CLng(.SearchMultiRetCol(0)) '--Book Vno
'                mskDt(43) = Format(.SearchMultiRetCol(1), "dd/mm/yyyy") '-- Book Vdt
'                txt(44) = CLng(.SearchMultiRetCol(2)) '--Book ItCtrlNo
'        End If
'    End With
'End If
End Sub
'--- Procedure Insert Code & Name In Setup TmpTbl As Per  Comp code,Year
Private Sub ProcInsertCodeNameInTmpTbl()
Dim TmpCn As Connection
Dim RsTmp As Recordset
Dim BrkName As String '- Broker
Dim PartyName As String '-Party
Dim ItName As String '-Item
Dim MillCode As Long '-MillCode
Dim MillName As String '--Mill
gSetupCn.Execute "delete from tmpGenTbl where UserName='" & gUserName & "'"
'--- Insert Code & their Name In Setup TmpTbl As Per  Comp code,Year
'- amt1=BrCode,Nar1=BrkName,amt2=PartyCode,Nar2=PartyName,Amt3=ItCode,Nar3=ItName,Amt4=MillCode,Nar4=MillName
'- Nar5=CCode,Nar6=CCyear
 gSetupCn.Execute "insert into tmpGenTbl (Nar5,Nar6,amt1,Amt2,Amt3,UserName) Select BkCCode,BkCYear,BkAcBrkCode," _
 & " BKAcCode,BkSubItCode,'" & gUserName & "' from tblbooking,tblbookingSub where BkSubIsCompleted=0 and " _
 & " tblbooking.vno=tblbookingSub.vno and tblbooking.vtype=tblbookingSub.Vtype and tblbooking.VYear=tblbookingSub.VYear" _
 & "  group by BkCCode,BkCYear,BkAcBrkCode,BKAcCode,BkSubItCode"
 GProcRstOpen RsTmp, "Select * from tmpGenTbl where Username='" & gUserName & "'", "O", gSetupCn
 With RsTmp
    Do While Not .EOF  '----Temparary Connection
        If gBackEndDB = gBackEndAccess Then '-- Access
            GProcConnectionOpen TmpCn, !Nar5 + !Nar6
        Else
            GProcConnectionOpenORA TmpCn, !Nar5 + !Nar6
        End If
        BrkName = GProcGetColumnValue("TblMastAccount", "ACCode", !Amt1, "N", "AcName", "S", "", TmpCn) '-Broker
        PartyName = GProcGetColumnValue("TblMastAccount", "ACCode", !AMT2, "N", "AcName", "S", "", TmpCn) '-Party
        ItName = GProcGetColumnValue("TblMastItem", "ItCode", !AMT3, "N", "ItName", "S", "", TmpCn) '-Item
        MillCode = GProcGetColumnValue("TblMastAccount,TblMastItem", "ItCode", !AMT3, "N", "AcCode", "N", "AcCode=ITMILLCODE", TmpCn)   '-Mill Code
        MillName = GProcGetColumnValue("TblMastAccount", "ACCode", CStr(MillCode), "N", "AcName", "S", "", TmpCn)   '- mill
        !Nar1 = BrkName
        !Nar2 = PartyName
        !Nar3 = ItName
        !Nar4 = MillName
        !Amt4 = MillCode
        .Update
        .MoveNext
    Loop
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

'-- ReNumber Grid Sr No
Private Sub ReNumberGridSrno()
Dim i As Long
With msGrid
    For i = 1 To .Rows - 1
        .TextMatrix(i, 0) = i
    Next i
End With
End Sub
Private Sub CalGridAmt()
Dim MillCode As Long
Dim StdPack As Double
CheckForNumberGrid
If Val(msGrid.TextMatrix(msGrid.Row, 1)) = 0 Then msGrid.TextMatrix(msGrid.Row, 2) = 0
MillCode = GProcGetColumnValue("TblMastAccount", "AcName", txttempAL(5), "S", "Accode", "N")
StdPack = GProcGetColumnValue("TblMastItem", "ITName", txttemp(4), "S", "ItStdpack", "N", " ItMillCode=" & MillCode)
If msGrid.Row = 0 Then Exit Sub
If msGrid.TextMatrix(msGrid.Row, 1) = 0 And msGrid.Row = 1 Then
   msGrid.TextMatrix(msGrid.Row, 1) = Val(txt(5))
   msGrid.TextMatrix(msGrid.Row, 2) = Val(txt(6))
End If


'-- Wt=Bag * Std Pack
If msGrid.TextMatrix(msGrid.Row, 2) = 0 Or (Abs(msGrid.TextMatrix(msGrid.Row, 2) - (msGrid.TextMatrix(msGrid.Row, 1) * StdPack)) > StdPack) Then
'If StdPack > 0 And (Abs(msGrid.TextMatrix(msGrid.Row, 2) - (msGrid.TextMatrix(msGrid.Row, 1) * StdPack)) > StdPack) Then
    msGrid.TextMatrix(msGrid.Row, 2) = msGrid.TextMatrix(msGrid.Row, 1) * StdPack
End If
End Sub


