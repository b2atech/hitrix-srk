VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form frmSetting 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00FBF2E1&
   BorderStyle     =   0  'None
   Caption         =   "Setting Entry"
   ClientHeight    =   10770
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   13365
   Icon            =   "frmSetting.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   10770
   ScaleWidth      =   13365
   ShowInTaskbar   =   0   'False
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
      Index           =   73
      Left            =   15840
      TabIndex        =   174
      Text            =   "73"
      Top             =   8460
      Visible         =   0   'False
      Width           =   420
   End
   Begin VB.TextBox txttemp 
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
      Index           =   73
      Left            =   2160
      MaxLength       =   40
      TabIndex        =   173
      Text            =   "txttemp(73)"
      ToolTipText     =   "Press F3 for Creation"
      Top             =   8400
      Width           =   4485
   End
   Begin VB.TextBox txttemp 
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
      Index           =   72
      Left            =   2190
      MaxLength       =   40
      TabIndex        =   171
      Text            =   "txttemp(72)"
      ToolTipText     =   "Press F3 for Creation"
      Top             =   7800
      Width           =   4485
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
      Index           =   72
      Left            =   15300
      TabIndex        =   170
      Text            =   "72"
      Top             =   8400
      Visible         =   0   'False
      Width           =   420
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
      Index           =   71
      Left            =   2010
      MaxLength       =   40
      TabIndex        =   26
      Text            =   "txttemp(71)"
      ToolTipText     =   "Press F3 for Creation"
      Top             =   960
      Width           =   4485
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
      Index           =   70
      Left            =   2010
      MaxLength       =   40
      TabIndex        =   28
      Text            =   "txttemp(70)"
      ToolTipText     =   "Press F3 for Creation"
      Top             =   1860
      Width           =   4485
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
      Index           =   71
      Left            =   16020
      TabIndex        =   169
      Text            =   "71"
      Top             =   6660
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
      Index           =   70
      Left            =   16020
      TabIndex        =   168
      Text            =   "70"
      Top             =   6300
      Visible         =   0   'False
      Width           =   420
   End
   Begin VB.TextBox txttemp 
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
      Index           =   69
      Left            =   8640
      MaxLength       =   40
      TabIndex        =   55
      Text            =   "txttemp(69)"
      ToolTipText     =   "Press F3 for Creation"
      Top             =   7020
      Width           =   4485
   End
   Begin VB.TextBox txttemp 
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
      Index           =   68
      Left            =   8640
      MaxLength       =   40
      TabIndex        =   54
      Text            =   "txttemp(68)"
      ToolTipText     =   "Press F3 for Creation"
      Top             =   6660
      Width           =   4485
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
      Index           =   69
      Left            =   15960
      TabIndex        =   165
      Text            =   "69"
      Top             =   6000
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
      Index           =   68
      Left            =   16080
      TabIndex        =   164
      Text            =   "68"
      Top             =   5640
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
      Index           =   67
      Left            =   16020
      TabIndex        =   160
      Text            =   "67"
      Top             =   5220
      Visible         =   0   'False
      Width           =   420
   End
   Begin VB.TextBox txttemp 
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
      Index           =   67
      Left            =   2205
      MaxLength       =   40
      TabIndex        =   46
      Text            =   "txttemp(67)"
      ToolTipText     =   "Press F3 for Creation"
      Top             =   5850
      Width           =   4485
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
      Index           =   66
      Left            =   16065
      TabIndex        =   154
      Text            =   "66"
      Top             =   4815
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
      Index           =   65
      Left            =   16065
      TabIndex        =   153
      Text            =   "65"
      Top             =   4410
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
      Index           =   64
      Left            =   16065
      TabIndex        =   152
      Text            =   "64"
      Top             =   3915
      Visible         =   0   'False
      Width           =   420
   End
   Begin VB.TextBox txttemp 
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
      Index           =   66
      Left            =   10170
      MaxLength       =   40
      TabIndex        =   151
      Text            =   "txttemp(66)"
      ToolTipText     =   "Press F3 for Creation"
      Top             =   1710
      Width           =   3135
   End
   Begin VB.TextBox txttemp 
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
      Index           =   65
      Left            =   10170
      MaxLength       =   40
      TabIndex        =   150
      Text            =   "txttemp(65)"
      ToolTipText     =   "Press F3 for Creation"
      Top             =   900
      Width           =   3135
   End
   Begin VB.TextBox txttemp 
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
      Index           =   64
      Left            =   10170
      MaxLength       =   40
      TabIndex        =   149
      Text            =   "txttemp(64)"
      ToolTipText     =   "Press F3 for Creation"
      Top             =   1350
      Width           =   3135
   End
   Begin VB.TextBox txttemp 
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
      Index           =   63
      Left            =   8550
      MaxLength       =   40
      TabIndex        =   41
      Text            =   "txttemp(63)"
      ToolTipText     =   "Press F3 for Creation"
      Top             =   3960
      Width           =   2250
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
      Index           =   63
      Left            =   16080
      TabIndex        =   147
      Text            =   "63"
      Top             =   3480
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
      Index           =   62
      Left            =   15345
      TabIndex        =   146
      Text            =   "62"
      Top             =   8010
      Visible         =   0   'False
      Width           =   420
   End
   Begin VB.TextBox txttemp 
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
      Index           =   62
      Left            =   2025
      MaxLength       =   40
      TabIndex        =   145
      Text            =   "txttemp(62)"
      ToolTipText     =   "Press F3 for Creation"
      Top             =   3375
      Width           =   4485
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
      Index           =   61
      Left            =   16065
      TabIndex        =   142
      Text            =   "61"
      Top             =   3060
      Visible         =   0   'False
      Width           =   420
   End
   Begin VB.TextBox txttemp 
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
      Index           =   61
      Left            =   2190
      MaxLength       =   40
      TabIndex        =   44
      Text            =   "txttemp(61)"
      ToolTipText     =   "Press F3 for Creation"
      Top             =   5445
      Width           =   4485
   End
   Begin VB.TextBox txttemp 
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
      Index           =   60
      Left            =   8550
      MaxLength       =   40
      TabIndex        =   42
      Text            =   "txttemp(60)"
      ToolTipText     =   "Press F3 for Creation"
      Top             =   4410
      Width           =   2250
   End
   Begin VB.TextBox txttemp 
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
      Index           =   59
      Left            =   8550
      MaxLength       =   40
      TabIndex        =   43
      Text            =   "txttemp(59)"
      ToolTipText     =   "Press F3 for Creation"
      Top             =   4770
      Width           =   2250
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
      Index           =   60
      Left            =   16110
      TabIndex        =   140
      Text            =   "60"
      Top             =   2610
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
      Index           =   59
      Left            =   16065
      TabIndex        =   139
      Text            =   "59"
      Top             =   2115
      Visible         =   0   'False
      Width           =   420
   End
   Begin VB.TextBox txttemp 
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
      Index           =   58
      Left            =   2025
      MaxLength       =   40
      TabIndex        =   40
      Text            =   "txttemp(58)"
      ToolTipText     =   "Press F3 for Creation"
      Top             =   2970
      Width           =   4485
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
      Index           =   58
      Left            =   16020
      TabIndex        =   138
      Text            =   "58"
      Top             =   1665
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
      Index           =   57
      Left            =   12105
      MaxLength       =   40
      TabIndex        =   63
      Text            =   "57"
      ToolTipText     =   "Press F3 for Creation"
      Top             =   9720
      Width           =   780
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
      Index           =   56
      Left            =   12105
      MaxLength       =   40
      TabIndex        =   62
      Text            =   "56"
      ToolTipText     =   "Press F3 for Creation"
      Top             =   9315
      Width           =   780
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
      Index           =   55
      Left            =   12105
      MaxLength       =   40
      TabIndex        =   60
      Text            =   "55"
      ToolTipText     =   "Press F3 for Creation"
      Top             =   8145
      Width           =   780
   End
   Begin VB.TextBox txttemp 
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
      Index           =   54
      Left            =   2025
      MaxLength       =   40
      TabIndex        =   39
      Text            =   "txttemp(54)"
      ToolTipText     =   "Press F3 for Creation"
      Top             =   2565
      Width           =   4485
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
      Index           =   54
      Left            =   20115
      MaxLength       =   40
      TabIndex        =   130
      Text            =   "54"
      ToolTipText     =   "Press F3 for Creation"
      Top             =   9450
      Visible         =   0   'False
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
      Height          =   360
      Index           =   52
      Left            =   20115
      TabIndex        =   128
      Text            =   "52"
      Top             =   1710
      Visible         =   0   'False
      Width           =   420
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
      Index           =   49
      Left            =   10920
      MaxLength       =   40
      TabIndex        =   47
      Text            =   "txttemp(49)"
      ToolTipText     =   "Press F3 for Creation"
      Top             =   4755
      Width           =   2250
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
      Index           =   52
      Left            =   14040
      MaxLength       =   40
      TabIndex        =   30
      Text            =   "txttemp(52)"
      ToolTipText     =   "Press F3 for Creation"
      Top             =   480
      Visible         =   0   'False
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
      Height          =   360
      Index           =   53
      Left            =   14535
      MaxLength       =   40
      TabIndex        =   24
      Text            =   "txt(53)"
      ToolTipText     =   "Press F3 for Creation"
      Top             =   5130
      Visible         =   0   'False
      Width           =   600
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
      Index           =   51
      Left            =   20115
      TabIndex        =   125
      Text            =   "51"
      Top             =   1305
      Visible         =   0   'False
      Width           =   420
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
      Index           =   51
      Left            =   7920
      MaxLength       =   40
      TabIndex        =   23
      Text            =   "txttemp(51)"
      ToolTipText     =   "Press F3 for Creation"
      Top             =   9750
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
      Index           =   50
      Left            =   14445
      MaxLength       =   40
      TabIndex        =   14
      Text            =   "txt(50)"
      ToolTipText     =   "Press F3 for Creation"
      Top             =   3150
      Visible         =   0   'False
      Width           =   645
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
      Index           =   49
      Left            =   20115
      TabIndex        =   124
      Text            =   "49"
      Top             =   900
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
      Index           =   48
      Left            =   20070
      TabIndex        =   123
      Text            =   "48"
      Top             =   450
      Visible         =   0   'False
      Width           =   420
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
      Left            =   8370
      MaxLength       =   40
      TabIndex        =   13
      Text            =   "txttemp(48)"
      ToolTipText     =   "Press F3 for Creation"
      Top             =   10005
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
      Index           =   47
      Left            =   12105
      MaxLength       =   40
      TabIndex        =   64
      Text            =   "47"
      ToolTipText     =   "Press F3 for Creation"
      Top             =   10215
      Width           =   780
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
      Index           =   46
      Left            =   12105
      MaxLength       =   40
      TabIndex        =   61
      Text            =   "46"
      ToolTipText     =   "Press F3 for Creation"
      Top             =   8505
      Width           =   780
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
      Left            =   12105
      MaxLength       =   40
      TabIndex        =   59
      Text            =   "45"
      ToolTipText     =   "Press F3 for Creation"
      Top             =   7785
      Width           =   780
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
      Left            =   13680
      MaxLength       =   40
      TabIndex        =   118
      Text            =   "txt(44)"
      ToolTipText     =   "Press F3 for Creation"
      Top             =   2430
      Visible         =   0   'False
      Width           =   600
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
      Left            =   13680
      MaxLength       =   40
      TabIndex        =   117
      Text            =   "txt(43)"
      ToolTipText     =   "Press F3 for Creation"
      Top             =   1665
      Visible         =   0   'False
      Width           =   600
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
      Left            =   13950
      MaxLength       =   40
      TabIndex        =   116
      Text            =   "txt(42)"
      ToolTipText     =   "Press F3 for Creation"
      Top             =   1275
      Visible         =   0   'False
      Width           =   645
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
      Left            =   13950
      MaxLength       =   40
      TabIndex        =   115
      Text            =   "txt(41)"
      ToolTipText     =   "Press F3 for Creation"
      Top             =   915
      Visible         =   0   'False
      Width           =   645
   End
   Begin VB.TextBox txttemp 
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
      Index           =   40
      Left            =   6870
      MaxLength       =   40
      TabIndex        =   38
      Text            =   "txttemp(40)"
      ToolTipText     =   "Press F3 for Creation"
      Top             =   3345
      Width           =   4485
   End
   Begin VB.TextBox txttemp 
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
      Index           =   39
      Left            =   6870
      MaxLength       =   40
      TabIndex        =   37
      Text            =   "txttemp(39)"
      ToolTipText     =   "Press F3 for Creation"
      Top             =   2940
      Width           =   4485
   End
   Begin VB.TextBox txttemp 
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
      Index           =   38
      Left            =   6870
      MaxLength       =   40
      TabIndex        =   36
      Text            =   "txttemp(38)"
      ToolTipText     =   "Press F3 for Creation"
      Top             =   2565
      Width           =   4485
   End
   Begin VB.TextBox txttemp 
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
      Index           =   37
      Left            =   6930
      MaxLength       =   40
      TabIndex        =   33
      Text            =   "txttemp(37)"
      ToolTipText     =   "Press F3 for Creation"
      Top             =   1725
      Width           =   3135
   End
   Begin VB.TextBox txttemp 
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
      Index           =   36
      Left            =   6930
      MaxLength       =   40
      TabIndex        =   32
      Text            =   "txttemp(36)"
      ToolTipText     =   "Press F3 for Creation"
      Top             =   1320
      Width           =   3135
   End
   Begin VB.TextBox txttemp 
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
      Index           =   35
      Left            =   6930
      MaxLength       =   40
      TabIndex        =   31
      Text            =   "txttemp(35)"
      ToolTipText     =   "Press F3 for Creation"
      Top             =   915
      Width           =   3135
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
      Left            =   20115
      MaxLength       =   40
      TabIndex        =   114
      Text            =   "40"
      ToolTipText     =   "Press F3 for Creation"
      Top             =   8955
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
      Left            =   20070
      MaxLength       =   40
      TabIndex        =   113
      Text            =   "39"
      ToolTipText     =   "Press F3 for Creation"
      Top             =   8550
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
      Index           =   38
      Left            =   20115
      MaxLength       =   40
      TabIndex        =   112
      Text            =   "38"
      ToolTipText     =   "Press F3 for Creation"
      Top             =   8190
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
      Index           =   37
      Left            =   20070
      MaxLength       =   40
      TabIndex        =   111
      Text            =   "37"
      ToolTipText     =   "Press F3 for Creation"
      Top             =   7785
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
      Index           =   36
      Left            =   20025
      MaxLength       =   40
      TabIndex        =   110
      Text            =   "36"
      ToolTipText     =   "Press F3 for Creation"
      Top             =   7425
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
      Left            =   20025
      MaxLength       =   40
      TabIndex        =   109
      Text            =   "35"
      ToolTipText     =   "Press F3 for Creation"
      Top             =   7065
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
      Left            =   14385
      MaxLength       =   40
      TabIndex        =   58
      Text            =   "34"
      ToolTipText     =   "Press F3 for Creation"
      Top             =   10395
      Visible         =   0   'False
      Width           =   465
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
      Left            =   14385
      MaxLength       =   40
      TabIndex        =   57
      Text            =   "33"
      ToolTipText     =   "Press F3 for Creation"
      Top             =   9990
      Visible         =   0   'False
      Width           =   465
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
      Left            =   14385
      MaxLength       =   40
      TabIndex        =   56
      Text            =   "32"
      ToolTipText     =   "Press F3 for Creation"
      Top             =   9585
      Visible         =   0   'False
      Width           =   465
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
      Left            =   14535
      MaxLength       =   40
      TabIndex        =   22
      Text            =   "txt(31)"
      ToolTipText     =   "Press F3 for Creation"
      Top             =   4740
      Visible         =   0   'False
      Width           =   600
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
      Left            =   14535
      MaxLength       =   40
      TabIndex        =   20
      Text            =   "txt(30)"
      ToolTipText     =   "Press F3 for Creation"
      Top             =   4320
      Visible         =   0   'False
      Width           =   600
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
      Left            =   14535
      MaxLength       =   40
      TabIndex        =   18
      Text            =   "txt(29)"
      ToolTipText     =   "Press F3 for Creation"
      Top             =   3930
      Visible         =   0   'False
      Width           =   600
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
      Left            =   14535
      MaxLength       =   40
      TabIndex        =   16
      Text            =   "txt(28)"
      ToolTipText     =   "Press F3 for Creation"
      Top             =   3555
      Visible         =   0   'False
      Width           =   600
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
      Left            =   14445
      MaxLength       =   40
      TabIndex        =   12
      Text            =   "txt(27)"
      ToolTipText     =   "Press F3 for Creation"
      Top             =   2745
      Visible         =   0   'False
      Width           =   645
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
      Left            =   14445
      MaxLength       =   40
      TabIndex        =   10
      Text            =   "txt(26)"
      ToolTipText     =   "Press F3 for Creation"
      Top             =   2355
      Visible         =   0   'False
      Width           =   645
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
      Left            =   14445
      MaxLength       =   40
      TabIndex        =   8
      Text            =   "txt(25)"
      ToolTipText     =   "Press F3 for Creation"
      Top             =   1980
      Visible         =   0   'False
      Width           =   645
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
      Left            =   14445
      MaxLength       =   40
      TabIndex        =   6
      Text            =   "txt(24)"
      ToolTipText     =   "Press F3 for Creation"
      Top             =   1605
      Visible         =   0   'False
      Width           =   645
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
      Left            =   14715
      MaxLength       =   40
      TabIndex        =   4
      Text            =   "txt(23)"
      ToolTipText     =   "Press F3 for Creation"
      Top             =   1320
      Visible         =   0   'False
      Width           =   645
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
      Left            =   14715
      MaxLength       =   40
      TabIndex        =   2
      Text            =   "txt(22)"
      ToolTipText     =   "Press F3 for Creation"
      Top             =   960
      Visible         =   0   'False
      Width           =   645
   End
   Begin VB.TextBox txttemp 
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
      Index           =   18
      Left            =   2190
      MaxLength       =   40
      TabIndex        =   50
      Text            =   "txttemp(18)"
      ToolTipText     =   "Press F3 for Creation"
      Top             =   7020
      Width           =   4485
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
      Index           =   8
      Left            =   7950
      MaxLength       =   40
      TabIndex        =   19
      Text            =   "txttemp(8)"
      ToolTipText     =   "Press F3 for Creation"
      Top             =   8895
      Visible         =   0   'False
      Width           =   345
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
      Index           =   9
      Left            =   7950
      MaxLength       =   40
      TabIndex        =   21
      Text            =   "txttemp(9)"
      ToolTipText     =   "Press F3 for Creation"
      Top             =   9315
      Visible         =   0   'False
      Width           =   345
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
      Index           =   5
      Left            =   8400
      MaxLength       =   40
      TabIndex        =   11
      Text            =   "txttemp(5)"
      ToolTipText     =   "Press F3 for Creation"
      Top             =   9600
      Visible         =   0   'False
      Width           =   345
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
      Index           =   12
      Left            =   2025
      MaxLength       =   40
      TabIndex        =   35
      Text            =   "txttemp(12)"
      ToolTipText     =   "Press F3 for Creation"
      Top             =   4560
      Width           =   4485
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
      Index           =   15
      Left            =   10920
      MaxLength       =   40
      TabIndex        =   45
      Text            =   "txttemp(15)"
      ToolTipText     =   "Press F3 for Creation"
      Top             =   4395
      Width           =   2250
   End
   Begin VB.TextBox txttemp 
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
      Index           =   16
      Left            =   2190
      MaxLength       =   40
      TabIndex        =   48
      Text            =   "txttemp(16)"
      ToolTipText     =   "Press F3 for Creation"
      Top             =   6210
      Width           =   4485
   End
   Begin VB.TextBox txttemp 
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
      Index           =   19
      Left            =   8670
      MaxLength       =   40
      TabIndex        =   51
      Text            =   "txttemp(19)"
      ToolTipText     =   "Press F3 for Creation"
      Top             =   5475
      Width           =   4485
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
      Left            =   15345
      TabIndex        =   102
      Text            =   "21"
      Top             =   7650
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
      Index           =   20
      Left            =   15345
      TabIndex        =   101
      Text            =   "20"
      Top             =   7290
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
      Index           =   19
      Left            =   15345
      TabIndex        =   100
      Text            =   "19"
      Top             =   6975
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
      Index           =   18
      Left            =   15345
      TabIndex        =   99
      Text            =   "18"
      Top             =   6570
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
      Index           =   17
      Left            =   15345
      TabIndex        =   98
      Text            =   "17"
      Top             =   6210
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
      Index           =   16
      Left            =   15345
      TabIndex        =   97
      Text            =   "16"
      Top             =   5895
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
      Index           =   15
      Left            =   15345
      TabIndex        =   96
      Text            =   "15"
      Top             =   5535
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
      Index           =   14
      Left            =   15345
      TabIndex        =   95
      Text            =   "14"
      Top             =   5175
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
      Index           =   13
      Left            =   15345
      TabIndex        =   94
      Text            =   "13"
      Top             =   4815
      Visible         =   0   'False
      Width           =   420
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
      Index           =   10
      Left            =   2010
      MaxLength       =   40
      TabIndex        =   27
      Text            =   "txttemp(10)"
      ToolTipText     =   "Press F3 for Creation"
      Top             =   1395
      Width           =   4485
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
      Index           =   11
      Left            =   2025
      MaxLength       =   40
      TabIndex        =   34
      Text            =   "txttemp(11)"
      ToolTipText     =   "Press F3 for Creation"
      Top             =   4125
      Width           =   4485
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
      Index           =   14
      Left            =   13800
      MaxLength       =   40
      TabIndex        =   29
      Text            =   "txttemp(14)"
      ToolTipText     =   "Press F3 for Creation"
      Top             =   120
      Visible         =   0   'False
      Width           =   1125
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
      Index           =   13
      Left            =   2010
      MaxLength       =   40
      TabIndex        =   25
      Text            =   "txttemp(13)"
      ToolTipText     =   "Press F3 for Creation"
      Top             =   555
      Width           =   4485
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
      Left            =   15345
      TabIndex        =   93
      Text            =   "12"
      Top             =   4545
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
      Left            =   15345
      TabIndex        =   92
      Text            =   "11"
      Top             =   4140
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
      Index           =   10
      Left            =   15345
      TabIndex        =   91
      Text            =   "10"
      Top             =   3735
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
      Index           =   9
      Left            =   15345
      TabIndex        =   90
      Text            =   "9"
      Top             =   3420
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
      Index           =   8
      Left            =   15345
      TabIndex        =   89
      Text            =   "8"
      Top             =   3060
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
      Index           =   7
      Left            =   15345
      TabIndex        =   88
      Text            =   "7"
      Top             =   2700
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
      Index           =   6
      Left            =   15345
      TabIndex        =   87
      Text            =   "6"
      Top             =   2340
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
      Index           =   5
      Left            =   15345
      TabIndex        =   86
      Text            =   "5"
      Top             =   2025
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
      Index           =   4
      Left            =   15345
      TabIndex        =   85
      Text            =   "4"
      Top             =   1710
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
      Index           =   3
      Left            =   15615
      TabIndex        =   84
      Text            =   "3"
      Top             =   1440
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
      Index           =   2
      Left            =   15615
      TabIndex        =   83
      Text            =   "2"
      Top             =   1080
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
      Index           =   1
      Left            =   15615
      TabIndex        =   82
      Text            =   "1"
      Top             =   765
      Visible         =   0   'False
      Width           =   420
   End
   Begin VB.TextBox txttemp 
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
      Index           =   21
      Left            =   8670
      MaxLength       =   40
      TabIndex        =   53
      Text            =   "txttemp(21)"
      ToolTipText     =   "Press F3 for Creation"
      Top             =   6285
      Width           =   4485
   End
   Begin VB.TextBox txttemp 
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
      Index           =   17
      Left            =   2205
      MaxLength       =   40
      TabIndex        =   49
      Text            =   "txttemp(17)"
      ToolTipText     =   "Press F3 for Creation"
      Top             =   6615
      Width           =   4485
   End
   Begin VB.TextBox txttemp 
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
      Index           =   20
      Left            =   8670
      MaxLength       =   40
      TabIndex        =   52
      Text            =   "txttemp(20)"
      ToolTipText     =   "Press F3 for Creation"
      Top             =   5880
      Width           =   4485
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
      Index           =   7
      Left            =   7950
      MaxLength       =   40
      TabIndex        =   17
      Text            =   "txttemp(7)"
      ToolTipText     =   "Press F3 for Creation"
      Top             =   8505
      Visible         =   0   'False
      Width           =   345
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
      Index           =   6
      Left            =   7950
      MaxLength       =   40
      TabIndex        =   15
      Text            =   "txttemp(6)"
      ToolTipText     =   "Press F3 for Creation"
      Top             =   8130
      Visible         =   0   'False
      Width           =   345
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
      Index           =   3
      Left            =   8400
      MaxLength       =   40
      TabIndex        =   7
      Text            =   "txttemp(3)"
      ToolTipText     =   "Press F3 for Creation"
      Top             =   8835
      Visible         =   0   'False
      Width           =   345
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
      Index           =   2
      Left            =   8400
      MaxLength       =   40
      TabIndex        =   5
      Text            =   "txttemp(2)"
      ToolTipText     =   "Press F3 for Creation"
      Top             =   8460
      Visible         =   0   'False
      Width           =   345
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
      Index           =   1
      Left            =   8580
      MaxLength       =   40
      TabIndex        =   3
      Text            =   "txttemp(1)"
      ToolTipText     =   "Press F3 for Creation"
      Top             =   7560
      Visible         =   0   'False
      Width           =   345
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
      Index           =   0
      Left            =   8520
      MaxLength       =   40
      TabIndex        =   1
      Text            =   "txttemp(0)"
      ToolTipText     =   "Press F3 for Creation"
      Top             =   8040
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
      Index           =   0
      Left            =   15615
      TabIndex        =   77
      Text            =   "0"
      Top             =   405
      Visible         =   0   'False
      Width           =   420
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
      Left            =   8400
      MaxLength       =   40
      TabIndex        =   9
      Text            =   "txttemp(4)"
      ToolTipText     =   "Press F3 for Creation"
      Top             =   9210
      Visible         =   0   'False
      Width           =   345
   End
   Begin VB.CommandButton CmdSave 
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
      Left            =   11970
      Style           =   1  'Graphical
      TabIndex        =   74
      ToolTipText     =   "Save Record"
      Top             =   2400
      Width           =   690
   End
   Begin VB.CommandButton cmdExit 
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
      Left            =   11970
      Style           =   1  'Graphical
      TabIndex        =   75
      ToolTipText     =   "Exit"
      Top             =   2985
      Width           =   690
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
      Left            =   14985
      Style           =   1  'Graphical
      TabIndex        =   78
      ToolTipText     =   "To Add New Record"
      Top             =   9090
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
      Left            =   15390
      Style           =   1  'Graphical
      TabIndex        =   79
      ToolTipText     =   "To Add New Record"
      Top             =   9225
      Visible         =   0   'False
      Width           =   1095
   End
   Begin Crystal.CrystalReport cryReport 
      Bindings        =   "frmSetting.frx":058A
      Left            =   15435
      Top             =   405
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      WindowControlBox=   -1  'True
      WindowMaxButton =   -1  'True
      WindowMinButton =   -1  'True
      PrintFileType   =   15
      PrintFileLinesPerPage=   60
   End
   Begin VB.Label lbl1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "TDS On Sales          :"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   225
      Index           =   0
      Left            =   360
      TabIndex        =   172
      Top             =   8400
      Width           =   1665
   End
   Begin VB.Label lbl1 
      BackStyle       =   0  'Transparent
      Caption         =   "TCS Payable A/c       :"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   360
      Index           =   57
      Left            =   6780
      TabIndex        =   167
      Top             =   7080
      Width           =   2205
   End
   Begin VB.Label lbl1 
      BackStyle       =   0  'Transparent
      Caption         =   "TCS Recivable A/c    :"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   360
      Index           =   56
      Left            =   6780
      TabIndex        =   166
      Top             =   6780
      Width           =   2205
   End
   Begin VB.Label lbl1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "For Mill Bill Direct Pay. To Mill"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   225
      Index           =   55
      Left            =   10920
      TabIndex        =   163
      Top             =   4080
      Width           =   2445
   End
   Begin VB.Label lbl1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "ForMill Bill  Statement A/c Setting "
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   225
      Index           =   54
      Left            =   180
      TabIndex        =   162
      Top             =   3780
      Width           =   2835
   End
   Begin VB.Label lbl1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "For Mill Bill Direct Payment To Party Statement A/c Setting "
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   225
      Index           =   53
      Left            =   -9480
      TabIndex        =   161
      Top             =   3060
      Width           =   4935
   End
   Begin VB.Label lbl1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "&L P Int Without GST :"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   225
      Index           =   52
      Left            =   360
      TabIndex        =   159
      Top             =   5895
      Width           =   1665
   End
   Begin VB.Label lbl1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "I"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   225
      Index           =   51
      Left            =   6660
      TabIndex        =   158
      Top             =   1755
      Width           =   45
   End
   Begin VB.Label lbl1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "C"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   225
      Index           =   50
      Left            =   6660
      TabIndex        =   157
      Top             =   1395
      Width           =   120
   End
   Begin VB.Label lbl1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "S"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   225
      Index           =   49
      Left            =   6660
      TabIndex        =   156
      Top             =   990
      Width           =   120
   End
   Begin VB.Label lbl1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "RCM GST Recivable A/c"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   225
      Index           =   48
      Left            =   10215
      TabIndex        =   155
      Top             =   585
      Width           =   1965
   End
   Begin VB.Label lbl1 
      BackStyle       =   0  'Transparent
      Caption         =   "Rounding Off A/c    : "
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   360
      Index           =   47
      Left            =   6840
      TabIndex        =   148
      Top             =   3960
      Width           =   1845
   End
   Begin VB.Label lbl1 
      BackStyle       =   0  'Transparent
      Caption         =   "Unloading Chg. A/c  :"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   360
      Index           =   46
      Left            =   180
      TabIndex        =   144
      Top             =   3360
      Width           =   2205
   End
   Begin VB.Label lbl1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "For Depot Statement A/c Setting "
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   225
      Index           =   32
      Left            =   225
      TabIndex        =   143
      Top             =   2295
      Width           =   2745
   End
   Begin VB.Label lbl1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "&Late Pay Interest GST :"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   225
      Index           =   31
      Left            =   360
      TabIndex        =   141
      Top             =   5445
      Width           =   1860
   End
   Begin VB.Label lbl1 
      BackStyle       =   0  'Transparent
      Caption         =   "Commission Item  :"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   360
      Index           =   19
      Left            =   6840
      TabIndex        =   137
      Top             =   4815
      Width           =   1665
   End
   Begin VB.Label lbl1 
      BackStyle       =   0  'Transparent
      Caption         =   "Brokerage Item      : "
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   360
      Index           =   14
      Left            =   6840
      TabIndex        =   136
      Top             =   4395
      Width           =   1845
   End
   Begin VB.Label lbl1 
      BackStyle       =   0  'Transparent
      Caption         =   "Commission Sale    :"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   360
      Index           =   5
      Left            =   180
      TabIndex        =   135
      Top             =   3015
      Width           =   2205
   End
   Begin VB.Label lbl1 
      BackStyle       =   0  'Transparent
      Caption         =   " Interest Rate   :"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   360
      Index           =   45
      Left            =   10665
      TabIndex        =   134
      Top             =   9720
      Width           =   1395
   End
   Begin VB.Label lbl1 
      BackStyle       =   0  'Transparent
      Caption         =   "Grace Days    :"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   360
      Index           =   44
      Left            =   10755
      TabIndex        =   133
      Top             =   9360
      Width           =   1290
   End
   Begin VB.Line Line5 
      BorderColor     =   &H00000080&
      X1              =   9360
      X2              =   13185
      Y1              =   10080
      Y2              =   10080
   End
   Begin VB.Label lbl1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   " Payment Interest Credit Note  :"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   225
      Index           =   43
      Left            =   9540
      TabIndex        =   132
      Top             =   9000
      Width           =   2625
   End
   Begin VB.Line Line4 
      BorderColor     =   &H00000080&
      X1              =   9360
      X2              =   13185
      Y1              =   8910
      Y2              =   8910
   End
   Begin VB.Label lbl1 
      BackStyle       =   0  'Transparent
      Caption         =   "Extra Grace Days    :"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   360
      Index           =   42
      Left            =   10305
      TabIndex        =   131
      Top             =   8145
      Width           =   1740
   End
   Begin VB.Label lbl1 
      BackStyle       =   0  'Transparent
      Caption         =   "Brokerage Sale A/c  :"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   360
      Index           =   41
      Left            =   225
      TabIndex        =   129
      Top             =   2610
      Width           =   2205
   End
   Begin VB.Label lbl1 
      BackStyle       =   0  'Transparent
      Caption         =   "Exempt  Sales  A/c   :"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   360
      Index           =   40
      Left            =   14100
      TabIndex        =   127
      Top             =   6480
      Width           =   1845
   End
   Begin VB.Label lbl1 
      BackStyle       =   0  'Transparent
      Caption         =   "Purchase Return     :"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   360
      Index           =   39
      Left            =   240
      TabIndex        =   126
      Top             =   1020
      Width           =   2070
   End
   Begin VB.Label lbl1 
      BackStyle       =   0  'Transparent
      Caption         =   "T.D.S. Rate   :"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   360
      Index           =   36
      Left            =   10800
      TabIndex        =   122
      Top             =   10215
      Width           =   1215
   End
   Begin VB.Label lbl1 
      BackStyle       =   0  'Transparent
      Caption         =   " Interest Rate For 365 Days   :"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   360
      Index           =   35
      Left            =   9540
      TabIndex        =   121
      Top             =   8505
      Width           =   2520
   End
   Begin VB.Label lbl1 
      BackStyle       =   0  'Transparent
      Caption         =   "Grace Days    :"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   360
      Index           =   34
      Left            =   10755
      TabIndex        =   120
      Top             =   7830
      Width           =   1350
   End
   Begin VB.Label lbl1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Late Payment Interest Debit Note  :"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   225
      Index           =   33
      Left            =   9540
      TabIndex        =   119
      Top             =   7515
      Width           =   2925
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H00EFB17A&
      BorderWidth     =   5
      FillColor       =   &H00EFB17A&
      Height          =   10725
      Left            =   0
      Top             =   0
      Width           =   13305
   End
   Begin VB.Label lbl1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "GST Account Setting  :"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   225
      Index           =   23
      Left            =   6870
      TabIndex        =   108
      Top             =   2205
      Width           =   1875
   End
   Begin VB.Label lbl1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "RCM GST Payable A/c"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   225
      Index           =   22
      Left            =   6930
      TabIndex        =   107
      Top             =   585
      Width           =   1815
   End
   Begin VB.Label lbl1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Payment JV Account Setting  :"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   225
      Index           =   30
      Left            =   6765
      TabIndex        =   106
      Top             =   5205
      Width           =   2550
   End
   Begin VB.Label lbl1 
      BackStyle       =   0  'Transparent
      Caption         =   "Discount   A/c.                 :"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   360
      Index           =   20
      Left            =   360
      TabIndex        =   105
      Top             =   7065
      Width           =   2205
   End
   Begin VB.Line Line2 
      BorderColor     =   &H000000C0&
      X1              =   0
      X2              =   13185
      Y1              =   7380
      Y2              =   7380
   End
   Begin VB.Label lbl1 
      BackStyle       =   0  'Transparent
      Caption         =   "Commission Sale    :"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   360
      Index           =   25
      Left            =   180
      TabIndex        =   104
      Top             =   4620
      Width           =   1845
   End
   Begin VB.Label lbl1 
      BackStyle       =   0  'Transparent
      Caption         =   "Hunk  Purchase  A/c     :"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   360
      Index           =   24
      Left            =   13860
      TabIndex        =   103
      Top             =   5880
      Width           =   2070
   End
   Begin VB.Line Line1 
      BorderColor     =   &H000000C0&
      X1              =   15
      X2              =   13365
      Y1              =   5145
      Y2              =   5130
   End
   Begin VB.Label lbl1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Receipt JV Account Setting  :"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   225
      Index           =   21
      Left            =   360
      TabIndex        =   68
      Top             =   5205
      Width           =   2445
   End
   Begin VB.Label lbl1 
      BackStyle       =   0  'Transparent
      Caption         =   "Brokerage Sale A/c  :"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   360
      Index           =   18
      Left            =   180
      TabIndex        =   67
      Top             =   4185
      Width           =   2205
   End
   Begin VB.Label lbl1 
      BackStyle       =   0  'Transparent
      Caption         =   "Trade Sales  A/c      :"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   360
      Index           =   17
      Left            =   240
      TabIndex        =   76
      Top             =   1425
      Width           =   2205
   End
   Begin VB.Label lbl1 
      BackStyle       =   0  'Transparent
      Caption         =   "Sales Return  A/c    :  "
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   360
      Index           =   16
      Left            =   240
      TabIndex        =   66
      Top             =   1860
      Width           =   2205
   End
   Begin VB.Label lbl1 
      BackStyle       =   0  'Transparent
      Caption         =   "Trade Purchase  A/c    :"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   270
      Index           =   15
      Left            =   180
      TabIndex        =   65
      Top             =   555
      Width           =   2070
   End
   Begin VB.Label lbl1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "&Late Pay Interest A/c  :"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   225
      Index           =   13
      Left            =   6765
      TabIndex        =   71
      Top             =   5475
      Width           =   1785
   End
   Begin VB.Label lbl1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "&Late Pay Interest A/c      :"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   225
      Index           =   12
      Left            =   360
      TabIndex        =   69
      Top             =   6255
      Width           =   1965
   End
   Begin VB.Label lbl1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "TDS On Purchase   :"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   225
      Index           =   2
      Left            =   390
      TabIndex        =   0
      Top             =   7845
      Width           =   1665
   End
   Begin VB.Shape Shape4 
      BorderColor     =   &H00EFB17A&
      BorderWidth     =   5
      FillColor       =   &H00EFB17A&
      Height          =   225
      Left            =   45
      Top             =   -135
      Width           =   3165
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
      Left            =   10530
      TabIndex        =   81
      Top             =   0
      Width           =   2475
   End
   Begin VB.Label lbl1 
      BackStyle       =   0  'Transparent
      Caption         =   "Discount   A/c.            :"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   360
      Index           =   4
      Left            =   6765
      TabIndex        =   73
      Top             =   6330
      Width           =   2205
   End
   Begin VB.Label lbl1 
      BackStyle       =   0  'Transparent
      Caption         =   "TDS  A/c            :"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   360
      Index           =   3
      Left            =   360
      TabIndex        =   70
      Top             =   6615
      Width           =   2205
   End
   Begin VB.Label lbl1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "TDS Paid A/c               :"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   225
      Index           =   1
      Left            =   6765
      TabIndex        =   72
      Top             =   5955
      Width           =   1785
   End
   Begin VB.Shape Shape1 
      FillColor       =   &H00FED8FC&
      FillStyle       =   0  'Solid
      Height          =   1560
      Left            =   11790
      Shape           =   4  'Rounded Rectangle
      Top             =   2085
      Width           =   1080
   End
   Begin VB.Label Label6 
      BackColor       =   &H00EFB17A&
      Caption         =   "    Settings"
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
      Left            =   -6285
      TabIndex        =   80
      Top             =   0
      Width           =   19680
   End
   Begin VB.Shape Shape3 
      BorderColor     =   &H000000C0&
      FillColor       =   &H00EDFDFE&
      FillStyle       =   0  'Solid
      Height          =   3165
      Left            =   9360
      Top             =   7425
      Width           =   3795
   End
End
Attribute VB_Name = "frmSetting"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Const tblName As String = "tblMastSetting"
Const MaxNo As Long = 73
Dim FormAction As Integer
Dim rstbl As Recordset
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case 13:
            Sendkeys "{TAB}"
    End Select
End Sub
Private Sub Form_Load()
'On Error GoTo ErrorHandler
    Dim X As String
    Dim i As Long
    Me.Top = 300
    Me.Left = Screen.Width / 2 - Me.Width / 2
    X = "Select * from " & tblName
    GProcRstOpen rstbl, X, "O"
    FormAction = vbDataActionAddNew
    Call GProcClearForm(Me, rstbl, MaxNo, True)
'    For i = 22 To 31
'        lblNature(i).Caption = ""
'        lblDescription(i).Caption = ""
'    Next i
'    For i = 41 To 44
'        lblNature(i).Caption = ""
'        lblDescription(i).Caption = ""
'    Next i
    If rstbl.EOF And rstbl.BOF Then
    Else
        ReadFields
    End If
    Exit Sub
ErrorHandler:
    rstbl.CancelUpdate
    GProcErrorHandler
End Sub
Private Sub cmdExit_Click()
Unload Me
End Sub
Private Sub cmdSave_Click()
'On Error GoTo ErrorHandler
If ValidateData = True Then
    gCn.Execute "delete from " & tblName
    FillTxtFromTemp
    GProcSaveRecord Me, rstbl, FormAction, MaxNo
    GProcGetSettingDetail
    Unload Me
End If
Exit Sub
ErrorHandler:
    rstbl.CancelUpdate
    GProcErrorHandler
End Sub
Private Sub txt_KeyPress(Index As Integer, KeyAscii As Integer)
Dim LeftPos As Long
Dim TopPos As Long
TopPos = txt(Index).Top + Me.Top + 650
LeftPos = txt(Index).Left + Me.Left
If KeyAscii = 13 Then Exit Sub '-- Enter
Select Case Index
       Case 22, 23, 24, 25, 26, 27, 50 '-- transaction code Sales
           gClsSearch.SearchMultiField "tblMastTransactionCode", "TranCode,iif(TranType=0,'Purchase',iif(TranType=1,'Purchase Goods Return',iif(TranType=2,'Purchase Debit Note',iif(TranType=3,'Sales', iif(TranType=4,'Sales Goods Return',iif(TranType=5,'Sales Credit Note', '') ) ) ) )),Nature,Description", Array("Tr. Code", "Tr. Type", "Nature", "Description"), Array(1000, 4000, 4000, 4000), "TranType = 3", "", "TranCode", txt(Index).Left + Me.Left, txt(Index).Top + Me.Top + 650
           KeyAscii = 0
           If gClsSearch.SearchMultiRetCol(0) <> "" Then
              txt(Index) = gClsSearch.SearchMultiRetCol(0)
           End If
       Case 28, 29, 30, 31, 53 '-- transaction code Purchase
           gClsSearch.SearchMultiField "tblMastTransactionCode", "TranCode,iif(TranType=0,'Purchase',iif(TranType=1,'Purchase Goods Return',iif(TranType=2,'Purchase Debit Note',iif(TranType=3,'Sales', iif(TranType=4,'Sales Goods Return',iif(TranType=5,'Sales Credit Note', '') ) ) ) )),Nature,Description", Array("Tr. Code", "Tr. Type", "Nature", "Description"), Array(1000, 4000, 4000, 4000), "TranType = 0", "", "TranCode", txt(Index).Left + Me.Left, txt(Index).Top + Me.Top + 650
           KeyAscii = 0
           If gClsSearch.SearchMultiRetCol(0) <> "" Then
              txt(Index) = gClsSearch.SearchMultiRetCol(0)
           End If
       Case 41, 42 '-- transaction code Sales
           gClsSearch.SearchMultiField "tblMastTransactionCode", "TranCode,iif(TranType=0,'Purchase',iif(TranType=1,'Purchase Goods Return',iif(TranType=2,'Purchase Debit Note',iif(TranType=3,'Sales', iif(TranType=4,'Sales Goods Return',iif(TranType=5,'Sales Credit Note', '') ) ) ) )),Nature,Description", Array("Tr. Code", "Tr. Type", "Nature", "Description"), Array(1000, 4000, 4000, 4000), "TranType = 4", "", "TranCode", txt(Index).Left + Me.Left, txt(Index).Top + Me.Top + 650
           KeyAscii = 0
           If gClsSearch.SearchMultiRetCol(0) <> "" Then
              txt(Index) = gClsSearch.SearchMultiRetCol(0)
           End If
       Case 43, 44 '-- transaction code Purchase
           gClsSearch.SearchMultiField "tblMastTransactionCode", "TranCode,iif(TranType=0,'Purchase',iif(TranType=1,'Purchase Goods Return',iif(TranType=2,'Purchase Debit Note',iif(TranType=3,'Sales', iif(TranType=4,'Sales Goods Return',iif(TranType=5,'Sales Credit Note', '') ) ) ) )),Nature,Description", Array("Tr. Code", "Tr. Type", "Nature", "Description"), Array(1000, 4000, 4000, 4000), "TranType = 1", "", "TranCode", txt(Index).Left + Me.Left, txt(Index).Top + Me.Top + 650
           KeyAscii = 0
           If gClsSearch.SearchMultiRetCol(0) <> "" Then
              txt(Index) = gClsSearch.SearchMultiRetCol(0)
           End If
      
      End Select
End Sub

Private Sub txt_LostFocus(Index As Integer)
Dim i As Long
'For i = 22 To 31 '-- transaction Code, nature discription
'    If txt(i) = "" Then
'       lblNature(i).Caption = ""
'       lblDescription(i).Caption = ""
'    Else
'       lblNature(i).Caption = GProcGetColumnValue("tblMastTransactionCode", "TranCode", txt(i), "S", "Nature", "S")
'       lblDescription(i).Caption = GProcGetColumnValue("tblMastTransactionCode", "TranCode", txt(i), "S", "Description", "S")
'    End If
'Next i
'For i = 41 To 44 '-- transaction Code, nature discription
'    If txt(i) = "" Then
'       lblNature(i).Caption = ""
'       lblDescription(i).Caption = ""
'    Else
'       lblNature(i).Caption = GProcGetColumnValue("tblMastTransactionCode", "TranCode", txt(i), "S", "Nature", "S")
'       lblDescription(i).Caption = GProcGetColumnValue("tblMastTransactionCode", "TranCode", txt(i), "S", "Description", "S")
'    End If
'Next i


End Sub

Private Sub txttemp_GotFocus(Index As Integer)
GProcSelectBox Me.ActiveControl
End Sub

Private Sub txttemp_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
Dim GName As String
If KeyCode = vbKeyF3 Then
Select Case Index
    Case 1, 2, 3, 4, 5, 7, 8, 9 '-- Tax
        GProcShowForm frmMastNarrationTax, frmMain.mnuMstfrm(6)
    Case Else '-- A/c
        'GName = GProcGetColumnValue("TblMastGroup", "AgCode", 90029, "N", "AgName", "S")
        GProcShowForm frmMastAccount, frmMain.mnuMstfrm(0) ', Array(frmMastAccount.txtTemp(2)), Array(GName)
End Select
End If
End Sub
Private Sub txtTemp_KeyPress(Index As Integer, KeyAscii As Integer)
Dim LeftPos As Long
Dim TopPos As Long
Dim X As String
Dim FindQStr As String
TopPos = txttemp(Index).Top + Me.Top + 650
LeftPos = txttemp(Index).Left + Me.Left
If KeyAscii = 13 Then Exit Sub '-- Enter
Select Case Index
     Case 1, 2, 3, 4, 5, 7, 8, 7, 9, 48, 51 '-- Tax
          X = " TblMastAccount as TblMastac1 "
          FindQStr = "Select Narration,MastTaxRate,tblMastAccount.AcName,tblMastAc1.AcName from " _
          & " tblMastNarration,TblMastaccount," & X & " where tblMastNarration.MastTaxPurAcCode=TblMastaccount.AcCode" _
          & " and tblMastNarration.MastTaxSAAcCode=TblMastac1.AcCode and NarrType='T'"
          gClsSearch.SearchMultiField FindQStr, "", Array("Tax", "Rate", "Purchase Account", "Sales Account"), Array(2000, 1000, 2000, 2000), "", "", "Narration", LeftPos, TopPos, False, , True
          KeyAscii = 0
          If gClsSearch.SearchMultiRetCol(0) <> "" Then
              txttemp(Index) = gClsSearch.SearchMultiRetCol(0)
          End If
'          If gClsSearch.SearchMultiRetCol(1) <> "" Then
'             txtTaxCode(Index) = gClsSearch.SearchMultiRetCol(1)
'          Else
'             txtTaxCode(Index) = ""
'          End If
'
'          If gClsSearch.SearchMultiRetCol(2) <> "" Then
'             lblNature(Index).Caption = gClsSearch.SearchMultiRetCol(2)
'          Else
'             lblNature(Index).Caption = ""
'          End If
'          If gClsSearch.SearchMultiRetCol(3) <> "" Then
'             lblDescription(Index).Caption = gClsSearch.SearchMultiRetCol(3)
'          Else
'             lblDescription(Index).Caption = ""
'          End If
     Case 35, 36, 37, 38, 39, 40, 64, 65, 66
           gClsSearch.SearchMultiField "tblMastaccount,tblMastGroup", "AcName,AgName", Array("Account Name", "Group Name"), Array(5000, 3000), "tblMastaccount.AgCode = tblMastGroup.AgCode and (tblMastaccount.AgCode=90044)  ", Chr(KeyAscii), "", 600 + Me.Left, 2500, True
           KeyAscii = 0
           If gClsSearch.SearchMultiRetCol(0) <> "" Then
               txttemp(Index) = gClsSearch.SearchMultiRetCol(0)
           End If
     Case 59, 60, 15, 49
           gClsSearch.SearchMultiField "tblMastItem", "ItName,ItCode,ITUnit", Array("Item", "Item", "Unit"), Array(5000, 0, 1000), " not ItName = '-' and  tblMastItem.IsOtherItem=1  ", Chr(KeyAscii), "ItName", 600 + Me.Left, 2500, True
           KeyAscii = 0
           If gClsSearch.SearchMultiRetCol(0) <> "" Then
               txttemp(Index) = gClsSearch.SearchMultiRetCol(0)
           End If
     Case Else  '-- A/C
          gClsSearch.SearchMultiField "tblMastaccount,tblMastGroup", "AcName,AgName", Array("Account Name", "Group Name"), Array(5000, 3000), "tblMastaccount.AgCode = tblMastGroup.AgCode and (tblMastaccount.AgCode<>90017) ", Chr(KeyAscii), "", LeftPos, TopPos, True
          KeyAscii = 0
          If gClsSearch.SearchMultiRetCol(0) <> "" Then
              txttemp(Index) = gClsSearch.SearchMultiRetCol(0)
          End If
End Select
End Sub
Private Sub ReadFields()
GProcShowRecord Me, rstbl, MaxNo
FillTempFromTxt

End Sub
Private Function ValidateData() As Boolean
Dim GName As String
Dim i As Long
'-- Create Tax Type
For i = 0 To 9
    If Trim(txttemp(i)) = "" Then
       txttemp(i) = "-"
    End If
Next i
'--Create A/c

For i = 10 To 21
    If Trim(txttemp(i)) <> "" Then
       If i <> 15 Then
            If GProcGetColumnValue("tblMastAccount", "AcName", txttemp(i), "S", "AcCode", "N") = 0 Then
               'GName = GProcGetColumnValue("TblMastGroup", "AgCode", 90029, "N", "AgName", "S")
               GProcShowForm frmMastAccount, frmMain.mnuMstfrm(0), Array(frmMastAccount.txt(1), frmMastAccount.txttemp(2)), Array(txttemp(i), GName)
               Exit Function
            End If
        End If
    Else
        txttemp(i) = "-"
    End If
Next i
For i = 35 To 40
    If Trim(txttemp(i)) <> "" Then
        If GProcGetColumnValue("tblMastAccount", "AcName", txttemp(i), "S", "AcCode", "N") = 0 Then
           'GName = GProcGetColumnValue("TblMastGroup", "AgCode", 90029, "N", "AgName", "S")
           GProcShowForm frmMastAccount, frmMain.mnuMstfrm(0), Array(frmMastAccount.txt(1), frmMastAccount.txttemp(2)), Array(txttemp(i), GName)
           Exit Function
        End If
    Else
        txttemp(i) = "-"
    End If
Next i
ValidateData = True
End Function
Private Sub FillTempFromTxt()
Dim i As Long
For i = 0 To 9 '-- Tax Type
    txttemp(i) = GProcGetColumnValue("tblMastNarration", "NarrCode", txt(i), "N", "Narration", "S")
Next i
txttemp(48) = GProcGetColumnValue("tblMastNarration", "NarrCode", txt(48), "N", "Narration", "S")
txttemp(51) = GProcGetColumnValue("tblMastNarration", "NarrCode", txt(51), "N", "Narration", "S")

For i = 10 To 21 '-- Account
    If i = 15 Then
       txttemp(i) = GProcGetColumnValue("tblMastItem", "ItCode", txt(i), "N", "ItName", "S")
    Else
       txttemp(i) = GProcGetColumnValue("tblMastAccount", "AcCode", txt(i), "N", "AcName", "S")
    End If
Next i
For i = 35 To 40 '-- Account
    txttemp(i) = GProcGetColumnValue("tblMastAccount", "AcCode", txt(i), "N", "AcName", "S")
Next i
txttemp(49) = GProcGetColumnValue("tblMastItem", "ItCode", txt(49), "N", "ItName", "S")


txttemp(52) = GProcGetColumnValue("tblMastAccount", "AcCode", txt(52), "N", "AcName", "S")
txttemp(54) = GProcGetColumnValue("tblMastAccount", "AcCode", txt(54), "N", "AcName", "S")
txttemp(58) = GProcGetColumnValue("tblMastAccount", "AcCode", txt(58), "N", "AcName", "S")
txttemp(59) = GProcGetColumnValue("tblMastItem", "ItCode", txt(59), "N", "ItName", "S")
txttemp(60) = GProcGetColumnValue("tblMastItem", "ItCode", txt(60), "N", "ItName", "S")
txttemp(61) = GProcGetColumnValue("tblMastAccount", "AcCode", txt(61), "N", "AcName", "S")
txttemp(62) = GProcGetColumnValue("tblMastAccount", "AcCode", txt(62), "N", "AcName", "S")
txttemp(63) = GProcGetColumnValue("tblMastAccount", "AcCode", txt(63), "N", "AcName", "S")
txttemp(64) = GProcGetColumnValue("tblMastAccount", "AcCode", txt(64), "N", "AcName", "S")
txttemp(65) = GProcGetColumnValue("tblMastAccount", "AcCode", txt(65), "N", "AcName", "S")
txttemp(66) = GProcGetColumnValue("tblMastAccount", "AcCode", txt(66), "N", "AcName", "S")
txttemp(67) = GProcGetColumnValue("tblMastAccount", "AcCode", txt(67), "N", "AcName", "S")
txttemp(68) = GProcGetColumnValue("tblMastAccount", "AcCode", txt(68), "N", "AcName", "S")
txttemp(69) = GProcGetColumnValue("tblMastAccount", "AcCode", txt(69), "N", "AcName", "S")
txttemp(70) = GProcGetColumnValue("tblMastAccount", "AcCode", txt(70), "N", "AcName", "S")
txttemp(71) = GProcGetColumnValue("tblMastAccount", "AcCode", txt(71), "N", "AcName", "S")
txttemp(72) = GProcGetColumnValue("tblMastAccount", "AcCode", txt(72), "N", "AcName", "S")
txttemp(73) = GProcGetColumnValue("tblMastAccount", "AcCode", txt(73), "N", "AcName", "S")

'For i = 22 To 31 '-- transaction Code, nature discription
'    If txt(i) = "" Then
'       lblNature(i).Caption = ""
'       lblDescription(i).Caption = ""
'    Else
'       lblNature(i).Caption = GProcGetColumnValue("tblMastTransactionCode", "TranCode", txt(i), "S", "Nature", "S")
'       lblDescription(i).Caption = GProcGetColumnValue("tblMastTransactionCode", "TranCode", txt(i), "S", "Description", "S")
'    End If
'Next i
'lblNature(50).Caption = GProcGetColumnValue("tblMastTransactionCode", "TranCode", txt(50), "S", "Nature", "S")
'lblDescription(50).Caption = GProcGetColumnValue("tblMastTransactionCode", "TranCode", txt(50), "S", "Description", "S")
'lblNature(53).Caption = GProcGetColumnValue("tblMastTransactionCode", "TranCode", txt(53), "S", "Nature", "S")
'lblDescription(53).Caption = GProcGetColumnValue("tblMastTransactionCode", "TranCode", txt(53), "S", "Description", "S")
'
'
'
'For i = 41 To 44 '-- transaction Code, nature discription
'    If txt(i) = "" Then
'       lblNature(i).Caption = ""
'       lblDescription(i).Caption = ""
'    Else
'       lblNature(i).Caption = GProcGetColumnValue("tblMastTransactionCode", "TranCode", txt(i), "S", "Nature", "S")
'       lblDescription(i).Caption = GProcGetColumnValue("tblMastTransactionCode", "TranCode", txt(i), "S", "Description", "S")
'    End If
'Next i


End Sub
Private Sub FillTxtFromTemp()
Dim i As Long
For i = 0 To 9 '-- Tax Type
    txt(i) = GProcGetColumnValue("tblMastNarration", "Narration", txttemp(i), "S", "NarrCode", "N")
Next i
txt(48) = GProcGetColumnValue("tblMastNarration", "Narration", txttemp(48), "S", "NarrCode", "N")
txt(51) = GProcGetColumnValue("tblMastNarration", "Narration", txttemp(51), "S", "NarrCode", "N")
For i = 10 To 21 '-- Account
    If i = 15 And txttemp(i) <> "-" Then
       txt(i) = GProcGetColumnValue("tblMastItem", "ItName", txttemp(i), "S", "ItCode", "N", "IsOtherItem = 1")
    Else
       txt(i) = GProcGetColumnValue("tblMastAccount", "AcName", txttemp(i), "S", "AcCode", "N")
    End If
Next i
txt(49) = GProcGetColumnValue("tblMastItem", "ItName", txttemp(49), "S", "ItCode", "N", "IsOtherItem = 1")

txt(52) = GProcGetColumnValue("tblMastAccount", "AcName", txttemp(52), "S", "AcCode", "N")
txt(54) = GProcGetColumnValue("tblMastAccount", "AcName", txttemp(54), "S", "AcCode", "N")
txt(58) = GProcGetColumnValue("tblMastAccount", "AcName", txttemp(58), "S", "AcCode", "N")
txt(59) = GProcGetColumnValue("tblMastItem", "ItName", txttemp(59), "S", "ItCode", "N", "IsOtherItem = 1")
txt(60) = GProcGetColumnValue("tblMastItem", "ItName", txttemp(60), "S", "ItCode", "N", "IsOtherItem = 1")
txt(61) = GProcGetColumnValue("tblMastAccount", "AcName", txttemp(61), "S", "AcCode", "N")
txt(62) = GProcGetColumnValue("tblMastAccount", "AcName", txttemp(62), "S", "AcCode", "N")
txt(63) = GProcGetColumnValue("tblMastAccount", "AcName", txttemp(63), "S", "AcCode", "N")
txt(64) = GProcGetColumnValue("tblMastAccount", "AcName", txttemp(64), "S", "AcCode", "N")
txt(65) = GProcGetColumnValue("tblMastAccount", "AcName", txttemp(65), "S", "AcCode", "N")
txt(66) = GProcGetColumnValue("tblMastAccount", "AcName", txttemp(66), "S", "AcCode", "N")
txt(67) = GProcGetColumnValue("tblMastAccount", "AcName", txttemp(67), "S", "AcCode", "N")
txt(68) = GProcGetColumnValue("tblMastAccount", "AcName", txttemp(68), "S", "AcCode", "N")
txt(69) = GProcGetColumnValue("tblMastAccount", "AcName", txttemp(69), "S", "AcCode", "N")
txt(70) = GProcGetColumnValue("tblMastAccount", "AcName", txttemp(70), "S", "AcCode", "N")
txt(71) = GProcGetColumnValue("tblMastAccount", "AcName", txttemp(71), "S", "AcCode", "N")
txt(72) = GProcGetColumnValue("tblMastAccount", "AcName", txttemp(72), "S", "AcCode", "N")
txt(73) = GProcGetColumnValue("tblMastAccount", "AcName", txttemp(73), "S", "AcCode", "N")

For i = 35 To 40 '-- Account
    txt(i) = GProcGetColumnValue("tblMastAccount", "AcName", txttemp(i), "S", "AcCode", "N")
Next i

End Sub

