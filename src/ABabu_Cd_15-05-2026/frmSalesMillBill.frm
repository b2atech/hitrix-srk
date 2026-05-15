VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form frmSalesMillBill 
   BackColor       =   &H00FBF2E1&
   BorderStyle     =   0  'None
   Caption         =   "Purchase Entry"
   ClientHeight    =   8625
   ClientLeft      =   0
   ClientTop       =   15
   ClientWidth     =   14550
   FillColor       =   &H000040C0&
   FillStyle       =   0  'Solid
   Icon            =   "frmSalesMillBill.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   8625
   ScaleWidth      =   14550
   ShowInTaskbar   =   0   'False
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
      Height          =   285
      Index           =   73
      Left            =   16020
      TabIndex        =   155
      Text            =   "73"
      Top             =   1920
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
      Height          =   285
      Index           =   72
      Left            =   16020
      TabIndex        =   154
      Text            =   "72"
      Top             =   1620
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
      Height          =   405
      Index           =   74
      Left            =   1920
      TabIndex        =   14
      Text            =   "74"
      Top             =   6960
      Width           =   1845
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
      Height          =   285
      Index           =   71
      Left            =   15420
      TabIndex        =   151
      Text            =   "71"
      Top             =   2940
      Visible         =   0   'False
      Width           =   345
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
      Index           =   70
      Left            =   1920
      TabIndex        =   17
      Text            =   "txtCode(70)"
      ToolTipText     =   "Press F3 for Creation"
      Top             =   7980
      Width           =   1095
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
      Height          =   405
      Index           =   70
      Left            =   3120
      TabIndex        =   18
      Text            =   "txttemp(70)"
      ToolTipText     =   "Press F3 for Creation"
      Top             =   7980
      Width           =   4335
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
      Height          =   285
      Index           =   70
      Left            =   15420
      TabIndex        =   148
      Text            =   "70"
      Top             =   2640
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
      Height          =   405
      Index           =   69
      Left            =   1920
      TabIndex        =   16
      Text            =   "69"
      Top             =   7440
      Width           =   1800
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
      Height          =   405
      Index           =   68
      Left            =   5220
      TabIndex        =   15
      Text            =   "68"
      Top             =   7020
      Width           =   1080
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
      Index           =   67
      Left            =   15420
      TabIndex        =   147
      Text            =   "67"
      Top             =   2160
      Visible         =   0   'False
      Width           =   300
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
      Index           =   66
      Left            =   15420
      TabIndex        =   146
      Text            =   "66"
      Top             =   1740
      Visible         =   0   'False
      Width           =   300
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
      Index           =   65
      Left            =   15420
      TabIndex        =   145
      Text            =   "65"
      Top             =   1380
      Visible         =   0   'False
      Width           =   300
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
      Index           =   64
      Left            =   15420
      TabIndex        =   144
      Text            =   "64"
      Top             =   960
      Visible         =   0   'False
      Width           =   300
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
      Index           =   63
      Left            =   15420
      TabIndex        =   143
      Text            =   "63"
      Top             =   600
      Visible         =   0   'False
      Width           =   300
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
      Index           =   62
      Left            =   15420
      TabIndex        =   142
      Text            =   "62"
      Top             =   120
      Visible         =   0   'False
      Width           =   300
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
      Index           =   61
      Left            =   14700
      TabIndex        =   138
      Text            =   "61"
      Top             =   5940
      Visible         =   0   'False
      Width           =   300
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
      Index           =   60
      Left            =   14700
      TabIndex        =   137
      Text            =   "60"
      Top             =   5580
      Visible         =   0   'False
      Width           =   300
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
      Index           =   59
      Left            =   14700
      TabIndex        =   136
      Text            =   "59"
      Top             =   5160
      Visible         =   0   'False
      Width           =   300
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
      Index           =   58
      Left            =   14700
      TabIndex        =   135
      Text            =   "58"
      Top             =   4800
      Visible         =   0   'False
      Width           =   300
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
      Index           =   57
      Left            =   14700
      TabIndex        =   134
      Text            =   "57"
      Top             =   4440
      Visible         =   0   'False
      Width           =   300
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
      Index           =   56
      Left            =   14700
      TabIndex        =   133
      Text            =   "56"
      Top             =   4020
      Visible         =   0   'False
      Width           =   300
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
      Index           =   55
      Left            =   14700
      TabIndex        =   132
      Text            =   "55"
      Top             =   3600
      Visible         =   0   'False
      Width           =   300
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
      Index           =   54
      Left            =   14700
      TabIndex        =   131
      Text            =   "54"
      Top             =   3300
      Visible         =   0   'False
      Width           =   300
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
      Index           =   53
      Left            =   14700
      TabIndex        =   130
      Text            =   "53"
      Top             =   2880
      Visible         =   0   'False
      Width           =   300
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
      Index           =   52
      Left            =   14700
      TabIndex        =   129
      Text            =   "52"
      Top             =   2460
      Visible         =   0   'False
      Width           =   300
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
      Index           =   51
      Left            =   14700
      TabIndex        =   128
      Text            =   "51"
      Top             =   2040
      Visible         =   0   'False
      Width           =   300
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
      Index           =   50
      Left            =   14700
      TabIndex        =   127
      Text            =   "50"
      Top             =   1620
      Visible         =   0   'False
      Width           =   300
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
      Index           =   49
      Left            =   14700
      TabIndex        =   126
      Text            =   "49"
      Top             =   1260
      Visible         =   0   'False
      Width           =   300
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
      Index           =   48
      Left            =   14700
      TabIndex        =   125
      Text            =   "48"
      Top             =   900
      Visible         =   0   'False
      Width           =   300
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
      Index           =   47
      Left            =   14700
      TabIndex        =   124
      Text            =   "47"
      Top             =   540
      Visible         =   0   'False
      Width           =   300
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
      Left            =   14700
      TabIndex        =   123
      Text            =   "46"
      Top             =   120
      Visible         =   0   'False
      Width           =   300
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
      Left            =   3540
      TabIndex        =   120
      Top             =   3120
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
         TabIndex        =   121
         Top             =   1320
         Width           =   1335
      End
      Begin VB.Label Label12 
         BackColor       =   &H00C0FFFF&
         Caption         =   "This Is No Work Party"
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
         TabIndex        =   122
         Top             =   420
         Width           =   6495
      End
   End
   Begin VB.CheckBox chkIsPaymenttoMill 
      BackColor       =   &H00FBF2E1&
      Caption         =   "Is Payment Direct To Mill  Of This Invoice ?"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   555
      Left            =   720
      TabIndex        =   13
      Top             =   6300
      Width           =   5775
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
      Left            =   1305
      TabIndex        =   1
      Text            =   "txtCode(7)"
      ToolTipText     =   "Press F3 for Creation"
      Top             =   855
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
      Left            =   1305
      TabIndex        =   106
      Text            =   "txtCode(5)"
      ToolTipText     =   "Press F3 for Creation"
      Top             =   1365
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
      Index           =   8
      Left            =   1305
      TabIndex        =   110
      Text            =   "txtCode(8)"
      ToolTipText     =   "Press F3 for Creation"
      Top             =   2520
      Width           =   1095
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
      Index           =   45
      Left            =   12030
      TabIndex        =   6
      Text            =   "txt(45)"
      Top             =   2130
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
      Height          =   375
      Index           =   44
      Left            =   16890
      TabIndex        =   104
      Text            =   "44"
      Top             =   3600
      Visible         =   0   'False
      Width           =   570
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
      Left            =   19800
      TabIndex        =   103
      Text            =   "43"
      Top             =   6105
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
      Index           =   42
      Left            =   19800
      TabIndex        =   102
      Text            =   "42"
      Top             =   5310
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
      Index           =   41
      Left            =   19800
      TabIndex        =   101
      Text            =   "41"
      Top             =   7335
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
      Index           =   40
      Left            =   16740
      TabIndex        =   100
      Text            =   "40"
      Top             =   1215
      Visible         =   0   'False
      Width           =   570
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
      Index           =   39
      Left            =   16695
      TabIndex        =   99
      Text            =   "39"
      Top             =   765
      Visible         =   0   'False
      Width           =   570
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
      Index           =   38
      Left            =   17775
      TabIndex        =   27
      Text            =   "38"
      Top             =   360
      Visible         =   0   'False
      Width           =   300
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
      Index           =   9
      Left            =   10020
      TabIndex        =   4
      Text            =   "9"
      Top             =   1680
      Width           =   1740
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
      Index           =   37
      Left            =   18810
      TabIndex        =   98
      Text            =   "37"
      Top             =   225
      Visible         =   0   'False
      Width           =   570
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
      Left            =   18540
      TabIndex        =   97
      Text            =   "36"
      Top             =   7605
      Visible         =   0   'False
      Width           =   570
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
      Left            =   13215
      Style           =   1  'Graphical
      TabIndex        =   95
      Top             =   5490
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
      Index           =   35
      Left            =   18585
      TabIndex        =   94
      Text            =   "35"
      Top             =   7200
      Visible         =   0   'False
      Width           =   570
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
      Left            =   17280
      TabIndex        =   21
      Text            =   "Cbo(22)"
      Top             =   4620
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.TextBox txttemp 
      Alignment       =   1  'Right Justify
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
      Height          =   360
      Index           =   35
      Left            =   17775
      MaxLength       =   5
      TabIndex        =   40
      Text            =   "txtemp(35)"
      Top             =   1665
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
      Index           =   22
      Left            =   18360
      TabIndex        =   91
      Text            =   "22"
      Top             =   4365
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.TextBox txt 
      Alignment       =   1  'Right Justify
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Index           =   16
      Left            =   10305
      TabIndex        =   12
      Text            =   "16"
      Top             =   7380
      Width           =   2400
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
      Index           =   15
      Left            =   10305
      TabIndex        =   10
      Text            =   "15"
      Top             =   6450
      Width           =   2400
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
      Left            =   18360
      TabIndex        =   90
      Text            =   "14"
      Top             =   4005
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
      Height          =   375
      Index           =   13
      Left            =   15255
      TabIndex        =   25
      Text            =   "13"
      Top             =   7110
      Visible         =   0   'False
      Width           =   480
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
      Index           =   11
      Left            =   15930
      TabIndex        =   24
      Text            =   "txttemp(11)"
      ToolTipText     =   "Press F3 for Creation & Press F2 for Selection"
      Top             =   6570
      Visible         =   0   'False
      Width           =   1335
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
      Left            =   18780
      TabIndex        =   37
      Text            =   "28"
      Top             =   2655
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
      Index           =   29
      Left            =   18900
      TabIndex        =   38
      Text            =   "29"
      Top             =   3105
      Visible         =   0   'False
      Width           =   330
   End
   Begin VB.CommandButton CmdBagDetails 
      Caption         =   "Cartoon &Bag Details"
      Height          =   720
      Left            =   17550
      TabIndex        =   85
      Top             =   5355
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
      Left            =   18405
      TabIndex        =   46
      Text            =   "8"
      Top             =   2835
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
      Index           =   34
      Left            =   18585
      TabIndex        =   84
      Text            =   "34"
      Top             =   6750
      Visible         =   0   'False
      Width           =   570
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
      Index           =   19
      Left            =   18450
      TabIndex        =   34
      Text            =   "19"
      Top             =   8100
      Visible         =   0   'False
      Width           =   885
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
      Left            =   18540
      TabIndex        =   83
      Text            =   "33"
      Top             =   6345
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
      Left            =   18495
      TabIndex        =   82
      Text            =   "32"
      Top             =   5940
      Visible         =   0   'False
      Width           =   420
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
      Height          =   405
      Index           =   8
      Left            =   2460
      TabIndex        =   112
      Text            =   "txttemp(8)"
      ToolTipText     =   "Press F3 for Creation"
      Top             =   2520
      Width           =   6210
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
      Height          =   405
      Index           =   5
      Left            =   2460
      TabIndex        =   108
      Text            =   "txttemp(5)"
      ToolTipText     =   "Press F3 for Creation"
      Top             =   1365
      Width           =   6210
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
      Left            =   2460
      TabIndex        =   2
      Text            =   "txttemp(7)"
      ToolTipText     =   "Press F3 for Creation"
      Top             =   855
      Width           =   6210
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
      Left            =   17460
      TabIndex        =   32
      Text            =   "txtDrBalance"
      Top             =   2160
      Visible         =   0   'False
      Width           =   525
   End
   Begin VB.CommandButton cmdLedger 
      Caption         =   "Ledger"
      Height          =   375
      Left            =   12915
      TabIndex        =   81
      Top             =   10980
      Visible         =   0   'False
      Width           =   1095
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
      Left            =   19800
      TabIndex        =   57
      Text            =   "27"
      Top             =   4905
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
      Index           =   26
      Left            =   19800
      TabIndex        =   56
      Text            =   "26"
      Top             =   6930
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
      Index           =   25
      Left            =   18405
      TabIndex        =   55
      Text            =   "25"
      Top             =   4770
      Visible         =   0   'False
      Width           =   480
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
      Index           =   21
      Left            =   10305
      TabIndex        =   11
      Text            =   "21"
      Top             =   6900
      Width           =   2400
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
      Left            =   435
      TabIndex        =   9
      Top             =   4140
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
      Left            =   18315
      TabIndex        =   79
      Text            =   "1"
      Top             =   495
      Visible         =   0   'False
      Width           =   420
   End
   Begin VB.TextBox txt 
      Alignment       =   1  'Right Justify
      Height          =   315
      Index           =   2
      Left            =   18270
      TabIndex        =   77
      Text            =   "2"
      Top             =   945
      Visible         =   0   'False
      Width           =   525
   End
   Begin VB.TextBox txt 
      Alignment       =   1  'Right Justify
      Height          =   315
      Index           =   5
      Left            =   18360
      TabIndex        =   76
      Text            =   "5"
      Top             =   1890
      Visible         =   0   'False
      Width           =   525
   End
   Begin VB.TextBox txt 
      Alignment       =   1  'Right Justify
      Height          =   315
      Index           =   3
      Left            =   18270
      TabIndex        =   75
      Text            =   "3"
      Top             =   1305
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
      Left            =   13365
      Style           =   1  'Graphical
      TabIndex        =   71
      Top             =   10980
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
      Left            =   12240
      Style           =   1  'Graphical
      TabIndex        =   70
      ToolTipText     =   "To Add New Record"
      Top             =   10980
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
      Left            =   13215
      Style           =   1  'Graphical
      TabIndex        =   68
      ToolTipText     =   "Exit"
      Top             =   6120
      Width           =   870
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
      Left            =   13215
      Style           =   1  'Graphical
      TabIndex        =   61
      ToolTipText     =   "To Save Record"
      Top             =   4845
      Width           =   870
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
      Left            =   17460
      Style           =   1  'Graphical
      TabIndex        =   69
      Top             =   3015
      Visible         =   0   'False
      Width           =   780
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
      Left            =   13215
      Style           =   1  'Graphical
      TabIndex        =   60
      ToolTipText     =   "To Delete Record"
      Top             =   4215
      Width           =   870
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
      Left            =   13215
      Style           =   1  'Graphical
      TabIndex        =   59
      ToolTipText     =   "To Modify Record"
      Top             =   3570
      Width           =   870
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
      Left            =   13215
      Style           =   1  'Graphical
      TabIndex        =   58
      ToolTipText     =   "To Add New Record"
      Top             =   2940
      Width           =   870
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
      Left            =   17280
      Style           =   1  'Graphical
      TabIndex        =   62
      Top             =   3375
      Visible         =   0   'False
      Width           =   780
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
      Left            =   13215
      Style           =   1  'Graphical
      TabIndex        =   63
      ToolTipText     =   "To Find Record"
      Top             =   2340
      Width           =   870
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
      Left            =   13710
      Style           =   1  'Graphical
      TabIndex        =   67
      ToolTipText     =   "Last"
      Top             =   1710
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
      Left            =   13215
      Style           =   1  'Graphical
      TabIndex        =   66
      ToolTipText     =   "First"
      Top             =   1710
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
      Left            =   13710
      Style           =   1  'Graphical
      TabIndex        =   65
      ToolTipText     =   "Next"
      Top             =   1080
      Width           =   375
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
      Left            =   18540
      TabIndex        =   31
      Text            =   "0"
      Top             =   90
      Visible         =   0   'False
      Width           =   300
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
      Left            =   13215
      Style           =   1  'Graphical
      TabIndex        =   64
      ToolTipText     =   "Previous"
      Top             =   1080
      Width           =   375
   End
   Begin VB.TextBox txt 
      Alignment       =   1  'Right Justify
      Height          =   315
      Index           =   4
      Left            =   18270
      TabIndex        =   74
      Text            =   "4"
      Top             =   1575
      Visible         =   0   'False
      Width           =   525
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
      Index           =   6
      Left            =   17325
      Locked          =   -1  'True
      TabIndex        =   41
      Text            =   "txttemp(6)"
      ToolTipText     =   "Press F3 for Creation"
      Top             =   6255
      Visible         =   0   'False
      Width           =   1050
   End
   Begin VB.TextBox txt 
      Height          =   315
      Index           =   6
      Left            =   18405
      TabIndex        =   73
      Text            =   "6"
      Top             =   2205
      Visible         =   0   'False
      Width           =   435
   End
   Begin VB.TextBox txt 
      Alignment       =   1  'Right Justify
      Height          =   315
      Index           =   7
      Left            =   18315
      TabIndex        =   72
      Text            =   "7"
      Top             =   2565
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
      Index           =   10
      Left            =   18360
      TabIndex        =   48
      Text            =   "10"
      Top             =   3240
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
      Index           =   11
      Left            =   18450
      TabIndex        =   53
      Text            =   "11"
      Top             =   3555
      Visible         =   0   'False
      Width           =   375
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
      Index           =   12
      Left            =   17775
      TabIndex        =   29
      Text            =   "12"
      Top             =   1215
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
      Height          =   360
      Index           =   23
      Left            =   19800
      TabIndex        =   35
      Text            =   "23"
      Top             =   6510
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
      Index           =   24
      Left            =   19800
      TabIndex        =   36
      Text            =   "24"
      Top             =   4485
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
      Index           =   20
      Left            =   18330
      TabIndex        =   42
      Text            =   "20"
      Top             =   8505
      Visible         =   0   'False
      Width           =   660
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
      Left            =   19800
      TabIndex        =   39
      Text            =   "30"
      Top             =   5715
      Visible         =   0   'False
      Width           =   330
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
      Index           =   31
      Left            =   15825
      TabIndex        =   22
      Text            =   "txt(31)"
      Top             =   7785
      Visible         =   0   'False
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
      Index           =   17
      Left            =   19800
      TabIndex        =   30
      Text            =   "17"
      Top             =   4080
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
      Index           =   18
      Left            =   18945
      TabIndex        =   33
      Text            =   "18"
      Top             =   2295
      Visible         =   0   'False
      Width           =   330
   End
   Begin MSMask.MaskEdBox mskDt 
      Height          =   375
      Index           =   2
      Left            =   17580
      TabIndex        =   19
      Top             =   -120
      Visible         =   0   'False
      Width           =   1005
      _ExtentX        =   1773
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
   Begin MSMask.MaskEdBox mskDt 
      Height          =   375
      Index           =   39
      Left            =   17775
      TabIndex        =   28
      Top             =   765
      Visible         =   0   'False
      Width           =   300
      _ExtentX        =   529
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
      Index           =   32
      Left            =   16635
      TabIndex        =   23
      Top             =   7800
      Visible         =   0   'False
      Width           =   1425
      _ExtentX        =   2514
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
      Bindings        =   "frmSalesMillBill.frx":058A
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
      Height          =   2850
      Left            =   315
      TabIndex        =   8
      Top             =   3435
      Width           =   12630
      _ExtentX        =   22278
      _ExtentY        =   5027
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
      Index           =   14
      Left            =   15885
      TabIndex        =   26
      Top             =   7110
      Visible         =   0   'False
      Width           =   1425
      _ExtentX        =   2514
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
   Begin MSMask.MaskEdBox mskDt 
      Height          =   375
      Index           =   10
      Left            =   17820
      TabIndex        =   20
      Top             =   2580
      Visible         =   0   'False
      Width           =   885
      _ExtentX        =   1561
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
      Left            =   10020
      TabIndex        =   5
      Top             =   2145
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
      Format          =   123338755
      CurrentDate     =   -273
   End
   Begin MSComCtl2.DTPicker dtpDt 
      Height          =   360
      Index           =   10
      Left            =   10020
      TabIndex        =   7
      Top             =   2655
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
      Format          =   123338755
      CurrentDate     =   -273
   End
   Begin MSComCtl2.DTPicker dtpDt 
      Height          =   360
      Index           =   33
      Left            =   10020
      TabIndex        =   3
      Top             =   840
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
      Format          =   123338755
      CurrentDate     =   -273
   End
   Begin VB.Label Label31 
      BackStyle       =   0  'Transparent
      Caption         =   "TDS On Amount :"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000040C0&
      Height          =   375
      Left            =   180
      TabIndex        =   153
      Top             =   7020
      Width           =   1890
   End
   Begin VB.Label Label16 
      BackStyle       =   0  'Transparent
      Caption         =   "Entry Dt.  : "
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
      Left            =   8970
      TabIndex        =   152
      Top             =   840
      Width           =   1095
   End
   Begin VB.Label Label13 
      BackStyle       =   0  'Transparent
      Caption         =   "Taxable Amount  :"
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
      Left            =   8280
      TabIndex        =   150
      Top             =   6960
      Width           =   1830
   End
   Begin VB.Label lblTcsLimit 
      BackColor       =   &H00FBF2E1&
      Caption         =   "This Is No Work Party"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   435
      Left            =   8040
      TabIndex        =   149
      Top             =   7980
      Width           =   6195
   End
   Begin VB.Label Label35 
      BackStyle       =   0  'Transparent
      Caption         =   "TDS Amt.           :    "
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000040C0&
      Height          =   375
      Left            =   180
      TabIndex        =   141
      Top             =   7500
      Width           =   1890
   End
   Begin VB.Label Label34 
      BackStyle       =   0  'Transparent
      Caption         =   "TDS Rate  :"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000040C0&
      Height          =   375
      Left            =   4020
      TabIndex        =   140
      Top             =   7020
      Width           =   1110
   End
   Begin VB.Label Label37 
      BackStyle       =   0  'Transparent
      Caption         =   "TDS A/c            :"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000040C0&
      Height          =   375
      Left            =   240
      TabIndex        =   139
      Top             =   8055
      Width           =   1890
   End
   Begin VB.Label lblMblNo 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "MBL No :  "
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
      Left            =   2430
      TabIndex        =   119
      Top             =   2160
      Width           =   975
   End
   Begin VB.Label lblMblNoParty 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Mobile No"
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
      Left            =   3420
      TabIndex        =   118
      Top             =   2160
      Width           =   4740
   End
   Begin VB.Label Label11 
      BackStyle       =   0  'Transparent
      Caption         =   "Date       : "
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
      Left            =   8970
      TabIndex        =   117
      Top             =   2700
      Width           =   1095
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "Date       : "
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
      Left            =   8970
      TabIndex        =   116
      Top             =   2138
      Width           =   1095
   End
   Begin VB.Label lblSr 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Sr."
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   285
      Left            =   10020
      TabIndex        =   115
      Top             =   1305
      Width           =   1575
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Bill Chr.  :"
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
      Left            =   8970
      TabIndex        =   114
      Top             =   1305
      Width           =   945
   End
   Begin VB.Label lblBalWt 
      BackStyle       =   0  'Transparent
      Caption         =   "Balance Weight"
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
      Left            =   16080
      TabIndex        =   113
      Top             =   4320
      Visible         =   0   'False
      Width           =   1140
   End
   Begin VB.Label lblWeight 
      BackStyle       =   0  'Transparent
      Caption         =   "Wt.       : "
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
      Height          =   375
      Left            =   15090
      TabIndex        =   111
      Top             =   4320
      Visible         =   0   'False
      Width           =   960
   End
   Begin VB.Label lblBalBag 
      BackStyle       =   0  'Transparent
      Caption         =   "Balance Bag"
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
      Left            =   15780
      TabIndex        =   109
      Top             =   3435
      Visible         =   0   'False
      Width           =   1050
   End
   Begin VB.Label lblBag 
      BackStyle       =   0  'Transparent
      Caption         =   " Bag     : "
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
      Height          =   375
      Left            =   14745
      TabIndex        =   107
      Top             =   3435
      Visible         =   0   'False
      Width           =   960
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Cr. Days  :"
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
      Height          =   255
      Left            =   12000
      TabIndex        =   105
      Top             =   1800
      Width           =   1170
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
      Left            =   7890
      TabIndex        =   96
      Top             =   1845
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
      Left            =   3420
      TabIndex        =   93
      Top             =   1845
      Width           =   2775
   End
   Begin VB.Label Label21 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "TIN        :  "
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
      Left            =   2430
      TabIndex        =   92
      Top             =   1845
      Width           =   975
   End
   Begin VB.Label Label14 
      BackStyle       =   0  'Transparent
      Caption         =   "L.R. No    :    "
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
      Height          =   255
      Left            =   14445
      TabIndex        =   89
      Top             =   7245
      Visible         =   0   'False
      Width           =   690
   End
   Begin VB.Label Label8 
      BackStyle       =   0  'Transparent
      Caption         =   "Transport  :    "
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
      Height          =   255
      Left            =   14625
      TabIndex        =   88
      Top             =   6615
      Visible         =   0   'False
      Width           =   1320
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Bill No    :"
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
      Left            =   8970
      TabIndex        =   87
      Top             =   1755
      Width           =   930
   End
   Begin VB.Label Label15 
      BackStyle       =   0  'Transparent
      Caption         =   "D&ue Days    :"
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
      Height          =   255
      Left            =   14580
      TabIndex        =   86
      Top             =   7830
      Visible         =   0   'False
      Width           =   1320
   End
   Begin VB.Shape Shape5 
      BorderColor     =   &H00EFB17A&
      BorderWidth     =   5
      FillColor       =   &H00EFB17A&
      Height          =   8820
      Left            =   45
      Top             =   -210
      Width           =   14460
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "&Party   :"
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
      Left            =   495
      TabIndex        =   43
      Top             =   1425
      Width           =   900
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "M&ill     :"
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
      Left            =   495
      TabIndex        =   45
      Top             =   900
      Width           =   960
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "&Broker  : "
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
      Left            =   495
      TabIndex        =   44
      Top             =   2520
      Width           =   870
   End
   Begin VB.Label Label9 
      BackStyle       =   0  'Transparent
      Caption         =   "Sales  A/C   :"
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
      Left            =   17235
      TabIndex        =   0
      Top             =   6660
      Visible         =   0   'False
      Width           =   1620
   End
   Begin VB.Label LblExmptKg 
      BackStyle       =   0  'Transparent
      Caption         =   "Exem/Kg.      :"
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
      Left            =   17370
      TabIndex        =   54
      Top             =   8505
      Width           =   1335
   End
   Begin VB.Label Label10 
      BackStyle       =   0  'Transparent
      Caption         =   "Ex. Mill Amount    :"
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
      Left            =   8280
      TabIndex        =   47
      Top             =   6495
      Width           =   1830
   End
   Begin VB.Label lblCharityPer 
      BackStyle       =   0  'Transparent
      Caption         =   " Kg/Pkg       "
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
      Left            =   8460
      TabIndex        =   49
      Top             =   9585
      Visible         =   0   'False
      Width           =   750
   End
   Begin VB.Label Label30 
      BackStyle       =   0  'Transparent
      Caption         =   "Bill Amount          :"
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
      Left            =   8280
      TabIndex        =   52
      Top             =   7425
      Width           =   1905
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
      Left            =   11610
      TabIndex        =   80
      Top             =   0
      Width           =   2475
   End
   Begin VB.Label LblExmptAmt 
      BackStyle       =   0  'Transparent
      Caption         =   "&Exempt. Amt :"
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
      Height          =   300
      Left            =   17190
      TabIndex        =   50
      Top             =   8190
      Width           =   1440
   End
   Begin VB.Label lblTaxType 
      BackStyle       =   0  'Transparent
      Caption         =   "&TaxType       : "
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
      Left            =   13185
      TabIndex        =   51
      Top             =   7650
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.Label LblSaleType 
      BackColor       =   &H00EFB17A&
      Caption         =   "     Sales ( Mill Bill )"
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
      TabIndex        =   78
      Top             =   0
      Width           =   14445
   End
   Begin VB.Shape Shape1 
      FillColor       =   &H00FED8FC&
      FillStyle       =   0  'Solid
      Height          =   6195
      Left            =   13050
      Shape           =   4  'Rounded Rectangle
      Top             =   675
      Width           =   1215
   End
   Begin VB.Shape Shape2 
      FillColor       =   &H00EDFDFE&
      FillStyle       =   0  'Solid
      Height          =   2490
      Left            =   315
      Top             =   630
      Width           =   8445
   End
   Begin VB.Shape Shape3 
      FillColor       =   &H00EDFDFE&
      FillStyle       =   0  'Solid
      Height          =   2490
      Left            =   8850
      Top             =   630
      Width           =   4095
   End
End
Attribute VB_Name = "frmSalesMillBill"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Const tblName As String = "tblSale"
Const IdField As String = "Vno"
Const OrderField As String = "Vdt,Vno"
'Const MaxNo As Long = 44
Const MaxNo As Long = 74
Const OtherTxtFromNo As Long = 0 '-- Other Txt Ctrl From No
Const OtherTxtToNo As Long = 0 '-- Other Txt Ctrl Less No
Dim rstbl As Recordset '-- TblSale
Dim rstblSub As Recordset '-- TblSaleSub
Dim varBookMark As Variant
Dim FirstTimeFlag As Boolean
Dim OldgFrmTypeStr As String
Dim FormAction As Integer
Dim frmTypeStr As String
Dim i As Long
Dim SalesAccode As Long
Dim Criteria As String
Dim TotWt As Double
Dim Mon As Integer
Dim MsOms As String
Dim X As String
Dim tmpRst As Recordset
Dim OldBag As Integer
Dim mLastEntryDate As Date
Dim mLastEntryBillNo As Long
Dim mLastEntryCrDays As Integer
Dim mLastEntryMill As String
Private Sub Cbo_KeyPress(Index As Integer, KeyAscii As Integer)
If Index = 22 Then KeyAscii = 0 '--Tax
End Sub
Private Sub cbo_LostFocus(Index As Integer)
If txttemp(8) = "" Then
   txttemp(8).SetFocus
End If
If txttemp(5) = "" Then
   txttemp(5).SetFocus
End If

'If Index = 22 And FormAction = vbDataActionAddNew Then '--Tax %
'   Dim wMillCode As String
'   Dim VATRtFromItem As Double
'   wMillCode = GProcGetColumnValue("tblMastAccount", "AcName", txttemp(7), "S", "AcCode", "N")
'   VATRtFromItem = GProcGetColumnValue("tblMastItem", "ItName", msGrid.TextMatrix(1, 3), "S", "ITTMP1", "N", "ItMillcode=" & wMillCode) '--Item Code
'   If VATRtFromItem > 0 Then
'      txt(23) = VATRtFromItem
'   Else
'       txt(23) = GProcGetColumnValue("tblMastNarration", "Narration", Cbo(22), "S", "MastTaxRate", "N")
'   End If
'   txt(23) = Format(txt(23), FStr)
'
'End If
End Sub
Private Sub chkIsPaymenttoMill_Click()
CalAmount
End Sub

'-- Cartoon Bag Detail Entry
Private Sub CmdBagDetails_Click()
'gfrmBagInwVNo = txt(0)
'gfrmBagInwVType = frmTypeStr
'frmBagEntry.Show
End Sub
'Dim GridFlag As Boolean
Private Sub cmdBtn_Click(Index As Integer)
'--for Log
Dim j As Long
Dim tmpRst As Recordset
Dim wLogNo As Integer
Dim AMD As String

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
    gClsSearch.SearchMultiField "tblSale,TblMastAccount", "vno,SlBillNo,Vdt,AcName", Array("Vno", "Bill No", "Vdt", "Party"), Array(0, 2500, 2000, 5000), " tblSale.SlAcDrCode=TblMastAccount.AcCode and vtype='" & frmTypeStr & "'  and VFirm = '" & gCCode & "'   and VYear=" & gCYear, "", "SlBillNo,Vdt", 0, 0
     If gClsSearch.SearchMultiRetCol(0) <> "" Then
        GProcRstOpen rstbl, "Select * from " & tblName & " where " & IdField & " = " & gClsSearch.SearchMultiRetCol(0) & " and " & Criteria, "R"
        gClsSearch.SearchMultiRetCol(1) = ""
        ReadFields
    End If
Case 6 'Add
    FormAction = vbDataActionAddNew
    Call GProcActivateControls(Me, True)
    Call GProcSetButtons(Me, vbDataActionAddNew)
    Call GProcClearForm(Me, rstbl, MaxNo, True, OtherTxtFromNo, OtherTxtToNo, True)
    SetControlEd
    
'    X = GProcGetColumnValue("tblMastAccount", "AcMillBillFirm", "" & gCCode & "", "S", "AcMillBillFirm", "S")
'    If X = gCCode Then
'       X = GProcGetColumnValue("tblMastAccount", "AcMillBillFirm", "" & gCCode & "", "S", "Acname", "S")
'           txttemp(7) = X
'       X = GProcGetColumnValue("tblMastAccount", "AcMillBillFirm", "" & gCCode & "", "S", "AcOurCode", "S")
'           txtCode(7) = X
'       SalesAccode = GProcGetColumnValue("tblMastAccount", "AcMillBillFirm", "" & gCCode & "", "S", "AcCode", "S")
'       txtCode(7).Enabled = False
'       txttemp(7).Enabled = False
'    End If
    dtpDt(33) = Format(Now, "dd/MM/yyyy")
    If Year(mLastEntryDate) = 1899 Then
        dtpDt(2) = Format(Now, "dd/MM/yyyy")
    Else
        dtpDt(2) = mLastEntryDate
        txtCode(7) = mLastEntryMill
        txt(9) = mLastEntryBillNo + 1
        txt(45) = mLastEntryCrDays
        
    End If
    If FirstTimeFlag = True Then
        FirstTimeFlag = False
    Else
        txtCode(7).SetFocus
'        txt(9).SetFocus
    End If
'    txt(0) = GProcGenerateId("TblPurch", "Vno", "Vtype='" & frmTypeStr & "' and VFirm = '" & gCCode & "'   and VYear=" & gCYear)
'    dtpDt(2) = GProcGenerateVDt("tblSale", "Vdt", "Vtype='" & frmTypeStr & "' and VFirm = '" & gCCode & "'   and VYear=" & gCYear)
'    dtpDt(2) = Format(Now, "dd/MM/yyyy")
    '-- Sales A/C
'    txttemp(6) = GProcGetColumnValue("tblMastAccount", "AcCode", CStr(SalesAccode), "N", "AcName", "S")
 '   txttemp(7) = txttemp(6).text
   txt(45) = 20
  
 
    FillCombo
    SetGrid
    lblBag.Visible = False
    lblWeight.Visible = False
    lblBalBag.Visible = False
    lblWeight.Visible = False
    lblBalBag.Caption = ""
    lblBalWt.Caption = ""
    lblTinNo.Caption = ""
    lblMsOms.Caption = ""
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
             '--- If Gate Pass Set Ctrl Enable
'            If IsRelatedGatePass = True Then
'                 Call GProcActivateControls(Me, False)
'                 msGrid.Enabled = True '-Grid
'                 txtVou(11).Enabled = True '--Bill No
'                 mskVou(12).Enabled = True '--Bill Dt
'                 For i = 25 To 30 '-Landing Cost
'                    txt(i).Enabled = True
'                 Next i
'                 txt(15).Enabled = True '-Ex amt
'                 txt(24).Enabled = True '-Ass value
'                 Cbo(12).Enabled = True '-Tax Type
'                 txt(13).Enabled = True '-Tax %
'                 txt(16).Enabled = True '-Ex /Kg
'                 txtOS(9).Enabled = True '-Due days
'                ' txt(21).Enabled = True '-Tot amt
'                 txt(18).Enabled = True 'add
'                 txt(19).Enabled = True 'Less
'                 txt(14).Enabled = True 'Tax amt
'                 txt(17).Enabled = True '-Roff
'                 SetControlEd
'            End If
        Else
            GProcRstOpen rstbl, "Select * from " & tblName & " where " & IdField & " = " & txt(0) & " and " & Criteria, "R"
            ReadFields
            FormAction = vbDataActionUpdate
            Call GProcActivateControls(Me, False)
            SetControlEd
            Call GProcSetButtons(Me, vbDataActionUpdate)
        
        End If
        txt(15).Enabled = True
        dtpDt(2).Enabled = True
        chkIsPaymenttoMill.Enabled = True
    'End If
Case 8 'Delete
    'If mURecDel = True Then
      If IsRelatedGatePass = True Then
            MsgBox "You can not Modify/Delete this record.Related Gate Pass is exists.", vbCritical + vbOKOnly, App.Title
            Exit Sub
      End If
      If IsRelatedRecord = False Then
         If MsgBox("Do You Really Delete This Record?", vbQuestion + vbYesNo + vbDefaultButton1, "Confirm Delete") = vbYes Then
            If gCYear >= 2023 And gCIsLog = 1 Then
               gCn.BeginTrans
               X = "Select max(LogNo) from tblSale_Log where Vno=" & Val(txt(0)) & " and " & Criteria & ""
               j = GProcRstOpen(tmpRst, X, "R")
               If j > 0 Then
                  wLogNo = IIf(IsNull(tmpRst.Fields(0)), 0, tmpRst.Fields(0)) + 1
               Else
                  wLogNo = 1
               End If
               gCn.Execute "Insert into tblSale_Log select tblSale.*," & wLogNo & " as LogNo,'D' as LogTp ,'" & gUserName & "' as UserName ,'" & Format(Now, "YYYY/MM/DD hh:mm:ss") & "' as LogDate from tblSale where Vno=" & Val(txt(0)) & " and " & Criteria & ""
               gCn.Execute "Insert into tblSaleSub_Log select tblSaleSub.*," & wLogNo & " as LogNo,'D' as LogTp ,'" & gUserName & "' as UserName ,'" & Format(Now, "YYYY/MM/DD hh:mm:ss") & "' as LogDate from tblSaleSub where Vno=" & Val(txt(0)) & " and " & Criteria & ""
               gCn.CommitTrans
               tmpRst.Close
           End If
           DeleteRelatedRecord
           GProcRstOpen rstbl, "Select * from " & tblName & " where " & IdField & " = " & txt(0) & " and " & Criteria, "O"
           rstbl.Delete
           cmdBtn_Click (3)
        End If
    End If
Case 9 'Print
Case 10 'Exit
'    DoEvents
    Unload Me
    Exit Sub
Case 11 'Label
    'frmLabel.Show 1
Case 12 'Optional
Case 13 'Save
        'txt(51) = CDate(Date)
        txttemp(6) = txttemp(7).text
     If ValidateData = True Then
        If Trim(txtGrid) <> "" Then
            msGrid.text = txtGrid
            DoEvents
            txtGrid = ""
        End If
        txtGrid.Visible = False
       ' If FormAction = vbDataActionAddNew Then txt(0) = GProcGenerateId(tblName, IdField, "vtype='" & FrmTypeStr & "' and VYear=" & gCYear)
        CalAmount
        FillTxtFromTemp
        '-- Save Record
        If FormAction = vbDataActionAddNew Then
           Mon = Month(dtpDt(2))
           txt(0) = GProcGenerateIdMonthwise("TblSale", "Vno", Mon, "Vdt", "Vtype='" & frmTypeStr & "' and VFirm = '" & gCCode & "'   and VYear=" & gCYear)
        End If
        gCn.BeginTrans
        GProcRstOpen rstbl, "Select * from " & tblName & " where " & IdField & " = " & txt(0) & " and " & Criteria, "O"
        GProcSaveRecord Me, rstbl, FormAction, MaxNo, OtherTxtFromNo, OtherTxtToNo
        SaveRelatedRecord
        gCn.CommitTrans
        '-- For Log
        If gCYear >= 2023 And gCIsLog = 1 Then
           If FormAction = vbDataActionAddNew Then
              AMD = "A"
           Else
              AMD = "M"
           End If
           X = "Select max(LogNo) from tblSale_Log where Vno=" & Val(txt(0)) & " and " & Criteria & ""
           j = GProcRstOpen(tmpRst, X, "R")
           If j > 0 Then
              wLogNo = IIf(IsNull(tmpRst.Fields(0)), 0, tmpRst.Fields(0)) + 1
           Else
              wLogNo = 1
           End If
           gCn.Execute "Insert into tblSale_Log select tblSale.*," & wLogNo & " as LogNo,'" & AMD & "' as LogTp ,'" & gUserName & "' as UserName ,'" & Format(Now, "YYYY/MM/DD hh:mm:ss") & "' as LogDate from tblSale where Vno=" & Val(txt(0)) & " and " & Criteria & ""
           gCn.Execute "Insert into tblSaleSub_Log select tblSaleSub.*," & wLogNo & " as LogNo,'" & AMD & "' as LogTp ,'" & gUserName & "' as UserName ,'" & Format(Now, "YYYY/MM/DD hh:mm:ss") & "' as LogDate from tblSaleSub where Vno=" & Val(txt(0)) & " and " & Criteria & ""
           tmpRst.Close
        End If
        mLastEntryDate = dtpDt(2)
        mLastEntryMill = txtCode(7)
        mLastEntryBillNo = Val(Mid(txt(9), 5, 15))
        mLastEntryCrDays = Val(txt(45))
        Call GProcActivateControls(Me, False)
        Call GProcSetButtons(Me, True)
        ReadFields
'        If FormAction = vbDataActionAddNew Then
            If cmdBtn(6).Visible = True Then cmdBtn(6).SetFocus
 '       Else
  '          If cmdBtn(7).Visible = True And cmdBtn(7).Enabled = True Then
   '            cmdBtn(7).SetFocus
    '        Else
     '          cmdBtn(10).SetFocus
      '      End If
       ' End If
        FormAction = vbDataActionClose
    End If
Case 14 'Cancel
    cmdBtn_Click (3)
    txtGrid.Visible = False
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
    msGrid.Col = 0
    msGrid.Row = 1
    txtGrid.Move msGrid.CellLeft + msGrid.Left, msGrid.CellTop + msGrid.Top, msGrid.CellWidth, msGrid.CellHeight
End Select
Exit Sub
ErrorRoutine:
    rstbl.CancelUpdate
    GProcErrorHandler
End Sub
Private Sub SetControlEd()
'txttemp(6).Enabled = False '----Sale A/c
'txttemp(7).Enabled = False '----Mill Name
'txtCode(7).Enabled = False
mskDt(32).Enabled = False
dtpDt(10).Enabled = False
If FormAction = vbDataActionAddNew Then
   txt(0).Enabled = False
End If
If Trim(txtCode(5)) <> "" Then
   txttemp(5).Enabled = False
Else
   txttemp(5).Enabled = True
End If
If Trim(txtCode(8)) <> "" Then
   txttemp(8).Enabled = False
Else
   txttemp(8).Enabled = True
End If
i = GProcGetColumnValue("tblMastAccount", "AcName", txttemp(7), "S", "AcComOn", "N")
If i = 0 Then
   txt(15).Visible = False
   Label10.Visible = False
Else
   txt(15).Visible = True
   Label10.Visible = True
End If


End Sub
Private Sub DeleteRelatedRecord()
With msGrid
    i = 1
    Do While i < .rows
       gCn.Execute "update tblBooKingDesp set BkDespTmpNo1 = 0 " _
                    & " where  tblBooKingDesp.Vno = " & Val(.TextMatrix(i, 1)) & "    and  tblBooKingDesp.Vtype =  '" & .TextMatrix(i, 7) & "'  and tblBooKingDesp.vyear = " & Val(.TextMatrix(i, 2)) & "   and tblBooKingDesp.Vfirm = '" & gCCode & "' and tblBooKingDesp.BkSrNo = " & Val(.TextMatrix(i, 9)) & ""
        i = i + 1
    Loop
 End With

''--- Voucher Sub
gCn.Execute "delete from tblSaleSub where vtype='" & frmTypeStr & "' and vno=" & txt(0) & " and Vfirm = '" & gCCode & "'  and " _
& " VYear=" & gCYear

''-- OutStanding
'gCn.Execute "delete from tblOutStanding where vtype='" & frmTypeStr & "' and vno=" & txt(0) & " and " _
'& " VYear=" & gCYear
''-- Voucher
'gCn.Execute "delete from tblVoucher where vtype='" & frmTypeStr & "' and vno=" & txt(0) & " and " _
'& " VYear=" & gCYear
'--- JV
'gCn.Execute "delete from tblVoucher where vtype='JV' and vno=" & txt(20) & " and " _
'& " VYear=" & gCYear
End Sub
Private Sub SaveRelatedRecord()
Dim i As Long
Dim MillCode As Long
Dim rstSub As Recordset
Dim CrCode As Long
Dim DrCode As Long
Dim BrkCode As Long
Dim TaxAcCode As Long
'Dim X As String
Dim BillDt
Dim DueDt
'--- Account Code
MillCode = GProcGetColumnValue("tblMastAccount", "AcName", txttemp(7), "S", "AcCode", "N")
DrCode = GProcGetColumnValue("tblMastAccount", "AcName", txttemp(5), "S", "AcCode", "N")
CrCode = GProcGetColumnValue("tblMastAccount", "AcName", txttemp(6), "S", "AcCode", "N")
BrkCode = GProcGetColumnValue("tblMastAccount", "AcName", txttemp(8), "S", "AcCode", "N")
TaxAcCode = GProcGetColumnValue("tblMastNarration", "Narration", Cbo(22), "S", "MastTaxSAAcCode", "N")
'-- Delete Records
DeleteRelatedRecord
'--- Save Grid Records
'--- TblVouSub
X = "select * from tblSaleSub where vtype='" & frmTypeStr & "' and vno=" & txt(0) & " and Vfirm = '" & gCCode & "' and " _
& " VYear=" & gCYear
GProcRstOpen rstSub, X, "O"
With msGrid
    i = 1
    Do While i < .rows
        rstSub.AddNew
        rstSub!Vno = CLng(txt(0)) '--Vno(0)
        rstSub!Vtype = txt(1) & vbNullString  '--Vtype(1)
        rstSub!Vdt = CDate(txt(2)) '--Vdt(2)
        rstSub!VYear = txt(3) '--VYear
        rstSub!VFirm = gCCode '--VFirm
        rstSub!SlSubItSrNo = i '--Sr No
        rstSub!SlSubBookNo = CLng(.TextMatrix(i, 1)) & vbNullString  '--Book No
        If IsDate(.TextMatrix(i, 3)) Then
           rstSub!SlSubBookDt = CDate(.TextMatrix(i, 3))
        Else
            rstSub!SlSubBookDt = Null
        End If
        '--Book Dt
        If CLng(.TextMatrix(i, 1)) = 0 Then '-- if Book No =0
            .TextMatrix(i, 3) = "" '--Book Dt =Null
        End If
        '--Itcode(4)
        rstSub!SlSubItCode = GProcGetColumnValue("tblMastItem", "ItName", .TextMatrix(i, 4), "S", "ItCode", "N", " ItMillCode=" & MillCode) '--Item Code
        rstSub!SlSubBag = CLng(.TextMatrix(i, 5)) '--Bag(4)
        rstSub!SlSubwt = CDbl(.TextMatrix(i, 6)) '-- Wt(5)
        rstSub!SlSubBkType = .TextMatrix(i, 7) & vbNullString
        rstSub!SlSubBkVyear = CLng(.TextMatrix(i, 2))
        rstSub!SlSubBkSrNo = CLng(.TextMatrix(i, 8))
        rstSub!SlSubBkDespSrNo = CLng(.TextMatrix(i, 9))
        '--Godown Code(4)
        rstSub!SlSubGodown = GProcGetColumnValue("tblMastNarration", "Narration", "-", "S", "NarrCode", "N", "")
        '-- P It Ctrl No =Booking It ctrl No
           
         gCn.Execute "update tblBooKingDesp set BkDespTmpNo1 = " & Val(.TextMatrix(i, 5)) & " " _
                               & " where  tblBooKingDesp.Vno = " & Val(.TextMatrix(i, 1)) & "    and  tblBooKingDesp.Vtype =  '" & .TextMatrix(i, 7) & "'  and tblBooKingDesp.vyear = " & Val(.TextMatrix(i, 2)) & "   and tblBooKingDesp.Vfirm = '" & gCCode & "' and tblBooKingDesp.BkSrNo = " & Val(.TextMatrix(i, 9)) & ""
        
        
        rstSub.Update
        i = i + 1
     Loop
 End With
 rstSub.Close
End Sub

Private Sub CmdClose_Click()
    DoEvents
    Frame2.Visible = False
    txtCode(8).SetFocus
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case 13:
            If Me.ActiveControl.Name <> "txtGrid" Then Sendkeys "{TAB}"
        Case 27:
            If (cmdBtn(13).Enabled = False) And TotWt = 0 Then
                Unload Me
            End If
    End Select
End Sub
Private Sub Form_Load()
'    Dim X As String
    Me.Top = 300
    Me.Left = Screen.Width / 2 - Me.Width / 2
    '--- TblAddLess
    frmTypeStr = gfrmTypeStr
    mLastEntryDate = vbNull
    SalesAccode = 0
    FirstTimeFlag = True

'    If Trim(mShowRec) <> "" Then
'       SalesAccode = mShowRecSaleAc
'    Else
'       SalesAccode = gClsSearch.SearchMultiRetCol(2)
'    End If
'    txtCode(7) = gClsSearch.SearchMultiRetCol(3)

'    X = "Select * from " & tblName & " where SlAcCrCode=" & SalesAccode & " and     Vtype='" & frmTypeStr & "'  and SlIsHank = 0  and VFirm  = '" & gCCode & "' and VYear=" & gCYear
'    If Trim(OrderField) <> "" Then X = X & " order by " & OrderField
'    GProcRstOpen rstbl, X, "O"
    
    Criteria = "Vtype='" & frmTypeStr & "' and VFirm  = '" & gCCode & "' and VYear=" & gCYear & ""
    If Trim(mShowRec) <> "" Then
       GProcRstOpen rstbl, "Select * from " & tblName & " where  " & Criteria & " and  " & IdField & " = " & CLng(Trim(mShowRec)), "R"
       ReadFields
       Call GProcSetButtons(Me, vbDataActionCancel)
       mShowRec = ""
    Else
       GProcRstOpen rstbl, "Select * from " & tblName & " where " & Criteria & " and  " & IdField & " = 0", "R"
       SetGrid
       SetRelatedRS
       cmdBtn_Click (6)
    End If
'
'
'
'
'    Criteria = "SlAcCrCode=" & SalesAccode & " and  Vtype='" & frmTypeStr & "'  and SlIsHank = 0  and VFirm  = '" & gCCode & "' and VYear=" & gCYear & ""
'    If Trim(mShowRec) <> "" Then
'       GProcRstOpen rstbl, "Select * from " & tblName & " where  " & Criteria & " and  " & IdField & " = " & CLng(Trim(mShowRec)), "R"
'       ReadFields
'       Call GProcSetButtons(Me, vbDataActionCancel)
'       mShowRec = ""
'    Else
'       GProcRstOpen rstbl, "Select * from " & tblName & " where " & Criteria & " and  " & IdField & " = 0", "R"
'       SetGrid
'       SetRelatedRS
'       cmdBtn_Click (6)
'    End If
    '---Grid
    lblSr.Caption = GProcGetColumnValue("tblMastAccount", "AcCode", CLng(SalesAccode), "N", "AcBillSrNo", "S")
    '---- Button Visiblity Set As Per User Previlage
    GProcSetButtonVisiblity Me
End Sub
Private Sub SetRelatedRS()
lblBag.Visible = False
lblWeight.Visible = False
lblBalBag.Visible = False
lblWeight.Visible = False
lblBalBag.Caption = ""
lblBalWt.Caption = ""



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
'Dim X As String
Set rstblSub = New Recordset
X = "select SlSubItSrNo,SlSubBookNo,SlSubBkVyear,SlSubBookDt,ItName,SlSubBag,SlSubWt" _
& " ,SlSubBkType,SlSubBkSrNo,SlSubBkDespSrNo from tblSaleSub,TblMastItem where " _
& " tblSaleSub.SlSubItCode=TblMastItem.ItCode " _
& " and tblSaleSub.vtype='" & frmTypeStr & "' and tblSaleSub.vno=" & txt(0) & " and tblSaleSub.VFirm = '" & gCCode & "' and tblSaleSub.VYear=" & gCYear & " order by SlSubItSrNo"
GProcRstOpen rstblSub, X, "R"
With msGrid
   .FormatString = ">Sr|>Bk.No         |>Bk.Year       |<Book Date    |<Count                                                             |>Bag           |>Weight          |||"
   .ColWidth(7) = 0
   .ColWidth(8) = 0
   .ColWidth(9) = 0
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
SetControlEd
SaveBtnEd
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
    If IsDate(mskDt(39)) = False Then
       mskDt(39) = dtpDt(2)
    End If
End If
'If Index = 2 Then ShowAccountBalance
'If IsDate(mskDt(Index)) = False Then mskDt(Index) = "__/__/____"
End Sub
Private Sub ShowAccountBalance()
Dim DrAcBal  As Double
Dim CrAcBal As Double
If IsDate(dtpDt(2)) Then
    DrAcBal = GProcGetAccountBalance(txttemp(5), dtpDt(2))
'    CrAcBal = GProcGetAccountBalance(txttemp(5), dtpDt(2))
    txtDrBalance = IIf(DrAcBal > 0, DrAcBal & " Cr", Abs(DrAcBal) & " Dr")
'    txtCrBalance = IIf(CrAcBal > 0, CrAcBal & " Cr", Abs(CrAcBal) & " Dr")
End If
End Sub
Private Sub txt_GotFocus(Index As Integer)
GProcSelectBox Me.ActiveControl
End Sub
Private Sub CheckForNumber()
GProcCheckForNumber Me, rstbl, MaxNo, Array(9), pOtherTxtFromNo:=OtherTxtFromNo, pOtherTxtToNo:=OtherTxtToNo
End Sub
Private Sub txt_LostFocus(Index As Integer)
If Index = 9 And Trim(txt(9)) <> "" And FormAction <> vbDataActionCancel Then
   Dim rstAddCheck As Recordset
   i = GProcGetColumnValue("tblMastAccount", "AcName", txttemp(7), "S", "AcCode", "N")
   X = "Select vno From " & tblName & " Where SlBillNo = '" & lblSr + "/" + txt(9) & "' and SlAcCrCode = " & i & "  and  vno <> " & CLng(txt(0)) & "  and Vtype='" & frmTypeStr & "'  and Vfirm = '" & gCCode & "'  and VYear=" & gCYear
        i = GProcRstOpen(rstAddCheck, X, "R")
        If i > 0 Then
            MsgBox "Same Bill No of Same Mill is Available in Voucher No. " & rstAddCheck.Fields(0), vbCritical + vbOKOnly, "Duplicate Record"
            dtpDt(2).SetFocus
        End If
End If

If Index = 45 Then
   msGrid.Col = 1
   msGrid.SetFocus
   msGrid_EnterCell
End If
CheckForNumber
CalAmount
SaveBtnEd
End Sub
Private Sub txt_KeyPress(Index As Integer, KeyAscii As Integer)
If KeyAscii = 13 Then Exit Sub
KeyAscii = GProcValidateKey(rstbl, Index, KeyAscii, txt(Index).text)
End Sub
Private Sub txtCode_LostFocus(Index As Integer)
If Trim(txtCode(Index)) = "" Then Exit Sub
Dim BrkAcCode As Long
Dim wIsNoWork As Integer
'Dim X As String
SetControlEd
Select Case Index
    Case 5  '-- Party  Code
         lblMblNoParty.Caption = ""
         txttemp(5) = GProcGetColumnValue("tblMastAccount", "AcOurCode", UCase(txtCode(5)), "S", "AcName", "S")
         If txttemp(5) <> "" Then
            wIsNoWork = GProcGetColumnValue("tblMastAccount", "AcOurCode", UCase(txtCode(5)), "S", "AcIsNonWork", "N")
            If wIsNoWork Then
               Frame2.Visible = True
               cmdClose.SetFocus
            Else
                lblMblNoParty.Caption = GProcGetColumnValue("TblMastAccount", "ACName", txttemp(Index), "S", "AcContNo", "S")
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
                X = "Select VFirmName,VBillNo,VBillDt,Amt1,Amt3,Amt2,amt1+amt3-Amt2 as Os,Amt5 from TmpGentbl where UserName = '" & gUserName & "' and amt5 > 60 and amt1+amt3-Amt2 > amt3  "
                i = GProcRstOpen(tmpRst, X, "R")
                If i > 0 Then
                   If MsgBox("Total O/s Bills Over 60 days are " + Str(i) + " Do You Want To See O/s List ? ", vbYesNo) = vbYes Then
                      gClsSearch.SearchMultiField X, "VFirmName,VBillNo,VBillDt,Amt1,Amt3,Amt2,os,Amt5", Array("Firm Name", "Bill No", "Date", "Inv.Amt", "Debit Amt", "Recd.Amt", "O/s Amt", "L.Days"), Array(2000, 2000, 1800, 1500, 1500, 1500, 1500, 1000), "", "", "VFirmName,VBillDt,VBillNo ", 0, 0, False, , True
                   End If
                End If
            End If
         End If
         '--- Broker From Master
         If txttemp(8) = "" Then
            BrkAcCode = GProcGetColumnValue("TblMastAccount", "ACName", txttemp(Index), "S", "AcBrkCode", "N")
            txttemp(8) = GProcGetColumnValue("TblMastAccount", "ACCode", CStr(BrkAcCode), "N", "ACname", "S")
            txtCode(8) = GProcGetColumnValue("TblMastAccount", "ACCode", CStr(BrkAcCode), "N", "AcOurCode", "S")
         End If
         If Trim(txttemp(5)) = "" Then
            txttemp(5).Enabled = True
            txttemp(5).SetFocus
         End If
    Case 7  '-- Mill  Code
         Dim MilTpCd As Long
'         Dim tmpRst As Recordset
'         Dim X As String
         MilTpCd = GProcGetColumnValue("tblMastNarration", "Narration", "Mill Bill", "S", "Narrcode", "N")
         X = "Select Acname from tblMastAccount,tblMastBillSerial where tblMastaccount.Accode = tblMastBillSerial.MillCode  and tblMastBillSerial.Vfirm = '" & gCCode & "'   and  tblMastaccount.AcOurCode = '" & txtCode(7) & "'   and  tblMastBillSerial.SaleType = " & MilTpCd & "    "
          i = GProcRstOpen(tmpRst, X, "R")
          If i = 0 Then
             txttemp(7) = ""
              MsgBox ("Mill Not in Firm")
              txtCode(7).SetFocus
              Exit Sub
          End If
         txttemp(7) = GProcGetColumnValue("tblMastAccount", "AcOurCode", UCase(txtCode(7)), "S", "AcName", "S")
         MillCtrlEd
         GBillChr
    Case 8  '-- Broker  Code
         txttemp(8) = GProcGetColumnValue("tblMastAccount", "AcOurCode", UCase(txtCode(8)), "S", "AcName", "S")
         If Trim(txttemp(8)) = "" Then
            txttemp(8).Enabled = True
            txttemp(8).SetFocus
         End If
End Select
End Sub

Private Sub txttemp_GotFocus(Index As Integer)
Dim AcCode As Long
GProcSelectBox Me.ActiveControl
End Sub
Private Sub txttemp_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
Dim GName As String
Dim OldFrmType As String
If KeyCode = vbKeyF3 Then
    Select Case Index
        Case 5 '-- Party
            GName = GProcGetColumnValue("TblMastGroup", "AgCode", 90017, "N", "AgName", "S")
            GProcShowForm frmMastAccount, frmMain.mnuMstfrm(0), Array(frmMastAccount.txttemp(2)), Array(GName)
        Case 6 '-- Sales A/c
            GName = GProcGetColumnValue("TblMastGroup", "AgCode", 90023, "N", "AgName", "S")
            GProcShowForm frmMastAccount, frmMain.mnuMstfrm(0), Array(frmMastAccount.txttemp(2)), Array(GName)
        Case 7 '-- millcode
            GName = GProcGetColumnValue("TblMastGroup", "AgCode", 90029, "N", "AgName", "S")
            GProcShowForm frmMastAccount, frmMain.mnuMstfrm(0), Array(frmMastAccount.txttemp(2)), Array(GName)
        Case 8 '-- Broker
           GName = GProcGetColumnValue("TblMastGroup", "AgCode", 90016, "N", "AgName", "S")
           GProcShowForm frmMastAccount, frmMain.mnuMstfrm(0), Array(frmMastAccount.txttemp(2)), Array(GName)
        Case 12 '--TaxType
            'OldgFrmTypeStr = gfrmTypeStr
            GProcShowForm frmMastNarrationTax, frmMain.mnuMstfrm(6)
            'gfrmTypeStr = OldgFrmTypeStr
        Case 11 '--Narration transport
            OldFrmType = gfrmTypeStr
            gfrmTypeStr = "R"
            GProcShowForm frmMastNarration, frmMain.mnuMstfrm(8), gfrmTypeStr
            gfrmTypeStr = OldFrmType
    End Select
End If
Dim LeftPos As Long
Dim TopPos As Long
If KeyCode = vbKeyF2 Then '-- Select
    TopPos = txttemp(Index).Top + Me.Top + 650
    LeftPos = txttemp(Index).Left + Me.Left
    Select Case Index
        Case 11 '--Narration transport
            gClsSearch.SearchMultiField "tblMastNarration", "Narration", Array("Narration"), Array(txttemp(Index).Width), "NarrType='R'", "", "Narration", LeftPos, TopPos, False
            If gClsSearch.SearchMultiRetCol(0) <> "" Then txttemp(Index) = gClsSearch.SearchMultiRetCol(0)
    End Select
End If
End Sub
Private Sub txtTemp_KeyPress(Index As Integer, KeyAscii As Integer)
Dim LeftPos As Long
Dim TopPos As Long
Dim AgCode As Long
Dim BrkAcCode As Long
Dim wIsNoWork As Integer
TopPos = txttemp(Index).Top + Me.Top + 650
LeftPos = txttemp(Index).Left + Me.Left
If KeyAscii = 13 Then Exit Sub
Select Case Index
    Case 5  '-- Party   Alias Account
        gClsSearch.SearchMultiFieldAccMaster Chr(KeyAscii), LeftPos, TopPos
        KeyAscii = 0
        lblTinNo.Caption = ""
        lblMsOms.Caption = ""
        txtCode(5) = ""
        If gClsSearch.SearchMultiRetColAccMaster(0) <> "" Then
            txttemp(Index) = gClsSearch.SearchMultiRetColAccMaster(0)
            txtCode(5) = GProcGetColumnValue("TblMastAccount", "ACName", txttemp(Index), "S", "AcOurCode", "S")
            lblTinNo.Caption = GProcGetColumnValue("TblMastAccount", "ACName", txttemp(Index), "S", "AcGSTIN", "S")
            lblMsOms.Caption = GProcGetColumnValue("TblMastAccount", "ACName", txttemp(Index), "S", "AcMsOmsParty", "N")
            lblMblNoParty.Caption = GProcGetColumnValue("TblMastAccount", "ACName", txttemp(Index), "S", "AcContNo", "S")
            If Val(lblMsOms.Caption) = 0 Then
               lblMsOms.Caption = "M.S."
            Else
               lblMsOms.Caption = "O.M.S."
            End If
        End If
        If txttemp(5) <> "" Then
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
                X = "Select VFirmName,VBillNo,VBillDt,Amt1,Amt3,Amt2,amt1+amt3-Amt2 as Os,Amt5 from TmpGentbl where UserName = '" & gUserName & "' and amt5 > 60 and amt1+amt3-Amt2 > amt3  "
                i = GProcRstOpen(tmpRst, X, "R")
                If i > 0 Then
                   If MsgBox("Total O/s Bills Over 60 days are " + Str(i) + " Do You Want To See O/s List ? ", vbYesNo) = vbYes Then
                      gClsSearch.SearchMultiField X, "VFirmName,VBillNo,VBillDt,Amt1,Amt3,Amt2,os,Amt5", Array("Firm Name", "Bill No", "Date", "Inv.Amt", "Debit Amt", "Recd.Amt", "O/s Amt", "L.Days"), Array(2000, 2000, 1800, 1500, 1500, 1500, 1500, 1000), "", "", "VFirmName,VBillDt,VBillNo ", 0, 0, False, , True
                   End If
                End If
             End If
        End If
        '-- Broker from master
        If txttemp(8) = "" Then
            BrkAcCode = GProcGetColumnValue("TblMastAccount", "ACName", txttemp(Index), "S", "AcBrkCode", "N")
            txttemp(8) = GProcGetColumnValue("TblMastAccount", "ACCode", CStr(BrkAcCode), "N", "ACname", "S")
            txtCode(8) = GProcGetColumnValue("TblMastAccount", "ACCode", CStr(BrkAcCode), "N", "AcOurCode", "S")
        End If
    Case 6 '-- Sales A/c  Alias Account
        gClsSearch.SearchMultiField "tblMastaccount", "AcName,AcCity", Array("Sales Account", "City"), Array(3000, 2000), "(Agcode=90023)", Chr(KeyAscii), "", LeftPos, TopPos, True
        KeyAscii = 0
        If gClsSearch.SearchMultiRetCol(0) <> "" Then
            txttemp(Index) = gClsSearch.SearchMultiRetCol(0)
        End If
    Case 7  '-- Mill  Alias Account
         Dim MilTpCd As Long
         Dim ShowMill As String
         MilTpCd = GProcGetColumnValue("tblMastNarration", "Narration", "Mill Bill", "S", "Narrcode", "N")
         ShowMill = "AgCode=90029 and   tblMastaccount.Accode = tblMastBillSerial.MillCode and  tblMastBillSerial.SaleType = " & MilTpCd & "  and tblMastBillSerial.Vfirm = '" & gCCode & "'  "
         gClsSearch.SearchMultiField "tblMastaccount,tblMastBillSerial", "BillSr,AcName,AcOurCode", Array("Mill Ini.", "Mill Name", "Code"), Array(0, 8000, 1000), ShowMill, Trim(Chr(KeyAscii)), "AcName", LeftPos, TopPos, True
'        gClsSearch.SearchMultiField "qryMillHelpWitIni", "BillSr,MillName,AcOurCode", Array("Mill Ini.", "Mill Name", "Code"), Array(1000, 8000, 1000), "SaleType = '30' and vfirm = '" & gCCode & "' ", Trim(Chr(KeyAscii)), "MillName", LeftPos, TopPos, True
        KeyAscii = 0
        txtCode(7) = ""
        If gClsSearch.SearchMultiRetCol(1) <> "" Then
            txttemp(Index) = gClsSearch.SearchMultiRetCol(1)
            txtCode(7) = gClsSearch.SearchMultiRetCol(2)
        End If
        MillCtrlEd
        GBillChr
    Case 8 '-- Broker
        gClsSearch.SearchMultiField "tblMastaccount", "AcName,AcOurCode,AcCity", Array("Broker", "Code", "City"), Array(5000, 1000, 3000), "(AgCode=90016 )", Trim(Chr(KeyAscii)), "", LeftPos, TopPos, True
        KeyAscii = 0
        txtCode(8) = ""
        If gClsSearch.SearchMultiRetCol(0) <> "" Then
            txttemp(Index) = gClsSearch.SearchMultiRetCol(0)
            txtCode(8) = gClsSearch.SearchMultiRetCol(1)
        End If
End Select
If Index = 11 Then '-- Transport
    gClsSearch.SearchMultiField "tblMastNarration", "Narration", Array("Transport"), Array(txttemp(Index).Width), " NarrType='R'", Chr(KeyAscii), "Narration", 0, 0, True
    KeyAscii = 0
    If gClsSearch.SearchMultiRetCol(0) <> "" Then
        txttemp(Index) = gClsSearch.SearchMultiRetCol(0) '--Transport
    End If
End If
End Sub
Private Sub MillCtrlEd()
Dim ExemptMill As Long
'-- Is Exempt Mill
End Sub
Private Sub txtTemp_LostFocus(Index As Integer)
Dim AcCode As Long
txttemp(6) = txttemp(7).text
SaveBtnEd
MillCtrlEd
ShowAccountBalance
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
                Case 2 '-- Booking Date
                    txtGrid.Locked = False
                
                Case 3 '-- Booking Date
                     BookingData
                     If .Col = 3 Then txtGrid.Locked = True
                
                Case 4 '-- Count
                    If .TextMatrix(.Row, 1) > 0 Then   '-- Book No
                        txtGrid.Locked = True
                    Else
                        txtGrid.Locked = False
                    End If
                Case Else
                    If .Col = 5 Then
                       OldBag = .TextMatrix(.Row, 5)
                    End If
                    txtGrid.Locked = False
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
        If ((Val(.TextMatrix(.Row, 5)) = 0) Or (Val(.TextMatrix(.Row, 6)) = 0)) And .rows > 2 Then
            .RemoveItem (.Row)
            txtGrid.text = ""
        End If
        txtGrid.Visible = False
        txt(15).SetFocus
        Exit Sub
    End If
End With
'If FormAction <> vbDataActionAddNew And FormAction <> vbDataActionUpdate Then Exit Sub
'---- Navigate Row ,Col
If KeyAscii = 13 Then
    KeyAscii = 0
    With msGrid
        If .Col = 6 Then 'rstblSub.Fields.Count - 2 Then
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
If msGrid.Col < rstblSub.Fields.Count - 1 Then
   txtGrid.Move msGrid.CellLeft + msGrid.Left, msGrid.CellTop + msGrid.Top, msGrid.CellWidth, msGrid.CellHeight
End If
txtGrid.Visible = True
txtGrid.Enabled = True
DoEvents
If KeyAscii <> 0 Then
    Select Case rstblSub.Fields(msGrid.Col).Type
        Case adChar, adVarWChar
            If msGrid.Col = 1 Or msGrid.Col = 4 Or msGrid.Col = 5 Then '-- Count ,godown
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
    txtGrid.text = ""
    txtGrid.Visible = False
End If
DoEvents
'If msGrid.Col >= 13 Then
    CalGridAmt
    CalAmount
'End If
End Sub
Private Sub CalGridAmt()
Dim MillCode As Long
Dim StdPack As Double
Dim NetRt As Double
Dim GAmt As Double
Dim GSubAmt As Double
Dim wCharityRt As Double
Dim wCharityOn As Integer
Dim wCharityAmt As Double
Dim CalWt As Double
CheckForNumberGrid
If msGrid.TextMatrix(msGrid.Row, 2) = 0 Then msGrid.TextMatrix(msGrid.Row, 2) = gCYear '-- Book year

If msGrid.TextMatrix(msGrid.Row, 1) = 0 Then msGrid.TextMatrix(msGrid.Row, 3) = "" '-- Book date
If msGrid.TextMatrix(msGrid.Row, 1) = 0 Then msGrid.TextMatrix(msGrid.Row, 2) = 0 '-- Book year
MillCode = GProcGetColumnValue("TblMastAccount", "AcName", txttemp(7), "S", "Accode", "N")
StdPack = GProcGetColumnValue("TblMastItem", "ITName", msGrid.TextMatrix(msGrid.Row, 4), "S", "ItStdpack", "N", " ItMillCode=" & MillCode)

If msGrid.Row = 0 Then Exit Sub
'-- Wt=Bag * Std Pack
CalWt = Format(Abs(((msGrid.TextMatrix(msGrid.Row, 5) * StdPack) - msGrid.TextMatrix(msGrid.Row, 6))), WtStr)
'If StdPack > 0 And ((msGrid.TextMatrix(msGrid.Row, 5) = 0) Or (CalWt >= StdPack)) Then
If OldBag <> msGrid.TextMatrix(msGrid.Row, 5) And StdPack > 0 And FormAction = vbDataActionAddNew Then
'If FormAction = vbDataActionAddNew And msGrid.TextMatrix(msGrid.Row, 6) = 0 Then
    msGrid.TextMatrix(msGrid.Row, 6) = msGrid.TextMatrix(msGrid.Row, 5) * StdPack
    OldBag = msGrid.TextMatrix(msGrid.Row, 5)
End If

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
                    msGrid.TextMatrix(RowIndex, i) = ""
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
    
    If msGrid.Col = 2 Then txtGrid.SelStart = 3
    
    
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
If KeyCode = vbKeyF3 Then
Select Case msGrid.Col
    Case 4 '-- Item
        GProcShowForm frmMastItem, frmMain.mnuMstfrm(3), Array(frmMastItem.txttemp(9)), Array(txttemp(7))
        Exit Sub
End Select
End If
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
Dim Tp As String
Dim MillCode As Long
Dim BalBag As Integer
Dim BalWt As Double
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
            Case 4 '-- Count
                If msGrid.TextMatrix(msGrid.Row, 1) > 0 Then
                    KeyAscii = 0
                Else
                    MillCode = GProcGetColumnValue("TblMastAccount", "AcName", txttemp(7), "S", "Accode", "N")
                    gClsSearch.SearchMultiField "tblMastItem,tblMastaccount", "ItName,AcName,ItStdpack,ItStdrateper,ItCode,InDailyReport", Array("Count Name", "Mill", "Std. Pack", "Rate Per", "ItCode", ""), Array(5000, 5000, 1000, 1000, 0, 0), "tblMastItem.ITMillCode=tblMastaccount.acCode and ITMillCode=" & MillCode, Chr(KeyAscii), "ItName", LeftPos, TopPos, True
                    If gClsSearch.SearchMultiRetCol(5) = 1 Then
                       MsgBox "This is No Use Count "
                    End If
                    txtGrid = ""
                    KeyAscii = 0
                    If gClsSearch.SearchMultiRetCol(0) <> "" And gClsSearch.SearchMultiRetCol(5) = 0 Then '-- Item Wt ,Rate From Master
                        txtGrid = gClsSearch.SearchMultiRetCol(0)
                        txtGrid_KeyPress 13
                    End If
                  End If
            Case Else
                KeyAscii = GProcValidateKey(rstblSub, msGrid.Col, KeyAscii, txtGrid)
        End Select
   End If
    If KeyAscii = 13 Then
        CheckForNumberTxtGrid
    End If
    If KeyAscii = 13 Then
        CheckForNumberTxtGrid
        If txtGrid = "0" And .Col = 1 Then
           KeyAscii = 27
        End If
    End If
    
    '--- Navigate Row ,Col
    If KeyAscii = 13 And .Col < 6 Then 'rstblSub.Fields.Count - 2 Then
        .Col = .Col + 1
    ElseIf KeyAscii = 13 And .Col = 6 Then '= rstblSub.Fields.Count - 2 Then
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
        If ((Val(.TextMatrix(.Row, 5)) = 0) Or (Val(.TextMatrix(.Row, 6)) = 0)) And .rows > 2 Then
            .RemoveItem (.Row)
            txtGrid.text = ""
        End If
        txtGrid.Visible = False
        If txt(15).Visible = True Then
            txt(15).SetFocus
        Else
            txt(16).SetFocus
        End If
        End If
        Exit Sub
'    End If
End With
End Sub
'-- Grid Required
Private Function ValidateGrid() As Boolean
Dim i As Long
Dim MillCode As Long
Dim RowIndex As Long
Dim ItCode As Long
'Dim X As String
Dim Rs1 As Recordset
Dim oldFrmTypeStr As String
Dim LRNo As String
Dim pItCtrlNo As Long
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
        For i = 0 To rstblSub.Fields.Count - 1
            Select Case i
                Case 4 '-- Count
                 '-Rate=0 ,count,wt=0 then delete row
                    If RowIndex > 1 And Trim(.TextMatrix(RowIndex, 4)) = "" And CDbl(.TextMatrix(RowIndex, 5)) = 0 And CDbl(.TextMatrix(RowIndex, 6)) = 0 Then
                        .RemoveItem (RowIndex)
                        GoTo Validate
                    End If
                    If Trim(.TextMatrix(RowIndex, i)) = "" Then
                        MsgBox "Check Count.", vbCritical + vbOKOnly, Me.Caption
                        .Col = i
                        msGrid.SetFocus
                        Exit Function
                    End If
                    MillCode = GProcGetColumnValue("TblMastAccount", "AcName", txttemp(7), "S", "Accode", "N")
                    If GProcGetColumnValue("TblMastItem", "ItName", .TextMatrix(RowIndex, i), "S", "Itcode", "N", "ITMillCode=" & MillCode) = 0 Then
                        MsgBox "Check Count.Mill and Count are Mismatch.", vbCritical + vbOKOnly, Me.Caption
                        .Col = i
                        msGrid.SetFocus
                        Exit Function
                    End If
                    If GProcGetColumnValue("TblMastItem", "ItName", .TextMatrix(RowIndex, i), "S", "InDailyReport", "N", "ITMillCode=" & MillCode) = 1 Then
                        MsgBox "Check Count is Marked as No Use", vbCritical + vbOKOnly, Me.Caption
                        .Col = i
                        msGrid.SetFocus
                        Exit Function
                    End If
                
                
                
                Case 5 '-- Bag
                    If CDbl(.TextMatrix(RowIndex, i)) = 0 Then
                        MsgBox "Check Bag.", vbCritical + vbOKOnly, Me.Caption
                        .Col = i
                        msGrid.SetFocus
                        Exit Function
                    End If

'                    '-- If Bag < Issue Bag
'                    If CDbl(.TextMatrix(RowIndex, 7)) < CDbl(.TextMatrix(RowIndex, 13)) Then
'                        MsgBox "Check Bag.Bag Less than issued.", vbCritical + vbOKOnly, Me.Caption
'                        .Col = i
'                        msGrid.SetFocus
'                        Exit Function
'                    End If
                Case 6 '-- Wt
                    If CDbl(.TextMatrix(RowIndex, i)) = 0 Then
                        MsgBox "Check Weight.", vbCritical + vbOKOnly, Me.Caption
                        .Col = i
                        msGrid.SetFocus
                        Exit Function
                    End If
'                    '-- If Wt < Issue wt
'                    If CDbl(.TextMatrix(RowIndex, 8)) < CDbl(.TextMatrix(RowIndex, 14)) Then
'                        MsgBox "Check Wt.Weight Less than issued.", vbCritical + vbOKOnly, Me.Caption
'                        .Col = i
'                        msGrid.SetFocus
'                        Exit Function
'                    End If
           End Select
        Next i
        '--- Duplicate LR No ,It Ctrl  No (Booking)
'        LRNo = .TextMatrix(RowIndex, 6)
'        pItCtrlNo = .TextMatrix(RowIndex, 19)
'        For i = 1 To .Rows - 1
'            '--- Duplicate LR No
'            If i <> RowIndex And .TextMatrix(i, 6) = LRNo Then
'                MsgBox "Duplicate LR No.", vbCritical + vbOKOnly, Me.Caption
'                .Col = 6
'                msGrid.SetFocus
'                Exit Function
'            End If
'            '-- Duplicate Booking Count Selection
' '           If i <> RowIndex And .TextMatrix(i, 19) = pItCtrlNo And pItCtrlNo > 0 Then
'  '              MsgBox "Duplicate Booking Count Selection.", vbCritical + vbOKOnly, Me.Caption
'   '             .Col = 1
'    '            msGrid.SetFocus
'     '           Exit Function
'      '      End If
'        Next i
        '-- Invalid Booking
        If Val(msGrid.TextMatrix(msGrid.Row, 1)) <> 0 Then
           MillCode = GProcGetColumnValue("TblMastAccount", "AcName", txttemp(7), "S", "Accode", "N")
           Dim BalBag As Integer
           Dim BalWt As Double
'           If Val(msGrid.TextMatrix(msGrid.Row, 1)) < 0 Then
 '             X = "Select BkParty,BkBroker,BkItCode,BkBag-BkCbag as BkBg,BkWt-BkCWt as BkKg ,BkRt,BkRtPer,Vdt,Vtype,VYear from tblBooking where VNo=" & msGrid.TextMatrix(msGrid.Row, 1) & " and vtype = 'OM' and VFirm  = '" & gCCode & "' and VYear < " & gCYear & " and BkMillCode = " & MillCode & " "
  '         Else
              X = "Select BkParty,BkBroker,BkItCode,BkBag-BkCbag as BkBg,BkWt-BkCWt as BkKg ,BkRt,BkRtPer,Vdt,Vtype,VYear from tblBooking where VNo=" & msGrid.TextMatrix(msGrid.Row, 1) & " and vtype = 'OM' and VFirm  = '" & gCCode & "' and VYear = " & Val(msGrid.TextMatrix(msGrid.Row, 2)) & " and BkMillCode = " & MillCode & " "
   '        End If
           
           If GProcRstOpen(Rs1, X, "R") > 0 Then
'              BalBag = Rs1.Fields("BkBg")
'              BalWt = Rs1.Fields("BkKg")
'              BalBag = BalBag - GProcGetColumnValue("tblSaleSub", "SlSubItCode", Rs1.Fields("BkItCode"), "N", "sum(SlSubBag)", "N", "VType = '" & frmTypeStr & "' and  SlSubBookNo = " & msGrid.TextMatrix(RowIndex, 1) & "  and SlSubBkType = 'OM' and SlSubBkVyear =  " & Rs1.Fields("Vyear") & "   and VFirm = '" & gCCode & "'  and  not (vno = " & txt(0) & " and VYear=" & gCYear & " and VFirm = '" & gCCode & "'  and vtype = '" & frmTypeStr & "')")
'              BalWt = BalWt - GProcGetColumnValue("tblSaleSub", "SlSubItCode", Rs1.Fields("BkItCode"), "N", "sum(SlSubWt)", "N", "VType = '" & frmTypeStr & "' and  SlSubBookNo = " & msGrid.TextMatrix(RowIndex, 1) & "  and SlSubBkType = 'OM' and SlSubBkVyear =  " & Rs1.Fields("Vyear") & "   and VFirm = '" & gCCode & "'  and  not (vno = " & txt(0) & " and VYear=" & gCYear & " and VFirm = '" & gCCode & "'  and vtype = '" & frmTypeStr & "')")
'              If Val(msGrid.TextMatrix(msGrid.Row, 4)) > BalBag Then
'                 MsgBox "Bags Greater Than Balance Bags.", vbInformation + vbOKOnly, Me.Caption
'                 .Col = 4
'                 msGrid.SetFocus
'                 Exit Function
'             End If
'             If Val(msGrid.TextMatrix(msGrid.Row, 5)) > BalWt Then
'                 MsgBox "Weight Greater Than Balance Weight", vbInformation + vbOKOnly, Me.Caption
'                 .Col = 5
'                 msGrid.SetFocus
'                 Exit Function
'             End If
'
           End If
           If Rs1.EOF And Rs1.BOF Then
              MsgBox "Check Booking No.", vbInformation + vbOKOnly, Me.Caption
              .Col = 1
              msGrid.SetFocus
              Exit Function
           End If
           Rs1.Close
        End If







'        If CLng(.TextMatrix(RowIndex, 1)) > 0 Then
'            MillCode = GProcGetColumnValue("TblMastAccount", "AcName", txttemp(5), "S", "Accode", "N")
'            ItCode = GProcGetColumnValue("tblMastItem", "ItName", .TextMatrix(RowIndex, 3), "S", "ItCode", "N", " ItMillCode=" & MillCode)
'            X = "Select * from tblvouSub where Vno=" & CLng(.TextMatrix(RowIndex, 1)) & " and  " _
'            & " VSUBITCODE=" & ItCode & " and  vsubitctrlno=" & .TextMatrix(RowIndex, 19) & " and Vtype='OT' and "
'            If gBackEndDB = gBackEndAccess Then
'                X = X + "  vdt=cdate('" & CDate(.TextMatrix(RowIndex, 2)) & "')"
''                X = X + "  vdt=cdate('" & Format(CDate(.TextMatrix(RowIndex, 2)), "dd/mm/yyyy") & "')"
'            Else '-- Oracle
'                X = X + "  vdt= to_date('" & CDate(.TextMatrix(RowIndex, 2)) & "','dd/MM/yyyy')"
'            End If
'            GProcRstOpen Rs1, X, "R"
'            If Rs1.EOF And Rs1.BOF Then
'                MsgBox "Check Booking No.", vbInformation + vbOKOnly, Me.Caption
'                .Col = 1
'                msGrid.SetFocus
'                Exit Function
'            End If
'            Rs1.Close
'        End If
      Next RowIndex
End With
ValidateGrid = True
ReNumberGridSrno
End Function
Private Sub FillCombo()
Call GProcFillCombo(Cbo(22), "select Narration from tblMastNarration where NarrType ='T' order by Narration")
Cbo(22) = GProcGetColumnValue("tblMastNarration", "NarrCode", CStr(gTaxCodeSM), "N", "Narration", "S")
End Sub
Private Sub SaveBtnEd()
'--- Vno,vdt,Party,A/c,Mill,Amount
If (CDbl(txt(0)) = 0 And FormAction <> vbDataActionAddNew) Or GProcIsDateValid(dtpDt(2)) = False Or txttemp(5) = "" Or txttemp(6) = "" Or txttemp(7) = "" Or CDbl(txt(16)) = 0 Then
    cmdBtn(13).Enabled = False
    cmdBtn(13).BackColor = &H80000004
Else
    cmdBtn(13).Enabled = True
    cmdBtn(13).BackColor = &HFFFF80
End If
End Sub
Private Function ValidateData() As Boolean
Dim i As Long
'Dim X As String
Dim Y As String
Dim GName As String
Dim MillCode As Long
Dim rstAddCheck As Recordset
Dim rstModiCheck As Recordset
Dim ExpAcCode As Long
MillCtrlEd
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
    
    '---Validate Vou.Date = Bill Date
'    If dtpDt(10) <> dtpDt(2) Then
'        MsgBox "Vou.Date Muse Be Same As Bill Date", vbInformation + vbOKOnly, Me.Caption
'        mskDt(12).SetFocus
'        Exit Function
'    End If
    '--- Validate Grid
    If ValidateGrid = False Then
        'MSGrid.SetFocus
        Exit Function
    End If
    '--- Vno
    If CLng(txt(0)) = 0 And FormAction <> vbDataActionAddNew Then
        MsgBox "Check Voucher No.", vbInformation + vbOKOnly, Me.Caption
        dtpDt(2).SetFocus
        Exit Function
    End If
    '--- Vdt
    If GProcIsDateValid(dtpDt(2)) = False Then
        MsgBox "Check Voucher Date.", vbInformation + vbOKOnly, Me.Caption
        dtpDt(2).SetFocus
        Exit Function
    End If
    '---Tax Type
    If Cbo(22) = "" Then
        MsgBox "Check Tax Type.", vbInformation + vbOKOnly, Me.Caption
        Cbo(22).SetFocus
        Exit Function
    End If
    '---- Party Account
    If txttemp(5) = "" Then
        MsgBox "Check Party.", vbInformation + vbOKOnly, Me.Caption
        txttemp(3).SetFocus
        Exit Function
    End If
        '--Create Party
        If GProcGetColumnValue("tblMastAccount", "AcName", txttemp(5), "S", "AcCode", "N") = 0 Then
           GName = GProcGetColumnValue("TblMastGroup", "AgCode", 90017, "N", "AgName", "S")
           GProcShowForm frmMastAccount, frmMain.mnuMstfrm(0), Array(frmMastAccount.txt(1), frmMastAccount.txttemp(2)), Array(txttemp(5), GName)
           Exit Function
        End If
    '---- Sales Account
    If txttemp(6) = "" Then
        MsgBox "Check Purchase Account.", vbInformation + vbOKOnly, Me.Caption
        txttemp(6).SetFocus
        Exit Function
    End If
        '--Create Party
        If GProcGetColumnValue("tblMastAccount", "AcName", txttemp(6), "S", "AcCode", "N") = 0 Then
           GName = GProcGetColumnValue("TblMastGroup", "AgCode", 90023, "N", "AgName", "S")
           GProcShowForm frmMastAccount, frmMain.mnuMstfrm(0), Array(frmMastAccount.txt(1), frmMastAccount.txttemp(2)), Array(txttemp(6), GName)
           Exit Function
        End If
    
    
    '---- Mill
    If txttemp(7) = "" Then
        MsgBox "Check Mill.", vbInformation + vbOKOnly, Me.Caption
        txttemp(7).SetFocus
        Exit Function
    End If
        '--Create Mill
        If GProcGetColumnValue("tblMastAccount", "AcName", txttemp(7), "S", "AcCode", "N") = 0 Then
           GName = GProcGetColumnValue("TblMastGroup", "AgCode", 90029, "N", "AgName", "S")
           GProcShowForm frmMastAccount, frmMain.mnuMstfrm(0), Array(frmMastAccount.txt(1), frmMastAccount.txttemp(2)), Array(txttemp(7), GName)
           Exit Function
        End If
    '--- Broker
        '--Create Broker
        If txttemp(8) = "" Then txttemp(8) = "-"
        If GProcGetColumnValue("tblMastAccount", "AcName", txttemp(8), "S", "AcCode", "N") = 0 Then
           GName = GProcGetColumnValue("TblMastGroup", "AgCode", 90016, "N", "AgName", "S")
           GProcShowForm frmMastAccount, frmMain.mnuMstfrm(0), Array(frmMastAccount.txt(1), frmMastAccount.txttemp(2)), Array(txttemp(8), GName)
           Exit Function
        End If
    '---- Amount
    If CDbl(txt(16)) <= 0 Then
        MsgBox "Check Bill Amount.", vbInformation + vbOKOnly, Me.Caption
        Exit Function
    End If
    i = GProcGetColumnValue("tblMastAccount", "AcName", txttemp(7), "S", "AcComOn", "N")
    If CDbl(txt(15)) <= 0 And i <> 0 Then
        MsgBox "Check Ex-Mill Amount.", vbInformation + vbOKOnly, Me.Caption
        Exit Function
    End If
    If txt(9) = "" Then
        MsgBox "Check Bill No.", vbInformation + vbOKOnly, Me.Caption
        txt(9).SetFocus
        Exit Function
    End If
    MillCode = GProcGetColumnValue("TblMastAccount", "AcName", txttemp(7), "S", "Accode", "N")
    X = "Select SlBillNo From " & tblName & " Where vno<> " & CLng(txt(0)) & " and SlAcCrCode = " & MillCode & " and SlBillNo = '" & txt(9) & "'  and Vtype='" & frmTypeStr & "'  and Vfirm = '" & gCCode & "' and  VYear=" & gCYear
        i = GProcRstOpen(rstAddCheck, X, "R")
        If i > 0 Then
            MsgBox "Record is Available of Bill No. " & txt(9), vbCritical + vbOKOnly, "Duplicate Record"
            txt(9).SetFocus
            Exit Function
        End If
        rstAddCheck.Close
     
     
     '-- Expenses A/c
'    If gExpensesAcCode = GProcGetColumnValue("tblmastAccount", "AcName", "-", "S", "AcCode", "N") Then
'        ExpAcCode = 0
'    Else
'        ExpAcCode = gExpensesAcCode
'    End If
'    If txtExpensesAmt > 0 And ExpAcCode = 0 Then
'        MsgBox "Check Purchase Expenses Account.", vbInformation + vbOKOnly, Me.Caption
'        GProcShowForm frmSetting, frmMain.mnuMstSetting(0)
'        Exit Function
'    End If
    '----Duplicate Records
    '----Addition Case
    If FormAction = vbDataActionAddNew Then
        X = "Select vno From " & tblName & " Where vno=" & CLng(txt(0)) & " and Vtype='" & frmTypeStr & "'  and Vfirm = '" & gCCode & "'  and VYear=" & gCYear
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
If IsDate(mskDt(32)) = False Then mskDt(32) = "__/__/____" '--Due Dt
If IsDate(dtpDt(2)) Then '--Vdt
    txt(2) = CDate(dtpDt(2))
Else
    txt(2) = ""
End If
If IsDate(mskDt(32)) Then '--Due Date
    txt(32) = CDate(mskDt(32))
Else
    txt(32) = ""
End If
If IsDate(dtpDt(10)) Then '--Due Date
    txt(10) = CDate(dtpDt(10))
Else
    txt(10) = ""
End If
If IsDate(dtpDt(33)) Then '--Due Date
    txt(33) = CDate(dtpDt(33))
Else
    txt(33) = ""
End If


'-- Dr Code
txt(5) = GProcGetColumnValue("tblMastAccount", "AcName", txttemp(5), "S", "AcCode", "N")
'-- Cr Code
txt(6) = GProcGetColumnValue("tblMastAccount", "AcName", txttemp(6), "S", "AcCode", "N")
'--Mill
txt(7) = GProcGetColumnValue("tblMastAccount", "AcName", txttemp(7), "S", "AcCode", "N")

'--TdsAc
If txttemp(70) = "" Then
   txttemp(70) = "-"
End If
txt(70) = GProcGetColumnValue("tblMastAccount", "AcName", txttemp(70), "S", "AcCode", "N")

'-- broker
txt(8) = GProcGetColumnValue("tblMastAccount", "AcName", txttemp(8), "S", "AcCode", "N")
'-- Transport
If txttemp(11) = "" Then
   txttemp(11) = "-"
End If
txt(11) = GProcGetColumnValue("tblMastNarration", "Narration", txttemp(11), "S", "NarrCode", "N")
txt(9) = lblSr + "/" + txt(9)
txt(36) = chkIsPaymenttoMill.Value
'-- TaxType
txt(22) = GProcGetColumnValue("tblMastNarration", "Narration", Cbo(22), "S", "NarrCode", "N")
txt(3) = gCYear
txt(4) = gCCode
End Sub
Private Sub FillTempFromTxt()
'Dim X As String
Dim MsOms As Integer
FillCombo
If IsDate(txt(2)) Then '--VDt
    dtpDt(2) = txt(2)
End If
If IsDate(txt(32)) Then '--Due Date
    mskDt(32) = txt(32)
Else
    mskDt(32) = "__/__/____"
End If
If IsDate(txt(10)) Then '--Due Date
    dtpDt(10) = txt(10)
End If
If IsDate(txt(33)) Then '--Entdt
    dtpDt(33) = txt(33)
End If



'-- Dr Code
txttemp(5) = GProcGetColumnValue("tblMastAccount", "AcCode", txt(5), "N", "AcName", "S")
txtCode(5) = GProcGetColumnValue("tblMastAccount", "AcCode", txt(5), "N", "AcOurCode", "S")
lblTinNo.Caption = GProcGetColumnValue("tblMastAccount", "AcCode", txt(5), "N", "AcGSTIN", "S")
MsOms = GProcGetColumnValue("tblMastAccount", "AcCode", txt(5), "N", "AcMsOmsParty", "N")
If MsOms = 0 Then
   lblMsOms.Caption = "M.S."
Else
   lblMsOms.Caption = "O.M.S."
End If
lblMblNoParty.Caption = GProcGetColumnValue("TblMastAccount", "AcCode", txt(5), "N", "AcContNo", "S")

'-- Cr Code
txttemp(6) = GProcGetColumnValue("tblMastAccount", "AcCode", txt(6), "N", "AcName", "S")
'--Mill
txttemp(7) = GProcGetColumnValue("tblMastAccount", "AcCode", txt(7), "N", "AcName", "S")
txtCode(7) = GProcGetColumnValue("tblMastAccount", "AcCode", txt(7), "N", "AcOurCode", "S")
'--TdsAc
txttemp(70) = GProcGetColumnValue("tblMastAccount", "AcCode", txt(70), "N", "AcName", "S")
txtCode(70) = GProcGetColumnValue("tblMastAccount", "AcCode", txt(70), "N", "AcOurCode", "S")

GBillChr
'--Broker
txttemp(8) = GProcGetColumnValue("tblMastAccount", "AcCode", txt(8), "N", "AcName", "S")
txtCode(8) = GProcGetColumnValue("tblMastAccount", "AcCode", txt(8), "N", "AcOurCode", "S")

'--Transport
txttemp(11) = GProcGetColumnValue("tblMastNarration", "NarrCode", txt(11), "N", "Narration", "S")

'-- TaxType
Cbo(22) = GProcGetColumnValue("tblMastNarration", "NarrCode", txt(22), "N", "Narration", "S")
txt(9) = Mid(txt(9), 5, 15)
chkIsPaymenttoMill.Value = Val(txt(36))
'-- Fill Related Record
FillRelatedRecord
ShowAccountBalance
MillCtrlEd
SetControlEd
End Sub
Private Sub FillRelatedRecord()
Dim RowIndex As Long
Dim ColIndex As Long
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
            Next ColIndex
            RowIndex = RowIndex + 1
            .MoveNext
            If Not .EOF Then msGrid.rows = msGrid.rows + 1
        Loop
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
Dim SubAmt As Double
Dim Amt As Double
Dim RowIndex As Integer
Dim MillAcCode As Long
Dim wMillCode As String
Dim ItemType As Integer
Dim IsTdsParty As Integer
Dim TdsFrom1stBill As Integer
Dim wPtyCd As Integer
Dim TDSLimitAmt As Double
Dim TdsAmt As Double
Dim AssValue As Double
Dim TdsOnAmt As Double
wPtyCd = GProcGetColumnValue("TblMastAccount", "AcName", txttemp(5), "S", "AcCode", "N")
wMillCode = GProcGetColumnValue("TblMastAccount", "AcName", txttemp(7), "S", "AcCode", "N")

RowIndex = 1
With msGrid
    Do While RowIndex < .rows
        '--Tot Wt
        If Val(.TextMatrix(RowIndex, 6)) > 0 Then
            TotWt = TotWt + Val(.TextMatrix(RowIndex, 6))
        End If
        RowIndex = RowIndex + 1
    Loop
End With
'-- Due Date
If IsNumeric(txt(31)) = False Then txt(31) = "0"
If IsNumeric(txt(45)) = False Then txt(45) = "0"
If IsNumeric(txt(10)) = False Then txt(10) = "0"

If Val(txt(21)) = 0 Then
   txt(21) = Format(txt(15), FStr)
End If
txt(31) = Format(txt(31), PStr)
txt(45) = Format(txt(45), PStr)
If IsDate(dtpDt(2)) Then mskDt(32) = DateAdd("d", (Val(txt(31)) + Val(txt(45))), dtpDt(2))
If IsDate(dtpDt(2)) Then dtpDt(10) = DateAdd("d", (Val(txt(45))), dtpDt(2))

'---- Sub amt
'--For TDS
If Format(dtpDt(2).Value, "YYYY/MM/dd") >= Format("2021/07/01", "YYYY/MM/dd") And chkIsPaymenttoMill.Value = 0 Then
    gCn.BeginTrans
    IsTdsParty = GProcGetColumnValue("TblMastAccount", "AcName", txttemp(5), "S", "AcIsLessTDSOnRec", "N")
    TdsFrom1stBill = GProcGetColumnValue("TblMastAccount", "AcName", txttemp(5), "S", "AcIsTDSfrom1stBill", "N")
    gCn.CommitTrans
Else
   IsTdsParty = 0
End If

Amt = Format(CDbl(txt(16)), FStr)
TDSLimitAmt = 0

AssValue = Val(txt(21))
'If gCYear >= 2023 Then
TdsAmt = 0
txt(68) = Format(0, FStr)
txt(69) = Format(0, FStr)
txttemp(70) = "-"
txtCode(70) = ""
If Format(dtpDt(2).Value, "YYYY/MM/dd") >= Format("2021/07/01", "YYYY/MM/dd") And TdsFrom1stBill = 1 Then
       TdsOnAmt = AssValue
       txt(74) = Format(TdsOnAmt, FStr)
       txt(68) = 0.1
       TdsAmt = (TdsOnAmt * Val(txt(68))) / 100
       TdsAmt = Format(GProcMakeRounding(TdsAmt), FStr)
       txt(69) = TdsAmt
       If TdsAmt > 0 Then
          txttemp(70) = txttemp(7)
          txtCode(70) = txtCode(7)
       Else
          txttemp(70) = "-"
          txtCode(70) = ""
       End If
End If
If Format(dtpDt(2).Value, "YYYY/MM/dd") >= Format("2021/07/01", "YYYY/MM/dd") And IsTdsParty = 1 And TdsFrom1stBill = 0 Then
       X = "Select sum(SlBillAmt) from tblSale where SlAcDrCode = " & wPtyCd & " and SlAcCrCode = " & wMillCode & " and vfirm = '" & gCCode & "' and vtype = 'SM' and vyear = " & gCYear & " and vdt <= '" & Format(dtpDt(2).Value, "YYYY/MM/dd") & "'  "
       X = X + " and  ltrim(str(vno))+vtype+vFirm+ltrim(str(vyear)) <> '" & Trim(Str(txt(0))) & "" & gfrmTypeStr & "" & gCCode & "" & LTrim(Str(gCYear)) & "'"
       i = GProcRstOpen(tmpRst, X, "R")
       TDSLimitAmt = IIf(IsNull(tmpRst.Fields(0)), 0, tmpRst.Fields(0))
       TDSLimitAmt = TDSLimitAmt + Amt
    
       If (TDSLimitAmt >= 5000000) Then
          txt(68) = 0.1
          TdsOnAmt = 0
          TdsAmt = (Val(txt(21)) * Val(txt(68))) / 100
          TdsAmt = Format(GProcMakeRounding(TdsAmt), FStr)
          txt(69) = TdsAmt
          If TdsAmt > 0 Then
             txttemp(70) = txttemp(7)
             txtCode(70) = txtCode(7)
          Else
             txttemp(70) = "-"
             txtCode(70) = ""
          End If
       End If
       lblTcsLimit.Caption = "Sales up to Date = " + Str(TDSLimitAmt)
       tmpRst.Close
End If
CheckForNumberGrid
CheckForNumber
End Sub
'--- Is Related Record
Private Function IsRelatedRecord() As Boolean
Dim RsG As Recordset
'Dim X As String
X = "Select Vno,Vdt from tblVoucher where VBillType= '" & frmTypeStr & "' and " _
& "  VBillVno = " & txt(0) & " and VFirm  = '" & gCCode & "' and  VBillVYear= " & gCYear & ""
GProcRstOpen RsG, X, "R"
If RsG.EOF Then
Else
    IsRelatedRecord = True
    MsgBox "You can not Modify/Delete this record.Related Record No. " & RsG.Fields(0) & " is exists.", vbCritical + vbOKOnly, App.Title
    Exit Function
End If
RsG.Close


'-- Audited
'If GProcISAudited(txt(0), txt(1), txt(23)) = True Then
'    MsgBox "Record is Audited.You can not Modify/Delete this Record.", vbCritical + vbOKOnly, App.Title
'    IsRelatedRecord = True
'    Exit Function
'End If
''-- Is Gate Pass
'X = "Select Vno,Vdt from tblVouSub where Vtype='GP' and " _
'& " VSubPItCtrlNo in (select VSubItCtrlNo from tblVouSub where vtype='" & frmTypeStr & "' and vno=" & txt(0) & " and " _
'& " VYear=" & gCYear & ")"
'GProcRstOpen RsG, X, "R"
'If RsG.EOF Then
'Else
'    IsRelatedRecord = True
'    MsgBox "You can not Modify/Delete this record.Related Gate Pass No. " & RsG.Fields(0) & " is exists.", vbCritical + vbOKOnly, App.Title
'    Exit Function
'End If
'RsG.Close
End Function
'-- Is Related GatePass
Private Function IsRelatedGatePass() As Boolean
Dim RsG As Recordset
'Dim X As String
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
'Dim X As String
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
'---------- Booking Vs Purchase
Private Sub PrepareBooking(Optional Chr As String)
'''Dim MillCode As Long
'''Dim QStr As String
'''Dim RsBk As Recordset
'''Dim Chr2 As String
'''If IsDate(dtpDt(2)) = True Then
'''    MillCode = GProcGetColumnValue("TblMastAccount", "AcName", txttemp(5), "S", "Accode", "N")
'''    QStr = " select tblAddless.Adgpno, tblVousub.VNo,format(tblVousub.vdt,'dd/mm/yyyy'),ItName,vsubbag-(vsubissbag),vsubwt-(vsubisswt),vsubbag,vsubwt,vsubitctrlno,VsubRt from " _
'''    & " tblVousub,tblmastItem,tblAddless Where " _
'''    & " tblVousub.vno = tblAddless.vno and  tblVousub.vtype  = tblAddless.vtype and  tblVousub.vyear = tblAddless.vyear and " _
'''    & " Itcode=VSUBITCODE and ItMillCode=" & MillCode & " and " _
'''    & " ((vsubbag-(vsubissbag))> 0) and tblVousub.vtype ='OT'  and VSUBBOOKISCOMPLETED=0 and "
'''    If gBackEndDB = gBackEndAccess Then
'''        QStr = QStr + " tblVousub.vdt<=cdate('" & CDate(dtpDt(2)) & "')"
'''    Else '-- Oracle
'''        QStr = QStr + " tblVousub.vdt <= to_date('" & CDate(dtpDt(2)) & "','dd/MM/yyyy')"
'''    End If
'''
'''
'''
'''    With gClsSearch
'''        .SearchMultiField QStr, "", Array("Contract No.", " Booking No.", "Date", "Count", "Bal Bag", "Bal Wt.", "Bag", "Wt", "itctrlno", "Rate"), Array(1200, 1200, 1500, 1500, 1000, 1000, 500, 1000, 0, 1000), "", Chr, "tblVousub.Vdt,tblAddless.Adgpno,tblVousub.vno,VSUBCTRNO", 0, 0, False, , True
'''        If .SearchMultiRetCol(1) <> "" Then
'''            QStr = "select tblvousub.*,tblmastItem.ItName from tblvousub,tblMastItem where " _
'''            & " ItCode=VSUBITCODE and VSUBITCTRLNO= " & CLng(.SearchMultiRetCol(8))
'''              If GProcRstOpen(RsBk, QStr, "R") > 0 Then
'''                    txtGrid = ""
'''                    txtGrid = CLng(.SearchMultiRetCol(1)) '--Book Vno
'''                    msGrid.TextMatrix(msGrid.Row, 1) = CLng(.SearchMultiRetCol(1)) '--Book Vno
'''                    msGrid.TextMatrix(msGrid.Row, 2) = Format(.SearchMultiRetCol(2), "dd/mm/yyyy") '-- Book Vdt
'''                    msGrid.TextMatrix(msGrid.Row, 3) = RsBk!ItName  '-- Count
'''                    msGrid.TextMatrix(msGrid.Row, 9) = .SearchMultiRetCol(4)  '-- Bal Bag
'''                    msGrid.TextMatrix(msGrid.Row, 10) = .SearchMultiRetCol(5) '-- Bal Wt
'''                    msGrid.TextMatrix(msGrid.Row, 19) = RsBk!VSubItCtrlNo  '--  P ItCtrl No =Book ItCtrl No
'''                    msGrid.TextMatrix(msGrid.Row, 11) = RsBk!vsubRt  '-- Rate
'''                    msGrid.TextMatrix(msGrid.Row, 12) = RsBk!vsubRTPER  '--Rate Per
'''                    CalGridAmt
'''                    CalAmount
'''              End If
'''        End If
'''    End With
    
'    With gClsSearch
'        .SearchMultiField QStr, "", Array("Booking No.", "Date", "Count", "Bal Bag", "Bal Wt.", "Rate", "Wt", "itctrlno"), Array(1000, 1500, 1500, 1000, 1000, 1000, 0, 0), "", Chr, "tblVousub.Vdt,tblVousub.vno,VSUBCTRNO ", 0, 0, False, , True
'        If .SearchMultiRetCol(0) <> "" Then
'            QStr = "select tblvousub.*,tblmastItem.ItName from tblvousub,tblMastItem where " _
'            & " ItCode=VSUBITCODE and VSUBITCTRLNO= " & CLng(.SearchMultiRetCol(7))
'              If GProcRstOpen(RsBk, QStr, "R") > 0 Then
'                    txtGrid = ""
'                    txtGrid = CLng(.SearchMultiRetCol(0)) '--Book Vno
'                    msGrid.TextMatrix(msGrid.Row, 1) = CLng(.SearchMultiRetCol(0)) '--Book Vno
'                    msGrid.TextMatrix(msGrid.Row, 2) = Format(.SearchMultiRetCol(1), "dd/mm/yyyy") '-- Book Vdt
'                    msGrid.TextMatrix(msGrid.Row, 3) = RsBk!ItName  '-- Count
'              '      msGrid.TextMatrix(msGrid.Row, 9) = .SearchMultiRetCol(3)  '-- Bal Bag
'               '     msGrid.TextMatrix(msGrid.Row, 10) = .SearchMultiRetCol(4) '-- Bal Wt
'                    msGrid.TextMatrix(msGrid.Row, 4) = RsBk!vsubRt  '-- Rate
'
'                    msGrid.TextMatrix(msGrid.Row, 19) = RsBk!VSubItCtrlNo  '--  P ItCtrl No =Book ItCtrl No
''                    msGrid.TextMatrix(msGrid.Row, 11) = RsBk!vsubRt  '-- Rate
'                    msGrid.TextMatrix(msGrid.Row, 12) = RsBk!vsubRTPER  '--Rate Per
'                    CalGridAmt
'                    CalAmount
'              End If
'        End If
'    End With
'End If
End Sub
Private Sub BookingData()
Dim RsG As Recordset
Dim WStdPack As Double
Dim wItcode As Long
Dim wMillCode As Long
Dim wParty As Long
Dim wBroker As Long
Dim BalBag As Integer
Dim BalWt As Double
Set RsG = New Recordset
Dim i As Integer
wMillCode = GProcGetColumnValue("tblMastAccount", "AcName", txttemp(7), "S", "AcCode", "N")
wParty = GProcGetColumnValue("tblMastAccount", "AcName", txttemp(5), "S", "AcCode", "N")
wBroker = GProcGetColumnValue("tblMastAccount", "AcName", txttemp(8), "S", "AcCode", "N")

If Val(msGrid.TextMatrix(msGrid.Row, 1)) <> 0 Then
'----
   gCmd.CommandText = "PrcPreparePendingBookingMillBill"
   gCmd.Parameters.Refresh
   gCmd.Parameters("@Fr_dt") = Format(gCYSDate - 400, "yyyy/mm/dd")
   gCmd.Parameters("@To_dt") = Format(gCYEDate, "yyyy/mm/dd")
   gCmd.Parameters("@UCode") = gUserName
   gCmd.Parameters("@VFirm") = Trim(gCCode)
   gCmd.Parameters("@Vtype") = "OM"
   gCmd.Parameters("@BkNo") = Val(msGrid.TextMatrix(msGrid.Row, 1))
   gCmd.Parameters("@BkYear") = Val(msGrid.TextMatrix(msGrid.Row, 2))
   gCmd.Parameters("@RptTp") = "P"
   gCmd.Parameters("@BkSrNo") = Val(msGrid.TextMatrix(msGrid.Row, 8))
   gCmd.Execute
   X = "Select Vno,PartyName,BrokerName,MillCode,Nar1,Amt3,Amt4-Amt8,Dt1, " _
       & "PartyCode,BrokerCode,MillName,Amt6 ,Amt7,Vtype,Vyear,Amt5,Vdt,Dt2,format(Amt9,'#####'),format(Amt11,'#####'),amt12 from TmpGentbl where TmpGentbl.UserName = '" & gUserName & "' and (amt4 - amt8) > 0   "

       i = GProcRstOpen(tmpRst, X, "R")
       If i > 0 Then
          gClsSearch.SearchMultiField X, "", Array("Booking No", "Party", "Broker", "Mill", "Count", "Rate", "Sh.Bag", "Sh.Date", "", "", "", "Std.Pck", "RtPer", "BkTp", "BkYear", "BkSrNo", "Bk.Dt", "Desp. Dt", "Desp. Bag", "BkDespSr", ""), Array(0, 2500, 1800, 1000, 2000, 1000, 800, 1500, 0, 0, 0, 0, 0, 0, 0, 0, 0, 1500, 800, 0, 0), "", "", "Vno,Vdt,dt1 ", 0, 0, False, , True
          If Val(gClsSearch.SearchMultiRetCol(0)) <> 0 Then
            If txttemp(5) <> "" And gClsSearch.SearchMultiRetCol(1) <> txttemp(5) Then
               MsgBox ("Please Check Booking Party ....")
               msGrid.Col = 1
               msGrid.SetFocus
               Exit Sub
            End If
            If txttemp(8) <> "" And gClsSearch.SearchMultiRetCol(2) <> txttemp(8) Then
               MsgBox ("Please Check Booking Broker ....")
               msGrid.Col = 1
               msGrid.SetFocus
               Exit Sub
            End If
            txttemp(5) = gClsSearch.SearchMultiRetCol(1)
            txttemp(8) = gClsSearch.SearchMultiRetCol(2)
            msGrid.TextMatrix(msGrid.Row, 2) = gClsSearch.SearchMultiRetCol(14) '-- Vyear
            msGrid.TextMatrix(msGrid.Row, 3) = CStr(Format(gClsSearch.SearchMultiRetCol(16), "dd/mm/yyyy"))
            msGrid.TextMatrix(msGrid.Row, 4) = gClsSearch.SearchMultiRetCol(4) '-- Item
            msGrid.TextMatrix(msGrid.Row, 7) = gClsSearch.SearchMultiRetCol(13) '-- Vtype
            msGrid.TextMatrix(msGrid.Row, 8) = gClsSearch.SearchMultiRetCol(15) '-- BkSrNo
            msGrid.TextMatrix(msGrid.Row, 9) = gClsSearch.SearchMultiRetCol(19) '-- BkSrNo
            If FormAction = vbDataActionAddNew Then
               If Val(gClsSearch.SearchMultiRetCol(18)) = 0 Or IsNull(gClsSearch.SearchMultiRetCol(18)) Then
                   msGrid.TextMatrix(msGrid.Row, 5) = gClsSearch.SearchMultiRetCol(6)
               Else
                   msGrid.TextMatrix(msGrid.Row, 5) = gClsSearch.SearchMultiRetCol(18)
               End If
               msGrid.TextMatrix(msGrid.Row, 6) = BalWt
            End If
            If FormAction = vbDataActionAddNew Then
               chkIsPaymenttoMill.Value = gClsSearch.SearchMultiRetCol(20)
            End If
            txtCode(5) = GProcGetColumnValue("tblMastAccount", "AcName", "'" & txttemp(5) & "'", "N", "AcOurCode", "S")
            txtCode(8) = GProcGetColumnValue("tblMastAccount", "AcName", "'" & txttemp(8) & "'", "N", "AcOurCode", "S")
            wItcode = GProcGetColumnValue("tblMastItem", "ItName", "" & msGrid.TextMatrix(msGrid.Row, 4) & "", "S", "ItCode", "N", "ItMillCode = " & wMillCode & "")
            WStdPack = GProcGetColumnValue("tblMastItem", "Itcode", "" & wItcode & "", "N", "ItStdpack", "N")
            msGrid.TextMatrix(msGrid.Row, 6) = Format(gClsSearch.SearchMultiRetCol(6) * WStdPack, WtStr)
            lblTinNo.Caption = GProcGetColumnValue("tblMastAccount", "AcName", "'" & txttemp(5) & "'", "N", "AcGSTIN", "S")
            MsOms = GProcGetColumnValue("tblMastAccount", "AcName", "'" & txttemp(5) & "'", "N", "AcMsOmsParty", "N")
            If MsOms = 0 Then
               lblMsOms.Caption = "M.S."
            Else
               lblMsOms.Caption = "O.M.S."
            End If
            msGrid.Col = 5
             txtGrid.Locked = False

            msGrid.SetFocus
        Else
            msGrid.TextMatrix(msGrid.Row, 0) = ""
        End If
'-----

'''   X = "Select BkParty,BkBroker,BkItCode,BkBag-BkCbag as BkBg,BkWt-BkCWt as BkKg ,BkRt,BkRtPer,Vdt,Vtype,VYear from tblBooking where VNo=" & msGrid.TextMatrix(msGrid.Row, 1) & " and vtype = 'OM' and VFirm  = '" & gCCode & "' and VYear<= " & gCYear & " and BkMillCode = " & wMillCode & " "
'''   If GProcRstOpen(RsG, X, "R") > 0 Then
'''      If txttemp(5) <> "" And (RsG.Fields("BkParty") <> wParty) Then
'''          MsgBox ("Please Check Booking Party ....")
'''          msGrid.Col = 1
'''          msGrid.SetFocus
'''          Exit Sub
'''      End If
'''      If (txttemp(8) <> "") And RsG.Fields("BkBroker") <> wBroker Then
'''          MsgBox ("Please Check Booking Broker ....")
'''          msGrid.Col = 1
'''          msGrid.SetFocus
'''          Exit Sub
'''      End If
'''      BalBag = RsG.Fields("BkBg")
'''      BalWt = RsG.Fields("BkKg")
'''      BalBag = BalBag - GProcGetColumnValue("tblSaleSub", "SlSubItCode", RsG.Fields("BkItCode"), "N", "sum(SlSubBag)", "N", "VType = '" & frmTypeStr & "' and  SlSubBookNo = " & msGrid.TextMatrix(msGrid.Row, 1) & "  and SlSubBkType = 'OM' and SlSubBkVyear =  " & RsG.Fields("Vyear") & "   and VFirm = '" & gCCode & "'  and  not (vno = " & txt(0) & " and VYear=" & gCYear & " and VFirm = '" & gCCode & "'  and vtype = '" & frmTypeStr & "')")
'''      BalWt = BalWt - GProcGetColumnValue("tblSaleSub", "SlSubItCode", RsG.Fields("BkItCode"), "N", "sum(SlSubWt)", "N", "VType = '" & frmTypeStr & "' and  SlSubBookNo = " & msGrid.TextMatrix(msGrid.Row, 1) & "  and SlSubBkType = 'OM' and SlSubBkVyear =  " & RsG.Fields("Vyear") & "   and VFirm = '" & gCCode & "'  and  not (vno = " & txt(0) & " and VYear=" & gCYear & " and VFirm = '" & gCCode & "'  and vtype = '" & frmTypeStr & "')")
'''      lblBag.Visible = True
'''      lblBalWt.Visible = True
'''      lblBalBag.Visible = True
'''      lblWeight.Visible = True
'''      lblBalBag.Caption = BalBag
'''      lblBalWt.Caption = BalWt
'''      txttemp(5) = GProcGetColumnValue("tblMastAccount", "AcCode", RsG.Fields("BkParty"), "N", "AcName", "S")
'''      txttemp(8) = GProcGetColumnValue("tblMastAccount", "AcCode", RsG.Fields("BkBroker"), "N", "AcName", "S")
'''      msGrid.TextMatrix(msGrid.Row, 2) = CStr(Format(RsG.Fields("Vdt"), "dd/mm/yyyy"))
'''      msGrid.TextMatrix(msGrid.Row, 3) = GProcGetColumnValue("tblMastItem", "ItCode", RsG.Fields("BkItCode"), "N", "ItName", "S", " ItMillCode=" & wMillCode) '--Item
'''      msGrid.TextMatrix(msGrid.Row, 6) = RsG.Fields("Vtype")
'''      msGrid.TextMatrix(msGrid.Row, 7) = RsG.Fields("Vyear")
'''      If FormAction = vbDataActionAddNew Then
'''         msGrid.TextMatrix(msGrid.Row, 4) = BalBag
'''         msGrid.TextMatrix(msGrid.Row, 5) = BalWt
'''      End If
'''      txtCode(5) = GProcGetColumnValue("tblMastAccount", "AcName", "'" & txttemp(5) & "'", "N", "AcOurCode", "S")
'''      txtCode(8) = GProcGetColumnValue("tblMastAccount", "AcName", "'" & txttemp(8) & "'", "N", "AcOurCode", "S")
'''      lblTinNo.Caption = GProcGetColumnValue("tblMastAccount", "AcName", "'" & txttemp(5) & "'", "N", "AcBST", "S")
'''      MsOms = GProcGetColumnValue("tblMastAccount", "AcName", "'" & txttemp(5) & "'", "N", "AcMsOmsParty", "N")
'''      If MsOms = 0 Then
'''         lblMsOms.Caption = "M.S."
'''      Else
'''         lblMsOms.Caption = "O.M.S."
'''      End If
'''      msGrid.Col = 4
'''      msGrid.SetFocus
   Else
      msGrid.TextMatrix(msGrid.Row, 3) = ""
      msGrid.TextMatrix(msGrid.Row, 4) = ""
      msGrid.TextMatrix(msGrid.Row, 7) = ""
      msGrid.TextMatrix(msGrid.Row, 2) = 0
      msGrid.TextMatrix(msGrid.Row, 5) = 0
      msGrid.TextMatrix(msGrid.Row, 6) = 0
      lblBalBag.Caption = 0
      lblBalWt.Caption = 0
'      txttemp(5) = ""
'      txttemp(8) = ""
'      txtCode(5) = ""
'      txtCode(8) = ""
'      lblTinNo.Caption = ""
'      lblMsOms.Caption = ""
      MsgBox ("Booking No Not Found....")
      msGrid.Col = 1
      msGrid.SetFocus
   End If
End If
End Sub
Private Function GBillChr()
Dim BilSr As String
Dim NarrCdBlTp As Long
NarrCdBlTp = GProcGetColumnValue("tblMastNarration", "Narration", "Mill Bill", "S", "NarrCode", "N", "NarrType = 'X'")
SalesAccode = GProcGetColumnValue("tblMastAccount", "AcName", txttemp(7), "S", "AcCode", "N")
If SalesAccode > 1 Then
    BilSr = ""
    BilSr = GProcGetColumnValue("tblMastBillSerial", "MillCode", "" & SalesAccode & "", "N", "BillSr", "S", "SaleType = " & NarrCdBlTp & " and VFirm = '" & gCCode & "' ")
    If BilSr = "" Then
       MsgBox ("Check Bill Sr not given...")
    End If
    lblSr.Caption = BilSr
End If
End Function

