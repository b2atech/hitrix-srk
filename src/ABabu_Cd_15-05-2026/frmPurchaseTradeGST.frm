VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form frmPurchaseTradeGST 
   BackColor       =   &H00FBF2E1&
   BorderStyle     =   0  'None
   Caption         =   "Purchase Entry"
   ClientHeight    =   9240
   ClientLeft      =   0
   ClientTop       =   15
   ClientWidth     =   13830
   FillColor       =   &H000040C0&
   FillStyle       =   0  'Solid
   Icon            =   "frmPurchaseTradeGST.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   9240
   ScaleWidth      =   13830
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame1 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFC0C0&
      Caption         =   "Packing Slip"
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
      Left            =   2820
      TabIndex        =   160
      Top             =   2520
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
         Left            =   4245
         TabIndex        =   165
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
         Left            =   2430
         TabIndex        =   164
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
         TabIndex        =   163
         Top             =   1065
         Visible         =   0   'False
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
         Left            =   720
         TabIndex        =   162
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
         TabIndex        =   161
         Top             =   495
         Width           =   2055
      End
      Begin VB.Label Label39 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "To     G. P.  No     :"
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
         Left            =   915
         TabIndex        =   167
         Top             =   1155
         Visible         =   0   'False
         Width           =   1740
      End
      Begin VB.Label Label38 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Voucher No         :"
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
         Left            =   915
         TabIndex        =   166
         Top             =   585
         Width           =   1710
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
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Index           =   72
      Left            =   2280
      TabIndex        =   43
      Text            =   "72"
      Top             =   6120
      Width           =   1485
   End
   Begin VB.TextBox txtGodown 
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
      Left            =   9885
      TabIndex        =   14
      Text            =   "txtGodown"
      ToolTipText     =   "Press F3 for Creation & Press F2 for Selection"
      Top             =   1440
      Width           =   3495
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
      Index           =   71
      Left            =   14640
      TabIndex        =   154
      Text            =   "71"
      Top             =   4380
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
      Height          =   405
      Index           =   65
      Left            =   2280
      TabIndex        =   47
      Text            =   "txtCode(65)"
      ToolTipText     =   "Press F3 for Creation"
      Top             =   7920
      Width           =   1485
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
      Index           =   65
      Left            =   2280
      Locked          =   -1  'True
      TabIndex        =   48
      Text            =   "txttemp(65)"
      ToolTipText     =   "Press F3 for Creation"
      Top             =   8400
      Width           =   3000
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
      Index           =   70
      Left            =   9480
      TabIndex        =   38
      Text            =   "70"
      Top             =   7680
      Width           =   2265
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
      Height          =   375
      Index           =   69
      Left            =   6540
      TabIndex        =   36
      Text            =   "69"
      Top             =   7695
      Width           =   735
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
      Height          =   375
      Index           =   68
      Left            =   7680
      TabIndex        =   37
      Text            =   "68"
      Top             =   7695
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
      Index           =   67
      Left            =   14640
      TabIndex        =   145
      Text            =   "67"
      Top             =   3960
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
      Index           =   66
      Left            =   14640
      TabIndex        =   144
      Text            =   "66"
      Top             =   3600
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
      Index           =   65
      Left            =   14640
      TabIndex        =   143
      Text            =   "65"
      Top             =   3300
      Visible         =   0   'False
      Width           =   420
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
      Index           =   64
      Left            =   2280
      TabIndex        =   45
      Text            =   "64"
      Top             =   7080
      Width           =   1485
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
      Index           =   63
      Left            =   2280
      TabIndex        =   44
      Text            =   "63"
      Top             =   6600
      Width           =   1485
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
      Index           =   62
      Left            =   14700
      TabIndex        =   46
      Text            =   "62"
      Top             =   2220
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
      Index           =   61
      Left            =   14640
      TabIndex        =   142
      Text            =   "61"
      Top             =   1920
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
      Index           =   60
      Left            =   14640
      TabIndex        =   141
      Text            =   "60"
      Top             =   1560
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
      Index           =   59
      Left            =   14640
      TabIndex        =   140
      Text            =   "59"
      Top             =   1200
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
      Index           =   58
      Left            =   14640
      TabIndex        =   139
      Text            =   "58"
      Top             =   840
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
      Index           =   57
      Left            =   14640
      TabIndex        =   138
      Text            =   "57"
      Top             =   480
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
      Index           =   56
      Left            =   14640
      TabIndex        =   137
      Text            =   "56"
      Top             =   120
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
      Index           =   55
      Left            =   13920
      TabIndex        =   136
      Text            =   "55"
      Top             =   5220
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
      Index           =   54
      Left            =   13920
      TabIndex        =   135
      Text            =   "54"
      Top             =   4860
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
      Index           =   53
      Left            =   13920
      TabIndex        =   134
      Text            =   "53"
      Top             =   4500
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
      Index           =   52
      Left            =   13920
      TabIndex        =   133
      Text            =   "52"
      Top             =   4140
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
      Index           =   51
      Left            =   13920
      TabIndex        =   132
      Text            =   "51"
      Top             =   3780
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
      Index           =   50
      Left            =   13920
      TabIndex        =   131
      Text            =   "50"
      Top             =   3420
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
      Index           =   49
      Left            =   13920
      TabIndex        =   130
      Text            =   "49"
      Top             =   3060
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
      Index           =   48
      Left            =   13920
      TabIndex        =   129
      Text            =   "48"
      Top             =   2700
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
      Index           =   47
      Left            =   13920
      TabIndex        =   128
      Text            =   "47"
      Top             =   2340
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
      Index           =   46
      Left            =   13920
      TabIndex        =   127
      Text            =   "46"
      Top             =   1980
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
      Index           =   45
      Left            =   13920
      TabIndex        =   126
      Text            =   "45"
      Top             =   1620
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
      Index           =   44
      Left            =   13920
      TabIndex        =   125
      Text            =   "44"
      Top             =   1260
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
      Index           =   43
      Left            =   13920
      TabIndex        =   124
      Text            =   "43"
      Top             =   900
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
      Index           =   42
      Left            =   13920
      TabIndex        =   123
      Text            =   "42"
      Top             =   540
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
      Index           =   41
      Left            =   13920
      TabIndex        =   122
      Text            =   "41"
      Top             =   180
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
      Index           =   40
      Left            =   16470
      TabIndex        =   121
      Text            =   "40"
      Top             =   7875
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
      Left            =   17460
      TabIndex        =   120
      Text            =   "39"
      Top             =   7920
      Visible         =   0   'False
      Width           =   570
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
      TabIndex        =   8
      Text            =   "txtCode(7)"
      ToolTipText     =   "Press F3 for Creation"
      Top             =   2205
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
      TabIndex        =   4
      Text            =   "txtCode(5)"
      ToolTipText     =   "Press F3 for Creation"
      Top             =   1080
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
      TabIndex        =   6
      Text            =   "txtCode(8)"
      ToolTipText     =   "Press F3 for Creation"
      Top             =   1800
      Visible         =   0   'False
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
      Index           =   38
      Left            =   9495
      TabIndex        =   35
      Text            =   "38"
      Top             =   7245
      Width           =   2265
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
      Index           =   37
      Left            =   7020
      TabIndex        =   34
      Text            =   "37"
      Top             =   7245
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
      Height          =   375
      Index           =   36
      Left            =   17370
      TabIndex        =   112
      Text            =   "36"
      Top             =   7515
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
      Height          =   345
      Index           =   14
      Left            =   12450
      Style           =   1  'Graphical
      TabIndex        =   110
      Top             =   8145
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
      Left            =   17415
      TabIndex        =   109
      Text            =   "35"
      Top             =   7110
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
      Index           =   22
      Left            =   17190
      TabIndex        =   106
      Text            =   "22"
      Top             =   4275
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.TextBox txt 
      Alignment       =   1  'Right Justify
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   450
      Index           =   16
      Left            =   8190
      TabIndex        =   42
      Text            =   "16"
      Top             =   8505
      Width           =   3570
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
      Index           =   15
      Left            =   9495
      TabIndex        =   25
      Text            =   "15"
      Top             =   4680
      Width           =   2265
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
      Left            =   17190
      TabIndex        =   104
      Text            =   "14"
      Top             =   3915
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
      Index           =   13
      Left            =   9885
      TabIndex        =   16
      Text            =   "13"
      Top             =   1845
      Width           =   1830
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
      Left            =   9885
      TabIndex        =   13
      Text            =   "txttemp(11)"
      ToolTipText     =   "Press F3 for Creation & Press F2 for Selection"
      Top             =   1035
      Width           =   3495
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
      Index           =   28
      Left            =   15480
      TabIndex        =   39
      Text            =   "28"
      Top             =   6945
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
      Height          =   360
      Index           =   29
      Left            =   15480
      TabIndex        =   40
      Text            =   "29"
      Top             =   7335
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.CommandButton CmdBagDetails 
      Caption         =   "Cartoon &Bag Details"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1080
      Left            =   3900
      TabIndex        =   97
      Top             =   5940
      Width           =   1530
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
      Left            =   17235
      TabIndex        =   54
      Text            =   "8"
      Top             =   2745
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
      Left            =   17415
      TabIndex        =   96
      Text            =   "34"
      Top             =   6660
      Visible         =   0   'False
      Width           =   570
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
      Index           =   19
      Left            =   15495
      TabIndex        =   28
      Text            =   "19"
      Top             =   7905
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
      Index           =   33
      Left            =   17370
      TabIndex        =   95
      Text            =   "33"
      Top             =   6255
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
      Left            =   17325
      TabIndex        =   94
      Text            =   "32"
      Top             =   5850
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
      Height          =   390
      Index           =   8
      Left            =   2505
      TabIndex        =   7
      Text            =   "txttemp(8)"
      ToolTipText     =   "Press F3 for Creation"
      Top             =   1800
      Visible         =   0   'False
      Width           =   5910
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
      Left            =   2505
      TabIndex        =   5
      Text            =   "txttemp(5)"
      ToolTipText     =   "Press F3 for Creation"
      Top             =   1080
      Width           =   5910
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
      Left            =   2505
      TabIndex        =   9
      Text            =   "txttemp(7)"
      ToolTipText     =   "Press F3 for Creation"
      Top             =   2205
      Width           =   5910
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
      Left            =   17010
      TabIndex        =   22
      Text            =   "txtCrBalance"
      Top             =   135
      Visible         =   0   'False
      Width           =   525
   End
   Begin VB.CommandButton cmdLedger 
      Caption         =   "Ledger"
      Height          =   375
      Left            =   15780
      TabIndex        =   93
      Top             =   9360
      Visible         =   0   'False
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
      Index           =   27
      Left            =   9495
      TabIndex        =   33
      Text            =   "27"
      Top             =   6810
      Width           =   2265
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
      Index           =   26
      Left            =   7005
      TabIndex        =   32
      Text            =   "26"
      Top             =   6825
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
      Height          =   375
      Index           =   25
      Left            =   17235
      TabIndex        =   69
      Text            =   "25"
      Top             =   4680
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
      Left            =   9495
      TabIndex        =   29
      Text            =   "21"
      Top             =   5940
      Width           =   2265
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
      TabIndex        =   24
      Top             =   3900
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
      Left            =   17145
      TabIndex        =   91
      Text            =   "1"
      Top             =   405
      Visible         =   0   'False
      Width           =   420
   End
   Begin VB.TextBox txt 
      Alignment       =   1  'Right Justify
      Height          =   315
      Index           =   2
      Left            =   17100
      TabIndex        =   89
      Text            =   "2"
      Top             =   855
      Visible         =   0   'False
      Width           =   525
   End
   Begin VB.TextBox txt 
      Alignment       =   1  'Right Justify
      Height          =   315
      Index           =   5
      Left            =   17190
      TabIndex        =   88
      Text            =   "5"
      Top             =   1800
      Visible         =   0   'False
      Width           =   525
   End
   Begin VB.TextBox txt 
      Alignment       =   1  'Right Justify
      Height          =   315
      Index           =   3
      Left            =   17100
      TabIndex        =   87
      Text            =   "3"
      Top             =   1215
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
      Left            =   15600
      Style           =   1  'Graphical
      TabIndex        =   83
      Top             =   9120
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
      Left            =   15720
      Style           =   1  'Graphical
      TabIndex        =   82
      ToolTipText     =   "To Add New Record"
      Top             =   9180
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
      Height          =   345
      Index           =   10
      Left            =   12450
      Style           =   1  'Graphical
      TabIndex        =   80
      ToolTipText     =   "Exit"
      Top             =   8565
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
      Height          =   345
      Index           =   13
      Left            =   12450
      Style           =   1  'Graphical
      TabIndex        =   73
      ToolTipText     =   "To Save Record"
      Top             =   7695
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
      Height          =   345
      Index           =   9
      Left            =   12450
      Style           =   1  'Graphical
      TabIndex        =   81
      Top             =   7260
      Width           =   870
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
      Height          =   345
      Index           =   8
      Left            =   12450
      Style           =   1  'Graphical
      TabIndex        =   72
      ToolTipText     =   "To Delete Record"
      Top             =   6825
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
      Height          =   345
      Index           =   7
      Left            =   12450
      Style           =   1  'Graphical
      TabIndex        =   71
      ToolTipText     =   "To Modify Record"
      Top             =   6375
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
      Height          =   345
      Index           =   6
      Left            =   12450
      Style           =   1  'Graphical
      TabIndex        =   70
      ToolTipText     =   "To Add New Record"
      Top             =   5940
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
      Left            =   16905
      Style           =   1  'Graphical
      TabIndex        =   74
      Top             =   8400
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
      Height          =   345
      Index           =   4
      Left            =   12450
      Style           =   1  'Graphical
      TabIndex        =   75
      ToolTipText     =   "To Find Record"
      Top             =   5505
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
      Height          =   345
      Index           =   3
      Left            =   12945
      Style           =   1  'Graphical
      TabIndex        =   79
      ToolTipText     =   "Last"
      Top             =   5070
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
      Height          =   345
      Index           =   2
      Left            =   12450
      Style           =   1  'Graphical
      TabIndex        =   78
      ToolTipText     =   "First"
      Top             =   5070
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
      Height          =   345
      Index           =   1
      Left            =   12945
      Style           =   1  'Graphical
      TabIndex        =   77
      ToolTipText     =   "Next"
      Top             =   4680
      Width           =   375
   End
   Begin VB.TextBox txt 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Index           =   0
      Left            =   6585
      TabIndex        =   3
      Text            =   "0"
      Top             =   645
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
      Height          =   345
      Index           =   0
      Left            =   12450
      Style           =   1  'Graphical
      TabIndex        =   76
      ToolTipText     =   "Previous"
      Top             =   4680
      Width           =   375
   End
   Begin VB.TextBox txt 
      Alignment       =   1  'Right Justify
      Height          =   315
      Index           =   4
      Left            =   17100
      TabIndex        =   86
      Text            =   "4"
      Top             =   1485
      Visible         =   0   'False
      Width           =   525
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
      Left            =   1410
      Locked          =   -1  'True
      TabIndex        =   49
      Text            =   "txttemp(6)"
      ToolTipText     =   "Press F3 for Creation"
      Top             =   4620
      Width           =   4020
   End
   Begin VB.TextBox txt 
      Height          =   315
      Index           =   6
      Left            =   17235
      TabIndex        =   85
      Text            =   "6"
      Top             =   2115
      Visible         =   0   'False
      Width           =   435
   End
   Begin VB.TextBox txt 
      Alignment       =   1  'Right Justify
      Height          =   315
      Index           =   7
      Left            =   17145
      TabIndex        =   84
      Text            =   "7"
      Top             =   2475
      Visible         =   0   'False
      Width           =   525
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
      Height          =   375
      Index           =   9
      Left            =   9885
      TabIndex        =   10
      Text            =   "9"
      Top             =   600
      Width           =   1830
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
      Left            =   17190
      TabIndex        =   56
      Text            =   "10"
      Top             =   3150
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
      Left            =   17280
      TabIndex        =   67
      Text            =   "11"
      Top             =   3465
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
      Index           =   12
      Left            =   15360
      TabIndex        =   15
      Text            =   "12"
      Top             =   300
      Visible         =   0   'False
      Width           =   555
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
      Index           =   23
      Left            =   7005
      TabIndex        =   30
      Text            =   "23"
      Top             =   6390
      Width           =   735
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
      Index           =   24
      Left            =   9495
      TabIndex        =   31
      Text            =   "24"
      Top             =   6375
      Width           =   2265
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
      Index           =   20
      Left            =   15525
      TabIndex        =   50
      Text            =   "20"
      Top             =   8340
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
      Index           =   30
      Left            =   9495
      TabIndex        =   41
      Text            =   "30"
      Top             =   8085
      Width           =   2265
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
      Left            =   9885
      TabIndex        =   18
      Text            =   "txt(31)"
      Top             =   2250
      Visible         =   0   'False
      Width           =   1830
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
      Index           =   17
      Left            =   9495
      TabIndex        =   26
      Text            =   "17"
      Top             =   5085
      Width           =   2265
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
      Index           =   18
      Left            =   9495
      TabIndex        =   27
      Text            =   "18"
      Top             =   5520
      Width           =   2265
   End
   Begin MSMask.MaskEdBox mskDt 
      Height          =   375
      Index           =   2
      Left            =   16260
      TabIndex        =   2
      Top             =   495
      Visible         =   0   'False
      Width           =   360
      _ExtentX        =   635
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
      Left            =   16320
      TabIndex        =   12
      Top             =   945
      Visible         =   0   'False
      Width           =   345
      _ExtentX        =   609
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
      Index           =   32
      Left            =   16320
      TabIndex        =   21
      Top             =   1920
      Visible         =   0   'False
      Width           =   525
      _ExtentX        =   926
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
      Bindings        =   "frmPurchaseTradeGST.frx":058A
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
      Height          =   1590
      Left            =   315
      TabIndex        =   23
      Top             =   2835
      Width           =   13275
      _ExtentX        =   23416
      _ExtentY        =   2805
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
      Left            =   16320
      TabIndex        =   20
      Top             =   1380
      Visible         =   0   'False
      Width           =   345
      _ExtentX        =   609
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
      Left            =   1305
      TabIndex        =   1
      Top             =   667
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
      Format          =   122683395
      CurrentDate     =   -273
   End
   Begin MSComCtl2.DTPicker dtpDt 
      Height          =   360
      Index           =   10
      Left            =   11745
      TabIndex        =   11
      Top             =   607
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
      Format          =   122683395
      CurrentDate     =   -273
   End
   Begin MSComCtl2.DTPicker dtpDt 
      Height          =   360
      Index           =   14
      Left            =   11745
      TabIndex        =   17
      Top             =   1845
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
      Format          =   122683395
      CurrentDate     =   -273
   End
   Begin MSComCtl2.DTPicker dtpDt 
      Height          =   360
      Index           =   32
      Left            =   11745
      TabIndex        =   19
      Top             =   2250
      Visible         =   0   'False
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
      Format          =   122683395
      CurrentDate     =   -273
   End
   Begin VB.Label Label19 
      BackStyle       =   0  'Transparent
      Caption         =   "TDS On Amount   :"
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
      Left            =   420
      TabIndex        =   159
      Top             =   6180
      Width           =   1890
   End
   Begin VB.Label lblPurWeight 
      BackColor       =   &H00EDFDFE&
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
      ForeColor       =   &H00000080&
      Height          =   255
      Left            =   6600
      TabIndex        =   158
      Top             =   5160
      Width           =   1335
   End
   Begin VB.Label lblPurBgas 
      BackColor       =   &H00EDFDFE&
      Caption         =   "Bags "
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
      Height          =   255
      Left            =   6600
      TabIndex        =   157
      Top             =   4740
      Width           =   1335
   End
   Begin VB.Label lblTKgs 
      BackColor       =   &H00EDFDFE&
      Caption         =   "T.Kgs.   :"
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
      Height          =   255
      Left            =   5700
      TabIndex        =   156
      Top             =   5160
      Width           =   855
   End
   Begin VB.Label lblTBags 
      BackColor       =   &H00EDFDFE&
      Caption         =   "T.Bags  :"
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
      Height          =   255
      Left            =   5700
      TabIndex        =   155
      Top             =   4740
      Width           =   795
   End
   Begin VB.Label lblPurchAmt 
      BackStyle       =   0  'Transparent
      Caption         =   "Total Purchase  Amount"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF00FF&
      Height          =   375
      Left            =   2340
      TabIndex        =   153
      Top             =   5460
      Width           =   3030
   End
   Begin VB.Label Label18 
      BackStyle       =   0  'Transparent
      Caption         =   "Total Purchase     :"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF00FF&
      Height          =   375
      Left            =   240
      TabIndex        =   152
      Top             =   5460
      Width           =   1890
   End
   Begin VB.Label Label34 
      BackStyle       =   0  'Transparent
      Caption         =   "TDS Rate              :"
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
      Left            =   420
      TabIndex        =   151
      Top             =   6660
      Width           =   1890
   End
   Begin VB.Label Label35 
      BackStyle       =   0  'Transparent
      Caption         =   "TDS Amt.              :"
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
      Left            =   420
      TabIndex        =   150
      Top             =   7140
      Width           =   1890
   End
   Begin VB.Label Label36 
      BackStyle       =   0  'Transparent
      Caption         =   "TDS A/c  Code     :"
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
      Left            =   420
      TabIndex        =   149
      Top             =   7980
      Width           =   1890
   End
   Begin VB.Label Label37 
      BackStyle       =   0  'Transparent
      Caption         =   "TDS A/c               :"
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
      Left            =   420
      TabIndex        =   148
      Top             =   8460
      Width           =   1890
   End
   Begin VB.Label Label16 
      BackStyle       =   0  'Transparent
      Caption         =   "On "
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000040&
      Height          =   375
      Left            =   7320
      TabIndex        =   147
      Top             =   7695
      Width           =   825
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "TCS @  : "
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000040&
      Height          =   375
      Left            =   5700
      TabIndex        =   146
      Top             =   7695
      Width           =   825
   End
   Begin VB.Label lblIGstRate 
      BackStyle       =   0  'Transparent
      Caption         =   "I  GST @  : "
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C000C0&
      Height          =   375
      Left            =   5940
      TabIndex        =   119
      Top             =   7260
      Width           =   1065
   End
   Begin VB.Label lblCGstRate 
      BackStyle       =   0  'Transparent
      Caption         =   "C GST @  : "
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C000C0&
      Height          =   375
      Left            =   5940
      TabIndex        =   118
      Top             =   6825
      Width           =   1065
   End
   Begin VB.Label lblSGstRate 
      BackStyle       =   0  'Transparent
      Caption         =   "S GST @  : "
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C000C0&
      Height          =   375
      Left            =   5940
      TabIndex        =   117
      Top             =   6390
      Width           =   1065
   End
   Begin VB.Label lbl3 
      BackStyle       =   0  'Transparent
      Caption         =   "%"
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
      Left            =   7830
      TabIndex        =   116
      Top             =   7260
      Width           =   255
   End
   Begin VB.Label lbl2 
      BackStyle       =   0  'Transparent
      Caption         =   "%"
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
      Left            =   7830
      TabIndex        =   115
      Top             =   6825
      Width           =   255
   End
   Begin VB.Label lblIgstAmt 
      BackStyle       =   0  'Transparent
      Caption         =   "I  GST Amt   : "
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
      Left            =   8130
      TabIndex        =   114
      Top             =   7260
      Width           =   1335
   End
   Begin VB.Label lblCgstAmt 
      BackStyle       =   0  'Transparent
      Caption         =   "C GST Amt   : "
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
      Left            =   8130
      TabIndex        =   113
      Top             =   6825
      Width           =   1335
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
      Left            =   7590
      TabIndex        =   111
      Top             =   1500
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
      Left            =   3105
      TabIndex        =   108
      Top             =   1500
      Width           =   2775
   End
   Begin VB.Label Label21 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "TIN  :  "
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
      Left            =   2475
      TabIndex        =   107
      Top             =   1500
      Width           =   615
   End
   Begin VB.Label Label20 
      BackStyle       =   0  'Transparent
      Caption         =   "Date     :    "
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
      Left            =   495
      TabIndex        =   105
      Top             =   675
      Width           =   870
   End
   Begin VB.Label Label14 
      BackStyle       =   0  'Transparent
      Caption         =   "L.R. No          :    "
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
      Left            =   8805
      TabIndex        =   103
      Top             =   1905
      Width           =   1005
   End
   Begin VB.Label Label13 
      BackStyle       =   0  'Transparent
      Caption         =   "Godown    :    "
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
      Left            =   8820
      TabIndex        =   102
      Top             =   1485
      Width           =   1230
   End
   Begin VB.Label Label8 
      BackStyle       =   0  'Transparent
      Caption         =   "Transport  :    "
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
      Left            =   8805
      TabIndex        =   101
      Top             =   1125
      Width           =   1140
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "Bill No     :    "
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
      Left            =   8805
      TabIndex        =   100
      Top             =   675
      Width           =   1185
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Vou. No :"
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
      Left            =   5535
      TabIndex        =   99
      Top             =   705
      Width           =   870
   End
   Begin VB.Label Label15 
      BackStyle       =   0  'Transparent
      Caption         =   "D&ue Days       :"
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
      Left            =   8805
      TabIndex        =   98
      Top             =   2295
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.Label Label31 
      BackStyle       =   0  'Transparent
      Caption         =   "Aft Tax Add   :"
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
      Left            =   13890
      TabIndex        =   63
      Top             =   7020
      Width           =   1335
   End
   Begin VB.Label Label24 
      BackStyle       =   0  'Transparent
      Caption         =   "Aft Tax Less  :"
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
      Left            =   13890
      TabIndex        =   64
      Top             =   7410
      Width           =   1470
   End
   Begin VB.Shape Shape5 
      BorderColor     =   &H00EFB17A&
      BorderWidth     =   5
      FillColor       =   &H00EFB17A&
      Height          =   9195
      Left            =   0
      Top             =   0
      Width           =   13770
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
      TabIndex        =   51
      Top             =   1140
      Width           =   900
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
      ForeColor       =   &H00800000&
      Height          =   375
      Left            =   495
      TabIndex        =   53
      Top             =   2205
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
      TabIndex        =   52
      Top             =   1800
      Visible         =   0   'False
      Width           =   870
   End
   Begin VB.Label Label9 
      BackStyle       =   0  'Transparent
      Caption         =   "Pur. A/C  :"
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
      Left            =   360
      TabIndex        =   0
      Top             =   4605
      Width           =   1620
   End
   Begin VB.Label LblExmptKg 
      BackStyle       =   0  'Transparent
      Caption         =   "Exem/Kg.  :"
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
      Left            =   14355
      TabIndex        =   68
      Top             =   8370
      Visible         =   0   'False
      Width           =   1155
   End
   Begin VB.Label lblSgstAmt 
      BackStyle       =   0  'Transparent
      Caption         =   "S GST Amt   : "
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
      Left            =   8130
      TabIndex        =   62
      Top             =   6390
      Width           =   1335
   End
   Begin VB.Label Label10 
      BackStyle       =   0  'Transparent
      Caption         =   "Sub Amt      :"
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
      Left            =   8130
      TabIndex        =   55
      Top             =   4680
      Width           =   1425
   End
   Begin VB.Label Label11 
      BackStyle       =   0  'Transparent
      Caption         =   "Add              :"
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
      Left            =   8130
      TabIndex        =   57
      Top             =   5085
      Width           =   1380
   End
   Begin VB.Label Label12 
      BackStyle       =   0  'Transparent
      Caption         =   "Less             :"
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
      Left            =   8130
      TabIndex        =   58
      Top             =   5520
      Width           =   1380
   End
   Begin VB.Label Label30 
      BackStyle       =   0  'Transparent
      Caption         =   "Amount        :"
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
      Left            =   6780
      TabIndex        =   66
      Top             =   8550
      Width           =   1410
   End
   Begin VB.Label Label17 
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
      Height          =   300
      Left            =   7605
      TabIndex        =   60
      Top             =   5985
      Width           =   1755
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
      Left            =   10845
      TabIndex        =   92
      Top             =   0
      Width           =   2475
   End
   Begin VB.Label Label28 
      BackStyle       =   0  'Transparent
      Caption         =   "ROff              :"
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
      Left            =   8130
      TabIndex        =   65
      Top             =   8130
      Width           =   1380
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
      Left            =   14130
      TabIndex        =   59
      Top             =   7965
      Visible         =   0   'False
      Width           =   1395
   End
   Begin VB.Label Label6 
      BackColor       =   &H00EFB17A&
      Caption         =   "     Trade  Purchase"
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
      TabIndex        =   90
      Top             =   0
      Width           =   13815
   End
   Begin VB.Shape Shape1 
      FillColor       =   &H00FED8FC&
      FillStyle       =   0  'Solid
      Height          =   4470
      Left            =   12195
      Shape           =   4  'Rounded Rectangle
      Top             =   4575
      Width           =   1350
   End
   Begin VB.Label lbl1 
      BackStyle       =   0  'Transparent
      Caption         =   "%"
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
      Index           =   1
      Left            =   7815
      TabIndex        =   61
      Top             =   6390
      Width           =   255
   End
   Begin VB.Shape Shape3 
      FillColor       =   &H00EDFDFE&
      FillStyle       =   0  'Solid
      Height          =   2220
      Left            =   315
      Top             =   495
      Width           =   8205
   End
   Begin VB.Shape Shape4 
      FillColor       =   &H00EDFDFE&
      FillStyle       =   0  'Solid
      Height          =   2220
      Left            =   8595
      Top             =   495
      Width           =   4965
   End
   Begin VB.Shape Shape2 
      FillColor       =   &H00EDFDFE&
      FillStyle       =   0  'Solid
      Height          =   4530
      Left            =   5520
      Top             =   4620
      Width           =   6405
   End
End
Attribute VB_Name = "frmPurchaseTradeGST"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Const tblName As String = "tblPurch"
Const IdField As String = "Vno"
Const OrderField As String = "Vdt,Vno"
Const MaxNo As Long = 72
Const OtherTxtFromNo As Long = 0 '-- Other Txt Ctrl From No
Const OtherTxtToNo As Long = 0 '-- Other Txt Ctrl Less No
Dim rstbl As Recordset '-- TblPurch
Dim rstblSub As Recordset '-- TblPurchSub
Dim varBookMark As Variant
Dim FirstTimeFlag As Boolean, Criteria As String
Dim OldgFrmTypeStr As String
Dim FormAction As Integer
Dim frmTypeStr As String
Dim mLastEntryDate As Date
Dim SalesAccode As Long
Dim ItemType As Integer
Dim i As Long
Dim X As String
Dim OldBag As Integer
Dim Mon As Integer

'-- Cartoon Bag Detail Entry
Private Sub CmdBagDetails_Click()
gfrmBagInwVNo = txt(0)
gfrmBagInwVType = frmTypeStr
frmCartoonEntry.Show
End Sub
'Dim GridFlag As Boolean
Private Sub cmdBtn_Click(Index As Integer)
'On Error GoTo ErrorRoutine
'--for Log
Dim j As Long
Dim tmpRst As Recordset
Dim wLogNo As Integer
Dim AMD As String

Select Case Index
Case 0 '---- Previous
     If GProcRstOpen(rstbl, "Select * from " & tblName & " where " & Criteria & "  and " & IdField & " = (Select max(" & IdField & ") from " & tblName & " where " & Criteria & " And " & IdField & "  < " & Val(txt(0)) & ")", "R") > 0 Then ReadFields Else MsgBox "This is First Record:", vbOKOnly + vbInformation
Case 1 '---- Next
      If GProcRstOpen(rstbl, "Select * from " & tblName & " where " & Criteria & "  and   " & IdField & " = (Select min(" & IdField & ") from " & tblName & " where " & IdField & " > '" & txt(0) & "' and " & Criteria & ") and " & Criteria, "R") > 0 Then ReadFields Else MsgBox "This is last Record:", vbOKOnly + vbInformation
Case 2 '---- First
         If GProcRstOpen(rstbl, "Select * from " & tblName & " where " & Criteria & "  and  " & IdField & " = (Select min(" & IdField & ") from " & tblName & " where " & IdField & " <> '' and " & Criteria & ")and " & Criteria, "R") > 0 Then ReadFields
Case 3 '---- Last
'        If GProcRstOpen(rstbl, "Select * from " & tblName & " where " & Criteria & "  and  " & IdField & " = (Select max(" & IdField & ") from " & tblName & " where " & IdField & " <> '' and " & Criteria & ")and " & Criteria, "R") > 0 Then ReadFields Else GProcClearForm Me, rstbl, MaxNo
        If GProcRstOpen(rstbl, "Select * from " & tblName & " where " & Criteria & "  and  " & IdField & " = (Select max(" & IdField & ") from " & tblName & " where vdt = (SELECT MAX(vdt) from " & tblName & " where " & Criteria & " )  and  " & IdField & " <> 0 and  " & Criteria & ")", "R") > 0 Then ReadFields Else GProcClearForm Me, rstbl, MaxNo

Case 4 'Find ' QryMastAccount
     gClsSearch.SearchMultiField "tblPurch,TblMastAccount", "Vno,AcName,Vdt,PurBillNo,PurBillDt", Array("Vno", "Party", "Vdt", "Bill No", "Bill Dt"), Array(txt(0).Width, 5000, 1500, 3000, 1500), " tblPurch.PurAcCrCode=TblMastAccount.AcCode and vtype='" & frmTypeStr & "' and VFirm = '" & gCCode & "'  and VYear=" & gCYear, "", "Vno,Vdt", 0, 0
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
    If Year(mLastEntryDate) = 1899 Then
        dtpDt(2) = Format(Now, "dd/MM/yyyy")
    Else
        dtpDt(2) = mLastEntryDate
    End If
    If FirstTimeFlag = True Then
        FirstTimeFlag = False
    Else
        dtpDt(2).SetFocus
    End If
    lblPurBgas.Caption = ""
    lblPurWeight.Caption = ""
'    txt(0) = GProcGenerateId("TblPurch", "Vno", "Vtype='" & frmTypeStr & "' and VFirm = '" & gCCode & "'   and VYear=" & gCYear)
'    dtpDt(2) = GProcGenerateVDt("tblPurch", "Vdt", "Vtype='" & frmTypeStr & "' and VFirm = '" & gCCode & "'   and VYear=" & gCYear)
'    dtpDt(2) = Format(Now, "dd/MM/yyyy")
    
    '--A/C
    If frmTypeStr = "PE" Then
       txttemp(6) = GProcGetColumnValue("tblMastAccount", "AcCode", CStr(gAcCodePYExempt), "N", "AcName", "S")
    Else
       txttemp(6) = GProcGetColumnValue("tblMastAccount", "AcCode", CStr(gAcCodePY), "N", "AcName", "S")
    End If
    SetGrid
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
        End If
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
              X = "Select max(LogNo) from tblPurch_Log where Vno=" & Val(txt(0)) & " and " & Criteria & ""
              j = GProcRstOpen(tmpRst, X, "R")
              If j > 0 Then
                 wLogNo = IIf(IsNull(tmpRst.Fields(0)), 0, tmpRst.Fields(0)) + 1
              Else
                 wLogNo = 1
              End If
              gCn.Execute "Insert into tblPurch_Log select tblPurch.*," & wLogNo & " as LogNo,'D' as LogTp ,'" & gUserName & "' as UserName ,'" & Format(Now, "YYYY/MM/DD hh:mm:ss") & "' as LogDate from tblPurch where Vno=" & Val(txt(0)) & " and " & Criteria & ""
              gCn.Execute "Insert into tblPurchSub_Log select tblPurchSub.*," & wLogNo & " as LogNo,'D' as LogTp ,'" & gUserName & "' as UserName ,'" & Format(Now, "YYYY/MM/DD hh:mm:ss") & "' as LogDate from tblPurchSub where Vno=" & Val(txt(0)) & " and " & Criteria & ""
              gCn.CommitTrans
              tmpRst.Close
           End If
           DeleteRelatedRecord
           gCn.Execute "delete from tblPurchSub where vtype='" & frmTypeStr & "' and vno=" & txt(0) & " and Vfirm = '" & gCCode & "'  and " _
                       & " VYear=" & gCYear
           GProcRstOpen rstbl, "Select * from " & tblName & " where " & IdField & " = " & txt(0) & " and " & Criteria, "O"
           rstbl.Delete
           cmdBtn_Click (3)
        End If
    End If
Case 9 'Print
    Frame1.Caption = "Commission Details Printing"
    '-- Inv No
    txtPInvno(0) = txt(0).text
    Frame1.Visible = True
    DoEvents
    txtPInvno(0).Enabled = True
    txtPInvno(0).SetFocus
Case 10 'Exit
'    DoEvents
    Unload Me
    Exit Sub
Case 11 'Label
    'frmLabel.Show 1
Case 12 'Optional
Case 13 'Save
        'txt(51) = CDate(Date)
        If Trim(txtGrid) <> "" Then
            msGrid.text = txtGrid
            DoEvents
            txtGrid = ""
        End If
        txtGrid.Visible = False
       ' If FormAction = vbDataActionAddNew Then txt(0) = GProcGenerateId(tblName, IdField, "vtype='" & FrmTypeStr & "' and VYear=" & gCYear)
        CalAmount
        FillTxtFromTemp
        If gfrmHunkStr = "HNK" Then
           txt(36) = 1
        Else
           txt(36) = 0
        End If
        If gfrmHunkStr = "EXM" Then
           txt(40) = 1
        Else
           txt(40) = 0
        End If
        '-- Save Record
        If FormAction = vbDataActionAddNew Then
           Mon = Month(dtpDt(2))
           txt(0) = GProcGenerateIdMonthwise("TblPurch", "Vno", Mon, "Vdt", "Vtype='" & frmTypeStr & "' and VFirm = '" & gCCode & "'   and VYear=" & gCYear)
        End If
        If ValidateData = True Then
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
               X = "Select max(LogNo) from tblPurch_Log where Vno=" & Val(txt(0)) & " and " & Criteria & ""
               j = GProcRstOpen(tmpRst, X, "R")
               If j > 0 Then
                  wLogNo = IIf(IsNull(tmpRst.Fields(0)), 0, tmpRst.Fields(0)) + 1
               Else
                  wLogNo = 1
               End If
               gCn.Execute "Insert into tblPurch_Log select tblPurch.*," & wLogNo & " as LogNo,'" & AMD & "' as LogTp ,'" & gUserName & "' as UserName ,'" & Format(Now, "YYYY/MM/DD hh:mm:ss") & "' as LogDate from tblPurch where Vno=" & Val(txt(0)) & " and " & Criteria & ""
               gCn.Execute "Insert into tblPurchSub_Log select tblPurchSub.*," & wLogNo & " as LogNo,'" & AMD & "' as LogTp ,'" & gUserName & "' as UserName ,'" & Format(Now, "YYYY/MM/DD hh:mm:ss") & "' as LogDate from tblPurchSub where Vno=" & Val(txt(0)) & " and " & Criteria & ""
               tmpRst.Close
            End If
            mLastEntryDate = dtpDt(2)
            FormAction = vbDataActionClose
            Call GProcActivateControls(Me, False)
            Call GProcSetButtons(Me, True)
            msGrid.Col = 0 '-----
            msGrid.Row = 1 '-----
            txtGrid.Move msGrid.CellLeft + msGrid.Left, msGrid.CellTop + msGrid.Top, msGrid.CellWidth, msGrid.CellHeight '-----
'            If FormAction = vbDataActionAddNew Then
                If cmdBtn(6).Visible = True Then cmdBtn(6).SetFocus
 '           Else
  '              If cmdBtn(7).Visible = True Then cmdBtn(7).SetFocus
   '         End If
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
'    SetGrid
    msGrid.Col = 0
    msGrid.Row = 1
    txtGrid.Move msGrid.CellLeft + msGrid.Left, msGrid.CellTop + msGrid.Top, msGrid.CellWidth, msGrid.CellHeight
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
Private Sub SetControlEd()
txttemp(6).Enabled = False '----Purch A/c
txt(16).Enabled = False '-- Amt
txt(21).Enabled = False '-- Ass Val
dtpDt(32).Enabled = False '--Due Dt
txtCrBalance.Enabled = False
If FormAction = vbDataActionAddNew Then
   txt(0).Enabled = False
End If
If Trim(txtCode(5)) <> "" Then
   txttemp(5).Enabled = False
Else
   txttemp(5).Enabled = True
End If
If Trim(txtCode(7)) <> "" Then
   txttemp(7).Enabled = False
Else
   txttemp(7).Enabled = True
End If
If Trim(txtCode(8)) <> "" Then
   txttemp(8).Enabled = False
Else
   txttemp(8).Enabled = True
End If
txtCode(65).Enabled = False
txttemp(65).Enabled = False
If gCYear >= 2023 Then
   txt(72).Visible = True
   txt(72).Enabled = False
   Label19.Visible = True
Else
   txt(72).Visible = False
   Label19.Visible = False
End If
End Sub
Private Sub DeleteRelatedRecord()
''--- Voucher Sub
gCn.Execute "delete from tblPurchSub where vtype='" & frmTypeStr & "' and vno=" & txt(0) & " and Vfirm = '" & gCCode & "'  and " _
& " VYear=" & gCYear
End Sub
Private Sub SaveRelatedRecord()
Dim i As Long
Dim MillCode As Long
Dim rstSub As Recordset
Dim CrCode As Long
Dim DrCode As Long
Dim BrkCode As Long
Dim TaxAcCode As Long
Dim wit As Long, wgd As Long, Y As String, maxy As Long
Dim BillDt
Dim DueDt
Dim Exempt As Integer
Dim JVTDSVNo As Long
Dim rstTmp As Recordset
'--- Account Code
MillCode = GProcGetColumnValue("tblMastAccount", "AcName", txttemp(7), "S", "AcCode", "N")
CrCode = GProcGetColumnValue("tblMastAccount", "AcName", txttemp(5), "S", "AcCode", "N")
DrCode = GProcGetColumnValue("tblMastAccount", "AcName", txttemp(6), "S", "AcCode", "N")
BrkCode = GProcGetColumnValue("tblMastAccount", "AcName", txttemp(8), "S", "AcCode", "N")
'TaxAcCode = GProcGetColumnValue("tblMastNarration", "Narration", Cbo(22), "S", "MastTaxPurAcCode", "N")
'-- Delete Records
DeleteRelatedRecord
'--- Save Grid Records
'--- TblVouSub
X = "select * from tblPurchSub where vtype='" & frmTypeStr & "' and vno=" & txt(0) & " and Vfirm = '" & gCCode & "' and " _
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
        rstSub!pursubitsrno = i '--Sr No
        '--- It Ctrl No(11)
        rstSub!PursubBookNo = Val(.TextMatrix(i, 1))
        rstSub!PurSubBkVyear = Val(.TextMatrix(i, 2))
        rstSub!PurSubItCode = Val(GProcGetColumnValue("tblMastItem", "ItName", .TextMatrix(i, 3), "S", "ItCode", "N", " ItMillCode=" & MillCode))
        rstSub!PurSubBag = Val(.TextMatrix(i, 6))
        rstSub!PurSubwt = Val(.TextMatrix(i, 7))
        rstSub!PurSubRt = Val(.TextMatrix(i, 8))
        rstSub!PurSubRtPer = Val(.TextMatrix(i, 9))
        rstSub!PurSubAmt = Val(.TextMatrix(i, 10))
        rstSub!PurSubLotNo = .TextMatrix(i, 5) & vbNullString
        If txtGodown.text = "" Then
           .TextMatrix(i, 4) = "-"
        Else
           .TextMatrix(i, 4) = txtGodown.text
        End If
        rstSub!PurSubGodown = Val(GProcGetColumnValue("tblMastNarration", "Narration", .TextMatrix(i, 4), "S", "NarrCode", "N", ""))
        rstSub!PurSubBkSrNo = Val(.TextMatrix(i, 13))
        rstSub.Update
        i = i + 1
     Loop
 End With
 rstSub.Close

'maxy = GProcGetColumnValue("tblPurchSub", "Vno", txt(0), "N", "max(PurSubItCtrlNo)", "N", Criteria)
'With msGrid
'    i = 1
'    gCn.Execute "update TblPurchSub set pursubitsrno = pursubitsrno+100 where " & Criteria & " and vno=" & txt(0)
'    Y = ""
'    Do While i < .rows
'       wit = Val(GProcGetColumnValue("tblMastItem", "ItName", .TextMatrix(i, 3), "S", "ItCode", "N", " ItMillCode=" & MillCode))
'       .TextMatrix(i, 4) = "-"
'       wgd = Val(GProcGetColumnValue("tblMastNarration", "Narration", .TextMatrix(i, 4), "S", "NarrCode", "N", ""))
'       If Val(.TextMatrix(i, 11)) = 0 Then
'           X = Val(txt(0)) & ",'" & txt(1) & "','" & Format(txt(2), "yyyy/mm/dd") & "'," & Val(txt(3)) & ",'" & gCCode & "'," & i & "," & Val(.TextMatrix(i, 1)) & ",'" & Val(.TextMatrix(i, 2)) & "',"
'           X = X & wit & ","
'           X = X & Val(.TextMatrix(i, 6)) & "," & Val(.TextMatrix(i, 7)) & "," & Val(.TextMatrix(i, 8)) & "," & Val(.TextMatrix(i, 9)) & "," & Val(.TextMatrix(i, 10)) & ",'" & .TextMatrix(i, 5) & "',"
'           X = X & wgd & "," & Val(.TextMatrix(i, 13)) & ""
'          gCn.Execute "Insert into TblPurchSub(vno,vtype,vdt,vyear,vfirm,pursubitsrno,PursubBookNo,PurSubBkVyear,PurSubItCode,PurSubBag,PurSubwt,PurSubRt,PurSubRtPer,PurSubAmt,PurSubLotNo,PurSubGodown,PurSubBkSrNo)" _
'            & " Values(" & X & ")"
'        Else
'          Y = Y & Val(.TextMatrix(i, 11)) & ","
'
'          gCn.Execute "Update TblPurchSub set pursubitsrno=" & i & ", vdt ='" & Format(txt(2), "yyyy/mm/dd") & "',PursubBookNo=" & Val(.TextMatrix(i, 1)) & ",PurSubBkVyear=" & Val(.TextMatrix(i, 2)) & ",PurSubBkSrNo=" & Val(.TextMatrix(i, 13)) & "   ," _
'                & "PurSubItCode=" & wit & ",PurSubBag=" & Val(.TextMatrix(i, 6)) & ",PurSubwt=" & Val(.TextMatrix(i, 7)) & ",PurSubRt=" & Val(.TextMatrix(i, 8)) & ",PurSubRtPer=" & Val(.TextMatrix(i, 9)) & ",PurSubAmt=" & Val(.TextMatrix(i, 10)) & ",PurSubLotNo='" & Trim(.TextMatrix(i, 5)) & "',PurSubGodown=" & wgd & " where PurSubItCtrlNo=" & Val(.TextMatrix(i, 11))
'        End If
'        i = i + 1
'     Loop
' End With
'''' rstSub.Close
'If Len(Y) >= 2 Then
' Y = Left(Y, Len(Y) - 1)
' gCn.Execute "delete from TblPurchSub where PurSubItCtrlNo <= " & maxy & " and PurSubItCtrlNo not in (" & Y & ") and " & Criteria & " and vno=" & txt(0)
'End If

'-- For TDS
'If FormAction = vbDataActionAddNew And Val(txt(64)) > 0 Then
'   Mon = Month(dtpDt(10))
'   JVTDSVNo = GProcGenerateIdMonthwise("TblVoucher", "Vno", Mon, "Vdt", "Vtype='JV' and VFirm = '" & gCCode & "'   and VYear=" & gCYear)
'Else
'   If Val(txt(64)) > 0 Then
'      If Val(txt(62)) > 0 And GProcGetColumnValue("tblVoucher", "Vno", txt(62), "N", "Vno", "N", " Vtype='JV' and  VFirm = '" & gCCode & "' and Vyear=" & gCYear) = 0 Then
'         JVTDSVNo = CLng(txt(62))
'      Else
'         Mon = Month(dtpDt(10))
'         JVTDSVNo = GProcGenerateIdMonthwise("TblVoucher", "Vno", Mon, "Vdt", "Vtype='JV' and VFirm = '" & gCCode & "'   and VYear=" & gCYear)
'      End If
'   End If
'End If
'If Val(txt(64)) > 0 Then
'   X = "select * from tblVoucher where vtype= 'JV' and vno=" & JVTDSVNo & " and Vfirm = '" & gCCode & "' and " _
'   & " VYear=" & gCYear
'   GProcRstOpen rstTmp, X, "O"
'   rstTmp.AddNew
'   rstTmp!Vno = JVTDSVNo  '--VNo
'   rstTmp!Vtype = "JV"  '--Vtype
'   rstTmp!Vdt = CDate(txt(10)) '--Vdt(2)
'   rstTmp!VYear = txt(3) '--VYear
'   rstTmp!VFirm = gCCode '--VFirm
'   rstTmp!VCtrNo = 1 '--Sr No
'   rstTmp!VDrAcCode = CrCode
'   rstTmp!VCrAcCode = txt(65) 'gTDSAcCodePay
'   rstTmp!vamt = Val(txt(64))
'   rstTmp!VNar1 = "T.D.S. On Rs. " + txt(16) + " @ " + txt(63) + " %"
'   rstTmp!VNar2 = "Against Bill No " + txt(9) + " Dtd. " + txt(10)
'   rstTmp!VIsRefEntType = 1
'   rstTmp!VTopCrDr = 1
'   rstTmp.Update
'   gCn.Execute "update tblPurch set PurTdsJvNo = " & JVTDSVNo & " where vtype='" & frmTypeStr & "' and vno=" & txt(0) & " and  VFirm = '" & gCCode & "'  and    " _
'               & " VYear=" & gCYear
'End If
'If Val(txt(64)) = 0 Then
'   gCn.Execute "update tblPurch set PurTdsJvNo = 0 where vtype='" & frmTypeStr & "' and vno=" & txt(0) & " and  VFirm = '" & gCCode & "'  and    " _
'               & " VYear=" & gCYear
'End If

'-------------------------------


End Sub

Private Sub cmdPrintBill_Click(Index As Integer)
Dim RsSA As Recordset
Dim RsTmp As Recordset
Dim X As String
Dim RecYN As Long
Dim Y As String

Select Case Index
Case 0, 1:
    '------ Validate
    If IsNumeric(txtPInvno(0)) = False Then txtPInvno(0) = 0
    '-------
    If Index = 0 Then GProcCrystalRptPreparation cryReport, "Printer"
    If Index = 1 Then GProcCrystalRptPreparation cryReport, "Window"
    gCn.Execute "Delete from TmpPackingSlip where  UserName='" & gUserName & "'"
    
    Y = "Select * from TmpPackingSlip"
    GProcRstOpen RsTmp, Y, "O"
    
    X = "Select a.Vno,a.Vdt,b.AcName,b.AcAdd1,b.AcAdd2,b.AcAdd3,b.AcCity,b.AcZip,b.AcOurCode,b.AcGSTIN,c.Narration,a.PurBillNo,a.PurBillDt,d.ItName,e.SrNo,e.BagNo,e.CartoonWt " _
        & " From tblPurch A,tblMastAccount B,tblMastNarration C, tblMastItem D, tblBags E  where a.vno = e.Invno and a.vtype = e.Invtype and a.vyear = e.Invyear and a.vfirm = e.Invfirm and " _
        & " a.vno = " & txtPInvno(0) & " and a.Vtype='" & frmTypeStr & "' and a.VFirm = '" & gCCode & "' And a.VYear = " & gCYear & " and a.PurAcCrCode = b.AcCode and e.InGodown = c.NarrCode and e.InItCode = d.ItCode " _
        & " order by e.SrNo "

    RecYN = GProcRstOpen(RsSA, X, "R", gCn)
                 
    If RecYN > 0 Then
       With RsSA
            .MoveFirst
            Do While Not .EOF
               If RsSA.Fields(14) <= 30 Then
                  RsTmp.AddNew
                  RsTmp!Vno = RsSA.Fields(0)
                  RsTmp!Vtype = frmTypeStr
                  RsTmp!VFirm = gCCode
                  RsTmp!VYear = gCYear
                  RsTmp!Vdt = RsSA.Fields(1)
                  RsTmp!MillName = RsSA.Fields(2)
                  RsTmp!MillAdd1 = RsSA.Fields(3) & vbNullString
                  RsTmp!MillAdd2 = RsSA.Fields(4) & vbNullString
                  RsTmp!MillAdd3 = RsSA.Fields(5) & vbNullString
                  RsTmp!MillCity = RsSA.Fields(6) & vbNullString
                  RsTmp!MillPin = RsSA.Fields(7) & vbNullString
                  RsTmp!MillCode = RsSA.Fields(8) & vbNullString
                  RsTmp!MillGstin = RsSA.Fields(9) & vbNullString
                  RsTmp!GodownName = RsSA.Fields(10) & vbNullString
                  RsTmp!VBillNo = RsSA.Fields(11) & vbNullString
                  If IsDate(RsSA.Fields(12)) Then
                     RsTmp!VBILLDT = RsSA.Fields(12)
                  End If
                  RsTmp!CountName = RsSA.Fields(13) & vbNullString
                  RsTmp!SrNo1 = RsSA.Fields(14)
                  RsTmp!BagNo1 = RsSA.Fields(15) & vbNullString
                  RsTmp!Weight1 = RsSA.Fields(16)
                  RsTmp!UserName = gUserName
                  RsTmp.Update
               ElseIf RsSA.Fields(14) > 30 And RsSA.Fields(14) <= 60 Then
                  gCn.Execute "update TmpPackingSlip set SrNo2 = " & RsSA.Fields(14) & ", BagNo2 = '" & RsSA.Fields(15) & "',Weight2 = " & RsSA.Fields(16) & "  where UserName='" & gUserName & "' and vno=" & RsSA.Fields(0) & " and  VFirm = '" & gCCode & "'  and    " _
                  & " SrNo1 = " & RsSA.Fields(14) - 30 & " and VYear=" & gCYear
               ElseIf RsSA.Fields(14) > 60 And RsSA.Fields(14) <= 90 Then
                  gCn.Execute "update TmpPackingSlip set SrNo3 = " & RsSA.Fields(14) & ", BagNo3 = '" & RsSA.Fields(15) & "',Weight3 = " & RsSA.Fields(16) & "  where UserName='" & gUserName & "' and vno=" & RsSA.Fields(0) & " and  VFirm = '" & gCCode & "'  and    " _
                  & " SrNo1 = " & RsSA.Fields(14) - 60 & " and VYear=" & gCYear
               ElseIf RsSA.Fields(14) > 90 And RsSA.Fields(14) <= 120 Then
                  gCn.Execute "update TmpPackingSlip set SrNo4 = " & RsSA.Fields(14) & ", BagNo4 = '" & RsSA.Fields(15) & "',Weight4 = " & RsSA.Fields(16) & "  where UserName='" & gUserName & "' and vno=" & RsSA.Fields(0) & " and  VFirm = '" & gCCode & "'  and    " _
                  & " SrNo1 = " & RsSA.Fields(14) - 90 & " and VYear=" & gCYear
'--Second Page
               ElseIf RsSA.Fields(14) > 120 And RsSA.Fields(14) <= 150 Then
                  RsTmp.AddNew
                  RsTmp!Vno = RsSA.Fields(0)
                  RsTmp!Vtype = frmTypeStr
                  RsTmp!VFirm = gCCode
                  RsTmp!VYear = gCYear
                  RsTmp!Vdt = RsSA.Fields(1)
                  RsTmp!MillName = RsSA.Fields(2)
                  RsTmp!MillAdd1 = RsSA.Fields(3) & vbNullString
                  RsTmp!MillAdd2 = RsSA.Fields(4) & vbNullString
                  RsTmp!MillAdd3 = RsSA.Fields(5) & vbNullString
                  RsTmp!MillCity = RsSA.Fields(6) & vbNullString
                  RsTmp!MillPin = RsSA.Fields(7) & vbNullString
                  RsTmp!MillCode = RsSA.Fields(8) & vbNullString
                  RsTmp!MillGstin = RsSA.Fields(9) & vbNullString
                  RsTmp!GodownName = RsSA.Fields(10) & vbNullString
                  RsTmp!VBillNo = RsSA.Fields(11) & vbNullString
                  If IsDate(RsSA.Fields(12)) Then
                     RsTmp!VBILLDT = RsSA.Fields(12)
                  End If
                  RsTmp!CountName = RsSA.Fields(13) & vbNullString
                  RsTmp!SrNo1 = RsSA.Fields(14)
                  RsTmp!BagNo1 = RsSA.Fields(15) & vbNullString
                  RsTmp!Weight1 = RsSA.Fields(16)
                  RsTmp!UserName = gUserName
                  RsTmp.Update
               ElseIf RsSA.Fields(14) > 150 And RsSA.Fields(14) <= 180 Then
                  gCn.Execute "update TmpPackingSlip set SrNo2 = " & RsSA.Fields(14) & ", BagNo2 = '" & RsSA.Fields(15) & "',Weight2 = " & RsSA.Fields(16) & "  where UserName='" & gUserName & "' and vno=" & RsSA.Fields(0) & " and  VFirm = '" & gCCode & "'  and    " _
                  & " SrNo1 = " & RsSA.Fields(14) - 30 & " and VYear=" & gCYear
               ElseIf RsSA.Fields(14) > 180 And RsSA.Fields(14) <= 210 Then
                  gCn.Execute "update TmpPackingSlip set SrNo3 = " & RsSA.Fields(14) & ", BagNo3 = '" & RsSA.Fields(15) & "',Weight3 = " & RsSA.Fields(16) & "  where UserName='" & gUserName & "' and vno=" & RsSA.Fields(0) & " and  VFirm = '" & gCCode & "'  and    " _
                  & " SrNo1 = " & RsSA.Fields(14) - 60 & " and VYear=" & gCYear
               ElseIf RsSA.Fields(14) > 210 And RsSA.Fields(14) <= 240 Then
                  gCn.Execute "update TmpPackingSlip set SrNo4 = " & RsSA.Fields(14) & ", BagNo4 = '" & RsSA.Fields(15) & "',Weight4 = " & RsSA.Fields(16) & "  where UserName='" & gUserName & "' and vno=" & RsSA.Fields(0) & " and  VFirm = '" & gCCode & "'  and    " _
                  & " SrNo1 = " & RsSA.Fields(14) - 90 & " and VYear=" & gCYear
               End If
               .MoveNext
            Loop
       End With
    End If
'    X = "select * from TmpPackingSlip where UserName='" & gUserName & "'"
    With cryReport
         X = "{TmpPackingSlip.UserName}='" & gUserName & "' "
         .SelectionFormula = X
         .ReportFileName = gReportPath & "rptPackingSlip_Purch.rpt"
         .Action = 1
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
            If (cmdBtn(13).Enabled = False) Then
                Unload Me
            End If
    End Select
End Sub
Private Sub Form_Load()
    Me.Top = 300
    Me.Left = Screen.Width / 2 - Me.Width / 2
    '--- TblAddLess
    frmTypeStr = gfrmTypeStr
    mLastEntryDate = vbNull
    X = "Select * from " & tblName & " where Vtype='" & frmTypeStr & "' and VFirm  = '" & gCCode & "' and VYear=" & gCYear
    Criteria = "vtype='" & frmTypeStr & "' and Vfirm = '" & gCCode & "' and  VYear=" & gCYear
    If Trim(OrderField) <> "" Then X = X & " order by " & OrderField
'    GProcRstOpen rstbl, X, "O"
    FirstTimeFlag = True
    If frmTypeStr = "PE" Then
       txttemp(6) = GProcGetColumnValue("tblMastAccount", "AcCode", CStr(gAcCodePYExempt), "N", "AcName", "S")
'       txttemp(35) = GProcGetColumnValue("tblMastTransactionCode", "TranCode", CStr(gTranCdPYExempt), "S", "TranCode", "S")
    Else
       txttemp(6) = GProcGetColumnValue("tblMastAccount", "AcCode", CStr(gAcCodePY), "N", "AcName", "S")
'       txttemp(35) = GProcGetColumnValue("tblMastTransactionCode", "TranCode", CStr(gTranCdPY), "S", "TranCode", "S")
    End If
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
    '---Grid
    '---- Button Visiblity Set As Per User Previlage
    GProcSetButtonVisiblity Me
'    cmdBtn_Click (6)
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
Set rstblSub = New Recordset
X = "select PurSubItSrNo,PurSubBookNo,PurSubBkVyear,ItName,Narration,PurSubLotNo,PurSubBag,PurSubWt" _
& " ,PurSubRt,PurSubRtPer,PurSubAmt,PurSubItCtrlNo,PurSubPItCtrlNo,PurSubBkSrNo from tblPurchSub,TblMastNarration,TblMastItem where " _
& " tblPurchSub.PurSubItCode=TblMastItem.ItCode and tblPurchSub.PurSubGodown=TblMastNarration.NarrCode " _
& " and tblPurchSub.vtype='" & frmTypeStr & "' and tblPurchSub.vno=" & txt(0) & " and tblPurchSub.VFirm = '" & gCCode & "' and tblPurchSub.VYear=" & gCYear & " order by PurSubItSrNo"
GProcRstOpen rstblSub, X, "R"
With msGrid
   .FormatString = ">Sr|>Bk.No    |>Bk. Year  |<Count                                                           |<Godown               |<Lot No                 |>Bag      |>Weight         |>Rate                |>Per    |>Amount                    |||"
   .ColWidth(1) = 0
   .ColWidth(2) = 0
   .ColWidth(4) = 0
'   .ColWidth(5) = 0
    For i = 11 To 13
        .ColWidth(i) = 0
    Next i
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
SetControlEd
If Index = 2 Then
    If GProcIsDateValid(dtpDt(Index)) = False Then
        dtpDt(Index).SetFocus
        Exit Sub
    End If
    If dtpDt(10) = "01/04/1899" And IsDate(dtpDt(2)) = True Then dtpDt(10) = dtpDt(2)
End If
If dtpDt(10) = "01/04/1899" And IsDate(dtpDt(2)) = True Then dtpDt(10) = dtpDt(2)

If Index = 10 And (dtpDt(14) = "01/04/1899" Or dtpDt(14) <> dtpDt(10)) And IsDate(dtpDt(10)) = True Then
   dtpDt(14) = dtpDt(10)
End If
'If Index = 2 Then ShowAccountBalance
'If IsDate(mskDt(Index)) = False Then mskDt(Index) = "__/__/____"
If Index = 14 Then
   msGrid.Col = 3
   msGrid_EnterCell
End If
End Sub
Private Sub ShowAccountBalance()
'Dim DrAcBal  As Double
Dim CrAcBal As Double
If IsDate(dtpDt(2)) Then
'    DrAcBal = GProcGetAccountBalance(txttemp(4), dtpDt(2))
    CrAcBal = GProcGetAccountBalance(txttemp(5), dtpDt(2))
'    txtDrBalance = IIf(DrAcBal > 0, DrAcBal & " Cr", Abs(DrAcBal) & " Dr")
    txtCrBalance = IIf(CrAcBal > 0, CrAcBal & " Cr", Abs(CrAcBal) & " Dr")
End If
End Sub
Private Sub txt_GotFocus(Index As Integer)
Dim Amt As Double
Dim Amt2 As Double
GProcSelectBox Me.ActiveControl
If Index = 30 And Val(txt(30)) = 0 Then
   Amt = Val(txt(21)) + Val(txt(24)) + Val(txt(27)) + Val(txt(38)) + Val(txt(70))
   Amt2 = GProcMakeRounding(Amt)
   txt(30) = Format(Amt2 - Format(Amt, FStr), FStr)
End If
End Sub
Private Sub CheckForNumber()
GProcCheckForNumber Me, rstbl, MaxNo, pOtherTxtFromNo:=OtherTxtFromNo, pOtherTxtToNo:=OtherTxtToNo
End Sub
Private Sub txt_LostFocus(Index As Integer)
If Index = 9 And Trim(txt(9)) <> "" And FormAction <> vbDataActionCancel Then
   Dim rstAddCheck As Recordset
   i = GProcGetColumnValue("tblMastAccount", "AcName", txttemp(5), "S", "AcCode", "N")
   X = "Select vno From " & tblName & " Where PurBillNo= '" & txt(9) & "' and PurAcCrCode = " & i & "  and  vno <> " & CLng(txt(0)) & "  and Vtype='" & frmTypeStr & "'  and Vfirm = '" & gCCode & "'  and VYear=" & gCYear
   i = GProcRstOpen(rstAddCheck, X, "R")
   If i > 0 Then
       MsgBox "Same Bill No of Same Mill is Available in Voucher No. " & rstAddCheck.Fields(0), vbCritical + vbOKOnly, "Duplicate Record"
       dtpDt(10).SetFocus
   End If
End If
CheckForNumber
SaveBtnEd
CalAmount
If Index = 31 Then
   msGrid.Col = 3
   msGrid_EnterCell
End If
End Sub
Private Sub txt_KeyPress(Index As Integer, KeyAscii As Integer)
If KeyAscii = 13 Then Exit Sub
KeyAscii = GProcValidateKey(rstbl, Index, KeyAscii, txt(Index).text)
End Sub

Private Sub txtCode_LostFocus(Index As Integer)
If Trim(txtCode(Index)) = "" Then Exit Sub
Dim BrkAcCode As Long
Dim AgCode As Long
Dim i As Integer
SetControlEd
Select Case Index
    Case 5  '-- Party  Code
         txttemp(5) = GProcGetColumnValue("tblMastAccount", "AcOurCode", UCase(txtCode(5)), "S", "AcName", "S")
         lblTinNo.Caption = GProcGetColumnValue("tblMastAccount", "AcOurCode", UCase(txtCode(5)), "S", "AcGSTIN", "S")
         i = GProcGetColumnValue("tblMastAccount", "AcOurCode", UCase(txtCode(5)), "S", "AcMsOmsParty", "N")
         If i = 0 Then
            lblMsOms.Caption = "M.S."
         Else
            lblMsOms.Caption = "O.M.S."
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
         AgCode = GProcGetColumnValue("tblMastAccount", "AcName", txttemp(5), "S", "AgCode", "N")
        If AgCode = 90029 Then
            txttemp(7) = txttemp(5)
            txtCode(7) = txtCode(5)
            txttemp(7).Enabled = False
            txtCode(7).Enabled = False
        Else
            txttemp(7).Enabled = True
            txtCode(7).Enabled = True
        End If
    Case 7  '-- Mill  Code
          Dim MilTpCd  As Long
          Dim X As String
          Dim tmpRst As Recordset
          MilTpCd = GProcGetColumnValue("tblMastNarration", "Narration", "Trade", "S", "Narrcode", "N")
          X = "Select Acname from tblMastAccount,tblMastBillSerial where tblMastaccount.Accode = tblMastBillSerial.MillCode and  tblMastBillSerial.SaleType = " & MilTpCd & "  and tblMastBillSerial.Vfirm = '" & gCCode & "'   and  tblMastaccount.AcOurCode = '" & txtCode(7) & "'   "
          i = GProcRstOpen(tmpRst, X, "R")
          If i = 0 Then
              MsgBox ("Mill Not in Firm")
              txttemp(7).text = ""
              txtCode(7).SetFocus
              Exit Sub
          End If
         txttemp(7) = GProcGetColumnValue("tblMastAccount", "AcOurCode", UCase(txtCode(7)), "S", "AcName", "S")
         If Trim(txttemp(7)) = "" Then
            txttemp(7).Enabled = True
            txttemp(7).SetFocus
         End If
    Case 8  '-- Broker  Code
         txttemp(8) = GProcGetColumnValue("tblMastAccount", "AcOurCode", UCase(txtCode(8)), "S", "AcName", "S")
         If Trim(txttemp(8)) = "" Then
            txttemp(8).Enabled = True
            txttemp(8).SetFocus
         End If
End Select
DoEvents
End Sub
Private Sub txtGodown_KeyPress(KeyAscii As Integer)
Dim LeftPos As Long
Dim TopPos As Long
TopPos = txtGodown.Top + Me.Top + 650
LeftPos = txtGodown.Left + Me.Left
If KeyAscii = 13 Then Exit Sub
gClsSearch.SearchMultiField "tblMastNarration", "Narration", Array("Godown"), Array(6000), " NarrType='G' ", Chr(KeyAscii), "Narration", LeftPos, TopPos, True
KeyAscii = 0
If gClsSearch.SearchMultiRetCol(0) <> "" Then
   txtGodown.text = gClsSearch.SearchMultiRetCol(0)
End If
End Sub

Private Sub txttemp_GotFocus(Index As Integer)
GProcSelectBox Me.ActiveControl
End Sub
Private Sub txttemp_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
Dim GName As String
Dim OldFrmType As String
If KeyCode = vbKeyF3 Then
    Select Case Index
        Case 5 '-- Party
            GName = GProcGetColumnValue("TblMastGroup", "AgCode", 90029, "N", "AgName", "S")
            GProcShowForm frmMastAccount, frmMain.mnuMstfrm(0), Array(frmMastAccount.txttemp(2)), Array(GName)
        Case 6 '-- Purchase A/c
            GName = GProcGetColumnValue("TblMastGroup", "AgCode", 90025, "N", "AgName", "S")
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
TopPos = txttemp(Index).Top + Me.Top + 650
LeftPos = txttemp(Index).Left + Me.Left
If KeyAscii = 13 Then Exit Sub
Select Case Index
    Case 5  '-- Party   Alias Account
        If frmTypeStr = "PE" Then
           gClsSearch.SearchMultiField "QryMastAccount", "Account,AcOurCode,CGName,AcContPer,AcCity,AgName,AcGSTIN,AcMsOmsParty", Array("Account Name", "Code", "Party Group", "Contact Person", "City", "Group", "TIN", "Ms/Oms"), Array(4000, 1000, 2000, 2000, 2000, 2500, 0, 0), "(Agcode in (90029) and AcIsExemptMill = 1 )", Trim(Chr(KeyAscii)), "", LeftPos, TopPos, True
        Else
           gClsSearch.SearchMultiField "QryMastAccount", "Account,AcOurCode,CGName,AcContPer,AcCity,AgName,AcGSTIN,AcMsOmsParty", Array("Account Name", "Code", "Party Group", "Contact Person", "City", "Group", "TIN", "Ms/Oms"), Array(4000, 1000, 2000, 2000, 2000, 2500, 0, 0), "(Agcode in (90015,90029) )", Trim(Chr(KeyAscii)), "", LeftPos, TopPos, True
        End If
        KeyAscii = 0
        lblTinNo.Caption = ""
        lblMsOms.Caption = ""
        txtCode(5) = ""
        If gClsSearch.SearchMultiRetCol(0) <> "" Then
            txttemp(Index) = gClsSearch.SearchMultiRetCol(0)
            lblTinNo.Caption = gClsSearch.SearchMultiRetCol(6)
            txtCode(5) = gClsSearch.SearchMultiRetCol(1)
            If gClsSearch.SearchMultiRetCol(7) = 0 Then
               lblMsOms.Caption = "M.S."
            Else
               lblMsOms.Caption = "O.M.S."
            End If
        End If
        AgCode = GProcGetColumnValue("tblMastAccount", "AcName", txttemp(5), "S", "AgCode", "N")
        If AgCode = 90029 Then
            txttemp(7) = txttemp(5) '-- Mill =Party
            txtCode(7) = txtCode(5)
            txttemp(7).Enabled = False
            txtCode(7) = txtCode(5)
            txtCode(7).Enabled = False
        Else
            txttemp(7).Enabled = True
            txtCode(7).Enabled = True
        End If
    Case 6 '-- Purchase A/c  Alias Account
        gClsSearch.SearchMultiField "tblMastaccount", "AcName,AcCity", Array("Purchase Account", "City"), Array(3000, 2000), "(Agcode=90025)", Trim(Chr(KeyAscii)), "", LeftPos, TopPos, True
        KeyAscii = 0
        If gClsSearch.SearchMultiRetCol(0) <> "" Then
            txttemp(Index) = gClsSearch.SearchMultiRetCol(0)
        End If
    Case 7  '-- Mill  Alias Account
         Dim MilTpCd As Long
         Dim ShowMill As String
         MilTpCd = GProcGetColumnValue("tblMastNarration", "Narration", "Trade", "S", "Narrcode", "N")
         ShowMill = "AgCode=90029 and   tblMastaccount.Accode = tblMastBillSerial.MillCode and  tblMastBillSerial.SaleType = " & MilTpCd & "  and tblMastBillSerial.Vfirm = '" & gCCode & "'  "
         gClsSearch.SearchMultiField "tblMastaccount,tblMastBillSerial", "BillSr,AcName,AcOurCode", Array("Mill Ini.", "Mill Name", "Code"), Array(0, 8000, 1000), ShowMill, Trim(Chr(KeyAscii)), "AcName", LeftPos, TopPos, True
'        gClsSearch.SearchMultiField "tblMastaccount", "AcName,AcOurCode,AcCity", Array("Mill Name", "Code", "City"), Array(3000, 1000, 2000), "(AgCode=90029 ) ", Trim(Chr(KeyAscii)), "AcName", LeftPos, TopPos, True
        KeyAscii = 0
        txtCode(7) = ""
        If gClsSearch.SearchMultiRetCol(1) <> "" Then
            txttemp(Index) = gClsSearch.SearchMultiRetCol(1)
            txtCode(7) = gClsSearch.SearchMultiRetCol(2)
        End If
        MillCtrlEd
    Case 8 '-- Broker
        gClsSearch.SearchMultiField "tblMastaccount", "AcName,AcCity,AcOurCode", Array("Broker", "City", ""), Array(5000, 3000, 0), "(AgCode=90016 )", Trim(Chr(KeyAscii)), "", LeftPos, TopPos, True
        KeyAscii = 0
        txtCode(8) = ""
        If gClsSearch.SearchMultiRetCol(0) <> "" Then
            txttemp(Index) = gClsSearch.SearchMultiRetCol(0)
            txtCode(8) = gClsSearch.SearchMultiRetCol(2)
        End If
    Case 35 '-- Transaction code
        gClsSearch.SearchMultiField "tblMastTransactionCode", "TranCode,iif(TranType=0,'Purchase',iif(TranType=1,'Purchase Goods Return',iif(TranType=2,'Purchase Debit Note',iif(TranType=3,'Sales', iif(TranType=4,'Sales Goods Return',iif(TranType=5,'Sales Credit Note', '') ) ) ) )),Nature,Description", Array("Tr. Code", "Tr. Type", "Nature", "Description"), Array(1000, 4000, 4000, 4000), "TranType = 0 and Trade = 1", "", "TranCode", 600 + Me.Left, 2500
        KeyAscii = 0
        If gClsSearch.SearchMultiRetCol(0) <> "" Then
            txttemp(Index) = gClsSearch.SearchMultiRetCol(0)
'            lblNature.Caption = GProcGetColumnValue("tblMastTransactionCode", "TranCode", txttemp(35), "S", "Nature", "S")
'            lblDescription.Caption = GProcGetColumnValue("tblMastTransactionCode", "TranCode", txttemp(35), "S", "Description", "S")
        End If
End Select

If Index = 11 Then '-- Transport
    gClsSearch.SearchMultiField "tblMastNarration", "Narration", Array("Transport"), Array(txttemp(Index).Width), " NarrType='R'", Trim(Chr(KeyAscii)), "Narration", 0, 0, True
    KeyAscii = 0
    If gClsSearch.SearchMultiRetCol(0) <> "" Then
        txttemp(Index) = gClsSearch.SearchMultiRetCol(0) '--Transport
    End If
End If
End Sub
Private Sub MillCtrlEd()
Dim ExemptMill As Long
'-- Is Exempt Mill
ExemptMill = GProcGetColumnValue("TblMastAccount", "AcName", txttemp(7), "S", "ACIsExemptMill", "N")
If ExemptMill = 1 Then
   LblExmptAmt.Visible = True
   LblExmptKg.Visible = True
   txt(19).Visible = True
   txt(20).Visible = True
    If FormAction = vbDataActionAddNew Or FormAction = vbDataActionUpdate Then
        txt(19).Enabled = True
        txt(20).Enabled = True
    End If
Else
    txt(19).Enabled = False
    txt(20).Enabled = False
    LblExmptAmt.Visible = False
    LblExmptKg.Visible = False
    txt(19).Visible = False
    txt(20).Visible = False
    txt(19) = 0
    txt(20) = 0
End If
End Sub
Private Sub txtTemp_LostFocus(Index As Integer)
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
'                Case 2 '-- Booking Date
'                     BookingData
'                     If .Col = 2 Then txtGrid.Locked = True

'                    txtGrid.Locked = True
                Case 3 '-- Count
                     BookingData
                     txtGrid.Locked = False
                Case 8 '-- rate
                    txtGrid.Locked = True
                Case 9 '-- rate Per
                    txtGrid.Locked = True
                Case Else
                    txtGrid.Locked = False
                    If .Col = 6 Then
                       OldBag = .TextMatrix(.Row, 6)
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
        If ((Val(.TextMatrix(.Row, 6)) = 0) Or (Val(.TextMatrix(.Row, 10)) = 0)) And .rows > 2 Then
            .RemoveItem (.Row)
            txtGrid.text = ""
        End If
        txtGrid.Visible = False
        txt(17).SetFocus
        Exit Sub
    End If
End With
'If FormAction <> vbDataActionAddNew And FormAction <> vbDataActionUpdate Then Exit Sub
'---- Navigate Row ,Col
If KeyAscii = 13 Then
    KeyAscii = 0
    With msGrid
        If .Col = 10 Then 'rstblSub.Fields.Count - 2 Then
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
   If msGrid.Col = 1 Or msGrid.Col = 2 Then
      msGrid.Col = 3
      If msGrid.Enabled = True Then
         msGrid.SetFocus
      End If
   End If
'   If msGrid.Col = 4 Or msGrid.Col = 5 Then
   If msGrid.Col = 4 Then
      msGrid.Col = 5
      KeyAscii = 0
      If msGrid.Enabled = True Then
          msGrid.SetFocus
      End If
   End If
   If msGrid.Col = 8 Then
   
      msGrid.Col = 10
      msGrid.SetFocus
   End If
   txtGrid.Move msGrid.CellLeft + msGrid.Left, msGrid.CellTop + msGrid.Top, msGrid.CellWidth, msGrid.CellHeight
End If
txtGrid.Visible = True
txtGrid.Enabled = True
DoEvents
If KeyAscii <> 0 Then
    Select Case rstblSub.Fields(msGrid.Col).Type
        Case adChar, adVarWChar
            If msGrid.Col = 1 Or msGrid.Col = 3 Or msGrid.Col = 4 Then '-- Count ,godown
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
Dim CalWt As Double
CheckForNumberGrid

If msGrid.TextMatrix(msGrid.Row, 2) = 0 Then msGrid.TextMatrix(msGrid.Row, 2) = gCYear '-- Book year
If msGrid.TextMatrix(msGrid.Row, 1) = 0 Then msGrid.TextMatrix(msGrid.Row, 2) = 0 '-- Book year
MillCode = GProcGetColumnValue("TblMastAccount", "AcName", txttemp(5), "S", "Accode", "N")
StdPack = GProcGetColumnValue("TblMastItem", "ITName", msGrid.TextMatrix(msGrid.Row, 3), "S", "ItStdpack", "N", " ItMillCode=" & MillCode)
'msGrid.TextMatrix(msGrid.Row, 9) = StdPack
If msGrid.Row = 0 Then Exit Sub
'-- Wt=Bag * Std Pack


If OldBag <> msGrid.TextMatrix(msGrid.Row, 6) And StdPack > 0 And FormAction = vbDataActionAddNew Then
    msGrid.TextMatrix(msGrid.Row, 7) = msGrid.TextMatrix(msGrid.Row, 6) * StdPack
    OldBag = msGrid.TextMatrix(msGrid.Row, 6)
End If
'-- Wt=Bag * Std Pack
'If msGrid.TextMatrix(msGrid.Row, 7) = 0 Then
'    msGrid.TextMatrix(msGrid.Row, 7) = msGrid.TextMatrix(msGrid.Row, 6) * StdPack
'End If
'-- Rate=Amt*Rate Per/Wt
If CDbl(msGrid.TextMatrix(msGrid.Row, 7)) > 0 Then
    msGrid.TextMatrix(msGrid.Row, 8) = CDbl(msGrid.TextMatrix(msGrid.Row, 10)) * CDbl(msGrid.TextMatrix(msGrid.Row, 9)) / CDbl(msGrid.TextMatrix(msGrid.Row, 7))
Else
    msGrid.TextMatrix(msGrid.Row, 8) = 0
End If
msGrid.TextMatrix(msGrid.Row, 8) = Format(msGrid.TextMatrix(msGrid.Row, 8), LStr)

'-- Amt=Wt*Rate/Rate Per
'If CDbl(msGrid.TextMatrix(msGrid.Row, 9)) > 0 Then
'    msGrid.TextMatrix(msGrid.Row, 10) = CDbl(msGrid.TextMatrix(msGrid.Row, 7)) * CDbl(msGrid.TextMatrix(msGrid.Row, 8)) / CDbl(msGrid.TextMatrix(msGrid.Row, 9))
'Else
'    msGrid.TextMatrix(msGrid.Row, 10) = 0
'End If
msGrid.TextMatrix(msGrid.Row, 10) = Format(msGrid.TextMatrix(msGrid.Row, 10), LStr)
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
    Case 3 '-- Item
        GProcShowForm frmMastItem, frmMain.mnuMstfrm(3), Array(frmMastItem.txttemp(9)), Array(txttemp(7))
        Exit Sub
    Case 4 '-- Godown
        OldgFrmTypeStr = gfrmTypeStr
        gfrmTypeStr = "G"
        GProcShowForm frmMastNarration, frmMain.mnuMstfrm(8), "G"
        gfrmTypeStr = OldgFrmTypeStr
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
'            Case 1 '--Booking
''                PrepareBooking Chr(KeyAscii)
'                KeyAscii = 0
            Case 3 '-- Count
                If msGrid.TextMatrix(msGrid.Row, 1) > 0 Then
                    KeyAscii = 0
                Else
                    MillCode = GProcGetColumnValue("TblMastAccount", "AcName", txttemp(7), "S", "Accode", "N")
'                    gClsSearch.SearchMultiField "tblMastItem,tblMastaccount", "ItName,AcName,ItStdpack,ItStdrateper", Array("Count Name", "Mill", "Std. Pack", "Rate Per"), Array(txtGrid.Width, 2000, 1000, 1000), "tblMastItem.ITMillCode=tblMastaccount.acCode and ITMillCode=" & MillCode, Chr(KeyAscii), "ItName", LeftPos, TopPos, True
                    gClsSearch.SearchMultiField "tblMastItem,tblMastaccount", "ItName,AcName,ItStdpack,ItStdrateper,InDailyReport", Array("Count Name", "Mill", "Std. Pack", "Rate Per", ""), Array(5000, 5000, 1000, 1000, 0), "tblMastItem.ITMillCode=tblMastaccount.acCode and ITMillCode=" & MillCode, Chr(KeyAscii), "ItName", LeftPos, TopPos, True
                    If gClsSearch.SearchMultiRetCol(4) = 1 Then
                       MsgBox "This is No Use Count "
                    End If
                    txtGrid = ""
                    KeyAscii = 0
                    If gClsSearch.SearchMultiRetCol(0) <> "" And gClsSearch.SearchMultiRetCol(4) = 0 Then '-- Item Wt ,Rate From Master
                        txtGrid = gClsSearch.SearchMultiRetCol(0)
                        If IsNumeric(.TextMatrix(.Row, 7)) = False Then .TextMatrix(.Row, 7) = 0
'                        If IsNumeric(.TextMatrix(.Row, 9)) = False Then .TextMatrix(.Row, 9) = 0
                        .TextMatrix(.Row, 9) = gClsSearch.SearchMultiRetCol(3)
                        If Val(txt(23)) + Val(txt(26)) + Val(txt(37)) = 0 Then
                            ItemType = GProcGetColumnValue("tblMastItem", "ItName", txtGrid, "S", "ItType", "N", "ItMillcode=" & MillCode) '--Item Code
                            If ItemType = 0 Then
                               If (Left(gCGStin, 2) = Left(lblTinNo.Caption, 2) Or Left(lblTinNo.Caption, 2) = "" Or Left(lblTinNo.Caption, 1) = "U") Then
                                  txt(23) = GProcGetColumnValue("tblMastNarration", "Narration", "G S T", "S", "CotCGSTRt", "N")
                                  txt(26) = GProcGetColumnValue("tblMastNarration", "Narration", "G S T", "S", "CotSGSTRt", "N")
                                  txt(37) = 0
                               Else
                                  txt(23) = 0
                                  txt(26) = 0
                                  txt(37) = GProcGetColumnValue("tblMastNarration", "Narration", "G S T", "S", "CotIGSTRt", "N")
                               End If
                            Else
                               If (Left(gCGStin, 2) = Left(lblTinNo.Caption, 2) Or Left(lblTinNo.Caption, 2) = "" Or Left(lblTinNo.Caption, 1) = "U") Then
                                  txt(23) = GProcGetColumnValue("tblMastNarration", "Narration", "G S T", "S", "PolCGSTRt", "N")
                                  txt(26) = GProcGetColumnValue("tblMastNarration", "Narration", "G S T", "S", "PolSGSTRt", "N")
                                  txt(37) = 0
                               Else
                                  txt(23) = 0
                                  txt(26) = 0
                                  txt(37) = GProcGetColumnValue("tblMastNarration", "Narration", "G S T", "S", "PolIGSTRt", "N")
                               End If
                            End If
                        End If
                        txtGrid_KeyPress 13
                    End If
                  End If
            Case 4  '-- Godown
                gClsSearch.SearchMultiField "tblMastNarration", "Narration", Array("Godown"), Array(6000), " NarrType='G' ", Chr(KeyAscii), "Narration", LeftPos, TopPos, True
                txtGrid = ""
                KeyAscii = 0
                If gClsSearch.SearchMultiRetCol(0) <> "" Then
                    txtGrid = gClsSearch.SearchMultiRetCol(0)
                    txtGrid_KeyPress 13
                End If
            Case Else
                KeyAscii = GProcValidateKey(rstblSub, msGrid.Col, KeyAscii, txtGrid)
        End Select
   End If
    If KeyAscii = 13 Then
        CheckForNumberTxtGrid
        If txtGrid = "" And .Col = 3 Then
           KeyAscii = 27
        End If
    End If
    '--- Navigate Row ,Col
    If KeyAscii = 13 And .Col < 10 Then 'rstblSub.Fields.Count - 2 Then
        .Col = .Col + 1
'        If .Col = 6 Then
'            If IsNumeric(txtGrid) = True Then '-Auto LR No
'                tmpAutoLRNo = CDbl(txtGrid)
'            Else
'                tmpAutoLRNo = ""
'            End If
'        End If
    ElseIf KeyAscii = 13 And .Col = 10 Then '= rstblSub.Fields.Count - 2 Then
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
        If (Val(.TextMatrix(.Row, 8)) = 0) And .rows > 2 Then
            .RemoveItem (.Row)
            txtGrid.text = ""
        End If
        txtGrid.Visible = False
        txt(17).SetFocus
        Exit Sub
    End If
End With
End Sub

'-- Grid Required
Private Function ValidateGrid() As Boolean
Dim i As Long
Dim MillCode As Long
Dim RowIndex As Long
Dim ItCode As Long
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
                Case 3 '-- Count
                 '-Rate=0 ,count,wt=0 then delete row
                    If RowIndex > 1 And Trim(.TextMatrix(RowIndex, 3)) = "" And CDbl(.TextMatrix(RowIndex, 7)) = 0 And CDbl(.TextMatrix(RowIndex, 8)) = 0 Then
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
                    



'                Case 4 '-- Godown
'                    If Trim(.TextMatrix(RowIndex, i)) = "" Then
'                        MsgBox "Check Godown.", vbCritical + vbOKOnly, Me.Caption
'                        .Col = i
'                        msGrid.SetFocus
'                        Exit Function
'                    End If
'                    If GProcGetColumnValue("TblMastNarration", "Narration", .TextMatrix(RowIndex, i), "S", "Narrcode", "N") = 0 Then
'                        'MSGrid.SetFocus
'                        oldFrmTypeStr = gfrmTypeStr
'                        gfrmTypeStr = "G"
'                        GProcShowForm frmMastNarration, frmMain.mnuMstfrm(8), Array(frmMastNarration.txt(2)), Array(.TextMatrix(RowIndex, i))
'                        gfrmTypeStr = oldFrmTypeStr
'                        Exit Function
'                    End If
                    
                Case 5 '-- Lot No
'                    If Trim(.TextMatrix(RowIndex, i)) = "" Then
'                        MsgBox "Check Lot No.", vbCritical + vbOKOnly, Me.Caption
'                        .Col = i
'                        MSGrid.SetFocus
'                        Exit Function
'                    End If
'                Case 6 '-- LR No
'                    If Trim(.TextMatrix(RowIndex, i)) = "" Then
'                        MsgBox "Check LR No.", vbCritical + vbOKOnly, Me.Caption
'                        .Col = i
'                        msGrid.SetFocus
'                        Exit Function
'                    End If
'                Case 7 '-- LR Date
'                    If IsDate(.TextMatrix(RowIndex, i)) = False Then
'                        MsgBox "Check LR Date.", vbCritical + vbOKOnly, Me.Caption
'                        .Col = i
'                        msGrid.SetFocus
'                        Exit Function
'                    End If
'                Case 8 '-- LR Amt
'                    If CDbl(.TextMatrix(RowIndex, i)) = 0 Then
'                        MsgBox "Check LR Amount.", vbCritical + vbOKOnly, Me.Caption
'                        .Col = i
'                        MSGrid.SetFocus
'                        Exit Function
'                    End If
                Case 6 '-- Bag
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
                Case 7 '-- Wt
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
                Case 10 '-- amount
                    If CDbl(.TextMatrix(RowIndex, i)) = 0 Then
                        MsgBox "Check amount.", vbCritical + vbOKOnly, Me.Caption
                        .Col = i
                        msGrid.SetFocus
                        Exit Function
                    End If
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
        If CLng(.TextMatrix(RowIndex, 1)) > 0 Then
            MillCode = GProcGetColumnValue("TblMastAccount", "AcName", txttemp(5), "S", "Accode", "N")
            ItCode = GProcGetColumnValue("tblMastItem", "ItName", .TextMatrix(RowIndex, 3), "S", "ItCode", "N", " ItMillCode=" & MillCode)
            X = "Select * from tblBooking where Vno=" & CLng(.TextMatrix(RowIndex, 1)) & " and  " _
            & " BkItCode=" & ItCode & " and  BkSrNo=" & .TextMatrix(RowIndex, 13) & " and Vtype='OT'  and VYear = " & .TextMatrix(RowIndex, 2) & " and Vfirm = '" & gCCode & "'"
            GProcRstOpen Rs1, X, "R"
            If Rs1.EOF And Rs1.BOF Then
                MsgBox "Check Booking No.", vbInformation + vbOKOnly, Me.Caption
                .Col = 1
                msGrid.SetFocus
                Exit Function
            End If
            Rs1.Close
        End If
      Next RowIndex
End With
ValidateGrid = True
ReNumberGridSrno
End Function
Private Sub FillCombo()
'Call GProcFillCombo(Cbo(22), "select Narration from tblMastNarration where NarrType ='T' order by Narration")
'If gfrmHunkStr = "HNK" Then
'   Cbo(22) = GProcGetColumnValue("tblMastNarration", "NarrCode", CStr(gTaxCodePYHunk), "N", "Narration", "S")
'Else
'   Cbo(22) = GProcGetColumnValue("tblMastNarration", "NarrCode", CStr(gTaxCodePY), "N", "Narration", "S")
'End If
''If FormAction = vbDataActionAddNew And Val(txt(23)) = 0 Then
''    txt(23) = GProcGetColumnValue("tblMastNarration", "Narration", Cbo(22), "S", "MastTaxRate", "N")
''    txt(23) = Format(txt(23), FStr)
''End If
End Sub
Private Sub SaveBtnEd()
'--- Vno,vdt,Party,A/c,Mill,Amount
If (CDbl(txt(0)) = 0 And FormAction <> vbDataActionAddNew) Or GProcIsDateValid(dtpDt(2)) = False Or txttemp(5) = "" Or txttemp(6) = "" Or txttemp(7) = "" Or CDbl(txt(16)) = 0 Or GProcIsDateValid(dtpDt(10)) = False Then
    cmdBtn(13).Enabled = False
    cmdBtn(13).BackColor = &H80000004
Else
    cmdBtn(13).Enabled = True
    cmdBtn(13).BackColor = &HFFFF80
End If
End Sub
Private Function ValidateData() As Boolean
Dim i As Long
Dim Y As String
Dim GName As String
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
'    If gCisGst = 1 Then
'        If GProcValidateTranCd(frmTypeStr, Cbo(22).text, txtTemp(35), lblTinNo.Caption, lblMsOms.Caption, False, IIf(gfrmHunkStr = "HNK", True, False)) = False Then
'           Select Case txtTemp(35).text
'                  Case "10"
'                       If gfrmHunkStr = "HNK" Then
'                          MsgBox "Required Data Tax = 'Tax Free', Tin No Should Not Balnk, And M.S. Party", vbInformation + vbOKOnly, Me.Caption
'                       Else
'                          MsgBox "Required Data Tax = 'V A T', Tin No Should Not Balnk, And M.S. Party", vbInformation + vbOKOnly, Me.Caption
'                       End If
'                  Case "20"
'                       MsgBox "Required Data Tax = 'U R D' And Tin No Balnk", vbInformation + vbOKOnly, Me.Caption
'                  Case "70"
'                       MsgBox "Required Data Tax = 'C S T', Tin No Should Not Balnk And O.M.S. Party", vbInformation + vbOKOnly, Me.Caption
'                  Case "40"
'                       MsgBox "Required Data Tax = 'C S T Against Form C', Tin No Should Not Balnk And O.M.S. Party", vbInformation + vbOKOnly, Me.Caption
'           End Select
'           txtTemp(35).SetFocus
'           Exit Function
'       End If
'    End If
    
    '---- Required Data
    
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
''    If Cbo(22) = "" Then
''        MsgBox "Check Tax Type.", vbInformation + vbOKOnly, Me.Caption
''        Cbo(22).SetFocus
''        Exit Function
''    End If
''
''
''    '-- Tax %
''    If CDbl(txt(23)) > 0 And Cbo(22) = "Tax Free" Then
''        MsgBox "Check Tax %.", vbInformation + vbOKOnly, Me.Caption
''        txt(23).SetFocus
''        Exit Function
''    End If
    
    
    '-- Tax Amount
    If (CDbl(txt(24)) > 0 Or CDbl(txt(23)) > 0) And Abs(Round(Val(txt(21)) * Val(txt(23)) / 100, 2) - Val(txt(24))) > 1 And FormAction = vbDataActionAddNew Then
        MsgBox "Tax Amount on Ass.Val not correct", vbInformation + vbOKOnly, Me.Caption
        txt(24).SetFocus
        Exit Function
    End If
    '--CGST Tax Amt
    If (CDbl(txt(27)) > 0 Or CDbl(txt(26)) > 0) And Abs(Round(Val(txt(21)) * Val(txt(26)) / 100, 2) - Val(txt(27))) > 1 And FormAction = vbDataActionAddNew Then
        MsgBox "Tax Amount on Ass.Val not correct", vbInformation + vbOKOnly, Me.Caption
        txt(27).SetFocus
        Exit Function
    End If
    '--IGST Tax Amt
    If (CDbl(txt(38)) > 0 Or CDbl(txt(37)) > 0) And Abs(Round(Val(txt(21)) * Val(txt(37)) / 100, 2) - Val(txt(38))) > 1 And FormAction = vbDataActionAddNew Then
        MsgBox "Tax Amount on Ass.Val not correct", vbInformation + vbOKOnly, Me.Caption
        txt(38).SetFocus
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
           GName = GProcGetColumnValue("TblMastGroup", "AgCode", 90015, "N", "AgName", "S")
           GProcShowForm frmMastAccount, frmMain.mnuMstfrm(0), Array(frmMastAccount.txt(1), frmMastAccount.txttemp(2)), Array(txttemp(5), GName)
           Exit Function
        End If
    '---- Purchase Account
    If txttemp(6) = "" Then
        MsgBox "Check Purchase Account.", vbInformation + vbOKOnly, Me.Caption
        txttemp(6).SetFocus
        Exit Function
    End If
        '--Create Party
        If GProcGetColumnValue("tblMastAccount", "AcName", txttemp(6), "S", "AcCode", "N") = 0 Then
           GName = GProcGetColumnValue("TblMastGroup", "AgCode", 90025, "N", "AgName", "S")
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
    '---- Godown Name
    If txtGodown = "" Then
        MsgBox "Check Godown Name", vbInformation + vbOKOnly, Me.Caption
        txtGodown.SetFocus
        Exit Function
    End If
    
    
    
    '---- Amount
    If CDbl(txt(16)) <= 0 Then
        MsgBox "Check Voucher Amount.", vbInformation + vbOKOnly, Me.Caption
        Exit Function
    End If
    '---Check Dublicate Bill No
    If Trim(txt(9)) <> "" Then
        X = "Select vno From " & tblName & " Where PurBillNo= '" & txt(9) & "' and PurAcCrCode = " & CLng(txt(5)) & "  and  vno <> " & CLng(txt(0)) & "  and Vtype='" & frmTypeStr & "'  and Vfirm = '" & gCCode & "'  and VYear=" & gCYear
        i = GProcRstOpen(rstAddCheck, X, "R")
        If i > 0 Then
            MsgBox "Same Bill No of Same Party is Available in Voucher No. " & rstAddCheck.Fields(0), vbCritical + vbOKOnly, "Duplicate Record"
'            txt(9).SetFocus
 '           Exit Function
        End If
    End If
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
'If IsDate(dtpDt(10)) = False Then dtpDt(10) = "__/__/____" '--BillDt
'If IsDate(dtpDt(32)) = False Then dtpDt(32) = "__/__/____" '--Due Dt
'If IsDate(dtpDt(14)) = False Then dtpDt(14) = "__/__/____" '--lr Dt

If IsDate(dtpDt(2)) Then '--Vdt
    txt(2) = CDate(dtpDt(2))
Else
    txt(2) = ""
End If
If IsDate(dtpDt(10)) Then '--Party Bill Date
    txt(10) = CDate(dtpDt(10))
Else
    txt(10) = ""
End If
If IsDate(dtpDt(32)) Then '--Due Date
    txt(32) = CDate(dtpDt(32))
Else
    txt(32) = ""
End If
If IsDate(dtpDt(14)) Then '--lr Date
    txt(14) = CDate(dtpDt(14))
Else
    txt(14) = ""
End If

'-- Cr Code
txt(5) = GProcGetColumnValue("tblMastAccount", "AcName", txttemp(5), "S", "AcCode", "N")
'-- Dr Code
txt(6) = GProcGetColumnValue("tblMastAccount", "AcName", txttemp(6), "S", "AcCode", "N")
'--Mill
txt(7) = GProcGetColumnValue("tblMastAccount", "AcName", txttemp(7), "S", "AcCode", "N")

'-- broker
If txttemp(8) = "" Then
   txttemp(8) = "-"
End If

txt(8) = GProcGetColumnValue("tblMastAccount", "AcName", txttemp(8), "S", "AcCode", "N")
'-- Transport
If txttemp(11) = "" Then
   txttemp(11) = "-"
End If
txt(11) = GProcGetColumnValue("tblMastNarration", "Narration", txttemp(11), "S", "NarrCode", "N")
txt(39) = GProcGetColumnValue("tblMastAccount", "AcName", "-", "S", "AcCode", "N")

'-- TaxType
txt(22) = GProcGetColumnValue("tblMastNarration", "Narration", "-", "S", "NarrCode", "N")
txt(35) = GProcGetColumnValue("tblMastTransactionCode", "TranCode", "-", "S", "TranCode", "S")

txt(3) = gCYear
txt(4) = gCCode
txt(65) = GProcGetColumnValue("tblMastAccount", "AcName", txttemp(65), "S", "AcCode", "N")

ShowAccountBalance
End Sub
Private Sub FillTempFromTxt()
Dim MsOms As Integer
FillCombo
If IsDate(txt(2)) Then '--VDt
    dtpDt(2) = txt(2)
End If
If IsDate(txt(10)) Then '--Party Bill Date
    dtpDt(10) = txt(10)
End If
If IsDate(txt(32)) Then '--Due Date
    dtpDt(32) = txt(32)
End If
If IsDate(txt(14)) Then '--LR Date
    dtpDt(14) = txt(14)
End If
'-- Cr Code
txttemp(5) = GProcGetColumnValue("tblMastAccount", "AcCode", txt(5), "N", "AcName", "S")
lblTinNo.Caption = GProcGetColumnValue("tblMastAccount", "AcCode", txt(5), "N", "AcGSTIN", "S")
MsOms = GProcGetColumnValue("tblMastAccount", "AcCode", txt(5), "N", "AcMsOmsParty", "N")
If MsOms = 0 Then
   lblMsOms.Caption = "M.S."
Else
   lblMsOms.Caption = "O.M.S."
End If
txtCode(5) = GProcGetColumnValue("tblMastAccount", "AcCode", txt(5), "N", "AcOurCode", "S")

'-- Dr Code
txttemp(6) = GProcGetColumnValue("tblMastAccount", "AcCode", txt(6), "N", "AcName", "S")
'--Mill
txttemp(7) = GProcGetColumnValue("tblMastAccount", "AcCode", txt(7), "N", "AcName", "S")
txtCode(7) = GProcGetColumnValue("tblMastAccount", "AcCode", txt(7), "N", "AcOurCode", "S")

'--Broker
txttemp(8) = GProcGetColumnValue("tblMastAccount", "AcCode", txt(8), "N", "AcName", "S")
txtCode(8) = GProcGetColumnValue("tblMastAccount", "AcCode", txt(8), "N", "AcOurCode", "S")

'--Transport
txttemp(11) = GProcGetColumnValue("tblMastNarration", "NarrCode", txt(11), "N", "Narration", "S")

'--Taxpayable Ac
'txttemp(39) = GProcGetColumnValue("tblMastAccount", "AcCode", txt(39), "N", "AcName", "S")
'-- TaxType
'Cbo(22) = GProcGetColumnValue("tblMastNarration", "NarrCode", txt(22), "N", "Narration", "S")
'txttemp(35) = GProcGetColumnValue("tblMastTransactionCode", "TranCode", txt(35), "S", "TranCode", "S")
'lblNature.Caption = GProcGetColumnValue("tblMastTransactionCode", "TranCode", txt(35), "S", "Nature", "S")
'lblDescription.Caption = GProcGetColumnValue("tblMastTransactionCode", "TranCode", txt(35), "S", "Description", "S")
'-- Fill Related Record
txtCode(65) = GProcGetColumnValue("tblMastAccount", "AcCode", txt(65), "N", "AcOurCode", "S")
txttemp(65) = GProcGetColumnValue("tblMastAccount", "AcCode", txt(65), "N", "AcName", "S")
FillRelatedRecord
ShowAccountBalance
MillCtrlEd
End Sub
Private Sub FillRelatedRecord()
Dim RowIndex As Long
Dim ColIndex As Long
Dim TotBag As Long
Dim TotWt As Double
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
                If ColIndex = 6 Then TotBag = TotBag + CLng(msGrid.TextMatrix(RowIndex, 6)) '-- Tot Bag
                If ColIndex = 7 Then TotWt = TotWt + CDbl(msGrid.TextMatrix(RowIndex, 7)) '-- Tot Wt
                txtGodown = msGrid.TextMatrix(RowIndex, 4)
            Next ColIndex
            RowIndex = RowIndex + 1
            .MoveNext
            If Not .EOF Then msGrid.rows = msGrid.rows + 1
        Loop
    End If
End With
lblPurBgas.Caption = TotBag
lblPurWeight.Caption = TotWt




'--- Fill Related Control
SetRelatedRS
End Sub
Private Sub ReadFields()
FormAction = vbDataActionClose
GProcShowRecord Me, rstbl, MaxNo, OtherTxtFromNo, OtherTxtToNo
FillTempFromTxt
End Sub
Private Sub CalAmount()
Dim TaxAmt As Double
Dim TaxAmtCGST As Double
Dim TaxAmtIGST As Double
Dim SubAmt As Double
Dim Amt As Double
Dim ROff As Double
Dim ExmptAmt As Double
Dim ExmptKg As Double

Dim AssValue As Double
Dim RowIndex As Integer
Dim TotBag As Long
Dim TotWt As Double
Dim LCost As Double
Dim wMillCode As Long
Dim IsExemptMill As Long
Dim wExemptAmt As Double
Dim wExemptRt As Double
Dim TaxCd As Long
Dim TcsAmt As Double
Dim X As String
Dim wPurPtyCode As Long
Dim tmpRst As Recordset
Dim TDSLimitAmt As Double
Dim TdsAmt As Double
Dim PtyPan As String
Dim TdsOnAmt As Double




'-- Is Exempt Mill
IsExemptMill = GProcGetColumnValue("TblMastAccount", "AcName", txttemp(5), "S", "ACIsExemptMill", "N")
wMillCode = GProcGetColumnValue("tblMastAccount", "AcName", txttemp(7), "S", "AcCode", "N")
wPurPtyCode = GProcGetColumnValue("tblMastAccount", "AcName", txttemp(5), "S", "AcCode", "N")
PtyPan = GProcGetColumnValue("tblMastAccount", "AcName", txttemp(5), "S", "AcPAN", "S")

If PtyPan = "" Then PtyPan = "XXXX"
'-- Due Date
If IsNumeric(txt(31)) = False Then txt(31) = "0"
txt(31) = Format(txt(31), PStr)
If IsDate(dtpDt(2)) Then dtpDt(32) = DateAdd("d", txt(31), dtpDt(2))
'---- Sub amt
CheckForNumberGrid
CheckForNumber
RowIndex = 1
wExemptAmt = 0
With msGrid
    Do While RowIndex < .rows
        '-- sub amt
        If Val(.TextMatrix(RowIndex, 10)) > 0 Then
            SubAmt = SubAmt + Val(.TextMatrix(RowIndex, 10))
        End If
        '--Tot Bag
        If Val(.TextMatrix(RowIndex, 6)) > 0 Then
            TotBag = TotBag + Val(.TextMatrix(RowIndex, 6))
        End If
        '--Tot Wt
        If Val(.TextMatrix(RowIndex, 7)) > 0 Then
            TotWt = TotWt + Val(.TextMatrix(RowIndex, 7))
        End If
        If IsExemptMill = 1 Then
           wExemptRt = GProcGetColumnValue("tblMastItem", "ItName", msGrid.TextMatrix(RowIndex, 3), "S", "ExemptRt", "N", "ItMillcode=" & wMillCode) '--Item Code
           If wExemptRt = 100 Then
              wExemptAmt = wExemptAmt + Val(.TextMatrix(RowIndex, 10))
           Else
              wExemptAmt = wExemptAmt + Round((Val(.TextMatrix(RowIndex, 10)) * wExemptRt) / 100, 2)
           End If
        End If
        If Val(txt(23)) + Val(txt(26)) + Val(txt(37)) = 0 And msGrid.TextMatrix(1, 3) <> "" Then
           ItemType = GProcGetColumnValue("tblMastItem", "ItName", msGrid.TextMatrix(1, 3), "S", "ItType", "N", "ItMillcode=" & wMillCode) '--Item Code
           If ItemType = 0 Then
              If (Left(gCGStin, 2) = Left(lblTinNo.Caption, 2) Or Left(lblTinNo.Caption, 2) = "" Or Left(lblTinNo.Caption, 1) = "U") Then
                 txt(23) = GProcGetColumnValue("tblMastNarration", "Narration", "G S T", "S", "CotCGSTRt", "N")
                 txt(26) = GProcGetColumnValue("tblMastNarration", "Narration", "G S T", "S", "CotSGSTRt", "N")
                 txt(37) = 0
              Else
                 txt(23) = 0
                 txt(26) = 0
                 txt(37) = GProcGetColumnValue("tblMastNarration", "Narration", "G S T", "S", "CotIGSTRt", "N")
              End If
           Else
              If (Left(gCGStin, 2) = Left(lblTinNo.Caption, 2) Or Left(lblTinNo.Caption, 2) = "" Or Left(lblTinNo.Caption, 1) = "U") Then
                 txt(23) = GProcGetColumnValue("tblMastNarration", "Narration", "G S T", "S", "PolCGSTRt", "N")
                 txt(26) = GProcGetColumnValue("tblMastNarration", "Narration", "G S T", "S", "PolSGSTRt", "N")
                 txt(37) = 0
              Else
                 txt(23) = 0
                 txt(26) = 0
                 txt(37) = GProcGetColumnValue("tblMastNarration", "Narration", "G S T", "S", "PolIGSTRt", "N")
              End If
           End If
        End If
        RowIndex = RowIndex + 1
    Loop
End With

lblPurBgas.Caption = TotBag
lblPurWeight.Caption = TotWt



'--- Subamt+Add -Less
If wExemptAmt > 0 Then
   txt(19) = wExemptAmt
Else
   txt(19) = 0
End If
AssValue = SubAmt + CDbl(txt(17)) - CDbl(txt(18)) - CDbl(txt(19))
'- 1 SGST

    If Abs((CDbl(txt(24)) - (AssValue * CDbl(txt(23)) / 100))) > 1 Then
       TaxAmt = AssValue * CDbl(txt(23)) / 100
    Else
       TaxAmt = CDbl(txt(24))
    End If
    TaxAmt = Format(TaxAmt, FStr)
      
    '- 2 CGST
    If Abs((CDbl(txt(27)) - (AssValue * CDbl(txt(26)) / 100))) > 1 Then
       TaxAmtCGST = AssValue * CDbl(txt(26)) / 100
    Else
       TaxAmtCGST = CDbl(txt(27))
    End If
    TaxAmtCGST = Format(TaxAmtCGST, FStr)
       
    '- 3 IGST
    If Abs((CDbl(txt(38)) - (AssValue * CDbl(txt(37)) / 100))) > 1 Then
       TaxAmtIGST = AssValue * CDbl(txt(37)) / 100
    Else
       TaxAmtIGST = CDbl(txt(38))
    End If
    TaxAmtIGST = Format(TaxAmtIGST, FStr)

'---TCS
    If Format(dtpDt(10).Value, "YYYY/MM/dd") >= Format("2020/01/10", "YYYY/MM/dd") Then
       txt(69) = "0.075"
       If Val(txt(68)) > 0 And Val(txt(70)) = 0 Then
          TcsAmt = (Val(txt(68)) * Val(txt(69))) / 100
          TcsAmt = Format(GProcMakeRounding(TcsAmt), FStr)
       Else
          TcsAmt = txt(70)
       End If
    End If
'--Roff

ROff = txt(30) '= 'CDbl(GProcMakeRounding(Amt) - Amt)
'-- SubAmt+Tax Amt +Add -Less+roff + Aft Tax Add - Aft Tax Less
Amt = SubAmt + TaxAmt + TaxAmtCGST + TaxAmtIGST + TcsAmt + CDbl(txt(17)) - CDbl(txt(18)) + ROff + CDbl(txt(28)) - CDbl(txt(29))
txt(15) = SubAmt
txt(24) = TaxAmt
txt(27) = TaxAmtCGST
txt(38) = TaxAmtIGST
txt(21) = AssValue
txt(16) = Amt
txt(70) = TcsAmt

'If Abs(SubAmt + CDbl(txt(17)) - CDbl(txt(18)) - AssValue) < 2 Then
'    ExmptAmt = 0
'    ExmptKg = 0
'End If
'txt(19) = 0 '--ExmptAmt
If Val(TotWt) > 0 Then
   txt(20) = Val(txt(19)) / TotWt '--ExmptKg
Else
   txt(20) = 0
End If
CheckForNumber
'TaxCd = GProcGetColumnValue("tblMastNarration", "Narration", Cbo(22).text, "S", "MastTaxPurAcCode", "N")
'txttemp(39) = GProcGetColumnValue("tblMastAccount", "AcCode", CStr(TaxCd), "N", "AcName", "S")
'--TDS
TdsAmt = 0
TdsOnAmt = 0
txt(63) = 0
txtCode(65) = ""
txttemp(65) = "-"
If Format(dtpDt(10).Value, "YYYY/MM/dd") >= Format("2021/07/01", "YYYY/MM/dd") And (gCIsTcsFirm = 1 Or gCIsDepotFirm = 1) Then
    X = "Select sum(PurBillAmt) from tblpurch,tblMastAccount where tblpurch.PurAcCrCode = tblMastAccount.AcCode and (PurAcCrCode = " & wPurPtyCode & " or tblMastAccount.AcPan = '" & PtyPan & "' ) and PurBillDt <= '" & Format(dtpDt(10).Value, "YYYY/MM/dd") & "'"
    X = X + " and  ltrim(str(vno))+vtype+vFirm+ltrim(str(vyear)) <> '" & Trim(Str(txt(0))) & "" & gfrmTypeStr & "" & gCCode & "" & LTrim(Str(gCYear)) & "' and vyear = " & gCYear & " and tblpurch.Vtype in ('PY','PO') and vfirm = '" & gCCode & "' "
    i = GProcRstOpen(tmpRst, X, "R")
       
    TDSLimitAmt = IIf(IsNull(tmpRst.Fields(0)), 0, tmpRst.Fields(0))
    TDSLimitAmt = TDSLimitAmt + Amt
    lblPurchAmt.Caption = TDSLimitAmt
    '--As Per Discussion with Arun Babu and Munole Sir On 06/08/2025 no tds up to 30 corod for mill code ZRA,ZRS,ZRP
    If (UCase(txtCode(5).text) = "ZRA" Or UCase(txtCode(5).text) = "ZRS" Or UCase(txtCode(5).text) = "ZRP" Or UCase(txtCode(5).text) = "RJS") And Format(dtpDt(10).Value, "YYYY/MM/dd") >= Format("2025/08/01", "YYYY/MM/dd") And gCYear = 2025 Then
        If TDSLimitAmt >= 300000000 And Val(txt(70)) = 0 Then
           txt(63) = 0.1
           If gCYear >= 2023 Then
              If TDSLimitAmt - Amt >= 300000000 Then
                 TdsOnAmt = AssValue
              Else
                 TdsOnAmt = TDSLimitAmt - 300000000
              End If
              If TdsOnAmt > AssValue Then
                 TdsOnAmt = AssValue
              End If
              TdsAmt = (TdsOnAmt * Val(txt(63))) / 100
           Else
               TdsAmt = (AssValue * Val(txt(63))) / 100
           End If
           TdsAmt = GProcMakeRounding(TdsAmt)
           TdsAmt = Format(TdsAmt, FStr)
        End If
    Else
        If TDSLimitAmt >= 5000000 And Val(txt(70)) = 0 Then
           txt(63) = 0.1
           If gCYear >= 2023 Then
              If TDSLimitAmt - Amt >= 5000000 Then
                 TdsOnAmt = AssValue
              Else
                 TdsOnAmt = TDSLimitAmt - 5000000
              End If
              If TdsOnAmt > AssValue Then
                 TdsOnAmt = AssValue
              End If
              TdsAmt = (TdsOnAmt * Val(txt(63))) / 100
           Else
               TdsAmt = (AssValue * Val(txt(63))) / 100
           End If
           TdsAmt = GProcMakeRounding(TdsAmt)
           TdsAmt = Format(TdsAmt, FStr)
        End If
    End If
    txtCode(65) = GProcGetColumnValue("tblMastAccount", "AcCode", "" & gTdsOnPurchCode & "", "N", "AcOurCode", "S")
    txttemp(65) = GProcGetColumnValue("tblMastAccount", "AcCode", "" & gTdsOnPurchCode & "", "N", "AcName", "S")
End If
If gCYear >= 2023 Then
   txt(72) = TdsOnAmt
End If
txt(64) = TdsAmt

End Sub
'--- Is Related Record
Private Function IsRelatedRecord() As Boolean
Dim RsG As Recordset
Dim RsG2 As Recordset
Dim Y As String
Dim BilNote As String
'-- Audited
'If GProcISAudited(txt(0), txt(1), txt(23)) = True Then
'    MsgBox "Record is Audited.You can not Modify/Delete this Record.", vbCritical + vbOKOnly, App.Title
'    IsRelatedRecord = True
'    Exit Function
'End If
'-- Is Related Payment
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
If frmTypeStr = "PE" Then
    BilNote = ""
    i = 0
    X = "Select PurSubItCtrlNo  from tblPurchSub where VType= '" & frmTypeStr & "' and " _
    & "  Vno = " & txt(0) & " and VFirm  = '" & gCCode & "' and  VYear= " & gCYear & ""
    GProcRstOpen RsG, X, "R"
    RsG.MoveFirst
    Do While Not RsG.EOF
       Y = "Select tblSale.SlBillNo,tblSale.vdt from tblSale,tblSaleSub where tblSale.vno = tblSaleSub.Vno and tblSale.Vtype = tblSaleSub.Vtype " _
           & " and tblSale.vyear = tblSaleSub.Vyear and tblSale.vfirm = tblSaleSub.Vfirm and tblSaleSub.SlSubPItCtrlNo =  " & RsG.Fields(0) & "  "
       If GProcRstOpen(RsG2, Y, "R") > 0 Then
          RsG2.MoveFirst
          Do While Not RsG2.EOF And i <= 10
             i = i + 1
             BilNote = BilNote + CStr(RsG2.Fields(0)) + "-" + CStr(RsG2.Fields(1)) + ", "
             RsG2.MoveNext
          Loop
       End If
       RsG.MoveNext
    Loop
    If BilNote <> "" Then
       IsRelatedRecord = True
       MsgBox "You can not Modify/Delete this record.Related Bill No. " & BilNote & " is exists.", vbCritical + vbOKOnly, App.Title
       Exit Function
    End If
End If
End Function
'-- Is Related GatePass
Private Function IsRelatedGatePass() As Boolean
Dim RsG As Recordset
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
Private Sub BookingData()
Dim RsG As Recordset
Dim WStdPack As Double
Dim wItcode As Long
Dim wMillCode As Long
Dim wParty As Long
Dim wBroker As Long
Dim BalBag As Integer
Dim BalWt As Double
Dim tmpRst As Recordset
Dim X As String
Set RsG = New Recordset
Dim i As Integer
wParty = GProcGetColumnValue("tblMastAccount", "AcName", txttemp(5), "S", "AcCode", "N")
If Val(msGrid.TextMatrix(msGrid.Row, 1)) <> 0 Then
'----
   gCmd.CommandText = "PrcPreparePendingBookingPurch"
   gCmd.Parameters.Refresh
   gCmd.Parameters("@Fr_dt") = Format(gCYSDate - 400, "yyyy/mm/dd")
   gCmd.Parameters("@To_dt") = Format(gCYEDate, "yyyy/mm/dd")
   gCmd.Parameters("@UCode") = gUserName
   gCmd.Parameters("@VFirm") = Trim(gCCode)
   gCmd.Parameters("@Vtype") = "OT"
   gCmd.Parameters("@BkNo") = Val(msGrid.TextMatrix(msGrid.Row, 1))
   gCmd.Parameters("@BkYear") = Val(msGrid.TextMatrix(msGrid.Row, 2))
   gCmd.Parameters("@RptTp") = "P"
   gCmd.Parameters("@BkSrNo") = Val(msGrid.TextMatrix(msGrid.Row, 13))
   gCmd.Execute
   X = "Select Vno,PartyName,BrokerName,MillCode,Nar1,Amt3,Amt4-Amt8,Dt1, " _
       & "PartyCode,BrokerCode,MillName,Amt6 ,Amt7,Vtype,Vyear,Amt5,Vdt,Dt2,format(Amt9,'#####'),format(Amt11,'#####') from TmpGentbl where TmpGentbl.UserName = '" & gUserName & "' and (amt4 - amt8) > 0   "

       i = GProcRstOpen(tmpRst, X, "R")
       If i > 0 Then
          gClsSearch.SearchMultiField X, "", Array("Booking No", "Party", "Broker", "Mill", "Count", "Rate", "Sh.Bag", "Sh.Date", "", "", "", "Std.Pck", "RtPer", "BkTp", "BkYear", "BkSrNo", "Bk.Dt", "Desp. Dt", "Desp. Bag", "BkDespSr"), Array(0, 0, 0, 0, 2000, 1000, 800, 1500, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0), "", "", "Vno,Vdt,dt1 ", 0, 0, False, , True
          If Val(gClsSearch.SearchMultiRetCol(0)) <> 0 Then
            If txttemp(5) <> "" And gClsSearch.SearchMultiRetCol(10) <> txttemp(5) Then
               MsgBox ("Please Check Booking Party ....")
               msGrid.Col = 1
               msGrid.SetFocus
               Exit Sub
            End If
'            If txttemp(8) <> "" And gClsSearch.SearchMultiRetCol(2) <> txttemp(8) Then
'               MsgBox ("Please Check Booking Broker ....")
'               msGrid.Col = 1
'               msGrid.SetFocus
'               Exit Sub
'            End If
'            txttemp(5) = gClsSearch.SearchMultiRetCol(1)
'            txttemp(8) = gClsSearch.SearchMultiRetCol(2)
            msGrid.TextMatrix(msGrid.Row, 2) = gClsSearch.SearchMultiRetCol(14) '-- Vyear
            msGrid.TextMatrix(msGrid.Row, 3) = gClsSearch.SearchMultiRetCol(4) '-- Item
            msGrid.TextMatrix(msGrid.Row, 13) = gClsSearch.SearchMultiRetCol(15) '-- BkSrNo
            If FormAction = vbDataActionAddNew Then
               msGrid.TextMatrix(msGrid.Row, 6) = gClsSearch.SearchMultiRetCol(6)
            End If
            wItcode = GProcGetColumnValue("tblMastItem", "ItName", "" & msGrid.TextMatrix(msGrid.Row, 3) & "", "S", "ItCode", "N", "ItMillCode = " & wParty & "")
            WStdPack = GProcGetColumnValue("tblMastItem", "Itcode", "" & wItcode & "", "N", "ItStdpack", "N")
            msGrid.TextMatrix(msGrid.Row, 7) = Format(gClsSearch.SearchMultiRetCol(6) * WStdPack, WtStr)
            msGrid.Col = 5
            txtGrid.Locked = False

            msGrid.SetFocus
        Else
            msGrid.TextMatrix(msGrid.Row, 0) = ""
        End If
   Else
      msGrid.TextMatrix(msGrid.Row, 2) = 0
      msGrid.TextMatrix(msGrid.Row, 3) = ""
      msGrid.TextMatrix(msGrid.Row, 13) = 0
      msGrid.TextMatrix(msGrid.Row, 6) = 0
      msGrid.TextMatrix(msGrid.Row, 7) = 0
      MsgBox ("Booking No Not Found....")
      msGrid.Col = 1
      msGrid.SetFocus
   End If
End If
End Sub

