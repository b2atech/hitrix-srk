VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form frmSalesGST 
   BackColor       =   &H00FBF2E1&
   BorderStyle     =   0  'None
   Caption         =   "Purchase Entry"
   ClientHeight    =   9300
   ClientLeft      =   0
   ClientTop       =   15
   ClientWidth     =   15630
   FillColor       =   &H000040C0&
   FillStyle       =   0  'Solid
   Icon            =   "frmSalesGST.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   9300
   ScaleWidth      =   15630
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame1 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFC0C0&
      Caption         =   "Sale Bill  Printing"
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
      Height          =   5835
      Left            =   3540
      TabIndex        =   137
      Top             =   1620
      Visible         =   0   'False
      Width           =   7485
      Begin VB.CheckBox chkIsDatewise 
         BackColor       =   &H00FFC0C0&
         Caption         =   "Is Datewise Printing ?"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   495
         Left            =   840
         TabIndex        =   198
         Top             =   1560
         Width           =   3075
      End
      Begin VB.CommandButton cmdPrintBill 
         Caption         =   "Whatsapp SMS PDF"
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
         Index           =   6
         Left            =   2580
         TabIndex        =   182
         Top             =   5100
         Width           =   2535
      End
      Begin VB.CommandButton cmdPrintBill 
         Caption         =   "Update IRN"
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
         Index           =   5
         Left            =   4800
         TabIndex        =   177
         Top             =   4380
         Width           =   2055
      End
      Begin VB.CommandButton cmdPrintBill 
         Caption         =   "e-Invoice - Screen"
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
         Left            =   360
         TabIndex        =   176
         Top             =   4380
         Width           =   2055
      End
      Begin VB.CommandButton cmdPrintBill 
         Caption         =   "&e-Invoice"
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
         Left            =   2580
         TabIndex        =   175
         Top             =   4380
         Width           =   2055
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
         Left            =   4080
         TabIndex        =   160
         Top             =   2940
         Width           =   1815
      End
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
         Left            =   2160
         TabIndex        =   159
         Top             =   2940
         Width           =   1575
      End
      Begin VB.CheckBox ChkGod 
         BackColor       =   &H00FFC0C0&
         Caption         =   "Godown Copy"
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
         Left            =   240
         TabIndex        =   158
         Top             =   2940
         Width           =   1575
      End
      Begin VB.CheckBox chkInvChallan 
         BackColor       =   &H00FFC0C0&
         Caption         =   "Delivery Challan"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   285
         Index           =   4
         Left            =   4320
         TabIndex        =   151
         Top             =   2340
         Width           =   3030
      End
      Begin VB.CheckBox chkInvChallan 
         BackColor       =   &H00FFC0C0&
         Caption         =   "Qudraplicate For Mill"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   285
         Index           =   3
         Left            =   4320
         TabIndex        =   150
         Top             =   1890
         Width           =   3030
      End
      Begin VB.CheckBox chkInvChallan 
         BackColor       =   &H00FFC0C0&
         Caption         =   "Triplicate For Accounts"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   285
         Index           =   2
         Left            =   4320
         TabIndex        =   149
         Top             =   1395
         Width           =   3030
      End
      Begin VB.CheckBox chkInvChallan 
         BackColor       =   &H00FFC0C0&
         Caption         =   "Duplicate For Transporter"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   285
         Index           =   1
         Left            =   4320
         TabIndex        =   148
         Top             =   900
         Width           =   3030
      End
      Begin VB.CheckBox chkInvChallan 
         BackColor       =   &H00FFC0C0&
         Caption         =   "Original For Buyer "
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   285
         Index           =   0
         Left            =   4320
         TabIndex        =   147
         Top             =   450
         Width           =   2445
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
         Left            =   4800
         TabIndex        =   142
         Top             =   3630
         Width           =   2055
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
         Left            =   2580
         TabIndex        =   141
         Top             =   3630
         Width           =   2055
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
         Left            =   1920
         MaxLength       =   10
         TabIndex        =   140
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
         Left            =   360
         TabIndex        =   139
         Top             =   3630
         Width           =   2055
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
         Left            =   1920
         MaxLength       =   10
         TabIndex        =   138
         Top             =   465
         Width           =   2055
      End
      Begin MSComCtl2.DTPicker dtpDtPrintBill 
         Height          =   360
         Left            =   1320
         TabIndex        =   200
         Top             =   2160
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
         Format          =   116916227
         CurrentDate     =   -273
      End
      Begin VB.Label Label33 
         BackStyle       =   0  'Transparent
         Caption         =   "Date       :    "
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
         Height          =   210
         Left            =   240
         TabIndex        =   199
         Top             =   2220
         Width           =   1320
      End
      Begin VB.Line Line2 
         X1              =   0
         X2              =   7680
         Y1              =   2760
         Y2              =   2760
      End
      Begin VB.Label Label39 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "To       Bill  No   :"
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
         Left            =   195
         TabIndex        =   144
         Top             =   1065
         Width           =   1590
      End
      Begin VB.Label Label38 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "From  Bill  No    :"
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
         Left            =   195
         TabIndex        =   143
         Top             =   555
         Width           =   1605
      End
      Begin VB.Line Line3 
         X1              =   0
         X2              =   7680
         Y1              =   3360
         Y2              =   3360
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
      Index           =   74
      Left            =   2220
      TabIndex        =   196
      Text            =   "74"
      Top             =   7140
      Width           =   1845
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
      Index           =   73
      Left            =   16200
      TabIndex        =   194
      Text            =   "73"
      Top             =   7440
      Visible         =   0   'False
      Width           =   480
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
      Index           =   72
      Left            =   16200
      TabIndex        =   193
      Text            =   "72"
      Top             =   7020
      Visible         =   0   'False
      Width           =   480
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
      Index           =   71
      Left            =   16200
      TabIndex        =   192
      Text            =   "71"
      Top             =   6600
      Visible         =   0   'False
      Width           =   480
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
      Index           =   70
      Left            =   16320
      TabIndex        =   191
      Text            =   "70"
      Top             =   6180
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
      Height          =   405
      Index           =   70
      Left            =   2220
      TabIndex        =   190
      Text            =   "txttemp(70)"
      ToolTipText     =   "Press F3 for Creation"
      Top             =   8640
      Width           =   4335
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
      Index           =   69
      Left            =   4680
      TabIndex        =   189
      Text            =   "69"
      Top             =   7680
      Width           =   1845
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
      Left            =   2220
      TabIndex        =   188
      Text            =   "txtCode(70)"
      ToolTipText     =   "Press F3 for Creation"
      Top             =   8160
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
      Height          =   405
      Index           =   68
      Left            =   2220
      TabIndex        =   187
      Text            =   "68"
      Top             =   7680
      Width           =   1005
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
      Left            =   10035
      TabIndex        =   179
      Text            =   "67"
      Top             =   2880
      Width           =   4005
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
      Height          =   585
      Index           =   66
      Left            =   10020
      MaxLength       =   64
      MultiLine       =   -1  'True
      TabIndex        =   178
      Text            =   "frmSalesGST.frx":058A
      Top             =   2160
      Width           =   4965
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
      Index           =   62
      Left            =   16680
      TabIndex        =   173
      Text            =   "62"
      Top             =   5760
      Visible         =   0   'False
      Width           =   480
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
      Index           =   61
      Left            =   16980
      TabIndex        =   172
      Text            =   "61"
      Top             =   5340
      Visible         =   0   'False
      Width           =   480
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
      Index           =   60
      Left            =   16680
      TabIndex        =   171
      Text            =   "60"
      Top             =   5280
      Visible         =   0   'False
      Width           =   480
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
      Left            =   60
      MaxLength       =   64
      TabIndex        =   47
      Text            =   "59"
      Top             =   6480
      Visible         =   0   'False
      Width           =   6840
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
      Index           =   65
      Left            =   10440
      TabIndex        =   35
      Text            =   "65"
      Top             =   7620
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
      Height          =   405
      Index           =   64
      Left            =   9540
      TabIndex        =   34
      Text            =   "64"
      Top             =   7620
      Width           =   645
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
      Left            =   7860
      TabIndex        =   33
      Text            =   "63"
      Top             =   7620
      Width           =   1665
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
      Left            =   4200
      TabIndex        =   161
      Top             =   720
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
         TabIndex        =   162
         Top             =   1320
         Width           =   1335
      End
      Begin VB.Label Label11 
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
         TabIndex        =   163
         Top             =   420
         Width           =   6495
      End
   End
   Begin VB.TextBox txtSlSub 
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
      Left            =   17730
      TabIndex        =   157
      Text            =   "txtSlSub(28)"
      ToolTipText     =   "Press F3 for Creation & Press F2 for Selection"
      Top             =   495
      Visible         =   0   'False
      Width           =   1710
   End
   Begin VB.TextBox txtSlSub 
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
      Index           =   20
      Left            =   15930
      TabIndex        =   156
      Text            =   "txtSlSub(20)"
      ToolTipText     =   "Press F3 for Creation & Press F2 for Selection"
      Top             =   540
      Visible         =   0   'False
      Width           =   1710
   End
   Begin VB.TextBox txtSlSub 
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
      Index           =   19
      Left            =   18045
      TabIndex        =   155
      Text            =   "txtSlSub(19)"
      ToolTipText     =   "Press F3 for Creation & Press F2 for Selection"
      Top             =   45
      Visible         =   0   'False
      Width           =   1710
   End
   Begin VB.TextBox txtSlSub 
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
      Left            =   16290
      TabIndex        =   154
      Text            =   "txtSlSub(18)"
      ToolTipText     =   "Press F3 for Creation & Press F2 for Selection"
      Top             =   90
      Visible         =   0   'False
      Width           =   1710
   End
   Begin VB.TextBox txtSlSub 
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
      Index           =   17
      Left            =   10035
      TabIndex        =   1
      Text            =   "txtSlSub(17)"
      ToolTipText     =   "Press F3 for Creation & Press F2 for Selection"
      Top             =   1335
      Width           =   1710
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
      Left            =   2325
      TabIndex        =   6
      Text            =   "txttemp(5)"
      ToolTipText     =   "Press F3 for Creation"
      Top             =   1245
      Width           =   5655
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
      Left            =   2325
      TabIndex        =   8
      Text            =   "txttemp(8)"
      ToolTipText     =   "Press F3 for Creation"
      Top             =   2400
      Width           =   5655
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
      Index           =   22
      Left            =   2325
      TabIndex        =   9
      Text            =   "txttemp(22)"
      ToolTipText     =   "Press F3 for Creation"
      Top             =   2925
      Width           =   5655
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
      Left            =   19080
      TabIndex        =   134
      Text            =   "47"
      Top             =   4995
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
      Index           =   46
      Left            =   19080
      TabIndex        =   133
      Text            =   "46"
      Top             =   4590
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
      Index           =   45
      Left            =   19035
      TabIndex        =   132
      Text            =   "45"
      Top             =   4140
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
      Height          =   405
      Index           =   8
      Left            =   1155
      TabIndex        =   7
      Text            =   "txtCode(8)"
      ToolTipText     =   "Press F3 for Creation"
      Top             =   2385
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
      Height          =   405
      Index           =   5
      Left            =   1155
      TabIndex        =   5
      Text            =   "txtCode(5)"
      ToolTipText     =   "Press F3 for Creation"
      Top             =   1245
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
      Index           =   7
      Left            =   1155
      TabIndex        =   3
      Text            =   "txtCode(7)"
      ToolTipText     =   "Press F3 for Creation"
      Top             =   735
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
      Height          =   375
      Index           =   44
      Left            =   18990
      TabIndex        =   126
      Text            =   "44"
      Top             =   3600
      Visible         =   0   'False
      Width           =   570
   End
   Begin VB.ComboBox Cbo 
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Index           =   44
      ItemData        =   "frmSalesGST.frx":058D
      Left            =   8865
      List            =   "frmSalesGST.frx":0597
      TabIndex        =   24
      Text            =   "Cbo(44)"
      Top             =   5256
      Visible         =   0   'False
      Width           =   1290
   End
   Begin VB.TextBox txt 
      Alignment       =   1  'Right Justify
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Index           =   43
      Left            =   8055
      TabIndex        =   23
      Text            =   "43"
      Top             =   5256
      Visible         =   0   'False
      Width           =   690
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
      Index           =   42
      Left            =   10440
      TabIndex        =   32
      Text            =   "42"
      Top             =   6990
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
      Height          =   405
      Index           =   41
      Left            =   8160
      TabIndex        =   29
      Text            =   "41"
      Top             =   6983
      Width           =   690
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
      Left            =   19080
      TabIndex        =   115
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
      Left            =   19035
      TabIndex        =   114
      Text            =   "39"
      Top             =   765
      Visible         =   0   'False
      Width           =   570
   End
   Begin VB.TextBox txt 
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
      Index           =   38
      Left            =   16695
      TabIndex        =   13
      Text            =   "38"
      Top             =   4950
      Visible         =   0   'False
      Width           =   435
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
      Left            =   10035
      TabIndex        =   10
      Text            =   "9"
      Top             =   735
      Width           =   1710
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
      TabIndex        =   113
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
      TabIndex        =   112
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
      Left            =   14325
      Style           =   1  'Graphical
      TabIndex        =   110
      Top             =   6767
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
      TabIndex        =   109
      Text            =   "35"
      Top             =   7200
      Visible         =   0   'False
      Width           =   570
   End
   Begin VB.TextBox txttemp 
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
      Height          =   360
      Index           =   35
      Left            =   18180
      MaxLength       =   5
      TabIndex        =   44
      Text            =   "txtemp(35)"
      Top             =   5715
      Visible         =   0   'False
      Width           =   1545
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
      TabIndex        =   103
      Text            =   "22"
      Top             =   4365
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
      Height          =   435
      Index           =   16
      Left            =   9060
      TabIndex        =   37
      Text            =   "16"
      Top             =   8520
      Width           =   3615
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
      Left            =   10440
      TabIndex        =   22
      Text            =   "15"
      Top             =   4860
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
      Left            =   18360
      TabIndex        =   101
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
      Left            =   17325
      TabIndex        =   18
      Text            =   "13"
      Top             =   2790
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
      Left            =   17280
      TabIndex        =   16
      Text            =   "txttemp(11)"
      ToolTipText     =   "Press F3 for Creation & Press F2 for Selection"
      Top             =   1800
      Visible         =   0   'False
      Width           =   1290
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
      Left            =   18900
      TabIndex        =   42
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
      TabIndex        =   43
      Text            =   "29"
      Top             =   3105
      Visible         =   0   'False
      Width           =   330
   End
   Begin VB.CommandButton CmdBagDetails 
      Caption         =   "XXX"
      Height          =   720
      Left            =   13200
      TabIndex        =   94
      Top             =   9840
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
      TabIndex        =   54
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
      TabIndex        =   93
      Text            =   "34"
      Top             =   6750
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
      Left            =   17730
      TabIndex        =   41
      Text            =   "19"
      Top             =   8685
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
      Height          =   375
      Index           =   33
      Left            =   18540
      TabIndex        =   92
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
      TabIndex        =   91
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
      Index           =   7
      Left            =   2325
      TabIndex        =   4
      Text            =   "txttemp(7)"
      ToolTipText     =   "Press F3 for Creation"
      Top             =   735
      Width           =   5655
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
      Left            =   16965
      TabIndex        =   39
      Text            =   "txtDrBalance"
      Top             =   1170
      Width           =   300
   End
   Begin VB.CommandButton cmdLedger 
      Caption         =   "Ledger"
      Height          =   375
      Left            =   12915
      TabIndex        =   90
      Top             =   10980
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
      Left            =   10440
      TabIndex        =   31
      Text            =   "27"
      Top             =   6564
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
      Height          =   405
      Index           =   26
      Left            =   8160
      TabIndex        =   28
      Text            =   "26"
      Top             =   6557
      Width           =   690
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
      TabIndex        =   66
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
      Left            =   10440
      TabIndex        =   26
      Text            =   "21"
      Top             =   5712
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
      TabIndex        =   21
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
      TabIndex        =   88
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
      TabIndex        =   86
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
      TabIndex        =   85
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
      TabIndex        =   84
      Text            =   "3"
      Top             =   1305
      Visible         =   0   'False
      Width           =   525
   End
   Begin VB.CommandButton cmdBtn 
      Caption         =   "Change GST Amt."
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
      Left            =   5040
      Style           =   1  'Graphical
      TabIndex        =   80
      Top             =   5700
      Width           =   1815
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
      TabIndex        =   79
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
      Left            =   14325
      Style           =   1  'Graphical
      TabIndex        =   77
      ToolTipText     =   "Exit"
      Top             =   7950
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
      Left            =   13365
      Style           =   1  'Graphical
      TabIndex        =   70
      ToolTipText     =   "To Save Record"
      Top             =   7950
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
      Left            =   14325
      Style           =   1  'Graphical
      TabIndex        =   78
      Top             =   7365
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
      Height          =   400
      Index           =   8
      Left            =   13365
      Style           =   1  'Graphical
      TabIndex        =   69
      ToolTipText     =   "To Delete Record"
      Top             =   7365
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
      Left            =   13365
      Style           =   1  'Graphical
      TabIndex        =   68
      ToolTipText     =   "To Modify Record"
      Top             =   6765
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
      Left            =   13365
      Style           =   1  'Graphical
      TabIndex        =   67
      ToolTipText     =   "To Add New Record"
      Top             =   6165
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
      Left            =   18360
      Style           =   1  'Graphical
      TabIndex        =   71
      Top             =   3915
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
      Left            =   13365
      Style           =   1  'Graphical
      TabIndex        =   72
      ToolTipText     =   "To Find Record"
      Top             =   5520
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
      Left            =   14820
      Style           =   1  'Graphical
      TabIndex        =   76
      ToolTipText     =   "Last"
      Top             =   6165
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
      Left            =   14325
      Style           =   1  'Graphical
      TabIndex        =   75
      ToolTipText     =   "First"
      Top             =   6165
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
      Left            =   14820
      Style           =   1  'Graphical
      TabIndex        =   74
      ToolTipText     =   "Next"
      Top             =   5520
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
      TabIndex        =   38
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
      Left            =   14325
      Style           =   1  'Graphical
      TabIndex        =   73
      ToolTipText     =   "Previous"
      Top             =   5520
      Width           =   375
   End
   Begin VB.TextBox txt 
      Alignment       =   1  'Right Justify
      Height          =   315
      Index           =   4
      Left            =   18270
      TabIndex        =   83
      Text            =   "4"
      Top             =   1575
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
      Left            =   1455
      Locked          =   -1  'True
      TabIndex        =   48
      Text            =   "txttemp(6)"
      ToolTipText     =   "Press F3 for Creation"
      Top             =   5265
      Width           =   4740
   End
   Begin VB.TextBox txt 
      Height          =   315
      Index           =   6
      Left            =   18405
      TabIndex        =   82
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
      TabIndex        =   81
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
      TabIndex        =   56
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
      TabIndex        =   64
      Text            =   "11"
      Top             =   3555
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
      Height          =   405
      Index           =   12
      Left            =   17235
      TabIndex        =   17
      Text            =   "12"
      Top             =   2250
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
      Height          =   390
      Index           =   23
      Left            =   8160
      TabIndex        =   27
      Text            =   "23"
      Top             =   6138
      Width           =   690
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
      Left            =   10440
      TabIndex        =   30
      Text            =   "24"
      Top             =   6138
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
      Left            =   18000
      TabIndex        =   50
      Text            =   "20"
      Top             =   9000
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
      Index           =   30
      Left            =   10440
      TabIndex        =   36
      Text            =   "30"
      Top             =   8100
      Width           =   2265
   End
   Begin VB.TextBox txt 
      Alignment       =   1  'Right Justify
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Index           =   31
      Left            =   1455
      TabIndex        =   45
      Text            =   "txt(31)"
      Top             =   5670
      Width           =   795
   End
   Begin VB.TextBox txt 
      Alignment       =   1  'Right Justify
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Index           =   17
      Left            =   10440
      TabIndex        =   25
      Text            =   "17"
      Top             =   5256
      Visible         =   0   'False
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
      Index           =   18
      Left            =   18945
      TabIndex        =   40
      Text            =   "18"
      Top             =   2295
      Visible         =   0   'False
      Width           =   330
   End
   Begin MSMask.MaskEdBox mskDt 
      Height          =   390
      Index           =   2
      Left            =   16245
      TabIndex        =   12
      Top             =   3420
      Visible         =   0   'False
      Width           =   1740
      _ExtentX        =   3069
      _ExtentY        =   688
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
      Height          =   390
      Index           =   39
      Left            =   16200
      TabIndex        =   15
      Top             =   3915
      Visible         =   0   'False
      Width           =   1740
      _ExtentX        =   3069
      _ExtentY        =   688
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
      Left            =   16290
      TabIndex        =   49
      Top             =   4455
      Visible         =   0   'False
      Width           =   1470
      _ExtentX        =   2593
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
      Bindings        =   "frmSalesGST.frx":05A4
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
      Height          =   1140
      Left            =   180
      TabIndex        =   20
      Top             =   3540
      Width           =   15270
      _ExtentX        =   26935
      _ExtentY        =   2011
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
      Left            =   17910
      TabIndex        =   19
      Top             =   2655
      Visible         =   0   'False
      Width           =   480
      _ExtentX        =   847
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
      Left            =   18945
      TabIndex        =   116
      Top             =   1665
      Visible         =   0   'False
      Width           =   795
      _ExtentX        =   1402
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
   Begin MSComCtl2.DTPicker dtpDt 
      Height          =   360
      Index           =   2
      Left            =   13215
      TabIndex        =   11
      Top             =   735
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
      Format          =   116916227
      CurrentDate     =   -273
   End
   Begin MSComCtl2.DTPicker dtpDt 
      Height          =   360
      Index           =   39
      Left            =   17190
      TabIndex        =   14
      Top             =   4950
      Visible         =   0   'False
      Width           =   1710
      _ExtentX        =   3016
      _ExtentY        =   635
      _Version        =   393216
      Enabled         =   0   'False
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
      Format          =   116916227
      CurrentDate     =   -273
   End
   Begin MSComCtl2.DTPicker dtpDt 
      Height          =   360
      Index           =   32
      Left            =   3195
      TabIndex        =   46
      Top             =   5715
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
      Format          =   116916227
      CurrentDate     =   -273
   End
   Begin MSMask.MaskEdBox mskDtSlSub 
      Height          =   390
      Index           =   19
      Left            =   13215
      TabIndex        =   2
      Top             =   1335
      Width           =   1740
      _ExtentX        =   3069
      _ExtentY        =   688
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
   Begin Crystal.CrystalReport CrystalReport1 
      Bindings        =   "frmSalesGST.frx":05B5
      Left            =   0
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
   Begin VB.Label Label32 
      BackStyle       =   0  'Transparent
      Caption         =   "TCS On : "
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
      Left            =   7020
      TabIndex        =   197
      Top             =   7680
      Width           =   1065
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
      Left            =   480
      TabIndex        =   195
      Top             =   7200
      Width           =   1890
   End
   Begin VB.Label Label37 
      BackStyle       =   0  'Transparent
      Caption         =   "TDS A/c             :"
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
      Left            =   480
      TabIndex        =   186
      Top             =   8700
      Width           =   1890
   End
   Begin VB.Label Label36 
      BackStyle       =   0  'Transparent
      Caption         =   "TDS A/c  Code   :"
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
      Left            =   480
      TabIndex        =   185
      Top             =   8220
      Width           =   1890
   End
   Begin VB.Label Label35 
      BackStyle       =   0  'Transparent
      Caption         =   "TDS Amt.  :    "
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
      Left            =   3480
      TabIndex        =   184
      Top             =   7740
      Width           =   1890
   End
   Begin VB.Label Label34 
      BackStyle       =   0  'Transparent
      Caption         =   "TDS Rate           :"
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
      Left            =   480
      TabIndex        =   183
      Top             =   7740
      Width           =   1890
   End
   Begin VB.Label Label29 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Ack. No.       :"
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
      Left            =   8700
      TabIndex        =   181
      Top             =   2940
      Width           =   1260
   End
   Begin VB.Label Label27 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "IRN               :"
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
      Left            =   8640
      TabIndex        =   180
      Top             =   2160
      Width           =   1290
   End
   Begin VB.Label Label26 
      BackStyle       =   0  'Transparent
      Caption         =   "e-Invoice No"
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
      Left            =   240
      TabIndex        =   174
      Top             =   6180
      Visible         =   0   'False
      Width           =   1320
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
      Left            =   180
      TabIndex        =   170
      Top             =   4740
      Width           =   6195
   End
   Begin VB.Label Label24 
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
      Left            =   10200
      TabIndex        =   169
      Top             =   7695
      Width           =   255
   End
   Begin VB.Label Label23 
      BackStyle       =   0  'Transparent
      Caption         =   "@  : "
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
      Left            =   7320
      TabIndex        =   168
      Top             =   8220
      Width           =   405
   End
   Begin VB.Label Label22 
      BackStyle       =   0  'Transparent
      Caption         =   "TCS On Amount          :"
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
      Left            =   13560
      TabIndex        =   167
      Top             =   8880
      Width           =   765
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00C00000&
      X1              =   6900
      X2              =   13035
      Y1              =   7500
      Y2              =   7500
   End
   Begin VB.Label Label20 
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
      Left            =   120
      TabIndex        =   166
      Top             =   5280
      Width           =   1620
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
      Left            =   3360
      TabIndex        =   165
      Top             =   2040
      Width           =   4020
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
      Left            =   2310
      TabIndex        =   164
      Top             =   2040
      Width           =   975
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Date       :    "
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
      Height          =   210
      Left            =   12060
      TabIndex        =   153
      Top             =   1402
      Width           =   1320
   End
   Begin VB.Label Label12 
      BackStyle       =   0  'Transparent
      Caption         =   "Consi. :"
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
      Left            =   330
      TabIndex        =   152
      Top             =   3015
      Width           =   1035
   End
   Begin VB.Label lblAdd2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Add Line2"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   17490
      TabIndex        =   146
      Top             =   990
      Visible         =   0   'False
      Width           =   960
   End
   Begin VB.Label lblAdd1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Add Line1 "
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   17490
      TabIndex        =   145
      Top             =   585
      Visible         =   0   'False
      Width           =   1020
   End
   Begin VB.Label Label18 
      BackStyle       =   0  'Transparent
      Caption         =   "Exempt/Kg   :"
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
      Left            =   16560
      TabIndex        =   136
      Top             =   8955
      Visible         =   0   'False
      Width           =   1305
   End
   Begin VB.Label Label16 
      BackStyle       =   0  'Transparent
      Caption         =   "Exempt.Amt"
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
      Left            =   16875
      TabIndex        =   135
      Top             =   8415
      Visible         =   0   'False
      Width           =   1530
   End
   Begin VB.Shape Shape3 
      BorderColor     =   &H00EFB17A&
      BorderWidth     =   5
      FillColor       =   &H00EFB17A&
      Height          =   9480
      Left            =   0
      Top             =   -195
      Width           =   15585
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "Stock "
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
      Left            =   450
      TabIndex        =   131
      Top             =   9465
      Visible         =   0   'False
      Width           =   2310
   End
   Begin VB.Label lblBag 
      BackStyle       =   0  'Transparent
      Caption         =   " Bag             : "
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
      Left            =   405
      TabIndex        =   130
      Top             =   9915
      Visible         =   0   'False
      Width           =   1590
   End
   Begin VB.Label lblWeight 
      BackStyle       =   0  'Transparent
      Caption         =   "Wt.               : "
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
      Left            =   450
      TabIndex        =   129
      Top             =   10275
      Visible         =   0   'False
      Width           =   1590
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
      Left            =   2160
      TabIndex        =   128
      Top             =   9915
      Visible         =   0   'False
      Width           =   2715
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
      Left            =   2160
      TabIndex        =   127
      Top             =   10275
      Visible         =   0   'False
      Width           =   2805
   End
   Begin VB.Label Label25 
      BackStyle       =   0  'Transparent
      Caption         =   " :"
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
      Left            =   10215
      TabIndex        =   125
      Top             =   5250
      Width           =   255
   End
   Begin VB.Label lblPer3 
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
      Left            =   8850
      TabIndex        =   124
      Top             =   6998
      Width           =   255
   End
   Begin VB.Label lblPer2 
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
      Left            =   8850
      TabIndex        =   123
      Top             =   6572
      Width           =   255
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
      Left            =   9120
      TabIndex        =   122
      Top             =   6146
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
      Left            =   9120
      TabIndex        =   121
      Top             =   6572
      Width           =   1335
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
      Left            =   9120
      TabIndex        =   120
      Top             =   6998
      Width           =   1335
   End
   Begin VB.Label lblSgst 
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
      Left            =   7050
      TabIndex        =   119
      Top             =   6150
      Width           =   1065
   End
   Begin VB.Label lblCgst 
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
      Left            =   7050
      TabIndex        =   118
      Top             =   6570
      Width           =   1065
   End
   Begin VB.Label lblIgst 
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
      Left            =   7050
      TabIndex        =   117
      Top             =   7005
      Width           =   1065
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
      Left            =   7185
      TabIndex        =   111
      Top             =   1680
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
      Left            =   3375
      TabIndex        =   108
      Top             =   1680
      Width           =   3090
   End
   Begin VB.Label Label21 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "GSTIN   :  "
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
      Left            =   2310
      TabIndex        =   107
      Top             =   1680
      Width           =   960
   End
   Begin VB.Label lblNature 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Nature"
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
      Height          =   240
      Left            =   16920
      TabIndex        =   106
      Top             =   6750
      Width           =   4485
   End
   Begin VB.Label lblDescription 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Description"
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
      Height          =   240
      Left            =   16920
      TabIndex        =   105
      Top             =   7155
      Width           =   4485
   End
   Begin VB.Label lbl1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Tr.  Code          : "
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C000C0&
      Height          =   225
      Index           =   0
      Left            =   16920
      TabIndex        =   104
      Top             =   6345
      Width           =   1290
   End
   Begin VB.Label Label19 
      BackStyle       =   0  'Transparent
      Caption         =   "Date :    "
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
      Left            =   2430
      TabIndex        =   102
      Top             =   5760
      Width           =   870
   End
   Begin VB.Label Label14 
      BackStyle       =   0  'Transparent
      Caption         =   "L.R. No        :    "
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
      Left            =   15930
      TabIndex        =   100
      Top             =   2925
      Visible         =   0   'False
      Width           =   1320
   End
   Begin VB.Label Label13 
      BackStyle       =   0  'Transparent
      Caption         =   "Date       :    "
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
      Height          =   210
      Left            =   12060
      TabIndex        =   99
      Top             =   780
      Width           =   1320
   End
   Begin VB.Label Label8 
      BackStyle       =   0  'Transparent
      Caption         =   "Transport      :    "
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
      Left            =   15930
      TabIndex        =   98
      Top             =   1845
      Visible         =   0   'False
      Width           =   1320
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "Booking No  :    "
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
      Left            =   8640
      TabIndex        =   97
      Top             =   1380
      Width           =   1365
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Bill No         :"
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
      Left            =   8640
      TabIndex        =   96
      Top             =   765
      Width           =   1230
   End
   Begin VB.Label Label15 
      BackStyle       =   0  'Transparent
      Caption         =   "D&ue Days     :"
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
      Left            =   180
      TabIndex        =   95
      Top             =   5760
      Width           =   1320
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
      Left            =   330
      TabIndex        =   51
      Top             =   1275
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
      Left            =   330
      TabIndex        =   53
      Top             =   780
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
      Left            =   330
      TabIndex        =   52
      Top             =   2385
      Width           =   870
   End
   Begin VB.Label Label9 
      BackStyle       =   0  'Transparent
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
      Left            =   180
      TabIndex        =   0
      Top             =   5655
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
      Left            =   18450
      TabIndex        =   65
      Top             =   8505
      Width           =   1335
   End
   Begin VB.Label Label10 
      BackStyle       =   0  'Transparent
      Caption         =   "Sub Total     :"
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
      Left            =   9075
      TabIndex        =   55
      Top             =   4890
      Width           =   1425
   End
   Begin VB.Label lblCharity 
      BackStyle       =   0  'Transparent
      Caption         =   "Charity (GF)  :"
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
      Left            =   6690
      TabIndex        =   57
      Top             =   5279
      Visible         =   0   'False
      Width           =   1380
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
      TabIndex        =   58
      Top             =   9585
      Visible         =   0   'False
      Width           =   750
   End
   Begin VB.Label Label30 
      BackStyle       =   0  'Transparent
      Caption         =   "Total Rs        :"
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
      Left            =   7560
      TabIndex        =   63
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
      Left            =   8595
      TabIndex        =   60
      Top             =   5757
      Width           =   1845
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
      Left            =   12585
      TabIndex        =   89
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
      Left            =   9030
      TabIndex        =   62
      Top             =   8138
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
      Left            =   18270
      TabIndex        =   59
      Top             =   8190
      Width           =   1440
   End
   Begin VB.Label LblSaleType 
      BackColor       =   &H00EFB17A&
      Caption         =   "     Sales"
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
      TabIndex        =   87
      Top             =   0
      Width           =   15615
   End
   Begin VB.Shape Shape1 
      FillColor       =   &H00FED8FC&
      FillStyle       =   0  'Solid
      Height          =   3720
      Left            =   13170
      Shape           =   4  'Rounded Rectangle
      Top             =   5040
      Width           =   2250
   End
   Begin VB.Label lblPer1 
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
      Left            =   8850
      TabIndex        =   61
      Top             =   6146
      Width           =   255
   End
   Begin VB.Shape Shape4 
      FillColor       =   &H00ECFFFF&
      FillStyle       =   0  'Solid
      Height          =   2145
      Left            =   270
      Top             =   6990
      Visible         =   0   'False
      Width           =   6420
   End
   Begin VB.Shape Shape6 
      FillColor       =   &H00EDFDFE&
      FillStyle       =   0  'Solid
      Height          =   2940
      Left            =   8325
      Top             =   540
      Width           =   7125
   End
   Begin VB.Shape Shape5 
      FillColor       =   &H00EDFDFE&
      FillStyle       =   0  'Solid
      Height          =   4365
      Left            =   6915
      Top             =   4740
      Width           =   6090
   End
   Begin VB.Shape Shape2 
      FillColor       =   &H00EDFDFE&
      FillStyle       =   0  'Solid
      Height          =   2940
      Left            =   180
      Top             =   540
      Width           =   7965
   End
End
Attribute VB_Name = "frmSalesGST"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Const tblName As String = "tblSale"
Const IdField As String = "Vno"
Const OrderField As String = "Vdt,Vno"
'Const MaxNo As Long = 44
Const MaxNo As Long = 47
Const OtherTxtFromNo As Long = 59 '-- Other Txt Ctrl From No
Const OtherTxtToNo As Long = 74 '-- Other Txt Ctrl Less No
Dim rstbl As Recordset '-- TblSale
Dim rstblSub As Recordset '-- TblSaleSub
Dim varBookMark As Variant
Dim FirstTimeFlag As Boolean
Dim OldgFrmTypeStr As String
Dim FormAction As Integer
Dim frmTypeStr As String
Dim mLastEntryDate As Date
Dim mLastEntryMill As String
Dim ChgRoff As Integer
Dim i As Long
Dim SalesAccode As Long
Dim Criteria As String
Dim ItemType As Integer
Dim OldBag As Integer
Dim OldCount As String
Dim TrnCd As String
Dim tmpRst As Recordset
Dim Mon As Integer
'-- Cartoon Bag Detail Entry
Private Sub CmdBagDetails_Click()
'gfrmBagInwVNo = txt(0)
'gfrmBagInwVType = frmTypeStr
'frmBagEntry.Show
End Sub
'Dim GridFlag As Boolean
Private Sub cmdBtn_Click(Index As Integer)
Dim X As String
ChgRoff = 0
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
'        If GProcRstOpen(rstbl, "Select * from " & tblName & " where " & Criteria & "  and  " & IdField & " = (Select max(" & IdField & ") from " & tblName & " where " & IdField & " <> '' and month(vdt) = " & Month(Date) & " and  " & Criteria & ") and " & Criteria, "R") > 0 Then ReadFields Else GProcClearForm Me, rstbl, MaxNo
        If GProcRstOpen(rstbl, "Select * from " & tblName & " where " & Criteria & "  and  " & IdField & " = (Select max(" & IdField & ") from " & tblName & " where vdt = (SELECT MAX(vdt) from tblSale where " & Criteria & " )  and  " & IdField & " <> 0 and  " & Criteria & ")", "R") > 0 Then ReadFields Else GProcClearForm Me, rstbl, MaxNo
Case 4 'Find ' QryMastAccount
    If gCYear < 2018 Then
       gClsSearch.SearchMultiField "tblSale,TblMastAccount", "vno,SlBillNo,Vdt,AcName", Array("Vno", "Bill No", "Vdt", "Party"), Array(0, txt(9).Width, 1500, 4000), " tblSale.SlAcDrCode=TblMastAccount.AcCode and vtype='" & frmTypeStr & "'   and VFirm = '" & gCCode & "'   and VYear=" & gCYear, "", "SlBillNo", 0, 0
       If gClsSearch.SearchMultiRetCol(0) <> "" Then
          GProcRstOpen rstbl, "Select * from " & tblName & " where " & IdField & " = " & gClsSearch.SearchMultiRetCol(0) & " and Vtype='" & frmTypeStr & "' and VFirm  = '" & gCCode & "' and VYear=" & gCYear & "     ", "R"
          gClsSearch.SearchMultiRetCol(1) = ""
          ReadFields
       End If
    Else
       gClsSearch.SearchMultiField "tblSale,TblMastAccount", "vno,SlBillNo,Vdt,AcName", Array("Vno", "Bill No", "Vdt", "Party"), Array(0, txt(9).Width, 1500, 4000), " tblSale.SlAcDrCode=TblMastAccount.AcCode and vtype='" & frmTypeStr & "'   and VFirm = '" & gCCode & "'   and VYear=" & gCYear, "", "(cast(substring(SlBillNo,5,15) as int) ) desc", 0, 0
       If gClsSearch.SearchMultiRetCol(0) <> "" Then
          GProcRstOpen rstbl, "Select * from " & tblName & " where " & IdField & " = " & gClsSearch.SearchMultiRetCol(0) & " and " & Criteria, "R"
          gClsSearch.SearchMultiRetCol(1) = ""
          ReadFields
       End If
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
        txtCode(7) = mLastEntryMill
    End If
    If FirstTimeFlag = True Then
        FirstTimeFlag = False
    Else
       txtSlSub(17).SetFocus
'       If txtCode(7).Enabled = True Then
'          txtCode(7).SetFocus
'       Else
'         txtCode(5).SetFocus
'       End If
    End If
    '-- Sales A/C
    txtTemp(6) = GProcGetColumnValue("tblMastAccount", "AcCode", CStr(SalesAccode), "N", "AcName", "S")
    txtTemp(35) = GProcGetColumnValue("tblMastTransactionCode", "TranCode", CStr(TrnCd), "S", "TranCode", "S")
    lblNature.Caption = GProcGetColumnValue("tblMastTransactionCode", "TranCode", txtTemp(35), "S", "Nature", "S")
    lblDescription.Caption = GProcGetColumnValue("tblMastTransactionCode", "TranCode", txtTemp(35), "S", "Description", "S")
    If frmTypeStr = "SY" Or frmTypeStr = "SE" Then
       GBillNo
    End If
    FillCombo
    SetGrid
    lblTinNo.Caption = ""
    lblMsOms.Caption = ""
    lblBalBag.Caption = ""
    lblBalWt.Caption = ""
    txt(31) = 20
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
            txt(9).Enabled = False
            txtCode(7).SetFocus
        End If
        txtTemp(8).Enabled = True
        txtCode(8).Enabled = True
        txt(31).Enabled = True
        SetControlEd
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
    Frame1.Caption = "Sales Bill Printing "
    '-- Inv No
    txtPInvno(0) = txt(9).text
    txtPInvno(1) = txt(9).text
    Frame1.Visible = True
    DoEvents
    txtPInvno(0).Enabled = True
    txtPInvno(1).Enabled = True
    chkInvChallan(0).Enabled = True
    chkInvChallan(1).Enabled = True
    chkInvChallan(2).Enabled = True
    chkInvChallan(3).Enabled = True
    chkInvChallan(4).Enabled = True
    ChkCus.Enabled = True
    chkOff.Enabled = True
    ChkGod.Enabled = True
    chkInvChallan(4).Value = 1
    ChkCus.Value = 1
    txtPInvno(0).SetFocus
    txtPInvno(0).SelStart = Len(txtPInvno(0))
    chkIsDatewise.Enabled = True
    dtpDtPrintBill.Enabled = True
    dtpDtPrintBill = Format(Now - 1, "dd/MM/yyyy")
Case 10 'Exit
'    DoEvents
    Unload Me
    Exit Sub
Case 11 'Label
    'frmLabel.Show 1
Case 12 'Optional
        If IsRelatedRecord = False Then
            GProcRstOpen rstbl, "Select * from " & tblName & " where " & IdField & " = " & txt(0) & " and " & Criteria, "R"
            ReadFields
            FormAction = vbDataActionUpdate
            ChgRoff = 1
            Call GProcActivateControls(Me, True)
            SetControlEd
            Call GProcSetButtons(Me, vbDataActionUpdate)
            txt(0).Enabled = False
            txt(9).Enabled = False
            txt(30).Enabled = True
            txtCode(7).SetFocus
        End If
Case 13 'Save
     lblTcsLimit.Caption = ""
  If gCYear < 2018 Then
     FillTxtFromTemp
     gCn.Execute "update TblSale set SlBillAmt = " & Val(txt(16)) & " where Vtype = '" & txt(1) & "' and Vyear = " & txt(3) & " and Vfirm = '" & txt(4) & "' and vno=" & txt(0)
     cmdBtn(13).Enabled = False
     cmdBtn_Click (14)
  Else
     If Trim(txtGrid) <> "" Then
        msGrid.text = txtGrid
        DoEvents
        txtGrid = ""
     End If
     txtGrid.Visible = False
     CalAmount
     FillTxtFromTemp
     DoEvents
     DoEvents
     '-- Save Record
     If ValidateData = True Then
        If FormAction = vbDataActionAddNew Then
            Mon = Month(dtpDt(2))
            txt(0) = GProcGenerateIdMonthwise("TblSale", "Vno", Mon, "Vdt", "Vtype='" & frmTypeStr & "' and VFirm = '" & gCCode & "'   and VYear=" & gCYear)
        End If
        txt(46) = GProcGetColumnValue("tblMastAccount", "AcName", "-", "S", "AcCode", "N")
        txt(72) = GProcGetColumnValue("tblMastNarration", "Narration", msGrid.TextMatrix(1, 4), "S", "NarrCode", "N", "")
        
        
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
        FormAction = vbDataActionClose
        Call GProcActivateControls(Me, False)
        Call GProcSetButtons(Me, True)
        msGrid.Col = 0 '-----
        msGrid.Row = 1 '-----
        txtGrid.Move msGrid.CellLeft + msGrid.Left, msGrid.CellTop + msGrid.Top, msGrid.CellWidth, msGrid.CellHeight '-----
        If cmdBtn(6).Visible = True Then cmdBtn(6).SetFocus
        ReadFields
        FormAction = vbDataActionClose
        cmdBtn(13).Enabled = False
'-- For WhatsApp SMS
'''        Dim RsSA As Recordset
'''        Dim aa As Boolean
'''        aa = False
'''        If MsgBox("Do You Want To Send SMS ? ", vbYesNo) = vbYes Then aa = True
'''
'''        If aa Then
'''
'''           GProcCrystalRptPreparation cryReport, "Window"
'''           X = "Select a.Vno,a.Vtype,a.Vdt,a.Vyear,a.VFirm,a.SlAcDrCode,a.SlAcCrCode,a.SlMillCode,a.SlBroker,a.SlBillNo,a.SlBillDt,a.SlTransport,a.SlLorryNo,a.SlLrNo,a.SlLrDate, " _
'''               & "a.SlSubAmt,a.SlBillAmt,a.SlAdd1,a.SlLess1,a.SlExemptAmt,a.SlExemptPerKg,a.SlTaxableAmt,a.SlTaxCode,a.SlTaxRate,a.SlTaxAmt,a.SlTaxableAmt2,a.SlTaxRate2,a.SlTaxAmt2, " _
'''               & " a.SlAftTaxAdd1,a.SlAftTaxLess1,a.SlRoff,a.DueDays,a.DueDate,a.SlEntDt,'" & gUserName & "',a.SlTranCd,a.SlIsHank,a.GpVno,a.GpNo,a.GpDate,a.GpVYear,a.SlTaxRate3, " _
'''               & " a.SlTaxAmt3,a.SlCharityRt,a.SlCharityOn,a.SlCrDays,a.SlTaxPayAc,a.SlIsExempt,b.SlSubItCode,b.SlSubBag,b.SlSubWt,b.SlSubRt,b.SlSubNetRate,b.SlSubRtPer,b.SlSubAmt, " _
'''               & " b.SlSubBookNo,b.SlSubBookDt,b.SlSubGodown,b.SlSubLotNo,a.SlTcsOnAmt,a.SlTcsRate,a.SlTcsAmt,a.SlTmpChr3,a.SlIRNNo,a.SlAckNo from tblSale A ,tblSaleSub B" _
'''               & " where a.vno = b.vno and a.vtype = b.vtype and a.vyear = b.vyear and a.vfirm = b.vfirm and  a.SlBillNo = '" & txt(9).text & "' and a.Vtype='" & frmTypeStr & "' and a.VFirm = '" & gCCode & "' And a.VYear = " & gCYear
'''
'''           gCn.Execute "Delete from tmpSale where  SlUser='" & gUserName & "'"
'''
'''           gCn.Execute " insert into tmpSale (Vno,Vtype,Vdt,Vyear,VFirm,SlAcDrCode,SlAcCrCode,SlMillCode,SlBroker,SlBillNo,SlBillDt,SlTransport,SlLorryNo,SlLrNo,SlLrDate, " _
'''                        & "SlSubAmt,SlBillAmt,SlAdd1,SlLess1,SlExemptAmt,SlExemptPerKg,SlTaxableAmt,SlTaxCode,SlTaxRate,SlTaxAmt,SlTaxableAmt2,SlTaxRate2,SlTaxAmt2," _
'''                        & " SlAftTaxAdd1,SlAftTaxLess1,SlRoff,DueDays,DueDate,SlEntDt,SlUser,SlTranCd,SlIsHank,GpVno,GpNo,GpDate,GpVYear,SlTaxRate3, " _
'''                        & " SlTaxAmt3,SlCharityRt,SlCharityOn,SlCrDays,SlTaxPayAc,SlIsExempt,SlSubItCode,SlSubBag,SlSubWt,SlSubRt,SlSubNetRate,SlSubRtPer,SlSubItAmt, " _
'''                        & " SlSubBookNo,SlSubBookDt,SlSubGodown,SlTmpChr1,SlTcsOnAmt,SlTcsRate,SlTcsAmt,SlTmpChr3,SlIRNNo,SlAckNo) " & X
'''           X = "select * from tmpSale where SlUser = '" & gUserName & "'"
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
'''                Tattach = gReportPath & "rptSaleBillPrintGst.rpt"
'''                Set objReport = objCrystal.OpenReport(Tattach, 1)
'''
''''                Dim crTab As CRAXDRT.DatabaseTable
'''                objReport.Database.LogOnServer "pdsodbc.dll", "Hitrix", "Vishnu", "sa", "123456"
'''
'''
''''                For Each crTab In objReport.Database.Tables
''''                    crTab.location = gDbLocation
''''                Next
'''
'''                objReport.RecordSelectionFormula = "{tmpSale.SlUser}='" & gUserName & "' "
'''
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
'''                    Case "{@WprintFor}"
'''                          objReport.FormulaFields(j).text = "Trim(" & Chr(39) + "Original For Buyer" + Chr(39) & ")"
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
'''           PtyMbl = GProcGetColumnValue("TblMastAccount", "AcCode", RsSA!SlAcDrCode, "N", "AcContNo", "S")
'''           PtyName = GProcGetColumnValue("TblMastAccount", "AcCode", RsSA!SlAcDrCode, "N", "AcName", "S")
'''
'''           BrkMbl = GProcGetColumnValue("TblMastAccount", "AcCode", RsSA!SlBroker, "N", "AcContNo", "S")
'''           BrkName = GProcGetColumnValue("TblMastAccount", "AcCode", RsSA!SlBroker, "N", "AcName", "S")
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
'''  End If
''''-- End For WhatsApp SMS

'-- For SMS
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
            Dim RsTmp As Recordset
            
            MillCode = GProcGetColumnValue("tblMastAccount", "AcName", txtTemp(7), "S", "AcCode", "N")

'            TktNm = GProcGetColumnValue("tblMastItem", "ItName", msGrid.TextMatrix(1, 3), "S", "ItTicket", "S", " ItMillCode=" & MillCode) '--Item Code

            SmsMblNo = lblMblNoParty.Caption

            If Val(Right(Trim(txtTemp(5)), 2)) > 0 Then
               PtyName = Trim(Left(Trim(txtTemp(5)), Len(Trim(txtTemp(5))) - 2))
            Else
               PtyName = Trim(txtTemp(5))
            End If
            
            If Val(Right(Trim(txtTemp(7)), 2)) > 0 Then
               TktNm = Trim(Left(Trim(txtTemp(7)), Len(Trim(txtTemp(7))) - 2))
            Else
               TktNm = Trim(txtTemp(7))
            End If
            
            
            
            SmsString = "Firm   :" & vbLf
            SmsString = SmsString + "*" + gCName + "*" & vbLf & vbLf
            SmsString = SmsString + "*GATE PASS*" & vbLf & vbLf
            SmsString = SmsString + "No     : *" + Trim(txt(9)) + "*" & vbLf
            SmsString = SmsString + "Date  : " + Trim(CStr(Format(dtpDt(2), "dd/mm/yy"))) & vbLf & vbLf
            
            SmsString = SmsString + "Party  : *" + PtyName + "*" & vbLf & vbLf
            
            SmsString = SmsString + "Broker : " + txtTemp(8) & vbLf
            SmsString = SmsString + "Mill   : " + TktNm & vbLf & vbLf
            SmsString = SmsString + "Count  : " & vbLf
            SmsString = SmsString + "*" + Trim(msGrid.TextMatrix(1, 3)) + "*" & vbLf & vbLf
            
            SmsString = SmsString + "Qty.      : " + msGrid.TextMatrix(1, 6) & vbLf
            SmsString = SmsString + "Weight : " + msGrid.TextMatrix(1, 8) + " Kgs " & vbLf
            
'CStr(Format(RtPer, "0.0"))
            
            SmsString = SmsString + "Rate     : *" + CStr(Format(msGrid.TextMatrix(1, 9), "######0.000")) + "*" + " / " + msGrid.TextMatrix(1, 10) + " Kgs " & vbLf
            SmsString = SmsString + "( Plus GST and TCS )" & vbLf & vbLf
            
            SmsString = SmsString + "Amount : *" + Trim(txt(16)) + "*" & vbLf & vbLf
            
            SmsString = SmsString + "*Due Date  : " + Trim(CStr(Format(dtpDt(32), "dd/mm/yy"))) + "*" & vbLf
            
            
            
            
            
            
'            SmsString = "GP-" + Trim(txt(9)) + " Dt-" + Trim(CStr(Format(dtpDt(2), "dd/mm/yy"))) & vbLf
'            SmsString = SmsString + "Pty-" + Left(txttemp(5), 20) & vbLf
'            SmsString = SmsString + "Cnt-" + Left(msGrid.TextMatrix(1, 3), 15) & vbLf
'            If TktNm = "" Or IsNull(TktNm) Then
'                SmsString = SmsString + "Pkg-" + msGrid.TextMatrix(1, 6) + "-" + Left(txttemp(7), 10) & vbLf
'            Else
'                SmsString = SmsString + "Pkg-" + msGrid.TextMatrix(1, 6) + "-" + Left(TktNm, 10) & vbLf
'            End If
'            SmsString = SmsString + "Rt-" + msGrid.TextMatrix(1, 9) & vbLf
'            SmsString = SmsString + "Amt-" + CStr(Int(txt(16))) & vbLf
'            SmsString = SmsString + "Brk-" + Left(txttemp(8), 15) & vbLf
'            SmsString = SmsString + "Firm-" + Left(gCName, 15)
'---For WhatsApp Text SMS
            gCn.BeginTrans
            gCn.Execute "Delete from tmpWhatsAppErr where username = '" & gUserName & "' and Vfirm = '" & gCCode & "' "
            gCn.CommitTrans
            
            If SmsMblNo = "" Then
               MsgBox ("Party Mobile No Not Found....")
            Else
              PtyName = txtTemp(5)
              SentWhatsApp "Txt", "", SmsString, SmsMblNo, PtyName
            End If
            '-- SMS To Broker
            SmsMblNoBrk = GProcGetColumnValue("TblMastAccount", "ACName", txtTemp(8), "S", "AcContNo", "S")
            If SmsMblNoBrk = "" Then
               MsgBox ("Broker Mobile No Not Found....")
            Else
              PtyName = txtTemp(8)
              SentWhatsApp "Txt", "", SmsString, SmsMblNoBrk, PtyName
            End If
            
            '--Error
            X = "Select * from tmpWhatsAppErr where username = '" & gUserName & "' and Vfirm = '" & gCCode & "'"
            If GProcRstOpen(RsTmp, X, "R", gCn) > 0 Then
               GProcWhatsAppError
            End If

'---End For WhatsApp Text SMS


'---For SMS
'            If SmsMblNo = "" Then
'               MsgBox ("Party Mobile No Not Found....")
'            Else
'                url = "https://api-alerts.kaleyra.com/v4/?"
'                params = "api_key=A967d646cc4202f5c44523b468617f117&method=sms&message=" & SmsString & "&to=" & SmsMblNo & "&sender=VISHTX"
'                HttpReq.Open "POST", url, False
'                HttpReq.setRequestHeader "Content-Type", "application/x-www-form-urlencoded"
'                HttpReq.send params
'            End If
'            '-- SMS To Broker
'            SmsMblNoBrk = GProcGetColumnValue("TblMastAccount", "ACName", txttemp(8), "S", "AcContNo", "S")
'            If SmsMblNoBrk = "" Then
'               MsgBox ("Broker Mobile No Not Found....")
'            Else
'                url = "https://api-alerts.kaleyra.com/v4/?"
'                params = "api_key=A967d646cc4202f5c44523b468617f117&method=sms&message=" & SmsString & "&to=" & SmsMblNoBrk & "&sender=VISHTX"
'                HttpReq.Open "POST", url, False
'                HttpReq.setRequestHeader "Content-Type", "application/x-www-form-urlencoded"
'                HttpReq.send params
'            End If
'--End For SMS
          End If
        End If
    End If
Case 14 'Cancel
    lblTcsLimit.Caption = ""
    
    FormAction = vbDataActionCancel
    cmdBtn_Click (3)
    txtGrid.Visible = False
    If rstbl.RecordCount <= 0 Then
        Call GProcActivateControls(Me, False)
        Call GProcSetButtons(Me, vbDataActionCancel)
    Else
        Call GProcActivateControls(Me, False)
        Call GProcSetButtons(Me, True)
    End If
    If cmdBtn(6).Visible = True Then cmdBtn(6).SetFocus
    msGrid.Col = 0
    msGrid.Row = 1
    txtGrid.Move msGrid.CellLeft + msGrid.Left, msGrid.CellTop + msGrid.Top, msGrid.CellWidth, msGrid.CellHeight
    If frmTypeStr = "SE" Or frmTypeStr = "SY" Then
        lblBag.Visible = True
        lblWeight.Visible = True
        lblBalBag.Visible = True
        lblBalWt.Visible = True
    Else
        lblBag.Visible = False
        lblWeight.Visible = False
        lblBalBag.Visible = False
        lblBalWt.Visible = False
    End If
End Select
Exit Sub
ErrorRoutine:
    rstbl.CancelUpdate
'    GProcErrorHandler
End Sub
Private Sub SetControlEd()
txtTemp(6).Enabled = False '----Purch A/c
txtTemp(35).Enabled = False
'-- Temp
If gCYear < 2018 Then
   txt(16).Enabled = True '-- Amt
Else
   txt(16).Enabled = False '-- Amt
End If
txt(21).Enabled = False '-- Ass Val
dtpDt(32).Enabled = False '--Due Dt
txtDrBalance.Enabled = False
txt(24).Enabled = False
txt(30).Enabled = False
txt(63).Enabled = False
txt(64).Enabled = False
txt(65).Enabled = False


If ChgRoff = 1 Then
   txt(24).Enabled = True
   txt(27).Enabled = True
   txt(42).Enabled = True
Else
   txt(24).Enabled = False
   txt(27).Enabled = False
   txt(42).Enabled = False
End If
If FormAction = vbDataActionAddNew Then
   txt(0).Enabled = False
End If
If frmTypeStr = "SY" Or frmTypeStr = "SE" Then
   Label6.Caption = "Purchase"
   lblBag.Caption = " Bill No.       : "
   lblWeight.Caption = " Vou. No.    : "
Else
   Label6.Caption = "Stock"
   lblBag.Caption = " Bag         : "
   lblWeight.Caption = " Weight      : "
End If
txt(23).Enabled = False
txt(26).Enabled = False
txt(41).Enabled = False

If Val(txt(17)) > 0 Then
   lblCharity.Visible = True
   txt(17).Visible = True
   txt(43).Visible = True
   Cbo(44).Visible = True
Else
   lblCharity.Visible = False
   txt(17).Visible = False
   txt(43).Visible = False
   Cbo(44).Visible = False
End If
If frmTypeStr = "SE" Then
   Label16.Visible = True
   txt(19).Visible = True
   txt(20).Visible = True
   Label18.Visible = True
End If
If gCCode = "GMW" Then
   Label26.Visible = True
   txt(59).Visible = True
End If

txtCode(70).Enabled = False
txtTemp(70).Enabled = False



End Sub
Private Sub DeleteRelatedRecord()
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
Dim X As String
Dim BillDt
Dim DueDt
'--- Account Code
MillCode = GProcGetColumnValue("tblMastAccount", "AcName", txtTemp(7), "S", "AcCode", "N")
DrCode = GProcGetColumnValue("tblMastAccount", "AcName", txtTemp(5), "S", "AcCode", "N")
CrCode = GProcGetColumnValue("tblMastAccount", "AcName", txtTemp(6), "S", "AcCode", "N")
BrkCode = GProcGetColumnValue("tblMastAccount", "AcName", txtTemp(8), "S", "AcCode", "N")
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
        '--- It Ctrl No(11)
        '--Pending
'        If CLng(.TextMatrix(i, 14)) = 0 Then 'FormAction = vbDataActionAddNew Then
'            rstSub!VSubItCtrlNo = GProcGenerateId("tblPurchSub", "PurSubItCtrlNo", " Vfirm = '" & gCCode & "' and  VYear=" & gCYear)   '--It Ctrl No
'        Else
'            rstSub!VSubItCtrlNo = CLng(.TextMatrix(i, 14))
'        End If
        '--end pending
        rstSub!SlSubPItCtrlNo = Val(.TextMatrix(i, 14))
        rstSub!SlSubBookNo = CLng(.TextMatrix(i, 1)) '--Book No
        '--Book Dt
        If IsDate(.TextMatrix(i, 2)) Then
            rstSub!SlSubBookDt = .TextMatrix(i, 2)
        Else
            rstSub!SlSubBookDt = Null
        End If
        rstSub!SlSubBkType = .TextMatrix(i, 15) & vbNullString
        rstSub!SlSubBkVyear = .TextMatrix(i, 16)
        rstSub!SlSubBkSrNo = .TextMatrix(i, 17)
        
        '--Itcode(4)
        rstSub!SlSubItCode = GProcGetColumnValue("tblMastItem", "ItName", .TextMatrix(i, 3), "S", "ItCode", "N", " ItMillCode=" & MillCode) '--Item Code
        rstSub!SlSubBag = CLng(.TextMatrix(i, 6)) '--Bag(6)
        rstSub!SlSubwt = CDbl(.TextMatrix(i, 8)) '-- Wt(7)
        rstSub!SlSubRt = CDbl(.TextMatrix(i, 9)) '--Rate(8)
        rstSub!SlSubRtPer = CDbl(.TextMatrix(i, 10)) '--Rate
        rstSub!SlSubNetRate = CDbl(.TextMatrix(i, 11)) '--Net Rate
        rstSub!SlSubAmt = CDbl(.TextMatrix(i, 12)) '--Amt(10)
        rstSub!SlSubLotNo = .TextMatrix(i, 5) '--Lot No(5)
        '--Godown Code(4)
'        .TextMatrix(i, 4) = "-"
        rstSub!SlSubGodown = GProcGetColumnValue("tblMastNarration", "Narration", .TextMatrix(i, 4), "S", "NarrCode", "N", "")
        '-- P It Ctrl No =Booking It ctrl No
        rstSub.Update
        i = i + 1
     Loop
 End With
 rstSub.Close
 cmdBtn(13).Enabled = False
End Sub

Private Sub CmdClose_Click()
    DoEvents
    Frame2.Visible = False
    txtCode(8).SetFocus
End Sub

Private Sub cmdPrintBill_Click(Index As Integer)
Dim RsSA As Recordset
Dim X As String
Dim XX As String
Dim BillSrNo As String
Dim InvNo As Long
Dim ToInvNo As Long
Dim Y As String
Dim BilNo As String
Dim YY As String
Dim tmpVno As Long
Dim tmpVtype As String
Dim tmpVyear As Long
Dim tmpVFirm As String
Dim GimBl As String
Dim GstinDepot As String
Dim PanDepot As String
Dim PrintFor As String
Dim PayTo As String
Dim aFile As String

Select Case Index
Case 0, 1, 6:
    If Index = 0 Then GProcCrystalRptPreparation cryReport, "Printer"
    If Index = 1 Then GProcCrystalRptPreparation cryReport, "Window"
    If Index = 3 Then GProcCrystalRptPreparation cryReport, "Window"
    If Index = 6 Then GProcCrystalRptPreparation cryReport, "Window"
    
    cryReport.WindowTitle = "Sale Bill Print"
    '-- A/C Bill SrNo
    If Val(txtPInvno(0)) > 0 Then
       InvNo = CLng(txtPInvno(0))
       ToInvNo = CLng(txtPInvno(1))
       BillSrNo = ""
    Else
       InvNo = CLng(Mid(txtPInvno(0), 5, 15))
       ToInvNo = CLng(Mid(txtPInvno(1), 5, 15))
       BillSrNo = Left(txtPInvno(0), 4)
    End If
    '------ Validate
    If IsNumeric(InvNo) = False Then InvNo = 0
    If IsNumeric(ToInvNo) = False Then ToInvNo = InvNo
    If ToInvNo < InvNo Then ToInvNo = InvNo
    
    If chkIsDatewise.Value = 1 Then
        X = "Select a.Vno,a.Vtype,a.Vdt,a.Vyear,a.VFirm,a.SlAcDrCode,a.SlAcCrCode,a.SlMillCode,a.SlBroker,a.SlBillNo,a.SlBillDt,a.SlTransport,a.SlLorryNo,a.SlLrNo,a.SlLrDate, " _
            & "a.SlSubAmt,a.SlBillAmt,a.SlAdd1,a.SlLess1,a.SlExemptAmt,a.SlExemptPerKg,a.SlTaxableAmt,a.SlTaxCode,a.SlTaxRate,a.SlTaxAmt,a.SlTaxableAmt2,a.SlTaxRate2,a.SlTaxAmt2, " _
            & " a.SlAftTaxAdd1,a.SlAftTaxLess1,a.SlRoff,a.DueDays,a.DueDate,a.SlEntDt,'" & gUserName & "',a.SlTranCd,a.SlIsHank,a.GpVno,a.GpNo,a.GpDate,a.GpVYear,a.SlTaxRate3, " _
            & " a.SlTaxAmt3,a.SlCharityRt,a.SlCharityOn,a.SlCrDays,a.SlTaxPayAc,a.SlIsExempt,b.SlSubItCode,b.SlSubBag,b.SlSubWt,b.SlSubRt,b.SlSubNetRate,b.SlSubRtPer,b.SlSubAmt, " _
            & " b.SlSubBookNo,b.SlSubBookDt,b.SlSubGodown,b.SlSubLotNo,a.SlTcsOnAmt,a.SlTcsRate,a.SlTcsAmt,a.SlTmpChr3,a.SlIRNNo,a.SlAckNo,a.SlTdsAmt from tblSale A ,tblSaleSub B" _
            & " where a.vno = b.vno and a.vtype = b.vtype and a.vyear = b.vyear and a.vfirm = b.vfirm and  a.VDt = '" & Format(dtpDtPrintBill, "yyyy/mm/dd") & "' and a.Vtype='" & frmTypeStr & "' and a.VFirm = '" & gCCode & "' And a.VYear = " & gCYear
    Else
        X = "("
        Do While InvNo <= ToInvNo
           If BillSrNo = "" Then
              X = X + "'" & InvNo & "',"
           Else
              If frmTypeStr = "SD" And txtCode(7) = "ISCCGMW" Then
                 GimBl = Right("000000" + CStr(InvNo), 6)
                 X = X + "'" & BillSrNo & "" & GimBl & "',"
              Else
                 X = X + "'" & BillSrNo & "" & InvNo & "',"
              End If
           End If
           InvNo = InvNo + 1
        Loop
        X = Left(X, Len(X) - 1) + ")"
        Y = X
        X = "Select a.Vno,a.Vtype,a.Vdt,a.Vyear,a.VFirm,a.SlAcDrCode,a.SlAcCrCode,a.SlMillCode,a.SlBroker,a.SlBillNo,a.SlBillDt,a.SlTransport,a.SlLorryNo,a.SlLrNo,a.SlLrDate, " _
            & "a.SlSubAmt,a.SlBillAmt,a.SlAdd1,a.SlLess1,a.SlExemptAmt,a.SlExemptPerKg,a.SlTaxableAmt,a.SlTaxCode,a.SlTaxRate,a.SlTaxAmt,a.SlTaxableAmt2,a.SlTaxRate2,a.SlTaxAmt2, " _
            & " a.SlAftTaxAdd1,a.SlAftTaxLess1,a.SlRoff,a.DueDays,a.DueDate,a.SlEntDt,'" & gUserName & "',a.SlTranCd,a.SlIsHank,a.GpVno,a.GpNo,a.GpDate,a.GpVYear,a.SlTaxRate3, " _
            & " a.SlTaxAmt3,a.SlCharityRt,a.SlCharityOn,a.SlCrDays,a.SlTaxPayAc,a.SlIsExempt,b.SlSubItCode,b.SlSubBag,b.SlSubWt,b.SlSubRt,b.SlSubNetRate,b.SlSubRtPer,b.SlSubAmt, " _
            & " b.SlSubBookNo,b.SlSubBookDt,b.SlSubGodown,b.SlSubLotNo,a.SlTcsOnAmt,a.SlTcsRate,a.SlTcsAmt,a.SlTmpChr3,a.SlIRNNo,a.SlAckNo,a.SlTdsAmt from tblSale A ,tblSaleSub B" _
            & " where a.vno = b.vno and a.vtype = b.vtype and a.vyear = b.vyear and a.vfirm = b.vfirm and  a.SlBillNo in " & X & " and a.Vtype='" & frmTypeStr & "' and a.VFirm = '" & gCCode & "' And a.VYear = " & gCYear
   End If
   gCn.BeginTrans
   
   
   gCn.Execute "Delete from tmpSale where  SlUser='" & gUserName & "'"
     
   gCn.Execute " insert into tmpSale (Vno,Vtype,Vdt,Vyear,VFirm,SlAcDrCode,SlAcCrCode,SlMillCode,SlBroker,SlBillNo,SlBillDt,SlTransport,SlLorryNo,SlLrNo,SlLrDate, " _
               & "SlSubAmt,SlBillAmt,SlAdd1,SlLess1,SlExemptAmt,SlExemptPerKg,SlTaxableAmt,SlTaxCode,SlTaxRate,SlTaxAmt,SlTaxableAmt2,SlTaxRate2,SlTaxAmt2," _
               & " SlAftTaxAdd1,SlAftTaxLess1,SlRoff,DueDays,DueDate,SlEntDt,SlUser,SlTranCd,SlIsHank,GpVno,GpNo,GpDate,GpVYear,SlTaxRate3, " _
               & " SlTaxAmt3,SlCharityRt,SlCharityOn,SlCrDays,SlTaxPayAc,SlIsExempt,SlSubItCode,SlSubBag,SlSubWt,SlSubRt,SlSubNetRate,SlSubRtPer,SlSubItAmt, " _
               & " SlSubBookNo,SlSubBookDt,SlSubGodown,SlTmpChr1,SlTcsOnAmt,SlTcsRate,SlTcsAmt,SlTmpChr3,SlIRNNo,SlAckNo,SlTdsAmt) " & X
   gCn.CommitTrans
   
   QrCodeData
     
     With cryReport
        .Formulas(10) = "wCSTNo = '" & gCCST & "'"
        .Formulas(11) = "wBSTNo = '" & gCBST & "'"
        If frmTypeStr = "SD" Then
           GstinDepot = GProcGetColumnValue("tblMastAccount", "AcCode", txt(7), "N", "AcGSTIN", "S")
           PanDepot = GProcGetColumnValue("tblMastAccount", "AcCode", txt(7), "N", "AcPAN", "S")
           PayTo = GProcGetColumnValue("tblMastAccount", "AcCode", "'" & gCDepotMainFirm & "' ", "N", "AcName", "S")
           .Formulas(12) = "wGSTNo = '" & GstinDepot & "'"
           .Formulas(13) = "wcPan = '" & PanDepot & "'"
        Else
           .Formulas(12) = "wGSTNo = '" & gCGStin & "'"
           .Formulas(13) = "wcPan = '" & gCPAN & "'"
           PayTo = gCName
        End If
        .Formulas(15) = "wPayTo = '" & PayTo & "'"
        .Formulas(21) = "wBank = '" & gCBankName & "'"
        .Formulas(22) = "wBankAcNo = '" & gCBankAcNo & "'"
        .Formulas(23) = "wBankIfsc = '" & gCBankRTGSCode & "'"
        
        X = "{tmpSale.SlUser}='" & gUserName & "' "
        .SelectionFormula = X
        If chkInvChallan(0).Value = 1 Then
           PrintFor = chkInvChallan(0).Caption
           .Formulas(14) = "wPrintFor = '" & PrintFor & "'"
           .ReportFileName = gReportPath & "rptSaleBillPrintGstQr.rpt"
           If Index = 0 Or Index = 1 Then
              .Action = 1
           End If
        End If
        If chkInvChallan(1).Value = 1 Then
           PrintFor = chkInvChallan(1).Caption
           .Formulas(14) = "wPrintFor = '" & PrintFor & "'"
           .ReportFileName = gReportPath & "rptSaleBillPrintGstQr.rpt"
           If Index = 0 Or Index = 1 Then
              .Action = 1
           End If
        End If
        If chkInvChallan(2).Value = 1 Then
           PrintFor = chkInvChallan(2).Caption
           .Formulas(14) = "wPrintFor = '" & PrintFor & "'"
           .ReportFileName = gReportPath & "rptSaleBillPrintGstQr.rpt"
           If Index = 0 Or Index = 1 Then
              .Action = 1
           End If
        End If
        If chkInvChallan(3).Value = 1 Then
           PrintFor = chkInvChallan(3).Caption
           .Formulas(14) = "wPrintFor = '" & PrintFor & "'"
           .ReportFileName = gReportPath & "rptSaleBillPrintGstQr.rpt"
           If Index = 0 Or Index = 1 Then
              .Action = 1
           End If
        End If
        
'        If Index = 0 And chkInvChallan(4).Value = 1 Then
'            Dim i As Integer
'            For i = 1 To 3
'            Select Case i
'            Case 1
'             .Formulas(20) = "wcopy = 'Godown Copy'"
'            Case 2
'             .Formulas(20) = "wcopy = 'Customer Copy'"
'            Case 3
'             .Formulas(20) = "wcopy = 'Office Copy'"
'            End Select
'            .ReportFileName = gReportPath & "rptSaleBillGP.rpt"
'            .Action = 1
'             Next i
'         Else
        If chkInvChallan(4).Value = 1 Then
           PrintFor = chkInvChallan(4).Caption
           .Formulas(14) = "wPrintFor = '" & PrintFor & "'"
'           If ChkGod.Value = 1 Then
'             .Formulas(20) = "wcopy = 'Godown Copy'"
'           ElseIf ChkCus.Value = 1 Then
'             .Formulas(20) = "wcopy = 'Customer Copy'"
'           ElseIf chkOff.Value = 1 Then
'             .Formulas(20) = "wcopy = 'Office Copy'"
'            End If
           If Dir("D:\HiReports\rptSaleBillGP.rpt") <> "" Then
              .ReportFileName = "D:\HiReports\rptSaleBillGP.rpt"
           Else
              .ReportFileName = gReportPath & "rptSaleBillGP.rpt"
           End If
           If Index = 0 Or Index = 1 Then
              .Action = 1
           End If
        End If
        If Index = 6 Then
           '--Error
           gCn.BeginTrans
           gCn.Execute "Delete from tmpWhatsAppErr where username = '" & gUserName & "' and Vfirm = '" & gCCode & "' "
           gCn.CommitTrans
           Dim PtyMbl As String
           Dim BrkMbl As String
           Dim FileNMWithPath As String
           Dim FileNM As String
           Dim PtyName As String
           Dim BrkName As String
'           Dim aFile As String
           Dim RsTmp As Recordset
           Dim Tattach As String
           Dim Mfile As String
'           Dim GstinDepot As String
'           Dim PanDepot As String
'           Dim PayTo As String
           X = "select * from tmpSale where SlUser = '" & gUserName & "'"
           If GProcRstOpen(RsSA, X, "R") <= 0 Then
              MsgBox "No Records..."
              Exit Sub
           End If
           RsSA.MoveFirst
           Do While Not RsSA.EOF
              With cryReport
                   Dim objCrystal As CRAXDRT.Application
                   Dim objReport As CRAXDRT.Report
    
                   Set objCrystal = New CRAXDRT.Application
                   Tattach = gReportPath & "rptSaleBillPrintGst.rpt"
                   Set objReport = objCrystal.OpenReport(Tattach, 1)
    
                   objReport.Database.LogOnServer "pdsodbc.dll", "Hitrix", "Vishnu", "sa", "123456"
                   objReport.RecordSelectionFormula = "{tmpSale.SlBillNo} = '" & RsSA!SlBillNo & "' and {tmpSale.SlUser}='" & gUserName & "' "
   
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
                        Case "{@wcpan}"
                              If frmTypeStr = "SD" Then
                                 PanDepot = GProcGetColumnValue("tblMastAccount", "AcCode", txt(7), "N", "AcPAN", "S")
                                 objReport.FormulaFields(j).text = "Trim(" & Chr(39) & UCase(PanDepot) & Chr(39) & ")"
                              Else
                                 objReport.FormulaFields(j).text = "Trim(" & Chr(39) & UCase(gCPAN) & Chr(39) & ")"
                              End If
                        Case "{@wCSTNo}"
                              objReport.FormulaFields(j).text = "Trim(" & Chr(39) & UCase(gCCST) & Chr(39) & ")"
                        Case "{@wBSTNo}"
                              objReport.FormulaFields(j).text = "Trim(" & Chr(39) & UCase(gCBST) & Chr(39) & ")"
                        Case "{@wGSTNo}"
                              If frmTypeStr = "SD" Then
                                 GstinDepot = GProcGetColumnValue("tblMastAccount", "AcCode", txt(7), "N", "AcGSTIN", "S")
                                 objReport.FormulaFields(j).text = "Trim(" & Chr(39) & UCase(GstinDepot) & Chr(39) & ")"
                              Else
                                 objReport.FormulaFields(j).text = "Trim(" & Chr(39) & UCase(gCGStin) & Chr(39) & ")"
                              End If
                        Case "{@wPayTo}"
                              If frmTypeStr = "SD" Then
                                 PayTo = GProcGetColumnValue("tblMastAccount", "AcCode", "'" & gCDepotMainFirm & "' ", "N", "AcName", "S")
                              Else
                                 PayTo = gCName
                              End If
                              objReport.FormulaFields(j).text = "Trim(" & Chr(39) & (PayTo) & Chr(39) & ")"
                        Case "{@wBank}"
                              objReport.FormulaFields(j).text = "Trim(" & Chr(39) & gCBankName & Chr(39) & ")"
                        Case "{@wBankAcNo}"
                              objReport.FormulaFields(j).text = "Trim(" & Chr(39) & gCBankAcNo & Chr(39) & ")"
                        Case "{@wBankIfsc}"
                              objReport.FormulaFields(j).text = "Trim(" & Chr(39) & gCBankRTGSCode & Chr(39) & ")"
                        Case "{@WprintFor}"
                              objReport.FormulaFields(j).text = "Trim(" & Chr(39) + "Original For Buyer" + Chr(39) & ")"
                        End Select
                    Next j
               End With
    
                       '--
               FileNM = "" & Format(Now, "ddMMyyyyhhmmss") & "" & ".pdf"
               FileNMWithPath = gReportPath & "" & FileNM & ""
               PtyMbl = GProcGetColumnValue("TblMastAccount", "AcCode", RsSA!SlAcDrCode, "N", "AcContNo", "S")
               PtyName = GProcGetColumnValue("TblMastAccount", "AcCode", RsSA!SlAcDrCode, "N", "AcName", "S")
    
               BrkMbl = GProcGetColumnValue("TblMastAccount", "AcCode", RsSA!SlBroker, "N", "AcContNo", "S")
               BrkName = GProcGetColumnValue("TblMastAccount", "AcCode", RsSA!SlBroker, "N", "AcName", "S")
               If PtyMbl = "" Or IsNull(PtyMbl) Then
                  MsgBox "Party - " + PtyName + " Phone Not found..."
               Else
                  ExportReportToPDF objReport, FileNMWithPath, "foo"
                  SentWhatsApp "PDF", FileNM, "", PtyMbl, PtyName
               End If
               If BrkMbl = "" Or IsNull(BrkMbl) Then
                  MsgBox "Broker - " + BrkName + " Phone Not found..."
               Else
                  ExportReportToPDF objReport, FileNMWithPath, "foo"
                  SentWhatsApp "PDF", FileNM, "", BrkMbl, BrkName
               End If

               RsSA.MoveNext
           Loop
           X = "Select * from tmpWhatsAppErr where username = '" & gUserName & "' and vfirm = '" & gCCode & "'"
           If GProcRstOpen(RsTmp, X, "R", gCn) > 0 Then
              GProcWhatsAppError
           End If
           gReportPath = App.Path & "\HIReports\"
           aFile = gReportPath & "*.pdf"
           If Len(Dir$(aFile)) > 0 Then
              Kill aFile
           End If
        End If
     End With
Case 3, 4:
    If Val(txtPInvno(0)) > 0 Then
       InvNo = CLng(txtPInvno(0))
       ToInvNo = CLng(txtPInvno(1))
       BillSrNo = ""
    Else
       InvNo = CLng(Mid(txtPInvno(0), 5, 15))
       ToInvNo = CLng(Mid(txtPInvno(1), 5, 15))
       BillSrNo = Left(txtPInvno(0), 4)
    End If
    '------ Validate
    If IsNumeric(InvNo) = False Then InvNo = 0
    If IsNumeric(ToInvNo) = False Then ToInvNo = InvNo
    If ToInvNo < InvNo Then ToInvNo = InvNo
    
    If chkIsDatewise.Value = 1 Then
        X = "Select a.Vno,a.Vtype,a.Vdt,a.Vyear,a.VFirm,a.SlAcDrCode,a.SlAcCrCode,a.SlMillCode,a.SlBroker,a.SlBillNo,a.SlBillDt,a.SlTransport,a.SlLorryNo,a.SlLrNo,a.SlLrDate, " _
            & "a.SlSubAmt,a.SlBillAmt,a.SlAdd1,a.SlLess1,a.SlExemptAmt,a.SlExemptPerKg,a.SlTaxableAmt,a.SlTaxCode,a.SlTaxRate,a.SlTaxAmt,a.SlTaxableAmt2,a.SlTaxRate2,a.SlTaxAmt2, " _
            & " a.SlAftTaxAdd1,a.SlAftTaxLess1,a.SlRoff,a.DueDays,a.DueDate,a.SlEntDt,'" & gUserName & "',a.SlTranCd,a.SlIsHank,a.GpVno,a.GpNo,a.GpDate,a.GpVYear,a.SlTaxRate3, " _
            & " a.SlTaxAmt3,a.SlCharityRt,a.SlCharityOn,a.SlCrDays,a.SlTaxPayAc,a.SlIsExempt,b.SlSubItCode,b.SlSubBag,b.SlSubWt,b.SlSubRt,b.SlSubNetRate,b.SlSubRtPer,b.SlSubAmt, " _
            & " b.SlSubBookNo,b.SlSubBookDt,b.SlSubGodown,b.SlSubLotNo,a.SlTcsOnAmt,a.SlTcsRate,a.SlTcsAmt,a.SlTmpChr3 from tblSale A ,tblSaleSub B" _
            & " where a.vno = b.vno and a.vtype = b.vtype and a.vyear = b.vyear and a.vfirm = b.vfirm and  a.VDt = '" & Format(dtpDtPrintBill, "yyyy/mm/dd") & "' and a.Vtype='" & frmTypeStr & "' and a.VFirm = '" & gCCode & "' And a.VYear = " & gCYear
    Else
        X = "("
        Do While InvNo <= ToInvNo
           If BillSrNo = "" Then
              X = X + "'" & InvNo & "',"
           Else
              If frmTypeStr = "SD" And txtCode(7) = "ISCCGMW" Then
                 GimBl = Right("000000" + CStr(InvNo), 6)
                 X = X + "'" & BillSrNo & "" & GimBl & "',"
              Else
                 X = X + "'" & BillSrNo & "" & InvNo & "',"
              End If
           End If
           InvNo = InvNo + 1
        Loop
        X = Left(X, Len(X) - 1) + ")"
        Y = X
        X = "Select a.Vno,a.Vtype,a.Vdt,a.Vyear,a.VFirm,a.SlAcDrCode,a.SlAcCrCode,a.SlMillCode,a.SlBroker,a.SlBillNo,a.SlBillDt,a.SlTransport,a.SlLorryNo,a.SlLrNo,a.SlLrDate, " _
            & "a.SlSubAmt,a.SlBillAmt,a.SlAdd1,a.SlLess1,a.SlExemptAmt,a.SlExemptPerKg,a.SlTaxableAmt,a.SlTaxCode,a.SlTaxRate,a.SlTaxAmt,a.SlTaxableAmt2,a.SlTaxRate2,a.SlTaxAmt2, " _
            & " a.SlAftTaxAdd1,a.SlAftTaxLess1,a.SlRoff,a.DueDays,a.DueDate,a.SlEntDt,'" & gUserName & "',a.SlTranCd,a.SlIsHank,a.GpVno,a.GpNo,a.GpDate,a.GpVYear,a.SlTaxRate3, " _
            & " a.SlTaxAmt3,a.SlCharityRt,a.SlCharityOn,a.SlCrDays,a.SlTaxPayAc,a.SlIsExempt,b.SlSubItCode,b.SlSubBag,b.SlSubWt,b.SlSubRt,b.SlSubNetRate,b.SlSubRtPer,b.SlSubAmt, " _
            & " b.SlSubBookNo,b.SlSubBookDt,b.SlSubGodown,b.SlSubLotNo,a.SlTcsOnAmt,a.SlTcsRate,a.SlTcsAmt,a.SlTmpChr3 from tblSale A ,tblSaleSub B" _
            & " where a.vno = b.vno and a.vtype = b.vtype and a.vyear = b.vyear and a.vfirm = b.vfirm and  a.SlBillNo in " & X & " and a.Vtype='" & frmTypeStr & "' and a.VFirm = '" & gCCode & "' And a.VYear = " & gCYear
   End If
   gCn.Execute "Delete from tmpSale where  SlUser='" & gUserName & "'"
     
   gCn.Execute " insert into tmpSale (Vno,Vtype,Vdt,Vyear,VFirm,SlAcDrCode,SlAcCrCode,SlMillCode,SlBroker,SlBillNo,SlBillDt,SlTransport,SlLorryNo,SlLrNo,SlLrDate, " _
               & "SlSubAmt,SlBillAmt,SlAdd1,SlLess1,SlExemptAmt,SlExemptPerKg,SlTaxableAmt,SlTaxCode,SlTaxRate,SlTaxAmt,SlTaxableAmt2,SlTaxRate2,SlTaxAmt2," _
               & " SlAftTaxAdd1,SlAftTaxLess1,SlRoff,DueDays,DueDate,SlEntDt,SlUser,SlTranCd,SlIsHank,GpVno,GpNo,GpDate,GpVYear,SlTaxRate3, " _
               & " SlTaxAmt3,SlCharityRt,SlCharityOn,SlCrDays,SlTaxPayAc,SlIsExempt,SlSubItCode,SlSubBag,SlSubWt,SlSubRt,SlSubNetRate,SlSubRtPer,SlSubItAmt, " _
               & " SlSubBookNo,SlSubBookDt,SlSubGodown,SlTmpChr1,SlTcsOnAmt,SlTcsRate,SlTcsAmt,SlTmpChr3) " & X

    XX = "Select * from tmpSale where  SlUser='" & gUserName & "'"
    XX = XX & " order by SlBillNo"
    If GProcRstOpen(RsSA, X, "R") <= 0 Then
       MsgBox "No Records..."
       Exit Sub
    End If
    
    If Index = 4 Then
        GProcCrystalRptPreparation cryReport, "Window"
        With cryReport
             .Formulas(4) = "wCompGstin = '" & gCGStin & "'"
             .SelectionFormula = "{tmpSale.SlUser} = '" & gUserName & "' "
             .ReportFileName = gReportPath & "rpte-Invoice.rpt"
             .Action = 1
        End With
        Exit Sub
    End If
    Dim LnCnt As Integer
    Dim FirmGstin As String
    Dim PartyName As String
    Dim PartyGstin As String
    Dim PartyAdd1 As String
    Dim PartyAdd2 As String
    Dim PartyCity As String
    Dim PartyPin As String
    Dim ItemHsn As String
'    Dim aFile As String
    Dim ii As Integer
    LnCnt = 1
             
    Dim ConsiName As String
    Dim ConsiGstin As String
    Dim ConsiAdd1 As String
    Dim ConsiAdd2 As String
    Dim ConsiCity As String
    Dim ConsiPin As String
             
             
             
    aFile = "D:\e-Inv-" + gCCode + "*.json"
    If Len(Dir$(aFile)) > 0 Then
          Kill aFile
    End If
    
    Open "D:\e-Inv-" + gCCode + ".json" For Output As #1
    
    Print #1, "["
    RsSA.MoveFirst
    Do While Not RsSA.EOF
       If frmTypeStr = "SD" Then
          FirmGstin = GProcGetColumnValue("tblMastAccount", "AcCode", RsSA.Fields("SlAcCrCode"), "N", "AcGSTIN", "S")
       Else
          FirmGstin = gCGStin
       End If
       PartyName = GProcGetColumnValue("tblMastAccount", "AcCode", RsSA.Fields("SlAcDrCode"), "N", "AcName", "S")
       PartyGstin = GProcGetColumnValue("tblMastAccount", "AcCode", RsSA.Fields("SlAcDrCode"), "N", "AcGSTIN", "S")
       PartyAdd1 = GProcGetColumnValue("tblMastAccount", "AcCode", RsSA.Fields("SlAcDrCode"), "N", "AcAdd1", "S")
       PartyAdd2 = GProcGetColumnValue("tblMastAccount", "AcCode", RsSA.Fields("SlAcDrCode"), "N", "AcAdd2", "S")
       PartyCity = GProcGetColumnValue("tblMastAccount", "AcCode", RsSA.Fields("SlAcDrCode"), "N", "AcCity", "S")
       PartyPin = GProcGetColumnValue("tblMastAccount", "AcCode", RsSA.Fields("SlAcDrCode"), "N", "AcZip", "S")
       If gCYear >= 2021 Then
          ItemHsn = GProcGetColumnValue("tblMastItem", "ItCode", RsSA.Fields("SlSubItCode"), "S", "ItHsn2", "S")
       Else
          ItemHsn = GProcGetColumnValue("tblMastItem", "ItCode", RsSA.Fields("SlSubItCode"), "S", "ItHsn", "S")
       End If
       If Val(Right(PartyName, 2)) > 0 Then
           PartyName = Trim(Left(PartyName, Len(PartyName) - 2))
       End If
       
       If Len(PartyPin) > 6 Then
          PartyPin = Left(PartyPin, 3) + Right(PartyPin, 3)
       End If
       If PartyAdd1 = "" Then PartyAdd1 = "null"
       If PartyAdd2 = "" Then PartyAdd2 = "null"
       '---Checks
       ii = InStr(1, PartyName, Chr(34))
       If ii > 0 Then
          MsgBox ("Name Contens " + """ ""  ..For '" & PartyName & "'   ")
          Close #1
          Exit Sub
       End If
       ii = InStr(1, PartyAdd1, Chr(34))
       If ii > 0 Then
          MsgBox ("Address Line 1 Contens " + """ ""  ..For '" & PartyName & "'   ")
          Close #1
          Exit Sub
       End If
       ii = InStr(1, PartyAdd2, Chr(34))
       If ii > 0 Then
          MsgBox ("Address Line 2 Contens " + """ ""  ..For '" & PartyName & "'   ")
          Close #1
          Exit Sub
       End If
       If Trim(PartyCity) = "" Then
          MsgBox ("City Not Found..For +'" & PartyName & "'   ")
          Close #1
          Exit Sub
       End If
       If Trim(PartyPin) = "" Then
          MsgBox ("Pin Not Found..For +'" & PartyName & "'   ")
          Close #1
          Exit Sub
       End If
       
       ConsiName = GProcGetColumnValue("tblMastNarration", "NarrCode", RsSA.Fields("SlTaxCode"), "N", "Narration", "S")
       If ConsiName <> "-" Then
          ConsiGstin = GProcGetColumnValue("tblMastNarration", "NarrCode", RsSA.Fields("SlTaxCode"), "N", "ConsiGSTIN", "S")
          ConsiAdd1 = GProcGetColumnValue("tblMastNarration", "NarrCode", RsSA.Fields("SlTaxCode"), "N", "GodAdd1", "S")
          ConsiAdd2 = GProcGetColumnValue("tblMastNarration", "NarrCode", RsSA.Fields("SlTaxCode"), "N", "GodAdd2", "S")
          ConsiCity = GProcGetColumnValue("tblMastNarration", "NarrCode", RsSA.Fields("SlTaxCode"), "N", "GodCity", "S")
          ConsiPin = GProcGetColumnValue("tblMastNarration", "NarrCode", RsSA.Fields("SlTaxCode"), "N", "ConsiPin", "N")
          ConsiPin = Left(ConsiPin, 6)
          ConsiGstin = UCase(ConsiGstin)
          If ConsiAdd1 = "" Then ConsiAdd1 = "null"
          If ConsiAdd2 = "" Then ConsiAdd2 = "null"
          If ConsiPin = 0 Then
             MsgBox ("Consignee Pin Code Not Found..For +'" & ConsiName & "'   ")
             Close #1
             Exit Sub
          End If
          If ConsiGstin = "" Then
             MsgBox ("Consignee GSTIN Not Found..For +'" & ConsiName & "'   ")
             Close #1
             Exit Sub
          End If
          If ConsiCity = "" Then
             MsgBox ("Consignee City Not Found..For +'" & ConsiName & "'   ")
             Close #1
             Exit Sub
          End If
       
       End If
       
        
       
       Print #1, "{"
       Print #1, Chr(34) + "Version" + Chr(34) + ":" + Chr(34) + "1.1" + Chr(34) + Chr(44)
       Print #1, Chr(34) + "TranDtls" + Chr(34) + ":{"
       Print #1, Space(15) + Chr(34) + "TaxSch" + Chr(34) + ":" + Chr(34) + "GST" + Chr(34) + Chr(44)
       Print #1, Space(15) + Chr(34) + "SupTyp" + Chr(34) + ":" + Chr(34) + "B2B" + Chr(34) + Chr(44)
       Print #1, Space(15) + Chr(34) + "IgstOnIntra" + Chr(34) + ":" + Chr(34) + "N" + Chr(34) + Chr(44)
       Print #1, Space(15) + Chr(34) + "RegRev" + Chr(34) + ":" + "null" + Chr(44)
       Print #1, Space(15) + Chr(34) + "EcmGstin" + Chr(34) + ":" + "null"
       Print #1, "}" + Chr(44)
       
       Print #1, Chr(34) + "DocDtls" + Chr(34) + ":{"
       Print #1, Space(15) + Chr(34) + "Typ" + Chr(34) + ":" + Chr(34) + "INV" + Chr(34) + Chr(44)
       Print #1, Space(15) + Chr(34) + "No" + Chr(34) + ":" + Chr(34) + RsSA.Fields("SlBillNo") + Chr(34) + Chr(44)
       Print #1, Space(15) + Chr(34) + "Dt" + Chr(34) + ":" + Chr(34) + CStr(RsSA.Fields("Vdt")) + Chr(34)
       Print #1, "}" + Chr(44)
      
       Print #1, Chr(34) + "SellerDtls" + Chr(34) + ":{"
       Print #1, Space(15) + Chr(34) + "Gstin" + Chr(34) + ":" + Chr(34) + "" & FirmGstin & "" + Chr(34) + Chr(44)
       Print #1, Space(15) + Chr(34) + "LglNm" + Chr(34) + ":" + Chr(34) + "" & UCase(gCName) & "" + Chr(34) + Chr(44)
       Print #1, Space(15) + Chr(34) + "TrdNm" + Chr(34) + ":" + "null" + Chr(44)
       Print #1, Space(15) + Chr(34) + "Addr1" + Chr(34) + ":" + Chr(34) + "" & gCAdd1 & "" + Chr(34) + Chr(44)
       Print #1, Space(15) + Chr(34) + "Addr2" + Chr(34) + ":" + Chr(34) + "" & gCAdd3 & "" + Chr(34) + Chr(44)
       If gCIBOffice = "I" Then
          Print #1, Space(15) + Chr(34) + "Loc" + Chr(34) + ":" + Chr(34) + "Ichalkaranji" + Chr(34) + Chr(44)
          Print #1, Space(15) + Chr(34) + "Pin" + Chr(34) + ":" + "416115" + Chr(44)
       Else
          Print #1, Space(15) + Chr(34) + "Loc" + Chr(34) + ":" + Chr(34) + "Mumbai" + Chr(34) + Chr(44)
          Print #1, Space(15) + Chr(34) + "Pin" + Chr(34) + ":" + "400009" + Chr(44)
       End If
       Print #1, Space(15) + Chr(34) + "Stcd" + Chr(34) + ":" + Chr(34) + "" & Left(FirmGstin, 2) & "" + Chr(34) + Chr(44)
       Print #1, Space(15) + Chr(34) + "Ph" + Chr(34) + ":" + "null" + Chr(44)
       Print #1, Space(15) + Chr(34) + "Em" + Chr(34) + ":" + "null"
       Print #1, "}" + Chr(44)
      
       Print #1, Chr(34) + "BuyerDtls" + Chr(34) + ":{"
       Print #1, Space(15) + Chr(34) + "Gstin" + Chr(34) + ":" + Chr(34) + "" & PartyGstin & "" + Chr(34) + Chr(44)
       Print #1, Space(15) + Chr(34) + "LglNm" + Chr(34) + ":" + Chr(34) + "" & PartyName & "" + Chr(34) + Chr(44)
       Print #1, Space(15) + Chr(34) + "TrdNm" + Chr(34) + ":" + "null" + Chr(44)
       Print #1, Space(15) + Chr(34) + "Pos" + Chr(34) + ":" + Chr(34) + "" & Left(PartyGstin, 2) & "" + Chr(34) + Chr(44)
       Print #1, Space(15) + Chr(34) + "Addr1" + Chr(34) + ":" + Chr(34) + "" & PartyAdd1 & "" + Chr(34) + Chr(44)
       Print #1, Space(15) + Chr(34) + "Addr2" + Chr(34) + ":" + Chr(34) + "" & PartyAdd2 & "" + Chr(34) + Chr(44)
       Print #1, Space(15) + Chr(34) + "Loc" + Chr(34) + ":" + Chr(34) + "" & PartyCity & "" + Chr(34) + Chr(44)
       Print #1, Space(15) + Chr(34) + "Pin" + Chr(34) + ":" + "" & PartyPin & "" + Chr(44)
       Print #1, Space(15) + Chr(34) + "Stcd" + Chr(34) + ":" + Chr(34) + "" & Left(PartyGstin, 2) & "" + Chr(34) + Chr(44)
       Print #1, Space(15) + Chr(34) + "Ph" + Chr(34) + ":" + "null" + Chr(44)
       Print #1, Space(15) + Chr(34) + "Em" + Chr(34) + ":" + "null"
       Print #1, "}" + Chr(44)
      
      
       Print #1, Chr(34) + "DispDtls" + Chr(34) + ":" + "null" + Chr(44)
       
       
       If ConsiName <> "-" Then
          Print #1, Chr(34) + "ShipDtls" + Chr(34) + ":{"
          Print #1, Space(15) + Chr(34) + "Gstin" + Chr(34) + ":" + Chr(34) + "" & ConsiGstin & "" + Chr(34) + Chr(44)
          Print #1, Space(15) + Chr(34) + "LglNm" + Chr(34) + ":" + Chr(34) + "" & ConsiName & "" + Chr(34) + Chr(44)
          Print #1, Space(15) + Chr(34) + "TrdNm" + Chr(34) + ":" + "null" + Chr(44)
          Print #1, Space(15) + Chr(34) + "Addr1" + Chr(34) + ":" + Chr(34) + "" & ConsiAdd1 & "" + Chr(34) + Chr(44)
          Print #1, Space(15) + Chr(34) + "Addr2" + Chr(34) + ":" + Chr(34) + "" & ConsiAdd2 & "" + Chr(34) + Chr(44)
          Print #1, Space(15) + Chr(34) + "Loc" + Chr(34) + ":" + Chr(34) + "" & ConsiCity & "" + Chr(34) + Chr(44)
          Print #1, Space(15) + Chr(34) + "Pin" + Chr(34) + ":" + "" & ConsiPin & "" + Chr(44)
          Print #1, Space(15) + Chr(34) + "Stcd" + Chr(34) + ":" + Chr(34) + "" & Left(ConsiGstin, 2) & "" + Chr(34)
          Print #1, "}" + Chr(44)
       
       Else
          Print #1, Chr(34) + "ShipDtls" + Chr(34) + ":" + "null" + Chr(44)
       End If
       Print #1, Chr(34) + "ValDtls" + Chr(34) + ":{"
       Print #1, Space(15) + Chr(34) + "AssVal" + Chr(34) + ":" + Trim(CStr(Format(RsSA.Fields(54) + RsSA.Fields("SlAdd1"), "############0.00"))) + Chr(44)
       Print #1, Space(15) + Chr(34) + "IgstVal" + Chr(34) + ":" + Trim(CStr(Format(RsSA.Fields("SlTaxAmt3"), "############0.00"))) + Chr(44)
       Print #1, Space(15) + Chr(34) + "CgstVal" + Chr(34) + ":" + Trim(CStr(Format(RsSA.Fields("SlTaxAmt"), "############0.00"))) + Chr(44)
       Print #1, Space(15) + Chr(34) + "SgstVal" + Chr(34) + ":" + Trim(CStr(Format(RsSA.Fields("SlTaxAmt2"), "############0.00"))) + Chr(44)
       Print #1, Space(15) + Chr(34) + "CesVal" + Chr(34) + ":" + Trim(CStr(Format(0, "############0.00"))) + Chr(44)
       Print #1, Space(15) + Chr(34) + "StCesVal" + Chr(34) + ":" + Trim(CStr(Format(0, "############0.00"))) + Chr(44)
       Print #1, Space(15) + Chr(34) + "Discount" + Chr(34) + ":" + Trim(CStr(Format(0, "############0.00"))) + Chr(44)
       Print #1, Space(15) + Chr(34) + "OthChrg" + Chr(34) + ":" + Trim(CStr(Format(RsSA.Fields("SlTcsAmt"), "############0.00"))) + Chr(44)
       Print #1, Space(15) + Chr(34) + "RndOffAmt" + Chr(34) + ":" + Trim(CStr(Format(RsSA.Fields("SlRoff"), "############0.00"))) + Chr(44)
       Print #1, Space(15) + Chr(34) + "TotInvVal" + Chr(34) + ":" + Trim(CStr(Format(RsSA.Fields("SlBillAmt"), "############0.00")))
       Print #1, "}" + Chr(44)
       
      
       Print #1, Space(25) + Chr(34) + "ItemList" + Chr(34) + ":[{"
       Print #1, Space(32) + Chr(34) + "SlNo" + Chr(34) + ":" + Chr(34) + "1" + Chr(34) + Chr(44)
       Print #1, Space(32) + Chr(34) + "PrdDesc" + Chr(34) + ":" + Chr(34) + "YARN" + Chr(34) + Chr(44)
       Print #1, Space(32) + Chr(34) + "IsServc" + Chr(34) + ":" + Chr(34) + "N" + Chr(34) + Chr(44)
       Print #1, Space(32) + Chr(34) + "HsnCd" + Chr(34) + ":" + Chr(34) + Trim(ItemHsn) + Chr(34) + Chr(44)
       Print #1, Space(32) + Chr(34) + "Qty" + Chr(34) + ":" + Trim(CStr(Format(RsSA.Fields("SlSubWt"), "############0.000"))) + Chr(44)
       Print #1, Space(32) + Chr(34) + "Unit" + Chr(34) + ":" + Chr(34) + "KGS" + Chr(34) + Chr(44)
       Print #1, Space(32) + Chr(34) + "UnitPrice" + Chr(34) + ":" + Trim(CStr(Format(RsSA.Fields("SlSubNetRate") / RsSA.Fields("SlSubRtPer"), "############0.000"))) + Chr(44)
       Print #1, Space(32) + Chr(34) + "TotAmt" + Chr(34) + ":" + Trim(CStr(Format(RsSA.Fields(54), "############0.00"))) + Chr(44)
       Print #1, Space(32) + Chr(34) + "Discount" + Chr(34) + ":" + Trim(CStr(Format(0, "############0.00"))) + Chr(44)
       Print #1, Space(32) + Chr(34) + "PreTaxVal" + Chr(34) + ":" + Trim(CStr(Format(0, "############0.00"))) + Chr(44)
       Print #1, Space(32) + Chr(34) + "AssAmt" + Chr(34) + ":" + Trim(CStr(Format(RsSA.Fields(54), "############0.00"))) + Chr(44)
       Print #1, Space(32) + Chr(34) + "GstRt" + Chr(34) + ":" + Trim(CStr(Format(RsSA.Fields("SlTaxRate") + RsSA.Fields("SlTaxRate2") + RsSA.Fields("SlTaxRate3"), "############0"))) + Chr(44)
       Print #1, Space(32) + Chr(34) + "IgstAmt" + Chr(34) + ":" + Trim(CStr(Format(RsSA.Fields("SlTaxAmt3"), "############0.00"))) + Chr(44)
       Print #1, Space(32) + Chr(34) + "CgstAmt" + Chr(34) + ":" + Trim(CStr(Format(RsSA.Fields("SlTaxAmt"), "############0.00"))) + Chr(44)
       Print #1, Space(32) + Chr(34) + "SgstAmt" + Chr(34) + ":" + Trim(CStr(Format(RsSA.Fields("SlTaxAmt2"), "############0.00"))) + Chr(44)
       Print #1, Space(32) + Chr(34) + "CesRt" + Chr(34) + ":" + Trim(CStr(Format(0, "############0.00"))) + Chr(44)
       Print #1, Space(32) + Chr(34) + "CesAmt" + Chr(34) + ":" + Trim(CStr(Format(0, "############0.00"))) + Chr(44)
       Print #1, Space(32) + Chr(34) + "CesNonAdvlAmt" + Chr(34) + ":" + Trim(CStr(Format(0, "############0.00"))) + Chr(44)
       Print #1, Space(32) + Chr(34) + "StateCesRt" + Chr(34) + ":" + Trim(CStr(Format(0, "############0.00"))) + Chr(44)
       Print #1, Space(32) + Chr(34) + "StateCesAmt" + Chr(34) + ":" + Trim(CStr(Format(0, "############0.00"))) + Chr(44)
       Print #1, Space(32) + Chr(34) + "StateCesNonAdvlAmt" + Chr(34) + ":" + Trim(CStr(Format(0, "############0.00"))) + Chr(44)
       Print #1, Space(32) + Chr(34) + "OthChrg" + Chr(34) + ":" + Trim(CStr(Format(0, "############0.00"))) + Chr(44)
       Print #1, Space(32) + Chr(34) + "TotItemVal" + Chr(34) + ":" + Trim(CStr(Format(RsSA.Fields(54) + RsSA.Fields("SlTaxAmt3") + RsSA.Fields("SlTaxAmt") + RsSA.Fields("SlTaxAmt2"), "############0.00")))
       Print #1, Space(30) + "}"
       Print #1, Space(30) + "]"
       RsSA.MoveNext
       If RsSA.EOF Then
          Print #1, Space(15) + "}"
       Else
          Print #1, "}" + Chr(44)
       End If
    Loop
    Print #1, "]"
    Close #1
    MsgBox ("e-Invoice Jason File Created....")
     Dim RetVal
'--Hitrix
'     RetVal = Shell("C:\Program Files (x86)\Google\Chrome\Application\chrome.exe https://einvoice1.gst.gov.in", vbMaximizedFocus)
'--VTC ICH
'     RetVal = Shell("C:\Program Files (x86)\Google\Chrome\Application\chrome.exe https://einvoice1.gst.gov.in", vbMaximizedFocus)
Case 5:
    Dim excel_app As Excel.Application
    Dim workbook As Excel.workbook
    Dim sheet_name As String
    Dim sheet As Excel.Worksheet
    Dim header_range As Excel.Range

    Dim flnm As String
    Dim a As Integer
    Dim PtyCd As Long
    
    flnm = "d:\UploadedInvoiceDetails"
    Set excel_app = New Excel.Application
    excel_app.Visible = False
    Set workbook = excel_app.Workbooks.Open(filename:=flnm)
    Set sheet = workbook.Sheets.Item("UploadedInvoiceDetails")
    If (sheet Is Nothing) Then
        MsgBox "Error Sales Sheet"
    End If
    a = 2
    Do While sheet.cells(a, 1) <> ""
       X = "select * from tblMastAccount where AcGSTIN = '" & sheet.cells(a, 9) & "'"
       i = GProcRstOpen(RsSA, X, "R")
       PtyCd = GProcGetColumnValue("tblMastAccount", "AcGSTIN", sheet.cells(a, 9), "S", "AcCode", "N")
       If i > 1 Then
          gCn.Execute "Update tblSale set SlIRNNo = '" & sheet.cells(a, 2) & "',SlAckNo = '" & sheet.cells(a, 3) & "'  where  vtype='" & frmTypeStr & "' and SlBillNo= '" & sheet.cells(a, 5) & "' and " _
          & " vfirm = '" & gCCode & "' and VYear=" & gCYear
       Else
          gCn.Execute "Update tblSale set SlIRNNo = '" & sheet.cells(a, 2) & "',SlAckNo = '" & sheet.cells(a, 3) & "'  where SlAcDrCode= " & PtyCd & " and vtype='" & frmTypeStr & "' and SlBillNo= '" & sheet.cells(a, 5) & "' and " _
          & " vfirm = '" & gCCode & "' and VYear=" & gCYear
       End If
       a = a + 1
    Loop
    workbook.Close SaveChanges:=True
    excel_app.Quit
    rstbl.Requery
    MsgBox ("e-Invoice IRN No Updated....")

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
    frmTypeStr = gfrmTypeStr
    mLastEntryDate = vbNull
    SalesAccode = 0
    If frmTypeStr = "SO" Then  ' --- Consignment Sale
        LblSaleType.Caption = "   Consignment Sale"
        TrnCd = gTranCdSO
    ElseIf frmTypeStr = "SD" Then ' --- Depot Sale
        LblSaleType.Caption = "   Depot Sale"
        TrnCd = gTranCdSD
    ElseIf frmTypeStr = "SY" Then
           SalesAccode = gAcCodeSY
           TrnCd = gTranCdSY
    ElseIf frmTypeStr = "SE" Then
           SalesAccode = gAcCodeSYExempt
           TrnCd = gTranCdSYExempt
    End If
    txtTemp(6) = GProcGetColumnValue("tblMastAccount", "AcCode", CStr(SalesAccode), "N", "AcName", "S")
    FirstTimeFlag = True
    Criteria = "Vtype='" & frmTypeStr & "' and VFirm  = '" & gCCode & "' and VYear=" & gCYear & " and vdt >= CONVERT(datetime, '01.07.2017', 104) "
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
    '---- Button Visiblity Set As Per User Previlage
    GProcSetButtonVisiblity Me
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
X = "select SlSubItSrNo,SlSubBookNo,SlSubBookDt,ItName,Narration,SlSubLotNo,SlSubBag,ItStdpack,SlSubWt" _
& " ,SlSubRt,SlSubRtPer,SlSubNetRate,SlSubAmt,SlSubItCtrlNo,SlSubPItCtrlNo,SlSubBkType,SlSubBkVyear,SlSubBkSrNo from tblSaleSub,TblMastNarration,TblMastItem where " _
& " tblSaleSub.SlSubItCode=TblMastItem.ItCode and tblSaleSub.SlSubGodown=TblMastNarration.NarrCode " _
& " and tblSaleSub.vtype='" & frmTypeStr & "' and tblSaleSub.vno=" & txt(0) & " and tblSaleSub.VFirm = '" & gCCode & "' and tblSaleSub.VYear=" & gCYear & " order by SlSubItSrNo"
GProcRstOpen rstblSub, X, "R"
With msGrid
   .FormatString = ">Sr|<Bk.No    |<Book Date  |<Count                                    |<Godown          |<Lot No            |>Bag  |>Std.Pack.|>Weight          |>Rate                |>Per     |>Net Rate           |>Amount               |||||"
   .ColWidth(1) = 0
   .ColWidth(2) = 0
'   .ColWidth(4) = 0
'   .ColWidth(5) = 0
    For i = 13 To 17
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
SetControlEd
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
'    If (dtpDt(39) = "01/04/1899") Then dtpDt(39) = ""
    If dtpDt(2) < CDate("01/07/2017") Then
       MsgBox "Date in VAT Applicabel " & txt(9), vbCritical + vbOKOnly, ""
       dtpDt(Index).SetFocus
       Exit Sub
    End If
End If
'If Index = 2 Then ShowAccountBalance

'If IsDate(mskDt(Index)) = False Then mskDt(Index) = "__/__/____"
If Index = 2 Then
   Dim Rs1 As Recordset
   Dim LastBillNo As String
   Dim MilCd As Long
   If FormAction = vbDataActionAddNew Then
        MilCd = GProcGetColumnValue("tblMastAccount", "AcName", txtTemp(7), "S", "AcCode", "N")
        If Val(txt(9)) > 0 Then
           LastBillNo = Val(txt(9)) - 1
        Else
              LastBillNo = Left(txt(9), 4) + Trim(CStr(Val(Mid(txt(9), 5, 15)) - 1))
        End If
        
        If Val(txtSlSub(17)) = 0 Then
            If GProcRstOpen(Rs1, "Select SlSubItCode,SlSubRt from tblSale,tblSaleSub where SlMillCode = " & MilCd & "  and tblsale.vtype = '" & frmTypeStr & "' and tblsale.Vfirm = '" & gCCode & "' and tblsale.vyear = " & gCYear & " and SlBillNo= '" & LastBillNo & "' and tblSale.vno = tblSaleSub.vno and tblSale.vtype = tblSaleSub.vtype and tblSale.vyear = tblSaleSub.vyear and tblSale.vfirm = tblSaleSub.vfirm", "R") > 0 Then
                msGrid.TextMatrix(1, 3) = GProcGetColumnValue("tblMastItem", "ItCode", Rs1.Fields(0), "N", "ItName", "S")
                msGrid.TextMatrix(1, 7) = GProcGetColumnValue("tblMastItem", "ItCode", Rs1.Fields(0), "N", "ItStdpack", "N")
                msGrid.TextMatrix(1, 9) = Rs1.Fields(1)
                msGrid.TextMatrix(1, 10) = GProcGetColumnValue("tblMastItem", "ItCode", Rs1.Fields(0), "N", "ItStdrateper", "N")
            End If
        End If
   End If
   msGrid.Col = 3
   msGrid_EnterCell
End If
End Sub
Private Sub ShowAccountBalance()
'Dim DrAcBal  As Double
'Dim CrAcBal As Double
'If IsDate(dtpDt(2)) Then
'    DrAcBal = GProcGetAccountBalance(txttemp(5), dtpDt(2))
''    CrAcBal = GProcGetAccountBalance(txttemp(5), dtpDt(2))
'    txtDrBalance = IIf(DrAcBal > 0, DrAcBal & " Cr", Abs(DrAcBal) & " Dr")
''    txtCrBalance = IIf(CrAcBal > 0, CrAcBal & " Cr", Abs(CrAcBal) & " Dr")
'End If
End Sub
Private Sub txt_GotFocus(Index As Integer)
GProcSelectBox Me.ActiveControl
End Sub
Private Sub CheckForNumber()
GProcCheckForNumber Me, rstbl, MaxNo, Array(9), Array(59, 66, 67), pOtherTxtFromNo:=OtherTxtFromNo, pOtherTxtToNo:=OtherTxtToNo
End Sub
Private Sub txt_LostFocus(Index As Integer)
If Index = 38 Then
 '  GatePassData
End If
If Index = 9 And FormAction <> vbDataActionCancel Then
   Dim X As String
   Dim rstAddCheck As Recordset
   X = "Select SlBillNo From " & tblName & " Where SlBillNo = '" & txt(9) & "'  and Vtype='" & frmTypeStr & "'  and Vfirm = '" & gCCode & "' and  VYear=" & gCYear
        i = GProcRstOpen(rstAddCheck, X, "R")
        If i > 0 Then
            MsgBox "Record is Available of Bill No. " & txt(9), vbCritical + vbOKOnly, "Duplicate Record"
            txt(9).SetFocus
        End If
End If
CheckForNumber
SaveBtnEd
CalAmount
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
Dim tmpBillSr As String
Dim NarrCdBlTp As Long
Dim BilSr As String
Dim wIsNoWork As Integer
Dim MilTpCd As Long
Dim X As String
SetControlEd
'If cmdBtn_Click(Index) <> 14 Then
Select Case Index
    Case 5  '-- Party  Code
         txtTemp(5) = GProcGetColumnValue("tblMastAccount", "AcOurCode", UCase(txtCode(5)), "S", "AcName", "S")
         wIsNoWork = GProcGetColumnValue("tblMastAccount", "AcOurCode", UCase(txtCode(5)), "S", "AcIsNonWork", "N")
         lblTinNo.Caption = GProcGetColumnValue("TblMastAccount", "ACName", txtTemp(Index), "S", "AcGSTIN", "S")
         lblMblNoParty.Caption = GProcGetColumnValue("TblMastAccount", "ACName", txtTemp(Index), "S", "AcContNo", "S")
'         lblAdd1.Caption = GProcGetColumnValue("TblMastAccount", "ACName", txttemp(Index), "S", "AcAdd1", "S")
'         lblAdd2.Caption = GProcGetColumnValue("TblMastAccount", "ACName", txttemp(Index), "S", "AcAdd2", "S")
         '--- Broker From Master
         If txtTemp(5) <> "" Then
            If wIsNoWork Then
               Frame2.Visible = True
               cmdClose.SetFocus
            Else
                txt(5) = GProcGetColumnValue("tblMastAccount", "AcName", txtTemp(5), "S", "AcCode", "N")
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
         If txtTemp(8) = "" Then
            BrkAcCode = GProcGetColumnValue("TblMastAccount", "ACName", txtTemp(Index), "S", "AcBrkCode", "N")
            If BrkAcCode = 1 Then
               txtTemp(8) = ""
               txtCode(8) = ""
            Else
               txtTemp(8) = GProcGetColumnValue("TblMastAccount", "ACCode", CStr(BrkAcCode), "N", "ACname", "S")
               txtCode(8) = GProcGetColumnValue("TblMastAccount", "ACCode", CStr(BrkAcCode), "N", "AcOurCode", "S")
            End If
         End If
         If Trim(txtTemp(5)) = "" Then
            If txtCode(5) <> "D" Then
               MsgBox ("Code Not Found ...")
            End If
            txtTemp(5) = ""
            txtTemp(5).Enabled = True
            txtTemp(5).SetFocus
         Else
            txtCode(8).SetFocus
         End If
    Case 7  '-- Mill  Code
        Select Case frmTypeStr
        Case "SY", "SD"
             If frmTypeStr = "SY" Then
                 MilTpCd = GProcGetColumnValue("tblMastNarration", "Narration", "Trade", "S", "Narrcode", "N")
             Else
                 MilTpCd = GProcGetColumnValue("tblMastNarration", "Narration", "Depot", "S", "Narrcode", "N")
             End If
             X = "Select Acname from tblMastAccount,tblMastBillSerial where tblMastaccount.Accode = tblMastBillSerial.MillCode and  tblMastBillSerial.SaleType = " & MilTpCd & "  and tblMastBillSerial.Vfirm = '" & gCCode & "'   and  tblMastaccount.AcOurCode = '" & txtCode(7) & "'   "
              i = GProcRstOpen(tmpRst, X, "R")
              If i = 0 Then
                 MsgBox ("Mill Not in Firm")
                 txtCode(7).SetFocus
                 Exit Sub
              End If
       End Select
              
         
         
         
         
         
         txtTemp(7) = GProcGetColumnValue("tblMastAccount", "AcOurCode", UCase(txtCode(7)), "S", "AcName", "S")
         If Trim(txtTemp(7)) = "" Then
            MsgBox ("Code Not Found ...")
            txtTemp(7) = ""
            txtTemp(7).Enabled = True
            txtTemp(7).SetFocus
         Else
            If txtCode(5).Enabled = True Then txtCode(5).SetFocus
         End If
         If (frmTypeStr = "SO" Or frmTypeStr = "SD") And txtTemp(7) <> "" Then  ' --- Consignment Sale
            txtTemp(6) = txtTemp(7)
         End If
    Case 8  '-- Broker  Code
         txtTemp(8) = GProcGetColumnValue("tblMastAccount", "AcOurCode", UCase(txtCode(8)), "S", "AcName", "S")
         wIsNoWork = GProcGetColumnValue("tblMastAccount", "AcName", txtTemp(8), "S", "AcIsNonWork", "N")
         If Trim(txtTemp(8)) = "" Then
            MsgBox ("Code Not Found ...")
            txtTemp(8) = ""
            txtTemp(8).Enabled = True
            txtTemp(8).SetFocus
         Else
            If wIsNoWork Then
               Frame2.Visible = True
               cmdClose.SetFocus
            Else
               txtTemp(22).SetFocus
            End If
'            If txt(9).Enabled = True Then txt(9).SetFocus
         End If
End Select
If txt(9) = "" Then
   If frmTypeStr = "SD" And txtCode(7) = "ISCCGMW" Then
      GBillNoGima
   Else
      GBillNo
   End If
End If
End Sub

Private Sub txtSlSub_LostFocus(Index As Integer)
Dim X As String
If txtTemp(7) = "" And txtTemp(8) = "" And Val(txtSlSub(17)) <> 0 Then
   gCmd.CommandText = "PrcPreparePendingBookingParty"
   gCmd.Parameters.Refresh
   gCmd.Parameters("@Fr_dt") = Format(gCYSDate - 400, "yyyy/mm/dd")
   gCmd.Parameters("@To_dt") = Format(gCYEDate, "yyyy/mm/dd")
   gCmd.Parameters("@UCode") = gUserName
   gCmd.Parameters("@VFirm") = Trim(gCCode)
   gCmd.Parameters("@Vtype") = "OP"
   gCmd.Parameters("@RptTp") = "P"
   gCmd.Parameters("@BkNo") = Val(txtSlSub(17))
   gCmd.Execute
   
'   X = "Select Vno,PartyName,BrokerName,MillCode,Nar1,Vdt,Amt1,Amt3,Amt4,Dt1, " _
'       & "PartyCode,BrokerCode,MillName,Amt6 ,Amt7,Vtype,Vyear,Amt5 from TmpGentbl where TmpGentbl.UserName = '" & gUserName & "' and (amt4 - amt8) > 0   "
   
   X = "Select Vno,PartyName,BrokerName,MillCode,Nar1,Amt3,Amt4-Amt8-Amt2,Dt1, " _
       & "PartyCode,BrokerCode,MillName,Amt6 ,Amt7,Vtype,Vyear,Amt5,Vdt from TmpGentbl where TmpGentbl.UserName = '" & gUserName & "' and (amt4 - amt8) > 0   "
   
   
   i = GProcRstOpen(tmpRst, X, "R")
   If i > 0 Then
'      gClsSearch.SearchMultiField X, "", Array("Booking No", "Party", "Broker", "Mill", "Count", "Date", "Bags", "Rate", "Sh.Bag", "Sh.Date", "", "", "", "Std.Pck", "RtPer", "BkTp", "BkYear", "BkSrNo"), Array(0, 2700, 1800, 1000, 2000, 1500, 800, 1000, 800, 1500, 0, 0, 0, 0, 0, 0, 0, 0), "", "", "Vno,Vdt,dt1 ", 0, 0, False, , True
      gClsSearch.SearchMultiField X, "", Array("Booking No", "Party", "Broker", "Mill", "Count", "Rate", "Sh.Bag", "Sh.Date", "", "", "", "Std.Pck", "RtPer", "BkTp", "BkYear", "BkSrNo", "Bk.Dt"), Array(0, 2700, 1800, 1000, 2000, 1000, 800, 1500, 0, 0, 0, 0, 0, 0, 0, 0, 0), "", "", "Vno,Vdt,dt1 ", 0, 0, False, , True
      If gClsSearch.SearchMultiRetCol(0) <> "" Then
         txtSlSub(19) = gClsSearch.SearchMultiRetCol(13)
         txtSlSub(20) = gClsSearch.SearchMultiRetCol(14)
         txtSlSub(28) = gClsSearch.SearchMultiRetCol(15)
         txtTemp(5) = gClsSearch.SearchMultiRetCol(1)
         txtTemp(8) = gClsSearch.SearchMultiRetCol(2)
         txtCode(7) = gClsSearch.SearchMultiRetCol(3)
         mskDtSlSub(19) = gClsSearch.SearchMultiRetCol(16)
         txtCode(5) = gClsSearch.SearchMultiRetCol(8)
         txtCode(8) = gClsSearch.SearchMultiRetCol(9)
         txtTemp(7) = gClsSearch.SearchMultiRetCol(10)
         msGrid.TextMatrix(1, 1) = gClsSearch.SearchMultiRetCol(0)
         msGrid.TextMatrix(1, 2) = gClsSearch.SearchMultiRetCol(16)
         msGrid.TextMatrix(1, 15) = gClsSearch.SearchMultiRetCol(13)
         msGrid.TextMatrix(1, 16) = gClsSearch.SearchMultiRetCol(14)
         msGrid.TextMatrix(1, 17) = gClsSearch.SearchMultiRetCol(15)
         msGrid.TextMatrix(1, 3) = gClsSearch.SearchMultiRetCol(4)
'         If gClsSearch.SearchMultiRetCol(6) = 0 Then
'             msGrid.TextMatrix(1, 6) = gClsSearch.SearchMultiRetCol(6)
'             msGrid.TextMatrix(1, 8) = gClsSearch.SearchMultiRetCol(6) * gClsSearch.SearchMultiRetCol(13)
'         Else
             msGrid.TextMatrix(1, 6) = gClsSearch.SearchMultiRetCol(6)
             msGrid.TextMatrix(1, 8) = gClsSearch.SearchMultiRetCol(6) * gClsSearch.SearchMultiRetCol(11)
'         End If
         msGrid.TextMatrix(1, 7) = gClsSearch.SearchMultiRetCol(11)
         msGrid.TextMatrix(1, 9) = gClsSearch.SearchMultiRetCol(5)
         msGrid.TextMatrix(1, 10) = gClsSearch.SearchMultiRetCol(12)
      Else
        txtSlSub(17) = ""
      End If
   Else
      txtSlSub(17) = ""
      MsgBox ("Booking No Not Found....")
      If txtCode(7).Enabled = True Then txtCode(7).SetFocus
   End If
End If
End Sub

Private Sub txttemp_GotFocus(Index As Integer)
Dim AcCode As Long
Dim BilSr As String
Dim NarrCdBlTp As Long
GProcSelectBox Me.ActiveControl
End Sub
Private Sub txttemp_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
Dim GName As String
Dim OldFrmType As String
If KeyCode = vbKeyF3 Then
    Select Case Index
        Case 5 '-- Party
            GName = GProcGetColumnValue("TblMastGroup", "AgCode", 90017, "N", "AgName", "S")
            GProcShowForm frmMastAccount, frmMain.mnuMstfrm(0), Array(frmMastAccount.txtTemp(2)), Array(GName)
        Case 6 '-- Sales A/c
            GName = GProcGetColumnValue("TblMastGroup", "AgCode", 90023, "N", "AgName", "S")
            GProcShowForm frmMastAccount, frmMain.mnuMstfrm(0), Array(frmMastAccount.txtTemp(2)), Array(GName)
        Case 7 '-- millcode
            GName = GProcGetColumnValue("TblMastGroup", "AgCode", 90029, "N", "AgName", "S")
            GProcShowForm frmMastAccount, frmMain.mnuMstfrm(0), Array(frmMastAccount.txtTemp(2)), Array(GName)
        Case 8 '-- Broker
           GName = GProcGetColumnValue("TblMastGroup", "AgCode", 90016, "N", "AgName", "S")
           GProcShowForm frmMastAccount, frmMain.mnuMstfrm(0), Array(frmMastAccount.txtTemp(2)), Array(GName)
        Case 12 '--TaxType
            'OldgFrmTypeStr = gfrmTypeStr
            GProcShowForm frmMastNarrationTax, frmMain.mnuMstfrm(6)
            'gfrmTypeStr = OldgFrmTypeStr
        Case 11 '--Narration transport
            OldFrmType = gfrmTypeStr
            gfrmTypeStr = "R"
            GProcShowForm frmMastNarration, frmMain.mnuMstfrm(8), gfrmTypeStr
            gfrmTypeStr = OldFrmType
        Case 22 '--Narration transport
            OldFrmType = gfrmTypeStr
            gfrmTypeStr = "C"
            GProcShowForm frmMastNarration, frmMain.mnuMstfrm(8), gfrmTypeStr
            gfrmTypeStr = OldFrmType
    End Select
End If
Dim LeftPos As Long
Dim TopPos As Long
If KeyCode = vbKeyF2 Then '-- Select
    TopPos = txtTemp(Index).Top + Me.Top + 650
    LeftPos = txtTemp(Index).Left + Me.Left
    Select Case Index
        Case 11 '--Narration transport
            gClsSearch.SearchMultiField "tblMastNarration", "Narration", Array("Narration"), Array(txtTemp(Index).Width), "NarrType='R'", "", "Narration", LeftPos, TopPos, False
            If gClsSearch.SearchMultiRetCol(0) <> "" Then txtTemp(Index) = gClsSearch.SearchMultiRetCol(0)
    End Select
End If
End Sub
Private Sub txtTemp_KeyPress(Index As Integer, KeyAscii As Integer)
Dim LeftPos As Long
Dim TopPos As Long
Dim AgCode As Long
Dim BrkAcCode As Long
Dim ShowMill As String
Dim NarrCdBlTp As Long
Dim BilSr As String
Dim X As String
Dim wIsNoWork As Integer
TopPos = txtTemp(Index).Top + Me.Top + 650
LeftPos = txtTemp(Index).Left + Me.Left
If KeyAscii = 13 Then Exit Sub
Select Case Index
    Case 5  '-- Party   Alias Account
        gClsSearch.SearchMultiFieldAccMaster Chr(KeyAscii), txtTemp(Index).Left + Me.Left, txtTemp(Index).Top + Me.Top + 1000
'        gClsSearch.SearchMultiFieldAccMaster Chr(KeyAscii), LeftPos, TopPos
        KeyAscii = 0
        lblTinNo.Caption = ""
        lblMsOms.Caption = ""
        txtCode(5) = ""
        If gClsSearch.SearchMultiRetColAccMaster(0) <> "" Then
            txtTemp(Index) = gClsSearch.SearchMultiRetColAccMaster(0)
            txtCode(5) = GProcGetColumnValue("TblMastAccount", "ACName", txtTemp(Index), "S", "AcOurCode", "S")
            lblTinNo.Caption = GProcGetColumnValue("TblMastAccount", "ACName", txtTemp(Index), "S", "AcGSTIN", "S")
            lblMsOms.Caption = GProcGetColumnValue("TblMastAccount", "ACName", txtTemp(Index), "S", "AcMsOmsParty", "N")
            If Val(lblMsOms.Caption) = 0 Then
               lblMsOms.Caption = "M.S."
            Else
               lblMsOms.Caption = "O.M.S."
            End If
            lblMblNoParty.Caption = GProcGetColumnValue("TblMastAccount", "ACName", txtTemp(Index), "S", "AcContNo", "S")
'            lblAdd1.Caption = GProcGetColumnValue("TblMastAccount", "ACName", txttemp(Index), "S", "AcAdd1", "S")
 '           lblAdd2.Caption = GProcGetColumnValue("TblMastAccount", "ACName", txttemp(Index), "S", "AcAdd2", "S")
            If txtTemp(5) <> "" Then
               wIsNoWork = GProcGetColumnValue("tblMastAccount", "AcOurCode", UCase(txtCode(5)), "S", "AcIsNonWork", "N")
               If wIsNoWork Then
                  Frame2.Visible = True
                  cmdClose.SetFocus
               Else
                    txt(5) = GProcGetColumnValue("tblMastAccount", "AcName", txtTemp(5), "S", "AcCode", "N")
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
        End If
        '-- Broker from master
        If txtTemp(8) = "" Then
            BrkAcCode = GProcGetColumnValue("TblMastAccount", "ACName", txtTemp(Index), "S", "AcBrkCode", "N")
            txtTemp(8) = GProcGetColumnValue("TblMastAccount", "ACCode", CStr(BrkAcCode), "N", "ACname", "S")
            txtCode(8) = GProcGetColumnValue("TblMastAccount", "ACCode", CStr(BrkAcCode), "N", "AcOurCode", "S")
            If Val(txtCode(8)) = 0 Then
               txtCode(8) = ""
            Else
               wIsNoWork = GProcGetColumnValue("tblMastAccount", "AcName", txtTemp(8), "S", "AcIsNonWork", "N")
               If wIsNoWork Then
                  Frame2.Visible = True
                  cmdClose.SetFocus
               End If
            End If
        End If
    Case 6 '-- Sales A/c  Alias Account
        gClsSearch.SearchMultiField "tblMastaccount", "AcName,AcCity", Array("Sales Account", "City"), Array(3000, 2000), "(Agcode=90023)", Chr(KeyAscii), "", LeftPos, TopPos, True
        KeyAscii = 0
        If gClsSearch.SearchMultiRetCol(0) <> "" Then
            txtTemp(Index) = gClsSearch.SearchMultiRetCol(0)
        End If
    Case 7  '-- Mill  Alias Account
          Dim MilTpCd As Long
        Select Case frmTypeStr
        Case "SY", "ST", "SE"
             MilTpCd = GProcGetColumnValue("tblMastNarration", "Narration", "Trade", "S", "Narrcode", "N")
             If frmTypeStr = "SE" Then
                ShowMill = "AgCode=90029 and AcIsExemptMill = 1 and isnull(AcConsignmentFirm,'') = '' and isnull(AcDepotFirm,'') = ''"
             Else
'                ShowMill = "AgCode=90029 and isnull(AcConsignmentFirm,'') = '' and isnull(AcDepotFirm,'') = ''"
                ShowMill = "AgCode=90029 and   tblMastaccount.Accode = tblMastBillSerial.MillCode and  tblMastBillSerial.SaleType = " & MilTpCd & "  and tblMastBillSerial.Vfirm = '" & gCCode & "'  "
             End If
             gClsSearch.SearchMultiField "tblMastaccount,tblMastBillSerial", "BillSr,AcName,AcOurCode", Array("Mill Ini.", "Mill Name", "Code"), Array(1000, 8000, 1000), ShowMill, Trim(Chr(KeyAscii)), "AcName", LeftPos, TopPos, True
'            gClsSearch.SearchMultiField "tblMastaccount", "AcName,AcCity,AcOurCode", Array("Mill Name", "City", ""), Array(8000, 2000, 0), ShowMill, Trim(Chr(KeyAscii)), "AcName", LeftPos, TopPos, True
            KeyAscii = 0
            txtCode(7) = ""
            If gClsSearch.SearchMultiRetCol(1) <> "" Then
                txtTemp(Index) = gClsSearch.SearchMultiRetCol(1)
                txtCode(7) = gClsSearch.SearchMultiRetCol(2)
                If (frmTypeStr = "SO" Or frmTypeStr = "SD") And txtTemp(7) <> "" Then  ' --- Consignment Sale
                    txtTemp(6) = gClsSearch.SearchMultiRetCol(1)
                End If
            End If
        Case "SO"
            gClsSearch.SearchMultiField "qryMillHelpWitIni", "BillSr,MillName,AcOurCode", Array("Mill Ini.", "Mill Name", "Code"), Array(1000, 8000, 1000), "SaleType = '27' and vfirm = '" & gCCode & "' ", Trim(Chr(KeyAscii)), "MillName", LeftPos, TopPos, True
            KeyAscii = 0
            txtCode(7) = ""
            If gClsSearch.SearchMultiRetCol(1) <> "" Then
                txtTemp(Index) = gClsSearch.SearchMultiRetCol(1)
                txtCode(7) = gClsSearch.SearchMultiRetCol(2)
                txtTemp(6) = gClsSearch.SearchMultiRetCol(1)
            End If
        Case "SD"
             MilTpCd = GProcGetColumnValue("tblMastNarration", "Narration", "Depot", "S", "Narrcode", "N")
             ShowMill = "AgCode=90029 and   tblMastaccount.Accode = tblMastBillSerial.MillCode and  tblMastBillSerial.SaleType = " & MilTpCd & "  and tblMastBillSerial.Vfirm = '" & gCCode & "'  "
             gClsSearch.SearchMultiField "tblMastaccount,tblMastBillSerial", "BillSr,AcName,AcOurCode", Array("Mill Ini.", "Mill Name", "Code"), Array(1000, 8000, 1000), ShowMill, Trim(Chr(KeyAscii)), "AcName", LeftPos, TopPos, True
            
 '           gClsSearch.SearchMultiField "qryMillHelpWitIni", "BillSr,MillName,AcOurCode", Array("Mill Ini.", "Mill Name", "Code"), Array(1000, 8000, 1000), "SaleType = '28' and vfirm = '" & gCCode & "' ", Trim(Chr(KeyAscii)), "MillName", LeftPos, TopPos, True
            KeyAscii = 0
            txtCode(7) = ""
            If gClsSearch.SearchMultiRetCol(1) <> "" Then
                txtTemp(Index) = gClsSearch.SearchMultiRetCol(1)
                txtCode(7) = gClsSearch.SearchMultiRetCol(2)
                txtTemp(6) = gClsSearch.SearchMultiRetCol(1)
            End If
        End Select
        MillCtrlEd
    Case 8 '-- Broker
        gClsSearch.SearchMultiField "tblMastaccount", "AcName,AcOurCode,AcCity", Array("Broker", "Code", "City"), Array(5000, 1000, 3000), "(AgCode=90016 )", Trim(Chr(KeyAscii)), "", LeftPos, TopPos, True
        KeyAscii = 0
        txtCode(8) = ""
        If gClsSearch.SearchMultiRetCol(0) <> "" Then
            txtTemp(Index) = gClsSearch.SearchMultiRetCol(0)
            txtCode(8) = gClsSearch.SearchMultiRetCol(1)
            wIsNoWork = GProcGetColumnValue("tblMastAccount", "AcName", txtTemp(8), "S", "AcIsNonWork", "N")
            If wIsNoWork Then
               Frame2.Visible = True
               cmdClose.SetFocus
            End If
        End If
End Select
If Index = 11 Then '-- Transport
    gClsSearch.SearchMultiField "tblMastNarration", "Narration", Array("Transport"), Array(txtTemp(Index).Width), " NarrType='R'", Chr(KeyAscii), "Narration", 0, 0, True
    KeyAscii = 0
    If gClsSearch.SearchMultiRetCol(0) <> "" Then
        txtTemp(Index) = gClsSearch.SearchMultiRetCol(0) '--Transport
    End If
End If
If Index = 22 Then '-- Consignee
    gClsSearch.SearchMultiField "tblMastNarration", "Narration", Array("Consignee Name"), Array(txtTemp(Index).Width), " NarrType='C'", Chr(KeyAscii), "Narration", 0, 0, True
    KeyAscii = 0
    If gClsSearch.SearchMultiRetCol(0) <> "" Then
        txtTemp(Index) = gClsSearch.SearchMultiRetCol(0) '--Transport
    End If
End If
End Sub
Private Sub MillCtrlEd()
Dim ExemptMill As Long
'-- Is Exempt Mill
ExemptMill = GProcGetColumnValue("TblMastAccount", "AcName", txtTemp(7), "S", "ACIsExemptMill", "N")
If frmTypeStr = "SE" Then
    If FormAction = vbDataActionAddNew Or FormAction = vbDataActionUpdate Then
        txt(19).Enabled = True
        txt(20).Enabled = True
        LblExmptAmt.Visible = True
        LblExmptKg.Visible = True
        txt(19).Visible = True
        txt(20).Visible = True
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
Dim AcCode As Long
Dim BilSr As String
Dim NarrCdBlTp As Long
Dim wIsNoWork As Integer
SaveBtnEd
MillCtrlEd
ShowAccountBalance
If txt(9) = "" Then
   If frmTypeStr = "SD" And txtCode(7) = "ISCCGMW" Then  '"ADMMGMW"
      GBillNoGima
   Else
      GBillNo
   End If
End If
''If Index = 5 Then
''   wIsNoWork = GProcGetColumnValue("tblMastAccount", "AcName", txttemp(5), "S", "AcIsNonWork", "N")
''   If wIsNoWork Then
''      Frame2.Visible = True
''      cmdClose.SetFocus
''   End If
''End If
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
                    txtGrid.Locked = True
                Case 3 '-- Count
                    If .TextMatrix(.Row, 1) > 0 Then   '-- Book No
                        txtGrid.Locked = True
                    Else
                        txtGrid.Locked = False
                    End If
                Case 4 '-- Godown
                    If Year(gCYSDate) >= 2023 Then
                       txtGrid.Locked = True
                    Else
                       txtGrid.Locked = False
                    End If
                Case 7 '-- Std Pack
                    txtGrid.Locked = True
                Case 10 '-- rate per
                    txtGrid.Locked = True
                Case 11 '-- Net rate
                    txtGrid.Locked = True
                Case 12 '-- Sub Amt rate
                    txtGrid.Locked = True
                Case Else
                    txtGrid.Locked = False
                    If .Col = 6 Then
                       OldBag = .TextMatrix(.Row, 6)
                       OldCount = .TextMatrix(.Row, 3)
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
        If ((Val(.TextMatrix(.Row, 6)) = 0) Or (Val(.TextMatrix(.Row, 12)) = 0)) And .rows > 2 Then
            .RemoveItem (.Row)
            txtGrid.text = ""
        End If
        txtGrid.Visible = False
        txt(31).SetFocus
        Exit Sub
    End If
End With
'If FormAction <> vbDataActionAddNew And FormAction <> vbDataActionUpdate Then Exit Sub
'---- Navigate Row ,Col
If KeyAscii = 13 Then
    KeyAscii = 0
    With msGrid
        If .Col = 9 Then 'rstblSub.Fields.Count - 2 Then
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
'--   If msGrid.Col = 5 Then
'   If msGrid.Col = 4 Or msGrid.Col = 5 Then
'--      msGrid.Col = 6
'--      KeyAscii = 0
'--      If msGrid.Enabled = True Then
'--         msGrid.SetFocus
'--      End If
      '---- For Exemption Stock
      '-- For Purchase Against Sale
      
'      If frmTypeStr = "SE" Or frmTypeStr = "SY" Then
'           Dim MillCode As Long
'           Dim ExRt As Double
'           Dim QStr As String
'           Dim LeftPos As Long
'           Dim TopPos As Long
'           Dim ItCd As Long
'           Dim ModiBags As Integer
'           Dim ModiWt As Double
'           Dim PurTp As String
'           If frmTypeStr = "SE" Then
'              PurTp = "PE"
'           Else
'              PurTp = "('PY','OY')"
'           End If
'           TopPos = txtGrid.Top + Me.Top + 650
'           LeftPos = txtGrid.Left + Me.Left
'           If FormAction = vbDataActionUpdate Then
'              ModiBags = Val(msGrid.TextMatrix(msGrid.Row, 6))
'              ModiWt = Val(msGrid.TextMatrix(msGrid.Row, 7))
'           Else
'              ModiBags = 0
'              ModiWt = 0
'           End If
'           MillCode = GProcGetColumnValue("TblMastAccount", "AcName", txttemp(7), "S", "Accode", "N")
'           ExRt = GProcGetColumnValue("tblMastItem", "Itname", msGrid.TextMatrix(msGrid.Row, 3), "S", "ExemptRt", "N", "ITMillCode=" & MillCode & "")
'           ItCd = GProcGetColumnValue("tblMastItem", "Itname", msGrid.TextMatrix(msGrid.Row, 3), "S", "Itcode", "N", "ITMillCode=" & MillCode & "")
'
'           QStr = "Select iif(tblPurch.vtype = 'OY','Op.Stock',tblPurch.PurBillNo),format(tblPurchSub.Vdt,'dd/MM/yyyy'),tblPurchSub.PurSubBag-tblPurchSub.PurSubIssBag+ " & ModiBags & ",tblPurchSub.PurSubWt-tblPurchSub.PurSubIssWt+" & ModiWt & ",tblPurchSub.PurSubItCtrlNo,tblPurchSub.Vno " _
'                  & " from tblPurch,tblPurchSub where PurSubItCode = " & ItCd & " and tblPurchSub.vdt <=('" & Format(dtpDt(2), "YYYY/MM/DD") & "') and tblPurchSub.PurSubBag-tblPurchSub.PurSubIssBag+" & ModiBags & " > 0 and " _
'                  & " tblPurchSub.VFirm = '" & gCCode & "'  and tblPurchSub.vtype in " & PurTp & " and tblpurch.vno = tblPurchSub.vno and tblpurch.vtype = tblPurchSub.vtype and tblpurch.vyear = tblPurchSub.vyear and tblpurch.vfirm = tblPurchSub.vfirm and tblPurch.vyear = " & gCYear & " "
'           gClsSearch.SearchMultiField QStr, "", Array("Bill No", "Date", "Bal. Bag", "Bal. Wt.", "PurSubItCtrlNo", "Vno"), Array(2000, 1500, 1000, 2000, 0, 0), "", "", "tblPurchSub.Vdt,tblPurchSub.Vno", LeftPos, TopPos, True, gCn, True
'           If gClsSearch.SearchMultiRetCol(0) <> "" Then
'              msGrid.TextMatrix(msGrid.Row, 14) = gClsSearch.SearchMultiRetCol(4)
'              msGrid.TextMatrix(msGrid.Row, 6) = gClsSearch.SearchMultiRetCol(2)
'              txtGrid = gClsSearch.SearchMultiRetCol(2)
'              msGrid.TextMatrix(msGrid.Row, 8) = gClsSearch.SearchMultiRetCol(3)
'              lblBalWt.Caption = gClsSearch.SearchMultiRetCol(5)
'              lblBalBag.Caption = gClsSearch.SearchMultiRetCol(0)
'           End If
'      End If
'--   End If
   If msGrid.Col = 7 Then
      If Val(msGrid.TextMatrix(msGrid.Row, 7)) = 0 Then
         msGrid.Col = 8
      Else
         msGrid.Col = 9
      End If
      KeyAscii = 0
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
If msGrid.Col = 3 And txtGrid <> "" Then txtGrid.SelStart = 0
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
Dim Rs1 As Recordset
Dim VAddWithTax As Double
Dim ItCd As Long
Dim X As String
Dim ExmptKg As Double
Dim ExmptAmt As Double
Dim CalWt As Double
CheckForNumberGrid
If msGrid.TextMatrix(msGrid.Row, 1) = 0 Then msGrid.TextMatrix(msGrid.Row, 2) = "" '-- Book date
MillCode = GProcGetColumnValue("TblMastAccount", "AcName", txtTemp(7), "S", "Accode", "N")
StdPack = GProcGetColumnValue("TblMastItem", "ITName", msGrid.TextMatrix(msGrid.Row, 3), "S", "ItStdpack", "N", " ItMillCode=" & MillCode)
wCharityRt = GProcGetColumnValue("TblMastItem", "ITName", msGrid.TextMatrix(msGrid.Row, 3), "S", "CharityRt", "N", " ItMillCode=" & MillCode)
wCharityOn = GProcGetColumnValue("TblMastItem", "ITName", msGrid.TextMatrix(msGrid.Row, 3), "S", "CharityOn", "N", " ItMillCode=" & MillCode)
ItCd = GProcGetColumnValue("tblMastItem", "ItName", msGrid.TextMatrix(msGrid.Row, 3), "S", "ItCode", "N", " ItMillCode=" & MillCode)


If msGrid.Row = 0 Then Exit Sub
'-- Wt=Bag * Std Pack
'--Weight Calculation only at Addition as per req by Vijay List on 04/11/2019

If ((OldBag <> msGrid.TextMatrix(msGrid.Row, 6) Or OldCount <> msGrid.TextMatrix(msGrid.Row, 3)) And StdPack > 0) And FormAction = vbDataActionAddNew Then
    msGrid.TextMatrix(msGrid.Row, 8) = msGrid.TextMatrix(msGrid.Row, 6) * StdPack
    OldBag = msGrid.TextMatrix(msGrid.Row, 6)
End If

'If OldBag <> msGrid.TextMatrix(msGrid.Row, 6) And StdPack = 0 Then
'    msGrid.TextMatrix(msGrid.Row, 8) = 0
'End If
If CDbl(msGrid.TextMatrix(msGrid.Row, 10)) > 0 Then
    GAmt = Round(CDbl(msGrid.TextMatrix(msGrid.Row, 8)) * CDbl(msGrid.TextMatrix(msGrid.Row, 9)) / CDbl(msGrid.TextMatrix(msGrid.Row, 10)), 2)
Else
    GAmt = 0
End If
If Format(dtpDt(2).Value, "YYYY/MM/dd") >= Format("2021/02/15", "YYYY/MM/dd") Then
   GSubAmt = GAmt
Else
    If frmTypeStr = "SE" And Val(msGrid.TextMatrix(msGrid.Row, 14)) > 0 Then
       X = "select tblpurch.vno,tblpurch.vtype,tblpurch.vyear,tblpurch.vfirm,tblPurch.PurExemptPerKg from tblpurch,tblpurchsub where tblpurch.vno = tblpurchsub.vno and tblpurch.vtype = tblpurchsub.vtype and " _
           & " tblpurch.vyear = tblpurchsub.vyear and tblpurch.vfirm = tblpurchsub.vfirm and tblpurch.vfirm = '" & gCCode & "' and tblpurchsub.PurSubItCtrlNo = '" & msGrid.TextMatrix(msGrid.Row, 14) & "'"
       If GProcRstOpen(Rs1, X, "R") > 0 Then
          ExmptKg = Rs1.Fields("PurExemptPerKg")
          ExmptAmt = (ExmptKg * Val(msGrid.TextMatrix(msGrid.Row, 8)))
       End If
       VAddWithTax = GAmt - ExmptAmt
        If VAddWithTax < 0 Then
            VAddWithTax = 0
            ExmptAmt = GAmt
        End If
        GSubAmt = Format(VAddWithTax - ((VAddWithTax * 100) / (100 + (Val(txt(23)) + Val(txt(26)) + Val(txt(41))))), LStr)
        GSubAmt = Round(GAmt - GSubAmt, 2)
    Else
       '--Sub Amt=(Amt)*100/(100+Tax %)
       GSubAmt = Format(((GAmt * 100) / (100 + (Val(txt(23)) + Val(txt(26)) + Val(txt(41))))), LStr)
    End If
End If
'- 8  - Weitht
'- 9  - Rate
'- 10  - Per
'- 11 - Net Rate
'- 12 - Sub Amt
If wCharityRt > 0 Then
'And lblCharity.Visible = False Then
   lblCharity.Visible = True
   txt(17).Visible = True
   txt(43).Visible = True
   Cbo(44).Visible = True
   
Else
   lblCharity.Visible = False
   txt(17).Visible = False
   txt(43).Visible = False
   Cbo(44).Visible = False
End If

If wCharityRt > 0 Then
   If wCharityOn = 0 Then   '-- On Kg
      wCharityAmt = Format(Val(msGrid.TextMatrix(msGrid.Row, 8)) * wCharityRt, FStr)
   Else                     '-- On Pkg
      wCharityAmt = Format(Val(msGrid.TextMatrix(msGrid.Row, 6)) * wCharityRt, FStr)
   End If
   txt(43) = wCharityRt
   Cbo(44).ListIndex = wCharityOn
Else
   wCharityAmt = 0
End If
If Format(dtpDt(2).Value, "YYYY/MM/dd") >= Format("2021/02/15", "YYYY/MM/dd") Then
   NetRt = Val(msGrid.TextMatrix(msGrid.Row, 9))
Else
    If frmTypeStr = "SD" And txtCode(7) = "ISCCGMW" Then
       NetRt = Val(msGrid.TextMatrix(msGrid.Row, 9))
    Else
        GSubAmt = GSubAmt - wCharityAmt
        If CDbl(msGrid.TextMatrix(msGrid.Row, 8)) > 0 Then
           NetRt = Format((GSubAmt / CDbl(msGrid.TextMatrix(msGrid.Row, 8))) * CDbl(msGrid.TextMatrix(msGrid.Row, 10)), LStr)
        Else
           NetRt = 0
        End If
    End If
End If
msGrid.TextMatrix(msGrid.Row, 11) = NetRt
'-- SubAmt=Wt*NetRate/Rate Per
If CDbl(msGrid.TextMatrix(msGrid.Row, 10)) > 0 Then
    msGrid.TextMatrix(msGrid.Row, 12) = CDbl(msGrid.TextMatrix(msGrid.Row, 8)) * CDbl(msGrid.TextMatrix(msGrid.Row, 11)) / CDbl(msGrid.TextMatrix(msGrid.Row, 10))
Else
    msGrid.TextMatrix(msGrid.Row, 12) = 0
End If
msGrid.TextMatrix(msGrid.Row, 12) = Format(msGrid.TextMatrix(msGrid.Row, 12), LStr)
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
        GProcShowForm frmMastItem, frmMain.mnuMstfrm(3), Array(frmMastItem.txtTemp(9)), Array(txtTemp(7))
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
Dim Tp As String
Dim MillCode As Long
Dim ItCd As Long
Dim BalBag As Long
Dim BalWt As Double
Dim QStr As String
Dim tmpRst As Recordset
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
                PrepareBooking Chr(KeyAscii)
                KeyAscii = 0
            Case 3 '-- Count
                If msGrid.TextMatrix(msGrid.Row, 1) > 0 Then
                    KeyAscii = 0
                Else
                    MillCode = GProcGetColumnValue("TblMastAccount", "AcName", txtTemp(7), "S", "Accode", "N")
                    gClsSearch.SearchMultiField "tblMastItem,tblMastaccount", "ItName,AcName,ItStdpack,ItStdrateper,ItCode,ExemptRt,InDailyReport", Array("Count Name", "Mill", "Std. Pack", "Rate Per", "ItCode", "ExemptRt", ""), Array(5000, 5000, 1000, 1000, 0, 0, 0), "tblMastItem.ITMillCode=tblMastaccount.acCode and ITMillCode=" & MillCode, Chr(KeyAscii), "ItName", LeftPos, TopPos, True
                    If gClsSearch.SearchMultiRetCol(0) = "" Then
                       Exit Sub
                    End If
                    If gClsSearch.SearchMultiRetCol(6) = 1 Then
                       MsgBox "This is No Use Count "
                    End If
                    txtGrid = ""
                    KeyAscii = 0
                    txtGrid = ""
                    KeyAscii = 0
                    If gClsSearch.SearchMultiRetCol(0) <> "" And gClsSearch.SearchMultiRetCol(6) = 0 Then '-- Item Wt ,Rate From Master
                        txtGrid = gClsSearch.SearchMultiRetCol(0)
                        .TextMatrix(.Row, 7) = gClsSearch.SearchMultiRetCol(2)
                        If IsNumeric(.TextMatrix(.Row, 8)) = False Then .TextMatrix(.Row, 8) = 0
                        .TextMatrix(.Row, 10) = gClsSearch.SearchMultiRetCol(3)
                        ItemType = GProcGetColumnValue("tblMastItem", "ItName", msGrid.TextMatrix(1, 3), "S", "ItType", "N", "ItMillcode=" & MillCode) '--Item Code
                        If ItemType = 0 Then
                           If (Left(gCGStin, 2) = Left(lblTinNo.Caption, 2) Or Left(lblTinNo.Caption, 2) = "" Or Left(lblTinNo.Caption, 1) = "U") Then
                              txt(23) = GProcGetColumnValue("tblMastNarration", "Narration", "G S T", "S", "CotCGSTRt", "N")
                              txt(26) = GProcGetColumnValue("tblMastNarration", "Narration", "G S T", "S", "CotSGSTRt", "N")
                              txt(41) = 0
                            Else
                              txt(23) = 0
                              txt(26) = 0
                              txt(41) = GProcGetColumnValue("tblMastNarration", "Narration", "G S T", "S", "CotIGSTRt", "N")
                            End If
                        Else
                            If (Left(gCGStin, 2) = Left(lblTinNo.Caption, 2) Or Left(lblTinNo.Caption, 2) = "" Or Left(lblTinNo.Caption, 1) = "U") Then
                               txt(23) = GProcGetColumnValue("tblMastNarration", "Narration", "G S T", "S", "PolCGSTRt", "N")
                               txt(26) = GProcGetColumnValue("tblMastNarration", "Narration", "G S T", "S", "PolSGSTRt", "N")
                               txt(41) = 0
                            Else
                               txt(23) = 0
                               txt(26) = 0
                               txt(41) = GProcGetColumnValue("tblMastNarration", "Narration", "G S T", "S", "PolIGSTRt", "N")
                            End If
                         End If
                        txtGrid_KeyPress 13
                    End If
                    If frmTypeStr = "SO" Or frmTypeStr = "SD" Then
                        If gClsSearch.SearchMultiRetCol(0) <> "" Then '-- Item Wt ,Rate From Master
                              Select Case frmTypeStr
                                     Case "SY"
                                        Tp = "PY"
                                     Case "SE"
                                        Tp = "PE"
                                     Case "SO"
                                        Tp = "PI"
                                     Case "SD"
                                        Tp = "PI"
                              End Select
                              BalBag = GProcGetColumnValue("tblPurchSub", "PurSubItCode", Val(gClsSearch.SearchMultiRetCol(4)), "N", "sum(PurSubBag)", "N", "vtype = '" & Tp & "' and VFirm = '" & gCCode & "' and Vdt<=('" & Format(dtpDt(2), "YYYY/MM/DD") & "') ")
                              BalBag = BalBag - GProcGetColumnValue("tblSaleSub", "SlSubItCode", Val(gClsSearch.SearchMultiRetCol(4)), "N", "sum(SlSubBag)", "N", "VType = '" & frmTypeStr & "' and VFirm = '" & gCCode & "' and Vdt<=('" & Format(dtpDt(2), "YYYY/MM/DD") & "') and  not (vno = " & txt(0) & " and VYear=" & gCYear & " and VFirm = '" & gCCode & "'  and vtype = '" & frmTypeStr & "')")
                                
                                
                              BalWt = GProcGetColumnValue("tblPurchSub", "PurSubItCode", Val(gClsSearch.SearchMultiRetCol(4)), "N", "sum(PurSubWt)", "N", "vtype = '" & Tp & "' and VFirm = '" & gCCode & "' and Vdt<=('" & Format(dtpDt(2), "YYYY/MM/DD") & "') ")
                              BalWt = BalWt - GProcGetColumnValue("tblSaleSub", "SlSubItCode", Val(gClsSearch.SearchMultiRetCol(4)), "N", "sum(SlSubWt)", "N", "VType = '" & frmTypeStr & "' and VFirm = '" & gCCode & "' and Vdt<=('" & Format(dtpDt(2), "YYYY/MM/DD") & "') and  not (vno = " & txt(0) & " and VYear=" & gCYear & " and VFirm = '" & gCCode & "'  and vtype = '" & frmTypeStr & "')")
                              
                              BalWt = Format(BalWt, WtStr)
                        
                              lblBag.Visible = True
                              lblWeight.Visible = True
                              lblBalBag.Visible = True
                              lblBalWt.Visible = True
                              lblBalBag.Caption = BalBag
                              lblBalWt.Caption = BalWt
                         End If
                    End If
                  End If
            Case 4  '-- Godown
                If Year(gCYSDate) >= 2023 Then
                   msGrid.Col = 5
                Else
                   gClsSearch.SearchMultiField "tblMastNarration", "Narration", Array("Godown"), Array(6000), " NarrType='G' ", Chr(KeyAscii), "Narration", LeftPos, TopPos, True
                   txtGrid = ""
                   KeyAscii = 0
                   If gClsSearch.SearchMultiRetCol(0) <> "" Then
                      txtGrid = gClsSearch.SearchMultiRetCol(0)
                      txtGrid_KeyPress 13
                   End If
                End If
            Case 5  '-- Lot No
                If Year(gCYSDate) >= 2022 Then
                   MillCode = GProcGetColumnValue("TblMastAccount", "AcName", txtTemp(7), "S", "Accode", "N")
                   ItCd = GProcGetColumnValue("tblMastItem", "ItName", msGrid.TextMatrix(msGrid.Row, 3), "S", "ItCode", "N", " ItMillCode=" & MillCode) '--Item Code
                   If ItCd > 0 Then
                      gCn.BeginTrans
                      If Year(gCYSDate) >= 2023 Then
                         gCmd.CommandText = "PrcPrepareLotwiseGodownStock"
                      Else
                         gCmd.CommandText = "PrcPrepareLotwiseStock"
                      End If
                      gCmd.Parameters.Refresh
                      gCmd.Parameters("@VFirm") = gCCode
                      gCmd.Parameters("@Fr_dt") = Format(gCYSDate, "yyyy/mm/dd")
                      gCmd.Parameters("@To_dt") = Format(dtpDt(2), "yyyy/mm/dd")
                      gCmd.Parameters("@MillCode") = MillCode
                      gCmd.Parameters("@ItemCode") = ItCd
                      gCmd.Parameters("@LotNo") = ""
                      If frmTypeStr = "SY" Then
                         gCmd.Parameters("@StkTP") = "T"
                      Else
                         gCmd.Parameters("@StkTP") = "D"
                      End If
                      If FormAction = vbDataActionUpdate Then
                         gCmd.Parameters("@Vno") = txt(0)
                         gCmd.Parameters("@Vtype") = frmTypeStr
                         gCmd.Parameters("@Vyear") = gCYear
                         gCmd.Parameters("@IsModi") = "Y"
                      Else
                         gCmd.Parameters("@Vno") = ""
                         gCmd.Parameters("@Vtype") = ""
                         gCmd.Parameters("@Vyear") = ""
                         gCmd.Parameters("@IsModi") = "N"
                      End If
                      gCmd.Parameters("@UCode") = gUserName
                      gCmd.Execute
                      gCn.CommitTrans
                      If Year(gCYSDate) >= 2023 Then
                         QStr = "Select GodownNm,LotNo,Amt1-Amt3,Amt2-Amt4,iif('" & frmTypeStr & "' = 'SY','Trade','Depot') from tmpLotwiseStock where Amt1-Amt3 <> 0 and username = '" & gUserName & "' "
                      Else
                         QStr = "Select LotNo,Amt1-Amt3,Amt2-Amt4,iif('" & frmTypeStr & "' = 'SY','Trade','Depot') from tmpLotwiseStock where username = '" & gUserName & "' "
                      End If
                      i = GProcRstOpen(tmpRst, QStr, "R")
                      If i > 0 Then
                         txtGrid = ""
                         If Year(gCYSDate) >= 2023 Then
                            gClsSearch.SearchMultiField QStr, "GodownNm,LotNo,Amt1-Amt3,Amt2-Amt4", Array("Godown", "Lot No", "Bal. Bags", "Bal. Weight", "Type"), Array(3000, 2000, 1000, 1500, 1000), "", "", "", 0, 0, False, , True
                         Else
                            gClsSearch.SearchMultiField QStr, "LotNo,Amt1-Amt3,Amt2-Amt4", Array("Lot No", "Bal. Bags", "Bal. Weight", "Type"), Array(2000, 1000, 1500, 1000), "", "", "", 0, 0, False, , True
                         End If
                         If gClsSearch.SearchMultiRetCol(0) <> "" Then
                            If Year(gCYSDate) >= 2023 Then
                               MsgBox ("Balance Bags = " + gClsSearch.SearchMultiRetCol(2))
                               txtGrid = gClsSearch.SearchMultiRetCol(1)
                               msGrid.TextMatrix(msGrid.Row, 4) = gClsSearch.SearchMultiRetCol(0)
                               msGrid.TextMatrix(msGrid.Row, 5) = gClsSearch.SearchMultiRetCol(1)
                            Else
                               txtGrid = gClsSearch.SearchMultiRetCol(0)
                               MsgBox ("Balance Bags = " + gClsSearch.SearchMultiRetCol(1))
                               msGrid.TextMatrix(msGrid.Row, 5) = gClsSearch.SearchMultiRetCol(0)
                            End If
                            msGrid.Col = 6
                            DoEvents
                            DoEvents
                            i = 0
                            gClsSearch.SearchMultiRetCol(0) = ""
                         End If
                      End If
                   End If
                End If
            Case Else
                KeyAscii = GProcValidateKey(rstblSub, msGrid.Col, KeyAscii, txtGrid)
        End Select
   End If
    If KeyAscii = 13 Then
        CheckForNumberTxtGrid
    End If
    '--- Navigate Row ,Col
    If KeyAscii = 13 And .Col < 9 Then 'rstblSub.Fields.Count - 2 Then
        .Col = .Col + 1
        If .Col = 4 And Year(gCYSDate) >= 2023 Then
           .Col = .Col + 1
        End If
    
    ElseIf KeyAscii = 13 And .Col = 9 Then '= rstblSub.Fields.Count - 2 Then
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
    If KeyAscii = 27 Or .Row = 2 Then
        KeyAscii = 0
        If (Val(.TextMatrix(.Row, 8)) = 0) And .rows > 2 Then
            .RemoveItem (.Row)
            txtGrid.text = ""
        End If
        txtGrid.Visible = False
        If frmTypeStr = "SE" Or frmTypeStr = "SY" Then
        Else
            lblBag.Visible = False
            lblWeight.Visible = False
            lblBalBag.Visible = False
            lblBalWt.Visible = False
        End If
        txtGrid.Visible = False
        txt(31).SetFocus
        Exit Sub
    End If
End With
End Sub

'-- Grid Required
Private Function ValidateGrid() As Boolean
Dim i As Long
Dim j As Long
Dim MillCode As Long
Dim RowIndex As Long
Dim ItCode As Long
Dim X As String
Dim Rs1 As Recordset
Dim oldFrmTypeStr As String
Dim LRNo As String
Dim pItCtrlNo As Long
Dim ItCd As Long
Dim wGodCd As Long

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
                    MillCode = GProcGetColumnValue("TblMastAccount", "AcName", txtTemp(7), "S", "Accode", "N")
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
                Case 4 '-- Godown
                    If Trim(.TextMatrix(RowIndex, i)) = "" Then
                        MsgBox "Check Godown.", vbCritical + vbOKOnly, Me.Caption
                        .Col = i
                        msGrid.SetFocus
                        Exit Function
                    End If
                    If GProcGetColumnValue("TblMastNarration", "Narration", .TextMatrix(RowIndex, i), "S", "Narrcode", "N") = 0 Then
                        'MSGrid.SetFocus
                        oldFrmTypeStr = gfrmTypeStr
                        gfrmTypeStr = "G"
                        GProcShowForm frmMastNarration, frmMain.mnuMstfrm(8), Array(frmMastNarration.txt(2)), Array(.TextMatrix(RowIndex, i))
                        gfrmTypeStr = oldFrmTypeStr
                        Exit Function
                    End If
                    If Year(gCYSDate) >= 2023 Then
                       MillCode = GProcGetColumnValue("TblMastAccount", "AcName", txtTemp(7), "S", "Accode", "N")
                       ItCd = GProcGetColumnValue("tblMastItem", "ItName", msGrid.TextMatrix(RowIndex, 3), "S", "ItCode", "N", " ItMillCode=" & MillCode) '--Item Code
                       If ItCd > 0 Then
                          gCn.BeginTrans
                          If Year(gCYSDate) >= 2023 Then
                             wGodCd = GProcGetColumnValue("tblMastNarration", "Narration", msGrid.TextMatrix(RowIndex, 4), "S", "NarrCode", "N", "")
                             gCmd.CommandText = "PrcPrepareLotwiseGodownStock"
                          Else
                             gCmd.CommandText = "PrcPrepareLotwiseStock"
                          End If
                          gCmd.Parameters.Refresh
                          gCmd.Parameters("@VFirm") = gCCode
                          gCmd.Parameters("@Fr_dt") = Format(gCYSDate, "yyyy/mm/dd")
                          gCmd.Parameters("@To_dt") = Format(dtpDt(2), "yyyy/mm/dd")
                          gCmd.Parameters("@MillCode") = MillCode
                          gCmd.Parameters("@ItemCode") = ItCd
                          gCmd.Parameters("@LotNo") = .TextMatrix(RowIndex, 5) & vbNullString
                          If frmTypeStr = "SY" Then
                             gCmd.Parameters("@StkTP") = "T"
                          Else
                             gCmd.Parameters("@StkTP") = "D"
                          End If
                          gCmd.Parameters("@Vno") = txt(0)
                          gCmd.Parameters("@Vtype") = frmTypeStr
                          gCmd.Parameters("@Vyear") = gCYear
                          gCmd.Parameters("@IsModi") = "Y"
                          gCmd.Parameters("@GodownCd") = wGodCd
                          gCmd.Parameters("@UCode") = gUserName
                          gCmd.Execute
                          gCn.CommitTrans
                          X = "Select GodownNm,LotNo,Amt1-Amt3,Amt2-Amt4,iif('" & frmTypeStr & "' = 'SY','Trade','Depot') from tmpLotwiseStock where username = '" & gUserName & "' "
                          j = GProcRstOpen(Rs1, X, "R")
                          If j > 0 Then
                             If Val(.TextMatrix(RowIndex, 6)) > Rs1.Fields(2) Then
                                MsgBox "Sales Bags are more than Balance Bags.", vbCritical + vbOKOnly, Me.Caption
                                .Col = 6
                                msGrid.SetFocus
                                Exit Function
                             End If
                             If .TextMatrix(RowIndex, 5) <> Rs1.Fields(1) Then
                                MsgBox "Please Check Lot No....", vbCritical + vbOKOnly, Me.Caption
                                .Col = 5
                                msGrid.SetFocus
                                Exit Function
                             End If
                             If .TextMatrix(RowIndex, 4) <> Rs1.Fields(0) Then
                                MsgBox "Please Check Godown....", vbCritical + vbOKOnly, Me.Caption
                                .Col = 4
                                msGrid.SetFocus
                                Exit Function
                             End If
                          
                          Else
                             MsgBox "Please Check Stock....", vbCritical + vbOKOnly, Me.Caption
                             .Col = 3
                             msGrid.SetFocus
                             Exit Function
                          End If
                       End If
                    End If
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
                Case 8 '-- Wt
                    If CDbl(.TextMatrix(RowIndex, i)) = 0 Then
                        MsgBox "Check Weight.", vbCritical + vbOKOnly, Me.Caption
                        .Col = i
                        msGrid.SetFocus
                        Exit Function
                    End If
                    If Abs((Val(.TextMatrix(RowIndex, 6)) * Val(.TextMatrix(RowIndex, 7))) - Val(.TextMatrix(RowIndex, 8))) > Val(.TextMatrix(RowIndex, 7)) Then
                        MsgBox "Please Check Weight ", vbCritical + vbOKOnly, Me.Caption
                    End If
'                    '-- If Wt < Issue wt
                
'                    If CDbl(.TextMatrix(RowIndex, 8)) < CDbl(.TextMatrix(RowIndex, 14)) Then
'                        MsgBox "Check Wt.Weight Less than issued.", vbCritical + vbOKOnly, Me.Caption
'                        .Col = i
'                        msGrid.SetFocus
'                        Exit Function
'                    End If
                Case 12 '-- amount
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
            MillCode = GProcGetColumnValue("TblMastAccount", "AcName", txtTemp(7), "S", "Accode", "N")
            ItCode = GProcGetColumnValue("tblMastItem", "ItName", .TextMatrix(RowIndex, 3), "S", "ItCode", "N", " ItMillCode=" & MillCode)
            X = "Select * from tblBooking where Vno=" & CLng(.TextMatrix(RowIndex, 1)) & " and  " _
            & " BkItCode=" & ItCode & " and  BkSrNo=" & .TextMatrix(RowIndex, 17) & " and Vtype='" & .TextMatrix(RowIndex, 15) & "'  and VYear = " & .TextMatrix(RowIndex, 16) & " and Vfirm = '" & gCCode & "'    and         "
            
            X = X + "  vdt='" & Format(CDate(.TextMatrix(RowIndex, 2)), "YYYY/MM/DD") & "'"
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
'If FormAction = vbDataActionAddNew And Val(txt(23)) = 0 Then
'    txt(23) = GProcGetColumnValue("tblMastNarration", "Narration", Cbo(22), "S", "MastTaxRate", "N")
'    txt(23) = Format(txt(23), FStr)
'End If
End Sub
Private Sub SaveBtnEd()
'--- Vno,vdt,Party,A/c,Mill,Amount
If (CDbl(txt(0)) = 0 And FormAction <> vbDataActionAddNew) Or GProcIsDateValid(dtpDt(2)) = False Or txtTemp(5) = "" Or txtTemp(6) = "" Or txtTemp(7) = "" Or CDbl(txt(16)) = 0 Or CDbl(txt(15)) = 0 Then
    cmdBtn(13).Enabled = False
    cmdBtn(13).BackColor = &H80000004
Else
    cmdBtn(13).Enabled = True
    cmdBtn(13).BackColor = &HFFFF80
End If
'-- Temp
If gCYear < 2018 Then
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
Dim MillCode As Long
Dim HsnNm As String
'If gCYear < 2018 Then
'    ValidateData = True
'    Exit Function
'End If
If gCYear >= 2021 Then
   MillCode = GProcGetColumnValue("TblMastAccount", "AcName", txtTemp(7), "S", "Accode", "N")
   HsnNm = GProcGetColumnValue("tblMastItem", "ItName", msGrid.TextMatrix(1, 3), "S", "ItHsn2", "S", " ItMillCode=" & MillCode)
   If Len(HsnNm) < 6 Then
      MsgBox "HSN Code Must be Minimum 6 Digit", vbInformation + vbOKOnly, Me.Caption
      txtTemp(5).SetFocus
      Exit Function
   End If
End If


If Len(Trim(lblTinNo.Caption)) = 0 Then
      MsgBox "Party GSTIN Not Found", vbInformation + vbOKOnly, Me.Caption
      txtTemp(5).SetFocus
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
    MillCtrlEd
    '---Validate Vou.Date = Bill Date
'    If mskdt(10) <> dtpDt(2) Then
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
    '-- Tax %
    
    '-- Tax Amount
    If (CDbl(txt(24)) > 0 Or CDbl(txt(23)) > 0) And Abs(Round(Val(txt(21)) * Val(txt(23)) / 100, 2) - Val(txt(24))) > 1 Then
        MsgBox "Tax Amount on Ass.Val not correct", vbInformation + vbOKOnly, Me.Caption
        txt(24).SetFocus
        Exit Function
    End If
    
    
    
    '---- Party Account
    If txtTemp(5) = "" Then
        MsgBox "Check Party.", vbInformation + vbOKOnly, Me.Caption
        txtTemp(3).SetFocus
        Exit Function
    End If
        '--Create Party
        If GProcGetColumnValue("tblMastAccount", "AcName", txtTemp(5), "S", "AcCode", "N") = 0 Then
           GName = GProcGetColumnValue("TblMastGroup", "AgCode", 90017, "N", "AgName", "S")
           GProcShowForm frmMastAccount, frmMain.mnuMstfrm(0), Array(frmMastAccount.txt(1), frmMastAccount.txtTemp(2)), Array(txtTemp(5), GName)
           Exit Function
        End If
    '---- Sales Account
    If txtTemp(6) = "" Then
        MsgBox "Check Purchase Account.", vbInformation + vbOKOnly, Me.Caption
        txtTemp(6).SetFocus
        Exit Function
    End If
        '--Create Party
        If GProcGetColumnValue("tblMastAccount", "AcName", txtTemp(6), "S", "AcCode", "N") = 0 Then
           GName = GProcGetColumnValue("TblMastGroup", "AgCode", 90023, "N", "AgName", "S")
           GProcShowForm frmMastAccount, frmMain.mnuMstfrm(0), Array(frmMastAccount.txt(1), frmMastAccount.txtTemp(2)), Array(txtTemp(6), GName)
           Exit Function
        End If
    '---- Mill
    If txtTemp(7) = "" Then
        MsgBox "Check Mill.", vbInformation + vbOKOnly, Me.Caption
        txtTemp(7).SetFocus
        Exit Function
    End If
        '--Create Mill
        If GProcGetColumnValue("tblMastAccount", "AcName", txtTemp(7), "S", "AcCode", "N") = 0 Then
           GName = GProcGetColumnValue("TblMastGroup", "AgCode", 90029, "N", "AgName", "S")
           GProcShowForm frmMastAccount, frmMain.mnuMstfrm(0), Array(frmMastAccount.txt(1), frmMastAccount.txtTemp(2)), Array(txtTemp(7), GName)
           Exit Function
        End If
    '--- Broker
        '--Create Broker
        If txtTemp(8) = "" Then txtTemp(8) = "-"
        If GProcGetColumnValue("tblMastAccount", "AcName", txtTemp(8), "S", "AcCode", "N") = 0 Then
           GName = GProcGetColumnValue("TblMastGroup", "AgCode", 90016, "N", "AgName", "S")
           GProcShowForm frmMastAccount, frmMain.mnuMstfrm(0), Array(frmMastAccount.txt(1), frmMastAccount.txtTemp(2)), Array(txtTemp(8), GName)
           Exit Function
        End If
    '---- Amount
'--Consign GSTIN
        If txtTemp(22) <> "-" Then
           GName = GProcGetColumnValue("tblMastNarration", "NarrCode", txt(22), "N", "ConsiGSTIN", "S")
           If IsNull(GName) Then
              MsgBox "Check Mill.", vbInformation + vbOKOnly, Me.Caption
              txtTemp(22).SetFocus
              Exit Function
           End If
        End If
    If CDbl(txt(16)) <= 0 Then
        MsgBox "Check Voucher Amount.", vbInformation + vbOKOnly, Me.Caption
        Exit Function
    End If
    If Trim(txt(38)) <> "" Then
        X = "Select SlBillNo From " & tblName & " Where GpNo= '" & txt(38) & "' and   vno <> " & CLng(txt(0)) & "  and Vtype='" & frmTypeStr & "'  and Vfirm = '" & gCCode & "'  and VYear=" & gCYear
        i = GProcRstOpen(rstAddCheck, X, "R")
        If i > 0 Then
            MsgBox "Same G.P. No is Available in Bill No. " & rstAddCheck.Fields(0), vbCritical + vbOKOnly, "Duplicate Record"
'            txt(9).SetFocus
 '           Exit Function
        End If
    End If
    '----Duplicate Records
    '----Addition Case
    If FormAction = vbDataActionAddNew Then
        X = "Select vno From " & tblName & " Where vno=" & CLng(txt(0)) & " and Vtype='" & frmTypeStr & "'  and Vfirm = '" & gCCode & "'  and VYear=" & gCYear
        i = GProcRstOpen(rstAddCheck, X, "R")
        If i > 0 Then
            MsgBox "Record is Available of Voucher No. " & CLng(txt(0)), vbCritical + vbOKOnly, "Duplicate Record"
            txt(9).SetFocus
            Exit Function
        End If
        rstAddCheck.Close
    End If
    ValidateData = True
End Function
Private Sub FillTxtFromTemp()
lblTcsLimit.Caption = ""
txt(1) = frmTypeStr '--VType
If IsDate(dtpDt(2)) Then '--Vdt
    txt(2) = CDate(dtpDt(2))
Else
    txt(2) = ""
End If
If IsDate(mskDt(10)) Then '--Bill Date
    txt(10) = CDate(mskDt(10))
Else
    txt(10) = ""
End If
If IsDate(mskDt(39)) Then '--G.P. Date
    txt(39) = CDate(mskDt(39))
Else
    txt(39) = ""
End If

If IsDate(dtpDt(32)) Then '--Due Date
    txt(32) = CDate(dtpDt(32))
Else
    txt(32) = ""
End If
If IsDate(mskDt(14)) Then '--lr Date
    txt(14) = CDate(mskDt(14))
Else
    txt(14) = ""
End If

'-- Dr Code
txt(5) = GProcGetColumnValue("tblMastAccount", "AcName", txtTemp(5), "S", "AcCode", "N")
'-- Cr Code
txt(6) = GProcGetColumnValue("tblMastAccount", "AcName", txtTemp(6), "S", "AcCode", "N")
'--Mill
txt(7) = GProcGetColumnValue("tblMastAccount", "AcName", txtTemp(7), "S", "AcCode", "N")
'-- broker
If txtTemp(8) = "" Then
   txtTemp(8) = "-"
End If
txt(8) = GProcGetColumnValue("tblMastAccount", "AcName", txtTemp(8), "S", "AcCode", "N")
'-- Transport
If txtTemp(11) = "" Then
   txtTemp(11) = "-"
End If
txt(11) = GProcGetColumnValue("tblMastNarration", "Narration", txtTemp(11), "S", "NarrCode", "N")
'-- Consignee
If txtTemp(22) = "" Then
   txtTemp(22) = "-"
End If

'--TdsAc
If txtTemp(70) = "" Then
   txtTemp(70) = "-"
End If
txt(70) = GProcGetColumnValue("tblMastAccount", "AcName", txtTemp(70), "S", "AcCode", "N")



txt(22) = GProcGetColumnValue("tblMastNarration", "Narration", txtTemp(22), "S", "NarrCode", "N")
txt(44) = Cbo(44).ListIndex
txt(3) = gCYear
txt(4) = gCCode
ShowAccountBalance
End Sub
Private Sub FillTempFromTxt()
Dim X As String
Dim MsOms As Integer
FillCombo
lblTcsLimit.Caption = ""
If IsDate(txt(2)) Then '--VDt
    dtpDt(2) = txt(2)
End If
If IsDate(txt(10)) Then '--Bill Date
    mskDt(10) = txt(10)
Else
    mskDt(10) = "__/__/____"
End If

If IsDate(txt(39)) Then '--G.P. Date
    mskDt(39) = txt(39)
Else
    mskDt(39) = "__/__/____"
End If
If IsDate(txt(32)) Then '--Due Date
    dtpDt(32) = txt(32)
Else
End If
If IsDate(txt(14)) Then '--LR Date
    mskDt(14) = txt(14)
Else
    mskDt(14) = "__/__/____"
End If
'-- Dr Code
txtTemp(5) = GProcGetColumnValue("tblMastAccount", "AcCode", txt(5), "N", "AcName", "S")
txtCode(5) = GProcGetColumnValue("tblMastAccount", "AcCode", txt(5), "N", "AcOurCode", "S")


lblTinNo.Caption = GProcGetColumnValue("TblMastAccount", "ACName", txtTemp(5), "S", "AcGSTIN", "S")
MsOms = GProcGetColumnValue("tblMastAccount", "AcCode", txt(5), "N", "AcMsOmsParty", "N")
If MsOms = 0 Then
   lblMsOms.Caption = "M.S."
Else
   lblMsOms.Caption = "O.M.S."
End If
'lblAdd1.Caption = GProcGetColumnValue("TblMastAccount", "AcCode", txt(5), "N", "AcAdd1", "S")
'lblAdd2.Caption = GProcGetColumnValue("TblMastAccount", "AcCode", txt(5), "N", "AcAdd2", "S")
lblMblNoParty.Caption = GProcGetColumnValue("TblMastAccount", "AcCode", txt(5), "N", "AcContNo", "S")

'-- Cr Code
txtTemp(6) = GProcGetColumnValue("tblMastAccount", "AcCode", txt(6), "N", "AcName", "S")

'--Mill
txtTemp(7) = GProcGetColumnValue("tblMastAccount", "AcCode", txt(7), "N", "AcName", "S")
txtCode(7) = GProcGetColumnValue("tblMastAccount", "AcCode", txt(7), "N", "AcOurCode", "S")

'--Broker
txtTemp(8) = GProcGetColumnValue("tblMastAccount", "AcCode", txt(8), "N", "AcName", "S")
txtCode(8) = GProcGetColumnValue("tblMastAccount", "AcCode", txt(8), "N", "AcOurCode", "S")

'--Transport
txtTemp(11) = GProcGetColumnValue("tblMastNarration", "NarrCode", txt(11), "N", "Narration", "S")

'--Consignee
txtTemp(22) = GProcGetColumnValue("tblMastNarration", "NarrCode", txt(22), "N", "Narration", "S")
'--TdsAc
txtTemp(70) = GProcGetColumnValue("tblMastAccount", "AcCode", txt(70), "N", "AcName", "S")
txtCode(70) = GProcGetColumnValue("tblMastAccount", "AcCode", txt(70), "N", "AcOurCode", "S")



'-- TaxType
Cbo(44).ListIndex = Val(txt(44))
'-- Fill Related Record
FillRelatedRecord
ShowAccountBalance
MillCtrlEd
SetControlEd
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
                If frmTypeStr = "SE" Or frmTypeStr = "SY" Then
'                   Label6.Visible = True
 '                  Label6.Caption = "Purchase"
                   lblBag.Visible = True
                   lblWeight.Visible = True
'                   lblBag.Caption = "V.No    :"
 '                  lblWeight = "Date    :"
                   lblBalBag.Visible = True
                   lblBalWt.Visible = True
                   lblBalBag.Caption = GProcGetColumnValue("tblPurchSub,tblPurch", "PurSubItCtrlNo", "'" & msGrid.TextMatrix(RowIndex, 14) & "'", "N", "PurBillNo", "S", "tblPurch.Vno = tblpurchSub.Vno and tblPurch.Vtype = tblpurchSub.Vtype and tblPurch.Vyear = tblpurchSub.Vyear and tblPurch.Vfirm = tblpurchSub.Vfirm")
                   lblBalWt.Caption = GProcGetColumnValue("tblPurchSub", "PurSubItCtrlNo", "'" & msGrid.TextMatrix(RowIndex, 14) & "'", "N", "Vno", "N")
                End If
            RowIndex = RowIndex + 1
            .MoveNext
            If Not .EOF Then msGrid.rows = msGrid.rows + 1
        Loop
    End If
End With
txtSlSub(17) = msGrid.TextMatrix(1, 1)
If IsDate(msGrid.TextMatrix(1, 2)) Then
   mskDtSlSub(19) = msGrid.TextMatrix(1, 2)
Else
   mskDtSlSub(19) = "__/__/____"
End If
'--- Fill Related Control
SetRelatedRS
If Val(txt(17)) > 0 Then
   lblCharity.Visible = True
   txt(17).Visible = True
Else
   lblCharity.Visible = False
   txt(17).Visible = False
End If
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
Dim Amt2 As Double
Dim ROff As Double
Dim ExmptAmt As Double
Dim ExmptKg As Double
Dim AssValue As Double
Dim RowIndex As Integer
Dim TotBag As Long
Dim TotWt As Double
Dim LCost As Double
Dim MillAcCode As Long
Dim IsExemptMill As Long
Dim wMillCode As Long
Dim wCharityRt As Double
Dim wCharityOn As Integer
Dim wCharityAmt As Double
Dim TotCharityAmt As Double
Dim TcsAmt As Double
Dim TaxCd As Long
Dim X As String
Dim wPtyCd As Long
Dim TcsLimitAmt As Double
Dim tmpRst As Recordset
Dim tmpRst2 As Recordset
Dim PtyOpBal As Double
Dim IsTdsParty As Integer
Dim TDSLimitAmt As Double
Dim TdsAmt As Double
Dim wDepoHoCd As Long
Dim TdsOnAmt As Double
Dim TcsOnAmt As Double
Dim PtyPan As String
Dim TdsFrom1stBill As Integer
'-- Is Exempt Mill
'IsExemptMill = GProcGetColumnValue("TblMastAccount", "AcName", txttemp(7), "S", "ACIsExemptMill", "N")
wMillCode = GProcGetColumnValue("tblMastAccount", "AcName", txtTemp(7), "S", "AcCode", "N")
wPtyCd = GProcGetColumnValue("TblMastAccount", "AcName", txtTemp(5), "S", "AcCode", "N")
PtyPan = GProcGetColumnValue("tblMastAccount", "AcName", txtTemp(5), "S", "AcPAN", "S")


'-- Due Date
If IsNumeric(txt(31)) = False Then txt(31) = "0"
txt(31) = Format(txt(31), PStr)
If IsDate(dtpDt(2)) Then dtpDt(32) = DateAdd("d", txt(31), dtpDt(2))
'---- Sub amt
CheckForNumberGrid
CheckForNumber
TotCharityAmt = 0
ExmptKg = 0
ExmptAmt = 0
RowIndex = 1
With msGrid
    Do While RowIndex < .rows
        '-- sub amt
        If Val(.TextMatrix(RowIndex, 12)) > 0 Then
            SubAmt = SubAmt + Val(.TextMatrix(RowIndex, 12))
            If frmTypeStr = "SE" And Val(.TextMatrix(RowIndex, 14)) > 0 Then
               X = "select tblpurch.vno,tblpurch.vtype,tblpurch.vyear,tblpurch.vfirm,tblPurch.PurExemptPerKg from tblpurch,tblpurchsub where tblpurch.vno = tblpurchsub.vno and tblpurch.vtype = tblpurchsub.vtype and " _
                   & " tblpurch.vyear = tblpurchsub.vyear and tblpurch.vfirm = tblpurchsub.vfirm and tblpurch.vfirm = '" & gCCode & "' and tblpurchsub.PurSubItCtrlNo = '" & .TextMatrix(RowIndex, 14) & "'"
                If GProcRstOpen(tmpRst, X, "R") > 0 Then
                   ExmptKg = tmpRst.Fields("PurExemptPerKg")
                   ExmptAmt = ExmptAmt + (ExmptKg * Val(.TextMatrix(RowIndex, 8)))
                End If
            End If
        End If
        '--Tot Bag
        If Val(.TextMatrix(RowIndex, 6)) > 0 Then
            TotBag = TotBag + Val(.TextMatrix(RowIndex, 6))
        End If
        '--Tot Wt
        If Val(.TextMatrix(RowIndex, 8)) > 0 Then
            TotWt = TotWt + Val(.TextMatrix(RowIndex, 8))
        End If
        
        If RowIndex = 1 And msGrid.TextMatrix(1, 3) <> "" Then
           ItemType = GProcGetColumnValue("tblMastItem", "ItName", msGrid.TextMatrix(1, 3), "S", "ItType", "N", "ItMillcode=" & wMillCode) '--Item Code
           If ItemType = 0 Then
              If Left(lblTinNo.Caption, 2) = "27" Or Left(lblTinNo.Caption, 1) = "U" Or Trim(lblTinNo.Caption) = "" Then
                 txt(23) = GProcGetColumnValue("tblMastNarration", "Narration", "G S T", "S", "CotCGSTRt", "N")
                 txt(26) = GProcGetColumnValue("tblMastNarration", "Narration", "G S T", "S", "CotSGSTRt", "N")
                 txt(41) = 0
              Else
                 txt(23) = 0
                 txt(26) = 0
                 txt(41) = GProcGetColumnValue("tblMastNarration", "Narration", "G S T", "S", "CotIGSTRt", "N")
              End If
           Else
              If Left(lblTinNo.Caption, 2) = "27" Or Left(lblTinNo.Caption, 1) = "U" Or Trim(lblTinNo.Caption) = "" Then
                 txt(23) = GProcGetColumnValue("tblMastNarration", "Narration", "G S T", "S", "PolCGSTRt", "N")
                 txt(26) = GProcGetColumnValue("tblMastNarration", "Narration", "G S T", "S", "PolSGSTRt", "N")
                 txt(41) = 0
              Else
                 txt(23) = 0
                 txt(26) = 0
                 txt(41) = GProcGetColumnValue("tblMastNarration", "Narration", "G S T", "S", "PolIGSTRt", "N")
              End If
           End If
        End If
        wCharityRt = GProcGetColumnValue("TblMastItem", "ITName", msGrid.TextMatrix(RowIndex, 3), "S", "CharityRt", "N", " ItMillCode=" & wMillCode)
        wCharityOn = GProcGetColumnValue("TblMastItem", "ITName", msGrid.TextMatrix(RowIndex, 3), "S", "CharityOn", "N", " ItMillCode=" & wMillCode)
        If wCharityRt > 0 Then
           If wCharityOn = 0 Then   '-- On Kg
              wCharityAmt = Format(Val(msGrid.TextMatrix(RowIndex, 8)) * wCharityRt, FStr)
           Else                     '-- On Pkg
              wCharityAmt = Format(Val(msGrid.TextMatrix(RowIndex, 6)) * wCharityRt, FStr)
           End If
        Else
           wCharityAmt = 0
        End If
        TotCharityAmt = TotCharityAmt + wCharityAmt
        RowIndex = RowIndex + 1
    Loop
End With
TotCharityAmt = Format(TotCharityAmt, FStr)
txt(17) = TotCharityAmt
txt(19) = Format(ExmptAmt, FStr)
txt(20) = ExmptKg

'--- Subamt+Add -Less

AssValue = SubAmt + CDbl(txt(17)) - CDbl(txt(18)) - CDbl(txt(19))
If AssValue < 0 Then
   AssValue = 0
End If
'- 1 SGST
TaxAmt = AssValue * CDbl(txt(23)) / 100
TaxAmt = Format(TaxAmt, FStr)

'- 2 CGST
TaxAmtCGST = AssValue * CDbl(txt(26)) / 100
TaxAmtCGST = Format(TaxAmtCGST, FStr)
   
'- 3 IGST
TaxAmtIGST = AssValue * CDbl(txt(41)) / 100
TaxAmtIGST = Format(TaxAmtIGST, FStr)


Amt = Format(SubAmt + TaxAmt + TaxAmtCGST + TaxAmtIGST + CDbl(txt(17)) - CDbl(txt(18)) + CDbl(txt(28)) - CDbl(txt(29)), FStr)
TcsLimitAmt = 0
'---TCS
If Format(dtpDt(2).Value, "YYYY/MM/dd") >= Format("2021/07/01", "YYYY/MM/dd") Then
    gCn.BeginTrans
    IsTdsParty = GProcGetColumnValue("TblMastAccount", "AcName", txtTemp(5), "S", "AcIsLessTDSOnRec", "N")
    TdsFrom1stBill = GProcGetColumnValue("TblMastAccount", "AcName", txtTemp(5), "S", "AcIsTDSfrom1stBill", "N")
    gCn.CommitTrans
Else
   IsTdsParty = 0
   TdsFrom1stBill = 0
End If
''--As per Munole's Phone On 02/04/2022 Tds Limit is applicable for Depot & Mill Bill
'If Format(dtpDt(2).Value, "YYYY/MM/dd") >= Format("2020/10/01", "YYYY/MM/dd") And (gCIsTcsFirm = 1 Or gCIsDepotFirm = 1) And IsTdsParty = 0 Then
'   X = "Select sum(SlBillAmt) from tblSale,tblMastAccount where tblSale.SlAcDrCode = tblMastAccount.AcCode and (SlAcDrCode = " & wPtyCd & " or tblMastAccount.AcPan = '" & PtyPan & "') and vfirm = '" & gCCode & "' and vtype in ('SY','SD') and vyear = " & gCYear & " and vdt <= '" & Format(dtpDt(2).Value, "YYYY/MM/dd") & "'  "
'   X = X + " and  ltrim(str(vno))+vtype+vFirm+ltrim(str(vyear)) <> '" & Trim(Str(txt(0))) & "" & gfrmTypeStr & "" & gCCode & "" & LTrim(Str(gCYear)) & "'"
'   i = GProcRstOpen(tmpRst, X, "R")
'   TcsLimitAmt = IIf(IsNull(tmpRst.Fields(0)), 0, tmpRst.Fields(0))
'   TcsLimitAmt = TcsLimitAmt + Amt
'
'   If GProcRstOpen(tmpRst2, "select Vamt from tblVoucher where vtype = '0P' and  VDrAcCode= " & wPtyCd & "  And Vyear = " & gCYear & " and vfirm ='" & gCCode & "'", "R") > 0 Then
'      TcsLimitAmt = TcsLimitAmt + IIf(IsNull(tmpRst2.Fields(0)), 0, tmpRst2.Fields(0))
'   End If
'   If TcsLimitAmt >= 5000000 Or (gCIsDepotFirm = 1 And gCYear <= 2021) Then
'      If gCYear >= 2021 Then
'         txt(64) = 0.1
'      Else
'         txt(64) = 0.075
'      End If
'      If gCYear >= 2023 Then
'         If TcsLimitAmt - Amt >= 5000000 Then
'            TcsOnAmt = Amt
'         Else
'             TcsOnAmt = TcsLimitAmt - 5000000
'         End If
'         If TcsOnAmt > Amt Then
'            TcsOnAmt = Amt
'         End If
'         TcsAmt = (TcsOnAmt * Val(txt(64))) / 100
'         TcsAmt = Format(Round(TcsAmt, 2), FStr)
'         txt(63) = Format(TcsOnAmt, FStr)
'         txt(65) = TcsAmt
'      Else
'         txt(63) = Format(Amt, FStr)
'         TcsAmt = (Amt * Val(txt(64))) / 100
'         If Format(dtpDt(2).Value, "YYYY/MM/dd") >= Format("2021/01/18", "YYYY/MM/dd") Then
'            TcsAmt = Format(Round(TcsAmt, 2), FStr)
'         Else
'            TcsAmt = Format(GProcMakeRounding(TcsAmt), FStr)
'         End If
'         txt(65) = TcsAmt
'      End If
'   Else
'      txt(63) = Amt
'      txt(64) = 0
'      txt(65) = 0
'   End If
'   lblTcsLimit.Caption = "Op.+Sales up to Date = " + Str(TcsLimitAmt)
'
'   tmpRst.Close
'   tmpRst2.Close
'Else
    TcsOnAmt = 0
    TcsAmt = 0
    txt(63) = Format(0, FStr)
    txt(64) = Format(0, FStr)
    txt(65) = Format(0, FStr)
'End If

Amt2 = GProcMakeRounding(Amt + TcsAmt)
ROff = Format(Amt2 - Format(Amt + TcsAmt, FStr), FStr)

'--For TDS
If Format(dtpDt(2).Value, "YYYY/MM/dd") >= Format("2021/07/01", "YYYY/MM/dd") And (gCIsTcsFirm = 1 Or gCIsDepotFirm = 1) And (IsTdsParty = 1 Or TdsFrom1stBill = 1) Then
   
   '--As Kanayalalji's phone on 05/04/2023 tds on party or salme pan no of party
   
   X = "Select sum(SlBillAmt) from tblSale,tblMastAccount where tblSale.SlAcDrCode = tblMastAccount.AcCode and (SlAcDrCode = " & wPtyCd & " or tblMastAccount.AcPan = '" & PtyPan & "') and vfirm = '" & gCCode & "' and vtype in ('SY','SD') and vyear = " & gCYear & " and vdt <= '" & Format(dtpDt(2).Value, "YYYY/MM/dd") & "'  "
   X = X + " and  ltrim(str(vno))+vtype+vFirm+ltrim(str(vyear)) <> '" & Trim(Str(txt(0))) & "" & gfrmTypeStr & "" & gCCode & "" & LTrim(Str(gCYear)) & "'"
   i = GProcRstOpen(tmpRst, X, "R")
   TDSLimitAmt = IIf(IsNull(tmpRst.Fields(0)), 0, tmpRst.Fields(0))
   TDSLimitAmt = TDSLimitAmt + Amt
   
   If TDSLimitAmt >= 5000000 Or (gCIsDepotFirm = 1 And gCYear <= 2021) Or TdsFrom1stBill = 1 Then
      txt(68) = 0.1
      If gCYear >= 2023 Then
         If TDSLimitAmt - Amt >= 5000000 Then
            TdsOnAmt = AssValue
         Else
             TdsOnAmt = TDSLimitAmt - 5000000
         End If
         If TdsOnAmt > AssValue Or TdsFrom1stBill = 1 Then
            TdsOnAmt = AssValue
         End If
         TdsAmt = (TdsOnAmt * Val(txt(68))) / 100
      Else
         TdsOnAmt = 0
         TdsAmt = (AssValue * Val(txt(68))) / 100
      End If
      TdsAmt = Format(GProcMakeRounding(TdsAmt), FStr)
      txt(69) = TdsAmt
      If TdsAmt > 0 Then
         If gCIsDepotFirm = 1 Then
            wDepoHoCd = GProcGetColumnValue("tblMastAccount", "AcName", txtTemp(7), "S", "AcDepotHoAc", "N")
            txtTemp(70) = GProcGetColumnValue("tblMastAccount", "AcCode", "" & wDepoHoCd & "", "N", "AcName", "S")
            txtCode(70) = GProcGetColumnValue("tblMastAccount", "AcCode", "" & wDepoHoCd & "", "N", "AcOurCode", "S")
         Else
            txtTemp(70) = GProcGetColumnValue("tblMastAccount", "AcCode", "" & gTdsOnSalesCode & "", "N", "AcName", "S")
            txtCode(70) = GProcGetColumnValue("tblMastAccount", "AcCode", "" & gTdsOnSalesCode & "", "N", "AcOurCode", "S")
        End If
      Else
         txtTemp(70) = "-"
         txtCode(70) = ""
      End If
   Else
      txt(68) = 0
      txt(69) = 0
      TdsOnAmt = 0
      txtTemp(70) = "-"
      txtCode(70) = ""
   End If
   lblTcsLimit.Caption = "Sales up to Date = " + Str(TDSLimitAmt)
   tmpRst.Close
Else
    TdsAmt = 0
    TdsOnAmt = 0
    txt(68) = Format(0, FStr)
    txt(69) = Format(0, FStr)
    txtTemp(70) = "-"
    txtCode(70) = ""
End If
If gCYear >= 2023 Then
   txt(74) = Format(TdsOnAmt, FStr)
End If
txt(15) = SubAmt
txt(24) = TaxAmt
txt(27) = TaxAmtCGST
txt(42) = TaxAmtIGST
txt(21) = AssValue
txt(30) = ROff
CheckForNumber
If gCYear >= 2018 Then
   txt(16) = Round(Amt2, 0)
End If

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
'--Is Related Receipt

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
Private Sub GatePassData()
Dim RsG As Recordset
Dim X As String
Dim TmpGpVno As Long
Dim GPSLType As Long
Dim wMillCode As String
Set RsG = New Recordset
Dim i As Integer
If frmTypeStr = "SY" Then '--- Trade
   GPSLType = 0
ElseIf frmTypeStr = "SO" Then '--Consignment
   GPSLType = 1
ElseIf frmTypeStr = "SD" Then '--Depot Sale
'   MillCondStr = " and GPmill='" & txttemp(3) & "'"
   GPSLType = 2
ElseIf frmTypeStr = "ST" Then '--SIT
   GPSLType = 3
Else
   '-- Trade
   If gfrmHunkStr = "HNK" Then
      GPSLType = 4
   Else
      GPSLType = 0
   End If
'   MillCondStr = ""
End If
'    CheckForNumberTmp


X = "Select GpNo,format(Vdt,'dd/MM/yyyy'),tblMastAccount.AcName,tblMastAccount_Br.AcName,tblMastAccount_Mill.AcName,Vno from tblGatePass,tblMastAccount,tblMastAccount tblMastAccount_Br,tblMastAccount tblMastAccount_Mill where tblGatePass.GpAcDrCode = tblMastAccount.AcCode and tblGatePass.GpBroker = tblMastAccount_Br.AcCode and tblGatePass.GpMillCode = tblMastAccount_Mill.AcCode and VFirm  = '" & gCCode & "' and VYear= " & gCYear & " and GpType=" & GPSLType & " and VnoBill = 0 "
gClsSearch.SearchMultiField X, "", Array("Gp No", "Gp Date", "Party", "Broker", "Mill", "Vno"), Array(1200, 1200, 3800, 2900, 3800, 0), "", "", "Vdt,GpNo", 0, 0, False, , True
If gClsSearch.SearchMultiRetCol(0) <> "" Then
   txt(38) = gClsSearch.SearchMultiRetCol(0)
   dtpDt(39) = Format(gClsSearch.SearchMultiRetCol(1), "dd/MM/yyyy")
   txtTemp(5) = gClsSearch.SearchMultiRetCol(2) '--Party
   txtTemp(8) = gClsSearch.SearchMultiRetCol(3) '--Broker
   txtTemp(7) = gClsSearch.SearchMultiRetCol(4) '--Mill
   TmpGpVno = gClsSearch.SearchMultiRetCol(5) '--Gp Vno
   If FormAction = vbDataActionAddNew Then dtpDt(2) = CDate(dtpDt(39))   '-- Inv Dt
   wMillCode = GProcGetColumnValue("tblMastAccount", "AcName", txtTemp(7), "S", "AcCode", "N")
   
   X = "Select * from tblGatePassSub where VNo=" & TmpGpVno & " and vtype = 'GP' and VFirm  = '" & gCCode & "' and VYear= " & gCYear & " and GpSubGpType=" & GPSLType & "  "
   If GProcRstOpen(RsG, X, "R") > 0 Then
      i = 1
      With msGrid
           Do While Not RsG.EOF
              i = msGrid.rows - 1
'             msGrid.TextMatrix(msGrid.Row, 1) = CLng(.SearchMultiRetCol(1)) '--Book Vno
'             msGrid.TextMatrix(msGrid.Row, 2) = Format(.SearchMultiRetCol(2), "dd/mm/yyyy") '-- Book Vdt
              msGrid.TextMatrix(i, 3) = GProcGetColumnValue("tblMastItem", "ItCode", RsG.Fields("GpSubItCode"), "N", "ItName", "S", " ItMillCode=" & wMillCode) '--Item
              msGrid.TextMatrix(i, 4) = GProcGetColumnValue("tblMastNarration", "NarrCode", RsG.Fields("GpSubGodown"), "N", "Narration", "S", "")
              msGrid.TextMatrix(i, 5) = RsG.Fields("GpSubLotNo") & vbNullString
              msGrid.TextMatrix(i, 6) = RsG.Fields("GpSubBag")
              msGrid.TextMatrix(i, 8) = RsG.Fields("GpSubWt")
              msGrid.TextMatrix(i, 9) = RsG.Fields("GpSubRt")
              msGrid.TextMatrix(i, 10) = RsG.Fields("GpSubRtPer")
              CalGridAmt
              CalAmount

              RsG.MoveNext
              If Not RsG.EOF Then msGrid.rows = msGrid.rows + 1
           Loop
      End With

      RsG.Close
   End If
End If
End Sub
Private Function GBillNo()
Dim BilSr As String
Dim NarrCdBlTp As Long
Select Case frmTypeStr
Case "SD"
   NarrCdBlTp = GProcGetColumnValue("tblMastNarration", "Narration", "Depot", "S", "NarrCode", "N", "NarrType = 'X'")
Case "SO"
   NarrCdBlTp = GProcGetColumnValue("tblMastNarration", "Narration", "Consignment", "S", "NarrCode", "N", "NarrType = 'X'")
Case "SY"
   NarrCdBlTp = GProcGetColumnValue("tblMastNarration", "Narration", "Trade", "S", "NarrCode", "N", "NarrType = 'X'")
Case "SE"
   NarrCdBlTp = GProcGetColumnValue("tblMastNarration", "Narration", "Exempt", "S", "NarrCode", "N", "NarrType = 'X'")
Case "ST"
   NarrCdBlTp = GProcGetColumnValue("tblMastNarration", "Narration", "Sale In Transit", "S", "NarrCode", "N", "NarrType = 'X'")
End Select
SalesAccode = GProcGetColumnValue("tblMastAccount", "AcName", txtTemp(6), "S", "AcCode", "N")
If SalesAccode > 1 Then
    BilSr = ""
    BilSr = GProcGetColumnValue("tblMastBillSerial", "MillCode", "" & SalesAccode & "", "N", "BillSr", "S", "SaleType = " & NarrCdBlTp & " and VFirm = '" & gCCode & "' ")
    If BilSr = "" Then
       MsgBox ("Check Bill Sr not given...")
    Else
       txt(9) = GProcGenerateInvNo(frmTypeStr, SalesAccode, BilSr) '--Inv No
    End If
End If
End Function
Private Function GBillNoGima()
Dim BilSr As String
Dim NarrCdBlTp As Long
NarrCdBlTp = GProcGetColumnValue("tblMastNarration", "Narration", "Depot", "S", "NarrCode", "N", "NarrType = 'X'")
SalesAccode = GProcGetColumnValue("tblMastAccount", "AcName", txtTemp(6), "S", "AcCode", "N")
If SalesAccode > 1 Then
    BilSr = ""
    BilSr = GProcGetColumnValue("tblMastBillSerial", "MillCode", "" & SalesAccode & "", "N", "BillSr", "S", "SaleType = " & NarrCdBlTp & " and VFirm = '" & gCCode & "' ")
    If BilSr = "" Then
       MsgBox ("Check Bill Sr not given...")
    Else
       txt(9) = GProcGenerateInvNoGima(frmTypeStr, SalesAccode, BilSr) '--Inv No
    End If
End If
End Function
Public Sub QrCodeData()
Dim X As String
Dim x1 As String
Dim rstTmp As Recordset
Dim rstTmp2 As Recordset
Dim QRFontEncoder As QRCode
Set QRFontEncoder = New QRCode
Dim i As Integer
Dim f As Integer
Dim EcLevel As Integer
Dim DataToPrint As String
Dim DataToEncode As String
Dim OnlyCorrectData As String
Dim Printable_string As String
Dim Truncated As Integer
Dim TotalRows As Integer
Dim TotalColumns As Integer
Dim msg As String
Dim Proportional As Integer
Dim ForceBinary As Integer
Dim EncMode As Integer
Dim ProcTilde As Integer
Dim PrefFormat As Integer
Dim ErrorCorrectionLevel As Integer
Dim EncodingModeLevel As Integer
Dim setVersion As Integer
Dim QrStr As String
x1 = "Select * from tmpSale where  SlUser='" & gUserName & "'"
If GProcRstOpen(rstTmp, x1, "R") > 0 Then
   With rstTmp
        rstTmp.MoveFirst
        Do While Not .EOF
           x1 = "Select '" & gCGStin & "' ,tblMastAccount.AcGSTIN,tmpSale.SlBillNo,tmpSale.Vdt,tblMastItem.ItHsn2,tmpSale.SlIRNNo,tmpSale.SlAckNo,tmpSale.SlSubItAmt+tmpSale.SlAdd1,tmpSale.SlTaxAmt+tmpSale.SlTaxAmt2+tmpSale.SlTaxAmt3,1,tblMastAccount_Ml.AcGSTIN,tmpSale.Vtype from " _
                 & " tmpSale,tblMastAccount,tblMastItem,tblMastAccount tblMastAccount_Ml where  " _
                 & " tmpSale.SlAcDrCode = tblMastAccount.Accode  and tmpSale.SlMillCode = tblMastAccount_Ml.Accode  and tmpSale.SlSubItCode = tblMastItem.Itcode and " _
                 & " tmpSale.Vno = " & rstTmp.Fields("Vno") & " and tmpSale.Vtype = '" & rstTmp.Fields("Vtype") & "' and tmpSale.Vyear = " & rstTmp.Fields("Vyear") & " and tmpSale.VFirm = '" & rstTmp.Fields("VFirm") & "' "
           If GProcRstOpen(rstTmp2, x1, "R") > 0 Then
              If rstTmp2.Fields(11) = "SD" Then
                 QrStr = "From GSTIN : " + rstTmp2.Fields(10)             '--Firm GStin
              Else
                 QrStr = "From GSTIN : " + rstTmp2.Fields(0)             '--Firm GStin
              End If
              QrStr = QrStr + "," + " To GSTIN : " + rstTmp2.Fields(1) '--Party GSTIN
              QrStr = QrStr + "," + " Bill No : " + rstTmp2.Fields(2) '--Bill No
              QrStr = QrStr + "," + " Date : " + CStr(Format(rstTmp2.Fields(3), "DD/MM/YYYY")) '-- Bill Date
              QrStr = QrStr + "," + " Taxable : " + CStr(Format(rstTmp2.Fields(7), "0.00")) '-- Taxable
              QrStr = QrStr + "," + " GST : " + CStr(Format(rstTmp2.Fields(8), "0.00")) '-- GST
              QrStr = QrStr + "," + " Items : " + CStr(rstTmp2.Fields(9)) '-- No of Items
              QrStr = QrStr + "," + " HSN Cd : " + rstTmp2.Fields(4) '--HSN Code
              QrStr = QrStr + "," + " IRN : " + rstTmp2.Fields(5) & vbNullString  '--IRN
              QrStr = QrStr + "," + " Ack : " + rstTmp2.Fields(6) & vbNullString '--Ack No
            
              Printable_string = ""
              DataToPrint = ""
              ProcTilde = 0
              DataToEncode = QrStr
              ErrorCorrectionLevel = 0
              EncodingModeLevel = 0
              setVersion = 0
              If DataToEncode = "" Then
                 DataToEncode = ""
              End If
              QRFontEncoder.FontEncode DataToEncode, ProcTilde, EncodingModeLevel, setVersion, ErrorCorrectionLevel, Printable_string
              QrStr = ""
              QrStr = Printable_string
              gCn.BeginTrans
              gCn.Execute "update tmpSale set QrCode = '" & QrStr & "' where tmpSale.Vno = " & rstTmp.Fields("Vno") & " and tmpSale.Vtype = '" & rstTmp.Fields("Vtype") & "' and tmpSale.Vyear = " & rstTmp.Fields("Vyear") & " and tmpSale.VFirm = '" & rstTmp.Fields("VFirm") & "' "
              gCn.CommitTrans
           End If
           rstTmp.MoveNext
        Loop
   End With
End If






















'X = "Select '" & gCGStin & "' ,AcGSTIN,tblAddLess.AdInvNo,tblAddLess.Vdt,tblMastItem.ITTMP2,tblAddLess.ADTMP6,tblAddLess.ADTMP8,sum(tblVouSub.VSubTaxableAmt),sum(tblVouSub.VsubCGSTAmt+tblVouSub.VsubSGSTAmt+tblVouSub.VsubIGSTAmt),max(tblvousub.VSubCtrNo) from " _
'     & " tblAddLess,tblVouSub,tblMastAccount,tblMastItem where tblAddLess.Vno = tblVouSub.Vno and tblAddLess.Vtype = tblVouSub.Vtype  and tblAddLess.Vyear = tblVouSub.Vyear and " _
'     & " tblAddLess.AdAcDrCode = tblMastAccount.Accode and tblVouSub.VSubItCode = tblMastItem.Itcode and " _
'     & " tblAddLess.Vno = " & wVno & " and tblAddLess.Vtype = '" & wVtype & "' and tblAddLess.Vyear = " & wVyear & " group by AcGSTIN,tblAddLess.AdInvNo,tblAddLess.Vdt,tblMastItem.ITTMP2,tblAddLess.ADTMP6,tblAddLess.ADTMP8  "
'If GProcRstOpen(rstTmp, X, "R") > 0 Then
'   QrStr = "From GSTIN : " + rstTmp.Fields(0)             '--Firm GStin
'   QrStr = QrStr + "," + " To GSTIN : " + rstTmp.Fields(1) '--Party GSTIN
'   QrStr = QrStr + "," + " Bill No : " + rstTmp.Fields(2) '--Bill No
'   QrStr = QrStr + "," + " Date : " + CStr(Format(rstTmp.Fields(3), "DD/MM/YYYY")) '-- Bill Date
'   QrStr = QrStr + "," + " Taxable : " + CStr(Format(rstTmp.Fields(7), "0.00")) '-- Taxable
'   QrStr = QrStr + "," + " GST : " + CStr(Format(rstTmp.Fields(8), "0.00")) '-- GST
'   QrStr = QrStr + "," + " Items : " + CStr(rstTmp.Fields(9)) '-- No of Items
'   QrStr = QrStr + "," + " HSN Cd : " + rstTmp.Fields(4) '--HSN Code
'   QrStr = QrStr + "," + " IRN : " + rstTmp.Fields(5) & vbNullString  '--IRN
'   QrStr = QrStr + "," + " Ack : " + rstTmp.Fields(6) & vbNullString '--Ack No
'
'   Printable_string = ""
'   DataToPrint = ""
'   ProcTilde = 0
'   DataToEncode = QrStr
''  'DataToEncode = "eyJhbGciOiJSUzI1NiIsImtpZCI6IjQ0NDQwNUM3ODFFNDgyNTA3MkIzNENBNEY4QkRDNjA2Qzg2QjU3MjAiLCJ0eXAiOiJKV1QiLCJ4NXQiOiJSRVFGeDRIa2dsQnlzMHlrLUwzR0JzaHJWeUEifQ.eyJkYXRhIjoie1wiU2VsbGVyR3N0aW5cIjpcIjI3QUNMUEowNzUyQjFaWVwiLFwiQnV5ZXJHc3RpblwiOlwiMTlBQUhIUjg4MjJDMVpaXCIsXCJEb2NOb1wiOlwiMjlcIixcIkRvY1R5cFwiOlwiSU5WXCIsXCJEb2NEdFwiOlwiMDUvMDQvMjAyMVwiLFwiVG90SW52VmFsXCI6MTAwMjEzLFwiSXRlbUNudFwiOjYsXCJNYWluSHNuQ29kZVwiOlwiNTIwODEyXCIsXCJJcm5cIjpcIjEzZmEzY2YwODlmYmVhNTc0OTFmMDE0M2Q5ZGRlZjQ0M2EwMzEwY2UwMTg3ZTY4NTQxOWVkMmUwZjFjZmU3M2NcIixcIklybkR0XCI6XCIyMDIxLTA0LTA1IDEyOjIzOjAwXCJ9IiwiaXNzIjoiTklDIn0.eR56ejp14WzftM1vSGez_tFnrLW7UpfAHxfrfoNRWoRNWLwXMJ2u1hE_AToQ10hFGLeHXSvY4McqXL7aPUxNooSgOApr_PsMJoa4xV0dNkKC4E4pJAWdmvXKT_iIWRy7DVbU5hqCHFpwMN0ItQaVZjatuTdIjzrMeTKYecB0uPERT2HhRtEstl6Dn1vMKIloMFmN7Nj1h3kiKiU4ROgjdb40BR6MnKAplnsziJSz7AtowBSZdwkV7yz95xxXnpZcSpmyRk5nID_TmQ9sCNi8DwTXPrVsBPL16BMDqmqCUhbr9TUZme0EXjWvaCxxfC739FRR2fDx9SSVvEkTavR9_g"
'   ErrorCorrectionLevel = 0
'   EncodingModeLevel = 0
'   setVersion = 0
'   If DataToEncode = "" Then
'      DataToEncode = ""
'   End If
'   QRFontEncoder.FontEncode DataToEncode, ProcTilde, EncodingModeLevel, setVersion, ErrorCorrectionLevel, Printable_string
'   QrStr = ""
'   QrStr = Printable_string
'   gCn.BeginTrans
'   gCn.Execute " Insert Into tmptblQrCode (VNo,VType,VYear,QrCode,UserName) " _
'                & " Values(" & wVno & ", '" & wVtype & "'," & wVyear & ",'" & QrStr & "','" & gUserName & "')"
'   gCn.CommitTrans
'End If
End Sub

Public Sub delay(PauseTime As Integer)
    Dim start As Single
    start = Timer
   Do While Timer < start + PauseTime
    If (Timer < start) Then 'midnight crossover
        start = start - (86400 + 1)
    End If
    DoEvents   ' Yield to other processes.
    Loop
End Sub



