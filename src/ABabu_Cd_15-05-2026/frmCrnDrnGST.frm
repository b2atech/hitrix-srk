VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form frmCrnDrnGST 
   BackColor       =   &H00FBF2E1&
   BorderStyle     =   0  'None
   Caption         =   "Purchase Entry"
   ClientHeight    =   8940
   ClientLeft      =   0
   ClientTop       =   15
   ClientWidth     =   14700
   FillColor       =   &H000040C0&
   FillStyle       =   0  'Solid
   Icon            =   "frmCrnDrnGST.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   8940
   ScaleWidth      =   14700
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame1 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFC0C0&
      Caption         =   "Printing"
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
      Height          =   5295
      Left            =   3480
      TabIndex        =   113
      Top             =   1920
      Visible         =   0   'False
      Width           =   6165
      Begin VB.CommandButton cmdPrintBill 
         Caption         =   "e-Inv.Screen"
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
         Left            =   540
         TabIndex        =   137
         Top             =   4440
         Width           =   1635
      End
      Begin VB.CommandButton cmdPrintBill 
         Caption         =   "e-Invoice"
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
         Left            =   2340
         TabIndex        =   136
         Top             =   4440
         Width           =   1635
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
         Left            =   4140
         TabIndex        =   135
         Top             =   4440
         Width           =   1635
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
         Left            =   2820
         TabIndex        =   126
         Top             =   1890
         Visible         =   0   'False
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
         Left            =   2820
         TabIndex        =   125
         Top             =   1395
         Visible         =   0   'False
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
         Left            =   2820
         TabIndex        =   124
         Top             =   900
         Visible         =   0   'False
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
         Left            =   2820
         TabIndex        =   123
         Top             =   450
         Visible         =   0   'False
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
         Left            =   4245
         TabIndex        =   118
         Top             =   3690
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
         TabIndex        =   117
         Top             =   3705
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
         MaxLength       =   10
         TabIndex        =   116
         Top             =   2775
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
         TabIndex        =   115
         Top             =   3690
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
         MaxLength       =   10
         TabIndex        =   114
         Top             =   2205
         Width           =   2055
      End
      Begin VB.Label Label39 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "To         No   :"
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
         TabIndex        =   120
         Top             =   2865
         Width           =   1275
      End
      Begin VB.Label Label38 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "From    No    :"
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
         TabIndex        =   119
         Top             =   2295
         Width           =   1290
      End
      Begin VB.Line Line3 
         X1              =   0
         X2              =   6120
         Y1              =   3435
         Y2              =   3435
      End
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
      Height          =   735
      Index           =   36
      Left            =   1440
      MaxLength       =   64
      MultiLine       =   -1  'True
      TabIndex        =   26
      Text            =   "frmCrnDrnGST.frx":058A
      Top             =   7920
      Width           =   4830
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
      Height          =   375
      Index           =   37
      Left            =   7500
      TabIndex        =   27
      Text            =   "37"
      Top             =   8280
      Width           =   4890
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
      Left            =   1740
      TabIndex        =   13
      Text            =   "txtCode(6)"
      ToolTipText     =   "Press F3 for Creation"
      Top             =   5700
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
      Index           =   31
      Left            =   7560
      TabIndex        =   18
      Text            =   "31"
      Top             =   7192
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
      Index           =   47
      Left            =   19080
      TabIndex        =   110
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
      TabIndex        =   109
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
      TabIndex        =   108
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
      Index           =   11
      Left            =   9015
      TabIndex        =   11
      Text            =   "txtCode(11)"
      ToolTipText     =   "Press F3 for Creation"
      Top             =   3060
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
      Left            =   1440
      TabIndex        =   3
      Text            =   "txtCode(7)"
      ToolTipText     =   "Press F3 for Creation"
      Top             =   3060
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
      TabIndex        =   106
      Text            =   "44"
      Top             =   3600
      Visible         =   0   'False
      Width           =   570
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
      Left            =   16065
      TabIndex        =   105
      Text            =   "43"
      Top             =   9045
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
      Height          =   405
      Index           =   32
      Left            =   10140
      TabIndex        =   22
      Text            =   "32"
      Top             =   7185
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
      Index           =   41
      Left            =   15975
      TabIndex        =   104
      Text            =   "41"
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
      Height          =   375
      Index           =   40
      Left            =   19080
      TabIndex        =   94
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
      TabIndex        =   93
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
      Left            =   16560
      TabIndex        =   23
      Text            =   "38"
      Top             =   8145
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
      Index           =   9
      Left            =   6375
      TabIndex        =   5
      Text            =   "9"
      Top             =   720
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
      Index           =   370
      Left            =   18810
      TabIndex        =   92
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
      Index           =   360
      Left            =   18540
      TabIndex        =   91
      Text            =   "360"
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
      Left            =   13305
      Style           =   1  'Graphical
      TabIndex        =   89
      Top             =   5730
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
      TabIndex        =   88
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
      Left            =   17460
      MaxLength       =   5
      TabIndex        =   39
      Text            =   "txtemp(35)"
      Top             =   5715
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
      Index           =   22
      Left            =   18360
      TabIndex        =   82
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
      Left            =   15795
      TabIndex        =   38
      Text            =   "16"
      Top             =   9360
      Visible         =   0   'False
      Width           =   465
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
      Index           =   25
      Left            =   16065
      TabIndex        =   32
      Text            =   "25"
      Top             =   5220
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
      Index           =   14
      Left            =   18360
      TabIndex        =   81
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
      Left            =   16245
      TabIndex        =   30
      Text            =   "13"
      Top             =   2790
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
      Height          =   405
      Index           =   28
      Left            =   10140
      TabIndex        =   19
      Text            =   "28"
      Top             =   6285
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
      Index           =   29
      Left            =   7560
      TabIndex        =   17
      Text            =   "29"
      Top             =   6742
      Width           =   690
   End
   Begin VB.CommandButton CmdBagDetails 
      Caption         =   "Cartoon &Bag Details"
      Height          =   720
      Left            =   17460
      TabIndex        =   77
      Top             =   4635
      Visible         =   0   'False
      Width           =   870
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
      Height          =   420
      Index           =   8
      Left            =   10140
      TabIndex        =   24
      Text            =   "8"
      Top             =   7635
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
      Height          =   375
      Index           =   34
      Left            =   18585
      TabIndex        =   76
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
      Left            =   17145
      TabIndex        =   36
      Text            =   "19"
      Top             =   8955
      Visible         =   0   'False
      Width           =   285
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
      Index           =   33
      Left            =   14940
      MaxLength       =   10
      TabIndex        =   9
      Text            =   "33"
      Top             =   6240
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
      Index           =   11
      Left            =   9015
      TabIndex        =   12
      Text            =   "txttemp(11)"
      ToolTipText     =   "Press F3 for Creation"
      Top             =   3600
      Width           =   3540
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
      Left            =   1440
      TabIndex        =   4
      Text            =   "txttemp(7)"
      ToolTipText     =   "Press F3 for Creation"
      Top             =   3600
      Width           =   6180
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
      Left            =   15885
      TabIndex        =   34
      Text            =   "txtDrBalance"
      Top             =   1170
      Width           =   300
   End
   Begin VB.CommandButton cmdLedger 
      Caption         =   "Ledger"
      Height          =   375
      Left            =   12915
      TabIndex        =   75
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
      Height          =   405
      Index           =   27
      Left            =   7560
      TabIndex        =   16
      Text            =   "27"
      Top             =   6292
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
      Index           =   26
      Left            =   16020
      TabIndex        =   51
      Text            =   "26"
      Top             =   8190
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
      Index           =   15
      Left            =   10440
      TabIndex        =   7
      Text            =   "15"
      Top             =   720
      Width           =   1740
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
      Left            =   10140
      TabIndex        =   15
      Text            =   "21"
      Top             =   5700
      Width           =   2265
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
      TabIndex        =   73
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
      TabIndex        =   71
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
      TabIndex        =   70
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
      TabIndex        =   69
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
      TabIndex        =   65
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
      TabIndex        =   64
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
      Left            =   13305
      Style           =   1  'Graphical
      TabIndex        =   62
      ToolTipText     =   "Exit"
      Top             =   7110
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
      Left            =   13305
      Style           =   1  'Graphical
      TabIndex        =   55
      ToolTipText     =   "To Save Record"
      Top             =   5085
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
      Left            =   13305
      Style           =   1  'Graphical
      TabIndex        =   63
      Top             =   6420
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
      Left            =   13305
      Style           =   1  'Graphical
      TabIndex        =   54
      ToolTipText     =   "To Delete Record"
      Top             =   4395
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
      Left            =   13305
      Style           =   1  'Graphical
      TabIndex        =   53
      ToolTipText     =   "To Modify Record"
      Top             =   3600
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
      Left            =   13305
      Style           =   1  'Graphical
      TabIndex        =   52
      ToolTipText     =   "To Add New Record"
      Top             =   2955
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
      TabIndex        =   56
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
      Left            =   13305
      Style           =   1  'Graphical
      TabIndex        =   57
      ToolTipText     =   "To Find Record"
      Top             =   2325
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
      Left            =   13800
      Style           =   1  'Graphical
      TabIndex        =   61
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
      Left            =   13305
      Style           =   1  'Graphical
      TabIndex        =   60
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
      Left            =   13800
      Style           =   1  'Graphical
      TabIndex        =   59
      ToolTipText     =   "Next"
      Top             =   1140
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
      Height          =   390
      Index           =   0
      Left            =   1680
      TabIndex        =   1
      Text            =   "0"
      Top             =   720
      Width           =   1740
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
      Left            =   13305
      Style           =   1  'Graphical
      TabIndex        =   58
      ToolTipText     =   "Previous"
      Top             =   1140
      Width           =   375
   End
   Begin VB.TextBox txt 
      Alignment       =   1  'Right Justify
      Height          =   315
      Index           =   4
      Left            =   18270
      TabIndex        =   68
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
      Height          =   405
      Index           =   6
      Left            =   1740
      Locked          =   -1  'True
      TabIndex        =   14
      Text            =   "txttemp(6)"
      ToolTipText     =   "Press F3 for Creation"
      Top             =   6180
      Width           =   4500
   End
   Begin VB.TextBox txt 
      Height          =   315
      Index           =   6
      Left            =   18405
      TabIndex        =   67
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
      TabIndex        =   66
      Text            =   "7"
      Top             =   2565
      Visible         =   0   'False
      Width           =   525
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
      Height          =   825
      Index           =   10
      Left            =   1440
      MaxLength       =   50
      MultiLine       =   -1  'True
      TabIndex        =   25
      Text            =   "frmCrnDrnGST.frx":058D
      Top             =   6720
      Width           =   4800
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
      TabIndex        =   49
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
      Left            =   16155
      TabIndex        =   29
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
      Left            =   17040
      TabIndex        =   10
      Text            =   "23"
      Top             =   420
      Visible         =   0   'False
      Width           =   660
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
      Index           =   24
      Left            =   16650
      TabIndex        =   37
      Text            =   "24"
      Top             =   8550
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
      Height          =   360
      Index           =   20
      Left            =   17640
      TabIndex        =   41
      Text            =   "20"
      Top             =   9135
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
      Height          =   405
      Index           =   30
      Left            =   10140
      TabIndex        =   20
      Text            =   "30"
      Top             =   6735
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
      Index           =   17
      Left            =   16470
      TabIndex        =   33
      Text            =   "17"
      Top             =   9405
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
      Index           =   18
      Left            =   18945
      TabIndex        =   35
      Text            =   "18"
      Top             =   2295
      Visible         =   0   'False
      Width           =   330
   End
   Begin MSMask.MaskEdBox mskDt 
      Height          =   390
      Index           =   2
      Left            =   15165
      TabIndex        =   21
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
      Left            =   15120
      TabIndex        =   28
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
      Left            =   15210
      TabIndex        =   40
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
      Bindings        =   "frmCrnDrnGST.frx":0590
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
   Begin MSMask.MaskEdBox mskDt 
      Height          =   375
      Index           =   14
      Left            =   16830
      TabIndex        =   31
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
      TabIndex        =   95
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
      Height          =   390
      Index           =   2
      Left            =   1680
      TabIndex        =   2
      Top             =   1500
      Width           =   1740
      _ExtentX        =   3069
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
      Format          =   115015683
      CurrentDate     =   -273
   End
   Begin MSComCtl2.DTPicker dtpDt 
      Height          =   390
      Index           =   18
      Left            =   10440
      TabIndex        =   8
      Top             =   1500
      Width           =   1740
      _ExtentX        =   3069
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
      Format          =   115015683
      CurrentDate     =   -273
   End
   Begin MSComCtl2.DTPicker dtpDt 
      Height          =   390
      Index           =   12
      Left            =   6375
      TabIndex        =   6
      Top             =   1500
      Width           =   1740
      _ExtentX        =   3069
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
      Format          =   115015683
      CurrentDate     =   -273
   End
   Begin VB.Label Label9 
      BackStyle       =   0  'Transparent
      Caption         =   "Nar.       :"
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
      Height          =   330
      Left            =   405
      TabIndex        =   140
      Top             =   6720
      Width           =   1260
   End
   Begin VB.Line Line2 
      X1              =   6360
      X2              =   12780
      Y1              =   8100
      Y2              =   8100
   End
   Begin VB.Line Line1 
      X1              =   6360
      X2              =   6360
      Y1              =   5520
      Y2              =   8160
   End
   Begin VB.Label Label12 
      BackStyle       =   0  'Transparent
      Caption         =   "IRN No   :"
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
      Height          =   330
      Left            =   405
      TabIndex        =   139
      Top             =   7920
      Width           =   1260
   End
   Begin VB.Label Label15 
      BackStyle       =   0  'Transparent
      Caption         =   "Ack No.   :"
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
      Height          =   330
      Left            =   6405
      TabIndex        =   138
      Top             =   8280
      Width           =   1260
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "Ref. Date : "
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
      Left            =   5220
      TabIndex        =   134
      Top             =   1560
      Width           =   1095
   End
   Begin VB.Label lblAcc1 
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
      Left            =   420
      TabIndex        =   133
      Top             =   2580
      Width           =   1620
   End
   Begin VB.Label Label11 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Entry No  :"
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
      Left            =   540
      TabIndex        =   132
      Top             =   720
      Width           =   975
   End
   Begin VB.Label lblHsn 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "HSN Code"
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
      Left            =   10980
      TabIndex        =   131
      Top             =   4920
      Width           =   960
   End
   Begin VB.Label Label10 
      BackStyle       =   0  'Transparent
      Caption         =   "HSN Code  :"
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
      Height          =   330
      Left            =   9720
      TabIndex        =   130
      Top             =   4920
      Width           =   1260
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "R. Bill Dt : "
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
      Left            =   9360
      TabIndex        =   129
      Top             =   1560
      Width           =   1095
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Ref. Bill   : "
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
      Left            =   9360
      TabIndex        =   128
      Top             =   840
      Width           =   1095
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Ref. No    :"
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
      Left            =   5160
      TabIndex        =   127
      Top             =   735
      Width           =   990
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
      Left            =   1440
      TabIndex        =   122
      Top             =   4500
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
      Left            =   1440
      TabIndex        =   121
      Top             =   4140
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
      Left            =   12960
      TabIndex        =   112
      Top             =   9135
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
      Left            =   15420
      TabIndex        =   111
      Top             =   8820
      Visible         =   0   'False
      Width           =   1530
   End
   Begin VB.Shape Shape3 
      BorderColor     =   &H00EFB17A&
      BorderWidth     =   5
      FillColor       =   &H00EFB17A&
      Height          =   13140
      Left            =   0
      Top             =   -4200
      Width           =   14520
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Date        : "
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
      TabIndex        =   107
      Top             =   1560
      Width           =   1095
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
      Left            =   8370
      TabIndex        =   103
      Top             =   7200
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
      Left            =   8370
      TabIndex        =   102
      Top             =   6750
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
      Left            =   8640
      TabIndex        =   101
      Top             =   6300
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
      Left            =   8640
      TabIndex        =   100
      Top             =   6750
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
      Left            =   8640
      TabIndex        =   99
      Top             =   7200
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
      Left            =   6480
      TabIndex        =   98
      Top             =   6300
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
      Left            =   6450
      TabIndex        =   97
      Top             =   6750
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
      Left            =   6480
      TabIndex        =   96
      Top             =   7200
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
      Left            =   6840
      TabIndex        =   90
      Top             =   4905
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
      Left            =   2430
      TabIndex        =   87
      Top             =   4905
      Width           =   3090
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
      Left            =   1440
      TabIndex        =   86
      Top             =   4905
      Width           =   840
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
      Left            =   15840
      TabIndex        =   85
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
      Left            =   15840
      TabIndex        =   84
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
      Left            =   15840
      TabIndex        =   83
      Top             =   6345
      Width           =   1290
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
      Left            =   14850
      TabIndex        =   80
      Top             =   2925
      Visible         =   0   'False
      Width           =   1320
   End
   Begin VB.Label Label13 
      BackStyle       =   0  'Transparent
      Caption         =   "Vehicle No   :    "
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
      Left            =   14895
      TabIndex        =   79
      Top             =   2385
      Visible         =   0   'False
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
      Left            =   14850
      TabIndex        =   78
      Top             =   1845
      Visible         =   0   'False
      Width           =   1320
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "&Party     :"
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
      Left            =   450
      TabIndex        =   42
      Top             =   3075
      Width           =   900
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "&Broker    : "
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
      Left            =   8040
      TabIndex        =   43
      Top             =   3075
      Width           =   870
   End
   Begin VB.Label lblAcc2 
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
      Left            =   420
      TabIndex        =   0
      Top             =   5715
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
      TabIndex        =   50
      Top             =   8505
      Width           =   1335
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
      Left            =   8520
      TabIndex        =   48
      Top             =   7680
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
      Left            =   8115
      TabIndex        =   45
      Top             =   5745
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
      Left            =   11745
      TabIndex        =   74
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
      Left            =   6450
      TabIndex        =   47
      Top             =   7725
      Visible         =   0   'False
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
      Left            =   17190
      TabIndex        =   44
      Top             =   8190
      Width           =   1440
   End
   Begin VB.Label LblSaleType 
      BackColor       =   &H00EFB17A&
      Caption         =   "     Credit Note / Debit Note"
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
      TabIndex        =   72
      Top             =   0
      Width           =   14715
   End
   Begin VB.Shape Shape1 
      FillColor       =   &H00FED8FC&
      FillStyle       =   0  'Solid
      Height          =   7830
      Left            =   13050
      Shape           =   4  'Rounded Rectangle
      Top             =   540
      Width           =   1350
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
      Left            =   8370
      TabIndex        =   46
      Top             =   6300
      Width           =   255
   End
   Begin VB.Shape Shape2 
      FillColor       =   &H00EDFDFE&
      FillStyle       =   0  'Solid
      Height          =   2925
      Left            =   180
      Top             =   2460
      Width           =   12540
   End
   Begin VB.Shape Shape5 
      FillColor       =   &H00EDFDFE&
      FillStyle       =   0  'Solid
      Height          =   3180
      Left            =   195
      Top             =   5550
      Width           =   12555
   End
   Begin VB.Shape Shape7 
      FillColor       =   &H00EDFDFE&
      FillStyle       =   0  'Solid
      Height          =   1620
      Left            =   300
      Top             =   540
      Width           =   4455
   End
   Begin VB.Shape Shape4 
      FillColor       =   &H00EDFDFE&
      FillStyle       =   0  'Solid
      Height          =   1620
      Left            =   4980
      Top             =   540
      Width           =   7755
   End
End
Attribute VB_Name = "frmCrnDrnGST"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Const tblName As String = "tblIntSale"
Const IdField As String = "Vno"
Const OrderField As String = "Vdt,CrDrNoteNo"
Const MaxNo As Long = 37
Const OtherTxtFromNo As Long = 0 '-- Other Txt Ctrl From No
Const OtherTxtToNo As Long = 0 '-- Other Txt Ctrl Less No
Dim rstbl As Recordset '-- TblSale
Dim rstblSub As Recordset '-- TblSaleSub
Dim varBookMark As Variant
Dim FirstTimeFlag As Boolean
Dim OldgFrmTypeStr As String
Dim FormAction As Integer
Dim frmTypeStr As String
Dim mLastEntryDate As Date
Dim mLastEntryMill As String
Dim i As Long
Dim SalesAccode As Long
Dim Criteria As String
Dim ItemType As Integer
Dim OldBag As Integer
Dim TrnCd As String
Dim BillsInFirm As String
Dim rstblBlChk As Recordset
Dim Mon As Integer
'Dim GridFlag As Boolean
Private Sub cmdBtn_Click(Index As Integer)
Dim X As String
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
        If GProcRstOpen(rstbl, "Select * from " & tblName & " where " & Criteria & "  and  " & IdField & " = (Select max(" & IdField & ") from " & tblName & " where vdt = (SELECT MAX(vdt) from tblIntSale where " & Criteria & " )  and  " & IdField & " <> 0 and  " & Criteria & ")", "R") > 0 Then ReadFields Else GProcClearForm Me, rstbl, MaxNo
Case 4 'Find ' QryMastAccount
        gClsSearch.SearchMultiField "tblIntSale,TblMastAccount", "vno,CrDrNoteNo,Vdt,AcName", Array("Vno", "Bill No", "Vdt", "Party"), Array(0, 1500, 1500, 4000), " tblIntSale.VDrAcCode=TblMastAccount.AcCode and vtype='" & frmTypeStr & "'   and VFirm = '" & gCCode & "'   and VYear=" & gCYear, "", "(cast(substring(CrDrNoteNo,7,15) as int) ) desc", 0, 0
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
        dtpDt(18) = Format(Now, "dd/MM/yyyy")
    Else
        dtpDt(2) = mLastEntryDate
        dtpDt(18) = mLastEntryDate
    End If
    If FirstTimeFlag = True Then
        FirstTimeFlag = False
    Else
       If txtCode(7).Enabled = True Then
          txtCode(7).SetFocus
       End If
    End If
    '-- Sales A/C
'    txttemp(6) = GProcGetColumnValue("tblMastAccount", "AcCode", CStr(SalesAccode), "N", "AcName", "S")
'    If gfrmTypeStr = "PV" Then
'       txt(33) = GProcGenerateNoWithIBCrDrNt("PV", gCCode)
'    Else
'       txt(33) = GProcGenerateNoWithIBCrDrNt("SV", gCCode)
'    End If
    lblTinNo.Caption = ""
    lblMsOms.Caption = ""
Case 7 'Modify
    'If mURecModi = True Then
    '    SearchHelp
        DoEvents
        If IsRelatedRecord = False Then
            GProcRstOpen rstbl, "Select * from " & tblName & " where " & IdField & " = " & txt(0) & " and " & Criteria, "R"
            ReadFields
            FormAction = vbDataActionUpdate
            Call GProcActivateControls(Me, True)
            SetControlEd
            Call GProcSetButtons(Me, vbDataActionUpdate)
            txt(0).Enabled = False
'            txt(33).Enabled = False
            txtCode(7).SetFocus
        End If
    'End If
Case 8 'Delete
    'If mURecDel = True Then
      If IsRelatedRecord = False Then
        If MsgBox("Do You Really Delete This Record?", vbQuestion + vbYesNo + vbDefaultButton1, "Confirm Delete") = vbYes Then
           If gCYear >= 2023 And gCIsLog = 1 Then
              gCn.BeginTrans
              X = "Select max(LogNo) from tblIntSale_Log where Vno=" & Val(txt(0)) & " and " & Criteria & ""
              j = GProcRstOpen(tmpRst, X, "R")
              If j > 0 Then
                 wLogNo = IIf(IsNull(tmpRst.Fields(0)), 0, tmpRst.Fields(0)) + 1
              Else
                 wLogNo = 1
              End If
              gCn.Execute "Insert into tblIntSale_Log select tblIntSale.*," & wLogNo & " as LogNo,'D' as LogTp ,'" & gUserName & "' as UserName ,'" & Format(Now, "YYYY/MM/DD hh:mm:ss") & "' as LogDate from tblIntSale where Vno=" & Val(txt(0)) & " and " & Criteria & ""
              gCn.CommitTrans
              tmpRst.Close
           End If
'           DeleteRelatedRecord
           GProcRstOpen rstbl, "Select * from " & tblName & " where " & IdField & " = " & txt(0) & " and " & Criteria, "O"
           rstbl.Delete
           cmdBtn_Click (3)
        End If
    End If
Case 9 'Print
    Frame1.Caption = "Sales Bill Printing "
    '-- Inv No
    txtPInvno(0) = txt(0).text
    txtPInvno(1) = txt(0).text
    Frame1.Visible = True
    DoEvents
    txtPInvno(0).Enabled = True
    txtPInvno(1).Enabled = True
    chkInvChallan(0).Enabled = True
    chkInvChallan(1).Enabled = True
    chkInvChallan(2).Enabled = True
    chkInvChallan(3).Enabled = True
    txtPInvno(0).SetFocus
    txtPInvno(0).SelStart = Len(txtPInvno(0))

Case 10 'Exit
'    DoEvents
    Unload Me
    Exit Sub
Case 11 'Label
    'frmLabel.Show 1
Case 12 'Optional
Case 13 'Save
        'txt(51) = CDate(Date)
'        If FormAction = vbDataActionAddNew Then txt(0) = GProcGenerateId(tblName, IdField, "vtype='" & frmTypeStr & "' and VYear=" & gCYear)
        CalAmount
        FillTxtFromTemp
        '-- Save Record
        If ValidateData = True Then
'            gCn.Execute "delete from " & tblName & "  where vtype='" & frmTypeStr & "' and vno=" & txt(0) & " and Vfirm = '" & gCCode & "'  and " _
'            & " VYear=" & gCYear
           DoEvents
           DoEvents
           DoEvents
            If FormAction = vbDataActionAddNew Then
               Mon = Month(dtpDt(2))
               txt(0) = GProcGenerateIdMonthwise("tblIntSale", "Vno", Mon, "Vdt", "Vtype='" & frmTypeStr & "' and VFirm = '" & gCCode & "'   and VYear=" & gCYear)
            End If
            txt(5) = 1
            txt(35) = 1
            gCn.BeginTrans
            GProcRstOpen rstbl, "Select * from " & tblName & " where " & IdField & " = " & txt(0) & " and " & Criteria, "O"
            GProcSaveRecord Me, rstbl, FormAction, MaxNo, OtherTxtFromNo, OtherTxtToNo
            gCn.CommitTrans
            '-- For Log
            If gCYear >= 2023 And gCIsLog = 1 Then
               If FormAction = vbDataActionAddNew Then
                  AMD = "A"
               Else
                  AMD = "M"
               End If
               X = "Select max(LogNo) from tblIntSale_Log where Vno=" & Val(txt(0)) & " and " & Criteria & ""
               j = GProcRstOpen(tmpRst, X, "R")
               If j > 0 Then
                  wLogNo = IIf(IsNull(tmpRst.Fields(0)), 0, tmpRst.Fields(0)) + 1
               Else
                  wLogNo = 1
               End If
               X = "select tblIntSale.*," & wLogNo & " as LogNo,'" & AMD & "' as LogTp ,'" & gUserName & "' as UserName ,'" & Format(Now, "YYYY/MM/DD hh:mm:ss") & "' as LogDate from tblIntSale where Vno=" & Val(txt(0)) & " and " & Criteria & ""
               gCn.Execute "Insert into tblIntSale_Log select tblIntSale.*," & wLogNo & " as LogNo,'" & AMD & "' as LogTp ,'" & gUserName & "' as UserName ,'" & Format(Now, "YYYY/MM/DD hh:mm:ss") & "' as LogDate from tblIntSale where Vno=" & Val(txt(0)) & " and " & Criteria & ""
               tmpRst.Close
            End If
'            SaveRelatedRecord
            mLastEntryDate = dtpDt(2)
            mLastEntryMill = txtCode(7)
            FormAction = vbDataActionClose
            Call GProcActivateControls(Me, False)
            Call GProcSetButtons(Me, True)
'            FormAction = vbDataActionClose
'            If FormAction = vbDataActionAddNew Then
                If cmdBtn(6).Visible = True Then cmdBtn(6).SetFocus
 '           Else
  '              If cmdBtn(7).Visible = True Then cmdBtn(7).SetFocus
   '         End If
            ReadFields
            FormAction = vbDataActionClose
'            Call GProcActivateControls(Me, False)
'            Call GProcSetButtons(Me, True)
            cmdBtn(13).Enabled = False
        End If
Case 14 'Cancel
    FormAction = vbDataActionCancel
    cmdBtn_Click (3)
    If rstbl.RecordCount <= 0 Then
        Call GProcActivateControls(Me, False)
        Call GProcSetButtons(Me, vbDataActionCancel)
    Else
        Call GProcActivateControls(Me, False)
        Call GProcSetButtons(Me, True)
    End If
    If cmdBtn(4).Enabled = True Then cmdBtn(4).SetFocus
End Select
Exit Sub
ErrorRoutine:
    rstbl.CancelUpdate
    GProcErrorHandler
End Sub
Private Sub SetControlEd()
If frmTypeStr = "PV" Or gfrmTypeStr = "PX" Then
   lblAcc2.Caption = "Debit Account :"
   lblAcc1.Caption = "Credit Account :"
   If gfrmTypeStr = "PX" Then
      LblSaleType.Caption = "           Credit Note With GST (Sales)"
   Else
      LblSaleType.Caption = "           Credit Note With GST (Purchase)"
   End If
End If
If frmTypeStr = "SV" Or gfrmTypeStr = "SI" Then
   lblAcc1.Caption = "Debit Account :"
   lblAcc2.Caption = "Credit Account :"
   If gfrmTypeStr = "SI" Then
      LblSaleType.Caption = "           Debit Note With GST (Sales)"
   Else
      LblSaleType.Caption = "           Debit Note With GST (Purchase)"
   End If
End If

'txt(8).Enabled = False '-- Amt
'If FormAction = vbDataActionUpdate Then
'   txt(33).Enabled = False
'End If
If FormAction = vbDataActionAddNew Then
   txt(0).Enabled = False
End If
End Sub
Private Sub DeleteRelatedRecord()
''--- Voucher Sub
'gCn.Execute "delete from tblSaleSub where vtype='" & frmTypeStr & "' and vno=" & txt(0) & " and Vfirm = '" & gCCode & "'  and " _
'& " VYear=" & gCYear
End Sub
Private Sub SaveRelatedRecord()
'Dim i As Long
'Dim MillCode As Long
'Dim rstSub As Recordset
'Dim CrCode As Long
'Dim DrCode As Long
'Dim BrkCode As Long
'Dim TaxAcCode As Long
'Dim X As String
'Dim BillDt
'Dim DueDt
''--- Account Code
'MillCode = GProcGetColumnValue("tblMastAccount", "AcName", txttemp(7), "S", "AcCode", "N")
'DrCode = GProcGetColumnValue("tblMastAccount", "AcName", txttemp(5), "S", "AcCode", "N")
'CrCode = GProcGetColumnValue("tblMastAccount", "AcName", txttemp(6), "S", "AcCode", "N")
'BrkCode = GProcGetColumnValue("tblMastAccount", "AcName", txttemp(8), "S", "AcCode", "N")
''-- Delete Records
'DeleteRelatedRecord
''--- Save Grid Records
''--- TblVouSub
'X = "select * from tblSaleSub where vtype='" & frmTypeStr & "' and vno=" & txt(0) & " and Vfirm = '" & gCCode & "' and " _
'& " VYear=" & gCYear
'GProcRstOpen rstSub, X, "O"
'With msGrid
'    i = 1
'    Do While i < .rows
'        rstSub.AddNew
'        rstSub!Vno = CLng(txt(0)) '--Vno(0)
'        rstSub!Vtype = txt(1) & vbNullString  '--Vtype(1)
'        rstSub!Vdt = CDate(txt(2)) '--Vdt(2)
'        rstSub!VYear = txt(3) '--VYear
'        rstSub!VFirm = gCCode '--VFirm
'        rstSub!SlSubItSrNo = i '--Sr No
'        '--- It Ctrl No(11)
'        '--Pending
''        If CLng(.TextMatrix(i, 14)) = 0 Then 'FormAction = vbDataActionAddNew Then
''            rstSub!VSubItCtrlNo = GProcGenerateId("tblPurchSub", "PurSubItCtrlNo", " Vfirm = '" & gCCode & "' and  VYear=" & gCYear)   '--It Ctrl No
''        Else
''            rstSub!VSubItCtrlNo = CLng(.TextMatrix(i, 14))
''        End If
'        '--end pending
'        rstSub!SlSubPItCtrlNo = Val(.TextMatrix(i, 14))
'        rstSub!SlSubBookNo = CLng(.TextMatrix(i, 1)) '--Book No
'        '--Book Dt
''        If CLng(.TextMatrix(i, 1)) = 0 Then '-- if Book No =0
''            .TextMatrix(i, 2) = "" '--Book Dt =Null
''            .TextMatrix(i, 19) = 0 ' Booking It ctrl No =0
''        End If
'        If IsDate(.TextMatrix(i, 2)) Then
'            rstSub!SlSubBookDt = .TextMatrix(i, 2)
'        Else
'            rstSub!SlSubBookDt = Null
'        End If
'        '--Itcode(4)
'        rstSub!SlSubItCode = GProcGetColumnValue("tblMastItem", "ItName", .TextMatrix(i, 3), "S", "ItCode", "N", " ItMillCode=" & MillCode) '--Item Code
'        rstSub!SlSubBag = CLng(.TextMatrix(i, 6)) '--Bag(6)
'        rstSub!SlSubwt = CDbl(.TextMatrix(i, 8)) '-- Wt(7)
'        rstSub!SlSubRt = CDbl(.TextMatrix(i, 9)) '--Rate(8)
'        rstSub!SlSubRtPer = CDbl(.TextMatrix(i, 10)) '--Rate
'        rstSub!SlSubNetRate = CDbl(.TextMatrix(i, 11)) '--Net Rate
'        rstSub!SlSubAmt = CDbl(.TextMatrix(i, 12)) '--Amt(10)
'        rstSub!SlSubLotNo = .TextMatrix(i, 5) '--Lot No(5)
'        '--Godown Code(4)
'        .TextMatrix(i, 4) = "-"
'        rstSub!SlSubGodown = GProcGetColumnValue("tblMastNarration", "Narration", .TextMatrix(i, 4), "S", "NarrCode", "N", "")
'        '-- P It Ctrl No =Booking It ctrl No
'        rstSub.Update
'        i = i + 1
'     Loop
' End With
' rstSub.Close
' cmdBtn(13).Enabled = False
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
Select Case Index
Case 0, 1:
    If Index = 0 Then GProcCrystalRptPreparation cryReport, "Printer"
    If Index = 1 Then GProcCrystalRptPreparation cryReport, "Window"
    If Index = 3 Then GProcCrystalRptPreparation cryReport, "Window"
    
    cryReport.WindowTitle = "Print"
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
    X = "Select a.Vno,a.Vtype,a.Vdt,a.Vyear,a.VFirm,a.VDrAcCode,a.VCrAcCode,a.Broker,a.VBillNo,a.VBillDate, " _
        & "a.Interest,a.VAmt,a.Interest,a.CgstRt,a.CgstAmt,a.SgstRt,a.SgstAmt, " _
        & " '" & gUserName & "',a.IgstRt, " _
        & " a.IgstAmt,a.ItCode,a.SlIRNNo,a.SlAckNo,a.Vnar2 " _
        & " from tblIntSale A " _
        & " where a.Vno in " & X & " and a.Vtype='" & frmTypeStr & "' and a.VFirm = '" & gCCode & "' And a.VYear = " & gCYear
   
   gCn.Execute "Delete from tmpSale where  SlUser='" & gUserName & "'"
     
   gCn.Execute " insert into tmpSale (Vno,Vtype,Vdt,Vyear,VFirm,SlAcDrCode,SlAcCrCode,SlBroker,SlBillNo,SlBillDt, " _
               & "SlSubAmt,SlBillAmt,SlTaxableAmt,SlTaxRate,SlTaxAmt,SlTaxRate2,SlTaxAmt2," _
               & " SlUser,SlTaxRate3,SlTaxAmt3,SlSubItCode,SlIRNNo,SlAckNo,SlTmpChr3) " & X
     With cryReport
        .Formulas(10) = "wCSTNo = '" & gCCST & "'"
        .Formulas(11) = "wBSTNo = '" & gCBST & "'"
        If frmTypeStr = "SD" Then
           GstinDepot = GProcGetColumnValue("tblMastAccount", "AcCode", txt(7), "N", "AcGSTIN", "S")
           PanDepot = GProcGetColumnValue("tblMastAccount", "AcCode", txt(7), "N", "AcPAN", "S")
           .Formulas(12) = "wGSTNo = '" & GstinDepot & "'"
           .Formulas(13) = "wcPan = '" & PanDepot & "'"
        Else
           .Formulas(12) = "wGSTNo = '" & gCGStin & "'"
           .Formulas(13) = "wcPan = '" & gCPAN & "'"
        End If
        X = "{tmpSale.SlUser}='" & gUserName & "' "
        .SelectionFormula = X
        PrintFor = chkInvChallan(0).Caption
'        .Formulas(14) = "wPrintFor = '" & PrintFor & "'"
        .ReportFileName = gReportPath & "rptCrDrNoteGST.rpt"
        If Index = 0 Or Index = 1 Then
          .Action = 1
        End If
     End With
Case 2:
    DoEvents
    Frame1.Visible = False
Case 3, 4:
    If Val(txtPInvno(0)) > 0 Then
       InvNo = CLng(txtPInvno(0))
       ToInvNo = CLng(txtPInvno(1))
       BillSrNo = ""
    Else
       InvNo = CLng(Mid(txtPInvno(0), 7, 15))
       ToInvNo = CLng(Mid(txtPInvno(1), 7, 15))
       BillSrNo = Left(txtPInvno(0), 6)
    End If
    
    '------ Validate
    If IsNumeric(InvNo) = False Then InvNo = 0
    If IsNumeric(ToInvNo) = False Then ToInvNo = InvNo
    If ToInvNo < InvNo Then ToInvNo = InvNo
    
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
    X = "Select a.Vno,a.Vtype,a.Vdt,a.Vyear,a.VFirm,a.VDrAcCode,a.VCrAcCode,a.Broker,a.CrDrNoteNo,a.Interest,a.VAmt,a.CgstRt,a.CgstAmt,a.SgstRt,a.SgstAmt, " _
        & " a.VAmt-a.Interest-a.CgstAmt-a.SgstAmt-a.IgstAmt,'" & gUserName & "',a.IgstRt,a.IgstAmt,a.ItCode " _
        & "  from tblIntSale A " _
        & " where  a.Vno in " & X & " and a.Vtype='" & frmTypeStr & "' and a.VFirm = '" & gCCode & "' And a.VYear = " & gCYear
   gCn.Execute "Delete from tmpSale where  SlUser='" & gUserName & "'"
     
   gCn.Execute " insert into tmpSale (Vno,Vtype,Vdt,Vyear,VFirm,SlAcDrCode,SlAcCrCode,SlBroker,SlBillNo, " _
               & "SlSubAmt,SlBillAmt,SlTaxRate,SlTaxAmt,SlTaxRate2,SlTaxAmt2,SlRoff,SlUser,SlTaxRate3,SlTaxAmt3,SlSubItCode " _
               & " ) " & X

    XX = "Select * from tmpSale where  SlUser='" & gUserName & "'"
    XX = XX & " order by SlBillNo"
    If GProcRstOpen(RsSA, X, "R") <= 0 Then
       MsgBox "No Records..."
       Exit Sub
    End If
    
    If Index = 3 Then
        GProcCrystalRptPreparation cryReport, "Window"
        With cryReport
             .Formulas(4) = "wCompGstin = '" & gCGStin & "'"
             .SelectionFormula = "{tmpSale.SlUser} = '" & gUserName & "' "
             .ReportFileName = gReportPath & "rpte-Invoice-IntDb.rpt"
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
    Dim aFile As String
    Dim ii As Integer
    LnCnt = 1
             
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
       If frmTypeStr = "PV" Or frmTypeStr = "PX" Then
            PartyName = GProcGetColumnValue("tblMastAccount", "AcCode", RsSA.Fields("VCrAcCode"), "N", "AcName", "S")
            PartyGstin = GProcGetColumnValue("tblMastAccount", "AcCode", RsSA.Fields("VCrAcCode"), "N", "AcGSTIN", "S")
            PartyAdd1 = GProcGetColumnValue("tblMastAccount", "AcCode", RsSA.Fields("VCrAcCode"), "N", "AcAdd1", "S")
            PartyAdd2 = GProcGetColumnValue("tblMastAccount", "AcCode", RsSA.Fields("VCrAcCode"), "N", "AcAdd2", "S")
            PartyCity = GProcGetColumnValue("tblMastAccount", "AcCode", RsSA.Fields("VCrAcCode"), "N", "AcCity", "S")
            PartyPin = GProcGetColumnValue("tblMastAccount", "AcCode", RsSA.Fields("VCrAcCode"), "N", "AcZip", "S")
       Else
            PartyName = GProcGetColumnValue("tblMastAccount", "AcCode", RsSA.Fields("VDrAcCode"), "N", "AcName", "S")
            PartyGstin = GProcGetColumnValue("tblMastAccount", "AcCode", RsSA.Fields("VDrAcCode"), "N", "AcGSTIN", "S")
            PartyAdd1 = GProcGetColumnValue("tblMastAccount", "AcCode", RsSA.Fields("VDrAcCode"), "N", "AcAdd1", "S")
            PartyAdd2 = GProcGetColumnValue("tblMastAccount", "AcCode", RsSA.Fields("VDrAcCode"), "N", "AcAdd2", "S")
            PartyCity = GProcGetColumnValue("tblMastAccount", "AcCode", RsSA.Fields("VDrAcCode"), "N", "AcCity", "S")
            PartyPin = GProcGetColumnValue("tblMastAccount", "AcCode", RsSA.Fields("VDrAcCode"), "N", "AcZip", "S")
       End If
       ItemHsn = GProcGetColumnValue("tblMastItem", "ItCode", RsSA.Fields(19), "S", "ItHsn2", "S")
        
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
       If frmTypeStr = "PV" Or frmTypeStr = "PX" Then
          Print #1, Space(15) + Chr(34) + "Typ" + Chr(34) + ":" + Chr(34) + "CRN" + Chr(34) + Chr(44)
       Else
          Print #1, Space(15) + Chr(34) + "Typ" + Chr(34) + ":" + Chr(34) + "DBN" + Chr(34) + Chr(44)
       End If
       Print #1, Space(15) + Chr(34) + "No" + Chr(34) + ":" + Chr(34) + CStr(RsSA.Fields(0)) + Chr(34) + Chr(44)
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
       Print #1, Chr(34) + "ShipDtls" + Chr(34) + ":" + "null" + Chr(44)
      
      
       Print #1, Chr(34) + "ValDtls" + Chr(34) + ":{"
       Print #1, Space(15) + Chr(34) + "AssVal" + Chr(34) + ":" + Trim(CStr(Format(RsSA.Fields(9), "############0.00"))) + Chr(44)
       Print #1, Space(15) + Chr(34) + "IgstVal" + Chr(34) + ":" + Trim(CStr(Format(RsSA.Fields(18), "############0.00"))) + Chr(44)
       Print #1, Space(15) + Chr(34) + "CgstVal" + Chr(34) + ":" + Trim(CStr(Format(RsSA.Fields(12), "############0.00"))) + Chr(44)
       Print #1, Space(15) + Chr(34) + "SgstVal" + Chr(34) + ":" + Trim(CStr(Format(RsSA.Fields(14), "############0.00"))) + Chr(44)
       Print #1, Space(15) + Chr(34) + "CesVal" + Chr(34) + ":" + Trim(CStr(Format(0, "############0.00"))) + Chr(44)
       Print #1, Space(15) + Chr(34) + "StCesVal" + Chr(34) + ":" + Trim(CStr(Format(0, "############0.00"))) + Chr(44)
       Print #1, Space(15) + Chr(34) + "Discount" + Chr(34) + ":" + Trim(CStr(Format(0, "############0.00"))) + Chr(44)
       Print #1, Space(15) + Chr(34) + "OthChrg" + Chr(34) + ":" + Trim(CStr(Format(0, "############0.00"))) + Chr(44)
       Print #1, Space(15) + Chr(34) + "RndOffAmt" + Chr(34) + ":" + Trim(CStr(Format(RsSA.Fields(15), "############0.00"))) + Chr(44)
       Print #1, Space(15) + Chr(34) + "TotInvVal" + Chr(34) + ":" + Trim(CStr(Format(RsSA.Fields(10), "############0.00")))
       Print #1, "}" + Chr(44)
       
      
       Print #1, Space(25) + Chr(34) + "ItemList" + Chr(34) + ":[{"
       Print #1, Space(32) + Chr(34) + "SlNo" + Chr(34) + ":" + Chr(34) + "1" + Chr(34) + Chr(44)
       Print #1, Space(32) + Chr(34) + "PrdDesc" + Chr(34) + ":" + Chr(34) + "Yarn" + Chr(34) + Chr(44)
       Print #1, Space(32) + Chr(34) + "IsServc" + Chr(34) + ":" + Chr(34) + "N" + Chr(34) + Chr(44)
       Print #1, Space(32) + Chr(34) + "HsnCd" + Chr(34) + ":" + Chr(34) + Trim(ItemHsn) + Chr(34) + Chr(44)
       Print #1, Space(32) + Chr(34) + "Qty" + Chr(34) + ":" + Trim(CStr(Format(0, "############0.000"))) + Chr(44)
       Print #1, Space(32) + Chr(34) + "Unit" + Chr(34) + ":" + Chr(34) + "OTH" + Chr(34) + Chr(44)
       Print #1, Space(32) + Chr(34) + "UnitPrice" + Chr(34) + ":" + Trim(CStr(Format(0, "############0.000"))) + Chr(44)
       Print #1, Space(32) + Chr(34) + "TotAmt" + Chr(34) + ":" + Trim(CStr(Format(RsSA.Fields(9), "############0.00"))) + Chr(44)
       Print #1, Space(32) + Chr(34) + "Discount" + Chr(34) + ":" + Trim(CStr(Format(0, "############0.00"))) + Chr(44)
       Print #1, Space(32) + Chr(34) + "PreTaxVal" + Chr(34) + ":" + Trim(CStr(Format(0, "############0.00"))) + Chr(44)
       Print #1, Space(32) + Chr(34) + "AssAmt" + Chr(34) + ":" + Trim(CStr(Format(RsSA.Fields(9), "############0.00"))) + Chr(44)
       Print #1, Space(32) + Chr(34) + "GstRt" + Chr(34) + ":" + Trim(CStr(Format(RsSA.Fields(11) + RsSA.Fields(13) + RsSA.Fields(17), "############0"))) + Chr(44)
       Print #1, Space(32) + Chr(34) + "IgstAmt" + Chr(34) + ":" + Trim(CStr(Format(RsSA.Fields(18), "############0.00"))) + Chr(44)
       Print #1, Space(32) + Chr(34) + "CgstAmt" + Chr(34) + ":" + Trim(CStr(Format(RsSA.Fields(12), "############0.00"))) + Chr(44)
       Print #1, Space(32) + Chr(34) + "SgstAmt" + Chr(34) + ":" + Trim(CStr(Format(RsSA.Fields(14), "############0.00"))) + Chr(44)
       Print #1, Space(32) + Chr(34) + "CesRt" + Chr(34) + ":" + Trim(CStr(Format(0, "############0.00"))) + Chr(44)
       Print #1, Space(32) + Chr(34) + "CesAmt" + Chr(34) + ":" + Trim(CStr(Format(0, "############0.00"))) + Chr(44)
       Print #1, Space(32) + Chr(34) + "CesNonAdvlAmt" + Chr(34) + ":" + Trim(CStr(Format(0, "############0.00"))) + Chr(44)
       Print #1, Space(32) + Chr(34) + "StateCesRt" + Chr(34) + ":" + Trim(CStr(Format(0, "############0.00"))) + Chr(44)
       Print #1, Space(32) + Chr(34) + "StateCesAmt" + Chr(34) + ":" + Trim(CStr(Format(0, "############0.00"))) + Chr(44)
       Print #1, Space(32) + Chr(34) + "StateCesNonAdvlAmt" + Chr(34) + ":" + Trim(CStr(Format(0, "############0.00"))) + Chr(44)
       Print #1, Space(32) + Chr(34) + "OthChrg" + Chr(34) + ":" + Trim(CStr(Format(0, "############0.00"))) + Chr(44)
       Print #1, Space(32) + Chr(34) + "TotItemVal" + Chr(34) + ":" + Trim(CStr(Format(RsSA.Fields(9) + RsSA.Fields(18) + RsSA.Fields(12) + RsSA.Fields(14), "############0.00")))
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
          gCn.Execute "Update tblIntSale set SlIRNNo = '" & sheet.cells(a, 2) & "',SlAckNo = '" & sheet.cells(a, 3) & "'  where vtype='" & frmTypeStr & "' and Vno= '" & sheet.cells(a, 5) & "' and " _
          & " vfirm = '" & gCCode & "' and VYear=" & gCYear
       Else
          gCn.Execute "Update tblIntSale set SlIRNNo = '" & sheet.cells(a, 2) & "',SlAckNo = '" & sheet.cells(a, 3) & "'  where VDrAcCode= " & PtyCd & " and vtype='" & frmTypeStr & "' and VNo= '" & sheet.cells(a, 5) & "' and " _
          & " vfirm = '" & gCCode & "' and VYear=" & gCYear
       End If
       a = a + 1
    Loop
    workbook.Close SaveChanges:=True
    excel_app.Quit
    rstbl.Requery
    MsgBox ("e-Invoice IRN No Updated....")
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
    Dim X As String
    Me.Top = 300
    Me.Left = Screen.Width / 2 - Me.Width / 2
    '--- TblAddLess
    frmTypeStr = gfrmTypeStr
    mLastEntryDate = vbNull
    SalesAccode = 0
    X = "Select CCode from tblMastCompany where CDepotMainFirmCompCd = '" & gCCode & "'"
    i = GProcRstOpen(rstblBlChk, X, "R")
    BillsInFirm = "'" & gCCode & "'"
    If i > 0 Then
       rstblBlChk.MoveFirst
       Do While Not rstblBlChk.EOF
          BillsInFirm = BillsInFirm + "," + "'" & rstblBlChk.Fields(0) & "'"
          rstblBlChk.MoveNext
       Loop
    End If
'    txttemp(6) = GProcGetColumnValue("tblMastAccount", "AcCode", CStr(SalesAccode), "N", "AcName", "S")
    FirstTimeFlag = True
'    Criteria = "Vtype='" & frmTypeStr & "' and VFirm  = '" & gCCode & "' and VYear=" & gCYear & " and vdt >= CONVERT(datetime, '01.07.2017', 104) "
    Criteria = "Vtype='" & frmTypeStr & "' and VFirm = '" & gCCode & "'  and VYear=" & gCYear & "   and vnolist = 0  "
    If Trim(mShowRec) <> "" Then
       GProcRstOpen rstbl, "Select * from " & tblName & " where  " & Criteria & " and  " & IdField & " = " & CLng(Trim(mShowRec)), "R"
       ReadFields
       Call GProcSetButtons(Me, vbDataActionCancel)
       mShowRec = ""
    Else
       GProcRstOpen rstbl, "Select * from " & tblName & " where " & Criteria & " and  " & IdField & " = 0", "R"
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
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
If FormAction = vbDataActionUpdate Then
    MsgBox "You are about to modify the record. Save the record and then exit.", vbInformation + vbOKOnly, Me.Caption
    Cancel = True
    Exit Sub
End If
rstbl.Close
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
    If dtpDt(2) < CDate("01/07/2017") Then
       MsgBox "Date in VAT Applicabel " & txt(9), vbCritical + vbOKOnly, ""
       dtpDt(Index).SetFocus
       Exit Sub
    End If
End If
End Sub
Private Sub txt_GotFocus(Index As Integer)
GProcSelectBox Me.ActiveControl
End Sub
Private Sub CheckForNumber()
GProcCheckForNumber Me, rstbl, MaxNo, Array(33, 25, 15), pOtherTxtFromNo:=OtherTxtFromNo, pOtherTxtToNo:=OtherTxtToNo
End Sub
Private Sub txt_LostFocus(Index As Integer)
'If Index = 33 And FormAction <> vbDataActionUpdate Then
'   Dim X As String
'   Dim rstAddCheck As Recordset
'   X = "Select CrDrNoteNo From " & tblName & " Where CrDrNoteNo = '" & txt(33) & "'  and Vtype='" & frmTypeStr & "'  and Vfirm = '" & gCCode & "' and  VYear=" & gCYear
'        i = GProcRstOpen(rstAddCheck, X, "R")
'        If i > 0 Then
'            MsgBox "Record is Available of Bill No. " & txt(33), vbCritical + vbOKOnly, "Duplicate Record"
'            txtCode(7).SetFocus
'        End If
'End If
CheckForNumber
SaveBtnEd
CalAmount
End Sub
Private Sub txt_KeyPress(Index As Integer, KeyAscii As Integer)
If KeyAscii = 13 Then Exit Sub
KeyAscii = GProcValidateKey(rstbl, Index, KeyAscii, txt(Index).text)
If Index = 15 And txttemp(7) <> "" Then
   Dim X As String
   Dim BillsInFirm  As String
   Dim PtyCd As Long
   PtyCd = GProcGetColumnValue("tblMastAccount", "AcName", txttemp(7), "S", "AcCode", "N")
    X = "Select CCode from tblMastCompany where CDepotMainFirmCompCd = '" & gCCode & "'"
    i = GProcRstOpen(rstblBlChk, X, "R")
    BillsInFirm = "'" & gCCode & "'"
    If i > 0 Then
       rstblBlChk.MoveFirst
       Do While Not rstblBlChk.EOF
          BillsInFirm = BillsInFirm + "," + "'" & rstblBlChk.Fields(0) & "'"
          rstblBlChk.MoveNext
       Loop
    End If
    
    If frmTypeStr = "SV" Or frmTypeStr = "PV" Then
        X = "select PurBillNo,format(tblPurch.PurBillDt,'dd/MM/yyyy'),PurBillAmt,tblMastAccount.AcOurCode,tblMastAccount.AcName,tblPurch.VType," _
            & " tblPurch.VYear,tblPurch.Vno,tblPurch.vfirm,tblPurch.PurTaxRate,tblPurch.PurTaxRate2,tblPurch.PurTaxRate3,tblPurchSub.PurSubItCode from tblPurch,tblPurchSub, tblMastAccount where  tblPurch.PurAcCrCode  = " & PtyCd & " and    tblPurch.PurBroker = tblMastAccount.AcCode and " _
            & "  tblPurch.Vno = tblPurchSub.Vno and  tblPurch.VType = tblPurchSub.Vtype and tblPurch.Vyear = tblPurchSub.Vyear and tblPurch.Vfirm = tblPurchSub.Vfirm and tblPurchSub.PurSubItSrNo = 1 and tblPurch.vfirm in (" & BillsInFirm & ") "
        gClsSearch.SearchMultiField X, "", Array("Bill No", "Date", "Bill Amt", "Broker Code", "Broker Name", "Vtype", "VYear", "INV Vno", "BillFirm", "C Rt", "S Rt", "I Rt", "Itcode"), Array(1500, 1500, 2000, 2000, 3000, 0, 0, 0, 0, 0, 0, 0, 0), "", "", "tblPurch.PurBillDt,tblPurch.PurBillNo", 0, 0, False, , True
    Else
        X = "select SlBillNo,format(tblSale.Vdt,'dd/MM/yyyy'),SlBillAmt,tblMastAccount.AcOurCode,tblMastAccount.AcName,tblSale.VType," _
            & " tblSale.VYear,tblSale.Vno,tblSale.vfirm,tblSale.SlTaxRate,tblSale.SlTaxRate2,tblSale.SlTaxRate3,tblSaleSub.SlSubItCode from tblSale,tblSaleSub, tblMastAccount where  tblSale.SlAcDrCode  = " & PtyCd & " and    tblSale.SlBroker = tblMastAccount.AcCode and " _
            & "  tblSale.Vno = tblSaleSub.Vno and  tblSale.VType = tblSaleSub.Vtype and tblSale.Vyear = tblSaleSub.Vyear and tblSale.Vfirm = tblSaleSub.Vfirm and tblSalesub.SlSubItSrNo = 1 and    tblSale.vfirm in (" & BillsInFirm & ") "
        gClsSearch.SearchMultiField X, "", Array("Bill No", "Date", "Bill Amt", "Broker Code", "Broker Name", "Vtype", "VYear", "INV Vno", "BillFirm", "C Rt", "S Rt", "I Rt", "Itcode"), Array(1500, 1500, 2000, 2000, 3000, 0, 0, 0, 0, 0, 0, 0, 0), "", "", "tblSale.Vdt,tblSale.SlBillNo", 0, 0, False, , True
    End If
     If gClsSearch.SearchMultiRetCol(0) <> "" Then
         txt(15) = gClsSearch.SearchMultiRetCol(0)
         dtpDt(18) = Format(gClsSearch.SearchMultiRetCol(1), "dd/MM/yyyy")
        txtCode(11) = gClsSearch.SearchMultiRetCol(3)
        txttemp(11) = gClsSearch.SearchMultiRetCol(4)
        txt(15) = gClsSearch.SearchMultiRetCol(0)
        txt(16) = gClsSearch.SearchMultiRetCol(7)
        txt(17) = gClsSearch.SearchMultiRetCol(5)
        txt(19) = gClsSearch.SearchMultiRetCol(6)
        txt(20) = gClsSearch.SearchMultiRetCol(2)
        txt(24) = gClsSearch.SearchMultiRetCol(8)
        txt(27) = gClsSearch.SearchMultiRetCol(9)
        txt(29) = gClsSearch.SearchMultiRetCol(10)
        txt(31) = gClsSearch.SearchMultiRetCol(11)
        txt(34) = gClsSearch.SearchMultiRetCol(12)
        txtCode(6).SetFocus
     End If
End If
End Sub

Private Sub txtCode_GotFocus(Index As Integer)
'If Index = 7 And txtCode(7) = "" Then
'   txtCode(7) = "D"
'   txtCode(7).SelStart = 2
'End If
End Sub
Private Sub txtCode_LostFocus(Index As Integer)
If Trim(txtCode(Index)) = "" Then Exit Sub
Dim BrkAcCode As Long
Dim tmpBillSr As String
Dim NarrCdBlTp As Long
Dim BilSr As String
SetControlEd
'If cmdBtn_Click(Index) <> 14 Then
Select Case Index
    Case 7  '-- Party  Code
         txttemp(7) = GProcGetColumnValue("tblMastAccount", "AcOurCode", UCase(txtCode(7)), "S", "AcName", "S")
         lblTinNo.Caption = GProcGetColumnValue("TblMastAccount", "ACName", txttemp(Index), "S", "AcGSTIN", "S")
         lblAdd1.Caption = GProcGetColumnValue("TblMastAccount", "ACName", txttemp(Index), "S", "AcAdd1", "S")
         lblAdd2.Caption = GProcGetColumnValue("TblMastAccount", "ACName", txttemp(Index), "S", "AcAdd2", "S")
         If Trim(txttemp(7)) = "" Then
            If txtCode(7) <> "D" Then
               MsgBox ("Code Not Found ...")
            End If
            txttemp(7) = ""
            txttemp(7).Enabled = True
            txttemp(7).SetFocus
         Else
            txtCode(11).SetFocus
         End If
    Case 6  '-- Sales   Code
         txttemp(6) = GProcGetColumnValue("tblMastAccount", "AcOurCode", UCase(txtCode(6)), "S", "AcName", "S")
         If Trim(txttemp(6)) = "" Then
            txttemp(6) = ""
            txttemp(6).Enabled = True
            txttemp(6).SetFocus
         End If
    Case 11  '-- Broker  Code
         txttemp(11) = GProcGetColumnValue("tblMastAccount", "AcOurCode", UCase(txtCode(11)), "S", "AcName", "S")
         If Trim(txttemp(11)) = "" Then
            MsgBox ("Code Not Found ...")
            txttemp(11) = ""
            txttemp(11).Enabled = True
            txttemp(11).SetFocus
'         Else
'            If txt(9).Enabled = True Then txt(9).SetFocus
         End If
End Select
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
        Case 7 '-- Party
            GName = GProcGetColumnValue("TblMastGroup", "AgCode", 90017, "N", "AgName", "S")
            GProcShowForm frmMastAccount, frmMain.mnuMstfrm(0), Array(frmMastAccount.txttemp(2)), Array(GName)
        Case 6 '-- Sales A/c
            GName = GProcGetColumnValue("TblMastGroup", "AgCode", 90023, "N", "AgName", "S")
            GProcShowForm frmMastAccount, frmMain.mnuMstfrm(0), Array(frmMastAccount.txttemp(2)), Array(GName)
        Case 11 '-- Broker
           GName = GProcGetColumnValue("TblMastGroup", "AgCode", 90016, "N", "AgName", "S")
           GProcShowForm frmMastAccount, frmMain.mnuMstfrm(0), Array(frmMastAccount.txttemp(2)), Array(GName)
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
TopPos = txttemp(Index).Top + Me.Top + 650
LeftPos = txttemp(Index).Left + Me.Left
If KeyAscii = 13 Then Exit Sub
Select Case Index
    Case 7  '-- Party   Alias Account
        gClsSearch.SearchMultiFieldAccMaster Chr(KeyAscii), txttemp(Index).Left + Me.Left, txttemp(Index).Top + Me.Top + 1000
        KeyAscii = 0
        lblTinNo.Caption = ""
        lblMsOms.Caption = ""
        txtCode(7) = ""
        If gClsSearch.SearchMultiRetColAccMaster(0) <> "" Then
            txttemp(Index) = gClsSearch.SearchMultiRetColAccMaster(0)
            txtCode(7) = GProcGetColumnValue("TblMastAccount", "ACName", txttemp(Index), "S", "AcOurCode", "S")
            lblTinNo.Caption = GProcGetColumnValue("TblMastAccount", "ACName", txttemp(Index), "S", "AcGSTIN", "S")
            lblMsOms.Caption = GProcGetColumnValue("TblMastAccount", "ACName", txttemp(Index), "S", "AcMsOmsParty", "N")
            If Val(lblMsOms.Caption) = 0 Then
               lblMsOms.Caption = "M.S."
            Else
               lblMsOms.Caption = "O.M.S."
            End If
            lblAdd1.Caption = GProcGetColumnValue("TblMastAccount", "ACName", txttemp(Index), "S", "AcAdd1", "S")
            lblAdd2.Caption = GProcGetColumnValue("TblMastAccount", "ACName", txttemp(Index), "S", "AcAdd2", "S")
        End If
    Case 6 '-- Sales A/c  Alias Account
        gClsSearch.SearchMultiFieldAccMaster Chr(KeyAscii), LeftPos, TopPos
        KeyAscii = 0
        If gClsSearch.SearchMultiRetColAccMaster(0) <> "" Then
            txttemp(Index) = gClsSearch.SearchMultiRetColAccMaster(0)
            txtCode(6) = GProcGetColumnValue("tblMastAccount", "AcName", txttemp(6), "S", "AcOurCode", "S")
        End If
'        gClsSearch.SearchMultiField "tblMastaccount", "AcName,AcCity", Array("Sales Account", "City"), Array(3000, 2000), "(Agcode=90023)", Chr(KeyAscii), "", LeftPos, TopPos, True
'        KeyAscii = 0
'        If gClsSearch.SearchMultiRetCol(0) <> "" Then
'            txttemp(Index) = gClsSearch.SearchMultiRetCol(0)
'        End If
    Case 11 '-- Broker
        gClsSearch.SearchMultiField "tblMastaccount", "AcName,AcOurCode,AcCity", Array("Broker", "Code", "City"), Array(5000, 1000, 3000), "(AgCode=90016 )", Trim(Chr(KeyAscii)), "", LeftPos, TopPos, True
        KeyAscii = 0
        txtCode(11) = ""
        If gClsSearch.SearchMultiRetCol(0) <> "" Then
            txttemp(Index) = gClsSearch.SearchMultiRetCol(0)
            txtCode(11) = gClsSearch.SearchMultiRetCol(1)
        End If
End Select
End Sub
Private Sub txtTemp_LostFocus(Index As Integer)
Dim AcCode As Long
Dim BilSr As String
Dim NarrCdBlTp As Long
SaveBtnEd
End Sub
Private Sub SaveBtnEd()
'--- Vno,vdt,Party,A/c,Mill,Amount
If (CDbl(txt(0)) = 0 And FormAction <> vbDataActionAddNew) Or GProcIsDateValid(dtpDt(2)) = False Or txttemp(7) = "" Or txttemp(6) = "" Or CDbl(txt(8)) = 0 Then
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
    
If Len(Trim(lblTinNo.Caption)) = 0 Then
      MsgBox "Party GSTIN Not Found", vbInformation + vbOKOnly, Me.Caption
      txttemp(7).SetFocus
      Exit Function
End If
    
    
    '---Validate Vou.Date = Bill Date
'    If mskdt(10) <> dtpDt(2) Then
'        MsgBox "Vou.Date Muse Be Same As Bill Date", vbInformation + vbOKOnly, Me.Caption
'        mskDt(12).SetFocus
'        Exit Function
'    End If
    '--- Validate Grid
    '--- Vno
    If CLng(txt(0)) = 0 And FormAction <> vbDataActionAddNew Then
        MsgBox "Check Voucher No.", vbInformation + vbOKOnly, Me.Caption
        dtpDt(2).SetFocus
        Exit Function
    End If
    If CLng(txt(34)) = 0 And FormAction <> vbDataActionAddNew Then
        MsgBox "ItemCode For HSN Not found..", vbInformation + vbOKOnly, Me.Caption
        dtpDt(2).SetFocus
        Exit Function
    End If
    
    
    '--- Vdt
    If GProcIsDateValid(dtpDt(2)) = False Then
        MsgBox "Check Voucher Date.", vbInformation + vbOKOnly, Me.Caption
        dtpDt(2).SetFocus
        Exit Function
    End If
    
    '---- Party Account
    If txttemp(7) = "" Then
        MsgBox "Check Party.", vbInformation + vbOKOnly, Me.Caption
        txttemp(7).SetFocus
        Exit Function
    End If
        '--Create Party
        If GProcGetColumnValue("tblMastAccount", "AcName", txttemp(7), "S", "AcCode", "N") = 0 Then
           GName = GProcGetColumnValue("TblMastGroup", "AgCode", 90017, "N", "AgName", "S")
           GProcShowForm frmMastAccount, frmMain.mnuMstfrm(0), Array(frmMastAccount.txt(1), frmMastAccount.txttemp(2)), Array(txttemp(7), GName)
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
       '--Create Broker
        If txttemp(11) = "" Then txttemp(11) = "-"
        If GProcGetColumnValue("tblMastAccount", "AcName", txttemp(11), "S", "AcCode", "N") = 0 Then
           GName = GProcGetColumnValue("TblMastGroup", "AgCode", 90016, "N", "AgName", "S")
           GProcShowForm frmMastAccount, frmMain.mnuMstfrm(0), Array(frmMastAccount.txt(1), frmMastAccount.txttemp(2)), Array(txttemp(11), GName)
           Exit Function
        End If
    '---- Amount
    If CDbl(txt(8)) <= 0 Then
        MsgBox "Check Voucher Amount.", vbInformation + vbOKOnly, Me.Caption
        Exit Function
    End If
    '----Duplicate Records
    '----Addition Case
'    If FormAction = vbDataActionAddNew Then
'        X = "Select vno From " & tblName & " Where vno=" & CLng(txt(0)) & " and Vtype='" & frmTypeStr & "'  and Vfirm = '" & gCCode & "'  and VYear=" & gCYear
'        i = GProcRstOpen(rstAddCheck, X, "R")
'        If i > 0 Then
'            MsgBox "Record is Available of Voucher No. " & CLng(txt(0)), vbCritical + vbOKOnly, "Duplicate Record"
'            txt(9).SetFocus
'            Exit Function
'        End If
'        rstAddCheck.Close
'    End If
    ValidateData = True
End Function
Private Sub FillTxtFromTemp()
txt(1) = frmTypeStr '--VType
If IsDate(dtpDt(2)) Then '--Vdt
    txt(2) = CDate(dtpDt(2))
Else
    txt(2) = ""
End If
If IsDate(dtpDt(18)) Then '--Vdt
    txt(18) = CDate(dtpDt(18))
Else
    txt(18) = ""
End If
If IsDate(dtpDt(12)) Then '--Ref.Vdt
    txt(12) = CDate(dtpDt(12))
Else
    txt(12) = ""
End If

If gfrmTypeStr = "PV" Or gfrmTypeStr = "PX" Then
    '-- Dr Code
    txt(7) = GProcGetColumnValue("tblMastAccount", "AcName", txttemp(6), "S", "AcCode", "N")
    '-- Cr Code
    txt(6) = GProcGetColumnValue("tblMastAccount", "AcName", txttemp(7), "S", "AcCode", "N")
End If
If gfrmTypeStr = "SV" Or gfrmTypeStr = "SI" Then
    '-- Dr Code
    txt(7) = GProcGetColumnValue("tblMastAccount", "AcName", txttemp(7), "S", "AcCode", "N")
    '-- Cr Code
    txt(6) = GProcGetColumnValue("tblMastAccount", "AcName", txttemp(6), "S", "AcCode", "N")
End If



'-- broker
If txttemp(11) = "" Then
   txttemp(11) = "-"
End If
txt(11) = GProcGetColumnValue("tblMastAccount", "AcName", txttemp(11), "S", "AcCode", "N")
txt(3) = gCYear
txt(4) = gCCode
End Sub
Private Sub FillTempFromTxt()
Dim X As String
Dim MsOms As Integer
If IsDate(txt(2)) Then '--VDt
    dtpDt(2) = txt(2)
End If
If IsDate(txt(18)) Then '--Bill Date
    dtpDt(18) = txt(18)
End If
If IsDate(txt(12)) Then '--Bill Date
    dtpDt(12) = txt(12)
End If

If gfrmTypeStr = "PV" Or gfrmTypeStr = "PX" Then
    '-- Dr Code
    txttemp(7) = GProcGetColumnValue("tblMastAccount", "AcCode", txt(6), "N", "AcName", "S")
    txtCode(7) = GProcGetColumnValue("tblMastAccount", "AcCode", txt(6), "N", "AcOurCode", "S")
    lblTinNo.Caption = GProcGetColumnValue("TblMastAccount", "ACName", txttemp(6), "S", "AcGSTIN", "S")
    MsOms = GProcGetColumnValue("tblMastAccount", "AcCode", txt(6), "N", "AcMsOmsParty", "N")
    If MsOms = 0 Then
       lblMsOms.Caption = "M.S."
    Else
       lblMsOms.Caption = "O.M.S."
    End If
    lblAdd1.Caption = GProcGetColumnValue("TblMastAccount", "AcCode", txt(6), "N", "AcAdd1", "S")
    lblAdd2.Caption = GProcGetColumnValue("TblMastAccount", "AcCode", txt(6), "N", "AcAdd2", "S")
    If gCYear >= 2021 Then
       lblHsn.Caption = GProcGetColumnValue("TblMastItem", "ItCode", txt(34), "N", "ItHsn2", "S")
    Else
       lblHsn.Caption = GProcGetColumnValue("TblMastItem", "ItCode", txt(34), "N", "ItHsn", "S")
    End If
    '-- Cr Code
    txttemp(6) = GProcGetColumnValue("tblMastAccount", "AcCode", txt(7), "N", "AcName", "S")
    txtCode(6) = GProcGetColumnValue("tblMastAccount", "AcCode", txt(7), "N", "AcOurCode", "S")
End If
If gfrmTypeStr = "SV" Or gfrmTypeStr = "SI" Then
    '-- Dr Code
    txttemp(7) = GProcGetColumnValue("tblMastAccount", "AcCode", txt(7), "N", "AcName", "S")
    txtCode(7) = GProcGetColumnValue("tblMastAccount", "AcCode", txt(7), "N", "AcOurCode", "S")
    lblTinNo.Caption = GProcGetColumnValue("TblMastAccount", "ACName", txttemp(7), "S", "AcGSTIN", "S")
    MsOms = GProcGetColumnValue("tblMastAccount", "AcCode", txt(7), "N", "AcMsOmsParty", "N")
    If MsOms = 0 Then
       lblMsOms.Caption = "M.S."
    Else
       lblMsOms.Caption = "O.M.S."
    End If
    lblAdd1.Caption = GProcGetColumnValue("TblMastAccount", "AcCode", txt(7), "N", "AcAdd1", "S")
    lblAdd2.Caption = GProcGetColumnValue("TblMastAccount", "AcCode", txt(7), "N", "AcAdd2", "S")
    If gCYear >= 2021 Then
       lblHsn.Caption = GProcGetColumnValue("TblMastItem", "ItCode", txt(34), "N", "ItHsn2", "S")
    Else
       lblHsn.Caption = GProcGetColumnValue("TblMastItem", "ItCode", txt(34), "N", "ItHsn", "S")
    End If
    '-- Cr Code
    txttemp(6) = GProcGetColumnValue("tblMastAccount", "AcCode", txt(6), "N", "AcName", "S")
    txtCode(6) = GProcGetColumnValue("tblMastAccount", "AcCode", txt(6), "N", "AcOurCode", "S")
End If




'--Broker
txttemp(11) = GProcGetColumnValue("tblMastAccount", "AcCode", txt(11), "N", "AcName", "S")
txtCode(11) = GProcGetColumnValue("tblMastAccount", "AcCode", txt(11), "N", "AcOurCode", "S")

'-- Fill Related Record
'FillRelatedRecord
SetControlEd
End Sub
Private Sub ReadFields()
FormAction = vbDataActionClose
GProcShowRecord Me, rstbl, MaxNo, OtherTxtFromNo, OtherTxtToNo
FillTempFromTxt
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

'X = "Select Vno,Vdt from tblVoucher where VBillType= '" & frmTypeStr & "' and " _
'& "  VBillVno = " & txt(0) & " and VFirm  = '" & gCCode & "' and  VBillVYear= " & gCYear & ""
'GProcRstOpen RsG, X, "R"
'If RsG.EOF Then
'Else
'    IsRelatedRecord = True
'    MsgBox "You can not Modify/Delete this record.Related Record No. " & RsG.Fields(0) & " is exists.", vbCritical + vbOKOnly, App.Title
'    Exit Function
'End If
'RsG.Close
End Function

Private Sub CalAmount()
Dim TaxAmt As Double
Dim TaxAmtCGST As Double
Dim TaxAmtIGST As Double
Dim SubAmt As Double
Dim Amt As Double
Dim Amt2 As Double
Dim ROff As Double
'-- Due Date
CheckForNumber
'--- Subamt+Add -Less
SubAmt = CDbl(txt(21))
If Abs((CDbl(txt(28)) - (SubAmt * CDbl(txt(27)) / 100))) > 1 Then
    TaxAmt = SubAmt * CDbl(txt(27)) / 100
Else
    TaxAmt = CDbl(txt(28))
End If
TaxAmt = Format(TaxAmt, FStr)

'- 2 CGST
If Abs((CDbl(txt(30)) - (SubAmt * CDbl(txt(29)) / 100))) > 1 Then
   TaxAmtCGST = SubAmt * CDbl(txt(29)) / 100
Else
   TaxAmtCGST = CDbl(txt(30))
End If
TaxAmtCGST = Format(TaxAmtCGST, FStr)
  
'- 3 IGST
If Abs((CDbl(txt(32)) - (SubAmt * CDbl(txt(31)) / 100))) > 1 Then
   TaxAmtIGST = SubAmt * CDbl(txt(31)) / 100
Else
   TaxAmtIGST = CDbl(txt(32))
End If
TaxAmtIGST = Format(TaxAmtIGST, FStr)
Amt = Format(SubAmt + TaxAmt + TaxAmtCGST + TaxAmtIGST, FStr)
Amt2 = GProcMakeRounding(Amt)
txt(21) = SubAmt
txt(28) = TaxAmt
txt(30) = TaxAmtCGST
txt(32) = TaxAmtIGST
CheckForNumber
txt(8) = Amt2
End Sub















