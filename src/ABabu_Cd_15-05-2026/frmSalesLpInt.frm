VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form frmSalesLpInt 
   BackColor       =   &H00FBF2E1&
   BorderStyle     =   0  'None
   Caption         =   "Purchase Entry"
   ClientHeight    =   8805
   ClientLeft      =   0
   ClientTop       =   15
   ClientWidth     =   14790
   FillColor       =   &H000040C0&
   FillStyle       =   0  'Solid
   Icon            =   "frmSalesLpInt.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   8805
   ScaleWidth      =   14790
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame2 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFC0C0&
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
      ForeColor       =   &H80000008&
      Height          =   3195
      Left            =   3780
      TabIndex        =   138
      Top             =   1260
      Visible         =   0   'False
      Width           =   7605
      Begin VB.TextBox txtIrnNo 
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
         Height          =   825
         Left            =   1440
         MaxLength       =   70
         MultiLine       =   -1  'True
         TabIndex        =   142
         Top             =   465
         Width           =   5655
      End
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
         Index           =   0
         Left            =   1260
         TabIndex        =   141
         Top             =   2550
         Width           =   1635
      End
      Begin VB.TextBox txtAckNo 
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
         Left            =   1440
         MaxLength       =   25
         TabIndex        =   140
         Top             =   1695
         Width           =   5655
      End
      Begin VB.CommandButton cmdUpdate 
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
         Index           =   1
         Left            =   3060
         TabIndex        =   139
         Top             =   2580
         Width           =   1635
      End
      Begin VB.Line Line1 
         X1              =   0
         X2              =   7560
         Y1              =   2355
         Y2              =   2355
      End
      Begin VB.Label Label20 
         AutoSize        =   -1  'True
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
         Height          =   240
         Left            =   300
         TabIndex        =   144
         Top             =   480
         Width           =   885
      End
      Begin VB.Label Label19 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Ack No   :"
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
         Left            =   255
         TabIndex        =   143
         Top             =   1725
         Width           =   900
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
      Height          =   390
      Index           =   31
      Left            =   8040
      TabIndex        =   16
      Text            =   "31"
      Top             =   7605
      Width           =   690
   End
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
      Height          =   4815
      Left            =   1440
      TabIndex        =   111
      Top             =   1095
      Visible         =   0   'False
      Width           =   7605
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
         Left            =   4860
         TabIndex        =   135
         Top             =   3990
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
         Left            =   3060
         TabIndex        =   134
         Top             =   3990
         Width           =   1635
      End
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
         Left            =   1260
         TabIndex        =   133
         Top             =   3990
         Width           =   1635
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
         Left            =   4440
         TabIndex        =   125
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
         Left            =   4440
         TabIndex        =   124
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
         Left            =   4440
         TabIndex        =   123
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
         Left            =   4440
         TabIndex        =   122
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
         Left            =   4440
         TabIndex        =   121
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
         Left            =   4860
         TabIndex        =   116
         Top             =   3210
         Width           =   1635
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
         Left            =   3060
         TabIndex        =   115
         Top             =   3210
         Width           =   1635
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
         Left            =   2160
         MaxLength       =   10
         TabIndex        =   114
         Top             =   2175
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
         Left            =   1260
         TabIndex        =   113
         Top             =   3210
         Width           =   1635
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
         Left            =   2160
         MaxLength       =   10
         TabIndex        =   112
         Top             =   1605
         Width           =   2055
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
         Left            =   255
         TabIndex        =   118
         Top             =   2265
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
         Left            =   255
         TabIndex        =   117
         Top             =   1695
         Width           =   1605
      End
      Begin VB.Line Line3 
         X1              =   0
         X2              =   7560
         Y1              =   2955
         Y2              =   2955
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
      Height          =   375
      Index           =   47
      Left            =   15240
      TabIndex        =   108
      Text            =   "47"
      Top             =   6975
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
      Left            =   15240
      TabIndex        =   107
      Text            =   "46"
      Top             =   6570
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
      Left            =   15195
      TabIndex        =   106
      Text            =   "45"
      Top             =   6120
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
      Index           =   11
      Left            =   1575
      TabIndex        =   3
      Text            =   "txtCode(11)"
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
      Height          =   390
      Index           =   7
      Left            =   1575
      TabIndex        =   1
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
      Left            =   15150
      TabIndex        =   104
      Text            =   "44"
      Top             =   5580
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
      TabIndex        =   103
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
      Height          =   360
      Index           =   32
      Left            =   10440
      TabIndex        =   19
      Text            =   "32"
      Top             =   7605
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
      TabIndex        =   102
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
      Left            =   17280
      TabIndex        =   92
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
      Left            =   17235
      TabIndex        =   91
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
      TabIndex        =   24
      Text            =   "38"
      Top             =   8145
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
      Index           =   9
      Left            =   16425
      TabIndex        =   22
      Text            =   "9"
      Top             =   7785
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
      Height          =   375
      Index           =   37
      Left            =   1455
      TabIndex        =   90
      Text            =   "37"
      Top             =   8160
      Width           =   4590
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
      Left            =   1455
      MaxLength       =   64
      MultiLine       =   -1  'True
      TabIndex        =   89
      Text            =   "frmSalesLpInt.frx":058A
      Top             =   7080
      Width           =   4650
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
      TabIndex        =   87
      Top             =   7110
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
      TabIndex        =   86
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
      TabIndex        =   36
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
      TabIndex        =   80
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
      TabIndex        =   35
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
      TabIndex        =   29
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
      TabIndex        =   79
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
      TabIndex        =   27
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
      Height          =   390
      Index           =   28
      Left            =   10440
      TabIndex        =   17
      Text            =   "28"
      Top             =   6705
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
      Index           =   29
      Left            =   8040
      TabIndex        =   15
      Text            =   "29"
      Top             =   7140
      Width           =   690
   End
   Begin VB.CommandButton CmdBagDetails 
      Caption         =   "Cartoon &Bag Details"
      Height          =   720
      Left            =   17460
      TabIndex        =   75
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
      Height          =   390
      Index           =   8
      Left            =   10440
      TabIndex        =   20
      Text            =   "8"
      Top             =   8055
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
      TabIndex        =   74
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
      TabIndex        =   33
      Text            =   "19"
      Top             =   8955
      Visible         =   0   'False
      Width           =   285
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
      Index           =   33
      Left            =   10935
      TabIndex        =   5
      Text            =   "33"
      Top             =   630
      Width           =   1740
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
      Index           =   11
      Left            =   2775
      TabIndex        =   4
      Text            =   "txttemp(11)"
      ToolTipText     =   "Press F3 for Creation"
      Top             =   2385
      Width           =   6540
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
      Left            =   2775
      TabIndex        =   2
      Text            =   "txttemp(7)"
      ToolTipText     =   "Press F3 for Creation"
      Top             =   735
      Width           =   6540
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
      TabIndex        =   31
      Text            =   "txtDrBalance"
      Top             =   1170
      Width           =   300
   End
   Begin VB.CommandButton cmdLedger 
      Caption         =   "Ledger"
      Height          =   375
      Left            =   12915
      TabIndex        =   73
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
      Left            =   8040
      TabIndex        =   14
      Text            =   "27"
      Top             =   6690
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
      TabIndex        =   49
      Text            =   "26"
      Top             =   8190
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
      Index           =   15
      Left            =   10935
      TabIndex        =   9
      Text            =   "15"
      Top             =   2295
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
      Left            =   10440
      TabIndex        =   13
      Text            =   "21"
      Top             =   6240
      Width           =   2265
   End
   Begin VB.TextBox txtGrid 
      BackColor       =   &H80000000&
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
      Left            =   435
      TabIndex        =   12
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
      TabIndex        =   71
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
      TabIndex        =   69
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
      TabIndex        =   68
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
      TabIndex        =   67
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
      TabIndex        =   63
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
      TabIndex        =   62
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
      TabIndex        =   60
      ToolTipText     =   "Exit"
      Top             =   8010
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
      TabIndex        =   53
      ToolTipText     =   "To Save Record"
      Top             =   6645
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
      TabIndex        =   61
      Top             =   7560
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
      TabIndex        =   52
      ToolTipText     =   "To Delete Record"
      Top             =   6195
      Visible         =   0   'False
      Width           =   870
   End
   Begin VB.CommandButton cmdBtn 
      Caption         =   " IRN"
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
      TabIndex        =   51
      ToolTipText     =   "To Modify Record"
      Top             =   2580
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
      TabIndex        =   50
      ToolTipText     =   "To Add New Record"
      Top             =   2115
      Visible         =   0   'False
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
      TabIndex        =   54
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
      TabIndex        =   55
      ToolTipText     =   "To Find Record"
      Top             =   1665
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
      TabIndex        =   59
      ToolTipText     =   "Last"
      Top             =   1170
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
      TabIndex        =   58
      ToolTipText     =   "First"
      Top             =   1170
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
      TabIndex        =   57
      ToolTipText     =   "Next"
      Top             =   720
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
      Left            =   10935
      TabIndex        =   7
      Text            =   "0"
      Top             =   1440
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
      TabIndex        =   56
      ToolTipText     =   "Previous"
      Top             =   720
      Width           =   375
   End
   Begin VB.TextBox txt 
      Alignment       =   1  'Right Justify
      Height          =   315
      Index           =   4
      Left            =   18270
      TabIndex        =   66
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
      TabIndex        =   21
      Text            =   "txttemp(6)"
      ToolTipText     =   "Press F3 for Creation"
      Top             =   6165
      Width           =   4740
   End
   Begin VB.TextBox txt 
      Height          =   315
      Index           =   6
      Left            =   18405
      TabIndex        =   65
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
      TabIndex        =   64
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
      TabIndex        =   41
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
      TabIndex        =   47
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
      TabIndex        =   26
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
      Left            =   10935
      TabIndex        =   8
      Text            =   "23"
      Top             =   1890
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
      Height          =   360
      Index           =   24
      Left            =   16650
      TabIndex        =   34
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
      TabIndex        =   38
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
      Height          =   360
      Index           =   30
      Left            =   10440
      TabIndex        =   18
      Text            =   "30"
      Top             =   7155
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
      TabIndex        =   30
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
      Left            =   17145
      TabIndex        =   32
      Text            =   "18"
      Top             =   2295
      Visible         =   0   'False
      Width           =   330
   End
   Begin MSMask.MaskEdBox mskDt 
      Height          =   390
      Index           =   2
      Left            =   15165
      TabIndex        =   23
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
      TabIndex        =   25
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
      TabIndex        =   37
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
      Bindings        =   "frmSalesLpInt.frx":058D
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
      Height          =   2715
      Left            =   180
      TabIndex        =   11
      Top             =   3165
      Width           =   14235
      _ExtentX        =   25109
      _ExtentY        =   4789
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
   Begin MSMask.MaskEdBox mskDt 
      Height          =   375
      Index           =   14
      Left            =   16830
      TabIndex        =   28
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
      TabIndex        =   93
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
      Left            =   10935
      TabIndex        =   6
      Top             =   1035
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
      Format          =   95944707
      CurrentDate     =   -273
   End
   Begin MSComCtl2.DTPicker dtpDt 
      Height          =   390
      Index           =   18
      Left            =   10935
      TabIndex        =   10
      Top             =   2700
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
      Format          =   95944707
      CurrentDate     =   -273
   End
   Begin VB.Label Label15 
      BackStyle       =   0  'Transparent
      Caption         =   "Ack No.     :"
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
      Left            =   240
      TabIndex        =   137
      Top             =   8220
      Width           =   1260
   End
   Begin VB.Label Label12 
      BackStyle       =   0  'Transparent
      Caption         =   "IRN No       :"
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
      Left            =   240
      TabIndex        =   136
      Top             =   7140
      Width           =   1260
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
      Left            =   9750
      TabIndex        =   132
      Top             =   1530
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
      Left            =   1530
      TabIndex        =   131
      Top             =   6660
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
      Left            =   270
      TabIndex        =   130
      Top             =   6660
      Width           =   1260
   End
   Begin VB.Shape Shape4 
      FillColor       =   &H00FED8FC&
      FillStyle       =   0  'Solid
      Height          =   2550
      Left            =   13050
      Shape           =   4  'Rounded Rectangle
      Top             =   6030
      Width           =   1350
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
      Left            =   9750
      TabIndex        =   129
      Top             =   2745
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
      Left            =   9750
      TabIndex        =   128
      Top             =   2385
      Width           =   1095
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "List No    : "
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
      Left            =   9750
      TabIndex        =   127
      Top             =   1935
      Width           =   1095
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Bill. No    :"
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
      Left            =   9720
      TabIndex        =   126
      Top             =   675
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
      Left            =   2775
      TabIndex        =   120
      Top             =   1620
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
      Left            =   2775
      TabIndex        =   119
      Top             =   1215
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
      TabIndex        =   110
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
      Left            =   13050
      TabIndex        =   109
      Top             =   8730
      Visible         =   0   'False
      Width           =   1530
   End
   Begin VB.Shape Shape3 
      BorderColor     =   &H00EFB17A&
      BorderWidth     =   5
      FillColor       =   &H00EFB17A&
      Height          =   8775
      Left            =   0
      Top             =   0
      Width           =   14760
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
      Left            =   9750
      TabIndex        =   105
      Top             =   1035
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
      Left            =   8850
      TabIndex        =   101
      Top             =   7650
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
      TabIndex        =   100
      Top             =   7200
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
      TabIndex        =   99
      Top             =   6720
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
      TabIndex        =   98
      Top             =   7170
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
      TabIndex        =   97
      Top             =   7635
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
      Left            =   6930
      TabIndex        =   96
      Top             =   6705
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
      Left            =   6930
      TabIndex        =   95
      Top             =   7155
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
      Left            =   6930
      TabIndex        =   94
      Top             =   7650
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
      Left            =   8490
      TabIndex        =   88
      Top             =   2025
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
      Left            =   3690
      TabIndex        =   85
      Top             =   2025
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
      Left            =   2745
      TabIndex        =   84
      Top             =   2025
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
      TabIndex        =   83
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
      TabIndex        =   82
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
      TabIndex        =   81
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
      TabIndex        =   78
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
      TabIndex        =   77
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
      TabIndex        =   76
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
      TabIndex        =   39
      Top             =   765
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
      Left            =   450
      TabIndex        =   40
      Top             =   2400
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
      Left            =   180
      TabIndex        =   0
      Top             =   6195
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
      TabIndex        =   48
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
      Left            =   9000
      TabIndex        =   46
      Top             =   8100
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
      TabIndex        =   43
      Top             =   6315
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
      TabIndex        =   72
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
      Left            =   6840
      TabIndex        =   45
      Top             =   8145
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
      TabIndex        =   42
      Top             =   8190
      Width           =   1440
   End
   Begin VB.Label LblSaleType 
      BackColor       =   &H00EFB17A&
      Caption         =   "     Sales ( Late Payment Interest )"
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
      TabIndex        =   70
      Top             =   0
      Width           =   14715
   End
   Begin VB.Shape Shape1 
      FillColor       =   &H00FED8FC&
      FillStyle       =   0  'Solid
      Height          =   2550
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
      Left            =   8850
      TabIndex        =   44
      Top             =   6735
      Width           =   255
   End
   Begin VB.Shape Shape2 
      FillColor       =   &H00EDFDFE&
      FillStyle       =   0  'Solid
      Height          =   2445
      Left            =   180
      Top             =   540
      Width           =   9240
   End
   Begin VB.Shape Shape6 
      FillColor       =   &H00EDFDFE&
      FillStyle       =   0  'Solid
      Height          =   2580
      Left            =   9585
      Top             =   540
      Width           =   3345
   End
   Begin VB.Shape Shape5 
      FillColor       =   &H00EDFDFE&
      FillStyle       =   0  'Solid
      Height          =   2580
      Left            =   6435
      Top             =   6030
      Width           =   6495
   End
End
Attribute VB_Name = "frmSalesLpInt"
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
    If frmTypeStr = "MI" Then
       gClsSearch.SearchMultiField "tblIntSale,TblMastAccount", "vno,CrDrNoteNo,Vdt,AcName", Array("Vno", "Bill No", "Vdt", "Party"), Array(1500, 0, 1500, 4000), " tblIntSale.VDrAcCode=TblMastAccount.AcCode and vtype='" & frmTypeStr & "'   and VFirm = '" & gCCode & "'   and VYear=" & gCYear, "", "(cast(substring(CrDrNoteNo,7,15) as int) ) desc", 0, 0
    Else
        gClsSearch.SearchMultiField "tblIntSale,TblMastAccount", "vno,CrDrNoteNo,Vdt,AcName", Array("Vno", "Bill No", "Vdt", "Party"), Array(0, 1500, 1500, 4000), " tblIntSale.VDrAcCode=TblMastAccount.AcCode and vtype='" & frmTypeStr & "'   and VFirm = '" & gCCode & "'   and VYear=" & gCYear, "", "(cast(substring(CrDrNoteNo,7,15) as int) ) desc", 0, 0
    End If
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
       If txtCode(7).Enabled = True Then
          txtCode(7).SetFocus
       End If
    End If
    '-- Sales A/C
'    txttemp(6) = GProcGetColumnValue("tblMastAccount", "AcCode", CStr(SalesAccode), "N", "AcName", "S")
     txt(33) = GProcGenerateNoWithIBCrDrNt("SI", gCCode)
    SetGrid
    lblTinNo.Caption = ""
    lblMsOms.Caption = ""
Case 7 'Modify
    txtIrnNo.Enabled = True
    txtAckNo.Enabled = True
    txtIrnNo = txt(36).text
    txtAckNo = txt(37).text
    Frame2.Visible = True
    txtIrnNo.SetFocus
     
     
     
     
Case 8 'Delete
    'If mURecDel = True Then
      If IsRelatedRecord = False Then
        If MsgBox("Do You Really Delete This Record?", vbQuestion + vbYesNo + vbDefaultButton1, "Confirm Delete") = vbYes Then
           DeleteRelatedRecord
           GProcRstOpen rstbl, "Select * from " & tblName & " where " & IdField & " = " & txt(0) & " and " & Criteria, "O"
           rstbl.Delete
           cmdBtn_Click (3)
        End If
    End If
Case 9 'Print
    Frame1.Caption = "Sales Bill Printing "
    '-- Inv No
    txtPInvno(0) = txt(33).text
    txtPInvno(1) = txt(33).text
    Frame1.Visible = True
    DoEvents
    txtPInvno(0).Enabled = True
    txtPInvno(1).Enabled = True
    chkInvChallan(0).Enabled = True
    chkInvChallan(1).Enabled = True
    chkInvChallan(2).Enabled = True
    chkInvChallan(3).Enabled = True
    chkInvChallan(4).Enabled = True
    chkInvChallan(4).Value = 1
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
        If Trim(txtGrid) <> "" Then
            msGrid.text = txtGrid
            DoEvents
            txtGrid = ""
        End If
        txtGrid.Visible = False
'        If FormAction = vbDataActionAddNew Then txt(0) = GProcGenerateId(tblName, IdField, "vtype='" & frmTypeStr & "' and VYear=" & gCYear)
        CalAmount
        FillTxtFromTemp
        '-- Save Record
        If ValidateData = True Then
            If FormAction = vbDataActionAddNew Then
               Mon = Month(dtpDt(2))
               txt(0) = GProcGenerateIdMonthwise("tblIntSale", "Vno", Mon, "Vdt", "Vtype='" & frmTypeStr & "' and VFirm = '" & gCCode & "'   and VYear=" & gCYear)
            End If
            txt(11) = GProcGetColumnValue("tblMastAccount", "AcName", "-", "S", "AcCode", "N")
            GProcRstOpen rstbl, "Select * from " & tblName & " where " & IdField & " = " & txt(0) & " and " & Criteria, "O"
            GProcSaveRecord Me, rstbl, FormAction, MaxNo, OtherTxtFromNo, OtherTxtToNo
            SaveRelatedRecord
            mLastEntryDate = dtpDt(2)
            mLastEntryMill = txtCode(7)
            FormAction = vbDataActionClose
            Call GProcActivateControls(Me, False)
            Call GProcSetButtons(Me, True)
'            FormAction = vbDataActionClose
            msGrid.Col = 0 '-----
            msGrid.Row = 1 '-----
            txtGrid.Move msGrid.CellLeft + msGrid.Left, msGrid.CellTop + msGrid.Top, msGrid.CellWidth, msGrid.CellHeight '-----
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
    txtGrid.Visible = False
    If rstbl.RecordCount <= 0 Then
        Call GProcActivateControls(Me, False)
        Call GProcSetButtons(Me, vbDataActionCancel)
    Else
        Call GProcActivateControls(Me, False)
        Call GProcSetButtons(Me, True)
    End If
    If cmdBtn(4).Visible = True Then cmdBtn(4).SetFocus
    msGrid.Col = 0
    If rstbl.RecordCount > 0 Then
    msGrid.Row = 1
    End If
    txtGrid.Move msGrid.CellLeft + msGrid.Left, msGrid.CellTop + msGrid.Top, msGrid.CellWidth, msGrid.CellHeight
End Select
Exit Sub
ErrorRoutine:
    rstbl.CancelUpdate
    GProcErrorHandler
End Sub
Private Sub SetControlEd()
txt(8).Enabled = False '-- Amt
txt(21).Enabled = False '-- Ass Val
If FormAction = vbDataActionAddNew Then
   txt(0).Enabled = False
End If
If frmTypeStr = "MI" Then
   Label4.Visible = False
   txt(33).Visible = False
   txt(28).Visible = False
   txt(30).Visible = False
   txt(32).Visible = False
   txt(27).Visible = False
   txt(29).Visible = False
   txt(31).Visible = False
   lblSgst.Visible = False
   lblCgst.Visible = False
   lblIgst.Visible = False
   lblSgstAmt.Visible = False
   lblCgstAmt.Visible = False
   lblIgstAmt.Visible = False
   lblPer1.Visible = False
   lblPer2.Visible = False
   lblPer3.Visible = False
   
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
        & " b.SlSubBookNo,b.SlSubBookDt from tblSale A ,tblSaleSub B" _
        & " where a.vno = b.vno and a.vtype = b.vtype and a.vyear = b.vyear and a.vfirm = b.vfirm and  a.SlBillNo in " & X & " and a.Vtype='" & frmTypeStr & "' and a.VFirm = '" & gCCode & "' And a.VNoListVYear = " & gCYear
   
   gCn.Execute "Delete from tmpSale where  SlUser='" & gUserName & "'"
     
   gCn.Execute " insert into tmpSale (Vno,Vtype,Vdt,Vyear,VFirm,SlAcDrCode,SlAcCrCode,SlMillCode,SlBroker,SlBillNo,SlBillDt,SlTransport,SlLorryNo,SlLrNo,SlLrDate, " _
               & "SlSubAmt,SlBillAmt,SlAdd1,SlLess1,SlExemptAmt,SlExemptPerKg,SlTaxableAmt,SlTaxCode,SlTaxRate,SlTaxAmt,SlTaxableAmt2,SlTaxRate2,SlTaxAmt2," _
               & " SlAftTaxAdd1,SlAftTaxLess1,SlRoff,DueDays,DueDate,SlEntDt,SlUser,SlTranCd,SlIsHank,GpVno,GpNo,GpDate,GpVYear,SlTaxRate3, " _
               & " SlTaxAmt3,SlCharityRt,SlCharityOn,SlCrDays,SlTaxPayAc,SlIsExempt,SlSubItCode,SlSubBag,SlSubWt,SlSubRt,SlSubNetRate,SlSubRtPer,SlSubItAmt, " _
               & " SlSubBookNo,SlSubBookDt) " & X
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
        If chkInvChallan(0).Value = 1 Then
           PrintFor = chkInvChallan(0).Caption
           .Formulas(14) = "wPrintFor = '" & PrintFor & "'"
           .ReportFileName = gReportPath & "rptSaleBillPrintGst.rpt"
           If Index = 0 Or Index = 1 Then
              .Action = 1
           End If
        End If
        If chkInvChallan(1).Value = 1 Then
           PrintFor = chkInvChallan(1).Caption
           .Formulas(14) = "wPrintFor = '" & PrintFor & "'"
           .ReportFileName = gReportPath & "rptSaleBillPrintGst.rpt"
           If Index = 0 Or Index = 1 Then
              .Action = 1
           End If
        End If
        If chkInvChallan(2).Value = 1 Then
           PrintFor = chkInvChallan(2).Caption
           .Formulas(14) = "wPrintFor = '" & PrintFor & "'"
           .ReportFileName = gReportPath & "rptSaleBillPrintGst.rpt"
           If Index = 0 Or Index = 1 Then
              .Action = 1
           End If
        End If
        If chkInvChallan(3).Value = 1 Then
           PrintFor = chkInvChallan(3).Caption
           .Formulas(14) = "wPrintFor = '" & PrintFor & "'"
           .ReportFileName = gReportPath & "rptSaleBillPrintGst.rpt"
           If Index = 0 Or Index = 1 Then
              .Action = 1
           End If
        End If
        If chkInvChallan(4).Value = 1 Then
           PrintFor = chkInvChallan(4).Caption
           .Formulas(14) = "wPrintFor = '" & PrintFor & "'"
           .ReportFileName = gReportPath & "rptSaleBillGP.rpt"
           If Index = 0 Or Index = 1 Then
              .Action = 1
           End If
        End If
'        If optInvChallan(0).Value = True Then
'            PrintFor = optInvChallan(0).Caption
'        ElseIf optInvChallan(1).Value = True Then
'            PrintFor = optInvChallan(1).Caption
'        ElseIf optInvChallan(2).Value = True Then
'            PrintFor = optInvChallan(2).Caption
'        ElseIf optInvChallan(3).Value = True Then
'            PrintFor = optInvChallan(3).Caption
'        Else
'            PrintFor = ""
'        End If
'        .Formulas(14) = "wPrintFor = '" & PrintFor & "'"
'        X = "{tmpSale.SlUser}='" & gUserName & "' "
'       .SelectionFormula = X
'       If optInvChallan(4).Value = True Then
'          .ReportFileName = gReportPath & "rptSaleBillGP.rpt"
'       Else
'          .ReportFileName = gReportPath & "rptSaleBillPrintGst.rpt"
'       End If
'         If Index = 0 Or Index = 1 Then
'           .Action = 1
'         End If
     End With
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
        & " where  a.CrDrNoteNo in " & X & " and a.Vtype='" & frmTypeStr & "' and a.VFirm = '" & gCCode & "' And a.VYear = " & gCYear
     
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
       PartyName = GProcGetColumnValue("tblMastAccount", "AcCode", RsSA.Fields("VDrAcCode"), "N", "AcName", "S")
       PartyGstin = GProcGetColumnValue("tblMastAccount", "AcCode", RsSA.Fields("VDrAcCode"), "N", "AcGSTIN", "S")
       PartyAdd1 = GProcGetColumnValue("tblMastAccount", "AcCode", RsSA.Fields("VDrAcCode"), "N", "AcAdd1", "S")
       PartyAdd2 = GProcGetColumnValue("tblMastAccount", "AcCode", RsSA.Fields("VDrAcCode"), "N", "AcAdd2", "S")
       PartyCity = GProcGetColumnValue("tblMastAccount", "AcCode", RsSA.Fields("VDrAcCode"), "N", "AcCity", "S")
       PartyPin = GProcGetColumnValue("tblMastAccount", "AcCode", RsSA.Fields("VDrAcCode"), "N", "AcZip", "S")
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
       Print #1, Space(15) + Chr(34) + "Typ" + Chr(34) + ":" + Chr(34) + "DBN" + Chr(34) + Chr(44)
       Print #1, Space(15) + Chr(34) + "No" + Chr(34) + ":" + Chr(34) + RsSA.Fields(8) + Chr(34) + Chr(44)
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
       Print #1, Space(32) + Chr(34) + "PrdDesc" + Chr(34) + ":" + Chr(34) + "Late Payment Interest on Yarn" + Chr(34) + Chr(44)
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
          gCn.Execute "Update tblIntSale set SlIRNNo = '" & sheet.cells(a, 2) & "',SlAckNo = '" & sheet.cells(a, 3) & "'  where vtype='" & frmTypeStr & "' and CrDrNoteNo= '" & sheet.cells(a, 5) & "' and " _
          & " vfirm = '" & gCCode & "' and VYear=" & gCYear
       Else
          gCn.Execute "Update tblIntSale set SlIRNNo = '" & sheet.cells(a, 2) & "',SlAckNo = '" & sheet.cells(a, 3) & "'  where VDrAcCode= " & PtyCd & " and vtype='" & frmTypeStr & "' and CrDrNoteNo= '" & sheet.cells(a, 5) & "' and " _
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

Private Sub cmdUpdate_Click(Index As Integer)
Select Case Index
Case 0:
     gCn.BeginTrans
     gCn.Execute "Update tblIntSale set SlIRNNo = '" & txtIrnNo & "',SlAckNo = '" & txtAckNo & "'  where vtype='" & frmTypeStr & "' and Vno = " & txt(0) & " and " _
                 & " vfirm = '" & gCCode & "' and VYear=" & gCYear
     gCn.CommitTrans
     Frame2.Visible = False
Case 1:
    DoEvents
    Frame2.Visible = False
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
    Criteria = "Vtype='" & frmTypeStr & "' and VFirm in (" & BillsInFirm & ") and VYear=" & gCYear & "  and vnolist <> 0   "
    If Trim(mShowRec) <> "" Then
       GProcRstOpen rstbl, "Select * from " & tblName & " where  " & Criteria & " and  " & IdField & " = " & CLng(Trim(mShowRec)), "R"
       ReadFields
       Call GProcSetButtons(Me, vbDataActionCancel)
       mShowRec = ""
    Else
       GProcRstOpen rstbl, "Select * from " & tblName & " where " & Criteria & " and  " & IdField & " = 0", "R"
       SetGrid
       SetRelatedRS
       cmdBtn_Click (14)
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
If frmTypeStr = "MI" Then
    X = "select VBillNo,VBillType,VBillDate,VBillAmt,VAmt,Vno,Vdt,Grace,LateDays,IntRt,Interest,CgstRt+SgstRt+IgstRt,CgstAmt,SgstAmt,IgstAmt,VCtrNo " _
    & " from tblVoucher where " _
    & " tblVoucher.VType='BR' and VnoList=" & txt(23) & " and JvNoInt=" & txt(0) & "  and tblVoucher.VFirm in (" & BillsInFirm & ") and  tblVoucher.VNoListVYear=" & gCYear & " " _
    & " order by tblVoucher.Vno,VCtrNo"
Else
    X = "select VBillNo,VBillType,VBillDate,VBillAmt,VAmt,Vno,Vdt,Grace,LateDays,IntRt,Interest,CgstRt+SgstRt+IgstRt,CgstAmt,SgstAmt,IgstAmt,VCtrNo " _
    & " from tblVoucher where " _
    & " tblVoucher.VType in ('BR','CR') and VnoList=" & txt(23) & "  and CrDrNoteNo = '" & txt(33) & "'  and tblVoucher.VFirm in (" & BillsInFirm & ") and  tblVoucher.VNoListVYear=" & gCYear & " " _
    & " order by tblVoucher.Vno,VCtrNo"
End If
GProcRstOpen rstblSub, X, "R"
With msGrid
   .FormatString = "<Bill No           |<Tp |<Bill Date    |>Bill Amount    |>Amt. Recd     |>Vou. No. |<Vou. Date  |>Grase|>OD   |>Int. Rt.|>Interest       |>GST Rt|>CGST      |>SGST      |>IGST      |"
   .ColWidth(15) = 0
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
    If (dtpDt(12) = "01/04/1899") And IsDate(dtpDt(2)) = True Then dtpDt(12) = dtpDt(2)
    If dtpDt(2) < CDate("01/07/2017") Then
       MsgBox "Date in VAT Applicabel " & txt(9), vbCritical + vbOKOnly, ""
       dtpDt(Index).SetFocus
       Exit Sub
    End If
End If
End Sub
Private Sub ShowAccountBalance()
End Sub
Private Sub txt_GotFocus(Index As Integer)
GProcSelectBox Me.ActiveControl
End Sub
Private Sub CheckForNumber()
GProcCheckForNumber Me, rstbl, MaxNo, Array(33, 25), pOtherTxtFromNo:=OtherTxtFromNo, pOtherTxtToNo:=OtherTxtToNo
End Sub
Private Sub txt_LostFocus(Index As Integer)
If Index = 33 And FormAction <> vbDataActionCancel Then
   Dim X As String
   Dim rstAddCheck As Recordset
   X = "Select CrDrNoteNo From " & tblName & " Where CrDrNoteNo = '" & txt(33) & "'  and Vtype='" & frmTypeStr & "'  and Vfirm = '" & gCCode & "' and  VYear=" & gCYear
        i = GProcRstOpen(rstAddCheck, X, "R")
        If i > 0 Then
            MsgBox "Record is Available of Bill No. " & txt(33), vbCritical + vbOKOnly, "Duplicate Record"
            txtCode(7).SetFocus
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
If Index = 7 And txtCode(7) = "" Then
   txtCode(7) = "D"
   txtCode(7).SelStart = 2
End If
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
    Case 11  '-- Broker  Code
         txttemp(11) = GProcGetColumnValue("tblMastAccount", "AcOurCode", UCase(txtCode(11)), "S", "AcName", "S")
         If Trim(txttemp(11)) = "" Then
            MsgBox ("Code Not Found ...")
            txttemp(11) = ""
            txttemp(11).Enabled = True
            txttemp(11).SetFocus
         Else
            If txt(33).Enabled = True Then txt(33).SetFocus
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
        gClsSearch.SearchMultiField "tblMastaccount", "AcName,AcCity", Array("Sales Account", "City"), Array(3000, 2000), "(Agcode=90023)", Chr(KeyAscii), "", LeftPos, TopPos, True
        KeyAscii = 0
        If gClsSearch.SearchMultiRetCol(0) <> "" Then
            txttemp(Index) = gClsSearch.SearchMultiRetCol(0)
        End If
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
    With msGrid
        Select Case .Col
            Case 0, 1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14
                txtGrid.Locked = True
            Case Else
                txtGrid.Locked = False
        End Select
    End With
End If
msGrid_KeyPress (0)
End Sub
Private Sub msGrid_KeyPress(KeyAscii As Integer)
'--- Check is Related Gate Pass
'If IsRelatedItCtrlGatePass(msGrid.TextMatrix(msGrid.Row, 14)) = True Then Exit Sub
With msGrid
    If KeyAscii = 27 Then '----Esc Key
        KeyAscii = 0
        If (Val(.TextMatrix(.Row, 10)) = 0) And .rows > 2 Then
            .RemoveItem (.Row)
            txtGrid.text = ""
        End If
        txtGrid.Visible = False
        txt(21).SetFocus
        Exit Sub
    End If
End With
'If FormAction <> vbDataActionAddNew And FormAction <> vbDataActionUpdate Then Exit Sub
'---- Navigate Row ,Col
If KeyAscii = 13 Then
    KeyAscii = 0
    With msGrid
        If .Col = 14 Then 'rstblSub.Fields.Count - 2 Then
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
    txtGrid.Move msGrid.CellLeft + msGrid.Left, msGrid.CellTop + msGrid.Top, msGrid.CellWidth, msGrid.CellHeight
'End If
txtGrid.Visible = True
txtGrid.Enabled = True
DoEvents
If KeyAscii <> 0 Then
    Select Case rstblSub.Fields(msGrid.Col).Type
        Case adChar, adVarWChar
            txtGrid.text = txtGrid.text & Chr(KeyAscii)
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
'If msGrid.Col = 3 And txtGrid <> "" Then txtGrid.SelStart = 0
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
'''Dim MillCode As Long
'''Dim StdPack As Double
'''Dim NetRt As Double
'''Dim GAmt As Double
'''Dim GSubAmt As Double
'''Dim wCharityRt As Double
'''Dim wCharityOn As Integer
'''Dim wCharityAmt As Double
'''Dim Rs1 As Recordset
'''Dim VAddWithTax As Double
'''Dim ItCd As Long
'''Dim X As String
'''Dim ExmptKg As Double
'''Dim ExmptAmt As Double
'''Dim CalWt As Double
'''CheckForNumberGrid
'''If msGrid.TextMatrix(msGrid.Row, 1) = 0 Then msGrid.TextMatrix(msGrid.Row, 2) = "" '-- Book date
'''MillCode = GProcGetColumnValue("TblMastAccount", "AcName", txttemp(7), "S", "Accode", "N")
'''StdPack = GProcGetColumnValue("TblMastItem", "ITName", msGrid.TextMatrix(msGrid.Row, 3), "S", "ItStdpack", "N", " ItMillCode=" & MillCode)
'''wCharityRt = GProcGetColumnValue("TblMastItem", "ITName", msGrid.TextMatrix(msGrid.Row, 3), "S", "CharityRt", "N", " ItMillCode=" & MillCode)
'''wCharityOn = GProcGetColumnValue("TblMastItem", "ITName", msGrid.TextMatrix(msGrid.Row, 3), "S", "CharityOn", "N", " ItMillCode=" & MillCode)
'''ItCd = GProcGetColumnValue("tblMastItem", "ItName", msGrid.TextMatrix(msGrid.Row, 3), "S", "ItCode", "N", " ItMillCode=" & MillCode)
'''
'''
'''If msGrid.Row = 0 Then Exit Sub
''''-- Wt=Bag * Std Pack
'''If OldBag <> msGrid.TextMatrix(msGrid.Row, 6) And StdPack > 0 Then
'''    msGrid.TextMatrix(msGrid.Row, 8) = msGrid.TextMatrix(msGrid.Row, 6) * StdPack
'''    OldBag = msGrid.TextMatrix(msGrid.Row, 6)
'''End If
'''
''''If OldBag <> msGrid.TextMatrix(msGrid.Row, 6) And StdPack = 0 Then
''''    msGrid.TextMatrix(msGrid.Row, 8) = 0
''''End If
'''If CDbl(msGrid.TextMatrix(msGrid.Row, 10)) > 0 Then
'''    GAmt = Round(CDbl(msGrid.TextMatrix(msGrid.Row, 8)) * CDbl(msGrid.TextMatrix(msGrid.Row, 9)) / CDbl(msGrid.TextMatrix(msGrid.Row, 10)), 2)
'''Else
'''    GAmt = 0
'''End If
'''If frmTypeStr = "SE" And Val(msGrid.TextMatrix(msGrid.Row, 14)) > 0 Then
'''   X = "select tblpurch.vno,tblpurch.vtype,tblpurch.vyear,tblpurch.vfirm,tblPurch.PurExemptPerKg from tblpurch,tblpurchsub where tblpurch.vno = tblpurchsub.vno and tblpurch.vtype = tblpurchsub.vtype and " _
'''       & " tblpurch.vyear = tblpurchsub.vyear and tblpurch.vfirm = tblpurchsub.vfirm and tblpurch.vfirm = '" & gCCode & "' and tblpurchsub.PurSubItCtrlNo = '" & msGrid.TextMatrix(msGrid.Row, 14) & "'"
'''   If GProcRstOpen(Rs1, X, "R") > 0 Then
'''      ExmptKg = Rs1.Fields("PurExemptPerKg")
'''      ExmptAmt = (ExmptKg * Val(msGrid.TextMatrix(msGrid.Row, 8)))
'''   End If
'''   VAddWithTax = GAmt - ExmptAmt
'''    If VAddWithTax < 0 Then
'''        VAddWithTax = 0
'''        ExmptAmt = GAmt
'''    End If
'''    GSubAmt = Format(VAddWithTax - ((VAddWithTax * 100) / (100 + (Val(txt(23)) + Val(txt(26)) + Val(txt(41))))), LStr)
'''    GSubAmt = Round(GAmt - GSubAmt, 2)
'''Else
'''   '--Sub Amt=(Amt)*100/(100+Tax %)
'''   GSubAmt = Format(((GAmt * 100) / (100 + (Val(txt(23)) + Val(txt(26)) + Val(txt(41))))), LStr)
'''End If
''''- 8  - Weitht
''''- 9  - Rate
''''- 10  - Per
''''- 11 - Net Rate
''''- 12 - Sub Amt
'''If wCharityRt > 0 Then
''''And lblCharity.Visible = False Then
'''   lblCharity.Visible = True
'''   txt(17).Visible = True
'''   txt(43).Visible = True
'''   Cbo(44).Visible = True
'''
'''Else
'''   lblCharity.Visible = False
'''   txt(17).Visible = False
'''   txt(43).Visible = False
'''   Cbo(44).Visible = False
'''End If
'''
'''If wCharityRt > 0 Then
'''   If wCharityOn = 0 Then   '-- On Kg
'''      wCharityAmt = Format(Val(msGrid.TextMatrix(msGrid.Row, 8)) * wCharityRt, FStr)
'''   Else                     '-- On Pkg
'''      wCharityAmt = Format(Val(msGrid.TextMatrix(msGrid.Row, 6)) * wCharityRt, FStr)
'''   End If
'''   txt(43) = wCharityRt
'''   Cbo(44).ListIndex = wCharityOn
'''Else
'''   wCharityAmt = 0
'''End If
'''
'''If frmTypeStr = "SD" And txtCode(7) = "ISCCGMW" Then
'''   NetRt = Val(msGrid.TextMatrix(msGrid.Row, 9))
'''Else
'''    GSubAmt = GSubAmt - wCharityAmt
'''    If CDbl(msGrid.TextMatrix(msGrid.Row, 8)) > 0 Then
'''       NetRt = Format((GSubAmt / CDbl(msGrid.TextMatrix(msGrid.Row, 8))) * CDbl(msGrid.TextMatrix(msGrid.Row, 10)), LStr)
'''    Else
'''       NetRt = 0
'''    End If
'''End If
'''
'''msGrid.TextMatrix(msGrid.Row, 11) = NetRt
''''-- SubAmt=Wt*NetRate/Rate Per
'''If CDbl(msGrid.TextMatrix(msGrid.Row, 10)) > 0 Then
'''    msGrid.TextMatrix(msGrid.Row, 12) = CDbl(msGrid.TextMatrix(msGrid.Row, 8)) * CDbl(msGrid.TextMatrix(msGrid.Row, 11)) / CDbl(msGrid.TextMatrix(msGrid.Row, 10))
'''Else
'''    msGrid.TextMatrix(msGrid.Row, 12) = 0
'''End If
'''msGrid.TextMatrix(msGrid.Row, 12) = Format(msGrid.TextMatrix(msGrid.Row, 12), LStr)
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
'--- Delete Row On Ctrl+Y
If KeyCode = vbKeyY And Shift = 2 And msGrid.rows > 2 Then
    msGrid = 4
    msGrid.RemoveItem (msGrid.Row)
    txtGrid.Visible = False
    txtGrid.text = ""
    CalAmount
    Exit Sub
End If
End Sub
Private Sub txtGrid_KeyPress(KeyAscii As Integer)
Dim LeftPos As Long
Dim TopPos As Long
Dim Tp As String
Dim MillCode As Long
Dim BalBag As Integer
Dim BalWt As Double
Dim QStr As String
TopPos = txtGrid.Top + Me.Top + 650
LeftPos = txtGrid.Left + Me.Left
DoEvents
'--- Check is Related Gate Pass
'If IsRelatedItCtrlGatePass(msGrid.TextMatrix(msGrid.Row, 14)) = True Then Exit Sub
With msGrid
    If KeyAscii <> 13 And KeyAscii <> 27 And KeyAscii <> 25 Then
       KeyAscii = GProcValidateKey(rstblSub, msGrid.Col, KeyAscii, txtGrid)
   End If
    If KeyAscii = 13 Then
        CheckForNumberTxtGrid
    End If
    '--- Navigate Row ,Col
    If KeyAscii = 13 And .Col < 14 Then 'rstblSub.Fields.Count - 2 Then
        .Col = .Col + 1
    ElseIf KeyAscii = 13 And .Col = 14 Then '= rstblSub.Fields.Count - 2 Then
        If .Row = (.rows - 1) Then
            .rows = .rows + 1
            .Row = .Row + 1
            GridAddNew
        Else
            .Row = .Row + 1
        End If
        .Col = 0
    End If
    '-- On ESC
    If KeyAscii = 27 Or .Row = 2 Then
        KeyAscii = 0
        If (Val(.TextMatrix(.Row, 10)) = 0) And .rows > 2 Then
            .RemoveItem (.Row)
            txtGrid.text = ""
        End If
        txtGrid.Visible = False
        txtGrid.Visible = False
        txt(21).SetFocus
        Exit Sub
    End If
End With
End Sub

'-- Grid Required
Private Function ValidateGrid() As Boolean
'''Dim i As Long
'''Dim MillCode As Long
'''Dim RowIndex As Long
'''Dim ItCode As Long
'''Dim X As String
'''Dim Rs1 As Recordset
'''Dim oldFrmTypeStr As String
'''Dim LRNo As String
'''Dim pItCtrlNo As Long
'''If Trim(txtGrid.text) <> "" Then
'''    msGrid.text = txtGrid.text
'''    CalGridAmt
'''End If
'''txtGrid.Visible = False
'''txtGrid.text = ""
'''CheckForNumberGrid
'''With msGrid
'''Validate:
'''    For RowIndex = 1 To .rows - 1
'''        For i = 0 To rstblSub.Fields.Count - 1
'''            Select Case i
'''                Case 3 '-- Count
'''                 '-Rate=0 ,count,wt=0 then delete row
'''                    If RowIndex > 1 And Trim(.TextMatrix(RowIndex, 3)) = "" And CDbl(.TextMatrix(RowIndex, 7)) = 0 And CDbl(.TextMatrix(RowIndex, 8)) = 0 Then
'''                        .RemoveItem (RowIndex)
'''                        GoTo Validate
'''                    End If
'''                    If Trim(.TextMatrix(RowIndex, i)) = "" Then
'''                        MsgBox "Check Count.", vbCritical + vbOKOnly, Me.Caption
'''                        .Col = i
'''                        msGrid.SetFocus
'''                        Exit Function
'''                    End If
'''                    MillCode = GProcGetColumnValue("TblMastAccount", "AcName", txttemp(7), "S", "Accode", "N")
'''                    If GProcGetColumnValue("TblMastItem", "ItName", .TextMatrix(RowIndex, i), "S", "Itcode", "N", "ITMillCode=" & MillCode) = 0 Then
'''                        MsgBox "Check Count.Mill and Count are Mismatch.", vbCritical + vbOKOnly, Me.Caption
'''                        .Col = i
'''                        msGrid.SetFocus
'''                        Exit Function
'''                    End If
''''                Case 4 '-- Godown
''''                    If Trim(.TextMatrix(RowIndex, i)) = "" Then
''''                        MsgBox "Check Godown.", vbCritical + vbOKOnly, Me.Caption
''''                        .Col = i
''''                        msGrid.SetFocus
''''                        Exit Function
''''                    End If
''''                    If GProcGetColumnValue("TblMastNarration", "Narration", .TextMatrix(RowIndex, i), "S", "Narrcode", "N") = 0 Then
''''                        'MSGrid.SetFocus
''''                        oldFrmTypeStr = gfrmTypeStr
''''                        gfrmTypeStr = "G"
''''                        GProcShowForm frmMastNarration, frmMain.mnuMstfrm(8), Array(frmMastNarration.txt(2)), Array(.TextMatrix(RowIndex, i))
''''                        gfrmTypeStr = oldFrmTypeStr
''''                        Exit Function
''''                    End If
'''
'''                Case 5 '-- Lot No
''''                    If Trim(.TextMatrix(RowIndex, i)) = "" Then
''''                        MsgBox "Check Lot No.", vbCritical + vbOKOnly, Me.Caption
''''                        .Col = i
''''                        MSGrid.SetFocus
''''                        Exit Function
''''                    End If
''''                Case 6 '-- LR No
''''                    If Trim(.TextMatrix(RowIndex, i)) = "" Then
''''                        MsgBox "Check LR No.", vbCritical + vbOKOnly, Me.Caption
''''                        .Col = i
''''                        msGrid.SetFocus
''''                        Exit Function
''''                    End If
''''                Case 7 '-- LR Date
''''                    If IsDate(.TextMatrix(RowIndex, i)) = False Then
''''                        MsgBox "Check LR Date.", vbCritical + vbOKOnly, Me.Caption
''''                        .Col = i
''''                        msGrid.SetFocus
''''                        Exit Function
''''                    End If
''''                Case 8 '-- LR Amt
''''                    If CDbl(.TextMatrix(RowIndex, i)) = 0 Then
''''                        MsgBox "Check LR Amount.", vbCritical + vbOKOnly, Me.Caption
''''                        .Col = i
''''                        MSGrid.SetFocus
''''                        Exit Function
''''                    End If
'''                Case 6 '-- Bag
'''                    If CDbl(.TextMatrix(RowIndex, i)) = 0 Then
'''                        MsgBox "Check Bag.", vbCritical + vbOKOnly, Me.Caption
'''                        .Col = i
'''                        msGrid.SetFocus
'''                        Exit Function
'''                    End If
'''
''''                    '-- If Bag < Issue Bag
''''                    If CDbl(.TextMatrix(RowIndex, 7)) < CDbl(.TextMatrix(RowIndex, 13)) Then
''''                        MsgBox "Check Bag.Bag Less than issued.", vbCritical + vbOKOnly, Me.Caption
''''                        .Col = i
''''                        msGrid.SetFocus
''''                        Exit Function
''''                    End If
'''                Case 8 '-- Wt
'''                    If CDbl(.TextMatrix(RowIndex, i)) = 0 Then
'''                        MsgBox "Check Weight.", vbCritical + vbOKOnly, Me.Caption
'''                        .Col = i
'''                        msGrid.SetFocus
'''                        Exit Function
'''                    End If
''''                    '-- If Wt < Issue wt
''''                    If CDbl(.TextMatrix(RowIndex, 8)) < CDbl(.TextMatrix(RowIndex, 14)) Then
''''                        MsgBox "Check Wt.Weight Less than issued.", vbCritical + vbOKOnly, Me.Caption
''''                        .Col = i
''''                        msGrid.SetFocus
''''                        Exit Function
''''                    End If
'''                Case 12 '-- amount
'''                    If CDbl(.TextMatrix(RowIndex, i)) = 0 Then
'''                        MsgBox "Check amount.", vbCritical + vbOKOnly, Me.Caption
'''                        .Col = i
'''                        msGrid.SetFocus
'''                        Exit Function
'''                    End If
'''           End Select
'''        Next i
'''        '--- Duplicate LR No ,It Ctrl  No (Booking)
''''        LRNo = .TextMatrix(RowIndex, 6)
''''        pItCtrlNo = .TextMatrix(RowIndex, 19)
''''        For i = 1 To .Rows - 1
''''            '--- Duplicate LR No
''''            If i <> RowIndex And .TextMatrix(i, 6) = LRNo Then
''''                MsgBox "Duplicate LR No.", vbCritical + vbOKOnly, Me.Caption
''''                .Col = 6
''''                msGrid.SetFocus
''''                Exit Function
''''            End If
''''            '-- Duplicate Booking Count Selection
'''' '           If i <> RowIndex And .TextMatrix(i, 19) = pItCtrlNo And pItCtrlNo > 0 Then
''''  '              MsgBox "Duplicate Booking Count Selection.", vbCritical + vbOKOnly, Me.Caption
''''   '             .Col = 1
''''    '            msGrid.SetFocus
''''     '           Exit Function
''''      '      End If
''''        Next i
'''        '-- Invalid Booking
'''        If CLng(.TextMatrix(RowIndex, 1)) > 0 Then
'''            MillCode = GProcGetColumnValue("TblMastAccount", "AcName", txttemp(5), "S", "Accode", "N")
'''            ItCode = GProcGetColumnValue("tblMastItem", "ItName", .TextMatrix(RowIndex, 3), "S", "ItCode", "N", " ItMillCode=" & MillCode)
'''            X = "Select * from tblvouSub where Vno=" & CLng(.TextMatrix(RowIndex, 1)) & " and  " _
'''            & " VSUBITCODE=" & ItCode & " and  vsubitctrlno=" & .TextMatrix(RowIndex, 19) & " and Vtype='OT' and "
'''            If gBackEndDB = gBackEndAccess Then
'''                X = X + "  vdt=cdate('" & CDate(.TextMatrix(RowIndex, 2)) & "')"
''''                X = X + "  vdt=cdate('" & Format(CDate(.TextMatrix(RowIndex, 2)), "dd/mm/yyyy") & "')"
'''            Else '-- Oracle
'''                X = X + "  vdt= to_date('" & CDate(.TextMatrix(RowIndex, 2)) & "','dd/MM/yyyy')"
'''            End If
'''            GProcRstOpen Rs1, X, "R"
'''            If Rs1.EOF And Rs1.BOF Then
'''                MsgBox "Check Booking No.", vbInformation + vbOKOnly, Me.Caption
'''                .Col = 1
'''                msGrid.SetFocus
'''                Exit Function
'''            End If
'''            Rs1.Close
'''        End If
'''      Next RowIndex
'''End With
ValidateGrid = True
End Function
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

'-- Dr Code
txt(7) = GProcGetColumnValue("tblMastAccount", "AcName", txttemp(7), "S", "AcCode", "N")
'-- Cr Code
txt(6) = GProcGetColumnValue("tblMastAccount", "AcName", txttemp(6), "S", "AcCode", "N")
'-- broker
If txttemp(11) = "" Then
   txttemp(11) = "-"
End If
txt(11) = GProcGetColumnValue("tblMastNarration", "Narration", txttemp(11), "S", "NarrCode", "N")
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


'--Broker
txttemp(11) = GProcGetColumnValue("tblMastAccount", "AcCode", txt(11), "N", "AcName", "S")
txtCode(11) = GProcGetColumnValue("tblMastAccount", "AcCode", txt(11), "N", "AcOurCode", "S")

'-- Fill Related Record
FillRelatedRecord
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
'''Dim TaxAmt As Double
'''Dim TaxAmtCGST As Double
'''Dim TaxAmtIGST As Double
'''Dim SubAmt As Double
'''Dim Amt As Double
'''Dim Amt2 As Double
'''Dim ROff As Double
'''Dim ExmptAmt As Double
'''Dim ExmptKg As Double
'''Dim AssValue As Double
'''Dim RowIndex As Integer
'''Dim TotBag As Long
'''Dim TotWt As Double
'''Dim LCost As Double
'''Dim MillAcCode As Long
'''Dim IsExemptMill As Long
'''Dim wMillCode As Long
'''Dim wCharityRt As Double
'''Dim wCharityOn As Integer
'''Dim wCharityAmt As Double
'''Dim TotCharityAmt As Double
'''Dim TaxCd As Long
'''Dim tmpRst As Recordset
'''Dim X As String
''''-- Is Exempt Mill
'''IsExemptMill = GProcGetColumnValue("TblMastAccount", "AcName", txttemp(7), "S", "ACIsExemptMill", "N")
'''wMillCode = GProcGetColumnValue("tblMastAccount", "AcName", txttemp(7), "S", "AcCode", "N")
'''If Val(txt(31)) = 0 Then
'''   txt(31) = 20
'''End If
'''
'''
''''-- Due Date
'''If IsNumeric(txt(31)) = False Then txt(31) = "0"
'''txt(31) = Format(txt(31), PStr)
'''If IsDate(dtpDt(2)) Then dtpDt(32) = DateAdd("d", txt(31), dtpDt(2))
''''---- Sub amt
'''CheckForNumberGrid
'''CheckForNumber
'''TotCharityAmt = 0
'''ExmptKg = 0
'''ExmptAmt = 0
'''RowIndex = 1
'''With msGrid
'''    Do While RowIndex < .rows
'''        '-- sub amt
'''        If Val(.TextMatrix(RowIndex, 12)) > 0 Then
'''            SubAmt = SubAmt + Val(.TextMatrix(RowIndex, 12))
'''            If frmTypeStr = "SE" And Val(.TextMatrix(RowIndex, 14)) > 0 Then
'''               X = "select tblpurch.vno,tblpurch.vtype,tblpurch.vyear,tblpurch.vfirm,tblPurch.PurExemptPerKg from tblpurch,tblpurchsub where tblpurch.vno = tblpurchsub.vno and tblpurch.vtype = tblpurchsub.vtype and " _
'''                   & " tblpurch.vyear = tblpurchsub.vyear and tblpurch.vfirm = tblpurchsub.vfirm and tblpurch.vfirm = '" & gCCode & "' and tblpurchsub.PurSubItCtrlNo = '" & .TextMatrix(RowIndex, 14) & "'"
'''                If GProcRstOpen(tmpRst, X, "R") > 0 Then
'''                   ExmptKg = tmpRst.Fields("PurExemptPerKg")
'''                   ExmptAmt = ExmptAmt + (ExmptKg * Val(.TextMatrix(RowIndex, 8)))
'''                End If
'''            End If
'''        End If
'''        '--Tot Bag
'''        If Val(.TextMatrix(RowIndex, 6)) > 0 Then
'''            TotBag = TotBag + Val(.TextMatrix(RowIndex, 6))
'''        End If
'''        '--Tot Wt
'''        If Val(.TextMatrix(RowIndex, 8)) > 0 Then
'''            TotWt = TotWt + Val(.TextMatrix(RowIndex, 8))
'''        End If
'''
'''        If RowIndex = 1 And msGrid.TextMatrix(1, 3) <> "" Then
'''           ItemType = GProcGetColumnValue("tblMastItem", "ItName", msGrid.TextMatrix(1, 3), "S", "ItType", "N", "ItMillcode=" & wMillCode) '--Item Code
'''           If ItemType = 0 Then
'''              If Left(lblTinNo.Caption, 2) = "27" Or Left(lblTinNo.Caption, 1) = "U" Or Trim(lblTinNo.Caption) = "" Then
'''                 txt(23) = GProcGetColumnValue("tblMastNarration", "Narration", "G S T", "S", "CotCGSTRt", "N")
'''                 txt(26) = GProcGetColumnValue("tblMastNarration", "Narration", "G S T", "S", "CotSGSTRt", "N")
'''                 txt(41) = 0
'''              Else
'''                 txt(23) = 0
'''                 txt(26) = 0
'''                 txt(41) = GProcGetColumnValue("tblMastNarration", "Narration", "G S T", "S", "CotIGSTRt", "N")
'''              End If
'''           Else
'''              If Left(lblTinNo.Caption, 2) = "27" Or Left(lblTinNo.Caption, 1) = "U" Or Trim(lblTinNo.Caption) = "" Then
'''                 txt(23) = GProcGetColumnValue("tblMastNarration", "Narration", "G S T", "S", "PolCGSTRt", "N")
'''                 txt(26) = GProcGetColumnValue("tblMastNarration", "Narration", "G S T", "S", "PolSGSTRt", "N")
'''                 txt(41) = 0
'''              Else
'''                 txt(23) = 0
'''                 txt(26) = 0
'''                 txt(41) = GProcGetColumnValue("tblMastNarration", "Narration", "G S T", "S", "PolIGSTRt", "N")
'''              End If
'''           End If
'''        End If
'''        wCharityRt = GProcGetColumnValue("TblMastItem", "ITName", msGrid.TextMatrix(RowIndex, 3), "S", "CharityRt", "N", " ItMillCode=" & wMillCode)
'''        wCharityOn = GProcGetColumnValue("TblMastItem", "ITName", msGrid.TextMatrix(RowIndex, 3), "S", "CharityOn", "N", " ItMillCode=" & wMillCode)
'''        If wCharityRt > 0 Then
'''           If wCharityOn = 0 Then   '-- On Kg
'''              wCharityAmt = Format(Val(msGrid.TextMatrix(RowIndex, 8)) * wCharityRt, FStr)
'''           Else                     '-- On Pkg
'''              wCharityAmt = Format(Val(msGrid.TextMatrix(RowIndex, 6)) * wCharityRt, FStr)
'''           End If
'''        Else
'''           wCharityAmt = 0
'''        End If
'''        TotCharityAmt = TotCharityAmt + wCharityAmt
'''        RowIndex = RowIndex + 1
'''    Loop
'''End With
'''TotCharityAmt = Format(TotCharityAmt, FStr)
'''txt(17) = TotCharityAmt
'''txt(19) = Format(ExmptAmt, FStr)
'''txt(20) = ExmptKg
'''
''''--- Subamt+Add -Less
'''AssValue = SubAmt + CDbl(txt(17)) - CDbl(txt(18)) - CDbl(txt(19))
'''If AssValue < 0 Then
'''   AssValue = 0
'''End If
'''
'''TaxAmt = AssValue * CDbl(txt(23)) / 100
'''TaxAmt = Format(TaxAmt, FStr)
'''
''''- 2 CGST
'''If Abs((CDbl(txt(27)) - (AssValue * CDbl(txt(26)) / 100))) > 1 Then
'''   TaxAmtCGST = AssValue * CDbl(txt(26)) / 100
'''Else
'''   TaxAmtCGST = CDbl(txt(27))
'''End If
'''TaxAmtCGST = Format(TaxAmtCGST, FStr)
'''
''''- 3 IGST
'''If Abs((CDbl(txt(42)) - (AssValue * CDbl(txt(41)) / 100))) > 1 Then
'''   TaxAmtIGST = AssValue * CDbl(txt(41)) / 100
'''Else
'''   TaxAmtIGST = CDbl(txt(42))
'''End If
'''TaxAmtIGST = Format(TaxAmtIGST, FStr)
'''Amt = Format(SubAmt + TaxAmt + TaxAmtCGST + TaxAmtIGST + CDbl(txt(17)) - CDbl(txt(18)) + CDbl(txt(28)) - CDbl(txt(29)), FStr)
'''Amt2 = GProcMakeRounding(Amt)
'''ROff = Format(Amt2 - Format(Amt, FStr), FStr)
'''txt(15) = SubAmt
'''txt(24) = TaxAmt
'''txt(27) = TaxAmtCGST
'''txt(42) = TaxAmtIGST
'''txt(21) = AssValue
'''txt(30) = ROff
'''txt(16) = Round(Amt2, 0)
'''CheckForNumber


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
