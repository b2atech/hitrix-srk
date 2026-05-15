VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form frmSalesOtherGST 
   BackColor       =   &H00FBF2E1&
   BorderStyle     =   0  'None
   Caption         =   "Purchase Entry"
   ClientHeight    =   9075
   ClientLeft      =   0
   ClientTop       =   15
   ClientWidth     =   14730
   FillColor       =   &H000040C0&
   FillStyle       =   0  'Solid
   Icon            =   "frmSalesOtherGST.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   9075
   ScaleWidth      =   14730
   ShowInTaskbar   =   0   'False
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
      Index           =   71
      Left            =   15360
      TabIndex        =   168
      Text            =   "71"
      Top             =   600
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
      Index           =   70
      Left            =   14940
      TabIndex        =   167
      Text            =   "70"
      Top             =   1380
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
      Index           =   69
      Left            =   15000
      TabIndex        =   166
      Text            =   "69"
      Top             =   960
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
      Index           =   68
      Left            =   14940
      TabIndex        =   165
      Text            =   "68"
      Top             =   600
      Visible         =   0   'False
      Width           =   465
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
      Left            =   855
      TabIndex        =   161
      Text            =   "67"
      Top             =   8400
      Width           =   5445
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
      Left            =   840
      MaxLength       =   64
      MultiLine       =   -1  'True
      TabIndex        =   159
      Text            =   "frmSalesOtherGST.frx":058A
      Top             =   7740
      Width           =   5445
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
      Height          =   5295
      Left            =   2310
      TabIndex        =   125
      Top             =   2055
      Visible         =   0   'False
      Width           =   7785
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
         Left            =   2760
         TabIndex        =   158
         Top             =   4200
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
         Left            =   540
         TabIndex        =   157
         Top             =   4200
         Width           =   2055
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
         Left            =   4980
         TabIndex        =   156
         Top             =   4200
         Width           =   2055
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
         Left            =   4500
         TabIndex        =   139
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
         Left            =   4500
         TabIndex        =   138
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
         Left            =   4500
         TabIndex        =   137
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
         Left            =   4500
         TabIndex        =   136
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
         Left            =   4500
         TabIndex        =   135
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
         Left            =   4980
         TabIndex        =   130
         Top             =   3330
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
         Left            =   2760
         TabIndex        =   129
         Top             =   3345
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
         Left            =   2160
         MaxLength       =   10
         TabIndex        =   128
         Top             =   2220
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
         Left            =   540
         TabIndex        =   127
         Top             =   3330
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
         Left            =   2160
         MaxLength       =   10
         TabIndex        =   126
         Top             =   1425
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
         TabIndex        =   132
         Top             =   2302
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
         TabIndex        =   131
         Top             =   1515
         Width           =   1605
      End
      Begin VB.Line Line3 
         X1              =   0
         X2              =   7740
         Y1              =   2835
         Y2              =   2835
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
      Index           =   65
      Left            =   10440
      TabIndex        =   29
      Text            =   "65"
      Top             =   7500
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
      Left            =   9480
      TabIndex        =   25
      Text            =   "64"
      Top             =   7500
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
      Index           =   63
      Left            =   7380
      TabIndex        =   24
      Text            =   "63"
      Top             =   7500
      Width           =   2085
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
      Index           =   62
      Left            =   15960
      TabIndex        =   153
      Text            =   "62"
      Top             =   6780
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
      Index           =   61
      Left            =   16080
      TabIndex        =   152
      Text            =   "61"
      Top             =   6540
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
      Index           =   60
      Left            =   15960
      TabIndex        =   151
      Text            =   "60"
      Top             =   6180
      Visible         =   0   'False
      Width           =   465
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
      Height          =   360
      Index           =   59
      Left            =   870
      MaxLength       =   64
      TabIndex        =   34
      Text            =   "59"
      Top             =   7260
      Width           =   5415
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
      Height          =   360
      Index           =   58
      Left            =   870
      MaxLength       =   25
      TabIndex        =   33
      Text            =   "58"
      Top             =   6870
      Width           =   5415
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
      Height          =   360
      Index           =   57
      Left            =   870
      MaxLength       =   25
      TabIndex        =   32
      Text            =   "57"
      Top             =   6480
      Width           =   5415
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
      Index           =   56
      Left            =   15840
      TabIndex        =   148
      Text            =   "56"
      Top             =   5850
      Visible         =   0   'False
      Width           =   465
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
      Height          =   405
      Index           =   55
      Left            =   1560
      TabIndex        =   147
      Text            =   "55"
      Top             =   5760
      Visible         =   0   'False
      Width           =   765
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
      Index           =   54
      Left            =   4380
      TabIndex        =   146
      Text            =   "54"
      Top             =   5820
      Visible         =   0   'False
      Width           =   1665
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
      Index           =   53
      Left            =   15210
      TabIndex        =   145
      Text            =   "53"
      Top             =   7920
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
      Index           =   52
      Left            =   15255
      TabIndex        =   144
      Text            =   "52"
      Top             =   7515
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
      Index           =   51
      Left            =   15165
      TabIndex        =   143
      Text            =   "51"
      Top             =   7065
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
      Index           =   50
      Left            =   15210
      TabIndex        =   142
      Text            =   "50"
      Top             =   6660
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
      Index           =   49
      Left            =   15255
      TabIndex        =   141
      Text            =   "49"
      Top             =   6300
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
      Index           =   48
      Left            =   15210
      TabIndex        =   140
      Text            =   "48"
      Top             =   5940
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
      Height          =   375
      Index           =   47
      Left            =   19080
      TabIndex        =   122
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
      TabIndex        =   121
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
      TabIndex        =   120
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
      Height          =   390
      Index           =   5
      Left            =   1575
      TabIndex        =   2
      Text            =   "txtCode(5)"
      ToolTipText     =   "Press F3 for Creation"
      Top             =   1590
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
      TabIndex        =   117
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
      ItemData        =   "frmSalesOtherGST.frx":058D
      Left            =   8865
      List            =   "frmSalesOtherGST.frx":0597
      TabIndex        =   18
      Text            =   "Cbo(44)"
      Top             =   5325
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
      Height          =   450
      Index           =   43
      Left            =   8055
      TabIndex        =   17
      Text            =   "43"
      Top             =   5325
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
      Height          =   360
      Index           =   42
      Left            =   10440
      TabIndex        =   28
      Text            =   "42"
      Top             =   6930
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
      Left            =   8040
      TabIndex        =   23
      Text            =   "41"
      Top             =   6930
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
      TabIndex        =   106
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
      TabIndex        =   105
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
      Left            =   10920
      TabIndex        =   7
      Text            =   "38"
      Top             =   2415
      Visible         =   0   'False
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
      Index           =   9
      Left            =   10920
      TabIndex        =   4
      Text            =   "9"
      Top             =   1320
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
      TabIndex        =   104
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
      TabIndex        =   103
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
      Left            =   13305
      Style           =   1  'Graphical
      TabIndex        =   101
      Top             =   6480
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
      TabIndex        =   100
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
      TabIndex        =   41
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
      TabIndex        =   94
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
      Left            =   9090
      TabIndex        =   31
      Text            =   "16"
      Top             =   8400
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
      TabIndex        =   16
      Text            =   "15"
      Top             =   4980
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
      TabIndex        =   93
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
      TabIndex        =   12
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
      Left            =   16200
      TabIndex        =   10
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
      TabIndex        =   39
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
      TabIndex        =   40
      Text            =   "29"
      Top             =   3105
      Visible         =   0   'False
      Width           =   330
   End
   Begin VB.CommandButton CmdBagDetails 
      Caption         =   "Cartoon &Bag Details"
      Height          =   720
      Left            =   13305
      TabIndex        =   87
      Top             =   2985
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
      TabIndex        =   47
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
      TabIndex        =   86
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
      Left            =   14280
      TabIndex        =   38
      Text            =   "19"
      Top             =   8550
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
      TabIndex        =   85
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
      TabIndex        =   84
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
      Height          =   390
      Index           =   5
      Left            =   2775
      TabIndex        =   3
      Text            =   "txttemp(5)"
      ToolTipText     =   "Press F3 for Creation"
      Top             =   1590
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
      TabIndex        =   36
      Text            =   "txtDrBalance"
      Top             =   1170
      Width           =   300
   End
   Begin VB.CommandButton cmdLedger 
      Caption         =   "Ledger"
      Height          =   375
      Left            =   12915
      TabIndex        =   83
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
      Height          =   360
      Index           =   27
      Left            =   10440
      TabIndex        =   27
      Text            =   "27"
      Top             =   6540
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
      Left            =   8040
      TabIndex        =   22
      Text            =   "26"
      Top             =   6540
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
      TabIndex        =   59
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
      TabIndex        =   20
      Text            =   "21"
      Top             =   5745
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
      TabIndex        =   15
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
      TabIndex        =   81
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
      TabIndex        =   79
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
      TabIndex        =   78
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
      TabIndex        =   77
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
      TabIndex        =   73
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
      TabIndex        =   72
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
      TabIndex        =   70
      ToolTipText     =   "Exit"
      Top             =   7650
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
      TabIndex        =   63
      ToolTipText     =   "To Save Record"
      Top             =   5835
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
      TabIndex        =   71
      Top             =   7065
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
      TabIndex        =   62
      ToolTipText     =   "To Delete Record"
      Top             =   5205
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
      TabIndex        =   61
      ToolTipText     =   "To Modify Record"
      Top             =   4560
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
      TabIndex        =   60
      ToolTipText     =   "To Add New Record"
      Top             =   3915
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
      TabIndex        =   64
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
      TabIndex        =   65
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
      Left            =   13800
      Style           =   1  'Graphical
      TabIndex        =   69
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
      TabIndex        =   68
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
      TabIndex        =   67
      ToolTipText     =   "Next"
      Top             =   1080
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
      Height          =   375
      Index           =   0
      Left            =   10920
      TabIndex        =   35
      Text            =   "0"
      Top             =   780
      Width           =   1680
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
      TabIndex        =   66
      ToolTipText     =   "Previous"
      Top             =   1080
      Width           =   375
   End
   Begin VB.TextBox txt 
      Alignment       =   1  'Right Justify
      Height          =   315
      Index           =   4
      Left            =   18270
      TabIndex        =   76
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
      Height          =   390
      Index           =   6
      Left            =   2775
      Locked          =   -1  'True
      TabIndex        =   1
      Text            =   "txttemp(6)"
      ToolTipText     =   "Press F3 for Creation"
      Top             =   720
      Width           =   6540
   End
   Begin VB.TextBox txt 
      Height          =   315
      Index           =   6
      Left            =   18405
      TabIndex        =   75
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
      TabIndex        =   74
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
      TabIndex        =   49
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
      TabIndex        =   57
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
      TabIndex        =   11
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
      Height          =   405
      Index           =   23
      Left            =   8040
      TabIndex        =   21
      Text            =   "23"
      Top             =   6150
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
      Height          =   360
      Index           =   24
      Left            =   10440
      TabIndex        =   26
      Text            =   "24"
      Top             =   6150
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
      Left            =   14235
      TabIndex        =   45
      Text            =   "20"
      Top             =   8955
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
      TabIndex        =   30
      Text            =   "30"
      Top             =   7980
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
      Left            =   16875
      TabIndex        =   42
      Text            =   "txt(31)"
      Top             =   5085
      Visible         =   0   'False
      Width           =   435
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
      Height          =   360
      Index           =   17
      Left            =   10440
      TabIndex        =   19
      Text            =   "17"
      Top             =   5355
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
      TabIndex        =   37
      Text            =   "18"
      Top             =   2295
      Visible         =   0   'False
      Width           =   330
   End
   Begin MSMask.MaskEdBox mskDt 
      Height          =   390
      Index           =   2
      Left            =   15165
      TabIndex        =   6
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
      TabIndex        =   9
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
      TabIndex        =   44
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
      Bindings        =   "frmSalesOtherGST.frx":05A4
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
      Height          =   1125
      Left            =   180
      TabIndex        =   14
      Top             =   3660
      Width           =   12750
      _ExtentX        =   22490
      _ExtentY        =   1984
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
      Left            =   16830
      TabIndex        =   13
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
      TabIndex        =   107
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
      Left            =   10920
      TabIndex        =   5
      Top             =   1845
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
      Format          =   116391939
      CurrentDate     =   -273
   End
   Begin MSComCtl2.DTPicker dtpDt 
      Height          =   360
      Index           =   39
      Left            =   10920
      TabIndex        =   8
      Top             =   2925
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
      Format          =   116391939
      CurrentDate     =   -273
   End
   Begin MSComCtl2.DTPicker dtpDt 
      Height          =   360
      Index           =   32
      Left            =   16380
      TabIndex        =   43
      Top             =   5490
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
      Format          =   116391939
      CurrentDate     =   -273
   End
   Begin VB.Label Label20 
      BackStyle       =   0  'Transparent
      Caption         =   "Vou. No   :"
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
      Left            =   9780
      TabIndex        =   170
      Top             =   780
      Width           =   1080
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
      TabIndex        =   169
      Top             =   5040
      Width           =   6195
   End
   Begin VB.Label Label19 
      BackStyle       =   0  'Transparent
      Caption         =   "Comm. On Amt.  :"
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
      Left            =   2640
      TabIndex        =   164
      Top             =   5820
      Visible         =   0   'False
      Width           =   1680
   End
   Begin VB.Label Label15 
      BackStyle       =   0  'Transparent
      Caption         =   "Comm. @ :"
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
      Left            =   420
      TabIndex        =   163
      Top             =   5820
      Visible         =   0   'False
      Width           =   1140
   End
   Begin VB.Label Label29 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Ack.  "
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
      Left            =   180
      TabIndex        =   162
      Top             =   8460
      Width           =   525
   End
   Begin VB.Label Label27 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "IRN "
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
      Left            =   180
      TabIndex        =   160
      Top             =   7740
      Width           =   390
   End
   Begin VB.Label Label12 
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
      TabIndex        =   155
      Top             =   7560
      Width           =   255
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "TCS Amt      : "
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
      Left            =   15120
      TabIndex        =   154
      Top             =   8580
      Width           =   1335
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "TCS On  : "
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
      Left            =   6510
      TabIndex        =   150
      Top             =   7560
      Width           =   885
   End
   Begin VB.Line Line1 
      X1              =   6420
      X2              =   12900
      Y1              =   7380
      Y2              =   7380
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Note  :"
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
      Left            =   150
      TabIndex        =   149
      Top             =   6525
      Width           =   900
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
      TabIndex        =   134
      Top             =   2520
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
      TabIndex        =   133
      Top             =   2115
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
      TabIndex        =   124
      Top             =   9000
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
      TabIndex        =   123
      Top             =   8595
      Visible         =   0   'False
      Width           =   1530
   End
   Begin VB.Shape Shape3 
      BorderColor     =   &H00EFB17A&
      BorderWidth     =   5
      FillColor       =   &H00EFB17A&
      Height          =   9045
      Left            =   45
      Top             =   0
      Width           =   14670
   End
   Begin VB.Label Label11 
      BackStyle       =   0  'Transparent
      Caption         =   "Date         : "
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
      TabIndex        =   119
      Top             =   2925
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Date         : "
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
      TabIndex        =   118
      Top             =   1845
      Width           =   1095
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
      TabIndex        =   116
      Top             =   5348
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
      TabIndex        =   115
      Top             =   6923
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
      TabIndex        =   114
      Top             =   6533
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
      TabIndex        =   113
      Top             =   6143
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
      TabIndex        =   112
      Top             =   6533
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
      TabIndex        =   111
      Top             =   6923
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
      TabIndex        =   110
      Top             =   6143
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
      TabIndex        =   109
      Top             =   6533
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
      TabIndex        =   108
      Top             =   6923
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
      TabIndex        =   102
      Top             =   2925
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
      TabIndex        =   99
      Top             =   2925
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
      TabIndex        =   98
      Top             =   2925
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
      TabIndex        =   97
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
      TabIndex        =   96
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
      TabIndex        =   95
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
      TabIndex        =   92
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
      TabIndex        =   91
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
      TabIndex        =   90
      Top             =   1845
      Visible         =   0   'False
      Width           =   1320
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "G.P. No   :    "
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
      Left            =   9750
      TabIndex        =   89
      Top             =   2475
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Bill No     :"
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
      TabIndex        =   88
      Top             =   1380
      Width           =   990
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
      TabIndex        =   46
      Top             =   1620
      Width           =   900
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
      Left            =   450
      TabIndex        =   0
      Top             =   750
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
      TabIndex        =   58
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
      TabIndex        =   48
      Top             =   4973
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
      TabIndex        =   50
      Top             =   5348
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
      TabIndex        =   51
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
      TabIndex        =   56
      Top             =   8520
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
      TabIndex        =   53
      Top             =   5790
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
      TabIndex        =   82
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
      TabIndex        =   55
      Top             =   7980
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
      TabIndex        =   52
      Top             =   8190
      Width           =   1440
   End
   Begin VB.Label LblSaleType 
      BackColor       =   &H00EFB17A&
      Caption         =   "     Sales ( Other )"
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
      Height          =   420
      Left            =   0
      TabIndex        =   80
      Top             =   45
      Width           =   17055
   End
   Begin VB.Shape Shape1 
      FillColor       =   &H00FED8FC&
      FillStyle       =   0  'Solid
      Height          =   7770
      Left            =   13050
      Shape           =   4  'Rounded Rectangle
      Top             =   675
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
      TabIndex        =   54
      Top             =   6143
      Width           =   255
   End
   Begin VB.Shape Shape4 
      FillColor       =   &H00ECFFFF&
      FillStyle       =   0  'Solid
      Height          =   2595
      Left            =   90
      Top             =   6300
      Width           =   6300
   End
   Begin VB.Shape Shape2 
      FillColor       =   &H00EDFDFE&
      FillStyle       =   0  'Solid
      Height          =   2940
      Left            =   180
      Top             =   540
      Width           =   9240
   End
   Begin VB.Shape Shape6 
      FillColor       =   &H00EDFDFE&
      FillStyle       =   0  'Solid
      Height          =   2940
      Left            =   9585
      Top             =   540
      Width           =   3345
   End
   Begin VB.Shape Shape5 
      FillColor       =   &H00EDFDFE&
      FillStyle       =   0  'Solid
      Height          =   4065
      Left            =   6435
      Top             =   4860
      Width           =   6495
   End
End
Attribute VB_Name = "frmSalesOtherGST"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Const tblName As String = "tblSale"
Const IdField As String = "Vno"
Const OrderField As String = "Vdt,Vno"
'Const MaxNo As Long = 44
Const MaxNo As Long = 71
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
'On Error GoTo ErrorRoutine
'-- For Log
Dim j As Long
Dim tmpRst As Recordset
Dim wLogNo As Integer
Dim AMD As String

Select Case Index
Case 0 '---- Previous
     If GProcRstOpen(rstbl, "Select * from " & tblName & " where " & Criteria & "  and " & IdField & " = (Select max(" & IdField & ") from " & tblName & " where " & Criteria & " And " & IdField & "  < " & Val(txt(0)) & ")", "R") > 0 Then ReadFields Else MsgBox "This is First Record:", vbOKOnly + vbInformation
     Call GProcActivateControls(Me, False)
Case 1 '---- Next
      If GProcRstOpen(rstbl, "Select * from " & tblName & " where " & Criteria & "  and   " & IdField & " = (Select min(" & IdField & ") from " & tblName & " where " & IdField & " > '" & txt(0) & "' and " & Criteria & ") and " & Criteria, "R") > 0 Then ReadFields Else MsgBox "This is last Record:", vbOKOnly + vbInformation
     Call GProcActivateControls(Me, False)
Case 2 '---- First
         If GProcRstOpen(rstbl, "Select * from " & tblName & " where " & Criteria & "  and  " & IdField & " = (Select min(" & IdField & ") from " & tblName & " where " & IdField & " <> '' and " & Criteria & ")and " & Criteria, "R") > 0 Then ReadFields
     Call GProcActivateControls(Me, False)
Case 3 '---- Last
'        If GProcRstOpen(rstbl, "Select * from " & tblName & " where " & Criteria & "  and  " & IdField & " = (Select max(" & IdField & ") from " & tblName & " where " & IdField & " <> '' and " & Criteria & ")and " & Criteria, "R") > 0 Then ReadFields Else GProcClearForm Me, rstbl, MaxNo
        If GProcRstOpen(rstbl, "Select * from " & tblName & " where " & Criteria & "  and  " & IdField & " = (Select max(" & IdField & ") from " & tblName & " where vdt = (SELECT MAX(vdt) from tblSale where " & Criteria & " )  and  " & IdField & " <> 0 and  " & Criteria & ")", "R") > 0 Then ReadFields Else GProcClearForm Me, rstbl, MaxNo
     Call GProcActivateControls(Me, False)

Case 4 'Find ' QryMastAccount
'    gClsSearch.SearchMultiField "tblSale,TblMastAccount", "vno,SlBillNo,Vdt,AcName", Array("Vno", "Bill No", "Vdt", "Party"), Array(0, txt(9).Width, 1500, 4000), " tblSale.SlAcDrCode=TblMastAccount.AcCode and vtype='" & frmTypeStr & "'  and SlIsHank = 0  and VFirm = '" & gCCode & "'   and VYear=" & gCYear, "", "(cast(substring(SlBillNo,6,15) as int) ) desc", 0, 0
    gClsSearch.SearchMultiField "tblSale,TblMastAccount", "vno,SlBillNo,Vdt,AcName", Array("Vno", "Bill No", "Vdt", "Party"), Array(txt(0).Width, txt(9).Width, 1500, 4000), " tblSale.SlAcDrCode=TblMastAccount.AcCode and vtype='" & frmTypeStr & "'  and SlIsHank = 0  and VFirm = '" & gCCode & "'   and VYear=" & gCYear, "", "SlBillNo", 0, 0
     If gClsSearch.SearchMultiRetCol(0) <> "" Then
        GProcRstOpen rstbl, "Select * from " & tblName & " where " & IdField & " = " & gClsSearch.SearchMultiRetCol(0) & " and " & Criteria, "R"
        gClsSearch.SearchMultiRetCol(1) = ""
        ReadFields
        Call GProcActivateControls(Me, False)
    
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
    End If
    If gfrmTypeStr = "SB" Then
       txt(9) = GProcGenerateInvNoBrokComm(frmTypeStr, gCCode, gCCode)  '--Inv No
    Else
'       GBillNoWithIB
    End If
    If gfrmHunkStr = "TCS" Then
       txt(9) = GProcGenerateInvNoBrokComm(frmTypeStr, "TCS", gCCode)   '--Inv No
    End If
    FillCombo
    SetGrid
'    SetControlEd
'    txttemp(6).SetFocus
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
'            txt(9).Enabled = False
            dtpDt(2).SetFocus
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
    If frmTypeStr = "SC" Then
       txtPInvno(1).Visible = False
       txtPInvno(1).Enabled = False
    Else
       txtPInvno(1).Visible = True
       txtPInvno(1).Enabled = True
    End If
    chkInvChallan(0).Value = 1
    chkInvChallan(0).Enabled = True
    chkInvChallan(1).Enabled = True
    chkInvChallan(2).Enabled = True
    chkInvChallan(3).Enabled = True
    chkInvChallan(4).Enabled = True
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
     lblTcsLimit.Caption = ""
        
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
               txt(0) = GProcGenerateIdMonthwise("TblSale", "Vno", Mon, "Vdt", "Vtype='" & frmTypeStr & "' and VFirm = '" & gCCode & "'   and VYear=" & gCYear)
            End If
            txt(46) = GProcGetColumnValue("tblMastAccount", "AcName", "-", "S", "AcCode", "N")
            txt(22) = GProcGetColumnValue("tblMastNarration", "Narration", "-", "S", "NarrCode", "N")
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
End Select
Exit Sub
ErrorRoutine:
    rstbl.CancelUpdate
    GProcErrorHandler
End Sub
Private Sub SetControlEd()
txttemp(35).Enabled = False
txt(16).Enabled = False '-- Amt
txt(21).Enabled = False '-- Ass Val
dtpDt(32).Enabled = False '--Due Dt
txtDrBalance.Enabled = False
txt(24).Enabled = False
'txt(30).Enabled = False
If FormAction = vbDataActionAddNew Then
   txt(0).Enabled = False
End If
'txt(23).Enabled = False
'txt(26).Enabled = False
'txt(41).Enabled = False
txt(24).Enabled = True
txt(27).Enabled = True
txt(42).Enabled = True

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
txt(63).Enabled = False
txt(64).Enabled = False
txt(65).Enabled = False
'If frmTypeStr = "SB" Then
'   Label19.Visible = True
'   Label15.Visible = True
'   txt(54).Visible = True
'   txt(55).Visible = True
'   If gCYear = 2021 Then
'      txt(54).Enabled = True
'      txt(55).Enabled = False
'   Else
'      txt(54).Enabled = False
'      txt(55).Enabled = False
'   End If
'End If
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
MillCode = GProcGetColumnValue("tblMastAccount", "AcName", "-", "S", "AcCode", "N")
DrCode = GProcGetColumnValue("tblMastAccount", "AcName", txttemp(5), "S", "AcCode", "N")
CrCode = GProcGetColumnValue("tblMastAccount", "AcName", txttemp(6), "S", "AcCode", "N")
BrkCode = GProcGetColumnValue("tblMastAccount", "AcName", "-", "S", "AcCode", "N")
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
        '--Itcode(4)
        rstSub!SlSubItCode = GProcGetColumnValue("tblMastItem", "ItName", .TextMatrix(i, 1), "S", "ItCode", "N", " ItMillCode=" & MillCode) '--Item Code
        rstSub!SlSubAmt = CDbl(.TextMatrix(i, 2))
        rstSub!SlSubGodown = GProcGetColumnValue("tblMastNarration", "Narration", "-", "S", "NarrCode", "N", "")
        '-- P It Ctrl No =Booking It ctrl No
        rstSub.Update
        i = i + 1
     Loop
 End With
 rstSub.Close
 cmdBtn(13).Enabled = False
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
    If frmTypeStr = "SC" Then
       X = "('" + Trim(txtPInvno(0)) + "')"
    Else
        If Val(txtPInvno(0)) > 0 Then
           InvNo = CLng(txtPInvno(0))
           ToInvNo = CLng(txtPInvno(1))
           BillSrNo = ""
        Else
           InvNo = CLng(Mid(txtPInvno(0), 6, 15))
           ToInvNo = CLng(Mid(txtPInvno(1), 6, 15))
           BillSrNo = Left(txtPInvno(0), 5)
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
             X = X + "'" & BillSrNo & "" & InvNo & "',"
           End If
           InvNo = InvNo + 1
        Loop
        X = Left(X, Len(X) - 1) + ")"
    End If

    Y = X
    X = "Select a.Vno,a.Vtype,a.Vdt,a.Vyear,a.VFirm,a.SlAcDrCode,a.SlAcCrCode,a.SlMillCode,a.SlBroker,a.SlBillNo,a.SlBillDt,a.SlTransport,a.SlLorryNo,a.SlLrNo,a.SlLrDate, " _
        & "a.SlSubAmt,a.SlBillAmt,a.SlAdd1,a.SlLess1,a.SlExemptAmt,a.SlExemptPerKg,a.SlTaxableAmt,a.SlTaxCode,a.SlTaxRate,a.SlTaxAmt,a.SlTaxableAmt2,a.SlTaxRate2,a.SlTaxAmt2, " _
        & " a.SlAftTaxAdd1,a.SlAftTaxLess1,a.SlRoff,a.DueDays,a.DueDate,a.SlEntDt,'" & gUserName & "',a.SlTranCd,a.SlIsHank,a.GpVno,a.GpNo,a.GpDate,a.GpVYear,a.SlTaxRate3, " _
        & " a.SlTaxAmt3,a.SlCharityRt,a.SlCharityOn,a.SlCrDays,a.SlTaxPayAc,a.SlIsExempt,b.SlSubItCode,b.SlSubBag,b.SlSubWt,b.SlSubRt,b.SlSubNetRate,b.SlSubRtPer,b.SlSubAmt, " _
        & " b.SlSubBookNo,b.SlSubBookDt,a.SlTmpChr1,a.SlTmpChr2,a.SlTmpChr3,a.SlIRNNo,a.SlAckNo from tblSale A ,tblSaleSub B" _
        & " where a.vno = b.vno and a.vtype = b.vtype and a.vyear = b.vyear and a.vfirm = b.vfirm and  a.SlBillNo in " & X & " and a.Vtype='" & frmTypeStr & "' and a.VFirm = '" & gCCode & "' And a.VYear = " & gCYear
     
   gCn.Execute "Delete from tmpSale where  SlUser='" & gUserName & "'"
     
   gCn.Execute " insert into tmpSale (Vno,Vtype,Vdt,Vyear,VFirm,SlAcDrCode,SlAcCrCode,SlMillCode,SlBroker,SlBillNo,SlBillDt,SlTransport,SlLorryNo,SlLrNo,SlLrDate, " _
               & "SlSubAmt,SlBillAmt,SlAdd1,SlLess1,SlExemptAmt,SlExemptPerKg,SlTaxableAmt,SlTaxCode,SlTaxRate,SlTaxAmt,SlTaxableAmt2,SlTaxRate2,SlTaxAmt2," _
               & " SlAftTaxAdd1,SlAftTaxLess1,SlRoff,DueDays,DueDate,SlEntDt,SlUser,SlTranCd,SlIsHank,GpVno,GpNo,GpDate,GpVYear,SlTaxRate3, " _
               & " SlTaxAmt3,SlCharityRt,SlCharityOn,SlCrDays,SlTaxPayAc,SlIsExempt,SlSubItCode,SlSubBag,SlSubWt,SlSubRt,SlSubNetRate,SlSubRtPer,SlSubItAmt, " _
               & " SlSubBookNo,SlSubBookDt,SlTmpChr1,SlTmpChr2,SlTmpChr3,SlIRNNo,SlAckNo) " & X
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
           .ReportFileName = gReportPath & "rptSaleOtherBillPrintGst.rpt"
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
    If frmTypeStr = "SC" Then
       X = "('" + txtPInvno(0) + "')"
    Else
        If Val(txtPInvno(0)) > 0 Then
           InvNo = CLng(txtPInvno(0))
           ToInvNo = CLng(txtPInvno(1))
           BillSrNo = ""
        Else
           InvNo = CLng(Mid(txtPInvno(0), 6, 15))
           ToInvNo = CLng(Mid(txtPInvno(1), 6, 15))
           BillSrNo = Left(txtPInvno(0), 5)
        End If
'        ------ Validate
        If IsNumeric(InvNo) = False Then InvNo = 0
        If IsNumeric(ToInvNo) = False Then ToInvNo = InvNo
        If ToInvNo < InvNo Then ToInvNo = InvNo
        
        X = "("
        Do While InvNo <= ToInvNo
           If BillSrNo = "" Then
              X = X + "'" & InvNo & "',"
           Else
              X = X + "'" & BillSrNo & "" & InvNo & "',"
           End If
           InvNo = InvNo + 1
        Loop
        X = Left(X, Len(X) - 1) + ")"
    End If
    Y = X
    X = "Select a.Vno,a.Vtype,a.Vdt,a.Vyear,a.VFirm,a.SlAcDrCode,a.SlAcCrCode,a.SlMillCode,a.SlBroker,a.SlBillNo,a.SlBillDt,a.SlTransport,a.SlLorryNo,a.SlLrNo,a.SlLrDate, " _
        & "a.SlSubAmt,a.SlBillAmt,a.SlAdd1,a.SlLess1,a.SlExemptAmt,a.SlExemptPerKg,a.SlTaxableAmt,a.SlTaxCode,a.SlTaxRate,a.SlTaxAmt,a.SlTaxableAmt2,a.SlTaxRate2,a.SlTaxAmt2, " _
        & " a.SlAftTaxAdd1,a.SlAftTaxLess1,a.SlRoff,a.DueDays,a.DueDate,a.SlEntDt,'" & gUserName & "',a.SlTranCd,a.SlIsHank,a.GpVno,a.GpNo,a.GpDate,a.GpVYear,a.SlTaxRate3, " _
        & " a.SlTaxAmt3,a.SlCharityRt,a.SlCharityOn,a.SlCrDays,a.SlTaxPayAc,a.SlIsExempt,b.SlSubItCode,b.SlSubBag,1,b.SlSubRt,b.SlSubAmt,1,b.SlSubAmt, " _
        & " b.SlSubBookNo,b.SlSubBookDt,b.SlSubGodown,b.SlSubLotNo,a.SlTcsOnAmt,a.SlTcsRate,a.SlTcsAmt,a.SlTmpChr3 from tblSale A ,tblSaleSub B" _
        & " where a.vno = b.vno and a.vtype = b.vtype and a.vyear = b.vyear and a.vfirm = b.vfirm and  a.SlBillNo in " & X & " and a.Vtype='" & frmTypeStr & "' and a.VFirm = '" & gCCode & "' And a.VYear = " & gCYear
     
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
    Dim ItemNm As String
    Dim aFile As String
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
       ItemNm = GProcGetColumnValue("tblMastItem", "ItCode", RsSA.Fields("SlSubItCode"), "S", "ItName", "S")
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
'       Print #1, Space(15) + Chr(34) + "SupTyp" + Chr(34) + ":" + Chr(34) + "SEZWOP" + Chr(34) + Chr(44)
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
       Print #1, Space(32) + Chr(34) + "PrdDesc" + Chr(34) + ":" + Chr(34) + Trim(ItemNm) + Chr(34) + Chr(44)
       If Left(Trim(ItemHsn), 2) = "99" Then
          Print #1, Space(32) + Chr(34) + "IsServc" + Chr(34) + ":" + Chr(34) + "Y" + Chr(34) + Chr(44)
       Else
          Print #1, Space(32) + Chr(34) + "IsServc" + Chr(34) + ":" + Chr(34) + "N" + Chr(34) + Chr(44)
       End If
       Print #1, Space(32) + Chr(34) + "HsnCd" + Chr(34) + ":" + Chr(34) + Trim(ItemHsn) + Chr(34) + Chr(44)
       Print #1, Space(32) + Chr(34) + "Qty" + Chr(34) + ":" + Trim(CStr(Format(RsSA.Fields(50), "############0.000"))) + Chr(44)
'       Print #1, Space(32) + Chr(34) + "Unit" + Chr(34) + ":" + Chr(34) + "NOS" + Chr(34) + Chr(44)
       Print #1, Space(32) + Chr(34) + "Unit" + Chr(34) + ":" + Chr(34) + "KGS" + Chr(34) + Chr(44)
       
       Print #1, Space(32) + Chr(34) + "UnitPrice" + Chr(34) + ":" + Trim(CStr(Format(RsSA.Fields(52), "############0.000"))) + Chr(44)
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
'    txttemp(6) = GProcGetColumnValue("tblMastAccount", "AcCode", CStr(SalesAccode), "N", "AcName", "S")
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
X = "select SlSubItSrNo,ItName,SlSubAmt,SlSubItCtrlNo from tblSaleSub,TblMastItem " _
& "  where tblSaleSub.SlSubItCode=TblMastItem.ItCode" _
& " and tblSaleSub.vtype='" & frmTypeStr & "' and tblSaleSub.vno=" & txt(0) & " and tblSaleSub.VFirm = '" & gCCode & "' and tblSaleSub.VYear=" & gCYear & " order by SlSubItSrNo"
GProcRstOpen rstblSub, X, "R"
With msGrid
   .FormatString = ">Sr|<Description                                         |>Amount               |"
   .ColWidth(3) = 0
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
    If (dtpDt(39) = "01/04/1899" Or dtpDt(39) <> dtpDt(2)) And IsDate(dtpDt(2)) = True Then dtpDt(39) = dtpDt(2)
    If dtpDt(2) < CDate("01/07/2017") Then
       MsgBox "Date in VAT Applicabel " & txt(9), vbCritical + vbOKOnly, ""
       dtpDt(Index).SetFocus
       Exit Sub
    End If
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
GProcCheckForNumber Me, rstbl, MaxNo, Array(9), Array(66, 67), pOtherTxtFromNo:=OtherTxtFromNo, pOtherTxtToNo:=OtherTxtToNo
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
Private Sub txtCode_LostFocus(Index As Integer)
If Trim(txtCode(Index)) = "" Then Exit Sub
Dim BrkAcCode As Long
Dim tmpBillSr As String
Dim NarrCdBlTp As Long
Dim BilSr As String
SetControlEd
'If cmdBtn_Click(Index) <> 14 Then
Select Case Index
    Case 5  '-- Party  Code
         txttemp(5) = GProcGetColumnValue("tblMastAccount", "AcOurCode", UCase(txtCode(5)), "S", "AcName", "S")
         lblTinNo.Caption = GProcGetColumnValue("TblMastAccount", "ACName", txttemp(Index), "S", "AcGSTIN", "S")
         lblAdd1.Caption = GProcGetColumnValue("TblMastAccount", "ACName", txttemp(Index), "S", "AcAdd1", "S")
         lblAdd2.Caption = GProcGetColumnValue("TblMastAccount", "ACName", txttemp(Index), "S", "AcAdd2", "S")
         If Trim(txttemp(5)) = "" Then
            If txtCode(5) <> "D" Then
               MsgBox ("Code Not Found ...")
            End If
            txttemp(5) = ""
            txttemp(5).Enabled = True
            txttemp(5).SetFocus
         Else
            txttemp(5).SetFocus
         End If
End Select
If txt(9) = "" Then
'   GBillNoWithIB
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
            GProcShowForm frmMastAccount, frmMain.mnuMstfrm(0), Array(frmMastAccount.txttemp(2)), Array(GName)
        Case 6 '-- Sales A/c
            GName = GProcGetColumnValue("TblMastGroup", "AgCode", 90023, "N", "AgName", "S")
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
Dim ShowMill As String
Dim NarrCdBlTp As Long
Dim BilSr As String
TopPos = txttemp(Index).Top + Me.Top + 650
LeftPos = txttemp(Index).Left + Me.Left
If KeyAscii = 13 Then Exit Sub
Select Case Index
    Case 5  '-- Party   Alias Account
'        gClsSearch.SearchMultiField "QryMastAccount", "Account,AcOurCode,CGName,AcContPer,AcCity,AgName,AcBST,AcMsOmsParty", Array("Account Name", "Code", "Party Group", "Contact Person", "City", "Group", "TIN", "Ms/Oms"), Array(4000, 1000, 2000, 2000, 2000, 2500, 0, 0), "(Agcode in (90017) )", Chr(KeyAscii), "", LeftPos, TopPos, True
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
            If Val(lblMsOms.Caption) = 0 Then
               lblMsOms.Caption = "M.S."
            Else
               lblMsOms.Caption = "O.M.S."
            End If
            lblAdd1.Caption = GProcGetColumnValue("TblMastAccount", "ACName", txttemp(Index), "S", "AcAdd1", "S")
            lblAdd2.Caption = GProcGetColumnValue("TblMastAccount", "ACName", txttemp(Index), "S", "AcAdd2", "S")
        End If
    Case 6 '-- Sales A/c  Alias Account
'        gClsSearch.SearchMultiField "tblMastaccount", "AcName,AcCity", Array("Sales Account", "City"), Array(3000, 2000), "(Agcode=90024)", Chr(KeyAscii), "", LeftPos, TopPos, True
        gClsSearch.SearchMultiFieldAccMaster Chr(KeyAscii), LeftPos, TopPos
        KeyAscii = 0
        If gClsSearch.SearchMultiRetColAccMaster(0) <> "" Then
            txttemp(Index) = gClsSearch.SearchMultiRetColAccMaster(0)
        End If
'        If gClsSearch.SearchMultiRetCol(0) <> "" Then
'            txttemp(Index) = gClsSearch.SearchMultiRetCol(0)
'        End If
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
End Sub
Private Sub txtTemp_LostFocus(Index As Integer)
Dim AcCode As Long
Dim BilSr As String
Dim NarrCdBlTp As Long
SaveBtnEd
MillCtrlEd
ShowAccountBalance
If txt(9) = "" Then
   'GBillNoWithIB
End If
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
        If Val(.TextMatrix(.Row, 2)) = 0 And .rows > 2 Then
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
        If .Col = 3 Then
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
'   msGrid.Col = 1
'   If msGrid.Enabled = True Then
'      msGrid.SetFocus
'   End If
'   KeyAscii = 0
'   msGrid.SetFocus
   txtGrid.Move msGrid.CellLeft + msGrid.Left, msGrid.CellTop + msGrid.Top, msGrid.CellWidth, msGrid.CellHeight
End If
txtGrid.Visible = True
txtGrid.Enabled = True
DoEvents
If KeyAscii <> 0 Then
    Select Case rstblSub.Fields(msGrid.Col).Type
        Case adChar, adVarWChar
            If msGrid.Col = 1 Then  '-- Count
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
If msGrid.Col = 1 And txtGrid <> "" Then txtGrid.SelStart = 0
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
'Dim MillCode As Long
'Dim StdPack As Double
'Dim NetRt As Double
'Dim GAmt As Double
'Dim GSubAmt As Double
'Dim wCharityRt As Double
'Dim wCharityOn As Integer
'Dim wCharityAmt As Double
'Dim Rs1 As Recordset
'Dim VAddWithTax As Double
'Dim ItCd As Long
'Dim X As String
'Dim ExmptKg As Double
'Dim ExmptAmt As Double
'Dim CalWt As Double
'CheckForNumberGrid
'If msGrid.TextMatrix(msGrid.Row, 1) = 0 Then msGrid.TextMatrix(msGrid.Row, 2) = "" '-- Book date
'MillCode = GProcGetColumnValue("TblMastAccount", "AcName", txttemp(7), "S", "Accode", "N")
'StdPack = GProcGetColumnValue("TblMastItem", "ITName", msGrid.TextMatrix(msGrid.Row, 3), "S", "ItStdpack", "N", " ItMillCode=" & MillCode)
'wCharityRt = GProcGetColumnValue("TblMastItem", "ITName", msGrid.TextMatrix(msGrid.Row, 3), "S", "CharityRt", "N", " ItMillCode=" & MillCode)
'wCharityOn = GProcGetColumnValue("TblMastItem", "ITName", msGrid.TextMatrix(msGrid.Row, 3), "S", "CharityOn", "N", " ItMillCode=" & MillCode)
'ItCd = GProcGetColumnValue("tblMastItem", "ItName", msGrid.TextMatrix(msGrid.Row, 3), "S", "ItCode", "N", " ItMillCode=" & MillCode)
'
'
'If msGrid.Row = 0 Then Exit Sub
''-- Wt=Bag * Std Pack
'If OldBag <> msGrid.TextMatrix(msGrid.Row, 6) And StdPack > 0 Then
'    msGrid.TextMatrix(msGrid.Row, 8) = msGrid.TextMatrix(msGrid.Row, 6) * StdPack
'    OldBag = msGrid.TextMatrix(msGrid.Row, 6)
'End If
'
''If OldBag <> msGrid.TextMatrix(msGrid.Row, 6) And StdPack = 0 Then
''    msGrid.TextMatrix(msGrid.Row, 8) = 0
''End If
'If CDbl(msGrid.TextMatrix(msGrid.Row, 10)) > 0 Then
'    GAmt = Round(CDbl(msGrid.TextMatrix(msGrid.Row, 8)) * CDbl(msGrid.TextMatrix(msGrid.Row, 9)) / CDbl(msGrid.TextMatrix(msGrid.Row, 10)), 2)
'Else
'    GAmt = 0
'End If
'If frmTypeStr = "SE" And Val(msGrid.TextMatrix(msGrid.Row, 14)) > 0 Then
'   X = "select tblpurch.vno,tblpurch.vtype,tblpurch.vyear,tblpurch.vfirm,tblPurch.PurExemptPerKg from tblpurch,tblpurchsub where tblpurch.vno = tblpurchsub.vno and tblpurch.vtype = tblpurchsub.vtype and " _
'       & " tblpurch.vyear = tblpurchsub.vyear and tblpurch.vfirm = tblpurchsub.vfirm and tblpurch.vfirm = '" & gCCode & "' and tblpurchsub.PurSubItCtrlNo = '" & msGrid.TextMatrix(msGrid.Row, 14) & "'"
'   If GProcRstOpen(Rs1, X, "R") > 0 Then
'      ExmptKg = Rs1.Fields("PurExemptPerKg")
'      ExmptAmt = (ExmptKg * Val(msGrid.TextMatrix(msGrid.Row, 8)))
'   End If
'   VAddWithTax = GAmt - ExmptAmt
'    If VAddWithTax < 0 Then
'        VAddWithTax = 0
'        ExmptAmt = GAmt
'    End If
'    GSubAmt = Format(VAddWithTax - ((VAddWithTax * 100) / (100 + (Val(txt(23)) + Val(txt(26)) + Val(txt(41))))), LStr)
'    GSubAmt = Round(GAmt - GSubAmt, 2)
'Else
'   '--Sub Amt=(Amt)*100/(100+Tax %)
'   GSubAmt = Format(((GAmt * 100) / (100 + (Val(txt(23)) + Val(txt(26)) + Val(txt(41))))), LStr)
'End If
''- 8  - Weitht
''- 9  - Rate
''- 10  - Per
''- 11 - Net Rate
''- 12 - Sub Amt
'If wCharityRt > 0 Then
''And lblCharity.Visible = False Then
'   lblCharity.Visible = True
'   txt(17).Visible = True
'   txt(43).Visible = True
'   Cbo(44).Visible = True
'
'Else
'   lblCharity.Visible = False
'   txt(17).Visible = False
'   txt(43).Visible = False
'   Cbo(44).Visible = False
'End If
'
'If wCharityRt > 0 Then
'   If wCharityOn = 0 Then   '-- On Kg
'      wCharityAmt = Format(Val(msGrid.TextMatrix(msGrid.Row, 8)) * wCharityRt, FStr)
'   Else                     '-- On Pkg
'      wCharityAmt = Format(Val(msGrid.TextMatrix(msGrid.Row, 6)) * wCharityRt, FStr)
'   End If
'   txt(43) = wCharityRt
'   Cbo(44).ListIndex = wCharityOn
'Else
'   wCharityAmt = 0
'End If
'
'If frmTypeStr = "SD" And txtCode(7) = "ISCCGMW" Then
'   NetRt = Val(msGrid.TextMatrix(msGrid.Row, 9))
'Else
'    GSubAmt = GSubAmt - wCharityAmt
'    If CDbl(msGrid.TextMatrix(msGrid.Row, 8)) > 0 Then
'       NetRt = Format((GSubAmt / CDbl(msGrid.TextMatrix(msGrid.Row, 8))) * CDbl(msGrid.TextMatrix(msGrid.Row, 10)), LStr)
'    Else
'       NetRt = 0
'    End If
'End If
'
'msGrid.TextMatrix(msGrid.Row, 11) = NetRt
''-- SubAmt=Wt*NetRate/Rate Per
'If CDbl(msGrid.TextMatrix(msGrid.Row, 10)) > 0 Then
'    msGrid.TextMatrix(msGrid.Row, 12) = CDbl(msGrid.TextMatrix(msGrid.Row, 8)) * CDbl(msGrid.TextMatrix(msGrid.Row, 11)) / CDbl(msGrid.TextMatrix(msGrid.Row, 10))
'Else
'    msGrid.TextMatrix(msGrid.Row, 12) = 0
'End If
'msGrid.TextMatrix(msGrid.Row, 12) = Format(msGrid.TextMatrix(msGrid.Row, 12), LStr)
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
    Case 1 '-- Item
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
Dim QStr As String
TopPos = txtGrid.Top + Me.Top + 650
LeftPos = txtGrid.Left + Me.Left
DoEvents
'--- Check is Related Gate Pass
'If IsRelatedItCtrlGatePass(msGrid.TextMatrix(msGrid.Row, 14)) = True Then Exit Sub
With msGrid
    If KeyAscii <> 13 And KeyAscii <> 27 And KeyAscii <> 25 Then
        '--- Search
        Select Case .Col
            Case 1 '-- Count
                    gClsSearch.SearchMultiField "tblMastItem", "ItName,ItCode,ITMAXRATE,ITMINRATE,ItBrkRtOld", Array("Item Name", "ItCode", "CGST Rt.", "SGST Rt.", "IGST Rt."), Array(5000, 0, 1000, 1000, 1000), "tblMastItem.IsOtherItem=1", Chr(KeyAscii), "ItName", LeftPos, TopPos, True
                    txtGrid = ""
                    KeyAscii = 0
                    If gClsSearch.SearchMultiRetCol(0) <> "" Then '-- Item Wt ,Rate From Master
                        txtGrid = gClsSearch.SearchMultiRetCol(0)
                        If (Left(gCGStin, 2) = Left(lblTinNo.Caption, 2) Or Left(lblTinNo.Caption, 2) = "" Or Left(lblTinNo.Caption, 1) = "U") Then
                           txt(23) = gClsSearch.SearchMultiRetCol(2)
                           txt(26) = gClsSearch.SearchMultiRetCol(3)
                           txt(41) = 0
                        Else
                           txt(23) = 0
                           txt(26) = 0
                           txt(41) = gClsSearch.SearchMultiRetCol(4)
                        End If
                        txtGrid_KeyPress 13
                    End If
            Case Else
                KeyAscii = GProcValidateKey(rstblSub, msGrid.Col, KeyAscii, txtGrid)
        End Select
   End If
    If KeyAscii = 13 Then
        CheckForNumberTxtGrid
    End If
    '--- Navigate Row ,Col
    If KeyAscii = 13 And .Col < 2 Then 'rstblSub.Fields.Count - 2 Then
        .Col = .Col + 1
    ElseIf KeyAscii = 13 And .Col = 2 Then '= rstblSub.Fields.Count - 2 Then
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
        If (Val(.TextMatrix(.Row, 2)) = 0) And .rows > 2 Then
            .RemoveItem (.Row)
            txtGrid.text = ""
        End If
        txtGrid.Visible = False
        txtGrid.Visible = False
        txt(23).SetFocus
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
Dim X As String
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
                Case 1 '-- Count
                 '-Rate=0 ,count,wt=0 then delete row
                    If RowIndex > 1 And Trim(.TextMatrix(RowIndex, 1)) = "" And CDbl(.TextMatrix(RowIndex, 2)) = 0 Then
                        .RemoveItem (RowIndex)
                        GoTo Validate
                    End If
                    If Trim(.TextMatrix(RowIndex, i)) = "" Then
                        MsgBox "Check Item", vbCritical + vbOKOnly, Me.Caption
                        .Col = i
                        msGrid.SetFocus
                        Exit Function
                    End If
                Case 2 '-- amount
                    If CDbl(.TextMatrix(RowIndex, i)) = 0 Then
                        MsgBox "Check amount.", vbCritical + vbOKOnly, Me.Caption
                        .Col = i
                        msGrid.SetFocus
                        Exit Function
                    End If
           End Select
        Next i
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
If (CDbl(txt(0)) = 0 And FormAction <> vbDataActionAddNew) Or GProcIsDateValid(dtpDt(2)) = False Or txttemp(5) = "" Or txttemp(6) = "" Or CDbl(txt(16)) = 0 Or CDbl(txt(15)) = 0 Then
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
Dim HsnNm As String
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
If gCYear >= 2021 Then
   HsnNm = GProcGetColumnValue("tblMastItem", "ItName", msGrid.TextMatrix(1, 1), "S", "ItHsn2", "S", "IsOtherItem= 1")
   If Len(HsnNm) < 6 Then
      MsgBox "HSN Code Must be Minimum 6 Digit", vbInformation + vbOKOnly, Me.Caption
      txttemp(5).SetFocus
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
    '---Tax Type
    '-- Tax %
    
    '-- Tax Amount
    If (CDbl(txt(24)) > 0 Or CDbl(txt(23)) > 0) And Abs(Round(Val(txt(21)) * Val(txt(23)) / 100, 2) - Val(txt(24))) > 1 Then
        MsgBox "Tax Amount on Ass.Val not correct", vbInformation + vbOKOnly, Me.Caption
        txt(24).SetFocus
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
    '---- Amount
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
If IsDate(dtpDt(39)) Then '--G.P. Date
    txt(39) = CDate(dtpDt(39))
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
txt(5) = GProcGetColumnValue("tblMastAccount", "AcName", txttemp(5), "S", "AcCode", "N")
'-- Cr Code
txt(6) = GProcGetColumnValue("tblMastAccount", "AcName", txttemp(6), "S", "AcCode", "N")
'--Mill
txt(7) = GProcGetColumnValue("tblMastAccount", "AcName", "-", "S", "AcCode", "N")
'-- broker
txt(8) = GProcGetColumnValue("tblMastAccount", "AcName", "-", "S", "AcCode", "N")
'-- Transport
If txttemp(11) = "" Then
   txttemp(11) = "-"
End If
txt(11) = GProcGetColumnValue("tblMastNarration", "Narration", txttemp(11), "S", "NarrCode", "N")
txt(44) = Cbo(44).ListIndex
txt(3) = gCYear
txt(4) = gCCode
ShowAccountBalance
End Sub
Private Sub FillTempFromTxt()
lblTcsLimit.Caption = ""
Dim X As String
Dim MsOms As Integer
FillCombo
If IsDate(txt(2)) Then '--VDt
    dtpDt(2) = txt(2)
End If
If IsDate(txt(10)) Then '--Bill Date
    mskDt(10) = txt(10)
Else
    mskDt(10) = "__/__/____"
End If

If IsDate(txt(39)) Then '--G.P. Date
    dtpDt(39) = txt(39)
End If
If IsDate(txt(32)) Then '--Due Date
    dtpDt(32) = txt(32)
End If
If IsDate(txt(14)) Then '--LR Date
    mskDt(14) = txt(14)
Else
    mskDt(14) = "__/__/____"
End If
'-- Dr Code
txttemp(5) = GProcGetColumnValue("tblMastAccount", "AcCode", txt(5), "N", "AcName", "S")
txtCode(5) = GProcGetColumnValue("tblMastAccount", "AcCode", txt(5), "N", "AcOurCode", "S")
lblTinNo.Caption = GProcGetColumnValue("TblMastAccount", "ACName", txttemp(5), "S", "AcGSTIN", "S")
MsOms = GProcGetColumnValue("tblMastAccount", "AcCode", txt(5), "N", "AcMsOmsParty", "N")
If MsOms = 0 Then
   lblMsOms.Caption = "M.S."
Else
   lblMsOms.Caption = "O.M.S."
End If
lblAdd1.Caption = GProcGetColumnValue("TblMastAccount", "AcCode", txt(5), "N", "AcAdd1", "S")
lblAdd2.Caption = GProcGetColumnValue("TblMastAccount", "AcCode", txt(5), "N", "AcAdd2", "S")

'-- Cr Code
txttemp(6) = GProcGetColumnValue("tblMastAccount", "AcCode", txt(6), "N", "AcName", "S")

'--Transport
txttemp(11) = GProcGetColumnValue("tblMastNarration", "NarrCode", txt(11), "N", "Narration", "S")

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
            Next ColIndex
            RowIndex = RowIndex + 1
            .MoveNext
            If Not .EOF Then msGrid.rows = msGrid.rows + 1
        Loop
    End If
End With
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
Dim TaxCd As Long
Dim tmpRst As Recordset
Dim X As String
Dim TcsAmt As Double
Dim tmpRst2 As Recordset
Dim PtyOpBal As Double
Dim IsTdsParty As Integer
Dim TcsLimitAmt As Double
Dim TcsOnAmt As Double
Dim wPtyCd As Long
Dim PtyPan As String
wPtyCd = GProcGetColumnValue("TblMastAccount", "AcName", txttemp(5), "S", "AcCode", "N")
PtyPan = GProcGetColumnValue("tblMastAccount", "AcName", txttemp(5), "S", "AcPAN", "S")

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
        If Val(.TextMatrix(RowIndex, 2)) > 0 Then
            SubAmt = SubAmt + Val(.TextMatrix(RowIndex, 2))
        End If
        If RowIndex = 1 And msGrid.TextMatrix(1, 1) <> "" Then
           If (Left(gCGStin, 2) = Left(lblTinNo.Caption, 2) Or Left(lblTinNo.Caption, 2) = "" Or Left(lblTinNo.Caption, 1) = "U") Then
               txt(23) = GProcGetColumnValue("tblMastItem", "ItName", msGrid.TextMatrix(1, 1), "S", "ITMAXRATE", "N", "IsOtherItem= 1") '--Item Code
               txt(26) = GProcGetColumnValue("tblMastItem", "ItName", msGrid.TextMatrix(1, 1), "S", "ITMINRATE", "N", "IsOtherItem= 1") '--Item Code
               txt(41) = 0
           Else
               txt(23) = 0
               txt(26) = 0
               txt(41) = GProcGetColumnValue("tblMastItem", "ItName", msGrid.TextMatrix(1, 1), "S", "ItBrkRtOld", "N", "IsOtherItem= 1") '--Item Code
           End If
        End If
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

TaxAmt = AssValue * CDbl(txt(23)) / 100
TaxAmt = Format(TaxAmt, FStr)

If Abs(Val(txt(24)) - TaxAmt) < 1 Then
   TaxAmt = txt(24)
End If

'- 2 CGST
TaxAmtCGST = AssValue * CDbl(txt(26)) / 100
TaxAmtCGST = Format(TaxAmtCGST, FStr)
   
If Abs(Val(txt(27)) - TaxAmtCGST) < 1 Then
   TaxAmtCGST = txt(27)
End If
   
   
   
'- 3 IGST
TaxAmtIGST = AssValue * CDbl(txt(41)) / 100
TaxAmtIGST = Format(TaxAmtIGST, FStr)
If Abs(Val(txt(42)) - TaxAmtIGST) < 1 Then
   TaxAmtIGST = txt(42)
End If

Amt = Format(SubAmt + TaxAmt + TaxAmtCGST + TaxAmtIGST + CDbl(txt(17)) - CDbl(txt(18)) + CDbl(txt(28)) - CDbl(txt(29)), FStr)


'---TCS

txt(65) = Format(0, FStr)


'If Format(dtpDt(2).Value, "YYYY/MM/dd") >= Format("2021/07/01", "YYYY/MM/dd") Then
'    gCn.BeginTrans
'    IsTdsParty = GProcGetColumnValue("TblMastAccount", "AcName", txttemp(5), "S", "AcIsLessTDSOnRec", "N")
'    gCn.CommitTrans
'Else
'   IsTdsParty = 0
'End If
'--As per Munole's Phone On 02/04/2022 Tds Limit is applicable for Depot & Mill Bill
'If Format(dtpDt(2).Value, "YYYY/MM/dd") >= Format("2020/10/01", "YYYY/MM/dd") And (gCIsTcsFirm = 1 Or gCIsDepotFirm = 1) And IsTdsParty = 0 Then
'   X = "Select sum(SlBillAmt) from tblSale,tblMastAccount where tblSale.SlAcDrCode = tblMastAccount.AcCode and (SlAcDrCode = " & wPtyCd & " or tblMastAccount.AcPan = '" & PtyPan & "') and vfirm = '" & gCCode & "' and vtype in ('SY','SD','SC','SB') and vyear = " & gCYear & " and vdt <= '" & Format(dtpDt(2).Value, "YYYY/MM/dd") & "'  "
'   X = X + " and  ltrim(str(vno))+vtype+vFirm+ltrim(str(vyear)) <> '" & Trim(Str(txt(0))) & "" & gfrmTypeStr & "" & gCCode & "" & LTrim(Str(gCYear)) & "'"
'   i = GProcRstOpen(tmpRst, X, "R")
'   TcsLimitAmt = IIf(IsNull(tmpRst.Fields(0)), 0, tmpRst.Fields(0))
'   TcsLimitAmt = TcsLimitAmt + Amt
'
'   If GProcRstOpen(tmpRst2, "select Vamt from tblVoucher where vtype = '0P' and  VDrAcCode= " & wPtyCd & "  And Vyear = " & gCYear & " and vfirm ='" & gCCode & "'", "R") > 0 Then
'      TcsLimitAmt = TcsLimitAmt + IIf(IsNull(tmpRst2.Fields(0)), 0, tmpRst2.Fields(0))
'   End If
'   lblTcsLimit.Caption = "Op.+Sales up to Date = " + Str(TcsLimitAmt)
'
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
'   tmpRst.Close
'   tmpRst2.Close
'Else
'    TcsOnAmt = 0
'    TcsAmt = 0
'    txt(63) = Format(0, FStr)
'    txt(64) = Format(0, FStr)
'    txt(65) = Format(0, FStr)
'End If

If Abs(Val(txt(65)) - TcsAmt) < 1 Then
   TcsAmt = txt(65)
   txt(65) = TcsAmt
End If




'--Old Calculation before 31/03/2023

'If Val(txt(64)) > 0 Then
'   txt(63) = Format(Amt, FStr)
'   TcsAmt = (Amt * Val(txt(64))) / 100
'   If Format(dtpDt(2).Value, "YYYY/MM/dd") >= Format("2021/01/18", "YYYY/MM/dd") Then
'      TcsAmt = Format(Round(TcsAmt, 2), FStr)
'   Else
'      TcsAmt = Format(GProcMakeRounding(TcsAmt), FStr)
'   End If
'   If Abs(Val(txt(65)) - TcsAmt) < 1 Then
'       TcsAmt = txt(65)
'   End If
'   txt(65) = TcsAmt
'Else
'   txt(63) = Format(Amt, FStr)
'   txt(65) = 0
'End If

Amt = Amt + Val(txt(65)) + Val(txt(30))


'Amt2 = GProcMakeRounding(Amt)
'ROff = Format(Amt2 - Format(Amt, FStr), FStr)
txt(15) = SubAmt
txt(24) = TaxAmt
txt(27) = TaxAmtCGST
txt(42) = TaxAmtIGST
txt(21) = AssValue
'txt(30) = ROff
txt(16) = Round(Amt, 2)
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
   txttemp(5) = gClsSearch.SearchMultiRetCol(2) '--Party
   txttemp(8) = gClsSearch.SearchMultiRetCol(3) '--Broker
   txttemp(7) = gClsSearch.SearchMultiRetCol(4) '--Mill
   TmpGpVno = gClsSearch.SearchMultiRetCol(5) '--Gp Vno
   If FormAction = vbDataActionAddNew Then dtpDt(2) = CDate(dtpDt(39))   '-- Inv Dt
   wMillCode = GProcGetColumnValue("tblMastAccount", "AcName", txttemp(7), "S", "AcCode", "N")
   
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
Private Function GBillNoWithIB()
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
Case "SC"
   NarrCdBlTp = GProcGetColumnValue("tblMastNarration", "Narration", "Trade", "S", "NarrCode", "N", "NarrType = 'X'")
End Select
SalesAccode = GProcGetColumnValue("tblMastAccount", "AcName", txttemp(6), "S", "AcCode", "N")
If SalesAccode > 1 Then
    BilSr = ""
    BilSr = GProcGetColumnValue("tblMastBillSerial", "MillCode", "" & SalesAccode & "", "N", "BillSr", "S", "SaleType = " & NarrCdBlTp & " and VFirm = '" & gCCode & "' ")
    If BilSr = "" Then
       MsgBox ("Check Bill Sr not given...")
    Else
       txt(9) = GProcGenerateInvNoWithIB(frmTypeStr, SalesAccode, BilSr)  '--Inv No
    End If
End If
End Function

