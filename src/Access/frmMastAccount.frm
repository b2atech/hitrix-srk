VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form frmMastAccount 
   BackColor       =   &H00C2E3FC&
   BorderStyle     =   0  'None
   Caption         =   " Account Master"
   ClientHeight    =   8475
   ClientLeft      =   0
   ClientTop       =   -105
   ClientWidth     =   10455
   DrawMode        =   10  'Mask Pen
   FillColor       =   &H000040C0&
   FillStyle       =   0  'Solid
   Icon            =   "frmMastAccount.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   8475
   ScaleWidth      =   10455
   ShowInTaskbar   =   0   'False
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
      Left            =   2745
      TabIndex        =   31
      Text            =   "43"
      Top             =   7605
      Width           =   1005
   End
   Begin VB.TextBox txt 
      Height          =   360
      Index           =   34
      Left            =   -45
      TabIndex        =   102
      Text            =   "34"
      Top             =   615
      Visible         =   0   'False
      Width           =   300
   End
   Begin VB.ComboBox cboBalance 
      Appearance      =   0  'Flat
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
      ItemData        =   "frmMastAccount.frx":058A
      Left            =   4635
      List            =   "frmMastAccount.frx":0594
      Style           =   2  'Dropdown List
      TabIndex        =   10
      Top             =   2475
      Width           =   1050
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
      Index           =   33
      Left            =   15
      TabIndex        =   100
      Text            =   "33"
      Top             =   2460
      Visible         =   0   'False
      Width           =   420
   End
   Begin VB.TextBox txtOpBal 
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
      Left            =   3945
      TabIndex        =   11
      Top             =   2475
      Width           =   675
   End
   Begin VB.TextBox txtOpBal 
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
      Index           =   0
      Left            =   2490
      TabIndex        =   9
      Top             =   2475
      Width           =   1425
   End
   Begin VB.TextBox txtTemp 
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
      Left            =   2490
      MaxLength       =   40
      TabIndex        =   7
      Text            =   "txttemp(16)"
      ToolTipText     =   "Press F3 For Creation"
      Top             =   2010
      Width           =   5520
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
      Height          =   360
      Index           =   6
      Left            =   2490
      TabIndex        =   3
      Text            =   "6"
      Top             =   1170
      Width           =   5520
   End
   Begin VB.TextBox txt 
      Appearance      =   0  'Flat
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
      Index           =   32
      Left            =   9960
      MaxLength       =   40
      TabIndex        =   97
      Text            =   "32"
      Top             =   7320
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.TextBox txt 
      Height          =   360
      Index           =   29
      Left            =   9480
      TabIndex        =   96
      Text            =   "29"
      Top             =   7320
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.TextBox txt 
      Height          =   360
      Index           =   28
      Left            =   9120
      TabIndex        =   95
      Text            =   "28"
      Top             =   7320
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.TextBox txt 
      Height          =   360
      Index           =   27
      Left            =   8760
      TabIndex        =   94
      Text            =   "27"
      Top             =   7320
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
      Index           =   21
      Left            =   8520
      TabIndex        =   93
      Text            =   "21"
      Top             =   7320
      Visible         =   0   'False
      Width           =   255
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
      Index           =   17
      Left            =   0
      TabIndex        =   92
      Text            =   "17"
      Top             =   3960
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
      Height          =   375
      Index           =   16
      Left            =   0
      TabIndex        =   91
      Text            =   "16"
      Top             =   3360
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
      Height          =   375
      Index           =   20
      Left            =   8520
      TabIndex        =   90
      Text            =   "20"
      Top             =   6960
      Visible         =   0   'False
      Width           =   375
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   5280
      Left            =   225
      TabIndex        =   80
      Top             =   3015
      Width           =   8295
      _ExtentX        =   14631
      _ExtentY        =   9313
      _Version        =   393216
      Tabs            =   4
      TabsPerRow      =   4
      TabHeight       =   617
      BackColor       =   12640511
      ForeColor       =   128
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "Addr&ess"
      TabPicture(0)   =   "frmMastAccount.frx":05A0
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Shape3"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "lbl1(6)"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "lbl1(8)"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "lbl1(7)"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "lbl1(12)"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "lbl1(13)"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "lbl1(11)"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "lbl1(10)"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "lbl1(9)"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "lbl1(3)"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "lbl1(24)"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "txt(5)"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).Control(12)=   "txt(4)"
      Tab(0).Control(12).Enabled=   0   'False
      Tab(0).Control(13)=   "cbo(7)"
      Tab(0).Control(13).Enabled=   0   'False
      Tab(0).Control(14)=   "cbo(9)"
      Tab(0).Control(14).Enabled=   0   'False
      Tab(0).Control(15)=   "txt(10)"
      Tab(0).Control(15).Enabled=   0   'False
      Tab(0).Control(16)=   "cbo(8)"
      Tab(0).Control(16).Enabled=   0   'False
      Tab(0).Control(17)=   "txt(23)"
      Tab(0).Control(17).Enabled=   0   'False
      Tab(0).Control(18)=   "txt(12)"
      Tab(0).Control(18).Enabled=   0   'False
      Tab(0).Control(19)=   "txt(11)"
      Tab(0).Control(19).Enabled=   0   'False
      Tab(0).Control(20)=   "txt(36)"
      Tab(0).Control(20).Enabled=   0   'False
      Tab(0).Control(21)=   "Chk(27)"
      Tab(0).Control(21).Enabled=   0   'False
      Tab(0).ControlCount=   22
      TabCaption(1)   =   "Personal Deta&ils"
      TabPicture(1)   =   "frmMastAccount.frx":05BC
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Chk(34)"
      Tab(1).Control(1)=   "Chk(28)"
      Tab(1).Control(2)=   "Chk(26)"
      Tab(1).Control(3)=   "Chk(29)"
      Tab(1).Control(4)=   "cbo(33)"
      Tab(1).Control(5)=   "txt(18)"
      Tab(1).Control(6)=   "txt(19)"
      Tab(1).Control(7)=   "txtTemp(3)"
      Tab(1).Control(8)=   "txtTemp(17)"
      Tab(1).Control(9)=   "Chk(21)"
      Tab(1).Control(10)=   "Chk(20)"
      Tab(1).Control(11)=   "Chk(32)"
      Tab(1).Control(12)=   "txt(26)"
      Tab(1).Control(13)=   "txt(31)"
      Tab(1).Control(14)=   "txt(30)"
      Tab(1).Control(15)=   "lbl1(23)"
      Tab(1).Control(16)=   "lblBillSrNo"
      Tab(1).Control(17)=   "lbl1(5)"
      Tab(1).Control(18)=   "lbl1(0)"
      Tab(1).Control(19)=   "lbl1(4)"
      Tab(1).Control(20)=   "lblOSLimit"
      Tab(1).Control(21)=   "lbl1(1)"
      Tab(1).Control(22)=   "Shape6"
      Tab(1).ControlCount=   23
      TabCaption(2)   =   "&Govt. Tax Nos"
      TabPicture(2)   =   "frmMastAccount.frx":05D8
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "txt(41)"
      Tab(2).Control(1)=   "txt(42)"
      Tab(2).Control(2)=   "cbo(35)"
      Tab(2).Control(3)=   "txt(35)"
      Tab(2).Control(4)=   "txt(15)"
      Tab(2).Control(5)=   "txt(13)"
      Tab(2).Control(6)=   "txt(14)"
      Tab(2).Control(7)=   "txt(22)"
      Tab(2).Control(8)=   "cbo(24)"
      Tab(2).Control(9)=   "txt(25)"
      Tab(2).Control(10)=   "lbl1(25)"
      Tab(2).Control(11)=   "lbl1(26)"
      Tab(2).Control(12)=   "lbl1(2)"
      Tab(2).Control(13)=   "lbl1(15)"
      Tab(2).Control(14)=   "lbl1(16)"
      Tab(2).Control(15)=   "lbl1(14)"
      Tab(2).Control(16)=   "lbl1(17)"
      Tab(2).Control(17)=   "lbl1(18)"
      Tab(2).Control(18)=   "lbl1(19)"
      Tab(2).Control(19)=   "Shape5"
      Tab(2).ControlCount=   20
      TabCaption(3)   =   "RTGS Details"
      TabPicture(3)   =   "frmMastAccount.frx":05F4
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "Shape8"
      Tab(3).Control(1)=   "Label15(2)"
      Tab(3).Control(2)=   "Label15(3)"
      Tab(3).Control(3)=   "Label15(4)"
      Tab(3).Control(4)=   "Label15(5)"
      Tab(3).Control(5)=   "txt(37)"
      Tab(3).Control(6)=   "txt(38)"
      Tab(3).Control(7)=   "txt(39)"
      Tab(3).Control(8)=   "txt(40)"
      Tab(3).ControlCount=   9
      Begin VB.CheckBox Chk 
         BackColor       =   &H00ECFFFF&
         Caption         =   "Is TDS Deducted Party ?"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   375
         Index           =   27
         Left            =   4860
         TabIndex        =   116
         Top             =   4725
         Width           =   3165
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
         Index           =   41
         Left            =   -72435
         MaxLength       =   20
         TabIndex        =   54
         Text            =   "41"
         Top             =   1395
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
         Height          =   360
         Index           =   42
         Left            =   -72435
         MaxLength       =   5
         TabIndex        =   55
         Text            =   "42"
         Top             =   1800
         Width           =   645
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
         Index           =   40
         Left            =   -72975
         TabIndex        =   107
         Text            =   "40"
         Top             =   1605
         Width           =   5760
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
         Index           =   39
         Left            =   -72975
         TabIndex        =   109
         Text            =   "39"
         Top             =   2160
         Width           =   5760
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
         Index           =   38
         Left            =   -72975
         TabIndex        =   108
         Text            =   "38"
         Top             =   2685
         Width           =   5760
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
         Left            =   -72975
         TabIndex        =   106
         Text            =   "37"
         Top             =   975
         Width           =   5760
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
         Index           =   36
         Left            =   2520
         TabIndex        =   28
         Text            =   "36"
         Top             =   3705
         Width           =   4995
      End
      Begin VB.ComboBox cbo 
         Appearance      =   0  'Flat
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
         ItemData        =   "frmMastAccount.frx":0610
         Left            =   -72435
         List            =   "frmMastAccount.frx":061D
         Style           =   2  'Dropdown List
         TabIndex        =   104
         Top             =   3555
         Width           =   2295
      End
      Begin VB.TextBox txt 
         Height          =   360
         Index           =   35
         Left            =   -75000
         TabIndex        =   103
         Text            =   "35"
         Top             =   345
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.CheckBox Chk 
         BackColor       =   &H00ECFFFF&
         Caption         =   "Is DbNT Account ?"
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
         Index           =   34
         Left            =   -69750
         TabIndex        =   101
         Top             =   3795
         Width           =   2595
      End
      Begin VB.CheckBox Chk 
         BackColor       =   &H00ECFFFF&
         Caption         =   "Is Service Tax Pay ?"
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
         Index           =   28
         Left            =   -69750
         TabIndex        =   47
         Top             =   3090
         Width           =   2415
      End
      Begin VB.CheckBox Chk 
         BackColor       =   &H80000018&
         Caption         =   "Is Local Brokerage By Self ?"
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
         Index           =   26
         Left            =   -69750
         TabIndex        =   46
         Top             =   2430
         Width           =   2775
      End
      Begin VB.CheckBox Chk 
         BackColor       =   &H00ECFFFF&
         Caption         =   "Is Exemption Mill ?"
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
         Index           =   29
         Left            =   -69750
         TabIndex        =   48
         Top             =   3450
         Width           =   2055
      End
      Begin VB.ComboBox cbo 
         Appearance      =   0  'Flat
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
         ItemData        =   "frmMastAccount.frx":064D
         Left            =   -69465
         List            =   "frmMastAccount.frx":0663
         Style           =   2  'Dropdown List
         TabIndex        =   45
         Top             =   1995
         Width           =   2010
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
         Index           =   15
         Left            =   -72435
         TabIndex        =   53
         Text            =   "15"
         Top             =   960
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
         Height          =   360
         Index           =   13
         Left            =   -72435
         TabIndex        =   57
         Text            =   "13"
         Top             =   2160
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
         Height          =   360
         Index           =   14
         Left            =   -72435
         TabIndex        =   59
         Text            =   "14"
         Top             =   2640
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
         Height          =   375
         Index           =   22
         Left            =   -72435
         TabIndex        =   61
         Text            =   "22"
         Top             =   3120
         Width           =   4335
      End
      Begin VB.ComboBox cbo 
         Appearance      =   0  'Flat
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
         ItemData        =   "frmMastAccount.frx":06A4
         Left            =   -72435
         List            =   "frmMastAccount.frx":06B4
         Style           =   2  'Dropdown List
         TabIndex        =   63
         Top             =   3975
         Width           =   2295
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
         Left            =   -69435
         TabIndex        =   65
         Text            =   "25"
         Top             =   3975
         Width           =   1335
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
         Index           =   18
         Left            =   -73095
         TabIndex        =   37
         Text            =   "18"
         Top             =   1995
         Width           =   2295
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
         Index           =   19
         Left            =   -73095
         TabIndex        =   41
         Text            =   "19"
         Top             =   2955
         Width           =   2295
      End
      Begin VB.TextBox txtTemp 
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
         Index           =   3
         Left            =   -73095
         MaxLength       =   40
         TabIndex        =   35
         Text            =   "txttemp(3)"
         ToolTipText     =   "Press F3 For Creation"
         Top             =   1515
         Width           =   6105
      End
      Begin VB.TextBox txtTemp 
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
         Index           =   17
         Left            =   -73095
         MaxLength       =   40
         TabIndex        =   33
         Text            =   "txttemp(17)"
         Top             =   1035
         Width           =   6105
      End
      Begin VB.CheckBox Chk 
         BackColor       =   &H00ECFFFF&
         Caption         =   "Don't Show Outstanding ?"
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
         Index           =   21
         Left            =   -70110
         TabIndex        =   51
         Top             =   4275
         Width           =   2595
      End
      Begin VB.CheckBox Chk 
         BackColor       =   &H00ECFFFF&
         Caption         =   "Is Non Work ?"
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
         Index           =   20
         Left            =   -73665
         TabIndex        =   49
         Top             =   4245
         Width           =   1545
      End
      Begin VB.CheckBox Chk 
         BackColor       =   &H00ECFFFF&
         Caption         =   "Is Ledger Total ?"
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
         Index           =   32
         Left            =   -71970
         TabIndex        =   50
         Top             =   4260
         Width           =   1770
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
         Index           =   26
         Left            =   -71745
         TabIndex        =   98
         Text            =   "26"
         Top             =   4830
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
         Height          =   375
         Index           =   31
         Left            =   -73095
         TabIndex        =   39
         Text            =   "31"
         Top             =   2475
         Width           =   2295
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
         Index           =   30
         Left            =   -73095
         TabIndex        =   43
         Text            =   "30"
         Top             =   3435
         Width           =   2295
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
         Index           =   11
         Left            =   2520
         TabIndex        =   24
         Text            =   "11"
         Top             =   2805
         Width           =   4995
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
         Index           =   12
         Left            =   2520
         TabIndex        =   26
         Text            =   "12"
         Top             =   3255
         Width           =   4995
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
         Index           =   23
         Left            =   2520
         TabIndex        =   30
         Text            =   "23"
         Top             =   4155
         Width           =   4995
      End
      Begin VB.ComboBox cbo 
         CausesValidation=   0   'False
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
         Left            =   5160
         TabIndex        =   22
         Text            =   "cbo(8)"
         Top             =   2310
         Width           =   2355
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
         Left            =   1755
         TabIndex        =   20
         Text            =   "10"
         Top             =   2310
         Width           =   2340
      End
      Begin VB.ComboBox cbo 
         CausesValidation=   0   'False
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
         Left            =   5160
         TabIndex        =   18
         Text            =   "cbo(9)"
         Top             =   1875
         Width           =   2355
      End
      Begin VB.ComboBox cbo 
         CausesValidation=   0   'False
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
         Left            =   1755
         TabIndex        =   16
         Text            =   "cbo(7)"
         Top             =   1875
         Width           =   2355
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
         Index           =   4
         Left            =   1755
         TabIndex        =   13
         Text            =   "4"
         Top             =   960
         Width           =   5760
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
         Index           =   5
         Left            =   1755
         MaxLength       =   40
         TabIndex        =   14
         Text            =   "5"
         Top             =   1365
         Width           =   5760
      End
      Begin VB.Label lbl1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Distance In KM   :"
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
         Height          =   240
         Index           =   24
         Left            =   630
         TabIndex        =   117
         Top             =   4635
         Width           =   1575
      End
      Begin VB.Label lbl1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "GSTIN   No  :"
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
         Index           =   25
         Left            =   -74235
         TabIndex        =   115
         Top             =   1395
         Width           =   1215
      End
      Begin VB.Label lbl1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "State Code  :"
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
         Index           =   26
         Left            =   -74235
         TabIndex        =   114
         Top             =   1800
         Width           =   1230
      End
      Begin VB.Label Label15 
         BackStyle       =   0  'Transparent
         Caption         =   "IFSC Code     :"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00645285&
         Height          =   255
         Index           =   5
         Left            =   -74505
         TabIndex        =   113
         Top             =   2250
         Width           =   1380
      End
      Begin VB.Label Label15 
         BackStyle       =   0  'Transparent
         Caption         =   "Branch   :"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00645285&
         Height          =   255
         Index           =   4
         Left            =   -74505
         TabIndex        =   112
         Top             =   2745
         Width           =   1425
      End
      Begin VB.Label Label15 
         BackStyle       =   0  'Transparent
         Caption         =   "A/c No          :"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00645285&
         Height          =   255
         Index           =   3
         Left            =   -74505
         TabIndex        =   111
         Top             =   1665
         Width           =   1290
      End
      Begin VB.Label Label15 
         BackStyle       =   0  'Transparent
         Caption         =   "Bank Name    :"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00645285&
         Height          =   255
         Index           =   2
         Left            =   -74505
         TabIndex        =   110
         Top             =   1035
         Width           =   1380
      End
      Begin VB.Label lbl1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Mbl No For SMS  :             :"
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
         Height          =   240
         Index           =   3
         Left            =   600
         TabIndex        =   27
         Top             =   3705
         Width           =   2475
      End
      Begin VB.Label lbl1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Commission Type :"
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
         Height          =   240
         Index           =   2
         Left            =   -74190
         TabIndex        =   105
         Top             =   3555
         Width           =   1680
      End
      Begin VB.Label lbl1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Mill Type   :"
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
         Height          =   240
         Index           =   23
         Left            =   -70680
         TabIndex        =   44
         Top             =   1995
         Width           =   1020
      End
      Begin VB.Label lbl1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "VAT TIN      :"
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
         Height          =   240
         Index           =   15
         Left            =   -74190
         TabIndex        =   56
         Top             =   2160
         Width           =   1140
      End
      Begin VB.Label lbl1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "CST TIN      :"
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
         Height          =   240
         Index           =   16
         Left            =   -74190
         TabIndex        =   58
         Top             =   2640
         Width           =   1140
      End
      Begin VB.Label lbl1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "PAN No.     :"
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
         Height          =   360
         Index           =   14
         Left            =   -74190
         TabIndex        =   52
         Top             =   960
         Width           =   1275
      End
      Begin VB.Label lbl1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "TAN           :"
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
         Height          =   240
         Index           =   17
         Left            =   -74190
         TabIndex        =   60
         Top             =   3120
         Width           =   1095
      End
      Begin VB.Label lbl1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Commission On  :"
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
         Height          =   240
         Index           =   18
         Left            =   -74190
         TabIndex        =   62
         Top             =   3975
         Width           =   1575
      End
      Begin VB.Label lbl1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "%    :"
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
         Index           =   19
         Left            =   -70035
         TabIndex        =   64
         Top             =   3975
         Width           =   480
      End
      Begin VB.Label lblBillSrNo 
         BackStyle       =   0  'Transparent
         Caption         =   "Bill Sr No.        :"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404080&
         Height          =   375
         Left            =   -74655
         TabIndex        =   36
         Top             =   1995
         Width           =   1590
      End
      Begin VB.Label lbl1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Broker Name    :"
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
         Height          =   240
         Index           =   5
         Left            =   -74655
         TabIndex        =   32
         Top             =   1035
         Width           =   1440
      End
      Begin VB.Label lbl1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Interest %        :"
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
         Left            =   -74655
         TabIndex        =   40
         Top             =   2955
         Width           =   1590
      End
      Begin VB.Label lbl1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Second Group  :"
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
         Index           =   4
         Left            =   -74655
         TabIndex        =   34
         Top             =   1515
         Width           =   1425
      End
      Begin VB.Label lblOSLimit 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "O/S. Limit        :"
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
         Height          =   240
         Left            =   -74655
         TabIndex        =   38
         Top             =   2475
         Width           =   1425
      End
      Begin VB.Label lbl1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Exemption %   :"
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
         Height          =   240
         Index           =   1
         Left            =   -74655
         TabIndex        =   42
         Top             =   3435
         Width           =   1410
      End
      Begin VB.Label lbl1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Contact Person    :"
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
         Index           =   9
         Left            =   600
         TabIndex        =   23
         Top             =   2805
         Width           =   1650
      End
      Begin VB.Label lbl1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Contact No.         :"
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
         Index           =   10
         Left            =   600
         TabIndex        =   25
         Top             =   3255
         Width           =   1635
      End
      Begin VB.Label lbl1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Email                  :"
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
         Height          =   240
         Index           =   11
         Left            =   600
         TabIndex        =   29
         Top             =   4155
         Width           =   1635
      End
      Begin VB.Label lbl1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Pin Code   :"
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
         Height          =   240
         Index           =   13
         Left            =   600
         TabIndex        =   19
         Top             =   2310
         Width           =   1035
      End
      Begin VB.Label lbl1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "State   :"
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
         Height          =   240
         Index           =   12
         Left            =   4350
         TabIndex        =   21
         Top             =   2310
         Width           =   705
      End
      Begin VB.Label lbl1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "City          :"
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
         Height          =   240
         Index           =   7
         Left            =   600
         TabIndex        =   15
         Top             =   1875
         Width           =   1005
      End
      Begin VB.Label lbl1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Dist     :"
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
         Height          =   240
         Index           =   8
         Left            =   4350
         TabIndex        =   17
         Top             =   1875
         Width           =   705
      End
      Begin VB.Label lbl1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Address    : "
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
         Height          =   240
         Index           =   6
         Left            =   600
         TabIndex        =   12
         Top             =   960
         Width           =   1080
      End
      Begin VB.Shape Shape3 
         BorderColor     =   &H00000080&
         FillColor       =   &H00EDFDFE&
         FillStyle       =   0  'Solid
         Height          =   4350
         Left            =   240
         Top             =   795
         Width           =   7815
      End
      Begin VB.Shape Shape6 
         FillColor       =   &H00EDFDFE&
         FillStyle       =   0  'Solid
         Height          =   5055
         Left            =   -74865
         Top             =   855
         Width           =   8055
      End
      Begin VB.Shape Shape5 
         FillColor       =   &H00EDFDFE&
         FillStyle       =   0  'Solid
         Height          =   3555
         Left            =   -74715
         Top             =   870
         Width           =   7785
      End
      Begin VB.Shape Shape8 
         BorderColor     =   &H00000080&
         FillColor       =   &H00EDFDFE&
         FillStyle       =   0  'Solid
         Height          =   3675
         Left            =   -74820
         Top             =   495
         Width           =   7815
      End
   End
   Begin VB.CommandButton cmdBtn 
      Caption         =   "Creation &To Firm"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   705
      Index           =   12
      Left            =   9000
      Style           =   1  'Graphical
      TabIndex        =   81
      Top             =   3435
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
      Index           =   11
      Left            =   9000
      Style           =   1  'Graphical
      TabIndex        =   78
      Top             =   2460
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
      Left            =   9000
      Style           =   1  'Graphical
      TabIndex        =   79
      ToolTipText     =   "Exit"
      Top             =   6630
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
      Left            =   9000
      Style           =   1  'Graphical
      TabIndex        =   70
      ToolTipText     =   "To Cancel Record"
      Top             =   6165
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
      Left            =   9000
      Style           =   1  'Graphical
      TabIndex        =   69
      ToolTipText     =   "To Save Record"
      Top             =   5730
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
      Left            =   9000
      Style           =   1  'Graphical
      TabIndex        =   71
      Top             =   2940
      Visible         =   0   'False
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
      Left            =   9000
      Style           =   1  'Graphical
      TabIndex        =   68
      ToolTipText     =   "To Delete Record"
      Top             =   5175
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
      Height          =   405
      Index           =   7
      Left            =   9000
      Style           =   1  'Graphical
      TabIndex        =   67
      ToolTipText     =   "To Modify Record"
      Top             =   4710
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
      Left            =   9000
      Style           =   1  'Graphical
      TabIndex        =   66
      ToolTipText     =   "To Add New Record"
      Top             =   4245
      Width           =   1095
   End
   Begin VB.CommandButton cmdBtn 
      Caption         =   "M&ill List"
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
      Left            =   9000
      Style           =   1  'Graphical
      TabIndex        =   77
      ToolTipText     =   "Show Mill List"
      Top             =   2040
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
      Left            =   9000
      Style           =   1  'Graphical
      TabIndex        =   72
      ToolTipText     =   "To Find Record"
      Top             =   1605
      Width           =   1095
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
      Left            =   9600
      Style           =   1  'Graphical
      TabIndex        =   76
      ToolTipText     =   "Last"
      Top             =   1155
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
      Left            =   9000
      Style           =   1  'Graphical
      TabIndex        =   75
      ToolTipText     =   "First"
      Top             =   1155
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
      Left            =   9600
      Style           =   1  'Graphical
      TabIndex        =   74
      ToolTipText     =   "Next"
      Top             =   675
      Width           =   495
   End
   Begin VB.TextBox txtTemp 
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
      Left            =   2490
      MaxLength       =   40
      TabIndex        =   5
      Text            =   "txttemp(2)"
      ToolTipText     =   "Press F3 For Creation"
      Top             =   1560
      Width           =   5520
   End
   Begin VB.TextBox txt 
      Height          =   360
      Index           =   24
      Left            =   0
      TabIndex        =   88
      Text            =   "24"
      Top             =   6000
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.TextBox txt 
      Height          =   360
      Index           =   9
      Left            =   0
      TabIndex        =   87
      Text            =   "9"
      Top             =   5520
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.TextBox txt 
      Height          =   360
      Index           =   8
      Left            =   0
      TabIndex        =   86
      Text            =   "8"
      Top             =   5040
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.TextBox txt 
      Height          =   360
      Index           =   7
      Left            =   1320
      TabIndex        =   85
      Text            =   "7"
      Top             =   5760
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.TextBox txt 
      Height          =   360
      Index           =   3
      Left            =   0
      TabIndex        =   84
      Text            =   "3"
      Top             =   4440
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.TextBox txt 
      Height          =   360
      Index           =   2
      Left            =   0
      TabIndex        =   83
      Text            =   "2"
      Top             =   1440
      Visible         =   0   'False
      Width           =   375
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
      Height          =   360
      Index           =   1
      Left            =   2490
      TabIndex        =   1
      Text            =   "1"
      Top             =   750
      Width           =   5520
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
      Left            =   9000
      Style           =   1  'Graphical
      TabIndex        =   73
      ToolTipText     =   "Previous"
      Top             =   675
      Width           =   495
   End
   Begin VB.TextBox txt 
      Appearance      =   0  'Flat
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
      Left            =   0
      MaxLength       =   40
      TabIndex        =   82
      Text            =   "0"
      Top             =   1080
      Visible         =   0   'False
      Width           =   375
   End
   Begin Crystal.CrystalReport CryReport 
      Left            =   0
      Top             =   0
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
   Begin VB.Shape Shape7 
      BorderColor     =   &H00C00000&
      BorderWidth     =   5
      FillColor       =   &H00C00000&
      Height          =   8400
      Left            =   0
      Top             =   15
      Width           =   10455
   End
   Begin VB.Label lblActionStatus 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "View Record.."
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   495
      Left            =   7875
      TabIndex        =   99
      Top             =   75
      Width           =   2475
   End
   Begin VB.Label lbl1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "&Opening Balance :"
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
      Height          =   240
      Index           =   20
      Left            =   690
      TabIndex        =   8
      Top             =   2475
      Width           =   1605
   End
   Begin VB.Label lbl1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Party Group        :"
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
      Height          =   240
      Index           =   21
      Left            =   690
      TabIndex        =   6
      Top             =   2010
      Width           =   1590
   End
   Begin VB.Label Label15 
      BackStyle       =   0  'Transparent
      Caption         =   "Alias Name         :"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00645285&
      Height          =   255
      Index           =   1
      Left            =   690
      TabIndex        =   2
      Top             =   1170
      Width           =   1605
   End
   Begin VB.Label Label15 
      BackStyle       =   0  'Transparent
      Caption         =   "Account &Name   :"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00645285&
      Height          =   360
      Index           =   0
      Left            =   690
      TabIndex        =   0
      Top             =   750
      Width           =   1665
   End
   Begin VB.Label Label6 
      BackColor       =   &H00C00000&
      Caption         =   "   Account Master"
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
      Height          =   390
      Left            =   0
      TabIndex        =   89
      Top             =   0
      Width           =   10575
   End
   Begin VB.Shape Shape1 
      FillColor       =   &H00BDDFD1&
      FillStyle       =   0  'Solid
      Height          =   6615
      Left            =   8760
      Shape           =   4  'Rounded Rectangle
      Top             =   600
      Width           =   1575
   End
   Begin VB.Label lbl1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "&Group Name      :"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H004D177D&
      Height          =   240
      Index           =   22
      Left            =   690
      TabIndex        =   4
      Top             =   1560
      Width           =   1620
   End
   Begin VB.Shape Shape2 
      BackColor       =   &H00808080&
      BorderColor     =   &H004D177D&
      FillColor       =   &H00EDFDFE&
      FillStyle       =   0  'Solid
      Height          =   2355
      Left            =   240
      Top             =   600
      Width           =   8295
   End
End
Attribute VB_Name = "frmMastAccount"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Const tblName As String = "tblMastAccount"
Const IdField As String = "AcCode"
Const OrderField As String = "AcName"
Const CriteriaStr As String = ""
Const MaxNo As Long = 43
Dim rstbl As Recordset
Dim varBookMark As Variant
Dim FirstTimeFlag As Boolean
Dim FormAction As Integer

Private Sub Cbo_KeyPress(Index As Integer, KeyAscii As Integer)
If KeyAscii = 13 Then Exit Sub
'-- Commission On
If Index = 24 Or Index = 35 Then KeyAscii = 0
End Sub

Private Sub cbo_LostFocus(Index As Integer)
Dim CityState As String
Dim CityDist As String
Dim CityZip As String
If Index = 7 Or Index = 8 Or Index = 9 Then
    cbo(Index).text = GProcProperCase(Trim(cbo(Index).text))
    If Index = 7 Then
       CityState = GProcGetColumnValue("TblMastAccount", "AcCity", cbo(7).text, "S", "AcState", "S")
       cbo(8).text = CityState
    End If
    If Index = 7 Then
       CityDist = GProcGetColumnValue("TblMastAccount", "AcCity", cbo(7).text, "S", "AcDist", "S")
       cbo(9).text = CityDist
    End If
    If Index = 7 Then
       CityZip = GProcGetColumnValue("TblMastAccount", "AcCity", cbo(7).text, "S", "AcZip", "S")
       txt(10).text = CityZip
    End If

End If
''If Index = 7 Or Index = 8 Or Index = 9 Then
''    cbo(Index).Text = GProcProperCase(Trim(cbo(Index).Text))
''End If
End Sub
Private Sub cboBalance_LostFocus()
If cboBalance.ListIndex = 0 Then
    txtOpBal(1) = "Cr"
Else
    txtOpBal(1) = "Dr"
End If
End Sub
Private Sub cmdBtn_Click(Index As Integer)
On Error GoTo ErrorRoutine
Select Case Index
Case 0 'Previous
    rstbl.MovePrevious
    If rstbl.BOF = True Then
        rstbl.MoveFirst
        MsgBox ("This is First Record"), vbInformation + vbOKOnly, "First Record"
    End If
    ReadFields
Case 1 'Next
    rstbl.MoveNext
    If rstbl.EOF = True Then
        rstbl.MoveLast
        MsgBox ("This is Last Record"), vbInformation + vbOKOnly, "Last Record"
    End If
    ReadFields
Case 2 'First
    rstbl.MoveFirst
    ReadFields
Case 3 'Last
    rstbl.MoveLast
    ReadFields
Case 4 'Find ' QryMastAccount
    gClsSearch.SearchMultiField "QryMastAccount", "Account,AcName,AcCity,AgName,AcMillTypeStr", Array("Account Name", "AcName", "City", "Group", "Mill Type"), Array(3000, 0, 1200, 2800, 2000), "", "", "Account", txt(1).Left + Me.Left, txt(1).Top + Me.Top + 650
   ' gClsSearch.SearchMultiSelField "QryMastAccount", "Account,AcName,AcCity,AgName,AcMillTypeStr", Array("Account Name", "AcName", "City", "Group", "Mill Type"), Array(3000, 0, 1200, 2800, 2000), "", "", "Account", txt(1).Left + Me.Left, txt(1).Top + Me.Top + 650
     If gClsSearch.SearchMultiRetCol(1) <> "" Then
        rstbl.CancelUpdate
        rstbl.MoveFirst
        rstbl.Find "AcName = '" & gClsSearch.SearchMultiRetCol(1) & "'"
        gClsSearch.SearchMultiRetCol(1) = ""
        ReadFields
    End If
Case 5 'Mill A/c List
    On Error GoTo ErrorHandler '--- Outstanding List
    gSelectedMenu = "Mill Account List"
    GProcCrystalRptPreparation CryReport, "Window"
    CryReport.Formulas(4) = ""
    CryReport.Formulas(5) = ""
   ' CryReport.SelectionFormula = "{tblOutStanding.VYear} = " & gCYear - 1 & " and ({tblOutStanding.Vtype}='SO' or {tblOutStanding.Vtype}='SY' or {tblOutStanding.Vtype}='ST'  or   {tblOutStanding.Vtype}='SD' or  {tblOutStanding.Vtype}='SM')"
    CryReport.ReportFileName = gReportPath & "RptLstMillAccount.rpt"
    CryReport.Action = 1
    Exit Sub
ErrorHandler:
        GProcErrorHandler
    'GenMastList ("LAC")
Case 6 'Add
    FormAction = vbDataActionAddNew
    Call GProcActivateControls(Me, True)
    Call GProcSetButtons(Me, vbDataActionAddNew)
    Call GProcClearForm(Me, rstbl, MaxNo, True)
    SetCtrlEd
    If FirstTimeFlag = True Then
        FirstTimeFlag = False
    Else
        txt(1).SetFocus
    End If
    FillCombo
    cbo(7) = "" '-City
    cbo(8) = "" '- State
    cbo(9) = "" '-Dist
Case 7 'Modify
    'If mURecModi = True Then
    '    SearchHelp
    '    DoEvents
        cmdBtn_Click 4 '-- Find
        If gClsSearch.SearchMultiRetCol(0) = "" Then Exit Sub
        FormAction = vbDataActionUpdate
        Call GProcActivateControls(Me, True)
        Call GProcSetButtons(Me, vbDataActionUpdate)
        SetCtrlEd
        txt(1).SetFocus
    'End If
Case 8 'Delete
    'If mURecDel = True Then
        If MsgBox("Do You Really Delete This Record?", vbQuestion + vbYesNo + vbDefaultButton1, "Confirm Delete") = vbYes Then
            If rstbl.RecordCount > 1 Then
                rstbl.Delete
                cmdBtn_Click (1)
            ElseIf rstbl.RecordCount = 1 Then
                rstbl.Delete
                Call GProcActivateControls(Me, False)
                Call GProcSetButtons(Me, vbDataActionCancel)
                Form_Load
            End If
        End If
    'End If
Case 9 'Print
Case 10 'Exit
    Unload Me
Case 11 '--  Account List Report
    frmMain.mnuRptGenrSubLst_Click (1)
Case 12 'Creation To Firm
    gfrmTypeStr = "Account"
    frmUtlCreationToFirm.Show 1
    rstbl.Requery
Case 13 'Save
    If ValidateData = True Then
        FillTxtFromTemp
        GProcSaveRecord Me, rstbl, FormAction, MaxNo
        WriteFields
'        gSearchMultiRetCol(0) = txt(1) '--AName
        Call GProcActivateControls(Me, False)
        Call GProcSetButtons(Me, True)
        If FormAction = vbDataActionAddNew Then
            If cmdBtn(6).Visible = True Then cmdBtn(6).SetFocus
        Else
            If cmdBtn(7).Visible = True Then cmdBtn(7).SetFocus
        End If
        FormAction = vbDataActionClose
    End If
Case 14 'Cancel
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
End Select
Exit Sub
ErrorRoutine:
    rstbl.CancelUpdate
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
    Me.Top = 50
    Me.Left = Screen.Width / 2 - Me.Width / 2
    X = "Select * from " & tblName
    If Trim(CriteriaStr) <> "" Then X = X & " where " & CriteriaStr
    If Trim(OrderField) <> "" Then X = X & " order by " & OrderField
    GProcRstOpen rstbl, X, "O"
    FirstTimeFlag = True
    '---- Button Visiblity Set As Per User Previlage
    GProcSetButtonVisiblity Me
    cboBalance.ListIndex = 0
    txtOpBal(1) = "Cr"
    SSTab1.Tab = 0
    cmdBtn_Click (6)
    cbo(24).ListIndex = 0
    cbo(33).ListIndex = 0
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
If FormAction = vbDataActionUpdate Then
    MsgBox "You are about to modify the record. Save the record and then exit.", vbInformation + vbOKOnly, Me.Caption
    Cancel = True
    Exit Sub
End If
'mShowRec = Space(5)
rstbl.Requery
rstbl.Close
End Sub

'Private Sub lstDrCr_DblClick()
'    Call lstDrCr_KeyPress(13)
'    DoEvents
'End Sub
'Private Sub lstDrCr_GotFocus()
'If Trim(txtOpBal(1).Text) <> "" Then
'    lstDrCr.Text = txtOpBal(1).Text
'Else
'    lstDrCr.ListIndex = 0
'End If
'End Sub
'Private Sub lstDrCr_KeyPress(KeyAscii As Integer)
'If KeyAscii = 13 Then
'    txtOpBal(1).Text = Trim(lstDrCr.Text)
'    lstDrCr.Visible = False
'    DoEvents
'End If
'End Sub
'Private Sub lstDrCr_LostFocus()
'    Call lstDrCr_KeyPress(13)
'    DoEvents
'End Sub
Private Sub SSTab1_Click(PreviousTab As Integer)
'SSTab1.SetFocus
End Sub

Private Sub txt_GotFocus(Index As Integer)
GProcSelectBox Me.ActiveControl
End Sub
Private Sub CheckForNumber()
'GProcCheckForNumber Me, Rstbl, MaxNo, Array(18, 15), Array(4, 5, 23, 36)
GProcCheckForNumber Me, rstbl, MaxNo, Array(18, 15, 41, 42), Array(23, 36)
End Sub
Private Sub txt_LostFocus(Index As Integer)
CheckForNumber
SaveBtnEd
txtOpBal(0).text = Format(txtOpBal(0), FStr)
If Index = 13 And txt(14) = "" Then
   txt(14) = txt(13)
End If


End Sub
Private Sub txt_KeyPress(Index As Integer, KeyAscii As Integer)
If KeyAscii = 13 Then Exit Sub
If Index = 18 Then
    KeyAscii = GProcCharacterOnly(KeyAscii)
Else
    KeyAscii = GProcValidateKey(rstbl, Index, KeyAscii, txt(Index).text)
End If
End Sub
Private Sub txtOpBal_GotFocus(Index As Integer)
GProcSelectBox Me.ActiveControl
End Sub
Private Sub txtOpBal_KeyPress(Index As Integer, KeyAscii As Integer)
If KeyAscii = 13 Then Exit Sub
'If Index = 1 Then
'   lstDrCr.Visible = True
'   lstDrCr.SetFocus
'   DoEvents
'End If
End Sub

Private Sub txtOpBal_LostFocus(Index As Integer)
If IsNumeric(txtOpBal(0)) = False Then txtOpBal(0) = "0.00"
txtOpBal(0).text = Format(txtOpBal(0), FStr)
'If Index = 0 Then txtOpBal(1).SetFocus
End Sub

Private Sub txttemp_GotFocus(Index As Integer)
GProcSelectBox Me.ActiveControl
End Sub

Private Sub txttemp_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
Dim GName As String
If KeyCode = vbKeyF3 Then
Select Case Index
    Case 16 '-- Company Group
        GProcShowForm frmMastCompGroup, frmMain.mnuMstfrm(2)
    Case 2, 3 '-- A/c Group
        GName = GProcGetColumnValue("TblMastGroup", "AgCode", 90017, "N", "AgName", "S")
        GProcShowForm frmMastGroup, frmMain.mnuMstfrm(1), Array(frmMastGroup.txtTemp(2)), Array(GName)
End Select
End If
End Sub
Private Sub txtTemp_KeyPress(Index As Integer, KeyAscii As Integer)
Dim LeftPos As Long
Dim TopPos As Long
TopPos = txtTemp(Index).Top + Me.Top + 650
LeftPos = txtTemp(Index).Left + Me.Left
If KeyAscii = 13 Then Exit Sub '-- Enter
'-- A/c Group
Select Case Index
Case 2, 3
    gClsSearch.SearchMultiField "tblMastGroup", "AGName", Array("Account Group"), Array(txtTemp(Index).Width), "", Chr(KeyAscii), "AgName", LeftPos, TopPos, True, , False
    KeyAscii = 0
'-- Company Group
Case 16
    gClsSearch.SearchMultiField "tblMastCompGroup", "CGName", Array("Company Group"), Array(txtTemp(Index).Width), " ", Chr(KeyAscii), "CGName", LeftPos, TopPos, True
    KeyAscii = 0
'-- Broker
Case 17
    gClsSearch.SearchMultiField "tblMastAccount", "AcName", Array("Broker"), Array(txtTemp(Index).Width), " (AgCode=90016) ", Chr(KeyAscii), "AcName", LeftPos, TopPos
    KeyAscii = 0
End Select
If gClsSearch.SearchMultiRetCol(0) <> "" Then txtTemp(Index) = gClsSearch.SearchMultiRetCol(0)
Set gClsSearch = Nothing
'KeyAscii = 0
End Sub
Private Sub SetCtrlEd()
Dim AgCode As Long
AgCode = GProcGetMainGroupCode(txtTemp(2))
'-- Mill,Sales,Purchase
'--Bill Sr No
If AgCode = 90029 Or AgCode = 90023 Or AgCode = 90025 Then
    txt(18).Visible = True
    lblBillSrNo.Visible = True
    If FormAction = vbDataActionAddNew Or FormAction = vbDataActionUpdate Then
        txt(18).Enabled = True
    Else
        txt(18).Enabled = False
    End If
Else
    lblBillSrNo.Visible = False
    txt(18).Visible = False
    txt(18).Enabled = False
    txt(18) = ""
End If
If AgCode <> 90017 Then
   lbl1(0).Caption = "T.D.S. % :"
End If



'-- Mill,Cr,Dr
'--OS Limit
If AgCode = 90029 Or AgCode = 90015 Or AgCode = 90017 Then
    txt(31).Visible = True
    lblOSLimit.Visible = True
    If FormAction = vbDataActionAddNew Or FormAction = vbDataActionUpdate Then
        txt(31).Enabled = True
    Else
        txt(31).Enabled = False
    End If
Else
    lblOSLimit.Visible = False
    txt(31).Visible = False
    txt(31) = 0
End If
'-- Mill Group
If GProcGetMainGroupCode(txtTemp(2)) = 90029 Then
    If FormAction = vbDataActionAddNew Or FormAction = vbDataActionUpdate Then
        cbo(33).Enabled = True '-- Mill Type
        cbo(24).Enabled = True '-- commission On
        cbo(35).Enabled = True '-- commission Type
        txt(25).Enabled = True '--Comm Rate
        Chk(26).Enabled = True '--Is Local Brk by self
'        Chk(27).Enabled = True '-- Is TDS Pay
        Chk(28).Enabled = True '--Service tax Pay
        If IsMillPurchaseSale = False Then
            Chk(29).Enabled = True '-- Exmpt Mill
            txt(30).Enabled = True '-- Exmpt Rate
        Else
            Chk(29).Enabled = False '-- Exemp Mill
            txt(30).Enabled = False '-- Exem Rate
        End If
   End If
Else
    cbo(33).Enabled = False '-- Mill Type
    cbo(35).Enabled = False '-- Comm Type
    cbo(33).ListIndex = 0
    cbo(24).Enabled = False
    txt(25).Enabled = False
    Chk(26).Enabled = False
 '   Chk(27).Enabled = False
    Chk(28).Enabled = False
    Chk(29).Enabled = False
    txt(30).Enabled = False
    txt(25) = "0.00"
    txt(30) = "0.00"
    Chk(26) = 0
  '  Chk(27) = 0
    Chk(28) = 0
    Chk(29) = 0
End If
txtOpBal(1).Enabled = False
End Sub
Private Function IsMillPurchaseSale() As Boolean
Dim X As String
Dim Rs1 As Recordset
X = "Select * from tblAddLess where AdMillCode=" & txt(0)
GProcRstOpen Rs1, X, "R"
If Rs1.EOF And Rs1.BOF Then
Else
    IsMillPurchaseSale = True
    Exit Function
End If
Rs1.Close
End Function
Private Sub txtTemp_LostFocus(Index As Integer)
Dim TopGroup As Long
If Index = 2 And txtTemp(3) = "" Then txtTemp(3) = txtTemp(2)
SaveBtnEd
SetCtrlEd
If IsNumeric(txtOpBal(0)) = False Then txtOpBal(0) = "0.00"
If txtOpBal(0) = 0 Then
    TopGroup = GProcGetColumnValue("TblMastGroup", "AgName", txtTemp(2), "S", "TopGroup", "N")
    If TopGroup = 90001 Then
        cboBalance.ListIndex = 1
        txtOpBal(1) = "Dr"
    Else
        cboBalance.ListIndex = 0
        txtOpBal(1) = "Cr"
    End If

End If
'TopGroup = GProcGetColumnValue("TblMastGroup", "AgName", txttemp(2), "S", "TopGroup", "N")
'If TopGroup = 90016 Then
'   lbl1(0).Caption = "Purch.Brok. % :"
'End If
End Sub
Private Sub FillCombo()
Call GProcFillCombo(cbo(7), "select Distinct AcCity from " & tblName & " where AcCity <> '' order by AcCity")
Call GProcFillCombo(cbo(8), "select Distinct AcState from " & tblName & " where AcState <> '' order by AcState")
Call GProcFillCombo(cbo(9), "select Distinct AcDist from " & tblName & " where AcDist <> '' order by AcDist")
End Sub
Private Sub SaveBtnEd()
'--- AcName,AgCode
If Trim(txt(1)) = "" Or Trim(txtTemp(2)) = "" Then
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
Dim rstAddCheck As Recordset
Dim rstModiCheck As Recordset
    CheckForNumber
    SetCtrlEd
    '-- AcCode
    If FormAction = vbDataActionAddNew Then txt(0) = GProcGenerateId(tblName, IdField, pStartCode:=gCSAcCode)
    If gCEAcCode > 0 And CLng(txt(0)) > gCEAcCode Then
        MsgBox "Year End is completed. You can not create new account.", vbCritical + vbOKOnly, Me.Caption
        txt(1).SetFocus
        Exit Function
    End If
    '----Required
    '--- Account Name
    If Trim(txt(1)) = "" Then
        MsgBox "Check Account Name.", vbCritical + vbOKOnly, Me.Caption
        txt(1).SetFocus
        Exit Function
    End If
    '---Group
    If Trim(txtTemp(2)) = "" Then
        MsgBox "Check Group Name.", vbCritical + vbOKOnly, Me.Caption
        txtTemp(2).SetFocus
        Exit Function
    End If
    '---Create Group
    If GProcGetColumnValue("tblMastGroup", "AgName", txtTemp(2), "S", "AgCode", "N") = 0 Then
        GProcShowForm frmMastGroup, frmMain.mnuMstfrm(1), Array(frmMastGroup.txt(1)), Array(txtTemp(2))
        Exit Function
    End If
    '---Group 2
    If Trim(txtTemp(3)) = "" Then txtTemp(3) = txtTemp(2)
    '---Create Group
    If GProcGetColumnValue("tblMastGroup", "AgName", txtTemp(3), "S", "AgCode", "N") = 0 Then
        GProcShowForm frmMastGroup, frmMain.mnuMstfrm(1), Array(frmMastGroup.txt(1)), Array(txtTemp(3))
        Exit Function
    End If
    '--- Company Group
    If txtTemp(16) = "" Then txtTemp(16) = "-"
'    If Trim(txttemp(16)) = "" Then
'        MsgBox "Check Company Group Name.", vbCritical + vbOKOnly, Me.Caption
'        txttemp(16).SetFocus
'        Exit Function
'    End If
    '---Create comp Group
    If GProcGetColumnValue("tblMastCompGroup", "CgName", txtTemp(16), "S", "CgCode", "N") = 0 Then
        GProcShowForm frmMastCompGroup, frmMain.mnuMstfrm(2), Array(frmMastCompGroup.txt(1)), Array(txtTemp(16))
        Exit Function
    End If
    '-- Mill Group Commission
    If GProcGetMainGroupCode(txtTemp(2)) = 29 Then
        If CDbl(txt(25)) = 0 And (cbo(33).ListIndex = 1 Or cbo(33).ListIndex = 2 Or cbo(33).ListIndex = 3) Then
            MsgBox "Check Commission Rate.", vbCritical + vbOKOnly, Me.Caption
            txt(25).SetFocus
            Exit Function
        End If
    End If
    '-- Commission Rate
    If CDbl(txt(25)) > 100 Then
        MsgBox "Check Commission Rate.", vbCritical + vbOKOnly, Me.Caption
        txt(25).SetFocus
        Exit Function
    End If
    '-- Interest Rate
    If CDbl(txt(19)) > 100 Then
        MsgBox "Check Interest %.", vbCritical + vbOKOnly, Me.Caption
        txt(19).SetFocus
        Exit Function
    End If
    If Chk(29) = 0 And CDbl(txt(30)) > 0 Then txt(30) = 0
    '-- Is Exemption Mill
    If Chk(29) = 1 And CDbl(txt(30)) = 0 Then
        MsgBox "Check Exemption %.", vbCritical + vbOKOnly, Me.Caption
        txt(30).SetFocus
        Exit Function
    End If
    '-- Exemption %
    If CDbl(txt(30)) > 100 Then
        MsgBox "Check Exemption %.", vbCritical + vbOKOnly, Me.Caption
        txt(30).SetFocus
        Exit Function
    End If
    '--- Opening
    If IsNumeric(txtOpBal(0)) = False Then txtOpBal(0) = "0.00"
    If txtOpBal(0) > 0 And cboBalance = "" Then
        MsgBox "Check Opening Balance Type.", vbCritical + vbOKOnly, Me.Caption
        cboBalance.SetFocus
        Exit Function
    End If
    '----Duplicate Records
    '-- Alias Name
    '----Addition Case
    If FormAction = vbDataActionAddNew And Trim(txt(6)) <> "" Then
        X = "Select ACCODE,AcAlName From " & tblName & " Where AcAlName='" & Trim(txt(6).text) & "'"
        i = GProcRstOpen(rstAddCheck, X, "R")
        If i > 0 Then
            MsgBox "Record is Available of A/c Alias Name  " & Trim(txt(6).text), vbCritical + vbOKOnly, "Duplicate Record"
            txt(6).SetFocus
            Exit Function
        End If
        rstAddCheck.Close
    End If
    '----Modification case
    If FormAction = vbDataActionUpdate And Trim(txt(6)) <> "" Then
        X = "Select ACCODE,AcAlName From " & tblName & " Where ((AcAlName='" & Trim(txt(6).text) & "') and (ACCODE<>" & txt(0).text & "))"
        i = GProcRstOpen(rstModiCheck, X, "R")
        If i > 0 Then
            MsgBox "Record is Available of A/c Alias Name  " & Trim(txt(6).text), vbCritical + vbOKOnly, "Duplicate Record"
            txt(6).SetFocus
            Exit Function
        End If
        rstModiCheck.Close
    End If
    '--- Bill SR No
    '----Addition Case
    If Trim(txt(18)) <> "" Then
    If FormAction = vbDataActionAddNew Then
        X = "Select AcBillSrNo From " & tblName & " Where AcBillSrNo='" & Trim(txt(18).text) & "'"
        i = GProcRstOpen(rstAddCheck, X, "R")
        If i > 0 Then
            MsgBox "Record is Available of Bill SrNo." & Trim(txt(18).text), vbCritical + vbOKOnly, "Duplicate Record"
            txt(18).SetFocus
            Exit Function
        End If
        rstAddCheck.Close
    End If
    '----Modification case
    If FormAction = vbDataActionUpdate Then
        X = "Select AcBillSrNo From " & tblName & " Where ((AcBillSrNo='" & Trim(txt(18).text) & "') and (ACCODE<>" & txt(0).text & "))"
        i = GProcRstOpen(rstModiCheck, X, "R")
        If i > 0 Then
            MsgBox "Record is Available of Bill SrNo." & Trim(txt(18).text), vbCritical + vbOKOnly, "Duplicate Record"
            txt(18).SetFocus
            Exit Function
        End If
        rstModiCheck.Close
    End If
    End If
    ValidateData = True
    '-- Account Name
    '----Addition Case
    If FormAction = vbDataActionAddNew Then
        X = "Select ACCODE,AcName From " & tblName & " Where AcName='" & Trim(txt(1).text) & "'"
        i = GProcRstOpen(rstAddCheck, X, "R")
        If i > 0 Then
            MsgBox "Record is Available of A/c Name  " & Trim(txt(1).text), vbCritical + vbOKOnly, "Duplicate Record"
            txt(1).SetFocus
            Exit Function
        End If
        rstAddCheck.Close
    End If
    '----Modification case
    If FormAction = vbDataActionUpdate Then
        X = "Select ACCODE,AcName From " & tblName & " Where ((AcName='" & Trim(txt(1).text) & "') and (ACCODE<>" & txt(0).text & "))"
        i = GProcRstOpen(rstModiCheck, X, "R")
        If i > 0 Then
            MsgBox "Record is Available of A/c Name  " & Trim(txt(1).text), vbCritical + vbOKOnly, "Duplicate Record"
            txt(1).SetFocus
            Exit Function
        End If
        rstModiCheck.Close
    End If
    ValidateData = True
End Function
Private Sub FillTxtFromTemp()
'-- A/c Group
txt(2) = GProcGetColumnValue("tblMastGroup", "AgName", txtTemp(2), "S", "AgCode", "N")
txt(3) = GProcGetColumnValue("tblMastGroup", "AgName", txtTemp(3), "S", "AgCode", "N")
 '--Com Group
If txtTemp(16) = "" Then txtTemp(16) = "-"
txt(16) = GProcGetColumnValue("tblMastCompGroup", "CgName", txtTemp(16), "S", "CgCode", "N")
'--Broker
If txtTemp(17) = "" Then txtTemp(17) = "-"
txt(17) = GProcGetColumnValue("tblMastAccount", "AcName", txtTemp(17), "S", "ACCode", "N")
txt(7) = cbo(7)
txt(8) = cbo(8)
txt(9) = cbo(9)
txt(24) = cbo(24).ListIndex
txt(33) = cbo(33).ListIndex '--Mill Type
txt(35) = cbo(35).ListIndex '-- Comm Type
txt(20) = Chk(20)
txt(21) = Chk(21)
txt(26) = Chk(26)
txt(27) = Chk(27)
txt(28) = Chk(28)
txt(29) = Chk(29)
txt(32) = Chk(32)
txt(34) = Chk(34) '--Is Db NT A/c

End Sub
Private Sub FillTempFromTxt()
Dim TopGroup As Long

txtTemp(2) = GProcGetColumnValue("tblMastGroup", "AgCode", txt(2), "N", "AgName", "S") '--Group
txtTemp(3) = GProcGetColumnValue("tblMastGroup", "AgCode", txt(3), "N", "AgName", "S")
txtTemp(16) = GProcGetColumnValue("tblMastCompGroup", "CGCode", txt(16), "N", "CgName", "S") '--Com Group
txtTemp(17) = GProcGetColumnValue("tblMastAccount", "AcCode", txt(17), "N", "AcName", "S") '--Broker
cbo(7) = txt(7) '--City
cbo(8) = txt(8) '-state
cbo(9) = txt(9) '--District
cbo(24).ListIndex = txt(24)  '--comm On
cbo(33).ListIndex = txt(33)  '--Mill Type
cbo(35).ListIndex = txt(35)  '--comm Type
Chk(20) = txt(20)
Chk(21) = txt(21)
Chk(26) = txt(26)
Chk(27) = txt(27)
Chk(28) = txt(28)
Chk(29) = txt(29)
Chk(32) = txt(32)
Chk(34) = txt(34) '--Is Db NT A/c
'txtOpBal(1) = "Cr"
cboBalance.ListIndex = 0
txtOpBal(1) = "Cr"
txtOpBal(0) = GProcGetColumnValue("tblVoucher", "VAcCode", txt(0), "N", "VAmt", "N", "vtype = '0P' and Vyear=" & gCYear, gCn)  '--Balance
If (txtOpBal(0) < 0) Then
   txtOpBal(0) = Format(Abs(Val(txtOpBal(0))), FStr)
   txtOpBal(1) = "Dr"
   cboBalance.ListIndex = 1
End If
txtOpBal(0) = Format(txtOpBal(0), FStr)
SetCtrlEd

TopGroup = GProcGetColumnValue("TblMastGroup", "AgName", txtTemp(2), "S", "AgCode", "N")
If TopGroup = 90016 Then
   lbl1(0).Caption = "Purch.Brok. % :"
Else
   lbl1(0).Caption = "Interest %        :"
End If


End Sub
Private Sub WriteFields()
Dim mRst1 As Recordset, mvno As Long, AcOCode As Long
AcOCode = GProcGetColumnValue("TblMastAccount", "AcName", "-", "S", "AcCode", "N")
If IsNumeric(txtOpBal(0)) = False Then txtOpBal(0) = "0.00"
If CDbl(txtOpBal(0)) = 0 Then
    gCn.Execute "delete from tblVoucher where Vtype='0P' and VAcCode=" & txt(0) & " and VYear=" & gCYear
Else
    If (GProcRstOpen(mRst1, "select * from tblvoucher where vtype = '0P' and vAccode = " & txt(0) & " and VYear=" & gCYear & " Order by vno", "R", gCn) > 0) Then
    '    mRst1!vamt = IIf(cboBalance = "Cr", txtOpBal(0), Val(txtOpBal(0)) * (-1))
    '    mRst1.Update
        gCn.Execute "update tblvoucher set Vamt=" & CDbl(IIf(cboBalance = "Cr", txtOpBal(0), Val(txtOpBal(0)) * (-1))) & " where vtype = '0P' and vAccode = " & txt(0) & " and VYear=" & gCYear
    Else
        mvno = GProcGenerateId("tblvoucher", "vno", "vtype = '0P' and VYear=" & gCYear, gCn)
        GProcCreateVoucher FormAction, mvno, 1, gCYSDate, "0P", txt(0), AcOCode, IIf(cboBalance = "Cr", txtOpBal(0), Val(txtOpBal(0)) * (-1))
    End If
    If mRst1.State = adStateOpen Then mRst1.Close
End If
End Sub
Private Sub ReadFields()
GProcShowRecord Me, rstbl, MaxNo
FillTempFromTxt
End Sub
