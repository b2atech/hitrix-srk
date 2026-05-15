VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form frmMastAccount 
   BackColor       =   &H00FBF2E1&
   BorderStyle     =   0  'None
   Caption         =   " Account Master"
   ClientHeight    =   9645
   ClientLeft      =   0
   ClientTop       =   -105
   ClientWidth     =   13605
   DrawMode        =   10  'Mask Pen
   FillColor       =   &H000040C0&
   FillStyle       =   0  'Solid
   Icon            =   "frmMastAccount.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   9645
   ScaleWidth      =   13605
   ShowInTaskbar   =   0   'False
   Begin VB.CheckBox chk 
      BackColor       =   &H00FBF2E1&
      Caption         =   "Is TDS Applicable From 1st Bill ?"
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
      Height          =   375
      Index           =   72
      Left            =   6960
      TabIndex        =   179
      Top             =   2760
      Width           =   4050
   End
   Begin VB.TextBox txt 
      Height          =   360
      Index           =   72
      Left            =   13860
      TabIndex        =   178
      Text            =   "72"
      Top             =   1680
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.CheckBox chk 
      BackColor       =   &H00FBF2E1&
      Caption         =   "Is TDS Applicable Party ?"
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
      Height          =   375
      Index           =   71
      Left            =   6960
      TabIndex        =   177
      Top             =   2280
      Width           =   2730
   End
   Begin VB.TextBox txt 
      Height          =   360
      Index           =   71
      Left            =   13920
      TabIndex        =   176
      Text            =   "71"
      Top             =   1320
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.TextBox txt 
      Height          =   360
      Index           =   70
      Left            =   16920
      TabIndex        =   175
      Text            =   "70"
      Top             =   6960
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.TextBox txt 
      Height          =   360
      Index           =   69
      Left            =   13860
      TabIndex        =   173
      Text            =   "69"
      Top             =   780
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.TextBox txt 
      Height          =   360
      Index           =   68
      Left            =   16920
      TabIndex        =   171
      Text            =   "68"
      Top             =   6540
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.TextBox txt 
      Height          =   360
      Index           =   67
      Left            =   16260
      TabIndex        =   167
      Text            =   "67"
      Top             =   120
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.TextBox txt 
      Height          =   360
      Index           =   66
      Left            =   15780
      TabIndex        =   166
      Text            =   "66"
      Top             =   360
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.TextBox txt 
      Height          =   360
      Index           =   65
      Left            =   15720
      TabIndex        =   164
      Text            =   "65"
      Top             =   60
      Visible         =   0   'False
      Width           =   375
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
      Index           =   2
      Left            =   2820
      TabIndex        =   9
      Top             =   2700
      Width           =   1425
   End
   Begin VB.TextBox txt 
      Height          =   360
      Index           =   63
      Left            =   17520
      TabIndex        =   161
      Text            =   "63"
      Top             =   8880
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
      Index           =   30
      Left            =   15840
      TabIndex        =   159
      Text            =   "30"
      Top             =   5880
      Visible         =   0   'False
      Width           =   945
   End
   Begin VB.TextBox txt 
      Height          =   360
      Index           =   62
      Left            =   17520
      TabIndex        =   157
      Text            =   "62"
      Top             =   8520
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.TextBox txt 
      Height          =   360
      Index           =   58
      Left            =   17010
      TabIndex        =   153
      Text            =   "58"
      Top             =   8505
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.TextBox txt 
      Height          =   360
      Index           =   57
      Left            =   17010
      TabIndex        =   152
      Text            =   "57"
      Top             =   8010
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.TextBox txt 
      Height          =   360
      Index           =   56
      Left            =   17055
      TabIndex        =   151
      Text            =   "56"
      Top             =   7605
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.TextBox txt 
      Height          =   360
      Index           =   55
      Left            =   17505
      TabIndex        =   146
      Text            =   "55"
      Top             =   7965
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.TextBox txt 
      Height          =   360
      Index           =   54
      Left            =   17550
      TabIndex        =   145
      Text            =   "54"
      Top             =   7650
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.TextBox txt 
      Height          =   360
      Index           =   53
      Left            =   17505
      TabIndex        =   143
      Text            =   "53"
      Top             =   7290
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.TextBox txt 
      Height          =   360
      Index           =   52
      Left            =   17460
      TabIndex        =   142
      Text            =   "52"
      Top             =   6930
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.TextBox txt 
      Height          =   360
      Index           =   51
      Left            =   17550
      TabIndex        =   141
      Text            =   "51"
      Top             =   6525
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.TextBox txt 
      Height          =   360
      Index           =   50
      Left            =   17550
      TabIndex        =   140
      Text            =   "50"
      Top             =   6120
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.TextBox txt 
      Height          =   360
      Index           =   46
      Left            =   17640
      TabIndex        =   136
      Text            =   "46"
      Top             =   5760
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.TextBox txt 
      Height          =   360
      Index           =   45
      Left            =   17730
      TabIndex        =   135
      Text            =   "45"
      Top             =   5355
      Visible         =   0   'False
      Width           =   375
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
      Index           =   45
      ItemData        =   "frmMastAccount.frx":058A
      Left            =   5445
      List            =   "frmMastAccount.frx":0594
      Style           =   2  'Dropdown List
      TabIndex        =   10
      Top             =   2820
      Visible         =   0   'False
      Width           =   1395
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
      Index           =   44
      Left            =   2805
      MaxLength       =   10
      TabIndex        =   2
      Text            =   "44"
      Top             =   915
      Width           =   1335
   End
   Begin VB.TextBox txt 
      Height          =   360
      Index           =   43
      Left            =   17685
      TabIndex        =   132
      Text            =   "43"
      Top             =   4995
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.TextBox txt 
      Height          =   360
      Index           =   34
      Left            =   17460
      TabIndex        =   115
      Text            =   "34"
      Top             =   585
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
      ItemData        =   "frmMastAccount.frx":05A9
      Left            =   4950
      List            =   "frmMastAccount.frx":05B3
      Style           =   2  'Dropdown List
      TabIndex        =   8
      Top             =   2235
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
      Left            =   17730
      TabIndex        =   114
      Text            =   "33"
      Top             =   1845
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
      Left            =   4260
      TabIndex        =   7
      Top             =   2235
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
      Left            =   2805
      TabIndex        =   6
      Top             =   2235
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
      Left            =   2805
      MaxLength       =   40
      TabIndex        =   5
      Text            =   "txttemp(16)"
      ToolTipText     =   "Press F3 For Creation"
      Top             =   1830
      Width           =   7095
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
      Left            =   16020
      MaxLength       =   50
      TabIndex        =   3
      Text            =   "6"
      Top             =   5535
      Visible         =   0   'False
      Width           =   300
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
      Left            =   13260
      MaxLength       =   40
      TabIndex        =   111
      Text            =   "32"
      Top             =   9720
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.TextBox txt 
      Height          =   360
      Index           =   29
      Left            =   16080
      TabIndex        =   110
      Text            =   "29"
      Top             =   6660
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.TextBox txt 
      Height          =   360
      Index           =   28
      Left            =   12300
      TabIndex        =   109
      Text            =   "28"
      Top             =   9720
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.TextBox txt 
      Height          =   360
      Index           =   27
      Left            =   11820
      TabIndex        =   108
      Text            =   "27"
      Top             =   9720
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
      Left            =   11100
      TabIndex        =   107
      Text            =   "21"
      Top             =   9720
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
      Left            =   17730
      TabIndex        =   106
      Text            =   "17"
      Top             =   2790
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
      Left            =   17640
      TabIndex        =   105
      Text            =   "16"
      Top             =   2340
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
      Left            =   10500
      TabIndex        =   104
      Text            =   "20"
      Top             =   9720
      Visible         =   0   'False
      Width           =   375
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   6180
      Left            =   300
      TabIndex        =   74
      Top             =   3240
      Width           =   10875
      _ExtentX        =   19182
      _ExtentY        =   10901
      _Version        =   393216
      Tabs            =   4
      Tab             =   1
      TabsPerRow      =   4
      TabHeight       =   617
      BackColor       =   15708538
      ForeColor       =   8388608
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
      TabPicture(0)   =   "frmMastAccount.frx":05BF
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "chk(69)"
      Tab(0).Control(1)=   "chk(68)"
      Tab(0).Control(2)=   "chk(67)"
      Tab(0).Control(3)=   "chk(66)"
      Tab(0).Control(4)=   "chk(65)"
      Tab(0).Control(5)=   "chk(54)"
      Tab(0).Control(6)=   "txt(49)"
      Tab(0).Control(7)=   "txt(41)"
      Tab(0).Control(8)=   "txt(42)"
      Tab(0).Control(9)=   "txt(36)"
      Tab(0).Control(10)=   "txt(11)"
      Tab(0).Control(11)=   "txt(12)"
      Tab(0).Control(12)=   "txt(23)"
      Tab(0).Control(13)=   "cbo(8)"
      Tab(0).Control(14)=   "txt(10)"
      Tab(0).Control(15)=   "cbo(9)"
      Tab(0).Control(16)=   "cbo(7)"
      Tab(0).Control(17)=   "txt(4)"
      Tab(0).Control(18)=   "txt(5)"
      Tab(0).Control(19)=   "lbl1(26)"
      Tab(0).Control(20)=   "lbl1(24)"
      Tab(0).Control(21)=   "lbl1(3)"
      Tab(0).Control(22)=   "lbl1(9)"
      Tab(0).Control(23)=   "lbl1(10)"
      Tab(0).Control(24)=   "lbl1(11)"
      Tab(0).Control(25)=   "lbl1(13)"
      Tab(0).Control(26)=   "lbl1(12)"
      Tab(0).Control(27)=   "lbl1(7)"
      Tab(0).Control(28)=   "lbl1(8)"
      Tab(0).Control(29)=   "lbl1(6)"
      Tab(0).Control(30)=   "Shape3"
      Tab(0).ControlCount=   31
      TabCaption(1)   =   "Personal Deta&ils"
      TabPicture(1)   =   "frmMastAccount.frx":05DB
      Tab(1).ControlEnabled=   -1  'True
      Tab(1).Control(0)=   "Shape6"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "lblOSLimit"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "lbl1(4)"
      Tab(1).Control(2).Enabled=   0   'False
      Tab(1).Control(3)=   "lbl1(0)"
      Tab(1).Control(3).Enabled=   0   'False
      Tab(1).Control(4)=   "lbl1(5)"
      Tab(1).Control(4).Enabled=   0   'False
      Tab(1).Control(5)=   "Label1"
      Tab(1).Control(5).Enabled=   0   'False
      Tab(1).Control(6)=   "lbl1(31)"
      Tab(1).Control(6).Enabled=   0   'False
      Tab(1).Control(7)=   "txt(31)"
      Tab(1).Control(7).Enabled=   0   'False
      Tab(1).Control(8)=   "txt(26)"
      Tab(1).Control(8).Enabled=   0   'False
      Tab(1).Control(9)=   "chk(20)"
      Tab(1).Control(9).Enabled=   0   'False
      Tab(1).Control(10)=   "chk(21)"
      Tab(1).Control(10).Enabled=   0   'False
      Tab(1).Control(11)=   "txtTemp(17)"
      Tab(1).Control(11).Enabled=   0   'False
      Tab(1).Control(12)=   "txtTemp(3)"
      Tab(1).Control(12).Enabled=   0   'False
      Tab(1).Control(13)=   "txt(19)"
      Tab(1).Control(13).Enabled=   0   'False
      Tab(1).Control(14)=   "txt(48)"
      Tab(1).Control(14).Enabled=   0   'False
      Tab(1).Control(15)=   "chk(55)"
      Tab(1).Control(15).Enabled=   0   'False
      Tab(1).Control(16)=   "txt(61)"
      Tab(1).Control(16).Enabled=   0   'False
      Tab(1).ControlCount=   17
      TabCaption(2)   =   "&Mill Details"
      TabPicture(2)   =   "frmMastAccount.frx":05F7
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Shape5"
      Tab(2).Control(1)=   "lbl1(19)"
      Tab(2).Control(2)=   "lbl1(18)"
      Tab(2).Control(3)=   "lbl1(2)"
      Tab(2).Control(4)=   "lbl1(23)"
      Tab(2).Control(5)=   "lbl1(25)"
      Tab(2).Control(6)=   "lblBillSrNo"
      Tab(2).Control(7)=   "lbl1(27)"
      Tab(2).Control(8)=   "lbl1(28)"
      Tab(2).Control(9)=   "lbl1(32)"
      Tab(2).Control(10)=   "lbl1(1)"
      Tab(2).Control(11)=   "lbl1(34)"
      Tab(2).Control(12)=   "lbl1(35)"
      Tab(2).Control(13)=   "txt(25)"
      Tab(2).Control(14)=   "cbo(24)"
      Tab(2).Control(15)=   "txt(35)"
      Tab(2).Control(16)=   "cbo(35)"
      Tab(2).Control(17)=   "cbo(33)"
      Tab(2).Control(18)=   "chk(26)"
      Tab(2).Control(19)=   "chk(27)"
      Tab(2).Control(20)=   "chk(28)"
      Tab(2).Control(21)=   "chk(29)"
      Tab(2).Control(22)=   "txtTemp(43)"
      Tab(2).Control(23)=   "txt(18)"
      Tab(2).Control(24)=   "chk(46)"
      Tab(2).Control(25)=   "txt(47)"
      Tab(2).Control(26)=   "txtTemp(50)"
      Tab(2).Control(27)=   "txtTemp(51)"
      Tab(2).Control(28)=   "chk(52)"
      Tab(2).Control(29)=   "chk(53)"
      Tab(2).Control(30)=   "chk(56)"
      Tab(2).Control(31)=   "chk(57)"
      Tab(2).Control(32)=   "chk(58)"
      Tab(2).Control(33)=   "txtTemp(62)"
      Tab(2).Control(34)=   "txtTemp(63)"
      Tab(2).Control(35)=   "txt(64)"
      Tab(2).Control(36)=   "chk(70)"
      Tab(2).ControlCount=   37
      TabCaption(3)   =   "&TIN-PAN-RTGS Details"
      TabPicture(3)   =   "frmMastAccount.frx":0613
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "Shape8"
      Tab(3).Control(1)=   "Label15(5)"
      Tab(3).Control(2)=   "Label15(4)"
      Tab(3).Control(3)=   "Label15(3)"
      Tab(3).Control(4)=   "Label15(2)"
      Tab(3).Control(5)=   "Label15(6)"
      Tab(3).Control(6)=   "lbl1(14)"
      Tab(3).Control(7)=   "lbl1(15)"
      Tab(3).Control(8)=   "lbl1(16)"
      Tab(3).Control(9)=   "lbl1(17)"
      Tab(3).Control(10)=   "Line1"
      Tab(3).Control(11)=   "lbl1(29)"
      Tab(3).Control(12)=   "lbl1(30)"
      Tab(3).Control(13)=   "txt(40)"
      Tab(3).Control(14)=   "txt(39)"
      Tab(3).Control(15)=   "txt(38)"
      Tab(3).Control(16)=   "txt(37)"
      Tab(3).Control(17)=   "txt(15)"
      Tab(3).Control(18)=   "txt(13)"
      Tab(3).Control(19)=   "txt(14)"
      Tab(3).Control(20)=   "txt(22)"
      Tab(3).Control(21)=   "txt(59)"
      Tab(3).Control(22)=   "txt(60)"
      Tab(3).ControlCount=   23
      Begin VB.CheckBox chk 
         BackColor       =   &H00FBF2E1&
         Caption         =   " Less TCS Amount In Commission"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF00FF&
         Height          =   375
         Index           =   70
         Left            =   -69120
         TabIndex        =   150
         Top             =   5280
         Width           =   4575
      End
      Begin VB.CheckBox chk 
         BackColor       =   &H00FBF2E1&
         Caption         =   "Is Exp.Sub A/c  Det.  (For Audit)  Req. ?"
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
         Height          =   375
         Index           =   69
         Left            =   -70800
         TabIndex        =   174
         Top             =   5400
         Width           =   4110
      End
      Begin VB.CheckBox chk 
         BackColor       =   &H00FBF2E1&
         Caption         =   "Is Opp. A/c wise Det.  (For Audit)  Req. ?"
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
         Height          =   375
         Index           =   68
         Left            =   -70800
         TabIndex        =   172
         Top             =   5700
         Width           =   4170
      End
      Begin VB.CheckBox chk 
         BackColor       =   &H00FBF2E1&
         Caption         =   "Is Asset Det. Req. ?"
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
         Height          =   375
         Index           =   67
         Left            =   -66540
         TabIndex        =   169
         Top             =   5400
         Width           =   2190
      End
      Begin VB.CheckBox chk 
         BackColor       =   &H00FBF2E1&
         Caption         =   "Is Loan Det. Req. ?"
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
         Height          =   375
         Index           =   66
         Left            =   -66540
         TabIndex        =   168
         Top             =   5700
         Width           =   2130
      End
      Begin VB.CheckBox chk 
         BackColor       =   &H00FBF2E1&
         Caption         =   "Is Opp. A/c wise Det. Req. ?"
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
         Height          =   375
         Index           =   65
         Left            =   -74280
         TabIndex        =   165
         Top             =   5700
         Width           =   3030
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
         Index           =   64
         Left            =   -72120
         TabIndex        =   45
         Text            =   "64"
         Top             =   4740
         Width           =   1335
      End
      Begin VB.TextBox txtTemp 
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
         Index           =   63
         Left            =   -72120
         MaxLength       =   60
         TabIndex        =   39
         Text            =   "txtTemp(63)"
         Top             =   2400
         Width           =   7095
      End
      Begin VB.TextBox txtTemp 
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
         Index           =   62
         Left            =   -72120
         MaxLength       =   60
         TabIndex        =   38
         Text            =   "txtTemp(62)"
         Top             =   1920
         Width           =   7095
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
         Left            =   9585
         TabIndex        =   30
         Text            =   "61"
         Top             =   3060
         Width           =   615
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
         Index           =   60
         Left            =   -70605
         MaxLength       =   11
         TabIndex        =   57
         Text            =   "60"
         Top             =   3015
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
         Index           =   59
         Left            =   -70605
         MaxLength       =   11
         TabIndex        =   53
         Text            =   "59"
         Top             =   1215
         Width           =   4335
      End
      Begin VB.CheckBox chk 
         BackColor       =   &H00FBF2E1&
         Caption         =   " Less GST Amount In Commission"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF00FF&
         Height          =   375
         Index           =   58
         Left            =   -69105
         TabIndex        =   149
         Top             =   4965
         Width           =   4575
      End
      Begin VB.CheckBox chk 
         BackColor       =   &H00FBF2E1&
         Caption         =   " Less Unloading Amount In Commission"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF00FF&
         Height          =   375
         Index           =   57
         Left            =   -69105
         TabIndex        =   148
         Top             =   4650
         Width           =   4575
      End
      Begin VB.CheckBox chk 
         BackColor       =   &H00FBF2E1&
         Caption         =   " Less Brokerage Amount In Commission"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF00FF&
         Height          =   375
         Index           =   56
         Left            =   -69105
         TabIndex        =   147
         Top             =   4335
         Width           =   4575
      End
      Begin VB.CheckBox chk 
         BackColor       =   &H00FBF2E1&
         Caption         =   "Is Interest Debit Note Monthly ?"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   375
         Index           =   55
         Left            =   3465
         TabIndex        =   31
         Top             =   3555
         Width           =   3660
      End
      Begin VB.CheckBox chk 
         BackColor       =   &H00FBF2E1&
         Caption         =   "Is Corporate Party ?"
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
         Height          =   375
         Index           =   54
         Left            =   -74280
         TabIndex        =   24
         Top             =   5400
         Width           =   2250
      End
      Begin VB.CheckBox chk 
         BackColor       =   &H00FBF2E1&
         Caption         =   "Is Tax In Depot Sale Paid By Mill ?"
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
         Height          =   375
         Index           =   53
         Left            =   -68400
         TabIndex        =   144
         Top             =   2850
         Width           =   3720
      End
      Begin VB.CheckBox chk 
         BackColor       =   &H00FBF2E1&
         Caption         =   "No Due Date Print On Sale Bill "
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
         Height          =   375
         Index           =   52
         Left            =   -71640
         TabIndex        =   41
         Top             =   2850
         Width           =   3240
      End
      Begin VB.TextBox txtTemp 
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
         Index           =   51
         Left            =   -72120
         MaxLength       =   60
         TabIndex        =   36
         Text            =   "txtTemp(51)"
         Top             =   1020
         Width           =   7095
      End
      Begin VB.TextBox txtTemp 
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
         Index           =   50
         Left            =   -72120
         MaxLength       =   60
         TabIndex        =   35
         Text            =   "txtTemp(50)"
         Top             =   645
         Width           =   7095
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
         Index           =   49
         Left            =   -71940
         MaxLength       =   50
         TabIndex        =   13
         Text            =   "49"
         Top             =   1440
         Width           =   6705
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
         Left            =   3480
         TabIndex        =   28
         Text            =   "48"
         Top             =   2565
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
         Index           =   47
         Left            =   -72120
         TabIndex        =   46
         Text            =   "47"
         Top             =   5505
         Width           =   1335
      End
      Begin VB.CheckBox chk 
         BackColor       =   &H00FBF2E1&
         Caption         =   "Payment Direct To Mill ?"
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
         Height          =   375
         Index           =   46
         Left            =   -74400
         TabIndex        =   40
         Top             =   2850
         Width           =   2655
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
         Left            =   -64785
         MaxLength       =   3
         TabIndex        =   51
         Text            =   "18"
         Top             =   810
         Visible         =   0   'False
         Width           =   405
      End
      Begin VB.TextBox txtTemp 
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
         Index           =   43
         Left            =   -72120
         MaxLength       =   60
         TabIndex        =   37
         Text            =   "txtTemp(43)"
         Top             =   1470
         Width           =   7095
      End
      Begin VB.CheckBox chk 
         BackColor       =   &H00FBF2E1&
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
         ForeColor       =   &H00800000&
         Height          =   375
         Index           =   29
         Left            =   -60600
         TabIndex        =   47
         Top             =   3000
         Width           =   2055
      End
      Begin VB.CheckBox chk 
         BackColor       =   &H00FBF2E1&
         Caption         =   "Mill Required in Commission O/s (Daily Report)"
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
         Index           =   28
         Left            =   -69105
         TabIndex        =   49
         Top             =   3975
         Width           =   4485
      End
      Begin VB.CheckBox chk 
         BackColor       =   &H00FBF2E1&
         Caption         =   "No Unloading Debit Note In Monthly Statement"
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
         Index           =   27
         Left            =   -69105
         TabIndex        =   50
         Top             =   3615
         Width           =   4470
      End
      Begin VB.CheckBox chk 
         BackColor       =   &H00FBF2E1&
         Caption         =   "No Brokerage Debit Note In  Monthly Statement"
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
         Index           =   26
         Left            =   -69105
         TabIndex        =   48
         Top             =   3255
         Width           =   4575
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
         ItemData        =   "frmMastAccount.frx":062F
         Left            =   -64695
         List            =   "frmMastAccount.frx":0645
         Style           =   2  'Dropdown List
         TabIndex        =   34
         Top             =   405
         Visible         =   0   'False
         Width           =   705
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
         Left            =   -70605
         MaxLength       =   15
         TabIndex        =   56
         Text            =   "22"
         Top             =   2565
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
         Left            =   -70605
         MaxLength       =   11
         TabIndex        =   55
         Text            =   "14"
         Top             =   2115
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
         Left            =   -70605
         MaxLength       =   11
         TabIndex        =   54
         Text            =   "13"
         Top             =   1665
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
         Index           =   15
         Left            =   -70605
         MaxLength       =   10
         TabIndex        =   52
         Text            =   "15"
         Top             =   765
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
         Index           =   41
         Left            =   -71925
         MaxLength       =   40
         TabIndex        =   22
         Text            =   "41"
         Top             =   4530
         Width           =   6705
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
         Index           =   42
         Left            =   -71925
         MaxLength       =   40
         TabIndex        =   23
         Text            =   "42"
         Top             =   5010
         Width           =   6705
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
         Left            =   -70605
         MaxLength       =   40
         TabIndex        =   58
         Text            =   "37"
         Top             =   3600
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
         Left            =   -70590
         MaxLength       =   40
         TabIndex        =   61
         Text            =   "38"
         Top             =   4950
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
         Left            =   -70605
         MaxLength       =   40
         TabIndex        =   60
         Text            =   "39"
         Top             =   4515
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
         Index           =   40
         Left            =   -70575
         MaxLength       =   20
         TabIndex        =   59
         Text            =   "40"
         Top             =   4050
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
         Left            =   -71925
         MaxLength       =   15
         TabIndex        =   20
         Text            =   "36"
         Top             =   3630
         Width           =   6705
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
         ItemData        =   "frmMastAccount.frx":0686
         Left            =   -72120
         List            =   "frmMastAccount.frx":0693
         Style           =   2  'Dropdown List
         TabIndex        =   42
         Top             =   3360
         Width           =   2295
      End
      Begin VB.TextBox txt 
         Height          =   360
         Index           =   35
         Left            =   -75000
         TabIndex        =   116
         Text            =   "35"
         Top             =   0
         Visible         =   0   'False
         Width           =   375
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
         ItemData        =   "frmMastAccount.frx":06C3
         Left            =   -72120
         List            =   "frmMastAccount.frx":06CD
         Style           =   2  'Dropdown List
         TabIndex        =   43
         Top             =   3780
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
         Left            =   -72120
         TabIndex        =   44
         Text            =   "25"
         Top             =   4215
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
         Height          =   375
         Index           =   19
         Left            =   3480
         TabIndex        =   29
         Text            =   "19"
         Top             =   3060
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
         Left            =   3480
         MaxLength       =   40
         TabIndex        =   26
         Text            =   "txttemp(3)"
         ToolTipText     =   "Press F3 For Creation"
         Top             =   1530
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
         Left            =   3510
         MaxLength       =   60
         TabIndex        =   25
         Text            =   "txttemp(17)"
         Top             =   1080
         Width           =   6105
      End
      Begin VB.CheckBox chk 
         BackColor       =   &H00FBF2E1&
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
         ForeColor       =   &H00800000&
         Height          =   375
         Index           =   21
         Left            =   3480
         TabIndex        =   33
         Top             =   4680
         Width           =   6735
      End
      Begin VB.CheckBox chk 
         BackColor       =   &H00FBF2E1&
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
         ForeColor       =   &H00800000&
         Height          =   375
         Index           =   20
         Left            =   3480
         TabIndex        =   32
         Top             =   4185
         Width           =   1545
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
         Left            =   12825
         TabIndex        =   112
         Text            =   "26"
         Top             =   5490
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
         Left            =   3480
         TabIndex        =   27
         Text            =   "31"
         Top             =   2085
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
         Left            =   -71925
         MaxLength       =   40
         TabIndex        =   18
         Text            =   "11"
         Top             =   2730
         Width           =   6705
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
         Left            =   -71925
         MaxLength       =   50
         TabIndex        =   19
         Text            =   "12"
         Top             =   3180
         Width           =   6705
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
         Left            =   -71925
         MaxLength       =   400
         TabIndex        =   21
         Text            =   "23"
         Top             =   4080
         Width           =   6705
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
         Left            =   -67575
         TabIndex        =   17
         Text            =   "cbo(8)"
         Top             =   2325
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
         Left            =   -71925
         MaxLength       =   10
         TabIndex        =   16
         Text            =   "10"
         Top             =   2325
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
         Left            =   -67575
         TabIndex        =   15
         Text            =   "cbo(9)"
         Top             =   1890
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
         Left            =   -71925
         TabIndex        =   14
         Text            =   "cbo(7)"
         Top             =   1890
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
         Left            =   -71925
         MaxLength       =   50
         TabIndex        =   11
         Text            =   "4"
         Top             =   630
         Width           =   6705
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
         Left            =   -71925
         MaxLength       =   50
         TabIndex        =   12
         Text            =   "5"
         Top             =   1020
         Width           =   6705
      End
      Begin VB.Label lbl1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Cm. Rt Direct Pay to Mill   :"
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
         Height          =   240
         Index           =   35
         Left            =   -74640
         TabIndex        =   170
         Top             =   4980
         Width           =   2415
      End
      Begin VB.Label lbl1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Cm. 2nd Day Bl Rt (depot) :"
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
         Height          =   240
         Index           =   34
         Left            =   -74625
         TabIndex        =   163
         Top             =   4680
         Width           =   2415
      End
      Begin VB.Label lbl1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Firm -  H.O. A/c         :"
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
         Height          =   240
         Index           =   1
         Left            =   -74370
         TabIndex        =   160
         Top             =   2460
         Width           =   1995
      End
      Begin VB.Label lbl1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Depot -  H.O. A/c       :"
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
         Height          =   240
         Index           =   32
         Left            =   -74370
         TabIndex        =   158
         Top             =   1980
         Width           =   1995
      End
      Begin VB.Label lbl1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Sr. No For Price List and Mill Booking  :"
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
         Height          =   240
         Index           =   31
         Left            =   6030
         TabIndex        =   156
         Top             =   3127
         Width           =   3465
      End
      Begin VB.Label lbl1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "GSTIN  Of Depot H.O.      :"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   240
         Index           =   30
         Left            =   -73335
         TabIndex        =   155
         Top             =   3060
         Width           =   2340
      End
      Begin VB.Label lbl1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "GSTIN        :"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   240
         Index           =   29
         Left            =   -72135
         TabIndex        =   154
         Top             =   1260
         Width           =   1110
      End
      Begin VB.Label lbl1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Depot In Firm             :"
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
         Height          =   240
         Index           =   28
         Left            =   -74370
         TabIndex        =   139
         Top             =   1080
         Width           =   2010
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Grase Days     :"
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
         Height          =   240
         Left            =   1920
         TabIndex        =   138
         Top             =   2610
         Width           =   1395
      End
      Begin VB.Label lbl1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Cash Disc. Rate    :"
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
         Height          =   240
         Index           =   27
         Left            =   -74145
         TabIndex        =   137
         Top             =   5550
         Width           =   1725
      End
      Begin VB.Label lblBillSrNo 
         BackStyle       =   0  'Transparent
         Caption         =   "Bill Sr No.                 :"
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
         Left            =   -64605
         TabIndex        =   131
         Top             =   810
         Visible         =   0   'False
         Width           =   825
      End
      Begin VB.Label lbl1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Mill Bill In Firm           :"
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
         Height          =   240
         Index           =   25
         Left            =   -74370
         TabIndex        =   130
         Top             =   1530
         Width           =   2010
      End
      Begin VB.Label lbl1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Consignment In Firm   :"
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
         Height          =   240
         Index           =   23
         Left            =   -74370
         TabIndex        =   129
         Top             =   690
         Width           =   2040
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00800000&
         X1              =   -74595
         X2              =   -64560
         Y1              =   3465
         Y2              =   3465
      End
      Begin VB.Label lbl1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "TAN NO      :"
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
         Height          =   240
         Index           =   17
         Left            =   -72135
         TabIndex        =   128
         Top             =   2610
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
         ForeColor       =   &H00800000&
         Height          =   240
         Index           =   16
         Left            =   -72135
         TabIndex        =   127
         Top             =   2160
         Width           =   1140
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
         ForeColor       =   &H00800000&
         Height          =   240
         Index           =   15
         Left            =   -72135
         TabIndex        =   126
         Top             =   1755
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
         ForeColor       =   &H00800000&
         Height          =   360
         Index           =   14
         Left            =   -72135
         TabIndex        =   125
         Top             =   765
         Width           =   1275
      End
      Begin VB.Label Label15 
         BackStyle       =   0  'Transparent
         Caption         =   "RTGS Details   :"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   360
         Index           =   6
         Left            =   -74190
         TabIndex        =   124
         Top             =   3600
         Width           =   1665
      End
      Begin VB.Label lbl1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Remark               :"
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
         Height          =   240
         Index           =   26
         Left            =   -73905
         TabIndex        =   123
         Top             =   5070
         Width           =   1650
      End
      Begin VB.Label lbl1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Proprietor Name   :"
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
         Height          =   240
         Index           =   24
         Left            =   -73905
         TabIndex        =   122
         Top             =   4590
         Width           =   1650
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
         ForeColor       =   &H00800000&
         Height          =   255
         Index           =   2
         Left            =   -72135
         TabIndex        =   121
         Top             =   3615
         Width           =   1380
      End
      Begin VB.Label Label15 
         BackStyle       =   0  'Transparent
         Caption         =   "A/c No           :"
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
         Index           =   3
         Left            =   -72135
         TabIndex        =   120
         Top             =   4065
         Width           =   1470
      End
      Begin VB.Label Label15 
         BackStyle       =   0  'Transparent
         Caption         =   "Branch          :"
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
         Index           =   4
         Left            =   -72135
         TabIndex        =   119
         Top             =   5040
         Width           =   1425
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
         ForeColor       =   &H00800000&
         Height          =   255
         Index           =   5
         Left            =   -72135
         TabIndex        =   118
         Top             =   4560
         Width           =   1380
      End
      Begin VB.Shape Shape8 
         BorderColor     =   &H00800000&
         FillColor       =   &H00FBF2E1&
         FillStyle       =   0  'Solid
         Height          =   4770
         Left            =   -74550
         Top             =   615
         Width           =   10020
      End
      Begin VB.Label lbl1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Mbl No For SMS  :"
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
         Height          =   240
         Index           =   3
         Left            =   -73905
         TabIndex        =   72
         Top             =   3247
         Width           =   1650
      End
      Begin VB.Label lbl1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Commission Type   :"
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
         Height          =   240
         Index           =   2
         Left            =   -74145
         TabIndex        =   117
         Top             =   3360
         Width           =   1800
      End
      Begin VB.Label lbl1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Commission On     :"
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
         Height          =   240
         Index           =   18
         Left            =   -74145
         TabIndex        =   79
         Top             =   3765
         Width           =   1755
      End
      Begin VB.Label lbl1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Commission Rate  :"
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
         Height          =   240
         Index           =   19
         Left            =   -74145
         TabIndex        =   80
         Top             =   4260
         Width           =   1725
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
         ForeColor       =   &H00800000&
         Height          =   240
         Index           =   5
         Left            =   1920
         TabIndex        =   75
         Top             =   1050
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
         ForeColor       =   &H00800000&
         Height          =   375
         Index           =   0
         Left            =   1920
         TabIndex        =   78
         Top             =   3060
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
         ForeColor       =   &H00800000&
         Height          =   375
         Index           =   4
         Left            =   1920
         TabIndex        =   76
         Top             =   1530
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
         ForeColor       =   &H00800000&
         Height          =   240
         Left            =   1920
         TabIndex        =   77
         Top             =   2085
         Width           =   1425
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
         ForeColor       =   &H00800000&
         Height          =   375
         Index           =   9
         Left            =   -73905
         TabIndex        =   70
         Top             =   2730
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
         ForeColor       =   &H00800000&
         Height          =   375
         Index           =   10
         Left            =   -73905
         TabIndex        =   71
         Top             =   3630
         Width           =   1650
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
         ForeColor       =   &H00800000&
         Height          =   240
         Index           =   11
         Left            =   -73905
         TabIndex        =   73
         Top             =   4080
         Width           =   1650
      End
      Begin VB.Label lbl1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Pin Code             :"
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
         Height          =   240
         Index           =   13
         Left            =   -73905
         TabIndex        =   68
         Top             =   2325
         Width           =   1635
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
         ForeColor       =   &H00800000&
         Height          =   240
         Index           =   12
         Left            =   -68580
         TabIndex        =   69
         Top             =   2325
         Width           =   705
      End
      Begin VB.Label lbl1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "City                    :"
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
         Height          =   240
         Index           =   7
         Left            =   -73905
         TabIndex        =   66
         Top             =   1890
         Width           =   1605
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
         ForeColor       =   &H00800000&
         Height          =   240
         Index           =   8
         Left            =   -68580
         TabIndex        =   67
         Top             =   1890
         Width           =   705
      End
      Begin VB.Label lbl1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Address              : "
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
         Height          =   240
         Index           =   6
         Left            =   -73905
         TabIndex        =   65
         Top             =   705
         Width           =   1680
      End
      Begin VB.Shape Shape3 
         BorderColor     =   &H00800000&
         FillColor       =   &H00FBF2E1&
         FillStyle       =   0  'Solid
         Height          =   5640
         Left            =   -74760
         Top             =   450
         Width           =   10515
      End
      Begin VB.Shape Shape6 
         BorderColor     =   &H00800000&
         FillColor       =   &H00FBF2E1&
         FillStyle       =   0  'Solid
         Height          =   4725
         Left            =   270
         Top             =   510
         Width           =   10440
      End
      Begin VB.Shape Shape5 
         BorderColor     =   &H00800000&
         FillColor       =   &H00FBF2E1&
         FillStyle       =   0  'Solid
         Height          =   5475
         Left            =   -74670
         Top             =   525
         Width           =   10260
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
      Left            =   15615
      Style           =   1  'Graphical
      TabIndex        =   95
      Top             =   3915
      Visible         =   0   'False
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
      Left            =   11790
      Style           =   1  'Graphical
      TabIndex        =   93
      Top             =   3090
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
      Left            =   11790
      Style           =   1  'Graphical
      TabIndex        =   94
      ToolTipText     =   "Exit"
      Top             =   7125
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
      Left            =   11790
      Style           =   1  'Graphical
      TabIndex        =   85
      ToolTipText     =   "To Cancel Record"
      Top             =   6660
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
      Left            =   11790
      Style           =   1  'Graphical
      TabIndex        =   84
      ToolTipText     =   "To Save Record"
      Top             =   6225
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
      Left            =   11790
      Style           =   1  'Graphical
      TabIndex        =   86
      Top             =   3570
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
      Left            =   11790
      Style           =   1  'Graphical
      TabIndex        =   83
      ToolTipText     =   "To Delete Record"
      Top             =   5445
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
      Left            =   11790
      Style           =   1  'Graphical
      TabIndex        =   82
      ToolTipText     =   "To Modify Record"
      Top             =   4980
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
      Left            =   11790
      Style           =   1  'Graphical
      TabIndex        =   81
      ToolTipText     =   "To Add New Record"
      Top             =   4515
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
      Left            =   15615
      Style           =   1  'Graphical
      TabIndex        =   92
      ToolTipText     =   "Show Mill List"
      Top             =   4860
      Visible         =   0   'False
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
      Left            =   11790
      Style           =   1  'Graphical
      TabIndex        =   87
      ToolTipText     =   "To Find Record"
      Top             =   2025
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
      Left            =   12390
      Style           =   1  'Graphical
      TabIndex        =   91
      ToolTipText     =   "Last"
      Top             =   1605
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
      Left            =   11790
      Style           =   1  'Graphical
      TabIndex        =   90
      ToolTipText     =   "First"
      Top             =   1605
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
      Left            =   12390
      Style           =   1  'Graphical
      TabIndex        =   89
      ToolTipText     =   "Next"
      Top             =   1125
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
      Left            =   2805
      MaxLength       =   40
      TabIndex        =   4
      Text            =   "txttemp(2)"
      ToolTipText     =   "Press F3 For Creation"
      Top             =   1380
      Width           =   7095
   End
   Begin VB.TextBox txt 
      Height          =   360
      Index           =   24
      Left            =   17640
      TabIndex        =   102
      Text            =   "24"
      Top             =   4635
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.TextBox txt 
      Height          =   360
      Index           =   9
      Left            =   17550
      TabIndex        =   101
      Text            =   "9"
      Top             =   4140
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.TextBox txt 
      Height          =   360
      Index           =   8
      Left            =   17685
      TabIndex        =   100
      Text            =   "8"
      Top             =   3780
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.TextBox txt 
      Height          =   360
      Index           =   7
      Left            =   1320
      TabIndex        =   99
      Text            =   "7"
      Top             =   5760
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.TextBox txt 
      Height          =   360
      Index           =   3
      Left            =   17685
      TabIndex        =   98
      Text            =   "3"
      Top             =   3285
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.TextBox txt 
      Height          =   360
      Index           =   2
      Left            =   17595
      TabIndex        =   97
      Text            =   "2"
      Top             =   1395
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
      Left            =   2805
      MaxLength       =   60
      TabIndex        =   1
      Text            =   "1"
      Top             =   510
      Width           =   7095
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
      Left            =   11790
      Style           =   1  'Graphical
      TabIndex        =   88
      ToolTipText     =   "Previous"
      Top             =   1125
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
      Left            =   17370
      MaxLength       =   40
      TabIndex        =   96
      Text            =   "0"
      Top             =   990
      Visible         =   0   'False
      Width           =   375
   End
   Begin Crystal.CrystalReport CryReport 
      Left            =   4950
      Top             =   0
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
   Begin VB.Label lbl1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Op. Bal. Without Int  :"
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
      Height          =   240
      Index           =   33
      Left            =   660
      TabIndex        =   162
      Top             =   2760
      Width           =   1905
   End
   Begin VB.Label lblMsOms 
      BackStyle       =   0  'Transparent
      Caption         =   "MS / OMS :"
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
      Left            =   4350
      TabIndex        =   134
      Top             =   2820
      Visible         =   0   'False
      Width           =   1785
   End
   Begin VB.Label Label15 
      BackStyle       =   0  'Transparent
      Caption         =   "Code                      :"
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
      Index           =   7
      Left            =   660
      TabIndex        =   133
      Top             =   1005
      Width           =   2100
   End
   Begin VB.Shape Shape9 
      BackColor       =   &H00C00000&
      BackStyle       =   1  'Opaque
      Height          =   150
      Left            =   360
      Shape           =   3  'Circle
      Top             =   1485
      Width           =   150
   End
   Begin VB.Shape Shape4 
      BackColor       =   &H00C00000&
      BackStyle       =   1  'Opaque
      Height          =   150
      Left            =   360
      Shape           =   3  'Circle
      Top             =   570
      Width           =   150
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
      Left            =   10575
      TabIndex        =   113
      Top             =   0
      Width           =   2475
   End
   Begin VB.Label lbl1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "&Opening Balance      :"
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
      Height          =   240
      Index           =   20
      Left            =   630
      TabIndex        =   64
      Top             =   2280
      Width           =   1905
   End
   Begin VB.Label lbl1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Party Group            :"
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
      Height          =   240
      Index           =   21
      Left            =   690
      TabIndex        =   63
      Top             =   1875
      Width           =   1830
   End
   Begin VB.Label Label15 
      BackStyle       =   0  'Transparent
      Caption         =   "Account &Name       :"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   360
      Index           =   0
      Left            =   690
      TabIndex        =   0
      Top             =   510
      Width           =   1980
   End
   Begin VB.Label Label6 
      BackColor       =   &H00EFB17A&
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
      TabIndex        =   103
      Top             =   0
      Width           =   13545
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00800000&
      FillColor       =   &H00FED8FC&
      FillStyle       =   0  'Solid
      Height          =   7200
      Left            =   11550
      Shape           =   4  'Rounded Rectangle
      Top             =   765
      Width           =   1575
   End
   Begin VB.Label lbl1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "A/c Group Name    :"
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
      Index           =   22
      Left            =   690
      TabIndex        =   62
      Top             =   1425
      Width           =   1860
   End
   Begin VB.Shape Shape2 
      BackColor       =   &H00808080&
      BorderColor     =   &H004D177D&
      FillColor       =   &H00FDF1E3&
      FillStyle       =   0  'Solid
      Height          =   2355
      Left            =   15390
      Top             =   990
      Width           =   1185
   End
   Begin VB.Shape Shape7 
      BorderColor     =   &H00EFB17A&
      BorderWidth     =   5
      FillColor       =   &H00EFB17A&
      Height          =   9630
      Left            =   45
      Top             =   0
      Width           =   13530
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
Public CriteriaStr As String
Const MaxNo As Long = 72
Dim rstbl As Recordset
Dim i As Integer
Dim varBookMark As Variant
Dim Isfind As Boolean
Dim FirstTimeFlag As Boolean
Dim X As String
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
    Cbo(Index).text = GProcProperCase(Trim(Cbo(Index).text))
    If Index = 7 Then
       CityState = GProcGetColumnValue("TblMastAccount", "AcCity", Cbo(7).text, "S", "AcState", "S")
       Cbo(8).text = CityState
    End If
    If Index = 7 Then
       CityDist = GProcGetColumnValue("TblMastAccount", "AcCity", Cbo(7).text, "S", "AcDist", "S")
       Cbo(9).text = CityDist
    End If
    If Index = 7 Then
       CityZip = GProcGetColumnValue("TblMastAccount", "AcCity", Cbo(7).text, "S", "AcZip", "S")
       txt(10).text = CityZip
    End If

End If
End Sub
Private Sub cboBalance_LostFocus()
If cboBalance.ListIndex = 0 Then
    txtOpBal(1) = "Cr"
Else
    txtOpBal(1) = "Dr"
End If
End Sub
Private Sub cmdBtn_Click(Index As Integer)
'On Error GoTo ErrorRoutine
Select Case Index
Case 0 '---- Previous
     If GProcRstOpen(rstbl, "Select * from " & tblName & " where " & OrderField & "  = (Select max(" & OrderField & " ) from " & tblName & " where  " & OrderField & " < '" & txt(1) & "' and " & CriteriaStr & " )and " & CriteriaStr, "R") > 0 Then ReadFields Else MsgBox "This is First Record:", vbOKOnly + vbInformation
Case 1 '---- Next
      If GProcRstOpen(rstbl, "Select * from " & tblName & " where " & OrderField & " = (Select min(" & OrderField & ") from " & tblName & " where " & OrderField & " > '" & txt(1) & "' and " & CriteriaStr & ")and " & CriteriaStr, "R") > 0 Then ReadFields Else MsgBox "This is last Record:", vbOKOnly + vbInformation
Case 2 '---- First
         If GProcRstOpen(rstbl, "Select * from " & tblName & " where " & OrderField & " = (Select min(" & OrderField & ") from " & tblName & " where " & OrderField & " <> '' and " & CriteriaStr & ")and " & CriteriaStr, "R") > 0 Then ReadFields
Case 3 '---- Last
        If GProcRstOpen(rstbl, "Select * from " & tblName & " where " & OrderField & " = (Select max(" & OrderField & ") from " & tblName & " where " & OrderField & " <> '' and " & CriteriaStr & ")and " & CriteriaStr, "R") > 0 Then ReadFields Else GProcClearForm Me, rstbl, MaxNo
Case 4 'Find ' QryMastAccount
     Isfind = True
     txt(44).Enabled = True
     txt(44).SetFocus
     Call GProcSetButtons(Me, vbDataActionCancel)
     Call GProcClearForm(Me, rstbl, MaxNo, True)
     Cbo(7) = "" '-City
     Cbo(8) = "" '- State
     Cbo(9) = "" '-Dist
     
     
'     gClsSearch.SearchMultiFieldAccMaster "", txt(1).Left + Me.Left, txt(1).Top + Me.Top + 650
'     If gClsSearch.SearchMultiRetColAccMaster(0) <> "" Then
'        GProcRstOpen rstbl, "Select * from " & tblName & " where " & OrderField & " = '" & gClsSearch.SearchMultiRetColAccMaster(0) & "' and " & CriteriaStr, "R"
'        gClsSearch.SearchMultiRetColAccMaster(0) = ""
'        ReadFields
'    End If
Case 5 'Mill A/c List
    On Error GoTo ErrorHandler '--- Outstanding List
    gSelectedMenu = "Mill Account List"
    GProcCrystalRptPreparation cryReport, "Window"
    cryReport.Formulas(4) = ""
    cryReport.Formulas(5) = ""
   ' CryReport.SelectionFormula = "{tblOutStanding.VYear} = " & gCYear - 1 & " and ({tblOutStanding.Vtype}='SO' or {tblOutStanding.Vtype}='SY' or {tblOutStanding.Vtype}='ST'  or   {tblOutStanding.Vtype}='SD' or  {tblOutStanding.Vtype}='SM')"
    cryReport.ReportFileName = gReportPath & "RptLstMillAccount.rpt"
    cryReport.Action = 1
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
 '       txt(1).SetFocus
    End If
    FillCombo
    Cbo(7) = "" '-City
    Cbo(8) = "" '- State
    Cbo(9) = "" '-Dist
Case 7 'Modify
    'If mURecModi = True Then
    '    SearchHelp
    '    DoEvents
    '    cmdBtn_Click 4 '-- Find
     '   If gClsSearch.SearchMultiRetCol(0) = "" Then Exit Sub
        FormAction = vbDataActionUpdate
        Call GProcActivateControls(Me, True)
        Call GProcSetButtons(Me, vbDataActionUpdate)
        SetCtrlEd
        txt(1).Enabled = True
        txt(1).SetFocus
    'End If
Case 8 'Delete
    'If mURecDel = True Then
     'If IsRelatedRecord = False Then
        If MsgBox("Do You Really Delete This Record?", vbQuestion + vbYesNo + vbDefaultButton1, "Confirm Delete") = vbYes Then
            gCn.Execute "delete from tblVoucher where Vtype='0P' and  (VCrAcCode=" & Val(txt(0)) & " or VDrAcCode=" & Val(txt(0)) & ") and VYear=" & gCYear & " and Vfirm ='" & gCCode & "'"
            gCn.Execute "Delete from " & tblName & " where " & IdField & " =  " & txt(0) & " and " & CriteriaStr, "O"
        End If
        cmdBtn_Click (3)
     'End If
Case 9 'Print
    Dim TmpRstbl As Recordset
    GProcRstOpen rstbl, "Select * from " & tblName & " ", "R"
    rstbl.MoveFirst
    With rstbl
       Do While Not rstbl.EOF
'          If rstbl.Fields("Acname") = "Shital Sujit Bhoje" Then
            FormAction = vbDataActionUpdate
            ReadFields
            cmdBtn_Click (13)
            cmdBtn_Click (1)
 '         End If
'          rstbl.MoveNext
       Loop
    End With

Case 10 'Exit
    gRstAccMast.Requery
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
        GProcRstOpen rstbl, "Select * from " & tblName & " where " & IdField & " =  " & txt(0) & " and " & CriteriaStr, "O"
        GProcSaveRecord Me, rstbl, FormAction, MaxNo
        DoEvents
        DoEvents
        DoEvents
        DoEvents
        WriteFields
'        gSearchMultiRetCol(0) = txt(1) '--AName
        FormAction = vbDataActionClose
        Call GProcActivateControls(Me, False)
        Call GProcSetButtons(Me, True)
        If FormAction = vbDataActionAddNew Then
            If cmdBtn(6).Visible = True Then cmdBtn(6).SetFocus
        Else
            If cmdBtn(7).Visible = True Then cmdBtn(7).SetFocus
        End If
'        If rstbl.Fields("AgCode") = 90017 Then
'           gRstAccMast.Requery
 '       End If
    End If
Case 14 'Cancel
'    Call GProcClearForm(Me, rstbl, MaxNo, True)
    cmdBtn_Click (3)
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
    CriteriaStr = " Accode > 1"
    If Trim(CriteriaStr) <> "" Then X = X & " where " & CriteriaStr
    If Trim(OrderField) <> "" Then X = X & " order by " & OrderField
'    GProcRstOpen rstbl, x, "O"
    If Trim(mShowRec) <> "" Then
       GProcRstOpen rstbl, "Select * from " & tblName & " where " & IdField & " = " & CLng(Trim(mShowRec)), "R"
       ReadFields
       Call GProcSetButtons(Me, vbDataActionCancel)
       mShowRec = ""
    Else
       GProcRstOpen rstbl, "Select * from " & tblName & " where " & IdField & " = 0", "R"
       cmdBtn_Click (6)
    End If
    FirstTimeFlag = True
    '---- Button Visiblity Set As Per User Previlage
    GProcSetButtonVisiblity Me
    cboBalance.ListIndex = 0
    txtOpBal(1) = "Cr"
    SSTab1.Tab = 0
    cmdBtn_Click (6)
    Cbo(24).ListIndex = 0
    Cbo(33).ListIndex = 0
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
GProcCheckForNumber Me, rstbl, MaxNo, Array(18, 15, 22, 44, 59, 60), Array(23, 36, 59, 60)
End Sub
Private Sub txt_LostFocus(Index As Integer)
CheckForNumber
SaveBtnEd
txtOpBal(0).text = Format(txtOpBal(0), FStr)
txtOpBal(2).text = Format(txtOpBal(2), FStr)
If Index = 42 Then
   SSTab1.Tab = 1
   txttemp(17).SetFocus
End If
If Index = 19 Then
   If Cbo(33).Enabled = False Then
      SSTab1.Tab = 3
      txt(15).SetFocus
   Else
      SSTab1.Tab = 2
      Cbo(33).SetFocus
   End If
End If
If Index = 30 Then
   SSTab1.Tab = 3
   txt(15).SetFocus
End If
If Index = 13 And txt(14) = "" Then
   txt(14) = txt(13)
End If
If Index = 14 And txt(13) = "" Then
   txt(13) = txt(14)
End If
'If txttemp(2) = "Sundry Debtors" Or txttemp(2) = "Brokers ( Sale Commission Agents)" Then
'   Cbo(7).text = "Ichalkaranji"
'End If
If Index = 44 And FormAction = vbDataActionAddNew Then
   X = GProcGetColumnValue("tblMastAccount", "AcOurCode", UCase(txt(44)), "S", "AcName", "S")
   If X <> "" Then
      MsgBox "Code Exist For A/s Name " + X, vbCritical + vbOKOnly, Me.Caption
      txt(44).SetFocus
   End If
End If
If Index = 44 And Isfind = True Then
   Isfind = False
   txt(1) = GProcGetColumnValue("tblMastAccount", "AcOurCode", UCase(txt(44)), "S", "AcName", "S")
   If txt(1) <> "" Then
      GProcRstOpen rstbl, "Select * from " & tblName & " where " & OrderField & " = '" & txt(1) & "' and " & CriteriaStr, "R"
      ReadFields
      Call GProcSetButtons(Me, vbDataActionCancel)
      For i = 0 To 7
         cmdBtn(i).Enabled = True
      Next i
   Else
      If txt(44) <> "" Then
         MsgBox ("Code Not Found...")
      End If
      gClsSearch.SearchMultiFieldAccMaster "", txt(1).Left + Me.Left, txt(1).Top + Me.Top + 650
      
      If gClsSearch.SearchMultiRetColAccMaster(0) <> "" Then
         GProcRstOpen rstbl, "Select * from " & tblName & " where " & OrderField & " = '" & gClsSearch.SearchMultiRetColAccMaster(0) & "' and " & CriteriaStr, "R"
         gClsSearch.SearchMultiRetColAccMaster(0) = ""
         ReadFields
         Call GProcSetButtons(Me, vbDataActionCancel)
         For i = 0 To 8
             cmdBtn(i).Enabled = True
         Next i
      
      End If
   End If
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
If IsNumeric(txtOpBal(2)) = False Then txtOpBal(2) = "0.00"
txtOpBal(0).text = Format(txtOpBal(0), FStr)
txtOpBal(2).text = Format(txtOpBal(2), FStr)
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
        GProcShowForm frmMastGroup, frmMain.mnuMstfrm(1), Array(frmMastGroup.txttemp(2)), Array(GName)
End Select
End If
End Sub
Private Sub txtTemp_KeyPress(Index As Integer, KeyAscii As Integer)
Dim LeftPos As Long
Dim TopPos As Long
TopPos = txttemp(Index).Top + Me.Top + 650
LeftPos = txttemp(Index).Left + Me.Left
If KeyAscii = 13 Then Exit Sub '-- Enter
'-- A/c Group
Select Case Index
Case 2, 3
'    gClsSearch.SearchMultiField "tblMastGroup", "AGName,Agcode", Array("Account Group", "Agcode"), Array(txttemp(Index).Width), "", Chr(KeyAscii), "AgName", LeftPos, TopPos, True, , False
    gClsSearch.SearchMultiField "tblMastGroup", "AGName,Agcode", Array("Account Group", "Agcode"), Array(txttemp(Index).Width, 0), " Not(Agcode = TopGroup)  ", Trim(Chr(KeyAscii)), "", LeftPos, TopPos, False, , False
    KeyAscii = 0
'-- Company Group
Case 16
    gClsSearch.SearchMultiField "tblMastCompGroup", "CGName", Array("Company Group"), Array(txttemp(Index).Width), " ", Chr(KeyAscii), "CGName", LeftPos, TopPos, True
    KeyAscii = 0
'-- Broker
Case 17
    gClsSearch.SearchMultiField "tblMastAccount", "AcName", Array("Broker"), Array(txttemp(Index).Width), " (AgCode=90016) ", Chr(KeyAscii), "AcName", LeftPos, TopPos
    KeyAscii = 0
Case 43, 50, 51
    gClsSearch.SearchMultiField "tblMastCompany", "CName", Array("Firm Name"), Array(txttemp(Index).Width), "", Chr(KeyAscii), "CName", LeftPos, TopPos
    KeyAscii = 0
'-- Depot Ho A/c
Case 62, 63
'    gClsSearch.SearchMultiField "tblMastAccount", "AcName", Array("Broker"), Array(txtTemp(Index).Width), "", Chr(KeyAscii), "AcName", LeftPos, TopPos
    gClsSearch.SearchMultiFieldAccMaster Chr(KeyAscii), txttemp(Index).Left + Me.Left, txttemp(Index).Top + Me.Top + 1000
    KeyAscii = 0
End Select
If Index = 62 Or Index = 63 Then
   If gClsSearch.SearchMultiRetColAccMaster(0) <> "" Then txttemp(Index) = gClsSearch.SearchMultiRetColAccMaster(0)
Else
   If gClsSearch.SearchMultiRetCol(0) <> "" Then txttemp(Index) = gClsSearch.SearchMultiRetCol(0)
End If
Set gClsSearch = Nothing
'KeyAscii = 0
End Sub
Private Sub SetCtrlEd()
Dim AgCode As Long
AgCode = GProcGetMainGroupCode(txttemp(2))
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
'-- Mill,Cr,Dr
'--OS Limit
If AgCode = 90029 Or AgCode = 90015 Or AgCode = 90017 Then
    lblMsOms.Visible = True
    Cbo(45).Visible = True
    txt(31).Visible = True
    lblOSLimit.Visible = True
    If FormAction = vbDataActionAddNew Or FormAction = vbDataActionUpdate Then
        txt(31).Enabled = True
    Else
        txt(31).Enabled = False
    End If
    chk(21).Caption = "Don't Show Outstanding ?"
Else
    lblOSLimit.Visible = False
    txt(31).Visible = False
    txt(31) = 0
    chk(21).Caption = "Not Required In Interest Calculation"
End If
'-- Mill Group
lbl1(26).Caption = "Remark               :"
If GProcGetMainGroupCode(txttemp(2)) = 90029 Then
   lbl1(26).Caption = "Godown   :"
    If FormAction = vbDataActionAddNew Or FormAction = vbDataActionUpdate Then
'        cbo(33).Enabled = True '-- Mill Type
        Cbo(24).Enabled = True '-- commission On
        Cbo(35).Enabled = True '-- commission Type
        txt(25).Enabled = True '--Comm Rate
        txt(47).Enabled = True '--Cash Disc Rate
        chk(26).Enabled = True '--Is Local Brk by self
        chk(27).Enabled = True '-- Is TDS Pay
        chk(28).Enabled = True '--Service tax Pay
        chk(46).Enabled = True '--is payment to our main firm
        txttemp(43).Enabled = True  '-- Mill Bill In firm
        txttemp(50).Enabled = True  '-- Consignment In firm
        txttemp(51).Enabled = True  '-- Depot In Firm firm
        If IsMillPurchaseSale = False Then
            chk(29).Enabled = True '-- Exmpt Mill
            txt(30).Enabled = True '-- Exmpt Rate
        Else
            chk(29).Enabled = False '-- Exemp Mill
            txt(30).Enabled = False '-- Exem Rate
        End If
   End If
Else
    Cbo(33).Enabled = False '-- Mill Type
    Cbo(35).Enabled = False '-- Comm Type
    Cbo(33).ListIndex = 0
    Cbo(24).Enabled = False
    txt(25).Enabled = False
    txt(47).Enabled = False '--Cash Disc Rate
    chk(26).Enabled = False
    chk(27).Enabled = False
    chk(28).Enabled = False
    chk(29).Enabled = False
    chk(46).Enabled = False '--is payment to our main firm
    txt(30).Enabled = False
    txttemp(43).Enabled = False  '-- Mill Bill In firm
    txttemp(50).Enabled = False  '-- Consignment In firm
    txttemp(51).Enabled = False  '-- Depot In Firm firm
    
    txt(25) = "0.00"
    txt(30) = "0.00"
    chk(26) = 0
    chk(27) = 0
    chk(28) = 0
    chk(29) = 0
    chk(46) = 0
    txttemp(43) = ""  '-- Main firm
End If
If AgCode = 90008 Then
   lbl1(0).Caption = "Depreciation % :"
Else
   lbl1(0).Caption = "Interest %        :"
End If
txtOpBal(1).Enabled = False

If AgCode = 90016 Then
   chk(55).Caption = "No Brokerage On Sale"
'   chk(55).Enabled = True
Else
  chk(55).Caption = "Is Interest Debit Note Monthly ?"
End If
End Sub
Private Function IsMillPurchaseSale() As Boolean

'--- Pending
'Dim X As String
'Dim Rs1 As Recordset
'X = "Select * from tblAddLess where AdMillCode=" & txt(0)
'GProcRstOpen Rs1, X, "R"
'If Rs1.EOF And Rs1.BOF Then
'Else
'    IsMillPurchaseSale = True
'    Exit Function
'End If
'Rs1.Close
'--end

End Function
Private Sub txtTemp_LostFocus(Index As Integer)
Dim TopGroup As Long
If Index = 2 And txttemp(3) = "" Then txttemp(3) = txttemp(2)
SaveBtnEd
SetCtrlEd
If IsNumeric(txtOpBal(0)) = False Then txtOpBal(0) = "0.00"
If IsNumeric(txtOpBal(2)) = False Then txtOpBal(2) = "0.00"
If txtOpBal(0) = 0 Then
    TopGroup = GProcGetColumnValue("TblMastGroup", "AgName", txttemp(2), "S", "TopGroup", "N")
    If TopGroup = 90001 Then
        cboBalance.ListIndex = 1
        txtOpBal(1) = "Dr"
    Else
        cboBalance.ListIndex = 0
        txtOpBal(1) = "Cr"
    End If
End If
End Sub
Private Sub FillCombo()
Call GProcFillCombo(Cbo(7), "select Distinct AcCity from " & tblName & " where AcCity <> '' order by AcCity")
Call GProcFillCombo(Cbo(8), "select Distinct AcState from " & tblName & " where AcState <> '' order by AcState")
Call GProcFillCombo(Cbo(9), "select Distinct AcDist from " & tblName & " where AcDist <> '' order by AcDist")
End Sub
Private Sub SaveBtnEd()
'--- AcName,AgCode
If Trim(txt(1)) = "" Or Trim(txttemp(2)) = "" Then
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
    
    X = GProcGetColumnValue("tblMastGroup", "AgName", txttemp(2), "S", "TopGroup", "N")
    Y = GProcGetColumnValue("tblMastGroup", "AgName", txttemp(3), "S", "TopGroup", "N")
    If X = 90003 Or Y = 90003 Then
       If X <> Y Then
          MsgBox "Check Second Group Name.", vbCritical + vbOKOnly, Me.Caption
          txttemp(3).SetFocus
          Exit Function
       End If
    End If
    If Val(txt(61)) > 0 Then
        X = "Select * From " & tblName & " Where Acsrno = " & Trim(txt(61).text) & " and AcCode <> " & Trim(txt(0).text) & "   "
        i = GProcRstOpen(rstAddCheck, X, "R")
        If i > 0 Then
            MsgBox "Record is Available of A/c Sr.No Name  " & Trim(txt(61).text), vbCritical + vbOKOnly, "Duplicate Record"
            txt(61).SetFocus
            Exit Function
        End If
       X = GProcGetColumnValue("TblMastAccount", "AcBST", txt(13).text, "S", "AcName", "S", "AcCode <> " & txt(0) & "")
       If X <> "" Then
            MsgBox "TIN No. Exist For A/s Name " + X, vbCritical + vbOKOnly, Me.Caption
'            txt(13).SetFocus
 '           Exit Function
       End If
    
    End If
    
    
    '----Required
    '--- Account Name
    If Trim(txt(1)) = "" Then
        MsgBox "Check Account Name.", vbCritical + vbOKOnly, Me.Caption
        txt(1).SetFocus
        Exit Function
    End If
    '---Group
    If Trim(txttemp(2)) = "" Then
        MsgBox "Check Group Name.", vbCritical + vbOKOnly, Me.Caption
        txttemp(2).SetFocus
        Exit Function
    End If
    '---Create Group
    If GProcGetColumnValue("tblMastGroup", "AgName", txttemp(2), "S", "AgCode", "N") = 0 Then
        GProcShowForm frmMastGroup, frmMain.mnuMstfrm(1), Array(frmMastGroup.txt(1)), Array(txttemp(2))
        Exit Function
    End If
    '---Group 2
    
'---Temp Cheng for 2nd grp = main grp
'    txtTemp(3) = txtTemp(2)
'-- Original    If Trim(txttemp(3)) = "" Then txttemp(3) = txttemp(2)
    '---Create Group
    If GProcGetColumnValue("tblMastGroup", "AgName", txttemp(3), "S", "AgCode", "N") = 0 Then
        GProcShowForm frmMastGroup, frmMain.mnuMstfrm(1), Array(frmMastGroup.txt(1)), Array(txttemp(3))
        Exit Function
    End If
    '--- Company Group
    If txttemp(16) = "" Then txttemp(16) = "-"
'    If Trim(txttemp(16)) = "" Then
'        MsgBox "Check Company Group Name.", vbCritical + vbOKOnly, Me.Caption
'        txttemp(16).SetFocus
'        Exit Function
'    End If
    '---Create comp Group
    If GProcGetColumnValue("tblMastCompGroup", "CgName", txttemp(16), "S", "CgCode", "N") = 0 Then
        GProcShowForm frmMastCompGroup, frmMain.mnuMstfrm(2), Array(frmMastCompGroup.txt(1)), Array(txttemp(16))
        Exit Function
    End If
    '-- Mill Group Commission
    If GProcGetMainGroupCode(txttemp(2)) = 29 Then
        If CDbl(txt(25)) = 0 And (Cbo(33).ListIndex = 1 Or Cbo(33).ListIndex = 2 Or Cbo(33).ListIndex = 3) Then
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
    
    If Cbo(33).ListIndex = 3 And txttemp(43).text = "" Then
        MsgBox "Check Main Firm Name", vbCritical + vbOKOnly, Me.Caption
        txttemp(43).SetFocus
        Exit Function
    End If
    
    '---BST TIN 11 chr
'    If txt(13) <> "" And Len(txt(13)) <> 11 Then
'        MsgBox "Check TIN No. Chr. Must be 11", vbCritical + vbOKOnly, Me.Caption
'        txt(13).SetFocus
'        Exit Function
'    End If
'    If txt(14) <> "" And Len(txt(14)) <> 11 Then
'        MsgBox "Check TIN No. Chr. Must be 11", vbCritical + vbOKOnly, Me.Caption
'        txt(14).SetFocus
'        Exit Function
'    End If
    '--Dublicate Tin
    If txt(13) <> "" And Left(txt(13), 1) <> "U" Then
       X = GProcGetColumnValue("TblMastAccount", "AcBST", txt(13).text, "S", "AcName", "S", "AcCode <> " & txt(0) & "")
       If X <> "" Then
            MsgBox "TIN No. Exist For A/s Name " + X, vbCritical + vbOKOnly, Me.Caption
'            txt(13).SetFocus
 '           Exit Function
       End If
    End If
    If txt(14) <> "" And Left(txt(13), 1) <> "U" Then
       X = GProcGetColumnValue("TblMastAccount", "AcCST", txt(14).text, "S", "AcName", "S", "AcCode <> " & txt(0) & "")
       If X <> "" Then
            MsgBox "TIN No. Exist For A/s Name " + X, vbCritical + vbOKOnly, Me.Caption
'            txt(14).SetFocus
 '           Exit Function
       End If
    End If
    If txt(59) <> "" And Left(txt(59), 1) <> "U" Then
       X = GProcGetColumnValue("TblMastAccount", "AcGSTIN", txt(59).text, "S", "AcName", "S", "AcCode <> " & txt(0) & "")
       If X <> "" Then
            MsgBox "GSTIN No. Exist For A/s Name " + X, vbCritical + vbOKOnly, Me.Caption
'            txt(14).SetFocus
 '           Exit Function
       End If
    End If
    
    ''--Dublicate Our Code
    If txt(44) <> "" Then
       X = GProcGetColumnValue("TblMastAccount", "AcOurCode", txt(44).text, "S", "AcName", "S", "AcCode <> " & txt(0) & "")
       If X <> "" Then
            MsgBox "Code Exist For A/s Name " + X, vbCritical + vbOKOnly, Me.Caption
            txt(44).SetFocus
            Exit Function
       End If
    End If
    '--Dublicate AcName
    If txt(1) <> "" Then
       X = GProcGetColumnValue("TblMastAccount", "AcName", txt(1).text, "S", "AcName", "S", "AcCode <> " & txt(0) & "")
       If X <> "" Then
            MsgBox "Dublicate A/c Name  ", vbCritical + vbOKOnly, Me.Caption
            txt(1).SetFocus
            Exit Function
       End If
    End If
    
    
    If chk(29) = 0 And CDbl(txt(30)) > 0 Then txt(30) = 0
    '-- Is Exemption Mill
    If chk(29) = 1 And CDbl(txt(30)) = 0 Then
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
txt(2) = GProcGetColumnValue("tblMastGroup", "AgName", txttemp(2), "S", "AgCode", "N")
txt(3) = GProcGetColumnValue("tblMastGroup", "AgName", txttemp(3), "S", "AgCode", "N")
 '--Com Group
If txttemp(16) = "" Then txttemp(16) = "-"
txt(16) = GProcGetColumnValue("tblMastCompGroup", "CgName", txttemp(16), "S", "CgCode", "N")
'--Broker
If txttemp(17) = "" Then txttemp(17) = "-"
txt(17) = GProcGetColumnValue("tblMastAccount", "AcName", txttemp(17), "S", "ACCode", "N")
'--Mill Bill In Firm
If txttemp(43) = "" Then
   txt(43) = ""
Else
   txt(43) = GProcGetColumnValue("tblMastCompany", "CName", txttemp(43), "S", "CCode", "N")
End If
If txttemp(62) = "" Then
   txttemp(62) = "-"
End If
txt(62) = GProcGetColumnValue("tblMastAccount", "AcName", txttemp(62), "S", "ACCode", "N")
If txttemp(63) = "" Then
   txttemp(63) = "-"
End If
txt(63) = GProcGetColumnValue("tblMastAccount", "AcName", txttemp(63), "S", "ACCode", "N")

'--Consignment In Firm
If txttemp(50) = "" Then
   txt(50) = ""
Else
   txt(50) = GProcGetColumnValue("tblMastCompany", "CName", txttemp(50), "S", "CCode", "N")
End If
'--Depot In Firm
If txttemp(51) = "" Then
   txt(51) = ""
Else
   txt(51) = GProcGetColumnValue("tblMastCompany", "CName", txttemp(51), "S", "CCode", "N")
End If


txt(7) = Cbo(7)
txt(8) = Cbo(8)
txt(9) = Cbo(9)
txt(24) = Cbo(24).ListIndex
txt(33) = Cbo(33).ListIndex '--Mill Type
txt(35) = Cbo(35).ListIndex '-- Comm Type
txt(45) = Cbo(45).ListIndex '-- MS/OMS party
txt(20) = chk(20)
txt(21) = chk(21)
txt(26) = chk(26)
txt(27) = chk(27)
txt(28) = chk(28)
txt(29) = chk(29)
txt(46) = chk(46) '--Is Payment to Main Firm
txt(52) = chk(52) '--Due Date Not Print In Sale Bill
txt(53) = chk(53) '--Is Tax Paid By Mill
txt(54) = chk(54) '--Is Corporate party
txt(55) = chk(55) '--Is Int DbNt Monthly
txt(56) = chk(56) '--Less Brokerage in Comm
txt(57) = chk(57) '--Less Unloading in comm
txt(58) = chk(58) '--Less GST in Comm
txt(65) = chk(65) '--Partywise detail req
txt(66) = chk(66) '--Loan detail req
txt(67) = chk(67) '--Asset detail req
txt(68) = chk(68) '--Partywise detail req for audit
txt(69) = chk(69)  '--Exp Sub A/c detail req for audit
txt(70) = chk(70)  '--Less TCS in Comm
txt(71) = chk(71)  '--Less TCS in Comm
txt(72) = chk(72)  '--TDS from 1st Bill

'txt(32) = Chk(32)
'txt(34) = Chk(34) '--Is Db NT A/c
End Sub
Private Sub FillTempFromTxt()
txttemp(2) = GProcGetColumnValue("tblMastGroup", "AgCode", txt(2), "N", "AgName", "S") '--Group
txttemp(3) = GProcGetColumnValue("tblMastGroup", "AgCode", txt(3), "N", "AgName", "S")
txttemp(16) = GProcGetColumnValue("tblMastCompGroup", "CGCode", txt(16), "N", "CgName", "S") '--Com Group
txttemp(17) = GProcGetColumnValue("tblMastAccount", "AcCode", txt(17), "N", "AcName", "S") '--Broker
If Val(txt(62)) > 0 Then
   txttemp(62) = GProcGetColumnValue("tblMastAccount", "AcCode", txt(62), "N", "AcName", "S")
Else
   txttemp(62) = ""
End If
If Val(txt(63)) > 0 Then
   txttemp(63) = GProcGetColumnValue("tblMastAccount", "AcCode", txt(63), "N", "AcName", "S")
Else
   txttemp(63) = ""
End If
If txt(43) = "" Then
   txttemp(43) = ""
Else
   txttemp(43) = GProcGetColumnValue("tblMastCompany", "CCode", txt(43), "S", "CName", "S") '--Mill Bill In  firm
End If
If txt(50) = "" Then
   txttemp(50) = ""
Else
   txttemp(50) = GProcGetColumnValue("tblMastCompany", "CCode", txt(50), "S", "CName", "S") '--Consignment In firm
End If
If txt(51) = "" Then
   txttemp(51) = ""
Else
   txttemp(51) = GProcGetColumnValue("tblMastCompany", "CCode", txt(51), "S", "CName", "S") '--Depot In firm
End If

Cbo(7) = txt(7) '--City
Cbo(8) = txt(8) '-state
Cbo(9) = txt(9) '--District
Cbo(24).ListIndex = txt(24)  '--comm On
Cbo(33).ListIndex = txt(33)  '--Mill Type
Cbo(35).ListIndex = txt(35)  '--comm Type
Cbo(45).ListIndex = txt(45)  '--MS/OMS

chk(20) = txt(20)
chk(21) = txt(21)
chk(26) = txt(26)
chk(27) = txt(27)
chk(28) = txt(28)
chk(29) = txt(29)
chk(46) = txt(46) '--Is Payment to Main Firm
chk(52) = txt(52) '--Is Payment to Main Firm
chk(53) = txt(53) '--Is Tax Paid By Mill
chk(54) = txt(54) '--Is Corporate party
chk(55) = txt(55) '--Is Int DbNt Monthly
chk(56) = txt(56) '--Less Brokerage in Comm
chk(57) = txt(57) '--Less Unloading in Comm
chk(58) = txt(58) '--Less GST in comm
chk(65) = txt(65) '--Party wise detail req
chk(66) = txt(66) '--Loan Det req
chk(67) = txt(67) '--Asset detail req
chk(68) = txt(68) '--Party wise detail req for audit
chk(69) = txt(69) '--Exp Sub A/c detail req for audit
chk(70) = txt(70) '--Less TCS in comm
chk(70) = txt(70) '--Less TCS in comm
chk(71) = txt(71) '--TDS Party
chk(72) = txt(72) '--TDS from 1st Bill

'Chk(34) = txt(34) '--Is Db NT A/c
'txtOpBal(1) = "Cr"
cboBalance.ListIndex = 0
txtOpBal(1) = "Cr"


'---Pending for opbal srk
'txtOpBal(0) = GProcGetColumnValue("tblVoucher", "VAcCode", txt(0), "N", "VAmt", "N", "vtype = '0P' and Vyear=" & gCYear, gCn)  '--Balance



'If (txtOpBal(0) < 0) Then
'   txtOpBal(0) = Format(Abs(Val(txtOpBal(0))), FStr)
'   txtOpBal(1) = "Dr"
'   cboBalance.ListIndex = 1
'End If
'txtOpBal(0) = Format(txtOpBal(0), FStr)

'--end
SetCtrlEd
End Sub
Private Sub WriteFields()
Dim tmpAccode As Long, Vno As Long
Dim X As String
Dim CmdPara As String
tmpAccode = GProcGetColumnValue("TblMastAccount", "AcName", "-", "S", "AcCode", "N")
gCn.Execute "delete from tblVoucher where Vtype='0P' and  (VCrAcCode=" & Val(txt(0)) & " or VDrAcCode=" & Val(txt(0)) & ") and VYear=" & gCYear & " and Vfirm ='" & gCCode & "'"
If CDbl(txtOpBal(0)) > 0 Then
   Vno = GProcGenerateId("tblvoucher", "vno", "vtype = '0P'  and VYear=" & gCYear & " and vfirm = '" & gCCode & "'")
   X = "VFirm,VYear,VType,VNo,VCtrNo,VDt,VCrAcCode,VDrAcCode,VAmt,Interest"

   CmdPara = "'" & gCCode & "'," & gCYear & ",'0P'," & Val(Vno) & ",1,'" & Format(gCYSDate, "yyyy/mm/dd") & "'," & IIf(cboBalance = "Cr", Val(txt(0)), Val(tmpAccode)) & "," & _
                  IIf(cboBalance = "Dr", txt(0), tmpAccode) & "," & Val(txtOpBal(0)) & "," & Val(txtOpBal(2)) & ""

   gCn.Execute "Insert into TblVoucher(" & X & ") values(" & CmdPara & ")"
   'GProcInsertRecord "TblVoucher", x, CmdPara, Vno, "select max(vno) from tblvoucher where vtype = '0P'  and VYear=" & gCYear & " and vfCode = " & gCCode
End If




'Dim mRst1 As Recordset, mvno As Long, AcOCode As Long
'AcOCode = GProcGetColumnValue("TblMastAccount", "AcName", "-", "S", "AcCode", "N")
'If IsNumeric(txtOpBal(0)) = False Then txtOpBal(0) = "0.00"
'If CDbl(txtOpBal(0)) = 0 Then
'    gCn.Execute "delete from tblVoucher where Vtype='0P' and Vfirm = '" & gCCode & "' and VAcCode=" & txt(0) & " and VYear=" & gCYear
'Else
'    If (GProcRstOpen(mRst1, "select * from tblvoucher where vtype = '0P' and vAccode = " & txt(0) & " and Vfirm = '" & gCCode & "' and VYear=" & gCYear & " Order by vno", "R", gCn) > 0) Then
'    '    mRst1!vamt = IIf(cboBalance = "Cr", txtOpBal(0), Val(txtOpBal(0)) * (-1))
'    '    mRst1.Update
'        gCn.Execute "update tblvoucher set Vamt=" & CDbl(IIf(cboBalance = "Cr", txtOpBal(0), Val(txtOpBal(0)) * (-1))) & " where vtype = '0P' and vAccode = " & txt(0) & " and VYear=" & gCYear
'    Else
'        mvno = GProcGenerateId("tblvoucher", "vno", "vtype = '0P' and VYear=" & gCYear, gCn)
'        GProcCreateVoucher FormAction, mvno, 1, gCYSDate, "0P", txt(0), AcOCode, IIf(cboBalance = "Cr", txtOpBal(0), Val(txtOpBal(0)) * (-1))
'    End If
'    If mRst1.State = adStateOpen Then mRst1.Close
'End If
End Sub
Private Sub ReadFields()
Dim RstFNRW As Recordset
GProcShowRecord Me, rstbl, MaxNo
FillTempFromTxt
txtOpBal(1) = "Cr"
cboBalance.ListIndex = 0
If GProcRstOpen(RstFNRW, "select VCrAcCode,VDrAcCode,Vamt,Interest from tblVoucher where vtype = '0P' and (VCrAcCode= " & txt(0) & " or VDrAcCode= " & txt(0) & ")  And Vyear = " & gCYear & " and vfirm ='" & gCCode & "'", "R") > 0 Then
   txtOpBal(0) = RstFNRW.Fields(2)
   txtOpBal(2) = RstFNRW.Fields(3)
   If txt(0) = RstFNRW.Fields(1) Then
      txtOpBal(1) = "Dr"
      cboBalance.ListIndex = 1
   End If
Else
   txtOpBal(1) = "Cr"
   txtOpBal(0) = "0.00"
   txtOpBal(2) = "0.00"
End If
SetCtrlEd
End Sub
