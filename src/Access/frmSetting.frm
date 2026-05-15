VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form frmSetting 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00C2E3FC&
   BorderStyle     =   0  'None
   Caption         =   "Setting Entry"
   ClientHeight    =   7785
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   10290
   Icon            =   "frmSetting.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7785
   ScaleWidth      =   10290
   ShowInTaskbar   =   0   'False
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
      Index           =   29
      Left            =   7095
      MaxLength       =   40
      TabIndex        =   30
      Text            =   "txttemp(29)"
      ToolTipText     =   "Press F3 for Creation"
      Top             =   7155
      Width           =   2685
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
      Index           =   28
      Left            =   7095
      MaxLength       =   40
      TabIndex        =   29
      Text            =   "txttemp(28)"
      ToolTipText     =   "Press F3 for Creation"
      Top             =   6795
      Width           =   2685
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
      Index           =   29
      Left            =   0
      TabIndex        =   92
      Text            =   "29"
      Top             =   495
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
      Index           =   28
      Left            =   0
      TabIndex        =   91
      Text            =   "28"
      Top             =   0
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
      Index           =   27
      Left            =   9840
      TabIndex        =   90
      Text            =   "27"
      Top             =   6600
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
      Index           =   26
      Left            =   9840
      TabIndex        =   89
      Text            =   "26"
      Top             =   6240
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
      Index           =   25
      Left            =   9840
      TabIndex        =   88
      Text            =   "25"
      Top             =   5880
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
      Index           =   24
      Left            =   9840
      TabIndex        =   87
      Text            =   "24"
      Top             =   5520
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
      Index           =   26
      Left            =   5370
      MaxLength       =   40
      TabIndex        =   10
      Text            =   "txttemp(26)"
      ToolTipText     =   "Press F3 for Creation"
      Top             =   1320
      Width           =   1365
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
      Index           =   27
      Left            =   2310
      MaxLength       =   40
      TabIndex        =   7
      Text            =   "txttemp(27)"
      ToolTipText     =   "Press F3 for Creation"
      Top             =   2640
      Width           =   1365
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
      Index           =   24
      Left            =   7080
      MaxLength       =   40
      TabIndex        =   20
      Text            =   "txttemp(24)"
      ToolTipText     =   "Press F3 for Creation"
      Top             =   4680
      Width           =   2685
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
      Index           =   25
      Left            =   2280
      MaxLength       =   40
      TabIndex        =   17
      Text            =   "txttemp(25)"
      ToolTipText     =   "Press F3 for Creation"
      Top             =   4680
      Width           =   2685
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
      Index           =   23
      Left            =   2280
      MaxLength       =   40
      TabIndex        =   24
      Text            =   "txttemp(23)"
      ToolTipText     =   "Press F3 for Creation"
      Top             =   6780
      Width           =   2685
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
      Index           =   23
      Left            =   9690
      TabIndex        =   82
      Text            =   "23"
      Top             =   2295
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
      Index           =   7
      Left            =   2280
      MaxLength       =   40
      TabIndex        =   21
      Text            =   "txttemp(7)"
      ToolTipText     =   "Press F3 for Creation"
      Top             =   5610
      Width           =   2685
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
      Index           =   8
      Left            =   7095
      MaxLength       =   40
      TabIndex        =   26
      Text            =   "txttemp(8)"
      ToolTipText     =   "Press F3 for Creation"
      Top             =   5595
      Width           =   2685
   End
   Begin VB.CheckBox Chk 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00EDFDFE&
      Caption         =   "Is &Mill Bill Amount Include DbNT ?"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   420
      Index           =   22
      Left            =   7020
      TabIndex        =   11
      Top             =   870
      Width           =   2850
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
      Index           =   22
      Left            =   9315
      TabIndex        =   81
      Text            =   "22"
      Top             =   3105
      Visible         =   0   'False
      Width           =   420
   End
   Begin VB.CheckBox Chk 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00EDFDFE&
      Caption         =   "Is Separate VAT A/C (Trade) ?"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   300
      Index           =   19
      Left            =   4095
      TabIndex        =   12
      Top             =   2055
      Width           =   3195
   End
   Begin VB.CheckBox Chk 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00EDFDFE&
      Caption         =   "Is Separate VAT A/C (Depot) ?"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   300
      Index           =   21
      Left            =   4095
      TabIndex        =   14
      Top             =   2700
      Width           =   3195
   End
   Begin VB.CheckBox Chk 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00EDFDFE&
      Caption         =   "Is Separate VAT A/C (Consign) ?"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   300
      Index           =   20
      Left            =   4095
      TabIndex        =   13
      Top             =   2385
      Width           =   3195
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
      Left            =   9510
      TabIndex        =   80
      Text            =   "21"
      Top             =   3135
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
      Left            =   9630
      TabIndex        =   79
      Text            =   "20"
      Top             =   2970
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
      Left            =   9240
      TabIndex        =   78
      Text            =   "19"
      Top             =   3045
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
      Left            =   9960
      TabIndex        =   77
      Text            =   "18"
      Top             =   4080
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
      Index           =   18
      Left            =   2280
      MaxLength       =   40
      TabIndex        =   22
      Text            =   "txttemp(18)"
      ToolTipText     =   "Press F3 for Creation"
      Top             =   6015
      Width           =   2685
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
      Left            =   9840
      TabIndex        =   76
      Text            =   "17"
      Top             =   5160
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
      Index           =   17
      Left            =   9690
      MaxLength       =   40
      TabIndex        =   31
      Text            =   "txttemp(17)"
      ToolTipText     =   "Press F3 for Creation"
      Top             =   6780
      Visible         =   0   'False
      Width           =   105
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
      Left            =   9450
      TabIndex        =   75
      Text            =   "16"
      Top             =   3000
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
      Left            =   9570
      TabIndex        =   74
      Text            =   "15"
      Top             =   3120
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
      Left            =   9510
      TabIndex        =   73
      Text            =   "14"
      Top             =   3045
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
      Left            =   9750
      TabIndex        =   72
      Text            =   "13"
      Top             =   2910
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
      Index           =   16
      Left            =   2280
      MaxLength       =   40
      TabIndex        =   16
      Text            =   "txttemp(16)"
      ToolTipText     =   "Press F3 for Creation"
      Top             =   4260
      Width           =   2685
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
      Index           =   15
      Left            =   2280
      MaxLength       =   40
      TabIndex        =   15
      Text            =   "txttemp(15)"
      ToolTipText     =   "Press F3 for Creation"
      Top             =   3855
      Width           =   2685
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
      Index           =   14
      Left            =   7095
      MaxLength       =   40
      TabIndex        =   19
      Text            =   "txttemp(14)"
      ToolTipText     =   "Press F3 for Creation"
      Top             =   4260
      Width           =   2685
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
      Index           =   13
      Left            =   7095
      MaxLength       =   40
      TabIndex        =   18
      Text            =   "txttemp(13)"
      ToolTipText     =   "Press F3 for Creation"
      Top             =   3840
      Width           =   2685
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
      Left            =   9525
      TabIndex        =   71
      Text            =   "12"
      Top             =   3150
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
      Left            =   9585
      TabIndex        =   70
      Text            =   "11"
      Top             =   3210
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
      Left            =   9345
      TabIndex        =   69
      Text            =   "10"
      Top             =   3180
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
      Left            =   9510
      TabIndex        =   68
      Text            =   "9"
      Top             =   3045
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
      Left            =   9375
      TabIndex        =   67
      Text            =   "8"
      Top             =   3075
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
      Left            =   9615
      TabIndex        =   66
      Text            =   "7"
      Top             =   2850
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
      Left            =   9465
      TabIndex        =   65
      Text            =   "6"
      Top             =   3075
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
      Left            =   9780
      TabIndex        =   64
      Text            =   "5"
      Top             =   2775
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
      Left            =   9615
      TabIndex        =   63
      Text            =   "4"
      Top             =   3075
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
      Left            =   9750
      TabIndex        =   62
      Text            =   "3"
      Top             =   3105
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
      Left            =   9795
      TabIndex        =   61
      Text            =   "2"
      Top             =   2790
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
      Left            =   9660
      TabIndex        =   60
      Text            =   "1"
      Top             =   3075
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
      Index           =   12
      Left            =   2280
      MaxLength       =   40
      TabIndex        =   25
      Text            =   "txttemp(12)"
      ToolTipText     =   "Press F3 for Creation"
      Top             =   7170
      Width           =   2565
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
      Index           =   11
      Left            =   7095
      MaxLength       =   40
      TabIndex        =   28
      Text            =   "txttemp(11)"
      ToolTipText     =   "Press F3 for Creation"
      Top             =   6405
      Width           =   2685
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
      Index           =   10
      Left            =   2280
      MaxLength       =   40
      TabIndex        =   23
      Text            =   "txttemp(10)"
      ToolTipText     =   "Press F3 for Creation"
      Top             =   6420
      Width           =   2685
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
      Index           =   9
      Left            =   7095
      MaxLength       =   40
      TabIndex        =   27
      Text            =   "txttemp(9)"
      ToolTipText     =   "Press F3 for Creation"
      Top             =   6000
      Width           =   2685
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
      Index           =   6
      Left            =   5415
      MaxLength       =   40
      TabIndex        =   9
      Text            =   "txttemp(6)"
      ToolTipText     =   "Press F3 for Creation"
      Top             =   900
      Width           =   1365
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
      Index           =   5
      Left            =   5415
      MaxLength       =   40
      TabIndex        =   8
      Text            =   "txttemp(5)"
      ToolTipText     =   "Press F3 for Creation"
      Top             =   525
      Width           =   1365
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
      Index           =   3
      Left            =   2325
      MaxLength       =   40
      TabIndex        =   5
      Text            =   "txttemp(3)"
      ToolTipText     =   "Press F3 for Creation"
      Top             =   1905
      Width           =   1365
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
      Index           =   2
      Left            =   2325
      MaxLength       =   40
      TabIndex        =   4
      Text            =   "txttemp(2)"
      ToolTipText     =   "Press F3 for Creation"
      Top             =   1530
      Width           =   1365
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
      Index           =   1
      Left            =   2325
      MaxLength       =   40
      TabIndex        =   3
      Text            =   "txttemp(1)"
      ToolTipText     =   "Press F3 for Creation"
      Top             =   1140
      Width           =   1365
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
      Index           =   0
      Left            =   2325
      MaxLength       =   40
      TabIndex        =   2
      Text            =   "txttemp(0)"
      ToolTipText     =   "Press F3 for Creation"
      Top             =   765
      Width           =   1365
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
      Left            =   9825
      TabIndex        =   55
      Text            =   "0"
      Top             =   3300
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
      Index           =   4
      Left            =   2325
      MaxLength       =   40
      TabIndex        =   6
      Text            =   "txttemp(4)"
      ToolTipText     =   "Press F3 for Creation"
      Top             =   2265
      Width           =   1365
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
      Left            =   8160
      Style           =   1  'Graphical
      TabIndex        =   52
      ToolTipText     =   "Save Record"
      Top             =   1635
      Width           =   1095
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
      Left            =   8160
      Style           =   1  'Graphical
      TabIndex        =   53
      ToolTipText     =   "Exit"
      Top             =   2085
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
      Left            =   9915
      Style           =   1  'Graphical
      TabIndex        =   56
      ToolTipText     =   "To Add New Record"
      Top             =   4290
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
      Left            =   9975
      Style           =   1  'Graphical
      TabIndex        =   57
      ToolTipText     =   "To Add New Record"
      Top             =   4725
      Visible         =   0   'False
      Width           =   1095
   End
   Begin Crystal.CrystalReport cryReport 
      Bindings        =   "frmSetting.frx":058A
      Left            =   -15
      Top             =   855
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
      BackStyle       =   0  'Transparent
      Caption         =   "TCS Payable A/c"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   360
      Index           =   28
      Left            =   5040
      TabIndex        =   94
      Top             =   6780
      Width           =   2205
   End
   Begin VB.Label lbl1 
      BackStyle       =   0  'Transparent
      Caption         =   "TCS Receivable A/c"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   360
      Index           =   19
      Left            =   5055
      TabIndex        =   93
      Top             =   7170
      Width           =   2205
   End
   Begin VB.Label lbl1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Hank Purchase  :"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   225
      Index           =   27
      Left            =   3840
      TabIndex        =   86
      Top             =   1380
      Width           =   1410
   End
   Begin VB.Label lbl1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Hank Trade Sale    :"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   225
      Index           =   26
      Left            =   480
      TabIndex        =   85
      Top             =   2700
      Width           =   1605
   End
   Begin VB.Label lbl1 
      BackStyle       =   0  'Transparent
      Caption         =   "Hunk  Purchase  A/C"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   360
      Index           =   24
      Left            =   5040
      TabIndex        =   84
      Top             =   4680
      Width           =   1845
   End
   Begin VB.Label lbl1 
      BackStyle       =   0  'Transparent
      Caption         =   "Hunk  Sales  A/C"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   360
      Index           =   25
      Left            =   360
      TabIndex        =   83
      Top             =   4680
      Width           =   1845
   End
   Begin VB.Label lbl1 
      BackStyle       =   0  'Transparent
      Caption         =   "TDS Paid A/C"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   360
      Index           =   23
      Left            =   375
      TabIndex        =   47
      Top             =   6780
      Width           =   2205
   End
   Begin VB.Label lbl1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "&VAT Accounting      :"
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
      Left            =   3990
      TabIndex        =   38
      Top             =   1770
      Width           =   1665
   End
   Begin VB.Line Line1 
      BorderColor     =   &H000000C0&
      X1              =   240
      X2              =   9960
      Y1              =   5055
      Y2              =   5055
   End
   Begin VB.Label lbl1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "&Other Expenses Account  :"
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
      TabIndex        =   43
      Top             =   5280
      Width           =   2250
   End
   Begin VB.Label lbl1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Mill Bill Debit Note A/C"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   225
      Index           =   20
      Left            =   375
      TabIndex        =   45
      Top             =   6015
      Width           =   1815
   End
   Begin VB.Label lbl1 
      BackStyle       =   0  'Transparent
      Caption         =   "SIT Sales  A/C"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   360
      Index           =   18
      Left            =   375
      TabIndex        =   42
      Top             =   4260
      Width           =   2205
   End
   Begin VB.Label lbl1 
      BackStyle       =   0  'Transparent
      Caption         =   "Trade Sales  A/C"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   360
      Index           =   17
      Left            =   375
      TabIndex        =   54
      Top             =   3840
      Width           =   2205
   End
   Begin VB.Label lbl1 
      BackStyle       =   0  'Transparent
      Caption         =   "SIT Purchase  A/C"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   360
      Index           =   16
      Left            =   5025
      TabIndex        =   41
      Top             =   4245
      Width           =   2205
   End
   Begin VB.Label lbl1 
      BackStyle       =   0  'Transparent
      Caption         =   "Trade Purchase  A/C"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   360
      Index           =   15
      Left            =   5025
      TabIndex        =   40
      Top             =   3840
      Width           =   2205
   End
   Begin VB.Label lbl1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Sale && Purchase &Account  :"
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
      Index           =   14
      Left            =   360
      TabIndex        =   39
      Top             =   3420
      Width           =   2310
   End
   Begin VB.Label lbl1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   " Discount A/C"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   225
      Index           =   13
      Left            =   5025
      TabIndex        =   49
      Top             =   5595
      Width           =   1110
   End
   Begin VB.Label lbl1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "&Late Pay Interest A/C"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   225
      Index           =   12
      Left            =   375
      TabIndex        =   44
      Top             =   5610
      Width           =   1695
   End
   Begin VB.Label lbl1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "SIT Purchase     :"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   225
      Index           =   11
      Left            =   3885
      TabIndex        =   37
      Top             =   900
      Width           =   1380
   End
   Begin VB.Label lbl1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Trade Purchase :"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   225
      Index           =   10
      Left            =   3885
      TabIndex        =   36
      Top             =   525
      Width           =   1410
   End
   Begin VB.Label lbl1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "SIT Sale                :"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   225
      Index           =   9
      Left            =   495
      TabIndex        =   35
      Top             =   2265
      Width           =   1455
   End
   Begin VB.Label lbl1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Mill Bill Sale         :"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   225
      Index           =   8
      Left            =   495
      TabIndex        =   34
      Top             =   1905
      Width           =   1440
   End
   Begin VB.Label lbl1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Depot Sale           :"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   225
      Index           =   7
      Left            =   495
      TabIndex        =   33
      Top             =   1530
      Width           =   1455
   End
   Begin VB.Label lbl1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Consignment Sale :"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   225
      Index           =   6
      Left            =   495
      TabIndex        =   32
      Top             =   1140
      Width           =   1635
   End
   Begin VB.Label lbl1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Trade Sale            :"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   225
      Index           =   2
      Left            =   495
      TabIndex        =   1
      Top             =   765
      Width           =   1485
   End
   Begin VB.Shape Shape4 
      BorderColor     =   &H00C00000&
      BorderWidth     =   5
      FillColor       =   &H00C00000&
      Height          =   7440
      Left            =   0
      Top             =   300
      Width           =   10290
   End
   Begin VB.Shape Shape3 
      BorderColor     =   &H000000C0&
      FillColor       =   &H000000C0&
      FillStyle       =   0  'Solid
      Height          =   30
      Left            =   0
      Top             =   8655
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
      Left            =   6990
      TabIndex        =   59
      Top             =   0
      Width           =   2475
   End
   Begin VB.Label lbl1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "&Tax Type  :"
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
      Index           =   0
      Left            =   360
      TabIndex        =   0
      Top             =   465
      Width           =   900
   End
   Begin VB.Label lbl1 
      BackStyle       =   0  'Transparent
      Caption         =   "B&rokerage A/C"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   360
      Index           =   5
      Left            =   375
      TabIndex        =   48
      Top             =   7170
      Width           =   2205
   End
   Begin VB.Label lbl1 
      BackStyle       =   0  'Transparent
      Caption         =   "TDS-194Q Receivable "
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   360
      Index           =   4
      Left            =   5025
      TabIndex        =   51
      Top             =   6405
      Width           =   2085
   End
   Begin VB.Label lbl1 
      BackStyle       =   0  'Transparent
      Caption         =   "TDS Payable A/C"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   360
      Index           =   3
      Left            =   375
      TabIndex        =   46
      Top             =   6420
      Width           =   2205
   End
   Begin VB.Label lbl1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "TDS-194Q Payable (Pur)"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   225
      Index           =   1
      Left            =   5025
      TabIndex        =   50
      Top             =   6000
      Width           =   2025
   End
   Begin VB.Shape Shape1 
      FillColor       =   &H00BDDFD1&
      FillStyle       =   0  'Solid
      Height          =   1455
      Left            =   7905
      Shape           =   4  'Rounded Rectangle
      Top             =   1335
      Width           =   1575
   End
   Begin VB.Label Label6 
      BackColor       =   &H00C00000&
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
      Left            =   -75
      TabIndex        =   58
      Top             =   -45
      Width           =   10455
   End
   Begin VB.Shape Shape2 
      BackColor       =   &H00808080&
      BorderColor     =   &H004D177D&
      FillColor       =   &H00ECFFFF&
      FillStyle       =   0  'Solid
      Height          =   4260
      Left            =   240
      Top             =   3315
      Width           =   9735
   End
   Begin VB.Shape Shape5 
      BorderColor     =   &H00000080&
      FillColor       =   &H00EDFDFE&
      FillStyle       =   0  'Solid
      Height          =   2745
      Left            =   240
      Top             =   390
      Width           =   9750
   End
End
Attribute VB_Name = "frmSetting"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Const tblName As String = "tblMastSetting"
Const MaxNo As Long = 29
Dim FormAction As Integer
Dim rstbl As Recordset
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case 13:
            Sendkeys "{TAB}"
    End Select
End Sub
Private Sub Form_Load()
On Error GoTo ErrorHandler
    Dim X As String
    Me.Top = 300
    Me.Left = Screen.Width / 2 - Me.Width / 2
    X = "Select * from " & tblName
    GProcRstOpen rstbl, X, "O"
    FormAction = vbDataActionAddNew
    Call GProcClearForm(Me, rstbl, MaxNo, True)
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
On Error GoTo ErrorHandler
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

Private Sub txttemp_GotFocus(Index As Integer)
GProcSelectBox Me.ActiveControl
End Sub

Private Sub txttemp_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
Dim GName As String
If KeyCode = vbKeyF3 Then
Select Case Index
    Case 0, 1, 2, 3, 4, 5, 6, 26, 27 '-- Tax
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
TopPos = txttemp(Index).Top + Me.Top + 650
LeftPos = txttemp(Index).Left + Me.Left
If KeyAscii = 13 Then Exit Sub '-- Enter
Select Case Index
     Case 0, 1, 2, 3, 4, 5, 6, 26, 27 '-- Tax
        gClsSearch.SearchMultiField "TblMastNarration", "Narration", Array("Tax Type"), Array(txttemp(Index).Width), "NarrType='T'", Chr(KeyAscii), "Narration", LeftPos, TopPos, True
        KeyAscii = 0
        If gClsSearch.SearchMultiRetCol(0) <> "" Then
            txttemp(Index) = gClsSearch.SearchMultiRetCol(0)
        End If
     Case Else  '-- A/C
        If Index = 8 Or Index = 18 Then '-- Db Nt A/c
            gClsSearch.SearchMultiField "QryMastaccount", "Account,AcName,AcCity,AgName", Array("Account", "AcName", "City", "Group"), Array(3000, 0, 2000, 3000), "AcIsDbNtAc=1", Chr(KeyAscii), "AcName", 0, 0, True
        Else
            gClsSearch.SearchMultiField "QryMastaccount", "Account,AcName,AcCity,AgName", Array("Account", "AcName", "City", "Group"), Array(3000, 0, 2000, 3000), "", Chr(KeyAscii), "AcName", 0, 0, True
        End If
        KeyAscii = 0
        If gClsSearch.SearchMultiRetCol(0) <> "" And gClsSearch.SearchMultiRetCol(1) = "" Then
            txttemp(Index) = gClsSearch.SearchMultiRetCol(0)
        Else
           If gClsSearch.SearchMultiRetCol(1) <> "" Then txttemp(Index) = gClsSearch.SearchMultiRetCol(1)
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
For i = 0 To 6
    If Trim(txttemp(i)) <> "" Then
        If GProcGetColumnValue("tblMastNarration", "Narration", txttemp(i), "S", "NarrCode", "N") = 0 Then
           GProcShowForm frmMastNarrationTax, frmMain.mnuMstfrm(6), Array(frmMastNarrationTax.txt(2)), Array(txttemp(i))
           Exit Function
        End If
    Else
        txttemp(i) = "Tax Free" 'GProcGetColumnValue("tblMastNarration", "Narration", txttemp(i), "S", "NarrCode", "N")
    End If
Next i
'--Create A/c
For i = 7 To 18
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

For i = 24 To 25
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


'--Create A/c : TDS Paid
If Trim(txttemp(23)) <> "" Then
    If GProcGetColumnValue("tblMastAccount", "AcName", txttemp(23), "S", "AcCode", "N") = 0 Then
       'GName = GProcGetColumnValue("TblMastGroup", "AgCode", 90029, "N", "AgName", "S")
       GProcShowForm frmMastAccount, frmMain.mnuMstfrm(0), Array(frmMastAccount.txt(1), frmMastAccount.txttemp(2)), Array(txttemp(23), GName)
       Exit Function
    End If
Else
    txttemp(23) = "-"
End If
ValidateData = True
End Function
Private Sub FillTempFromTxt()
Dim i As Long
For i = 0 To 6 '-- Tax Type
    txttemp(i) = GProcGetColumnValue("tblMastNarration", "NarrCode", txt(i), "N", "Narration", "S")
Next i
For i = 26 To 27 '-- Tax Type
    txttemp(i) = GProcGetColumnValue("tblMastNarration", "NarrCode", txt(i), "N", "Narration", "S")
Next i



For i = 7 To 18 '-- Account
    txttemp(i) = GProcGetColumnValue("tblMastAccount", "AcCode", txt(i), "N", "AcName", "S")
Next i
For i = 24 To 25 '-- Account
    txttemp(i) = GProcGetColumnValue("tblMastAccount", "AcCode", txt(i), "N", "AcName", "S")
Next i

For i = 28 To 29 '-- Account
    txttemp(i) = GProcGetColumnValue("tblMastAccount", "AcCode", txt(i), "N", "AcName", "S")
Next i


'- A/c
txttemp(23) = GProcGetColumnValue("tblMastAccount", "AcCode", txt(23), "N", "AcName", "S")
For i = 19 To 21 '-- Is No separate VAT Tax Voucher creation
    Chk(i) = IIf(txt(i) = 0, 1, 0)
Next i
Chk(22) = txt(22) '- Mill Bill With DbNt
End Sub
Private Sub FillTxtFromTemp()
Dim i As Long
For i = 0 To 6 '-- Tax Type
    txt(i) = GProcGetColumnValue("tblMastNarration", "Narration", txttemp(i), "S", "NarrCode", "N")
Next i
For i = 26 To 27 '-- Tax Type
    txt(i) = GProcGetColumnValue("tblMastNarration", "Narration", txttemp(i), "S", "NarrCode", "N")
Next i



For i = 7 To 18 '-- Account
    txt(i) = GProcGetColumnValue("tblMastAccount", "AcName", txttemp(i), "S", "AcCode", "N")
Next i
'- A/c
txt(23) = GProcGetColumnValue("tblMastAccount", "AcName", txttemp(23), "S", "AcCode", "N")
txt(24) = GProcGetColumnValue("tblMastAccount", "AcName", txttemp(24), "S", "AcCode", "N")
txt(25) = GProcGetColumnValue("tblMastAccount", "AcName", txttemp(25), "S", "AcCode", "N")


txt(28) = GProcGetColumnValue("tblMastAccount", "AcName", txttemp(28), "S", "AcCode", "N")
txt(29) = GProcGetColumnValue("tblMastAccount", "AcName", txttemp(29), "S", "AcCode", "N")


For i = 19 To 21 '-- Is No separate VAT Tax Voucher creation
    txt(i) = IIf(Chk(i) = 0, 1, 0)
Next i
txt(22) = Chk(22) '- Mill Bill With DbNt
End Sub
