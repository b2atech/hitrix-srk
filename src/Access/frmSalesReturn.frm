VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form frmSalesReturn 
   BackColor       =   &H00C2E3FC&
   BorderStyle     =   0  'None
   Caption         =   " Sale Return"
   ClientHeight    =   7215
   ClientLeft      =   0
   ClientTop       =   -105
   ClientWidth     =   10590
   DrawMode        =   10  'Mask Pen
   FillColor       =   &H000040C0&
   FillStyle       =   0  'Solid
   Icon            =   "frmSalesReturn.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7215
   ScaleWidth      =   10590
   ShowInTaskbar   =   0   'False
   Begin Crystal.CrystalReport cryReport 
      Left            =   4545
      Top             =   3930
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
   Begin VB.CommandButton cmdLedger 
      Caption         =   "Le&dger"
      Height          =   375
      Index           =   1
      Left            =   3600
      TabIndex        =   90
      Top             =   2205
      Width           =   1605
   End
   Begin VB.CommandButton cmdLedger 
      Caption         =   "Le&dger"
      Height          =   375
      Index           =   0
      Left            =   6780
      TabIndex        =   89
      Top             =   1125
      Width           =   1650
   End
   Begin VB.TextBox txtVSub 
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
      Index           =   12
      Left            =   8610
      TabIndex        =   88
      Text            =   "txtVSub(12)"
      Top             =   7860
      Visible         =   0   'False
      Width           =   360
   End
   Begin VB.TextBox txttmpInvNo 
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
      Left            =   6795
      TabIndex        =   9
      Text            =   "txttmpInvNo"
      Top             =   1845
      Width           =   1635
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
      Left            =   5310
      TabIndex        =   86
      Text            =   "txtCrBalance"
      Top             =   1125
      Width           =   1380
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
      Left            =   1410
      TabIndex        =   85
      Text            =   "txtDrBalance"
      Top             =   2220
      Width           =   1935
   End
   Begin VB.ComboBox Cbo 
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
      Left            =   3495
      TabIndex        =   34
      Text            =   "Cbo(12)"
      Top             =   4830
      Width           =   1200
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
      Height          =   345
      Index           =   4
      Left            =   7935
      TabIndex        =   84
      Text            =   "4"
      Top             =   8400
      Visible         =   0   'False
      Width           =   315
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
      Height          =   345
      Index           =   5
      Left            =   7230
      TabIndex        =   83
      Text            =   "5"
      Top             =   8445
      Visible         =   0   'False
      Width           =   435
   End
   Begin VB.TextBox txt 
      Height          =   360
      Index           =   2
      Left            =   1905
      TabIndex        =   82
      Text            =   "2"
      Top             =   8325
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.TextBox txt 
      Height          =   360
      Index           =   3
      Left            =   1620
      TabIndex        =   81
      Text            =   "3"
      Top             =   8370
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.TextBox txt 
      Height          =   360
      Index           =   6
      Left            =   2205
      TabIndex        =   80
      Text            =   "6"
      Top             =   8325
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.TextBox txt 
      Height          =   285
      Index           =   7
      Left            =   6705
      TabIndex        =   79
      Text            =   "7"
      Top             =   8415
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.TextBox txt 
      Height          =   360
      Index           =   9
      Left            =   3570
      TabIndex        =   78
      Text            =   "9"
      Top             =   8325
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.TextBox txt 
      Appearance      =   0  'Flat
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
      Height          =   345
      Index           =   12
      Left            =   5190
      TabIndex        =   77
      Text            =   "12"
      Top             =   8430
      Visible         =   0   'False
      Width           =   915
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
      Height          =   345
      Index           =   16
      Left            =   2640
      TabIndex        =   76
      Text            =   "16"
      Top             =   8370
      Visible         =   0   'False
      Width           =   435
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
      Height          =   345
      Index           =   20
      Left            =   1140
      TabIndex        =   75
      Text            =   "20"
      Top             =   8370
      Visible         =   0   'False
      Width           =   315
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
      Height          =   345
      Index           =   15
      Left            =   3090
      TabIndex        =   74
      Text            =   "15"
      Top             =   8340
      Visible         =   0   'False
      Width           =   435
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
      Height          =   345
      Index           =   1
      Left            =   4830
      TabIndex        =   73
      Text            =   "1"
      Top             =   8385
      Visible         =   0   'False
      Width           =   315
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
      Height          =   345
      Index           =   21
      Left            =   3975
      TabIndex        =   72
      Text            =   "21"
      Top             =   8340
      Visible         =   0   'False
      Width           =   315
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
      Height          =   345
      Index           =   22
      Left            =   4365
      TabIndex        =   71
      Text            =   "22"
      Top             =   8385
      Visible         =   0   'False
      Width           =   315
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
      Height          =   345
      Index           =   23
      Left            =   8400
      TabIndex        =   70
      Text            =   "23"
      Top             =   8415
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.TextBox txtVou 
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
      Left            =   1395
      TabIndex        =   47
      Text            =   "txtVou(7)"
      Top             =   5355
      Width           =   3930
   End
   Begin VB.TextBox txtVou 
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
      Left            =   1395
      TabIndex        =   48
      Text            =   "txtVou(8)"
      Top             =   5760
      Width           =   3930
   End
   Begin VB.TextBox txtVou 
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
      Left            =   1395
      TabIndex        =   49
      Text            =   "txtVou(9)"
      Top             =   6150
      Width           =   3930
   End
   Begin VB.TextBox txtVou 
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
      Left            =   1395
      TabIndex        =   50
      Text            =   "txtVou(10)"
      Top             =   6570
      Width           =   3930
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
      Left            =   6795
      TabIndex        =   28
      Text            =   "txt(10)"
      Top             =   3915
      Width           =   1680
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
      Left            =   6795
      TabIndex        =   45
      Text            =   "txt(11)"
      Top             =   6045
      Width           =   1680
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
      Left            =   4710
      TabIndex        =   35
      Text            =   "txt(13)"
      Top             =   4830
      Width           =   615
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
      Left            =   6795
      TabIndex        =   37
      Text            =   "txt(14)"
      Top             =   5220
      Width           =   1680
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
      Left            =   6795
      TabIndex        =   43
      Text            =   "txt(17)"
      Top             =   5640
      Width           =   1680
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
      Left            =   6795
      TabIndex        =   41
      Text            =   "txt(19)"
      Top             =   4830
      Width           =   1680
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
      Left            =   6795
      ScrollBars      =   1  'Horizontal
      TabIndex        =   39
      Text            =   "txt(18)"
      Top             =   4440
      Width           =   1680
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
      Index           =   4
      Left            =   1410
      MaxLength       =   40
      TabIndex        =   13
      Text            =   "txttemp(4)"
      Top             =   1830
      Width           =   3810
   End
   Begin VB.TextBox txtVSub 
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
      Index           =   6
      Left            =   1980
      TabIndex        =   22
      Text            =   "txtVSub(6)"
      Top             =   3885
      Width           =   1395
   End
   Begin VB.TextBox txtVSub 
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
      Index           =   7
      Left            =   3585
      TabIndex        =   24
      Text            =   "txtVSub(7)"
      Top             =   3885
      Width           =   1395
   End
   Begin VB.TextBox txtVSub 
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
      Index           =   8
      Left            =   5265
      TabIndex        =   26
      Text            =   "txtVSub(8)"
      Top             =   3885
      Width           =   1395
   End
   Begin VB.TextBox txtVSub 
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
      Left            =   1395
      TabIndex        =   32
      Text            =   "txtVSub(10)"
      Top             =   4830
      Width           =   1980
   End
   Begin VB.TextBox txttmpVSub 
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
      Left            =   1410
      MaxLength       =   40
      TabIndex        =   17
      Text            =   "txttmpVSub(4)"
      Top             =   3075
      Width           =   3810
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
      Index           =   8
      Left            =   8790
      TabIndex        =   18
      Text            =   "txt(8)"
      Top             =   8295
      Visible         =   0   'False
      Width           =   435
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
      Index           =   5
      Left            =   1410
      MaxLength       =   40
      TabIndex        =   15
      Text            =   "txttemp(5)"
      Top             =   2640
      Width           =   3810
   End
   Begin VB.TextBox txtVSub 
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
      Index           =   5
      Left            =   360
      TabIndex        =   20
      Text            =   "txtVSub(5)"
      Top             =   3885
      Width           =   1395
   End
   Begin VB.TextBox txttmpInvYear 
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
      Left            =   6795
      TabIndex        =   11
      Text            =   "txttmpInvYear"
      Top             =   2760
      Width           =   1635
   End
   Begin VB.TextBox txttmpInvType 
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
      Left            =   6795
      TabIndex        =   7
      Text            =   "txttmpInvType"
      Top             =   2295
      Width           =   1635
   End
   Begin VB.TextBox txttmpVSub 
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
      Left            =   1395
      MaxLength       =   40
      TabIndex        =   30
      Text            =   "txttmpVSub(17)"
      Top             =   4440
      Width           =   1980
   End
   Begin VB.TextBox txt 
      Alignment       =   1  'Right Justify
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
      Left            =   1410
      MaxLength       =   40
      TabIndex        =   1
      Text            =   "txt(0)"
      Top             =   675
      Width           =   1560
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
      Index           =   3
      Left            =   1410
      MaxLength       =   40
      TabIndex        =   5
      Text            =   "txtTemp(3)"
      Top             =   1125
      Width           =   3810
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
      Left            =   9210
      Style           =   1  'Graphical
      TabIndex        =   65
      Top             =   7755
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
      Left            =   9210
      Style           =   1  'Graphical
      TabIndex        =   64
      Top             =   7275
      Visible         =   0   'False
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
      Left            =   9060
      Style           =   1  'Graphical
      TabIndex        =   63
      ToolTipText     =   "Exit"
      Top             =   5550
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
      Left            =   9060
      Style           =   1  'Graphical
      TabIndex        =   55
      Top             =   4905
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
      Left            =   9060
      Style           =   1  'Graphical
      TabIndex        =   54
      ToolTipText     =   "To Save Record"
      Top             =   4425
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
      Left            =   8580
      Style           =   1  'Graphical
      TabIndex        =   62
      Top             =   7620
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
      Left            =   9060
      Style           =   1  'Graphical
      TabIndex        =   53
      ToolTipText     =   "To Delete Record"
      Top             =   3765
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
      Height          =   400
      Index           =   7
      Left            =   9060
      Style           =   1  'Graphical
      TabIndex        =   52
      ToolTipText     =   "To Modify Record"
      Top             =   3285
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
      Left            =   9060
      Style           =   1  'Graphical
      TabIndex        =   51
      ToolTipText     =   "To Add New Record"
      Top             =   2805
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
      Index           =   5
      Left            =   8625
      Style           =   1  'Graphical
      TabIndex        =   56
      Top             =   7785
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
      Left            =   9060
      Style           =   1  'Graphical
      TabIndex        =   57
      ToolTipText     =   "To Find Record"
      Top             =   2145
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
      Left            =   9660
      Style           =   1  'Graphical
      TabIndex        =   61
      ToolTipText     =   "Last"
      Top             =   1665
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
      Left            =   9060
      Style           =   1  'Graphical
      TabIndex        =   60
      ToolTipText     =   "First"
      Top             =   1665
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
      Left            =   9660
      Style           =   1  'Graphical
      TabIndex        =   59
      ToolTipText     =   "Next"
      Top             =   1185
      Width           =   495
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
      Left            =   9060
      Style           =   1  'Graphical
      TabIndex        =   58
      ToolTipText     =   "Previous"
      Top             =   1185
      Width           =   495
   End
   Begin MSMask.MaskEdBox mskDt 
      Height          =   360
      Index           =   2
      Left            =   6795
      TabIndex        =   3
      Top             =   675
      Width           =   1635
      _ExtentX        =   2884
      _ExtentY        =   635
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
   Begin VB.Label Label15 
      BackStyle       =   0  'Transparent
      Caption         =   "Tax Amt  :"
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
      Height          =   360
      Index           =   19
      Left            =   5670
      TabIndex        =   87
      Top             =   5190
      Width           =   1035
   End
   Begin VB.Line Line4 
      BorderColor     =   &H00000080&
      X1              =   5565
      X2              =   5565
      Y1              =   7020
      Y2              =   4350
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00000080&
      X1              =   5565
      X2              =   5565
      Y1              =   1695
      Y2              =   3540
   End
   Begin VB.Line Line1 
      BorderColor     =   &H000000C0&
      X1              =   90
      X2              =   8580
      Y1              =   4350
      Y2              =   4350
   End
   Begin VB.Line Line3 
      BorderColor     =   &H000000C0&
      X1              =   135
      X2              =   8565
      Y1              =   3525
      Y2              =   3525
   End
   Begin VB.Label Label15 
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
      ForeColor       =   &H00645285&
      Height          =   360
      Index           =   24
      Left            =   5340
      TabIndex        =   36
      Top             =   4845
      Width           =   255
   End
   Begin VB.Label Label15 
      BackStyle       =   0  'Transparent
      Caption         =   "&Narration  :"
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
      Height          =   360
      Index           =   20
      Left            =   270
      TabIndex        =   46
      Top             =   5355
      Width           =   1035
   End
   Begin VB.Label Label15 
      BackStyle       =   0  'Transparent
      Caption         =   "Rt Per:"
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
      Height          =   330
      Index           =   6
      Left            =   5250
      TabIndex        =   25
      Top             =   3570
      Width           =   825
   End
   Begin VB.Label Label15 
      BackStyle       =   0  'Transparent
      Caption         =   "Type :"
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
      Height          =   360
      Index           =   9
      Left            =   5700
      TabIndex        =   6
      Top             =   2295
      Width           =   1035
   End
   Begin VB.Label Label15 
      BackStyle       =   0  'Transparent
      Caption         =   "Year :"
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
      Height          =   360
      Index           =   10
      Left            =   5700
      TabIndex        =   10
      Top             =   2760
      Width           =   1035
   End
   Begin VB.Label Label15 
      BackStyle       =   0  'Transparent
      Caption         =   "Amount    :"
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
      Index           =   23
      Left            =   5700
      TabIndex        =   44
      Top             =   6045
      Width           =   1035
   End
   Begin VB.Label Label15 
      BackStyle       =   0  'Transparent
      Caption         =   "Roff          :"
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
      Height          =   360
      Index           =   22
      Left            =   5700
      TabIndex        =   42
      Top             =   5640
      Width           =   1035
   End
   Begin VB.Label Label15 
      BackStyle       =   0  'Transparent
      Caption         =   "&Tax Type :"
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
      Height          =   360
      Index           =   18
      Left            =   3570
      TabIndex        =   33
      Top             =   4455
      Width           =   1035
   End
   Begin VB.Label Label15 
      BackStyle       =   0  'Transparent
      Caption         =   "Less  :"
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
      Height          =   360
      Index           =   17
      Left            =   5670
      TabIndex        =   40
      Top             =   4830
      Width           =   1035
   End
   Begin VB.Label Label15 
      BackStyle       =   0  'Transparent
      Caption         =   "Add     :"
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
      Height          =   360
      Index           =   16
      Left            =   5700
      TabIndex        =   38
      Top             =   4440
      Width           =   1035
   End
   Begin VB.Label Label15 
      BackStyle       =   0  'Transparent
      Caption         =   "Amount    :"
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
      Index           =   15
      Left            =   6795
      TabIndex        =   27
      Top             =   3570
      Width           =   1035
   End
   Begin VB.Label Label15 
      BackStyle       =   0  'Transparent
      Caption         =   "Wt  :"
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
      Height          =   255
      Index           =   14
      Left            =   1980
      TabIndex        =   21
      Top             =   3570
      Width           =   1125
   End
   Begin VB.Label Label15 
      BackStyle       =   0  'Transparent
      Caption         =   "Bag          :"
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
      Height          =   255
      Index           =   13
      Left            =   375
      TabIndex        =   19
      Top             =   3570
      Width           =   1035
   End
   Begin VB.Label Label15 
      BackStyle       =   0  'Transparent
      Caption         =   "&Lot No      :"
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
      Height          =   360
      Index           =   12
      Left            =   270
      TabIndex        =   31
      Top             =   4830
      Width           =   1035
   End
   Begin VB.Label Label15 
      BackStyle       =   0  'Transparent
      Caption         =   "M&ill        :"
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
      Height          =   255
      Index           =   11
      Left            =   270
      TabIndex        =   14
      Top             =   2640
      Width           =   1035
   End
   Begin VB.Label Label15 
      BackStyle       =   0  'Transparent
      Caption         =   "In&voice  No:"
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
      Height          =   360
      Index           =   8
      Left            =   5700
      TabIndex        =   8
      Top             =   1830
      Width           =   1035
   End
   Begin VB.Label Label15 
      BackStyle       =   0  'Transparent
      Caption         =   "Sa&les (Dr) A/c :"
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
      Height          =   660
      Index           =   7
      Left            =   270
      TabIndex        =   12
      Top             =   1830
      Width           =   1185
   End
   Begin VB.Label Label15 
      BackStyle       =   0  'Transparent
      Caption         =   "Rate      :"
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
      Height          =   330
      Index           =   5
      Left            =   3585
      TabIndex        =   23
      Top             =   3570
      Width           =   1005
   End
   Begin VB.Label Label15 
      BackStyle       =   0  'Transparent
      Caption         =   "&Godown :"
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
      Height          =   360
      Index           =   4
      Left            =   270
      TabIndex        =   29
      Top             =   4440
      Width           =   1035
   End
   Begin VB.Label Label15 
      BackStyle       =   0  'Transparent
      Caption         =   "Item       :"
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
      Height          =   330
      Index           =   3
      Left            =   270
      TabIndex        =   16
      Top             =   3075
      Width           =   1035
   End
   Begin VB.Label Label15 
      BackStyle       =   0  'Transparent
      Caption         =   "&Party      :"
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
      Index           =   2
      Left            =   270
      TabIndex        =   4
      Top             =   1140
      Width           =   1035
   End
   Begin VB.Label Label15 
      BackStyle       =   0  'Transparent
      Caption         =   "Vou Date :"
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
      Height          =   330
      Index           =   1
      Left            =   5700
      TabIndex        =   2
      Top             =   675
      Width           =   1035
   End
   Begin VB.Label Label15 
      BackStyle       =   0  'Transparent
      Caption         =   "Vou &No   :"
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
      Height          =   330
      Index           =   0
      Left            =   270
      TabIndex        =   0
      Top             =   675
      Width           =   1035
   End
   Begin VB.Shape Shape2 
      BackColor       =   &H00808080&
      BorderColor     =   &H004D177D&
      FillColor       =   &H00ECFFFF&
      FillStyle       =   0  'Solid
      Height          =   5340
      Index           =   1
      Left            =   120
      Top             =   1680
      Width           =   8475
   End
   Begin VB.Shape Shape2 
      BackColor       =   &H00808080&
      BorderColor     =   &H004D177D&
      FillColor       =   &H00ECFFFF&
      FillStyle       =   0  'Solid
      Height          =   1110
      Index           =   0
      Left            =   120
      Top             =   495
      Width           =   8475
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Vou Date :"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00400040&
      Height          =   240
      Index           =   1
      Left            =   6000
      TabIndex        =   69
      Top             =   1080
      Width           =   990
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Vou No   :"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00400040&
      Height          =   240
      Index           =   0
      Left            =   3060
      TabIndex        =   68
      Top             =   1200
      Width           =   930
   End
   Begin VB.Shape Shape7 
      BorderColor     =   &H00C00000&
      BorderWidth     =   5
      FillColor       =   &H00C00000&
      Height          =   8565
      Left            =   0
      Top             =   -1380
      Width           =   10590
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
      TabIndex        =   67
      Top             =   15
      Width           =   2475
   End
   Begin VB.Label LblFrmCaption 
      BackColor       =   &H00C00000&
      Caption         =   "   Trade Sales Return"
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
      Left            =   -60
      TabIndex        =   66
      Top             =   0
      Width           =   10620
   End
   Begin VB.Shape Shape1 
      FillColor       =   &H00BDDFD1&
      FillStyle       =   0  'Solid
      Height          =   5265
      Left            =   8820
      Shape           =   4  'Rounded Rectangle
      Top             =   930
      Width           =   1575
   End
End
Attribute VB_Name = "frmSalesReturn"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Const tblName As String = "tblAddLess"
Const IdField As String = "Vno"
Const OrderField As String = "Vdt,Vno"
Const MaxNo As Long = 23
Dim WhereCriteia As String
Dim Rstbl As Recordset '-- tblAddLess
Dim rstblSub As Recordset '-- tblVouSub
Dim rsVou As Recordset '--tblVoucher
Dim varBookMark As Variant
Dim FirstTimeFlag As Boolean
Dim oldFrmTypeStr As String
Dim FormAction As Integer
Dim frmTypeStr As String
Dim OldInvRefNo As Long
Private Sub Cbo_KeyPress(Index As Integer, KeyAscii As Integer)
If Index = 12 Then KeyAscii = 0 '--Tax
End Sub
Private Sub cbo_LostFocus(Index As Integer)
If Index = 12 And FormAction = vbDataActionAddNew Then '--Tax %
    txt(13) = GProcGetColumnValue("tblMastNarration", "Narration", Cbo(12), "S", "MastTaxRate", "N")
    txt(13) = Format(txt(13), FStr)
End If
End Sub
Private Sub cmdBtn_Click(Index As Integer)
Dim FindQStr As String
Dim x As String
Dim Hank As String
On Error GoTo ErrorRoutine
Select Case Index
Case 0 'Previous
    Rstbl.MovePrevious
    If Rstbl.BOF = True Then
        Rstbl.MoveFirst
        MsgBox ("This is First Record"), vbInformation + vbOKOnly, "First Record"
    End If
    ReadFields
Case 1 'Next
    Rstbl.MoveNext
    If Rstbl.EOF = True Then
        Rstbl.MoveLast
        MsgBox ("This is Last Record"), vbInformation + vbOKOnly, "Last Record"
    End If
    ReadFields
Case 2 'First
    Rstbl.MoveFirst
    ReadFields
Case 3 'Last
    Rstbl.MoveLast
    ReadFields
Case 4 'Find
    If gBackEndDB = gBackEndAccess Then
        x = " iif(tblAddSA.Vtype='SY','Trade',tblAddSA.Vtype),tblAddSA.AdInvNo,tblAddSA.Vdt," _
    & " TblAddLess.AdBillAmt from tblAddLess,tblAddLess as tblAddSA,TblVouSub,tblVouSub as TblSubSA "
    Else '-- Oracle
        x = "decode(tblAddSA.Vtype,'SY','Trade',tblAddSA.Vtype),tblAddSA.AdInvNo,tblAddSA.Vdt," _
    & " TblAddLess.AdBillAmt from tblAddLess,tblAddLess tblAddSA,TblVouSub,tblVouSub TblSubSA "
    End If
    '-- Vno,Vdt,Party,Inv Type,Inv No,Inv Dt,Amt
    FindQStr = "Select tblAddLess.Vno,tblAddLess.Vdt,AcName, " & x & " ,tblMastAccount where " _
    & "tblAddLess.AdAcCrcode=tblMastAccount.AcCode and tblAddLess.Vtype=TblVouSub.Vtype and " _
    & " tblAddLess.VNo=TblVouSub.VNo and tblAddLess.VYear=TblVouSub.VYear and TblVouSub.VSubPItCtrlNo=" _
    & " TblSubSA.VSubItCtrlNo and TblSubSA.Vno=tblAddSA.vno and tblSubSA.Vtype=tblAddSA.Vtype" _
    & " and tblSubSA.VYear=tblAddSA.VYear and tblAddLess.Vtype='" & frmTypeStr & "' and tblAddLess.Vyear=" & gCYear
    gClsSearch.SearchMultiField FindQStr, "", Array("VNo", "Vdt", "Party", "Inv Type", "Inv No", "Inv Dt", "Amt"), Array(800, 1100, 2500, 900, 900, 1100, 1000), "", "", "tblAddLess.Vno,tblAddLess.Vdt", 0, 0, False, , True
    If gClsSearch.SearchMultiRetCol(0) <> "" Then
        Rstbl.MoveFirst
        Rstbl.Find "Vno = " & gClsSearch.SearchMultiRetCol(0)
        gClsSearch.SearchMultiRetCol(0) = ""
        ReadFields
    End If
Case 6 'Add
    FormAction = vbDataActionAddNew
    Call GProcActivateControls(Me, True)
    SetControlEd
    Call GProcSetButtons(Me, vbDataActionAddNew)
    Call GProcClearForm(Me, Rstbl, MaxNo, True)
    If FirstTimeFlag = True Then
        FirstTimeFlag = False
    Else
       txt(0).SetFocus
    End If
    FillCombo
    txt(0) = GProcGenerateId("TblAddLess", "Vno", "Vtype='" & frmTypeStr & "' and VYear=" & gCYear)
    mskDt(2) = GProcGenerateVDt("tblAddLess", "Vdt", "Vtype='" & frmTypeStr & "' and VYear=" & gCYear)
Case 7 'Modify
    FormAction = vbDataActionUpdate
    Hank = ""
    Call GProcChkPass(frmTypeStr, FormAction, txt(0), txt(2), txtTemp(3), gUserName, Hank)
    If gChkPassTrueFalse = True Then
        'If mURecModi = True Then
        '    SearchHelp
        '    DoEvents
            If IsRelatedRecord = False Then
                FormAction = vbDataActionUpdate
                Call GProcActivateControls(Me, True)
                SetControlEd
                Call GProcSetButtons(Me, vbDataActionUpdate)
                txt(0).Enabled = False
                mskDt(2).SetFocus
                OldInvRefNo = txtVSub(12)
            End If
        'End If
    Else
       FormAction = vbDataActionClose
    End If
Case 8 'Delete
    FormAction = vbDataActionDelete
    Hank = ""
    Call GProcChkPass(frmTypeStr, FormAction, txt(0), txt(2), txtTemp(3), gUserName, Hank)
    If gChkPassTrueFalse = True Then
        'If mURecDel = True Then
          If IsRelatedRecord = False Then
            If MsgBox("Do You Really Delete This Record?", vbQuestion + vbYesNo + vbDefaultButton1, "Confirm Delete") = vbYes Then
                If Rstbl.RecordCount > 1 Then
                    '-- Delete Related Record
                    DeleteRelatedRecord
                    Rstbl.Delete
                    cmdBtn_Click (1)
                ElseIf Rstbl.RecordCount = 1 Then
                    '-- Delete Related Record
                    DeleteRelatedRecord
                    Rstbl.Delete
                    Call GProcActivateControls(Me, False)
                    Call GProcSetButtons(Me, vbDataActionCancel)
                    Form_Load
                End If
            End If
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
    If ValidateData = True Then
        'If FormAction = vbDataActionAddNew Then txt(0) = GProcGenerateId(tblName, IdField, "vtype='" & FrmTypeStr & "' and VYear=" & gCYear)
        CalAmount
        FillTxtFromTemp
        '-- Save Record
        GProcSaveRecord Me, Rstbl, FormAction, MaxNo
        SaveRelatedRecord
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
    Call GProcClearForm(Me, Rstbl, MaxNo, True)
    If Rstbl.RecordCount <= 0 Then
        Call GProcActivateControls(Me, False)
        Call GProcSetButtons(Me, vbDataActionCancel)
    Else
        If FormAction = vbDataActionAddNew Then
            Rstbl.MoveLast
        Else
            Rstbl.CancelUpdate
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
    Rstbl.CancelUpdate
    GProcErrorHandler
End Sub

Private Sub cmdLedger_Click(Index As Integer)
On Error GoTo ErrorHandler
If Index = 0 Then
    frmRptAccount.txt(1) = txtTemp(3)
Else
    frmRptAccount.txt(1) = txtTemp(4)
End If
frmMain.mnuRptRSubAccLeg_Click (0)
Exit Sub
ErrorHandler:
    GProcErrorHandler
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
        Select Case KeyCode
            Case 13:
                 SendKeys "{TAB}"
            Case 27:
                If (cmdBtn(13).Enabled = False) Then
                    Unload Me
                End If
        End Select
End Sub
Private Sub Form_Load()
Dim x As String
Me.Top = 300
Me.Left = Screen.Width / 2 - Me.Width / 2
'--- TblAddLess
frmTypeStr = gfrmTypeStr
x = "Select * from " & tblName & " where Vtype='" & frmTypeStr & "' and VYear=" & gCYear
If Trim(OrderField) <> "" Then x = x & " order by " & OrderField
GProcRstOpen Rstbl, x, "O"
FirstTimeFlag = True
If frmTypeStr = "RO" Then  ' --- Consignment Sale
    LblFrmCaption.Caption = "   Consignment \Depot Sales Return"
End If
'---- Button Visiblity Set As Per User Previlage
GProcSetButtonVisiblity Me
cmdBtn_Click (6)
'-- Set Related Rs
SetRelatedRS
End Sub
Private Sub SetRelatedRS()
Dim x As String
'Set rsOS = New Recordset
Set rsVou = New Recordset
Set rstblSub = New Recordset
'---TblVoucher
x = "Select * from tblVoucher where Vtype='" & frmTypeStr & "' and vno=" & txt(0) & " and VYear=" & gCYear & " order by VCtrNo"
GProcRstOpen rsVou, x, "O"
'-- tblVouSub
x = "select * from tblvouSub where vtype='" & frmTypeStr & "' and vno=" & txt(0) & " and VYear=" & gCYear & " order by VSubCtrNo"
GProcRstOpen rstblSub, x, "R"
''---TblOutStanding
'x = "Select * from tblOutStanding where Vtype='" & frmTypeStr & "' and vno=" & txt(0) & " and VYear=" & gCYear
'GProcRstOpen rsOS, x, "O"
End Sub
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
If FormAction = vbDataActionUpdate Then
    MsgBox "You are about to modify the record. Save the record and then exit.", vbInformation + vbOKOnly, Me.Caption
    Cancel = True
    Exit Sub
End If
Rstbl.Close
End Sub

Private Sub mskDt_GotFocus(Index As Integer)
'On Error Resume Next
GProcSelectBox Me.ActiveControl
End Sub
Private Sub mskDt_LostFocus(Index As Integer)
If Index = 2 Then
    If GProcIsDateValid(mskDt(Index)) = False Then
        mskDt(Index).SetFocus
        MsgBox ("Date not within financial year")
        Exit Sub
    End If
End If
ShowAccountBalance
End Sub
Private Sub ShowAccountBalance()
Dim DrAcBal  As Double
Dim CrAcBal As Double
If IsDate(mskDt(2)) Then
    DrAcBal = GProcGetAccountBalance(txtTemp(4), mskDt(2))
    CrAcBal = GProcGetAccountBalance(txtTemp(3), mskDt(2))
    txtDrBalance = IIf(DrAcBal > 0, DrAcBal & " Cr", Abs(DrAcBal) & " Dr")
    txtCrBalance = IIf(CrAcBal > 0, CrAcBal & " Cr", Abs(CrAcBal) & " Dr")
End If
End Sub

Private Sub txt_GotFocus(Index As Integer)
GProcSelectBox Me.ActiveControl
End Sub
Private Sub txt_LostFocus(Index As Integer)
GProcCheckForNumber Me, Rstbl, MaxNo
CalAmount
End Sub
Private Sub CheckForNumberTmp()
Dim i As Long
For i = 6 To 8
If IsNumeric(txtVSub(i)) = False Then txtVSub(i) = "0.00"
Next i
If IsNumeric(txtVSub(5)) = False Then txtVSub(5) = "0"
End Sub
Private Sub CalAmount()
Dim TaxAmt As Double
Dim SubAmt As Double
Dim Amt As Double
Dim ROff As Double
GProcCheckForNumber Me, Rstbl, MaxNo
CheckForNumberTmp
'---- Sub amt =Wt*Rate/RatePer
If txtVSub(8) = 0 Then txtVSub(8) = 1
SubAmt = CDbl(txtVSub(6)) * CDbl(txtVSub(7)) / CDbl(txtVSub(8))
SubAmt = Format(SubAmt, FStr)
'--- (Subamt+Add -Less) * Tax %
TaxAmt = (SubAmt + CDbl(txt(18)) - CDbl(txt(19))) * CDbl(txt(13)) / 100
TaxAmt = Format(TaxAmt, FStr)
'--Roff
ROff = txt(17)
'-- SubAmt+Tax Amt +Add -Less+roff
Amt = SubAmt + TaxAmt + CDbl(txt(18)) - CDbl(txt(19)) + ROff
txt(10) = SubAmt
txt(14) = TaxAmt
txt(11) = Amt
GProcCheckForNumber Me, Rstbl, MaxNo
CheckForNumberTmp
End Sub
Private Sub txt_KeyPress(Index As Integer, KeyAscii As Integer)
If KeyAscii = vbKeyReturn Then Exit Sub
KeyAscii = GProcValidateKey(Rstbl, Index, KeyAscii, txt(Index).Text)
End Sub

Private Sub txttmpInvNo_GotFocus()
GProcSelectBox Me.ActiveControl
End Sub
Private Sub txttmpInvNo_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyReturn Then Exit Sub
GetSalesDetails
KeyAscii = 0
End Sub

Private Sub txtVSub_GotFocus(Index As Integer)
GProcSelectBox Me.ActiveControl
End Sub
Private Sub txtVSub_KeyPress(Index As Integer, KeyAscii As Integer)
If KeyAscii = vbKeyReturn Then Exit Sub
KeyAscii = GProcValidateKey(rstblSub, Index, KeyAscii, txtVSub(Index).Text)
End Sub
Private Sub txtVSub_LostFocus(Index As Integer)
CheckForNumberTmp
CalAmount
End Sub
Private Sub txttmpVSub_GotFocus(Index As Integer)
GProcSelectBox Me.ActiveControl
End Sub
Private Sub txttmpVSub_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyF3 Then
    Select Case Index
        Case 4 '-- item
            GProcShowForm frmMastItem, frmMain.mnuMstfrm(3), Array(frmMastItem.txtTemp(9)), Array(txtTemp(3))
        Case 17 '--Godown
            oldFrmTypeStr = gfrmTypeStr
            gfrmTypeStr = "G"
            GProcShowForm frmMastNarration, frmMain.mnuMstfrm(4), "G"
            gfrmTypeStr = oldFrmTypeStr
    End Select
End If
End Sub

Private Sub txttmpVSub_KeyPress(Index As Integer, KeyAscii As Integer)
Dim LeftPos As Long
Dim TopPos As Long
Dim MillCode As Long
TopPos = txttmpVSub(Index).Top + Me.Top + 650
LeftPos = txttmpVSub(Index).Left + Me.Left
If KeyAscii = 13 Then Exit Sub
Select Case Index
    Case 4 '-- Item
       MillCode = GProcGetColumnValue("TblMastAccount", "AcName", txtTemp(5), "S", "Accode", "N")
       gClsSearch.SearchMultiField "tblMastItem", "ItName,ItStdpack", Array("Count Name", "Weight"), Array(txttmpVSub(Index).Width, 0), "ITMillCode=" & MillCode, Chr(KeyAscii), "ItName", LeftPos, TopPos, True
       KeyAscii = 0
       If gClsSearch.SearchMultiRetCol(0) <> "" Then
            txttmpVSub(Index) = gClsSearch.SearchMultiRetCol(0)
            txtVSub(8) = gClsSearch.SearchMultiRetCol(1)
       End If
    Case 17 '-- Godown
        gClsSearch.SearchMultiField "tblMastNarration", "Narration", Array("Godown"), Array(txttmpVSub(Index).Width, 0), " NarrType='G'", Chr(KeyAscii), "Narration", 0, 0, True
        KeyAscii = 0
        If gClsSearch.SearchMultiRetCol(0) <> "" Then
            txttmpVSub(Index) = gClsSearch.SearchMultiRetCol(0) '--Godown
        End If
End Select
End Sub
Private Sub txttemp_GotFocus(Index As Integer)
GProcSelectBox Me.ActiveControl
End Sub
Private Sub txttemp_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
Dim GName As String
If KeyCode = vbKeyF3 Then
    Select Case Index
        Case 3 '-- Party
            GName = GProcGetColumnValue("TblMastGroup", "AgCode", 90017, "N", "AgName", "S")
            GProcShowForm frmMastAccount, frmMain.mnuMstfrm(0), Array(frmMastAccount.txtTemp(2)), Array(GName)
        Case 4 '-- Dr A/c
            GName = GProcGetColumnValue("TblMastGroup", "AgCode", 90023, "N", "AgName", "S")
            GProcShowForm frmMastAccount, frmMain.mnuMstfrm(0), Array(frmMastAccount.txtTemp(2)), Array(GName)
        Case 5 '-- millcode
            GName = GProcGetColumnValue("TblMastGroup", "AgCode", 90029, "N", "AgName", "S")
            GProcShowForm frmMastAccount, frmMain.mnuMstfrm(0), Array(frmMastAccount.txtTemp(2)), Array(GName)
    End Select
End If
End Sub

Private Sub txtTemp_KeyPress(Index As Integer, KeyAscii As Integer)
Dim LeftPos As Long
Dim TopPos As Long
Dim AgCode As Long
TopPos = txtTemp(Index).Top + Me.Top + 650
LeftPos = txtTemp(Index).Left + Me.Left
If KeyAscii = 13 Then Exit Sub
Select Case Index
   Case 3  '-- Party   Alias Account
        gClsSearch.SearchMultiField "QryMastaccount", "Account,AcName,AcCity,AgName", Array("Party", "AcName", "City", "Group"), Array(3000, 0, 2000, 3000), " (Agcode=90017 or GpCode=90017 or Agcode=90015 or GpCode=90015 or Agcode=90029 or GpCode=90029)", Chr(KeyAscii), "acName", LeftPos, TopPos
        KeyAscii = 0
        If gClsSearch.SearchMultiRetCol(0) <> "" And gClsSearch.SearchMultiRetCol(1) = "" Then
            txtTemp(Index) = gClsSearch.SearchMultiRetCol(0)
        Else
            If gClsSearch.SearchMultiRetCol(1) <> "" Then txtTemp(Index) = gClsSearch.SearchMultiRetCol(1)
        End If
    Case 4 '-- Sales A/c  Alias Account
        gClsSearch.SearchMultiField "QryMastaccount", "Account,AcName,AcCity,AgName", Array("Sales Account", "AcName", "City", "Group"), Array(3000, 0, 2000, 3000), "(AgCode=90023 or GpCode=90023)", Chr(KeyAscii), "AcName", LeftPos, TopPos, True
        KeyAscii = 0
        If gClsSearch.SearchMultiRetCol(0) <> "" And gClsSearch.SearchMultiRetCol(1) = "" Then
            txtTemp(Index) = gClsSearch.SearchMultiRetCol(0)
        Else
            If gClsSearch.SearchMultiRetCol(1) <> "" Then txtTemp(Index) = gClsSearch.SearchMultiRetCol(1)
        End If
    Case 5  '-- Mill  Alias Account
        gClsSearch.SearchMultiField "QryMastaccount", "Account,AcName,AcCity,AgName", Array("Mill", "AcName", "City", "Group"), Array(3000, 0, 2000, 3000), "((AgCode=90029 or GpCode=90029))", Chr(KeyAscii), "AcName", LeftPos, TopPos, True
        KeyAscii = 0
        If gClsSearch.SearchMultiRetCol(0) <> "" And gClsSearch.SearchMultiRetCol(1) = "" Then
            txtTemp(Index) = gClsSearch.SearchMultiRetCol(0)
        Else
            If gClsSearch.SearchMultiRetCol(1) <> "" Then txtTemp(Index) = gClsSearch.SearchMultiRetCol(1)
        End If
End Select
End Sub
Private Sub txtTemp_LostFocus(Index As Integer)
SaveBtnEd
ShowAccountBalance
End Sub
Private Sub SaveBtnEd()
'--- Vno,vdt,Party,A/c,Mill,Amount,Item,Inv.No
If CDbl(txt(0)) = 0 Or GProcIsDateValid(mskDt(2)) = False Or txtTemp(3) = "" Or txtTemp(5) = "" Or txtTemp(4) = "" Or CDbl(txt(11)) = 0 Or txttmpVSub(4) = "" Or txttmpInvNo = "" Then
     cmdBtn(13).Enabled = False
     cmdBtn(13).BackColor = &H80000004
Else
     cmdBtn(13).Enabled = True
     cmdBtn(13).BackColor = &HFFFF80
End If
End Sub
Private Function ValidateData() As Boolean
Dim i As Long
Dim x As String
Dim Y As String
Dim RsSA As Recordset
Dim GName As String
Dim PartyCode As Long
Dim rstAddCheck As Recordset
Dim rstModiCheck As Recordset
    GProcCheckForNumber Me, Rstbl, MaxNo
     '--- Vno
    If CLng(txt(0).Text) <= 0 Then
        MsgBox "Check Voucher No.", vbInformation + vbOKOnly, Me.Caption
        txt(0).SetFocus
        Exit Function
    End If
    '--- Vdt
    If GProcIsDateValid(mskDt(2)) = False Then
        MsgBox "Check Voucher Date.", vbInformation + vbOKOnly, Me.Caption
        mskDt(2).SetFocus
        Exit Function
    End If
    '---- party
    If txtTemp(3) = "" Then
        MsgBox "Check Party.", vbInformation + vbOKOnly, Me.Caption
        txtTemp(3).SetFocus
        Exit Function
    End If
    '---- debitor
    If txtTemp(4) = "" Then
        MsgBox "Check debitor Account.", vbInformation + vbOKOnly, Me.Caption
        txtTemp(4).SetFocus
        Exit Function
    End If
     '--Create Sale A/c
    If GProcGetColumnValue("tblMastAccount", "AcName", txtTemp(4), "S", "AcCode", "N") = 0 Then
       GName = GProcGetColumnValue("TblMastGroup", "AgCode", 90029, "N", "AgName", "S")
       GProcShowForm frmMastAccount, frmMain.mnuMstfrm(0), Array(frmMastAccount.txt(1), frmMastAccount.txtTemp(2)), Array(txtTemp(4), GName)
       Exit Function
    End If
    '---- Item
    If txttmpVSub(4) = "" Then
        MsgBox "Check Item.", vbInformation + vbOKOnly, Me.Caption
       ' txttmpVSub(4).SetFocus
        Exit Function
    End If
    '---- Godown
    If txttmpVSub(17) = "" Then
        MsgBox "Check Godown.", vbInformation + vbOKOnly, Me.Caption
        txttmpVSub(17).SetFocus
        Exit Function
    End If
    If GProcGetColumnValue("TblMastNarration", "Narration", txttmpVSub(17), "S", "Narrcode", "N") = 0 Then
        oldFrmTypeStr = gfrmTypeStr
        gfrmTypeStr = "G"
        GProcShowForm frmMastNarration, frmMain.mnuMstfrm(4), Array(frmMastNarration.txt(2)), Array(txttmpVSub(17))
        gfrmTypeStr = oldFrmTypeStr
        Exit Function
    End If
    '---- Bag
    If CLng(txtVSub(5)) <= 0 Then
        MsgBox "Check Bag.", vbInformation + vbOKOnly, Me.Caption
        txtVSub(5).SetFocus
        Exit Function
    End If
   '---- weight
    If CDbl(txtVSub(6)) <= 0 Then
        MsgBox "Check Weight.", vbInformation + vbOKOnly, Me.Caption
        txtVSub(6).SetFocus
        Exit Function
    End If
    '----Basic Amount
    If CDbl(txt(10)) <= 0 Then
        MsgBox "Check Basic Amount.", vbInformation + vbOKOnly, Me.Caption
        Exit Function
    End If
    '----Rate
    If CDbl(txtVSub(7)) <= 0 Then
        MsgBox "Check Rate.", vbInformation + vbOKOnly, Me.Caption
        txtVSub(7).SetFocus
        Exit Function
    End If
    '----Rate Per
    If CDbl(txtVSub(8)) <= 0 Then
        MsgBox "Check Rate Per.", vbInformation + vbOKOnly, Me.Caption
        txtVSub(8).SetFocus
        Exit Function
    End If
    ' ---Tax Type
    If Cbo(12) = "" Then
        MsgBox "Check Tax Type.", vbInformation + vbOKOnly, Me.Caption
        Cbo(12).SetFocus
        Exit Function
    End If
    '-- Tax %
    If CDbl(txt(13)) > 0 And Cbo(12) = "Tax Free" Then
        MsgBox "Check Tax %.", vbInformation + vbOKOnly, Me.Caption
        txt(13).SetFocus
        Exit Function
    End If
    '----Bill Amount
    If CDbl(txt(11)) <= 0 Then
        MsgBox "Check Voucher Amount.", vbInformation + vbOKOnly, Me.Caption
        Exit Function
    End If
    '--- Valid Invoice
    PartyCode = GProcGetColumnValue("tblMastaccount", "AcName", txtTemp(3), "S", "AcCode", "N")
    x = "Select tblvouSub.* from tblvouSub,tblAddLess where tblvouSub.Vno=tblAddLess.Vno and " _
    & " tblvouSub.Vtype=tblAddLess.VType and tblvouSub.VYear=tblAddLess.VYear and " _
    & " VSubItCtrlNo=" & txtVSub(12) & " and tblAddLess.AdAcDrCode=" & PartyCode
    If GProcRstOpen(RsSA, x, "R") = 0 Then
        MsgBox "Check Invoice No.", vbInformation + vbOKOnly, Me.Caption
        txttmpInvNo.SetFocus
        Exit Function
    Else
        '--Bag
        If CLng(txtVSub(5)) > RsSA.Fields("VSubBag") Then
            MsgBox "Check Bag.Sold Bags are " & RsSA.Fields("VSubBag"), vbInformation + vbOKOnly, Me.Caption
            txtVSub(5).SetFocus
            Exit Function
        End If
        '--Wt
        If CLng(txtVSub(6)) > RsSA.Fields("VSubWt") Then
            MsgBox "Check Wt.Sold Weight is " & RsSA.Fields("VSubWt"), vbInformation + vbOKOnly, Me.Caption
            txtVSub(6).SetFocus
            Exit Function
        End If
    End If
    '----Duplicate Records
    '----Addition Case
    If FormAction = vbDataActionAddNew Then
     x = "Select VNo From " & tblName & " Where VNo=" & CLng(txt(0)) & " and VType='" & frmTypeStr & "' and VYear=" & gCYear
        i = GProcRstOpen(rstAddCheck, x, "R")
        If i > 0 Then
            MsgBox "Record is Available of Voucher No. " & CLng(txt(0)), vbCritical + vbOKOnly, "Duplicate Record"
            txt(0).SetFocus
            Exit Function
        End If
        rstAddCheck.Close
    End If
ValidateData = True
End Function
Private Sub FillCombo()
Call GProcFillCombo(Cbo(12), "select Narration from tblMastNarration where NarrType ='T' order by Narration")
Cbo(12) = "Tax Free" '-- Tax Type
End Sub
Private Sub FillTxtFromTemp()
txt(1) = frmTypeStr '--VType
If IsDate(mskDt(2)) Then '--Vdt
    txt(2) = CDate(mskDt(2))
Else
    txt(2) = ""
End If
'-- Party
txt(3) = GProcGetColumnValue("tblMastAccount", "AcName", txtTemp(3), "S", "AcCode", "N")
'-- Dr Code
txt(4) = GProcGetColumnValue("tblMastAccount", "AcName", txtTemp(4), "S", "AcCode", "N")
'--Mill
txt(5) = GProcGetColumnValue("tblMastAccount", "AcName", txtTemp(5), "S", "AcCode", "N")
'-- TaxType
txt(12) = GProcGetColumnValue("tblMastNarration", "Narration", Cbo(12), "S", "NarrCode", "N")
'--year
txt(23) = gCYear
ShowAccountBalance
End Sub

Private Sub FillTempFromTxt()
If IsDate(txt(2)) Then '--VDt
    mskDt(2) = txt(2)
Else
    mskDt(2) = "__/__/____"
End If
'--Party
txtTemp(3) = GProcGetColumnValue("tblMastAccount", "AcCode", txt(3), "N", "AcName", "S")
'-- Dr Code
txtTemp(4) = GProcGetColumnValue("tblMastAccount", "AcCode", txt(4), "N", "AcName", "S")
'--Mill
txtTemp(5) = GProcGetColumnValue("tblMastAccount", "AcCode", txt(5), "N", "AcName", "S")
'-- TaxType
Cbo(12) = GProcGetColumnValue("tblMastNarration", "NarrCode", txt(12), "N", "Narration", "S")
'-- Fill Related Record
FillRelatedRecord
ShowAccountBalance
End Sub
Private Sub FillRelatedRecord()
'--- Fill Related Control
SetRelatedRS
'==== Voucher
txtVou(7) = rsVou!VNar1
txtVou(8) = rsVou!VNar2
txtVou(9) = rsVou!VNar3
txtVou(10) = rsVou!VNAR4
'==== Vou Sub
txtVSub(5) = rstblSub!vsubBAG
txtVSub(6) = rstblSub!vsubwt
txtVSub(12) = rstblSub!VSubPItCtrlNo
txtVSub(7) = rstblSub!vsubRt
txtVSub(8) = rstblSub!vsubRTPER
txtVSub(10) = rstblSub!vsubLOTNO
'-- Inv No
txttmpInvNo = GProcGetColumnValue("tblAddLess,tblVouSub", "VSubItCtrlNo", txtVSub(12), "N", "AdInvNo", "S", "tblAddLess.Vno=tblVouSub.Vno " _
& " and tblAddLess.Vtype=tblVouSub.Vtype and tblAddLess.VYear=tblVouSub.VYear")
'-- Inv Type
txttmpInvType = GProcGetColumnValue("tblAddLess,tblVouSub", "VSubItCtrlNo", txtVSub(12), "N", "tblAddLess.Vtype", "S", "tblAddLess.Vno=tblVouSub.Vno " _
& " and tblAddLess.Vtype=tblVouSub.Vtype and tblAddLess.VYear=tblVouSub.VYear")
If txttmpInvType = "SY" Then
    txttmpInvType = "Trade Sales"
End If
'-- Inv Year
txttmpInvYear = GProcGetColumnValue("tblAddLess,tblVouSub", "VSubItCtrlNo", txtVSub(12), "N", "tblAddLess.Vyear", "S", "tblAddLess.Vno=tblVouSub.Vno " _
& " and tblAddLess.Vtype=tblVouSub.Vtype and tblAddLess.VYear=tblVouSub.VYear")
'-- Item
txttmpVSub(4) = GProcGetColumnValue("tblMastItem", "Itcode", rstblSub!vsubITCODE, "N", "ItName", "S")
'-- Godown
txttmpVSub(17) = GProcGetColumnValue("tblMastNarration", "NarrCode", rstblSub!VSubGodownCode, "N", "Narration", "S")
End Sub
Private Sub SetControlEd()
txt(10).Enabled = False '--Sub amt
txt(14).Enabled = False '--Tax
txt(11).Enabled = False '--Amt
txttmpInvType.Enabled = False '--Inv Type
txttmpInvYear.Enabled = False '--Inv Year
txttmpVSub(4).Enabled = False '--Item
txtTemp(5).Enabled = False '--Mill
txtCrBalance.Enabled = False
txtDrBalance.Enabled = False
CheckForNumberTmp
End Sub
Private Sub DeleteRelatedRecord()
Dim InvVno As Long
Dim InvVtype As String
Dim InvVYear As Long
'--- Voucher Sub
gCn.Execute "delete from tblVouSub where vtype='" & frmTypeStr & "' and vno=" & txt(0) & " and " _
& " VYear=" & gCYear
'-- Voucher
gCn.Execute "delete from tblVoucher where vtype='" & frmTypeStr & "' and vno=" & txt(0) & " and " _
& " VYear=" & gCYear
'---- Update OutStanding (OutRecAmt)
InvVno = GProcGetColumnValue("tblVousub", "VSubItCtrlNo", CStr(txtVSub(12)), "N", "Vno", "N")
InvVtype = GProcGetColumnValue("tblVousub", "VSubItCtrlNo", CStr(txtVSub(12)), "N", "Vtype", "N")
InvVYear = GProcGetColumnValue("tblVousub", "VSubItCtrlNo", CStr(txtVSub(12)), "N", "VYear", "N")
gCn.Execute "update tblOutstanding set OutRecAmt=0 where Vno=" & InvVno & " and Vtype='" & InvVtype & "' and Vyear=" & InvVYear
End Sub
Private Sub SaveRelatedRecord()
Dim X1 As String
Dim CrCode As Long
Dim DrCode As Long
Dim TaxAcCode As Long
Dim InvVno As Long
Dim InvVtype As String
Dim InvVYear As Long
CrCode = txt(3)
DrCode = txt(4)
TaxAcCode = GProcGetColumnValue("tblMastNarration", "Narration", Cbo(12), "S", "MastTaxSAAcCode", "N")
DeleteRelatedRecord
'=========Voucher Sub
X1 = "select * from tblVouSub where vtype='" & frmTypeStr & "' and vno=" & txt(0) & " and " _
& " VYear=" & gCYear
GProcRstOpen rstblSub, X1, "O"
With rstblSub
    .AddNew
    !Vno = CLng(txt(0).Text) '--Vno
    !Vtype = txt(1) '--Vtype
    !Vdt = CDate(mskDt(2).Text) '--Vdt
    !vsubLOTNO = Val(txtVSub(10).Text) '--Lot No
    !VSubCtrNo = 1 '- Ctr no
    '-- Item code
    !vsubITCODE = GProcGetColumnValue("tblMastItem", "ItName", txttmpVSub(4), "S", "ItCode", "N")
    '-- It ctrl No
    !VSubItCtrlNo = GProcGenerateId("tblVouSub", " VSubItCtrlNo", "VYear=" & gCYear)
    '-- PIT Ctrl No ( Sale It Ctrl No)
    !VSubPItCtrlNo = CLng(txtVSub(12))
    '-- Godown code
    !VSubGodownCode = GProcGetColumnValue("tblMastNarration", "Narration", txttmpVSub(17), "S", "NarrCode", "N")
    !vsubBAG = Val(txtVSub(5)) '-Bag
    !vsubwt = Val(txtVSub(6)) '--Wt
    !vsubRt = Val(txtVSub(7)) '--Rate
    !vsubRTPER = Val(txtVSub(8)) '--Rate Per
    !vsubAMT = Val(txt(10)) '--Sub Amt
    !VYear = Val(txt(23)) '--Year
    .Update
End With
'==========Voucher  For Trade Sales
If frmTypeStr = "RY" Then
    If GProcGetColumnValue("tblMastAccount", "AcCode", CStr(TaxAcCode), "N", "AcName", "S") = "-" Then
         '-- Bill Amt
        GProcCreateVoucher FormAction, CLng(txt(0)), 1, CDate(txt(2)), txt(1), CrCode, DrCode, Val(txt(11)), " ", " ", " ", txtVou(7), txtVou(8), txtVou(9), txtVou(10)
        GProcCreateVoucher FormAction, CLng(txt(0)), 2, CDate(txt(2)), txt(1), DrCode, CrCode, Val(txt(11)) * -1, " ", " ", " ", txtVou(7), txtVou(8), txtVou(9), txtVou(10)
    Else '-- If Tax A/c
        '-- Bill -Tax Amt
        GProcCreateVoucher FormAction, txt(0), 1, txt(2), txt(1), CrCode, DrCode, txt(11) - txt(14), " ", " ", " ", txtVou(7), txtVou(8), txtVou(9), txtVou(10)
        GProcCreateVoucher FormAction, txt(0), 2, txt(2), txt(1), DrCode, CrCode, -(txt(11) - txt(14)), " ", " ", " ", txtVou(7), txtVou(8), txtVou(9), txtVou(10)
        '-- Tax Amt
        If CDbl(txt(14)) > 0 Then
            GProcCreateVoucher FormAction, txt(0), 3, txt(2), txt(1), CrCode, TaxAcCode, txt(14), " ", " ", " ", txtVou(7), txtVou(8), txtVou(9), txtVou(10)
            GProcCreateVoucher FormAction, txt(0), 4, txt(2), txt(1), TaxAcCode, CrCode, -txt(14), " ", " ", " ", txtVou(7), txtVou(8), txtVou(9), txtVou(10)
        End If
    End If
End If
'---- Update OutStanding (OutRecAmt)
InvVno = GProcGetColumnValue("tblVousub", "VSubItCtrlNo", CStr(OldInvRefNo), "N", "Vno", "N")
InvVtype = GProcGetColumnValue("tblVousub", "VSubItCtrlNo", CStr(OldInvRefNo), "N", "Vtype", "N")
InvVYear = GProcGetColumnValue("tblVousub", "VSubItCtrlNo", CStr(OldInvRefNo), "N", "VYear", "N")
gCn.Execute "update tblOutstanding set OutRecAmt=0 where Vno=" & InvVno & " and Vtype='" & InvVtype & "' and Vyear=" & InvVYear

InvVno = GProcGetColumnValue("tblVousub", "VSubItCtrlNo", CStr(txtVSub(12)), "N", "Vno", "N")
InvVtype = GProcGetColumnValue("tblVousub", "VSubItCtrlNo", CStr(txtVSub(12)), "N", "Vtype", "N")
InvVYear = GProcGetColumnValue("tblVousub", "VSubItCtrlNo", CStr(txtVSub(12)), "N", "VYear", "N")
gCn.Execute "update tblOutstanding set OutRecAmt=" & txt(11) & " where Vno=" & InvVno & " and Vtype='" & InvVtype & "' and Vyear=" & InvVYear
End Sub
Private Sub ReadFields()
    FormAction = vbDataActionClose
    GProcShowRecord Me, Rstbl, MaxNo
    FillTempFromTxt
End Sub
Private Function IsRelatedRecord() As Boolean
'-- Audited
If GProcISAudited(txt(0), txt(1), txt(23)) = True Then
    MsgBox "Record is Audited.You can not Modify/Delete this Record.", vbCritical + vbOKOnly, App.Title
    IsRelatedRecord = True
    Exit Function
End If
End Function
'-- Get Sales Details
Private Sub GetSalesDetails()
Dim x As String
Dim QStr As String
Dim RsSA As Recordset
Dim InvVtype As String
If frmTypeStr = "RY" Then
    InvVtype = " TblAddLess.Vtype='SY'"
Else
    InvVtype = " TblAddLess.Vtype in (SO,SD)"
End If
If IsDate(mskDt(2)) = False Then Exit Sub
    '-- Search Sales
    'Inv Type,VSubItCtrlNo, Inv No, Inv Dt, Count, Bag, Wt,Rate
    If gBackEndDB = gBackEndAccess Then
        x = " tblAddLess.Vdt<=Cdate('" & mskDt(2) & "') "
    Else '-- Oracle
        x = " tblAddLess.Vdt<=to_date('" & CDate(mskDt(2)) & "','dd/MM/yyyy') "
    End If
 QStr = "Select tblAddLess.Vtype,VSubItCtrlNo,AdInvNo,tblAddLess.Vdt,ItName,VSubBag,VSubWt,VSubRt " _
   & "from tblVouSub,tblAddLess,tblMastAccount,TblMastItem where tblAddLess.AdAcDrCode=tblMastAccount.AcCode and  " _
   & " tblVouSub.VSubItCode=TblMastItem.ItCode and tblVouSub.Vno=tblAddLess.Vno and tblVouSub.Vtype=tblAddLess.Vtype" _
   & " and tblVouSub.Vyear=tblAddLess.Vyear and " _
   & " TblMastAccount.AcName='" & txtTemp(3) & "' and " & x & "" _
   & " and " & InvVtype & " and VSubITCtrlNo not in (select VSubPItCtrlNo from tblVouSub where vtype='" & frmTypeStr & "' and  vyear<=" & gCYear & " and Vno<>" & txt(0) & ")"
    gClsSearch.SearchMultiField QStr, "", Array("Inv Type", "", "Inv No", "Inv Dt", "Count", "Bag", "Wt", "Rate"), Array(0, 0, 1000, 1500, 2000, 1000, 1000, 1500), "", "", "tblAddLess.Vdt,AdInvNo", 0, 0, False, , True
    If gClsSearch.SearchMultiRetCol(0) <> "" Then
        txtVSub(12) = gClsSearch.SearchMultiRetCol(1)
        txttmpInvNo = gClsSearch.SearchMultiRetCol(2) '--Inv No
        '--- Fill Sales Details
        If gBackEndDB = gBackEndAccess Then
            x = " TblMastAccount as TblMastAc1 "
        Else '-- Oracle
            x = " TblMastAccount TblMastAc1 "
        End If
        QStr = "Select tblVouSub.*,AdTaxCode,AdTaxP,TblMastAccount.AcName as DrAc," _
        & " TblMastAc1.AcName as MillAc,ITName from tblVouSub,tblAddLess,TblMastAccount, " _
        & x & ", TblMastItem where tblVouSub.Vno=tblAddLess.Vno " _
        & " and tblVouSub.Vtype=tblAddLess.Vtype and tblVouSub.VYear=tblAddLess.VYear " _
        & " and tblAddLess.AdAcCrCode=TblMastAccount.AcCode and " _
        & " tblAddLess.AdMillCode=TblMastAc1.AcCode and " _
        & " tblMastItem.ItCode=VSubItCode and VSubITCtrlNo=" & txtVSub(12)
        GProcRstOpen RsSA, QStr, "R"
        With RsSA
            If .EOF Then
                txttmpInvType = ""
                txttmpInvYear = ""
                txtTemp(4) = ""
                txtTemp(5) = ""
                txttmpVSub(4) = ""
                txtVSub(5) = 0
                txtVSub(6) = 0
                txtVSub(7) = 0
                txtVSub(8) = 0
                txttmpVSub(17) = ""
                txtVSub(10) = 0
                Cbo(12) = ""
                txtVSub(12) = 0
                txt(13) = 0
            Else
                txttmpInvType = !Vtype '--Inv Type
                If txttmpInvType = "SY" Then
                    txttmpInvType = "Trade Sales"
                End If
                txttmpInvYear = !VYear '--Inv Year
                txtTemp(4) = !DrAc '--DR A/c
                txtTemp(5) = !MillAc '--Mill A/c
                txttmpVSub(4) = !ItName '--Item
                txtVSub(5) = !vsubBAG '--Bag
                txtVSub(6) = !vsubwt '--Wt
                txtVSub(7) = !vsubRt '--Rate
                txtVSub(8) = !vsubRTPER '--Rate Per
                '--Godown
                txttmpVSub(17) = GProcGetColumnValue("TblMastNarration", "NarrCode", !VSubGodownCode, "N", "Narration", "S")
                txtVSub(10) = !vsubLOTNO & vbNullString '--Lot no
                '-- Tax
                Cbo(12) = GProcGetColumnValue("TblMastNarration", "NarrCode", !AdTaxCode, "N", "Narration", "S")
                '--It Vtrl No
                txtVSub(12) = !VSubItCtrlNo
                txt(13) = !AdTaxP '--Tax %
            End If
          .Close
        End With
    End If
CalAmount
ShowAccountBalance
End Sub
