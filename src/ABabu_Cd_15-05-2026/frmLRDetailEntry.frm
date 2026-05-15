VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form frmLrDetailEntry 
   BackColor       =   &H00FBF2E1&
   BorderStyle     =   0  'None
   Caption         =   "Goods Inward"
   ClientHeight    =   9300
   ClientLeft      =   0
   ClientTop       =   15
   ClientWidth     =   14250
   FillColor       =   &H000040C0&
   FillStyle       =   0  'Solid
   Icon            =   "frmLRDetailEntry.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   9300
   ScaleWidth      =   14250
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox txtBalWt 
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
      Left            =   5340
      MaxLength       =   40
      TabIndex        =   5
      Text            =   "txtBalWt"
      Top             =   11505
      Visible         =   0   'False
      Width           =   1560
   End
   Begin VB.TextBox txtTotWt 
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
      Left            =   5340
      MaxLength       =   40
      TabIndex        =   3
      Text            =   "txtTotWt"
      Top             =   11085
      Visible         =   0   'False
      Width           =   1560
   End
   Begin VB.TextBox txttmpWt 
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
      Left            =   2445
      MaxLength       =   40
      TabIndex        =   4
      Text            =   "txttmpWt"
      Top             =   11505
      Visible         =   0   'False
      Width           =   1500
   End
   Begin VB.TextBox txttmpBag 
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
      Left            =   2445
      MaxLength       =   40
      TabIndex        =   2
      Text            =   "txttmpBag"
      Top             =   11085
      Visible         =   0   'False
      Width           =   1500
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
      Left            =   11805
      Style           =   1  'Graphical
      TabIndex        =   7
      ToolTipText     =   "Exit"
      Top             =   7500
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
      Left            =   9570
      Style           =   1  'Graphical
      TabIndex        =   6
      ToolTipText     =   "To Save Record"
      Top             =   7500
      Width           =   1095
   End
   Begin VB.TextBox txtGrid 
      BackColor       =   &H00C0FFC0&
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
      Left            =   630
      TabIndex        =   1
      Top             =   1485
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
      Left            =   18885
      TabIndex        =   31
      Text            =   "1"
      Top             =   840
      Visible         =   0   'False
      Width           =   420
   End
   Begin VB.TextBox txt 
      Alignment       =   1  'Right Justify
      Height          =   315
      Index           =   2
      Left            =   18825
      TabIndex        =   20
      Text            =   "2"
      Top             =   1380
      Visible         =   0   'False
      Width           =   525
   End
   Begin VB.TextBox txt 
      Alignment       =   1  'Right Justify
      Height          =   315
      Index           =   5
      Left            =   18825
      TabIndex        =   19
      Text            =   "5"
      Top             =   2580
      Visible         =   0   'False
      Width           =   525
   End
   Begin VB.TextBox txt 
      Alignment       =   1  'Right Justify
      Height          =   315
      Index           =   3
      Left            =   18825
      TabIndex        =   18
      Text            =   "3"
      Top             =   1740
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
      Left            =   8550
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   11115
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
      Left            =   9315
      Style           =   1  'Graphical
      TabIndex        =   8
      ToolTipText     =   "To Add New Record"
      Top             =   11040
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.TextBox txt 
      Alignment       =   1  'Right Justify
      Height          =   315
      Index           =   4
      Left            =   18825
      TabIndex        =   17
      Text            =   "4"
      Top             =   2100
      Visible         =   0   'False
      Width           =   525
   End
   Begin VB.TextBox txt 
      Alignment       =   1  'Right Justify
      Height          =   315
      Index           =   7
      Left            =   18825
      TabIndex        =   16
      Text            =   "7"
      Top             =   3300
      Visible         =   0   'False
      Width           =   525
   End
   Begin VB.TextBox txt 
      Alignment       =   1  'Right Justify
      Height          =   315
      Index           =   9
      Left            =   18825
      TabIndex        =   15
      Text            =   "9"
      Top             =   3660
      Visible         =   0   'False
      Width           =   525
   End
   Begin VB.TextBox txt 
      Alignment       =   1  'Right Justify
      Height          =   315
      Index           =   12
      Left            =   18825
      TabIndex        =   14
      Text            =   "12"
      Top             =   4020
      Visible         =   0   'False
      Width           =   525
   End
   Begin VB.TextBox txt 
      Alignment       =   1  'Right Justify
      Height          =   315
      Index           =   20
      Left            =   18825
      TabIndex        =   13
      Text            =   "20"
      Top             =   4380
      Visible         =   0   'False
      Width           =   525
   End
   Begin VB.TextBox txt 
      Alignment       =   1  'Right Justify
      Height          =   315
      Index           =   21
      Left            =   18825
      TabIndex        =   12
      Text            =   "21"
      Top             =   4740
      Visible         =   0   'False
      Width           =   525
   End
   Begin VB.TextBox txt 
      Alignment       =   1  'Right Justify
      Height          =   315
      Index           =   22
      Left            =   18825
      TabIndex        =   11
      Text            =   "22"
      Top             =   5100
      Visible         =   0   'False
      Width           =   525
   End
   Begin VB.TextBox txt 
      Alignment       =   1  'Right Justify
      Height          =   315
      Index           =   23
      Left            =   18825
      TabIndex        =   10
      Text            =   "23"
      Top             =   5340
      Visible         =   0   'False
      Width           =   525
   End
   Begin Crystal.CrystalReport cryReport 
      Bindings        =   "frmLRDetailEntry.frx":058A
      Left            =   0
      Top             =   600
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
      Height          =   8400
      Left            =   180
      TabIndex        =   0
      Top             =   615
      Width           =   8295
      _ExtentX        =   14631
      _ExtentY        =   14817
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
   Begin VB.Label lblBalWt5 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "BalWt5"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   375
      Left            =   12855
      TabIndex        =   77
      Top             =   4860
      Width           =   975
   End
   Begin VB.Label lblBalWt4 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "BalWt4"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   375
      Left            =   12855
      TabIndex        =   76
      Top             =   4080
      Width           =   975
   End
   Begin VB.Label lblBalWt3 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "BalWt3"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   375
      Left            =   12855
      TabIndex        =   75
      Top             =   3360
      Width           =   975
   End
   Begin VB.Label lblBalWt2 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "BalWt2"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   375
      Left            =   12855
      TabIndex        =   74
      Top             =   2700
      Width           =   975
   End
   Begin VB.Label lblBalWt1 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "BalWt1"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   375
      Left            =   12855
      TabIndex        =   73
      Top             =   2040
      Width           =   975
   End
   Begin VB.Label lblBalBag5 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "BalBg5"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   375
      Left            =   12270
      TabIndex        =   72
      Top             =   4860
      Width           =   480
   End
   Begin VB.Label lblBalBag4 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "BalBg4"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   375
      Left            =   12270
      TabIndex        =   71
      Top             =   4080
      Width           =   480
   End
   Begin VB.Label lblBalBag3 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "BalBg3"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   375
      Left            =   12270
      TabIndex        =   70
      Top             =   3360
      Width           =   480
   End
   Begin VB.Label lblBalBag2 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "BalBg2"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   375
      Left            =   12270
      TabIndex        =   69
      Top             =   2700
      Width           =   480
   End
   Begin VB.Label lblBalBag1 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "BalBg1"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   375
      Left            =   12270
      TabIndex        =   68
      Top             =   2040
      Width           =   480
   End
   Begin VB.Label lblLRWt5 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "LrWt5"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   375
      Left            =   11250
      TabIndex        =   67
      Top             =   4860
      Width           =   930
   End
   Begin VB.Label lblLRWt4 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "LrWt4"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   375
      Left            =   11250
      TabIndex        =   66
      Top             =   4080
      Width           =   930
   End
   Begin VB.Label lblLRWt3 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "LrWt3"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   375
      Left            =   11250
      TabIndex        =   65
      Top             =   3360
      Width           =   930
   End
   Begin VB.Label lblLRWt2 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "LrWt2"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   375
      Left            =   11250
      TabIndex        =   64
      Top             =   2700
      Width           =   930
   End
   Begin VB.Label lblLRWt1 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "LrWt1"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   375
      Left            =   11250
      TabIndex        =   63
      Top             =   2040
      Width           =   930
   End
   Begin VB.Label lblLRBag5 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "LrBg5"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   375
      Left            =   10650
      TabIndex        =   62
      Top             =   4860
      Width           =   480
   End
   Begin VB.Label lblLRBag4 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "LrBg4"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   375
      Left            =   10650
      TabIndex        =   61
      Top             =   4080
      Width           =   480
   End
   Begin VB.Label lblLRBag3 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "LrBg3"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   375
      Left            =   10650
      TabIndex        =   60
      Top             =   3360
      Width           =   480
   End
   Begin VB.Label lblLRBag2 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "LrBg2"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   375
      Left            =   10650
      TabIndex        =   59
      Top             =   2700
      Width           =   480
   End
   Begin VB.Label lblLRBag1 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "LRBg1"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   375
      Left            =   10650
      TabIndex        =   58
      Top             =   2040
      Width           =   480
   End
   Begin VB.Label lblInWt5 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Weight5"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   375
      Left            =   9540
      TabIndex        =   57
      Top             =   4860
      Width           =   975
   End
   Begin VB.Label lblInWt4 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Weight4"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   375
      Left            =   9540
      TabIndex        =   56
      Top             =   4080
      Width           =   975
   End
   Begin VB.Label lblInWt3 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Weight3"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   375
      Left            =   9540
      TabIndex        =   55
      Top             =   3360
      Width           =   975
   End
   Begin VB.Label lblInWt2 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Weight2"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   375
      Left            =   9540
      TabIndex        =   54
      Top             =   2700
      Width           =   975
   End
   Begin VB.Label lblInWt1 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Weight1"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   375
      Left            =   9540
      TabIndex        =   53
      Top             =   2040
      Width           =   975
   End
   Begin VB.Label lblInBag5 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Bags5"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   375
      Left            =   9000
      TabIndex        =   52
      Top             =   4860
      Width           =   480
   End
   Begin VB.Label lblInBag4 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Bags4"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   375
      Left            =   9000
      TabIndex        =   51
      Top             =   4080
      Width           =   480
   End
   Begin VB.Label lblInBag3 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Bags3"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   375
      Left            =   9000
      TabIndex        =   50
      Top             =   3360
      Width           =   480
   End
   Begin VB.Label lblInBag2 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Bags2"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   375
      Left            =   9000
      TabIndex        =   49
      Top             =   2700
      Width           =   480
   End
   Begin VB.Label lblInBag1 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Bags1"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   375
      Left            =   9000
      TabIndex        =   48
      Top             =   2040
      Width           =   480
   End
   Begin VB.Label lbl1 
      BackStyle       =   0  'Transparent
      Caption         =   "5)"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   375
      Index           =   18
      Left            =   8685
      TabIndex        =   47
      Top             =   4860
      Width           =   210
   End
   Begin VB.Label lbl1 
      BackStyle       =   0  'Transparent
      Caption         =   "4)"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   375
      Index           =   17
      Left            =   8685
      TabIndex        =   46
      Top             =   4080
      Width           =   210
   End
   Begin VB.Label lbl1 
      BackStyle       =   0  'Transparent
      Caption         =   "3)"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   375
      Index           =   16
      Left            =   8685
      TabIndex        =   45
      Top             =   3360
      Width           =   210
   End
   Begin VB.Label lbl1 
      BackStyle       =   0  'Transparent
      Caption         =   "2)"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   375
      Index           =   15
      Left            =   8685
      TabIndex        =   44
      Top             =   2700
      Width           =   210
   End
   Begin VB.Label lbl1 
      BackStyle       =   0  'Transparent
      Caption         =   "1)"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   375
      Index           =   14
      Left            =   8685
      TabIndex        =   43
      Top             =   2040
      Width           =   210
   End
   Begin VB.Label lbl1 
      BackStyle       =   0  'Transparent
      Caption         =   "BALANCE"
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
      Height          =   375
      Index           =   13
      Left            =   12630
      TabIndex        =   42
      Top             =   780
      Width           =   1110
   End
   Begin VB.Label lbl1 
      BackStyle       =   0  'Transparent
      Caption         =   "LR DETAIL"
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
      Height          =   375
      Index           =   12
      Left            =   10950
      TabIndex        =   41
      Top             =   780
      Width           =   1110
   End
   Begin VB.Label lbl1 
      BackStyle       =   0  'Transparent
      Caption         =   "INWARD"
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
      Height          =   375
      Index           =   11
      Left            =   9420
      TabIndex        =   40
      Top             =   780
      Width           =   1110
   End
   Begin VB.Line Line8 
      BorderColor     =   &H00000080&
      X1              =   12810
      X2              =   12810
      Y1              =   1140
      Y2              =   6000
   End
   Begin VB.Line Line7 
      BorderColor     =   &H00000080&
      X1              =   12195
      X2              =   12195
      Y1              =   675
      Y2              =   6015
   End
   Begin VB.Line Line6 
      BorderColor     =   &H00000080&
      X1              =   11190
      X2              =   11190
      Y1              =   1140
      Y2              =   6000
   End
   Begin VB.Line Line5 
      BorderColor     =   &H00000080&
      X1              =   10590
      X2              =   10590
      Y1              =   660
      Y2              =   6000
   End
   Begin VB.Line Line4 
      BorderColor     =   &H00000080&
      X1              =   9510
      X2              =   9510
      Y1              =   1140
      Y2              =   6000
   End
   Begin VB.Line Line3 
      BorderColor     =   &H00000080&
      X1              =   8580
      X2              =   13860
      Y1              =   1140
      Y2              =   1140
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00000080&
      X1              =   8925
      X2              =   8925
      Y1              =   660
      Y2              =   6000
   End
   Begin VB.Label lbl1 
      BackStyle       =   0  'Transparent
      Caption         =   "Weight"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   375
      Index           =   10
      Left            =   13095
      TabIndex        =   39
      Top             =   1260
      Width           =   750
   End
   Begin VB.Label lbl1 
      BackStyle       =   0  'Transparent
      Caption         =   "Bags"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   375
      Index           =   9
      Left            =   12240
      TabIndex        =   38
      Top             =   1260
      Width           =   450
   End
   Begin VB.Label lbl1 
      BackStyle       =   0  'Transparent
      Caption         =   "Weight"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   375
      Index           =   8
      Left            =   11415
      TabIndex        =   37
      Top             =   1260
      Width           =   750
   End
   Begin VB.Label lbl1 
      BackStyle       =   0  'Transparent
      Caption         =   "Bags"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   375
      Index           =   7
      Left            =   10620
      TabIndex        =   36
      Top             =   1260
      Width           =   510
   End
   Begin VB.Label lbl1 
      BackStyle       =   0  'Transparent
      Caption         =   "Weight"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   375
      Index           =   6
      Left            =   9810
      TabIndex        =   35
      Top             =   1260
      Width           =   750
   End
   Begin VB.Label lbl1 
      BackStyle       =   0  'Transparent
      Caption         =   "Bags"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   375
      Index           =   5
      Left            =   9000
      TabIndex        =   34
      Top             =   1260
      Width           =   480
   End
   Begin VB.Label lbl1 
      BackStyle       =   0  'Transparent
      Caption         =   "Sr."
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   375
      Index           =   4
      Left            =   8640
      TabIndex        =   33
      Top             =   1260
      Width           =   300
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00000080&
      X1              =   8580
      X2              =   13860
      Y1              =   1620
      Y2              =   1620
   End
   Begin VB.Shape Shape3 
      BorderColor     =   &H00000080&
      Height          =   5355
      Left            =   8595
      Top             =   660
      Width           =   5340
   End
   Begin VB.Shape Shape1 
      FillColor       =   &H00FED8FC&
      FillStyle       =   0  'Solid
      Height          =   990
      Left            =   8550
      Shape           =   4  'Rounded Rectangle
      Top             =   7185
      Width           =   5235
   End
   Begin VB.Shape Shape5 
      BorderColor     =   &H00EFB17A&
      BorderWidth     =   5
      FillColor       =   &H00EFB17A&
      Height          =   9240
      Left            =   0
      Top             =   30
      Width           =   14190
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
      ForeColor       =   &H0000FFFF&
      Height          =   495
      Left            =   11340
      TabIndex        =   32
      Top             =   0
      Width           =   2475
   End
   Begin VB.Label Label27 
      BackStyle       =   0  'Transparent
      Caption         =   "Landing Cost Per Bag"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   735
      Left            =   12255
      TabIndex        =   30
      Top             =   10860
      Width           =   975
   End
   Begin VB.Label Label26 
      BackStyle       =   0  'Transparent
      Caption         =   "Landing Cost"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   255
      Left            =   12375
      TabIndex        =   29
      Top             =   11220
      Width           =   1455
   End
   Begin VB.Label Label25 
      BackStyle       =   0  'Transparent
      Caption         =   "Other Less"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   255
      Left            =   11415
      TabIndex        =   28
      Top             =   11340
      Width           =   1215
   End
   Begin VB.Label Label24 
      BackStyle       =   0  'Transparent
      Caption         =   "Other Add"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   255
      Left            =   11295
      TabIndex        =   27
      Top             =   10980
      Width           =   1095
   End
   Begin VB.Label Label22 
      BackStyle       =   0  'Transparent
      Caption         =   "Godown Rent"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   255
      Left            =   11175
      TabIndex        =   26
      Top             =   11220
      Width           =   1575
   End
   Begin VB.Label Label21 
      BackStyle       =   0  'Transparent
      Caption         =   "Freight"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   255
      Left            =   11175
      TabIndex        =   25
      Top             =   11100
      Width           =   1335
   End
   Begin VB.Label Label20 
      BackStyle       =   0  'Transparent
      Caption         =   "Total Amount"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   255
      Left            =   11415
      TabIndex        =   24
      Top             =   10920
      Width           =   1455
   End
   Begin VB.Label Label19 
      BackStyle       =   0  'Transparent
      Caption         =   "Landing Cost Details :-"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   375
      Left            =   11100
      TabIndex        =   23
      Top             =   11340
      Width           =   3135
   End
   Begin VB.Label Label13 
      BackStyle       =   0  'Transparent
      Caption         =   "Ass.Value"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   255
      Left            =   11055
      TabIndex        =   22
      Top             =   11220
      Width           =   1095
   End
   Begin VB.Label Label6 
      BackColor       =   &H00EFB17A&
      Caption         =   "   L.R. Details"
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
      Left            =   -15
      TabIndex        =   21
      Top             =   15
      Width           =   14835
   End
   Begin VB.Shape Shape2 
      BackColor       =   &H00808080&
      BorderColor     =   &H004D177D&
      FillColor       =   &H00EDFDFE&
      FillStyle       =   0  'Solid
      Height          =   8685
      Left            =   0
      Top             =   465
      Width           =   14055
   End
End
Attribute VB_Name = "frmLrDetailEntry"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Const tblName As String = "tblSITLRDetails"
Const IdField As String = "Vno"
Const OrderField As String = "VFrim,Vdt,Vno"
Const MaxNo As Long = 10
Dim varBookMark As Variant
Dim FirstTimeFlag As Boolean
Dim OldgFrmTypeStr As String
Dim FormAction As Integer
Dim frmTypeStr As String
Dim rstbl As Recordset '-- TblAddLess
Dim rstblSub As Recordset '-- TblSITLRDetail
Dim tmpAutoLRNo As String '- Auto LR No
Dim tmpAutoItSrNo As String
Dim tmpAutoLrDt As String
Dim AutoBalBg As String
Dim Inbg1 As Long
Dim Inbg2 As Long
Dim Inbg3 As Long
Dim Inbg4 As Long
Dim Inbg5 As Long
Dim InWt1 As Double
Dim InWt2 As Double
Dim InWt3 As Double
Dim InWt4 As Double
Dim InWt5 As Double

Dim CalBg1 As Long
Dim CalBg2 As Long
Dim CalBg3 As Long
Dim CalBg4 As Long
Dim CalBg5 As Long

Dim CalWt1 As Double
Dim CalWt2 As Double
Dim CalWt3 As Double
Dim CalWt4 As Double
Dim CalWt5 As Double
Dim ChkBal As Boolean
Dim i As Long


Private Sub cmdBtn_Click(Index As Integer)
'On Error GoTo ErrorRoutine
Dim i As Long
Select Case Index
Case 10 'Exit
'    DoEvents
    Unload Me
    Exit Sub
Case 13 'Save
    If ValidateData = True Then
        If Trim(txtGrid) <> "" Then
            msGrid.text = txtGrid
            DoEvents
            txtGrid = ""
        End If
        txtGrid.Visible = False
        CalAmount
        FillTxtFromTemp
        SaveRelatedRecord
        Unload Me
        Exit Sub
    End If
End Select
Exit Sub
ErrorRoutine:
'    Rstbl.CancelUpdate
    GProcErrorHandler
End Sub
Private Sub SaveRelatedRecord()
Dim i As Long, X As String
Dim MillCode As Long
Dim RsG As Recordset
'--- Save Grid Records
CheckForNumberGrid
   X = "Select PurMillCode from tblPurch where VType='" & gfrmBagInwVType & "' and Vno=" & gfrmBagInwVNo & " and VFirm  = '" & gCCode & "' and VYear=" & gCYear & " "
   i = GProcRstOpen(RsG, X, "R")
   If i > 0 Then
      MillCode = RsG.Fields(0)
   End If
'--- Save Grid Records To tblBags
gCn.Execute "delete from tblSITLRDetails where VType='" & gfrmBagInwVType & "' and Vno=" & gfrmBagInwVNo & " and VFirm  = '" & gCCode & "' and VYear=" & gCYear & ""
X = "select * from tblSITLRDetails where VType='" & gfrmBagInwVType & "' and Vno=" & gfrmBagInwVNo & " and VFirm  = '" & gCCode & "'   and VYear=" & gCYear & ""
GProcRstOpen rstblSub, X, "O"
With msGrid
    i = 1
    Do While i < .rows
       If .TextMatrix(i, 2) <> "" Then '- LrNo
            rstblSub.AddNew
            '-- SNo
            rstblSub!Vno = gfrmBagInwVNo    '--Inw No
            rstblSub!Vtype = gfrmBagInwVType    '--Inw Type
            rstblSub!VYear = gCYear   '--Inw Year
            rstblSub!VFirm = gCCode   '--Inw Firm
            rstblSub!LrDItSrNo = .TextMatrix(i, 0)   '--Item Sr.No. In TblPurchSub
            rstblSub!LrDCtrNo = i
            rstblSub!LrDItCode = GProcGetColumnValue("tblMastItem", "ItName", .TextMatrix(i, 1), "S", "ItCode", "N", " ItMillCode=" & MillCode) '--Item Code
            rstblSub!LrDBag = CLng(.TextMatrix(i, 5))   '-- Bags
            rstblSub!LrDWt = CDbl(.TextMatrix(i, 6))   '-- Wt.
            rstblSub!LrDLRNo = .TextMatrix(i, 2) & vbNullString    '-- LR No
            If IsDate(.TextMatrix(i, 3)) = True Then
               rstblSub!LrdLRDate = .TextMatrix(i, 3)   '-- LR Date
            Else
               rstblSub!LrdLRDate = Null
            End If
            rstblSub!LrdLRAmt = CDbl(.TextMatrix(i, 4))   '-- LR Amount
            rstblSub!SlBillNo = .TextMatrix(i, 7) & vbNullString    '--
            rstblSub!SlVno = CLng(.TextMatrix(i, 12))   '--
            rstblSub!SlVtype = .TextMatrix(i, 13) & vbNullString  '--
            rstblSub!SlVyear = CLng(.TextMatrix(i, 14))   '--
            rstblSub!SlVfirm = .TextMatrix(i, 15) & vbNullString
            rstblSub.Update
        End If
        i = i + 1
     Loop
 End With
 rstblSub.Close
End Sub
Private Sub Form_Load()
    Dim X As String
    Dim RsG As Recordset
    Me.Top = 300
    Me.Left = Screen.Width / 2 - Me.Width / 2
    FirstTimeFlag = True
'    SetGrid
    FormAction = 5
    '---Grid
    lblInBag1.Caption = ""
    lblInWt1.Caption = ""
    Inbg1 = 0
    InWt1 = 0
    
    lblInBag2.Caption = ""
    lblInWt2.Caption = ""
    Inbg2 = 0
    InWt2 = 0
    
    lblInBag3.Caption = ""
    lblInWt3.Caption = ""
    Inbg3 = 0
    InWt3 = 0
    
    lblInBag4.Caption = ""
    lblInWt4.Caption = ""
    Inbg4 = 0
    InWt4 = 0
    
    lblInBag5.Caption = ""
    lblInWt5.Caption = ""
    Inbg5 = 0
    InWt5 = 0
    
    
    X = "Select PurSubBag,PurSubWt,PurSubItSrNo from tblPurchSub where VType='" & gfrmBagInwVType & "' and Vno=" & gfrmBagInwVNo & " and VYear=" & gCYear & " and VFirm  = '" & gCCode & "' order by PurSubItSrNo"
    If GProcRstOpen(RsG, X, "R") > 0 Then
        RsG.MoveFirst
        Do While Not RsG.EOF
           If RsG.Fields(2) = 1 Then
              lblInBag1.Caption = CStr(RsG.Fields(0))
              lblInWt1.Caption = Format(CStr(RsG.Fields(1)), WtStr)
              Inbg1 = RsG.Fields(0)
              InWt1 = RsG.Fields(1)
           End If
           If RsG.Fields(2) = 2 Then
              lblInBag2.Caption = CStr(RsG.Fields(0))
              lblInWt2.Caption = Format(CStr(RsG.Fields(1)), WtStr)
              Inbg2 = RsG.Fields(0)
              InWt2 = RsG.Fields(1)
           End If
           If RsG.Fields(2) = 3 Then
              lblInBag3.Caption = CStr(RsG.Fields(0))
              lblInWt3.Caption = Format(CStr(RsG.Fields(1)), WtStr)
              Inbg3 = RsG.Fields(0)
              InWt3 = RsG.Fields(1)
           End If
           If RsG.Fields(2) = 4 Then
              lblInBag4.Caption = CStr(RsG.Fields(0))
              lblInWt4.Caption = Format(CStr(RsG.Fields(1)), WtStr)
              Inbg4 = RsG.Fields(0)
              InWt4 = RsG.Fields(1)
           End If
           If RsG.Fields(2) = 5 Then
              lblInBag5.Caption = CStr(RsG.Fields(0))
              lblInWt5.Caption = Format(CStr(RsG.Fields(1)), WtStr)
              Inbg5 = RsG.Fields(0)
              InWt5 = RsG.Fields(1)
           End If
           RsG.MoveNext
        Loop
    End If
    FillRelatedRecord
    CalAmount
    msGrid_EnterCell
End Sub
'---------------- Grid
Private Sub SetGrid()
Dim X As String
'--- Set Grid
Set rstblSub = New Recordset
X = "select LrDItSrNo,Itname,LrDLRNo,LrdLRDate,LrdLRAmt,LrDBag,LrDWt,SlBillNo,vno,VType,VYear,Vfirm,SlVno,SlVtype,SlVyear,SlVfirm from tblSITLRDetails,tblMastItem  where " _
& " tblSITLRDetails.LrDItCode=TblMastItem.ItCode and VType='" & gfrmBagInwVType & "' and Vno=" & gfrmBagInwVNo & " and VFirm  = '" & gCCode & "'  and VYear=" & gCYear & " order by VNo,LrDItSrNo,LrDCtrNo,LrDLRNo"
GProcRstOpen rstblSub, X, "R"
With msGrid
    .FormatString = ">Sr.|<Count Description              |<L.R.No       |<L.R.Date    |>L.R.Amount    |>Bags |>Weight       |>Sale Bill No |VNo |VType|Vyear|Vfirm|SlVNo |SlVType|SlVyear|SlVfirm"
    .ColWidth(4) = 0 '-Lr.Amount
    .ColWidth(8) = 0 '-VNo
    .ColWidth(9) = 0 '-VType
    .ColWidth(10) = 0 '-Vyear
    .ColWidth(11) = 0 '-Vfirm
    .ColWidth(12) = 0 '-SaleVno
    .ColWidth(13) = 0 '-saleVtype
    .ColWidth(14) = 0 '-SaleVyear
    .ColWidth(15) = 0 '-SaleFirm
    .rows = 2
    GridAddNew
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
                 .TextMatrix(.Row, i) = Null & vbNullString
            Case Else
                 .TextMatrix(.Row, i) = ""
        End Select
    Next i
'ReNumberGridSrno
.Refresh
End With
End Sub

Private Sub msGrid_Click()
On Error Resume Next
    msGrid_EnterCell
End Sub
Private Sub msGrid_GotFocus()
cmdBtn(10).Cancel = False
End Sub
Private Sub MSGrid_LostFocus()
CalAmount
End Sub
Private Sub msGrid_EnterCell()
If FormAction <> vbDataActionCancel Then
    CheckForNumberTxtGrid
    DoEvents
    txtGrid = ""
    DoEvents
    AutoBalBg = "lblBalBag" + Trim(CStr(msGrid.Row))
    If msGrid.TextMatrix(msGrid.Row, 2) = "" And IsNumeric(tmpAutoLRNo) = True Then msGrid.TextMatrix(msGrid.Row, 2) = tmpAutoLRNo  ''-Auto LR No
    If msGrid.TextMatrix(msGrid.Row, 0) = "0" And IsNumeric(tmpAutoItSrNo) = True Then msGrid.TextMatrix(msGrid.Row, 0) = tmpAutoItSrNo  ''-Auto It Sr No
    If msGrid.TextMatrix(msGrid.Row, 3) = "" And IsDate(tmpAutoLrDt) = True Then msGrid.TextMatrix(msGrid.Row, 3) = tmpAutoLrDt  ''-Auto LR No
    If msGrid.text <> "" Then txtGrid.text = msGrid.text
'    If msGrid.TextMatrix(msGrid.Row, 7) = "" Then msGrid.TextMatrix(msGrid.Row, 7) = 0
    '--- Lock TextBox
    With msGrid
        Select Case .Col
            Case 0: '---- Sr No
                txtGrid.Locked = False
                DoEvents
            Case 1:
                .Col = .Col + 1
            Case 7 '-- Sale Bill
                txtGrid.Locked = True
            Case Else
                If msGrid.TextMatrix(msGrid.Row, 7) = "" Then
                   txtGrid.Locked = False
                Else
                   txtGrid.Locked = True
                End If
        End Select
    End With
    DoEvents
    msGrid_KeyPress (0)
    DoEvents
End If
End Sub
Private Sub msGrid_KeyPress(KeyAscii As Integer)
With msGrid
    If KeyAscii = 27 Then '----Esc Key
        KeyAscii = 0
        If (Val(.TextMatrix(.Row, 6)) = 0) And .rows > 2 Then
            .RemoveItem (.Row)
            txtGrid.text = ""
        End If
        txtGrid.Visible = False
        cmdBtn(13).SetFocus
        Exit Sub
    End If
End With
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
            .Col = 0
        Else    '-- Next col
            .Col = .Col + 1
        End If
    End With
End If
DoEvents
'---- Set TxtGrid
If msGrid.text <> "" Then txtGrid.text = msGrid.text
If msGrid.Col < rstblSub.Fields.Count - 1 Then
   If msGrid.Col = 4 Then
      msGrid.Col = 5
   End If
    txtGrid.Move msGrid.CellLeft + msGrid.Left, msGrid.CellTop + msGrid.Top, msGrid.CellWidth, msGrid.CellHeight
End If
txtGrid.Visible = True
txtGrid.Enabled = True
If txtGrid.Visible = True Then txtGrid.SetFocus
DoEvents
If KeyAscii <> 0 Then
    Select Case rstblSub.Fields(msGrid.Col).Type
        Case adChar, adVarWChar
            If msGrid.Col = 1 Then  '-- Count ,godown
                txtGrid_KeyPress (KeyAscii)
            Else
                txtGrid.text = txtGrid.text & Chr(KeyAscii)
            End If
        Case Else
            txtGrid.text = Chr(GProcValidateKey(rstblSub, msGrid.Col, KeyAscii, txtGrid.text)) 'Chr(GProcNumberWithDecimal(KeyAscii, txtGrid))
    End Select
End If
DoEvents
'If msGrid.Text <> "" Then
'    txtGrid.SelStart = Len(txtGrid.Text) + 1
'Else
'    txtGrid.SelStart = 1
'End If
DoEvents
If txtGrid.Visible = True Then txtGrid.SetFocus
End Sub
Private Sub msGrid_LeaveCell()
If txtGrid.Visible Then
    msGrid.text = txtGrid.text
    DoEvents
    txtGrid.text = ""
    txtGrid.Visible = False
    If IsNumeric(msGrid.TextMatrix(msGrid.Row, 2)) = True Then ''-Auto LR No
        tmpAutoLRNo = CDbl(msGrid.TextMatrix(msGrid.Row, 2)) + 1
    Else
        tmpAutoLRNo = tmpAutoLRNo
    End If
    If CDbl(msGrid.TextMatrix(msGrid.Row, 0)) > 0 Then ''-Auto ItSr No
        tmpAutoItSrNo = msGrid.TextMatrix(msGrid.Row, 0)
    Else
        tmpAutoItSrNo = tmpAutoItSrNo
    End If
    If IsDate(msGrid.TextMatrix(msGrid.Row, 3)) = True Then  ''-Auto Lr Date
        tmpAutoLrDt = msGrid.TextMatrix(msGrid.Row, 3)
    Else
        tmpAutoLrDt = tmpAutoLrDt
    End If
End If
DoEvents
CalAmount
CalGridAmt
DoEvents
End Sub
Private Sub txtGrid_GotFocus()
    txtGrid.Alignment = ProcSetAlignment(rstblSub, msGrid.Col)
'    DoEvents
    Call GProcSelectBox(txtGrid)
End Sub
Private Sub txtGrid_KeyPress(KeyAscii As Integer)
Dim LeftPos As Long
Dim TopPos As Long
Dim MillCode As Long
Dim X As String
Dim RsG As Recordset
Dim i As Integer
TopPos = txtGrid.Top + Me.Top + 650
LeftPos = txtGrid.Left + Me.Left
DoEvents
With msGrid
    If KeyAscii <> 13 And KeyAscii <> 27 Then
        '--- Search
        Select Case .Col
            Case 0 '-- Count
            Case Else
                KeyAscii = GProcValidateKey(rstblSub, msGrid.Col, KeyAscii, txtGrid)
        End Select
   End If
   If KeyAscii = 13 Then CheckForNumberTxtGrid
    If KeyAscii = 13 Then
        CheckForNumberTxtGrid
        If txtGrid = "" And .Col = 2 Then
           KeyAscii = 27
        End If
    End If
   
    '--- Navigate Row ,Col
    If KeyAscii = 13 And .Col < 6 Then
        .Col = .Col + 1
        If .Col = 2 Then
            If IsNumeric(txtGrid) = True Then '-Auto LR No
                tmpAutoLRNo = CDbl(txtGrid)
            Else
                tmpAutoLRNo = ""
            End If
        End If
        If .Col = 0 Then
            If IsNumeric(txtGrid) = True Then '-Auto LR No
                tmpAutoItSrNo = CLng(txtGrid)
            Else
                tmpAutoItSrNo = ""
            End If
        End If
        If .Col = 3 Then
            If IsDate(txtGrid) = True Then
                tmpAutoLrDt = txtGrid.text
            Else
                tmpAutoLrDt = ""
            End If
        End If
    ElseIf KeyAscii = 13 And .Col = 6 Then
        If .Row = (.rows - 1) Then
            .rows = .rows + 1
            .Row = .Row + 1
            GridAddNew
        Else
            .Row = .Row + 1
        End If
        .Col = 0
    End If
'    txtGrid.SetFocus
    '-- On ESC
    With msGrid
        If KeyAscii = 27 Then '----Esc Key
            KeyAscii = 0
            If (Val(.TextMatrix(.Row, 6)) = 0) And .rows > 2 Then
                .RemoveItem (.Row)
                txtGrid.text = ""
            End If
            txtGrid.Visible = False
            If cmdBtn(13).Enabled = True Then
               cmdBtn(13).SetFocus
            Else
               cmdBtn(10).SetFocus
            End If
            Exit Sub
        End If
    End With
End With
If txtGrid.Enabled = True And txtGrid.Visible = True Then txtGrid.SetFocus
End Sub
Private Sub txtGrid_KeyDown(KeyCode As Integer, Shift As Integer)
'--- Delete Row On Ctrl+Y
If KeyCode = vbKeyY And Shift = 2 And msGrid.rows > 2 Then
    If (msGrid.TextMatrix(msGrid.Row, 7) <> "") Then
        MsgBox "Sale Bill Avaiable Can not delete..."
    Else
        msGrid.RemoveItem (msGrid.Row)
        txtGrid.Visible = False
        txtGrid.text = ""
        CalGridAmt
        CalAmount
    End If
'    msGrid = 4

'    ReNumberGridSrno
    CalAmount
    Exit Sub
End If
End Sub
Private Sub CalGridAmt()
Dim X As String
Dim RsG As Recordset
Dim StdPack As Double
'If msGrid.TextMatrix(msGrid.Row, 1) = "" And Val(msGrid.TextMatrix(msGrid.Row, 0)) > 0 Then
If Val(msGrid.TextMatrix(msGrid.Row, 0)) > 0 Then
   X = "Select itname from tblPurchSub,tblMastItem where VType='" & gfrmBagInwVType & "' and Vno=" & gfrmBagInwVNo & " and VFirm  = '" & gCCode & "'  and VYear=" & gCYear & " and  tblPurchSub.PurSubItCode=TblMastItem.ItCode and " _
       & " PurSubItSrNo = " & CLng(msGrid.TextMatrix(msGrid.Row, 0))
       i = GProcRstOpen(RsG, X, "R")
       If i > 0 Then
          msGrid.TextMatrix(msGrid.Row, 1) = RsG.Fields(0)
      End If
End If

If Val(msGrid.TextMatrix(msGrid.Row, 6)) = 0 And Val(msGrid.TextMatrix(msGrid.Row, 0)) > 0 Then
   X = "Select ItStdpack from tblPurchSub,tblMastItem where VType='" & gfrmBagInwVType & "' and Vno=" & gfrmBagInwVNo & " and VFirm  = '" & gCCode & "'  and VYear=" & gCYear & " and  tblPurchSub.PurSubItCode=TblMastItem.ItCode   and " _
       & " PurSubItSrNo = " & CLng(msGrid.TextMatrix(msGrid.Row, 0))
       i = GProcRstOpen(RsG, X, "R")
       If i > 0 Then
          StdPack = RsG.Fields(0)
       Else
          StdPack = 0
       End If
    msGrid.TextMatrix(msGrid.Row, 6) = msGrid.TextMatrix(msGrid.Row, 5) * StdPack
End If
If IsDate(msGrid.TextMatrix(msGrid.Row, 3)) = False And IsDate(gfrmBagLrDt) = True Then
    msGrid.TextMatrix(msGrid.Row, 3) = gfrmBagLrDt
End If
SaveBtnEd
End Sub
Private Sub CheckForNumber()
GProcCheckForNumber Me, rstbl, MaxNo
End Sub
Private Sub CheckForNumberGrid()
Dim RowIndex As Long
With rstblSub
For RowIndex = 1 To msGrid.rows - 1
    For i = 1 To .Fields.Count - 1
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
Private Function ProcSetAlignment(pRstbl As Recordset, pColIndex As Long)
Select Case pRstbl.Fields(pColIndex).Type
    Case adInteger, adDouble, 2, 3, 131, adDecimal
        ProcSetAlignment = 1
    Case Else
        ProcSetAlignment = 0
End Select
End Function
'-- Grid Required
Private Function ValidateGrid() As Boolean
Dim LRNo As String
Dim i As Long
Dim MillCode As Long
Dim RowIndex As Long
Dim oldFrmTypeStr As String
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
                    If RowIndex > 1 And Trim(.TextMatrix(RowIndex, i)) = "" And CDbl(.TextMatrix(RowIndex, 5)) = 0 And CDbl(.TextMatrix(RowIndex, 6)) = 0 Then
                        .RemoveItem (RowIndex)
                        GoTo Validate
                    End If
                    If Trim(.TextMatrix(RowIndex, i)) = "" Then
                        MsgBox "Check Count.", vbCritical + vbOKOnly, Me.Caption
                        .Col = i
                        msGrid.SetFocus
                        Exit Function
                    End If
                Case 2 '-- lrno
                    If Trim(.TextMatrix(RowIndex, i)) = "" Then
                        MsgBox "Check L.R. No.", vbCritical + vbOKOnly, Me.Caption
                        .Col = i
                        msGrid.SetFocus
                        Exit Function
                    End If
                Case 3 '-- Lr Date
                    If IsDate(.TextMatrix(RowIndex, i)) = False Then
                        MsgBox "Check LR. Date.", vbCritical + vbOKOnly, Me.Caption
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
                Case 6 '-- Wt
                    If CDbl(.TextMatrix(RowIndex, i)) = 0 Then
                        MsgBox "Check Weight.", vbCritical + vbOKOnly, Me.Caption
                        .Col = i
                        msGrid.SetFocus
                        Exit Function
                    End If
           End Select
        Next i
      
'--- Duplicate LR No ,It Ctrl  No (Booking)
        LRNo = .TextMatrix(RowIndex, 2)
        For i = 1 To .rows - 1
'           --- Duplicate LR No
            If i <> RowIndex And .TextMatrix(i, 2) = LRNo Then
                MsgBox "Duplicate LR No.", vbCritical + vbOKOnly, Me.Caption
                .Col = 2
                msGrid.SetFocus
                Exit Function
            End If
        Next i
      
      
      Next RowIndex
End With





ValidateGrid = True
'ReNumberGridSrno
End Function
Private Sub SaveBtnEd()
'--- Vno,vdt,Party,A/c,Mill,Amount
Dim X As String
Dim RsG As Recordset
Dim PurWt As Double

If (InWt1 + InWt2 + InWt3 + InWt4 + InWt5) - (CalWt1 + CalWt2 + CalWt3 + CalWt4 + CalWt5) <> 0 Or ChkBal = True Then
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
Dim ExpAcCode As Long
Dim rstAddCheck As Recordset
Dim rstModiCheck As Recordset
    '--- Validate Grid
    If ValidateGrid = False Then
        msGrid.SetFocus
        Exit Function
    End If
    ValidateData = True
End Function
Private Sub FillTxtFromTemp()
End Sub
Private Sub FillTempFromTxt()
End Sub
Private Sub FillRelatedRecord()
Dim RowIndex As Long
Dim ColIndex As Long
Dim TotBag As Long
Dim TotWt As Double
tmpAutoLRNo = "" ''-Auto LR No
tmpAutoItSrNo = ""
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
                If ColIndex = 4 Then TotBag = TotBag + CLng(msGrid.TextMatrix(RowIndex, 4)) '-- Tot Bag
                If ColIndex = 5 Then TotWt = TotWt + CDbl(msGrid.TextMatrix(RowIndex, 5)) '-- Tot Wt
            Next ColIndex
            RowIndex = RowIndex + 1
            .MoveNext
            If Not .EOF Then msGrid.rows = msGrid.rows + 1
        Loop
    End If
End With
'txttmpTotBag = TotBag '-- Tot Bag
'txttmpTotWt = Format(TotWt, WtStr) '-- Tot Wt
'--- Fill Related Control
End Sub
Private Sub ReadFields()
'FormAction = vbDataActionClose
GProcShowRecord Me, rstbl, MaxNo
FillTempFromTxt
End Sub
Private Sub CalAmount()
Dim TaxAmt As Double
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
Dim IsExemptMill As Long
Dim MillAcCode  As Long
ChkBal = False
CalGridAmt
CalBg1 = 0
CalBg2 = 0
CalBg3 = 0
CalBg4 = 0
CalBg5 = 0

CalWt1 = 0
CalWt2 = 0
CalWt3 = 0
CalWt4 = 0
CalWt5 = 0

RowIndex = 1
With msGrid
    Do While RowIndex < .rows
        If Val(.TextMatrix(RowIndex, 0)) = 0 Then
           .TextMatrix(RowIndex, 0) = 1
        End If
        '-- bag
        If Val(.TextMatrix(RowIndex, 0)) = 1 Then
            CalBg1 = CalBg1 + Val(.TextMatrix(RowIndex, 5))
        End If
        '-- Weight
        If Val(.TextMatrix(RowIndex, 0)) = 1 Then
            CalWt1 = Format(CalWt1 + Val(.TextMatrix(RowIndex, 6)), WtStr)
        End If
        
        If Val(.TextMatrix(RowIndex, 0)) = 2 Then
            CalBg2 = CalBg2 + Val(.TextMatrix(RowIndex, 5))
        End If
        '-- Weight
        If Val(.TextMatrix(RowIndex, 0)) = 2 Then
            CalWt2 = Format(CalWt2 + Val(.TextMatrix(RowIndex, 6)), WtStr)
        End If
        
        If Val(.TextMatrix(RowIndex, 0)) = 3 Then
            CalBg3 = CalBg3 + Val(.TextMatrix(RowIndex, 5))
        End If
        '-- Weight
        If Val(.TextMatrix(RowIndex, 0)) = 3 Then
            CalWt3 = Format(CalWt3 + Val(.TextMatrix(RowIndex, 6)), WtStr)
        End If
        
        If Val(.TextMatrix(RowIndex, 0)) = 4 Then
            CalBg4 = CalBg4 + Val(.TextMatrix(RowIndex, 5))
        End If
        '-- Weight
        If Val(.TextMatrix(RowIndex, 0)) = 4 Then
            CalWt4 = Format(CalWt4 + Val(.TextMatrix(RowIndex, 6)), WtStr)
        End If
        
        If Val(.TextMatrix(RowIndex, 0)) = 5 Then
            CalBg5 = CalBg5 + Val(.TextMatrix(RowIndex, 5))
        End If
        '-- Weight
        If Val(.TextMatrix(RowIndex, 0)) = 5 Then
            CalWt5 = Format(CalWt5 + Val(.TextMatrix(RowIndex, 6)), WtStr)
        End If
        RowIndex = RowIndex + 1
    Loop
End With

If CalBg1 > 0 Then
   lblLRBag1.Caption = CStr(CalBg1)
Else
   lblLRBag1.Caption = ""
End If
If CalWt1 > 0 Then
   lblLRWt1.Caption = Format(CStr(CalWt1), WtStr)
Else
   lblLRWt1.Caption = ""
End If

If CalBg2 > 0 Then
   lblLRBag2.Caption = CStr(CalBg2)
Else
   lblLRBag2.Caption = ""
End If
If CalWt2 > 0 Then
   lblLRWt2.Caption = Format(CStr(CalWt2), WtStr)
Else
   lblLRWt2.Caption = ""
End If

If CalBg3 > 0 Then
   lblLRBag3.Caption = CStr(CalBg3)
Else
   lblLRBag3.Caption = ""
End If
If CalWt3 > 0 Then
   lblLRWt3.Caption = Format(CStr(CalWt3), WtStr)
Else
   lblLRWt3.Caption = ""
End If


If CalBg4 > 0 Then
   lblLRBag4.Caption = CStr(CalBg4)
Else
   lblLRBag4.Caption = ""
End If
If CalWt4 > 0 Then
   lblLRWt4.Caption = Format(CStr(CalWt4), WtStr)
Else
   lblLRWt4.Caption = ""
End If

If CalBg5 > 0 Then
   lblLRBag5.Caption = CStr(CalBg5)
Else
   lblLRBag5.Caption = ""
End If
If CalWt5 > 0 Then
   lblLRWt5.Caption = Format(CStr(CalWt5), WtStr)
Else
   lblLRWt5.Caption = ""
End If



If Inbg1 - CalBg1 <> 0 Then
   lblBalBag1.Caption = CStr(Inbg1 - CalBg1)
Else
   lblBalBag1.Caption = ""
End If

If InWt1 - CalWt1 <> 0 Then
   lblBalWt1.Caption = Format(CStr(InWt1 - CalWt1), WtStr)
Else
   lblBalWt1.Caption = ""
End If

If Inbg2 - CalBg2 <> 0 Then
   lblBalBag2.Caption = CStr(Inbg2 - CalBg2)
Else
   lblBalBag2.Caption = ""
End If
If InWt2 - CalWt2 <> 0 Then
   lblBalWt2.Caption = Format(CStr(InWt2 - CalWt2), WtStr)
Else
   lblBalWt2.Caption = ""
End If

If Inbg3 - CalBg3 <> 0 Then
   lblBalBag3.Caption = CStr(Inbg3 - CalBg3)
Else
   lblBalBag3.Caption = ""
End If
If InWt3 - CalWt3 <> 0 Then
   lblBalWt3.Caption = Format(CStr(InWt3 - CalWt3), WtStr)
Else
   lblBalWt3.Caption = ""
End If

If Inbg4 - CalBg4 <> 0 Then
   lblBalBag4.Caption = CStr(Inbg4 - CalBg4)
Else
   lblBalBag4.Caption = ""
End If
If InWt4 - CalWt4 <> 0 Then
   lblBalWt4.Caption = Format(CStr(InWt4 - CalWt4), WtStr)
Else
   lblBalWt4.Caption = ""
End If

If Inbg5 - CalBg5 <> 0 Then
   lblBalBag5.Caption = CStr(Inbg5 - CalBg5)
Else
   lblBalBag5.Caption = ""
End If
If InWt5 - CalWt5 <> 0 Then
   lblBalWt5.Caption = Format(CStr(InWt5 - CalWt5), WtStr)
Else
   lblBalWt5.Caption = ""
End If


If (Inbg1 - CalBg1 < 0) Or (InWt1 - CalWt1 < 0) Then
   ChkBal = True
End If
If (Inbg2 - CalBg2 < 0) Or (InWt2 - CalWt2 < 0) Then
   ChkBal = True
End If
If (Inbg3 - CalBg3 < 0) Or (InWt3 - CalWt3 < 0) Then
   ChkBal = True
End If
If (Inbg4 - CalBg4 < 0) Or (InWt4 - CalWt4 < 0) Then
   ChkBal = True
End If
If (Inbg5 - CalBg5 < 0) Or (InWt5 - CalWt5 < 0) Then
   ChkBal = True
End If
SaveBtnEd
End Sub
'''''-- Is Related GatePass
Private Function IsRelatedGatePass() As Boolean
''''Dim RsG As Recordset
''''Dim X As String
'''''-- Is Gate Pass
''''X = "Select Vno,Vdt from tblVouSub where Vtype='GP' and " _
''''& " VSubPItCtrlNo in (select VSubItCtrlNo from tblVouSub where vtype='" & frmTypeStr & "' and vno=" & txt(0) & " and " _
''''& " VYear=" & gCYear & ")"
''''GProcRstOpen RsG, X, "R"
''''If RsG.EOF Then
''''Else
''''    IsRelatedGatePass = True
''''   ' MsgBox "You can not Modify/Delete this record.Related Gate Pass No. " & RsG.Fields(0) & " is exists.", vbCritical + vbOKOnly, App.Title
''''    Exit Function
''''End If
''''RsG.Close
End Function
'--- Is Related Record
Private Function IsRelatedRecord() As Boolean
Dim RsR As Recordset
Dim RsG As Recordset
Dim X As String
'-- Audited
If GProcISAudited(txt(0), txt(1), txt(23)) = True Then
    MsgBox "Record is Audited.You can not Modify/Delete this Record.", vbCritical + vbOKOnly, App.Title
    IsRelatedRecord = True
    Exit Function
End If
RsR.Close
End Function
