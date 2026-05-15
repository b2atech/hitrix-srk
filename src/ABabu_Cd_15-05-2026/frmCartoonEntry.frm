VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form frmCartoonEntry 
   BackColor       =   &H00C2E3FC&
   BorderStyle     =   0  'None
   Caption         =   "Goods Inward"
   ClientHeight    =   9135
   ClientLeft      =   0
   ClientTop       =   15
   ClientWidth     =   15405
   FillColor       =   &H000040C0&
   FillStyle       =   0  'Solid
   Icon            =   "frmCartoonEntry.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   9135
   ScaleWidth      =   15405
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox txtGrid 
      BackColor       =   &H00C0FFC0&
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
      Left            =   540
      TabIndex        =   121
      Top             =   1680
      Visible         =   0   'False
      Width           =   1035
   End
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
      Left            =   5580
      MaxLength       =   40
      TabIndex        =   4
      Text            =   "txtBalWt"
      Top             =   10965
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
      Left            =   5580
      MaxLength       =   40
      TabIndex        =   2
      Text            =   "txtTotWt"
      Top             =   10545
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
      Left            =   2685
      MaxLength       =   40
      TabIndex        =   3
      Text            =   "txttmpWt"
      Top             =   10965
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
      Left            =   2685
      MaxLength       =   40
      TabIndex        =   1
      Text            =   "txttmpBag"
      Top             =   10545
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
      Left            =   11745
      Style           =   1  'Graphical
      TabIndex        =   6
      ToolTipText     =   "Exit"
      Top             =   8025
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
      Left            =   10545
      Style           =   1  'Graphical
      TabIndex        =   5
      ToolTipText     =   "To Save Record"
      Top             =   8025
      Width           =   1095
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
      Left            =   14715
      TabIndex        =   30
      Text            =   "1"
      Top             =   9630
      Visible         =   0   'False
      Width           =   420
   End
   Begin VB.TextBox txt 
      Alignment       =   1  'Right Justify
      Height          =   315
      Index           =   2
      Left            =   15210
      TabIndex        =   19
      Text            =   "2"
      Top             =   9675
      Visible         =   0   'False
      Width           =   525
   End
   Begin VB.TextBox txt 
      Alignment       =   1  'Right Justify
      Height          =   315
      Index           =   5
      Left            =   16830
      TabIndex        =   18
      Text            =   "5"
      Top             =   9585
      Visible         =   0   'False
      Width           =   525
   End
   Begin VB.TextBox txt 
      Alignment       =   1  'Right Justify
      Height          =   315
      Index           =   3
      Left            =   15750
      TabIndex        =   17
      Text            =   "3"
      Top             =   9630
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
      Left            =   8790
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   10575
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
      Left            =   9555
      Style           =   1  'Graphical
      TabIndex        =   7
      ToolTipText     =   "To Add New Record"
      Top             =   10500
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.TextBox txt 
      Alignment       =   1  'Right Justify
      Height          =   315
      Index           =   4
      Left            =   16290
      TabIndex        =   16
      Text            =   "4"
      Top             =   9585
      Visible         =   0   'False
      Width           =   525
   End
   Begin VB.TextBox txt 
      Alignment       =   1  'Right Justify
      Height          =   315
      Index           =   7
      Left            =   19170
      TabIndex        =   15
      Text            =   "7"
      Top             =   8775
      Visible         =   0   'False
      Width           =   525
   End
   Begin VB.TextBox txt 
      Alignment       =   1  'Right Justify
      Height          =   315
      Index           =   9
      Left            =   17415
      TabIndex        =   14
      Text            =   "9"
      Top             =   9630
      Visible         =   0   'False
      Width           =   525
   End
   Begin VB.TextBox txt 
      Alignment       =   1  'Right Justify
      Height          =   315
      Index           =   12
      Left            =   17910
      TabIndex        =   13
      Text            =   "12"
      Top             =   9585
      Visible         =   0   'False
      Width           =   525
   End
   Begin VB.TextBox txt 
      Alignment       =   1  'Right Justify
      Height          =   315
      Index           =   20
      Left            =   18495
      TabIndex        =   12
      Text            =   "20"
      Top             =   9585
      Visible         =   0   'False
      Width           =   525
   End
   Begin VB.TextBox txt 
      Alignment       =   1  'Right Justify
      Height          =   315
      Index           =   21
      Left            =   18855
      TabIndex        =   11
      Text            =   "21"
      Top             =   9585
      Visible         =   0   'False
      Width           =   525
   End
   Begin VB.TextBox txt 
      Alignment       =   1  'Right Justify
      Height          =   315
      Index           =   22
      Left            =   18675
      TabIndex        =   10
      Text            =   "22"
      Top             =   9675
      Visible         =   0   'False
      Width           =   525
   End
   Begin VB.TextBox txt 
      Alignment       =   1  'Right Justify
      Height          =   315
      Index           =   23
      Left            =   19260
      TabIndex        =   9
      Text            =   "23"
      Top             =   8325
      Visible         =   0   'False
      Width           =   525
   End
   Begin Crystal.CrystalReport cryReport 
      Bindings        =   "frmCartoonEntry.frx":058A
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
      Height          =   8220
      Left            =   225
      TabIndex        =   0
      Top             =   585
      Width           =   7920
      _ExtentX        =   13970
      _ExtentY        =   14499
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
   Begin VB.Line Line9 
      BorderColor     =   &H00000080&
      X1              =   8280
      X2              =   15015
      Y1              =   4440
      Y2              =   4440
   End
   Begin VB.Shape Shape5 
      BorderColor     =   &H00C00000&
      BorderWidth     =   5
      FillColor       =   &H00000080&
      Height          =   9090
      Left            =   0
      Top             =   0
      Width           =   15390
   End
   Begin VB.Label lblBalWt 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "BalWt1"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   285
      Index           =   1
      Left            =   13860
      TabIndex        =   120
      Top             =   1320
      Width           =   1020
   End
   Begin VB.Label lblBalWt 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "BalWt2"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   285
      Index           =   2
      Left            =   13860
      TabIndex        =   119
      Top             =   1620
      Width           =   1020
   End
   Begin VB.Label lblBalWt 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "BalWt3"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   285
      Index           =   3
      Left            =   13860
      TabIndex        =   118
      Top             =   1935
      Width           =   1020
   End
   Begin VB.Label lblBalWt 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "BalWt4"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   285
      Index           =   4
      Left            =   13860
      TabIndex        =   117
      Top             =   2220
      Width           =   1020
   End
   Begin VB.Label lblBalWt 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "BalWt5"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   285
      Index           =   5
      Left            =   13860
      TabIndex        =   116
      Top             =   2535
      Width           =   1020
   End
   Begin VB.Label lblBalWt 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "BalWt6"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   285
      Index           =   6
      Left            =   13860
      TabIndex        =   115
      Top             =   2850
      Width           =   1020
   End
   Begin VB.Label lblBalWt 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "BalWt7"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   285
      Index           =   7
      Left            =   13860
      TabIndex        =   114
      Top             =   3165
      Width           =   1020
   End
   Begin VB.Label lblBalWt 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "BalWt8"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   285
      Index           =   8
      Left            =   13860
      TabIndex        =   113
      Top             =   3480
      Width           =   1020
   End
   Begin VB.Label lblBalWt 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "BalWt9"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   285
      Index           =   9
      Left            =   13860
      TabIndex        =   112
      Top             =   3795
      Width           =   1020
   End
   Begin VB.Label lblBalWt 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "BalWt10"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   285
      Index           =   10
      Left            =   13860
      TabIndex        =   111
      Top             =   4110
      Width           =   1020
   End
   Begin VB.Label lblBalBag 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "BalBg2"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   285
      Index           =   2
      Left            =   12060
      TabIndex        =   110
      Top             =   1605
      Width           =   570
   End
   Begin VB.Label lblBalBag 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "BalBg3"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   285
      Index           =   3
      Left            =   12060
      TabIndex        =   109
      Top             =   1920
      Width           =   570
   End
   Begin VB.Label lblBalBag 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "BalBg4"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   285
      Index           =   4
      Left            =   12060
      TabIndex        =   108
      Top             =   2205
      Width           =   570
   End
   Begin VB.Label lblBalBag 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "BalBg5"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   285
      Index           =   5
      Left            =   12060
      TabIndex        =   107
      Top             =   2520
      Width           =   570
   End
   Begin VB.Label lblBalBag 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "BalBg6"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   285
      Index           =   6
      Left            =   12060
      TabIndex        =   106
      Top             =   2835
      Width           =   570
   End
   Begin VB.Label lblBalBag 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "BalBg7"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   285
      Index           =   7
      Left            =   12060
      TabIndex        =   105
      Top             =   3150
      Width           =   570
   End
   Begin VB.Label lblBalBag 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "BalBg8"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   285
      Index           =   8
      Left            =   12060
      TabIndex        =   104
      Top             =   3465
      Width           =   570
   End
   Begin VB.Label lblBalBag 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "BalBg1"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   285
      Index           =   1
      Left            =   12060
      TabIndex        =   103
      Top             =   1320
      Width           =   570
   End
   Begin VB.Label lbl1 
      BackStyle       =   0  'Transparent
      Caption         =   "Totall balance           :"
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
      Index           =   55
      Left            =   8505
      TabIndex        =   102
      Top             =   5460
      Width           =   2190
   End
   Begin VB.Label lbl1 
      BackStyle       =   0  'Transparent
      Caption         =   "Total Cartoon Detail  :"
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
      Index           =   53
      Left            =   8520
      TabIndex        =   101
      Top             =   5040
      Width           =   2190
   End
   Begin VB.Label lbl1 
      BackStyle       =   0  'Transparent
      Caption         =   "Total Inward              :"
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
      Index           =   54
      Left            =   8505
      TabIndex        =   100
      Top             =   4620
      Width           =   2325
   End
   Begin VB.Label TBlWt 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "TBlWt"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   285
      Left            =   11970
      TabIndex        =   99
      Top             =   5505
      Width           =   1650
   End
   Begin VB.Label TBlBg 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "TBlBg"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   285
      Left            =   10665
      TabIndex        =   98
      Top             =   5505
      Width           =   1065
   End
   Begin VB.Label TLrWt 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "TLrWt"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   285
      Left            =   11970
      TabIndex        =   97
      Top             =   5085
      Width           =   1650
   End
   Begin VB.Label TLrBg 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "TLrBg"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   285
      Left            =   10665
      TabIndex        =   96
      Top             =   5085
      Width           =   1065
   End
   Begin VB.Label TInWt 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "TInWt"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   285
      Left            =   11970
      TabIndex        =   95
      Top             =   4665
      Width           =   1650
   End
   Begin VB.Label TInBg 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "TInBg"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   285
      Left            =   10665
      TabIndex        =   94
      Top             =   4665
      Width           =   1065
   End
   Begin VB.Label lblBalBag 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "BalBg10"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   285
      Index           =   10
      Left            =   12060
      TabIndex        =   93
      Top             =   4095
      Width           =   570
   End
   Begin VB.Label lblBalBag 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "BalBg9"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   285
      Index           =   9
      Left            =   12060
      TabIndex        =   92
      Top             =   3780
      Width           =   570
   End
   Begin VB.Label lblLRWt 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "LrWt10"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   285
      Index           =   10
      Left            =   11130
      TabIndex        =   91
      Top             =   4095
      Width           =   840
   End
   Begin VB.Label lblLRWt 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "LrWt9"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   285
      Index           =   9
      Left            =   11130
      TabIndex        =   90
      Top             =   3780
      Width           =   840
   End
   Begin VB.Label lblLRWt 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "LrWt8"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   285
      Index           =   8
      Left            =   11130
      TabIndex        =   89
      Top             =   3465
      Width           =   840
   End
   Begin VB.Label lblLRWt 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "LrWt7"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   285
      Index           =   7
      Left            =   11130
      TabIndex        =   88
      Top             =   3150
      Width           =   840
   End
   Begin VB.Label lblLRWt 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "LrWt6"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   285
      Index           =   6
      Left            =   11130
      TabIndex        =   87
      Top             =   2835
      Width           =   840
   End
   Begin VB.Label lblLRWt 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "LrWt5"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   285
      Index           =   5
      Left            =   11130
      TabIndex        =   86
      Top             =   2520
      Width           =   840
   End
   Begin VB.Label lblLRBag 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "LrBg10"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   285
      Index           =   10
      Left            =   10485
      TabIndex        =   85
      Top             =   4095
      Width           =   525
   End
   Begin VB.Label lblLRBag 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "LrBg9"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   285
      Index           =   9
      Left            =   10485
      TabIndex        =   84
      Top             =   3780
      Width           =   525
   End
   Begin VB.Label lblLRBag 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "LrBg8"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   285
      Index           =   8
      Left            =   10485
      TabIndex        =   83
      Top             =   3465
      Width           =   525
   End
   Begin VB.Label lblLRBag 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "LrBg7"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   285
      Index           =   7
      Left            =   10485
      TabIndex        =   82
      Top             =   3150
      Width           =   525
   End
   Begin VB.Label lblLRBag 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "LrBg6"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   285
      Index           =   6
      Left            =   10485
      TabIndex        =   81
      Top             =   2835
      Width           =   525
   End
   Begin VB.Label lblLRBag 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "LrBg5"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   285
      Index           =   5
      Left            =   10485
      TabIndex        =   80
      Top             =   2520
      Width           =   525
   End
   Begin VB.Label lblInWt 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Weight10"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   285
      Index           =   10
      Left            =   9525
      TabIndex        =   79
      Top             =   4095
      Width           =   840
   End
   Begin VB.Label lblInWt 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Weight9"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   285
      Index           =   9
      Left            =   9525
      TabIndex        =   78
      Top             =   3780
      Width           =   840
   End
   Begin VB.Label lblInWt 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Weight8"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   285
      Index           =   8
      Left            =   9525
      TabIndex        =   77
      Top             =   3465
      Width           =   840
   End
   Begin VB.Label lblInWt 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Weight7"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   285
      Index           =   7
      Left            =   9525
      TabIndex        =   76
      Top             =   3150
      Width           =   840
   End
   Begin VB.Label lblInWt 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Weight6"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   285
      Index           =   6
      Left            =   9525
      TabIndex        =   75
      Top             =   2835
      Width           =   840
   End
   Begin VB.Label lblInWt 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Weight5"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   285
      Index           =   5
      Left            =   9525
      TabIndex        =   74
      Top             =   2520
      Width           =   840
   End
   Begin VB.Label lbl1 
      BackStyle       =   0  'Transparent
      Caption         =   "7)"
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
      Index           =   19
      Left            =   8505
      TabIndex        =   73
      Top             =   3150
      Width           =   210
   End
   Begin VB.Label lblInBag 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Bags10"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   285
      Index           =   10
      Left            =   8865
      TabIndex        =   72
      Top             =   4095
      Width           =   525
   End
   Begin VB.Label lblInBag 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Bags9"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   285
      Index           =   9
      Left            =   8865
      TabIndex        =   71
      Top             =   3780
      Width           =   525
   End
   Begin VB.Label lblInBag 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Bags8"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   285
      Index           =   8
      Left            =   8865
      TabIndex        =   70
      Top             =   3465
      Width           =   525
   End
   Begin VB.Label lblInBag 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Bags7"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   285
      Index           =   7
      Left            =   8865
      TabIndex        =   69
      Top             =   3150
      Width           =   525
   End
   Begin VB.Label lblInBag 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Bags6"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   285
      Index           =   6
      Left            =   8865
      TabIndex        =   68
      Top             =   2835
      Width           =   525
   End
   Begin VB.Label lblInBag 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Bags5"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   285
      Index           =   5
      Left            =   8865
      TabIndex        =   67
      Top             =   2520
      Width           =   525
   End
   Begin VB.Line Line7 
      BorderColor     =   &H00000080&
      X1              =   12015
      X2              =   12015
      Y1              =   540
      Y2              =   4440
   End
   Begin VB.Label lbl1 
      BackStyle       =   0  'Transparent
      Caption         =   "10)"
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
      Index           =   22
      Left            =   8415
      TabIndex        =   66
      Top             =   4095
      Width           =   255
   End
   Begin VB.Label lbl1 
      BackStyle       =   0  'Transparent
      Caption         =   "9)"
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
      Index           =   21
      Left            =   8505
      TabIndex        =   65
      Top             =   3780
      Width           =   210
   End
   Begin VB.Label lbl1 
      BackStyle       =   0  'Transparent
      Caption         =   "8)"
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
      Index           =   20
      Left            =   8505
      TabIndex        =   64
      Top             =   3465
      Width           =   210
   End
   Begin VB.Label lbl1 
      BackStyle       =   0  'Transparent
      Caption         =   "6)"
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
      Index           =   3
      Left            =   8505
      TabIndex        =   63
      Top             =   2835
      Width           =   210
   End
   Begin VB.Label lblLRWt 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "LrWt4"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   285
      Index           =   4
      Left            =   11130
      TabIndex        =   62
      Top             =   2205
      Width           =   840
   End
   Begin VB.Label lblLRWt 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "LrWt3"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   285
      Index           =   3
      Left            =   11130
      TabIndex        =   61
      Top             =   1920
      Width           =   840
   End
   Begin VB.Label lblLRWt 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "LrWt2"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   285
      Index           =   2
      Left            =   11130
      TabIndex        =   60
      Top             =   1605
      Width           =   840
   End
   Begin VB.Label lblLRWt 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "LrWt1"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   285
      Index           =   1
      Left            =   11130
      TabIndex        =   59
      Top             =   1305
      Width           =   840
   End
   Begin VB.Label lblLRBag 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "LrBg4"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   285
      Index           =   4
      Left            =   10485
      TabIndex        =   58
      Top             =   2205
      Width           =   525
   End
   Begin VB.Label lblLRBag 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "LrBg3"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   285
      Index           =   3
      Left            =   10485
      TabIndex        =   57
      Top             =   1920
      Width           =   525
   End
   Begin VB.Label lblLRBag 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "LrBg2"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   285
      Index           =   2
      Left            =   10485
      TabIndex        =   56
      Top             =   1605
      Width           =   525
   End
   Begin VB.Label lblLRBag 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "LrBg1"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   285
      Index           =   1
      Left            =   10485
      TabIndex        =   55
      Top             =   1305
      Width           =   525
   End
   Begin VB.Label lblInWt 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Weight4"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   285
      Index           =   4
      Left            =   9525
      TabIndex        =   54
      Top             =   2205
      Width           =   840
   End
   Begin VB.Label lblInWt 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Weight3"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   285
      Index           =   3
      Left            =   9525
      TabIndex        =   53
      Top             =   1920
      Width           =   840
   End
   Begin VB.Label lblInWt 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Weight2"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   285
      Index           =   2
      Left            =   9525
      TabIndex        =   52
      Top             =   1605
      Width           =   840
   End
   Begin VB.Label lblInWt 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Weight1"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   285
      Index           =   1
      Left            =   9525
      TabIndex        =   51
      Top             =   1305
      Width           =   840
   End
   Begin VB.Label lblInBag 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Bags4"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   285
      Index           =   4
      Left            =   8865
      TabIndex        =   50
      Top             =   2205
      Width           =   525
   End
   Begin VB.Label lblInBag 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Bags3"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   285
      Index           =   3
      Left            =   8865
      TabIndex        =   49
      Top             =   1920
      Width           =   525
   End
   Begin VB.Label lblInBag 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Bags2"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   285
      Index           =   2
      Left            =   8865
      TabIndex        =   48
      Top             =   1605
      Width           =   525
   End
   Begin VB.Label lblInBag 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Bags1"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   285
      Index           =   1
      Left            =   8865
      TabIndex        =   47
      Top             =   1305
      Width           =   525
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
      Left            =   8520
      TabIndex        =   46
      Top             =   2505
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
      Left            =   8520
      TabIndex        =   45
      Top             =   2220
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
      Left            =   8520
      TabIndex        =   44
      Top             =   1905
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
      Left            =   8520
      TabIndex        =   43
      Top             =   1605
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
      Left            =   8520
      TabIndex        =   42
      Top             =   1305
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
      Left            =   12330
      TabIndex        =   41
      Top             =   630
      Width           =   1110
   End
   Begin VB.Label lbl1 
      BackStyle       =   0  'Transparent
      Caption         =   "CARTOON DTL"
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
      Left            =   10515
      TabIndex        =   40
      Top             =   630
      Width           =   1515
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
      Left            =   9210
      TabIndex        =   39
      Top             =   630
      Width           =   1110
   End
   Begin VB.Line Line8 
      BorderColor     =   &H00000080&
      X1              =   12645
      X2              =   12645
      Y1              =   945
      Y2              =   4440
   End
   Begin VB.Line Line6 
      BorderColor     =   &H00000080&
      X1              =   11025
      X2              =   11025
      Y1              =   945
      Y2              =   4440
   End
   Begin VB.Line Line5 
      BorderColor     =   &H00000080&
      X1              =   10395
      X2              =   10395
      Y1              =   540
      Y2              =   4440
   End
   Begin VB.Line Line4 
      BorderColor     =   &H00000080&
      X1              =   9405
      X2              =   9420
      Y1              =   945
      Y2              =   4440
   End
   Begin VB.Line Line3 
      BorderColor     =   &H00000080&
      X1              =   8325
      X2              =   15060
      Y1              =   945
      Y2              =   960
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00000080&
      X1              =   8775
      X2              =   8775
      Y1              =   540
      Y2              =   4440
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
      Left            =   14130
      TabIndex        =   38
      Top             =   975
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
      Left            =   12120
      TabIndex        =   37
      Top             =   975
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
      Left            =   11205
      TabIndex        =   36
      Top             =   975
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
      Left            =   10500
      TabIndex        =   35
      Top             =   975
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
      Left            =   9600
      TabIndex        =   34
      Top             =   975
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
      Left            =   8835
      TabIndex        =   33
      Top             =   975
      Width           =   570
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
      Left            =   8520
      TabIndex        =   32
      Top             =   975
      Width           =   630
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00000080&
      X1              =   8325
      X2              =   15060
      Y1              =   1245
      Y2              =   1245
   End
   Begin VB.Shape Shape3 
      BorderColor     =   &H00000080&
      Height          =   5445
      Left            =   8325
      Top             =   540
      Width           =   6720
   End
   Begin VB.Image Image1 
      Height          =   1545
      Left            =   13080
      Picture         =   "frmCartoonEntry.frx":059B
      Stretch         =   -1  'True
      Top             =   10440
      Visible         =   0   'False
      Width           =   1395
   End
   Begin VB.Shape Shape1 
      FillColor       =   &H00BDDFD1&
      FillStyle       =   0  'Solid
      Height          =   675
      Left            =   10365
      Shape           =   4  'Rounded Rectangle
      Top             =   7890
      Width           =   2670
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
      Left            =   16605
      TabIndex        =   31
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
      Left            =   12495
      TabIndex        =   29
      Top             =   10320
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
      Left            =   12615
      TabIndex        =   28
      Top             =   10680
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
      Left            =   11655
      TabIndex        =   27
      Top             =   10800
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
      Left            =   11535
      TabIndex        =   26
      Top             =   10440
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
      Left            =   11415
      TabIndex        =   25
      Top             =   10680
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
      Left            =   11415
      TabIndex        =   24
      Top             =   10560
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
      Left            =   11655
      TabIndex        =   23
      Top             =   10380
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
      Left            =   11340
      TabIndex        =   22
      Top             =   10800
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
      Left            =   11295
      TabIndex        =   21
      Top             =   10680
      Width           =   1095
   End
   Begin VB.Label Label6 
      BackColor       =   &H00C00000&
      Caption         =   "Cartoon No. Entry"
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
      Left            =   45
      TabIndex        =   20
      Top             =   15
      Width           =   19410
   End
   Begin VB.Shape Shape2 
      BackColor       =   &H00808080&
      BorderColor     =   &H004D177D&
      FillColor       =   &H00EDFDFE&
      FillStyle       =   0  'Solid
      Height          =   10590
      Left            =   225
      Top             =   -1620
      Width           =   15015
   End
End
Attribute VB_Name = "frmCartoonEntry"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Const tblName As String = "tblBags"
Const IdField As String = "InVno"
Const OrderField As String = "InVdt,InVno"
Const MaxNo As Long = 13
Dim varBookMark As Variant
Dim FirstTimeFlag As Boolean
Dim OldgFrmTypeStr As String
Dim FormAction As Integer
Dim frmTypeStr As String
Dim rstbl As Recordset '-- TblAddLess
Dim rstblSub As Recordset '-- TblSITLRDetail
Dim tmpAutoLRNo As String '- Auto LR No
Dim tmpAutoItSrNo As String
Dim tmpAutoLotNo As String
Dim tmpAutoWeight As String

Dim InBg(10) As Integer
Dim InWt(10) As Double

Dim CalBg(10) As Integer
Dim CalWt(10) As Double

Dim TotalInBag As Integer
Dim TotalInWt As Double
Dim TotalLrBag As Integer
Dim TotalLrWt As Double
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
Dim BlDt As Date
Dim RsG As Recordset
Dim Tp As Integer
Dim GodownCd As Long
'--- Save Grid Records
CheckForNumberGrid
   X = "Select PurMillCode,tblPurch.Vdt,PurSubGodown from tblPurch,tblPurchSub where tblPurch.Vno = tblPurchSub.vno and tblPurch.Vtype = tblPurchSub.vtype and " _
       & " tblPurch.Vyear = tblPurchSub.vyear and tblPurch.Vfirm = tblPurchSub.vfirm and tblPurchSub.PurSubItSrNo = 1 and " _
       & " tblPurch.vfirm = '" & gCCode & "' and  tblPurch.VType='" & gfrmBagInwVType & "' and tblPurch.Vno=" & gfrmBagInwVNo & " and tblPurch.VYear=" & gCYear & " "
   i = GProcRstOpen(RsG, X, "R")
   If i > 0 Then
      MillCode = RsG.Fields(0)
      BlDt = RsG.Fields(1)
      GodownCd = RsG.Fields(2)
 '     Tp = RsG.Fields(2)
   End If
'--- Save Grid Records To tblBags
gCn.Execute "delete from tblBags where InVFirm = '" & gCCode & "' and InVType='" & gfrmBagInwVType & "' and InVNo=" & gfrmBagInwVNo & " and INVYear=" & gCYear & ""
X = "select * from tblBags where InVFirm = '" & gCCode & "' and InVType='" & gfrmBagInwVType & "' and InVNo=" & gfrmBagInwVNo & " and INVYear=" & gCYear & ""
GProcRstOpen rstblSub, X, "O"
With msGrid


X = "select InItSrNo,Itname,BagNo,CartoonWt,SlBillNo,InVNo,InVType,INVYear,InVFirm,SlVNo,SlVType,SlVdt,SlVYear,SlVFirm,SlItSrNo from tblBags,tblMastItem where " _
& " tblBags.InItCode=TblMastItem.ItCode and InVFirm = '" & gCCode & "' and InVType='" & gfrmBagInwVType & "' and InVNo=" & gfrmBagInwVNo & " and INVYear=" & gCYear & " order by InVNo,InVType,InItSrNo"



'0 - InItSrNo
'1 - Itname
'2 - BagNo
'3 - CartoonWt
'4 - SlBillNo
'5 - InVNo
'6 - InVType
'7 - INVYear
'8 - InVFirm
'9 - SlVNo
'10 - SlVType
'11 - SlVdt
'12 - SlVYear
'13 - SlVFirm
'14 - SlItSrNo
    
    i = 1
    Do While i < .rows
       If .TextMatrix(i, 3) <> "" Then '- Bag No
            rstblSub.AddNew
            '-- SNo
            rstblSub!InvNo = gfrmBagInwVNo    '--Inw No
            rstblSub!InVtype = gfrmBagInwVType    '--Inw Type
            rstblSub!InvYear = gCYear   '--Inw Year
            rstblSub!InVFirm = gCCode   '--V Firm
            rstblSub!InVdt = BlDt
            rstblSub!InGodown = GodownCd
            rstblSub!InItSrNo = .TextMatrix(i, 0)   '--Item Sr.No. In TblVouSub
            rstblSub!InItCode = GProcGetColumnValue("tblMastItem", "ItName", .TextMatrix(i, 1), "S", "ItCode", "N", " ItMillCode=" & MillCode) '--Item Code
            rstblSub!BagNo = CLng(.TextMatrix(i, 2))   '-- Bag No
            rstblSub!Cartoon = 1
            rstblSub!CartoonWt = CDbl(.TextMatrix(i, 3))   '-- Wt.
            rstblSub!SlBillNo = .TextMatrix(i, 4) & vbNullString    '-- Sale Bill No
            rstblSub!SlVno = CLng(.TextMatrix(i, 9)) '-- SaleVno
            rstblSub!SlVtype = .TextMatrix(i, 10) & vbNullString  '-- SaleVtype
            If IsDate(.TextMatrix(i, 11)) Then
                rstblSub!SlVdt = CDate(.TextMatrix(i, 11)) & vbNullString  '-- SaleVdt
            Else
                rstblSub!SlVdt = BlDt      '-- SaleVdt
            End If
            rstblSub!SlVyear = CLng(.TextMatrix(i, 12)) & vbNullString  '-- SaleVyear
            rstblSub!SlVfirm = .TextMatrix(i, 13) & vbNullString  '-- SaleVfirm
            rstblSub!SlItSrNo = CLng(.TextMatrix(i, 14)) & vbNullString  '-- SaleSr No
            rstblSub!SrNo = i '-- Sr No
            
            
            rstblSub.Update
        End If
        i = i + 1
     Loop
 End With
 rstblSub.Close
End Sub

''    Select Case KeyCode
''        Case 13:
''            If Me.ActiveControl.Name <> "txtGrid" Then SendKeys "{TAB}"
''            'SendKeys "{TAB}"
'''        Case 27:
'''            If (cmdBtn(13).Enabled = False) Then
'''                Unload Me
'''            End If
''    End Select

Private Sub Form_Load()
    Dim X As String
    Dim RsG As Recordset
    Me.Top = 300
    Me.Left = Screen.Width / 2 - Me.Width / 2
    FirstTimeFlag = True
    FormAction = 5
    For i = 1 To 10
        InBg(i) = 0
        InWt(i) = 0
        lblInBag(i).Caption = ""
        lblInWt(i).Caption = ""
    
        lblLRBag(i).Caption = ""
        lblLRWt(i).Caption = ""
        
        lblBalBag(i).Caption = ""
        lblBalWt(i).Caption = ""
    
    Next i
    
    TotalInBag = 0
    TotalInWt = 0

    
    X = "Select PurSubBag,PurSubWt,PurSubItSrNo from tblPurchSub where vfirm = '" & gCCode & "' and VType='" & gfrmBagInwVType & "' and Vno=" & gfrmBagInwVNo & " and VYear=" & gCYear & "  order by PurSubItSrNo "
    If GProcRstOpen(RsG, X, "R") > 0 Then
       With RsG
            If Not .EOF Then
                RsG.MoveFirst
                Do While Not .EOF
                   InBg(RsG.Fields(2)) = RsG.Fields(0)
                   InWt(RsG.Fields(2)) = RsG.Fields(1)
                   lblInBag(RsG.Fields(2)).Caption = CStr(RsG.Fields(0))
                   lblInWt(RsG.Fields(2)).Caption = Format(CStr(RsG.Fields(1)), WtStr)
                   
                   TotalInBag = TotalInBag + RsG.Fields(0)
                   TotalInWt = TotalInWt + RsG.Fields(1)
                   
                   RsG.MoveNext
                Loop
            End If
       End With
    End If
    
    TInBg.Caption = CStr(TotalInBag)
    TInWt.Caption = Format(CStr(TotalInWt), WtStr)
    FillRelatedRecord
    CalAmount
End Sub
'---------------- Grid
Private Sub SetGrid()
Dim X As String
'--- Set Grid
Set rstblSub = New Recordset
X = "select InItSrNo,Itname,BagNo,CartoonWt,SlBillNo,InVNo,InVType,INVYear,InVFirm,SlVNo,SlVType,SlVdt,SlVYear,SlVFirm,SlItSrNo from tblBags,tblMastItem where " _
& " tblBags.InItCode=TblMastItem.ItCode and InVFirm = '" & gCCode & "' and InVType='" & gfrmBagInwVType & "' and InVNo=" & gfrmBagInwVNo & " and INVYear=" & gCYear & " order by InVNo,InVType,SrNo"
GProcRstOpen rstblSub, X, "R"
With msGrid
    .FormatString = ">Sr.|<Count Description                       |>Bag No.    |>Weight    |<Sale Bill No    ||||||||||"
    .ColWidth(5) = 0 '-InVNo
    .ColWidth(6) = 0 '-InVType
    .ColWidth(7) = 0 '-InVyear
    .ColWidth(8) = 0 '-InVfirm
    .ColWidth(9) = 0 '-SlVno
    .ColWidth(10) = 0 '-SlVtype
    .ColWidth(11) = 0 '-SlVdt
    .ColWidth(12) = 0 '-SlVyear
    .ColWidth(13) = 0 '-SlVFirm
    .ColWidth(14) = 0 '-SlItSrNo

'    If FormAction = vbDataActionAddNew Then
        .rows = 2
        GridAddNew
 '   Else
 '       .Rows = 1
   ' End If
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
Dim X As String
Dim InProdTp As Integer

Dim RsG As Recordset
If FormAction <> vbDataActionCancel Then
    CheckForNumberTxtGrid
    DoEvents
    txtGrid = ""
    DoEvents
    If (msGrid.TextMatrix(msGrid.Row, 2)) = "" Then msGrid.TextMatrix(msGrid.Row, 2) = tmpAutoLRNo   ''-Auto Bag No
    If Val(msGrid.TextMatrix(msGrid.Row, 3)) = 0 And tmpAutoWeight <> "" Then msGrid.TextMatrix(msGrid.Row, 3) = tmpAutoWeight  ''-Auto weight
    If msGrid.TextMatrix(msGrid.Row, 0) = "0" And IsNumeric(tmpAutoItSrNo) = True Then msGrid.TextMatrix(msGrid.Row, 0) = tmpAutoItSrNo  ''-Auto It Sr No
    
    If msGrid.text <> "" Then txtGrid.text = msGrid.text
    '--- Lock TextBox
    With msGrid
        Select Case .Col
            Case 0: '---- Sr No
                txtGrid.Locked = False
                DoEvents
            Case 1:
                .Col = .Col + 1
            Case Else
                 If msGrid.TextMatrix(msGrid.Row, 4) = "" Then
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
        If .TextMatrix(.Row, 2) = "" And .rows > 2 Then
            .RemoveItem (.Row)
            txtGrid.text = ""
        End If
        txtGrid.Visible = False
        cmdBtn(10).SetFocus
        Exit Sub
    End If
End With
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
    txtGrid.Move msGrid.CellLeft + msGrid.Left, msGrid.CellTop + msGrid.Top, msGrid.CellWidth, msGrid.CellHeight
End If
txtGrid.Visible = True
txtGrid.Enabled = True
txtGrid.SetFocus
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
    If Val(msGrid.TextMatrix(msGrid.Row, 2)) > 0 Then ''-Auto LR No
        tmpAutoLRNo = CDbl(msGrid.TextMatrix(msGrid.Row, 2)) + 1
    Else
        tmpAutoLRNo = tmpAutoLRNo
    End If
    If CDbl(msGrid.TextMatrix(msGrid.Row, 0)) > 0 Then ''-Auto It Sr No
        tmpAutoItSrNo = msGrid.TextMatrix(msGrid.Row, 0)
    Else
        tmpAutoItSrNo = tmpAutoItSrNo
    End If
End If
DoEvents
'If msGrid.Col = 0 Then
    CalGridAmt
'End If
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
            Case 1 '-- Count
            Case Else
                KeyAscii = GProcValidateKey(rstblSub, msGrid.Col, KeyAscii, txtGrid)
        End Select
   End If
   If KeyAscii = 13 Then CheckForNumberTxtGrid
    '--- Navigate Row ,Col
    If KeyAscii = 13 And .Col < 3 Then
        .Col = .Col + 1
        If .Col = 2 Then
            If IsNumeric(txtGrid) = True Then  '-Auto LR No
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
            If txtGrid <> "" Then '-Auto Lot No
                tmpAutoWeight = txtGrid
            Else
                tmpAutoWeight = ""
            End If
        End If
    ElseIf KeyAscii = 13 And .Col = 3 Then
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
            If (Val(.TextMatrix(.Row, 3)) = 0) And .rows > 2 Then
                .RemoveItem (.Row)
                txtGrid.text = ""
            End If
            txtGrid.Visible = False
            cmdBtn(10).SetFocus
            Exit Sub
        End If
    End With
End With
If txtGrid.Enabled = True And txtGrid.Visible = True Then txtGrid.SetFocus
End Sub
Private Sub txtGrid_KeyDown(KeyCode As Integer, Shift As Integer)
'--- Delete Row On Ctrl+Y

'--- Very IMP After Bill Pl check
If KeyCode = vbKeyY And Shift = 2 And msGrid.rows > 2 Then

'    If (msGrid.TextMatrix(msGrid.Row, 7) <> 0) Then
'        MsgBox "G.P. Avaiable Can not delete..."
'    Else
        msGrid.RemoveItem (msGrid.Row)
        txtGrid.Visible = False
        txtGrid.text = ""
'    End If
'    msGrid = 4

'    ReNumberGridSrno
    CalAmount
    Exit Sub
End If
'--- Show Form
If KeyCode = vbKeyF3 Then
Select Case msGrid.Col
'    Case 1 '-- Item
'        GProcShowForm frmMastItem, frmMain.mnuMstfrm(3), Array(frmMastItem.txttemp(9)), Array(txttemp(5))
'        Exit Sub
'    Case 2 '-- Godown
'        OldgFrmTypeStr = gfrmTypeStr
'        gfrmTypeStr = "G"
'        GProcShowForm frmMastNarration, frmMain.mnuMstfrm(4), "G"
'        gfrmTypeStr = OldgFrmTypeStr
End Select
End If
End Sub
Private Sub CalGridAmt()
Dim X As String
Dim RsG As Recordset
Dim StdPack As Double
'If msGrid.TextMatrix(msGrid.Row, 1) = "" And Val(msGrid.TextMatrix(msGrid.Row, 0)) > 0 Then
If Val(msGrid.TextMatrix(msGrid.Row, 0)) > 0 Then
   X = "Select itname from tblPurchSub,tblMastItem where vfirm = '" & gCCode & "' and VType='" & gfrmBagInwVType & "' and Vno=" & gfrmBagInwVNo & " and VYear=" & gCYear & " and  tblPurchSub.PurSubItCode=TblMastItem.ItCode and " _
       & " PurSubItSrNo = " & CLng(msGrid.TextMatrix(msGrid.Row, 0))
       i = GProcRstOpen(RsG, X, "R")
       If i > 0 Then
          msGrid.TextMatrix(msGrid.Row, 1) = RsG.Fields(0)
      End If
End If
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
                    If RowIndex > 1 And Trim(.TextMatrix(RowIndex, i)) = "" And CDbl(.TextMatrix(RowIndex, 3)) = 0 Then
                        .RemoveItem (RowIndex)
                        GoTo Validate
                    End If
                    If Trim(.TextMatrix(RowIndex, i)) = "" Then
                        MsgBox "Check Count.", vbCritical + vbOKOnly, Me.Caption
                        .Col = i
                        msGrid.SetFocus
                        Exit Function
                    End If
                Case 2 '-- Bagno
                    If Trim(.TextMatrix(RowIndex, i)) = "" Then
                        MsgBox "Check Bag No.", vbCritical + vbOKOnly, Me.Caption
                        .Col = i
                        msGrid.SetFocus
                        Exit Function
                    End If
                Case 3 '-- Wt
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
                MsgBox "Duplicate Bag No." + LRNo, vbCritical + vbOKOnly, Me.Caption
                .Row = i
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

If (TotalInWt) - (TotalLrWt) <> 0 And (TotalInBag) - (TotalLrBag) <> 0 Then
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
''''txt(1) = frmTypeStr '--VType
''''If IsDate(mskVou(12)) = False Then mskVou(12) = "__/__/____" '--BillDt
''''If IsDate(mskOS(10)) = False Then mskOS(10) = "__/__/____" '--Due Dt
''''If IsDate(mskDt(2)) Then '--Vdt
''''    txt(2) = CDate(mskDt(2))
''''Else
''''    txt(2) = ""
''''End If
''''If IsDate(mskDt(9)) Then '-- Book dt
''''    txt(9) = CDate(mskDt(9))
''''Else
''''    txt(9) = ""
''''End If
'''''--- Bill No
''''txt(7) = mskVou(12) '--gpDt
''''txt(36) = txtVou(11)
'''''-- Cr Code
''''txt(3) = GProcGetColumnValue("tblMastAccount", "AcName", txttemp(3), "S", "AcCode", "N")
'''''-- Dr Code
''''txt(4) = GProcGetColumnValue("tblMastAccount", "AcName", txttemp(4), "S", "AcCode", "N")
'''''--Mill
''''txt(5) = GProcGetColumnValue("tblMastAccount", "AcName", txttemp(5), "S", "AcCode", "N")
'''''-- TaxType
''''txt(12) = GProcGetColumnValue("tblMastNarration", "Narration", Cbo(12), "S", "NarrCode", "N")
'''''If txttemp(12) = "" Then txttemp(12) = "-"
'''''txt(12) = GProcGetColumnValue("tblMastNarration", "Narration", txttemp(12), "S", "NarrCode", "N")
''''txt(23) = gCYear
'''' '--Total Amt
'''''If CDbl(txt(21)) = 0 Then txt(21) = txt(11)
'''''-- JV
'''''gCn.Execute "delete from tblVoucher where vtype='JV' and vno=" & txt(20) & " and " _
'''''& " VYear=" & gCYear
'''''If CDbl(txt(21)) = CDbl(txt(11)) Then
'''''    txt(20) = 0
'''''End If
''''ShowAccountBalance
End Sub
Private Sub FillTempFromTxt()
''''Dim X As String
''''FillCombo
''''If IsDate(txt(2)) Then '--VDt
''''    mskDt(2) = txt(2)
''''Else
''''    mskDt(2) = "__/__/____"
''''End If
''''If IsDate(txt(9)) Then '--Book Dt
''''    mskDt(9) = txt(9)
''''Else
''''    mskDt(9) = "__/__/____"
''''End If
'''''-- Cr Code
''''txttemp(3) = GProcGetColumnValue("tblMastAccount", "AcCode", txt(3), "N", "AcName", "S")
'''''-- Dr Code
''''txttemp(4) = GProcGetColumnValue("tblMastAccount", "AcCode", txt(4), "N", "AcName", "S")
'''''--Mill
''''txttemp(5) = GProcGetColumnValue("tblMastAccount", "AcCode", txt(5), "N", "AcName", "S")
'''''-- TaxType
''''Cbo(12) = GProcGetColumnValue("tblMastNarration", "NarrCode", txt(12), "N", "Narration", "S")
'''''txttemp(12) = GProcGetColumnValue("tblMastNarration", "NarrCode", txt(12), "N", "Narration", "S")
'-- Fill Related Record
FillRelatedRecord
'-- Expenses
'txtExpensesAmt = CDbl(txt(21)) - CDbl(txt(11))
''''ShowAccountBalance
''''MillCtrlEd
'''''--Total & Sub Diff. Amt   (Total - Sub) amt
''''If IsNumeric(txttmpAmt) = False Then txttmpAmt = 0
''''txttmpAmt = txt(11)
''''txttmpSubTotDiff = Format(CDbl(txttmpAmt) - CDbl(txt(10)), FStr)
End Sub
Private Sub FillRelatedRecord()
Dim RowIndex As Long
Dim ColIndex As Long
Dim TotBag As Long
Dim TotWt As Double
tmpAutoLRNo = "" ''-Auto LR No
tmpAutoItSrNo = ""
tmpAutoLotNo = ""
tmpAutoWeight = ""
'tmpAutoGrade = ""
'tmpAutoChases = ""
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
                If ColIndex = 2 Then TotBag = TotBag + 1 '---CLng(msGrid.TextMatrix(RowIndex, 4)) '-- Tot Bag
                If ColIndex = 3 Then TotWt = TotWt + CDbl(msGrid.TextMatrix(RowIndex, 3)) '-- Tot Wt
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
TotalLrBag = 0
TotalLrWt = 0
For i = 1 To 10
    CalBg(i) = 0
    CalWt(i) = 0
    lblLRBag(i).Caption = ""
    lblLRWt(i).Caption = ""
    lblBalBag(i).Caption = ""
    lblBalWt(i).Caption = ""
Next i

RowIndex = 1
With msGrid
    Do While RowIndex < .rows
        '-- bag
        If Val(.TextMatrix(RowIndex, 3)) > 0 Then
           If .TextMatrix(RowIndex, 0) <> "" Then
              CalBg(.TextMatrix(RowIndex, 0)) = CalBg(.TextMatrix(RowIndex, 0)) + 1
              CalWt(.TextMatrix(RowIndex, 0)) = Format(CalWt(.TextMatrix(RowIndex, 0)) + Val(.TextMatrix(RowIndex, 3)), WtStr)
           End If
        End If
        RowIndex = RowIndex + 1
    Loop
End With

For i = 1 To 10
    If CalBg(i) > 0 Then
       lblLRBag(i).Caption = CStr(CalBg(i))
       lblLRWt(i).Caption = Format(CStr(CalWt(i)), WtStr)
       lblBalBag(i).Caption = CStr(InBg(i) - CalBg(i))
       lblBalWt(i).Caption = Format(CStr(InWt(i) - CalWt(i)), WtStr)
       TotalLrBag = TotalLrBag + CalBg(i)
       TotalLrWt = TotalLrWt + CalWt(i)
    End If
Next i
TLrBg.Caption = CStr(TotalLrBag)
TLrWt.Caption = Format(CStr(TotalLrWt), WtStr)

TBlBg.Caption = CStr(TotalInBag - TotalLrBag)
TBlWt.Caption = Format(CStr(TotalInWt - TotalLrWt), WtStr)



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
'If GProcISAudited(txt(0), txt(1), txt(23)) = True Then
'    MsgBox "Record is Audited.You can not Modify/Delete this Record.", vbCritical + vbOKOnly, App.Title
'    IsRelatedRecord = True
'    Exit Function
'End If
'-- Purchase Return
'X = "Select Vtype,Vdt,Vno from tblVouSub where Vtype in ('VY','VI') and VSubPItCtrlNo in" _
'& " (select VSubItCtrlNo from tblVouSub where vtype='" & frmTypeStr & "' and vno=" & txt(0) & " and " _
'& " VYear=" & gCYear & ")"
'GProcRstOpen RsR, X, "R"
'If RsR.EOF And RsR.BOF Then
'Else
'    MsgBox "Related Purchase Return Record " & RsR.Fields("Vno") & " is exists.You can not Modify/Delete this Record.", vbCritical + vbOKOnly, App.Title
'    IsRelatedRecord = True
'    Exit Function
'End If
RsR.Close
End Function
