VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form frmReceiptGST 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00C2E3FC&
   BorderStyle     =   0  'None
   Caption         =   "Receipt Entry"
   ClientHeight    =   7965
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   15870
   Icon            =   "frmReceiptGST.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7965
   ScaleWidth      =   15870
   ShowInTaskbar   =   0   'False
   Begin VB.CheckBox ChkIsTcs 
      BackColor       =   &H00FFC0FF&
      Caption         =   "Is TCS Entry ?"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   4725
      TabIndex        =   88
      Top             =   7470
      Width           =   2400
   End
   Begin VB.TextBox txt 
      Height          =   375
      Index           =   32
      Left            =   18315
      TabIndex        =   84
      Text            =   "32"
      Top             =   4410
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.TextBox txt 
      Height          =   375
      Index           =   31
      Left            =   18315
      TabIndex        =   83
      Text            =   "31"
      Top             =   3960
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.TextBox txt 
      Height          =   375
      Index           =   30
      Left            =   18225
      TabIndex        =   82
      Text            =   "30"
      Top             =   3510
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.TextBox txt 
      Height          =   375
      Index           =   29
      Left            =   18270
      TabIndex        =   81
      Text            =   "29"
      Top             =   3060
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.TextBox txt 
      Height          =   375
      Index           =   28
      Left            =   18270
      TabIndex        =   80
      Text            =   "28"
      Top             =   2655
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.TextBox txt 
      Height          =   375
      Index           =   27
      Left            =   18180
      TabIndex        =   79
      Text            =   "27"
      Top             =   2250
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.TextBox txt 
      Height          =   375
      Index           =   26
      Left            =   18225
      TabIndex        =   78
      Text            =   "26"
      Top             =   1800
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.TextBox txt 
      Height          =   375
      Index           =   25
      Left            =   18225
      TabIndex        =   77
      Text            =   "25"
      Top             =   1395
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.TextBox txt 
      Height          =   375
      Index           =   24
      Left            =   18135
      TabIndex        =   76
      Text            =   "24"
      Top             =   990
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.TextBox txt 
      Height          =   375
      Index           =   23
      Left            =   18180
      TabIndex        =   75
      Text            =   "23"
      Top             =   540
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.TextBox txt 
      Height          =   375
      Index           =   22
      Left            =   17235
      TabIndex        =   74
      Text            =   "22"
      Top             =   90
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.TextBox txt 
      Height          =   345
      Index           =   21
      Left            =   720
      TabIndex        =   73
      Text            =   "21"
      Top             =   8160
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.TextBox txt 
      Height          =   345
      Index           =   20
      Left            =   16425
      TabIndex        =   72
      Text            =   "20"
      Top             =   6165
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFC0C0&
      Caption         =   "Voucher Printing"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2835
      Left            =   3915
      TabIndex        =   63
      Top             =   4140
      Visible         =   0   'False
      Width           =   5175
      Begin VB.CommandButton cmdPrint 
         Caption         =   "S&hort Note"
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
         Left            =   1950
         TabIndex        =   71
         Top             =   2160
         Width           =   1335
      End
      Begin VB.TextBox txtno 
         Appearance      =   0  'Flat
         Height          =   360
         Index           =   0
         Left            =   1950
         MaxLength       =   7
         TabIndex        =   68
         Top             =   600
         Width           =   1575
      End
      Begin VB.CommandButton cmdPrint 
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
         Left            =   480
         TabIndex        =   67
         Top             =   1560
         Width           =   1335
      End
      Begin VB.TextBox txtno 
         Appearance      =   0  'Flat
         Height          =   360
         Index           =   1
         Left            =   1950
         MaxLength       =   7
         TabIndex        =   66
         Top             =   1080
         Width           =   1575
      End
      Begin VB.CommandButton cmdPrint 
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
         Left            =   1935
         TabIndex        =   65
         Top             =   1560
         Width           =   1335
      End
      Begin VB.CommandButton cmdPrint 
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
         Left            =   3480
         TabIndex        =   64
         Top             =   1560
         Width           =   1335
      End
      Begin VB.Label Label31 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "From VNo :"
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
         Left            =   600
         TabIndex        =   70
         Top             =   600
         Width           =   1065
      End
      Begin VB.Label Label18 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "To VNo     :"
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
         Left            =   600
         TabIndex        =   69
         Top             =   1080
         Width           =   1050
      End
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
      Left            =   13065
      Style           =   1  'Graphical
      TabIndex        =   62
      ToolTipText     =   "Exit"
      Top             =   3075
      Width           =   975
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
      Left            =   13065
      Style           =   1  'Graphical
      TabIndex        =   61
      ToolTipText     =   "To Cancel Record"
      Top             =   2595
      Width           =   975
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
      Left            =   11940
      Style           =   1  'Graphical
      TabIndex        =   60
      ToolTipText     =   "To Save Record"
      Top             =   3090
      Width           =   975
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
      Left            =   13065
      Style           =   1  'Graphical
      TabIndex        =   59
      Top             =   1305
      Width           =   975
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
      Left            =   11940
      Style           =   1  'Graphical
      TabIndex        =   58
      ToolTipText     =   "To Delete Record"
      Top             =   2550
      Width           =   975
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
      Left            =   11940
      Style           =   1  'Graphical
      TabIndex        =   57
      ToolTipText     =   "To Modify Record"
      Top             =   2100
      Width           =   975
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
      Left            =   11940
      Style           =   1  'Graphical
      TabIndex        =   56
      ToolTipText     =   "To Add New Record"
      Top             =   1650
      Width           =   975
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
      Left            =   13065
      Style           =   1  'Graphical
      TabIndex        =   55
      ToolTipText     =   "To Find Record"
      Top             =   2115
      Width           =   975
   End
   Begin VB.CommandButton cmdIntDbNt 
      Caption         =   "&Interest Debit Note"
      Height          =   525
      Left            =   13065
      TabIndex        =   54
      Top             =   585
      Width           =   975
   End
   Begin VB.ListBox List1 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   870
      Left            =   1485
      TabIndex        =   9
      Top             =   2835
      Visible         =   0   'False
      Width           =   1830
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
      Index           =   11
      Left            =   2145
      TabIndex        =   8
      Text            =   "11"
      Top             =   2850
      Width           =   1755
   End
   Begin VB.TextBox txttemp1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Index           =   8
      Left            =   1515
      TabIndex        =   7
      Text            =   "8"
      Top             =   2850
      Width           =   555
   End
   Begin VB.TextBox txtCrBalance 
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
      Height          =   375
      Left            =   8100
      TabIndex        =   50
      Text            =   "txtCrBalance"
      Top             =   1230
      Width           =   1785
   End
   Begin VB.TextBox txtDrBalance 
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
      Height          =   375
      Left            =   8100
      TabIndex        =   49
      Text            =   "txtDrBalance"
      Top             =   1695
      Width           =   1785
   End
   Begin VB.CommandButton cmdLedger 
      Caption         =   "Ledger"
      Height          =   375
      Index           =   0
      Left            =   10065
      TabIndex        =   48
      Top             =   1230
      Width           =   1200
   End
   Begin VB.CommandButton cmdLedger 
      Caption         =   "Ledger"
      Height          =   375
      Index           =   1
      Left            =   10080
      TabIndex        =   47
      Top             =   1695
      Width           =   1185
   End
   Begin VB.TextBox txtGrid 
      BackColor       =   &H8000000A&
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
      Left            =   480
      TabIndex        =   14
      Top             =   4815
      Visible         =   0   'False
      Width           =   1065
   End
   Begin VB.TextBox txt 
      Height          =   325
      Index           =   18
      Left            =   17085
      MaxLength       =   40
      TabIndex        =   46
      Text            =   "18"
      Top             =   2025
      Visible         =   0   'False
      Width           =   270
   End
   Begin VB.TextBox txttemp1 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   375
      Index           =   7
      Left            =   5235
      MaxLength       =   40
      TabIndex        =   11
      Text            =   "txttemp1(7)"
      ToolTipText     =   "Press F3 for Creation & Press F2 for Selection"
      Top             =   2400
      Width           =   6060
   End
   Begin VB.TextBox txt 
      Height          =   345
      Index           =   17
      Left            =   16920
      TabIndex        =   45
      Text            =   "17"
      Top             =   6270
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.TextBox txt 
      Height          =   345
      Index           =   16
      Left            =   16920
      TabIndex        =   44
      Text            =   "16"
      Top             =   5790
      Visible         =   0   'False
      Width           =   495
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
      ForeColor       =   &H00000080&
      Height          =   375
      Index           =   8
      Left            =   5235
      MaxLength       =   40
      TabIndex        =   12
      Text            =   "8"
      ToolTipText     =   "Press F3 for Creation & Press F2 for Selection"
      Top             =   2850
      Width           =   6060
   End
   Begin VB.TextBox txt 
      Height          =   345
      Index           =   15
      Left            =   16920
      TabIndex        =   43
      Text            =   "15"
      Top             =   5310
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.TextBox txt 
      Height          =   375
      Index           =   12
      Left            =   16590
      TabIndex        =   41
      Text            =   "12"
      Top             =   3420
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.TextBox txt 
      Height          =   325
      Index           =   10
      Left            =   16980
      TabIndex        =   40
      Text            =   "10"
      Top             =   3120
      Visible         =   0   'False
      Width           =   360
   End
   Begin VB.TextBox txt 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFFFF&
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Index           =   0
      Left            =   1515
      TabIndex        =   1
      Text            =   "0"
      Top             =   585
      Width           =   1680
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
      Index           =   5
      Left            =   1530
      MaxLength       =   40
      TabIndex        =   4
      Text            =   "txttemp(5)"
      ToolTipText     =   "Press F3 for Creation"
      Top             =   1230
      Width           =   6435
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
      Index           =   4
      Left            =   1515
      MaxLength       =   40
      TabIndex        =   5
      Text            =   "txttemp(4)"
      ToolTipText     =   "Press F3 for Creation"
      Top             =   1695
      Width           =   6435
   End
   Begin VB.TextBox txt 
      Height          =   375
      Index           =   14
      Left            =   16920
      TabIndex        =   33
      Text            =   "14"
      Top             =   3990
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.TextBox txt 
      Alignment       =   1  'Right Justify
      Height          =   360
      Index           =   13
      Left            =   17025
      TabIndex        =   32
      Text            =   "13"
      Top             =   3585
      Visible         =   0   'False
      Width           =   345
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
      Left            =   11940
      Style           =   1  'Graphical
      TabIndex        =   23
      ToolTipText     =   "Previous"
      Top             =   585
      Width           =   450
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
      Left            =   12465
      Style           =   1  'Graphical
      TabIndex        =   24
      ToolTipText     =   "Next"
      Top             =   585
      Width           =   450
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
      Left            =   11940
      Style           =   1  'Graphical
      TabIndex        =   25
      ToolTipText     =   "First"
      Top             =   1035
      Width           =   450
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
      Left            =   12465
      Style           =   1  'Graphical
      TabIndex        =   26
      ToolTipText     =   "Last"
      Top             =   1035
      Width           =   450
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
      Left            =   16110
      Style           =   1  'Graphical
      TabIndex        =   27
      Top             =   6930
      Visible         =   0   'False
      Width           =   975
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
      Left            =   16800
      Style           =   1  'Graphical
      TabIndex        =   34
      Top             =   4440
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
      Left            =   16860
      Style           =   1  'Graphical
      TabIndex        =   35
      Top             =   4875
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.TextBox txt 
      Height          =   345
      Index           =   3
      Left            =   16620
      TabIndex        =   38
      Text            =   "3"
      Top             =   1650
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.TextBox txt 
      Alignment       =   1  'Right Justify
      Height          =   360
      Index           =   5
      Left            =   16620
      TabIndex        =   31
      Text            =   "5"
      Top             =   2370
      Visible         =   0   'False
      Width           =   345
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
      Index           =   6
      Left            =   1515
      TabIndex        =   6
      Text            =   "6"
      Top             =   2400
      Width           =   2400
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
      ForeColor       =   &H00000080&
      Height          =   375
      Index           =   9
      Left            =   5235
      MaxLength       =   40
      TabIndex        =   13
      Text            =   "9"
      ToolTipText     =   "Press F3 for Creation & Press F2 for Selection"
      Top             =   3300
      Width           =   6060
   End
   Begin VB.TextBox txt 
      Height          =   325
      Index           =   7
      Left            =   16710
      TabIndex        =   29
      Text            =   "7"
      Top             =   2940
      Visible         =   0   'False
      Width           =   240
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
      Height          =   345
      Index           =   19
      Left            =   16620
      TabIndex        =   28
      Text            =   "19"
      Top             =   2910
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.TextBox txt 
      Alignment       =   1  'Right Justify
      Height          =   360
      Index           =   4
      Left            =   17160
      TabIndex        =   30
      Text            =   "4"
      Top             =   1590
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.TextBox txt 
      Height          =   345
      Index           =   2
      Left            =   16620
      TabIndex        =   37
      Text            =   "2"
      Top             =   1290
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
      Height          =   345
      Index           =   1
      Left            =   16800
      TabIndex        =   36
      Text            =   "1"
      Top             =   870
      Visible         =   0   'False
      Width           =   555
   End
   Begin MSMask.MaskEdBox mskDt 
      Height          =   360
      Index           =   2
      Left            =   4890
      TabIndex        =   2
      Top             =   585
      Width           =   1620
      _ExtentX        =   2858
      _ExtentY        =   635
      _Version        =   393216
      ClipMode        =   1
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
      Height          =   360
      Index           =   12
      Left            =   1515
      TabIndex        =   10
      Top             =   3285
      Visible         =   0   'False
      Width           =   1590
      _ExtentX        =   2805
      _ExtentY        =   635
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
      Height          =   360
      Index           =   13
      Left            =   8640
      TabIndex        =   3
      Top             =   585
      Width           =   1605
      _ExtentX        =   2831
      _ExtentY        =   635
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
   Begin MSFlexGridLib.MSFlexGrid msGrid 
      Height          =   3450
      Left            =   120
      TabIndex        =   15
      Top             =   3915
      Width           =   15615
      _ExtentX        =   27543
      _ExtentY        =   6085
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
   Begin Crystal.CrystalReport cryReport 
      Bindings        =   "frmReceiptGST.frx":058A
      Left            =   165
      Top             =   660
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
      Caption         =   "Press F4 To Show All Bills"
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
      Height          =   255
      Index           =   6
      Left            =   180
      TabIndex        =   87
      Top             =   7470
      Width           =   4680
   End
   Begin VB.Label lblBalanceAmt 
      BackStyle       =   0  'Transparent
      Caption         =   "UnAdjusted Amount :"
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
      Height          =   255
      Left            =   11970
      TabIndex        =   86
      Top             =   7470
      Width           =   2115
   End
   Begin VB.Label lbl1 
      BackStyle       =   0  'Transparent
      Caption         =   "UnAdjusted Amount :"
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
      Height          =   255
      Index           =   5
      Left            =   9900
      TabIndex        =   85
      Top             =   7470
      Width           =   2115
   End
   Begin VB.Label lbl1 
      BackStyle       =   0  'Transparent
      Caption         =   "&Party        :"
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
      Index           =   3
      Left            =   360
      TabIndex        =   53
      Top             =   1695
      Width           =   1215
   End
   Begin VB.Label lblBalance 
      BackColor       =   &H00ECFFFE&
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
      Left            =   8535
      TabIndex        =   52
      Top             =   3075
      Width           =   1095
   End
   Begin VB.Label lbl1 
      BackStyle       =   0  'Transparent
      Caption         =   "Ban&k       :"
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
      Height          =   240
      Index           =   2
      Left            =   4185
      TabIndex        =   51
      Top             =   2400
      Width           =   1455
   End
   Begin VB.Shape Shape6 
      BorderColor     =   &H00C00000&
      BorderWidth     =   5
      FillColor       =   &H00FF0000&
      Height          =   7935
      Left            =   0
      Top             =   0
      Width           =   15840
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
      Left            =   10560
      TabIndex        =   42
      Top             =   0
      Width           =   3180
   End
   Begin VB.Label lbl1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Entry &No.  :"
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
      Height          =   240
      Index           =   0
      Left            =   360
      TabIndex        =   0
      Top             =   585
      Width           =   1035
   End
   Begin VB.Label lbl1 
      BackStyle       =   0  'Transparent
      Caption         =   "&Rcon.Date :"
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
      Height          =   360
      Index           =   10
      Left            =   7320
      TabIndex        =   17
      Top             =   585
      Width           =   1335
   End
   Begin VB.Label lbl1 
      BackStyle       =   0  'Transparent
      Caption         =   "Am&ount    :"
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
      Height          =   255
      Index           =   9
      Left            =   360
      TabIndex        =   19
      Top             =   2400
      Width           =   1215
   End
   Begin VB.Label lbl1 
      BackStyle       =   0  'Transparent
      Caption         =   "Date         :"
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
      Height          =   360
      Index           =   8
      Left            =   360
      TabIndex        =   21
      Top             =   3285
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.Label lbl1 
      BackStyle       =   0  'Transparent
      Caption         =   "&Bank        :"
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
      Left            =   360
      TabIndex        =   18
      Top             =   1230
      Width           =   1215
   End
   Begin VB.Label lbl1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Da&te   :"
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
      Height          =   240
      Index           =   1
      Left            =   3960
      TabIndex        =   16
      Top             =   585
      Width           =   780
   End
   Begin VB.Shape Shape1 
      FillColor       =   &H00BDDFD1&
      FillStyle       =   0  'Solid
      Height          =   3360
      Left            =   11550
      Shape           =   4  'Rounded Rectangle
      Top             =   450
      Width           =   2745
   End
   Begin VB.Label Label6 
      BackColor       =   &H00C00000&
      Caption         =   "   Bank Receipts"
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
      Left            =   45
      TabIndex        =   39
      Top             =   30
      Width           =   14385
   End
   Begin VB.Label lbl1 
      BackStyle       =   0  'Transparent
      Caption         =   "Narration :"
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
      Height          =   240
      Index           =   11
      Left            =   4185
      TabIndex        =   22
      Top             =   2850
      Width           =   1455
   End
   Begin VB.Label lbl1 
      BackStyle       =   0  'Transparent
      Caption         =   "R&ef No.     :"
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
      Height          =   255
      Index           =   13
      Left            =   360
      TabIndex        =   20
      Top             =   2850
      Width           =   1095
   End
   Begin VB.Shape Shape2 
      BackColor       =   &H00808080&
      BorderColor     =   &H00000080&
      FillColor       =   &H00F5F5DE&
      FillStyle       =   0  'Solid
      Height          =   585
      Left            =   240
      Top             =   465
      Width           =   11190
   End
   Begin VB.Shape Shape3 
      BackColor       =   &H00ECFFFE&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00000080&
      FillColor       =   &H00ECFFFF&
      Height          =   1065
      Left            =   240
      Top             =   1110
      Width           =   11190
   End
   Begin VB.Shape Shape5 
      BorderColor     =   &H00000080&
      FillColor       =   &H00ECFFFE&
      FillStyle       =   0  'Solid
      Height          =   1575
      Left            =   240
      Top             =   2235
      Width           =   3810
   End
   Begin VB.Shape Shape4 
      BorderColor     =   &H00000080&
      FillColor       =   &H00F5F5DE&
      FillStyle       =   0  'Solid
      Height          =   1575
      Left            =   4125
      Top             =   2235
      Width           =   7305
   End
End
Attribute VB_Name = "frmReceiptGST"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Const tblName As String = "tblVoucher"
Const IdField As String = "Vno"
Const MaxNo As Long = 31
Dim CriteriaStr As String
Dim rstbl As Recordset, ShowFrmFlag As Boolean
Dim varBookMark As Variant
Dim FirstTimeFlag As Boolean, mLastEntryBank As String, mLastEntryDate As Date
Dim FormAction As Integer
Dim frmTypeStr As String
Dim i As Long
Dim wCotCGStRt As Double
Dim wCotSGStRt As Double
Dim wCotIGStRt As Double
Dim wPolCGStRt As Double
Dim wPolSGStRt As Double
Dim wPolIGStRt As Double
'-- RecVsSale
Dim rsRecVsSale As Recordset
Dim OldgFrmTypeStr As String
Const OrderField As String = "Vdt,Vno,VCtrNo"
Private Sub cmdBtn_Click(Index As Integer)
Dim FindQStr As String
Dim X As String
'On Error GoTo ErrorRoutine
Select Case Index
Case 0 'Previous
    If GProcRstOpen(rstbl, "Select * from " & tblName & " where vno = (Select max(Vno) from " & tblName & " where Vno <> 0  and Vno < " & txt(0) & " and  " & CriteriaStr & "  ) and  " & CriteriaStr & " order by " & OrderField & " ", "O") > 0 Then
       ReadFields
    Else
      MsgBox "This is First Record ", vbOKOnly + vbInformation
    End If
'    rstbl.MovePrevious
'    If rstbl.BOF = True Then
'        rstbl.MoveFirst
'        MsgBox ("This is First Record"), vbInformation + vbOKOnly, "First Record"
'    End If
'    ReadFields
Case 1 'Next
    If GProcRstOpen(rstbl, "Select * from " & tblName & " where Vno = (Select min(Vno) from " & tblName & " where Vno <> 0  and Vno > " & txt(0) & "  and " & CriteriaStr & ") and  " & CriteriaStr & " order by " & OrderField & "  ", "O") > 0 Then
       ReadFields
    Else
       MsgBox "This is Last Record ", vbOKOnly + vbInformation
    End If
'    rstbl.MoveNext
'    If rstbl.EOF = True Then
'        rstbl.MoveLast
'        MsgBox ("This is Last Record"), vbInformation + vbOKOnly, "Last Record"
'    End If
'    ReadFields
Case 2 'First
    If GProcRstOpen(rstbl, "Select * from " & tblName & " where Vno = (Select min(Vno) from " & tblName & " where Vno <> 0 and " & CriteriaStr & ") and  " & CriteriaStr & " order by " & OrderField & "  ", "O") > 0 Then ReadFields
'    rstbl.MoveFirst
'    ReadFields
Case 3 'Last
    If GProcRstOpen(rstbl, "Select * from " & tblName & " where " & CriteriaStr & "  and  " & IdField & " = (Select max(" & IdField & ") from " & tblName & " where vdt = (SELECT MAX(vdt) from " & tblName & " where " & CriteriaStr & " )  and  " & IdField & " <> 0 and  " & CriteriaStr & " ) order by " & OrderField & " ", "R") > 0 Then ReadFields Else GProcClearForm Me, rstbl, MaxNo
'    rstbl.MoveLast
'    ReadFields
Case 4 'Find
    If gBackEndDB = gBackEndAccess Then
        X = " vNar1 & ' ' & VNar2 from " _
    & " tblvoucher,TblMastAccount,TblMastAccount as TblAc1 "
    Else '-- Oracle
        X = " vNar1 || ' ' || VNar2 from " _
    & " tblvoucher,TblMastAccount,TblMastAccount TblAc1 "
    End If
    FindQStr = "select Vno,Vdt,TblMastAccount.AcName,TblAc1.AcName,Vamt," & X & " where Vtype='" & frmTypeStr & "' and " _
    & " Vyear=" & gCYear & " and VCtrNo=1 and VAcCode=TblMastAccount.AcCode and VAcOCode=TblAc1.AcCode"
    gClsSearch.SearchMultiField FindQStr, "", Array("Vno", "Vdt", "Account", "Credit A/c", "Amount", "Narration"), Array(700, 1200, 3500, 2500, 1000, 2000), "", " ", "vno,vdt", 0, 0, False, , True
    If gClsSearch.SearchMultiRetCol(0) <> "" Then
        If GProcRstOpen(rstbl, "Select * from " & tblName & " where Vno = (Select max(" & gClsSearch.SearchMultiRetCol(0) & ") from " & tblName & " where Vno <> 0 and " & CriteriaStr & " )and  " & CriteriaStr & " order by " & OrderField & "  ", "O") > 0 Then ReadFields
        gClsSearch.SearchMultiRetCol(0) = ""
    
'        rstbl.MoveFirst
'        rstbl.Find "Vno = " & gClsSearch.SearchMultiRetCol(0)
'        gClsSearch.SearchMultiRetCol(0) = ""
'        ReadFields
    End If
Case 5 'List
    'GenMastList ("LAC")
Case 6 'Add
    FormAction = vbDataActionAddNew
    lblBalanceAmt.Caption = ""
    Call GProcActivateControls(Me, True)
    Call GProcSetButtons(Me, vbDataActionAddNew)
    Call GProcClearForm(Me, rstbl, MaxNo, True)
    txt(0).text = GProcGenerateId(tblName, "VNO", CriteriaStr, gCn)
    mskDt(2) = GProcGenerateVDt(tblName, "VDt", CriteriaStr)
    If (mLastEntryBank <> "") Then
       txtTemp(5).text = mLastEntryBank
       mskDt(2) = mLastEntryDate
    End If
    If FirstTimeFlag = True Then
        FirstTimeFlag = False
    Else
        txt(0).SetFocus
    End If
    FillCombo
    SetGrid
Case 7 'Modify
    'If mURecModi = True Then
    '    SearchHelp
    '    DoEvents
    If IsRelatedRecord = False Then
        FormAction = vbDataActionUpdate
        Call GProcActivateControls(Me, True)
        Call GProcSetButtons(Me, vbDataActionUpdate)
        '--Update Outstanding Balance
        GProcUpdateOutstandingBal CLng(txt(0)), UCase(txt(3)), gCYear, True
        txt(0).Enabled = False
'        txt(1).SetFocus
    End If
Case 8 'Delete
   ' If mURecDel = True Then
     If IsRelatedRecord = False Then
        If MsgBox("Do You Really Delete This Record?", vbQuestion + vbYesNo + vbDefaultButton1, "Confirm Delete") = vbYes Then
           gCn.Execute "Delete from tblvoucher where " & CriteriaStr & " and vno = " & txt(0)
           '--Update Outstanding Balance
           GProcUpdateOutstandingBal CLng(txt(0)), UCase(txt(3)), gCYear, False, True
            '-- Delete Related Record
            DeleteRelatedRecord
            If rstbl.RecordCount > 1 Then
                rstbl.Requery
                cmdBtn_Click (1)
            ElseIf rstbl.RecordCount = 1 Then
                Call GProcActivateControls(Me, False)
                Call GProcSetButtons(Me, vbDataActionCancel)
                Form_Load
            End If
        End If
        DoEvents
     End If
   ' End If
Case 9 'Print
    Frame1.Visible = True
    DoEvents
    txtno(0).Enabled = True
    txtno(1).Enabled = True
    txtno(0) = txt(0)
    txtno(1) = txt(0)
    txtno(0).SetFocus
Case 10 'Exit
    DoEvents
    Unload Me
Case 13 'Save
    If ValidateData = True Then
       If Trim(txtGrid) <> "" Then
            msGrid.text = txtGrid
            DoEvents
            txtGrid = ""
        End If
        txtGrid.Visible = False
        FillTxtFromTemp
   '     GProcSaveRecord Me, rstbl, gFormAction
        gCn.Execute "Delete from tblvoucher where " & CriteriaStr & " and vno = " & txt(0)
        GProcCreateVoucher FormAction, txt(0), 1, txt(2), frmTypeStr, txt(4), txt(5), Val(txt(6)), txt(11), txt(12), txt(13), txt(7), txt(8), txt(9), txt(10), False, txt(20), 0, GSTEnt:=1
        GProcCreateVoucher FormAction, txt(0), 2, txt(2), frmTypeStr, txt(5), txt(4), Val(txt(6)) * (-1), txt(11), txt(12), txt(13), txt(7), txt(8), txt(9), txt(10), False, txt(20), 0, GSTEnt:=1
        '--- Related Rcord Rec Vs Sale
        SaveRelatedRecord
        '--Update Outstanding Balance
        GProcUpdateOutstandingBal CLng(txt(0)), UCase(txt(3)), gCYear
        rstbl.Requery
        txt(32) = 1
        Call GProcActivateControls(Me, False)
        Call GProcSetButtons(Me, True)
        If FormAction = vbDataActionAddNew Then
            If cmdBtn(6).Visible = True Then cmdBtn(6).SetFocus
        Else
            If cmdBtn(7).Visible = True Then cmdBtn(7).SetFocus
        End If
        FormAction = vbDataActionClose
        msGrid.Refresh
        msGrid.Col = 0
        msGrid.Row = 1
        txtGrid.Move msGrid.CellLeft + msGrid.Left, msGrid.CellTop + msGrid.Top, msGrid.CellWidth, msGrid.CellHeight
        lblBalanceAmt.Caption = ""
    End If
Case 14 'Cancel
    If GProcRstOpen(rstbl, "Select * from " & tblName & " where " & CriteriaStr & "  and  " & IdField & " = (Select max(" & IdField & ") from " & tblName & " where vdt = (SELECT MAX(vdt) from " & tblName & " where " & CriteriaStr & " )  and  " & IdField & " <> 0 and  " & CriteriaStr & " ) order by " & OrderField & " ", "R") > 0 Then ReadFields Else GProcClearForm Me, rstbl, MaxNo
    txtGrid.Visible = False
    If rstbl.RecordCount <= 0 Then
        Call GProcClearForm(Me, rstbl, MaxNo, True)
        Call GProcActivateControls(Me, False)
        Call GProcSetButtons(Me, vbDataActionCancel)
    Else
        If FormAction = vbDataActionAddNew Then
'            rstbl.MoveLast
        Else
             '--Update Outstanding Balance
            GProcUpdateOutstandingBal CLng(txt(0)), UCase(txt(3)), gCYear
            rstbl.CancelUpdate
        End If
        Call GProcClearForm(Me, rstbl, MaxNo, True)
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
    msGrid.Col = 0
    If msGrid.Row >= 1 Then msGrid.Row = 1
    txtGrid.Move msGrid.CellLeft + msGrid.Left, msGrid.CellTop + msGrid.Top, msGrid.CellWidth, msGrid.CellHeight
    lblBalanceAmt.Caption = ""
End Select
Exit Sub
ErrorRoutine:
    rstbl.CancelUpdate
    GProcErrorHandler
End Sub
'------  Print Voucher
Private Sub txtno_GotFocus(Index As Integer)
GProcSelectBox Me.ActiveControl
End Sub
Private Sub txtno_KeyPress(Index As Integer, KeyAscii As Integer)
KeyAscii = GProcNumberOnly(KeyAscii)
End Sub
Private Sub cmdPrint_Click(Index As Integer)
Dim Rs1 As Recordset
Dim X As String
Dim Vno As Long
Select Case Index
Case 0, 1:
    '------ Validate
    If IsNumeric(txtno(0)) = False Then txtno(0) = 0
    If IsNumeric(txtno(1)) = False Then txtno(1) = txtno(0)
    If CLng(txtno(1)) < CLng(txtno(0)) Then txtno(1) = txtno(0)
    '-------
    If Index = 0 Then GProcCrystalRptPreparation frmMain.cryReport, "Printer"
    If Index = 1 Then GProcCrystalRptPreparation frmMain.cryReport, "Window"
    frmMain.cryReport.WindowTitle = "Receipt Voucher"
    Vno = CLng(txtno(0))
    Do While Vno <= txtno(1)
        X = "Select * From tblVoucher where  vNo=" & Vno & " and VCTRNO =1 and VType='" & frmTypeStr & "' and VYear=" & gCYear
        If GProcRstOpen(Rs1, X, "R") > 0 Then
            With frmMain.cryReport
                .SelectionFormula = "{tblVoucher.vNo} = " & Rs1!Vno & " and {tblVoucher.VCTRNO}=1 and {tblVoucher.VType}='" & Rs1!Vtype & "' and {tblVoucher.VYear}=" & Rs1!VYear
                .Formulas(5) = "wAmountInWords = '" & GProcAmountInWords(Abs(Rs1!vamt)) & "'"
                .ReportFileName = gReportPath & "rptVouReceiptPrint.rpt"
                .Action = 1
            End With
        End If
        Vno = Vno + 1
    Loop
Case 2:
    DoEvents
    Frame1.Visible = False
Case 3:
    '------ Validate
    If IsNumeric(txtno(0)) = False Then txtno(0) = 0
    If IsNumeric(txtno(1)) = False Then txtno(1) = txtno(0)
    If CLng(txtno(1)) < CLng(txtno(0)) Then txtno(1) = txtno(0)
    '-------
    If Index = 3 Then GProcCrystalRptPreparation frmMain.cryReport, "Printer"
    If Index = 3 Then GProcCrystalRptPreparation frmMain.cryReport, "Window"
    frmMain.cryReport.WindowTitle = "Short Payment Voucher"
    Vno = CLng(txtno(0))
    Do While Vno <= txtno(1)
        X = "Select * From qryshortrecvoucher where  RecVNo=" & Vno & " and RecVType='" & frmTypeStr & "' and RecVYear=" & gCYear
        If GProcRstOpen(Rs1, X, "R") > 0 Then
            With frmMain.cryReport
                .SelectionFormula = "{qryshortrecvoucher.RecVno} = " & Rs1!RecVno & "  and {qryshortrecvoucher.RecVType}='" & Rs1!RecVType & "' and {qryshortrecvoucher.RecVYear}=" & Rs1!RecVYear
                .ReportFileName = gReportPath & "rptVouReceiptShort.rpt"
                .Action = 1
            End With
        End If
        Vno = Vno + 1
    Loop
End Select
End Sub
Private Sub cmdIntDbNt_Click()
frmLatePayIntDbNt.Show
End Sub

Private Sub cmdLedger_Click(Index As Integer)
On Error GoTo ErrorHandler
If Index = 0 Then
    frmRptAccount.txt(1) = txtTemp(5)
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
            Sendkeys "{TAB}"
        Case 27:
'            If (cmdBtn(13).Enabled = False) Then
'                Unload Me
'            End If
    End Select
End Sub
Private Sub Form_Load()
    Dim X As String
    Dim Y As String
    Dim rstTmp As Recordset
    mLastEntryBank = ""
    Me.Top = 300
    List1.Clear
    List1.AddItem "Cheque"
    List1.AddItem "Transfer"
    List1.AddItem "DD."
    List1.AddItem "Hundi"
    List1.AddItem "Other"
    
    lblBalance.Caption = ""
    Me.Left = Screen.Width / 2 - Me.Width / 2
    frmTypeStr = gfrmTypeStr
    CriteriaStr = " VType = '" & frmTypeStr & "' and VYear=" & gCYear
    Y = "Select * from " & tblName & " where  vno = 0 and   " & CriteriaStr & " and vctrno = 1"
    If Trim(OrderField) <> "" Then Y = Y & " order by " & OrderField
    GProcRstOpen rstbl, Y, "O"
    Select Case frmTypeStr
    Case "BR" '--Bank
        Label6.Caption = "Bank Receipts"
    Case "CR" '--Cash
        Label6.Caption = "Cash Receipts"
        lbl1(10).Visible = False
        mskDt(13).Visible = False
        mskDt(12).Visible = False
        lbl1(8).Visible = False
        lbl1(13).Visible = False
        txttemp1(8).Visible = False
        txt(11).Visible = False
        txttemp1(7).MaxLength = 40
    End Select
    lblBalance.Caption = ""
    FirstTimeFlag = True
    If GProcRstOpen(rstTmp, "select CotCGSTRt,CotSGSTRt,CotIGSTRt,PolCGSTRt,PolSGSTRt,PolIGSTRt from tblMastNarration where Narration = 'G S T'", "R", gCn) > 0 Then
       wCotCGStRt = rstTmp.Fields(0)
       wCotSGStRt = rstTmp.Fields(1)
       wCotIGStRt = rstTmp.Fields(2)
       wPolCGStRt = rstTmp.Fields(3)
       wPolSGStRt = rstTmp.Fields(4)
       wPolIGStRt = rstTmp.Fields(5)
    End If
    '---Grid
    SetGrid
    '---- Button Visiblity Set As Per User Previlage
    GProcSetButtonVisiblity Me
    cmdBtn_Click (6)
End Sub
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
'If gFormAction = vbDataActionUpdate Then
'    MsgBox "You are about to modify the record. Save the record and then exit.", vbInformation + vbOKOnly, Me.Caption
'    Cancel = True
'    Exit Sub
'End If
rstbl.Close
End Sub

Private Sub mskDt_LostFocus(Index As Integer)
SaveBtnEd
If Index = 2 Then ShowAccountBalance
End Sub
Private Sub mskDt_GotFocus(Index As Integer)
GProcSelectBox Me.ActiveControl
End Sub

Private Sub txt_GotFocus(Index As Integer)
GProcSelectBox Me.ActiveControl
End Sub
Private Sub txt_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
Dim OldFrmType As String
Dim LeftPos As Long
Dim TopPos As Long
If KeyCode = vbKeyF3 Then '--F3 Create
    Select Case Index
        Case 8, 9 '--Narration
            OldFrmType = gfrmTypeStr
            gfrmTypeStr = "N"
            GProcShowForm frmMastNarration, frmMain.mnuMstfrm(5), , gfrmTypeStr
            gfrmTypeStr = OldFrmType
    End Select
End If
If KeyCode = vbKeyF2 Then '-- Select
    TopPos = txt(Index).Top + Me.Top + 650
    LeftPos = txt(Index).Left + Me.Left
    Select Case Index
        Case 8, 9 '--Narration
            gClsSearch.SearchMultiField "tblMastNarration", "Narration", Array("Narration"), Array(txt(Index).Width), "NarrType='N'", "", "Narration", LeftPos, TopPos, False
            If gClsSearch.SearchMultiRetCol(0) <> "" Then txt(Index) = gClsSearch.SearchMultiRetCol(0)
    End Select
End If
End Sub
Private Sub txt_LostFocus(Index As Integer)
GProcCheckForNumber Me, rstbl, MaxNo
SaveBtnEd
End Sub
Private Sub txt_KeyPress(Index As Integer, KeyAscii As Integer)
    KeyAscii = GProcValidateKey(rstbl, Index, KeyAscii, txt(Index).text)
End Sub
Private Sub FillCombo()
If (frmTypeStr = "CR") Then
   txtTemp(5) = "Cash In Hand"
End If
End Sub
Private Sub SaveBtnEd()
'--Vno,Vdt,Amount,Cr A/c,Dr A/c
If txt(0) = "" Or GProcIsDateValid(mskDt(2)) = False Or CDbl(txt(6)) = 0 Or txtTemp(4) = "" Or txtTemp(5) = "" Then
    cmdBtn(13).Enabled = False
    cmdBtn(13).BackColor = &H80000004
Else
    cmdBtn(13).Enabled = True
    cmdBtn(13).BackColor = &HFFFF80
End If
End Sub
Private Sub FillTxtFromTemp()
txt(1) = 1
txt(2) = IIf(IsDate(mskDt(2)), mskDt(2), "") '--Vdt
txt(3) = frmTypeStr
txt(12) = IIf(IsDate(mskDt(12)), mskDt(12), "") '--Chqdt
txt(13) = IIf(IsDate(mskDt(13)), mskDt(13), "") '--Rcondt
If Left(frmTypeStr, 1) = "B" Then
   txt(7) = Left(txttemp1(8) + Space(3), 3) + ". No. " + Left(txt(11) + Space(10), 10) + " " + txttemp1(7)
Else
   txt(7) = txttemp1(7)
End If
txt(4) = GProcGetColumnValue("tblMastAccount", "AcName", txtTemp(4), "S", "Accode", "N")
txt(5) = GProcGetColumnValue("tblMastAccount", "AcName", txtTemp(5), "S", "AcCode", "N")
txt(20) = ChkIsTcs.Value
ShowAccountBalance
End Sub
Private Sub FillTempFromTxt()
mskDt(2) = IIf(IsDate(txt(2)), CDate(txt(2)), "__/__/____") '--Vdt
If IsDate(txt(12)) Then
   mskDt(12) = CDate(txt(12))
Else
   mskDt(12) = "__/__/____"
End If
If IsDate(txt(13)) Then
   mskDt(13) = CDate(txt(13))
Else
   mskDt(13) = "__/__/____"
End If
If Left(frmTypeStr, 1) = "B" Then
    txttemp1(7) = Mid(txt(7), 21)
    txttemp1(8) = Left(txt(7), 3)
Else
    txttemp1(7) = txt(7)
End If
txtTemp(4) = GProcGetColumnValue("tblMastAccount", "AcCode", txt(4), "N", "AcName", "S")
txtTemp(5) = GProcGetColumnValue("tblMastAccount", "AcCode", txt(5), "N", "AcName", "S")
If IsNull(txt(20)) Then
   ChkIsTcs.Value = 0
Else
   ChkIsTcs.Value = Val(txt(20))
End If
lblBalance.Caption = ""
'-- Fill Related Record Rec Vs Sale
FillRelatedRecord
ShowAccountBalance
End Sub
Private Function ValidateGrid() As Boolean
Dim TotAmt As Double
Dim i As Long
Dim RCount As Long
Dim j As Long
Dim InvNo As Long
Dim InVtype As String
Dim InvYear As Long
Dim IntAcCode As Long
Dim TDSPaidAcCode As Long
Dim BillDet As String
If Trim(txtGrid.text) <> "" Then
        msGrid.text = txtGrid.text
End If
txtGrid.Visible = False
txtGrid.text = ""
BillDet = ""
CheckForNumberGrid
With msGrid
    '-- Duplicate invoice selection
    For i = 1 To .Rows - 1
        If Trim(.TextMatrix(i, 1)) <> "" Then
           BillDet = BillDet + "B.No." + .TextMatrix(i, 1) + "-" + CStr(Val(.TextMatrix(i, 4)) - Val(.TextMatrix(i, 5)) + Val(.TextMatrix(i, 7)) + Val(.TextMatrix(i, 11))) + ":"
        End If
      
        '-- Int A/c
           If CDbl(.TextMatrix(i, 7)) > 0 Then
                If gLatePayIntAcCode = GProcGetColumnValue("tblmastAccount", "AcName", "-", "S", "AcCode", "N") Then
                    IntAcCode = 0
                Else
                    IntAcCode = gLatePayIntAcCode
                End If
                If IntAcCode = 0 Then
                    MsgBox "Check Late Pay Interest Account.", vbInformation + vbOKOnly, Me.Caption
                    GProcShowForm frmSetting, frmMain.mnuMstSetting(0)
                    Exit Function
                End If
           End If
        '---
        '-- TDS Paid A/c ( For Less TDS Amt JV )
           If CDbl(.TextMatrix(i, 8)) > 0 Then
                If gTDSPaidAcCode = GProcGetColumnValue("tblmastAccount", "AcName", "-", "S", "AcCode", "N") Then
                    TDSPaidAcCode = 0
                Else
                    TDSPaidAcCode = gTDSPaidAcCode
                End If
                If TDSPaidAcCode = 0 Then
                    MsgBox "Check TDS Paid Account.", vbInformation + vbOKOnly, Me.Caption
                    GProcShowForm frmSetting, frmMain.mnuMstSetting(0)
                    Exit Function
                End If
           End If
        '---
        InvNo = .TextMatrix(i, 20)
        InVtype = .TextMatrix(i, 0)
        InvYear = .TextMatrix(i, 19)
        For j = 1 To .Rows - 1
            If j <> i And InvNo = .TextMatrix(j, 20) And InVtype = .TextMatrix(j, 0) And InvYear = .TextMatrix(j, 19) And InvNo <> 0 Then
                MsgBox "Duplicate Invoice Selection.", vbOKOnly, Me.Caption
                msGrid.SetFocus
                Exit Function
            End If
        Next j
    Next i
    If BillDet <> "" Then
        BillDet = Left(Trim(BillDet) + Space(120), 120)
        txt(8) = Left(BillDet, 40)
        txt(9) = Mid(BillDet, 41, 40)
        txt(10) = Mid(BillDet, 81, 40)
    End If
    '-- Adjusted Amt
    If .Rows > 1 Then
       TotAmt = 0
       For i = 1 To .Rows - 1
           '-- Adj amt > Bill Amt
'---TDS-194Q   Delete - Val(.TextMatrix(i, 5))
           If Val(.TextMatrix(i, 4)) > Val(.TextMatrix(i, 3)) Then
'           If Val(.TextMatrix(i, 4)) - Val(.TextMatrix(i, 5)) > Val(.TextMatrix(i, 3)) Then
                MsgBox "Adjusted Amount is more than Balance Bill Amount. ", vbOKOnly, Me.Caption
                msGrid.SetFocus
                Exit Function
           End If
'--- TDS-194Q
           
          If Val(.TextMatrix(i, 5)) > 0 And Val(.TextMatrix(i, 4)) + Val(.TextMatrix(i, 5)) > Val(.TextMatrix(i, 3)) Then
                MsgBox "Pl Check TDS Amount....", vbOKOnly, Me.Caption
                msGrid.SetFocus
                Exit Function
           End If
'-- End TDS-194Q

           If IsDate(.TextMatrix(i, 2)) Then
                If CDate(.TextMatrix(i, 2)) >= CDate("01/07/2017") Then
                     '-- Check GST Rates
'                     If (Val(.TextMatrix(i, 7)) + Val(.TextMatrix(i, 5)) + Val(.TextMatrix(i, 8))) > 0 And (Val(.TextMatrix(i, 13)) + Val(.TextMatrix(i, 14)) + Val(.TextMatrix(i, 15))) = 0 Then
'                         MsgBox "Please Check GST Rates", vbOKOnly, Me.Caption
'                          msGrid.SetFocus
'                          Exit Function
'                     End If
                End If
           End If
'--- TDS-194Q delete - Val(.TextMatrix(i, 5))
           
           TotAmt = TotAmt + Val(.TextMatrix(i, 4)) + Val(.TextMatrix(i, 7)) - Val(.TextMatrix(i, 8)) + Val(.TextMatrix(i, 9))
'           TotAmt = TotAmt + Val(.TextMatrix(i, 4)) - Val(.TextMatrix(i, 5)) + Val(.TextMatrix(i, 7)) - Val(.TextMatrix(i, 8)) + Val(.TextMatrix(i, 9))
'-- End TDS-194Q
        Next i
        If (TotAmt > 0) And (TotAmt <> Val(txt(6))) Then
           MsgBox "Total of Adjusted Amount " & TotAmt & " & Received Amount not tally.", vbOKOnly, Me.Caption
           msGrid.SetFocus
'           Exit Function
        End If
    End If
End With
ValidateGrid = True
End Function
Private Function ValidateData() As Boolean
Dim X As String
Dim GName As String
Dim rstAddCheck As Recordset
GProcCheckForNumber Me, rstbl, MaxNo
'---- Required Data
If ValidateGrid = False Then Exit Function
'--- Vno
If CLng(txt(0)) = 0 Then
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
'---- Dr A/c
If txtTemp(4) = "" Then
    MsgBox "Check Account.", vbInformation + vbOKOnly, Me.Caption
    txtTemp(4).SetFocus
    Exit Function
End If
'--Create Dr A/c
If GProcGetColumnValue("tblMastAccount", "AcName", txtTemp(4), "S", "AcCode", "N") = 0 Then
   GName = GProcGetColumnValue("TblMastGroup", "AgCode", 90017, "N", "AgName", "S")
   GProcShowForm frmMastAccount, frmMain.mnuMstfrm(0), Array(frmMastAccount.txt(1), frmMastAccount.txtTemp(2)), Array(txtTemp(4), GName)
   Exit Function
End If
'--- Cr A/c
If txtTemp(5) = "" Then
    MsgBox "Check Account.", vbInformation + vbOKOnly, Me.Caption
    txtTemp(5).SetFocus
    Exit Function
End If
'--Create Cr A/c
If GProcGetColumnValue("tblMastAccount", "AcName", txtTemp(5), "S", "AcCode", "N") = 0 Then
   GName = GProcGetColumnValue("TblMastGroup", "AgCode", 90019, "N", "AgName", "S")
   GProcShowForm frmMastAccount, frmMain.mnuMstfrm(0), Array(frmMastAccount.txt(1), frmMastAccount.txtTemp(2)), Array(txtTemp(5), GName)
   Exit Function
End If
'--- Cr A/c ,Dr A/c
If txtTemp(5) = txtTemp(4) Then
    MsgBox "Check Account.Debit and Credit Accounts are Same.", vbInformation + vbOKOnly, Me.Caption
    txtTemp(5).SetFocus
    Exit Function
End If
'---- Amount
If CDbl(txt(6)) <= 0 Then
    MsgBox "Check Voucher Amount.", vbInformation + vbOKOnly, Me.Caption
    Exit Function
End If
'----Duplicate Records
'----Addition Case
If FormAction = vbDataActionAddNew Then
    X = "Select vno From " & tblName & " Where vno=" & CLng(txt(0)) & " and Vtype='" & frmTypeStr & "' and VYear=" & gCYear
    i = GProcRstOpen(rstAddCheck, X, "R")
    If i > 0 Then
        MsgBox "Record is Available of Voucher No. " & CLng(txt(0)), vbCritical + vbOKOnly, "Duplicate Record"
        txt(0).SetFocus
        Exit Function
    End If
    rstAddCheck.Close
End If
ValidateData = True
End Function
Private Sub ReadFields()
GProcShowRecord Me, rstbl, MaxNo
FillTempFromTxt
End Sub
Private Sub txttemp_GotFocus(Index As Integer)
GProcSelectBox Me.ActiveControl
End Sub
Private Sub txttemp_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
Dim GName As String
If KeyCode = vbKeyF3 Then
    Select Case Index
        Case 5 '-- cr A/c
         If frmTypeStr <> "CR" Then '-- Bank Receipt
            GName = GProcGetColumnValue("TblMastGroup", "AgCode", 90019, "N", "AgName", "S")
            GProcShowForm frmMastAccount, frmMain.mnuMstfrm(0), Array(frmMastAccount.txtTemp(2)), Array(GName)
         End If
        Case 4 '-- dr A/c
            GName = GProcGetColumnValue("TblMastGroup", "AgCode", 90017, "N", "AgName", "S")
            GProcShowForm frmMastAccount, frmMain.mnuMstfrm(0), Array(frmMastAccount.txtTemp(2)), Array(GName)
    End Select
End If
End Sub
Private Sub txtTemp_KeyPress(Index As Integer, KeyAscii As Integer)
Dim LeftPos As Long
Dim TopPos As Long
TopPos = txtTemp(Index).Top + Me.Top + 650
LeftPos = txtTemp(Index).Left + Me.Left
If KeyAscii = 13 Then Exit Sub
'--Cash Receipt
If (frmTypeStr = "CR") And (Index = 5) Then
   KeyAscii = 0
   txt(Index) = 2
   txtTemp(Index) = GProcGetColumnValue("tblmastaccount", "Accode", 2, "N", "Acname", "S")
   Exit Sub
End If
Select Case Index
    Case 5: '--Bank Receipt
        'gClsSearch.SearchMultiField "tblMastAccount", "AcName,Accode", Array("Account Name", "Aid"), Array(txttemp(index).Width, 0), "AgCode=19 ", Chr(KeyAscii), "AcName", LeftPos, TopPos, True
        gClsSearch.SearchMultiField "QryMastaccount", "Account,AcName,AcCity,AgName", Array("Account", "AcName", "City", "Group"), Array(4000, 0, 2000, 3000), "AgCode=90019", Chr(KeyAscii), "AcName", LeftPos, TopPos, True
        KeyAscii = 0
        If gClsSearch.SearchMultiRetCol(0) <> "" And gClsSearch.SearchMultiRetCol(1) = "" Then
            txtTemp(Index) = gClsSearch.SearchMultiRetCol(0)
        Else
            If gClsSearch.SearchMultiRetCol(1) <> "" Then txtTemp(Index) = gClsSearch.SearchMultiRetCol(1)
        End If
    Case 4: '--Account
        gClsSearch.SearchMultiField "QryMastaccount", "Account,AcName,AcCity,AgName", Array("Account", "AcName", "City", "Group"), Array(3000, 0, 2000, 3000), "", Chr(KeyAscii), "AcName", LeftPos, TopPos, True
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
If Index = 5 Then
   mLastEntryBank = txtTemp(5).text
   mLastEntryDate = mskDt(2)
End If
End Sub

Private Sub txttemp1_GotFocus(Index As Integer)
If Index = 8 Then
   List1.Visible = True
   DoEvents
   List1.SetFocus
End If
End Sub
Private Sub list1_DblClick()
    Call list1_KeyPress(13)
    DoEvents
End Sub
Private Sub list1_GotFocus()
If Trim(txttemp1(8).text) <> "" Then
    List1.text = txttemp1(8).text
Else
    List1.ListIndex = 0
End If
End Sub
Private Sub list1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    txttemp1(8).text = Trim(List1.text)
    List1.Visible = False
    DoEvents
    txt(11).SetFocus
End If
End Sub
Private Sub list1_LostFocus()
    Call list1_KeyPress(13)
    DoEvents
End Sub
'--- Related Record
Private Function IsRelatedRecord() As Boolean
'-- Audited
If GProcISAudited(txt(0), txt(3), gCYear) = True Then
    MsgBox "Record is Audited.You can not Modify/Delete this Record.", vbCritical + vbOKOnly, App.Title
    IsRelatedRecord = True
    Exit Function
End If
If Val(txt(32)) = 0 Then
   MsgBox "Entry Done In VAT Receipt. You can not Modify/Delete this Record.", vbCritical + vbOKOnly, App.Title
   IsRelatedRecord = True
   Exit Function
End If
'''---e invoice check
''Dim X As String
''Dim Rs1 As Recordset
''X = "select * from tbleInvoiceIRN where vtype= 'SN' and vno=" & txt(0) & " and VYear=" & gCYear & " "
''
''i = GProcRstOpen(Rs1, X, "R")
''If i > 0 Then
''   MsgBox "e-Invoice Done", vbCritical + vbOKOnly, App.Title
''   IsRelatedRecord = True
''   Exit Function
''End If
'--- Related Record Receipt
If rstbl.EOF And rstbl.BOF Then
Else
    If rstbl.Fields("VIsRefEntType") = 1 Then
        MsgBox "Related Record is Present.You can not Modify/Delete this Record.", vbCritical + vbOKOnly, App.Title
        IsRelatedRecord = True
        Exit Function
    End If
End If
End Function
Private Sub ShowAccountBalance()
Dim DrAcBal  As Double
Dim CrAcBal As Double
If IsDate(mskDt(2)) Then
    DrAcBal = GProcGetAccountBalance(txtTemp(4), mskDt(2))
    CrAcBal = GProcGetAccountBalance(txtTemp(5), mskDt(2))
    txtDrBalance = IIf(DrAcBal > 0, DrAcBal & " Cr", Abs(DrAcBal) & " Dr")
    txtCrBalance = IIf(CrAcBal > 0, CrAcBal & " Cr", Abs(CrAcBal) & " Dr")
End If
End Sub
'---------------- Grid
Private Sub SetGrid()
Dim X As String, x1Rec As Double
Set rsRecVsSale = New Recordset
X = "select InvTp, InvBillNo, InvDt, BillAmt, AdjAmt, Discount, LateDays,IntReceived,LessTDS,GSTRecd,GSTPend,IntDbAmt,GSTDbAmt,CGSTRt,SGSTRt,IGSTRt,JvNoDisc,JVNOINT,JvNoLessTDS,InvYear,InvNo,DrCrNtNo,Interest" _
& " from tblRecVsSale where " _
& " Recvtype='" & frmTypeStr & "' and Recvno=" & txt(0) & " and RecVYear=" & gCYear & " order by RecVCtrlNo"
x1Rec = GProcRstOpen(rsRecVsSale, X, "R")
With msGrid
    .Clear
'---------------------------------------------------------------------------------------- TDS-194Q
    .FormatString = "<Tp |>Inv. No      |<Inv. Date    |>Invoice Amt    |>Adj. Amount   |>TDS-194Q |>Late|>Int Recev.|>Less TDS|>GST Recd|>GST Pend|>Ind.Db.Amt.|>GST Db.|>C Rt |>S Rt |>I Rt    |||||||>Interest        "
    .ColWidth(16) = 0
    .ColWidth(17) = 0
    .ColWidth(18) = 0
    .ColWidth(19) = 0
    .ColWidth(20) = 0 '-JV No. TDS
    .ColWidth(21) = 0 '-JV No. TDS
    If FormAction = vbDataActionAddNew Then
        .Rows = 2
        GridAddNew
    Else
        .Rows = 2
    End If
End With
End Sub
Private Sub GridAddNew()
Dim i As Long
With msGrid
    For i = 0 To rsRecVsSale.Fields.Count - 1
        Select Case rsRecVsSale.Fields(i).Type
            Case adInteger, 2, 3 '-- Integer
                .TextMatrix(.Row, i) = 0
            Case adDecimal, 131, adNumeric, adDouble '-- Double
                .TextMatrix(.Row, i) = Format(0, GProcNumberFormat(rsRecVsSale.Fields(i).Precision, rsRecVsSale.Fields(i).NumericScale))
            Case adDate '-- Date
                 .TextMatrix(.Row, i) = ""
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
   ' cmdBtn(10).Cancel = False
End Sub
Private Sub msGrid_EnterCell()
If FormAction <> vbDataActionClose Then
    CheckForNumberTxtGrid
    DoEvents
    txtGrid = ""
    DoEvents
    If msGrid.text <> "" Then txtGrid.text = msGrid.text
    '--- Lock TextBox
    With msGrid
        Select Case .Col
'-----------------------------------TDS-194Q Delete 5
            Case 0, 1, 2, 6, 11, 12: '--Type,Inv No, Inv Dt, Debit Note
'            Case 0, 1, 2, 6, 11, 12, 5: '--Type,Inv No, Inv Dt, Debit Note
                txtGrid.Locked = True
            Case Else
                txtGrid.Locked = False
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
        If (Val(.TextMatrix(.Row, 0)) = 0) And .Rows > 2 Then
            .RemoveItem (.Row)
            txtGrid.text = ""
        End If
        txtGrid.Visible = False
        If cmdBtn(13).Enabled = True Then
           cmdBtn(13).SetFocus
        End If
        'txt(18).SetFocus
        Exit Sub
    End If
    '---- Navigate Row ,Col
    If KeyAscii = 13 Then
        KeyAscii = 0
        If .Col = 12 Then  '    rsRecVsSale.Fields.Count - 5 Then  '-4
            If .Row = (.Rows - 1) Then '-- Add New Row
                .Rows = .Rows + 1
                .Row = .Row + 1
'                GridAddNew
            Else '--Next Row
                .Row = .Row + 1
            End If
            .Col = 0
        Else    '-- Next col
            .Col = .Col + 1
        End If
    End If
    
'--'-TDS-194Q

'    If .Col = 6 And Val(.TextMatrix(.Row, 5)) = 0 Then
'       .Col = 7
'    End If
'    If .Col = 6 And Val(.TextMatrix(.Row, 5)) > 0 Then
'       .Col = 13
'    End If

'---End TDS-194Q


'    If .Col = 8 Then
'        If Val(.TextMatrix(.Row, 7)) = 0 Then
'           .Col = 12
'        End If
'    End If
End With

'---- Set TxtGrid
If Trim(msGrid.text) <> "" Then
    txtGrid = Trim(msGrid.text)
End If
If msGrid.Col < (rsRecVsSale.Fields.Count) - 6 Then '- -3
'    If msGrid.Col = 6 Then msGrid.Col = 7
'    If msGrid.Col = 10 Then msGrid.Col = 13
'    If msGrid.Col = 10 Then msGrid.Col = 13
'    If msGrid.Col = 7 Then msGrid.Col = 9
'    If msGrid.Col = 14 Then msGrid.Col = 15
    
    txtGrid.Move msGrid.CellLeft + msGrid.Left, msGrid.CellTop + msGrid.Top, msGrid.CellWidth, msGrid.CellHeight
    msGrid.SetFocus
End If
CalAmount
txtGrid.Visible = True
txtGrid.Enabled = True
DoEvents
If KeyAscii <> 0 Then
    Select Case rsRecVsSale.Fields(msGrid.Col).Type
        Case adChar, adVarWChar
            If msGrid.Col = 0 Then  '-- Invoice
                txtGrid_KeyPress (KeyAscii)
            Else
                txtGrid.text = txtGrid.text & Chr(KeyAscii)
            End If
        Case Else
            txtGrid.text = Chr(GProcValidateKey(rsRecVsSale, msGrid.Col, KeyAscii, txtGrid.text)) 'Chr(GProcNumberWithDecimal(KeyAscii, txtGrid))
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
End Sub
Private Sub msGrid_LeaveCell()
If txtGrid.Visible Then
    msGrid.text = txtGrid.text
    DoEvents
    txtGrid.text = ""
    txtGrid.Visible = False
End If
CalGridAmt
End Sub
Private Sub CalBalance()
Dim k As Integer
With msGrid
    k = 1
    lblBalance.Caption = ""
    DoEvents
    Do While k < .Rows
'--TDS-194Q Delete - Val(.TextMatrix(k, 5))
        lblBalance = Val(lblBalance) + Val(.TextMatrix(k, 4)) + Val(.TextMatrix(k, 7)) + Val(.TextMatrix(k, 9))
'        lblBalance = Val(lblBalance) + Val(.TextMatrix(k, 4)) - Val(.TextMatrix(k, 5)) + Val(.TextMatrix(k, 7)) + Val(.TextMatrix(k, 9))
        k = k + 1
        DoEvents
    Loop
    DoEvents
    If Val(lblBalance) > Val(txt(6)) Then MsgBox "Amount Should be less than or equal to " & txt(6), vbCritical + vbOKOnly, "Amount not Tally"
End With
End Sub
Private Sub txtGrid_GotFocus()
    txtGrid.Alignment = ProcSetAlignment(rsRecVsSale, msGrid.Col)
    DoEvents
    Call GProcSelectBox(txtGrid)
End Sub
Private Sub txtGrid_KeyPress(KeyAscii As Integer)
Dim LeftPos As Long
Dim TopPos As Long
Dim PartyCode As Long
TopPos = txtGrid.Top + Me.Top + 650
LeftPos = txtGrid.Left + Me.Left
DoEvents
With msGrid
    If KeyAscii <> 13 And KeyAscii <> 27 Then
        '--- Search
        Select Case .Col
            Case 0 '-- Invoice
                DoEvents
                KeyAscii = 0
                Pending_Invoice (1)
            Case Else
                KeyAscii = GProcValidateKey(rsRecVsSale, msGrid.Col, KeyAscii, txtGrid)
        End Select
   End If
  If KeyAscii = 13 Then
     CheckForNumberTxtGrid
     '----Adj Amt <=Inv Amt & Adj Dr Note Amt <  Dr Note Amt
     Select Case .Col
     Case 4:
          If Val(txtGrid) > Val(.TextMatrix(.Row, 3)) Then
                MsgBox "Adjustment Amount must less than Invoice Amount", vbCritical + vbOKOnly, "Adjustment Amount is less than Invoice Amount"
                DoEvents
                txtGrid.SetFocus
                Exit Sub
            End If
        Case 5:
            If Val(txtGrid) > Val(.TextMatrix(.Row, 3)) Then
'-----------------------TDS-194Q
                MsgBox "TDS Amount must less than Invoice Amount", vbCritical + vbOKOnly, "Adjustment Dr Note Amount is less than Dr Note Amount"
                DoEvents
                txtGrid.SetFocus
                Exit Sub
            End If
        End Select
    End If
    '--- Navigate Row ,Col
    If KeyAscii = 13 And .Col < 12 Then  'rsRecVsSale.Fields.Count - 5 Then '- -4
        .Col = .Col + 1
    ElseIf KeyAscii = 13 And .Col = 12 Then   '    rsRecVsSale.Fields.Count - 5 Then  '- -4
        If .Row = (.Rows - 1) Then
            .Rows = .Rows + 1
            .Row = .Row + 1
'            GridAddNew
        Else
            .Row = .Row + 1
        End If
        .Col = 0
    End If
    '-- On ESC
    With msGrid
        If KeyAscii = 27 Then '----Esc Key
            KeyAscii = 0
            If (Val(.TextMatrix(.Row, 4)) + Val(.TextMatrix(.Row, 7)) + Val(.TextMatrix(.Row, 9)) = 0) And .Rows > 2 Then
                .RemoveItem (.Row)
                txtGrid.text = ""
            End If
            txtGrid.Visible = False
            If cmdBtn(13).Enabled = True Then
               cmdBtn(13).SetFocus
            End If
            Exit Sub
        End If
    End With
End With
End Sub
Private Sub txtGrid_KeyDown(KeyCode As Integer, Shift As Integer)
'--- Delete Row On Ctrl+Y
If KeyCode = vbKeyY And Shift = 2 And msGrid.Rows > 2 Then
    msGrid = 4
    msGrid.RemoveItem (msGrid.Row)
    txtGrid.Visible = False
    txtGrid.text = ""
    'CalAmount
    Exit Sub
End If
If KeyCode = vbKeyF4 Then
   DoEvents
   Pending_Invoice (2)
End If
End Sub
Private Sub CheckForNumberGrid()
Dim RowIndex As Long
With rsRecVsSale
For RowIndex = 1 To msGrid.Rows - 1
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
With rsRecVsSale
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
                txtGrid = ""
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

Private Sub DeleteRelatedRecord()
'--O/s
gCn.Execute "delete from tblOutStanding where vtype='SN' and VNo in (select JVNOINT from tblRecVsSale where " _
& " INTRECEIVED>0 and RecVtype='" & frmTypeStr & "' and RecVNo=" & txt(0) & " and  RecVyear= " & gCYear & ") and " _
& " VYear=" & gCYear
'--- JV Int
gCn.Execute "delete from tblVoucher where Vtype='SN' and VNo in (select JVNOINT from tblRecVsSale where " _
& " INTRECEIVED>0 and RecVtype='" & frmTypeStr & "' and RecVNo=" & txt(0) & " and  RecVyear= " & gCYear & ")  and " _
 & " Vyear= " & gCYear

 '-- JV Less TDS
 gCn.Execute "delete from tblVoucher where Vtype='JV' and VNo in (select JVNOLessTDS from tblRecVsSale where " _
& " LessTDS>0 and RecVtype='" & frmTypeStr & "' and RecVNo=" & txt(0) & " and  RecVyear= " & gCYear & ")  and " _
 & " Vyear= " & gCYear


'---JV TDS-194Q
 gCn.Execute "delete from tblVoucher where Vtype='JV' and VNo in (select JvNoDisc from tblRecVsSale where " _
& " Discount>0 and RecVtype='" & frmTypeStr & "' and RecVNo=" & txt(0) & " and  RecVyear= " & gCYear & ")  and " _
 & " Vyear= " & gCYear
'---JV Discount
' gCn.Execute "delete from tblVoucher where Vtype='PN' and VNo in (select JvNoDisc from tblRecVsSale where " _
'& " Discount>0 and RecVtype='" & frmTypeStr & "' and RecVNo=" & txt(0) & " and  RecVyear= " & gCYear & ")  and " _
' & " Vyear= " & gCYear

'--End JV TDS-194Q


'--- Rec Vs Sale
gCn.Execute "delete from tblRecVsSale where RecVtype='" & frmTypeStr & "' and RecVNo=" & txt(0) & " and " _
 & " RecVyear= " & gCYear
End Sub
Private Sub SaveRelatedRecord()
Dim i As Long, X As String
Dim j As Long
Dim Y As String
Dim Vsr As Long
Dim IntRecAmt As Double
Dim JVVNo As Long
Dim JVTDSVNo As Long '- Less TDS JV VNo
Dim JVDiscNo As Long '- Less Discount JV VNo
Dim DbNtNo As String  '- Less Discount JV VNo
Dim LessTDSAmt As Double
Dim LessDiscAmt As Double
Dim wTaxableAmt As Double
Dim wCgstRt As Double
Dim wCgstAmt As Double
Dim wSCgstRt As Double
Dim wSgstAmt As Double
Dim wIgstRt As Double
Dim wIgstAmt As Double
Dim NetDbAmt As Double
Dim rsRecVsSale As Recordset, rsOutst As Recordset, rstTmp As Recordset
Dim wBroker As Long
Dim wMillCd As Long
Dim TaxAcCodeCGST As Long
Dim TaxAcCodeSGST As Long
Dim TaxAcCodeIGST As Long
Dim PartyCd As Long
Dim TaxAcCodeCGSTDisc As Long
Dim TaxAcCodeSGSTDisc As Long
Dim TaxAcCodeIGSTDisc As Long
Dim IndDbAmt As Double

Dim Nt1 As String
Dim Nt2 As String
Dim Nt3 As String
Dim Nt4 As String
TaxAcCodeCGST = GProcGetColumnValue("tblMastNarration", "Narration", "G S T", "S", "CGSTInPutAc", "N")
TaxAcCodeSGST = GProcGetColumnValue("tblMastNarration", "Narration", "G S T", "S", "SGSTInPutAc", "N")
TaxAcCodeIGST = GProcGetColumnValue("tblMastNarration", "Narration", "G S T", "S", "IGSTInPutAc", "N")
Dim RowIndex As Integer
RowIndex = 1
With msGrid
    Do While RowIndex < .Rows
       If (Val(.TextMatrix(RowIndex, 4)) + Val(.TextMatrix(RowIndex, 5)) + Val(.TextMatrix(RowIndex, 7)) + Val(.TextMatrix(RowIndex, 8)) + Val(.TextMatrix(RowIndex, 9)) = 0) And .Row >= 2 Then
          .RemoveItem (RowIndex)
       End If
       RowIndex = RowIndex + 1
    Loop
End With
'-- Delete Records
DeleteRelatedRecord
'--- Save Grid Records
'--- tblRecSale
X = "select * from tblRecVsSale where Recvtype='" & frmTypeStr & "' and Recvno=" & txt(0) & " and " _
& " RecVYear=" & gCYear
GProcRstOpen rsRecVsSale, X, "O"
With msGrid
    i = 1
    Do While i < .Rows
        DoEvents
        If CDec(Val(.TextMatrix(i, 4))) + CDec(Val(.TextMatrix(i, 5))) + CDec(Val(.TextMatrix(i, 7))) + CDec(Val(.TextMatrix(i, 8))) + CDec(Val(.TextMatrix(i, 9))) > 0 Then
        rsRecVsSale.AddNew
        rsRecVsSale!AdjNo = GProcGenerateId("tblRecVsSale", "ADJNO", "")
        rsRecVsSale!RecVno = CLng(txt(0)) '--RecVNo
        rsRecVsSale!RecVType = UCase(txt(3)) '--RecVtype
        rsRecVsSale!RecVdt = CDate(txt(2)) '--Vdt(2)
        rsRecVsSale!RecVCtrlNo = 1 'i '--Sr No(3)
        rsRecVsSale!RecVYear = gCYear
        rsRecVsSale!RecAmt = CDec(txt(6))
        rsRecVsSale!InvBillno = .TextMatrix(i, 1) & vbNullString
        rsRecVsSale!RecVCtrlNo = 1 'i
        rsRecVsSale!InvNo = CLng(.TextMatrix(i, 20)) '--Invno
        rsRecVsSale!InvTp = .TextMatrix(i, 0) '--InvTp
        rsRecVsSale!InvDt = CDate(.TextMatrix(i, 2)) '--Vdt(2)
        rsRecVsSale!InvYear = .TextMatrix(i, 19)
        rsRecVsSale!BillAmt = CDec(.TextMatrix(i, 3))
        rsRecVsSale!PartyCode = GProcGetColumnValue("tblMastAccount", "AcName", txtTemp(4), "S", "AcCode", "N")
        rsRecVsSale!AdjAmt = CDec(.TextMatrix(i, 4))
        rsRecVsSale!Discount = CDec(.TextMatrix(i, 5))
        rsRecVsSale!LateDays = CInt(.TextMatrix(i, 6))
        rsRecVsSale!Interest = CDec(.TextMatrix(i, 22))
        rsRecVsSale!IntReceived = CDec(Val(.TextMatrix(i, 7)))
        IntRecAmt = CDec(Val(.TextMatrix(i, 7)))
        rsRecVsSale!LessTDS = CDec(Val(.TextMatrix(i, 8))) '- Less TDS
        LessTDSAmt = CDec(Val(.TextMatrix(i, 8)))
        LessDiscAmt = CDec(.TextMatrix(i, 5))
        rsRecVsSale!CGSTRt = Val(.TextMatrix(i, 13))
        rsRecVsSale!SGSTRt = Val(.TextMatrix(i, 14))
        rsRecVsSale!IGSTRt = Val(.TextMatrix(i, 15))
        rsRecVsSale!GSTRecd = Val(.TextMatrix(i, 9))
        rsRecVsSale!GSTPend = Val(.TextMatrix(i, 10))
        JVVNo = 0
        JVTDSVNo = 0
        JVDiscNo = 0
        DbNtNo = ""
        IndDbAmt = 0
        wCgstAmt = 0
        wSgstAmt = 0
        wIgstAmt = 0
        '------------------------- Interest Rec JV
        If FormAction = vbDataActionAddNew And IntRecAmt > 0 Then
            JVVNo = GProcGenerateId("tblVoucher", "Vno", " vyear=" & gCYear & " and Vtype='SN'")
            DbNtNo = GProcGenerateCrDrNtNo("SN", "B", "BD")
        Else
            If IntRecAmt > 0 Then
                If CLng(.TextMatrix(i, 17)) > 0 And GProcGetColumnValue("tblVoucher", "Vno", .TextMatrix(i, 17), "N", "Vno", "N", " Vtype='SN' and Vyear=" & gCYear) = 0 Then
                    JVVNo = CLng(.TextMatrix(i, 17))
                Else
                    JVVNo = GProcGenerateId("tblVoucher", "Vno", " vyear=" & gCYear & " and Vtype='SN'")
                End If
                If Len(.TextMatrix(i, 21)) > 0 And GProcGetColumnValue("tblVoucher", "Vno", .TextMatrix(i, 17), "N", "VDrCrNtNo", "S", " Vtype='SN' and Vyear=" & gCYear) = "" Then
                    DbNtNo = .TextMatrix(i, 21) & vbNullString
                Else
                    DbNtNo = GProcGenerateCrDrNtNo("SN", "B", "BD")
                End If
            End If
        End If
        If Val(.TextMatrix(i, 11)) + Val(.TextMatrix(i, 12)) > 0 Then
           If Val(msGrid.TextMatrix(msGrid.Row, 13)) + Val(msGrid.TextMatrix(msGrid.Row, 14)) > 0 Then
              wCgstAmt = Val(.TextMatrix(i, 12)) / 2
              wSgstAmt = Val(.TextMatrix(i, 12)) / 2
           End If
           If Val(msGrid.TextMatrix(msGrid.Row, 15)) > 0 Then
              wIgstAmt = Val(.TextMatrix(i, 12))
           End If
           IndDbAmt = Val(.TextMatrix(i, 11))
           NetDbAmt = (IndDbAmt) + (wCgstAmt + wSgstAmt + wIgstAmt)
           Nt1 = Left("Int.Recd.Bill-" & .TextMatrix(i, 1) & "-" & "BillAmt-" & CDbl(.TextMatrix(i, 3)) & "-" & CDate(.TextMatrix(i, 2)), 50)
           If LessTDSAmt > 0 Then
              Nt2 = Left("Int-" + CStr(IndDbAmt) + " Less TDS-" + CStr(LessTDSAmt) + " GST-" + CStr(wCgstAmt + wSgstAmt + wIgstAmt), 40)
           Else
              Nt2 = Left("Int-" + CStr(IndDbAmt) + " GST-" + CStr(wCgstAmt + wSgstAmt + wIgstAmt), 40)
           End If
           Nt3 = "Debit Not No : " + DbNtNo
           Y = "Select ITTMP2 from tblOutStanding,tblMastItem where Vno = " & CLng(.TextMatrix(i, 20)) & " and vtype = '" & .TextMatrix(i, 0) & "' and vyear = " & .TextMatrix(i, 19) & " and OutItemCode = tblMastItem.Itcode "
           j = GProcRstOpen(rstTmp, Y, "R")
           Nt4 = ""
           If j > 0 Then
              Nt4 = rstTmp.Fields(0)
           End If
           rstTmp.Close
           
           
           GProcCreateVoucher FormAction, JVVNo, 1, mskDt(2), "SN", txt(4), gLatePayIntAcCode, -(NetDbAmt), .TextMatrix(i, 1), CDate(.TextMatrix(i, 2)), , Nar1:=Nt1, Nar2:=Nt2, Nar3:=Nt3, IsRefEntryType:=True, Taxable:=IndDbAmt, Crt:=Val(.TextMatrix(i, 13)), Camt:=wCgstAmt, Srt:=Val(.TextMatrix(i, 14)), Samt:=wSgstAmt, Irt:=Val(.TextMatrix(i, 15)), Iamt:=wIgstAmt, AgtVno:=CLng(.TextMatrix(i, 20)), AgtVtype:=.TextMatrix(i, 0), AgtVyear:=.TextMatrix(i, 19), DrCrNtNo:=DbNtNo, DrCrNtType:="B", HsnCd:=Nt4
           GProcCreateVoucher FormAction, JVVNo, 2, mskDt(2), "SN", gLatePayIntAcCode, txt(4), (IndDbAmt), .TextMatrix(i, 1), CDate(.TextMatrix(i, 2)), , Nar1:=Nt1, Nar2:=Nt2, Nar3:=Nt3, IsRefEntryType:=True, Taxable:=IndDbAmt, Crt:=Val(.TextMatrix(i, 13)), Camt:=wCgstAmt, Srt:=Val(.TextMatrix(i, 14)), Samt:=wSgstAmt, Irt:=Val(.TextMatrix(i, 15)), Iamt:=wIgstAmt, AgtVno:=CLng(.TextMatrix(i, 20)), AgtVtype:=.TextMatrix(i, 0), AgtVyear:=.TextMatrix(i, 19), DrCrNtNo:=DbNtNo, DrCrNtType:="B", HsnCd:=Nt4
           Vsr = 2
           If wCgstAmt > 0 Then
              Vsr = Vsr + 1
              GProcCreateVoucher FormAction, JVVNo, Vsr, mskDt(2), "SN", TaxAcCodeCGST, txt(4), (wCgstAmt), .TextMatrix(i, 1), CDate(.TextMatrix(i, 2)), , "Debit Note No :  " & DbNtNo & " ", "", "", "", IsRefEntryType:=True, Taxable:=wTaxableAmt, Crt:=Val(.TextMatrix(i, 13)), Camt:=wCgstAmt, Srt:=Val(.TextMatrix(i, 14)), Samt:=wSgstAmt, Irt:=Val(.TextMatrix(i, 15)), Iamt:=wIgstAmt, AgtVno:=CLng(.TextMatrix(i, 20)), AgtVtype:=.TextMatrix(i, 0), AgtVyear:=.TextMatrix(i, 19), DrCrNtNo:=DbNtNo, DrCrNtType:="B"
           End If
           If wSgstAmt > 0 Then
              Vsr = Vsr + 1
              GProcCreateVoucher FormAction, JVVNo, Vsr, mskDt(2), "SN", TaxAcCodeSGST, txt(4), (wSgstAmt), .TextMatrix(i, 1), CDate(.TextMatrix(i, 2)), , "Debit Note No :  " & DbNtNo & " ", "", "", "", IsRefEntryType:=True, Taxable:=wTaxableAmt, Crt:=Val(.TextMatrix(i, 13)), Camt:=wCgstAmt, Srt:=Val(.TextMatrix(i, 14)), Samt:=wSgstAmt, Irt:=Val(.TextMatrix(i, 15)), Iamt:=wIgstAmt, AgtVno:=CLng(.TextMatrix(i, 20)), AgtVtype:=.TextMatrix(i, 0), AgtVyear:=.TextMatrix(i, 19), DrCrNtNo:=DbNtNo, DrCrNtType:="B"
           End If
           If wIgstAmt > 0 Then
              Vsr = Vsr + 1
              GProcCreateVoucher FormAction, JVVNo, Vsr, mskDt(2), "SN", TaxAcCodeIGST, txt(4), (wIgstAmt), .TextMatrix(i, 1), CDate(.TextMatrix(i, 2)), , "Debit Note No :  " & DbNtNo & " ", "", "", "", IsRefEntryType:=True, Taxable:=wTaxableAmt, Crt:=Val(.TextMatrix(i, 13)), Camt:=wCgstAmt, Srt:=Val(.TextMatrix(i, 14)), Samt:=wSgstAmt, Irt:=Val(.TextMatrix(i, 15)), Iamt:=wIgstAmt, AgtVno:=CLng(.TextMatrix(i, 20)), AgtVtype:=.TextMatrix(i, 0), AgtVyear:=.TextMatrix(i, 19), DrCrNtNo:=DbNtNo, DrCrNtType:="B"
           End If
            '--- Pend GST OS
            If Val(.TextMatrix(i, 10)) > 0 Then
               Y = "Select OutBrokerCode,OutMillCode from tblOutStanding where Vno = " & CLng(.TextMatrix(i, 20)) & " and vtype = '" & .TextMatrix(i, 0) & "' and vyear = " & .TextMatrix(i, 19) & ""
               j = GProcRstOpen(rstTmp, Y, "R")
               If j > 0 Then
                  wBroker = Val(rstTmp.Fields(0) & vbNullString)
                  wMillCd = Val(rstTmp.Fields(1) & vbNullString)
               Else
                  wBroker = GProcGetColumnValue("tblMastAccount", "AcName", "-", "S", "AcCode", "N")
               End If
               PartyCd = GProcGetColumnValue("tblMastAccount", "AcName", txtTemp(4), "S", "AcCode", "N")
               GProcCreateOutStanding JVVNo, CDate(txt(2)), "SN", DbNtNo, PartyCd, wBroker, Val(.TextMatrix(i, 10)), 0, CDate(txt(2)), wMillCd, 0, 0
            End If
            '-------------------------------
        End If
        '-------------------------------
'        '------------------------- Less TDS  JV
        If FormAction = vbDataActionAddNew And LessTDSAmt > 0 Then
            JVTDSVNo = GProcGenerateId("tblVoucher", "Vno", " vyear=" & gCYear & " and Vtype='JV'")
        Else
            If LessTDSAmt > 0 Then
                If CLng(.TextMatrix(i, 18)) > 0 And GProcGetColumnValue("tblVoucher", "Vno", .TextMatrix(i, 18), "N", "Vno", "N", " Vtype='JV' and Vyear=" & gCYear) = 0 Then
                    JVTDSVNo = CLng(.TextMatrix(i, 18))
                Else
                    JVTDSVNo = GProcGenerateId("tblVoucher", "Vno", " vyear=" & gCYear & " and Vtype='JV'")
                End If
'-- TDS-194Q
            Else
               JVTDSVNo = 0
'--TDS-194Q
            End If
        End If
        If LessTDSAmt > 0 Then
            GProcCreateVoucher FormAction, JVTDSVNo, 1, mskDt(2), "JV", txt(4), gTDSPaidAcCode, LessTDSAmt, IsRefEntryType:=True, Nar1:="TDS on Interest " & IntRecAmt, Nar2:="Agt Bill-" & .TextMatrix(i, 1) & "-" & CDate(.TextMatrix(i, 2))
            GProcCreateVoucher FormAction, JVTDSVNo, 2, mskDt(2), "JV", gTDSPaidAcCode, txt(4), -LessTDSAmt, IsRefEntryType:=True, Nar1:="TDS on Interest " & IntRecAmt, Nar2:="Agt Bill-" & .TextMatrix(i, 1) & "-" & CDate(.TextMatrix(i, 2))
        End If
'        '-------------------------------
        
        
        
'------------------------- TDS-194Q  delete For OLD Less Disc  JV
        If FormAction = vbDataActionAddNew And LessDiscAmt > 0 Then
            JVDiscNo = GProcGenerateId("tblVoucher", "Vno", " vyear=" & gCYear & " and Vtype='JV'")
'            DbNtNo = GProcGenerateCrDrNtNo("SN", "B", "BC")
        Else
            If LessDiscAmt > 0 Then
                If CLng(.TextMatrix(i, 16)) > 0 Then
                    JVDiscNo = CLng(.TextMatrix(i, 16))
                Else
                    JVDiscNo = GProcGenerateId("tblVoucher", "Vno", " vyear=" & gCYear & " and Vtype='JV'")
                End If
            Else
               JVDiscNo = 0
            End If
        End If
        If LessDiscAmt > 0 Then
           Nt1 = Left("TDS Agt. Bill-" & .TextMatrix(i, 1) & "-" & CDate(.TextMatrix(i, 2)), 40)
           Nt2 = ""
           Nt3 = ""
           GProcCreateVoucher FormAction, JVDiscNo, 1, mskDt(2), "JV", txt(4), gBankCommAcCode, LessDiscAmt, .TextMatrix(i, 1), CDate(.TextMatrix(i, 2)), , Nar1:=Nt1, Nar2:=Nt2, Nar3:=Nt3, IsRefEntryType:=True, AgtVno:=CLng(.TextMatrix(i, 20)), AgtVtype:=.TextMatrix(i, 0), AgtVyear:=.TextMatrix(i, 19)
           GProcCreateVoucher FormAction, JVDiscNo, 2, mskDt(2), "JV", gBankCommAcCode, txt(4), -(LessDiscAmt), .TextMatrix(i, 1), CDate(.TextMatrix(i, 2)), , Nar1:=Nt1, Nar2:=Nt2, Nar3:=Nt3, IsRefEntryType:=True, AgtVno:=CLng(.TextMatrix(i, 20)), AgtVtype:=.TextMatrix(i, 0), AgtVyear:=.TextMatrix(i, 19)
        End If
'-------------------------End TDS-194Q
        
        
        
        
        
        
        
        rsRecVsSale!IntDbAmt = IndDbAmt
        rsRecVsSale!GSTDbAmt = wCgstAmt + wSgstAmt + wIgstAmt
        rsRecVsSale!JvNoLessTDS = JVTDSVNo '-- TDS JV No
        rsRecVsSale!JVNOINT = JVVNo
        rsRecVsSale!JvNoDisc = JVDiscNo
        rsRecVsSale!DrCrNtNo = DbNtNo & vbNullString
        rsRecVsSale.Update
        
'        GProcUpdateGSTPenging CLng(.TextMatrix(i, 19)), .TextMatrix(i, 0), CLng(.TextMatrix(i, 19)), .TextMatrix(i, 19)
        
        
        End If
        DoEvents
        i = i + 1
     Loop
 rsRecVsSale.Close
FillRelatedRecord
End With
End Sub
Private Sub FillRelatedRecord()
Dim RowIndex As Long
Dim ColIndex As Long
'--Fill Grid
SetGrid
With rsRecVsSale
    If Not .EOF Then
        .MoveFirst
        RowIndex = 1
        msGrid.Rows = 2
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
            If Not .EOF Then msGrid.Rows = msGrid.Rows + 1
        Loop
    End If
End With
End Sub
Private Sub Pending_Invoice(Index)
Dim X As String
Dim wFromCd As Long
Dim LateDayCd As String
Dim vamt As Double
Dim rstOutStanding As Recordset
Dim i As Long, j As Long
Dim NewRow As Boolean
Dim TotCol(0, 1) As Long
Dim wItemType
Dim wMsOmsParty As String
vamt = Val(txt(6))
With msGrid
     For i = 1 To .Rows - 1
         vamt = vamt - (Val(.TextMatrix(i, 4)) + Val(.TextMatrix(i, 7)) - Val(.TextMatrix(i, 8)) + Val(.TextMatrix(i, 9)))
     Next i
End With






If Trim(txtTemp(4)) <> "" Then
    wFromCd = GProcGetColumnValue("tblMastAccount", "AcName", txtTemp(4), "S", "AcCode", "N")
    If Index = 2 Then
       LateDayCd = "cdate('" & (CDate(mskDt(2))) & "') " & " - cdate(outDueDate)"
       X = "select VType, outbillNo,format(vdt,'dd/mm/yyyy'), OutAmount,OutAmount-OutRecAmt" _
           & "," & LateDayCd & ", VYear,Vno,OutAmount,OutDbNtAmt,tblMastAccount.Acname,iif(isnull(OutItemCode),0,  OutItemCode) from tblOutstanding, tblMastAccount where tblOutstanding.OutBrokerCode = tblMastAccount.AcCode and OutAcCode=" & wFromCd & " "
    Else
       LateDayCd = "cdate('" & (CDate(mskDt(2))) & "') " & " - cdate(outDueDate)"
       X = "select VType, outbillNo,format(vdt,'dd/mm/yyyy'), OutAmount,OutAmount-OutRecAmt" _
           & "," & LateDayCd & ", VYear,Vno,OutAmount,OutDbNtAmt,tblMastAccount.Acname,iif(isnull(OutItemCode),0,  OutItemCode) from tblOutstanding, tblMastAccount where tblOutstanding.OutBrokerCode = tblMastAccount.AcCode and OutAcCode=" & wFromCd & " and " _
           & " (OutAmount-OutRecAmt) > 0  and vtype <> 'RY'" 'or (OutDbNtAmt-OutRecDbAmt) > 0)
    End If
    TotCol(0, 0) = 4
    TotCol(0, 1) = vamt 'CDbl(txt(6))
    gClsSearch.SearchMultiSelField X, "", Array("Type", "Invoice No", "Date", "Bill Amt", "Balance Amt", "Due Days", "VYear", "INV Vno", "Invoice Amt", "DbNt Amt", "Broker", ""), Array(800, 1200, 1200, 1500, 1500, 800, 0, 0, 0, 0, 2000, 0), "", "", "Vdt,vno desc", 0, 0, , True, TotOfSelCols:=TotCol
'    vamt = CDbl(txt(6))
    With msGrid
        If gClsSearch.SearchMultiSelRetCol(0, 0) <> "" Then
                j = .Row
                If Index <> 2 Then
                    txtGrid = gClsSearch.SearchMultiSelRetCol(0, 0)
                End If
                For i = 0 To UBound(gClsSearch.SearchMultiSelRetCol)
                    'For j = 1 To msfBillDtl.Rows - 1
                       'i + 1
                     If .Rows - 1 = j Then
                     Else
                        .Rows = .Rows + 1
                     End If
                    .TextMatrix(j, 0) = gClsSearch.SearchMultiSelRetCol(i, 0) '--Type
                    .TextMatrix(j, 1) = gClsSearch.SearchMultiSelRetCol(i, 1) '--Bill No
                    .TextMatrix(j, 2) = Format(gClsSearch.SearchMultiSelRetCol(i, 2), "dd/mm/yyyy") '--VDt
                    .TextMatrix(j, 3) = gClsSearch.SearchMultiSelRetCol(i, 3) '--Bill Amt
                    If vamt >= Val(gClsSearch.SearchMultiSelRetCol(i, 4)) Then  '- Adj Amt
                       .TextMatrix(j, 4) = gClsSearch.SearchMultiSelRetCol(i, 4)
                        vamt = vamt - Val(gClsSearch.SearchMultiSelRetCol(i, 4))
                    Else
                       .TextMatrix(j, 4) = vamt
                       vamt = 0
                    End If
                    .TextMatrix(j, 22) = Format(Round((((Val(.TextMatrix(j, 4)) * 1.5) / 3000) * Val(.TextMatrix(j, 6))), 2), "###0.00")
                    
                    
'                    .TextMatrix(j, 4) = gClsSearch.SearchMultiSelRetCol(i, 4) '--adj amt
                    .TextMatrix(j, 6) = CLng(gClsSearch.SearchMultiSelRetCol(i, 5)) '--Late Day
                    .TextMatrix(j, 19) = CLng(gClsSearch.SearchMultiSelRetCol(i, 6)) '--Year
                    .TextMatrix(j, 20) = gClsSearch.SearchMultiSelRetCol(i, 7) '--Vno
                    If gClsSearch.SearchMultiSelRetCol(i, 11) > 0 And Format(gClsSearch.SearchMultiSelRetCol(i, 2), "dd/mm/yyyy") >= CDate("01/07/2017") Then
                        wItemType = GProcGetColumnValue("tblMastItem", "ItCode", "" & gClsSearch.SearchMultiSelRetCol(i, 11) & "", "N", "ItType", "N")
                        wMsOmsParty = GProcGetColumnValue("tblMastAccount", "AcName", txtTemp(4), "S", "AcGSTIN", "S")
                        If Left(wMsOmsParty, 2) = "27" Or Left(wMsOmsParty, 1) = "U" Or wMsOmsParty = "" Or IsNull(wMsOmsParty) Then
                           If wItemType = 0 Then  '--Cotton Rate
                              .TextMatrix(j, 13) = wCotCGStRt
                              .TextMatrix(j, 14) = wCotSGStRt
                              .TextMatrix(j, 15) = 0
                           Else
                              If Format(.TextMatrix(j, 2), "YYYY/MM/dd") < Format("2025/09/22", "YYYY/MM/dd") Then
                                 .TextMatrix(j, 13) = 6
                                 .TextMatrix(j, 14) = 6
                                 .TextMatrix(j, 15) = 0
                              Else
                                 .TextMatrix(j, 13) = wPolCGStRt
                                 .TextMatrix(j, 14) = wPolSGStRt
                                 .TextMatrix(j, 15) = 0
                              End If
                           End If
                        Else
                           If wItemType = 0 Then  '--Cotton Rate
                              .TextMatrix(j, 13) = 0
                              .TextMatrix(j, 14) = 0
                              .TextMatrix(j, 15) = wCotIGStRt
                           Else
                              .TextMatrix(j, 13) = 0
                              .TextMatrix(j, 14) = 0
                              If Format(.TextMatrix(j, 2), "YYYY/MM/dd") < Format("2025/09/22", "YYYY/MM/dd") Then
                                 .TextMatrix(j, 15) = 12
                              Else
                                 .TextMatrix(j, 15) = wPolIGStRt
                              End If
                           End If
                        End If
                    Else
                       .TextMatrix(j, 13) = 0
                       .TextMatrix(j, 14) = 0
                       .TextMatrix(j, 15) = 0
                    End If
                    .TextMatrix(j, 22) = Format(Round((((Val(.TextMatrix(j, 4)) * 1.5) / 3000) * Val(.TextMatrix(j, 6))), 2), "###0.00")
                     j = j + 1
                Next i
'                For i = 1 To .Rows - 1
'                    If (Val(.TextMatrix(i, 4)) + Val(.TextMatrix(i, 7)) - Val(.TextMatrix(i, 8)) + Val(.TextMatrix(i, 9))) > 0 Then
'                        vamt = vamt - (Val(.TextMatrix(i, 4)) + Val(.TextMatrix(i, 7)) - Val(.TextMatrix(i, 8)) + Val(.TextMatrix(i, 9)))
'                    Else
'                        If vamt <= Val(.TextMatrix(i, 4)) Then  '- Adj Amt
'                            .TextMatrix(i, 4) = Val(.TextMatrix(i, 3))
'                            vamt = vamt - Val(.TextMatrix(i, 3))
'                        Else
'                            .TextMatrix(i, 4) = vamt
'                            vamt = 0
'                            Exit For
'                        End If
'                    End If
'                Next i
                .Col = 4
        End If
    End With
Else
    MsgBox "Select Party Compulsary", vbCritical + vbOKOnly, "Select Party"
    DoEvents
    txtTemp(4).SetFocus
End If
End Sub



Private Sub txttemp1_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
Dim OldFrmType As String
Dim LeftPos As Long
Dim TopPos As Long
If KeyCode = vbKeyF3 Then '--F3 Create
    Select Case Index
        Case 7 '--Narration
            OldFrmType = gfrmTypeStr
            gfrmTypeStr = "N"
            GProcShowForm frmMastNarration, frmMain.mnuMstfrm(5), , gfrmTypeStr
            gfrmTypeStr = OldFrmType
    End Select
End If
If KeyCode = vbKeyF2 Then '-- Select
    TopPos = txttemp1(Index).Top + Me.Top + 650
    LeftPos = txttemp1(Index).Left + Me.Left
    Select Case Index
        Case 7 '--Narration
            gClsSearch.SearchMultiField "tblMastNarration", "Narration", Array("Narration"), Array(txttemp1(Index).Width), "NarrType='N'", "", "Narration", LeftPos, TopPos, False
            If gClsSearch.SearchMultiRetCol(0) <> "" Then txttemp1(Index) = gClsSearch.SearchMultiRetCol(0)
    End Select
End If
End Sub
Private Sub CalGridAmt()
Dim X As String
Dim ItTp As Long
Dim Stat As String
Dim Camt, Samt, Iamt As Double
Dim IndDbAmt As Double
Dim DbAm As Double
Dim CalGst As Double
Dim IntGst As Double
Dim rstTmp As Recordset
If Val(msGrid.TextMatrix(msGrid.Row, 7)) > 0 Then
   If Val(msGrid.TextMatrix(msGrid.Row, 9)) > 0 Then
      IntGst = Round((Val(msGrid.TextMatrix(msGrid.Row, 7)) * (Val(msGrid.TextMatrix(msGrid.Row, 13)) + Val(msGrid.TextMatrix(msGrid.Row, 14)) + Val(msGrid.TextMatrix(msGrid.Row, 15)))) / 100, 2)
      If Abs(IntGst - Val(msGrid.TextMatrix(msGrid.Row, 9))) < 1 Then
         IndDbAmt = Val(msGrid.TextMatrix(msGrid.Row, 7))
         If Val(msGrid.TextMatrix(msGrid.Row, 13)) + Val(msGrid.TextMatrix(msGrid.Row, 13)) > 0 Then
            Camt = Val(msGrid.TextMatrix(msGrid.Row, 9)) / 2
            Samt = Val(msGrid.TextMatrix(msGrid.Row, 9)) / 2
         Else
            Iamt = Val(msGrid.TextMatrix(msGrid.Row, 9))
         End If
      Else
        IndDbAmt = ((Val(msGrid.TextMatrix(msGrid.Row, 7)) + Val(msGrid.TextMatrix(msGrid.Row, 9))) * 100) / (100 + (Val(msGrid.TextMatrix(msGrid.Row, 13)) + Val(msGrid.TextMatrix(msGrid.Row, 14)) + Val(msGrid.TextMatrix(msGrid.Row, 15))))
        Camt = Round((IndDbAmt * msGrid.TextMatrix(msGrid.Row, 13)) / 100, 0)
        Samt = Round((IndDbAmt * msGrid.TextMatrix(msGrid.Row, 14)) / 100, 0)
        Iamt = Round((IndDbAmt * msGrid.TextMatrix(msGrid.Row, 15)) / 100, 0)
        DbAm = Val(msGrid.TextMatrix(msGrid.Row, 7)) + Val(msGrid.TextMatrix(msGrid.Row, 9))               'IndDbAmt + (Camt + Samt + Iamt)
        DbAm = GProcMakeRounding(DbAm)
        IndDbAmt = DbAm - (Camt + Samt + Iamt)
      End If
   Else
      IndDbAmt = Val(msGrid.TextMatrix(msGrid.Row, 7))
      Camt = Round((IndDbAmt * msGrid.TextMatrix(msGrid.Row, 13)) / 100, 0)
      Samt = Round((IndDbAmt * msGrid.TextMatrix(msGrid.Row, 14)) / 100, 0)
      Iamt = Round((IndDbAmt * msGrid.TextMatrix(msGrid.Row, 15)) / 100, 0)
   End If
   msGrid.TextMatrix(msGrid.Row, 11) = Format(IndDbAmt, FStr)
   msGrid.TextMatrix(msGrid.Row, 12) = Format(Camt + Samt + Iamt, FStr)

'----- TDS-194Q

'ElseIf Val(msGrid.TextMatrix(msGrid.Row, 5)) > 0 Then
'   Camt = Round(Val(msGrid.TextMatrix(msGrid.Row, 5)) * msGrid.TextMatrix(msGrid.Row, 13) / 100, 0)
'   Samt = Round(Val(msGrid.TextMatrix(msGrid.Row, 5)) * msGrid.TextMatrix(msGrid.Row, 14) / 100, 0)
'   Iamt = Round(Val(msGrid.TextMatrix(msGrid.Row, 5)) * msGrid.TextMatrix(msGrid.Row, 15) / 100, 0)
'   msGrid.TextMatrix(msGrid.Row, 10) = Camt + Samt + Iamt

'-----End TDS-194Q


Else
   msGrid.TextMatrix(msGrid.Row, 10) = 0
   msGrid.TextMatrix(msGrid.Row, 11) = 0
   msGrid.TextMatrix(msGrid.Row, 12) = 0
End If
End Sub
Private Sub CalAmount()
Dim Bal As Double
Dim RowIndex As Integer
Bal = Val(txt(6))
RowIndex = 1
With msGrid
    Do While RowIndex < .Rows
'----- TDS-194Q Delete - Val(.TextMatrix(RowIndex, 5))
       
       Bal = Bal - (Val(.TextMatrix(RowIndex, 4)) + Val(.TextMatrix(RowIndex, 7)) - Val(.TextMatrix(RowIndex, 8)) + Val(.TextMatrix(RowIndex, 9)))
'       Bal = Bal - (Val(.TextMatrix(RowIndex, 4)) - Val(.TextMatrix(RowIndex, 5)) + Val(.TextMatrix(RowIndex, 7)) - Val(.TextMatrix(RowIndex, 8)) + Val(.TextMatrix(RowIndex, 9)))
       RowIndex = RowIndex + 1
    Loop
End With
lblBalanceAmt.Caption = Format(Bal, FStr)
End Sub
