VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form frmLatePaymentInt 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00FBF2E1&
   BorderStyle     =   0  'None
   Caption         =   "Late Payment Interest List"
   ClientHeight    =   9405
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   14625
   Icon            =   "frmLatePaymentInt.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   9405
   ScaleWidth      =   14625
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox txt 
      Height          =   325
      Index           =   47
      Left            =   16290
      TabIndex        =   103
      Text            =   "47"
      Top             =   8010
      Visible         =   0   'False
      Width           =   360
   End
   Begin VB.TextBox txt 
      Height          =   325
      Index           =   46
      Left            =   16200
      TabIndex        =   96
      Text            =   "46"
      Top             =   7650
      Visible         =   0   'False
      Width           =   360
   End
   Begin VB.TextBox txt 
      Height          =   325
      Index           =   45
      Left            =   16200
      TabIndex        =   95
      Text            =   "45"
      Top             =   7200
      Visible         =   0   'False
      Width           =   360
   End
   Begin VB.TextBox txt 
      Height          =   325
      Index           =   44
      Left            =   16200
      TabIndex        =   94
      Text            =   "44"
      Top             =   6750
      Visible         =   0   'False
      Width           =   360
   End
   Begin VB.CheckBox chkIsMonthlyDbNt 
      BackColor       =   &H00ECFFFF&
      Caption         =   "Is Monthly Debit Note ?"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C000C0&
      Height          =   240
      Left            =   1845
      MaskColor       =   &H00ECFFFF&
      TabIndex        =   5
      Top             =   1665
      Width           =   4110
   End
   Begin VB.TextBox txt 
      Height          =   325
      Index           =   43
      Left            =   16200
      TabIndex        =   91
      Text            =   "43"
      Top             =   6300
      Visible         =   0   'False
      Width           =   360
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
      Index           =   42
      Left            =   6810
      TabIndex        =   9
      Text            =   "42"
      Top             =   2025
      Width           =   945
   End
   Begin VB.TextBox txt 
      Height          =   325
      Index           =   41
      Left            =   16245
      TabIndex        =   89
      Text            =   "41"
      Top             =   5985
      Visible         =   0   'False
      Width           =   360
   End
   Begin VB.TextBox txt 
      Height          =   325
      Index           =   40
      Left            =   16245
      TabIndex        =   88
      Text            =   "40"
      Top             =   5625
      Visible         =   0   'False
      Width           =   360
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
      Index           =   41
      Left            =   1845
      TabIndex        =   4
      Text            =   "txttemp(41)"
      ToolTipText     =   "Press F3 for Creation"
      Top             =   1215
      Width           =   5910
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
      Index           =   39
      Left            =   3075
      TabIndex        =   7
      Text            =   "39"
      Top             =   2025
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
      Height          =   390
      Index           =   38
      Left            =   4830
      TabIndex        =   8
      Text            =   "38"
      Top             =   2025
      Width           =   945
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
      Index           =   37
      Left            =   1860
      TabIndex        =   6
      Text            =   "37"
      Top             =   2025
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
      Height          =   390
      Index           =   36
      Left            =   16245
      TabIndex        =   10
      Text            =   "36"
      Top             =   5220
      Visible         =   0   'False
      Width           =   405
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
      Index           =   35
      Left            =   1845
      TabIndex        =   1
      Text            =   "35"
      Top             =   765
      Width           =   1125
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
      Height          =   325
      Index           =   34
      Left            =   10035
      TabIndex        =   81
      Text            =   "34"
      Top             =   8865
      Width           =   1440
   End
   Begin VB.TextBox txt 
      Height          =   325
      Index           =   33
      Left            =   16245
      TabIndex        =   80
      Text            =   "33"
      Top             =   4410
      Visible         =   0   'False
      Width           =   360
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
      Height          =   325
      Index           =   32
      Left            =   2250
      TabIndex        =   79
      Text            =   "32"
      Top             =   8865
      Width           =   1620
   End
   Begin VB.TextBox txt 
      Height          =   325
      Index           =   31
      Left            =   16245
      TabIndex        =   78
      Text            =   "31"
      Top             =   3690
      Visible         =   0   'False
      Width           =   360
   End
   Begin VB.TextBox txt 
      Height          =   325
      Index           =   30
      Left            =   16200
      TabIndex        =   77
      Text            =   "30"
      Top             =   3330
      Visible         =   0   'False
      Width           =   360
   End
   Begin VB.TextBox txt 
      Height          =   325
      Index           =   29
      Left            =   16155
      TabIndex        =   76
      Text            =   "29"
      Top             =   2970
      Visible         =   0   'False
      Width           =   360
   End
   Begin VB.TextBox txt 
      Height          =   325
      Index           =   28
      Left            =   16245
      TabIndex        =   75
      Text            =   "28"
      Top             =   2655
      Visible         =   0   'False
      Width           =   360
   End
   Begin VB.TextBox txt 
      Height          =   325
      Index           =   27
      Left            =   16245
      TabIndex        =   74
      Text            =   "27"
      Top             =   2295
      Visible         =   0   'False
      Width           =   360
   End
   Begin VB.TextBox txt 
      Height          =   325
      Index           =   26
      Left            =   16200
      TabIndex        =   73
      Text            =   "26"
      Top             =   1935
      Visible         =   0   'False
      Width           =   360
   End
   Begin VB.TextBox txt 
      Height          =   325
      Index           =   25
      Left            =   16200
      TabIndex        =   72
      Text            =   "25"
      Top             =   1620
      Visible         =   0   'False
      Width           =   360
   End
   Begin VB.TextBox txt 
      Height          =   325
      Index           =   24
      Left            =   16200
      TabIndex        =   71
      Text            =   "24"
      Top             =   1260
      Visible         =   0   'False
      Width           =   360
   End
   Begin VB.TextBox txt 
      Height          =   325
      Index           =   23
      Left            =   16200
      TabIndex        =   70
      Text            =   "23"
      Top             =   900
      Visible         =   0   'False
      Width           =   360
   End
   Begin VB.TextBox txt 
      Height          =   325
      Index           =   22
      Left            =   16200
      TabIndex        =   69
      Text            =   "22"
      Top             =   540
      Visible         =   0   'False
      Width           =   360
   End
   Begin VB.TextBox txt 
      Height          =   325
      Index           =   21
      Left            =   16200
      TabIndex        =   68
      Text            =   "21"
      Top             =   180
      Visible         =   0   'False
      Width           =   360
   End
   Begin VB.TextBox txt 
      Height          =   325
      Index           =   20
      Left            =   18270
      TabIndex        =   67
      Text            =   "20"
      Top             =   4770
      Visible         =   0   'False
      Width           =   360
   End
   Begin VB.TextBox txt 
      Height          =   325
      Index           =   19
      Left            =   18225
      TabIndex        =   66
      Text            =   "19"
      Top             =   4410
      Visible         =   0   'False
      Width           =   360
   End
   Begin VB.TextBox txt 
      Height          =   325
      Index           =   18
      Left            =   18225
      TabIndex        =   65
      Text            =   "18"
      Top             =   4050
      Visible         =   0   'False
      Width           =   360
   End
   Begin VB.TextBox txt 
      Height          =   325
      Index           =   17
      Left            =   18180
      TabIndex        =   64
      Text            =   "17"
      Top             =   3735
      Visible         =   0   'False
      Width           =   360
   End
   Begin VB.TextBox txt 
      Height          =   325
      Index           =   12
      Left            =   18135
      TabIndex        =   63
      Text            =   "12"
      Top             =   3375
      Visible         =   0   'False
      Width           =   360
   End
   Begin VB.TextBox txt 
      Height          =   325
      Index           =   11
      Left            =   18135
      TabIndex        =   62
      Text            =   "11"
      Top             =   3015
      Visible         =   0   'False
      Width           =   360
   End
   Begin VB.TextBox txt 
      Height          =   325
      Index           =   9
      Left            =   18180
      TabIndex        =   61
      Text            =   "9"
      Top             =   2655
      Visible         =   0   'False
      Width           =   330
   End
   Begin VB.TextBox txt 
      Alignment       =   1  'Right Justify
      Height          =   360
      Index           =   6
      Left            =   18180
      TabIndex        =   60
      Text            =   "6"
      Top             =   1935
      Visible         =   0   'False
      Width           =   345
   End
   Begin VB.TextBox txt 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   375
      Index           =   16
      Left            =   17550
      MaxLength       =   50
      TabIndex        =   19
      Text            =   "16"
      ToolTipText     =   "Press F3 for Creation & Press F2 for Selection"
      Top             =   1800
      Visible         =   0   'False
      Width           =   315
   End
   Begin VB.TextBox txt 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   375
      Index           =   15
      Left            =   17550
      MaxLength       =   50
      TabIndex        =   18
      Text            =   "15"
      ToolTipText     =   "Press F3 for Creation & Press F2 for Selection"
      Top             =   1395
      Visible         =   0   'False
      Width           =   360
   End
   Begin VB.CheckBox chkDontShowInUnadj 
      BackColor       =   &H00FFC0FF&
      Caption         =   "Dont Show In Unadjusted Receipt"
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
      Left            =   6345
      TabIndex        =   59
      Top             =   9540
      Width           =   4515
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
      Left            =   3420
      TabIndex        =   50
      Top             =   4095
      Visible         =   0   'False
      Width           =   6030
      Begin VB.OptionButton OptType 
         BackColor       =   &H00FFC0C0&
         Caption         =   "Int. Debit Note "
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
         Height          =   330
         Index           =   1
         Left            =   3915
         TabIndex        =   93
         Top             =   1440
         Width           =   1815
      End
      Begin VB.OptionButton OptType 
         BackColor       =   &H00FFC0C0&
         Caption         =   "Int. Debit Note Check List"
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
         Height          =   330
         Index           =   0
         Left            =   495
         TabIndex        =   92
         Top             =   1440
         Width           =   2985
      End
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
         Left            =   5085
         TabIndex        =   58
         Top             =   2340
         Visible         =   0   'False
         Width           =   1335
      End
      Begin VB.TextBox txtno 
         Appearance      =   0  'Flat
         Height          =   360
         Index           =   0
         Left            =   1950
         MaxLength       =   7
         TabIndex        =   55
         Top             =   420
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
         Left            =   270
         TabIndex        =   54
         Top             =   2145
         Width           =   1335
      End
      Begin VB.TextBox txtno 
         Appearance      =   0  'Flat
         Height          =   360
         Index           =   1
         Left            =   1950
         MaxLength       =   7
         TabIndex        =   53
         Top             =   900
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
         TabIndex        =   52
         Top             =   2145
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
         TabIndex        =   51
         Top             =   2145
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
         TabIndex        =   57
         Top             =   420
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
         TabIndex        =   56
         Top             =   900
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
      Left            =   11880
      Style           =   1  'Graphical
      TabIndex        =   49
      ToolTipText     =   "Exit"
      Top             =   1710
      Width           =   1065
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
      Left            =   10710
      Style           =   1  'Graphical
      TabIndex        =   48
      ToolTipText     =   "To Cancel Record"
      Top             =   1215
      Width           =   1065
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
      Left            =   10710
      Style           =   1  'Graphical
      TabIndex        =   47
      ToolTipText     =   "To Save Record"
      Top             =   720
      Width           =   1065
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
      Left            =   10710
      Style           =   1  'Graphical
      TabIndex        =   46
      Top             =   1710
      Width           =   1065
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
      Left            =   9585
      Style           =   1  'Graphical
      TabIndex        =   45
      ToolTipText     =   "To Delete Record"
      Top             =   1710
      Width           =   1065
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
      Left            =   9585
      Style           =   1  'Graphical
      TabIndex        =   44
      ToolTipText     =   "To Modify Record"
      Top             =   1215
      Width           =   1065
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
      Left            =   9585
      Style           =   1  'Graphical
      TabIndex        =   43
      ToolTipText     =   "To Add New Record"
      Top             =   720
      Width           =   1065
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
      Left            =   13005
      Style           =   1  'Graphical
      TabIndex        =   42
      ToolTipText     =   "To Find Record"
      Top             =   720
      Width           =   1065
   End
   Begin VB.CommandButton cmdIntDbNt 
      Caption         =   "&Interest Debit Note"
      Height          =   900
      Left            =   13005
      TabIndex        =   41
      Top             =   1215
      Width           =   1065
   End
   Begin VB.TextBox txt 
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Index           =   10
      Left            =   17505
      TabIndex        =   15
      Text            =   "10"
      Top             =   2565
      Visible         =   0   'False
      Width           =   405
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
      Left            =   18630
      TabIndex        =   39
      Text            =   "txtCrBalance"
      Top             =   180
      Visible         =   0   'False
      Width           =   615
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
      Left            =   18630
      TabIndex        =   38
      Text            =   "txtDrBalance"
      Top             =   585
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.CommandButton cmdLedger 
      Caption         =   "Ledger"
      Height          =   375
      Index           =   0
      Left            =   18225
      TabIndex        =   37
      Top             =   5175
      Visible         =   0   'False
      Width           =   1200
   End
   Begin VB.CommandButton cmdLedger 
      Caption         =   "Ledger"
      Height          =   375
      Index           =   1
      Left            =   18225
      TabIndex        =   36
      Top             =   5535
      Visible         =   0   'False
      Width           =   1185
   End
   Begin VB.TextBox txtGrid 
      BackColor       =   &H8000000A&
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
      Left            =   480
      TabIndex        =   12
      Top             =   4815
      Visible         =   0   'False
      Width           =   885
   End
   Begin VB.TextBox txt 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   375
      Index           =   13
      Left            =   17550
      MaxLength       =   50
      TabIndex        =   16
      Text            =   "13"
      ToolTipText     =   "Press F3 for Creation & Press F2 for Selection"
      Top             =   585
      Visible         =   0   'False
      Width           =   360
   End
   Begin VB.TextBox txt 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFFFF&
      Enabled         =   0   'False
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
      Left            =   17550
      TabIndex        =   13
      Text            =   "0"
      Top             =   135
      Visible         =   0   'False
      Width           =   285
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
      Left            =   11880
      Style           =   1  'Graphical
      TabIndex        =   21
      ToolTipText     =   "Previous"
      Top             =   720
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
      Left            =   12450
      Style           =   1  'Graphical
      TabIndex        =   22
      ToolTipText     =   "Next"
      Top             =   720
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
      Left            =   11880
      Style           =   1  'Graphical
      TabIndex        =   23
      ToolTipText     =   "First"
      Top             =   1215
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
      Left            =   12450
      Style           =   1  'Graphical
      TabIndex        =   24
      ToolTipText     =   "Last"
      Top             =   1215
      Width           =   495
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
      Left            =   18360
      Style           =   1  'Graphical
      TabIndex        =   25
      Top             =   6480
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
      Left            =   18585
      Style           =   1  'Graphical
      TabIndex        =   29
      Top             =   8325
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
      Left            =   18585
      Style           =   1  'Graphical
      TabIndex        =   30
      Top             =   7920
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.TextBox txt 
      Height          =   345
      Index           =   3
      Left            =   18135
      TabIndex        =   33
      Text            =   "3"
      Top             =   900
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.TextBox txt 
      Alignment       =   1  'Right Justify
      Height          =   360
      Index           =   5
      Left            =   18135
      TabIndex        =   28
      Text            =   "5"
      Top             =   1575
      Visible         =   0   'False
      Width           =   345
   End
   Begin VB.TextBox txt 
      Alignment       =   1  'Right Justify
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Index           =   8
      Left            =   17505
      TabIndex        =   14
      Text            =   "8"
      Top             =   2160
      Visible         =   0   'False
      Width           =   330
   End
   Begin VB.TextBox txt 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   375
      Index           =   14
      Left            =   17550
      MaxLength       =   50
      TabIndex        =   17
      Text            =   "14"
      ToolTipText     =   "Press F3 for Creation & Press F2 for Selection"
      Top             =   990
      Visible         =   0   'False
      Width           =   405
   End
   Begin VB.TextBox txt 
      Height          =   325
      Index           =   7
      Left            =   18180
      TabIndex        =   26
      Text            =   "7"
      Top             =   2340
      Visible         =   0   'False
      Width           =   330
   End
   Begin VB.TextBox txt 
      Alignment       =   1  'Right Justify
      Height          =   360
      Index           =   4
      Left            =   18135
      TabIndex        =   27
      Text            =   "4"
      Top             =   1260
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.TextBox txt 
      Height          =   345
      Index           =   2
      Left            =   18135
      TabIndex        =   32
      Text            =   "2"
      Top             =   540
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
      Left            =   18135
      TabIndex        =   31
      Text            =   "1"
      Top             =   180
      Visible         =   0   'False
      Width           =   375
   End
   Begin MSFlexGridLib.MSFlexGrid msGrid 
      Height          =   6060
      Left            =   210
      TabIndex        =   11
      Top             =   2745
      Width           =   14145
      _ExtentX        =   24950
      _ExtentY        =   10689
      _Version        =   393216
      FixedCols       =   0
      RowHeightMin    =   275
      ForeColor       =   0
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
   Begin Crystal.CrystalReport cryReport 
      Bindings        =   "frmLatePaymentInt.frx":058A
      Left            =   -45
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
   Begin MSComCtl2.DTPicker dtpDt 
      Height          =   360
      Index           =   2
      Left            =   6075
      TabIndex        =   3
      Top             =   780
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
      Format          =   81264643
      CurrentDate     =   -273
   End
   Begin MSComCtl2.DTPicker dtpDt 
      Height          =   360
      Index           =   40
      Left            =   3825
      TabIndex        =   2
      Top             =   780
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
      Format          =   81264643
      CurrentDate     =   -273
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "+"
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
      Left            =   6120
      TabIndex        =   102
      Top             =   8910
      Width           =   120
   End
   Begin VB.Label lblInt2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Interest Amount 2"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C000C0&
      Height          =   240
      Left            =   6525
      TabIndex        =   101
      Top             =   8910
      Width           =   1665
   End
   Begin VB.Label lblTDS 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "T D S  Amount"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C000C0&
      Height          =   240
      Left            =   11565
      TabIndex        =   100
      Top             =   8910
      Width           =   2535
   End
   Begin VB.Label lblInt 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Interest Amount"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C000C0&
      Height          =   240
      Left            =   4005
      TabIndex        =   99
      Top             =   8910
      Width           =   1905
   End
   Begin VB.Label lbl1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "TDS JV NO   :"
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
      Index           =   9
      Left            =   8505
      TabIndex        =   98
      Top             =   8910
      Width           =   1275
   End
   Begin VB.Label lbl1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "INTEREST JV NO   :"
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
      Index           =   8
      Left            =   270
      TabIndex        =   97
      Top             =   8910
      Width           =   1830
   End
   Begin VB.Label lbl1 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "TDS @  :"
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
      Index           =   7
      Left            =   5895
      TabIndex        =   90
      Top             =   2070
      Width           =   810
   End
   Begin VB.Label lbl1 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "+"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   285
      Index           =   6
      Left            =   2790
      TabIndex        =   87
      Top             =   2055
      Width           =   135
   End
   Begin VB.Label lbl1 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Party  :"
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
      Index           =   5
      Left            =   945
      TabIndex        =   86
      Top             =   1260
      Width           =   675
   End
   Begin VB.Label lbl1 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "To  :"
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
      Index           =   4
      Left            =   5580
      TabIndex        =   85
      Top             =   840
      Width           =   420
   End
   Begin VB.Label lblApplyJV 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Interest JV Pending"
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
      Height          =   990
      Left            =   7875
      TabIndex        =   84
      Top             =   1305
      Width           =   1155
      WordWrap        =   -1  'True
   End
   Begin VB.Label lbl1 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Int. @  :"
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
      Index           =   3
      Left            =   4005
      TabIndex        =   83
      Top             =   2070
      Width           =   705
   End
   Begin VB.Label lbl1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Grace Days  :"
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
      Left            =   390
      TabIndex        =   82
      Top             =   2070
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
      Left            =   17055
      TabIndex        =   40
      Top             =   5265
      Width           =   1095
   End
   Begin VB.Shape Shape6 
      BorderColor     =   &H00EFB17A&
      BorderWidth     =   5
      FillColor       =   &H00EFB17A&
      Height          =   9330
      Left            =   0
      Top             =   0
      Width           =   14550
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
      Left            =   10575
      TabIndex        =   35
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
      Left            =   630
      TabIndex        =   0
      Top             =   765
      Width           =   1035
   End
   Begin VB.Label lbl1 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "From  :"
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
      Left            =   2985
      TabIndex        =   20
      Top             =   840
      Width           =   675
   End
   Begin VB.Shape Shape1 
      FillColor       =   &H00BDDFD1&
      FillStyle       =   0  'Solid
      Height          =   1830
      Left            =   9315
      Shape           =   4  'Rounded Rectangle
      Top             =   540
      Width           =   4950
   End
   Begin VB.Label Label6 
      BackColor       =   &H00EFB17A&
      Caption         =   "   Late Payment Interest List"
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
      TabIndex        =   34
      Top             =   30
      Width           =   14475
   End
   Begin VB.Shape Shape2 
      BackColor       =   &H00808080&
      BorderColor     =   &H00000080&
      FillColor       =   &H00ECFFFF&
      FillStyle       =   0  'Solid
      Height          =   1980
      Left            =   225
      Top             =   585
      Width           =   8970
   End
End
Attribute VB_Name = "frmLatePaymentInt"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Const tblName As String = "tblVoucher"
Const IdField As String = "VnoList"
Const MaxNo As Long = 47
Dim CriteriaStr As String
Dim rstbl As Recordset, ShowFrmFlag As Boolean
Dim rsRecVsSale As Recordset
Dim varBookMark As Variant
Dim FirstTimeFlag As Boolean, mLastEntryBank As String, mLastEntryDate As Date
Dim FormAction As Integer
Dim frmTypeStr As String
Dim i As Long
Dim OldTdsOn As Double
Dim OldgFrmTypeStr As String
Dim MainFirm As String
Dim CriteriaStrForGenId As String
Const OrderField As String = "Vdt,VnoList"
Private Sub cmdBtn_Click(Index As Integer)
Dim FindQStr As String
Dim X As String
Dim Cdt As String
'On Error GoTo ErrorRoutine
Select Case Index
Case 0 'Previous
    If GProcRstOpen(rstbl, "Select * from " & tblName & " where VnoList = (Select max(VnoList) from " & tblName & " where VnoList <> 0  and VnoList < " & txt(35) & " and  " & CriteriaStr & "  ) and  " & CriteriaStr & "  ", "O") > 0 Then
       ReadFields
    Else
      MsgBox "This is First Record ", vbOKOnly + vbInformation
    End If
Case 1 'Next
    If GProcRstOpen(rstbl, "Select * from " & tblName & " where VnoList = (Select min(VnoList) from " & tblName & " where VnoList <> 0  and VnoList > " & txt(35) & "  and " & CriteriaStr & ") and  " & CriteriaStr & "   ", "O") > 0 Then
       ReadFields
    Else
       MsgBox "This is Last Record ", vbOKOnly + vbInformation
    End If
Case 2 'First
    If GProcRstOpen(rstbl, "Select * from " & tblName & " where VnoList = (Select min(VnoList) from " & tblName & " where VnoList <> 0 and " & CriteriaStr & ") and  " & CriteriaStr & "   ", "O") > 0 Then ReadFields
Case 3 'Last
     If GProcRstOpen(rstbl, "Select * from " & tblName & " where VnoList = (Select max(VnoList) from " & tblName & " where VnoList <> 0 and " & CriteriaStr & " )and  " & CriteriaStr & "   ", "O") > 0 Then ReadFields
Case 4 'Find
    gClsSearch.SearchMultiField "tblVoucher", "VnoList,Vdt,sum(Interest)", Array("Vno", "Vdt", ""), Array(1400, 1300, 0), CriteriaStr + " group by VnoList,Vdt", , , , , True
    If gClsSearch.SearchMultiRetCol(0) <> "" Then
        If GProcRstOpen(rstbl, "Select * from " & tblName & " where VnoList = (Select max(" & gClsSearch.SearchMultiRetCol(0) & ") from " & tblName & " where VnoList <> 0 and " & CriteriaStr & " ) and  " & CriteriaStr & "   ", "O") > 0 Then ReadFields
    End If
Case 5 'List
    'GenMastList ("LAC")
Case 6 'Add
    FormAction = vbDataActionAddNew
    Call GProcActivateControls(Me, True)
    Call GProcSetButtons(Me, vbDataActionAddNew)
    Call GProcClearForm(Me, rstbl, MaxNo, True)
    txt(35).text = GProcGenerateId(tblName, "VnoList", CriteriaStrForGenId, gCn)
    txt(37) = gLPGrase
    txt(39) = gLPExtraGrase
    txt(38) = Format(gLPIntRt, FStr)
    txt(42) = Format(gTDSRate, FStr)
    Cdt = Format(Now, "dd/MM/yyyy")
    If CDate(Cdt) > gCYEDate Then
       dtpDt(2) = Format(gCYEDate, "dd/MM/yyyy")
       dtpDt(40) = Format(gCYEDate, "dd/MM/yyyy")
    Else
       dtpDt(2) = Format(Now, "dd/MM/yyyy")
       dtpDt(40) = Format(Now, "dd/MM/yyyy")
    End If
    If FirstTimeFlag = True Then
        FirstTimeFlag = False
    Else
        txt(35).SetFocus
    End If
'    FillCombo
    SetGrid
Case 7 'Modify
    'If mURecModi = True Then
    '    SearchHelp
    '    DoEvents
    If IsRelatedRecord = False Then
        FormAction = vbDataActionUpdate
        Call GProcActivateControls(Me, True)
        Call GProcSetButtons(Me, vbDataActionUpdate)
        txt(35).SetFocus
        SetControlEd
    End If
Case 8 'Delete
   ' If mURecDel = True Then
     If IsRelatedRecord = False Then
        If MsgBox("Do You Really Delete This Record?", vbQuestion + vbYesNo + vbDefaultButton1, "Confirm Delete") = vbYes Then
'           gCn.Execute "Delete from tblvoucher where " & CriteriaStr & " and vno = " & txt(0)
            '-- Delete Related Record
            DeleteRelatedRecord
            cmdBtn_Click (0)
            If GProcRstOpen(rstbl, "Select * from " & tblName & " where VnoList = (Select max(VnoList) from " & tblName & " where VnoList <> 0 and " & CriteriaStr & " )and  " & CriteriaStr & "   ", "O") > 0 Then
               ReadFields
            Else
                SetGrid
                Call GProcClearForm(Me, rstbl, MaxNo, True)
            End If
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
    OptType(0).Enabled = True
    OptType(1).Enabled = True
    OptType(0).Value = True
    txtno(0) = txt(35)
    txtno(1) = txt(35)
    txtno(0).SetFocus
Case 10 'Exit
    DoEvents
    Unload Me
Case 13 'Save
       If txtGrid.Visible = True Then
'       If txtGrid.Enabled = True Then
            msGrid.text = txtGrid
            DoEvents
            txtGrid = ""
        End If
        txtGrid.Visible = False
    If ValidateData = True Then
       If txtGrid.Visible = True Then
'       If txtGrid.Enabled = True Then
            msGrid.text = txtGrid
            DoEvents
            txtGrid = ""
        End If
        txtGrid.Visible = False
        FillTxtFromTemp
        SaveRelatedRecord
        DoEvents
        DoEvents
        txtGrid = ""
        If GProcRstOpen(rstbl, "Select * from " & tblName & " where VnoList = " & txt(35) & " and  " & CriteriaStr & "   ", "O") > 0 Then ReadFields
'        FillTempFromTxt
'        ReadFields
        rstbl.Requery
        FormAction = vbDataActionClose
        Call GProcActivateControls(Me, False)
        Call GProcSetButtons(Me, True)
        msGrid.Col = 0
        msGrid.Row = 1
        txtGrid.Move msGrid.CellLeft + msGrid.Left, msGrid.CellTop + msGrid.Top, msGrid.CellWidth, msGrid.CellHeight
        lblBalance.Caption = ""
        If FormAction = vbDataActionAddNew Then
            If cmdBtn(6).Visible = True Then cmdBtn(6).SetFocus
        Else
            If cmdBtn(7).Visible = True Then cmdBtn(7).SetFocus
        End If
    End If
Case 14 'Cancel
    txtGrid.Visible = False
    If GProcRstOpen(rstbl, "Select * from " & tblName & " where VnoList = (Select max(VnoList) from " & tblName & " where VnoList <> 0 and " & CriteriaStr & " ) and  " & CriteriaStr & "   ", "O") > 0 Then ReadFields
    
    
    If rstbl.RecordCount <= 0 Then
        Call GProcClearForm(Me, rstbl, MaxNo, True)
        Call GProcActivateControls(Me, False)
        Call GProcSetButtons(Me, vbDataActionCancel)
    Else
        If FormAction = vbDataActionAddNew Then
            rstbl.MoveLast
        Else
            rstbl.CancelUpdate
        End If
        txtGrid = ""
        ReadFields
        Call GProcActivateControls(Me, False)
        Call GProcSetButtons(Me, True)
    End If
    FormAction = vbDataActionClose
    msGrid.Col = 0
    If msGrid.Row >= 1 Then msGrid.Row = 1
    txtGrid.Move msGrid.CellLeft + msGrid.Left, msGrid.CellTop + msGrid.Top, msGrid.CellWidth, msGrid.CellHeight
    lblBalance.Caption = ""
    If FormAction = vbDataActionAddNew Then
        If cmdBtn(6).Visible = True Then cmdBtn(6).SetFocus
    Else
        If cmdBtn(7).Visible = True And cmdBtn(7).Enabled = True Then
           cmdBtn(7).SetFocus
        Else
           cmdBtn(10).SetFocus
        End If
    End If
End Select
Exit Sub
ErrorRoutine:
    rstbl.CancelUpdate
    GProcErrorHandler
End Sub

Private Sub dtpDt_GotFocus(Index As Integer)
GProcSelectBox Me.ActiveControl
End Sub

Private Sub dtpDt_LostFocus(Index As Integer)
SaveBtnEd
End Sub

Private Sub txtGrid_DblClick()
If msGrid.Col = 13 Then
   If msGrid.TextMatrix(msGrid.Row, 4) = msGrid.TextMatrix(msGrid.Row - 1, 4) Then
      If txtGrid.text = "" Then
         txtGrid.text = msGrid.TextMatrix(msGrid.Row - 1, 13)  '  IIf(msGrid.TextMatrix(msGrid.Row, 13) = "", msGrid.TextMatrix(msGrid.Row - 1, 0), msGrid.TextMatrix(msGrid.Row - 1, 13))
      Else
         txtGrid.text = ""
      End If
   Else
      If txtGrid.text = "" Then
          txtGrid.text = msGrid.TextMatrix(msGrid.Row, 0)
      Else
         txtGrid.text = ""
      End If
   End If
End If
If msGrid.Col = 12 Then
   If txtGrid.text = "Y" Then
      txtGrid.text = "N"
   Else
      txtGrid.text = "Y"
   End If
End If
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
        X = "select VBillNo,VBillType, VBillDate,VBillAmt,tblMastAccount.AcOurCode,tblMastAccount.AcName,tblMastAccount.AcCity,tblMastAccount_Br.AcOurCode,tblMastAccount_Br.AcName, VAmt,DueOnDt,tblVoucher.Vdt,LateDays,Interest,JvNoInt,Grace+GraceExtra,IntRt,DbNtMonthly,DbNtInBillNo,'" & gUserName & "'" _
            & " from tblVoucher,tblMastAccount,tblMastAccount tblMastAccount_Br,tblSale where " _
            & " tblVoucher.VType='" & frmTypeStr & "' and VnoList=" & Vno & " and tblVoucher.VFirm  = '" & gCCode & "' and  tblVoucher.VYear=" & gCYear & " and VCrAcCode = tblMastAccount.AcCode and  tblSale.SlBroker = tblMastAccount_Br.AcCode  and " _
            & " VBillVno = tblSale.Vno and VBillType = tblSale.Vtype and VBillVYear = tblSale.Vyear and tblVoucher.VFirm = tblSale.Vfirm   order by tblMastAccount.AcName,tblMastAccount_Br.AcName,tblVoucher.Vno,VCtrNo"
               
            gCn.Execute "Delete from TmpGentbl where  UserName='" & gUserName & "'"
     
            gCn.Execute " insert into TmpGentbl (VBillNo,VBillType,VBillDt,Amt1,PartyCode,PartyName,Nar1,BrokerCode,BrokerName,Amt2,Dt1,Vdt,Amt3,Amt4,Amt5,Amt6,amt7,amt8,Nar2,UserName) " & X
'            X1 = "{TmpGentbl.UserName}='" & gUserName & "' and {@IntAmt} > 0  "
            With frmMain.cryReport
              .SelectionFormula = "{TmpGentbl.UserName}='" & gUserName & "'"
              .Formulas(5) = "wIntRt =  '" & txt(1) & "'"
              .Formulas(6) = "wGrase =  '" & txt(2) & "'"
              If OptType(0).Value = True Then
                 .ReportFileName = gReportPath & "rptIntDbNtList.rpt"
              Else
                 .ReportFileName = gReportPath & "rptIntDbNt.rpt"
              End If
              .Action = 1
              .PageZoom (120)

            End With
 '       End If
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
'   DeleteRelatedRecord
   ApplyIntDbNt
End Sub
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case 13:
            If Me.ActiveControl.Name <> "txtGrid" Then Sendkeys "{TAB}"
        Case 27:
'            If (cmdBtn(13).Enabled = False) Then
'                Unload Me
'            End If
    End Select
End Sub
Private Sub Form_Load()
    Dim X As String
    Dim Y As String
    Me.Top = 300
    Me.Left = Screen.Width / 2 - Me.Width / 2
    frmTypeStr = gfrmTypeStr
    X = GProcGetColumnValue("tblMastCompany", "CCode", gCCode, "S", "CDepotMainFirmCompCd", "S")
    If X = "" Then
       MainFirm = gCCode
    Else
       MainFirm = X
    End If
    CriteriaStr = " VType = '" & frmTypeStr & "' and  VnoList <> 0  and Interest > 0 and VFirm  = '" & gCCode & "' and VYear=" & gCYear
    CriteriaStrForGenId = " VType = '" & frmTypeStr & "' and  VnoList <> 0 and VFirm  = '" & gCCode & "' and VYear=" & gCYear
    Y = "Select * from " & tblName & " where " & CriteriaStr & " "
    If Trim(OrderField) <> "" Then Y = Y & " order by " & OrderField
    GProcRstOpen rstbl, Y, "O"
    FirstTimeFlag = True
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
Private Sub txt_GotFocus(Index As Integer)
GProcSelectBox Me.ActiveControl
End Sub
'Private Sub txt_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
'Dim OldFrmType As String
'Dim LeftPos As Long
'Dim TopPos As Long
'If KeyCode = vbKeyF3 Then '--F3 Create
'    Select Case Index
'        Case 13, 14, 15, 16 '--Narration
'            OldFrmType = gfrmTypeStr
'            gfrmTypeStr = "N"
'            GProcShowForm frmMastNarration, frmMain.mnuMstfrm(5), , gfrmTypeStr
'            gfrmTypeStr = OldFrmType
'    End Select
'End If
'If KeyCode = vbKeyF2 Then '-- Select
'    TopPos = txt(Index).Top + Me.Top + 650
'    LeftPos = txt(Index).Left + Me.Left
'    Select Case Index
'        Case 13, 14, 15, 16 '--Narration
'            gClsSearch.SearchMultiField "tblMastNarration", "Narration", Array("Narration"), Array(txt(Index).Width), "NarrType='N'", "", "Narration", LeftPos, TopPos, False
'            If gClsSearch.SearchMultiRetCol(0) <> "" Then txt(Index) = gClsSearch.SearchMultiRetCol(0)
'    End Select
'End If
'End Sub
Private Sub txt_LostFocus(Index As Integer)
GProcCheckForNumber Me, rstbl, MaxNo
If FormAction = vbDataActionAddNew And Index = 42 Then
   PreperReport
End If
If FormAction = vbDataActionUpdate And Index = 35 Then
   msGrid.Col = 9
   msGrid.SetFocus
   msGrid_EnterCell
End If


SaveBtnEd
End Sub
Private Sub txt_KeyPress(Index As Integer, KeyAscii As Integer)
    KeyAscii = GProcValidateKey(rstbl, Index, KeyAscii, txt(Index).text)
End Sub
Private Sub SaveBtnEd()
'--Vno,Vdt,Amount,Cr A/c,Dr A/c
If txt(35) = "" Or GProcIsDateValid(dtpDt(2)) = False Then
    cmdBtn(13).Enabled = False
    cmdBtn(13).BackColor = &H80000004
Else
    cmdBtn(13).Enabled = True
    cmdBtn(13).BackColor = &HFFFF80
End If
End Sub
Private Sub FillTxtFromTemp()
txt(1) = frmTypeStr
txt(2) = IIf(IsDate(dtpDt(2)), dtpDt(2), "") '--Vdt
txt(40) = IIf(IsDate(dtpDt(40)), dtpDt(40), "") '--Vdt
txt(41) = GProcGetColumnValue("tblMastAccount", "AcName", txttemp(41), "S", "AcCode", "N")
txt(3) = gCYear
txt(4) = gCCode
ShowAccountBalance
End Sub
Private Sub FillTempFromTxt()
dtpDt(2) = IIf(IsDate(txt(2)), CDate(txt(2)), "__/__/____") '--Vdt
dtpDt(40) = IIf(IsDate(txt(40)), CDate(txt(40)), "__/__/____") '--Int From Vdt
txttemp(41) = GProcGetColumnValue("tblMastAccount", "AcCode", txt(41), "N", "AcName", "S")
If Val(txt(32)) > 0 Then
   lblApplyJV.Visible = False
   cmdIntDbNt.Visible = False
'   cmdBtn(7).Visible = False
'   cmdBtn(8).Visible = False
Else
   lblApplyJV.Visible = True
   cmdIntDbNt.Visible = True
   cmdBtn(7).Visible = True
'   cmdBtn(8).Visible = True
End If
'-- Fill Related Record Rec Vs Sale
FillRelatedRecord
'ShowAccountBalance
End Sub
Private Function ValidateGrid() As Boolean
'--- Pending
Dim TotAmt As Double
Dim i As Long
Dim RCount As Long
Dim J As Long
Dim ChkTds As Double
Dim BilChk As Boolean
Dim PartyChk As Boolean
'If Trim(txtGrid.text) <> "" Then
'   msGrid.text = txtGrid.text
'End If
'txtGrid.Visible = False
'txtGrid.text = ""
CheckForNumberGrid
With msGrid
     For i = 1 To .rows - 1
         If Val(.TextMatrix(i, 11)) > 0 And Val(.TextMatrix(i, 9)) = 0 Then
                MsgBox "Please Check Interest Amount. ", vbOKOnly, Me.Caption
                msGrid.SetFocus
                Exit Function
         End If
         If Val(.TextMatrix(i, 11)) > 0 Then
            ChkTds = GProcMakeRounding((Val(.TextMatrix(i, 10)) * Val(txt(42))) / 100)
            If Val(.TextMatrix(i, 11)) <> ChkTds Then
                MsgBox "Please Check TDS Amount. ", vbOKOnly, Me.Caption
                .Col = 10
                .SetFocus
                Exit Function
            End If
         End If
         BilChk = False
         PartyChk = False
         If Trim(.TextMatrix(i, 13)) <> "" Then
             For J = 1 To .rows - 1
                 If .TextMatrix(J, 0) = .TextMatrix(i, 13) Then
                    BilChk = True
                    If .TextMatrix(J, 4) = .TextMatrix(i, 4) Then
                       PartyChk = True
                    End If
                 End If
                 If BilChk = True And PartyChk = True Then
                    Exit For
                 End If
             Next J
             If BilChk = False Then
                MsgBox "Please Check Bill No. ", vbOKOnly, Me.Caption
               .Row = i
               .Col = 13
               .SetFocus
               Exit Function
             End If
             If PartyChk = False Then
                MsgBox "Please Check Bill No. And Party Name ", vbOKOnly, Me.Caption
               .Row = i
               .Col = 13
               .SetFocus
               Exit Function
             End If
         End If
         
         
     Next i
End With



'With msGrid
'    '-- Duplicate invoice selection
'    For i = 1 To .rows - 1
'        If Trim(.TextMatrix(i, 1)) <> "" Then
'           BillDet = BillDet + .TextMatrix(i, 1) + "-" + CStr(Val(.TextMatrix(i, 4)) + Val(.TextMatrix(i, 6)) + Val(.TextMatrix(i, 9))) + ":"
'        End If
'
'        '-- Int A/c
'           If CDbl(.TextMatrix(i, 9)) > 0 Then
'                If gLatePayIntAcCode = GProcGetColumnValue("tblmastAccount", "AcName", "-", "S", "AcCode", "N") Then
'                    IntAcCode = 0
'                Else
'                    IntAcCode = gLatePayIntAcCode
'                End If
'                If IntAcCode = 0 Then
'                    MsgBox "Check Late Pay Interest Account.", vbInformation + vbOKOnly, Me.Caption
'                    GProcShowForm frmSetting, frmMain.mnuMstSetting(0)
'                    Exit Function
'                End If
'           End If
'        '---
'        '-- TDS Paid A/c ( For Less TDS Amt JV )
'           If CDbl(.TextMatrix(i, 13)) > 0 Then
'                If gTDSPaidAcCode = GProcGetColumnValue("tblmastAccount", "AcName", "-", "S", "AcCode", "N") Then
'                    TDSPaidAcCode = 0
'                Else
'                    TDSPaidAcCode = gTDSPaidAcCode
'                End If
'                If TDSPaidAcCode = 0 Then
'                    MsgBox "Check TDS Paid Account.", vbInformation + vbOKOnly, Me.Caption
'                    GProcShowForm frmSetting, frmMain.mnuMstSetting(0)
'                    Exit Function
'                End If
'           End If
'        '---
'        InvNo = .TextMatrix(i, 11)
'        InVtype = .TextMatrix(i, 0)
'        InvYear = .TextMatrix(i, 10)
'        For j = 1 To .rows - 1
'            If j <> i And InvNo = .TextMatrix(j, 11) And InVtype = .TextMatrix(j, 0) And InvYear = .TextMatrix(j, 10) And InvNo <> 0 Then
'                MsgBox "Duplicate Invoice Selection.", vbOKOnly, Me.Caption
'                msGrid.SetFocus
'                Exit Function
'            End If
'        Next j
'    Next i
'    If BillDet <> "" Then
'        BillDet = Left(Trim(BillDet) + Space(120), 120)
'        txt(8) = Left(BillDet, 40)
'        txt(9) = Mid(BillDet, 41, 40)
'        txt(10) = Mid(BillDet, 81, 40)
'    End If
'    '-- Adjusted Amt
'    If .rows > 1 Then
'       TotAmt = 0
'       For i = 1 To .rows - 1
'           '-- Adj amt > Bill Amt
'           If CDbl(.TextMatrix(i, 4)) - CDbl(.TextMatrix(i, 5)) > CDbl(.TextMatrix(i, 3)) Then
'                MsgBox "Adjusted Amount is more than Balance Bill Amount. ", vbOKOnly, Me.Caption
'                msGrid.SetFocus
'                Exit Function
'           End If
'            '-- Adj DB amt > Db Amt
''           If CDbl(.TextMatrix(i, 6)) > CDbl(.TextMatrix(i, 5)) Then
''                MsgBox "Adjusted DbNt Amount is more than Balance DbNt Amount. ", vbOKOnly, Me.Caption
''                msGrid.SetFocus
''                Exit Function
''           End If
'           TotAmt = TotAmt + Val(.TextMatrix(i, 4)) - Val(.TextMatrix(i, 5)) + Val(.TextMatrix(i, 9)) - Val(.TextMatrix(i, 13))
'        Next i
'        If (TotAmt > 0) And (TotAmt <> Val(txt(6))) Then
'           MsgBox "Total of Adjusted Amount " & TotAmt & " & Received Amount not tally.", vbOKOnly, Me.Caption
'     '      msGrid.SetFocus
'     '      Exit Function
'        End If
'    End If
'End With
ValidateGrid = True
End Function
Private Function ValidateData() As Boolean
Dim X As String
Dim GName As String
Dim BnkTot As Double
Dim RowIndex As Integer
Dim rstAddCheck As Recordset
GProcCheckForNumber Me, rstbl, MaxNo
If ValidateGrid = False Then
        'MSGrid.SetFocus
    Exit Function
End If

RowIndex = 1
'''If txttemp(6).text <> "" Then
'''    With msGrid
'''        Do While RowIndex < .rows
'''            '-- sub amt
'''            If Val(.TextMatrix(RowIndex, 6)) > 0 Then
'''                BnkTot = BnkTot + Val(.TextMatrix(RowIndex, 6))
'''            End If
'''            RowIndex = RowIndex + 1
'''        Loop
'''    End With
'''    If BnkTot > 0 Then
'''        MsgBox "For Bill Detail Party Name Should Be Blank.", vbInformation + vbOKOnly, Me.Caption
'''        txttemp(6).SetFocus
'''        Exit Function
'''    End If
'''End If
'''
'''
'''
''''---- Required Data
'''If ValidateGrid = False Then Exit Function
''''--- Vno
'''If CLng(txt(0)) = 0 Then
'''    MsgBox "Check Voucher No.", vbInformation + vbOKOnly, Me.Caption
'''    txt(0).SetFocus
'''    Exit Function
'''End If
''''--- Vdt
'''If GProcIsDateValid(mskDt(2)) = False Then
'''    MsgBox "Check Voucher Date.", vbInformation + vbOKOnly, Me.Caption
'''    mskDt(2).SetFocus
'''    Exit Function
'''End If
''''---- party
''''If txttemp(6) = "" Then
''''    MsgBox "Check Account.", vbInformation + vbOKOnly, Me.Caption
''''    txttemp(6).SetFocus
''''    Exit Function
''''End If
''''--Bank A/c
''''If GProcGetColumnValue("tblMastAccount", "AcName", txttemp(6), "S", "AcCode", "N") = 0 Then
''''   GName = GProcGetColumnValue("TblMastGroup", "AgCode", 90017, "N", "AgName", "S")
''''   GProcShowForm frmMastAccount, frmMain.mnuMstfrm(0), Array(frmMastAccount.txt(1), frmMastAccount.txttemp(2)), Array(txttemp(6), GName)
''''   Exit Function
''''End If
''''--- Cr A/c
'''If txttemp(7) = "" Then
'''    MsgBox "Check Account.", vbInformation + vbOKOnly, Me.Caption
'''    txttemp(7).SetFocus
'''    Exit Function
'''End If
''''--Create Cr A/c
'''If GProcGetColumnValue("tblMastAccount", "AcName", txttemp(7), "S", "AcCode", "N") = 0 Then
'''   GName = GProcGetColumnValue("TblMastGroup", "AgCode", 90019, "N", "AgName", "S")
'''   GProcShowForm frmMastAccount, frmMain.mnuMstfrm(0), Array(frmMastAccount.txt(1), frmMastAccount.txttemp(2)), Array(txttemp(7), GName)
'''   Exit Function
'''End If
''''--- Cr A/c ,Dr A/c
'''If txttemp(6) = txttemp(7) Then
'''    MsgBox "Check Account.Debit and Credit Accounts are Same.", vbInformation + vbOKOnly, Me.Caption
'''    txttemp(6).SetFocus
'''    Exit Function
'''End If
'''
''''--- Pending
'''
''''---- Amount
''''If CDbl(txt(6)) <= 0 Then
''''    MsgBox "Check Voucher Amount.", vbInformation + vbOKOnly, Me.Caption
''''    Exit Function
''''End If
''''----Duplicate Records
''''----Addition Case
'''
'''
''''If FormAction = vbDataActionAddNew Then
''''    X = "Select vno From " & tblName & " Where vno=" & CLng(txt(0)) & " and Vtype='" & frmTypeStr & "' and VYear=" & gCYear
''''    i = GProcRstOpen(rstAddCheck, X, "R")
''''    If i > 0 Then
''''        MsgBox "Record is Available of Voucher No. " & CLng(txt(0)), vbCritical + vbOKOnly, "Duplicate Record"
''''        txt(0).SetFocus
''''        Exit Function
''''    End If
''''    rstAddCheck.Close
''''End If
ValidateData = True
End Function
Private Sub ReadFields()
GProcShowRecord Me, rstbl, MaxNo
FillTempFromTxt
End Sub
'--- Related Record
Private Function IsRelatedRecord() As Boolean
'-- Audited
'If GProcISAudited(txt(0), txt(3), gCYear) = True Then
'    MsgBox "Record is Audited.You can not Modify/Delete this Record.", vbCritical + vbOKOnly, App.Title
'    IsRelatedRecord = True
'    Exit Function
'End If
'--- Related Record Receipt
'If rstbl.EOF And rstbl.BOF Then
'Else
'    If rstbl.Fields("VIsRefEntType") = 1 Then
'        MsgBox "Related Record is Present.You can not Modify/Delete this Record.", vbCritical + vbOKOnly, App.Title
'        IsRelatedRecord = True
'        Exit Function
'    End If
'End If
End Function
Private Sub ShowAccountBalance()
Dim DrAcBal  As Double
Dim CrAcBal As Double
'If IsDate(mskDt(2)) Then
'    DrAcBal = GProcGetAccountBalance(txttemp(6), mskDt(2))
'    CrAcBal = GProcGetAccountBalance(txttemp(7), mskDt(2))
'    txtDrBalance = IIf(DrAcBal > 0, DrAcBal & " Cr", Abs(DrAcBal) & " Dr")
'    txtCrBalance = IIf(CrAcBal > 0, CrAcBal & " Cr", Abs(CrAcBal) & " Dr")
'End If
End Sub
'---------------- Grid
Private Sub SetGrid()
Dim X As String, x1Rec As Double
Set rsRecVsSale = New Recordset

'0 - VBillNo
'1 - VBillType
'2 - VBillDate
'3 - Bill Amt
'4 - Pty Name
'5 - Brok NM
'6 - Amt Recd
'7 - Rec Dt
'8 - OD
'9 - Int
'10 - TDS On Amt
'11 - Tds
'12 - DbNtInt (Y/N)
'13 - DbNtIn Bill No
'14 - DbNtMonthly
'15 - Due On
'16 - Vno
'17 - VctrNo
'18 - Grase
'19 - IntRt



X = "select VBillNo,VBillType, VBillDate,VBillAmt,tblMastAccount.AcName,tblMastAccount_Br.AcName, VAmt,tblVoucher.Vdt,LateDays,Interest,TDSOn,LessTDS,DBNtInt,DbNtInBillNo,DbNtMonthly,DueOnDt,tblVoucher.Vno,VCtrNo,Grace,IntRt" _
& " from tblVoucher,tblMastAccount,tblMastAccount tblMastAccount_Br,tblSale where " _
& " tblVoucher.VType='" & frmTypeStr & "' and VnoList=" & txt(35) & " and tblVoucher.VFirm  = '" & gCCode & "' and  tblVoucher.VYear=" & gCYear & " and VCrAcCode = tblMastAccount.AcCode and  tblSale.SlBroker = tblMastAccount_Br.AcCode  and " _
& " VBillVno = tblSale.Vno and VBillType = tblSale.Vtype and VBillVYear = tblSale.Vyear and tblVoucher.VFirm = tblSale.Vfirm   order by tblMastAccount.AcIsIntDbNtMonthly,tblMastAccount.AcName,tblMastAccount_Br.AcName,tblVoucher.Vno,VCtrNo"
x1Rec = GProcRstOpen(rsRecVsSale, X, "R")
With msGrid
    .Clear
    .FormatString = "<Bill No           |<Tp |<Bill Date    |>Bill Amount    |<Party Name                     |<Broker Name |>Amt. Recd     |<Pmt. Date  |>OD |>Late Int.        |>TDS On Amt.  |>TDS Amt.  |DN|DN in Bill No||||||"
    .ColWidth(1) = 0
    .ColWidth(2) = 0
    .ColWidth(5) = 0
    .ColWidth(7) = 0
    .ColWidth(14) = 0
    .ColWidth(15) = 0
    .ColWidth(16) = 0
    .ColWidth(17) = 0
    .ColWidth(18) = 0
    .ColWidth(19) = 0
    If FormAction = vbDataActionAddNew Then
        .rows = 2
        GridAddNew
    Else
        .rows = 2
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
With msGrid
     If ((.TextMatrix(.Row, 0)) = "") And .Row = 1 And txtGrid.Visible = True Then
        msGrid.Col = 9
        msGrid.SetFocus
        If txtGrid.Visible = True Then
           txtGrid.SetFocus
        End If
     End If
End With
End Sub
Private Sub MSGrid_LostFocus()
CalAmount
End Sub
Private Sub msGrid_EnterCell()
If FormAction <> vbDataActionClose Then
    CheckForNumberTxtGrid
    DoEvents
    txtGrid = ""
    DoEvents
    If msGrid.text <> "" Then txtGrid.text = msGrid.text
    '--- Lock TextBox
'0 - VBillNo
'1 - VBillType
'2 - VBillDate
'3 - Bill Amt
'4 - Pty Name
'5 - Brok NM
'6 - Amt Recd
'7 - Rec Dt
'8 - OD
'9 - Int
'10 - TDS On Amt
'11 - Tds
'12 - DbNtInt (Y/N)
'13 - DbNtIn Bill No
'14 - DbNtMonthly
'15 - Due On
'16 - Vno
'17 - VctrNo
'18 - Grase
'19 - IntRt
    
    With msGrid
        Select Case .Col
            Case 0, 1, 2, 3, 4, 5, 6, 7, 8, 11
                txtGrid.Locked = True
            Case Else
                txtGrid.Locked = False
                If .Col = 9 Then
                   OldTdsOn = Val(.TextMatrix(.Row, 9))
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
        If (Val(.TextMatrix(.Row, 9)) = 0) And .rows > 2 Then
            .RemoveItem (.Row)
            txtGrid.text = ""
            If cmdBtn(13).Enabled = True Then
                cmdBtn(13).SetFocus
            Else
                cmdBtn(14).SetFocus
            End If
        End If
        txtGrid.Visible = False
        Exit Sub
    End If
    '---- Navigate Row ,Col
    If KeyAscii = 13 Then
        KeyAscii = 0
        If .Col = rsRecVsSale.Fields.Count - 10 Then
            If .Row = (.rows - 1) Then '-- Add New Row
                .rows = .rows + 1
                .Row = .Row + 1
'                GridAddNew
            Else '--Next Row
                .Row = .Row + 1
            End If
            .Col = 9
        Else    '-- Next col
            .Col = .Col + 1
        End If
    End If
End With

'---- Set TxtGrid
If Trim(msGrid.text) <> "" Then
    txtGrid = Trim(msGrid.text)
End If
If msGrid.Col < (rsRecVsSale.Fields.Count - 1) Then '- -3
    If msGrid.Col = 1 Then
       msGrid.Col = 3
    End If
    If msGrid.Col = 2 Then
       msGrid.Col = 3
    End If
    If msGrid.Col = 5 Then
       msGrid.Col = 6
    End If
    If msGrid.Col = 7 Then
       msGrid.Col = 8
    End If
    txtGrid.Move msGrid.CellLeft + msGrid.Left, msGrid.CellTop + msGrid.Top, msGrid.CellWidth, msGrid.CellHeight
End If
txtGrid.Visible = True
txtGrid.Enabled = True
DoEvents
If KeyAscii <> 0 Then
    Select Case rsRecVsSale.Fields(msGrid.Col).Type
        Case adChar, adVarWChar
             '-- Pending
'            If msGrid.Col = 0 Then  '-- Invoice
'                txtGrid_KeyPress (KeyAscii)
'            Else
                txtGrid.text = txtGrid.text & Chr(KeyAscii)
'            End If
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
CalAmount
End Sub
Private Sub CalBalance()
'Dim k As Integer
'Dim TAmt As Double
'With msGrid
'    k = 1
'    DoEvents
'    Do While k < .rows
'       TAmt = TAmt + Val(.TextMatrix(k, 6))
'       k = k + 1
'    Loop
'    txt(8) = Format(TAmt, FStr)
'End With
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
            Case 1 '-- Invoice
                DoEvents
                KeyAscii = 0
            Case Else
                KeyAscii = GProcValidateKey(rsRecVsSale, msGrid.Col, KeyAscii, txtGrid)
        End Select
   End If
   
    If KeyAscii = 13 Then
        CheckForNumberTxtGrid
    End If
    '--- Navigate Row ,Col
    
    If KeyAscii = 13 And .Col < rsRecVsSale.Fields.Count - 10 Then '- -4
        If .Col <= 8 Then
           .Col = 9
        Else
            .Col = .Col + 1
        End If
    ElseIf KeyAscii = 13 And .Col = rsRecVsSale.Fields.Count - 10 Then '- -4
        If .Row = (.rows - 1) Then
            .rows = .rows + 1
            .Row = .Row + 1
        Else
            .Row = .Row + 1
        End If
        .Col = 9
    End If
    '-- On ESC
    With msGrid
        If KeyAscii = 27 Then '----Esc Key
            KeyAscii = 0
            If (Val(.TextMatrix(.Row, 9)) = 0) And .rows > 2 Then
                .RemoveItem (.Row)
                txtGrid.text = ""
            End If
            txtGrid.Visible = False
            If cmdBtn(13).Enabled = True Then
                cmdBtn(13).SetFocus
            Else
                cmdBtn(14).SetFocus
            End If
            Exit Sub
        End If
    End With
End With
End Sub
Private Sub txtGrid_KeyDown(KeyCode As Integer, Shift As Integer)
'--- Delete Row On Ctrl+Y
If KeyCode = vbKeyY And Shift = 2 And msGrid.rows > 2 Then
    msGrid = 9
    msGrid.RemoveItem (msGrid.Row)
    txtGrid.Visible = False
    txtGrid.text = ""
    CalAmount
    msGrid.SetFocus
    Exit Sub
End If
'If msGrid.Col >= 9 And msGrid.Col <= rsRecVsSale.Fields.Count - 7 Then  '- -4
'        If msGrid.Row < (msGrid.rows - 1) Then
'           msGrid.Row = msGrid.Row + 1
'           msGrid_KeyPress (0)
''           txtGrid.SetFocus
'        End If
'End If
End Sub
Private Sub CheckForNumberGrid()
Dim RowIndex As Long
With rsRecVsSale
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
Dim Y As String
Dim RsP As Recordset
Dim JvNo As Long
Dim JvNoTDs As Long
JvNo = 0
JvNoTDs = 0
Y = "Select * from " & tblName & " where VnoList = " & txt(35) & " and " & CriteriaStr & " "
If (GProcRstOpen(RsP, Y, "O") > 0) Then
    JvNo = RsP.Fields("JvNoInt")
    JvNoTDs = RsP.Fields("JvNoLessTDS")
    
End If
Y = "select VCtrNo from tblVoucher where VType='JV' and VNo=" & JvNo & " and VFirm  = '" & MainFirm & "' and  VYear=" & gCYear & "  order by VCtrNo"
i = GProcRstOpen(RsP, Y, "R")
With RsP
     If Not .EOF Then
       .MoveFirst
       Do While Not .EOF
          gCn.Execute "Delete from tblvoucher where VType = 'JV' and VFirm = '" & MainFirm & "'  and VYear = " & Val(txt(3)) & "  and VCtrNo = " & RsP.Fields(0) & "   and  vno = " & JvNo
          .MoveNext
       Loop
     End If
End With
Y = "select VCtrNo from tblVoucher where VType='JV' and VNo=" & JvNoTDs & " and VFirm  = '" & MainFirm & "' and  VYear=" & gCYear & "  order by VCtrNo"
i = GProcRstOpen(RsP, Y, "R")
With RsP
     If Not .EOF Then
       .MoveFirst
       Do While Not .EOF
          gCn.Execute "Delete from tblvoucher where VType = 'JV' and VFirm = '" & MainFirm & "'  and VYear = " & Val(txt(3)) & "  and VCtrNo = " & RsP.Fields(0) & "   and  vno = " & JvNoTDs
          .MoveNext
       Loop
     End If
End With
DoEvents
gCn.Execute "update tblVoucher set Interest = 0, LessTDS = 0 ,DueOnDt =  '', VnoList = 0 ,Grace = 0,IntRt = 0,LateDays = 0,JvNoInt=0 ,JvNoLessTDS = 0,GraceExtra =0,IntFromDate = null,IntParty = 0,TdsRt = 0,TDSOn=0,DBNtInt = null,DbNtMonthly = 0,DbNtInBillNo = null " _
                       & " where VType = '" & UCase(txt(1)) & "' and VFirm = '" & txt(4) & "'  and VYear = " & Val(txt(3)) & "  and VnoList = " & Val(txt(35)) & " "
End Sub
Private Sub SaveRelatedRecord()
Dim i As Long, X As String
Dim RowIndex As Integer
RowIndex = 1
With msGrid
    Do While RowIndex < .rows
       If (Val(.TextMatrix(RowIndex, 9)) = 0) And .rows > 2 Then
          .RemoveItem (RowIndex)
       End If
       RowIndex = RowIndex + 1
    Loop
End With
'-- Delete Records
DeleteRelatedRecord
With msGrid
    i = 1
    Do While i < .rows
        DoEvents
        If CDec(Val(.TextMatrix(i, 9))) > 0 Then
           gCn.Execute "update tblVoucher set Interest = " & Val(.TextMatrix(i, 9)) & ", LessTDS = " & Val(.TextMatrix(i, 11)) & ",DueOnDt = '" & IIf(IsDate(.TextMatrix(i, 15)), Format(.TextMatrix(i, 15), "yyyy/mm/dd"), "") & "', VnoList = " & Val(txt(35)) & " ,Grace = " & Val(.TextMatrix(i, 18)) & ",IntRt = " & Val(.TextMatrix(i, 19)) & ",LateDays = " & Val(.TextMatrix(i, 8)) & ",  " _
                       & " GraceExtra = " & Val(txt(39)) & ",IntFromDate = '" & IIf(IsDate(dtpDt(40)), Format(dtpDt(40), "yyyy/mm/dd"), "") & "' , IntParty = " & Val(txt(41)) & ",  TdsRt = " & Val(txt(42)) & ",TDSOn = " & Val(.TextMatrix(i, 10)) & ",DBNtInt = '" & UCase(.TextMatrix(i, 12)) & "',DbNtMonthly = " & Val(.TextMatrix(i, 14)) & ",DbNtInBillNo = '" & UCase(.TextMatrix(i, 13)) & "'," _
                       & " JvNoInt=" & Val(txt(32)) & " ,JvNoLessTDS = " & Val(txt(34)) & "  " _
                       & " where VNo = " & Val(.TextMatrix(i, 16)) & " and  VType = '" & UCase(txt(1)) & "' and  VFirm = '" & txt(4) & "'     and VYear = " & Val(txt(3)) & "  and VCtrNo = " & Val(.TextMatrix(i, 17)) & " "
            DoEvents
        End If
       i = i + 1
     Loop
End With
If FormAction <> vbDataActionAddNew Then
   ApplyIntDbNt
End If


End Sub
Private Sub ApplyIntDbNt()
Dim Y As String
Dim RsP As Recordset
Dim wJvNo As Long
Dim wJvNoTDS As Long
Dim wParty As Long
Dim wIntAmt As Double
Dim wTdsAmt As Double
Dim wBlNo As String
Dim wBlNoTds As String
Dim rsDbNt As Recordset
Dim i As Integer
Dim J As Integer
Dim wVBillNo As String
Dim wVBillVno As Long
Dim wVBillType As String
Dim wVBillDate As Date
Dim wVBillVYear As Integer
Dim wVBillAmt As Double
Dim wTdsRt As Double
Dim wTdsOn As Double
Dim Mon As Integer
Mon = Month(dtpDt(2))
GProcRstOpen rsDbNt, "Select * from tblIntDbNote where VnoList = " & txt(35) & " and " & CriteriaStr & " ", "O"
Y = "Select * from " & tblName & " where VnoList = " & txt(35) & " and " & CriteriaStr & " "
'If (GProcRstOpen(Rsp, Y, "O") > 0) Then
'    Rsp.MoveFirst
'    Do While Not Rsp.EOF
'       If Val(Rsp.Fields("JvNoInt")) > 0 Then
'          gCn.Execute "Delete from tblvoucher where VType = 'JV' and VFirm = '" & txt(4) & "'  and VYear = " & Val(txt(3)) & "  and Vno = " & Val(Rsp.Fields("VnoList")) & " "
'          gCn.Execute "Delete from tblIntDbNote where VType = 'JV' and VFirm = '" & txt(4) & "'  and VYear = " & Val(txt(3)) & "  and Vno = " & Val(Rsp.Fields("VnoList")) & " "
'          DoEvents
'       End If
'       If Val(Rsp.Fields("JvNoLessTDS")) > 0 Then
'          gCn.Execute "Delete from tblvoucher where VType = 'JV' and VFirm = '" & txt(4) & "'  and VYear = " & Val(txt(3)) & "  and Vno = " & Val(Rsp.Fields("JvNoLessTDS")) & " "
'          gCn.Execute "Delete from tblIntDbNote where VType = 'JV' and VFirm = '" & txt(4) & "'  and VYear = " & Val(txt(3)) & "  and Vno = " & Val(Rsp.Fields("JvNoLessTDS")) & " "
'          DoEvents
'       End If
'       Rsp.MoveNext
'    Loop
'End If
DoEvents
Y = "Select * from " & tblName & " where VnoList = " & txt(35) & " and DBNtInt = 'Y' and " & CriteriaStr & " order by VCrAcCode,DbNtInBillNo"
If (GProcRstOpen(RsP, Y, "O") > 0) Then
    wParty = 0
    If Val(txt(32)) > 0 Then
       wJvNo = txt(32)
    Else
       wJvNo = GProcGenerateIdMonthwise("TblVoucher", "Vno", Mon, "Vdt", "Vtype='JV' and VFirm = '" & MainFirm & "' and VYear=" & gCYear)
    End If
    i = 1
    J = 1
    RsP.MoveFirst
    Do While Not RsP.EOF
       wParty = RsP.Fields("VCrAcCode")
       wIntAmt = 0
       wTdsAmt = 0
       wTdsOn = 0
       wBlNo = RsP.Fields("DbNtInBillNo")
       wBlNoTds = ""
       Do While Not RsP.EOF
           If RsP.Fields("DbNtInBillNo") <> wBlNo Or RsP.Fields("DbNtInBillNo") = "" Then
              wIntAmt = Val(RsP.Fields("Interest"))
              wTdsAmt = Val(RsP.Fields("LessTDS"))
              wTdsOn = Val(RsP.Fields("TDSOn"))
              wBlNo = RsP.Fields("VBillNo")
              wVBillVno = RsP.Fields("VBillVno")
              wVBillType = RsP.Fields("VBillType")
              wVBillDate = RsP.Fields("VBillDate")
              wVBillVYear = RsP.Fields("VBillVYear")
              wVBillAmt = RsP.Fields("VBillAmt")
              RsP.MoveNext
              Exit Do
           Else
              wIntAmt = wIntAmt + Val(RsP.Fields("Interest"))
              wTdsAmt = wTdsAmt + Val(RsP.Fields("LessTDS"))
              wTdsOn = wTdsOn + Val(RsP.Fields("TDSOn"))
              If RsP.Fields("VBillNo") = wBlNo Then
                 wVBillVno = RsP.Fields("VBillVno")
                 wVBillType = RsP.Fields("VBillType")
                 wVBillDate = RsP.Fields("VBillDate")
                 wVBillVYear = RsP.Fields("VBillVYear")
                 wVBillAmt = RsP.Fields("VBillAmt")
              End If
              RsP.MoveNext
              If Not RsP.EOF Then
                 If RsP.Fields("DbNtInBillNo") <> wBlNo Then Exit Do
              End If
           End If
       Loop
       '--- Interest Db Nt.
        If wIntAmt > 0 Then
            rstbl.AddNew
            rstbl!Vno = wJvNo '--VNo
            rstbl!Vtype = "JV"  '--Vtype
            rstbl!Vdt = CDate(txt(2)) '--Vdt(2)
            rstbl!VYear = txt(3) '--VYear
            rstbl!VFirm = MainFirm '--VFirm
            rstbl!VCtrNo = i '--Sr No
            i = i + 1
            rstbl!VCrAcCode = gLatePayIntAcCodeRec
            rstbl!VDrAcCode = wParty
            rstbl!VTopCrDr = 0
            rstbl!vamt = wIntAmt
            rstbl!VNar1 = "Late Payment Debited "
            rstbl!VNar2 = "Agt. Bill No : " & wBlNo
            rstbl!VNar3 = ""
            rstbl!VNAR4 = ""
            rstbl!Interest = wIntAmt
            rstbl!VBillNo = wBlNo
            rstbl!VBillVno = wVBillVno
            rstbl!VBillType = wVBillType
            rstbl!VBillDate = wVBillDate
            rstbl!VBillVYear = wVBillVYear
            rstbl!VBillAmt = wVBillAmt
            rstbl!VBillFirm = gCCode
            rstbl.Update
            gCn.Execute "update tblVoucher set JvNoInt = " & wJvNo & " " _
                        & " where VCrAcCode = " & wParty & " and  VnoList = " & txt(35) & " and " & CriteriaStr & ""
            DoEvents
        End If
        '--- TDS Db Nt.
        If wTdsAmt > 0 Then
            If J = 1 Then
               If Val(txt(34)) > 0 Then
                  wJvNoTDS = txt(34)
               Else
                   wJvNoTDS = GProcGenerateIdMonthwise("TblVoucher", "Vno", Mon, "Vdt", "Vtype='JV' and VFirm = '" & MainFirm & "' and VYear=" & gCYear)
               End If
            End If
            rstbl.AddNew
            rstbl!Vno = wJvNoTDS '--VNo
            rstbl!Vtype = "JV"  '--Vtype
            rstbl!Vdt = CDate(txt(2)) '--Vdt(2)
            rstbl!VYear = txt(3) '--VYear
            rstbl!VFirm = MainFirm '--VFirm
            rstbl!VCtrNo = J '--Sr No
            J = J + 1
            rstbl!VDrAcCode = gTDSAcCodeRec
            rstbl!VCrAcCode = wParty
            rstbl!VTopCrDr = 1
            rstbl!vamt = wTdsAmt
            rstbl!VNar1 = "T D S @ " + CStr(gTDSRate) + " % On Int. Of Rs. " + CStr(wTdsOn) + " Debited"
            rstbl!VNar2 = "Agt. Bill No : " & wBlNo
            rstbl!VNar3 = ""
            rstbl!VNAR4 = ""
            rstbl!LessTDS = wTdsAmt
            rstbl!VBillNo = wBlNo
            rstbl!VBillVno = wVBillVno
            rstbl!VBillType = wVBillType
            rstbl!VBillDate = wVBillDate
            rstbl!VBillVYear = wVBillVYear
            rstbl!VBillAmt = wVBillAmt
            rstbl!TdsRt = wTdsRt
            rstbl!TDSOn = wTdsOn
            rstbl!VBillFirm = gCCode
            rstbl.Update
            gCn.Execute "update tblVoucher set JvNoLessTDS = " & wJvNoTDS & " " _
                        & " where VCrAcCode = " & wParty & " and  VnoList = " & txt(35) & " and " & CriteriaStr & " and LessTDS <> 0"
            DoEvents
       End If
'''       RsP.MoveNext
    Loop
End If

'---For Db Nt Without A/c

Y = "Select * from " & tblName & " where VnoList = " & txt(35) & " and DBNtInt <> 'Y' and " & CriteriaStr & " order by VCrAcCode,DbNtInBillNo"

If (GProcRstOpen(RsP, Y, "O") > 0) Then
    wParty = 0
    wJvNo = GProcGenerateIdMonthwise("tblIntDbNote", "Vno", Mon, "Vdt", "Vtype='JV' and VFirm = '" & gCCode & "' and VYear=" & gCYear)
    i = 1
    J = 1
    RsP.MoveFirst
    Do While Not RsP.EOF
       wParty = RsP.Fields("VCrAcCode")
       wIntAmt = 0
       wTdsAmt = 0
       wTdsOn = 0
       wBlNo = RsP.Fields("DbNtInBillNo")
       wBlNoTds = ""
       Do While Not RsP.EOF
           If RsP.Fields("DbNtInBillNo") <> wBlNo Or RsP.Fields("DbNtInBillNo") = "" Then
              wIntAmt = Val(RsP.Fields("Interest"))
              wTdsAmt = Val(RsP.Fields("LessTDS"))
              wTdsOn = Val(RsP.Fields("TDSOn"))
              wBlNo = RsP.Fields("VBillNo")
              wVBillVno = RsP.Fields("VBillVno")
              wVBillType = RsP.Fields("VBillType")
              wVBillDate = RsP.Fields("VBillDate")
              wVBillVYear = RsP.Fields("VBillVYear")
              wVBillAmt = RsP.Fields("VBillAmt")
              RsP.MoveNext
              Exit Do
           Else
              wIntAmt = wIntAmt + Val(RsP.Fields("Interest"))
              wTdsAmt = wTdsAmt + Val(RsP.Fields("LessTDS"))
              wTdsOn = wTdsOn + Val(RsP.Fields("TDSOn"))
              If RsP.Fields("VBillNo") = wBlNo Then
                 wVBillVno = RsP.Fields("VBillVno")
                 wVBillType = RsP.Fields("VBillType")
                 wVBillDate = RsP.Fields("VBillDate")
                 wVBillVYear = RsP.Fields("VBillVYear")
                 wVBillAmt = RsP.Fields("VBillAmt")
              End If
              RsP.MoveNext
           End If
       Loop
       '--- Interest Db Nt.
        If wIntAmt > 0 Then
            rsDbNt.AddNew
            rsDbNt!Vno = wJvNo '--VNo
            rsDbNt!Vtype = "JV"  '--Vtype
            rsDbNt!Vdt = CDate(txt(2)) '--Vdt(2)
            rsDbNt!VYear = txt(3) '--VYear
            rsDbNt!VFirm = MainFirm '--VFirm
            rsDbNt!VCtrNo = i '--Sr No
            i = i + 1
            rsDbNt!VCrAcCode = gLatePayIntAcCodeRec
            rsDbNt!VDrAcCode = wParty
            rsDbNt!VTopCrDr = 0
            rsDbNt!vamt = wIntAmt
            rsDbNt!VNar1 = "Late Payment Debited "
            rsDbNt!VNar2 = "Agt. Bill No : " & wBlNo
            rsDbNt!VNar3 = ""
            rsDbNt!VNAR4 = ""
            rsDbNt!Interest = wIntAmt
            rsDbNt!VBillNo = wBlNo
            rsDbNt!VBillVno = wVBillVno
            rsDbNt!VBillType = wVBillType
            rsDbNt!VBillDate = wVBillDate
            rsDbNt!VBillVYear = wVBillVYear
            rsDbNt!VBillAmt = wVBillAmt
            rsDbNt!VBillFirm = gCCode
            rsDbNt.Update
            gCn.Execute "update tblVoucher set JvNoInt = " & wJvNo & " " _
                        & " where VCrAcCode = " & wParty & " and  VnoList = " & txt(35) & " and " & CriteriaStr & ""
            DoEvents
        End If
'''       RsP.MoveNext
    Loop
End If

'If (GProcRstOpen(RsP, Y, "O") > 0) Then
'    wParty = 0
'    wJvNo = GProcGenerateIdMonthwise("tblIntDbNote", "Vno", Mon, "Vdt", "Vtype='JV' and VFirm = '" & gCCode & "' and VYear=" & gCYear)
'    i = 1
'    RsP.MoveFirst
'    Do While Not RsP.EOF
'       wParty = RsP.Fields("VCrAcCode")
'       wIntAmt = 0
'       wTdsAmt = 0
'       wTdsOn = 0
'       wBlNo = ""
'       wBlNoTds = ""
'       wVBillNo = RsP.Fields("VBillNo")
'       wVBillVno = RsP.Fields("VBillVno")
'       wVBillType = RsP.Fields("VBillType")
'       wVBillDate = RsP.Fields("VBillDate")
'       wVBillVYear = RsP.Fields("VBillVYear")
'       wVBillAmt = RsP.Fields("VBillAmt")
'       wTdsRt = RsP.Fields("TdsRt")
'       wTdsOn = RsP.Fields("TDSOn")
'       wIntAmt = wIntAmt + Val(RsP.Fields("Interest"))
'       wTdsAmt = wTdsAmt + Val(RsP.Fields("LessTDS"))
'       wTdsOn = wTdsOn + Val(RsP.Fields("TDSOn"))
'       wBlNo = CStr(RsP.Fields("VBillNo"))
'       wBlNoTds = CStr(RsP.Fields("VBillNo"))
'         '--- Interest Db Nt.
'        If wIntAmt > 0 Then
''            wJvNo = GProcGenerateIdMonthwise("tblIntDbNote", "Vno", Mon, "Vdt", "Vtype='JV' and VFirm = '" & gCCode & "' and VYear=" & gCYear)
'            i = 1
'            rsDbNt.AddNew
'            rsDbNt!Vno = wJvNo '--VNo
'            rsDbNt!Vtype = "JV"  '--Vtype
'            rsDbNt!Vdt = CDate(txt(2)) '--Vdt(2)
'            rsDbNt!VYear = txt(3) '--VYear
'            rsDbNt!VFirm = gCCode '--VFirm
'            rsDbNt!VCtrNo = i '--Sr No
'            i = i + 1
'            rsDbNt!VCrAcCode = gLatePayIntAcCodeRec
'            rsDbNt!VDrAcCode = wParty
'            rsDbNt!VTopCrDr = 0
'            rsDbNt!vamt = wIntAmt
'            rsDbNt!VNar1 = "Interest Debited Against Bill No : " & wBlNo
'            rsDbNt!VNar2 = ""
'            rsDbNt!VNar3 = ""
'            rsDbNt!VNAR4 = ""
'            rsDbNt!Interest = wIntAmt
'            rsDbNt!VBillNo = wVBillNo
'            rsDbNt!VBillVno = wVBillVno
'            rsDbNt!VBillType = wVBillType
'            rsDbNt!VBillDate = wVBillDate
'            rsDbNt!VBillVYear = wVBillVYear
'            rsDbNt!VBillAmt = wVBillAmt
'            rsDbNt.Update
'            gCn.Execute "update tblVoucher set JvNoInt = " & wJvNo & " " _
'                        & " where VCrAcCode = " & wParty & " and  VnoList = " & txt(35) & " and " & CriteriaStr & ""
'            DoEvents
'        End If
'        RsP.MoveNext
'    Loop
'End If
If GProcRstOpen(rstbl, "Select * from " & tblName & " where VnoList = " & txt(35) & " and  " & CriteriaStr & "   ", "O") > 0 Then ReadFields
End Sub
Private Sub FillRelatedRecord()
Dim RowIndex As Long
Dim ColIndex As Long
Dim BnkTot As Double
'--Fill Grid
SetGrid
BnkTot = 0
With rsRecVsSale
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
'                If ColIndex = 32 Then BnkTot = BnkTot + Val(msGrid.TextMatrix(RowIndex, 32)) '-- jv
            Next ColIndex
                If msGrid.TextMatrix(RowIndex, 14) = 1 Then
                   msGrid.Row = RowIndex
                   Dim k As Integer
                   For k = 0 To 13
                   msGrid.Col = k
                   msGrid.CellForeColor = &HC0&
                   Next k
                End If
            RowIndex = RowIndex + 1
            .MoveNext
            If Not .EOF Then msGrid.rows = msGrid.rows + 1
        Loop
    End If
End With
CalAmount
End Sub
Private Sub CalAmount()
Dim BnkTot As Double
Dim BnkTot2 As Double
Dim BnkTot3 As Double
Dim RowIndex As Integer
RowIndex = 1
BnkTot = 0
BnkTot2 = 0
BnkTot3 = 0
With msGrid
    Do While RowIndex < .rows
        '-- sub amt
        If Val(.TextMatrix(RowIndex, 9)) > 0 Then
            If UCase(.TextMatrix(RowIndex, 12)) = "Y" Then
               BnkTot = BnkTot + Val(.TextMatrix(RowIndex, 9))
               BnkTot2 = BnkTot2 + Val(.TextMatrix(RowIndex, 11))
            Else
               BnkTot3 = BnkTot3 + Val(.TextMatrix(RowIndex, 9))
            End If
        End If
        RowIndex = RowIndex + 1
    Loop
End With
lblInt.Caption = Format(BnkTot, "###0.00")
lblTDS.Caption = Format(BnkTot2, "###0.00")
lblInt2.Caption = Format(BnkTot3, "###0.00")
'txt(8) = Format(BnkTot, "###0.00")
End Sub
Private Sub SetControlEd()
If FormAction = vbDataActionUpdate Then
   dtpDt(40).Enabled = False
   dtpDt(2).Enabled = False
   txttemp(41).Enabled = False
   txt(37).Enabled = False
   txt(38).Enabled = False
   txt(39).Enabled = False
   txt(42).Enabled = False
   chkIsMonthlyDbNt.Enabled = False
End If
txt(32).Enabled = False
txt(34).Enabled = False
End Sub
Private Sub PreperReport()
Dim X As String
Dim wFromCd As Long
Dim LateDayCd As String
Dim vamt As Double
Dim Lint As Double
Dim LTds As Double
Dim RsTmp As Recordset
Dim i As Long, J As Long
Dim NewRow As Boolean
Dim IntRt As Double
Dim TotCol(0, 1) As Long
               
'0 - VBillNo
'1 - VBillType
'2 - VBillDate
'3 - Bill Amt
'4 - Pty Name
'5 - Brok NM
'6 - Amt Recd
'7 - Rec Dt
'8 - OD
'9 - Int
'10 - TDS On Amt
'11 - Tds
'12 - DbNtInt (Y/N)
'13 - DbNtIn Bill No
'14 - DbNtMonthly
'15 - Due On
'16 - Vno
'17 - VctrNo
'18 - Grase
'19 - IntRt

X = "Select a.VBillNo,d.Vtype,a.VBillDate,a.VBillVYear,d.VFirm,d.SlBillAmt,a.VAmt,a.Vdt,b.AcOurCode,b.Acname,dateadd(dd,d.DueDays+d.SlCrDays+iif(b.AcDueDays>0,b.AcDueDays," & Val(txt(37)) & ") ,d.Vdt) , " _
                   & " c.AcOurCode,c.AcName,a.Vtype,'" & gUserName & "', DateDiff(day, d.Vdt, a.Vdt)-(d.DueDays+d.SlCrDays)-iif(b.AcDueDays>0,b.AcDueDays," & Val(txt(37)) & ") as OD,a.vno,a.VCtrNo,a.TDSOn,b.AcIntPer,b.AcIsIntDbNtMonthly,iif(b.AcDueDays>0,b.AcDueDays," & Val(txt(37)) & ") as Gease,DbNtInBillNo from tblVoucher a , tblMastAccount b, tblMastAccount c, tblsale d " _
                   & " where a.VBillVno = d.VNo and a.VBillType = d.VType and a.VBillVYear = d.Vyear and a.VFirm = d.VFirm " _
                   & " and d.SlAcDrCode = b.AcCode and d.SlBroker = c.AcCode and a.Vtype = 'BR' " _
                   & " and a.VDt >= '" & Format(dtpDt(40), "yyyy/mm/dd") & "'  and a.VDt <= '" & Format(dtpDt(2), "yyyy/mm/dd") & "' and VnoList = 0 "
 
X = X + " and a.Vfirm = '" & gCCode & "' and a.VBillDate < '" & Format("2017/07/01", "yyyy/mm/dd") & "'   "

If txttemp(41) <> "" Then
   X = X + " and b.Acname = '" & txttemp(41) & "'"
End If

If chkIsMonthlyDbNt.Value = 1 Then
   X = X + " and b.AcIsIntDbNtMonthly = 1"
End If

X = X + " order by b.AcIsIntDbNtMonthly,b.Acname,c.AcName,a.vno,a.VCtrNo"

If GProcRstOpen(RsTmp, X, "R", gCn) > 0 Then
   RsTmp.MoveFirst
   Do While Not RsTmp.EOF
      With msGrid
           If Val(RsTmp.Fields(15)) > Val(txt(39)) Then
                J = .rows - 1
                .TextMatrix(J, 0) = RsTmp.Fields(0)                    '0 - VBillNo
                .TextMatrix(J, 1) = RsTmp.Fields(1)                    '1 - VBillType
                .TextMatrix(J, 2) = RsTmp.Fields(2)                    '2 - VBillDate
                .TextMatrix(J, 3) = CDbl(RsTmp.Fields(5))              '3 - Bill Amt
                .TextMatrix(J, 4) = RsTmp.Fields(9)                    '4 - Pty Name
                .TextMatrix(J, 5) = RsTmp.Fields(12)                   '5 - Brok NM
                .TextMatrix(J, 6) = CDbl(RsTmp.Fields(6))              '6 - Amt Recd
                .TextMatrix(J, 7) = RsTmp.Fields(7)                    '7 - Rec Dt
                .TextMatrix(J, 8) = RsTmp.Fields(15)                   '8 - OD
                If RsTmp.Fields("AcIntPer") > 0 Then
                   IntRt = RsTmp.Fields("AcIntPer")
                Else
                   IntRt = Val(txt(38))
                End If
                Lint = GProcMakeRounding(((CDbl(RsTmp.Fields(6)) * IntRt) / 100) / 365) * Val(RsTmp.Fields(15))
                DoEvents
                Lint = Format(Lint, "#######0.00")
                .TextMatrix(J, 9) = Format(Lint, "##########0.00")      '9 - Int
                If RsTmp.Fields(18) > 0 Then
                   .TextMatrix(J, 10) = RsTmp.Fields(18)                '10 - TDS On Amt
                Else
                   .TextMatrix(J, 10) = Format(Lint, "##########0.00")  '10 - TDS On Amt
                End If
                LTds = GProcMakeRounding((.TextMatrix(J, 10) * Val(txt(42))) / 100)
                DoEvents
                
                .TextMatrix(J, 11) = LTds         '11 - Tds
                .TextMatrix(J, 12) = "Y"                                                         '12 - DbNtInt (Y/N)
                .TextMatrix(J, 13) = RsTmp.Fields("DbNtInBillNo") & vbNullString                 '13 - DbNtIn Bill No
                .TextMatrix(J, 14) = RsTmp.Fields("AcIsIntDbNtMonthly")                          '14 - DbNtMonthly
                .TextMatrix(J, 15) = RsTmp.Fields(10)                                            '15 - Due On
                .TextMatrix(J, 16) = RsTmp.Fields(16)                                            '16 - Vno
                .TextMatrix(J, 17) = RsTmp.Fields(17)                                            '17 - VctrNo
                .TextMatrix(J, 18) = RsTmp.Fields("Gease")                                       '18 - Grase
                .TextMatrix(J, 19) = IntRt                                                       '19 - IntRt
                If .TextMatrix(J, 14) = 1 Then
                   .Row = J
                   Dim k As Integer
                   For k = 0 To 13
                   .Col = k
                   .CellForeColor = &HC0&
                   Next k
'                   .Row.C
 '                   Row.cells.ForeColor = System.Drawing.color.Red
'                   .ForeColor = vbRed
                End If
                J = J + 1
                .rows = .rows + 1
                .Col = 9
           End If
      End With
      RsTmp.MoveNext
   Loop
msGrid.Col = 9
msGrid.SetFocus
txtGrid.SetFocus
End If

End Sub
Private Sub txtTemp_KeyPress(Index As Integer, KeyAscii As Integer)
Dim LeftPos As Long
Dim TopPos As Long
Dim AgCode As Long
Dim BrkAcCode As Long
TopPos = txttemp(Index).Top + Me.Top + 650
LeftPos = txttemp(Index).Left + Me.Left
If KeyAscii = 13 Then Exit Sub
Select Case Index
    Case 41  '-- Party   Alias Account
        gClsSearch.SearchMultiField "QryMastAccount", "AcOurCode,Account,CGName,AcContPer,AcCity,AgName,AcBST,AcMsOmsParty", Array("Code", "Account Name", "Party Group", "Contact Person", "City", "Group", "TIN", "Ms/Oms"), Array(1000, 4000, 2000, 2000, 2000, 2500, 0, 0), "(Agcode in (90017) )", Chr(KeyAscii), "", LeftPos, TopPos, True
        KeyAscii = 0
        If gClsSearch.SearchMultiRetCol(1) <> "" Then
            txttemp(Index) = gClsSearch.SearchMultiRetCol(1)
        End If
End Select
End Sub
Private Sub CalGridAmt()
Dim TmpAmt As Double
If OldTdsOn <> Val(msGrid.TextMatrix(msGrid.Row, 9)) And OldTdsOn <> 0 Then
    msGrid.TextMatrix(msGrid.Row, 10) = msGrid.TextMatrix(msGrid.Row, 9)
    OldTdsOn = Val(msGrid.TextMatrix(msGrid.Row, 10))
End If
If Val(msGrid.TextMatrix(msGrid.Row, 10)) > 0 Then
    TmpAmt = GProcMakeRounding((msGrid.TextMatrix(msGrid.Row, 10) * Val(txt(42))) / 100)
    msGrid.TextMatrix(msGrid.Row, 11) = TmpAmt
    msGrid.TextMatrix(msGrid.Row, 11) = Format(msGrid.TextMatrix(msGrid.Row, 11), FStr)
Else
    msGrid.TextMatrix(msGrid.Row, 11) = 0
    msGrid.TextMatrix(msGrid.Row, 11) = Format(msGrid.TextMatrix(msGrid.Row, 11), FStr)
End If

If Val(msGrid.TextMatrix(msGrid.Row, 9)) = 0 Then
    msGrid.TextMatrix(msGrid.Row, 10) = 0
    msGrid.TextMatrix(msGrid.Row, 10) = Format(msGrid.TextMatrix(msGrid.Row, 11), FStr)
    msGrid.TextMatrix(msGrid.Row, 11) = 0
    msGrid.TextMatrix(msGrid.Row, 11) = Format(msGrid.TextMatrix(msGrid.Row, 11), FStr)
End If

End Sub
