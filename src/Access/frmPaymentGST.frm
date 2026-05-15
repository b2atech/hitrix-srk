VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form frmPaymentGST 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00C2E3FC&
   BorderStyle     =   0  'None
   Caption         =   "Payment Entry"
   ClientHeight    =   9135
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   14520
   Icon            =   "frmPaymentGST.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   9135
   ScaleWidth      =   14520
   ShowInTaskbar   =   0   'False
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
      Height          =   375
      Index           =   37
      Left            =   5250
      MaxLength       =   40
      TabIndex        =   95
      Text            =   "37"
      Top             =   3870
      Width           =   6060
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
      Height          =   3375
      Left            =   2970
      TabIndex        =   60
      Top             =   4500
      Visible         =   0   'False
      Width           =   6960
      Begin VB.CheckBox chkIsChqPrint 
         BackColor       =   &H00FFC0C0&
         Caption         =   "Cheque Print"
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
         Left            =   4560
         TabIndex        =   85
         Top             =   360
         Width           =   2295
      End
      Begin VB.CheckBox Check2 
         BackColor       =   &H00FFC0C0&
         Caption         =   "Chq Print For NEFT"
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
         Left            =   4560
         TabIndex        =   90
         Top             =   1170
         Width           =   2175
      End
      Begin VB.TextBox Text1 
         Height          =   375
         Left            =   360
         TabIndex        =   89
         Top             =   2160
         Width           =   6135
      End
      Begin VB.CheckBox Check1 
         BackColor       =   &H00FFC0C0&
         Caption         =   "Chq Print For RTGS"
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
         Left            =   4560
         TabIndex        =   88
         Top             =   765
         Width           =   2175
      End
      Begin VB.OptionButton Option2 
         BackColor       =   &H00FFC0C0&
         Caption         =   "Bearer"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   3360
         TabIndex        =   87
         Top             =   1650
         Width           =   1215
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H00FFC0C0&
         Caption         =   "A/c Payee"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   1560
         TabIndex        =   86
         Top             =   1650
         Width           =   1335
      End
      Begin VB.CommandButton cmdPrint 
         Caption         =   "&RTGS"
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
         Left            =   3720
         TabIndex        =   70
         Top             =   2760
         Width           =   1020
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
         Left            =   4800
         TabIndex        =   65
         Top             =   2760
         Width           =   1335
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
         Left            =   2280
         TabIndex        =   64
         Top             =   2760
         Width           =   1335
      End
      Begin VB.TextBox txtno 
         Appearance      =   0  'Flat
         Height          =   360
         Index           =   1
         Left            =   2670
         MaxLength       =   7
         TabIndex        =   63
         Top             =   840
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
         Left            =   840
         TabIndex        =   62
         Top             =   2760
         Width           =   1335
      End
      Begin VB.TextBox txtno 
         Appearance      =   0  'Flat
         Height          =   360
         Index           =   0
         Left            =   2670
         MaxLength       =   7
         TabIndex        =   61
         Top             =   360
         Width           =   1575
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
         Left            =   1320
         TabIndex        =   67
         Top             =   840
         Width           =   1050
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
         Left            =   1320
         TabIndex        =   66
         Top             =   360
         Width           =   1065
      End
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
      Height          =   375
      Index           =   38
      Left            =   13770
      MaxLength       =   40
      TabIndex        =   94
      Text            =   "38"
      Top             =   8685
      Visible         =   0   'False
      Width           =   300
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
      Height          =   325
      Index           =   36
      Left            =   2565
      MaxLength       =   40
      TabIndex        =   93
      Text            =   "36"
      Top             =   9765
      Visible         =   0   'False
      Width           =   690
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
      Height          =   325
      Index           =   35
      Left            =   1665
      MaxLength       =   40
      TabIndex        =   92
      Text            =   "35"
      Top             =   9765
      Visible         =   0   'False
      Width           =   690
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
      Height          =   325
      Index           =   34
      Left            =   900
      MaxLength       =   40
      TabIndex        =   91
      Text            =   "34"
      Top             =   9945
      Visible         =   0   'False
      Width           =   690
   End
   Begin VB.TextBox txt 
      Alignment       =   1  'Right Justify
      Height          =   360
      Index           =   33
      Left            =   14760
      TabIndex        =   84
      Text            =   "33"
      Top             =   3960
      Visible         =   0   'False
      Width           =   345
   End
   Begin VB.TextBox txt 
      Height          =   375
      Index           =   22
      Left            =   16470
      TabIndex        =   83
      Text            =   "22"
      Top             =   360
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.TextBox txt 
      Height          =   375
      Index           =   23
      Left            =   16515
      TabIndex        =   82
      Text            =   "23"
      Top             =   810
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.TextBox txt 
      Height          =   375
      Index           =   24
      Left            =   16470
      TabIndex        =   81
      Text            =   "24"
      Top             =   1260
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.TextBox txt 
      Height          =   375
      Index           =   25
      Left            =   16560
      TabIndex        =   80
      Text            =   "25"
      Top             =   1665
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.TextBox txt 
      Height          =   375
      Index           =   26
      Left            =   16560
      TabIndex        =   79
      Text            =   "26"
      Top             =   2070
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.TextBox txt 
      Height          =   375
      Index           =   27
      Left            =   16515
      TabIndex        =   78
      Text            =   "27"
      Top             =   2520
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.TextBox txt 
      Height          =   375
      Index           =   28
      Left            =   16605
      TabIndex        =   77
      Text            =   "28"
      Top             =   2925
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.TextBox txt 
      Height          =   375
      Index           =   29
      Left            =   16605
      TabIndex        =   76
      Text            =   "29"
      Top             =   3330
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.TextBox txt 
      Height          =   375
      Index           =   30
      Left            =   16560
      TabIndex        =   75
      Text            =   "30"
      Top             =   3780
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.TextBox txt 
      Height          =   375
      Index           =   31
      Left            =   16650
      TabIndex        =   74
      Text            =   "31"
      Top             =   4230
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.TextBox txt 
      Height          =   375
      Index           =   32
      Left            =   16650
      TabIndex        =   73
      Text            =   "32"
      Top             =   4680
      Visible         =   0   'False
      Width           =   495
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
      Height          =   325
      Index           =   21
      Left            =   15180
      MaxLength       =   40
      TabIndex        =   71
      Text            =   "21"
      Top             =   7890
      Visible         =   0   'False
      Width           =   690
   End
   Begin VB.CheckBox chkDontShowInUnadj 
      BackColor       =   &H00FFC0FF&
      Caption         =   "Dont Show In Unadjusted Payment"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   7785
      TabIndex        =   17
      Top             =   8505
      Width           =   3315
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
      Height          =   390
      Left            =   8895
      TabIndex        =   69
      Text            =   "txtCrBalance"
      Top             =   1320
      Width           =   2385
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
      Height          =   390
      Left            =   8895
      TabIndex        =   68
      Text            =   "txtDrBalance"
      Top             =   1755
      Width           =   2385
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
      Left            =   12270
      Style           =   1  'Graphical
      TabIndex        =   39
      ToolTipText     =   "Last"
      Top             =   1140
      Width           =   495
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
      Height          =   325
      Index           =   20
      Left            =   15180
      MaxLength       =   40
      TabIndex        =   57
      Text            =   "20"
      Top             =   7530
      Visible         =   0   'False
      Width           =   690
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
      Height          =   325
      Index           =   19
      Left            =   15180
      MaxLength       =   40
      TabIndex        =   56
      Text            =   "19"
      Top             =   7170
      Visible         =   0   'False
      Width           =   690
   End
   Begin VB.TextBox txttemp 
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
      Left            =   2340
      MaxLength       =   40
      TabIndex        =   18
      Text            =   "txttemp(20)"
      ToolTipText     =   "Press F3 for Creation"
      Top             =   8505
      Width           =   5040
   End
   Begin VB.TextBox txtGrid 
      BackColor       =   &H8000000A&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   315
      TabIndex        =   16
      Top             =   6135
      Visible         =   0   'False
      Width           =   795
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
      Height          =   325
      Index           =   18
      Left            =   15180
      MaxLength       =   40
      TabIndex        =   55
      Text            =   "18"
      Top             =   6690
      Visible         =   0   'False
      Width           =   690
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
      Left            =   1575
      TabIndex        =   8
      Top             =   3195
      Visible         =   0   'False
      Width           =   1440
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
      Height          =   375
      Index           =   7
      Left            =   5250
      MaxLength       =   40
      TabIndex        =   11
      Text            =   "txttemp1(7)"
      Top             =   2595
      Width           =   6060
   End
   Begin VB.TextBox txt 
      Height          =   345
      Index           =   17
      Left            =   15315
      TabIndex        =   54
      Text            =   "17"
      Top             =   6255
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.TextBox txt 
      Height          =   345
      Index           =   16
      Left            =   15315
      TabIndex        =   53
      Text            =   "16"
      Top             =   5775
      Visible         =   0   'False
      Width           =   495
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
      Height          =   345
      Index           =   8
      Left            =   1545
      TabIndex        =   7
      Text            =   "8"
      Top             =   3165
      Width           =   555
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
      Height          =   390
      Index           =   8
      Left            =   5250
      MaxLength       =   40
      TabIndex        =   12
      Text            =   "8"
      ToolTipText     =   "Press F3 for Creation & Press F2 for Selection"
      Top             =   3000
      Width           =   6060
   End
   Begin VB.TextBox txt 
      Height          =   345
      Index           =   15
      Left            =   15315
      TabIndex        =   52
      Text            =   "15"
      Top             =   5295
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.TextBox txt 
      Height          =   375
      Index           =   12
      Left            =   15330
      TabIndex        =   51
      Text            =   "12"
      Top             =   3180
      Visible         =   0   'False
      Width           =   375
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
      Height          =   375
      Index           =   10
      Left            =   12510
      TabIndex        =   14
      Text            =   "10"
      Top             =   8685
      Visible         =   0   'False
      Width           =   390
   End
   Begin VB.TextBox txt 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFFFF&
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
      Index           =   0
      Left            =   1560
      TabIndex        =   1
      Text            =   "0"
      Top             =   585
      Width           =   1320
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
      Left            =   1560
      MaxLength       =   40
      TabIndex        =   5
      Text            =   "txttemp(5)"
      ToolTipText     =   "Press F3 for Creation"
      Top             =   1800
      Width           =   6885
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
      Index           =   4
      Left            =   1560
      MaxLength       =   40
      TabIndex        =   4
      Text            =   "txttemp(4)"
      ToolTipText     =   "Press F3 for Creation"
      Top             =   1365
      Width           =   6885
   End
   Begin VB.TextBox txt 
      Alignment       =   1  'Right Justify
      Height          =   375
      Index           =   14
      Left            =   14940
      TabIndex        =   45
      Text            =   "14"
      Top             =   0
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.TextBox txt 
      Alignment       =   1  'Right Justify
      Height          =   360
      Index           =   13
      Left            =   15300
      TabIndex        =   44
      Text            =   "13"
      Top             =   3615
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
      Left            =   11670
      Style           =   1  'Graphical
      TabIndex        =   36
      ToolTipText     =   "Previous"
      Top             =   690
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
      Left            =   12270
      Style           =   1  'Graphical
      TabIndex        =   37
      ToolTipText     =   "Next"
      Top             =   690
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
      Left            =   11670
      Style           =   1  'Graphical
      TabIndex        =   38
      ToolTipText     =   "First"
      Top             =   1140
      Width           =   495
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
      TabIndex        =   35
      ToolTipText     =   "To Find Record"
      Top             =   720
      Width           =   1095
   End
   Begin VB.CommandButton cmdBtn 
      Caption         =   "&Mill Vou.Print"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   5
      Left            =   14985
      Style           =   1  'Graphical
      TabIndex        =   33
      Top             =   2655
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
      Left            =   11670
      Style           =   1  'Graphical
      TabIndex        =   28
      ToolTipText     =   "To Add New Record"
      Top             =   1665
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
      Left            =   11670
      Style           =   1  'Graphical
      TabIndex        =   29
      ToolTipText     =   "To Modify Record"
      Top             =   2115
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
      Left            =   11670
      Style           =   1  'Graphical
      TabIndex        =   30
      ToolTipText     =   "To Delete Record"
      Top             =   2565
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
      Left            =   13005
      Style           =   1  'Graphical
      TabIndex        =   32
      Top             =   1665
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
      Left            =   11700
      Style           =   1  'Graphical
      TabIndex        =   31
      ToolTipText     =   "To Save Record"
      Top             =   3240
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
      Left            =   13050
      Style           =   1  'Graphical
      TabIndex        =   34
      ToolTipText     =   "To Cancel Record"
      Top             =   2565
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
      Left            =   13050
      Style           =   1  'Graphical
      TabIndex        =   40
      ToolTipText     =   "Exit"
      Top             =   3195
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
      Left            =   15330
      Style           =   1  'Graphical
      TabIndex        =   46
      ToolTipText     =   "To Add New Record"
      Top             =   4425
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
      Left            =   15255
      Style           =   1  'Graphical
      TabIndex        =   47
      ToolTipText     =   "To Add New Record"
      Top             =   4860
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.TextBox txt 
      Height          =   345
      Index           =   3
      Left            =   15360
      TabIndex        =   50
      Text            =   "3"
      Top             =   1575
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.TextBox txt 
      Alignment       =   1  'Right Justify
      Height          =   360
      Index           =   5
      Left            =   15360
      TabIndex        =   43
      Text            =   "5"
      Top             =   2295
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
      Height          =   390
      Index           =   6
      Left            =   1560
      TabIndex        =   6
      Text            =   "6"
      Top             =   2625
      Width           =   2235
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
      Height          =   390
      Index           =   9
      Left            =   5250
      MaxLength       =   40
      TabIndex        =   13
      Text            =   "9"
      ToolTipText     =   "Press F3 for Creation & Press F2 for Selection"
      Top             =   3420
      Width           =   6060
   End
   Begin VB.TextBox txt 
      Height          =   325
      Index           =   7
      Left            =   15300
      TabIndex        =   41
      Text            =   "7"
      Top             =   2655
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.TextBox txt 
      CausesValidation=   0   'False
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
      Left            =   2145
      TabIndex        =   9
      Text            =   "11"
      Top             =   3165
      Width           =   1665
   End
   Begin VB.TextBox txt 
      Alignment       =   1  'Right Justify
      Height          =   360
      Index           =   4
      Left            =   15360
      TabIndex        =   42
      Text            =   "4"
      Top             =   1935
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.TextBox txt 
      Height          =   345
      Index           =   2
      Left            =   15360
      TabIndex        =   49
      Text            =   "2"
      Top             =   1215
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
      Left            =   15435
      TabIndex        =   48
      Text            =   "1"
      Top             =   720
      Visible         =   0   'False
      Width           =   555
   End
   Begin MSMask.MaskEdBox mskDt 
      Height          =   360
      Index           =   2
      Left            =   4365
      TabIndex        =   2
      Top             =   585
      Width           =   1560
      _ExtentX        =   2752
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
      Index           =   12
      Left            =   1530
      TabIndex        =   10
      Top             =   3570
      Width           =   1560
      _ExtentX        =   2752
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
      Left            =   8085
      TabIndex        =   3
      Top             =   600
      Width           =   1560
      _ExtentX        =   2752
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
      Height          =   4065
      Left            =   120
      TabIndex        =   15
      Top             =   4440
      Width           =   14205
      _ExtentX        =   25056
      _ExtentY        =   7170
      _Version        =   393216
      FixedCols       =   0
      RowHeightMin    =   275
      BackColorFixed  =   -2147483647
      ForeColorFixed  =   -2147483634
      BackColorBkg    =   12632256
      Appearance      =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
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
      Index           =   6
      Left            =   4230
      TabIndex        =   72
      Top             =   2655
      Width           =   1005
   End
   Begin VB.Shape Shape4 
      BorderColor     =   &H00C00000&
      BorderWidth     =   5
      FillColor       =   &H00C00000&
      Height          =   9480
      Left            =   0
      Top             =   -405
      Width           =   14505
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
      Left            =   11250
      TabIndex        =   58
      Top             =   0
      Width           =   2475
   End
   Begin VB.Label lbl1 
      BackStyle       =   0  'Transparent
      Caption         =   "Mill &Exp. Description :"
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
      Height          =   360
      Index           =   5
      Left            =   180
      TabIndex        =   22
      Top             =   8550
      Width           =   2190
   End
   Begin VB.Label lbl1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Entry &No.   :"
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
      Left            =   285
      TabIndex        =   0
      Top             =   585
      Width           =   1095
   End
   Begin VB.Label lbl1 
      BackStyle       =   0  'Transparent
      Caption         =   "Rcon.Date  :"
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
      Index           =   10
      Left            =   6810
      TabIndex        =   25
      Top             =   600
      Width           =   1335
   End
   Begin VB.Label lbl1 
      BackStyle       =   0  'Transparent
      Caption         =   "Amoun&t     :"
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
      Left            =   285
      TabIndex        =   23
      Top             =   2625
      Width           =   1215
   End
   Begin VB.Label lbl1 
      BackStyle       =   0  'Transparent
      Caption         =   "Date        :"
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
      Index           =   8
      Left            =   285
      TabIndex        =   26
      Top             =   3570
      Width           =   1095
   End
   Begin VB.Label lbl1 
      BackStyle       =   0  'Transparent
      Caption         =   "&To             :"
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
      Left            =   285
      TabIndex        =   21
      Top             =   1800
      Width           =   1215
   End
   Begin VB.Label lbl1 
      BackStyle       =   0  'Transparent
      Caption         =   "From         :"
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
      Left            =   285
      TabIndex        =   20
      Top             =   1365
      Width           =   1095
   End
   Begin VB.Label lbl1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Date   :"
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
      Left            =   3405
      TabIndex        =   19
      Top             =   585
      Width           =   675
   End
   Begin VB.Shape Shape1 
      FillColor       =   &H00BDDFD1&
      FillStyle       =   0  'Solid
      Height          =   3585
      Left            =   11520
      Shape           =   4  'Rounded Rectangle
      Top             =   465
      Width           =   2745
   End
   Begin VB.Label lbl1 
      BackStyle       =   0  'Transparent
      Caption         =   "&Narration :"
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
      Left            =   4230
      TabIndex        =   27
      Top             =   3060
      Width           =   960
   End
   Begin VB.Label lbl1 
      BackStyle       =   0  'Transparent
      Caption         =   "Ref No. :"
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
      Index           =   13
      Left            =   285
      TabIndex        =   24
      Top             =   3165
      Width           =   1095
   End
   Begin VB.Shape Shape2 
      BackColor       =   &H00808080&
      BorderColor     =   &H004D177D&
      FillColor       =   &H00F5F5DE&
      FillStyle       =   0  'Solid
      Height          =   615
      Left            =   105
      Top             =   465
      Width           =   11250
   End
   Begin VB.Label Label6 
      BackColor       =   &H00C00000&
      Caption         =   "  Bank Payment"
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
      Left            =   30
      TabIndex        =   59
      Top             =   0
      Width           =   14460
   End
   Begin VB.Shape Shape5 
      BackColor       =   &H00808080&
      BorderColor     =   &H004D177D&
      FillColor       =   &H80000005&
      FillStyle       =   0  'Solid
      Height          =   1155
      Left            =   135
      Top             =   1170
      Width           =   11250
   End
   Begin VB.Shape Shape6 
      BorderColor     =   &H00000080&
      FillColor       =   &H00F5F5DE&
      FillStyle       =   0  'Solid
      Height          =   1980
      Left            =   4095
      Top             =   2385
      Width           =   7305
   End
   Begin VB.Shape Shape7 
      BorderColor     =   &H00000080&
      FillColor       =   &H00ECFFFE&
      FillStyle       =   0  'Solid
      Height          =   1575
      Left            =   135
      Top             =   2475
      Width           =   3810
   End
End
Attribute VB_Name = "frmPaymentGST"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Const tblName As String = "tblVoucher"
Const OrderField As String = "Vdt,Vno,VCtrNo"
Const IdField As String = "vno"
Const MaxNo As Long = 37
Dim CriteriaStr As String
Dim rstbl As Recordset, ShowFrmFlag As Boolean
Dim varBookMark As Variant
Dim FirstTimeFlag As Boolean
Dim FormAction As Integer
Dim frmTypeStr As String
Dim rsRecVsSale As Recordset '-- TblRecVsSale
Dim oldFrmTypeStr As String
Dim i As Long
Dim IsMillVouPrint As Boolean
Dim LastBankAcName As String
Private Sub SetControlEd()
Dim AgCode As Long
If FormAction = vbDataActionAddNew Or FormAction = vbDataActionUpdate Then
      AgCode = GProcGetColumnValue("tblMastAccount", "AcName", txttemp(5), "S", "AcMillType", "N")
      If AgCode = 1 Or AgCode = 2 Then '--Mill consi/Depot
            txttemp(20).Enabled = True
      Else
            txttemp(20) = ""
            txttemp(20).Enabled = False
      End If
Else
    txttemp(20).Enabled = False
End If
txtDrBalance.Enabled = False
txtCrBalance.Enabled = False
End Sub
Private Sub cmdBtn_Click(Index As Integer)
'On Error GoTo ErrorRoutine
Dim FindQStr As String
Dim X As String
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
    FindQStr = "select Vno,vdt,TblMastAccount.AcName,TblAc1.AcName,Vamt, " & X & " where Vtype='" & frmTypeStr & "' and " _
    & " Vyear=" & gCYear & " and VCtrNo=1 and VAcCode=TblMastAccount.AcCode and VAcOCode=TblAc1.AcCode"
    gClsSearch.SearchMultiField FindQStr, "", Array("Vno", "Vdt", "Account", "Credit A/c", "Amount", "Narration"), Array(700, 1200, 2500, 2500, 1000, 2000), "", " ", "vno,vdt", 0, 0, False, , True
    If gClsSearch.SearchMultiRetCol(0) <> "" Then
        If GProcRstOpen(rstbl, "Select * from " & tblName & " where Vno = (Select max(" & gClsSearch.SearchMultiRetCol(0) & ") from " & tblName & " where Vno <> 0 and " & CriteriaStr & " )and  " & CriteriaStr & " order by " & OrderField & "  ", "O") > 0 Then ReadFields
        gClsSearch.SearchMultiRetCol(0) = ""
'        rstbl.MoveFirst
'        rstbl.Find "Vno = " & gClsSearch.SearchMultiRetCol(0)
'        gClsSearch.SearchMultiRetCol(0) = ""
'        ReadFields
    End If
Case 5, 9 'Print
    If Index = 5 Then
        IsMillVouPrint = True '-- Mill Voucher Print
    Else
        IsMillVouPrint = False
    End If
    Frame1.Visible = True
    DoEvents
    txtno(0).Enabled = True
    txtno(1).Enabled = True
    txtno(0) = txt(0)
    txtno(1) = txt(0)
    txtno(0).SetFocus
    chkIsChqPrint.Visible = True
    chkIsChqPrint.Enabled = True
    chkIsChqPrint.Value = 1
    Option1.Enabled = True
    Option2.Enabled = True
    Option1.Value = True
    Check1.Enabled = True
    Text1.Enabled = True
    Check2.Enabled = True
    
Case 6 'Add
    FormAction = vbDataActionAddNew
    Call GProcActivateControls(Me, True)
    Call GProcSetButtons(Me, vbDataActionAddNew)
    Call GProcClearForm(Me, rstbl, MaxNo, True)
    txt(0).text = GProcGenerateId(tblName, "VNO", CriteriaStr, gCn)
    mskDt(2) = GProcGenerateVDt(tblName, "VDt", CriteriaStr)
    SetControlEd
    If FirstTimeFlag = True Then
        FirstTimeFlag = False
    Else
        txt(0).SetFocus
    End If
    FillCombo
    SetGrid
    If frmTypeStr = "BP" Then txttemp(4) = LastBankAcName '-- Last Bank AcName
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
        SetControlEd
        mskDt(2).SetFocus
    End If
Case 8 'Delete
   ' If mURecDel = True Then
     If IsRelatedRecord = False Then
        If MsgBox("Do You Really Delete This Record?", vbQuestion + vbYesNo + vbDefaultButton1, "Confirm Delete") = vbYes Then
           gCn.Execute "Delete from tblvoucher where " & CriteriaStr & " and vno = " & txt(0)
            '--Update Outstanding Balance
           GProcUpdateOutstandingBal CLng(txt(0)), UCase(txt(3)), gCYear, False, True
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
     End If
   ' End If
Case 10 'Exit
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
        '-- Save Record
        txt(33) = 1  ' Entry In GST Payment
        GProcCreateVoucher FormAction, CLng(txt(0)), 1, CDate(txt(2)), frmTypeStr, CLng(txt(4)), CLng(txt(5)), Val(txt(6)), txt(11), txt(12), txt(13), txt(7), txt(8), txt(9), txt(10), False, txt(20), 0, Val(txt(21)), GSTEnt:=1, wIrnNo:=txt(37)
        GProcCreateVoucher FormAction, txt(0), 2, txt(2), frmTypeStr, txt(5), txt(4), Val(txt(6)) * (-1), txt(11), txt(12), txt(13), txt(7), txt(8), txt(9), txt(10), False, txt(20), 0, Val(txt(21)), GSTEnt:=1, wIrnNo:=txt(37)
        SaveRelatedRecord
        '--Update Outstanding Balance
        GProcUpdateOutstandingBal CLng(txt(0)), UCase(txt(3)), gCYear
        rstbl.Requery
        Call GProcActivateControls(Me, False)
        Call GProcSetButtons(Me, True)
        If FormAction = vbDataActionAddNew Then
            If cmdBtn(6).Visible = True Then cmdBtn(6).SetFocus
        Else
            If cmdBtn(7).Visible = True Then cmdBtn(7).SetFocus
        End If
        FormAction = vbDataActionClose
        msGrid.Col = 0
        msGrid.Row = 1
        txtGrid.Move msGrid.CellLeft + msGrid.Left, msGrid.CellTop + msGrid.Top, msGrid.CellWidth, msGrid.CellHeight
        LastBankAcName = txttemp(4)
    End If
Case 14 'Cancel
    If GProcRstOpen(rstbl, "Select * from " & tblName & " where " & CriteriaStr & "  and  " & IdField & " = (Select max(" & IdField & ") from " & tblName & " where vdt = (SELECT MAX(vdt) from " & tblName & " where " & CriteriaStr & " )  and  " & IdField & " <> 0 and  " & CriteriaStr & " ) order by " & OrderField & " ", "R") > 0 Then ReadFields Else GProcClearForm Me, rstbl, MaxNo
    txtGrid.Visible = False
    Call GProcClearForm(Me, rstbl, MaxNo, True)
    If rstbl.RecordCount <= 0 Then
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
    msGrid.Row = 1
    txtGrid.Move msGrid.CellLeft + msGrid.Left, msGrid.CellTop + msGrid.Top, msGrid.CellWidth, msGrid.CellHeight
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
    Vno = CLng(txtno(0))
    Do While Vno <= txtno(1)
     If Index = 0 Then GProcCrystalRptPreparation frmMain.cryReport, "Printer"
        If Index = 1 Then GProcCrystalRptPreparation frmMain.cryReport, "Window"
        X = "Select * From tblVoucher where  vNo=" & Vno & " and VCTRNO =1 and VType='" & frmTypeStr & "' and VYear=" & gCYear
        If GProcRstOpen(Rs1, X, "R") > 0 Then
            With frmMain.cryReport
                .SelectionFormula = "{tblVoucher.vNo} = " & Rs1!Vno & " and {tblVoucher.VCTRNO}=1 and {tblVoucher.VType}='" & Rs1!Vtype & "' and {tblVoucher.VYear}=" & Rs1!VYear
                .Formulas(5) = "wAmountInWords = '" & GProcAmountInWords(Abs(Rs1!vamt)) & "'"
                .WindowTitle = "Payment Voucher"
                If IsMillVouPrint = True Then  '-- Mill Payment Voucher Print
                    .ReportFileName = gReportPath & "rptVouMillPaymentPrint.rpt"
                    .SubreportToChange = "rptSubRptMillPaymentPrint.rpt"
                    If gBackEndDB = gBackEndAccess Then
                        .DataFiles(0) = gDbLocation
                    Else '-- Oracle
                       .Connect = "DSN=" + gODBCDSN + " ;UID=Scott;PWD=" + gODBCPwd + ";DBQ=" + gODBCHostStr
                    End If
                    .Action = 1
                Else
                
                    .SelectionFormula = "{tblVoucher.vNo} = " & Rs1!Vno & " and {tblVoucher.VCTRNO}=1 and {tblVoucher.VType}='" & Rs1!Vtype & "' and {tblVoucher.VYear}=" & Rs1!VYear
                    .Formulas(5) = "wAmountInWords = '" & GProcAmountInWords(Abs(Rs1!vamt)) & "'"
                    If Check1.Value = 1 Then
                       .Formulas(6) = "wnm = '1'"
                    ElseIf Check2.Value = 1 Then
                       .Formulas(6) = "wnm = '2'"
                    Else
                       .Formulas(6) = "wnm = '0'"
                    End If
                    
                    If chkIsChqPrint.Value = 1 Or Check1.Value = 1 Or Check2.Value = 1 Then
                        
                        If Text1.text <> "" Then
                           .Formulas(8) = "wprintnm = '" & Text1.text & "'"
                        End If
                        
                        If Option1.Value = True Then
                           .Formulas(7) = "wacpy = 'A/C PAYEE'"
                        Else
                           .Formulas(7) = "wacpy = ''"
                        End If
                        
                       .ReportFileName = gReportPath & "rptVouChequePrint.rpt"
                    End If
                    .Action = 1
                End If
'                    .SelectionFormula = "{tblVoucher.vNo} = " & Rs1!Vno & " and {tblVoucher.VCTRNO}=1 and {tblVoucher.VType}='" & Rs1!Vtype & "' and {tblVoucher.VYear}=" & Rs1!VYear
'                    .Formulas(5) = "wAmountInWords = '" & GProcAmountInWords(Abs(Rs1!vamt)) & "'"
'                    .ReportFileName = gReportPath & "rptVouPaymentPrint.rpt"
'                    .Action = 1
'                End If
            End With
        End If
        Vno = Vno + 1
    
''        If Index = 0 Then GProcCrystalRptPreparation frmMain.cryReport, "Printer"
''        If Index = 1 Then GProcCrystalRptPreparation frmMain.cryReport, "Window"
''        If IsMillVouPrint = True Then
''           X = "Select * From qryshortrecvoucher where  RecVNo=" & Vno & " and RecVType='" & frmTypeStr & "' and RecVYear=" & gCYear
''        Else
''          X = "Select * From tblVoucher where  vNo=" & Vno & " and VCTRNO =1 and VType='" & frmTypeStr & "' and VYear=" & gCYear
''        End If
''        If GProcRstOpen(Rs1, X, "R") > 0 Then
''            With frmMain.cryReport
''                 If IsMillVouPrint = True Then
''                    .SelectionFormula = "{qryshortrecvoucher.RecVno} = " & Rs1!RecVno & "  and {qryshortrecvoucher.RecVType}='" & Rs1!RecVType & "' and {qryshortrecvoucher.RecVYear}=" & Rs1!RecVYear
''                    .Formulas(5) = "wAmountInWord = '" & GProcAmountInWords(Abs(Rs1!RecAmt)) & "'"
''                    .Formulas(4) = "VoucherTitle " = "Bank Payment Voucher"
''                    .ReportFileName = gReportPath & "rptVouMillPayment.rpt"
''                    If gBackEndDB = gBackEndAccess Then
''                        .DataFiles(0) = gDbLocation
''                    Else '-- Oracle
''                       .Connect = "DSN=" + gODBCDSN + " ;UID=Scott;PWD=" + gODBCPwd + ";DBQ=" + gODBCHostStr
''                    End If
''                    .Action = 1
''                Else
''                    .SelectionFormula = "{tblVoucher.vNo} = " & Rs1!Vno & " and {tblVoucher.VCTRNO}=1 and {tblVoucher.VType}='" & Rs1!Vtype & "' and {tblVoucher.VYear}=" & Rs1!VYear
''                    .Formulas(5) = "wAmountInWords = '" & GProcAmountInWords(Abs(Rs1!vamt)) & "'"
''                    .ReportFileName = gReportPath & "rptVouPaymentPrint.rpt"
''                    .Action = 1
''                End If
''            End With
''        End If
''        Vno = Vno + 1
    Loop
Case 3:
    '------ Validate
    If IsNumeric(txtno(0)) = False Then txtno(0) = 0
    If IsNumeric(txtno(1)) = False Then txtno(1) = txtno(0)
    If CLng(txtno(1)) < CLng(txtno(0)) Then txtno(1) = txtno(0)
    '-------
    Vno = CLng(txtno(0))
    Do While Vno <= txtno(1)
     If Index = 0 Then GProcCrystalRptPreparation frmMain.cryReport, "Printer"
        If Index = 3 Then GProcCrystalRptPreparation frmMain.cryReport, "Window"
        X = "Select * From tblVoucher where  vNo=" & Vno & " and VCTRNO =1 and VType='" & frmTypeStr & "' and VYear=" & gCYear
        If GProcRstOpen(Rs1, X, "R") > 0 Then
            With frmMain.cryReport
                .WindowTitle = "Payment Voucher"
                .SelectionFormula = "{tblVoucher.vNo} = " & Rs1!Vno & " and {tblVoucher.VCTRNO}=1 and {tblVoucher.VType}='" & Rs1!Vtype & "' and {tblVoucher.VYear}=" & Rs1!VYear
                .Formulas(5) = "wAmountInWords = '" & GProcAmountInWords(Abs(Rs1!vamt)) & "'"
                '.Formulas(6) = "Wpanno = '" & gCPAN & "'"
                .ReportFileName = gReportPath & "rptRTGSPrint.rpt"
                .Action = 1
            End With
        End If
        Vno = Vno + 1
    Loop
Case 2:
    DoEvents
    Frame1.Visible = False
End Select
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
    Me.Top = 300
    List1.Clear
    List1.AddItem "Chq. "
    List1.AddItem "T.T. "
    List1.AddItem "D.D. "
    List1.AddItem "NEFT"
    List1.AddItem "Other"
    Me.Left = Screen.Width / 2 - Me.Width / 2
    frmTypeStr = gfrmTypeStr
    CriteriaStr = " VType = '" & frmTypeStr & "' and Vyear=" & gCYear
    Y = "Select * from " & tblName & " where vno = 0 and " & CriteriaStr & " and vctrno = 1"
    If Trim(OrderField) <> "" Then Y = Y & " order by " & OrderField
    GProcRstOpen rstbl, Y, "O"
    Select Case frmTypeStr
    Case "BP" '--Bank
    Case "CP" '-- Cash
        Label6.Caption = "Cash Payments"
        lbl1(10).Visible = True
        mskDt(13).Visible = True
        lbl1(10).Caption = "Date"
        mskDt(12).Visible = False
        txt(14).Enabled = True
        txt(14).Visible = True
        lbl1(8).Visible = False
        lbl1(13).Visible = False
        txttemp1(8).Visible = False
        txt(11).Visible = False
        txttemp1(7).MaxLength = 60
    End Select
    '---- Button Visiblity Set As Per User Previlage
    GProcSetButtonVisiblity Me
    '---Grid
    SetGrid
    FirstTimeFlag = True
    LastBankAcName = ""
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
If mskDt(12) = "__/__/____" Then mskDt(12) = mskDt(2)
If mskDt(13) = "__/__/____" Then mskDt(13) = mskDt(2)
SaveBtnEd
ShowAccountBalance
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
            GProcShowForm frmMastNarration, frmMain.mnuMstfrm(5), gfrmTypeStr
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
If (frmTypeStr = "CP") Then
   txttemp(4) = "Cash In Hand"
End If
End Sub
Private Sub SaveBtnEd()
'--Vno,Vdt,Amount,Cr A/c,Dr A/c
If txt(0) = "" Or GProcIsDateValid(mskDt(2)) = False Or CDbl(txt(6)) = 0 Or txttemp(4) = "" Or txttemp(5) = "" Then
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
   txt(7) = Left(txttemp1(8) + Space(5), 5) + " No. " + Left(txt(11) + Space(10), 10) + txttemp1(7)
Else
   txt(7) = txttemp1(7)
End If
txt(4) = GProcGetColumnValue("tblMastAccount", "AcName", txttemp(4), "S", "Accode", "N")
txt(5) = GProcGetColumnValue("tblMastAccount", "AcName", txttemp(5), "S", "AcCode", "N")
txt(20) = GProcGetColumnValue("tblMastNarration", "Narration", txttemp(20), "S", "NarrCode", "N")
txt(21) = chkDontShowInUnadj.Value
LastBankAcName = txttemp(4)
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
    txttemp1(8) = Left(txt(7), 5)
Else
    txttemp1(7) = txt(7)
End If
If IsNull(txt(21)) Then
   chkDontShowInUnadj.Value = 0
Else
   chkDontShowInUnadj.Value = Val(txt(21))
End If
txttemp(4) = GProcGetColumnValue("tblMastAccount", "AcCode", txt(4), "N", "AcName", "S")
txttemp(5) = GProcGetColumnValue("tblMastAccount", "AcCode", txt(5), "N", "AcName", "S")
txttemp(20) = GProcGetColumnValue("tblMastNarration", "NarrCode", txt(20), "N", "Narration", "S")
'-- Fill Related Record
FillRelatedRecord
SetControlEd
LastBankAcName = txttemp(4)
ShowAccountBalance
End Sub
Private Function ValidateData() As Boolean
Dim X As String
Dim GName As String
Dim rstAddCheck As Recordset
Dim MillAgCode As Long
GProcCheckForNumber Me, rstbl, MaxNo
SetControlEd
'---- Required Data
'--- Validate Grid
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
If txttemp(4) = "" Then
    MsgBox "Check Account.", vbInformation + vbOKOnly, Me.Caption
    txttemp(4).SetFocus
    Exit Function
End If
'--Create Dr A/c
If GProcGetColumnValue("tblMastAccount", "AcName", txttemp(4), "S", "AcCode", "N") = 0 Then
   GName = GProcGetColumnValue("TblMastGroup", "AgCode", 90019, "N", "AgName", "S")
   GProcShowForm frmMastAccount, frmMain.mnuMstfrm(0), Array(frmMastAccount.txt(1), frmMastAccount.txttemp(2)), Array(txttemp(4), GName)
   Exit Function
End If
'--- Cr A/c
If txttemp(5) = "" Then
    MsgBox "Check Account.", vbInformation + vbOKOnly, Me.Caption
    txttemp(5).SetFocus
    Exit Function
End If
'--Create Cr A/c
If GProcGetColumnValue("tblMastAccount", "AcName", txttemp(5), "S", "AcCode", "N") = 0 Then
   GName = GProcGetColumnValue("TblMastGroup", "AgCode", 90015, "N", "AgName", "S")
   GProcShowForm frmMastAccount, frmMain.mnuMstfrm(0), Array(frmMastAccount.txt(1), frmMastAccount.txttemp(2)), Array(txttemp(5), GName)
   Exit Function
End If
'--- Cr A/c ,Dr A/c
If txttemp(5) = txttemp(4) Then
    MsgBox "Check Account.Debit and Credit Accounts are Same.", vbInformation + vbOKOnly, Me.Caption
    txttemp(5).SetFocus
    Exit Function
End If
'---- Amount
If CDbl(txt(6)) <= 0 Then
    MsgBox "Check Voucher Amount.", vbInformation + vbOKOnly, Me.Caption
    Exit Function
End If
'---Mill Expenses
MillAgCode = GProcGetColumnValue("tblMastAccount", "AcName", txttemp(5), "S", "ACMILLTYPE", "N", " ACMILLTYPE<>0")
If txttemp(20) = "" And (MillAgCode = 1 Or MillAgCode = 2) Then
    MsgBox "Check Mill Expenses Description.", vbInformation + vbOKOnly, Me.Caption
    txttemp(20).SetFocus
    Exit Function
End If
If Trim(txttemp(20)) <> "" And GProcGetColumnValue("tblMastNarration", "Narration", txttemp(20), "S", "NarrCode", "N") = 0 Then
    oldFrmTypeStr = gfrmTypeStr
    gfrmTypeStr = "M"
    GProcShowForm frmMastNarration, frmMain.mnuMstfrm(7), Array(frmMastNarration.txt(2)), Array(txttemp(20))
    gfrmTypeStr = oldFrmTypeStr
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
        Case 4 '-- Dr A/c
         If frmTypeStr <> "CP" Then '-- Bank Payment
            GName = GProcGetColumnValue("TblMastGroup", "AgCode", 90019, "N", "AgName", "S")
            GProcShowForm frmMastAccount, frmMain.mnuMstfrm(0), Array(frmMastAccount.txttemp(2)), Array(GName)
         End If
        Case 5 '-- Cr A/c
            GName = GProcGetColumnValue("TblMastGroup", "AgCode", 90015, "N", "AgName", "S")
            GProcShowForm frmMastAccount, frmMain.mnuMstfrm(0), Array(frmMastAccount.txttemp(2)), Array(GName)
        Case 20 '-- Mill Exp Description
             oldFrmTypeStr = gfrmTypeStr
            gfrmTypeStr = "M"
            GProcShowForm frmMastNarration, frmMain.mnuMstfrm(7)
            gfrmTypeStr = oldFrmTypeStr
    End Select
End If
SetControlEd
End Sub
Private Sub txtTemp_KeyPress(Index As Integer, KeyAscii As Integer)
Dim LeftPos As Long
Dim TopPos As Long
TopPos = txttemp(Index).Top + Me.Top + 650
LeftPos = txttemp(Index).Left + Me.Left
If KeyAscii = 13 Then Exit Sub
'--Cash Payment
If (frmTypeStr = "CP") And (Index = 4) Then
   KeyAscii = 0
   txt(Index) = 2
   txttemp(Index) = GProcGetColumnValue("tblmastaccount", "Accode", 2, "N", "Acname", "S")
   Exit Sub
End If
Select Case Index
    Case 4: '--Bank payment
        'gClsSearch.SearchMultiField "tblMastAccount", "AcName,Accode", Array("Account Name", "Aid"), Array(txttemp(index).Width, 0), "AgCode=19 ", Chr(KeyAscii), "AcName", LeftPos, TopPos, True
        gClsSearch.SearchMultiField "QryMastaccount", "Account,AcName,AcCity,AgName", Array("Account", "AcName", "City", "Group"), Array(4000, 0, 2000, 3000), "AgCode=90019", Chr(KeyAscii), "AcName", LeftPos, TopPos, True
        KeyAscii = 0
        If gClsSearch.SearchMultiRetCol(0) <> "" And gClsSearch.SearchMultiRetCol(1) = "" Then
            txttemp(Index) = gClsSearch.SearchMultiRetCol(0)
        Else
            If gClsSearch.SearchMultiRetCol(1) <> "" Then txttemp(Index) = gClsSearch.SearchMultiRetCol(1)
        End If
    Case 5: '--Account
        gClsSearch.SearchMultiField "QryMastaccount", "Account,AcName,AcCity,AgName", Array("Account", "AcName", "City", "Group"), Array(4000, 0, 2000, 3000), "", Chr(KeyAscii), "AcName", LeftPos, TopPos, True
        KeyAscii = 0
        If gClsSearch.SearchMultiRetCol(0) <> "" And gClsSearch.SearchMultiRetCol(1) = "" Then
            txttemp(Index) = gClsSearch.SearchMultiRetCol(0)
        Else
            If gClsSearch.SearchMultiRetCol(1) <> "" Then txttemp(Index) = gClsSearch.SearchMultiRetCol(1)
        End If
    Case 20 '--Mill Expes. Narration
        gClsSearch.SearchMultiField "TblMastNarration", "Narration", Array("Mill Expenses Description"), Array(4000), "NarrType='M'", Chr(KeyAscii), "Narration", LeftPos, TopPos, True
        KeyAscii = 0
        If gClsSearch.SearchMultiRetCol(0) <> "" Then
            txttemp(Index) = gClsSearch.SearchMultiRetCol(0)
        End If
End Select
SetControlEd
End Sub

Private Sub txtTemp_LostFocus(Index As Integer)
SaveBtnEd
SetControlEd
ShowAccountBalance
End Sub

Private Sub txttemp1_GotFocus(Index As Integer)
If Index = 8 Then
   List1.Visible = True
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
    If txt(11).Enabled = True Then txt(11).SetFocus
End If
End Sub
Private Sub list1_LostFocus()
    Call list1_KeyPress(13)
    DoEvents
End Sub
Private Sub ShowAccountBalance()
Dim DrAcBal  As Double
Dim CrAcBal As Double
If IsDate(mskDt(2)) Then
    DrAcBal = GProcGetAccountBalance(txttemp(5), mskDt(2))
    CrAcBal = GProcGetAccountBalance(txttemp(4), mskDt(2))
    txtDrBalance = IIf(DrAcBal > 0, DrAcBal & " Cr", Abs(DrAcBal) & " Dr")
    txtCrBalance = IIf(CrAcBal > 0, CrAcBal & " Cr", Abs(CrAcBal) & " Dr")
End If
End Sub
'--- Related Record
Private Function IsRelatedRecord() As Boolean
'-- Audited
If GProcISAudited(txt(0), txt(3), txt(17)) = True Then
    MsgBox "Record is Audited.You can not Modify/Delete this Record.", vbCritical + vbOKOnly, App.Title
    IsRelatedRecord = True
    Exit Function
End If
If Val(txt(33)) = 0 Then
   MsgBox "Entry Done In VAT Payment. You can not Modify/Delete this Record.", vbCritical + vbOKOnly, App.Title
   IsRelatedRecord = True
   Exit Function
End If

'--- Related Record Payment
If rstbl.EOF And rstbl.BOF Then
Else
    If rstbl.Fields("VIsRefEntType") = 1 Then
        MsgBox "Related Record is Present.You can not Modify/Delete this Record.", vbCritical + vbOKOnly, App.Title
        IsRelatedRecord = True
        Exit Function
    End If
End If
End Function
Private Sub SetGrid()
Dim X As String
Set rsRecVsSale = New Recordset
'X = "select InvTp, InvBillNo, InvDt, BillAmt, AdjAmt, Discount, JvNoDisc, LateDays, Interest, IntReceived ,InvYear,InvNo,JVNOINT" _
'& " ,LessTDS, JvNoLessTDS,CGSTRt,SGSTRt,IGSTRt from tblRecVsSale where " _
'& " Recvtype='" & frmTypeStr & "' and Recvno=" & txt(0) & " and RecVYear=" & gCYear & " order by RecVCtrlNo"
'x1Rec = GProcRstOpen(rsRecVsSale, X, "R")
'With msGrid
'    .Clear
'    .FormatString = "<Type|>Inv. No      |<Inv. Date      |>Invoice Amt    |>Adj. Amount   |>Discount ||>Late |>Late Pay Int|>Int Recev.||||>Less TDS ||CGST Rt|SGST Rt|IGST Rt"




'--  InvTp,InvNo, InvDt, BillAmt,Discount,AddAMt,LessAmt, AdjAmt, LateDays, Interest, IntReceived ,InvYear
X = "select InvTp,InvBillNo,InvDt, BillAmt, AdjAmt,Discount,JvNoDisc, LateDays, Interest, IntReceived ,InvYear,InvNo,JVNOINT" _
& " ,LessTDS, JvNoLessTDS,CGSTRt,SGSTRt,IGSTRt from tblRecVsSale where " _
& " Recvtype='" & frmTypeStr & "' and Recvno=" & txt(0) & "  and RecVYear=" & gCYear & " order by RecVCtrlNo"
GProcRstOpen rsRecVsSale, X, "R"
With msGrid
    .Clear
    .FormatString = "<Type|>Bill No      |<Bill Date      |>Bill Amount    |>Adj. Amount   |>Discount ||>Late |>Late Pay Int|>Int. Paid.||||>Less TDS ||CGST Rt|SGST Rt|IGST Rt"
    .ColWidth(6) = 0
    .ColWidth(10) = 0
    .ColWidth(11) = 0
    .ColWidth(12) = 0
    .ColWidth(14) = 0 '-JV No. TDS
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
Private Sub MSGrid_LostFocus()
'CalAmount
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
            Case 0, 1, 2, 3, 6, 7, 8, 10, 11, 12, 14:  '--Type,Inv No, Inv Dt, Debit Note
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
        If .TextMatrix(.Row, 0) = "" And .Rows > 2 Then
            .RemoveItem (.Row)
            txtGrid.text = ""
        End If
        txtGrid.Visible = False
        If cmdBtn(13).Enabled = True Then
           cmdBtn(13).SetFocus
        End If
        Exit Sub
    End If
    '---- Navigate Row ,Col
    If KeyAscii = 13 Then
        KeyAscii = 0
        If .Col = rsRecVsSale.Fields.Count Then
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
    If .Col = 10 Then
        If Val(.TextMatrix(.Row, 9)) = 0 Then
           .Col = 17
        End If
    End If
    If KeyAscii <> 13 And KeyAscii <> 27 And .Col = 0 Then
        DoEvents
        KeyAscii = 0
        Pending_In
    End If
End With
'---- Set TxtGrid
If msGrid.text <> "" Then txtGrid.text = msGrid.text
If msGrid.Col < (rsRecVsSale.Fields.Count) Then
    If msGrid.Col = 6 Then msGrid.Col = 7
    If msGrid.Col = 10 Then msGrid.Col = 13
    If msGrid.Col = 10 Then msGrid.Col = 13
    If msGrid.Col = 7 Then msGrid.Col = 9
    If msGrid.Col = 14 Then msGrid.Col = 15
    txtGrid.Move msGrid.CellLeft + msGrid.Left, msGrid.CellTop + msGrid.Top, msGrid.CellWidth, msGrid.CellHeight
    msGrid.SetFocus
End If
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
'----Late Payment Interest Calcualtion
'---Interest Rate 1.5
Select Case msGrid.Col
Case 4:
    With msGrid
        .TextMatrix(.Row, 8) = Format(Round((((Val(.TextMatrix(.Row, 7)) * 1.5) / 3000) * Val(.TextMatrix(.Row, 8))), 2), "###0.00")
    End With
Case 7:
    'CalBalance
End Select
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
                Pending_In
                msGrid.Col = 4
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
                   MsgBox "Adjustment Amount is less than Invoice Amount", vbCritical + vbOKOnly, "Adjustment Amount is less than Invoice Amount"
                   DoEvents
                   txtGrid.SetFocus
                   Exit Sub
               End If
           Case 5:
               If Val(txtGrid) > Val(.TextMatrix(.Row, 3)) Then
                   MsgBox "Discount Amount must less than Invoice Amount", vbCritical + vbOKOnly, "Adjustment Dr Note Amount is less than Dr Note Amount"
                   DoEvents
                   txtGrid.SetFocus
                   Exit Sub
               End If
         End Select
    End If
    '--- Navigate Row ,Col
    If KeyAscii = 13 And .Col < rsRecVsSale.Fields.Count - 1 Then
        .Col = .Col + 1
    ElseIf KeyAscii = 13 And .Col = rsRecVsSale.Fields.Count - 1 Then
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
            If (Val(.TextMatrix(.Row, 4)) + Val(.TextMatrix(.Row, 9)) = 0) And .Rows > 2 Then
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
If KeyCode = vbKeyY And Shift = 2 Then
    msGrid = 4
    If msGrid.Rows > 2 Then
        msGrid.RemoveItem (msGrid.Row)
    Else
        SetGrid
    End If
    txtGrid.Visible = False
    txtGrid.text = ""
    'ReNumberGridSrno
    'CalAmount
    Exit Sub
End If
'--- Show Form
If KeyCode = vbKeyF3 Then
Select Case msGrid.Col
    Case 1 '-- Item
        'GProcShowForm frmMastItem, frmMain.mnuMstfrm(3), Array(frmMastItem.txttemp(9)), Array(txttemp(5))
        'Exit Sub
End Select
End If
End Sub
Private Sub CheckForNumberGrid()
Dim RowIndex As Long
Dim i As Long
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
'-- Grid Required
Private Function ValidateGrid() As Boolean
Dim TotAmt As Double
Dim i As Long
Dim RCount As Long
Dim j As Long
Dim InvNo As Long
Dim InVtype As String
Dim InvYear As Long
Dim BillDet As String
If Trim(txtGrid.text) <> "" Then
        msGrid.text = txtGrid.text
End If
txtGrid.Visible = False
txtGrid.text = ""
BillDet = ""
CheckForNumberGrid
With msGrid
    '-- Duplicate in selection
    For i = 1 To .Rows - 1
        If Trim(.TextMatrix(i, 1)) <> "" Then
           BillDet = BillDet + .TextMatrix(i, 1) + "-" + CStr(Val(.TextMatrix(i, 4)) + Val(.TextMatrix(i, 9))) + ":"
        End If
        InvNo = .TextMatrix(i, 11)
        InVtype = .TextMatrix(i, 0)
        InvYear = .TextMatrix(i, 10)
        For j = 1 To .Rows - 1
            If j <> i And InvNo = .TextMatrix(j, 11) And InVtype = .TextMatrix(j, 0) And InvYear = .TextMatrix(j, 10) And InvNo <> 0 Then
                MsgBox "Duplicate Inward Selection.", vbOKOnly, Me.Caption
                msGrid.SetFocus
                Exit Function
            End If
        Next j
    Next i
    If BillDet <> "" Then
        BillDet = Left(Trim(BillDet) + Space(120), 120)
        txt(8) = Left(BillDet, 40)
        txt(10) = Mid(BillDet, 41, 40)
'        txt(10) = Mid(BillDet, 81, 40)
    End If
    
    '-- Adjusted Amt
    If .Rows > 1 Then
       TotAmt = 0
       For i = 1 To .Rows - 1
           '-- Adj amt > Bill Amt
           If CDbl(.TextMatrix(i, 4)) - CDbl(.TextMatrix(i, 5)) > CDbl(.TextMatrix(i, 3)) Then
                MsgBox "Adjusted Amount is more than Balance Bill Amount. ", vbOKOnly, Me.Caption
                msGrid.SetFocus
                Exit Function
           End If
           '-- Check GST Rates
           If (Val(.TextMatrix(i, 9)) + Val(.TextMatrix(i, 5))) > 0 And (Val(.TextMatrix(i, 15)) + Val(.TextMatrix(i, 16)) + Val(.TextMatrix(i, 17))) = 0 Then
               MsgBox "Please Check GST Rates", vbOKOnly, Me.Caption
                msGrid.SetFocus
                Exit Function
           End If
           TotAmt = TotAmt + Val(.TextMatrix(i, 4)) - Val(.TextMatrix(i, 5)) + Val(.TextMatrix(i, 9)) - Val(.TextMatrix(i, 13))
        Next i
        If (TotAmt > 0) And (TotAmt <> Val(txt(6))) Then
           MsgBox "Total of Adjusted Amount " & TotAmt & " & Paid Amount not tally.", vbOKOnly, Me.Caption
           Exit Function
        End If
    End If
End With
ValidateGrid = True
End Function
Private Sub DeleteRelatedRecord()
'--- JV Int
gCn.Execute "delete from tblVoucher where Vtype='PN' and VNo in (select JVNOINT from tblRecVsSale where " _
& " INTRECEIVED>0 and RecVtype='" & frmTypeStr & "' and RecVNo=" & txt(0) & " and  RecVyear= " & gCYear & ")  and " _
 & " Vyear= " & gCYear

 '-- JV Less TDS
 gCn.Execute "delete from tblVoucher where Vtype='JV' and VNo in (select JVNOLessTDS from tblRecVsSale where " _
& " LessTDS>0 and RecVtype='" & frmTypeStr & "' and RecVNo=" & txt(0) & " and  RecVyear= " & gCYear & ")  and " _
 & " Vyear= " & gCYear

'---JV Discount
 gCn.Execute "delete from tblVoucher where Vtype='SN' and VNo in (select JvNoDisc from tblRecVsSale where " _
& " Discount>0 and RecVtype='" & frmTypeStr & "' and RecVNo=" & txt(0) & " and  RecVyear= " & gCYear & ")  and " _
 & " Vyear= " & gCYear


'--- Rec Vs Sale
gCn.Execute "delete from tblRecVsSale where RecVtype='" & frmTypeStr & "' and RecVNo=" & txt(0) & " and " _
 & " RecVyear= " & txt(17)
End Sub
Private Sub SaveRelatedRecord()
Dim i As Long, X As String
Dim rsRecVsSale As Recordset
Dim IntRecAmt As Double
Dim JVVNo As Long
Dim JVTDSVNo As Long '- Less TDS JV VNo
Dim JVDiscNo As Long '- Less Discount JV VNo
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
Dim rsOutst As Recordset, rstTmp As Recordset
Dim wBroker As String
Dim TaxAcCodeCGST As Long
Dim TaxAcCodeSGST As Long
Dim TaxAcCodeIGST As Long
Dim Nt2 As String
Dim Nt1 As String

Dim TaxAcCodeCGSTDisc As Long
Dim TaxAcCodeSGSTDisc As Long
Dim TaxAcCodeIGSTDisc As Long


TaxAcCodeCGST = GProcGetColumnValue("tblMastNarration", "Narration", "G S T", "S", "CGSTInPutAc", "N")
TaxAcCodeSGST = GProcGetColumnValue("tblMastNarration", "Narration", "G S T", "S", "SGSTInPutAc", "N")
TaxAcCodeIGST = GProcGetColumnValue("tblMastNarration", "Narration", "G S T", "S", "IGSTInPutAc", "N")



'-- Delete Records
DeleteRelatedRecord
'--- Save Grid Records
'--- TblRecVsSale
X = "select * from tblRecVsSale where Recvtype='" & frmTypeStr & "' and Recvno=" & txt(0) & " and " _
& " RecVYear=" & gCYear
GProcRstOpen rsRecVsSale, X, "O"
CheckForNumberGrid
With msGrid
    i = 1
    Do While i < .Rows
        If CDec(.TextMatrix(i, 7)) > 0 Or CDec(.TextMatrix(i, 10)) > 0 Then
        DoEvents
        rsRecVsSale.AddNew
        rsRecVsSale!AdjNo = GProcGenerateId("tblRecVsSale", "ADJNO", "")
        rsRecVsSale!RecVno = CLng(txt(0)) '--RecVNo
        rsRecVsSale!RecVType = UCase(txt(3)) '--RecVtype
        rsRecVsSale!RecVdt = CDate(txt(2)) '--RecVdt
        rsRecVsSale!RecVCtrlNo = i '--RecSrNo
        rsRecVsSale!RecVYear = gCYear '-- RecYear
        rsRecVsSale!RecAmt = CDbl(txt(6)) '--Rec Amt
        'rsRecVsSale!JVNoDisc
        'rsRecVsSale!JvNoIt
        'rsRecVsSale!Nar1 = Left(Trim(txttemp1(8)), 3) + ". No. " + Left(txt(19) + Space(10), 10) + " " + Left(txttemp1(7) + Space(20), 20)
        rsRecVsSale!InvBillno = .TextMatrix(i, 1) & vbNullString
'        rsRecVsSale!InvBillno = .TextMatrix(i, 1) '--Inv Bill No
        rsRecVsSale!InvNo = CLng(.TextMatrix(i, 11)) '--Invno
        rsRecVsSale!InvTp = .TextMatrix(i, 0) '--InvTp
        rsRecVsSale!InvDt = CDate(.TextMatrix(i, 2)) '--InvDt
        rsRecVsSale!InvYear = .TextMatrix(i, 10)
        rsRecVsSale!BillAmt = CDec(.TextMatrix(i, 3))
        rsRecVsSale!Discount = CDec(.TextMatrix(i, 5)) '- Discount
        rsRecVsSale!PartyCode = GProcGetColumnValue("tblMastAccount", "AcName", txttemp(5), "S", "AcCode", "N")
        rsRecVsSale!AdjAmt = CDec(.TextMatrix(i, 4))
        rsRecVsSale!LateDays = CInt(.TextMatrix(i, 7))
        rsRecVsSale!Interest = CDec(.TextMatrix(i, 8))
        rsRecVsSale!IntReceived = CDec(Val(.TextMatrix(i, 9)))
        IntRecAmt = CDec(Val(.TextMatrix(i, 9)))
        rsRecVsSale!LessTDS = CDec(Val(.TextMatrix(i, 13))) '- Less TDS
        LessTDSAmt = CDec(Val(.TextMatrix(i, 13)))
        LessDiscAmt = CDec(.TextMatrix(i, 5))
        rsRecVsSale!CGSTRt = Val(.TextMatrix(i, 15))
        rsRecVsSale!SGSTRt = Val(.TextMatrix(i, 16))
        rsRecVsSale!IGSTRt = Val(.TextMatrix(i, 17))
        JVVNo = 0
        JVTDSVNo = 0
        JVDiscNo = 0
        '------------------------- Interest Rec JV
        If FormAction = vbDataActionAddNew And IntRecAmt > 0 Then
            JVVNo = GProcGenerateId("tblVoucher", "Vno", " vyear=" & gCYear & " and Vtype='PN'")
        Else
            If IntRecAmt > 0 Then
                If CLng(.TextMatrix(i, 12)) > 0 And GProcGetColumnValue("tblVoucher", "Vno", .TextMatrix(i, 12), "N", "Vno", "N", " Vtype='PN' and Vyear=" & gCYear) = 0 Then
                    JVVNo = CLng(.TextMatrix(i, 12))
                Else
                    JVVNo = GProcGenerateId("tblVoucher", "Vno", " vyear=" & gCYear & " and Vtype='PN'")
                End If
            End If
        End If
        If IntRecAmt > 0 Then
           wTaxableAmt = Round(((IntRecAmt) * 100) / (100 + (Val(.TextMatrix(i, 15)) + Val(.TextMatrix(i, 16))) + Val(.TextMatrix(i, 17))), 2)
           wCgstAmt = Round((wTaxableAmt * Val(.TextMatrix(i, 15))) / 100, 2)
           wSgstAmt = Round((wTaxableAmt * Val(.TextMatrix(i, 16))) / 100, 2)
           wIgstAmt = Round((wTaxableAmt * Val(.TextMatrix(i, 17))) / 100, 2)
           NetDbAmt = (IntRecAmt) - (wCgstAmt + wSgstAmt + wIgstAmt)
           Nt1 = Left("Int. Paid Bill-" & .TextMatrix(i, 1) & "-" & CDate(.TextMatrix(i, 2)), 40)
           If LessTDSAmt > 0 Then
              Nt2 = "Less : T D S Rs " + CStr(LessTDSAmt)
           Else
              Nt2 = ""
           End If
           
           GProcCreateVoucher FormAction, JVVNo, 1, mskDt(2), "PN", txt(5), gLatePayIntAcCode, (IntRecAmt), .TextMatrix(i, 1), CDate(.TextMatrix(i, 2)), , Nar1:=Nt1, Nar2:=Nt2, IsRefEntryType:=True, Taxable:=wTaxableAmt, CRt:=Val(.TextMatrix(i, 15)), Camt:=wCgstAmt, SRt:=Val(.TextMatrix(i, 16)), Samt:=wSgstAmt, IRt:=Val(.TextMatrix(i, 17)), Iamt:=wIgstAmt, AgtVno:=CLng(.TextMatrix(i, 11)), AgtVtype:=.TextMatrix(i, 0), AgtVyear:=.TextMatrix(i, 10)
           GProcCreateVoucher FormAction, JVVNo, 2, mskDt(2), "PN", gLatePayIntAcCode, txt(5), -NetDbAmt, .TextMatrix(i, 1), CDate(.TextMatrix(i, 2)), , Nar1:=Nt1, Nar2:=Nt2, IsRefEntryType:=True, Taxable:=wTaxableAmt, CRt:=Val(.TextMatrix(i, 15)), Camt:=wCgstAmt, SRt:=Val(.TextMatrix(i, 16)), Samt:=wSgstAmt, IRt:=Val(.TextMatrix(i, 17)), Iamt:=wIgstAmt, AgtVno:=CLng(.TextMatrix(i, 11)), AgtVtype:=.TextMatrix(i, 0), AgtVyear:=.TextMatrix(i, 10)
           If wCgstAmt > 0 Then
              GProcCreateVoucher FormAction, JVVNo, 3, mskDt(2), "PN", TaxAcCodeCGST, txt(5), -(wCgstAmt), .TextMatrix(i, 1), CDate(.TextMatrix(i, 2)), , "", "", "", "", IsRefEntryType:=True, Taxable:=wTaxableAmt, CRt:=Val(.TextMatrix(i, 15)), Camt:=wCgstAmt, SRt:=Val(.TextMatrix(i, 16)), Samt:=wSgstAmt, IRt:=Val(.TextMatrix(i, 17)), Iamt:=wIgstAmt, AgtVno:=CLng(.TextMatrix(i, 11)), AgtVtype:=.TextMatrix(i, 0), AgtVyear:=.TextMatrix(i, 10)
           End If
           If wSgstAmt > 0 Then
              GProcCreateVoucher FormAction, JVVNo, 4, mskDt(2), "PN", TaxAcCodeSGST, txt(5), -(wSgstAmt), .TextMatrix(i, 1), CDate(.TextMatrix(i, 2)), , "", "", "", "", IsRefEntryType:=True, Taxable:=wTaxableAmt, CRt:=Val(.TextMatrix(i, 15)), Camt:=wCgstAmt, SRt:=Val(.TextMatrix(i, 16)), Samt:=wSgstAmt, IRt:=Val(.TextMatrix(i, 17)), Iamt:=wIgstAmt, AgtVno:=CLng(.TextMatrix(i, 11)), AgtVtype:=.TextMatrix(i, 0), AgtVyear:=.TextMatrix(i, 10)
           End If
           If wIgstAmt > 0 Then
              GProcCreateVoucher FormAction, JVVNo, 5, mskDt(2), "PN", TaxAcCodeIGST, txt(5), -(wIgstAmt), .TextMatrix(i, 1), CDate(.TextMatrix(i, 2)), , "", "", "", "", IsRefEntryType:=True, Taxable:=wTaxableAmt, CRt:=Val(.TextMatrix(i, 15)), Camt:=wCgstAmt, SRt:=Val(.TextMatrix(i, 16)), Samt:=wSgstAmt, IRt:=Val(.TextMatrix(i, 17)), Iamt:=wIgstAmt, AgtVno:=CLng(.TextMatrix(i, 11)), AgtVtype:=.TextMatrix(i, 0), AgtVyear:=.TextMatrix(i, 10)
           End If
        End If
        '-------------------------------
'        '------------------------- Less TDS  JV
        If FormAction = vbDataActionAddNew And LessTDSAmt > 0 Then
            JVTDSVNo = GProcGenerateId("tblVoucher", "Vno", " vyear=" & gCYear & " and Vtype='JV'")
        Else
            If LessTDSAmt > 0 Then
                If CLng(.TextMatrix(i, 14)) > 0 And GProcGetColumnValue("tblVoucher", "Vno", .TextMatrix(i, 14), "N", "Vno", "N", " Vtype='JV' and Vyear=" & gCYear) = 0 Then
                    JVTDSVNo = CLng(.TextMatrix(i, 14))
                Else
                    JVTDSVNo = GProcGenerateId("tblVoucher", "Vno", " vyear=" & gCYear & " and Vtype='JV'")
                End If
            End If
        End If
        If LessTDSAmt > 0 Then
            GProcCreateVoucher FormAction, JVTDSVNo, 1, mskDt(2), "JV", gTDSPaidAcCode, txt(5), LessTDSAmt, IsRefEntryType:=True, Nar1:="TDS on Interest " & IntRecAmt, Nar2:="Agt Bill-" & .TextMatrix(i, 1) & "-" & CDate(.TextMatrix(i, 2))
            GProcCreateVoucher FormAction, JVTDSVNo, 2, mskDt(2), "JV", txt(5), gTDSPaidAcCode, -LessTDSAmt, IsRefEntryType:=True, Nar1:="TDS on Interest " & IntRecAmt, Nar2:="Agt Bill-" & .TextMatrix(i, 1) & "-" & CDate(.TextMatrix(i, 2))
        End If
        '------------------------- Less Disc  JV
        If FormAction = vbDataActionAddNew And LessDiscAmt > 0 Then
            JVDiscNo = GProcGenerateId("tblVoucher", "Vno", " vyear=" & gCYear & " and Vtype='SN'")
        Else
            If LessDiscAmt > 0 Then
                If CLng(.TextMatrix(i, 6)) > 0 And GProcGetColumnValue("tblVoucher", "Vno", .TextMatrix(i, 6), "N", "Vno", "N", " Vtype='SN' and Vyear=" & gCYear) = 0 Then
                    JVDiscNo = CLng(.TextMatrix(i, 6))
                Else
                    JVDiscNo = GProcGenerateId("tblVoucher", "Vno", " vyear=" & gCYear & " and Vtype='SN'")
                End If
            End If
        End If
        If LessDiscAmt > 0 Then
           TaxAcCodeCGSTDisc = GProcGetColumnValue("tblMastNarration", "Narration", "G S T", "S", "CGSTPayAc", "N")
           TaxAcCodeSGSTDisc = GProcGetColumnValue("tblMastNarration", "Narration", "G S T", "S", "SGSTPayAc", "N")
           TaxAcCodeIGSTDisc = GProcGetColumnValue("tblMastNarration", "Narration", "G S T", "S", "IGSTPayAc", "N")
           wTaxableAmt = Round(((LessDiscAmt) * 100) / (100 + (Val(.TextMatrix(i, 15)) + Val(.TextMatrix(i, 16))) + Val(.TextMatrix(i, 17))), 2)
           wCgstAmt = Round((wTaxableAmt * Val(.TextMatrix(i, 15))) / 100, 2)
           wSgstAmt = Round((wTaxableAmt * Val(.TextMatrix(i, 16))) / 100, 2)
           wIgstAmt = Round((wTaxableAmt * Val(.TextMatrix(i, 17))) / 100, 2)
           NetDbAmt = (LessDiscAmt) - (wCgstAmt + wSgstAmt + wIgstAmt)
           Nt2 = ""
           Nt1 = Left("Int. Paid Agt. Bill-" & .TextMatrix(i, 1) & "-" & CDate(.TextMatrix(i, 2)), 40)
           GProcCreateVoucher FormAction, JVDiscNo, 1, mskDt(2), "SN", txt(5), gBillDbNtAcCode, -LessDiscAmt, .TextMatrix(i, 1), CDate(.TextMatrix(i, 2)), , Nar1:=Nt1, Nar2:=Nt2, IsRefEntryType:=True, Taxable:=wTaxableAmt, CRt:=Val(.TextMatrix(i, 15)), Camt:=wCgstAmt, SRt:=Val(.TextMatrix(i, 16)), Samt:=wSgstAmt, IRt:=Val(.TextMatrix(i, 17)), Iamt:=wIgstAmt, AgtVno:=CLng(.TextMatrix(i, 11)), AgtVtype:=.TextMatrix(i, 0), AgtVyear:=.TextMatrix(i, 10)
           GProcCreateVoucher FormAction, JVDiscNo, 2, mskDt(2), "SN", gBillDbNtAcCode, txt(5), (NetDbAmt), .TextMatrix(i, 1), CDate(.TextMatrix(i, 2)), , Nar1:=Nt1, Nar2:=Nt2, IsRefEntryType:=True, Taxable:=wTaxableAmt, CRt:=Val(.TextMatrix(i, 15)), Camt:=wCgstAmt, SRt:=Val(.TextMatrix(i, 16)), Samt:=wSgstAmt, IRt:=Val(.TextMatrix(i, 17)), Iamt:=wIgstAmt, AgtVno:=CLng(.TextMatrix(i, 11)), AgtVtype:=.TextMatrix(i, 0), AgtVyear:=.TextMatrix(i, 10)
           If wCgstAmt > 0 Then
              GProcCreateVoucher FormAction, JVDiscNo, 3, mskDt(2), "SN", TaxAcCodeCGSTDisc, txt(5), (wCgstAmt), .TextMatrix(i, 1), CDate(.TextMatrix(i, 2)), , "", "", "", "", IsRefEntryType:=True, Taxable:=wTaxableAmt, CRt:=Val(.TextMatrix(i, 15)), Camt:=wCgstAmt, SRt:=Val(.TextMatrix(i, 16)), Samt:=wSgstAmt, IRt:=Val(.TextMatrix(i, 17)), Iamt:=wIgstAmt, AgtVno:=CLng(.TextMatrix(i, 11)), AgtVtype:=.TextMatrix(i, 0), AgtVyear:=.TextMatrix(i, 10)
           End If
           If wSgstAmt > 0 Then
              GProcCreateVoucher FormAction, JVDiscNo, 4, mskDt(2), "SN", TaxAcCodeSGSTDisc, txt(5), (wSgstAmt), .TextMatrix(i, 1), CDate(.TextMatrix(i, 2)), , "", "", "", "", IsRefEntryType:=True, Taxable:=wTaxableAmt, CRt:=Val(.TextMatrix(i, 15)), Camt:=wCgstAmt, SRt:=Val(.TextMatrix(i, 16)), Samt:=wSgstAmt, IRt:=Val(.TextMatrix(i, 17)), Iamt:=wIgstAmt, AgtVno:=CLng(.TextMatrix(i, 11)), AgtVtype:=.TextMatrix(i, 0), AgtVyear:=.TextMatrix(i, 10)
           End If
           If wIgstAmt > 0 Then
              GProcCreateVoucher FormAction, JVDiscNo, 5, mskDt(2), "SN", TaxAcCodeIGSTDisc, txt(5), (wIgstAmt), .TextMatrix(i, 1), CDate(.TextMatrix(i, 2)), , "", "", "", "", IsRefEntryType:=True, Taxable:=wTaxableAmt, CRt:=Val(.TextMatrix(i, 15)), Camt:=wCgstAmt, SRt:=Val(.TextMatrix(i, 16)), Samt:=wSgstAmt, IRt:=Val(.TextMatrix(i, 17)), Iamt:=wIgstAmt, AgtVno:=CLng(.TextMatrix(i, 11)), AgtVtype:=.TextMatrix(i, 0), AgtVyear:=.TextMatrix(i, 10)
           End If
        End If
        '-------------------------------
        rsRecVsSale!JvNoLessTDS = JVTDSVNo '-- TDS JV No
        rsRecVsSale!JVNOINT = JVVNo
        rsRecVsSale!JvNoDisc = JVDiscNo
        rsRecVsSale.Update
        DoEvents
        End If
        i = i + 1
     Loop
 rsRecVsSale.Close
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
Private Sub Pending_In()
Dim wFromCd As Long, mItem
Dim rstOutStanding As Recordset
Dim LateDayCd As String
Dim X As String
Dim DateCd As String
Dim vamt As Double
Dim i As Long
Dim j As Long
Dim TotCol(0, 1) As Long

If Trim(txttemp(5)) <> "" Then
    wFromCd = GProcGetColumnValue("tblMastAccount", "AcName", txttemp(5), "S", "AcCode", "N")
    If gBackEndDB = gBackEndAccess Then
    Else '-- Oracle
        DateCd = " and vdt<= to_date('" & CDate(mskDt(2)) & "','dd/MM/yyyy')"
    End If
    If gBackEndDB = gBackEndAccess Then
        LateDayCd = "cdate('" & (CDate(mskDt(2)) + 3) & "') " & "- cdate(outDueDate)"
        X = "select VType, outbillNo,format(vdt,'dd/mm/yyyy'),OutAmount, OutAmount-OutRecAmt , " & LateDayCd & ", VYear,Vno from tblOutstanding where OutAcCode=" & wFromCd & DateCd & " and (OutAmount-OutRecAmt) > 0 "
'        X = "select VType, outbillNo,format(vdt,'dd/mm/yyyy'),iif(vtype='SM',OutAmountMillBill,OutAmount), iif(vtype='SM',OutAmountMillBill-OutRecAmtMillBill,OutAmount-OutRecAmt) , " & LateDayCd & ", VYear,Vno from tblOutstanding where  iif(vtype='SM',OutMillCode=" & wFromCd & DateCd & ",OutAcCode=" & wFromCd & DateCd & ") and iif(vtype='SM',(OutAmountMillBill-OutRecAmtMillBill) > 0,(OutAmount-OutRecAmt) > 0) "
    Else
        LateDayCd = "to_date('" & CDate(mskDt(2)) + 3 & "','dd/MM/yyyy') - outDueDate"
         X = "select VType,outbillNo,Vdt,OutAmount-OutRecAmt , OutAmount-OutRecAmt , " & LateDayCd & ", VYear,Vno from tblOutstanding where OutAcCode=" & wFromCd & DateCd & " and (OutAmount-OutRecAmt) > 0 "
    End If
    TotCol(0, 0) = 4
    TotCol(0, 1) = CDbl(txt(6))
    vamt = CDbl(txt(6))
    gClsSearch.SearchMultiSelField X, "", Array("In.Type", "Bill .No", "Date", "Amount", "Balance Amt", "Due Date", "Year", "Vno"), Array(900, 900, 1200, 1500, 1500, 1000, 0, 1000), "", "", "Vdt,Vno", 0, 0, , True, TotOfSelCols:=TotCol
    With msGrid
        If gClsSearch.SearchMultiSelRetCol(0, 0) <> "" Then
             j = .Row
             txtGrid = gClsSearch.SearchMultiSelRetCol(0, 0)
            For i = 0 To UBound(gClsSearch.SearchMultiSelRetCol)
                If .Rows - 1 = j Then
                Else
                    .Rows = .Rows + 1
                End If
                .TextMatrix(j, 0) = gClsSearch.SearchMultiSelRetCol(i, 0) '--Vtype
                .TextMatrix(j, 1) = gClsSearch.SearchMultiSelRetCol(i, 1) '-Bill No
                .TextMatrix(j, 11) = gClsSearch.SearchMultiSelRetCol(i, 7) '-VNo
                .TextMatrix(j, 2) = Format(gClsSearch.SearchMultiSelRetCol(i, 2), "dd/mm/yyyy") '--Vdt
                .TextMatrix(j, 3) = gClsSearch.SearchMultiSelRetCol(i, 3) '--Amt
                .TextMatrix(j, 4) = gClsSearch.SearchMultiSelRetCol(i, 4) '--Amt
                .TextMatrix(j, 7) = gClsSearch.SearchMultiSelRetCol(i, 5) '-- Late Days
                .TextMatrix(j, 10) = gClsSearch.SearchMultiSelRetCol(i, 6) '--Year
                j = j + 1
            Next i
                For i = 1 To .Rows - 1
                    If vamt <= CDbl(.TextMatrix(i, 4)) Then  '- Adj Amt
                        .TextMatrix(i, 4) = vamt
                        .TextMatrix(i, 8) = Format(Round((((Val(.TextMatrix(i, 4)) * 1.5) / 3000) * Val(.TextMatrix(i, 7))), 2), "###0.00")
                        vamt = 0
                        Exit For
                    Else
                        .TextMatrix(i, 4) = CDbl(.TextMatrix(i, 4))
                        .TextMatrix(i, 8) = Format(Round((((Val(.TextMatrix(i, 4)) * 1.5) / 3000) * Val(.TextMatrix(i, 7))), 2), "###0.00")
                        vamt = vamt - CDbl(.TextMatrix(i, 4))
                    End If
                Next i
                .Col = 4
    End If
    End With
Else
    MsgBox "Select From Account Compulsary", vbCritical + vbOKOnly, "Select From"
    DoEvents
    txttemp(5).SetFocus
End If
End Sub
Private Sub CalGridAmt()
Dim X As String
Dim ItTp As Long
Dim Stat As String
Dim rstTmp As Recordset
If Val(msGrid.TextMatrix(msGrid.Row, 11)) > 0 Then
    X = "Select OutItemCode,ItType from tblOutstanding,tblMastItem where vno = " & msGrid.TextMatrix(msGrid.Row, 11) & " and VType = '" & msGrid.TextMatrix(msGrid.Row, 0) & "' and Vyear = " & msGrid.TextMatrix(msGrid.Row, 10) & "" _
        & " and tblOutStanding.OutItemCode = tblMastItem.Itcode"
    i = GProcRstOpen(rstTmp, X, "R")
    If i > 0 Then
       If Val(rstTmp.Fields(0)) > 0 And FormAction = vbDataActionAddNew Then
            ItTp = rstTmp.Fields(1)
            Stat = GProcGetColumnValue("tblMastAccount", "AcName", txttemp(4), "S", "AcGSTIN", "S")
            If ItTp = 0 Then
               If Left(Stat, 2) = "" Or Left(Stat, 1) = "U" Or Left(Stat, 2) = "27" Then
                  msGrid.TextMatrix(msGrid.Row, 15) = GProcGetColumnValue("tblMastNarration", "Narration", "G S T", "S", "CotCGSTRt", "N")
                  msGrid.TextMatrix(msGrid.Row, 16) = GProcGetColumnValue("tblMastNarration", "Narration", "G S T", "S", "CotSGSTRt", "N")
               Else
                  msGrid.TextMatrix(msGrid.Row, 17) = GProcGetColumnValue("tblMastNarration", "Narration", "G S T", "S", "CotIGSTRt", "N")
               End If
            Else
               If Left(Stat, 2) = "" Or Left(Stat, 1) = "U" Or Left(Stat, 2) = "27" Then
                  msGrid.TextMatrix(msGrid.Row, 15) = GProcGetColumnValue("tblMastNarration", "Narration", "G S T", "S", "PolCGSTRt", "N")
                  msGrid.TextMatrix(msGrid.Row, 16) = GProcGetColumnValue("tblMastNarration", "Narration", "G S T", "S", "PolSGSTRt", "N")
               Else
                  msGrid.TextMatrix(msGrid.Row, 17) = GProcGetColumnValue("tblMastNarration", "Narration", "G S T", "S", "PolIGSTRt", "N")
               End If
            End If
      End If
    End If
End If
End Sub





