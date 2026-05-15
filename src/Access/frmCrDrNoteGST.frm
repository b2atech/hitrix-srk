VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form frmCrDrNoteGST 
   Appearance      =   0  'Flat
   AutoRedraw      =   -1  'True
   BackColor       =   &H00C2E3FC&
   BorderStyle     =   0  'None
   Caption         =   "Journal Entry"
   ClientHeight    =   9180
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   14310
   FillColor       =   &H00C00000&
   ForeColor       =   &H00C00000&
   Icon            =   "frmCrDrNoteGST.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   9180
   ScaleWidth      =   14310
   ShowInTaskbar   =   0   'False
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
      Index           =   39
      Left            =   3735
      MaxLength       =   12
      TabIndex        =   104
      Text            =   "39"
      Top             =   5760
      Width           =   1410
   End
   Begin VB.TextBox txtIRN 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Index           =   6
      Left            =   0
      MultiLine       =   -1  'True
      TabIndex        =   103
      Text            =   "frmCrDrNoteGST.frx":058A
      ToolTipText     =   "Press F3 for Creation"
      Top             =   0
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
      Height          =   3015
      Left            =   3510
      TabIndex        =   58
      Top             =   2520
      Visible         =   0   'False
      Width           =   5715
      Begin VB.CommandButton cmdPrint 
         Caption         =   "e-Invoice Screen"
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
         Index           =   6
         Left            =   135
         TabIndex        =   94
         Top             =   2250
         Width           =   1890
      End
      Begin VB.CommandButton cmdPrint 
         Caption         =   "e-Invoice"
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
         Left            =   2130
         TabIndex        =   93
         Top             =   2250
         Width           =   1575
      End
      Begin VB.CommandButton cmdPrint 
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
         Left            =   3840
         TabIndex        =   92
         Top             =   2250
         Width           =   1800
      End
      Begin VB.TextBox txtno 
         Appearance      =   0  'Flat
         Height          =   360
         Index           =   0
         Left            =   1950
         MaxLength       =   7
         TabIndex        =   63
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
         TabIndex        =   62
         Top             =   1560
         Width           =   1335
      End
      Begin VB.TextBox txtno 
         Appearance      =   0  'Flat
         Height          =   360
         Index           =   1
         Left            =   1950
         MaxLength       =   7
         TabIndex        =   61
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
         TabIndex        =   60
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
         TabIndex        =   59
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
         TabIndex        =   65
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
         TabIndex        =   64
         Top             =   1080
         Width           =   1050
      End
   End
   Begin VB.TextBox txtIRN 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   645
      Index           =   3
      Left            =   1530
      MaxLength       =   64
      MultiLine       =   -1  'True
      TabIndex        =   100
      Text            =   "frmCrDrNoteGST.frx":0596
      ToolTipText     =   "Press F3 for Creation"
      Top             =   4455
      Width           =   3630
   End
   Begin VB.TextBox txtIRN 
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
      Index           =   4
      Left            =   1530
      MaxLength       =   40
      TabIndex        =   99
      Text            =   "txtIRN(4)"
      ToolTipText     =   "Press F3 for Creation"
      Top             =   5265
      Width           =   3630
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
      Left            =   90
      MaxLength       =   40
      TabIndex        =   96
      Text            =   "38"
      ToolTipText     =   "Press F3 for Creation & Press F2 for Selection"
      Top             =   8685
      Visible         =   0   'False
      Width           =   360
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
      Left            =   90
      MaxLength       =   40
      TabIndex        =   95
      Text            =   "37"
      ToolTipText     =   "Press F3 for Creation & Press F2 for Selection"
      Top             =   7425
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
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Index           =   36
      Left            =   0
      TabIndex        =   90
      Text            =   "36"
      Top             =   0
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
      Height          =   405
      Index           =   36
      Left            =   1620
      MaxLength       =   40
      TabIndex        =   5
      Text            =   "txttemp(36)"
      ToolTipText     =   "Press F3 for Creation"
      Top             =   2655
      Width           =   5010
   End
   Begin VB.CheckBox chkIsLastYearBill 
      BackColor       =   &H00F5F5DE&
      Caption         =   "Is Last year Bill ?"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   10350
      TabIndex        =   88
      Top             =   1575
      Width           =   1905
   End
   Begin VB.TextBox txt 
      Height          =   345
      Index           =   35
      Left            =   17550
      TabIndex        =   87
      Text            =   "35"
      Top             =   2655
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.TextBox txt 
      Height          =   345
      Index           =   33
      Left            =   17550
      TabIndex        =   86
      Text            =   "33"
      Top             =   2295
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.TextBox txt 
      Height          =   345
      Index           =   32
      Left            =   17505
      TabIndex        =   85
      Text            =   "32"
      Top             =   1890
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.TextBox txt 
      Height          =   345
      Index           =   31
      Left            =   17550
      TabIndex        =   84
      Text            =   "31"
      Top             =   1440
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.TextBox txt 
      Height          =   345
      Index           =   30
      Left            =   17550
      TabIndex        =   83
      Text            =   "30"
      Top             =   1035
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.TextBox txt 
      Height          =   345
      Index           =   22
      Left            =   17595
      TabIndex        =   82
      Text            =   "22"
      Top             =   585
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.TextBox txt 
      Height          =   345
      Index           =   21
      Left            =   17550
      TabIndex        =   81
      Text            =   "21"
      Top             =   225
      Visible         =   0   'False
      Width           =   495
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
      Height          =   390
      Index           =   34
      Left            =   5940
      TabIndex        =   2
      Text            =   "34"
      Top             =   675
      Width           =   1800
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
      Index           =   29
      Left            =   9945
      TabIndex        =   15
      Text            =   "29"
      Top             =   5535
      Width           =   1935
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
      Index           =   28
      Left            =   6705
      TabIndex        =   12
      Text            =   "28"
      Top             =   5535
      Width           =   855
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
      Index           =   27
      Left            =   9945
      TabIndex        =   14
      Text            =   "27"
      Top             =   5130
      Width           =   1935
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
      Index           =   26
      Left            =   6705
      TabIndex        =   11
      Text            =   "26"
      Top             =   5130
      Width           =   855
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
      Index           =   25
      Left            =   9945
      TabIndex        =   13
      Text            =   "25"
      Top             =   4725
      Width           =   1935
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
      Index           =   24
      Left            =   6705
      TabIndex        =   10
      Text            =   "24"
      Top             =   4725
      Width           =   855
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
      Index           =   23
      Left            =   9945
      TabIndex        =   9
      Text            =   "23"
      Top             =   4050
      Width           =   1935
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
      Left            =   14400
      TabIndex        =   23
      Top             =   8280
      Visible         =   0   'False
      Width           =   885
   End
   Begin VB.TextBox txt 
      Height          =   345
      Index           =   20
      Left            =   15030
      TabIndex        =   57
      Text            =   "20"
      Top             =   6885
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.TextBox txt 
      Height          =   345
      Index           =   19
      Left            =   15075
      TabIndex        =   56
      Text            =   "19"
      Top             =   7380
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.TextBox txt 
      Height          =   345
      Index           =   18
      Left            =   15075
      TabIndex        =   55
      Text            =   "18"
      Top             =   7785
      Visible         =   0   'False
      Width           =   495
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
      Index           =   20
      Left            =   3090
      MaxLength       =   40
      TabIndex        =   20
      Text            =   "txttemp(20)"
      ToolTipText     =   "Press F3 for Creation"
      Top             =   8400
      Visible         =   0   'False
      Width           =   7290
   End
   Begin VB.TextBox txt 
      Height          =   345
      Index           =   17
      Left            =   15030
      TabIndex        =   54
      Text            =   "17"
      Top             =   6075
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.TextBox txt 
      Height          =   345
      Index           =   16
      Left            =   15075
      TabIndex        =   53
      Text            =   "16"
      Top             =   5355
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
      Height          =   390
      Index           =   8
      Left            =   3090
      MaxLength       =   40
      TabIndex        =   18
      Text            =   "8"
      ToolTipText     =   "Press F3 for Creation & Press F2 for Selection"
      Top             =   7575
      Width           =   7290
   End
   Begin VB.TextBox txt 
      Height          =   345
      Index           =   15
      Left            =   14940
      TabIndex        =   52
      Text            =   "15"
      Top             =   6435
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.TextBox txt 
      Height          =   375
      Index           =   12
      Left            =   15105
      TabIndex        =   51
      Text            =   "12"
      Top             =   2940
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.TextBox txt 
      Height          =   325
      Index           =   10
      Left            =   15075
      TabIndex        =   50
      Text            =   "10"
      Top             =   2655
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
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Index           =   0
      Left            =   2295
      TabIndex        =   1
      Text            =   "0"
      Top             =   705
      Width           =   1575
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
      Left            =   1530
      MaxLength       =   40
      TabIndex        =   6
      Text            =   "txttemp(5)"
      ToolTipText     =   "Press F3 for Creation"
      Top             =   3825
      Width           =   5850
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
      Left            =   1620
      MaxLength       =   40
      TabIndex        =   4
      Text            =   "txttemp(4)"
      ToolTipText     =   "Press F3 for Creation"
      Top             =   1710
      Width           =   6525
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
      Height          =   375
      Index           =   14
      Left            =   1800
      TabIndex        =   16
      Text            =   "14"
      Top             =   6345
      Width           =   990
   End
   Begin VB.TextBox txt 
      Alignment       =   1  'Right Justify
      Height          =   360
      Index           =   13
      Left            =   15075
      TabIndex        =   43
      Text            =   "13"
      Top             =   3375
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
      Left            =   12870
      Style           =   1  'Graphical
      TabIndex        =   34
      ToolTipText     =   "Previous"
      Top             =   1065
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
      Left            =   13380
      Style           =   1  'Graphical
      TabIndex        =   35
      ToolTipText     =   "Next"
      Top             =   1065
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
      Left            =   12870
      Style           =   1  'Graphical
      TabIndex        =   36
      ToolTipText     =   "First"
      Top             =   1515
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
      Left            =   13380
      Style           =   1  'Graphical
      TabIndex        =   37
      ToolTipText     =   "Last"
      Top             =   1515
      Width           =   450
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
      Left            =   12870
      Style           =   1  'Graphical
      TabIndex        =   33
      ToolTipText     =   "To Find Record"
      Top             =   2520
      Width           =   960
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
      Left            =   12870
      Style           =   1  'Graphical
      TabIndex        =   38
      Top             =   2070
      Visible         =   0   'False
      Width           =   960
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
      Left            =   12870
      Style           =   1  'Graphical
      TabIndex        =   28
      ToolTipText     =   "To Add New Record"
      Top             =   3465
      Width           =   960
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
      Left            =   12870
      Style           =   1  'Graphical
      TabIndex        =   29
      ToolTipText     =   "To Modify Record"
      Top             =   3960
      Width           =   960
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
      Left            =   12870
      Style           =   1  'Graphical
      TabIndex        =   30
      ToolTipText     =   "To Delete Record"
      Top             =   4455
      Width           =   960
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
      Left            =   12870
      Style           =   1  'Graphical
      TabIndex        =   39
      Top             =   5715
      Width           =   960
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
      Left            =   12870
      Style           =   1  'Graphical
      TabIndex        =   31
      ToolTipText     =   "To Save Record"
      Top             =   6210
      Width           =   960
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
      Left            =   12870
      Style           =   1  'Graphical
      TabIndex        =   32
      ToolTipText     =   "To Cancel Record"
      Top             =   6705
      Width           =   960
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
      Left            =   12870
      Style           =   1  'Graphical
      TabIndex        =   40
      ToolTipText     =   "Exit"
      Top             =   7875
      Width           =   960
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
      Left            =   14580
      Style           =   1  'Graphical
      TabIndex        =   45
      ToolTipText     =   "To Add New Record"
      Top             =   4275
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
      Left            =   14805
      Style           =   1  'Graphical
      TabIndex        =   46
      ToolTipText     =   "To Add New Record"
      Top             =   4860
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.TextBox txt 
      Height          =   345
      Index           =   3
      Left            =   15195
      TabIndex        =   49
      Text            =   "3"
      Top             =   1335
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.TextBox txt 
      Alignment       =   1  'Right Justify
      Height          =   360
      Index           =   5
      Left            =   15195
      TabIndex        =   42
      Text            =   "5"
      Top             =   2055
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
      Left            =   9945
      TabIndex        =   21
      Text            =   "6"
      Top             =   6390
      Width           =   1935
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
      Left            =   3090
      MaxLength       =   40
      TabIndex        =   19
      Text            =   "9"
      ToolTipText     =   "Press F3 for Creation & Press F2 for Selection"
      Top             =   7965
      Width           =   7290
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
      Index           =   7
      Left            =   3090
      MaxLength       =   40
      TabIndex        =   17
      Text            =   "7"
      ToolTipText     =   "Press F3 for Creation & Press F2 for Selection"
      Top             =   7185
      Width           =   7290
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
      Index           =   11
      Left            =   10395
      TabIndex        =   7
      Text            =   "11"
      Top             =   2115
      Width           =   1890
   End
   Begin VB.TextBox txt 
      Alignment       =   1  'Right Justify
      Height          =   360
      Index           =   4
      Left            =   15195
      TabIndex        =   41
      Text            =   "4"
      Top             =   1695
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.TextBox txt 
      Height          =   345
      Index           =   2
      Left            =   15195
      TabIndex        =   48
      Text            =   "2"
      Top             =   975
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
      Left            =   15075
      TabIndex        =   47
      Text            =   "1"
      Top             =   615
      Visible         =   0   'False
      Width           =   555
   End
   Begin MSMask.MaskEdBox mskDt 
      Height          =   390
      Index           =   2
      Left            =   10350
      TabIndex        =   3
      Top             =   705
      Width           =   1890
      _ExtentX        =   3334
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
   Begin Crystal.CrystalReport cryReport 
      Bindings        =   "frmCrDrNoteGST.frx":05A0
      Left            =   14355
      Top             =   225
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
      Height          =   1695
      Left            =   14400
      TabIndex        =   22
      Top             =   7605
      Visible         =   0   'False
      Width           =   870
      _ExtentX        =   1535
      _ExtentY        =   2990
      _Version        =   393216
      FixedCols       =   0
      RowHeightMin    =   275
      BackColorFixed  =   -2147483647
      ForeColorFixed  =   -2147483634
      BackColorBkg    =   12632256
      Appearance      =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSMask.MaskEdBox mskDt 
      Height          =   390
      Index           =   12
      Left            =   10395
      TabIndex        =   8
      Top             =   2655
      Width           =   1890
      _ExtentX        =   3334
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
   Begin MSMask.MaskEdBox mskDtIRN 
      Height          =   390
      Index           =   5
      Left            =   1530
      TabIndex        =   101
      Top             =   5745
      Width           =   1470
      _ExtentX        =   2593
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
   Begin VB.Label lbl1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "HSN  : "
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
      Index           =   13
      Left            =   3060
      TabIndex        =   105
      Top             =   5805
      Width           =   645
   End
   Begin VB.Label lbl1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Ack. Dt.  : "
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
      Index           =   12
      Left            =   450
      TabIndex        =   102
      Top             =   5805
      Width           =   960
   End
   Begin VB.Label lbl1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Ack. No. : "
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
      Index           =   10
      Left            =   450
      TabIndex        =   98
      Top             =   5355
      Width           =   960
   End
   Begin VB.Label lbl1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "IRN No.   : "
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
      Index           =   9
      Left            =   450
      TabIndex        =   97
      Top             =   4455
      Width           =   1095
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "TCS  Rate   :"
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
      Left            =   495
      TabIndex        =   91
      Top             =   6390
      Width           =   1305
   End
   Begin VB.Label lbl1 
      BackStyle       =   0  'Transparent
      Caption         =   "Broker      :"
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
      Height          =   495
      Index           =   8
      Left            =   360
      TabIndex        =   89
      Top             =   2715
      Width           =   1230
   End
   Begin VB.Label lblDrCrNote 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Voucher &No.    "
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
      Left            =   4095
      TabIndex        =   80
      Top             =   765
      Width           =   1410
   End
   Begin VB.Label lblGstin 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Gstin"
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
      Left            =   1620
      TabIndex        =   79
      Top             =   2250
      Width           =   480
   End
   Begin VB.Label lbl1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "GSTIN     : "
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
      Index           =   6
      Left            =   450
      TabIndex        =   78
      Top             =   2250
      Width           =   1020
   End
   Begin VB.Label Label9 
      BackStyle       =   0  'Transparent
      Caption         =   "Total Amount            :"
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
      Left            =   7605
      TabIndex        =   77
      Top             =   6435
      Width           =   2175
   End
   Begin VB.Label lbl1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Invoice  Date     :"
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
      Height          =   240
      Index           =   7
      Left            =   8595
      TabIndex        =   76
      Top             =   2700
      Width           =   1605
   End
   Begin VB.Label lbl1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Against Invoice  :"
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
      Index           =   2
      Left            =   8595
      TabIndex        =   75
      Top             =   2160
      Width           =   1635
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FF00FF&
      BorderWidth     =   2
      X1              =   225
      X2              =   12330
      Y1              =   6255
      Y2              =   6255
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "I  GST Amount        :"
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
      Left            =   7830
      TabIndex        =   74
      Top             =   5520
      Width           =   2055
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "I  GST Rate  :"
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
      Left            =   5355
      TabIndex        =   73
      Top             =   5520
      Width           =   1260
   End
   Begin VB.Shape Shape4 
      BorderColor     =   &H00C00000&
      BorderWidth     =   5
      FillColor       =   &H00C00000&
      Height          =   9555
      Left            =   0
      Top             =   -405
      Width           =   14250
   End
   Begin VB.Label Label17 
      BackStyle       =   0  'Transparent
      Caption         =   "Taxable Amount       :"
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
      Left            =   7605
      TabIndex        =   72
      Top             =   4050
      Width           =   2175
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "C GST Rate  :"
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
      Left            =   5355
      TabIndex        =   71
      Top             =   4710
      Width           =   1305
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "C GST Amount        :"
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
      Left            =   7830
      TabIndex        =   70
      Top             =   4710
      Width           =   2055
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "S GST Rate  :"
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
      Left            =   5355
      TabIndex        =   69
      Top             =   5115
      Width           =   1260
   End
   Begin VB.Label Label8 
      BackStyle       =   0  'Transparent
      Caption         =   "S GST Amount        :"
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
      Left            =   7830
      TabIndex        =   68
      Top             =   5115
      Width           =   2055
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
      Left            =   11340
      TabIndex        =   67
      Top             =   -45
      Width           =   2475
   End
   Begin VB.Label Label6 
      BackColor       =   &H00C00000&
      Caption         =   "  JV"
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
      Left            =   0
      TabIndex        =   66
      Top             =   0
      Width           =   14235
   End
   Begin VB.Label lbl1 
      BackStyle       =   0  'Transparent
      Caption         =   "Mill &Exp. Description   :"
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
      Height          =   495
      Index           =   5
      Left            =   450
      TabIndex        =   27
      Top             =   8400
      Visible         =   0   'False
      Width           =   2190
   End
   Begin VB.Label lbl1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Voucher &No.    "
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
      Index           =   0
      Left            =   495
      TabIndex        =   0
      Top             =   765
      Width           =   1410
   End
   Begin VB.Label lbl1 
      BackStyle       =   0  'Transparent
      Caption         =   "Debit       :"
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
      ForeColor       =   &H00800000&
      Height          =   375
      Index           =   4
      Left            =   450
      TabIndex        =   26
      Top             =   3825
      Width           =   1320
   End
   Begin VB.Label lbl1 
      BackStyle       =   0  'Transparent
      Caption         =   "Credit      :"
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
      ForeColor       =   &H00800000&
      Height          =   375
      Index           =   3
      Left            =   450
      TabIndex        =   25
      Top             =   1740
      Width           =   1095
   End
   Begin VB.Label lbl1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Da&te                  :"
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
      Index           =   1
      Left            =   8550
      TabIndex        =   24
      Top             =   750
      Width           =   1575
   End
   Begin VB.Shape Shape1 
      FillColor       =   &H00BDDFD1&
      FillStyle       =   0  'Solid
      Height          =   8235
      Left            =   12645
      Shape           =   4  'Rounded Rectangle
      Top             =   585
      Width           =   1365
   End
   Begin VB.Label lbl1 
      BackStyle       =   0  'Transparent
      Caption         =   "Narration                   :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   240
      Index           =   11
      Left            =   450
      TabIndex        =   44
      Top             =   7185
      Width           =   2265
   End
   Begin VB.Shape Shape3 
      BackColor       =   &H00808080&
      BorderColor     =   &H00C00000&
      FillColor       =   &H00ECFFFF&
      FillStyle       =   0  'Solid
      Height          =   1935
      Left            =   225
      Top             =   7020
      Width           =   12135
   End
   Begin VB.Shape Shape5 
      BackColor       =   &H00808080&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00800000&
      FillColor       =   &H00F5F5DE&
      FillStyle       =   0  'Solid
      Height          =   3285
      Left            =   225
      Top             =   3555
      Width           =   12135
   End
   Begin VB.Shape Shape2 
      BackColor       =   &H00808080&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00800000&
      FillColor       =   &H00ECFFFF&
      FillStyle       =   0  'Solid
      Height          =   1800
      Left            =   225
      Top             =   1395
      Width           =   8085
   End
   Begin VB.Shape Shape7 
      BackColor       =   &H00808080&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00800000&
      FillColor       =   &H00F5F5DE&
      FillStyle       =   0  'Solid
      Height          =   675
      Left            =   225
      Top             =   540
      Width           =   12180
   End
   Begin VB.Shape Shape6 
      BackColor       =   &H00808080&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00800000&
      FillColor       =   &H00F5F5DE&
      FillStyle       =   0  'Solid
      Height          =   1800
      Left            =   8415
      Top             =   1395
      Width           =   3990
   End
End
Attribute VB_Name = "frmCrDrNoteGST"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Const tblName As String = "tblVoucher"
Const OrderField As String = "Vdt,Vno"
Const IdField As String = "vno"
Const MaxNo As Long = 39
Dim CriteriaStr As String
Dim rstbl As Recordset, ShowFrmFlag As Boolean
Dim varBookMark As Variant
Dim FindFlag As Boolean
Dim FirstTimeFlag As Boolean
Dim FormAction As Integer
Dim frmTypeStr As String
Dim TaxAmt As Double
Dim wCotCGStRt As Double
Dim wCotSGStRt As Double
Dim wCotIGStRt As Double
Dim wPolCGStRt As Double
Dim wPolSGStRt As Double
Dim wPolIGStRt As Double
Dim i As Long
Dim rsOS As Recordset '--tbleInvoiceIRN
'-- RecVsSale
Dim rsRecVsSale As Recordset
Private Sub chkIsLastYearBill_LostFocus()
SetControlEd
End Sub

Private Sub cmdBtn_Click(Index As Integer)
Dim FindQStr As String
Dim X As String
Dim BillSr As String
Dim wBrokerCd As Long
Dim wMillCd As Long
Dim Rs1 As Recordset
'On Error GoTo ErrorRoutine
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
Case 4 'Find
    If gBackEndDB = gBackEndAccess Then
        X = " vNar1 & ' ' & VNar2 from " _
    & " tblvoucher,TblMastAccount,TblMastAccount as TblAc1 "
    Else
        X = " vNar1 || ' ' || VNar2 from " _
    & " tblvoucher,TblMastAccount,TblMastAccount TblAc1 "
    End If
    FindQStr = "select Vno,VDRCRNTNO,Vdt,TblMastAccount.AcName,TblAc1.AcName,abs(Vamt)," & X & " where Vtype='" & frmTypeStr & "' and " _
    & " Vyear=" & gCYear & " and VCtrNo=1 and VAcCode=TblMastAccount.AcCode and VAcOCode=TblAc1.AcCode and (VDrCrNtType =  '" & gfrmHunkStr & "' )"
    gClsSearch.SearchMultiField FindQStr, "", Array("Vno", "Doc.No", "Vdt", "Credit A/c", "Debit A/c", "Amount", "Narration"), Array(700, 1200, 1200, 2500, 2500, 1000, 2000), "", " ", "vno,vdt", 0, 0, False, , True
    If gClsSearch.SearchMultiRetCol(0) <> "" Then
        rstbl.MoveFirst
        rstbl.Find "Vno = " & gClsSearch.SearchMultiRetCol(0)
        gClsSearch.SearchMultiRetCol(0) = ""
        ReadFields
    End If
Case 5 'List
    'GenMastList ("LAC")
Case 6 'Add
     txt(22) = "A"
     FormAction = vbDataActionAddNew
     Call GProcActivateControls(Me, True)
'     lbl1(3).Enabled = True
'     lbl1(4).Enabled = True
     Call GProcSetButtons(Me, vbDataActionAddNew)
     Call GProcClearForm(Me, rstbl, MaxNo, True)
     txt(0).text = GProcGenerateId(tblName, "VNo", CriteriaStr, gCn)
     Select Case frmTypeStr
     Case "SN"
          If gfrmHunkStr = "B" Then
             BillSr = "BD"
          ElseIf gfrmHunkStr = "T" Then
             BillSr = "TC"
          Else
             BillSr = "SD"
          End If
     Case "PN"
          If gfrmHunkStr = "B" Then
             BillSr = "BC"
          Else
             BillSr = "SC"
          End If
     End Select
     txt(34) = GProcGenerateCrDrNtNo(frmTypeStr, gfrmHunkStr, BillSr)
     mskDt(2) = GProcGenerateVDt(tblName, "VDT", CriteriaStr)
     SetControlEd
     If FirstTimeFlag = True Then
        FirstTimeFlag = False
     Else
        txt(0).SetFocus
     End If
     FillCombo
     'SetGrid
Case 7 'Modify
'    Call GProcChkPass(frmTypeStr, FormAction, txt(34), mskDt(2), txtTemp(4), gUserName, "")
'    If gChkPassTrueFalse = True Then
'       txt(22) = "M"
        If IsRelatedRecord = False Then
            FormAction = vbDataActionUpdate
            Call GProcActivateControls(Me, True)
            Call GProcSetButtons(Me, vbDataActionUpdate)
            SetControlEd
            txt(0).Enabled = False
            txt(34).Enabled = False
             '--Update Outstanding Balance
            GProcUpdateOutstandingBal CLng(txt(0)), UCase(txt(3)), gCYear, True
            mskDt(2).SetFocus
        End If
'    Else
 '      FormAction = vbDataActionClose
 '   End If
     txtIRN(3).Enabled = True
     txtIRN(4).Enabled = True
     mskDtIRN(5).Enabled = True
     Call GProcSetButtons(Me, vbDataActionUpdate)
Case 8 'Delete
   ' If mURecDel = True Then
    'Call GProcChkPass("DRCRNOTE", FormAction, txt(0), mskDt(2), txtTemp(4), gUserName, "")
    'If gChkPassTrueFalse = True Then
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
     End If
    'End If
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
        If gfrmHunkStr = "B" Then
            txt(35) = "B"
        ElseIf gfrmHunkStr = "T" Then
            txt(35) = "T"
        Else
            txt(35) = "S"
        End If
        gCn.Execute "Delete from tblvoucher where " & CriteriaStr & " and vno = " & txt(0)
        Dim TaxAcCodeCGST As Long
        Dim TaxAcCodeSGST As Long
        Dim TaxAcCodeIGST As Long
        txt(13) = ""
        '--- Account Code
        TaxAcCodeCGST = GProcGetColumnValue("tblMastNarration", "Narration", "G S T", "S", "CGSTInPutAc", "N")
        TaxAcCodeSGST = GProcGetColumnValue("tblMastNarration", "Narration", "G S T", "S", "SGSTInPutAc", "N")
        TaxAcCodeIGST = GProcGetColumnValue("tblMastNarration", "Narration", "G S T", "S", "IGSTInPutAc", "N")
        
        If lbl1(3) = "Credit" Then
            TaxAmt = Val(txt(25)) + Val(txt(27)) + Val(txt(29))
            GProcCreateVoucher FormAction, txt(0), 1, txt(2), frmTypeStr, txt(4), txt(5), Val(txt(6)), txt(11), txt(12), txt(13), txt(7), txt(8), txt(9), txt(10), txt(19), txt(20), Taxable:=Val(txt(23)), Crt:=Val(txt(24)), Camt:=Val(txt(25)), Srt:=Val(txt(26)), Samt:=Val(txt(27)), Irt:=Val(txt(28)), Iamt:=Val(txt(29)), AgtVno:=CLng(txt(30)), AgtVtype:=txt(31), AgtVyear:=CLng(txt(32)), GSTEnt:=0, DrCrNtNo:=txt(34), DrCrNtType:=txt(35), Ramt:=txt(14), wBroker:=txt(36), HsnCd:=txt(39)
            GProcCreateVoucher FormAction, txt(0), 2, txt(2), frmTypeStr, txt(5), txt(4), (Val(txt(6)) - TaxAmt) * (-1), txt(11), txt(12), txt(13), txt(7), txt(8), txt(9), txt(10), txt(19), txt(20), Taxable:=Val(txt(23)), Crt:=Val(txt(24)), Camt:=Val(txt(25)), Srt:=Val(txt(26)), Samt:=Val(txt(27)), Irt:=Val(txt(28)), Iamt:=Val(txt(29)), AgtVno:=CLng(txt(30)), AgtVtype:=txt(31), AgtVyear:=CLng(txt(32)), GSTEnt:=0, DrCrNtNo:=txt(34), DrCrNtType:=txt(35), Ramt:=txt(14), wBroker:=txt(36), HsnCd:=txt(39)
            i = 3
            If Val(txt(25)) > 0 Then
               GProcCreateVoucher FormAction, txt(0), i, txt(2), frmTypeStr, TaxAcCodeCGST, txt(4), Val(txt(25)) * -1, txt(11), txt(12), txt(13), txt(7), txt(8), txt(9), txt(10), txt(19), txt(20), Taxable:=Val(txt(23)), Crt:=Val(txt(24)), Camt:=Val(txt(25)), Srt:=Val(txt(26)), Samt:=Val(txt(27)), Irt:=Val(txt(28)), Iamt:=Val(txt(29)), AgtVno:=CLng(txt(30)), AgtVtype:=txt(31), AgtVyear:=CLng(txt(32)), GSTEnt:=0, DrCrNtNo:=txt(34), DrCrNtType:=txt(35)
               i = i + 1
               DoEvents
            End If
            If Val(txt(27)) > 0 Then
               GProcCreateVoucher FormAction, txt(0), i, txt(2), frmTypeStr, TaxAcCodeSGST, txt(4), Val(txt(27)) * -1, txt(11), txt(12), txt(13), txt(7), txt(8), txt(9), txt(10), txt(19), txt(20), Taxable:=Val(txt(23)), Crt:=Val(txt(24)), Camt:=Val(txt(25)), Srt:=Val(txt(26)), Samt:=Val(txt(27)), Irt:=Val(txt(28)), Iamt:=Val(txt(29)), AgtVno:=CLng(txt(30)), AgtVtype:=txt(31), AgtVyear:=CLng(txt(32)), GSTEnt:=0, DrCrNtNo:=txt(34), DrCrNtType:=txt(35)
               i = i + 1
               DoEvents
            End If
            If Val(txt(29)) > 0 Then
               GProcCreateVoucher FormAction, txt(0), i, txt(2), frmTypeStr, TaxAcCodeIGST, txt(4), Val(txt(29)) * -1, txt(11), txt(12), txt(13), txt(7), txt(8), txt(9), txt(10), txt(19), txt(20), Taxable:=Val(txt(23)), Crt:=Val(txt(24)), Camt:=Val(txt(25)), Srt:=Val(txt(26)), Samt:=Val(txt(27)), Irt:=Val(txt(28)), Iamt:=Val(txt(29)), AgtVno:=CLng(txt(30)), AgtVtype:=txt(31), AgtVyear:=CLng(txt(32)), GSTEnt:=0, DrCrNtNo:=txt(34), DrCrNtType:=txt(35)
               i = i + 1
               DoEvents
            End If
        Else
            TaxAmt = Val(txt(25)) + Val(txt(27)) + Val(txt(29))
            GProcCreateVoucher FormAction, txt(0), 1, txt(2), frmTypeStr, txt(4), txt(5), Val(txt(6)) * (-1), txt(11), txt(12), txt(13), txt(7), txt(8), txt(9), txt(10), txt(19), txt(20), Taxable:=Val(txt(23)), Crt:=Val(txt(24)), Camt:=Val(txt(25)), Srt:=Val(txt(26)), Samt:=Val(txt(27)), Irt:=Val(txt(28)), Iamt:=Val(txt(29)), AgtVno:=CLng(txt(30)), AgtVtype:=txt(31), AgtVyear:=CLng(txt(32)), GSTEnt:=0, DrCrNtNo:=txt(34), DrCrNtType:=txt(35), Ramt:=txt(14), wBroker:=txt(36), HsnCd:=txt(39)
            GProcCreateVoucher FormAction, txt(0), 2, txt(2), frmTypeStr, txt(5), txt(4), Val(txt(6)) - TaxAmt, txt(11), txt(12), txt(13), txt(7), txt(8), txt(9), txt(10), txt(19), txt(20), Taxable:=Val(txt(23)), Crt:=Val(txt(24)), Camt:=Val(txt(25)), Srt:=Val(txt(26)), Samt:=Val(txt(27)), Irt:=Val(txt(28)), Iamt:=Val(txt(29)), AgtVno:=CLng(txt(30)), AgtVtype:=txt(31), AgtVyear:=CLng(txt(32)), GSTEnt:=0, DrCrNtNo:=txt(34), DrCrNtType:=txt(35), Ramt:=txt(14), wBroker:=txt(36), HsnCd:=txt(39)
            i = 3
            If Val(txt(25)) > 0 Then
               GProcCreateVoucher FormAction, txt(0), i, txt(2), frmTypeStr, TaxAcCodeCGST, txt(4), Val(txt(25)), txt(11), txt(12), txt(13), txt(7), txt(8), txt(9), txt(10), txt(19), txt(20), Taxable:=Val(txt(23)), Crt:=Val(txt(24)), Camt:=Val(txt(25)), Srt:=Val(txt(26)), Samt:=Val(txt(27)), Irt:=Val(txt(28)), Iamt:=Val(txt(29)), AgtVno:=CLng(txt(30)), AgtVtype:=txt(31), AgtVyear:=CLng(txt(32)), GSTEnt:=0, DrCrNtNo:=txt(34), DrCrNtType:=txt(35)
               i = i + 1
               DoEvents
            End If
            If Val(txt(27)) > 0 Then
               GProcCreateVoucher FormAction, txt(0), i, txt(2), frmTypeStr, TaxAcCodeSGST, txt(4), Val(txt(27)), txt(11), txt(12), txt(13), txt(7), txt(8), txt(9), txt(10), txt(19), txt(20), Taxable:=Val(txt(23)), Crt:=Val(txt(24)), Camt:=Val(txt(25)), Srt:=Val(txt(26)), Samt:=Val(txt(27)), Irt:=Val(txt(28)), Iamt:=Val(txt(29)), AgtVno:=CLng(txt(30)), AgtVtype:=txt(31), AgtVyear:=CLng(txt(32)), GSTEnt:=0, DrCrNtNo:=txt(34), DrCrNtType:=txt(35)
               i = i + 1
               DoEvents
            End If
            If Val(txt(29)) > 0 Then
               GProcCreateVoucher FormAction, txt(0), i, txt(2), frmTypeStr, TaxAcCodeIGST, txt(4), Val(txt(29)), txt(11), txt(12), txt(13), txt(7), txt(8), txt(9), txt(10), txt(19), txt(20), Taxable:=Val(txt(23)), Crt:=Val(txt(24)), Camt:=Val(txt(25)), Srt:=Val(txt(26)), Samt:=Val(txt(27)), Irt:=Val(txt(28)), Iamt:=Val(txt(29)), AgtVno:=CLng(txt(30)), AgtVtype:=txt(31), AgtVyear:=CLng(txt(32)), GSTEnt:=0, DrCrNtNo:=txt(34), DrCrNtType:=txt(35)
               i = i + 1
               DoEvents
            End If
        End If
        rstbl.Requery
        
        '--- Related Rcord Rec Vs Sale
        'If chkIsLastYearBill = 0 And (frmTypeStr = "SN" And gfrmHunkStr = "S") Or (frmTypeStr = "PN" And gfrmHunkStr = "B") Then
        If chkIsLastYearBill = 0 And ((frmTypeStr = "SN" And gfrmHunkStr = "S") Or (frmTypeStr = "PN" And gfrmHunkStr = "B")) Then
            SaveRelatedRecord
        End If
        If (frmTypeStr = "SN" And (gfrmHunkStr = "B" Or gfrmHunkStr = "T")) And txt(19) <> 1 Then
            gCn.Execute "delete from tblOutStanding where vtype='SN' and VNo = " & txt(0) & " and vyear = " & gCYear & ""
           X = "Select OutBrokerCode,OutMillCode from tblOutStanding where vno = " & txt(30) & " and vtype = '" & txt(31) & "' and vyear = " & txt(32) & ""
           If GProcRstOpen(Rs1, X, "R") > 0 Then
              wBrokerCd = Rs1.Fields(0)
              wMillCd = Rs1.Fields(1)
           Else
              wBrokerCd = txt(36)
           End If
           GProcCreateOutStanding txt(0), CDate(txt(2)), "SN", txt(34), txt(4), wBrokerCd, Val(txt(6)), 0, CDate(txt(2)), wMillCd, 0, 0
        End If
        If gfrmHunkStr = "B" Then  '--tbleInvoiceIRN
           gCn.Execute "delete from tbleInvoiceIRN where Vtype=  '" & frmTypeStr & "' and VNo = " & Val(txt(0)) & " and vyear = " & gCYear & ""
           If txtIRN(3) <> "" And IsDate(mskDtIRN(5)) Then
              gCn.Execute "insert into tbleInvoiceIRN (Vno,Vtype,Vyear,IrnNo,AckNo,AckDt,eQrCode) " _
                         & " Values(" & Val(txt(0)) & ", '" & frmTypeStr & "'," & gCYear & ",'" & txtIRN(3) & "','" & txtIRN(4) & "','" & Format(mskDtIRN(5), "yyyy/mm/dd") & "','" & txtIRN(6) & "')"
           End If
        End If
        
        '--Update Outstanding Balance
        GProcUpdateOutstandingBal CLng(txt(0)), UCase(txt(3)), gCYear
        rstbl.Find "vno = " & txt(0)
        ReadFields
        Call GProcActivateControls(Me, False)
        lbl1(3).Enabled = False
        lbl1(4).Enabled = False
        Call GProcSetButtons(Me, True)
        If FormAction = vbDataActionAddNew Then
            If cmdBtn(6).Visible = True Then cmdBtn(6).SetFocus
        Else
            If cmdBtn(7).Visible = True Then cmdBtn(7).SetFocus
        End If
        FormAction = vbDataActionClose
'        msGrid.Col = 0
'        msGrid.Row = 1
'        txtGrid.Move msGrid.CellLeft + msGrid.Left, msGrid.CellTop + msGrid.Top, msGrid.CellWidth, msGrid.CellHeight
    End If
Case 14 'Cancel
'    txtGrid.Visible = False
    Call GProcClearForm(Me, rstbl, MaxNo, True)
    If rstbl.RecordCount <= 0 Then
        Call GProcActivateControls(Me, False)
        Call GProcSetButtons(Me, vbDataActionCancel)
    Else
        If FormAction = vbDataActionAddNew Then
            rstbl.MoveLast
        Else
            '--Update Outstanding Balance
            GProcUpdateOutstandingBal CLng(txt(0)), UCase(txt(3)), gCYear
            rstbl.CancelUpdate
        End If
        DoEvents
        DoEvents
        DoEvents
        ReadFields
        Call GProcActivateControls(Me, False)
        Call GProcSetButtons(Me, True)
    End If
'    If (lbl1(3) = "Credit" And frmTypeStr = "SN") Or (lbl1(3) = "Debit" And frmTypeStr <> "SN") Then lbl1_Click (3)
    lbl1(3).Enabled = False
    lbl1(4).Enabled = False
    FormAction = vbDataActionClose
    If FormAction = vbDataActionAddNew Then
        If cmdBtn(6).Visible = True Then cmdBtn(6).SetFocus
    Else
        If cmdBtn(7).Visible = True And cmdBtn(7).Enabled = True Then cmdBtn(7).SetFocus
    End If
    FormAction = vbDataActionClose
'    msGrid.Col = 0
'    If msGrid.Row >= 1 Then msGrid.Row = 1
'    txtGrid.Move msGrid.CellLeft + msGrid.Left, msGrid.CellTop + msGrid.Top, msGrid.CellWidth, msGrid.CellHeight
End Select
Exit Sub
ErrorRoutine:
    If rstbl.EOF And rstbl.BOF Then
    Else
        rstbl.CancelUpdate
    End If
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
    Dim Y As String
    Dim rstTmp As Recordset
    Me.Top = 300
    Me.Left = Screen.Width / 2 - Me.Width / 2
    frmTypeStr = gfrmTypeStr
    CriteriaStr = " VType = '" & frmTypeStr & "'  and VYear=" & gCYear
    Y = "Select * from " & tblName & " where " & CriteriaStr & " and vctrno = 1 and (VDrCrNtType =  '" & gfrmHunkStr & "' or VDrCrNtType = '' )  "
    If Trim(OrderField) <> "" Then Y = Y & " order by " & OrderField
    GProcRstOpen rstbl, Y, "O"
    SetRelatedRS  '--tbleInvoiceIRN
    Select Case frmTypeStr
    Case "JV"
        Label6.Caption = "    Journal Voucher"
        lbl1(3).Caption = "Credit"
        lbl1(4).Caption = "Debit"
    Case "SN"
        If gfrmHunkStr = "B" Then
           Label6.Caption = "    Debit Note (Sale)"
        ElseIf gfrmHunkStr = "T" Then
           Label6.Caption = "    Debit Note (TCS)"
        Else
           Label6.Caption = "    Debit Note (Purchase)"
        End If
        lbl1(3).Caption = "Debit"
        lbl1(4).Caption = "Credit"
        lblDrCrNote.Caption = "Debit Note No  :"
    Case "PN"
        If gfrmHunkStr = "B" Then
           Label6.Caption = "    Credit Note (Sale)"
        Else
           Label6.Caption = "    Credit Note (Purchase)"
        End If
        lbl1(3).Caption = "Credit"
        lbl1(4).Caption = "Debit"
        lblDrCrNote.Caption = "Credit Note No  :"
    End Select
    If GProcRstOpen(rstTmp, "select CotCGSTRt,CotSGSTRt,CotIGSTRt,PolCGSTRt,PolSGSTRt,PolIGSTRt from tblMastNarration where Narration = 'G S T'", "R", gCn) > 0 Then
       wCotCGStRt = rstTmp.Fields(0)
       wCotSGStRt = rstTmp.Fields(1)
       wCotIGStRt = rstTmp.Fields(2)
       wPolCGStRt = rstTmp.Fields(3)
       wPolSGStRt = rstTmp.Fields(4)
       wPolIGStRt = rstTmp.Fields(5)
    End If
    
    FirstTimeFlag = True
    '---- Button Visiblity Set As Per User Previlage
    GProcSetButtonVisiblity Me
    cmdBtn_Click (6)
    '---Grid
    'SetGrid
End Sub
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
'If gFormAction = vbDataActionUpdate Then
'    MsgBox "You are about to modify the record. Save the record and then exit.", vbInformation + vbOKOnly, Me.Caption
'    Cancel = True
'    Exit Sub
'End If
rstbl.Close
End Sub
Private Sub lbl1_Click(Index As Integer)
'If Index = 3 Or Index = 4 Then
'    Dim x As String
'    x = lbl1(3)
'    lbl1(3) = lbl1(4)
'    lbl1(4) = x
'End If
End Sub

Private Sub mskDt_LostFocus(Index As Integer)
SaveBtnEd
End Sub
Private Sub mskDt_GotFocus(Index As Integer)
GProcSelectBox Me.ActiveControl
End Sub

Private Sub txt_GotFocus(Index As Integer)
GProcSelectBox Me.ActiveControl
If gfrmHunkStr <> "T" Then
If Index = 7 And txt(7) = "" Then
    txt(7) = "Ref. " + lblDrCrNote.Caption + " " + txt(34)
    txt(8) = Left(Left(txttemp(5).text, 15) + "-" + CStr(txt(23)) + " +GST-" + CStr(Format(Val(txt(25)) + Val(txt(27)) + Val(txt(29)), FStr)), 40)
    If IsDate(mskDt(12)) Then
       txt(9) = "Agt - " + txt(11) + " - " + (mskDt(12))
    Else
       txt(9) = "Agt - " + txt(11)
    End If
End If
End If
End Sub
Private Sub txt_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
Dim OldFrmType As String
Dim LeftPos As Long
Dim TopPos As Long
If KeyCode = vbKeyF3 Then '--F3 Create
    Select Case Index
        Case 7, 8, 9 '--Narration
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
        Case 7, 8, 9 '--Narration
            gClsSearch.SearchMultiField "tblMastNarration", "Narration", Array("Narration"), Array(txt(Index).Width), "NarrType='N'", "", "Narration", LeftPos, TopPos, False
            If gClsSearch.SearchMultiRetCol(0) <> "" Then txt(Index) = gClsSearch.SearchMultiRetCol(0)
    End Select
End If
End Sub
Private Sub txt_LostFocus(Index As Integer)
GProcCheckForNumber Me, rstbl, MaxNo, Array(34, 11)
SaveBtnEd
CalAmount
End Sub
Private Sub txt_KeyPress(Index As Integer, KeyAscii As Integer)
KeyAscii = GProcValidateKey(rstbl, Index, KeyAscii, txt(Index).text)
If Index = 11 And chkIsLastYearBill.Value = 0 Then
   Dim X As String
   Dim Pcode As Long
   Dim wItemType As Long
   Dim wItcd As Long
   Dim wMsOmsParty As String
   Dim rstTmp As Recordset
   Pcode = GProcGetColumnValue("tblMastAccount", "AcName", txttemp(4), "S", "Accode", "N") ', gCn)
   If frmTypeStr = "PN" And gfrmHunkStr = "B" Then
      X = "select tblOutstanding.OutBillNo,format(tblOutstanding.Vdt,'dd/mm/yyyy') as BillDt,tblOutstanding.OutAmount,tblOutstanding.Vno,tblOutstanding.Vtype,tblOutstanding.Vyear,tblOutstanding.OutItemCode from tblOutstanding where " _
          & " tblOutstanding.OutAcCode =  " & Pcode & " and tblOutstanding.vdt <= cdate('" & CDate(mskDt(2)) & "')"
   
   Else
      X = "select tblOutstanding.OutBillNo,format(iif(LEFT(tblOutstanding.VTYPE,1) = 'P',tblAddLess.AdGpdt,tblOutstanding.Vdt),'dd/mm/yyyy') as BillDt,tblOutstanding.OutAmount,tblOutstanding.Vno,tblOutstanding.Vtype,tblOutstanding.Vyear,tblOutstanding.OutItemCode from tblOutstanding, tblAddLess where " _
          & "tblOutstanding.Vno = tblAddLess.Vno and tblOutstanding.Vtype = tblAddLess.Vtype and tblOutstanding.Vyear = tblAddLess.Vyear and tblOutstanding.OutAcCode =  " & Pcode & " and tblOutstanding.vdt <= cdate('" & CDate(mskDt(2)) & "')"
   End If
   
''   X = "select tblOutstanding.OutBillNo,format(iif(LEFT(tblOutstanding.VTYPE,1) = 'P',tblOutstanding.OutPurBillDt,tblOutstanding.Vdt),'dd/mm/yyyy') as BillDt,tblOutstanding.OutAmount,tblOutstanding.Vno,tblOutstanding.Vtype,tblOutstanding.Vyear,tblOutstanding.OutItemCode from tblOutstanding where " _
''       & " tblOutstanding.OutAcCode =  " & Pcode & " and tblOutstanding.vdt <= cdate('" & CDate(mskDt(2)) & "')"
''
'''   X = "select tblOutstanding.OutBillNo,format(iif(LEFT(tblOutstanding.VTYPE,1) = 'P',tblAddLess.AdGpdt,tblOutstanding.Vdt),'dd/mm/yyyy') as BillDt,tblOutstanding.OutAmount,tblOutstanding.Vno,tblOutstanding.Vtype,tblOutstanding.Vyear,tblOutstanding.OutItemCode from tblOutstanding, tblAddLess where " _
'''       & "tblOutstanding.Vno = tblAddLess.Vno and tblOutstanding.Vtype = tblAddLess.Vtype and tblOutstanding.Vyear = tblAddLess.Vyear and tblOutstanding.OutAcCode =  " & Pcode & " and tblOutstanding.vdt <= cdate('" & CDate(mskDt(2)) & "')"
   gClsSearch.SearchMultiField X, "", Array("Bill No", "Bill Date", "Bill Amount", "", "", "", ""), Array(2000, 2000, 3000, 0, 0, 0, 0), "", "", "tblOutstanding.vdt,tblOutstanding.vno", 0, 0, False, , True
   KeyAscii = 0
   If gClsSearch.SearchMultiRetCol(0) <> "" Then
      txt(11) = gClsSearch.SearchMultiRetCol(0)
      mskDt(12) = gClsSearch.SearchMultiRetCol(1)
      txt(30) = gClsSearch.SearchMultiRetCol(3)
      txt(31) = gClsSearch.SearchMultiRetCol(4)
      txt(32) = gClsSearch.SearchMultiRetCol(5)
      If IsNull(gClsSearch.SearchMultiRetCol(6)) Or gClsSearch.SearchMultiRetCol(6) = "" Then
         wItcd = 0
      Else
         wItcd = gClsSearch.SearchMultiRetCol(6) & vbNullString
      End If
'--For HSN
      If gfrmHunkStr <> "S" Then
         txt(39) = GProcGetColumnValue("tblMastItem", "ItCode", "" & gClsSearch.SearchMultiRetCol(6) & "", "N", "ITTMP2", "S")
      End If
'--
      
      If Format(gClsSearch.SearchMultiRetCol(1), "dd/mm/yyyy") >= CDate("01/07/2017") Then
         If gfrmHunkStr = "S" Then
            X = "Select VSubItCode from tblVouSub where vno = " & gClsSearch.SearchMultiRetCol(3) & " and VType = '" & gClsSearch.SearchMultiRetCol(4) & "' and Vyear = " & gClsSearch.SearchMultiRetCol(5) & " and VSubCtrNo = 1"
            i = GProcRstOpen(rstTmp, X, "R")
            If i > 0 Then
               wItcd = rstTmp.Fields(0)
               wItemType = GProcGetColumnValue("tblMastItem", "ItCode", "" & wItcd & "", "N", "ItType", "N")
            End If
         Else
             wItemType = GProcGetColumnValue("tblMastItem", "ItCode", "" & gClsSearch.SearchMultiRetCol(6) & "", "N", "ItType", "N")
         End If
         wMsOmsParty = GProcGetColumnValue("tblMastAccount", "AcName", txttemp(4), "S", "AcGSTIN", "S")
         If Left(wMsOmsParty, 2) = "27" Or Left(wMsOmsParty, 1) = "U" Or wMsOmsParty = "" Or IsNull(wMsOmsParty) Then
            If wItemType = 0 Then  '--Cotton Rate
               txt(24) = wCotCGStRt
               txt(26) = wCotSGStRt
               txt(28) = 0
            Else
                If Format(gClsSearch.SearchMultiRetCol(1), "dd/mm/yyyy") < CDate("22/09/2025") Then
                   txt(24) = 6
                   txt(26) = 6
                 Else
                  txt(24) = wPolCGStRt
                  txt(26) = wPolSGStRt
                 End If
               txt(28) = 0
            End If
         Else
            If wItemType = 0 Then  '--Cotton Rate
              txt(24) = 0
              txt(26) = 0
              txt(28) = wCotIGStRt
            Else
              txt(24) = 0
              txt(26) = 0
              If Format(gClsSearch.SearchMultiRetCol(1), "dd/mm/yyyy") < CDate("22/09/2025") Then
                 txt(28) = 12
              Else
                 txt(28) = wPolIGStRt
              End If
            End If
         End If
      Else
         txt(24) = 0
         txt(26) = 0
         txt(28) = 0
      End If
      txt(23).SetFocus
   End If
End If
End Sub
Private Sub FillCombo()
'    cbo(1).Clear
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
txt(12) = IIf(IsDate(mskDt(12)), mskDt(12), "")
txt(13) = IIf(IsDate(mskDt(13)), mskDt(13), "")
txt(4) = GProcGetColumnValue("tblMastAccount", "AcName", txttemp(4), "S", "Accode", "N") ', gCn)
txt(5) = GProcGetColumnValue("tblMastAccount", "AcName", txttemp(5), "S", "AcCode", "N") ', gCn)
txt(20) = GProcGetColumnValue("tblMastNarration", "Narration", txttemp(20), "S", "NarrCode", "N")
txt(36) = GProcGetColumnValue("tblMastAccount", "AcName", txttemp(36), "S", "AcCode", "N") ', gCn)
If IsDate(mskDt(12)) = True Then
'If mskDt(12) < gCYSDate Then txt(29) = "SY"
End If
End Sub
Private Sub SetControlEd()
Dim AgCode1 As Long
Dim AgCode2 As Long
If FormAction = vbDataActionAddNew Or FormAction = vbDataActionUpdate Then
      AgCode1 = GProcGetColumnValue("tblMastAccount", "AcName", txttemp(5), "S", "AgCode", "N")
      AgCode2 = GProcGetColumnValue("tblMastAccount", "AcName", txttemp(4), "S", "AgCode", "N")
      If (AgCode1 = 90029 Or AgCode2 = 90029) Then  '--Mill
            txttemp(20).Enabled = True
      Else
            txttemp(20) = ""
            txttemp(20).Enabled = False
      End If
Else
    txttemp(20).Enabled = False
End If
If gfrmHunkStr = "S" Then
    txt(25).Enabled = True
    txt(27).Enabled = True
    txt(29).Enabled = True
'Else
'    txt(25).Enabled = False
'    txt(27).Enabled = False
'    txt(29).Enabled = False
End If
txt(6).Enabled = False
mskDt(12).Enabled = False
Label3.Visible = False
txt(14).Visible = False
If gfrmHunkStr = "T" Then
   Label3.Visible = True
   txt(14).Visible = True
   
   chkIsLastYearBill.Visible = False
   lbl1(2).Visible = False
   txt(11).Visible = False
   lbl1(7).Visible = False
   mskDt(12).Visible = False
   txt(24).Visible = False
   txt(25).Visible = False
   txt(26).Visible = False
   txt(27).Visible = False
   txt(28).Visible = False
   txt(29).Visible = False
   Label1.Visible = False
   Label2.Visible = False
   Label5.Visible = False
   Label4.Visible = False
   Label8.Visible = False
   Label7.Visible = False
   Label17.Caption = "TCS On Amount     :"
   Label9.Caption = "TCS Amount        :"
   txt(6).Enabled = True
End If
If frmTypeStr = "JV" Then
   msGrid.Enabled = False
   msGrid.Visible = False
   Shape5.Visible = False
   Label17.Visible = False
   Label2.Visible = False
   Label1.Visible = False
   Label5.Visible = False
   Label4.Visible = False
   Label7.Visible = False
   Label8.Visible = False
   txt(11).Visible = False
   mskDt(12).Visible = False
   Line1.Visible = False
   Label9.Visible = False
   lbl1(2).Visible = False
   lbl1(7).Visible = False
   For i = 36 To 42
       txt(i).Visible = False
   Next i
End If
If chkIsLastYearBill.Value = 1 Then
   mskDt(12).Enabled = True
Else
   mskDt(12).Enabled = False
End If
End Sub
Private Sub FillTempFromTxt()
mskDt(2) = IIf(IsDate(txt(2)), CDate(txt(2)), "__/__/____") '--Vdt
If IsDate(txt(12)) Then
   mskDt(12) = CDate(txt(12))
End If


txt(6) = Format(Abs(Val(txt(6))), "##0.00")
txttemp(4) = GProcGetColumnValue("tblMastAccount", "AcCode", txt(4), "N", "AcName", "S") ', gCn)
lblGstin.Caption = GProcGetColumnValue("tblMastAccount", "AcCode", txt(4), "N", "AcGSTIN", "S")

txttemp(5) = GProcGetColumnValue("tblMastAccount", "AcCode", txt(5), "N", "AcName", "S") ', gCn)
txttemp(20) = GProcGetColumnValue("tblMastNarration", "NarrCode", txt(20), "N", "Narration", "S")
txttemp(36) = GProcGetColumnValue("tblMastAccount", "AcCode", txt(36), "N", "AcName", "S") ', gCn)
'-- Fill Related Record Rec Vs Sale
'FillRelatedRecord
FillRelatedRecord   '---tbleInvoiceIRN
SetControlEd
End Sub
Private Function ValidateGrid() As Boolean
'Dim TotAmt As Double
'Dim i As Long
'Dim RCount As Long
'Dim j As Long
'Dim InvNo As Long
'Dim InVtype As String
'Dim InvYear As Long
'Dim IntAcCode As Long
'Dim BillDet As String
'If Trim(txtGrid.text) <> "" Then
'        msGrid.text = txtGrid.text
'End If
'txtGrid.Visible = False
'txtGrid.text = ""
'BillDet = ""
'CheckForNumberGrid
'With msGrid
'    '-- Duplicate invoice selection
'    For i = 1 To .rows - 1
'        If Trim(.TextMatrix(i, 1)) <> "" Then
'           BillDet = BillDet + .TextMatrix(i, 1) + "-" + CStr(Val(.TextMatrix(i, 4)) + Val(.TextMatrix(i, 6)) + Val(.TextMatrix(i, 9))) + ":"
'        End If
'        '-- Int A/c
'           If CDbl(.TextMatrix(i, 9)) > 0 Then
'                If gLatePayIntAcCode = GProcGetColumnValue("tblmastAccount", "AcName", "-", "S", "AcCode", "N") Then
'                    IntAcCode = 0
'                Else
'                    IntAcCode = gLatePayIntAcCode
'                End If
'                If IntAcCode = 0 Then
'                    MsgBox "Check Late Pay Interest Account.", vbInformation + vbOKOnly, Me.Caption
'                    GProcShowForm frmSetting, frmMain.mnuMstfrm(12)
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
'           If CDbl(.TextMatrix(i, 4)) > CDbl(.TextMatrix(i, 3)) Then
'                MsgBox "Adjusted Amount is more than Balance Bill Amount. ", vbOKOnly, Me.Caption
'                msGrid.SetFocus
'                Exit Function
'           End If
'            '-- Adj DB amt > Db Amt
'           If CDbl(.TextMatrix(i, 6)) > CDbl(.TextMatrix(i, 5)) Then
'                MsgBox "Adjusted DbNt Amount is more than Balance DbNt Amount. ", vbOKOnly, Me.Caption
'                msGrid.SetFocus
'                Exit Function
'           End If
'           TotAmt = TotAmt + Val(.TextMatrix(i, 4)) + Val(.TextMatrix(i, 6)) + Val(.TextMatrix(i, 9))
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
Dim MillAgCode1 As Long
Dim MillAgCode2 As Long
Dim oldFrmTypeStr As String
Dim IsMonthLastDt As Boolean
Dim rstAddCheck As Recordset
GProcCheckForNumber Me, rstbl, MaxNo, Array(34, 11)
SetControlEd
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
If gfrmHunkStr <> "T" Then
If txt(11) = "" Then
    MsgBox "Check Against Bill No.", vbInformation + vbOKOnly, Me.Caption
    txt(11).SetFocus
    Exit Function
End If
End If
'----  A/c
If txttemp(4) = "" Then
    MsgBox "Check Account.", vbInformation + vbOKOnly, Me.Caption
    txttemp(4).SetFocus
    Exit Function
End If
'-- Dr A/c
If GProcGetColumnValue("tblMastAccount", "AcName", txttemp(4), "S", "AcCode", "N") = 0 Then
  ' GName = GProcGetColumnValue("TblMastGroup", "AgCode", 90019, "N", "AgName", "S")
   GProcShowForm frmMastAccount, frmMain.mnuMstfrm(0), Array(frmMastAccount.txt(1), frmMastAccount.txttemp(2)), Array(txttemp(4), GName)
   Exit Function
End If
'---  A/c
If txttemp(5) = "" Then
    MsgBox "Check Account.", vbInformation + vbOKOnly, Me.Caption
    txttemp(5).SetFocus
    Exit Function
End If
'--Create  A/c
If GProcGetColumnValue("tblMastAccount", "AcName", txttemp(5), "S", "AcCode", "N") = 0 Then
   'GName = GProcGetColumnValue("TblMastGroup", "AgCode", 90015, "N", "AgName", "S")
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
If gfrmHunkStr <> "T" Then

'---Cgst Amt
If Abs(Round(Val(txt(23)) * Val(txt(24)) / 100, 2) - Val(txt(25))) > 1 Then
    MsgBox "Check CGST Amount.", vbInformation + vbOKOnly, Me.Caption
    txt(25).SetFocus
    Exit Function
End If

'---Sgst Amt
If Abs(Round(Val(txt(23)) * Val(txt(26)) / 100, 2) - Val(txt(27))) > 1 Then
    MsgBox "Check SGST Amount.", vbInformation + vbOKOnly, Me.Caption
    txt(27).SetFocus
    Exit Function
End If
'---Igst Amt
If Abs(Round(Val(txt(23)) * Val(txt(28)) / 100, 2) - Val(txt(29))) > 1 Then
    MsgBox "Check IGST Amount.", vbInformation + vbOKOnly, Me.Caption
    txt(29).SetFocus
    Exit Function
End If
End If
'---Mill Expenses (not on month last date)
'MillAgCode1 = GProcGetColumnValue("tblMastAccount", "AcName", txttemp(5), "S", "AgCode", "N", " ACMILLTYPE<>0")
'MillAgCode2 = GProcGetColumnValue("tblMastAccount", "AcName", txttemp(4), "S", "AgCode", "N", " ACMILLTYPE<>0")
'If Format(DateAdd("d", 1, mskDt(2)), "dd") = 1 Then
'    IsMonthLastDt = True
'Else
'    IsMonthLastDt = False
'End If
'If IsMonthLastDt = False Then
'    If txttemp(20) = "" And (MillAgCode1 = 90029 Or MillAgCode2 = 90029) Then
'        MsgBox "Check Mill Expenses Description.", vbInformation + vbOKOnly, Me.Caption
'        txttemp(20).SetFocus
'        Exit Function
'    End If
'    If Trim(txttemp(20)) <> "" And GProcGetColumnValue("tblMastNarration", "Narration", txttemp(20), "S", "NarrCode", "N") = 0 Then
'        oldFrmTypeStr = gfrmTypeStr
'        gfrmTypeStr = "M"
'        GProcShowForm frmMastNarration, frmMain.mnuMstfrm(7), Array(frmMastNarration.txt(2)), Array(txttemp(20))
'        gfrmTypeStr = oldFrmTypeStr
'        Exit Function
'    End If
'End If
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
If Index = 36 And frmTypeStr = "SN" And txt(36) = 0 And txttemp(4) <> "" Then
   Dim X As String
   Dim BrCd As Long
   Dim Ptcd As Long
   Dim TmpRec As Recordset
   Ptcd = GProcGetColumnValue("tblMastAccount", "AcName", txttemp(4), "S", "Accode", "N") ', gCn)
   X = "Select OutBrokerCode from tblOutStanding where vtype = 'SY' and OutAcCode = " & Ptcd & " order by Vdt desc"
   If GProcRstOpen(TmpRec, X, "R", gCn) > 0 Then
      txttemp(36) = GProcGetColumnValue("tblMastAccount", "AcCode", TmpRec.Fields(0), "N", "AcName", "S")
   End If
End If
End Sub

Private Sub txttemp_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
Dim GName As String
Dim oldFrmTypeStr As String
If KeyCode = vbKeyF3 Then
    Select Case Index
        Case 4 '-- A/c
            GProcShowForm frmMastAccount, frmMain.mnuMstfrm(0), Array(frmMastAccount.txttemp(2)), Array(GName)
        Case 5 '-- A/c
            'GName = GProcGetColumnValue("TblMastGroup", "AgCode", 90015, "N", "AgName", "S")
            GProcShowForm frmMastAccount, frmMain.mnuMstfrm(0), Array(frmMastAccount.txttemp(2)), Array(GName)
        Case 20 '-- Mill Exp Description
             oldFrmTypeStr = gfrmTypeStr
            gfrmTypeStr = "M"
            GProcShowForm frmMastNarration, frmMain.mnuMstfrm(7)
            gfrmTypeStr = oldFrmTypeStr
    End Select
End If
End Sub
Private Sub txtTemp_KeyPress(Index As Integer, KeyAscii As Integer)
If KeyAscii = 13 Then Exit Sub
Dim LeftPos As Long
Dim TopPos As Long
TopPos = txttemp(Index).Top + Me.Top + 650
LeftPos = txttemp(Index).Left + Me.Left
Select Case Index
    Case 4, 5
        gClsSearch.SearchMultiField "QryMastaccount", "Account,AcName,AcCity,AgName", Array("Account", "AcName", "City", "Group"), Array(3000, 0, 2000, 3000), "", Chr(KeyAscii), "AcName", LeftPos, TopPos, True
        KeyAscii = 0
        If gClsSearch.SearchMultiRetCol(0) <> "" And gClsSearch.SearchMultiRetCol(1) = "" Then
            txttemp(Index) = gClsSearch.SearchMultiRetCol(0)
        Else
            If gClsSearch.SearchMultiRetCol(1) <> "" Then txttemp(Index) = gClsSearch.SearchMultiRetCol(1)
        End If
        If Index = 4 Then
           lblGstin.Caption = GProcGetColumnValue("tblMastAccount", "AcName", txttemp(Index), "S", "AcGSTIN", "S")
        End If
    Case 20 '--Mill Expes. Narration
        gClsSearch.SearchMultiField "TblMastNarration", "Narration", Array("Mill Expenses Description"), Array(4000), "NarrType='M'", Chr(KeyAscii), "Narration", LeftPos, TopPos, True
        KeyAscii = 0
        If gClsSearch.SearchMultiRetCol(0) <> "" Then
            txttemp(Index) = gClsSearch.SearchMultiRetCol(0)
        End If
    Case 36 '--broker
        gClsSearch.SearchMultiField "QryMastaccount", "Account,AcName", Array("Broker", "AcName"), Array(txttemp(Index).Width, 0), "(AgCode=90016 or GpCode=90016)", Chr(KeyAscii), "AcName", LeftPos, TopPos, True
        KeyAscii = 0
        If gClsSearch.SearchMultiRetCol(0) <> "" And gClsSearch.SearchMultiRetCol(1) = "" Then
            txttemp(Index) = gClsSearch.SearchMultiRetCol(0)
        Else
            If gClsSearch.SearchMultiRetCol(1) <> "" Then txttemp(Index) = gClsSearch.SearchMultiRetCol(1)
        End If
End Select
SetControlEd
End Sub
'--- Related Record
Private Function IsRelatedRecord() As Boolean
'-- Audited
If GProcISAudited(txt(0), txt(3), txt(17)) = True Then
    MsgBox "Record is Audited.You can not Modify/Delete this Record.", vbCritical + vbOKOnly, App.Title
    IsRelatedRecord = True
    Exit Function
End If
'-- Related Record JV
If rstbl.EOF And rstbl.BOF Then
Else
    If rstbl.Fields("VIsRefEntType") = 1 Then
        MsgBox "Related Record is Present.You can not Modify/Delete this Record.", vbCritical + vbOKOnly, App.Title
        IsRelatedRecord = True
        Exit Function
    End If
End If
End Function
Private Sub txtTemp_LostFocus(Index As Integer)
SaveBtnEd
SetControlEd
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
Dim Y As String
Dim RsSA1 As Recordset
Select Case Index
Case 0, 1:
    '------ Validate
    If IsNumeric(txtno(0)) = False Then txtno(0) = 0
    If IsNumeric(txtno(1)) = False Then txtno(1) = txtno(0)
    If CLng(txtno(1)) < CLng(txtno(0)) Then txtno(1) = txtno(0)
    '-------
    If Index = 0 Then GProcCrystalRptPreparation CryReport, "Printer"
    If Index = 1 Then GProcCrystalRptPreparation CryReport, "Window"
    CryReport.WindowTitle = "Journal Entry"
    Vno = CLng(txtno(0))
    Do While Vno <= txtno(1)
        X = "Select * From tblVoucher where  vNo=" & Vno & " and VCTRNO =1 and VType='" & frmTypeStr & "' and VYear=" & gCYear
        If GProcRstOpen(Rs1, X, "R") > 0 Then
            With CryReport
                .SelectionFormula = "{tblVoucher.vNo} = " & Rs1!Vno & " and {tblVoucher.VCTRNO}=1 and {tblVoucher.VType}='" & Rs1!Vtype & "' and {tblVoucher.VYear}=" & Rs1!VYear
                .Formulas(5) = "wCompGstin = '" & gCGSTIN & "'"
                .Formulas(6) = "wAmountInWords = '" & GProcAmountInWords(Abs(Rs1!vamt)) & "'"
                Y = "Select IrnNo,AckNo,AckDt from tbleInvoiceIRN where vno = " & Rs1!Vno & " and Vtype  = '" & Rs1!Vtype & "' and vyear = " & Rs1!VYear & ""
                
                i = GProcRstOpen(RsSA1, Y, "R")
                If i > 0 Then
                   .Formulas(7) = "wIrnNo = '" & RsSA1.Fields(0) & "'"
                   .Formulas(8) = "wAckNo = '" & RsSA1.Fields(1) & "'"
                   .Formulas(9) = "wAckDt = '" & RsSA1.Fields(2) & "'"
                Else
                   .Formulas(7) = "wIrnNo = ''"
                   .Formulas(8) = "wAckNo = ''"
                   .Formulas(9) = "wAckDt = ''"
                End If
                If frmTypeStr = "SN" Then '-- Debit Note
                     If gfrmHunkStr = "T" Then
                        .ReportFileName = gReportPath & "rptTCSDbNtPrint.rpt"
                     Else
                        .ReportFileName = gReportPath & "rptVouDbNtPrint.rpt"
                     End If
                ElseIf frmTypeStr = "PN" Then '-- credit Note
                    .ReportFileName = gReportPath & "rptVouCrNtPrint.rpt"
                Else '-- Jv
                    .ReportFileName = gReportPath & "rptVouJVPrint.rpt"
                End If
                .Action = 1
            End With
        End If
        Vno = Vno + 1
    Loop
Case 4, 6:
     If gfrmHunkStr <> "B" Then
        MsgBox ("e-Invoice Is Only For B2B...")
        Exit Sub
     End If

    Dim RsSA As Recordset
    If IsNumeric(txtno(0)) = False Then txtno(0) = 0
    If IsNumeric(txtno(1)) = False Then txtno(1) = txtno(0)
    If CLng(txtno(1)) < CLng(txtno(0)) Then txtno(1) = txtno(0)

    X = "Select a.Vno,a.Vtype,a.Vdt,a.Vyear,a.VAcCode,a.VAcOCode,a.VDRCRNTNO,val(mid(a.VDRCRNTNO,4,15)),a.VTaxableAmt,abs(a.VAmt),a.VCGSTRt,a.VCGSTAmt,a.VSGSTRt,a.VSGSTAmt, " _
        & " a.VIGSTRt,a.VIGSTAmt,abs(a.VAmt) - (a.VTaxableAmt+a.VCGSTAmt+a.VSGSTAmt+a.VIGSTAmt),'" & gUserName & "',a.VAgtVno,a.VAgtVtype,a.VAgtVyear " _
        & " from qrytblVoucherSr1 A " _
        & " where a.Vno >= " & Val(txtno(0)) & " and a.Vno <= " & Val(txtno(1)) & " and a.VType = '" & frmTypeStr & "' and VDRCRNTTYPE = '" & gfrmHunkStr & "' and a.VCGSTAmt+a.VSGSTAmt+a.VIGSTAmt > 0  and vyear =  " & gCYear
    gCn.BeginTrans
    gCn.Execute "Delete from tmpGenTbl2 where  USERNAME='" & gUserName & "'"

    gCn.Execute " insert into tmpGenTbl2 (Vno,Vtype,Vdt,Vyear,VACCODE,VACOCODE,VBILLNO,VsrNo,Amt1,Amt2,Amt3,Amt4,Amt5,Amt6,Amt7,Amt8,Amt9," _
                & " USERNAME,Amt10,nar1,Amt11  ) " & X
    gCn.CommitTrans
    X = "Select * from tmpGenTbl2 where  USERNAME='" & gUserName & "' order by VsrNo"
    If GProcRstOpen(RsSA, X, "R") <= 0 Then
       MsgBox "No Records..."
       Exit Sub
    End If
     
    If Index = 6 Then
        GProcCrystalRptPreparation CryReport, "Window"
        With CryReport
             .Formulas(4) = "wCompGstin = '" & gCGSTIN & "'"
             .SelectionFormula = "{tmpGenTbl2.USERNAME} = '" & gUserName & "' "
             .ReportFileName = gReportPath & "rpte-Invoice-CrDrNote.rpt"
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
    Dim aFile As String
    Dim ItemName As String
    Dim ii As Integer
    LnCnt = 1

    aFile = "D:\e-Inv-" + gCCode + "*.json"
    If Len(Dir$(aFile)) > 0 Then
          Kill aFile
    End If

    Open "D:\e-Inv-" + gCCode + ".json" For Output As #1

    Print #1, "["
    RsSA.MoveFirst
    Do While Not RsSA.EOF
       FirmGstin = gCGSTIN
       PartyName = GProcGetColumnValue("tblMastAccount", "AcCode", RsSA.Fields("VACCODE"), "N", "AcName", "S")
       
       PartyGstin = GProcGetColumnValue("tblMastAccount", "AcCode", RsSA.Fields("VACCODE"), "N", "AcGSTIN", "S")
       PartyAdd1 = GProcGetColumnValue("tblMastAccount", "AcCode", RsSA.Fields("VACCODE"), "N", "AcAdd1", "S")
       PartyAdd2 = GProcGetColumnValue("tblMastAccount", "AcCode", RsSA.Fields("VACCODE"), "N", "AcAdd2", "S")
       PartyCity = GProcGetColumnValue("tblMastAccount", "AcCode", RsSA.Fields("VACCODE"), "N", "AcCity", "S")
       PartyPin = GProcGetColumnValue("tblMastAccount", "AcCode", RsSA.Fields("VACCODE"), "N", "AcZip", "S")
       ItemHsn = ""   '--GProcGetColumnValue("tblMastItem", "ItCode", RsSA.fields("VACOCODE"), "N", "ITTMP2", "S")
       If Len(PartyPin) > 6 Then
          PartyPin = Left(PartyPin, 3) + Right(PartyPin, 3)
       End If
       If Trim(PartyAdd1) = "" Then PartyAdd1 = "null"
       If Trim(PartyAdd2) = "" Then PartyAdd2 = "null"
       
       X = "Select OutItemCode from tblOutStanding where vno = " & RsSA.Fields("Amt10") & " and Vtype = '" & RsSA.Fields("Nar1") & "' and vyear = " & RsSA.Fields("Amt11") & ""
       
       i = GProcRstOpen(Rs1, X, "R")
       If i > 0 Then
          ItemName = GProcGetColumnValue("tblMastAccount", "AcCode", RsSA.Fields("VACOCODE"), "N", "AcName", "S")
          ItemHsn = GProcGetColumnValue("tblMastItem", "ItCode", Rs1.Fields(0), "N", "ITTMP2", "S")
       Else '--- For last year Bill
          ItemName = "Yarn"
          If (RsSA.Fields("Amt3") + RsSA.Fields("Amt5") + RsSA.Fields("Amt7")) <= 5 Then
             ItemHsn = "5205"
          Else
             ItemHsn = "5509"
          End If
       End If
       
'       If i > 0 And Rs1.Fields(0) <> "" Then
'          ItemHsn = GProcGetColumnValue("tblMastItem", "ItCode", Rs1.Fields(0), "N", "ITTMP2", "S")
'       End If
       If Trim(ItemHsn) = "" Then
          MsgBox ("HSN Code Not Found For... '" & PartyName & "'   ")
          Close #1
          Exit Sub
       End If
       ItemName = GProcGetColumnValue("tblMastAccount", "AcCode", RsSA.Fields("VACOCODE"), "N", "AcName", "S")
       
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

       Print #1, "{"
       Print #1, Chr(34) + "Version" + Chr(34) + ":" + Chr(34) + "1.1" + Chr(34) + Chr(44)
       Print #1, Chr(34) + "TranDtls" + Chr(34) + ":{"
       Print #1, Space(15) + Chr(34) + "TaxSch" + Chr(34) + ":" + Chr(34) + "GST" + Chr(34) + Chr(44)
       Print #1, Space(15) + Chr(34) + "SupTyp" + Chr(34) + ":" + Chr(34) + "B2B" + Chr(34) + Chr(44)
       Print #1, Space(15) + Chr(34) + "IgstOnIntra" + Chr(34) + ":" + Chr(34) + "N" + Chr(34) + Chr(44)
       Print #1, Space(15) + Chr(34) + "RegRev" + Chr(34) + ":" + "null" + Chr(44)
       Print #1, Space(15) + Chr(34) + "EcmGstin" + Chr(34) + ":" + "null"
       Print #1, "}" + Chr(44)

       Print #1, Chr(34) + "DocDtls" + Chr(34) + ":{"
       If frmTypeStr = "SN" Then
          Print #1, Space(15) + Chr(34) + "Typ" + Chr(34) + ":" + Chr(34) + "DBN" + Chr(34) + Chr(44)
       End If
       If frmTypeStr = "PN" Then
          Print #1, Space(15) + Chr(34) + "Typ" + Chr(34) + ":" + Chr(34) + "CRN" + Chr(34) + Chr(44)
       End If
       Print #1, Space(15) + Chr(34) + "No" + Chr(34) + ":" + Chr(34) + RsSA.Fields("VBILLNO") + Chr(34) + Chr(44)
       Print #1, Space(15) + Chr(34) + "Dt" + Chr(34) + ":" + Chr(34) + CStr(RsSA.Fields("Vdt")) + Chr(34)
       Print #1, "}" + Chr(44)

       Print #1, Chr(34) + "SellerDtls" + Chr(34) + ":{"
       Print #1, Space(15) + Chr(34) + "Gstin" + Chr(34) + ":" + Chr(34) + "" & FirmGstin & "" + Chr(34) + Chr(44)
       Print #1, Space(15) + Chr(34) + "LglNm" + Chr(34) + ":" + Chr(34) + "" & UCase(gCName) & "" + Chr(34) + Chr(44)
       Print #1, Space(15) + Chr(34) + "TrdNm" + Chr(34) + ":" + "null" + Chr(44)
       Print #1, Space(15) + Chr(34) + "Addr1" + Chr(34) + ":" + Chr(34) + "" & gCAdd1 & "" + Chr(34) + Chr(44)
       Print #1, Space(15) + Chr(34) + "Addr2" + Chr(34) + ":" + Chr(34) + "" & gCAdd3 & "" + Chr(34) + Chr(44)
       Print #1, Space(15) + Chr(34) + "Loc" + Chr(34) + ":" + Chr(34) + "Ichalkaranji" + Chr(34) + Chr(44)
       Print #1, Space(15) + Chr(34) + "Pin" + Chr(34) + ":" + "416115" + Chr(44)
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
       Print #1, Chr(34) + "ShipDtls" + Chr(34) + ":" + "null" + Chr(44)


       Print #1, Chr(34) + "ValDtls" + Chr(34) + ":{"
       Print #1, Space(15) + Chr(34) + "AssVal" + Chr(34) + ":" + Trim(CStr(Format(RsSA.Fields("Amt1"), "############0.00"))) + Chr(44)
       Print #1, Space(15) + Chr(34) + "IgstVal" + Chr(34) + ":" + Trim(CStr(Format(RsSA.Fields("Amt8"), "############0.00"))) + Chr(44)
       Print #1, Space(15) + Chr(34) + "CgstVal" + Chr(34) + ":" + Trim(CStr(Format(RsSA.Fields("Amt4"), "############0.00"))) + Chr(44)
       Print #1, Space(15) + Chr(34) + "SgstVal" + Chr(34) + ":" + Trim(CStr(Format(RsSA.Fields("Amt6"), "############0.00"))) + Chr(44)
       Print #1, Space(15) + Chr(34) + "CesVal" + Chr(34) + ":" + Trim(CStr(Format(0, "############0.00"))) + Chr(44)
       Print #1, Space(15) + Chr(34) + "StCesVal" + Chr(34) + ":" + Trim(CStr(Format(0, "############0.00"))) + Chr(44)
       Print #1, Space(15) + Chr(34) + "Discount" + Chr(34) + ":" + Trim(CStr(Format(0, "############0.00"))) + Chr(44)
       Print #1, Space(15) + Chr(34) + "OthChrg" + Chr(34) + ":" + Trim(CStr(Format(0, "############0.00"))) + Chr(44)
       Print #1, Space(15) + Chr(34) + "RndOffAmt" + Chr(34) + ":" + Trim(CStr(Format(RsSA.Fields("Amt9"), "############0.00"))) + Chr(44)
       Print #1, Space(15) + Chr(34) + "TotInvVal" + Chr(34) + ":" + Trim(CStr(Format(RsSA.Fields("Amt2"), "############0.00")))
       Print #1, "}" + Chr(44)


       Print #1, Space(25) + Chr(34) + "ItemList" + Chr(34) + ":[{"
       Print #1, Space(32) + Chr(34) + "SlNo" + Chr(34) + ":" + Chr(34) + "1" + Chr(34) + Chr(44)
       Print #1, Space(32) + Chr(34) + "PrdDesc" + Chr(34) + ":" + Chr(34) + "" & ItemName & "" + Chr(34) + Chr(44)
       Print #1, Space(32) + Chr(34) + "IsServc" + Chr(34) + ":" + Chr(34) + "N" + Chr(34) + Chr(44)
       Print #1, Space(32) + Chr(34) + "HsnCd" + Chr(34) + ":" + Chr(34) + Trim(ItemHsn) + Chr(34) + Chr(44)
       Print #1, Space(32) + Chr(34) + "Qty" + Chr(34) + ":" + Trim(CStr(Format(0, "############0.000"))) + Chr(44)
       Print #1, Space(32) + Chr(34) + "Unit" + Chr(34) + ":" + Chr(34) + "OTH" + Chr(34) + Chr(44)
       Print #1, Space(32) + Chr(34) + "UnitPrice" + Chr(34) + ":" + Trim(CStr(Format(0, "############0.000"))) + Chr(44)
       Print #1, Space(32) + Chr(34) + "TotAmt" + Chr(34) + ":" + Trim(CStr(Format(RsSA.Fields("Amt1"), "############0.00"))) + Chr(44)
       Print #1, Space(32) + Chr(34) + "Discount" + Chr(34) + ":" + Trim(CStr(Format(0, "############0.00"))) + Chr(44)
       Print #1, Space(32) + Chr(34) + "PreTaxVal" + Chr(34) + ":" + Trim(CStr(Format(0, "############0.00"))) + Chr(44)
       Print #1, Space(32) + Chr(34) + "AssAmt" + Chr(34) + ":" + Trim(CStr(Format(RsSA.Fields("Amt1"), "############0.00"))) + Chr(44)
       Print #1, Space(32) + Chr(34) + "GstRt" + Chr(34) + ":" + Trim(CStr(Format(RsSA.Fields("Amt3") + RsSA.Fields("Amt5") + RsSA.Fields("Amt7"), "############0"))) + Chr(44)
       Print #1, Space(32) + Chr(34) + "IgstAmt" + Chr(34) + ":" + Trim(CStr(Format(RsSA.Fields("Amt8"), "############0.00"))) + Chr(44)
       Print #1, Space(32) + Chr(34) + "CgstAmt" + Chr(34) + ":" + Trim(CStr(Format(RsSA.Fields("Amt4"), "############0.00"))) + Chr(44)
       Print #1, Space(32) + Chr(34) + "SgstAmt" + Chr(34) + ":" + Trim(CStr(Format(RsSA.Fields("Amt6"), "############0.00"))) + Chr(44)
       Print #1, Space(32) + Chr(34) + "CesRt" + Chr(34) + ":" + Trim(CStr(Format(0, "############0.00"))) + Chr(44)
       Print #1, Space(32) + Chr(34) + "CesAmt" + Chr(34) + ":" + Trim(CStr(Format(0, "############0.00"))) + Chr(44)
       Print #1, Space(32) + Chr(34) + "CesNonAdvlAmt" + Chr(34) + ":" + Trim(CStr(Format(0, "############0.00"))) + Chr(44)
       Print #1, Space(32) + Chr(34) + "StateCesRt" + Chr(34) + ":" + Trim(CStr(Format(0, "############0.00"))) + Chr(44)
       Print #1, Space(32) + Chr(34) + "StateCesAmt" + Chr(34) + ":" + Trim(CStr(Format(0, "############0.00"))) + Chr(44)
       Print #1, Space(32) + Chr(34) + "StateCesNonAdvlAmt" + Chr(34) + ":" + Trim(CStr(Format(0, "############0.00"))) + Chr(44)
       Print #1, Space(32) + Chr(34) + "OthChrg" + Chr(34) + ":" + Trim(CStr(Format(0, "############0.00"))) + Chr(44)
       Print #1, Space(32) + Chr(34) + "TotItemVal" + Chr(34) + ":" + Trim(CStr(Format(RsSA.Fields("Amt1") + RsSA.Fields("Amt8") + RsSA.Fields("Amt4") + RsSA.Fields("Amt6"), "############0.00")))
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
    
Case 5:
    Dim excel_app As Excel.Application
    Dim workbook As Excel.workbook
    Dim sheet_name As String
    Dim sheet As Excel.Worksheet
    Dim header_range As Excel.Range

    Dim flnm As String
    Dim a As Integer
    Dim PtyCd As Long
'    Dim i As Integer
    Dim wVno As Long
    Dim wVtype As String
    Dim wVyear As Long
    
    
    
    flnm = "d:\UploadedInvoiceDetails"
    Set excel_app = New Excel.Application
    excel_app.Visible = False
    Set workbook = excel_app.Workbooks.Open(filename:=flnm)
    Set sheet = workbook.Sheets.Item("UploadedInvoiceDetails")
    If (sheet Is Nothing) Then
        MsgBox "Error Sales Sheet"
    End If
    a = 2
    Do While sheet.Cells(a, 1) <> ""
       X = "select vno,vtype,vyear from tblVoucher where vtype= '" & frmTypeStr & "' and VDRCRNTNO= '" & sheet.Cells(a, 5) & "' and " _
          & " VCtrNo = 1 and VYear =" & gCYear
       i = GProcRstOpen(RsSA, X, "R")
       If i > 0 Then
          wVno = RsSA.Fields(0)
          wVtype = RsSA.Fields(1)
          wVyear = RsSA.Fields(2)
          gCn.Execute "delete from tbleInvoiceIRN where Vtype=  '" & frmTypeStr & "' and VNo = " & wVno & " and vyear = " & wVyear & ""
              
          gCn.Execute "insert into tbleInvoiceIRN (Vno,Vtype,Vyear,VBillNo,IrnNo,AckNo,AckDt,eQrCode) " _
                      & " Values(" & wVno & ", '" & wVtype & "'," & wVyear & ",'" & sheet.Cells(a, 5) & "'  ,'" & sheet.Cells(a, 2) & "','" & sheet.Cells(a, 3) & "','" & Format(sheet.Cells(a, 4), "yyyy/mm/dd") & "','" & sheet.Cells(a, 11) & "')"

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
'---------------- Grid
Private Sub SetGrid()
'Dim X As String, x1Rec As Double
'Set rsRecVsSale = New Recordset
'X = "select InvTp, InvBillNo, InvDt, BillAmt, AdjAmt, DbNtAmt, AdjDbNtAmt, LateDays, Interest, IntReceived ,InvYear,InvNo,JVNOINT" _
'& " from tblRecVsSale where " _
'& " Recvtype='" & frmTypeStr & "' and Recvno=" & txt(0) & " and RecVYear=" & gCYear & " order by RecVCtrlNo"
'x1Rec = GProcRstOpen(rsRecVsSale, X, "R")
'With msGrid
'    .Clear
'    .FormatString = "<Type|>Inv. No   |<Inv. Date      |>Invoice Amt |>Adj. Amount|>DbNt  Amt|>Adj DbNt Amt|>Late Days|>Late Pay Int|>Int Received|||"
'    .ColWidth(10) = 0
'    .ColWidth(11) = 0
'    .ColWidth(12) = 0
'    If FormAction = vbDataActionAddNew Then
'        .rows = 2
'        GridAddNew
'    Else
'        .rows = 2
'    End If
'End With
End Sub
Private Sub GridAddNew()
'Dim i As Long
'With msGrid
'    For i = 0 To rsRecVsSale.Fields.Count - 1
'        Select Case rsRecVsSale.Fields(i).Type
'            Case adInteger, 2, 3 '-- Integer
'                .TextMatrix(.Row, i) = 0
'            Case adDecimal, 131, adNumeric, adDouble '-- Double
'                .TextMatrix(.Row, i) = Format(0, GProcNumberFormat(rsRecVsSale.Fields(i).Precision, rsRecVsSale.Fields(i).NumericScale))
'            Case adDate '-- Date
'                 .TextMatrix(.Row, i) = ""
'            Case Else
'                 .TextMatrix(.Row, i) = ""
'        End Select
'    Next i
''ReNumberGridSrno
'.Refresh
'End With
End Sub
Private Sub msGrid_Click()
'On Error Resume Next
'    msGrid_EnterCell
End Sub
Private Sub msGrid_GotFocus()
   ' cmdBtn(10).Cancel = False
End Sub
Private Sub MSGrid_LostFocus()
'CalAmount
End Sub
Private Sub msGrid_EnterCell()
'If FormAction <> vbDataActionClose Then
'    CheckForNumberTxtGrid
'    DoEvents
'    txtGrid = ""
'    DoEvents
'    If msGrid.text <> "" Then txtGrid.text = msGrid.text
'    '--- Lock TextBox
'    With msGrid
'        Select Case .Col
'            Case 0, 1, 2, 3, 5, 7, 8, 10, 11, 12: '--Type,Inv No, Inv Dt, Debit Note
'                txtGrid.Locked = True
'            Case Else
'                txtGrid.Locked = False
'        End Select
'    End With
'    DoEvents
'    msGrid_KeyPress (0)
'    DoEvents
'End If
'End Sub
'Private Sub msGrid_KeyPress(KeyAscii As Integer)
'With msGrid
'    If KeyAscii = 27 Then '----Esc Key
'        KeyAscii = 0
'        If (Val(.TextMatrix(.Row, 0)) = 0) And .rows > 2 Then
'            .RemoveItem (.Row)
'            txtGrid.text = ""
'        End If
'        txtGrid.Visible = False
'        'txt(18).SetFocus
'        Exit Sub
'    End If
'
'    '---- Navigate Row ,Col
'    If KeyAscii = 13 Then
'        KeyAscii = 0
'        If .Col = rsRecVsSale.Fields.Count - 4 Then
'            If .Row = (.rows - 1) Then '-- Add New Row
'                .rows = .rows + 1
'                .Row = .Row + 1
''                GridAddNew
'            Else '--Next Row
'                .Row = .Row + 1
'            End If
'            .Col = 0
'        Else    '-- Next col
'            .Col = .Col + 1
'        End If
'    End If
'
'    If KeyAscii <> 13 And KeyAscii <> 27 And .Col = 0 Then
'        DoEvents
'        KeyAscii = 0
'        Pending_Invoice
'    End If
'End With
'
''---- Set TxtGrid
'If Trim(msGrid.text) <> "" Then
'    txtGrid = Trim(msGrid.text)
'End If
'If msGrid.Col < (rsRecVsSale.Fields.Count - 3) Then
'    txtGrid.Move msGrid.CellLeft + msGrid.Left, msGrid.CellTop + msGrid.Top, msGrid.CellWidth, msGrid.CellHeight
'End If
'txtGrid.Visible = True
'txtGrid.Enabled = True
'DoEvents
'If KeyAscii <> 0 Then
'    Select Case rsRecVsSale.Fields(msGrid.Col).Type
'        Case adChar, adVarWChar
'            If msGrid.Col = 0 Then  '-- Invoice
'                txtGrid_KeyPress (KeyAscii)
'            Else
'                txtGrid.text = txtGrid.text & Chr(KeyAscii)
'            End If
'        Case Else
'            txtGrid.text = Chr(GProcValidateKey(rsRecVsSale, msGrid.Col, KeyAscii, txtGrid.text)) 'Chr(GProcNumberWithDecimal(KeyAscii, txtGrid))
'    End Select
'End If
'DoEvents
'If msGrid.text <> "" Then
'    txtGrid.SelStart = Len(txtGrid.text) + 1
'Else
'    txtGrid.SelStart = 1
'End If
'DoEvents
'If txtGrid.Visible = True Then txtGrid.SetFocus
'End Sub
'Private Sub msGrid_LeaveCell()
'If txtGrid.Visible Then
'    msGrid.text = txtGrid.text
'    DoEvents
'    txtGrid.text = ""
'    txtGrid.Visible = False
'End If
''----Late Payment Interest Calcualtion
''---Interest Rate 1.5
'Select Case msGrid.Col
'Case 4:
'    With msGrid
'        .TextMatrix(.Row, 8) = Format(Round((((Val(.TextMatrix(.Row, 4)) * 1.5) / 3000) * Val(.TextMatrix(.Row, 7))), 2), "###0.00")
'    End With
'   ' CalBalance
'Case 6, 9:
'   ' CalBalance
'End Select
End Sub
Private Sub txtGrid_GotFocus()
'    txtGrid.Alignment = ProcSetAlignment(rsRecVsSale, msGrid.Col)
'    DoEvents
'    Call GProcSelectBox(txtGrid)
End Sub
Private Sub txtGrid_KeyPress(KeyAscii As Integer)
'Dim LeftPos As Long
'Dim TopPos As Long
'Dim PartyCode As Long
'TopPos = txtGrid.Top + Me.Top + 650
'LeftPos = txtGrid.Left + Me.Left
'DoEvents
'With msGrid
'    If KeyAscii <> 13 And KeyAscii <> 27 Then
'        '--- Search
'        Select Case .Col
'            Case 0 '-- Invoice
'                DoEvents
'                KeyAscii = 0
'                Pending_Invoice
''                PartyCode = GProcGetColumnValue("TblMastAccount", "AcName", txttemp(4), "S", "Accode", "N")
''                '''''gClsSearch.SearchMultiField "tblMastItem,tblMastaccount", "ItName,AcName,ItStdpack,ItStdrateper", Array("Count Name", "Mill", "Std. Pack", "Rate Per"), Array(txtGrid.Width, 2000, 1000, 1000), "tblMastItem.ITMillCode=tblMastaccount.acCode and ITMillCode=" & MillCode, Chr(KeyAscii), "ItName", LeftPos, TopPos, True
''                PendingInvoice
''                txtGrid = ""
''                KeyAscii = 0
''                If gClsSearch.SearchMultiRetCol(0) <> "" Then '-- Item Wt ,Rate From Master
''                    txtGrid = gClsSearch.SearchMultiRetCol(0)
''                    If IsNumeric(.TextMatrix(.Row, 5)) = False Then .TextMatrix(.Row, 5) = 0
''                    If IsNumeric(.TextMatrix(.Row, 7)) = False Then .TextMatrix(.Row, 7) = 0
''                    'If .TextMatrix(.Row, 5) = 0 Then .TextMatrix(.Row, 5) = gClsSearch.SearchMultiRetCol(2)
''                    If .TextMatrix(.Row, 7) = 0 Then .TextMatrix(.Row, 7) = gClsSearch.SearchMultiRetCol(3)
''                    txtGrid_KeyPress 13
''                End If
'            Case Else
'                KeyAscii = GProcValidateKey(rsRecVsSale, msGrid.Col, KeyAscii, txtGrid)
'        End Select
'   End If
'
'    If KeyAscii = 13 Then
'        CheckForNumberTxtGrid
'        '----Adj Amt <=Inv Amt & Adj Dr Note Amt <  Dr Note Amt
'        Select Case .Col
'        Case 4:
'            If Val(txtGrid) > Val(.TextMatrix(.Row, 3)) Then
'                MsgBox "Adjustment Amount is less than Invoice Amount", vbCritical + vbOKOnly, "Adjustment Amount is less than Invoice Amount"
'                DoEvents
'                txtGrid.SetFocus
'                Exit Sub
'            End If
'        Case 6:
'            If Val(txtGrid) > Val(.TextMatrix(.Row, 5)) Then
'                MsgBox "Adjustment Dr Note Amount is less than Dr Note Amount", vbCritical + vbOKOnly, "Adjustment Dr Note Amount is less than Dr Note Amount"
'                DoEvents
'                txtGrid.SetFocus
'                Exit Sub
'            End If
'        End Select
'    End If
'
'    '--- Navigate Row ,Col
'    If KeyAscii = 13 And .Col < rsRecVsSale.Fields.Count - 4 Then
'        .Col = .Col + 1
'    ElseIf KeyAscii = 13 And .Col = rsRecVsSale.Fields.Count - 4 Then
'        If .Row = (.rows - 1) Then
'            .rows = .rows + 1
'            .Row = .Row + 1
''            GridAddNew
'        Else
'            .Row = .Row + 1
'        End If
'        .Col = 0
'    End If
'    '-- On ESC
'    With msGrid
'        If KeyAscii = 27 Then '----Esc Key
'            KeyAscii = 0
'            If (Val(.TextMatrix(.Row, 4)) = 0) And .rows > 2 Then
'                .RemoveItem (.Row)
'                txtGrid.text = ""
'            End If
'            txtGrid.Visible = False
'            txt(18).SetFocus
'            Exit Sub
'        End If
'    End With
'End With
'End Sub
'Private Sub txtGrid_KeyDown(KeyCode As Integer, Shift As Integer)
''--- Delete Row On Ctrl+Y
'If KeyCode = vbKeyY And Shift = 2 And msGrid.rows > 2 Then
'    msGrid = 4
'    msGrid.RemoveItem (msGrid.Row)
'    txtGrid.Visible = False
'    txtGrid.text = ""
'    'CalAmount
'    Exit Sub
'End If
End Sub
Private Sub CheckForNumberGrid()
'Dim RowIndex As Long
'With rsRecVsSale
'For RowIndex = 1 To msGrid.rows - 1
'    For i = 1 To .Fields.Count - 1
'        Select Case .Fields(i).Type
'            Case adInteger, 2, 3 '-- Integer
'                If IsNumeric(msGrid.TextMatrix(RowIndex, i)) = False Then
'                    msGrid.TextMatrix(RowIndex, i) = 0
'                Else
'                    msGrid.TextMatrix(RowIndex, i) = CLng(msGrid.TextMatrix(RowIndex, i))
'                End If
'            Case adDecimal, 131, adNumeric, adDouble '-- Double
'                If IsNumeric(msGrid.TextMatrix(RowIndex, i)) = False Then
'                    msGrid.TextMatrix(RowIndex, i) = Format(0, GProcNumberFormat(.Fields(i).Precision, .Fields(i).NumericScale))
'                Else
'                    msGrid.TextMatrix(RowIndex, i) = Format(msGrid.TextMatrix(RowIndex, i), GProcNumberFormat(.Fields(i).Precision, .Fields(i).NumericScale))
'                End If
'            Case adChar, adVarWChar
'            Case adDate '-- Date
'                If IsDate(msGrid.TextMatrix(RowIndex, i)) = False Then
'                    msGrid.TextMatrix(RowIndex, i) = ""
'                End If
'        End Select
'    Next i
'  Next RowIndex
'End With
End Sub
'---  Check For Number Grid
Private Sub CheckForNumberTxtGrid()
'With rsRecVsSale
'    Select Case .Fields(msGrid.Col).Type
'        Case adInteger, 2, 3 '-- Integer
'            If IsNumeric(txtGrid) = False Then
'                txtGrid = 0
'            Else
'                txtGrid = CLng(txtGrid)
'            End If
'            txtGrid.MaxLength = .Fields(msGrid.Col).Precision
'        Case adDecimal, 131, adNumeric, adDouble '-- Double
'            If IsNumeric(txtGrid) = False Then
'                txtGrid = Format(0, GProcNumberFormat(.Fields(msGrid.Col).Precision, .Fields(msGrid.Col).NumericScale))
'            Else
'                txtGrid = Format(txtGrid, GProcNumberFormat(.Fields(msGrid.Col).Precision, .Fields(msGrid.Col).NumericScale))
'            End If
'            txtGrid.MaxLength = .Fields(msGrid.Col).Precision
'        Case adDate '-- Date
'            If IsDate(txtGrid) = False Then
'                txtGrid = ""
'            Else
'                txtGrid = CDate(txtGrid)
'            End If
'            txtGrid.MaxLength = 10
'        Case Else
'            txtGrid.MaxLength = .Fields(msGrid.Col).DefinedSize
'    End Select
'End With
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
'--- JV
gCn.Execute "delete from tblVoucher where Vtype='JV' and VNo in (select JVNOINT from tblRecVsSale where " _
& " INTRECEIVED>0 and RecVtype='" & frmTypeStr & "' and RecVNo=" & txt(0) & " and  RecVyear= " & gCYear & ")  and " _
 & " Vyear= " & gCYear
'--- Rec Vs Sale
gCn.Execute "delete from tblRecVsSale where RecVtype='" & frmTypeStr & "' and RecVNo=" & txt(0) & " and " _
 & " RecVyear= " & gCYear
 
gCn.Execute "delete from tblOutStanding where vtype='" & frmTypeStr & "' and VNo = " & txt(0) & " and vyear = " & gCYear & ""
 
End Sub

Private Sub FillRelatedRecord()
SetRelatedRS '--tbleInvoiceIRN

If rsOS.EOF Or rsOS.BOF Then
  txtIRN(3) = ""
  txtIRN(4) = ""
  mskDtIRN(5) = "__/__/____"
  txtIRN(6) = ""
Else
  txtIRN(3) = rsOS.Fields("IrnNo") & vbNullString
  txtIRN(4) = rsOS.Fields("AckNo") & vbNullString
  If IsDate(rsOS.Fields("AckDt")) = True Then '--Due Dt
     mskDtIRN(5) = rsOS.Fields("AckDt")
  Else
     mskDtIRN(5) = "__/__/____"
  End If
  txtIRN(6) = rsOS.Fields("eQrCode") & vbNullString
End If
'Dim RowIndex As Long
'Dim ColIndex As Long
''--Fill Grid
'SetGrid
'With rsRecVsSale
'    If Not .EOF Then
'        .MoveFirst
'        RowIndex = 1
'        msGrid.rows = 2
'        Do While Not .EOF
'            For ColIndex = 0 To .Fields.Count - 1
'                Select Case .Fields(ColIndex).Type
'                    Case adInteger, 2, 3 '-- Integer
'                        msGrid.TextMatrix(RowIndex, ColIndex) = CLng(.Fields(ColIndex))
'                    Case adDecimal, 131, adNumeric, adDouble  '-- Double
'                        msGrid.TextMatrix(RowIndex, ColIndex) = Format(.Fields(ColIndex), GProcNumberFormat(.Fields(ColIndex).Precision, .Fields(ColIndex).NumericScale))
'                    Case Else
'                        msGrid.TextMatrix(RowIndex, ColIndex) = .Fields(ColIndex) & vbNullString
'                End Select
'            Next ColIndex
'            RowIndex = RowIndex + 1
'            .MoveNext
'            If Not .EOF Then msGrid.rows = msGrid.rows + 1
'        Loop
'    End If
'End With
End Sub
Private Sub Pending_Invoice()
Dim X As String
Dim LateDayCd As String
Dim vamt As Double
Dim DrAcCode As Long
Dim CrAcCode As Long
Dim rstOutStanding As Recordset
Dim i As Long, j As Long
Dim NewRow As Boolean
Dim TotCol(0, 1) As Long
If Trim(txttemp(4)) <> "" Then
    
    CrAcCode = GProcGetColumnValue("tblMastAccount", "AcName", txttemp(4), "S", "AcCode", "N")
    DrAcCode = GProcGetColumnValue("tblMastAccount", "AcName", txttemp(5), "S", "AcCode", "N")
    If gBackEndDB = gBackEndAccess Then
        LateDayCd = "cdate('" & (CDate(mskDt(2)) + 3) & "') " & " - cdate(outDueDate)"
        X = "select VType, outbillNo,format(vdt,'dd/mm/yyyy'), OutAmount-OutRecAmt, OutDbNtAmt-OutRecDbAmt" _
        & "," & LateDayCd & ", VYear,Vno,OutAmount,OutDbNtAmt from tblOutstanding where (OutAcCode=" & DrAcCode & " or OutAcCode= " & CrAcCode & " ) and " _
        & " ((OutAmount-OutRecAmt) > 0 or (OutDbNtAmt-OutRecDbAmt) > 0) "
    Else
        LateDayCd = "to_date('" & CDate(mskDt(2)) + 3 & "','dd/MM/yyyy') - outDueDate"
        X = "select VType, outbillNo,vdt, OutAmount-OutRecAmt, OutDbNtAmt-OutRecDbAmt" _
        & "," & LateDayCd & ", VYear,Vno,OutAmount,OutDbNtAmt from tblOutstanding where (OutAcCode=" & DrAcCode & " or OutAcCode= " & CrAcCode & " ) and " _
        & " ((OutAmount-OutRecAmt) > 0 or (OutDbNtAmt-OutRecDbAmt) > 0) "
    End If
    '" & IIf(gBackEndDB = gBackEndAccess, "cdate(vdt)", "to_date(vdt,'dd/mm/yyyy')") & "
    ' If FormAction = vbDataActionAddNew Then '-- Record Add
        
   ' Else
    '    X = "select VType, outbillNo, Vdt, OutAmount-OutRecAmt+outrectmp as BalAmt, " _
   '     & " OutDbNtAmt-OutRecDbAmt+Outrecdbtmp as DebitNoteBal, " & LataDayCd & ", VYear,Vno,OutAmount,OutDbNtAmt from tblOutstanding where " _
   '     & " OutAcCode=" & wFromCd & " and ((OutAmount-OutRecAmt+outrectmp) > 0 or (OutDbNtAmt-OutRecDbAmt+Outrecdbtmp) > 0) "
  '  End If
    TotCol(0, 0) = 3
    TotCol(0, 1) = CDbl(txt(6))
    gClsSearch.SearchMultiSelField X, "", Array("Type", "Invoice No", "Date", "Balance Amt", "DbNt Balance", "Due Days", "VYear", "INV Vno", "Invoice Amt", "DbNt Amt"), Array(800, 1200, 1200, 1500, 1500, 800, 0, 0, 1200, 1200), "", "", "Vdt,vno desc", 0, 0, , True, TotOfSelCols:=TotCol
    vamt = CDbl(txt(6))
    With msGrid
        If gClsSearch.SearchMultiSelRetCol(0, 0) <> "" Then
                j = .Row
                 txtGrid = gClsSearch.SearchMultiSelRetCol(0, 0)
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
                    .TextMatrix(j, 3) = gClsSearch.SearchMultiSelRetCol(i, 3) '--Bal Amt
                    .TextMatrix(j, 5) = gClsSearch.SearchMultiSelRetCol(i, 4) '--Db Nt amt
                    .TextMatrix(j, 7) = CLng(gClsSearch.SearchMultiSelRetCol(i, 5)) '--Late Day
                    .TextMatrix(j, 10) = CLng(gClsSearch.SearchMultiSelRetCol(i, 6)) '--Year
                    .TextMatrix(j, 11) = gClsSearch.SearchMultiSelRetCol(i, 7) '--Vno
                     j = j + 1
                Next i
                For i = 1 To .Rows - 1
                    If vamt <= CDbl(.TextMatrix(i, 3)) Then  '- Adj Amt
                        .TextMatrix(i, 4) = vamt
                        .TextMatrix(i, 8) = Format(Round((((Val(.TextMatrix(i, 4)) * 1.5) / 3000) * Val(.TextMatrix(i, 7))), 2), "###0.00")
                        vamt = 0
                        Exit For
                    Else
                        .TextMatrix(i, 4) = CDbl(.TextMatrix(i, 3))
                        .TextMatrix(i, 8) = Format(Round((((Val(.TextMatrix(i, 4)) * 1.5) / 3000) * Val(.TextMatrix(i, 7))), 2), "###0.00")
                        vamt = vamt - CDbl(.TextMatrix(i, 3))
                    End If
                Next i
                .Col = 4
        End If
    End With
Else
    MsgBox "Select Party Compulsary", vbCritical + vbOKOnly, "Select Party"
    DoEvents
    txttemp(4).SetFocus
End If
End Sub
Private Sub CalAmount()
Dim CTaxAmt As Double
Dim STaxAmt As Double
Dim ITaxAmt As Double
Dim Taxable As Double
Taxable = Val(txt(23))
If gfrmHunkStr = "S" Then
    If Val(txt(25)) = 0 Then
       CTaxAmt = Round((Taxable * Val(txt(24))) / 100, 2)
       txt(25) = Format(CTaxAmt, FStr)
    End If
    If Val(txt(27)) = 0 Then
       STaxAmt = Round((Taxable * Val(txt(26))) / 100, 2)
       txt(27) = Format(STaxAmt, FStr)
    End If
    If Val(txt(29)) = 0 Then
       ITaxAmt = Round((Taxable * Val(txt(28))) / 100, 2)
       txt(29) = Format(ITaxAmt, FStr)
    End If
Else
'    CTaxAmt = Round((Taxable * Val(txt(24))) / 100, 2)
'    CTaxAmt = GProcMakeRounding(CTaxAmt)
'    txt(25) = Format(CTaxAmt, FStr)
'
'    STaxAmt = Round((Taxable * Val(txt(26))) / 100, 2)
'    STaxAmt = GProcMakeRounding(STaxAmt)
'    txt(27) = Format(STaxAmt, FStr)
'
'    ITaxAmt = Round((Taxable * Val(txt(28))) / 100, 2)
'    ITaxAmt = GProcMakeRounding(ITaxAmt)
'    txt(29) = Format(ITaxAmt, FStr)
    If Abs(Val(txt(25)) - Round((Taxable * Val(txt(24))) / 100, 2)) > 1 Then
       CTaxAmt = Round((Taxable * Val(txt(24))) / 100, 2)
       CTaxAmt = GProcMakeRounding(CTaxAmt)
       txt(25) = Format(CTaxAmt, FStr)
    Else
       txt(25) = Format(Val(txt(25)), FStr)
    End If
    If Abs(Val(txt(27)) - Round((Taxable * Val(txt(26))) / 100, 2)) > 1 Then
       STaxAmt = Round((Taxable * Val(txt(26))) / 100, 2)
       STaxAmt = GProcMakeRounding(STaxAmt)
       txt(27) = Format(STaxAmt, FStr)
    Else
       txt(27) = Format(Val(txt(27)), FStr)
    End If
    If Abs(Val(txt(29)) - Round((Taxable * Val(txt(28))) / 100, 2)) > 1 Then
       ITaxAmt = Round((Taxable * Val(txt(28))) / 100, 2)
       ITaxAmt = GProcMakeRounding(ITaxAmt)
       txt(29) = Format(ITaxAmt, FStr)
    Else
       txt(29) = Format(Val(txt(29)), FStr)
    End If
    
    
    
    
End If
'txt(6) = Format(Val(txt(23)) + (Val(txt(25)) + Val(txt(27)) + Val(txt(29))), FStr)
If gfrmHunkStr = "T" Then
   txt(14).text = 0.1
   ITaxAmt = Round((Taxable * Val(txt(14))) / 100, 2)
   ITaxAmt = GProcMakeRounding(ITaxAmt)
   txt(6) = Format(ITaxAmt, FStr)
Else
   txt(6) = Format(Val(txt(23)) + (Val(txt(25)) + Val(txt(27)) + Val(txt(29))), FStr)
End If

txt(6) = GProcMakeRounding(txt(6))
'txt(7) = "Ref. " + lblDrCrNote.Caption + " " + txt(34)
'txt(8) = Left(Left(txttemp(5).text, 15) + "-" + CStr(txt(23)) + " +GST-" + CStr(Format(Val(txt(25)) + Val(txt(27)) + Val(txt(29)), FStr)), 40)
'If IsDate(mskDt(12)) Then
'   txt(9) = "Agt - " + txt(11) + " - " + (mskDt(12))
'Else
'   txt(9) = "Agt - " + txt(11)
'End If
End Sub
Private Sub SaveRelatedRecord()
Dim X As String
Dim i As Long
Dim BlAmt As Double
Dim rstTmp As Recordset
BlAmt = 0
If GProcRstOpen(rstTmp, "Select OutAmount from tblOutStanding where vno = " & txt(28) & " and vtype = '" & txt(29) & "' and vyear = " & Val(txt(30)) & "", "R", gCn) > 0 Then
   BlAmt = rstTmp.Fields(0)
End If
'--- Save Grid Records
'--- tblRecSale
gCn.Execute "delete from tblRecVsSale where RecVtype='" & frmTypeStr & "' and RecVNo=" & txt(0) & " and " _
 & " RecVyear= " & gCYear

X = "select * from tblRecVsSale where Recvtype='" & frmTypeStr & "' and Recvno=" & txt(0) & " and " _
& " RecVYear=" & gCYear
GProcRstOpen rsRecVsSale, X, "O"
        If Val(txt(6)) > 0 Then
           rsRecVsSale.AddNew
           rsRecVsSale!AdjNo = GProcGenerateId("tblRecVsSale", "ADJNO", "")
           rsRecVsSale!RecVno = CLng(txt(0)) '--RecVNo
           rsRecVsSale!RecVType = UCase(txt(3)) '--RecVtype
           rsRecVsSale!RecVdt = CDate(txt(2)) '--Vdt(2)
           rsRecVsSale!RecVCtrlNo = 1 'i '--Sr No(3)
           rsRecVsSale!RecVYear = gCYear
           rsRecVsSale!RecAmt = CDec(txt(6))
           rsRecVsSale!InvBillno = txt(11) & vbNullString
           rsRecVsSale!RecVCtrlNo = 1 'i
           rsRecVsSale!InvNo = CLng(txt(30)) '--Invno
           rsRecVsSale!InvTp = txt(31) '--InvTp
           rsRecVsSale!InvDt = mskDt(12) '--Vdt(2)
           rsRecVsSale!InvYear = CLng(txt(32))
           rsRecVsSale!BillAmt = BlAmt
           rsRecVsSale!PartyCode = txt(4)
           rsRecVsSale!AdjAmt = CDec(txt(6))
           rsRecVsSale!DrCrNtNo = txt(34) & vbNullString
           rsRecVsSale.Update
        End If
        DoEvents
 rsRecVsSale.Close
End Sub

Private Sub SetRelatedRS()
'--tbleInvoiceIRN
Dim X As String
Set rsOS = New Recordset
X = "Select * from tbleInvoiceIRN where Vtype='" & frmTypeStr & "' and vno=" & txt(0) & " and VYear=" & gCYear
GProcRstOpen rsOS, X, "O"
End Sub

