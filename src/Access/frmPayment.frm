VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form frmPayment 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00C2E3FC&
   BorderStyle     =   0  'None
   Caption         =   "Payment Entry"
   ClientHeight    =   7200
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   11970
   Icon            =   "frmPayment.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7200
   ScaleWidth      =   11970
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox txt 
      Height          =   375
      Index           =   22
      Left            =   0
      TabIndex        =   72
      Text            =   "22"
      Top             =   480
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.TextBox txt 
      Height          =   375
      Index           =   21
      Left            =   0
      TabIndex        =   71
      Text            =   "21"
      Top             =   0
      Visible         =   0   'False
      Width           =   375
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
      Left            =   6735
      TabIndex        =   69
      Text            =   "txtCrBalance"
      Top             =   1005
      Width           =   1755
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
      Left            =   6735
      TabIndex        =   68
      Text            =   "txtDrBalance"
      Top             =   1440
      Width           =   1755
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
      Height          =   2295
      Left            =   2055
      TabIndex        =   60
      Top             =   3480
      Visible         =   0   'False
      Width           =   5760
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
         Left            =   3120
         TabIndex        =   70
         Top             =   1560
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
         Left            =   4200
         TabIndex        =   65
         Top             =   1560
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
         Left            =   1680
         TabIndex        =   64
         Top             =   1560
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
         Left            =   240
         TabIndex        =   62
         Top             =   1560
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
      Left            =   11145
      Style           =   1  'Graphical
      TabIndex        =   38
      ToolTipText     =   "Last"
      Top             =   1100
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
      Left            =   7845
      MaxLength       =   40
      TabIndex        =   57
      Text            =   "20"
      Top             =   2775
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
      Left            =   8010
      MaxLength       =   40
      TabIndex        =   56
      Text            =   "19"
      Top             =   2670
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
      Left            =   1560
      MaxLength       =   40
      TabIndex        =   9
      Text            =   "txttemp(20)"
      ToolTipText     =   "Press F3 for Creation"
      Top             =   1890
      Width           =   5040
   End
   Begin VB.TextBox txtGrid 
      BackColor       =   &H8000000A&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   315
      TabIndex        =   26
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
      Left            =   8025
      MaxLength       =   40
      TabIndex        =   55
      Text            =   "18"
      Top             =   2430
      Visible         =   0   'False
      Width           =   690
   End
   Begin VB.ListBox List1 
      Height          =   1035
      Left            =   4575
      TabIndex        =   14
      Top             =   2400
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.TextBox txttemp1 
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
      Index           =   7
      Left            =   1560
      MaxLength       =   40
      TabIndex        =   21
      Text            =   "txttemp1(7)"
      Top             =   3270
      Width           =   7005
   End
   Begin VB.TextBox txt 
      Height          =   345
      Index           =   17
      Left            =   12030
      TabIndex        =   54
      Text            =   "17"
      Top             =   6120
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.TextBox txt 
      Height          =   345
      Index           =   16
      Left            =   12030
      TabIndex        =   53
      Text            =   "16"
      Top             =   5640
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.TextBox txttemp1 
      Height          =   345
      Index           =   8
      Left            =   4560
      TabIndex        =   13
      Text            =   "8"
      Top             =   2400
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
      Index           =   8
      Left            =   1560
      MaxLength       =   40
      TabIndex        =   22
      Text            =   "8"
      ToolTipText     =   "Press F3 for Creation & Press F2 for Selection"
      Top             =   3675
      Width           =   7005
   End
   Begin VB.TextBox txt 
      Height          =   345
      Index           =   15
      Left            =   12030
      TabIndex        =   52
      Text            =   "15"
      Top             =   5160
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.TextBox txt 
      Height          =   375
      Index           =   12
      Left            =   12045
      TabIndex        =   51
      Text            =   "12"
      Top             =   3045
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.TextBox txt 
      Height          =   325
      Index           =   10
      Left            =   12015
      TabIndex        =   50
      Text            =   "10"
      Top             =   2760
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
      Left            =   1560
      TabIndex        =   1
      Text            =   "0"
      Top             =   585
      Width           =   1320
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
      Index           =   5
      Left            =   1560
      MaxLength       =   40
      TabIndex        =   7
      Text            =   "txttemp(5)"
      ToolTipText     =   "Press F3 for Creation"
      Top             =   1440
      Width           =   5040
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
      Index           =   4
      Left            =   1560
      MaxLength       =   40
      TabIndex        =   5
      Text            =   "txttemp(4)"
      ToolTipText     =   "Press F3 for Creation"
      Top             =   1005
      Width           =   5040
   End
   Begin VB.TextBox txt 
      Height          =   375
      Index           =   14
      Left            =   11160
      TabIndex        =   44
      Text            =   "14"
      Top             =   6360
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.TextBox txt 
      Alignment       =   1  'Right Justify
      Height          =   360
      Index           =   13
      Left            =   12015
      TabIndex        =   43
      Text            =   "13"
      Top             =   3480
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
      Left            =   10545
      Style           =   1  'Graphical
      TabIndex        =   35
      ToolTipText     =   "Previous"
      Top             =   650
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
      Left            =   11145
      Style           =   1  'Graphical
      TabIndex        =   36
      ToolTipText     =   "Next"
      Top             =   650
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
      Left            =   10545
      Style           =   1  'Graphical
      TabIndex        =   37
      ToolTipText     =   "First"
      Top             =   1100
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
      Left            =   10545
      Style           =   1  'Graphical
      TabIndex        =   34
      ToolTipText     =   "To Find Record"
      Top             =   1550
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
      Left            =   10545
      Style           =   1  'Graphical
      TabIndex        =   32
      Top             =   2460
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
      Left            =   10545
      Style           =   1  'Graphical
      TabIndex        =   27
      ToolTipText     =   "To Add New Record"
      Top             =   3240
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
      Left            =   10545
      Style           =   1  'Graphical
      TabIndex        =   28
      ToolTipText     =   "To Modify Record"
      Top             =   3690
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
      Left            =   10545
      Style           =   1  'Graphical
      TabIndex        =   29
      ToolTipText     =   "To Delete Record"
      Top             =   4140
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
      Left            =   10545
      Style           =   1  'Graphical
      TabIndex        =   31
      Top             =   1995
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
      TabIndex        =   30
      ToolTipText     =   "To Save Record"
      Top             =   4680
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
      Left            =   10545
      Style           =   1  'Graphical
      TabIndex        =   33
      ToolTipText     =   "To Cancel Record"
      Top             =   5130
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
      Left            =   10545
      Style           =   1  'Graphical
      TabIndex        =   39
      ToolTipText     =   "Exit"
      Top             =   5670
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
      Left            =   12045
      Style           =   1  'Graphical
      TabIndex        =   45
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
      Left            =   11970
      Style           =   1  'Graphical
      TabIndex        =   46
      ToolTipText     =   "To Add New Record"
      Top             =   4725
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.TextBox txt 
      Height          =   345
      Index           =   3
      Left            =   12075
      TabIndex        =   49
      Text            =   "3"
      Top             =   1440
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.TextBox txt 
      Alignment       =   1  'Right Justify
      Height          =   360
      Index           =   5
      Left            =   12075
      TabIndex        =   42
      Text            =   "5"
      Top             =   2160
      Visible         =   0   'False
      Width           =   345
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
      Index           =   6
      Left            =   1560
      TabIndex        =   11
      Text            =   "6"
      Top             =   2400
      Width           =   1560
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
      Index           =   9
      Left            =   1560
      MaxLength       =   40
      TabIndex        =   23
      Text            =   "9"
      ToolTipText     =   "Press F3 for Creation & Press F2 for Selection"
      Top             =   4095
      Width           =   7005
   End
   Begin VB.TextBox txt 
      Height          =   325
      Index           =   7
      Left            =   12015
      TabIndex        =   40
      Text            =   "7"
      Top             =   2520
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.TextBox txt 
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
      Index           =   11
      Left            =   5025
      TabIndex        =   15
      Text            =   "11"
      Top             =   2400
      Width           =   1575
   End
   Begin VB.TextBox txt 
      Alignment       =   1  'Right Justify
      Height          =   360
      Index           =   4
      Left            =   12075
      TabIndex        =   41
      Text            =   "4"
      Top             =   1800
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.TextBox txt 
      Height          =   345
      Index           =   2
      Left            =   12075
      TabIndex        =   48
      Text            =   "2"
      Top             =   1080
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
      Left            =   11955
      TabIndex        =   47
      Text            =   "1"
      Top             =   720
      Visible         =   0   'False
      Width           =   555
   End
   Begin MSMask.MaskEdBox mskDt 
      Height          =   360
      Index           =   2
      Left            =   5040
      TabIndex        =   3
      Top             =   585
      Width           =   1560
      _ExtentX        =   2752
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
   Begin MSMask.MaskEdBox mskDt 
      Height          =   360
      Index           =   12
      Left            =   5040
      TabIndex        =   19
      Top             =   2850
      Width           =   1560
      _ExtentX        =   2752
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
   Begin MSMask.MaskEdBox mskDt 
      Height          =   360
      Index           =   13
      Left            =   1560
      TabIndex        =   17
      Top             =   2850
      Width           =   1560
      _ExtentX        =   2752
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
   Begin MSFlexGridLib.MSFlexGrid msGrid 
      Height          =   2025
      Left            =   120
      TabIndex        =   25
      Top             =   5025
      Width           =   9975
      _ExtentX        =   17595
      _ExtentY        =   3572
      _Version        =   393216
      FixedCols       =   0
      RowHeightMin    =   275
      BackColorFixed  =   -2147483647
      ForeColorFixed  =   -2147483634
      BackColorBkg    =   12632256
      Appearance      =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Shape Shape4 
      BorderColor     =   &H00C00000&
      BorderWidth     =   5
      FillColor       =   &H00C00000&
      Height          =   7245
      Left            =   0
      Top             =   -120
      Width           =   11940
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
      Left            =   9360
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
      Height          =   495
      Index           =   5
      Left            =   285
      TabIndex        =   8
      Top             =   1890
      Width           =   1290
   End
   Begin VB.Label lbl1 
      BackStyle       =   0  'Transparent
      Caption         =   "&Inward Details  :"
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
      Left            =   195
      TabIndex        =   24
      Top             =   4710
      Width           =   2175
   End
   Begin VB.Shape Shape3 
      BackColor       =   &H00808080&
      BorderColor     =   &H004D177D&
      FillColor       =   &H00ECFFFF&
      FillStyle       =   0  'Solid
      Height          =   2520
      Left            =   105
      Top             =   4635
      Width           =   10095
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
      Left            =   285
      TabIndex        =   16
      Top             =   2850
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
      TabIndex        =   10
      Top             =   2400
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
      Left            =   3960
      TabIndex        =   18
      Top             =   2850
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
      TabIndex        =   6
      Top             =   1440
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
      TabIndex        =   4
      Top             =   1005
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
      Left            =   4080
      TabIndex        =   2
      Top             =   585
      Width           =   675
   End
   Begin VB.Shape Shape1 
      FillColor       =   &H00BDDFD1&
      FillStyle       =   0  'Solid
      Height          =   5745
      Left            =   10395
      Shape           =   4  'Rounded Rectangle
      Top             =   510
      Width           =   1395
   End
   Begin VB.Label lbl1 
      BackStyle       =   0  'Transparent
      Caption         =   "&Narration    :"
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
      Left            =   285
      TabIndex        =   20
      Top             =   3270
      Width           =   1455
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
      Left            =   3600
      TabIndex        =   12
      Top             =   2400
      Width           =   1095
   End
   Begin VB.Shape Shape2 
      BackColor       =   &H00808080&
      BorderColor     =   &H004D177D&
      FillColor       =   &H00ECFFFF&
      FillStyle       =   0  'Solid
      Height          =   4125
      Left            =   105
      Top             =   465
      Width           =   10080
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
      Width           =   11895
   End
End
Attribute VB_Name = "frmPayment"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Const tblName As String = "tblVoucher"
Const OrderField As String = "Vdt,Vno"
Const IdField As String = "vno"
Const MaxNo As Long = 20
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
      AgCode = GProcGetColumnValue("tblMastAccount", "AcName", txttemp(5), "S", "AgCode", "N")
      If AgCode = 90029 Then '--Mill
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
On Error GoTo ErrorRoutine
Dim FindQStr As String
Dim X As String
Dim Hank As String
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
    Else '-- Oracle
        X = " vNar1 || ' ' || VNar2 from " _
    & " tblvoucher,TblMastAccount,TblMastAccount TblAc1 "
    End If
    FindQStr = "select Vno,Vdt,TblMastAccount.AcName,TblAc1.AcName,Vamt, " & X & " where Vtype='" & frmTypeStr & "' and " _
    & " Vyear=" & gCYear & " and VCtrNo=1 and VAcCode=TblMastAccount.AcCode and VAcOCode=TblAc1.AcCode"
    gClsSearch.SearchMultiField FindQStr, "", Array("Vno", "Vdt", "Account", "Credit A/c", "Amount", "Narration"), Array(700, 1200, 2500, 2500, 1000, 2000), "", " ", "vno,vdt", 0, 0, False, , True
    If gClsSearch.SearchMultiRetCol(0) <> "" Then
        rstbl.MoveFirst
        rstbl.Find "Vno = " & gClsSearch.SearchMultiRetCol(0)
        gClsSearch.SearchMultiRetCol(0) = ""
        ReadFields
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
Case 6 'Add
    FormAction = vbDataActionAddNew
    txt(22) = "A"
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
    FormAction = vbDataActionUpdate
    Hank = ""
    'Call GProcChkPass(frmTypeStr, FormAction, txt(0), txt(2), txtTemp(5), gUserName, Hank)
    'If gChkPassTrueFalse = True Then
        'If mURecModi = True Then
        '    SearchHelp
        '    DoEvents
        txt(22) = "M"
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
    'Else
    '   FormAction = vbDataActionClose
    'End If
Case 8 'Delete
'    FormAction = vbDataActionDelete
'6    Hank = ""
'    Call GProcChkPass(frmTypeStr, FormAction, txt(0), txt(2), txttemp(5), gUserName, Hank)
 '   If gChkPassTrueFalse = True Then
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
  '  End If
Case 10 'Exit
    Unload Me
Case 13 'Save
    txt(21) = CDate(Date)
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
        GProcCreateVoucher FormAction, CLng(txt(0)), 1, CDate(txt(2)), frmTypeStr, CLng(txt(4)), CLng(txt(5)), Val(txt(6)), txt(11), txt(12), txt(13), txt(7), txt(8), txt(9), txt(10), False, txt(20)
        GProcCreateVoucher FormAction, txt(0), 2, txt(2), frmTypeStr, txt(5), txt(4), Val(txt(6)) * (-1), txt(11), txt(12), txt(13), txt(7), txt(8), txt(9), txt(10), False, txt(20)
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
    txtGrid.Visible = False
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
                    .ReportFileName = gReportPath & "rptVouPaymentPrint.rpt"
                    .Action = 1
                End If
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
            If (cmdBtn(13).Enabled = False) Then
                Unload Me
            End If
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
    List1.AddItem "Hundi"
    List1.AddItem "Other"
    Me.Left = Screen.Width / 2 - Me.Width / 2
    frmTypeStr = gfrmTypeStr
    CriteriaStr = " VType = '" & frmTypeStr & "' and Vyear=" & gCYear
    Y = "Select * from " & tblName & " where " & CriteriaStr & " and vctrno = 1"
    If Trim(OrderField) <> "" Then Y = Y & " order by " & OrderField
    GProcRstOpen rstbl, Y, "O"
    Select Case frmTypeStr
    Case "BP" '--Bank
    Case "CP" '-- Cash
        Label6.Caption = "Cash Payments"
        lbl1(10).Visible = False
        mskDt(13).Visible = False
        mskDt(12).Visible = False
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
MillAgCode = GProcGetColumnValue("tblMastAccount", "AcName", txttemp(5), "S", "AgCode", "N", " ACMILLTYPE<>0")
If txttemp(20) = "" And MillAgCode = 90029 Then
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
        gClsSearch.SearchMultiField "QryMastaccount", "Account,AcName,AcCity,AgName", Array("Account", "AcName", "City", "Group"), Array(3000, 0, 2000, 3000), "AgCode=90019", Chr(KeyAscii), "AcName", LeftPos, TopPos, True
        KeyAscii = 0
        If gClsSearch.SearchMultiRetCol(0) <> "" And gClsSearch.SearchMultiRetCol(1) = "" Then
            txttemp(Index) = gClsSearch.SearchMultiRetCol(0)
        Else
            If gClsSearch.SearchMultiRetCol(1) <> "" Then txttemp(Index) = gClsSearch.SearchMultiRetCol(1)
        End If
    Case 5: '--Account
        gClsSearch.SearchMultiField "QryMastaccount", "Account,AcName,AcCity,AgName", Array("Account", "AcName", "City", "Group"), Array(3000, 0, 2000, 3000), "", Chr(KeyAscii), "AcName", LeftPos, TopPos, True
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

'X = "select InvTp, InvBillNo, InvDt, BillAmt, AdjAmt, DbNtAmt, AdjDbNtAmt, LateDays, Interest, IntReceived ,InvYear,InvNo,JVNOINT" _




X = "select InvTp, InvBillNo, InvDt, BillAmt,Discount,AddAmt,  LessAmt, AdjAmt, LateDays, Interest, IntReceived ,InvYear,InvNo,JVNOINT" _
& " from tblRecVsSale where " _
& " Recvtype='" & frmTypeStr & "' and Recvno=" & txt(0) & "  and RecVYear=" & gCYear & " order by RecVCtrlNo"
GProcRstOpen rsRecVsSale, X, "R"
With msGrid
    .Clear
    .FormatString = "<Type |> Bill. No|< Date          |>Bill Amount       |>Discount    |>Add     |>Less    |>Adj. Amoun        |>Late Days|>Late Pay Int|>Int Received|Year|VNo|Bill Bal|jVNO"
    .ColWidth(11) = 0
    .ColWidth(12) = 0
    .ColWidth(13) = 0
    .ColWidth(14) = 0
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
            Case 1, 2, 3, 8, 9, 11: '--Type,Inv No, Inv Dt, Debit Note
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
        'txt(18).SetFocus
        Exit Sub
    End If
    '---- Navigate Row ,Col
    If KeyAscii = 13 Then
        KeyAscii = 0
        If .Col = rsRecVsSale.Fields.Count - 4 Then
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
    If KeyAscii <> 13 And KeyAscii <> 27 And .Col = 0 Then
        DoEvents
        KeyAscii = 0
        Pending_In
    End If
End With
'---- Set TxtGrid
If msGrid.text <> "" Then txtGrid.text = msGrid.text
If msGrid.Col < (rsRecVsSale.Fields.Count - 1) Then
    txtGrid.Move msGrid.CellLeft + msGrid.Left, msGrid.CellTop + msGrid.Top, msGrid.CellWidth, msGrid.CellHeight
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
Case 7:
    With msGrid
        .TextMatrix(.Row, 9) = Format(Round((((Val(.TextMatrix(.Row, 7)) * 1.5) / 3000) * Val(.TextMatrix(.Row, 8))), 2), "###0.00")
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
                'txtGrid_KeyPress 13
            Case Else
                KeyAscii = GProcValidateKey(rsRecVsSale, msGrid.Col, KeyAscii, txtGrid)
        End Select
   End If
    If KeyAscii = 13 Then
        CheckForNumberTxtGrid
        '----Adj Amt <=Inv Amt & Adj Dr Note Amt <  Dr Note Amt
        Select Case .Col
        Case 7:
            If Val(txtGrid) > Val(.TextMatrix(.Row, 3)) Then
                MsgBox "Adjustment Amount is more than Invoice Amount", vbCritical + vbOKOnly, "Adjustment Amount is less than Invoice Amount"
                DoEvents
                txtGrid.SetFocus
                Exit Sub
            End If
        End Select
    End If
    '--- Navigate Row ,Col
    If KeyAscii = 13 And .Col < rsRecVsSale.Fields.Count - 4 Then
        .Col = .Col + 1
    ElseIf KeyAscii = 13 And .Col = rsRecVsSale.Fields.Count - 4 Then
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
            If (Val(.TextMatrix(.Row, 3)) = 0) And .Rows > 3 Then
                .RemoveItem (.Row)
                txtGrid.text = ""
            End If
            txtGrid.Visible = False
            txt(18).SetFocus
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
'           BillDet = BillDet + .TextMatrix(i, 1) + "-" + CStr(Val(.TextMatrix(i, 4)) + Val(.TextMatrix(i, 6)) + Val(.TextMatrix(i, 9))) + ":"
           BillDet = BillDet + .TextMatrix(i, 1) + "-" + CStr(Val(.TextMatrix(i, 7)) + Val(.TextMatrix(i, 10))) + ":"
        End If
        InvNo = .TextMatrix(i, 12)
        InVtype = .TextMatrix(i, 0)
        InvYear = .TextMatrix(i, 11)
        For j = 1 To .Rows - 1
            If j <> i And InvNo = .TextMatrix(j, 12) And InVtype = .TextMatrix(j, 0) And InvYear = .TextMatrix(j, 11) And InvNo <> 0 Then
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
           If CDbl(.TextMatrix(i, 7)) > CDbl(.TextMatrix(i, 3)) Then
                MsgBox "Adjusted Amount is more than Balance Bill Amount. ", vbOKOnly, Me.Caption
                msGrid.SetFocus
                Exit Function
           End If
           TotAmt = TotAmt + Val(.TextMatrix(i, 7)) + Val(.TextMatrix(i, 10))
        Next i
        If (TotAmt > 0) And (TotAmt <> Val(txt(6))) Then
           MsgBox "Total of Adjusted Amount " & TotAmt & " & Paid Amount not tally.", vbOKOnly, Me.Caption
        End If
    End If
End With
ValidateGrid = True
End Function
Private Sub DeleteRelatedRecord()
'--- JV Int

gCn.Execute "delete from tblVoucher where Vtype='JV' and VNo in (select JVNOINT from tblRecVsSale where " _
& " INTRECEIVED>0 and RecVtype='" & frmTypeStr & "' and RecVNo=" & txt(0) & " and  RecVyear= " & gCYear & ")  and " _
 & " Vyear= " & gCYear

DoEvents

'--- Rec Vs Sale


gCn.Execute "delete from tblRecVsSale where RecVtype='" & frmTypeStr & "' and RecVNo=" & txt(0) & " and " _
 & " RecVyear= " & txt(17)
End Sub
Private Sub SaveRelatedRecord()
Dim i As Long, X As String
Dim rsRecVsSale As Recordset
Dim JVVNo As Long
Dim JVTDSVNo As Long '- Less TDS JV VNo
Dim IntRecAmt As Double



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
        rsRecVsSale!InvBillno = CLng(.TextMatrix(i, 1)) '--Inv Bill No
        rsRecVsSale!InvNo = CLng(.TextMatrix(i, 12)) '--Invno
        rsRecVsSale!InvTp = .TextMatrix(i, 0) '--InvTp
        rsRecVsSale!InvDt = CDate(.TextMatrix(i, 2)) '--InvDt
        rsRecVsSale!InvYear = .TextMatrix(i, 11)
        rsRecVsSale!BillAmt = CDec(.TextMatrix(i, 3))
        rsRecVsSale!Discount = CDec(.TextMatrix(i, 4)) '- Discount
        rsRecVsSale!AddAmt = CDec(.TextMatrix(i, 5)) '-Add
        rsRecVsSale!LessAmt = CDec(.TextMatrix(i, 6)) 'Less
        rsRecVsSale!PartyCode = GProcGetColumnValue("tblMastAccount", "AcName", txttemp(5), "S", "AcCode", "N")
        rsRecVsSale!AdjAmt = CDec(.TextMatrix(i, 7))
        rsRecVsSale!LateDays = CInt(.TextMatrix(i, 8))
        rsRecVsSale!Interest = CDec(.TextMatrix(i, 9))
        rsRecVsSale!IntReceived = CDec(Val(.TextMatrix(i, 10)))
        
        IntRecAmt = CDec(Val(.TextMatrix(i, 10)))

        
        JVVNo = 0
        JVTDSVNo = 0
        '------------------------- Interest Rec JV
        If FormAction = vbDataActionAddNew And IntRecAmt > 0 Then
            JVVNo = GProcGenerateId("tblVoucher", "Vno", " vyear=" & gCYear & " and Vtype='JV'")
        Else
            If IntRecAmt > 0 Then
                If CLng(.TextMatrix(i, 13)) > 0 And GProcGetColumnValue("tblVoucher", "Vno", .TextMatrix(i, 13), "N", "Vno", "N", " Vtype='JV' and Vyear=" & gCYear) = 0 Then
                    JVVNo = CLng(.TextMatrix(i, 13))
                Else
                    JVVNo = GProcGenerateId("tblVoucher", "Vno", " vyear=" & gCYear & " and Vtype='JV'")
                End If
            End If
        End If
        If IntRecAmt > 0 Then
            GProcCreateVoucher FormAction, JVVNo, 1, mskDt(2), "JV", txt(5), gLatePayIntAcCode, IntRecAmt, IsRefEntryType:=True
            GProcCreateVoucher FormAction, JVVNo, 2, mskDt(2), "JV", gLatePayIntAcCode, txt(5), -IntRecAmt, IsRefEntryType:=True
        End If
        '-------------------------------
        
        rsRecVsSale!JVNOINT = JVVNo
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
If Trim(txttemp(5)) <> "" Then
    wFromCd = GProcGetColumnValue("tblMastAccount", "AcName", txttemp(5), "S", "AcCode", "N")
    If gBackEndDB = gBackEndAccess Then
'        DateCd = " and vdt<=cdate('" & CDate(mskDt(2)) & "')"
    Else '-- Oracle
        DateCd = " and vdt<= to_date('" & CDate(mskDt(2)) & "','dd/MM/yyyy')"
    End If
    If gBackEndDB = gBackEndAccess Then
        LateDayCd = "cdate('" & (CDate(mskDt(2)) + 3) & "') " & "- cdate(outDueDate)"
'        X = "select VType, outbillNo,format(vdt,'dd/mm/yyyy'),OutAmount-OutRecAmt , OutAmount-OutRecAmt , " & LateDayCd & ", VYear,Vno from tblOutstanding where OutAcCode=" & wFromCd & DateCd & " and (OutAmount-OutRecAmt) > 0 "
        X = "select VType, outbillNo,format(vdt,'dd/mm/yyyy'),OutAmount, OutAmount-OutRecAmt , " & LateDayCd & ", VYear,Vno from tblOutstanding where OutAcCode=" & wFromCd & DateCd & " and (OutAmount-OutRecAmt) > 0 "
    Else
        LateDayCd = "to_date('" & CDate(mskDt(2)) + 3 & "','dd/MM/yyyy') - outDueDate"
         X = "select VType,outbillNo,Vdt,OutAmount-OutRecAmt , OutAmount-OutRecAmt , " & LateDayCd & ", VYear,Vno from tblOutstanding where OutAcCode=" & wFromCd & DateCd & " and (OutAmount-OutRecAmt) > 0 "
    End If
   ' If FormAction = vbDataActionAddNew Then
      
   ' Else
   '     X = "select VType, Vno,Vdt,OutAmount(OutRecAmt- OutRecTmp), OutAmount-(OutRecAmt- OutRecTmp) ," & LataDayCd & " , VYear from tblOutstanding where OutAcCode=" & wFromCd & DateCd & " and (OutAmount-(OutRecAmt- OutRecTmp)) > 0 "
   ' End If
    vamt = CDbl(txt(6))
    gClsSearch.SearchMultiSelField X, "", Array("In.Type", "Bill .No", "Date", "Amount", "Balance Amt", "Due Date", "Year", "Vno"), Array(900, 900, 1200, 1500, 1500, 1000, 0, 1000), "", "", "Vdt,Vno", 0, 0, , True
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
                .TextMatrix(j, 12) = gClsSearch.SearchMultiSelRetCol(i, 7) '-VNo
                .TextMatrix(j, 2) = Format(gClsSearch.SearchMultiSelRetCol(i, 2), "dd/mm/yyyy") '--Vdt
                .TextMatrix(j, 3) = gClsSearch.SearchMultiSelRetCol(i, 3) '--Amt
                .TextMatrix(j, 14) = gClsSearch.SearchMultiSelRetCol(i, 4) '--Bal
                .TextMatrix(j, 8) = gClsSearch.SearchMultiSelRetCol(i, 5) '-- Late Days
                .TextMatrix(j, 11) = gClsSearch.SearchMultiSelRetCol(i, 6) '--Year
                j = j + 1
            Next i
             For i = 1 To .Rows - 1 '- Adj Amt
                If vamt <= Val(.TextMatrix(i, 14)) Then
                    .TextMatrix(i, 7) = vamt
                    .TextMatrix(.Row, 9) = Format(Round((((Val(.TextMatrix(.Row, 7)) * 1.5) / 3000) * Val(.TextMatrix(.Row, 8))), 2), "###0.00")
                    vamt = 0
                    Exit For
                Else
                    .TextMatrix(i, 7) = CDbl(.TextMatrix(i, 14))
                    .TextMatrix(.Row, 9) = Format(Round((((Val(.TextMatrix(.Row, 7)) * 1.5) / 3000) * Val(.TextMatrix(.Row, 8))), 2), "###0.00")
                    vamt = vamt - CDbl(.TextMatrix(i, 14))
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
