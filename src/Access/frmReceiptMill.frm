VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form frmReceiptMill 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00C2E3FC&
   BorderStyle     =   0  'None
   Caption         =   "Receipt Entry"
   ClientHeight    =   7245
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   11445
   Icon            =   "frmReceiptMill.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7245
   ScaleWidth      =   11445
   ShowInTaskbar   =   0   'False
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
      Left            =   2850
      TabIndex        =   62
      Top             =   4380
      Visible         =   0   'False
      Width           =   5175
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
         TabIndex        =   67
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
         Left            =   1935
         TabIndex        =   66
         Top             =   1560
         Width           =   1335
      End
      Begin VB.TextBox txtno 
         Appearance      =   0  'Flat
         Height          =   360
         Index           =   1
         Left            =   1950
         MaxLength       =   7
         TabIndex        =   65
         Top             =   1080
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
         TabIndex        =   64
         Top             =   1560
         Width           =   1335
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
         TabIndex        =   68
         Top             =   600
         Width           =   1065
      End
   End
   Begin VB.ListBox List1 
      Height          =   1035
      Left            =   1530
      TabIndex        =   9
      Top             =   2820
      Visible         =   0   'False
      Width           =   1335
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
      Index           =   11
      Left            =   1920
      TabIndex        =   8
      Text            =   "11"
      Top             =   2805
      Width           =   1575
   End
   Begin VB.TextBox txttemp1 
      Height          =   345
      Index           =   8
      Left            =   1515
      TabIndex        =   7
      Text            =   "8"
      Top             =   2805
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
      Left            =   6225
      TabIndex        =   58
      Text            =   "txtCrBalance"
      Top             =   1230
      Width           =   2055
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
      Left            =   6225
      TabIndex        =   57
      Text            =   "txtDrBalance"
      Top             =   1650
      Width           =   2055
   End
   Begin VB.CommandButton cmdLedger 
      Caption         =   "Ledger"
      Height          =   375
      Index           =   0
      Left            =   8445
      TabIndex        =   56
      Top             =   1230
      Width           =   1200
   End
   Begin VB.CommandButton cmdLedger 
      Caption         =   "Ledger"
      Height          =   375
      Index           =   1
      Left            =   8445
      TabIndex        =   55
      Top             =   1650
      Width           =   1200
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
      Top             =   5445
      Visible         =   0   'False
      Width           =   885
   End
   Begin VB.TextBox txt 
      Height          =   325
      Index           =   18
      Left            =   11745
      MaxLength       =   40
      TabIndex        =   54
      Text            =   "18"
      Top             =   1755
      Visible         =   0   'False
      Width           =   270
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
      ForeColor       =   &H00000080&
      Height          =   375
      Index           =   7
      Left            =   5145
      MaxLength       =   20
      TabIndex        =   11
      Text            =   "txttemp1(7)"
      Top             =   2370
      Width           =   4500
   End
   Begin VB.TextBox txt 
      Height          =   345
      Index           =   17
      Left            =   11580
      TabIndex        =   53
      Text            =   "17"
      Top             =   6000
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.TextBox txt 
      Height          =   345
      Index           =   16
      Left            =   11580
      TabIndex        =   52
      Text            =   "16"
      Top             =   5520
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
      ForeColor       =   &H00000080&
      Height          =   375
      Index           =   8
      Left            =   5145
      MaxLength       =   40
      TabIndex        =   12
      Text            =   "8"
      ToolTipText     =   "Press F3 for Creation & Press F2 for Selection"
      Top             =   2805
      Width           =   4500
   End
   Begin VB.TextBox txt 
      Height          =   345
      Index           =   15
      Left            =   11580
      TabIndex        =   51
      Text            =   "15"
      Top             =   5040
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.TextBox txt 
      Height          =   375
      Index           =   12
      Left            =   11550
      TabIndex        =   49
      Text            =   "12"
      Top             =   3090
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.TextBox txt 
      Height          =   325
      Index           =   10
      Left            =   11640
      TabIndex        =   48
      Text            =   "10"
      Top             =   2850
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
      Left            =   1515
      TabIndex        =   1
      Text            =   "0"
      Top             =   570
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
      Left            =   1515
      MaxLength       =   40
      TabIndex        =   4
      Text            =   "txttemp(5)"
      ToolTipText     =   "Press F3 for Creation"
      Top             =   1230
      Width           =   4590
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
      Left            =   1515
      MaxLength       =   40
      TabIndex        =   5
      Text            =   "txttemp(4)"
      ToolTipText     =   "Press F3 for Creation"
      Top             =   1650
      Width           =   4575
   End
   Begin VB.TextBox txt 
      Height          =   375
      Index           =   14
      Left            =   11580
      TabIndex        =   41
      Text            =   "14"
      Top             =   3720
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.TextBox txt 
      Alignment       =   1  'Right Justify
      Height          =   360
      Index           =   13
      Left            =   11685
      TabIndex        =   40
      Text            =   "13"
      Top             =   3315
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
      Left            =   10155
      Style           =   1  'Graphical
      TabIndex        =   29
      ToolTipText     =   "Previous"
      Top             =   1125
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
      Left            =   10680
      Style           =   1  'Graphical
      TabIndex        =   30
      ToolTipText     =   "Next"
      Top             =   1125
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
      Left            =   10155
      Style           =   1  'Graphical
      TabIndex        =   31
      ToolTipText     =   "First"
      Top             =   1575
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
      Left            =   10680
      Style           =   1  'Graphical
      TabIndex        =   32
      ToolTipText     =   "Last"
      Top             =   1575
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
      Left            =   10155
      Style           =   1  'Graphical
      TabIndex        =   28
      ToolTipText     =   "To Find Record"
      Top             =   2025
      Width           =   975
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
      Left            =   9600
      Style           =   1  'Graphical
      TabIndex        =   33
      Top             =   330
      Visible         =   0   'False
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
      Left            =   10155
      Style           =   1  'Graphical
      TabIndex        =   23
      ToolTipText     =   "To Add New Record"
      Top             =   2670
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
      Left            =   10155
      Style           =   1  'Graphical
      TabIndex        =   24
      ToolTipText     =   "To Modify Record"
      Top             =   3120
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
      Left            =   10155
      Style           =   1  'Graphical
      TabIndex        =   25
      ToolTipText     =   "To Delete Record"
      Top             =   3570
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
      Left            =   10155
      Style           =   1  'Graphical
      TabIndex        =   34
      Top             =   4215
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
      Left            =   10155
      Style           =   1  'Graphical
      TabIndex        =   26
      ToolTipText     =   "To Save Record"
      Top             =   4845
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
      Left            =   10155
      Style           =   1  'Graphical
      TabIndex        =   27
      ToolTipText     =   "To Cancel Record"
      Top             =   5295
      Width           =   975
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
      Left            =   10155
      Style           =   1  'Graphical
      TabIndex        =   35
      ToolTipText     =   "Exit"
      Top             =   5910
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
      Left            =   11460
      Style           =   1  'Graphical
      TabIndex        =   42
      Top             =   4170
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
      Left            =   11520
      Style           =   1  'Graphical
      TabIndex        =   43
      Top             =   4605
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.TextBox txt 
      Height          =   345
      Index           =   3
      Left            =   11580
      TabIndex        =   46
      Text            =   "3"
      Top             =   1320
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.TextBox txt 
      Alignment       =   1  'Right Justify
      Height          =   360
      Index           =   5
      Left            =   11580
      TabIndex        =   39
      Text            =   "5"
      Top             =   2040
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
      Left            =   1515
      TabIndex        =   6
      Text            =   "6"
      Top             =   2370
      Width           =   1500
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
      ForeColor       =   &H00000080&
      Height          =   375
      Index           =   9
      Left            =   5145
      MaxLength       =   40
      TabIndex        =   13
      Text            =   "9"
      ToolTipText     =   "Press F3 for Creation & Press F2 for Selection"
      Top             =   3240
      Width           =   4500
   End
   Begin VB.TextBox txt 
      Height          =   325
      Index           =   7
      Left            =   11670
      TabIndex        =   37
      Text            =   "7"
      Top             =   2610
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
      Left            =   11580
      TabIndex        =   36
      Text            =   "19"
      Top             =   2580
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.TextBox txt 
      Alignment       =   1  'Right Justify
      Height          =   360
      Index           =   4
      Left            =   11580
      TabIndex        =   38
      Text            =   "4"
      Top             =   1680
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.TextBox txt 
      Height          =   345
      Index           =   2
      Left            =   11580
      TabIndex        =   45
      Text            =   "2"
      Top             =   960
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
      Left            =   11460
      TabIndex        =   44
      Text            =   "1"
      Top             =   600
      Visible         =   0   'False
      Width           =   555
   End
   Begin MSMask.MaskEdBox mskDt 
      Height          =   360
      Index           =   2
      Left            =   4785
      TabIndex        =   2
      Top             =   570
      Width           =   1320
      _ExtentX        =   2328
      _ExtentY        =   635
      _Version        =   393216
      ClipMode        =   1
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
      Left            =   1515
      TabIndex        =   10
      Top             =   3240
      Visible         =   0   'False
      Width           =   1440
      _ExtentX        =   2540
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
      Left            =   8445
      TabIndex        =   3
      Top             =   570
      Width           =   1200
      _ExtentX        =   2117
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
      Height          =   3345
      Left            =   240
      TabIndex        =   15
      Top             =   3780
      Width           =   9600
      _ExtentX        =   16933
      _ExtentY        =   5900
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
   Begin Crystal.CrystalReport cryReport 
      Bindings        =   "frmReceiptMill.frx":058A
      Left            =   0
      Top             =   0
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
      Caption         =   "Party        :"
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
      TabIndex        =   61
      Top             =   1650
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
      Left            =   3930
      TabIndex        =   60
      Top             =   3180
      Width           =   1095
   End
   Begin VB.Label lbl1 
      BackStyle       =   0  'Transparent
      Caption         =   "Ban&k Name  :"
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
      Left            =   3840
      TabIndex        =   59
      Top             =   2370
      Width           =   1365
   End
   Begin VB.Shape Shape6 
      BorderColor     =   &H00C00000&
      BorderWidth     =   5
      FillColor       =   &H00C00000&
      Height          =   7230
      Left            =   15
      Top             =   0
      Width           =   11430
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
      Height          =   465
      Left            =   8535
      TabIndex        =   50
      Top             =   45
      Width           =   2940
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
      Top             =   570
      Width           =   1035
   End
   Begin VB.Label lbl1 
      BackStyle       =   0  'Transparent
      Caption         =   "&Rcon.Date    :"
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
      Left            =   7125
      TabIndex        =   17
      Top             =   570
      Width           =   1335
   End
   Begin VB.Label lbl1 
      BackStyle       =   0  'Transparent
      Caption         =   "Am&ount   :"
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
      Top             =   2370
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
      Top             =   3240
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.Label lbl1 
      BackStyle       =   0  'Transparent
      Caption         =   "Cre&dit       :"
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
      Caption         =   "Da&te    :"
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
      Left            =   3885
      TabIndex        =   16
      Top             =   570
      Width           =   735
   End
   Begin VB.Shape Shape1 
      FillColor       =   &H00BDDFD1&
      FillStyle       =   0  'Solid
      Height          =   5790
      Left            =   9975
      Shape           =   4  'Rounded Rectangle
      Top             =   795
      Width           =   1335
   End
   Begin VB.Label Label6 
      BackColor       =   &H00C00000&
      Caption         =   "  Mill Receipts"
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
      Left            =   15
      TabIndex        =   47
      Top             =   15
      Width           =   11955
   End
   Begin VB.Label lbl1 
      BackStyle       =   0  'Transparent
      Caption         =   "Narration      :"
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
      Left            =   3840
      TabIndex        =   22
      Top             =   2805
      Width           =   1455
   End
   Begin VB.Label lbl1 
      BackStyle       =   0  'Transparent
      Caption         =   "Re&f No.     :"
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
      Top             =   2805
      Width           =   1095
   End
   Begin VB.Shape Shape2 
      BackColor       =   &H00808080&
      BorderColor     =   &H00000080&
      FillColor       =   &H00ECFFFF&
      FillStyle       =   0  'Solid
      Height          =   585
      Left            =   240
      Top             =   450
      Width           =   9600
   End
   Begin VB.Shape Shape3 
      BackColor       =   &H00ECFFFE&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00000080&
      FillColor       =   &H00ECFFFF&
      Height          =   1050
      Left            =   240
      Top             =   1095
      Width           =   9600
   End
   Begin VB.Shape Shape5 
      BorderColor     =   &H00000080&
      FillColor       =   &H00ECFFFE&
      FillStyle       =   0  'Solid
      Height          =   1545
      Left            =   240
      Top             =   2190
      Width           =   3420
   End
   Begin VB.Shape Shape4 
      BorderColor     =   &H00000080&
      FillColor       =   &H00ECFFFE&
      FillStyle       =   0  'Solid
      Height          =   1530
      Left            =   3735
      Top             =   2190
      Width           =   6105
   End
End
Attribute VB_Name = "frmReceiptMill"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Const tblName As String = "TblMillRecPay"
Const IdField As String = "Vno"
Const MaxNo As Long = 18
Dim CriteriaStr As String
Dim Rstbl As Recordset, ShowFrmFlag As Boolean
Dim varBookMark As Variant
Dim FirstTimeFlag As Boolean, mLastEntryBank As String, mLastEntryDate As Date
Dim FormAction As Integer
Dim frmTypeStr As String
Dim i As Long
'-- RecVsSale
Dim rsRecVsSale As Recordset
Dim OldgFrmTypeStr As String
Const OrderField As String = "Vdt,Vno"
Private Sub cmdBtn_Click(Index As Integer)
Dim FindQStr As String
Dim x As String
Dim Hank As String
'On Error GoTo ErrorRoutine
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
        x = " MRecNar1 & ' ' & MRecNar2 from " _
    & " tblmillrecpay,TblMastAccount,TblMastAccount as TblAc1 "
    Else '-- Oracle
        x = " MRecNar1 || ' ' || MRecNar2 from " _
    & " tblmillrecpay,TblMastAccount,TblMastAccount TblAc1 "
    End If
    FindQStr = "select Vno,Vdt,TblMastAccount.AcName,TblAc1.AcName,MRecAmt," & x & " where Vtype='" & frmTypeStr & "' and " _
    & " Vyear=" & gCYear & " and VCtrNo=1 and MRECACCODE=TblMastAccount.AcCode and MRECACOCODE=TblAc1.AcCode"
    gClsSearch.SearchMultiField FindQStr, "", Array("Vno", "Vdt", "Account", "Credit A/c", "Amount", "Narration"), Array(700, 1200, 2500, 2500, 1000, 2000), "", " ", "vno,vdt", 0, 0, False, , True
    If gClsSearch.SearchMultiRetCol(0) <> "" Then
        Rstbl.MoveFirst
        Rstbl.Find "Vno = " & gClsSearch.SearchMultiRetCol(0)
        gClsSearch.SearchMultiRetCol(0) = ""
        ReadFields
    End If
Case 5 'List
    'GenMastList ("LAC")
Case 6 'Add
    FormAction = vbDataActionAddNew
    Call GProcActivateControls(Me, True)
    Call GProcSetButtons(Me, vbDataActionAddNew)
    Call GProcClearForm(Me, Rstbl, MaxNo, True)
    txt(0).Text = GProcGenerateId(tblName, "VNO", CriteriaStr, gCn)
    mskDt(2) = GProcGenerateVDt(tblName, "VDt", CriteriaStr)
    If (mLastEntryBank <> "") Then
       txtTemp(5).Text = mLastEntryBank
       mskDt(2) = mLastEntryDate
    End If
    If FirstTimeFlag = True Then
        FirstTimeFlag = False
    Else
        txt(0).SetFocus
    End If
    SetGrid
Case 7 'Modify
    FormAction = vbDataActionUpdate
    Hank = ""
    Call GProcChkPass(frmTypeStr, FormAction, txt(0), txt(2), txtTemp(4), gUserName, Hank)
    If gChkPassTrueFalse = True Then
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
    Else
       FormAction = vbDataActionClose
    End If
Case 8 'Delete
    FormAction = vbDataActionDelete
    Hank = ""
    Call GProcChkPass(frmTypeStr, FormAction, txt(0), txt(2), txtTemp(4), gUserName, Hank)
    If gChkPassTrueFalse = True Then
        ' If mURecDel = True Then
          If IsRelatedRecord = False Then
             If MsgBox("Do You Really Delete This Record?", vbQuestion + vbYesNo + vbDefaultButton1, "Confirm Delete") = vbYes Then
                gCn.Execute "Delete from tblmillrecpay where " & CriteriaStr & " and vno = " & txt(0)
                '--Update Outstanding Balance
                GProcUpdateOutstandingBal CLng(txt(0)), UCase(txt(3)), gCYear, False, True
                 '-- Delete Related Record
                 DeleteRelatedRecord
                 If Rstbl.RecordCount > 1 Then
                     Rstbl.Requery
                     cmdBtn_Click (1)
                 ElseIf Rstbl.RecordCount = 1 Then
                     Call GProcActivateControls(Me, False)
                     Call GProcSetButtons(Me, vbDataActionCancel)
                     Form_Load
                 End If
             End If
             DoEvents
          End If
        ' End If
     End If
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
            msGrid.Text = txtGrid
            DoEvents
            txtGrid = ""
        End If
        txtGrid.Visible = False
        FillTxtFromTemp
   '     GProcSaveRecord Me, rstbl, gFormAction
        gCn.Execute "Delete from tblmillrecpay where " & CriteriaStr & " and vno = " & txt(0)
        GProcCreateMillRecPay FormAction, txt(0), 1, txt(2), frmTypeStr, txt(4), txt(5), Val(txt(6)), txt(11), txt(12), txt(13), txt(7), txt(8), txt(9), txt(10)
        GProcCreateMillRecPay FormAction, txt(0), 2, txt(2), frmTypeStr, txt(5), txt(4), Val(txt(6)) * (-1), txt(11), txt(12), txt(13), txt(7), txt(8), txt(9), txt(10)
        '--- Related Rcord Rec Vs Sale
        SaveRelatedRecord
        '--Update Outstanding Balance
        GProcUpdateOutstandingBal CLng(txt(0)), UCase(txt(3)), gCYear
        Rstbl.Requery
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
        LblBalance.Caption = ""
    End If
Case 14 'Cancel
    txtGrid.Visible = False
    Call GProcClearForm(Me, Rstbl, MaxNo, True)
    If Rstbl.RecordCount <= 0 Then
        Call GProcActivateControls(Me, False)
        Call GProcSetButtons(Me, vbDataActionCancel)
    Else
        If FormAction = vbDataActionAddNew Then
            Rstbl.MoveLast
        Else
             '--Update Outstanding Balance
            GProcUpdateOutstandingBal CLng(txt(0)), UCase(txt(3)), gCYear
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
    msGrid.Col = 0
    If msGrid.Row >= 1 Then msGrid.Row = 1
    txtGrid.Move msGrid.CellLeft + msGrid.Left, msGrid.CellTop + msGrid.Top, msGrid.CellWidth, msGrid.CellHeight
    LblBalance.Caption = ""
End Select
Exit Sub
ErrorRoutine:
    Rstbl.CancelUpdate
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
Dim x As String
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
        x = "Select * From TblMillRecPay where  vNo=" & Vno & " and VCTRNO =1 and VType='" & frmTypeStr & "' and VYear=" & gCYear
        If GProcRstOpen(Rs1, x, "R") > 0 Then
            With frmMain.cryReport
                .SelectionFormula = "{TblMillRecPay.vNo} = " & Rs1!Vno & " and {TblMillRecPay.VCTRNO}=1 and {TblMillRecPay.VType}='" & Rs1!Vtype & "' and {TblMillRecPay.VYear}=" & Rs1!VYear
                .Formulas(5) = "wAmountInWords = '" & GProcAmountInWords(Abs(Rs1!MRecAmt)) & "'"
                .ReportFileName = gReportPath & "rptVouMillReceiptPrint.rpt"
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
Private Sub cmdLedger_Click(Index As Integer)
'On Error GoTo ErrorHandler
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
            SendKeys "{TAB}"
        Case 27:
'            If (cmdBtn(13).Enabled = False) Then
'                Unload Me
'            End If
    End Select
End Sub
Private Sub Form_Load()
    Dim x As String
    Dim Y As String
    mLastEntryBank = ""
    Me.Top = 300
    List1.Clear
    List1.AddItem "Cheque"
    List1.AddItem "Transfer"
    List1.AddItem "DD."
    List1.AddItem "Hundi"
    List1.AddItem "Other"
    
    LblBalance.Caption = ""
    Me.Left = Screen.Width / 2 - Me.Width / 2
    frmTypeStr = gfrmTypeStr
    CriteriaStr = " VType = '" & frmTypeStr & "' and VYear=" & gCYear
    Y = "Select * from " & tblName & " where " & CriteriaStr & " and vctrno = 1"
    If Trim(OrderField) <> "" Then Y = Y & " order by " & OrderField
    GProcRstOpen Rstbl, Y, "O"
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
Rstbl.Close
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
GProcCheckForNumber Me, Rstbl, MaxNo
SaveBtnEd
End Sub
Private Sub txt_KeyPress(Index As Integer, KeyAscii As Integer)
    KeyAscii = GProcValidateKey(Rstbl, Index, KeyAscii, txt(Index).Text)
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
'If Left(frmTypeStr, 1) = "B" Then
   txt(7) = Left(txttemp1(8) + Space(3), 3) + ". No. " + Left(txt(11) + Space(10), 10) + " " + txttemp1(7)
'Else
'   txt(7) = txttemp1(7)
'End If
txt(4) = GProcGetColumnValue("tblMastAccount", "AcName", txtTemp(4), "S", "Accode", "N")
txt(5) = GProcGetColumnValue("tblMastAccount", "AcName", txtTemp(5), "S", "AcCode", "N")
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
'If Left(frmTypeStr, 1) = "B" Then
    txttemp1(7) = Mid(txt(7), 21)
    txttemp1(8) = Left(txt(7), 3)
'Else
'    txttemp1(7) = txt(7)
'End If
txtTemp(4) = GProcGetColumnValue("tblMastAccount", "AcCode", txt(4), "N", "AcName", "S")
txtTemp(5) = GProcGetColumnValue("tblMastAccount", "AcCode", txt(5), "N", "AcName", "S")
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
If Trim(txtGrid.Text) <> "" Then
        msGrid.Text = txtGrid.Text
End If
txtGrid.Visible = False
txtGrid.Text = ""
CheckForNumberGrid
With msGrid
    '-- Duplicate invoice selection
    For i = 1 To .Rows - 1
        '-- Int A/c
           If CDbl(.TextMatrix(i, 9)) > 0 Then
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
        InvNo = .TextMatrix(i, 11)
        InVtype = .TextMatrix(i, 0)
        InvYear = .TextMatrix(i, 10)
        For j = 1 To .Rows - 1
            If j <> i And InvNo = .TextMatrix(j, 11) And InVtype = .TextMatrix(j, 0) And InvYear = .TextMatrix(j, 10) And InvNo <> 0 Then
                MsgBox "Duplicate Invoice Selection.", vbOKOnly, Me.Caption
                msGrid.SetFocus
                Exit Function
            End If
        Next j
    Next i
    '-- Adjusted Amt
    If .Rows > 1 Then
       TotAmt = 0
       For i = 1 To .Rows - 1
           '-- Adj amt > Bill Amt
           If CDbl(.TextMatrix(i, 4)) > CDbl(.TextMatrix(i, 3)) Then
                MsgBox "Adjusted Amount is more than Balance Bill Amount. ", vbOKOnly, Me.Caption
                msGrid.SetFocus
                Exit Function
           End If
            '-- Adj DB amt > Db Amt
           If CDbl(.TextMatrix(i, 6)) > CDbl(.TextMatrix(i, 5)) Then
                MsgBox "Adjusted DbNt Amount is more than Balance DbNt Amount. ", vbOKOnly, Me.Caption
                msGrid.SetFocus
                Exit Function
           End If
           TotAmt = TotAmt + Val(.TextMatrix(i, 4)) + Val(.TextMatrix(i, 6)) + Val(.TextMatrix(i, 9))
        Next i
        If (TotAmt > 0) And (TotAmt <> Val(txt(6))) Then
           MsgBox "Total of Adjusted Amount " & TotAmt & " & Received Amount not tally.", vbOKOnly, Me.Caption
     '      msGrid.SetFocus
     '      Exit Function
        End If
    End If
End With
ValidateGrid = True
End Function
Private Function ValidateData() As Boolean
Dim x As String
Dim GName As String
Dim rstAddCheck As Recordset
GProcCheckForNumber Me, Rstbl, MaxNo
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
    x = "Select vno From " & tblName & " Where vno=" & CLng(txt(0)) & " and Vtype='" & frmTypeStr & "' and VYear=" & gCYear
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
Private Sub ReadFields()
GProcShowRecord Me, Rstbl, MaxNo
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
            GName = GProcGetColumnValue("TblMastGroup", "AgCode", 90019, "N", "AgName", "S")
            GProcShowForm frmMastAccount, frmMain.mnuMstfrm(0), Array(frmMastAccount.txtTemp(2)), Array(GName)
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
Select Case Index
    Case 5: '--Cr
        gClsSearch.SearchMultiField "QryMastaccount", "Account,AcName,AcCity,AgName", Array("Credit Account", "AcName", "City", "Group"), Array(3000, 0, 2000, 3000), "(AGCODE=90029 and ACMILLTYPE=3)", Chr(KeyAscii), "AcName", LeftPos, TopPos, True
        KeyAscii = 0
        If gClsSearch.SearchMultiRetCol(0) <> "" And gClsSearch.SearchMultiRetCol(1) = "" Then
            txtTemp(Index) = gClsSearch.SearchMultiRetCol(0)
        Else
            If gClsSearch.SearchMultiRetCol(1) <> "" Then txtTemp(Index) = gClsSearch.SearchMultiRetCol(1)
        End If
    Case 4: '--dr
        gClsSearch.SearchMultiField "QryMastaccount", "Account,AcName,AcCity,AgName", Array("Debit Account", "AcName", "City", "Group"), Array(3000, 0, 2000, 3000), "(AGCODE=90015 or AGCODE=90017)", Chr(KeyAscii), "AcName", LeftPos, TopPos, True
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
   mLastEntryBank = txtTemp(5).Text
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
If Trim(txttemp1(8).Text) <> "" Then
    List1.Text = txttemp1(8).Text
Else
    List1.ListIndex = 0
End If
End Sub
Private Sub list1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    txttemp1(8).Text = Trim(List1.Text)
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
If txt(17) = 1 Then
    MsgBox "Record is Audited.You can not Modify/Delete this Record.", vbCritical + vbOKOnly, App.Title
    IsRelatedRecord = True
    Exit Function
End If
End Function
Private Sub ShowAccountBalance()
Dim DrAcBal  As Double
Dim CrAcBal As Double
If IsDate(mskDt(2)) Then
   ' DrAcBal = GProcGetAccountBalance(txttemp(4), mskDt(2))
   ' CrAcBal = GProcGetAccountBalance(txttemp(5), mskDt(2))
    txtDrBalance = IIf(DrAcBal > 0, DrAcBal & " Cr", Abs(DrAcBal) & " Dr")
    txtCrBalance = IIf(CrAcBal > 0, CrAcBal & " Cr", Abs(CrAcBal) & " Dr")
End If
End Sub
'---------------- Grid
Private Sub SetGrid()
Dim x As String, x1Rec As Double
Set rsRecVsSale = New Recordset
x = "select InvTp, InvBillNo, InvDt, BillAmt, AdjAmt, DbNtAmt, AdjDbNtAmt, LateDays, Interest, IntReceived ,InvYear,InvNo,JVNOINT" _
& " from tblRecVsSale where " _
& " Recvtype='" & frmTypeStr & "' and Recvno=" & txt(0) & " and RecVYear=" & gCYear & " order by RecVCtrlNo"
x1Rec = GProcRstOpen(rsRecVsSale, x, "R")
With msGrid
    .Clear
    .FormatString = "<Type      |>Inv. No        |<Inv. Date          |>Invoice Amt        |>Adj. Amount       |>DbNt  Amt      |>Adj DbNt Amt    |>Late Days  |>Late Pay Int|>Int Received|||"
    .ColWidth(8) = 0
    .ColWidth(9) = 0
    .ColWidth(10) = 0
    .ColWidth(11) = 0
    .ColWidth(12) = 0
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
'On Error Resume Next
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
    If msGrid.Text <> "" Then txtGrid.Text = msGrid.Text
    '--- Lock TextBox
    With msGrid
        Select Case .Col
            Case 0, 1, 2, 3, 5, 7, 8, 10, 11, 12: '--Type,Inv No, Inv Dt, Debit Note
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
            txtGrid.Text = ""
        End If
        txtGrid.Visible = False
        'txt(18).SetFocus
        Exit Sub
    End If
    
    '---- Navigate Row ,Col
    If KeyAscii = 13 Then
        KeyAscii = 0
        If .Col = rsRecVsSale.Fields.Count - 6 Then
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
        Pending_Invoice
    End If
End With

'---- Set TxtGrid
If Trim(msGrid.Text) <> "" Then
    txtGrid = Trim(msGrid.Text)
End If
If msGrid.Col < (rsRecVsSale.Fields.Count - 5) Then
    If msGrid.Col = 8 Then msGrid.Col = 13
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
                txtGrid.Text = txtGrid.Text & Chr(KeyAscii)
            End If
        Case Else
            txtGrid.Text = Chr(GProcValidateKey(rsRecVsSale, msGrid.Col, KeyAscii, txtGrid.Text)) 'Chr(GProcNumberWithDecimal(KeyAscii, txtGrid))
    End Select
End If
DoEvents
If msGrid.Text <> "" Then
    txtGrid.SelStart = Len(txtGrid.Text) + 1
Else
    txtGrid.SelStart = 1
End If
DoEvents
If txtGrid.Visible = True Then txtGrid.SetFocus
End Sub
Private Sub msGrid_LeaveCell()
If txtGrid.Visible Then
    msGrid.Text = txtGrid.Text
    DoEvents
    txtGrid.Text = ""
    txtGrid.Visible = False
End If
'----Late Payment Interest Calcualtion
'---Interest Rate 1.5
Select Case msGrid.Col
Case 4:
    With msGrid
    '-- int
      '  .TextMatrix(.Row, 8) = Format(Round((((Val(.TextMatrix(.Row, 4)) * 1.5) / 3000) * Val(.TextMatrix(.Row, 7))), 2), "###0.00")
    End With
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
                Pending_Invoice
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
        Case 6:
            If Val(txtGrid) > Val(.TextMatrix(.Row, 5)) Then
                MsgBox "Adjustment Dr Note Amount is less than Dr Note Amount", vbCritical + vbOKOnly, "Adjustment Dr Note Amount is less than Dr Note Amount"
                DoEvents
                txtGrid.SetFocus
                Exit Sub
            End If
        End Select
    End If
    
    '--- Navigate Row ,Col
    If KeyAscii = 13 And .Col < rsRecVsSale.Fields.Count - 6 Then
        .Col = .Col + 1
    ElseIf KeyAscii = 13 And .Col = rsRecVsSale.Fields.Count - 6 Then
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
            If (Val(.TextMatrix(.Row, 4)) = 0) And .Rows > 2 Then
                .RemoveItem (.Row)
                txtGrid.Text = ""
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
If KeyCode = vbKeyY And Shift = 2 And msGrid.Rows > 2 Then
    msGrid = 4
    msGrid.RemoveItem (msGrid.Row)
    txtGrid.Visible = False
    txtGrid.Text = ""
    'CalAmount
    Exit Sub
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
'--- JV
'gCn.Execute "delete from tblVoucher where Vtype='JV' and VNo in (select JVNOINT from tblRecVsSale where " _
'& " INTRECEIVED>0 and RecVtype='" & frmTypeStr & "' and RecVNo=" & txt(0) & " and  RecVyear= " & gCYear & ")  and " _
' & " Vyear= " & gCYear
'--- Rec Vs Sale
gCn.Execute "delete from tblRecVsSale where RecVtype='" & frmTypeStr & "' and RecVNo=" & txt(0) & " and " _
 & " RecVyear= " & gCYear
End Sub

Private Sub SaveRelatedRecord()
Dim i As Long, x As String
Dim IntRecAmt As Double
Dim JVVNo As Long
Dim rsRecVsSale As Recordset, rsOutst As Recordset
'-- Delete Records
DeleteRelatedRecord
'--- Save Grid Records
'--- tblRecSale
x = "select * from tblRecVsSale where Recvtype='" & frmTypeStr & "' and Recvno=" & txt(0) & " and " _
& " RecVYear=" & gCYear
GProcRstOpen rsRecVsSale, x, "O"
With msGrid
    i = 1
    Do While i < .Rows
        DoEvents
        If CDec(Val(.TextMatrix(i, 4))) + CDec(Val(.TextMatrix(i, 6))) + CDec(Val(.TextMatrix(i, 9))) > 0 Then
        rsRecVsSale.AddNew
        rsRecVsSale!AdjNo = GProcGenerateId("tblRecVsSale", "ADJNO", "")
        rsRecVsSale!RecVno = CLng(txt(0)) '--RecVNo
        rsRecVsSale!RecVType = UCase(txt(3)) '--RecVtype
        rsRecVsSale!RecVdt = CDate(txt(2)) '--Vdt(2)
        rsRecVsSale!RecVCtrlNo = 1 'i '--Sr No(3)
        rsRecVsSale!RecVYear = gCYear
        rsRecVsSale!RecAmt = CDec(txt(6))
        'rsRecVsSale!Nar1 = Left(Trim(txttemp1(8)), 3) + ". No. " + Left(txt(19) + Space(10), 10) + " " + Left(txttemp1(7) + Space(20), 20)
        rsRecVsSale!InvBillno = .TextMatrix(i, 1) & vbNullString
        rsRecVsSale!RecVCtrlNo = 1 'i
        rsRecVsSale!InvNo = CLng(.TextMatrix(i, 11)) '--Invno
        rsRecVsSale!InvTp = .TextMatrix(i, 0) '--InvTp
        rsRecVsSale!InvDt = CDate(.TextMatrix(i, 2)) '--Vdt(2)
        rsRecVsSale!InvYear = .TextMatrix(i, 10)
        
        rsRecVsSale!BillAmt = CDec(.TextMatrix(i, 3))
        rsRecVsSale!PartyCode = GProcGetColumnValue("tblMastAccount", "AcName", txtTemp(4), "S", "AcCode", "N")
        rsRecVsSale!AdjAmt = CDec(.TextMatrix(i, 4))
        rsRecVsSale!DbNtAmt = CDec(.TextMatrix(i, 5))
        rsRecVsSale!AdjDbNtAmt = CDec(Val(.TextMatrix(i, 6)))
        
        rsRecVsSale!LateDays = CInt(.TextMatrix(i, 7))
        rsRecVsSale!Interest = CDec(.TextMatrix(i, 8))
        rsRecVsSale!IntReceived = CDec(Val(.TextMatrix(i, 9)))
        IntRecAmt = CDec(Val(.TextMatrix(i, 9)))
        'rsRecVsSale!JVNoDisc
        JVVNo = 0
        '------------------------- Interest Rec JV
'        If FormAction = vbDataActionAddNew And IntRecAmt > 0 Then
'            JVVNo = GProcGenerateId("tblVoucher", "Vno", " vyear=" & gCYear & " and Vtype='JV'")
'        Else
'            If IntRecAmt > 0 Then
'                If CLng(.TextMatrix(i, 12)) > 0 And GProcGetColumnValue("tblVoucher", "Vno", .TextMatrix(i, 12), "N", "Vno", "N", " Vtype='JV' and Vyear=" & gCYear) = 0 Then
'                    JVVNo = CLng(.TextMatrix(i, 12))
'                Else
'                    JVVNo = GProcGenerateId("tblVoucher", "Vno", " vyear=" & gCYear & " and Vtype='JV'")
'                End If
'            End If
'        End If
'        If IntRecAmt > 0 Then
'            GProcCreateVoucher FormAction, JVVNo, 1, mskDt(2), "JV", gLatePayIntAcCode, txt(4), IntRecAmt, IsRefEntryType:=True
'            GProcCreateVoucher FormAction, JVVNo, 2, mskDt(2), "JV", txt(4), gLatePayIntAcCode, -IntRecAmt, IsRefEntryType:=True
'        End If
        '-------------------------------
        rsRecVsSale!JVNOINT = JVVNo
        rsRecVsSale.Update
        End If
        DoEvents
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
Private Sub Pending_Invoice()
Dim x As String
Dim wFromCd As Long
Dim CrAcCode As Long
Dim LateDayCd As String
Dim vamt As Double
Dim rstOutStanding As Recordset
Dim i As Long, j As Long
Dim NewRow As Boolean
Dim TotCol(0, 1) As Long
If Trim(txtTemp(4)) <> "" Then
    wFromCd = GProcGetColumnValue("tblMastAccount", "AcName", txtTemp(4), "S", "AcCode", "N")
    CrAcCode = GProcGetColumnValue("tblMastAccount", "AcName", txtTemp(5), "S", "AcCode", "N")
    If gBackEndDB = gBackEndAccess Then
        LateDayCd = "cdate('" & (CDate(mskDt(2)) + 3) & "') " & " - cdate(outDueDate)"
        x = "select VType, outbillNo,format(vdt,'dd/mm/yyyy'), OutAmount-OutRecAmt, OutDbNtAmt-OutRecDbAmt" _
        & "," & LateDayCd & ", VYear,Vno,OutAmount,OutDbNtAmt from tblOutstanding where Vtype='SM' and OutMillCode=" & CrAcCode & " and OutAcCode=" & wFromCd & " and " _
        & " ((OutAmount-OutRecAmt) > 0 or (OutDbNtAmt-OutRecDbAmt) > 0) "
    Else
        LateDayCd = "to_date('" & CDate(mskDt(2)) + 3 & "','dd/MM/yyyy') - outDueDate"
        x = "select VType, outbillNo,vdt, OutAmount-OutRecAmt, OutDbNtAmt-OutRecDbAmt" _
        & "," & LateDayCd & ", VYear,Vno,OutAmount,OutDbNtAmt from tblOutstanding where Vtype='SM' and OutMillCode=" & CrAcCode & " and OutAcCode=" & wFromCd & " and " _
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
    gClsSearch.SearchMultiSelField x, "", Array("Type", "Invoice No", "Date", "Balance Amt", "DbNt Balance", "Due Days", "VYear", "INV Vno", "Invoice Amt", "DbNt Amt"), Array(800, 1200, 1200, 1500, 1500, 800, 0, 0, 1200, 1200), "", "", "Vdt,vno desc", 0, 0, , True, TotOfSelCols:=TotCol
    vamt = CDbl(txt(6))
    With msGrid
        If gClsSearch.SearchMultiSelRetCol(0, 0) <> "" Then
                j = .Row
                 txtGrid = gClsSearch.SearchMultiSelRetCol(0, 0)
                For i = 0 To UBound(gClsSearch.SearchMultiSelRetCol)
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
                        '--int
                       ' .TextMatrix(i, 8) = Format(Round((((Val(.TextMatrix(i, 4)) * 1.5) / 3000) * Val(.TextMatrix(i, 7))), 2), "###0.00")
                        vamt = 0
                        Exit For
                    Else
                        .TextMatrix(i, 4) = CDbl(.TextMatrix(i, 3))
                        '-- int
                       ' .TextMatrix(i, 8) = Format(Round((((Val(.TextMatrix(i, 4)) * 1.5) / 3000) * Val(.TextMatrix(i, 7))), 2), "###0.00")
                        vamt = vamt - CDbl(.TextMatrix(i, 3))
                    End If
                Next i
                .Col = 4
        End If
    End With
Else
    MsgBox "Select Party Compulsary", vbCritical + vbOKOnly, "Select Party"
    DoEvents
    txtTemp(4).SetFocus
End If
End Sub
