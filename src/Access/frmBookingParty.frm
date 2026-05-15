VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form frmBookingParty 
   BackColor       =   &H00C2E3FC&
   BorderStyle     =   0  'None
   Caption         =   "Party Booking Entry"
   ClientHeight    =   7125
   ClientLeft      =   0
   ClientTop       =   15
   ClientWidth     =   10875
   DrawMode        =   10  'Mask Pen
   FillColor       =   &H000040C0&
   FillStyle       =   0  'Solid
   Icon            =   "frmBookingParty.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7125
   ScaleWidth      =   10875
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFC0C0&
      Caption         =   "Party Book Printing"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3705
      Left            =   2790
      TabIndex        =   60
      Top             =   1620
      Visible         =   0   'False
      Width           =   5175
      Begin VB.CommandButton cmdPrint 
         Caption         =   "&Whatsapp - Broker"
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
         Left            =   2700
         TabIndex        =   74
         Top             =   2970
         Width           =   2280
      End
      Begin VB.CommandButton cmdPrint 
         Caption         =   "&Whatsapp - Party"
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
         Left            =   225
         TabIndex        =   73
         Top             =   2970
         Width           =   2280
      End
      Begin VB.CheckBox Check1 
         BackColor       =   &H00FFC0C0&
         Caption         =   "Print Ticket Name"
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
         Height          =   375
         Left            =   1440
         TabIndex        =   68
         Top             =   1530
         Width           =   2535
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
         TabIndex        =   65
         Top             =   2250
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
         TabIndex        =   64
         Top             =   2250
         Width           =   1335
      End
      Begin VB.TextBox txtno 
         Appearance      =   0  'Flat
         Height          =   360
         Index           =   1
         Left            =   1950
         MaxLength       =   7
         TabIndex        =   63
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
         TabIndex        =   62
         Top             =   2250
         Width           =   1335
      End
      Begin VB.TextBox txtno 
         Appearance      =   0  'Flat
         Height          =   360
         Index           =   0
         Left            =   1980
         MaxLength       =   7
         TabIndex        =   61
         Top             =   600
         Width           =   1575
      End
      Begin VB.Label Label18 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "To      :"
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
         TabIndex        =   67
         Top             =   1080
         Width           =   660
      End
      Begin VB.Label Label3 
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
         Height          =   240
         Index           =   1
         Left            =   600
         TabIndex        =   66
         Top             =   600
         Width           =   675
      End
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
      Height          =   375
      Index           =   15
      Left            =   0
      TabIndex        =   72
      Text            =   "15"
      Top             =   0
      Visible         =   0   'False
      Width           =   465
   End
   Begin VB.CheckBox ChkExGST 
      BackColor       =   &H00C0E0FF&
      Caption         =   "Is Rate Excluding GST ?"
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
      Height          =   405
      Left            =   5445
      TabIndex        =   12
      Top             =   4905
      Width           =   2820
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
      Height          =   375
      Index           =   14
      Left            =   2010
      TabIndex        =   14
      Text            =   "14"
      Top             =   6435
      Width           =   915
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
      Height          =   375
      Index           =   13
      Left            =   0
      TabIndex        =   70
      Text            =   "13"
      Top             =   0
      Visible         =   0   'False
      Width           =   465
   End
   Begin VB.CheckBox Chk 
      BackColor       =   &H00C0E0FF&
      Caption         =   "Print Export ?"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H004D177D&
      Height          =   375
      Index           =   1
      Left            =   3060
      TabIndex        =   11
      Top             =   4875
      Width           =   1785
   End
   Begin VB.TextBox txt 
      Alignment       =   1  'Right Justify
      Height          =   315
      Index           =   12
      Left            =   11460
      TabIndex        =   69
      Text            =   "12"
      Top             =   3900
      Visible         =   0   'False
      Width           =   525
   End
   Begin VB.CheckBox Chk 
      BackColor       =   &H00C0E0FF&
      Caption         =   "Print Ticket Name ?"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H004D177D&
      Height          =   375
      Index           =   0
      Left            =   330
      TabIndex        =   10
      Top             =   4875
      Width           =   2325
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
      Index           =   10
      ItemData        =   "frmBookingParty.frx":058A
      Left            =   2010
      List            =   "frmBookingParty.frx":0594
      TabIndex        =   13
      Text            =   "Cbo(10)"
      Top             =   6030
      Width           =   2145
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
      Index           =   11
      Left            =   5955
      TabIndex        =   15
      Text            =   "11"
      Top             =   6030
      Width           =   3255
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
      Height          =   375
      Index           =   10
      Left            =   -105
      TabIndex        =   59
      Text            =   "10"
      Top             =   5055
      Visible         =   0   'False
      Width           =   465
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
      Height          =   375
      Index           =   9
      Left            =   -60
      TabIndex        =   58
      Text            =   "9"
      Top             =   5400
      Visible         =   0   'False
      Width           =   465
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
      Index           =   3
      Left            =   1320
      TabIndex        =   6
      Text            =   "txttemp(3)"
      ToolTipText     =   "Press F3 for Creation"
      Top             =   720
      Width           =   3945
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
      Left            =   9690
      Style           =   1  'Graphical
      TabIndex        =   16
      ToolTipText     =   "To Save Record"
      Top             =   4200
      Width           =   915
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
      Left            =   9690
      Style           =   1  'Graphical
      TabIndex        =   17
      Top             =   4680
      Width           =   915
   End
   Begin VB.TextBox txttmpTotBag 
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
      Left            =   2010
      TabIndex        =   55
      Text            =   "txttmpTotBag"
      Top             =   5475
      Width           =   1230
   End
   Begin VB.TextBox txttmpTotWt 
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
      Left            =   4770
      TabIndex        =   54
      Text            =   "txttmpTotWt"
      Top             =   5475
      Width           =   1380
   End
   Begin VB.TextBox txtGrid 
      BackColor       =   &H8000000A&
      BorderStyle     =   0  'None
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
      Left            =   495
      TabIndex        =   18
      Top             =   3390
      Visible         =   0   'False
      Width           =   1035
   End
   Begin MSFlexGridLib.MSFlexGrid msGrid 
      Height          =   2955
      Left            =   195
      TabIndex        =   9
      Top             =   1860
      Width           =   9165
      _ExtentX        =   16166
      _ExtentY        =   5212
      _Version        =   393216
      FixedCols       =   0
      RowHeightMin    =   275
      BackColorFixed  =   -2147483647
      ForeColorFixed  =   -2147483634
      BackColorBkg    =   12632256
      AllowUserResizing=   1
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
      Index           =   5
      ItemData        =   "frmBookingParty.frx":05B1
      Left            =   6840
      List            =   "frmBookingParty.frx":05BB
      TabIndex        =   4
      Text            =   "Cbo(5)"
      Top             =   1275
      Width           =   2370
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
      Index           =   0
      Left            =   6840
      TabIndex        =   1
      Text            =   "0"
      Top             =   720
      Width           =   1200
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
      Left            =   10875
      TabIndex        =   52
      Text            =   "1"
      Top             =   2430
      Width           =   420
   End
   Begin VB.TextBox txt 
      Alignment       =   1  'Right Justify
      Height          =   315
      Index           =   2
      Left            =   11040
      TabIndex        =   41
      Text            =   "2"
      Top             =   1440
      Visible         =   0   'False
      Width           =   525
   End
   Begin VB.TextBox txt 
      Alignment       =   1  'Right Justify
      Height          =   315
      Index           =   5
      Left            =   11040
      TabIndex        =   40
      Text            =   "5"
      Top             =   2640
      Width           =   525
   End
   Begin VB.TextBox txt 
      Alignment       =   1  'Right Justify
      Height          =   315
      Index           =   3
      Left            =   11040
      TabIndex        =   39
      Text            =   "3"
      Top             =   1800
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
      TabIndex        =   36
      Top             =   8055
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
      TabIndex        =   35
      ToolTipText     =   "To Add New Record"
      Top             =   7980
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
      Left            =   9690
      Style           =   1  'Graphical
      TabIndex        =   33
      ToolTipText     =   "Exit"
      Top             =   5265
      Width           =   915
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
      Left            =   9690
      Style           =   1  'Graphical
      TabIndex        =   26
      Top             =   2055
      Width           =   915
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
      Left            =   9690
      Style           =   1  'Graphical
      TabIndex        =   25
      ToolTipText     =   "To Delete Record"
      Top             =   3645
      Width           =   915
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
      Height          =   405
      Index           =   7
      Left            =   9690
      Style           =   1  'Graphical
      TabIndex        =   24
      ToolTipText     =   "To Modify Record"
      Top             =   3165
      Width           =   915
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
      Left            =   9690
      Style           =   1  'Graphical
      TabIndex        =   23
      ToolTipText     =   "To Add New Record"
      Top             =   2700
      Width           =   915
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
      Left            =   9750
      Style           =   1  'Graphical
      TabIndex        =   32
      Top             =   5970
      Visible         =   0   'False
      Width           =   915
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
      Left            =   9690
      Style           =   1  'Graphical
      TabIndex        =   27
      ToolTipText     =   "Find Record"
      Top             =   1590
      Width           =   915
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
      Left            =   10170
      Style           =   1  'Graphical
      TabIndex        =   31
      ToolTipText     =   "Last"
      Top             =   1110
      Width           =   435
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
      Left            =   9690
      Style           =   1  'Graphical
      TabIndex        =   30
      ToolTipText     =   "First"
      Top             =   1110
      Width           =   435
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
      Left            =   10170
      Style           =   1  'Graphical
      TabIndex        =   29
      ToolTipText     =   "Next"
      Top             =   630
      Width           =   435
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
      Left            =   9690
      Style           =   1  'Graphical
      TabIndex        =   28
      ToolTipText     =   "Previous"
      Top             =   630
      Width           =   435
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
      Left            =   1320
      TabIndex        =   8
      Text            =   "txttemp(4)"
      ToolTipText     =   "Press F3 for Creation"
      Top             =   1275
      Width           =   3945
   End
   Begin VB.TextBox txt 
      Alignment       =   1  'Right Justify
      Height          =   315
      Index           =   4
      Left            =   11040
      TabIndex        =   38
      Text            =   "4"
      Top             =   2160
      Width           =   525
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
      Height          =   375
      Index           =   6
      Left            =   7620
      TabIndex        =   20
      Text            =   "6"
      Top             =   5475
      Width           =   1590
   End
   Begin VB.TextBox txt 
      Alignment       =   1  'Right Justify
      Height          =   315
      Index           =   7
      Left            =   11040
      TabIndex        =   37
      Text            =   "7"
      Top             =   3360
      Visible         =   0   'False
      Width           =   525
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
      Height          =   375
      Index           =   8
      Left            =   3405
      TabIndex        =   34
      Text            =   "8"
      Top             =   5385
      Visible         =   0   'False
      Width           =   495
   End
   Begin MSMask.MaskEdBox mskDt 
      Height          =   375
      Index           =   2
      Left            =   8070
      TabIndex        =   2
      Top             =   720
      Width           =   1140
      _ExtentX        =   2011
      _ExtentY        =   661
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
   Begin Crystal.CrystalReport cryReport 
      Bindings        =   "frmBookingParty.frx":05D8
      Left            =   9405
      Top             =   6075
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      WindowControlBox=   -1  'True
      WindowMaxButton =   -1  'True
      WindowMinButton =   -1  'True
      PrintFileType   =   15
      PrintFileLinesPerPage=   60
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "Side Days   :"
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
      Left            =   450
      TabIndex        =   71
      Top             =   6480
      Width           =   1185
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "De&livery Place  :"
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
      Height          =   435
      Left            =   435
      TabIndex        =   21
      Top             =   6030
      Width           =   1425
   End
   Begin VB.Label Label10 
      BackStyle       =   0  'Transparent
      Caption         =   "Delivery Period :"
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
      Height          =   435
      Left            =   4455
      TabIndex        =   22
      Top             =   6030
      Width           =   1425
   End
   Begin VB.Line Line1 
      BorderColor     =   &H000000C0&
      X1              =   180
      X2              =   9330
      Y1              =   5925
      Y2              =   5925
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "&Broker :"
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
      Index           =   0
      Left            =   420
      TabIndex        =   5
      Top             =   720
      Width           =   1140
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "&Rate Per     :"
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
      Left            =   5550
      TabIndex        =   3
      Top             =   1275
      Width           =   1305
   End
   Begin VB.Label Label31 
      BackStyle       =   0  'Transparent
      Caption         =   "Total Weight :"
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
      Left            =   3420
      TabIndex        =   57
      Top             =   5475
      Width           =   1230
   End
   Begin VB.Label Label23 
      BackStyle       =   0  'Transparent
      Caption         =   "Total Bag :"
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
      Left            =   915
      TabIndex        =   56
      Top             =   5475
      Width           =   1440
   End
   Begin VB.Shape Shape8 
      BorderColor     =   &H00C00000&
      BorderWidth     =   5
      FillColor       =   &H00C00000&
      Height          =   7125
      Left            =   0
      Top             =   -45
      Width           =   10845
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Booking &No :"
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
      Left            =   5550
      TabIndex        =   0
      Top             =   720
      Width           =   1215
   End
   Begin VB.Label Label30 
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
      ForeColor       =   &H00000080&
      Height          =   375
      Left            =   6375
      TabIndex        =   19
      Top             =   5475
      Width           =   1095
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
      Left            =   8340
      TabIndex        =   53
      Top             =   45
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
      TabIndex        =   51
      Top             =   7800
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
      TabIndex        =   50
      Top             =   8160
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
      TabIndex        =   49
      Top             =   8280
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
      TabIndex        =   48
      Top             =   7920
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
      TabIndex        =   47
      Top             =   8160
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
      TabIndex        =   46
      Top             =   8040
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
      TabIndex        =   45
      Top             =   7860
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
      Left            =   11175
      TabIndex        =   44
      Top             =   7590
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
      TabIndex        =   43
      Top             =   8160
      Width           =   1095
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Par&ty   :"
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
      Left            =   420
      TabIndex        =   7
      Top             =   1275
      Width           =   1140
   End
   Begin VB.Label Label6 
      BackColor       =   &H00C00000&
      Caption         =   "    Party Booking"
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
      Left            =   0
      TabIndex        =   42
      Top             =   0
      Width           =   10815
   End
   Begin VB.Shape Shape1 
      FillColor       =   &H00BDDFD1&
      FillStyle       =   0  'Solid
      Height          =   5445
      Left            =   9540
      Shape           =   4  'Rounded Rectangle
      Top             =   465
      Width           =   1200
   End
   Begin VB.Shape Shape4 
      BorderColor     =   &H00000080&
      FillColor       =   &H00EDFDFE&
      FillStyle       =   0  'Solid
      Height          =   1590
      Left            =   210
      Top             =   5355
      Width           =   9135
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H00000080&
      FillColor       =   &H00EDFDFE&
      FillStyle       =   0  'Solid
      Height          =   1305
      Left            =   210
      Top             =   510
      Width           =   5145
   End
   Begin VB.Shape Shape6 
      BorderColor     =   &H00000080&
      FillColor       =   &H00EDFDFE&
      FillStyle       =   0  'Solid
      Height          =   1305
      Left            =   5415
      Top             =   510
      Width           =   3930
   End
End
Attribute VB_Name = "frmBookingParty"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Const tblName As String = "tblBooking"
Const IdField As String = "Vno"
Const OrderField As String = "Vdt,Vno"
Const MaxNo As Long = 15
Dim rstbl As Recordset '-- TblBooking
Dim rstblSub As Recordset '-- TblBookingSub
Dim varBookMark As Variant
Dim FirstTimeFlag As Boolean
Dim OldgFrmTypeStr As String
Dim FormAction As Integer
Dim frmTypeStr As String
Dim IsLockCtrl As Boolean '-- Set True If Related GatePass & Allow IsComplete grid entry only
Dim i As Long

Private Sub Cbo_KeyPress(Index As Integer, KeyAscii As Integer)
If Index = 5 Then KeyAscii = 0 '--Rate Type
End Sub
Private Sub cmdBtn_Click(Index As Integer)
On Error GoTo ErrorRoutine
Dim X As String
Dim RSAc As Recordset
Dim RsTmp As Recordset
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
    '--- insert ITCode,ITName,Ticket in SetUp tmpGenTbl Table
    GProcInsertToSetupDbTbl False, "select AcName,AcCode,'" & gUserName & "' from tblMastAccount", "tmpGenTbl", "Nar1,VAcCode,UserName", "Username='" & gUserName & "'", ""
    gClsSearch.SearchMultiField "TblBooking,tmpGenTbl", "TblBooking.Vno,TblBooking.Vdt,Nar1,BkBillAmt", Array("Vno", "Vdt", "Broker", "Amount"), Array(txt(0).Width, 1500, 3000, 1500), " BkAcBrkCode=VAcCode and userName='" & gUserName & "' and TblBooking.vtype='" & frmTypeStr & "' and TblBooking.VYear=" & gCYear & " and BkCCode='" & gCCode & "'", "", "TblBooking.Vno,TblBooking.Vdt", 0, 0, Cn:=gSetupCn
    If gClsSearch.SearchMultiRetCol(0) <> "" Then
        rstbl.MoveFirst
        rstbl.Find "Vno = " & gClsSearch.SearchMultiRetCol(0)
        gClsSearch.SearchMultiRetCol(0) = ""
        ReadFields
    End If
Case 6 'Add
    FormAction = vbDataActionAddNew
    Call GProcActivateControls(Me, True)
    Call GProcSetButtons(Me, vbDataActionAddNew)
    Call GProcClearForm(Me, rstbl, MaxNo, True)
    If FirstTimeFlag = True Then
        FirstTimeFlag = False
    Else
        txt(0).SetFocus
    End If
    FillCombo
    cbo(10) = "Ichalkaranji"
    txt(0) = GProcGenerateId(tblName, "Vno", "Vtype='" & frmTypeStr & "' and VYear=" & gCYear, gSetupCn)
    mskDt(2) = GProcGenerateVDt(tblName, "Vdt", "Vtype='" & frmTypeStr & "' and VYear=" & gCYear & " and BkCCode='" & gCCode & "'", gSetupCn)
    IsLockCtrl = False
    SetControlEd
    SetGrid
Case 7 'Modify
    FormAction = vbDataActionUpdate
    Hank = ""
    Call GProcChkPass(frmTypeStr, FormAction, txt(0), txt(2), txttemp(3), gUserName, Hank)
    If gChkPassTrueFalse = True Then
        ' If IsRelatedRecord = False Then
             IsLockCtrl = IsRelatedRecord
             FormAction = vbDataActionUpdate
             Call GProcActivateControls(Me, True)
             SetControlEd
             Call GProcSetButtons(Me, vbDataActionUpdate)
             txt(0).Enabled = False
             mskDt(2).SetFocus
             If IsLockCtrl = True Then '--Allow  Booking Complete
                 Call GProcActivateControls(Me, False)
                 msGrid.Enabled = True
             End If
        ' End If
    Else
       FormAction = vbDataActionClose
    End If
Case 8 'Delete
    FormAction = vbDataActionDelete
    Hank = ""
'    Call GProcChkPass(frmTypeStr, FormAction, txt(0), txt(2), txttemp(3), gUserName, Hank)
'    If gChkPassTrueFalse = True Then
          If IsRelatedRecord = False Then
            If MsgBox("Do You Really Delete This Record?", vbQuestion + vbYesNo + vbDefaultButton1, "Confirm Delete") = vbYes Then
                If rstbl.RecordCount > 1 Then
                     '-- Delete Related Record
                    DeleteRelatedRecord
                    rstbl.Delete
                    cmdBtn_Click (1)
                ElseIf rstbl.RecordCount = 1 Then
                    '-- Delete Related Record
                    DeleteRelatedRecord
                    rstbl.Delete
                    Call GProcActivateControls(Me, False)
                    Call GProcSetButtons(Me, vbDataActionCancel)
                    Form_Load
                End If
            End If
        End If
 '   End If
Case 9 'Print
    Frame1.Visible = True
    txtno(0).Enabled = True
    txtno(1).Enabled = True
    txtno(0) = txt(0)
    txtno(1) = txt(0)
    txtno(0).SetFocus
    Check1.Enabled = True
    

'    GProcCrystalRptPreparation frmMain.cryReport, "Window"
'    With frmMain.cryReport
'       .WindowTitle = "CST Booking"
'        X = "{tmpgenTbl.Username}='" & gUserName & "' and  {tblBooking.Vno}=" & txt(0) & " and {tblBooking.VYear}=" & gCYear & " and {tblBooking.BkCCode}='" & gCCode & "' and {tblBooking.Vtype}='" & txt(1) & "'"
'       .SelectionFormula = X
'       X = "Select ACADD1,ACADD2,ACCITY,ACCST  from tblMastAccount where acname='" & txttemp(4) & "'"
'       .Formulas(5) = "WCSTNo = '" & gCCST & "'" '-gCCST
'       .Formulas(6) = "WVATNo = '" & gCBST & "'" '-gCBST
'       .Formulas(7) = "BrkName = '" & txttemp(3) & "'"  '- Broker
'       .DataFiles(0) = App.Path & "\HIData\setup.mdb"
'       gReportPath = App.Path & "\HiReports\"
'       If GProcRstOpen(RSAc, X, "R") > 0 Then
'           .Formulas(8) = "PartyName = '" & txttemp(4) & "'" '- Party
'           .Formulas(9) = "PartyAdd1= '" & RSAc!ACADD1 & "'" '-Add1
'           .Formulas(10) = "PartyAdd2= '" & RSAc!ACADD2 & "'" '-Add2
'           .Formulas(11) = "PartyCity = '" & RSAc!ACCITY & "'" '-City
'           .Formulas(12) = "PartyCSTNo = '" & RSAc!ACCST & "'" '-CST
'       End If
'       .ReportFileName = gReportPath & "rptBookingPartyPrint.rpt"
'       .Action = 1
'       Exit Sub
'    End With
Case 10 'Exit
    Unload Me
    Exit Sub
Case 11 'Label
    'frmLabel.Show 1
Case 12 'Optional
Case 13 'Save
    If ValidateData = True Then
        If Trim(txtGrid) <> "" Then
            msGrid.text = txtGrid
            DoEvents
            txtGrid = ""
        End If
        txtGrid.Visible = False
       ' If FormAction = vbDataActionAddNew Then txt(0) = GProcGenerateId(tblName, IdField, "vtype='" & frmTypeStr & "' and VYear=" & gCYear, gSetupCn)
        CalAmount
        FillTxtFromTemp
        '-- Save Record
        GProcSaveRecord Me, rstbl, FormAction, MaxNo
        SaveRelatedRecord
        Call GProcActivateControls(Me, False)
        Call GProcSetButtons(Me, True)
        If FormAction = vbDataActionAddNew Then
            If cmdBtn(6).Visible = True Then cmdBtn(6).SetFocus
        Else
            If cmdBtn(7).Visible = True Then cmdBtn(7).SetFocus
        End If
        FormAction = vbDataActionClose
''-----------For SMS
        Dim aa As Boolean
        aa = False
        If MsgBox("Do You Want To Send Whatsapp SMS ? ", vbYesNo) = vbYes Then aa = True

        If aa Then

            Dim HttpReq As New WinHttpRequest
            Dim URL As String
            Dim params As String
            Dim params2 As String
            Dim SmsMblNo As String
            Dim SmsString As String
            Dim Pside As String
            Dim RtNar As String
            Dim PtyName As String
            'Dim RSTmp As String
            
            gCn.BeginTrans
            gCn.Execute "Delete from tmpWhatsAppErr where username = '" & gUserName & "'"
            gCn.CommitTrans
      
            PtyName = txttemp(4)

            SmsMblNo = GProcGetColumnValue("TblMastAccount", "AcName", txttemp(4), "S", "AcFaxNo", "S")
            SmsString = "BO-" + Trim(txt(0)) & vbLf
            SmsString = SmsString + "Dt-" + Trim(CStr(Format(mskDt(2), "dd/mm/yyyy"))) & vbLf
            SmsString = SmsString + "Pty-" + Left(txttemp(4), 30) & vbLf
            SmsString = SmsString + "Cnt-" + msGrid.TextMatrix(1, 1) + "-" + msGrid.TextMatrix(1, 2) & vbLf
            SmsString = SmsString + "Bg-" + msGrid.TextMatrix(1, 3) + ",Wt-" + msGrid.TextMatrix(1, 4) & vbLf
            SmsString = SmsString + "Rt-" + msGrid.TextMatrix(1, 5) + "/" + msGrid.TextMatrix(1, 6) + "Kg" & vbLf
            If ChkExGST.Value = 1 Then
                RtNar = "Rate Excluding GST"
            Else
                RtNar = "Rate Including GST"
            End If
            SmsString = SmsString + RtNar & vbLf
            If txt(14).text = 25 Then
                Pside = "Regular"
            ElseIf txt(14).text = 4 Then
                Pside = "Second Day"
            ElseIf txt(14).text = 1 Then
               Pside = "Next Day"
            End If
            SmsString = SmsString + "Payment Condition-" + Pside & vbLf
            SmsString = SmsString + "Delivey Period -" + Left(txt(11), 25) & vbLf
            SmsString = SmsString + "Brk-" + Left(txttemp(3), 25) & vbLf
            SmsString = SmsString + "Firm-" + gCName
'            SmsString = Left(SmsString, 160)
            If SmsMblNo = "" Then
               MsgBox ("Party Mobile No Not Found....")
            Else
               SentWhatsApp "Txt", "", SmsString, SmsMblNo, PtyName
'               URL = "http://alerts.sinfini.com/api/web2sms.php"
'               params = "workingkey= Ad583e28d45e23d78611667b060ed2e40&sender=MCROCO&to=" & SmsMblNo & " &message=" & SmsString & ""
'               HttpReq.Open "POST", URL, False
'               HttpReq.SetRequestHeader "Content-Type", "application/x-www-form-urlencoded"
'               HttpReq.Send params
            End If
            '----For Broker SMS
            DoEvents
            SmsMblNo = GProcGetColumnValue("TblMastAccount", "AcName", txttemp(3), "S", "AcFaxNo", "S")
            PtyName = txttemp(3)
            If SmsMblNo = "" Then
               MsgBox ("Broker Mobile No Not Found....")
            Else
               SentWhatsApp "Txt", "", SmsString, SmsMblNo, PtyName

'               URL = "http://alerts.sinfini.com/api/web2sms.php"
'               params = "workingkey= Ad583e28d45e23d78611667b060ed2e40&sender=MCROCO&to=" & SmsMblNo & " &message=" & SmsString & ""
'               HttpReq.Open "POST", URL, False
'               HttpReq.SetRequestHeader "Content-Type", "application/x-www-form-urlencoded"
'               HttpReq.Send params
            End If
            X = "Select * from tmpWhatsAppErr where username = '" & gUserName & "'"
            If GProcRstOpen(RsTmp, X, "R", gCn) > 0 Then
               GProcWhatsAppError
            End If
           
            MsgBox ("Message send succesfully....")
            DoEvents
        End If
''------End SMS
    End If
Case 14 'Cancel
    txtGrid.Visible = False
    DoEvents
    Call GProcClearForm(Me, rstbl, MaxNo, True)
    If rstbl.RecordCount <= 0 Then
        Call GProcActivateControls(Me, False)
        Call GProcSetButtons(Me, vbDataActionCancel)
    Else
        If FormAction = vbDataActionAddNew Then
            rstbl.CancelUpdate
            rstbl.MoveLast
        Else
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
    FormAction = vbDataActionCancel
    DoEvents
    msGrid.Col = 0
    If msGrid.Rows <= 1 Then msGrid.Rows = msGrid.Rows + 1
    msGrid.Row = 1
    txtGrid.Move msGrid.CellLeft + msGrid.Left, msGrid.CellTop + msGrid.Top, msGrid.CellWidth, msGrid.CellHeight
End Select
Exit Sub
ErrorRoutine:
    rstbl.CancelUpdate
    GProcErrorHandler
End Sub
'-- Set Control Enability
Private Sub SetControlEd()
txt(6).Enabled = False '--Bill Amt
txttmpTotBag.Enabled = False
txttmpTotWt.Enabled = False
End Sub
Private Sub DeleteRelatedRecord()
'--- Booking Sub
gSetupCn.Execute "delete from tblBookingSub where vtype='" & frmTypeStr & "' and vno=" & txt(0) & " and " _
& " VYear=" & gCYear
End Sub
Private Sub SaveRelatedRecord()
Dim i As Long
Dim rsSub As Recordset
Dim X As String
'-- Delete Records
DeleteRelatedRecord
'--- Save Grid Records
'--- TblVouSub
X = "select * from tblBookingSub where vtype='" & frmTypeStr & "' and vno=" & txt(0) & " and " _
& " VYear=" & gCYear
GProcRstOpen rsSub, X, "O", gSetupCn
With msGrid
    i = 1
    Do While i < .Rows
        rsSub.AddNew
        rsSub!Vno = CLng(txt(0)) '--Vno
        rsSub!Vtype = txt(1) '--Vtype
        rsSub!Vdt = CDate(txt(2)) '--Vdt
        rsSub!BkSubCtrNo = i '--Sr No
        '--Itcode
        rsSub!BksubITCODE = CLng(.TextMatrix(i, 8))
        rsSub!BksubBAG = CLng(.TextMatrix(i, 3)) '--Bag
        rsSub!Bksubwt = CDbl(.TextMatrix(i, 4)) '-- Wt
        rsSub!BksubRt = CDbl(.TextMatrix(i, 5)) '--Rate
        rsSub!BksubRTPER = CDbl(.TextMatrix(i, 6)) '--Rate Per
        rsSub!BksubAMT = CDbl(.TextMatrix(i, 7)) '--Amt
        rsSub!BkSubItCtrlNo = GProcGenerateId("tblBookingSub", "BkSubItCtrlNo", "", gSetupCn) '--It Ctrl No
        rsSub!BkSubIsCompleted = IIf(Trim(UCase(.TextMatrix(i, 9))) = "Y", 1, 0) '-Is complete
        rsSub!VYear = txt(7) '--VYear
        rsSub.Update
        i = i + 1
     Loop
 End With
 rsSub.Close
End Sub

Private Sub cmdPrint_Click(Index As Integer)
Dim XX As String
Dim YY As String
Dim RsGp As Recordset
Dim RSAc As Recordset
Dim GpNo As Long
Dim X As String
Dim InvNo As Variant
Dim RsSA As Recordset
Dim RsTmp As Recordset
Dim LrString As String
Dim SmsMblNo As String
Dim SmsString As String
Dim Bnm As String
Dim Pnm As String
Dim PAd1 As String
Dim PAd2 As String
Dim PCity As String
Dim PGst As String

Select Case Index
    Case 0, 1, 3, 4:
'----sss
    If IsNumeric(txtno(0)) = False Then txtno(0) = 0
    If IsNumeric(txtno(1)) = False Then txtno(1) = txtno(0)
    If CLng(txtno(1)) < CLng(txtno(0)) Then txtno(1) = txtno(0)
    '-------
    If Index = 0 Then GProcCrystalRptPreparation CryReport, "Printer"
    If Index = 1 Then GProcCrystalRptPreparation CryReport, "Window"
    If Index = 3 Then GProcCrystalRptPreparation CryReport, "Window"
    If Index = 4 Then GProcCrystalRptPreparation CryReport, "Window"
    
    CryReport.WindowTitle = "Party Booking"
    GpNo = CLng(txtno(0))
    Do While GpNo <= txtno(1)
        XX = "Select * From tblBooking where  vNo=" & GpNo & " and tblBooking.VYear=" & gCYear & " and tblBooking.BkCCode='" & gCCode & "' and tblBooking.Vtype='" & txt(1) & "'"
        If GProcRstOpen(RsGp, XX, "R", gSetupCn) > 0 Then
            With CryReport
                .SelectionFormula = "{tblBooking.vNo} = " & RsGp!Vno & "  and {tmpGentbl.USERNAME} = '" & gUserName & "' and {tblBooking.VType}='" & RsGp!Vtype & "' and {tblBooking.VYear}=" & RsGp!VYear
                .DataFiles(0) = App.Path & "\HIData\setup.mdb"
                .Formulas(5) = "WGSTNo = '" & gCGSTIN & "'"  '-gCCST
                .Formulas(6) = "WVATNo = '" & gCBST & "'" '-gCBST
                 YY = "Select AcName,ACADD1,ACADD2,ACCITY,AcGSTIN  from tblMastAccount where AcCode =" & RsGp.Fields(3) & ""
                    If Check1.Value = 1 Then
                        .Formulas(15) = "wMnm = 'Y'"
                    Else
                       .Formulas(15) = "wMnm = 'N'"
                    End If
                 
                 If GProcRstOpen(RSAc, YY, "R") > 0 Then
                    .Formulas(7) = "BrkName = '" & RSAc!AcName & "'"   '- Broker
                 End If
'--Party
                 YY = "Select AcName,ACADD1,ACADD2,ACCITY,AcGSTIN  from tblMastAccount where AcCode =" & RsGp.Fields(4) & ""
                 If GProcRstOpen(RSAc, YY, "R") > 0 Then
                    .Formulas(9) = "PartyAdd1= '" & RSAc!ACADD1 & "'" '-Add1
                    .Formulas(10) = "PartyAdd2= '" & RSAc!ACADD2 & "'" '-Add2
                    .Formulas(11) = "PartyCity = '" & RSAc!ACCITY & "'" '-City
                    .Formulas(12) = "PartyGSTNo = '" & RSAc!AcGSTIN & "'" '-CST
                    .Formulas(8) = "PartyName = '" & RSAc!AcName & "'" '- Party
                     Pnm = RSAc!AcName
                     PAd1 = RSAc!ACADD1 & vbNullString
                     PAd2 = RSAc!ACADD2 & vbNullString
                     PCity = RSAc!ACCITY & vbNullString
                     PGst = RSAc!AcGSTIN
                 
                End If
                gReportPath = App.Path & "\HiReports\"
                 '--Report File At  D:\HiReports if exist
                    .ReportFileName = gReportPath & "rptBookingPartyPrint.rpt"
                     If Index <> 3 And Index <> 4 Then
                       .Action = 1
                     End If

'-------------------Mail
                 If (Index = 3 Or Index = 4) Then
                    Dim objCrystal As CRAXDRT.Application
                    Dim objReport As CRAXDRT.Report
                    Dim Tattach As String
                    Dim Mfile As String
                    Dim Mcp As String
                    Dim PtyName As String
                    Dim PtyMailId As String
                    Dim PtyMailIdCc As String
                    Dim PtyMailIdBcc As String
                    Dim MailSub As String
                    Dim MailStr As String
                    Set objCrystal = New CRAXDRT.Application
                    Tattach = gReportPath & "rptBookingPartyPrint.rpt"
                    Set objReport = objCrystal.OpenReport(Tattach, 1)
                    
                    Dim crTab As CRAXDRT.DatabaseTable
                    For Each crTab In objReport.Database.Tables
'                        crTab.location = gDbLocation
                        crTab.Location = gDataPath & "setup.mdb"   '   gDbLocation
                    
                    Next
                    
                    
                    objReport.RecordSelectionFormula = "{tblBooking.vNo} = " & RsGp!Vno & "  and {tmpGentbl.USERNAME} = '" & gUserName & "' and {tblBooking.VType}='" & RsGp!Vtype & "' and {tblBooking.VYear}=" & RsGp!VYear
                    
                    Dim j As Integer
                    For j = 1 To objReport.FormulaFields.Count
                        Select Case objReport.FormulaFields(j).Name
                        Case "{@FirmNm}"
                            objReport.FormulaFields(j).text = "Trim(" & Chr(39) & UCase(gCName) & Chr(39) & ")"
                        Case "{@wAdd1}"
                            objReport.FormulaFields(j).text = "Trim(" & Chr(39) & gCAdd1 & Chr(39) & ")"
                        Case "{@wAdd2}"
                            objReport.FormulaFields(j).text = "Trim(" & Chr(39) & gCAdd2 & Chr(39) & ")"
                        Case "{@wAdd3}"
                            objReport.FormulaFields(j).text = "Trim(" & Chr(39) & gCAdd3 & Chr(39) & ")"
                        Case "{@wPhNo}"
                            objReport.FormulaFields(j).text = "Trim(" & Chr(39) & gCPhNo & Chr(39) & ")"
                        
                        Case "{@PartyAdd1}"
                            objReport.FormulaFields(j).text = "Trim(" & Chr(39) & PAd1 & Chr(39) & ")"
                        Case "{@PartyAdd2}"
                            objReport.FormulaFields(j).text = "Trim(" & Chr(39) & PAd2 & Chr(39) & ")"
                        Case "{@PartyCity}"
                            objReport.FormulaFields(j).text = "Trim(" & Chr(39) & PCity & Chr(39) & ")"
                        Case "{@PartyGSTNo}"
                            objReport.FormulaFields(j).text = "Trim(" & Chr(39) & UCase(PGst) & Chr(39) & ")"
                        Case "{@PartyName}"
                            objReport.FormulaFields(j).text = "Trim(" & Chr(39) & UCase(Pnm) & Chr(39) & ")"
                        Case "{@wCompGstin}"
                            objReport.FormulaFields(j).text = "Trim(" & Chr(39) & UCase(gCGSTIN) & Chr(39) & ")"
                        Case "{@wInsuNo}"
                            objReport.FormulaFields(j).text = "Trim(" & Chr(39) & UCase(gCFileNo) & Chr(39) & ")"
                        Case "{@Wcopy}"
                            If Index = 3 And Index = 4 Then
                               objReport.FormulaFields(j).text = "Trim(" & Chr(39) & "WhatsApp Copy" & Chr(39) & ")"
                            Else
                               objReport.FormulaFields(j).text = "Trim(" & Chr(39) & "Mail Copy" & Chr(39) & ")"
                            End If
                        Case "{@BrkName}"
                            objReport.FormulaFields(j).text = "Trim(" & Chr(39) & UCase(Bnm) & Chr(39) & ")"
                        End Select
                   
                    Next j
                    Dim PtyMbl As String
                    Dim BrkName As String
                    Dim BrkMbl As String
                    Dim BrkCd As Long
                      
                    Dim FileNMWithPath As String
                    Dim FileNM As String
                    Dim aFile As String
                    'FileNM = "" & Format(Now, "ddMMyyyyhhmmss") & "" & ".pdf"
                    FileNM = "Outstanding " & Format(Now, "ddMMyyyyhhmmss") & "" & ".pdf"
                    
                    FileNMWithPath = gReportPath & "" & FileNM & ""
                    PtyMbl = GProcGetColumnValue("TblMastAccount", "AcCode", RsGp!BKAcCode, "N", "AcFaxNo", "S")
                    PtyName = GProcGetColumnValue("TblMastAccount", "AcCode", RsGp!BKAcCode, "N", "AcName", "S")
                    If Index = 3 Then
                       ExportReportToPDF objReport, FileNMWithPath, "foo"
                       SentWhatsApp "PDF", FileNM, "", PtyMbl, PtyName
                    End If

                    If Index = 4 Then
                      '-- To Broker
                          BrkCd = RsGp!BkAcBrkCode
                          BrkName = GProcGetColumnValue("TblMastAccount", "AcCode", "" & BrkCd & "", "N", "AcName", "S")
                          BrkMbl = GProcGetColumnValue("TblMastAccount", "AcCode", "" & BrkCd & "", "N", "AcFaxNo", "S")
                          ExportReportToPDF objReport, FileNMWithPath, "foo"
                          SentWhatsApp "PDF", FileNM, "", BrkMbl, BrkName
                     End If

                     gReportPath = App.Path & "\HIReports_Access\"
                     aFile = gReportPath & "*.pdf"
                     If Len(Dir$(aFile)) > 0 Then
                        Kill aFile
                     End If
                 End If
            End With
        End If
        GpNo = GpNo + 1
    Loop
    X = "Select * from tmpWhatsAppErr where username = '" & gUserName & "'"
    If GProcRstOpen(RsTmp, X, "R", gCn) > 0 Then
       GProcWhatsAppError
    End If

Case 2:
    DoEvents
    Frame1.Visible = False
End Select

End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case 13:
            If Me.ActiveControl.Name <> "txtGrid" Then Sendkeys "{TAB}"
            'SendKeys "{TAB}"
'        Case 27:
'            If (cmdBtn(13).Enabled = False) Then
'                Unload Me
'            End If
    End Select
End Sub
Private Sub Form_Load()
    Dim X As String
    Me.Top = 300
    Me.Left = Screen.Width / 2 - Me.Width / 2
    '--- TblBooking
    frmTypeStr = gfrmTypeStr
    X = "Select * from " & tblName & " where Vtype='" & frmTypeStr & "' and VYear=" & gCYear & " and BkCCode='" & gCCode & "'"
    If Trim(OrderField) <> "" Then X = X & " order by " & OrderField
    GProcRstOpen rstbl, X, "O", gSetupCn
    FirstTimeFlag = True
    '---Grid
    SetGrid
    SetRelatedRS
    '---- Button Visiblity Set As Per User Previlage
    GProcSetButtonVisiblity Me
    cmdBtn_Click (6)
End Sub
Private Sub SetRelatedRS()
End Sub
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
If FormAction = vbDataActionUpdate Then
    MsgBox "You are about to modify the record. Save the record and then exit.", vbInformation + vbOKOnly, Me.Caption
    Cancel = True
    Exit Sub
End If
rstbl.Requery
rstbl.Close
End Sub
Private Sub mskDt_GotFocus(Index As Integer)
GProcSelectBox Me.ActiveControl
End Sub
Private Sub mskDt_LostFocus(Index As Integer)
If Index = 2 Then
    If GProcIsDateValid(mskDt(Index)) = False Then
        mskDt(Index).SetFocus
        Exit Sub
    End If
End If
SaveBtnEd
End Sub
Private Sub txt_GotFocus(Index As Integer)
GProcSelectBox Me.ActiveControl
End Sub
Private Sub txt_LostFocus(Index As Integer)
GProcCheckForNumber Me, rstbl, MaxNo
CalAmount
SaveBtnEd
End Sub
Private Sub txt_KeyPress(Index As Integer, KeyAscii As Integer)
KeyAscii = GProcValidateKey(rstbl, Index, KeyAscii, txt(Index).text)
End Sub
Private Sub txttemp_GotFocus(Index As Integer)
GProcSelectBox Me.ActiveControl
End Sub
Private Sub txttemp_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
Dim GName As String
If KeyCode = vbKeyF3 Then
    Select Case Index
        Case 3 '-- Broker code
            GName = GProcGetColumnValue("TblMastGroup", "AgCode", 90016, "N", "AgName", "S")
            GProcShowForm frmMastAccount, frmMain.mnuMstfrm(0), Array(frmMastAccount.txttemp(2)), Array(GName)
        Case 4 '-- Party code
            GName = GProcGetColumnValue("TblMastGroup", "AgCode", 90017, "N", "AgName", "S")
            GProcShowForm frmMastAccount, frmMain.mnuMstfrm(0), Array(frmMastAccount.txttemp(2)), Array(GName)
    End Select
End If
End Sub
Private Sub txtTemp_KeyPress(Index As Integer, KeyAscii As Integer)
Dim LeftPos As Long
Dim TopPos As Long
TopPos = txttemp(Index).Top + Me.Top + 650
LeftPos = txttemp(Index).Left + Me.Left
If KeyAscii = 13 Then Exit Sub
Select Case Index
    Case 3  '-- Broker Alias Account
        gClsSearch.SearchMultiField "QryMastaccount", "Account,AcName,AcCity,AgName", Array("Broker", "AcName", "City", "Group"), Array(3000, 0, 2000, 3000), "(AgCode=90016 or GpCode=90016)", Chr(KeyAscii), "AcName", LeftPos, TopPos, True
        KeyAscii = 0
        If gClsSearch.SearchMultiRetCol(0) <> "" And gClsSearch.SearchMultiRetCol(1) = "" Then
            txttemp(Index) = gClsSearch.SearchMultiRetCol(0)
        Else
            If gClsSearch.SearchMultiRetCol(1) <> "" Then txttemp(Index) = gClsSearch.SearchMultiRetCol(1)
        End If
    Case 4  '-- Party  Alias Account
        gClsSearch.SearchMultiField "QryMastaccount", "Account,AcName,AcCity,AgName", Array("Party", "AcName", "City", "Group"), Array(3000, 0, 2000, 3000), "(AgCode=90017 or AgCode=90015)", Chr(KeyAscii), "AcName", LeftPos, TopPos, True
        KeyAscii = 0
        If gClsSearch.SearchMultiRetCol(0) <> "" And gClsSearch.SearchMultiRetCol(1) = "" Then
            txttemp(Index) = gClsSearch.SearchMultiRetCol(0)
        Else
            If gClsSearch.SearchMultiRetCol(1) <> "" Then txttemp(Index) = gClsSearch.SearchMultiRetCol(1)
        End If
End Select
End Sub
Private Sub txtTemp_LostFocus(Index As Integer)
SaveBtnEd
End Sub
'---------------- Grid
Private Sub SetGrid()
Dim X As String
Dim i As Long
Set rstblSub = New Recordset
'--- insert ITCode,ITName,Ticket in SetUp tmpGenTbl Table
GProcInsertToSetupDbTbl False, "select itname,ITTICKET,Itcode,AcName,'" & gUserName & "' from tblMastItem,tblMastAccount where ItMillCode=AcCode", "tmpGenTbl", "Nar1,Nar2,VAcCode,Nar5,UserName", "Username='" & gUserName & "'", ""

X = "select BkSubCtrNo,Nar1,Nar2,BkSubBag,BkSubWt" _
& " ,BkSubRt,BkSubRtPer,BkSubAmt,VAcCode,iif(BkSubIsCompleted=1,'Y','') from TblBookingSub,tmpGenTbl where tmpGenTbl.UserName='" & gUserName & "'" _
& " and BkSUBItCODE=VAcCode and TblBookingSub.vtype='" & frmTypeStr & "' and TblBookingSub.vno=" & txt(0) & " and TblBookingSub.VYear=" & gCYear & " order by BkSubCtrNo"
GProcRstOpen rstblSub, X, "R", gSetupCn
With msGrid
    .FormatString = ">SR|<Count                                     |<Ticket     |>Bag   |>Weight      |>Rate              |>Rt Per|>Amount          |<ItCode|<Is Complete(Y)"
    .ColWidth(8) = 0
    If FormAction = vbDataActionAddNew Then
        .Rows = 2
        GridAddNew
    Else
        .Rows = 1
    End If
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
                 .TextMatrix(.Row, i) = Null
            Case Else
                 .TextMatrix(.Row, i) = ""
        End Select
    Next i
ReNumberGridSrno
.Refresh
End With
End Sub
Private Sub msGrid_Click()
On Error Resume Next
Dim i As Long
msGrid_EnterCell
End Sub
Private Sub msGrid_GotFocus()
If LCase(Me.ActiveControl.Name) <> "txtgrid" Then msGrid.Col = 1
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
    If msGrid.text <> "" Then txtGrid.text = msGrid.text
    '--- Lock TextBox
    If IsLockCtrl = False Then
        With msGrid
            Select Case .Col
                Case 0: '---- Sr No
                    txtGrid.Visible = True
                    txtGrid.Locked = True
                    ReNumberGridSrno
                    DoEvents
                    .Col = .Col + 1
                Case 2 '-- Ticket
                    txtGrid.Locked = True
                Case Else
                    txtGrid.Locked = False
            End Select
        End With
    Else
        If msGrid.Col = 9 Then '-Is complete
            txtGrid.Locked = False
        Else
            txtGrid.Locked = True
        End If
    End If
    DoEvents
    msGrid_KeyPress (0)
    DoEvents
End If
End Sub
Private Sub msGrid_KeyPress(KeyAscii As Integer)
With msGrid
    If KeyAscii = 27 Then '----Esc Key
        KeyAscii = 0
        If (Val(.TextMatrix(.Row, 7)) = 0) And .Rows > 2 Then '-- amt =0
            .RemoveItem (.Row)
            txtGrid.text = ""
        End If
        txtGrid.Visible = False
        Exit Sub
    End If
End With

'If FormAction <> vbDataActionAddNew And FormAction <> vbDataActionUpdate Then Exit Sub
'---- Navigate Row ,Col
If KeyAscii = 13 Then
    KeyAscii = 0
    With msGrid
        If .Col = rstblSub.Fields.Count - 1 Then
            If .Row = (.Rows - 1) Then '-- Add New Row
                .Rows = .Rows + 1
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
'---- Set TxtGrid
If msGrid.text <> "" Then txtGrid.text = msGrid.text
If msGrid.Col <= rstblSub.Fields.Count - 1 And msGrid.Col <> 8 Then
    txtGrid.Move msGrid.CellLeft + msGrid.Left, msGrid.CellTop + msGrid.Top, msGrid.CellWidth, msGrid.CellHeight
End If
txtGrid.Visible = True
txtGrid.Enabled = True
DoEvents
If KeyAscii <> 0 Then
    Select Case rstblSub.Fields(msGrid.Col).Type
        Case adChar, adVarWChar
            If msGrid.Col = 1 Then  '-- Count
                txtGrid_KeyPress (KeyAscii)
            Else
                txtGrid.text = txtGrid.text & Chr(KeyAscii)
            End If
        Case Else
            txtGrid.text = Chr(GProcValidateKey(rstblSub, msGrid.Col, KeyAscii, txtGrid.text))
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
DoEvents
If msGrid.Col >= 1 Then
    CalGridAmt
    CalAmount
End If
DoEvents
End Sub
Private Sub txtGrid_GotFocus()
    txtGrid.Alignment = ProcSetAlignment(rstblSub, msGrid.Col)
    DoEvents
    Call GProcSelectBox(txtGrid)
End Sub
Private Sub txtGrid_KeyPress(KeyAscii As Integer)
Dim LeftPos As Long
Dim TopPos As Long
TopPos = txtGrid.Top + Me.Top + 650
LeftPos = txtGrid.Left + Me.Left
DoEvents
With msGrid
    If KeyAscii <> 13 And KeyAscii <> 27 Then
        If .Col <> 9 And IsLockCtrl = True Then
            KeyAscii = 0
        Else
            '--- Search
            Select Case .Col
                Case 1  '-- Count
                    gClsSearch.SearchMultiField "tblMastItem,tblMastaccount", "ItName,AcName,ItStdpack,ItStdrateper,ItCode,ITTICKET", Array("Count Name", "Mill", "Std. Pack", "Rate Per", "It Code", "ITTICKET"), Array(txtGrid.Width, 2000, 1000, 1000, 0, 0), "tblMastItem.ITMillCode=tblMastaccount.acCode", Chr(KeyAscii), "ItName", LeftPos, TopPos
                    txtGrid = ""
                    KeyAscii = 0
                    If gClsSearch.SearchMultiRetCol(0) <> "" Then '-- Rate Per,ItCode From Master
                        txtGrid = gClsSearch.SearchMultiRetCol(0)
                        If .TextMatrix(.Row, 6) = 0 Then .TextMatrix(.Row, 6) = gClsSearch.SearchMultiRetCol(3) '-- Rate Per
                        .TextMatrix(.Row, 8) = gClsSearch.SearchMultiRetCol(4) '--Itcode
                        .TextMatrix(.Row, 2) = gClsSearch.SearchMultiRetCol(5) '--Ticket
                        txtGrid_KeyPress 13
                    End If
                Case 9 '- Is Complete Y/N
                        If UCase(Chr(KeyAscii)) = "Y" Then 'Y
                            txtGrid = ""
                            KeyAscii = Asc("Y")
                        Else
                            txtGrid = ""
                            KeyAscii = 0
                        End If
                Case Else
                    KeyAscii = GProcValidateKey(rstblSub, msGrid.Col, KeyAscii, txtGrid)
            End Select
        End If
   End If
    If KeyAscii = 13 Then CheckForNumberTxtGrid
    '--- Navigate Row ,Col
    If KeyAscii = 13 And .Col < 9 Then
        .Col = .Col + 1
    ElseIf KeyAscii = 13 And .Col = 9 Then
        If .Row = (.Rows - 1) Then
            .Rows = .Rows + 1
            .Row = .Row + 1
            GridAddNew
        Else
            .Row = .Row + 1
        End If
        .Col = 0
    End If
    '-- On ESC
    With msGrid
        If KeyAscii = 27 Then '----Esc Key
            KeyAscii = 0
            If (Val(.TextMatrix(.Row, 7)) = 0) And .Rows > 2 Then '-- Amt=0
                .RemoveItem (.Row)
                txtGrid.text = ""
            End If
            txtGrid.Visible = False
            Exit Sub
        End If
    End With
End With
End Sub
Private Sub txtGrid_KeyDown(KeyCode As Integer, Shift As Integer)
Dim GName As String
'--- Delete Row On Ctrl+Y
If KeyCode = vbKeyY And Shift = 2 And msGrid.Rows > 2 Then
    msGrid = 4
    msGrid.RemoveItem (msGrid.Row)
    txtGrid.Visible = False
    txtGrid.text = ""
    ReNumberGridSrno
    CalAmount
    Exit Sub
End If
'--- Show Form
If KeyCode = vbKeyF3 Then
Select Case msGrid.Col
    Case 1 '-- count
        GProcShowForm frmMastItem, frmMain.mnuMstfrm(3)
        Exit Sub
End Select
End If
End Sub
Private Sub CalGridAmt()
Dim StdPack As Double
Dim Bags As Long
CheckForNumberGrid
If msGrid.Row = 0 Then Exit Sub
StdPack = GProcGetColumnValue("TblMastItem", "ITCode", msGrid.TextMatrix(msGrid.Row, 8), "N", "ItStdpack", "N")
With msGrid
    '-- Wt=Bag * Std Pack
    If .TextMatrix(.Row, 4) = 0 Then
        .TextMatrix(.Row, 4) = .TextMatrix(.Row, 3) * StdPack
    End If
    '--Amt=Rate * Wt/Rate Per
    If CDbl(.TextMatrix(.Row, 6)) > 0 Then
        .TextMatrix(.Row, 7) = CDbl(.TextMatrix(.Row, 5)) * CDbl(.TextMatrix(.Row, 4)) / CDbl(.TextMatrix(.Row, 6))
    Else
        .TextMatrix(.Row, 7) = 0
    End If
    .TextMatrix(.Row, 7) = Format(.TextMatrix(.Row, 7), FStr)
'    '--Rate =Amt*Rate Per/Wt
'    If CDbl(.TextMatrix(.Row, 11)) > 0 Then
'        .TextMatrix(.Row, 12) = CDbl(.TextMatrix(.Row, 14)) * CDbl(.TextMatrix(.Row, 13)) / CDbl(.TextMatrix(.Row, 11))
'    Else
'        .TextMatrix(.Row, 12) = 0
'    End If
'    .TextMatrix(.Row, 12) = Format(.TextMatrix(.Row, 12), LStr)
End With
End Sub
Private Sub CheckForNumberGrid()
Dim RowIndex As Long
With rstblSub
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
                    msGrid.TextMatrix(RowIndex, i) = Null
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
                txtGrid = Null
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
'-- ReNumber Grid Sr No
Private Sub ReNumberGridSrno()
Dim i As Long
With msGrid
    For i = 1 To .Rows - 1
        .TextMatrix(i, 0) = i
    Next i
End With
End Sub
'-- Grid Required
Private Function ValidateGrid() As Boolean
Dim i As Long
Dim ItCode  As Long
Dim RowIndex As Long
If Trim(txtGrid.text) <> "" Then
    msGrid.text = txtGrid.text
    CalGridAmt
End If
txtGrid.Visible = False
txtGrid.text = ""
CheckForNumberGrid
With msGrid
Validate:
    For RowIndex = 1 To .Rows - 1
        For i = 0 To rstblSub.Fields.Count - 1
            Select Case i
                Case 1 '-- Count
                    '-count,wt=0 then delete row
                    If RowIndex > 1 And Trim(.TextMatrix(RowIndex, i)) = "" And CDbl(.TextMatrix(RowIndex, 4)) = 0 Then
                        .RemoveItem (RowIndex)
                        GoTo Validate
                    End If
                    '--Create Count
                    If GProcGetColumnValue("TblMastItem", "ItName", .TextMatrix(RowIndex, i), "S", "Itcode", "N") = 0 Then
                        MsgBox "Check Count.", vbCritical + vbOKOnly, Me.Caption
                        .Col = i
                        msGrid.SetFocus
                        Exit Function
                    End If
                Case 3 '-- Bag
                    If CDbl(.TextMatrix(RowIndex, i)) = 0 Then
                        MsgBox "Check Bag.", vbCritical + vbOKOnly, Me.Caption
                        .Col = i
                        msGrid.SetFocus
                        Exit Function
                    End If
                Case 4 '-- Wt
                    If CDbl(.TextMatrix(RowIndex, i)) = 0 Then
                        MsgBox "Check Weight.", vbCritical + vbOKOnly, Me.Caption
                        .Col = i
                        msGrid.SetFocus
                        Exit Function
                    End If
                Case 5 '-- Rate
                    If CDbl(.TextMatrix(RowIndex, i)) = 0 Then
                        MsgBox "Check Rate.", vbCritical + vbOKOnly, Me.Caption
                        .Col = i
                        msGrid.SetFocus
                        Exit Function
                    End If
           End Select
        Next i
        '--- Duplicate Count
        ItCode = .TextMatrix(RowIndex, 8)
        For i = 1 To .Rows - 1
            If i <> RowIndex And .TextMatrix(i, 8) = ItCode Then
                MsgBox "Duplicate count For Mill.", vbCritical + vbOKOnly, Me.Caption
                .Col = 1
                msGrid.SetFocus
                Exit Function
            End If
        Next i
      Next RowIndex
End With
ValidateGrid = True
ReNumberGridSrno
End Function
Private Sub SaveBtnEd()
'--- Vno,vdt,Broker,Amount
If CDbl(txt(0)) = 0 Or GProcIsDateValid(mskDt(2)) = False Or txttemp(3) = "" Or CDbl(txt(6)) = 0 Then
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
Dim AcCode As Long
Dim rstAddCheck As Recordset
Dim rstModiCheck As Recordset
    '--- Validate Grid
    If ValidateGrid = False Then
        Exit Function
    End If
    '---- Required Data
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
    '---- Broker
    If txttemp(3) = "" Then
        MsgBox "Check Broker.", vbInformation + vbOKOnly, Me.Caption
        txttemp(3).SetFocus
        Exit Function
    End If
    '--Create Broker
    If GProcGetColumnValue("tblMastAccount", "AcName", txttemp(3), "S", "AcCode", "N") = 0 Then
        GName = GProcGetColumnValue("TblMastGroup", "AgCode", 90016, "N", "AgName", "S")
        GProcShowForm frmMastAccount, frmMain.mnuMstfrm(0), Array(frmMastAccount.txt(1), frmMastAccount.txttemp(2)), Array(txttemp(3), GName)
        Exit Function
    End If
    '--Create Party
    If Trim(txttemp(4)) <> "" Then
        If GProcGetColumnValue("tblMastAccount", "AcName", txttemp(4), "S", "AcCode", "N") = 0 Then
            GName = GProcGetColumnValue("TblMastGroup", "AgCode", 90017, "N", "AgName", "S")
            GProcShowForm frmMastAccount, frmMain.mnuMstfrm(0), Array(frmMastAccount.txt(1), frmMastAccount.txttemp(2)), Array(txttemp(4), GName)
            Exit Function
        End If
    End If
    '---- Amount
    If CDbl(txt(6)) <= 0 Then
        MsgBox "Check Booking Amount.", vbInformation + vbOKOnly, Me.Caption
        Exit Function
    End If
    '----Duplicate Records
    '----Addition Case
    If FormAction = vbDataActionAddNew Then
        X = "Select vno From " & tblName & " Where vno=" & CLng(txt(0)) & " and Vtype='" & frmTypeStr & "' and VYear=" & gCYear
        i = GProcRstOpen(rstAddCheck, X, "R", gSetupCn)
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
Call GProcFillCombo(cbo(10), "select Distinct DeliveryPlace from " & tblName & " where DeliveryPlace <> '' order by DeliveryPlace", gSetupCn)
End Sub
Private Sub FillTxtFromTemp()
txt(1) = frmTypeStr '--VType
If IsDate(mskDt(2)) Then '--Vdt
    txt(2) = CDate(mskDt(2))
Else
    txt(2) = ""
End If
'-- Broker Code
txt(3) = GProcGetColumnValue("tblMastAccount", "AcName", txttemp(3), "S", "AcCode", "N")
'-- Party Code
txt(4) = GProcGetColumnValue("tblMastAccount", "AcName", txttemp(4), "S", "AcCode", "N")
'-- Book Rt Type
txt(5) = cbo(5).ListIndex
txt(10) = cbo(10) '-- Del. Place
txt(7) = gCYear '-Year
txt(8) = gCCode  '-Ccode
txt(9) = gDbYr '-CYear code
txt(12) = Chk(0).Value
txt(13) = Chk(1).Value
txt(15) = ChkExGST.Value
End Sub
Private Sub FillTempFromTxt()
Dim X As String
If IsDate(txt(2)) Then '--VDt
    mskDt(2) = txt(2)
Else
    mskDt(2) = "__/__/____"
End If
If IsNull(txt(12)) Then txt(12) = 0
Chk(0).Value = Val(txt(12))
If IsNull(txt(13)) Then txt(13) = 0
Chk(1).Value = Val(txt(13))
If IsNull(txt(15)) Then txt(15) = 0
ChkExGST.Value = Val(txt(15))
cbo(5).ListIndex = txt(5) '-- Rt Type
cbo(10) = txt(10) '-- Del. Place
'--Broker
txttemp(3) = GProcGetColumnValue("tblMastAccount", "AcCode", txt(3), "N", "AcName", "S")
'--Party
txttemp(4) = GProcGetColumnValue("tblMastAccount", "AcCode", txt(4), "N", "AcName", "S")
'-- Fill Related Record
FillRelatedRecord
End Sub
Private Sub FillRelatedRecord()
Dim RowIndex As Long
Dim ColIndex As Long
Dim TotBag As Long
Dim TotWt As Double
'--Fill Grid
SetGrid
With rstblSub
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
                If ColIndex = 3 Then TotBag = TotBag + CLng(msGrid.TextMatrix(RowIndex, ColIndex)) '-- Tot Bag
                If ColIndex = 4 Then TotWt = TotWt + CDbl(msGrid.TextMatrix(RowIndex, ColIndex)) '-- Tot Wt
            Next ColIndex
            RowIndex = RowIndex + 1
            .MoveNext
            If Not .EOF Then msGrid.Rows = msGrid.Rows + 1
        Loop
    End If
End With
txttmpTotBag = TotBag '-- Tot Bag
txttmpTotWt = Format(TotWt, WtStr) '-- Tot Wt
'--- Fill Related Control
SetRelatedRS
End Sub
Private Sub ReadFields()
FormAction = vbDataActionClose
GProcShowRecord Me, rstbl, MaxNo
FillTempFromTxt
End Sub
Private Sub CalAmount()
Dim Amt As Double
Dim RowIndex As Integer
Dim TotBag As Long
Dim TotWt As Double
CalGridAmt
'---- Sub amt
GProcCheckForNumber Me, rstbl, MaxNo
CheckForNumberGrid
RowIndex = 1
With msGrid
    Do While RowIndex < .Rows
        '-- sub amt
        If Val(.TextMatrix(RowIndex, 7)) > 0 Then
            Amt = Amt + Val(.TextMatrix(RowIndex, 7))
        End If
        '--Tot Bag
        If Val(.TextMatrix(RowIndex, 3)) > 0 Then
            TotBag = TotBag + Val(.TextMatrix(RowIndex, 3))
        End If
         '--Tot Wt
        If Val(.TextMatrix(RowIndex, 4)) > 0 Then
            TotWt = TotWt + Val(.TextMatrix(RowIndex, 4))
        End If
        RowIndex = RowIndex + 1
    Loop
End With
txttmpTotBag = TotBag
txttmpTotWt = Format(TotWt, WtStr)
Amt = Format(Amt, FStr)
txt(6) = Amt
GProcCheckForNumber Me, rstbl, MaxNo
End Sub
'--- Is Related Record
Private Function IsRelatedRecord() As Boolean
Dim TmpCn As Connection
Dim Rs1 As Recordset
Dim Rs2 As Recordset
Dim RstCompMast As Recordset
Dim X As String
'=== Check Is Related Booking in Gate Pass Record for All company
'--- company Recordset
        GProcRstOpen RstCompMast, "Select * from tblCompanyYear,tblMastCompany  where compCode=cCode", "R", gSetupCn
        If RstCompMast.EOF = False Then RstCompMast.MoveFirst
        Do While Not RstCompMast.EOF
            '----Open Temp Connection for company
            If gBackEndDB = gBackEndAccess Then
                GProcConnectionOpen TmpCn, RstCompMast!CompCode + RstCompMast!CompYear
            Else
                GProcConnectionOpenORA TmpCn, RstCompMast!CompCode + RstCompMast!CompYear
            End If
            '-- Open RS2 For Grid ItCtrlNo
            X = "Select BkSubItCtrlNo From tblBookingSub where Vtype='" & txt(1) & "' and Vno=" & txt(0) & " and Vyear=" & gCYear
            If GProcRstOpen(Rs2, X, "R", gSetupCn) > 0 Then
                Do While Not Rs2.EOF
                    '-- Open Rs1 To Check Is Gate Pass
                    X = "Select Vno,Vdt from tblVouSub where Vtype='GP' and VSUBBKITCTRLNO =" & Rs2.Fields(0)
                    If GProcRstOpen(Rs1, X, "R") > 0 Then
                        IsRelatedRecord = True
                        MsgBox "You can not Modify/Delete this record.Related Gate Pass No. " & Rs1.Fields("Vno") & " is exists in " & RstCompMast!CompCode & " company.", vbCritical + vbOKOnly, App.Title
                        Rs1.Close
                        Exit Function
                    End If
                    Rs1.Close
                    Rs2.MoveNext
                Loop
                Rs2.Close
              End If
            RstCompMast.MoveNext
        Loop
        If RstCompMast.State = adStateOpen Then RstCompMast.Close
        If TmpCn.State = adStateOpen Then TmpCn.Close
End Function
