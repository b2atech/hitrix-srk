VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form frmTDSEntry 
   BackColor       =   &H00C2E3FC&
   BorderStyle     =   0  'None
   Caption         =   "Purchase Entry"
   ClientHeight    =   7170
   ClientLeft      =   0
   ClientTop       =   15
   ClientWidth     =   11010
   DrawMode        =   10  'Mask Pen
   FillColor       =   &H000040C0&
   FillStyle       =   0  'Solid
   Icon            =   "frmTDSEntry.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7170
   ScaleWidth      =   11010
   ShowInTaskbar   =   0   'False
   Visible         =   0   'False
   Begin VB.TextBox Txt 
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
      Index           =   22
      Left            =   7440
      TabIndex        =   81
      Text            =   "22"
      Top             =   5235
      Width           =   1710
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFC0C0&
      Caption         =   "TDS Printing"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4275
      Left            =   2520
      TabIndex        =   70
      Top             =   1440
      Visible         =   0   'False
      Width           =   7065
      Begin VB.TextBox txtAck4 
         Height          =   420
         Left            =   1920
         TabIndex        =   85
         Text            =   "TxtAck4"
         Top             =   3000
         Width           =   4770
      End
      Begin VB.TextBox txtAck3 
         Height          =   420
         Left            =   1920
         TabIndex        =   84
         Text            =   "TxtAck3"
         Top             =   2520
         Width           =   4770
      End
      Begin VB.TextBox txtAck2 
         Height          =   420
         Left            =   1920
         TabIndex        =   83
         Text            =   "TxtAck2"
         Top             =   2040
         Width           =   4770
      End
      Begin VB.TextBox txtAck1 
         Height          =   420
         Left            =   1920
         TabIndex        =   82
         Text            =   "TxtAck1"
         Top             =   1560
         Width           =   4770
      End
      Begin VB.TextBox TxtAuthBy 
         Height          =   420
         Left            =   1950
         TabIndex        =   76
         Text            =   "TxtAuthBy"
         Top             =   1065
         Width           =   4770
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
         Left            =   3570
         TabIndex        =   79
         Top             =   3675
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
         Left            =   1995
         TabIndex        =   78
         Top             =   3675
         Width           =   1335
      End
      Begin VB.TextBox txtno 
         Appearance      =   0  'Flat
         Height          =   360
         Index           =   1
         Left            =   5145
         MaxLength       =   7
         TabIndex        =   74
         Top             =   495
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
         Left            =   420
         TabIndex        =   77
         Top             =   3675
         Width           =   1335
      End
      Begin VB.TextBox txtno 
         Appearance      =   0  'Flat
         Height          =   360
         Index           =   0
         Left            =   1950
         MaxLength       =   7
         TabIndex        =   72
         Top             =   495
         Width           =   1575
      End
      Begin VB.Label Label23 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "4th Q.Ack.No    :"
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
         Left            =   360
         TabIndex        =   89
         Top             =   3120
         Width           =   1485
      End
      Begin VB.Label Label15 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "3rd  Q.Ack.No   :"
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
         Left            =   360
         TabIndex        =   88
         Top             =   2640
         Width           =   1485
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "2nd Q.Ack.No   :"
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
         Left            =   360
         TabIndex        =   87
         Top             =   2160
         Width           =   1470
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "1St Q.Ack.No   :"
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
         Left            =   360
         TabIndex        =   86
         Top             =   1560
         Width           =   1455
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Authorised &By  :"
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
         Left            =   360
         TabIndex        =   75
         Top             =   1065
         Width           =   1410
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
         Left            =   3870
         TabIndex        =   73
         Top             =   495
         Width           =   1050
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "From VNo      :"
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
         Left            =   360
         TabIndex        =   71
         Top             =   495
         Width           =   1365
      End
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
      Index           =   19
      ItemData        =   "frmTDSEntry.frx":058A
      Left            =   1650
      List            =   "frmTDSEntry.frx":0594
      TabIndex        =   20
      Text            =   "Cbo(19)"
      Top             =   5235
      Width           =   2475
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
      Height          =   360
      Index           =   21
      Left            =   405
      TabIndex        =   25
      Text            =   "txttemp(21)"
      ToolTipText     =   "Press F3 for Creation"
      Top             =   6570
      Width           =   3720
   End
   Begin VB.TextBox Txt 
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
      Index           =   21
      Left            =   10905
      TabIndex        =   69
      Text            =   "21"
      Top             =   525
      Visible         =   0   'False
      Width           =   420
   End
   Begin VB.TextBox Txt 
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
      Height          =   315
      Index           =   16
      Left            =   10965
      TabIndex        =   68
      Text            =   "16"
      Top             =   6300
      Visible         =   0   'False
      Width           =   420
   End
   Begin VB.TextBox Txt 
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
      Left            =   10980
      TabIndex        =   67
      Text            =   "15"
      Top             =   5040
      Visible         =   0   'False
      Width           =   420
   End
   Begin VB.TextBox Txt 
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
      Left            =   8340
      TabIndex        =   18
      Text            =   "18"
      Top             =   165
      Visible         =   0   'False
      Width           =   315
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
      Left            =   495
      TabIndex        =   14
      Top             =   3390
      Visible         =   0   'False
      Width           =   1035
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
      Index           =   9
      ItemData        =   "frmTDSEntry.frx":05A6
      Left            =   5790
      List            =   "frmTDSEntry.frx":05B9
      TabIndex        =   12
      Text            =   "Cbo(9)"
      Top             =   1470
      Width           =   2400
   End
   Begin VB.TextBox Txt 
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
      Height          =   360
      Index           =   0
      Left            =   1650
      TabIndex        =   1
      Text            =   "0"
      Top             =   675
      Width           =   1155
   End
   Begin VB.CommandButton cmdLedger 
      Caption         =   "Ledger"
      Height          =   375
      Index           =   0
      Left            =   8400
      TabIndex        =   66
      Top             =   1080
      Width           =   1215
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
      Height          =   360
      Left            =   5790
      TabIndex        =   10
      Text            =   "txtDrBalance"
      Top             =   1080
      Width           =   2400
   End
   Begin VB.TextBox Txt 
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
      Left            =   10935
      TabIndex        =   64
      Text            =   "1"
      Top             =   2190
      Visible         =   0   'False
      Width           =   420
   End
   Begin VB.TextBox Txt 
      Alignment       =   1  'Right Justify
      Height          =   375
      Index           =   2
      Left            =   10980
      TabIndex        =   53
      Text            =   "2"
      Top             =   1230
      Visible         =   0   'False
      Width           =   420
   End
   Begin VB.TextBox Txt 
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
      Left            =   9645
      TabIndex        =   3
      Text            =   "5"
      Top             =   315
      Visible         =   0   'False
      Width           =   270
   End
   Begin VB.TextBox Txt 
      Alignment       =   1  'Right Justify
      Height          =   375
      Index           =   3
      Left            =   10965
      TabIndex        =   52
      Text            =   "3"
      Top             =   1560
      Visible         =   0   'False
      Width           =   420
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
      TabIndex        =   46
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
      TabIndex        =   45
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
      Left            =   9570
      Style           =   1  'Graphical
      TabIndex        =   37
      ToolTipText     =   "Exit"
      Top             =   5955
      Width           =   855
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
      Left            =   8535
      Style           =   1  'Graphical
      TabIndex        =   31
      Top             =   5955
      Width           =   855
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
      Left            =   7635
      Style           =   1  'Graphical
      TabIndex        =   29
      ToolTipText     =   "To Save Record"
      Top             =   5955
      Width           =   855
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
      Left            =   9525
      Style           =   1  'Graphical
      TabIndex        =   44
      Top             =   7155
      Visible         =   0   'False
      Width           =   855
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
      Left            =   6630
      Style           =   1  'Graphical
      TabIndex        =   28
      ToolTipText     =   "To Delete Record"
      Top             =   5955
      Width           =   855
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
      Left            =   5730
      Style           =   1  'Graphical
      TabIndex        =   27
      ToolTipText     =   "To Modify Record"
      Top             =   5955
      Width           =   855
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
      Height          =   420
      Index           =   6
      Left            =   4830
      Style           =   1  'Graphical
      TabIndex        =   26
      ToolTipText     =   "To Add New Record"
      Top             =   5955
      Width           =   855
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
      Index           =   5
      Left            =   4830
      Style           =   1  'Graphical
      TabIndex        =   30
      Top             =   6495
      Width           =   855
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
      Left            =   5880
      Style           =   1  'Graphical
      TabIndex        =   32
      ToolTipText     =   "To Find Record"
      Top             =   6495
      Width           =   855
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
      Left            =   9090
      Style           =   1  'Graphical
      TabIndex        =   36
      ToolTipText     =   "Last"
      Top             =   6495
      Width           =   615
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
      Left            =   8355
      Style           =   1  'Graphical
      TabIndex        =   35
      ToolTipText     =   "First"
      Top             =   6495
      Width           =   615
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
      Left            =   7620
      Style           =   1  'Graphical
      TabIndex        =   34
      ToolTipText     =   "Next"
      Top             =   6495
      Width           =   615
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
      Left            =   6885
      Style           =   1  'Graphical
      TabIndex        =   33
      ToolTipText     =   "Previous"
      Top             =   6495
      Width           =   615
   End
   Begin VB.TextBox Txt 
      Alignment       =   1  'Right Justify
      Height          =   375
      Index           =   4
      Left            =   10950
      TabIndex        =   51
      Text            =   "4"
      Top             =   1875
      Visible         =   0   'False
      Width           =   420
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
      Height          =   360
      Index           =   4
      Left            =   5790
      TabIndex        =   9
      Text            =   "txttemp(4)"
      ToolTipText     =   "Press F3 for Creation"
      Top             =   675
      Width           =   4635
   End
   Begin VB.TextBox Txt 
      Alignment       =   1  'Right Justify
      Height          =   375
      Index           =   6
      Left            =   10980
      TabIndex        =   50
      Text            =   "6"
      Top             =   3000
      Visible         =   0   'False
      Width           =   420
   End
   Begin VB.TextBox Txt 
      Alignment       =   1  'Right Justify
      Height          =   375
      Index           =   7
      Left            =   10980
      TabIndex        =   49
      Text            =   "7"
      Top             =   3360
      Visible         =   0   'False
      Width           =   420
   End
   Begin VB.TextBox Txt 
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
      Left            =   10905
      TabIndex        =   38
      Text            =   "8"
      Top             =   2550
      Visible         =   0   'False
      Width           =   450
   End
   Begin VB.TextBox Txt 
      Alignment       =   1  'Right Justify
      Height          =   375
      Index           =   9
      Left            =   10980
      TabIndex        =   48
      Text            =   "9"
      Top             =   3720
      Visible         =   0   'False
      Width           =   420
   End
   Begin VB.TextBox Txt 
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
      Left            =   10980
      TabIndex        =   39
      Text            =   "10"
      Top             =   120
      Visible         =   0   'False
      Width           =   420
   End
   Begin VB.TextBox Txt 
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
      Index           =   11
      Left            =   10980
      TabIndex        =   42
      Text            =   "11"
      Top             =   5835
      Visible         =   0   'False
      Width           =   420
   End
   Begin VB.TextBox Txt 
      Alignment       =   1  'Right Justify
      Height          =   375
      Index           =   12
      Left            =   10980
      TabIndex        =   47
      Text            =   "12"
      Top             =   4080
      Visible         =   0   'False
      Width           =   420
   End
   Begin VB.TextBox Txt 
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
      Left            =   10980
      TabIndex        =   40
      Text            =   "13"
      Top             =   4515
      Visible         =   0   'False
      Width           =   420
   End
   Begin VB.TextBox Txt 
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
      Left            =   10980
      TabIndex        =   41
      Text            =   "14"
      Top             =   5430
      Visible         =   0   'False
      Width           =   420
   End
   Begin VB.TextBox Txt 
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
      Left            =   7995
      TabIndex        =   16
      Text            =   "17"
      Top             =   120
      Visible         =   0   'False
      Width           =   315
   End
   Begin VB.TextBox Txt 
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
      Index           =   19
      Left            =   30
      TabIndex        =   21
      Text            =   "19"
      Top             =   5040
      Visible         =   0   'False
      Width           =   270
   End
   Begin VB.TextBox Txt 
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
      Index           =   20
      Left            =   10965
      TabIndex        =   43
      Text            =   "20"
      Top             =   930
      Visible         =   0   'False
      Width           =   420
   End
   Begin MSMask.MaskEdBox mskDt 
      Height          =   360
      Index           =   3
      Left            =   2955
      TabIndex        =   2
      Top             =   675
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
   Begin Crystal.CrystalReport cryReport 
      Bindings        =   "frmTDSEntry.frx":05FB
      Left            =   2820
      Top             =   30
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
      Height          =   2910
      Left            =   105
      TabIndex        =   13
      Top             =   2100
      Width           =   10785
      _ExtentX        =   19024
      _ExtentY        =   5133
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
   Begin MSMask.MaskEdBox mskDt 
      Height          =   360
      Index           =   7
      Left            =   1650
      TabIndex        =   5
      Top             =   1080
      Width           =   1155
      _ExtentX        =   2037
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
      Index           =   8
      Left            =   1650
      TabIndex        =   7
      Top             =   1485
      Width           =   1155
      _ExtentX        =   2037
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
      Index           =   20
      Left            =   1650
      TabIndex        =   23
      Top             =   5640
      Width           =   2475
      _ExtentX        =   4366
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
   Begin VB.Shape Shape5 
      FillColor       =   &H00BDDFD1&
      FillStyle       =   0  'Solid
      Height          =   1230
      Left            =   4575
      Shape           =   4  'Rounded Rectangle
      Top             =   5820
      Width           =   6075
   End
   Begin VB.Label Label10 
      BackStyle       =   0  'Transparent
      Caption         =   "Total Tax Amount      :"
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
      Height          =   510
      Left            =   4695
      TabIndex        =   80
      Top             =   5250
      Width           =   2505
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Capital A/c  ( Person resposible for Deduction of Tax )     :"
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
      Height          =   480
      Left            =   405
      TabIndex        =   24
      Top             =   6060
      Width           =   3885
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "To Date      :"
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
      Left            =   405
      TabIndex        =   6
      Top             =   1485
      Width           =   1230
   End
   Begin VB.Label LblExmptAmt 
      BackStyle       =   0  'Transparent
      Caption         =   "BSR Code   :"
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
      Height          =   300
      Left            =   6660
      TabIndex        =   17
      Top             =   120
      Visible         =   0   'False
      Width           =   1440
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Payment Nature     :"
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
      Height          =   555
      Left            =   4695
      TabIndex        =   11
      Top             =   1470
      Width           =   1140
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Cheque No. :"
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
      Height          =   300
      Left            =   6600
      TabIndex        =   15
      Top             =   105
      Visible         =   0   'False
      Width           =   1440
   End
   Begin VB.Shape Shape8 
      BorderColor     =   &H00C00000&
      BorderWidth     =   5
      FillColor       =   &H00C00000&
      Height          =   7290
      Left            =   0
      Top             =   -135
      Width           =   10995
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "&Vou. No     :"
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
      Left            =   405
      TabIndex        =   0
      Top             =   675
      Width           =   1110
   End
   Begin VB.Label Label17 
      BackStyle       =   0  'Transparent
      Caption         =   "Place         :"
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
      Height          =   300
      Left            =   405
      TabIndex        =   19
      Top             =   5235
      Width           =   1530
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
      TabIndex        =   65
      Top             =   45
      Width           =   2385
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
      TabIndex        =   63
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
      TabIndex        =   62
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
      TabIndex        =   61
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
      TabIndex        =   60
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
      TabIndex        =   59
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
      TabIndex        =   58
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
      TabIndex        =   57
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
      TabIndex        =   56
      Top             =   7590
      Width           =   3135
   End
   Begin VB.Label Label16 
      BackStyle       =   0  'Transparent
      Caption         =   "Date           : "
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
      Left            =   405
      TabIndex        =   22
      Top             =   5640
      Width           =   1335
   End
   Begin VB.Label Label14 
      BackStyle       =   0  'Transparent
      Caption         =   "From Date  :"
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
      Left            =   405
      TabIndex        =   4
      Top             =   1080
      Width           =   1230
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
      TabIndex        =   55
      Top             =   8160
      Width           =   1095
   End
   Begin VB.Label Label9 
      BackStyle       =   0  'Transparent
      Caption         =   "Party     :"
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
      Height          =   570
      Left            =   4695
      TabIndex        =   8
      Top             =   675
      Width           =   1035
   End
   Begin VB.Label Label6 
      BackColor       =   &H00C00000&
      Caption         =   "    TDS Entry"
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
      Height          =   405
      Left            =   -15
      TabIndex        =   54
      Top             =   0
      Width           =   10995
   End
   Begin VB.Shape Shape4 
      BorderColor     =   &H00000080&
      FillColor       =   &H00EDFDFE&
      FillStyle       =   0  'Solid
      Height          =   630
      Left            =   4380
      Top             =   5100
      Width           =   6480
   End
   Begin VB.Shape Shape6 
      BorderColor     =   &H00000080&
      FillColor       =   &H00EDFDFE&
      FillStyle       =   0  'Solid
      Height          =   1530
      Left            =   165
      Top             =   495
      Width           =   4170
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H00000080&
      FillColor       =   &H00EDFDFE&
      FillStyle       =   0  'Solid
      Height          =   1530
      Left            =   4380
      Top             =   495
      Width           =   6480
   End
   Begin VB.Shape Shape3 
      BorderColor     =   &H00000080&
      FillColor       =   &H00EDFDFE&
      FillStyle       =   0  'Solid
      Height          =   1935
      Left            =   165
      Top             =   5100
      Width           =   4170
   End
End
Attribute VB_Name = "frmTDSEntry"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Const tblName As String = "tblTDS"
Const IdField As String = "Vno"
Const OrderField As String = "Vdt,Vno"
Const MaxNo As Long = 22
Dim rstbl As Recordset '-- TblTDS
Dim rstblSub As Recordset '-- TBlTDs
Dim varBookMark As Variant
Dim FirstTimeFlag As Boolean
Dim OldgFrmTypeStr As String
Dim FormAction As Integer
Dim frmTypeStr As String
Dim i As Long
'=== Form  Code
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case 13:
            If Me.ActiveControl.Name <> "txtGrid" Then SendKeys "{TAB}"
    End Select
End Sub
Private Sub Form_Load()
    Dim X As String
    Me.Top = 300
    Me.Left = Screen.Width / 2 - Me.Width / 2
    '--- TblTDS
    frmTypeStr = gfrmTypeStr
    X = "Select * from " & tblName & " where VCtrlNo=1 and VYear=" & gCYear
    If Trim(OrderField) <> "" Then X = X & " order by " & OrderField
    GProcRstOpen rstbl, X, "O"
    FirstTimeFlag = True
    '---Grid
    SetGrid
    '---- Button Visiblity Set As Per User Previlage
    GProcSetButtonVisiblity Me
    cmdBtn_Click (6)
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
'===== Command Ctrls Code
Private Sub cmdBtn_Click(Index As Integer)
On Error GoTo ErrorRoutine
Dim i As Long
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
    '-- From Date","To Date", "Party", "Tax Amt", "Challano No.","Ch. Date
    gClsSearch.SearchMultiField "tblTDS,TblMastAccount", "vno,TDSFromDt,TDSToDt,AcName,TDSTaxAmt,TDSchNo,TDSChdt", Array("VNo", "From Date", "To Date", "Party", "Tax Amt", "Challano No.", "Ch. Date"), Array(txt(0).Width, 1000, 1000, 2000, 1500, 1200, 1200), " tblTds.TDSAcCode=TblMastAccount.AcCode and vCtrlNo=1 and VYear=" & gCYear, "", "Vdt,Vno", 0, 0
    If gClsSearch.SearchMultiRetCol(0) <> "" Then
        rstbl.MoveFirst
        rstbl.Find "Vno = " & gClsSearch.SearchMultiRetCol(0)
        gClsSearch.SearchMultiRetCol(0) = ""
        ReadFields
    End If
Case 5 'Print
    Frame1.Visible = True
    txtno(0).Enabled = True
    txtno(1).Enabled = True
    TxtAuthBy.Enabled = True
    txtAck1.Enabled = True
    txtAck2.Enabled = True
    txtAck3.Enabled = True
    txtAck4.Enabled = True
    txtno(0) = txt(0)
    txtno(1) = txt(0)
    txtno(0).SetFocus
Case 6 'Add
    FormAction = vbDataActionAddNew
    Call GProcActivateControls(Me, True)
    Call GProcSetButtons(Me, vbDataActionAddNew)
    Call GProcClearForm(Me, rstbl, MaxNo, True)
    SetControlEd
    If FirstTimeFlag = True Then
        FirstTimeFlag = False
    Else
        txt(0).SetFocus
    End If
    txt(0) = GProcGenerateId(tblName, "Vno", "VYear=" & gCYear)
    mskDt(3) = GProcGenerateVDt(tblName, "Vdt", "VYear=" & gCYear) '- Vdt
   ' mskDt(6) = GProcGenerateVDt(tblName, "TDSChDt", "VYear=" & gCYear) '-Ch.Date
    mskDt(7) = gCYSDate  '- From Date
    mskDt(8) = gCYEDate  '- To Date
    mskDt(20) = GProcGenerateVDt(tblName, "Vdt", "VYear=" & gCYear) '-date
    FillCombo
    SetGrid
Case 7 'Modify
    'If IsRelatedRecord = False Then
        FormAction = vbDataActionUpdate
        Call GProcActivateControls(Me, True)
        Call GProcSetButtons(Me, vbDataActionUpdate)
        SetControlEd
        txt(0).Enabled = False
        mskDt(3).SetFocus
    'End If
Case 8 'Delete
    '  If IsRelatedRecord = False Then
        If MsgBox("Do You Really Delete This Record?", vbQuestion + vbYesNo + vbDefaultButton1, "Confirm Delete") = vbYes Then
            If rstbl.RecordCount > 1 Then
                '-- Delete Related Record
                DeleteRelatedRecord
                'Rstbl.Delete
                cmdBtn_Click (1)
            ElseIf rstbl.RecordCount = 1 Then
                '-- Delete Related Record
                DeleteRelatedRecord
                'Rstbl.Delete
                Call GProcActivateControls(Me, False)
                Call GProcSetButtons(Me, vbDataActionCancel)
                Form_Load
                SetControlEd
            End If
        End If
    'End If
Case 9 ''Print
Case 10 'Exit
'    DoEvents
    Unload Me
    Exit Sub
Case 11 'Label
    'frmLabel.Show 1
Case 12 'Optional
Case 13 'Save
    If ValidateData = True Then
        If Trim(txtGrid) <> "" Then
            msGrid.Text = txtGrid
            DoEvents
            txtGrid = ""
        End If
        txtGrid.Visible = False
        If FormAction = vbDataActionAddNew Then txt(0) = GProcGenerateId(tblName, IdField, "VYear=" & gCYear)
        FillTxtFromTemp
        '-- Save Record
        SaveRelatedRecord
        rstbl.Requery
        Call GProcActivateControls(Me, False)
        Call GProcSetButtons(Me, True)
        SetControlEd
        If FormAction = vbDataActionAddNew Then
            If cmdBtn(6).Visible = True Then cmdBtn(6).SetFocus
        Else
            If cmdBtn(7).Visible = True Then cmdBtn(7).SetFocus
        End If
        FormAction = vbDataActionClose
        msGrid.Col = 0
        msGrid.Row = 1
        txtGrid.Move msGrid.CellLeft + msGrid.Left, msGrid.CellTop + msGrid.Top, msGrid.CellWidth, msGrid.CellHeight
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
    SetControlEd
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
Private Sub cmdLedger_Click(Index As Integer)
On Error GoTo ErrorHandler
If Index = 0 Then frmRptAccount.txt(1) = txttemp(4)
frmMain.mnuRptRSubAccLeg_Click (0)
Exit Sub
ErrorHandler:
    GProcErrorHandler
End Sub
'=== Cbo Ctrls Code
Private Sub Cbo_KeyPress(Index As Integer, KeyAscii As Integer)
If Index = 9 Then KeyAscii = 0 '-- Pay Type
End Sub
Private Sub cbo_LostFocus(Index As Integer)
SaveBtnEd
End Sub

'=== Msk Ctrls Code
Private Sub mskDt_GotFocus(Index As Integer)
GProcSelectBox Me.ActiveControl
End Sub
Private Sub mskDt_LostFocus(Index As Integer)
Select Case Index
    Case 3, 7, 8
        If GProcIsDateValid(mskDt(Index)) = False Then
            mskDt(Index).SetFocus
            Exit Sub
        End If
End Select
If Index = 3 Then ShowAccountBalance
SaveBtnEd
End Sub
'== Txt Ctrls code
Private Sub txt_GotFocus(Index As Integer)
GProcSelectBox Me.ActiveControl
End Sub
Private Sub txt_LostFocus(Index As Integer)
CheckForNumber
CalAmount
SaveBtnEd
End Sub
Private Sub txt_KeyPress(Index As Integer, KeyAscii As Integer)
KeyAscii = GProcValidateKey(rstbl, Index, KeyAscii, txt(Index).Text)
End Sub
'===  TxtTemp Ctrls Code
Private Sub txttemp_GotFocus(Index As Integer)
GProcSelectBox Me.ActiveControl
End Sub
Private Sub txttemp_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
Dim GName As String
If KeyCode = vbKeyF3 Then
    Select Case Index
        Case 4 '-- Party
            GName = GProcGetColumnValue("TblMastGroup", "AgCode", 90015, "N", "AgName", "S")
            GProcShowForm frmMastAccount, frmMain.mnuMstfrm(0), Array(frmMastAccount.txttemp(2)), Array(GName)
        Case 21 '-- Capital a/c
            GName = GProcGetColumnValue("TblMastGroup", "AgCode", 90004, "N", "AgName", "S")
            GProcShowForm frmMastAccount, frmMain.mnuMstfrm(0), Array(frmMastAccount.txttemp(2)), Array(GName)
    End Select
End If
End Sub
Private Sub txtTemp_KeyPress(Index As Integer, KeyAscii As Integer)
Dim LeftPos As Long
Dim TopPos As Long
Dim AgCode As Long
TopPos = txttemp(Index).Top + Me.Top + 650
LeftPos = txttemp(Index).Left + Me.Left
If KeyAscii = 13 Then Exit Sub
Select Case Index
    Case 4  '-- Party   Alias Account
        gClsSearch.SearchMultiField "QryMastaccount", "Account,AcName,AcCity,AgName", Array("Party", "AcName", "City", "Group"), Array(3000, 0, 2000, 3000), "", Chr(KeyAscii), "acName", LeftPos, TopPos, True
        KeyAscii = 0
        If gClsSearch.SearchMultiRetCol(0) <> "" And gClsSearch.SearchMultiRetCol(1) = "" Then
            txttemp(Index) = gClsSearch.SearchMultiRetCol(0)
        Else
            If gClsSearch.SearchMultiRetCol(1) <> "" Then txttemp(Index) = gClsSearch.SearchMultiRetCol(1)
        End If
    Case 21 '-- Capital A/c  Alias Account
        gClsSearch.SearchMultiField "QryMastaccount", "Account,AcName,AcCity,AgName", Array("Capital Account", "AcName", "City", "Group"), Array(3000, 0, 2000, 3000), "(AgCode=90004 or GpCode=90004)", Chr(KeyAscii), "AcName", 0, 0, True
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
ShowAccountBalance
End Sub
'====  Procedures
Private Sub DeleteRelatedRecord()
'--- Tbl TDS
gCn.Execute "delete from tblTDS where vno=" & txt(0) & " and VYear=" & gCYear
End Sub
Private Sub SaveRelatedRecord()
Dim i As Integer
Dim j As Long
Dim X As String
Dim rsSub As Recordset
'-- Delete Record
DeleteRelatedRecord
'--- Save Records
'--- TblTDS
X = "select * from TblTDS where vno=" & txt(0) & " and VYear=" & gCYear
GProcRstOpen rsSub, X, "O"
With msGrid
    i = 1
    Do While i < .Rows
        rsSub.AddNew
        For j = 0 To MaxNo '--- Fill 0 to Max
            If txt(j) = "" Then
                rsSub.Fields(j) = Null
            Else
                rsSub.Fields(j) = txt(j)
            End If
        Next j
        '--Fill Grid
        rsSub!VCtrlNo = i '-- VCtrlNo (sr no )
        rsSub!TDSBillNo = .TextMatrix(i, 1) '--TDSBillNo (BillNo)
        rsSub!TDSCrDt = CDate(.TextMatrix(i, 2)) '--TDSCrDate (Cr Date)
        rsSub!TDSCrAmt = CDbl(.TextMatrix(i, 3)) '--TDSCramt (Cr Amt)
        rsSub!TDSTDSAmt = CDbl(.TextMatrix(i, 4)) '--TDSAmt (TDS Amt)
        rsSub!TDSSurchgAmt = CDbl(.TextMatrix(i, 5)) '--TDSSurchgAmt ( Surchg. Amt)
        rsSub!TDSEdCessAmt = CDbl(.TextMatrix(i, 6)) '--Ed Cess Amt (Ed Cess Amt)
        rsSub!TDSTaxAmt = CDbl(.TextMatrix(i, 7)) '--TDSTaxAmt (Tax Amt)
        rsSub!TDSChNo = .TextMatrix(i, 8) '--Challan No.
        If IsDate(.TextMatrix(i, 9)) = True Then '-- Challan Dt
            rsSub!TDSChDt = CDate(.TextMatrix(i, 9))
        Else
            rsSub!TDSChDt = Null
        End If
        rsSub!TDSChqNo = .TextMatrix(i, 10) '--Chq No.
        rsSub!TDSBSRNo = .TextMatrix(i, 11) '--BSR Code
'        If IsDate(.TextMatrix(i, 10)) = True Then '-- Deposit Dt
'            rsSub!TDSDepositDt = CDate(.TextMatrix(i, 10))
'        Else
'            rsSub!TDSDepositDt = Null
'        End If
        rsSub.Update
        i = i + 1
     Loop
 End With
 rsSub.Close
End Sub
Private Sub ShowAccountBalance()
Dim DrAcBal  As Double
Dim CrAcBal As Double
If IsDate(mskDt(3)) Then
    DrAcBal = GProcGetAccountBalance(txttemp(4), mskDt(3))
    txtDrBalance = IIf(DrAcBal > 0, DrAcBal & " Cr", Abs(DrAcBal) & " Dr")
End If
End Sub
Private Sub FillCombo()
'- Place
Call GProcFillCombo(cbo(19), "select distinct TDSPlace from tblTDS order by TDSPlace")
End Sub
Private Sub SaveBtnEd()
'--- Vno,vdt,Party,DedTax A/c,From Dt,To Dt,Place,Date
If CDbl(txt(0)) = 0 Or GProcIsDateValid(mskDt(3)) = False Or txttemp(4) = "" Or txttemp(21) = "" Or GProcIsDateValid(mskDt(7)) = False Or GProcIsDateValid(mskDt(8)) = False Or IsDate(mskDt(20)) = False Then
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
Dim GName As String
Dim rstCheck As Recordset
Dim AcCode As Long
    '--- Validate Grid
    If ValidateGrid = False Then
        'MSGrid.SetFocus
        Exit Function
    End If
    CalAmount
    '---- Required Data
    '--- Vno
    If CLng(txt(0)) = 0 Then
        MsgBox "Check Voucher No.", vbInformation + vbOKOnly, Me.Caption
        txt(0).SetFocus
        Exit Function
    End If
    '--- Vdt
    If GProcIsDateValid(mskDt(3)) = False Then
        MsgBox "Check Voucher Date.", vbInformation + vbOKOnly, Me.Caption
        mskDt(3).SetFocus
        Exit Function
    End If
    '--- From Date
    If GProcIsDateValid(mskDt(7)) = False Then
        MsgBox "Check From Date.", vbInformation + vbOKOnly, Me.Caption
        mskDt(7).SetFocus
        Exit Function
    End If
    '--- To Date
    If GProcIsDateValid(mskDt(8)) = False Then
        MsgBox "Check To Date.", vbInformation + vbOKOnly, Me.Caption
        mskDt(8).SetFocus
        Exit Function
    End If
    '--- Date
    If IsDate(mskDt(20)) = False Then
        MsgBox "Check Date.", vbInformation + vbOKOnly, Me.Caption
        mskDt(20).SetFocus
        Exit Function
    End If
    '---Place
    If cbo(19) = "" Then
        MsgBox "Check Place.", vbInformation + vbOKOnly, Me.Caption
        cbo(19).SetFocus
        Exit Function
    End If
    '---- Party Account
    If txttemp(4) = "" Then
        MsgBox "Check Party.", vbInformation + vbOKOnly, Me.Caption
        txttemp(4).SetFocus
        Exit Function
    End If
        '--Create Party
        If GProcGetColumnValue("tblMastAccount", "AcName", txttemp(4), "S", "AcCode", "N") = 0 Then
           GName = GProcGetColumnValue("TblMastGroup", "AgCode", 90015, "N", "AgName", "S")
           GProcShowForm frmMastAccount, frmMain.mnuMstfrm(0), Array(frmMastAccount.txt(1), frmMastAccount.txttemp(2)), Array(txttemp(4), GName)
           Exit Function
        End If
    '---- Capital A/c
    If txttemp(21) = "" Then
        MsgBox "Check Capital Account.", vbInformation + vbOKOnly, Me.Caption
        txttemp(21).SetFocus
        Exit Function
    End If
        '--Create Capital A/c
        If GProcGetColumnValue("tblMastAccount", "AcName", txttemp(21), "S", "AcCode", "N") = 0 Then
           GName = GProcGetColumnValue("TblMastGroup", "AgCode", 90004, "N", "AgName", "S")
           GProcShowForm frmMastAccount, frmMain.mnuMstfrm(0), Array(frmMastAccount.txt(1), frmMastAccount.txttemp(2)), Array(txttemp(21), GName)
           Exit Function
        End If
    '----Duplicate Records
    '----Addition Case
    If FormAction = vbDataActionAddNew Then
        X = "Select vno From " & tblName & " Where vno=" & CLng(txt(0)) & " and VCtrlNo=1 and VYear=" & gCYear
        i = GProcRstOpen(rstCheck, X, "R")
        If i > 0 Then
            MsgBox "Record is Available of No. " & CLng(txt(0)), vbCritical + vbOKOnly, "Duplicate Record"
            txt(0).SetFocus
            Exit Function
        End If
        rstCheck.Close
    End If
    '--- Duplicate Record for From-to Date
    AcCode = GProcGetColumnValue("tblMastAccount", "AcName", txttemp(4), "S", "AcCode", "N")
    If gBackEndDB = gBackEndAccess Then
        X = "Select vno From " & tblName & " Where vno<>" & CLng(txt(0)) & " and VYear=" & gCYear & " and " _
        & " Vctrlno=1  and cdate('" & mskDt(7) & "') between TDSFromDt and TDSToDt and TDSAcCode=" & AcCode
    Else
        X = "Select vno From " & tblName & " Where vno<>" & CLng(txt(0)) & " and VYear=" & gCYear & " and " _
        & " Vctrlno=1  and to_date('" & mskDt(7) & "','dd/MM/yyyy') between TDSFromDt and TDSToDt and TDSAcCode=" & AcCode
    End If
    i = GProcRstOpen(rstCheck, X, "R")
    If i > 0 Then
        MsgBox "Record is Available of No. " & rstCheck.Fields(0) & " for from Date, to Date Period.", vbCritical + vbOKOnly, "Duplicate Record"
        mskDt(7).SetFocus
        Exit Function
    End If
    rstCheck.Close
    ValidateData = True
End Function
Private Sub FillTxtFromTemp()
If IsDate(mskDt(3)) Then '--Vdt
    txt(3) = CDate(mskDt(3))
Else
    txt(3) = ""
End If
'If IsDate(mskDt(6)) Then '-- ch dt
'    txt(6) = CDate(mskDt(6))
'Else
'    txt(6) = ""
'End If
If IsDate(mskDt(7)) Then '-- From dt
    txt(7) = CDate(mskDt(7))
Else
    txt(7) = ""
End If
If IsDate(mskDt(8)) Then '-- To dt
    txt(8) = CDate(mskDt(8))
Else
    txt(8) = ""
End If
If IsDate(mskDt(20)) Then '-- date
    txt(20) = CDate(mskDt(20))
Else
    txt(20) = ""
End If
txt(11) = "" '-Cr Dt
txt(2) = gCYear
'-- Party Code
txt(4) = GProcGetColumnValue("tblMastAccount", "AcName", txttemp(4), "S", "AcCode", "N")
'-- Capital Code
txt(21) = GProcGetColumnValue("tblMastAccount", "AcName", txttemp(21), "S", "AcCode", "N")
txt(9) = cbo(9).ListIndex '- Pay Type
txt(19) = cbo(19) '-Place
ShowAccountBalance
End Sub
Private Sub FillTempFromTxt()
Dim X As String
FillCombo
If IsDate(txt(3)) Then '--VDt
    mskDt(3) = txt(3)
Else
    mskDt(3) = "__/__/____"
End If
'If IsDate(txt(6)) Then '-- Ch.Dt
'    mskDt(6) = txt(6)
'Else
'    mskDt(6) = "__/__/____"
'End If
If IsDate(txt(7)) Then '-- From Dt
    mskDt(7) = txt(7)
Else
    mskDt(7) = "__/__/____"
End If
If IsDate(txt(8)) Then '-- To Dt
    mskDt(8) = txt(8)
Else
    mskDt(8) = "__/__/____"
End If
If IsDate(txt(20)) Then '-- Date
    mskDt(20) = txt(20)
Else
    mskDt(20) = "__/__/____"
End If
'-- Party Code
txttemp(4) = GProcGetColumnValue("tblMastAccount", "AcCode", txt(4), "N", "AcName", "S")
'-- Capital Code
txttemp(21) = GProcGetColumnValue("tblMastAccount", "AcCode", txt(21), "N", "AcName", "S")
cbo(9).ListIndex = txt(9) '-- Pay Type
cbo(19) = txt(19) '-- Place
'-- Fill Related Record
FillRelatedRecord
ShowAccountBalance
End Sub
Private Sub FillRelatedRecord()
Dim RowIndex As Long
Dim ColIndex As Long
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
            Next ColIndex
            RowIndex = RowIndex + 1
            .MoveNext
            If Not .EOF Then msGrid.Rows = msGrid.Rows + 1
        Loop
    End If
End With
End Sub
Private Sub ReadFields()
FormAction = vbDataActionClose
GProcShowRecord Me, rstbl, MaxNo
FillTempFromTxt
End Sub
'=========== Grid Code
Private Sub SetGrid()
Dim X As String
'--- Set Grid
Set rstblSub = New Recordset
X = "select VCtrlNo,TDSBillNo,TDSCrDt,TDSCrAmt,TDSTDSAmt,TDSSurchgAmt,TDSEdCessAmt,TDSTaxAmt,TDSChNo," _
& " TDSChDt,TDSChqNo,TDSBSRNo from tblTDS where vno=" & txt(0) & " and VYear=" & gCYear & " order by VCtrlNo"
GProcRstOpen rstblSub, X, "R"
With msGrid
'SR|<Bill No  |<Cr Date      |<Cr Amount  |>TDS Amount |>Surchg.Amount |>Ed. Cess Amount |> Tax Amount| |
    .FormatString = ">SR|<Bill No |<Cr Date        |>Cr Amount |>TDS Amt   |>Surchrg Amt|>Ed. Cess Amt|>      Tax Amt|<Challan No.|<Ch. Date     |Chq/DD No.    |BSR Code      "
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
                 .TextMatrix(.Row, i) = Null & vbNullString
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
    msGrid_EnterCell
End Sub
Private Sub msGrid_GotFocus()
cmdBtn(10).Cancel = False
End Sub
Private Sub msGrid_EnterCell()
If FormAction <> vbDataActionCancel Then
    CheckForNumberTxtGrid
    DoEvents
    txtGrid = ""
    DoEvents
    If msGrid.Text <> "" Then txtGrid.Text = msGrid.Text
    '--- Lock TextBox
    With msGrid
        Select Case .Col
            Case 0: '---- Sr No
                txtGrid.Visible = True
                txtGrid.Locked = True
                ReNumberGridSrno
                DoEvents
                .Col = .Col + 1
            Case 7 '-- Tot Amount
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
        If (Val(.TextMatrix(.Row, 7)) = 0) And .Rows > 2 Then
            .RemoveItem (.Row)
            txtGrid.Text = ""
        End If
        txtGrid.Visible = False
        txt(18).SetFocus
        Exit Sub
    End If
End With

'If FormAction <> vbDataActionAddNew And FormAction <> vbDataActionUpdate Then Exit Sub
'---- Navigate Row ,Col
If KeyAscii = 13 Then
    KeyAscii = 0
    With msGrid
        If .Col = 11 Then 'rstblSub.Fields.Count - 2 Then
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
If msGrid.Text <> "" Then txtGrid.Text = msGrid.Text
If msGrid.Col < rstblSub.Fields.Count Then
    txtGrid.Move msGrid.CellLeft + msGrid.Left, msGrid.CellTop + msGrid.Top, msGrid.CellWidth, msGrid.CellHeight
End If
txtGrid.Visible = True
txtGrid.Enabled = True
DoEvents
If KeyAscii <> 0 Then
    Select Case rstblSub.Fields(msGrid.Col).Type
        Case adChar, adVarWChar
            txtGrid.Text = txtGrid.Text & Chr(KeyAscii)
        Case Else
            txtGrid.Text = Chr(GProcValidateKey(rstblSub, msGrid.Col, KeyAscii, txtGrid.Text)) 'Chr(GProcNumberWithDecimal(KeyAscii, txtGrid))
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
DoEvents
If msGrid.Col >= 1 Then
    CalGridAmt
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
Dim MillCode As Long
TopPos = txtGrid.Top + Me.Top + 650
LeftPos = txtGrid.Left + Me.Left
DoEvents
With msGrid
    If KeyAscii <> 13 And KeyAscii <> 27 Then
        '--- Search
        Select Case .Col
            Case Else
                KeyAscii = GProcValidateKey(rstblSub, msGrid.Col, KeyAscii, txtGrid)
        End Select
   End If
    If KeyAscii = 13 Then CheckForNumberTxtGrid
    '--- Navigate Row ,Col
    If KeyAscii = 13 And .Col < 11 Then
        .Col = .Col + 1
    ElseIf KeyAscii = 13 And .Col = 11 Then
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
            If (Val(.TextMatrix(.Row, 7)) = 0) And .Rows > 2 Then
                .RemoveItem (.Row)
                txtGrid.Text = ""
            End If
            txtGrid.Visible = False
            txt(17).SetFocus
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
    ReNumberGridSrno
    Exit Sub
End If
End Sub
Private Sub CalGridAmt()
Dim TotTaxAmt As Double
CheckForNumberGrid
'-- Tot Tax Amt  =TDS + Surcharge + ECess
TotTaxAmt = CDbl(msGrid.TextMatrix(msGrid.Row, 4)) + CDbl(msGrid.TextMatrix(msGrid.Row, 5)) + CDbl(msGrid.TextMatrix(msGrid.Row, 6))
msGrid.TextMatrix(msGrid.Row, 7) = Format(TotTaxAmt, FStr)
End Sub
Private Sub SetControlEd()
txt(22).Enabled = False '-- Total Tax Amt
txtDrBalance.Enabled = False
End Sub
Private Sub CalAmount()
Dim TotTaxAmt As Double
Dim RowIndex As Integer
CheckForNumberGrid
RowIndex = 1
With msGrid
    Do While RowIndex < .Rows
        '-- Total Tax Amt
        If CDbl(.TextMatrix(RowIndex, 7)) > 0 Then
            TotTaxAmt = TotTaxAmt + CDbl(.TextMatrix(RowIndex, 7))
        End If
        RowIndex = RowIndex + 1
    Loop
End With
txt(22) = Format(TotTaxAmt, FStr) '-- Total Tax Amt
End Sub
Private Sub CheckForNumber()
GProcCheckForNumber Me, rstbl, MaxNo
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
                    msGrid.TextMatrix(RowIndex, i) = Null & vbNullString
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
                txtGrid = Null & vbNullString
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
Dim RowIndex As Long
If Trim(txtGrid.Text) <> "" Then
    msGrid.Text = txtGrid.Text
    CalGridAmt
End If
txtGrid.Visible = False
txtGrid.Text = ""
CheckForNumberGrid
With msGrid
Validate:
    For RowIndex = 1 To .Rows - 1
        For i = 0 To rstblSub.Fields.Count - 1
            Select Case i
                Case 2 '-- Cr Date
                    '- Tax,Cr Amt=0 then delete row
                    If RowIndex > 1 And CDbl(.TextMatrix(RowIndex, 7)) = 0 And CDbl(.TextMatrix(RowIndex, 3)) = 0 Then
                        .RemoveItem (RowIndex)
                        GoTo Validate
                    End If
                    If GProcIsDateValid(.TextMatrix(RowIndex, i)) = False Then
                        MsgBox "Check Cr Date.", vbCritical + vbOKOnly, Me.Caption
                        .Col = i
                        msGrid.SetFocus
                        Exit Function
                    End If
                Case 3 '-- Cr Amt
                    If CDbl(.TextMatrix(RowIndex, i)) = 0 Then
                        MsgBox "Check Cr Amount.", vbCritical + vbOKOnly, Me.Caption
                        .Col = i
                        msGrid.SetFocus
                        Exit Function
                    End If
                Case 7 '--  Tax Amt
                    If CDbl(.TextMatrix(RowIndex, i)) = 0 Then
                        MsgBox "Check Tax Amount.", vbCritical + vbOKOnly, Me.Caption
                        .Col = i
                        msGrid.SetFocus
                        Exit Function
                    End If
           End Select
        Next i
      Next RowIndex
End With
ValidateGrid = True
ReNumberGridSrno
End Function
'==========  Print
Private Sub txtno_GotFocus(Index As Integer)
GProcSelectBox Me.ActiveControl
End Sub
Private Sub txtno_KeyPress(Index As Integer, KeyAscii As Integer)
KeyAscii = GProcNumberOnly(KeyAscii)
End Sub
Private Sub TxtAuthBy_GotFocus()
GProcSelectBox Me.ActiveControl
End Sub
Private Sub TxtAuthBy_LostFocus()
TxtAuthBy = GProcProperCase(TxtAuthBy)
txtAck1.SetFocus
End Sub
Private Sub cmdPrint_Click(Index As Integer)
Dim X As String
Dim Rs As Recordset
Select Case Index
    Case 0, 1:
        '------ Validate
        If IsNumeric(txtno(0)) = False Then txtno(0) = 0
        If IsNumeric(txtno(1)) = False Then txtno(1) = txtno(0)
        If CLng(txtno(1)) < CLng(txtno(0)) Then txtno(1) = txtno(0)
        '------- Set Crystal Properties
        If Index = 0 Then GProcCrystalRptPreparation CryReport, "Printer"
        If Index = 1 Then GProcCrystalRptPreparation CryReport, "Window"
        X = "Select * From tblTDS where VCtrlNo=1 and vno>= " & txtno(0) & " and Vno<=" & txtno(1) & " and VYear=" & gCYear & " order by vdt,vno"
        If GProcRstOpen(Rs, X, "R") <= 0 Then
           MsgBox "No Records..."
           Exit Sub
        End If
        With CryReport
            .Formulas(5) = "CPanNo = '" & gCPAN & "'" '-Pan No
            .Formulas(6) = "CTDSNo = '" & gCTDS & "'" '-TDS No
            .Formulas(7) = "CTDSCircle= '" & gCTDSCircle & "'" '-TDS  circle
            .Formulas(8) = "CDesignation = '" & gCDesignation & "'" '- Designation
            .Formulas(9) = "AuthoriseBy = '" & GProcProperCase(Trim(TxtAuthBy)) & "'" '-- Authorise By
            .Formulas(10) = "Ack1 = '" & GProcProperCase(Trim(txtAck1)) & "'" '-- Ack1
            .Formulas(11) = "Ack2 = '" & GProcProperCase(Trim(txtAck2)) & "'" '--
            .Formulas(14) = "Ack3 = '" & GProcProperCase(Trim(txtAck3)) & "'" '--
            .Formulas(13) = "Ack4 = '" & GProcProperCase(Trim(txtAck4)) & "'" '--
            .WindowTitle = "TDS Form"
'            .ReportFileName = gReportPath & "rptTDSFormPrint.rpt"
            
            .ReportFileName = gReportPath & "rptTDSFormPrint.rpt"
            
            
            Rs.MoveFirst
            Do While Not Rs.EOF
               '-- Set Selection formula
               X = "{tblTDS.Vno}=" & Rs.Fields("Vno") & " and {tblTDS.VYear}=" & gCYear
              .Formulas(12) = "AmtInWords = '" & GProcAmountInWords(Rs!TDSTotTaxAmt) & "'"
               .SelectionFormula = X
               .Action = 1
               Rs.MoveNext
            Loop
        End With
    Case 2:
        DoEvents
        Frame1.Visible = False
    End Select
End Sub

