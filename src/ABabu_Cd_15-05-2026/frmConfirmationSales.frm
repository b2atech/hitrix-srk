VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form frmConfirmationSales 
   BackColor       =   &H00C2E3FC&
   BorderStyle     =   0  'None
   Caption         =   "Sales Confirmation Entry"
   ClientHeight    =   7200
   ClientLeft      =   0
   ClientTop       =   15
   ClientWidth     =   9120
   DrawMode        =   10  'Mask Pen
   FillColor       =   &H000040C0&
   FillStyle       =   0  'Solid
   Icon            =   "frmConfirmationSales.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7200
   ScaleWidth      =   9120
   ShowInTaskbar   =   0   'False
   Visible         =   0   'False
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
      Index           =   32
      Left            =   10680
      TabIndex        =   83
      Text            =   "32"
      ToolTipText     =   "Press F1 for Selection, F3 for Creation"
      Top             =   1440
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
      Height          =   360
      Index           =   31
      Left            =   240
      TabIndex        =   81
      Text            =   "31"
      ToolTipText     =   "Press F1 for Selection, F3 for Creation"
      Top             =   6120
      Width           =   2025
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
      Index           =   12
      ItemData        =   "frmConfirmationSales.frx":058A
      Left            =   2790
      List            =   "frmConfirmationSales.frx":0594
      TabIndex        =   12
      Text            =   "Cbo(12)"
      Top             =   5040
      Width           =   1230
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
      Index           =   30
      Left            =   5040
      TabIndex        =   15
      Text            =   "30"
      ToolTipText     =   "Press F1 for Selection, F3 for Creation"
      Top             =   5520
      Width           =   2025
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
      Index           =   29
      Left            =   5040
      TabIndex        =   14
      Text            =   "29"
      ToolTipText     =   "Press F1 for Selection, F3 for Creation"
      Top             =   5040
      Width           =   2025
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
      Index           =   28
      Left            =   4080
      TabIndex        =   13
      Text            =   "28"
      ToolTipText     =   "Press F1 for Selection, F3 for Creation"
      Top             =   5040
      Width           =   585
   End
   Begin VB.TextBox txt 
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
      Index           =   27
      Left            =   9000
      TabIndex        =   76
      Text            =   "27"
      ToolTipText     =   "Press F1 for Selection, F3 for Creation"
      Top             =   3840
      Visible         =   0   'False
      Width           =   465
   End
   Begin VB.TextBox txt 
      Alignment       =   1  'Right Justify
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
      Index           =   26
      Left            =   5040
      TabIndex        =   16
      Text            =   "26"
      ToolTipText     =   "Press F1 for Selection, F3 for Creation"
      Top             =   6120
      Width           =   2025
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
      Index           =   25
      Left            =   5040
      TabIndex        =   11
      Text            =   "25"
      ToolTipText     =   "Press F1 for Selection, F3 for Creation"
      Top             =   4440
      Width           =   2025
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
      Index           =   24
      Left            =   4320
      TabIndex        =   10
      Text            =   "24"
      ToolTipText     =   "Press F1 for Selection, F3 for Creation"
      Top             =   4440
      Width           =   585
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
      Index           =   23
      ItemData        =   "frmConfirmationSales.frx":05A6
      Left            =   9960
      List            =   "frmConfirmationSales.frx":05B0
      TabIndex        =   22
      Text            =   "Cbo(23)"
      Top             =   2160
      Visible         =   0   'False
      Width           =   930
   End
   Begin VB.TextBox txt 
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
      Index           =   22
      Left            =   10050
      TabIndex        =   32
      Text            =   "22"
      ToolTipText     =   "Press F1 for Selection, F3 for Creation"
      Top             =   6780
      Visible         =   0   'False
      Width           =   465
   End
   Begin VB.TextBox txt 
      Alignment       =   1  'Right Justify
      Height          =   375
      Index           =   23
      Left            =   7230
      TabIndex        =   74
      Text            =   "23"
      Top             =   885
      Visible         =   0   'False
      Width           =   420
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFC0C0&
      Caption         =   "Sale Confirmation Letter"
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
      Left            =   1200
      TabIndex        =   66
      Top             =   2040
      Visible         =   0   'False
      Width           =   4980
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
         Left            =   3345
         TabIndex        =   71
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
         Left            =   1860
         TabIndex        =   70
         Top             =   1560
         Width           =   1335
      End
      Begin VB.TextBox txtno 
         Appearance      =   0  'Flat
         Height          =   360
         Index           =   1
         Left            =   2640
         MaxLength       =   7
         TabIndex        =   69
         Top             =   1080
         Width           =   1695
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
         Left            =   345
         TabIndex        =   68
         Top             =   1560
         Width           =   1335
      End
      Begin VB.TextBox txtno 
         Appearance      =   0  'Flat
         Height          =   360
         Index           =   0
         Left            =   2640
         MaxLength       =   7
         TabIndex        =   67
         Top             =   600
         Width           =   1695
      End
      Begin VB.Label Label23 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "To No     :"
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
         TabIndex        =   73
         Top             =   1080
         Width           =   915
      End
      Begin VB.Label Label31 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "From No :"
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
         TabIndex        =   72
         Top             =   600
         Width           =   930
      End
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
      Index           =   6
      Left            =   1515
      TabIndex        =   3
      Text            =   "txttemp(6)"
      ToolTipText     =   "Press F3 for Creation"
      Top             =   1200
      Width           =   5145
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
      Index           =   5
      Left            =   1515
      TabIndex        =   4
      Text            =   "txttemp(5)"
      ToolTipText     =   "Press F3 for Creation"
      Top             =   1800
      Width           =   5145
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
      ItemData        =   "frmConfirmationSales.frx":05C3
      Left            =   10200
      List            =   "frmConfirmationSales.frx":05DC
      TabIndex        =   21
      Text            =   "Cbo(9)"
      Top             =   2760
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.TextBox txt 
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
      Index           =   21
      Left            =   10050
      TabIndex        =   31
      Text            =   "21"
      ToolTipText     =   "Press F1 for Selection, F3 for Creation"
      Top             =   6435
      Visible         =   0   'False
      Width           =   465
   End
   Begin VB.TextBox txt 
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
      Index           =   16
      Left            =   10050
      TabIndex        =   28
      Text            =   "16"
      Top             =   5400
      Visible         =   0   'False
      Width           =   465
   End
   Begin VB.TextBox txt 
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
      Index           =   15
      Left            =   10050
      TabIndex        =   24
      Text            =   "15"
      ToolTipText     =   "Press F1 for Selection, F3 for Creation"
      Top             =   3885
      Visible         =   0   'False
      Width           =   465
   End
   Begin VB.TextBox txt 
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
      Index           =   18
      Left            =   10050
      TabIndex        =   30
      Text            =   "18"
      Top             =   6060
      Visible         =   0   'False
      Width           =   465
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
      Height          =   360
      Index           =   0
      Left            =   1515
      TabIndex        =   1
      Text            =   "0"
      Top             =   615
      Width           =   1500
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
      Left            =   7770
      TabIndex        =   64
      Text            =   "1"
      Top             =   405
      Visible         =   0   'False
      Width           =   420
   End
   Begin VB.TextBox txt 
      Alignment       =   1  'Right Justify
      Height          =   375
      Index           =   2
      Left            =   7650
      TabIndex        =   53
      Text            =   "2"
      Top             =   450
      Visible         =   0   'False
      Width           =   420
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
      Index           =   5
      Left            =   7545
      TabIndex        =   44
      Text            =   "5"
      Top             =   735
      Visible         =   0   'False
      Width           =   270
   End
   Begin VB.TextBox txt 
      Alignment       =   1  'Right Justify
      Height          =   375
      Index           =   3
      Left            =   8085
      TabIndex        =   52
      Text            =   "3"
      Top             =   630
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
      TabIndex        =   48
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
      TabIndex        =   47
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
      Left            =   7560
      Style           =   1  'Graphical
      TabIndex        =   45
      ToolTipText     =   "Exit"
      Top             =   6390
      Width           =   1080
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
      Left            =   7560
      Style           =   1  'Graphical
      TabIndex        =   38
      Top             =   5790
      Width           =   1080
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
      Left            =   7560
      Style           =   1  'Graphical
      TabIndex        =   36
      ToolTipText     =   "To Save Record"
      Top             =   5310
      Width           =   1080
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
      Left            =   6750
      Style           =   1  'Graphical
      TabIndex        =   46
      Top             =   7155
      Visible         =   0   'False
      Width           =   870
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
      Left            =   7560
      Style           =   1  'Graphical
      TabIndex        =   35
      ToolTipText     =   "To Delete Record"
      Top             =   4635
      Width           =   1080
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
      Left            =   7560
      Style           =   1  'Graphical
      TabIndex        =   34
      ToolTipText     =   "To Modify Record"
      Top             =   4155
      Width           =   1080
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
      Left            =   7560
      Style           =   1  'Graphical
      TabIndex        =   33
      ToolTipText     =   "To Add New Record"
      Top             =   3675
      Width           =   1080
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
      Left            =   7560
      Style           =   1  'Graphical
      TabIndex        =   37
      Top             =   3015
      Width           =   1080
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
      Left            =   7560
      Style           =   1  'Graphical
      TabIndex        =   39
      ToolTipText     =   "To Find Record"
      Top             =   2505
      Width           =   1080
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
      Left            =   8115
      Style           =   1  'Graphical
      TabIndex        =   43
      ToolTipText     =   "Last"
      Top             =   2025
      Width           =   525
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
      Left            =   7560
      Style           =   1  'Graphical
      TabIndex        =   42
      ToolTipText     =   "First"
      Top             =   2025
      Width           =   525
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
      Left            =   8115
      Style           =   1  'Graphical
      TabIndex        =   41
      ToolTipText     =   "Next"
      Top             =   1545
      Width           =   525
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
      Left            =   7560
      Style           =   1  'Graphical
      TabIndex        =   40
      ToolTipText     =   "Previous"
      Top             =   1545
      Width           =   525
   End
   Begin VB.TextBox txt 
      Alignment       =   1  'Right Justify
      Height          =   375
      Index           =   4
      Left            =   7875
      TabIndex        =   51
      Text            =   "4"
      Top             =   615
      Visible         =   0   'False
      Width           =   300
   End
   Begin VB.TextBox txt 
      Alignment       =   1  'Right Justify
      Height          =   375
      Index           =   6
      Left            =   8460
      TabIndex        =   50
      Text            =   "6"
      Top             =   495
      Visible         =   0   'False
      Width           =   420
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
      Index           =   7
      Left            =   1605
      TabIndex        =   5
      Text            =   "7"
      ToolTipText     =   "Press F1 for Selection"
      Top             =   2955
      Width           =   5430
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
      Index           =   8
      Left            =   360
      TabIndex        =   7
      Text            =   "8"
      Top             =   4425
      Width           =   750
   End
   Begin VB.TextBox txt 
      Alignment       =   1  'Right Justify
      Height          =   375
      Index           =   9
      Left            =   8295
      TabIndex        =   49
      Text            =   "9"
      Top             =   795
      Visible         =   0   'False
      Width           =   420
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
      Index           =   10
      Left            =   2880
      TabIndex        =   9
      Text            =   "10"
      Top             =   4425
      Width           =   1335
   End
   Begin VB.TextBox txt 
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
      Index           =   11
      Left            =   1605
      TabIndex        =   6
      Text            =   "11"
      ToolTipText     =   "Press F1 for Selection, F3 for Creation"
      Top             =   3360
      Width           =   5055
   End
   Begin VB.TextBox txt 
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
      Index           =   12
      Left            =   10050
      TabIndex        =   25
      Text            =   "12"
      ToolTipText     =   "Press F1 for Selection, F3 for Creation"
      Top             =   4275
      Visible         =   0   'False
      Width           =   465
   End
   Begin VB.TextBox txt 
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
      Index           =   13
      Left            =   10050
      TabIndex        =   26
      Text            =   "13"
      ToolTipText     =   "Press F1 for Selection, F3 for Creation"
      Top             =   4620
      Visible         =   0   'False
      Width           =   465
   End
   Begin VB.TextBox txt 
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
      Index           =   14
      Left            =   10050
      TabIndex        =   27
      Text            =   "14"
      ToolTipText     =   "Press F1 for Selection, F3 for Creation"
      Top             =   5025
      Visible         =   0   'False
      Width           =   465
   End
   Begin VB.TextBox txt 
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
      Index           =   17
      Left            =   10050
      TabIndex        =   29
      Text            =   "17"
      Top             =   5730
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
      Height          =   360
      Index           =   19
      Left            =   1320
      TabIndex        =   8
      Text            =   "19"
      Top             =   4425
      Width           =   1335
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
      Index           =   20
      Left            =   10320
      TabIndex        =   23
      Text            =   "20"
      Top             =   3480
      Visible         =   0   'False
      Width           =   255
   End
   Begin MSMask.MaskEdBox mskDt 
      Height          =   375
      Index           =   2
      Left            =   4980
      TabIndex        =   2
      Top             =   615
      Width           =   2055
      _ExtentX        =   3625
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
      Bindings        =   "frmConfirmationSales.frx":0611
      Left            =   4125
      Top             =   15
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      WindowControlBox=   -1  'True
      WindowMaxButton =   -1  'True
      WindowMinButton =   -1  'True
      PrintFileType   =   15
      PrintFileLinesPerPage=   60
   End
   Begin VB.Label Label12 
      BackStyle       =   0  'Transparent
      Caption         =   "Broker    : "
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
      Left            =   360
      TabIndex        =   84
      Top             =   1800
      Width           =   1380
   End
   Begin VB.Label Label11 
      BackStyle       =   0  'Transparent
      Caption         =   "Sub Rate "
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
      Left            =   240
      TabIndex        =   82
      Top             =   5760
      Width           =   1380
   End
   Begin VB.Label Label10 
      BackStyle       =   0  'Transparent
      Caption         =   "%"
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
      Left            =   4680
      TabIndex        =   80
      Top             =   5070
      Width           =   225
   End
   Begin VB.Label Label16 
      BackStyle       =   0  'Transparent
      Caption         =   "TaxType    :"
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
      Left            =   1680
      TabIndex        =   79
      Top             =   5040
      Width           =   1335
   End
   Begin VB.Label Label9 
      BackStyle       =   0  'Transparent
      Caption         =   "Bill Amount  : "
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
      Left            =   3480
      TabIndex        =   78
      Top             =   6150
      Width           =   1380
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "R.Off       : "
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
      Left            =   3840
      TabIndex        =   77
      Top             =   5550
      Width           =   1380
   End
   Begin VB.Line Line3 
      BorderColor     =   &H000000C0&
      X1              =   120
      X2              =   7185
      Y1              =   3840
      Y2              =   3840
   End
   Begin VB.Line Line2 
      BorderColor     =   &H000000C0&
      X1              =   120
      X2              =   7185
      Y1              =   4320
      Y2              =   4320
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Mill            : "
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
      Left            =   240
      TabIndex        =   75
      Top             =   3360
      Width           =   1380
   End
   Begin VB.Line Line1 
      BorderColor     =   &H000000C0&
      X1              =   120
      X2              =   7185
      Y1              =   6570
      Y2              =   6570
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Da&te      :"
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
      Left            =   3780
      TabIndex        =   17
      Top             =   615
      Width           =   855
   End
   Begin VB.Label Label8 
      BackStyle       =   0  'Transparent
      Caption         =   "Bag                  Weitht              Net  Rate      Per                       Amount"
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
      Left            =   600
      TabIndex        =   20
      Top             =   3960
      Width           =   6570
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "Count        :"
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
      Height          =   540
      Left            =   240
      TabIndex        =   19
      Top             =   2895
      Width           =   1380
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "Party      :"
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
      Left            =   360
      TabIndex        =   18
      Top             =   1200
      Width           =   1110
   End
   Begin VB.Shape Shape8 
      BorderColor     =   &H00C00000&
      BorderWidth     =   5
      FillColor       =   &H00C00000&
      Height          =   7245
      Left            =   15
      Top             =   -45
      Width           =   9045
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "&No          :"
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
      Left            =   360
      TabIndex        =   0
      Top             =   615
      Width           =   915
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
      Left            =   6150
      TabIndex        =   65
      Top             =   30
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
   Begin VB.Label Label6 
      BackColor       =   &H00C00000&
      Caption         =   "  Performa Invoice Entry"
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
      Left            =   0
      TabIndex        =   54
      Top             =   0
      Width           =   9045
   End
   Begin VB.Shape Shape1 
      FillColor       =   &H00BDDFD1&
      FillStyle       =   0  'Solid
      Height          =   5715
      Left            =   7380
      Shape           =   4  'Rounded Rectangle
      Top             =   1335
      Width           =   1455
   End
   Begin VB.Shape Shape6 
      BorderColor     =   &H00000080&
      FillColor       =   &H00EDFDFE&
      FillStyle       =   0  'Solid
      Height          =   4455
      Left            =   120
      Top             =   2700
      Width           =   7080
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H00000080&
      FillColor       =   &H00EDFDFE&
      FillStyle       =   0  'Solid
      Height          =   2040
      Left            =   120
      Top             =   465
      Width           =   7080
   End
End
Attribute VB_Name = "frmConfirmationSales"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Const tblName As String = "tblConfirmation"
Const IdField As String = "Vno"
Const OrderField As String = "Vdt,Vno"
Const MaxNo As Long = 32
Dim CondStr As String '-- Rstbl Recordset Condition String
Dim Rstbl As Recordset
Dim FirstTimeFlag As Boolean
Dim OldFrmType As String
Dim FormAction As Integer
Dim LastQualityRemark As String '- Last Quality Remark
Dim LastPackRemark As String '- Last Pack Remark
Private Sub Cbo_KeyPress(Index As Integer, KeyAscii As Integer)
If KeyAscii = 13 Then Exit Sub
If (Index = 9) Then KeyAscii = 0   '-Unit
End Sub
Private Sub cmdBtn_Click(Index As Integer)
On Error GoTo ErrorRoutine
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
    '-- Vno,Vdt,Acname,CfItemNarr,CfQty,CfRate
    gClsSearch.SearchMultiField "tblConfirmation,tblMastAccount", "Vno,Vdt,Acname,CfItemNarr,CfQty,CfRate", Array("No", "Date", "Seller", "Goods Desc.", "Qty", "Rate"), Array(1400, 1500, 2000, 1800, 900, 900), " CfCrAcCode=AcCode and " & CondStr, "", "Vdt,Vno ", txt(0).Left + Me.Left, txt(0).Top + Me.Top + 650
    If gClsSearch.SearchMultiRetCol(0) <> "" Then
        Rstbl.MoveFirst
        Rstbl.Find "Vno= " & gClsSearch.SearchMultiRetCol(0)
        gClsSearch.SearchMultiRetCol(0) = ""
        ReadFields
    End If
Case 5 'Print
    Frame1.Visible = True
    DoEvents
    txtno(0).Enabled = True
    txtno(1).Enabled = True
    txtno(0) = txt(0)
    txtno(1) = txt(0)
    txtno(0).SetFocus
Case 6 'Add
    FormAction = vbDataActionAddNew
    Call GProcActivateControls(Me, True)
    Call GProcSetButtons(Me, vbDataActionAddNew)
    Call GProcClearForm(Me, Rstbl, MaxNo, True)
    If FirstTimeFlag = True Then
        FirstTimeFlag = False
    Else
        txt(0).SetFocus
    End If
    txt(0) = GProcGenerateId(tblName, IdField, CondStr) '-Vno
    mskDt(2) = GProcGenerateVDt(tblName, "Vdt", CondStr) '- Vdt
    txt(15) = LastQualityRemark '- Quality Remark
    txt(12) = LastPackRemark '- Pack Remark
Case 7 'Modify
    'cmdBtn_Click 4 '-- Find
    'If gClsSearch.SearchMultiRetCol(0) = "" Then Exit Sub
    FormAction = vbDataActionUpdate
    Call GProcActivateControls(Me, True)
    Call GProcSetButtons(Me, vbDataActionUpdate)
    txt(0).Enabled = False
    mskDt(2).SetFocus
Case 8 'Delete
    If MsgBox("Do You Really Want to Delete This Record?", vbQuestion + vbYesNo + vbDefaultButton1, "Confirm Delete") = vbYes Then
        If Rstbl.RecordCount > 1 Then
            Rstbl.Delete
            cmdBtn_Click (1)
        ElseIf Rstbl.RecordCount = 1 Then
            Rstbl.Delete
            Call GProcActivateControls(Me, False)
            Call GProcSetButtons(Me, vbDataActionCancel)
            Form_Load
        End If
    End If
Case 9 'Print
Case 10 'Exit
    Unload Me
Case 11 'Label
    'frmLabel.Show 1
Case 12 'Creation To Firm
Case 13 'Save
    If ValidateData = True Then
        'If FormAction = vbDataActionAddNew Then txt(0) = GProcGenerateId(tblName, IdField, CondStr)
        FillTxtFromTemp
        GProcSaveRecord Me, Rstbl, FormAction, MaxNo
        Call GProcActivateControls(Me, False)
        Call GProcSetButtons(Me, True)
        If FormAction = vbDataActionAddNew Then
            If cmdBtn(6).Visible = True Then cmdBtn(6).SetFocus
        Else
            If cmdBtn(7).Visible = True Then cmdBtn(7).SetFocus
        End If
        FormAction = vbDataActionClose
    End If
Case 14 'Cancel
    Call GProcClearForm(Me, Rstbl, MaxNo, True)
    If Rstbl.RecordCount <= 0 Then
        Call GProcActivateControls(Me, False)
        Call GProcSetButtons(Me, vbDataActionCancel)
    Else
        If FormAction = vbDataActionAddNew Then
            Rstbl.MoveLast
        Else
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
End Select
Exit Sub
ErrorRoutine:
    Rstbl.CancelUpdate
    GProcErrorHandler
End Sub
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case 13:
            SendKeys "{TAB}"
        Case 27:
            If (cmdBtn(13).Enabled = False) Then
                Unload Me
            End If
    End Select
End Sub
Private Sub Form_Load()
    Dim X As String
    Me.Top = 300
    Me.Left = Screen.Width / 2 - Me.Width / 2
    CondStr = " Vtype='SA' and VYear=" & gCYear
    X = "Select * from " & tblName & " where " & CondStr
    If Trim(OrderField) <> "" Then X = X & " order by " & OrderField
    GProcRstOpen Rstbl, X, "O"
    '---- Button Visiblity Set As Per User Previlage
    GProcSetButtonVisiblity Me
    FirstTimeFlag = True
    cmdBtn_Click (6)
    Cbo(9).ListIndex = 0
    FillCombo
    Cbo(12).ListIndex = 0
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
If FormAction = vbDataActionUpdate Then
    MsgBox "You are about to modify the record. Save the record and then exit.", vbInformation + vbOKOnly, Me.Caption
    Cancel = True
    Exit Sub
End If
'mShowRec = Space(5)
Rstbl.Requery
Rstbl.Close
End Sub

Private Sub mskDt_GotFocus(Index As Integer)
GProcSelectBox Me.ActiveControl
End Sub
Private Sub txt_GotFocus(Index As Integer)
GProcSelectBox Me.ActiveControl
End Sub
Private Sub txt_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
Dim TopPos As Long
Dim LeftPos As Long
If KeyCode = vbKeyF3 Then '--F3 Create
    Select Case Index
        Case 11, 12, 13, 14, 15, 21, 22 '--Narration (quality Remark,Packing,Delivery,Payment,Remark )
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
        Case 11, 12, 13, 14, 15, 21, 22 '-- Narration (quality Remark,Packing,Delivery,Payment,Remark)
            gClsSearch.SearchMultiField "tblMastNarration", "Narration", Array("Narration"), Array(txt(Index).Width), "NarrType='N'", "", "Narration", 0, 0, True
            If gClsSearch.SearchMultiRetCol(0) <> "" Then txt(Index) = gClsSearch.SearchMultiRetCol(0)
        Case 7 '-- Goods Description
            gClsSearch.SearchMultiField "tblMastItem", "ItName", Array("Item"), Array(txt(Index).Width), "", "", "ItName", LeftPos, TopPos, True
            If gClsSearch.SearchMultiRetCol(0) <> "" Then txt(Index) = gClsSearch.SearchMultiRetCol(0)
    End Select
End If
End Sub
Private Sub txt_LostFocus(Index As Integer)
GProcCheckForNumber Me, Rstbl, MaxNo
SaveBtnEd
CalAmount
End Sub
Private Sub txt_KeyPress(Index As Integer, KeyAscii As Integer)
Select Case Index
    Case 7 '----Count Name
       gClsSearch.SearchMultiField "tblMastItem,tblMastAccount", "ItName,ItCode,Itticket,Acname,ItStdpack,ItStdrateper,ITUnit", Array("Item", "Item", "Ticket", "Mill", "Std.Packing", "Rate Per", "Unit"), Array(2000, 0, 1000, 2000, 1000, 1000, 1000), " tblMastItem.ItMillcode=tblMastAccount.AcCode", "", "ItName", txt(1).Left + Me.Left, txt(1).Top + Me.Top + 650
       KeyAscii = 0
       If gClsSearch.SearchMultiRetCol(1) <> "" Then
          txt(7) = gClsSearch.SearchMultiRetCol(0)
          txt(11) = gClsSearch.SearchMultiRetCol(3)
          txt(32) = GProcGetColumnValue("tblMastAccount", "AcName", txt(11), "S", "AcCode", "N")
          txt(8).SetFocus
       End If
 End Select
End Sub
Private Sub txttemp_GotFocus(Index As Integer)
GProcSelectBox Me.ActiveControl
End Sub
Private Sub txtTemp_LostFocus(Index As Integer)
SaveBtnEd
End Sub
Private Sub txtTemp_KeyPress(Index As Integer, KeyAscii As Integer)
Dim LeftPos As Long
Dim TopPos As Long
Dim BrkAcCode As Long
TopPos = txttemp(Index).Top + Me.Top + 650
LeftPos = txttemp(Index).Left + Me.Left
If KeyAscii = 13 Then Exit Sub
Select Case Index
    Case 6 '----Party Name
        gClsSearch.SearchMultiField "QryMastaccount", "Account,AcName,ACCode", Array("Party", "AcName", ""), Array(txttemp(Index).Width, 0, 0), " ", Chr(KeyAscii), "AcName", txttemp(Index).Left + Me.Left, txttemp(Index).Top + Me.Top + 1000
        KeyAscii = 0
        If gClsSearch.SearchMultiRetCol(1) <> "" Then
            txttemp(6) = gClsSearch.SearchMultiRetCol(1)
        End If
        '-- Broker from master
        If txttemp(5) = "" Then
            BrkAcCode = GProcGetColumnValue("TblMastAccount", "ACName", txttemp(Index), "S", "AcBrkCode", "N")
            txttemp(5) = GProcGetColumnValue("TblMastAccount", "ACCode", CStr(BrkAcCode), "N", "ACname", "S")
        End If
    Case 5 '-- Broker
        gClsSearch.SearchMultiField "QryMastaccount", "Account,AcName,AcCity,AgName", Array("Broker", "AcName", "City", "Group"), Array(3000, 0, 2000, 3000), "(AgCode=90016 or GpCode=90016)", Chr(KeyAscii), "AcName", LeftPos, TopPos, True
        KeyAscii = 0
        If gClsSearch.SearchMultiRetCol(0) <> "" And gClsSearch.SearchMultiRetCol(1) = "" Then
             txttemp(Index) = gClsSearch.SearchMultiRetCol(0)
        Else
             txttemp(Index) = gClsSearch.SearchMultiRetCol(1)
        End If
 End Select
End Sub
Private Sub txttemp_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
Dim GName As String
If KeyCode = vbKeyF3 Then
    Select Case Index
        Case 5, 6 '-- Credit ,Debit A/c
            GName = GProcGetColumnValue("TblMastGroup", "AgCode", 90015, "N", "AgName", "S")
            GProcShowForm frmMastAccount, frmMain.mnuMstfrm(0), Array(frmMastAccount.txttemp(2)), Array(GName)
    End Select
End If
End Sub
Private Sub SaveBtnEd()
'--- Vno,Vdt,Cr,Dr,Item
If CLng(txt(0)) = 0 Or GProcIsDateValid(mskDt(2)) = False Or Trim(txttemp(5)) = "" Or Trim(txttemp(6)) = "" Or Trim(txt(7)) = "" Then
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
    GProcCheckForNumber Me, Rstbl, MaxNo
    '--- Required
    '--- VNo
    If txt(0) = 0 Then
        MsgBox "Check Confirmation No.", vbCritical + vbOKOnly, Me.Caption
        txt(0).SetFocus
        Exit Function
    End If
    '--- VDt
    If GProcIsDateValid(mskDt(2)) = False Then
        MsgBox "Check Confirmation Date.", vbCritical + vbOKOnly, Me.Caption
        mskDt(2).SetFocus
        Exit Function
    End If
    '--- Credit A/c
    If Trim(txttemp(5)) = "" Then
        MsgBox "Check Seller.", vbCritical + vbOKOnly, Me.Caption
        txttemp(5).SetFocus
        Exit Function
    End If
    '-- Create Credit A/c
    If GProcGetColumnValue("tblMastAccount", "AcName", txttemp(5), "S", "AcCode", "N") = 0 Then
        GName = GProcGetColumnValue("TblMastGroup", "AgCode", 90015, "N", "AgName", "S")
        GProcShowForm frmMastAccount, frmMain.mnuMstfrm(0), Array(frmMastAccount.txt(1), frmMastAccount.txttemp(2)), Array(txttemp(5), GName)
        Exit Function
    End If
    '--- Debit A/c
    If Trim(txttemp(6)) = "" Then
        MsgBox "Check Buyer.", vbCritical + vbOKOnly, Me.Caption
        txttemp(6).SetFocus
        Exit Function
    End If
    '-- Create Debit A/c
    If GProcGetColumnValue("tblMastAccount", "AcName", txttemp(6), "S", "AcCode", "N") = 0 Then
        GName = GProcGetColumnValue("TblMastGroup", "AgCode", 90015, "N", "AgName", "S")
        GProcShowForm frmMastAccount, frmMain.mnuMstfrm(0), Array(frmMastAccount.txt(1), frmMastAccount.txttemp(2)), Array(txttemp(6), GName)
        Exit Function
    End If
    '-- Goods Desc.
    If Trim(txt(7)) = "" Then
        MsgBox "Check Goods Description.", vbCritical + vbOKOnly, App.Title
        txt(7).SetFocus
        Exit Function
    End If
    '----Duplicate Records
    '----Addition Case
    If FormAction = vbDataActionAddNew Then
        X = "Select Vno From " & tblName & " Where " & CondStr & " and Vno=" & CLng(txt(0))
        i = GProcRstOpen(rstCheck, X, "R")
        If i > 0 Then
            MsgBox "Record already exists of confirmation No. " & txt(0), vbCritical + vbOKOnly, "Duplicate Record"
            txt(0).SetFocus
            Exit Function
        End If
        rstCheck.Close
    End If
    ValidateData = True
End Function
Private Sub FillTxtFromTemp()
txt(3) = "SA" '-- Sales
txt(1) = gCYear '-Year
txt(9) = Cbo(9) '-- Unit
txt(23) = Cbo(23) '-- Price Type
txt(2) = CDate(mskDt(2)) '- Vdt
'-- Dr
txt(6) = GProcGetColumnValue("TBLMASTaccount", "acname", txttemp(6), "S", "accode", "N", "", gCn)
txt(5) = GProcGetColumnValue("TBLMASTaccount", "acname", txttemp(5), "S", "accode", "N", "", gCn)

txt(27) = GProcGetColumnValue("tblMastNarration", "Narration", Cbo(12), "S", "NarrCode", "N")
End Sub
Private Sub FillTempFromTxt()
Cbo(9) = txt(9) '-- Unit
Cbo(23) = txt(23) '-- Price Type
mskDt(2) = CDate(txt(2)) '- Vdt
'-- Dr
txttemp(6) = GProcGetColumnValue("TBLMASTaccount", "accode", txt(6), "N", "acname", "N", "", gCn)
txttemp(5) = GProcGetColumnValue("TBLMASTaccount", "accode", txt(5), "N", "acname", "N", "", gCn)
End Sub
Private Sub ReadFields()
GProcShowRecord Me, Rstbl, MaxNo
FillTempFromTxt
End Sub
Private Sub FillCombo()
Call GProcFillCombo(Cbo(12), "select Narration from tblMastNarration where NarrType ='T' order by Narration")
End Sub
'===  Print
Private Sub txtno_GotFocus(Index As Integer)
GProcSelectBox Me.ActiveControl
End Sub
Private Sub txtno_KeyPress(Index As Integer, KeyAscii As Integer)
KeyAscii = GProcNumberOnly(KeyAscii)
End Sub
Private Sub cmdPrint_Click(Index As Integer)
Dim X As String
Select Case Index
Case 0, 1:
    '------ Validate
    If IsNumeric(txtno(0)) = False Then txtno(0) = 0
    If IsNumeric(txtno(1)) = False Then txtno(1) = txtno(0)
    If CLng(txtno(1)) < CLng(txtno(0)) Then txtno(1) = txtno(0)
    '-------
    If Index = 0 Then GProcCrystalRptPreparation cryReport, "Printer"
    If Index = 1 Then GProcCrystalRptPreparation cryReport, "Window"
    cryReport.WindowTitle = "Sales Confirmation Letter"
    With cryReport
        .SelectionFormula = "{tblconfirmation.vNo} >= " & CLng(txtno(0)) & " and {tblconfirmation.vNo} <= " & CLng(txtno(1)) & "  and {tblconfirmation.VType}='SA' and {tblconfirmation.VYear}=" & gCYear
        .ReportFileName = gReportPath & "rptProFormaInvoice.rpt"
        .Action = 1
    End With
Case 2:
    DoEvents
    Frame1.Visible = False
End Select
End Sub
Private Sub CalAmount()
Dim TaxAmt As Double
Dim SubAmt As Double
Dim Amt As Double
Dim Rate As Double
Dim ROff As Double
Dim AssVal As Double
Dim X As String
Dim Qty As Long
Dim Wt As Double
Dim MillCode As Long
Qty = CLng(txt(8))
If Qty = 0 Then txt(19) = 0
If (CDbl(txt(19)) = 0) Then
    MillCode = GProcGetColumnValue("TblMastAccount", "AcName", txt(11), "S", "AcCode", "N")
    Wt = Qty * CDbl(GProcGetColumnValue("tblMastItem", "ItName", txt(7), "S", "ItStdpack", "N", " ItMillCode= " & MillCode))
    txt(19) = Wt
End If
CkeckForNumber
'-- Rate Per
If txt(24) = 0 Then txt(24) = 1
'---- Amt =[Wt* Net Rate/Rate Per]
Amt = (CDbl(txt(19)) * CDbl(txt(10)) / CDbl(txt(24)))
Amt = GProcMakeRounding(Amt)
'--------- ( Universal ) ----------------------------
'If frmTypeStr = "ST" Then
'   '-- Amt -LR Amt
'    Amt = Amt - txtVSub(21)
'    Amt = GProcMakeRounding(Amt)
'End If
'------------------( Universal End)
    
'--Sub Amt=(Amt)*100/(100+Tax %)
    
SubAmt = Amt * 100 / (100 + CDbl(txt(28)))
SubAmt = Format(SubAmt, FStr)
    
'---Tax Amt= Sub amt*Taxp/100
TaxAmt = SubAmt * CDbl(txt(28)) / 100
TaxAmt = Round(Format(TaxAmt, FStr), 0)
    
'--ROff = Amt - SubAmt - TaxAmt
ROff = Amt - SubAmt - TaxAmt
ROff = Format(ROff, FStr)
    
'--Rate=SubAmt * Rate Per /Wt
If CDbl(txt(19)) > 0 Then
   Rate = SubAmt * CDbl(txt(24)) / CDbl(txt(19))
   Rate = Format(Rate, LStr)
End If
AssVal = SubAmt '--Ass .Value
ROff = Amt - SubAmt - TaxAmt  'Round(Amt, 0) - Amt
ROff = Format(ROff, FStr)
txt(31) = Rate '--Rate
txt(25) = Format(SubAmt, FStr) '--SubAmt
txt(29) = TaxAmt '--Tax
txt(26) = Amt '--Amt
txt(30) = ROff '--round off
txt(25) = SubAmt '--Sub amt
'txt(24) = AssVal '-- Ass. Value
CkeckForNumber
End Sub


Private Sub CkeckForNumber()
GProcCheckForNumber Me, Rstbl, MaxNo, Array(31)
End Sub

