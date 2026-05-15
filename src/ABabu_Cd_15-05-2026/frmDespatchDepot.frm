VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form frmDespatchDepot 
   BackColor       =   &H00FBF2E1&
   BorderStyle     =   0  'None
   Caption         =   "Purchase Entry"
   ClientHeight    =   8430
   ClientLeft      =   0
   ClientTop       =   15
   ClientWidth     =   10485
   FillColor       =   &H000040C0&
   FillStyle       =   0  'Solid
   Icon            =   "frmDespatchDepot.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   8430
   ScaleWidth      =   10485
   ShowInTaskbar   =   0   'False
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
      Left            =   11520
      TabIndex        =   73
      Text            =   "30"
      Top             =   3840
      Visible         =   0   'False
      Width           =   375
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
      Left            =   11520
      TabIndex        =   72
      Text            =   "29"
      Top             =   3480
      Visible         =   0   'False
      Width           =   375
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
      Left            =   11520
      TabIndex        =   71
      Text            =   "28"
      Top             =   3120
      Visible         =   0   'False
      Width           =   375
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
      Index           =   27
      Left            =   11520
      TabIndex        =   70
      Text            =   "27"
      Top             =   2760
      Visible         =   0   'False
      Width           =   375
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
      Index           =   26
      Left            =   11520
      TabIndex        =   69
      Text            =   "26"
      Top             =   2400
      Visible         =   0   'False
      Width           =   375
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
      Left            =   11520
      TabIndex        =   68
      Text            =   "25"
      Top             =   2040
      Visible         =   0   'False
      Width           =   375
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
      Left            =   11520
      TabIndex        =   67
      Text            =   "24"
      Top             =   1680
      Visible         =   0   'False
      Width           =   375
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
      Index           =   23
      Left            =   11520
      TabIndex        =   66
      Text            =   "23"
      Top             =   1320
      Visible         =   0   'False
      Width           =   375
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
      Index           =   22
      Left            =   11520
      TabIndex        =   65
      Text            =   "22"
      Top             =   900
      Visible         =   0   'False
      Width           =   375
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
      Index           =   21
      Left            =   11520
      TabIndex        =   64
      Text            =   "21"
      Top             =   540
      Visible         =   0   'False
      Width           =   375
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
      Index           =   20
      Left            =   11520
      TabIndex        =   63
      Text            =   "20"
      Top             =   180
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.TextBox txtCode 
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
      Left            =   2145
      TabIndex        =   2
      Text            =   "txtCode(6)"
      ToolTipText     =   "Press F3 for Creation"
      Top             =   1920
      Width           =   1095
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
      Left            =   9120
      Style           =   1  'Graphical
      TabIndex        =   61
      ToolTipText     =   "To Find Record"
      Top             =   2400
      Width           =   975
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
      Left            =   9675
      Style           =   1  'Graphical
      TabIndex        =   60
      ToolTipText     =   "Next"
      Top             =   1320
      Width           =   420
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
      Left            =   9675
      Style           =   1  'Graphical
      TabIndex        =   59
      ToolTipText     =   "Last"
      Top             =   1800
      Width           =   420
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
      Left            =   9120
      Style           =   1  'Graphical
      TabIndex        =   58
      ToolTipText     =   "To Modify Record"
      Top             =   3600
      Width           =   975
   End
   Begin VB.Frame Frame1 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFC0C0&
      Caption         =   "Booking Printing"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   2865
      Left            =   12540
      TabIndex        =   49
      Top             =   3180
      Visible         =   0   'False
      Width           =   6165
      Begin VB.CommandButton cmdPrintBill 
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
         Left            =   4605
         TabIndex        =   55
         Top             =   2010
         Width           =   1335
      End
      Begin VB.CommandButton cmdPrintBill 
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
         Left            =   1710
         TabIndex        =   54
         Top             =   2025
         Width           =   1335
      End
      Begin VB.TextBox txtPInvno 
         Appearance      =   0  'Flat
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
         Index           =   1
         Left            =   2820
         MaxLength       =   7
         TabIndex        =   53
         Top             =   975
         Width           =   2055
      End
      Begin VB.CommandButton cmdPrintBill 
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
         Left            =   225
         TabIndex        =   52
         Top             =   2010
         Width           =   1335
      End
      Begin VB.TextBox txtPInvno 
         Appearance      =   0  'Flat
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
         Index           =   0
         Left            =   2820
         MaxLength       =   7
         TabIndex        =   51
         Top             =   405
         Width           =   2055
      End
      Begin VB.CommandButton cmdPrintBill 
         Caption         =   "&Mail"
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
         Left            =   3105
         TabIndex        =   50
         Top             =   2025
         Width           =   1335
      End
      Begin VB.Label Label39 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "To Booking No         :"
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
         Top             =   975
         Width           =   1995
      End
      Begin VB.Label Label38 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "From Booking  No    :"
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
         Top             =   495
         Width           =   2010
      End
      Begin VB.Line Line3 
         X1              =   0
         X2              =   6120
         Y1              =   1665
         Y2              =   1665
      End
   End
   Begin VB.TextBox txttemp 
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
      Index           =   6
      Left            =   2145
      TabIndex        =   3
      Text            =   "txttemp(6)"
      ToolTipText     =   "Press F3 for Creation"
      Top             =   2460
      Width           =   6315
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
      Index           =   16
      Left            =   12060
      TabIndex        =   46
      Text            =   "16"
      Top             =   5220
      Visible         =   0   'False
      Width           =   375
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
      Index           =   15
      Left            =   12060
      TabIndex        =   45
      Text            =   "15"
      Top             =   4980
      Visible         =   0   'False
      Width           =   375
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
      Index           =   13
      Left            =   12060
      TabIndex        =   44
      Text            =   "13"
      Top             =   4320
      Visible         =   0   'False
      Width           =   375
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
      Index           =   12
      Left            =   12060
      TabIndex        =   43
      Text            =   "12"
      Top             =   4020
      Visible         =   0   'False
      Width           =   375
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
      Index           =   9
      Left            =   12060
      TabIndex        =   42
      Text            =   "9"
      Top             =   3000
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.TextBox txttemp 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Index           =   17
      Left            =   2145
      TabIndex        =   4
      Text            =   "txttemp(17)"
      ToolTipText     =   "Press F3 for Creation"
      Top             =   3000
      Width           =   6315
   End
   Begin VB.TextBox txt 
      Alignment       =   1  'Right Justify
      Enabled         =   0   'False
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
      Index           =   19
      Left            =   2145
      TabIndex        =   6
      Text            =   "19"
      Top             =   4005
      Width           =   1695
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
      Left            =   9120
      Style           =   1  'Graphical
      TabIndex        =   38
      Top             =   4740
      Width           =   975
   End
   Begin VB.TextBox txt 
      Alignment       =   1  'Right Justify
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Index           =   18
      Left            =   2145
      TabIndex        =   5
      Text            =   "18"
      Top             =   3525
      Width           =   750
   End
   Begin VB.TextBox txt 
      Alignment       =   1  'Right Justify
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
      Left            =   12060
      TabIndex        =   9
      Text            =   "10"
      Top             =   3360
      Visible         =   0   'False
      Width           =   375
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
      Index           =   14
      Left            =   12060
      TabIndex        =   36
      Text            =   "14"
      Top             =   4680
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.CommandButton CmdBagDetails 
      Caption         =   "Cartoon &Bag Details"
      Height          =   720
      Left            =   12720
      TabIndex        =   34
      Top             =   180
      Visible         =   0   'False
      Width           =   870
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
      Left            =   12060
      TabIndex        =   12
      Text            =   "8"
      Top             =   2700
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.CommandButton cmdLedger 
      Caption         =   "Ledger"
      Height          =   375
      Left            =   12600
      TabIndex        =   33
      Top             =   2220
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.TextBox txtGrid 
      BackColor       =   &H80000000&
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
      Left            =   2640
      TabIndex        =   8
      Top             =   5520
      Visible         =   0   'False
      Width           =   1035
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
      Left            =   12060
      TabIndex        =   31
      Text            =   "1"
      Top             =   180
      Visible         =   0   'False
      Width           =   420
   End
   Begin VB.TextBox txt 
      Alignment       =   1  'Right Justify
      Height          =   375
      Index           =   2
      Left            =   12060
      TabIndex        =   29
      Text            =   "2"
      Top             =   540
      Visible         =   0   'False
      Width           =   420
   End
   Begin VB.TextBox txt 
      Alignment       =   1  'Right Justify
      Height          =   360
      Index           =   5
      Left            =   12060
      TabIndex        =   28
      Text            =   "5"
      Top             =   1680
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.TextBox txt 
      Alignment       =   1  'Right Justify
      Height          =   375
      Index           =   3
      Left            =   12060
      TabIndex        =   27
      Text            =   "3"
      Top             =   900
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
      Left            =   13200
      Style           =   1  'Graphical
      TabIndex        =   23
      Top             =   6720
      Visible         =   0   'False
      Width           =   1095
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
      Index           =   11
      Left            =   11460
      Style           =   1  'Graphical
      TabIndex        =   22
      ToolTipText     =   "To Add New Record"
      Top             =   8325
      Visible         =   0   'False
      Width           =   1275
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
      Left            =   9120
      Style           =   1  'Graphical
      TabIndex        =   20
      ToolTipText     =   "Exit"
      Top             =   5700
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
      Left            =   9120
      Style           =   1  'Graphical
      TabIndex        =   16
      ToolTipText     =   "To Save Record"
      Top             =   5220
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
      Left            =   12720
      Style           =   1  'Graphical
      TabIndex        =   21
      Top             =   1860
      Width           =   1275
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
      Left            =   9120
      Style           =   1  'Graphical
      TabIndex        =   15
      ToolTipText     =   "To Delete Record"
      Top             =   4080
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
      Height          =   405
      Index           =   6
      Left            =   9120
      Style           =   1  'Graphical
      TabIndex        =   14
      ToolTipText     =   "To Add New Record"
      Top             =   3120
      Width           =   975
   End
   Begin VB.CommandButton cmdBtn 
      Caption         =   "Cancel Bag"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   810
      Index           =   5
      Left            =   12840
      Style           =   1  'Graphical
      TabIndex        =   17
      Top             =   900
      Width           =   1005
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
      Left            =   9120
      Style           =   1  'Graphical
      TabIndex        =   19
      ToolTipText     =   "First"
      Top             =   1800
      Width           =   420
   End
   Begin VB.TextBox txt 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFFFF&
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
      Left            =   2100
      TabIndex        =   0
      Text            =   "0"
      Top             =   900
      Width           =   1830
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
      Left            =   9120
      Style           =   1  'Graphical
      TabIndex        =   18
      ToolTipText     =   "Previous"
      Top             =   1320
      Width           =   420
   End
   Begin VB.TextBox txt 
      Alignment       =   1  'Right Justify
      Height          =   375
      Index           =   4
      Left            =   12060
      TabIndex        =   26
      Text            =   "4"
      Top             =   1260
      Visible         =   0   'False
      Width           =   420
   End
   Begin VB.TextBox txt 
      Height          =   360
      Index           =   6
      Left            =   12060
      TabIndex        =   25
      Text            =   "6"
      Top             =   1980
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.TextBox txt 
      Alignment       =   1  'Right Justify
      Height          =   360
      Index           =   7
      Left            =   12060
      TabIndex        =   24
      Text            =   "7"
      Top             =   2340
      Visible         =   0   'False
      Width           =   375
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
      Index           =   11
      Left            =   12060
      TabIndex        =   13
      Text            =   "11"
      Top             =   3720
      Visible         =   0   'False
      Width           =   375
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
      Index           =   17
      Left            =   12060
      TabIndex        =   10
      Text            =   "17"
      Top             =   5520
      Visible         =   0   'False
      Width           =   375
   End
   Begin MSMask.MaskEdBox mskDt 
      Height          =   375
      Index           =   2
      Left            =   6060
      TabIndex        =   1
      Top             =   960
      Width           =   1560
      _ExtentX        =   2752
      _ExtentY        =   661
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
      Bindings        =   "frmDespatchDepot.frx":058A
      Left            =   4110
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
   Begin MSFlexGridLib.MSFlexGrid msGrid 
      Height          =   3510
      Left            =   2400
      TabIndex        =   7
      Top             =   4680
      Width           =   5550
      _ExtentX        =   9790
      _ExtentY        =   6191
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
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Mill  Code  :"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   375
      Left            =   720
      TabIndex        =   62
      Top             =   1920
      Width           =   1230
   End
   Begin VB.Shape Shape1 
      FillColor       =   &H00FED8FC&
      FillStyle       =   0  'Solid
      Height          =   5355
      Left            =   8970
      Shape           =   4  'Rounded Rectangle
      Top             =   1155
      Width           =   1305
   End
   Begin VB.Label Label13 
      BackStyle       =   0  'Transparent
      Caption         =   "Total    :"
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
      Left            =   15030
      TabIndex        =   48
      Top             =   6795
      Width           =   1365
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Despatch Details"
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
      Left            =   300
      TabIndex        =   47
      Top             =   4620
      Width           =   1665
   End
   Begin VB.Label Label12 
      BackStyle       =   0  'Transparent
      Caption         =   "Weight"
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
      Height          =   255
      Left            =   720
      TabIndex        =   41
      Top             =   4140
      Width           =   825
   End
   Begin VB.Label Label11 
      BackStyle       =   0  'Transparent
      Caption         =   "Bag"
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
      Left            =   720
      TabIndex        =   40
      Top             =   3540
      Width           =   825
   End
   Begin VB.Label Label9 
      BackStyle       =   0  'Transparent
      Caption         =   "Count "
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
      Left            =   720
      TabIndex        =   39
      Top             =   3045
      Width           =   870
   End
   Begin VB.Shape Shape5 
      BorderColor     =   &H00EFB17A&
      BorderWidth     =   5
      FillColor       =   &H00EFB17A&
      Height          =   8400
      Left            =   0
      Top             =   0
      Width           =   10455
   End
   Begin VB.Label Label20 
      BackStyle       =   0  'Transparent
      Caption         =   "Date              :    "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   255
      Left            =   4800
      TabIndex        =   37
      Top             =   960
      Width           =   1500
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Entry No    :"
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
      Left            =   840
      TabIndex        =   35
      Top             =   900
      Width           =   1095
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "M&ill Name  :"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   315
      Left            =   720
      TabIndex        =   11
      Top             =   2535
      Width           =   1230
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
      Left            =   7290
      TabIndex        =   32
      Top             =   0
      Width           =   2475
   End
   Begin VB.Label Label6 
      BackColor       =   &H00EFB17A&
      Caption         =   "     Despatch Details   ( Depot )"
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
      TabIndex        =   30
      Top             =   45
      Width           =   10440
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H00000080&
      FillColor       =   &H00EDFDFE&
      FillStyle       =   0  'Solid
      Height          =   2850
      Left            =   315
      Top             =   1665
      Width           =   8505
   End
   Begin VB.Shape Shape3 
      BorderColor     =   &H00000080&
      FillColor       =   &H00EDFDFE&
      FillStyle       =   0  'Solid
      Height          =   810
      Left            =   300
      Top             =   660
      Width           =   7845
   End
End
Attribute VB_Name = "frmDespatchDepot"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Const tblName As String = "tblBooking"
Const IdField As String = "Vno"
Const OrderField As String = "Vdt,Vno"
Const MaxNo As Long = 30
Const OtherTxtFromNo As Long = 0 '-- Other Txt Ctrl From No
Const OtherTxtToNo As Long = 0 '-- Other Txt Ctrl Less No
Dim rstbl As Recordset '-- TblPurch
Dim rstblSub As Recordset '-- TblPurchSub
Dim varBookMark As Variant
Dim FirstTimeFlag As Boolean, Criteria As String
Dim OldgFrmTypeStr As String
Dim FormAction As Integer
Dim frmTypeStr As String
Dim i As Long
Dim Rs1 As Recordset
Dim TotBag As Long
Dim Mon As Integer
'Dim GridFlag As Boolean
Private Sub cmdBtn_Click(Index As Integer)
'On Error GoTo ErrorRoutine
Select Case Index
Case 0 '---- Previous
     If GProcRstOpen(rstbl, "Select * from " & tblName & " where " & Criteria & "  and " & IdField & " = (Select max(" & IdField & ") from " & tblName & " where " & Criteria & " And " & IdField & "  < " & Val(txt(0)) & ")", "R") > 0 Then ReadFields Else MsgBox "This is First Record:", vbOKOnly + vbInformation
Case 1 '---- Next
      If GProcRstOpen(rstbl, "Select * from " & tblName & " where " & Criteria & "  and   " & IdField & " = (Select min(" & IdField & ") from " & tblName & " where " & IdField & " > '" & txt(0) & "' and " & Criteria & ") and " & Criteria, "R") > 0 Then ReadFields Else MsgBox "This is last Record:", vbOKOnly + vbInformation
Case 2 '---- First
         If GProcRstOpen(rstbl, "Select * from " & tblName & " where " & Criteria & "  and  " & IdField & " = (Select min(" & IdField & ") from " & tblName & " where " & IdField & " <> '' and " & Criteria & ")and " & Criteria, "R") > 0 Then ReadFields
Case 3 '---- Last
        If GProcRstOpen(rstbl, "Select * from " & tblName & " where " & Criteria & "  and  " & IdField & " = (Select max(" & IdField & ") from " & tblName & " where " & IdField & " <> '' and " & Criteria & ")and " & Criteria, "R") > 0 Then ReadFields Else GProcClearForm Me, rstbl, MaxNo
Case 4 'Find ' QryMastAccount
       gClsSearch.SearchMultiField "tblBooking,TblMastAccount", "Vno,AcName,Vdt", Array("Vno", "Mill", "Vdt"), Array(txt(0).Width, 4000, 2000), " tblBooking.BkMillCode=TblMastAccount.AcCode and vtype='" & frmTypeStr & "' and VFirm = '" & gCCode & "'   and VYear=" & gCYear, "", "Vno,Vdt", 0, 0
       If gClsSearch.SearchMultiRetCol(0) <> "" Then
          GProcRstOpen rstbl, "Select * from " & tblName & " where " & IdField & " = " & gClsSearch.SearchMultiRetCol(0) & " and " & Criteria, "R"
          gClsSearch.SearchMultiRetCol(1) = ""
          ReadFields
          FormAction = vbDataActionUpdate
          Call GProcActivateControls(Me, True)
          SetControlEd
          Call GProcSetButtons(Me, vbDataActionUpdate)
          txt(0).SetFocus
       End If
Case 6 'Add
    FormAction = vbDataActionAddNew
    Call GProcActivateControls(Me, True)
    SetControlEd
    Call GProcSetButtons(Me, vbDataActionAddNew)
    Call GProcClearForm(Me, rstbl, MaxNo, True, OtherTxtFromNo, OtherTxtToNo)
    txt(0) = GProcGenerateId("tblBooking", "Vno", "Vtype='" & frmTypeStr & "' and VFirm = '" & gCCode & "'   and VYear=" & gCYear)
    mskDt(2) = GProcGenerateVDt("tblBooking", "Vdt", "Vtype='" & frmTypeStr & "' and VFirm = '" & gCCode & "'   and VYear=" & gCYear)
    '--A/C
    SetGrid
 Case 7 'Modify
        If IsRelatedRecord = False Then
            cmdBtn_Click 4 '-- Find
            If gClsSearch.SearchMultiRetCol(0) = "" Then Exit Sub
            FormAction = vbDataActionUpdate
            Call GProcActivateControls(Me, True)
            SetControlEd
            Call GProcSetButtons(Me, vbDataActionUpdate)
            txt(0).SetFocus
        End If
Case 8 'Delete
    'If mURecDel = True Then
      If IsRelatedGatePass = True Then
            MsgBox "You can not Modify/Delete this record.Related Gate Pass is exists.", vbCritical + vbOKOnly, App.Title
            Exit Sub
      End If
      If IsRelatedRecord = False Then
        If MsgBox("Do You Really Delete This Record?", vbQuestion + vbYesNo + vbDefaultButton1, "Confirm Delete") = vbYes Then
           DeleteRelatedRecord
           cmdBtn_Click (3)
        End If
    End If
Case 9 'Print
    Frame1.Caption = "Booking Confirmation Printing "
        '-- Inv No
    txtPInvno(0) = txt(0).text
    txtPInvno(1) = txt(0).text
    Frame1.Visible = True
    DoEvents
    txtPInvno(0).Enabled = True
    txtPInvno(1).Enabled = True
    txtPInvno(0).SetFocus
Case 10 'Exit
'    DoEvents
    Unload Me
    Exit Sub
Case 11 'Label
    'frmLabel.Show 1
Case 12 'Optional
Case 13 'Save
    If ValidateGrid = True Then
        'txt(51) = CDate(Date)
        If Trim(txtGrid) <> "" Then
            msGrid.text = txtGrid
            DoEvents
            txtGrid = ""
        End If
        txtGrid.Visible = False
       ' If FormAction = vbDataActionAddNew Then txt(0) = GProcGenerateId(tblName, IdField, "vtype='" & FrmTypeStr & "' and VYear=" & gCYear)
        CalAmount
        FillTxtFromTemp
        '-- Save Record
'        If FormAction = vbDataActionAddNew Then
'           Mon = Month(mskDt(2))
'           txt(0) = GProcGenerateIdMonthwise("TblBooking", "Vno", Mon, "Vdt", "Vtype='" & frmTypeStr & "' and VFirm = '" & gCCode & "'   and VYear=" & gCYear)
'        End If
        GProcRstOpen rstbl, "Select * from " & tblName & " where " & IdField & " = " & txt(0) & " and " & Criteria, "O"
        txt(30) = 1
        GProcSaveRecord Me, rstbl, FormAction, MaxNo, OtherTxtFromNo, OtherTxtToNo
        SaveRelatedRecord
        ReadFields
        Call GProcActivateControls(Me, False)
        Call GProcSetButtons(Me, True)
        FormAction = vbDataActionClose
        msGrid.Col = 0 '-----
        msGrid.Row = 1 '-----
        txtGrid.Move msGrid.CellLeft + msGrid.Left, msGrid.CellTop + msGrid.Top, msGrid.CellWidth, msGrid.CellHeight '-----
'        If FormAction = vbDataActionAddNew Then
            If cmdBtn(6).Visible = True Then cmdBtn(6).SetFocus
 '       Else
  '          If cmdBtn(7).Visible = True Then cmdBtn(7).SetFocus
   '     End If
    End If
Case 14 'Cancel
    cmdBtn_Click (3)
    If rstbl.RecordCount <= 0 Then
        Call GProcActivateControls(Me, False)
        Call GProcSetButtons(Me, vbDataActionCancel)
    Else
        Call GProcActivateControls(Me, False)
        Call GProcSetButtons(Me, True)
    End If
    If FormAction = vbDataActionAddNew Then
        If cmdBtn(6).Visible = True Then cmdBtn(6).SetFocus
    Else
        If cmdBtn(7).Visible = True Then cmdBtn(7).SetFocus
    End If
    FormAction = vbDataActionCancel
    If msGrid.Enabled = True Then
       msGrid.Col = 0
       msGrid.Row = 1
       txtGrid.Move msGrid.CellLeft + msGrid.Left, msGrid.CellTop + msGrid.Top, msGrid.CellWidth, msGrid.CellHeight
    End If
End Select
Exit Sub
ErrorRoutine:
    rstbl.CancelUpdate
    GProcErrorHandler
End Sub
Private Sub SetControlEd()
txt(19).Enabled = False
End Sub
Private Sub DeleteRelatedRecord()
Dim mNo As Integer
mNo = GProcGetColumnValue("tblBooKingDesp", "Vno", txt(0), "N", "max(BkSrNo)", "N", Criteria)
For i = 1 To mNo

Dim X As String
X = "select * from tblBooKingDesp where vno=" & txt(0) & " and BkSrNo=" & i & " and " & Criteria
   gCn.Execute "Delete from tblBooKingDesp where vno=" & txt(0) & " and BkSrNo=" & i & " and " & Criteria
Next i
End Sub
Private Sub SaveRelatedRecord()
Dim i As Long
Dim MillCode As Long
Dim rstSub As Recordset
Dim ItemCode As Long
Dim X As String, wit As Long, wgd As Long, Y As String, maxy As Long
Dim BillDt
Dim DueDt
'--- Account Code
MillCode = GProcGetColumnValue("tblMastAccount", "AcName", txttemp(6), "S", "AcCode", "N")
ItemCode = GProcGetColumnValue("tblMastItem", "ItName", txttemp(17), "S", "Itcode", "N", "ItMillCode = " & MillCode & "")
'-- Delete Records
DeleteRelatedRecord
'--- Save Grid Records
'--- TblVouSub
X = "select * from tblBooKingDesp where vtype='" & frmTypeStr & "' and vno=" & txt(0) & " and Vfirm = '" & gCCode & "' and " _
& " VYear=" & gCYear
GProcRstOpen rstSub, X, "O"
'maxy = GProcGetColumnValue("tblPurchSub", "Vno", txt(0), "N", "max(PurSubItCtrlNo)", "N", Criteria)

With msGrid
    i = 1
    Do While i < .rows
        If Val(.TextMatrix(i, 1)) > 0 Then
            rstSub.AddNew
            rstSub!Vno = CLng(txt(0)) '--Vno(0)
            rstSub!Vtype = txt(1) & vbNullString  '--Vtype(1)
            rstSub!Vdt = CDate(txt(2)) '--Vdt(2)
            rstSub!VYear = txt(3) '--VYear
            rstSub!VFirm = gCCode '--VFirm
            rstSub!BkSrNo = i '--Sr No
            rstSub!BkDespItSrNo = Val(.TextMatrix(i, 0))   'i '--Sr No
            rstSub!BkDespDate = CDate(txt(2))  '--Des.Date
            rstSub!BkDespBag = CLng(.TextMatrix(i, 1)) '--Bag(1)
            rstSub!BkDespTmpChr2 = .TextMatrix(i, 2) & vbNullString
            rstSub!BkDespTmpChr1 = UCase(.TextMatrix(i, 3)) & vbNullString
            rstSub!BkDespItCode = ItemCode
            rstSub.Update
        End If
        i = i + 1
     Loop
 End With
 rstSub.Close
End Sub
Private Sub cmdPrintBill_Click(Index As Integer)
Dim RsSA As Recordset
Dim X As String
Dim XX As String
Dim BillSrNo As String
Dim InvNo As Long
Dim Y As String
Dim BilNo As String
Dim YY As String
Select Case Index
Case 0, 1, 3:
    '------ Validate
    If IsNumeric(txtPInvno(0)) = False Then txtPInvno(0) = 0
    If IsNumeric(txtPInvno(1)) = False Then txtPInvno(1) = txtPInvno(0)
    If CLng(txtPInvno(1)) < CLng(txtPInvno(0)) Then txtPInvno(1) = txtPInvno(0)
    '-------
    If Index = 0 Then GProcCrystalRptPreparation cryReport, "Printer"
    If Index = 1 Then GProcCrystalRptPreparation cryReport, "Window"
    If Index = 3 Then GProcCrystalRptPreparation cryReport, "Window"
    
    cryReport.WindowTitle = "Booking"
    '-- A/C Bill SrNo
    InvNo = CLng(txtPInvno(0))
    X = "("
    Do While InvNo <= txtPInvno(1)
       If BillSrNo = "" Then
          X = X + "'" & InvNo & "',"
       Else
          X = X + "'" & BillSrNo & "-" & InvNo & "',"
       End If
       InvNo = InvNo + 1
    Loop
    X = Left(X, Len(X) - 1) + ")"
    Y = X
    X = "Select a.Vno,a.Vtype,a.Vdt,a.Vyear,a.VFirm,BkParty,BkMillCode,BkBroker,BkRefNo,BkRefDt,BkSubAmt,BkBillAmt,BkRoff,DueDays,DueDate,BkEntDt,'" & gUserName & "',BkItCode,BkBag,BkWt,BkRt,BkRtPer,BkRateType,BkDeleAdd,BkPartyGroup,BkExtraCount,BkNote,BkCbag,BkCWt," _
        & " BkCDate,BkTmpNo1,BkTmpNo2,BkTmpNo3,BkTmpChr1,BkTmpChr2,BkTmpChr3,BkSubBag,BkSubWt,BkSubDDate " _
        & " From tblBooking A left outer join tblBookingSub B on a.vno = b.vno and a.vtype = b.vtype and a.vfirm = b.vfirm and a.vyear = b.vyear where a.vno in " & X & " and a.Vtype='" & frmTypeStr & "' and a.VFirm = '" & gCCode & "' And a.VYear = " & gCYear
     
   gCn.Execute "Delete from TmpBooking where  BkUser='" & gUserName & "'"
     
   gCn.Execute " insert into TmpBooking (Vno,Vtype,Vdt,Vyear,VFirm,BkParty,BkMillCode,BkBroker,BkRefNo,BkRefDt,BkSubAmt,BkBillAmt,BkRoff,DueDays,DueDate,BkEntDt,BkUser,BkItCode,BkBag,BkWt,BkRt,BkRtPer,BkRateType,BkDeleAdd,BkPartyGroup,BkExtraCount,BkNote,BkCbag,BkCWt," _
              & " BkCDate,BkTmpNo1,BkTmpNo2,BkTmpNo3,BkTmpChr1,BkTmpChr2,BkTmpChr3,BkSubBag,BkSubWt,BkSubDDate ) " & X
     
'     If GProcRstOpen(RsSA, X, "R") <= 0 Then
'       MsgBox "No Records..."
'       Exit Sub
'     Else
'        YY = "Vno in " & Y & " and Vtype='" & frmTypeStr & "' and VFirm = '" & gCCode & "' And VYear = " & gCYear
'        gCn.Execute "Delete from tblBookingSubTmp where  BkUserName='" & gUserName & "'"
'        gCn.Execute "Delete from tblBookingSubTmp where  " & YY & " "
'
'        gCn.Execute "Delete from tblBookingTmp where  BkUserName='" & gUserName & "'"
'        gCn.Execute "Delete from tblBookingTmp where  " & YY & " "
'
'
'        If GProcRstOpen(RsSA, XX, "R") > 0 Then
'            gCn.Execute " insert into tblBookingSubTmp (Vno,Vtype,Vdt,Vyear,VFirm,BkSubItSrNo,BkSubBag,BkSubWt,BkSubDDate,BkSubItCode,BkSubTmpNo1,BkSubTmpNo2,BkSubTmpNo3,BkSubTmpChr1,BkSubTmpChr2,BkSubTmpChr3,BkUserName  " _
'                        & " BkCDate,BkTmpNo1,BkTmpNo2,BkTmpNo3,BkTmpChr1,BkTmpChr2,BkTmpChr3 ) " & XX
'
'        End If
'     End If
     
     
     X = "select * from TmpBooking where BkUser='" & gUserName & "'"
     If GProcRstOpen(RsSA, X, "R") <= 0 Then
        MsgBox "No Records..."
        Exit Sub
     Else
        RsSA.MoveFirst
        Do While Not RsSA.EOF
           With cryReport
              .Formulas(10) = "wCSTNo = '" & gCCST & "'"
              .Formulas(11) = "wBSTNo = '" & gCBST & "'"
              .SelectionFormula = "{TmpBooking.Vno} = " & RsSA!Vno & " and {TmpBooking.VType}='" & RsSA!Vtype & "' and {TmpBooking.VFirm} = '" & RsSA!VFirm & "' and  {TmpBooking.VYear}=" & RsSA!VYear
              .ReportFileName = gReportPath & "rptBookingPrint_MillBill.rpt"
'              .ReportFileName = gReportPath & "xx.rpt"
              If Index = 0 Or Index = 1 Then
                 .Action = 1
'              Else
'                 '-- For Mail
'                 Dim Tattach As String
'                 Dim EmailId As String
'                 Dim MailToParty As String
'                 Dim Subject As String
'                 Dim Massage As String
'
''                 EmailId = "sanjivtavate@gmail.com"
''                 MailToParty = GProcGetColumnValue("tblMastAccount", "AcCode", RsSA!BkSCode, "N", "AcName", "S")
'                 Subject = "Order Confirmation From " + gCName
'                 Massage = "Dear Sir, <br> <br>"
'                 Massage = Massage + "<font color=red>Please see the attachment <br>"
'                 Massage = Massage + "<font size=3> <font color=blue> Line 2 -------------- "
'
'
'                 '------ Export To PDF
'                .SelectionFormula = "{tblBookingTmp.BkInvNo} = '" & RsSA!BkInvNo & "' and {tblBookingTmp.VType}='" & RsSA!Vtype & "' and {tblBookingTmp.VBranch} = '" & RsSA!VBranch & "' and {tblBookingTmp.BkFirm} = '" & RsSA!BkFirm & "' and  {tblBookingTmp.VYear}=" & RsSA!VYear
'
'                 Tattach = gReportPath & "rptBookingPrint.rpt"
''                 Tattach = gReportPath & "xx.rpt"
'
'                 Report_ExportToPDF frmMain.cryReport, gReportPath & "rptBookingPrint.rpt"
'
'                 If chkMailSupplier.Value = 1 Then
'                    MailToParty = GProcGetColumnValue("tblMastAccount", "AcCode", RsSA!BkSCode, "N", "AcName", "S")
'                    EmailId = GProcGetColumnValue("tblMastAccount", "AcCode", RsSA!BkSCode, "N", "AcEmail", "S")
'                    SendEmail frmMain.cryReport, Tattach, EmailId, MailToParty, Subject, Massage
'                 End If
'                 If chkMailBuyer.Value = 1 Then
'                    MailToParty = GProcGetColumnValue("tblMastAccount", "AcCode", RsSA!BkBCode, "N", "AcName", "S")
'                    EmailId = GProcGetColumnValue("tblMastAccount", "AcCode", RsSA!BkBCode, "N", "AcEmail", "S")
'                    SendEmail frmMain.cryReport, Tattach, EmailId, MailToParty, Subject, Massage
'                 End If
'                 Kill Mid(Tattach, 1, Len(Tattach) - 4) & ".pdf"
              End If
'              Exit Sub
              RsSA.MoveNext
           End With
        Loop
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
    FirstTimeFlag = True
    
    '--- TblAddLess
    frmTypeStr = gfrmTypeStr
    X = "Select * from " & tblName & " where Vtype='" & frmTypeStr & "' and VFirm  = '" & gCCode & "' and VYear=" & gCYear
    If Trim(OrderField) <> "" Then X = X & " order by " & OrderField
    Criteria = "vtype='" & frmTypeStr & "' and Vfirm = '" & gCCode & "' and  VYear=" & gCYear
'    GProcRstOpen rstbl, X, "O"
    If Trim(mShowRec) <> "" Then
       GProcRstOpen rstbl, "Select * from " & tblName & " where  " & Criteria & " and  " & IdField & " = " & CLng(Trim(mShowRec)), "R"
       ReadFields
       Call GProcSetButtons(Me, vbDataActionCancel)
       mShowRec = ""
    Else
       GProcRstOpen rstbl, "Select * from " & tblName & " where " & Criteria & " and  " & IdField & " = 0", "R"
       cmdBtn_Click (6)
    End If
    '---Grid
    SetGrid
    SetRelatedRS
    '---- Button Visiblity Set As Per User Previlage
    GProcSetButtonVisiblity Me
    cmdBtn_Click (6)
End Sub
Private Sub SetRelatedRS()
'Dim X As String
'Set rsOS = New Recordset
'Set rsVou = New Recordset
''---TblVoucher
'X = "Select * from tblVoucher where Vtype='" & frmTypeStr & "' and vno=" & txt(0) & " and VYear=" & gCYear & " order by VCtrNo"
'GProcRstOpen rsVou, X, "O"
''---TblOutStanding
'X = "Select * from tblOutStanding where Vtype='" & frmTypeStr & "' and vno=" & txt(0) & " and VYear=" & gCYear
'GProcRstOpen rsOS, X, "O"
End Sub
Private Sub SetGrid()
Dim X As String
Set rstblSub = New Recordset
    
X = "select BkDespItSrNo,BkDespBag,BkDespTmpChr2,BkDespTmpChr1" _
    & "  from tblBooKingDesp where " _
    & "  tblBooKingDesp.vtype='" & frmTypeStr & "' and tblBooKingDesp.vno=" & txt(0) & " and tblBooKingDesp.VFirm = '" & gCCode & "' and " _
    & "  tblBooKingDesp.VYear=" & gCYear & " order by BkSrNo"

GProcRstOpen rstblSub, X, "R"
With msGrid
   .FormatString = ">Sr.No.|>Bag   |<Lot No.                                    |<Recd"
    If FormAction = vbDataActionAddNew Then
        .rows = 2
        GridAddNew
    Else
        .rows = 1
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
                 .TextMatrix(.Row, i) = ""
            Case Else
                 .TextMatrix(.Row, i) = ""
        End Select
    Next i
    ReNumberGridSrno
End With
End Sub
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
If FormAction = vbDataActionUpdate Then
    MsgBox "You are about to modify the record. Save the record and then exit.", vbInformation + vbOKOnly, Me.Caption
    Cancel = True
    Exit Sub
End If
rstbl.Close
End Sub
Private Sub MSGrid_LostFocus()
CalAmount
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

If IsDate(mskDt(Index)) = False Then mskDt(Index) = "__/__/____"
End Sub
Private Sub txt_GotFocus(Index As Integer)
GProcSelectBox Me.ActiveControl
End Sub
Private Sub CheckForNumber()
GProcCheckForNumber Me, rstbl, MaxNo, pOtherTxtFromNo:=OtherTxtFromNo, pOtherTxtToNo:=OtherTxtToNo
End Sub
Private Sub txt_LostFocus(Index As Integer)
'If Index = 0  Then
'   GProcRstOpen rstbl, "Select * from " & tblName & " where " & IdField & " = " & Val(txt(0)) & " and " & Criteria, "R"
'   If rstbl.EOF = True Then
'      MsgBox "Entry No Not Found", vbCritical + vbOKOnly, Me.Caption
'      cmdBtn_Click (3) '--last record
'      ReadFields
'      cmdBtn_Click (14)
'      cmdBtn(6).SetFocus
'      Exit Sub
'   End If
'   ReadFields
'End If
CheckForNumber
SaveBtnEd
CalAmount
End Sub
Private Sub txt_KeyPress(Index As Integer, KeyAscii As Integer)
If KeyAscii = 13 Then Exit Sub
KeyAscii = GProcValidateKey(rstbl, Index, KeyAscii, txt(Index).text)
End Sub
'---------------- Grid
Private Sub msGrid_Click()
'On Error Resume Next
    msGrid_EnterCell
End Sub
Private Sub msGrid_EnterCell()
If FormAction <> vbDataActionClose Then
    CheckForNumberTxtGrid
    DoEvents
    txtGrid = ""
    DoEvents
    If msGrid.text <> "" Then txtGrid.text = msGrid.text
    '--- Lock TextBox
    '--- Check is Related Gate Pass
'    If IsRelatedItCtrlGatePass(msGrid.TextMatrix(msGrid.Row, 14)) = False Then
        With msGrid
            Select Case .Col
                Case 0: '---- Sr No
                    txtGrid.Locked = True
                    ReNumberGridSrno
                    txtGrid = ""
                    .Col = .Col + 1
'                Case 3: '---- Sh Date
'                    txtGrid.Locked = True
'                    txtGrid = ""
'                    .Col = .Col + 1
                Case Else
                    txtGrid.Locked = False
            End Select
        End With
'    Else '-Related Gate Pass
 '         txtGrid.Locked = True
  '  End If
End If
msGrid_KeyPress (0)
End Sub
Private Sub msGrid_KeyPress(KeyAscii As Integer)
Dim X As String
'--- Check is Related Gate Pass
'If IsRelatedItCtrlGatePass(msGrid.TextMatrix(msGrid.Row, 14)) = True Then Exit Sub
With msGrid
    If KeyAscii = 27 Then '----Esc Key
        KeyAscii = 0
        If ((Val(.TextMatrix(.Row, 1)) = 0)) And .rows > 2 Then
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
'If FormAction <> vbDataActionAddNew And FormAction <> vbDataActionUpdate Then Exit Sub
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
'---- Set TxtGrid
If msGrid.text <> "" Then txtGrid.text = msGrid.text
If msGrid.Col < rstblSub.Fields.Count Then
   txtGrid.Move msGrid.CellLeft + msGrid.Left, msGrid.CellTop + msGrid.Top, msGrid.CellWidth, msGrid.CellHeight
End If
txtGrid.Visible = True
txtGrid.Enabled = True
DoEvents
If KeyAscii <> 0 Then
    Select Case rstblSub.Fields(msGrid.Col).Type
        Case adChar, adVarWChar
'            If msGrid.Col = 1 Or msGrid.Col = 3 Or msGrid.Col = 4 Then '-- Count ,godown
 '               txtGrid_KeyPress (KeyAscii)
  '          Else
                txtGrid.text = txtGrid.text & Chr(KeyAscii)
   '         End If
        Case Else
            txtGrid.text = Chr(GProcValidateKey(rstblSub, msGrid.Col, KeyAscii, txtGrid.text)) 'Chr(GProcNumberWithDecimal(KeyAscii, txtGrid))
    End Select
End If
DoEvents
If msGrid.text <> "" Then
    txtGrid.SelStart = Len(txtGrid.text) + 1
Else
    txtGrid.SelStart = 1
End If
DoEvents
CalAmount
If TotBag > Val(txt(18)) Then
'   MsgBox "Check Sh.Bags Total.", vbCritical + vbOKOnly, Me.Caption
   msGrid.Col = 1
End If
If msGrid.TextMatrix(msGrid.Row, 0) > 0 And msGrid.Col = 1 Then
'    If IsDate(msGrid.TextMatrix(msGrid.Row, 3)) = False Then
'       X = "Select BkShDate from tblBooKing where vno=" & txt(0) & " and BkSrNo =  " & msGrid.TextMatrix(msGrid.Row, 0) & " And " & Criteria
'       GProcRstOpen Rs1, X, "R"
'       If Rs1.EOF Then
'          MsgBox ("Please Check Sr. No ...")
'          msGrid.Col = 0
'          Exit Sub
'       Else
'          msGrid.TextMatrix(msGrid.Row, 3) = Format(Rs1.Fields(0), "dd/MM/yyyy")
'       End If
'    End If
End If
If txtGrid.Visible = True Then txtGrid.SetFocus: Call GProcSelectBox(txtGrid)
End Sub
Private Sub msGrid_LeaveCell()
If txtGrid.Visible Then
    msGrid.text = txtGrid.text
    txtGrid.text = ""
    txtGrid.Visible = False
End If
DoEvents
CalGridAmt
End Sub
Private Sub CalGridAmt()
Dim MillCode As Long
Dim StdPack As Double
Dim TotBag As Long
Dim RowIndex As Integer
CheckForNumberGrid
Dim X As String
RowIndex = 1
With msGrid
    Do While RowIndex < .rows
        '--Tot Bag
        If Val(.TextMatrix(RowIndex, 1)) > 0 Then
            TotBag = TotBag + Val(.TextMatrix(RowIndex, 1))
        End If
        RowIndex = RowIndex + 1
    Loop
End With
'lblBags.Caption = TotBag
End Sub
Private Sub CheckForNumberGrid()
Dim RowIndex As Long
With rstblSub
For RowIndex = 1 To msGrid.rows - 1
    For i = 0 To .Fields.Count - 1
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
'                    msGrid.TextMatrix(RowIndex, i) = ""
                    msGrid.TextMatrix(RowIndex, i) = mskDt(2)
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
       ' Case adChar, adVarWChar '--String
       '     txtGrid.MaxLength = .Fields(msGrid.Col).DefinedSize
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
Private Sub txtCode_LostFocus(Index As Integer)
Dim X As String
Dim tmpRst As Recordset
If Trim(txtCode(Index)) = "" Then Exit Sub
If Index = 6 Then
   txttemp(6) = GProcGetColumnValue("tblMastAccount", "AcOurCode", UCase(txtCode(6)), "S", "AcName", "S")
   If Trim(txttemp(6)) = "" Then
      MsgBox ("Code Not Found ...")
      txttemp(6) = ""
      txttemp(6).Enabled = True
      txttemp(6).SetFocus
   Else
      txttemp(6).SetFocus
   End If
   Dim MilTpCd As Long
   MilTpCd = GProcGetColumnValue("tblMastNarration", "Narration", "Depot", "S", "Narrcode", "N")
   X = "Select Acname from tblMastAccount,tblMastBillSerial where tblMastaccount.Accode = tblMastBillSerial.MillCode  and tblMastBillSerial.Vfirm = '" & gCCode & "'   and  tblMastaccount.AcOurCode = '" & txtCode(6) & "'   and  tblMastBillSerial.SaleType = " & MilTpCd & "    "
   i = GProcRstOpen(tmpRst, X, "R")
   If i = 0 Then
      txttemp(6) = ""
      txtCode(6) = ""
      MsgBox ("Mill Not in Firm")
      txtCode(6).SetFocus
      Exit Sub
   End If
   tmpRst.Close
End If
End Sub
Private Sub txtGrid_GotFocus()
   txtGrid.Alignment = ProcSetAlignment(rstblSub, msGrid.Col)
    DoEvents
    Call GProcSelectBox(txtGrid)
End Sub
Private Function ProcSetAlignment(pRstbl As Recordset, pColIndex As Long)
Select Case pRstbl.Fields(pColIndex).Type
    Case adInteger, adDouble, 2, 3, 131, adDecimal
        ProcSetAlignment = 1
    Case Else
        ProcSetAlignment = 0
End Select
End Function
Private Sub txtGrid_KeyDown(KeyCode As Integer, Shift As Integer)
'--- Check is Related Gate Pass
'If IsRelatedItCtrlGatePass(msGrid.TextMatrix(msGrid.Row, 14)) = True Then Exit Sub
'--- Delete Row On Ctrl+Y
If KeyCode = vbKeyY And Shift = 2 And msGrid.rows > 2 Then
    msGrid = 1
    msGrid.RemoveItem (msGrid.Row)
    txtGrid.Visible = False
    txtGrid.text = ""
    ReNumberGridSrno
    CalAmount
    Exit Sub
End If
End Sub
'-- ReNumber Grid Sr No
Private Sub ReNumberGridSrno()
Dim i As Long
With msGrid
    For i = 1 To .rows - 1
        .TextMatrix(i, 0) = i
    Next i
End With
End Sub
Private Sub txtGrid_KeyPress(KeyAscii As Integer)
Dim LeftPos As Long
Dim TopPos As Long
Dim MillCode As Long
TopPos = txtGrid.Top + Me.Top + 650
LeftPos = txtGrid.Left + Me.Left
DoEvents
'--- Check is Related Gate Pass
'If IsRelatedItCtrlGatePass(msGrid.TextMatrix(msGrid.Row, 14)) = True Then Exit Sub
With msGrid
    If KeyAscii <> 13 And KeyAscii <> 27 And KeyAscii <> 25 Then
        '--- Search
        Select Case .Col
            Case 0 '--Booking
                KeyAscii = GProcValidateKey(rstblSub, msGrid.Col, KeyAscii, txtGrid)
            Case Else
                KeyAscii = GProcValidateKey(rstblSub, msGrid.Col, KeyAscii, txtGrid)
        End Select
   End If
    If KeyAscii = 13 Then
        CheckForNumberTxtGrid
    End If
    '--- Navigate Row ,Col
    If KeyAscii = 13 And .Col < 3 Then 'rstblSub.Fields.Count - 2 Then
       .Col = .Col + 1
    ElseIf KeyAscii = 13 And .Col = 3 Then '= rstblSub.Fields.Count - 2 Then
       If .Row = (.rows - 1) Then
          .rows = .rows + 1
          .Row = .Row + 1
           GridAddNew
        Else
           .Row = .Row + 1
        End If
        .Col = 0
    End If
    '-- On ESC
    If KeyAscii = 27 Then
        KeyAscii = 0
        If (Val(.TextMatrix(.Row, 1)) = 0) And .rows > 2 Then
            .RemoveItem (.Row)
            txtGrid.text = ""
        End If
        If TotBag > Val(txt(18)) Then
           MsgBox "Check Sh.Bags Total.", vbCritical + vbOKOnly, Me.Caption
           msGrid.Col = 1
           txtGrid.SetFocus
           Exit Sub
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
End Sub

'-- Grid Required
Private Function ValidateGrid() As Boolean
Dim TotBag As Long
Dim i As Long
Dim RowIndex As Long
Dim X As String
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
        If Val(.TextMatrix(RowIndex, 1)) > 0 Then
            TotBag = TotBag + Val(.TextMatrix(RowIndex, 1))
        End If
        For i = 0 To rstblSub.Fields.Count - 1
            Select Case i
'                Case 2 '-- Bag
'                    If CDbl(.TextMatrix(RowIndex, i)) = 0 Then
'                        MsgBox "Check Bag.", vbCritical + vbOKOnly, Me.Caption
'                        .Col = i
'                        msGrid.SetFocus
'                        Exit Function
'                    End If
'                Case 3 '-- Wt
'                    If CDbl(.TextMatrix(RowIndex, i)) = 0 Then
'                        MsgBox "Check Weight.", vbCritical + vbOKOnly, Me.Caption
'                        .Col = i
'                        msGrid.SetFocus
'                        Exit Function
'                    End If
           End Select
        Next i
      Next RowIndex
End With
If TotBag > Val(txt(18)) Then
   MsgBox "Check Despatch Detail Bag Total.", vbCritical + vbOKOnly, Me.Caption
   Exit Function
End If
ValidateGrid = True
ReNumberGridSrno
End Function
Private Sub SaveBtnEd()
'--- Vno,vdt,Party,A/c,Mill,Amount
If (CDbl(txt(0)) = 0 And FormAction <> vbDataActionAddNew) Then
    cmdBtn(13).Enabled = False
    cmdBtn(13).BackColor = &H80000004
Else
    cmdBtn(13).Enabled = True
    cmdBtn(13).BackColor = &HFFFF80
End If
End Sub
Private Sub FillTxtFromTemp()
txt(1) = frmTypeStr '--VType
If IsDate(mskDt(2)) Then '--Vdt
    txt(2) = CDate(mskDt(2))
Else
    txt(2) = ""
End If

'-- Party Code
txt(5) = GProcGetColumnValue("tblMastAccount", "AcName", "-", "S", "AcCode", "N")
'-- Mill Code
txt(6) = GProcGetColumnValue("tblMastAccount", "AcName", txttemp(6), "S", "AcCode", "N")
'--broker Code
txt(7) = GProcGetColumnValue("tblMastAccount", "AcName", "-", "S", "AcCode", "N")

'--Count
txt(17) = GProcGetColumnValue("tblMastItem", "ItName", txttemp(17), "S", "Itcode", "N", "ItMillCode = " & txt(6) & "")
txt(3) = gCYear
txt(4) = gCCode
End Sub
Private Sub FillTempFromTxt()
Dim X As String
Dim MsOms As Integer
If IsDate(txt(2)) Then '--VDt
    mskDt(2) = txt(2)
Else
    mskDt(2) = "__/__/____"
End If
'-- Party Code

'-- Mill Code
txtCode(6) = GProcGetColumnValue("tblMastAccount", "AcCode", txt(6), "N", "AcOurCode", "S")
txttemp(6) = GProcGetColumnValue("tblMastAccount", "AcCode", txt(6), "N", "AcName", "S")

'--Item
txttemp(17) = GProcGetColumnValue("tblMastItem", "Itcode", txt(17), "N", "ItName", "S", "ItMillCode = " & txt(6) & "")
'-- Fill Related Record
FillRelatedRecord
CalGridAmt
End Sub
Private Sub FillRelatedRecord()
Dim RowIndex As Long
Dim RowIndex2 As Long
Dim ColIndex As Long
Dim ColIndex2 As Long
'Dim TotBag As Long
'Dim TotWt As Double
'tmpAutoLRNo = "" ''-Auto LR No
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
                    Case 135
                        If Format(.Fields(ColIndex) & vbNullString, "DD/MM/YYYY") <> "01/01/1900" Then
                           msGrid.TextMatrix(RowIndex, ColIndex) = .Fields(ColIndex) & vbNullString
                        Else
                           msGrid.TextMatrix(RowIndex, ColIndex) = ""
                        End If
                    Case Else
                        msGrid.TextMatrix(RowIndex, ColIndex) = .Fields(ColIndex) & vbNullString
                End Select
'                If ColIndex = 9 Then TotBag = TotBag + CLng(msGrid.TextMatrix(RowIndex, 9)) '-- Tot Bag
'                If ColIndex = 10 Then TotWt = TotWt + CDbl(msGrid.TextMatrix(RowIndex, 10)) '-- Tot Wt
            Next ColIndex
            RowIndex = RowIndex + 1
            .MoveNext
            If Not .EOF Then msGrid.rows = msGrid.rows + 1
        Loop
    Else
        msGrid.rows = msGrid.rows + 1
        ReNumberGridSrno
    End If
End With


'--- Fill Related Control
SetRelatedRS
End Sub
Private Sub ReadFields()
FormAction = vbDataActionClose
GProcShowRecord Me, rstbl, MaxNo, OtherTxtFromNo, OtherTxtToNo
FillTempFromTxt
End Sub
Private Sub CalAmount()
Dim Amt As Double
Dim RowIndex As Integer
Dim MillCode As Long
Dim StdPack As Double
RowIndex = 1
TotBag = 0
'With msGrid
'    Do While RowIndex < .rows
'       '--Tot Bag
'       If Val(.TextMatrix(RowIndex, 2)) > 0 Then
'          TotBag = TotBag + Val(.TextMatrix(RowIndex, 2))
'       End If
'       RowIndex = RowIndex + 1
'    Loop
'End With
MillCode = GProcGetColumnValue("TblMastAccount", "AcName", txttemp(6), "S", "Accode", "N")
StdPack = GProcGetColumnValue("TblMastItem", "ITName", txttemp(17), "S", "ItStdpack", "N", " ItMillCode=" & MillCode)
'-- Wt=Bag * Std Pack
If StdPack > 0 Then
   txt(19) = Format(Val(txt(18)) * StdPack, WtStr)
End If
End Sub
'--- Is Related Record
Private Function IsRelatedRecord() As Boolean
Dim RsG As Recordset
Dim X As String
'-- Audited
'If GProcISAudited(txt(0), txt(1), txt(23)) = True Then
'    MsgBox "Record is Audited.You can not Modify/Delete this Record.", vbCritical + vbOKOnly, App.Title
'    IsRelatedRecord = True
'    Exit Function
'End If
''-- Is Gate Pass
'X = "Select Vno,Vdt from tblVouSub where Vtype='GP' and " _
'& " VSubPItCtrlNo in (select VSubItCtrlNo from tblVouSub where vtype='" & frmTypeStr & "' and vno=" & txt(0) & " and " _
'& " VYear=" & gCYear & ")"
'GProcRstOpen RsG, X, "R"
'If RsG.EOF Then
'Else
'    IsRelatedRecord = True
'    MsgBox "You can not Modify/Delete this record.Related Gate Pass No. " & RsG.Fields(0) & " is exists.", vbCritical + vbOKOnly, App.Title
'    Exit Function
'End If
'RsG.Close
End Function
'-- Is Related GatePass
Private Function IsRelatedGatePass() As Boolean
Dim RsG As Recordset
Dim X As String
'-- Is Gate Pass
'X = "Select Vno,Vdt from tblVouSub where Vtype='GP' and " _
'& " VSubPItCtrlNo in (select VSubItCtrlNo from tblVouSub where vtype='" & frmTypeStr & "' and vno=" & txt(0) & " and " _
'& " VYear=" & gCYear & ")"
'GProcRstOpen RsG, X, "R"
'If RsG.EOF Then
'Else
'    IsRelatedGatePass = True
'    MsgBox "You can not Modify/Delete this record.Related Gate Pass No. " & RsG.Fields(0) & " is exists.", vbCritical + vbOKOnly, App.Title
'    Exit Function
'End If
'RsG.Close
End Function
'-- Is Related It Ctrl GatePass
Private Function IsRelatedItCtrlGatePass(ItCtrlNo As Long) As Boolean
Dim RsG As Recordset
Dim X As String
'-- Is Gate Pass
'X = "Select Vno,Vdt from tblVouSub where Vtype='GP' and " _
'& " VSubPItCtrlNo = " & ItCtrlNo
'GProcRstOpen RsG, X, "R"
'If RsG.EOF Then
'Else
'    IsRelatedItCtrlGatePass = False
'    Exit Function
'End If
'RsG.Close
End Function
Private Sub txtTemp_KeyPress(Index As Integer, KeyAscii As Integer)
Dim LeftPos As Long
Dim TopPos As Long
Dim MillCode As Long
TopPos = txttemp(Index).Top + Me.Top + 650
LeftPos = txttemp(Index).Left + Me.Left
If KeyAscii = 13 Then Exit Sub
Select Case Index
Case 6  '-- Mill
     Dim MilTpCd As Long
     Dim ShowMill As String
     MilTpCd = GProcGetColumnValue("tblMastNarration", "Narration", "Depot", "S", "Narrcode", "N")
     ShowMill = "AgCode=90029 and   tblMastaccount.Accode = tblMastBillSerial.MillCode and  tblMastBillSerial.SaleType = " & MilTpCd & "  and tblMastBillSerial.Vfirm = '" & gCCode & "'  "
     gClsSearch.SearchMultiField "tblMastaccount,tblMastBillSerial", "BillSr,AcName,AcOurCode", Array("Mill Ini.", "Mill Name", "Code"), Array(1000, 8000, 1000), ShowMill, Trim(Chr(KeyAscii)), "AcName", LeftPos, TopPos, True
     KeyAscii = 0
     txtCode(6) = ""
     If gClsSearch.SearchMultiRetCol(1) <> "" Then
        txttemp(Index) = gClsSearch.SearchMultiRetCol(1)
        txtCode(6) = gClsSearch.SearchMultiRetCol(2)
     End If
Case 17  '-- Count
     MillCode = GProcGetColumnValue("TblMastAccount", "AcName", txttemp(6), "S", "Accode", "N")
     gClsSearch.SearchMultiField "tblMastItem,tblMastaccount", "ItName,AcName,ItStdpack,ItStdrateper,InDailyReport", Array("Count Name", "Mill", "Std. Pack", "Rate Per", ""), Array(5000, 5000, 1000, 1000, 0), "tblMastItem.ITMillCode=tblMastaccount.acCode and ITMillCode=" & MillCode, Chr(KeyAscii), "ItName", LeftPos, TopPos, True
     If gClsSearch.SearchMultiRetCol(4) = 1 Then
        MsgBox "This is No Use Count "
     End If
     KeyAscii = 0
     If gClsSearch.SearchMultiRetCol(0) <> "" And gClsSearch.SearchMultiRetCol(4) = 0 Then
        txttemp(Index) = gClsSearch.SearchMultiRetCol(0)
     End If
End Select
End Sub
