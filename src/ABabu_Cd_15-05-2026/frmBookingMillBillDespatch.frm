VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form frmBookingMillBillDespatch 
   BackColor       =   &H00FBF2E1&
   BorderStyle     =   0  'None
   Caption         =   "Purchase Entry"
   ClientHeight    =   9570
   ClientLeft      =   0
   ClientTop       =   15
   ClientWidth     =   13935
   FillColor       =   &H000040C0&
   FillStyle       =   0  'Solid
   Icon            =   "frmBookingMillBillDespatch.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   9570
   ScaleWidth      =   13935
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
      Index           =   29
      Left            =   18495
      TabIndex        =   69
      Text            =   "19"
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
      Index           =   28
      Left            =   18405
      TabIndex        =   68
      Text            =   "18"
      Top             =   3915
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
      Left            =   18405
      TabIndex        =   67
      Text            =   "27"
      Top             =   3510
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
      Left            =   18405
      TabIndex        =   66
      Text            =   "26"
      Top             =   3105
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
      Left            =   18405
      TabIndex        =   65
      Text            =   "25"
      Top             =   2700
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
      Left            =   18405
      TabIndex        =   64
      Text            =   "24"
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
      Index           =   23
      Left            =   18405
      TabIndex        =   63
      Text            =   "23"
      Top             =   1935
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
      Left            =   18405
      TabIndex        =   62
      Text            =   "22"
      Top             =   1530
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
      Left            =   18405
      TabIndex        =   61
      Text            =   "21"
      Top             =   1170
      Visible         =   0   'False
      Width           =   375
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
      Index           =   20
      Left            =   2025
      TabIndex        =   1
      Text            =   "20"
      Top             =   667
      Width           =   1500
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
      Left            =   18360
      TabIndex        =   60
      Text            =   "19"
      Top             =   450
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
      Index           =   18
      Left            =   18360
      TabIndex        =   59
      Text            =   "18"
      Top             =   90
      Visible         =   0   'False
      Width           =   375
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
      Left            =   4365
      TabIndex        =   48
      Top             =   4950
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
         TabIndex        =   54
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
         TabIndex        =   53
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
         TabIndex        =   52
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
         TabIndex        =   51
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
         TabIndex        =   50
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
         TabIndex        =   49
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
         TabIndex        =   56
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
         TabIndex        =   55
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
      Left            =   17775
      TabIndex        =   47
      Text            =   "10"
      Top             =   3645
      Visible         =   0   'False
      Width           =   375
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
      Index           =   6
      Left            =   1980
      TabIndex        =   3
      Text            =   "txttemp(6)"
      ToolTipText     =   "Press F3 for Creation"
      Top             =   1215
      Width           =   6225
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
      Index           =   13
      Left            =   1980
      TabIndex        =   4
      Text            =   "txttemp(13)"
      ToolTipText     =   "Press F3 for Creation"
      Top             =   1755
      Width           =   6270
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
      Left            =   17820
      TabIndex        =   43
      Text            =   "16"
      Top             =   5985
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
      Left            =   17820
      TabIndex        =   42
      Text            =   "15"
      Top             =   5625
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
      Left            =   17820
      TabIndex        =   41
      Text            =   "13"
      Top             =   4815
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
      Left            =   17820
      TabIndex        =   40
      Text            =   "12"
      Top             =   4365
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
      Height          =   375
      Index           =   9
      Left            =   17820
      TabIndex        =   39
      Text            =   "9"
      Top             =   3240
      Visible         =   0   'False
      Width           =   345
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
      Left            =   12180
      Style           =   1  'Graphical
      TabIndex        =   38
      Top             =   1365
      Width           =   1140
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
      Left            =   17820
      TabIndex        =   36
      Text            =   "14"
      Top             =   5265
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.CommandButton CmdBagDetails 
      Caption         =   "Cartoon &Bag Details"
      Height          =   720
      Left            =   16065
      TabIndex        =   35
      Top             =   2565
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
      Height          =   375
      Index           =   8
      Left            =   17775
      TabIndex        =   9
      Text            =   "8"
      Top             =   2880
      Visible         =   0   'False
      Width           =   345
   End
   Begin VB.CommandButton cmdLedger 
      Caption         =   "Ledger"
      Height          =   375
      Left            =   14985
      TabIndex        =   34
      Top             =   5040
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
      Left            =   540
      TabIndex        =   7
      Top             =   3960
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
      Left            =   17640
      TabIndex        =   32
      Text            =   "1"
      Top             =   450
      Visible         =   0   'False
      Width           =   420
   End
   Begin VB.TextBox txt 
      Alignment       =   1  'Right Justify
      Height          =   315
      Index           =   2
      Left            =   17595
      TabIndex        =   30
      Text            =   "2"
      Top             =   855
      Visible         =   0   'False
      Width           =   525
   End
   Begin VB.TextBox txt 
      Alignment       =   1  'Right Justify
      Height          =   315
      Index           =   5
      Left            =   17685
      TabIndex        =   29
      Text            =   "5"
      Top             =   1890
      Visible         =   0   'False
      Width           =   525
   End
   Begin VB.TextBox txt 
      Alignment       =   1  'Right Justify
      Height          =   315
      Index           =   3
      Left            =   17595
      TabIndex        =   28
      Text            =   "3"
      Top             =   1215
      Visible         =   0   'False
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
      Left            =   15435
      Style           =   1  'Graphical
      TabIndex        =   24
      Top             =   6435
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
      Left            =   15255
      Style           =   1  'Graphical
      TabIndex        =   23
      ToolTipText     =   "To Add New Record"
      Top             =   5850
      Visible         =   0   'False
      Width           =   1095
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
      Left            =   12180
      Style           =   1  'Graphical
      TabIndex        =   21
      ToolTipText     =   "Exit"
      Top             =   1845
      Width           =   1140
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
      Left            =   12180
      Style           =   1  'Graphical
      TabIndex        =   14
      ToolTipText     =   "To Save Record"
      Top             =   885
      Width           =   1140
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
      Left            =   16335
      Style           =   1  'Graphical
      TabIndex        =   22
      Top             =   4680
      Visible         =   0   'False
      Width           =   1140
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
      Left            =   10605
      Style           =   1  'Graphical
      TabIndex        =   13
      ToolTipText     =   "To Delete Record"
      Top             =   1845
      Width           =   1140
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
      Left            =   10605
      Style           =   1  'Graphical
      TabIndex        =   12
      ToolTipText     =   "To Modify Record"
      Top             =   1365
      Width           =   1140
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
      Left            =   10605
      Style           =   1  'Graphical
      TabIndex        =   11
      ToolTipText     =   "To Add New Record"
      Top             =   885
      Width           =   1140
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
      Height          =   400
      Index           =   5
      Left            =   16290
      Style           =   1  'Graphical
      TabIndex        =   15
      Top             =   3960
      Visible         =   0   'False
      Width           =   1140
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
      Left            =   8895
      Style           =   1  'Graphical
      TabIndex        =   16
      ToolTipText     =   "To Find Record"
      Top             =   1845
      Width           =   1140
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
      Left            =   9525
      Style           =   1  'Graphical
      TabIndex        =   20
      ToolTipText     =   "Last"
      Top             =   1380
      Width           =   510
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
      Left            =   8895
      Style           =   1  'Graphical
      TabIndex        =   19
      ToolTipText     =   "First"
      Top             =   1395
      Width           =   510
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
      Left            =   9525
      Style           =   1  'Graphical
      TabIndex        =   18
      ToolTipText     =   "Next"
      Top             =   900
      Width           =   510
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
      Left            =   17730
      TabIndex        =   0
      Text            =   "0"
      Top             =   90
      Visible         =   0   'False
      Width           =   255
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
      Left            =   8895
      Style           =   1  'Graphical
      TabIndex        =   17
      ToolTipText     =   "Previous"
      Top             =   900
      Width           =   510
   End
   Begin VB.TextBox txt 
      Alignment       =   1  'Right Justify
      Height          =   315
      Index           =   4
      Left            =   17685
      TabIndex        =   27
      Text            =   "4"
      Top             =   1575
      Visible         =   0   'False
      Width           =   525
   End
   Begin VB.TextBox txt 
      Height          =   315
      Index           =   6
      Left            =   17730
      TabIndex        =   26
      Text            =   "6"
      Top             =   2205
      Visible         =   0   'False
      Width           =   435
   End
   Begin VB.TextBox txt 
      Alignment       =   1  'Right Justify
      Height          =   315
      Index           =   7
      Left            =   17685
      TabIndex        =   25
      Text            =   "7"
      Top             =   2520
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
      Height          =   360
      Index           =   11
      Left            =   17820
      TabIndex        =   10
      Text            =   "11"
      Top             =   4005
      Visible         =   0   'False
      Width           =   375
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
      Index           =   17
      Left            =   1980
      MaxLength       =   25
      TabIndex        =   5
      Text            =   "17"
      Top             =   2295
      Width           =   6270
   End
   Begin MSMask.MaskEdBox mskDt 
      Height          =   375
      Index           =   10
      Left            =   6645
      TabIndex        =   2
      Top             =   675
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
      Bindings        =   "frmBookingMillBillDespatch.frx":058A
      Left            =   11070
      Top             =   45
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
      Height          =   5865
      Left            =   315
      TabIndex        =   6
      Top             =   3015
      Width           =   13335
      _ExtentX        =   23521
      _ExtentY        =   10345
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
      Caption         =   "Doc.  No.          :    "
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   345
      Left            =   315
      TabIndex        =   70
      Top             =   675
      Width           =   1635
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Vehicle No     :"
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
      Left            =   315
      TabIndex        =   58
      Top             =   2340
      Width           =   1590
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Transport        :"
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
      Left            =   315
      TabIndex        =   57
      Top             =   1755
      Width           =   1590
   End
   Begin VB.Label lblWeight 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Weight"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   375
      Left            =   11520
      TabIndex        =   46
      Top             =   8955
      Width           =   1500
   End
   Begin VB.Label lblBags 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Bags"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   375
      Left            =   10215
      TabIndex        =   45
      Top             =   8955
      Width           =   1365
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
      Left            =   315
      TabIndex        =   44
      Top             =   9135
      Width           =   1365
   End
   Begin VB.Shape Shape5 
      BorderColor     =   &H00EFB17A&
      BorderWidth     =   5
      FillColor       =   &H00EFB17A&
      Height          =   9540
      Left            =   0
      Top             =   0
      Width           =   13905
   End
   Begin VB.Label Label20 
      BackStyle       =   0  'Transparent
      Caption         =   "Despatch Date   :    "
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   345
      Left            =   4950
      TabIndex        =   37
      Top             =   675
      Width           =   1635
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "M&ill                   :"
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
      Left            =   315
      TabIndex        =   8
      Top             =   1215
      Width           =   1635
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
      Left            =   11070
      TabIndex        =   33
      Top             =   0
      Width           =   2475
   End
   Begin VB.Label Label6 
      BackColor       =   &H00EFB17A&
      Caption         =   "     Booking ( Mill Bill ) Despatch Details"
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
      TabIndex        =   31
      Top             =   45
      Width           =   13860
   End
   Begin VB.Shape Shape1 
      FillColor       =   &H00FED8FC&
      FillStyle       =   0  'Solid
      Height          =   2145
      Left            =   8685
      Shape           =   4  'Rounded Rectangle
      Top             =   540
      Width           =   4995
   End
End
Attribute VB_Name = "frmBookingMillBillDespatch"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Const tblName As String = "tblBookingSub"
Const IdField As String = "BkSubDespVno"
Const OrderField As String = "BkSubDespDt,BkSubDespVno"
Const MaxNo As Long = 21
Const OtherTxtFromNo As Long = 0 '-- Other Txt Ctrl From No
Const OtherTxtToNo As Long = 0 '-- Other Txt Ctrl Less No
Dim rstbl As Recordset '-- TblPurch
Dim rstblSub As Recordset '-- TblPurchSub
Dim varBookMark As Variant
Dim FirstTimeFlag As Boolean, Criteria As String
Dim OldgFrmTypeStr As String
Dim FormAction As Integer
Dim frmTypeStr As String
Dim MillCd As Long
Dim i As Long
Dim Mon As Integer
'Dim GridFlag As Boolean
Private Sub cmdBtn_Click(Index As Integer)
'On Error GoTo ErrorRoutine
Select Case Index
Case 0 'Previous
    If GProcRstOpen(rstbl, "Select * from " & tblName & " where BkSubDespVno = (Select max(BkSubDespVno) from " & tblName & " where BkSubDespVno <> 0  and BkSubDespVno < " & txt(20) & " and  " & Criteria & "  ) and  " & Criteria & "  ", "O") > 0 Then
       ReadFields
    Else
      MsgBox "This is First Record ", vbOKOnly + vbInformation
    End If
Case 1 'Next
    If GProcRstOpen(rstbl, "Select * from " & tblName & " where BkSubDespVno = (Select min(BkSubDespVno) from " & tblName & " where BkSubDespVno <> 0  and BkSubDespVno > " & txt(20) & "  and " & Criteria & ") and  " & Criteria & "   ", "O") > 0 Then
       ReadFields
    Else
       MsgBox "This is Last Record ", vbOKOnly + vbInformation
    End If
Case 2 'First
    If GProcRstOpen(rstbl, "Select * from " & tblName & " where BkSubDespVno = (Select min(BkSubDespVno) from " & tblName & " where BkSubDespVno <> 0 and " & Criteria & ") and  " & Criteria & "   ", "O") > 0 Then ReadFields
Case 3 'Last
     If GProcRstOpen(rstbl, "Select * from " & tblName & " where BkSubDespVno = (Select max(BkSubDespVno) from " & tblName & " where BkSubDespVno <> 0 and " & Criteria & " )and  " & Criteria & "   ", "O") > 0 Then ReadFields
Case 4 'Find
    gClsSearch.SearchMultiField "qryDespatchDetailHelp", "BkSubDespVno,BkSubDespDt,Mill,Transport,TruckNo", Array("Doc. No", "Despatch Dt", "Mill Name", "Transport", "Vehicle No"), Array(1000, 1500, 5500, 3000, 2000), Criteria, , , , , True
    If gClsSearch.SearchMultiRetCol(0) <> "" Then
        If GProcRstOpen(rstbl, "Select * from " & tblName & " where BkSubDespVno = (Select max(" & gClsSearch.SearchMultiRetCol(0) & ") from " & tblName & " where BkSubDespVno <> 0 and " & Criteria & " )and  " & Criteria & "   ", "O") > 0 Then ReadFields
    End If
Case 6 'Add
    lblBags.Caption = ""
    lblWeight.Caption = ""
    FormAction = vbDataActionAddNew
    Call GProcActivateControls(Me, True)
    SetControlEd
    Call GProcSetButtons(Me, vbDataActionAddNew)
    Call GProcClearForm(Me, rstbl, MaxNo, True, OtherTxtFromNo, OtherTxtToNo)
    If FirstTimeFlag = True Then
        FirstTimeFlag = False
    Else
        mskDt(10).SetFocus
    End If
    txt(20) = GProcGenerateId("TblBookingSub", "BkSubDespVno", "Vtype='" & frmTypeStr & "' and VFirm = '" & gCCode & "'   and BkSubDespVyear=" & gCYear)
    mskDt(10) = Format(Now, "dd/MM/yyyy")
    SetGrid
Case 7 'Modify
        If IsRelatedRecord = False Then
            ReadFields
            FormAction = vbDataActionUpdate
            Call GProcActivateControls(Me, True)
            SetControlEd
            Call GProcSetButtons(Me, vbDataActionUpdate)
            txt(0).Enabled = False
            mskDt(10).SetFocus
        End If
    'End If
'Case 8 'Delete
'    'If mURecDel = True Then
'      If IsRelatedGatePass = True Then
'            MsgBox "You can not Modify/Delete this record.Related Gate Pass is exists.", vbCritical + vbOKOnly, App.Title
'            Exit Sub
'      End If
'      If IsRelatedRecord = False Then
'        If MsgBox("Do You Really Delete This Record?", vbQuestion + vbYesNo + vbDefaultButton1, "Confirm Delete") = vbYes Then
'            If rstbl.RecordCount > 1 Then
'                 '-- Delete Related Record
'                DeleteRelatedRecord
'                rstbl.Delete
''                 '--- Bag Details
''                gCn.Execute "delete From tblBags where InwNo=" & txt(0) & " and InwType='" & frmTypeStr & "'"
'                cmdBtn_Click (1)
'            ElseIf rstbl.RecordCount = 1 Then
'                '-- Delete Related Record
'                DeleteRelatedRecord
'                rstbl.Delete
'                '--- Bag Details
''                gCn.Execute "delete From tblBags where InwNo=" & txt(0) & " and InwType='" & frmTypeStr & "'"
'                Call GProcActivateControls(Me, False)
'                Call GProcSetButtons(Me, vbDataActionCancel)
'                Form_Load
'            End If
'        End If
'    End If
'Case 9 'Print
'    Frame1.Caption = "Booking Confirmation Printing "
'        '-- Inv No
'    txtPInvno(0) = txt(0).text
'    txtPInvno(1) = txt(0).text
'    Frame1.Visible = True
'    DoEvents
'    txtPInvno(0).Enabled = True
'    txtPInvno(1).Enabled = True
'    txtPInvno(0).SetFocus
Case 10 'Exit
'    DoEvents
    Unload Me
    Exit Sub
Case 11 'Label
    'frmLabel.Show 1
Case 12 'Optional
Case 13 'Save
        'txt(51) = CDate(Date)
        If ValidateData = True Then
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
'        GProcSaveRecord Me, rstbl, FormAction, MaxNo, OtherTxtFromNo, OtherTxtToNo
        SaveRelatedRecord
        Call GProcActivateControls(Me, False)
        Call GProcSetButtons(Me, True)
        FormAction = vbDataActionClose
        msGrid.Col = 0 '-----
        msGrid.Row = 1 '-----
        txtGrid.Move msGrid.CellLeft + msGrid.Left, msGrid.CellTop + msGrid.Top, msGrid.CellWidth, msGrid.CellHeight '-----
        If FormAction = vbDataActionAddNew Then
            If cmdBtn(6).Visible = True Then cmdBtn(6).SetFocus
        Else
            If cmdBtn(7).Visible = True Then cmdBtn(7).SetFocus
        End If
    End If
Case 14 'Cancel
    txtGrid.Visible = False
    If rstbl.RecordCount <= 0 Then
        Call GProcClearForm(Me, rstbl, MaxNo, True, OtherTxtFromNo, OtherTxtToNo)
        Call GProcActivateControls(Me, False)
        Call GProcSetButtons(Me, vbDataActionCancel)
    Else
        If FormAction = vbDataActionAddNew Then
            rstbl.MoveLast
        Else
            rstbl.CancelUpdate
        End If
        Call GProcClearForm(Me, rstbl, MaxNo, True, OtherTxtFromNo, OtherTxtToNo)
        ReadFields
        Call GProcActivateControls(Me, False)
        Call GProcSetButtons(Me, True)
    End If
    FormAction = vbDataActionCancel
    If msGrid.Enabled = True Then
       msGrid.Col = 0
       msGrid.Row = 1
       txtGrid.Move msGrid.CellLeft + msGrid.Left, msGrid.CellTop + msGrid.Top, msGrid.CellWidth, msGrid.CellHeight
    End If
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
Private Sub SetControlEd()
If FormAction = vbDataActionAddNew Then
   txt(20).Enabled = False
End If
End Sub
Private Sub DeleteRelatedRecord()
'''--- Voucher Sub
'gCn.Execute "delete from tblBookingSub where vtype='" & frmTypeStr & "' and vno=" & txt(0) & " and Vfirm = '" & gCCode & "'  and " _
'& " VYear=" & gCYear

''-- OutStanding
'gCn.Execute "delete from tblOutStanding where vtype='" & frmTypeStr & "' and vno=" & txt(0) & " and " _
'& " VYear=" & gCYear
''-- Voucher
'gCn.Execute "delete from tblVoucher where vtype='" & frmTypeStr & "' and vno=" & txt(0) & " and " _
'& " VYear=" & gCYear
'--- JV
'gCn.Execute "delete from tblVoucher where vtype='JV' and vno=" & txt(20) & " and " _
'& " VYear=" & gCYear
End Sub
Private Sub SaveRelatedRecord()
Dim i As Long, X As String
Dim RowIndex As Integer
RowIndex = 1
With msGrid
    Do While RowIndex < .rows
       If (Val(.TextMatrix(RowIndex, 6)) = 0) And .rows > 2 Then
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
        If CDec(Val(.TextMatrix(i, 6))) > 0 Then
           gCn.Execute "update tblBooKingSub set BkSubDespBag = " & Val(.TextMatrix(i, 6)) & ", BkSubDespWt = " & Val(.TextMatrix(i, 7)) & ",BkSubDespDt = '" & IIf(IsDate(txt(10)), Format(txt(10), "yyyy/mm/dd"), "") & "', BkSubDespVno = " & Val(txt(20)) & " ,BkSubTmpChr1 = '" & txt(17) & "' ,BkSubTrnpt = '" & txt(13) & "',BkSubDespVyear = " & gCYear & " " _
                       & " where VNo = " & Val(.TextMatrix(i, 0)) & " and  VType = '" & UCase(.TextMatrix(i, 8)) & "' and  VFirm = '" & .TextMatrix(i, 10) & "'     and VYear = " & Val(.TextMatrix(i, 9)) & "  and BkSubItSrNo = " & Val(.TextMatrix(i, 11)) & " "
            DoEvents
        End If
       i = i + 1
     Loop
End With
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
    '--- TblAddLess
    frmTypeStr = gfrmTypeStr
    X = "Select * from tblBookingSub,tblMastNarration where Vtype='" & frmTypeStr & "' and VFirm  = '" & gCCode & "' and tblBookingSub.BkSubTrnpt = tblMastNarration.NarrCode and BkSubDespVyear=" & gCYear
    If Trim(OrderField) <> "" Then X = X & " order by " & OrderField
    Criteria = "vtype='" & frmTypeStr & "' and Vfirm = '" & gCCode & "' and  BkSubDespVyear=" & gCYear
    GProcRstOpen rstbl, X, "O"
    FirstTimeFlag = True
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
X = "select tblBooKingSub.Vno,tblBooKingSub.Vdt,tblMastItem.Itname,tblBooKingSub.BkSubDDate,tblBooKingSub.BkSubBag,tblBooKingSub.BkSubWt,tblBooKingSub.BkSubDespBag,tblBooKingSub.BkSubDespWt,VType,VYear,VFirm,BkSubItSrNo" _
& "  from tblBookingSub,tblMastItem where " _
& "  tblBookingSub.vtype='" & frmTypeStr & "' and tblBookingSub.BkSubDespVno=" & txt(20) & " and tblBookingSub.VFirm = '" & gCCode & "' and tblBookingSub.BkSubItCode = tblMastItem.Itcode  and tblBookingSub.BkSubDespVyear=" & gCYear & " order by BkSubDDate"
GProcRstOpen rstblSub, X, "R"
With msGrid
   .FormatString = ">Booking No|<Bookung Date|<Count Name                                         |<Sh.Date      |>Sh.Bag |>Sh.Weight  |>Recd.Bag|>Recd.Weight||||"
    .ColWidth(8) = 0
    .ColWidth(9) = 0
    .ColWidth(10) = 0
    .ColWidth(11) = 0
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
End With
End Sub
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
If FormAction = vbDataActionUpdate Then
    MsgBox "You are about to modify the record. Save the record and then exit.", vbInformation + vbOKOnly, Me.Caption
    Cancel = True
    Exit Sub
End If
'rstbl.Close
End Sub
Private Sub mskDt_GotFocus(Index As Integer)
GProcSelectBox Me.ActiveControl
End Sub
Private Sub mskDt_LostFocus(Index As Integer)
If Index = 10 Then
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
'GProcCheckForNumber Me, rstbl, MaxNo, pOtherTxtFromNo:=OtherTxtFromNo, pOtherTxtToNo:=OtherTxtToNo
End Sub
Private Sub txt_LostFocus(Index As Integer)
CheckForNumber
SaveBtnEd
CalAmount

End Sub
Private Sub txt_KeyPress(Index As Integer, KeyAscii As Integer)
If KeyAscii = 13 Then Exit Sub
KeyAscii = GProcValidateKey(rstbl, Index, KeyAscii, txt(Index).text)
End Sub
Private Sub txttemp_GotFocus(Index As Integer)
GProcSelectBox Me.ActiveControl
End Sub
Private Sub txtTemp_KeyPress(Index As Integer, KeyAscii As Integer)
Dim LeftPos As Long
Dim TopPos As Long
Dim AgCode As Long
Dim wPartyCode As Long
Dim BrkAcCode As Long
Dim MillCode As Long
TopPos = txttemp(Index).Top + Me.Top + 650
LeftPos = txttemp(Index).Left + Me.Left
If KeyAscii = 13 Then Exit Sub
Select Case Index
    Case 6  '-- Mill  Alias Account
        gClsSearch.SearchMultiField "tblMastaccount", "AcName,AcCity", Array("Mill Name", "City"), Array(7000, 2000), "(AgCode=90029  and AcMillBillFirm = '" & gCCode & "') ", Chr(KeyAscii), "AcName", LeftPos, TopPos, True
        KeyAscii = 0
        If gClsSearch.SearchMultiRetCol(0) <> "" Then
            txttemp(Index) = gClsSearch.SearchMultiRetCol(0)
        End If
    Case 13  '-- Transport
         gClsSearch.SearchMultiField "tblMastNarration", "Narration", Array("Transport"), Array(txttemp(Index).Width), " NarrType='R'", Chr(KeyAscii), "Narration", 0, 0, True
         KeyAscii = 0
         If gClsSearch.SearchMultiRetCol(0) <> "" Then
             txttemp(Index) = gClsSearch.SearchMultiRetCol(0)
         End If
End Select
End Sub
Private Sub txtTemp_LostFocus(Index As Integer)
SaveBtnEd
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
        With msGrid
            Select Case .Col
                Case 1, 2, 3, 4, 5: '---- Sr No
                    txtGrid.Locked = True
                Case Else
                    txtGrid.Locked = False
            End Select
        End With
End If
msGrid_KeyPress (0)
End Sub
Private Sub msGrid_KeyPress(KeyAscii As Integer)
'--- Check is Related Gate Pass
'If IsRelatedItCtrlGatePass(msGrid.TextMatrix(msGrid.Row, 14)) = True Then Exit Sub
With msGrid
    If KeyAscii = 27 Then '----Esc Key
        KeyAscii = 0
        If ((Val(.TextMatrix(.Row, 6)) = 0)) And .rows > 2 Then
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
        If .Col = 7 Then 'rstblSub.Fields.Count - 2 Then
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
   
With msGrid
    If KeyAscii <> 13 And KeyAscii <> 27 And .Col = 1 And Val(.TextMatrix(.Row, 0)) > 0 Then
        DoEvents
        KeyAscii = 0
        Pending_Despatch (.TextMatrix(.Row, 0))
        Exit Sub
        KeyAscii = 0
    End If
End With





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
If txtGrid.Visible = True Then txtGrid.SetFocus
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
Dim TotWt As Double
Dim RowIndex As Integer
CheckForNumberGrid
RowIndex = 1
With msGrid
    Do While RowIndex < .rows
        '--Tot Bag
        If Val(.TextMatrix(RowIndex, 6)) > 0 Then
            TotBag = TotBag + Val(.TextMatrix(RowIndex, 6))
        End If
        '--Tot Wt
        If Val(.TextMatrix(RowIndex, 7)) > 0 Then
            TotWt = TotWt + Val(.TextMatrix(RowIndex, 7))
        End If
        RowIndex = RowIndex + 1
    Loop
End With
lblBags.Caption = TotBag
lblWeight.Caption = TotWt
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
    CalAmount
    Exit Sub
End If
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
            Case 1 '--Booking
                KeyAscii = GProcValidateKey(rstblSub, msGrid.Col, KeyAscii, txtGrid)
            Case Else
                KeyAscii = GProcValidateKey(rstblSub, msGrid.Col, KeyAscii, txtGrid)
        End Select
   End If
    If KeyAscii = 13 Then
        CheckForNumberTxtGrid
    End If
    '--- Navigate Row ,Col
    If KeyAscii = 13 And .Col < 7 Then 'rstblSub.Fields.Count - 2 Then
        .Col = .Col + 1
    ElseIf KeyAscii = 13 And .Col = 7 Then '= rstblSub.Fields.Count - 2 Then
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
        If (Val(.TextMatrix(.Row, 6)) = 0) And .rows > 2 Then
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
End Sub
'-- Grid Required
Private Function ValidateGrid() As Boolean
'Dim TotBag As Long
'Dim TotWt As Double
'Dim i As Long
'Dim RowIndex As Long
'Dim X As String
'Dim Rs1 As Recordset
'Dim oldFrmTypeStr As String
'If Trim(txtGrid.text) <> "" Then
'    msGrid.text = txtGrid.text
'    CalGridAmt
'End If
'txtGrid.Visible = False
'txtGrid.text = ""
'CheckForNumberGrid
'With msGrid
'Validate:
'    For RowIndex = 1 To .rows - 1
'        If Val(.TextMatrix(RowIndex, 2)) > 0 Then
'            TotBag = TotBag + Val(.TextMatrix(RowIndex, 2))
'        End If
'        '--Tot Wt
'        If Val(.TextMatrix(RowIndex, 3)) > 0 Then
'            TotWt = TotWt + Val(.TextMatrix(RowIndex, 3))
'        End If
'        For i = 0 To rstblSub.Fields.Count - 1
'            Select Case i
'                Case 1 '-- Date
'                    If CDbl(.TextMatrix(RowIndex, 2)) > 0 And IsDate(.TextMatrix(RowIndex, i)) = False Then
'                        MsgBox "Check Date.", vbCritical + vbOKOnly, Me.Caption
'                        .Col = i
'                        msGrid.SetFocus
'                        Exit Function
'                    End If
'                    If CDbl(.TextMatrix(RowIndex, 2)) > 0 And IsDate(.TextMatrix(RowIndex, i)) = True Then
'                        MsgBox "Check Date.", vbCritical + vbOKOnly, Me.Caption
'                        .Col = i
'                        msGrid.SetFocus
'                        Exit Function
'                    End If
'                    If IsDate(.TextMatrix(RowIndex, i)) = True Then
'                       If CDate(.TextMatrix(RowIndex, i)) < mskDt(2) < mskDt(2) Then
'                           MsgBox "Delivery Date Must Be Less Than Booking Date.", vbCritical + vbOKOnly, Me.Caption
'                           .Col = i
'                           msGrid.SetFocus
'                        Exit Function
'                       End If
'                    End If
'
''                Case 2 '-- Bag
''                    If CDbl(.TextMatrix(RowIndex, i)) = 0 Then
''                        MsgBox "Check Bag.", vbCritical + vbOKOnly, Me.Caption
''                        .Col = i
''                        msGrid.SetFocus
''                        Exit Function
''                    End If
''                Case 3 '-- Wt
''                    If CDbl(.TextMatrix(RowIndex, i)) = 0 Then
''                        MsgBox "Check Weight.", vbCritical + vbOKOnly, Me.Caption
''                        .Col = i
''                        msGrid.SetFocus
''                        Exit Function
''                    End If
'           End Select
'        Next i
'      Next RowIndex
'End With
'If TotBag > Val(txt(18)) Then
'   MsgBox "Check Despatch Detail Bag Total.", vbCritical + vbOKOnly, Me.Caption
'   Exit Function
'End If
'If TotWt > Val(txt(19)) Then
'   MsgBox "Check Despatch Detail Weight Total.", vbCritical + vbOKOnly, Me.Caption
'   Exit Function
'End If
ValidateGrid = True
End Function
Private Sub SaveBtnEd()
'--- Vno,vdt,Party,A/c,Mill,Amount
If (CDbl(txt(20)) = 0 And FormAction <> vbDataActionAddNew) Or GProcIsDateValid(mskDt(10)) = False Or txttemp(6) = "" Or txttemp(13) = "" Or txt(17) = "" Then
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
Dim rstAddCheck As Recordset
Dim rstModiCheck As Recordset
Dim ExpAcCode As Long
'    MillCtrlEd
    '--- Validate Grid
    If ValidateGrid = False Then
        'MSGrid.SetFocus
        Exit Function
    End If
    '--- Vno
    If CLng(txt(20)) = 0 And FormAction <> vbDataActionAddNew Then
        MsgBox "Check Doc No.", vbInformation + vbOKOnly, Me.Caption
        txt(20).SetFocus
        Exit Function
    End If
    '--- Vdt
    If GProcIsDateValid(mskDt(10)) = False Then
        MsgBox "Check Despatch Date.", vbInformation + vbOKOnly, Me.Caption
        mskDt(10).SetFocus
        Exit Function
    End If
    '---- Mill Account
    If txttemp(6) = "" Then
        MsgBox "Check Mill Name", vbInformation + vbOKOnly, Me.Caption
        txttemp(6).SetFocus
        Exit Function
    End If
    '--- Transport
    If txttemp(13) = "" Then
        MsgBox "Check Transport Name", vbInformation + vbOKOnly, Me.Caption
        txttemp(13).SetFocus
        Exit Function
    End If
    '--- Transport
    If txt(17) = "" Then
        MsgBox "Check Vehicle No", vbInformation + vbOKOnly, Me.Caption
        txt(17).SetFocus
        Exit Function
    End If
    '----Duplicate Records
    '----Addition Case
    If FormAction = vbDataActionAddNew Then
        X = "Select BkSubDespVno From " & tblName & " Where BkSubDespVno=" & CLng(txt(20)) & " and Vtype='" & frmTypeStr & "'  and Vfirm = '" & gCCode & "'  and BkSubDespVyear=" & gCYear
        i = GProcRstOpen(rstAddCheck, X, "R")
        If i > 0 Then
            MsgBox "Record is Doc. No. " & CLng(txt(20)), vbCritical + vbOKOnly, "Duplicate Record"
            txt(20).SetFocus
            Exit Function
        End If
        rstAddCheck.Close
    End If
    ValidateData = True
End Function
Private Sub FillTxtFromTemp()
txt(1) = frmTypeStr '--VType
If IsDate(mskDt(10)) Then '--Vdt
    txt(10) = CDate(mskDt(10))
Else
    txt(10) = ""
End If
'--Transport
If txttemp(13) = "" Then
   txttemp(13) = "-"
End If
txt(13) = GProcGetColumnValue("tblMastNarration", "Narration", txttemp(13), "S", "NarrCode", "N")
txt(3) = gCYear
txt(4) = gCCode
txt(21) = gCYear
End Sub
Private Sub FillTempFromTxt()
Dim X As String
Dim MilCd As Long
Dim RsG As Recordset


If IsDate(txt(10)) Then '--VDt
    mskDt(10) = txt(10)
Else
    mskDt(10) = "__/__/____"
End If
'-- Mill Code
X = "select BkMillCode,Acname from tblBooking,tblMastAccount where vno = " & rstbl.Fields("Vno") & " and vtype = '" & rstbl.Fields("Vtype") & "' and vyear = " & rstbl.Fields("Vyear") & " and vfirm = '" & rstbl.Fields("Vfirm") & "' and tblBooking.BkMillCode = tblMastAccount.AcCode"

GProcRstOpen RsG, X, "R"
If RsG.EOF Then
   txttemp(6) = "-"
Else
   txttemp(6) = RsG.Fields(1)
End If



txttemp(13) = GProcGetColumnValue("tblMastNarration", "NarrCode", rstbl.Fields("BkSubTrnpt"), "N", "Narration", "S")

'-- Fill Related Record
FillRelatedRecord
CalGridAmt
End Sub
Private Sub FillRelatedRecord()
Dim RowIndex As Long
Dim ColIndex As Long
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
Dim TotBag As Long
Dim TotWt As Double
'---- Sub amt
CheckForNumberGrid
CheckForNumber
RowIndex = 1
With msGrid
    Do While RowIndex < .rows
        '--Tot Bag
        If Val(.TextMatrix(RowIndex, 6)) > 0 Then
            TotBag = TotBag + Val(.TextMatrix(RowIndex, 6))
        End If
        '--Tot Wt
        If Val(.TextMatrix(RowIndex, 7)) > 0 Then
            TotWt = TotWt + Val(.TextMatrix(RowIndex, 7))
        End If
        RowIndex = RowIndex + 1
    Loop
End With
lblBags.Caption = TotBag
lblWeight.Caption = TotWt
CheckForNumber
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

Private Sub Pending_Despatch(BkNo As Double)
Dim X As String
Dim RsG As Recordset
MillCd = GProcGetColumnValue("tblMastAccount", "AcName", txttemp(6), "S", "Accode", "N")
If Trim(txttemp(6)) <> "" Then
    X = "select tblBooKingSub.Vno,tblBooKingSub.Vdt,tblMastItem.Itname,tblBooKingSub.BkSubDDate,tblBooKingSub.BkSubBag,tblBooKingSub.BkSubWt,tblBooKingSub.VType,tblBooKingSub.VYear,tblBooKingSub.VFirm,tblBooKingSub.BkSubItSrNo" _
    & "  from tblBookingSub,tblMastItem,tblBooking where " _
    & "  tblBooking.vno = tblBookingSub.vno and tblBooking.vtype = tblBookingSub.vtype and tblBooking.vyear = tblBookingSub.vyear and  tblBooking.vfirm = tblBookingSub.vfirm and tblBooking.BkMillCode = " & MillCd & " and " _
    & "  tblBookingSub.vtype='" & frmTypeStr & "' and tblBookingSub.Vno=" & BkNo & " and tblBookingSub.VFirm = '" & gCCode & "' and tblBookingSub.vyear<=" & gCYear & "  and tblBooKingSub.BkSubItCode = tblMastItem.ItCode order by BkSubDDate"
    GProcRstOpen RsG, X, "R"
    If RsG.RecordCount = 1 Then
        With msGrid
               .TextMatrix(.Row, 1) = Format(RsG.Fields(1), "dd/MM/yyyy")
               .TextMatrix(.Row, 2) = RsG.Fields(2)
               .TextMatrix(.Row, 3) = Format(RsG.Fields(3), "dd/MM/yyyy")
               .TextMatrix(.Row, 4) = Format(RsG.Fields(4), "#######0")
               .TextMatrix(.Row, 5) = Format(RsG.Fields(5), "#######0.000")
               .TextMatrix(.Row, 6) = Format(RsG.Fields(4), "#######0")
               .TextMatrix(.Row, 7) = Format(RsG.Fields(5), "#######0.000")
               .TextMatrix(.Row, 8) = RsG.Fields(6)
               .TextMatrix(.Row, 9) = RsG.Fields(7)
               .TextMatrix(.Row, 10) = RsG.Fields(8)
               .TextMatrix(.Row, 11) = RsG.Fields(9)
               .Col = 6
        End With
    Else
        gClsSearch.SearchMultiField X, "", Array("Booking No", "Date", "Count Name", "Sh.Date", "Sh.Bag", "Sh.Weight", "", "", "", ""), Array(1200, 1500, 5000, 1500, 800, 1700, 0, 0, 0, 0), "", "", "", 0, 0, False, , True
        With msGrid
             If gClsSearch.SearchMultiRetCol(0) <> "" Then
               .TextMatrix(.Row, 1) = Format(gClsSearch.SearchMultiRetCol(1), "dd/MM/yyyy")
               .TextMatrix(.Row, 2) = gClsSearch.SearchMultiRetCol(2)
               .TextMatrix(.Row, 3) = Format(gClsSearch.SearchMultiRetCol(3), "dd/MM/yyyy")
               .TextMatrix(.Row, 4) = Format(gClsSearch.SearchMultiRetCol(4), "#######0")
               .TextMatrix(.Row, 5) = Format(gClsSearch.SearchMultiRetCol(5), "#######0.000")
               .TextMatrix(.Row, 6) = Format(gClsSearch.SearchMultiRetCol(4), "#######0")
               .TextMatrix(.Row, 7) = Format(gClsSearch.SearchMultiRetCol(5), "#######0.000")
               .TextMatrix(.Row, 8) = gClsSearch.SearchMultiRetCol(6)
               .TextMatrix(.Row, 9) = gClsSearch.SearchMultiRetCol(7)
               .TextMatrix(.Row, 10) = gClsSearch.SearchMultiRetCol(8)
               .TextMatrix(.Row, 11) = gClsSearch.SearchMultiRetCol(9)
               .Col = 6
            End If
        End With
    End If
End If
End Sub

