VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Begin VB.Form frmJVGST 
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
   Icon            =   "frmJVGST.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   9180
   ScaleWidth      =   14310
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox txt 
      Height          =   345
      Index           =   30
      Left            =   16155
      TabIndex        =   81
      Text            =   "30"
      Top             =   1935
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.TextBox txt 
      Height          =   345
      Index           =   29
      Left            =   16200
      TabIndex        =   80
      Text            =   "29"
      Top             =   1485
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.TextBox txt 
      Height          =   345
      Index           =   28
      Left            =   16245
      TabIndex        =   79
      Text            =   "28"
      Top             =   1080
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
      Height          =   2295
      Left            =   3000
      TabIndex        =   55
      Top             =   3960
      Visible         =   0   'False
      Width           =   5085
      Begin VB.TextBox txtno 
         Appearance      =   0  'Flat
         Height          =   360
         Index           =   0
         Left            =   1950
         MaxLength       =   7
         TabIndex        =   60
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
         TabIndex        =   59
         Top             =   1560
         Width           =   1335
      End
      Begin VB.TextBox txtno 
         Appearance      =   0  'Flat
         Height          =   360
         Index           =   1
         Left            =   1950
         MaxLength       =   7
         TabIndex        =   58
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
         TabIndex        =   57
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
         TabIndex        =   56
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
         TabIndex        =   62
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
         TabIndex        =   61
         Top             =   1080
         Width           =   1050
      End
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
      Height          =   345
      Index           =   27
      Left            =   5940
      TabIndex        =   13
      Text            =   "27"
      Top             =   5310
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
      Left            =   2745
      TabIndex        =   12
      Text            =   "26"
      Top             =   5310
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
      Height          =   345
      Index           =   25
      Left            =   5940
      TabIndex        =   11
      Text            =   "25"
      Top             =   4905
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
      Left            =   2745
      TabIndex        =   10
      Text            =   "24"
      Top             =   4905
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
      Height          =   345
      Index           =   23
      Left            =   5940
      TabIndex        =   9
      Text            =   "23"
      Top             =   4500
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
      Index           =   22
      Left            =   2745
      TabIndex        =   8
      Text            =   "22"
      Top             =   4500
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
      Height          =   345
      Index           =   21
      Left            =   5940
      TabIndex        =   7
      Text            =   "21"
      Top             =   3870
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
      TabIndex        =   19
      Top             =   8280
      Visible         =   0   'False
      Width           =   885
   End
   Begin VB.TextBox txt 
      Height          =   345
      Index           =   20
      Left            =   15030
      TabIndex        =   54
      Text            =   "20"
      Top             =   6885
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.TextBox txt 
      Height          =   345
      Index           =   19
      Left            =   15075
      TabIndex        =   53
      Text            =   "19"
      Top             =   7380
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.TextBox txt 
      Height          =   345
      Index           =   18
      Left            =   15075
      TabIndex        =   52
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
      TabIndex        =   14
      Text            =   "txttemp(20)"
      ToolTipText     =   "Press F3 for Creation"
      Top             =   7185
      Width           =   7290
   End
   Begin VB.TextBox txt 
      Height          =   345
      Index           =   17
      Left            =   15030
      TabIndex        =   51
      Text            =   "17"
      Top             =   6075
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.TextBox txt 
      Height          =   345
      Index           =   16
      Left            =   15075
      TabIndex        =   50
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
      TabIndex        =   16
      Text            =   "8"
      ToolTipText     =   "Press F3 for Creation & Press F2 for Selection"
      Top             =   8115
      Width           =   7290
   End
   Begin VB.TextBox txt 
      Height          =   345
      Index           =   15
      Left            =   14940
      TabIndex        =   49
      Text            =   "15"
      Top             =   6435
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.TextBox txt 
      Height          =   375
      Index           =   12
      Left            =   15105
      TabIndex        =   48
      Text            =   "12"
      Top             =   2940
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.TextBox txt 
      Height          =   325
      Index           =   10
      Left            =   15075
      TabIndex        =   47
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
      Left            =   10350
      TabIndex        =   1
      Text            =   "0"
      Top             =   705
      Width           =   1890
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
      TabIndex        =   4
      Text            =   "txttemp(5)"
      ToolTipText     =   "Press F3 for Creation"
      Top             =   1530
      Width           =   6525
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
      TabIndex        =   3
      Text            =   "txttemp(4)"
      ToolTipText     =   "Press F3 for Creation"
      Top             =   885
      Width           =   6525
   End
   Begin VB.TextBox txt 
      Height          =   375
      Index           =   14
      Left            =   15195
      TabIndex        =   40
      Text            =   "14"
      Top             =   3735
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.TextBox txt 
      Alignment       =   1  'Right Justify
      Height          =   360
      Index           =   13
      Left            =   15075
      TabIndex        =   39
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
      TabIndex        =   30
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
      TabIndex        =   31
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
      TabIndex        =   32
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
      TabIndex        =   33
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
      TabIndex        =   29
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
      TabIndex        =   34
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
      TabIndex        =   24
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
      TabIndex        =   25
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
      TabIndex        =   26
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
      TabIndex        =   35
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
      TabIndex        =   27
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
      TabIndex        =   28
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
      TabIndex        =   36
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
      TabIndex        =   42
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
      TabIndex        =   43
      ToolTipText     =   "To Add New Record"
      Top             =   4860
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.TextBox txt 
      Height          =   345
      Index           =   3
      Left            =   15195
      TabIndex        =   46
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
      TabIndex        =   38
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
      Left            =   1560
      TabIndex        =   5
      Text            =   "6"
      Top             =   2295
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
      Height          =   375
      Index           =   9
      Left            =   3090
      MaxLength       =   40
      TabIndex        =   17
      Text            =   "9"
      ToolTipText     =   "Press F3 for Creation & Press F2 for Selection"
      Top             =   8505
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
      TabIndex        =   15
      Text            =   "7"
      ToolTipText     =   "Press F3 for Creation & Press F2 for Selection"
      Top             =   7725
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
      Left            =   10350
      TabIndex        =   6
      Text            =   "11"
      Top             =   1890
      Width           =   1890
   End
   Begin VB.TextBox txt 
      Alignment       =   1  'Right Justify
      Height          =   360
      Index           =   4
      Left            =   15195
      TabIndex        =   37
      Text            =   "4"
      Top             =   1695
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.TextBox txt 
      Height          =   345
      Index           =   2
      Left            =   15195
      TabIndex        =   45
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
      TabIndex        =   44
      Text            =   "1"
      Top             =   615
      Visible         =   0   'False
      Width           =   555
   End
   Begin MSMask.MaskEdBox mskDt 
      Height          =   390
      Index           =   2
      Left            =   10350
      TabIndex        =   2
      Top             =   1200
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
      Bindings        =   "frmJVGST.frx":058A
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
      TabIndex        =   18
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
      Left            =   10350
      TabIndex        =   78
      Top             =   2430
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
      Left            =   3645
      TabIndex        =   77
      Top             =   6210
      Width           =   2175
   End
   Begin VB.Label lbl1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "GST Details    :"
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
      Index           =   6
      Left            =   225
      TabIndex        =   76
      Top             =   3150
      Width           =   1410
   End
   Begin VB.Line Line2 
      X1              =   8370
      X2              =   12420
      Y1              =   1710
      Y2              =   1710
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
      Left            =   8550
      TabIndex        =   75
      Top             =   2475
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
      Left            =   8550
      TabIndex        =   74
      Top             =   1935
      Width           =   1635
   End
   Begin VB.Label lblTotal 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Total"
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
      Left            =   5670
      TabIndex        =   73
      Top             =   6210
      Width           =   2175
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FF00FF&
      BorderWidth     =   2
      X1              =   5670
      X2              =   7920
      Y1              =   5985
      Y2              =   5985
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
      Left            =   3870
      TabIndex        =   72
      Top             =   5295
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
      Left            =   1395
      TabIndex        =   71
      Top             =   5295
      Width           =   1260
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   " Amount  :"
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
      Left            =   450
      TabIndex        =   70
      Top             =   2340
      Width           =   1140
   End
   Begin VB.Shape Shape4 
      BorderColor     =   &H00C00000&
      BorderWidth     =   5
      FillColor       =   &H00C00000&
      Height          =   9510
      Left            =   0
      Top             =   -360
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
      Left            =   3645
      TabIndex        =   69
      Top             =   3870
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
      Left            =   1395
      TabIndex        =   68
      Top             =   4485
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
      Left            =   3870
      TabIndex        =   67
      Top             =   4485
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
      Left            =   1395
      TabIndex        =   66
      Top             =   4890
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
      Left            =   3870
      TabIndex        =   65
      Top             =   4890
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
      TabIndex        =   64
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
      TabIndex        =   63
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
      TabIndex        =   23
      Top             =   7185
      Width           =   2190
   End
   Begin VB.Label lbl1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Entry &No.           :"
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
      Left            =   8550
      TabIndex        =   0
      Top             =   765
      Width           =   1575
   End
   Begin VB.Label lbl1 
      BackStyle       =   0  'Transparent
      Caption         =   "Debit      :"
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
      TabIndex        =   22
      Top             =   1530
      Width           =   1320
   End
   Begin VB.Label lbl1 
      BackStyle       =   0  'Transparent
      Caption         =   "Credit     :"
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
      TabIndex        =   21
      Top             =   885
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
      TabIndex        =   20
      Top             =   1245
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
      TabIndex        =   41
      Top             =   7725
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
      Height          =   3375
      Left            =   225
      Top             =   3465
      Width           =   12135
   End
   Begin VB.Shape Shape2 
      BackColor       =   &H00808080&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00800000&
      FillColor       =   &H00ECFFFF&
      FillStyle       =   0  'Solid
      Height          =   2475
      Left            =   225
      Top             =   540
      Width           =   7995
   End
   Begin VB.Shape Shape7 
      BackColor       =   &H00808080&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00800000&
      FillColor       =   &H00F5F5DE&
      FillStyle       =   0  'Solid
      Height          =   2475
      Left            =   8370
      Top             =   540
      Width           =   4035
   End
End
Attribute VB_Name = "frmJVGST"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Const tblName As String = "tblVoucher"
Const OrderField As String = "Vdt,Vno"
Const IdField As String = "vno"
Const MaxNo As Long = 30
Dim CriteriaStr As String
Dim rstbl As Recordset, ShowFrmFlag As Boolean
Dim varBookMark As Variant
Dim FindFlag As Boolean
Dim FirstTimeFlag As Boolean
Dim FormAction As Integer
Dim frmTypeStr As String
Dim TaxAmt As Double
Dim i As Long
'-- RecVsSale
Dim rsRecVsSale As Recordset
Private Sub cmdBtn_Click(Index As Integer)
Dim FindQStr As String
Dim X As String
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
    FindQStr = "select Vno,Vdt,TblMastAccount.AcName,TblAc1.AcName,Vamt," & X & " where Vtype='" & frmTypeStr & "' and " _
    & " Vyear=" & gCYear & " and VCtrNo=1 and VAcCode=TblMastAccount.AcCode and VAcOCode=TblAc1.AcCode"
    gClsSearch.SearchMultiField FindQStr, "", Array("Vno", "Vdt", "Credit A/c", "Debit A/c", "Amount", "Narration"), Array(700, 1200, 2500, 2500, 1000, 2000), "", " ", "vno,vdt", 0, 0, False, , True
    If gClsSearch.SearchMultiRetCol(0) <> "" Then
        rstbl.MoveFirst
        rstbl.Find "Vno = " & gClsSearch.SearchMultiRetCol(0)
        gClsSearch.SearchMultiRetCol(0) = ""
        ReadFields
    End If
Case 5 'List
    'GenMastList ("LAC")
Case 6 'Add
     FormAction = vbDataActionAddNew
     Call GProcActivateControls(Me, True)
'     lbl1(3).Enabled = True
'     lbl1(4).Enabled = True
     Call GProcSetButtons(Me, vbDataActionAddNew)
     Call GProcClearForm(Me, rstbl, MaxNo, True)
     txt(0).Text = GProcGenerateId(tblName, "VNo", CriteriaStr, gCn)
     mskDt(2) = GProcGenerateVDt(tblName, "VDT", CriteriaStr)
     SetControlEd
     If FirstTimeFlag = True Then
        FirstTimeFlag = False
     Else
        txt(0).SetFocus
     End If
     FillCombo
     SetGrid
Case 7 'Modify
    'If mURecModi = True Then
    '    SearchHelp
    '    DoEvents
    If IsRelatedRecord = False Then
        FormAction = vbDataActionUpdate
        Call GProcActivateControls(Me, True)
'        lbl1(3).Enabled = True
'        lbl1(4).Enabled = True
        Call GProcSetButtons(Me, vbDataActionUpdate)
        SetControlEd
        txt(0).Enabled = False
         '--Update Outstanding Balance
        GProcUpdateOutstandingBal CLng(txt(0)), UCase(txt(3)), gCYear, True
        mskDt(2).SetFocus
    End If
Case 8 'Delete
   ' If mURecDel = True Then
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
        gCn.Execute "Delete from tblvoucher where " & CriteriaStr & " and vno = " & txt(0)
        Dim TaxAcCodeCGST As Long
        Dim TaxAcCodeSGST As Long
        Dim TaxAcCodeIGST As Long
        If lbl1(3) = "Credit" Then
            GProcCreateVoucher FormAction, txt(0), 1, txt(2), frmTypeStr, txt(4), txt(5), Val(txt(6)), txt(11), txt(12), txt(13), txt(7), txt(8), txt(9), txt(10), False, txt(20), txt(21), txt(22), txt(23), txt(24), txt(25), txt(26), txt(27), txt(28), txt(29), txt(30)
            GProcCreateVoucher FormAction, txt(0), 2, txt(2), frmTypeStr, txt(5), txt(4), Val(txt(6)) * (-1), txt(11), txt(12), txt(13), txt(7), txt(8), txt(9), txt(10), False, txt(20), txt(21), txt(22), txt(23), txt(24), txt(25), txt(26), txt(27), txt(28), txt(29), txt(30)
'            If frmTypeStr = "SN" Then lbl1_Click (3)
        Else
            '--- Account Code
            TaxAcCodeCGST = GProcGetColumnValue("tblMastNarration", "Narration", "G S T", "S", "CGSTPayAc", "N")
            TaxAcCodeSGST = GProcGetColumnValue("tblMastNarration", "Narration", "G S T", "S", "SGSTPayAc", "N")
            TaxAcCodeIGST = GProcGetColumnValue("tblMastNarration", "Narration", "G S T", "S", "IGSTPayAc", "N")
            
            TaxAmt = Val(txt(23)) + Val(txt(25)) + Val(txt(27))
        
            GProcCreateVoucher FormAction, txt(0), 1, txt(2), frmTypeStr, txt(4), txt(5), Val(txt(6)) * (-1), txt(11), txt(12), txt(13), txt(7), txt(8), txt(9), txt(10), False, txt(20), txt(21), txt(22), txt(23), txt(24), txt(25), txt(26), txt(27), txt(28), txt(29), txt(30)
            GProcCreateVoucher FormAction, txt(0), 2, txt(2), frmTypeStr, txt(5), txt(4), Val(txt(6)) - TaxAmt, txt(11), txt(12), txt(13), txt(7), txt(8), txt(9), txt(10), False, txt(20), txt(21), txt(22), txt(23), txt(24), txt(25), txt(26), txt(27), txt(28), txt(29), txt(30)

            i = 3
            If Val(txt(23)) > 0 Then
               GProcCreateVoucher FormAction, txt(0), i, txt(2), frmTypeStr, TaxAcCodeCGST, txt(4), Val(txt(23)), txt(11), txt(12), txt(13), txt(7), txt(8), txt(9), txt(10), False, txt(20), txt(21), txt(22), txt(23), txt(24), txt(25), txt(26), txt(27), txt(28), txt(29), txt(30)
               i = i + 1
               DoEvents
            End If
            If Val(txt(25)) > 0 Then
               GProcCreateVoucher FormAction, txt(0), i, txt(2), frmTypeStr, TaxAcCodeSGST, txt(4), Val(txt(25)), txt(11), txt(12), txt(13), txt(7), txt(8), txt(9), txt(10), False, txt(20), txt(21), txt(22), txt(23), txt(24), txt(25), txt(26), txt(27), txt(28), txt(29), txt(30)
               i = i + 1
               DoEvents
            End If
            If Val(txt(27)) > 0 Then
               GProcCreateVoucher FormAction, txt(0), i, txt(2), frmTypeStr, TaxAcCodeIGST, txt(4), Val(txt(27)), txt(11), txt(12), txt(13), txt(7), txt(8), txt(9), txt(10), False, txt(20), txt(21), txt(22), txt(23), txt(24), txt(25), txt(26), txt(27), txt(28), txt(29), txt(30)
               i = i + 1
               DoEvents
            End If
'            If frmTypeStr <> "SN" Then lbl1_Click (3)
        End If
        rstbl.Requery
        '--- Related Rcord Rec Vs Sale
        SaveRelatedRecord
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
        msGrid.Col = 0
        msGrid.Row = 1
        txtGrid.Move msGrid.CellLeft + msGrid.Left, msGrid.CellTop + msGrid.Top, msGrid.CellWidth, msGrid.CellHeight
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
    msGrid.Col = 0
    If msGrid.Row >= 1 Then msGrid.Row = 1
    txtGrid.Move msGrid.CellLeft + msGrid.Left, msGrid.CellTop + msGrid.Top, msGrid.CellWidth, msGrid.CellHeight
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
            SendKeys "{TAB}"
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
    Me.Left = Screen.Width / 2 - Me.Width / 2
    frmTypeStr = gfrmTypeStr
    CriteriaStr = " VType = '" & frmTypeStr & "' and VYear=" & gCYear
    Y = "Select * from " & tblName & " where " & CriteriaStr & " and vctrno = 1"
    If Trim(OrderField) <> "" Then Y = Y & " order by " & OrderField
    GProcRstOpen rstbl, Y, "O"
    Select Case frmTypeStr
    Case "JV"
        Label6.Caption = "    Journal Voucher"
        lbl1(3).Caption = "Credit"
        lbl1(4).Caption = "Debit"
    Case "SN"
        Label6.Caption = "    Debit Note (Against Sales)"
        lbl1(3).Caption = "Debit"
        lbl1(4).Caption = "Credit"
'        lbl1_Click (3)
    Case "PN"
        Label6.Caption = "    Credit Note (Against Purchse)"
        lbl1(3).Caption = "Credit"
        lbl1(4).Caption = "Debit"
    End Select
    FirstTimeFlag = True
    '---- Button Visiblity Set As Per User Previlage
    GProcSetButtonVisiblity Me
    cmdBtn_Click (6)
    '---Grid
    SetGrid
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
GProcCheckForNumber Me, rstbl, MaxNo
SaveBtnEd
CalAmount
End Sub
Private Sub txt_KeyPress(Index As Integer, KeyAscii As Integer)
KeyAscii = GProcValidateKey(rstbl, Index, KeyAscii, txt(Index).Text)
If Index = 11 Then
   Dim X As String
   Dim Pcode As Long
   Pcode = GProcGetColumnValue("tblMastAccount", "AcName", txttemp(4), "S", "Accode", "N") ', gCn)
   X = "select tblOutstanding.OutBillNo,format(iif(LEFT(tblOutstanding.VTYPE,1) = 'P',tblAddLess.AdGpdt,tblOutstanding.Vdt),'dd/mm/yyyy') as BillDt,tblOutstanding.OutAmount,tblOutstanding.Vno,tblOutstanding.Vtype,tblOutstanding.Vyear from tblOutstanding, tblAddLess where " _
       & "tblOutstanding.Vno = tblAddLess.Vno and tblOutstanding.Vtype = tblAddLess.Vtype and tblOutstanding.Vyear = tblAddLess.Vyear and tblOutstanding.OutAcCode =  " & Pcode & " and tblOutstanding.vdt <= cdate('" & CDate(mskDt(2)) & "')"
   gClsSearch.SearchMultiField X, "", Array("Bill No", "Bill Date", "Bill Amount", "", "", ""), Array(2000, 2000, 3000, 0, 0, 0), "", "", "tblOutstanding.vdt,tblOutstanding.vno", 0, 0, False, , True
   KeyAscii = 0
   If gClsSearch.SearchMultiRetCol(0) <> "" Then
      txt(11) = gClsSearch.SearchMultiRetCol(0)
      mskDt(12) = gClsSearch.SearchMultiRetCol(1)
      txt(28) = gClsSearch.SearchMultiRetCol(3)
      txt(29) = gClsSearch.SearchMultiRetCol(4)
      txt(30) = gClsSearch.SearchMultiRetCol(5)
      txt(22).SetFocus
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
txt(21).Enabled = False
txt(23).Enabled = False
txt(25).Enabled = False
txt(27).Enabled = False
mskDt(12).Enabled = False
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
   lblTotal.Visible = False
   lbl1(2).Visible = False
   lbl1(7).Visible = False
   For i = 21 To 27
       txt(i).Visible = False
   Next i
End If
End Sub
Private Sub FillTempFromTxt()
mskDt(2) = IIf(IsDate(txt(2)), CDate(txt(2)), "__/__/____") '--Vdt
mskDt(12) = IIf(IsDate(txt(12)), CDate(txt(12)), "__/__/____") '--Bill Date

txt(6) = Format(Abs(Val(txt(6))), "##0.00")
txttemp(4) = GProcGetColumnValue("tblMastAccount", "AcCode", txt(4), "N", "AcName", "S") ', gCn)
txttemp(5) = GProcGetColumnValue("tblMastAccount", "AcCode", txt(5), "N", "AcName", "S") ', gCn)
txttemp(20) = GProcGetColumnValue("tblMastNarration", "NarrCode", txt(20), "N", "Narration", "S")
lblTotal.Caption = txt(6)

'-- Fill Related Record Rec Vs Sale
FillRelatedRecord
SetControlEd
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
Dim BillDet As String
If Trim(txtGrid.Text) <> "" Then
        msGrid.Text = txtGrid.Text
End If
txtGrid.Visible = False
txtGrid.Text = ""
BillDet = ""
CheckForNumberGrid
With msGrid
    '-- Duplicate invoice selection
    For i = 1 To .Rows - 1
        If Trim(.TextMatrix(i, 1)) <> "" Then
           BillDet = BillDet + .TextMatrix(i, 1) + "-" + CStr(Val(.TextMatrix(i, 4)) + Val(.TextMatrix(i, 6)) + Val(.TextMatrix(i, 9))) + ":"
        End If
        '-- Int A/c
           If CDbl(.TextMatrix(i, 9)) > 0 Then
                If gLatePayIntAcCode = GProcGetColumnValue("tblmastAccount", "AcName", "-", "S", "AcCode", "N") Then
                    IntAcCode = 0
                Else
                    IntAcCode = gLatePayIntAcCode
                End If
                If IntAcCode = 0 Then
                    MsgBox "Check Late Pay Interest Account.", vbInformation + vbOKOnly, Me.Caption
                    GProcShowForm frmSetting, frmMain.mnuMstfrm(12)
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
    If BillDet <> "" Then
        BillDet = Left(Trim(BillDet) + Space(120), 120)
        txt(8) = Left(BillDet, 40)
        txt(9) = Mid(BillDet, 41, 40)
        txt(10) = Mid(BillDet, 81, 40)
    End If
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
Dim X As String
Dim GName As String
Dim MillAgCode1 As Long
Dim MillAgCode2 As Long
Dim oldFrmTypeStr As String
Dim IsMonthLastDt As Boolean
Dim rstAddCheck As Recordset
GProcCheckForNumber Me, rstbl, MaxNo
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
'---Mill Expenses (not on month last date)
MillAgCode1 = GProcGetColumnValue("tblMastAccount", "AcName", txttemp(5), "S", "AgCode", "N", " ACMILLTYPE<>0")
MillAgCode2 = GProcGetColumnValue("tblMastAccount", "AcName", txttemp(4), "S", "AgCode", "N", " ACMILLTYPE<>0")
If Format(DateAdd("d", 1, mskDt(2)), "dd") = 1 Then
    IsMonthLastDt = True
Else
    IsMonthLastDt = False
End If
If IsMonthLastDt = False Then
    If txttemp(20) = "" And (MillAgCode1 = 90029 Or MillAgCode2 = 90029) Then
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
    Case 20 '--Mill Expes. Narration
        gClsSearch.SearchMultiField "TblMastNarration", "Narration", Array("Mill Expenses Description"), Array(4000), "NarrType='M'", Chr(KeyAscii), "Narration", LeftPos, TopPos, True
        KeyAscii = 0
        If gClsSearch.SearchMultiRetCol(0) <> "" Then
            txttemp(Index) = gClsSearch.SearchMultiRetCol(0)
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
Select Case Index
Case 0, 1:
    '------ Validate
    If IsNumeric(txtno(0)) = False Then txtno(0) = 0
    If IsNumeric(txtno(1)) = False Then txtno(1) = txtno(0)
    If CLng(txtno(1)) < CLng(txtno(0)) Then txtno(1) = txtno(0)
    '-------
    If Index = 0 Then GProcCrystalRptPreparation cryReport, "Printer"
    If Index = 1 Then GProcCrystalRptPreparation cryReport, "Window"
    cryReport.WindowTitle = "Journal Entry"
    Vno = CLng(txtno(0))
    Do While Vno <= txtno(1)
        X = "Select * From tblVoucher where  vNo=" & Vno & " and VCTRNO =1 and VType='" & frmTypeStr & "' and VYear=" & gCYear
        If GProcRstOpen(Rs1, X, "R") > 0 Then
            With cryReport
                .SelectionFormula = "{tblVoucher.vNo} = " & Rs1!Vno & " and {tblVoucher.VCTRNO}=1 and {tblVoucher.VType}='" & Rs1!Vtype & "' and {tblVoucher.VYear}=" & Rs1!VYear
                .Formulas(5) = "wAmountInWords = '" & GProcAmountInWords(Abs(Rs1!vamt)) & "'"
                If frmTypeStr = "SN" Then '-- Debit Note
                    .ReportFileName = gReportPath & "rptVouDbNtPrint.rpt"
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
Case 2:
    DoEvents
    Frame1.Visible = False
End Select
End Sub
'---------------- Grid
Private Sub SetGrid()
Dim X As String, x1Rec As Double
Set rsRecVsSale = New Recordset
X = "select InvTp, InvBillNo, InvDt, BillAmt, AdjAmt, DbNtAmt, AdjDbNtAmt, LateDays, Interest, IntReceived ,InvYear,InvNo,JVNOINT" _
& " from tblRecVsSale where " _
& " Recvtype='" & frmTypeStr & "' and Recvno=" & txt(0) & " and RecVYear=" & gCYear & " order by RecVCtrlNo"
x1Rec = GProcRstOpen(rsRecVsSale, X, "R")
With msGrid
    .Clear
    .FormatString = "<Type|>Inv. No   |<Inv. Date      |>Invoice Amt |>Adj. Amount|>DbNt  Amt|>Adj DbNt Amt|>Late Days|>Late Pay Int|>Int Received|||"
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
        Pending_Invoice
    End If
End With

'---- Set TxtGrid
If Trim(msGrid.Text) <> "" Then
    txtGrid = Trim(msGrid.Text)
End If
If msGrid.Col < (rsRecVsSale.Fields.Count - 3) Then
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
        .TextMatrix(.Row, 8) = Format(Round((((Val(.TextMatrix(.Row, 4)) * 1.5) / 3000) * Val(.TextMatrix(.Row, 7))), 2), "###0.00")
    End With
   ' CalBalance
Case 6, 9:
   ' CalBalance
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
'                PartyCode = GProcGetColumnValue("TblMastAccount", "AcName", txttemp(4), "S", "Accode", "N")
'                '''''gClsSearch.SearchMultiField "tblMastItem,tblMastaccount", "ItName,AcName,ItStdpack,ItStdrateper", Array("Count Name", "Mill", "Std. Pack", "Rate Per"), Array(txtGrid.Width, 2000, 1000, 1000), "tblMastItem.ITMillCode=tblMastaccount.acCode and ITMillCode=" & MillCode, Chr(KeyAscii), "ItName", LeftPos, TopPos, True
'                PendingInvoice
'                txtGrid = ""
'                KeyAscii = 0
'                If gClsSearch.SearchMultiRetCol(0) <> "" Then '-- Item Wt ,Rate From Master
'                    txtGrid = gClsSearch.SearchMultiRetCol(0)
'                    If IsNumeric(.TextMatrix(.Row, 5)) = False Then .TextMatrix(.Row, 5) = 0
'                    If IsNumeric(.TextMatrix(.Row, 7)) = False Then .TextMatrix(.Row, 7) = 0
'                    'If .TextMatrix(.Row, 5) = 0 Then .TextMatrix(.Row, 5) = gClsSearch.SearchMultiRetCol(2)
'                    If .TextMatrix(.Row, 7) = 0 Then .TextMatrix(.Row, 7) = gClsSearch.SearchMultiRetCol(3)
'                    txtGrid_KeyPress 13
'                End If
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
gCn.Execute "delete from tblVoucher where Vtype='JV' and VNo in (select JVNOINT from tblRecVsSale where " _
& " INTRECEIVED>0 and RecVtype='" & frmTypeStr & "' and RecVNo=" & txt(0) & " and  RecVyear= " & gCYear & ")  and " _
 & " Vyear= " & gCYear
'--- Rec Vs Sale
gCn.Execute "delete from tblRecVsSale where RecVtype='" & frmTypeStr & "' and RecVNo=" & txt(0) & " and " _
 & " RecVyear= " & gCYear
End Sub

Private Sub SaveRelatedRecord()
Dim i As Long, X As String
Dim IntRecAmt As Double
Dim JVVNo As Long
Dim rsRecVsSale As Recordset, rsOutst As Recordset
'-- Delete Records
DeleteRelatedRecord
'--- Save Grid Records
'--- tblRecSale
X = "select * from tblRecVsSale where Recvtype='" & frmTypeStr & "' and Recvno=" & txt(0) & " and " _
& " RecVYear=" & gCYear
GProcRstOpen rsRecVsSale, X, "O"
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
        rsRecVsSale!PartyCode = GProcGetColumnValue("tblMastAccount", "AcName", txttemp(4), "S", "AcCode", "N")
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
        If FormAction = vbDataActionAddNew And IntRecAmt > 0 Then
            JVVNo = GProcGenerateId("tblVoucher", "Vno", " vyear=" & gCYear & " and Vtype='JV'")
        Else
            If IntRecAmt > 0 Then
                If CLng(.TextMatrix(i, 12)) > 0 And GProcGetColumnValue("tblVoucher", "Vno", .TextMatrix(i, 12), "N", "Vno", "N", " Vtype='JV' and Vyear=" & gCYear) = 0 Then
                    JVVNo = CLng(.TextMatrix(i, 12))
                Else
                    JVVNo = GProcGenerateId("tblVoucher", "Vno", " vyear=" & gCYear & " and Vtype='JV'")
                End If
            End If
        End If
        If IntRecAmt > 0 Then
            GProcCreateVoucher FormAction, JVVNo, 1, mskDt(2), "JV", gLatePayIntAcCode, txt(4), IntRecAmt, IsRefEntryType:=True
            GProcCreateVoucher FormAction, JVVNo, 2, mskDt(2), "JV", txt(4), gLatePayIntAcCode, -IntRecAmt, IsRefEntryType:=True
        End If
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
CTaxAmt = 0
STaxAmt = 0
ITaxAmt = 0

Taxable = Round((Val(txt(6)) * 100) / (100 + Val(txt(22)) + Val(txt(24)) + Val(txt(26))), 2)

CTaxAmt = (Taxable * Val(txt(22))) / 100
CTaxAmt = GProcMakeRounding(CTaxAmt)
txt(23) = Format(CTaxAmt, FStr)

STaxAmt = (Taxable * Val(txt(24))) / 100
STaxAmt = GProcMakeRounding(STaxAmt)
txt(25) = Format(STaxAmt, FStr)

ITaxAmt = (Taxable * Val(txt(26))) / 100
ITaxAmt = GProcMakeRounding(ITaxAmt)
txt(27) = Format(ITaxAmt, FStr)

txt(21) = Format(Val(txt(6)) - (CTaxAmt + STaxAmt + ITaxAmt), FStr)
lblTotal.Caption = txt(6)
End Sub

