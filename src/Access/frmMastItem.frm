VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form frmMastItem 
   BackColor       =   &H00C0E0FF&
   BorderStyle     =   0  'None
   Caption         =   " Count Master"
   ClientHeight    =   7995
   ClientLeft      =   0
   ClientTop       =   60
   ClientWidth     =   12015
   FillColor       =   &H000040C0&
   FillStyle       =   0  'Solid
   Icon            =   "frmMastItem.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7995
   ScaleWidth      =   12015
   ShowInTaskbar   =   0   'False
   Begin VB.CheckBox Check1 
      BackColor       =   &H00C0FFC0&
      Caption         =   "Print Ticket Name In Sale Bill ?"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF00FF&
      Height          =   405
      Left            =   5895
      TabIndex        =   66
      Top             =   7155
      Width           =   3960
   End
   Begin VB.TextBox txt 
      BackColor       =   &H00FFFFFF&
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
      Index           =   24
      Left            =   0
      TabIndex        =   65
      Text            =   "24"
      Top             =   0
      Visible         =   0   'False
      Width           =   315
   End
   Begin VB.TextBox txt 
      BackColor       =   &H00FFFFFF&
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
      Index           =   23
      Left            =   0
      TabIndex        =   64
      Text            =   "23"
      Top             =   7695
      Visible         =   0   'False
      Width           =   315
   End
   Begin VB.CheckBox chkIsIncludingGst 
      BackColor       =   &H00EDFDFE&
      Caption         =   "Is Rate Excluding GST ?"
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
      Height          =   405
      Left            =   5820
      TabIndex        =   63
      Top             =   4920
      Width           =   4095
   End
   Begin VB.TextBox txt 
      BackColor       =   &H00FFFFFF&
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
      Index           =   22
      Left            =   13590
      TabIndex        =   62
      Text            =   "22"
      Top             =   3195
      Visible         =   0   'False
      Width           =   315
   End
   Begin VB.TextBox txt 
      BackColor       =   &H00FFFFFF&
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
      Index           =   21
      Left            =   13140
      TabIndex        =   61
      Text            =   "21"
      Top             =   3105
      Visible         =   0   'False
      Width           =   315
   End
   Begin VB.TextBox txt 
      BackColor       =   &H00FFFFFF&
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
      Index           =   20
      Left            =   13140
      TabIndex        =   60
      Text            =   "20"
      Top             =   2700
      Visible         =   0   'False
      Width           =   315
   End
   Begin VB.TextBox txt 
      BackColor       =   &H00FFFFFF&
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
      Left            =   13095
      TabIndex        =   59
      Text            =   "19"
      Top             =   2295
      Visible         =   0   'False
      Width           =   315
   End
   Begin VB.TextBox txt 
      BackColor       =   &H00FFFFFF&
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
      Index           =   18
      Left            =   13095
      TabIndex        =   58
      Text            =   "18"
      Top             =   1890
      Visible         =   0   'False
      Width           =   315
   End
   Begin VB.ComboBox Cbo 
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
      ItemData        =   "frmMastItem.frx":058A
      Left            =   6960
      List            =   "frmMastItem.frx":0597
      TabIndex        =   10
      Text            =   "Cbo"
      Top             =   3840
      Width           =   1935
   End
   Begin VB.TextBox txt 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFFFF&
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
      Left            =   12195
      TabIndex        =   55
      Text            =   "17"
      Top             =   7920
      Visible         =   0   'False
      Width           =   1350
   End
   Begin VB.ComboBox Cbo 
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
      ItemData        =   "frmMastItem.frx":05B7
      Left            =   7515
      List            =   "frmMastItem.frx":05C4
      TabIndex        =   16
      Text            =   "Cbo(8)"
      Top             =   5940
      Width           =   2055
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
      Height          =   360
      Index           =   15
      Left            =   2280
      TabIndex        =   14
      Text            =   "15"
      Top             =   6855
      Width           =   1515
   End
   Begin VB.TextBox txt 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFFFF&
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
      Left            =   12240
      TabIndex        =   51
      Text            =   "16"
      Top             =   7560
      Visible         =   0   'False
      Width           =   1350
   End
   Begin VB.CommandButton cmdBtn 
      Caption         =   "Creation &To Firm"
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
      Index           =   12
      Left            =   10485
      Style           =   1  'Graphical
      TabIndex        =   50
      Top             =   2550
      Width           =   1095
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
      Height          =   360
      Index           =   14
      Left            =   2280
      TabIndex        =   7
      Text            =   "14"
      Top             =   4395
      Width           =   2115
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
      Index           =   13
      Left            =   2280
      TabIndex        =   8
      Text            =   "13"
      Top             =   4950
      Width           =   2115
   End
   Begin VB.TextBox txt 
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
      Height          =   360
      Index           =   12
      Left            =   6975
      TabIndex        =   11
      Text            =   "12"
      Top             =   4335
      Width           =   2760
   End
   Begin VB.TextBox txt 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFFFF&
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
      Left            =   12690
      TabIndex        =   49
      Text            =   "11"
      Top             =   3645
      Visible         =   0   'False
      Width           =   540
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
      Height          =   360
      Index           =   10
      Left            =   7515
      TabIndex        =   17
      Text            =   "10"
      Top             =   6360
      Width           =   1155
   End
   Begin VB.TextBox txt 
      BackColor       =   &H00FFFFFF&
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
      Index           =   7
      Left            =   12780
      TabIndex        =   47
      Text            =   "7"
      Top             =   4590
      Visible         =   0   'False
      Width           =   315
   End
   Begin VB.TextBox txt 
      BackColor       =   &H00FFFFFF&
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
      Index           =   6
      Left            =   12780
      TabIndex        =   46
      Text            =   "6"
      Top             =   5130
      Visible         =   0   'False
      Width           =   315
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
      Height          =   375
      Index           =   9
      Left            =   2280
      TabIndex        =   3
      Text            =   "txttemp(9)"
      ToolTipText     =   "Press F3 For Creation"
      Top             =   1920
      Width           =   5880
   End
   Begin VB.ComboBox Cbo 
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
      ItemData        =   "frmMastItem.frx":05D6
      Left            =   2280
      List            =   "frmMastItem.frx":05E3
      TabIndex        =   12
      Text            =   "Cbo(7)"
      Top             =   5940
      Width           =   1830
   End
   Begin VB.ComboBox Cbo 
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
      ItemData        =   "frmMastItem.frx":05F5
      Left            =   6990
      List            =   "frmMastItem.frx":05FF
      TabIndex        =   9
      Text            =   "Cbo"
      Top             =   3360
      Width           =   1935
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
      Height          =   360
      Index           =   8
      Left            =   2280
      TabIndex        =   13
      Text            =   "8"
      Top             =   6360
      Width           =   1515
   End
   Begin VB.TextBox txt 
      BackColor       =   &H00FFFFFF&
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
      Index           =   9
      Left            =   12825
      TabIndex        =   45
      Text            =   "9"
      Top             =   4185
      Visible         =   0   'False
      Width           =   315
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
      Height          =   360
      Index           =   5
      Left            =   2280
      TabIndex        =   6
      Text            =   "5"
      Top             =   3840
      Width           =   2115
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
      Height          =   360
      Index           =   4
      Left            =   2280
      TabIndex        =   5
      Text            =   "4"
      Top             =   3360
      Width           =   2115
   End
   Begin VB.TextBox txt 
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
      Height          =   375
      Index           =   3
      Left            =   2280
      TabIndex        =   2
      Text            =   "3"
      Top             =   1440
      Width           =   3630
   End
   Begin VB.TextBox txt 
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
      Height          =   375
      Index           =   2
      Left            =   2280
      TabIndex        =   4
      Text            =   "2"
      Top             =   2400
      Width           =   3630
   End
   Begin VB.TextBox txt 
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
      Height          =   375
      Index           =   1
      Left            =   2280
      TabIndex        =   1
      Text            =   "1"
      Top             =   960
      Width           =   5880
   End
   Begin VB.CommandButton cmdBtn 
      Caption         =   "Update Brokerage"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   720
      Index           =   11
      Left            =   10440
      Style           =   1  'Graphical
      TabIndex        =   42
      ToolTipText     =   "Label"
      Top             =   6300
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
      Left            =   10485
      Style           =   1  'Graphical
      TabIndex        =   40
      ToolTipText     =   "Exit from form"
      Top             =   5640
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
      Left            =   10485
      Style           =   1  'Graphical
      TabIndex        =   33
      ToolTipText     =   "cancel"
      Top             =   5160
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
      Left            =   10485
      Style           =   1  'Graphical
      TabIndex        =   32
      ToolTipText     =   "Save Record"
      Top             =   4680
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
      Left            =   12735
      Style           =   1  'Graphical
      TabIndex        =   41
      ToolTipText     =   "Print Record"
      Top             =   6255
      Visible         =   0   'False
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
      Left            =   10485
      Style           =   1  'Graphical
      TabIndex        =   31
      ToolTipText     =   "Delete Record"
      Top             =   4185
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
      Height          =   405
      Index           =   7
      Left            =   10485
      Style           =   1  'Graphical
      TabIndex        =   30
      ToolTipText     =   " Modify Record"
      Top             =   3705
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
      Left            =   10485
      Style           =   1  'Graphical
      TabIndex        =   29
      ToolTipText     =   " Add New Record"
      Top             =   3225
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
      Index           =   5
      Left            =   10485
      Style           =   1  'Graphical
      TabIndex        =   34
      ToolTipText     =   "List Of Record"
      Top             =   2085
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
      Left            =   10485
      Style           =   1  'Graphical
      TabIndex        =   35
      ToolTipText     =   "Find Record"
      Top             =   1620
      Width           =   1095
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
      Left            =   11085
      Style           =   1  'Graphical
      TabIndex        =   39
      ToolTipText     =   "Last"
      Top             =   1155
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
      Left            =   10485
      Style           =   1  'Graphical
      TabIndex        =   38
      ToolTipText     =   "First"
      Top             =   1155
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
      Left            =   11085
      Style           =   1  'Graphical
      TabIndex        =   37
      ToolTipText     =   "Next"
      Top             =   675
      Width           =   495
   End
   Begin VB.TextBox txt 
      BackColor       =   &H00FFFFFF&
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
      Index           =   0
      Left            =   2760
      TabIndex        =   43
      Text            =   "0"
      Top             =   600
      Visible         =   0   'False
      Width           =   315
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
      Left            =   10485
      Style           =   1  'Graphical
      TabIndex        =   36
      ToolTipText     =   "Previous"
      Top             =   675
      Width           =   495
   End
   Begin MSMask.MaskEdBox mskDt 
      Height          =   360
      Left            =   2280
      TabIndex        =   15
      Top             =   7275
      Width           =   1515
      _ExtentX        =   2672
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
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "HSN Code :"
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
      Index           =   16
      Left            =   5760
      TabIndex        =   57
      Top             =   4380
      Width           =   1080
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Item Type  :"
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
      Index           =   15
      Left            =   5760
      TabIndex        =   56
      Top             =   3870
      Width           =   1035
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00000080&
      X1              =   5205
      X2              =   5205
      Y1              =   5745
      Y2              =   7665
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "&Purch Brkrage On :"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H004D177D&
      Height          =   240
      Index           =   14
      Left            =   5715
      TabIndex        =   54
      Top             =   6000
      Width           =   1680
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Before Date          :"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H004D177D&
      Height          =   240
      Index           =   13
      Left            =   480
      TabIndex        =   53
      Top             =   7335
      Width           =   1680
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Old Brokerage Rt. :"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H004D177D&
      Height          =   240
      Index           =   12
      Left            =   480
      TabIndex        =   52
      Top             =   6915
      Width           =   1680
   End
   Begin VB.Image Image1 
      Height          =   1485
      Left            =   8460
      Picture         =   "frmMastItem.frx":0611
      Stretch         =   -1  'True
      Top             =   990
      Width           =   1365
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Min Rate           :"
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
      Left            =   480
      TabIndex        =   24
      Top             =   4395
      Width           =   1500
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Max Rate          :"
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
      Index           =   1
      Left            =   480
      TabIndex        =   25
      Top             =   4905
      Width           =   1500
   End
   Begin VB.Shape Shape6 
      BorderColor     =   &H00C00000&
      BorderWidth     =   5
      FillColor       =   &H00C00000&
      Height          =   8160
      Left            =   0
      Top             =   -225
      Width           =   11970
   End
   Begin VB.Label lblActionStatus 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "View Record.."
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   375
      Left            =   9090
      TabIndex        =   48
      Top             =   15
      Width           =   2475
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "&Purch Brok Rt.    :"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H004D177D&
      Height          =   240
      Index           =   0
      Left            =   5715
      TabIndex        =   28
      Top             =   6420
      Width           =   1590
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Brokerage Rate :"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H004D177D&
      Height          =   375
      Index           =   10
      Left            =   480
      TabIndex        =   27
      Top             =   6360
      Width           =   1470
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "M&ill Name         :"
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
      Index           =   9
      Left            =   480
      TabIndex        =   19
      Top             =   1920
      Width           =   1560
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "&Brokerage On     :"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H004D177D&
      Height          =   240
      Index           =   8
      Left            =   480
      TabIndex        =   26
      Top             =   5940
      Width           =   1560
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Item Unit  :"
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
      Index           =   7
      Left            =   5760
      TabIndex        =   23
      Top             =   3405
      Width           =   1035
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "&Rate Scale        :"
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
      Left            =   480
      TabIndex        =   22
      Top             =   3840
      Width           =   1500
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Std Packing      :"
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
      Index           =   5
      Left            =   480
      TabIndex        =   21
      Top             =   3360
      Width           =   1485
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Short Name      :"
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
      Index           =   4
      Left            =   480
      TabIndex        =   18
      Top             =   1440
      Width           =   1545
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Ticket Name     :"
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
      Left            =   480
      TabIndex        =   20
      Top             =   2400
      Width           =   1545
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Count &Name     :"
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
      Index           =   2
      Left            =   480
      TabIndex        =   0
      Top             =   960
      Width           =   1530
   End
   Begin VB.Label Label6 
      BackColor       =   &H00C00000&
      Caption         =   "   Count Master"
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
      TabIndex        =   44
      Top             =   0
      Width           =   11985
   End
   Begin VB.Shape Shape1 
      FillColor       =   &H00BDDFD1&
      FillStyle       =   0  'Solid
      Height          =   5730
      Left            =   10290
      Shape           =   4  'Rounded Rectangle
      Top             =   495
      Width           =   1485
   End
   Begin VB.Shape Shape2 
      BackColor       =   &H00808080&
      BorderColor     =   &H004D177D&
      FillColor       =   &H00ECFFFF&
      FillStyle       =   0  'Solid
      Height          =   2175
      Left            =   240
      Top             =   720
      Width           =   9705
   End
   Begin VB.Shape Shape3 
      BorderColor     =   &H00000080&
      FillColor       =   &H00EDFDFE&
      FillStyle       =   0  'Solid
      Height          =   2370
      Left            =   270
      Top             =   3150
      Width           =   9705
   End
   Begin VB.Shape Shape4 
      BorderColor     =   &H00000080&
      FillColor       =   &H00EDFDFE&
      FillStyle       =   0  'Solid
      Height          =   1935
      Left            =   240
      Top             =   5745
      Width           =   9750
   End
End
Attribute VB_Name = "frmMastItem"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Const tblName As String = "tblMastItem"
Const IdField As String = "ItCode"
Const OrderField As String = "ItName"
Const MaxNo As Long = 24
Dim rstbl As Recordset
Dim varBookMark As Variant
Dim FirstTimeFlag As Boolean
Dim FormAction As Integer
Private Sub Cbo_KeyPress(Index As Integer, KeyAscii As Integer)
If KeyAscii = 13 Then Exit Sub
If Index = 7 Or Index = 6 Then KeyAscii = 0
End Sub
Private Sub cmdBtn_Click(Index As Integer)
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
    gClsSearch.SearchMultiField "tblMastItem,tblMastAccount", "ItName,ItCode,Itticket,Acname,ItStdpack,ItStdrateper,ITUnit,ItBrokRt", Array("Item", "Item", "Ticket", "Mill", "Std.Packing", "Rate Per", "Unit", "Brk.Rt"), Array(2000, 0, 1000, 2000, 1000, 1000, 1000, 1200), " tblMastItem.ItMillcode=tblMastAccount.AcCode and tblMastItem.ItOtherItem = 0 ", "", "ItName", txt(1).Left + Me.Left, txt(1).Top + Me.Top + 650
    If gClsSearch.SearchMultiRetCol(0) <> "" Then
        rstbl.MoveFirst
        rstbl.Find "ItCode= " & gClsSearch.SearchMultiRetCol(1)
        gClsSearch.SearchMultiRetCol(0) = ""
        ReadFields
    End If
Case 5 'List
    frmMain.mnuRptGenrSubLst_Click (5)
Case 6 'Add
    FormAction = vbDataActionAddNew
    Call GProcActivateControls(Me, True)
    Call GProcSetButtons(Me, vbDataActionAddNew)
    Call GProcClearForm(Me, rstbl, MaxNo, True)
    If FirstTimeFlag = True Then
        FirstTimeFlag = False
    Else
        txt(1).SetFocus
    End If
    
Case 7 'Modify
    'If mURecModi = True Then
    '    SearchHelp
    '    DoEvents
    'If IsRelatedRecord = False Then
        cmdBtn_Click 4 '-- Find
        If gClsSearch.SearchMultiRetCol(0) = "" Then Exit Sub
        FormAction = vbDataActionUpdate
        Call GProcActivateControls(Me, True)
        Call GProcSetButtons(Me, vbDataActionUpdate)
        '--Related Mill & Item Record
        If IsRelatedRecord = True Then txttemp(9).Enabled = False
        txt(1).SetFocus
    'End If
Case 8 'Delete
    'If mURecDel = True Then
     'If IsRelatedRecord = False Then
        If MsgBox("Do You Really Want to Delete This Record?", vbQuestion + vbYesNo + vbDefaultButton1, "Confirm Delete") = vbYes Then
            If rstbl.RecordCount > 1 Then
                rstbl.Delete
                cmdBtn_Click (1)
            ElseIf rstbl.RecordCount = 1 Then
                rstbl.Delete
                Call GProcActivateControls(Me, False)
                Call GProcSetButtons(Me, vbDataActionCancel)
                Form_Load
            End If
        End If
    'End If
Case 9 'Print
Case 10 'Exit
    Unload Me
Case 11 'Label
     If gCYear = 2023 Then
        Dim i As Long
        Dim X As String
        Dim tmp1 As Recordset
        X = "select max(ItBrkRtOld) from tblMastItem"
        i = GProcRstOpen(tmp1, X, "R", gCn)
      '  If tmp1.Fields(0) = 0 Then
'            gCn.Execute "Update tblMastItem set ItOldRtBefore =  CDate('01/05/2023') "
'            DoEvents
'            DoEvents
 '           gCn.Execute "Update tblMastItem set ItBrkRtOld = ItBrokRt"
            gCn.Execute "Update tblMastItem set ItBrokRt = 21 where val(itname) > 0 and val(itname) <= 59 "
            gCn.Execute "Update tblMastItem set ItBrokRt = 25 where val(itname) >= 60 and val(itname) <= 89"
            gCn.Execute "Update tblMastItem set ItBrokRt = 30 where val(itname) >= 90 "
            gCn.Execute "Update tblMastItem set ItBrokRt = 0.50 where ItBrokRt = 0.25"
            DoEvents
            DoEvents
            
'            gCn.Execute "Update tblMastItem set ITTMP3 = 1 where ItBrokRt = 20"
'            gCn.Execute "Update tblMastItem set ITTMP3 = 2 where ItBrokRt = 25"
'
'            gCn.Execute "Update tblMastItem set ItBrokRt = 25 where ItBrokRt = 20 and ITTMP3 = 1"
'            gCn.Execute "Update tblMastItem set ItBrokRt = 30 where ItBrokRt = 25 and ITTMP3 = 2"
            MsgBox ("Complited...")
        'End If
     End If
    'frmLabel.Show 1
Case 12 'Creation To Firm
    gfrmTypeStr = "Item"
    frmUtlCreationToFirm.Show 1
    rstbl.Requery
Case 13 'Save
    If ValidateData = True Then
        FillTxtFromTemp
        GProcSaveRecord Me, rstbl, FormAction, MaxNo
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
    Call GProcClearForm(Me, rstbl, MaxNo, True)
    If rstbl.RecordCount <= 0 Then
        Call GProcActivateControls(Me, False)
        Call GProcSetButtons(Me, vbDataActionCancel)
    Else
        If FormAction = vbDataActionAddNew Then
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
    FormAction = vbDataActionClose
End Select
Exit Sub
ErrorRoutine:
    rstbl.CancelUpdate
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
    Me.Top = 300
    Me.Left = Screen.Width / 2 - Me.Width / 2
    X = "Select * from " & tblName & " where ItOtherItem <> 1"
    If Trim(OrderField) <> "" Then X = X & " order by " & OrderField
    GProcRstOpen rstbl, X, "O"
    '---- Button Visiblity Set As Per User Previlage
    GProcSetButtonVisiblity Me
    FirstTimeFlag = True
    cmdBtn_Click (6)
    cbo(6).ListIndex = 0
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
If FormAction = vbDataActionUpdate Then
    MsgBox "You are about to modify the record. Save the record and then exit.", vbInformation + vbOKOnly, Me.Caption
    Cancel = True
    Exit Sub
End If
'mShowRec = Space(5)
rstbl.Requery
rstbl.Close
End Sub
Private Sub txt_GotFocus(Index As Integer)
GProcSelectBox Me.ActiveControl
End Sub
Private Sub txt_LostFocus(Index As Integer)
If Index = 1 Then
    Dim m As Integer
    For m = 1 To Len(txt(Index))
    If (Mid(txt(Index), m, 1)) = "'" Then
    txt(1) = Left(txt(Index), m - 1) + Mid(txt(Index), m + 1)
    End If
Next m
End If
GProcCheckForNumber Me, rstbl, MaxNo
SaveBtnEd
End Sub
Private Sub txt_KeyPress(Index As Integer, KeyAscii As Integer)
If KeyAscii = 13 Then Exit Sub
KeyAscii = GProcValidateKey(rstbl, Index, KeyAscii, txt(Index).text)
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
Dim MillCode As Long
TopPos = txttemp(Index).Top + Me.Top + 650
LeftPos = txttemp(Index).Left + Me.Left
If KeyAscii = 13 Then Exit Sub
If Index = 9 Then '-- Mill Name
    gClsSearch.SearchMultiField "QryMastaccount", "Account,AcName,AcCity,AgName", Array("Mill", "AcName", "City", "Group"), Array(3000, 0, 2000, 3000), " (AgCode=90029 or GpCode=90029)", Chr(KeyAscii), "AcName", LeftPos, TopPos, True
    KeyAscii = 0
    If gClsSearch.SearchMultiRetCol(0) <> "" And gClsSearch.SearchMultiRetCol(1) = "" Then
        txttemp(Index) = gClsSearch.SearchMultiRetCol(0)
    Else
        If gClsSearch.SearchMultiRetCol(1) <> "" Then txttemp(Index) = gClsSearch.SearchMultiRetCol(1)
    End If
    MillCode = GProcGetColumnValue("tblMastAccount", "AcName", txttemp(9), "S", "AcCode", "N")
    '-- Ticket
    If txt(2) = "" Then txt(2) = GProcGetColumnValue("tblMastItem", "ItMillCode", CStr(MillCode), "N", "ItTicket", "S")
End If
End Sub
Private Sub txttemp_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
Dim GName As String
If KeyCode = vbKeyF3 Then
    Select Case Index
        Case 9 '-- mill
            GName = GProcGetColumnValue("TblMastGroup", "AgCode", 90029, "N", "AgName", "S")
            GProcShowForm frmMastAccount, frmMain.mnuMstfrm(0), Array(frmMastAccount.txttemp(2)), Array(GName)
    End Select
End If
End Sub
Private Sub SaveBtnEd()
'--- Name,Ticket,Short,Unit,Mill
If Trim(txt(1)) = "" Or Trim(txt(2)) = "" Or Trim(txt(3)) = "" Or Trim(cbo(6)) = "" Or Trim(txttemp(9)) = "" Then
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
Dim MillCode As Long
Dim rstAddCheck As Recordset
Dim rstModiCheck As Recordset
    GProcCheckForNumber Me, rstbl, MaxNo
    '--- Required
    '-- ITCode
    If FormAction = vbDataActionAddNew Then txt(0) = GProcGenerateId(tblName, IdField, pStartCode:=gCSItCode)
    If gCEItCode > 0 And CLng(txt(0)) > gCEItCode Then
        MsgBox "Year End is completed. You can not create new count.", vbCritical + vbOKOnly, Me.Caption
        txt(1).SetFocus
        Exit Function
    End If
    '--- Name
    If Trim(txt(1)) = "" Then
        MsgBox "Check Item Name.", vbCritical + vbOKOnly, Me.Caption
        txt(1).SetFocus
        Exit Function
    End If
    '--- Ticket
    If Trim(txt(2)) = "" Then
        MsgBox "Check Item Ticket.", vbCritical + vbOKOnly, Me.Caption
        txt(2).SetFocus
        Exit Function
    End If
    '---Short
    If Trim(txt(3)) = "" Then
        MsgBox "Check Item Short.", vbCritical + vbOKOnly, Me.Caption
        txt(3).SetFocus
        Exit Function
    End If
    '---Unit
    If Trim(cbo(6)) = "" Then
        MsgBox "Check Unit.", vbCritical + vbOKOnly, Me.Caption
        cbo(6).SetFocus
        Exit Function
    End If
    '--- Mill
    If Trim(txttemp(9)) = "" Then
        MsgBox "Check Mill.", vbCritical + vbOKOnly, Me.Caption
        txttemp(9).SetFocus
        Exit Function
    End If
    '-- Create Mill
    If GProcGetColumnValue("tblMastAccount", "AcName", txttemp(9), "S", "AcCode", "N") = 0 Then
        GName = GProcGetColumnValue("TblMastGroup", "AgCode", 90029, "N", "AgName", "S")
        GProcShowForm frmMastAccount, frmMain.mnuMstfrm(0), Array(frmMastAccount.txt(1), frmMastAccount.txttemp(2)), Array(txttemp(9), GName)
        Exit Function
    End If
    '-- Max Rate
    If CDbl(txt(13)) < CDbl(txt(14)) Then
        MsgBox "Max Rate is less than Min Rate.", vbCritical + vbOKOnly, App.Title
        txt(13).SetFocus
        Exit Function
    End If
    '----Duplicate Records
    '----Addition Case
    MillCode = GProcGetColumnValue("TblMastAccount", "AcName", txttemp(9), "S", "Accode", "N")
    If FormAction = vbDataActionAddNew Then
        X = "Select ItCode,ItName From " & tblName & " Where ItMillCode=" & MillCode & " and ItName='" & Trim(txt(1).text) & "'"
        i = GProcRstOpen(rstAddCheck, X, "R")
        If i > 0 Then
            MsgBox "Record already exists  " & Trim(txt(1).text), vbCritical + vbOKOnly, "Duplicate Record"
            txt(1).SetFocus
            Exit Function
        End If
        rstAddCheck.Close
    End If
    '----Modification case
    If FormAction = vbDataActionUpdate Then
        X = "Select ItCode,ItName From " & tblName & " Where (ItMillCode=" & MillCode & " and (ItName='" & Trim(txt(1).text) & "') and (ItCode<>" & txt(0).text & "))"
        i = GProcRstOpen(rstModiCheck, X, "O")
        If i > 0 Then
            MsgBox "Record already exists  " & Trim(txt(1).text), vbCritical + vbOKOnly, "Duplicate Record"
            txt(1).SetFocus
            Exit Function
        End If
        rstModiCheck.Close
    End If
    ValidateData = True
End Function
Private Sub FillTxtFromTemp()
txt(6) = cbo(6)
txt(7) = cbo(7)
txt(17) = cbo(8)
txt(11) = chkIsIncludingGst.Value
txt(24) = Check1.Value
If cbo(22).ListIndex = -1 Then
    txt(22) = 0
Else
    txt(22) = cbo(22).ListIndex '-- Sl Type
End If
txt(9) = GProcGetColumnValue("TBLMASTaccount", "acname", txttemp(9), "S", "accode", "N", "", gCn)
If IsDate(mskDt) Then
    txt(16) = CDate(mskDt)
Else
    txt(16) = ""
End If
End Sub
Private Sub FillTempFromTxt()
cbo(6) = txt(6)
cbo(7) = txt(7)
cbo(8) = txt(17)
cbo(22).ListIndex = txt(22) '--SL Type
txttemp(9) = GProcGetColumnValue("TBLMASTaccount", "accode", txt(9), "N", "acname", "N", "", gCn)
mskDt = txt(16)
chkIsIncludingGst.Value = Val(txt(11))
Check1.Value = Val(txt(24))
End Sub
Private Sub ReadFields()
GProcShowRecord Me, rstbl, MaxNo
FillTempFromTxt
End Sub
'--- Is Releted Mill & Item Record
Private Function IsRelatedRecord() As Boolean
Dim X As String
Dim MillCode As Long
Dim Rs1 As Recordset
MillCode = GProcGetColumnValue("tblMastAccount", "AcName", txttemp(9), "S", "AcCode", "N")
X = "select tblAddLess.Vno,tblAddLess.Vtype,tblAddLess.Vdt from tblAddLess,tblVouSub where " _
& " tblAddLess.Vno=tblVouSub.Vno and tblAddLess.Vtype=tblVouSub.Vtype and tblAddLess.VYear=tblVouSub.VYear and " _
& " tblAddLess.AdMillCode = " & MillCode & " And tblVouSub.vSubItCode=" & txt(0)
GProcRstOpen Rs1, X, "R"
If Rs1.EOF And Rs1.BOF Then
Else
    IsRelatedRecord = True
End If
End Function
