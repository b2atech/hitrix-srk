VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form frmMastCompany 
   BackColor       =   &H00C2E3FC&
   BorderStyle     =   0  'None
   Caption         =   "Company Creation"
   ClientHeight    =   7185
   ClientLeft      =   0
   ClientTop       =   -45
   ClientWidth     =   9675
   Icon            =   "frmMastCompany.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7185
   ScaleWidth      =   9675
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox txt 
      Appearance      =   0  'Flat
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
      Left            =   9465
      MaxLength       =   3
      TabIndex        =   64
      Text            =   "20"
      Top             =   540
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.ComboBox cbo 
      Appearance      =   0  'Flat
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
      Index           =   19
      ItemData        =   "frmMastCompany.frx":058A
      Left            =   6330
      List            =   "frmMastCompany.frx":0597
      TabIndex        =   5
      Text            =   "cbo(19)"
      Top             =   615
      Width           =   1245
   End
   Begin VB.TextBox txt 
      Appearance      =   0  'Flat
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
      Left            =   45
      MaxLength       =   3
      TabIndex        =   63
      Text            =   "19"
      Top             =   1185
      Visible         =   0   'False
      Width           =   210
   End
   Begin VB.TextBox txt 
      Appearance      =   0  'Flat
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
      Left            =   5280
      MaxLength       =   3
      TabIndex        =   33
      Text            =   "18"
      Top             =   6480
      Width           =   2295
   End
   Begin VB.TextBox txt 
      Appearance      =   0  'Flat
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
      Left            =   1800
      MaxLength       =   3
      TabIndex        =   31
      Text            =   "17"
      Top             =   6480
      Width           =   2415
   End
   Begin VB.TextBox txt 
      Appearance      =   0  'Flat
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
      Index           =   15
      Left            =   1800
      MaxLength       =   3
      TabIndex        =   27
      Text            =   "15"
      Top             =   5985
      Width           =   2415
   End
   Begin VB.TextBox txt 
      Appearance      =   0  'Flat
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
      Left            =   5280
      MaxLength       =   3
      TabIndex        =   29
      Text            =   "16"
      Top             =   5985
      Width           =   2295
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
      Left            =   8085
      Style           =   1  'Graphical
      TabIndex        =   44
      ToolTipText     =   "Previous"
      Top             =   2025
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
      Left            =   8685
      Style           =   1  'Graphical
      TabIndex        =   45
      ToolTipText     =   "Next"
      Top             =   2025
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
      Left            =   8085
      Style           =   1  'Graphical
      TabIndex        =   46
      ToolTipText     =   "First"
      Top             =   2505
      Width           =   495
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
      Left            =   8685
      Style           =   1  'Graphical
      TabIndex        =   47
      ToolTipText     =   "Last"
      Top             =   2505
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
      Left            =   8085
      Style           =   1  'Graphical
      TabIndex        =   48
      ToolTipText     =   "To Find Record"
      Top             =   2985
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
      Left            =   9675
      Style           =   1  'Graphical
      TabIndex        =   49
      Top             =   2430
      Visible         =   0   'False
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
      Left            =   8085
      Style           =   1  'Graphical
      TabIndex        =   39
      ToolTipText     =   "To Add New Record"
      Top             =   3630
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
      Left            =   8085
      Style           =   1  'Graphical
      TabIndex        =   40
      ToolTipText     =   "To Modify Record"
      Top             =   4110
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
      Left            =   8085
      Style           =   1  'Graphical
      TabIndex        =   41
      ToolTipText     =   "To Delete Record"
      Top             =   4590
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
      Left            =   6120
      Style           =   1  'Graphical
      TabIndex        =   50
      ToolTipText     =   "To Add New Record"
      Top             =   7290
      Visible         =   0   'False
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
      Left            =   8085
      Style           =   1  'Graphical
      TabIndex        =   42
      ToolTipText     =   "To Save Record"
      Top             =   5265
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
      Left            =   8085
      Style           =   1  'Graphical
      TabIndex        =   43
      ToolTipText     =   "To Cancel Record"
      Top             =   5745
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
      Left            =   8085
      Style           =   1  'Graphical
      TabIndex        =   53
      ToolTipText     =   "Exit"
      Top             =   6315
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
      Left            =   5280
      Style           =   1  'Graphical
      TabIndex        =   51
      ToolTipText     =   "To Add New Record"
      Top             =   7395
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
      Left            =   4320
      Style           =   1  'Graphical
      TabIndex        =   52
      ToolTipText     =   "To Add New Record"
      Top             =   7395
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.CheckBox Chk 
      Alignment       =   1  'Right Justify
      BackColor       =   &H80000018&
      Caption         =   "Hide"
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
      Index           =   14
      Left            =   4320
      TabIndex        =   38
      Top             =   5550
      Width           =   3135
   End
   Begin VB.ComboBox cbo 
      Appearance      =   0  'Flat
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
      Left            =   5280
      TabIndex        =   37
      Text            =   "cbo(13)"
      Top             =   4665
      Width           =   2295
   End
   Begin VB.ComboBox cbo 
      Appearance      =   0  'Flat
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
      Left            =   1800
      TabIndex        =   15
      Text            =   "cbo(6)"
      Top             =   3330
      Width           =   5775
   End
   Begin VB.ComboBox cbo 
      Appearance      =   0  'Flat
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
      Left            =   1800
      TabIndex        =   13
      Text            =   "cbo(5)"
      Top             =   2910
      Width           =   5775
   End
   Begin VB.ComboBox cbo 
      Appearance      =   0  'Flat
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
      Left            =   1800
      TabIndex        =   11
      Text            =   "cbo(4)"
      Top             =   2490
      Width           =   5775
   End
   Begin VB.ComboBox cbo 
      Appearance      =   0  'Flat
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
      Index           =   3
      Left            =   1800
      TabIndex        =   10
      Text            =   "cbo(3)"
      Top             =   2070
      Width           =   5775
   End
   Begin VB.ComboBox cbo 
      Appearance      =   0  'Flat
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
      Index           =   2
      Left            =   1800
      TabIndex        =   9
      Text            =   "cbo(2)"
      Top             =   1665
      Width           =   5775
   End
   Begin VB.TextBox txt 
      Appearance      =   0  'Flat
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
      Left            =   9480
      MaxLength       =   3
      TabIndex        =   60
      Text            =   "14"
      Top             =   3315
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.TextBox txt 
      Appearance      =   0  'Flat
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
      Left            =   9480
      MaxLength       =   3
      TabIndex        =   59
      Text            =   "13"
      Top             =   3795
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.TextBox txt 
      Appearance      =   0  'Flat
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
      Left            =   1800
      MaxLength       =   3
      TabIndex        =   17
      Text            =   "12"
      Top             =   3780
      Width           =   5775
   End
   Begin VB.TextBox txt 
      Appearance      =   0  'Flat
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
      Left            =   5280
      MaxLength       =   3
      TabIndex        =   35
      Text            =   "11"
      Top             =   4215
      Width           =   2295
   End
   Begin VB.TextBox txt 
      Appearance      =   0  'Flat
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
      Left            =   1800
      MaxLength       =   3
      TabIndex        =   25
      Text            =   "10"
      Top             =   5550
      Width           =   2415
   End
   Begin VB.TextBox txt 
      Appearance      =   0  'Flat
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
      Left            =   1800
      MaxLength       =   3
      TabIndex        =   23
      Text            =   "9"
      Top             =   5115
      Width           =   2415
   End
   Begin VB.TextBox txt 
      Appearance      =   0  'Flat
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
      Left            =   1800
      MaxLength       =   3
      TabIndex        =   21
      Text            =   "8"
      Top             =   4665
      Width           =   2415
   End
   Begin VB.TextBox txt 
      Appearance      =   0  'Flat
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
      Left            =   1800
      MaxLength       =   3
      TabIndex        =   19
      Text            =   "7"
      Top             =   4215
      Width           =   2415
   End
   Begin VB.TextBox txt 
      Appearance      =   0  'Flat
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
      Left            =   9480
      MaxLength       =   3
      TabIndex        =   58
      Text            =   "6"
      Top             =   2835
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.TextBox txt 
      Appearance      =   0  'Flat
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
      Left            =   9480
      MaxLength       =   3
      TabIndex        =   57
      Text            =   "5"
      Top             =   2355
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.TextBox txt 
      Appearance      =   0  'Flat
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
      Left            =   9480
      MaxLength       =   3
      TabIndex        =   56
      Text            =   "4"
      Top             =   1875
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.TextBox txt 
      Appearance      =   0  'Flat
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
      Index           =   3
      Left            =   9480
      MaxLength       =   3
      TabIndex        =   55
      Text            =   "3"
      Top             =   1395
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.TextBox txt 
      Appearance      =   0  'Flat
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
      Index           =   2
      Left            =   9480
      MaxLength       =   3
      TabIndex        =   54
      Text            =   "2"
      Top             =   915
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.TextBox txt 
      Appearance      =   0  'Flat
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
      Index           =   1
      Left            =   1800
      MaxLength       =   3
      TabIndex        =   7
      Text            =   "1"
      Top             =   1050
      Width           =   5775
   End
   Begin VB.TextBox txt 
      Appearance      =   0  'Flat
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
      Left            =   1800
      MaxLength       =   3
      TabIndex        =   1
      Text            =   "0"
      Top             =   600
      Width           =   825
   End
   Begin MSMask.MaskEdBox mskDt 
      Height          =   360
      Index           =   20
      Left            =   3885
      TabIndex        =   3
      Top             =   600
      Width           =   1170
      _ExtentX        =   2064
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
   Begin VB.Label Label19 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Start Date :"
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
      Left            =   2715
      TabIndex        =   2
      Top             =   600
      Width           =   1065
   End
   Begin VB.Label Label18 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Designation :"
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
      Left            =   5145
      TabIndex        =   4
      Top             =   615
      Width           =   1140
   End
   Begin VB.Image Image1 
      Height          =   1245
      Left            =   7875
      Picture         =   "frmMastCompany.frx":05BA
      Stretch         =   -1  'True
      Top             =   495
      Width           =   1455
   End
   Begin VB.Shape Shape4 
      BorderColor     =   &H00C00000&
      BorderWidth     =   5
      FillColor       =   &H00C00000&
      Height          =   6810
      Left            =   0
      Top             =   360
      Width           =   9675
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
      Left            =   6855
      TabIndex        =   61
      Top             =   30
      Width           =   2475
   End
   Begin VB.Label Label16 
      BackStyle       =   0  'Transparent
      Caption         =   "AC/ STO No :"
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
      Height          =   540
      Left            =   4320
      TabIndex        =   32
      Top             =   6480
      Width           =   1020
      WordWrap        =   -1  'True
   End
   Begin VB.Label Label10 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "File No :"
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
      Left            =   360
      TabIndex        =   30
      Top             =   6480
      Width           =   735
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "E-mail Id :"
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
      Left            =   360
      TabIndex        =   26
      Top             =   5985
      Width           =   885
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "STC No :"
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
      Left            =   4320
      TabIndex        =   28
      Top             =   5985
      Width           =   795
   End
   Begin VB.Shape Shape1 
      FillColor       =   &H00BDDFD1&
      FillStyle       =   0  'Solid
      Height          =   5175
      Left            =   7875
      Shape           =   4  'Rounded Rectangle
      Top             =   1830
      Width           =   1515
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "TDS  :"
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
      Left            =   4320
      TabIndex        =   34
      Top             =   4215
      Width           =   555
   End
   Begin VB.Label Label9 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "PAN :"
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
      Left            =   360
      TabIndex        =   24
      Top             =   5550
      Width           =   525
   End
   Begin VB.Label Label11 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "TAN  :"
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
      Left            =   360
      TabIndex        =   22
      Top             =   5115
      Width           =   555
   End
   Begin VB.Label Label12 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "CST &TIN  :"
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
      Left            =   360
      TabIndex        =   18
      Top             =   4215
      Width           =   900
   End
   Begin VB.Label Label13 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "VAT TIN  :"
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
      Left            =   360
      TabIndex        =   20
      Top             =   4665
      Width           =   900
   End
   Begin VB.Label Label15 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "TDS Cir  :"
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
      Left            =   4320
      TabIndex        =   36
      Top             =   4665
      Width           =   855
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Add&ress    :"
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
      Left            =   360
      TabIndex        =   8
      Top             =   1665
      Width           =   1020
   End
   Begin VB.Label Label14 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Short &Name  :"
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
      Top             =   600
      Width           =   1305
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Capital Name  : "
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
      Left            =   360
      TabIndex        =   16
      Top             =   3780
      Width           =   1410
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "&Phone     :"
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
      Left            =   360
      TabIndex        =   12
      Top             =   2910
      Width           =   915
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Bill &Heading  :"
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
      Left            =   360
      TabIndex        =   14
      Top             =   3330
      Width           =   1215
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Name      :"
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
      TabIndex        =   6
      Top             =   1050
      Width           =   975
   End
   Begin VB.Label Label17 
      BackColor       =   &H00C00000&
      Caption         =   "   Company Creation"
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
      TabIndex        =   62
      Top             =   0
      Width           =   10455
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H00000080&
      FillColor       =   &H00EDFDFE&
      FillStyle       =   0  'Solid
      Height          =   1020
      Left            =   165
      Top             =   480
      Width           =   7530
   End
   Begin VB.Shape Shape3 
      BorderColor     =   &H00000080&
      FillColor       =   &H00EDFDFE&
      FillStyle       =   0  'Solid
      Height          =   5490
      Left            =   165
      Top             =   1560
      Width           =   7530
   End
End
Attribute VB_Name = "frmMastCompany"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Const tblName As String = "tblMastCompany"
Const IdField As String = "CCode"
Const OrderField As String = "CName"
Const MaxNo As Long = 20
Dim rstbl As Recordset
Dim varBookMark As Variant
Dim FirstTimeFlag As Boolean
Dim FormAction As Integer

Private Sub Cbo_KeyPress(Index As Integer, KeyAscii As Integer)
If Index = 19 Then KeyAscii = 0 '-Designation
End Sub

Private Sub cmdBtn_Click(Index As Integer)
On Error GoTo ErrorRoutine
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
    gClsSearch.SearchMultiField tblName, "CName", Array("Company Name"), Array(txt(1).Width), "", "", "CName", txt(1).Left + Me.Left, txt(1).Top + Me.Top + 650, False, gSetupCn
    If gClsSearch.SearchMultiRetCol(0) <> "" Then
        rstbl.MoveFirst
        rstbl.Find "CName= '" & gClsSearch.SearchMultiRetCol(0) & "'"
        gClsSearch.SearchMultiRetCol(0) = ""
        ReadFields
    End If
Case 5 'List
    'GenMastList ("LAC")
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
    mskDt(20) = Format("01/04/" & Year(Date), "dd/mm/yyyy") '- Start Date
    FillCombo
Case 7 'Modify
    'If mURecModi = True Then
    '    SearchHelp
    '    DoEvents
        FormAction = vbDataActionUpdate
        Call GProcActivateControls(Me, True)
        Call GProcSetButtons(Me, vbDataActionUpdate)
        txt(0).Enabled = False '-- Short Name
        txt(1).SetFocus
    'End If
Case 8 'Delete
    'If mURecDel = True Then
        If MsgBox("Do You Really Delete This Record?", vbQuestion + vbYesNo + vbDefaultButton1, "Confirm Delete") = vbYes Then
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
    'frmLabel.Show 1
Case 12 'Optional
Case 13 'Save
    If ValidateData = True Then
           ' If FormAction = vbDataActionAddNew Then
        FillTxtFromTemp
        GProcSaveRecord Me, rstbl, FormAction, MaxNo
        '-- Create Company Database
        CreateCmpStartDtDb
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

Private Sub Form_Load()
    Dim X As String
    Dim Y As String
    Me.Top = 300
    Me.Left = Screen.Width / 2 - Me.Width / 2
    Y = "Select * from " & tblName & " order by " & OrderField
    GProcRstOpen rstbl, Y, "O", gSetupCn
    'MaxNo = rstbl.Fields.Count - 1
    FirstTimeFlag = True
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
'mShowRec = Space(5)
rstbl.Close
End Sub
Private Sub mskDt_GotFocus(Index As Integer)
GProcSelectBox Me.ActiveControl
End Sub
Private Sub mskDt_LostFocus(Index As Integer)
If IsDate(mskDt(Index)) = False Then mskDt(Index) = "__/__/____" '- Start Date
SaveBtnEd
End Sub
Private Sub txt_GotFocus(Index As Integer)
GProcSelectBox Me.ActiveControl
End Sub
Private Sub txt_LostFocus(Index As Integer)
CheckForNumber
SaveBtnEd
End Sub
Private Sub CheckForNumber()
GProcCheckForNumber Me, rstbl, MaxNo, Array(7, 8, 9, 10, 11, 15, 16, 17, 18)
End Sub
Private Sub txt_KeyPress(Index As Integer, KeyAscii As Integer)
KeyAscii = GProcValidateKey(rstbl, Index, KeyAscii, txt(Index).Text)
End Sub
Private Sub txttemp_GotFocus(Index As Integer)
GProcSelectBox Me.ActiveControl
End Sub

Private Sub txtTemp_KeyPress(Index As Integer, KeyAscii As Integer)
End Sub

Private Sub txtTemp_LostFocus(Index As Integer)
SaveBtnEd
End Sub

Private Sub FillCombo()
Call GProcFillCombo(cbo(2), "select Distinct CAdd1 from " & tblName & " where CAdd1 <> ''", gSetupCn)
Call GProcFillCombo(cbo(3), "select Distinct CAdd2 from " & tblName & " where CAdd2 <> ''", gSetupCn)
Call GProcFillCombo(cbo(4), "select Distinct CAdd3 from " & tblName & " where CAdd3 <> ''", gSetupCn)
Call GProcFillCombo(cbo(5), "select Distinct CPh from " & tblName & " where CPh <> ''", gSetupCn)
Call GProcFillCombo(cbo(6), "select Distinct CHead from " & tblName & " where CHead <> ''", gSetupCn)
Call GProcFillCombo(cbo(13), "select Distinct CTDSCir from " & tblName & " where CTDSCir <> ''", gSetupCn)
End Sub
Private Sub SaveBtnEd()
'--- Short Name,Name
If txt(0) = "" Or txt(1) = "" Or IsDate(mskDt(20)) = False Then
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
Dim rstAddCheck As Recordset
Dim rstModiCheck As Recordset
    '---- Required Fields
        '--- Short Name
        If Trim(txt(0)) = "" Then
            MsgBox "Enter Company Short Name.", vbCritical + vbOKOnly, Me.Caption
            txt(0).SetFocus
            Exit Function
        End If
        '-- Start Date
        If IsDate(mskDt(20)) = False Then
            MsgBox "Enter Company Start Date.", vbCritical + vbOKOnly, Me.Caption
            mskDt(20).SetFocus
            Exit Function
        End If
        If Format(mskDt(20), "dd/mm") <> "01/04" Then
            MsgBox "Check Company Start Date. Date/Month must be 01/04", vbCritical + vbOKOnly, Me.Caption
            mskDt(20).SetFocus
            Exit Function
        End If
        '--- Name
        If Trim(txt(1)) = "" Then
            MsgBox "Enter Company Name.", vbCritical + vbOKOnly, Me.Caption
            txt(1).SetFocus
            Exit Function
        End If
    '----Duplicate Records
    '----Addition Case
    If FormAction = vbDataActionAddNew Then
        X = "Select CCode From " & tblName & " Where CCode='" & Trim(txt(0).Text) & "'"
        i = GProcRstOpen(rstAddCheck, X, "R", gSetupCn)
        If i > 0 Then
            MsgBox "Record is Available of Company Short Name." & Trim(txt(0).Text), vbCritical + vbOKOnly, "Duplicate Record"
            txt(0).SetFocus
            Exit Function
        End If
        rstAddCheck.Close
    End If
    '--- Start Year Confirmation For Database Creation
    If MsgBox("Company will get create for start date " & mskDt(20) & ". Are you sure?", vbInformation + vbYesNoCancel) <> vbYes Then
        mskDt(20).SetFocus
        Exit Function
    End If
    '----Modification case
'    If FormAction = vbDataActionUpdate Then
'        x = "Select compCode From " & tblName & " Where ((CompCode='" & Trim(txt(0).Text) & "') and (ACCODE<>" & txt(0).Text & "))"
'        i = GProcRstOpen(gCn, rstModiCheck, x, "O")
'        If i > 0 Then
'            MsgBox "Record is Available of A/c Name  " & Trim(txt(1).Text), vbCritical + vbOKOnly, "Duplicate Record"
'            txt(1).SetFocus
'            Exit Function
'        End If
'        rstModiCheck.Close
'    End If
    ValidateData = True
End Function
Private Sub FillTxtFromTemp()
Dim i As Long
For i = 2 To 6
    txt(i) = cbo(i)
Next i
txt(13) = cbo(13)
txt(14) = Chk(14)
txt(19) = cbo(19).ListIndex
If IsDate(mskDt(20)) = True Then '- Start Dt
    txt(20) = CDate(mskDt(20))
Else
    txt(20) = ""
End If
End Sub
Private Sub FillTempFromTxt()
Dim i As Long
For i = 2 To 6
    cbo(i) = txt(i)
Next i
cbo(13) = txt(13)
Chk(14) = txt(14)
cbo(19).ListIndex = txt(19)
If IsDate(txt(20)) = True Then '- Start Date
    mskDt(20) = CDate(txt(20))
Else
    mskDt(20) = "__/__/____"
End If
End Sub
Private Sub ReadFields()
GProcShowRecord Me, rstbl, MaxNo
FillTempFromTxt
End Sub
Private Sub CreateCmpStartDtDb()
Dim CYr As String

If FormAction = vbDataActionAddNew And IsDate(mskDt(20)) = True Then
    CYr = Format(Format(mskDt(20), "yy"), "00")
    CYr = CYr + Format(CLng(CYr) + 1, "00")
    '-- Check Is Year Entry already Exist
    If GProcGetColumnValue("tblCompanyYear", "CompCode", txt(0), "S", "CompYear", "S", "CompFdt=cdate('" & mskDt(20) & "')", gSetupCn) = "" Then
        '- Insert Year Entry
        gSetupCn.Execute "insert into tblCompanyYear (CompCode,CompYear,CompFdt,CompTdt) values " _
        & " ('" & txt(0) & "','" & CYr & "',cdate('" & mskDt(20) & "'),cdate('" & (DateAdd("yyyy", 1, mskDt(20)) - 1) & "'))"
        '- Create Database
        If gBackEndDB = gBackEndAccess Then FileCopy gDataPath & "\Hitrix.mdb", gDataPath & "\" & txt(0) & CYr & ".mdb"
    End If
End If
End Sub

