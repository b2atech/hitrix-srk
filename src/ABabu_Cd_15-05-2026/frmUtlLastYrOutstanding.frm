VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form frmUtlLastYrOutstanding 
   BackColor       =   &H00C2E3FC&
   BorderStyle     =   0  'None
   Caption         =   " "
   ClientHeight    =   6300
   ClientLeft      =   0
   ClientTop       =   60
   ClientWidth     =   9780
   FillColor       =   &H000040C0&
   FillStyle       =   0  'Solid
   Icon            =   "frmUtlLastYrOutstanding.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6300
   ScaleWidth      =   9780
   ShowInTaskbar   =   0   'False
   Begin Crystal.CrystalReport CryReport 
      Left            =   4140
      Top             =   2775
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
   Begin VB.ComboBox cbo 
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
      ItemData        =   "frmUtlLastYrOutstanding.frx":058A
      Left            =   1950
      List            =   "frmUtlLastYrOutstanding.frx":059D
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   720
      Width           =   1980
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
      Left            =   3750
      TabIndex        =   58
      Text            =   "21"
      Top             =   6030
      Visible         =   0   'False
      Width           =   375
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
      Left            =   4515
      TabIndex        =   57
      Text            =   "20"
      Top             =   6030
      Visible         =   0   'False
      Width           =   375
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
      Left            =   4110
      TabIndex        =   56
      Text            =   "19"
      Top             =   6000
      Visible         =   0   'False
      Width           =   375
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
      Left            =   4890
      TabIndex        =   55
      Text            =   "18"
      Top             =   6030
      Visible         =   0   'False
      Width           =   375
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
      Index           =   17
      Left            =   5280
      TabIndex        =   54
      Text            =   "17"
      Top             =   6000
      Visible         =   0   'False
      Width           =   375
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
      Index           =   16
      Left            =   5640
      TabIndex        =   53
      Text            =   "16"
      Top             =   6030
      Visible         =   0   'False
      Width           =   375
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
      Index           =   15
      Left            =   6000
      TabIndex        =   52
      Text            =   "15"
      Top             =   6030
      Visible         =   0   'False
      Width           =   375
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
      Left            =   1950
      TabIndex        =   11
      Text            =   "txttemp(5)"
      ToolTipText     =   "Press F3 for Creation"
      Top             =   2355
      Width           =   5535
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
      Height          =   360
      Index           =   14
      Left            =   6030
      TabIndex        =   19
      Text            =   "14"
      ToolTipText     =   "Press F3 For Creation"
      Top             =   3915
      Width           =   1455
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
      Index           =   13
      Left            =   1950
      TabIndex        =   21
      Text            =   "8"
      Top             =   4680
      Width           =   1980
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
      Height          =   360
      Index           =   12
      Left            =   1950
      TabIndex        =   17
      Text            =   "12"
      ToolTipText     =   "Press F3 For Creation"
      Top             =   3915
      Width           =   1980
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
      Index           =   11
      Left            =   6690
      TabIndex        =   51
      Text            =   "11"
      Top             =   6030
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
      Index           =   10
      Left            =   7830
      TabIndex        =   50
      Text            =   "10"
      Top             =   5970
      Visible         =   0   'False
      Width           =   315
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
      Index           =   8
      Left            =   6030
      TabIndex        =   28
      Text            =   "9"
      Top             =   5145
      Width           =   1455
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
      Height          =   360
      Index           =   4
      Left            =   1950
      TabIndex        =   9
      Text            =   "4"
      Top             =   1140
      Width           =   1980
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
      Index           =   1
      Left            =   2520
      TabIndex        =   49
      Text            =   "1"
      Top             =   6090
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
      Height          =   375
      Index           =   6
      Left            =   1950
      TabIndex        =   13
      Text            =   "txttemp(6)"
      ToolTipText     =   "Press F3 for Creation"
      Top             =   2820
      Width           =   5535
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
      Height          =   360
      Index           =   5
      Left            =   7440
      TabIndex        =   48
      Text            =   "5"
      Top             =   5970
      Visible         =   0   'False
      Width           =   360
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
      Index           =   7
      Left            =   6030
      TabIndex        =   23
      Text            =   "7"
      Top             =   4680
      Width           =   1455
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
      Left            =   7050
      TabIndex        =   46
      Text            =   "6"
      Top             =   6000
      Visible         =   0   'False
      Width           =   315
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
      Index           =   11
      Left            =   1950
      TabIndex        =   15
      Text            =   "txttemp(11)"
      ToolTipText     =   "Press F3 For Creation"
      Top             =   3360
      Width           =   5535
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
      Index           =   9
      Left            =   1950
      TabIndex        =   25
      Text            =   "9"
      Top             =   5145
      Width           =   720
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
      Height          =   360
      Index           =   3
      Left            =   6390
      TabIndex        =   41
      Text            =   "3"
      Top             =   6030
      Visible         =   0   'False
      Width           =   285
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
      Index           =   2
      Left            =   6030
      TabIndex        =   7
      Text            =   "2"
      Top             =   1590
      Width           =   1455
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
      Left            =   6000
      Style           =   1  'Graphical
      TabIndex        =   44
      ToolTipText     =   "Optional"
      Top             =   7200
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
      Left            =   4800
      Style           =   1  'Graphical
      TabIndex        =   43
      ToolTipText     =   "Label"
      Top             =   7200
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
      Left            =   8280
      Style           =   1  'Graphical
      TabIndex        =   40
      ToolTipText     =   "Exit from Form"
      Top             =   5520
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
      Left            =   8280
      Style           =   1  'Graphical
      TabIndex        =   33
      ToolTipText     =   "Cancel"
      Top             =   5040
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
      Left            =   8280
      Style           =   1  'Graphical
      TabIndex        =   32
      ToolTipText     =   "Save Record"
      Top             =   4560
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
      Left            =   7200
      Style           =   1  'Graphical
      TabIndex        =   42
      ToolTipText     =   "Print Record"
      Top             =   7320
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
      Left            =   8280
      Style           =   1  'Graphical
      TabIndex        =   31
      ToolTipText     =   "Delete Record"
      Top             =   3960
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
      Left            =   8280
      Style           =   1  'Graphical
      TabIndex        =   30
      ToolTipText     =   " Modify Record"
      Top             =   3480
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
      Left            =   8280
      Style           =   1  'Graphical
      TabIndex        =   29
      ToolTipText     =   " Add New Record"
      Top             =   3000
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
      Left            =   8280
      Style           =   1  'Graphical
      TabIndex        =   39
      ToolTipText     =   "List Of Record"
      Top             =   2280
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
      Left            =   8280
      Style           =   1  'Graphical
      TabIndex        =   34
      ToolTipText     =   "Find Record"
      Top             =   1800
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
      Left            =   8880
      Style           =   1  'Graphical
      TabIndex        =   38
      ToolTipText     =   "Last"
      Top             =   1320
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
      Left            =   8280
      Style           =   1  'Graphical
      TabIndex        =   37
      ToolTipText     =   "First"
      Top             =   1320
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
      Left            =   8880
      Style           =   1  'Graphical
      TabIndex        =   36
      ToolTipText     =   "Next"
      Top             =   840
      Width           =   495
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
      Left            =   6030
      TabIndex        =   3
      Text            =   " 0"
      Top             =   720
      Width           =   1455
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
      Left            =   8280
      Style           =   1  'Graphical
      TabIndex        =   35
      ToolTipText     =   "Previous"
      Top             =   840
      Width           =   495
   End
   Begin MSMask.MaskEdBox mskDt 
      Height          =   360
      Index           =   3
      Left            =   6030
      TabIndex        =   5
      Top             =   1170
      Width           =   1455
      _ExtentX        =   2566
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
      Index           =   10
      Left            =   2745
      TabIndex        =   26
      Top             =   5145
      Width           =   1185
      _ExtentX        =   2090
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
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "&Amount            :"
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
      Index           =   11
      Left            =   4290
      TabIndex        =   22
      Top             =   4680
      Width           =   1515
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Bill No             :"
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
      Index           =   14
      Left            =   390
      TabIndex        =   8
      Top             =   1170
      Width           =   1470
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Debit Note Amt  :"
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
      Left            =   4290
      TabIndex        =   27
      Top             =   5145
      Width           =   1515
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Ba&g               :"
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
      Index           =   12
      Left            =   390
      TabIndex        =   20
      Top             =   4680
      Width           =   1305
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "&Due Day        :"
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
      Index           =   8
      Left            =   390
      TabIndex        =   24
      Top             =   5145
      Width           =   1290
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "&Broker          :"
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
      Left            =   390
      TabIndex        =   12
      Top             =   2820
      Width           =   1290
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "&Party             :"
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
      Left            =   390
      TabIndex        =   10
      Top             =   2355
      Width           =   1335
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Voucher Year   :"
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
      Left            =   4290
      TabIndex        =   6
      Top             =   1590
      Width           =   1530
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Voucher &Date   :"
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
      Index           =   6
      Left            =   4290
      TabIndex        =   4
      Top             =   1170
      Width           =   1530
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Invoice &Type   :"
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
      Index           =   5
      Left            =   390
      TabIndex        =   0
      Top             =   720
      Width           =   1455
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Voucher &No      :"
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
      Left            =   4290
      TabIndex        =   2
      Top             =   720
      Width           =   1530
   End
   Begin VB.Shape Shape2 
      BackColor       =   &H00808080&
      BorderColor     =   &H004D177D&
      FillColor       =   &H00ECFFFF&
      FillStyle       =   0  'Solid
      Height          =   1515
      Index           =   1
      Left            =   180
      Top             =   570
      Width           =   7635
   End
   Begin VB.Shape Shape6 
      BorderColor     =   &H00C00000&
      BorderWidth     =   5
      FillColor       =   &H00C00000&
      Height          =   6270
      Left            =   0
      Top             =   15
      Width           =   9780
   End
   Begin VB.Shape Shape5 
      BorderColor     =   &H000000C0&
      FillColor       =   &H000000C0&
      FillStyle       =   0  'Solid
      Height          =   45
      Left            =   0
      Top             =   6255
      Width           =   9735
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
      Left            =   6975
      TabIndex        =   47
      Top             =   45
      Width           =   2475
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "&Mill               :"
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
      Left            =   390
      TabIndex        =   14
      Top             =   3360
      Width           =   1305
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Short Name  :"
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
      Index           =   4
      Left            =   390
      TabIndex        =   16
      Top             =   3900
      Width           =   1215
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Ticket Name    :"
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
      Index           =   3
      Left            =   4290
      TabIndex        =   18
      Top             =   3900
      Width           =   1395
   End
   Begin VB.Label Label6 
      BackColor       =   &H00C00000&
      Caption         =   " Last Year Outstanding"
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
      TabIndex        =   45
      Top             =   0
      Width           =   9735
   End
   Begin VB.Shape Shape1 
      FillColor       =   &H00BDDFD1&
      FillStyle       =   0  'Solid
      Height          =   5415
      Left            =   7995
      Shape           =   4  'Rounded Rectangle
      Top             =   675
      Width           =   1575
   End
   Begin VB.Shape Shape2 
      BackColor       =   &H00808080&
      BorderColor     =   &H004D177D&
      FillColor       =   &H00ECFFFF&
      FillStyle       =   0  'Solid
      Height          =   2295
      Index           =   0
      Left            =   180
      Top             =   2190
      Width           =   7635
   End
   Begin VB.Shape Shape4 
      BorderColor     =   &H00000080&
      FillColor       =   &H00EDFDFE&
      FillStyle       =   0  'Solid
      Height          =   1275
      Left            =   180
      Top             =   4560
      Width           =   7635
   End
End
Attribute VB_Name = "frmUtlLastYrOutstanding"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Const tblName As String = "tblOutStanding"
Const IdField As String = "vno"
Const OrderField As String = "vdt,vno"
Const MaxNo As Long = 21
Dim ConditionStr As String
Dim rstbl As Recordset
Dim varBookMark As Variant
Dim FirstTimeFlag As Boolean
Dim FormAction As Integer

Private Sub Cbo_KeyPress(Index As Integer, KeyAscii As Integer)
If KeyAscii = 13 Then Exit Sub
'-- Inv Type
If Index = 1 Then KeyAscii = 0
End Sub
Private Sub cbo_LostFocus(Index As Integer)
If Index = 1 And FormAction = vbDataActionAddNew Then  '--Inv Type
    txt(0) = GProcGenerateId(tblName, "Vno", "Vtype='" & GetVType & "' and VYear=" & gCYear - 1)
    mskDt(3) = GProcGenerateVDt(tblName, "Vdt", "Vtype='" & GetVType & "' and VYear=" & gCYear - 1)
End If
End Sub
Private Function GetVType() As String
Select Case Cbo(1).ListIndex
    Case 0 '--Trade
        GetVType = "SY"
    Case 1 '--Consign
        GetVType = "SO"
    Case 2 '--Depot
        GetVType = "SD"
    Case 3 '--SIT
        GetVType = "ST"
    Case 4 '-- Mill Bill
        GetVType = "SM"
End Select
End Function

Private Sub cmdBtn_Click(Index As Integer)
On Error GoTo ErrorRoutine
Dim FindStr As String
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
    If gBackEndDB = gBackEndAccess Then '--Access
        gClsSearch.SearchMultiField "tblOutstanding,tblMastAccount", "VNo,Vtype,iif(VType='SY','Trade',iif(Vtype='SO','Consignment',iif(Vtype='SD','Depot',iif(Vtype='ST','Sale In Transit','Mill Bill') ))),VDt,OutBillNo,AcName,OutAmount,OutItShort,OutTicket", Array("VNo", "Vtype", "Type", "Date", "Bill No", "Party", "Amount", "Short", "Ticket"), Array(800, 0, 1500, 1200, 1000, 1800, 1200, 1000, 1000), "  outaccode=AcCode and Vyear=" & gCYear - 1 & " and vtype in ('SO','SY','ST','SD','SM') ", "", "Vdt,Vno", 0, 0
    Else '--Oracle
        gClsSearch.SearchMultiField "tblOutstanding,tblMastAccount", "VNo,Vtype,decode(VType,'SY','Trade','SO','Consignment','SD','Depot','ST','Sale In Transit','Mill Bill'),VDt,OutBillNo,AcName,OutAmount,OutItShort,OutTicket", Array("VNo", "Vtype", "Type", "Date", "Bill No", "Party", "Amount", "Short", "Ticket"), Array(800, 0, 1500, 1200, 1000, 1800, 1200, 1000, 1000), "  outaccode=AcCode and Vyear=" & gCYear - 1 & " and vtype in ('SO','SY','ST','SD','SM') ", "", "Vdt,Vno", 0, 0
    End If
    If gClsSearch.SearchMultiRetCol(1) <> "" Then
        rstbl.CancelUpdate
        rstbl.MoveFirst
        FindStr = gClsSearch.SearchMultiRetCol(0) & "-" & gClsSearch.SearchMultiRetCol(1)
        rstbl.Find "VnoVType='" & FindStr & "'"
        gClsSearch.SearchMultiRetCol(0) = ""
        ReadFields
    End If
Case 5 'List
        On Error GoTo ErrorHandler '--- Outstanding List
        GProcCrystalRptPreparation cryReport, "Window"
        cryReport.Formulas(4) = ""
        cryReport.Formulas(5) = ""
        cryReport.SelectionFormula = "{tblOutStanding.VYear} = " & gCYear - 1 & " and ({tblOutStanding.Vtype}='SO' or {tblOutStanding.Vtype}='SY' or {tblOutStanding.Vtype}='ST'  or   {tblOutStanding.Vtype}='SD' or  {tblOutStanding.Vtype}='SM')"
        cryReport.ReportFileName = gReportPath & "rptUtlLastYrOutStandingSale.rpt"
        cryReport.Action = 1
        Exit Sub
ErrorHandler:
        GProcErrorHandler
Case 6 'Add
    FormAction = vbDataActionAddNew
    Call GProcActivateControls(Me, True)
    Call GProcSetButtons(Me, vbDataActionAddNew)
    Call GProcClearForm(Me, rstbl, MaxNo, True)
    If FirstTimeFlag = True Then
       FirstTimeFlag = False
    Else
       Cbo(1).SetFocus
    End If
    txt(2) = gCYear - 1
    SetControlEd
    'cbo_LostFocus (1)
Case 7 'Modify
    'If mURecModi = True Then
    '    SearchHelp
    '    DoEvents
    If IsRelatedRecord = False Then
        FormAction = vbDataActionUpdate
        Call GProcActivateControls(Me, True)
        Call GProcSetButtons(Me, vbDataActionUpdate)
        Cbo(1).SetFocus
        txt(0).Enabled = False
        SetControlEd
   End If
Case 8 'Delete
    'If mURecDel = True Then
     If IsRelatedRecord = False Then
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
    End If
Case 9 'Print
Case 10 'Exit
    DoEvents
    Unload Me
Case 11 'Label
    'frmLabel.Show 1
Case 12 'Optional
Case 13 'Save
    If ValidateData = True Then
        If FormAction = vbDataActionAddNew Then txt(0) = GProcGenerateId(tblName, "Vno", "Vtype='" & GetVType & "' and VYear=" & CLng(txt(2)))
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
Private Sub SetControlEd()
txt(2).Enabled = False '--Year
mskDt(10).Enabled = False '--Due Dt
End Sub
Private Sub CalCulate()
'-- Due date
If IsDate(mskDt(3)) = True Then
    mskDt(10) = DateAdd("d", txt(9), mskDt(3))
End If
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
    ConditionStr = " where Vyear=" & gCYear - 1 & " and Vtype in ('SO','SY','ST','SD','SM')"
    If gBackEndDB = gBackEndAccess Then
        X = "Select " & tblName & ".*,Vno & '-' & Vtype as VnoVType from " & tblName & ConditionStr
    Else '-- Oracle
        X = "Select " & tblName & ".*,Vno||'-'||Vtype as VnoVType from " & tblName & ConditionStr
    End If
    If Trim(OrderField) <> "" Then X = X & " order by " & OrderField
    GProcRstOpen rstbl, X, "O"
    '---- Button Visiblity Set As Per User Previlage
    GProcSetButtonVisiblity Me
    FirstTimeFlag = True
    cmdBtn_Click (6)
    Cbo(1).ListIndex = 0
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
If FormAction = vbDataActionUpdate Then
    MsgBox "You are about to modify the record. Save the record and then exit.", vbInformation + vbOKOnly, Me.Caption
    Cancel = True
    Exit Sub
End If
If rstbl.State = adStateOpen Then rstbl.Close
End Sub

Private Sub mskDt_GotFocus(Index As Integer)
'On Error Resume Next
GProcSelectBox Me.ActiveControl
End Sub
Private Sub mskDt_LostFocus(Index As Integer)
If Index = 3 Then
    If GProcIsDateValid(mskDt(Index), gCYear - 1) = False Then
        mskDt(Index).SetFocus
        MsgBox ("Date not within financial year")
        Exit Sub
    End If
End If
CalCulate
End Sub
Private Sub txt_GotFocus(Index As Integer)
GProcSelectBox Me.ActiveControl
End Sub

Private Sub txt_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
Dim GName As String
If KeyCode = vbKeyF3 Then
    Select Case Index
        Case 12 '-- short
            GProcShowForm frmMastItem, frmMain.mnuMstfrm(3), Array(frmMastItem.txttemp(9)), Array(txttemp(11))
        Case 14 '-- ticket
            GProcShowForm frmMastItem, frmMain.mnuMstfrm(3), Array(frmMastItem.txttemp(9)), Array(txttemp(11))
    End Select
End If
End Sub

Private Sub txt_LostFocus(Index As Integer)
GProcCheckForNumber Me, rstbl, MaxNo
SaveBtnEd
CalCulate
End Sub
Private Sub txt_KeyPress(Index As Integer, KeyAscii As Integer)
Dim LeftPos As Long
Dim TopPos As Long
Dim MillCode As Long
TopPos = txt(Index).Top + Me.Top + 650
LeftPos = txt(Index).Left + Me.Left
If KeyAscii = 13 Then Exit Sub
KeyAscii = GProcValidateKey(rstbl, Index, KeyAscii, txt(Index).text)
'--short
If Index = 12 Then
    MillCode = GProcGetColumnValue("TblMastAccount", "AcName", txttemp(11), "S", "Accode", "N")
    gClsSearch.SearchMultiField "tblMastItem", "ItShort", Array("Short"), Array(txt(Index).Width, 0), "ITMillCode=" & MillCode, Chr(KeyAscii), "ItShort", LeftPos, TopPos, True
    KeyAscii = 0
    If gClsSearch.SearchMultiRetCol(0) <> "" Then
        txt(Index) = gClsSearch.SearchMultiRetCol(0)
    End If
End If
'---ticket
If Index = 14 Then
    MillCode = GProcGetColumnValue("TblMastAccount", "AcName", txttemp(11), "S", "Accode", "N")
    gClsSearch.SearchMultiField "tblMastItem", "ItTicket", Array("Ticket"), Array(txt(Index).Width, 0), "ITMillCode=" & MillCode, Chr(KeyAscii), "ItTicket", LeftPos, TopPos, True
    KeyAscii = 0
    If gClsSearch.SearchMultiRetCol(0) <> "" Then
        txt(Index) = gClsSearch.SearchMultiRetCol(0)
    End If
End If
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
Dim BrkAcCode As Long
TopPos = txttemp(Index).Top + Me.Top + 650
LeftPos = txttemp(Index).Left + Me.Left
If KeyAscii = 13 Then Exit Sub
'--Party
  If Index = 5 Then
    gClsSearch.SearchMultiField "QryMastaccount", "Account,AcName,AcCity,AgName", Array("Party", "AcName", "City", "Group"), Array(3000, 0, 2000, 3000), "(AgCode=90017 or GpCode=90017)", Chr(KeyAscii), "AcName", LeftPos, TopPos, True
    KeyAscii = 0
    If gClsSearch.SearchMultiRetCol(0) <> "" And gClsSearch.SearchMultiRetCol(1) = "" Then
        txttemp(Index) = gClsSearch.SearchMultiRetCol(0)
    Else
        If gClsSearch.SearchMultiRetCol(1) <> "" Then txttemp(Index) = gClsSearch.SearchMultiRetCol(1)
    End If
    '-- Broker from master
    If txttemp(6) = "" Then
       BrkAcCode = GProcGetColumnValue("TblMastAccount", "ACName", txttemp(5), "S", "AcBrkCode", "N")
       txttemp(6) = GProcGetColumnValue("TblMastAccount", "ACCode", CStr(BrkAcCode), "N", "ACname", "S")
    End If
 End If
 '-- Broker
If Index = 6 Then
    gClsSearch.SearchMultiField "QryMastaccount", "Account,AcName", Array("Broker", "AcName"), Array(txttemp(Index).Width, 0), "(AgCode=90016 or GpCode=90016)", Chr(KeyAscii), "AcName", LeftPos, TopPos, True
    KeyAscii = 0
    If gClsSearch.SearchMultiRetCol(0) <> "" And gClsSearch.SearchMultiRetCol(1) = "" Then
        txttemp(Index) = gClsSearch.SearchMultiRetCol(0)
    Else
        If gClsSearch.SearchMultiRetCol(1) <> "" Then txttemp(Index) = gClsSearch.SearchMultiRetCol(1)
    End If
End If
 '-- Mill Name
 If Index = 11 Then
    If Cbo(1).ListIndex = 0 Then '--Trade
        gClsSearch.SearchMultiField "QryMastaccount", "Account,AcName,AcCity,AgName", Array("Mill", "AcName", "City", "Group"), Array(3000, 0, 2000, 3000), " (AgCode=90029 and (AcMillType =0 or AcMillType=5) )", Chr(KeyAscii), "AcName", LeftPos, TopPos, True
    ElseIf Cbo(1).ListIndex = 1 Then '--Cons
        gClsSearch.SearchMultiField "QryMastaccount", "Account,AcName,AcCity,AgName", Array("Mill", "AcName", "City", "Group"), Array(3000, 0, 2000, 3000), " (AgCode=90029 and AcMillType =1  )", Chr(KeyAscii), "AcName", LeftPos, TopPos, True
    ElseIf Cbo(1).ListIndex = 2 Then '--Depot
        gClsSearch.SearchMultiField "QryMastaccount", "Account,AcName,AcCity,AgName", Array("Mill", "AcName", "City", "Group"), Array(3000, 0, 2000, 3000), " (AgCode=90029 and AcMillType =2  )", Chr(KeyAscii), "AcName", LeftPos, TopPos, True
    ElseIf Cbo(1).ListIndex = 3 Then '--SIT
        gClsSearch.SearchMultiField "QryMastaccount", "Account,AcName,AcCity,AgName", Array("Mill", "AcName", "City", "Group"), Array(3000, 0, 2000, 3000), " (AgCode=90029 and (AcMillType =4 or AcMillType=5) )", Chr(KeyAscii), "AcName", LeftPos, TopPos, True
    Else '-- Mill Bill
        gClsSearch.SearchMultiField "QryMastaccount", "Account,AcName,AcCity,AgName", Array("Mill", "AcName", "City", "Group"), Array(3000, 0, 2000, 3000), " (AgCode=90029 and AcMillType =3)", Chr(KeyAscii), "AcName", LeftPos, TopPos, True
    End If
    KeyAscii = 0
    If gClsSearch.SearchMultiRetCol(0) <> "" And gClsSearch.SearchMultiRetCol(1) = "" Then
        txttemp(Index) = gClsSearch.SearchMultiRetCol(0)
    Else
        If gClsSearch.SearchMultiRetCol(1) <> "" Then txttemp(Index) = gClsSearch.SearchMultiRetCol(1)
    End If
    MillCode = GProcGetColumnValue("tblMastAccount", "AcName", txttemp(11), "S", "AcCode", "N")
    '-- Ticket
    If txt(14) = "" Then txt(14) = GProcGetColumnValue("tblMastItem", "ItMillCode", CStr(MillCode), "N", "ItTicket", "S")
    '--Short
    If txt(12) = "" Then txt(12) = GProcGetColumnValue("tblMastItem", "ItMillCode", CStr(MillCode), "N", "ItShort", "S")
End If
End Sub
Private Sub txttemp_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
Dim GName As String
If KeyCode = vbKeyF3 Then
    Select Case Index
        Case 5 '-- Party
            GName = GProcGetColumnValue("TblMastGroup", "AgCode", 90017, "N", "AgName", "S")
            GProcShowForm frmMastAccount, frmMain.mnuMstfrm(0), Array(frmMastAccount.txttemp(2)), Array(GName)
        Case 6 '-- Broker
            GName = GProcGetColumnValue("TblMastGroup", "AgCode", 90016, "N", "AgName", "S")
            GProcShowForm frmMastAccount, frmMain.mnuMstfrm(0), Array(frmMastAccount.txttemp(2)), Array(GName)
        Case 11 '-- mill
            GName = GProcGetColumnValue("TblMastGroup", "AgCode", 90029, "N", "AgName", "S")
            GProcShowForm frmMastAccount, frmMain.mnuMstfrm(0), Array(frmMastAccount.txttemp(2)), Array(GName)
    End Select
End If
End Sub
Private Sub SaveBtnEd()
'---Vno,Bill No,vdt,Party , Broker, Mill
If CLng(txt(0)) = 0 Or Trim(txt(4)) = "" Or GProcIsDateValid(mskDt(3), txt(2)) = False Or Trim(txttemp(5)) = "" Or Trim(txttemp(6)) = "" Or Trim(txttemp(11)) = "" Then
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
    CalCulate
    '--- Required
    '-- Vdt
    If GProcIsDateValid(mskDt(3), txt(2)) = False Then
        MsgBox "Check Voucher Date.", vbInformation + vbOKOnly, Me.Caption
        mskDt(3).SetFocus
        Exit Function
    End If
    '-- Vno
    If txt(0) <= 0 Then
        MsgBox "Check Voucher No.", vbCritical + vbOKOnly, Me.Caption
        txt(0).SetFocus
        Exit Function
    End If
    '---Bill No
    If Trim(txt(4)) = "" Then
        MsgBox "Check Bill No.", vbCritical + vbOKOnly, Me.Caption
        txt(4).SetFocus
        Exit Function
    End If
    '-- Party
    If Trim(txttemp(5)) = "" Then
        MsgBox "Check Party Account.", vbCritical + vbOKOnly, Me.Caption
        txttemp(5).SetFocus
        Exit Function
    End If
    '-- Create Party
        If GProcGetColumnValue("tblMastAccount", "AcName", txttemp(5), "S", "AcCode", "N") = 0 Then
           GName = GProcGetColumnValue("TblMastGroup", "AgCode", 90017, "N", "AgName", "S")
           GProcShowForm frmMastAccount, frmMain.mnuMstfrm(0), Array(frmMastAccount.txt(1), frmMastAccount.txttemp(2)), Array(txttemp(5), GName)
           Exit Function
        End If
    '-- Broker
    If Trim(txttemp(6)) = "" Then
        MsgBox "Check Broker Account.", vbCritical + vbOKOnly, Me.Caption
        txttemp(6).SetFocus
        Exit Function
    End If
    '-- Create Broker
    If GProcGetColumnValue("tblMastAccount", "AcName", txttemp(6), "S", "AcCode", "N") = 0 Then
       GName = GProcGetColumnValue("TblMastGroup", "AgCode", 90016, "N", "AgName", "S")
       GProcShowForm frmMastAccount, frmMain.mnuMstfrm(0), Array(frmMastAccount.txt(1), frmMastAccount.txttemp(2)), Array(txttemp(6), GName)
       Exit Function
    End If
    '--- Mill
    If Trim(txttemp(11)) = "" Then
        MsgBox "Check Mill.", vbCritical + vbOKOnly, Me.Caption
        txttemp(11).SetFocus
        Exit Function
    End If
    '-- Create Mill
    If GProcGetColumnValue("tblMastAccount", "AcName", txttemp(11), "S", "AcCode", "N") = 0 Then
        GName = GProcGetColumnValue("TblMastGroup", "AgCode", 90029, "N", "AgName", "S")
        GProcShowForm frmMastAccount, frmMain.mnuMstfrm(0), Array(frmMastAccount.txt(1), frmMastAccount.txttemp(2)), Array(txttemp(11), GName)
        Exit Function
    End If
'    '--- Ticket
'    If Trim(txt(14)) = "" Then
'        MsgBox "Check Item Ticket.", vbCritical + vbOKOnly, Me.Caption
'        txt(14).SetFocus
'        Exit Function
'    End If
'    ''-- Create Ticket
'    'GProcShowForm frmMastItem, frmMain.mnuMstfrm(3), Array(frmMastItem.txttemp(9)), Array(txttemp(11))
'    'GProcShowForm frmMastItem, frmMain.mnuMstfrm(3), Array(frmMastItem.txttemp(9)), Array(txttemp(11))
'    '---Short
'    If Trim(txt(12)) = "" Then
'        MsgBox "Check Item Short.", vbCritical + vbOKOnly, Me.Caption
'        txt(12).SetFocus
'        Exit Function
'    End If
    '---Amt
    If CDbl(txt(7)) + CDbl(txt(9)) < 0 Then
        MsgBox "Check Amount.", vbCritical + vbOKOnly, Me.Caption
       ' txt(7).SetFocus
        Exit Function
    End If
    '----Addition Case
    If FormAction = vbDataActionAddNew Then
      X = "Select VNo From " & tblName & " Where VNo=" & CLng(txt(0)) & " and vtype='" & GetVType & "' and VYear=" & gCYear - 1
        i = GProcRstOpen(rstAddCheck, X, "R")
        If i > 0 Then
            MsgBox "Record already exists  " & Trim(txt(0).text), vbCritical + vbOKOnly, "Duplicate Record"
            txt(0).SetFocus
            Exit Function
        End If
        rstAddCheck.Close
    End If
    ValidateData = True
End Function
Private Sub FillTxtFromTemp()
txt(1) = GetVType '--Vtype
If IsDate(mskDt(3)) Then '-- Vdt
    txt(3) = CDate(mskDt(3))
Else
    txt(3) = ""
End If
If IsDate(mskDt(10)) Then '--Due Date
    txt(10) = CDate(mskDt(10))
Else
    txt(10) = ""
End If
'--Party
txt(5) = GProcGetColumnValue("tblMastAccount", "AcName", txttemp(5), "S", "AcCode", "N")
'-- Broker
txt(6) = GProcGetColumnValue("TBLMASTAccount", "AcName", txttemp(6), "S", "AcCode", "N")
'-- Mill
txt(11) = GProcGetColumnValue("TBLMASTAccount", "AcName", txttemp(11), "S", "AcCode", "N")
End Sub
Private Sub FillTempFromTxt()
Dim X As String
SetControlEd
If IsDate(txt(3)) Then '--VDt
    mskDt(3) = txt(3)
Else
    mskDt(3) = "__/__/____"
End If
If IsDate(txt(10)) Then '--Due date
    mskDt(10) = txt(10)
Else
    mskDt(10) = "__/__/____"
End If
'-- Vtype
Select Case txt(1)
    Case "SY" '-- Trade
        Cbo(1).ListIndex = 0
    Case "SO" '--Consg
        Cbo(1).ListIndex = 1
    Case "SD" '--Depot
        Cbo(1).ListIndex = 2
    Case "ST" '--SIT
        Cbo(1).ListIndex = 3
    Case "SM" '--Mill Bill
        Cbo(1).ListIndex = 4
End Select
'-- Party
txttemp(5) = GProcGetColumnValue("tblMastAccount", "AcCode", txt(5), "N", "AcName", "S")
'--Broker
txttemp(6) = GProcGetColumnValue("TBLMASTAccount", "AcCode", txt(6), "N", "AcName", "S")
'--Mill
txttemp(11) = GProcGetColumnValue("TBLMASTAccount", "AcCode", txt(11), "N", "AcName", "S")
End Sub
Private Sub ReadFields()
GProcShowRecord Me, rstbl, MaxNo
FillTempFromTxt
End Sub
'-- Related Record
Private Function IsRelatedRecord() As Boolean
Dim X As String
Dim Rs1 As Recordset
'-- Audited
If GProcISAudited(txt(0), txt(1), txt(2)) = True Then
    MsgBox "Record is Audited.You can not Modify/Delete this Record.", vbCritical + vbOKOnly, App.Title
    IsRelatedRecord = True
    Exit Function
End If
'--- Receipt
X = "Select RecVno from tblRecVsSale where InvNo=" & txt(0) & " and InvTp='" & txt(1) & "' and InvYear=" & txt(2)
GProcRstOpen Rs1, X, "R"
If Rs1.EOF And Rs1.BOF Then
Else
    MsgBox "Related Receipt Record " & Rs1.Fields("RecVno") & " is exists.You can not Modify/Delete this Record.", vbCritical + vbOKOnly, App.Title
    IsRelatedRecord = True
    Exit Function
End If
Rs1.Close
End Function

