VERSION 5.00
Begin VB.Form frmMastDelAdd 
   BackColor       =   &H00FBF2E1&
   BorderStyle     =   0  'None
   Caption         =   " Narration Master"
   ClientHeight    =   6600
   ClientLeft      =   0
   ClientTop       =   60
   ClientWidth     =   10635
   FillColor       =   &H000040C0&
   FillStyle       =   0  'Solid
   Icon            =   "frmMastDelAdd.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6600
   ScaleWidth      =   10635
   ShowInTaskbar   =   0   'False
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
      Index           =   1
      Left            =   2325
      TabIndex        =   1
      Text            =   "txtCode(1)"
      ToolTipText     =   "Press F3 for Creation"
      Top             =   585
      Width           =   1095
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
      Left            =   2295
      MaxLength       =   40
      TabIndex        =   3
      Text            =   "11"
      Top             =   1575
      Width           =   5985
   End
   Begin VB.TextBox txt 
      Height          =   360
      Index           =   1
      Left            =   6615
      TabIndex        =   43
      Text            =   "1"
      Top             =   7110
      Visible         =   0   'False
      Width           =   630
   End
   Begin VB.TextBox txt 
      Height          =   360
      Index           =   8
      Left            =   8190
      TabIndex        =   42
      Text            =   "8"
      Top             =   7110
      Visible         =   0   'False
      Width           =   630
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
      Index           =   8
      Left            =   5805
      TabIndex        =   10
      Text            =   "cbo(8)"
      Top             =   4050
      Visible         =   0   'False
      Width           =   2595
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
      Index           =   7
      Left            =   2325
      TabIndex        =   9
      Text            =   "cbo(7)"
      Top             =   4005
      Visible         =   0   'False
      Width           =   2595
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
      Index           =   6
      Left            =   5805
      TabIndex        =   8
      Text            =   "cbo(6)"
      Top             =   3555
      Visible         =   0   'False
      Width           =   2595
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
      Index           =   6
      Left            =   7875
      TabIndex        =   37
      Text            =   "6"
      Top             =   7515
      Visible         =   0   'False
      Width           =   525
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
      Index           =   4
      Left            =   2325
      MaxLength       =   60
      TabIndex        =   6
      Text            =   "4"
      Top             =   3060
      Width           =   5985
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
      Height          =   795
      Index           =   12
      Left            =   11340
      Style           =   1  'Graphical
      TabIndex        =   36
      Top             =   6615
      Visible         =   0   'False
      Width           =   1095
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
      Index           =   5
      Left            =   2325
      TabIndex        =   7
      Text            =   "cbo(5)"
      Top             =   3525
      Width           =   2595
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
      Index           =   9
      Left            =   2325
      MaxLength       =   10
      TabIndex        =   11
      Text            =   "9"
      Top             =   4500
      Visible         =   0   'False
      Width           =   2595
   End
   Begin VB.TextBox txt 
      Height          =   360
      Index           =   7
      Left            =   7380
      TabIndex        =   35
      Text            =   "7"
      Top             =   7110
      Visible         =   0   'False
      Width           =   630
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
      Index           =   0
      Left            =   5580
      TabIndex        =   32
      Text            =   "0"
      Top             =   7200
      Visible         =   0   'False
      Width           =   795
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
      Index           =   10
      Left            =   2325
      MaxLength       =   25
      TabIndex        =   12
      Text            =   "10"
      Top             =   4995
      Width           =   6015
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
      Left            =   9360
      TabIndex        =   13
      Text            =   "5"
      Top             =   7155
      Visible         =   0   'False
      Width           =   525
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
      Index           =   3
      Left            =   2325
      MaxLength       =   60
      TabIndex        =   5
      Text            =   "3"
      Top             =   2550
      Width           =   5985
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
      Index           =   2
      Left            =   2325
      MaxLength       =   60
      TabIndex        =   4
      Text            =   "2"
      Top             =   2070
      Width           =   5985
   End
   Begin VB.TextBox txtTemp 
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
      Index           =   1
      Left            =   2325
      TabIndex        =   2
      Text            =   "txtTemp(1)"
      Top             =   1095
      Width           =   5985
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
      Left            =   10080
      Style           =   1  'Graphical
      TabIndex        =   26
      Top             =   7110
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
      Left            =   8925
      Style           =   1  'Graphical
      TabIndex        =   25
      ToolTipText     =   "Exit"
      Top             =   4065
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
      Left            =   6345
      Style           =   1  'Graphical
      TabIndex        =   24
      ToolTipText     =   "To Cancel Record"
      Top             =   5715
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
      Left            =   4680
      Style           =   1  'Graphical
      TabIndex        =   23
      ToolTipText     =   "To Save Record"
      Top             =   5715
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
      Left            =   11880
      Style           =   1  'Graphical
      TabIndex        =   22
      Top             =   5940
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
      Left            =   3330
      Style           =   1  'Graphical
      TabIndex        =   21
      ToolTipText     =   "To Delete Record"
      Top             =   5715
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
      Left            =   1980
      Style           =   1  'Graphical
      TabIndex        =   20
      ToolTipText     =   "To Modify Record"
      Top             =   5715
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
      Left            =   765
      Style           =   1  'Graphical
      TabIndex        =   19
      ToolTipText     =   "To Add New Record"
      Top             =   5715
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
      Left            =   8925
      Style           =   1  'Graphical
      TabIndex        =   18
      Top             =   3195
      Visible         =   0   'False
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
      Left            =   8925
      Style           =   1  'Graphical
      TabIndex        =   17
      ToolTipText     =   "To Find Record"
      Top             =   2340
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
      Left            =   9525
      Style           =   1  'Graphical
      TabIndex        =   16
      ToolTipText     =   "Last"
      Top             =   1620
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
      Left            =   8925
      Style           =   1  'Graphical
      TabIndex        =   15
      ToolTipText     =   "First"
      Top             =   1620
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
      Left            =   9525
      Style           =   1  'Graphical
      TabIndex        =   14
      ToolTipText     =   "Next"
      Top             =   900
      Width           =   495
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
      Left            =   8925
      Style           =   1  'Graphical
      TabIndex        =   0
      ToolTipText     =   "Previous"
      Top             =   900
      Width           =   495
   End
   Begin VB.Label lbl1 
      BackStyle       =   0  'Transparent
      Caption         =   "Party Code      :"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   375
      Index           =   10
      Left            =   765
      TabIndex        =   45
      Top             =   630
      Width           =   1500
   End
   Begin VB.Label lbl1 
      BackStyle       =   0  'Transparent
      Caption         =   "Contact To     :"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   375
      Index           =   9
      Left            =   810
      TabIndex        =   44
      Top             =   1620
      Width           =   1500
   End
   Begin VB.Label lbl1 
      BackStyle       =   0  'Transparent
      Caption         =   "Dist.              :"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   495
      Index           =   8
      Left            =   810
      TabIndex        =   41
      Top             =   3960
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.Label lbl1 
      BackStyle       =   0  'Transparent
      Caption         =   "PIN .             :"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   315
      Index           =   7
      Left            =   810
      TabIndex        =   40
      Top             =   4545
      Visible         =   0   'False
      Width           =   1425
   End
   Begin VB.Label lbl1 
      BackStyle       =   0  'Transparent
      Caption         =   "Tal.     :"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   495
      Index           =   2
      Left            =   5085
      TabIndex        =   39
      Top             =   4095
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.Label lbl1 
      BackStyle       =   0  'Transparent
      Caption         =   "State  :"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   495
      Index           =   1
      Left            =   5085
      TabIndex        =   38
      Top             =   3600
      Visible         =   0   'False
      Width           =   660
   End
   Begin VB.Shape Shape4 
      BackColor       =   &H00C00000&
      BackStyle       =   1  'Opaque
      Height          =   150
      Left            =   495
      Shape           =   3  'Circle
      Top             =   1170
      Width           =   150
   End
   Begin VB.Shape Shape2 
      FillColor       =   &H00FED8FC&
      FillStyle       =   0  'Solid
      Height          =   960
      Left            =   495
      Shape           =   4  'Rounded Rectangle
      Top             =   5445
      Width           =   7365
   End
   Begin VB.Shape Shape3 
      BorderColor     =   &H00EFB17A&
      BorderWidth     =   5
      FillColor       =   &H00EFB17A&
      Height          =   6525
      Left            =   45
      Top             =   45
      Width           =   10545
   End
   Begin VB.Label lbl1 
      BackStyle       =   0  'Transparent
      Caption         =   "    :"
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
      Index           =   4
      Left            =   2115
      TabIndex        =   34
      Top             =   3285
      Visible         =   0   'False
      Width           =   495
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
      Height          =   495
      Left            =   7935
      TabIndex        =   33
      Top             =   60
      Width           =   2475
   End
   Begin VB.Label lbl1 
      BackStyle       =   0  'Transparent
      Caption         =   "Party &Name   :"
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
      Index           =   0
      Left            =   795
      TabIndex        =   31
      Top             =   1095
      Width           =   1650
   End
   Begin VB.Label lbl1 
      BackStyle       =   0  'Transparent
      Caption         =   "&Phone No.     :"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   375
      Index           =   6
      Left            =   795
      TabIndex        =   30
      Top             =   4995
      Width           =   1575
   End
   Begin VB.Label lbl1 
      BackStyle       =   0  'Transparent
      Caption         =   "City               :"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   495
      Index           =   5
      Left            =   795
      TabIndex        =   29
      Top             =   3525
      Width           =   1695
   End
   Begin VB.Label lbl1 
      BackStyle       =   0  'Transparent
      Caption         =   "Add&ress         :"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   375
      Index           =   3
      Left            =   795
      TabIndex        =   28
      Top             =   2070
      Width           =   1815
   End
   Begin VB.Label Label6 
      BackColor       =   &H00EFB17A&
      Caption         =   "   Delivery Address"
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
      TabIndex        =   27
      Top             =   45
      Width           =   10695
   End
   Begin VB.Shape Shape1 
      FillColor       =   &H00FED8FC&
      FillStyle       =   0  'Solid
      Height          =   4245
      Left            =   8595
      Shape           =   4  'Rounded Rectangle
      Top             =   630
      Width           =   1695
   End
End
Attribute VB_Name = "frmMastDelAdd"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Const tblName As String = "tblMastDeleAdd"
Const IdField As String = "DelCode"
Const OrderField As String = "PartyCode"
Const MaxNo As Long = 11
Dim rstbl As Recordset
Dim mNarrTp As String
Dim i As Long
Dim FirstTimeFlag As Boolean
Dim FormAction As Integer
Private Sub Cbo_KeyPress(Index As Integer, KeyAscii As Integer)
If KeyAscii = 13 Then Exit Sub
End Sub

Private Sub cbo_LostFocus(Index As Integer)
Dim CityState As String
Dim CityDist As String
Dim CityTal As String
Dim CityZip As String
If Index = 7 Or Index = 8 Or Index = 5 Or Index = 6 Then
    Cbo(Index).text = GProcProperCase(Trim(Cbo(Index).text))
    If Index = 5 Then
       CityState = GProcGetColumnValue("tblMastDeleAdd", "DCity", Cbo(5).text, "S", "DState", "S")
       Cbo(6).text = CityState
    End If
    If Index = 5 Then
       CityDist = GProcGetColumnValue("tblMastDeleAdd", "DCity", Cbo(5).text, "S", "DDist", "S")
       Cbo(7).text = CityDist
    End If
    If Index = 5 Then
       CityTal = GProcGetColumnValue("tblMastDeleAdd", "DCity", Cbo(5).text, "S", "DTal", "S")
       Cbo(8).text = CityTal
    End If
    
    If Index = 5 Then
       CityZip = GProcGetColumnValue("tblMastDeleAdd", "DCity", Cbo(5).text, "S", "DZip", "S")
       txt(9).text = CityZip
    End If
End If
Cbo(Index).text = GProcProperCase(Trim(Cbo(Index).text))
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
    gClsSearch.SearchMultiField "tblMastAccount,tblMastDeleAdd", "AcName,DAdd1,DAdd2,ContactPer,DCity,DelCode", Array("Party Name", "Address 1", "Address 2", "Contact  ", "City", "Code"), Array(3000, 3000, 3000, 3000, 2000, 0), " tblMastDeleAdd.PartyCode=tblMastAccount.AcCode", "", "AcName", txttemp(1).Left + Me.Left, txttemp(1).Top + Me.Top + 650
    If gClsSearch.SearchMultiRetCol(0) <> "" Then
        rstbl.MoveFirst
        rstbl.Find "DelCode= " & gClsSearch.SearchMultiRetCol(5) & ""
        gClsSearch.SearchMultiRetCol(5) = ""
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
    txtCode(1).SetFocus
    End If
    FillCombo
Case 7 'Modify
    'If mURecModi = True Then
    '    SearchHelp
    '    DoEvents
'        cmdBtn_Click 4 '-- Find
'        If gClsSearch.SearchMultiRetCol(0) = "" Then Exit Sub
        FormAction = vbDataActionUpdate
        Call GProcActivateControls(Me, True)
        Call GProcSetButtons(Me, vbDataActionUpdate)
        txtCode(1).SetFocus
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
Case 12 'Creation To Firm
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

Private Sub Form_Load()
    Dim X As String
    Me.Top = 300
    Me.Left = Screen.Width / 2 - Me.Width / 2
    X = "Select * from " & tblName & " "
    If Trim(OrderField) <> "" Then X = X & " order by " & OrderField
    GProcSetButtonVisiblity Me
    GProcRstOpen rstbl, X, "O"
    FirstTimeFlag = True
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

Private Sub txt_GotFocus(Index As Integer)
GProcSelectBox Me.ActiveControl
End Sub

Private Sub txt_KeyPress(Index As Integer, KeyAscii As Integer)
If KeyAscii = 13 Then Exit Sub
KeyAscii = GProcValidateKey(rstbl, Index, KeyAscii, txt(Index).text)
End Sub
Private Sub txt_LostFocus(Index As Integer)
GProcCheckForNumber Me, rstbl, MaxNo
SaveBtnEd
End Sub
Private Sub FillCombo()
Call GProcFillCombo(Cbo(5), "select Distinct DCity from " & tblName & " where DCity <> '' order by DCity")
Call GProcFillCombo(Cbo(6), "select Distinct DState from " & tblName & " where DState <> '' order by DState")
Call GProcFillCombo(Cbo(7), "select Distinct DDist from " & tblName & " where DDist <> '' order by DDist")
Call GProcFillCombo(Cbo(8), "select Distinct DTal from " & tblName & " where DTal <> '' order by DTal")
End Sub
Private Sub SaveBtnEd()
'--- Narration
If txt(2) = "" Then
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
    GProcCheckForNumber Me, rstbl, MaxNo
    '---- Required
    '-- NarrCode
    If FormAction = vbDataActionAddNew Then txt(0) = GProcGenerateId(tblName, IdField, pStartCode:=gCSNarrCode)
    If gCENarrCode > 0 And CLng(txt(0)) > gCENarrCode Then
        MsgBox "Year End is completed. You can not create new narration.", vbCritical + vbOKOnly, Me.Caption
        txt(2).SetFocus
        Exit Function
    End If
    '--- Narration
    If Trim(txt(2)) = "" Then
        MsgBox "Check Address Line 1", vbCritical + vbOKOnly, Me.Caption
        txt(2).SetFocus
        Exit Function
    End If
    If Trim(txt(2)) = "" Then
        MsgBox "Check Address Line 1", vbCritical + vbOKOnly, Me.Caption
        txt(2).SetFocus
        Exit Function
    End If
    If Trim(txtCode(1)) = "" Or Trim(txtCode(1)) = "D" Then
        MsgBox "Check Party Code", vbCritical + vbOKOnly, Me.Caption
        txtCode(1).SetFocus
        Exit Function
    End If
    '----Duplicate Records
    '----Addition Case
'    If FormAction = vbDataActionAddNew Then
'        X = "Select NArration From " & tblName & " Where Narration='" & Trim(txt(2).Text) & "'"
'        i = GProcRstOpen(rstAddCheck, X, "R")
'        If i > 0 Then
'            MsgBox "Record is Available of Same Narration  " & Trim(txt(2).Text), vbCritical + vbOKOnly, "Duplicate Record"
'            txt(2).SetFocus
'            Exit Function
'        End If
'        rstAddCheck.Close
'    End If
'    '----Modification case
'    If FormAction = vbDataActionUpdate Then
'        X = "Select narration From " & tblName & " Where ((narration='" & Trim(txt(2).Text) & "') and (narrCode<>" & txt(1).Text & "))"
'        i = GProcRstOpen(rstModiCheck, X, "O")
'        If i > 0 Then
'            MsgBox "Record is Available of Same Narration  " & Trim(txt(2).Text), vbCritical + vbOKOnly, "Duplicate Record"
'            txt(2).SetFocus
'            Exit Function
'        End If
'        rstModiCheck.Close
'    End If
    ValidateData = True
End Function
Private Sub FillTxtFromTemp()
'txt(0) = mNarrTp
'-- Tax Purchase,Sale Accode
txt(1) = GProcGetColumnValue("TblMastAccount", "AcName", txttemp(1), "S", "AcCode", "N")
txt(5) = Cbo(5)
txt(6) = Cbo(6)
txt(7) = Cbo(7)
txt(8) = Cbo(8)
End Sub
Private Sub FillTempFromTxt()
txttemp(1) = GProcGetColumnValue("tblMastAccount", "accode", txt(1), "N", "acname", "S", "", gCn)
Cbo(5) = txt(5)
Cbo(6) = txt(6)
Cbo(7) = txt(7)
Cbo(8) = txt(8)
End Sub
Private Sub ReadFields()
GProcShowRecord Me, rstbl, MaxNo
FillTempFromTxt
End Sub
Private Sub txtCode_GotFocus(Index As Integer)
If Index = 1 And txtCode(1) = "" Then
   txtCode(1) = "D"
   txtCode(1).SelStart = 2
End If
End Sub
Private Sub txtCode_LostFocus(Index As Integer)
If Trim(txtCode(Index)) = "" Then Exit Sub
Select Case Index
    Case 1  '-- Party  Code
         txttemp(1) = GProcGetColumnValue("tblMastAccount", "AcOurCode", UCase(txtCode(1)), "S", "AcName", "S")
         If Trim(txttemp(1)) = "" Then
            MsgBox ("Code Not Found ...")
            txttemp(1) = ""
            txttemp(1).Enabled = True
            txttemp(1).SetFocus
         Else
            txt(11).SetFocus
         End If
End Select
End Sub

Private Sub txttemp_GotFocus(Index As Integer)
GProcSelectBox Me.ActiveControl
End Sub
Private Sub txttemp_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
Dim GName As String
If KeyCode = vbKeyF3 Then
    Select Case Index
        Case 1 '-- Party
            GName = GProcGetColumnValue("TblMastGroup", "AgCode", 90017, "N", "AgName", "S")
            GProcShowForm frmMastAccount, frmMain.mnuMstfrm(0), Array(frmMastAccount.txttemp(2)), Array(GName)
    End Select
End If
End Sub
Private Sub txtTemp_KeyPress(Index As Integer, KeyAscii As Integer)
Dim LeftPos As Long
Dim TopPos As Long
Dim MillCode As Long
TopPos = txttemp(Index).Top + Me.Top + 650
LeftPos = txttemp(Index).Left + Me.Left
If KeyAscii = 13 Then Exit Sub
If Index = 1 Then '-- Party Name
'    gClsSearch.SearchMultiField "TblMastaccount", "AcName,AcCity,AcAlName,Accode", Array("AcName", "City", "Alias Name", "Accode"), Array(4000, 1000, 2000, 0), " (not AcName = '-') and agcode = 90017   ", "", " ", txttemp(1).Left + Me.Left, txttemp(1).Top + Me.Top + 650
    gClsSearch.SearchMultiFieldAccMaster Chr(KeyAscii), LeftPos, TopPos
    KeyAscii = 0
    If gClsSearch.SearchMultiRetColAccMaster(0) <> "" Then
        txttemp(Index) = gClsSearch.SearchMultiRetColAccMaster(0)
        txtCode(1) = GProcGetColumnValue("TblMastAccount", "ACName", txttemp(Index), "S", "AcOurCode", "S")
    End If
End If
End Sub
Private Sub txtTemp_LostFocus(Index As Integer)
SaveBtnEd
End Sub

