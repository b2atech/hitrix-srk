VERSION 5.00
Begin VB.Form frmRepoCrea 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Report Creation"
   ClientHeight    =   7350
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9435
   Icon            =   "frmRepoCrea.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7350
   ScaleWidth      =   9435
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txt 
      Appearance      =   0  'Flat
      Height          =   360
      Index           =   23
      Left            =   2025
      MaxLength       =   50
      TabIndex        =   64
      Text            =   "23"
      Top             =   4455
      Width           =   5055
   End
   Begin VB.TextBox txt 
      Appearance      =   0  'Flat
      Height          =   360
      Index           =   22
      Left            =   30
      MaxLength       =   50
      TabIndex        =   63
      Text            =   "22"
      Top             =   6735
      Visible         =   0   'False
      Width           =   315
   End
   Begin VB.TextBox txt 
      Appearance      =   0  'Flat
      Height          =   360
      Index           =   21
      Left            =   540
      MaxLength       =   50
      TabIndex        =   62
      Text            =   "21"
      Top             =   6945
      Visible         =   0   'False
      Width           =   315
   End
   Begin VB.CheckBox chk1 
      Caption         =   "Monthly Totals"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Index           =   21
      Left            =   5400
      TabIndex        =   60
      Top             =   2400
      Width           =   1935
   End
   Begin VB.TextBox txt 
      Appearance      =   0  'Flat
      Height          =   360
      Index           =   20
      Left            =   5760
      MaxLength       =   50
      TabIndex        =   58
      Text            =   "20"
      Top             =   6225
      Width           =   495
   End
   Begin VB.TextBox txt 
      Appearance      =   0  'Flat
      Height          =   360
      Index           =   19
      Left            =   1080
      MaxLength       =   50
      TabIndex        =   52
      Text            =   "19"
      Top             =   6240
      Visible         =   0   'False
      Width           =   315
   End
   Begin VB.TextBox txt 
      Appearance      =   0  'Flat
      Height          =   360
      Index           =   18
      Left            =   5295
      MaxLength       =   50
      TabIndex        =   51
      Text            =   "18"
      Top             =   2760
      Width           =   1470
   End
   Begin VB.TextBox txt 
      Appearance      =   0  'Flat
      Height          =   360
      Index           =   17
      Left            =   2025
      MaxLength       =   50
      TabIndex        =   50
      Text            =   "17"
      Top             =   6660
      Width           =   5055
   End
   Begin VB.TextBox txt 
      Appearance      =   0  'Flat
      Height          =   360
      Index           =   16
      Left            =   2025
      MaxLength       =   50
      TabIndex        =   49
      Text            =   "16"
      Top             =   6225
      Width           =   495
   End
   Begin VB.TextBox txt 
      Appearance      =   0  'Flat
      Height          =   360
      Index           =   15
      Left            =   3615
      MaxLength       =   50
      TabIndex        =   48
      Text            =   "15"
      Top             =   2400
      Width           =   1470
   End
   Begin VB.TextBox txt 
      Appearance      =   0  'Flat
      Height          =   360
      Index           =   14
      Left            =   660
      MaxLength       =   50
      TabIndex        =   47
      Text            =   "14"
      Top             =   6090
      Visible         =   0   'False
      Width           =   315
   End
   Begin VB.TextBox txt 
      Appearance      =   0  'Flat
      Height          =   360
      Index           =   13
      Left            =   75
      MaxLength       =   50
      TabIndex        =   46
      Text            =   "13"
      Top             =   6060
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.TextBox txt 
      Appearance      =   0  'Flat
      Height          =   360
      Index           =   12
      Left            =   2025
      MaxLength       =   50
      TabIndex        =   45
      Text            =   "12"
      Top             =   5775
      Width           =   5055
   End
   Begin VB.TextBox txt 
      Appearance      =   0  'Flat
      Height          =   360
      Index           =   11
      Left            =   2025
      MaxLength       =   50
      TabIndex        =   44
      Text            =   "11"
      Top             =   5385
      Width           =   5055
   End
   Begin VB.TextBox txt 
      Appearance      =   0  'Flat
      Height          =   360
      Index           =   10
      Left            =   2025
      MaxLength       =   50
      TabIndex        =   43
      Text            =   "10"
      Top             =   4920
      Width           =   5055
   End
   Begin VB.TextBox txt 
      Appearance      =   0  'Flat
      Height          =   360
      Index           =   9
      Left            =   2025
      MaxLength       =   50
      TabIndex        =   42
      Text            =   "9"
      Top             =   3135
      Width           =   5055
   End
   Begin VB.TextBox txt 
      Appearance      =   0  'Flat
      Height          =   360
      Index           =   8
      Left            =   2025
      MaxLength       =   50
      TabIndex        =   41
      Text            =   "8"
      Top             =   3630
      Width           =   5055
   End
   Begin VB.TextBox txt 
      Appearance      =   0  'Flat
      Height          =   360
      Index           =   7
      Left            =   2025
      MaxLength       =   50
      TabIndex        =   40
      Text            =   "7"
      Top             =   4020
      Width           =   5055
   End
   Begin VB.TextBox txt 
      Appearance      =   0  'Flat
      Height          =   360
      Index           =   6
      Left            =   5295
      MaxLength       =   50
      TabIndex        =   39
      Text            =   "6"
      Top             =   2025
      Width           =   510
   End
   Begin VB.TextBox txt 
      Appearance      =   0  'Flat
      Height          =   360
      Index           =   5
      Left            =   5295
      MaxLength       =   50
      TabIndex        =   38
      Text            =   "5"
      Top             =   1590
      Width           =   510
   End
   Begin VB.TextBox txt 
      Appearance      =   0  'Flat
      Height          =   360
      Index           =   4
      Left            =   2025
      MaxLength       =   50
      TabIndex        =   37
      Text            =   "4"
      Top             =   2055
      Width           =   390
   End
   Begin VB.TextBox txt 
      Appearance      =   0  'Flat
      Height          =   360
      Index           =   3
      Left            =   2025
      MaxLength       =   50
      TabIndex        =   36
      Text            =   "3"
      Top             =   1605
      Width           =   390
   End
   Begin VB.TextBox txt 
      Appearance      =   0  'Flat
      Height          =   360
      Index           =   2
      Left            =   2025
      MaxLength       =   50
      TabIndex        =   35
      Text            =   "2"
      Top             =   1110
      Width           =   5055
   End
   Begin VB.TextBox txt 
      Appearance      =   0  'Flat
      Height          =   360
      Index           =   1
      Left            =   2025
      MaxLength       =   50
      TabIndex        =   34
      Text            =   "1"
      Top             =   705
      Width           =   5055
   End
   Begin VB.TextBox txt 
      Appearance      =   0  'Flat
      Height          =   360
      Index           =   0
      Left            =   2025
      MaxLength       =   50
      TabIndex        =   33
      Text            =   "0"
      Top             =   240
      Width           =   5055
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
      Left            =   7680
      Style           =   1  'Graphical
      TabIndex        =   32
      ToolTipText     =   "Previous"
      Top             =   465
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
      Left            =   8280
      Style           =   1  'Graphical
      TabIndex        =   31
      ToolTipText     =   "Next"
      Top             =   465
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
      Left            =   7680
      Style           =   1  'Graphical
      TabIndex        =   30
      ToolTipText     =   "First"
      Top             =   945
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
      Left            =   8280
      Style           =   1  'Graphical
      TabIndex        =   29
      ToolTipText     =   "Last"
      Top             =   945
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
      Left            =   7680
      Style           =   1  'Graphical
      TabIndex        =   28
      ToolTipText     =   "To Find Record"
      Top             =   1425
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
      Left            =   7680
      Style           =   1  'Graphical
      TabIndex        =   27
      Top             =   1905
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
      Left            =   7680
      Style           =   1  'Graphical
      TabIndex        =   26
      ToolTipText     =   "To Add New Record"
      Top             =   2385
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
      Left            =   7680
      Style           =   1  'Graphical
      TabIndex        =   25
      ToolTipText     =   "To Modify Record"
      Top             =   2865
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
      Left            =   7680
      Style           =   1  'Graphical
      TabIndex        =   24
      ToolTipText     =   "To Delete Record"
      Top             =   3345
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
      Left            =   7680
      Style           =   1  'Graphical
      TabIndex        =   23
      Top             =   3825
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
      Left            =   7680
      Style           =   1  'Graphical
      TabIndex        =   22
      ToolTipText     =   "To Save Record"
      Top             =   4305
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
      Left            =   7680
      Style           =   1  'Graphical
      TabIndex        =   21
      ToolTipText     =   "To Cancel Record"
      Top             =   4785
      Width           =   1095
   End
   Begin VB.CommandButton cmdBtn 
      Caption         =   "&Exit"
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
      Left            =   7680
      Style           =   1  'Graphical
      TabIndex        =   20
      ToolTipText     =   "Exit"
      Top             =   5265
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
      Left            =   7680
      Style           =   1  'Graphical
      TabIndex        =   19
      Top             =   5745
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
      Left            =   7680
      Style           =   1  'Graphical
      TabIndex        =   18
      Top             =   6225
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.CheckBox chk1 
      Caption         =   "Single A/c Reuired"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Index           =   13
      Left            =   240
      TabIndex        =   16
      Top             =   2760
      Width           =   2415
   End
   Begin VB.CheckBox chk1 
      Caption         =   "Date Reuired"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Index           =   14
      Left            =   240
      TabIndex        =   0
      Top             =   2400
      Width           =   1695
   End
   Begin VB.Label lbl1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "(Oracle)"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   22
      Left            =   1125
      TabIndex        =   65
      Top             =   4470
      Width           =   825
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
      ForeColor       =   &H000000FF&
      Height          =   495
      Left            =   6915
      TabIndex        =   61
      Top             =   -90
      Width           =   2475
   End
   Begin VB.Label lbl1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Procedure No. to Run :"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   21
      Left            =   3360
      TabIndex        =   59
      Top             =   6225
      Width           =   2310
   End
   Begin VB.Label lbl1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "[M/S]"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Index           =   20
      Left            =   2640
      TabIndex        =   57
      Top             =   6225
      Width           =   465
   End
   Begin VB.Label lbl1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "[10/12/17]"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Index           =   19
      Left            =   5880
      TabIndex        =   56
      Top             =   1680
      Width           =   780
   End
   Begin VB.Label lbl1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "[24/36/48/72]"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Index           =   18
      Left            =   2520
      TabIndex        =   55
      Top             =   2040
      Width           =   1020
   End
   Begin VB.Label lbl1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "[P/G]"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Index           =   17
      Left            =   2520
      TabIndex        =   54
      Top             =   1560
      Width           =   405
   End
   Begin VB.Label lbl1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Muli/Single Line:"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   16
      Left            =   240
      TabIndex        =   53
      Top             =   6225
      Width           =   1740
   End
   Begin VB.Shape Shape1 
      FillColor       =   &H00BDDFD1&
      FillStyle       =   0  'Solid
      Height          =   6825
      Left            =   7440
      Shape           =   4  'Rounded Rectangle
      Top             =   120
      Width           =   1575
   End
   Begin VB.Label lbl1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "A/c Name Field :"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   8
      Left            =   3480
      TabIndex        =   17
      Top             =   2805
      Width           =   1680
   End
   Begin VB.Label lbl1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Mulitple Fields:"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   15
      Left            =   240
      TabIndex        =   15
      Top             =   6705
      Width           =   1590
   End
   Begin VB.Label lbl1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Where Condition:"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   10
      Left            =   240
      TabIndex        =   14
      Top             =   3615
      Width           =   1770
   End
   Begin VB.Label lbl1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Table Required :"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   9
      Left            =   240
      TabIndex        =   13
      Top             =   3135
      Width           =   1695
   End
   Begin VB.Label lbl1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Date Field :"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   7
      Left            =   2160
      TabIndex        =   12
      Top             =   2445
      Width           =   1185
   End
   Begin VB.Label lbl1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Field Lengths :"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   14
      Left            =   240
      TabIndex        =   11
      Top             =   5745
      Width           =   1500
   End
   Begin VB.Label lbl1 
      BackStyle       =   0  'Transparent
      Caption         =   "Totals Fields :"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   13
      Left            =   240
      TabIndex        =   10
      Top             =   5385
      Width           =   1530
   End
   Begin VB.Label lbl1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Order By :"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   12
      Left            =   240
      TabIndex        =   9
      Top             =   4905
      Width           =   1080
   End
   Begin VB.Label lbl1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Fields List    :"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   11
      Left            =   240
      TabIndex        =   8
      Top             =   4005
      Width           =   1365
   End
   Begin VB.Label lbl1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Left Margin :"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   6
      Left            =   3840
      TabIndex        =   7
      Top             =   2010
      Width           =   1365
   End
   Begin VB.Label lbl1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Report CPI :"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   5
      Left            =   3840
      TabIndex        =   6
      Top             =   1575
      Width           =   1290
   End
   Begin VB.Label lbl1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Page Length :"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   4
      Left            =   240
      TabIndex        =   5
      Top             =   2010
      Width           =   1395
   End
   Begin VB.Label lbl1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Report Type :"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   3
      Left            =   240
      TabIndex        =   4
      Top             =   1575
      Width           =   1410
   End
   Begin VB.Label lbl1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Report Title2 :"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   2
      Left            =   240
      TabIndex        =   3
      Top             =   1144
      Width           =   1485
   End
   Begin VB.Label lbl1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Report Title1 :"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   1
      Left            =   240
      TabIndex        =   2
      Top             =   711
      Width           =   1485
   End
   Begin VB.Label lbl1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Report Name :"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Index           =   0
      Left            =   240
      TabIndex        =   1
      Top             =   240
      Width           =   1500
   End
End
Attribute VB_Name = "frmRepoCrea"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Const tblName As String = "UdfReports"
Const MaxNo As Long = 23
'Const IdField As String = "AcCode"
Dim rstbl As Recordset
Dim varBookMark As Variant
Dim FindFlag As Boolean
Dim FirstTimeFlag As Boolean
Dim FormAction As Integer

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
Case 7 'Modify
    'If mURecModi = True Then
    '    SearchHelp
    '    DoEvents
        FormAction = vbDataActionUpdate
        Call GProcActivateControls(Me, True)
        Call GProcSetButtons(Me, vbDataActionUpdate)
        txt(0).SetFocus
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
    DoEvents
    Unload Me
Case 11 'Label
    'frmLabel.Show 1
Case 12 'Optional
Case 13 'Save
    If ValidateData = True Then
       FillTxtFromTemp
       GProcSaveRecord Me, rstbl, FormAction, MaxNo
       Call GProcActivateControls(Me, False)
       Call GProcSetButtons(Me, True)
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
    FormAction = vbDataActionClose
End Select
Exit Sub
ErrorRoutine:
    GProcErrorHandler
End Sub
Private Sub Form_Load()
    Dim X As String
    Dim Y As String
    Me.Top = 300
    Me.Left = Screen.Width / 2 - Me.Width / 2
    Me.Caption = "Report Creation "
    Y = "Select * from " & tblName & " order by RepoName"
    GProcRstOpen rstbl, Y, "O", gSetupCn
    'MaxNo = rstbl.Fields.Count - 1
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
rstbl.Requery
rstbl.Close
End Sub
Private Sub txt_GotFocus(Index As Integer)
GProcSelectBox Me.ActiveControl
End Sub
Private Sub txt_LostFocus(Index As Integer)
GProcCheckForNumber Me, rstbl, MaxNo
SaveBtnEd
End Sub
Private Sub txt_KeyPress(Index As Integer, KeyAscii As Integer)
KeyAscii = GProcValidateKey(rstbl, Index, KeyAscii, txt(Index).Text)
End Sub
Private Sub txttemp_GotFocus(Index As Integer)
End Sub
Private Sub txtTemp_LostFocus(Index As Integer)
SaveBtnEd
End Sub
Private Sub SaveBtnEd()
'--- AcName,AgCode,Agcode1
If txt(0) = "" Then
    cmdBtn(10).Enabled = False
    cmdBtn(10).BackColor = &H80000004
Else
    cmdBtn(10).Enabled = True
    cmdBtn(10).BackColor = &HFFFF80
End If
End Sub
Private Function ValidateData() As Boolean
    ValidateData = True
End Function
Private Sub FillTxtFromTemp()
txt(13) = chk1(13)
txt(14) = chk1(14)
txt(21) = chk1(21)
If txt(23) = "" Then txt(23) = txt(7)
End Sub
Private Sub FillTempFromTxt()
chk1(13) = Val(txt(13))
chk1(14) = Val(txt(14))
chk1(21) = Val(txt(21))
End Sub
Private Sub ReadFields()
GProcShowRecord Me, rstbl, MaxNo
FillTempFromTxt
End Sub

