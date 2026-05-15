VERSION 5.00
Begin VB.Form frmMastNarrationTax 
   BackColor       =   &H00FBF2E1&
   BorderStyle     =   0  'None
   Caption         =   " Tax Master"
   ClientHeight    =   8400
   ClientLeft      =   0
   ClientTop       =   60
   ClientWidth     =   10890
   FillColor       =   &H000040C0&
   FillStyle       =   0  'Solid
   Icon            =   "frmMastNarrationTax.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   8400
   ScaleWidth      =   10890
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox txt 
      Height          =   360
      Index           =   23
      Left            =   12870
      TabIndex        =   66
      Text            =   "23"
      Top             =   7020
      Visible         =   0   'False
      Width           =   690
   End
   Begin VB.TextBox txtTemp 
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
      Left            =   2430
      MaxLength       =   40
      TabIndex        =   14
      Text            =   "txttemp(23)"
      ToolTipText     =   "Press F3 For Creation"
      Top             =   5310
      Width           =   3435
   End
   Begin VB.TextBox txt 
      Height          =   360
      Index           =   22
      Left            =   12915
      TabIndex        =   64
      Text            =   "22"
      Top             =   6660
      Visible         =   0   'False
      Width           =   690
   End
   Begin VB.TextBox txt 
      Height          =   360
      Index           =   21
      Left            =   12915
      TabIndex        =   63
      Text            =   "21"
      Top             =   6255
      Visible         =   0   'False
      Width           =   690
   End
   Begin VB.TextBox txt 
      Height          =   360
      Index           =   20
      Left            =   12915
      TabIndex        =   62
      Text            =   "20"
      Top             =   5850
      Visible         =   0   'False
      Width           =   690
   End
   Begin VB.TextBox txt 
      Height          =   360
      Index           =   19
      Left            =   12915
      TabIndex        =   61
      Text            =   "19"
      Top             =   5445
      Visible         =   0   'False
      Width           =   690
   End
   Begin VB.TextBox txt 
      Height          =   360
      Index           =   18
      Left            =   12870
      TabIndex        =   60
      Text            =   "18"
      Top             =   4995
      Visible         =   0   'False
      Width           =   690
   End
   Begin VB.TextBox txt 
      Height          =   360
      Index           =   17
      Left            =   12915
      TabIndex        =   59
      Text            =   "17"
      Top             =   4590
      Visible         =   0   'False
      Width           =   690
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
      Left            =   5490
      MaxLength       =   40
      TabIndex        =   7
      Text            =   "16"
      Top             =   2790
      Width           =   1440
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
      Left            =   5490
      MaxLength       =   40
      TabIndex        =   6
      Text            =   "15"
      Top             =   2295
      Width           =   1440
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
      Left            =   5490
      MaxLength       =   40
      TabIndex        =   5
      Text            =   "14"
      Top             =   1800
      Width           =   1440
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
      Left            =   2250
      MaxLength       =   40
      TabIndex        =   4
      Text            =   "13"
      Top             =   2790
      Width           =   1440
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
      Left            =   2250
      MaxLength       =   40
      TabIndex        =   3
      Text            =   "12"
      Top             =   2295
      Width           =   1440
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
      Left            =   2250
      MaxLength       =   40
      TabIndex        =   2
      Text            =   "11"
      Top             =   1800
      Width           =   1440
   End
   Begin VB.TextBox txtTemp 
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
      Left            =   1395
      MaxLength       =   40
      TabIndex        =   10
      Text            =   "txttemp(22)"
      ToolTipText     =   "Press F3 For Creation"
      Top             =   4770
      Width           =   3435
   End
   Begin VB.TextBox txtTemp 
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
      Left            =   1395
      MaxLength       =   40
      TabIndex        =   9
      Text            =   "txttemp(21)"
      ToolTipText     =   "Press F3 For Creation"
      Top             =   4365
      Width           =   3435
   End
   Begin VB.TextBox txtTemp 
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
      Left            =   1395
      MaxLength       =   40
      TabIndex        =   8
      Text            =   "txttemp(20)"
      ToolTipText     =   "Press F3 For Creation"
      Top             =   3960
      Width           =   3435
   End
   Begin VB.TextBox txtTemp 
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
      Left            =   5220
      MaxLength       =   40
      TabIndex        =   13
      Text            =   "txttemp(19)"
      ToolTipText     =   "Press F3 For Creation"
      Top             =   4725
      Width           =   3435
   End
   Begin VB.TextBox txtTemp 
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
      Left            =   5220
      MaxLength       =   40
      TabIndex        =   12
      Text            =   "txttemp(18)"
      ToolTipText     =   "Press F3 For Creation"
      Top             =   4365
      Width           =   3435
   End
   Begin VB.TextBox txtTemp 
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
      Left            =   5220
      MaxLength       =   40
      TabIndex        =   11
      Text            =   "txttemp(17)"
      ToolTipText     =   "Press F3 For Creation"
      Top             =   3960
      Width           =   3435
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
      Left            =   2475
      MaxLength       =   40
      TabIndex        =   16
      Text            =   "10"
      Top             =   6615
      Width           =   1440
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
      Left            =   12600
      Style           =   1  'Graphical
      TabIndex        =   46
      Top             =   3690
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.TextBox txtTemp 
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
      Left            =   2520
      MaxLength       =   40
      TabIndex        =   18
      Text            =   "txttemp(9)"
      ToolTipText     =   "Press F3 For Creation"
      Top             =   7560
      Width           =   5775
   End
   Begin VB.TextBox txtTemp 
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
      Left            =   2490
      MaxLength       =   40
      TabIndex        =   17
      Text            =   "txttemp(8)"
      ToolTipText     =   "Press F3 For Creation"
      Top             =   7095
      Width           =   5775
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
      Left            =   3495
      MaxLength       =   40
      TabIndex        =   45
      Text            =   "9"
      Top             =   8880
      Visible         =   0   'False
      Width           =   915
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
      Index           =   8
      Left            =   2505
      MaxLength       =   40
      TabIndex        =   44
      Text            =   "8"
      Top             =   8910
      Visible         =   0   'False
      Width           =   915
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
      Index           =   7
      Left            =   2490
      MaxLength       =   40
      TabIndex        =   15
      Text            =   "7"
      Top             =   6075
      Width           =   1440
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
      Index           =   1
      Left            =   1845
      MaxLength       =   40
      TabIndex        =   42
      Text            =   "1"
      Top             =   8835
      Visible         =   0   'False
      Width           =   915
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
      Left            =   4440
      TabIndex        =   41
      Text            =   "0"
      Top             =   8865
      Visible         =   0   'False
      Width           =   795
   End
   Begin VB.TextBox txt 
      Height          =   360
      Index           =   6
      Left            =   13020
      TabIndex        =   37
      Text            =   "6"
      Top             =   2985
      Visible         =   0   'False
      Width           =   690
   End
   Begin VB.TextBox txt 
      Height          =   360
      Index           =   5
      Left            =   13020
      TabIndex        =   36
      Text            =   "5"
      Top             =   2505
      Visible         =   0   'False
      Width           =   690
   End
   Begin VB.TextBox txt 
      Height          =   360
      Index           =   4
      Left            =   13020
      TabIndex        =   35
      Text            =   "4"
      Top             =   2025
      Visible         =   0   'False
      Width           =   690
   End
   Begin VB.TextBox txt 
      Height          =   360
      Index           =   3
      Left            =   13020
      TabIndex        =   34
      Text            =   "3"
      Top             =   1545
      Visible         =   0   'False
      Width           =   690
   End
   Begin VB.TextBox txt 
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
      Index           =   2
      Left            =   2490
      TabIndex        =   1
      Text            =   "2"
      Top             =   690
      Width           =   5775
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
      Left            =   6555
      Style           =   1  'Graphical
      TabIndex        =   39
      Top             =   8850
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
      Left            =   9225
      Style           =   1  'Graphical
      TabIndex        =   33
      ToolTipText     =   "Exit"
      Top             =   7290
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
      Left            =   9225
      Style           =   1  'Graphical
      TabIndex        =   26
      ToolTipText     =   "To Cancel Record"
      Top             =   6660
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
      Left            =   9225
      Style           =   1  'Graphical
      TabIndex        =   25
      ToolTipText     =   "To Save Record"
      Top             =   6120
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
      Left            =   5475
      Style           =   1  'Graphical
      TabIndex        =   38
      Top             =   8850
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
      Left            =   9225
      Style           =   1  'Graphical
      TabIndex        =   24
      ToolTipText     =   "To Delete Record"
      Top             =   5310
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
      Left            =   9225
      Style           =   1  'Graphical
      TabIndex        =   23
      ToolTipText     =   "To Modify Record"
      Top             =   4725
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
      Left            =   9225
      Style           =   1  'Graphical
      TabIndex        =   22
      ToolTipText     =   "To Add New Record"
      Top             =   4185
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
      Left            =   9225
      Style           =   1  'Graphical
      TabIndex        =   28
      Top             =   2745
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
      Left            =   9225
      Style           =   1  'Graphical
      TabIndex        =   27
      ToolTipText     =   "To Find Record"
      Top             =   3285
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
      Left            =   9825
      Style           =   1  'Graphical
      TabIndex        =   32
      ToolTipText     =   "Last"
      Top             =   2250
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
      Left            =   9225
      Style           =   1  'Graphical
      TabIndex        =   31
      ToolTipText     =   "First"
      Top             =   2250
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
      Left            =   9825
      Style           =   1  'Graphical
      TabIndex        =   30
      ToolTipText     =   "Next"
      Top             =   1620
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
      Left            =   9225
      Style           =   1  'Graphical
      TabIndex        =   29
      ToolTipText     =   "Previous"
      Top             =   1620
      Width           =   495
   End
   Begin VB.Label lbl1 
      BackStyle       =   0  'Transparent
      Caption         =   "R C M Sales A/c   :"
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
      Height          =   315
      Index           =   16
      Left            =   495
      TabIndex        =   65
      Top             =   5355
      Width           =   1875
   End
   Begin VB.Label lbl1 
      BackStyle       =   0  'Transparent
      Caption         =   "VAT Details"
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
      Index           =   15
      Left            =   5310
      TabIndex        =   58
      Top             =   5985
      Width           =   2055
   End
   Begin VB.Label lbl1 
      BackStyle       =   0  'Transparent
      Caption         =   "Input GST A/c  (Purchase)"
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
      Height          =   315
      Index           =   14
      Left            =   1395
      TabIndex        =   57
      Top             =   3600
      Width           =   3585
   End
   Begin VB.Label lbl1 
      BackStyle       =   0  'Transparent
      Caption         =   "Output GST A/c   (Sale)    "
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
      Height          =   285
      Index           =   13
      Left            =   5220
      TabIndex        =   56
      Top             =   3600
      Width           =   2760
   End
   Begin VB.Label lbl1 
      BackStyle       =   0  'Transparent
      Caption         =   "CGST  :"
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
      Height          =   375
      Index           =   12
      Left            =   540
      TabIndex        =   55
      Top             =   3960
      Width           =   810
   End
   Begin VB.Label lbl1 
      BackStyle       =   0  'Transparent
      Caption         =   "SGST  :"
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
      Height          =   375
      Index           =   11
      Left            =   540
      TabIndex        =   54
      Top             =   4410
      Width           =   810
   End
   Begin VB.Label lbl1 
      BackStyle       =   0  'Transparent
      Caption         =   "IGST   :"
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
      Height          =   375
      Index           =   10
      Left            =   540
      TabIndex        =   53
      Top             =   4815
      Width           =   810
   End
   Begin VB.Label lbl1 
      BackStyle       =   0  'Transparent
      Caption         =   "Cotton GST Rates"
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
      Index           =   9
      Left            =   2250
      TabIndex        =   52
      Top             =   1410
      Width           =   2055
   End
   Begin VB.Label lbl1 
      BackStyle       =   0  'Transparent
      Caption         =   "Polyster GST Rates"
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
      Index           =   8
      Left            =   5505
      TabIndex        =   51
      Top             =   1395
      Width           =   2055
   End
   Begin VB.Label lbl1 
      BackStyle       =   0  'Transparent
      Caption         =   "CGST  Rate  :"
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
      Height          =   375
      Index           =   7
      Left            =   540
      TabIndex        =   50
      Top             =   1845
      Width           =   1530
   End
   Begin VB.Label lbl1 
      BackStyle       =   0  'Transparent
      Caption         =   "SGST  Rate  :"
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
      Height          =   375
      Index           =   4
      Left            =   540
      TabIndex        =   49
      Top             =   2385
      Width           =   1530
   End
   Begin VB.Label lbl1 
      BackStyle       =   0  'Transparent
      Caption         =   "IGST  Rate   :"
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
      Height          =   375
      Index           =   2
      Left            =   540
      TabIndex        =   48
      Top             =   2880
      Width           =   1530
   End
   Begin VB.Label lbl1 
      BackStyle       =   0  'Transparent
      Caption         =   "Tax %  (Polister)   :"
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
      Index           =   1
      Left            =   585
      TabIndex        =   47
      Top             =   6615
      Width           =   1935
   End
   Begin VB.Shape Shape2 
      BackColor       =   &H00C00000&
      BackStyle       =   1  'Opaque
      Height          =   150
      Left            =   270
      Shape           =   3  'Circle
      Top             =   720
      Width           =   150
   End
   Begin VB.Shape Shape4 
      BorderColor     =   &H00EFB17A&
      BorderWidth     =   5
      FillColor       =   &H00EFB17A&
      Height          =   8370
      Left            =   0
      Top             =   0
      Width           =   10875
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
      TabIndex        =   43
      Top             =   60
      Width           =   2475
   End
   Begin VB.Label lbl1 
      BackStyle       =   0  'Transparent
      Caption         =   "Tax D&escription   :"
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
      Height          =   375
      Index           =   0
      Left            =   570
      TabIndex        =   0
      Top             =   690
      Width           =   2055
   End
   Begin VB.Label lbl1 
      BackStyle       =   0  'Transparent
      Caption         =   "Sale Tax A/c        :"
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
      Left            =   570
      TabIndex        =   21
      Top             =   7545
      Width           =   1815
   End
   Begin VB.Label lbl1 
      BackStyle       =   0  'Transparent
      Caption         =   "&Purchase Tax A/c :"
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
      Left            =   570
      TabIndex        =   20
      Top             =   7095
      Width           =   1695
   End
   Begin VB.Label lbl1 
      BackStyle       =   0  'Transparent
      Caption         =   "Tax %  (Cotton)     :"
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
      Left            =   570
      TabIndex        =   19
      Top             =   6075
      Width           =   1935
   End
   Begin VB.Label Label6 
      BackColor       =   &H00EFB17A&
      Caption         =   "   Tax Master"
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
      TabIndex        =   40
      Top             =   0
      Width           =   10875
   End
   Begin VB.Shape Shape1 
      FillColor       =   &H00FED8FC&
      FillStyle       =   0  'Solid
      Height          =   6750
      Left            =   9060
      Shape           =   4  'Rounded Rectangle
      Top             =   1335
      Width           =   1425
   End
   Begin VB.Shape Shape3 
      BackColor       =   &H00808080&
      BorderColor     =   &H004D177D&
      FillColor       =   &H00ECFFFF&
      FillStyle       =   0  'Solid
      Height          =   2190
      Left            =   270
      Top             =   1215
      Width           =   8580
   End
   Begin VB.Shape Shape6 
      BackColor       =   &H00808080&
      BorderColor     =   &H004D177D&
      FillColor       =   &H00ECFFFF&
      FillStyle       =   0  'Solid
      Height          =   2280
      Left            =   270
      Top             =   5850
      Width           =   8535
   End
   Begin VB.Shape Shape5 
      BackColor       =   &H00808080&
      BorderColor     =   &H004D177D&
      FillColor       =   &H00ECFFFF&
      FillStyle       =   0  'Solid
      Height          =   2280
      Left            =   270
      Top             =   3510
      Width           =   8535
   End
End
Attribute VB_Name = "frmMastNarrationTax"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Const tblName As String = "tblMastNarration"
Const IdField As String = "narrcode"
Const OrderField As String = "Narration"
Const MaxNo As Long = 23
Dim rstbl As Recordset
Dim i As Integer
Dim FirstTimeFlag As Boolean
Dim FormAction As Integer

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
        X = " TblMastAccount as TblMastac1 "
    Else '-- Oracle
        X = " TblMastAccount TblMastac1 "
    End If
    FindQStr = "Select Narration,NarrCode,MastTaxRate,tblMastAccount.AcName,tblMastAc1.AcName from " _
    & " tblMastNarration,TblMastaccount," & X & " where tblMastNarration.MastTaxPurAcCode=TblMastaccount.AcCode" _
    & " and tblMastNarration.MastTaxSAAcCode=TblMastac1.AcCode and NarrType='T'"
    gClsSearch.SearchMultiField FindQStr, "", Array("Tax", "NarrationCd", "Rate", "Purchase Account", "Sales Account"), Array(2000, 0, 1000, 2000, 2000), "", "", "Narration", txt(2).Left + Me.Left, txt(2).Top + Me.Top + 650, False, , True
    If gClsSearch.SearchMultiRetCol(0) <> "" Then
        rstbl.MoveFirst
        rstbl.Find "Narration= '" & gClsSearch.SearchMultiRetCol(0) & "'"
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
    txt(2).SetFocus
    End If
    FillCombo

Case 7 'Modify
    'If mURecModi = True Then
    '    SearchHelp
    '    DoEvents
        cmdBtn_Click 4 '-- Find
        If gClsSearch.SearchMultiRetCol(0) = "" Then Exit Sub
        FormAction = vbDataActionUpdate
        Call GProcActivateControls(Me, True)
        Call GProcSetButtons(Me, vbDataActionUpdate)
        txt(2).SetFocus
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
     gfrmTypeStr = "Tax"
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
    X = "Select * from " & tblName & " where NarrType='T'"
    If Trim(OrderField) <> "" Then X = X & " order by " & OrderField
    '---- Button Visiblity Set As Per User Previlage
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
rstbl.Requery
rstbl.Close
End Sub

Private Sub txt_GotFocus(Index As Integer)
GProcSelectBox Me.ActiveControl
End Sub
Private Sub txt_LostFocus(Index As Integer)
GProcCheckForNumber Me, rstbl, MaxNo
SetControlEd
SaveBtnEd
End Sub
Private Sub txt_KeyPress(Index As Integer, KeyAscii As Integer)
If KeyAscii = 13 Then Exit Sub
KeyAscii = GProcValidateKey(rstbl, Index, KeyAscii, txt(Index).text)
End Sub

Private Sub FillCombo()
'Call GProFillCombo(cbo(7), "select Distinct AcCity from " & tblName & " where AcCity <> '' order by AcCity")
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
    If FormAction = vbDataActionAddNew Then txt(1) = GProcGenerateId(tblName, IdField, pStartCode:=gCSNarrCode)
    If gCENarrCode > 0 And CLng(txt(1)) > gCENarrCode Then
        MsgBox "Year End is completed. You can not create new narration.", vbCritical + vbOKOnly, Me.Caption
        txt(2).SetFocus
        Exit Function
    End If
    '--- Tax Account Purchase
    If txttemp(8) = "" Then txttemp(8) = "-"
    If GProcGetColumnValue("tblMastAccount", "AcName", txttemp(8), "S", "AcCode", "N") = 0 Then
        'GName = GProcGetColumnValue("TblMastGroup", "AgCode", 25, "N", "AgName", "S")
        GProcShowForm frmMastAccount, frmMain.mnuMstfrm(0), Array(frmMastAccount.txt(1)), Array(txttemp(8))
        Exit Function
    End If
    '--- Tax Account Sale
    If txttemp(9) = "" Then txttemp(9) = "-"
    If GProcGetColumnValue("tblMastAccount", "AcName", txttemp(9), "S", "AcCode", "N") = 0 Then
        'GName = GProcGetColumnValue("TblMastGroup", "AgCode", 25, "N", "AgName", "S")
        GProcShowForm frmMastAccount, frmMain.mnuMstfrm(0), Array(frmMastAccount.txt(1)), Array(txttemp(9))
        Exit Function
    End If
    '--- Narration
    If Trim(txt(2)) = "" Then
        MsgBox "Check Narration.", vbCritical + vbOKOnly, Me.Caption
        txt(2).SetFocus
        Exit Function
    End If
    '----Duplicate Records
    '----Addition Case
    If FormAction = vbDataActionAddNew Then
         X = "Select NArration From " & tblName & " Where Narration='" & Trim(txt(2).text) & "'  "
        i = GProcRstOpen(rstAddCheck, X, "R")
        If i > 0 Then
            MsgBox "Record is Available of Same Narration  " & Trim(txt(2).text), vbCritical + vbOKOnly, "Duplicate Record"
            txt(2).SetFocus
            Exit Function
        End If
        rstAddCheck.Close
    End If
    '----Modification case
    If FormAction = vbDataActionUpdate Then
        X = "Select narration From " & tblName & " Where ((narration='" & Trim(txt(2).text) & "') and (narrCode<>" & txt(1).text & "))"
        i = GProcRstOpen(rstModiCheck, X, "O")
        If i > 0 Then
            MsgBox "Record is Available of Same Narration  " & Trim(txt(2).text), vbCritical + vbOKOnly, "Duplicate Record"
            txt(2).SetFocus
            Exit Function
        End If
        rstModiCheck.Close
    End If
    ValidateData = True
End Function
Private Sub FillTxtFromTemp()
txt(0) = "T"
'-- Tax AcCode
For i = 8 To 9
    If txttemp(i) = "" Then
       txttemp(i) = "-"
    End If
Next i
For i = 17 To 22
    If txttemp(i) = "" Then
       txttemp(i) = "-"
    End If
Next i
txt(8) = GProcGetColumnValue("TblMastAccount", "ACName", txttemp(8), "S", "AcCode", "N")
txt(9) = GProcGetColumnValue("TblMastAccount", "ACName", txttemp(9), "S", "AcCode", "N")
txt(17) = GProcGetColumnValue("TblMastAccount", "ACName", txttemp(17), "S", "AcCode", "N")
txt(18) = GProcGetColumnValue("TblMastAccount", "ACName", txttemp(18), "S", "AcCode", "N")
txt(19) = GProcGetColumnValue("TblMastAccount", "ACName", txttemp(19), "S", "AcCode", "N")
txt(20) = GProcGetColumnValue("TblMastAccount", "ACName", txttemp(20), "S", "AcCode", "N")
txt(21) = GProcGetColumnValue("TblMastAccount", "ACName", txttemp(21), "S", "AcCode", "N")
txt(22) = GProcGetColumnValue("TblMastAccount", "ACName", txttemp(22), "S", "AcCode", "N")
txt(23) = GProcGetColumnValue("TblMastAccount", "ACName", txttemp(23), "S", "AcCode", "N")
End Sub
Private Sub FillTempFromTxt()
'-- Tax AcName
txttemp(8) = GProcGetColumnValue("TblMastAccount", "AcCode", txt(8), "N", "ACName", "S")
txttemp(9) = GProcGetColumnValue("TblMastAccount", "AcCode", txt(9), "N", "ACName", "S")

txttemp(17) = GProcGetColumnValue("TblMastAccount", "AcCode", txt(17), "N", "ACName", "S")
txttemp(18) = GProcGetColumnValue("TblMastAccount", "AcCode", txt(18), "N", "ACName", "S")
txttemp(19) = GProcGetColumnValue("TblMastAccount", "AcCode", txt(19), "N", "ACName", "S")
txttemp(20) = GProcGetColumnValue("TblMastAccount", "AcCode", txt(20), "N", "ACName", "S")
txttemp(21) = GProcGetColumnValue("TblMastAccount", "AcCode", txt(21), "N", "ACName", "S")
txttemp(22) = GProcGetColumnValue("TblMastAccount", "AcCode", txt(22), "N", "ACName", "S")
txttemp(23) = GProcGetColumnValue("TblMastAccount", "AcCode", txt(23), "N", "ACName", "S")
End Sub
Private Sub ReadFields()
GProcShowRecord Me, rstbl, MaxNo
FillTempFromTxt
End Sub

Private Sub txttemp_GotFocus(Index As Integer)
GProcSelectBox Me.ActiveControl
End Sub

Private Sub txttemp_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
Dim GName As String
If KeyCode = vbKeyF3 Then
    Select Case Index
        Case 8, 9, 17, 18, 19, 20, 21, 22, 23 '-- Party
            'GName = GProcGetColumnValue("TblMastAccount", "AgCode", 15, "N", "AgName", "S")
            GProcShowForm frmMastAccount, frmMain.mnuMstfrm(0)
    End Select
End If
End Sub

Private Sub txtTemp_KeyPress(Index As Integer, KeyAscii As Integer)
Dim LeftPos As Long
Dim TopPos As Long
TopPos = txttemp(Index).Top + Me.Top + 650
LeftPos = txttemp(Index).Left + Me.Left
If KeyAscii = 13 Then Exit Sub
Select Case Index
    Case 8, 9, 17, 18, 19, 20, 21, 22 '-- Tax Account ,Alias Name
        gClsSearch.SearchMultiField "tblMastAccount", "AcName,Agcode", Array("Account Name", "AcCode"), Array(txttemp(Index).Width, 0), " (Agcode=90044 )  ", Trim(Chr(KeyAscii)), "", LeftPos, TopPos, False, , False
        KeyAscii = 0
        If gClsSearch.SearchMultiRetCol(0) <> "" Then
            txttemp(Index) = gClsSearch.SearchMultiRetCol(0)
        End If
    Case 23  '-- Tax Account ,Alias Name
        gClsSearch.SearchMultiField "tblMastAccount", "AcName,Agcode", Array("Account Name", "AcCode"), Array(txttemp(Index).Width, 0), " (Agcode=90023 )  ", Trim(Chr(KeyAscii)), "", LeftPos, TopPos, False, , False
        KeyAscii = 0
        If gClsSearch.SearchMultiRetCol(0) <> "" Then
            txttemp(Index) = gClsSearch.SearchMultiRetCol(0)
        End If
End Select
End Sub
Private Sub SetControlEd()
If txt(2) = "G S T" Then
   For i = 11 To 16
       txt(i).Enabled = True
   Next i
   txt(7).Enabled = False
   txt(10).Enabled = False
   txttemp(8).Enabled = False
   txttemp(9).Enabled = False
   For i = 17 To 22
       txttemp(i).Enabled = True
   Next i
Else
   For i = 11 To 16
       txt(i).Enabled = False
   Next i
   txt(7).Enabled = True
   txt(10).Enabled = True
   txttemp(8).Enabled = True
   txttemp(9).Enabled = True
   For i = 17 To 22
       txttemp(i).Enabled = False
   Next i
End If
End Sub

