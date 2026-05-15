VERSION 5.00
Begin VB.Form frmMastNarrationTaxGST 
   BackColor       =   &H00C2E3FC&
   BorderStyle     =   0  'None
   Caption         =   " Tax Master"
   ClientHeight    =   7200
   ClientLeft      =   0
   ClientTop       =   60
   ClientWidth     =   10725
   FillColor       =   &H000040C0&
   FillStyle       =   0  'Solid
   Icon            =   "frmMastNarrationTaxGST.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7200
   ScaleWidth      =   10725
   ShowInTaskbar   =   0   'False
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
      Index           =   23
      Left            =   11475
      MaxLength       =   40
      TabIndex        =   59
      Text            =   "23"
      Top             =   1395
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
      Index           =   22
      Left            =   11520
      MaxLength       =   40
      TabIndex        =   58
      Text            =   "22"
      Top             =   990
      Visible         =   0   'False
      Width           =   420
   End
   Begin VB.TextBox txtTemp 
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
      Index           =   23
      Left            =   2745
      MaxLength       =   40
      TabIndex        =   13
      Text            =   "txttemp(23)"
      ToolTipText     =   "Press F3 For Creation"
      Top             =   6210
      Width           =   5865
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
      Index           =   21
      Left            =   11385
      MaxLength       =   40
      TabIndex        =   56
      Text            =   "21"
      Top             =   585
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
      Index           =   20
      Left            =   10935
      MaxLength       =   40
      TabIndex        =   55
      Text            =   "20"
      Top             =   2160
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
      Index           =   19
      Left            =   10890
      MaxLength       =   40
      TabIndex        =   54
      Text            =   "19"
      Top             =   1800
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
      Index           =   18
      Left            =   10890
      MaxLength       =   40
      TabIndex        =   53
      Text            =   "18"
      Top             =   1395
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
      Index           =   17
      Left            =   10845
      MaxLength       =   40
      TabIndex        =   52
      Text            =   "17"
      Top             =   1035
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
      Index           =   16
      Left            =   10845
      MaxLength       =   40
      TabIndex        =   51
      Text            =   "16"
      Top             =   585
      Visible         =   0   'False
      Width           =   420
   End
   Begin VB.TextBox txtTemp 
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
      Index           =   18
      Left            =   5175
      MaxLength       =   40
      TabIndex        =   12
      Text            =   "txttemp(18)"
      ToolTipText     =   "Press F3 For Creation"
      Top             =   5130
      Width           =   3615
   End
   Begin VB.TextBox txtTemp 
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
      Left            =   5175
      MaxLength       =   40
      TabIndex        =   11
      Text            =   "txttemp(17)"
      ToolTipText     =   "Press F3 For Creation"
      Top             =   4590
      Width           =   3615
   End
   Begin VB.TextBox txtTemp 
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
      Index           =   16
      Left            =   5175
      MaxLength       =   40
      TabIndex        =   10
      Text            =   "txttemp(16)"
      ToolTipText     =   "Press F3 For Creation"
      Top             =   4095
      Width           =   3615
   End
   Begin VB.TextBox txtTemp 
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
      Index           =   21
      Left            =   1350
      MaxLength       =   40
      TabIndex        =   9
      Text            =   "txttemp(21)"
      ToolTipText     =   "Press F3 For Creation"
      Top             =   5130
      Width           =   3615
   End
   Begin VB.TextBox txtTemp 
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
      Left            =   1350
      MaxLength       =   40
      TabIndex        =   8
      Text            =   "txttemp(20)"
      ToolTipText     =   "Press F3 For Creation"
      Top             =   4590
      Width           =   3615
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
      Index           =   15
      Left            =   5175
      MaxLength       =   40
      TabIndex        =   6
      Text            =   "15"
      Top             =   2550
      Width           =   915
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
      Index           =   14
      Left            =   5175
      MaxLength       =   40
      TabIndex        =   5
      Text            =   "14"
      Top             =   2055
      Width           =   915
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
      Index           =   13
      Left            =   5175
      MaxLength       =   40
      TabIndex        =   4
      Text            =   "13"
      Top             =   1560
      Width           =   915
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
      Height          =   360
      Index           =   12
      Left            =   1935
      MaxLength       =   40
      TabIndex        =   3
      Text            =   "12"
      Top             =   2565
      Width           =   915
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
      Height          =   360
      Index           =   11
      Left            =   1935
      MaxLength       =   40
      TabIndex        =   2
      Text            =   "11"
      Top             =   2070
      Width           =   915
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
      Height          =   360
      Index           =   10
      Left            =   1935
      MaxLength       =   40
      TabIndex        =   1
      Text            =   "10"
      Top             =   1575
      Width           =   915
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
      Left            =   12150
      Style           =   1  'Graphical
      TabIndex        =   44
      Top             =   4005
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
      Left            =   10845
      MaxLength       =   40
      TabIndex        =   19
      Text            =   "txttemp(9)"
      ToolTipText     =   "Press F3 For Creation"
      Top             =   3330
      Visible         =   0   'False
      Width           =   1050
   End
   Begin VB.TextBox txtTemp 
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
      Index           =   19
      Left            =   1350
      MaxLength       =   40
      TabIndex        =   7
      Text            =   "txttemp(19)"
      ToolTipText     =   "Press F3 For Creation"
      Top             =   4095
      Width           =   3615
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
      Left            =   12375
      MaxLength       =   40
      TabIndex        =   43
      Text            =   "9"
      Top             =   3465
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
      Left            =   12465
      MaxLength       =   40
      TabIndex        =   42
      Text            =   "8"
      Top             =   2970
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
      Left            =   10980
      MaxLength       =   40
      TabIndex        =   16
      Text            =   "7"
      Top             =   2700
      Visible         =   0   'False
      Width           =   720
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
      Left            =   12600
      MaxLength       =   40
      TabIndex        =   40
      Text            =   "1"
      Top             =   135
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
      Left            =   12420
      TabIndex        =   39
      Text            =   "0"
      Top             =   540
      Visible         =   0   'False
      Width           =   795
   End
   Begin VB.TextBox txt 
      Height          =   360
      Index           =   6
      Left            =   12600
      TabIndex        =   35
      Text            =   "6"
      Top             =   2565
      Visible         =   0   'False
      Width           =   690
   End
   Begin VB.TextBox txt 
      Height          =   360
      Index           =   5
      Left            =   12555
      TabIndex        =   34
      Text            =   "5"
      Top             =   2025
      Visible         =   0   'False
      Width           =   690
   End
   Begin VB.TextBox txt 
      Height          =   360
      Index           =   4
      Left            =   12600
      TabIndex        =   33
      Text            =   "4"
      Top             =   1530
      Visible         =   0   'False
      Width           =   690
   End
   Begin VB.TextBox txt 
      Height          =   360
      Index           =   3
      Left            =   12690
      TabIndex        =   32
      Text            =   "3"
      Top             =   945
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
      Left            =   11880
      TabIndex        =   14
      Text            =   "2"
      Top             =   3870
      Visible         =   0   'False
      Width           =   375
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
      Left            =   12240
      Style           =   1  'Graphical
      TabIndex        =   37
      Top             =   6165
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
      Left            =   9390
      Style           =   1  'Graphical
      TabIndex        =   31
      ToolTipText     =   "Exit"
      Top             =   5715
      Width           =   870
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
      Left            =   9390
      Style           =   1  'Graphical
      TabIndex        =   24
      ToolTipText     =   "To Cancel Record"
      Top             =   5175
      Width           =   870
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
      Left            =   9390
      Style           =   1  'Graphical
      TabIndex        =   23
      ToolTipText     =   "To Save Record"
      Top             =   4635
      Width           =   870
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
      Left            =   12015
      Style           =   1  'Graphical
      TabIndex        =   36
      Top             =   5580
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
      Left            =   9390
      Style           =   1  'Graphical
      TabIndex        =   22
      ToolTipText     =   "To Delete Record"
      Top             =   3690
      Width           =   870
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
      Left            =   9390
      Style           =   1  'Graphical
      TabIndex        =   21
      ToolTipText     =   "To Modify Record"
      Top             =   3150
      Width           =   870
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
      Left            =   9390
      Style           =   1  'Graphical
      TabIndex        =   20
      ToolTipText     =   "To Add New Record"
      Top             =   2655
      Width           =   870
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
      Left            =   12285
      Style           =   1  'Graphical
      TabIndex        =   26
      Top             =   4995
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
      Left            =   9390
      Style           =   1  'Graphical
      TabIndex        =   25
      ToolTipText     =   "To Find Record"
      Top             =   1800
      Width           =   870
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
      Left            =   9855
      Style           =   1  'Graphical
      TabIndex        =   30
      ToolTipText     =   "Last"
      Top             =   1305
      Width           =   405
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
      Left            =   9390
      Style           =   1  'Graphical
      TabIndex        =   29
      ToolTipText     =   "First"
      Top             =   1305
      Width           =   405
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
      Left            =   9855
      Style           =   1  'Graphical
      TabIndex        =   28
      ToolTipText     =   "Next"
      Top             =   810
      Width           =   405
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
      Left            =   9390
      Style           =   1  'Graphical
      TabIndex        =   27
      ToolTipText     =   "Previous"
      Top             =   810
      Width           =   405
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
      Height          =   285
      Index           =   10
      Left            =   540
      TabIndex        =   57
      Top             =   6255
      Width           =   2760
   End
   Begin VB.Shape Shape5 
      BackColor       =   &H00808080&
      BorderColor     =   &H004D177D&
      FillColor       =   &H00ECFFFF&
      FillStyle       =   0  'Solid
      Height          =   930
      Left            =   315
      Top             =   5940
      Width           =   8670
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
      Index           =   9
      Left            =   495
      TabIndex        =   50
      Top             =   5220
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
      Index           =   8
      Left            =   495
      TabIndex        =   49
      Top             =   4680
      Width           =   810
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
      Index           =   7
      Left            =   495
      TabIndex        =   48
      Top             =   4140
      Width           =   810
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
      Index           =   4
      Left            =   5175
      TabIndex        =   47
      Top             =   1125
      Width           =   2055
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
      Left            =   480
      TabIndex        =   46
      Top             =   2610
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
      Index           =   1
      Left            =   480
      TabIndex        =   45
      Top             =   2115
      Width           =   1530
   End
   Begin VB.Shape Shape4 
      BorderColor     =   &H00C00000&
      BorderWidth     =   5
      FillColor       =   &H00C00000&
      Height          =   7110
      Left            =   0
      Top             =   45
      Width           =   10695
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
      TabIndex        =   41
      Top             =   60
      Width           =   2475
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
      Index           =   0
      Left            =   1920
      TabIndex        =   0
      Top             =   1140
      Width           =   2055
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
      Index           =   6
      Left            =   5175
      TabIndex        =   18
      Top             =   3690
      Width           =   2760
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
      Index           =   5
      Left            =   1350
      TabIndex        =   17
      Top             =   3690
      Width           =   3585
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
      Index           =   3
      Left            =   480
      TabIndex        =   15
      Top             =   1575
      Width           =   1530
   End
   Begin VB.Label Label6 
      BackColor       =   &H00C00000&
      Caption         =   "   Tax Master (GST)"
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
      TabIndex        =   38
      Top             =   45
      Width           =   10695
   End
   Begin VB.Shape Shape1 
      FillColor       =   &H00BDDFD1&
      FillStyle       =   0  'Solid
      Height          =   6075
      Left            =   9150
      Shape           =   4  'Rounded Rectangle
      Top             =   570
      Width           =   1335
   End
   Begin VB.Shape Shape2 
      BackColor       =   &H00808080&
      BorderColor     =   &H004D177D&
      FillColor       =   &H00ECFFFF&
      FillStyle       =   0  'Solid
      Height          =   2595
      Left            =   240
      Top             =   660
      Width           =   8715
   End
   Begin VB.Shape Shape3 
      BackColor       =   &H00808080&
      BorderColor     =   &H004D177D&
      FillColor       =   &H00ECFFFF&
      FillStyle       =   0  'Solid
      Height          =   2280
      Left            =   270
      Top             =   3465
      Width           =   8670
   End
End
Attribute VB_Name = "frmMastNarrationTaxGST"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Const tblName As String = "tblMastNarration"
Const IdField As String = "narrcode"
Const OrderField As String = "Narration"
Const MaxNo As Long = 23
'Const CriteriaStr As String = " narrtp = '" & gfrmTypeStr & "'"
Dim rstbl As Recordset
'Dim varBookMark As Variant
Dim FirstTimeFlag As Boolean
Dim FormAction As Integer

Private Sub cmdBtn_Click(Index As Integer)
Dim FindQStr As String
Dim X As String
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
    FindQStr = "Select Narration from " _
    & " tblMastNarration where tblMastNarration.Narration = 'G S T' and NarrType='T'"
    gClsSearch.SearchMultiField FindQStr, "", Array("Tax Name"), Array(3000), "", "", "Narration", txt(10).Left + Me.Left, txt(10).Top + Me.Top + 650, False, , True
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
        txt(10).SetFocus
    End If
    txt(2) = "G S T"
Case 7 'Modify
    'If mURecModi = True Then
    '    SearchHelp
    '    DoEvents
        cmdBtn_Click 4 '-- Find
        If gClsSearch.SearchMultiRetCol(0) = "" Then Exit Sub
        FormAction = vbDataActionUpdate
        Call GProcActivateControls(Me, True)
        Call GProcSetButtons(Me, vbDataActionUpdate)
        txt(10).SetFocus
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
'Case 12 'Creation To Firm
'     gfrmTypeStr = "Tax"
'     frmUtlCreationToFirm.Show 1
'     Rstbl.Requery
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
    X = "Select * from " & tblName & " where NarrType='T' and Narration = 'G S T'"
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
SaveBtnEd
End Sub
Private Sub txt_KeyPress(Index As Integer, KeyAscii As Integer)
If KeyAscii = 13 Then Exit Sub
KeyAscii = GProcValidateKey(rstbl, Index, KeyAscii, txt(Index).text)
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
        txt(10).SetFocus
        Exit Function
    End If
    '--- Tax Account Purchase
    If txttemp(16) = "" Then txttemp(16) = "-"
    If GProcGetColumnValue("tblMastAccount", "AcName", txttemp(16), "S", "AcCode", "N") = 0 Then
        GProcShowForm frmMastAccount, frmMain.mnuMstfrm(0), Array(frmMastAccount.txt(1)), Array(txttemp(16))
        Exit Function
    End If
    If txttemp(17) = "" Then txttemp(17) = "-"
    If GProcGetColumnValue("tblMastAccount", "AcName", txttemp(17), "S", "AcCode", "N") = 0 Then
        GProcShowForm frmMastAccount, frmMain.mnuMstfrm(0), Array(frmMastAccount.txt(1)), Array(txttemp(17))
        Exit Function
    End If
    If txttemp(18) = "" Then txttemp(18) = "-"
    If GProcGetColumnValue("tblMastAccount", "AcName", txttemp(18), "S", "AcCode", "N") = 0 Then
        GProcShowForm frmMastAccount, frmMain.mnuMstfrm(0), Array(frmMastAccount.txt(1)), Array(txttemp(18))
        Exit Function
    End If
    If txttemp(19) = "" Then txttemp(19) = "-"
    If GProcGetColumnValue("tblMastAccount", "AcName", txttemp(19), "S", "AcCode", "N") = 0 Then
        GProcShowForm frmMastAccount, frmMain.mnuMstfrm(0), Array(frmMastAccount.txt(1)), Array(txttemp(19))
        Exit Function
    End If
    If txttemp(20) = "" Then txttemp(20) = "-"
    If GProcGetColumnValue("tblMastAccount", "AcName", txttemp(20), "S", "AcCode", "N") = 0 Then
        GProcShowForm frmMastAccount, frmMain.mnuMstfrm(0), Array(frmMastAccount.txt(1)), Array(txttemp(20))
        Exit Function
    End If
    If txttemp(21) = "" Then txttemp(21) = "-"
    If GProcGetColumnValue("tblMastAccount", "AcName", txttemp(21), "S", "AcCode", "N") = 0 Then
        GProcShowForm frmMastAccount, frmMain.mnuMstfrm(0), Array(frmMastAccount.txt(1)), Array(txttemp(21))
        Exit Function
    End If
    '--- Narration
    If Trim(txt(2)) = "" Then txt(2) = "G S T"
    '----Duplicate Records
    '----Addition Case
    If FormAction = vbDataActionAddNew Then
         X = "Select NArration From " & tblName & " Where Narration='" & Trim(txt(2).text) & "'"
        i = GProcRstOpen(rstAddCheck, X, "R")
        If i > 0 Then
            MsgBox "Record is Available of Same Narration  " & Trim(txt(2).text), vbCritical + vbOKOnly, "Duplicate Record"
            txt(10).SetFocus
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
            txt(10).SetFocus
            Exit Function
        End If
        rstModiCheck.Close
    End If
    ValidateData = True
End Function
Private Sub FillTxtFromTemp()
txt(0) = "T"
txt(2) = "G S T"
txt(8) = GProcGetColumnValue("TblMastAccount", "ACName", "-", "S", "AcCode", "N")
txt(9) = GProcGetColumnValue("TblMastAccount", "ACName", "-", "S", "AcCode", "N")

'-- Tax AcCode
txt(16) = GProcGetColumnValue("TblMastAccount", "ACName", txttemp(16), "S", "AcCode", "N")
txt(17) = GProcGetColumnValue("TblMastAccount", "ACName", txttemp(17), "S", "AcCode", "N")
txt(18) = GProcGetColumnValue("TblMastAccount", "ACName", txttemp(18), "S", "AcCode", "N")
txt(19) = GProcGetColumnValue("TblMastAccount", "ACName", txttemp(19), "S", "AcCode", "N")
txt(20) = GProcGetColumnValue("TblMastAccount", "ACName", txttemp(20), "S", "AcCode", "N")
txt(21) = GProcGetColumnValue("TblMastAccount", "ACName", txttemp(21), "S", "AcCode", "N")
txt(23) = GProcGetColumnValue("TblMastAccount", "ACName", txttemp(23), "S", "AcCode", "N")
End Sub
Private Sub FillTempFromTxt()
'-- Tax AcName
txttemp(16) = GProcGetColumnValue("TblMastAccount", "AcCode", txt(16), "N", "ACName", "S")
txttemp(17) = GProcGetColumnValue("TblMastAccount", "AcCode", txt(17), "N", "ACName", "S")
txttemp(18) = GProcGetColumnValue("TblMastAccount", "AcCode", txt(18), "N", "ACName", "S")
txttemp(19) = GProcGetColumnValue("TblMastAccount", "AcCode", txt(19), "N", "ACName", "S")
txttemp(20) = GProcGetColumnValue("TblMastAccount", "AcCode", txt(20), "N", "ACName", "S")
txttemp(21) = GProcGetColumnValue("TblMastAccount", "AcCode", txt(21), "N", "ACName", "S")
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
        Case 16, 17, 18, 19, 20, 21, 23 '-- Party
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
    Case 16, 17, 18, 19, 20, 21, 23 '-- Tax Account ,Alias Name
        gClsSearch.SearchMultiField "QryMastaccount", "Account,AcName,AgName", Array("Party", "AcName", "Group"), Array(txttemp(Index).Width, 0, 3000), " (Agcode=90038 or agcode = 90044)", Chr(KeyAscii), "acName", LeftPos, TopPos, True
        KeyAscii = 0
        If gClsSearch.SearchMultiRetCol(0) <> "" And gClsSearch.SearchMultiRetCol(1) = "" Then
            txttemp(Index) = gClsSearch.SearchMultiRetCol(0)
        Else
           If gClsSearch.SearchMultiRetCol(1) <> "" Then txttemp(Index) = gClsSearch.SearchMultiRetCol(1)
        End If
    Case 23 '-- rcm sale Account ,Alias Name
        gClsSearch.SearchMultiField "QryMastaccount", "Account,AcName,AgName", Array("Party", "AcName", "Group"), Array(txttemp(Index).Width, 0, 3000), " (Agcode=90023 )", Chr(KeyAscii), "acName", LeftPos, TopPos, True
        KeyAscii = 0
        If gClsSearch.SearchMultiRetCol(0) <> "" And gClsSearch.SearchMultiRetCol(1) = "" Then
            txttemp(Index) = gClsSearch.SearchMultiRetCol(0)
        Else
           If gClsSearch.SearchMultiRetCol(1) <> "" Then txttemp(Index) = gClsSearch.SearchMultiRetCol(1)
        End If
End Select
End Sub
