VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form frmLatePaymentIntCreditGST 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00FBF2E1&
   BorderStyle     =   0  'None
   Caption         =   "Late Payment Interest List"
   ClientHeight    =   9405
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   17655
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmLatePaymentIntCreditGST.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   9405
   ScaleWidth      =   17655
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox txt 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   325
      Index           =   57
      Left            =   18000
      TabIndex        =   119
      Text            =   "57"
      Top             =   4860
      Visible         =   0   'False
      Width           =   360
   End
   Begin VB.TextBox txt 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   325
      Index           =   56
      Left            =   18960
      TabIndex        =   115
      Text            =   "56"
      Top             =   8280
      Visible         =   0   'False
      Width           =   360
   End
   Begin VB.TextBox txt 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   325
      Index           =   55
      Left            =   18960
      TabIndex        =   114
      Text            =   "55"
      Top             =   7875
      Visible         =   0   'False
      Width           =   360
   End
   Begin VB.TextBox txt 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   325
      Index           =   54
      Left            =   18960
      TabIndex        =   113
      Text            =   "54"
      Top             =   7470
      Visible         =   0   'False
      Width           =   360
   End
   Begin VB.TextBox txt 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   325
      Index           =   53
      Left            =   18960
      TabIndex        =   112
      Text            =   "53"
      Top             =   7065
      Visible         =   0   'False
      Width           =   360
   End
   Begin VB.TextBox txt 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   325
      Index           =   52
      Left            =   18960
      TabIndex        =   111
      Text            =   "52"
      Top             =   6705
      Visible         =   0   'False
      Width           =   360
   End
   Begin VB.TextBox txt 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   325
      Index           =   51
      Left            =   18960
      TabIndex        =   110
      Text            =   "51"
      Top             =   6390
      Visible         =   0   'False
      Width           =   360
   End
   Begin VB.TextBox txt 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   325
      Index           =   50
      Left            =   18960
      TabIndex        =   109
      Text            =   "50"
      Top             =   5985
      Visible         =   0   'False
      Width           =   360
   End
   Begin VB.TextBox txt 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   325
      Index           =   49
      Left            =   18240
      TabIndex        =   108
      Text            =   "49"
      Top             =   8730
      Visible         =   0   'False
      Width           =   360
   End
   Begin VB.TextBox txt 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   325
      Index           =   48
      Left            =   18240
      TabIndex        =   107
      Text            =   "48"
      Top             =   8370
      Visible         =   0   'False
      Width           =   360
   End
   Begin VB.CheckBox chkIsGSTDbNt 
      BackColor       =   &H00ECFFFF&
      Caption         =   "Is GST Debit Note ?"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C000C0&
      Height          =   240
      Left            =   4860
      MaskColor       =   &H00ECFFFF&
      TabIndex        =   7
      Top             =   1665
      Width           =   2805
   End
   Begin VB.TextBox txt 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   325
      Index           =   47
      Left            =   18330
      TabIndex        =   105
      Text            =   "47"
      Top             =   8010
      Visible         =   0   'False
      Width           =   360
   End
   Begin VB.TextBox txt 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   325
      Index           =   46
      Left            =   18240
      TabIndex        =   98
      Text            =   "46"
      Top             =   7650
      Visible         =   0   'False
      Width           =   360
   End
   Begin VB.TextBox txt 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   325
      Index           =   45
      Left            =   18240
      TabIndex        =   97
      Text            =   "45"
      Top             =   7200
      Visible         =   0   'False
      Width           =   360
   End
   Begin VB.TextBox txt 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   325
      Index           =   44
      Left            =   18240
      TabIndex        =   96
      Text            =   "44"
      Top             =   6750
      Visible         =   0   'False
      Width           =   360
   End
   Begin VB.CheckBox chkIsMonthlyDbNt 
      BackColor       =   &H00ECFFFF&
      Caption         =   "Is Monthly Debit Note ?"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C000C0&
      Height          =   240
      Left            =   1845
      MaskColor       =   &H00ECFFFF&
      TabIndex        =   6
      Top             =   1665
      Width           =   2805
   End
   Begin VB.TextBox txt 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   325
      Index           =   43
      Left            =   18240
      TabIndex        =   93
      Text            =   "43"
      Top             =   6300
      Visible         =   0   'False
      Width           =   360
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
      Index           =   42
      Left            =   6810
      TabIndex        =   11
      Text            =   "42"
      Top             =   2025
      Width           =   945
   End
   Begin VB.TextBox txt 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   325
      Index           =   41
      Left            =   18285
      TabIndex        =   91
      Text            =   "41"
      Top             =   5985
      Visible         =   0   'False
      Width           =   360
   End
   Begin VB.TextBox txt 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   325
      Index           =   40
      Left            =   18285
      TabIndex        =   90
      Text            =   "40"
      Top             =   5625
      Visible         =   0   'False
      Width           =   360
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
      Index           =   41
      Left            =   1845
      TabIndex        =   5
      Text            =   "txttemp(41)"
      ToolTipText     =   "Press F3 for Creation"
      Top             =   1215
      Width           =   5910
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
      Index           =   39
      Left            =   3075
      TabIndex        =   9
      Text            =   "39"
      Top             =   2025
      Visible         =   0   'False
      Width           =   765
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
      Index           =   38
      Left            =   4830
      TabIndex        =   10
      Text            =   "38"
      Top             =   2025
      Width           =   945
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
      Index           =   37
      Left            =   1860
      TabIndex        =   8
      Text            =   "37"
      Top             =   2025
      Width           =   765
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
      Index           =   36
      Left            =   18285
      TabIndex        =   12
      Text            =   "36"
      Top             =   5220
      Visible         =   0   'False
      Width           =   405
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
      Index           =   35
      Left            =   1845
      TabIndex        =   1
      Text            =   "35"
      Top             =   765
      Width           =   1125
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
      Height          =   325
      Index           =   34
      Left            =   10035
      TabIndex        =   83
      Text            =   "34"
      Top             =   8865
      Visible         =   0   'False
      Width           =   1440
   End
   Begin VB.TextBox txt 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   325
      Index           =   33
      Left            =   18285
      TabIndex        =   82
      Text            =   "33"
      Top             =   4410
      Visible         =   0   'False
      Width           =   360
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
      Height          =   325
      Index           =   32
      Left            =   2250
      TabIndex        =   81
      Text            =   "32"
      Top             =   8865
      Visible         =   0   'False
      Width           =   1620
   End
   Begin VB.TextBox txt 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   325
      Index           =   31
      Left            =   18285
      TabIndex        =   80
      Text            =   "31"
      Top             =   3690
      Visible         =   0   'False
      Width           =   360
   End
   Begin VB.TextBox txt 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   325
      Index           =   30
      Left            =   18240
      TabIndex        =   79
      Text            =   "30"
      Top             =   3330
      Visible         =   0   'False
      Width           =   360
   End
   Begin VB.TextBox txt 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   325
      Index           =   29
      Left            =   18195
      TabIndex        =   78
      Text            =   "29"
      Top             =   2970
      Visible         =   0   'False
      Width           =   360
   End
   Begin VB.TextBox txt 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   325
      Index           =   28
      Left            =   18285
      TabIndex        =   77
      Text            =   "28"
      Top             =   2655
      Visible         =   0   'False
      Width           =   360
   End
   Begin VB.TextBox txt 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   325
      Index           =   27
      Left            =   18285
      TabIndex        =   76
      Text            =   "27"
      Top             =   2295
      Visible         =   0   'False
      Width           =   360
   End
   Begin VB.TextBox txt 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   325
      Index           =   26
      Left            =   18240
      TabIndex        =   75
      Text            =   "26"
      Top             =   1935
      Visible         =   0   'False
      Width           =   360
   End
   Begin VB.TextBox txt 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   325
      Index           =   25
      Left            =   18240
      TabIndex        =   74
      Text            =   "25"
      Top             =   1620
      Visible         =   0   'False
      Width           =   360
   End
   Begin VB.TextBox txt 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   325
      Index           =   24
      Left            =   18240
      TabIndex        =   73
      Text            =   "24"
      Top             =   1260
      Visible         =   0   'False
      Width           =   360
   End
   Begin VB.TextBox txt 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   325
      Index           =   23
      Left            =   18240
      TabIndex        =   72
      Text            =   "23"
      Top             =   900
      Visible         =   0   'False
      Width           =   360
   End
   Begin VB.TextBox txt 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   325
      Index           =   22
      Left            =   18240
      TabIndex        =   71
      Text            =   "22"
      Top             =   540
      Visible         =   0   'False
      Width           =   360
   End
   Begin VB.TextBox txt 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   325
      Index           =   21
      Left            =   18240
      TabIndex        =   70
      Text            =   "21"
      Top             =   180
      Visible         =   0   'False
      Width           =   360
   End
   Begin VB.TextBox txt 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   325
      Index           =   20
      Left            =   20310
      TabIndex        =   69
      Text            =   "20"
      Top             =   4770
      Visible         =   0   'False
      Width           =   360
   End
   Begin VB.TextBox txt 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   325
      Index           =   19
      Left            =   20265
      TabIndex        =   68
      Text            =   "19"
      Top             =   4410
      Visible         =   0   'False
      Width           =   360
   End
   Begin VB.TextBox txt 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   325
      Index           =   18
      Left            =   20265
      TabIndex        =   67
      Text            =   "18"
      Top             =   4050
      Visible         =   0   'False
      Width           =   360
   End
   Begin VB.TextBox txt 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   325
      Index           =   17
      Left            =   20220
      TabIndex        =   66
      Text            =   "17"
      Top             =   3735
      Visible         =   0   'False
      Width           =   360
   End
   Begin VB.TextBox txt 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   325
      Index           =   12
      Left            =   20175
      TabIndex        =   65
      Text            =   "12"
      Top             =   3375
      Visible         =   0   'False
      Width           =   360
   End
   Begin VB.TextBox txt 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   325
      Index           =   11
      Left            =   20175
      TabIndex        =   64
      Text            =   "11"
      Top             =   3015
      Visible         =   0   'False
      Width           =   360
   End
   Begin VB.TextBox txt 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   325
      Index           =   9
      Left            =   20220
      TabIndex        =   63
      Text            =   "9"
      Top             =   2655
      Visible         =   0   'False
      Width           =   330
   End
   Begin VB.TextBox txt 
      Alignment       =   1  'Right Justify
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Index           =   6
      Left            =   20220
      TabIndex        =   62
      Text            =   "6"
      Top             =   1935
      Visible         =   0   'False
      Width           =   345
   End
   Begin VB.TextBox txt 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   375
      Index           =   16
      Left            =   19590
      MaxLength       =   50
      TabIndex        =   21
      Text            =   "16"
      ToolTipText     =   "Press F3 for Creation & Press F2 for Selection"
      Top             =   1800
      Visible         =   0   'False
      Width           =   315
   End
   Begin VB.TextBox txt 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   375
      Index           =   15
      Left            =   19590
      MaxLength       =   50
      TabIndex        =   20
      Text            =   "15"
      ToolTipText     =   "Press F3 for Creation & Press F2 for Selection"
      Top             =   1395
      Visible         =   0   'False
      Width           =   360
   End
   Begin VB.CheckBox chkDontShowInUnadj 
      BackColor       =   &H00FFC0FF&
      Caption         =   "Dont Show In Unadjusted Receipt"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   6345
      TabIndex        =   61
      Top             =   9540
      Width           =   4515
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
      Height          =   3615
      Left            =   2880
      TabIndex        =   52
      Top             =   4095
      Visible         =   0   'False
      Width           =   8070
      Begin VB.Frame Frame2 
         Height          =   1275
         Left            =   4380
         TabIndex        =   116
         Top             =   420
         Width           =   2955
         Begin VB.OptionButton OptCopy 
            Caption         =   "Customer Copy"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   0
            Left            =   300
            TabIndex        =   118
            Top             =   300
            Width           =   2115
         End
         Begin VB.OptionButton OptCopy 
            Caption         =   "Office Copy"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   1
            Left            =   300
            TabIndex        =   117
            Top             =   840
            Width           =   2115
         End
      End
      Begin VB.OptionButton OptType 
         BackColor       =   &H00FFC0C0&
         Caption         =   "Int. Credit Note "
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
         Height          =   330
         Index           =   1
         Left            =   3915
         TabIndex        =   95
         Top             =   1980
         Width           =   1815
      End
      Begin VB.OptionButton OptType 
         BackColor       =   &H00FFC0C0&
         Caption         =   "Int. Credit Note Check List"
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
         Height          =   330
         Index           =   0
         Left            =   495
         TabIndex        =   94
         Top             =   1980
         Width           =   2985
      End
      Begin VB.CommandButton cmdPrint 
         Caption         =   "S&hort Note"
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
         Left            =   6960
         TabIndex        =   60
         Top             =   2820
         Visible         =   0   'False
         Width           =   1335
      End
      Begin VB.TextBox txtno 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   0
         Left            =   1950
         MaxLength       =   7
         TabIndex        =   57
         Top             =   420
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
         Left            =   1200
         TabIndex        =   56
         Top             =   2760
         Width           =   1335
      End
      Begin VB.TextBox txtno 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   1
         Left            =   1950
         MaxLength       =   7
         TabIndex        =   55
         Top             =   900
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
         Left            =   2940
         TabIndex        =   54
         Top             =   2760
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
         Left            =   4920
         TabIndex        =   53
         Top             =   2760
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
         TabIndex        =   59
         Top             =   420
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
         TabIndex        =   58
         Top             =   900
         Width           =   1050
      End
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
      Left            =   13185
      Style           =   1  'Graphical
      TabIndex        =   51
      ToolTipText     =   "Exit"
      Top             =   1710
      Width           =   1065
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
      Left            =   12015
      Style           =   1  'Graphical
      TabIndex        =   50
      ToolTipText     =   "To Cancel Record"
      Top             =   1215
      Width           =   1065
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
      Left            =   12015
      Style           =   1  'Graphical
      TabIndex        =   49
      ToolTipText     =   "To Save Record"
      Top             =   720
      Width           =   1065
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
      TabIndex        =   48
      Top             =   1710
      Width           =   1065
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
      Left            =   10890
      Style           =   1  'Graphical
      TabIndex        =   47
      ToolTipText     =   "To Delete Record"
      Top             =   1710
      Width           =   1065
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
      Left            =   10890
      Style           =   1  'Graphical
      TabIndex        =   46
      ToolTipText     =   "To Modify Record"
      Top             =   1215
      Width           =   1065
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
      Left            =   10890
      Style           =   1  'Graphical
      TabIndex        =   45
      ToolTipText     =   "To Add New Record"
      Top             =   720
      Width           =   1065
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
      Left            =   14310
      Style           =   1  'Graphical
      TabIndex        =   44
      ToolTipText     =   "To Find Record"
      Top             =   720
      Width           =   1065
   End
   Begin VB.CommandButton cmdIntDbNt 
      Caption         =   "&Interest Debit Note"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   900
      Left            =   14310
      TabIndex        =   43
      Top             =   1215
      Width           =   1065
   End
   Begin VB.TextBox txt 
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Index           =   10
      Left            =   19545
      TabIndex        =   17
      Text            =   "10"
      Top             =   2565
      Visible         =   0   'False
      Width           =   405
   End
   Begin VB.TextBox txtCrBalance 
      Enabled         =   0   'False
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
      Left            =   20670
      TabIndex        =   41
      Text            =   "txtCrBalance"
      Top             =   180
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.TextBox txtDrBalance 
      Enabled         =   0   'False
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
      Left            =   20670
      TabIndex        =   40
      Text            =   "txtDrBalance"
      Top             =   585
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.CommandButton cmdLedger 
      Caption         =   "Ledger"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   0
      Left            =   19440
      TabIndex        =   39
      Top             =   5040
      Visible         =   0   'False
      Width           =   1200
   End
   Begin VB.CommandButton cmdLedger 
      Caption         =   "Ledger"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   1
      Left            =   19560
      TabIndex        =   38
      Top             =   5535
      Visible         =   0   'False
      Width           =   1185
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
      Left            =   480
      TabIndex        =   14
      Top             =   4815
      Visible         =   0   'False
      Width           =   885
   End
   Begin VB.TextBox txt 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   375
      Index           =   13
      Left            =   19590
      MaxLength       =   50
      TabIndex        =   18
      Text            =   "13"
      ToolTipText     =   "Press F3 for Creation & Press F2 for Selection"
      Top             =   585
      Visible         =   0   'False
      Width           =   360
   End
   Begin VB.TextBox txt 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFFFF&
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
      Height          =   360
      Index           =   0
      Left            =   19590
      TabIndex        =   15
      Text            =   "0"
      Top             =   135
      Visible         =   0   'False
      Width           =   285
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
      Left            =   13185
      Style           =   1  'Graphical
      TabIndex        =   23
      ToolTipText     =   "Previous"
      Top             =   720
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
      Left            =   13755
      Style           =   1  'Graphical
      TabIndex        =   24
      ToolTipText     =   "Next"
      Top             =   720
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
      Left            =   13185
      Style           =   1  'Graphical
      TabIndex        =   25
      ToolTipText     =   "First"
      Top             =   1215
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
      Left            =   13755
      Style           =   1  'Graphical
      TabIndex        =   26
      ToolTipText     =   "Last"
      Top             =   1215
      Width           =   495
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
      Left            =   20160
      Style           =   1  'Graphical
      TabIndex        =   27
      Top             =   6600
      Visible         =   0   'False
      Width           =   975
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
      Left            =   19560
      Style           =   1  'Graphical
      TabIndex        =   31
      Top             =   8325
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
      Left            =   19680
      Style           =   1  'Graphical
      TabIndex        =   32
      Top             =   7800
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.TextBox txt 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Index           =   3
      Left            =   20175
      TabIndex        =   35
      Text            =   "3"
      Top             =   900
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.TextBox txt 
      Alignment       =   1  'Right Justify
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Index           =   5
      Left            =   20175
      TabIndex        =   30
      Text            =   "5"
      Top             =   1575
      Visible         =   0   'False
      Width           =   345
   End
   Begin VB.TextBox txt 
      Alignment       =   1  'Right Justify
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Index           =   8
      Left            =   19545
      TabIndex        =   16
      Text            =   "8"
      Top             =   2160
      Visible         =   0   'False
      Width           =   330
   End
   Begin VB.TextBox txt 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   375
      Index           =   14
      Left            =   19590
      MaxLength       =   50
      TabIndex        =   19
      Text            =   "14"
      ToolTipText     =   "Press F3 for Creation & Press F2 for Selection"
      Top             =   990
      Visible         =   0   'False
      Width           =   405
   End
   Begin VB.TextBox txt 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   325
      Index           =   7
      Left            =   20220
      TabIndex        =   28
      Text            =   "7"
      Top             =   2340
      Visible         =   0   'False
      Width           =   330
   End
   Begin VB.TextBox txt 
      Alignment       =   1  'Right Justify
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Index           =   4
      Left            =   20175
      TabIndex        =   29
      Text            =   "4"
      Top             =   1260
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.TextBox txt 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Index           =   2
      Left            =   20175
      TabIndex        =   34
      Text            =   "2"
      Top             =   540
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
      Left            =   20175
      TabIndex        =   33
      Text            =   "1"
      Top             =   180
      Visible         =   0   'False
      Width           =   375
   End
   Begin MSFlexGridLib.MSFlexGrid msGrid 
      Height          =   6060
      Left            =   75
      TabIndex        =   13
      Top             =   2745
      Width           =   17460
      _ExtentX        =   30798
      _ExtentY        =   10689
      _Version        =   393216
      FixedCols       =   0
      RowHeightMin    =   275
      ForeColor       =   0
      BackColorFixed  =   -2147483647
      ForeColorFixed  =   -2147483634
      BackColorBkg    =   12632256
      Appearance      =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin Crystal.CrystalReport cryReport 
      Bindings        =   "frmLatePaymentIntCreditGST.frx":058A
      Left            =   -45
      Top             =   660
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      WindowControlBox=   -1  'True
      WindowMaxButton =   -1  'True
      WindowMinButton =   -1  'True
      PrintFileType   =   15
      PrintFileLinesPerPage=   60
   End
   Begin MSComCtl2.DTPicker dtpDt 
      Height          =   360
      Index           =   2
      Left            =   8685
      TabIndex        =   4
      Top             =   780
      Width           =   1710
      _ExtentX        =   3016
      _ExtentY        =   635
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      CustomFormat    =   "dd/MM/yyyy"
      Format          =   115015683
      CurrentDate     =   -273
   End
   Begin MSComCtl2.DTPicker dtpDt 
      Height          =   360
      Index           =   40
      Left            =   6435
      TabIndex        =   3
      Top             =   765
      Width           =   1710
      _ExtentX        =   3016
      _ExtentY        =   635
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      CustomFormat    =   "dd/MM/yyyy"
      Format          =   115015683
      CurrentDate     =   -273
   End
   Begin MSComCtl2.DTPicker dtpDt 
      Height          =   360
      Index           =   11
      Left            =   3780
      TabIndex        =   2
      Top             =   765
      Width           =   1710
      _ExtentX        =   3016
      _ExtentY        =   635
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      CustomFormat    =   "dd/MM/yyyy"
      Format          =   115015683
      CurrentDate     =   -273
   End
   Begin VB.Label lbl1 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Date :"
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
      Index           =   10
      Left            =   3090
      TabIndex        =   106
      Top             =   810
      Width           =   555
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "+"
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
      Left            =   6120
      TabIndex        =   104
      Top             =   8910
      Width           =   120
   End
   Begin VB.Label lblInt2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Interest Amount 2"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C000C0&
      Height          =   240
      Left            =   6525
      TabIndex        =   103
      Top             =   8910
      Visible         =   0   'False
      Width           =   1665
   End
   Begin VB.Label lblTDS 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "T D S  Amount"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C000C0&
      Height          =   240
      Left            =   11565
      TabIndex        =   102
      Top             =   8910
      Visible         =   0   'False
      Width           =   2535
   End
   Begin VB.Label lblInt 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Interest Amount"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C000C0&
      Height          =   240
      Left            =   4005
      TabIndex        =   101
      Top             =   8910
      Visible         =   0   'False
      Width           =   1905
   End
   Begin VB.Label lbl1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "TDS JV NO   :"
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
      Index           =   9
      Left            =   8505
      TabIndex        =   100
      Top             =   8910
      Visible         =   0   'False
      Width           =   1275
   End
   Begin VB.Label lbl1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "INTEREST JV NO   :"
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
      Index           =   8
      Left            =   270
      TabIndex        =   99
      Top             =   8910
      Visible         =   0   'False
      Width           =   1830
   End
   Begin VB.Label lbl1 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "TDS @  :"
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
      Index           =   7
      Left            =   5895
      TabIndex        =   92
      Top             =   2070
      Width           =   810
   End
   Begin VB.Label lbl1 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "+"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   285
      Index           =   6
      Left            =   2790
      TabIndex        =   89
      Top             =   2055
      Visible         =   0   'False
      Width           =   135
   End
   Begin VB.Label lbl1 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Party  :"
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
      Left            =   945
      TabIndex        =   88
      Top             =   1260
      Width           =   675
   End
   Begin VB.Label lbl1 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "To  :"
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
      Left            =   8190
      TabIndex        =   87
      Top             =   840
      Width           =   420
   End
   Begin VB.Label lblApplyJV 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Interest JV Pending"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   990
      Left            =   9270
      TabIndex        =   86
      Top             =   1305
      Width           =   1155
      WordWrap        =   -1  'True
   End
   Begin VB.Label lbl1 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Int. @  :"
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
      Left            =   4005
      TabIndex        =   85
      Top             =   2070
      Width           =   705
   End
   Begin VB.Label lbl1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Grace Days  :"
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
      Index           =   2
      Left            =   390
      TabIndex        =   84
      Top             =   2070
      Width           =   1215
   End
   Begin VB.Label lblBalance 
      BackColor       =   &H00ECFFFE&
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
      Left            =   19095
      TabIndex        =   42
      Top             =   5265
      Width           =   1095
   End
   Begin VB.Shape Shape6 
      BorderColor     =   &H00EFB17A&
      BorderWidth     =   5
      FillColor       =   &H00EFB17A&
      Height          =   9330
      Left            =   0
      Top             =   0
      Width           =   17610
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
      Left            =   11970
      TabIndex        =   37
      Top             =   0
      Width           =   3180
   End
   Begin VB.Label lbl1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Entry &No.  :"
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
      Left            =   630
      TabIndex        =   0
      Top             =   765
      Width           =   1035
   End
   Begin VB.Label lbl1 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "From  :"
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
      Left            =   5595
      TabIndex        =   22
      Top             =   795
      Width           =   675
   End
   Begin VB.Shape Shape1 
      FillColor       =   &H00BDDFD1&
      FillStyle       =   0  'Solid
      Height          =   1830
      Left            =   10620
      Shape           =   4  'Rounded Rectangle
      Top             =   540
      Width           =   4950
   End
   Begin VB.Label Label6 
      BackColor       =   &H00EFB17A&
      Caption         =   "   Early Payment Interest List"
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
      Left            =   45
      TabIndex        =   36
      Top             =   30
      Width           =   17535
   End
   Begin VB.Shape Shape2 
      BackColor       =   &H00808080&
      BorderColor     =   &H00000080&
      FillColor       =   &H00ECFFFF&
      FillStyle       =   0  'Solid
      Height          =   1980
      Left            =   225
      Top             =   585
      Width           =   10275
   End
End
Attribute VB_Name = "frmLatePaymentIntCreditGST"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Const tblName As String = "tblVoucher"
Const IdField As String = "VnoList"
Const MaxNo As Long = 57
Dim CriteriaStr As String
Dim rstbl As Recordset, ShowFrmFlag As Boolean
Dim rsRecVsSale As Recordset
Dim varBookMark As Variant
Dim FirstTimeFlag As Boolean, mLastEntryBank As String, mLastEntryDate As Date, mLastEntryDate2 As Date
Dim FormAction As Integer
Dim frmTypeStr As String
Dim i As Long
Dim OldTdsOn As Double
Dim OldgFrmTypeStr As String
Dim MainFirm As String
Dim CriteriaStrForGenId As String
Dim BillsInFirm As String
Dim rstblBlChk As Recordset
Const OrderField As String = "Vdt,VnoList"
Private Sub cmdBtn_Click(Index As Integer)
Dim FindQStr As String
Dim X As String
Dim Cdt As String
Dim RsP As Recordset
Dim RsP2 As Recordset
Dim JvNoTDs As Long
Dim Y As String
Dim jj As Long
Dim XX As String
Dim tmpRst As Recordset
Dim wLogNo As Long
'On Error GoTo ErrorRoutine
Select Case Index
Case 0 'Previous
    If GProcRstOpen(rstbl, "Select * from " & tblName & " where VnoList = (Select max(VnoList) from " & tblName & " where VnoList <> 0  and VnoList < " & txt(35) & " and  " & CriteriaStr & "  ) and  " & CriteriaStr & "  ", "O") > 0 Then
       ReadFields
    Else
      MsgBox "This is First Record ", vbOKOnly + vbInformation
    End If
Case 1 'Next
    If GProcRstOpen(rstbl, "Select * from " & tblName & " where VnoList = (Select min(VnoList) from " & tblName & " where VnoList <> 0  and VnoList > " & txt(35) & "  and " & CriteriaStr & ") and  " & CriteriaStr & "   ", "O") > 0 Then
       ReadFields
    Else
       MsgBox "This is Last Record ", vbOKOnly + vbInformation
    End If
Case 2 'First
    If GProcRstOpen(rstbl, "Select * from " & tblName & " where VnoList = (Select min(VnoList) from " & tblName & " where VnoList <> 0 and " & CriteriaStr & ") and  " & CriteriaStr & "   ", "O") > 0 Then ReadFields
Case 3 'Last
     If GProcRstOpen(rstbl, "Select * from " & tblName & " where VnoList = (Select max(VnoList) from " & tblName & " where VnoList <> 0 and " & CriteriaStr & " )and  " & CriteriaStr & "   ", "O") > 0 Then ReadFields
Case 4 'Find
    gClsSearch.SearchMultiField "tblVoucher", "Distinct(VnoList),Vdt", Array("Vno", "Vdt"), Array(1400, 1800), CriteriaStr, , , , , True
    If gClsSearch.SearchMultiRetCol(0) <> "" Then
        If GProcRstOpen(rstbl, "Select * from " & tblName & " where VnoList = (Select max(" & gClsSearch.SearchMultiRetCol(0) & ") from " & tblName & " where VnoList <> 0 and " & CriteriaStr & " ) and  " & CriteriaStr & "   ", "O") > 0 Then ReadFields
    End If
Case 5 'List
    'GenMastList ("LAC")
Case 6 'Add
    OldTdsOn = 0
    FormAction = vbDataActionAddNew
    Call GProcActivateControls(Me, True)
    Call GProcSetButtons(Me, vbDataActionAddNew)
    Call GProcClearForm(Me, rstbl, MaxNo, True)
    txt(35).text = GProcGenerateId(tblName, "VnoList", CriteriaStrForGenId, gCn)
    txt(37) = 5
    txt(38) = Format(gLpIntRtCrNote, FStr)
    txt(42) = 0  '-- Format(gTDSRate, FStr)
    Cdt = Format(Now, "dd/MM/yyyy")
    'If CDate(Cdt) > gCYEDate Then
    '   dtpDt(2) = Format(gCYEDate, "dd/MM/yyyy")
    '   dtpDt(40) = Format(gCYEDate, "dd/MM/yyyy")
    'Else
    '   dtpDt(2) = Format(Now, "dd/MM/yyyy")
    '   dtpDt(40) = Format(Now, "dd/MM/yyyy")
    'End If
    
    If Year(mLastEntryDate) = 1899 Then
        dtpDt(2) = Format(Now, "dd/MM/yyyy")
    Else
        dtpDt(2) = mLastEntryDate
        dtpDt(40) = mLastEntryDate
        dtpDt(11) = mLastEntryDate2
    End If
    dtpDt(2) = Cdt
    dtpDt(40) = Cdt
    dtpDt(11) = Cdt
    txt(35).Enabled = False
    If FirstTimeFlag = True Then
        FirstTimeFlag = False
    Else
        dtpDt(11).SetFocus
'        txt(35).SetFocus
    End If
'    FillCombo
    SetGrid
    chkIsGSTDbNt.Value = 1
Case 7 'Modify
    'If mURecModi = True Then
    '    SearchHelp
    '    DoEvents
    If IsRelatedRecord = False Then
        FormAction = vbDataActionUpdate
        Call GProcActivateControls(Me, True)
        Call GProcSetButtons(Me, vbDataActionUpdate)
        SetControlEd
        txt(35).SetFocus
    End If
Case 8 'Delete
   ' If mURecDel = True Then
     If IsRelatedRecord = False Then
        If MsgBox("Do You Really Delete This Record?", vbQuestion + vbYesNo + vbDefaultButton1, "Confirm Delete") = vbYes Then
'           gCn.Execute "Delete from tblvoucher where " & CriteriaStr & " and vno = " & txt(0)
            '-- Delete Related Record
           If gCYear >= 2023 And gCIsLog = 1 Then
              gCn.BeginTrans
              XX = "Select max(LogNo) from tblVoucher_Log where VType = 'PN' and VFirm = '" & gCCode & "'  and VYear = " & Val(txt(3)) & " and  vnolist  = " & txt(35)
              jj = GProcRstOpen(tmpRst, XX, "R")
              If IIf(IsNull(tmpRst.Fields(0)), 0, tmpRst.Fields(0)) > 0 Then
                 wLogNo = IIf(IsNull(tmpRst.Fields(0)), 0, tmpRst.Fields(0)) + 1
              Else
                 wLogNo = 1
              End If
              gCn.Execute "Insert into tblVoucher_Log select tblVoucher.*," & wLogNo & " as LogNo,'D' as LogTp ,'" & gUserName & "' as UserName ,'" & Format(Now, "YYYY/MM/DD hh:mm:ss") & "' as LogDate from tblVoucher where VType = 'PN' and VFirm = '" & gCCode & "'  and VYear = " & Val(txt(3)) & " and  vnolist  = " & txt(35)
              gCn.CommitTrans
              tmpRst.Close
           End If
           If gCIsDepotFirm = 1 And gCDepotMainFirm > 0 Then
              gCn.BeginTrans
              XX = "Select max(LogNo) from tblVoucher_Log where VType = 'PN' and VFirm = '" & gCDepotMainFirmCompCd & "'  and VYear = " & Val(txt(3)) & " and  vnolist  = " & txt(35)
              jj = GProcRstOpen(tmpRst, XX, "R")
              If IIf(IsNull(tmpRst.Fields(0)), 0, tmpRst.Fields(0)) > 0 Then
                 wLogNo = IIf(IsNull(tmpRst.Fields(0)), 0, tmpRst.Fields(0)) + 1
              Else
                 wLogNo = 1
              End If
              gCn.Execute "Insert into tblVoucher_Log select tblVoucher.*," & wLogNo & " as LogNo,'D' as LogTp ,'" & gUserName & "' as UserName ,'" & Format(Now, "YYYY/MM/DD hh:mm:ss") & "' as LogDate from tblVoucher where VType = 'PN' and VFirm = '" & gCDepotMainFirmCompCd & "'  and VYear = " & Val(txt(3)) & " and  vnolist  = " & txt(35)
              gCn.CommitTrans
              tmpRst.Close
           End If
            DeleteRelatedRecord
            cmdBtn_Click (0)
            If GProcRstOpen(rstbl, "Select * from " & tblName & " where VnoList = (Select max(VnoList) from " & tblName & " where VnoList <> 0 and " & CriteriaStr & " )and  " & CriteriaStr & "   ", "O") > 0 Then
               ReadFields
            Else
                SetGrid
                Call GProcClearForm(Me, rstbl, MaxNo, True)
            End If
            If rstbl.RecordCount > 1 Then
                rstbl.Requery
                cmdBtn_Click (1)
            ElseIf rstbl.RecordCount = 1 Then
                Call GProcActivateControls(Me, False)
                Call GProcSetButtons(Me, vbDataActionCancel)
                Form_Load
            End If
        End If
        DoEvents
     End If
   ' End If
Case 9 'Print
    Frame1.Visible = True
    Frame2.Visible = True
    OptCopy(0).Visible = True
    OptCopy(1).Visible = True
    OptCopy(0).Value = True
    OptCopy(0).Enabled = True
    OptCopy(1).Enabled = True
    DoEvents
    DoEvents
    txtno(0).Enabled = True
    txtno(1).Enabled = True
    OptType(0).Enabled = True
    OptType(1).Enabled = True
    OptType(0).Value = True
    txtno(0) = txt(35)
    txtno(1) = txt(35)
    txtno(0).SetFocus
Case 10 'Exit
    DoEvents
    Unload Me
Case 13 'Save
      OldTdsOn = 0
       If txtGrid.Visible = True Then
'       If txtGrid.Enabled = True Then
            msGrid.text = txtGrid
            DoEvents
            txtGrid = ""
        End If
        txtGrid.Visible = False
    If ValidateData = True Then
       If txtGrid.Visible = True Then
'       If txtGrid.Enabled = True Then
            msGrid.text = txtGrid
            DoEvents
            txtGrid = ""
        End If
        txtGrid.Visible = False
        FillTxtFromTemp
        SaveRelatedRecord
        
        mLastEntryDate = dtpDt(2)
        mLastEntryDate2 = dtpDt(11)
        
        DoEvents
        DoEvents
        txtGrid = ""
        rsRecVsSale.Requery
        rstbl.Requery
        
        If GProcRstOpen(rstbl, "Select * from " & tblName & " where VnoList = " & txt(35) & " and  " & CriteriaStr & "   ", "O") > 0 Then ReadFields
'        FillTempFromTxt
'        ReadFields

        FormAction = vbDataActionClose
        Call GProcActivateControls(Me, False)
        Call GProcSetButtons(Me, True)
        msGrid.Col = 0
        msGrid.Row = 1
        txtGrid.Move msGrid.CellLeft + msGrid.Left, msGrid.CellTop + msGrid.Top, msGrid.CellWidth, msGrid.CellHeight
        lblBalance.Caption = ""
'        If FormAction = vbDataActionAddNew Then
            If cmdBtn(6).Visible = True Then cmdBtn(6).SetFocus
 '       Else
  '          If cmdBtn(7).Visible = True Then cmdBtn(7).SetFocus
   '     End If
    End If
Case 14 'Cancel
    FormAction = vbDataActionClose
    txtGrid.Visible = False
    If GProcRstOpen(rstbl, "Select * from " & tblName & " where VnoList = (Select max(VnoList) from " & tblName & " where VnoList <> 0 and " & CriteriaStr & " ) and  " & CriteriaStr & "   ", "O") > 0 Then ReadFields
    
    If rstbl.RecordCount <= 0 Then
        Call GProcClearForm(Me, rstbl, MaxNo, True)
        Call GProcActivateControls(Me, False)
        Call GProcSetButtons(Me, vbDataActionCancel)
    Else
        If FormAction = vbDataActionAddNew Then
            rstbl.MoveLast
        Else
            rstbl.CancelUpdate
        End If
        txtGrid = ""
        ReadFields
        Call GProcActivateControls(Me, False)
        Call GProcSetButtons(Me, True)
    End If
    FormAction = vbDataActionClose
    msGrid.Col = 0
    If msGrid.Row >= 1 Then msGrid.Row = 1
    txtGrid.Move msGrid.CellLeft + msGrid.Left, msGrid.CellTop + msGrid.Top, msGrid.CellWidth, msGrid.CellHeight
    lblBalance.Caption = ""
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

Private Sub dtpDt_GotFocus(Index As Integer)
GProcSelectBox Me.ActiveControl
End Sub

Private Sub dtpDt_LostFocus(Index As Integer)
SaveBtnEd
End Sub

Private Sub txtGrid_DblClick()
If msGrid.Col = 13 Then
   If msGrid.TextMatrix(msGrid.Row, 4) = msGrid.TextMatrix(msGrid.Row - 1, 4) Then
      If txtGrid.text = "" Then
         txtGrid.text = msGrid.TextMatrix(msGrid.Row - 1, 13)  '  IIf(msGrid.TextMatrix(msGrid.Row, 13) = "", msGrid.TextMatrix(msGrid.Row - 1, 0), msGrid.TextMatrix(msGrid.Row - 1, 13))
      Else
         txtGrid.text = ""
      End If
   Else
      If txtGrid.text = "" Then
          txtGrid.text = msGrid.TextMatrix(msGrid.Row, 0)
      Else
         txtGrid.text = ""
      End If
   End If
End If
If msGrid.Col = 12 Then
   If txtGrid.text = "Y" Then
      txtGrid.text = "N"
   Else
      txtGrid.text = "Y"
   End If
End If
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
Dim wMainFirmNm As String
Select Case Index
Case 0, 1:
    '------ Validate
    If IsNumeric(txtno(0)) = False Then txtno(0) = 0
    If IsNumeric(txtno(1)) = False Then txtno(1) = txtno(0)
    If CLng(txtno(1)) < CLng(txtno(0)) Then txtno(1) = txtno(0)
    '-------
    If Index = 0 Then GProcCrystalRptPreparation frmMain.cryReport, "Printer"
    If Index = 1 Then GProcCrystalRptPreparation frmMain.cryReport, "Window"
    frmMain.cryReport.WindowTitle = "Receipt Voucher"
    Vno = CLng(txtno(0))
    Do While Vno <= txtno(1)
       If gCYear >= 2021 Then '--ItHsn2,ItHsn
         X = "select tblVoucher.Vno,tblVoucher.VFirm,tblMastCompany.Cname,VBillNo,VBillType, VBillDate,VBillAmt,tblMastAccount.AcOurCode,tblMastAccount.AcName,tblMastAccount.AcAdd1,tblMastAccount.AcAdd2,tblMastAccount.AcAdd3,tblMastAccount.AcGSTIN,tblMastAccount_Br.AcOurCode,tblMastAccount_Br.AcName, VAmt,DueOnDt,tblVoucher.VRefDate,LateDays,InterestCredit,JvNoDisc,Grace,IntRt,DbNtMonthly,DbNtInBillNo,CgstRt,SgstRt,IgstRt,CgstAmt,SgstAmt,IgstAmt,ItHsn2,tblVoucher.CrDrNoteNo,'" & gUserName & "'," _
             & " tblVoucher.Vno,tblVoucher.Vdt,tblMastAccount.AcContNo  from tblVoucher,tblMastAccount,tblMastAccount tblMastAccount_Br,tblSale,tblSaleSub,tblMastItem,tblMastCompany where " _
             & " tblVoucher.VType='" & frmTypeStr & "' and VnoList=" & Vno & " and tblVoucher.VFirm in (" & BillsInFirm & ") and  tblVoucher.VYear=" & gCYear & " and VCrAcCode = tblMastAccount.AcCode and  tblSale.SlBroker = tblMastAccount_Br.AcCode and tblVoucher.VFirm = dbo.tblMastCompany.CCode  and " _
             & " VBillVno = tblSale.Vno and VBillType = tblSale.Vtype and VBillVYear = tblSale.Vyear and tblVoucher.VFirm = tblSale.Vfirm  and tblSale.Vno = tblSaleSub.vno  and tblSale.Vtype = tblSaleSub.vtype and tblSale.Vyear = tblSaleSub.vyear and tblsale.VFirm = tblSalesub.Vfirm and tblsalesub.SlSubItCode = tblMastItem.ItCode and tblsalesub.SlSubItSrNo = 1 " _
             & " order by tblMastAccount.AcName,tblVoucher.Vno,VCtrNo"
       Else
         X = "select tblVoucher.Vno,tblVoucher.VFirm,tblMastCompany.Cname,VBillNo,VBillType, VBillDate,VBillAmt,tblMastAccount.AcOurCode,tblMastAccount.AcName,tblMastAccount.AcAdd1,tblMastAccount.AcAdd2,tblMastAccount.AcAdd3,tblMastAccount.AcGSTIN,tblMastAccount_Br.AcOurCode,tblMastAccount_Br.AcName, VAmt,DueOnDt,tblVoucher.VRefDate,LateDays,InterestCredit,JvNoDisc,Grace,IntRt,DbNtMonthly,DbNtInBillNo,CgstRt,SgstRt,IgstRt,CgstAmt,SgstAmt,IgstAmt,ItHsn,tblVoucher.CrDrNoteNo,'" & gUserName & "'," _
             & " tblVoucher.Vno,tblVoucher.Vdt,tblMastAccount.AcContNo  from tblVoucher,tblMastAccount,tblMastAccount tblMastAccount_Br,tblSale,tblSaleSub,tblMastItem,tblMastCompany where " _
             & " tblVoucher.VType='" & frmTypeStr & "' and VnoList=" & Vno & " and tblVoucher.VFirm in (" & BillsInFirm & ") and  tblVoucher.VYear=" & gCYear & " and VCrAcCode = tblMastAccount.AcCode and  tblSale.SlBroker = tblMastAccount_Br.AcCode and tblVoucher.VFirm = dbo.tblMastCompany.CCode  and " _
             & " VBillVno = tblSale.Vno and VBillType = tblSale.Vtype and VBillVYear = tblSale.Vyear and tblVoucher.VFirm = tblSale.Vfirm  and tblSale.Vno = tblSaleSub.vno  and tblSale.Vtype = tblSaleSub.vtype and tblSale.Vyear = tblSaleSub.vyear and tblsale.VFirm = tblSalesub.Vfirm and tblsalesub.SlSubItCode = tblMastItem.ItCode and tblsalesub.SlSubItSrNo = 1 " _
             & " order by tblMastAccount.AcName,tblVoucher.Vno,VCtrNo"
       End If
               
               
            gCn.Execute "Delete from TmpGentbl where  UserName='" & gUserName & "'"
     
            gCn.Execute " insert into TmpGentbl (Vno,VFirm,VFirmName,VBillNo,VBillType,VBillDt,Amt1,PartyCode,PartyName,Nar1,Nar4,Nar6,Nar5,BrokerCode,BrokerName,Amt2,Dt1,Vdt,Amt3,Amt4,Amt5,Amt6,amt7,amt8,Nar2,Amt9,Amt10,Amt11,Amt12,Amt13,Amt14,Nar3,Nar7,UserName,Amt15,Dt2,Nar8) " & X
'            X1 = "{TmpGentbl.UserName}='" & gUserName & "' and {@IntAmt} > 0  "
            With frmMain.cryReport
               If gCDepotMainFirm > 0 Then
                  wMainFirmNm = GProcGetColumnValue("tblMastCompany", "CCode", gCDepotMainFirmCompCd, "S", "CName", "S")
                  .Formulas(0) = "FirmNm = '" & UCase(wMainFirmNm) & "'"
              End If
              .SelectionFormula = "{TmpGentbl.UserName}='" & gUserName & "'"
              .Formulas(5) = "wIntRt =  '" & txt(1) & "'"
              .Formulas(6) = "wGrase =  '" & txt(2) & "'"
              If OptType(0).Value = True Then
                 .ReportFileName = gReportPath & "rptIntCrNtListGST.rpt"
              Else
                 .Formulas(7) = "wGSTNo = '" & gCGStin & "'"
                 If OptCopy(0).Value = True Then
                    .Formulas(8) = "wCopy =  'Customer Copy'"
                 Else
                    .Formulas(8) = "wCopy =  'Office Copy'"
                 End If
                 .ReportFileName = gReportPath & "rptIntCrNtGSTPlain.rpt"
              End If
              .Action = 1
              .PageZoom (120)

            End With
 '       End If
        Vno = Vno + 1
    Loop
Case 2:
    DoEvents
    Frame1.Visible = False
Case 3:
    '------ Validate
    If IsNumeric(txtno(0)) = False Then txtno(0) = 0
    If IsNumeric(txtno(1)) = False Then txtno(1) = txtno(0)
    If CLng(txtno(1)) < CLng(txtno(0)) Then txtno(1) = txtno(0)
    '-------
    If Index = 3 Then GProcCrystalRptPreparation frmMain.cryReport, "Printer"
    If Index = 3 Then GProcCrystalRptPreparation frmMain.cryReport, "Window"
    frmMain.cryReport.WindowTitle = "Short Payment Voucher"
    Vno = CLng(txtno(0))
    Do While Vno <= txtno(1)
        X = "Select * From qryshortrecvoucher where  RecVNo=" & Vno & " and RecVType='" & frmTypeStr & "' and RecVYear=" & gCYear
        If GProcRstOpen(Rs1, X, "R") > 0 Then
            With frmMain.cryReport
                .SelectionFormula = "{qryshortrecvoucher.RecVno} = " & Rs1!RecVno & "  and {qryshortrecvoucher.RecVType}='" & Rs1!RecVType & "' and {qryshortrecvoucher.RecVYear}=" & Rs1!RecVYear
                .ReportFileName = gReportPath & "rptVouReceiptShort.rpt"
                .Action = 1
            End With
        End If
        Vno = Vno + 1
    Loop
End Select
End Sub
Private Sub cmdIntDbNt_Click()
'   DeleteRelatedRecord
   ApplyIntDbNt
End Sub
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case 13:
            If Me.ActiveControl.Name <> "txtGrid" Then Sendkeys "{TAB}"
        Case 27:
'            If (cmdBtn(13).Enabled = False) Then
'                Unload Me
'            End If
    End Select
End Sub
Private Sub Form_Load()
    Dim X As String
    Dim Y As String
    Me.Top = 300
    Me.Left = Screen.Width / 2 - Me.Width / 2
    frmTypeStr = gfrmTypeStr
'    X = GProcGetColumnValue("tblMastCompany", "CCode", gCCode, "S", "CDepotMainFirmCompCd", "S")
'    If X = "" Then
'       MainFirm = gCCode
'    Else
'       MainFirm = X
'    End If
    
'    X = "Select CCode from tblMastCompany where CDepotMainFirmCompCd = '" & gCCode & "'"
'    i = GProcRstOpen(rstblBlChk, X, "R")
    BillsInFirm = "'" & gCCode & "'"
'    If i > 0 Then
'       rstblBlChk.MoveFirst
'       Do While Not rstblBlChk.EOF
'          BillsInFirm = BillsInFirm + "," + "'" & rstblBlChk.Fields(0) & "'"
'          rstblBlChk.MoveNext
'       Loop
'    End If
    CriteriaStr = " VType = '" & frmTypeStr & "' and  VnoList <> 0  and InterestCredit > 0 and VFirm in (" & BillsInFirm & ") and VYear=" & gCYear
    CriteriaStrForGenId = " VType = '" & frmTypeStr & "' and  VnoList <> 0 and VFirm in (" & BillsInFirm & ") and VYear=" & gCYear
    Y = "Select * from " & tblName & " where " & CriteriaStr & " "
    If Trim(OrderField) <> "" Then Y = Y & " order by " & OrderField
    GProcRstOpen rstbl, Y, "O"
    FirstTimeFlag = True
    '---Grid
    SetGrid
    '---- Button Visiblity Set As Per User Previlage
    GProcSetButtonVisiblity Me
    cmdBtn_Click (6)
    chkIsGSTDbNt.Value = 1
End Sub
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
'If gFormAction = vbDataActionUpdate Then
'    MsgBox "You are about to modify the record. Save the record and then exit.", vbInformation + vbOKOnly, Me.Caption
'    Cancel = True
'    Exit Sub
'End If
rstbl.Close
End Sub
Private Sub txt_GotFocus(Index As Integer)
GProcSelectBox Me.ActiveControl
End Sub
'Private Sub txt_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
'Dim OldFrmType As String
'Dim LeftPos As Long
'Dim TopPos As Long
'If KeyCode = vbKeyF3 Then '--F3 Create
'    Select Case Index
'        Case 13, 14, 15, 16 '--Narration
'            OldFrmType = gfrmTypeStr
'            gfrmTypeStr = "N"
'            GProcShowForm frmMastNarration, frmMain.mnuMstfrm(5), , gfrmTypeStr
'            gfrmTypeStr = OldFrmType
'    End Select
'End If
'If KeyCode = vbKeyF2 Then '-- Select
'    TopPos = txt(Index).Top + Me.Top + 650
'    LeftPos = txt(Index).Left + Me.Left
'    Select Case Index
'        Case 13, 14, 15, 16 '--Narration
'            gClsSearch.SearchMultiField "tblMastNarration", "Narration", Array("Narration"), Array(txt(Index).Width), "NarrType='N'", "", "Narration", LeftPos, TopPos, False
'            If gClsSearch.SearchMultiRetCol(0) <> "" Then txt(Index) = gClsSearch.SearchMultiRetCol(0)
'    End Select
'End If
'End Sub
Private Sub txt_LostFocus(Index As Integer)
GProcCheckForNumber Me, rstbl, MaxNo
If FormAction = vbDataActionAddNew And Index = 42 Then
   PreperReport
End If
If FormAction = vbDataActionUpdate And Index = 35 Then
   msGrid.Col = 9
   msGrid.SetFocus
   msGrid_EnterCell
End If


SaveBtnEd
End Sub
Private Sub txt_KeyPress(Index As Integer, KeyAscii As Integer)
    KeyAscii = GProcValidateKey(rstbl, Index, KeyAscii, txt(Index).text)
End Sub
Private Sub SaveBtnEd()
'--Vno,Vdt,Amount,Cr A/c,Dr A/c
If txt(35) = "" Or GProcIsDateValid(dtpDt(2)) = False Or GProcIsDateValid(dtpDt(11)) = False Or GProcIsDateValid(dtpDt(40)) = False Then
    cmdBtn(13).Enabled = False
    cmdBtn(13).BackColor = &H80000004
Else
    cmdBtn(13).Enabled = True
    cmdBtn(13).BackColor = &HFFFF80
End If
End Sub
Private Sub FillTxtFromTemp()
txt(1) = frmTypeStr
txt(2) = IIf(IsDate(dtpDt(2)), dtpDt(2), "") '--Vdt
txt(40) = IIf(IsDate(dtpDt(40)), dtpDt(40), "") '--Vdt
txt(11) = IIf(IsDate(dtpDt(11)), dtpDt(11), "") '--Vdt
txt(41) = GProcGetColumnValue("tblMastAccount", "AcName", txttemp(41), "S", "AcCode", "N")
txt(56) = chkIsGSTDbNt.Value
txt(57) = IIf(IsDate(dtpDt(2)), dtpDt(2), "") '--Vdt
txt(3) = gCYear
txt(4) = gCCode
ShowAccountBalance
End Sub
Private Sub FillTempFromTxt()
If IsDate(txt(57)) Then
   dtpDt(2).Value = IIf(IsDate(txt(57)), CDate(txt(57)), Format(Now, "dd/MM/yyyy"))  '--Int From Vdt
Else
   dtpDt(2) = Format(Now, "dd/MM/yyyy")
End If



'dtpDt(2) = IIf(IsDate(txt(2)), CDate(txt(2)), Format(Now, "dd/MM/yyyy")) '--Vdt
If IsDate(txt(40)) Then
   dtpDt(40).Value = IIf(IsDate(txt(40)), CDate(txt(40)), Format(Now, "dd/MM/yyyy"))  '--Int From Vdt
Else
   dtpDt(40) = Format(Now, "dd/MM/yyyy")
End If
If IsDate(txt(11)) Then
   dtpDt(11).Value = IIf(IsDate(txt(11)), CDate(txt(11)), Format(Now, "dd/MM/yyyy"))  '--Int From Vdt
Else
   dtpDt(11) = Format(Now, "dd/MM/yyyy")
End If


txttemp(41) = GProcGetColumnValue("tblMastAccount", "AcCode", txt(41), "N", "AcName", "S")
If Val(txt(31)) > 0 Then
   lblApplyJV.Visible = False
   cmdIntDbNt.Visible = False
'   cmdBtn(7).Visible = False
'   cmdBtn(8).Visible = False
Else
   lblApplyJV.Visible = True
   cmdIntDbNt.Visible = True
   cmdBtn(7).Visible = True
'   cmdBtn(8).Visible = True
End If
chkIsGSTDbNt.Value = Val(txt(56))
'-- Fill Related Record Rec Vs Sale
FillRelatedRecord
'ShowAccountBalance
End Sub
Private Function ValidateGrid() As Boolean
'--- Pending
Dim TotAmt As Double
Dim i As Long
Dim RCount As Long
Dim j As Long
Dim ChkTds As Double
Dim BilChk As Boolean
Dim PartyChk As Boolean
'If Trim(txtGrid.text) <> "" Then
'   msGrid.text = txtGrid.text
'End If
'txtGrid.Visible = False
'txtGrid.text = ""
CheckForNumberGrid
With msGrid
     For i = 1 To .rows - 1
         If Val(.TextMatrix(i, 11)) > 0 And Val(.TextMatrix(i, 9)) = 0 Then
                MsgBox "Please Check Interest Amount. ", vbOKOnly, Me.Caption
                msGrid.SetFocus
                Exit Function
         End If
         If Val(.TextMatrix(i, 11)) > 0 Then
            ChkTds = GProcMakeRounding((Val(.TextMatrix(i, 10)) * Val(txt(42))) / 100)
            If Abs(Val(.TextMatrix(i, 11)) - ChkTds) > 1 Then
                MsgBox "Please Check TDS Amount. ", vbOKOnly, Me.Caption
                .Col = 10
                .SetFocus
                Exit Function
            End If
         End If
         If UCase(.TextMatrix(i, 12)) = "Y" Or UCase(.TextMatrix(i, 12)) = "N" Then
         Else
                MsgBox "Please Check YN ", vbOKOnly, Me.Caption
                .Col = 12
                .SetFocus
                Exit Function
         End If
         
         
         
         BilChk = False
         PartyChk = False
         If Trim(.TextMatrix(i, 13)) <> "" Then
             For j = 1 To .rows - 1
                 If .TextMatrix(j, 0) = .TextMatrix(i, 13) Then
                    BilChk = True
                    If .TextMatrix(j, 4) = .TextMatrix(i, 4) Then
                       PartyChk = True
                    End If
                 End If
                 If BilChk = True And PartyChk = True Then
                    Exit For
                 End If
             Next j
             If BilChk = False Then
                MsgBox "Please Check Bill No. ", vbOKOnly, Me.Caption
               .Row = i
               .Col = 13
               .SetFocus
               Exit Function
             End If
             If PartyChk = False Then
                MsgBox "Please Check Bill No. And Party Name ", vbOKOnly, Me.Caption
               .Row = i
               .Col = 13
               .SetFocus
               Exit Function
             End If
         End If
         
         
     Next i
End With



'With msGrid
'    '-- Duplicate invoice selection
'    For i = 1 To .rows - 1
'        If Trim(.TextMatrix(i, 1)) <> "" Then
'           BillDet = BillDet + .TextMatrix(i, 1) + "-" + CStr(Val(.TextMatrix(i, 4)) + Val(.TextMatrix(i, 6)) + Val(.TextMatrix(i, 9))) + ":"
'        End If
'
'        '-- Int A/c
'           If CDbl(.TextMatrix(i, 9)) > 0 Then
'                If gLatePayIntAcCode = GProcGetColumnValue("tblmastAccount", "AcName", "-", "S", "AcCode", "N") Then
'                    IntAcCode = 0
'                Else
'                    IntAcCode = gLatePayIntAcCode
'                End If
'                If IntAcCode = 0 Then
'                    MsgBox "Check Late Pay Interest Account.", vbInformation + vbOKOnly, Me.Caption
'                    GProcShowForm frmSetting, frmMain.mnuMstSetting(0)
'                    Exit Function
'                End If
'           End If
'        '---
'        '-- TDS Paid A/c ( For Less TDS Amt JV )
'           If CDbl(.TextMatrix(i, 13)) > 0 Then
'                If gTDSPaidAcCode = GProcGetColumnValue("tblmastAccount", "AcName", "-", "S", "AcCode", "N") Then
'                    TDSPaidAcCode = 0
'                Else
'                    TDSPaidAcCode = gTDSPaidAcCode
'                End If
'                If TDSPaidAcCode = 0 Then
'                    MsgBox "Check TDS Paid Account.", vbInformation + vbOKOnly, Me.Caption
'                    GProcShowForm frmSetting, frmMain.mnuMstSetting(0)
'                    Exit Function
'                End If
'           End If
'        '---
'        InvNo = .TextMatrix(i, 11)
'        InVtype = .TextMatrix(i, 0)
'        InvYear = .TextMatrix(i, 10)
'        For j = 1 To .rows - 1
'            If j <> i And InvNo = .TextMatrix(j, 11) And InVtype = .TextMatrix(j, 0) And InvYear = .TextMatrix(j, 10) And InvNo <> 0 Then
'                MsgBox "Duplicate Invoice Selection.", vbOKOnly, Me.Caption
'                msGrid.SetFocus
'                Exit Function
'            End If
'        Next j
'    Next i
'    If BillDet <> "" Then
'        BillDet = Left(Trim(BillDet) + Space(120), 120)
'        txt(8) = Left(BillDet, 40)
'        txt(9) = Mid(BillDet, 41, 40)
'        txt(10) = Mid(BillDet, 81, 40)
'    End If
'    '-- Adjusted Amt
'    If .rows > 1 Then
'       TotAmt = 0
'       For i = 1 To .rows - 1
'           '-- Adj amt > Bill Amt
'           If CDbl(.TextMatrix(i, 4)) - CDbl(.TextMatrix(i, 5)) > CDbl(.TextMatrix(i, 3)) Then
'                MsgBox "Adjusted Amount is more than Balance Bill Amount. ", vbOKOnly, Me.Caption
'                msGrid.SetFocus
'                Exit Function
'           End If
'            '-- Adj DB amt > Db Amt
''           If CDbl(.TextMatrix(i, 6)) > CDbl(.TextMatrix(i, 5)) Then
''                MsgBox "Adjusted DbNt Amount is more than Balance DbNt Amount. ", vbOKOnly, Me.Caption
''                msGrid.SetFocus
''                Exit Function
''           End If
'           TotAmt = TotAmt + Val(.TextMatrix(i, 4)) - Val(.TextMatrix(i, 5)) + Val(.TextMatrix(i, 9)) - Val(.TextMatrix(i, 13))
'        Next i
'        If (TotAmt > 0) And (TotAmt <> Val(txt(6))) Then
'           MsgBox "Total of Adjusted Amount " & TotAmt & " & Received Amount not tally.", vbOKOnly, Me.Caption
'     '      msGrid.SetFocus
'     '      Exit Function
'        End If
'    End If
'End With
ValidateGrid = True
End Function
Private Function ValidateData() As Boolean
Dim X As String
Dim GName As String
Dim BnkTot As Double
Dim RowIndex As Integer
Dim rstAddCheck As Recordset
GProcCheckForNumber Me, rstbl, MaxNo
If ValidateGrid = False Then
        'MSGrid.SetFocus
    Exit Function
End If
If GProcIsDateValid(dtpDt(11)) = False Then
   MsgBox "Check Voucher Date.", vbInformation + vbOKOnly, Me.Caption
   dtpDt(11).SetFocus
   Exit Function
End If
RowIndex = 1
'''If txttemp(6).text <> "" Then
'''    With msGrid
'''        Do While RowIndex < .rows
'''            '-- sub amt
'''            If Val(.TextMatrix(RowIndex, 6)) > 0 Then
'''                BnkTot = BnkTot + Val(.TextMatrix(RowIndex, 6))
'''            End If
'''            RowIndex = RowIndex + 1
'''        Loop
'''    End With
'''    If BnkTot > 0 Then
'''        MsgBox "For Bill Detail Party Name Should Be Blank.", vbInformation + vbOKOnly, Me.Caption
'''        txttemp(6).SetFocus
'''        Exit Function
'''    End If
'''End If
'''
'''
'''
''''---- Required Data
'''If ValidateGrid = False Then Exit Function
''''--- Vno
'''If CLng(txt(0)) = 0 Then
'''    MsgBox "Check Voucher No.", vbInformation + vbOKOnly, Me.Caption
'''    txt(0).SetFocus
'''    Exit Function
'''End If
''''--- Vdt
'''If GProcIsDateValid(mskDt(2)) = False Then
'''    MsgBox "Check Voucher Date.", vbInformation + vbOKOnly, Me.Caption
'''    mskDt(2).SetFocus
'''    Exit Function
'''End If
''''---- party
''''If txttemp(6) = "" Then
''''    MsgBox "Check Account.", vbInformation + vbOKOnly, Me.Caption
''''    txttemp(6).SetFocus
''''    Exit Function
''''End If
''''--Bank A/c
''''If GProcGetColumnValue("tblMastAccount", "AcName", txttemp(6), "S", "AcCode", "N") = 0 Then
''''   GName = GProcGetColumnValue("TblMastGroup", "AgCode", 90017, "N", "AgName", "S")
''''   GProcShowForm frmMastAccount, frmMain.mnuMstfrm(0), Array(frmMastAccount.txt(1), frmMastAccount.txttemp(2)), Array(txttemp(6), GName)
''''   Exit Function
''''End If
''''--- Cr A/c
'''If txttemp(7) = "" Then
'''    MsgBox "Check Account.", vbInformation + vbOKOnly, Me.Caption
'''    txttemp(7).SetFocus
'''    Exit Function
'''End If
''''--Create Cr A/c
'''If GProcGetColumnValue("tblMastAccount", "AcName", txttemp(7), "S", "AcCode", "N") = 0 Then
'''   GName = GProcGetColumnValue("TblMastGroup", "AgCode", 90019, "N", "AgName", "S")
'''   GProcShowForm frmMastAccount, frmMain.mnuMstfrm(0), Array(frmMastAccount.txt(1), frmMastAccount.txttemp(2)), Array(txttemp(7), GName)
'''   Exit Function
'''End If
''''--- Cr A/c ,Dr A/c
'''If txttemp(6) = txttemp(7) Then
'''    MsgBox "Check Account.Debit and Credit Accounts are Same.", vbInformation + vbOKOnly, Me.Caption
'''    txttemp(6).SetFocus
'''    Exit Function
'''End If
'''
''''--- Pending
'''
''''---- Amount
''''If CDbl(txt(6)) <= 0 Then
''''    MsgBox "Check Voucher Amount.", vbInformation + vbOKOnly, Me.Caption
''''    Exit Function
''''End If
''''----Duplicate Records
''''----Addition Case
'''
'''
''''If FormAction = vbDataActionAddNew Then
''''    X = "Select vno From " & tblName & " Where vno=" & CLng(txt(0)) & " and Vtype='" & frmTypeStr & "' and VYear=" & gCYear
''''    i = GProcRstOpen(rstAddCheck, X, "R")
''''    If i > 0 Then
''''        MsgBox "Record is Available of Voucher No. " & CLng(txt(0)), vbCritical + vbOKOnly, "Duplicate Record"
''''        txt(0).SetFocus
''''        Exit Function
''''    End If
''''    rstAddCheck.Close
''''End If
ValidateData = True
End Function
Private Sub ReadFields()
GProcShowRecord Me, rstbl, MaxNo
FillTempFromTxt
End Sub
'--- Related Record
Private Function IsRelatedRecord() As Boolean
'-- Audited
'If GProcISAudited(txt(0), txt(3), gCYear) = True Then
'    MsgBox "Record is Audited.You can not Modify/Delete this Record.", vbCritical + vbOKOnly, App.Title
'    IsRelatedRecord = True
'    Exit Function
'End If
'--- Related Record Receipt
'If rstbl.EOF And rstbl.BOF Then
'Else
'    If rstbl.Fields("VIsRefEntType") = 1 Then
'        MsgBox "Related Record is Present.You can not Modify/Delete this Record.", vbCritical + vbOKOnly, App.Title
'        IsRelatedRecord = True
'        Exit Function
'    End If
'End If
End Function
Private Sub ShowAccountBalance()
Dim DrAcBal  As Double
Dim CrAcBal As Double
'If IsDate(mskDt(2)) Then
'    DrAcBal = GProcGetAccountBalance(txttemp(6), mskDt(2))
'    CrAcBal = GProcGetAccountBalance(txttemp(7), mskDt(2))
'    txtDrBalance = IIf(DrAcBal > 0, DrAcBal & " Cr", Abs(DrAcBal) & " Dr")
'    txtCrBalance = IIf(CrAcBal > 0, CrAcBal & " Cr", Abs(CrAcBal) & " Dr")
'End If
End Sub
'---------------- Grid
Private Sub SetGrid()
Dim X As String, x1Rec As Double
Set rsRecVsSale = New Recordset

'0 - VBillNo
'1 - VBillType
'2 - VBillDate
'3 - Bill Amt
'4 - Pty Name
'5 - Brok NM
'6 - Amt Recd
'7 - Rec Dt
'8 - OD
'9 - Int
'10 - TDS On Amt
'11 - Tds
'12 - DbNtInt (Y/N)
'13 - DbNtIn Bill No
'14 - DbNtMonthly
'15 - Due On
'16 - Vno
'17 - VctrNo
'18 - Grase
'19 - IntRt
'20 - CRt
'21 - SRt
'22 - IRt
'23 - JvNoInt
'24 - JvNoLessTDS
'25 - HSN Code
'26 - VbillFirm
'27 - CrDrNoteNo
'28 -VnolistVyear
'29 -Vyear

If gCYear >= 2021 Then '--ItHsn2,ItHsn
    X = "select VBillNo,VBillType, VBillDate,VBillAmt,iif(tblMastAccount.AcIsIntDbNtMonthly = 1,' * '+tblMastAccount.AcName,tblMastAccount.AcName),tblMastAccount_Br.AcName, VAmt,tblVoucher.Vdt,LateDays,InterestCredit,TDSOn,LessTDS,DBNtInt,DbNtInBillNo,DbNtMonthly,DueOnDt,tblVoucher.Vno,VCtrNo,Grace,IntRt,CgstRt,SgstRt,IgstRt,JvNoDisc,JvNoLessTDS,tblMastItem.ItHsn2,tblVoucher.VBillFirm,tblVoucher.CrDrNoteNo,tblVoucher.VNoListVYear,tblVoucher.Vyear" _
    & " from tblVoucher,tblMastAccount,tblMastAccount tblMastAccount_Br,tblSale,tblSaleSub,tblMastItem where " _
    & " tblVoucher.VType='" & frmTypeStr & "' and VnoList=" & txt(35) & " and tblVoucher.VFirm in (" & BillsInFirm & ") and  tblVoucher.VYear=" & gCYear & " and VCrAcCode = tblMastAccount.AcCode and  tblSale.SlBroker = tblMastAccount_Br.AcCode  and " _
    & " VBillVno = tblSale.Vno and VBillType = tblSale.Vtype and VBillVYear = tblSale.Vyear and (case when isnull(tblvoucher.vBillfirm,'0')='0' then tblvoucher.vfirm else tblvoucher.vBillfirm end)=tblsale.vfirm and tblSale.vno = tblSaleSub.Vno and tblSale.vtype = tblSaleSub.Vtype and tblSale.vyear = tblSaleSub.Vyear and tblSale.vfirm = tblSaleSub.Vfirm and  tblsaleSub.SlSubItSrNo = 1 and " _
    & " tblSaleSub.SlSubItCode = tblMastItem.ItCode"  ' order by tblMastAccount.AcIsIntDbNtMonthly,tblMastAccount.AcName,tblMastItem.ItHsn,tblsale.vdt,tblsale.Vno"
    X = X + " order by tblvoucher.vfirm,tblMastAccount.AcIsIntDbNtMonthly,tblMastAccount.AcOurCode,tblVoucher.Vdt,tblMastItem.ItHsn2,tblVoucher.vno"
Else
    X = "select VBillNo,VBillType, VBillDate,VBillAmt,iif(tblMastAccount.AcIsIntDbNtMonthly = 1,' * '+tblMastAccount.AcName,tblMastAccount.AcName),tblMastAccount_Br.AcName, VAmt,tblVoucher.Vdt,LateDays,InterestCredit,TDSOn,LessTDS,DBNtInt,DbNtInBillNo,DbNtMonthly,DueOnDt,tblVoucher.Vno,VCtrNo,Grace,IntRt,CgstRt,SgstRt,IgstRt,JvNoDisc,JvNoLessTDS,tblMastItem.ItHsn,tblVoucher.VBillFirm,tblVoucher.CrDrNoteNo,tblVoucher.VNoListVYear,tblVoucher.Vyear" _
    & " from tblVoucher,tblMastAccount,tblMastAccount tblMastAccount_Br,tblSale,tblSaleSub,tblMastItem where " _
    & " tblVoucher.VType='" & frmTypeStr & "' and VnoList=" & txt(35) & " and tblVoucher.VFirm in (" & BillsInFirm & ") and  tblVoucher.VYear=" & gCYear & " and VCrAcCode = tblMastAccount.AcCode and  tblSale.SlBroker = tblMastAccount_Br.AcCode  and " _
    & " VBillVno = tblSale.Vno and VBillType = tblSale.Vtype and VBillVYear = tblSale.Vyear and (case when isnull(tblvoucher.vBillfirm,'0')='0' then tblvoucher.vfirm else tblvoucher.vBillfirm end)=tblsale.vfirm and tblSale.vno = tblSaleSub.Vno and tblSale.vtype = tblSaleSub.Vtype and tblSale.vyear = tblSaleSub.Vyear and tblSale.vfirm = tblSaleSub.Vfirm and  tblsaleSub.SlSubItSrNo = 1 and " _
    & " tblSaleSub.SlSubItCode = tblMastItem.ItCode"  ' order by tblMastAccount.AcIsIntDbNtMonthly,tblMastAccount.AcName,tblMastItem.ItHsn,tblsale.vdt,tblsale.Vno"
    X = X + " order by tblvoucher.vfirm,tblMastAccount.AcIsIntDbNtMonthly,tblMastAccount.AcOurCode,tblVoucher.Vdt,tblMastItem.ItHsn,tblVoucher.vno"
End If
x1Rec = GProcRstOpen(rsRecVsSale, X, "R")
With msGrid
    .Clear
    .FormatString = "<Bill No           |<Tp |<Bill Date    |>Bill Amount    |<Party Name                     |<Broker Name |>Amt. Recd     |<Pmt. Date  |>OD |>Discount         |>TDS On Amt.  |>TDS Amt.  |DN|DN in Bill No|||||||Cgst Rt|Sgst Rt|Igst Rt||TDS JV No.|HSN Cd||Debit Nt. No ||"
    .ColWidth(1) = 0
    .ColWidth(2) = 0
    .ColWidth(5) = 0
    .ColWidth(7) = 0
    .ColWidth(14) = 0
    .ColWidth(15) = 0
    .ColWidth(16) = 0
    .ColWidth(17) = 0
    .ColWidth(18) = 0
    .ColWidth(19) = 0
    .ColWidth(23) = 0
'    .ColWidth(24) = 0
    .ColWidth(26) = 0
    
    .ColWidth(10) = 0
    .ColWidth(11) = 0
    .ColWidth(20) = 0
    .ColWidth(21) = 0
    .ColWidth(22) = 0
    .ColWidth(24) = 0
    .ColWidth(25) = 0
    
    .ColWidth(28) = 0
    .ColWidth(29) = 0

    If FormAction = vbDataActionAddNew Then
        .rows = 2
        GridAddNew
    Else
        .rows = 2
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
With msGrid
     If ((.TextMatrix(.Row, 0)) = "") And .Row = 1 And txtGrid.Visible = True Then
        msGrid.Col = 9
        msGrid.SetFocus
        If txtGrid.Visible = True Then
           txtGrid.SetFocus
        End If
     End If
End With
End Sub
Private Sub MSGrid_LostFocus()
CalAmount
End Sub
Private Sub msGrid_EnterCell()
If FormAction <> vbDataActionClose Then
    CheckForNumberTxtGrid
    DoEvents
    txtGrid = ""
    DoEvents
    If msGrid.text <> "" Then txtGrid.text = msGrid.text
    '--- Lock TextBox
'0 - VBillNo
'1 - VBillType
'2 - VBillDate
'3 - Bill Amt
'4 - Pty Name
'5 - Brok NM
'6 - Amt Recd
'7 - Rec Dt
'8 - OD
'9 - Int
'10 - TDS On Amt
'11 - Tds
'12 - DbNtInt (Y/N)
'13 - DbNtIn Bill No
'14 - DbNtMonthly
'15 - Due On
'16 - Vno
'17 - VctrNo
'18 - Grase
'19 - IntRt
'20 - CRt
'21 - SRt
'22 - IRt
'23 - JvNoInt
'24 - JvNoLessTDS
'25 - Hsn Cd
'26 - VbillFirm
'28 -VnolistVyear
'29 -Vyear
    
    With msGrid
        Select Case .Col
            Case 0, 1, 2, 3, 4, 5, 6, 7, 8, 25, 26, 28, 29
                txtGrid.Locked = True
            Case Else
                txtGrid.Locked = False
'                If .Col = 9 Then
'                   OldTdsOn = Val(.TextMatrix(.Row, 9))
'                End If
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
        If (Val(.TextMatrix(.Row, 9)) = 0) And .rows > 2 Then
            OldTdsOn = 0
            .RemoveItem (.Row)
            txtGrid.text = ""
            If cmdBtn(13).Enabled = True Then
                cmdBtn(13).SetFocus
            Else
                cmdBtn(14).SetFocus
            End If
        End If
        txtGrid.Visible = False
        Exit Sub
    End If
    '---- Navigate Row ,Col
    If KeyAscii = 13 Then
        KeyAscii = 0
        If .Col = rsRecVsSale.Fields.Count - 14 Then
            If .Row = (.rows - 1) Then '-- Add New Row
                .rows = .rows + 1
                .Row = .Row + 1
'                GridAddNew
            Else '--Next Row
                .Row = .Row + 1
            End If
            .Col = 9
        Else    '-- Next col
            .Col = .Col + 1
        End If
    End If
End With

'---- Set TxtGrid
If Trim(msGrid.text) <> "" Then
    txtGrid = Trim(msGrid.text)
End If

If msGrid.Col < (rsRecVsSale.Fields.Count) Then   '- -3
    If msGrid.Col = 1 Then
       msGrid.Col = 3
    End If
    If msGrid.Col = 2 Then
       msGrid.Col = 3
    End If
    If msGrid.Col = 5 Then
       msGrid.Col = 6
    End If
    If msGrid.Col = 7 Then
       msGrid.Col = 8
    End If
    If msGrid.Col >= 10 And msGrid.Col <= 11 Then
       msGrid.Col = 12
    End If
    
    txtGrid.Move msGrid.CellLeft + msGrid.Left, msGrid.CellTop + msGrid.Top, msGrid.CellWidth, msGrid.CellHeight
End If
txtGrid.Visible = True
'SRK
txtGrid.Enabled = True
DoEvents
If KeyAscii <> 0 Then
    Select Case rsRecVsSale.Fields(msGrid.Col).Type
        Case adChar, adVarWChar
             '-- Pending
'            If msGrid.Col = 0 Then  '-- Invoice
'                txtGrid_KeyPress (KeyAscii)
'            Else
                txtGrid.text = txtGrid.text & Chr(KeyAscii)
'            End If
        Case Else
            txtGrid.text = Chr(GProcValidateKey(rsRecVsSale, msGrid.Col, KeyAscii, txtGrid.text)) 'Chr(GProcNumberWithDecimal(KeyAscii, txtGrid))
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
'If FormAction = 5 Or FormAction = 6 Then
If txtGrid.Visible Then
    msGrid.text = txtGrid.text
    DoEvents
    txtGrid.text = ""
    txtGrid.Visible = False
End If
CalGridAmt
CalAmount
'End If
End Sub
Private Sub CalBalance()
'Dim k As Integer
'Dim TAmt As Double
'With msGrid
'    k = 1
'    DoEvents
'    Do While k < .rows
'       TAmt = TAmt + Val(.TextMatrix(k, 6))
'       k = k + 1
'    Loop
'    txt(8) = Format(TAmt, FStr)
'End With
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
            Case 1 '-- Invoice
                DoEvents
                KeyAscii = 0
            Case Else
                KeyAscii = GProcValidateKey(rsRecVsSale, msGrid.Col, KeyAscii, txtGrid)
        End Select
   End If
   
    If KeyAscii = 13 Then
        CheckForNumberTxtGrid
    End If
    '--- Navigate Row ,Col
    
    If KeyAscii = 13 And .Col < rsRecVsSale.Fields.Count - 17 Then '- -4
        If .Col <= 8 Then
           .Col = 9
        Else
            .Col = .Col + 1
        End If
    ElseIf KeyAscii = 13 And .Col = rsRecVsSale.Fields.Count - 17 Then '- -4
        If .Row = (.rows - 1) Then
            .rows = .rows + 1
            .Row = .Row + 1
        Else
            .Row = .Row + 1
        End If
        .Col = 9
    End If
    '-- On ESC
    With msGrid
        If KeyAscii = 27 Then '----Esc Key
            KeyAscii = 0
            If (Val(.TextMatrix(.Row, 9)) = 0) And .rows > 2 Then
                OldTdsOn = 0
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
End With
End Sub
Private Sub txtGrid_KeyDown(KeyCode As Integer, Shift As Integer)
'--- Delete Row On Ctrl+Y
If KeyCode = vbKeyY And Shift = 2 And msGrid.rows > 2 Then
    msGrid = 9
    msGrid.RemoveItem (msGrid.Row)
    txtGrid.Visible = False
    txtGrid.text = ""
    CalAmount
    msGrid.SetFocus
    Exit Sub
End If
'If msGrid.Col >= 9 And msGrid.Col <= rsRecVsSale.Fields.Count - 7 Then  '- -4
'        If msGrid.Row < (msGrid.rows - 1) Then
'           msGrid.Row = msGrid.Row + 1
'           msGrid_KeyPress (0)
''           txtGrid.SetFocus
'        End If
'End If
End Sub
Private Sub CheckForNumberGrid()
Dim RowIndex As Long
With rsRecVsSale
For RowIndex = 1 To msGrid.rows - 1
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
Dim Y As String
Dim Y2 As String
Dim j As Long
Dim RsP As Recordset
Dim RsP2 As Recordset
Dim JvNo As Long
Dim JvNoTDs As Long
JvNo = 0
JvNoTDs = 0
gCn.Execute "Delete from tblVoucher where VType = 'PN' and VFirm = '" & gCCode & "'  and VYear = " & Val(txt(3)) & " and  vnolist  = " & txt(35)
If gCIsDepotFirm = 1 And gCDepotMainFirm > 0 Then
    gCn.Execute "Delete from tblVoucher where VType = 'PN' and VFirm = '" & gCDepotMainFirmCompCd & "'  and VYear = " & Val(txt(3)) & " and  vnolist  = " & txt(35)
End If
'---Delete JV OF TDS is pending 27/01/2018
'Y = "Select * from " & tblName & " where VnoList = " & txt(35) & " and JvNoLessTDS > 0  and " & CriteriaStr & " order by JvNoLessTDS  "
'i = GProcRstOpen(RsP, Y, "R")
'If Not RsP.EOF Then
'   JvNoTDs = RsP.Fields("JvNoLessTDS")
'   Y = "select VCtrNo from tblVoucher where VType='JV' and VNo=" & JvNoTDs & " and VFirm  = '" & gCCode & "' and  VYear=" & gCYear & "  order by VCtrNo"
'   i = GProcRstOpen(RsP2, Y, "R")
'   With RsP2
'        If Not .EOF Then
'           .MoveFirst
'           Do While Not .EOF
'              gCn.Execute "Delete from tblvoucher where VType = 'JV' and VFirm = '" & gCCode & "'  and VYear = " & Val(txt(3)) & "  and VCtrNo = " & RsP2.Fields(0) & "   and  vno = " & JvNoTDs
'              .MoveNext
'           Loop
'         End If
'    End With
'End If
DoEvents
gCn.Execute "update tblVoucher set InterestCredit = 0, JvNoIntInDepot = 0, LessTDS = 0 ,DueOnDt =  '', VnoList = 0 ,Grace = 0,IntRt = 0,LateDays = 0,JvNoDisc=0 ,JvNoLessTDS = 0,GraceExtra =0,IntFromDate = null,IntParty = 0,TdsRt = 0,TDSOn=0,DBNtInt = null,DbNtMonthly = 0,DbNtInBillNo = null,CrDrNoteNo = null,VBillFirm = null " _
                       & " where VType = '" & UCase(txt(1)) & "' and VFirm in (" & BillsInFirm & ")  and VYear = " & Val(txt(3)) & "  and VnoList = " & Val(txt(35)) & " "

rsRecVsSale.Requery

End Sub
Private Sub SaveRelatedRecord()
Dim i As Long, X As String
Dim RowIndex As Integer
Dim wCgstAmt As Double
Dim wSgstAmt As Double
Dim wIgstAmt As Double
RowIndex = 1
With msGrid
    Do While RowIndex < .rows
       If (Val(.TextMatrix(RowIndex, 9)) = 0) And .rows > 2 Then
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
        If CDec(Val(.TextMatrix(i, 9))) > 0 Then
           wCgstAmt = (Val(.TextMatrix(i, 9)) * Val(.TextMatrix(i, 20))) / 100
           wSgstAmt = (Val(.TextMatrix(i, 9)) * Val(.TextMatrix(i, 21))) / 100
           wIgstAmt = (Val(.TextMatrix(i, 9)) * Val(.TextMatrix(i, 22))) / 100
           wCgstAmt = Format(wCgstAmt, FStr)
           wSgstAmt = Format(wSgstAmt, FStr)
           wIgstAmt = Format(wIgstAmt, FStr)
           gCn.Execute "update tblVoucher set InterestCredit = " & Val(.TextMatrix(i, 9)) & ", LessTDS = " & Val(.TextMatrix(i, 11)) & ",DueOnDt = '" & IIf(IsDate(.TextMatrix(i, 15)), Format(.TextMatrix(i, 15), "yyyy/mm/dd"), "") & "', VnoList = " & Val(txt(35)) & " ,Grace = " & Val(.TextMatrix(i, 18)) & ",IntRt = " & Val(.TextMatrix(i, 19)) & ",LateDays = " & Val(.TextMatrix(i, 8)) & ",  " _
                       & " GraceExtra = " & Val(txt(39)) & ",IntFromDate = '" & IIf(IsDate(dtpDt(40)), Format(dtpDt(40), "yyyy/mm/dd"), "") & "' , IntParty = " & Val(txt(41)) & ",  TdsRt = " & Val(txt(42)) & ",TDSOn = " & Val(.TextMatrix(i, 10)) & ",DBNtInt = '" & UCase(.TextMatrix(i, 12)) & "',DbNtMonthly = " & Val(.TextMatrix(i, 14)) & ",DbNtInBillNo = '" & UCase(.TextMatrix(i, 13)) & "'," _
                       & " JvNoDisc=" & Val(.TextMatrix(i, 23)) & " ,JvNoLessTDS = " & Val(.TextMatrix(i, 24)) & ", CrDrNoteNo='" & .TextMatrix(i, 27) & "', " _
                       & " CgstRt=" & Val(.TextMatrix(i, 20)) & ",SgstRt=" & Val(.TextMatrix(i, 21)) & ",IgstRt=" & Val(.TextMatrix(i, 22)) & ",   " _
                       & " CgstAmt=" & wCgstAmt & ",SgstAmt=" & wSgstAmt & ",IgstAmt=" & wIgstAmt & ",VRefDate = '" & IIf(IsDate(dtpDt(11)), Format(dtpDt(11), "yyyy/mm/dd"), "") & "',IsGstDbNt = " & chkIsGSTDbNt.Value & ", VBillFirm = '" & UCase(.TextMatrix(i, 26)) & "',    " _
                       & "  IntToDate = '" & IIf(IsDate(dtpDt(2)), Format(dtpDt(2), "yyyy/mm/dd"), "") & "', VNoListVYear = " & gCYear & "    " _
                       & " where VNo = " & Val(.TextMatrix(i, 16)) & " and  VType = '" & UCase(txt(1)) & "' and  VFirm = '" & .TextMatrix(i, 26) & "'     and VYear = " & Val(.TextMatrix(i, 29)) & " and VCtrNo = " & Val(.TextMatrix(i, 17)) & " "
            
            DoEvents
        End If
       i = i + 1
     Loop
End With
If FormAction <> vbDataActionAddNew Then
   ApplyIntDbNt
End If
End Sub
Private Sub ApplyIntDbNt()
Dim Y As String
Dim RsP As Recordset
Dim wJvNo As Long
Dim wJvNoTmp As Long
Dim wJvNoTDS As Long
Dim wJvNoIntInDepot As Long
Dim wParty As Long
Dim wIntAmt As Double
Dim wTdsAmt As Double
Dim wBlNo As String
Dim wBlNoTds As String
Dim rsDbNt As Recordset
Dim rstSaleInt As Recordset
Dim i As Integer
Dim j As Integer
Dim wVBillNo As String
Dim wVBillVno As Long
Dim wVBillType As String
Dim wVBillDate As Date
Dim wVBillVYear As Integer
Dim wVBillFirm As String
Dim wVBillAmt As Double
Dim wTdsRt As Double
Dim wTdsOn As Double
Dim Mon As Integer
Dim wCgstAmt As Double
Dim wSgstAmt As Double
Dim wIgstAmt As Double
Dim wCgstRt As Double
Dim wSgstRt As Double
Dim wIgstRt As Double
Dim wDbNtNoTmp As String
Dim wDbNtNo As String
Dim wHsnCode As String
Dim wDbNtInBillNo As String
Dim wBroker As Integer
Dim wItcode As Integer
Dim wDbNtTot As Double
Dim wDbNtTotTmp As Double
Dim WcrNtAc As Long
'--For Log
Dim jj As Long
Dim XX As String
Dim tmpRst As Recordset
Dim wLogNo As Integer
Dim AMD As String
Dim wTp As String

If gCIsDepotFirm = 1 And gCDepotMainFirm > 0 Then
   WcrNtAc = GProcGetColumnValue("tblMastCompany", "CCode", gCCode, "S", "CCrNoteAc", "N")
Else
   WcrNtAc = 0
End If
Y = "Select * from tblVoucher where VType = 'PN' and  VFirm  = '" & gCCode & "' and VYear=" & gCYear
GProcRstOpen rstSaleInt, Y, "O"
Mon = Month(dtpDt(11))
DoEvents

If gCYear >= 2021 Then '--ItHsn2,ItHsn
    Y = "select tblVoucher.VNo,tblVoucher.VType,tblVoucher.Vdt,tblVoucher.VYear,tblVoucher.VFirm,VCtrNo, VCrAcCode,VDrAcCode,VAmt,VBillNo,VBillVno,VBillType,VBillDate,VBillVYear," _
        & "VBillAmt,LateDays,Interest,InterestCredit,JvNoDisc,JvNoInt,LessTDS,JvNoLessTDS,VnoList,DueOnDt,Grace,IntRt,GraceExtra,IntParty," _
        & "TdsRt,TDSOn,DbNtInt,DbNtMonthly,VBillFirm,DbNtInBillNo,JvNoIntInDepot,CgstRt,CgstAmt,SgstRt,SgstAmt,IgstRt,IgstAmt,CrDrNoteNo," _
        & "tblSaleSub.SlSubItCode , tblMastItem.ItHsn2,tblSale.SlBroker from tblVoucher,tblSaleSub,tblMastItem,tblSale,tblMastAccount " _
        & "where tblVoucher.VBillVno = tblSaleSub.VNo and tblVoucher.VBillType = tblSaleSub.VType and tblVoucher.VBillVYear  = tblSaleSub.VYear and tblVoucher.VCrAcCode = tblMastAccount.AcCode  " _
        & "and (case when isnull(tblvoucher.vBillfirm,'0')='0' then tblvoucher.vfirm else tblvoucher.vBillfirm end) = tblSaleSub.VFirm and tblSaleSub.SlSubItSrNo = 1 and tblSale.Vno = tblSaleSub.Vno and tblSale.Vtype = tblSaleSub.Vtype and tblSale.Vyear = tblSaleSub.Vyear and tblSale.Vfirm = tblSaleSub.Vfirm  " _
        & "and tblSaleSub.SlSubItCode = tblMastItem.ItCode and VnoList = " & txt(35) & " and DBNtInt = 'Y' and tblVoucher.VType = 'BR' and tblVoucher.VnoList <> 0 and tblVoucher.InterestCredit > 0 and tblVoucher.VFirm in (" & BillsInFirm & ")  and tblVoucher.VNoListVYear= " & gCYear & " "
    Y = Y + " order by VBillFirm,tblMastAccount.AcIsIntDbNtMonthly,tblMastAccount.AcOurCode,tblVoucher.VBillType,tblVoucher.Vdt,tblMastItem.ItHsn2,tblVoucher.vno"
Else
    Y = "select tblVoucher.VNo,tblVoucher.VType,tblVoucher.Vdt,tblVoucher.VYear,tblVoucher.VFirm,VCtrNo, VCrAcCode,VDrAcCode,VAmt,VBillNo,VBillVno,VBillType,VBillDate,VBillVYear," _
        & "VBillAmt,LateDays,Interest,InterestCredit,JvNoDisc,JvNoInt,LessTDS,JvNoLessTDS,VnoList,DueOnDt,Grace,IntRt,GraceExtra,IntParty," _
        & "TdsRt,TDSOn,DbNtInt,DbNtMonthly,VBillFirm,DbNtInBillNo,JvNoIntInDepot,CgstRt,CgstAmt,SgstRt,SgstAmt,IgstRt,IgstAmt,CrDrNoteNo," _
        & "tblSaleSub.SlSubItCode , tblMastItem.ItHsn,tblSale.SlBroker from tblVoucher,tblSaleSub,tblMastItem,tblSale,tblMastAccount " _
        & "where tblVoucher.VBillVno = tblSaleSub.VNo and tblVoucher.VBillType = tblSaleSub.VType and tblVoucher.VBillVYear  = tblSaleSub.VYear and tblVoucher.VCrAcCode = tblMastAccount.AcCode  " _
        & "and (case when isnull(tblvoucher.vBillfirm,'0')='0' then tblvoucher.vfirm else tblvoucher.vBillfirm end) = tblSaleSub.VFirm and tblSaleSub.SlSubItSrNo = 1 and tblSale.Vno = tblSaleSub.Vno and tblSale.Vtype = tblSaleSub.Vtype and tblSale.Vyear = tblSaleSub.Vyear and tblSale.Vfirm = tblSaleSub.Vfirm  " _
        & "and tblSaleSub.SlSubItCode = tblMastItem.ItCode and VnoList = " & txt(35) & " and DBNtInt = 'Y' and tblVoucher.VType = 'BR' and tblVoucher.VnoList <> 0 and tblVoucher.InterestCredit > 0 and tblVoucher.VFirm in (" & BillsInFirm & ")  and tblVoucher.VNoListVYear= " & gCYear & " "
    Y = Y + " order by VBillFirm,tblMastAccount.AcIsIntDbNtMonthly,tblMastAccount.AcOurCode,tblVoucher.VBillType,tblVoucher.Vdt,tblMastItem.ItHsn,tblVoucher.vno"
End If
If (GProcRstOpen(RsP, Y, "O") > 0) Then
    wParty = 0
    i = 1
    j = 1
    RsP.MoveFirst
    Do While Not RsP.EOF
       wParty = RsP.Fields("VCrAcCode")
       If gCYear >= 2021 Then '--ItHsn2,ItHsn
          wHsnCode = RsP.Fields("ItHsn2") & vbNullString
       Else
          wHsnCode = RsP.Fields("ItHsn") & vbNullString
       End If
       wDbNtInBillNo = RsP.Fields("DbNtInBillNo")
       wIntAmt = 0
       wTdsAmt = 0
       wTdsOn = 0
       wCgstAmt = 0
       wSgstAmt = 0
       wIgstAmt = 0
       wCgstRt = 0
       wSgstRt = 0
       wIgstRt = 0
       wDbNtTot = 0
       wBlNo = RsP.Fields("DbNtInBillNo")
       wBlNoTds = ""
       Do While Not RsP.EOF
'           If RsP.Fields("VCrAcCode") <> wParty And RsP.Fields("ItHsn") <> wHsnCode And RsP.Fields("DbNtInBillNo") <> wDbNtInBillNo Then
           If RsP.Fields("VCrAcCode") <> wParty And RsP.Fields("DbNtInBillNo") <> wDbNtInBillNo Then
              Exit Do
           End If
           wJvNoTmp = Val(RsP.Fields("JvNoDisc"))
           wJvNoIntInDepot = Val(RsP.Fields("JvNoIntInDepot"))
           wDbNtNoTmp = RsP.Fields("CrDrNoteNo") & vbNullString
           If wJvNoTmp > 0 Then
              wJvNo = wJvNoTmp
           Else
              wJvNo = GProcGenerateIdMonthwise("TblVoucher", "Vno", Mon, "Vdt", "Vtype='PN' and VFirm = '" & gCCode & "'   and VYear=" & gCYear)
           End If
           If wDbNtNoTmp <> "" Then
              wDbNtNo = wDbNtNoTmp
           Else
               wDbNtNo = GProcGenerateNoWithIBCrDrNt("PN", gCCode)
           End If
           DoEvents
           DoEvents
           DoEvents
           '-- for depot credit note
           If gCIsDepotFirm = 1 And gCDepotMainFirm > 0 Then
              If wJvNoIntInDepot > 0 Then
                 wJvNoIntInDepot = Val(RsP.Fields("JvNoIntInDepot"))
              Else
                  wJvNoIntInDepot = GProcGenerateIdMonthwise("TblVoucher", "Vno", Mon, "Vdt", "Vtype='PN' and VFirm = '" & gCDepotMainFirmCompCd & "'   and VYear=" & gCYear)
              End If
           End If
           '---
           If RsP.Fields("DbNtInBillNo") <> wBlNo Or RsP.Fields("DbNtInBillNo") = "" Then
              wJvNoTmp = Val(RsP.Fields("JvNoDisc"))
              wDbNtNoTmp = RsP.Fields("CrDrNoteNo") & vbNullString
              wIntAmt = Val(RsP.Fields("InterestCredit"))
              wTdsAmt = Val(RsP.Fields("LessTDS"))
              wTdsOn = Val(RsP.Fields("TDSOn"))
'              wCgstAmt = Round((Val(RsP.Fields("Interest")) * Val(RsP.Fields("CgstRt"))) / 100, 2)
'              wSgstAmt = Round((Val(RsP.Fields("Interest")) * Val(RsP.Fields("SgstRt"))) / 100, 2)
'              wIgstAmt = Round((Val(RsP.Fields("Interest")) * Val(RsP.Fields("IgstRt"))) / 100, 2)
              
              wDbNtTotTmp = Round(Val(RsP.Fields("InterestCredit")), 2)     '--+ Round((Val(RsP.Fields("InterestCredit")) * Val(RsP.Fields("CgstRt"))) / 100, 2) + Round((Val(RsP.Fields("Interest")) * Val(RsP.Fields("SgstRt"))) / 100, 2) + Round((Val(RsP.Fields("Interest")) * Val(RsP.Fields("IgstRt"))) / 100, 2)
              wDbNtTotTmp = GProcMakeRounding(wDbNtTotTmp)
              wDbNtTot = wDbNtTot + wDbNtTotTmp
              
'              wCgstRt = Val(RsP.Fields("CgstRt"))
'              wSgstRt = Val(RsP.Fields("SgstRt"))
'              wIgstRt = Val(RsP.Fields("IgstRt"))
              wBlNo = RsP.Fields("VBillNo")
              wVBillVno = RsP.Fields("VBillVno")
              wVBillType = RsP.Fields("VBillType")
              wVBillDate = RsP.Fields("VBillDate")
              wVBillVYear = RsP.Fields("VBillVYear")
              wVBillFirm = RsP.Fields("VBillFirm")
              wVBillAmt = RsP.Fields("VBillAmt")
              wBroker = RsP.Fields("SlBroker")
              wItcode = RsP.Fields("SlSubItCode")
              gCn.Execute "update tblVoucher set JvNoDisc = " & wJvNo & ",JvNoIntInDepot = " & wJvNoIntInDepot & ", CrDrNoteNo= '" & wDbNtNo & "' " _
                        & " where vno = " & RsP.Fields("Vno") & " and Vtype = '" & RsP.Fields("VType") & "' and Vyear = " & RsP.Fields("VYear") & " and Vfirm = '" & RsP.Fields("Vfirm") & "'  and VCtrNo = " & RsP.Fields("VCtrNo") & ""
              RsP.MoveNext
              Exit Do
           Else
              wJvNoTmp = Val(RsP.Fields("JvNoDisc"))
              wDbNtNoTmp = RsP.Fields("CrDrNoteNo") & vbNullString
              wIntAmt = wIntAmt + Val(RsP.Fields("InterestCredit"))
              wTdsAmt = wTdsAmt + Val(RsP.Fields("LessTDS"))
              wTdsOn = wTdsOn + Val(RsP.Fields("TDSOn"))
'              wCgstAmt = wCgstAmt + Round((Val(RsP.Fields("Interest")) * Val(RsP.Fields("CgstRt"))) / 100, 2)
'              wSgstAmt = wSgstAmt + Round((Val(RsP.Fields("Interest")) * Val(RsP.Fields("SgstRt"))) / 100, 2)
'              wIgstAmt = wIgstAmt + Round((Val(RsP.Fields("Interest")) * Val(RsP.Fields("IgstRt"))) / 100, 2)
              
              wDbNtTotTmp = Round(Val(RsP.Fields("InterestCredit")), 2)    '--+ Round((Val(RsP.Fields("Interest")) * Val(RsP.Fields("CgstRt"))) / 100, 2) + Round((Val(RsP.Fields("Interest")) * Val(RsP.Fields("SgstRt"))) / 100, 2) + Round((Val(RsP.Fields("Interest")) * Val(RsP.Fields("IgstRt"))) / 100, 2)
              wDbNtTotTmp = GProcMakeRounding(wDbNtTotTmp)
              wDbNtTot = wDbNtTot + wDbNtTotTmp
              
              
'              wCgstRt = Val(RsP.Fields("CgstRt"))
'              wSgstRt = Val(RsP.Fields("SgstRt"))
'              wIgstRt = Val(RsP.Fields("IgstRt"))
              If RsP.Fields("VBillNo") = wBlNo Then
                 wVBillVno = RsP.Fields("VBillVno")
                 wVBillType = RsP.Fields("VBillType")
                 wVBillDate = RsP.Fields("VBillDate")
                 wVBillVYear = RsP.Fields("VBillVYear")
                 wVBillFirm = RsP.Fields("VBillFirm")
                 wVBillAmt = RsP.Fields("VBillAmt")
                 wBroker = RsP.Fields("SlBroker")
                 wItcode = RsP.Fields("SlSubItCode")
              End If
              gCn.Execute "update tblVoucher set JvNoDisc = " & wJvNo & ",JvNoIntInDepot = " & wJvNoIntInDepot & ", CrDrNoteNo= '" & wDbNtNo & "' " _
                        & " where vno = " & RsP.Fields("Vno") & " and Vtype = '" & RsP.Fields("VType") & "' and Vyear = " & RsP.Fields("VYear") & " and Vfirm = '" & RsP.Fields("Vfirm") & "'  and VCtrNo = " & RsP.Fields("VCtrNo") & ""
              RsP.MoveNext
              If Not RsP.EOF Then
                 If RsP.Fields("DbNtInBillNo") <> wBlNo Then Exit Do
              End If
           End If
       Loop
       '--- Interest Db Nt.
       If wIntAmt > 0 Then
            i = 1
            gCn.BeginTrans
            rstSaleInt.AddNew
            rstSaleInt!Vno = wJvNo '--VNo
            rstSaleInt!Vtype = "PN"  '--Vtype
            rstSaleInt!Vdt = CDate(txt(11)) '--Vdt(2)
            rstSaleInt!VYear = txt(3) '--VYear
            rstSaleInt!VFirm = gCCode '--VFirm
            rstSaleInt!VCtrNo = i '--Sr No
            i = i + 1
            rstSaleInt!VCrAcCode = wParty
            If gCIsDepotFirm = 1 And gCDepotMainFirm > 0 Then
                rstSaleInt!VDrAcCode = gCDepotMainFirm
            Else
                rstSaleInt!VDrAcCode = gDiscAcCodeRec
            End If
            rstSaleInt!vamt = wDbNtTot
            rstSaleInt!VNar1 = "Early Payment Credited "
            rstSaleInt!VNar2 = "Agt. Bill No : " & wBlNo
'            rstSaleInt!VNar3 = ""
'            rstSaleInt!VNAR4 = ""
            rstSaleInt!InterestCredit = wIntAmt
            rstSaleInt!VBillNo = wBlNo
            rstSaleInt!CrDrNoteNo = wDbNtNo
            rstSaleInt!VBillVno = wVBillVno
            rstSaleInt!VBillType = wVBillType
            rstSaleInt!VBillDate = wVBillDate
            rstSaleInt!VBillVYear = wVBillVYear
            rstSaleInt!VBillAmt = wVBillAmt
            rstSaleInt!VBillFirm = wVBillFirm
            rstSaleInt!VnoList = Val(txt(35))
            rstSaleInt!CgstRt = wCgstRt
            rstSaleInt!CgstAmt = wCgstAmt
            rstSaleInt!SgstRt = wSgstRt
            rstSaleInt!SgstAmt = wSgstAmt
            rstSaleInt!IgstRt = wIgstRt
            rstSaleInt!IgstAmt = wIgstAmt
'            rstSaleInt!Broker = wBroker
'            rstSaleInt!ItCode = wItcode
            rstSaleInt!IsGstDbNt = chkIsGSTDbNt.Value
            rstSaleInt!VIsRefEntType = 1
            rstSaleInt.Update
            gCn.CommitTrans
            '-- For Log
            If gCYear >= 2023 And gCIsLog = 1 Then
               XX = "Select max(LogNo) from tblVoucher_Log where Vno=" & wJvNo & " and vtype = 'PN' and Vfirm = '" & gCCode & "' and vyear = " & txt(3) & " and VCtrNo = " & i - 1 & " "
               jj = GProcRstOpen(tmpRst, XX, "R")
               If IIf(IsNull(tmpRst.Fields(0)), 0, tmpRst.Fields(0)) > 0 Then
                  wLogNo = IIf(IsNull(tmpRst.Fields(0)), 0, tmpRst.Fields(0)) + 1
                  AMD = "M"
               Else
                  wLogNo = 1
                  AMD = "A"
               End If
               gCn.Execute "Insert into tblVoucher_Log select tblVoucher.*," & wLogNo & " as LogNo,'" & AMD & "' as LogTp ,'" & gUserName & "' as UserName ,'" & Format(Now, "YYYY/MM/DD hh:mm:ss") & "' as LogDate from tblVoucher where Vno=" & wJvNo & " and vtype = 'PN' and Vfirm = '" & gCCode & "' and vyear = " & txt(3) & " and VCtrNo = " & i - 1 & " "
               tmpRst.Close
            End If
           '-- For Credite Note of Depot
           If gCIsDepotFirm = 1 And gCDepotMainFirm > 0 Then
                gCn.BeginTrans
                rstSaleInt.AddNew
                rstSaleInt!Vno = wJvNoIntInDepot '--VNo
                rstSaleInt!Vtype = "PN"  '--Vtype
                rstSaleInt!Vdt = CDate(txt(11)) '--Vdt(2)
                rstSaleInt!VYear = txt(3) '--VYear
                rstSaleInt!VFirm = gCDepotMainFirmCompCd '--VFirm
                rstSaleInt!VCtrNo = i - 1 '--Sr No
                rstSaleInt!VCrAcCode = WcrNtAc
                rstSaleInt!VDrAcCode = gDiscAcCodeRec
                rstSaleInt!vamt = wDbNtTot
                rstSaleInt!VNar1 = "Early Payment Credited "
                rstSaleInt!VNar2 = "Agt. Bill No : " & wBlNo
                rstSaleInt!VNar3 = Left(GProcGetColumnValue("tblMastAccount", "AcCode", "" & wParty & "", "N", "AcName", "S"), 50)
    '            rstSaleInt!VNAR4 = ""
                rstSaleInt!InterestCredit = wIntAmt
'                rstSaleInt!VBillNo = wBlNo
                rstSaleInt!CrDrNoteNo = wDbNtNo
'                rstSaleInt!VBillVno = wVBillVno
'                rstSaleInt!VBillType = wVBillType
'                rstSaleInt!VBillDate = wVBillDate
'                rstSaleInt!VBillVYear = wVBillVYear
'                rstSaleInt!VBillAmt = wVBillAmt
'                rstSaleInt!VBillFirm = wVBillFirm
                rstSaleInt!VnoList = Val(txt(35))
                rstSaleInt!CgstRt = wCgstRt
                rstSaleInt!CgstAmt = wCgstAmt
                rstSaleInt!SgstRt = wSgstRt
                rstSaleInt!SgstAmt = wSgstAmt
                rstSaleInt!IgstRt = wIgstRt
                rstSaleInt!IgstAmt = wIgstAmt
    '            rstSaleInt!Broker = wBroker
    '            rstSaleInt!ItCode = wItcode
                rstSaleInt!IsGstDbNt = chkIsGSTDbNt.Value
                rstSaleInt!VIsRefEntType = 1
                rstSaleInt.Update
                gCn.CommitTrans
                If gCYear >= 2023 And gCIsLog = 1 Then
                   XX = "Select max(LogNo) from tblVoucher_Log where Vno=" & wJvNoIntInDepot & " and vtype = 'PN' and Vfirm = '" & gCDepotMainFirmCompCd & "' and vyear = " & txt(3) & " and VCtrNo = " & i - 1 & " "
                   jj = GProcRstOpen(tmpRst, XX, "R")
                   If IIf(IsNull(tmpRst.Fields(0)), 0, tmpRst.Fields(0)) > 0 Then
                      wLogNo = IIf(IsNull(tmpRst.Fields(0)), 0, tmpRst.Fields(0)) + 1
                      AMD = "M"
                   Else
                      wLogNo = 1
                      AMD = "A"
                   End If
                   gCn.Execute "Insert into tblVoucher_Log select tblVoucher.*," & wLogNo & " as LogNo,'" & AMD & "' as LogTp ,'" & gUserName & "' as UserName ,'" & Format(Now, "YYYY/MM/DD hh:mm:ss") & "' as LogDate from tblVoucher where Vno=" & wJvNoIntInDepot & " and vtype = 'PN' and Vfirm = '" & gCDepotMainFirmCompCd & "' and vyear = " & txt(3) & " and VCtrNo = " & i - 1 & " "
                   tmpRst.Close
                End If
                
                
                
           End If
        
        End If
    Loop
End If

If GProcRstOpen(rstbl, "Select * from " & tblName & " where VnoList = " & txt(35) & " and  " & CriteriaStr & "   ", "O") > 0 Then ReadFields
End Sub
Private Sub FillRelatedRecord()
Dim RowIndex As Long
Dim ColIndex As Long
Dim BnkTot As Double, c1 As Integer, r1 As Integer

'--Fill Grid
SetGrid
BnkTot = 0
With rsRecVsSale
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
'                If Val(.Fields(14)) = 1 And ColIndex = 0 Then
'                   msGrid.CellForeColor = &HC0&
'                Else
'                   msGrid.CellForeColor = &H0&
'                End If
'                If ColIndex = 32 Then BnkTot = BnkTot + Val(msGrid.TextMatrix(RowIndex, 32)) '-- jv
            Next ColIndex
                If msGrid.TextMatrix(RowIndex, 14) = 1 Then
                   c1 = msGrid.Col
                   r1 = msGrid.Row
                   msGrid.Row = RowIndex
                   Dim k As Integer
                   For k = 0 To 13
                   msGrid.Col = k
                   msGrid.CellForeColor = &HC0&
                   Next k
                   msGrid.Col = c1
                   msGrid.Row = r1
                End If
            RowIndex = RowIndex + 1
            .MoveNext
            If Not .EOF Then msGrid.rows = msGrid.rows + 1
        Loop
    End If
End With
CalAmount
End Sub
Private Sub CalAmount()
Dim BnkTot As Double
Dim BnkTot2 As Double
Dim BnkTot3 As Double
Dim RowIndex As Integer
RowIndex = 1
BnkTot = 0
BnkTot2 = 0
BnkTot3 = 0
With msGrid
    Do While RowIndex < .rows
        '-- sub amt
        If Val(.TextMatrix(RowIndex, 9)) > 0 Then
            If UCase(.TextMatrix(RowIndex, 12)) = "Y" Then
               BnkTot = BnkTot + Val(.TextMatrix(RowIndex, 9))
               BnkTot2 = BnkTot2 + Val(.TextMatrix(RowIndex, 11))
            Else
               BnkTot3 = BnkTot3 + Val(.TextMatrix(RowIndex, 9))
            End If
        End If
        RowIndex = RowIndex + 1
    Loop
End With
lblInt.Caption = Format(BnkTot, "###0.00")
lblTDS.Caption = Format(BnkTot2, "###0.00")
lblInt2.Caption = Format(BnkTot3, "###0.00")
'txt(8) = Format(BnkTot, "###0.00")
End Sub
Private Sub SetControlEd()
txt(35).Enabled = False

If FormAction = vbDataActionUpdate Then
   dtpDt(40).Enabled = False
   dtpDt(2).Enabled = False
   txttemp(41).Enabled = False
   txt(37).Enabled = False
   txt(38).Enabled = False
   txt(39).Enabled = False
   txt(42).Enabled = False
   chkIsMonthlyDbNt.Enabled = False
   chkIsGSTDbNt.Enabled = False
   txt(35).Enabled = True
End If
txt(32).Enabled = False
txt(34).Enabled = False

End Sub
Private Sub PreperReport()
Dim X As String
Dim wFromCd As Long
Dim LateDayCd As String
Dim CrNtOnAmt As Double
Dim vamt As Double
Dim Lint As Double
Dim LTds As Double
Dim RsTmp As Recordset
Dim i As Long, j As Long
Dim NewRow As Boolean
Dim IntRt As Double
Dim TotCol(0, 1) As Long
Dim CotCrt As Double
Dim CotSrt As Double
Dim CotIrt As Double
Dim PolCrt As Double
Dim PolSrt As Double
Dim PolIrt As Double
Dim wgTin As String
CotCrt = 0
CotSrt = 0
CotIrt = 0
PolCrt = 0
PolSrt = 0
PolIrt = 0

X = "select CotCGSTRt,CotSGSTRt,CotIGSTRt,PolCGSTRt,PolSGSTRt,PolIGSTRt from tblMastNarration where NarrType = 'T' and Narration = 'G S T'"

If GProcRstOpen(RsTmp, X, "R", gCn) > 0 Then
   CotCrt = Format(RsTmp.Fields(0), FStr)
   CotSrt = Format(RsTmp.Fields(1), FStr)
   CotIrt = Format(RsTmp.Fields(2), FStr)
   PolCrt = Format(RsTmp.Fields(3), FStr)
   PolSrt = Format(RsTmp.Fields(4), FStr)
   PolIrt = Format(RsTmp.Fields(5), FStr)
End If
DoEvents
DoEvents
               
'X = "Select CCode from tblMastCompany where CDepotMainFirmCompCd = '" & gCCode & "'"
'i = GProcRstOpen(rstblBlChk, X, "R")
'BillsInFirm = "'" & gCCode & "'"
'If i > 0 Then
'   rstblBlChk.MoveFirst
'   Do While Not rstblBlChk.EOF
'      BillsInFirm = BillsInFirm + "," + "'" & rstblBlChk.Fields(0) & "'"
'      rstblBlChk.MoveNext
'   Loop
'End If
'0 - VBillNo
'1 - VBillType
'2 - VBillDate
'3 - Bill Amt
'4 - Pty Name
'5 - Brok NM
'6 - Amt Recd
'7 - Rec Dt
'8 - OD
'9 - Int
'10 - TDS On Amt
'11 - Tds
'12 - DbNtInt (Y/N)
'13 - DbNtIn Bill No
'14 - DbNtMonthly
'15 - Due On
'16 - Vno
'17 - VctrNo
'18 - Grase
'19 - IntRt

If gCYear >= 2021 Then
    X = "Select a.VBillNo,d.Vtype,a.VBillDate,a.VBillVYear,d.VFirm,d.SlBillAmt,a.VAmt,a.Vdt,b.AcOurCode,b.Acname,dateadd(dd,d.DueDays+d.SlCrDays+iif(b.AcDueDays>0,b.AcDueDays," & Val(txt(37)) & ") ,d.Vdt) , " _
                       & " c.AcOurCode,c.AcName,a.Vtype,'" & gUserName & "', DateDiff(day, d.Vdt, a.Vdt)-(d.DueDays+d.SlCrDays)-iif(b.AcDueDays>0,b.AcDueDays," & Val(txt(37)) & ") as OD,a.vno,a.VCtrNo,a.TDSOn,b.AcIntPer,b.AcIsIntDbNtMonthly,iif(b.AcDueDays>0,b.AcDueDays," & Val(txt(37)) & ") as Gease,DbNtInBillNo,d.SltaxRate,d.SltaxRate2,d.SltaxRate3,f.ittype,f.ItHsn2,a.Vyear as ReceiptYear from tblVoucher a , tblMastAccount b, tblMastAccount c, tblsale d,tblsaleSub e, tblMastItem f " _
                       & " where a.VBillVno = d.VNo and a.VBillType = d.VType and a.VBillVYear = d.Vyear and a.VFirm = d.VFirm " _
                       & " and d.SlAcDrCode = b.AcCode and d.SlBroker = c.AcCode and a.Vtype = 'BR' and d.vno = e.vno and d.vtype = e.vtype and d.vyear = e.vyear and d.vfirm = e.vfirm and  e.SlSubItSrNo = 1 and e.SlSubItCode = f.ItCode" _
                       & " and a.VDt >= '" & Format(dtpDt(40), "yyyy/mm/dd") & "'  and a.VDt <= '" & Format(dtpDt(2), "yyyy/mm/dd") & "' and VnoList = 0 "
Else
    X = "Select a.VBillNo,d.Vtype,a.VBillDate,a.VBillVYear,d.VFirm,d.SlBillAmt,a.VAmt,a.Vdt,b.AcOurCode,b.Acname,dateadd(dd,d.DueDays+d.SlCrDays+iif(b.AcDueDays>0,b.AcDueDays," & Val(txt(37)) & ") ,d.Vdt) , " _
                       & " c.AcOurCode,c.AcName,a.Vtype,'" & gUserName & "', DateDiff(day, d.Vdt, a.Vdt)-(d.DueDays+d.SlCrDays)-iif(b.AcDueDays>0,b.AcDueDays," & Val(txt(37)) & ") as OD,a.vno,a.VCtrNo,a.TDSOn,b.AcIntPer,b.AcIsIntDbNtMonthly,iif(b.AcDueDays>0,b.AcDueDays," & Val(txt(37)) & ") as Gease,DbNtInBillNo,d.SltaxRate,d.SltaxRate2,d.SltaxRate3,f.ittype,f.ItHsn,a.Vyear as ReceiptYear from tblVoucher a , tblMastAccount b, tblMastAccount c, tblsale d,tblsaleSub e, tblMastItem f " _
                       & " where a.VBillVno = d.VNo and a.VBillType = d.VType and a.VBillVYear = d.Vyear and a.VFirm = d.VFirm " _
                       & " and d.SlAcDrCode = b.AcCode and d.SlBroker = c.AcCode and a.Vtype = 'BR' and d.vno = e.vno and d.vtype = e.vtype and d.vyear = e.vyear and d.vfirm = e.vfirm and  e.SlSubItSrNo = 1 and e.SlSubItCode = f.ItCode" _
                       & " and a.VDt >= '" & Format(dtpDt(40), "yyyy/mm/dd") & "'  and a.VDt <= '" & Format(dtpDt(2), "yyyy/mm/dd") & "' and VnoList = 0 "
End If
If chkIsGSTDbNt.Value = 1 Then
   X = X + " and a.Vfirm  in (" & BillsInFirm & ") and a.VBillDate >= '" & Format("2017/07/01", "yyyy/mm/dd") & "'  "
Else
   X = X + " and a.Vfirm  in (" & BillsInFirm & ") and a.VBillDate < '" & Format("2017/07/01", "yyyy/mm/dd") & "'  "
End If
If txttemp(41) <> "" Then
   X = X + " and b.Acname = '" & txttemp(41) & "'"
End If

If chkIsMonthlyDbNt.Value = 1 Then
   X = X + " and b.AcIsIntDbNtMonthly = 1"
End If
If gCYear >= 2021 Then
   X = X + " order by d.VFirm,b.AcIsIntDbNtMonthly,b.AcOurCode,d.Vtype,d.Vdt,f.ItHsn2,d.vno"
Else
   X = X + " order by d.VFirm,b.AcIsIntDbNtMonthly,b.AcOurCode,d.Vtype,d.Vdt,f.ItHsn,d.vno"
End If
If GProcRstOpen(RsTmp, X, "R", gCn) > 0 Then
   RsTmp.MoveFirst
   Do While Not RsTmp.EOF
      With msGrid
           If Val(RsTmp.Fields(15)) < 0 Then
                j = .rows - 1
                .TextMatrix(j, 0) = RsTmp.Fields(0)                    '0 - VBillNo
                .TextMatrix(j, 1) = RsTmp.Fields(1)                    '1 - VBillType
                .TextMatrix(j, 2) = RsTmp.Fields(2)                    '2 - VBillDate
                .TextMatrix(j, 3) = CDbl(RsTmp.Fields(5))              '3 - Bill Amt
                .TextMatrix(j, 4) = RsTmp.Fields(9)                    '4 - Pty Name
                .TextMatrix(j, 5) = RsTmp.Fields(12)                   '5 - Brok NM
                .TextMatrix(j, 6) = CDbl(RsTmp.Fields(6))              '6 - Amt Recd
                .TextMatrix(j, 7) = RsTmp.Fields(7)                    '7 - Rec Dt
                .TextMatrix(j, 8) = RsTmp.Fields(15)                   '8 - OD
'                If RsTmp.Fields("AcIntPer") > 0 Then
 '                  IntRt = RsTmp.Fields("AcIntPer")
  '              Else
                   IntRt = Val(txt(38))
   '             End If
                Lint = Round((Val(RsTmp.Fields(5)) * IntRt) / 100, 0) '-- Int On Bill Amt
'                Lint = (((CDbl(RsTmp.Fields(5)) * IntRt) / 100) / 365) * (Val(RsTmp.Fields(15)))
                If Abs(Val(RsTmp.Fields(5)) - (Val(RsTmp.Fields(6)) + Lint)) <= Lint + 10 Then
                   CrNtOnAmt = Val(CDbl(RsTmp.Fields(5)))
                Else
                   CrNtOnAmt = Val(CDbl(RsTmp.Fields(6)))
                End If
                DoEvents
                Lint = Format(Round(((CrNtOnAmt * IntRt) / 100), 0), "#######0.00")
 '               Lint = (((CrNtOnAmt * IntRt) / 100) / 365) * (Val(RsTmp.Fields(15)))
                Lint = GProcMakeRounding(Lint)
                DoEvents
                .TextMatrix(j, 9) = Format(Lint, "##########0.00")      '9 - Int
                If RsTmp.Fields(18) > 0 Then
                   .TextMatrix(j, 10) = RsTmp.Fields(18)                '10 - TDS On Amt
                Else
                   .TextMatrix(j, 10) = Format(Lint, "##########0.00")  '10 - TDS On Amt
                End If
                LTds = GProcMakeRounding((.TextMatrix(j, 10) * Val(txt(42))) / 100)
                DoEvents
                
                .TextMatrix(j, 28) = gCYear
                .TextMatrix(j, 29) = RsTmp.Fields("ReceiptYear")
                
                
                
                .TextMatrix(j, 11) = LTds         '11 - Tds
                .TextMatrix(j, 12) = "Y"                                                         '12 - DbNtInt (Y/N)
                .TextMatrix(j, 13) = RsTmp.Fields("DbNtInBillNo") & vbNullString                 '13 - DbNtIn Bill No
                .TextMatrix(j, 14) = RsTmp.Fields("AcIsIntDbNtMonthly")                          '14 - DbNtMonthly
                .TextMatrix(j, 15) = RsTmp.Fields(10)                                            '15 - Due On
                .TextMatrix(j, 16) = RsTmp.Fields(16)                                            '16 - Vno
                .TextMatrix(j, 17) = RsTmp.Fields(17)                                            '17 - VctrNo
                .TextMatrix(j, 18) = RsTmp.Fields("Gease")                                       '18 - Grase
                .TextMatrix(j, 19) = IntRt                                                       '19 - IntRt
                '---Tax Rate From Item Type
                If Format(dtpDt(11), "YYYY/MM/dd") <= Format("2018/04/01", "YYYY/MM/dd") Then
                   wgTin = GProcGetColumnValue("TblMastAccount", "ACName", RsTmp.Fields(9), "S", "AcGSTIN", "S")
                   If Left(wgTin, 2) = "27" Or Left(wgTin, 1) = "U" Or Trim(wgTin) = "" Then
                      .TextMatrix(j, 20) = Format(0, FStr)      '--              Format(RsTmp.Fields("SltaxRate") / 2, FStr)             '20 - Cgst rt
                      .TextMatrix(j, 21) = Format(0, FStr)      '--Format(RsTmp.Fields("SltaxRate") / 2, FStr)             '21 - Sgst rt
                      .TextMatrix(j, 22) = Format(0, FStr)     '-- 0
                   Else
                      .TextMatrix(j, 20) = Format(0, FStr)  '--            '20 - Cgst rt
                      .TextMatrix(j, 21) = Format(0, FStr)  '--            '21 - Sgst rt
                      .TextMatrix(j, 22) = Format(0, FStr)  '--    Format(RsTmp.Fields("SltaxRate"), FStr)
                   End If
                ElseIf Format(dtpDt(11), "YYYY/MM/dd") >= Format("2018/11/06", "YYYY/MM/dd") Then
                    If RsTmp.Fields(1) = "SY" Then
                       .TextMatrix(j, 20) = Format(0, FStr)  '--         Format(RsTmp.Fields("SltaxRate"), FStr)                     '20 - Cgst rt
                       .TextMatrix(j, 21) = Format(0, FStr)  '--         Format(RsTmp.Fields("SltaxRate2"), FStr)                    '20 - Sgst rt
                       .TextMatrix(j, 22) = Format(0, FStr)  '--         Format(RsTmp.Fields("SltaxRate3"), FStr)                    '20 - Igst rt
                    Else
                       .TextMatrix(j, 20) = Format(0, FStr)                     '20 - Cgst rt
                       .TextMatrix(j, 21) = Format(0, FStr)                    '20 - Sgst rt
                       .TextMatrix(j, 22) = Format(0, FStr)                    '20 - Igst rt
                    End If
                Else
                    If RsTmp.Fields(1) = "SM" Then
                        wgTin = GProcGetColumnValue("TblMastAccount", "ACName", RsTmp.Fields(9), "S", "AcGSTIN", "S")
                        If RsTmp.Fields("ittype") = 0 Then
                           If Left(wgTin, 2) = "27" Or Left(wgTin, 1) = "U" Or Trim(wgTin) = "" Then
                              .TextMatrix(j, 20) = Format(0, FStr)   '--   Format(CotCrt, FStr)                                      '20 - Cgst rt
                              .TextMatrix(j, 21) = Format(0, FStr)   '--   Format(CotSrt, FStr)                                      '21 - Sgst rt
                              .TextMatrix(j, 22) = Format(0, FStr)   '--   0
                           Else
                              .TextMatrix(j, 20) = Format(0, FStr)
                              .TextMatrix(j, 21) = Format(0, FStr)
                              .TextMatrix(j, 22) = Format(0, FStr)  ' -- Format(CotIrt, FStr)
                           End If
                        Else
                           If Left(wgTin, 2) = "27" Or Left(wgTin, 1) = "U" Or Trim(wgTin) = "" Then
                              .TextMatrix(j, 20) = Format(0, FStr)   '--  Format(PolCrt, FStr)                                      '20 - Cgst rt
                              .TextMatrix(j, 21) = Format(0, FStr)   '--  Format(PolSrt, FStr)                                      '21 - Sgst rt
                              .TextMatrix(j, 22) = Format(0, FStr)
                              If Format(RsTmp.Fields(2), "YYYY/MM/dd") <= Format("2017/10/12", "yyyy/mm/dd") Then
                                .TextMatrix(j, 20) = Format(0, FStr)  '--  Format(9, FStr)                                      '20 - Cgst rt
                                .TextMatrix(j, 21) = Format(0, FStr)  '--  Format(9, FStr)                                      '21 - Sgst rt
                                .TextMatrix(j, 22) = Format(0, FStr)
                              End If
                           Else
                              .TextMatrix(j, 20) = Format(0, FStr)
                              .TextMatrix(j, 21) = Format(0, FStr)
                              .TextMatrix(j, 22) = Format(0, FStr)  '--  Format(PolIrt, FStr)                                      '22 - Igst Rt
                              If Format(RsTmp.Fields(2), "YYYY/MM/dd") <= Format("2017/10/12", "YYYY/MM/dd") Then
                                .TextMatrix(j, 20) = Format(0, FStr)
                                .TextMatrix(j, 21) = Format(0, FStr)
                                .TextMatrix(j, 22) = Format(0, FStr)  '-- Format(18, FStr)
                              End If
                           
                           End If
                        End If
                    Else
                        '---Tax Rate From Sale Bill
                        .TextMatrix(j, 20) = Format(0, FStr)  '--  Format(RsTmp.Fields("SltaxRate"), FStr)                     '20 - Cgst rt
                        .TextMatrix(j, 21) = Format(0, FStr)  '--   Format(RsTmp.Fields("SltaxRate2"), FStr)                    '20 - Sgst rt
                        .TextMatrix(j, 22) = Format(0, FStr)  '--   Format(RsTmp.Fields("SltaxRate3"), FStr)                    '20 - Igst rt
                    End If
                End If
                If gCYear >= 2021 Then  '--ItHsn2,ItHsn
                   .TextMatrix(j, 25) = RsTmp.Fields("ItHsn2") & vbNullString                                        '25 - Hsn Cd
                Else
                   .TextMatrix(j, 25) = RsTmp.Fields("ItHsn") & vbNullString                                        '25 - Hsn Cd
                End If
                .TextMatrix(j, 26) = RsTmp.Fields("VFirm") & vbNullString                                        '25 - Hsn Cd
                If Format(RsTmp.Fields(2), "YYYY/MM/dd") <= Format("2017/07/01", "YYYY/MM/dd") Then
                   .TextMatrix(j, 20) = Format(0, FStr)
                   .TextMatrix(j, 21) = Format(0, FStr)
                   .TextMatrix(j, 22) = Format(0, FStr)
                   .TextMatrix(j, 25) = ""  '-- HSN
                End If
                
                If .TextMatrix(j, 14) = 1 Then
                   .Row = j
                   Dim k As Integer
                   For k = 0 To 13
                   .Col = k
                   .CellForeColor = &HC0&
                   Next k
'                   .Row.C
 '                   Row.cells.ForeColor = System.Drawing.color.Red
'                   .ForeColor = vbRed
                End If
                j = j + 1
                .rows = .rows + 1
                .Col = 9
           End If
      End With
      RsTmp.MoveNext
   Loop

'----
Dim RowIndex As Integer
Dim LastBl As String
Dim TmpParty As String
Dim TmpHSN As String
Dim TmpRowIndex As Integer

RowIndex = 1
With msGrid
    Do While RowIndex < .rows
       If .TextMatrix(RowIndex, 4) = "" Then
           RowIndex = RowIndex + 1
       Else
            TmpParty = .TextMatrix(RowIndex, 4) & vbNullString
            TmpHSN = .TextMatrix(RowIndex, 25) & vbNullString
            TmpRowIndex = RowIndex
            Do While RowIndex < .rows And .TextMatrix(RowIndex, 4) = TmpParty And .TextMatrix(RowIndex, 25) = TmpHSN
               LastBl = .TextMatrix(RowIndex, 0) & vbNullString
               RowIndex = RowIndex + 1
            Loop
            RowIndex = TmpRowIndex
            DoEvents
            DoEvents
            DoEvents
            Do While RowIndex < .rows And .TextMatrix(RowIndex, 4) = TmpParty And .TextMatrix(RowIndex, 25) = TmpHSN
               .TextMatrix(RowIndex, 13) = LastBl & vbNullString
               RowIndex = RowIndex + 1
            Loop
    End If
    Loop
End With

If msGrid.rows >= 2 Then msGrid.rows = msGrid.rows - 1
'----
msGrid.Col = 9
msGrid.SetFocus
txtGrid.SetFocus
End If

End Sub
Private Sub txtTemp_KeyPress(Index As Integer, KeyAscii As Integer)
Dim LeftPos As Long
Dim TopPos As Long
Dim AgCode As Long
Dim BrkAcCode As Long
TopPos = txttemp(Index).Top + Me.Top + 650
LeftPos = txttemp(Index).Left + Me.Left
If KeyAscii = 13 Then Exit Sub
Select Case Index
    Case 41  '-- Party   Alias Account
'        gClsSearch.SearchMultiField "QryMastAccount", "AcOurCode,Account,CGName,AcContPer,AcCity,AgName,AcBST,AcMsOmsParty", Array("Code", "Account Name", "Party Group", "Contact Person", "City", "Group", "TIN", "Ms/Oms"), Array(1000, 4000, 2000, 2000, 2000, 2500, 0, 0), "(Agcode in (90017) )", Chr(KeyAscii), "", LeftPos, TopPos, True
        gClsSearch.SearchMultiFieldAccMaster Chr(KeyAscii), txttemp(Index).Left + Me.Left, txttemp(Index).Top + Me.Top + 1000
        KeyAscii = 0
        If gClsSearch.SearchMultiRetColAccMaster(0) <> "" Then
            txttemp(Index) = gClsSearch.SearchMultiRetColAccMaster(0)
        End If
End Select
End Sub
Private Sub CalGridAmt()
Dim TmpAmt As Double
If OldTdsOn <> Val(msGrid.TextMatrix(msGrid.Row, 9)) And OldTdsOn <> 0 Then
    msGrid.TextMatrix(msGrid.Row, 10) = msGrid.TextMatrix(msGrid.Row, 9)
    OldTdsOn = Val(msGrid.TextMatrix(msGrid.Row, 10))
End If
If Val(msGrid.TextMatrix(msGrid.Row, 10)) > 0 Then
    TmpAmt = GProcMakeRounding((msGrid.TextMatrix(msGrid.Row, 10) * Val(txt(42))) / 100)
    If Abs(Val(msGrid.TextMatrix(msGrid.Row, 11)) - TmpAmt) > 1 Then
        msGrid.TextMatrix(msGrid.Row, 11) = TmpAmt
        msGrid.TextMatrix(msGrid.Row, 11) = Format(msGrid.TextMatrix(msGrid.Row, 11), FStr)
    End If
Else
    msGrid.TextMatrix(msGrid.Row, 11) = 0
    msGrid.TextMatrix(msGrid.Row, 11) = Format(msGrid.TextMatrix(msGrid.Row, 11), FStr)
End If

If Val(msGrid.TextMatrix(msGrid.Row, 9)) = 0 Then
    msGrid.TextMatrix(msGrid.Row, 10) = 0
    msGrid.TextMatrix(msGrid.Row, 10) = Format(msGrid.TextMatrix(msGrid.Row, 11), FStr)
    msGrid.TextMatrix(msGrid.Row, 11) = 0
    msGrid.TextMatrix(msGrid.Row, 11) = Format(msGrid.TextMatrix(msGrid.Row, 11), FStr)
End If

End Sub
