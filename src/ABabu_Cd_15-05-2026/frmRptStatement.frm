VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form frmRptStatement 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Reports"
   ClientHeight    =   8400
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   14040
   Icon            =   "frmRptStatement.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   8400
   ScaleWidth      =   14040
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
      Height          =   390
      Index           =   11
      Left            =   2400
      TabIndex        =   90
      Text            =   "11"
      Top             =   3780
      Width           =   4020
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
      Height          =   420
      Index           =   10
      Left            =   2400
      TabIndex        =   88
      Text            =   "10"
      Top             =   4260
      Visible         =   0   'False
      Width           =   1710
   End
   Begin VB.OptionButton OptAllBills 
      BackColor       =   &H00EDFDFE&
      Caption         =   "All Bills Statement"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   315
      Left            =   3480
      TabIndex        =   86
      Top             =   2280
      Visible         =   0   'False
      Width           =   2355
   End
   Begin VB.OptionButton opt2ndDay 
      BackColor       =   &H00EDFDFE&
      Caption         =   "2nd Day Pay Bills Statement"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   315
      Left            =   3480
      TabIndex        =   85
      Top             =   1860
      Visible         =   0   'False
      Width           =   2835
   End
   Begin VB.OptionButton optRegilur 
      BackColor       =   &H00EDFDFE&
      Caption         =   "Regular Bills Statement"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   315
      Left            =   3480
      TabIndex        =   84
      Top             =   1440
      Visible         =   0   'False
      Width           =   2355
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
      Height          =   390
      Index           =   9
      Left            =   2385
      TabIndex        =   13
      Text            =   "9"
      Top             =   6720
      Width           =   4020
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
      Height          =   390
      Index           =   8
      Left            =   2385
      TabIndex        =   12
      Text            =   "8"
      Top             =   6300
      Width           =   4020
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
      Index           =   7
      Left            =   2385
      TabIndex        =   11
      Text            =   "7"
      Top             =   5880
      Width           =   1800
   End
   Begin VB.CommandButton cmdCrystal 
      Caption         =   "Mail"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   480
      Index           =   3
      Left            =   3780
      Picture         =   "frmRptStatement.frx":058A
      TabIndex        =   80
      Top             =   7680
      Visible         =   0   'False
      Width           =   1185
   End
   Begin VB.CheckBox chkPaymentDirectToMill 
      BackColor       =   &H00EDFDFE&
      Caption         =   "Statement For Payment Direct To Mill"
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
      Left            =   1620
      TabIndex        =   79
      Top             =   1035
      Visible         =   0   'False
      Width           =   4695
   End
   Begin VB.Frame FrameApplyDbNt 
      BackColor       =   &H00FFFFC0&
      Height          =   8160
      Left            =   6615
      TabIndex        =   29
      Top             =   45
      Width           =   7260
      Begin VB.CommandButton cmdCrystal 
         Caption         =   "Genrate Bills ?"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   480
         Index           =   2
         Left            =   2655
         Style           =   1  'Graphical
         TabIndex        =   30
         Top             =   7470
         Width           =   1770
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Report                 Entry   "
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF00FF&
         Height          =   240
         Index           =   27
         Left            =   4680
         TabIndex        =   78
         Top             =   1170
         Width           =   2310
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Cash Discount Debit Note"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000080FF&
         Height          =   240
         Index           =   26
         Left            =   4365
         TabIndex        =   77
         Top             =   3915
         Width           =   2385
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Unloading Charges Debit Nt"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000080FF&
         Height          =   240
         Index           =   25
         Left            =   4455
         TabIndex        =   76
         Top             =   720
         Width           =   2625
      End
      Begin VB.Label Label20 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Cash Disc."
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
         Left            =   5790
         TabIndex        =   75
         Top             =   4950
         Width           =   1275
      End
      Begin VB.Label Label19 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Db. Nt. No"
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
         Left            =   5805
         TabIndex        =   74
         Top             =   4590
         Width           =   1275
      End
      Begin VB.Label Label14 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Cash Disc."
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
         Left            =   4350
         TabIndex        =   73
         Top             =   4950
         Width           =   1275
      End
      Begin VB.Line Line6 
         X1              =   4185
         X2              =   4185
         Y1              =   585
         Y2              =   7245
      End
      Begin VB.Label Label13 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Db. Nt. No"
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
         Left            =   5805
         TabIndex        =   72
         Top             =   1575
         Width           =   1275
      End
      Begin VB.Label Label12 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Unload Amt"
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
         Left            =   5790
         TabIndex        =   71
         Top             =   1935
         Width           =   1275
      End
      Begin VB.Label lblUnloadAmt 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Unload Amt"
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
         Left            =   4350
         TabIndex        =   70
         Top             =   1890
         Width           =   1275
      End
      Begin VB.Label lblBrkEntryDbNtNo 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Db. Nt. No"
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
         Left            =   2670
         TabIndex        =   69
         Top             =   4545
         Width           =   1275
      End
      Begin VB.Label lblBrkEntryBrkAmt 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Brok .Amt"
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
         Left            =   2670
         TabIndex        =   68
         Top             =   4905
         Width           =   1275
      End
      Begin VB.Label lblBrkEntryCgst 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Brok.Cgst"
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
         Left            =   2670
         TabIndex        =   67
         Top             =   5265
         Width           =   1275
      End
      Begin VB.Label lblBrkEntrySgst 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Brok.Sgst"
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
         Left            =   2670
         TabIndex        =   66
         Top             =   5670
         Width           =   1275
      End
      Begin VB.Label lblBrkEntryIgst 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Brok.Igst"
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
         Left            =   2670
         TabIndex        =   65
         Top             =   6075
         Width           =   1275
      End
      Begin VB.Label lblBrkEntryTotalAmt 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
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
         Height          =   240
         Left            =   2670
         TabIndex        =   64
         Top             =   6435
         Width           =   1275
      End
      Begin VB.Label lblBrkAmt 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Brok. Amt"
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
         Left            =   1170
         TabIndex        =   63
         Top             =   4950
         Width           =   1275
      End
      Begin VB.Label lblBrkCgst 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Brok.Cgst"
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
         Left            =   1170
         TabIndex        =   62
         Top             =   5310
         Width           =   1275
      End
      Begin VB.Label lblBrkSgst 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Brok.Sgst"
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
         Left            =   1170
         TabIndex        =   61
         Top             =   5715
         Width           =   1275
      End
      Begin VB.Label lblBrkIgst 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Brok.Igst"
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
         Left            =   1170
         TabIndex        =   60
         Top             =   6120
         Width           =   1275
      End
      Begin VB.Label lblBrkTotalAmt 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
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
         ForeColor       =   &H00000080&
         Height          =   240
         Left            =   1170
         TabIndex        =   59
         Top             =   6480
         Width           =   1275
      End
      Begin VB.Line Line5 
         BorderColor     =   &H000040C0&
         X1              =   0
         X2              =   7245
         Y1              =   7245
         Y2              =   7245
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Total      :"
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
         Index           =   24
         Left            =   180
         TabIndex        =   58
         Top             =   6480
         Width           =   900
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "IGST      :"
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
         Index           =   23
         Left            =   180
         TabIndex        =   57
         Top             =   6120
         Width           =   885
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "SGST     :"
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
         Index           =   22
         Left            =   180
         TabIndex        =   56
         Top             =   5760
         Width           =   900
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "CGST     :"
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
         Index           =   21
         Left            =   180
         TabIndex        =   55
         Top             =   5355
         Width           =   900
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Amount  :"
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
         Index           =   20
         Left            =   180
         TabIndex        =   54
         Top             =   4950
         Width           =   915
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Bill  No  :"
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
         Index           =   19
         Left            =   180
         TabIndex        =   53
         Top             =   4545
         Width           =   870
      End
      Begin VB.Label lblCmTotalAmt 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
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
         ForeColor       =   &H00000080&
         Height          =   240
         Left            =   1170
         TabIndex        =   52
         Top             =   3465
         Width           =   1275
      End
      Begin VB.Label lblCmEntryTotalAmt 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
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
         Height          =   240
         Left            =   2670
         TabIndex        =   51
         Top             =   3465
         Width           =   1275
      End
      Begin VB.Label lblCmEntryIgst 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Comm.Igst"
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
         Left            =   2670
         TabIndex        =   50
         Top             =   3105
         Width           =   1275
      End
      Begin VB.Label lblCmEntrySgst 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Comm.Sgst"
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
         Left            =   2670
         TabIndex        =   49
         Top             =   2700
         Width           =   1275
      End
      Begin VB.Label lblCmEntryCgst 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Comm.Cgst"
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
         Left            =   2670
         TabIndex        =   48
         Top             =   2295
         Width           =   1275
      End
      Begin VB.Label lblCmEntryCmAmt 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Comm.Amt"
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
         Left            =   2670
         TabIndex        =   47
         Top             =   1935
         Width           =   1275
      End
      Begin VB.Line Line4 
         BorderColor     =   &H000040C0&
         X1              =   0
         X2              =   7245
         Y1              =   1035
         Y2              =   1035
      End
      Begin VB.Label lblMainFirm 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Main Firm  "
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
         Left            =   1395
         TabIndex        =   46
         Top             =   270
         Width           =   3645
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Main Firm  :"
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
         Index           =   18
         Left            =   135
         TabIndex        =   45
         Top             =   270
         Width           =   1140
      End
      Begin VB.Label lblCmEntryDbNtNo 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Db. Nt. No"
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
         Left            =   2670
         TabIndex        =   44
         Top             =   1575
         Width           =   1275
      End
      Begin VB.Line Line3 
         BorderColor     =   &H000040C0&
         X1              =   0
         X2              =   7245
         Y1              =   3825
         Y2              =   3825
      End
      Begin VB.Label lblCmIgst 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Comm.Igst"
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
         Left            =   1170
         TabIndex        =   43
         Top             =   3105
         Width           =   1275
      End
      Begin VB.Label lblCmSgst 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Comm.Sgst"
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
         Left            =   1170
         TabIndex        =   42
         Top             =   2700
         Width           =   1275
      End
      Begin VB.Label lblCmCgst 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Comm.Cgst"
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
         Left            =   1170
         TabIndex        =   41
         Top             =   2295
         Width           =   1275
      End
      Begin VB.Label lblCmAmt 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Comm.Amt"
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
         Left            =   1170
         TabIndex        =   40
         Top             =   1935
         Width           =   1275
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Report                 Entry "
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF00FF&
         Height          =   240
         Index           =   17
         Left            =   1530
         TabIndex        =   39
         Top             =   1125
         Width           =   2190
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Bill  No  :"
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
         Index           =   16
         Left            =   180
         TabIndex        =   38
         Top             =   1530
         Width           =   870
      End
      Begin VB.Line Line2 
         BorderColor     =   &H000040C0&
         X1              =   0
         X2              =   7245
         Y1              =   4275
         Y2              =   4275
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Total      :"
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
         Index           =   15
         Left            =   180
         TabIndex        =   37
         Top             =   3465
         Width           =   900
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "IGST      :"
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
         Index           =   14
         Left            =   180
         TabIndex        =   36
         Top             =   3105
         Width           =   885
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "SGST     :"
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
         Index           =   13
         Left            =   180
         TabIndex        =   35
         Top             =   2700
         Width           =   900
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "CGST     :"
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
         Index           =   12
         Left            =   180
         TabIndex        =   34
         Top             =   2295
         Width           =   900
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Amount  :"
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
         Index           =   11
         Left            =   180
         TabIndex        =   33
         Top             =   1935
         Width           =   915
      End
      Begin VB.Line Line1 
         BorderColor     =   &H000040C0&
         X1              =   0
         X2              =   7245
         Y1              =   585
         Y2              =   585
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Brokerage  Bill Details"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000080FF&
         Height          =   240
         Index           =   10
         Left            =   180
         TabIndex        =   32
         Top             =   3915
         Width           =   2130
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Commission  Bill Details"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000080FF&
         Height          =   240
         Index           =   9
         Left            =   135
         TabIndex        =   31
         Top             =   720
         Width           =   2295
      End
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
      Height          =   390
      Index           =   6
      Left            =   2385
      TabIndex        =   9
      Text            =   "6"
      Top             =   4755
      Width           =   4020
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
      Index           =   5
      Left            =   2385
      TabIndex        =   10
      Text            =   "5"
      Top             =   5175
      Width           =   1800
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
      Height          =   420
      Index           =   4
      Left            =   3510
      MaxLength       =   60
      TabIndex        =   7
      Text            =   "4"
      Top             =   2640
      Width           =   1710
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
      Height          =   420
      Index           =   3
      Left            =   2385
      TabIndex        =   8
      Text            =   "3"
      Top             =   3315
      Width           =   1710
   End
   Begin VB.CommandButton cmdCrystal 
      Height          =   480
      Index           =   0
      Left            =   1185
      Picture         =   "frmRptStatement.frx":09CC
      Style           =   1  'Graphical
      TabIndex        =   20
      Top             =   7695
      Width           =   1185
   End
   Begin VB.CommandButton cmdCrystal 
      Height          =   480
      Index           =   1
      Left            =   2490
      Picture         =   "frmRptStatement.frx":0E0E
      Style           =   1  'Graphical
      TabIndex        =   21
      ToolTipText     =   "Print"
      Top             =   7695
      Width           =   1185
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
      Height          =   420
      Index           =   0
      Left            =   1635
      TabIndex        =   4
      Text            =   "0"
      Top             =   1395
      Width           =   1710
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
      Height          =   420
      Index           =   1
      Left            =   1635
      TabIndex        =   5
      Text            =   "1"
      Top             =   1800
      Width           =   1710
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
      Height          =   420
      Index           =   2
      Left            =   1635
      TabIndex        =   6
      Text            =   "2"
      Top             =   2205
      Width           =   1710
   End
   Begin VB.TextBox txtMillName 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   1590
      TabIndex        =   3
      Top             =   630
      Width           =   4815
   End
   Begin VB.CommandButton cmdExit 
      Cancel          =   -1  'True
      Caption         =   "E&xit"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   480
      Left            =   5055
      TabIndex        =   22
      Top             =   7695
      Width           =   1185
   End
   Begin Crystal.CrystalReport CrystalReport1 
      Left            =   270
      Top             =   7830
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
   Begin MSComctlLib.ProgressBar prgBar 
      Height          =   255
      Left            =   360
      TabIndex        =   23
      Top             =   7215
      Visible         =   0   'False
      Width           =   6180
      _ExtentX        =   10901
      _ExtentY        =   450
      _Version        =   393216
      Appearance      =   1
   End
   Begin MSComCtl2.DTPicker dtpToDt 
      Height          =   405
      Left            =   4650
      TabIndex        =   2
      Top             =   180
      Width           =   1755
      _ExtentX        =   3096
      _ExtentY        =   714
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
      Format          =   123600899
      CurrentDate     =   38050
   End
   Begin MSComCtl2.DTPicker dtpFrDt 
      Height          =   405
      Left            =   1635
      TabIndex        =   1
      Top             =   180
      Width           =   1845
      _ExtentX        =   3254
      _ExtentY        =   714
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
      Format          =   123600899
      CurrentDate     =   38050
   End
   Begin MSComCtl2.DTPicker dtpBillDt 
      Height          =   360
      Left            =   15435
      TabIndex        =   19
      Top             =   5580
      Visible         =   0   'False
      Width           =   1305
      _ExtentX        =   2302
      _ExtentY        =   635
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      CustomFormat    =   "dd/MM/yyyy"
      Format          =   123600899
      CurrentDate     =   38050
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Add Description      :"
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
      Index           =   32
      Left            =   540
      TabIndex        =   89
      Top             =   3840
      Visible         =   0   'False
      Width           =   1815
      WordWrap        =   -1  'True
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Add                       :"
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
      Index           =   31
      Left            =   540
      TabIndex        =   87
      Top             =   4320
      Visible         =   0   'False
      Width           =   1815
      WordWrap        =   -1  'True
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Description           :"
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
      Index           =   30
      Left            =   540
      TabIndex        =   83
      Top             =   6300
      Width           =   1860
      WordWrap        =   -1  'True
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Less Amount        :"
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
      Index           =   29
      Left            =   540
      TabIndex        =   82
      Top             =   5940
      Width           =   1695
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Less Amounts In Brokerage Statement   :"
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
      Index           =   28
      Left            =   540
      TabIndex        =   81
      Top             =   5580
      Width           =   6735
      WordWrap        =   -1  'True
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Unloading Chg. On Bags/Pkgs  : "
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
      Height          =   435
      Index           =   8
      Left            =   540
      TabIndex        =   28
      Top             =   2730
      Width           =   2910
      WordWrap        =   -1  'True
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Less Amounts In Commission Statement   :"
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
      Height          =   345
      Index           =   7
      Left            =   540
      TabIndex        =   27
      Top             =   3015
      Width           =   6735
      WordWrap        =   -1  'True
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Other  Rs.             :"
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
      Left            =   540
      TabIndex        =   26
      Top             =   5220
      Width           =   1740
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Other Description   :"
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
      Height          =   390
      Index           =   5
      Left            =   540
      TabIndex        =   25
      Top             =   4800
      Width           =   1860
      WordWrap        =   -1  'True
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Less  Freight          :"
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
      Left            =   540
      TabIndex        =   24
      Top             =   3360
      Width           =   2280
      WordWrap        =   -1  'True
   End
   Begin VB.Shape Shape2 
      FillColor       =   &H00BDDFD1&
      FillStyle       =   0  'Solid
      Height          =   780
      Left            =   960
      Shape           =   4  'Rounded Rectangle
      Top             =   7530
      Width           =   5520
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "CGST %  :"
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
      Left            =   540
      TabIndex        =   16
      Top             =   1440
      Width           =   945
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "IGST  %   :"
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
      Left            =   540
      TabIndex        =   17
      Top             =   2295
      Width           =   975
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "SGST %  :"
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
      Left            =   540
      TabIndex        =   18
      Top             =   1845
      Width           =   945
   End
   Begin VB.Label lblMillName 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "&Mill         :"
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
      Left            =   495
      TabIndex        =   15
      Top             =   660
      Width           =   900
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "&To          :"
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
      Left            =   3600
      TabIndex        =   14
      Top             =   262
      Width           =   900
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "From      : "
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
      Left            =   540
      TabIndex        =   0
      Top             =   262
      Width           =   975
   End
   Begin VB.Shape Shape6 
      BorderColor     =   &H00000080&
      FillColor       =   &H00EDFDFE&
      FillStyle       =   0  'Solid
      Height          =   7065
      Left            =   315
      Top             =   90
      Width           =   6165
   End
End
Attribute VB_Name = "frmRptStatement"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub cmdCrystal_Click(Index As Integer)
Dim FromDt As Date
Dim ToDt As Date
Dim X As String
Dim Y As String
Dim wMillCode As Integer
Dim wCommRt As Double
Dim wCommRt2 As Double
Dim wCommOn As Integer
Dim wCashDiscRt As Double
Dim wDepoConsi As String
Dim wMainFirmNm As String
Dim wMainFirmAd1 As String
Dim wMainFirmAd2 As String
Dim wMainFirmGSTIN As String
Dim tmpRst3 As Recordset
Dim wRCommOn As Double
Dim wRCommAmt As Double
Dim wRCommCgst As Double
Dim wRCommSgst As Double
Dim wRCommIgst As Double

Dim wRBrkAmt As Double
Dim wRBrkCgst As Double
Dim wRBrkSgst As Double
Dim wRBrkIgst As Double
Dim YN As Double
Dim tmpVno As Long
Dim Mon As Integer
Dim tmpBillNo As String
Dim BilSr As String
Dim PreVnoCM As Long
Dim PreVnoBrk As Long
Dim PreVnoJv As Long
Dim PreVnoJvUnloading As Long
Dim LessUnloading As Long
Dim wUnloadAmt As Double
Dim wBrokBilYN As Integer
Dim wEntSrNo As Integer
Dim wDepotHoAc As Long
Dim wFirmHoAc As Long
Dim wMillMailId As String
Dim MailCommInvNo As String
Dim MailBrokInvNo As String
Dim objCrystal As CRAXDRT.Application
Dim objReport As CRAXDRT.Report
Dim Tattach As String
Dim Mfile As String
Dim Mcp As String
Dim PtyMailId As String
Dim PtyMailIdCc As String
Dim PtyMailIdBcc As String
Dim MailSub As String
Dim MailStr As String
Dim j As Integer
Dim wMFirmPan As String
Dim w2ndDay As Integer
Dim DayStr As String
Dim DayStrBrk As String

'--for log
Dim wLogNo As Integer
Dim XX As String
Dim jj As Long
Dim tmpRstLog As Recordset
Dim AMD As String

If optRegilur.Value = True Then w2ndDay = 1
If opt2ndDay.Value = True Then w2ndDay = 2
If OptAllBills.Value = True Then w2ndDay = 0
                    
If optRegilur.Value = True Then DayStr = "Commission (Regular)"
If opt2ndDay.Value = True Then DayStr = "Commission (2nd Day)"
If OptAllBills.Value = True Then DayStr = "Commission"

If optRegilur.Value = True Then DayStrBrk = "Brokerage (Regular)"
If opt2ndDay.Value = True Then DayStrBrk = "Brokerage (2nd Day)"
If OptAllBills.Value = True Then DayStrBrk = "Brokerage"

FromDt = Format(dtpFrDt.Value, "dd/MM/yyyy")
ToDt = Format(dtpToDt.Value, "dd/MM/yyyy")
GProcCrystalRptPreparation CrystalReport1, IIf(Index = 0, "Window", "Printer")
With CrystalReport1
    Select Case gSelectedMenu
        Case "Consignment/ Depot Statement"  '--- "Consignment/ Depot Statement"
            Dim OpAmt As Double
            If ValidateData = False Then Exit Sub
                If txtMillName <> "" Then
                   wMillCode = GProcGetColumnValue("tblMastAccount", "AcName", txtMillName, "S", "AcCode", "N")
                   wDepotHoAc = GProcGetColumnValue("tblMastAccount", "AcName", txtMillName, "S", "AcDepotHoAc", "N")
                   wFirmHoAc = GProcGetColumnValue("tblMastAccount", "AcName", txtMillName, "S", "AcFirmHoAc", "N")
                Else
                   wMillCode = 0
                   wDepotHoAc = 0
                   wFirmHoAc = 0
                End If
                If opt2ndDay.Value = True Then
                   wCommRt = GProcGetColumnValue("tblMastAccount", "AcName", txtMillName, "S", "AcComRt2ndDay", "N")
                Else
                   wCommRt = GProcGetColumnValue("tblMastAccount", "AcName", txtMillName, "S", "AcComRt", "N")
                End If
                wCommOn = GProcGetColumnValue("tblMastAccount", "AcName", txtMillName, "S", "AcComOn", "N")
                wCashDiscRt = GProcGetColumnValue("tblMastAccount", "AcName", txtMillName, "S", "AcCashDiscRt", "N")
                wBrokBilYN = GProcGetColumnValue("tblMastAccount", "AcName", txtMillName, "S", "AcIsLocalBrkBySelf", "N")
                X = "select VType from tblsale where SlMillCode = " & wMillCode & " and vyear = " & gCYear & " and vfirm = '" & gCCode & "' and vtype in ('SD','SO')"
                wDepoConsi = ""
                If GProcRstOpen(tmpRst3, X, "R", gCn) > 0 Then
                   If tmpRst3.Fields(0) = "SO" Then
                      wDepoConsi = "Consignment"
                   Else
                      wDepoConsi = "Depot"
                   End If
                End If
                
                '-- Op Bal
                 gCmd.CommandText = "PrcPrepareOpBalance"
                 gCmd.Parameters.Refresh
                 gCmd.Parameters("@Ldr") = ""
                 gCmd.Parameters("@VFirm") = Trim(gCCode)
                 gCmd.Parameters("@To_dt") = Format(FromDt, "yyyy/mm/dd")
                 gCmd.Parameters("@FyStart_dt") = Format(gCYSDate, "yyyy/mm/dd")
                 gCmd.Parameters("@UCode") = gUserName
                 gCmd.Parameters("@CGSTAc") = gCgstAcCode
                 gCmd.Parameters("@SGSTAc") = gSgstAcCode
                 gCmd.Parameters("@IGSTAc") = gIgstAcCode
                 gCmd.Parameters("@TcsRec") = gTcsRec
                 gCmd.Parameters("@TcsPay") = gTcsPay
                 gCmd.Parameters("@CGSTRcmPayAc") = gCgstRCMAcCode
                 gCmd.Parameters("@SGSTRcmPayAc") = gSgstRCMAcCode
                 gCmd.Parameters("@IGSTRcmPayAc") = gIgstRCMAcCode
                 If wMillCode <> 0 Then gCmd.Parameters("@Accode") = wMillCode
                 gCmd.Execute
            
                
                 OpAmt = GProcGetColumnValue("TmpAccLedger", "AcName", txtMillName, "S", "VDrAmt-VCrAmt", "N")
                
                gCmd.CommandText = "PrcPrepareMillStatement"
                gCmd.Parameters.Refresh
                gCmd.Parameters("@Fr_dt") = Format(dtpFrDt, "yyyy/mm/dd")
                gCmd.Parameters("@To_dt") = Format(dtpToDt, "yyyy/mm/dd")
                gCmd.Parameters("@YearStart_Dt") = Format(gCYSDate, "yyyy/mm/dd")
                gCmd.Parameters("@UCode") = gUserName
                gCmd.Parameters("@VFirm") = Trim(gCCode)
                If wMillCode <> 0 Then gCmd.Parameters("@MillCode") = wMillCode
                gCmd.Parameters("@CommRt") = wCommRt
                gCmd.Parameters("@CommOn") = wCommOn
                gCmd.Parameters("@CashDiscRt") = wCashDiscRt
                gCmd.Parameters("@SrvTxRt") = Val(txt(0))
                gCmd.Parameters("@SWCessRt") = Val(txt(1))
                gCmd.Parameters("@KKCessRt") = Val(txt(2))
                gCmd.Parameters("@FreightPd") = Val(txt(3))
                gCmd.Parameters("@UnloadBags") = Val(txt(4))
                If optRegilur.Value = True Then gCmd.Parameters("@2ndDayBills") = 1
                If opt2ndDay.Value = True Then gCmd.Parameters("@2ndDayBills") = 2
                If OptAllBills.Value = True Then gCmd.Parameters("@2ndDayBills") = 3
                gCmd.Execute
                gCn.BeginTrans
                gCn.Execute "delete from TmpGentbl where UserName='" & gUserName & "' and nar6 = 'Op.Cal'  "
                gCn.Execute "delete from TmpGentbl where UserName='" & gUserName & "' and nar6 = 'H' and amt1+amt2+amt3 = 0"
                gCn.CommitTrans
                
                X = "{TmpGentbl.UserName}='" & gUserName & "' "
               .SelectionFormula = X
               If gCDepotMainFirm > 0 Then
                  wMainFirmNm = GProcGetColumnValue("tblMastAccount", "AcCode", "'" & gCDepotMainFirm & "' ", "N", "AcName", "S")
                  If Val(Right(wMainFirmNm, 1)) > 0 Then
                     wMainFirmNm = Left(wMainFirmNm, Len(wMainFirmNm) - 1)
                  End If
                  wMainFirmAd1 = GProcGetColumnValue("tblMastAccount", "AcCode", "'" & gCDepotMainFirm & "' ", "N", "AcAdd1", "S")
                  wMainFirmAd2 = GProcGetColumnValue("tblMastAccount", "AcCode", "'" & gCDepotMainFirm & "' ", "N", "AcAdd2", "S")
                  wMainFirmGSTIN = GProcGetColumnValue("tblMastAccount", "AcCode", "'" & gCDepotMainFirm & "' ", "N", "AcGSTIN", "S")
                  wMFirmPan = GProcGetColumnValue("tblMastAccount", "AcCode", "'" & gCDepotMainFirm & "' ", "N", "AcPAN", "S")
                  .Formulas(0) = "FirmNm = '" & UCase(wMainFirmNm) & "'"
                  .Formulas(1) = "wAdd1 = '" & wMainFirmAd1 & "'"
                  .Formulas(2) = "wAdd2 = '" & wMainFirmAd2 & "'"
                  .Formulas(16) = "wMainFirmGst = '" & wMainFirmGSTIN & "'"
                  .Formulas(7) = "Wpanno  = '" & wMFirmPan & "'"
               Else
                  .Formulas(16) = "wMainFirmGst = '" & gCGStin & "'"
                  .Formulas(7) = "Wpanno  = '" & gCPAN & "'"
               End If
               .Formulas(4) = "wFrToDt = 'From : ' & '" & dtpFrDt & "' & '  To ' & '" & dtpToDt & "'"
               .Formulas(5) = "wToDt = '" & dtpToDt & "'"
               .Formulas(6) = "STRegNo='" & gCSTRegNo & "'"
               .Formulas(8) = "wHamaliRt  = " & gCHamaliRt & ""
               .Formulas(9) = "wLevyRt  = " & gCLevyRt & ""
               .Formulas(10) = "wOpBal  = " & OpAmt & ""
               .Formulas(11) = "wOtherNote  = '" & txt(6) & "'"
               .Formulas(12) = "wOtherAmt  = " & txt(5) & ""
               .Formulas(13) = "wDepoConsi  = '" & wDepoConsi & "'"
               If cmdCrystal(2).Value = True Then
                  '-- For Commission
                  '--Unloading
                  X = "Select sum(amt2) from TmpGentbl where nar6 = 'B' AND amt6 = 0 and UserName = '" & gUserName & "'"
                  If GProcRstOpen(tmpRst3, X, "R", gCn) > 0 Then
                     wUnloadAmt = Round(IIf(IsNull(tmpRst3.Fields(0)), 0, tmpRst3.Fields(0) * gCHamaliRt), 2)
                     wUnloadAmt = wUnloadAmt + Round(((wUnloadAmt * gCLevyRt) / 100), 2)
                     wUnloadAmt = GProcMakeRounding(wUnloadAmt)
                  End If
                  '--Brokerage
                  If wBrokBilYN = 0 Then
                     X = "Select sum(iif(TmpGentbl.Nar1 = 'Bag',amt2*amt1,iif((TmpGentbl.Nar1 = 'Basic Amt' or TmpGentbl.Nar1 = 'Amt') ,round((amt12*amt1)/100,0),0))) from TmpGentbl where nar6 = 'A' and UserName = '" & gUserName & "'"
                     If GProcRstOpen(tmpRst3, X, "R", gCn) > 0 Then
                        wRBrkAmt = tmpRst3.Fields(0) - Val(txt(7))
                     End If
                  End If
                  '--Other Less = val(txt(5))
                  '--Freight Paid = val(txt(3))
                  X = "Select sum(amt1),sum(amt2),sum(amt8),sum(amt18) from TmpGentbl where nar6 = 'C' and Amt9 = 2 and UserName = '" & gUserName & "'"
                  If GProcRstOpen(tmpRst3, X, "R", gCn) > 0 Then
                     If gCDepotMainFirmCompCd = "" Or IsNull(gCDepotMainFirmCompCd) Then
                        MsgBox ("Please Check Main Firm In Company Master")
                        Exit Sub
                     End If
                     If gCommissionSaleAc = 0 Then
                        MsgBox ("Please Check Commission Sales A/c In Master Setting")
                        Exit Sub
                     End If
                     If gBrokerageSaleAc = 0 Then
                        MsgBox ("Please Check Brokerage Sales A/c In Master Setting")
                        Exit Sub
                     End If
                     LessUnloading = GProcGetColumnValue("tblMastAccount", "AcName", txtMillName, "S", "AcIsUnloadingLessInComm", "N")
                     
                     wRCommOn = tmpRst3.Fields(0) - (IIf(wBrokBilYN = 1, wRBrkAmt, 0) + IIf(LessUnloading = 1, wUnloadAmt, 0) + Val(txt(5)) + Val(txt(3))) - tmpRst3.Fields(1) - tmpRst3.Fields(3)
                     wRCommAmt = Round((wRCommOn * wCommRt) / 100, 2)
                     wRCommAmt = GProcMakeRounding(wRCommAmt)
                     lblCmAmt.Caption = Str(wRCommAmt)
                     wRCommCgst = Round((wRCommAmt * Val(txt(0))) / 100, 2)
                     wRCommCgst = GProcMakeRounding(wRCommCgst)
                     lblCmCgst.Caption = Str(wRCommCgst)
                     wRCommSgst = Round((wRCommAmt * Val(txt(1))) / 100, 2)
                     wRCommSgst = GProcMakeRounding(wRCommSgst)
                     lblCmSgst.Caption = Str(wRCommSgst)
                  
                     wRCommIgst = Round((wRCommAmt * Val(txt(2))) / 100, 2)
                     wRCommIgst = GProcMakeRounding(wRCommIgst)
                     lblCmIgst.Caption = Str(wRCommIgst)
                     lblCmTotalAmt.Caption = Str(wRCommAmt + wRCommCgst + wRCommSgst + wRCommIgst)
                                        
                     If wFirmHoAc > 0 Then
                        Y = "Select SlBillNo,tblSale.SlSubAmt,SlTaxAmt,SlTaxAmt2,SlTaxAmt3,SlBillAmt,ItName,tblSale.Vno from tblsale,tblSaleSub,tblMastItem where tblSale.VFirm = '" & gCDepotMainFirmCompCd & "' and SlAcDrCode = " & wFirmHoAc & " and tblSale.Vdt = ('" & Format(dtpToDt.Value, "yyyy/mm/dd") & "') " _
                            & " and tblSale.Vno = tblSaleSub.Vno and tblSale.Vtype = tblSaleSub.Vtype and tblSale.Vyear = tblSaleSub.Vyear and tblSale.Vfirm = tblSaleSub.Vfirm and tblSaleSub.SlSubItCode = tblMastItem.ItCode and left(tblMastItem.ItName,4) = 'Comm'    and tblSale.SlMillCode = " & wMillCode & "  and SlTmpChr1 = '" & DayStr & "'  "
                     Else
                        Y = "Select SlBillNo,tblSale.SlSubAmt,SlTaxAmt,SlTaxAmt2,SlTaxAmt3,SlBillAmt,ItName,tblSale.Vno from tblsale,tblSaleSub,tblMastItem where tblSale.VFirm = '" & gCDepotMainFirmCompCd & "' and SlAcDrCode = " & wMillCode & " and tblSale.Vdt = ('" & Format(dtpToDt.Value, "yyyy/mm/dd") & "') and SlTmpChr1 = '" & DayStr & "'  " _
                            & " and tblSale.Vno = tblSaleSub.Vno and tblSale.Vtype = tblSaleSub.Vtype and tblSale.Vyear = tblSaleSub.Vyear and tblSale.Vfirm = tblSaleSub.Vfirm and tblSaleSub.SlSubItCode = tblMastItem.ItCode and left(tblMastItem.ItName,4) = 'Comm'"
                     End If
                     
                     
                     If GProcRstOpen(tmpRst3, Y, "R", gCn) > 0 Then
                        lblCmEntryDbNtNo.Caption = tmpRst3.Fields(0)
                        lblCmEntryCmAmt.Caption = tmpRst3.Fields(1)
                        lblCmEntryCgst.Caption = tmpRst3.Fields(2)
                        lblCmEntrySgst.Caption = tmpRst3.Fields(3)
                        lblCmEntryIgst.Caption = tmpRst3.Fields(4)
                        lblCmEntryTotalAmt.Caption = tmpRst3.Fields(5)
                        PreVnoCM = tmpRst3.Fields(7)
                     End If
                  End If
                  '-- For Brokerage
                     
                     lblBrkAmt.Caption = Str(wRBrkAmt)
                     wRBrkCgst = Round((wRBrkAmt * Val(txt(0))) / 100, 2)
                     wRBrkCgst = GProcMakeRounding(wRBrkCgst)
                     lblBrkCgst.Caption = Str(wRBrkCgst)
                     wRBrkSgst = Round((wRBrkAmt * Val(txt(1))) / 100, 2)
                     wRBrkSgst = GProcMakeRounding(wRBrkSgst)
                     lblBrkSgst.Caption = Str(wRBrkSgst)

                     wRBrkIgst = Round((wRBrkAmt * Val(txt(2))) / 100, 2)
                     wRBrkIgst = GProcMakeRounding(wRBrkIgst)
                     lblBrkIgst.Caption = Str(wRBrkIgst)
                     lblBrkTotalAmt.Caption = Str(wRBrkAmt + wRBrkCgst + wRBrkSgst + wRBrkIgst)

                     If wFirmHoAc > 0 Then
                        Y = "Select SlBillNo,tblSale.SlSubAmt,SlTaxAmt,SlTaxAmt2,SlTaxAmt3,SlBillAmt,ItName,tblSale.Vno from tblsale,tblSaleSub,tblMastItem where tblSale.VFirm = '" & gCDepotMainFirmCompCd & "' and SlAcDrCode = " & wFirmHoAc & " and tblSale.Vdt = ('" & Format(dtpToDt.Value, "yyyy/mm/dd") & "') " _
                            & " and tblSale.Vno = tblSaleSub.Vno and tblSale.Vtype = tblSaleSub.Vtype and tblSale.Vyear = tblSaleSub.Vyear and tblSale.Vfirm = tblSaleSub.Vfirm and tblSaleSub.SlSubItCode = tblMastItem.ItCode and left(tblMastItem.ItName,4) = 'Brok'  and tblSale.SlMillCode = " & wMillCode & " and SlTmpChr1 = '" & DayStrBrk & "'  "
                     Else
                        Y = "Select SlBillNo,tblSale.SlSubAmt,SlTaxAmt,SlTaxAmt2,SlTaxAmt3,SlBillAmt,ItName,tblSale.Vno from tblsale,tblSaleSub,tblMastItem where tblSale.VFirm = '" & gCDepotMainFirmCompCd & "' and SlAcDrCode = " & wMillCode & " and tblSale.Vdt = ('" & Format(dtpToDt.Value, "yyyy/mm/dd") & "') " _
                            & " and tblSale.Vno = tblSaleSub.Vno and tblSale.Vtype = tblSaleSub.Vtype and tblSale.Vyear = tblSaleSub.Vyear and tblSale.Vfirm = tblSaleSub.Vfirm and tblSaleSub.SlSubItCode = tblMastItem.ItCode and left(tblMastItem.ItName,4) = 'Brok' and SlTmpChr1 = '" & DayStrBrk & "'  "
                     End If
                     If GProcRstOpen(tmpRst3, Y, "R", gCn) > 0 Then
                        lblBrkEntryDbNtNo.Caption = tmpRst3.Fields(0)
                        lblBrkEntryBrkAmt.Caption = tmpRst3.Fields(1)
                        lblBrkEntryCgst.Caption = tmpRst3.Fields(2)
                        lblBrkEntrySgst.Caption = tmpRst3.Fields(3)
                        lblBrkEntryIgst.Caption = tmpRst3.Fields(4)
                        lblBrkEntryTotalAmt.Caption = tmpRst3.Fields(5)
                        PreVnoBrk = tmpRst3.Fields(7)
                     
                     End If
                  If MsgBox("Genreate Bills ? Are you sure?", vbYesNo) = vbYes Then
                     '===== For Commission
                     
                     If PreVnoCM <> 0 Then
                        gCn.Execute "delete from tblSaleSub where vtype='SB' and vno=" & PreVnoCM & " and Vfirm = '" & gCDepotMainFirmCompCd & "'  and " _
                        & " VYear=" & gCYear
                        
                        gCn.Execute "delete from tblSale where vtype='SB' and vno=" & PreVnoCM & " and Vfirm = '" & gCDepotMainFirmCompCd & "'  and " _
                        & " VYear=" & gCYear
                        tmpBillNo = lblCmEntryDbNtNo.Caption  '--Inv No
                        Mon = Month(dtpToDt)
                        tmpVno = PreVnoCM
                     Else
                        tmpBillNo = GProcGenerateInvNoBrokComm("SB", "'" & gCDepotMainFirmCompCd & "'", gCDepotMainFirmCompCd) '--Inv No
                        Mon = Month(dtpToDt)
                        tmpVno = GProcGenerateIdMonthwise("TblSale", "Vno", Mon, "Vdt", "Vtype='SB' and VFirm = '" & gCDepotMainFirmCompCd & "'   and VYear=" & gCYear)
                     End If
                     
                     '---For Save In tblSale
                     
                     X = "select * from tblSale where vtype='SB' and tblSale.VFirm = '" & gCDepotMainFirmCompCd & "' and " _
                         & " VYear=" & gCYear
                     GProcRstOpen tmpRst3, X, "O"
                     If wRCommAmt > 0 Then
                        gCn.BeginTrans
                        tmpRst3.AddNew
                        tmpRst3!Vno = tmpVno '--Vno(0)
                        tmpRst3!Vtype = "SB" & vbNullString  '--Vtype(1)
                        tmpRst3!Vdt = dtpToDt.Value  '--Vdt(2)
                        tmpRst3!VYear = gCYear '--VYear
                        tmpRst3!VFirm = gCDepotMainFirmCompCd  '--VFirm
                        If wFirmHoAc > 0 Then
                            tmpRst3!SlAcDrCode = wFirmHoAc
                        Else
                            tmpRst3!SlAcDrCode = wMillCode
                        End If
                        tmpRst3!SlAcCrCode = gCommissionSaleAc
                        tmpRst3!SlMillCode = wMillCode
                        tmpRst3!SlBroker = GProcGetColumnValue("tblMastAccount", "AcName", "-", "S", "AcCode", "N", "")
                        tmpRst3!SlBillNo = tmpBillNo
                        tmpRst3!SlBillDt = dtpToDt.Value
                        tmpRst3!SlSubAmt = wRCommAmt
                        tmpRst3!SlTmpNo1 = wRCommOn
                        tmpRst3!SlTmpNo2 = wCommRt
                        tmpRst3!SlBillAmt = wRCommAmt + wRCommCgst + wRCommSgst + wRCommIgst
                        tmpRst3!SlTaxableAmt = wRCommAmt
                        tmpRst3!SlTaxRate = Val(txt(0))
                        tmpRst3!SlTaxAmt = wRCommCgst
                        tmpRst3!SlTaxRate2 = Val(txt(1))
                        tmpRst3!SlTaxAmt2 = wRCommSgst
                        tmpRst3!SlTaxRate3 = Val(txt(2))
                        tmpRst3!SlTaxAmt3 = wRCommIgst
                        tmpRst3!SlTransport = GProcGetColumnValue("tblMastNarration", "Narration", "-", "S", "NarrCode", "N")
                        tmpRst3!SlTaxCode = GProcGetColumnValue("tblMastNarration", "Narration", "-", "S", "NarrCode", "N")
                        tmpRst3!DueDays = 0
                        tmpRst3!DueDate = dtpToDt.Value
                        If optRegilur.Value = True Then tmpRst3!SlTmpChr1 = "Commission (Regular)"
                        If opt2ndDay.Value = True Then tmpRst3!SlTmpChr1 = "Commission (2nd Day)"
                        If OptAllBills.Value = True Then tmpRst3!SlTmpChr1 = "Commission"
                        tmpRst3!SlTmpChr2 = "For The Month Of " + Format(dtpToDt, "MMM-YY")
                        tmpRst3!SlTmpChr3 = "[ Including G S T ]"
                        If optRegilur.Value = True Then tmpRst3!SlTmpNo3 = 1
                        If opt2ndDay.Value = True Then tmpRst3!SlTmpNo3 = 2
                        tmpRst3.Update
                        gCn.CommitTrans
                        
                        '---For Save In tblSaleSub
                        
                        X = "select * from tblSaleSub where vtype='SB' and tblSaleSub.VFirm = '" & gCDepotMainFirmCompCd & "' and " _
                            & " VYear=" & gCYear
                        GProcRstOpen tmpRst3, X, "O"
                        
                        tmpRst3.AddNew
                        tmpRst3!Vno = tmpVno '--Vno(0)
                        tmpRst3!Vtype = "SB" & vbNullString  '--Vtype(1)
                        tmpRst3!Vdt = dtpToDt.Value  '--Vdt(2)
                        tmpRst3!VYear = gCYear '--VYear
                        tmpRst3!VFirm = gCDepotMainFirmCompCd '--VFirm
                        tmpRst3!SlSubItSrNo = 1 '--Sr No
                        '--Itcode(4)
                        tmpRst3!SlSubItCode = gCommissionItem '--Item Code
                        tmpRst3!SlSubAmt = wRCommAmt
                        tmpRst3!SlSubGodown = GProcGetColumnValue("tblMastNarration", "Narration", "-", "S", "NarrCode", "N", "")
                        '-- P It Ctrl No =Booking It ctrl No
                        tmpRst3.Update
                        
                        '--For Log Add & Modi
                        If gCYear >= 2023 And gCIsLog = 1 Then
                           If PreVnoCM = 0 Then
                              AMD = "A"
                           Else
                              AMD = "M"
                           End If
                           gCn.BeginTrans
                           XX = "Select max(LogNo) from tblSale_Log where vtype='SB' and vno=" & tmpVno & " and Vfirm = '" & gCDepotMainFirmCompCd & "'  and" _
                                & " VYear=" & gCYear

                           jj = GProcRstOpen(tmpRstLog, XX, "R")
                           If jj > 0 Then
                              wLogNo = IIf(IsNull(tmpRstLog.Fields(0)), 0, tmpRstLog.Fields(0)) + 1
                           Else
                              wLogNo = 1
                           End If
                           gCn.Execute "Insert into tblSale_Log select tblSale.*," & wLogNo & " as LogNo,'" & AMD & "' as LogTp ,'" & gUserName & "' as UserName ,'" & Format(Now, "YYYY/MM/DD hh:mm:ss") & "' as LogDate from tblSale where vtype='SB' and vno=" & tmpVno & " and Vfirm = '" & gCDepotMainFirmCompCd & "'  and" _
                                       & " VYear=" & gCYear
                           
                           gCn.Execute "Insert into tblSaleSub_Log select tblSaleSub.*," & wLogNo & " as LogNo,'" & AMD & "' as LogTp ,'" & gUserName & "' as UserName ,'" & Format(Now, "YYYY/MM/DD hh:mm:ss") & "' as LogDate from tblSaleSub where vtype='SB' and vno=" & tmpVno & " and Vfirm = '" & gCDepotMainFirmCompCd & "'  and" _
                                       & " VYear=" & gCYear
                           
                           gCn.CommitTrans
                           tmpRstLog.Close
                        End If
                     
                     
                     End If
                     '-- end commission
                     '===== For Brokerage
                     If PreVnoBrk <> 0 Then
                        gCn.Execute "delete from tblSaleSub where vtype='SB' and vno=" & PreVnoBrk & " and Vfirm = '" & gCDepotMainFirmCompCd & "'  and " _
                        & " VYear=" & gCYear
                        
                        gCn.Execute "delete from tblSale where vtype='SB' and vno=" & PreVnoBrk & " and Vfirm = '" & gCDepotMainFirmCompCd & "'  and " _
                        & " VYear=" & gCYear
                        tmpBillNo = lblBrkEntryDbNtNo.Caption  '--Inv No
                        Mon = Month(dtpToDt)
                        tmpVno = PreVnoBrk
                     Else
                        tmpBillNo = GProcGenerateInvNoBrokComm("SB", "'" & gCDepotMainFirmCompCd & "'", gCDepotMainFirmCompCd) '--Inv No
                        Mon = Month(dtpToDt)
                        tmpVno = GProcGenerateIdMonthwise("TblSale", "Vno", Mon, "Vdt", "Vtype='SB' and VFirm = '" & gCDepotMainFirmCompCd & "'   and VYear=" & gCYear)
                     End If
                     
                     '---For Save In tblSale
                     
                     X = "select * from tblSale where vtype='SB' and tblSale.VFirm = '" & gCDepotMainFirmCompCd & "' and " _
                         & " VYear=" & gCYear
                     GProcRstOpen tmpRst3, X, "O"
                     If wRBrkAmt > 0 Then
                        tmpRst3.AddNew
                        tmpRst3!Vno = tmpVno '--Vno(0)
                        tmpRst3!Vtype = "SB" & vbNullString  '--Vtype(1)
                        tmpRst3!Vdt = dtpToDt.Value  '--Vdt(2)
                        tmpRst3!VYear = gCYear '--VYear
                        tmpRst3!VFirm = gCDepotMainFirmCompCd  '--VFirm
                        If wFirmHoAc > 0 Then
                            tmpRst3!SlAcDrCode = wFirmHoAc
                        Else
                            tmpRst3!SlAcDrCode = wMillCode
                        End If
                        tmpRst3!SlAcCrCode = gBrokerageSaleAc
                        tmpRst3!SlMillCode = wMillCode
                        tmpRst3!SlBroker = GProcGetColumnValue("tblMastAccount", "AcName", "-", "S", "AcCode", "N", "")
                        tmpRst3!SlBillNo = tmpBillNo
                        tmpRst3!SlBillDt = dtpToDt.Value
                        tmpRst3!SlSubAmt = wRBrkAmt
                        tmpRst3!SlBillAmt = wRBrkAmt + wRBrkCgst + wRBrkSgst + wRBrkIgst
                        tmpRst3!SlTaxableAmt = wRBrkAmt
                        tmpRst3!SlTaxRate = Val(txt(0))
                        tmpRst3!SlTaxAmt = wRBrkCgst
                        tmpRst3!SlTaxRate2 = Val(txt(1))
                        tmpRst3!SlTaxAmt2 = wRBrkSgst
                        tmpRst3!SlTaxRate3 = Val(txt(2))
                        tmpRst3!SlTaxAmt3 = wRBrkIgst
                        tmpRst3!SlTransport = GProcGetColumnValue("tblMastNarration", "Narration", "-", "S", "NarrCode", "N")
                        tmpRst3!SlTaxCode = GProcGetColumnValue("tblMastNarration", "Narration", "-", "S", "NarrCode", "N")
                        tmpRst3!DueDays = 0
                        tmpRst3!DueDate = dtpToDt.Value
                        If optRegilur.Value = True Then tmpRst3!SlTmpChr1 = "Brokerage (Regular)"
                        If opt2ndDay.Value = True Then tmpRst3!SlTmpChr1 = "Brokerage (2nd Day)"
                        If OptAllBills.Value = True Then tmpRst3!SlTmpChr1 = "Brokerage"
                        tmpRst3!SlTmpChr2 = "For  The Month  Of " + Format(dtpToDt, "MMM-YY")
                        tmpRst3!SlTmpChr3 = "[Including G S T ]"
                        If optRegilur.Value = True Then tmpRst3!SlTmpNo3 = 1
                        If opt2ndDay.Value = True Then tmpRst3!SlTmpNo3 = 2
                        tmpRst3.Update
                                              
                        '---For Save In tblSaleSub
                        
                        X = "select * from tblSaleSub where vtype='SB' and tblSaleSub.VFirm = '" & gCDepotMainFirmCompCd & "' and " _
                            & " VYear=" & gCYear
                        GProcRstOpen tmpRst3, X, "O"
                        
                        tmpRst3.AddNew
                        tmpRst3!Vno = tmpVno '--Vno(0)
                        tmpRst3!Vtype = "SB" & vbNullString  '--Vtype(1)
                        tmpRst3!Vdt = dtpToDt.Value  '--Vdt(2)
                        tmpRst3!VYear = gCYear '--VYear
                        tmpRst3!VFirm = gCDepotMainFirmCompCd '--VFirm
                        tmpRst3!SlSubItSrNo = 1 '--Sr No
                        '--Itcode(4)
                        tmpRst3!SlSubItCode = gBrokerageItem  '--Item Code
                        tmpRst3!SlSubAmt = wRBrkAmt
                        tmpRst3!SlSubGodown = GProcGetColumnValue("tblMastNarration", "Narration", "-", "S", "NarrCode", "N", "")
                        '-- P It Ctrl No =Booking It ctrl No
                        tmpRst3.Update
                     
                        '--For Log Add & Modi
                        If gCYear >= 2023 And gCIsLog = 1 Then
                           If PreVnoBrk = 0 Then
                              AMD = "A"
                           Else
                              AMD = "M"
                           End If
                           gCn.BeginTrans
                           XX = "Select max(LogNo) from tblSale_Log where vtype='SB' and vno=" & tmpVno & " and Vfirm = '" & gCDepotMainFirmCompCd & "'  and" _
                                & " VYear=" & gCYear

                           jj = GProcRstOpen(tmpRstLog, XX, "R")
                           If jj > 0 Then
                              wLogNo = IIf(IsNull(tmpRstLog.Fields(0)), 0, tmpRstLog.Fields(0)) + 1
                           Else
                              wLogNo = 1
                           End If
                           gCn.Execute "Insert into tblSale_Log select tblSale.*," & wLogNo & " as LogNo,'" & AMD & "' as LogTp ,'" & gUserName & "' as UserName ,'" & Format(Now, "YYYY/MM/DD hh:mm:ss") & "' as LogDate from tblSale where vtype='SB' and vno=" & tmpVno & " and Vfirm = '" & gCDepotMainFirmCompCd & "'  and" _
                                       & " VYear=" & gCYear
                           
                           gCn.Execute "Insert into tblSaleSub_Log select tblSaleSub.*," & wLogNo & " as LogNo,'" & AMD & "' as LogTp ,'" & gUserName & "' as UserName ,'" & Format(Now, "YYYY/MM/DD hh:mm:ss") & "' as LogDate from tblSaleSub where vtype='SB' and vno=" & tmpVno & " and Vfirm = '" & gCDepotMainFirmCompCd & "'  and" _
                                       & " VYear=" & gCYear
                           
                           gCn.CommitTrans
                           tmpRstLog.Close
                        End If
                     
                     End If
                     '-- For Unloading
                     If wFirmHoAc > 0 Then
                         Y = "Select Vno from tblVoucher where tblVoucher.VFirm = '" & gCDepotMainFirmCompCd & "' and VDrAcCode = " & wFirmHoAc & " and VCrAcCode = " & gUnloadingChg & " and tblVoucher.Vdt = ('" & Format(dtpToDt.Value, "yyyy/mm/dd") & "')  and VBillVno  = " & wMillCode & " and LateDays = " & w2ndDay & "   "
                     Else
                         Y = "Select Vno from tblVoucher where tblVoucher.VFirm = '" & gCDepotMainFirmCompCd & "' and VDrAcCode = " & wMillCode & " and VCrAcCode = " & gUnloadingChg & " and tblVoucher.Vdt = ('" & Format(dtpToDt.Value, "yyyy/mm/dd") & "') and LateDays = " & w2ndDay & "  "
                     End If
                     If GProcRstOpen(tmpRst3, Y, "R", gCn) > 0 Then
                        PreVnoJvUnloading = tmpRst3.Fields(0)
                     End If
                     If PreVnoJvUnloading <> 0 Then
                        gCn.Execute "delete from tblVoucher where vtype='JV' and vno=" & PreVnoJvUnloading & " and Vfirm = '" & gCDepotMainFirmCompCd & "'  and " _
                        & " VYear=" & gCYear
                        Mon = Month(dtpToDt)
                        tmpVno = PreVnoJvUnloading
                     Else
                        Mon = Month(dtpToDt)
                        tmpVno = GProcGenerateIdMonthwise("TblVoucher", "Vno", Mon, "Vdt", "Vtype='JV' and VFirm = '" & gCDepotMainFirmCompCd & "'   and VYear=" & gCYear)
                     End If
                     X = "select * from tblVoucher where vtype='JV' and tblVoucher.VFirm = '" & gCDepotMainFirmCompCd & "' and " _
                         & " VYear=" & gCYear
                     GProcRstOpen tmpRst3, X, "O"
                     wEntSrNo = 1
                     If wUnloadAmt > 0 Then
                        tmpRst3.AddNew
                        tmpRst3!Vno = tmpVno '--Vno(0)
                        tmpRst3!Vtype = "JV" & vbNullString  '--Vtype(1)
                        tmpRst3!Vdt = dtpToDt.Value  '--Vdt(2)
                        tmpRst3!VYear = gCYear '--VYear
                        tmpRst3!VFirm = gCDepotMainFirmCompCd  '--VFirm
                        tmpRst3!VCtrNo = wEntSrNo
                        tmpRst3!VCrAcCode = gUnloadingChg
                        If wFirmHoAc > 0 Then
                            tmpRst3!VDrAcCode = wFirmHoAc
                        Else
                            tmpRst3!VDrAcCode = wMillCode
                        End If
                        tmpRst3!VBillVno = wMillCode
                        tmpRst3!vamt = wUnloadAmt
                        tmpRst3!VNar1 = "Unloading Chg. For - " + CStr(Format(dtpToDt.Value, "mmm-yyyy"))
                        tmpRst3!VTopCrDr = 1
                        If optRegilur.Value = True Then tmpRst3!LateDays = 1
                        If opt2ndDay.Value = True Then tmpRst3!LateDays = 2
                        tmpRst3.Update
                        
                        '--For Log Add & Modi
                        If gCYear >= 2023 And gCIsLog = 1 Then
                           If PreVnoJvUnloading = 0 Then
                              AMD = "A"
                           Else
                              AMD = "M"
                           End If
                           gCn.BeginTrans
                           XX = "Select max(LogNo) from tblVoucher_Log where vtype='JV' and vno=" & tmpVno & " and Vfirm = '" & gCDepotMainFirmCompCd & "'  and" _
                                & " VYear=" & gCYear

                           jj = GProcRstOpen(tmpRstLog, XX, "R")
                           If jj > 0 Then
                              wLogNo = IIf(IsNull(tmpRstLog.Fields(0)), 0, tmpRstLog.Fields(0)) + 1
                           Else
                              wLogNo = 1
                           End If
                           gCn.Execute "Insert into tblVoucher_Log select tblVoucher.*," & wLogNo & " as LogNo,'" & AMD & "' as LogTp ,'" & gUserName & "' as UserName ,'" & Format(Now, "YYYY/MM/DD hh:mm:ss") & "' as LogDate from tblVoucher where vtype='JV' and vno=" & tmpVno & " and Vfirm = '" & gCDepotMainFirmCompCd & "'  and" _
                                       & " VYear=" & gCYear
                           
                           gCn.CommitTrans
                           tmpRstLog.Close
                        End If
                     
                     End If
                     '--end unloading
                     
                     '--- For JV In Depot
                     
                     If wDepotHoAc > 0 Then
                        Y = "Select Vno from tblVoucher where tblVoucher.VFirm = '" & gCCode & "' and VDrAcCode = " & wDepotHoAc & " and VCrAcCode = " & gCDepotMainFirm & " and tblVoucher.Vdt = ('" & Format(dtpToDt.Value, "yyyy/mm/dd") & "')  and VBillVno  = " & wMillCode & " and LateDays = " & w2ndDay & " "
                     Else
                        Y = "Select Vno from tblVoucher where tblVoucher.VFirm = '" & gCCode & "' and VDrAcCode = " & wMillCode & " and VCrAcCode = " & gCDepotMainFirm & " and tblVoucher.Vdt = ('" & Format(dtpToDt.Value, "yyyy/mm/dd") & "') and LateDays = " & w2ndDay & " "
                     End If
                     If GProcRstOpen(tmpRst3, Y, "R", gCn) > 0 Then
                        PreVnoJv = tmpRst3.Fields(0)
                     End If
                     If PreVnoJv <> 0 Then
                        gCn.Execute "delete from tblVoucher where vtype='JV' and vno=" & PreVnoJv & " and Vfirm = '" & gCCode & "'  and " _
                        & " VYear=" & gCYear
                        Mon = Month(dtpToDt)
                        tmpVno = PreVnoJv
                     Else
                        Mon = Month(dtpToDt)
                        tmpVno = GProcGenerateIdMonthwise("TblVoucher", "Vno", Mon, "Vdt", "Vtype='JV' and VFirm = '" & gCCode & "'   and VYear=" & gCYear)
                     End If
                     
                     X = "select * from tblVoucher where vtype='JV' and tblVoucher.VFirm = '" & gCCode & "' and " _
                         & " VYear=" & gCYear
                     GProcRstOpen tmpRst3, X, "O"
                     wEntSrNo = 1
                     If wRCommAmt > 0 Then
                        tmpRst3.AddNew
                        tmpRst3!Vno = tmpVno '--Vno(0)
                        tmpRst3!Vtype = "JV" & vbNullString  '--Vtype(1)
                        tmpRst3!Vdt = dtpToDt.Value  '--Vdt(2)
                        tmpRst3!VYear = gCYear '--VYear
                        tmpRst3!VFirm = gCCode  '--VFirm
                        tmpRst3!VCtrNo = wEntSrNo
                        tmpRst3!VCrAcCode = gCDepotMainFirm
                        If wDepotHoAc > 0 Then
                           tmpRst3!VDrAcCode = wDepotHoAc
                        Else
                           tmpRst3!VDrAcCode = wMillCode
                        End If
                        tmpRst3!VBillVno = wMillCode
                        tmpRst3!vamt = wRCommAmt
                        tmpRst3!VBillAmt = wCommOn
                        tmpRst3!IntRt = wCommRt
                        If optRegilur.Value = True Then tmpRst3!VNar1 = "Commission ( Regular ) For - " + CStr(Format(dtpToDt.Value, "mmm-yyyy"))
                        If opt2ndDay.Value = True Then tmpRst3!VNar1 = "Commission ( 2nd Day ) For - " + CStr(Format(dtpToDt.Value, "mmm-yyyy"))
                        If OptAllBills.Value = True Then tmpRst3!VNar1 = "Commission For - " + CStr(Format(dtpToDt.Value, "mmm-yyyy"))
'                        tmpRst3!VNar2 = "[ Including G S T ]"
                        tmpRst3!VTopCrDr = 1
                        If optRegilur.Value = True Then tmpRst3!LateDays = 1
                        If opt2ndDay.Value = True Then tmpRst3!LateDays = 2
                        tmpRst3.Update
                        If wRCommCgst > 0 Then
                           wEntSrNo = wEntSrNo + 1
                           tmpRst3.AddNew
                           tmpRst3!Vno = tmpVno '--Vno(0)
                           tmpRst3!Vtype = "JV" & vbNullString  '--Vtype(1)
                           tmpRst3!Vdt = dtpToDt.Value  '--Vdt(2)
                           tmpRst3!VYear = gCYear '--VYear
                           tmpRst3!VFirm = gCCode  '--VFirm
                           tmpRst3!VCtrNo = wEntSrNo
                           tmpRst3!VCrAcCode = gCDepotMainFirm
                           If wDepotHoAc > 0 Then
                               tmpRst3!VDrAcCode = wDepotHoAc
                           Else
                               tmpRst3!VDrAcCode = wMillCode
                           End If
                           tmpRst3!vamt = wRCommCgst
                           tmpRst3!VNar1 = "C Gst On Commission For - " + CStr(Format(dtpToDt.Value, "mmm-yyyy"))
                           tmpRst3!VTopCrDr = 1
                           tmpRst3!VBillVno = wMillCode
                           If optRegilur.Value = True Then tmpRst3!LateDays = 1
                           If opt2ndDay.Value = True Then tmpRst3!LateDays = 2
                           tmpRst3.Update
                        End If
                        If wRCommSgst > 0 Then
                           wEntSrNo = wEntSrNo + 1
                           tmpRst3.AddNew
                           tmpRst3!Vno = tmpVno '--Vno(0)
                           tmpRst3!Vtype = "JV" & vbNullString  '--Vtype(1)
                           tmpRst3!Vdt = dtpToDt.Value  '--Vdt(2)
                           tmpRst3!VYear = gCYear '--VYear
                           tmpRst3!VFirm = gCCode  '--VFirm
                           tmpRst3!VCtrNo = wEntSrNo
                           tmpRst3!VCrAcCode = gCDepotMainFirm
                           If wDepotHoAc > 0 Then
                               tmpRst3!VDrAcCode = wDepotHoAc
                           Else
                               tmpRst3!VDrAcCode = wMillCode
                           End If
                           tmpRst3!vamt = wRCommSgst
                           tmpRst3!VNar1 = "S Gst On Commission For - " + CStr(Format(dtpToDt.Value, "mmm-yyyy"))
                           tmpRst3!VTopCrDr = 1
                           tmpRst3!VBillVno = wMillCode
                           If optRegilur.Value = True Then tmpRst3!LateDays = 1
                           If opt2ndDay.Value = True Then tmpRst3!LateDays = 2
                           tmpRst3.Update
                        End If
                        If wRCommIgst > 0 Then
                           wEntSrNo = wEntSrNo + 1
                           tmpRst3.AddNew
                           tmpRst3!Vno = tmpVno '--Vno(0)
                           tmpRst3!Vtype = "JV" & vbNullString  '--Vtype(1)
                           tmpRst3!Vdt = dtpToDt.Value  '--Vdt(2)
                           tmpRst3!VYear = gCYear '--VYear
                           tmpRst3!VFirm = gCCode  '--VFirm
                           tmpRst3!VCtrNo = wEntSrNo
                           tmpRst3!VCrAcCode = gCDepotMainFirm
                           If wDepotHoAc > 0 Then
                              tmpRst3!VDrAcCode = wDepotHoAc
                           Else
                              tmpRst3!VDrAcCode = wMillCode
                           End If
                           tmpRst3!vamt = wRCommIgst
                           tmpRst3!VNar1 = "I Gst On Commission For - " + CStr(Format(dtpToDt.Value, "mmm-yyyy"))
                           tmpRst3!VTopCrDr = 1
                           tmpRst3!VBillVno = wMillCode
                           If optRegilur.Value = True Then tmpRst3!LateDays = 1
                           If opt2ndDay.Value = True Then tmpRst3!LateDays = 2
                           tmpRst3.Update
                        End If
                     End If
                     If wRBrkAmt > 0 Then
                        wEntSrNo = wEntSrNo + 1
                        tmpRst3.AddNew
                        tmpRst3!Vno = tmpVno '--Vno(0)
                        tmpRst3!Vtype = "JV" & vbNullString  '--Vtype(1)
                        tmpRst3!Vdt = dtpToDt.Value  '--Vdt(2)
                        tmpRst3!VYear = gCYear '--VYear
                        tmpRst3!VFirm = gCCode  '--VFirm
                        tmpRst3!VCtrNo = wEntSrNo
                        tmpRst3!VCrAcCode = gCDepotMainFirm
                        If wDepotHoAc > 0 Then
                           tmpRst3!VDrAcCode = wDepotHoAc
                        Else
                           tmpRst3!VDrAcCode = wMillCode
                        End If
                        tmpRst3!vamt = wRBrkAmt '-- wRBrkAmt + wRBrkCgst + wRBrkSgst + wRBrkIgst
                        If optRegilur.Value = True Then tmpRst3!VNar1 = "Brokerage Paid ( Regular ) For - " + CStr(Format(dtpToDt.Value, "mmm-yyyy"))
                        If opt2ndDay.Value = True Then tmpRst3!VNar1 = "Brokerage Paid ( 2nd Day ) For - " + CStr(Format(dtpToDt.Value, "mmm-yyyy"))
                        If OptAllBills.Value = True Then tmpRst3!VNar1 = "Brokerage Paid For - " + CStr(Format(dtpToDt.Value, "mmm-yyyy"))
 '                       tmpRst3!VNar2 = "[ Including G S T ]"
                        tmpRst3!VTopCrDr = 1
                        tmpRst3!VBillVno = wMillCode
                        If optRegilur.Value = True Then tmpRst3!LateDays = 1
                        If opt2ndDay.Value = True Then tmpRst3!LateDays = 2
                        tmpRst3.Update
                        If wRBrkCgst > 0 Then
                           wEntSrNo = wEntSrNo + 1
                           tmpRst3.AddNew
                           tmpRst3!Vno = tmpVno '--Vno(0)
                           tmpRst3!Vtype = "JV" & vbNullString  '--Vtype(1)
                           tmpRst3!Vdt = dtpToDt.Value  '--Vdt(2)
                           tmpRst3!VYear = gCYear '--VYear
                           tmpRst3!VFirm = gCCode  '--VFirm
                           tmpRst3!VCtrNo = wEntSrNo
                           tmpRst3!VCrAcCode = gCDepotMainFirm
                           If wDepotHoAc > 0 Then
                              tmpRst3!VDrAcCode = wDepotHoAc
                           Else
                              tmpRst3!VDrAcCode = wMillCode
                           End If
                           tmpRst3!vamt = wRBrkCgst
                           tmpRst3!VNar1 = "C Gst On Brokerage For - " + CStr(Format(dtpToDt.Value, "mmm-yyyy"))
                           tmpRst3!VTopCrDr = 1
                           tmpRst3!VBillVno = wMillCode
                           If optRegilur.Value = True Then tmpRst3!LateDays = 1
                           If opt2ndDay.Value = True Then tmpRst3!LateDays = 2
                           tmpRst3.Update
                        End If
                        If wRBrkSgst > 0 Then
                           wEntSrNo = wEntSrNo + 1
                           tmpRst3.AddNew
                           tmpRst3!Vno = tmpVno '--Vno(0)
                           tmpRst3!Vtype = "JV" & vbNullString  '--Vtype(1)
                           tmpRst3!Vdt = dtpToDt.Value  '--Vdt(2)
                           tmpRst3!VYear = gCYear '--VYear
                           tmpRst3!VFirm = gCCode  '--VFirm
                           tmpRst3!VCtrNo = wEntSrNo
                           tmpRst3!VCrAcCode = gCDepotMainFirm
                           If wDepotHoAc > 0 Then
                              tmpRst3!VDrAcCode = wDepotHoAc
                           Else
                              tmpRst3!VDrAcCode = wMillCode
                           End If
                           tmpRst3!vamt = wRBrkSgst
                           tmpRst3!VNar1 = "S Gst On Brokerage For - " + CStr(Format(dtpToDt.Value, "mmm-yyyy"))
                           tmpRst3!VTopCrDr = 1
                           tmpRst3!VBillVno = wMillCode
                           If optRegilur.Value = True Then tmpRst3!LateDays = 1
                           If opt2ndDay.Value = True Then tmpRst3!LateDays = 2
                           tmpRst3.Update
                        End If
                        If wRBrkIgst > 0 Then
                           wEntSrNo = wEntSrNo + 1
                           tmpRst3.AddNew
                           tmpRst3!Vno = tmpVno '--Vno(0)
                           tmpRst3!Vtype = "JV" & vbNullString  '--Vtype(1)
                           tmpRst3!Vdt = dtpToDt.Value  '--Vdt(2)
                           tmpRst3!VYear = gCYear '--VYear
                           tmpRst3!VFirm = gCCode  '--VFirm
                           tmpRst3!VCtrNo = wEntSrNo
                           tmpRst3!VCrAcCode = gCDepotMainFirm
                           If wDepotHoAc > 0 Then
                              tmpRst3!VDrAcCode = wDepotHoAc
                           Else
                              tmpRst3!VDrAcCode = wMillCode
                           End If
                           tmpRst3!vamt = wRBrkIgst
                           tmpRst3!VNar1 = "I Gst On Brokerage For - " + CStr(Format(dtpToDt.Value, "mmm-yyyy"))
                           tmpRst3!VTopCrDr = 1
                           tmpRst3!VBillVno = wMillCode
                           If optRegilur.Value = True Then tmpRst3!LateDays = 1
                           If opt2ndDay.Value = True Then tmpRst3!LateDays = 2
                           tmpRst3.Update
                        End If
                     End If
                     If wUnloadAmt > 0 Then
                           wEntSrNo = wEntSrNo + 1
                           tmpRst3.AddNew
                           tmpRst3!Vno = tmpVno '--Vno(0)
                           tmpRst3!Vtype = "JV" & vbNullString  '--Vtype(1)
                           tmpRst3!Vdt = dtpToDt.Value  '--Vdt(2)
                           tmpRst3!VYear = gCYear '--VYear
                           tmpRst3!VFirm = gCCode  '--VFirm
                           tmpRst3!VCtrNo = wEntSrNo
                           tmpRst3!VCrAcCode = gCDepotMainFirm
                           If wDepotHoAc > 0 Then
                              tmpRst3!VDrAcCode = wDepotHoAc
                           Else
                              tmpRst3!VDrAcCode = wMillCode
                           End If
                           tmpRst3!vamt = wUnloadAmt
                           tmpRst3!VNar1 = "Unloading Chg. For - " + CStr(Format(dtpToDt.Value, "mmm-yyyy"))
                           tmpRst3!VTopCrDr = 1
                           tmpRst3!VBillVno = wMillCode
                           If optRegilur.Value = True Then tmpRst3!LateDays = 1
                           If opt2ndDay.Value = True Then tmpRst3!LateDays = 2
                           tmpRst3.Update
                           '--For Log Add & Modi
                           If gCYear >= 2023 And gCIsLog = 1 Then
                              If PreVnoJv = 0 Then
                                 AMD = "A"
                              Else
                                 AMD = "M"
                              End If
                              gCn.BeginTrans
                              XX = "Select max(LogNo) from tblVoucher_Log where vtype='JV' and vno=" & tmpVno & " and Vfirm = '" & gCCode & "'  and" _
                                   & " VYear=" & gCYear
    
                              jj = GProcRstOpen(tmpRstLog, XX, "R")
                              If jj > 0 Then
                                 wLogNo = IIf(IsNull(tmpRstLog.Fields(0)), 0, tmpRstLog.Fields(0)) + 1
                              Else
                                 wLogNo = 1
                              End If
                              gCn.Execute "Insert into tblVoucher_Log select tblVoucher.*," & wLogNo & " as LogNo,'" & AMD & "' as LogTp ,'" & gUserName & "' as UserName ,'" & Format(Now, "YYYY/MM/DD hh:mm:ss") & "' as LogDate from tblVoucher where vtype='JV' and vno=" & tmpVno & " and Vfirm = '" & gCCode & "'  and" _
                                          & " VYear=" & gCYear
                               
                              gCn.CommitTrans
                              tmpRstLog.Close
                           End If
                     End If
                     
                     '-- end
                  End If
                  cmdExit.SetFocus
                  Exit Sub
               Else
                        
                     If wFirmHoAc > 0 Then
                        Y = "Select SlBillNo,tblSale.SlSubAmt,SlTaxAmt,SlTaxAmt2,SlTaxAmt3,SlBillAmt,ItName,tblSale.Vno from tblsale,tblSaleSub,tblMastItem where tblSale.VFirm = '" & gCDepotMainFirmCompCd & "' and SlAcDrCode = " & wFirmHoAc & " and tblSale.Vdt = ('" & Format(dtpToDt.Value, "yyyy/mm/dd") & "') " _
                            & " and tblSale.Vno = tblSaleSub.Vno and tblSale.Vtype = tblSaleSub.Vtype and tblSale.Vyear = tblSaleSub.Vyear and tblSale.Vfirm = tblSaleSub.Vfirm and tblSaleSub.SlSubItCode = tblMastItem.ItCode and left(tblMastItem.ItName,4) = 'Comm'    and tblSale.SlMillCode = " & wMillCode & " and SlTmpChr1 = '" & DayStr & "'  "
                     Else
                        Y = "Select SlBillNo,tblSale.SlSubAmt,SlTaxAmt,SlTaxAmt2,SlTaxAmt3,SlBillAmt,ItName,tblSale.Vno from tblsale,tblSaleSub,tblMastItem where tblSale.VFirm = '" & gCDepotMainFirmCompCd & "' and SlAcDrCode = " & wMillCode & " and tblSale.Vdt = ('" & Format(dtpToDt.Value, "yyyy/mm/dd") & "') and SlTmpChr1 = '" & DayStr & "' " _
                            & " and tblSale.Vno = tblSaleSub.Vno and tblSale.Vtype = tblSaleSub.Vtype and tblSale.Vyear = tblSaleSub.Vyear and tblSale.Vfirm = tblSaleSub.Vfirm and tblSaleSub.SlSubItCode = tblMastItem.ItCode and left(tblMastItem.ItName,4) = 'Comm'"
                     End If
                     
                     
                     If GProcRstOpen(tmpRst3, Y, "R", gCn) > 0 Then
                        .Formulas(14) = "wCommInvNo = '" & tmpRst3.Fields(0) & "'"
                        MailCommInvNo = tmpRst3.Fields(0)
                     Else
                        .Formulas(14) = "wCommInvNo = ''"
                         MailCommInvNo = ""
                     End If
                        If wFirmHoAc > 0 Then
                            Y = "Select SlBillNo,tblSale.SlSubAmt,SlTaxAmt,SlTaxAmt2,SlTaxAmt3,SlBillAmt,ItName,tblSale.Vno from tblsale,tblSaleSub,tblMastItem where tblSale.VFirm = '" & gCDepotMainFirmCompCd & "' and SlAcDrCode = " & wFirmHoAc & " and tblSale.Vdt = ('" & Format(dtpToDt.Value, "yyyy/mm/dd") & "') " _
                            & " and tblSale.Vno = tblSaleSub.Vno and tblSale.Vtype = tblSaleSub.Vtype and tblSale.Vyear = tblSaleSub.Vyear and tblSale.Vfirm = tblSaleSub.Vfirm and tblSaleSub.SlSubItCode = tblMastItem.ItCode and left(tblMastItem.ItName,4) = 'Brok' and SlTmpChr1 = '" & DayStrBrk & "' "
                        Else
                            Y = "Select SlBillNo,tblSale.SlSubAmt,SlTaxAmt,SlTaxAmt2,SlTaxAmt3,SlBillAmt,ItName,tblSale.Vno from tblsale,tblSaleSub,tblMastItem where tblSale.VFirm = '" & gCDepotMainFirmCompCd & "' and SlAcDrCode = " & wMillCode & " and tblSale.Vdt = ('" & Format(dtpToDt.Value, "yyyy/mm/dd") & "') " _
                            & " and tblSale.Vno = tblSaleSub.Vno and tblSale.Vtype = tblSaleSub.Vtype and tblSale.Vyear = tblSaleSub.Vyear and tblSale.Vfirm = tblSaleSub.Vfirm and tblSaleSub.SlSubItCode = tblMastItem.ItCode and left(tblMastItem.ItName,4) = 'Brok' and SlTmpChr1 = '" & DayStrBrk & "' "
                        End If
                     If GProcRstOpen(tmpRst3, Y, "R", gCn) > 0 Then
                        .Formulas(15) = "wBrokInvNo = '" & tmpRst3.Fields(0) & "'"
                        MailBrokInvNo = tmpRst3.Fields(0)
                     Else
                        .Formulas(15) = "wBrokInvNo = ''"
                        MailBrokInvNo = ""
                     End If
                     .Formulas(17) = "wBrokLess = " & Val(txt(7)) & ""
                     .Formulas(18) = "wBrokLessDes1 = '" & txt(8) & "'"
                     .Formulas(19) = "wBrokLessDes2 = '" & txt(9) & "'"
                     .Formulas(20) = "w2ndBills = ''"
                     If opt2ndDay.Value = True Then .Formulas(20) = "w2ndBills = 'Only 2nd Day Bills'"
                     If optRegilur.Value = True Then .Formulas(20) = "w2ndBills = 'Only Regular Bills'"
                     If cmdCrystal(3).Value = True Then
'                        Dim objCrystal As CRAXDRT.Application
'                        Dim objReport As CRAXDRT.Report
'                        Dim Tattach As String
'                        Dim Mfile As String
'                        Dim Mcp As String
'                        Dim PtyMailId As String
'                        Dim PtyMailIdCc As String
'                        Dim PtyMailIdBcc As String
'                        Dim MailSub As String
'                        Dim MailStr As String
                        Set objCrystal = New CRAXDRT.Application
                         Tattach = gReportPath & "rptMillStatement.rpt"
                        Set objReport = objCrystal.OpenReport(Tattach, 1)
                        objReport.RecordSelectionFormula = "{TmpGentbl.UserName}='" & gUserName & "'"
                        If gCIBOffice = "I" Then
                            objReport.Database.LogOnServer "pdsodbc.dll", "Hitrix", "Vishnu", "sa", "123456"
                        Else
                            objReport.Database.LogOnServer "pdsodbc.dll", "Hitrix", "VishnuMum", "sa", "123456"
                         End If
                        '---   X"
'                        Dim j As Integer
                        For j = 1 To objReport.FormulaFields.Count
                        Select Case objReport.FormulaFields(j).Name
                        Case "{@FirmNm}"
                                  If gCDepotMainFirm > 0 Then
                                     wMainFirmNm = GProcGetColumnValue("tblMastAccount", "AcCode", "'" & gCDepotMainFirm & "' ", "N", "AcName", "S")
                                     objReport.FormulaFields(j).text = "Trim(" & Chr(39) & UCase(wMainFirmNm) & Chr(39) & ")"
                                  Else
                                      objReport.FormulaFields(j).text = "Trim(" & Chr(39) & UCase(gCName) & Chr(39) & ")"
                                  End If
                        Case "{@wAdd1}"
                                  If gCDepotMainFirm > 0 Then
                                       wMainFirmAd1 = GProcGetColumnValue("tblMastAccount", "AcCode", "'" & gCDepotMainFirm & "' ", "N", "AcAdd1", "S")
                                      objReport.FormulaFields(j).text = "Trim(" & Chr(39) & wMainFirmAd1 & Chr(39) & ")"
                                  Else
                                      objReport.FormulaFields(j).text = "Trim(" & Chr(39) & gCAdd1 & Chr(39) & ")"
                                  End If
                        Case "{@wAdd2}"
                                  If gCDepotMainFirm > 0 Then
                                       wMainFirmAd2 = GProcGetColumnValue("tblMastAccount", "AcCode", "'" & gCDepotMainFirm & "' ", "N", "AcAdd2", "S")
                                      objReport.FormulaFields(j).text = "Trim(" & Chr(39) & wMainFirmAd2 & Chr(39) & ")"
                                  Else
                                      objReport.FormulaFields(j).text = "Trim(" & Chr(39) & gCAdd2 & Chr(39) & ")"
                                  End If
                        Case "{@wPhNo}"
                            objReport.FormulaFields(j).text = "Trim(" & Chr(39) & gCPhNo & Chr(39) & ")"
                        Case "{@wFrToDt}"
                           objReport.FormulaFields(j).text = "Trim(" & Chr(39) & "'From : " & dtpFrDt & "   To " & dtpToDt & "'" & Chr(39) & ")"
                        Case "{@wToDt}"
                           objReport.FormulaFields(j).text = "Trim(" & Chr(39) & dtpToDt & Chr(39) & ")"
                        Case "{@STRegNo}"
                             objReport.FormulaFields(j).text = "Trim(" & Chr(39) & gCSTRegNo & Chr(39) & ")"
                        Case "{@Wpanno}"
                            objReport.FormulaFields(j).text = "Trim(" & Chr(39) & gCPAN & Chr(39) & ")"
                        Case "{@wHamaliRt}"
                            objReport.FormulaFields(j).text = gCHamaliRt
                        Case "{@wLevyRt }"
                            objReport.FormulaFields(j).text = gCLevyRt
                        Case "{@wOpBal }"
                            objReport.FormulaFields(j).text = OpAmt
                        Case "{@wOtherNote}"
                            objReport.FormulaFields(j).text = "Trim(" & Chr(39) & txt(6) & Chr(39) & ")"
                        Case "{@wOtherAmt }"
                            objReport.FormulaFields(j).text = "Trim(" & Chr(39) & txt(5) & Chr(39) & ")"
                        Case "{@wDepoConsi }"
                            objReport.FormulaFields(j).text = "Trim(" & Chr(39) & wDepoConsi & Chr(39) & ")"
                        Case "{@wDepoConsi }"
                            objReport.FormulaFields(j).text = "Trim(" & Chr(39) & wDepoConsi & Chr(39) & ")"
                        Case "{@wCommInvNo})"
                                 objReport.FormulaFields(j).text = "Trim(" & Chr(39) & MailCommInvNo & Chr(39) & ")"
                        Case "{@wBrokInvNo}"
                                 objReport.FormulaFields(j).text = "Trim(" & Chr(39) & MailBrokInvNo & Chr(39) & ")"
                        Case "{@wMainFirmGst}"
                                  If gCDepotMainFirm > 0 Then
                                     objReport.FormulaFields(j).text = "Trim(" & Chr(39) & wMainFirmGSTIN & Chr(39) & ")"
                                   Else
                                     objReport.FormulaFields(j).text = "Trim(" & Chr(39) & gCGStin & Chr(39) & ")"
                                   End If
                       Case "{@w2ndBills}"
                             objReport.FormulaFields(j).text = "Trim(" & Chr(39) & "" & Chr(39) & ")"
                             If opt2ndDay.Value = True Then objReport.FormulaFields(j).text = "Trim(" & Chr(39) & "Only 2nd Day Bills" & Chr(39) & ")"
                             If optRegilur.Value = True Then objReport.FormulaFields(j).text = "Trim(" & Chr(39) & "Only Regular Bills" & Chr(39) & ")"
                       End Select
                       Next j
                       wMillMailId = GProcGetColumnValue("tblMastAccount", "AcName", txtMillName, "S", "AcEmail", "S")
                       If wMillMailId = "" Or IsNull(wMillMailId) Then
                            MsgBox "Mail ID Not Found"
                             Exit Sub
                        End If
                        PtyMailId = wMillMailId
                        
                        Mfile = CStr(App.Path) + "\MailPDF\" + gCCode + "-" + gUserName + "-" + "Depot_Statement"
                        ExportReportToPDF objReport, Mfile & ".pdf", "foo"
                        Mfile = Mfile + ".pdf"
                        
                        
                        
'                        ExportReportToPDF objReport, Mid(Tattach, 1, Len(Tattach) - 4) & ".pdf", "foo"
'                        Mfile = gReportPath & "rptMillStatement.pdf"
                        MailSub = "Monthly Statement  for the month of " + MonthName(Month(dtpToDt)) + "  " + CStr(Year(dtpToDt))
                        MailStr = "Dear Sir, " + Chr(13) + Chr(10)
                        MailStr = MailStr + "Attaching herewith the monthly Statement for the month of " + MonthName(Month(dtpToDt)) + "  " + CStr(Year(dtpToDt)) + Chr(13) + Chr(10)
                        
            
            '                   MailStr = MailStr + "Please find attached herewith GST Booking Order" & vbLf
            '                   MailStr = MailStr + "Note : This is system generated email, Please do not reply."
                         gMailsendToMill = False
                         Call SendEmail1("" & Trim(gCMailId) & "", "" & PtyMailId & "", MailSub, MailStr, "" & PtyMailIdCc & "", "" & PtyMailIdBcc & "", Mfile)
'                         Kill Mid(Tattach, 1, Len(Tattach) - 4) & ".pdf"
                         Exit Sub
                     Else
                         .ReportFileName = gReportPath & "rptMillStatement.rpt"
                     End If
               End If
    Case "Mill Billing Statement"
            If ValidateData = False Then Exit Sub
                If txtMillName <> "" Then
                   wMillCode = GProcGetColumnValue("tblMastAccount", "AcName", txtMillName, "S", "AcCode", "N")
                Else
                   wMillCode = 0
                End If
                wBrokBilYN = GProcGetColumnValue("tblMastAccount", "AcName", txtMillName, "S", "AcIsBrokerageLessInComm", "N")
                wCommRt = GProcGetColumnValue("tblMastAccount", "AcName", txtMillName, "S", "AcComRt", "N")
                DoEvents
                DoEvents
                DoEvents
                DoEvents
                If chkPaymentDirectToMill.Value = 1 Then
                   wCommRt2 = GProcGetColumnValue("tblMastAccount", "AcName", txtMillName, "S", "AcComRt2ndDay", "N")
                   If wCommRt2 > 0 Then
                      wCommRt = wCommRt2
                   End If
                End If
                wCommOn = GProcGetColumnValue("tblMastAccount", "AcName", txtMillName, "S", "AcComOn", "N")
                
                
                gCmd.CommandText = "PrcPrepareMillStatementMillBill"
                gCmd.Parameters.Refresh
                gCmd.Parameters("@Fr_dt") = Format(dtpFrDt, "yyyy/mm/dd")
                gCmd.Parameters("@To_dt") = Format(dtpToDt, "yyyy/mm/dd")
'                gCmd.Parameters("@YearStart_Dt") = Format(gCYSDate, "yyyy/mm/dd")
                gCmd.Parameters("@UCode") = gUserName
                gCmd.Parameters("@IsDirectPayment") = chkPaymentDirectToMill.Value
                gCmd.Parameters("@VFirm") = Trim(gCCode)
                If wMillCode <> 0 Then gCmd.Parameters("@MillCode") = wMillCode
                gCmd.Parameters("@CommRt") = wCommRt
                gCmd.Parameters("@CommOn") = wCommOn
                gCmd.Parameters("@SrvTxRt") = Val(txt(0))
                gCmd.Parameters("@SWCessRt") = Val(txt(1))
                gCmd.Parameters("@KKCessRt") = Val(txt(2))
                gCmd.Parameters("@FreightPd") = Val(txt(3))
                gCmd.Parameters("@FreightAdd") = Val(txt(10))
                gCmd.Execute
                X = "{TmpGentbl.UserName}='" & gUserName & "' "
               .SelectionFormula = X
               .Formulas(4) = "wFrToDt = 'From : ' & '" & dtpFrDt & "' & '  To ' & '" & dtpToDt & "'"
               .Formulas(5) = "wToDt = '" & dtpToDt & "'"
               .Formulas(6) = "STRegNo='" & gCSTRegNo & "'"
               .Formulas(7) = "Wpanno  = '" & gCPAN & "'"
               .Formulas(8) = "wHamaliRt  = " & gCHamaliRt & ""
               .Formulas(9) = "wLevyRt  = " & gCLevyRt & ""
               .Formulas(10) = "wOpBal  = " & OpAmt & ""
               .Formulas(11) = "wGSTNo = '" & gCGStin & "'"
               .Formulas(12) = "wCommissionOn = " & wCommOn & ""
               .Formulas(15) = "wOtherNote  = '" & txt(6) & "'"
               .Formulas(16) = "wOtherAmt  = " & txt(5) & ""
               .Formulas(17) = "wBrokLess = " & Val(txt(7)) & ""
               .Formulas(18) = "wBrokLessDes1 = '" & txt(8) & "'"
               .Formulas(19) = "wBrokLessDes2 = '" & txt(9) & "'"
               If chkPaymentDirectToMill.Value = 1 Then
                  .Formulas(20) = "wDirectPayment = 'Direct Payment To Mill'"
               Else
                  .Formulas(20) = "wDirectPayment = ''"
               End If
               .Formulas(21) = "wAddFrghtNote  = '" & txt(11) & "'"
               
               
               If cmdCrystal(2).Value = True Then
               
                  '-- For Brokerage
                  
                  X = "Select sum(iif(nar1 = 'Bag',round(Amt2*Amt1,0),iif(nar1 = 'Kg',round(Amt12*Amt1,0),round((Amt6*Amt1)/100,0)))) from TmpGentbl where nar6 = 'A' and UserName = '" & gUserName & "'"
                  If GProcRstOpen(tmpRst3, X, "R", gCn) > 0 Then
                     Y = GProcGetColumnValue("TblMastAccount", "ACName", txtMillName, "S", "AcMillBillFirm", "S")
                     lblMainFirm.Caption = GProcGetColumnValue("tblMastCompany", "CCode", Y, "S", "CName", "S")
                     If IsNull(tmpRst3.Fields(0)) Then
                         wRBrkAmt = 0
                         lblBrkAmt.Caption = ""
                         wRBrkCgst = 0
                         lblBrkCgst.Caption = ""
                         wRBrkSgst = 0
                         lblBrkSgst.Caption = ""
    
                         wRBrkIgst = 0
                         wRBrkIgst = 0
                         lblBrkTotalAmt.Caption = ""
                     Else
                         wRBrkAmt = tmpRst3.Fields(0) - Val(txt(7))
                         lblBrkAmt.Caption = Str(wRBrkAmt)
                         wRBrkCgst = Round((wRBrkAmt * Val(txt(0))) / 100, 2)
                         wRBrkCgst = GProcMakeRounding(wRBrkCgst)
                         lblBrkCgst.Caption = Str(wRBrkCgst)
                         wRBrkSgst = Round((wRBrkAmt * Val(txt(1))) / 100, 2)
                         wRBrkSgst = GProcMakeRounding(wRBrkSgst)
                         lblBrkSgst.Caption = Str(wRBrkSgst)
    
                         wRBrkIgst = Round((wRBrkAmt * Val(txt(2))) / 100, 2)
                         wRBrkIgst = GProcMakeRounding(wRBrkIgst)
                         lblBrkIgst.Caption = Str(wRBrkIgst)
                         lblBrkTotalAmt.Caption = Str(wRBrkAmt + wRBrkCgst + wRBrkSgst + wRBrkIgst)
                     End If
                     Y = "Select SlBillNo,tblSale.SlSubAmt,SlTaxAmt,SlTaxAmt2,SlTaxAmt3,SlBillAmt,ItName,tblSale.Vno from tblsale,tblSaleSub,tblMastItem where tblSale.VFirm = '" & Y & "' and SlAcDrCode = " & wMillCode & " and tblSale.Vdt = ('" & Format(dtpToDt.Value, "yyyy/mm/dd") & "') " _
                         & " and tblSale.Vno = tblSaleSub.Vno and tblSale.Vtype = tblSaleSub.Vtype and tblSale.Vyear = tblSaleSub.Vyear and tblSale.Vfirm = tblSaleSub.Vfirm and tblSaleSub.SlSubItCode = tblMastItem.ItCode and left(tblMastItem.ItName,4) = 'Brok' and SlTmpNo3 = " & chkPaymentDirectToMill.Value & " "
                     If GProcRstOpen(tmpRst3, Y, "R", gCn) > 0 Then
                        lblBrkEntryDbNtNo.Caption = tmpRst3.Fields(0)
                        lblBrkEntryBrkAmt.Caption = tmpRst3.Fields(1)
                        lblBrkEntryCgst.Caption = tmpRst3.Fields(2)
                        lblBrkEntrySgst.Caption = tmpRst3.Fields(3)
                        lblBrkEntryIgst.Caption = tmpRst3.Fields(4)
                        lblBrkEntryTotalAmt.Caption = tmpRst3.Fields(5)
                        PreVnoBrk = tmpRst3.Fields(7)
                     
                     End If
                  End If
                  '-- For Commission
                  X = "Select sum(amt1) from TmpGentbl where nar6 = 'B' and UserName = '" & gUserName & "'"
                  If GProcRstOpen(tmpRst3, X, "R", gCn) > 0 Then
                     Y = GProcGetColumnValue("TblMastAccount", "ACName", txtMillName, "S", "AcMillBillFirm", "S")
                     If Y = "" Or IsNull(Y) Then
                        MsgBox ("Please Check Mill Bill In Firm At Account Master")
                        Exit Sub
                     End If
                     If gAcCodeSYHank = 0 Then
                        MsgBox ("Please Check Commission Sales A/c In Master Setting")
                        Exit Sub
                     End If
                     If gAcCodeST = 0 Then
                        MsgBox ("Please Check Brokerage Sales A/c In Master Setting")
                        Exit Sub
                     End If
                     lblMainFirm.Caption = GProcGetColumnValue("tblMastCompany", "CCode", Y, "S", "CName", "S")
                     
                     If wBrokBilYN = 1 Then
                        wRCommOn = Val(tmpRst3.Fields(0)) - (Val(txt(3)) + Val(txt(5)) + wRBrkAmt) + Val(txt(10))
                     Else
                        wRCommOn = Val(tmpRst3.Fields(0)) - (Val(txt(3)) + Val(txt(5))) + Val(txt(10))
                     End If
                     wRCommAmt = Round((wRCommOn * wCommRt) / 100, 2)
'                     wRCommAmt = wRCommAmt
                     wRCommAmt = GProcMakeRounding(wRCommAmt)
                     lblCmAmt.Caption = Str(wRCommAmt)
                     wRCommCgst = Round((wRCommAmt * Val(txt(0))) / 100, 2)
                     wRCommCgst = GProcMakeRounding(wRCommCgst)
                     lblCmCgst.Caption = Str(wRCommCgst)
                     wRCommSgst = Round((wRCommAmt * Val(txt(1))) / 100, 2)
                     wRCommSgst = GProcMakeRounding(wRCommSgst)
                     lblCmSgst.Caption = Str(wRCommSgst)
                  
                     wRCommIgst = Round((wRCommAmt * Val(txt(2))) / 100, 2)
                     wRCommIgst = GProcMakeRounding(wRCommIgst)
                     lblCmIgst.Caption = Str(wRCommIgst)
                     lblCmTotalAmt.Caption = Str(wRCommAmt + wRCommCgst + wRCommSgst + wRCommIgst)
                  
                     Y = "Select SlBillNo,tblSale.SlSubAmt,SlTaxAmt,SlTaxAmt2,SlTaxAmt3,SlBillAmt,ItName,tblSale.Vno from tblsale,tblSaleSub,tblMastItem where tblSale.VFirm = '" & Y & "' and SlAcDrCode = " & wMillCode & " and tblSale.Vdt = ('" & Format(dtpToDt.Value, "yyyy/mm/dd") & "') " _
                         & " and tblSale.Vno = tblSaleSub.Vno and tblSale.Vtype = tblSaleSub.Vtype and tblSale.Vyear = tblSaleSub.Vyear and tblSale.Vfirm = tblSaleSub.Vfirm and tblSaleSub.SlSubItCode = tblMastItem.ItCode and left(tblMastItem.ItName,4) = 'Comm'  and SlTmpNo3 = " & chkPaymentDirectToMill.Value & " "
                     
                     
                     If GProcRstOpen(tmpRst3, Y, "R", gCn) > 0 Then
                        lblCmEntryDbNtNo.Caption = tmpRst3.Fields(0)
                        lblCmEntryCmAmt.Caption = tmpRst3.Fields(1)
                        lblCmEntryCgst.Caption = tmpRst3.Fields(2)
                        lblCmEntrySgst.Caption = tmpRst3.Fields(3)
                        lblCmEntryIgst.Caption = tmpRst3.Fields(4)
                        lblCmEntryTotalAmt.Caption = tmpRst3.Fields(5)
                        PreVnoCM = tmpRst3.Fields(7)
                     End If
                  End If
                  If MsgBox("Genreate Bills ? Are you sure?", vbYesNo) = vbYes Then
                     Y = GProcGetColumnValue("TblMastAccount", "ACName", txtMillName, "S", "AcMillBillFirm", "S")
                     
                     '===== For Commission
                     
                     If PreVnoCM <> 0 Then
                        gCn.Execute "delete from tblSaleSub where vtype='SB' and vno=" & PreVnoCM & " and Vfirm = '" & Y & "'  and " _
                        & " VYear=" & gCYear
                        
                        gCn.Execute "delete from tblSale where vtype='SB' and vno=" & PreVnoCM & " and Vfirm = '" & Y & "'  and " _
                        & " VYear=" & gCYear
                        tmpBillNo = lblCmEntryDbNtNo.Caption  '--Inv No
                        Mon = Month(dtpToDt)
                        tmpVno = PreVnoCM
                     Else
                        tmpBillNo = GProcGenerateInvNoBrokComm("SB", "'" & Y & "'", "" & Y & "") '--Inv No
                        Mon = Month(dtpToDt)
                        tmpVno = GProcGenerateIdMonthwise("TblSale", "Vno", Mon, "Vdt", "Vtype='SB' and VFirm = '" & Y & "'   and VYear=" & gCYear)
                     End If
                     
                     '---For Save In tblSale
                     If wRCommAmt > 0 Then
                        X = "select * from tblSale where vtype='SB' and tblSale.VFirm = '" & Y & "' and " _
                            & " VYear=" & gCYear
                        GProcRstOpen tmpRst3, X, "O"
                        tmpRst3.AddNew
                        tmpRst3!Vno = tmpVno '--Vno(0)
                        tmpRst3!Vtype = "SB" & vbNullString  '--Vtype(1)
                        tmpRst3!Vdt = dtpToDt.Value  '--Vdt(2)
                        tmpRst3!VYear = gCYear '--VYear
                        tmpRst3!VFirm = Y  '--VFirm
                        tmpRst3!SlAcDrCode = wMillCode
                        tmpRst3!SlAcCrCode = gAcCodeSYHank
                        tmpRst3!SlMillCode = wMillCode
                        tmpRst3!SlBroker = GProcGetColumnValue("tblMastAccount", "AcName", "-", "S", "AcCode", "N", "")
                        tmpRst3!SlBillNo = tmpBillNo
                        tmpRst3!SlBillDt = dtpToDt.Value
                        tmpRst3!SlSubAmt = wRCommAmt
                        tmpRst3!SlTmpNo1 = wRCommOn
                        tmpRst3!SlTmpNo2 = wCommRt
                        tmpRst3!SlBillAmt = wRCommAmt + wRCommCgst + wRCommSgst + wRCommIgst
                        tmpRst3!SlTaxableAmt = wRCommAmt
                        tmpRst3!SlTaxRate = Val(txt(0))
                        tmpRst3!SlTaxAmt = wRCommCgst
                        tmpRst3!SlTaxRate2 = Val(txt(1))
                        tmpRst3!SlTaxAmt2 = wRCommSgst
                        tmpRst3!SlTaxRate3 = Val(txt(2))
                        tmpRst3!SlTaxAmt3 = wRCommIgst
                        tmpRst3!SlTransport = GProcGetColumnValue("tblMastNarration", "Narration", "-", "S", "NarrCode", "N")
                        tmpRst3!SlTaxCode = GProcGetColumnValue("tblMastNarration", "Narration", "-", "S", "NarrCode", "N")
                        tmpRst3!DueDays = 0
                        tmpRst3!DueDate = dtpToDt.Value
                        If chkPaymentDirectToMill.Value = 1 Then
                           tmpRst3!SlTmpChr1 = "Commission-Direct"
                        Else
                           tmpRst3!SlTmpChr1 = "Commission"
                        End If
                        tmpRst3!SlTmpChr2 = "For The Month Of " + Format(dtpToDt, "MMM-YY")
                        tmpRst3!SlTmpChr3 = "[Including G S T ]"
                        If chkPaymentDirectToMill.Value = 1 Then
                           tmpRst3!SlTmpNo3 = 1
                        Else
                           tmpRst3!SlTmpNo3 = 0
                        End If
                        tmpRst3.Update
                                              
                        '---For Save In tblSaleSub
                        
                        X = "select * from tblSaleSub where vtype='SB' and tblSaleSub.VFirm = '" & Y & "' and " _
                            & " VYear=" & gCYear
                        GProcRstOpen tmpRst3, X, "O"
                        
                        
                        
                        tmpRst3.AddNew
                        tmpRst3!Vno = tmpVno '--Vno(0)
                        tmpRst3!Vtype = "SB" & vbNullString  '--Vtype(1)
                        tmpRst3!Vdt = dtpToDt.Value  '--Vdt(2)
                        tmpRst3!VYear = gCYear '--VYear
                        tmpRst3!VFirm = Y '--VFirm
                        tmpRst3!SlSubItSrNo = 1 '--Sr No
                        '--Itcode(4)
                        If chkPaymentDirectToMill.Value = 1 Then
                           tmpRst3!SlSubItCode = gAcCodeSYExempt '--Item Code
                        Else
                           tmpRst3!SlSubItCode = gCommissionItem '--Item Code
                        End If
                        tmpRst3!SlSubAmt = wRCommAmt
                        tmpRst3!SlSubGodown = GProcGetColumnValue("tblMastNarration", "Narration", "-", "S", "NarrCode", "N", "")
                        '-- P It Ctrl No =Booking It ctrl No
                        tmpRst3.Update
                        
                        '--For Log Add & Modi
                        If gCYear >= 2023 And gCIsLog = 1 Then
                           If PreVnoCM = 0 Then
                              AMD = "A"
                           Else
                              AMD = "M"
                           End If
                           gCn.BeginTrans
                           XX = "Select max(LogNo) from tblSale_Log where vtype='SB' and vno=" & tmpVno & " and Vfirm = '" & Y & "'  and" _
                                & " VYear=" & gCYear

                           jj = GProcRstOpen(tmpRstLog, XX, "R")
                           If jj > 0 Then
                              wLogNo = IIf(IsNull(tmpRstLog.Fields(0)), 0, tmpRstLog.Fields(0)) + 1
                           Else
                              wLogNo = 1
                           End If
                           gCn.Execute "Insert into tblSale_Log select tblSale.*," & wLogNo & " as LogNo,'" & AMD & "' as LogTp ,'" & gUserName & "' as UserName ,'" & Format(Now, "YYYY/MM/DD hh:mm:ss") & "' as LogDate from tblSale where vtype='SB' and vno=" & tmpVno & " and Vfirm = '" & Y & "'  and" _
                                       & " VYear=" & gCYear
                           
                           gCn.Execute "Insert into tblSaleSub_Log select tblSaleSub.*," & wLogNo & " as LogNo,'" & AMD & "' as LogTp ,'" & gUserName & "' as UserName ,'" & Format(Now, "YYYY/MM/DD hh:mm:ss") & "' as LogDate from tblSaleSub where vtype='SB' and vno=" & tmpVno & " and Vfirm = '" & Y & "'  and" _
                                       & " VYear=" & gCYear
                           
                           gCn.CommitTrans
                           tmpRstLog.Close
                        End If
                     
                     End If
                     
                     
                     
                     
                     '-- end commission
                  
                     '===== For Brokerage
                     
                     If PreVnoBrk <> 0 Then
                        gCn.Execute "delete from tblSaleSub where vtype='SB' and vno=" & PreVnoBrk & " and Vfirm = '" & Y & "'  and " _
                        & " VYear=" & gCYear
                        
                        gCn.Execute "delete from tblSale where vtype='SB' and vno=" & PreVnoBrk & " and Vfirm = '" & Y & "'  and " _
                        & " VYear=" & gCYear
                        tmpBillNo = lblBrkEntryDbNtNo.Caption  '--Inv No
                        Mon = Month(dtpToDt)
                        tmpVno = PreVnoBrk
                     Else
                        tmpBillNo = GProcGenerateInvNoBrokComm("SB", "'" & Y & "'", "" & Y & "") '--Inv No
                        Mon = Month(dtpToDt)
                        tmpVno = GProcGenerateIdMonthwise("TblSale", "Vno", Mon, "Vdt", "Vtype='SB' and VFirm = '" & Y & "'   and VYear=" & gCYear)
                     End If
                     
                     '---For Save In tblSale
                     If wRBrkAmt > 0 Then
                        X = "select * from tblSale where vtype='SB' and tblSale.VFirm = '" & Y & "' and " _
                            & " VYear=" & gCYear
                        GProcRstOpen tmpRst3, X, "O"
                        tmpRst3.AddNew
                        tmpRst3!Vno = tmpVno '--Vno(0)
                        tmpRst3!Vtype = "SB" & vbNullString  '--Vtype(1)
                        tmpRst3!Vdt = dtpToDt.Value  '--Vdt(2)
                        tmpRst3!VYear = gCYear '--VYear
                        tmpRst3!VFirm = Y  '--VFirm
                        tmpRst3!SlAcDrCode = wMillCode
                        tmpRst3!SlAcCrCode = gAcCodeST
                        tmpRst3!SlMillCode = wMillCode
                        tmpRst3!SlBroker = GProcGetColumnValue("tblMastAccount", "AcName", "-", "S", "AcCode", "N", "")
                        tmpRst3!SlBillNo = tmpBillNo
                        tmpRst3!SlBillDt = dtpToDt.Value
                        tmpRst3!SlSubAmt = wRBrkAmt
                        tmpRst3!SlBillAmt = wRBrkAmt + wRBrkCgst + wRBrkSgst + wRBrkIgst
                        tmpRst3!SlTaxableAmt = wRBrkAmt
                        tmpRst3!SlTaxRate = Val(txt(0))
                        tmpRst3!SlTaxAmt = wRBrkCgst
                        tmpRst3!SlTaxRate2 = Val(txt(1))
                        tmpRst3!SlTaxAmt2 = wRBrkSgst
                        tmpRst3!SlTaxRate3 = Val(txt(2))
                        tmpRst3!SlTaxAmt3 = wRBrkIgst
                        tmpRst3!SlTransport = GProcGetColumnValue("tblMastNarration", "Narration", "-", "S", "NarrCode", "N")
                        tmpRst3!SlTaxCode = GProcGetColumnValue("tblMastNarration", "Narration", "-", "S", "NarrCode", "N")
                        tmpRst3!DueDays = 0
                        tmpRst3!DueDate = dtpToDt.Value
                        If chkPaymentDirectToMill.Value = 1 Then
                           tmpRst3!SlTmpChr1 = "Brokerage-Direct"
                        Else
                           tmpRst3!SlTmpChr1 = "Brokerage   "
                        End If
                        tmpRst3!SlTmpChr2 = "For  The Month  Of " + Format(dtpToDt, "MMM-YY")
                        tmpRst3!SlTmpChr3 = "[ Including G S T ]"
                        If chkPaymentDirectToMill.Value = 1 Then
                           tmpRst3!SlTmpNo3 = 1
                        Else
                           tmpRst3!SlTmpNo3 = 0
                        End If
                        tmpRst3.Update
                                              
                        '---For Save In tblSaleSub
                        
                        X = "select * from tblSaleSub where vtype='SB' and tblSaleSub.VFirm = '" & Y & "' and " _
                            & " VYear=" & gCYear
                        GProcRstOpen tmpRst3, X, "O"
                        tmpRst3.AddNew
                        tmpRst3!Vno = tmpVno '--Vno(0)
                        tmpRst3!Vtype = "SB" & vbNullString  '--Vtype(1)
                        tmpRst3!Vdt = dtpToDt.Value  '--Vdt(2)
                        tmpRst3!VYear = gCYear '--VYear
                        tmpRst3!VFirm = Y '--VFirm
                        tmpRst3!SlSubItSrNo = 1 '--Sr No
                        '--Itcode(4)
                        If chkPaymentDirectToMill.Value = 1 Then
                            tmpRst3!SlSubItCode = gAcCodePYHank  '--Item Code
                        Else

                            tmpRst3!SlSubItCode = gBrokerageItem  '--Item Code
                        End If
                        tmpRst3!SlSubAmt = wRBrkAmt
                        tmpRst3!SlSubGodown = GProcGetColumnValue("tblMastNarration", "Narration", "-", "S", "NarrCode", "N", "")
                        '-- P It Ctrl No =Booking It ctrl No
                        tmpRst3.Update
                     
                        '--For Log Add & Modi
                        If gCYear >= 2023 And gCIsLog = 1 Then
                           If PreVnoBrk = 0 Then
                              AMD = "A"
                           Else
                              AMD = "M"
                           End If
                           gCn.BeginTrans
                           XX = "Select max(LogNo) from tblSale_Log where vtype='SB' and vno=" & tmpVno & " and Vfirm = '" & Y & "'  and" _
                                & " VYear=" & gCYear

                           jj = GProcRstOpen(tmpRstLog, XX, "R")
                           If jj > 0 Then
                              wLogNo = IIf(IsNull(tmpRstLog.Fields(0)), 0, tmpRstLog.Fields(0)) + 1
                           Else
                              wLogNo = 1
                           End If
                           gCn.Execute "Insert into tblSale_Log select tblSale.*," & wLogNo & " as LogNo,'" & AMD & "' as LogTp ,'" & gUserName & "' as UserName ,'" & Format(Now, "YYYY/MM/DD hh:mm:ss") & "' as LogDate from tblSale where vtype='SB' and vno=" & tmpVno & " and Vfirm = '" & Y & "'  and" _
                                       & " VYear=" & gCYear
                           
                           gCn.Execute "Insert into tblSaleSub_Log select tblSaleSub.*," & wLogNo & " as LogNo,'" & AMD & "' as LogTp ,'" & gUserName & "' as UserName ,'" & Format(Now, "YYYY/MM/DD hh:mm:ss") & "' as LogDate from tblSaleSub where vtype='SB' and vno=" & tmpVno & " and Vfirm = '" & Y & "'  and" _
                                       & " VYear=" & gCYear
                           
                           gCn.CommitTrans
                           tmpRstLog.Close
                        End If
                     End If
                  End If
                  cmdExit.SetFocus
                  Exit Sub
               Else
                   Y = GProcGetColumnValue("TblMastAccount", "ACName", txtMillName, "S", "AcMillBillFirm", "S")
                   Y = "Select SlBillNo,tblSale.SlSubAmt,SlTaxAmt,SlTaxAmt2,SlTaxAmt3,SlBillAmt,ItName,tblSale.Vno from tblsale,tblSaleSub,tblMastItem where tblSale.VFirm = '" & Y & "' and SlAcDrCode = " & wMillCode & " and tblSale.Vdt = ('" & Format(dtpToDt.Value, "yyyy/mm/dd") & "') " _
                        & " and tblSale.Vno = tblSaleSub.Vno and tblSale.Vtype = tblSaleSub.Vtype and tblSale.Vyear = tblSaleSub.Vyear and tblSale.Vfirm = tblSaleSub.Vfirm and tblSaleSub.SlSubItCode = tblMastItem.ItCode and left(tblMastItem.ItName,4) = 'Comm'  and SlTmpNo3 = " & chkPaymentDirectToMill.Value & "  "
                   If GProcRstOpen(tmpRst3, Y, "R", gCn) > 0 Then
                     .Formulas(13) = "wCommInvNo = '" & tmpRst3.Fields(0) & "'"
                      MailCommInvNo = tmpRst3.Fields(0)
                   Else
                     .Formulas(13) = "wCommInvNo = ''"
                      MailCommInvNo = ""
                   End If
                 '-- For Brokerage
                   Y = GProcGetColumnValue("TblMastAccount", "ACName", txtMillName, "S", "AcMillBillFirm", "S")
                   Y = "Select SlBillNo,tblSale.SlSubAmt,SlTaxAmt,SlTaxAmt2,SlTaxAmt3,SlBillAmt,ItName,tblSale.Vno from tblsale,tblSaleSub,tblMastItem where tblSale.VFirm = '" & Y & "' and SlAcDrCode = " & wMillCode & " and tblSale.Vdt = ('" & Format(dtpToDt.Value, "yyyy/mm/dd") & "') " _
                       & " and tblSale.Vno = tblSaleSub.Vno and tblSale.Vtype = tblSaleSub.Vtype and tblSale.Vyear = tblSaleSub.Vyear and tblSale.Vfirm = tblSaleSub.Vfirm and tblSaleSub.SlSubItCode = tblMastItem.ItCode and left(tblMastItem.ItName,4) = 'Brok'  and SlTmpNo3 = " & chkPaymentDirectToMill.Value & " "
                   
                   
                   
                   If GProcRstOpen(tmpRst3, Y, "R", gCn) > 0 Then
                     .Formulas(14) = "wBrokInvNo = '" & tmpRst3.Fields(0) & "'"
                      MailBrokInvNo = tmpRst3.Fields(0)
                   Else
                     .Formulas(14) = "wBrokInvNo = ''"
                      MailBrokInvNo = ""
                   End If
                   
                   
                   If cmdCrystal(3).Value = True Then
'                        Dim objCrystal As CRAXDRT.Application
'                        Dim objReport As CRAXDRT.Report
'                        Dim Tattach As String
'                        Dim Mfile As String
'                        Dim Mcp As String
'                        Dim PtyMailId As String
'                        Dim PtyMailIdCc As String
'                        Dim PtyMailIdBcc As String
'                        Dim MailSub As String
'                        Dim MailStr As String
                        Set objCrystal = New CRAXDRT.Application
                         Tattach = gReportPath & "rptMillStatementMillBill.rpt"
                        Set objReport = objCrystal.OpenReport(Tattach, 1)
                        objReport.RecordSelectionFormula = "{TmpGentbl.UserName}='" & gUserName & "'"
                        If gCIBOffice = "I" Then
                            objReport.Database.LogOnServer "pdsodbc.dll", "Hitrix", "Vishnu", "sa", "123456"
                        Else
                            objReport.Database.LogOnServer "pdsodbc.dll", "Hitrix", "VishnuMum", "sa", "123456"
                         End If
                        '---   X"
'                        Dim j As Integer
                        For j = 1 To objReport.FormulaFields.Count
                        Select Case objReport.FormulaFields(j).Name
                        Case "{@FirmNm}"
                                   objReport.FormulaFields(j).text = "Trim(" & Chr(39) & UCase(gCName) & Chr(39) & ")"
                        Case "{@wAdd1}"
                                   objReport.FormulaFields(j).text = "Trim(" & Chr(39) & gCAdd1 & Chr(39) & ")"
                        Case "{@wAdd2}"
                                   objReport.FormulaFields(j).text = "Trim(" & Chr(39) & gCAdd2 & Chr(39) & ")"
                        Case "{@wPhNo}"
                            objReport.FormulaFields(j).text = "Trim(" & Chr(39) & gCPhNo & Chr(39) & ")"
                        Case "{@wFrToDt}"
                           objReport.FormulaFields(j).text = "Trim(" & Chr(39) & "'From : " & dtpFrDt & "   To " & dtpToDt & "'" & Chr(39) & ")"
                        Case "{@wToDt}"
                           objReport.FormulaFields(j).text = "Trim(" & Chr(39) & dtpToDt & Chr(39) & ")"
                        Case "{@STRegNo}"
                             objReport.FormulaFields(j).text = "Trim(" & Chr(39) & gCSTRegNo & Chr(39) & ")"
                        Case "{@Wpanno}"
                            objReport.FormulaFields(j).text = "Trim(" & Chr(39) & gCPAN & Chr(39) & ")"
                        Case "{@wHamaliRt}"
                            objReport.FormulaFields(j).text = gCHamaliRt
                        Case "{@wLevyRt }"
                            objReport.FormulaFields(j).text = gCLevyRt
                        Case "{@wOpBal }"
                            objReport.FormulaFields(j).text = OpAmt
                        Case "{@wGSTNo}"
                            objReport.FormulaFields(j).text = "Trim(" & Chr(39) & gCGStin & Chr(39) & ")"
                        Case "{@wCommissionOn}"
                            objReport.FormulaFields(j).text = wCommOn
                        Case "{@wOtherNote}"
                            objReport.FormulaFields(j).text = "Trim(" & Chr(39) & txt(6) & Chr(39) & ")"
                        Case "{@wwAddFrghtNote}"
                            objReport.FormulaFields(j).text = "Trim(" & Chr(39) & txt(11) & Chr(39) & ")"
                        Case "{@wOtherAmt}"
                            objReport.FormulaFields(j).text = Val(txt(5))    '"Trim(" & Chr(39) & txt(5) & Chr(39) & ")"
                        Case "{@wCommInvNo}"
                                 objReport.FormulaFields(j).text = "Trim(" & Chr(39) & MailCommInvNo & Chr(39) & ")"
                        Case "{@wBrokInvNo}"
                                 objReport.FormulaFields(j).text = "Trim(" & Chr(39) & MailBrokInvNo & Chr(39) & ")"
                        Case "{@wGSTNo}"
                                   objReport.FormulaFields(j).text = "Trim(" & Chr(39) & gCGStin & Chr(39) & ")"
                        Case "{@wDirectPayment}"
                              If chkPaymentDirectToMill.Value = 1 Then
                                 objReport.FormulaFields(j).text = "Trim(" & Chr(39) & "Direct Payment To Mill" & Chr(39) & ")"
                              Else
                                 objReport.FormulaFields(j).text = "Trim(" & Chr(39) & "" & Chr(39) & ")"
                              End If
                       End Select
                       Next j
                       wMillMailId = GProcGetColumnValue("tblMastAccount", "AcName", txtMillName, "S", "AcEmail", "S")
                       If wMillMailId = "" Or IsNull(wMillMailId) Then
                            MsgBox "Mail ID Not Found"
                             Exit Sub
                        End If
                        PtyMailId = wMillMailId
                        
                        Mfile = CStr(App.Path) + "\MailPDF\" + gCCode + "-" + gUserName + "-" + "Mill_Bill_Statement"
                        ExportReportToPDF objReport, Mfile & ".pdf", "foo"
                        Mfile = Mfile + ".pdf"
                        
                        
                        
'                        ExportReportToPDF objReport, Mid(Tattach, 1, Len(Tattach) - 4) & ".pdf", "foo"
'                        Mfile = gReportPath & "rptMillStatementMillBill.pdf"
                        MailSub = "Monthly Statement of Direct Sales for the month of " + MonthName(Month(dtpToDt)) + "  " + CStr(Year(dtpToDt))
                        MailStr = "Dear Sir, " + Chr(13) + Chr(10)
                        MailStr = MailStr + "Attaching herewith the monthly Statement of Direct Sales for the month of " + MonthName(Month(dtpToDt)) + "  " + CStr(Year(dtpToDt)) + Chr(13) + Chr(10)
            '                   MailStr = MailStr + "Please find attached herewith GST Booking Order" & vbLf
            '                   MailStr = MailStr + "Note : This is system generated email, Please do not reply."
                         gMailsendToMill = False
                         Call SendEmail1("" & Trim(gCMailId) & "", "" & PtyMailId & "", MailSub, MailStr, "" & PtyMailIdCc & "", "" & PtyMailIdBcc & "", Mfile)
'                         Kill Mid(Tattach, 1, Len(Tattach) - 4) & ".pdf"
                         Exit Sub
                Else
                   .ReportFileName = gReportPath & "rptMillStatementMillBill.rpt"
               End If
           End If
    End Select
    .Action = 1
    .PageZoom (120)
End With
End Sub
Private Sub cmdExit_Click()
    DoEvents
    Unload Me
End Sub

Private Sub dtpFrDt_GotFocus()
GProcSelectBox Me.ActiveControl
End Sub

Private Sub dtpToDt_GotFocus()
GProcSelectBox Me.ActiveControl
End Sub
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
    Case 13:
       Sendkeys "{TAB}"
    Case 27:
        Unload Me
    End Select
End Sub
Private Sub Form_Load()
Dim i As Long
Me.Top = 700
Me.Left = Screen.Width / 2 - Me.Width / 2
dtpFrDt.Value = Format(gCYSDate, "dd/MM/yyyy")
dtpToDt.Value = Format(gCYEDate, "dd/MM/yyyy")
dtpBillDt.Value = Format(gCYSDate, "dd/MM/yyyy")
FrameApplyDbNt.Visible = True
Me.Caption = gSelectedMenu
For i = 0 To 5
    txt(i) = "0.00"
Next i
txt(6) = ""
txt(7) = "0.00"
txt(8) = ""
txt(9) = ""
txt(10) = "0.00"
txt(11) = ""

Select Case gSelectedMenu
    Case "Consignment/ Depot Statement" '--"Consignment/ Depot Statement"
          txt(6).Enabled = True
          lblMainFirm.Caption = GProcGetColumnValue("tblMastCompany", "CCode", gCDepotMainFirmCompCd, "S", "CName", "S")
          cmdCrystal(3).Visible = True
          optRegilur.Visible = True
          optRegilur.Enabled = True
          opt2ndDay.Visible = True
          opt2ndDay.Enabled = True
          OptAllBills.Visible = True
          OptAllBills.Enabled = True
          OptAllBills.Value = True
    Case "Mill Billing Statement"
         chkPaymentDirectToMill.Visible = True
         chkPaymentDirectToMill.Enabled = True
          cmdCrystal(3).Visible = True
         Label1(31).Visible = True
         Label1(32).Visible = True
         txt(10).Visible = True
         txt(10).Enabled = True
'        txt(3).Visible = False
End Select
 '-- Set InActivate Ctrl Back Color
GProcSetInActivateCtrlBackClr Me
End Sub
Private Sub txt_GotFocus(Index As Integer)
GProcSelectBox Me.ActiveControl
End Sub

Private Sub txt_KeyPress(Index As Integer, KeyAscii As Integer)
If Index <= 5 Then
    KeyAscii = GProcNumberOnly(KeyAscii)
End If
End Sub

Private Sub txt_LostFocus(Index As Integer)
ValidateData
End Sub

Private Sub txtMillName_GotFocus()
GProcSelectBox Me.ActiveControl
End Sub

Private Sub txtMillName_KeyPress(KeyAscii As Integer)
Dim X As String
If KeyAscii = 13 Then Exit Sub '-- Enter
Select Case gSelectedMenu
    Case "Consignment/ Depot Statement", "Mill Billing Statement" '"Consignment/ Depot Statement"
           If gSelectedMenu = "Consignment/ Depot Statement" Then
               gClsSearch.SearchMultiField "qryMillHelpWitIni", "BillSr,MillName,AcOurCode,AcGSTINDepotHo,AcMsOmsParty", Array("Mill Ini.", "Mill Name", "Code", "", ""), Array(1000, 8000, 1000, 0, 0), "SaleType in ('27','28','30') and vfirm = '" & gCCode & "' ", Trim(Chr(KeyAscii)), "MillName", txtMillName.Left + Me.Left, txtMillName.Top + Me.Top + 1000, True
           Else
               gClsSearch.SearchMultiField "qryMillHelpWitIni", "BillSr,MillName,AcOurCode,AcGSTIN,AcMsOmsParty", Array("Mill Ini.", "Mill Name", "Code", "", ""), Array(1000, 8000, 1000, 0, 0), "SaleType in ('27','28','30') and vfirm = '" & gCCode & "' ", Trim(Chr(KeyAscii)), "MillName", txtMillName.Left + Me.Left, txtMillName.Top + Me.Top + 1000, True
           End If
            KeyAscii = 0
            If gClsSearch.SearchMultiRetCol(1) <> "" Then
               txtMillName = gClsSearch.SearchMultiRetCol(1)
               If Left(gClsSearch.SearchMultiRetCol(3), 2) = "27" Then
                  txt(0) = Format(gCSrvTaxRt, FStr)
                  txt(1) = Format(gCSBCessRt, FStr)
                  txt(2) = 0
               Else
                  txt(0) = 0
                  txt(1) = 0
                  txt(2) = Format(gCKKCessRt, FStr)
               End If
               If gSelectedMenu = "Mill Billing Statement" Then
                  X = GProcGetColumnValue("TblMastAccount", "ACName", txtMillName, "S", "AcMillBillFirm", "S")
                  lblMainFirm.Caption = GProcGetColumnValue("tblMastCompany", "CCode", X, "S", "CName", "S")
               Else
'                  X = GProcGetColumnValue("TblMastAccount", "ACName", txtMillName, "S", "AcDepotFirm", "S")
'                  If X = "" Then '-- For Consignment Main Firm
'                     X = GProcGetColumnValue("TblMastAccount", "ACName", txtMillName, "S", "AcConsignmentFirm", "S")
'                  End If
'                  lblMainFirm.Caption = GProcGetColumnValue("tblMastCompany", "CCode", gCDepotMainFirmCompCd, "S", "CName", "S")
               End If
            End If
'        gClsSearch.SearchMultiField "tblMastaccount", "AcName,AcCity", Array("Mill Name", "City"), Array(7000, 3000), "(AgCode=90029 ) ", Chr(KeyAscii), "AcName", txtMillName.Left + Me.Left, txtMillName.Top + Me.Top + 1000, False
End Select
'If gClsSearch.SearchMultiRetCol(0) <> "" Then
'    txtMillName = gClsSearch.SearchMultiRetCol(0)
'    KeyAscii = 0
'End If
End Sub
'------------ Validate Selection
Private Function ValidateData() As Boolean
Dim i As Long
'-- Mill Name
If txtMillName = "" Then
    MsgBox "Check Mill Name", vbOKOnly + vbInformation, Me.Caption
    txtMillName.SetFocus
    Exit Function
End If
'-- %
For i = 0 To 5
    If IsNumeric(txt(i)) = False Then
        txt(i) = "0.00"
    Else
        txt(i) = Format(txt(i), FStr)
    End If
Next i
txt(3) = Format(txt(3), PStr)
ValidateData = True
End Function
