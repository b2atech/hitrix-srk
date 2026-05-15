VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form frmPurchaseOtherGST 
   BackColor       =   &H00FBF2E1&
   BorderStyle     =   0  'None
   Caption         =   "Purchase Entry"
   ClientHeight    =   9675
   ClientLeft      =   0
   ClientTop       =   15
   ClientWidth     =   14625
   FillColor       =   &H000040C0&
   FillStyle       =   0  'Solid
   Icon            =   "frmPurchaseOtherGST.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   9675
   ScaleWidth      =   14625
   ShowInTaskbar   =   0   'False
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
      Index           =   71
      Left            =   14880
      TabIndex        =   189
      Text            =   "71"
      Top             =   180
      Visible         =   0   'False
      Width           =   570
   End
   Begin VB.Frame Frame1 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFC0C0&
      Caption         =   "RCM Bill  Printing"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   5655
      Left            =   3240
      TabIndex        =   167
      Top             =   2340
      Visible         =   0   'False
      Width           =   6165
      Begin VB.TextBox txtPInvno 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Index           =   0
         Left            =   2820
         MaxLength       =   10
         TabIndex        =   180
         Top             =   3105
         Width           =   2055
      End
      Begin VB.CommandButton cmdPrintBill 
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
         Left            =   720
         TabIndex        =   179
         Top             =   5010
         Width           =   1335
      End
      Begin VB.TextBox txtPInvno 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Index           =   1
         Left            =   2820
         MaxLength       =   10
         TabIndex        =   178
         Top             =   3675
         Width           =   2055
      End
      Begin VB.CommandButton cmdPrintBill 
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
         Left            =   2430
         TabIndex        =   177
         Top             =   5025
         Width           =   1335
      End
      Begin VB.CommandButton cmdPrintBill 
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
         Left            =   4245
         TabIndex        =   176
         Top             =   5010
         Width           =   1335
      End
      Begin VB.CheckBox chkInvChallan 
         BackColor       =   &H00FFC0C0&
         Caption         =   "Original For Buyer "
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   285
         Index           =   0
         Left            =   2820
         TabIndex        =   175
         Top             =   450
         Visible         =   0   'False
         Width           =   2445
      End
      Begin VB.CheckBox chkInvChallan 
         BackColor       =   &H00FFC0C0&
         Caption         =   "Duplicate For Transporter"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   285
         Index           =   1
         Left            =   2820
         TabIndex        =   174
         Top             =   900
         Visible         =   0   'False
         Width           =   3030
      End
      Begin VB.CheckBox chkInvChallan 
         BackColor       =   &H00FFC0C0&
         Caption         =   "Triplicate For Accounts"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   285
         Index           =   2
         Left            =   2820
         TabIndex        =   173
         Top             =   1395
         Visible         =   0   'False
         Width           =   3030
      End
      Begin VB.CheckBox chkInvChallan 
         BackColor       =   &H00FFC0C0&
         Caption         =   "Qudraplicate For Mill"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   285
         Index           =   3
         Left            =   2820
         TabIndex        =   172
         Top             =   1890
         Visible         =   0   'False
         Width           =   3030
      End
      Begin VB.CheckBox chkInvChallan 
         BackColor       =   &H00FFC0C0&
         Caption         =   "Delivery Challan"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   285
         Index           =   4
         Left            =   2820
         TabIndex        =   171
         Top             =   2340
         Visible         =   0   'False
         Width           =   3030
      End
      Begin VB.CheckBox ChkGod 
         BackColor       =   &H00FFC0C0&
         Caption         =   "Godown Copy"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   195
         Left            =   240
         TabIndex        =   170
         Top             =   4320
         Visible         =   0   'False
         Width           =   1575
      End
      Begin VB.CheckBox chkOff 
         BackColor       =   &H00FFC0C0&
         Caption         =   "Office Copy"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   195
         Left            =   2160
         TabIndex        =   169
         Top             =   4320
         Visible         =   0   'False
         Width           =   1575
      End
      Begin VB.CheckBox ChkCus 
         BackColor       =   &H00FFC0C0&
         Caption         =   "Customer Copy"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   195
         Left            =   4080
         TabIndex        =   168
         Top             =   4320
         Visible         =   0   'False
         Width           =   1815
      End
      Begin VB.Line Line3 
         X1              =   0
         X2              =   6120
         Y1              =   4800
         Y2              =   4800
      End
      Begin VB.Label Label38 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "From  Bill  No    :"
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
         Left            =   915
         TabIndex        =   182
         Top             =   3195
         Width           =   1605
      End
      Begin VB.Label Label39 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "To       Bill  No   :"
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
         Left            =   915
         TabIndex        =   181
         Top             =   3765
         Width           =   1590
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
      Height          =   405
      Index           =   72
      Left            =   2250
      TabIndex        =   38
      Text            =   "72"
      Top             =   5280
      Width           =   1965
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
      Index           =   70
      Left            =   10680
      TabIndex        =   34
      Text            =   "70"
      Top             =   8400
      Width           =   2265
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
      Height          =   405
      Index           =   69
      Left            =   9540
      TabIndex        =   33
      Text            =   "69"
      Top             =   8340
      Width           =   825
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
      Height          =   375
      Index           =   68
      Left            =   7920
      TabIndex        =   32
      Text            =   "68"
      Top             =   8340
      Width           =   1530
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
      Height          =   405
      Index           =   67
      Left            =   2250
      Locked          =   -1  'True
      TabIndex        =   11
      Text            =   "txttemp(67)"
      ToolTipText     =   "Press F3 for Creation"
      Top             =   4380
      Width           =   4740
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
      Index           =   67
      Left            =   17160
      TabIndex        =   184
      Text            =   "67"
      Top             =   5700
      Visible         =   0   'False
      Width           =   570
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
      Height          =   375
      Index           =   66
      Left            =   17820
      TabIndex        =   183
      Text            =   "66"
      Top             =   6060
      Visible         =   0   'False
      Width           =   570
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
      Index           =   65
      Left            =   17100
      TabIndex        =   166
      Text            =   "65"
      Top             =   5220
      Visible         =   0   'False
      Width           =   570
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
      Height          =   405
      Index           =   65
      Left            =   2250
      Locked          =   -1  'True
      TabIndex        =   42
      Text            =   "txttemp(65)"
      ToolTipText     =   "Press F3 for Creation"
      Top             =   6660
      Width           =   4740
   End
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
      Height          =   405
      Index           =   65
      Left            =   2250
      TabIndex        =   41
      Text            =   "txtCode(65)"
      ToolTipText     =   "Press F3 for Creation"
      Top             =   6180
      Width           =   1485
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
      Height          =   405
      Index           =   64
      Left            =   2250
      TabIndex        =   40
      Text            =   "64"
      Top             =   5760
      Width           =   1965
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
      Height          =   405
      Index           =   63
      Left            =   6030
      TabIndex        =   39
      Text            =   "63"
      Top             =   5340
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
      Height          =   405
      Index           =   62
      Left            =   5025
      TabIndex        =   43
      Text            =   "62"
      Top             =   6180
      Width           =   1965
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
      Height          =   375
      Index           =   61
      Left            =   17880
      TabIndex        =   163
      Text            =   "61"
      Top             =   5700
      Visible         =   0   'False
      Width           =   570
   End
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
      Index           =   5
      Left            =   2430
      TabIndex        =   5
      Text            =   "txtCode(5)"
      ToolTipText     =   "Press F3 for Creation"
      Top             =   540
      Width           =   1905
   End
   Begin VB.CheckBox chkIsRCM 
      BackColor       =   &H00FBF2E1&
      Caption         =   "Is Reverse Charge Mechanism Applicable ?"
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
      Left            =   1305
      TabIndex        =   7
      Top             =   1845
      Width           =   4650
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
      Height          =   405
      Index           =   60
      Left            =   4230
      TabIndex        =   51
      Text            =   "60"
      Top             =   9052
      Width           =   2265
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
      Height          =   405
      Index           =   59
      Left            =   1890
      TabIndex        =   48
      Text            =   "59"
      Top             =   9045
      Width           =   645
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
      Height          =   405
      Index           =   58
      Left            =   4230
      TabIndex        =   50
      Text            =   "58"
      Top             =   8625
      Width           =   2265
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
      Index           =   57
      Left            =   1890
      TabIndex        =   47
      Text            =   "57"
      Top             =   8625
      Width           =   645
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
      Height          =   405
      Index           =   56
      Left            =   4230
      TabIndex        =   49
      Text            =   "56"
      Top             =   8190
      Width           =   2265
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
      Index           =   54
      Left            =   17325
      TabIndex        =   156
      Text            =   "54"
      Top             =   0
      Visible         =   0   'False
      Width           =   375
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
      Index           =   53
      Left            =   15705
      TabIndex        =   155
      Text            =   "53"
      Top             =   585
      Visible         =   0   'False
      Width           =   375
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
      Index           =   55
      Left            =   1890
      TabIndex        =   46
      Text            =   "55"
      Top             =   8190
      Width           =   645
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
      Index           =   54
      Left            =   15240
      Locked          =   -1  'True
      TabIndex        =   52
      Text            =   "txttemp(54)"
      ToolTipText     =   "Press F3 for Creation"
      Top             =   6360
      Visible         =   0   'False
      Width           =   885
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
      Index           =   51
      Left            =   16875
      TabIndex        =   154
      Text            =   "51"
      Top             =   45
      Visible         =   0   'False
      Width           =   375
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
      Index           =   50
      Left            =   16425
      TabIndex        =   153
      Text            =   "50"
      Top             =   405
      Visible         =   0   'False
      Width           =   375
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
      Index           =   49
      Left            =   16425
      TabIndex        =   152
      Text            =   "49"
      Top             =   45
      Visible         =   0   'False
      Width           =   375
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
      Index           =   48
      Left            =   16020
      TabIndex        =   151
      Text            =   "48"
      Top             =   360
      Visible         =   0   'False
      Width           =   375
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
      Index           =   47
      Left            =   16065
      TabIndex        =   150
      Text            =   "47"
      Top             =   45
      Visible         =   0   'False
      Width           =   375
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
      Height          =   375
      Index           =   52
      Left            =   1935
      TabIndex        =   44
      Text            =   "52"
      Top             =   7230
      Width           =   1965
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
      Height          =   375
      Index           =   46
      Left            =   17955
      TabIndex        =   137
      Text            =   "46"
      Top             =   5220
      Visible         =   0   'False
      Width           =   570
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
      Index           =   39
      Left            =   15885
      Locked          =   -1  'True
      TabIndex        =   135
      Text            =   "txttemp(39)"
      ToolTipText     =   "Press F3 for Creation"
      Top             =   6840
      Visible         =   0   'False
      Width           =   825
   End
   Begin VB.CheckBox chkIsCapitalGoods 
      BackColor       =   &H00FBF2E1&
      Caption         =   "Is Capital Goods Purchase"
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
      Left            =   15750
      TabIndex        =   134
      Top             =   4590
      Visible         =   0   'False
      Width           =   3255
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
      Height          =   375
      Index           =   45
      Left            =   17910
      TabIndex        =   133
      Text            =   "45"
      Top             =   4815
      Visible         =   0   'False
      Width           =   570
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
      Height          =   375
      Index           =   44
      Left            =   2250
      MaxLength       =   50
      TabIndex        =   37
      Text            =   "44"
      Top             =   4860
      Width           =   4740
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
      Height          =   375
      Index           =   43
      Left            =   17955
      TabIndex        =   132
      Text            =   "43"
      Top             =   4005
      Visible         =   0   'False
      Width           =   570
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
      Height          =   375
      Index           =   42
      Left            =   17865
      TabIndex        =   131
      Text            =   "42"
      Top             =   3555
      Visible         =   0   'False
      Width           =   570
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
      Height          =   375
      Index           =   41
      Left            =   17820
      TabIndex        =   130
      Text            =   "41"
      Top             =   3150
      Visible         =   0   'False
      Width           =   570
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
      Height          =   375
      Index           =   40
      Left            =   17820
      TabIndex        =   129
      Text            =   "40"
      Top             =   2790
      Visible         =   0   'False
      Width           =   570
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
      Height          =   375
      Index           =   39
      Left            =   17820
      TabIndex        =   128
      Text            =   "39"
      Top             =   2430
      Visible         =   0   'False
      Width           =   570
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
      Height          =   375
      Index           =   36
      Left            =   17865
      TabIndex        =   127
      Text            =   "36"
      Top             =   1980
      Visible         =   0   'False
      Width           =   570
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
      Height          =   405
      Index           =   26
      Left            =   8145
      TabIndex        =   26
      Text            =   "26"
      Top             =   6690
      Width           =   645
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
      Height          =   405
      Index           =   37
      Left            =   8145
      TabIndex        =   28
      Text            =   "37"
      Top             =   7140
      Width           =   645
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
      Index           =   27
      Left            =   10665
      TabIndex        =   27
      Text            =   "27"
      Top             =   6705
      Width           =   2265
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
      Index           =   38
      Left            =   10665
      TabIndex        =   29
      Text            =   "38"
      Top             =   7155
      Width           =   2265
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
      Left            =   13170
      Style           =   1  'Graphical
      TabIndex        =   116
      Top             =   6480
      Width           =   960
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
      Height          =   375
      Index           =   35
      Left            =   17820
      TabIndex        =   115
      Text            =   "35"
      Top             =   1620
      Visible         =   0   'False
      Width           =   570
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
      Left            =   16245
      TabIndex        =   54
      Text            =   "Cbo(22)"
      Top             =   5535
      Visible         =   0   'False
      Width           =   840
   End
   Begin VB.TextBox txttemp 
      Alignment       =   1  'Right Justify
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
      Height          =   360
      Index           =   35
      Left            =   15930
      MaxLength       =   5
      TabIndex        =   55
      Text            =   "txtemp(35)"
      Top             =   7470
      Visible         =   0   'False
      Width           =   780
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
      Index           =   22
      Left            =   17190
      TabIndex        =   109
      Text            =   "22"
      Top             =   4275
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.TextBox txt 
      Alignment       =   1  'Right Justify
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Index           =   16
      Left            =   10665
      TabIndex        =   36
      Text            =   "16"
      Top             =   9225
      Width           =   2265
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
      Index           =   15
      Left            =   10665
      TabIndex        =   20
      Text            =   "15"
      Top             =   4620
      Width           =   2265
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
      Left            =   17190
      TabIndex        =   105
      Text            =   "14"
      Top             =   3915
      Visible         =   0   'False
      Width           =   375
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
      Index           =   13
      Left            =   15885
      TabIndex        =   17
      Text            =   "13"
      Top             =   1800
      Visible         =   0   'False
      Width           =   615
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
      Index           =   11
      Left            =   15840
      TabIndex        =   15
      Text            =   "txttemp(11)"
      ToolTipText     =   "Press F3 for Creation & Press F2 for Selection"
      Top             =   945
      Visible         =   0   'False
      Width           =   1290
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
      Index           =   28
      Left            =   10665
      TabIndex        =   30
      Text            =   "28"
      Top             =   7590
      Width           =   2265
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
      Index           =   29
      Left            =   10665
      TabIndex        =   31
      Text            =   "29"
      Top             =   7995
      Width           =   2265
   End
   Begin VB.CommandButton CmdBagDetails 
      Caption         =   "Cartoon &Bag Details"
      Height          =   720
      Left            =   13170
      TabIndex        =   101
      Top             =   2985
      Width           =   960
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
      Height          =   375
      Index           =   8
      Left            =   17235
      TabIndex        =   59
      Text            =   "8"
      Top             =   2745
      Visible         =   0   'False
      Width           =   345
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
      Height          =   375
      Index           =   34
      Left            =   17775
      TabIndex        =   100
      Text            =   "34"
      Top             =   1215
      Visible         =   0   'False
      Width           =   570
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
      Index           =   19
      Left            =   16875
      TabIndex        =   53
      Text            =   "19"
      Top             =   6840
      Visible         =   0   'False
      Width           =   435
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
      Height          =   375
      Index           =   33
      Left            =   17820
      TabIndex        =   99
      Text            =   "33"
      Top             =   765
      Visible         =   0   'False
      Width           =   420
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
      Height          =   375
      Index           =   32
      Left            =   17775
      TabIndex        =   98
      Text            =   "32"
      Top             =   360
      Visible         =   0   'False
      Width           =   420
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
      Index           =   8
      Left            =   15885
      TabIndex        =   13
      Text            =   "txttemp(8)"
      ToolTipText     =   "Press F3 for Creation"
      Top             =   3015
      Visible         =   0   'False
      Width           =   1095
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
      Left            =   435
      TabIndex        =   6
      Text            =   "txttemp(5)"
      ToolTipText     =   "Press F3 for Creation"
      Top             =   960
      Width           =   7125
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
      Index           =   7
      Left            =   15885
      TabIndex        =   14
      Text            =   "txttemp(7)"
      ToolTipText     =   "Press F3 for Creation"
      Top             =   3465
      Visible         =   0   'False
      Width           =   1140
   End
   Begin VB.TextBox txtCrBalance 
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
      Height          =   375
      Left            =   15975
      TabIndex        =   19
      Text            =   "txtCrBalance"
      Top             =   2610
      Visible         =   0   'False
      Width           =   435
   End
   Begin VB.CommandButton cmdLedger 
      Caption         =   "Ledger"
      Height          =   375
      Left            =   15075
      TabIndex        =   97
      Top             =   9180
      Visible         =   0   'False
      Width           =   1095
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
      Height          =   375
      Index           =   25
      Left            =   17235
      TabIndex        =   73
      Text            =   "25"
      Top             =   4680
      Visible         =   0   'False
      Width           =   480
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
      Index           =   21
      Left            =   10665
      TabIndex        =   23
      Text            =   "21"
      Top             =   5835
      Width           =   2265
   End
   Begin VB.TextBox txtGrid 
      BackColor       =   &H80000000&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   645
      TabIndex        =   9
      Top             =   3240
      Visible         =   0   'False
      Width           =   1035
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
      Height          =   375
      Index           =   1
      Left            =   17145
      TabIndex        =   95
      Text            =   "1"
      Top             =   405
      Visible         =   0   'False
      Width           =   420
   End
   Begin VB.TextBox txt 
      Alignment       =   1  'Right Justify
      Height          =   315
      Index           =   2
      Left            =   17100
      TabIndex        =   93
      Text            =   "2"
      Top             =   855
      Visible         =   0   'False
      Width           =   525
   End
   Begin VB.TextBox txt 
      Alignment       =   1  'Right Justify
      Height          =   315
      Index           =   5
      Left            =   17190
      TabIndex        =   92
      Text            =   "5"
      Top             =   1800
      Visible         =   0   'False
      Width           =   525
   End
   Begin VB.TextBox txt 
      Alignment       =   1  'Right Justify
      Height          =   315
      Index           =   3
      Left            =   17100
      TabIndex        =   91
      Text            =   "3"
      Top             =   1215
      Visible         =   0   'False
      Width           =   525
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
      Left            =   14445
      Style           =   1  'Graphical
      TabIndex        =   87
      Top             =   9135
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
      Left            =   14760
      Style           =   1  'Graphical
      TabIndex        =   86
      ToolTipText     =   "To Add New Record"
      Top             =   9000
      Visible         =   0   'False
      Width           =   1095
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
      Left            =   13170
      Style           =   1  'Graphical
      TabIndex        =   84
      ToolTipText     =   "Exit"
      Top             =   7950
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
      Left            =   13170
      Style           =   1  'Graphical
      TabIndex        =   77
      ToolTipText     =   "To Save Record"
      Top             =   5835
      Width           =   960
   End
   Begin VB.CommandButton cmdBtn 
      Caption         =   "RCM Bill Print"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   645
      Index           =   9
      Left            =   13170
      Style           =   1  'Graphical
      TabIndex        =   85
      Top             =   7050
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
      Left            =   13170
      Style           =   1  'Graphical
      TabIndex        =   76
      ToolTipText     =   "To Delete Record"
      Top             =   5205
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
      Height          =   405
      Index           =   7
      Left            =   13170
      Style           =   1  'Graphical
      TabIndex        =   75
      ToolTipText     =   "To Modify Record"
      Top             =   4560
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
      Left            =   13170
      Style           =   1  'Graphical
      TabIndex        =   74
      ToolTipText     =   "To Add New Record"
      Top             =   3930
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
      Left            =   16905
      Style           =   1  'Graphical
      TabIndex        =   78
      Top             =   8400
      Visible         =   0   'False
      Width           =   780
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
      Left            =   13170
      Style           =   1  'Graphical
      TabIndex        =   79
      ToolTipText     =   "To Find Record"
      Top             =   2340
      Width           =   960
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
      Left            =   13665
      Style           =   1  'Graphical
      TabIndex        =   83
      ToolTipText     =   "Last"
      Top             =   1710
      Width           =   465
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
      Left            =   13170
      Style           =   1  'Graphical
      TabIndex        =   82
      ToolTipText     =   "First"
      Top             =   1710
      Width           =   465
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
      Left            =   13665
      Style           =   1  'Graphical
      TabIndex        =   81
      ToolTipText     =   "Next"
      Top             =   1080
      Width           =   465
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
      Index           =   0
      Left            =   8850
      TabIndex        =   12
      Text            =   "0"
      Top             =   585
      Width           =   1965
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
      Left            =   13170
      Style           =   1  'Graphical
      TabIndex        =   80
      ToolTipText     =   "Previous"
      Top             =   1080
      Width           =   465
   End
   Begin VB.TextBox txt 
      Alignment       =   1  'Right Justify
      Height          =   315
      Index           =   4
      Left            =   17100
      TabIndex        =   90
      Text            =   "4"
      Top             =   1485
      Visible         =   0   'False
      Width           =   525
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
      Index           =   6
      Left            =   2250
      Locked          =   -1  'True
      TabIndex        =   10
      Text            =   "txttemp(6)"
      ToolTipText     =   "Press F3 for Creation"
      Top             =   3930
      Width           =   4740
   End
   Begin VB.TextBox txt 
      Height          =   315
      Index           =   6
      Left            =   17235
      TabIndex        =   89
      Text            =   "6"
      Top             =   2115
      Visible         =   0   'False
      Width           =   435
   End
   Begin VB.TextBox txt 
      Alignment       =   1  'Right Justify
      Height          =   315
      Index           =   7
      Left            =   17145
      TabIndex        =   88
      Text            =   "7"
      Top             =   2475
      Visible         =   0   'False
      Width           =   525
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
      Height          =   375
      Index           =   9
      Left            =   8850
      TabIndex        =   1
      Text            =   "9"
      Top             =   1230
      Width           =   1965
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
      Left            =   17190
      TabIndex        =   61
      Text            =   "10"
      Top             =   3150
      Visible         =   0   'False
      Width           =   420
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
      Left            =   17280
      TabIndex        =   71
      Text            =   "11"
      Top             =   3465
      Visible         =   0   'False
      Width           =   375
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
      Index           =   12
      Left            =   15885
      TabIndex        =   16
      Text            =   "12"
      Top             =   1395
      Visible         =   0   'False
      Width           =   795
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
      Height          =   405
      Index           =   23
      Left            =   8145
      TabIndex        =   24
      Text            =   "23"
      Top             =   6270
      Width           =   645
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
      Index           =   24
      Left            =   10665
      TabIndex        =   25
      Text            =   "24"
      Top             =   6270
      Width           =   2265
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
      Index           =   20
      Left            =   17325
      TabIndex        =   56
      Text            =   "20"
      Top             =   6885
      Visible         =   0   'False
      Width           =   570
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
      Index           =   30
      Left            =   10665
      TabIndex        =   35
      Text            =   "30"
      Top             =   8820
      Width           =   2265
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
      Height          =   375
      Index           =   31
      Left            =   8850
      TabIndex        =   3
      Text            =   "txt(31)"
      Top             =   1890
      Width           =   1965
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
      Index           =   17
      Left            =   10665
      TabIndex        =   21
      Text            =   "17"
      Top             =   5010
      Width           =   2265
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
      Index           =   18
      Left            =   10665
      TabIndex        =   22
      Text            =   "18"
      Top             =   5415
      Width           =   2265
   End
   Begin MSMask.MaskEdBox mskDt 
      Height          =   375
      Index           =   2
      Left            =   11370
      TabIndex        =   0
      Top             =   593
      Width           =   1515
      _ExtentX        =   2672
      _ExtentY        =   661
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
   Begin MSMask.MaskEdBox mskDt 
      Height          =   375
      Index           =   10
      Left            =   11370
      TabIndex        =   2
      Top             =   1230
      Width           =   1515
      _ExtentX        =   2672
      _ExtentY        =   661
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
   Begin MSMask.MaskEdBox mskDt 
      Height          =   375
      Index           =   32
      Left            =   11370
      TabIndex        =   4
      Top             =   1890
      Width           =   1515
      _ExtentX        =   2672
      _ExtentY        =   661
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
      Bindings        =   "frmPurchaseOtherGST.frx":058A
      Left            =   4110
      Top             =   0
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
      Height          =   1485
      Left            =   315
      TabIndex        =   8
      Top             =   2355
      Width           =   12570
      _ExtentX        =   22172
      _ExtentY        =   2619
      _Version        =   393216
      FixedCols       =   0
      RowHeightMin    =   275
      BackColorFixed  =   -2147483647
      ForeColorFixed  =   -2147483634
      BackColorBkg    =   12632256
      Appearance      =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSMask.MaskEdBox mskDt 
      Height          =   375
      Index           =   14
      Left            =   16020
      TabIndex        =   18
      Top             =   2250
      Visible         =   0   'False
      Width           =   300
      _ExtentX        =   529
      _ExtentY        =   661
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
   Begin MSMask.MaskEdBox mskDt 
      Height          =   375
      Index           =   53
      Left            =   4905
      TabIndex        =   45
      Top             =   7230
      Width           =   1515
      _ExtentX        =   2672
      _ExtentY        =   661
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
   Begin VB.Label Label43 
      BackStyle       =   0  'Transparent
      Caption         =   "TDS On Amount   :"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000040C0&
      Height          =   255
      Left            =   360
      TabIndex        =   188
      Top             =   5340
      Width           =   1770
   End
   Begin VB.Label Label42 
      BackStyle       =   0  'Transparent
      Caption         =   "%"
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
      Left            =   10440
      TabIndex        =   187
      Top             =   8460
      Width           =   255
   End
   Begin VB.Label Label41 
      BackStyle       =   0  'Transparent
      Caption         =   "TCS On : "
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   375
      Left            =   7065
      TabIndex        =   186
      Top             =   8400
      Width           =   825
   End
   Begin VB.Label Label40 
      BackStyle       =   0  'Transparent
      Caption         =   "Sub Group  A/c     :"
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
      Left            =   360
      TabIndex        =   185
      Top             =   4380
      Width           =   2430
   End
   Begin VB.Label Label37 
      BackStyle       =   0  'Transparent
      Caption         =   "TDS A/c               :"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000040C0&
      Height          =   375
      Left            =   420
      TabIndex        =   165
      Top             =   6720
      Width           =   1890
   End
   Begin VB.Label Label36 
      BackStyle       =   0  'Transparent
      Caption         =   "TDS A/c  Code     :"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000040C0&
      Height          =   375
      Left            =   420
      TabIndex        =   164
      Top             =   6240
      Width           =   1890
   End
   Begin VB.Label Label35 
      BackStyle       =   0  'Transparent
      Caption         =   "TDS Amt. :"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000040C0&
      Height          =   375
      Left            =   420
      TabIndex        =   162
      Top             =   5820
      Width           =   1170
   End
   Begin VB.Label Label34 
      BackStyle       =   0  'Transparent
      Caption         =   "TDS Rate  :"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000040C0&
      Height          =   375
      Left            =   4860
      TabIndex        =   161
      Top             =   5400
      Width           =   1110
   End
   Begin VB.Label Label33 
      BackStyle       =   0  'Transparent
      Caption         =   "J.V. No    :"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000040C0&
      Height          =   375
      Left            =   3960
      TabIndex        =   160
      Top             =   6300
      Width           =   990
   End
   Begin VB.Label Label25 
      BackStyle       =   0  'Transparent
      Caption         =   "S GST       :"
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
      Left            =   2790
      TabIndex        =   159
      Top             =   8633
      Width           =   1245
   End
   Begin VB.Label lblRCMTotal 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "RCM Total"
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
      Height          =   375
      Left            =   13440
      TabIndex        =   158
      Top             =   7560
      Width           =   2325
   End
   Begin VB.Label lblRcmSubAmt 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "RcmSubAmount"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000040C0&
      Height          =   375
      Left            =   4275
      TabIndex        =   157
      Top             =   7800
      Width           =   2205
   End
   Begin VB.Label Label32 
      BackStyle       =   0  'Transparent
      Caption         =   "I  GST @  : "
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
      Left            =   675
      TabIndex        =   149
      Top             =   9113
      Width           =   1065
   End
   Begin VB.Label Label29 
      BackStyle       =   0  'Transparent
      Caption         =   "C GST @  : "
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
      Left            =   675
      TabIndex        =   148
      Top             =   8633
      Width           =   1065
   End
   Begin VB.Label Label27 
      BackStyle       =   0  'Transparent
      Caption         =   "S GST @  : "
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
      Left            =   675
      TabIndex        =   147
      Top             =   8198
      Width           =   1065
   End
   Begin VB.Label Label26 
      BackStyle       =   0  'Transparent
      Caption         =   "C GST       :"
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
      Left            =   2790
      TabIndex        =   146
      Top             =   8198
      Width           =   1125
   End
   Begin VB.Label Label23 
      BackStyle       =   0  'Transparent
      Caption         =   "I  GST        :"
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
      Left            =   2790
      TabIndex        =   145
      Top             =   9113
      Width           =   1215
   End
   Begin VB.Label Label22 
      BackStyle       =   0  'Transparent
      Caption         =   "Narration              :"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000040C0&
      Height          =   375
      Left            =   360
      TabIndex        =   144
      Top             =   4860
      Width           =   1890
   End
   Begin VB.Label Label18 
      BackStyle       =   0  'Transparent
      Caption         =   "Sub Amt   :"
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
      Left            =   2790
      TabIndex        =   143
      Top             =   7800
      Width           =   1110
   End
   Begin VB.Label Label14 
      BackStyle       =   0  'Transparent
      Caption         =   "Sales A/c"
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
      Height          =   285
      Left            =   7140
      TabIndex        =   142
      Top             =   5340
      Visible         =   0   'False
      Width           =   1155
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00000080&
      X1              =   360
      X2              =   6930
      Y1              =   7665
      Y2              =   7665
   End
   Begin VB.Label Label13 
      BackStyle       =   0  'Transparent
      Caption         =   "RCM Bill No  :"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000040C0&
      Height          =   375
      Left            =   585
      TabIndex        =   141
      Top             =   7230
      Width           =   1665
   End
   Begin VB.Label Label8 
      BackStyle       =   0  'Transparent
      Caption         =   "Date  :"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000040C0&
      Height          =   375
      Left            =   4095
      TabIndex        =   140
      Top             =   7230
      Width           =   675
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "&Party / Credit A/c :"
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
      Left            =   495
      TabIndex        =   139
      Top             =   585
      Width           =   1800
   End
   Begin VB.Label Label9 
      BackStyle       =   0  'Transparent
      Caption         =   "Purch/Debit  A/c   :"
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
      Left            =   360
      TabIndex        =   138
      Top             =   3930
      Width           =   2430
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H00000080&
      FillColor       =   &H00EDFDFE&
      FillStyle       =   0  'Solid
      Height          =   2475
      Left            =   360
      Top             =   7095
      Width           =   6585
   End
   Begin VB.Label Label16 
      BackStyle       =   0  'Transparent
      Caption         =   "&Tax Payable A/c  : "
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
      Left            =   15840
      TabIndex        =   136
      Top             =   7200
      Width           =   1785
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00EFB17A&
      BorderWidth     =   3
      X1              =   7650
      X2              =   7680
      Y1              =   0
      Y2              =   2160
   End
   Begin VB.Label lblPer1 
      BackStyle       =   0  'Transparent
      Caption         =   "%"
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
      Left            =   8850
      TabIndex        =   126
      Top             =   6263
      Width           =   255
   End
   Begin VB.Label lblSgstAmt 
      BackStyle       =   0  'Transparent
      Caption         =   "S GST Amt   : "
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
      Left            =   9255
      TabIndex        =   125
      Top             =   6263
      Width           =   1335
   End
   Begin VB.Label lblCgstAmt 
      BackStyle       =   0  'Transparent
      Caption         =   "C GST Amt   : "
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
      Left            =   9255
      TabIndex        =   124
      Top             =   6698
      Width           =   1335
   End
   Begin VB.Label lblIgstAmt 
      BackStyle       =   0  'Transparent
      Caption         =   "I  GST Amt   : "
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
      Left            =   9255
      TabIndex        =   123
      Top             =   7148
      Width           =   1335
   End
   Begin VB.Label lblPer2 
      BackStyle       =   0  'Transparent
      Caption         =   "%"
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
      Left            =   8865
      TabIndex        =   122
      Top             =   6698
      Width           =   255
   End
   Begin VB.Label lblPer3 
      BackStyle       =   0  'Transparent
      Caption         =   "%"
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
      Left            =   8865
      TabIndex        =   121
      Top             =   7148
      Width           =   255
   End
   Begin VB.Label lblSgstRate 
      BackStyle       =   0  'Transparent
      Caption         =   "S GST @  : "
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
      Left            =   7065
      TabIndex        =   120
      Top             =   6263
      Width           =   1065
   End
   Begin VB.Label lblCgstRate 
      BackStyle       =   0  'Transparent
      Caption         =   "C GST @  : "
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
      Left            =   7080
      TabIndex        =   119
      Top             =   6698
      Width           =   1065
   End
   Begin VB.Label lblIgstRate 
      BackStyle       =   0  'Transparent
      Caption         =   "I  GST @  : "
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
      Left            =   7065
      TabIndex        =   118
      Top             =   7148
      Width           =   1065
   End
   Begin VB.Label lblMsOms 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Ms/Oms"
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
      Left            =   6780
      TabIndex        =   117
      Top             =   1380
      Width           =   735
   End
   Begin VB.Label lblTinNo 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Tin No"
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
      Left            =   1350
      TabIndex        =   114
      Top             =   1425
      Width           =   2775
   End
   Begin VB.Label Label21 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "GSTIN  :  "
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
      Left            =   405
      TabIndex        =   113
      Top             =   1425
      Width           =   900
   End
   Begin VB.Label lblNature 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Nature"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C000C0&
      Height          =   240
      Left            =   15705
      TabIndex        =   112
      Top             =   8055
      Visible         =   0   'False
      Width           =   3225
   End
   Begin VB.Label lblDescription 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Description"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C000C0&
      Height          =   240
      Left            =   15615
      TabIndex        =   111
      Top             =   8280
      Visible         =   0   'False
      Width           =   3495
   End
   Begin VB.Label lbl1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Tr.  Code          : "
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C000C0&
      Height          =   225
      Index           =   0
      Left            =   15705
      TabIndex        =   110
      Top             =   7920
      Visible         =   0   'False
      Width           =   1290
   End
   Begin VB.Label Label20 
      BackStyle       =   0  'Transparent
      Caption         =   "Date   :    "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   255
      Left            =   10845
      TabIndex        =   108
      Top             =   653
      Width           =   600
   End
   Begin VB.Label Label19 
      BackStyle       =   0  'Transparent
      Caption         =   "Date   :    "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   255
      Left            =   10890
      TabIndex        =   107
      Top             =   1980
      Width           =   870
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Date   :    "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   255
      Left            =   10890
      TabIndex        =   106
      Top             =   1305
      Width           =   870
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "Bill No     :    "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   255
      Left            =   7830
      TabIndex        =   104
      Top             =   1305
      Width           =   1185
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Vou. No :"
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
      Left            =   7830
      TabIndex        =   103
      Top             =   630
      Width           =   870
   End
   Begin VB.Label Label15 
      BackStyle       =   0  'Transparent
      Caption         =   "D&ue Days       :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   255
      Left            =   7815
      TabIndex        =   102
      Top             =   1935
      Width           =   975
   End
   Begin VB.Label Label31 
      BackStyle       =   0  'Transparent
      Caption         =   "Aft Tax Add   :"
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
      Height          =   285
      Left            =   9210
      TabIndex        =   67
      Top             =   7628
      Width           =   1335
   End
   Begin VB.Label Label24 
      BackStyle       =   0  'Transparent
      Caption         =   "Aft Tax Less  :"
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
      Left            =   9210
      TabIndex        =   68
      Top             =   7988
      Width           =   1470
   End
   Begin VB.Shape Shape5 
      BorderColor     =   &H00EFB17A&
      BorderWidth     =   5
      FillColor       =   &H00EFB17A&
      Height          =   9705
      Left            =   60
      Top             =   -45
      Width           =   14520
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "M&ill        :"
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
      Left            =   15885
      TabIndex        =   58
      Top             =   4185
      Visible         =   0   'False
      Width           =   960
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "&Broker    : "
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
      Left            =   15930
      TabIndex        =   57
      Top             =   3870
      Visible         =   0   'False
      Width           =   870
   End
   Begin VB.Label LblExmptKg 
      BackStyle       =   0  'Transparent
      Caption         =   "Exem/Kg.      :"
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
      Left            =   17100
      TabIndex        =   72
      Top             =   7830
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.Label Label10 
      BackStyle       =   0  'Transparent
      Caption         =   "Sub Amt      :"
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
      Left            =   9210
      TabIndex        =   60
      Top             =   4605
      Width           =   1425
   End
   Begin VB.Label Label11 
      BackStyle       =   0  'Transparent
      Caption         =   "Add              :"
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
      Left            =   9210
      TabIndex        =   62
      Top             =   5055
      Width           =   1380
   End
   Begin VB.Label Label12 
      BackStyle       =   0  'Transparent
      Caption         =   "Less             :"
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
      Left            =   9210
      TabIndex        =   63
      Top             =   5445
      Width           =   1380
   End
   Begin VB.Label Label30 
      BackStyle       =   0  'Transparent
      Caption         =   "Amount        :"
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
      Left            =   9210
      TabIndex        =   70
      Top             =   9270
      Width           =   1410
   End
   Begin VB.Label Label17 
      BackStyle       =   0  'Transparent
      Caption         =   "Taxable Value  :"
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
      Height          =   300
      Left            =   9090
      TabIndex        =   65
      Top             =   5850
      Width           =   1485
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
      Left            =   11925
      TabIndex        =   96
      Top             =   0
      Width           =   2475
   End
   Begin VB.Label Label28 
      BackStyle       =   0  'Transparent
      Caption         =   "ROff              :"
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
      Height          =   285
      Left            =   9210
      TabIndex        =   69
      Top             =   8835
      Width           =   1380
   End
   Begin VB.Label LblExmptAmt 
      BackStyle       =   0  'Transparent
      Caption         =   "&Exempt. Amt :"
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
      Height          =   300
      Left            =   15705
      TabIndex        =   64
      Top             =   7920
      Visible         =   0   'False
      Width           =   1440
   End
   Begin VB.Label lblTaxType 
      BackStyle       =   0  'Transparent
      Caption         =   "&TaxType        : "
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
      Left            =   16020
      TabIndex        =   66
      Top             =   5985
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.Label Label6 
      BackColor       =   &H00EFB17A&
      Caption         =   "     Other  Purchase"
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
      TabIndex        =   94
      Top             =   0
      Width           =   14580
   End
   Begin VB.Shape Shape1 
      FillColor       =   &H00FED8FC&
      FillStyle       =   0  'Solid
      Height          =   7995
      Left            =   13005
      Shape           =   4  'Rounded Rectangle
      Top             =   675
      Width           =   1305
   End
End
Attribute VB_Name = "frmPurchaseOtherGST"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Const tblName As String = "tblPurch"
Const IdField As String = "Vno"
Const OrderField As String = "Vdt,Vno"
Const MaxNo As Long = 72
Const OtherTxtFromNo As Long = 0 '-- Other Txt Ctrl From No
Const OtherTxtToNo As Long = 0 '-- Other Txt Ctrl Less No
Dim rstbl As Recordset '-- TblPurch
Dim rstblSub As Recordset '-- TblPurchSub
Dim varBookMark As Variant
Dim FirstTimeFlag As Boolean
Dim OldgFrmTypeStr As String
Dim FormAction As Integer
Dim frmTypeStr As String
Dim i As Long
Dim Criteria As String
Dim Mon As Integer

Private Sub Cbo_KeyPress(Index As Integer, KeyAscii As Integer)
If Index = 22 Then KeyAscii = 0 '--Tax
End Sub
Private Sub cbo_LostFocus(Index As Integer)
'If Index = 22 And FormAction = vbDataActionAddNew Then '--Tax %
'   Dim wMillCode As String
'   Dim VATRtFromItem As Double
'   wMillCode = GProcGetColumnValue("tblMastAccount", "AcName", txttemp(7), "S", "AcCode", "N")
'   VATRtFromItem = GProcGetColumnValue("tblMastItem", "ItName", msGrid.TextMatrix(1, 3), "S", "ITTMP1", "N", "ItMillcode=" & wMillCode) '--Item Code
'   If VATRtFromItem > 0 Then
'      txt(23) = VATRtFromItem
'   Else
'       txt(23) = GProcGetColumnValue("tblMastNarration", "Narration", Cbo(22), "S", "MastTaxRate", "N")
'   End If
'   txt(23) = Format(txt(23), FStr)
'
'End If
End Sub
Private Sub chkIsRCM_LostFocus()
Dim RcmSaleAcCd As Long
    If chkIsRCM.Value = 1 Then
       RcmSaleAcCd = GProcGetColumnValue("tblMastNarration", "Narration", "G S T", "S", "RCMSaleAc", "N")
       txttemp(54) = GProcGetColumnValue("tblMastAccount", "AcCode", "" & RcmSaleAcCd & "", "N", "AcName", "S")
    Else
       txttemp(54) = ""
    End If
    SetControlEd
    CalAmount
    GenRCMBillNo
End Sub

'-- Cartoon Bag Detail Entry
Private Sub CmdBagDetails_Click()
'gfrmBagInwVNo = txt(0)
'gfrmBagInwVType = frmTypeStr
'frmBagEntry.Show
End Sub
'Dim GridFlag As Boolean
Private Sub cmdBtn_Click(Index As Integer)
'On Error GoTo ErrorRoutine
Dim X As String
Dim j As Long
Dim tmpRst As Recordset
Dim wLogNo As Integer
Dim AMD As String
Select Case Index
Case 0 '---- Previous
     If GProcRstOpen(rstbl, "Select * from " & tblName & " where " & Criteria & "  and " & IdField & " = (Select max(" & IdField & ") from " & tblName & " where " & Criteria & " And " & IdField & "  < " & Val(txt(0)) & ")", "R") > 0 Then ReadFields Else MsgBox "This is First Record:", vbOKOnly + vbInformation
Case 1 '---- Next
      If GProcRstOpen(rstbl, "Select * from " & tblName & " where " & Criteria & "  and   " & IdField & " = (Select min(" & IdField & ") from " & tblName & " where " & IdField & " > '" & txt(0) & "' and " & Criteria & ") and " & Criteria, "R") > 0 Then ReadFields Else MsgBox "This is last Record:", vbOKOnly + vbInformation
Case 2 '---- First
         If GProcRstOpen(rstbl, "Select * from " & tblName & " where " & Criteria & "  and  " & IdField & " = (Select min(" & IdField & ") from " & tblName & " where " & IdField & " <> '' and " & Criteria & ")and " & Criteria, "R") > 0 Then ReadFields
Case 3 '---- Last
'        If GProcRstOpen(rstbl, "Select * from " & tblName & " where " & Criteria & "  and  " & IdField & " = (Select max(" & IdField & ") from " & tblName & " where " & IdField & " <> '' and " & Criteria & ")and " & Criteria, "R") > 0 Then ReadFields Else GProcClearForm Me, rstbl, MaxNo
       If GProcRstOpen(rstbl, "Select * from " & tblName & " where " & Criteria & "  and  " & IdField & " = (Select max(" & IdField & ") from " & tblName & " where vdt = (SELECT MAX(vdt) from " & tblName & " where " & Criteria & " )  and  " & IdField & " <> 0 and  " & Criteria & ")", "R") > 0 Then ReadFields Else GProcClearForm Me, rstbl, MaxNo

Case 4 'Find ' QryMastAccount
     gClsSearch.SearchMultiField "tblPurch,TblMastAccount", "Vno,AcName,Vdt,PurBillNo,PurBillDt", Array("Vno", "Party", "Vdt", "Bill No", "Bill Dt"), Array(txt(0).Width, 4000, 1080, 1000, 1200), " tblPurch.PurAcCrCode=TblMastAccount.AcCode and vtype='" & frmTypeStr & "' and VFirm = '" & gCCode & "'   and VYear=" & gCYear, "", "Vno,Vdt", 0, 0
     If gClsSearch.SearchMultiRetCol(0) <> "" Then
        GProcRstOpen rstbl, "Select * from " & tblName & " where " & IdField & " = " & gClsSearch.SearchMultiRetCol(0) & " and " & Criteria, "R"
        gClsSearch.SearchMultiRetCol(1) = ""
        ReadFields
    End If
Case 6 'Add
    FormAction = vbDataActionAddNew
    Call GProcActivateControls(Me, True)
    SetControlEd
    Call GProcSetButtons(Me, vbDataActionAddNew)
    Call GProcClearForm(Me, rstbl, MaxNo, True, OtherTxtFromNo, OtherTxtToNo)
    If FirstTimeFlag = True Then
        FirstTimeFlag = False
    Else
        mskDt(2).SetFocus
    End If
'    txt(0) = GProcGenerateId("TblPurch", "Vno", "Vtype='" & frmTypeStr & "' and VFirm = '" & gCCode & "'   and VYear=" & gCYear)
'    mskDt(2) = GProcGenerateVDt("tblPurch", "Vdt", "Vtype='" & frmTypeStr & "' and VFirm = '" & gCCode & "'   and VYear=" & gCYear)
    mskDt(2) = Format(Now, "dd/MM/yyyy")
    
    
    '--A/C
'    txttemp(35) = GProcGetColumnValue("tblMastTransactionCode", "TranCode", CStr(gTranCdPY), "S", "TranCode", "S")
    FillCombo
    SetGrid
Case 7 'Modify
    'If mURecModi = True Then
    '    SearchHelp
    '    DoEvents
        If IsRelatedRecord = False Then
            GProcRstOpen rstbl, "Select * from " & tblName & " where " & IdField & " = " & txt(0) & " and " & Criteria, "R"
            ReadFields
            FormAction = vbDataActionUpdate
            Call GProcActivateControls(Me, True)
            SetControlEd
            Call GProcSetButtons(Me, vbDataActionUpdate)
            txt(0).Enabled = False
            mskDt(2).SetFocus
             '--- If Gate Pass Set Ctrl Enable
'            If IsRelatedGatePass = True Then
'                 Call GProcActivateControls(Me, False)
'                 msGrid.Enabled = True '-Grid
'                 txtVou(11).Enabled = True '--Bill No
'                 mskVou(12).Enabled = True '--Bill Dt
'                 For i = 25 To 30 '-Landing Cost
'                    txt(i).Enabled = True
'                 Next i
'                 txt(15).Enabled = True '-Ex amt
'                 txt(24).Enabled = True '-Ass value
'                 Cbo(12).Enabled = True '-Tax Type
'                 txt(13).Enabled = True '-Tax %
'                 txt(16).Enabled = True '-Ex /Kg
'                 txtOS(9).Enabled = True '-Due days
'                ' txt(21).Enabled = True '-Tot amt
'                 txt(18).Enabled = True 'add
'                 txt(19).Enabled = True 'Less
'                 txt(14).Enabled = True 'Tax amt
'                 txt(17).Enabled = True '-Roff
'                 SetControlEd
'            End If
        End If
    'End If
Case 8 'Delete
    'If mURecDel = True Then
      If IsRelatedGatePass = True Then
            MsgBox "You can not Modify/Delete this record.Related Gate Pass is exists.", vbCritical + vbOKOnly, App.Title
            Exit Sub
      End If
      If IsRelatedRecord = False Then
        If MsgBox("Do You Really Delete This Record?", vbQuestion + vbYesNo + vbDefaultButton1, "Confirm Delete") = vbYes Then
           If gCYear >= 2023 And gCIsLog = 1 Then
              gCn.BeginTrans
              X = "Select max(LogNo) from tblPurch_Log where Vno=" & Val(txt(0)) & " and " & Criteria & ""
              j = GProcRstOpen(tmpRst, X, "R")
              If j > 0 Then
                 wLogNo = IIf(IsNull(tmpRst.Fields(0)), 0, tmpRst.Fields(0)) + 1
              Else
                 wLogNo = 1
              End If
              gCn.Execute "Insert into tblPurch_Log select tblPurch.*," & wLogNo & " as LogNo,'D' as LogTp ,'" & gUserName & "' as UserName ,'" & Format(Now, "YYYY/MM/DD hh:mm:ss") & "' as LogDate from tblPurch where Vno=" & Val(txt(0)) & " and " & Criteria & ""
              gCn.Execute "Insert into tblPurchSub_Log select tblPurchSub.*," & wLogNo & " as LogNo,'D' as LogTp ,'" & gUserName & "' as UserName ,'" & Format(Now, "YYYY/MM/DD hh:mm:ss") & "' as LogDate from tblPurchSub where Vno=" & Val(txt(0)) & " and " & Criteria & ""
              gCn.CommitTrans
              tmpRst.Close
           End If
           DeleteRelatedRecord
           gCn.Execute "delete from tblPurchSub where vtype='" & frmTypeStr & "' and vno=" & txt(0) & " and Vfirm = '" & gCCode & "'  and " _
                       & " VYear=" & gCYear
           GProcRstOpen rstbl, "Select * from " & tblName & " where " & IdField & " = " & txt(0) & " and " & Criteria, "O"
           rstbl.Delete
           cmdBtn_Click (3)
        End If
    End If
Case 9 'Print
    Frame1.Caption = "Sales Bill Printing "
    '-- Inv No
    txtPInvno(0) = txt(52).text
    txtPInvno(1) = txt(52).text
    Frame1.Visible = True
    DoEvents
    txtPInvno(0).Enabled = True
    txtPInvno(1).Enabled = True
'    chkInvChallan(0).Enabled = True
'    chkInvChallan(1).Enabled = True
'    chkInvChallan(2).Enabled = True
'    chkInvChallan(3).Enabled = True
'    chkInvChallan(4).Enabled = True
'    ChkCus.Enabled = True
'    chkOff.Enabled = True
'    ChkGod.Enabled = True
'    chkInvChallan(4).Value = 1
'    ChkCus.Value = 1
    txtPInvno(0).SetFocus
    txtPInvno(0).SelStart = Len(txtPInvno(0))



Case 10 'Exit
'    DoEvents
    Unload Me
    Exit Sub
Case 11 'Label
    'frmLabel.Show 1
Case 12 'Optional
Case 13 'Save
        'txt(51) = CDate(Date)
        If Trim(txtGrid) <> "" Then
            msGrid.text = txtGrid
            DoEvents
            txtGrid = ""
        End If
        txtGrid.Visible = False
       ' If FormAction = vbDataActionAddNew Then txt(0) = GProcGenerateId(tblName, IdField, "vtype='" & FrmTypeStr & "' and VYear=" & gCYear)
        
        CalAmount
        FillTxtFromTemp
'        If gfrmHunkStr = "HNK" Then
'           txt(43) = "HNK"
'        Else
'           txt(43) = ""
'        End If
        '-- Save Record
        
        If FormAction = vbDataActionAddNew Then
           Mon = Month(mskDt(2))
           txt(0) = GProcGenerateIdMonthwise("TblPurch", "Vno", Mon, "Vdt", "Vtype='" & frmTypeStr & "' and VFirm = '" & gCCode & "'   and VYear=" & gCYear)
        End If
        If ValidateData = True Then
            gCn.BeginTrans
            GProcRstOpen rstbl, "Select * from " & tblName & " where " & IdField & " = " & txt(0) & " and " & Criteria, "O"
            GProcSaveRecord Me, rstbl, FormAction, MaxNo, OtherTxtFromNo, OtherTxtToNo
            SaveRelatedRecord
            gCn.CommitTrans
            '-- For Log
            If gCYear >= 2023 And gCIsLog = 1 Then
               If FormAction = vbDataActionAddNew Then
                  AMD = "A"
               Else
                  AMD = "M"
               End If
               X = "Select max(LogNo) from tblPurch_Log where Vno=" & Val(txt(0)) & " and " & Criteria & ""
               j = GProcRstOpen(tmpRst, X, "R")
               If j > 0 Then
                  wLogNo = IIf(IsNull(tmpRst.Fields(0)), 0, tmpRst.Fields(0)) + 1
               Else
                  wLogNo = 1
               End If
               gCn.Execute "Insert into tblPurch_Log select tblPurch.*," & wLogNo & " as LogNo,'" & AMD & "' as LogTp ,'" & gUserName & "' as UserName ,'" & Format(Now, "YYYY/MM/DD hh:mm:ss") & "' as LogDate from tblPurch where Vno=" & Val(txt(0)) & " and " & Criteria & ""
               gCn.Execute "Insert into tblPurchSub_Log select tblPurchSub.*," & wLogNo & " as LogNo,'" & AMD & "' as LogTp ,'" & gUserName & "' as UserName ,'" & Format(Now, "YYYY/MM/DD hh:mm:ss") & "' as LogDate from tblPurchSub where Vno=" & Val(txt(0)) & " and " & Criteria & ""
               tmpRst.Close
            End If
            FormAction = vbDataActionClose
            Call GProcActivateControls(Me, False)
            Call GProcSetButtons(Me, True)
            msGrid.Col = 0 '-----
            msGrid.Row = 1 '-----
            txtGrid.Move msGrid.CellLeft + msGrid.Left, msGrid.CellTop + msGrid.Top, msGrid.CellWidth, msGrid.CellHeight '-----
            If FormAction = vbDataActionAddNew Then
                If cmdBtn(6).Visible = True Then cmdBtn(6).SetFocus
            Else
                If cmdBtn(7).Visible = True Then cmdBtn(7).SetFocus
            End If
    End If
Case 14 'Cancel
    cmdBtn_Click (3)
    txtGrid.Visible = False
    If rstbl.RecordCount <= 0 Then
        Call GProcActivateControls(Me, False)
        Call GProcSetButtons(Me, vbDataActionCancel)
    Else
        Call GProcActivateControls(Me, False)
        Call GProcSetButtons(Me, True)
    End If
    If FormAction = vbDataActionAddNew Then
        If cmdBtn(6).Visible = True Then cmdBtn(6).SetFocus
    Else
        If cmdBtn(7).Visible = True Then cmdBtn(7).SetFocus
    End If
    FormAction = vbDataActionCancel
    msGrid.Col = 0
    msGrid.Row = 1
    txtGrid.Move msGrid.CellLeft + msGrid.Left, msGrid.CellTop + msGrid.Top, msGrid.CellWidth, msGrid.CellHeight
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
Private Sub SetControlEd()
txt(16).Enabled = False '-- Amt
txt(21).Enabled = False '-- Ass Val
mskDt(32).Enabled = False '--Due Dt
txtCrBalance.Enabled = False
txttemp(39).Enabled = False
txt(62).Enabled = False
txt(64).Enabled = False

If FormAction = vbDataActionAddNew Then
   txt(0).Enabled = False
End If
txt(56).Enabled = False
txt(58).Enabled = False
txt(60).Enabled = False
mskDt(53).Enabled = False
If chkIsRCM.Value = 1 Then
  txt(23).Enabled = False
  txt(24).Enabled = False
  txt(26).Enabled = False
  txt(27).Enabled = False
  txt(37).Enabled = False
  txt(38).Enabled = False
  txt(55).Enabled = True
  txt(57).Enabled = True
  txt(59).Enabled = True
  txt(52).Enabled = True
Else
  txt(23).Enabled = True
  txt(24).Enabled = True
  txt(26).Enabled = True
  txt(27).Enabled = True
  txt(37).Enabled = True
  txt(38).Enabled = True
  txt(55).Enabled = False
  txt(57).Enabled = False
  txt(59).Enabled = False
  txt(52).Enabled = False
End If
If gCYear >= 2023 Then
   txt(72).Visible = True
   txt(72).Enabled = True
   Label19.Visible = True
Else
   txt(72).Visible = False
   Label43.Visible = False
End If

End Sub
Private Sub DeleteRelatedRecord()
''--- Voucher Sub
'gCn.Execute "delete from tblPurchSub where vtype='" & frmTypeStr & "' and vno=" & txt(0) & " and Vfirm = '" & gCCode & "'  and " _
'& " VYear=" & gCYear

''-- OutStanding
'gCn.Execute "delete from tblOutStanding where vtype='" & frmTypeStr & "' and vno=" & txt(0) & " and " _
'& " VYear=" & gCYear
''-- Voucher
'gCn.Execute "delete from tblVoucher where vtype='" & frmTypeStr & "' and vno=" & txt(0) & " and " _
'& " VYear=" & gCYear
'--- JV
gCn.Execute "delete from tblVoucher where vtype='JV' and vno=" & txt(62) & " and Vfirm = '" & gCCode & "'  and " _
& " VYear=" & gCYear
End Sub
Private Sub SaveRelatedRecord()
Dim i As Long
Dim MillCode As Long
Dim rstTmp As Recordset
Dim CrCode As Long
Dim DrCode As Long
Dim BrkCode As Long
Dim TaxAcCode As Long
Dim JVTDSVNo As Long
Dim X As String, wit As Long, wgd As Long, Y As String, maxy As Long
Dim BillDt
Dim DueDt
'--- Account Code
MillCode = GProcGetColumnValue("TblMastAccount", "AcName", "-", "S", "Accode", "N")

CrCode = GProcGetColumnValue("tblMastAccount", "AcName", txttemp(5), "S", "AcCode", "N")
DrCode = GProcGetColumnValue("tblMastAccount", "AcName", txttemp(6), "S", "AcCode", "N")
BrkCode = GProcGetColumnValue("tblMastAccount", "AcName", txttemp(8), "S", "AcCode", "N")
TaxAcCode = GProcGetColumnValue("tblMastNarration", "Narration", Cbo(22), "S", "MastTaxPurAcCode", "N")
'-- Delete Records
DeleteRelatedRecord
'--- Save Grid Records
maxy = GProcGetColumnValue("tblPurchSub", "Vno", txt(0), "N", "max(PurSubItCtrlNo)", "N", Criteria)
With msGrid
    i = 1
    gCn.Execute "update TblPurchSub set pursubitsrno = pursubitsrno+100 where " & Criteria & " and vno=" & txt(0)
    Y = ""
    Do While i < .rows
       wit = Val(GProcGetColumnValue("tblMastItem", "ItName", .TextMatrix(i, 1), "S", "ItCode", "N", " ItMillCode=" & MillCode))
'       wgd = Val(GProcGetColumnValue("tblMastNarration", "Narration", .TextMatrix(i, 4), "S", "NarrCode", "N", ""))
       If Val(.TextMatrix(i, 7)) = 0 Then
           X = Val(txt(0)) & ",'" & txt(1) & "','" & Format(txt(2), "yyyy/mm/dd") & "'," & Val(txt(3)) & ",'" & gCCode & "'," & i & ","
           X = X & wit & ","
           X = X & Val(.TextMatrix(i, 2)) & "," & Val(.TextMatrix(i, 3)) & "," & Val(.TextMatrix(i, 4)) & "," & Val(.TextMatrix(i, 5)) & "," & Val(.TextMatrix(i, 6)) & ",1"
          gCn.Execute "Insert into TblPurchSub(vno,vtype,vdt,vyear,vfirm,pursubitsrno,PurSubItCode,PurSubBag,PurSubwt,PurSubRt,PurSubRtPer,PurSubAmt,PurSubGodown)" _
            & " Values(" & X & ")"
       Else
          Y = Y & Val(.TextMatrix(i, 7)) & ","
          gCn.Execute "Update TblPurchSub set pursubitsrno=" & i & ", vdt ='" & Format(txt(2), "yyyy/mm/dd") & "'   ," _
                & "PurSubItCode=" & wit & ",PurSubBag=" & Val(.TextMatrix(i, 2)) & ",PurSubwt=" & Val(.TextMatrix(i, 3)) & ",PurSubRt=" & Val(.TextMatrix(i, 4)) & ",PurSubRtPer=" & Val(.TextMatrix(i, 5)) & ",PurSubAmt=" & Val(.TextMatrix(i, 6)) & " where PurSubItCtrlNo=" & Val(.TextMatrix(i, 7))
       End If
       i = i + 1
    Loop
End With
If Len(Y) >= 2 Then
   Y = Left(Y, Len(Y) - 1)
   gCn.Execute "delete from TblPurchSub where PurSubItCtrlNo <= " & maxy & " and PurSubItCtrlNo not in (" & Y & ") and " & Criteria & " and vno=" & txt(0)
End If

'-- For TDS
If FormAction = vbDataActionAddNew And Val(txt(64)) > 0 Then
           Mon = Month(mskDt(2))
           JVTDSVNo = GProcGenerateIdMonthwise("TblVoucher", "Vno", Mon, "Vdt", "Vtype='JV' and VFirm = '" & gCCode & "'   and VYear=" & gCYear)
Else
   If Val(txt(64)) > 0 Then
      If Val(txt(62)) > 0 And GProcGetColumnValue("tblVoucher", "Vno", txt(62), "N", "Vno", "N", " Vtype='JV' and  VFirm = '" & gCCode & "' and Vyear=" & gCYear) = 0 Then
         JVTDSVNo = CLng(txt(62))
      Else
         Mon = Month(mskDt(2))
        JVTDSVNo = GProcGenerateIdMonthwise("TblVoucher", "Vno", Mon, "Vdt", "Vtype='JV' and VFirm = '" & gCCode & "'   and VYear=" & gCYear)
      End If
      
   End If
End If
If Val(txt(64)) > 0 Then
        X = "select * from tblVoucher where vtype= 'JV' and vno=" & JVTDSVNo & " and Vfirm = '" & gCCode & "' and " _
        & " VYear=" & gCYear
        GProcRstOpen rstTmp, X, "O"
         rstTmp.AddNew
         rstTmp!Vno = JVTDSVNo  '--VNo
         rstTmp!Vtype = "JV"  '--Vtype
         rstTmp!Vdt = CDate(txt(2)) '--Vdt(2)
         rstTmp!VYear = txt(3) '--VYear
         rstTmp!VFirm = gCCode '--VFirm
         rstTmp!VCtrNo = 1 '--Sr No
         rstTmp!VDrAcCode = CrCode
          rstTmp!VCrAcCode = txt(65) 'gTDSAcCodePay
          rstTmp!vamt = Val(txt(64))
            rstTmp!VNar1 = "T.D.S. On Rs. " + txt(15) + " @ " + txt(63) + " %"
            rstTmp!VNar2 = "Against Bill No " + txt(9) + " Dtd. " + txt(10)
            rstTmp!VIsRefEntType = 1
'            rstTmp!VIsAudited = Val(txt(19))
'            rstTmp!VBillNo = txt(9) & vbNullString
'            rstTmp!VBillVno = Val(txt(0))
'            rstTmp!VBillType = "PO"
'            If IsDate(txt(2)) = True Then
'               rstTmp!VBillDate = CDate(txt(2))
'            Else
'               rstTmp!VBillDate = Null
'            End If
'            rstTmp!VBillVYear = gCYear
'            rstTmp!VBillAmt = Val(txt(16))
            rstTmp!VTopCrDr = 1
            rstTmp.Update
   gCn.Execute "update tblPurch set PurTdsJvNo = " & JVTDSVNo & " where vtype='" & frmTypeStr & "' and vno=" & txt(0) & " and  VFirm = '" & gCCode & "'  and    " _
               & " VYear=" & gCYear
End If
If Val(txt(64)) = 0 Then
   gCn.Execute "update tblPurch set PurTdsJvNo = 0 where vtype='" & frmTypeStr & "' and vno=" & txt(0) & " and  VFirm = '" & gCCode & "'  and    " _
               & " VYear=" & gCYear
End If


'-------------------------------




End Sub

Private Sub cmdPrintBill_Click(Index As Integer)
Dim RsSA As Recordset
Dim X As String
Dim XX As String
Dim BillSrNo As String
Dim InvNo As Long
Dim ToInvNo As Long
Dim Y As String
Dim BilNo As String
Dim YY As String
Dim tmpVno As Long
Dim tmpVtype As String
Dim tmpVyear As Long
Dim tmpVFirm As String
Dim GimBl As String
Dim GstinDepot As String
Dim PanDepot As String
Dim PrintFor As String
Dim PayTo As String
Select Case Index
Case 0, 1:
    If Index = 0 Then GProcCrystalRptPreparation cryReport, "Printer"
    If Index = 1 Then GProcCrystalRptPreparation cryReport, "Window"
    If Index = 3 Then GProcCrystalRptPreparation cryReport, "Window"
    
    cryReport.WindowTitle = "RCM Bill Print"
    '-- A/C Bill SrNo
    If Val(txtPInvno(0)) > 0 Then
       InvNo = CLng(txtPInvno(0))
       ToInvNo = CLng(txtPInvno(1))
       BillSrNo = ""
    Else
       InvNo = CLng(Mid(txtPInvno(0), 5, 15))
       ToInvNo = CLng(Mid(txtPInvno(1), 5, 15))
       BillSrNo = Left(txtPInvno(0), 4)
    End If
    '------ Validate
    If IsNumeric(InvNo) = False Then InvNo = 0
    If IsNumeric(ToInvNo) = False Then ToInvNo = InvNo
    If ToInvNo < InvNo Then ToInvNo = InvNo
    
    
    
    X = "("
    Do While InvNo <= ToInvNo
       If BillSrNo = "" Then
          X = X + "'" & InvNo & "',"
       Else
             X = X + "'" & BillSrNo & "" & InvNo & "',"
       End If
       InvNo = InvNo + 1
    Loop
    X = Left(X, Len(X) - 1) + ")"
    Y = X
    X = "Select a.Vno,a.Vtype,a.Vdt,a.Vyear,a.VFirm,a.PurAcCrCode,a.PurAcDrCode,a.PurMillCode,a.PurBroker,a.PurRCMBillNo,a.PurRCMBillDt,a.PurTransport,a.PurLorryNo,a.PurBillNo,a.PurBillDt, " _
        & "a.PurSubAmt,a.PurBillAmt,a.PurAdd1,a.PurLess1,a.PurExemptAmt,a.PurExemptPerKg,a.PurTaxableAmt,a.PurTaxCode,a.PurRCMCgstRt,a.PurRCMCgstAmt,0,a.PurRCMSgstRt,a.PurRCMSgstAmt, " _
        & " a.PurAftTaxAdd1,a.PurAftTaxLess1,a.PurRoff,a.DueDays,a.DueDate,a.PurEntDt,'" & gUserName & "',a.PurTranCd,a.PurIsHank,a.PurRCMIgstRt, " _
        & " a.PurRCMIgstAmt,0,0,0,0,0,b.PurSubItCode,b.PurSubBag,b.PurSubWt,b.PurSubRt,b.PurSubNetRate,b.PurSubRtPer,b.PurSubAmt, " _
        & " 0,b.vdt,b.PurSubGodown,b.PurSubLotNo,a.PurTmpChr1 from tblPurch A ,tblPurchSub B" _
        & " where a.vno = b.vno and a.vtype = b.vtype and a.vyear = b.vyear and a.vfirm = b.vfirm and  a.PurRCMBillNo in " & X & " and a.Vtype='" & frmTypeStr & "' and a.VFirm = '" & gCCode & "' And a.VYear = " & gCYear
     
   gCn.Execute "Delete from tmpSale where  SlUser='" & gUserName & "'"
     
   gCn.Execute " insert into tmpSale (Vno,Vtype,Vdt,Vyear,VFirm,SlAcDrCode,SlAcCrCode,SlMillCode,SlBroker,SlBillNo,SlBillDt,SlTransport,SlLorryNo,SlLrNo,SlLrDate, " _
               & "SlSubAmt,SlBillAmt,SlAdd1,SlLess1,SlExemptAmt,SlExemptPerKg,SlTaxableAmt,SlTaxCode,SlTaxRate,SlTaxAmt,SlTaxableAmt2,SlTaxRate2,SlTaxAmt2," _
               & " SlAftTaxAdd1,SlAftTaxLess1,SlRoff,DueDays,DueDate,SlEntDt,SlUser,SlTranCd,SlIsHank,SlTaxRate3, " _
               & " SlTaxAmt3,SlCharityRt,SlCharityOn,SlCrDays,SlTaxPayAc,SlIsExempt,SlSubItCode,SlSubBag,SlSubWt,SlSubRt,SlSubNetRate,SlSubRtPer,SlSubItAmt, " _
               & " SlSubBookNo,SlSubBookDt,SlSubGodown,SlTmpChr1,SlTmpChr2) " & X
     With cryReport
        .Formulas(10) = "wCSTNo = '" & gCCST & "'"
        .Formulas(11) = "wBSTNo = '" & gCBST & "'"
        .Formulas(12) = "wGSTNo = '" & gCGStin & "'"
        .Formulas(13) = "wcPan = '" & gCPAN & "'"
        PayTo = gCName
        .Formulas(15) = "wPayTo = '" & PayTo & "'"
        .Formulas(20) = "wcopy = 'Office Copy'"
        
        X = "{tmpSale.SlUser}='" & gUserName & "' "
        .SelectionFormula = X
        PrintFor = chkInvChallan(1).Caption
        .Formulas(14) = "wPrintFor = '" & PrintFor & "'"
        .ReportFileName = gReportPath & "rptRCMBillPrintGst.rpt"
        If Index = 0 Or Index = 1 Then
            .Action = 1
        End If
     End With
Case 2:
    DoEvents
    Frame1.Visible = False
End Select

End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case 13:
            If Me.ActiveControl.Name <> "txtGrid" Then Sendkeys "{TAB}"
        Case 27:
            If (cmdBtn(13).Enabled = False) And txtGrid.Enabled = False Then
                Unload Me
            End If
    End Select
End Sub
Private Sub Form_Load()
    Dim X As String
    Me.Top = 300
    Me.Left = Screen.Width / 2 - Me.Width / 2
    '--- TblAddLess
    frmTypeStr = gfrmTypeStr
    X = "Select * from " & tblName & " where Vtype='" & frmTypeStr & "'  and VFirm  = '" & gCCode & "' and VYear=" & gCYear
    If Trim(OrderField) <> "" Then X = X & " order by " & OrderField
    Criteria = "vtype='" & frmTypeStr & "' and Vfirm = '" & gCCode & "' and  VYear=" & gCYear
'    GProcRstOpen rstbl, X, "O"
    FirstTimeFlag = True
    If Trim(mShowRec) <> "" Then
       GProcRstOpen rstbl, "Select * from " & tblName & " where  " & Criteria & " and  " & IdField & " = " & CLng(Trim(mShowRec)), "R"
       ReadFields
       Call GProcSetButtons(Me, vbDataActionCancel)
       mShowRec = ""
    Else
       GProcRstOpen rstbl, "Select * from " & tblName & " where " & Criteria & " and  " & IdField & " = 0", "R"
       SetGrid
       SetRelatedRS
       cmdBtn_Click (6)
    End If
    '---Grid
    '---- Button Visiblity Set As Per User Previlage
    GProcSetButtonVisiblity Me
'    cmdBtn_Click (6)
End Sub
Private Sub SetRelatedRS()
'Dim X As String
'Set rsOS = New Recordset
'Set rsVou = New Recordset
''---TblVoucher
'X = "Select * from tblVoucher where Vtype='" & frmTypeStr & "' and vno=" & txt(0) & " and VYear=" & gCYear & " order by VCtrNo"
'GProcRstOpen rsVou, X, "O"
''---TblOutStanding
'X = "Select * from tblOutStanding where Vtype='" & frmTypeStr & "' and vno=" & txt(0) & " and VYear=" & gCYear
'GProcRstOpen rsOS, X, "O"
End Sub
Private Sub SetGrid()
Dim X As String
Set rstblSub = New Recordset
X = "select PurSubItSrNo,ItName,PurSubBag,PurSubWt" _
& " ,PurSubRt,PurSubRtPer,PurSubAmt,PurSubItCtrlNo,PurSubPItCtrlNo from tblPurchSub,TblMastItem where " _
& " tblPurchSub.PurSubItCode=TblMastItem.ItCode  " _
& " and tblPurchSub.vtype='" & frmTypeStr & "' and tblPurchSub.vno=" & txt(0) & " and tblPurchSub.VFirm = '" & gCCode & "' and tblPurchSub.VYear=" & gCYear & " order by PurSubItSrNo"
GProcRstOpen rstblSub, X, "R"
With msGrid
   .FormatString = ">Sr|<Item Description                                                  |>Nos          |>Weight           |>Rate                  |>Per    |>Amount                 ||"
    For i = 7 To 8
        .ColWidth(i) = 0
    Next i
    If FormAction = vbDataActionAddNew Then
        .rows = 2
        GridAddNew
    Else
        .rows = 1
    End If
End With
End Sub
Private Sub GridAddNew()
Dim i As Long
With msGrid
    For i = 0 To rstblSub.Fields.Count - 1
        Select Case rstblSub.Fields(i).Type
            Case adInteger, 2, 3 '-- Integer
                .TextMatrix(.Row, i) = 0
            Case adDecimal, 131, adNumeric, adDouble '-- Double
                .TextMatrix(.Row, i) = Format(0, GProcNumberFormat(rstblSub.Fields(i).Precision, rstblSub.Fields(i).NumericScale))
            Case adDate '-- Date
                 .TextMatrix(.Row, i) = ""
            Case Else
                 .TextMatrix(.Row, i) = ""
        End Select
    Next i
    ReNumberGridSrno
End With
End Sub
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
If FormAction = vbDataActionUpdate Then
    MsgBox "You are about to modify the record. Save the record and then exit.", vbInformation + vbOKOnly, Me.Caption
    Cancel = True
    Exit Sub
End If
rstbl.Close
End Sub
Private Sub MSGrid_LostFocus()
CalAmount
End Sub
Private Sub mskDt_GotFocus(Index As Integer)
GProcSelectBox Me.ActiveControl
End Sub
Private Sub mskDt_LostFocus(Index As Integer)
If Index = 2 Then
    If GProcIsDateValid(mskDt(Index)) = False Then
        mskDt(Index).SetFocus
        Exit Sub
    End If
End If
'If Index = 2 Then ShowAccountBalance
If IsDate(mskDt(Index)) = False Then mskDt(Index) = "__/__/____"
If Index = 2 And IsDate(mskDt(10)) = False Then
   mskDt(10) = mskDt(2)
End If
End Sub
Private Sub ShowAccountBalance()
'Dim DrAcBal  As Double
Dim CrAcBal As Double
If IsDate(mskDt(2)) Then
'    DrAcBal = GProcGetAccountBalance(txttemp(4), mskDt(2))
    CrAcBal = GProcGetAccountBalance(txttemp(5), mskDt(2))
'    txtDrBalance = IIf(DrAcBal > 0, DrAcBal & " Cr", Abs(DrAcBal) & " Dr")
    txtCrBalance = IIf(CrAcBal > 0, CrAcBal & " Cr", Abs(CrAcBal) & " Dr")
End If
End Sub
Private Sub txt_GotFocus(Index As Integer)
GProcSelectBox Me.ActiveControl
If Index = 52 Then
   GenRCMBillNo
End If
End Sub
Private Sub CheckForNumber()
GProcCheckForNumber Me, rstbl, MaxNo, Array(52), pOtherTxtFromNo:=OtherTxtFromNo, pOtherTxtToNo:=OtherTxtToNo
End Sub
Private Sub txt_LostFocus(Index As Integer)
SetControlEd
CheckForNumber
SaveBtnEd
CalAmount
End Sub
Private Sub txt_KeyPress(Index As Integer, KeyAscii As Integer)
If KeyAscii = 13 Then Exit Sub
KeyAscii = GProcValidateKey(rstbl, Index, KeyAscii, txt(Index).text)
End Sub
Private Sub txtCode_LostFocus(Index As Integer)
If Index = 5 Then
   txttemp(5) = GProcGetColumnValue("tblMastAccount", "AcOurCode", UCase(txtCode(5)), "S", "AcName", "S")
   lblTinNo.Caption = GProcGetColumnValue("TblMastAccount", "ACName", txttemp(Index), "S", "AcGSTIN", "S")
   If Trim(txttemp(5)) = "" Then
      txttemp(5) = ""
      txttemp(5).Enabled = True
   End If
   txttemp(5).SetFocus
End If
If Index = 65 Then
   txttemp(65) = GProcGetColumnValue("tblMastAccount", "AcOurCode", UCase(txtCode(65)), "S", "AcName", "S")
   If Trim(txttemp(65)) = "" Then
      txttemp(65) = ""
      txttemp(65).Enabled = True
   End If
   txttemp(65).SetFocus
End If
End Sub

Private Sub txttemp_GotFocus(Index As Integer)
Dim GpCd As Long
Dim X As String
Dim rstblBlChk As Recordset
GProcSelectBox Me.ActiveControl
If Index = 67 And (txttemp(67) = "" Or txttemp(67) = "-") And txttemp(6) <> "" Then
   GpCd = GProcGetColumnValue("tblMastAccount", "AcName", "" & txttemp(6) & "", "S", "AcCode", "N")
   X = "select * from tblMastSubGroup where SubGpCode = " & GpCd & ""
   i = GProcRstOpen(rstblBlChk, X, "R")
   If i > 0 Then
      gClsSearch.SearchMultiField "tblMastSubGroup", "SubOurCode,SubName", Array("Code", "Sub A/c Name"), Array(2000, 5000), "SubGpCode = " & GpCd & "", "", "SubName", 0, 0, False
      If gClsSearch.SearchMultiRetCol(0) <> "" Then
         txttemp(67) = gClsSearch.SearchMultiRetCol(1)
      End If
   End If
End If
End Sub
Private Sub txttemp_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
Dim GName As String
Dim OldFrmType As String
If KeyCode = vbKeyF3 Then
    Select Case Index
        Case 5 '-- Party
            GName = GProcGetColumnValue("TblMastGroup", "AgCode", 90015, "N", "AgName", "S")
            GProcShowForm frmMastAccount, frmMain.mnuMstfrm(0), Array(frmMastAccount.txttemp(2)), Array(GName)
        Case 6 '-- Purchase A/c
            GName = GProcGetColumnValue("TblMastGroup", "AgCode", 90025, "N", "AgName", "S")
            GProcShowForm frmMastAccount, frmMain.mnuMstfrm(0), Array(frmMastAccount.txttemp(2)), Array(GName)
        Case 7 '-- millcode
            GName = GProcGetColumnValue("TblMastGroup", "AgCode", 90029, "N", "AgName", "S")
            GProcShowForm frmMastAccount, frmMain.mnuMstfrm(0), Array(frmMastAccount.txttemp(2)), Array(GName)
        Case 8 '-- Broker
           GName = GProcGetColumnValue("TblMastGroup", "AgCode", 90016, "N", "AgName", "S")
           GProcShowForm frmMastAccount, frmMain.mnuMstfrm(0), Array(frmMastAccount.txttemp(2)), Array(GName)
        Case 12 '--TaxType
            'OldgFrmTypeStr = gfrmTypeStr
            GProcShowForm frmMastNarrationTax, frmMain.mnuMstfrm(6)
            'gfrmTypeStr = OldgFrmTypeStr
        Case 11 '--Narration transport
            OldFrmType = gfrmTypeStr
            gfrmTypeStr = "R"
            GProcShowForm frmMastNarration, frmMain.mnuMstfrm(8), gfrmTypeStr
            gfrmTypeStr = OldFrmType
    End Select
End If
Dim LeftPos As Long
Dim TopPos As Long
If KeyCode = vbKeyF2 Then '-- Select
    TopPos = txttemp(Index).Top + Me.Top + 650
    LeftPos = txttemp(Index).Left + Me.Left
    Select Case Index
        Case 11 '--Narration transport
            gClsSearch.SearchMultiField "tblMastNarration", "Narration", Array("Narration"), Array(txttemp(Index).Width), "NarrType='R'", "", "Narration", LeftPos, TopPos, False
            If gClsSearch.SearchMultiRetCol(0) <> "" Then txttemp(Index) = gClsSearch.SearchMultiRetCol(0)
    End Select
End If
End Sub
Private Sub txtTemp_KeyPress(Index As Integer, KeyAscii As Integer)
Dim LeftPos As Long
Dim TopPos As Long
Dim AgCode As Long
Dim BrkAcCode As Long
Dim GpCd As Long
Dim X As String
Dim rstblBlChk As Recordset
TopPos = txttemp(Index).Top + Me.Top + 650
LeftPos = txttemp(Index).Left + Me.Left
If KeyAscii = 13 Then Exit Sub
Select Case Index
    Case 5  '-- Party   Alias Account
'        gClsSearch.SearchMultiField "tblMastaccount,tblMastGroup", "AcName,AgName,AcContPer,AcCity,AcGSTIN,AcMsOmsParty,AcOurCode", Array("Party", "Group", "Contact Person", "City", "TIN", "Ms/Oms", ""), Array(5000, 3000, 3000, 2000, 0, 0, 0), "(tblMastaccount.Agcode in (90015,90029) ) and tblMastaccount.AgCode = tblMastGroup.AgCode ", Chr(KeyAscii), "", LeftPos, TopPos, True
        gClsSearch.SearchMultiFieldAccMaster Chr(KeyAscii), txttemp(Index).Left + Me.Left, txttemp(Index).Top + Me.Top + 1000
        KeyAscii = 0
        lblTinNo.Caption = ""
        lblMsOms.Caption = ""
        If gClsSearch.SearchMultiRetColAccMaster(0) <> "" Then
            txttemp(Index) = gClsSearch.SearchMultiRetColAccMaster(0)
            txtCode(5) = GProcGetColumnValue("TblMastAccount", "ACName", txttemp(Index), "S", "AcOurCode", "S")
            lblTinNo.Caption = GProcGetColumnValue("TblMastAccount", "ACName", txttemp(Index), "S", "AcGSTIN", "S")
            lblMsOms.Caption = GProcGetColumnValue("TblMastAccount", "ACName", txttemp(Index), "S", "AcMsOmsParty", "N")
            If Val(lblMsOms.Caption) = 0 Then
               lblMsOms.Caption = "M.S."
            Else
               lblMsOms.Caption = "O.M.S."
            End If
        End If
'        AgCode = GProcGetColumnValue("tblMastAccount", "AcName", txttemp(5), "S", "AgCode", "N")
        txttemp(7) = "-"
'        If AgCode = 90029 Then
'            txttemp(7) = txttemp(5) '-- Mill =Party
'            txttemp(7).Enabled = False
'        Else
'            txttemp(7).Enabled = True
'        End If
    Case 65  '-- TDS Account
        gClsSearch.SearchMultiFieldAccMaster Chr(KeyAscii), txttemp(Index).Left + Me.Left, txttemp(Index).Top + Me.Top + 1000
        KeyAscii = 0
        If gClsSearch.SearchMultiRetColAccMaster(0) <> "" Then
            txttemp(Index) = gClsSearch.SearchMultiRetColAccMaster(0)
            txtCode(65) = GProcGetColumnValue("TblMastAccount", "ACName", txttemp(Index), "S", "AcOurCode", "S")
        End If
    Case 6 '-- Purchase A/c  Alias Account
        gClsSearch.SearchMultiFieldAccMaster Chr(KeyAscii), txttemp(Index).Left + Me.Left, txttemp(Index).Top + Me.Top + 1000
'        gClsSearch.SearchMultiField "tblMastaccount", "AcName,AcCity", Array("Purchase Account", "City"), Array(3000, 2000), "", Chr(KeyAscii), "", LeftPos, TopPos, True
        KeyAscii = 0
        If gClsSearch.SearchMultiRetColAccMaster(0) <> "" Then
            txttemp(Index) = gClsSearch.SearchMultiRetColAccMaster(0)
        End If
    Case 7  '-- Mill  Alias Account
        gClsSearch.SearchMultiField "tblMastaccount", "AcName,AcCity", Array("Mill Name", "City"), Array(3000, 2000), "(AgCode=90029 ) ", Chr(KeyAscii), "AcName", LeftPos, TopPos, True
        KeyAscii = 0
        If gClsSearch.SearchMultiRetCol(0) <> "" Then
            txttemp(Index) = gClsSearch.SearchMultiRetCol(0)
        End If
        MillCtrlEd
    Case 8 '-- Broker
        gClsSearch.SearchMultiField "tblMastaccount", "AcName,AcCity", Array("Broker", "City"), Array(5000, 3000), "(AgCode=90016 )", Chr(KeyAscii), "", LeftPos, TopPos, True
        KeyAscii = 0
        If gClsSearch.SearchMultiRetCol(0) <> "" Then
            txttemp(Index) = gClsSearch.SearchMultiRetCol(0)
        End If
    Case 35 '-- Transaction code
        gClsSearch.SearchMultiField "tblMastTransactionCode", "TranCode,iif(TranType=0,'Purchase',iif(TranType=1,'Purchase Goods Return',iif(TranType=2,'Purchase Debit Note',iif(TranType=3,'Sales', iif(TranType=4,'Sales Goods Return',iif(TranType=5,'Sales Credit Note', '') ) ) ) )),Nature,Description", Array("Tr. Code", "Tr. Type", "Nature", "Description"), Array(1000, 4000, 4000, 4000), "TranType = 0 and Trade = 1", "", "TranCode", 600 + Me.Left, 2500
        KeyAscii = 0
        If gClsSearch.SearchMultiRetCol(0) <> "" Then
            txttemp(Index) = gClsSearch.SearchMultiRetCol(0)
            lblNature.Caption = GProcGetColumnValue("tblMastTransactionCode", "TranCode", txttemp(35), "S", "Nature", "S")
            lblDescription.Caption = GProcGetColumnValue("tblMastTransactionCode", "TranCode", txttemp(35), "S", "Description", "S")
        End If
     Case 67
        GpCd = GProcGetColumnValue("tblMastAccount", "AcName", "" & txttemp(6) & "", "S", "AcCode", "N")
        X = "select * from tblMastSubGroup where SubGpCode = " & GpCd & ""
        i = GProcRstOpen(rstblBlChk, X, "R")
        If i > 0 Then
           gClsSearch.SearchMultiField "tblMastSubGroup", "SubOurCode,SubName", Array("Code", "Sub A/c Name"), Array(2000, 5000), "SubGpCode = " & GpCd & "", "", "SubName", 0, 0, False
           KeyAscii = 0
           If gClsSearch.SearchMultiRetCol(0) <> "" Then
              txttemp(67) = gClsSearch.SearchMultiRetCol(1)
           End If
        Else
          txttemp(67) = ""
        End If
End Select

If Index = 11 Then '-- Transport
    gClsSearch.SearchMultiField "tblMastNarration", "Narration", Array("Transport"), Array(txttemp(Index).Width), " NarrType='R'", Chr(KeyAscii), "Narration", 0, 0, True
    KeyAscii = 0
    If gClsSearch.SearchMultiRetCol(0) <> "" Then
        txttemp(Index) = gClsSearch.SearchMultiRetCol(0) '--Transport
    End If
End If
End Sub
Private Sub MillCtrlEd()
Dim ExemptMill As Long
'-- Is Exempt Mill
ExemptMill = GProcGetColumnValue("TblMastAccount", "AcName", txttemp(7), "S", "ACIsExemptMill", "N")
If ExemptMill = 1 Then
    If FormAction = vbDataActionAddNew Or FormAction = vbDataActionUpdate Then
        txt(19).Enabled = True
        txt(20).Enabled = True
        LblExmptAmt.Visible = True
        LblExmptKg.Visible = True
        txt(19).Visible = True
        txt(20).Visible = True
    End If
Else
    txt(19).Enabled = False
    txt(20).Enabled = False
    LblExmptAmt.Visible = False
    LblExmptKg.Visible = False
    txt(19).Visible = False
    txt(20).Visible = False
    txt(19) = 0
    txt(20) = 0
End If
End Sub
Private Sub txtTemp_LostFocus(Index As Integer)
SaveBtnEd
MillCtrlEd
ShowAccountBalance
End Sub
'---------------- Grid
Private Sub msGrid_Click()
'On Error Resume Next
    msGrid_EnterCell
End Sub
Private Sub msGrid_EnterCell()
If FormAction <> vbDataActionClose Then
    CheckForNumberTxtGrid
    DoEvents
    txtGrid = ""
    DoEvents
    If msGrid.text <> "" Then txtGrid.text = msGrid.text
    '--- Lock TextBox
    '--- Check is Related Gate Pass
'    If IsRelatedItCtrlGatePass(msGrid.TextMatrix(msGrid.Row, 14)) = False Then
        With msGrid
            Select Case .Col
                Case 0: '---- Sr No
                    txtGrid.Locked = True
                    ReNumberGridSrno
                    txtGrid = ""
                    .Col = .Col + 1
                Case Else
                    txtGrid.Locked = False
            End Select
        End With
'    Else '-Related Gate Pass
 '         txtGrid.Locked = True
  '  End If
End If
msGrid_KeyPress (0)
End Sub
Private Sub msGrid_KeyPress(KeyAscii As Integer)
'--- Check is Related Gate Pass
'If IsRelatedItCtrlGatePass(msGrid.TextMatrix(msGrid.Row, 14)) = True Then Exit Sub
With msGrid
    If KeyAscii = 27 Then '----Esc Key
        KeyAscii = 0
        If ((Val(.TextMatrix(.Row, 6)) = 0)) And .rows > 2 Then
            .RemoveItem (.Row)
            txtGrid.text = ""
        End If
        txtGrid.Visible = False
        txttemp(6).SetFocus
        Exit Sub
    End If
End With
'If FormAction <> vbDataActionAddNew And FormAction <> vbDataActionUpdate Then Exit Sub
'---- Navigate Row ,Col
If KeyAscii = 13 Then
    KeyAscii = 0
    With msGrid
        If .Col = 5 Then 'rstblSub.Fields.Count - 2 Then
            If .Row = (.rows - 1) Then '-- Add New Row
                .rows = .rows + 1
                .Row = .Row + 1
                GridAddNew
            Else '--Next Row
                .Row = .Row + 1
            End If
            .Col = 1
        Else    '-- Next col
            .Col = .Col + 1
        End If
    End With
End If
'---- Set TxtGrid


If msGrid.text <> "" Then txtGrid.text = msGrid.text
If msGrid.Col < rstblSub.Fields.Count - 1 Then
    txtGrid.Move msGrid.CellLeft + msGrid.Left, msGrid.CellTop + msGrid.Top, msGrid.CellWidth, msGrid.CellHeight
End If
txtGrid.Visible = True
txtGrid.Enabled = True
DoEvents
If KeyAscii <> 0 Then
    Select Case rstblSub.Fields(msGrid.Col).Type
        Case adChar, adVarWChar
            If msGrid.Col = 1 Then  '-- Count
                txtGrid_KeyPress (KeyAscii)
            Else
                txtGrid.text = txtGrid.text & Chr(KeyAscii)
            End If
        Case Else
            txtGrid.text = Chr(GProcValidateKey(rstblSub, msGrid.Col, KeyAscii, txtGrid.text)) 'Chr(GProcNumberWithDecimal(KeyAscii, txtGrid))
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
If txtGrid.Visible Then
    msGrid.text = txtGrid.text
    txtGrid.text = ""
    txtGrid.Visible = False
End If
DoEvents
'If msGrid.Col >= 13 Then
    CalGridAmt
    CalAmount
'End If
End Sub
Private Sub CalGridAmt()
Dim MillCode As Long
Dim StdPack As Double
CheckForNumberGrid
MillCode = GProcGetColumnValue("TblMastAccount", "AcName", txttemp(5), "S", "Accode", "N")
If msGrid.Row = 0 Then Exit Sub

'-- Amt=Wt*Rate/Rate Per
If CDbl(msGrid.TextMatrix(msGrid.Row, 5)) > 0 Then
    msGrid.TextMatrix(msGrid.Row, 6) = CDbl(msGrid.TextMatrix(msGrid.Row, 3)) * CDbl(msGrid.TextMatrix(msGrid.Row, 4)) / CDbl(msGrid.TextMatrix(msGrid.Row, 5))
Else
    msGrid.TextMatrix(msGrid.Row, 6) = 0
End If
msGrid.TextMatrix(msGrid.Row, 6) = Format(msGrid.TextMatrix(msGrid.Row, 6), LStr)
End Sub
Private Sub CheckForNumberGrid()
Dim RowIndex As Long
With rstblSub
For RowIndex = 1 To msGrid.rows - 1
    For i = 0 To .Fields.Count - 1
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
With rstblSub
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
       ' Case adChar, adVarWChar '--String
       '     txtGrid.MaxLength = .Fields(msGrid.Col).DefinedSize
        Case adDate '-- Date
            If IsDate(txtGrid) = False Then
                txtGrid = vbNullString
            Else
                txtGrid = CDate(txtGrid)
            End If
            txtGrid.MaxLength = 10
         Case Else
            txtGrid.MaxLength = .Fields(msGrid.Col).DefinedSize
    End Select
End With
End Sub
Private Sub txtGrid_GotFocus()
   txtGrid.Alignment = ProcSetAlignment(rstblSub, msGrid.Col)
    DoEvents
    Call GProcSelectBox(txtGrid)
End Sub
Private Function ProcSetAlignment(pRstbl As Recordset, pColIndex As Long)
Select Case pRstbl.Fields(pColIndex).Type
    Case adInteger, adDouble, 2, 3, 131, adDecimal
        ProcSetAlignment = 1
    Case Else
        ProcSetAlignment = 0
End Select
End Function
Private Sub txtGrid_KeyDown(KeyCode As Integer, Shift As Integer)
'--- Check is Related Gate Pass
'If IsRelatedItCtrlGatePass(msGrid.TextMatrix(msGrid.Row, 14)) = True Then Exit Sub
'--- Delete Row On Ctrl+Y
If KeyCode = vbKeyY And Shift = 2 And msGrid.rows > 2 Then
    msGrid = 1
    msGrid.RemoveItem (msGrid.Row)
    txtGrid.Visible = False
    txtGrid.text = ""
    ReNumberGridSrno
    CalAmount
    Exit Sub
End If
'--- Show Form
If KeyCode = vbKeyF3 Then
Select Case msGrid.Col
    Case 1 '-- Item
        GProcShowForm frmMastItem, frmMain.mnuMstfrm(3), Array(frmMastItem.txttemp(9)), Array(txttemp(7))
        Exit Sub
'    Case 4 '-- Godown
'        OldgFrmTypeStr = gfrmTypeStr
'        gfrmTypeStr = "G"
'        GProcShowForm frmMastNarration, frmMain.mnuMstfrm(8), "G"
'        gfrmTypeStr = OldgFrmTypeStr
End Select
End If
End Sub
'-- ReNumber Grid Sr No
Private Sub ReNumberGridSrno()
Dim i As Long
With msGrid
    For i = 1 To .rows - 1
        .TextMatrix(i, 0) = i
    Next i
End With
End Sub
Private Sub txtGrid_KeyPress(KeyAscii As Integer)
Dim LeftPos As Long
Dim TopPos As Long
Dim MillCode As Long
TopPos = txtGrid.Top + Me.Top + 650
LeftPos = txtGrid.Left + Me.Left
DoEvents
'--- Check is Related Gate Pass
'If IsRelatedItCtrlGatePass(msGrid.TextMatrix(msGrid.Row, 14)) = True Then Exit Sub
With msGrid
    If KeyAscii <> 13 And KeyAscii <> 27 And KeyAscii <> 25 Then
        '--- Search
        Select Case .Col
            Case 1 '-- Count
'                If msGrid.TextMatrix(msGrid.Row, 1) <> "" Then
 '                   KeyAscii = 0
  '              Else
                    MillCode = GProcGetColumnValue("TblMastAccount", "AcName", "-", "S", "Accode", "N")
                    gClsSearch.SearchMultiField "tblMastItem,tblMastaccount", "ItName", Array("Item Name"), Array(5000), "tblMastItem.ITMillCode=tblMastaccount.acCode and ITMillCode=" & MillCode, Chr(KeyAscii), "ItName", LeftPos, TopPos, True
                    txtGrid = ""
                    KeyAscii = 0
                    If gClsSearch.SearchMultiRetCol(0) <> "" Then '-- Item Wt ,Rate From Master
                        txtGrid = gClsSearch.SearchMultiRetCol(0)
                        If .TextMatrix(.Row, 5) = 0 Then .TextMatrix(.Row, 5) = 1
                        txtGrid_KeyPress 13
                    End If
   '               End If
            Case Else
                KeyAscii = GProcValidateKey(rstblSub, msGrid.Col, KeyAscii, txtGrid)
        End Select
   End If
    If KeyAscii = 13 Then
        CheckForNumberTxtGrid
        If txtGrid = "" And .Col = 1 Then
           KeyAscii = 27
        End If
    End If
    
    '--- Navigate Row ,Col
    If KeyAscii = 13 And .Col < 5 Then 'rstblSub.Fields.Count - 2 Then
        .Col = .Col + 1
    ElseIf KeyAscii = 13 And .Col = 5 Then '= rstblSub.Fields.Count - 2 Then
        If .Row = (.rows - 1) Then
            .rows = .rows + 1
            .Row = .Row + 1
            GridAddNew
        Else
            .Row = .Row + 1
        End If
        .Col = 1
    End If
    '-- On ESC
    If KeyAscii = 27 Then
        KeyAscii = 0
        If (Val(.TextMatrix(.Row, 6)) = 0) And .rows > 2 Then
            .RemoveItem (.Row)
            txtGrid.text = ""
        End If
        txtGrid.Visible = False
        txttemp(6).SetFocus
        Exit Sub
    End If
End With
End Sub

'-- Grid Required
Private Function ValidateGrid() As Boolean
Dim i As Long
Dim MillCode As Long
Dim RowIndex As Long
Dim ItCode As Long
Dim X As String
Dim Rs1 As Recordset
Dim oldFrmTypeStr As String
Dim LRNo As String
Dim pItCtrlNo As Long
If Trim(txtGrid.text) <> "" Then
    msGrid.text = txtGrid.text
    CalGridAmt
End If
txtGrid.Visible = False
txtGrid.text = ""
CheckForNumberGrid
With msGrid
Validate:
    For RowIndex = 1 To .rows - 1
        For i = 0 To rstblSub.Fields.Count - 1
            Select Case i
                Case 1 '-- Count
                 '-Rate=0 ,count,wt=0 then delete row
                    If RowIndex > 1 And (Trim(.TextMatrix(RowIndex, 1)) = "" Or CDbl(.TextMatrix(RowIndex, 3)) = 0 Or CDbl(.TextMatrix(RowIndex, 4)) = 0) Then
                        .RemoveItem (RowIndex)
                        GoTo Validate
                    End If
                    If Trim(.TextMatrix(RowIndex, i)) = "" Then
                        MsgBox "Check Count.", vbCritical + vbOKOnly, Me.Caption
                        .Col = i
                        msGrid.SetFocus
                        Exit Function
                    End If
                    MillCode = GProcGetColumnValue("TblMastAccount", "AcName", "-", "S", "Accode", "N")
                    If GProcGetColumnValue("TblMastItem", "ItName", .TextMatrix(RowIndex, i), "S", "Itcode", "N", "ITMillCode=" & MillCode) = 0 Then
                        MsgBox "Check Count.Mill and Count are Mismatch.", vbCritical + vbOKOnly, Me.Caption
                        .Col = i
                        msGrid.SetFocus
                        Exit Function
                    End If
'                Case 2 '-- Bag
'                    If CDbl(.TextMatrix(RowIndex, i)) = 0 Then
'                        MsgBox "Check Bag.", vbCritical + vbOKOnly, Me.Caption
'                        .Col = i
'                        msGrid.SetFocus
'                        Exit Function
'                    End If
'                    '-- If Bag < Issue Bag
'                    If CDbl(.TextMatrix(RowIndex, 7)) < CDbl(.TextMatrix(RowIndex, 13)) Then
'                        MsgBox "Check Bag.Bag Less than issued.", vbCritical + vbOKOnly, Me.Caption
'                        .Col = i
'                        msGrid.SetFocus
'                        Exit Function
'                    End If
                Case 3 '-- Wt
                    If CDbl(.TextMatrix(RowIndex, i)) = 0 Then
                        MsgBox "Check Weight.", vbCritical + vbOKOnly, Me.Caption
                        .Col = i
                        msGrid.SetFocus
                        Exit Function
                    End If
'                    '-- If Wt < Issue wt
'                    If CDbl(.TextMatrix(RowIndex, 8)) < CDbl(.TextMatrix(RowIndex, 14)) Then
'                        MsgBox "Check Wt.Weight Less than issued.", vbCritical + vbOKOnly, Me.Caption
'                        .Col = i
'                        msGrid.SetFocus
'                        Exit Function
'                    End If
                Case 6 '-- amount
                    If CDbl(.TextMatrix(RowIndex, i)) = 0 Then
                        MsgBox "Check amount.", vbCritical + vbOKOnly, Me.Caption
                        .Col = i
                        msGrid.SetFocus
                        Exit Function
                    End If
           End Select
        Next i
      Next RowIndex
End With
ValidateGrid = True
ReNumberGridSrno
End Function
Private Sub FillCombo()
Call GProcFillCombo(Cbo(22), "select Narration from tblMastNarration where NarrType ='T' order by Narration")
'If gfrmHunkStr = "HNK" Then
'   Cbo(22) = GProcGetColumnValue("tblMastNarration", "NarrCode", CStr(gTaxCodePYHunk), "N", "Narration", "S")
'Else
'   Cbo(22) = GProcGetColumnValue("tblMastNarration", "NarrCode", CStr(gTaxCodePY), "N", "Narration", "S")
'End If
'If FormAction = vbDataActionAddNew And Val(txt(23)) = 0 Then
'    txt(23) = GProcGetColumnValue("tblMastNarration", "Narration", Cbo(22), "S", "MastTaxRate", "N")
'    txt(23) = Format(txt(23), FStr)
'End If
End Sub
Private Sub SaveBtnEd()
'--- Vno,vdt,Party,A/c,Mill,Amount
If (CDbl(txt(0)) = 0 And FormAction <> vbDataActionAddNew) Or GProcIsDateValid(mskDt(2)) = False Or txttemp(5) = "" Or txttemp(6) = "" Or CDbl(txt(16)) = 0 Then
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
Dim rstAddCheck As Recordset
Dim rstModiCheck As Recordset
Dim ExpAcCode As Long
    MillCtrlEd
If FormAction <> vbDataActionAddNew Then
   If Len(txt(0)) = 5 Then
      i = Val(Left(txt(0), 1))
   Else
      i = Val(Left(txt(0), 2))
   End If
   If i <> Month(mskDt(2)) Then
      MsgBox "Cant Change Month Of Voucher Date", vbInformation + vbOKOnly, Me.Caption
      mskDt(2).SetFocus
      Exit Function
   End If
End If
    
    '---Validate Vou.Date = Bill Date
'    If mskdt(10) <> mskDt(2) Then
'        MsgBox "Vou.Date Muse Be Same As Bill Date", vbInformation + vbOKOnly, Me.Caption
'        mskDt(12).SetFocus
'        Exit Function
'    End If
    '--- Validate Grid
    If ValidateGrid = False Then
        'MSGrid.SetFocus
        Exit Function
    End If
'    If GProcValidateTranCd(frmTypeStr, Cbo(22).text, txttemp(35), lblTinNo.Caption, lblMsOms.Caption) = False Then
'       Select Case txttemp(35).text
'              Case "10"
'                   MsgBox "Required Data Tax = 'V A T', Tin No Should Not Balnk, And M.S. Party", vbInformation + vbOKOnly, Me.Caption
'              Case "15"
'                   MsgBox "Required Data Tax = 'U R D', Tin No Should Balnk", vbInformation + vbOKOnly, Me.Caption
'              Case "20"
'                   MsgBox "Required Data Tax = 'U R D' And Tin No Balnk", vbInformation + vbOKOnly, Me.Caption
'              Case "70"
'                   MsgBox "Required Data Tax = 'C S T', Tin No Should Not Balnk And O.M.S. Party", vbInformation + vbOKOnly, Me.Caption
'              Case "40"
'                   MsgBox "Required Data Tax = 'C S T Against Form C', Tin No Should Not Balnk And O.M.S. Party", vbInformation + vbOKOnly, Me.Caption
'       End Select
'       txttemp(35).SetFocus
'       Exit Function
'
'    '---- Required Data
'    End If
    '--- Vno
    If CLng(txt(0)) = 0 And FormAction <> vbDataActionAddNew Then
        MsgBox "Check Voucher No.", vbInformation + vbOKOnly, Me.Caption
        mskDt(2).SetFocus
        Exit Function
    End If
    '--- Vdt
    If GProcIsDateValid(mskDt(2)) = False Then
        MsgBox "Check Voucher Date.", vbInformation + vbOKOnly, Me.Caption
        mskDt(2).SetFocus
        Exit Function
    End If
    '---Tax Type
    If Cbo(22) = "" Then
        MsgBox "Check Tax Type.", vbInformation + vbOKOnly, Me.Caption
        Cbo(22).SetFocus
        Exit Function
    End If
    
    
    '-- Tax %
    If CDbl(txt(23)) > 0 And Cbo(22) = "Tax Free" Then
        MsgBox "Check Tax %.", vbInformation + vbOKOnly, Me.Caption
        txt(23).SetFocus
        Exit Function
    End If
    If chkIsRCM.Value = 1 And txttemp(54) = "" Then
        MsgBox "Check RCM Sale A/c", vbInformation + vbOKOnly, Me.Caption
        txttemp(54).SetFocus
        Exit Function
    End If
    
    '-- Tax Amount
    If (CDbl(txt(24)) > 0 Or CDbl(txt(23)) > 0) And Abs(Round(Val(txt(21)) * Val(txt(23)) / 100, 2) - Val(txt(24))) > 1 Then
        MsgBox "Tax Amount on Ass.Val not correct", vbInformation + vbOKOnly, Me.Caption
        txt(24).SetFocus
        Exit Function
    End If
    '--CGST Tax Amt
    If (CDbl(txt(27)) > 0 Or CDbl(txt(26)) > 0) And Abs(Round(Val(txt(21)) * Val(txt(26)) / 100, 2) - Val(txt(27))) > 1 Then
        MsgBox "Tax Amount on Ass.Val not correct", vbInformation + vbOKOnly, Me.Caption
        txt(27).SetFocus
        Exit Function
    End If
    '--IGST Tax Amt
    If (CDbl(txt(38)) > 0 Or CDbl(txt(37)) > 0) And Abs(Round(Val(txt(21)) * Val(txt(37)) / 100, 2) - Val(txt(38))) > 1 Then
        MsgBox "Tax Amount on Ass.Val not correct", vbInformation + vbOKOnly, Me.Caption
        txt(38).SetFocus
        Exit Function
    End If
    '---- Party Account
    If txttemp(5) = "" Then
        MsgBox "Check Party.", vbInformation + vbOKOnly, Me.Caption
        txttemp(3).SetFocus
        Exit Function
    End If
        '--Create Party
        If GProcGetColumnValue("tblMastAccount", "AcName", txttemp(5), "S", "AcCode", "N") = 0 Then
           GName = GProcGetColumnValue("TblMastGroup", "AgCode", 90015, "N", "AgName", "S")
           GProcShowForm frmMastAccount, frmMain.mnuMstfrm(0), Array(frmMastAccount.txt(1), frmMastAccount.txttemp(2)), Array(txttemp(5), GName)
           Exit Function
        End If
    '---- Purchase Account
    If txttemp(6) = "" Then
        MsgBox "Check Purchase Account.", vbInformation + vbOKOnly, Me.Caption
        txttemp(6).SetFocus
        Exit Function
    End If
        '--Create Party
        If GProcGetColumnValue("tblMastAccount", "AcName", txttemp(6), "S", "AcCode", "N") = 0 Then
           GName = GProcGetColumnValue("TblMastGroup", "AgCode", 90025, "N", "AgName", "S")
           GProcShowForm frmMastAccount, frmMain.mnuMstfrm(0), Array(frmMastAccount.txt(1), frmMastAccount.txttemp(2)), Array(txttemp(6), GName)
           Exit Function
        End If
    '---- Mill
'    If txttemp(7) = "" Then
'        MsgBox "Check Mill.", vbInformation + vbOKOnly, Me.Caption
'        txttemp(7).SetFocus
'        Exit Function
'    End If
        '--Create Mill
'        If GProcGetColumnValue("tblMastAccount", "AcName", txttemp(7), "S", "AcCode", "N") = 0 Then
'           GName = GProcGetColumnValue("TblMastGroup", "AgCode", 90029, "N", "AgName", "S")
'           GProcShowForm frmMastAccount, frmMain.mnuMstfrm(0), Array(frmMastAccount.txt(1), frmMastAccount.txttemp(2)), Array(txttemp(7), GName)
'           Exit Function
'        End If
    '--- Broker
        '--Create Broker
        If txttemp(8) = "" Then txttemp(8) = "-"
        If GProcGetColumnValue("tblMastAccount", "AcName", txttemp(8), "S", "AcCode", "N") = 0 Then
           GName = GProcGetColumnValue("TblMastGroup", "AgCode", 90016, "N", "AgName", "S")
           GProcShowForm frmMastAccount, frmMain.mnuMstfrm(0), Array(frmMastAccount.txt(1), frmMastAccount.txttemp(2)), Array(txttemp(8), GName)
           Exit Function
        End If
    '---- Amount
    If CDbl(txt(16)) <= 0 Then
        MsgBox "Check Voucher Amount.", vbInformation + vbOKOnly, Me.Caption
        Exit Function
    End If
    '---Check Dublicate Bill No
    If Trim(txt(9)) <> "" Then
        X = "Select vno From " & tblName & " Where PurBillNo= '" & txt(9) & "' and PurAcCrCode = " & CLng(txt(5)) & "  and  vno <> " & CLng(txt(0)) & "  and Vtype='" & frmTypeStr & "'  and Vfirm = '" & gCCode & "'  and VYear=" & gCYear
        i = GProcRstOpen(rstAddCheck, X, "R")
        If i > 0 Then
            MsgBox "Same Bill No of Same Party is Available in Voucher No. " & rstAddCheck.Fields(0), vbCritical + vbOKOnly, "Duplicate Record"
            txt(9).SetFocus
            Exit Function
        End If
    End If
     
    '---Check Sub Groub Ac
    If Val(txt(67)) > 1 Then
        X = "select * from tblMastSubGroup where SubCode = " & txt(67) & ""
        i = GProcRstOpen(rstAddCheck, X, "R")
        If i > 0 And Val(txt(6)) <> rstAddCheck.Fields("SubGpCode") Then
            MsgBox "Please Check Sub Group A/c " & rstAddCheck.Fields(1), vbCritical + vbOKOnly, ""
            txttemp(67).SetFocus
            Exit Function
        End If
    
    End If
     
     '-- Expenses A/c
'    If gExpensesAcCode = GProcGetColumnValue("tblmastAccount", "AcName", "-", "S", "AcCode", "N") Then
'        ExpAcCode = 0
'    Else
'        ExpAcCode = gExpensesAcCode
'    End If
'    If txtExpensesAmt > 0 And ExpAcCode = 0 Then
'        MsgBox "Check Purchase Expenses Account.", vbInformation + vbOKOnly, Me.Caption
'        GProcShowForm frmSetting, frmMain.mnuMstSetting(0)
'        Exit Function
'    End If
    '----Duplicate Records
    '----Addition Case
    If FormAction = vbDataActionAddNew Then
        X = "Select vno From " & tblName & " Where vno=" & CLng(txt(0)) & " and Vtype='" & frmTypeStr & "'  and Vfirm = '" & gCCode & "'  and VYear=" & gCYear
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
Private Sub FillTxtFromTemp()
txt(1) = frmTypeStr '--VType
If IsDate(mskDt(10)) = False Then mskDt(10) = "__/__/____" '--BillDt
If IsDate(mskDt(32)) = False Then mskDt(32) = "__/__/____" '--Due Dt
If IsDate(mskDt(14)) = False Then mskDt(14) = "__/__/____" '--lr Dt

If IsDate(mskDt(2)) Then '--Vdt
    txt(2) = CDate(mskDt(2))
Else
    txt(2) = ""
End If
If IsDate(mskDt(10)) Then '--Party Bill Date
    txt(10) = CDate(mskDt(10))
Else
    txt(10) = ""
End If
If IsDate(mskDt(32)) Then '--Due Date
    txt(32) = CDate(mskDt(32))
Else
    txt(32) = ""
End If
If IsDate(mskDt(14)) Then '--lr Date
    txt(14) = CDate(mskDt(14))
Else
    txt(14) = ""
End If

If IsDate(mskDt(53)) Then '--Party Bill Date
    txt(53) = CDate(mskDt(53))
Else
    txt(53) = ""
End If


'-- Cr Code
txt(5) = GProcGetColumnValue("tblMastAccount", "AcName", txttemp(5), "S", "AcCode", "N")
'-- Dr Code
txt(6) = GProcGetColumnValue("tblMastAccount", "AcName", txttemp(6), "S", "AcCode", "N")
'--Mill
txt(7) = GProcGetColumnValue("tblMastAccount", "AcName", "-", "S", "AcCode", "N")

'-- broker
txt(8) = GProcGetColumnValue("tblMastAccount", "AcName", "-", "S", "AcCode", "N")
'--TDS A/c
If txttemp(65) = "" Then
   txttemp(65) = "-"
End If
txt(65) = GProcGetColumnValue("tblMastAccount", "AcName", txttemp(65), "S", "AcCode", "N")

If txttemp(67) = "" Then
   txttemp(67) = "-"
End If

'--Sub Group A/c
txt(67) = GProcGetColumnValue("tblMastSubGroup", "SubName", txttemp(67), "S", "SubCode", "N")
'-- Transport
If txttemp(11) = "" Then
   txttemp(11) = "-"
End If
txt(11) = GProcGetColumnValue("tblMastNarration", "Narration", txttemp(11), "S", "NarrCode", "N")
If txttemp(54) = "" Then
   txttemp(54) = "-"
End If
txt(54) = GProcGetColumnValue("tblMastAccount", "AcName", txttemp(54), "S", "AcCode", "N")
'-- TaxType
txt(22) = GProcGetColumnValue("tblMastNarration", "Narration", Cbo(22), "S", "NarrCode", "N")
txt(35) = GProcGetColumnValue("tblMastTransactionCode", "TranCode", txttemp(35), "S", "TranCode", "S")
txt(39) = GProcGetColumnValue("tblMastAccount", "AcName", "-", "S", "AcCode", "N")


txt(3) = gCYear
txt(4) = gCCode
txt(46) = chkIsCapitalGoods.Value
txt(51) = chkIsRCM.Value
ShowAccountBalance
End Sub
Private Sub FillTempFromTxt()
Dim X As String
Dim MsOms As Integer
Dim RcmTotal As Double
FillCombo
If IsDate(txt(2)) Then '--VDt
    mskDt(2) = txt(2)
Else
    mskDt(2) = "__/__/____"
End If
If IsDate(txt(10)) Then '--Party Bill Date
    mskDt(10) = txt(10)
Else
    mskDt(10) = "__/__/____"
End If
If IsDate(txt(32)) Then '--Due Date
    mskDt(32) = txt(32)
Else
    mskDt(32) = "__/__/____"
End If
If IsDate(txt(14)) Then '--LR Date
    mskDt(14) = txt(14)
Else
    mskDt(14) = "__/__/____"
End If
If IsDate(txt(53)) Then '--LR Date
    mskDt(53) = txt(53)
Else
    mskDt(53) = "__/__/____"
End If

'-- Cr Code
txttemp(5) = GProcGetColumnValue("tblMastAccount", "AcCode", txt(5), "N", "AcName", "S")
txtCode(5) = GProcGetColumnValue("tblMastAccount", "AcCode", txt(5), "N", "AcOurCode", "S")

lblTinNo.Caption = GProcGetColumnValue("tblMastAccount", "AcCode", txt(5), "N", "AcGSTIN", "S")
MsOms = GProcGetColumnValue("tblMastAccount", "AcCode", txt(5), "N", "AcMsOmsParty", "N")
If MsOms = 0 Then
   lblMsOms.Caption = "M.S."
Else
   lblMsOms.Caption = "O.M.S."
End If

'-- Dr Code
txttemp(6) = GProcGetColumnValue("tblMastAccount", "AcCode", txt(6), "N", "AcName", "S")
'--Mill
txttemp(7) = GProcGetColumnValue("tblMastAccount", "AcCode", txt(7), "N", "AcName", "S")

'--Broker
txttemp(8) = GProcGetColumnValue("tblMastAccount", "AcCode", txt(8), "N", "AcName", "S")
'--Transport
txttemp(11) = GProcGetColumnValue("tblMastNarration", "NarrCode", txt(11), "N", "Narration", "S")
'--Taxpayable Ac
'txttemp(39) = GProcGetColumnValue("tblMastAccount", "AcCode", "-", "S", "AcName", "S")
txttemp(54) = GProcGetColumnValue("tblMastAccount", "AcCode", txt(54), "N", "AcName", "S")
'--Tds A/c
txttemp(65) = GProcGetColumnValue("tblMastAccount", "AcCode", txt(65), "N", "AcName", "S")
txtCode(65) = GProcGetColumnValue("tblMastAccount", "AcCode", txt(65), "N", "AcOurCode", "S")
'--Sub Group A/c
txttemp(67) = GProcGetColumnValue("tblMastSubGroup", "SubCode", txt(67), "N", "SubName", "S")

'-- TaxType
Cbo(22) = GProcGetColumnValue("tblMastNarration", "NarrCode", txt(22), "N", "Narration", "S")
txttemp(35) = GProcGetColumnValue("tblMastTransactionCode", "TranCode", txt(35), "S", "TranCode", "S")
lblNature.Caption = GProcGetColumnValue("tblMastTransactionCode", "TranCode", txt(35), "S", "Nature", "S")
lblDescription.Caption = GProcGetColumnValue("tblMastTransactionCode", "TranCode", txt(35), "S", "Description", "S")

chkIsCapitalGoods.Value = Val(txt(46))
chkIsRCM.Value = Val(txt(51))

If chkIsRCM.Value = 1 Then
   lblRcmSubAmt.Caption = txt(16)
   RcmTotal = Val(txt(16)) + Val(txt(56)) + Val(txt(58)) + Val(txt(56))
   RcmTotal = GProcMakeRounding(RcmTotal)
   lblRCMTotal.Caption = RcmTotal
Else
   lblRcmSubAmt.Caption = 0
   lblRCMTotal.Caption = 0
End If
'-- Fill Related Record
FillRelatedRecord
ShowAccountBalance
MillCtrlEd
End Sub
Private Sub FillRelatedRecord()
Dim RowIndex As Long
Dim ColIndex As Long
'Dim TotBag As Long
'Dim TotWt As Double
'tmpAutoLRNo = "" ''-Auto LR No
'--Fill Grid
SetGrid
With rstblSub
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
'                If ColIndex = 9 Then TotBag = TotBag + CLng(msGrid.TextMatrix(RowIndex, 9)) '-- Tot Bag
'                If ColIndex = 10 Then TotWt = TotWt + CDbl(msGrid.TextMatrix(RowIndex, 10)) '-- Tot Wt
            Next ColIndex
            RowIndex = RowIndex + 1
            .MoveNext
            If Not .EOF Then msGrid.rows = msGrid.rows + 1
        Loop
    End If
End With
'--- Fill Related Control
SetRelatedRS
End Sub
Private Sub ReadFields()
FormAction = vbDataActionClose
GProcShowRecord Me, rstbl, MaxNo, OtherTxtFromNo, OtherTxtToNo
FillTempFromTxt
End Sub
Private Sub CalAmount()
Dim TaxAmt As Double
Dim TaxAmtCGST As Double
Dim TaxAmtIGST As Double
Dim SubAmt As Double
Dim Amt As Double
Dim ROff As Double
Dim ExmptAmt As Double
Dim ExmptKg As Double

Dim AssValue As Double
Dim RowIndex As Integer
Dim TotBag As Long
Dim TotWt As Double
Dim LCost As Double
Dim MillAcCode As Long
Dim IsExemptMill As Long
Dim TaxCd As Long
Dim RcmSgst As Double
Dim RcmCgst As Double
Dim RcmIgst As Double
Dim RcmTotal As Double
Dim TdsAmt As Double
Dim TcsAmt As Double
'-- Is Exempt Mill
IsExemptMill = GProcGetColumnValue("TblMastAccount", "AcName", txttemp(5), "S", "ACIsExemptMill", "N")
'-- Due Date
If IsNumeric(txt(31)) = False Then txt(31) = "0"
txt(31) = Format(txt(31), PStr)
If IsDate(mskDt(2)) Then mskDt(32) = DateAdd("d", txt(31), mskDt(2))
'---- Sub amt
CheckForNumberGrid
CheckForNumber
RowIndex = 1
With msGrid
    Do While RowIndex < .rows
        '-- sub amt
        If Val(.TextMatrix(RowIndex, 6)) > 0 Then
            SubAmt = SubAmt + Val(.TextMatrix(RowIndex, 6))
        End If
        '--Tot Bag
        If Val(.TextMatrix(RowIndex, 2)) > 0 Then
            TotBag = TotBag + Val(.TextMatrix(RowIndex, 2))
        End If
        '--Tot Wt
        If Val(.TextMatrix(RowIndex, 3)) > 0 Then
            TotWt = TotWt + Val(.TextMatrix(RowIndex, 3))
        End If
        If RowIndex = 1 And msGrid.TextMatrix(1, 1) <> "" Then
           If (Left(gCGStin, 2) = Left(lblTinNo.Caption, 2) Or Left(lblTinNo.Caption, 2) = "" Or Left(lblTinNo.Caption, 1) = "U") Then
               txt(23) = GProcGetColumnValue("tblMastItem", "ItName", msGrid.TextMatrix(1, 1), "S", "ITMAXRATE", "N", "IsOtherItem= 1") '--Item Code
               txt(26) = GProcGetColumnValue("tblMastItem", "ItName", msGrid.TextMatrix(1, 1), "S", "ITMINRATE", "N", "IsOtherItem= 1") '--Item Code
               txt(37) = 0
           Else
               txt(23) = 0
               txt(26) = 0
               txt(37) = GProcGetColumnValue("tblMastItem", "ItName", msGrid.TextMatrix(1, 1), "S", "ItBrkRtOld", "N", "IsOtherItem= 1") '--Item Code
           End If
        End If
        RowIndex = RowIndex + 1
    Loop
End With
'--- Subamt+Add -Less
AssValue = 0
TaxAmt = 0
TaxAmtCGST = 0
TaxAmtIGST = 0
If chkIsRCM.Value = 0 Then
    AssValue = SubAmt + CDbl(txt(17)) - CDbl(txt(18))
    '- 1 SGST
    If Abs((CDbl(txt(24)) - (AssValue * CDbl(txt(23)) / 100))) > 1 Then
       TaxAmt = AssValue * CDbl(txt(23)) / 100
    Else
       TaxAmt = CDbl(txt(24))
    End If
    TaxAmt = Format(TaxAmt, FStr)
       
    '- 2 CGST
    If Abs((CDbl(txt(27)) - (AssValue * CDbl(txt(26)) / 100))) > 1 Then
       TaxAmtCGST = AssValue * CDbl(txt(26)) / 100
    Else
       TaxAmtCGST = CDbl(txt(27))
    End If
    TaxAmtCGST = Format(TaxAmtCGST, FStr)
       
    '- 3 IGST
    If Abs((CDbl(txt(38)) - (AssValue * CDbl(txt(37)) / 100))) > 1 Then
       TaxAmtIGST = AssValue * CDbl(txt(37)) / 100
    Else
       TaxAmtIGST = CDbl(txt(38))
    End If
    TaxAmtIGST = Format(TaxAmtIGST, FStr)
End If
'---TCS
    If Format(mskDt(2), "YYYY/MM/DD") >= Format("2020/01/10", "YYYY/MM/dd") Then
       txt(69) = "0.075"
       If Val(txt(68)) > 0 And Val(txt(70)) = 0 Then
          TcsAmt = (Val(txt(68)) * Val(txt(69))) / 100
          TcsAmt = Format(GProcMakeRounding(TcsAmt), FStr)
       Else
          TcsAmt = Format(txt(70), FStr)
       End If
       txt(70) = TcsAmt
    End If
    If Val(txt(70)) = 0 Then txt(69) = "0.00"
'--Roff

ROff = txt(30) '= 'CDbl(GProcMakeRounding(Amt) - Amt)
'-- SubAmt+Tax Amt +Add -Less+roff + Aft Tax Add - Aft Tax Less
Amt = SubAmt + TaxAmt + TaxAmtCGST + TaxAmtIGST + CDbl(txt(17)) - CDbl(txt(18)) + ROff + CDbl(txt(28)) - CDbl(txt(29)) + Val(txt(70))

If gCYear >= 2023 Then
  txt(72) = txt(15)
  TdsAmt = (Val(txt(72)) * Val(txt(63))) / 100
  txt(64) = GProcMakeRounding(TdsAmt)
Else
  TdsAmt = (Val(txt(15)) * Val(txt(63))) / 100
  txt(64) = GProcMakeRounding(TdsAmt)
End If

txt(15) = SubAmt
txt(24) = TaxAmt
txt(27) = TaxAmtCGST
txt(38) = TaxAmtIGST
txt(21) = AssValue
txt(16) = Amt

'-- For RCM Bill
If chkIsRCM.Value = 1 Then
   txt(55) = txt(23)
   txt(57) = txt(26)
   txt(59) = txt(37)
   DoEvents
   lblRcmSubAmt.Caption = txt(16)
   RcmSgst = Round(SubAmt * CDbl(txt(55)) / 100, 2)
   RcmSgst = Format(RcmSgst, FStr)
   RcmCgst = Round(SubAmt * CDbl(txt(57)) / 100, 2)
   RcmCgst = Format(RcmCgst, FStr)
   RcmIgst = Round(SubAmt * CDbl(txt(59)) / 100, 2)
   RcmIgst = Format(RcmIgst, FStr)
   txt(56) = RcmSgst
   txt(58) = RcmCgst
   txt(60) = RcmIgst
   RcmTotal = SubAmt + RcmSgst + RcmCgst + RcmIgst
   RcmTotal = GProcMakeRounding(RcmTotal)
   lblRCMTotal.Caption = RcmTotal
   txt(23) = 0
   txt(24) = 0
   txt(26) = 0
   txt(27) = 0
   txt(37) = 0
   txt(38) = 0
   mskDt(53) = mskDt(2)
Else
   lblRcmSubAmt.Caption = 0
   lblRCMTotal.Caption = 0
   txt(55) = 0
   txt(56) = 0
   txt(57) = 0
   txt(58) = 0
   txt(59) = 0
   txt(60) = 0
   txt(52) = ""
End If
'--
'If Abs(SubAmt + CDbl(txt(17)) - CDbl(txt(18)) - AssValue) < 2 Then
'    ExmptAmt = 0
'    ExmptKg = 0
'End If
txt(19) = 0 '--ExmptAmt
txt(20) = 0 '--ExmptKg
CheckForNumber
TaxCd = GProcGetColumnValue("tblMastNarration", "Narration", Cbo(22).text, "S", "MastTaxPurAcCode", "N")
txttemp(39) = GProcGetColumnValue("tblMastAccount", "AcCode", CStr(TaxCd), "N", "AcName", "S")

End Sub
'--- Is Related Record
Private Function IsRelatedRecord() As Boolean
Dim RsG As Recordset
Dim X As String
'-- Audited
'If GProcISAudited(txt(0), txt(1), txt(23)) = True Then
'    MsgBox "Record is Audited.You can not Modify/Delete this Record.", vbCritical + vbOKOnly, App.Title
'    IsRelatedRecord = True
'    Exit Function
'End If
'-- Is related payment
X = "Select Vno,Vdt from tblVoucher where VBillType= '" & frmTypeStr & "' and " _
& "  VBillVno = " & txt(0) & " and VFirm  = '" & gCCode & "' and  VBillVYear= " & gCYear & ""
GProcRstOpen RsG, X, "R"
If RsG.EOF Then
Else
    IsRelatedRecord = True
    MsgBox "You can not Modify/Delete this record.Related No. " & RsG.Fields(0) & " is exists.", vbCritical + vbOKOnly, App.Title
    Exit Function
End If
RsG.Close
End Function
'-- Is Related GatePass
Private Function IsRelatedGatePass() As Boolean
Dim RsG As Recordset
Dim X As String
'-- Is Gate Pass
'X = "Select Vno,Vdt from tblVouSub where Vtype='GP' and " _
'& " VSubPItCtrlNo in (select VSubItCtrlNo from tblVouSub where vtype='" & frmTypeStr & "' and vno=" & txt(0) & " and " _
'& " VYear=" & gCYear & ")"
'GProcRstOpen RsG, X, "R"
'If RsG.EOF Then
'Else
'    IsRelatedGatePass = True
'    MsgBox "You can not Modify/Delete this record.Related Gate Pass No. " & RsG.Fields(0) & " is exists.", vbCritical + vbOKOnly, App.Title
'    Exit Function
'End If
'RsG.Close
End Function
'-- Is Related It Ctrl GatePass
Private Function IsRelatedItCtrlGatePass(ItCtrlNo As Long) As Boolean
Dim RsG As Recordset
Dim X As String
'-- Is Gate Pass
'X = "Select Vno,Vdt from tblVouSub where Vtype='GP' and " _
'& " VSubPItCtrlNo = " & ItCtrlNo
'GProcRstOpen RsG, X, "R"
'If RsG.EOF Then
'Else
'    IsRelatedItCtrlGatePass = False
'    Exit Function
'End If
'RsG.Close
End Function
'---------- Booking Vs Purchase
Private Sub PrepareBooking(Optional Chr As String)
Dim MillCode As Long
Dim QStr As String
Dim RsBk As Recordset
Dim Chr2 As String
If IsDate(mskDt(2)) = True Then
    MillCode = GProcGetColumnValue("TblMastAccount", "AcName", txttemp(5), "S", "Accode", "N")
    QStr = " select tblAddless.Adgpno, tblVousub.VNo,format(tblVousub.vdt,'dd/mm/yyyy'),ItName,vsubbag-(vsubissbag),vsubwt-(vsubisswt),vsubbag,vsubwt,vsubitctrlno,VsubRt from " _
    & " tblVousub,tblmastItem,tblAddless Where " _
    & " tblVousub.vno = tblAddless.vno and  tblVousub.vtype  = tblAddless.vtype and  tblVousub.vyear = tblAddless.vyear and " _
    & " Itcode=VSUBITCODE and ItMillCode=" & MillCode & " and " _
    & " ((vsubbag-(vsubissbag))> 0) and tblVousub.vtype ='OT'  and VSUBBOOKISCOMPLETED=0 and "
    If gBackEndDB = gBackEndAccess Then
        QStr = QStr + " tblVousub.vdt<=cdate('" & CDate(mskDt(2)) & "')"
    Else '-- Oracle
        QStr = QStr + " tblVousub.vdt <= to_date('" & CDate(mskDt(2)) & "','dd/MM/yyyy')"
    End If
    
    
    
    With gClsSearch
        .SearchMultiField QStr, "", Array("Contract No.", " Booking No.", "Date", "Count", "Bal Bag", "Bal Wt.", "Bag", "Wt", "itctrlno", "Rate"), Array(1200, 1200, 1500, 1500, 1000, 1000, 500, 1000, 0, 1000), "", Chr, "tblVousub.Vdt,tblAddless.Adgpno,tblVousub.vno,VSUBCTRNO", 0, 0, False, , True
        If .SearchMultiRetCol(1) <> "" Then
            QStr = "select tblvousub.*,tblmastItem.ItName from tblvousub,tblMastItem where " _
            & " ItCode=VSUBITCODE and VSUBITCTRLNO= " & CLng(.SearchMultiRetCol(8))
              If GProcRstOpen(RsBk, QStr, "R") > 0 Then
                    txtGrid = ""
                    txtGrid = CLng(.SearchMultiRetCol(1)) '--Book Vno
                    msGrid.TextMatrix(msGrid.Row, 1) = CLng(.SearchMultiRetCol(1)) '--Book Vno
                    msGrid.TextMatrix(msGrid.Row, 2) = Format(.SearchMultiRetCol(2), "dd/mm/yyyy") '-- Book Vdt
                    msGrid.TextMatrix(msGrid.Row, 3) = RsBk!ItName  '-- Count
                    msGrid.TextMatrix(msGrid.Row, 9) = .SearchMultiRetCol(4)  '-- Bal Bag
                    msGrid.TextMatrix(msGrid.Row, 10) = .SearchMultiRetCol(5) '-- Bal Wt
                    msGrid.TextMatrix(msGrid.Row, 19) = RsBk!VSubItCtrlNo  '--  P ItCtrl No =Book ItCtrl No
                    msGrid.TextMatrix(msGrid.Row, 11) = RsBk!vsubRt  '-- Rate
                    msGrid.TextMatrix(msGrid.Row, 12) = RsBk!vsubRTPER  '--Rate Per
                    CalGridAmt
                    CalAmount
              End If
        End If
    End With
    
'    With gClsSearch
'        .SearchMultiField QStr, "", Array("Booking No.", "Date", "Count", "Bal Bag", "Bal Wt.", "Rate", "Wt", "itctrlno"), Array(1000, 1500, 1500, 1000, 1000, 1000, 0, 0), "", Chr, "tblVousub.Vdt,tblVousub.vno,VSUBCTRNO ", 0, 0, False, , True
'        If .SearchMultiRetCol(0) <> "" Then
'            QStr = "select tblvousub.*,tblmastItem.ItName from tblvousub,tblMastItem where " _
'            & " ItCode=VSUBITCODE and VSUBITCTRLNO= " & CLng(.SearchMultiRetCol(7))
'              If GProcRstOpen(RsBk, QStr, "R") > 0 Then
'                    txtGrid = ""
'                    txtGrid = CLng(.SearchMultiRetCol(0)) '--Book Vno
'                    msGrid.TextMatrix(msGrid.Row, 1) = CLng(.SearchMultiRetCol(0)) '--Book Vno
'                    msGrid.TextMatrix(msGrid.Row, 2) = Format(.SearchMultiRetCol(1), "dd/mm/yyyy") '-- Book Vdt
'                    msGrid.TextMatrix(msGrid.Row, 3) = RsBk!ItName  '-- Count
'              '      msGrid.TextMatrix(msGrid.Row, 9) = .SearchMultiRetCol(3)  '-- Bal Bag
'               '     msGrid.TextMatrix(msGrid.Row, 10) = .SearchMultiRetCol(4) '-- Bal Wt
'                    msGrid.TextMatrix(msGrid.Row, 4) = RsBk!vsubRt  '-- Rate
'
'                    msGrid.TextMatrix(msGrid.Row, 19) = RsBk!VSubItCtrlNo  '--  P ItCtrl No =Book ItCtrl No
''                    msGrid.TextMatrix(msGrid.Row, 11) = RsBk!vsubRt  '-- Rate
'                    msGrid.TextMatrix(msGrid.Row, 12) = RsBk!vsubRTPER  '--Rate Per
'                    CalGridAmt
'                    CalAmount
'              End If
'        End If
'    End With
End If
End Sub

Private Sub GenRCMBillNo()
If chkIsRCM.Value = 1 And txt(52) = "" Then
    Dim Y As String
    Dim G_Rs As Recordset
    Dim GInvNo As Long
    
    Y = "Select Max(cast(SUBSTRING(PurRCMBillNo," & 5 & ",13) as int)) from tblpurch where VYear=" & gCYear & " and VFirm  = '" & gCCode & "' and Vtype='PO' and PurRCMBillNo <> ''"
    GProcRstOpen G_Rs, Y, "R", gCn
    If G_Rs.Fields(0) = Null Or IsNumeric(G_Rs.Fields(0)) = False Then
        GInvNo = 1
    Else
        GInvNo = CLng(G_Rs.Fields(0)) + 1
    End If
    txt(52) = "RCM" & "-" & GInvNo
    If IsDate(mskDt(53)) = False Then
       mskDt(53) = mskDt(2)
    End If
    G_Rs.Close
End If
End Sub

