VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form frmRptTaxReports 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Tax Reports"
   ClientHeight    =   7260
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9075
   Icon            =   "frmRptTaxReports.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7260
   ScaleWidth      =   9075
   Begin VB.CommandButton cmdCrystal 
      Caption         =   "Excle"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   480
      Index           =   3
      Left            =   540
      Picture         =   "frmRptTaxReports.frx":058A
      TabIndex        =   66
      Top             =   6420
      Visible         =   0   'False
      Width           =   1185
   End
   Begin VB.CommandButton cmdCrystal 
      Caption         =   "Mail"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   480
      Index           =   2
      Left            =   1980
      Picture         =   "frmRptTaxReports.frx":09CC
      TabIndex        =   65
      Top             =   6420
      Visible         =   0   'False
      Width           =   1185
   End
   Begin VB.OptionButton optBrokList 
      BackColor       =   &H00EDFDFE&
      Caption         =   "All Brokers"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   270
      Index           =   2
      Left            =   5280
      TabIndex        =   64
      Top             =   6480
      Visible         =   0   'False
      Width           =   3375
   End
   Begin VB.OptionButton optBrokList 
      BackColor       =   &H00EDFDFE&
      Caption         =   "No Brokerage A/c"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   270
      Index           =   1
      Left            =   5280
      TabIndex        =   63
      Top             =   6060
      Visible         =   0   'False
      Width           =   3375
   End
   Begin VB.OptionButton optBrokList 
      BackColor       =   &H00EDFDFE&
      Caption         =   "Excluding No Brokerage A/c"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   195
      Index           =   0
      Left            =   5280
      TabIndex        =   62
      Top             =   5700
      Visible         =   0   'False
      Width           =   3375
   End
   Begin VB.CheckBox chkGenVoucher 
      BackColor       =   &H80000018&
      Caption         =   "Generate Brokerage Vou."
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
      Height          =   330
      Left            =   1680
      TabIndex        =   61
      Top             =   4440
      Visible         =   0   'False
      Width           =   7170
   End
   Begin VB.CheckBox chkIsExcludingSale 
      BackColor       =   &H80000018&
      Caption         =   "Excluding Paid Brokerage"
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
      Height          =   330
      Left            =   1680
      TabIndex        =   58
      Top             =   4080
      Visible         =   0   'False
      Width           =   3150
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
      Index           =   6
      Left            =   1680
      MaxLength       =   250
      TabIndex        =   12
      Text            =   "6"
      Top             =   3600
      Visible         =   0   'False
      Width           =   6810
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
      Index           =   0
      Left            =   1680
      TabIndex        =   3
      ToolTipText     =   "Press F3 for Creation"
      Top             =   780
      Width           =   1215
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
      Index           =   2
      Left            =   1680
      TabIndex        =   7
      ToolTipText     =   "Press F3 for Creation"
      Top             =   1680
      Width           =   1215
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
      Index           =   3
      Left            =   1680
      TabIndex        =   5
      ToolTipText     =   "Press F3 for Creation"
      Top             =   1260
      Width           =   1215
   End
   Begin VB.TextBox txt 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   5
      Left            =   10920
      TabIndex        =   37
      Text            =   "5"
      Top             =   3000
      Width           =   2100
   End
   Begin VB.TextBox txt 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   4
      Left            =   12855
      TabIndex        =   52
      Text            =   "4"
      Top             =   5475
      Width           =   795
   End
   Begin VB.TextBox txt 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   3
      Left            =   10665
      TabIndex        =   50
      Text            =   "3"
      Top             =   5475
      Width           =   1140
   End
   Begin VB.TextBox txt 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   2
      Left            =   12855
      TabIndex        =   48
      Text            =   "2"
      Top             =   4980
      Width           =   795
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
      Index           =   1
      Left            =   1680
      TabIndex        =   11
      Text            =   "1"
      Top             =   3060
      Visible         =   0   'False
      Width           =   1830
   End
   Begin VB.Frame Frame4 
      BackColor       =   &H80000018&
      Caption         =   "Sale Type"
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
      Height          =   1020
      Left            =   9585
      TabIndex        =   39
      Top             =   3885
      Width           =   3450
      Begin VB.OptionButton OptBrkSAType 
         BackColor       =   &H80000018&
         Caption         =   "Mill Bill"
         ForeColor       =   &H00000080&
         Height          =   360
         Index           =   5
         Left            =   2205
         TabIndex        =   45
         Top             =   540
         Width           =   990
      End
      Begin VB.OptionButton OptBrkSAType 
         BackColor       =   &H80000018&
         Caption         =   "SIT"
         ForeColor       =   &H00000080&
         Height          =   360
         Index           =   2
         Left            =   2205
         TabIndex        =   42
         Top             =   225
         Width           =   1050
      End
      Begin VB.OptionButton OptBrkSAType 
         BackColor       =   &H80000018&
         Caption         =   "Depot"
         ForeColor       =   &H00000080&
         Height          =   360
         Index           =   4
         Left            =   1125
         TabIndex        =   44
         Top             =   540
         Width           =   1335
      End
      Begin VB.OptionButton OptBrkSAType 
         BackColor       =   &H80000018&
         Caption         =   "Consign"
         ForeColor       =   &H00000080&
         Height          =   360
         Index           =   3
         Left            =   120
         TabIndex        =   43
         Top             =   540
         Width           =   1335
      End
      Begin VB.OptionButton OptBrkSAType 
         BackColor       =   &H80000018&
         Caption         =   "Trade"
         ForeColor       =   &H00000080&
         Height          =   360
         Index           =   1
         Left            =   1125
         TabIndex        =   41
         Top             =   225
         Width           =   1335
      End
      Begin VB.OptionButton OptBrkSAType 
         BackColor       =   &H80000018&
         Caption         =   "All"
         ForeColor       =   &H00000080&
         Height          =   360
         Index           =   0
         Left            =   120
         TabIndex        =   40
         Top             =   225
         Value           =   -1  'True
         Width           =   1335
      End
   End
   Begin VB.CheckBox ChkPartyOnNewPage 
      BackColor       =   &H80000018&
      Caption         =   "Is Partywise Print  ?"
      ForeColor       =   &H00000080&
      Height          =   330
      Left            =   6600
      TabIndex        =   13
      Top             =   2880
      Visible         =   0   'False
      Width           =   2190
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H80000018&
      Caption         =   "&Sales / Purchase Report On"
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
      Height          =   810
      Left            =   13140
      TabIndex        =   23
      Top             =   1125
      Width           =   3480
      Begin VB.OptionButton OptBag 
         BackColor       =   &H80000018&
         Caption         =   "Bags"
         ForeColor       =   &H00000080&
         Height          =   360
         Left            =   2520
         TabIndex        =   26
         Top             =   300
         Width           =   930
      End
      Begin VB.OptionButton OptGrossAmt 
         BackColor       =   &H80000018&
         Caption         =   "&Gross Amt"
         ForeColor       =   &H00000080&
         Height          =   360
         Left            =   165
         TabIndex        =   24
         Top             =   300
         Value           =   -1  'True
         Width           =   1080
      End
      Begin VB.OptionButton OptNetAmt 
         BackColor       =   &H80000018&
         Caption         =   "&Net Amt"
         ForeColor       =   &H00000080&
         Height          =   360
         Left            =   1425
         TabIndex        =   25
         Top             =   300
         Width           =   1020
      End
   End
   Begin VB.TextBox txtItCode 
      Height          =   375
      Left            =   11115
      TabIndex        =   54
      Top             =   6120
      Visible         =   0   'False
      Width           =   570
   End
   Begin VB.TextBox txtItem 
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
      Left            =   1680
      TabIndex        =   9
      Top             =   2175
      Visible         =   0   'False
      Width           =   6885
   End
   Begin VB.CheckBox ChkSummary 
      BackColor       =   &H80000018&
      Caption         =   "Is Summary ?"
      ForeColor       =   &H00000080&
      Height          =   330
      Left            =   9585
      TabIndex        =   35
      Top             =   2550
      Width           =   1830
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H80000018&
      Height          =   1725
      Left            =   11340
      TabIndex        =   31
      Top             =   645
      Width           =   1710
      Begin VB.OptionButton Opt 
         BackColor       =   &H80000018&
         Caption         =   "Pe&nding"
         ForeColor       =   &H00000080&
         Height          =   360
         Index           =   1
         Left            =   225
         TabIndex        =   33
         Top             =   675
         Value           =   -1  'True
         Width           =   1005
      End
      Begin VB.OptionButton Opt 
         BackColor       =   &H80000018&
         Caption         =   "&Received"
         ForeColor       =   &H00000080&
         Height          =   360
         Index           =   2
         Left            =   225
         TabIndex        =   34
         Top             =   1200
         Width           =   1215
      End
      Begin VB.OptionButton Opt 
         BackColor       =   &H80000018&
         Caption         =   "&All"
         ForeColor       =   &H00000080&
         Height          =   360
         Index           =   0
         Left            =   225
         TabIndex        =   32
         Top             =   210
         Width           =   1215
      End
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
      Height          =   375
      Index           =   0
      Left            =   1680
      TabIndex        =   10
      Text            =   "0"
      Top             =   2625
      Visible         =   0   'False
      Width           =   1830
   End
   Begin VB.TextBox txtParty 
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
      Left            =   3075
      TabIndex        =   8
      Top             =   1725
      Visible         =   0   'False
      Width           =   5490
   End
   Begin VB.CheckBox ChkConsolidate 
      BackColor       =   &H00EDFDFE&
      Caption         =   "Is All Firm Report ?"
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
      Height          =   330
      Left            =   6600
      TabIndex        =   22
      Top             =   2520
      Visible         =   0   'False
      Width           =   1995
   End
   Begin VB.CommandButton cmdCrystal 
      Height          =   480
      Index           =   1
      Left            =   1995
      Picture         =   "frmRptTaxReports.frx":0E0E
      Style           =   1  'Graphical
      TabIndex        =   15
      ToolTipText     =   "Print"
      Top             =   5835
      Width           =   1185
   End
   Begin VB.CommandButton cmdCrystal 
      Height          =   480
      Index           =   0
      Left            =   540
      Picture         =   "frmRptTaxReports.frx":1478
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   5835
      Width           =   1185
   End
   Begin VB.CommandButton cmdExit 
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
      Left            =   3465
      TabIndex        =   16
      Top             =   5835
      Width           =   1185
   End
   Begin VB.TextBox txtBroker 
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
      Left            =   3075
      TabIndex        =   6
      Top             =   1260
      Visible         =   0   'False
      Width           =   5490
   End
   Begin VB.TextBox txtMill 
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
      Left            =   3075
      TabIndex        =   4
      Top             =   780
      Visible         =   0   'False
      Width           =   5490
   End
   Begin Crystal.CrystalReport CrystalReport1 
      Left            =   8700
      Top             =   4740
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
   Begin MSComctlLib.ProgressBar prgBar 
      Height          =   255
      Left            =   255
      TabIndex        =   53
      Top             =   5370
      Visible         =   0   'False
      Width           =   8250
      _ExtentX        =   14552
      _ExtentY        =   450
      _Version        =   393216
      Appearance      =   1
   End
   Begin MSComCtl2.DTPicker dtpToDt 
      Height          =   360
      Left            =   6780
      TabIndex        =   2
      Top             =   300
      Width           =   1785
      _ExtentX        =   3149
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
      Format          =   124649475
      CurrentDate     =   38050
   End
   Begin MSComCtl2.DTPicker dtpFrDt 
      Height          =   360
      Left            =   1680
      TabIndex        =   1
      Top             =   300
      Width           =   1785
      _ExtentX        =   3149
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
      Format          =   124649475
      CurrentDate     =   38050
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H80000018&
      Height          =   1740
      Left            =   9585
      TabIndex        =   28
      Top             =   630
      Width           =   1695
      Begin VB.OptionButton optPartyGroup 
         BackColor       =   &H00C0FFFF&
         Caption         =   "Party Groupwise"
         ForeColor       =   &H00000080&
         Height          =   195
         Left            =   120
         MaskColor       =   &H00C0FFFF&
         TabIndex        =   56
         Top             =   1380
         Visible         =   0   'False
         Width           =   1455
      End
      Begin VB.OptionButton optOnlyParty 
         BackColor       =   &H00C0FFFF&
         Caption         =   "Partywise"
         ForeColor       =   &H00000080&
         Height          =   195
         Left            =   150
         MaskColor       =   &H00C0FFFF&
         TabIndex        =   55
         Top             =   960
         Visible         =   0   'False
         Width           =   1035
      End
      Begin VB.OptionButton OptParty 
         BackColor       =   &H80000018&
         Caption         =   "Partywise"
         ForeColor       =   &H00004080&
         Height          =   360
         Left            =   150
         MaskColor       =   &H00C0FFFF&
         TabIndex        =   30
         Top             =   495
         Value           =   -1  'True
         Width           =   1335
      End
      Begin VB.OptionButton OptBrk 
         BackColor       =   &H80000018&
         Caption         =   "Brokerwise"
         ForeColor       =   &H00004080&
         Height          =   360
         Left            =   150
         TabIndex        =   29
         Top             =   180
         Width           =   1335
      End
   End
   Begin MSComCtl2.DTPicker dtpVouDt 
      Height          =   360
      Left            =   3420
      TabIndex        =   59
      Top             =   4860
      Visible         =   0   'False
      Width           =   1785
      _ExtentX        =   3149
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
      Format          =   124649475
      CurrentDate     =   38050
   End
   Begin VB.Label Label11 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Voucher Date  :"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   330
      Left            =   1740
      TabIndex        =   60
      Top             =   4920
      Visible         =   0   'False
      Width           =   1620
   End
   Begin VB.Label Label10 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "&Broker Codes :"
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
      Left            =   300
      TabIndex        =   57
      Top             =   3600
      Visible         =   0   'False
      Width           =   1305
   End
   Begin VB.Shape Shape2 
      FillColor       =   &H00BDDFD1&
      FillStyle       =   0  'Solid
      Height          =   1305
      Left            =   285
      Shape           =   4  'Rounded Rectangle
      Top             =   5685
      Width           =   4530
   End
   Begin VB.Line Line1 
      BorderColor     =   &H000040C0&
      X1              =   9435
      X2              =   13800
      Y1              =   3450
      Y2              =   3450
   End
   Begin VB.Label Label9 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Form 'F' No"
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
      Left            =   9600
      TabIndex        =   36
      Top             =   3000
      Width           =   990
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "ECess %"
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
      Left            =   11955
      TabIndex        =   51
      Top             =   5475
      Width           =   825
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Surcharge %"
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
      Left            =   9510
      TabIndex        =   49
      Top             =   5475
      Width           =   1125
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "TDS %        :"
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
      Left            =   315
      TabIndex        =   47
      Top             =   3150
      Visible         =   0   'False
      Width           =   1155
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "TDS Limit     :"
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
      Left            =   270
      TabIndex        =   46
      Top             =   2700
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Bro&kerage Report   :"
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
      Height          =   330
      Left            =   9585
      TabIndex        =   38
      Top             =   3510
      Width           =   3375
   End
   Begin VB.Line Line3 
      BorderColor     =   &H000040C0&
      X1              =   9180
      X2              =   9180
      Y1              =   150
      Y2              =   5940
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "&Form Issue Receive Report   :"
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
      Height          =   330
      Left            =   9585
      TabIndex        =   27
      Top             =   285
      Width           =   2925
   End
   Begin VB.Label lblItem 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "&Item Name    :"
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
      Left            =   270
      TabIndex        =   21
      Top             =   2175
      Visible         =   0   'False
      Width           =   1245
   End
   Begin VB.Label lblParty 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "&Party Name   :"
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
      Left            =   270
      TabIndex        =   20
      Top             =   1725
      Visible         =   0   'False
      Width           =   1275
   End
   Begin VB.Label lblBroker 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "&Broker Name :"
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
      Left            =   270
      TabIndex        =   19
      Top             =   1260
      Visible         =   0   'False
      Width           =   1260
   End
   Begin VB.Label lblMill 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "&Mill Name      :"
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
      Left            =   270
      TabIndex        =   18
      Top             =   780
      Visible         =   0   'False
      Width           =   1290
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "To   :"
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
      Left            =   6120
      TabIndex        =   17
      Top             =   360
      Width           =   480
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "From &Date    :"
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
      Left            =   270
      TabIndex        =   0
      Top             =   315
      Width           =   1290
   End
   Begin VB.Shape Shape5 
      BorderColor     =   &H00000080&
      FillColor       =   &H00EDFDFE&
      FillStyle       =   0  'Solid
      Height          =   6990
      Left            =   105
      Top             =   120
      Width           =   8790
   End
End
Attribute VB_Name = "frmRptTaxReports"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim wMillCode As Long, wPartyCode As Long, wBrkCode As Long, wItemCode As Long, wItList As String
Dim CondStr As String
Dim formName1 As String
Dim formName2 As String
Dim formName3 As String

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
 '-- Activate Controls to False
GProcActivateControls Me, False
dtpFrDt.Enabled = True
dtpToDt.Enabled = True
dtpFrDt.Value = Format(gCYSDate, "dd/MM/yyyy")
If Format(Now, "dd/MM/yyyy") > Format(gCYEDate, "dd/MM/yyyy") Then
   dtpToDt.Value = Format(gCYEDate, "dd/MM/yyyy")
Else
   dtpToDt.Value = Format(Now, "dd/MM/yyyy")
End If
dtpVouDt.Value = Format(gCYEDate, "dd/MM/yyyy")
Me.Caption = gSelectedMenu
For i = 0 To 4
    txt(i) = "0.00"
Next i
txt(5) = "" '-F Form No
txt(6) = ""
Select Case gSelectedMenu
    '========== General Report
    Case "Total Sales Summary" 'Total Sales
         txtMill.Visible = True
         txtBroker.Visible = True
         txtParty.Visible = True
         txtMill.Enabled = True
         txtBroker.Enabled = True
         txtParty.Enabled = True
         lblMill.Visible = True
         lblBroker.Visible = True
         lblParty.Visible = True
         ChkConsolidate.Visible = True
         ChkConsolidate.Enabled = True
'         ChkConsolidate.Value = 1
    Case "Count / Brokerwise Sales", "Party / Countwise Sales", "Sales Against Booking"
         txtMill.Visible = True
         txtBroker.Visible = True
         txtParty.Visible = True
         txtMill.Enabled = True
         txtBroker.Enabled = True
         txtParty.Enabled = True
         lblMill.Visible = True
         lblBroker.Visible = True
         lblParty.Visible = True
         ChkConsolidate.Visible = True
         ChkConsolidate.Enabled = True
'         ChkConsolidate.Value = 1
         txtItem.Visible = True
         txtItem.Enabled = True
         lblItem.Visible = True
         txtCode(0).Enabled = True
         txtCode(2).Enabled = True
         txtCode(3).Enabled = True
         If gSelectedMenu = "Party / Countwise Sales" Then
            ChkPartyOnNewPage.Visible = True
            ChkPartyOnNewPage.Enabled = True
            ChkPartyOnNewPage.Caption = "Summery"
         End If
    Case "Broker Ledger", "Brokerage Detail ", "Brokerage Summery", "Brokerage Voucher", "Countwise Brokarage"
         If gSelectedMenu = "Brokerage Summery" Or gSelectedMenu = "Brokerage Voucher" Then
            txt(0).Visible = True
            txt(1).Visible = True
            txt(0).Enabled = True
            txt(1).Enabled = True
            If gCYear >= 2025 Then
               txt(0).text = 20000
            Else
               txt(0).text = 15000
            End If
            txt(1).text = 2
            Label5.Visible = True
            Label6.Visible = True
            chkIsExcludingSale.Visible = True
            chkIsExcludingSale.Enabled = True
         End If
         If gSelectedMenu = "Brokerage Detail " Then
            chkIsExcludingSale.Visible = True
            chkIsExcludingSale.Enabled = True
         End If
         If gSelectedMenu = "Brokerage Voucher" Then
            ChkPartyOnNewPage.Visible = True
            ChkPartyOnNewPage.Enabled = True
            ChkPartyOnNewPage.Caption = "Grand Total Required ?"
            Label10.Visible = True
         End If
         If gSelectedMenu = "Broker Ledger" Then
            ChkPartyOnNewPage.Visible = True
            ChkPartyOnNewPage.Enabled = True
            ChkPartyOnNewPage.Caption = "Is Summery ?"
            chkGenVoucher.Visible = True
            chkGenVoucher.Enabled = True
            chkGenVoucher.Caption = "Transfer Brokerage of No Brokerage to Broker to 9777777 ?"
         End If
         If gSelectedMenu = "Brokerage Summery" Then
            chkGenVoucher.Visible = True
            chkGenVoucher.Enabled = True
            Label11.Visible = True
            dtpVouDt.Visible = True
            dtpVouDt.Enabled = True
            optBrokList(0).Visible = True
            optBrokList(1).Visible = True
            optBrokList(2).Visible = True
            optBrokList(0).Enabled = True
            optBrokList(1).Enabled = True
            optBrokList(2).Enabled = True
            optBrokList(0).Value = True
         End If
         txtCode(3).Enabled = True
         txtCode(3).Visible = True
         txtBroker.Visible = True
         txtBroker.Enabled = True
         lblBroker.Visible = True
         ChkConsolidate.Visible = True
         ChkConsolidate.Enabled = True
         txt(6).Enabled = True
         txt(6).Visible = True
    Case "Trade Purchase TDS Report", "Depot Sales TDS Report", "Mill Bill TDS Report", "Millwise Inward", "Daily Purchase Report", "Daily Purchase Report ( Mill Bill )", "Other Purchase TDS Report"
         cmdCrystal(2).Visible = True
         cmdCrystal(3).Visible = True
         txtCode(0).Visible = True
         txtCode(0).Enabled = True
         txtCode(2).Enabled = True
         txtCode(2).Visible = True
         txtCode(3).Visible = False
         txtMill.Visible = True
         txtParty.Visible = True
         txtMill.Visible = True
         txtMill.Enabled = True
         txtParty.Enabled = True
         lblMill.Visible = True
         lblParty.Visible = True
         ChkConsolidate.Visible = True
         ChkConsolidate.Enabled = True
         optBrokList(1).Visible = True
         optBrokList(1).Value = True
         optBrokList(1).Enabled = True
         optBrokList(1).Caption = "Detail"
         optBrokList(2).Visible = True
         optBrokList(2).Enabled = True
         optBrokList(2).Caption = "Summery"
         optBrokList(0).Visible = True
         optBrokList(0).Enabled = True
         optBrokList(0).Caption = "Daywise"
         If gSelectedMenu = "Depot Sales TDS Report" Or gSelectedMenu = "Mill Bill TDS Report" Then
            chkIsExcludingSale.Visible = True
            chkIsExcludingSale.Enabled = True
            chkIsExcludingSale.Caption = "Only 2nd Day Payment"
            chkGenVoucher.Visible = True
            chkGenVoucher.Enabled = True
            chkGenVoucher.Caption = "Without 2nd Day Payment"
         End If
         If gSelectedMenu = "Millwise Inward" Or gSelectedMenu = "Daily Purchase Report" Or gSelectedMenu = "Daily Purchase Report ( Mill Bill )" Then
            txtCode(2).Visible = False
            txtParty.Visible = False
            lblParty.Visible = False
            optBrokList(0).Visible = False
            optBrokList(1).Visible = False
            optBrokList(2).Visible = False
            cmdCrystal(2).Visible = False
            cmdCrystal(3).Visible = False
            dtpFrDt.Value = Format(Now, "dd/MM/yyyy")
            dtpToDt.Value = Format(Now, "dd/MM/yyyy")
         End If
         If gSelectedMenu = "Daily Purchase Report" Or gSelectedMenu = "Daily Purchase Report ( Mill Bill )" Then
            ChkConsolidate.Value = 1
         End If
End Select
 '-- Set InActivate Ctrl Back Color
GProcSetInActivateCtrlBackClr Me
End Sub
Private Sub cmdExit_Click()
DoEvents
Unload Me
End Sub
Private Sub txtCode_LostFocus(Index As Integer)
If Trim(txtCode(Index)) = "" Then Exit Sub
Select Case Index
    Case 2  '-- Party  Code
         txtParty = GProcGetColumnValue("tblMastAccount", "AcOurCode", UCase(txtCode(2)), "S", "AcName", "S")
         wPartyCode = GProcGetColumnValue("tblMastAccount", "AcName", txtParty, "S", "AcCode", "N")
    Case 0  '-- Mill  Code
         txtMill = GProcGetColumnValue("tblMastAccount", "AcOurCode", UCase(txtCode(0)), "S", "AcName", "S")
    Case 3  '-- Broker  Code
         txtBroker = GProcGetColumnValue("tblMastAccount", "AcOurCode", UCase(txtCode(3)), "S", "AcName", "S")
End Select
'      txtBroker = GProcGetColumnValue("tblMastAccount", "AcOurCode", UCase(txtCode(3)), "S", "AcName", "S")
End Sub

Private Sub txtMill_GotFocus()
GProcSelectBox Me.ActiveControl
End Sub
Private Sub txtBroker_GotFocus()
GProcSelectBox Me.ActiveControl
End Sub
Private Sub txtParty_GotFocus()
GProcSelectBox Me.ActiveControl
End Sub
Private Sub txtItem_GotFocus()
GProcSelectBox Me.ActiveControl
End Sub
'-- mill
Private Sub txtMill_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then Exit Sub '-- Enter
   gClsSearch.SearchMultiField "tblMastaccount", "AcName,AcCity,ACCode", Array("Mill Name", "City", "Code"), Array(3000, 2000, 0), "(AgCode=90029 ) ", Chr(KeyAscii), "AcName", txtMill.Left + Me.Left, txtMill.Top + Me.Top + 1000
KeyAscii = 0
If gClsSearch.SearchMultiRetCol(0) <> "" Then
    txtMill = gClsSearch.SearchMultiRetCol(0)
End If
End Sub
'-- Party
Private Sub txtParty_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then Exit Sub '-- Enter
        
        gClsSearch.SearchMultiFieldAccMaster Chr(KeyAscii), txtParty.Left + Me.Left, txtParty.Top + Me.Top + 1000
        KeyAscii = 0
        If gClsSearch.SearchMultiRetColAccMaster(0) <> "" Then
            txtParty = gClsSearch.SearchMultiRetColAccMaster(0)
       End If
'gClsSearch.SearchMultiField "QryMastAccount", "AcOurCode,Account,CGName,AcContPer,AcCity,AgName,AcBST,ACCode", Array("Code", "Account Name", "Party Group", "Contact Person", "City", "Group", "TIN", "Code"), Array(1000, 4000, 2000, 2000, 2000, 2500, 0, 0), "(Agcode in (90017) )", Chr(KeyAscii), "", txtparty.Left + Me.Left, txtparty.Top + Me.Top + 1000, False
'KeyAscii = 0
'If gClsSearch.SearchMultiRetCol(1) <> "" Then
'    txtparty = gClsSearch.SearchMultiRetCol(1)
'End If
End Sub
'-- Broker
Private Sub txtBroker_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then Exit Sub '-- Enter
gClsSearch.SearchMultiField "tblMastaccount", "AcName,AcCity,ACCode", Array("Broker", "City", ""), Array(5000, 3000, 0), "(AgCode=90016 )", Chr(KeyAscii), "", txtBroker.Left + Me.Left, txtBroker.Top + Me.Top + 1000
'gClsSearch.SearchMultiField "tblMastaccount", "AcOurCode,AcName,AcCity,ACCode", Array("Code", "Broker", "City", ""), Array(1000, 5000, 3000, 0), "(AgCode=90016 )", Chr(KeyAscii), "", txtBroker.Left + Me.Left, txtBroker.Top + Me.Top + 1000
KeyAscii = 0
If gClsSearch.SearchMultiRetCol(0) <> "" Then
    txtBroker = gClsSearch.SearchMultiRetCol(0)
End If
End Sub
'-- Item
Private Sub txtItem_KeyPress(KeyAscii As Integer)
Dim i As Integer, X As String
'If gSelectedMenu = "Count / Brokerwise Sales" Then
'     gClsSearch.SearchMultiSelField "tblMastItem", "distinct ItName", Array("Item Name"), Array(2000)
'     wItList = ""
'     If gClsSearch.SearchMultiSelRetCol(0, 0) <> "" Then
'        For i = 0 To UBound(gClsSearch.SearchMultiSelRetCol)
'            wItList = wItList + "'" & gClsSearch.SearchMultiSelRetCol(i, 0) & "',"
'            txtItem = gClsSearch.SearchMultiSelRetCol(i, 0)
'        Next i
'    End If
'    wItList = Left(wItList, Len(wItList) - 1)
If gSelectedMenu = "C Form Receivable" Then
   gClsSearch.SearchMultiField "tblMastCompGroup", "CGName", Array("Company Group"), Array(2000), " ", Chr(KeyAscii), "CGName", txtItem.Left + Me.Left, txtItem.Top + Me.Top + 1000
   KeyAscii = 0
   If gClsSearch.SearchMultiRetCol(0) <> "" Then
      txtItem = gClsSearch.SearchMultiRetCol(0)
   End If
Else
        If txtMill <> "" Then
           wMillCode = GProcGetColumnValue("tblMastAccount", "AcName", txtMill, "S", "AcCode", "N")
           gClsSearch.SearchMultiField "tblMastItem,tblMastAccount", "ItName,ItCode,Itticket,Acname", Array("Item", "", "Ticket", "Mill"), Array(3000, 0, 0, 3000), "  tblMastItem.ITMillCode=tblMastaccount.acCode and ITMillCode=" & wMillCode, Chr(KeyAscii), "ItName", txtMill.Left + Me.Left, txtMill.Top + Me.Top + 1000
        Else
           gClsSearch.SearchMultiField "tblMastItem,tblMastAccount", "ItName,ItCode,Itticket,Acname", Array("Item", "", "Ticket", "Mill"), Array(3000, 0, 0, 3000), "  tblMastItem.ItMillcode=tblMastAccount.AcCode", Chr(KeyAscii), "ItName", txtMill.Left + Me.Left, txtMill.Top + Me.Top + 1000
        End If
'   gClsSearch.SearchMultiField "tblMastItem,tblMastAccount", "ItCode,ItName,Itticket,Acname", Array("", "Item", "Ticket", "Mill"), Array(0, 2000, 1000, 2000), " tblMastItem.ItMillcode=tblMastAccount.AcCode", "", "ItName", txtItem.Left + Me.Left, txtItem.Top + Me.Top + 1000
   KeyAscii = 0
   If gClsSearch.SearchMultiRetCol(1) <> "" Then
      txtItem = gClsSearch.SearchMultiRetCol(0)
      wItemCode = gClsSearch.SearchMultiRetCol(1) '-- ItCode
   End If
End If
End Sub
Private Sub txt_GotFocus(Index As Integer)
GProcSelectBox Me.ActiveControl
End Sub
Private Sub txt_KeyPress(Index As Integer, KeyAscii As Integer)
If Index < 5 Then KeyAscii = GProcNumberOnly(KeyAscii)
End Sub
Private Sub txt_LostFocus(Index As Integer)
ValidateData
End Sub
'------- Validate Data
Private Sub ValidateData()
Dim i As Long
For i = 0 To 4
    If IsNumeric(txt(i)) = False Then
        txt(i) = "0.00"
    Else
        txt(i) = Format(txt(i), FStr)
    End If
Next i
End Sub
'---- Preview / Print
Private Sub cmdCrystal_Click(Index As Integer)
Dim x1 As String
Dim RptOptStr As String
Dim PrevBrkStr As String
Dim X As String
Dim NetAm As String
Dim NetAmPu As String
Dim OthP As String
Dim OthN As String
Dim OthPurP As String
Dim OthPurN As String
Dim Rs As Recordset
Dim RsTmp As Recordset
x1 = ""
Dim ShowRepo As Boolean
Dim i As Integer
Dim j As Integer
Dim k As Integer
Dim BPaidDate As Date
Dim wJvNo As Long
Dim wParty As Long
                    
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



ShowRepo = False
ValidateData
With CrystalReport1
    GProcCrystalRptPreparation CrystalReport1, IIf(Index = 0, "Window", "Printer")
    Select Case gSelectedMenu
        Case "Total Sales Summary"  '-Total Sales"
                If txtParty <> "" Then
                   wPartyCode = GProcGetColumnValue("tblMastAccount", "AcName", txtParty, "S", "AcCode", "N")
                Else
                   wPartyCode = 0
                End If
                If txtBroker <> "" Then
                   wBrkCode = GProcGetColumnValue("tblMastAccount", "AcName", txtBroker, "S", "AcCode", "N")
                Else
                   wBrkCode = 0
                End If
                If txtMill <> "" Then
                   wMillCode = GProcGetColumnValue("tblMastAccount", "AcName", txtMill, "S", "AcCode", "N")
                Else
                   wMillCode = 0
                End If
                gCmd.CommandText = "PrcPrepareSaleStatus"
                gCmd.Parameters.Refresh
                gCmd.Parameters("@Fr_dt") = Format(dtpFrDt, "yyyy/mm/dd")
                gCmd.Parameters("@To_dt") = Format(dtpToDt, "yyyy/mm/dd")
                gCmd.Parameters("@UCode") = gUserName
                If ChkConsolidate.Value = 1 Then
                   gCmd.Parameters("@VFirm") = ""
                Else
                   gCmd.Parameters("@VFirm") = Trim(gCCode)
                End If
                If wPartyCode <> 0 Then gCmd.Parameters("@Accode") = wPartyCode
                If wBrkCode <> 0 Then gCmd.Parameters("@Accode_Br") = wBrkCode
                If wMillCode <> 0 Then gCmd.Parameters("@MillCode") = wMillCode
                gCmd.Execute
                x1 = "{TmpGentbl.UserName}='" & gUserName & "' "
               .SelectionFormula = x1
               .Formulas(4) = "wFrToDt = 'From : ' & '" & dtpFrDt & "' & '  To ' & '" & dtpToDt & "'"
               If ChkConsolidate.Value = 1 Then
                  .Formulas(5) = "wAllFirm = 'All Firm '"
               Else
                  .Formulas(5) = "wAllFirm = 'Single Firm '"
               End If
               .ReportFileName = gReportPath & "rptSaleSummary.rpt"
        Case "Count / Brokerwise Sales", "Party / Countwise Sales", "Sales Against Booking"
                If txtParty <> "" Then
                   wPartyCode = GProcGetColumnValue("tblMastAccount", "AcName", txtParty, "S", "AcCode", "N")
                Else
                   wPartyCode = 0
                End If
                If txtBroker <> "" Then
                   wBrkCode = GProcGetColumnValue("tblMastAccount", "AcName", txtBroker, "S", "AcCode", "N")
                Else
                   wBrkCode = 0
                End If
                If txtMill <> "" Then
                   wMillCode = GProcGetColumnValue("tblMastAccount", "AcName", txtMill, "S", "AcCode", "N")
                Else
                   wMillCode = 0
                End If
                If txtItem = "" Then
                   wItemCode = 0
                End If
                gCmd.CommandText = "PrcPrepareSaleItemDetail"
                gCmd.Parameters.Refresh
                gCmd.Parameters("@Fr_dt") = Format(dtpFrDt, "yyyy/mm/dd")
                gCmd.Parameters("@To_dt") = Format(dtpToDt, "yyyy/mm/dd")
                gCmd.Parameters("@UCode") = gUserName
                If ChkConsolidate.Value = 1 Then
                   gCmd.Parameters("@VFirm") = ""
                Else
                   gCmd.Parameters("@VFirm") = Trim(gCCode)
                End If
                If wPartyCode <> 0 Then gCmd.Parameters("@Accode") = wPartyCode
                If wBrkCode <> 0 Then gCmd.Parameters("@Accode_Br") = wBrkCode
                If wMillCode <> 0 Then gCmd.Parameters("@MillCode") = wMillCode
                If wItemCode <> 0 Then gCmd.Parameters("@ItemCode") = wItemCode
                gCmd.Execute
                If gSelectedMenu = "Party / Countwise Sales" And ChkPartyOnNewPage.Value = 1 Then
                   gCn.Execute "delete from TmpGentbl2  where UserName = '" & gUserName & "'   "
                   
                   X = "Select PartyName,PartyCode,sum(Amt2),'" & gUserName & "' from tmpGenTbl where UserName = '" & gUserName & "' group by PartyName,PartyCode"
                   gCn.Execute "Insert into tmpGenTbl2(PartyName,PartyCode,Amt2,UserName) " & X
                   X = "{TmpGentbl2.UserName}='" & gUserName & "' "
                   .SelectionFormula = X
                   .Formulas(4) = "wFrToDt = 'From : ' & '" & dtpFrDt & "' & '  To ' & '" & dtpToDt & "'"
                   If ChkConsolidate.Value = 1 Then
                      .Formulas(5) = "wAllFirm = 'All Firm '"
                   Else
                      .Formulas(5) = "wAllFirm = 'Single Firm '"
                   End If
                   .ReportFileName = gReportPath & "rptSlPtyCountSum.rpt"
                   
                   
                Else
                  x1 = "{TmpGentbl.UserName}='" & gUserName & "' "
                  .SelectionFormula = x1
                  .Formulas(4) = "wFrToDt = 'From : ' & '" & dtpFrDt & "' & '  To ' & '" & dtpToDt & "'"
                  If ChkConsolidate.Value = 1 Then
                     .Formulas(5) = "wAllFirm = 'All Firm '"
                  Else
                     .Formulas(5) = "wAllFirm = 'Single Firm '"
                  End If
                  If gSelectedMenu = "Party / Countwise Sales" Then
                    .ReportFileName = gReportPath & "rptSalePartyCountwise.rpt"
                  ElseIf gSelectedMenu = "Sales Against Booking" Then
                     x1 = x1 + " and {TmpGentbl.VSrNo} <> 0 "
                    .SelectionFormula = x1
                    .ReportFileName = gReportPath & "rptSaleAgtBooking.rpt"
                  Else
                    .ReportFileName = gReportPath & "rptSaleCountBrokerwise.rpt"
                  End If
                End If
    Case "Broker Ledger"
         Dim wAgCode As Long
         Dim wAcCode As Long
         Dim wAcOurCode As String
         Dim wBrokerageAcName As String
         Dim wEBRSLS As Long
         Dim rstTmp As Recordset
         If txtBroker = "" Then
            wAgCode = 90016
            wAcCode = 0
         Else
            wAgCode = 0
            wAcCode = GProcGetColumnValue("tblMastAccount", "AcName", txtBroker, "S", "AcCode", "N")
            wAcOurCode = GProcGetColumnValue("tblMastAccount", "AcName", txtBroker, "S", "AcOurCode", "S")
         End If
         wBrokerageAcName = GProcGetColumnValue("tblMastAccount", "AcCode", CStr(gBrokerageSaleAc), "N", "AcName", "S")
         
         wEBRSLS = GProcGetColumnValue("tblMastAccount", "AcOurCode", "EBRSLS", "S", "AcCode", "N")
         
         gCn.Execute "delete from TmpAccLedger  where UserName = '" & gUserName & "'   "
        
         gCn.Execute "delete from tmpSelectionCodes where UserName = '" & gUserName & "'"
         i = Len(txt(6))
         If i = 0 And txtBroker <> "" Then
            wItList = GProcGetColumnValue("tblMastAccount", "AcName", txtBroker, "S", "AcOurCode", "S")
            gCn.Execute "Insert into tmpSelectionCodes (USERNAME,OurCode ) Values('" & gUserName & "','" & wItList & "')"
         End If
         If i = 0 And txtBroker = "" Then
            gCn.Execute "Insert into tmpSelectionCodes (USERNAME,OurCode ) select '" & gUserName & "',AcOurCode from tblMastAccount where AgCode=  90016"
         End If
         If i > 0 Then
            j = 1
            X = txt(6)
            Do While j <= i
               k = InStr(X, ",")
               x1 = Mid(X, j, k - 1)
               gCn.Execute "Insert into tmpSelectionCodes (USERNAME,OurCode ) Values('" & gUserName & "','" & x1 & "')"
               j = j + k
            Loop
         End If
        '-- For Op Balance
         X = "select * from tblMastCompany"
         GProcRstOpen rstTmp, X, "R"
         rstTmp.MoveFirst
         Do While Not rstTmp.EOF
            If (rstTmp.Fields("CDepotMainFirmCompCd") = gCCode And ChkConsolidate = 0) Then
               gCmd.CommandText = "PrcPrepareAccLedgerBrokSaleOpBal"
               gCmd.Parameters.Refresh
               gCmd.Parameters("@From_dt") = Format(gCYSDate, "yyyy/mm/dd")
               gCmd.Parameters("@To_dt") = Format(dtpFrDt, "yyyy/mm/dd")
               gCmd.Parameters("@YrEnd_dt") = Format(gCYEDate, "yyyy/mm/dd")
               gCmd.Parameters("@UCode") = gUserName
'               If ChkConsolidate.Value = 1 Then
 '                 gCmd.Parameters("@VFirm") = ""
  '             Else
                  gCmd.Parameters("@VFirm") = rstTmp.Fields("CCode")
   '            End If
               If ChkPartyOnNewPage.Value = 1 Then
                  gCmd.Parameters("@BrkTp") = "S"
               Else
                  gCmd.Parameters("@BrkTp") = ""
               End If
               If rstTmp.Fields("CDepotMainFirmCompCd") <> 0 Then
                  gCmd.Parameters("@DepotFirm") = "D"
               Else
                  gCmd.Parameters("@DepotFirm") = "F"
               End If
               If wAcCode <> 0 Then gCmd.Parameters("@Accode") = wAcCode
               If wAgCode <> 0 Then gCmd.Parameters("@Agcode") = wAgCode
               If wAcOurCode = "9000" Then
                  gCmd.Parameters("@AcOurCode") = "9000"
               Else
                  gCmd.Parameters("@AcOurCode") = ""
               End If
               If chkGenVoucher.Value = 1 Then
                  gCmd.Parameters("@TfrTo9777777") = "1"
               Else
                  gCmd.Parameters("@TfrTo9777777") = "0"
               End If
               gCmd.Parameters("@BrokkerageAc") = wBrokerageAcName
               gCmd.Parameters("@EBRSLSCode") = wEBRSLS
               gCmd.Execute
            End If
            If ChkConsolidate = 1 Then
'            If (rstTmp.Fields("CDepotMainFirmCompCd") <> 0 And ChkConsolidate = 1) Then
               gCmd.CommandText = "PrcPrepareAccLedgerBrokSaleOpBal"
               gCmd.Parameters.Refresh
               gCmd.Parameters("@From_dt") = Format(gCYSDate, "yyyy/mm/dd")
               gCmd.Parameters("@To_dt") = Format(dtpFrDt, "yyyy/mm/dd")
               gCmd.Parameters("@YrEnd_dt") = Format(gCYEDate, "yyyy/mm/dd")
               gCmd.Parameters("@UCode") = gUserName
'               If ChkConsolidate.Value = 1 Then
 '                 gCmd.Parameters("@VFirm") = ""
  '             Else
                  gCmd.Parameters("@VFirm") = rstTmp.Fields("CCode")
   '            End If
               If ChkPartyOnNewPage.Value = 1 Then
                  gCmd.Parameters("@BrkTp") = "S"
               Else
                  gCmd.Parameters("@BrkTp") = ""
               End If
               If rstTmp.Fields("CDepotMainFirmCompCd") <> 0 Then
                  gCmd.Parameters("@DepotFirm") = "D"
               Else
                  gCmd.Parameters("@DepotFirm") = "F"
               End If
               If wAcCode <> 0 Then gCmd.Parameters("@Accode") = wAcCode
               If wAgCode <> 0 Then gCmd.Parameters("@Agcode") = wAgCode
               If wAcOurCode = "9000" Then
                  gCmd.Parameters("@AcOurCode") = "9000"
               Else
                  gCmd.Parameters("@AcOurCode") = ""
               End If
               If chkGenVoucher.Value = 1 Then
                  gCmd.Parameters("@TfrTo9777777") = "1"
               Else
                  gCmd.Parameters("@TfrTo9777777") = "0"
               End If
               
               gCmd.Parameters("@BrokkerageAc") = wBrokerageAcName
               gCmd.Parameters("@EBRSLSCode") = wEBRSLS
               gCmd.Execute
            End If
            rstTmp.MoveNext
         Loop
         If ChkConsolidate = 0 Then
            gCmd.CommandText = "PrcPrepareAccLedgerBrokSaleOpBal"
            gCmd.Parameters.Refresh
            gCmd.Parameters("@From_dt") = Format(gCYSDate, "yyyy/mm/dd")
            gCmd.Parameters("@To_dt") = Format(dtpFrDt, "yyyy/mm/dd")
            gCmd.Parameters("@YrEnd_dt") = Format(gCYEDate, "yyyy/mm/dd")
            gCmd.Parameters("@UCode") = gUserName
            If ChkConsolidate.Value = 1 Then
               gCmd.Parameters("@VFirm") = ""
            Else
               gCmd.Parameters("@VFirm") = Trim(gCCode)
            End If
            If ChkPartyOnNewPage.Value = 1 Then
               gCmd.Parameters("@BrkTp") = "S"
            Else
               gCmd.Parameters("@BrkTp") = ""
            End If
            gCmd.Parameters("@DepotFirm") = "F"
            If wAcCode <> 0 Then gCmd.Parameters("@Accode") = wAcCode
            If wAgCode <> 0 Then gCmd.Parameters("@Agcode") = wAgCode
            If wAcOurCode = "9000" Then
               gCmd.Parameters("@AcOurCode") = "9000"
            Else
               gCmd.Parameters("@AcOurCode") = ""
            End If
            If chkGenVoucher.Value = 1 Then
               gCmd.Parameters("@TfrTo9777777") = "1"
            Else
               gCmd.Parameters("@TfrTo9777777") = "0"
            End If
            gCmd.Parameters("@BrokkerageAc") = wBrokerageAcName
            gCmd.Parameters("@EBRSLSCode") = wEBRSLS
            gCmd.Execute
       End If
       X = "SELECT   2,Vfirm,VFirmName ,'0P','" & Format(dtpFrDt, "yyyy/mm/dd") & "',AcName,VAcPan ,'Opening Balance', " _
              & " case when sum(VDrAmt)-sum(VCrAmt) > 0 then sum(VDrAmt)-sum(VCrAmt) else 0 end, " _
              & " case when sum(VCrAmt)-sum(VDrAmt) > 0 then sum(VCrAmt)-sum(VDrAmt) else 0 end, '" & gUserName & "' From tmpAccLedger where  Vsrno=1  and username= '" & gUserName & "' " _
              & "group by Vfirm,VFirmName,AcName,VAcPan Having (Sum(VDrAmt) - Sum(VCrAmt)) <> 0"
       
       
'       X = "SELECT   2,'" & Trim(gCCode) & "' ,'" & gCName & "' ,'0P','" & Format(dtpFrDt, "yyyy/mm/dd") & "',AcName,VAcPan ,'Opening Balance', " _
'              & " case when sum(VDrAmt)-sum(VCrAmt) > 0 then sum(VDrAmt)-sum(VCrAmt) else 0 end, " _
'              & " case when sum(VCrAmt)-sum(VDrAmt) > 0 then sum(VCrAmt)-sum(VDrAmt) else 0 end, '" & gUserName & "' From tmpAccLedger where  Vsrno=1  and username= '" & gUserName & "' " _
'              & "group by AcName,VAcPan Having (Sum(VDrAmt) - Sum(VCrAmt)) <> 0"
       gCn.BeginTrans
       gCn.Execute " insert into tmpAccLedger (Vsrno,Vfirm,VFirmName,Vtype,Vdt,AcName,VAcPan ,AcOName,VDrAmt,VCrAmt,username) " & X
       gCn.CommitTrans
        
        '--End  For Op Balance
         
         X = "select * from tblMastCompany"
         GProcRstOpen rstTmp, X, "R"
         rstTmp.MoveFirst
         Do While Not rstTmp.EOF
            If (rstTmp.Fields("CDepotMainFirmCompCd") = gCCode And ChkConsolidate = 0) Then
               gCmd.CommandText = "PrcPrepareAccLedgerBrokSale"
               gCmd.Parameters.Refresh
               gCmd.Parameters("@From_dt") = Format(dtpFrDt, "yyyy/mm/dd")
               gCmd.Parameters("@To_dt") = Format(dtpToDt, "yyyy/mm/dd")
               gCmd.Parameters("@YrEnd_dt") = Format(gCYEDate, "yyyy/mm/dd")
               gCmd.Parameters("@UCode") = gUserName
               gCmd.Parameters("@VFirm") = rstTmp.Fields("CCode")
               If ChkPartyOnNewPage.Value = 1 Then
                  gCmd.Parameters("@BrkTp") = "S"
               Else
                  gCmd.Parameters("@BrkTp") = ""
               End If
               If rstTmp.Fields("CDepotMainFirmCompCd") <> 0 Then
                  gCmd.Parameters("@DepotFirm") = "D"
               Else
                  gCmd.Parameters("@DepotFirm") = "F"
               End If
               If wAcCode <> 0 Then gCmd.Parameters("@Accode") = wAcCode
               If wAgCode <> 0 Then gCmd.Parameters("@Agcode") = wAgCode
               If wAcOurCode = "9000" Then
                  gCmd.Parameters("@AcOurCode") = "9000"
               Else
                  gCmd.Parameters("@AcOurCode") = ""
               End If
               If chkGenVoucher.Value = 1 Then
                  gCmd.Parameters("@TfrTo9777777") = "1"
               Else
                  gCmd.Parameters("@TfrTo9777777") = "0"
               End If
               gCmd.Parameters("@BrokkerageAc") = wBrokerageAcName
               gCmd.Parameters("@EBRSLSCode") = wEBRSLS
               gCmd.Execute
            End If
            If ChkConsolidate = 1 Then
               gCmd.CommandText = "PrcPrepareAccLedgerBrokSale"
               gCmd.Parameters.Refresh
               gCmd.Parameters("@From_dt") = Format(dtpFrDt, "yyyy/mm/dd")
               gCmd.Parameters("@To_dt") = Format(dtpToDt, "yyyy/mm/dd")
               gCmd.Parameters("@YrEnd_dt") = Format(gCYEDate, "yyyy/mm/dd")
               gCmd.Parameters("@UCode") = gUserName
               gCmd.Parameters("@VFirm") = rstTmp.Fields("CCode")
               If ChkPartyOnNewPage.Value = 1 Then
                  gCmd.Parameters("@BrkTp") = "S"
               Else
                  gCmd.Parameters("@BrkTp") = ""
               End If
               If rstTmp.Fields("CDepotMainFirmCompCd") <> 0 Then
                  gCmd.Parameters("@DepotFirm") = "D"
               Else
                  gCmd.Parameters("@DepotFirm") = "F"
               End If
               If wAcCode <> 0 Then gCmd.Parameters("@Accode") = wAcCode
               If wAgCode <> 0 Then gCmd.Parameters("@Agcode") = wAgCode
               If wAcOurCode = "9000" Then
                  gCmd.Parameters("@AcOurCode") = "9000"
               Else
                  gCmd.Parameters("@AcOurCode") = ""
               End If
               If chkGenVoucher.Value = 1 Then
                  gCmd.Parameters("@TfrTo9777777") = "1"
               Else
                  gCmd.Parameters("@TfrTo9777777") = "0"
               End If
               gCmd.Parameters("@BrokkerageAc") = wBrokerageAcName
               gCmd.Parameters("@EBRSLSCode") = wEBRSLS
               gCmd.Execute
            End If
            rstTmp.MoveNext
         Loop
         If ChkConsolidate = 0 Then
            gCmd.CommandText = "PrcPrepareAccLedgerBrokSale"
            gCmd.Parameters.Refresh
            gCmd.Parameters("@From_dt") = Format(dtpFrDt, "yyyy/mm/dd")
            gCmd.Parameters("@To_dt") = Format(dtpToDt, "yyyy/mm/dd")
            gCmd.Parameters("@YrEnd_dt") = Format(gCYEDate, "yyyy/mm/dd")
            gCmd.Parameters("@UCode") = gUserName
            If ChkConsolidate.Value = 1 Then
               gCmd.Parameters("@VFirm") = ""
            Else
               gCmd.Parameters("@VFirm") = Trim(gCCode)
            End If
            If ChkPartyOnNewPage.Value = 1 Then
               gCmd.Parameters("@BrkTp") = "S"
            Else
               gCmd.Parameters("@BrkTp") = ""
            End If
            gCmd.Parameters("@DepotFirm") = "F"
            If wAcCode <> 0 Then gCmd.Parameters("@Accode") = wAcCode
            If wAgCode <> 0 Then gCmd.Parameters("@Agcode") = wAgCode
            If wAcOurCode = "9000" Then
               gCmd.Parameters("@AcOurCode") = "9000"
            Else
               gCmd.Parameters("@AcOurCode") = ""
            End If
            If chkGenVoucher.Value = 1 Then
               gCmd.Parameters("@TfrTo9777777") = "1"
            Else
               gCmd.Parameters("@TfrTo9777777") = "0"
            End If
            gCmd.Parameters("@BrokkerageAc") = wBrokerageAcName
            gCmd.Parameters("@EBRSLSCode") = wEBRSLS
            gCmd.Execute
         End If
         X = "{tmpaccledger.UserName}= '" & gUserName & "' "
         .SelectionFormula = X
         .Formulas(4) = "wFrToDt = 'From : ' & '" & dtpFrDt & "' & '  To ' & '" & dtpToDt & "'"
         If ChkPartyOnNewPage.Value = 1 Then
                .Formulas(5) = "wRptName = 'Account Ledger (Brokerage)  Summery  '"
                .Formulas(6) = "wOpDt = '" & dtpFrDt & "' "
                If chkGenVoucher.Value = 1 Then
'                   gCn.Execute "delete from TmpAccLedger  where OrderNo =1 and VDrAmt > 0 and vdt >= '" & Format(dtpFrDt, "dd/MM/yyyy") & "' and isnull(vtype,'') not in ('0P') and UserName = '" & gUserName & "'  "

'                   gCn.BeginTrans
'                   gCn.Execute "delete from TmpAccLedger  where OrderNo =1 and VCrAmt > 0  and isnull(vtype,'') not in ('0P') and UserName = '" & gUserName & "'  "
'                   gCn.CommitTrans
                   
                   
                   gCn.BeginTrans
'                   gCn.Execute "delete from TmpAccLedger  where  VCrAmt+ vdramt > 0  and isnull(vtype,'') not in ('0P') and UserName = '" & gUserName & "'  "
                   gCn.CommitTrans
                
                End If
                .ReportFileName = gReportPath & "rptAccountLedgerBrokerageSummery.rpt"
         Else
                .Formulas(5) = "wRptName = 'Account Ledger (Brokerage)'"
                If ChkConsolidate.Value = 1 Then
                  .ReportFileName = gReportPath & "rptAccountLedgerBrokerageAllFirm.rpt"
                Else
                  .ReportFileName = gReportPath & "rptAccountLedgerBrokerage.rpt"
                End If
          End If
    Case "Brokerage Detail ", "Brokerage Summery", "Brokerage Voucher", "Countwise Brokarage"
         BPaidDate = CDate("" & gCYear & "/10/01")
         gCn.BeginTrans
         gCn.Execute "update tblSale set SlTmpNo3 = 0,SlPrevBrokPd = 'N' where vyear = " & gCYear & ""
         gCn.CommitTrans
         If txtBroker = "" Then
            wAcCode = 0
         Else
            wAcCode = GProcGetColumnValue("tblMastAccount", "AcName", txtBroker, "S", "AcCode", "N")
         End If
         gCn.Execute "delete from TmpGentbl  where UserName = '" & gUserName & "'   "
         gCn.Execute "delete from tmpSelectionCodes where UserName = '" & gUserName & "'"
         i = Len(txt(6))
         If i = 0 And txtBroker <> "" Then
            wItList = GProcGetColumnValue("tblMastAccount", "AcName", txtBroker, "S", "AcOurCode", "S")
            gCn.Execute "Insert into tmpSelectionCodes (USERNAME,OurCode ) Values('" & gUserName & "','" & wItList & "')"
         End If
         If i = 0 And txtBroker = "" Then
            gCn.Execute "Insert into tmpSelectionCodes (USERNAME,OurCode ) select '" & gUserName & "',AcOurCode from tblMastAccount where AgCode=  90016"
         End If
         If i > 0 Then
            j = 1
            X = txt(6)
            Do While j <= i
               k = InStr(X, ",")
               x1 = Mid(X, j, k - 1)
               gCn.Execute "Insert into tmpSelectionCodes (USERNAME,OurCode ) Values('" & gUserName & "','" & x1 & "')"
               j = j + k
            Loop
         End If
         If ChkConsolidate = 1 Then
            X = "select * from tblMastCompany where CIsDepotFirm = 0 "
            GProcRstOpen rstTmp, X, "R"
            rstTmp.MoveFirst
            Do While Not rstTmp.EOF
               Mcp = rstTmp.Fields("CCode")
               x1 = "select * from tblMastCompany where CDepotMainFirmCompCd = '" & Mcp & "'"
               i = GProcRstOpen(Rs, x1, "R")
               If i > 0 Then
                  Rs.MoveFirst
                  Do While Not Rs.EOF
                     If chkIsExcludingSale And gSelectedMenu = "Brokerage Voucher" Then
                        gCn.BeginTrans
                        gCn.Execute "update tblSale set SlTmpNo3 = 1 where vfirm = '" & Rs.Fields("CCode") & "' and vyear = " & gCYear & " and vdt < '" & Format(BPaidDate, "yyyy/mm/dd") & "'  and  SlBroker in (select VCrAcCode from tblVoucher where vtype = 'JT' and vfirm = '" & Mcp & "' and vyear = " & gCYear & " and vdt < '" & Format(gCYEDate, "YYYY/MM/DD") & "'  and VDrAcCode = 15506) "
                        gCn.Execute "update tblSale set SlPrevBrokPd = 'Y' where vfirm = '" & Rs.Fields("CCode") & "' and vyear = " & gCYear & "  and  SlBroker in (select VCrAcCode from tblVoucher where vtype = 'JT' and vfirm = '" & Mcp & "' and vyear = " & gCYear & " and vdt < '" & Format(gCYEDate, "YYYY/MM/DD") & "'  and VDrAcCode = 15506) "
                        gCn.CommitTrans
                      End If
                      gCmd.CommandText = "PrcPrepareBrokSale"
                      gCmd.Parameters.Refresh
                      gCmd.Parameters("@Start_dt") = IIf(gSelectedMenu = "Countwise Brokarage" Or gSelectedMenu = "Brokerage Summery", Format(dtpFrDt, "yyyy/mm/dd"), Format(gCYSDate, "yyyy/mm/dd"))
                      gCmd.Parameters("@From_dt") = Format(dtpFrDt, "yyyy/mm/dd")
                      gCmd.Parameters("@To_dt") = Format(dtpToDt, "yyyy/mm/dd")
                      gCmd.Parameters("@TDSLimit") = Val(txt(0))
                      gCmd.Parameters("@TDSRate") = Val(txt(1))
                      gCmd.Parameters("@UCode") = gUserName
                      gCmd.Parameters("@VFirm") = Rs.Fields("CCode")
                      If gSelectedMenu = "Brokerage Detail " Then gCmd.Parameters("@BrkTp") = "D"
                      If gSelectedMenu = "Brokerage Summery" Then gCmd.Parameters("@BrkTp") = "V"
                      If gSelectedMenu = "Brokerage Voucher" Then gCmd.Parameters("@BrkTp") = "V"
                      If gSelectedMenu = "Countwise Brokarage" Then gCmd.Parameters("@BrkTp") = "V"
                      gCmd.Parameters("@VFirmAdd1") = Trim(gCAdd1)
                      gCmd.Parameters("@VFirmAdd2") = Trim(gCAdd2)
                      gCmd.Parameters("@VFirmAdd3") = Trim(gCAdd3)
                      gCmd.Parameters("@VFirmCph") = Trim(gCPhNo)
                      gCmd.Parameters("@VFirmGstin") = Trim(gCGStin)
                      gCmd.Parameters("@DepotFirm") = "D"
                      If wAcCode <> 0 Then gCmd.Parameters("@Accode") = wAcCode
                      gCmd.Execute
                    Rs.MoveNext
                  Loop
                End If
                    If chkIsExcludingSale Then
                       gCn.BeginTrans
                       gCn.Execute "update tblSale set SlTmpNo3 = 1       where vFirm = '" & Mcp & "' and vyear = " & gCYear & " and vdt < '" & Format(BPaidDate, "yyyy/mm/dd") & "' and SlBroker in (select VCrAcCode from tblVoucher where vtype = 'JT' and vfirm = '" & Mcp & "' and vyear = " & gCYear & " and vdt < '" & Format(gCYEDate, "YYYY/MM/DD") & "'  and VDrAcCode = 15506)"
                       gCn.Execute "update tblSale set SlPrevBrokPd = 'Y' where vfirm = '" & Mcp & "' and vyear = " & gCYear & "  and  SlBroker in (select VCrAcCode from tblVoucher where vtype = 'JT' and vfirm = '" & Mcp & "' and vyear = " & gCYear & " and vdt < '" & Format(gCYEDate, "YYYY/MM/DD") & "'  and VDrAcCode = 15506) "
                       gCn.CommitTrans
                     End If
                    gCmd.CommandText = "PrcPrepareBrokSale"
                    gCmd.Parameters.Refresh
                    gCmd.Parameters("@Start_dt") = IIf(gSelectedMenu = "Countwise Brokarage" Or gSelectedMenu = "Brokerage Summery", Format(dtpFrDt, "yyyy/mm/dd"), Format(gCYSDate, "yyyy/mm/dd"))
                    gCmd.Parameters("@From_dt") = Format(dtpFrDt, "yyyy/mm/dd")
                    gCmd.Parameters("@To_dt") = Format(dtpToDt, "yyyy/mm/dd")
                    gCmd.Parameters("@TDSLimit") = Val(txt(0))
                    gCmd.Parameters("@TDSRate") = Val(txt(1))
                    gCmd.Parameters("@UCode") = gUserName
                    gCmd.Parameters("@DepotFirm") = "F"
                    If gSelectedMenu = "Brokerage Detail " Then gCmd.Parameters("@BrkTp") = "D"
                    If gSelectedMenu = "Brokerage Summery" Then gCmd.Parameters("@BrkTp") = "V"
                    If gSelectedMenu = "Brokerage Voucher" Then gCmd.Parameters("@BrkTp") = "V"
                    If gSelectedMenu = "Countwise Brokarage" Then gCmd.Parameters("@BrkTp") = "V"
                    gCmd.Parameters("@VFirm") = Mcp
                    gCmd.Parameters("@VFirmAdd1") = Trim(gCAdd1)
                    gCmd.Parameters("@VFirmAdd2") = Trim(gCAdd2)
                    gCmd.Parameters("@VFirmAdd3") = Trim(gCAdd3)
                    gCmd.Parameters("@VFirmCph") = Trim(gCPhNo)
                    gCmd.Parameters("@VFirmGstin") = Trim(gCGStin)
                    If wAcCode <> 0 Then gCmd.Parameters("@Accode") = wAcCode
                    gCmd.Execute
'               End If
               rstTmp.MoveNext
            Loop
         Else
            X = "select * from tblMastCompany"
            GProcRstOpen rstTmp, X, "R"
            rstTmp.MoveFirst
            Do While Not rstTmp.EOF
               If (rstTmp.Fields("CDepotMainFirmCompCd") = gCCode And ChkConsolidate = 0) Then
                   If chkIsExcludingSale Then
                      gCn.BeginTrans
                      gCn.Execute "update tblSale set SlTmpNo3 = 1 where vfirm = '" & rstTmp.Fields("CCode") & "' and vyear = " & gCYear & " and vdt < '" & Format(BPaidDate, "yyyy/mm/dd") & "'  and  SlBroker in (select VCrAcCode from tblVoucher where vtype = 'JT' and vfirm = '" & gCCode & "' and vyear = " & gCYear & " and vdt < '" & Format(gCYEDate, "YYYY/MM/DD") & "'  and VDrAcCode = 15506) "
                      gCn.Execute "update tblSale set SlPrevBrokPd = 'Y' where vfirm = '" & rstTmp.Fields("CCode") & "' and vyear = " & gCYear & "  and  SlBroker in (select VCrAcCode from tblVoucher where vtype = 'JT' and vfirm = '" & gCCode & "' and vyear = " & gCYear & " and vdt < '" & Format(gCYEDate, "YYYY/MM/DD") & "'  and VDrAcCode = 15506) "
                      gCn.CommitTrans
                   End If
                   gCmd.CommandText = "PrcPrepareBrokSale"
                   gCmd.Parameters.Refresh
                   gCmd.Parameters("@Start_dt") = IIf(gSelectedMenu = "Countwise Brokarage" Or gSelectedMenu = "Brokerage Summery", Format(dtpFrDt, "yyyy/mm/dd"), Format(gCYSDate, "yyyy/mm/dd"))
                   gCmd.Parameters("@From_dt") = Format(dtpFrDt, "yyyy/mm/dd")
                   gCmd.Parameters("@To_dt") = Format(dtpToDt, "yyyy/mm/dd")
                   gCmd.Parameters("@TDSLimit") = Val(txt(0))
                   gCmd.Parameters("@TDSRate") = Val(txt(1))
                   gCmd.Parameters("@UCode") = gUserName
                   gCmd.Parameters("@VFirm") = rstTmp.Fields("CCode")
                   If gSelectedMenu = "Brokerage Detail " Then gCmd.Parameters("@BrkTp") = "D"
                   If gSelectedMenu = "Brokerage Summery" Then gCmd.Parameters("@BrkTp") = "V"
                   If gSelectedMenu = "Brokerage Voucher" Then gCmd.Parameters("@BrkTp") = "V"
                   If gSelectedMenu = "Countwise Brokarage" Then gCmd.Parameters("@BrkTp") = "V"
                   If ChkConsolidate.Value = 1 Then
                      gCmd.Parameters("@VFirm") = ""
                      gCmd.Parameters("@VFirmAdd1") = ""
                      gCmd.Parameters("@VFirmAdd2") = ""
                      gCmd.Parameters("@VFirmAdd3") = ""
                      gCmd.Parameters("@VFirmCph") = ""
                      gCmd.Parameters("@VFirmGstin") = ""
                   Else
                      gCmd.Parameters("@VFirmAdd1") = Trim(gCAdd1)
                      gCmd.Parameters("@VFirmAdd2") = Trim(gCAdd2)
                      gCmd.Parameters("@VFirmAdd3") = Trim(gCAdd3)
                      gCmd.Parameters("@VFirmCph") = Trim(gCPhNo)
                      gCmd.Parameters("@VFirmGstin") = Trim(gCGStin)
                   End If
                   If rstTmp.Fields("CDepotMainFirmCompCd") <> 0 Then
                      gCmd.Parameters("@DepotFirm") = "D"
                   Else
                      gCmd.Parameters("@DepotFirm") = "F"
                   End If
                   If wAcCode <> 0 Then gCmd.Parameters("@Accode") = wAcCode
                   gCmd.Execute
               End If
               rstTmp.MoveNext
            Loop
            If ChkConsolidate = 0 Then
               If chkIsExcludingSale Then
                  gCn.BeginTrans
                  gCn.Execute "update tblSale set SlTmpNo3 = 1       where vFirm = '" & Trim(gCCode) & "' and vyear = " & gCYear & " and vdt < '" & Format(BPaidDate, "yyyy/mm/dd") & "' and SlBroker in (select VCrAcCode from tblVoucher where vtype = 'JT' and vfirm = '" & gCCode & "' and vyear = " & gCYear & " and vdt < '" & Format(gCYEDate, "YYYY/MM/DD") & "'  and VDrAcCode = 15506)"
                  gCn.Execute "update tblSale set SlPrevBrokPd = 'Y' where vfirm = '" & Trim(gCCode) & "' and vyear = " & gCYear & "  and  SlBroker in (select VCrAcCode from tblVoucher where vtype = 'JT' and vfirm = '" & gCCode & "' and vyear = " & gCYear & " and vdt < '" & Format(gCYEDate, "YYYY/MM/DD") & "'  and VDrAcCode = 15506) "
                  gCn.CommitTrans
                End If
               gCmd.CommandText = "PrcPrepareBrokSale"
               gCmd.Parameters.Refresh
               gCmd.Parameters("@Start_dt") = IIf(gSelectedMenu = "Countwise Brokarage" Or gSelectedMenu = "Brokerage Summery", Format(dtpFrDt, "yyyy/mm/dd"), Format(gCYSDate, "yyyy/mm/dd"))
               gCmd.Parameters("@From_dt") = Format(dtpFrDt, "yyyy/mm/dd")
               gCmd.Parameters("@To_dt") = Format(dtpToDt, "yyyy/mm/dd")
               gCmd.Parameters("@TDSLimit") = Val(txt(0))
               gCmd.Parameters("@TDSRate") = Val(txt(1))
               gCmd.Parameters("@UCode") = gUserName
               gCmd.Parameters("@DepotFirm") = "F"
               If gSelectedMenu = "Brokerage Detail " Then gCmd.Parameters("@BrkTp") = "D"
               If gSelectedMenu = "Brokerage Summery" Then gCmd.Parameters("@BrkTp") = "V"
               If gSelectedMenu = "Brokerage Voucher" Then gCmd.Parameters("@BrkTp") = "V"
               If gSelectedMenu = "Countwise Brokarage" Then gCmd.Parameters("@BrkTp") = "V"
               If ChkConsolidate.Value = 1 Then
                  gCmd.Parameters("@VFirm") = ""
                  gCmd.Parameters("@VFirmAdd1") = ""
                  gCmd.Parameters("@VFirmAdd2") = ""
                  gCmd.Parameters("@VFirmAdd3") = ""
                  gCmd.Parameters("@VFirmCph") = ""
                  gCmd.Parameters("@VFirmGstin") = ""
               Else
                  gCmd.Parameters("@VFirm") = Trim(gCCode)
                  gCmd.Parameters("@VFirmAdd1") = Trim(gCAdd1)
                  gCmd.Parameters("@VFirmAdd2") = Trim(gCAdd2)
                  gCmd.Parameters("@VFirmAdd3") = Trim(gCAdd3)
                  gCmd.Parameters("@VFirmCph") = Trim(gCPhNo)
                  gCmd.Parameters("@VFirmGstin") = Trim(gCGStin)
               End If
               If wAcCode <> 0 Then gCmd.Parameters("@Accode") = wAcCode
               gCmd.Execute
            End If
         End If
         X = "{TmpGentbl.UserName}= '" & gUserName & "' "
         .SelectionFormula = X
         .Formulas(4) = "wFrToDt = 'From : ' & '" & dtpFrDt & "' & '  To ' & '" & dtpToDt & "'"
         .Formulas(5) = "wRptName = 'Brokerage On Sales Detail'"
         .Formulas(6) = "wTdsRt = '" & txt(1) & "'"
         If gSelectedMenu = "Brokerage Detail " Then .ReportFileName = gReportPath & "rptBrokerageDetail.rpt"
         If gSelectedMenu = "Brokerage Voucher" Then
            .Formulas(7) = "wGTot = " & ChkPartyOnNewPage.Value & ""
            If ChkConsolidate = 1 Then
              .Formulas(8) = "wFrToDt2 = 'From : ' & '" & BPaidDate & "' & '  To ' & '" & dtpToDt & "'"
              .ReportFileName = gReportPath & "rptBrokerageVoucherAllFirm.rpt"
            Else
              .Formulas(8) = "wFrToDt2 = 'From : ' & '" & BPaidDate & "' & '  To ' & '" & dtpToDt & "'"
              .ReportFileName = gReportPath & "rptBrokerageVoucher.rpt"
            End If
         End If
         If gSelectedMenu = "Brokerage Summery" Then
            
            If optBrokList(0).Value = True Then
               gCn.BeginTrans
               gCn.Execute "delete from TmpGentbl  where UserName = '" & gUserName & "' and amt13 = 1  "
               gCn.CommitTrans
            End If
            If optBrokList(1).Value = True Then
               gCn.BeginTrans
               gCn.Execute "delete from TmpGentbl  where UserName = '" & gUserName & "' and amt13 = 0  "
               gCn.CommitTrans
            End If
            
            If chkGenVoucher.Value = 1 Then
               Dim wTdsOn As Double
               Dim wTds As Double
               Dim TdsCrCd As Long
               wTds = 0
               wTdsOn = 0
               TdsCrCd = GProcGetColumnValue("TblMastAccount", "AcOurCode", "LCLTTBR", "S", "AcCode", "N")
               x1 = "Select * from tblVoucher where VType = 'JT' and  VFirm  = '" & gCCode & "' and VYear=" & gCYear
               GProcRstOpen RsTmp, x1, "O"
               X = "select Vfirm,BrokerCode,sum(amt1),sum(amt3) from TmpGentbl where UserName = '" & gUserName & "' and vfirm = '" & gCCode & "' group by Vfirm,BrokerCode"
               k = Month(dtpVouDt)
               wJvNo = GProcGenerateIdMonthwise("TblVoucher", "Vno", k, "Vdt", "Vtype='JT' and VFirm = '" & gCCode & "'   and VYear=" & gCYear)
               i = 1
               If GProcRstOpen(Rs, X, "R", gCn) > 0 Then
                  Rs.MoveFirst
                  Do While Not Rs.EOF
                     If Rs.Fields(3) > 0 Then
                        wParty = GProcGetColumnValue("TblMastAccount", "AcOurCode", Rs.Fields(1), "S", "AcCode", "N")
                     
                        If (Rs.Fields(2) < Val(txt(0))) And (Rs.Fields(2) + Rs.Fields(3) >= Val(txt(0))) Then
                           wTdsOn = Rs.Fields(2) + Rs.Fields(3)
                        ElseIf Rs.Fields(2) >= Val(txt(0)) Then
                           wTdsOn = Rs.Fields(3)
                        Else
                           wTdsOn = 0
                        End If
                        wTds = (wTdsOn * Val(txt(1))) / 100
                        DoEvents
                        DoEvents
                        DoEvents
                        DoEvents
                        wTds = GProcMakeRounding(wTds)
                        DoEvents
                        DoEvents
                        DoEvents
                        RsTmp.AddNew
                        RsTmp!Vno = wJvNo
                        RsTmp!Vtype = "JT"
                        RsTmp!Vdt = dtpVouDt
                        RsTmp!VYear = gCYear
                        RsTmp!VFirm = gCCode
                        RsTmp!VCtrNo = i
                        RsTmp!VCrAcCode = wParty
                        RsTmp!VDrAcCode = 15506
                        RsTmp!vamt = Rs.Fields(3)
                        RsTmp!VNar1 = "As Per Bill"
                        RsTmp!VNar2 = "From " + "" & dtpFrDt & "" + " To " + "" & dtpToDt & ""
                        RsTmp!VTopCrDr = 1
                        RsTmp!SubGroupCd = 1
                        RsTmp!VNoListVYear = gCYear
                        RsTmp!LessTDS = wTds
                        RsTmp!TdsRt = Val(txt(1))
                        RsTmp!TDSOn = wTdsOn
                        RsTmp!VCrAcCodeTDS = TdsCrCd
                        RsTmp!VDrAcCodeTDS = wParty
                        RsTmp!DbNtMonthly = 0  '--Tds Nature
                        i = i + 1
                        RsTmp.Update
                     End If
                     Rs.MoveNext
                  Loop
               End If
               MsgBox ("Brokerage Voucher Generated...")
               Exit Sub
            Else
               .ReportFileName = gReportPath & "rptBrokerageTotal.rpt"
            End If
         End If
         If gSelectedMenu = "Countwise Brokarage" Then .ReportFileName = gReportPath & "rptBrokerageCount.rpt"
    Case "Trade Purchase TDS Report", "Other Purchase TDS Report"
          If txtParty <> "" Then
             wPartyCode = GProcGetColumnValue("tblMastAccount", "AcName", txtParty, "S", "AcCode", "N")
          Else
             wPartyCode = 0
          End If
          If txtBroker <> "" Then
             wBrkCode = GProcGetColumnValue("tblMastAccount", "AcName", txtBroker, "S", "AcCode", "N")
          Else
             wBrkCode = 0
          End If
          If txtMill <> "" Then
             wMillCode = GProcGetColumnValue("tblMastAccount", "AcName", txtMill, "S", "AcCode", "N")
          Else
             wMillCode = 0
          End If
          gCmd.CommandText = "PrcPreparePurchaseGST"
          gCmd.Parameters.Refresh
          gCmd.Parameters("@Fr_dt") = Format(dtpFrDt, "yyyy/mm/dd")
          gCmd.Parameters("@To_dt") = Format(dtpToDt, "yyyy/mm/dd")
          gCmd.Parameters("@UCode") = gUserName
          If ChkConsolidate.Value = 1 Then
             gCmd.Parameters("@VFirm") = ""
          Else
             gCmd.Parameters("@VFirm") = Trim(gCCode)
          End If
          If gSelectedMenu = "Other Purchase TDS Report" Then
             gCmd.Parameters("@Vtype") = "PO"
          Else
             gCmd.Parameters("@Vtype") = "PY"
          End If
          If wPartyCode <> 0 Then gCmd.Parameters("@Accode") = wPartyCode
          If wBrkCode <> 0 Then gCmd.Parameters("@Accode_Br") = wBrkCode
          If wMillCode <> 0 Then gCmd.Parameters("@MillCode") = wMillCode
          gCmd.Execute
          If cmdCrystal(2).Value = True Then
             If txtParty = "" Then
                MsgBox "Select Party Name"
                Exit Sub
             End If
             PtyMailId = GProcGetColumnValue("tblMastAccount", "AcName", "'" & txtParty & "'", "N", "AcEmail", "S")
             If PtyMailId = "" Or IsNull(PtyMailId) Then
                MsgBox "Party Mail ID Not Found"
                Exit Sub
             End If
             Set objCrystal = New CRAXDRT.Application
             Tattach = gReportPath & "rptPurchTradePartywise.rpt"
             Set objReport = objCrystal.OpenReport(Tattach, 1)
             objReport.RecordSelectionFormula = "{TmpGentbl.UserName}='" & gUserName & "' and {TmpGentbl.Nar6} = 'A' and {TmpGentbl.Amt9} = 1 "
             objReport.Database.LogOnServer "pdsodbc.dll", "Hitrix", "Vishnu", "sa", "123456"
                    
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
                    Case "{@RepoName}"
                        objReport.FormulaFields(j).text = "Trim(" & Chr(39) & "Partywise Trade Purchase" & Chr(39) & ")"
                    Case "{@wFrToDt}"
                        objReport.FormulaFields(j).text = Chr(39) & "From : " + "" & Format(dtpFrDt.Value, "dd/MM/yyyy") & "" + " To : " + "" & Format(dtpToDt.Value, "dd/MM/yyyy") & "" & Chr(39)
                 End Select
             Next j
             Mfile = CStr(App.Path) + "\MailPDF\" + gCCode + "-" + gUserName + "-" + "Purchase_TDS"
             ExportReportToPDF objReport, Mfile & ".pdf", "foo"
             Mfile = CStr(App.Path) + "\MailPDF\" + gCCode + "-" + gUserName + "-" + "Purchase_TDS.pdf"
                   
             MailSub = "TDS On Purchase"
             MailStr = "Dear Sir, " + Chr(13) + Chr(10)

             MailStr = MailStr + "Please find attached TDS On Purchase" + Chr(13) + Chr(10)
             MailStr = MailStr + "Note : This is system generated email, Please do not reply."
             gMailsendToMill = False
             Call SendEmail1("" & Trim(gCMailId) & "", "" & PtyMailId & "", MailSub, MailStr, "" & PtyMailIdCc & "", "" & PtyMailIdBcc & "", Mfile)
             MsgBox "Mail Sent"
             Exit Sub
          ElseIf cmdCrystal(3).Value = True Then
             GProcTDSReportInExcle dtpFrDt.Value, dtpToDt.Value, "Partywise Trade Purchase"
             Exit Sub
          Else
            x1 = "{TmpGentbl.UserName}='" & gUserName & "' and {TmpGentbl.Nar6} = 'A' and {TmpGentbl.Amt9} = 1 "
            .SelectionFormula = x1
            .Formulas(4) = "wFrToDt = 'From : ' & '" & Format(dtpFrDt.Value, "dd/MM/yyyy") & "' & '  To ' & '" & Format(dtpToDt, "dd/MM/yyyy") & "'"
            If gSelectedMenu = "Other Purchase TDS Report" Then
                If optBrokList(1).Value = True Then
                   .Formulas(5) = "RepoName = 'Partywise Other Purchase'  "
                   .ReportFileName = gReportPath & "rptPurchTradePartywise.rpt"
                ElseIf optBrokList(0).Value = True Then
                   .Formulas(5) = "RepoName = 'Datewise Other Purchase'  "
                   .ReportFileName = gReportPath & "rptPurchTradeDatewiseSumm.rpt"
                Else
                   .Formulas(5) = "RepoName = 'Partywise Other Purchase Summery'  "
                   .ReportFileName = gReportPath & "rptPurchTradePartywiseSumm.rpt"
                End If
            Else
                If optBrokList(1).Value = True Then
                   .Formulas(5) = "RepoName = 'Partywise Trade Purchase'  "
                   .ReportFileName = gReportPath & "rptPurchTradePartywise.rpt"
                ElseIf optBrokList(0).Value = True Then
                   .Formulas(5) = "RepoName = 'Datewise Trade Purchase'  "
                   .ReportFileName = gReportPath & "rptPurchTradeDatewiseSumm.rpt"
                Else
                   .Formulas(5) = "RepoName = 'Partywise Trade Purchase Summery'  "
                   .ReportFileName = gReportPath & "rptPurchTradePartywiseSumm.rpt"
                End If
            End If
         End If
    Case "Depot Sales TDS Report", "Mill Bill TDS Report"
          If txtParty <> "" Then
             wPartyCode = GProcGetColumnValue("tblMastAccount", "AcName", txtParty, "S", "AcCode", "N")
          Else
             wPartyCode = 0
          End If
          If txtBroker <> "" Then
             wBrkCode = GProcGetColumnValue("tblMastAccount", "AcName", txtBroker, "S", "AcCode", "N")
          Else
             wBrkCode = 0
          End If
          If txtMill <> "" Then
             wMillCode = GProcGetColumnValue("tblMastAccount", "AcName", txtMill, "S", "AcCode", "N")
          Else
             wMillCode = 0
          End If
          gCmd.CommandText = "PrcPrepareSaleRegisterGST"
          gCmd.Parameters.Refresh
          gCmd.Parameters("@Fr_dt") = Format(dtpFrDt, "yyyy/mm/dd")
          gCmd.Parameters("@To_dt") = Format(dtpToDt, "yyyy/mm/dd")
          gCmd.Parameters("@UCode") = gUserName
          If chkIsExcludingSale.Value = 1 Then gCmd.Parameters("@Is2ndDay") = 1
          
          If chkGenVoucher.Value = 1 Then gCmd.Parameters("@IsRegiDay") = 1
          
          If ChkConsolidate.Value = 1 Then
             gCmd.Parameters("@VFirm") = ""
          Else
             gCmd.Parameters("@VFirm") = Trim(gCCode)
          End If
          If wPartyCode <> 0 Then gCmd.Parameters("@Accode") = wPartyCode
          If wMillCode <> 0 Then gCmd.Parameters("@MillCode") = wMillCode
          If gSelectedMenu = "Mill Bill TDS Report" Then
             gCmd.Parameters("@Vtype") = "SM"
          Else
             gCmd.Parameters("@Vtype") = "SD"
          End If
          gCmd.Execute
          
          If cmdCrystal(2).Value = True Then
             If txtMill = "" Then
                MsgBox "Select Mill Name"
                Exit Sub
             End If
             PtyMailId = GProcGetColumnValue("tblMastAccount", "AcName", "'" & txtMill & "'", "N", "AcEmail", "S")
             If PtyMailId = "" Or IsNull(PtyMailId) Then
                MsgBox "Mill Mail ID Not Found"
                Exit Sub
             End If
             Set objCrystal = New CRAXDRT.Application
             Tattach = gReportPath & "rptDepoMillBillSalePartywise.rpt"
             If gSelectedMenu = "Mill Bill TDS Report" And optBrokList(1).Value = True Then
                  gCn.BeginTrans
                  gCn.Execute "delete from TmpGentbl  where UserName = '" & gUserName & "' and amt15 = 1  "
                  gCn.CommitTrans
             End If
             Set objReport = objCrystal.OpenReport(Tattach, 1)
             objReport.RecordSelectionFormula = "{TmpGentbl.UserName}='" & gUserName & "' and {TmpGentbl.Nar6} = 'A' and {TmpGentbl.Amt14} = 1 "
             objReport.Database.LogOnServer "pdsodbc.dll", "Hitrix", "Vishnu", "sa", "123456"
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
                    Case "{@RepoName}"
                        objReport.FormulaFields(j).text = "Trim(" & Chr(39) & "Partywise Trade Purchase" & Chr(39) & ")"
                        If gSelectedMenu = "Mill Bill TDS Report" Then
                           If chkIsExcludingSale.Value = 1 Then
                              objReport.FormulaFields(j).text = "Trim(" & Chr(39) & "Mill Bill (Direct Sales) Report (For 2nd Day Payment)" & Chr(39) & ")"
                           ElseIf chkIsExcludingSale.Value = 1 Then
                              objReport.FormulaFields(j).text = "Trim(" & Chr(39) & "Mill Bill (Direct Sales) Report (Without 2nd Day Payment)" & Chr(39) & ")"
                           Else
                              .Formulas(5) = "RepoName = 'Mill Bill (Direct Sales)'  "
                              objReport.FormulaFields(j).text = "Trim(" & Chr(39) & "Mill Bill (Direct Sales)" & Chr(39) & ")"
                           End If
                        Else
                           If chkIsExcludingSale.Value = 1 Then
                              objReport.FormulaFields(j).text = "Trim(" & Chr(39) & "Depot Sales Report (For 2nd Day Payment)" & Chr(39) & ")"
                           ElseIf chkIsExcludingSale.Value = 1 Then
                              objReport.FormulaFields(j).text = "Trim(" & Chr(39) & "Depot Sales Report (For Without 2nd Day Payment)" & Chr(39) & ")"
                           Else
                              objReport.FormulaFields(j).text = "Trim(" & Chr(39) & "Depot Sales Report" & Chr(39) & ")"
                           End If
                        End If
                    Case "{@wFrToDt}"
                        objReport.FormulaFields(j).text = Chr(39) & "From : " + "" & Format(dtpFrDt.Value, "dd/MM/yyyy") & "" + " To : " + "" & Format(dtpToDt.Value, "dd/MM/yyyy") & "" & Chr(39)
                    Case "{@wFrToDt}"
                 End Select
             Next j
             Mfile = CStr(App.Path) + "\MailPDF\" + gCCode + "-" + gUserName + "-" + "Purchase_TDS"
             ExportReportToPDF objReport, Mfile & ".pdf", "foo"
             Mfile = CStr(App.Path) + "\MailPDF\" + gCCode + "-" + gUserName + "-" + "Purchase_TDS.pdf"
                   
             MailSub = "TDS Details"
             MailStr = "Dear Sir, " + Chr(13) + Chr(10)

             MailStr = MailStr + "Please find attached TDS Details" + Chr(13) + Chr(10)
             MailStr = MailStr + "Note : This is system generated email, Please do not reply."
             gMailsendToMill = False
             Call SendEmail1("" & Trim(gCMailId) & "", "" & PtyMailId & "", MailSub, MailStr, "" & PtyMailIdCc & "", "" & PtyMailIdBcc & "", Mfile)
             MsgBox "Mail Sent"
             Exit Sub
          ElseIf cmdCrystal(3).Value = True Then
             If gSelectedMenu = "Mill Bill TDS Report" Then
                gCn.BeginTrans
                gCn.Execute "delete from TmpGentbl  where UserName = '" & gUserName & "' and amt15 = 1  "
                gCn.CommitTrans
                GProcTDSReportInExcleMillDepot dtpFrDt.Value, dtpToDt.Value, "Mill Bill TDS Report"
             Else
                GProcTDSReportInExcleMillDepot dtpFrDt.Value, dtpToDt.Value, "Depot Sales TDS Report"
             End If
             Exit Sub
          Else
            x1 = "{TmpGentbl.UserName}='" & gUserName & "' and {TmpGentbl.Nar6} = 'A' and {TmpGentbl.Amt14} = 1 "
            .SelectionFormula = x1
            .Formulas(4) = "wFrToDt = 'From : ' & '" & Format(dtpFrDt.Value, "dd/MM/yyyy") & "' & '  To ' & '" & Format(dtpToDt, "dd/MM/yyyy") & "'"
            .Formulas(6) = "wCPan = '" & gCPAN & "' "
            If optBrokList(1).Value = True Then
               If gSelectedMenu = "Mill Bill TDS Report" Then
                  gCn.BeginTrans
                  gCn.Execute "delete from TmpGentbl  where UserName = '" & gUserName & "' and amt15 = 1  "
                  gCn.CommitTrans
                  If chkIsExcludingSale.Value = 1 Then
                     .Formulas(5) = "RepoName = 'Mill Bill (Direct Sales) Report (For 2nd Day Payment)'  "
                  ElseIf chkIsExcludingSale.Value = 1 Then
                     .Formulas(5) = "RepoName = 'Mill Bill (Direct Sales) Report (Without 2nd Day Payment)'  "
                  Else
                     .Formulas(5) = "RepoName = 'Mill Bill (Direct Sales)'  "
                  End If
               Else
                  If chkIsExcludingSale.Value = 1 Then
                     .Formulas(5) = "RepoName = 'Depot Sales Report (For 2nd Day Payment)'  "
                  ElseIf chkIsExcludingSale.Value = 1 Then
                     .Formulas(5) = "RepoName = 'Depot Sales Report (For Without 2nd Day Payment)'  "
                  Else
                     .Formulas(5) = "RepoName = 'Depot Sales Report'  "
                  End If
               End If
               .ReportFileName = gReportPath & "rptDepoMillBillSalePartywise.rpt"
            Else
               If gSelectedMenu = "Mill Bill TDS Report" Then
                  gCn.BeginTrans
                  gCn.Execute "delete from TmpGentbl  where UserName = '" & gUserName & "' and amt15 = 1  "
                  gCn.CommitTrans
                  If chkIsExcludingSale.Value = 1 Then
                     .Formulas(5) = "RepoName = 'Mill Bill (Direct Sales) TDS Summery (For 2nd Day Payment)'  "
                  ElseIf chkIsExcludingSale.Value = 1 Then
                     .Formulas(5) = "RepoName = 'Mill Bill (Direct Sales) TDS Summery (For Without 2nd Day Payment)'  "
                  Else
                     .Formulas(5) = "RepoName = 'Mill Bill (Direct Sales) TDS Summery'  "
                  End If
               Else
                  If chkIsExcludingSale.Value = 1 Then
                     .Formulas(5) = "RepoName = 'Depot Sales TDS Summery (For 2nd Day Payment)'  "
                  ElseIf chkIsExcludingSale.Value = 1 Then
                     .Formulas(5) = "RepoName = 'Depot Sales TDS Summery (For Without 2nd Day Payment)'  "
                  Else
                     .Formulas(5) = "RepoName = 'Depot Sales TDS Summery'  "
                  End If
               End If
               If optBrokList(0).Value = True Then
                  .ReportFileName = gReportPath & "rptDepoMillBillSaleDaywise.rpt"
               Else
                  .ReportFileName = gReportPath & "rptDepoMillBillSaleSumm.rpt"
               End If
            End If
        End If
    Case "Millwise Inward"
          gCmd.CommandText = "PrcPrepareDailyInward"
          gCmd.Parameters.Refresh
          gCmd.Parameters("@Fr_dt") = Format(dtpFrDt, "yyyy/mm/dd")
          gCmd.Parameters("@To_dt") = Format(dtpToDt, "yyyy/mm/dd")
          gCmd.Parameters("@UCode") = gUserName
          If ChkConsolidate.Value = 1 Then
             gCmd.Parameters("@VFirm") = ""
          Else
             gCmd.Parameters("@VFirm") = Trim(gCCode)
          End If
          If wMillCode <> 0 Then gCmd.Parameters("@MillCode") = wMillCode
          gCmd.Execute
          x1 = "{TmpGentbl.UserName}='" & gUserName & "'"
          .SelectionFormula = x1
          .Formulas(4) = "wFrToDt = 'From : ' & '" & Format(dtpFrDt.Value, "dd/MM/yyyy") & "' & '  To ' & '" & Format(dtpToDt, "dd/MM/yyyy") & "'"
          .ReportFileName = gReportPath & "rptDailyInward.rpt"
    Case "Daily Purchase Report ( Mill Bill )"
          If txtMill <> "" Then
             wMillCode = GProcGetColumnValue("tblMastAccount", "AcName", txtMill, "S", "AcCode", "N")
          Else
             wMillCode = 0
          End If
          gCmd.CommandText = "PrcPrepareDailyEntryMillBill"
          gCmd.Parameters.Refresh
          gCmd.Parameters("@Fr_dt") = Format(dtpFrDt, "yyyy/mm/dd")
          gCmd.Parameters("@To_dt") = Format(dtpToDt, "yyyy/mm/dd")
          gCmd.Parameters("@UCode") = gUserName
          If ChkConsolidate.Value = 1 Then
             gCmd.Parameters("@VFirm") = ""
          Else
             gCmd.Parameters("@VFirm") = Trim(gCCode)
          End If
          gCmd.Parameters("@Vtype") = "SM"
          gCmd.Parameters("@Vyear") = gCYear
          If wMillCode <> 0 Then gCmd.Parameters("@MillCode") = wMillCode
          gCmd.Execute
          x1 = "{TmpGentbl.UserName}='" & gUserName & "' "
          .SelectionFormula = x1
          .Formulas(4) = "wFrToDt = 'From : ' & '" & Format(dtpFrDt.Value, "dd/MM/yyyy") & "' & '  To ' & '" & Format(dtpToDt, "dd/MM/yyyy") & "'"
          .ReportFileName = gReportPath & "rptDailyEntryMillBill.rpt"
    Case "Daily Purchase Report"
          If txtMill <> "" Then
             wMillCode = GProcGetColumnValue("tblMastAccount", "AcName", txtMill, "S", "AcCode", "N")
          Else
             wMillCode = 0
          End If
          gCmd.CommandText = "PrcPreparePurchaseDetail"
          gCmd.Parameters.Refresh
          gCmd.Parameters("@Fr_dt") = Format(dtpFrDt, "yyyy/mm/dd")
          gCmd.Parameters("@To_dt") = Format(dtpToDt, "yyyy/mm/dd")
          gCmd.Parameters("@UCode") = gUserName
          If ChkConsolidate.Value = 1 Then
             gCmd.Parameters("@VFirm") = ""
          Else
             gCmd.Parameters("@VFirm") = Trim(gCCode)
          End If
          gCmd.Parameters("@Vtype") = "PY"
          If wMillCode <> 0 Then gCmd.Parameters("@MillCode") = wMillCode
          gCmd.Execute
          x1 = "{TmpGentbl.UserName}='" & gUserName & "' and {TmpGentbl.Nar6} = 'A'"
          .SelectionFormula = x1
          .Formulas(4) = "wFrToDt = 'From : ' & '" & Format(dtpFrDt.Value, "dd/MM/yyyy") & "' & '  To ' & '" & Format(dtpToDt, "dd/MM/yyyy") & "'"
          .ReportFileName = gReportPath & "rptRegisterPurchDaily.rpt"
    End Select
    .Action = 1
    .PageZoom (120)
End With
Screen.MousePointer = vbDefault
End Sub






'            If ChkConsolidate = 1 Then
'               If rstTmp.Fields("CDepotMainFirmCompCd") = "'" & rstTmp.Fields("CCode") & "'" Then
'                   If chkIsExcludingSale And gSelectedMenu = "Brokerage Voucher" Then
'                     gCn.BeginTrans
'                     gCn.Execute "update tblSale set SlTmpNo3 = 1 where vfirm = '" & rstTmp.Fields("CCode") & "' and vyear = " & gCYear & " and vdt < '" & Format(BPaidDate, "yyyy/mm/dd") & "'  and  SlBroker in (select VCrAcCode from tblVoucher where vtype = 'JT' and vfirm = '" & rstTmp.Fields("CDepotMainFirmCompCd") & "' and vyear = " & gCYear & " and vdt < '" & Format(gCYEDate, "YYYY/MM/DD") & "'  and VDrAcCode = 15508) "
'                     gCn.Execute "update tblSale set SlPrevBrokPd = 'Y' where vfirm = '" & rstTmp.Fields("CCode") & "' and vyear = " & gCYear & "  and  SlBroker in (select VCrAcCode from tblVoucher where vtype = 'JT' and vfirm = '" & rstTmp.Fields("CDepotMainFirmCompCd") & "' and vyear = " & gCYear & " and vdt < '" & Format(gCYEDate, "YYYY/MM/DD") & "'  and VDrAcCode = 15508) "
'                     gCn.CommitTrans
'                   End If
'               Else
'                   If chkIsExcludingSale And gSelectedMenu = "Brokerage Voucher" Then
'                     gCn.BeginTrans
'                     gCn.Execute "update tblSale set SlTmpNo3 = 1 where vfirm = '" & rstTmp.Fields("CCode") & "' and vyear = " & gCYear & " and vdt < '" & Format(BPaidDate, "yyyy/mm/dd") & "'  and  SlBroker in (select VCrAcCode from tblVoucher where vtype = 'JT' and vfirm = '" & rstTmp.Fields("CCode") & "' and vyear = " & gCYear & " and vdt < '" & Format(gCYEDate, "YYYY/MM/DD") & "'  and VDrAcCode = 15508) "
'                     gCn.Execute "update tblSale set SlPrevBrokPd = 'Y' where vfirm = '" & rstTmp.Fields("CCode") & "' and vyear = " & gCYear & "  and  SlBroker in (select VCrAcCode from tblVoucher where vtype = 'JT' and vfirm = '" & rstTmp.Fields("CCode") & "' and vyear = " & gCYear & " and vdt < '" & Format(gCYEDate, "YYYY/MM/DD") & "'  and VDrAcCode = 15508) "
'                     gCn.CommitTrans
'                   End If
'               End If
'               gCmd.CommandText = "PrcPrepareBrokSale"
'               gCmd.Parameters.Refresh
'               gCmd.Parameters("@Start_dt") = IIf(gSelectedMenu = "Countwise Brokarage" Or gSelectedMenu = "Brokerage Summery", Format(dtpFrDt, "yyyy/mm/dd"), Format(gCYSDate, "yyyy/mm/dd"))
'               gCmd.Parameters("@From_dt") = Format(dtpFrDt, "yyyy/mm/dd")
'               gCmd.Parameters("@To_dt") = Format(dtpToDt, "yyyy/mm/dd")
'               gCmd.Parameters("@TDSLimit") = Val(txt(0))
'               gCmd.Parameters("@TDSRate") = Val(txt(1))
'               gCmd.Parameters("@UCode") = gUserName
'               gCmd.Parameters("@VFirm") = rstTmp.Fields("CCode")
'               If gSelectedMenu = "Brokerage Detail " Then gCmd.Parameters("@BrkTp") = "D"
'               If gSelectedMenu = "Brokerage Summery" Then gCmd.Parameters("@BrkTp") = "V"
'               If gSelectedMenu = "Brokerage Voucher" Then gCmd.Parameters("@BrkTp") = "V"
'               If gSelectedMenu = "Countwise Brokarage" Then gCmd.Parameters("@BrkTp") = "V"
'               If ChkConsolidate.Value = 1 Then
'                  gCmd.Parameters("@VFirmAdd1") = ""
'                  gCmd.Parameters("@VFirmAdd2") = ""
'                  gCmd.Parameters("@VFirmAdd3") = ""
'                  gCmd.Parameters("@VFirmCph") = ""
'                  gCmd.Parameters("@VFirmGstin") = ""
'               Else
'                  gCmd.Parameters("@VFirmAdd1") = Trim(gCAdd1)
'                  gCmd.Parameters("@VFirmAdd2") = Trim(gCAdd2)
'                  gCmd.Parameters("@VFirmAdd3") = Trim(gCAdd3)
'                  gCmd.Parameters("@VFirmCph") = Trim(gCPhNo)
'                  gCmd.Parameters("@VFirmGstin") = Trim(gCGStin)
'               End If
'               If rstTmp.Fields("CDepotMainFirmCompCd") <> 0 Then
'                  gCmd.Parameters("@DepotFirm") = "D"
'               Else
'                  gCmd.Parameters("@DepotFirm") = "F"
'               End If
'               If wAcCode <> 0 Then gCmd.Parameters("@Accode") = wAcCode
'               gCmd.Execute
'            End If





