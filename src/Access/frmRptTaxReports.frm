VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form frmRptTaxReports 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Tax Reports"
   ClientHeight    =   5760
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10425
   Icon            =   "frmRptTaxReports.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5760
   ScaleWidth      =   10425
   Begin VB.CheckBox ChkSumTin 
      BackColor       =   &H80000018&
      Caption         =   "RCM Purch"
      ForeColor       =   &H00000080&
      Height          =   375
      Left            =   9045
      TabIndex        =   52
      Top             =   2010
      Width           =   1245
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
      Left            =   7320
      TabIndex        =   30
      Text            =   "5"
      Top             =   2460
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
      Left            =   9255
      TabIndex        =   46
      Text            =   "4"
      Top             =   4935
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
      Left            =   7065
      TabIndex        =   44
      Text            =   "3"
      Top             =   4935
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
      Left            =   9255
      TabIndex        =   42
      Text            =   "2"
      Top             =   4440
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
      Index           =   1
      Left            =   7065
      TabIndex        =   40
      Text            =   "1"
      Top             =   4470
      Width           =   1140
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
      Left            =   5985
      TabIndex        =   32
      Top             =   3345
      Width           =   3450
      Begin VB.OptionButton OptBrkSAType 
         BackColor       =   &H80000018&
         Caption         =   "Mill Bill"
         ForeColor       =   &H00000080&
         Height          =   360
         Index           =   5
         Left            =   2205
         TabIndex        =   38
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
         TabIndex        =   35
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
         TabIndex        =   37
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
         TabIndex        =   36
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
         TabIndex        =   34
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
         TabIndex        =   33
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
      Left            =   7335
      TabIndex        =   28
      Top             =   2010
      Width           =   1740
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
      Left            =   255
      TabIndex        =   15
      Top             =   3585
      Width           =   3480
      Begin VB.OptionButton OptBag 
         BackColor       =   &H80000018&
         Caption         =   "Bags"
         ForeColor       =   &H00000080&
         Height          =   360
         Left            =   2520
         TabIndex        =   18
         Top             =   300
         Width           =   930
      End
      Begin VB.OptionButton OptGrossAmt 
         BackColor       =   &H80000018&
         Caption         =   "&Gross Amt"
         ForeColor       =   &H00000080&
         Height          =   360
         Left            =   165
         TabIndex        =   16
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
         TabIndex        =   17
         Top             =   300
         Width           =   1020
      End
   End
   Begin VB.TextBox txtItCode 
      Height          =   375
      Left            =   6465
      TabIndex        =   51
      Top             =   5325
      Visible         =   0   'False
      Width           =   570
   End
   Begin VB.TextBox txtItem 
      Height          =   375
      Left            =   1680
      TabIndex        =   11
      Top             =   2175
      Width           =   3960
   End
   Begin VB.CheckBox ChkSummary 
      BackColor       =   &H80000018&
      Caption         =   "Is Summary ?"
      ForeColor       =   &H00000080&
      Height          =   330
      Left            =   5985
      TabIndex        =   27
      Top             =   2010
      Width           =   1830
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H80000018&
      Height          =   1245
      Left            =   7740
      TabIndex        =   23
      Top             =   645
      Width           =   1710
      Begin VB.OptionButton Opt 
         BackColor       =   &H80000018&
         Caption         =   "Pe&nding"
         ForeColor       =   &H00000080&
         Height          =   360
         Index           =   1
         Left            =   105
         TabIndex        =   25
         Top             =   495
         Value           =   -1  'True
         Width           =   1005
      End
      Begin VB.OptionButton Opt 
         BackColor       =   &H80000018&
         Caption         =   "&Received"
         ForeColor       =   &H00000080&
         Height          =   360
         Index           =   2
         Left            =   105
         TabIndex        =   26
         Top             =   840
         Width           =   1215
      End
      Begin VB.OptionButton Opt 
         BackColor       =   &H80000018&
         Caption         =   "&All"
         ForeColor       =   &H00000080&
         Height          =   360
         Index           =   0
         Left            =   105
         TabIndex        =   24
         Top             =   150
         Width           =   1215
      End
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
      Index           =   0
      Left            =   1680
      TabIndex        =   13
      Text            =   "0"
      Top             =   2625
      Width           =   1380
   End
   Begin VB.TextBox txtParty 
      Height          =   375
      Left            =   1680
      TabIndex        =   9
      Top             =   1725
      Width           =   3960
   End
   Begin VB.CheckBox ChkConsolidate 
      BackColor       =   &H80000018&
      Caption         =   "Is &Consolidate Report ?"
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
      Left            =   3270
      TabIndex        =   14
      Top             =   2625
      Width           =   2415
   End
   Begin VB.CommandButton cmdCrystal 
      Height          =   480
      Index           =   1
      Left            =   2475
      Picture         =   "frmRptTaxReports.frx":058A
      Style           =   1  'Graphical
      TabIndex        =   48
      ToolTipText     =   "Print"
      Top             =   4695
      Width           =   1185
   End
   Begin VB.CommandButton cmdCrystal 
      Height          =   480
      Index           =   0
      Left            =   1020
      Picture         =   "frmRptTaxReports.frx":0BF4
      Style           =   1  'Graphical
      TabIndex        =   47
      Top             =   4695
      Width           =   1185
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
      Left            =   3945
      TabIndex        =   49
      Top             =   4695
      Width           =   1185
   End
   Begin VB.TextBox txtBroker 
      Height          =   375
      Left            =   1680
      TabIndex        =   7
      Top             =   1260
      Width           =   3960
   End
   Begin VB.TextBox txtMill 
      Height          =   375
      Left            =   1680
      TabIndex        =   5
      Top             =   780
      Width           =   3960
   End
   Begin Crystal.CrystalReport CrystalReport1 
      Left            =   210
      Top             =   5460
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
   Begin MSComctlLib.ProgressBar prgBar 
      Height          =   255
      Left            =   255
      TabIndex        =   50
      Top             =   3135
      Visible         =   0   'False
      Width           =   5370
      _ExtentX        =   9472
      _ExtentY        =   450
      _Version        =   393216
      Appearance      =   1
   End
   Begin MSComCtl2.DTPicker dtpToDt 
      Height          =   360
      Left            =   3990
      TabIndex        =   3
      Top             =   300
      Width           =   1650
      _ExtentX        =   2910
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
      Format          =   101449731
      CurrentDate     =   38050
   End
   Begin MSComCtl2.DTPicker dtpFrDt 
      Height          =   360
      Left            =   1680
      TabIndex        =   1
      Top             =   300
      Width           =   1575
      _ExtentX        =   2778
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
      Format          =   101449731
      CurrentDate     =   38050
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H80000018&
      Height          =   1260
      Left            =   5985
      TabIndex        =   20
      Top             =   630
      Width           =   1695
      Begin VB.OptionButton OptParty 
         BackColor       =   &H80000018&
         Caption         =   "Partywise"
         ForeColor       =   &H00004080&
         Height          =   360
         Left            =   150
         TabIndex        =   22
         Top             =   495
         Width           =   1335
      End
      Begin VB.OptionButton OptBrk 
         BackColor       =   &H80000018&
         Caption         =   "Brokerwise"
         ForeColor       =   &H00004080&
         Height          =   360
         Left            =   150
         TabIndex        =   21
         Top             =   180
         Value           =   -1  'True
         Width           =   1335
      End
   End
   Begin VB.Shape Shape2 
      FillColor       =   &H00BDDFD1&
      FillStyle       =   0  'Solid
      Height          =   765
      Left            =   765
      Shape           =   4  'Rounded Rectangle
      Top             =   4545
      Width           =   4590
   End
   Begin VB.Line Line1 
      BorderColor     =   &H000040C0&
      X1              =   5835
      X2              =   10440
      Y1              =   2910
      Y2              =   2910
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
      Left            =   6000
      TabIndex        =   29
      Top             =   2460
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
      Left            =   8355
      TabIndex        =   45
      Top             =   4935
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
      Left            =   5910
      TabIndex        =   43
      Top             =   4935
      Width           =   1125
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "TDS %"
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
      Left            =   8355
      TabIndex        =   41
      Top             =   4455
      Width           =   615
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "TDS Limit"
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
      Left            =   5925
      TabIndex        =   39
      Top             =   4440
      Width           =   855
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
      Left            =   5985
      TabIndex        =   31
      Top             =   2970
      Width           =   3375
   End
   Begin VB.Line Line3 
      BorderColor     =   &H000040C0&
      X1              =   5805
      X2              =   5805
      Y1              =   150
      Y2              =   5490
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
      Left            =   5985
      TabIndex        =   19
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
      TabIndex        =   10
      Top             =   2175
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
      TabIndex        =   8
      Top             =   1725
      Width           =   1275
   End
   Begin VB.Label lblVAT 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "&VAT OP. Bal  :"
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
      TabIndex        =   12
      Top             =   2610
      Width           =   1305
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
      TabIndex        =   6
      Top             =   1260
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
      TabIndex        =   4
      Top             =   780
      Width           =   1290
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "To :"
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
      Left            =   3420
      TabIndex        =   2
      Top             =   315
      Width           =   360
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
      Height          =   5370
      Left            =   105
      Top             =   120
      Width           =   10275
   End
End
Attribute VB_Name = "frmRptTaxReports"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim wMillCode As Long
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
dtpToDt.Value = Format(gCYEDate, "dd/MM/yyyy")
Me.Caption = gSelectedMenu
For i = 0 To 4
    txt(i) = "0.00"
Next i
txt(5) = "" '-F Form No
Select Case gSelectedMenu
    Case "For Eof() Bof() Error"
         dtpFrDt.Enabled = False
         dtpToDt.Enabled = False
         cmdCrystal(0).Caption = "Delete"
         cmdCrystal(1).Visible = False
         Label4.Caption = "Sale Give Bill No/Purchase Vou.No"
         txt(5).Enabled = True
         Label9.Caption = "Bill/Vou.No :"
         OptBrkSAType(0).Value = True
         OptBrkSAType(0).Enabled = True
         OptBrkSAType(0).Caption = "Sale"
         OptBrkSAType(1).Enabled = True
         OptBrkSAType(1).Caption = "Purch"

    '========== Tax Reports
    '---------Form Receivable,issued
    Case "C Form Receivable", "E2 From Issuable", "E2 Form Receivable" '-"C Form Receivable",
        txtBroker.Enabled = True
        txtMill.Enabled = True
        txtParty.Enabled = True
        ChkConsolidate.Enabled = True
        ChkSummary.Enabled = True
        ChkPartyOnNewPage.Enabled = True
        Opt(0).Enabled = True '--All
        Opt(1).Enabled = True '-- Pending
        Opt(2).Enabled = True '--Receive
        Opt(1) = True
        If gSelectedMenu <> "E2 Form Receivable" Then
           OptBrk.Enabled = True
           OptParty.Enabled = True
        End If
        If gSelectedMenu = "C Form Receivable" Then
           OptGrossAmt.Value = False
           OptGrossAmt.Enabled = True
           OptGrossAmt.Caption = "New Report"
        End If
        OptBrk = True
        ChkConsolidate = 1
        If gSelectedMenu = "E2 From Issuable" Then
             Opt(2).Caption = "Issued"
        End If
    Case "C Form Issued But E1Form Not Received"
            txtBroker.Enabled = True
            ChkConsolidate.Enabled = True
    Case "C Form Issuable", "E1 Form Receivable", "F Form Issuable" '-- "C Form Issuable", "E1 Form Receivable",F Form Issuable
        txtParty.Enabled = True
        txtMill.Enabled = True
        txtBroker.Enabled = True
        ChkPartyOnNewPage.Enabled = True
        ChkConsolidate.Enabled = True
        ChkSummary.Enabled = True
        If gSelectedMenu = "E1 Form Receivable" Then
           Opt(2).Caption = "Received"
        End If
        If gSelectedMenu = "C Form Issuable" Or gSelectedMenu = "E1 Form Receivable" Then
           OptBrk.Value = True
           OptParty.Enabled = True
           OptParty.Caption = "Print Covering Letter"
           OptBrk.Enabled = True
           OptBrk.Caption = "New Report"
        End If
        If gSelectedMenu <> "E1 Form Receivable" And gSelectedMenu <> "E2 Form Receivable" Then
           Opt(2).Caption = "Issued"
           OptGrossAmt.Value = False
           OptGrossAmt.Enabled = True
           OptGrossAmt.Caption = "Partywise Summary"
        End If
        Opt(0).Enabled = True '--All
        Opt(1).Enabled = True  '-- Pending
        Opt(2).Enabled = True  '--Receive
        Opt(1) = True
        ChkConsolidate = 1
    Case "F Form Statement" '-- "F Form Statement"
        txtMill.Enabled = True '--Mill
        txt(5).Enabled = True '- 'F' From No
    Case "Requirement Of C Form", "Requirment of C Form Invoicewise", "Requirment of E2 Form Invoicewise"  '--  "Requirement Of C Form"
        ChkConsolidate.Enabled = True
        ChkConsolidate = 1
        txtParty.Enabled = True
        Opt(0) = True '--All
    Case "Requirement Of F Form", "Requirment of F Form Invoicewise" '-- "Requirement Of F Form"
        ChkConsolidate.Enabled = True
        ChkConsolidate = 1
    Case "Pending C Form Receivable Slab Wise" '-- "Pending C Form Receivable Slab Wise"
        dtpFrDt.Enabled = False
        dtpToDt.Enabled = False
        txtBroker.Enabled = True
        txtParty.Enabled = True
        ChkConsolidate.Enabled = True
        ChkConsolidate = 1
    Case "Pending E1 Form Receivable Slab Wise" '-"Pending E1 Form Receivable Slab Wise"
        dtpFrDt.Enabled = False
        dtpToDt.Enabled = False
        txtMill.Enabled = True
        ChkConsolidate.Enabled = True
        ChkConsolidate = 1
    Case "V A T Computation" '-- "V A T Computation"
        txt(0).Enabled = True '-- OP bal
    Case "V A T Computation 2 And 5 %"
        txt(0).Enabled = True '-- OP bal
    Case "TDS Register", "TDS Summary" '- "TDS Register" ,summary
        txtParty.Enabled = True '-Party
        LblMill.Caption = "Interest A/c   :"
        lblBroker.Caption = "T.D.S. A/c     :"
        txtMill.Enabled = True
        txtBroker.Enabled = True
        ChkSummary.Enabled = True
    '========== General Report
    Case "Total Sales Summary" 'Total Sales
        ChkConsolidate.Enabled = True
        OptGrossAmt.Enabled = True
        OptNetAmt.Enabled = True
        OptBag.Enabled = True
        OptGrossAmt = True
    Case "Countwise Sales Summary" 'Countwise Sales
        ChkConsolidate.Enabled = True
        txtItem.Enabled = True
        OptGrossAmt.Enabled = True
        OptNetAmt.Enabled = True
        OptBag.Enabled = True
        OptGrossAmt = True
    Case "Millwise Sales Summary" 'Millwise Sales
        ChkConsolidate.Enabled = True
        txtMill.Enabled = True
        OptGrossAmt.Enabled = True
        OptNetAmt.Enabled = True
        OptBag.Enabled = True
        OptGrossAmt = True
    Case "Mill and Countwise Sales Summary" 'Mill and Countwise Sales
        ChkConsolidate.Enabled = True
        txtMill.Enabled = True
        txtItem.Enabled = True
        OptGrossAmt.Enabled = True
        OptNetAmt.Enabled = True
        OptBag.Enabled = True
        OptGrossAmt = True
    Case "Broker and Partywise Sales Summary" 'Broker and Partywise Sales
        ChkConsolidate.Enabled = True
        txtBroker.Enabled = True
        txtParty.Enabled = True
        OptGrossAmt.Enabled = True
        OptNetAmt.Enabled = True
        OptBag.Enabled = True
        OptGrossAmt = True
    Case "Broker and Countwise Sales Summary" 'Broker and Countwise Sales
        ChkConsolidate.Enabled = True
        txtBroker.Enabled = True
        txtItem.Enabled = True
        OptGrossAmt.Enabled = True
        OptNetAmt.Enabled = True
        OptBag.Enabled = True
        OptGrossAmt = True
    Case "Partywise Sales Summary" '-Partywise Sales
        ChkConsolidate.Enabled = True
        txtParty.Enabled = True
        OptGrossAmt.Enabled = True
        OptNetAmt.Enabled = True
        OptBag.Enabled = True
        OptGrossAmt = True
    Case "Party and Countwise Sales Summary" '-Party and Countwise Sales
        ChkConsolidate.Enabled = True
        txtParty.Enabled = True
        txtItem.Enabled = True
        OptGrossAmt.Enabled = True
        OptNetAmt.Enabled = True
        OptBag.Enabled = True
        OptGrossAmt = True
    Case "Monthwise Sales Summary", "Monthwise Purchase Summary", "Monthwise Sales/Purchase New" '--Monthwise Sales
        ChkConsolidate.Enabled = True
        OptGrossAmt.Enabled = True
        OptNetAmt.Enabled = True
        OptBag.Enabled = True
        OptGrossAmt = True
        If gSelectedMenu = "Monthwise Sales/Purchase New" Then
            OptGrossAmt.Enabled = False
            OptNetAmt.Enabled = False
            OptBag.Enabled = False
            OptBag.Value = True
            ChkSummary.Enabled = True
            ChkSummary.Caption = "Sales Report"
            ChkPartyOnNewPage.Enabled = True
            ChkPartyOnNewPage.Caption = "Purchase Report"
            ChkSumTin.Visible = True
            ChkSumTin.Enabled = True
            ChkSumTin.Caption = "RCM Purch"
       End If
        
     Case "Total Purchase Summary" '--"Total Purchase Summary"
        ChkConsolidate.Enabled = True
        OptGrossAmt.Enabled = True
        OptNetAmt.Enabled = True
        OptBag.Enabled = True
        OptGrossAmt = True
     Case "Countwise Purchase Summary" '--"Countwise Purchase Summary"
        ChkConsolidate.Enabled = True
        txtItem.Enabled = True
        OptGrossAmt.Enabled = True
        OptNetAmt.Enabled = True
        OptBag.Enabled = True
        OptGrossAmt = True
     Case "Millwise Purchase Summary" '-- "Millwise Purchase Summary"
        ChkConsolidate.Enabled = True
        txtMill.Enabled = True
        OptGrossAmt.Enabled = True
        OptNetAmt.Enabled = True
        OptBag.Enabled = True
        OptGrossAmt = True
     Case "Monthwise Sales/Purchase Summary" '--"Monthwise Sales/Purchase Summary"
        ChkConsolidate.Enabled = True
        OptGrossAmt.Enabled = True
        OptNetAmt.Enabled = True
        OptBag.Enabled = True
        OptGrossAmt = True
    Case "Total Inward Summary" '--"Total Inward Summary"
        ChkConsolidate.Enabled = True
        OptGrossAmt.Enabled = True
        OptNetAmt.Enabled = True
        OptBag.Enabled = True
        OptGrossAmt = True
     Case "Countwise Inward Summary" '--"Countwise Inward Summary"
        ChkConsolidate.Enabled = True
        txtItem.Enabled = True
        OptGrossAmt.Enabled = True
        OptNetAmt.Enabled = True
        OptBag.Enabled = True
        OptGrossAmt = True
     Case "Millwise Inward Summary" '-- "Millwise Inward Summary"
        ChkConsolidate.Enabled = True
        txtMill.Enabled = True
        OptGrossAmt.Enabled = True
        OptNetAmt.Enabled = True
        OptBag.Enabled = True
        OptGrossAmt = True
        '--- Brokerage Reports
     Case "Brokerage Detail Report" '-- "Brokerage Detail Report"
        txtBroker.Enabled = True
        ChkSummary.Enabled = True
        ChkPartyOnNewPage.Enabled = True
        ChkPartyOnNewPage.Caption = "Include GST"
        
        For i = 0 To 5
            OptBrkSAType(i).Enabled = True
        Next i
     Case "Brokerage Voucher Report", "Brokerage Total Report", "Inward/Purchase Brokerage Report", "Inward/Purchase Brokerage Detail Report", "Inward/Purchase Brokarage Total" '-- "Brokerage Voucher Report", "Brokerage Total Report"
        txtBroker.Enabled = True
        txtMill.Enabled = True
        txtItem.Enabled = True
        ChkPartyOnNewPage.Enabled = True
        ChkPartyOnNewPage.Caption = "Brokerage Bill"
        For i = 0 To 5
            OptBrkSAType(i).Enabled = True
        Next i
        For i = 1 To 4
            txt(i).Enabled = True
        Next i
       If gSelectedMenu = "Inward/Purchase Brokerage Detail Report" Then
           ChkSummary.Enabled = True
           Label7.Visible = False
           Label8.Visible = False
           txt(3).Visible = False
           txt(4).Visible = False
        End If
        txt(1) = 5000 '-TDS Limit
        txt(2) = 10 '-tds%
        txt(3) = 0 '-Surcharge %
        txt(4) = 0 '-Surcharge %
    Case "Transactionwise Sales Details", "Transactionwise Purchase Details"
        txtBroker.Enabled = False
        txtMill.Enabled = False
        txtParty.Enabled = False
        ChkConsolidate.Enabled = False
        ChkSummary.Enabled = False
        ChkPartyOnNewPage.Enabled = False
        
        OptGrossAmt.Enabled = False
        OptGrossAmt.Value = False
        Opt(0).Enabled = False '--All
        Opt(1).Enabled = False '-- Pending
        Opt(2).Enabled = False '--Receive
        Opt(1) = False
        OptBrk.Enabled = False
        OptParty.Enabled = False
        OptBrk = False
        ChkConsolidate = 0
        OptBrkSAType(0).Visible = False
        OptBrkSAType(1).Visible = False
        OptBrkSAType(2).Visible = False
        OptBrkSAType(3).Visible = False
        OptBrkSAType(4).Visible = False
        OptBrkSAType(5).Visible = False
        OptBrkSAType(5).Enabled = False
        OptBrkSAType(0).Enabled = False
    '========== General Report End
    Case "Partywise Sales VAT Audit"
        LblMill.Caption = "Sales A/c"
        txtMill.Visible = True
        txtMill.Enabled = True
    Case "Cash - Bank  Status"
        ChkConsolidate.Enabled = True
        ChkConsolidate = 1
        dtpFrDt.Enabled = False
        dtpToDt.Enabled = True
End Select
 '-- Set InActivate Ctrl Back Color
GProcSetInActivateCtrlBackClr Me
End Sub
Private Sub cmdExit_Click()
DoEvents
Unload Me
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
If gSelectedMenu = "Partywise Sales VAT Audit" Then
   gClsSearch.SearchMultiField "QryMastaccount", "Account,AcName,Accode,AcCity,AcMillTypeStr", Array("Sales Account", "AcName", "acCode", "City", "Mill Type"), Array(3000, 0, 0, 2000, 2000), " ((AgCode=90023 or GpCode=90023) or AcMillType=1 or AcMillType=2)", , "AcName"
ElseIf gSelectedMenu = "TDS Register" Then
   gClsSearch.SearchMultiField "QryMastaccount", "Account,AcName,Accode,AcCity,AcMillTypeStr", Array("Sales Account", "AcName", "acCode", "City", "Mill Type"), Array(3000, 0, 0, 2000, 2000), "", Chr(KeyAscii), "AcName"
Else
   gClsSearch.SearchMultiField "QryMastaccount", "Account,AcName,ACCode", Array("Mill", "AcName", ""), Array(txtMill.Width, 0, 0), " (AgCode=90029 or GpCode=90029)", Chr(KeyAscii), "AcName", txtMill.Left + Me.Left, txtMill.Top + Me.Top + 1000, False
End If
KeyAscii = 0
If gClsSearch.SearchMultiRetCol(1) <> "" Then
    txtMill = gClsSearch.SearchMultiRetCol(1)
    wMillCode = gClsSearch.SearchMultiRetCol(2)

End If
End Sub
'-- Party
Private Sub txtParty_KeyPress(KeyAscii As Integer)
gClsSearch.SearchMultiField "QryMastaccount", "Account,AcName,ACCode", Array("Party", "AcName", ""), Array(txtParty.Width, 0, 0), "", Chr(KeyAscii), "AcName", txtParty.Left + Me.Left, txtParty.Top + Me.Top + 1000, False
KeyAscii = 0
If gClsSearch.SearchMultiRetCol(1) <> "" Then
    txtParty = gClsSearch.SearchMultiRetCol(1)
End If
End Sub
'-- Broker
Private Sub txtBroker_KeyPress(KeyAscii As Integer)
If gSelectedMenu = "TDS Register" Then
   gClsSearch.SearchMultiField "QryMastaccount", "Account,AcName,Accode,AcCity,AcMillTypeStr", Array("Sales Account", "AcName", "acCode", "City", "Mill Type"), Array(3000, 0, 0, 2000, 2000), "", Chr(KeyAscii), "AcName"
Else
   gClsSearch.SearchMultiField "QryMastaccount", "Account,AcName,ACCode", Array("Broker", "AcName", ""), Array(txtBroker.Width, 0, 0), "(AgCode=90016)", Chr(KeyAscii), "AcName", txtBroker.Left + Me.Left, txtBroker.Top + Me.Top + 1000, False
End If
KeyAscii = 0
If gClsSearch.SearchMultiRetCol(1) <> "" Then
    txtBroker = gClsSearch.SearchMultiRetCol(1)
End If
End Sub
'-- Item
Private Sub txtItem_KeyPress(KeyAscii As Integer)
gClsSearch.SearchMultiField "tblMastItem,tblMastAccount", "ItCode,ItName,Itticket,Acname", Array("", "Item", "Ticket", "Mill"), Array(0, 2000, 1000, 2000), " tblMastItem.ItMillcode=tblMastAccount.AcCode", "", "ItName", txtItem.Left + Me.Left, txtItem.Top + Me.Top + 1000
KeyAscii = 0
If gClsSearch.SearchMultiRetCol(1) <> "" Then
     txtItem = gClsSearch.SearchMultiRetCol(1)
     txtItCode = gClsSearch.SearchMultiRetCol(0) '-- ItCode
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
Dim X1 As String
Dim RptOptStr As String
Dim PrevBrkStr As String
Dim X As String
Dim Rs As Recordset
X1 = ""
ValidateData
RptOptStr = IIf(OptGrossAmt = True, "GAmt", IIf(OptNetAmt = True, "NAmt", "Bag"))
With CrystalReport1
    GProcCrystalRptPreparation CrystalReport1, IIf(Index = 0, "Window", "Printer")
    Select Case gSelectedMenu
        Case "For Eof() Bof() Error"
             Dim Del As Boolean
             Dim DelVno As Long
             Dim DelVtype As String
             Del = False
             If OptBrkSAType(0).Value = True Then
                X = "Select Vno from tblAddLess where trim(AdInvNo) = '" & Trim(txt(5)) & "' and vtype = 'SY' and vyear = " & gCYear & "   "
                If GProcRstOpen(Rs, X, "R", gCn) > 0 Then
                   DelVno = Rs.Fields(0)
                   DelVtype = "SY"
                   If MsgBox("Want To Delete Sale Bill No '" & Trim(txt(5)) & "'  ? Are you sure?", vbYesNo) = vbNo Then
                      Exit Sub
                   End If
                Else
                   MsgBox (" Bill No Not Found....")
                   Exit Sub
                End If
             End If
             If OptBrkSAType(1).Value = True Then
                X = "Select Vno from tblAddLess where Vno = " & Val(txt(5)) & " and vtype = 'PY' and vyear = " & gCYear & "   "
                If GProcRstOpen(Rs, X, "R", gCn) > 0 Then
                   DelVno = Rs.Fields(0)
                   DelVtype = "PY"
                   If MsgBox("Want To Delete Purchase Vou. No '" & Trim(txt(5)) & "'  ? Are you sure?", vbYesNo) = vbNo Then
                      Exit Sub
                   End If
                Else
                   MsgBox (" Purchase Vou. No Not Found....")
                   Exit Sub
                End If
             End If
             X = "Select Vno from tblVousub where Vno = " & DelVno & " and vtype = '" & DelVtype & "' and vyear = " & gCYear & "   "
             If GProcRstOpen(Rs, X, "R", gCn) = 0 Then Del = True
             
             X = "Select Vno from tblAddLess where Vno = " & DelVno & " and vtype = '" & DelVtype & "' and vyear = " & gCYear & "   "
             If GProcRstOpen(Rs, X, "R", gCn) = 0 Then Del = True
             
             X = "Select Vno from tblVoucher where Vno = " & DelVno & " and vtype = '" & DelVtype & "' and vyear = " & gCYear & "   "
             If GProcRstOpen(Rs, X, "R", gCn) = 0 Then Del = True
             
             X = "Select Vno from tblOutstanding where Vno = " & DelVno & " and vtype = '" & DelVtype & "' and vyear = " & gCYear & "   "
             If GProcRstOpen(Rs, X, "R", gCn) = 0 Then Del = True
             
             If Del Then
                gCn.BeginTrans
                gCn.Execute "delete from tblVousub where Vno = " & DelVno & " and vtype = '" & DelVtype & "' and vyear = " & gCYear & ""
                gCn.Execute "delete from tblAddLess where Vno = " & DelVno & " and vtype = '" & DelVtype & "' and vyear = " & gCYear & ""
                gCn.Execute "delete from tblVoucher where Vno = " & DelVno & " and vtype = '" & DelVtype & "' and vyear = " & gCYear & ""
                gCn.Execute "delete from tblOutstanding where Vno = " & DelVno & " and vtype = '" & DelVtype & "' and vyear = " & gCYear & ""
                gCn.CommitTrans
                MsgBox ("Deleted....")
             Else
                MsgBox ("Related Record is Exist Can not be delete....")
             End If
             Exit Sub
        '================ Tax Report
        Case "V A T Computation"  '---"V A T Computation
            GProcPrepareVATComputation Format(dtpFrDt.Value, "dd/MM/yyyy"), Format(dtpToDt.Value, "dd/MM/yyyy"), CrystalReport1, CDbl(txt(0))
            Exit Sub
        Case "V A T Computation 2 And 5 %"
            GProcPrepareVATComputation25 Format(dtpFrDt.Value, "dd/MM/yyyy"), Format(dtpToDt.Value, "dd/MM/yyyy"), CrystalReport1, CDbl(txt(0))
            Exit Sub
        Case "V A T Computation - Detail" '---"V A T Computation - Detail"
            GProcPrepareVATComputDet Format(dtpFrDt.Value, "dd/MM/yyyy"), Format(dtpToDt.Value, "dd/MM/yyyy"), CrystalReport1
            Exit Sub
        Case "Purchase Tax Return Register" '-- "Purchase Tax Return Register"
        Case "Exemption Sales Detail"  ' "Exemption Sales Detail"
            X1 = "{tblAddLess.VDt} in Date(" & Format(dtpFrDt.Value, "yyyy,mm,dd") & ") to Date(" & Format(dtpToDt.Value, "yyyy,mm,dd") & ")  " '
            X1 = X1 + " and ({tblAddLess.VType}='SY') and {tblAddLess.ADEXEMPTAMT}>0 and {tblVouSub_GP.VType}='GP'"
            .SelectionFormula = X1
            .Formulas(4) = "wFrToDt = 'From : ' & '" & Format(dtpFrDt.Value, "dd/MM/yyyy") & "' & '  To ' & '" & Format(dtpToDt.Value, "dd/MM/yyyy") & "'"
            .ReportFileName = gReportPath & "rptTaxExmptSalesReg.rpt"
        Case "Purchase Register" '--"Purchase Register"
            X1 = "{tblAddLess.VDt} in Date(" & Format(dtpFrDt.Value, "yyyy,mm,dd") & ") to Date(" & Format(dtpToDt.Value, "yyyy,mm,dd") & ")  " '
            X1 = X1 + " and ({tblAddLess.VType}='PY' or {tblAddLess.VType}='PT' )"
            .SelectionFormula = X1
            .Formulas(4) = "wFrToDt = 'From : ' & '" & Format(dtpFrDt.Value, "dd/MM/yyyy") & "' & '  To ' & '" & Format(dtpToDt.Value, "dd/MM/yyyy") & "'"
            .ReportFileName = gReportPath & "rptTaxPurchaseReg.rpt"
        Case "Sales Register" '-- "Sales Register"
            X1 = "{tblAddLess.VDt} in Date(" & Format(dtpFrDt.Value, "yyyy,mm,dd") & ") to Date(" & Format(dtpToDt.Value, "yyyy,mm,dd") & ")  " '
            X1 = X1 + " and ({tblAddLess.VType}='ST' or {tblAddLess.VType}='SY' or {tblAddLess.VType}='SD' or {tblAddLess.VType}='SO' )"
            .SelectionFormula = X1
            .Formulas(4) = "wFrToDt = 'From : ' & '" & Format(dtpFrDt.Value, "dd/MM/yyyy") & "' & '  To ' & '" & Format(dtpToDt.Value, "dd/MM/yyyy") & "'"
            .ReportFileName = gReportPath & "rptTaxSaleReg.rpt"
        Case "Partywise Sales VAT Audit" '-- "Partywise Sales VAT Audit"
             X1 = "{tblAddLess.VDt} in Date(" & Format(dtpFrDt.Value, "yyyy,mm,dd") & ") to Date(" & Format(dtpToDt.Value, "yyyy,mm,dd") & ")  " '
             If txtMill <> "" Then
                X1 = X1 + "and ({tblAddLess.VType}='SY' or {tblAddLess.VType}='SD' or {tblAddLess.VType}='SO')"
                X1 = X1 & " and {tblAddLess.AdAcCrcode}=" & wMillCode
             Else
                X1 = X1 + " and ({tblAddLess.VType}='SY' or {tblAddLess.VType}='SD' or {tblAddLess.VType}='SO' )"
             End If
'             X1 = X1 + " and ({tblAddLess.VType}='SY')"
            .SelectionFormula = X1
            .Formulas(4) = "wFrToDt = 'From : ' & '" & Format(dtpFrDt.Value, "dd/MM/yyyy") & "' & '  To ' & '" & Format(dtpToDt.Value, "dd/MM/yyyy") & "'"
            .Formulas(5) = "wSaleAc = 'Sales A/c : ' & '" & txtMill & "'"
            .ReportFileName = gReportPath & "rptTaxAuditSaleReg.rpt"
        Case "Partywise Purchase VAT Audit" '-- "Partywise Purchase VAT Audit"
            X1 = "{tblAddLess.VDt} in Date(" & Format(dtpFrDt.Value, "yyyy,mm,dd") & ") to Date(" & Format(dtpToDt.Value, "yyyy,mm,dd") & ")  " '
'            X1 = X1 + " and ({tblAddLess.VType}='PY' or {tblAddLess.VType}='PT' )"
            X1 = X1 + " and ({tblAddLess.VType}='PY')"
            .SelectionFormula = X1
            .Formulas(4) = "wFrToDt = 'From : ' & '" & Format(dtpFrDt.Value, "dd/MM/yyyy") & "' & '  To ' & '" & Format(dtpToDt.Value, "dd/MM/yyyy") & "'"
            .ReportFileName = gReportPath & "rptTaxAuditPurchaseReg.rpt"
        Case "Monthly Sales Detail VAT Audit" '-- "Monthly Sales Detail VAT Audit"
            X1 = "{tblAddLess.VDt} in Date(" & Format(dtpFrDt.Value, "yyyy,mm,dd") & ") to Date(" & Format(dtpToDt.Value, "yyyy,mm,dd") & ")  " '
            X1 = X1 + " and ({tblAddLess.VType}='ST' or {tblAddLess.VType}='SY' or {tblAddLess.VType}='SD' or {tblAddLess.VType}='SO' )"
            .SelectionFormula = X1
            .Formulas(4) = "wFrToDt = 'From : ' & '" & Format(dtpFrDt.Value, "dd/MM/yyyy") & "' & '  To ' & '" & Format(dtpToDt.Value, "dd/MM/yyyy") & "'"
            .ReportFileName = gReportPath & "rptTaxAuditSalesMonthly.rpt"
        Case "Monthly Purchase Detail VAT Audit" '---"Monthly Purchase Detail VAT Audit"
             X1 = "{tblAddLess.VDt} in Date(" & Format(dtpFrDt.Value, "yyyy,mm,dd") & ") to Date(" & Format(dtpToDt.Value, "yyyy,mm,dd") & ")  " '
             X1 = X1 + " and ({tblAddLess.VType}='PY' or {tblAddLess.VType}='PT' or {tblAddLess.VType}='PI')"
             .SelectionFormula = X1
             .Formulas(4) = "wFrToDt = 'From : ' & '" & Format(dtpFrDt.Value, "dd/MM/yyyy") & "' & '  To ' & '" & Format(dtpToDt.Value, "dd/MM/yyyy") & "'"
             .ReportFileName = gReportPath & "rptTaxAuditPurchaseMonthly.rpt"
            
         Case "Annexure - G (Received C And E1 Form)" '--Annexure - G (Received C And E1 Form)
              formName1 = "C Form"
              formName2 = "E1 Form"
              gCn.Execute "delete from tmpGenTbl where Username='" & gUserName & "'"
              If gBackEndDB = gBackEndAccess Then
                 CondStr = " tbladdLess.vdt>= cdate('" & Format(dtpFrDt.Value, "dd/MM/yyyy") & "')  and tbladdLess.vdt<= cdate('" & Format(dtpToDt.Value, "dd/MM/yyyy") & "')"
                 CondStr = CondStr + " and ADCFORMNO <>''"
              Else '-- Oracle
                 CondStr = " tbladdLess.vdt>= to_date('" & dtpFrDt.Value & "','dd/MM/yyyy') and tbladdLess.vdt<= to_date('" & dtpToDt.Value & "','dd/MM/yyyy')"
                 CondStr = CondStr + " and ADCFORMNO is not null "
              End If
              '-- C form
              X1 = " select sum(ADBILLAMT),TblMastAccount.AcName,'" & gUserName & "'" _
              & " ,ADCFORMNO,TblMastAccount.AcBST,tblmastaccount.accity,tblmastaccount.AcState, ' " & formName1 & "' from tbladdLess,tblmastaccount,tblMastNarration where" _
              & " ADACDRCODE=tblmastaccount.Accode and " _
              & " ADTAXCODE=NARRCODE and instr(narration,'C Form')>0 and tbladdLess.Vtype in ('SY','ST','SO')  and " & CondStr & "group by TblMastAccount.AcName,ADCFORMNO,TblMastAccount.AcBst,tblmastaccount.accity,tblmastaccount.AcState "
    
              gCn.Execute "insert into tmpGentbl (Amt1,Nar1,username,nar2,nar3,nar4,Nar5,Nar6)" & X1
              
              '-- E1 Form
              
              If gBackEndDB = gBackEndAccess Then
                 CondStr = " tbladdLess.vdt>= cdate('" & Format(dtpFrDt.Value, "dd/MM/yyyy") & "')  and tbladdLess.vdt<= cdate('" & Format(dtpToDt.Value, "dd/MM/yyyy") & "')"
                 CondStr = CondStr + " and ADE1FORMNO <>''"
              Else '-- Oracle
                 CondStr = " tbladdLess.vdt>= to_date('" & dtpFrDt.Value & "','dd/MM/yyyy') and tbladdLess.vdt<= to_date('" & dtpToDt.Value & "','dd/MM/yyyy')"
                 CondStr = CondStr + " and ADE1FORMNO is not null "
              End If
              X1 = " select sum(ADBILLAMT),TblMastAccount.AcName,'" & gUserName & "'" _
              & " ,ADE1FORMNO,TblMastAccount.AcCst,tblmastaccount.accity,tblmastaccount.AcState, ' " & formName2 & "' from tbladdLess,tblmastaccount,tblMastNarration where" _
              & " ADACCRCODE=tblmastaccount.Accode and " _
              & " ADTAXCODE=NARRCODE and instr(narration,'C Form')>0 and tbladdLess.Vtype in ('PT')  and " & CondStr & "group by TblMastAccount.AcName,ADE1FORMNO,TblMastAccount.AcCst,tblmastaccount.accity,tblmastaccount.AcState "
              gCn.Execute "insert into tmpGentbl (Amt1,Nar1,username,nar2,nar3,nar4,Nar5,Nar6)" & X1
              X1 = "{tmpGenTbl.UserName}='" & gUserName & "'"
             .SelectionFormula = X1
             .Formulas(4) = "wFrToDt = 'From : ' & '" & Format(dtpFrDt.Value, "dd/MM/yyyy") & "' & '  To ' & '" & Format(dtpToDt.Value, "dd/MM/yyyy") & "'"
             .ReportFileName = gReportPath & "rptTaxRptCFormRecdVatAuditNew.rpt"
         
         Case "Annexure - I  (Pending  C And E1 Form)"
              formName1 = "C Form"
              formName2 = "E1 Form"
              gCn.Execute "delete from tmpGenTbl where Username='" & gUserName & "'"
              If gBackEndDB = gBackEndAccess Then
                 CondStr = " tbladdLess.vdt>= cdate('" & Format(dtpFrDt.Value, "dd/MM/yyyy") & "')  and tbladdLess.vdt<= cdate('" & Format(dtpToDt.Value, "dd/MM/yyyy") & "')"
'                 CondStr = CondStr + " and ADCFORMNO ='' or ADCFORMNO is null"
                 CondStr = CondStr + " and (ADCFORMNO ='' or ADCFORMNO is null)"
              
              Else '-- Oracle
                 CondStr = " tbladdLess.vdt>= to_date('" & dtpFrDt.Value & "','dd/MM/yyyy') and tbladdLess.vdt<= to_date('" & dtpToDt.Value & "','dd/MM/yyyy')"
                 CondStr = CondStr + " and ADCFORMNO is null "
              End If
              '-- C form
              X1 = " select tblAddLess.AdInvNo,tblAddLess.VDt,tblAddLess.AdAssVal,tblAddLess.AdTaxAmt,tblAddLess.AdTaxP,tblAddLess.AdBillAmt,TblMastAccount.AcName,'" & gUserName & "'" _
              & " ,ADCFORMNO,TblMastAccount.AcBST,tblmastaccount.accity,tblmastaccount.AcState, ' " & formName1 & "' from tbladdLess,tblmastaccount,tblMastNarration where" _
              & " ADACDRCODE=tblmastaccount.Accode and " _
              & " ADTAXCODE=NARRCODE and (instr(narration,'C Form')>0  or instr(narration,'C S T')>0) and tbladdLess.Vtype in ('SY','ST','SO')  and " & CondStr
    
              gCn.Execute "insert into tmpGentbl (Nar7,Vdt,Amt1,Amt2,Amt3,Amt4,Nar1,username,nar2,nar3,nar4,Nar5,Nar6)" & X1
              
              '-- E1 Form
              
              If gBackEndDB = gBackEndAccess Then
                 CondStr = " tbladdLess.vdt>= cdate('" & Format(dtpFrDt.Value, "dd/MM/yyyy") & "')  and tbladdLess.vdt<= cdate('" & Format(dtpToDt.Value, "dd/MM/yyyy") & "')"
                 CondStr = CondStr + " and (ADE1FORMNO ='' or ADE1FORMNO is null)"
              Else '-- Oracle
                 CondStr = " tbladdLess.vdt>= to_date('" & dtpFrDt.Value & "','dd/MM/yyyy') and tbladdLess.vdt<= to_date('" & dtpToDt.Value & "','dd/MM/yyyy')"
                 CondStr = CondStr + " and ADE1FORMNO is null "
              End If
              X1 = " select tblAddLess.AdInvNo,tblAddLess.VDt,tblAddLess.AdAssVal,tblAddLess.AdTaxAmt,tblAddLess.AdTaxP,tblAddLess.AdBillAmt,TblMastAccount.AcName,'" & gUserName & "'" _
              & " ,ADE1FORMNO,TblMastAccount.AcCst,tblmastaccount.accity,tblmastaccount.AcState, ' " & formName2 & "' from tbladdLess,tblmastaccount,tblMastNarration where" _
              & " ADACCRCODE=tblmastaccount.Accode and " _
              & " ADTAXCODE=NARRCODE and instr(narration,'C Form')>0 and tbladdLess.Vtype in ('PT')  and " & CondStr
              
              gCn.Execute "insert into tmpGentbl (Nar7,Vdt,Amt1,Amt2,Amt3,Amt4,Nar1,username,nar2,nar3,nar4,Nar5,Nar6)" & X1
              
              X1 = "{tmpGenTbl.UserName}='" & gUserName & "'"
             .SelectionFormula = X1
             .Formulas(4) = "wFrToDt = 'From : ' & '" & Format(dtpFrDt.Value, "dd/MM/yyyy") & "' & '  To ' & '" & Format(dtpToDt.Value, "dd/MM/yyyy") & "'"
             .ReportFileName = gReportPath & "rptTaxRptCFormRecVatAuditNew.rpt"
         
         Case "Annexure - J Section -1   (Partywise VAT Sales)", "Annexure - J Section -1   (R.D. Partywise VAT Sales)", "Annexure - J Section -1   (U.R.D. Partywise VAT Sales)"
              gCn.Execute "delete from tmpGenTbl where Username='" & gUserName & "'"
              If gBackEndDB = gBackEndAccess Then
                 Select Case gSelectedMenu
                 Case "Annexure - J Section -1   (Partywise VAT Sales)"
                     CondStr = " tbladdLess.vdt>= cdate('" & Format(dtpFrDt.Value, "dd/MM/yyyy") & "')  and tbladdLess.vdt<= cdate('" & Format(dtpToDt.Value, "dd/MM/yyyy") & "')"
                     .Formulas(5) = "RDURD = 'RD & URD Partywise VAT Sales' "
                 Case "Annexure - J Section -1   (R.D. Partywise VAT Sales)"
                     CondStr = " tbladdLess.vdt>= cdate('" & Format(dtpFrDt.Value, "dd/MM/yyyy") & "')  and tbladdLess.vdt<= cdate('" & Format(dtpToDt.Value, "dd/MM/yyyy") & "')"
                     CondStr = CondStr + " and (tblMastAccount.AcBst <>'' or tblMastAccount.AcBst is not null)"
                     .Formulas(5) = "RDURD = 'RD Partywise VAT Sales' "
                 Case "Annexure - J Section -1   (U.R.D. Partywise VAT Sales)"
                     CondStr = " tbladdLess.vdt>= cdate('" & Format(dtpFrDt.Value, "dd/MM/yyyy") & "')  and tbladdLess.vdt<= cdate('" & Format(dtpToDt.Value, "dd/MM/yyyy") & "')"
                     CondStr = CondStr + " and (tblMastAccount.AcBst ='' or tblMastAccount.AcBst is null)"
                     .Formulas(5) = "RDURD = 'URD Partywise VAT Sales' "
                 End Select
              Else '-- Oracle
                 CondStr = " tbladdLess.vdt>= to_date('" & dtpFrDt.Value & "','dd/MM/yyyy') and tbladdLess.vdt<= to_date('" & dtpToDt.Value & "','dd/MM/yyyy')"
                 CondStr = CondStr + " and ADCFORMNO is not null "
              End If
              X1 = " select sum(ADBILLAMT),sum(AdAssVal),sum(AdTaxAmt),TblMastAccount.AcName,'" & gUserName & "'" _
              & " ,TblMastAccount.AcBST,tblmastaccount.accity,tblmastaccount.AcState from tbladdLess,tblmastaccount,tblMastNarration where" _
              & " ADACDRCODE=tblmastaccount.Accode and " _
              & " ADTAXCODE=NARRCODE and instr(narration,'V A T')>0 and tbladdLess.Vtype in ('SY','SO')  and " & CondStr & "group by TblMastAccount.AcName,TblMastAccount.AcBst,tblmastaccount.accity,tblmastaccount.AcState "
    
              gCn.Execute "insert into tmpGentbl (Amt1,Amt2,Amt3,Nar1,username,nar3,nar4,Nar5)" & X1
             
            .SelectionFormula = "{tmpgenTbl.Username}='" & gUserName & "'"
            .Formulas(4) = "wFrToDt = 'From : ' & '" & Format(dtpFrDt.Value, "dd/MM/yyyy") & "' & '  To ' & '" & Format(dtpToDt.Value, "dd/MM/yyyy") & "'"
            .ReportFileName = gReportPath & "rptTaxAuditSaleRegNew.rpt"
         Case "Annexure - J Section -2   (Partywise VAT Purchases)"
              gCn.Execute "delete from tmpGenTbl where Username='" & gUserName & "'"
              If gBackEndDB = gBackEndAccess Then
                 CondStr = " tbladdLess.vdt>= cdate('" & Format(dtpFrDt.Value, "dd/MM/yyyy") & "')  and tbladdLess.vdt<= cdate('" & Format(dtpToDt.Value, "dd/MM/yyyy") & "')"
              Else '-- Oracle
                 CondStr = " tbladdLess.vdt>= to_date('" & dtpFrDt.Value & "','dd/MM/yyyy') and tbladdLess.vdt<= to_date('" & dtpToDt.Value & "','dd/MM/yyyy')"
                 CondStr = CondStr + " and ADCFORMNO is not null "
              End If
              X1 = " select sum(ADBILLAMT),sum(AdAssVal),sum(AdTaxAmt),TblMastAccount.AcName,'" & gUserName & "'" _
              & " ,TblMastAccount.AcBST,tblmastaccount.accity,tblmastaccount.AcState from tbladdLess,tblmastaccount,tblMastNarration where" _
              & " ADACCRCODE=tblmastaccount.Accode and " _
              & " ADTAXCODE=NARRCODE and instr(narration,'V A T')>0 and tbladdLess.Vtype in ('PY')  and " & CondStr & "group by TblMastAccount.AcName,TblMastAccount.AcBst,tblmastaccount.accity,tblmastaccount.AcState "
    
              gCn.Execute "insert into tmpGentbl (Amt1,Amt2,Amt3,Nar1,username,nar3,nar4,Nar5)" & X1
             
            .SelectionFormula = "{tmpgenTbl.Username}='" & gUserName & "'"
            .Formulas(4) = "wFrToDt = 'From : ' & '" & Format(dtpFrDt.Value, "dd/MM/yyyy") & "' & '  To ' & '" & Format(dtpToDt.Value, "dd/MM/yyyy") & "'"
            .ReportFileName = gReportPath & "rptTaxAuditPurchaseRegNew.rpt"
         Case "Annexure - J Section -6   (Supllierwise CST Purchases)"
              formName1 = "Oms Purchase"
              formName2 = "Sale In Transit"
              formName3 = "Branch Transfer"
             
              gCn.Execute "delete from tmpGenTbl where Username='" & gUserName & "'"
              If gBackEndDB = gBackEndAccess Then
                 CondStr = " tbladdLess.vdt>= cdate('" & Format(dtpFrDt.Value, "dd/MM/yyyy") & "')  and tbladdLess.vdt<= cdate('" & Format(dtpToDt.Value, "dd/MM/yyyy") & "')"
              Else '-- Oracle
                 CondStr = " tbladdLess.vdt>= to_date('" & dtpFrDt.Value & "','dd/MM/yyyy') and tbladdLess.vdt<= to_date('" & dtpToDt.Value & "','dd/MM/yyyy')"
                 CondStr = CondStr + " and ADCFORMNO is not null "
              End If
              '-- Oms Purchase
              X1 = " select sum(ADBILLAMT),sum(AdAssVal),sum(AdTaxAmt),TblMastAccount.AcName,'" & gUserName & "'" _
              & " ,TblMastAccount.AcBST,tblmastaccount.accity,tblmastaccount.AcState,  ' " & formName1 & "' from tbladdLess,tblmastaccount,tblMastNarration where" _
              & " ADACCRCODE=tblmastaccount.Accode and " _
              & " ADTAXCODE=NARRCODE and instr(narration,'C S T')>0 and tbladdLess.Vtype in ('PY')  and " & CondStr & "group by TblMastAccount.AcName,TblMastAccount.AcBst,tblmastaccount.accity,tblmastaccount.AcState "
    
              gCn.Execute "insert into tmpGentbl (Amt1,Amt2,Amt3,Nar1,username,nar3,nar4,Nar5,Nar6)" & X1
             
              '-- SIT Purchase
              X1 = " select sum(ADBILLAMT),sum(AdAssVal),sum(AdTaxAmt),TblMastAccount.AcName,'" & gUserName & "'" _
              & " ,TblMastAccount.AcBST,tblmastaccount.accity,tblmastaccount.AcState, ' " & formName2 & "' from tbladdLess,tblmastaccount,tblMastNarration where" _
              & " ADACCRCODE=tblmastaccount.Accode and " _
              & " ADTAXCODE=NARRCODE and tbladdLess.Vtype in ('PT')  and " & CondStr & "group by TblMastAccount.AcName,TblMastAccount.AcBst,tblmastaccount.accity,tblmastaccount.AcState "
    
              gCn.Execute "insert into tmpGentbl (Amt1,Amt2,Amt3,Nar1,username,nar3,nar4,Nar5,Nar6)" & X1
            
              '-- Branch Tfr
              X1 = " select sum(ADBILLAMT),sum(AdAssVal),sum(AdTaxAmt),TblMastAccount.AcName,'" & gUserName & "'" _
              & " ,TblMastAccount.AcBST,tblmastaccount.accity,tblmastaccount.AcState, ' " & formName3 & "' from tbladdLess,tblmastaccount,tblMastNarration where" _
              & " ADACCRCODE=tblmastaccount.Accode and " _
              & " ADTAXCODE=NARRCODE and tbladdLess.Vtype in ('PI')  and " & CondStr & "group by TblMastAccount.AcName,TblMastAccount.AcBst,tblmastaccount.accity,tblmastaccount.AcState "
    
              gCn.Execute "insert into tmpGentbl (Amt1,Amt2,Amt3,Nar1,username,nar3,nar4,Nar5,Nar6)" & X1
            .SelectionFormula = "{tmpgenTbl.Username}='" & gUserName & "'"
            .Formulas(4) = "wFrToDt = 'From : ' & '" & Format(dtpFrDt.Value, "dd/MM/yyyy") & "' & '  To ' & '" & Format(dtpToDt.Value, "dd/MM/yyyy") & "'"
            .ReportFileName = gReportPath & "rptTaxAuditCSTPurNew.rpt"
            
         Case "Annexure - J Section -6   (Supllierwise CST Sales)"
             
              gCn.Execute "delete from tmpGenTbl where Username='" & gUserName & "'"
              If gBackEndDB = gBackEndAccess Then
                 CondStr = " tbladdLess.vdt>= cdate('" & Format(dtpFrDt.Value, "dd/MM/yyyy") & "')  and tbladdLess.vdt<= cdate('" & Format(dtpToDt.Value, "dd/MM/yyyy") & "')"
              Else '-- Oracle
                 CondStr = " tbladdLess.vdt>= to_date('" & dtpFrDt.Value & "','dd/MM/yyyy') and tbladdLess.vdt<= to_date('" & dtpToDt.Value & "','dd/MM/yyyy')"
                 CondStr = CondStr + " and ADCFORMNO is not null "
              End If
             
              '-- SIT Sales
              X1 = " select sum(ADBILLAMT),sum(AdAssVal),sum(AdTaxAmt),TblMastAccount.AcName,'" & gUserName & "'" _
              & " ,TblMastAccount.AcBST,TblMastAccount.AcCST,tblmastaccount.accity,tblmastaccount.AcState, ' " & formName2 & "' from tbladdLess,tblmastaccount,tblMastNarration where" _
              & " ADACdRCODE=tblmastaccount.Accode and " _
              & " ADTAXCODE=NARRCODE and tbladdLess.Vtype in ('ST')  and " & CondStr & "group by TblMastAccount.AcName,TblMastAccount.AcBst,TblMastAccount.AcCst,tblmastaccount.accity,tblmastaccount.AcState "
    
              gCn.Execute "insert into tmpGentbl (Amt1,Amt2,Amt3,Nar1,username,nar3,nar7,nar4,Nar5,Nar6)" & X1
            
            .SelectionFormula = "{tmpgenTbl.Username}='" & gUserName & "'"
            .Formulas(4) = "wFrToDt = 'From : ' & '" & Format(dtpFrDt.Value, "dd/MM/yyyy") & "' & '  To ' & '" & Format(dtpToDt.Value, "dd/MM/yyyy") & "'"
            .ReportFileName = gReportPath & "rptTaxAuditCSTSaleNew.rpt"
            
            
        Case "C Form Receivable" '-- "C Form Receivable"
            GProcFormIssueReceive Format(dtpFrDt.Value, "dd/MM/yyyy"), Format(dtpToDt.Value, "dd/MM/yyyy"), "CR", txtMill, txtParty, txtBroker, ChkConsolidate, IIf(Opt(0) = True, "All", IIf(Opt(1) = True, "Pending", "Receive"))
            GProcCrystalRptPreparation CrystalReport1, IIf(Index = 0, "Window", "Printer")
            .Formulas(4) = "wFrToDt = 'From : ' & '" & Format(dtpFrDt.Value, "dd/MM/yyyy") & "' & '  To ' & '" & Format(dtpToDt.Value, "dd/MM/yyyy") & "'"
            If ChkConsolidate = 1 Then .Formulas(6) = "wAllFirmStr= 'All Firm'" '- wAllFirmStr
            .Formulas(15) = "wCSTNo = '" & gCCST & "'"
            .Formulas(16) = "wBSTNo = '" & gCBST & "'"
            .Formulas(17) = "wPanNo = '" & gCPAN & "'"
            .DataFiles(0) = App.Path & "\HIData\setup.mdb"
            .SelectionFormula = "{tmpgenTbl.Username}='" & gUserName & "'"
            gReportPath = App.Path & "\HiReports\"
            '-- Opt Str
            If Opt(0) = True Then
                '.Formulas(7) = "Str='All'"
            ElseIf Opt(1) = True Then
                .Formulas(7) = "Str='Pending'"
            Else
                .Formulas(7) = "Str='Received'"
            End If
            If OptGrossAmt.Value = True Then
                .ReportFileName = gReportPath & "rptTaxRptCFormRecNew.rpt"
            Else
            If OptBrk = True Then '-- Broker wise
                If ChkSummary = 1 Then '--summary
                    .ReportFileName = gReportPath & "rptTaxRptCFormRecSumm.rpt"
                Else '--detail
                    .Formulas(8) = "IsNewPage='" & ChkPartyOnNewPage & "'" '-- Partywise Print
                    '-for Received Show Form No
                    If Opt(2) = True Then
                        .Formulas(9) = "IsShowFormNo='1'"
                    Else
                        .Formulas(9) = "IsShowFormNo='0'"
                    End If
                    .ReportFileName = gReportPath & "rptTaxRptCFormRec.rpt"
                End If
            Else '-- Party wise
                If ChkSummary = 1 Then  '--summary
                       .Formulas(14) = "wCSTNo = '" & gCCST & "'"
                       .Formulas(5) = "wBSTNo = '" & gCBST & "'"
                      .ReportFileName = gReportPath & "rptTaxRptCFormRecPartywiseSumm.rpt"
                Else '--detail
                    .Formulas(8) = "IsNewPage='" & ChkPartyOnNewPage & "'" '-- Partywise Print
                    '-for Received Show Form No
                    If Opt(2) = True Then
                        .Formulas(9) = "IsShowFormNo='1'"
                    Else
                        .Formulas(9) = "IsShowFormNo='0'"
                    End If
                    .ReportFileName = gReportPath & "rptTaxRptCFormRecPartywise.rpt"
                End If
            End If
            End If
        Case "C Form Issued But E1Form Not Received"
            GProcFormIssueReceive Format(dtpFrDt.Value, "dd/MM/yyyy"), Format(dtpToDt.Value, "dd/MM/yyyy"), "CIButNotE1", "", txtParty, txtBroker, ChkConsolidate, IIf(Opt(0) = True, "All", IIf(Opt(1) = True, "Pending", "Receive"))
            GProcCrystalRptPreparation CrystalReport1, IIf(Index = 0, "Window", "Printer")
            .Formulas(4) = "wFrToDt = 'From : ' & '" & Format(dtpFrDt.Value, "dd/MM/yyyy") & "' & '  To ' & '" & Format(dtpToDt.Value, "dd/MM/yyyy") & "'"
            .SelectionFormula = "{tmpgenTbl.Username}='" & gUserName & "'"
            .DataFiles(0) = App.Path & "\HIData\setup.mdb"
            gReportPath = App.Path & "\HiReports\"
            .Formulas(14) = "wCSTNo = '" & gCCST & "'"
            .Formulas(5) = "wBSTNo = '" & gCBST & "'"
            .ReportFileName = gReportPath & "rptTaxRptCFormIssuedE1NotRecd.rpt"
        
        
        Case "C Form Issuable"  '-- "C Form Issuable"
            GProcFormIssueReceive Format(dtpFrDt.Value, "dd/MM/yyyy"), Format(dtpToDt.Value, "dd/MM/yyyy"), "CI", "", txtParty, "", ChkConsolidate, IIf(Opt(0) = True, "All", IIf(Opt(1) = True, "Pending", "Receive"))
            GProcCrystalRptPreparation CrystalReport1, IIf(Index = 0, "Window", "Printer")
            .Formulas(4) = "wFrToDt = 'From : ' & '" & Format(dtpFrDt.Value, "dd/MM/yyyy") & "' & '  To ' & '" & Format(dtpToDt.Value, "dd/MM/yyyy") & "'"
            If ChkConsolidate = 1 Then .Formulas(6) = "wAllFirmStr= 'All Firm'" '- wAllFirmStr
            .Formulas(8) = "IsNewPage='" & ChkPartyOnNewPage & "'" '-- Partywise Print
            '-- Opt Str
            If Opt(0) = True Then
                '.Formulas(7) = "Str='All'"
            ElseIf Opt(1) = True Then
                .Formulas(7) = "Str='Pending'"
            Else
                .Formulas(7) = "Str='Issued'"
            End If
            '-for Received Show Form No
'            If Opt(2) = True Then
                .Formulas(9) = "IsShowFormNo='1'"
 '           Else
  '              .Formulas(9) = "IsShowFormNo='0'"
   '         End If
            .SelectionFormula = "{tmpgenTbl.Username}='" & gUserName & "'"
            .DataFiles(0) = App.Path & "\HIData\setup.mdb"
            gReportPath = App.Path & "\HiReports\"
            If ChkSummary = 1 Then  '--summary
               .Formulas(14) = "wCSTNo = '" & gCCST & "'"
               .Formulas(5) = "wBSTNo = '" & gCBST & "'"
               .ReportFileName = gReportPath & "rptTaxRptCFormIssueSumm.rpt"
            Else
               .ReportFileName = gReportPath & "rptTaxRptCFormIssue.rpt"
            End If
           If OptParty = True Then
               .ReportFileName = gReportPath & "rptTaxRptCFormIssueLtr.rpt"
           End If
           If OptBrk = True Then
               .ReportFileName = gReportPath & "rptTaxRptCFormIssueshort.rpt"
           End If
            
        Case "F Form Issuable"  '== "F Form Issuable"
            GProcFormIssueReceive Format(dtpFrDt.Value, "dd/MM/yyyy"), Format(dtpToDt.Value, "dd/MM/yyyy"), "FI", "", txtParty, "", ChkConsolidate, IIf(Opt(0) = True, "All", IIf(Opt(1) = True, "Pending", "Receive"))
            GProcCrystalRptPreparation CrystalReport1, IIf(Index = 0, "Window", "Printer")
            .Formulas(4) = "wFrToDt = 'From : ' & '" & Format(dtpFrDt.Value, "dd/MM/yyyy") & "' & '  To ' & '" & Format(dtpToDt.Value, "dd/MM/yyyy") & "'"
            If ChkConsolidate = 1 Then .Formulas(6) = "wAllFirmStr= 'All Firm'" '- wAllFirmStr
            .Formulas(8) = "IsNewPage='" & ChkPartyOnNewPage & "'" '-- Partywise Print
            '-- Opt Str
            If Opt(0) = True Then
                '.Formulas(7) = "Str='All'"
            ElseIf Opt(1) = True Then
                .Formulas(7) = "Str='Pending'"
            Else
                .Formulas(7) = "Str='Received'"
            End If
            '-for Received Show Form No
            If Opt(2) = True Then
                .Formulas(9) = "IsShowFormNo='1'"
            Else
                .Formulas(9) = "IsShowFormNo='0'"
            End If
            .SelectionFormula = "{tmpgenTbl.Username}='" & gUserName & "'"
            .DataFiles(0) = App.Path & "\HIData\setup.mdb"
            gReportPath = App.Path & "\HiReports\"
            If ChkSummary = 1 Then  '--summary
               .Formulas(14) = "wCSTNo = '" & gCCST & "'"
               .Formulas(5) = "wBSTNo = '" & gCBST & "'"
               .ReportFileName = gReportPath & "rptTaxRptFFormIssueSumm.rpt"
            Else
               .ReportFileName = gReportPath & "rptTaxRptFFormIssue.rpt"
            End If
        Case "E1 Form Receivable" '-- "E1 Form Receivable"
            GProcFormIssueReceive Format(dtpFrDt.Value, "dd/MM/yyyy"), Format(dtpToDt.Value, "dd/MM/yyyy"), "E1R", txtMill, txtParty, txtBroker, ChkConsolidate, IIf(Opt(0) = True, "All", IIf(Opt(1) = True, "Pending", "Receive"))
            GProcCrystalRptPreparation CrystalReport1, IIf(Index = 0, "Window", "Printer")
            .Formulas(4) = "wFrToDt = 'From : ' & '" & Format(dtpFrDt.Value, "dd/MM/yyyy") & "' & '  To ' & '" & Format(dtpToDt.Value, "dd/MM/yyyy") & "'"
            If ChkConsolidate = 1 Then .Formulas(6) = "wAllFirmStr= 'All Firm'" '- wAllFirmStr
            .Formulas(8) = "IsNewPage='" & ChkPartyOnNewPage & "'" '-- Partywise Print
            '-- Opt Str
            If Opt(0) = True Then
                '.Formulas(7) = "Str='All'"
            ElseIf Opt(1) = True Then
                .Formulas(7) = "Str='Pending'"
            Else
                .Formulas(7) = "Str='Received'"
            End If
             '-for Received Show Form No
    '        If Opt(2) = True Then
                .Formulas(9) = "IsShowFormNo='1'"
     '       Else
      '          .Formulas(9) = "IsShowFormNo='0'"
       '     End If
            .SelectionFormula = "{tmpgenTbl.Username}='" & gUserName & "'"
            .DataFiles(0) = App.Path & "\HIData\setup.mdb"
            gReportPath = App.Path & "\HiReports\"
            If ChkSummary = 1 Then  '--summary
               .Formulas(14) = "wCSTNo = '" & gCCST & "'"
               .Formulas(5) = "wBSTNo = '" & gCBST & "'"
               .ReportFileName = gReportPath & "rptTaxRptE1RecSumBrokerwise.rpt"
            Else
                If OptParty.Value = True Then
                   .ReportFileName = gReportPath & "rptTaxRptE1CovrLtr.rpt"
                Else
                   If OptGrossAmt.Value = True Then
                        .ReportFileName = gReportPath & "rptTaxRptE1RecSumPartywise.rpt"
                   Else
                       .ReportFileName = gReportPath & "rptTaxRptE1FormRec.rpt"
                   End If
                End If
'               .ReportFileName = gReportPath & "rptTaxRptE1FormRec.rpt"
            End If
        Case "E2 Form Receivable" '-- "E2 Form Receivable"
            GProcFormIssueReceive Format(dtpFrDt.Value, "dd/MM/yyyy"), Format(dtpToDt.Value, "dd/MM/yyyy"), "E2R", "", txtParty, "", ChkConsolidate, IIf(Opt(0) = True, "All", IIf(Opt(1) = True, "Pending", "Receive"))
            GProcCrystalRptPreparation CrystalReport1, IIf(Index = 0, "Window", "Printer")
            .Formulas(4) = "wFrToDt = 'From : ' & '" & Format(dtpFrDt.Value, "dd/MM/yyyy") & "' & '  To ' & '" & Format(dtpToDt.Value, "dd/MM/yyyy") & "'"
            If ChkConsolidate = 1 Then .Formulas(6) = "wAllFirmStr= 'All Firm'" '- wAllFirmStr
            .Formulas(8) = "IsNewPage='" & ChkPartyOnNewPage & "'" '-- Partywise Print
            '-- Opt Str
            If Opt(0) = True Then
                '.Formulas(7) = "Str='All'"
            ElseIf Opt(1) = True Then
                .Formulas(7) = "Str='Pending'"
            Else
                .Formulas(7) = "Str='Received'"
            End If
            .Formulas(9) = "IsShowFormNo='1'"
            .SelectionFormula = "{tmpgenTbl.Username}='" & gUserName & "'"
            .DataFiles(0) = App.Path & "\HIData\setup.mdb"
            gReportPath = App.Path & "\HiReports\"
            If ChkSummary = 1 Then  '--summary
               .ReportFileName = gReportPath & "rptTaxRptE2FormRecSumm.rpt"
            Else
               .ReportFileName = gReportPath & "rptTaxRptE2FormRec.rpt"
            End If
        Case "E2 From Issuable" '-- "C Form Receivable"
            GProcFormIssueReceive Format(dtpFrDt.Value, "dd/MM/yyyy"), Format(dtpToDt.Value, "dd/MM/yyyy"), "E2I", txtMill, txtParty, txtBroker, ChkConsolidate, IIf(Opt(0) = True, "All", IIf(Opt(1) = True, "Pending", "Receive"))
            GProcCrystalRptPreparation CrystalReport1, IIf(Index = 0, "Window", "Printer")
            .Formulas(4) = "wFrToDt = 'From : ' & '" & Format(dtpFrDt.Value, "dd/MM/yyyy") & "' & '  To ' & '" & Format(dtpToDt.Value, "dd/MM/yyyy") & "'"
            If ChkConsolidate = 1 Then .Formulas(6) = "wAllFirmStr= 'All Firm'" '- wAllFirmStr
            .DataFiles(0) = App.Path & "\HIData\setup.mdb"
            .SelectionFormula = "{tmpgenTbl.Username}='" & gUserName & "'"
            gReportPath = App.Path & "\HiReports\"
            '-- Opt Str
            If Opt(0) = True Then
                '.Formulas(7) = "Str='All'"
            ElseIf Opt(1) = True Then
                .Formulas(7) = "Str='Pending'"
            Else
                .Formulas(7) = "Str='Received'"
            End If
            If OptBrk = True Then '-- Broker wise
                If ChkSummary = 1 Then '--summary
                    .ReportFileName = gReportPath & "rptTaxRptE2FormIssuSumm.rpt"
                Else '--detail
                    .Formulas(8) = "IsNewPage='" & ChkPartyOnNewPage & "'" '-- Partywise Print
                    '-for Received Show Form No
                    If Opt(2) = True Then
                        .Formulas(9) = "IsShowFormNo='1'"
                    Else
                        .Formulas(9) = "IsShowFormNo='0'"
                    End If
                    .ReportFileName = gReportPath & "rptTaxRptE2FormIssu.rpt"
                End If
            Else '-- Party wise
                If ChkSummary = 1 Then  '--summary
                    .ReportFileName = gReportPath & "rptTaxRptE2FormIssuPartywiseSumm.rpt"
                Else '--detail
                    .Formulas(8) = "IsNewPage='" & ChkPartyOnNewPage & "'" '-- Partywise Print
                    '-for Received Show Form No
                    If Opt(2) = True Then
                        .Formulas(9) = "IsShowFormNo='1'"
                    Else
                        .Formulas(9) = "IsShowFormNo='0'"
                    End If
                    .ReportFileName = gReportPath & "rptTaxRptE2FormIssuPartywise.rpt"
                End If
            End If
        Case "Requirement Of C Form", "Requirment of C Form Invoicewise" '-- "Requirement Of C Form"
            GProcFormIssueReceive Format(dtpFrDt.Value, "dd/MM/yyyy"), Format(dtpToDt.Value, "dd/MM/yyyy"), "CI", "", "", "", ChkConsolidate, IIf(Opt(0) = True, "All", IIf(Opt(1) = True, "Pending", "Receive"))
            GProcCrystalRptPreparation CrystalReport1, IIf(Index = 0, "Window", "Printer")
            .Formulas(4) = "wFrToDt = '" & Format(dtpFrDt.Value, "dd/MM/yy") & "' & '-' & '" & Format(dtpToDt.Value, "dd/MM/yy") & "'"
            If ChkConsolidate = 1 Then .Formulas(6) = "wAllFirmStr= 'All Firm'" '- wAllFirmStr
            .Formulas(7) = "wCSTNo = '" & gCCST & "'"
            .Formulas(8) = "wBSTNo = '" & gCBST & "'"
            .SelectionFormula = "{tmpgenTbl.Username}='" & gUserName & "'"
            .DataFiles(0) = App.Path & "\HIData\setup.mdb"
            gReportPath = App.Path & "\HiReports\"
            If gSelectedMenu = "Requirment of C Form Invoicewise" Then
                .ReportFileName = gReportPath & "rptTaxRptCFormRequireNew.rpt"
            Else
                .ReportFileName = gReportPath & "rptTaxRptCFormRequire.rpt"
            End If
'            .Action = 1
'            '--- Max (billAmt)  Party Detail
'            gSetupCn.BeginTrans
 '           gSetupCn.Execute "delete from tmpgenTbl where Amt1<( select max(amt1) from tmpgentbl where username='" & gUserName & "') and Username='" & gUserName & "'"
 '           gSetupCn.CommitTrans
 '           .DataFiles(0) = App.Path & "\HIData\setup.mdb"
 '           .SelectionFormula = "{tmpgenTbl.Username}='" & gUserName & "'"
 '           .WindowTitle = "Bill Of Higest Value"
 '           .ReportFileName = gReportPath & "rptTaxRptCFormRequirePartyDet.rpt"
 
         Case "Requirment of E2 Form", "Requirment of E2 Form Invoicewise" '-- "Requirement Of E2 Form"
            GProcFormIssueReceive Format(dtpFrDt.Value, "dd/MM/yyyy"), Format(dtpToDt.Value, "dd/MM/yyyy"), "E2I", "", txtParty, "", ChkConsolidate, IIf(Opt(0) = True, "All", IIf(Opt(1) = True, "Pending", "Receive"))
            GProcCrystalRptPreparation CrystalReport1, IIf(Index = 0, "Window", "Printer")
            .Formulas(4) = "wFrToDt = '" & Format(dtpFrDt.Value, "dd/MM/yy") & "' & '-' & '" & Format(dtpToDt.Value, "dd/MM/yy") & "'"
            If ChkConsolidate = 1 Then .Formulas(6) = "wAllFirmStr= 'All Firm'" '- wAllFirmStr
            .Formulas(7) = "wCSTNo = '" & gCCST & "'"
            .Formulas(8) = "wBSTNo = '" & gCBST & "'"
            .SelectionFormula = "{tmpgenTbl.Username}='" & gUserName & "'"
            .DataFiles(0) = App.Path & "\HIData\setup.mdb"
            gReportPath = App.Path & "\HiReports\"
            If gSelectedMenu = "Requirment of E2 Form Invoicewise" Then
               .ReportFileName = gReportPath & "rptTaxRptE2FormRequireInvoicewise.rpt"
            Else
               .ReportFileName = gReportPath & "rptTaxRptE2FormRequire.rpt"
            End If

 
        Case "Requirement Of F Form", "Requirment of F Form Invoicewise" '-- Requirement Of F Form
            GProcFormIssueReceive Format(dtpFrDt.Value, "dd/MM/yyyy"), Format(dtpToDt.Value, "dd/MM/yyyy"), "FR", "", "", "", ChkConsolidate
            GProcCrystalRptPreparation CrystalReport1, IIf(Index = 0, "Window", "Printer")
            .Formulas(4) = "wFrToDt = '" & Format(dtpFrDt.Value, "dd/MM/yy") & "' & '-' & '" & Format(dtpToDt.Value, "dd/MM/yy") & "'"
            If ChkConsolidate = 1 Then .Formulas(6) = "wAllFirmStr= 'All Firm'" '- wAllFirmStr
            .Formulas(7) = "wCSTNo = '" & gCCST & "'"
            .Formulas(8) = "wBSTNo = '" & gCBST & "'"
            .SelectionFormula = "{tmpgenTbl.Username}='" & gUserName & "'"
            .DataFiles(0) = App.Path & "\HIData\setup.mdb"
            gReportPath = App.Path & "\HiReports\"
            If gSelectedMenu = "Requirment of F Form Invoicewise" Then
               .ReportFileName = gReportPath & "rptTaxRptFFormRequireNew.rpt"
            Else
               .ReportFileName = gReportPath & "rptTaxRptFFormRequire.rpt"
            End If
'            .Action = 1
'            '--- Max (billAmt)  Party Detail
'            gSetupCn.BeginTrans
'            gSetupCn.Execute "delete from tmpgenTbl where Amt1<( select max(amt1) from tmpgentbl where username='" & gUserName & "') and Username='" & gUserName & "'"
'            gSetupCn.CommitTrans
'            .DataFiles(0) = App.Path & "\HIData\setup.mdb"
'            .SelectionFormula = "{tmpgenTbl.Username}='" & gUserName & "'"
 '           .WindowTitle = "Bill Of Higest Value"
 '           .ReportFileName = gReportPath & "rptTaxRptFFormRequirePartyDet.rpt"
        Case "F Form Statement" '-- "F Form Statement"
            .Formulas(4) = "wFrToDt='From : '& '" & dtpFrDt.Value & "' & ' To ' & '" & dtpToDt.Value & "'"
            .Formulas(5) = "FFormNo='" & txt(5) & "'"
            '----On Despatch Date
            X1 = "{tblAddLess.VType}='PI' and {tblAddLess.ADGPDT} in Date(" & Format(dtpFrDt.Value, "yyyy,mm,dd") & ") to Date(" & Format(dtpToDt.Value, "yyyy,mm,dd") & ")  " '
            If txtMill <> "" Then '-mill
                X1 = X1 + " and {tblMastAccount.AcName} = '" & txtMill & "'"
            End If
            .SelectionFormula = X1
            .ReportFileName = gReportPath & "rptTaxRptFFormSTM.rpt"
        Case "Pending C Form Receivable Slab Wise" '--"Pending C Form Receivable Slab Wise"
            GProcFormIssueReceive Format(dtpFrDt.Value, "dd/MM/yyyy"), Format(dtpToDt.Value, "dd/MM/yyyy"), "CRSlabWise", txtMill, txtParty, txtBroker, ChkConsolidate
            GProcCrystalRptPreparation CrystalReport1, IIf(Index = 0, "Window", "Printer")
            .Formulas(4) = "wFrToDt = 'From : ' & '" & Format(dtpFrDt.Value, "dd/MM/yyyy") & "' & '  To ' & '" & Format(dtpToDt.Value, "dd/MM/yyyy") & "'"
            If ChkConsolidate = 1 Then .Formulas(6) = "wAllFirmStr= 'All Firm'" '- wAllFirmStr
            .DataFiles(0) = App.Path & "\HIData\setup.mdb"
            .SelectionFormula = "{tmpgenTbl.Username}='" & gUserName & "'"
            gReportPath = App.Path & "\HiReports\"
            .ReportFileName = gReportPath & "rptTaxRptCFormRecPendingSlabSumm.rpt"
       Case "Pending E1 Form Receivable Slab Wise" '--"Pending E1 Form Receivable Slab Wise"
            GProcFormIssueReceive Format(dtpFrDt.Value, "dd/MM/yyyy"), Format(dtpToDt.Value, "dd/MM/yyyy"), "E1RSlabWise", txtMill, "", "", ChkConsolidate
            GProcCrystalRptPreparation CrystalReport1, IIf(Index = 0, "Window", "Printer")
            .Formulas(4) = "wFrToDt = 'From : ' & '" & Format(dtpFrDt.Value, "dd/MM/yyyy") & "' & '  To ' & '" & Format(dtpToDt.Value, "dd/MM/yyyy") & "'"
            If ChkConsolidate = 1 Then .Formulas(6) = "wAllFirmStr= 'All Firm'" '- wAllFirmStr
            .DataFiles(0) = App.Path & "\HIData\setup.mdb"
            .SelectionFormula = "{tmpgenTbl.Username}='" & gUserName & "'"
            gReportPath = App.Path & "\HiReports\"
            .ReportFileName = gReportPath & "rptTaxRptE1FormRecPendingSlabSumm.rpt"
        Case "Transactionwise Sales Details"
             X1 = "{tblAddLess.VDt} in Date(" & Format(dtpFrDt.Value, "yyyy,mm,dd") & ") to Date(" & Format(dtpToDt.Value, "yyyy,mm,dd") & ")  " '
             X1 = X1 + " and ({tblAddLess.VType}='ST' or {tblAddLess.VType}='SY' or {tblAddLess.VType}='SD' or {tblAddLess.VType}='SO' or {tblAddLess.VType}='RY' or {tblAddLess.VType}='RO' )"
             .SelectionFormula = X1
             .Formulas(4) = "wFrToDt = 'From : ' & '" & Format(dtpFrDt.Value, "dd/MM/yyyy") & "' & '  To ' & '" & Format(dtpToDt.Value, "dd/MM/yyyy") & "'"
             .ReportFileName = gReportPath & "rptTransactionwiseSales.rpt"
        Case "Transactionwise Purchase Details"
             X1 = "{tblAddLess.AdGpDt} in Date(" & Format(dtpFrDt.Value, "yyyy,mm,dd") & ") to Date(" & Format(dtpToDt.Value, "yyyy,mm,dd") & ")  " '
             X1 = X1 + " and ({tblAddLess.VType}='PY' or {tblAddLess.VType}='PT' or {tblAddLess.VType}='PI' or {tblAddLess.VType}='PO' or {tblAddLess.VType}='VY' or  {tblAddLess.VType}='VI') "
             .SelectionFormula = X1
             .Formulas(4) = "wFrToDt = 'From : ' & '" & Format(dtpFrDt.Value, "dd/MM/yyyy") & "' & '  To ' & '" & Format(dtpToDt.Value, "dd/MM/yyyy") & "'"
             .ReportFileName = gReportPath & "rptTransactionwisePurch.rpt"
            
            '================ Tax Report End
            '        Case "RCREC", "RCISS", "RE1REC", "RCFORM", "REFORM", "RFSTATE"
'            .Formulas(4) = "wFrToDt = 'From : ' & '" & Format(dtpFrDt.Value, "dd/MM/yyyy") & "' & '  To ' & '" & Format(dtpToDt.Value, "dd/MM/yyyy") & "'"
'            X1 = "{tblAddLess.VDt} in Date(" & Format(dtpFrDt.Value, "yyyy,mm,dd") & ") to Date(" & Format(dtpToDt.Value, "yyyy,mm,dd") & ")  " '
'            X1 = X1 '+ " and {tblAddLess.VType}='" & Left(gSelectedMenu, 2) & "'"
'            .SelectionFormula = X1
'            .ReportFileName = gReportPath & "rptTradeSaleRegiMon.rpt"
        '================ General Report -Sales
        Case "Total Sales Summary"  '-Total Sales"
            GProcSalesSummary Format(dtpFrDt.Value, "dd/MM/yyyy"), Format(dtpToDt.Value, "dd/MM/yyyy"), ChkConsolidate, RptOptStr, "Total"
            .Formulas(4) = "wFrToDt = 'From : ' & '" & Format(dtpFrDt.Value, "dd/MM/yyyy") & "' & '  To ' & '" & Format(dtpToDt.Value, "dd/MM/yyyy") & "'"
            If ChkConsolidate = 1 Then .Formulas(6) = "wAllFirmStr= 'All Firm'" '- wAllFirmStr
            '-- Opt Str
            .Formulas(7) = "Str='" & IIf(OptGrossAmt = True, "Gross Amount", IIf(OptNetAmt = True, "Net Amount", "Bags")) & "'"
            .DataFiles(0) = App.Path & "\HIData\setup.mdb"
            .SelectionFormula = "{tmpgenTbl.Username}='" & gUserName & "'"
            gReportPath = App.Path & "\HiReports\"
            .ReportFileName = gReportPath & "rptSalesSummTotal.rpt"
        Case "Countwise Sales Summary" '-- "Countwise Sales Summary"
            GProcSalesSummary Format(dtpFrDt.Value, "dd/MM/yyyy"), Format(dtpToDt.Value, "dd/MM/yyyy"), ChkConsolidate, RptOptStr, "CountWise", Item:=txtItem
            .Formulas(4) = "wFrToDt = 'From : ' & '" & Format(dtpFrDt.Value, "dd/MM/yyyy") & "' & '  To ' & '" & Format(dtpToDt.Value, "dd/MM/yyyy") & "'"
            If ChkConsolidate = 1 Then .Formulas(6) = "wAllFirmStr= 'All Firm'" '- wAllFirmStr
            '-- Opt Str
            .Formulas(7) = "Str='" & IIf(OptGrossAmt = True, "Gross Amount", IIf(OptNetAmt = True, "Net Amount", "Bags")) & "'"
            .DataFiles(0) = App.Path & "\HIData\setup.mdb"
            .SelectionFormula = "{tmpgenTbl.Username}='" & gUserName & "'"
            gReportPath = App.Path & "\HiReports\"
            .ReportFileName = gReportPath & "rptSalesSummCountwise.rpt"
        Case "Millwise Sales Summary" '--"Millwise Sales Summary"
            GProcSalesSummary Format(dtpFrDt.Value, "dd/MM/yyyy"), Format(dtpToDt.Value, "dd/MM/yyyy"), ChkConsolidate, RptOptStr, "MillWise", txtMill
            .Formulas(4) = "wFrToDt = 'From : ' & '" & Format(dtpFrDt.Value, "dd/MM/yyyy") & "' & '  To ' & '" & Format(dtpToDt.Value, "dd/MM/yyyy") & "'"
            If ChkConsolidate = 1 Then .Formulas(6) = "wAllFirmStr= 'All Firm'" '- wAllFirmStr
            '-- Opt Str
            .Formulas(7) = "Str='" & IIf(OptGrossAmt = True, "Gross Amount", IIf(OptNetAmt = True, "Net Amount", "Bags")) & "'"
            .DataFiles(0) = App.Path & "\HIData\setup.mdb"
            .SelectionFormula = "{tmpgenTbl.Username}='" & gUserName & "'"
            gReportPath = App.Path & "\HiReports\"
            .ReportFileName = gReportPath & "rptSalesSummMillwise.rpt"
        Case "Mill and Countwise Sales Summary" '-- "Mill and Countwise Sales Summary"
            GProcSalesSummary Format(dtpFrDt.Value, "dd/MM/yyyy"), Format(dtpToDt.Value, "dd/MM/yyyy"), ChkConsolidate, RptOptStr, "MillCountWise", txtMill, Item:=txtItem
            .Formulas(4) = "wFrToDt = 'From : ' & '" & Format(dtpFrDt.Value, "dd/MM/yyyy") & "' & '  To ' & '" & Format(dtpToDt.Value, "dd/MM/yyyy") & "'"
            If ChkConsolidate = 1 Then .Formulas(6) = "wAllFirmStr= 'All Firm'" '- wAllFirmStr
            '-- Opt Str
            .Formulas(7) = "Str='" & IIf(OptGrossAmt = True, "Gross Amount", IIf(OptNetAmt = True, "Net Amount", "Bags")) & "'"
            .DataFiles(0) = App.Path & "\HIData\setup.mdb"
            .SelectionFormula = "{tmpgenTbl.Username}='" & gUserName & "'"
            gReportPath = App.Path & "\HiReports\"
            .ReportFileName = gReportPath & "rptSalesSummMillCountwise.rpt"
        Case "Broker and Partywise Sales Summary" '--"Broker and Partywise Sales Summary"
            GProcSalesSummary Format(dtpFrDt.Value, "dd/MM/yyyy"), Format(dtpToDt.Value, "dd/MM/yyyy"), ChkConsolidate, RptOptStr, "BrokerPartyWise", Broker:=txtBroker, Party:=txtParty
            .Formulas(4) = "wFrToDt = 'From : ' & '" & Format(dtpFrDt.Value, "dd/MM/yyyy") & "' & '  To ' & '" & Format(dtpToDt.Value, "dd/MM/yyyy") & "'"
            If ChkConsolidate = 1 Then .Formulas(6) = "wAllFirmStr= 'All Firm'" '- wAllFirmStr
            '-- Opt Str
            .Formulas(7) = "Str='" & IIf(OptGrossAmt = True, "Gross Amount", IIf(OptNetAmt = True, "Net Amount", "Bags")) & "'"
            .DataFiles(0) = App.Path & "\HIData\setup.mdb"
            .SelectionFormula = "{tmpgenTbl.Username}='" & gUserName & "'"
            gReportPath = App.Path & "\HiReports\"
            .ReportFileName = gReportPath & "rptSalesSummBrokerPartywise.rpt"
        Case "Broker and Countwise Sales Summary" '--"Broker and Countwise Sales Summary
            GProcSalesSummary Format(dtpFrDt.Value, "dd/MM/yyyy"), Format(dtpToDt.Value, "dd/MM/yyyy"), ChkConsolidate, RptOptStr, "BrokerCountWise", Broker:=txtBroker, Item:=txtItem
            .Formulas(4) = "wFrToDt = 'From : ' & '" & Format(dtpFrDt.Value, "dd/MM/yyyy") & "' & '  To ' & '" & Format(dtpToDt.Value, "dd/MM/yyyy") & "'"
            If ChkConsolidate = 1 Then .Formulas(6) = "wAllFirmStr= 'All Firm'" '- wAllFirmStr
            '-- Opt Str
            .Formulas(7) = "Str='" & IIf(OptGrossAmt = True, "Gross Amount", IIf(OptNetAmt = True, "Net Amount", "Bags")) & "'"
            .DataFiles(0) = App.Path & "\HIData\setup.mdb"
            .SelectionFormula = "{tmpgenTbl.Username}='" & gUserName & "'"
            gReportPath = App.Path & "\HiReports\"
            .ReportFileName = gReportPath & "rptSalesSummBrokerCountwise.rpt"
        Case "Partywise Sales Summary" '--  "Partywise Sales Summary"
            GProcSalesSummary Format(dtpFrDt.Value, "dd/MM/yyyy"), Format(dtpToDt.Value, "dd/MM/yyyy"), ChkConsolidate, RptOptStr, "PartyWise", "", txtParty
            .Formulas(4) = "wFrToDt = 'From : ' & '" & Format(dtpFrDt.Value, "dd/MM/yyyy") & "' & '  To ' & '" & Format(dtpToDt.Value, "dd/MM/yyyy") & "'"
            If ChkConsolidate = 1 Then .Formulas(6) = "wAllFirmStr= 'All Firm'" '- wAllFirmStr
            '-- Opt Str
            .Formulas(7) = "Str='" & IIf(OptGrossAmt = True, "Gross Amount", IIf(OptNetAmt = True, "Net Amount", "Bags")) & "'"
            .DataFiles(0) = App.Path & "\HIData\setup.mdb"
            .SelectionFormula = "{tmpgenTbl.Username}='" & gUserName & "'"
            gReportPath = App.Path & "\HiReports\"
            .ReportFileName = gReportPath & "rptSalesSummPartywise.rpt"
        Case "Party and Countwise Sales Summary" '-- "Party and Countwise Sales Summary"
            GProcSalesSummary Format(dtpFrDt.Value, "dd/MM/yyyy"), Format(dtpToDt.Value, "dd/MM/yyyy"), ChkConsolidate, RptOptStr, "PartyCountWise", Party:=txtParty, Item:=txtItem
            .Formulas(4) = "wFrToDt = 'From : ' & '" & Format(dtpFrDt.Value, "dd/MM/yyyy") & "' & '  To ' & '" & Format(dtpToDt.Value, "dd/MM/yyyy") & "'"
            If ChkConsolidate = 1 Then .Formulas(6) = "wAllFirmStr= 'All Firm'" '- wAllFirmStr
            '-- Opt Str
            .Formulas(7) = "Str='" & IIf(OptGrossAmt = True, "Gross Amount", IIf(OptNetAmt = True, "Net Amount", "Bags")) & "'"
            .DataFiles(0) = App.Path & "\HIData\setup.mdb"
            .SelectionFormula = "{tmpgenTbl.Username}='" & gUserName & "'"
            gReportPath = App.Path & "\HiReports\"
            .ReportFileName = gReportPath & "rptSalesSummPartyCountwise.rpt"
        Case "Monthwise Sales Summary" '-- "Monthwise Sales Summary"
            GProcSalesSummary Format(dtpFrDt.Value, "dd/MM/yyyy"), Format(dtpToDt.Value, "dd/MM/yyyy"), ChkConsolidate, RptOptStr, "MonthWise"
            .Formulas(4) = "wFrToDt = 'From : ' & '" & Format(dtpFrDt.Value, "dd/MM/yyyy") & "' & '  To ' & '" & Format(dtpToDt.Value, "dd/MM/yyyy") & "'"
            If ChkConsolidate = 1 Then .Formulas(6) = "wAllFirmStr= 'All Firm'" '- wAllFirmStr
            '-- Opt Str
            .Formulas(7) = "Str='" & IIf(OptGrossAmt = True, "Gross Amount", IIf(OptNetAmt = True, "Net Amount", "Bags")) & "'"
            .DataFiles(0) = App.Path & "\HIData\setup.mdb"
            .SelectionFormula = "{tmpgenTbl.Username}='" & gUserName & "'"
            gReportPath = App.Path & "\HiReports\"
            .ReportFileName = gReportPath & "rptSalesSummMonthwise.rpt"
        Case "Monthwise Purchase Summary"
            GProcPurchaseSummary Format(dtpFrDt.Value, "dd/MM/yyyy"), Format(dtpToDt.Value, "dd/MM/yyyy"), ChkConsolidate, RptOptStr, "MonthWisePurchase"
            .Formulas(4) = "wFrToDt = 'From : ' & '" & Format(dtpFrDt.Value, "dd/MM/yyyy") & "' & '  To ' & '" & Format(dtpToDt.Value, "dd/MM/yyyy") & "'"
            If ChkConsolidate = 1 Then .Formulas(6) = "wAllFirmStr= 'All Firm'" '- wAllFirmStr
            '-- Opt Str
            .Formulas(7) = "Str='" & IIf(OptGrossAmt = True, "Gross Amount", IIf(OptNetAmt = True, "Net Amount", "Bags")) & "'"
            .DataFiles(0) = App.Path & "\HIData\setup.mdb"
            .SelectionFormula = "{tmpgenTbl.Username}='" & gUserName & "'"
            gReportPath = App.Path & "\HiReports\"
            .ReportFileName = gReportPath & "rptPurchaseSummMonthwise.rpt"
        Case "Monthwise Sales/Purchase New" '-- "Monthwise Sales/Purchase Summary"
             If ChkSummary.Value = 1 Then
                X1 = "('SY')"
                .Formulas(6) = "RpNm='Monthwise Sales Summery'"
             ElseIf ChkSumTin = 1 Then
                X1 = "('PO') and left(tblAddless.adtmp5,3) = 'RCM' "
                .Formulas(6) = "RpNm='Monthwise RCM Sales Summery'"
             Else
                X1 = "('PY')"
                .Formulas(6) = "RpNm='Monthwise Purchase Summery'"
             End If
              X = "select year(tblAddless.vdt),month(tblAddless.vdt),format(tblAddless.vdt,'mmmm'),sum( iif(VSubCtrNo=1, AdBillAmt,0)),sum(iif(AdAssVal > 0 ,iif(tblMastNarration.Narration = 'V A T',AdAssVal,0) ,iif(VSubCtrNo=1, AdSubAmt,0))),sum(tblVouSub.VsubCGSTAmt),sum(tblVouSub.VsubSGSTAmt),sum(tblVouSub.VsubIGSTAmt)," _
                   & " sum(iif(tblMastItem.ItOtherItem = 0,VSUBBAG,0)),sum(iif(tblMastItem.ItOtherItem = 0,VSubWt,0)),sum(iif(tblMastNarration.Narration = 'V A T',AdTaxAmt,0)),sum(iif(VSubCtrNo=1,AdRoff+AdAftTaxAdd+AdAftTaxLess,0)),'" & gUserName & "' from tblVousub,tblMastItem,tblAddLess,tblMastNarration where   tblVouSub.VSubItCode = tblMastItem.Itcode   and  tblAddLess.Vno = tblVouSub.Vno  and  tblAddLess.Vtype = tblVouSub.Vtype and " _
                   & " tblAddLess.Vyear = tblVouSub.Vyear and tblAddLess.vtype in " & X1 & " and tbladdless.VDt>= cdate('" & Format(dtpFrDt.Value, "dd/MM/yyyy") & "')  and tbladdless.VDt<= cdate('" & Format(dtpToDt.Value, "dd/MM/yyyy") & "')" _
                   & " and tblAddLess.AdTaxCode = tblMastNarration.NarrCode    group by year(tblAddless.vdt),month(tblAddless.vdt),format(tblAddless.vdt,'mmmm')"
             GProcInsertToSetupDbTbl ChkConsolidate, X, "tmpGenTbl", "Vyear,Vno,Nar1,Amt10,Amt6,Amt14,Amt15,Amt16,Amt4,Amt5,Amt9,Amt11,username", "username='" & gUserName & "'", "Nar5"
             .Formulas(4) = "wFrToDt = 'From : ' & '" & Format(dtpFrDt.Value, "dd/MM/yyyy") & "' & '  To ' & '" & Format(dtpToDt.Value, "dd/MM/yyyy") & "'"
             If ChkConsolidate = 1 Then .Formulas(6) = "wAllFirmStr= 'All Firm'" '- wAllFirmStr
             '-- Opt Str
             .Formulas(7) = "Str='" & IIf(OptGrossAmt = True, "Gross Amount", IIf(OptNetAmt = True, "Net Amount", "Bags")) & "'"
             .DataFiles(0) = App.Path & "\HIData\setup.mdb"
             .SelectionFormula = "{tmpgenTbl.Username}='" & gUserName & "'"
             gReportPath = App.Path & "\HiReports\"
             .ReportFileName = gReportPath & "rptPurchSummMonthwiseNew.rpt" '"rptPurchaseSalesSummMonthwise.rpt"
            
         Case "TDS Register", "TDS Summary"     '- "TDS Register", "TDS Summary"
            Dim wPartyCd As Long
            Dim wTdsAcCd As Long
            '--- For Interest A/c
            X = "SELECT 'I',TBLVOUCHER.VNO, TBLVOUCHER.VDT, TBLVOUCHER.VTYPE,abs(TBLVOUCHER.VAMT),TBLMASTACCOUNT.ACNAME, iif(isnull(TBLRECVSSALE.INTRECEIVED),0,TBLRECVSSALE.INTRECEIVED), iif(isnull(TBLRECVSSALE.LESSTDS),0,TBLRECVSSALE.LESSTDS), iif(isnull(TBLRECVSSALE.GSTDBAMT),0,TBLRECVSSALE.GSTDBAMT), TBLRECVSSALE.INVBILLNO,TBLRECVSSALE.INVDT,'" & gUserName & "' " _
                & " FROM (TBLVOUCHER INNER JOIN TBLMASTACCOUNT ON TBLVOUCHER.VAcOCode = TBLMASTACCOUNT.AcCode) LEFT JOIN TBLRECVSSALE ON (TBLVOUCHER.VDt = TBLRECVSSALE.RecVDt) AND (TBLVOUCHER.VAcOCode = TBLRECVSSALE.PartyCode) AND (TBLVOUCHER.VNo = TBLRECVSSALE.JvNoInt) " _
                & " WHERE TBLVOUCHER.VACCODE = " & wMillCode & " and TBLVOUCHER.VDt>= cdate('" & dtpFrDt.Value & "')  and TBLVOUCHER.VDt<= cdate('" & dtpToDt.Value & "') "
            
            If txtParty <> "" Then
               wPartyCd = GProcGetColumnValue("tblMastAccount", "AcName", txtParty, "S", "AcCode", "N")
               X = X + " and tblVoucher.Vacocode = " & wPartyCd & ""
            End If
            gCn.BeginTrans
            gCn.Execute "Delete from tmpGenTbl where  USERNAME = '" & gUserName & "'"
            gCn.Execute " insert into tmpGenTbl (VBILLNO,Vno,Vdt,Vtype,Amt1,Nar1,Amt2,Amt3,Amt4,Nar2,VBILLDT,USERNAME) " & X
            
            '---For TDS A/c
            If txtBroker <> "" Then
               wTdsAcCd = GProcGetColumnValue("tblMastAccount", "AcName", txtBroker, "S", "AcCode", "N")
            
                X = "SELECT 'T',TBLVOUCHER.VNO, TBLVOUCHER.VDT, TBLVOUCHER.VTYPE,abs(TBLVOUCHER.VAMT),TBLMASTACCOUNT.ACNAME,Vnar1,Vnar2,'" & gUserName & "' " _
                    & " FROM TBLVOUCHER,TBLMASTACCOUNT WHERE TBLVOUCHER.VACOCODE = tblmastAccount.accode and TBLVOUCHER.VACCODE = " & wTdsAcCd & "  and  TBLVOUCHER.VDt>= cdate('" & dtpFrDt.Value & "')  and TBLVOUCHER.VDt<= cdate('" & dtpToDt.Value & "') "
                If txtParty <> "" Then
                    X = X + " and TBLVOUCHER.VACOCODE = " & wPartyCd & ""
                End If
                gCn.Execute " insert into tmpGenTbl (VBILLNO,Vno,Vdt,Vtype,Amt1,Nar1,Nar3,Nar4,USERNAME) " & X
            End If
            
            
            gCn.CommitTrans
            .Formulas(4) = "wFrToDt = 'From : ' & '" & Format(dtpFrDt.Value, "dd/MM/yyyy") & "' & '  To ' & '" & Format(dtpToDt.Value, "dd/MM/yyyy") & "'"
            .Formulas(5) = "CPanNo = '" & gCPAN & "'" '-Pan No
            .Formulas(6) = "CTDSNo = '" & gCTDS & "'" '-TDS No
            .Formulas(7) = "CTDSCircle= '" & gCTDSCircle & "'" '-TDS  circle
            .Formulas(8) = "CDesignation = '" & gCDesignation & "'" '- Designation
            .SelectionFormula = "{tmpgenTbl.Username}='" & gUserName & "'"
            If ChkSummary.Value = 1 Then
               .ReportFileName = gReportPath & "RptTaxTDSRegSum.rpt"
            Else
               .ReportFileName = gReportPath & "RptTaxTDSReg.rpt"
            End If
         '------------ General Report Purchase
         Case "Total Purchase Summary"  '-"Total Purchase Summary"
            GProcPurchaseSummary Format(dtpFrDt.Value, "dd/MM/yyyy"), Format(dtpToDt.Value, "dd/MM/yyyy"), ChkConsolidate, RptOptStr, "Total"
            .Formulas(4) = "wFrToDt = 'From : ' & '" & Format(dtpFrDt.Value, "dd/MM/yyyy") & "' & '  To ' & '" & Format(dtpToDt.Value, "dd/MM/yyyy") & "'"
            If ChkConsolidate = 1 Then .Formulas(6) = "wAllFirmStr= 'All Firm'" '- wAllFirmStr
            '-- Opt Str
            .Formulas(7) = "Str='" & IIf(OptGrossAmt = True, "Gross Amount", IIf(OptNetAmt = True, "Net Amount", "Bags")) & "'"
            .DataFiles(0) = App.Path & "\HIData\setup.mdb"
            .SelectionFormula = "{tmpgenTbl.Username}='" & gUserName & "'"
            gReportPath = App.Path & "\HiReports\"
            .ReportFileName = gReportPath & "rptPurchaseSummTotal.rpt"
        Case "Countwise Purchase Summary" '-- "Countwise Purchase Summary"
            GProcPurchaseSummary Format(dtpFrDt.Value, "dd/MM/yyyy"), Format(dtpToDt.Value, "dd/MM/yyyy"), ChkConsolidate, RptOptStr, "CountWise", Item:=txtItem
            .Formulas(4) = "wFrToDt = 'From : ' & '" & Format(dtpFrDt.Value, "dd/MM/yyyy") & "' & '  To ' & '" & Format(dtpToDt.Value, "dd/MM/yyyy") & "'"
            If ChkConsolidate = 1 Then .Formulas(6) = "wAllFirmStr= 'All Firm'" '- wAllFirmStr
            '-- Opt Str
            .Formulas(7) = "Str='" & IIf(OptGrossAmt = True, "Gross Amount", IIf(OptNetAmt = True, "Net Amount", "Bags")) & "'"
            .DataFiles(0) = App.Path & "\HIData\setup.mdb"
            .SelectionFormula = "{tmpgenTbl.Username}='" & gUserName & "'"
            gReportPath = App.Path & "\HiReports\"
            .ReportFileName = gReportPath & "rptPurchaseSummCountwise.rpt"
        Case "Millwise Purchase Summary" '--"Millwise Purchase Summary"
            GProcPurchaseSummary Format(dtpFrDt.Value, "dd/MM/yyyy"), Format(dtpToDt.Value, "dd/MM/yyyy"), ChkConsolidate, RptOptStr, "MillWise", txtMill
            .Formulas(4) = "wFrToDt = 'From : ' & '" & Format(dtpFrDt.Value, "dd/MM/yyyy") & "' & '  To ' & '" & Format(dtpToDt.Value, "dd/MM/yyyy") & "'"
            If ChkConsolidate = 1 Then .Formulas(6) = "wAllFirmStr= 'All Firm'" '- wAllFirmStr
            '-- Opt Str
            .Formulas(7) = "Str='" & IIf(OptGrossAmt = True, "Gross Amount", IIf(OptNetAmt = True, "Net Amount", "Bags")) & "'"
            .DataFiles(0) = App.Path & "\HIData\setup.mdb"
            .SelectionFormula = "{tmpgenTbl.Username}='" & gUserName & "'"
            gReportPath = App.Path & "\HiReports\"
            .ReportFileName = gReportPath & "rptPurchaseSummMillwise.rpt"
        Case "Monthwise Sales/Purchase Summary" '-- "Monthwise Sales/Purchase Summary"
            GProcPurchaseSummary Format(dtpFrDt.Value, "dd/MM/yyyy"), Format(dtpToDt.Value, "dd/MM/yyyy"), ChkConsolidate, RptOptStr, "MonthWise"
            .Formulas(4) = "wFrToDt = 'From : ' & '" & Format(dtpFrDt.Value, "dd/MM/yyyy") & "' & '  To ' & '" & Format(dtpToDt.Value, "dd/MM/yyyy") & "'"
            If ChkConsolidate = 1 Then .Formulas(6) = "wAllFirmStr= 'All Firm'" '- wAllFirmStr
            '-- Opt Str
            .Formulas(7) = "Str='" & IIf(OptGrossAmt = True, "Gross Amount", IIf(OptNetAmt = True, "Net Amount", "Bags")) & "'"
            .DataFiles(0) = App.Path & "\HIData\setup.mdb"
            .SelectionFormula = "{tmpgenTbl.Username}='" & gUserName & "'"
            gReportPath = App.Path & "\HiReports\"
            .ReportFileName = gReportPath & "rptPurchaseSalesSummMonthwise.rpt"
        Case "Total Inward Summary"  '-"Total Inward Summary"
            GProcInwardSummary Format(dtpFrDt.Value, "dd/MM/yyyy"), Format(dtpToDt.Value, "dd/MM/yyyy"), ChkConsolidate, RptOptStr, "Total"
            .Formulas(4) = "wFrToDt = 'From : ' & '" & Format(dtpFrDt.Value, "dd/MM/yyyy") & "' & '  To ' & '" & Format(dtpToDt.Value, "dd/MM/yyyy") & "'"
            If ChkConsolidate = 1 Then .Formulas(6) = "wAllFirmStr= 'All Firm'" '- wAllFirmStr
            '-- Opt Str
            .Formulas(7) = "Str='" & IIf(OptGrossAmt = True, "Gross Amount", IIf(OptNetAmt = True, "Net Amount", "Bags")) & "'"
            .DataFiles(0) = App.Path & "\HIData\setup.mdb"
            .SelectionFormula = "{tmpgenTbl.Username}='" & gUserName & "'"
            gReportPath = App.Path & "\HiReports\"
            .ReportFileName = gReportPath & "rptInwardSummTotal.rpt"
        Case "Countwise Inward Summary" '-- "Countwise Inward Summary"
            GProcInwardSummary Format(dtpFrDt.Value, "dd/MM/yyyy"), Format(dtpToDt.Value, "dd/MM/yyyy"), ChkConsolidate, RptOptStr, "CountWise", Item:=txtItem
            .Formulas(4) = "wFrToDt = 'From : ' & '" & Format(dtpFrDt.Value, "dd/MM/yyyy") & "' & '  To ' & '" & Format(dtpToDt.Value, "dd/MM/yyyy") & "'"
            If ChkConsolidate = 1 Then .Formulas(6) = "wAllFirmStr= 'All Firm'" '- wAllFirmStr
            '-- Opt Str
            .Formulas(7) = "Str='" & IIf(OptGrossAmt = True, "Gross Amount", IIf(OptNetAmt = True, "Net Amount", "Bags")) & "'"
            .DataFiles(0) = App.Path & "\HIData\setup.mdb"
            .SelectionFormula = "{tmpgenTbl.Username}='" & gUserName & "'"
            gReportPath = App.Path & "\HiReports\"
            .ReportFileName = gReportPath & "rptInwardSummCountwise.rpt"
        Case "Millwise Inward Summary" '--"Millwise Inward Summary"
            GProcInwardSummary Format(dtpFrDt.Value, "dd/MM/yyyy"), Format(dtpToDt.Value, "dd/MM/yyyy"), ChkConsolidate, RptOptStr, "MillWise", txtMill
            .Formulas(4) = "wFrToDt = 'From : ' & '" & Format(dtpFrDt.Value, "dd/MM/yyyy") & "' & '  To ' & '" & Format(dtpToDt.Value, "dd/MM/yyyy") & "'"
            If ChkConsolidate = 1 Then .Formulas(6) = "wAllFirmStr= 'All Firm'" '- wAllFirmStr
            '-- Opt Str
            .Formulas(7) = "Str='" & IIf(OptGrossAmt = True, "Gross Amount", IIf(OptNetAmt = True, "Net Amount", "Bags")) & "'"
            .DataFiles(0) = App.Path & "\HIData\setup.mdb"
            .SelectionFormula = "{tmpgenTbl.Username}='" & gUserName & "'"
            gReportPath = App.Path & "\HiReports\"
            .ReportFileName = gReportPath & "rptInwardSummMillwise.rpt"
           '----------- Brokerge Report
         Case "Brokerage Detail Report" '--"Brokerage Detail Report"
            .Formulas(4) = "wFrToDt = 'From : ' & '" & Format(dtpFrDt.Value, "dd/MM/yyyy") & "' & '  To ' & '" & Format(dtpToDt.Value, "dd/MM/yyyy") & "'"
            X1 = "{tblAddLess.VDt} in Date(" & Format(dtpFrDt.Value, "yyyy,mm,dd") & ") to Date(" & Format(dtpToDt.Value, "yyyy,mm,dd") & ")  "
            If txtBroker <> "" Then X1 = X1 + " and {tblMastAccount_Brk.AcName}='" & txtBroker & "'"
            '-- Sale Type Str
            If OptBrkSAType(0) = True Then '--all
                .Formulas(7) = "Str=''"
                X1 = X1 + " and ({tblAddLess.Vtype}='SY' or {tblAddLess.Vtype}='SD' or {tblAddLess.Vtype}='SO' or {tblAddLess.Vtype}='SM' or {tblAddLess.Vtype}='ST' or {tblAddLess.Vtype}='RY') "
            ElseIf OptBrkSAType(1) = True Then '--trade
                .Formulas(7) = "Str='Trade Sales'"
                X1 = X1 + " and {tblAddLess.Vtype} = 'SY'"
            ElseIf OptBrkSAType(2) = True Then '--sit
                .Formulas(7) = "Str='SIT Sales'"
                X1 = X1 + " and {tblAddLess.Vtype} = 'ST'"
            ElseIf OptBrkSAType(3) = True Then '--consig
                .Formulas(7) = "Str='Consignment Sales'"
                X1 = X1 + " and {tblAddLess.Vtype} = 'SO'"
            ElseIf OptBrkSAType(4) = True Then '--Depot
                .Formulas(7) = "Str='Depot Sales'"
                X1 = X1 + " and {tblAddLess.Vtype} = 'SD'"
            Else '--Mill bill
                .Formulas(7) = "Str='Mill Bill Sales'"
                X1 = X1 + " and {tblAddLess.Vtype} = 'SM'"
            End If
            .SelectionFormula = X1 & " and {@BrkAmt}<>0"
            If ChkSummary.Value = 1 Then
               .ReportFileName = gReportPath & "rptBrokerageSumm.rpt"
            Else
               .ReportFileName = gReportPath & "rptBrokerageRptDetail.rpt"
            End If
            
         Case "Brokerage Total Report", "Brokerage Voucher Report" '--"Brokerage Total Report",'--"Brokerage Voucher Report"
            .Formulas(4) = "wFrToDt = 'From ' & '" & Format(dtpFrDt.Value, "dd/MM/yyyy") & "' & '  To ' & '" & Format(dtpToDt.Value, "dd/MM/yyyy") & "'"
            X1 = "{tblAddLess.VDt} in Date(" & Format(dtpFrDt.Value, "yyyy,mm,dd") & ") to Date(" & Format(dtpToDt.Value, "yyyy,mm,dd") & ")  "
            If txtBroker <> "" Then X1 = X1 + " and {tblMastAccount_Brk.AcName}='" & txtBroker & "'" '--Broker
            '-- Sale Type Str
            If OptBrkSAType(0) = True Then '--all
                .Formulas(7) = "Str=''"
                X1 = X1 + " and ({tblAddLess.Vtype}='SY' or {tblAddLess.Vtype}='SD' or {tblAddLess.Vtype}='SO' or {tblAddLess.Vtype}='SM' or {tblAddLess.Vtype}='ST' or {tblAddLess.Vtype}='RY')"
                PrevBrkStr = "(tblAddLess.Vtype in ('SY','SD','SO','SM','ST','RY'))"
            ElseIf OptBrkSAType(1) = True Then '--trade
                .Formulas(7) = "Str='Trade Sales'"
                X1 = X1 + " and {tblAddLess.Vtype} = 'SY'"
                PrevBrkStr = "(tblAddLess.Vtype in ('SY'))"
            ElseIf OptBrkSAType(2) = True Then '--sit
                .Formulas(7) = "Str='SIT Sales'"
                X1 = X1 + " and {tblAddLess.Vtype} = 'ST'"
                PrevBrkStr = "(tblAddLess.Vtype in ('ST'))"
            ElseIf OptBrkSAType(3) = True Then '--consig
                .Formulas(7) = "Str='Consignment Sales'"
                X1 = X1 + " and {tblAddLess.Vtype} = 'SO'"
                PrevBrkStr = "(tblAddLess.Vtype in ('SO'))"
            ElseIf OptBrkSAType(4) = True Then '--Depot
                .Formulas(7) = "Str='Depot Sales'"
                X1 = X1 + " and {tblAddLess.Vtype} = 'SD'"
                PrevBrkStr = "(tblAddLess.Vtype in ('SD'))"
            Else '--Mill bill
                .Formulas(7) = "Str='Mill Bill Sales'"
                X1 = X1 + " and {tblAddLess.Vtype} = 'SM'"
                PrevBrkStr = "(tblAddLess.Vtype in ('SM'))"
            End If
            '--- Pevious Brokerage
            gCn.BeginTrans
            gCn.Execute "delete from tmpGenTbl where Username='" & gUserName & "'"
            If gBackEndDB = gBackEndAccess Then
                PrevBrkStr = "tblAddLess.Vdt<cdate('" & Format(dtpFrDt.Value, "dd/MM/yyyy") & "') and " & PrevBrkStr
            Else
                PrevBrkStr = "tblAddLess.Vdt<To_date('" & dtpFrDt.Value & "', 'dd/MM/yyyy') and " & PrevBrkStr
            End If
            'If txtBroker <> "" Then PrevBrkStr = PrevBrkStr + " and tblMastAccount.AcName}='" & txtBroker & "'" '-Broker
            If gBackEndDB = gBackEndAccess Then
                X = "sum(format((iif(ITBROKON='Bag',VSUBBAG,iif(ITBROKON='Kg',VSUBWT/100,iif(tblAddLess.Vdt >= cdate('01/04/2023'),(AdSubAmt/100),(ADBILLAMT/100)))) * iif(tblAddLess.Vdt >= cdate('01/04/2023'),ITBROKRT,iif(itbrkrtold = 0,ITBROKRT,ITBROKRT ))),'###########')) "
            Else
                 X = "sum(Round(decode(ITBROKON,'Bag',VSUBBAG,decode(ITBROKON,'Kg',VSUBWT/100,ADBILLAMT/100)) * ITBROKRT))"
            End If
            X = " select AcCode,( select " & X & " from tblOutStanding,TblAddLess,TblVouSub,tblMastItem where OUTBROKERCODE =acCode " _
            & " and  VSUBITCODE =ItCode and tblAddLess.Vtype=TblVouSub.Vtype and  tblAddLess.Vno=TblVouSub.Vno and  tblAddLess.VYear=TblVouSub.VYear and  " _
            & " tblAddLess.Vtype=tblOutStanding.Vtype and  tblAddLess.Vno=tblOutStanding.Vno and  tblAddLess.VYear=tblOutStanding.VYear and  " _
            & PrevBrkStr & ") , " & IIf(gBackEndDB = gBackEndAccess, "cdate('" & Format(CDate(dtpFrDt.Value - 1), "dd/MM/yyyy") & "')", "to_date('" & CDate(dtpFrDt.Value) - 1 & "','dd/MM/yyyy')") & "," & txt(1) & "," & txt(2) & "," & txt(3) & "," & txt(4) & ",'" & gUserName & "'" _
            & " from tblmastaccount " & IIf(txtBroker <> "", " where AcName='" & txtBroker & "'", "") & " group by AcCode"
            gCn.Execute "insert into TmpGenTbl (VAccode,Amt1,Vdt,Amt2,Amt3,Amt4,Amt5,UserName) " & X
            gCn.Execute "update tmpgentbl set amt1=0 where amt1 is null and UserName='" & gUserName & "'"
            gCn.CommitTrans
            '----- Prev Brk End
            If gBackEndDB = gBackEndAccess Then
                .DataFiles(0) = gDbLocation
                gReportPath = App.Path & "\HiReports_Access\"
            Else '-- Oracle
                .Connect = "DSN=" + gODBCDSN + " ;UID=Scott;PWD=" + gODBCPwd + ";DBQ=" + gODBCHostStr
                gReportPath = App.Path & "\HiReports_ORA\"
            End If
            X1 = X1 + " and {tmpGenTbl.UserName}='" & gUserName & "' and {@Brkamt}<>0"
            .SelectionFormula = X1
            If gSelectedMenu = "Brokerage Total Report" Then '--Total Report
                .ReportFileName = gReportPath & "rptBrokerageRptSumm.rpt" '-Voucher Report
            Else
                .ReportFileName = gReportPath & "rptBrokerageRptVoucher.rpt" '-Voucher Report
            End If
''            .Formulas(4) = "wFrToDt = 'From : ' & '" & Format(dtpFrDt.Value, "dd/MM/yyyy") & "' & '  To ' & '" & Format(dtpToDt.Value, "dd/MM/yyyy") & "'"
''            X1 = "{tblAddLess.VDt} in Date(" & Format(dtpFrDt.Value, "yyyy,mm,dd") & ") to Date(" & Format(dtpToDt.Value, "yyyy,mm,dd") & ")  "
''            If txtBroker <> "" Then X1 = X1 + " and {tblMastAccount_Brk.AcName}='" & txtBroker & "'"
''            '-- Sale Type Str
''            If OptBrkSAType(0) = True Then '--all
''                .Formulas(7) = "Str=''"
''                X1 = X1 + " and ({tblAddLess.Vtype}='SY' or {tblAddLess.Vtype}='SD' or {tblAddLess.Vtype}='SO' or {tblAddLess.Vtype}='SM' or {tblAddLess.Vtype}='ST')"
''            ElseIf OptBrkSAType(1) = True Then '--trade
''                .Formulas(7) = "Str='Trade Sales'"
''                X1 = X1 + " and {tblAddLess.Vtype} = 'SY'"
''            ElseIf OptBrkSAType(2) = True Then '--sit
''                .Formulas(7) = "Str='SIT Sales'"
''                X1 = X1 + " and {tblAddLess.Vtype} = 'ST'"
''            ElseIf OptBrkSAType(3) = True Then '--consig
''                .Formulas(7) = "Str='Consignment Sales'"
''                X1 = X1 + " and {tblAddLess.Vtype} = 'SO'"
''            ElseIf OptBrkSAType(4) = True Then '--Depot
''                .Formulas(7) = "Str='Depot Sales'"
''                X1 = X1 + " and {tblAddLess.Vtype} = 'SD'"
''            Else '--Mill bill
''                .Formulas(7) = "Str='Mill Bill Sales'"
''                X1 = X1 + " and {tblAddLess.Vtype} = 'SM'"
''            End If
''            If ChkPartyOnNewPage.Value = 1 Then
''                .Formulas(9) = "RptOn='1'"
''            Else
''                .Formulas(9) = "RptOn='0'"
''            End If
''
''            .SelectionFormula = X1 & " and {@BrkAmt}>0"
''            If ChkSummary.Value = 1 Then
''               .ReportFileName = gReportPath & "rptBrokerageSumm.rpt"
''            Else
''               .ReportFileName = gReportPath & "rptBrokerageRptDetail.rpt"
''            End If
''         Case "Brokerage Total Report", "Brokerage Voucher Report" '--"Brokerage Total Report",'--"Brokerage Voucher Report"
''            .Formulas(4) = "wFrToDt = 'From ' & '" & Format(dtpFrDt.Value, "dd/MM/yyyy") & "' & '  To ' & '" & Format(dtpToDt.Value, "dd/MM/yyyy") & "'"
''            X1 = "{tblAddLess.VDt} in Date(" & Format(dtpFrDt.Value, "yyyy,mm,dd") & ") to Date(" & Format(dtpToDt.Value, "yyyy,mm,dd") & ")  "
''            If txtBroker <> "" Then X1 = X1 + " and {tblMastAccount_Brk.AcName}='" & txtBroker & "'" '--Broker
''            '-- Sale Type Str
''            If OptBrkSAType(0) = True Then '--all
''                .Formulas(7) = "Str=''"
''                X1 = X1 + " and ({tblAddLess.Vtype}='SY' or {tblAddLess.Vtype}='SD' or {tblAddLess.Vtype}='SO' or {tblAddLess.Vtype}='SM' or {tblAddLess.Vtype}='ST')"
''                PrevBrkStr = "(tblAddLess.Vtype in ('SY','SD','SO','SM','ST'))"
''            ElseIf OptBrkSAType(1) = True Then '--trade
''                .Formulas(7) = "Str='Trade Sales'"
''                X1 = X1 + " and {tblAddLess.Vtype} = 'SY'"
''                PrevBrkStr = "(tblAddLess.Vtype in ('SY'))"
''            ElseIf OptBrkSAType(2) = True Then '--sit
''                .Formulas(7) = "Str='SIT Sales'"
''                X1 = X1 + " and {tblAddLess.Vtype} = 'ST'"
''                PrevBrkStr = "(tblAddLess.Vtype in ('ST'))"
''            ElseIf OptBrkSAType(3) = True Then '--consig
''                .Formulas(7) = "Str='Consignment Sales'"
''                X1 = X1 + " and {tblAddLess.Vtype} = 'SO'"
''                PrevBrkStr = "(tblAddLess.Vtype in ('SO'))"
''            ElseIf OptBrkSAType(4) = True Then '--Depot
''                .Formulas(7) = "Str='Depot Sales'"
''                X1 = X1 + " and {tblAddLess.Vtype} = 'SD'"
''                PrevBrkStr = "(tblAddLess.Vtype in ('SD'))"
''            Else '--Mill bill
''                .Formulas(7) = "Str='Mill Bill Sales'"
''                X1 = X1 + " and {tblAddLess.Vtype} = 'SM'"
''                PrevBrkStr = "(tblAddLess.Vtype in ('SM'))"
''            End If
''            '--- Pevious Brokerage
''            gCn.BeginTrans
''            gCn.Execute "delete from tmpGenTbl where Username='" & gUserName & "'"
''            If gBackEndDB = gBackEndAccess Then
''                PrevBrkStr = "tblAddLess.Vdt<cdate('" & Format(dtpFrDt.Value, "dd/MM/yyyy") & "') and " & PrevBrkStr
''            Else
''                PrevBrkStr = "tblAddLess.Vdt<To_date('" & dtpFrDt.Value & "', 'dd/MM/yyyy') and " & PrevBrkStr
''            End If
''            'If txtBroker <> "" Then PrevBrkStr = PrevBrkStr + " and tblMastAccount.AcName}='" & txtBroker & "'" '-Broker
''            '(Brk Accode,Prev Brk amt,TDS Limit,TDS %,SCharge%,ECess%,UserName)
''            If gBackEndDB = gBackEndAccess Then
'''                X = "sum((iif(ITBROKON='Bag',VSUBBAG,iif(ITBROKON='Kg',VSUBWT/100,ADBILLAMT/100)) * ITBROKRT))"
''                X = "sum((iif(ITBROKON='Bag',VSUBBAG,iif(ITBROKON='Kg',VSUBWT/100,ADBILLAMT/100)) * iif(iif(datediff('d',tbladdless.vdt,ItOldRtBefore) <= 0,ItBrokRt,ItBrkRtOld) > 0,iif(datediff('d',tbladdless.vdt,ItOldRtBefore) <= 0,ItBrokRt,ItBrkRtOld),ITBROKRT) ))"
''            Else
''                X = "sum(Round(decode(ITBROKON,'Bag',VSUBBAG,decode(ITBROKON,'Kg',VSUBWT/100,ADBILLAMT/100)) * ITBROKRT))"
''            End If
''            X = " select AcCode,( select " & X & " from tblOutStanding,TblAddLess,TblVouSub,tblMastItem where OUTBROKERCODE =acCode " _
''            & " and  VSUBITCODE =ItCode and tblAddLess.Vtype=TblVouSub.Vtype and  tblAddLess.Vno=TblVouSub.Vno and  tblAddLess.VYear=TblVouSub.VYear and  " _
''            & " tblAddLess.Vtype=tblOutStanding.Vtype and  tblAddLess.Vno=tblOutStanding.Vno and  tblAddLess.VYear=tblOutStanding.VYear and  " _
''            & PrevBrkStr & ") , " & IIf(gBackEndDB = gBackEndAccess, "cdate('" & Format(CDate(dtpFrDt.Value - 1), "dd/MM/yyyy") & "')", "to_date('" & CDate(dtpFrDt.Value) - 1 & "','dd/MM/yyyy')") & "," & txt(1) & "," & txt(2) & "," & txt(3) & "," & txt(4) & ",'" & gUserName & "'" _
''            & " from tblmastaccount " & IIf(txtBroker <> "", " where AcName='" & txtBroker & "'", "") & " group by AcCode"
''            gCn.Execute "insert into TmpGenTbl (VAccode,Amt1,Vdt,Amt2,Amt3,Amt4,Amt5,UserName) " & X
''            gCn.Execute "update tmpgentbl set amt1=0 where amt1 is null and UserName='" & gUserName & "'"
''            gCn.CommitTrans
''            '----- Prev Brk End
''            If gBackEndDB = gBackEndAccess Then
''                .DataFiles(0) = gDbLocation
''                gReportPath = App.Path & "\HiReports_Access\"
''            Else '-- Oracle
''                .Connect = "DSN=" + gODBCDSN + " ;UID=Scott;PWD=" + gODBCPwd + ";DBQ=" + gODBCHostStr
''                gReportPath = App.Path & "\HiReports_ORA\"
''            End If
''            X1 = X1 + " and {tmpGenTbl.UserName}='" & gUserName & "' and {@Brkamt}>0"
''            .SelectionFormula = X1
''            If gSelectedMenu = "Brokerage Total Report" Then '--Total Report
''                If ChkPartyOnNewPage.Value = 1 Then
''                   .ReportFileName = gReportPath & "rptBrokerageRptVoucherMB.rpt"
''                Else
''                   .ReportFileName = gReportPath & "rptBrokerageRptSumm.rpt" '-Voucher Report
''                End If
''            Else
''                .ReportFileName = gReportPath & "rptBrokerageRptVoucher.rpt" '-Voucher Report
''            End If
         Case "Inward/Purchase Brokerage Report" '--"Inward/Purchase Brokerage Report"
            .Formulas(4) = "wFrToDt = 'From ' & '" & Format(dtpFrDt.Value, "dd/MM/yyyy") & "' & '  To ' & '" & Format(dtpToDt.Value, "dd/MM/yyyy") & "'"
            X1 = "{tblAddLess.VDt} in Date(" & Format(dtpFrDt.Value, "yyyy,mm,dd") & ") to Date(" & Format(dtpToDt.Value, "yyyy,mm,dd") & ")  "
            If txtBroker <> "" Then X1 = X1 + " and {tblMastAccount_Brk.AcName}='" & txtBroker & "'" '--Broker
            '-- Sale Type Str
            If OptBrkSAType(0) = True Then '--all
                .Formulas(7) = "Str=''"
                X1 = X1 + " and ({tblAddLess.Vtype}='PY' or {tblAddLess.Vtype}='PI' or {tblAddLess.Vtype}='PT')"
                PrevBrkStr = "(tblAddLess.Vtype in ('PY','PI','PT'))"
            ElseIf OptBrkSAType(1) = True Then '--trade
                .Formulas(7) = "Str='Trade Sales'"
                X1 = X1 + " and {tblAddLess.Vtype} = 'PY'"
                PrevBrkStr = "(tblAddLess.Vtype in ('PY'))"
            ElseIf OptBrkSAType(2) = True Then '--sit
                .Formulas(7) = "Str='SIT Sales'"
                X1 = X1 + " and {tblAddLess.Vtype} = 'PT'"
                PrevBrkStr = "(tblAddLess.Vtype in ('PT'))"
            ElseIf OptBrkSAType(3) = True Then '--consig
                .Formulas(7) = "Str='Consignment Sales'"
                X1 = X1 + " and {tblAddLess.Vtype} = 'PI'"
                PrevBrkStr = "(tblAddLess.Vtype in ('PI'))"
            ElseIf OptBrkSAType(4) = True Then '--Depot
                .Formulas(7) = "Str='Depot Sales'"
                X1 = X1 + " and {tblAddLess.Vtype} = 'PI'"
                PrevBrkStr = "(tblAddLess.Vtype in ('PI'))"
            End If
            '--- Pevious Brokerage
            gCn.BeginTrans
            gCn.Execute "delete from tmpGenTbl where Username='" & gUserName & "'"
            If gBackEndDB = gBackEndAccess Then
                PrevBrkStr = "tblAddLess.Vdt<cdate('" & Format(dtpFrDt.Value, "dd/MM/yyyy") & "') and " & PrevBrkStr
            Else
                PrevBrkStr = "tblAddLess.Vdt<To_date('" & dtpFrDt.Value & "', 'dd/MM/yyyy') and " & PrevBrkStr
            End If
            If gBackEndDB = gBackEndAccess Then
                X = "sum((iif(ItBrokOnPurch='Bag',VSUBBAG,iif(ItBrokOnPurch='Kg',VSUBWT/100,ADBILLAMT/100)) * ITBROKRT))"
                
'                X = "sum(VSUBBAG*ITOUTBRKRT)"
            Else
                X = "sum(Round(decode(ITBROKON,'Bag',VSUBBAG,decode(ITBROKON,'Kg',VSUBWT/100,ADBILLAMT/100)) * ITBROKRT))"
            End If
            X = " select AcCode,( select " & X & " from tblOutStanding,TblAddLess,TblVouSub,tblMastItem where OUTBROKERCODE =acCode " _
            & " and  VSUBITCODE =ItCode and tblAddLess.Vtype=TblVouSub.Vtype and  tblAddLess.Vno=TblVouSub.Vno and  tblAddLess.VYear=TblVouSub.VYear and  " _
            & " tblAddLess.Vtype=tblOutStanding.Vtype and  tblAddLess.Vno=tblOutStanding.Vno and  tblAddLess.VYear=tblOutStanding.VYear and  " _
            & PrevBrkStr & ") , " & IIf(gBackEndDB = gBackEndAccess, "cdate('" & Format(CDate(dtpFrDt.Value - 1), "dd/MM/yyyy") & "')", "to_date('" & CDate(dtpFrDt.Value) - 1 & "','dd/MM/yyyy')") & "," & txt(1) & "," & txt(2) & "," & txt(3) & "," & txt(4) & ",'" & gUserName & "'" _
            & " from tblmastaccount " & IIf(txtBroker <> "", " where AcName='" & txtBroker & "'", "") & " group by AcCode"
            gCn.Execute "insert into TmpGenTbl (VAccode,Amt1,Vdt,Amt2,Amt3,Amt4,Amt5,UserName) " & X
            gCn.Execute "update tmpgentbl set amt1=0 where amt1 is null and UserName='" & gUserName & "'"
            gCn.CommitTrans
            '----- Prev Brk End
            If gBackEndDB = gBackEndAccess Then
                .DataFiles(0) = gDbLocation
                gReportPath = App.Path & "\HiReports_Access\"
            Else '-- Oracle
                .Connect = "DSN=" + gODBCDSN + " ;UID=Scott;PWD=" + gODBCPwd + ";DBQ=" + gODBCHostStr
                gReportPath = App.Path & "\HiReports_ORA\"
            End If
            X1 = X1 + " and {tmpGenTbl.UserName}='" & gUserName & "' and {@Brkamt}>0"
            .SelectionFormula = X1
            .ReportFileName = gReportPath & "rptBrokeragePurchRptVoucher.rpt" '-Voucher Report
         Case "Inward/Purchase Brokerage Detail Report" '--"Inward/Purchase Brokerage Detail Report"
            .Formulas(4) = "wFrToDt = 'From : ' & '" & Format(dtpFrDt.Value, "dd/MM/yyyy") & "' & '  To ' & '" & Format(dtpToDt.Value, "dd/MM/yyyy") & "'"
            X1 = "{tblAddLess.VDt} in Date(" & Format(dtpFrDt.Value, "yyyy,mm,dd") & ") to Date(" & Format(dtpToDt.Value, "yyyy,mm,dd") & ")  "
            If txtBroker <> "" Then X1 = X1 + " and {tblMastAccount_Brk.AcName}='" & txtBroker & "'"
            If txtMill <> "" Then X1 = X1 + " and {tblMastAccount.AcName}='" & txtMill & "'"
            If txtItem <> "" Then X1 = X1 + " and {tblMastItem.ItName}='" & txtItem & "'"
            '-- Sale Type Str
            If OptBrkSAType(0) = True Then '--all
                .Formulas(7) = "Str=''"
                X1 = X1 + " and ({tblAddLess.Vtype}='PY' or {tblAddLess.Vtype}='PT' or {tblAddLess.Vtype}='PI')"
            ElseIf OptBrkSAType(1) = True Then '--trade
                .Formulas(7) = "Str='Trade Sales'"
                X1 = X1 + " and {tblAddLess.Vtype} = 'PY'"
            ElseIf OptBrkSAType(2) = True Then '--sit
                .Formulas(7) = "Str='SIT Sales'"
                X1 = X1 + " and {tblAddLess.Vtype} = 'PT'"
            ElseIf OptBrkSAType(3) = True Then '--consig
                .Formulas(7) = "Str='Consignment Sales'"
                X1 = X1 + " and {tblAddLess.Vtype} = 'PI'"
            ElseIf OptBrkSAType(4) = True Then '--Depot
                .Formulas(7) = "Str='Depot Sales'"
                X1 = X1 + " and {tblAddLess.Vtype} = 'PI'"
            End If
            .SelectionFormula = X1 & " and {@BrkAmt}+ {@BrkAmt2} >0"
            If ChkSummary.Value = 1 Then
                .Formulas(8) = "TdsPer='" & txt(2) & "'"
               .ReportFileName = gReportPath & "RptBrokerageSummary.rpt"
            Else
               .ReportFileName = gReportPath & "rptBrokeragePurchRptDetail.rpt"
            End If
         Case "Inward/Purchase Brokarage Total" '--Inward/Purchase Brokarage Total
            .Formulas(4) = "wFrToDt = 'From ' & '" & Format(dtpFrDt.Value, "dd/MM/yyyy") & "' & '  To ' & '" & Format(dtpToDt.Value, "dd/MM/yyyy") & "'"
            X1 = "{tblAddLess.VDt} in Date(" & Format(dtpFrDt.Value, "yyyy,mm,dd") & ") to Date(" & Format(dtpToDt.Value, "yyyy,mm,dd") & ")  "
            If txtBroker <> "" Then X1 = X1 + " and {tblMastAccount_Brk.AcName}='" & txtBroker & "'" '--Broker
            '-- Sale Type Str
            If OptBrkSAType(0) = True Then '--all
                .Formulas(7) = "Str=''"
                X1 = X1 + " and ({tblAddLess.Vtype}='PY' or {tblAddLess.Vtype}='PT' or {tblAddLess.Vtype}='PI')"
                PrevBrkStr = "(tblAddLess.Vtype in ('PY','PT','PI'))"
            ElseIf OptBrkSAType(1) = True Then '--trade
                .Formulas(7) = "Str='Trade Sales'"
                X1 = X1 + " and {tblAddLess.Vtype} = 'PY'"
                PrevBrkStr = "(tblAddLess.Vtype in ('PY'))"
            ElseIf OptBrkSAType(2) = True Then '--sit
                .Formulas(7) = "Str='SIT Sales'"
                X1 = X1 + " and {tblAddLess.Vtype} = 'PT'"
                PrevBrkStr = "(tblAddLess.Vtype in ('PT'))"
            ElseIf OptBrkSAType(3) = True Then '--consig
                .Formulas(7) = "Str='Consignment Sales'"
                X1 = X1 + " and {tblAddLess.Vtype} = 'PI'"
                PrevBrkStr = "(tblAddLess.Vtype in ('PI'))"
            ElseIf OptBrkSAType(4) = True Then '--Depot
                .Formulas(7) = "Str='Depot Sales'"
                X1 = X1 + " and {tblAddLess.Vtype} = 'PI'"
                PrevBrkStr = "(tblAddLess.Vtype in ('PI'))"
            End If
            '--- Pevious Brokerage
            gCn.BeginTrans
            gCn.Execute "delete from tmpGenTbl where Username='" & gUserName & "'"
            If gBackEndDB = gBackEndAccess Then
                PrevBrkStr = "tblAddLess.Vdt<cdate('" & Format(dtpFrDt.Value, "dd/MM/yyyy") & "') and " & PrevBrkStr
            Else
                PrevBrkStr = "tblAddLess.Vdt<To_date('" & dtpFrDt.Value & "', 'dd/MM/yyyy') and " & PrevBrkStr
            End If
            
            If gBackEndDB = gBackEndAccess Then
                X = "sum((iif(ItBrokOnPurch='Bag',VSUBBAG,iif(ItBrokOnPurch='Kg',VSUBWT/100,ADBILLAMT/100)) * ITBROKRT))"
'                X = "sum(VSUBBAG*ITOUTBRKRT)"
            Else
                X = "sum(Round(decode(ITBROKON,'Bag',VSUBBAG,decode(ITBROKON,'Kg',VSUBWT/100,ADBILLAMT/100)) * ITBROKRT))"
            End If
            X = " select AcCode,( select " & X & " from tblOutStanding,TblAddLess,TblVouSub,tblMastItem where OUTBROKERCODE =acCode " _
            & " and  VSUBITCODE =ItCode and tblAddLess.Vtype=TblVouSub.Vtype and  tblAddLess.Vno=TblVouSub.Vno and  tblAddLess.VYear=TblVouSub.VYear and  " _
            & " tblAddLess.Vtype=tblOutStanding.Vtype and  tblAddLess.Vno=tblOutStanding.Vno and  tblAddLess.VYear=tblOutStanding.VYear and  " _
            & PrevBrkStr & ") , " & IIf(gBackEndDB = gBackEndAccess, "cdate('" & Format(CDate(dtpFrDt.Value - 1), "dd/MM/yyyy") & "')", "to_date('" & CDate(dtpFrDt.Value) - 1 & "','dd/MM/yyyy')") & "," & txt(1) & "," & txt(2) & "," & txt(3) & "," & txt(4) & ",'" & gUserName & "'" _
            & " from tblmastaccount " & IIf(txtBroker <> "", " where AcName='" & txtBroker & "'", "") & " group by AcCode"
            gCn.Execute "insert into TmpGenTbl (VAccode,Amt1,Vdt,Amt2,Amt3,Amt4,Amt5,UserName) " & X
            gCn.Execute "update tmpgentbl set amt1=0 where amt1 is null and UserName='" & gUserName & "'"
            gCn.CommitTrans
            '----- Prev Brk End
            If gBackEndDB = gBackEndAccess Then
                .DataFiles(0) = gDbLocation
                gReportPath = App.Path & "\HiReports_Access\"
            Else '-- Oracle
                .Connect = "DSN=" + gODBCDSN + " ;UID=Scott;PWD=" + gODBCPwd + ";DBQ=" + gODBCHostStr
                gReportPath = App.Path & "\HiReports_ORA\"
            End If
            X1 = X1 + " and {tmpGenTbl.UserName}='" & gUserName & "' and {@Brkamt}>0"
            .SelectionFormula = X1
            .ReportFileName = gReportPath & "rptPurchaseBrokerageRptSumm.rpt" '-Voucher Report
        Case "Cash - Bank  Status"
             Dim mTempCn As Connection, mCurrFirmDbNm As String
             Dim RstCompMast As Recordset
             Dim CurrentFirmFlag As Boolean
             Dim ComCd As String
             GProcRstOpen RstCompMast, "Select * from tblmastcompany " & IIf(CurrentFirmFlag = True, " where cname='" & gCName & "'", " "), "R", gSetupCn
             mCurrFirmDbNm = gDbNm
             RstCompMast.MoveFirst
             Do While Not RstCompMast.EOF
'                If RstCompMast!CIsPersonalFirm = 0 Then
                    gDbNm = RstCompMast!CCode + Right(mCurrFirmDbNm, 4)
                    ComCd = RstCompMast!CCode
                    GProcConnectionOpen mTempCn
                    mTempCn.BeginTrans
                    mTempCn.Execute "delete from tmpSelection where UserName='" & gUserName & "'"
                    If gBackEndDB = gBackEndAccess Then
                       mTempCn.Execute "insert into tmpSelection (dt1,dt2,UserName) values ('" & dtpFrDt.Value & "','" & dtpToDt.Value & "','" & gUserName & "')"
                    Else
                       mTempCn.Execute "insert into tmpSelection (dt1,dt2,UserName) values (to_date('" & dtpFrDt.Value & "','dd/MM/yyyy'),to_date('" & dtpToDt.Value & "','dd/MM/yyyy'),'" & gUserName & "') "
                    End If
                    mTempCn.CommitTrans
                    mTempCn.Close
 '               End If
                RstCompMast.MoveNext
             Loop
             gDbNm = mCurrFirmDbNm
             GProcConnectionOpen gCn
             GProcFormIssueReceive Format(dtpFrDt.Value, "dd/MM/yyyy"), Format(dtpToDt.Value, "dd/MM/yyyy"), "CsBnkStatus", "", "", "", ChkConsolidate, IIf(Opt(0) = True, "All", IIf(Opt(1) = True, "Pending", "Receive"))
            .Formulas(4) = "wFrToDt = 'Date : ' & '" & Format(dtpToDt.Value, "dd/MM/yyyy") & "' "
'            .Formulas(4) = "wFrToDt = 'Date : ' & '" & Format(dtpFrDt.Value, "dd/MM/yyyy") & "' & '  To ' & '" & Format(dtpToDt.Value, "dd/MM/yyyy") & "'"
            .DataFiles(0) = App.Path & "\HIData\setup.mdb"
            .SelectionFormula = "{tmpgenTbl.Username}='" & gUserName & "'"
            gReportPath = App.Path & "\HiReports\"
            '-- Opt Str
            .ReportFileName = gReportPath & "rptCashBankStatus.rpt"
        '================ General Report End
    End Select
    .Action = 1
End With
Screen.MousePointer = vbDefault
End Sub



