VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form frmRptGST 
   BackColor       =   &H80000013&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Register"
   ClientHeight    =   5670
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10215
   Icon            =   "frmRptGST.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5670
   ScaleWidth      =   10215
   Begin VB.CheckBox ChkRegular 
      BackColor       =   &H80000018&
      Caption         =   "Regular O/s"
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
      Left            =   8400
      TabIndex        =   39
      Top             =   1680
      Width           =   1560
   End
   Begin VB.CheckBox Chkonly 
      BackColor       =   &H80000018&
      Caption         =   "2nd Day O/s"
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
      Left            =   8400
      TabIndex        =   38
      Top             =   1200
      Width           =   1560
   End
   Begin VB.CheckBox ChkUnAdjustRec 
      BackColor       =   &H80000018&
      Caption         =   "Include Unadjusted Receipts?"
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
      Left            =   5715
      TabIndex        =   21
      Top             =   2175
      Width           =   4230
   End
   Begin VB.CheckBox ChkStkValExpense 
      BackColor       =   &H80000018&
      Caption         =   "Include &Expenses in valuation ?"
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
      Left            =   10275
      TabIndex        =   22
      Top             =   1575
      Visible         =   0   'False
      Width           =   3570
   End
   Begin VB.CheckBox ChkStkValTax 
      BackColor       =   &H80000018&
      Caption         =   "Include &Tax in valuation ?"
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
      Left            =   5700
      TabIndex        =   19
      Top             =   1395
      Width           =   4350
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
      Height          =   375
      Index           =   6
      Left            =   1380
      TabIndex        =   34
      Top             =   3585
      Width           =   3960
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
      Height          =   375
      Index           =   1
      Left            =   6765
      TabIndex        =   17
      Text            =   "1"
      Top             =   615
      Width           =   3165
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
      Height          =   375
      Index           =   5
      Left            =   1380
      TabIndex        =   6
      Top             =   1170
      Width           =   3960
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
      Height          =   375
      Index           =   4
      Left            =   1380
      TabIndex        =   14
      Top             =   3060
      Width           =   3960
   End
   Begin VB.CheckBox chkIsDespatchDt 
      BackColor       =   &H80000018&
      Caption         =   "Is Report For Despatch Date ?"
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
      Height          =   450
      Left            =   7590
      TabIndex        =   25
      Top             =   3600
      Width           =   2475
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
      Left            =   7590
      TabIndex        =   23
      Top             =   2805
      Width           =   2430
   End
   Begin VB.CheckBox chkMonthwise 
      BackColor       =   &H80000018&
      Caption         =   "Is Monthwise Report ?"
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
      Left            =   7590
      TabIndex        =   24
      Top             =   3180
      Width           =   2310
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
      Height          =   375
      Index           =   3
      Left            =   1380
      TabIndex        =   12
      Top             =   2580
      Width           =   3960
   End
   Begin VB.CommandButton cmdCrystal 
      Height          =   480
      Index           =   0
      Left            =   1065
      Picture         =   "frmRptGST.frx":058A
      Style           =   1  'Graphical
      TabIndex        =   29
      Top             =   4905
      Width           =   1185
   End
   Begin VB.CommandButton cmdCrystal 
      Height          =   480
      Index           =   1
      Left            =   2490
      Picture         =   "frmRptGST.frx":09CC
      Style           =   1  'Graphical
      TabIndex        =   30
      ToolTipText     =   "Print"
      Top             =   4905
      Width           =   1185
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
      Height          =   375
      Index           =   2
      Left            =   1380
      TabIndex        =   10
      Top             =   2100
      Width           =   3960
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
      Height          =   375
      Index           =   1
      Left            =   1380
      TabIndex        =   8
      Top             =   1620
      Width           =   3960
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
      Height          =   375
      Index           =   0
      Left            =   1380
      TabIndex        =   4
      Top             =   720
      Width           =   3960
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H80000018&
      Height          =   1200
      Left            =   5715
      TabIndex        =   26
      Top             =   2820
      Width           =   1815
      Begin VB.OptionButton opsummary 
         BackColor       =   &H80000018&
         Caption         =   "Summary"
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
         Height          =   315
         Left            =   120
         TabIndex        =   37
         Top             =   840
         Visible         =   0   'False
         Width           =   1605
      End
      Begin VB.OptionButton OptShort 
         BackColor       =   &H80000018&
         Caption         =   "Short"
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
         Height          =   315
         Left            =   120
         TabIndex        =   28
         Top             =   480
         Width           =   1605
      End
      Begin VB.OptionButton OptDetail 
         BackColor       =   &H80000018&
         Caption         =   "Detail"
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
         Height          =   315
         Left            =   120
         TabIndex        =   27
         Top             =   120
         Value           =   -1  'True
         Width           =   1605
      End
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
      Height          =   495
      Left            =   3915
      TabIndex        =   31
      Top             =   4905
      Width           =   1335
   End
   Begin Crystal.CrystalReport CrystalReport2 
      Left            =   0
      Top             =   6270
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
   Begin MSComctlLib.ProgressBar prgBar 
      Height          =   255
      Left            =   105
      TabIndex        =   32
      Top             =   4395
      Visible         =   0   'False
      Width           =   9645
      _ExtentX        =   17013
      _ExtentY        =   450
      _Version        =   393216
      Appearance      =   1
   End
   Begin MSComCtl2.DTPicker dtpToDt 
      Height          =   360
      Left            =   3540
      TabIndex        =   2
      Top             =   255
      Width           =   1800
      _ExtentX        =   3175
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
      Format          =   52166659
      CurrentDate     =   38050
   End
   Begin MSComCtl2.DTPicker dtpFrDt 
      Height          =   360
      Left            =   1380
      TabIndex        =   0
      Top             =   255
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
      Format          =   52166659
      CurrentDate     =   38050
   End
   Begin Crystal.CrystalReport CrystalReport1 
      Bindings        =   "frmRptGST.frx":1036
      Left            =   0
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
   Begin VB.Label Label11 
      BackStyle       =   0  'Transparent
      Caption         =   "&Outstanding Reports :"
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
      Left            =   5700
      TabIndex        =   20
      Top             =   1860
      Width           =   2115
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "From &Date  :"
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
      Left            =   120
      TabIndex        =   36
      Top             =   255
      Width           =   1170
   End
   Begin VB.Shape Shape3 
      FillColor       =   &H00BDDFD1&
      FillStyle       =   0  'Solid
      Height          =   765
      Left            =   765
      Shape           =   4  'Rounded Rectangle
      Top             =   4800
      Width           =   4785
   End
   Begin VB.Line Line2 
      BorderColor     =   &H000040C0&
      X1              =   10080
      X2              =   5580
      Y1              =   2670
      Y2              =   2670
   End
   Begin VB.Label Label10 
      BackStyle       =   0  'Transparent
      Caption         =   "Stock &Valuation :"
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
      Left            =   5685
      TabIndex        =   18
      Top             =   1080
      Width           =   2115
   End
   Begin VB.Label Label9 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "&Sizer    :"
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
      Left            =   120
      TabIndex        =   35
      Top             =   3585
      Width           =   750
   End
   Begin VB.Label Label8 
      BackStyle       =   0  'Transparent
      Caption         =   "&Daily Stock &Statement To Mill  :"
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
      Left            =   5685
      TabIndex        =   15
      Top             =   255
      Width           =   3435
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "C. C. To :"
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
      Left            =   5685
      TabIndex        =   16
      Top             =   615
      Width           =   840
   End
   Begin VB.Label Label7 
      ForeColor       =   &H000040C0&
      Height          =   330
      Left            =   1545
      TabIndex        =   33
      Top             =   5070
      Width           =   2460
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "&Item          :"
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
      Left            =   120
      TabIndex        =   5
      Top             =   1170
      Width           =   1035
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "&Godown    :"
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
      Left            =   120
      TabIndex        =   13
      Top             =   3060
      Width           =   1005
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "&Broker       :"
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
      Left            =   120
      TabIndex        =   11
      Top             =   2580
      Width           =   1050
   End
   Begin VB.Label lblParty 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "&Party         : "
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
      Left            =   120
      TabIndex        =   9
      Top             =   2100
      Width           =   1125
   End
   Begin VB.Label lblTaxType 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "&Tax Type   :"
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
      Left            =   120
      TabIndex        =   7
      Top             =   1620
      Width           =   1035
   End
   Begin VB.Label lblMillName 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "&Mill            :"
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
      Left            =   120
      TabIndex        =   3
      Top             =   735
      Width           =   1080
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
      Left            =   3135
      TabIndex        =   1
      Top             =   255
      Width           =   360
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00000080&
      FillColor       =   &H00EDFDFE&
      FillStyle       =   0  'Solid
      Height          =   4140
      Left            =   45
      Top             =   90
      Width           =   5430
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H00000080&
      FillColor       =   &H00EDFDFE&
      FillStyle       =   0  'Solid
      Height          =   4140
      Left            =   5580
      Top             =   90
      Width           =   4515
   End
End
Attribute VB_Name = "frmRptGST"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim wNarrCode As Long, wMillCode As Long, wPartyCode As Long, wBrkCode As Long, wItemCode As Long
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
    Case 13:
       Sendkeys "{TAB}"
    Case 27:
        Unload Me
    End Select
End Sub
Private Sub Form_Load()
wBrkCode = 0
wMillCode = 0
wPartyCode = 0
wNarrCode = 0
wItemCode = 0
Me.Top = 700
Me.Left = Screen.Width / 2 - Me.Width / 2
 '-- Activate Controls to False
GProcActivateControls Me, False
txt(1) = "" '-CCTo
dtpFrDt.Enabled = True
dtpToDt.Enabled = True
dtpFrDt.Value = Format(gCYSDate, "dd/MM/yyyy")
dtpToDt.Value = Format(gCYEDate, "dd/MM/yyyy")
Me.Caption = gSelectedMenu
Select Case gSelectedMenu
'==== Sale Register
Case "Trade Sale Register", "TDS Report" '--"Trade Sale Register":
    If gSelectedMenu = "TDS Report" Then
       txttemp(1).Enabled = False '-Tax
       txttemp(5).Enabled = False '-Count
       txttemp(2).Enabled = True '-Count
       txttemp(3).Enabled = False '-Broker
       OptDetail.Enabled = True
       OptShort.Enabled = True
       OptShort.Caption = "Summery"
       ChkConsolidate.Enabled = True
       chkMonthwise.Enabled = True
       ChkConsolidate.Caption = "Yarn Purchase"
       chkMonthwise.Caption = "Other Purchase"
    Else
        OptDetail.Enabled = True
        OptShort.Enabled = True
        Chkonly.Enabled = True
        txttemp(0).Enabled = True '-Mill
        txttemp(1).Enabled = True '-Tax
        txttemp(2).Enabled = True '-Party
        txttemp(5).Enabled = True '-Count
        txttemp(3).Enabled = True '-Broker
        opsummary.Visible = True
        opsummary.Enabled = True
    End If
Case "Trade Sale Register GST", "Trade Purchase Register GST" '--"Trade Sale Register":
    txttemp(0).Enabled = True '-Mill
    txttemp(2).Enabled = True '-Party
    txttemp(5).Enabled = True '-Count
    txttemp(3).Enabled = True '-Broker
    txttemp(4).Enabled = True '-Godown
    Label4.Caption = "Transport :"
    ChkConsolidate.Enabled = True
    ChkConsolidate.Caption = "Is Transportwise?"
    ChkUnAdjustRec.Enabled = True
    ChkUnAdjustRec.Caption = "With E-way Bill No ?"
    chkMonthwise.Enabled = True
    chkMonthwise.Caption = "With Booking No.?"
    
Case "Debit Note Register GST", "Credit Note Register GST"
     OptDetail.Enabled = True
     OptDetail.Caption = "On Sales"
     OptShort.Enabled = True
     OptShort.Caption = "On Purch"
     opsummary.Visible = False
Case "Insurance Report"
    OptDetail.Enabled = False
    OptShort.Enabled = False
    txttemp(0).Enabled = True  '--Mill
    txttemp(2).Enabled = True '-Party
    chkMonthwise.Enabled = False
    Label6.Caption = "Insu.Amt"
    txt(1).Enabled = True
Case "Form GSTR - 2"
    txttemp(2).Enabled = True
    OptDetail.Enabled = True
    OptShort.Enabled = True
    OptShort.Caption = "Summary"
    ChkConsolidate.Enabled = True
    ChkConsolidate.Caption = "Is Datewise Report"
Case "Form GSTR - 1", "Form GSTR - 2", "Monthwise Summery GST", "HSN Summery", "Purchase And Expences Detail (Audit)", "Purchase With Payment Detaill (Audit)"
    txttemp(0).Visible = False
    txttemp(1).Visible = False
    txttemp(2).Visible = False
    txttemp(3).Visible = False
    txttemp(4).Visible = False
    txttemp(5).Visible = False
    txttemp(6).Visible = False
    lblMillName.Visible = False
    Label5.Visible = False
    lblTaxType.Visible = False
    lblParty.Visible = False
    Label3.Visible = False
    Label4.Visible = False
    Label9.Visible = False
    Label8.Visible = False
    Label6.Visible = False
    txt(1).Visible = False
    Label10.Visible = False
    ChkStkValTax.Visible = False
    Label11.Visible = False
    ChkUnAdjustRec.Visible = False
    ChkConsolidate.Visible = False
    chkMonthwise.Visible = False
    chkIsDespatchDt.Visible = False
    If gSelectedMenu = "Form GSTR - 1" Then
       OptDetail.Value = True
       OptDetail.Enabled = True
       OptDetail.Caption = "Excel"
       OptShort.Caption = "Json"
       OptShort.Enabled = True
    End If
    
    If gSelectedMenu = "Form GSTR - 2" Then
        txttemp(2).Visible = True
        txttemp(2).Enabled = True
        lblParty.Visible = True
        OptDetail.Visible = True
        OptShort.Visible = True
        opsummary.Visible = True
        OptDetail.Enabled = True
        OptShort.Enabled = True
        opsummary.Enabled = True
        OptDetail.Caption = "Summery"
        OptShort.Caption = "Detail"
        opsummary.Caption = "GSTR-2"
    Else
''        OptDetail.Visible = False
''        OptShort.Visible = False
''        opsummary.Visible = False
''        OptDetail.Enabled = False
''        OptShort.Enabled = False
''        opsummary.Enabled = False
''        OptDetail.Caption = "B2B"
''        OptShort.Caption = "Dr.Nt./Cr.Nt."
''        opsummary.Caption = "HSN wise"
  End If
  If gSelectedMenu = "Monthwise Summery GST" Or gSelectedMenu = "HSN Summery" Then
        OptDetail.Enabled = True
        OptShort.Enabled = True
        OptDetail.Value = True
        OptDetail.Visible = True
        OptShort.Visible = True
        OptDetail.Caption = "Pur. Bill Dt."
        OptShort.Caption = "Pur. Recd Dt"
        Chkonly.Enabled = True
        If gSelectedMenu = "Monthwise Summery GST" Then
         Chkonly.Caption = "Is Datewise"
        End If
   End If

End Select
 '-- Set InActivate Ctrl Back Color
GProcSetInActivateCtrlBackClr Me
End Sub
Private Sub cmdExit_Click()
    Unload Me
End Sub

Private Sub OptShort_Click()
 '-- MillWise Register For Detail Trade Sale  only
If gSelectedMenu = "Trade Sale Register" Then
    If OptShort.Value = True Then txttemp(0) = ""
End If
End Sub

Private Sub txt_GotFocus(Index As Integer)
GProcSelectBox Me.ActiveControl
End Sub

Private Sub txttemp_GotFocus(Index As Integer)
GProcSelectBox Me.ActiveControl
End Sub
Private Sub txtTemp_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = 13 Then Exit Sub '-- Enter
    Select Case Index
    Case 0 '----Mill Name
        If gSelectedMenu = "Receipt Register" Or gSelectedMenu = "Payment/RTGS Register" Then
           gClsSearch.SearchMultiField "QryMastaccount", "Account,AcName,ACCode", Array("Bank", "AcName", ""), Array(txttemp(Index).Width, 0, 0), " (AgCode=90019)", Chr(KeyAscii), "AcName", txttemp(Index).Left + Me.Left, txttemp(Index).Top + Me.Top + 1000
        Else
           gClsSearch.SearchMultiField "QryMastaccount", "Account,AcName,ACCode", Array("Mill", "AcName", ""), Array(txttemp(Index).Width, 0, 0), " (AgCode=90029 or GpCode=90029)", Chr(KeyAscii), "AcName", txttemp(Index).Left + Me.Left, txttemp(Index).Top + Me.Top + 1000
        End If
        KeyAscii = 0
        If gClsSearch.SearchMultiRetCol(1) <> "" Then
            txttemp(Index) = gClsSearch.SearchMultiRetCol(1)
            wMillCode = gClsSearch.SearchMultiRetCol(2)
        End If
    Case 1: '--Tax
        gClsSearch.SearchMultiField "tblMastNarration", "Narration,NarrCode", Array("Tax Type", ""), Array(txttemp(1).Width, 0), "narrtype='T'", Chr(KeyAscii), "Narration", txttemp(Index).Left + Me.Left, txttemp(Index).Top + Me.Top + 1000
        KeyAscii = 0
        If gClsSearch.SearchMultiRetCol(0) <> "" Then
            txttemp(1) = gClsSearch.SearchMultiRetCol(0)
            wNarrCode = gClsSearch.SearchMultiRetCol(1)
        End If
    Case 2 '----Party Name
        gClsSearch.SearchMultiField "QryMastaccount", "Account,AcName,ACCode", Array("Party", "AcName", ""), Array(txttemp(Index).Width, 0, 0), " ", Chr(KeyAscii), "AcName", txttemp(Index).Left + Me.Left, txttemp(Index).Top + Me.Top + 1000
        KeyAscii = 0
        If gClsSearch.SearchMultiRetCol(1) <> "" Then
            txttemp(Index) = gClsSearch.SearchMultiRetCol(1)
            wPartyCode = gClsSearch.SearchMultiRetCol(2)
        End If
     Case 3 '--Broker
        gClsSearch.SearchMultiField "QryMastaccount", "Account,AcName,ACCode", Array("Broker", "AcName", ""), Array(txttemp(Index).Width, 0, 0), " (AgCode=90016 or GpCode=90016)", Chr(KeyAscii), "AcName", txttemp(Index).Left + Me.Left, txttemp(Index).Top + Me.Top + 1000
        KeyAscii = 0
        If gClsSearch.SearchMultiRetCol(1) <> "" Then
            txttemp(Index) = gClsSearch.SearchMultiRetCol(1)
            wBrkCode = gClsSearch.SearchMultiRetCol(2)
        End If
     Case 4 '--Godown,Transport
        gClsSearch.SearchMultiField "tblMastNarration", "Narration", Array("Godown"), Array(txttemp(Index).Width), " NarrType in ('G','R')", Chr(KeyAscii), "Narration", txttemp(Index).Left + Me.Left, txttemp(Index).Top + Me.Top + 1000
        KeyAscii = 0
        If gClsSearch.SearchMultiRetCol(0) <> "" Then
            txttemp(Index) = gClsSearch.SearchMultiRetCol(0)
        End If
     Case 5 '-- Item
        gClsSearch.SearchMultiField "tblMastItem,tblMastAccount", "ItName,ItCode,Itticket,Acname", Array("Item", "", "Ticket", "Mill"), Array(2000, 0, 1000, 2000), " tblMastItem.ItMillcode=tblMastAccount.AcCode", Chr(KeyAscii), "ItName", txttemp(Index).Left + Me.Left, txttemp(Index).Top + Me.Top + 1000
        KeyAscii = 0
        If gClsSearch.SearchMultiRetCol(0) <> "" Then
           wItemCode = gClsSearch.SearchMultiRetCol(1) '-Itcode
           txttemp(Index) = gClsSearch.SearchMultiRetCol(0) '-ItName
           txttemp(0) = gClsSearch.SearchMultiRetCol(3) '-Mill
           wMillCode = GProcGetColumnValue("tblMastAccount", "AcName", txttemp(0), "S", "AcCode", "N")
        End If
      Case 6 '--Sizer
        gClsSearch.SearchMultiField "tbladdLess", " distinct AdTmp5", Array("Sizer"), Array(txttemp(Index).Width), " Vtype in ('GP')", Chr(KeyAscii), "AdTmp5", txttemp(Index).Left + Me.Left, txttemp(Index).Top + Me.Top + 1000
        KeyAscii = 0
        If gClsSearch.SearchMultiRetCol(0) <> "" Then
            txttemp(Index) = gClsSearch.SearchMultiRetCol(0)
        End If
    End Select
End Sub
Private Sub cmdCrystal_Click(Index As Integer)
Dim X1 As String
Dim X As String
Dim DtCondStr As String
Dim flnm As String
Dim BarCnt As Double
Dim RecYN As Long
Dim tmp1 As Recordset
Dim excel_app As Excel.Application
Dim workbook As Excel.workbook
Dim sheet_name As String
Dim sheet As Excel.Worksheet
Dim header_range As Excel.Range
Dim a As Long
Dim TotCount As Double
Dim xlApp As Object
Dim xlWB As Object

X1 = ""
If txttemp(0) = "" Then wMillCode = 0 '- Mill Code
With CrystalReport1
    GProcCrystalRptPreparation CrystalReport1, IIf(Index = 0, "Window", "Printer")
    Select Case gSelectedMenu
        Case "Trade Sale Register GST"
            .Formulas(4) = "wFrToDt='From : '& '" & dtpFrDt.Value & "' & ' To ' & '" & dtpToDt.Value & "'"
            .Formulas(5) = "wCompGstin = '" & gCGSTIN & "'"
            X1 = "{tblAddLess.VType}='SY' and {tblAddLess.VDt} in Date(" & Format(dtpFrDt.Value, "yyyy,mm,dd") & ") to Date(" & Format(dtpToDt.Value, "yyyy,mm,dd") & ")  " '
            '-- Party
            If txttemp(2) <> "" Then X1 = X1 + " and {tblAddLess.AdAcDrCode} = " & wPartyCode
            '-- Mill
            If txttemp(0) <> "" Then X1 = X1 + " and {tblMastItem.ITMillCode} = " & wMillCode
            '-- Count
            If txttemp(5) <> "" Then X1 = X1 + " and {tblVouSub.VSubITCode} = " & wItemCode
            
            '-- Broker
            If txttemp(3) <> "" Then X1 = X1 + " and {tblOutstanding.OutBrokerCode} = " & wBrkCode
            
            '-- Godown
            'If txtTemp(4) <> "" Then X1 = X1 + " and {tblMastNarration.Narrcode} = " & wNarrCode
            If txttemp(4) <> "" Then X1 = X1 + " and {tblMastNarration.Narration} = '" & txttemp(4) & "'"

            .SelectionFormula = X1
            If chkMonthwise.Value = 1 Then
               .ReportFileName = gReportPath & "rptTradeSaleRegiBrooking.rpt"
            Else
               .ReportFileName = gReportPath & "rptTradeSaleRegiGST.rpt"
            End If
        Case "R C M Sale Register"
            .Formulas(4) = "wFrToDt='From : '& '" & dtpFrDt.Value & "' & ' To ' & '" & dtpToDt.Value & "'"
            .Formulas(5) = "wCompGstin = '" & gCGSTIN & "'"
            X1 = "{tblAddLess.VType}='PO' and {tblAddLess.VDt} in Date(" & Format(dtpFrDt.Value, "yyyy,mm,dd") & ") to Date(" & Format(dtpToDt.Value, "yyyy,mm,dd") & ")  " '
            X1 = X1 + " and {tblAddLess.ADTMP2}=1 and {tblVouSub.Vsubctrno} = 1 " '
            .SelectionFormula = X1
            .ReportFileName = gReportPath & "rptRCMSaleReg.rpt"
        Case "Debit Note Register GST"
            .Formulas(4) = "wFrToDt='From : '& '" & dtpFrDt.Value & "' & ' To ' & '" & dtpToDt.Value & "'"
            .Formulas(5) = "wCompGstin = '" & gCGSTIN & "'"
            If OptDetail.Value = True Then
               X1 = "{qrytblVoucherSr1.VType}='SN' and {qrytblVoucherSr1.VDrCrNtType} = 'B'  and {qrytblVoucherSr1.VDt} in Date(" & Format(dtpFrDt.Value, "yyyy,mm,dd") & ") to Date(" & Format(dtpToDt.Value, "yyyy,mm,dd") & ")  " '
            End If
            If OptShort.Value = True Then
               X1 = "{qrytblVoucherSr1.VType}='SN' and {qrytblVoucherSr1.VDrCrNtType} = 'S'  and {qrytblVoucherSr1.VDt} in Date(" & Format(dtpFrDt.Value, "yyyy,mm,dd") & ") to Date(" & Format(dtpToDt.Value, "yyyy,mm,dd") & ")  " '
            End If
            .SelectionFormula = X1
            .ReportFileName = gReportPath & "rptDebitNoteRegGST.rpt"
        Case "Trade Purchase Register GST"
           .Formulas(4) = "wFrToDt='From : '& '" & dtpFrDt.Value & "' & ' To ' & '" & dtpToDt.Value & "'"
           .Formulas(5) = "wCompGstin = '" & gCGSTIN & "'"
            If ChkConsolidate.Value = 1 Then
                X1 = "Select tblAddLess.Vno,tblAddLess.Vdt,tblAddLess.AdGpDt,tblAddLess.AdInvNo,tblAddLess.ADTMP6,tblAddLess.AdDbNtDt,tblMastAccount.AcName,tblMastAccount.AcGSTIN,tblMastItem.Itname,tblMastNarration.Narration,tblVouSub.VSubTaxableAmt,tblAddLess.AdRoff,tblAddLess.AdBillAmt,tblAddLess.AdAftTaxAdd,tblAddLess.AdAftTaxLess,tblVouSub.VsubCGSTRt,tblVouSub.VsubSGSTRt,tblVouSub.VsubIGSTRt,tblVouSub.VSubCtrNo,sum(tblVouSub.VsubCGSTAmt), " _
                     & " sum(tblVouSub.VsubSGSTAmt),sum(tblVouSub.VsubIGSTAmt),sum(iif(tblMastItem.ItOtherItem = 0,VSubBag,0)),sum(iif(tblMastItem.ItOtherItem = 0,VSubWt,0)),tblMastItem.Ittmp2,tblMastItem.Itname,tblVouSub.VsubCtrno,tblAddLess.AdTcsAmt,'" & gUserName & "'" _
                     & " from tblAddLess,tblVouSub,tblMastAccount,tblMastItem,tblMastNarration where tblAddLess.Vno = tblVouSub.Vno and tblAddLess.Vtype = tblVouSub.Vtype and tblAddLess.Vyear = tblVouSub.Vyear and tblAddLess.AdAcCrCode = tblMastAccount.Accode and tblVousub.VsubItcode = tblMastItem.Itcode " _
                     & " and val(tblAddLess.AdTransport) = tblMastNarration.Narrcode and tblAddLess.Vtype = 'PY' and tblAddLess.AdGpDt>=cdate('" & Format(dtpFrDt.Value, "yyyy,mm,dd") & "')  and tblAddLess.AdGpDt<=Cdate('" & Format(dtpToDt.Value, "yyyy,mm,dd") & "') GROUP BY tblAddLess.Vno,tblAddLess.AdGpDt,tblAddLess.Vdt,tblAddLess.AdInvNo,tblMastAccount.AcName,tblMastAccount.AcGSTIN,tblMastNarration.Narration,tblVouSub.VSubTaxableAmt,tblAddLess.AdRoff,tblAddLess.AdBillAmt,tblAddLess.AdAftTaxAdd,tblAddLess.AdAftTaxLess,tblVouSub.VsubCGSTRt,tblVouSub.VsubSGSTRt,tblVouSub.VsubIGSTRt,tblVouSub.VSubCtrNo,tblMastItem.Ittmp2,tblMastItem.Itname,tblVouSub.VsubCtrno,tblAddLess.AdTcsAmt,tblAddLess.ADTMP6,tblAddLess.AdDbNtDt"
                    gCn.BeginTrans
                    gCn.Execute "delete from tmpRptTbl Where Username='" & gUserName & "'"
                    gCn.Execute "insert into tmpRptTbl (Vno,Vdt,VBillDt,VBillNo,Nar7,Dt1,Nar1,Nar2,Nar6,Nar8,Amt1,Amt2,Amt3,Amt9,Amt10,Amt11,Amt12,Amt13,Amt14,Amt4,Amt5,Amt6,Amt7,Amt8,Nar4,Nar5,VACOCODE,Amt15,USERNAME) " & X1
                    gCn.CommitTrans
            Else
                X1 = "Select tblAddLess.Vno,tblAddLess.Vdt,tblAddLess.AdGpDt,tblAddLess.AdInvNo,tblAddLess.ADTMP6,tblAddLess.AdDbNtDt,tblMastAccount.AcName,tblMastAccount.AcGSTIN,tblMastItem.Itname,tblVouSub.VSubTaxableAmt,tblAddLess.AdRoff,tblAddLess.AdBillAmt,tblAddLess.AdAftTaxAdd,tblAddLess.AdAftTaxLess,tblVouSub.VsubCGSTRt,tblVouSub.VsubSGSTRt,tblVouSub.VsubIGSTRt,tblVouSub.VSubCtrNo,sum(tblVouSub.VsubCGSTAmt), " _
                     & " sum(tblVouSub.VsubSGSTAmt),sum(tblVouSub.VsubIGSTAmt),sum(iif(tblMastItem.ItOtherItem = 0,VSubBag,0)),sum(iif(tblMastItem.ItOtherItem = 0,VSubWt,0)),tblMastItem.Ittmp2,tblMastItem.Itname,tblVouSub.VsubCtrno,tblAddLess.AdTcsAmt,'" & gUserName & "'" _
                     & " from tblAddLess,tblVouSub,tblMastAccount,tblMastItem where tblAddLess.Vno = tblVouSub.Vno and tblAddLess.Vtype = tblVouSub.Vtype and tblAddLess.Vyear = tblVouSub.Vyear and tblAddLess.AdAcCrCode = tblMastAccount.Accode and tblVousub.VsubItcode = tblMastItem.Itcode " _
                     & " and tblAddLess.Vtype = 'PY' and tblAddLess.AdGpDt>=cdate('" & Format(dtpFrDt.Value, "yyyy,mm,dd") & "')  and tblAddLess.AdGpDt<=Cdate('" & Format(dtpToDt.Value, "yyyy,mm,dd") & "') GROUP BY tblAddLess.Vno,tblAddLess.AdGpDt,tblAddLess.Vdt,tblAddLess.AdInvNo,tblMastAccount.AcName,tblMastAccount.AcGSTIN,tblVouSub.VSubTaxableAmt,tblAddLess.AdRoff,tblAddLess.AdBillAmt,tblAddLess.AdAftTaxAdd,tblAddLess.AdAftTaxLess,tblVouSub.VsubCGSTRt,tblVouSub.VsubSGSTRt,tblVouSub.VsubIGSTRt,tblVouSub.VSubCtrNo,tblMastItem.Ittmp2,tblMastItem.Itname,tblVouSub.VsubCtrno,tblAddLess.AdTcsAmt,tblAddLess.ADTMP6,tblAddLess.AdDbNtDt"
                    gCn.BeginTrans
                    gCn.Execute "delete from tmpRptTbl Where Username='" & gUserName & "'"
                    gCn.Execute "insert into tmpRptTbl (Vno,Vdt,VBillDt,VBillNo,Nar7,Dt1,Nar1,Nar2,Nar6,Amt1,Amt2,Amt3,Amt9,Amt10,Amt11,Amt12,Amt13,Amt14,Amt4,Amt5,Amt6,Amt7,Amt8,Nar4,Nar5,VACOCODE,Amt15,USERNAME) " & X1
                    gCn.CommitTrans
            End If
            X1 = "{tmpRptTbl.UserName}='" & gUserName & "' "
            If Trim(txttemp(0)) <> "" Then X1 = X1 + " and {tmpRptTbl.Nar1} = '" & txttemp(0) & "'"
            If Trim(txttemp(5)) <> "" Then X1 = X1 + " and {tmpRptTbl.Nar5} = '" & txttemp(5) & "'"
            If Trim(txttemp(4)) <> "" Then X1 = X1 + " and {tmpRptTbl.Nar8}  = '" & txttemp(4) & "' "
            
           .SelectionFormula = X1
           If ChkUnAdjustRec.Value = 1 Then
               .ReportFileName = gReportPath & "rptTradePurchRegGSTNew.rpt"
           ElseIf ChkConsolidate.Value = 1 Then
               .ReportFileName = gReportPath & "rptTradePurchRegGSTTrpt.rpt"
           Else
               .ReportFileName = gReportPath & "rptTradePurchRegGST.rpt"
           End If
        
       Case "Purchase And Expences Detail (Audit)"
            gCn.BeginTrans
            gCn.Execute "delete from tmpGenTbl where username='" & gUserName & "'"
            gCn.CommitTrans
            X1 = "SELECT tblmastaccount.AcName,tblvoucher.Vtype,tblmastaccount_O.AcGSTIN, TopGroup, iif(sum(vamt)<0,abs(sum(vamt)),0) AS Debit, iif(sum(vamt)>0,sum(vamt),0) AS Credit,'" & gUserName & "' " _
                 & " FROM tblvoucher, tblmastaccount, TblMastGroup,tblmastaccount tblmastaccount_O " _
                 & " Where tblmastaccount.AgCode = TblMastGroup.AgCode And tblvoucher.vacCode = tblmastaccount.AcCode And tblvoucher.vacOCode = tblMastAccount_O.AcCode " _
                 & " and (TopGroup = 90003 or tblmastaccount.Agcode = 90008)and tblvoucher.Vdt>=cdate('" & Format(dtpFrDt.Value, "yyyy,mm,dd") & "')  and tblvoucher.Vdt<=Cdate('" & Format(dtpToDt.Value, "yyyy,mm,dd") & "') " _
                 & " GROUP BY tblmastaccount.AcName,tblvoucher.Vtype,tblmastaccount_O.AcGSTIN, TopGroup " _
                 & " Having Sum(vamt) <> 0 ORDER BY tblmastaccount.AcName "

            gCn.BeginTrans
            gCn.Execute " Insert Into TmpGentbl (Nar1,Nar2,Nar3,Vno,Amt1,Amt2,UserName)" & X1
            gCn.CommitTrans

            gCn.BeginTrans
            gCn.Execute "delete from tmpGenTbl where username='" & gUserName & "' and vno = 90001 and nar2 = '0P'  "
            gCn.CommitTrans




            X1 = "{TmpGentbl.UserName}='" & gUserName & "' "
            .SelectionFormula = X1
            .Formulas(4) = "wFrToDt = 'From : ' & '" & Format(dtpFrDt.Value, "dd/MM/yyyy") & "' & '  To ' & '" & Format(dtpToDt, "dd/MM/yyyy") & "'"
            .Formulas(5) = "wCompGstin = '" & gCGSTIN & "'"
            .ReportFileName = gReportPath & "rptPurchAndExpDetail.rpt"
        Case "Purchase With Payment Detaill (Audit)"
             X = "select tblAddless.vno,tblAddless.vtype,tblAddless.vyear,acname,iif(isnull(AcPAN),'',AcPAN),AcGSTIN,iif(isnull(AcAdd1),'',AcAdd1),iif(isnull(AcAdd2),'',AcAdd2),iif(isnull(AcCity),'',AcCity)," _
                 & " format(tblAddless.Vdt,'DD/MM/YYYY') , tblMastItem.ItName, tblVouSub.vsubBAG, IIf(VSubCtrNo = 1, AdBillAmt, 0), tblMastNarration.Narration,format(tblAddLess.AdGpDt,'DD/MM/YYYY'), " _
                 & " tblMastNarration.GodAdd1,tblMastNarration.GodAdd2,tblMastNarration.GodCity,tblVouSub.vsubWt,'" & gUserName & "' From tblAddless, tblMastAccount, tblVouSub, tblMastItem, tblMastNarration " _
                 & " where  tblAddless.AdAcCrCode = tblMastAccount.AcCode and tblAddless.Vtype = 'PY'  and tblAddless.Vno = tblVouSub.Vno and  tblAddless.Vtype = tblVouSub.Vtype and tblAddless.Vyear = tblVouSub.Vyear " _
                 & " and tblVouSub.VSubItCode = tblMastItem.ItCode and tblAddless.AdTransport = tblMastNarration.NarrCode and tblAddLess.vdt>=cdate('" & Format(dtpFrDt.Value, "yyyy,mm,dd") & "')  and tblAddLess.vdt<=Cdate('" & Format(dtpToDt.Value, "yyyy,mm,dd") & "') " _
                 & " order by tblVouSub.vno,tblVouSub.vtype,tblVouSub.vyear,tblVouSub.VSubCtrNo"
             
             flnm = "D:\Purch-Audit-" + gCCode + "-" + Format(dtpFrDt, "mmm") + ".xlsx"
             FileCopy gReportPath & "\Purch-Audit-Blank.xlsx", flnm
             
             
             Set excel_app = New Excel.Application
             prgBar.Max = 50000 'BarCnt
             prgBar.Value = 0
             prgBar.Visible = True
'             excel_app.Visible = True
             excel_app.Visible = False
             Set workbook = excel_app.Workbooks.Open(filename:=flnm)
             RecYN = GProcRstOpen(tmp1, X, "R", gCn)
             If RecYN > 0 Then
                Set sheet = workbook.Sheets.Item("SUNDRY CREDITORS")
                If (sheet Is Nothing) Then
                    MsgBox "Error SUNDRY CREDITORS Sheet"
                End If
                a = 5
                With tmp1
                     .MoveFirst
                     Do While Not .EOF
                        sheet.Cells(a, 1) = .Fields(3)
                        sheet.Cells(a, 2) = .Fields(4) & vbNullString
                        sheet.Cells(a, 3) = .Fields(5) & vbNullString
                        sheet.Cells(a, 4) = IIf(IsNull(.Fields(6)), "", .Fields(6)) + " " + IIf(IsNull(.Fields(7)), "", .Fields(7)) + " " + IIf(IsNull(.Fields(8)), "", .Fields(8))
                        sheet.Cells(a, 5) = Format(.Fields(9), "MM/DD/YYYY")
                        sheet.Cells(a, 6) = .Fields(10) & vbNullString
                        sheet.Cells(a, 7) = .Fields(11)
                        sheet.Cells(a, 8) = .Fields(12)
                        sheet.Cells(a, 9) = .Fields(13) & vbNullString
                        sheet.Cells(a, 10) = Format(.Fields(14), "MM/DD/YYYY")
                        sheet.Cells(a, 11) = .Fields(13) & vbNullString
                        sheet.Cells(a, 12) = IIf(IsNull(.Fields(15)), "", .Fields(15)) + " " + IIf(IsNull(.Fields(16)), "", .Fields(16)) + " " + IIf(IsNull(.Fields(17)), "", .Fields(17))
                        sheet.Cells(a, 14) = .Fields(18)
                        
                        a = a + 1
                        prgBar.Value = prgBar.Value + 1
                       .MoveNext
                     Loop
                End With
             End If
             
             X = "select acname,iif(isnull(AcPAN),'',AcPAN),iif(isnull(AcAdd1),'',AcAdd1),iif(isnull(AcAdd2),'',AcAdd2),iif(isnull(AcCity),'',AcCity), " _
                 & " sum(iif(vtype = '0P',vamt,0)), sum(  iif(vtype <> '0P'  and vamt > 0,vamt,0)) ,sum(iif(vtype <> '0P'  and vamt < 0,vamt,0)) " _
                 & " from tblvoucher,tblMastAccount,tblMastGroup where vaccode = tblMastAccount.Accode and tblMastAccount.Agcode = tblMastGroup.agcode and (tblMastAccount.AgCode in (90015,90016,90046,90029) or left(tblMastGroup.Agname,11) = 'Transporter') and " _
                 & " tblvoucher.vdt>=cdate('" & Format(dtpFrDt.Value, "yyyy,mm,dd") & "')  and tblvoucher.vdt<=Cdate('" & Format(dtpToDt.Value, "yyyy,mm,dd") & "') " _
                 & " group by acname,iif(isnull(AcPAN),'',AcPAN),iif(isnull(AcAdd1),'',AcAdd1),iif(isnull(AcAdd2),'',AcAdd2),iif(isnull(AcCity),'',AcCity)"

             RecYN = GProcRstOpen(tmp1, X, "R", gCn)
             If RecYN > 0 Then
                Set sheet = workbook.Sheets.Item("SUNDRY CREDITORS all")
                If (sheet Is Nothing) Then
                    MsgBox "Error SUNDRY CREDITORS all Sheet"
                End If
                a = 5
                With tmp1
                     .MoveFirst
                     Do While Not .EOF
                        sheet.Cells(a, 1) = .Fields(0)
                        sheet.Cells(a, 2) = .Fields(1) & vbNullString
                        sheet.Cells(a, 3) = IIf(IsNull(.Fields(2)), "", .Fields(2)) + " " + IIf(IsNull(.Fields(3)), "", .Fields(3)) + " " + IIf(IsNull(.Fields(4)), "", .Fields(4))
                        sheet.Cells(a, 4) = .Fields(5)
'                        If .Fields(5) > 0 Then
'                            sheet.Cells(a, 4) = CStr(.Fields(5)) + " Cr"
'                        Else
'                            sheet.Cells(a, 4) = CStr(.Fields(5)) + " Dr"
'                        End If
                        sheet.Cells(a, 5) = .Fields(6)
                        sheet.Cells(a, 6) = .Fields(7)
                        sheet.Cells(a, 7) = .Fields(5) + .Fields(6) + .Fields(7)
                        
                        a = a + 1
                        prgBar.Value = prgBar.Value + 1
                       .MoveNext
                     Loop
                End With
             End If

             
             
             
             
             
             
             
             
             
             workbook.Close SaveChanges:=True
             Set xlApp = CreateObject("Excel.Application")
             
             xlApp.Visible = True
             Set xlWB = xlApp.Workbooks.Open("" & flnm & "")
             excel_app.Quit
'             MsgBox "Compileted"
             Exit Sub
             
             
             
             
        Case "Credit Note Register GST"
            .Formulas(4) = "wFrToDt='From : '& '" & dtpFrDt.Value & "' & ' To ' & '" & dtpToDt.Value & "'"
            .Formulas(5) = "wCompGstin = '" & gCGSTIN & "'"
            If OptDetail.Value = True Then
               X1 = "{qrytblVoucherSr1.VType}='PN' and {qrytblVoucherSr1.VDrCrNtType} = 'B' and {qrytblVoucherSr1.VDt} in Date(" & Format(dtpFrDt.Value, "yyyy,mm,dd") & ") to Date(" & Format(dtpToDt.Value, "yyyy,mm,dd") & ")  " '
            End If
            If OptShort.Value = True Then
               X1 = "{qrytblVoucherSr1.VType}='PN' and {qrytblVoucherSr1.VDrCrNtType} = 'S' and {qrytblVoucherSr1.VDt} in Date(" & Format(dtpFrDt.Value, "yyyy,mm,dd") & ") to Date(" & Format(dtpToDt.Value, "yyyy,mm,dd") & ")  " '
            End If
            .SelectionFormula = X1
            .ReportFileName = gReportPath & "rptCreditNoteRegGST.rpt"
        Case "Purchase Return Register"
            .Formulas(4) = "wFrToDt='From : '& '" & dtpFrDt.Value & "' & ' To ' & '" & dtpToDt.Value & "'"
            .Formulas(5) = "wCompGstin = '" & gCGSTIN & "'"
            X1 = "{tblAddLess.VType}='VY' and {tblAddLess.VDt} in Date(" & Format(dtpFrDt.Value, "yyyy,mm,dd") & ") to Date(" & Format(dtpToDt.Value, "yyyy,mm,dd") & ")  " '
            .SelectionFormula = X1
            .ReportFileName = gReportPath & "rptTradePurchRtnRegiGST.rpt"
        Case "Sales Return Register"
            .Formulas(4) = "wFrToDt='From : '& '" & dtpFrDt.Value & "' & ' To ' & '" & dtpToDt.Value & "'"
            .Formulas(5) = "wCompGstin = '" & gCGSTIN & "'"
            X1 = "{tblAddLess.VType}='RY' and {tblAddLess.VDt} in Date(" & Format(dtpFrDt.Value, "yyyy,mm,dd") & ") to Date(" & Format(dtpToDt.Value, "yyyy,mm,dd") & ")  " '
            .SelectionFormula = X1
            .ReportFileName = gReportPath & "rptTradeSaleRtnRegiGST.rpt"
        Case "Form GSTR - 3B"
             Dim ig As Integer
             Dim rstTmp As Recordset
             Dim SaleTaxable As Double, SaleIgst As Double, SaleCgst As Double, SaleSgst As Double
             Dim RcmTaxable As Double, RcmIgst As Double, RcmCgst As Double, RcmSgst As Double
             Dim DbNtTaxable As Double, DbNtIgst As Double, DbNtCgst As Double, DbNtSgst As Double
             Dim CrNtTaxable As Double, CrNtIgst As Double, CrNtCgst As Double, CrNtSgst As Double
             Dim RTxableSale As Double, RIgstSale As Double, RCgstSale As Double, RSgstSale As Double
             Dim PurIgst As Double, PurCgst As Double, PurSgst As Double
            
            .Formulas(4) = "wFrToDt='From : '& '" & dtpFrDt.Value & "' & ' To ' & '" & dtpToDt.Value & "'"
            .Formulas(18) = "wCompGstin = '" & gCGSTIN & "'"
            '--Sale
             X1 = "Select sum(tblAddLess.AdSubAmt),sum(tblVouSub.VsubIGSTAmt),sum(tblVouSub.VsubCGSTAmt),sum(tblVouSub.VsubSGSTAmt) " _
                  & " from tblAddLess,tblVouSub where tblAddLess.Vno = tblVouSub.Vno and tblAddLess.Vtype = tblVouSub.Vtype and tblAddLess.Vyear = tblVouSub.Vyear  " _
                  & " and tblAddLess.Vtype in ('SY','SD') and tblAddLess.Vdt>=cdate('" & Format(dtpFrDt.Value, "yyyy,mm,dd") & "')  and tblAddLess.Vdt<=Cdate('" & Format(dtpToDt.Value, "yyyy,mm,dd") & "')"
             ig = GProcRstOpen(rstTmp, X1, "R")
             If ig > 0 Then
                SaleTaxable = Val(rstTmp.Fields(0) & vbNullString)
                SaleIgst = Val(rstTmp.Fields(1) & vbNullString)
                SaleCgst = Val(rstTmp.Fields(2) & vbNullString)
                SaleSgst = Val(rstTmp.Fields(3) & vbNullString)
             End If
             DoEvents
            '--Sales Rtn
            X1 = "Select sum(tblVouSub.VsubIGSTAmt),sum(tblVouSub.VsubCGSTAmt),sum(tblVouSub.VsubSGSTAmt),sum(tblAddLess.AdSubAmt) from tblAddLess,tblVouSub where tblAddLess.Vno = tblVouSub.Vno and tblAddLess.Vtype = tblVouSub.Vtype and tblAddLess.Vyear = tblVouSub.Vyear " _
                 & " and tblAddLess.Vtype = 'RY' and tblAddLess.vdt>=cdate('" & Format(dtpFrDt.Value, "yyyy,mm,dd") & "')  and tblAddLess.Vdt<=Cdate('" & Format(dtpToDt.Value, "yyyy,mm,dd") & "') "
             ig = GProcRstOpen(rstTmp, X1, "R")
             If ig > 0 Then
                SaleTaxable = SaleTaxable - Val(rstTmp.Fields(3) & vbNullString)
                SaleIgst = SaleIgst - Val(rstTmp.Fields(0) & vbNullString)
                SaleCgst = SaleCgst - Val(rstTmp.Fields(1) & vbNullString)
                SaleSgst = SaleSgst - Val(rstTmp.Fields(2) & vbNullString)
             End If
             DoEvents
             '-- Sale RCM
             X1 = "Select sum(tblAddLess.AdBillAmt),sum(tblAddLess.AdLCostAdd),sum(tblAddLess.AdLCostFreight),sum(tblAddLess.AdLCostGRent) " _
                  & " from tblAddLess where tblAddLess.Vtype = 'PO' and ADTMP2 = 1 and tblAddLess.AdBookDt>=cdate('" & Format(dtpFrDt.Value, "yyyy,mm,dd") & "')  and tblAddLess.AdBookDt<=Cdate('" & Format(dtpToDt.Value, "yyyy,mm,dd") & "')"
             ig = GProcRstOpen(rstTmp, X1, "R")
             If ig > 0 Then
                RcmTaxable = Val(rstTmp.Fields(0) & vbNullString)
                RcmIgst = Val(rstTmp.Fields(1) & vbNullString)
                RcmCgst = Val(rstTmp.Fields(2) & vbNullString)
                RcmSgst = Val(rstTmp.Fields(3) & vbNullString)
             End If
             DoEvents
             '---Debit Note Sales
             X1 = "Select sum(VTaxableAmt),sum(abs(VIGSTAmt)),sum(abs(VCGSTAmt)),sum(abs(VSGSTAmt)) " _
                  & " from qrytblVoucherSr1 where qrytblVoucherSr1.Vtype = 'SN' and qrytblVoucherSr1.VDrCrNtType = 'B'  and VCGSTAmt+VSGSTAmt+VIGSTAmt > 0 and qrytblVoucherSr1.Vdt>=cdate('" & Format(dtpFrDt.Value, "yyyy,mm,dd") & "')  and qrytblVoucherSr1.Vdt<=Cdate('" & Format(dtpToDt.Value, "yyyy,mm,dd") & "')"
             ig = GProcRstOpen(rstTmp, X1, "R")
             If ig > 0 Then
                DbNtTaxable = Val(rstTmp.Fields(0) & vbNullString)
                DbNtIgst = Val(rstTmp.Fields(1) & vbNullString)
                DbNtCgst = Val(rstTmp.Fields(2) & vbNullString)
                DbNtSgst = Val(rstTmp.Fields(3) & vbNullString)
             End If
             DoEvents
             '---Credit Note Sales
             X1 = "Select sum(VTaxableAmt),sum(abs(VIGSTAmt)),sum(abs(VCGSTAmt)),sum(abs(VSGSTAmt)) " _
                  & " from qrytblVoucherSr1 where qrytblVoucherSr1.Vtype = 'PN' and qrytblVoucherSr1.VDrCrNtType = 'B' and VCGSTAmt+VSGSTAmt+VIGSTAmt > 0 and qrytblVoucherSr1.Vdt>=cdate('" & Format(dtpFrDt.Value, "yyyy,mm,dd") & "')  and qrytblVoucherSr1.Vdt<=Cdate('" & Format(dtpToDt.Value, "yyyy,mm,dd") & "')"
             ig = GProcRstOpen(rstTmp, X1, "R")
             If ig > 0 Then
                CrNtTaxable = Val(rstTmp.Fields(0) & vbNullString)
                CrNtIgst = Val(rstTmp.Fields(1) & vbNullString)
                CrNtCgst = Val(rstTmp.Fields(2) & vbNullString)
                CrNtSgst = Val(rstTmp.Fields(3) & vbNullString)
             End If
             DoEvents
            '--Purchase
            X1 = "Select sum(tblVouSub.VsubIGSTAmt),sum(tblVouSub.VsubCGSTAmt),sum(tblVouSub.VsubSGSTAmt) " _
                 & " from tblAddLess,tblVouSub where tblAddLess.Vno = tblVouSub.Vno and tblAddLess.Vtype = tblVouSub.Vtype and tblAddLess.Vyear = tblVouSub.Vyear " _
                 & " and tblAddLess.Vtype = 'PY' and tblAddLess.AdGpDt>=cdate('" & Format(dtpFrDt.Value, "yyyy,mm,dd") & "')  and tblAddLess.AdGpDt<=Cdate('" & Format(dtpToDt.Value, "yyyy,mm,dd") & "') "
             ig = GProcRstOpen(rstTmp, X1, "R")
             If ig > 0 Then
                PurIgst = Val(rstTmp.Fields(0) & vbNullString)
                PurCgst = Val(rstTmp.Fields(1) & vbNullString)
                PurSgst = Val(rstTmp.Fields(2) & vbNullString)
             End If
             DoEvents
            '--Purchase Other No RCM
            X1 = "Select sum(tblVouSub.VsubIGSTAmt),sum(tblVouSub.VsubCGSTAmt),sum(tblVouSub.VsubSGSTAmt) " _
                 & " from tblAddLess,tblVouSub where tblAddLess.Vno = tblVouSub.Vno and tblAddLess.Vtype = tblVouSub.Vtype and tblAddLess.Vyear = tblVouSub.Vyear " _
                 & " and tblAddLess.Vtype = 'PO' and tblAddLess.ADTMP2 = 0  and tblAddLess.AdGpDt>=cdate('" & Format(dtpFrDt.Value, "yyyy,mm,dd") & "')  and tblAddLess.AdGpDt<=Cdate('" & Format(dtpToDt.Value, "yyyy,mm,dd") & "') "
             
             
             ig = GProcRstOpen(rstTmp, X1, "R")
             If ig > 0 Then
                PurIgst = PurIgst + Val(rstTmp.Fields(0) & vbNullString)
                PurCgst = PurCgst + Val(rstTmp.Fields(1) & vbNullString)
                PurSgst = PurSgst + Val(rstTmp.Fields(2) & vbNullString)
             End If
             
             '---Debit Note Purchase (-)
             X1 = "Select sum(VTaxableAmt),sum(abs(VIGSTAmt)),sum(abs(VCGSTAmt)),sum(abs(VSGSTAmt)) " _
                  & " from qrytblVoucherSr1 where qrytblVoucherSr1.Vtype = 'SN' and qrytblVoucherSr1.VDrCrNtType = 'S'  and VCGSTAmt+VSGSTAmt+VIGSTAmt > 0 and qrytblVoucherSr1.Vdt>=cdate('" & Format(dtpFrDt.Value, "yyyy,mm,dd") & "')  and qrytblVoucherSr1.Vdt<=Cdate('" & Format(dtpToDt.Value, "yyyy,mm,dd") & "')"
             ig = GProcRstOpen(rstTmp, X1, "R")
             If ig > 0 Then
                PurIgst = PurIgst - Val(rstTmp.Fields(1) & vbNullString)
                PurCgst = PurCgst - Val(rstTmp.Fields(2) & vbNullString)
                PurSgst = PurSgst - Val(rstTmp.Fields(3) & vbNullString)
             End If
             DoEvents
             '---Credit Note Purchase (+)
             X1 = "Select sum(VTaxableAmt),sum(abs(VIGSTAmt)),sum(abs(VCGSTAmt)),sum(abs(VSGSTAmt)) " _
                  & " from qrytblVoucherSr1 where qrytblVoucherSr1.Vtype = 'PN' and qrytblVoucherSr1.VDrCrNtType = 'S' and VCGSTAmt+VSGSTAmt+VIGSTAmt > 0 and qrytblVoucherSr1.Vdt>=cdate('" & Format(dtpFrDt.Value, "yyyy,mm,dd") & "')  and qrytblVoucherSr1.Vdt<=Cdate('" & Format(dtpToDt.Value, "yyyy,mm,dd") & "')"
             ig = GProcRstOpen(rstTmp, X1, "R")
             If ig > 0 Then
                PurIgst = PurIgst + Val(rstTmp.Fields(1) & vbNullString)
                PurCgst = PurCgst + Val(rstTmp.Fields(2) & vbNullString)
                PurSgst = PurSgst + Val(rstTmp.Fields(3) & vbNullString)
             End If
             
             DoEvents
            '--Purch Rtn
            X1 = "Select sum(tblVouSub.VsubIGSTAmt),sum(tblVouSub.VsubCGSTAmt),sum(tblVouSub.VsubSGSTAmt) " _
                 & " from tblAddLess,tblVouSub where tblAddLess.Vno = tblVouSub.Vno and tblAddLess.Vtype = tblVouSub.Vtype and tblAddLess.Vyear = tblVouSub.Vyear " _
                 & " and tblAddLess.Vtype = 'VY' and tblAddLess.vDt>=cdate('" & Format(dtpFrDt.Value, "yyyy,mm,dd") & "')  and tblAddLess.vDt<=Cdate('" & Format(dtpToDt.Value, "yyyy,mm,dd") & "') "
             ig = GProcRstOpen(rstTmp, X1, "R")
             If ig > 0 Then
                PurIgst = PurIgst - Val(rstTmp.Fields(0) & vbNullString)
                PurCgst = PurCgst - Val(rstTmp.Fields(1) & vbNullString)
                PurSgst = PurSgst - Val(rstTmp.Fields(2) & vbNullString)
             End If
             DoEvents
             
             RTxableSale = SaleTaxable + DbNtTaxable - CrNtTaxable
             RIgstSale = SaleIgst + DbNtIgst - CrNtIgst
             RCgstSale = SaleCgst + DbNtCgst - CrNtCgst
             RSgstSale = SaleSgst + DbNtSgst - CrNtSgst
             
             
            .Formulas(5) = "wRTxableSale= " & RTxableSale & ""
            .Formulas(6) = "wRIgstSale= " & RIgstSale & ""
            .Formulas(7) = "wRCgstSale= " & RCgstSale & ""
            .Formulas(8) = "wRSgstSale= " & RSgstSale & ""
             
            .Formulas(9) = "wRcmTaxable= " & RcmTaxable & ""
            .Formulas(10) = "wRcmIgst= " & RcmIgst & ""
            .Formulas(11) = "wRcmCgst= " & RcmCgst & ""
            .Formulas(12) = "wRcmSgst= " & RcmSgst & ""
             
            .Formulas(13) = "wPurIgst= " & PurIgst & ""
            .Formulas(14) = "wPurCgst= " & PurCgst & ""
            .Formulas(15) = "wPurSgst= " & PurSgst & ""
            gCn.BeginTrans
            gCn.Execute "delete from tmpRptTbl Where Username='" & gUserName & "'"
            gCn.CommitTrans
            X1 = "{tmpRptTbl.UserName}='" & gUserName & "' "
            .SelectionFormula = X1
            .ReportFileName = gReportPath & "rptFormGSTR-3B.rpt"
        Case "Insurance Report":
            .Formulas(4) = "wFrToDt='From : '& '" & dtpFrDt.Value & "' & ' To ' & '" & dtpToDt.Value & "'"
            X1 = "({tblAddLess.VType}='PT' or {tblAddLess.VType}='PY') and {tblAddLess.VDt} in Date(" & Format(dtpFrDt.Value, "yyyy,mm,dd") & ") to Date(" & Format(dtpToDt.Value, "yyyy,mm,dd") & ")  " '
            If Trim(txttemp(1)) <> "" Then '-Tax
                 X1 = X1 + " and {tblAddLess.AdTaxCode} = " & wNarrCode & ""
                .Formulas(5) = "wTaxType='Tax Type : ' & '" & txttemp(1) & "'"
            End If
            If Trim(txttemp(0)) <> "" Then '-Mill
                X1 = X1 + " and {tblAddLess.AdMillCode} = " & wMillCode & ""
                .Formulas(6) = "wMillName='Mill Name : ' & '" & txttemp(0) & "'"
            End If
            '-- Party
            If txttemp(2) <> "" Then X1 = X1 + " and {tblAddLess.AdAcCrCode} = " & wPartyCode
           
            .Formulas(8) = "InsAmt='" & Val(txt(1)) & "'"
            
           .ReportFileName = gReportPath & "rptInsurance.rpt"
           .SelectionFormula = X1
        Case "Other Purchase Register GST"
            .Formulas(4) = "wFrToDt='From : '& '" & dtpFrDt.Value & "' & ' To ' & '" & dtpToDt.Value & "'"
            .Formulas(5) = "wCompGstin = '" & gCGSTIN & "'"
            X1 = "Select tblAddLess.Vno,tblAddLess.Vdt,tblAddLess.AdGpDt,tblAddLess.AdInvNo,tblMastAccount.AcName,tblMastAccount.AcGSTIN,tblMastAccountP.AcName,tblAddLess.AdSubAmt,tblAddLess.AdRoff,tblAddLess.AdBillAmt,tblAddLess.ADTMP2,tblAddLess.ADTMP5,tblAddLess.AdBookDt,sum(tblVouSub.VsubCGSTAmt), " _
                 & " sum(tblVouSub.VsubSGSTAmt),sum(tblVouSub.VsubIGSTAmt),'" & gUserName & "' " _
                 & " from tblAddLess,tblVouSub,tblMastAccount,tblMastAccount tblMastAccountp where tblAddLess.Vno = tblVouSub.Vno and tblAddLess.Vtype = tblVouSub.Vtype and tblAddLess.Vyear = tblVouSub.Vyear and tblAddLess.AdAcCrCode = tblMastAccount.Accode and tblAddLess.AdAcDrCode = tblMastAccountP.Accode" _
                 & " and tblAddLess.Vtype = 'PO' and tblAddLess.ADTMP2 = 0 and tblAddLess.AdGpDt>=cdate('" & Format(dtpFrDt.Value, "yyyy,mm,dd") & "')  and tblAddLess.AdGpDt<=Cdate('" & Format(dtpToDt.Value, "yyyy,mm,dd") & "') GROUP BY tblAddLess.Vno,tblAddLess.AdGpDt,tblAddLess.Vdt,tblAddLess.AdInvNo,tblMastAccount.AcName,tblMastAccountp.AcName,tblMastAccount.AcGSTIN,tblAddLess.AdSubAmt,tblAddLess.AdRoff,tblAddLess.AdBillAmt,tblAddLess.ADTMP2,tblAddLess.ADTMP5,tblAddLess.AdBookDt"
            gCn.BeginTrans
            gCn.Execute "delete from tmpRptTbl Where Username='" & gUserName & "'"
            gCn.Execute "insert into tmpRptTbl (Vno,Vdt,VBillDt,VBillNo,Nar1,Nar2,Nar4,Amt1,Amt2,Amt6,Amt7,Nar3,Dt1,Amt3,Amt4,Amt5,USERNAME) " & X1
            gCn.CommitTrans
            X1 = "{tmpRptTbl.UserName}='" & gUserName & "' "
            .SelectionFormula = X1
            .ReportFileName = gReportPath & "rptOtherPurchRegGST.rpt"
        Case "GSTR - 9"
             GProcPurpareGSTR9 dtpFrDt.Value, dtpToDt.Value, OptDetail.Value, OptShort.Value
             Exit Sub
            
        Case "Form GSTR - 1"
             If OptDetail.Value = True Then
                GProcPurpareGSTR1_V21 dtpFrDt.Value, dtpToDt.Value, "E"
             Else
                GProcPurpareGSTR1_V21 dtpFrDt.Value, dtpToDt.Value, "J"
             End If
             Exit Sub
''             X1 = " tblAddLess.VType in ('SY')  and tblAddLess.vdt>=cdate('" & Format(dtpFrDt.Value, "yyyy,mm,dd") & "')  and tblAddLess.vdt<=Cdate('" & Format(dtpToDt.Value, "yyyy,mm,dd") & "') and isnull(tblmastaccount.AcGSTIN)"  '
''             X1 = "Select Distinct AcName,'" & gUserName & "' from tblAddLess,tblMastAccount where tblAddLess.AdAcDrCode = tblMastAccount.Accode  and " & X1
''             If GProcRstOpen(tmp1, X1, "R", gCn) > 0 Then
''                gCn.BeginTrans
''                gCn.Execute "delete from tmpGenTbl where username='" & gUserName & "'"
''                gCn.Execute " Insert Into TmpGentbl (Nar1,UserName) " & X1
''                gCn.CommitTrans
''                X1 = "{TmpGentbl.UserName}='" & gUserName & "'"
''               .SelectionFormula = X1
''               .Formulas(4) = "wFrToDt = 'From : ' & '" & Format(dtpFrDt.Value, "dd/MM/yyyy") & "' & '  To ' & '" & Format(dtpToDt, "dd/MM/yyyy") & "'"
''               .ReportFileName = gReportPath & "rptGstinCheck.rpt"
''                .Action = 1
''                Exit Sub
''             End If
''             '---FOr BarCount
''             gCn.BeginTrans
''             gCn.Execute "delete from tmpGenTbl where username='" & gUserName & "'"
''             '-B2B
''             X1 = " tblAddLess.VType in ('SY')  and tblAddLess.vdt>=cdate('" & Format(dtpFrDt.Value, "yyyy,mm,dd") & "')  and tblAddLess.vdt<=Cdate('" & Format(dtpToDt.Value, "yyyy,mm,dd") & "') and len(tblmastaccount.AcGSTIN) > 3 "  '
''             gCn.Execute " Insert Into TmpGentbl (vno,Nar6,vyear,Nar1,vtype,VbillNo,Vdt,Amt1,Nar4,Nar5,Nar2,Nar3,Amt2,Amt3,Amt4,UserName)" _
''                         & " select tblAddLess.vno,tblAddLess.vtype,tblAddLess.vyear,ucase(tblmastaccount.AcGSTIN),left(tblmastaccount.AcGSTIN,2),tblAddLess.AdInvNo,tblAddLess.Vdt,tblAddLess.AdBillAmt,tblMastState.StateName, 'N','Regular','', iif(tblVouSub.VsubIGSTRt>0,tblVouSub.VsubIGSTRt,tblVouSub.VsubCGSTRt+tblVouSub.VsubSGSTRt),sum(tblAddLess.AdSubAmt),0,'" & gUserName & "'" _
''                         & " From tblmastaccount,tblAddLess,tblVouSub,tblMastState where tblAddLess.AdAcDrCode = tblMastAccount.Accode and tblAddLess.Vno = tblVouSub.Vno and tblAddLess.Vtype = tblVouSub.Vtype and tblAddLess.Vyear = tblVouSub.Vyear and left(tblMastAccount.AcGSTIN,2) = tblMastState.StateCd  and " & X1 & "" _
''                         & " group by tblAddLess.vno,tblAddLess.vtype,tblAddLess.vyear,tblmastaccount.AcGSTIN,AdInvNo,tblAddLess.Vdt,tblAddLess.AdBillAmt,tblMastState.StateName,iif(tblVouSub.VsubIGSTRt>0,tblVouSub.VsubIGSTRt,tblVouSub.VsubCGSTRt+tblVouSub.VsubSGSTRt)"
''             '--B2CL
''             X1 = " tblAddLess.VType in ('SY')  and tblAddLess.vdt>=cdate('" & Format(dtpFrDt.Value, "yyyy,mm,dd") & "')  and tblAddLess.vdt<=Cdate('" & Format(dtpToDt.Value, "yyyy,mm,dd") & "') and (len(tblmastaccount.AcGSTIN) <= 3 or isnull(tblmastaccount.AcGSTIN)) and tblAddLess.AdBillAmt >= 250000"  '
''             gCn.Execute " Insert Into TmpGentbl (vno,Nar6,vyear,Nar1,vtype,VbillNo,Vdt,Amt1,Nar4,Nar5,Nar2,Nar3,Amt2,Amt3,Amt4,UserName)" _
''                         & " select tblAddLess.vno,tblAddLess.vtype,tblAddLess.vyear,ucase(tblmastaccount.AcGSTIN),left(tblmastaccount.AcGSTIN,2),tblAddLess.AdInvNo,tblAddLess.Vdt,tblAddLess.AdBillAmt,tblMastState.StateName, 'N','Regular','', iif(tblVouSub.VsubIGSTRt>0,tblVouSub.VsubIGSTRt,tblVouSub.VsubCGSTRt+tblVouSub.VsubSGSTRt),sum(tblVouSub.VSubAmt),0,'" & gUserName & "'" _
''                         & " From tblmastaccount,tblAddLess,tblVouSub,tblMastState where tblAddLess.AdAcDrCode = tblMastAccount.Accode and tblAddLess.Vno = tblVouSub.Vno and tblAddLess.Vtype = tblVouSub.Vtype and tblAddLess.Vyear = tblVouSub.Vyear and left(tblMastAccount.AcGSTIN,2) = tblMastState.StateCd  and " & X1 & "" _
''                         & " group by tblAddLess.vno,tblAddLess.vtype,tblAddLess.vyear,tblmastaccount.AcGSTIN,AdInvNo,tblAddLess.Vdt,tblAddLess.AdBillAmt,tblMastState.StateName,iif(tblVouSub.VsubIGSTRt>0,tblVouSub.VsubIGSTRt,tblVouSub.VsubCGSTRt+tblVouSub.VsubSGSTRt)"
''             '---B2CS
''             X1 = " tblAddLess.VType in ('SY')  and tblAddLess.vdt>=cdate('" & Format(dtpFrDt.Value, "yyyy,mm,dd") & "')  and tblAddLess.vdt<=Cdate('" & Format(dtpToDt.Value, "yyyy,mm,dd") & "') and (len(tblmastaccount.AcGSTIN) <= 3 or isnull(tblmastaccount.AcGSTIN)) and tblAddLess.AdBillAmt < 250000"  '
''             gCn.Execute " Insert Into TmpGentbl (vno,Nar6,vyear,Nar1,vtype,VbillNo,Vdt,Amt1,Nar4,Nar5,Nar2,Nar3,Amt2,Amt3,Amt4,UserName)" _
''                         & " select tblAddLess.vno,tblAddLess.vtype,tblAddLess.vyear,ucase(tblmastaccount.AcGSTIN),left(tblmastaccount.AcGSTIN,2),tblAddLess.AdInvNo,tblAddLess.Vdt,tblAddLess.AdBillAmt,tblMastState.StateName, 'N','Regular','', iif(tblVouSub.VsubIGSTRt>0,tblVouSub.VsubIGSTRt,tblVouSub.VsubCGSTRt+tblVouSub.VsubSGSTRt),sum(tblVouSub.VSubAmt),0,'" & gUserName & "'" _
''                         & " From tblmastaccount,tblAddLess,tblVouSub,tblMastState where tblAddLess.AdAcDrCode = tblMastAccount.Accode and tblAddLess.Vno = tblVouSub.Vno and tblAddLess.Vtype = tblVouSub.Vtype and tblAddLess.Vyear = tblVouSub.Vyear and left(tblMastAccount.AcGSTIN,2) = tblMastState.StateCd  and " & X1 & "" _
''                         & " group by tblAddLess.vno,tblAddLess.vtype,tblAddLess.vyear,tblmastaccount.AcGSTIN,AdInvNo,tblAddLess.Vdt,tblAddLess.AdBillAmt,tblMastState.StateName,iif(tblVouSub.VsubIGSTRt>0,tblVouSub.VsubIGSTRt,tblVouSub.VsubCGSTRt+tblVouSub.VsubSGSTRt)"
''             '--CDNR
''             X1 = " qrytblVoucherSr1.VType in ('SN','PN')  and qrytblVoucherSr1.vdt>=cdate('" & Format(dtpFrDt.Value, "yyyy,mm,dd") & "')  and qrytblVoucherSr1.vdt<=Cdate('" & Format(dtpToDt.Value, "yyyy,mm,dd") & "') "  '
''             gCn.Execute " Insert Into TmpGentbl (Nar1,Vtype,VBILLNO,VBILLDT,VNO,Vdt,Nar2,Nar3,Nar5,Amt1,Amt2,Amt3,Amt4,Nar4,UserName)" _
''                         & " Select tblmastaccount.AcGSTIN,left(tblmastaccount.AcGSTIN,2),qrytblVoucherSr1.VBillNo,qrytblVoucherSr1.VBillDt,qrytblVoucherSr1.Vno,qrytblVoucherSr1.Vdt,iif(qrytblVoucherSr1.vtype = 'SN','D','C'),'07-Others',tblMastState.StateName,abs(qrytblVoucherSr1.VAmt),qrytblVoucherSr1.VCGSTRt+qrytblVoucherSr1.VSGSTRt+qrytblVoucherSr1.VIGSTRt,qrytblVoucherSr1.VTaxableAmt,0," _
''                         & " iif(qrytblVoucherSr1.VBillDt < cdate('2017/07/01'),'Y','N'),'" & gUserName & "' from qrytblVoucherSr1,tblMastAccount,tblMastState where qrytblVoucherSr1.VacCode = tblMastAccount.AcCode and left(tblMastAccount.AcGSTIN,2) = tblMastState.StateCd   and " & X1 & " "
''             '--HSN
''             X1 = " tblAddLess.VType in ('SY')  and tblAddLess.vdt>=cdate('" & Format(dtpFrDt.Value, "yyyy,mm,dd") & "')  and tblAddLess.vdt<=Cdate('" & Format(dtpToDt.Value, "yyyy,mm,dd") & "') "  '
''             gCn.Execute " Insert Into TmpGentbl (Nar1,Nar2,Nar3,Amt1,Amt2,Amt3,Amt4,Amt5,Amt6,Amt7,UserName)" _
''                         & " select tblMastItem.ITTMP2,iif(tblMastItem.ItOtherItem=0,'Yarn',''),tblMastItem.ItUnit,sum(tblVouSub.VSubBag),sum(tblAddLess.AdBillAmt),sum(tblVouSub.VSubAmt),sum(tblVouSub.VsubIGSTAmt),sum(tblVouSub.VsubCGSTAmt),sum(tblVouSub.VsubSGSTAmt),0,'" & gUserName & "'" _
''                         & " From tblAddLess,tblVouSub,tblMastItem where tblVouSub.VSubItCode = tblMastItem.ItCode and tblAddLess.Vno = tblVouSub.Vno and tblAddLess.Vtype = tblVouSub.Vtype and tblAddLess.Vyear = tblVouSub.Vyear and " & X1 & "" _
''                         & " group by tblMastItem.ITTMP2,iif(tblMastItem.ItOtherItem=0,'Yarn',''),tblMastItem.ItUnit"
''             DoEvents
''             DoEvents
''             DoEvents
''             BarCnt = GProcRstOpen(tmp1, "select * from TmpGentbl where TmpGentbl.UserName ='" & gUserName & "'", "R", gCn)
''             BarCnt = BarCnt * 1.7
''             '---End Bar Count
''             flnm = "D:\GSTR1-" + gCCode + "-" + Format(dtpFrDt, "mmm") + ".xlsx"
''             FileCopy gReportPath & "\GSTR1-Blank-1.8.xlsx", flnm
''             gCn.BeginTrans
''             gCn.Execute "delete from tmpGenTbl where username='" & gUserName & "'"
''             '---Sales Detail
''             '--B2B
''             X1 = " tblAddLess.VType in ('SY')  and tblAddLess.vdt>=cdate('" & Format(dtpFrDt.Value, "yyyy,mm,dd") & "')  and tblAddLess.vdt<=Cdate('" & Format(dtpToDt.Value, "yyyy,mm,dd") & "') and len(tblmastaccount.AcGSTIN) > 3 "  '
''             gCn.Execute " Insert Into TmpGentbl (vno,Nar6,vyear,Nar1,vtype,Nar7,VbillNo,Vdt,Amt1,Nar4,Nar5,Nar2,Nar3,Amt2,Amt3,Amt4,UserName,Amt5)" _
''                         & " select tblAddLess.vno,tblAddLess.vtype,tblAddLess.vyear,ucase(tblmastaccount.AcGSTIN),left(tblmastaccount.AcGSTIN,2),tblMastAccount.Acname,tblAddLess.AdInvNo,tblAddLess.Vdt,tblAddLess.AdBillAmt-tblAddLess.AdTcsAmt,tblMastState.StateName, 'N','Regular','', iif(tblVouSub.VsubIGSTRt>0,tblVouSub.VsubIGSTRt,tblVouSub.VsubCGSTRt+tblVouSub.VsubSGSTRt),sum(tblAddLess.AdSubAmt),0,'" & gUserName & "',val(mid(tblAddLess.AdInvNo,5,10))" _
''                         & " From tblmastaccount,tblAddLess,tblVouSub,tblMastState where tblAddLess.AdAcDrCode = tblMastAccount.Accode and tblAddLess.Vno = tblVouSub.Vno and tblAddLess.Vtype = tblVouSub.Vtype and tblAddLess.Vyear = tblVouSub.Vyear and left(tblMastAccount.AcGSTIN,2) = tblMastState.StateCd  and " & X1 & "" _
''                         & " group by tblAddLess.vno,tblAddLess.vtype,tblAddLess.vyear,tblmastaccount.AcGSTIN,tblMastAccount.Acname,AdInvNo,tblAddLess.Vdt,tblAddLess.AdBillAmt-tblAddLess.AdTcsAmt,tblMastState.StateName,iif(tblVouSub.VsubIGSTRt>0,tblVouSub.VsubIGSTRt,tblVouSub.VsubCGSTRt+tblVouSub.VsubSGSTRt)"
''
''             gCn.CommitTrans
''             Set excel_app = New excel.Application
''             prgBar.Max = BarCnt
''             prgBar.Value = 0
''             prgBar.Visible = True
'''             excel_app.Visible = True
''             excel_app.Visible = False
''             Set workbook = excel_app.Workbooks.Open(filename:=flnm)
''             RecYN = GProcRstOpen(tmp1, "select * from TmpGentbl where TmpGentbl.UserName ='" & gUserName & "' order by Vdt,Amt5 ", "R", gCn)
''             If RecYN > 0 Then
''                Set sheet = workbook.Sheets.Item("b2b")
''                If (sheet Is Nothing) Then
''                    MsgBox "Error Sales Sheet"
''                End If
''                a = 5
''                With tmp1
''                     .MoveFirst
''                     Do While Not .EOF
''                        sheet.Cells(a, 1) = !Nar1
''                        sheet.Cells(a, 2) = !Nar7
''                        sheet.Cells(a, 3) = !VBILLNO
''                        DoEvents
''                        sheet.Cells(a, 4) = Format(!Vdt, "dd-MMM-YYYY")
''                        sheet.Cells(a, 5) = !Amt1
''                        sheet.Cells(a, 6) = !Nar4
''                        sheet.Cells(a, 7) = "N"
''                        sheet.Cells(a, 8) = ""
''                        sheet.Cells(a, 9) = "Regular"
''                        sheet.Cells(a, 10) = ""
''                        sheet.Cells(a, 11) = !AMT2
''                        sheet.Cells(a, 12) = !AMT3
''                        sheet.Cells(a, 13) = 0
''                        a = a + 1
''                        prgBar.Value = prgBar.Value + 1
''                       .MoveNext
''                     Loop
''                End With
''             End If
''             '--B2CL
''             gCn.BeginTrans
''             gCn.Execute "delete from tmpGenTbl where username='" & gUserName & "'"
''             X1 = " tblAddLess.VType in ('SY')  and tblAddLess.vdt>=cdate('" & Format(dtpFrDt.Value, "yyyy,mm,dd") & "')  and tblAddLess.vdt<=Cdate('" & Format(dtpToDt.Value, "yyyy,mm,dd") & "') and (len(tblmastaccount.AcGSTIN) <= 3 or isnull(tblmastaccount.AcGSTIN)) and tblAddLess.AdBillAmt >= 250000"  '
''             gCn.Execute " Insert Into TmpGentbl (vno,Nar6,vyear,Nar1,vtype,VbillNo,Vdt,Amt1,Nar4,Nar5,Nar2,Nar3,Amt2,Amt3,Amt4,UserName)" _
''                         & " select tblAddLess.vno,tblAddLess.vtype,tblAddLess.vyear,ucase(tblmastaccount.AcGSTIN),left(tblmastaccount.AcGSTIN,2),tblAddLess.AdInvNo,tblAddLess.Vdt,tblAddLess.AdBillAmt,tblMastState.StateName, 'N','Regular','', iif(tblVouSub.VsubIGSTRt>0,tblVouSub.VsubIGSTRt,tblVouSub.VsubCGSTRt+tblVouSub.VsubSGSTRt),sum(tblVouSub.VSubAmt),0,'" & gUserName & "'" _
''                         & " From tblmastaccount,tblAddLess,tblVouSub,tblMastState where tblAddLess.AdAcDrCode = tblMastAccount.Accode and tblAddLess.Vno = tblVouSub.Vno and tblAddLess.Vtype = tblVouSub.Vtype and tblAddLess.Vyear = tblVouSub.Vyear and left(tblMastAccount.AcGSTIN,2) = tblMastState.StateCd  and " & X1 & "  group by tblAddLess.vno,tblAddLess.vtype,tblAddLess.vyear,tblmastaccount.AcGSTIN,AdInvNo,tblAddLess.Vdt,tblAddLess.AdBillAmt,tblMastState.StateName,iif(tblVouSub.VsubIGSTRt>0,tblVouSub.VsubIGSTRt,tblVouSub.VsubCGSTRt+tblVouSub.VsubSGSTRt)"
''             gCn.CommitTrans
''             RecYN = GProcRstOpen(tmp1, "select * from TmpGentbl where TmpGentbl.UserName ='" & gUserName & "' order by Vbillno,Vdt", "R", gCn)
''             If RecYN > 0 Then
''                Set sheet = workbook.Sheets.Item("b2cl")
''                If (sheet Is Nothing) Then
''                    MsgBox "Error Sales Sheet"
''                End If
''                a = 5
''                With tmp1
''                     .MoveFirst
''                     Do While Not .EOF
''                        sheet.Cells(a, 1) = !VBILLNO
''                        DoEvents
''                        sheet.Cells(a, 2) = Format(!Vdt, "dd-MMM-YYYY")
''                        sheet.Cells(a, 3) = !Amt1
''                        sheet.Cells(a, 4) = !Nar4
''                        sheet.Cells(a, 5) = ""
''                        sheet.Cells(a, 6) = !AMT2
''                        sheet.Cells(a, 7) = !AMT3
''                        sheet.Cells(a, 8) = 0
''                        a = a + 1
''                        prgBar.Value = prgBar.Value + 1
''                       .MoveNext
''                     Loop
''                End With
''             End If
''             '--B2CS
''             gCn.BeginTrans
''             gCn.Execute "delete from tmpGenTbl where username='" & gUserName & "'"
''             X1 = " tblAddLess.VType in ('SY')  and tblAddLess.vdt>=cdate('" & Format(dtpFrDt.Value, "yyyy,mm,dd") & "')  and tblAddLess.vdt<=Cdate('" & Format(dtpToDt.Value, "yyyy,mm,dd") & "') and (len(tblmastaccount.AcGSTIN) <= 3 or isnull(tblmastaccount.AcGSTIN)) and tblAddLess.AdBillAmt < 250000"  '
''             gCn.Execute " Insert Into TmpGentbl (Nar4,Nar5,Nar2,Amt2,Amt3,Amt4,UserName)" _
''                         & " select ucase(tblmastaccount.AcGSTIN),left(tblmastaccount.AcGSTIN,2),tblMastState.StateName,iif(tblVouSub.VsubIGSTRt>0,tblVouSub.VsubIGSTRt,tblVouSub.VsubCGSTRt+tblVouSub.VsubSGSTRt),sum(tblVouSub.VSubAmt),0,'" & gUserName & "'" _
''                         & " From tblmastaccount,tblAddLess,tblVouSub,tblMastState where tblAddLess.AdAcDrCode = tblMastAccount.Accode and tblAddLess.Vno = tblVouSub.Vno and tblAddLess.Vtype = tblVouSub.Vtype and tblAddLess.Vyear = tblVouSub.Vyear and left(tblMastAccount.AcGSTIN,2) = tblMastState.StateCd  and " & X1 & " group by ucase(tblmastaccount.AcGSTIN),left(tblmastaccount.AcGSTIN,2),tblMastState.StateName,iif(tblVouSub.VsubIGSTRt>0,tblVouSub.VsubIGSTRt,tblVouSub.VsubCGSTRt+tblVouSub.VsubSGSTRt)"
''             gCn.CommitTrans
''             If GProcRstOpen(tmp1, "select * from TmpGentbl where TmpGentbl.UserName ='" & gUserName & "' order by Vbillno,Vdt", "R", gCn) > 0 Then
''                Set sheet = workbook.Sheets.Item("b2cs")
''                If (sheet Is Nothing) Then
''                    MsgBox "Error Sales Sheet"
''                End If
''                a = 5
''                With tmp1
''                     .MoveFirst
''                     Do While Not .EOF
''                        sheet.Cells(a, 1) = "OE"
''                        DoEvents
''                        sheet.Cells(a, 2) = !Nar2
''                        sheet.Cells(a, 4) = !AMT2
''                        sheet.Cells(a, 5) = !AMT3
''                        sheet.Cells(a, 6) = 0
''                        a = a + 1
''                        prgBar.Value = prgBar.Value + 1
''                       .MoveNext
''                     Loop
''                End With
''             End If
''             '--CDNR
''             gCn.BeginTrans
''             gCn.Execute "delete from tmpGenTbl where username='" & gUserName & "'"
''             X1 = " qrytblVoucherSr1.VType in ('SN','PN') and qrytblVoucherSr1.VDRCRNTTYPE = 'B' and qrytblVoucherSr1.vdt>=cdate('" & Format(dtpFrDt.Value, "yyyy,mm,dd") & "')  and qrytblVoucherSr1.vdt<=Cdate('" & Format(dtpToDt.Value, "yyyy,mm,dd") & "') "  '
''             gCn.Execute " Insert Into TmpGentbl (Nar1,Vtype,Nar7,VBILLNO,VBILLDT,Nar6,Vdt,Nar2,Nar3,Nar5,Amt1,Amt2,Amt3,Amt4,Nar4,UserName)" _
''                         & " Select tblmastaccount.AcGSTIN,left(tblmastaccount.AcGSTIN,2),tblMastAccount.Acname,qrytblVoucherSr1.VBillNo,qrytblVoucherSr1.VBillDt,iif((qrytblVoucherSr1.VDrCrNtNo = '' or isnull(qrytblVoucherSr1.VDrCrNtNo)), cstr(qrytblVoucherSr1.Vno),qrytblVoucherSr1.VDrCrNtNo)  ,qrytblVoucherSr1.Vdt,iif(qrytblVoucherSr1.vtype = 'SN','D','C'),'07-Others',tblMastState.StateName,abs(qrytblVoucherSr1.VAmt),qrytblVoucherSr1.VCGSTRt+qrytblVoucherSr1.VSGSTRt+qrytblVoucherSr1.VIGSTRt,qrytblVoucherSr1.VTaxableAmt,0," _
''                         & " iif(qrytblVoucherSr1.VBillDt < cdate('2017/07/01'),'Y','N'),'" & gUserName & "' from qrytblVoucherSr1,tblMastAccount,tblMastState where qrytblVoucherSr1.VacCode = tblMastAccount.AcCode and left(tblMastAccount.AcGSTIN,2) = tblMastState.StateCd   and " & X1 & " "
''             '---Sales Return              '--CDNR
''             X1 = " tblAddLess.VType in ('RY')  and tblAddLess.vdt>=cdate('" & Format(dtpFrDt.Value, "yyyy,mm,dd") & "')  and tblAddLess.vdt<=Cdate('" & Format(dtpToDt.Value, "yyyy,mm,dd") & "') and len(tblmastaccount.AcGSTIN) > 3 "  '
''             gCn.Execute " Insert Into TmpGentbl (Nar1,Nar7,Vtype,Nar6,Vdt,VBILLNO,VBILLDT,Nar2,Nar3,Nar5,Amt1,Amt2,Amt3,Amt4,Nar4,UserName)" _
''                         & " select  ucase(tblmastaccount.AcGSTIN),tblMastAccount.Acname,left(tblmastaccount.AcGSTIN,2),tblAddLess.vno,tblAddLess.Vdt,tblAddLess.AdInvNo,tblAddLess.AdGpDt,'C','01-Sales Return', tblMastState.StateName,tblAddLess.AdBillAmt,iif(tblVouSub.VsubIGSTRt>0,tblVouSub.VsubIGSTRt,tblVouSub.VsubCGSTRt+tblVouSub.VsubSGSTRt),sum(tblVouSub.VSubAmt),0,iif(tblAddLess.AdGpDt < cdate('2017/07/01'),'Y','N'),'" & gUserName & "'" _
''                         & " From tblmastaccount,tblAddLess,tblVouSub,tblMastState where tblAddLess.AdAcCrCode = tblMastAccount.Accode and tblAddLess.Vno = tblVouSub.Vno and tblAddLess.Vtype = tblVouSub.Vtype and tblAddLess.Vyear = tblVouSub.Vyear and left(tblMastAccount.AcGSTIN,2) = tblMastState.StateCd  and " & X1 & " group by ucase(tblmastaccount.AcGSTIN),tblMastAccount.Acname,left(tblmastaccount.AcGSTIN,2),tblAddLess.vno,tblAddLess.Vdt,tblAddLess.AdInvNo,tblAddLess.AdGpDt,'C','01-Sales Return', tblMastState.StateName,tblAddLess.AdBillAmt,iif(tblVouSub.VsubIGSTRt>0,tblVouSub.VsubIGSTRt,tblVouSub.VsubCGSTRt+tblVouSub.VsubSGSTRt)"
''             DoEvents
''             gCn.CommitTrans
''             If GProcRstOpen(tmp1, "select * from TmpGentbl where TmpGentbl.UserName ='" & gUserName & "' order by Vbillno,Vdt", "R", gCn) > 0 Then
''                Set sheet = workbook.Sheets.Item("cdnr")
''                If (sheet Is Nothing) Then
''                    MsgBox "Error Sales Sheet"
''                End If
''                a = 5
''                With tmp1
''                     .MoveFirst
''                     Do While Not .EOF
''                        sheet.Cells(a, 1) = !Nar1
''                        DoEvents
''                        DoEvents
'''-V1.7
''                        sheet.Cells(a, 2) = !Nar7
''                        sheet.Cells(a, 3) = !Nar6
''                        sheet.Cells(a, 4) = Format(!Vdt, "dd-MMM-YYYY")
''                        sheet.Cells(a, 5) = !Nar2
''                        sheet.Cells(a, 6) = !Nar5
''                        sheet.Cells(a, 7) = "N"
''                        sheet.Cells(a, 8) = "Regular"
''                        If !Nar2 = "D" Then
''                            sheet.Cells(a, 9) = !Amt1
''                            sheet.Cells(a, 11) = !AMT2
''                            sheet.Cells(a, 12) = !AMT3
''                            sheet.Cells(a, 13) = !Amt4
''                       Else
''                            sheet.Cells(a, 9) = !Amt1 * -1
''                            sheet.Cells(a, 11) = !AMT2 * -1
''                            sheet.Cells(a, 12) = !AMT3 * -1
''                            sheet.Cells(a, 13) = !Amt4 * -1
''                       End If
''                        a = a + 1
''                        prgBar.Value = prgBar.Value + 1
''                       .MoveNext
''                     Loop
''                End With
''             End If
''             '--CDNUR
''
''             gCn.BeginTrans
''             gCn.Execute "delete from tmpGenTbl where username='" & gUserName & "'"
''             X1 = " qrytblVoucherSr1.VType in ('SN','PN') and qrytblVoucherSr1.VDrCrNtType = 'B' and qrytblVoucherSr1.vdt>=cdate('" & Format(dtpFrDt.Value, "yyyy,mm,dd") & "')  and qrytblVoucherSr1.vdt<=Cdate('" & Format(dtpToDt.Value, "yyyy,mm,dd") & "') and (len(tblmastaccount.AcGSTIN) <= 3 or isnull(tblmastaccount.AcGSTIN))"  '
''             gCn.Execute " Insert Into TmpGentbl (Nar1,Vtype,VBILLNO,VBILLDT,Nar6,Vdt,Nar2,Nar3,Nar5,Amt1,Amt2,Amt3,Amt4,Nar4,UserName)" _
''                         & " Select tblmastaccount.AcGSTIN,left(tblmastaccount.AcGSTIN,2),qrytblVoucherSr1.VBillNo,qrytblVoucherSr1.VBillDt,iif((qrytblVoucherSr1.VDrCrNtNo = '' or isnull(qrytblVoucherSr1.VDrCrNtNo)), cstr(qrytblVoucherSr1.Vno),qrytblVoucherSr1.VDrCrNtNo)   ,qrytblVoucherSr1.Vdt,iif(qrytblVoucherSr1.vtype = 'SN','D','C'),'07-Others',tblMastState.StateName,abs(qrytblVoucherSr1.VAmt),qrytblVoucherSr1.VCGSTRt+qrytblVoucherSr1.VSGSTRt+qrytblVoucherSr1.VIGSTRt,qrytblVoucherSr1.VTaxableAmt,0," _
''                         & " iif(qrytblVoucherSr1.VBillDt < cdate('2017/07/01'),'Y','N'),'" & gUserName & "' from qrytblVoucherSr1,tblMastAccount,tblMastState where qrytblVoucherSr1.VacCode = tblMastAccount.AcCode and left(tblMastAccount.AcGSTIN,2) = tblMastState.StateCd and qrytblVoucherSr1.VCGSTAmt+qrytblVoucherSr1.VSGSTAmt+qrytblVoucherSr1.VIGSTAmt > 0  and " & X1 & " "
''             DoEvents
''             gCn.CommitTrans
''             If GProcRstOpen(tmp1, "select * from TmpGentbl where TmpGentbl.UserName ='" & gUserName & "'", "R", gCn) > 0 Then
''                Set sheet = workbook.Sheets.Item("cdnur")
''                If (sheet Is Nothing) Then
''                    MsgBox "Error Sales Sheet"
''                End If
''                a = 5
''                With tmp1
''                     .MoveFirst
''                     Do While Not .EOF
''                        sheet.Cells(a, 1) = "B2CL"
''                        DoEvents
''                        sheet.Cells(a, 2) = !Nar6
''                        sheet.Cells(a, 3) = Format(!Vdt, "dd-MMM-YYYY")
''                        sheet.Cells(a, 4) = !Nar2
''                        sheet.Cells(a, 5) = !Nar5
''                        sheet.Cells(a, 6) = !Amt1
''                        sheet.Cells(a, 8) = !AMT2
''                        sheet.Cells(a, 9) = !AMT3
''                        sheet.Cells(a, 10) = !Amt4
''
''                        a = a + 1
''                        prgBar.Value = prgBar.Value + 1
''                       .MoveNext
''                     Loop
''                End With
''             End If
''             '--HSN
''             gCn.BeginTrans
''             gCn.Execute "delete from tmpGenTbl where username='" & gUserName & "'"
''             X1 = " tblAddLess.VType in ('SY')  and tblAddLess.vdt>=cdate('" & Format(dtpFrDt.Value, "yyyy,mm,dd") & "')  and tblAddLess.vdt<=Cdate('" & Format(dtpToDt.Value, "yyyy,mm,dd") & "') "  '
''             gCn.Execute " Insert Into TmpGentbl (Nar1,Nar2,Nar3,Amt1,Amt2,Amt8,Amt3,Amt4,Amt5,Amt6,Amt7,UserName)" _
''                         & " select tblMastItem.ITTMP2,iif(tblMastItem.ItOtherItem=0,'Yarn',''),'KGS-KILOGRAMS',sum(tblVouSub.VSubWt),sum(tblAddLess.AdBillAmt),tblVouSub.VsubCGSTRt+tblVouSub.VsubSGSTRt+tblVouSub.VsubIGSTRt,sum(tblAddLess.AdSubAmt),sum(tblVouSub.VsubIGSTAmt),sum(tblVouSub.VsubCGSTAmt),sum(tblVouSub.VsubSGSTAmt),0,'" & gUserName & "'" _
''                         & " From tblAddLess,tblVouSub,tblMastItem where tblVouSub.VSubItCode = tblMastItem.ItCode and tblAddLess.Vno = tblVouSub.Vno and tblAddLess.Vtype = tblVouSub.Vtype and tblAddLess.Vyear = tblVouSub.Vyear and " & X1 & " group by tblMastItem.ITTMP2,iif(tblMastItem.ItOtherItem=0,'Yarn',''),'KGS-KILOGRAMS',tblVouSub.VsubCGSTRt+tblVouSub.VsubSGSTRt+tblVouSub.VsubIGSTRt"
''             gCn.CommitTrans
''
''
''
'''             '--Sales Rtn HSN
'''             gCn.BeginTrans
'''             X1 = " tblAddLess.VType in ('RY')  and tblAddLess.vdt>=cdate('" & Format(dtpFrDt.Value, "yyyy,mm,dd") & "')  and tblAddLess.vdt<=Cdate('" & Format(dtpToDt.Value, "yyyy,mm,dd") & "') "  '
'''             gCn.Execute " Insert Into TmpGentbl (Nar1,Nar2,Nar3,Amt1,Amt2,Amt8,Amt3,Amt4,Amt5,Amt6,Amt7,UserName)" _
'''                         & " select tblMastItem.ITTMP2,iif(tblMastItem.ItOtherItem=0,'Sales Rtn','Other Sales'),'KGS-KILOGRAMS',sum(tblVouSub.VSubWt),sum(tblAddLess.AdBillAmt)*-1,tblVouSub.VsubCGSTRt+tblVouSub.VsubSGSTRt+tblVouSub.VsubIGSTRt,sum(tblVouSub.VSubAmt)*-1,sum(tblVouSub.VsubIGSTAmt)*-1,sum(tblVouSub.VsubCGSTAmt)*-1,sum(tblVouSub.VsubSGSTAmt)*-1,0,'" & gUserName & "'" _
'''                         & " From tblAddLess,tblVouSub,tblMastItem where tblVouSub.VSubItCode = tblMastItem.ItCode and tblAddLess.Vno = tblVouSub.Vno and tblAddLess.Vtype = tblVouSub.Vtype and tblAddLess.Vyear = tblVouSub.Vyear and " & X1 & " group by tblMastItem.ITTMP2,iif(tblMastItem.ItOtherItem=0,'Sales Rtn','Other Sales'),'KGS-KILOGRAMS',tblVouSub.VsubCGSTRt+tblVouSub.VsubSGSTRt+tblVouSub.VsubIGSTRt"
'''             gCn.CommitTrans
'''             '--DrNote HSN
'''             gCn.BeginTrans
'''             X1 = " qrytblVoucherSr1.VType = 'SN' and qrytblVoucherSr1.VDRCRNTTYPE = 'B' and qrytblVoucherSr1.vdt>=cdate('" & Format(dtpFrDt.Value, "yyyy,mm,dd") & "')  and qrytblVoucherSr1.vdt<=Cdate('" & Format(dtpToDt.Value, "yyyy,mm,dd") & "') "  '
'''             gCn.Execute " Insert Into TmpGentbl (Nar1,Nar2,Nar3,Amt1,Amt2,Amt8,Amt3,Amt4,Amt5,Amt6,Amt7,UserName)" _
'''                         & " select qrytblVoucherSr1.VItHSN,'Dr Note','No',sum(qrytblVoucherSr1.VCtrNo),sum(abs(qrytblVoucherSr1.Vamt)),qrytblVoucherSr1.VCGSTRt+qrytblVoucherSr1.VSGSTRt+qrytblVoucherSr1.VIGSTRt,sum(qrytblVoucherSr1.VTaxableAmt),sum(qrytblVoucherSr1.VIGSTAmt),sum(qrytblVoucherSr1.VCGSTAmt),sum(qrytblVoucherSr1.VSGSTAmt),0,'" & gUserName & "'" _
'''                         & " From qrytblVoucherSr1 where  " & X1 & " group by qrytblVoucherSr1.VItHSN,'Dr Note','No',qrytblVoucherSr1.VCGSTRt+qrytblVoucherSr1.VSGSTRt+qrytblVoucherSr1.VIGSTRt"
'''             gCn.CommitTrans
'''             '--CrrNote HSN
'''             gCn.BeginTrans
'''             X1 = " qrytblVoucherSr1.VType = 'PN' and qrytblVoucherSr1.VDRCRNTTYPE = 'B' and qrytblVoucherSr1.vdt>=cdate('" & Format(dtpFrDt.Value, "yyyy,mm,dd") & "')  and qrytblVoucherSr1.vdt<=Cdate('" & Format(dtpToDt.Value, "yyyy,mm,dd") & "') "  '
'''             gCn.Execute " Insert Into TmpGentbl (Nar1,Nar2,Nar3,Amt1,Amt2,Amt8,Amt3,Amt4,Amt5,Amt6,Amt7,UserName)" _
'''                         & " select qrytblVoucherSr1.VItHSN,'Cr Note','No',sum(qrytblVoucherSr1.VCtrNo),sum(abs(qrytblVoucherSr1.Vamt)*-1),qrytblVoucherSr1.VCGSTRt+qrytblVoucherSr1.VSGSTRt+qrytblVoucherSr1.VIGSTRt,sum(qrytblVoucherSr1.VTaxableAmt)*-1,sum(qrytblVoucherSr1.VIGSTAmt)*-1,sum(qrytblVoucherSr1.VCGSTAmt)*-1,sum(qrytblVoucherSr1.VSGSTAmt)*-1,0,'" & gUserName & "'" _
'''                         & " From qrytblVoucherSr1 where  " & X1 & " group by qrytblVoucherSr1.VItHSN,'Cr Note','No',qrytblVoucherSr1.VCGSTRt+qrytblVoucherSr1.VSGSTRt+qrytblVoucherSr1.VIGSTRt"
'''             gCn.CommitTrans
''
''
''             If GProcRstOpen(tmp1, "select * from TmpGentbl where TmpGentbl.UserName ='" & gUserName & "' order by Vbillno,Vdt", "R", gCn) > 0 Then
''                Set sheet = workbook.Sheets.Item("hsn")
''                If (sheet Is Nothing) Then
''                    MsgBox "Error Sales Sheet"
''                End If
''                a = 5
''                With tmp1
''                     .MoveFirst
''                     Do While Not .EOF
''                        sheet.Cells(a, 1) = !Nar1
''                        DoEvents
''                        sheet.Cells(a, 2) = !Nar2
''                        sheet.Cells(a, 3) = !Nar3
''                        sheet.Cells(a, 4) = !Amt1
''                        sheet.Cells(a, 5) = !AMT2
''                        sheet.Cells(a, 6) = !AMT8
''                        sheet.Cells(a, 7) = !AMT3
''                        sheet.Cells(a, 8) = !Amt4
''                        sheet.Cells(a, 9) = !Amt5
''                        sheet.Cells(a, 10) = !AMT6
''                        sheet.Cells(a, 11) = !Amt7
''                        a = a + 1
'''                        prgBar.Value = prgBar.Value + 1
''                       .MoveNext
''                     Loop
''                End With
''             End If
'''--Docs
'''-1) yarn Sale
''             a = 5
''             X1 = "Select left(adinvno,4),min(val(mid(adinvno,5,15))),max(val(mid(adinvno,5,15))),max(val(mid(adinvno,5,15)))-min(val(mid(adinvno,5,15)))+1 from tblAddLess where vtype = 'SY' and tblAddLess.vdt>=cdate('" & Format(dtpFrDt.Value, "yyyy,mm,dd") & "')  and tblAddLess.vdt<=Cdate('" & Format(dtpToDt.Value, "yyyy,mm,dd") & "') and (adtmp5 is null or adtmp5 = '')  group by left(adinvno,4)"
''             If GProcRstOpen(tmp1, X1, "R", gCn) > 0 Then
''                If Not IsNull(tmp1.Fields(0)) Then
''                    Set sheet = workbook.Sheets.Item("docs")
''                    If (sheet Is Nothing) Then
''                        MsgBox "Error Docs Sheet"
''                    End If
''                    sheet.Cells(a, 1) = "Invoices for outward supply"
''                    DoEvents
''                    sheet.Cells(a, 2) = Left(tmp1.Fields(0), 4) + CStr(tmp1.Fields(1))
''                    sheet.Cells(a, 3) = Left(tmp1.Fields(0), 4) + CStr(tmp1.Fields(2))
''                    sheet.Cells(a, 4) = tmp1.Fields(3)
''                    X1 = "Select * from tblAddLess where vtype = 'SY' and tblAddLess.vdt>=cdate('" & Format(dtpFrDt.Value, "yyyy,mm,dd") & "')  and tblAddLess.vdt<=Cdate('" & Format(dtpToDt.Value, "yyyy,mm,dd") & "') and (adtmp5 is null or adtmp5 = '') "
''
''                    TotCount = GProcRstOpen(tmp1, X1, "R", gCn)
''                    sheet.Cells(a, 5) = Val(sheet.Cells(a, 4)) - TotCount
''                    a = a + 1
''                End If
''             End If
'''-1) Other Sale
''             X1 = "Select min(val(mid(adinvno,5,15))),max(val(mid(adinvno,5,15))),max(val(mid(adinvno,5,15)))-min(val(mid(adinvno,5,15)))+1 from tblAddLess where vtype = 'SY' and tblAddLess.vdt>=cdate('" & Format(dtpFrDt.Value, "yyyy,mm,dd") & "')  and tblAddLess.vdt<=Cdate('" & Format(dtpToDt.Value, "yyyy,mm,dd") & "') and adtmp5 <> '' "
''             If GProcRstOpen(tmp1, X1, "R", gCn) > 0 Then
''                If Not IsNull(tmp1.Fields(0)) Then
''                    sheet.Cells(a, 1) = "Invoices for outward supply"
''                    DoEvents
''                    sheet.Cells(a, 2) = "FAB-" + Trim(Str(tmp1.Fields(0)))
''                    sheet.Cells(a, 3) = "FAB-" + Trim(Str(tmp1.Fields(1)))
''                    sheet.Cells(a, 4) = tmp1.Fields(2)
''                    X1 = "Select * from tblAddLess where vtype = 'SY' and tblAddLess.vdt>=cdate('" & Format(dtpFrDt.Value, "yyyy,mm,dd") & "')  and tblAddLess.vdt<=Cdate('" & Format(dtpToDt.Value, "yyyy,mm,dd") & "') and adtmp5 <> '' "
''
''                    TotCount = GProcRstOpen(tmp1, X1, "R", gCn)
''                    sheet.Cells(a, 5) = Val(sheet.Cells(a, 4)) - TotCount
''                    a = a + 1
''                End If
''             End If
'''3) Debit Note
''             X1 = "Select min(val(mid(VDRCRNTNO,4,15))),max(val(mid(VDRCRNTNO,4,15))),max(val(mid(VDRCRNTNO,4,15)))-min(val(mid(VDRCRNTNO,4,15)))+1 from qrytblVoucherSr1 where vtype = 'SN' and qrytblVoucherSr1.vdt>=cdate('" & Format(dtpFrDt.Value, "yyyy,mm,dd") & "')  and qrytblVoucherSr1.vdt<=Cdate('" & Format(dtpToDt.Value, "yyyy,mm,dd") & "') and VDRCRNTTYPE = 'B' "
''             If GProcRstOpen(tmp1, X1, "R", gCn) > 0 Then
''                If Not IsNull(tmp1.Fields(0)) Then
''                    sheet.Cells(a, 1) = "Debit Note"
''                    DoEvents
''                    sheet.Cells(a, 2) = "BD-" + Trim(Str(tmp1.Fields(0)))
''                    sheet.Cells(a, 3) = "BD-" + Trim(Str(tmp1.Fields(1)))
''                    sheet.Cells(a, 4) = tmp1.Fields(2)
''                    X1 = "Select * from qrytblVoucherSr1 where vtype = 'SN' and qrytblVoucherSr1.vdt>=cdate('" & Format(dtpFrDt.Value, "yyyy,mm,dd") & "')  and qrytblVoucherSr1.vdt<=Cdate('" & Format(dtpToDt.Value, "yyyy,mm,dd") & "') and VDRCRNTTYPE = 'B' "
''
''                    TotCount = GProcRstOpen(tmp1, X1, "R", gCn)
''                    sheet.Cells(a, 5) = Val(sheet.Cells(a, 4)) - TotCount
''                    a = a + 1
''                End If
''             End If
'''4) Credit Note
''             X1 = "Select min(val(mid(VDRCRNTNO,4,15))),max(val(mid(VDRCRNTNO,4,15))),max(val(mid(VDRCRNTNO,4,15)))-min(val(mid(VDRCRNTNO,4,15)))+1 from qrytblVoucherSr1 where vtype = 'PN' and qrytblVoucherSr1.vdt>=cdate('" & Format(dtpFrDt.Value, "yyyy,mm,dd") & "')  and qrytblVoucherSr1.vdt<=Cdate('" & Format(dtpToDt.Value, "yyyy,mm,dd") & "') and VDRCRNTTYPE = 'B' "
''             If GProcRstOpen(tmp1, X1, "R", gCn) > 0 Then
''                If Not IsNull(tmp1.Fields(0)) Then
''                    sheet.Cells(a, 1) = "Credit Note"
''                    DoEvents
''                    sheet.Cells(a, 2) = "BC-" + Trim(Str(tmp1.Fields(0)))
''                    sheet.Cells(a, 3) = "BC-" + Trim(Str(tmp1.Fields(1)))
''                    sheet.Cells(a, 4) = tmp1.Fields(2)
''                    X1 = "Select * from qrytblVoucherSr1 where vtype = 'PN' and qrytblVoucherSr1.vdt>=cdate('" & Format(dtpFrDt.Value, "yyyy,mm,dd") & "')  and qrytblVoucherSr1.vdt<=Cdate('" & Format(dtpToDt.Value, "yyyy,mm,dd") & "') and VDRCRNTTYPE = 'B' "
''                    TotCount = GProcRstOpen(tmp1, X1, "R", gCn)
''                    sheet.Cells(a, 5) = Val(sheet.Cells(a, 4)) - TotCount
''                    a = a + 1
''                End If
''             End If
'''5) Credit Note (sales Rtn)
''             X1 = "Select min(vno),max(vno),max(vno)-min(vno)+1 from tblAddLess where vtype = 'RY' and tblAddLess.vdt>=cdate('" & Format(dtpFrDt.Value, "yyyy,mm,dd") & "')  and tblAddLess.vdt<=Cdate('" & Format(dtpToDt.Value, "yyyy,mm,dd") & "') "
''             If GProcRstOpen(tmp1, X1, "R", gCn) > 0 Then
''                If Not IsNull(tmp1.Fields(0)) Then
''                    Set sheet = workbook.Sheets.Item("docs")
''                    If (sheet Is Nothing) Then
''                        MsgBox "Error Docs Sheet"
''                    End If
''                    sheet.Cells(a, 1) = "Credit Note"
''                    DoEvents
''                    sheet.Cells(a, 2) = tmp1.Fields(0)
''                    sheet.Cells(a, 3) = tmp1.Fields(1)
''                    sheet.Cells(a, 4) = tmp1.Fields(2)
''                    X1 = "Select * from tblAddLess where vtype = 'RY' and tblAddLess.vdt>=cdate('" & Format(dtpFrDt.Value, "yyyy,mm,dd") & "')  and tblAddLess.vdt<=Cdate('" & Format(dtpToDt.Value, "yyyy,mm,dd") & "') "
''
''                    TotCount = GProcRstOpen(tmp1, X1, "R", gCn)
''                    sheet.Cells(a, 5) = Val(sheet.Cells(a, 4)) - TotCount
''                    a = a + 1
''                End If
''             End If
''             workbook.Close SaveChanges:=True
''             excel_app.Quit
''             MsgBox "Compileted"
''             Exit Sub
            
            Case "Form GSTR - 2"
             
             '---FOr BarCount
             gCn.BeginTrans
             gCn.Execute "delete from tmpGenTbl where username='" & gUserName & "'"
             gCn.CommitTrans
        
             X1 = " tblAddLess.VType in ('PY','PO') and tblAddLess.AdGpDt>=cdate('" & Format(dtpFrDt.Value, "yyyy,mm,dd") & "')  and tblAddLess.AdGpDt<=Cdate('" & Format(dtpToDt.Value, "yyyy,mm,dd") & "') and len(tblmastaccount.AcGSTIN) > 3 and ADTMP2 = 0"  '
             gCn.Execute " Insert Into TmpGentbl (vno,Nar6,vyear,Nar1,vtype,Nar7,VbillNo,Vdt,Amt1,Nar4,Nar5,Nar2,Nar3,Amt2,Amt3,Amt4,Amt5,Amt6,Amt7,Amt8,UserName)" _
                         & " select tblAddLess.vno,tblAddLess.vtype,tblAddLess.vyear,ucase(tblmastaccount.AcGSTIN),left(tblmastaccount.AcGSTIN,2),tblmastaccount.AcName,tblAddLess.AdInvNo,tblAddLess.AdGpDt,tblAddLess.AdBillAmt,tblMastState.StateName, 'N','Regular','', iif(tblVouSub.VsubIGSTRt>0,tblVouSub.VsubIGSTRt,tblVouSub.VsubCGSTRt+tblVouSub.VsubSGSTRt),sum(tblVouSub.VSubAmt),sum(VsubIGSTAmt),sum(VsubCGSTAmt),sum(VsubSGSTAmt),0,1,'" & gUserName & "'" _
                         & " From tblmastaccount,tblAddLess,tblVouSub,tblMastState where tblAddLess.AdAcCrCode = tblMastAccount.Accode and tblAddLess.Vno = tblVouSub.Vno and tblAddLess.Vtype = tblVouSub.Vtype and tblAddLess.Vyear = tblVouSub.Vyear and left(tblMastAccount.AcGSTIN,2) = tblMastState.StateCd  and " & X1 & "" _
                         & " group by tblAddLess.vno,tblAddLess.vtype,tblAddLess.vyear,tblmastaccount.AcGSTIN,tblmastaccount.AcName,AdInvNo,tblAddLess.AdGpDt,tblAddLess.AdBillAmt,tblMastState.StateName,iif(tblVouSub.VsubIGSTRt>0,tblVouSub.VsubIGSTRt,tblVouSub.VsubCGSTRt+tblVouSub.VsubSGSTRt)"
            
             X1 = " tblAddLess.VType in ('PO') and tblAddLess.AdBookDt>=cdate('" & Format(dtpFrDt.Value, "yyyy,mm,dd") & "')  and tblAddLess.AdBookDt<=Cdate('" & Format(dtpToDt.Value, "yyyy,mm,dd") & "') and ADTMP2 = 1  "
             gCn.Execute " Insert Into TmpGentbl (vno,Nar6,vyear,Nar1,vtype,Nar7,VbillNo,Vdt,Amt1,Nar4,Amt2,Amt3,Amt4,Amt5,UserName)" _
                         & " select tblAddLess.vno,tblAddLess.vtype,tblAddLess.vyear,ucase(tblmastaccount.AcGSTIN),left(tblmastaccount.AcGSTIN,2),tblmastaccount.AcName,tblAddLess.ADTMP5,tblAddLess.AdBookDt,tblAddLess.AdSubAmt,tblMastState.StateName, iif(tblVouSub.VsubIGSTRt>0,tblVouSub.VsubIGSTRt,tblVouSub.VsubCGSTRt+tblVouSub.VsubSGSTRt),sum(VsubIGSTAmt),sum(VsubCGSTAmt),sum(VsubSGSTAmt),'" & gUserName & "'" _
                         & " From tblmastaccount,tblAddLess,tblVouSub,tblMastState where tblAddLess.AdAcCrCode = tblMastAccount.Accode and tblAddLess.Vno = tblVouSub.Vno and tblAddLess.Vtype = tblVouSub.Vtype and tblAddLess.Vyear = tblVouSub.Vyear and left(tblMastAccount.AcGSTIN,2) = tblMastState.StateCd  and " & X1 & "" _
                         & " group by tblAddLess.vno,tblAddLess.vtype,tblAddLess.vyear,tblmastaccount.AcGSTIN,tblmastaccount.AcName,ADTMP5,tblAddLess.AdBookDt,tblAddLess.AdSubAmt,tblMastState.StateName,iif(tblVouSub.VsubIGSTRt>0,tblVouSub.VsubIGSTRt,tblVouSub.VsubCGSTRt+tblVouSub.VsubSGSTRt)"
        
        
             X1 = " qrytblVoucherSr1.VType in ('SN','PN') and qrytblVoucherSr1.VDRCRNTTYPE = 'S' and qrytblVoucherSr1.vdt>=cdate('" & Format(dtpFrDt.Value, "yyyy,mm,dd") & "')  and qrytblVoucherSr1.vdt<=Cdate('" & Format(dtpToDt.Value, "yyyy,mm,dd") & "') "  '
             gCn.Execute " Insert Into TmpGentbl (vno,Nar6,vyear,Nar1,vtype,Nar7,VbillNo,Vdt,Amt1,Nar4,Nar5,Nar2,Nar3,Amt2,Amt3,Amt4,Amt5,Amt6,Amt7,Amt8,UserName)" _
                         & " Select qrytblVoucherSr1.Vno,qrytblVoucherSr1.Vtype,qrytblVoucherSr1.Vyear,ucase(tblmastaccount.AcGSTIN),left(tblmastaccount.AcGSTIN,2),tblMastAccount.Acname,qrytblVoucherSr1.VDRCRNTNO,qrytblVoucherSr1.VDt,abs(qrytblVoucherSr1.vamt),tblMastState.StateName,'N','Regular','',iif(qrytblVoucherSr1.VIGSTRt>0,qrytblVoucherSr1.VIGSTRt,qrytblVoucherSr1.VCGSTRt+qrytblVoucherSr1.VSGSTRt),sum(qrytblVoucherSr1.VTaxableAmt)," _
                         & " sum(qrytblVoucherSr1.VIGSTAmt),sum(qrytblVoucherSr1.VCGSTAmt),sum(qrytblVoucherSr1.VSGSTAmt),0,1,'" & gUserName & "' " _
                         & " from qrytblVoucherSr1,tblMastAccount,tblMastState where qrytblVoucherSr1.VacCode = tblMastAccount.AcCode and left(tblMastAccount.AcGSTIN,2) = tblMastState.StateCd   and " & X1 & " " _
                         & " group by qrytblVoucherSr1.Vno,qrytblVoucherSr1.Vtype,qrytblVoucherSr1.Vyear,ucase(tblmastaccount.AcGSTIN),left(tblmastaccount.AcGSTIN,2),tblMastAccount.Acname,qrytblVoucherSr1.VDRCRNTNO,qrytblVoucherSr1.VDt,abs(qrytblVoucherSr1.vamt),  qrytblVoucherSr1.Vdt,tblMastState.StateName,'N','Regular','',iif(qrytblVoucherSr1.VIGSTRt>0,qrytblVoucherSr1.VIGSTRt,qrytblVoucherSr1.VCGSTRt+qrytblVoucherSr1.VSGSTRt)"

             DoEvents
             DoEvents
             DoEvents
             
             BarCnt = GProcRstOpen(tmp1, "select * from TmpGentbl where TmpGentbl.UserName ='" & gUserName & "'", "R", gCn)
             '---End Bar Count
             flnm = "D:\GSTR2-" + gCCode + "-" + Format(dtpFrDt, "mmm") + ".xlsx"
             FileCopy gReportPath & "\GSTR2-Blank.xlsx", flnm
             gCn.BeginTrans
             gCn.Execute "delete from tmpGenTbl where username='" & gUserName & "'"
             '---Purchase Detail
             '--B2B
             X1 = " tblAddLess.VType in ('PY','PO') and tblAddLess.AdGpDt>=cdate('" & Format(dtpFrDt.Value, "yyyy,mm,dd") & "')  and tblAddLess.AdGpDt<=Cdate('" & Format(dtpToDt.Value, "yyyy,mm,dd") & "') and len(tblmastaccount.AcGSTIN) > 3 and ADTMP2 = 0"  '
             gCn.Execute " Insert Into TmpGentbl (vno,Nar6,vyear,Nar1,vtype,Nar7,VbillNo,Vdt,Amt1,Nar4,Nar5,Nar2,Nar3,Amt2,Amt3,Amt4,Amt5,Amt6,Amt7,Amt8,UserName)" _
                         & " select tblAddLess.vno,tblAddLess.vtype,tblAddLess.vyear,ucase(tblmastaccount.AcGSTIN),left(tblmastaccount.AcGSTIN,2),tblmastaccount.AcName,tblAddLess.AdInvNo,tblAddLess.AdGpDt,tblAddLess.AdBillAmt,tblMastState.StateName, 'N','Regular','', iif(tblVouSub.VsubIGSTRt>0,tblVouSub.VsubIGSTRt,tblVouSub.VsubCGSTRt+tblVouSub.VsubSGSTRt),sum(tblVouSub.VSubTaxableAmt),sum(VsubIGSTAmt),sum(VsubCGSTAmt),sum(VsubSGSTAmt),0,1,'" & gUserName & "'" _
                         & " From tblmastaccount,tblAddLess,tblVouSub,tblMastState where tblAddLess.AdAcCrCode = tblMastAccount.Accode and tblAddLess.Vno = tblVouSub.Vno and tblAddLess.Vtype = tblVouSub.Vtype and tblAddLess.Vyear = tblVouSub.Vyear and left(tblMastAccount.AcGSTIN,2) = tblMastState.StateCd  and " & X1 & "" _
                         & " group by tblAddLess.vno,tblAddLess.vtype,tblAddLess.vyear,tblmastaccount.AcGSTIN,tblmastaccount.AcName,AdInvNo,tblAddLess.AdGpDt,tblAddLess.AdBillAmt,tblMastState.StateName,iif(tblVouSub.VsubIGSTRt>0,tblVouSub.VsubIGSTRt,tblVouSub.VsubCGSTRt+tblVouSub.VsubSGSTRt)"
             gCn.CommitTrans
             
             gCn.BeginTrans
             
             
             X1 = " qrytblVoucherSr1.VType in ('SN','PN') and qrytblVoucherSr1.VDRCRNTTYPE = 'S' and qrytblVoucherSr1.vdt>=cdate('" & Format(dtpFrDt.Value, "yyyy,mm,dd") & "')  and qrytblVoucherSr1.vdt<=Cdate('" & Format(dtpToDt.Value, "yyyy,mm,dd") & "') "  '
             gCn.Execute " Insert Into TmpGentbl (vno,Nar6,vyear,Nar1,vtype,Nar7,VbillNo,Vdt,Amt1,Nar4,Nar5,Nar2,Nar3,Amt2,Amt3,Amt4,Amt5,Amt6,Amt7,Amt8,UserName)" _
                         & " Select qrytblVoucherSr1.Vno,qrytblVoucherSr1.Vtype,qrytblVoucherSr1.Vyear,ucase(tblmastaccount.AcGSTIN),left(tblmastaccount.AcGSTIN,2),tblMastAccount.Acname,qrytblVoucherSr1.VDRCRNTNO,qrytblVoucherSr1.VDt,abs(qrytblVoucherSr1.vamt),tblMastState.StateName,'N','Regular','',iif(qrytblVoucherSr1.VIGSTRt>0,qrytblVoucherSr1.VIGSTRt,qrytblVoucherSr1.VCGSTRt+qrytblVoucherSr1.VSGSTRt),sum(qrytblVoucherSr1.VTaxableAmt)," _
                         & " sum(qrytblVoucherSr1.VIGSTAmt),sum(qrytblVoucherSr1.VCGSTAmt),sum(qrytblVoucherSr1.VSGSTAmt),0,1,'" & gUserName & "' " _
                         & " from qrytblVoucherSr1,tblMastAccount,tblMastState where qrytblVoucherSr1.VacCode = tblMastAccount.AcCode and left(tblMastAccount.AcGSTIN,2) = tblMastState.StateCd   and " & X1 & " " _
                         & " group by qrytblVoucherSr1.Vno,qrytblVoucherSr1.Vtype,qrytblVoucherSr1.Vyear,ucase(tblmastaccount.AcGSTIN),left(tblmastaccount.AcGSTIN,2),tblMastAccount.Acname,qrytblVoucherSr1.VDRCRNTNO,qrytblVoucherSr1.VDt,abs(qrytblVoucherSr1.vamt),  qrytblVoucherSr1.Vdt,tblMastState.StateName,'N','Regular','',iif(qrytblVoucherSr1.VIGSTRt>0,qrytblVoucherSr1.VIGSTRt,qrytblVoucherSr1.VCGSTRt+qrytblVoucherSr1.VSGSTRt)"
             
             
             gCn.CommitTrans
             
             
             DoEvents
             DoEvents
             DoEvents
             DoEvents
             If OptShort.Value = True Then    '--B2B Summery
                If txttemp(2).text <> "" Then
                   X1 = "{TmpGentbl.UserName}='" & gUserName & "' and {TmpGentbl.Nar7}= '" & txttemp(2).text & "' "
                Else
                   X1 = "{TmpGentbl.UserName}='" & gUserName & "'"
                End If
                .SelectionFormula = X1
                .ReportFileName = gReportPath & "rptGSTR2Summery.rpt"
                .Action = 1
                Exit Sub
             End If
             If OptDetail.Value = True Then
                If txttemp(2).text <> "" Then
                   X1 = "{TmpGentbl.UserName}='" & gUserName & "' and {TmpGentbl.Nar7}= '" & txttemp(2).text & "' "
                Else
                   X1 = "{TmpGentbl.UserName}='" & gUserName & "'"
                End If
                .SelectionFormula = X1
                If ChkConsolidate.Value = 1 Then
                   .ReportFileName = gReportPath & "rptGSTR2Datewise.rpt"
                Else
                   .ReportFileName = gReportPath & "rptGSTR2Detail.rpt"
                End If
                .Action = 1
                Exit Sub
             End If
             
             
             
             Set excel_app = New Excel.Application
             prgBar.Max = BarCnt
             prgBar.Value = 0
             prgBar.Visible = True
'             excel_app.Visible = True
             excel_app.Visible = False
             Set workbook = excel_app.Workbooks.Open(filename:=flnm)
             RecYN = GProcRstOpen(tmp1, "select * from TmpGentbl where TmpGentbl.UserName ='" & gUserName & "'", "R", gCn)
             If RecYN > 0 Then
                Set sheet = workbook.Sheets.Item("b2b")
                If (sheet Is Nothing) Then
                    MsgBox "Error Sales Sheet"
                End If
                a = 5
                With tmp1
                     .MoveFirst
                     Do While Not .EOF
                        sheet.Cells(a, 1) = !Nar1
                        sheet.Cells(a, 2) = Left(!VbillNo, 15)
                        DoEvents
                        sheet.Cells(a, 3) = Format(!Vdt, "dd-MMM-YYYY")
                        sheet.Cells(a, 4) = !Amt1
                        sheet.Cells(a, 5) = !Nar4
                        sheet.Cells(a, 6) = "N"
                        sheet.Cells(a, 7) = "Regular"
                        sheet.Cells(a, 8) = !AMT2
                        sheet.Cells(a, 9) = !AMT3
                        sheet.Cells(a, 10) = !Amt4
                        sheet.Cells(a, 11) = !Amt5
                        sheet.Cells(a, 12) = !Amt6
                        sheet.Cells(a, 13) = 0
                        sheet.Cells(a, 14) = "Inputs"
                        sheet.Cells(a, 15) = !Amt4
                        sheet.Cells(a, 16) = !Amt5
                        sheet.Cells(a, 17) = !Amt6
                        sheet.Cells(a, 18) = 0
                        a = a + 1
                        prgBar.Value = prgBar.Value + 1
                       .MoveNext
                     Loop
                End With
             End If
             '--B2BUR
             DoEvents
             DoEvents
             DoEvents
             gCn.BeginTrans
             gCn.Execute "delete from tmpGenTbl where username='" & gUserName & "'"
             X1 = " tblAddLess.VType in ('PO') and tblAddLess.AdBookDt>=cdate('" & Format(dtpFrDt.Value, "yyyy,mm,dd") & "')  and tblAddLess.AdBookDt<=Cdate('" & Format(dtpToDt.Value, "yyyy,mm,dd") & "') and ADTMP2 = 1  "
             gCn.Execute " Insert Into TmpGentbl (vno,Nar6,vyear,Nar1,vtype,Nar7,VbillNo,Vdt,Amt1,Nar4,Amt2,Amt3,Amt4,Amt5,UserName)" _
                         & " select tblAddLess.vno,tblAddLess.vtype,tblAddLess.vyear,ucase(tblmastaccount.AcGSTIN),left(tblmastaccount.AcGSTIN,2),tblmastaccount.AcName,tblAddLess.ADTMP5,tblAddLess.AdBookDt,tblAddLess.AdSubAmt,tblMastState.StateName, iif(tblVouSub.VsubIGSTRt>0,tblVouSub.VsubIGSTRt,tblVouSub.VsubCGSTRt+tblVouSub.VsubSGSTRt),sum(VsubIGSTAmt),sum(VsubCGSTAmt),sum(VsubSGSTAmt),'" & gUserName & "'" _
                         & " From tblmastaccount,tblAddLess,tblVouSub,tblMastState where tblAddLess.AdAcCrCode = tblMastAccount.Accode and tblAddLess.Vno = tblVouSub.Vno and tblAddLess.Vtype = tblVouSub.Vtype and tblAddLess.Vyear = tblVouSub.Vyear and left(tblMastAccount.AcGSTIN,2) = tblMastState.StateCd  and " & X1 & "" _
                         & " group by tblAddLess.vno,tblAddLess.vtype,tblAddLess.vyear,tblmastaccount.AcGSTIN,tblmastaccount.AcName,ADTMP5,tblAddLess.AdBookDt,tblAddLess.AdSubAmt,tblMastState.StateName,iif(tblVouSub.VsubIGSTRt>0,tblVouSub.VsubIGSTRt,tblVouSub.VsubCGSTRt+tblVouSub.VsubSGSTRt)"
             gCn.CommitTrans
             DoEvents
             DoEvents
             If GProcRstOpen(tmp1, "select * from TmpGentbl where TmpGentbl.UserName ='" & gUserName & "'", "R", gCn) > 0 Then
                Set sheet = workbook.Sheets.Item("b2bur")
                If (sheet Is Nothing) Then
                    MsgBox "Error B2BUR Sheet"
                End If
                a = 5
                With tmp1
                     .MoveFirst
                     Do While Not .EOF
                        sheet.Cells(a, 1) = !Nar7
                        DoEvents
                        sheet.Cells(a, 2) = Left(!VbillNo, 15)
                        sheet.Cells(a, 3) = Format(!Vdt, "dd-MMM-YYYY")
                        sheet.Cells(a, 4) = !Amt1 + !AMT3 + !Amt4 + !Amt5
                        sheet.Cells(a, 5) = !Nar4
                        If Left(!Nar4, 2) = "27" Then
                            sheet.Cells(a, 6) = "Intra State"
                        Else
                            sheet.Cells(a, 6) = "Inter State"
                        End If
                        sheet.Cells(a, 7) = !AMT2
                        sheet.Cells(a, 8) = !Amt1
                        sheet.Cells(a, 9) = !AMT3
                        sheet.Cells(a, 10) = !Amt4
                        sheet.Cells(a, 11) = !Amt5
                        sheet.Cells(a, 12) = 0
                        sheet.Cells(a, 13) = "Input services"
                        sheet.Cells(a, 14) = !AMT3
                        sheet.Cells(a, 15) = !Amt4
                        sheet.Cells(a, 16) = !Amt5
                        sheet.Cells(a, 17) = 0
                        a = a + 1
                        prgBar.Value = prgBar.Value + 1
                       .MoveNext
                     Loop
                End With
             End If
             workbook.Close SaveChanges:=True
             excel_app.Quit
             MsgBox "Compileted"
             Exit Sub

        Case "Monthwise Summery GST"
             If Chkonly.Value = 1 Then
                GProcPurpareGSTSummery dtpFrDt.Value, dtpToDt.Value, OptDetail.Value, True
             Else
                GProcPurpareGSTSummery dtpFrDt.Value, dtpToDt.Value, OptDetail.Value, False
             End If
            X1 = "{TmpGentbl.UserName}='" & gUserName & "' "
            .Formulas(4) = "wFrToDt = 'From : ' & '" & Format(dtpFrDt.Value, "dd/MM/yyyy") & "' & '  To ' & '" & Format(dtpToDt, "dd/MM/yyyy") & "'"
            .Formulas(15) = "wCompGstin = '" & gCGSTIN & "'"
            If OptDetail.Value = True Then
                .Formulas(5) = "wRptHd = 'Report On : Bill Date'"
            ElseIf OptShort.Value = True Then
                .Formulas(5) = "wRptHd = 'Report On : Received Date'"
            End If
            .SelectionFormula = X1
            If Chkonly.Value = 1 Then
              .ReportFileName = gReportPath & "rptMonthwiseGST.rpt"
            Else
              .ReportFileName = gReportPath & "rptMonthwiseGSTSumm.rpt"
            End If
        Case "HSN Summery"
             GProcPurpareHSNSummery dtpFrDt.Value, dtpToDt.Value, OptDetail.Value
             X1 = "{TmpGentbl.UserName}='" & gUserName & "'"
             .SelectionFormula = X1
             .Formulas(4) = "wFrToDt = 'From : ' & '" & Format(dtpFrDt.Value, "dd/MM/yyyy") & "' & '  To ' & '" & Format(dtpToDt, "dd/MM/yyyy") & "'"
             .Formulas(15) = "wCompGstin = '" & gCGSTIN & "'"
             If OptDetail.Value = True Then
                .Formulas(6) = "wOnDt = 'On Purch Bill Date'"
             Else
                .Formulas(6) = "wOnDt = 'On Purch Recd. Date'"
             End If
             
             .ReportFileName = gReportPath & "rptHSNSummery.rpt"
             .Action = 1
             Exit Sub
            
        
        
        Case "Partywise Sale/Purchase Summery"
               GprocPrepareSalePurSummGST Format(dtpFrDt.Value, "dd/MM/yyyy"), Format(dtpToDt.Value, "dd/MM/yyyy"), IIf(OptDetail.Value = True, "S", "P")
               .Formulas(4) = "wFrToDt = 'From : ' & '" & Format(dtpFrDt.Value, "dd/MM/yyyy") & "' & '  To ' & '" & Format(dtpToDt.Value, "dd/MM/yyyy") & "'"
               If OptDetail.Value = True Then
                 .Formulas(5) = "wRptName='Partywise Sales Summery'"
               Else
                 .Formulas(5) = "wRptName='Partywise Purchase Summery'"
               End If
               X1 = "{tmpGenTbl2.UserName}='" & gUserName & "'"
               .SelectionFormula = X1
               .ReportFileName = gReportPath & "rptSalePurSummGST.rpt"
        Case "GSTR - 2A  -  B2B  (Mismatch)"
             Dim ExSr As Long
             gCn.BeginTrans
             gCn.Execute "delete from tmpGenTbl where username='" & gUserName & "'"
             '---Purchase Detail
             '--B2B
             X1 = " tblAddLess.VType in ('PY','PO') and tblAddLess.AdGpDt>=cdate('" & Format(dtpFrDt.Value, "yyyy,mm,dd") & "')  and tblAddLess.AdGpDt<=Cdate('" & Format(dtpToDt.Value, "yyyy,mm,dd") & "') and len(tblmastaccount.AcGSTIN) > 3 and ADTMP2 = 0  and (isnull(adTmp5)  or adTmp5 = '' or left(adTmp5,3)= 'CLO')  "
             gCn.Execute " Insert Into TmpGentbl (vno,Nar6,vyear,Nar1,vtype,Nar7,Nar4,Vdt,Amt1,Amt2,Amt3,Amt4,Amt5,UserName)" _
                         & " select tblAddLess.vno,tblAddLess.vtype,tblAddLess.vyear,ucase(tblmastaccount.AcGSTIN),left(tblmastaccount.AcGSTIN,2),tblmastaccount.AcName,tblAddLess.AdInvNo,tblAddLess.AdGpDt,tblAddLess.AdBillAmt, sum(tblVouSub.VSubTaxableAmt),sum(VsubIGSTAmt),sum(VsubCGSTAmt),sum(VsubSGSTAmt),'" & gUserName & "'" _
                         & " From tblmastaccount,tblAddLess,tblVouSub where tblAddLess.AdAcCrCode = tblMastAccount.Accode and tblAddLess.Vno = tblVouSub.Vno and tblAddLess.Vtype = tblVouSub.Vtype and tblAddLess.Vyear = tblVouSub.Vyear  and " & X1 & "" _
                         & " group by tblAddLess.vno,tblAddLess.vtype,tblAddLess.vyear,tblmastaccount.AcGSTIN,tblmastaccount.AcName,AdInvNo,tblAddLess.AdGpDt,tblAddLess.AdBillAmt"
             gCn.CommitTrans
             gCn.BeginTrans
'             X1 = " tblAddLess.VType in ('PY') and tblAddLess.AdGpDt>=cdate('" & Format(dtpFrDt.Value, "yyyy,mm,dd") & "')  and tblAddLess.AdGpDt<=Cdate('" & Format(dtpToDt.Value, "yyyy,mm,dd") & "') and len(tblmastaccount.AcGSTIN) > 3   and left(adTmp5,3)= 'COT'   "  '
'             gCn.Execute " Insert Into TmpGentbl (vno,Nar6,vyear,Nar1,vtype,Nar7,Nar4,Vdt,Amt1,Amt2,Amt3,Amt4,Amt5,UserName)" _
'                         & " select tblAddLess.vno,tblAddLess.vtype,tblAddLess.vyear,ucase(tblmastaccount.AcGSTIN),left(tblmastaccount.AcGSTIN,2),tblmastaccount.AcName,tblAddLess.AdInvNo,tblAddLess.AdGpDt,tblAddLess.AdBillAmt, sum(tblAddLess.AdAssVal),sum(AdClothIgstAmt),sum(AdClothCgstAmt),sum(AdClothSgstAmt),'" & gUserName & "'" _
'                         & " From tblmastaccount,tblAddLess where tblAddLess.AdAcCrCode = tblMastAccount.Accode and  " & X1 & "" _
'                         & " group by tblAddLess.vno,tblAddLess.vtype,tblAddLess.vyear,tblmastaccount.AcGSTIN,tblmastaccount.AcName,AdInvNo,tblAddLess.AdGpDt,tblAddLess.AdBillAmt"
'             gCn.CommitTrans
             
             flnm = "D:\GSTR2A" + gCCode + ".xlsx"
         
             Set excel_app = New Excel.Application
             excel_app.Visible = False
             Set workbook = excel_app.Workbooks.Open(filename:=flnm)
             Set sheet = workbook.Sheets.Item("b2b")
             gCn.BeginTrans
             gCn.Execute "delete from tmpRptTbl where username='" & gUserName & "'"
             gCn.CommitTrans
             X1 = "Select * from tmpRptTbl where username='" & gUserName & "'"
             GProcRstOpen tmp1, X1, "O", gCn
             Set sheet = workbook.Sheets.Item("b2b")
             If (sheet Is Nothing) Then
                 MsgBox "Error Sales Sheet"
                 Exit Sub
             End If
             gCn.BeginTrans
             a = 7
             ExSr = 0
'             Do While Len(sheet.Cells(a, 1)) > 0
             Do While ExSr < 5
                If Right(sheet.Cells(a, 3), 5) = "Total" Or sheet.Cells(a, 3) = "" Then
                    a = a + 1
                Else
                    tmp1.AddNew
                    tmp1!Nar7 = Left(sheet.Cells(a, 2), 40) '-- Party Name
                    tmp1!Nar1 = sheet.Cells(a, 1)  '-- GSTIN
    '                tmp1!Vtype = Left(sheet.Cells(a, 2), 2)  '--State Code
                    tmp1!Nar4 = sheet.Cells(a, 3)   '--Bill No
                    tmp1!Vdt = Format(CDate(sheet.Cells(a, 5)), "yyyy/MM/DD")   '--Bill Date
                    tmp1!Amt1 = Val(sheet.Cells(a, 6))  '-- Bill Amt
                    tmp1!AMT2 = Val(sheet.Cells(a, 10))  '-- Taxable
                    tmp1!AMT3 = Val(sheet.Cells(a, 11))  ' - I
                    tmp1!Amt4 = Val(sheet.Cells(a, 12))   '-- C
                    tmp1!Amt5 = Val(sheet.Cells(a, 13))  '-- S
                    tmp1!UserName = gUserName
                    tmp1.Update
                    a = a + 1
                End If
                If Len(sheet.Cells(a, 1)) = 0 And Len(sheet.Cells(a + 1, 1)) = 0 Then ExSr = ExSr + 1
             Loop
             gCn.CommitTrans
             workbook.Close SaveChanges:=True
             excel_app.Quit
             If GProcRstOpen(tmp1, "select * from TmpGentbl where TmpGentbl.UserName ='" & gUserName & "'", "R", gCn) > 0 Then
                With tmp1
                     .MoveFirst
                     Do While Not .EOF
                        '-- For Matched Data
'                        gCn.Execute "Update tmpRptTbl set tmpRptTbl.Nar2 = 'Match' where tmpRptTbl.Nar1 = '" & tmp1.Fields("Nar1") & "' and tmpRptTbl.Nar4 = '" & tmp1.Fields("Nar4") & "' and tmpRptTbl.VDt = Cdate('" & Format(tmp1.Fields("Vdt"), "yyyy,mm,dd") & "') and  " _

                        gCn.Execute "Update tmpRptTbl set tmpRptTbl.Nar2 = 'Match' where tmpRptTbl.Nar1 = '" & tmp1.Fields("Nar1") & "' and tmpRptTbl.VDt = Cdate('" & Format(tmp1.Fields("Vdt"), "yyyy,mm,dd") & "') and  " _
                            & " tmpRptTbl.amt3 = " & tmp1.Fields("Amt3") & " and tmpRptTbl.amt4 = " & tmp1.Fields("Amt4") & " and tmpRptTbl.amt5 = " & tmp1.Fields("Amt5") & " and tmpRptTbl.UserName = '" & gUserName & "' "
                        '-- For MisMatched Data
                        gCn.Execute "Update tmpRptTbl set tmpRptTbl.Nar2 = 'A - Mismatch'  where tmpRptTbl.Nar1 = '" & tmp1.Fields("Nar1") & "' and tmpRptTbl.Nar4 = '" & tmp1.Fields("Nar4") & "' and tmpRptTbl.VDt = Cdate('" & Format(tmp1.Fields("Vdt"), "yyyy,mm,dd") & "')  and  " _
                            & " ( tmpRptTbl.amt3 <> " & tmp1.Fields("Amt3") & " or tmpRptTbl.amt4 <> " & tmp1.Fields("Amt4") & " or tmpRptTbl.amt5 <> " & tmp1.Fields("Amt5") & ") and tmpRptTbl.UserName = '" & gUserName & "' "
                       .MoveNext
                     Loop
                End With
             End If
             If GProcRstOpen(tmp1, "select * from tmpRptTbl where tmpRptTbl.UserName ='" & gUserName & "'", "R", gCn) > 0 Then
                With tmp1
                     .MoveFirst
                     Do While Not .EOF
                        '-- For Matched Data
                        gCn.Execute "Update TmpGentbl set TmpGentbl.Nar2 = 'Match' where TmpGentbl.Nar1 = '" & tmp1.Fields("Nar1") & "'  and TmpGentbl.VDt = Cdate('" & Format(tmp1.Fields("Vdt"), "yyyy,mm,dd") & "') and  " _
                            & " TmpGentbl.amt3 = " & tmp1.Fields("Amt3") & " and TmpGentbl.amt4 = " & tmp1.Fields("Amt4") & " and TmpGentbl.amt5 = " & tmp1.Fields("Amt5") & " and TmpGentbl.UserName = '" & gUserName & "' "
                        '-- For MisMatched Data
                        gCn.Execute "Update TmpGentbl set TmpGentbl.Nar2 = 'A - Mismatch',TmpGentbl.amt6 = " & tmp1.Fields("Amt1") & ",TmpGentbl.amt7 = " & tmp1.Fields("Amt2") & ",TmpGentbl.amt8 = " & tmp1.Fields("Amt3") & ",TmpGentbl.amt9 = " & tmp1.Fields("Amt4") & ",TmpGentbl.amt10 = " & tmp1.Fields("Amt5") & "" _
                            & " where TmpGentbl.Nar1 = '" & tmp1.Fields("Nar1") & "' and TmpGentbl.Nar4 = '" & tmp1.Fields("Nar4") & "' and TmpGentbl.VDt = Cdate('" & Format(tmp1.Fields("Vdt"), "yyyy,mm,dd") & "')  and  " _
                            & " ( TmpGentbl.amt3 <> " & tmp1.Fields("Amt3") & " or TmpGentbl.amt4 <> " & tmp1.Fields("Amt4") & " or TmpGentbl.amt5 <> " & tmp1.Fields("Amt5") & ") and TmpGentbl.UserName = '" & gUserName & "' "
                       .MoveNext
                     Loop
                End With
             End If
             gCn.Execute "delete from tmpGenTbl where tmpGenTbl.Nar2 = 'Match' and tmpGenTbl.UserName = '" & gUserName & "' "
             gCn.Execute "delete from tmpRptTbl where tmpRptTbl.Nar2 = 'Match' and tmpRptTbl.UserName = '" & gUserName & "' "
             gCn.Execute "delete from tmpRptTbl where tmpRptTbl.Nar2 = 'A - Mismatch' and tmpRptTbl.UserName = '" & gUserName & "' "
             '-- For Credit Not Received
             gCn.Execute "Update tmpGenTbl set Nar2 = 'B - GST Not Received From Supplier'  where tmpGenTbl.UserName = '" & gUserName & "' and (isnull(nar2)= TRUE or nar2 = '') "
             gCn.Execute "Update tmpRptTbl set Nar2 = 'C - GST Received But Not In Our Books' where tmpRptTbl.UserName = '" & gUserName & "' and (isnull(nar2) = TRUE or nar2 = '') "
             gCn.BeginTrans
             X1 = "Select vno,Nar6,vyear,Nar1,nar2,vtype,Nar7,Nar4,Vdt,Amt1,Amt2,Amt3,Amt4,Amt5,UserName from tmpRptTbl where username='" & gUserName & "'"
             GProcRstOpen tmp1, X1, "O", gCn
             gCn.Execute " Insert Into TmpGentbl (vno,Nar6,vyear,Nar1,Nar2,vtype,Nar7,Nar4,Vdt,Amt1,Amt2,Amt3,Amt4,Amt5,UserName) " & X1
             gCn.CommitTrans
             '---Bill date check is pending
             X1 = "{TmpGentbl.UserName}='" & gUserName & "'"
            .Formulas(4) = "wFrToDt = 'From : ' & '" & Format(dtpFrDt.Value, "dd/MM/yyyy") & "' & '  To ' & '" & Format(dtpToDt, "dd/MM/yyyy") & "'"
            .Formulas(5) = "wCompGstin = '" & gCGSTIN & "'"
            .SelectionFormula = X1
            .ReportFileName = gReportPath & "rptFormGSTR-2AMisMatch.rpt"
               
        End Select
    .Action = 1
End With
End Sub


