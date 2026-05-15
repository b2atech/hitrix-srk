VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form frmRptRegi 
   BackColor       =   &H80000013&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Register"
   ClientHeight    =   6525
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10215
   Icon            =   "frmRptRegi.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6525
   ScaleWidth      =   10215
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
      Index           =   3
      Left            =   3480
      TabIndex        =   40
      Text            =   "3"
      Top             =   4320
      Visible         =   0   'False
      Width           =   645
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
      Index           =   2
      Left            =   1680
      TabIndex        =   39
      Text            =   "2"
      Top             =   4320
      Visible         =   0   'False
      Width           =   645
   End
   Begin VB.CheckBox Chkonly 
      BackColor       =   &H80000018&
      Caption         =   "Only 2nd Day"
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
      Left            =   8115
      TabIndex        =   38
      Top             =   1770
      Width           =   1875
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
      Width           =   3570
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
      Width           =   3570
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
      Left            =   7455
      TabIndex        =   25
      Top             =   3600
      Width           =   2610
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
      Left            =   7455
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
      Left            =   7455
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
      Left            =   1185
      Picture         =   "frmRptRegi.frx":058A
      Style           =   1  'Graphical
      TabIndex        =   29
      Top             =   5625
      Width           =   1185
   End
   Begin VB.CommandButton cmdCrystal 
      Height          =   480
      Index           =   1
      Left            =   2610
      Picture         =   "frmRptRegi.frx":09CC
      Style           =   1  'Graphical
      TabIndex        =   30
      ToolTipText     =   "Print"
      Top             =   5625
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
      Top             =   735
      Width           =   3960
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H80000018&
      Height          =   1200
      Left            =   5715
      TabIndex        =   26
      Top             =   2820
      Width           =   1635
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
         Width           =   1425
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
         Width           =   1425
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
         Width           =   1335
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
      Left            =   4035
      TabIndex        =   31
      Top             =   5625
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
      Top             =   4920
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
      Format          =   47906819
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
      Format          =   47906819
      CurrentDate     =   38050
   End
   Begin Crystal.CrystalReport CrystalReport1 
      Bindings        =   "frmRptRegi.frx":1036
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
   Begin VB.Label Label13 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "To Days   :"
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
      Left            =   2400
      TabIndex        =   42
      Top             =   4320
      Visible         =   0   'False
      Width           =   960
   End
   Begin VB.Label Label12 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "From   :"
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
      Left            =   840
      TabIndex        =   41
      Top             =   4320
      Visible         =   0   'False
      Width           =   690
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
      Left            =   885
      Shape           =   4  'Rounded Rectangle
      Top             =   5520
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
      Left            =   1680
      TabIndex        =   33
      Top             =   5880
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
      Height          =   4725
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
Attribute VB_Name = "frmRptRegi"
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
'==== Outstanding Report
Case "Outstanding Broker with Partywise", "Mill Bill Outstanding Broker with Partywise", "Outstanding Brokerwise", "Pending GST Amt In L.P. Int "
    If gSelectedMenu = "Outstanding Broker with Partywise" Then
       ChkStkValTax.Enabled = True
       ChkStkValTax.Caption = "Sellected Firm"
    End If
    ChkConsolidate.Enabled = True
    ChkConsolidate = 1
    txttemp(0).Enabled = True '-Mill
    txttemp(2).Enabled = True '-Party
    txttemp(3).Enabled = True '-Broker
    Chkonly.Enabled = True
    Label12.Visible = True
    Label13.Visible = True
    txt(2) = 0
    txt(2).Visible = True
    txt(2).Enabled = True
    txt(3) = 0
    txt(3).Visible = True
    txt(3).Enabled = True

    '-- Unadjusted Receipt
    If gSelectedMenu = "Outstanding Broker with Partywise" Then ChkUnAdjustRec.Enabled = True
Case "Purchase Outstanding Partywise" '-"Purchase Outstanding Partywise"
    ChkConsolidate.Enabled = True
    ChkConsolidate = 1
    txttemp(2).Enabled = True '-Party
    txttemp(3).Enabled = True '-Broker
Case "Unadjusted Recepit", "Unadjusted Payment" '"Unadjusted Recepit ,"Unadjusted Payment"
    txttemp(2).Enabled = True '--Party
Case "Outstanding Late Payment Interest Report", "Outstanding Late Payment Interest Report (Mill Bill)" '-"Outstanding Late Payment Interest Report"
    txttemp(2).Enabled = True '-Party
    txttemp(3).Enabled = True '-Broker
    ChkConsolidate.Enabled = True
    ChkConsolidate = 1
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
Case "Sale In Transit Sale Register" '--"Sale In Transit Sale Register"
    OptDetail.Enabled = True
    OptShort.Enabled = True
    Chkonly.Enabled = True
    txttemp(0).Enabled = True 'Mill
    txttemp(2).Enabled = True '-Party
    txttemp(5).Enabled = True '-Count
    txttemp(3).Enabled = True '-Broker
    chkMonthwise.Enabled = True
    ChkUnAdjustRec.Caption = "Only E2 Sales Register"
    ChkUnAdjustRec.Enabled = True
    ChkConsolidate.Enabled = True
    ChkConsolidate.Caption = "New Register (G.P.)"
Case "Consignment/Depot Sale Register" '--"Consignment/Depot Sale Register":
    OptDetail.Enabled = True
    OptShort.Enabled = True
    Chkonly.Enabled = True
    txttemp(0).Enabled = True '-Mill
    txttemp(2).Enabled = True '-Party
    txttemp(5).Enabled = True '-Count
    txttemp(1).Enabled = True '-Tax
    txttemp(3).Enabled = True '-Broker
    chkMonthwise.Enabled = True
Case "Mill Bill Register" '--"Mill Bill Register"
    txttemp(0).Enabled = True '--Mill
    txttemp(2).Enabled = True '-Party
    txttemp(5).Enabled = True '-Count
    txttemp(3).Enabled = True '-Broker
    Chkonly.Enabled = True
    chkMonthwise.Enabled = True
    
Case "Receipt Register", "Payment Register"
     OptDetail.Caption = "Bank"
     OptShort.Caption = "Cash"
     OptDetail.Enabled = True
     OptShort.Enabled = True
     lblMillName.Caption = "Bank"
     txttemp(0).Enabled = True
     txttemp(2).Enabled = True
    
'==== Purchase Register
Case "Trade Purchase Register" '--"Trade Purchase Register"
    OptDetail.Enabled = True
    OptShort.Enabled = True
    txttemp(5).Enabled = True '-Count
    txttemp(0).Enabled = True '-Mill
    txttemp(1).Enabled = True '-Tax
    txttemp(2).Enabled = True '-Party
    txttemp(3).Enabled = True '-Broker
    chkMonthwise.Enabled = True
    opsummary.Visible = True
    opsummary.Enabled = True
Case "Purchase Register For ITC"
    txttemp(0).Enabled = True '-Mill
    txttemp(2).Enabled = True '-Party
    txttemp(1).Enabled = True '-Tax
Case "SIT Purchase Register" '--"SIT Purchase Register"
    OptDetail.Enabled = True
    OptShort.Enabled = True
    txttemp(5).Enabled = True '-Count
    txttemp(0).Enabled = True  '--Mill
    txttemp(2).Enabled = True '-Party
    txttemp(3).Enabled = True '-Broker
    chkMonthwise.Enabled = True
    ChkUnAdjustRec.Caption = "Only E2 Sales Register"
    ChkUnAdjustRec.Enabled = True
Case "Insurance Report(SIT)"
    txttemp(0).Enabled = True '-Party
    txttemp(2).Enabled = True '-Mill
    txt(1).Enabled = True
    Label6.Caption = "Sr. No. :"

Case "Other Purchase Register" '--"Other Purchase Register"
    txttemp(1).Enabled = True '--Tax
    txttemp(2).Enabled = True '-Party
    txttemp(3).Enabled = True '-Broker
    chkMonthwise.Enabled = True
Case "Consignment/Depot Purchase Register" '-"Consignment/Depot Purchase Register":
    txttemp(0).Enabled = True '--Mill
    txttemp(3).Enabled = True '-Broker
    'txttemp(5).Enabled = True '-Count
    chkIsDespatchDt.Enabled = True '--Report for despatch date
    chkMonthwise.Enabled = True
'==== Booking Register
Case "Mill Booking Register", "Pending Mill Booking Register", "Canceled Mill Booking Register", "Mill Booking Vs Dispatch Register" '-"Mill Booking Register"
    txttemp(0).Enabled = True '--Mill
    txttemp(5).Enabled = True '-Item
    txttemp(3).Enabled = True '-Broker
Case "Mill Booking Register", "Pending Mill Booking Register", "Canceled Mill Booking Register", "Mill Booking Vs Dispatch Register", "Trade Purchase Booking Register", "Pending Trade Purchase Booking Register" '-"Mill Booking Register"
    txttemp(0).Enabled = True '--Mill
    txttemp(5).Enabled = True '-Item
    txttemp(3).Enabled = True '-Broker
    If gSelectedMenu = "Trade Purchase Booking Register" Or gSelectedMenu = "Pending Trade Purchase Booking Register" Then
       ChkConsolidate.Enabled = True
       ChkConsolidate.Caption = "Is Millwise Report"
    End If
    If gSelectedMenu = "Pending Trade Purchase Booking Register" Then
       chkMonthwise.Enabled = True
       chkMonthwise.Caption = "Is Countwise Report"
    End If
    
Case "SIT Booking Register", "Pending SIT Booking Register", "Canceled SIT Booking Register", "SIT Booking Vs Dispatch Register", "Trade Booking Vs Dispatch Register" '-"SIT Booking Register"
    txttemp(0).Enabled = True '--Mill
    txttemp(5).Enabled = True '-Item
    txttemp(2).Enabled = True '-Party
    
Case "Consignment/Depot Booking Register", "Pending Consignment/Depot Booking Register", "Canceled Consignment/Depot Booking Register", "Consignment/Depot Booking Vs Dispatch Register" '-"Consignment/Depot Booking Register"
    txttemp(0).Enabled = True '--Mill
Case "Party Booking Register", "Pending Party Booking Register", "Canceled Party Booking Register", "Party Booking Vs Dispatch Register" '-"Party Booking Register"
    txttemp(2).Enabled = True '-Party
    txttemp(3).Enabled = True '-Broker
    txttemp(0).Enabled = True '--Mill
    txttemp(5).Enabled = True '-Item
'-====== Other Report
Case "Mill Bill Receipt Register" '-"Mill Bill Receipt Register"
     txttemp(0).Enabled = True '-Mill
     txttemp(2).Enabled = True '-Party
Case "Sale In Transit Register", "Trade Purchase - Sale Register" '-"Sale In Transit Register"
    txttemp(0).Enabled = True '-Mill
    txttemp(5).Enabled = True '-Item
    txttemp(2).Enabled = True '-Party
    ChkConsolidate.Enabled = True
    ChkConsolidate.Caption = "New Report"
    chkMonthwise.Enabled = True
Case "Late Pay Debit Note Register", "Late Pay Debit Note Register (Mill Bill)" '"Late Pay Debit Note Register"
    txttemp(2).Enabled = True '-Party
    txttemp(3).Enabled = True '-Broker
    ChkConsolidate.Enabled = True
    ChkConsolidate = 1
Case "Item Ledger" '-"Item Ledger"
    txttemp(5).Enabled = True '-Item
    txttemp(0).Enabled = True '-Mill
Case "Daily Stock To Mill" '-"Daily Stock To Mill"
    txttemp(0).Enabled = True '-Mill
    dtpFrDt.Enabled = False '-On Date
    txt(1).Enabled = True '-C. C. To
Case "Trade & SIT Stock Valuation Register", "Trade & SIT Stock Valuation Summary" '-"Trade & SIT Stock Valuation
    txttemp(5).Enabled = True '-Item
    txttemp(0).Enabled = True '-Mill
    dtpFrDt.Enabled = False
    ChkStkValTax.Enabled = True '-Tax Amt in Valuation
    ChkStkValExpense.Enabled = True '-- Expenses Amt In Valuation
'====Godown Reports
Case "Pending Gate Pass Report" '--"Pending Gate Pass Report"
    txttemp(2).Enabled = True '-Party
    txttemp(4).Enabled = True '-Godown
    txttemp(5).Enabled = True '-Item
    txttemp(0).Enabled = True '-Mill
    txttemp(3).Enabled = True '-Broker
Case "Gate Pass Register" '-"Gate Pass Register"
    txttemp(2).Enabled = True '-Party
    txttemp(0).Enabled = True '-Mill
    txttemp(5).Enabled = True '-Item
   ' txttemp(4).Enabled = True '-Godown
Case "Delivery Register" '-"Delivery Register"
    txttemp(2).Enabled = True '-Party
    txttemp(4).Enabled = True '-Godown
    txttemp(0).Enabled = True '-Mill
    txttemp(5).Enabled = True '-Item
    
Case "Godownwise Inward Register" '"Godownwise Inward Register"
    txttemp(5).Enabled = True '-Item
    txttemp(4).Enabled = True '-Godown
    txttemp(0).Enabled = True '-Mill
Case "Countwise Goods Delivery Report", "Count & Partywise Goods Delivery Report", "Count & Sizerwise Goods Delivery Report" '"Countwise Goods Delivery Report"
    txttemp(5).Enabled = True '-Item
    txttemp(2).Enabled = True '-Party
    txttemp(6).Enabled = True '-Sizer
Case "Party and Bag No wise Goods Delivery Report" '-"Party and Bag No wise Goods Delivery Report"
    txttemp(5).Enabled = True '-Item
    txttemp(2).Enabled = True '-Party
   
Case "Payment/RTGS Register"
     OptDetail.Enabled = True
     OptShort.Enabled = True
     OptShort.Caption = "Summary"
     lblMillName.Caption = "Bank"
     txttemp(0).Enabled = True
     txttemp(2).Enabled = True
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
X1 = ""
If txttemp(0) = "" Then wMillCode = 0 '- Mill Code
With CrystalReport1
    GProcCrystalRptPreparation CrystalReport1, IIf(Index = 0, "Window", "Printer")
    Select Case gSelectedMenu

        Case "Credit Note Register" '-"Credit Note Register"
            .Formulas(4) = "wFrToDt='From : '& '" & dtpFrDt.Value & "' & ' To ' & '" & dtpToDt.Value & "'"
            .Formulas(5) = "RptName ='" & gSelectedMenu & "'"
            .ReportFileName = gReportPath & "RptRegVoucher.rpt"
            X1 = "{TBLVoucher.VTYPE}='PN' and {TBLVoucher.VCTRNO}=1 and {TBLVoucher.VDt} in Date(" & Format(dtpFrDt.Value, "yyyy,mm,dd") & ") to Date(" & Format(dtpToDt.Value, "yyyy,mm,dd") & ")  "
            .SelectionFormula = X1
        Case "Debit Note Register" '-"Debit Note Register"
            .Formulas(4) = "wFrToDt='From : '& '" & dtpFrDt.Value & "' & ' To ' & '" & dtpToDt.Value & "'"
            .Formulas(5) = "RptName ='" & gSelectedMenu & "'"
            .ReportFileName = gReportPath & "RptRegVoucher.rpt"
            X1 = "{TBLVoucher.VTYPE}='SN' and {TBLVoucher.VCTRNO}=1 and {TBLVoucher.VDt} in Date(" & Format(dtpFrDt.Value, "yyyy,mm,dd") & ") to Date(" & Format(dtpToDt.Value, "yyyy,mm,dd") & ") "
            .SelectionFormula = X1
        Case "Receipt Register"
            .Formulas(4) = "wFrToDt='From : '& '" & dtpFrDt.Value & "' & ' To ' & '" & dtpToDt.Value & "'"
            .Formulas(6) = "wBank =  '" & txttemp(0) & "' "
            .ReportFileName = gReportPath & "RptReceiptReg.rpt"
            If OptDetail.Value = True Then '-Bank
              .Formulas(5) = "RptName ='Bank Receipt Register '"
               X1 = "{qryRecPayDetails.VTYPE}='BR' and {qryRecPayDetails.VDt} in Date(" & Format(dtpFrDt.Value, "yyyy,mm,dd") & ") to Date(" & Format(dtpToDt.Value, "yyyy,mm,dd") & ") "
               If txttemp(0) <> "" Then X1 = X1 & " and {qryRecPayDetails.VacOCode}=" & wMillCode
               If txttemp(2) <> "" Then X1 = X1 & " and {qryRecPayDetails.VacCode}=" & wPartyCode
            Else '--Cash
              .Formulas(5) = "RptName ='Cash Receipt Register '"
               X1 = "{qryRecPayDetails.VTYPE}='CR' and {qryRecPayDetails.VDt} in Date(" & Format(dtpFrDt.Value, "yyyy,mm,dd") & ") to Date(" & Format(dtpToDt.Value, "yyyy,mm,dd") & ") "
               If txttemp(2) <> "" Then X1 = X1 & " and {qryRecPayDetails.VacCode}=" & wPartyCode
            End If
            .SelectionFormula = X1
            
        Case "Payment/RTGS Register"
            .Formulas(4) = "wFrToDt='From : '& '" & dtpFrDt.Value & "' & ' To ' & '" & dtpToDt.Value & "'"
            .Formulas(6) = "wBank =  '" & txttemp(0) & "' "
            .Formulas(7) = "wcBnkac= '" & gCBankAcNo & "'"
            If OptDetail.Value = True Then
                .ReportFileName = gReportPath & "RptPaymentReg.rpt"
            Else
                .ReportFileName = gReportPath & "RptPaymentSummary.rpt"
            End If
            .Formulas(5) = "RptName ='Bank Payment Register '"
            X1 = "{tblVoucher.VTYPE}='BP' and {tblVoucher.VCtrNo}= 1 and left({tblVoucher.Vnar1},4) ='NEFT' and {tblVoucher.VDt} in Date(" & Format(dtpFrDt.Value, "yyyy,mm,dd") & ") to Date(" & Format(dtpToDt.Value, "yyyy,mm,dd") & ") "
            If txttemp(0) <> "" Then X1 = X1 & " and {tblVoucher.VacCode}=" & wMillCode
            If txttemp(2) <> "" Then X1 = X1 & " and {tblVoucher.VacOCode}=" & wPartyCode
            .SelectionFormula = X1
            
        Case "Internal Voucher Register" '-"Internal Voucher Register"
            .Formulas(4) = "wFrToDt='From : '& '" & dtpFrDt.Value & "' & ' To ' & '" & dtpToDt.Value & "'"
            .Formulas(5) = "RptName ='" & gSelectedMenu & "'"
            .ReportFileName = gReportPath & "RptRegVoucher.rpt"
            X1 = "{TBLVoucher.VTYPE}='JV' and {TBLVoucher.VCTRNO}=1 and {TBLVoucher.VDt} in Date(" & Format(dtpFrDt.Value, "yyyy,mm,dd") & ") to Date(" & Format(dtpToDt.Value, "yyyy,mm,dd") & ") "
            .SelectionFormula = X1
         Case "Mill Bill Receipt Register" '-"Mill Bill Receipt Register"
            .Formulas(4) = "wFrToDt='From : '& '" & dtpFrDt.Value & "' & ' To ' & '" & dtpToDt.Value & "'"
            .ReportFileName = gReportPath & "RptRegMillBillReceipt.rpt"
            X1 = "{tblMillRecPay.VTYPE}='MR' and {tblMillRecPay.VCTRNO}=1 and {tblMillRecPay.VDt} in Date(" & Format(dtpFrDt.Value, "yyyy,mm,dd") & ") to Date(" & Format(dtpToDt.Value, "yyyy,mm,dd") & ") "
            '-- Mill
            If txttemp(0) <> "" Then X1 = X1 & " and {tblMillRecPay.MRECACOCODE}=" & wMillCode
            '-- Party
            If txttemp(2) <> "" Then X1 = X1 & " and {tblMillRecPay.MRECACCODE}=" & wPartyCode
            .SelectionFormula = X1
        '======== Sale Register
        Case "Trade Sale Register" '--"Trade Sale Register"
            .Formulas(4) = "wFrToDt='From : '& '" & dtpFrDt.Value & "' & ' To ' & '" & dtpToDt.Value & "'"
            X1 = "{tblAddLess.VType}='SY' and {tblAddLess.VDt} in Date(" & Format(dtpFrDt.Value, "yyyy,mm,dd") & ") to Date(" & Format(dtpToDt.Value, "yyyy,mm,dd") & ")  " '
            '-- Party
            If txttemp(2) <> "" Then X1 = X1 + " and {tblAddLess.AdAcDrCode} = " & wPartyCode
            '-- Broker
            If txttemp(3) <> "" Then X1 = X1 + " and {tblOutstanding.OutBrokerCode} = " & wBrkCode
            '-- Count
            If txttemp(5) <> "" Then X1 = X1 + " and {tblVouSub.VSubITCode} = " & wItemCode
            '----- 2nd day register
            If Chkonly = 1 Then X1 = X1 & " and {tblOutstanding.OutDueDays}=" & 1
            If Trim(txttemp(1)) <> "" Then '-- Tax
                X1 = X1 + " and {tblAddLess.AdTaxCode} = " & wNarrCode & ""
                X = X + " Tax Type : " & txttemp(1)
            End If
            .Formulas(5) = "wTaxType='" & X & "'"
            If OptDetail.Value = True Then
                '--- Mill
                If txttemp(0) <> "" Then X1 = X1 + " and {tblMastItem.ITMillCode} = " & wMillCode
                If chkMonthwise = 1 Then '--MonthWise
                    .ReportFileName = gReportPath & "rptTradeSaleRegiMon.rpt"
                Else
                    .ReportFileName = gReportPath & "rptTradeSaleRegi.rpt" '--Detail
                End If
                .SelectionFormula = X1
            ElseIf OptShort.Value = True Then
                 .SelectionFormula = X1
                 If chkMonthwise = 1 Then '--MonthWise
                     .ReportFileName = gReportPath & "rptTradeSaleRegiSimMon.rpt" '-Short
                 Else
                     .ReportFileName = gReportPath & "rptTradeSaleRegiSim.rpt" '-Short
                 End If
             Else
                     .SelectionFormula = X1
                     .ReportFileName = gReportPath & "rptTradeSaleRegisumm.rpt" '-Short
            End If
        Case "Sale In Transit Sale Register": '--"Sale In Transit Sale Register":
            .Formulas(4) = "wFrToDt='From : '& '" & dtpFrDt.Value & "' & ' To ' & '" & dtpToDt.Value & "'"
            X1 = "{tblAddLess.VType}='ST' and {tblAddLess.VDt} in Date(" & Format(dtpFrDt.Value, "yyyy,mm,dd") & ") to Date(" & Format(dtpToDt.Value, "yyyy,mm,dd") & ")  " '
            '-- Party
            If txttemp(2) <> "" Then X1 = X1 + " and {tblAddLess.AdAcDrCode} = " & wPartyCode
            '-- Broker
            If txttemp(3) <> "" Then X1 = X1 + " and {tblOutstanding.OutBrokerCode} = " & wBrkCode
            '-- Count
            If txttemp(5) <> "" Then X1 = X1 + " and {tblVouSub.VSubITCode} = " & wItemCode
            '----- 2nd day register
            If Chkonly = 1 Then X1 = X1 & " and {tblOutstanding.OutDueDays}=" & 2
            If Trim(txttemp(0)) <> "" Then '--Mill
                X1 = X1 + " and {tblAddLess.AdMillCode} = " & wMillCode & ""
                .Formulas(6) = "wMillName='Mill Name : ' & '" & txttemp(0) & "'"
            End If
            '---- Only E2 Sale
            If ChkUnAdjustRec = 1 Then
               X1 = X1 + " and {tblAddLess.adtmp2} = 1"
           ' Else
            '   X1 = X1 + " and {tblAddLess.adtmp2} <> 1"
            End If

            .SelectionFormula = X1
            If OptDetail.Value = True Then '-Detail
                If chkMonthwise = 1 Then '--MonthWise
                    .ReportFileName = gReportPath & "rptSITSaleRegiMon.rpt"
                Else
                    If ChkConsolidate.Value = 1 Then
                       .ReportFileName = gReportPath & "rptSITSaleRegiNew.rpt"
                    Else
                       .ReportFileName = gReportPath & "rptSITSaleRegi.rpt"
                    End If
                End If
            Else '-Short
                If chkMonthwise = 1 Then '--MonthWise
                    .ReportFileName = gReportPath & "rptSITSaleRegiSimMon.rpt"
                Else
                    .ReportFileName = gReportPath & "rptSITSaleRegiSim.rpt"
                End If
            End If
        Case "Consignment/Depot Sale Register": '--"Consignment/Depot Sale Register":
            .Formulas(4) = "wFrToDt='From : '& '" & dtpFrDt.Value & "' & ' To ' & '" & dtpToDt.Value & "'"
            X1 = "({tblAddLess.VType} = 'SO' or {tblAddLess.VType} = 'SD') and {tblAddLess.VDt} in Date(" & Format(dtpFrDt.Value, "yyyy,mm,dd") & ") to Date(" & Format(dtpToDt.Value, "yyyy,mm,dd") & ")  " '
            If Trim(txttemp(1)) <> "" Then '--Tax
                X1 = X1 + " and {tblAddLess.AdTaxCode} = " & wNarrCode & ""
                .Formulas(5) = "wTaxType='Tax Type : ' & '" & txttemp(1) & "'"
            End If
            If Trim(txttemp(0)) <> "" Then '--mill
                X1 = X1 + " and {tblAddLess.AdMillCode} = " & wMillCode & ""
                .Formulas(6) = "wMillName='Mill Name : ' & '" & txttemp(0) & "'"
            End If
            '----- 2nd day register
            If Chkonly = 1 Then X1 = X1 & " and {tblOutstanding.OutDueDays}=" & 1
            '-- Party
            If txttemp(2) <> "" Then X1 = X1 + " and {tblAddLess.AdAcDrCode} = " & wPartyCode
            '-- Broker
            If txttemp(3) <> "" Then X1 = X1 + " and {tblOutstanding.OutBrokerCode} = " & wBrkCode
            '-- Count
            If txttemp(5) <> "" Then X1 = X1 + " and {tblVouSub.VSubITCode} = " & wItemCode
            .SelectionFormula = X1
            If OptDetail.Value = True Then '--Detail
                If chkMonthwise = 1 Then '--MonthWise
                    .ReportFileName = gReportPath & "rptConsiDepotSaleRegiMon.rpt"
                Else
                    .ReportFileName = gReportPath & "rptConsiDepotSaleRegi.rpt"
                End If
            Else '-Short
                If chkMonthwise = 1 Then '--MonthWise
                    .ReportFileName = gReportPath & "rptConsiDepotSaleRegiSimMon.rpt"
                Else
                    .ReportFileName = gReportPath & "rptConsiDepotSaleRegiSim.rpt"
                End If
            End If
        Case "Mill Bill Register": '--"Mill Bill Register":
            .Formulas(4) = "wFrToDt='From : '& '" & dtpFrDt.Value & "' & ' To ' & '" & dtpToDt.Value & "'"
            X1 = " {tblAddLess.VType}='SM' and {tblAddLess.VDt} in Date(" & Format(dtpFrDt.Value, "yyyy,mm,dd") & ") to Date(" & Format(dtpToDt.Value, "yyyy,mm,dd") & ")  " '
            If Trim(txttemp(0)) <> "" Then '-Mill
                X1 = X1 + " and {tblAddLess.AdMillCode} = " & wMillCode & ""
                .Formulas(6) = "wMillName='Mill Name : ' & '" & txttemp(0) & "'"
            End If
            '----- 2nd day register
            If Chkonly = 1 Then X1 = X1 & " and {tblOutstanding.OutDueDays}=" & 1
            '-- Party
            If txttemp(2) <> "" Then X1 = X1 + " and {tblAddLess.AdAcDrCode} = " & wPartyCode
            '-- Broker
            If txttemp(3) <> "" Then X1 = X1 + " and {tblOutstanding.OutBrokerCode} = " & wBrkCode
            '-- Count
            If txttemp(5) <> "" Then X1 = X1 + " and {tblVouSub.VSubITCode} = " & wItemCode
            .SelectionFormula = X1
            If chkMonthwise = 1 Then '--MonthWise
                .ReportFileName = gReportPath & "rptMillBillRegiMon.rpt"
            Else
                .ReportFileName = gReportPath & "rptMillBillRegi.rpt"
            End If
        '======== Purchase Register
        Case "Trade Purchase Register": '--"Trade Purchase Register":
            .Formulas(4) = "wFrToDt='From : '& '" & dtpFrDt.Value & "' & ' To ' & '" & dtpToDt.Value & "'"
            X1 = "{tblAddLess.VType}='PY' and {tblAddLess.VDt} in Date(" & Format(dtpFrDt.Value, "yyyy,mm,dd") & ") to Date(" & Format(dtpToDt.Value, "yyyy,mm,dd") & ")  " '
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
            If OptDetail.Value = True Then '-Detail
            '-- Broker
            If txttemp(3) <> "" Then X1 = X1 + " and {tblOutstanding.OutBrokerCode} = " & wBrkCode
                 '-- Count
                If txttemp(5) <> "" Then X1 = X1 + " and {tblVouSub.VSubITCode} = " & wItemCode
                .SelectionFormula = X1
                If chkMonthwise = 1 Then '--MonthWise
                    .ReportFileName = gReportPath & "rptTradePurchRegiMon.rpt"
                Else
                    .ReportFileName = gReportPath & "rptTradePurchRegi.rpt"
                End If
            ElseIf OptShort.Value = True Then '-Short
                .SelectionFormula = X1
                 If chkMonthwise = 1 Then '--MonthWise
                    .ReportFileName = gReportPath & "rptTradePurchRegiSimMon.rpt"
                 Else
                    .ReportFileName = gReportPath & "rptTradePurchRegiSim.rpt"
                 End If
             Else
                 .SelectionFormula = X1
                    .ReportFileName = gReportPath & "rptTradePurchRegiSumm.rpt"
            End If
        Case "Purchase Register For ITC":
            .Formulas(4) = "wFrToDt='From : '& '" & dtpFrDt.Value & "' & ' To ' & '" & dtpToDt.Value & "'"
            X1 = "{tblAddLess.VType}='PY' and {tblAddLess.AdgpDt} in Date(" & Format(dtpFrDt.Value, "yyyy,mm,dd") & ") to Date(" & Format(dtpToDt.Value, "yyyy,mm,dd") & ")  " '
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
            .SelectionFormula = X1
            .ReportFileName = gReportPath & "rptPurchRegiITC.rpt"
            
        Case "SIT Purchase Register": '-"SIT Purchase Register":
            .Formulas(4) = "wFrToDt='From : '& '" & dtpFrDt.Value & "' & ' To ' & '" & dtpToDt.Value & "'"
            X1 = "{tblAddLess.VType}='PT' and {tblAddLess.VDt} in Date(" & Format(dtpFrDt.Value, "yyyy,mm,dd") & ") to Date(" & Format(dtpToDt.Value, "yyyy,mm,dd") & ")  " '
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
            If OptDetail.Value = True Then '-Detail
            '-- Broker
            If txttemp(3) <> "" Then X1 = X1 + " and {tblOutstanding.OutBrokerCode} = " & wBrkCode
                '-- Count
            If txttemp(5) <> "" Then X1 = X1 + " and {tblVouSub.VSubITCode} = " & wItemCode
            '---- Only E2 Sale
            If ChkUnAdjustRec = 1 Then
               X1 = X1 + " and {tblAddLess.AdGpNo} = 1"
            'Else
            '   X1 = X1 + " and {tblAddLess.AdGpNo} <> 1"
            End If

                If chkMonthwise = 1 Then '--MonthWise
                    .ReportFileName = gReportPath & "rptSITPurchRegiMon.rpt"
                Else
                    .ReportFileName = gReportPath & "rptSITPurchRegi.rpt"
                End If
                .SelectionFormula = X1
            Else '-Short
                If chkMonthwise = 1 Then '--MonthWise
                    .ReportFileName = gReportPath & "rptSITPurchRegiSimMon.rpt"
                Else
                    .ReportFileName = gReportPath & "rptSITPurchRegiSim.rpt"
                End If
                .SelectionFormula = X1
            End If
        Case "Other Purchase Register": '--"Other Purchase Register":
            .Formulas(4) = "wFrToDt='From : '& '" & dtpFrDt.Value & "' & ' To ' & '" & dtpToDt.Value & "'"
            X1 = "{tblAddLess.VType}='PO' and {tblAddLess.VDt} in Date(" & Format(dtpFrDt.Value, "yyyy,mm,dd") & ") to Date(" & Format(dtpToDt.Value, "yyyy,mm,dd") & ")  " '
            If Trim(txttemp(1)) <> "" Then '--Tax
                 X1 = X1 + " and {tblAddLess.AdTaxCode} = " & wNarrCode & ""
                 .Formulas(5) = "wTaxType='Tax Type : ' & '" & txttemp(1) & "'"
            End If
            '-- Party
            If txttemp(2) <> "" Then X1 = X1 + " and {tblAddLess.AdAcCrCode} = " & wPartyCode
            '-- Broker
            If txttemp(3) <> "" Then X1 = X1 + " and {tblOutstanding.OutBrokerCode} = " & wBrkCode
            .SelectionFormula = X1
            If chkMonthwise = 1 Then '--MonthWise
                .ReportFileName = gReportPath & "rptOtherPurchRegiMon.rpt"
            Else
                .ReportFileName = gReportPath & "rptOtherPurchRegi.rpt"
            End If
        Case "Consignment/Depot Purchase Register": '--"Consignment/Depot Purchase Register":
            .Formulas(4) = "wFrToDt='From : '& '" & dtpFrDt.Value & "' & ' To ' & '" & dtpToDt.Value & "'"
            If chkIsDespatchDt = 0 Then '-----On Received Date
                X1 = "{tblAddLess.VType}='PI' and {tblAddLess.VDt} in Date(" & Format(dtpFrDt.Value, "yyyy,mm,dd") & ") to Date(" & Format(dtpToDt.Value, "yyyy,mm,dd") & ")  " '
                .SortFields(0) = "+{tblAddLess.VDt}"
            Else '----On Despatch Date
                X1 = "{tblAddLess.VType}='PI' and {tblAddLess.ADGPDT} in Date(" & Format(dtpFrDt.Value, "yyyy,mm,dd") & ") to Date(" & Format(dtpToDt.Value, "yyyy,mm,dd") & ")  " '
                .SortFields(0) = "+{tblAddLess.ADGPDT}"
            End If
            If Trim(txttemp(0)) <> "" Then '-mill
                X1 = X1 + " and {tblAddLess.AdMillCode} = " & wMillCode & ""
                .Formulas(6) = "wMillName='Mill Name : ' & '" & txttemp(0) & "'"
            End If
            '-- Broker
            If txttemp(3) <> "" Then X1 = X1 + " and {tblOutstanding.OutBrokerCode} = " & wBrkCode
            '-- Count
            If txttemp(5) <> "" Then X1 = X1 + " and {tblVouSub.VSubITCode} = " & wItemCode
            .SelectionFormula = X1
            If chkMonthwise = 1 Then '--MonthWise
                .ReportFileName = gReportPath & "rptConsiDepotInRegiMon.rpt"
            Else
                .ReportFileName = gReportPath & "rptConsiDepotInRegi.rpt"
            End If
        Case "Insurance Report(SIT)": '-"SIT Purchase Register":
            .Formulas(4) = "wFrToDt='From : '& '" & dtpFrDt.Value & "' & ' To ' & '" & dtpToDt.Value & "'"
            X1 = "{tblAddLess.VType}='PT' and {tblAddLess.ADTMP2} = 1 and {tblAddLess.VDt} in Date(" & Format(dtpFrDt.Value, "yyyy,mm,dd") & ") to Date(" & Format(dtpToDt.Value, "yyyy,mm,dd") & ")  " '
            If Trim(txttemp(1)) <> "" Then '-Tax
                 X1 = X1 + " and {tblAddLess.AdTaxCode} = " & wNarrCode & ""
                .Formulas(5) = "wTaxType='Tax Type : ' & '" & txttemp(1) & "'"
            End If
            If Trim(txttemp(0)) <> "" Then '-Mill
                X1 = X1 + " and {tblAddLess.AdMillCode} = " & wMillCode & ""
                .Formulas(6) = "wMillName='Mill Name : ' & '" & txttemp(0) & "'"
            End If
            .Formulas(7) = "PolicyNo='" & gCSTRegNo & "'"
            
            Dim OpStr As String
            Dim OpAmt As Double
            Dim ix As Integer
            Dim rstOp As Recordset
            Dim Sr As Integer
            OpStr = "select sum(AdBillAmt) from tblAddless where vtype = 'PT' and AdAcCrCode = " & wPartyCode & "  and tblAddLess.VDt >= " & CDate(Format("05/04/" & gCYear, "dd/mm/yyyy")) & "  and VYear = " & gCYear & "  and tblAddless.Vdt < cdate('" & dtpFrDt.Value & "')"
            
            ix = GProcRstOpen(rstOp, OpStr, "O")
            If IsNull(rstOp.Fields(0)) Then
               OpAmt = 0
            Else
               OpAmt = rstOp.Fields(0)
            End If
            DoEvents
            OpAmt = 250000000 - OpAmt
            .Formulas(8) = "OpBal=" & OpAmt & ""
            .Formulas(9) = "SrNo=" & Val(txt(1)) & ""
            
            '-- Party
             If txttemp(2) <> "" Then X1 = X1 + " and {tblAddLess.AdAcCrCode} = " & wPartyCode
             .ReportFileName = gReportPath & "rptSITPurchRegiInsu.rpt"
             .SelectionFormula = X1
            
        '====== Booking Register
        Case "Mill Booking Register", "Pending Mill Booking Register", "Canceled Mill Booking Register" '-"Mill Booking Register"
            .Formulas(4) = "wFrToDt='From : '& '" & dtpFrDt.Value & "' & ' To ' & '" & dtpToDt.Value & "'"
            X1 = "{tblAddLess.VType}='OM' and {tblAddLess.VDt} in Date(" & Format(dtpFrDt.Value, "yyyy,mm,dd") & ") to Date(" & Format(dtpToDt.Value, "yyyy,mm,dd") & ")  " '
            '-Mill
            If Trim(txttemp(0)) <> "" Then X1 = X1 + " and {tblmastAccount.AcName} = '" & txttemp(0) & "'"
            '-Count
            If Trim(txttemp(5)) <> "" Then X1 = X1 + " and {tblMastItem.ItName} = '" & txttemp(5) & "'"
            '-Broker
            If Trim(txttemp(3)) <> "" Then X1 = X1 + " and {tblmastAccount_Brk.AcName} = '" & txttemp(3) & "'"
            Select Case gSelectedMenu
                Case "Mill Booking Register"     '-- Register Rpt
                    .ReportFileName = gReportPath & "rptBookingMillBillReg.rpt"
                Case "Pending Mill Booking Register" '-- Pending Rpt
                    X1 = X1 + " and {@BalBag}<>0 and {tblvouSub.VSUBBOOKISCOMPLETED}=0"
                    .ReportFileName = gReportPath & "rptBookingMillBillPending.rpt"
                Case Else '-- Canceled Rpt
                    X1 = X1 + " and {@BalBag}<>0 and  {tblvouSub.VSUBBOOKISCOMPLETED}=1"
                    .ReportFileName = gReportPath & "rptBookingMillBillCancel.rpt"
            End Select
            .SelectionFormula = X1
        Case "Trade Purchase Booking Register", "Pending Trade Purchase Booking Register"
            .Formulas(4) = "wFrToDt='From : '& '" & dtpFrDt.Value & "' & ' To ' & '" & dtpToDt.Value & "'"
            X1 = "{tblAddLess.VType}='OT' and {tblAddLess.VDt} in Date(" & Format(dtpFrDt.Value, "yyyy,mm,dd") & ") to Date(" & Format(dtpToDt.Value, "yyyy,mm,dd") & ")  " '
            '-Mill
            If Trim(txttemp(0)) <> "" Then X1 = X1 + " and {tblmastAccount.AcName} = '" & txttemp(0) & "'"
            
            If Trim(txttemp(5)) <> "" Then X1 = X1 + " and {tblMastItem.ItName} = '" & txttemp(5) & "'"
            '-Broker
            If Trim(txttemp(3)) <> "" Then X1 = X1 + " and {tblmastAccount_Brk.AcName} = '" & txttemp(3) & "'"
            Select Case gSelectedMenu
             Case "Trade Purchase Booking Register"  '-- Register Rpt
                 If ChkConsolidate.Value = 1 Then
                    .ReportFileName = gReportPath & "rptBookingTradeRegMillwise.rpt"
                 Else
                    .ReportFileName = gReportPath & "rptBookingTradeReg.rpt"
                 End If
             Case "Pending Trade Purchase Booking Register" '-- Pending Rpt
                X1 = X1 + " and {@BalBag}<>0 and {tblvouSub.VSUBBOOKISCOMPLETED}=0"
                 If ChkConsolidate.Value = 1 Then
                     .ReportFileName = gReportPath & "rptBookingTradePurchPendingMill.rpt"
                 Else
                     If chkMonthwise.Value = 1 Then
                        .ReportFileName = gReportPath & "rptBookingTradePurchPendingCount.rpt"
                     Else
                        .ReportFileName = gReportPath & "rptBookingTradePurchPending.rpt"
                     End If
                 End If
            End Select
            .SelectionFormula = X1
            
        Case "Mill Booking Vs Dispatch Register" '"Mill Booking Vs Dispatch Register"
            .Formulas(4) = "wFrToDt='From : '& '" & dtpFrDt.Value & "' & ' To ' & '" & dtpToDt.Value & "'"
            X1 = "{tblAddLess.VType}='OM' and {tblAddLess.VDt} in Date(" & Format(dtpFrDt.Value, "yyyy,mm,dd") & ") to Date(" & Format(dtpToDt.Value, "yyyy,mm,dd") & ")  " '
            '-Mill
            If Trim(txttemp(0)) <> "" Then X1 = X1 + " and {tblmastAccount.AcName} = '" & txttemp(0) & "'"
            '-Count
            If Trim(txttemp(5)) <> "" Then X1 = X1 + " and {tblMastItem.ItName} = '" & txttemp(5) & "'"
             '-Broker
            If Trim(txttemp(3)) <> "" Then X1 = X1 + " and {tblmastAccount_Brk.AcName} = '" & txttemp(3) & "'"
            .ReportFileName = gReportPath & "rptBookingMillBillVsDispatch.rpt"
            .SelectionFormula = X1
        Case "SIT Booking Register", "Pending SIT Booking Register", "Canceled SIT Booking Register" '-"SIT Booking Register"
            .Formulas(4) = "wFrToDt='From : '& '" & dtpFrDt.Value & "' & ' To ' & '" & dtpToDt.Value & "'"
            X1 = "{tblAddLess.VType}='OS' and {tblAddLess.VDt} in Date(" & Format(dtpFrDt.Value, "yyyy,mm,dd") & ") to Date(" & Format(dtpToDt.Value, "yyyy,mm,dd") & ")  " '
            '-Mill
            If Trim(txttemp(0)) <> "" Then X1 = X1 + " and {tblmastAccount.AcName} = '" & txttemp(0) & "'"
            Select Case gSelectedMenu
             Case "SIT Booking Register"  '-- Register Rpt
                .ReportFileName = gReportPath & "rptBookingSITReg.rpt"
             Case "Pending SIT Booking Register" '-- Pending Rpt
                X1 = X1 + " and {@BalBag}<>0 and {tblvouSub.VSUBBOOKISCOMPLETED}=0"
                .ReportFileName = gReportPath & "rptBookingSITPending.rpt"
             Case Else '-- Canceled Rpt
                    X1 = X1 + " and {@BalBag}<>0 and  {tblvouSub.VSUBBOOKISCOMPLETED}=1"
                    .ReportFileName = gReportPath & "rptBookingSITCancel.rpt"
            End Select
            .SelectionFormula = X1
        Case "SIT Booking Vs Dispatch Register" '"SIT Booking Vs Dispatch Register"
            .Formulas(4) = "wFrToDt='From : '& '" & dtpFrDt.Value & "' & ' To ' & '" & dtpToDt.Value & "'"
            X1 = "{tblAddLess.VType}='OS' and {tblAddLess.VDt} in Date(" & Format(dtpFrDt.Value, "yyyy,mm,dd") & ") to Date(" & Format(dtpToDt.Value, "yyyy,mm,dd") & ")  " '
             '-Mill
            If Trim(txttemp(0)) <> "" Then X1 = X1 + " and {tblmastAccount.AcName} = '" & txttemp(0) & "'"
            .ReportFileName = gReportPath & "rptBookingSITVsDispatch.rpt"
            .SelectionFormula = X1
        Case "Trade Booking Vs Dispatch Register" '"SIT Booking Vs Dispatch Register"
            X1 = "SELECT tblAddLess.vno, tblAddLess.vdt, tblMastAccount.Acname, tblMastItem.Itname, tblVousub.vSubBag, tblVousub.VsubWt, tblAddLess_Pu.Vno, tblVousub_Pu.VsubPitctrlno, tblVouSub_Pu.Vdt, tblVouSub_Pu.vSubBag,tblVouSub_Pu.vSubWt,tblVousub.VsubRt,tblVouSub_Pu.vSubRt, '" & gUserName & "'" _
                 & " FROM tblAddLess, tblMastAccount, tblVouSub, tblMastItem, tblAddLess AS tblAddLess_Pu, tblVousub AS tblVousub_Pu " _
                 & " WHERE tblAddLess.Vtype = 'OT' and tblAddLess.AdMillCode = tblMastAccount.Accode and tblAddLess.Vno = tblVousub.Vno and tblAddLess.Vtype = tblVousub.Vtype and " _
                 & " tblAddLess.vyear = tblVouSub.Vyear and tblVousub.VsubItcode = tblMastItem.Itcode and " _
                 & " tblVousub.VsubItctrlno =   tblVousub_Pu.VsubPitctrlno and tblAddLess_Pu.Vno = tblVousub_Pu.Vno and tblAddLess_Pu.Vtype = tblVousub_Pu.Vtype and tblAddLess_Pu.vyear = tblVouSub_Pu.Vyear"
            X1 = X1 + " and tblAddLess.vdt>=cdate('" & Format(dtpFrDt.Value, "yyyy,mm,dd") & "')  and tblAddLess.Vdt<=Cdate('" & Format(dtpToDt.Value, "yyyy,mm,dd") & "')"
            gCn.BeginTrans
            gCn.Execute "delete from tmpGenTbl where username='" & gUserName & "'"
            gCn.Execute " Insert Into TmpGentbl (vno,Vdt,Nar1,Nar2,Amt1,Amt2,Amt3,Amt4,VBILLDT,Amt5,Amt6,Amt7,Amt8,UserName)" & X1
            gCn.CommitTrans
            .Formulas(4) = "wFrToDt='From : '& '" & dtpFrDt.Value & "' & ' To ' & '" & dtpToDt.Value & "'"
            .SelectionFormula = "{TmpGentbl.UserName}='" & gUserName & "'"
             If Trim(txttemp(0)) <> "" Then .SelectionFormula = .SelectionFormula + " and {tmpGenTbl.Nar1} = '" & txttemp(0) & "'"
             If Trim(txttemp(5)) <> "" Then .SelectionFormula = .SelectionFormula + " and {tmpGenTbl.Nar2} = '" & txttemp(5) & "'"
            .ReportFileName = gReportPath & "rptBookingTradeVsDispatch.rpt"
            
        Case "Consignment/Depot Booking Register", "Pending Consignment/Depot Booking Register", "Canceled Consignment/Depot Booking Register" '-"Consignment/Depot Booking Register"
            .Formulas(4) = "wFrToDt='From : '& '" & dtpFrDt.Value & "' & ' To ' & '" & dtpToDt.Value & "'"
            X1 = "{tblAddLess.VType}='OI' and {tblAddLess.VDt} in Date(" & Format(dtpFrDt.Value, "yyyy,mm,dd") & ") to Date(" & Format(dtpToDt.Value, "yyyy,mm,dd") & ")  " '
            '-Mill
            If Trim(txttemp(0)) <> "" Then X1 = X1 + " and {tblmastAccount.AcName} = '" & txttemp(0) & "'"
             Select Case gSelectedMenu
                Case "Consignment/Depot Booking Register"  '-- Register Rpt
                    .ReportFileName = gReportPath & "rptBookingInReg.rpt"
                Case "Pending Consignment/Depot Booking Register"  '-- Pending Rpt
                    X1 = X1 + " and {@BalBag}<>0 and {tblvouSub.VSUBBOOKISCOMPLETED}=0"
                    .ReportFileName = gReportPath & "rptBookingInPending.rpt"
                Case Else
                    X1 = X1 + " and {@BalBag}<>0 and {tblvouSub.VSUBBOOKISCOMPLETED}=1"
                    .ReportFileName = gReportPath & "rptBookingInCancel.rpt"
             End Select
             .SelectionFormula = X1
        Case "Consignment/Depot Booking Vs Dispatch Register" '-"Consignment/Depot Booking Vs Dispatch Register"
            .Formulas(4) = "wFrToDt='From : '& '" & dtpFrDt.Value & "' & ' To ' & '" & dtpToDt.Value & "'"
            X1 = "{tblAddLess.VType}='OI' and {tblAddLess.VDt} in Date(" & Format(dtpFrDt.Value, "yyyy,mm,dd") & ") to Date(" & Format(dtpToDt.Value, "yyyy,mm,dd") & ")  " '
            '-Mill
            If Trim(txttemp(0)) <> "" Then X1 = X1 + " and {tblmastAccount.AcName} = '" & txttemp(0) & "'"
            .ReportFileName = gReportPath & "rptBookingInVsDispatch.rpt"
            .SelectionFormula = X1
        Case "Pending Party Booking Register", "Canceled Party Booking Register" '-"Party Booking Register"
               .Formulas(4) = "wFrToDt='From : '& '" & dtpFrDt.Value & "' & ' To ' & '" & dtpToDt.Value & "'"
                X1 = "select 'A',tblAddLess_Book.vno,tblAddLess_Book.vtype,tblAddLess_Book.vyear,tblAddLess_Book.vdt,tblMastAccount.Acname ,tblMastAccount_Br.Acname ,tblMastAccount_Mill.Acname , " _
                     & " tblMastItem.ItName , tblvousub_Book.VSubRt, tblvousub_Book.VSubBag, tblvousub_Book.VSubCtrNo, tblvousub_Book.VSubItCtrlNo,'" & gUserName & "' " _
                     & " from tblAddLess_Book,tblMastAccount,tblMastAccount tblMastAccount_Br,tblMastAccount tblMastAccount_Mill,tblMastItem,tblvousub_Book " _
                     & " where tblAddLess_Book.AdAcDrCode = tblMastAccount.AcCode and tblAddLess_Book.AdAcCrCode = tblMastAccount_Br.AcCode and tblAddLess_Book.AdMillCode = tblMastAccount_Mill.AcCode and " _
                     & " tblAddLess_Book.Vno = tblvousub_Book.Vno And tblAddLess_Book.Vtype = tblvousub_Book.Vtype And tblAddLess_Book.VYear = tblvousub_Book.VYear And tblvousub_Book.VSubItcode = tblMastItem.ItCode " _
                     & " and tblAddLess_Book.vdt >= cdate('" & Format(dtpFrDt.Value, "yyyy,mm,dd") & "') and  tblAddLess_Book.vdt <= cdate('" & Format(dtpToDt.Value, "yyyy,mm,dd") & "') and tblVouSub_Book.VSubBookIsCompleted=0 and tblAddLess_Book.VType ='OP'"
                
                If Trim(txttemp(2)) <> "" Then X1 = X1 + " and tblmastAccount.AcName = '" & txttemp(2) & "'"
                If Trim(txttemp(3)) <> "" Then X1 = X1 + " and tblmastAccount_Br.AcName = '" & txttemp(3) & "'"
                If Trim(txttemp(0)) <> "" Then X1 = X1 + " and tblmastAccount_MILL.AcName = '" & txttemp(0) & "'"
                If Trim(txttemp(5)) <> "" Then X1 = X1 + " and tblMastItem.ItName = '" & txttemp(5) & "'"
                
                
                gCn.BeginTrans
                gCn.Execute "delete from tmpGenTbl2 where username='" & gUserName & "'"
                gCn.Execute " Insert Into TmpGentbl2 (Nar5,vno,vtype,vyear,Vdt,Nar1,Nar2,Nar3,Nar4,Amt1,Amt2,Amt3,Amt4,UserName)" & X1
                gCn.CommitTrans
                                 
                X1 = "select 'B',tblvousub.VSubBkItCtrlNo,sum(tblvousub.VSubBag),'" & gUserName & "' " _
                     & " from tblvousub where tblvousub.Vtype = 'GP' and  VSubBkItCtrlNo <> 0 and tblvousub.vdt >= cdate('" & Format(dtpFrDt.Value, "yyyy,mm,dd") & "') and  tblvousub.vdt <= cdate('" & Format(dtpToDt.Value, "yyyy,mm,dd") & "') " _
                     & " group by tblvousub.VSubBkItCtrlNo"
                gCn.BeginTrans
                gCn.Execute " Insert Into TmpGentbl2 (Nar5,Amt4,Amt5,UserName)" & X1
                gCn.CommitTrans
                
                
                gCn.BeginTrans
                RecYN = GProcRstOpen(tmp1, "select * from TmpGentbl2 where TmpGentbl2.UserName ='" & gUserName & "' and Nar5 = 'A' ", "O", gCn)
                If RecYN > 0 Then
                   With tmp1
                        .MoveFirst
                        Do While Not .EOF
                           BarCnt = GProcGetColumnValue("TmpGentbl2", "Amt4", .Fields("Amt4"), "N", "Amt5", "N", "username = '" & gUserName & "' and Nar5 = 'B'")
                           tmp1!Amt6 = BarCnt
                           .Update
                          .MoveNext
                        Loop
                   End With
                End If
                gCn.CommitTrans
                gCn.BeginTrans
                gCn.Execute "delete from tmpGenTbl2 where username='" & gUserName & "' and Nar5 = 'B'"
                gCn.Execute "delete from tmpGenTbl2 where username='" & gUserName & "' and Nar5 = 'A' and Amt2-Amt6 = 0"
                gCn.CommitTrans
                .Formulas(6) = "wRptOn= 'Party'"
                X1 = "{tmpGenTbl2.UserName}='" & gUserName & "'"
                .SelectionFormula = X1
                .ReportFileName = gReportPath & "rptBookingPartywisePendingNew.rpt"
        
''               .Formulas(4) = "wFrToDt='From : '& '" & dtpFrDt.Value & "' & ' To ' & '" & dtpToDt.Value & "'"
''                X1 = "{tblAddLess_Book.VType}='OP' and {tblAddLess_Book.VDt} in Date(" & Format(dtpFrDt.Value, "yyyy,mm,dd") & ") to Date(" & Format(dtpToDt.Value, "yyyy,mm,dd") & ")  " '
''                '-Party
''                If Trim(txtTemp(2)) <> "" Then X1 = X1 + " and {tblmastAccount.AcName} = '" & txtTemp(2) & "'"
''                If Trim(txtTemp(3)) <> "" Then X1 = X1 + " and {tblmastAccount_Br.AcName} = '" & txtTemp(3) & "'"
''                If Trim(txtTemp(0)) <> "" Then X1 = X1 + " and {tblmastAccount_MILL.AcName} = '" & txtTemp(0) & "'"
''                If Trim(txtTemp(5)) <> "" Then X1 = X1 + " and {tblMastItem.ItName} = '" & txtTemp(5) & "'"
''
''                Select Case gSelectedMenu
''                    Case "Pending Party Booking Register"
''                        If OptDetail.Value = True Then
''                          .Formulas(6) = "wRptOn= 'Party'"
''                           X1 = X1 + " and {@BalBag}<>0  and {tblVouSub_Book.VSubBookIsCompleted}=0"
''                          .ReportFileName = gReportPath & "rptBookingPartywisePendingNew.rpt"
''                        ElseIf OptShort.Value = True Then
''                          .Formulas(6) = "wRptOn= 'Broker'"
''                           X1 = X1 + " and {@BalBag}<>0  and {tblVouSub_Book.VSubBookIsCompleted}=0"
''                          .ReportFileName = gReportPath & "rptBookingPartywisePendingNew.rpt"
''                        ElseIf opsummary.Value = True Then
''                          .Formulas(6) = "wRptOn= 'Count'"
''                           X1 = X1 + " and {@BalBag}<>0  and {tblVouSub_Book.VSubBookIsCompleted}=0"
''                          .ReportFileName = gReportPath & "rptBookingPartywisePendingNew.rpt"
''                        Else
'''                          .ReportFileName = gReportPath & "rptBookingPartyPending.rpt"
''                        End If
''                End Select
                .SelectionFormula = X1
        Case "Party Booking Register"
            .Formulas(4) = "wFrToDt='From : '& '" & dtpFrDt.Value & "' & ' To ' & '" & dtpToDt.Value & "'"
            X1 = "{tblAddLess_Book.VType}='OP' and {tblAddLess_Book.VDt} in Date(" & Format(dtpFrDt.Value, "yyyy,mm,dd") & ") to Date(" & Format(dtpToDt.Value, "yyyy,mm,dd") & ")  " '
            '-Mill
            If Trim(txttemp(0)) <> "" Then X1 = X1 + " and {tblmastAccount.AcName} = '" & txttemp(0) & "'"
            
            If Trim(txttemp(5)) <> "" Then X1 = X1 + " and {tblMastItem.ItName} = '" & txttemp(5) & "'"
            '-Broker
            If Trim(txttemp(3)) <> "" Then X1 = X1 + " and {tblmastAccount_Brk.AcName} = '" & txttemp(3) & "'"
            .ReportFileName = gReportPath & "rptBookingPartyRegNew.rpt"
            .SelectionFormula = X1
'        Case "Party Booking Register", "Pending Party Booking Register", "Canceled Party Booking Register" '-"Party Booking Register"
'            '--- Prepare TmpGenTbl For Party Booking Register
'            GProcPrepareTmpTblForPartyBkReg dtpFrDt.Value, dtpToDt.Value
'            .Formulas(4) = "wFrToDt='From : '& '" & dtpFrDt.Value & "' & ' To ' & '" & dtpToDt.Value & "'"
'            X1 = "{tmpGenTbl.UserName}='" & gUserName & "'"
'            '-Party
'            If Trim(txtTemp(2)) <> "" Then X1 = X1 + " and {tmpGenTbl.Nar2} = '" & txtTemp(2) & "'"
'            '-Broker
'            If Trim(txtTemp(3)) <> "" Then X1 = X1 + " and {tmpGenTbl.Nar1} = '" & txtTemp(3) & "'"
'            .DataFiles(0) = App.Path & "\HIData\setup.mdb"
'             gReportPath = App.Path & "\HiReports\"
'            Select Case gSelectedMenu
'                Case "Party Booking Register"  '-- Register Rpt
'                    .ReportFileName = gReportPath & "rptBookingPartyReg.rpt"
'                Case "Canceled Party Booking Register" '- Canceled
'                    X1 = X1 + " and {@BalBag}<>0 and {tblBookingSub.BkSubIsCompleted}=1"
'                    .ReportFileName = gReportPath & "rptBookingPartyCancel.rpt"
'                Case Else  '-- Pending Rpt
'                    X1 = X1 + " and {@BalBag}<>0  and {tblBookingSub.BkSubIsCompleted}=0"
'                    .ReportFileName = gReportPath & "rptBookingPartyPending.rpt"
'            End Select
'            .SelectionFormula = X1
''        Case "Party Booking Vs Dispatch Register" '-"Party Booking Vs Dispatch Register"
''            '--- Prepare TmpGenTbl For Party Booking Vs Sales
''            GProcPrepareTmpTblForPartyBkVsSA dtpFrDt.Value, dtpToDt.Value
''            .Formulas(4) = "wFrToDt='From : '& '" & dtpFrDt.Value & "' & ' To ' & '" & dtpToDt.Value & "'"
''            X1 = "{QryBookingPartyVsSales.UserName}='" & gUserName & "'"
''            '-Party
''            If Trim(txttemp(2)) <> "" Then X1 = X1 + " and {QryBookingPartyVsSales.Party} = '" & txttemp(2) & "'"
''            '-Broker
''            If Trim(txttemp(3)) <> "" Then X1 = X1 + " and {QryBookingPartyVsSales.Broker} = '" & txttemp(3) & "'"
''            .DataFiles(0) = App.Path & "\HIData\setup.mdb"
''             gReportPath = App.Path & "\HiReports\"
''            .ReportFileName = gReportPath & "rptBookingPartyVsDispatch.rpt"
''            .SelectionFormula = X1
        Case "Party Booking Vs Dispatch Register"
            X1 = "SELECT tblAddLess_Book.vno, tblAddLess_Book.vdt, tblMastAccount.Acname, tblMastAccount_B.Acname,tblMastItem.Itname, tblvousub_Book.vSubBag, tblvousub_Book.VsubWt, tblAddLess_Gp.Vno, tblvousub_Book.VSubItCtrlNo, tblVouSub_Gp.Vdt, tblVouSub_Gp.vSubBag,tblVouSub_Gp.vSubWt,tblVouSub_Book.vSubrt,tblVouSub_Gp.vSubrt, '" & gUserName & "'" _
                 & " FROM tblAddLess_Book, tblMastAccount,tblMastAccount As tblMastAccount_B, tblvousub_Book, tblMastItem, tblAddLess AS tblAddLess_Gp, tblVousub AS tblVousub_Gp " _
                 & " WHERE tblAddLess_Book.Vtype = 'OP' and tblAddLess_Gp.AdAcCrCode = tblMastAccount.Accode and tblAddLess_Book.AdAcCrCode = tblMastAccount_B.Accode and tblAddLess_Book.Vno = tblVousub_Book.Vno and tblAddLess_Book.Vtype = tblVousub_Book.Vtype and " _
                 & " tblAddLess_Book.vyear = tblVouSub_Book.Vyear and tblVousub_Book.VsubItcode = tblMastItem.Itcode and " _
                 & " tblVousub_Book.VsubItctrlno =   tblVousub_Gp.VSubBkItCtrlNo and tblAddLess_Gp.Vno = tblVousub_Gp.Vno and tblAddLess_Gp.Vtype = tblVousub_Gp.Vtype and tblAddLess_Gp.vyear = tblVouSub_Gp.Vyear"
            X1 = X1 + " and tblAddLess_Book.vdt>=cdate('" & Format(dtpFrDt.Value, "yyyy,mm,dd") & "')  and tblAddLess_Book.Vdt<=Cdate('" & Format(dtpToDt.Value, "yyyy,mm,dd") & "')"
            gCn.BeginTrans
            gCn.Execute "delete from tmpGenTbl where username='" & gUserName & "'"
            gCn.Execute " Insert Into TmpGentbl (vno,Vdt,Nar1,Nar3,Nar2,Amt1,Amt2,Amt3,Amt4,VBILLDT,Amt5,Amt6,Amt7,Amt8,UserName)" & X1
            gCn.CommitTrans
            .Formulas(4) = "wFrToDt='From : '& '" & dtpFrDt.Value & "' & ' To ' & '" & dtpToDt.Value & "'"
            .SelectionFormula = "{TmpGentbl.UserName}='" & gUserName & "'"
             If Trim(txttemp(2)) <> "" Then .SelectionFormula = .SelectionFormula + " and {tmpGenTbl.Nar1} = '" & txttemp(2) & "'"
             If Trim(txttemp(5)) <> "" Then .SelectionFormula = .SelectionFormula + " and {tmpGenTbl.Nar2} = '" & txttemp(5) & "'"
             If Trim(txttemp(3)) <> "" Then .SelectionFormula = .SelectionFormula + " and {tmpGenTbl.Nar3} = '" & txttemp(3) & "'"
            .ReportFileName = gReportPath & "rptBookingPartyVsDispatchNew.rpt"
        '====== Godown Reports
''            .Formulas(4) = "wFrToDt='From : '& '" & dtpFrDt.Value & "' & ' To ' & '" & dtpToDt.Value & "'"
''            X1 = "{tblAddLess.VType}='GP' and {tblAddLess.VDt} in Date(" & Format(dtpFrDt.Value, "yyyy,mm,dd") & ") to Date(" & Format(dtpToDt.Value, "yyyy,mm,dd") & ")  " '
''            X1 = X1 + " and isnull({tblAddLess.AdLrDt})" '-- Delivery Dt Null
''            If Trim(txttemp(0)) <> "" Then '-Mill
''                X1 = X1 + " and {tblmastAccount_mill.AcName} = '" & txttemp(0) & "'"
''            End If
''            If Trim(txttemp(2)) <> "" Then '-Party
''                X1 = X1 + " and {tblmastAccount.AcName} = '" & txttemp(2) & "'"
''            End If
''            If Trim(txttemp(4)) <> "" Then '-Godown'
''                X1 = X1 + " and {tblMastNarration.Narration} = '" & txttemp(4) & "'"
''            End If
''            If Trim(txttemp(5)) <> "" Then '-Item
''                X1 = X1 + " and {tblmastItem.ItName} = '" & txttemp(5) & "'"
''            End If
''            .SelectionFormula = X1
''            .ReportFileName = gReportPath & "rptGatePassRptPending.rpt"
''        Case "Gate Pass Register" '-"Gate Pass Register"
''            .Formulas(4) = "wFrToDt='From : '& '" & dtpFrDt.Value & "' & ' To ' & '" & dtpToDt.Value & "'"
''            X1 = "{tblAddLess.VType}='GP' and {tblAddLess.VDt} in Date(" & Format(dtpFrDt.Value, "yyyy,mm,dd") & ") to Date(" & Format(dtpToDt.Value, "yyyy,mm,dd") & ")  " '
''            If Trim(txttemp(2)) <> "" Then '-Party
''                X1 = X1 + " and {tblmastAccount.AcName} = '" & txttemp(2) & "'"
''            End If
''            If Trim(txttemp(5)) <> "" Then '-Item
''                X1 = X1 + " and {tblmastItem.ItName} = '" & txttemp(5) & "'"
''            End If
''            If Trim(txttemp(0)) <> "" Then '-Mill
''                X1 = X1 + " and {tblmastAccount_Mill.AcName} = '" & txttemp(0) & "'"
''            End If
''            .SelectionFormula = X1
''            .ReportFileName = gReportPath & "rptGatePassRptRegi.rpt"
''        Case "Delivery Register" '-"Delivery Register"
''            .Formulas(4) = "wFrToDt='From : '& '" & dtpFrDt.Value & "' & ' To ' & '" & dtpToDt.Value & "'"
''            X1 = "{tblAddLess.VType}='GP' and isnull({tblAddLess.AdLrDt})=False and {tblAddLess.AdLrDt} in Date(" & Format(dtpFrDt.Value, "yyyy,mm,dd") & ") to Date(" & Format(dtpToDt.Value, "yyyy,mm,dd") & ")  " '
''            '-- Delivery Dt
''            If Trim(txttemp(2)) <> "" Then '-Party
''                X1 = X1 + " and {tblmastAccount.AcName} = '" & txttemp(2) & "'"
''            End If
''            If Trim(txttemp(4)) <> "" Then '-Godown'
''                X1 = X1 + " and {tblMastNarration.Narration} = '" & txttemp(4) & "'"
''            End If
''            If Trim(txttemp(5)) <> "" Then '-Item
''                X1 = X1 + " and {tblmastItem.ItName} = '" & txttemp(5) & "'"
''            End If
''            If Trim(txttemp(0)) <> "" Then '-Mill
''                X1 = X1 + " and {tblmastAccount_mill.AcName} = '" & txttemp(0) & "'"
''            End If
''            .SelectionFormula = X1
''            .ReportFileName = gReportPath & "rptGatePassRptDelivery.rpt"
        Case "Pending Gate Pass Report" '--"Pending Gate Pass Report"
            If gCYear >= 2025 Then
             X1 = "select tblGpSub.Vno,tblGpSub.Vdt,tblMastItem.Itname,tblMastAccount.AcName,tblMastAccount_Br.AcName,tblMastAccount_Ml.AcName,tblMastAccount_Ml.AcIsLocalBrkBySelf,tblMastNarration.Narration,tblMastItem.ItTicket," _
                  & " tblVouSub.VSubBag , tblVouSub.VSubWt,tblVouSub.VSubRt, tblGpsub.VSubCtrNo, tblGpsub.VSubBag, tblGpsub.VSubWt,'" & gUserName & "' from tblGpsub ,tblMastItem,tblAddLess,tblMastAccount, tblMastAccount tblMastAccount_Br, tblMastAccount tblMastAccount_Ml,tblVouSub,tblMastNarration " _
                  & " where isnull(tblGpSub.VDeliveryDt) and tblGpSub.VSubItCode = tblMastItem.ItCode and tblGpsub.Vno = tblAddLess.Vno and  tblGpsub.Vtype = tblAddLess.Vtype and tblGpsub.Vyear = tblAddLess.Vyear and " _
                  & " tblAddLess.AdAcCrCode = tblMastAccount.AcCode And tblAddLess.AdAcDrCode = tblMastAccount_Br.AcCode And tblAddLess.AdMillCode = tblMastAccount_Ml.AcCode " _
                  & " and tblAddLess.Vno = tblVouSub.Vno and  tblAddLess.Vtype = tblVouSub.Vtype and tblAddLess.Vyear = tblVouSub.Vyear and tblVouSub.VSubGodownCode = tblMastNarration.NarrCode and tblGpSub.Vdt>=Cdate('" & dtpFrDt.Value & "') and tblGpSub.Vdt<=Cdate('" & dtpToDt.Value & "')  "
             gCn.BeginTrans
             gCn.Execute "delete from tmpGenTbl2 where Username='" & gUserName & "'"
             gCn.Execute "insert into tmpGenTbl2 (Vno,Vdt,Nar1,Nar2,Nar3,Nar4,Amt6,Nar5,Nar6,Amt1,Amt2,Amt5,VACCODE,Amt3,Amt4,UserName) " & X1
             gCn.CommitTrans
            .Formulas(4) = "wFrToDt='From : '& '" & dtpFrDt.Value & "' & ' To ' & '" & dtpToDt.Value & "'"
             X1 = "{tmpGenTbl2.UserName}='" & gUserName & "' and {tmpGenTbl2.Amt6} <> 1 "
            If Trim(txttemp(2)) <> "" Then '-Party
                X1 = X1 + " and {tmpGenTbl2.Nar2} = '" & txttemp(2) & "'"
            End If
            If Trim(txttemp(5)) <> "" Then '-Party
                X1 = X1 + " and {tmpGenTbl2.Nar1} = '" & txttemp(5) & "'"
            End If
            If Trim(txttemp(3)) <> "" Then '-Party
                X1 = X1 + " and {tmpGenTbl2.Nar3} = '" & txttemp(3) & "'"
            End If
            If Trim(txttemp(0)) <> "" Then '-Party
                X1 = X1 + " and {tmpGenTbl2.Nar4} = '" & txttemp(0) & "'"
            End If
            If Trim(txttemp(4)) <> "" Then '-Party
                X1 = X1 + " and {tmpGenTbl2.Nar5} = '" & txttemp(4) & "'"
            End If
             .SelectionFormula = X1
             If ChkConsolidate.Value = 1 Then
                .ReportFileName = gReportPath & "rptGatePassRptPendingGod.rpt"
             Else
                .ReportFileName = gReportPath & "rptGatePassRptPending.rpt"
             End If
          Else
            .Formulas(4) = "wFrToDt='From : '& '" & dtpFrDt.Value & "' & ' To ' & '" & dtpToDt.Value & "'"
            X1 = "{tblAddLess.VType}='GP' and {tblAddLess.VDt} in Date(" & Format(dtpFrDt.Value, "yyyy,mm,dd") & ") to Date(" & Format(dtpToDt.Value, "yyyy,mm,dd") & ")  " '
            X1 = X1 + " and isnull({tblAddLess.AdLrDt})" '-- Delivery Dt Null
            If Trim(txttemp(0)) <> "" Then '-Mill
                X1 = X1 + " and {tblmastAccount_mill.AcName} = '" & txttemp(0) & "'"
            End If
            If Trim(txttemp(2)) <> "" Then '-Party
                X1 = X1 + " and {tblmastAccount.AcName} = '" & txttemp(2) & "'"
            End If
            If Trim(txttemp(4)) <> "" Then '-Godown'
                X1 = X1 + " and {tblMastNarration.Narration} = '" & txttemp(4) & "'"
            End If
            If Trim(txttemp(5)) <> "" Then '-Item
                X1 = X1 + " and {tblmastItem.ItName} = '" & txttemp(5) & "'"
            End If
            .SelectionFormula = X1
            .ReportFileName = gReportPath & "rptGatePassRptPendingOld.rpt"
         End If
        Case "Gate Pass Register" '-"Gate Pass Register"
            .Formulas(4) = "wFrToDt='From : '& '" & dtpFrDt.Value & "' & ' To ' & '" & dtpToDt.Value & "'"
            If OptShort.Value = True Then
                
                X1 = "select tblGpSub.Vno,tblGpSub.VSubCtrNo,tblGpSub.vdt,tblMastAccount.Acname ,tblMastAccount_Br.Acname,tblMastItem.Itname,tblMastAccount_Mill.Acname,tblMastNarration.Narration,tblGpSub.VSubBag,tblGpSub.VSubwt, tblvousub.VSUBTMP2," _
                     & " tblMastItem.ItStdrateper,tblMastItem.ItStdpack,tblvouSub.VSubRt ,tblMastItem.ItTicket,tblvouSub.VSUBTMP7,'" & gUserName & "' from tblGpSub,tbladdLess,tblmastaccount ,tblvouSub,tblmastaccount tblmastaccount_Br,tblMastItem,tblmastaccount tblmastaccount_Mill,tblMastNarration " _
                     & " Where tblGpSub.vtype = 'GP' and tblGpSub.Vno = tblAddLess.Vno  and tblGpSub.Vtype= tblAddLess.Vtype and tblGpSub.Vyear = tblAddLess.Vyear and " _
                     & " tblAddless.AdAcCrCode = tblmastaccount.accode and tblAddless.AdAcDrCode = tblmastaccount_Br.accode and tblAddless.AdMillCode = tblmastaccount_Mill.accode and " _
                     & " tblGpSub.Vno = tblvouSub.Vno  and tblGpSub.Vtype= tblvouSub.Vtype and tblGpSub.Vyear = tblvouSub.Vyear  and tblVousub.VsubItcode = tblMastItem.Itcode and tblGpSub.VSubGodownCode = tblMastNarration.NarrCode  "
                X1 = X1 + " and tblAddLess.vdt>=cdate('" & Format(dtpFrDt.Value, "yyyy,mm,dd") & "')  and tblAddLess.vdt<=Cdate('" & Format(dtpToDt.Value, "yyyy,mm,dd") & "')"
                
                If Trim(txttemp(2)) <> "" Then '-Party
                    X1 = X1 + " and tblmastAccount.AcName = '" & txttemp(2) & "'"
                End If
                If Trim(txttemp(0)) <> "" Then '-Mill
                    X1 = X1 + " and tblAddLess.AdMillCode = " & wMillCode & ""
                End If
                gCn.BeginTrans
                gCn.Execute "delete from tmpGenTbl2 where Username='" & gUserName & "'"
                gCn.Execute "insert into tmpGenTbl2 (Vno,VSRNO,Vdt,Nar1,Nar2,Nar3,Nar4,Nar5,Amt1,Amt2,Amt3,Amt4,Amt5,Amt8,Nar6,Nar7,UserName) " & X1
                gCn.CommitTrans
                X1 = "{tmpGenTbl2.UserName}='" & gUserName & "'"
                
                
                
                
                
'                X1 = "{tblAddLess.VType}='GP' and {tblAddLess.VDt} in Date(" & Format(dtpFrDt.Value, "yyyy,mm,dd") & ") to Date(" & Format(dtpToDt.Value, "yyyy,mm,dd") & ")  " '
                .SelectionFormula = X1
               .Formulas(5) = "wC= " & Val(txt(4)) & ""
               .Formulas(6) = "wB= " & Val(txt(5)) & ""
               .Formulas(7) = "wG= " & Val(txt(6)) & ""
                If Trim(txttemp(0)) <> "" Then '-Mill
                  .Formulas(8) = "wMillname='Mill Name : ' & '" & txttemp(0) & "'"
                End If
               .ReportFileName = gReportPath & "rptGatePassRptRegi_Spl2.rpt"
            Else
                X1 = "{tblAddLess.VType}='GP' and {tblAddLess.VDt} in Date(" & Format(dtpFrDt.Value, "yyyy,mm,dd") & ") to Date(" & Format(dtpToDt.Value, "yyyy,mm,dd") & ")  " '
                If Trim(txttemp(2)) <> "" Then '-Party
                    X1 = X1 + " and {tblmastAccount.AcName} = '" & txttemp(2) & "'"
                End If
                If Trim(txttemp(0)) <> "" Then '-Mill
                    X1 = X1 + " and {tblAddLess.AdMillCode} = " & wMillCode & ""
                End If
                If Trim(txttemp(3)) <> "" Then
                    X1 = X1 + " and {tblmastAccount_Brk.AcName} = '" & txttemp(3) & "'"
                End If
                If Trim(txttemp(5)) <> "" Then
                    X1 = X1 + " and {tblmastItem.ItName} = '" & txttemp(5) & "'"
                End If
                If Trim(txttemp(4)) <> "" Then
                    X1 = X1 + " and {tblMastNarration.Narration} = '" & txttemp(4) & "'"
                End If
                .SelectionFormula = X1
                .ReportFileName = gReportPath & "rptGatePassRptRegi.rpt"
            End If
        Case "Delivery Register" '-"Delivery Register"
        
             X = "Select tblAddLess.Vno,tblAddLess.Vdt,tblGpSub.VDeliveryDt ,tblMastAccount.Acname,tblMastAccount_Br.Acname,tblMastAccount_Ml.Acname,tblMastItem.Itname,tblMastItem.ItTicket,tblMastNarration.Narration,tblMastNarration_S.Narration," _
                 & " tblGpSub.VSubCtrNo , tblGpSub.VSubBag, tblGpSub.VSubWt,tblGpSub.VehicleNo,'" & gUserName & "' from tblGpSub,tblMastAccount,tblAddLess, tblMastAccount tblMastAccount_Br, tblMastAccount tblMastAccount_Ml,tblMastItem,tblMastNarration,tblMastNarration tblMastNarration_S " _
                 & " WHERE isdate(VDeliveryDt) and tblAddLess.Vno = tblGpSub.Vno and tblAddLess.Vtype = tblGpSub.Vtype and tblAddLess.Vyear = tblGpSub.Vyear and tblAddLess.AdAcCrCode = tblMastAccount.AcCode and " _
                 & " tblAddLess.AdAcDrCode = tblMastAccount_Br.AcCode and tblAddLess.AdMillCode = tblMastAccount_Ml.AcCode and tblGpSub.VSubItCode = tblMastItem.Itcode and tblGpSub.VSubGodownCode = tblMastNarration.NarrCode and tblGpSub.VSizer = tblMastNarration_S.NarrCode"
             
             X = X + " and tblGpSub.VDeliveryDt >= cdate('" & Format(dtpFrDt.Value, "yyyy,mm,dd") & "')  and tblGpSub.VDeliveryDt <= Cdate('" & Format(dtpToDt.Value, "yyyy,mm,dd") & "') "
             
             If Trim(txttemp(2)) <> "" Then X = X + " and tblMastAccount.Acname = '" & Trim(txttemp(2)) & "'"
             If Trim(txttemp(0)) <> "" Then X = X + " and tblMastAccount_Ml.Acname = '" & Trim(txttemp(0)) & "'"
             If Trim(txttemp(4)) <> "" Then X = X + " and tblMastNarration.Narration = '" & Trim(txttemp(4)) & "'"
             If Trim(txttemp(5)) <> "" Then X = X + " and tblMastItem.ItName= '" & Trim(txttemp(5)) & "'"
             
             gCn.BeginTrans
             gCn.Execute "delete from tmpGenTbl2 where Username='" & gUserName & "'"
             gCn.Execute "insert into tmpGenTbl2 (Vno,Vdt,Dt1,Nar1,Nar2,Nar3,Nar4,NAR6,Nar5,Nar7,VSRNO,Amt1,Amt2,Nar8,UserName) " & X
             gCn.CommitTrans
             X1 = "{tmpGenTbl2.UserName}='" & gUserName & "'"
             .Formulas(4) = "wFrToDt='From : '& '" & dtpFrDt.Value & "' & ' To ' & '" & dtpToDt.Value & "'"
             .SelectionFormula = X1
             .ReportFileName = gReportPath & "rptGatePassRptDelivery.rpt"


        Case "Godownwise Inward Register"  '- "Godownwise Inward Register"
            .Formulas(4) = "wFrToDt='From : '& '" & dtpFrDt.Value & "' & ' To ' & '" & dtpToDt.Value & "'"
            '-Vtype & Vdt
            X1 = "({tblVouSub.VType}='PI' or {tblVouSub.VType}='PY' or {tblVouSub.VType}='PT') and {tblVouSub.Vdt} in Date(" & Format(dtpFrDt.Value, "yyyy,mm,dd") & ") to Date(" & Format(dtpToDt.Value, "yyyy,mm,dd") & ")  " '
            If Trim(txttemp(5)) <> "" Then '-Item
                X1 = X1 + " and {tblmastItem.ItName} = '" & txttemp(5) & "'"
            End If
            If Trim(txttemp(4)) <> "" Then '-Godown'
                X1 = X1 + " and {tblMastNarration.Narration} = '" & txttemp(4) & "'"
            End If
            If Trim(txttemp(0)) <> "" Then X1 = X1 + " and {tblMastAccount.Acname} = '" & Trim(txttemp(0)) & "'"
            .SelectionFormula = X1
            .ReportFileName = gReportPath & "rptGodownRptInwardReg.rpt"
        Case "Countwise Goods Delivery Report", "Count & Partywise Goods Delivery Report", "Count & Sizerwise Goods Delivery Report" '-"Countwise Goods Delivery Report"
            .Formulas(4) = "wFrToDt='From : '& '" & dtpFrDt.Value & "' & ' To ' & '" & dtpToDt.Value & "'"
             '-- Deliveried
            X1 = "{tblAddLess.VType}='GP' and isnull({tblAddLess.AdLrDt})=False and {tblAddLess.Vdt} in Date(" & Format(dtpFrDt.Value, "yyyy,mm,dd") & ") to Date(" & Format(dtpToDt.Value, "yyyy,mm,dd") & ")  " '
            If Trim(txttemp(2)) <> "" Then '-Party
                X1 = X1 + " and {tblmastAccount.AcName} = '" & txttemp(2) & "'"
            End If
            If Trim(txttemp(6)) <> "" Then '-Sizer
                X1 = X1 + " and {tblAddLess.AdTmp5} = '" & txttemp(6) & "'"
            End If
            If Trim(txttemp(5)) <> "" Then '-Item
                X1 = X1 + " and {tblMastItem.ItName} = '" & txttemp(5) & "'"
            End If
            .SelectionFormula = X1
            If gSelectedMenu = "Countwise Goods Delivery Report" Then '-count wise
                .ReportFileName = gReportPath & "rptGPRptDelCountWise.rpt"
            ElseIf gSelectedMenu = "Count & Partywise Goods Delivery Report" Then '-Party Wse
                .ReportFileName = gReportPath & "rptGPRptDelPartywise.rpt"
            Else '-Sizer Wise
                .ReportFileName = gReportPath & "rptGPRptDelSizerwise.rpt"
            End If
        Case "Party and Bag No wise Goods Delivery Report" '--"Party and Bag No wise Goods Delivery Report"
            .Formulas(4) = "wFrToDt='From : '& '" & dtpFrDt.Value & "' & ' To ' & '" & dtpToDt.Value & "'"
             '-- Deliveried
            X1 = "{tblAddLess.VType}='GP' and isnull({tblAddLess.AdLrDt})=False and {tblAddLess.Vdt} in Date(" & Format(dtpFrDt.Value, "yyyy,mm,dd") & ") to Date(" & Format(dtpToDt.Value, "yyyy,mm,dd") & ")  " '
           '-- TblBag
           X1 = X1 + " and isnull({tblBags.InwType})=True and {tblBags.InwNo}=0"
            If Trim(txttemp(2)) <> "" Then '-Party
                X1 = X1 + " and {tblmastAccount.AcName} = '" & txttemp(2) & "'"
            End If
            If Trim(txttemp(5)) <> "" Then '-Item
                X1 = X1 + " and {tblMastItem.ItName} = '" & txttemp(5) & "'"
            End If
            .SelectionFormula = X1
            .ReportFileName = gReportPath & "rptGPRptDelPartyBagWise.rpt"
        '======= Other Register
        Case "Sale In Transit Register" '-- "Sale In Transit Register"
            .Formulas(4) = "wFrToDt='From : '& '" & dtpFrDt.Value & "' & ' To ' & '" & dtpToDt.Value & "'"
            '-- Purchase - PT ,Sale - ST
            X1 = "{tblAddLess.VType}='ST' and {tblAddLess_In.VType}='PT' and {tblAddLess_In.AdGpdt} in Date(" & Format(dtpFrDt.Value, "yyyy,mm,dd") & ") to Date(" & Format(dtpToDt.Value, "yyyy,mm,dd") & ")  " '
            If Trim(txttemp(0)) <> "" Then '-Mill
                X1 = X1 + " and {tblmastAccount_InMill.AcName} = '" & txttemp(0) & "'"
                .Formulas(5) = "Str='Mill : '& '" & txttemp(0) & "'"
            End If
           .SelectionFormula = X1
           If ChkConsolidate.Value = 1 Then
              .ReportFileName = gReportPath & "rptSITRegNew.rpt"
           Else
              .ReportFileName = gReportPath & "rptSITReg.rpt"
           End If
        Case "Trade Purchase - Sale Register"
            .Formulas(4) = "wFrToDt='From : '& '" & dtpFrDt.Value & "' & ' To ' & '" & dtpToDt.Value & "'"
            '-- Purchase - PY ,Sale - SY
            X1 = "{tblAddLess.VType}='SY' and {tblAddLess_In.VType}='PY' and {tblAddLess.Vdt} in Date(" & Format(dtpFrDt.Value, "yyyy,mm,dd") & ") to Date(" & Format(dtpToDt.Value, "yyyy,mm,dd") & ")  " '
            If Trim(txttemp(0)) <> "" Then '-Mill
                X1 = X1 + " and {tblmastAccount_InMill.AcName} = '" & txttemp(0) & "'"
                .Formulas(5) = "Str='Mill : '& '" & txttemp(0) & "'"
            End If
            If Trim(txttemp(5)) <> "" Then '-Item
                X1 = X1 + " and {tblMastItem_In.ItName} = '" & txttemp(5) & "'"
            End If
            If Trim(txttemp(2)) <> "" Then X1 = X1 + " and {tblmastAccount.AcName} = '" & txttemp(2) & "'"
           .SelectionFormula = X1
           If chkMonthwise.Value = 1 Then
              .ReportFileName = gReportPath & "rptTradePurSaleRegMonthwise.rpt"
           Else
              .ReportFileName = gReportPath & "rptTradePurSaleReg.rpt"
           End If
           
        Case "Late Pay Debit Note Register", "Late Pay Debit Note Register (Mill Bill)" '"Late Pay Debit Note Register"
                If gSelectedMenu = "Late Pay Debit Note Register" Then
                   GProcLatePayDbNtReg Format(dtpFrDt.Value, "dd/MM/yyyy"), Format(dtpToDt.Value, "dd/MM/yyyy"), ChkConsolidate, txttemp(3), txttemp(2)
                Else '-Mill Bill
                   GProcLatePayDbNtReg Format(dtpFrDt.Value, "dd/MM/yyyy"), Format(dtpToDt.Value, "dd/MM/yyyy"), ChkConsolidate, txttemp(3), txttemp(2), True
                End If
                .Formulas(4) = "wFrToDt = 'From : ' & '" & Format(dtpFrDt.Value, "dd/MM/yyyy") & "' & '  To ' & '" & Format(dtpToDt.Value, "dd/MM/yyyy") & "'"
                If ChkConsolidate = 1 Then .Formulas(6) = "wAllFirmStr= 'All Firm'" '- wAllFirmStr
                .DataFiles(0) = App.Path & "\HIData\setup.mdb"
                .SelectionFormula = "{tmpgenTbl.Username}='" & gUserName & "'"
                gReportPath = App.Path & "\HiReports\"
                .ReportFileName = gReportPath & "rptLatePayDbNtReg.rpt"
        '=== "Missing Transaction Report
        Case "Missing Transaction Report - Booking", "Missing Transaction Report - Vouchers", "Missing Transaction Report - Sales", "Missing Transaction Report - Gate Pass", "Missing Transaction Report - Purchase"
                Select Case gSelectedMenu '-- Vtype
                    Case "Missing Transaction Report - Booking"
                        X = "BK"
                    Case "Missing Transaction Report - Vouchers"
                        X = "VU"
                    Case "Missing Transaction Report - Sales"
                        X = "SA"
                    Case "Missing Transaction Report - Gate Pass"
                        X = "GP"
                    Case "Missing Transaction Report - Purchase"
                        X = "PU"
                End Select
                '-- Prepare Misssing Transaction Report
                GProcPrepareMisssingTransRpt Format(dtpFrDt.Value, "dd/MM/yyyy"), Format(dtpToDt.Value, "dd/MM/yyyy"), X
                '--Set Locate
                If gBackEndDB = gBackEndAccess Then
                    .DataFiles(0) = gDbLocation
                Else '-- Oracle
                    .Connect = "DSN=" + gODBCDSN + " ;UID=Scott;PWD=" + gODBCPwd + ";DBQ=" + gODBCHostStr
                End If
                .Formulas(3) = "RptName = 'Missing Transaction Report'"
                .Formulas(4) = "wFrToDt = 'From : ' & '" & Format(dtpFrDt.Value, "dd/MM/yyyy") & "' & '  To ' & '" & Format(dtpToDt.Value, "dd/MM/yyyy") & "'"
                .SelectionFormula = "{tmpgenTbl.Username}='" & gUserName & "'"
                .ReportFileName = gReportPath & "rptMissingTransRpt.rpt"
        '===== Outstanding Report
        Case "Outstanding Broker with Partywise", "Outstanding Brokerwise", "Mill Bill Outstanding Broker with Partywise" '--"Outstanding Broker with Partywise", "Outstanding Brokerwise",Mill Bill Outstanding Broker with Partywise
               Dim Fdays As Integer, Tdays As Integer
               Fdays = CDbl(txt(2))
               Tdays = CDbl(txt(3))
        
               GProcCollectCommonOutstanding dtpFrDt, dtpToDt, txttemp(0), txttemp(3), txttemp(2), IIf(ChkConsolidate = 1, False, True), ChkUnAdjustRec
               CrystalReport1.DataFiles(0) = App.Path & "\HIData\setup.mdb"
                If ChkStkValTax.Value = 1 Then
                  .SelectionFormula = "{tmptbloutstanding.Username}='" & gUserName & "'   and {tmptbloutstanding.VcompSNm} <> 'Mata'  and {tmptbloutstanding.VcompSNm} <> 'Utc2'"
                Else
                  .SelectionFormula = "{tmptbloutstanding.Username}='" & gUserName & "'  and {tmptbloutstanding.OutIsShowOuts} = 0 "
                End If
                If Chkonly = 1 Then
                   .SelectionFormula = "{tmptbloutstanding.Username}='" & gUserName & "' and {tmptbloutstanding.outduedays} = 2 and {tmptbloutstanding.OutIsShowOuts} = 0 "
                Else
                    .SelectionFormula = "{tmptbloutstanding.Username}='" & gUserName & "' and {tmptbloutstanding.OutIsShowOuts} = 0 "
                End If
                If Fdays + Tdays > 0 Then
                     .SelectionFormula = .SelectionFormula + " and {@mDueDays} >  " & Fdays & " and {@mDueDays} <  " & Tdays & " "
                End If

                
               gReportPath = App.Path & "\HiReports\"
               If gSelectedMenu = "Outstanding Brokerwise" Then '"Outstanding Brokerwise"
                    .ReportFileName = gReportPath & "rptOutstBrokerwise.rpt"
               Else
                    If ChkUnAdjustRec = 1 Then
                        .ReportFileName = gReportPath & "rptOutstBrkPartyWithUnadjustRec.rpt"
                    Else
                        .ReportFileName = gReportPath & "rptOutstBrokerPartywise.rpt"
                    End If
               End If
               '-Detail
               'CrystalReport1.Formulas(4) = "wFrToDt='From : '& '" & dtpFrDt.Value & "' & ' To ' & '" & dtpToDt.Value & "'"
               '.ReportFileName = gReportPath & "rptOutstBrokerPartyFirm.rpt"
        Case "Purchase Outstanding Partywise" '-"Purchase Outstanding Partywise"
            GProcPurchaseOutstanding dtpFrDt, dtpToDt, txttemp(3), txttemp(2), txttemp(0), IIf(ChkConsolidate = 1, False, True)
            CrystalReport1.DataFiles(0) = App.Path & "\HIData\setup.mdb"
            .SelectionFormula = "{tmptbloutstanding.Username}='" & gUserName & "'"
            gReportPath = App.Path & "\HiReports\"
            .ReportFileName = gReportPath & "rptOutstPurchPartywise.rpt"
        Case "Outstanding Late Payment Interest Report", "Outstanding Late Payment Interest Report (Mill Bill)" '-"Outstanding Late Payment Interest Report"
            If gSelectedMenu = "Outstanding Late Payment Interest Report" Then
                GProcOutstandingSALatePayIntOS Format(dtpFrDt.Value, "dd/MM/yyyy"), Format(dtpToDt.Value, "dd/MM/yyyy"), ChkConsolidate, txttemp(3), txttemp(2)
            Else '-Mill Bill
                GProcOutstandingSALatePayIntOS Format(dtpFrDt.Value, "dd/MM/yyyy"), Format(dtpToDt.Value, "dd/MM/yyyy"), ChkConsolidate, txttemp(3), txttemp(2), True
            End If
            .Formulas(4) = "wFrToDt = 'From : ' & '" & Format(dtpFrDt.Value, "dd/MM/yyyy") & "' & '  To ' & '" & Format(dtpToDt.Value, "dd/MM/yyyy") & "'"
            If ChkConsolidate = 1 Then .Formulas(6) = "wAllFirmStr= 'All Firm'" '- wAllFirmStr
            .DataFiles(0) = App.Path & "\HIData\setup.mdb"
            .SelectionFormula = "{tmpgenTbl.Username}='" & gUserName & "'"
            gReportPath = App.Path & "\HiReports\"
            .ReportFileName = gReportPath & "rptOutstSALatePayIntOS.rpt"
        Case "Unadjusted Recepit" '--Unadjusted Recepit
            X1 = " ({QryRegUnadjustReceipt.Vtype}= 'BR' or {QryRegUnadjustReceipt.Vtype}='CR') and   {QryRegUnadjustReceipt.UnadjustAmt}>0 and {QryRegUnadjustReceipt.VDt} in Date(" & Format(dtpFrDt.Value, "yyyy,mm,dd") & ") to Date(" & Format(dtpToDt.Value, "yyyy,mm,dd") & ")  " '
            .Formulas(8) = "wFrToDt = 'From : ' & '" & Format(dtpFrDt.Value, "dd/MM/yyyy") & "' & '  To ' & '" & Format(dtpToDt.Value, "dd/MM/yyyy") & "'"
            .Formulas(9) = "wRptName='Unadjusted Receipt Register'"
            If txttemp(2) <> "" Then '--Party
                X1 = X1 + " and {QryRegUnadjustReceipt.PARTY}='" & txttemp(2) & "'"
                .SelectionFormula = X1
                .Formulas(6) = "wPartyName='Party Name : ' & '" & txttemp(2) & "'"
                .ReportFileName = gReportPath & "rptUnAdjustReceiptPartyRegi.rpt"
            Else
                .SelectionFormula = X1
                .ReportFileName = gReportPath & "rptUnAdjustReceiptRegi.rpt"
            End If
        Case "Unadjusted Payment" '--Unadjusted Payment
            X1 = " ({QryRegUnadjustReceipt.Vtype}= 'BP' or {QryRegUnadjustReceipt.Vtype}='CP') and   {QryRegUnadjustReceipt.UnadjustAmt}>0 and {QryRegUnadjustReceipt.VDt} in Date(" & Format(dtpFrDt.Value, "yyyy,mm,dd") & ") to Date(" & Format(dtpToDt.Value, "yyyy,mm,dd") & ")  " '
            .Formulas(8) = "wFrToDt = 'From : ' & '" & Format(dtpFrDt.Value, "dd/MM/yyyy") & "' & '  To ' & '" & Format(dtpToDt.Value, "dd/MM/yyyy") & "'"
            .Formulas(9) = "wRptName='Unadjusted Payment Register'"
            If txttemp(2) <> "" Then '--Party
                X1 = X1 + " and {QryRegUnadjustReceipt.PARTY}='" & txttemp(2) & "'"
                .SelectionFormula = X1
                .Formulas(6) = "wPartyName='Party Name : ' & '" & txttemp(2) & "'"
                .ReportFileName = gReportPath & "rptUnAdjustReceiptPartyRegi.rpt"
            Else
                .SelectionFormula = X1
                .ReportFileName = gReportPath & "rptUnAdjustReceiptRegi.rpt"
            End If
       Case "Pending GST Amt In L.P. Int "
            GProcCollectCommonOutstanding dtpFrDt, dtpToDt, txttemp(0), txttemp(3), txttemp(2), IIf(ChkConsolidate = 1, False, True), ChkUnAdjustRec, True
            CrystalReport1.DataFiles(0) = App.Path & "\HIData\setup.mdb"
            CrystalReport1.Formulas(4) = "wFrToDt = 'From : ' & '" & Format(dtpFrDt.Value, "dd/MM/yyyy") & "' & '  To ' & '" & Format(dtpToDt.Value, "dd/MM/yyyy") & "'"
            gReportPath = App.Path & "\HiReports\"
            .SelectionFormula = "{tmptbloutstanding.Username}='" & gUserName & "'"
            .ReportFileName = gReportPath & "rptOutstGST.rpt"
            
        '===== Outstanding Reports End
        '=== Stock
        Case "Item Ledger" '--Item Ledger
            GprocPrepareStkItemLedger Format(dtpFrDt.Value, "dd/MM/yyyy"), Format(dtpToDt.Value, "dd/MM/yyyy"), txttemp(5), txttemp(0)
            .Formulas(4) = "wFrToDt = 'From : ' & '" & Format(dtpFrDt.Value, "dd/MM/yyyy") & "' & '  To ' & '" & Format(dtpToDt.Value, "dd/MM/yyyy") & "'"
           ' If txttemp(0) <> "" Then X1 = "Mill : " & txttemp(0) '-Mill
           ' If txttemp(5) <> "" Then X1 = X1 + " Item : " & txttemp(5) '-Item
           ' .Formulas(5) = "Str='" & X1 & "'"
            .SelectionFormula = "{tmpgenTbl.Username}='" & gUserName & "'"
            .ReportFileName = gReportPath & "rptStkItemLedger.rpt"
       Case "Daily Stock To Mill" '=== "Daily Stock To Mill"
            If txttemp(0) = "" Then wMillCode = 0 '-Mill
             GprocPrepareDailyStkToMill Format(dtpFrDt.Value, "dd/MM/yyyy"), Format(dtpToDt.Value, "dd/MM/yyyy"), wMillCode
            .Formulas(4) = "wFrToDt = ' On Date : ' & '" & Format(dtpToDt.Value, "dd/MM/yyyy") & "'"
             'X1 = "{tmpgenTbl.Username}='" & gUserName & "' and {tmpGenTbl.VSRNO}=3"  '
             X1 = "{tmpgenTbl.Username}='" & gUserName & "'"  '
            .SelectionFormula = X1
            .Formulas(5) = "CcTo ='" & txt(1) & "'" '-CCTo
            .ReportFileName = gReportPath & "rptStkDailyStkToMill_1.rpt"
            .Action = 1
             X1 = "{tmpgenTbl.Username}='" & gUserName & "'"  '
            .SelectionFormula = X1
            .Formulas(5) = ""
            .ReportFileName = gReportPath & "rptStkDailyStkToMill_2.rpt"
        Case "Trade & SIT Stock Valuation Register", "Trade & SIT Stock Valuation Summary" '-"Trade & SIT Stock Valuation Register"
            '=== Prepare Outward (in tmpgentbl )
            If gBackEndDB = gBackEndAccess Then
                X1 = "sum(iif(tblOut.vtype='RY',-tblOut.vsubbag,tblOut.vsubBag)),sum(iif(tblOut.vtype='RY',-tblOut.vsubWt,tblOut.vsubWt)) "
                DtCondStr = " tblOut.vdt<=cdate('" & Format(dtpToDt.Value, "dd/MM/yyyy") & "')"
            Else '-oracle
                X1 = "sum(decode(tblOut.vtype,'RY',-tblOut.vsubbag,tblOut.vsubBag)),sum(decode(tblOut.vtype,'RY',-tblOut.vsubWt,tblOut.vsubWt)) "
                DtCondStr = " tblOut.vdt<=to_date('" & Format(dtpToDt.Value, "dd/MM/yyyy") & "','dd/MM/yyyy')"
            End If
            gCn.Execute "delete from tmpGenTbl where username='" & gUserName & "'"
            '--- Insert Without GatePass
            'tblin.vtype in ('PT','PY')
           gCn.Execute "insert into tmpgentbl (Username,vno,amt1,amt2) " _
            & " select '" & gUserName & "',tblIn.vsubItCtrlNo,0,0 from tblvousub tblIn where tblin.vtype in ('PY','PT') and VsubItCtrlno not in(Select VsubPitctrlNo From tblvousub where vtype='GP' and VSUBGPSLTYPE in (0,3,4) and " _
            & " vdt<=" & IIf(gBackEndDB = gBackEndAccess, " cdate('" & Format(dtpToDt.Value, "dd/MM/yyyy") & "')", "to_Date('" & Format(dtpToDt.Value, "dd/MM/yyyy") & "' ,'dd/MM/yyyy')") & ")" _
            & " and vdt<=" & IIf(gBackEndDB = gBackEndAccess, " cdate('" & Format(dtpToDt.Value, "dd/MM/yyyy") & "')", "to_Date('" & Format(dtpToDt.Value, "dd/MM/yyyy") & "' ,'dd/MM/yyyy')")

             '-- Insert Balance Inward' SY ST
            gCn.Execute "insert into tmpgentbl (Username,vno,amt1,amt2)  select '" & gUserName & "',tblIn.vsubItCtrlNo, " & X1 & " from tblvousub tblOut,tblvousub tblGP,tblvousub tblIn where " _
            & " TblIn.VsubItCtrlNo=tblGp.VsubPitCtrlNo and tblGp.VsubItCtrlno=tblOut.VSubPitctrlNo and tblGp.vtype in ('GP','RY') and tblOut.Vtype in('SY','ST') and " & DtCondStr & " group by tblIn.vsubItCtrlNo"

            '--Pur Rtn
             gCn.Execute "insert into tmpgentbl (Username,vno,amt1,amt2,Amt3)  select '" & gUserName & "',tblIn.vsubItCtrlNo, " & X1 & ",tblIn.VsubItCtrlNo from tblvousub tblOut,tblvousub tblIn where " _
            & " TblIn.VsubItCtrlNo=tblOut.VsubPItCtrlNo and tblOut.vtype='VY' and " & DtCondStr & " group by tblIn.vsubItCtrlNo,tblIn.VsubPItCtrlNo"

            '--Sale Rtn
             gCn.Execute "insert into tmpgentbl (Username,vno,amt1,amt2,Amt3)  select '" & gUserName & "',tblIn.vsubItCtrlNo, " & X1 & ",tblIn.VsubItCtrlNo from tblvousub tblOut,tblvousub tblIn where " _
            & " TblIn.VsubItCtrlNo=tblOut.VsubPItCtrlNo and tblOut.vtype='RY' and " & DtCondStr & " group by tblIn.vsubItCtrlNo,tblIn.VsubPItCtrlNo"



            gCn.Execute "update tmpGenTbl set Nar1 = 'A' where username='" & gUserName & "'"
            
            gCn.Execute "insert into tmpgentbl (Nar1,Username,vno,amt1,amt2)  select 'B','" & gUserName & "',vno,sum(amt1),sum(amt2) from tmpgentbl where username='" & gUserName & "' and Nar1 = 'A' group by nar1,vno"
            gCn.Execute "update tmpGenTbl set amt1=0 where amt1 is null and username='" & gUserName & "'"
            gCn.Execute "update tmpGenTbl set amt2=0 where amt2 is null and username='" & gUserName & "'"

            gCn.Execute "Delete from tmpGenTbl where username='" & gUserName & "' and nar1 =  'A' "

            
            .Formulas(4) = "wFrToDt = 'On Date : ' & '" & Format(dtpToDt.Value, "dd/MM/yyyy") & "'"
            .Formulas(5) = "IsTaxInVal='" & ChkStkValTax & "'"
            '== Inward '
            X1 = " {tmpGenTbl.Username}='" & gUserName & "' and {tblmastItem.ItOtherItem} =  0 and {@BalanceBag}<>0 and ({TBLVOUSUB.VTYPE}='PY' or {TBLVOUSUB.VTYPE}='PT') and {tblAddLess.Vdt} <= Date(" & Format(dtpToDt.Value, "yyyy,mm,dd") & ") " '
            '- Mill
            If Trim(txttemp(0)) <> "" Then X1 = X1 + " and {tblmastAccount.AcName} = '" & txttemp(0) & "'"
            '- Item
            If Trim(txttemp(5)) <> "" Then X1 = X1 + " and {tblmastItem.ItName} = '" & txttemp(5) & "'"
            .SelectionFormula = X1
            If gSelectedMenu = "Trade & SIT Stock Valuation Register" Then ' Register
                .ReportFileName = gReportPath & "RptStkRptTradeSITStkVal.rpt"
            Else '-summary
                .ReportFileName = gReportPath & "RptStkRptTradeSITStkValSumm.rpt"
            End If
        Case "TDS Report"
              gCn.BeginTrans
              gCn.Execute "delete from tmpRptTbl where username='" & gUserName & "'"
              .Formulas(4) = "wFrToDt='From : '& '" & dtpFrDt.Value & "' & ' To ' & '" & dtpToDt.Value & "'"
              .Formulas(5) = "wcpan= '" & gCPAN & "'"
              .Formulas(6) = "wctan= '" & gCTAN & "'"
              If chkMonthwise.Value = 1 Then
                    X1 = "SELECT tblMastAccount.acname,tblMastAccount.AcPAN, tblMastAccount.AcTDSNature,tblMastAccountP.Acname,vno, vdt, AdInvno, AdGpdt, AdDbNtNo, AdDbNtRt, AdDbNtAmt,AdSubAmt,'" & gUserName & "' From tblAddLess, tblMastAccount,tblMastAccount tblMastAccountP" _
                          & " WHERE tblAddLess.AdAcCrCode = tblMastAccount.AcCode and tblAddLess.AdAcDrCode = tblMastAccountP.AcCode and  tblAddLess.Vtype = 'PO' and tblAddLess.AdDbNtAmt > 0 and " _
                          & "tblAddLess.vdt>=cdate('" & Format(dtpFrDt.Value, "yyyy,mm,dd") & "')  and tblAddLess.vdt<=Cdate('" & Format(dtpToDt.Value, "yyyy,mm,dd") & "')"
                          gCn.Execute " Insert Into tmpRptTbl (Nar1,Nar2,Nar6,Nar4,Vno,vdt,VBILLNO,VBILLDT,Nar3,Amt1,Amt2,Amt3,UserName)" & X1
              Else
                    X1 = "SELECT acname,AcPAN, tblAddLess.vno, tblAddLess.vdt, AdInvno, AdGpdt, AdDbNtNo, AdOtherAdd, AdOtherLess,AdSubAmt,'" & gUserName & "' From tblAddLess,tblVousub, tblMastAccount" _
                         & " WHERE tblAddLess.AdAcCrCode = tblMastAccount.AcCode and  tblAddLess.Vtype in ('PY') and tblAddLess.Vno = tblVouSub.Vno and" _
                         & " tblAddless.Vtype = tblVouSub.Vtype and tblAddless.Vyear = tblVousub.Vyear and tblAddLess.AdOtherLess > 0 and tblVousub.vsubctrno = 1 and " _
                         & "tblAddLess.vdt>=cdate('" & Format(dtpFrDt.Value, "yyyy,mm,dd") & "')  and tblAddLess.vdt<=Cdate('" & Format(dtpToDt.Value, "yyyy,mm,dd") & "')"
                         gCn.Execute " Insert Into tmpRptTbl (Nar1,Nar2,Vno,vdt,VBILLNO,VBILLDT,Nar3,Amt1,Amt2,Amt3,UserName)" & X1
              End If

               DoEvents
               gCn.CommitTrans
               X1 = "{tmpRptTbl.UserName}='" & gUserName & "'"
               If Trim(txttemp(2)) <> "" Then X1 = X1 + " and {tmpRptTbl.Nar1} = '" & txttemp(2) & "'"
               .SelectionFormula = X1
               If OptDetail.Value = True Then
                   .ReportFileName = gReportPath & "rptTSDReg.rpt"
               Else
                   .ReportFileName = gReportPath & "rptTSDRegSummery.rpt"
               End If
        
        End Select
    .Action = 1
End With

If gSelectedMenu = "Outstanding Broker with Partywise" Or gSelectedMenu = "Outstanding Brokerwise" Or gSelectedMenu = "Mill Bill Outstanding Broker with Partywise" Or gSelectedMenu = "Pending GST Amt In L.P. Int " Then
       Dim mTempCn2 As Connection, mCurrFirmDbNm As String
       Dim RstCompMast As Recordset
       Dim RsP As Recordset

       mCurrFirmDbNm = gDbNm
       GProcRstOpen RstCompMast, "Select * from tblmastcompany where CIsHide = 0 ", "R", gSetupCn
       RstCompMast.MoveFirst
       Do While Not RstCompMast.EOF
          gDbNm = RstCompMast!CCode + Right(mCurrFirmDbNm, 4)
          If Dir(gDataPath & gDbNm & ".mdb") <> "" Then
'          If Dir(gDataPath & RstCompMast!CompCode + RstCompMast!CompYear & ".mdb") <> "" Then
               If gBackEndDB = gBackEndORA Then
                   GProcConnectionOpenORA mTempCn2
               Else
                   GProcConnectionOpen mTempCn2
               End If
    '----------For Updating o/s Balance for ToDate
               '================Outstanding Vs Receipt
               mTempCn2.BeginTrans
               mTempCn2.Execute "Update tblOutStanding set OutRecAMT = 0, OutRecDbAMT = 0"
               If (GProcRstOpen(RsP, "select * from tblRecVsSale where format(RecVDt,'yyyy/MM/dd') <= '" & Format(gCYEDate, "yyyy/MM/dd") & "'", "O", mTempCn2) > 0) Then
                   RsP.MoveFirst
                   Do While Not RsP.EOF
                      If RsP!InvTp = "SM" Then
                          mTempCn2.Execute "Update tblOutStanding set OutRecAMT = OutRecAmt+" & RsP!AdjAmt & "+" & RsP!AdjDbNtAmt & " where vtype='" & RsP!InvTp & "' and vno=" & RsP!InvNo & " and VYear=" & RsP!InvYear
                      Else
                          mTempCn2.Execute "Update tblOutStanding set OutRecAMT = OutRecAmt+" & RsP!AdjAmt & "+" & RsP!Discount & " , OutRecDbAMT = OutRecDbAmt+" & RsP!AdjDbNtAmt & " where vtype='" & RsP!InvTp & "' and vno=" & RsP!InvNo & " and VYear=" & RsP!InvYear
                      End If
                       RsP.MoveNext
                   Loop
                   mTempCn2.CommitTrans
                   RsP.Close
               End If
          End If
          RstCompMast.MoveNext
       Loop
      If RstCompMast.State = adStateOpen Then RstCompMast.Close
      gDbNm = mCurrFirmDbNm
      GProcConnectionOpen gCn
End If



End Sub
