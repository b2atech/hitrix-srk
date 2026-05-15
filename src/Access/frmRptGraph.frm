VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form frmRptGraph 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Report"
   ClientHeight    =   5160
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9750
   Icon            =   "frmRptGraph.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5160
   ScaleWidth      =   9750
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame5 
      BackColor       =   &H80000018&
      Caption         =   "&Amount /Bag In Format"
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
      Height          =   1170
      Left            =   6150
      TabIndex        =   27
      Top             =   3375
      Width           =   3210
      Begin VB.OptionButton OptNumFormat 
         BackColor       =   &H80000018&
         Caption         =   "Crore"
         Height          =   360
         Index           =   4
         Left            =   1035
         TabIndex        =   32
         Top             =   645
         Width           =   855
      End
      Begin VB.OptionButton OptNumFormat 
         BackColor       =   &H80000018&
         Caption         =   "Lac"
         Height          =   360
         Index           =   3
         Left            =   120
         TabIndex        =   31
         Top             =   645
         Width           =   855
      End
      Begin VB.OptionButton OptNumFormat 
         BackColor       =   &H80000018&
         Caption         =   "Thousand"
         Height          =   360
         Index           =   2
         Left            =   2085
         TabIndex        =   30
         Top             =   300
         Width           =   1080
      End
      Begin VB.OptionButton OptNumFormat 
         BackColor       =   &H80000018&
         Caption         =   "Hundred"
         Height          =   360
         Index           =   1
         Left            =   1035
         TabIndex        =   29
         Top             =   300
         Width           =   1080
      End
      Begin VB.OptionButton OptNumFormat 
         BackColor       =   &H80000018&
         Caption         =   "No"
         Height          =   360
         Index           =   0
         Left            =   120
         TabIndex        =   28
         Top             =   300
         Value           =   -1  'True
         Width           =   855
      End
   End
   Begin VB.CheckBox ChkPuVsSA 
      BackColor       =   &H80000018&
      Caption         =   "Is Purchase &Vs Sales ?"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   330
      Left            =   3870
      TabIndex        =   11
      Top             =   2145
      Width           =   2085
   End
   Begin VB.CheckBox ChkConsolidate 
      BackColor       =   &H80000018&
      Caption         =   "Is &Consolidate Report ?"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   330
      Left            =   1635
      TabIndex        =   10
      Top             =   2145
      Width           =   2085
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H80000018&
      Caption         =   "Sales / Purchase &Report On"
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
      Height          =   750
      Left            =   6150
      TabIndex        =   24
      Top             =   2505
      Width           =   3210
      Begin VB.OptionButton OptAmt 
         BackColor       =   &H80000018&
         Caption         =   "Amount"
         Height          =   360
         Left            =   495
         TabIndex        =   25
         Top             =   240
         Value           =   -1  'True
         Width           =   1080
      End
      Begin VB.OptionButton OptBag 
         BackColor       =   &H80000018&
         Caption         =   "Bags"
         Height          =   360
         Left            =   1950
         TabIndex        =   26
         Top             =   240
         Width           =   930
      End
   End
   Begin VB.TextBox txtTemp 
      Height          =   375
      Index           =   0
      Left            =   1620
      TabIndex        =   5
      Top             =   720
      Width           =   4335
   End
   Begin VB.TextBox txtTemp 
      Height          =   375
      Index           =   2
      Left            =   1620
      TabIndex        =   9
      Top             =   1605
      Width           =   4335
   End
   Begin VB.TextBox txtTemp 
      Height          =   375
      Index           =   1
      Left            =   1620
      TabIndex        =   7
      Top             =   1155
      Width           =   4335
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
      Left            =   4230
      TabIndex        =   40
      Top             =   4170
      Width           =   1335
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H80000018&
      Caption         =   "&Purchase Type"
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
      Height          =   1020
      Left            =   6150
      TabIndex        =   19
      Top             =   1335
      Width           =   3210
      Begin VB.OptionButton OptPUType 
         BackColor       =   &H80000018&
         Caption         =   "Inward"
         Height          =   360
         Index           =   3
         Left            =   1635
         TabIndex        =   23
         Top             =   600
         Width           =   990
      End
      Begin VB.OptionButton OptPUType 
         BackColor       =   &H80000018&
         Caption         =   "SIT"
         Height          =   360
         Index           =   2
         Left            =   120
         TabIndex        =   22
         Top             =   600
         Width           =   1110
      End
      Begin VB.OptionButton OptPUType 
         BackColor       =   &H80000018&
         Caption         =   "Trade"
         Height          =   360
         Index           =   1
         Left            =   1635
         TabIndex        =   21
         Top             =   225
         Width           =   1335
      End
      Begin VB.OptionButton OptPUType 
         BackColor       =   &H80000018&
         Caption         =   "All (Trade,SIT)"
         Height          =   360
         Index           =   0
         Left            =   120
         TabIndex        =   20
         Top             =   225
         Value           =   -1  'True
         Width           =   1635
      End
   End
   Begin VB.Frame Frame4 
      BackColor       =   &H80000018&
      Caption         =   "&Sale Type"
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
      Height          =   1020
      Left            =   6150
      TabIndex        =   12
      Top             =   150
      Width           =   3210
      Begin VB.OptionButton OptSAType 
         BackColor       =   &H80000018&
         Caption         =   "All"
         Height          =   360
         Index           =   0
         Left            =   120
         TabIndex        =   13
         Top             =   225
         Value           =   -1  'True
         Width           =   855
      End
      Begin VB.OptionButton OptSAType 
         BackColor       =   &H80000018&
         Caption         =   "Trade"
         Height          =   360
         Index           =   1
         Left            =   1125
         TabIndex        =   14
         Top             =   225
         Width           =   855
      End
      Begin VB.OptionButton OptSAType 
         BackColor       =   &H80000018&
         Caption         =   "Consign"
         Height          =   360
         Index           =   3
         Left            =   120
         TabIndex        =   16
         Top             =   540
         Width           =   975
      End
      Begin VB.OptionButton OptSAType 
         BackColor       =   &H80000018&
         Caption         =   "Depot"
         Height          =   360
         Index           =   4
         Left            =   1125
         TabIndex        =   17
         Top             =   540
         Width           =   975
      End
      Begin VB.OptionButton OptSAType 
         BackColor       =   &H80000018&
         Caption         =   "SIT"
         Height          =   360
         Index           =   2
         Left            =   2205
         TabIndex        =   15
         Top             =   225
         Width           =   930
      End
      Begin VB.OptionButton OptSAType 
         BackColor       =   &H80000018&
         Caption         =   "Mill Bill"
         Height          =   360
         Index           =   5
         Left            =   2205
         TabIndex        =   18
         Top             =   540
         Width           =   870
      End
   End
   Begin VB.CommandButton cmdPrint 
      Height          =   495
      Index           =   0
      Left            =   1095
      Picture         =   "frmRptGraph.frx":058A
      Style           =   1  'Graphical
      TabIndex        =   37
      Top             =   4170
      Width           =   1335
   End
   Begin VB.CommandButton cmdPrint 
      Height          =   495
      Index           =   1
      Left            =   2670
      Picture         =   "frmRptGraph.frx":09CC
      Style           =   1  'Graphical
      TabIndex        =   38
      ToolTipText     =   "Print"
      Top             =   4170
      Width           =   1335
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H80000018&
      Caption         =   "&Graph Type"
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
      Height          =   855
      Left            =   300
      TabIndex        =   33
      Top             =   2670
      Width           =   5625
      Begin VB.OptionButton OptGraphType 
         BackColor       =   &H80000018&
         Caption         =   "Pie"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   2
         Left            =   3885
         TabIndex        =   36
         Top             =   360
         Width           =   1425
      End
      Begin VB.OptionButton OptGraphType 
         BackColor       =   &H80000018&
         Caption         =   "Line"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   1
         Left            =   2085
         TabIndex        =   35
         Top             =   360
         Width           =   1425
      End
      Begin VB.OptionButton OptGraphType 
         BackColor       =   &H80000018&
         Caption         =   "Bar"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   0
         Left            =   210
         TabIndex        =   34
         Top             =   360
         Value           =   -1  'True
         Width           =   1425
      End
   End
   Begin Crystal.CrystalReport CrystalReport1 
      Left            =   5550
      Top             =   4080
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
   Begin MSComCtl2.DTPicker DTPicker1 
      Height          =   375
      Left            =   1635
      TabIndex        =   1
      Top             =   225
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   661
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
      Format          =   88670211
      CurrentDate     =   38050
   End
   Begin MSComCtl2.DTPicker DTPicker2 
      Height          =   375
      Left            =   4410
      TabIndex        =   3
      Top             =   225
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   661
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
      Format          =   88670211
      CurrentDate     =   38050
   End
   Begin MSComctlLib.ProgressBar prgBar 
      Height          =   255
      Left            =   270
      TabIndex        =   39
      Top             =   3615
      Visible         =   0   'False
      Width           =   5595
      _ExtentX        =   9869
      _ExtentY        =   450
      _Version        =   393216
      Appearance      =   1
   End
   Begin VB.Line Line1 
      BorderColor     =   &H000040C0&
      X1              =   6060
      X2              =   6060
      Y1              =   75
      Y2              =   5025
   End
   Begin VB.Line Line3 
      BorderColor     =   &H000040C0&
      X1              =   6045
      X2              =   90
      Y1              =   2535
      Y2              =   2535
   End
   Begin VB.Shape Shape2 
      FillColor       =   &H00BDDFD1&
      FillStyle       =   0  'Solid
      Height          =   825
      Left            =   645
      Shape           =   4  'Rounded Rectangle
      Top             =   3990
      Width           =   5220
   End
   Begin VB.Label lblMillName 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "&Mill              :"
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
      Left            =   240
      TabIndex        =   4
      Top             =   720
      Width           =   1200
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "&Broker         :"
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
      Left            =   240
      TabIndex        =   8
      Top             =   1605
      Width           =   1170
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "&Item            :"
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
      Left            =   240
      TabIndex        =   6
      Top             =   1155
      Width           =   1155
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "To Date   :"
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
      Left            =   3315
      TabIndex        =   2
      Top             =   225
      Width           =   975
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "From &Date   :"
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
      Left            =   240
      TabIndex        =   0
      Top             =   225
      Width           =   1230
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00000080&
      FillColor       =   &H00EDFDFE&
      FillStyle       =   0  'Solid
      Height          =   4980
      Left            =   75
      Top             =   60
      Width           =   9540
   End
End
Attribute VB_Name = "frmRptGraph"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim i As Integer
Dim wMillCode As Long, wBrkCode As Long, wItemCode As Long

Private Sub cmdExit_Click()
Unload Me
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
Select Case KeyCode
Case 13
    Sendkeys "{TAB}"
Case 27
    Unload Me
End Select
End Sub
Private Sub Form_Load()
Dim i As Long
Me.Top = 700
Me.Left = Screen.Width / 2 - Me.Width / 2
 '-- Activate Controls
GProcActivateControls Me, False
DTPicker1.Enabled = True
DTPicker2.Enabled = True
'-- Set Values
DTPicker1.Value = Format(gCYSDate, "dd/MM/yyyy")
DTPicker2.Value = Format(gCYEDate, "dd/MM/yyyy")
For i = 0 To 2 '-- Graph Type
    OptGraphType(i).Enabled = True
Next i
For i = 0 To 4 '-- Number Format
    OptNumFormat(i).Enabled = True
Next i
ChkConsolidate.Enabled = True
OptAmt.Enabled = True
OptBag.Enabled = True
Me.Caption = gSelectedMenu
Select Case gSelectedMenu
    Case "TCS  Return"
        Frame1.Visible = False
        Frame2.Visible = False
        Frame3.Visible = False
        Frame4.Visible = False
        Frame5.Visible = False
        ChkConsolidate.Visible = False
        ChkPuVsSA.Visible = True
        ChkPuVsSA.Enabled = True
        ChkPuVsSA.Caption = "TCS Db.Note On O/s"
        txttemp(0).Visible = False
        txttemp(1).Visible = False
        txttemp(2).Visible = False
        lblMillName.Visible = False
        Label5.Visible = False
        Label3.Visible = False

    Case "TDS (194Q) Purchase Report", "TCS AND TDS (194Q) Details"
        Frame1.Visible = False
        Frame2.Visible = False
        Frame3.Visible = False
        Frame4.Visible = False
        Frame5.Visible = False
        ChkConsolidate.Visible = False
        txttemp(0).Visible = False
        txttemp(1).Visible = False
        txttemp(2).Visible = False
        lblMillName.Visible = False
        Label5.Visible = False
        Label3.Visible = False
        If gSelectedMenu = "TCS AND TDS (194Q) Details" Then
           DTPicker1.Enabled = False
           DTPicker2.Enabled = False
        End If
    
    Case "Sales Graph" '== Sales Graph
        For i = 0 To 5
            OptSAType(i).Enabled = True
        Next i
    Case "Countwise Sales Graph"  '==  Countwise Sales Graph
        For i = 0 To 5
            OptSAType(i).Enabled = True
        Next i
        txttemp(1).Enabled = True
    Case "Millwise Sales Graph"  '==  Millwise Sales Graph
        For i = 0 To 5
            OptSAType(i).Enabled = True
        Next i
        txttemp(0).Enabled = True
    Case "Brokerwise Sales Graph"  '==  Brokerwise Sales Graph
        For i = 0 To 5
            OptSAType(i).Enabled = True
        Next i
        txttemp(2).Enabled = True
    Case "Purchase Graph" '== "Purchase Graph"
        For i = 0 To 3
            OptPUType(i).Enabled = True
        Next i
        ChkPuVsSA.Enabled = True
    Case "Countwise Purchase Graph"  '== "Countwise Purchase Graph"
        For i = 0 To 3
            OptPUType(i).Enabled = True
        Next i
        txttemp(1).Enabled = True
        ChkPuVsSA.Enabled = True
    Case "Millwise Purchase Graph"  '==  "Millwise Purchase Graph"
        For i = 0 To 3
            OptPUType(i).Enabled = True
        Next i
        txttemp(0).Enabled = True
        ChkPuVsSA.Enabled = True
End Select
 '-- Set InActivate Ctrl Back Color
GProcSetInActivateCtrlBackClr Me
End Sub
Private Sub cmdPrint_Click(Index As Integer)
On Error GoTo ErrorRoutine
Dim X As String
Dim DivideBy As Long
Dim DivideByStr As String
Dim SAType As String '-Sale Type
Dim PUType As String '- Purchase Type
Dim RptOptStr As String '- Report Str (Amt/Bag)
RptOptStr = IIf(OptAmt = True, "Amt", "Bag")
If txttemp(0) = "" Then wMillCode = 0
If txttemp(1) = "" Then wItemCode = 0
If txttemp(2) = "" Then wBrkCode = 0
'== Number Format
If OptNumFormat(1) = True Then '--Hundred
    DivideBy = 100
    DivideByStr = "Hundred"
ElseIf OptNumFormat(2) = True Then  '--Thousand
    DivideBy = 1000
    DivideByStr = "Thousand"
ElseIf OptNumFormat(3) = True Then  '--Lac
    DivideBy = 100000
    DivideByStr = "Lac"
ElseIf OptNumFormat(4) = True Then  '--Crore
    DivideBy = 10000000
    DivideByStr = "Crore"
End If
'== Sale Type
If OptSAType(0) = True Then '--all
    SAType = ""
ElseIf OptSAType(1) = True Then '--trade
    SAType = "SY"
ElseIf OptSAType(2) = True Then '--sit
    SAType = "ST"
ElseIf OptSAType(3) = True Then '--consig
    SAType = "SO"
ElseIf OptSAType(4) = True Then '--Depot
    SAType = "SD"
Else '--Mill bill
    SAType = "SM"
End If
'== Purchase Type
If OptPUType(0) = True Then '--all
    PUType = ""
ElseIf OptPUType(1) = True Then '--trade
    PUType = "PY"
ElseIf OptPUType(2) = True Then '--sit
    PUType = "PT"
Else '--Inward
    PUType = "PI"
End If
With CrystalReport1
    GProcCrystalRptPreparation CrystalReport1, IIf(Index = 0, "Window", "Printer")
    Select Case gSelectedMenu
        Case "TCS  Return"
             If ChkPuVsSA.Value = 1 Then '-- For TCS in Bill
              
                X = "Select 'A',OutAcCode,AcName,AcPAN,sum(OutAmount),'" & gUserName & "' from tblOutstanding,tblMastAccount where tblOutstanding.OutAcCode = tblMastAccount.AcCode and tblOutstanding.vtype in ('SY','SN') and " _
                    & " Vdt < cdate('01/10/2020') group by OutAcCode,AcName,AcPAN"
             
                gCn.BeginTrans
                gCn.Execute "Delete from TmpGentbl2 where  UserName='" & gUserName & "'"
                gCn.Execute " Insert Into TmpGentbl2 (vtype,Vno,Nar2,Nar1,Amt1,UserName) " & X
                gCn.CommitTrans
             
                X = "Select 'B',PartyCode,AcName,AcPAN,sum(AdjAmt),'" & gUserName & "' from tblRecVsSale,tblMastAccount where tblRecVsSale.PartyCode = tblMastAccount.AcCode and tblRecVsSale.InvTp in ('SY','SN') and " _
                    & " RecVDt < cdate('01/10/2020') group by PartyCode,AcName,AcPAN"
             
                gCn.BeginTrans
                gCn.Execute " Insert Into TmpGentbl2 (vtype,Vno,Nar2,Nar1,Amt2,UserName) " & X
                gCn.CommitTrans
             
             
                X = "Select 'C',qrytblVoucherSr1.VacCode,AcName,AcPAN,sum(abs(qrytblVoucherSr1.Vamt)),'" & gUserName & "'  from  qrytblVoucherSr1,tblMastAccount where qrytblVoucherSr1.VacCode = tblMastAccount.AcCode and qrytblVoucherSr1.Vtype in ('BR','CR') and " _
                    & " VDt < cdate('01/10/2020') and tblMastAccount.Agcode = 90017 group by qrytblVoucherSr1.VacCode,AcName,AcPAN"
                
                gCn.BeginTrans
                gCn.Execute " Insert Into TmpGentbl2 (vtype,Vno,Nar2,Nar1,Amt3,UserName) " & X
                gCn.CommitTrans
             
                X = "Select 'D',Vno,Nar2,Nar1,sum(Amt1)-sum(Amt2),sum(Amt3),'" & gUserName & "' from TmpGentbl2 where username = '" & gUserName & "' group by Vno,Nar2,Nar1"
                gCn.BeginTrans
                gCn.Execute " Insert Into TmpGentbl2 (vtype,Vno,Nar2,Nar1,Amt2,Amt1,UserName) " & X
                gCn.Execute "delete from tmpGenTbl2 where Username='" & gUserName & "' and vtype <> 'D' "
                gCn.CommitTrans
                gCn.BeginTrans
                gCn.Execute "delete from tmpGenTbl2 where Username='" & gUserName & "' and vtype = 'D' and amt2+amt1 < 5000000 "
                
                gCn.CommitTrans
                X = "{TmpGentbl2.UserName}='" & gUserName & "'"
                .SelectionFormula = X
                .Formulas(4) = "wFrToDt = 'From : ' & '" & Format(DTPicker1.Value, "dd/MM/yyyy") & "' & '  To ' & '" & Format(DTPicker2.Value, "dd/MM/yyyy") & "'"
                .Formulas(5) = "wcPAN= '" & gCPAN & "'"
                
                .ReportFileName = gReportPath & "rptTcsOnOS.rpt"
             Else
                X = "Select 'A',AcCode,AcName,AcPAN,sum(AdTcsOn),sum(AdTcsAmt),'" & gUserName & "' from tblAddLess,tblMastAccount where tblAddLess.AdAcDrCode = tblMastAccount.AcCode and " _
                    & "vdt >= cdate('" & DTPicker1.Value & "') and vdt <= cdate('" & DTPicker2.Value & "') and tblAddLess.AdTcsAmt > 0 and tblAddLess.Vtype = 'SY' " _
                    & "group by AcCode,AcName,AcPAN"
                gCn.BeginTrans
                gCn.Execute "Delete from TmpGentbl2 where  UserName='" & gUserName & "'"
                gCn.Execute " Insert Into TmpGentbl2 (vtype,Vno,Nar2,Nar1,Amt2,Amt1,UserName) " & X
                gCn.CommitTrans
            
                X = "Select 'B',AcCode,AcName,AcPAN,sum(VTaxableAmt),sum(abs(VAmt)),'" & gUserName & "' from tblVoucher,tblMastAccount where tblVoucher.VAcCode = tblMastAccount.AcCode and " _
                    & "vdt >= cdate('" & DTPicker1.Value & "') and vdt <= cdate('" & DTPicker2.Value & "') and tblVoucher.Vtype = 'SN' and tblVoucher.VDRCRNTTYPE = 'T' and tblVoucher.VCtrNo = 1 " _
                    & "group by AcCode,AcName,AcPAN"
                gCn.BeginTrans
                gCn.Execute " Insert Into TmpGentbl2 (vtype,Vno,Nar2,Nar1,Amt2,Amt1,UserName) " & X
                gCn.CommitTrans
                
                X = "Select 'C',Vno,Nar2,Nar1,sum(Amt2),sum(Amt1),'" & gUserName & "' from TmpGentbl2 where username = '" & gUserName & "' group by Vno,Nar2,Nar1"
                gCn.BeginTrans
                gCn.Execute " Insert Into TmpGentbl2 (vtype,Vno,Nar2,Nar1,Amt2,Amt1,UserName) " & X
                gCn.Execute "delete from tmpGenTbl2 where Username='" & gUserName & "' and vtype <> 'C' "
                gCn.CommitTrans
                
                
                X = "{TmpGentbl2.UserName}='" & gUserName & "'"
                .SelectionFormula = X
                .Formulas(4) = "wFrToDt = 'From : ' & '" & Format(DTPicker1.Value, "dd/MM/yyyy") & "' & '  To ' & '" & Format(DTPicker2.Value, "dd/MM/yyyy") & "'"
                .Formulas(5) = "wcPAN= '" & gCPAN & "'"
                
                .ReportFileName = gReportPath & "rptTcsReturn.rpt"
             End If
     
        '==Sales Graph
        Case "Sales Graph", "Countwise Sales Graph", "Millwise Sales Graph", "Brokerwise Sales Graph"
            GProcGraphSales Format(DTPicker1.Value, "dd/MM/yyyy"), Format(DTPicker2.Value, "dd/MM/yyyy"), ChkConsolidate, RptOptStr, SAType, DivideBy, wMillCode, wBrkCode, wItemCode
            .Formulas(4) = "wFrToDt = 'From : ' & '" & Format(DTPicker1.Value, "dd/MM/yyyy") & "' & '  To ' & '" & Format(DTPicker2.Value, "dd/MM/yyyy") & "'"
            If ChkConsolidate = 1 Then .Formulas(5) = "wAllFirmStr= 'All Firm'" '- wAllFirmStr
            Select Case SAType '-- Graph Title
                Case ""
                    X = "Total Sales"
                Case "SY"
                    X = "Trade Sales"
                Case "SD"
                    X = "Depot Sales"
                Case "ST"
                    X = "SIT Sales"
                Case "SO"
                    X = "Consignment Sales"
                Case "SM"
                    X = "Mill Bill Sales"
            End Select
            '-- GraphTitle
            .Formulas(6) = "GraphTitle='" & X & " Month VS " & IIf(OptAmt = True, "Amount", "Bags") & IIf(DivideByStr <> "", " ( in " & DivideByStr & ")", "") & "'"
            If wBrkCode > 0 Then
                .Formulas(7) = "Str='( For Broker : " & txttemp(2) & ")'"
            ElseIf wMillCode > 0 Then
                .Formulas(7) = "Str='( For Mill : " & txttemp(0) & ")'"
            ElseIf wItemCode > 0 Then
                .Formulas(7) = "Str='( For Count : " & txttemp(1) & ")'"
            End If
            .DataFiles(0) = App.Path & "\HIData\setup.mdb"
            .SelectionFormula = "{tmpgenTbl.Username}='" & gUserName & "'"
            gReportPath = App.Path & "\HiReports\"
            If OptGraphType(0) = True Then '-Bar
                .ReportFileName = gReportPath & "rptGraphSalesBar.rpt"
            ElseIf OptGraphType(1) = True Then     '-Line
                .ReportFileName = gReportPath & "rptGraphSalesLine.rpt"
            Else '- Pie
                .ReportFileName = gReportPath & "rptGraphSalesPie.rpt"
            End If
        '==Purchase Graph
        Case "Purchase Graph", "Countwise Purchase Graph", "Millwise Purchase Graph"
            If ChkPuVsSA = 0 Then '== Purchase
                GProcGraphPurchase Format(DTPicker1.Value, "dd/MM/yyyy"), Format(DTPicker2.Value, "dd/MM/yyyy"), ChkConsolidate, RptOptStr, PUType, DivideBy, wMillCode, wItemCode
            Else '= Purchase VS Sales
                GProcGraphPurchaseVsSale Format(DTPicker1.Value, "dd/MM/yyyy"), Format(DTPicker2.Value, "dd/MM/yyyy"), ChkConsolidate, RptOptStr, PUType, DivideBy, wMillCode, wItemCode
            End If
            Select Case PUType '-- Graph Title
                Case ""
                    If ChkPuVsSA = 0 Then
                        X = "Total Purchase"
                    Else
                        X = "Total Purchase and Sales"
                    End If
                Case "PY"
                    If ChkPuVsSA = 0 Then
                        X = "Trade Purchase"
                    Else
                        X = "Trade Purchase and Sales"
                    End If
                Case "PT"
                    If ChkPuVsSA = 0 Then
                        X = "SIT Purchase"
                    Else
                        X = "SIT Purchase and Sales"
                    End If
                Case "PI"
                    If ChkPuVsSA = 0 Then
                        X = "Purchase Inward"
                    Else
                        X = "Purchase Inward and Sales"
                    End If
            End Select
            '-- GraphTitle
            .Formulas(6) = "GraphTitle='" & X & " Month VS " & IIf(OptAmt = True, "Amount", "Bags") & IIf(DivideByStr <> "", " ( in " & DivideByStr & ")", "") & "'"
            .Formulas(4) = "wFrToDt = 'From : ' & '" & Format(DTPicker1.Value, "dd/MM/yyyy") & "' & '  To ' & '" & Format(DTPicker2.Value, "dd/MM/yyyy") & "'"
            If ChkConsolidate = 1 Then .Formulas(5) = "wAllFirmStr= 'All Firm'" '- wAllFirmStr
            If wMillCode > 0 Then '-Mill
                .Formulas(7) = "Str='( For Mill : " & txttemp(0) & ")'"
            ElseIf wItemCode > 0 Then '-Count
                .Formulas(7) = "Str='( For Count : " & txttemp(1) & ")'"
            End If
            .DataFiles(0) = App.Path & "\HIData\setup.mdb"
            .SelectionFormula = "{tmpgenTbl.Username}='" & gUserName & "'"
            gReportPath = App.Path & "\HiReports\"
            If OptGraphType(0) = True Then '-Bar
                 If ChkPuVsSA = 0 Then
                    .ReportFileName = gReportPath & "rptGraphPurchaseBar.rpt"
                 Else
                    .ReportFileName = gReportPath & "rptGraphPurVsSaleBar.rpt"
                 End If
            ElseIf OptGraphType(1) = True Then     '-Line
                If ChkPuVsSA = 0 Then
                    .ReportFileName = gReportPath & "rptGraphPurchaseLine.rpt"
                Else
                    .ReportFileName = gReportPath & "rptGraphPurVsSaleLine.rpt"
                End If
            Else '- Pie
                If ChkPuVsSA = 0 Then
                    .ReportFileName = gReportPath & "rptGraphPurchasePie.rpt"
                Else
                    .ReportFileName = gReportPath & "rptGraphPurVsSalePie.rpt"
                End If
            End If
         Case "TCS AND TDS (194Q) Details"
              Dim JunDt As String
              '---Tcs On Collected (Sale)
              JunDt = "01/06/" + Trim(Year(gCYSDate))
              X = "Select 'A',tblmastaccount.AcName, tblAccO.acname as  AcOName,tblAccO.AcPan,sum(iif(vamt>0,vamt,0)) AS Credit, " _
                  & " sum(iif(tblvoucher.Vdt<= cdate('" & JunDt & "'),iif(vamt>0,vamt,0),0)),sum(iif(tblvoucher.Vdt> cdate('" & JunDt & "'),iif(vamt>0,vamt,0),0)),'" & gUserName & "' " _
                  & " From tblvoucher, tblMastAccount , tblMastAccount as  tblAccO  Where tblvoucher.vacCode = tblMastAccount.AcCode  and tblvoucher.VAcOCode = tblAccO.AcCode " _
                  & " and tblvoucher.Vdt>= cdate('" & DTPicker1.Value & "') and tblvoucher.Vdt<=cdate('" & DTPicker2.Value & "')  and tblmastaccount.AcCode = " & gAcCodeTcsPay & " " _
                  & " and tblAccO.acname <> '-' group by tblmastaccount.AcName, tblAccO.acname,tblAccO.AcPan"
              gCn.BeginTrans
              gCn.Execute "Delete from TmpGentbl2 where  UserName='" & gUserName & "'"
              gCn.Execute " Insert Into TmpGentbl2 (VTYPE,Nar1,Nar2,Nar3,Amt1,Amt2,Amt3,UserName) " & X
              gCn.CommitTrans
              '---Tcs On deducted(Purchase)
              X = "Select 'A',tblmastaccount.AcName, tblAccO.acname as  AcOName,tblAccO.AcPan,sum(iif(vamt<0,abs(vamt),0)) AS Debit, " _
                  & " sum(iif(tblvoucher.Vdt<= cdate('" & JunDt & "'),iif(vamt<0,abs(vamt),0),0)),sum(iif(tblvoucher.Vdt> cdate('" & JunDt & "'),iif(vamt<0,abs(vamt),0),0)),'" & gUserName & "' " _
                  & " From tblvoucher, tblMastAccount , tblMastAccount as  tblAccO  Where tblvoucher.vacCode = tblMastAccount.AcCode  and tblvoucher.VAcOCode = tblAccO.AcCode " _
                  & " and tblvoucher.Vdt>= cdate('" & DTPicker1.Value & "') and tblvoucher.Vdt<=cdate('" & DTPicker2.Value & "')  and tblmastaccount.AcCode = " & gAcCodeTcsRec & " " _
                  & " and tblAccO.acname <> '-' group by tblmastaccount.AcName, tblAccO.acname,tblAccO.AcPan"
              
              gCn.BeginTrans
              gCn.Execute " Insert Into TmpGentbl2 (VTYPE,Nar1,Nar2,Nar3,Amt4,Amt5,Amt6,UserName) " & X
              gCn.CommitTrans
              '---Tds On Purchase 194 Q
              X = "Select 'A',tblmastaccount.AcName, tblAccO.acname as  AcOName,tblAccO.AcPan,sum(iif(vamt>0,vamt,0)) AS Credit,'" & gUserName & "' " _
                  & " From tblvoucher, tblMastAccount , tblMastAccount as  tblAccO  Where tblvoucher.vacCode = tblMastAccount.AcCode  and tblvoucher.VAcOCode = tblAccO.AcCode " _
                  & " and tblvoucher.Vdt>= cdate('" & DTPicker1.Value & "') and tblvoucher.Vdt<=cdate('" & DTPicker2.Value & "')  and tblmastaccount.AcCode = " & gServiceTaxAcCode & " " _
                  & " and tblAccO.acname <> '-' group by tblmastaccount.AcName, tblAccO.acname,tblAccO.AcPan"
         
              gCn.BeginTrans
              gCn.Execute " Insert Into TmpGentbl2 (VTYPE,Nar1,Nar2,Nar3,Amt7,UserName) " & X
              gCn.CommitTrans
         
              X = "select 'B',Nar2,Nar3,sum(amt1),sum(amt2),sum(amt3),sum(amt4),sum(amt5),sum(amt6),sum(amt7),'" & gUserName & "' " _
                  & " from TmpGentbl2 where UserName = '" & gUserName & "' group by Nar2,Nar3 "
                    
              gCn.BeginTrans
              gCn.Execute " Insert Into TmpGentbl2 (VTYPE,Nar2,Nar3,Amt1,Amt2,Amt3,Amt4,Amt5,Amt6,Amt7,UserName) " & X
              gCn.CommitTrans
              
              gCn.BeginTrans
              gCn.Execute "Delete from TmpGentbl2 where vtype = 'A' and UserName='" & gUserName & "'"
              gCn.CommitTrans
              
              gCn.BeginTrans
              gCn.Execute "Delete from TmpGentbl2 where amt1+amt2+amt3+amt4+amt5+amt6+amt7 = 0 and UserName='" & gUserName & "'"
              gCn.CommitTrans
         
         
              X = "{TmpGentbl2.UserName}='" & gUserName & "'"
             .SelectionFormula = X
             .Formulas(4) = "wFrToDt = 'From : ' & '" & Format(DTPicker1.Value, "dd/MM/yyyy") & "' & '  To ' & '" & Format(DTPicker2.Value, "dd/MM/yyyy") & "'"
             .Formulas(5) = "wcPAN= '" & gCPAN & "'"
             .ReportFileName = gReportPath & "rptTcsTdsDetails.rpt"
         Case "TDS (194Q) Purchase Report"
              JunDt = "01/07/" + Trim(Year(gCYSDate))
              X = "Select 'A',AcCode,AcName,AcPAN,AcAdd1,AcAdd2,AcCity,sum(iif(vdt < cdate('" & JunDt & "'),AdBillAmt,0)),sum(iif(vdt < cdate('" & JunDt & "') and AdOtherLess > 0 ,AdSubAmt,0)), " _
                  & " sum(iif(vdt >= cdate('" & JunDt & "'),AdBillAmt,0)),sum(iif(vdt >= cdate('" & JunDt & "') and AdOtherLess > 0    ,AdSubAmt,0)),sum(iif(vdt >= cdate('" & JunDt & "'),AdOtherLess,0)),'" & gUserName & "' " _
                  & " from tblAddLess,tblMastAccount where tblAddLess.AdAcCrCode = tblMastAccount.AcCode and " _
                    & "vdt >= cdate('" & DTPicker1.Value & "') and vdt <= cdate('" & DTPicker2.Value & "') and tblAddLess.Vtype = 'PY' " _
                    & "group by AcCode,AcName,AcPAN,AcAdd1,AcAdd2,AcCity"
                gCn.BeginTrans
                gCn.Execute "Delete from TmpGentbl2 where  UserName='" & gUserName & "'"
                gCn.Execute " Insert Into TmpGentbl2 (vtype,Vno,Nar2,Nar1,Nar3,Nar4,Nar5,Amt1,Amt2,Amt3,Amt4,Amt5,UserName) " & X
                gCn.CommitTrans
                X = "{TmpGentbl2.UserName}='" & gUserName & "'"
                .SelectionFormula = X
                .Formulas(4) = "wFrToDt = 'From : ' & '" & Format(DTPicker1.Value, "dd/MM/yyyy") & "' & '  To ' & '" & Format(DTPicker2.Value, "dd/MM/yyyy") & "'"
                .Formulas(5) = "wcPAN= '" & gCPAN & "'"
                .ReportFileName = gReportPath & "rptTds194QPurch.rpt"
    
    End Select
    .Action = 1
End With
Exit Sub
ErrorRoutine:
    GProcErrorHandler
End Sub

Private Sub txttemp_GotFocus(Index As Integer)
GProcSelectBox Me.ActiveControl
End Sub
Private Sub txtTemp_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = 13 Then Exit Sub '-- Enter
    Select Case Index
     Case 0 '----Mill Name
        gClsSearch.SearchMultiField "QryMastaccount", "Account,AcName,ACCode", Array("Mill", "AcName", ""), Array(txttemp(Index).Width, 0, 0), " (AgCode=90029 or GpCode=90029)", Chr(KeyAscii), "AcName", txttemp(Index).Left + Me.Left, txttemp(Index).Top + Me.Top + 1000
        KeyAscii = 0
        If gClsSearch.SearchMultiRetCol(1) <> "" Then
            txttemp(Index) = gClsSearch.SearchMultiRetCol(1)
            wMillCode = gClsSearch.SearchMultiRetCol(2)
        End If
     Case 1 '-- Item
        gClsSearch.SearchMultiField "tblMastItem,tblMastAccount", "ItName,ItCode,Itticket,Acname", Array("Item", "", "Ticket", "Mill"), Array(2000, 0, 1000, 2000), " tblMastItem.ItMillcode=tblMastAccount.AcCode", Chr(KeyAscii), "ItName", txttemp(Index).Left + Me.Left, txttemp(Index).Top + Me.Top + 1000
        KeyAscii = 0
        If gClsSearch.SearchMultiRetCol(0) <> "" Then
            txttemp(Index) = gClsSearch.SearchMultiRetCol(0) '-ItName
            txttemp(0) = gClsSearch.SearchMultiRetCol(3) '-Mill
            wItemCode = gClsSearch.SearchMultiRetCol(1)
        End If
     Case 2 '--Broker
        gClsSearch.SearchMultiField "QryMastaccount", "Account,AcName,ACCode", Array("Broker", "AcName", ""), Array(txttemp(Index).Width, 0, 0), " (AgCode=90016 or GpCode=90016)", Chr(KeyAscii), "AcName", txttemp(Index).Left + Me.Left, txttemp(Index).Top + Me.Top + 1000
        KeyAscii = 0
        If gClsSearch.SearchMultiRetCol(1) <> "" Then
            txttemp(Index) = gClsSearch.SearchMultiRetCol(1)
            wBrkCode = gClsSearch.SearchMultiRetCol(2)
        End If
    End Select
End Sub

