VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form frmRptPartyIntCal 
   BackColor       =   &H80000013&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Late Payment Interest Calculation"
   ClientHeight    =   7650
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   13860
   Icon            =   "frmRptPartyIntCal.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7650
   ScaleWidth      =   13860
   Begin VB.TextBox txt 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
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
      Left            =   3300
      TabIndex        =   8
      Top             =   4140
      Width           =   1635
   End
   Begin VB.TextBox txt 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
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
      Left            =   3300
      TabIndex        =   7
      Top             =   3540
      Width           =   1635
   End
   Begin VB.TextBox txt 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
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
      Left            =   3300
      TabIndex        =   6
      Top             =   2940
      Width           =   1635
   End
   Begin VB.TextBox txt 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
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
      Left            =   3300
      TabIndex        =   4
      Top             =   1620
      Width           =   1635
   End
   Begin VB.TextBox txt 
      Appearance      =   0  'Flat
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
      Left            =   3300
      TabIndex        =   3
      Top             =   1020
      Width           =   1635
   End
   Begin VB.CommandButton cmdSPE 
      Height          =   495
      Index           =   0
      Left            =   3405
      Picture         =   "frmRptPartyIntCal.frx":058A
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   8400
      Width           =   1335
   End
   Begin VB.CommandButton cmdSPE 
      Height          =   495
      Index           =   1
      Left            =   4980
      Picture         =   "frmRptPartyIntCal.frx":09CC
      Style           =   1  'Graphical
      TabIndex        =   10
      ToolTipText     =   "Print"
      Top             =   8400
      Width           =   1335
   End
   Begin Crystal.CrystalReport CrystalReport1 
      Left            =   14520
      Top             =   8220
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
   Begin VB.CommandButton cmdSPE 
      Caption         =   "E&xit"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   2
      Left            =   12360
      TabIndex        =   11
      Top             =   5820
      Width           =   1335
   End
   Begin MSComCtl2.DTPicker DTPicker1 
      Height          =   375
      Left            =   3300
      TabIndex        =   5
      Top             =   2280
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
      Format          =   16580611
      CurrentDate     =   38050
   End
   Begin VB.TextBox txt 
      Appearance      =   0  'Flat
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
      Left            =   3300
      TabIndex        =   2
      Top             =   480
      Width           =   4335
   End
   Begin VB.ListBox lstParty 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Left            =   3300
      Sorted          =   -1  'True
      TabIndex        =   12
      Top             =   540
      Width           =   4335
   End
   Begin MSComctlLib.ProgressBar prgBar 
      Height          =   255
      Left            =   4020
      TabIndex        =   13
      Top             =   9180
      Visible         =   0   'False
      Width           =   5595
      _ExtentX        =   9869
      _ExtentY        =   450
      _Version        =   393216
      Appearance      =   1
   End
   Begin VB.Label lblBroker 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Broker Name"
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
      Left            =   10200
      TabIndex        =   36
      Top             =   1080
      Width           =   3225
   End
   Begin VB.Label lblShortInt 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Short Interest"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   285
      Left            =   8940
      TabIndex        =   35
      Top             =   6600
      Width           =   3045
   End
   Begin VB.Label Label13 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Broker           :"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   240
      Left            =   8460
      TabIndex        =   34
      Top             =   1080
      Width           =   1350
   End
   Begin VB.Label lblReqInt 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Int.Required"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   285
      Left            =   8940
      TabIndex        =   33
      Top             =   5580
      Width           =   2895
   End
   Begin VB.Label lblIntRecdAt 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Int.Recd.At@"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   285
      Left            =   3540
      TabIndex        =   32
      Top             =   6600
      Width           =   2565
   End
   Begin VB.Label Label12 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Interest Short         :"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   285
      Left            =   6420
      TabIndex        =   31
      Top             =   6600
      Width           =   2175
   End
   Begin VB.Label Label11 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Interest Must be    :"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   285
      Left            =   6420
      TabIndex        =   30
      Top             =   5580
      Width           =   2145
   End
   Begin VB.Label Label10 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Interest Recd. @   :"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   285
      Left            =   1140
      TabIndex        =   29
      Top             =   6600
      Width           =   2145
   End
   Begin VB.Label lblLateDays 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Late Days "
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   285
      Left            =   3600
      TabIndex        =   28
      Top             =   5580
      Width           =   2340
   End
   Begin VB.Label Label9 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Late Days               :"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   285
      Left            =   1140
      TabIndex        =   27
      Top             =   5580
      Width           =   2130
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Interest Received   :"
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
      Index           =   3
      Left            =   1260
      TabIndex        =   26
      Top             =   2940
      Width           =   1890
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Interest Rate        :"
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
      Index           =   2
      Left            =   1260
      TabIndex        =   25
      Top             =   4200
      Width           =   1740
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Grase Days           :"
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
      Left            =   1260
      TabIndex        =   24
      Top             =   3600
      Width           =   1785
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Received Agt.Bill  :"
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
      Left            =   1260
      TabIndex        =   23
      Top             =   1680
      Width           =   1815
   End
   Begin VB.Label lblDueDate 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Due Date"
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
      Left            =   10200
      TabIndex        =   22
      Top             =   3300
      Width           =   2910
   End
   Begin VB.Label lblDueDays 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Due Days"
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
      Left            =   10200
      TabIndex        =   21
      Top             =   4080
      Width           =   2625
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Due Date       :"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   240
      Left            =   8460
      TabIndex        =   20
      Top             =   3300
      Width           =   1350
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Due Days      :"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   240
      Left            =   8460
      TabIndex        =   19
      Top             =   4080
      Width           =   1305
   End
   Begin VB.Label lblBillAmount 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Bill Amount "
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
      Left            =   10200
      TabIndex        =   18
      Top             =   2520
      Width           =   2850
   End
   Begin VB.Label lblBillDate 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Bill Date    "
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
      Left            =   10200
      TabIndex        =   16
      Top             =   1800
      Width           =   2730
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Bill Amount   :"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   240
      Left            =   8460
      TabIndex        =   17
      Top             =   2520
      Width           =   1350
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Bill Date        :"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   240
      Left            =   8460
      TabIndex        =   15
      Top             =   1800
      Width           =   1350
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Bill No                   :"
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
      Left            =   1260
      TabIndex        =   14
      Top             =   1080
      Width           =   1830
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00000080&
      FillColor       =   &H00F4FDD9&
      FillStyle       =   0  'Solid
      Height          =   4590
      Left            =   8040
      Top             =   240
      Width           =   5520
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Party Name           :"
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
      Left            =   1260
      TabIndex        =   1
      Top             =   540
      Width           =   1830
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Received Date      :"
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
      Left            =   1260
      TabIndex        =   0
      Top             =   2280
      Width           =   1800
   End
   Begin VB.Shape Shape5 
      BorderColor     =   &H00000080&
      FillColor       =   &H00EDFDFE&
      FillStyle       =   0  'Solid
      Height          =   4650
      Left            =   435
      Top             =   240
      Width           =   7380
   End
   Begin VB.Shape Shape2 
      FillColor       =   &H00BDDFD1&
      FillStyle       =   0  'Solid
      Height          =   765
      Left            =   2955
      Shape           =   4  'Rounded Rectangle
      Top             =   8280
      Width           =   5220
   End
   Begin VB.Shape Shape3 
      BorderColor     =   &H00000080&
      FillColor       =   &H00FFC0FF&
      FillStyle       =   0  'Solid
      Height          =   2250
      Left            =   480
      Top             =   5160
      Width           =   11700
   End
End
Attribute VB_Name = "frmRptPartyIntCal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim i As Integer, LnCnt As Integer, PgCnt As Integer, wDotLine As String
Dim wAcCode As Long, mScreen As Boolean, mP12 As String, mPc As String
Dim wPartyCd(20) As String, DueDate As Date
'==== Format Text Boxes Value
Private Sub CheckForNumber()
Dim i As Integer
'-- Format txttmp
For i = 2 To 5
    If IsNumeric(txt(i)) = False Then
        txt(i) = 0
    End If
    '--format
    If i = 4 Then '-Rate Double
        txt(i) = Format(txt(i), PStr)
    Else
        txt(i) = Format(txt(i), FStr)
    End If
Next i
End Sub
'-- Validate Data
Private Function ValidateData() As Boolean
ValidateData = True
Select Case gSelectedMenu
'    Case "Sarafi Interest Report", "Sarafi Interest Report ANK" '-"Sarafi Interest Report"
'        '--Interest %
'        If CDbl(txttmp(0)) = 0 Or CDbl(txttmp(0)) > 100 Then
'            ValidateData = False
'            MsgBox "Check Interest %", vbOKOnly + vbInformation
'            txttmp(0).SetFocus
'            Exit Function
'        End If
        '-- Total Days for int Rate Yearly
        '-Days of year
End Select
End Function
Private Sub cmdSPE_Click(Index As Integer)
Dim sTp As String
Dim FromDt As Date
Dim ToDt As Date
Dim X As String
Select Case Index
Case 0
    mScreen = True
Case 1
    mScreen = False
Case 2
   Unload Me
   Exit Sub
End Select
End Sub
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
Select Case KeyCode
Case 13
    SendKeys "{TAB}"
Case 27
    Unload Me
End Select
End Sub
Private Sub Form_Load()
Dim i As Long
Me.Top = 700
Me.Left = Screen.Width / 2 - Me.Width / 2
 '-- Activate & Set  Controls
GProcActivateControls Me, True
'-- Set Values
DTPicker1.Value = Format(Now, "dd/MM/yyyy")
DueDate = Format(Now, "dd/MM/yyyy")
lblBillAmount.Caption = ""
lblBillDate.Caption = ""
lblDueDate.Caption = ""
lblDueDays.Caption = ""
lblIntRecdAt.Caption = ""
lblLateDays.Caption = ""
lblReqInt.Caption = ""
lblShortInt.Caption = ""



CheckForNumber
 '-- Set InActivate Ctrl Back Color
GProcSetInActivateCtrlBackClr Me
End Sub
Private Sub txt_GotFocus(Index As Integer)
'If index = 1 And lstParty.Visible = True Then
 '   lstParty.SetFocus
'Else
    GProcSelectBox Me.ActiveControl
'End If
End Sub
Private Sub txt_KeyPress(Index As Integer, KeyAscii As Integer)
Dim LeftPos As Long
Dim TopPos As Long
TopPos = txt(Index).Top + Me.Top + 650
LeftPos = txt(Index).Left + Me.Left
If KeyAscii = 13 Then Exit Sub
Select Case Index
    Case 1 '-Account
        '-Selected Group
        gClsSearch.SearchMultiField "QryMastaccount", "Account,AcName,Accode,AcCity,AgName", Array("Account", "AcCode", "AcName", "City", "Group"), Array(3000, 0, 0, 2000, 3000), "", Chr(KeyAscii), "AcName", LeftPos, TopPos
        KeyAscii = 0
        If gClsSearch.SearchMultiRetCol(0) <> "" And gClsSearch.SearchMultiRetCol(1) = "" Then
            txt(Index) = gClsSearch.SearchMultiRetCol(0)
            wAcCode = gClsSearch.SearchMultiRetCol(2)
        ElseIf gClsSearch.SearchMultiRetCol(0) <> "" And gClsSearch.SearchMultiRetCol(1) <> "" Then
            txt(Index) = gClsSearch.SearchMultiRetCol(1)
            wAcCode = gClsSearch.SearchMultiRetCol(2)
        End If
   Case 0
        Pending_Invoice
   
End Select
End Sub

Private Sub Pending_Invoice()
Dim X As String
Dim wFromCd As Long
Dim LateDayCd As String
Dim vamt As Double
Dim rstOutStanding As Recordset
Dim i As Long, j As Long
Dim NewRow As Boolean
If Trim(txt(1)) <> "" Then
    wFromCd = GProcGetColumnValue("tblMastAccount", "AcName", txt(1), "S", "AcCode", "N")
    X = "select VType, outbillNo,format(vdt,'dd/mm/yyyy'), OutAmount,OutDueDays,outDueDate" _
        & ",tblMastAccount.Acname from tblOutstanding, tblMastAccount where tblOutstanding.OutBrokerCode = tblMastAccount.AcCode and OutAcCode=" & wFromCd & " and " _
        & " ((OutAmount-OutRecAmt) > 0 or (OutDbNtAmt-OutRecDbAmt) > 0) "
    gClsSearch.SearchMultiField X, "", Array("Type", "Invoice No", "Inv.Date", "Bill Amount", "Due Days", "Due Date", "Broker"), Array(800, 1200, 1200, 2000, 800, 1200, 3000), "", "", "Vdt,outbillNo", 0, 0, False, , True
    If gClsSearch.SearchMultiRetCol(0) <> "" Then
       txt(0).Text = gClsSearch.SearchMultiRetCol(1)
       txt(2).Text = gClsSearch.SearchMultiRetCol(3)
       lblBillDate.Caption = Format(gClsSearch.SearchMultiRetCol(2), "dd/mm/yyyy")
       lblBillAmount.Caption = gClsSearch.SearchMultiRetCol(3)
       lblDueDays.Caption = gClsSearch.SearchMultiRetCol(4)
       lblDueDate.Caption = Format(gClsSearch.SearchMultiRetCol(5), "dd/mm/yyyy")
       DueDate = Format(gClsSearch.SearchMultiRetCol(5))
       lblBroker.Caption = gClsSearch.SearchMultiRetCol(6)
    End If
    txt(2).SetFocus
Else
    MsgBox "Select Party Compulsary", vbCritical + vbOKOnly, "Select Party"
    DoEvents
    txt(1).SetFocus
End If
End Sub

Private Sub txt_LostFocus(Index As Integer)
CheckForNumber
CalAmount
End Sub

Private Sub CalAmount()
Dim Ldays As Integer
Dim RIntRate As Double
Dim CalInt As Double
Dim Grase As Integer
Dim Short As Double

Grase = CDbl(txt(4))
If IsDate(DueDate) Then
   Ldays = CDate(DTPicker1.Value) - DueDate
Else
   Ldays = 0
End If
lblLateDays.Caption = Ldays - Grase
'----- Int Recd @
If CDbl(txt(2)) > 0 And Ldays > 0 Then
   RIntRate = (CDbl(txt(3)) * 30 * 100) / (CDbl(txt(2)) * (Ldays - Grase))
Else
   RIntRate = 0
End If
RIntRate = Format(RIntRate, FStr)
lblIntRecdAt.Caption = RIntRate
'--Int As per Our Int Rate
CalInt = Round(((CDbl(txt(2)) * CDbl(txt(5))) / (12 * 365)) * (Ldays - Grase), 0)




CalInt = Format(CalInt, PStr)

lblReqInt.Caption = CalInt

Short = CalInt - CDbl(txt(3))
Short = Format(Short, PStr)

lblShortInt.Caption = Short

End Sub
