VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form frmRpTWhatsApp 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Report"
   ClientHeight    =   5895
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9855
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmRptWhatsApp.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5895
   ScaleWidth      =   9855
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtLate 
      Alignment       =   1  'Right Justify
      Height          =   375
      Left            =   4440
      TabIndex        =   22
      Top             =   2400
      Visible         =   0   'False
      Width           =   1035
   End
   Begin VB.CommandButton cmdView 
      Caption         =   "View"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   0
      Left            =   10260
      Style           =   1  'Graphical
      TabIndex        =   20
      Top             =   2580
      Visible         =   0   'False
      Width           =   675
   End
   Begin VB.TextBox txtTemp 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Index           =   8
      Left            =   6660
      MaxLength       =   15
      TabIndex        =   8
      Top             =   1575
      Visible         =   0   'False
      Width           =   2655
   End
   Begin VB.TextBox txtTemp 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Index           =   7
      Left            =   6660
      MaxLength       =   15
      TabIndex        =   6
      Top             =   1125
      Visible         =   0   'False
      Width           =   2655
   End
   Begin VB.TextBox txtTemp 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Index           =   6
      Left            =   6660
      MaxLength       =   15
      TabIndex        =   4
      Top             =   690
      Visible         =   0   'False
      Width           =   2655
   End
   Begin VB.TextBox txtTemp 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Index           =   5
      Left            =   10680
      MaxLength       =   70
      TabIndex        =   10
      Top             =   5040
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.TextBox txtTemp 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Index           =   4
      Left            =   10680
      MaxLength       =   70
      TabIndex        =   9
      Top             =   5700
      Visible         =   0   'False
      Width           =   435
   End
   Begin VB.TextBox txtTemp 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   0
      Left            =   1620
      TabIndex        =   3
      Top             =   720
      Visible         =   0   'False
      Width           =   4335
   End
   Begin VB.TextBox txtTemp 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   2
      Left            =   1620
      TabIndex        =   7
      Top             =   1605
      Visible         =   0   'False
      Width           =   4335
   End
   Begin VB.TextBox txtTemp 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   1
      Left            =   1620
      TabIndex        =   5
      Top             =   1155
      Visible         =   0   'False
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
      Left            =   7500
      TabIndex        =   12
      Top             =   3540
      Width           =   915
   End
   Begin VB.CommandButton cmdPrint 
      Caption         =   "Send"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   0
      Left            =   7500
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   2700
      Width           =   975
   End
   Begin VB.CommandButton cmdPrint 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   1
      Left            =   9420
      Picture         =   "frmRptWhatsApp.frx":058A
      Style           =   1  'Graphical
      TabIndex        =   17
      ToolTipText     =   "Print"
      Top             =   6480
      Visible         =   0   'False
      Width           =   195
   End
   Begin Crystal.CrystalReport CrystalReport1 
      Left            =   8880
      Top             =   7200
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
      Format          =   87425027
      CurrentDate     =   38050
   End
   Begin MSComCtl2.DTPicker DTPicker2 
      Height          =   375
      Left            =   4410
      TabIndex        =   2
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
      Format          =   87425027
      CurrentDate     =   38050
   End
   Begin MSComctlLib.ProgressBar prgBar 
      Height          =   255
      Left            =   240
      TabIndex        =   18
      Top             =   5400
      Visible         =   0   'False
      Width           =   9135
      _ExtentX        =   16113
      _ExtentY        =   450
      _Version        =   393216
      Appearance      =   1
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Outstanding Morer Than late Days    :"
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
      Left            =   660
      TabIndex        =   21
      Top             =   2460
      Visible         =   0   'False
      Width           =   3525
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Mobile No "
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
      Left            =   7380
      TabIndex        =   19
      Top             =   240
      Width           =   1020
   End
   Begin VB.Line Line3 
      BorderColor     =   &H000040C0&
      X1              =   9600
      X2              =   90
      Y1              =   2115
      Y2              =   2115
   End
   Begin VB.Shape Shape2 
      FillColor       =   &H00BDDFD1&
      FillStyle       =   0  'Solid
      Height          =   2145
      Left            =   6900
      Shape           =   4  'Rounded Rectangle
      Top             =   2370
      Width           =   2160
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
      TabIndex        =   14
      Top             =   720
      Visible         =   0   'False
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
      TabIndex        =   16
      Top             =   1605
      Visible         =   0   'False
      Width           =   1170
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Party           :"
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
      TabIndex        =   15
      Top             =   1155
      Visible         =   0   'False
      Width           =   1185
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
      TabIndex        =   13
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
      Height          =   5160
      Left            =   120
      Top             =   60
      Width           =   9540
   End
End
Attribute VB_Name = "frmRpTWhatsApp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim i As Integer
Dim wMillCode As Long, wBrkCode As Long, wPtyCode As Long
Private Sub cmdExit_Click()
Unload Me
End Sub
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
Select Case KeyCode
Case 13
'     SendKeys "{TAB}"
Case 27
    Unload Me
End Select
End Sub
Private Sub Form_Load()
Dim i As Long
Dim cdt As String
Me.Top = 700
Me.Left = Screen.Width / 2 - Me.Width / 2
 '-- Activate Controls
GProcActivateControls Me, False
DTPicker1.Enabled = True
DTPicker2.Enabled = True
'-- Set Values
DTPicker1.Value = Format(gCYSDate, "dd/MM/yyyy")
cdt = Format(Now, "dd/MM/yyyy")

If CDate(cdt) > gCYEDate Then
   DTPicker2.Value = Format(gCYEDate, "dd/MM/yyyy")
Else
   DTPicker2.Value = Format(Now, "dd/MM/yyyy")
End If
Me.Caption = gSelectedMenu
Select Case gSelectedMenu
       Case "Outsading To Broker", "Outsading To Party"
            DTPicker1.Visible = False
            Label1.Visible = False
            txttemp(2).Visible = True
            txttemp(2).Enabled = True
            txttemp(1).Visible = True
            txttemp(1).Enabled = True
            Label5.Visible = True
            If gSelectedMenu = "Outsading To Party" Then
               txttemp(7).Visible = True
               txttemp(7).Enabled = True
            Else
               txttemp(8).Visible = True
               txttemp(8).Enabled = True
            End If
            Label3.Visible = True
End Select
GProcSetInActivateCtrlBackClr Me
End Sub
Private Sub cmdPrint_Click(Index As Integer)
Dim X As String
Dim PtyMbl As String
Dim FileNMWithPath As String
Dim FileNM As String
Dim RsTmp As Recordset
Dim Tattach As String
Dim objCrystal As CRAXDRT.Application
Dim objReport As CRAXDRT.Report
Dim aFile As String
Dim PtyName As String


If txttemp(0) = "" Then wMillCode = 0
If txttemp(1) = "" Then wPtyCode = 0
If txttemp(2) = "" Then wBrkCode = 0
With CrystalReport1
     GProcCrystalRptPreparation CrystalReport1, "Window"
   
     Select Case gSelectedMenu
       Case "Outsading To Broker", "Outsading To Party"
             GProcCollectCommonOutstanding DTPicker1.Value, DTPicker2.Value, "", txttemp(2), txttemp(1), False, False
             gSetupCn.BeginTrans
             gSetupCn.Execute "Delete from tmptbloutstanding where username = '" & gUserName & "' and tmptbloutstanding.OutDueDate >  cdate('" & DTPicker2.Value & "')"
             gSetupCn.CommitTrans
             gCn.BeginTrans
             gCn.Execute "Delete from tmpWhatsAppErr where username = '" & gUserName & "'"
             gCn.CommitTrans
             CrystalReport1.DataFiles(0) = App.Path & "\HIData\setup.mdb"
             If gSelectedMenu = "Outsading To Party" Then
                X = "Select Distinct OutAcName from tmptblOutStanding where username =  '" & gUserName & "'"
             Else
                X = "Select Distinct OutBrokerName from tmptblOutStanding where username =  '" & gUserName & "'"
             End If
             If GProcRstOpen(RsTmp, X, "R", gSetupCn) <= 0 Then
                MsgBox "No Records..."
                Exit Sub
             End If
             gReportPath = App.Path & "\HiReports\"
             aFile = gReportPath & "*.pdf"
             If Len(Dir$(aFile)) > 0 Then
                 Kill aFile
             End If
             i = 0
             RsTmp.MoveFirst
             Do While Not RsTmp.EOF
                With CrystalReport1
                    Set objCrystal = New CRAXDRT.Application
                    If gSelectedMenu = "Outsading To Party" Then
                       Tattach = gReportPath & "rptOutstPartywiseWhatsApp.rpt"
                    Else
                       Tattach = gReportPath & "rptOutstBrokerwiseWhatsApp.rpt"
                    End If
                    Set objReport = objCrystal.OpenReport(Tattach, 1)
                    
                    Dim crTab As CRAXDRT.DatabaseTable
                    For Each crTab In objReport.Database.Tables
                        crTab.Location = gDataPath & "setup.mdb"   '   gDbLocation
                    Next
                    If gSelectedMenu = "Outsading To Party" Then
                       objReport.RecordSelectionFormula = "{tmptbloutstanding.Username}='" & gUserName & "'  and {tmptbloutstanding.OutIsShowOuts} = 0 and {tmptbloutstanding.OutAcName} = '" & RsTmp.Fields(0) & "'  "
                    Else
                       objReport.RecordSelectionFormula = "{tmptbloutstanding.Username}='" & gUserName & "'  and {tmptbloutstanding.OutIsShowOuts} = 0 and {tmptbloutstanding.OutBrokerName} = '" & RsTmp.Fields(0) & "' "
                    End If
                    Dim j As Integer
                    For j = 1 To objReport.FormulaFields.Count
                         Select Case objReport.FormulaFields(j).Name
                         Case "{@FirmNm}"
                             objReport.FormulaFields(j).text = "Trim(" & Chr(39) & UCase(gCName) & Chr(39) & ")"
                         Case "{@wAdd1}"
                             objReport.FormulaFields(j).text = "Trim(" & Chr(39) & gCAdd1 & Chr(39) & ")"
                         Case "{@wAdd2}"
                             objReport.FormulaFields(j).text = "Trim(" & Chr(39) & gCAdd2 & Chr(39) & ")"
                         Case "{@wAdd3}"
                             objReport.FormulaFields(j).text = "Trim(" & Chr(39) & gCAdd3 & Chr(39) & ")"
                         Case "{@wPhNo}"
                             objReport.FormulaFields(j).text = "Trim(" & Chr(39) & gCPhNo & Chr(39) & ")"
                         Case "{@wfrtodt}"
                             objReport.FormulaFields(j).text = "Trim(" & Chr(39) & DTPicker2.Value & Chr(39) & ")"
                        End Select
                   Next j
                   If gSelectedMenu = "Outsading To Party" Then
                      .ReportFileName = gReportPath & "rptOutstPartywiseWhatsApp.rpt"
                   Else
                      .ReportFileName = gReportPath & "rptOutstBrokerwiseWhatsApp.rpt"
                   End If
                   i = i + 1
                   FileNM = "M" + "" & Format(Now, "ddMMyyyyhhmmss") & "" & ".pdf"
                   FileNMWithPath = gReportPath & "" & FileNM & ""
                   If gSelectedMenu = "Outsading To Party" Then
                      If txttemp(7) = "" Then
                          PtyMbl = GProcGetColumnValue("TblMastAccount", "AcName", RsTmp!OutAcName, "S", "AcFaxNo", "S")
                      Else
                          PtyMbl = txttemp(7)
                      End If
                   Else
                      If txttemp(8) = "" Then
                          PtyMbl = GProcGetColumnValue("TblMastAccount", "AcName", RsTmp!OutBrokerName, "S", "AcFaxNo", "S")
                      Else
                          PtyMbl = txttemp(8)
                      End If
                   End If
                   PtyName = RsTmp.Fields(0)

'                   PtyMbl = "9325748022"
'                   PtyMbl = "9922997663"
                   If PtyMbl = "" Then
                      i = i - 1
                   Else
                      ExportReportToPDF objReport, FileNMWithPath, "foo"
                      SentWhatsApp "PDF", FileNM, "", PtyMbl, PtyName
                   End If
                End With
                RsTmp.MoveNext
             Loop
             X = "Select * from tmpWhatsAppErr where username = '" & gUserName & "'"
             If GProcRstOpen(RsTmp, X, "R", gCn) > 0 Then
                GProcWhatsAppError
             End If
             MsgBox CStr(i) + " Whatsapp Msg Sent.."
             aFile = gReportPath & "*.pdf"
             If Len(Dir$(aFile)) > 0 Then
                 Kill aFile
             End If
End Select
End With
End Sub

Private Sub txttemp_GotFocus(Index As Integer)
GProcSelectBox Me.ActiveControl
End Sub
Private Sub txtTemp_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = 13 Then Exit Sub '-- Enter
    Select Case Index
     Case 0 '----Mill Name
'        gClsSearch.SearchMultiField "QryMastaccount", "Account,ACCode", Array("Name", ""), Array(txtTemp(Index).Width, 0, 0), "", Chr(KeyAscii), "AcName", txtTemp(Index).Left + Me.Left, txtTemp(Index).Top + Me.Top + 1000
'        gClsSearch.SearchMultiField "QryMastaccount", "Account,AcCODE,AgName", Array("Party", "AcName", "Group"), Array(txtTemp(Index).Width, 0, 3000), " ", Chr(KeyAscii), "acName", , txtTemp(Index).Left + Me.Left, txtTemp(Index).Top + Me.Top + 1000
'
'        KeyAscii = 0
'
'        If gClsSearch.SearchMultiRetCol(1) <> "" Then
'            txtTemp(Index) = gClsSearch.SearchMultiRetCol(0)
'            wMillCode = gClsSearch.SearchMultiRetCol(1)
'            If Index = 0 Then
'               txtTemp(6).text = GProcGetColumnValue("TblMastAccount", "AcName", txtTemp(0), "S", "AcFaxNo", "S")
'            End If
'            If Index = 1 Then
'               txtTemp(7).text = GProcGetColumnValue("TblMastAccount", "AcName", txtTemp(1), "S", "AcFaxNo", "S")
'            End If
'            If Index = 2 Then
'               txtTemp(8).text = GProcGetColumnValue("TblMastAccount", "AcName", txtTemp(2), "S", "AcFaxNo", "S")
'            End If
'        End If
     Case 1 '-- Party
        gClsSearch.SearchMultiField "QryMastaccount", "Account,AcName,ACCode,AcFaxNo", Array("Broker", "AcName", "", ""), Array(txttemp(Index).Width, 0, 0, 0), " (AgCode=90017 or GpCode=90017)", Chr(KeyAscii), "AcName", txttemp(Index).Left + Me.Left, txttemp(Index).Top + Me.Top + 1000
        KeyAscii = 0
        If gClsSearch.SearchMultiRetCol(1) <> "" Then
            txttemp(Index) = gClsSearch.SearchMultiRetCol(1)
            wPtyCode = gClsSearch.SearchMultiRetCol(2)
            txttemp(7) = gClsSearch.SearchMultiRetCol(3)
        End If
     Case 2 '--Broker
        gClsSearch.SearchMultiField "QryMastaccount", "Account,AcName,ACCode,AcFaxNo", Array("Broker", "AcName", "", ""), Array(txttemp(Index).Width, 0, 0, 0), " (AgCode=90016 or GpCode=90016)", Chr(KeyAscii), "AcName", txttemp(Index).Left + Me.Left, txttemp(Index).Top + Me.Top + 1000
        KeyAscii = 0
        If gClsSearch.SearchMultiRetCol(1) <> "" Then
            txttemp(Index) = gClsSearch.SearchMultiRetCol(1)
            wBrkCode = gClsSearch.SearchMultiRetCol(2)
            txttemp(8) = gClsSearch.SearchMultiRetCol(3)
        End If
    End Select
End Sub
Public Sub delay(PauseTime As Integer)
    Dim start As Single
    start = Timer
   Do While Timer < start + PauseTime
    If (Timer < start) Then 'midnight crossover
        start = start - (86400 + 1)
    End If
    DoEvents   ' Yield to other processes.
    Loop
End Sub


