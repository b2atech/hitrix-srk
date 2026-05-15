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
   ClientWidth     =   11535
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
   ScaleWidth      =   11535
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtFromBrkCd 
      Height          =   375
      Left            =   2880
      TabIndex        =   28
      Top             =   4560
      Visible         =   0   'False
      Width           =   1035
   End
   Begin VB.OptionButton OptBills 
      BackColor       =   &H00EDFDFE&
      Caption         =   "Due Bills"
      BeginProperty Font 
         Name            =   "Verdana"
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
      Left            =   780
      TabIndex        =   27
      Top             =   3960
      Width           =   2235
   End
   Begin VB.OptionButton OptBills 
      BackColor       =   &H00EDFDFE&
      Caption         =   "All Bills"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   240
      Index           =   0
      Left            =   780
      TabIndex        =   26
      Top             =   3480
      Width           =   2235
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
      Index           =   10
      Left            =   1140
      MaxLength       =   40
      TabIndex        =   9
      Top             =   1575
      Visible         =   0   'False
      Width           =   1215
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
      Index           =   9
      Left            =   1140
      MaxLength       =   40
      TabIndex        =   6
      Top             =   1125
      Visible         =   0   'False
      Width           =   1215
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
      Index           =   3
      Left            =   1140
      MaxLength       =   40
      TabIndex        =   3
      Top             =   690
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.TextBox txtLate 
      Alignment       =   1  'Right Justify
      Height          =   375
      Left            =   4440
      TabIndex        =   12
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
      TabIndex        =   24
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
      Left            =   6900
      MaxLength       =   40
      TabIndex        =   11
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
      Left            =   6900
      MaxLength       =   40
      TabIndex        =   8
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
      Left            =   6900
      MaxLength       =   15
      TabIndex        =   5
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
      TabIndex        =   14
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
      TabIndex        =   13
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
      Left            =   2460
      TabIndex        =   4
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
      Left            =   2460
      TabIndex        =   10
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
      Left            =   2460
      TabIndex        =   7
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
      TabIndex        =   16
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
      TabIndex        =   15
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
      TabIndex        =   21
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
      Format          =   105185283
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
      Format          =   105185283
      CurrentDate     =   38050
   End
   Begin MSComctlLib.ProgressBar prgBar 
      Height          =   255
      Left            =   240
      TabIndex        =   22
      Top             =   5400
      Visible         =   0   'False
      Width           =   9135
      _ExtentX        =   16113
      _ExtentY        =   450
      _Version        =   393216
      Appearance      =   1
   End
   Begin Crystal.CrystalReport cryReport2 
      Left            =   10080
      Top             =   720
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      WindowControlBox=   -1  'True
      WindowMaxButton =   -1  'True
      WindowMinButton =   -1  'True
      PrintFileType   =   15
      PrintFileLinesPerPage=   60
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "From Broker Code  :"
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
      Left            =   840
      TabIndex        =   29
      Top             =   4560
      Visible         =   0   'False
      Width           =   1920
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
      TabIndex        =   25
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
      Left            =   7680
      TabIndex        =   23
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
      Caption         =   "&Mill       :"
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
      TabIndex        =   18
      Top             =   720
      Visible         =   0   'False
      Width           =   780
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "&Broker  :"
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
      TabIndex        =   20
      Top             =   1605
      Visible         =   0   'False
      Width           =   750
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Party    :"
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
      TabIndex        =   19
      Top             =   1155
      Visible         =   0   'False
      Width           =   765
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
      TabIndex        =   17
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
     Sendkeys "{TAB}"
Case 27
    Unload Me
End Select
End Sub
Private Sub Form_Load()
Dim i As Long
Dim Cdt As String
Me.Top = 700
Me.Left = Screen.Width / 2 - Me.Width / 2
 '-- Activate Controls
GProcActivateControls Me, False
DTPicker1.Enabled = True
DTPicker2.Enabled = True
'-- Set Values
DTPicker1.Value = Format(gCYSDate, "dd/MM/yyyy")
Cdt = Format(Now, "dd/MM/yyyy")

If CDate(Cdt) > gCYEDate Then
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
            txttemp(9).Visible = True
            txttemp(9).Enabled = True
            txttemp(10).Visible = True
            txttemp(10).Enabled = True
            Label5.Visible = True
            Label3.Visible = True
            If gSelectedMenu = "Outsading To Party" Then
               txttemp(7).Visible = True
               txttemp(7).Enabled = True
               txtFromBrkCd.Visible = True
               txtFromBrkCd.Enabled = True
               Label4.Visible = True
               Label4.Caption = "From Party Code   :"
            Else
               txttemp(8).Visible = True
               txttemp(8).Enabled = True
               txtFromBrkCd.Visible = True
               txtFromBrkCd.Enabled = True
               Label4.Visible = True
            End If
       Case "Bank Detail Mail"
            txttemp(1).Visible = True
            txttemp(1).Enabled = True
            txttemp(9).Visible = True
            txttemp(9).Enabled = True
            txttemp(7).Visible = True
            txttemp(7).Enabled = True
            OptBills(0).Visible = False
            OptBills(1).Visible = False
            Label6.Caption = "Mail ID"
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
Dim PtyMailIdBcc As String
Dim Mfile As String
Dim MailStr As String
Dim MailSub As String
Dim PtyMailId As String
Dim PtyMailIdCc As String
Dim i As Long
Dim FBrkCd As Long
If txttemp(0) = "" Then wMillCode = 0
If txttemp(1) = "" Then wPtyCode = 0
If txttemp(2) = "" Then wBrkCode = 0
With CrystalReport1
     GProcCrystalRptPreparation CrystalReport1, "Window"
     Select Case gSelectedMenu
       Case "Outsading To Broker", "Outsading To Party"
            gCmd.CommandText = "PrcPrepareOutStangingSale"
            gCmd.Parameters.Refresh
            gCmd.Parameters("@To_dt") = Format(DTPicker2.Value, "yyyy/mm/dd")
            gCmd.Parameters("@UCode") = gUserName
            gCmd.Parameters("@IntDbAc") = gLatePayIntAcCodeRecGST
            gCmd.Parameters("@IntDbAcWoGST") = gLPIntAcCodeWithoutGST
            gCmd.Parameters("@IntTDSAc") = gTDSAcCodeRec
            gCmd.Parameters("@IntCrAc") = gLatePayIntAcCodePay
'           If chkBox4.Value = 1 Then gCmd.Parameters("@Is2ndDay") = 1
'           If ChkConsolidate.Value = 1 Then
                gCmd.Parameters("@VFirm") = ""
'           Else
'               gCmd.Parameters("@VFirm") = Trim(gCCode)
'           End If
            If OptBills(1).Value = True Then
               gCmd.Parameters("@DueOn") = "D"
            Else
               gCmd.Parameters("@DueOn") = "B"
            End If
'           If chkBox3.Value = 1 Then gCmd.Parameters("@IsDirectPay") = 1
            gCmd.Parameters("@X_dt1") = Format(DTPicker1.Value, "yyyy/mm/dd")
            gCmd.Parameters("@X_dt2") = Format(DTPicker2.Value, "yyyy/mm/dd")
            If wPtyCode <> 0 Then gCmd.Parameters("@Accode") = wPtyCode
            If wBrkCode <> 0 Then gCmd.Parameters("@Accode_Br") = wBrkCode
            If wMillCode <> 0 Then gCmd.Parameters("@MillCode") = wMillCode
            gCmd.Execute
             
             
            gCn.BeginTrans
            gCn.Execute "Delete from tmpWhatsAppErr where username = '" & gUserName & "'"
            gCn.CommitTrans
            If gSelectedMenu = "Outsading To Party" Then
               X = "Select Distinct PartyName,PartyCode from TmpGentbl where username =  '" & gUserName & "' order by PartyCode"
            Else
               X = "Select Distinct BrokerName,BrokerCode from TmpGentbl where username =  '" & gUserName & "' order by BrokerCode"
            End If
             If GProcRstOpen(RsTmp, X, "R") <= 0 Then
                MsgBox "No Records..."
                Exit Sub
             End If
             gReportPath = App.Path & "\HiReports\"
             aFile = gReportPath & "*.pdf"
             If Len(Dir$(aFile)) > 0 Then
                 Kill aFile
                 DoEvents
                 DoEvents
                 DoEvents
                 DoEvents
                 DoEvents
             End If
             i = 0
             RsTmp.MoveFirst
             Do While Not RsTmp.EOF
                If txtFromBrkCd <> "" Then
Loop1:
                   If gSelectedMenu = "Outsading To Party" Then
                      If Val(Mid(txtFromBrkCd, 2, 10)) > Val(Mid(RsTmp.Fields(1), 2, 10)) Then
                         RsTmp.MoveNext
                         GoTo Loop1
                      End If
                   Else
                      If Val(txtFromBrkCd) > Val(RsTmp.Fields(1)) Then
                         RsTmp.MoveNext
                         GoTo Loop1
                      End If
                   End If
                End If
             
                With CrystalReport1
                    Set objCrystal = New CRAXDRT.Application
                    If gSelectedMenu = "Outsading To Party" Then
                       Tattach = gReportPath & "rptOutstPartywiseWhatsApp.rpt"
                    Else
                       Tattach = gReportPath & "rptOutstBrokerwiseWhatsApp.rpt"
                    End If
                    Set objReport = objCrystal.OpenReport(Tattach, 1)

                    Dim crTab As CRAXDRT.DatabaseTable
                    objReport.Database.LogOnServer "pdsodbc.dll", "Hitrix", "Vishnu", "sa", "123456"
                    If gSelectedMenu = "Outsading To Party" Then
                       objReport.RecordSelectionFormula = "{TmpGentbl.Username}='" & gUserName & "'  and {TmpGentbl.PartyName} = '" & RsTmp.Fields(0) & "' "
                    Else
                       objReport.RecordSelectionFormula = "{TmpGentbl.Username}='" & gUserName & "'   and {TmpGentbl.BrokerName} = '" & RsTmp.Fields(0) & "' "
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
                   FileNM = "" & Format(Now, "ddMMyyyyhhmmss") & "" & ".pdf"
                   FileNMWithPath = gReportPath & "" & FileNM & ""
                  If gSelectedMenu = "Outsading To Party" Then
                      If txttemp(7) = "" Then
                          PtyMbl = GProcGetColumnValue("TblMastAccount", "AcName", RsTmp!PartyName, "S", "AcContNo", "S")
                      Else
                          PtyMbl = txttemp(7)
                      End If
                   Else
                      If txttemp(8) = "" Then
                          PtyMbl = GProcGetColumnValue("TblMastAccount", "AcName", RsTmp!BrokerName, "S", "AcContNo", "S")
                      Else
                          PtyMbl = txttemp(8)
                      End If
                   End If
                   PtyName = RsTmp.Fields(0)
'PtyMbl = "9422049083"
                   If PtyMbl = "" Then
                      i = i - 1
                   Else
                      ExportReportToPDF objReport, FileNMWithPath, "foo"
                      delay (3)
                      SentWhatsApp "PDF", FileNM, "", PtyMbl, PtyName
                      delay (3)
                   End If
                End With
                If i / 25 - Int(i / 25) = 0 Then
                   MsgBox CStr(i) + " Whatsapp " + Str(i) + " Msg Sent.."
                End If
                
                
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
       Case "Bank Detail Mail"
             PtyMailIdBcc = ""
             Mfile = ""
             MailSub = "Important Information"
             
             MailStr = "Dear Sirs, " + Chr(13) + Chr(10) + Chr(13) + Chr(10)
             MailStr = MailStr + "Now a days there is cyber fraud by giving same mail ID and " + Chr(13) + Chr(10)
             MailStr = MailStr + "get payment in their account." + Chr(13) + Chr(10)
             
             MailStr = MailStr + "Kindly note that we do not change the bank account." + Chr(13) + Chr(10)
             MailStr = MailStr + "Please do not consider any such request for our payment." + Chr(13) + Chr(10)
             MailStr = MailStr + "Our account number is same in Tamilnad Mercantile Bank Ltd., " + Chr(13) + Chr(10)
             MailStr = MailStr + "Ichalkaranji Branch." + Chr(13) + Chr(10) + Chr(13) + Chr(10) + Chr(13) + Chr(10)
             
             MailStr = MailStr + "With regards," + Chr(13) + Chr(10)
             MailStr = MailStr + "Arun Kumar Goenka" + Chr(13) + Chr(10)
                                
             If txttemp(7) <> "" Then
                PtyMailId = txttemp(7)
                i = i + 1
                Call SendEmail1("" & Trim(gCMailId) & "", "" & PtyMailId & "", MailSub, MailStr, "" & PtyMailIdCc & "", "" & PtyMailIdBcc & "", Mfile)
             Else
                X = "Select  AcEmail,AcName from tblMastAccount where AcEmail <> '' and AgCode = 90017"
                      
                If GProcRstOpen(RsTmp, X, "R") <= 0 Then
                   MsgBox "No Records..."
                   Exit Sub
                End If
                i = 0
                RsTmp.MoveFirst
                Do While Not RsTmp.EOF
                   PtyMailId = RsTmp.Fields(0)
                   i = i + 1
                   Call SendEmail1("" & Trim(gCMailId) & "", "" & PtyMailId & "", MailSub, MailStr, "" & PtyMailIdCc & "", "" & PtyMailIdBcc & "", Mfile)
                      
                   RsTmp.MoveNext
                Loop
             End If
             MsgBox CStr(i) + " Whatsapp Msg Sent.."
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
            gClsSearch.SearchMultiField "tblMastaccount", "AcName,AcOurCode,AcCity,ACCode", Array("Mill Name", "Code", "City", "Master Code"), Array(8000, 1500, 2000, 0), "(AgCode=90029 ) ", Chr(KeyAscii), "AcName", txttemp(Index).Left + Me.Left, txttemp(Index).Top + Me.Top + 1000
            KeyAscii = 0
            If gClsSearch.SearchMultiRetCol(0) <> "" Then
                txttemp(Index) = gClsSearch.SearchMultiRetCol(0)
                wMillCode = gClsSearch.SearchMultiRetCol(3)
            End If
     Case 1 '-- Party
        gClsSearch.SearchMultiFieldAccMaster Chr(KeyAscii), txttemp(Index).Left + Me.Left, txttemp(Index).Top + Me.Top + 1000
        KeyAscii = 0
        If gClsSearch.SearchMultiRetColAccMaster(0) <> "" Then
            txttemp(Index) = gClsSearch.SearchMultiRetColAccMaster(0)
            wPtyCode = GProcGetColumnValue("tblMastAccount", "AcName", txttemp(Index), "S", "AcCode", "N")
        End If
     Case 2 '--Broker
        gClsSearch.SearchMultiField "tblMastaccount", "AcName,AcOurCode,AcCity,ACCode", Array("Broker Name", "Code", "City", "Master Code"), Array(8000, 1500, 2000, 0), "(AgCode=90016 ) ", Chr(KeyAscii), "AcName", txttemp(Index).Left + Me.Left, txttemp(Index).Top + Me.Top + 1000
        KeyAscii = 0
        If gClsSearch.SearchMultiRetCol(0) <> "" Then
            txttemp(Index) = gClsSearch.SearchMultiRetCol(0)
            wBrkCode = gClsSearch.SearchMultiRetCol(3)
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
Private Sub txtTemp_LostFocus(Index As Integer)
Select Case Index
    Case 9  '-- Party  Code
         txttemp(1) = GProcGetColumnValue("tblMastAccount", "AcOurCode", UCase(txttemp(9)), "S", "AcName", "S")
         wPtyCode = GProcGetColumnValue("tblMastAccount", "AcName", txttemp(1), "S", "AcCode", "N")
'    Case 3  '-- Mill  Code
'         txttemp(3) = GProcGetColumnValue("tblMastAccount", "AcOurCode", UCase(txtCode(0)), "S", "AcName", "S")
'         wMillCode = GProcGetColumnValue("tblMastAccount", "AcName", txttemp(Index), "S", "AcCode", "N")
    Case 10  '-- Broker  Code
         txttemp(2) = GProcGetColumnValue("tblMastAccount", "AcOurCode", UCase(txttemp(10)), "S", "AcName", "S")
         wBrkCode = GProcGetColumnValue("tblMastAccount", "AcName", txttemp(2), "S", "AcCode", "N")
End Select

End Sub
