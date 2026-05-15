VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form frmRptAuditTrial 
   BackColor       =   &H8000000C&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Register"
   ClientHeight    =   6030
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10215
   Icon            =   "frmRptAuditTrial.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6030
   ScaleWidth      =   10215
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
      TabIndex        =   3
      Top             =   1575
      Visible         =   0   'False
      Width           =   3570
   End
   Begin VB.CommandButton cmdCrystal 
      Height          =   480
      Index           =   0
      Left            =   1305
      Picture         =   "frmRptAuditTrial.frx":058A
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   5325
      Width           =   1185
   End
   Begin VB.CommandButton cmdCrystal 
      Height          =   480
      Index           =   1
      Left            =   2730
      Picture         =   "frmRptAuditTrial.frx":09CC
      Style           =   1  'Graphical
      TabIndex        =   8
      ToolTipText     =   "Print"
      Top             =   5325
      Width           =   1185
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H80000018&
      Height          =   4560
      Left            =   5895
      TabIndex        =   4
      Top             =   120
      Width           =   4035
      Begin VB.OptionButton OptAll 
         BackColor       =   &H80000018&
         Caption         =   "All Type Of Transactions"
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
         TabIndex        =   22
         Top             =   240
         Width           =   3735
      End
      Begin VB.OptionButton OptPurchRtn 
         BackColor       =   &H80000018&
         Caption         =   "Purchase Return"
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
         TabIndex        =   21
         Top             =   4260
         Width           =   2415
      End
      Begin VB.OptionButton OptSales_Rtn 
         BackColor       =   &H80000018&
         Caption         =   "Sales Return"
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
         TabIndex        =   20
         Top             =   3900
         Width           =   3735
      End
      Begin VB.OptionButton OptJv 
         BackColor       =   &H80000018&
         Caption         =   "JV"
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
         TabIndex        =   19
         Top             =   3540
         Width           =   3735
      End
      Begin VB.OptionButton OptCreditNote_Purchase 
         BackColor       =   &H80000018&
         Caption         =   "Credit Note (Purchase)"
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
         TabIndex        =   18
         Top             =   3180
         Width           =   3735
      End
      Begin VB.OptionButton OptDebitNote_Purchase 
         BackColor       =   &H80000018&
         Caption         =   "DebitNote ( Purchase )"
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
         TabIndex        =   17
         Top             =   2820
         Width           =   3735
      End
      Begin VB.OptionButton OptCreditNote_Sales 
         BackColor       =   &H80000018&
         Caption         =   "Credit Note (Sales)"
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
         TabIndex        =   16
         Top             =   2460
         Width           =   3735
      End
      Begin VB.OptionButton OptDebitNote_Sales 
         BackColor       =   &H80000018&
         Caption         =   "DebitNote ( Sales )"
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
         TabIndex        =   15
         Top             =   2100
         Width           =   3735
      End
      Begin VB.OptionButton OptReceipt 
         BackColor       =   &H80000018&
         Caption         =   "Receipt"
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
         TabIndex        =   14
         Top             =   1740
         Width           =   1335
      End
      Begin VB.OptionButton OptPayment 
         BackColor       =   &H80000018&
         Caption         =   "Payment"
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
         TabIndex        =   13
         Top             =   1380
         Width           =   1335
      End
      Begin VB.OptionButton OptPurchase 
         BackColor       =   &H80000018&
         Caption         =   "Purchase"
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
         TabIndex        =   6
         Top             =   1020
         Width           =   1455
      End
      Begin VB.OptionButton OptSales 
         BackColor       =   &H80000018&
         Caption         =   "Sales"
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
         TabIndex        =   5
         Top             =   660
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
      Left            =   4155
      TabIndex        =   9
      Top             =   5325
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
      Left            =   135
      TabIndex        =   10
      Top             =   4860
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
      Format          =   123731971
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
      Format          =   123731971
      CurrentDate     =   38050
   End
   Begin Crystal.CrystalReport CrystalReport1 
      Bindings        =   "frmRptAuditTrial.frx":1036
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
      TabIndex        =   12
      Top             =   255
      Width           =   1170
   End
   Begin VB.Shape Shape3 
      FillColor       =   &H00BDDFD1&
      FillStyle       =   0  'Solid
      Height          =   765
      Left            =   1005
      Shape           =   4  'Rounded Rectangle
      Top             =   5220
      Width           =   4785
   End
   Begin VB.Label Label7 
      ForeColor       =   &H000040C0&
      Height          =   330
      Left            =   1710
      TabIndex        =   11
      Top             =   5625
      Width           =   2460
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
      Height          =   4620
      Left            =   45
      Top             =   90
      Width           =   5430
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H00000080&
      FillColor       =   &H00EDFDFE&
      FillStyle       =   0  'Solid
      Height          =   4620
      Left            =   5580
      Top             =   90
      Width           =   4515
   End
End
Attribute VB_Name = "frmRptAuditTrial"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
    Case 13:
       Sendkeys "{TAB}"
    Case 27:
        Unload Me
    End Select
End Sub
Private Sub Form_Load()
Me.Top = 700
Me.Left = Screen.Width / 2 - Me.Width / 2
 '-- Activate Controls to False
'GProcActivateControls Me, False
OptAll.Value = True
dtpFrDt.Enabled = True
dtpToDt.Enabled = True
dtpFrDt.Value = Format(gCYSDate, "dd/MM/yyyy")
If gCYear = Year(Date) Then
   dtpToDt.Value = Format(Date, "dd/MM/yyyy")
Else
dtpToDt.Value = Format(gCYEDate, "dd/MM/yyyy")
End If
Me.Caption = gSelectedMenu
Select Case gSelectedMenu
'==== Outstanding Report
Case "Audit Trail Reports"
End Select
 '-- Set InActivate Ctrl Back Color
GProcSetInActivateCtrlBackClr Me
End Sub
Private Sub cmdExit_Click()
    Unload Me
End Sub
Private Sub cmdCrystal_Click(Index As Integer)
Dim x1 As String, X As String, DtCondStr As String, flnm As String
Dim BarCnt As Double
Dim RecYN As Long
Dim tmp1 As Recordset
Dim excel_app As Excel.Application
Dim workbook As Excel.workbook
Dim sheet_name As String
Dim sheet As Excel.Worksheet
Dim header_range As Excel.Range
Dim a As Long
Dim GPer As Integer
Dim xlApp As Object
Dim xlWB As Object

x1 = ""
With CrystalReport1
    GProcCrystalRptPreparation CrystalReport1, IIf(Index = 0, "Window", "Printer")
    Select Case gSelectedMenu
        Case "Audit Trail Reports"
             flnm = "D:\AuditLog-" + gCCode + ".xlsx"
             FileCopy gReportPath & "\AuditLog_Blank.xlsx", flnm
             Set excel_app = New Excel.Application
             excel_app.Visible = False
             Set workbook = excel_app.Workbooks.Open(filename:=flnm)
            
             gCn.Execute "delete from tmpGenTbl2 where username = '" & gUserName & "'"
              
'             gCn.BeginTrans
'             RecYN = GProcRstOpen(tmp1, "select * from tblvousub_log ", "R", gCn)
'             gCn.CommitTrans
'             BarCnt = RecYN
'
'             gCn.BeginTrans
'             RecYN = GProcRstOpen(tmp1, "select * from tblvoucher_log where VCtrNo = 1 ", "R", gCn)
'             gCn.CommitTrans
'             BarCnt = BarCnt + RecYN
'             DoEvents
'             DoEvents
'             DoEvents
'             DoEvents
'             BarCnt = BarCnt * 1.25
 
'             prgBar.Max = BarCnt
'             prgBar.Value = 0
'             prgBar.Visible = True
              
              
             '-- For Sales Log
             
             If OptSales.Value = True Or OptAll.Value = True Then
                x1 = "select 'Sales',tblSale_Log.vno,tblSale_Log.vtype,tblSale_Log.vyear,tblSale_Log.Vfirm,tblSale_Log.SlBillNo,tblSale_Log.Vdt,tblSale_Log.LogNo,tblSale_Log.LogTp,tblSale_Log.LogDate," _
                     & " tblSale_Log.UserName,tblMastAccount.AcName,tblMastItem.Itname,tblMastAccount_Ml.AcName,SlSubBag ,SlSubWt,SlSubRt,SlSubRtPer,tblSaleSub_Log.SlSubAmt,tblSale_Log.SlTaxableAmt,tblSale_Log.SlTaxRate2,tblSale_Log.SlTaxAmt2," _
                     & " tblSale_Log.SlTaxRate,tblSale_Log.SlTaxAmt,tblSale_Log.SlTaxRate3,tblSale_Log.SlTaxAmt3,tblSale_Log.SlTcsRate,tblSale_Log.SlTcsAmt,tblSale_Log.SlRoff,tblSale_Log.SlBillAmt,tblMastAccount_Br.Acname,'" & gUserName & "'" _
                     & " from tblSale_Log,tblMastAccount,tblMastItem,tblSaleSub_Log,tblMastAccount tblMastAccount_Ml,tblMastAccount tblMastAccount_Br" _
                     & " where  tblSale_Log.SlAcDrCode = tblMastAccount.Accode and tblSale_Log.Vno = tblSaleSub_Log.Vno and tblSale_Log.Vtype = tblSaleSub_Log.Vtype and tblSale_Log.Vfirm = tblSaleSub_Log.Vfirm and tblSale_Log.SlBroker = tblMastAccount_Br.Accode  and" _
                     & " tblSale_Log.VYear = tblSaleSub_Log.VYear and tblSale_Log.UserName = tblSaleSub_Log.UserName and tblSale_Log.LogNo = tblSaleSub_Log.LogNo And tblSaleSub_Log.SlSubItCode = tblMastItem.ItCode And tblMastItem.ItMillCode = tblMastAccount_Ml.AcCode and tblSale_Log.vtype in ('SY','SC','SB') " _
                     & " and tblSale_Log.Vdt>= ('" & Format(dtpFrDt.Value, "yyyy/mm/dd") & "')  and tblSale_Log.Vdt<= ('" & Format(dtpToDt.Value, "yyyy/mm/dd") & "') " _
                     & " and tblSale_Log.vfirm = '" & gCCode & "' and " _
                     & " convert(char(7),tblSale_Log.vno)+tblSale_Log.vtype+convert(char(4),tblSale_Log.vyear)+tblSale_Log.vfirm in " _
                     & " (select convert(char(7),tblSale_Log.vno)+tblSale_Log.vtype+convert(char(4),tblSale_Log.vyear)+tblSale_Log.vfirm " _
                     & " from tblSale_Log group by convert(char(7),tblSale_Log.vno)+tblSale_Log.vtype+convert(char(4),tblSale_Log.vyear)+tblSale_Log.vfirm " _
                     & " having count(convert(char(7),tblSale_Log.vno)+tblSale_Log.vtype+convert(char(4),tblSale_Log.vyear)+tblSale_Log.vfirm) > 1) "
                     
                gCn.Execute "Insert into tmpGenTbl2(Nar13,VNO,VTYPE,VYEAR,vfirm,VBILLNO,VDt,VSRNO,Nar1,Dt1,Nar5,PartyName,Nar3,MillName,Amt1,Amt10,Amt9,Amt4,Amt5,Amt6,Amt7,Amt8,Amt3,Amt2,Amt11,Amt12,Amt13,Amt14,Amt15,Amt16,Nar8,UserName) " & x1
                RecYN = GProcRstOpen(tmp1, "select * from TmpGentbl2 where Nar13 = 'Sales' and TmpGentbl2.UserName ='" & gUserName & "' order by Nar8,vbillno,VSRNO  ", "R", gCn)
                If RecYN > 0 Then
                   Set sheet = workbook.Sheets.Item("Sales")
                   If (sheet Is Nothing) Then
                       MsgBox "Error Sales Sheet"
                   End If
                   sheet.cells(1, 1) = gCName
                   sheet.cells(2, 1) = "GSTIN : " + gCGStin
                   sheet.cells(3, 1) = "Audit Log"
                   sheet.cells(4, 1) = "Sales"
                   sheet.cells(5, 1) = "From : " & dtpFrDt.Value & "  To : " & dtpToDt.Value & ""
                   a = 8
                   With tmp1
                     .MoveFirst
                     Do While Not .EOF
                        sheet.cells(a, 1) = !Nar1
                        sheet.cells(a, 2) = !dt1
                        sheet.cells(a, 3) = !Nar5
                        If !Vtype = "SY" Then
                           sheet.cells(a, 4) = "Trade Sale"
                        ElseIf !Vtype = "SC" Then
                           sheet.cells(a, 4) = "Other Sale"
                        ElseIf !Vtype = "SB" Then
                           sheet.cells(a, 4) = "Brokerage & Commission Sale "
                        End If
                        sheet.cells(a, 5) = !VbillNo
                        sheet.cells(a, 6) = !Vdt
                        sheet.cells(a, 7) = !PartyName
                        sheet.cells(a, 8) = !Nar8
                        
                        sheet.cells(a, 9) = !Amt6      '--Taxable
                        sheet.cells(a, 10) = !Amt7      '--C Rt
                        sheet.cells(a, 11) = !Amt8     '-Camt
                        sheet.cells(a, 12) = !Amt3     'S Rt
                        sheet.cells(a, 13) = !Amt2     'SAmt
                        sheet.cells(a, 14) = !Amt11    ' Irt
                        sheet.cells(a, 15) = !Amt12    '- Iamt
                        sheet.cells(a, 16) = !Amt13    '-Tcs Rt
                        sheet.cells(a, 17) = !Amt14    '-TcsAmt
                        sheet.cells(a, 18) = !Amt15    '-- Roff
                        sheet.cells(a, 19) = !Amt16    '--Bill Amt
                        sheet.cells(a, 20) = !Nar3     '--Count
                        sheet.cells(a, 21) = !MillName
                        sheet.cells(a, 22) = !Amt1     '--Bag
                        sheet.cells(a, 23) = !Amt10    '-Wt
                        sheet.cells(a, 24) = !Amt9     '--Rate
                        sheet.cells(a, 25) = !Amt4     '-Per
                        sheet.cells(a, 26) = !Amt5     '--Amt
                        a = a + 1
'                        prgBar.Value = prgBar.Value + 1
                       .MoveNext
                     Loop
                   End With
                End If
             End If
                
             '-- For Purchase Log
             
             If OptPurchase.Value = True Or OptAll.Value = True Then

                x1 = "select 'Purchase',tblPurch_Log.vno,tblPurch_Log.vtype,tblPurch_Log.vyear,tblPurch_Log.Vfirm,tblPurch_Log.PurBillNo,tblPurch_Log.Vdt,tblPurch_Log.LogNo,tblPurch_Log.LogTp,tblPurch_Log.LogDate," _
                     & " tblPurch_Log.UserName,tblMastAccount.AcName,tblMastItem.Itname,tblMastAccount_Ml.AcName,PurSubBag ,PurSubWt,PurSubRt,PurSubRtPer,tblPurchSub_Log.PurSubAmt,tblPurch_Log.PurTaxableAmt,tblPurch_Log.PurTaxRate,tblPurch_Log.PurTaxAmt," _
                     & " tblPurch_Log.PurTaxRate2,tblPurch_Log.PurTaxAmt2,tblPurch_Log.PurTaxRate3,tblPurch_Log.PurTaxAmt3,tblPurch_Log.PurTcsRate,tblPurch_Log.PurTcsAmt,tblPurch_Log.PurRoff,tblPurch_Log.PurBillAmt,tblMastAccount_Br.Acname,tblPurch_Log.PurBillDt,PurSubItSrNo,'" & gUserName & "'" _
                     & " from tblPurch_Log,tblMastAccount,tblMastItem,tblPurchSub_Log,tblMastAccount tblMastAccount_Ml,tblMastAccount tblMastAccount_Br" _
                     & " where  tblPurch_Log.PurAcCrCode = tblMastAccount.Accode and tblPurch_Log.Vno = tblPurchSub_Log.Vno and tblPurch_Log.Vtype = tblPurchSub_Log.Vtype and tblPurch_Log.Vfirm = tblPurchSub_Log.Vfirm and tblPurch_Log.PurBroker = tblMastAccount_Br.Accode  and" _
                     & " tblPurch_Log.VYear = tblPurchSub_Log.VYear and tblPurch_Log.UserName = tblPurchSub_Log.UserName and tblPurch_Log.LogNo = tblPurchSub_Log.LogNo And tblPurchSub_Log.PurSubItCode = tblMastItem.ItCode And tblMastItem.ItMillCode = tblMastAccount_Ml.AcCode and tblPurch_Log.vtype in ('PY','PO') " _
                     & " and tblPurch_Log.Vdt>= ('" & Format(dtpFrDt.Value, "yyyy/mm/dd") & "')  and tblPurch_Log.Vdt<= ('" & Format(dtpToDt.Value, "yyyy/mm/dd") & "') " _
                     & " and tblPurch_Log.vfirm = '" & gCCode & "' and " _
                     & " convert(char(7),tblPurch_Log.vno)+tblPurch_Log.vtype+convert(char(4),tblPurch_Log.vyear)+tblPurch_Log.vfirm in " _
                     & " (select convert(char(7),tblPurch_Log.vno)+tblPurch_Log.vtype+convert(char(4),tblPurch_Log.vyear)+tblPurch_Log.vfirm " _
                     & " from tblPurch_Log group by convert(char(7),tblPurch_Log.vno)+tblPurch_Log.vtype+convert(char(4),tblPurch_Log.vyear)+tblPurch_Log.vfirm " _
                     & " having count(convert(char(7),tblPurch_Log.vno)+tblPurch_Log.vtype+convert(char(4),tblPurch_Log.vyear)+tblPurch_Log.vfirm) > 1) "
                     
                     
                
                
                gCn.Execute "Insert into tmpGenTbl2(Nar13,VNO,VTYPE,VYEAR,vfirm,VBILLNO,VDt,VSRNO,Nar1,Dt1,Nar5,PartyName,Nar3,MillName,Amt1,Amt10,Amt9,Amt4,Amt5,Amt6,Amt7,Amt8,Amt3,Amt2,Amt13,Amt12,Amt11,Amt14,Amt15,Amt16,Nar8,Dt3,Amt17,UserName) " & x1

                RecYN = GProcRstOpen(tmp1, "select * from TmpGentbl2 where Nar13 = 'Purchase' and TmpGentbl2.UserName ='" & gUserName & "' order by Nar8,vdt,vno,amt18,VSRNO ,Amt17 ", "R", gCn)
                If RecYN > 0 Then
                   Set sheet = workbook.Sheets.Item("Purchase")
                   If (sheet Is Nothing) Then
                       MsgBox "Error Purchase Sheet"
                   End If
                   sheet.cells(1, 1) = gCName
                   sheet.cells(2, 1) = "GSTIN : " + gCGStin
                   sheet.cells(3, 1) = "Audit Log"
                   sheet.cells(4, 1) = "Purchase"
                   sheet.cells(5, 1) = "From : " & dtpFrDt.Value & "  To : " & dtpToDt.Value & ""
                   a = 8
                   With tmp1
                     .MoveFirst
                     Do While Not .EOF
                        sheet.cells(a, 1) = !Nar1
                        sheet.cells(a, 2) = !dt1
                        sheet.cells(a, 3) = !Nar5
                        If !Vtype = "PY" Then
                           sheet.cells(a, 4) = "Trade Purchase"
                        ElseIf !Vtype = "PO" Then
                           sheet.cells(a, 4) = "Other Purchase"
                        End If
                        sheet.cells(a, 5) = !Vdt
                        sheet.cells(a, 6) = !VbillNo
                        sheet.cells(a, 7) = !Dt3
                        sheet.cells(a, 8) = !PartyName
                        sheet.cells(a, 9) = !Nar8        '--Broker
                        If !amt17 = 1 Then
                            sheet.cells(a, 10) = !Amt6      '--Taxable
                            sheet.cells(a, 11) = !Amt7      '--C Rt
                            sheet.cells(a, 12) = !Amt8     '-Camt
                            sheet.cells(a, 13) = !Amt3     'S Rt
                            sheet.cells(a, 14) = !Amt2     'SAmt
                            sheet.cells(a, 15) = !Amt13    ' Irt
                            sheet.cells(a, 16) = !Amt12    '- Iamt
                            sheet.cells(a, 17) = !Amt11    '-Tcs Rt
                            sheet.cells(a, 18) = !Amt14    '-TcsAmt
                            sheet.cells(a, 19) = !Amt15    '-- Roff
                            sheet.cells(a, 20) = !Amt16    '--Bill Amt
                        End If
                        sheet.cells(a, 21) = !Nar3     '--Count
                        sheet.cells(a, 22) = !MillName
                        sheet.cells(a, 23) = !Amt1     '--Bag
                        sheet.cells(a, 24) = !Amt10    '-Wt
                        sheet.cells(a, 25) = !Amt9     '--Rate
                        sheet.cells(a, 26) = !Amt4     '-Per
                        sheet.cells(a, 27) = !Amt5     '--Amt
                        
                        a = a + 1
'                        prgBar.Value = prgBar.Value + 1
                       .MoveNext
                     Loop
                   End With
                End If
             End If

             '-- For Sales Return Log
             If OptSales_Rtn.Value = True Or OptAll.Value = True Then
                x1 = "select 'Sales Rtn',tblSale_Log.vno,tblSale_Log.vtype,tblSale_Log.vyear,tblSale_Log.Vfirm,tblSale_Log.SlBillNo,tblSale_Log.Vdt,tblSale_Log.LogNo,tblSale_Log.LogTp,tblSale_Log.LogDate," _
                     & " tblSale_Log.UserName,tblMastAccount.AcName,tblMastItem.Itname,tblMastAccount_Ml.AcName,SlSubBag ,SlSubWt,SlSubRt,SlSubRtPer,tblSaleSub_Log.SlSubAmt,tblSale_Log.SlTaxableAmt,tblSale_Log.SlTaxRate2,tblSale_Log.SlTaxAmt2," _
                     & " tblSale_Log.SlTaxRate,tblSale_Log.SlTaxAmt,tblSale_Log.SlTaxRate3,tblSale_Log.SlTaxAmt3,tblSale_Log.SlTcsRate,tblSale_Log.SlTcsAmt,tblSale_Log.SlRoff,tblSale_Log.SlBillAmt,tblMastAccount_Br.Acname,tblSale_Log.SlBillDt,'" & gUserName & "'" _
                     & " from tblSale_Log,tblMastAccount,tblMastItem,tblSaleSub_Log,tblMastAccount tblMastAccount_Ml,tblMastAccount tblMastAccount_Br" _
                     & " where  tblSale_Log.SlAcDrCode = tblMastAccount.Accode and tblSale_Log.Vno = tblSaleSub_Log.Vno and tblSale_Log.Vtype = tblSaleSub_Log.Vtype and tblSale_Log.Vfirm = tblSaleSub_Log.Vfirm and tblSale_Log.SlBroker = tblMastAccount_Br.Accode  and" _
                     & " tblSale_Log.VYear = tblSaleSub_Log.VYear and tblSale_Log.UserName = tblSaleSub_Log.UserName and tblSale_Log.LogNo = tblSaleSub_Log.LogNo And tblSaleSub_Log.SlSubItCode = tblMastItem.ItCode And tblMastItem.ItMillCode = tblMastAccount_Ml.AcCode and tblSale_Log.vtype = 'RY' " _
                     & " and tblSale_Log.Vdt>= ('" & Format(dtpFrDt.Value, "yyyy/mm/dd") & "')  and tblSale_Log.Vdt<= ('" & Format(dtpToDt.Value, "yyyy/mm/dd") & "') " _
                     & " and tblSale_Log.vfirm = '" & gCCode & "' and " _
                     & " convert(char(7),tblSale_Log.vno)+tblSale_Log.vtype+convert(char(4),tblSale_Log.vyear)+tblSale_Log.vfirm in " _
                     & " (select convert(char(7),tblSale_Log.vno)+tblSale_Log.vtype+convert(char(4),tblSale_Log.vyear)+tblSale_Log.vfirm " _
                     & " from tblSale_Log group by convert(char(7),tblSale_Log.vno)+tblSale_Log.vtype+convert(char(4),tblSale_Log.vyear)+tblSale_Log.vfirm " _
                     & " having count(convert(char(7),tblSale_Log.vno)+tblSale_Log.vtype+convert(char(4),tblSale_Log.vyear)+tblSale_Log.vfirm) > 1) "
                     
                gCn.Execute "Insert into tmpGenTbl2(Nar13,VNO,VTYPE,VYEAR,vfirm,VBILLNO,VDt,VSRNO,Nar1,Dt1,Nar5,PartyName,Nar3,MillName,Amt1,Amt10,Amt9,Amt4,Amt5,Amt6,Amt7,Amt8,Amt3,Amt2,Amt11,Amt12,Amt13,Amt14,Amt15,Amt16,Nar8,Dt2,UserName) " & x1
                RecYN = GProcRstOpen(tmp1, "select * from TmpGentbl2 where Nar13 = 'Sales Rtn' and TmpGentbl2.UserName ='" & gUserName & "' order by Nar8,vbillno,VSRNO  ", "R", gCn)
                If RecYN > 0 Then
                   Set sheet = workbook.Sheets.Item("Sales_Rtn")
                   If (sheet Is Nothing) Then
                       MsgBox "Error Sales Sheet"
                   End If
                   sheet.cells(1, 1) = gCName
                   sheet.cells(2, 1) = "GSTIN : " + gCGStin
                   sheet.cells(3, 1) = "Audit Log"
                   sheet.cells(4, 1) = "Sales Return"
                   sheet.cells(5, 1) = "From : " & dtpFrDt.Value & "  To : " & dtpToDt.Value & ""
                   a = 8
                   With tmp1
                     .MoveFirst
                     Do While Not .EOF
                        sheet.cells(a, 1) = !Nar1
                        sheet.cells(a, 2) = !dt1
                        sheet.cells(a, 3) = !Nar5
                        sheet.cells(a, 4) = "Sale Rtn"
                        sheet.cells(a, 5) = !Vno
                        sheet.cells(a, 6) = !Vdt
                        sheet.cells(a, 7) = !VbillNo
                        sheet.cells(a, 8) = !dt2
                        sheet.cells(a, 9) = !PartyName
                        sheet.cells(a, 10) = !Nar8
                        
                        sheet.cells(a, 11) = !Amt6      '--Taxable
                        sheet.cells(a, 12) = !Amt7      '--C Rt
                        sheet.cells(a, 13) = !Amt8     '-Camt
                        sheet.cells(a, 14) = !Amt3     'S Rt
                        sheet.cells(a, 15) = !Amt2     'SAmt
                        sheet.cells(a, 16) = !Amt11    ' Irt
                        sheet.cells(a, 17) = !Amt12    '- Iamt
                        sheet.cells(a, 18) = !Amt13    '-Tcs Rt
                        sheet.cells(a, 19) = !Amt14    '-TcsAmt
                        sheet.cells(a, 20) = !Amt15    '-- Roff
                        sheet.cells(a, 21) = !Amt16    '--Bill Amt
                        sheet.cells(a, 22) = !Nar3     '--Count
                        sheet.cells(a, 23) = !MillName
                        sheet.cells(a, 24) = !Amt1     '--Bag
                        sheet.cells(a, 25) = !Amt10    '-Wt
                        sheet.cells(a, 26) = !Amt9     '--Rate
                        sheet.cells(a, 27) = !Amt4     '-Per
                        sheet.cells(a, 28) = !Amt5     '--Amt
                        a = a + 1
'                        prgBar.Value = prgBar.Value + 1
                       .MoveNext
                     Loop
                   End With
                End If
             End If
'
'             '-- For Purchase Return Log

             If OptPurchRtn.Value = True Or OptAll.Value = True Then
                x1 = "select 'Purchase Rtn',tblPurch_Log.vno,tblPurch_Log.vtype,tblPurch_Log.vyear,tblPurch_Log.Vfirm,tblPurch_Log.PurBillNo,tblPurch_Log.Vdt,tblPurch_Log.LogNo,tblPurch_Log.LogTp,tblPurch_Log.LogDate," _
                     & " tblPurch_Log.UserName,tblMastAccount.AcName,tblMastItem.Itname,tblMastAccount_Ml.AcName,PurSubBag ,PurSubWt,PurSubRt,PurSubRtPer,tblPurchSub_Log.PurSubAmt,tblPurch_Log.PurTaxableAmt,tblPurch_Log.PurTaxRate,tblPurch_Log.PurTaxAmt," _
                     & " tblPurch_Log.PurTaxRate2,tblPurch_Log.PurTaxAmt2,tblPurch_Log.PurTaxRate3,tblPurch_Log.PurTaxAmt3,tblPurch_Log.PurTcsRate,tblPurch_Log.PurTcsAmt,tblPurch_Log.PurRoff,tblPurch_Log.PurBillAmt,tblMastAccount_Br.Acname,tblPurch_Log.PurBillDt,PurSubItSrNo,'" & gUserName & "'" _
                     & " from tblPurch_Log,tblMastAccount,tblMastItem,tblPurchSub_Log,tblMastAccount tblMastAccount_Ml,tblMastAccount tblMastAccount_Br" _
                     & " where  tblPurch_Log.PurAcCrCode = tblMastAccount.Accode and tblPurch_Log.Vno = tblPurchSub_Log.Vno and tblPurch_Log.Vtype = tblPurchSub_Log.Vtype and tblPurch_Log.Vfirm = tblPurchSub_Log.Vfirm and tblPurch_Log.PurBroker = tblMastAccount_Br.Accode  and" _
                     & " tblPurch_Log.VYear = tblPurchSub_Log.VYear and tblPurch_Log.UserName = tblPurchSub_Log.UserName and tblPurch_Log.LogNo = tblPurchSub_Log.LogNo And tblPurchSub_Log.PurSubItCode = tblMastItem.ItCode And tblMastItem.ItMillCode = tblMastAccount_Ml.AcCode and tblPurch_Log.vtype = 'VY' " _
                     & " and tblPurch_Log.Vdt>= ('" & Format(dtpFrDt.Value, "yyyy/mm/dd") & "')  and tblPurch_Log.Vdt<= ('" & Format(dtpToDt.Value, "yyyy/mm/dd") & "') " _
                     & " and tblPurch_Log.vfirm = '" & gCCode & "' and " _
                     & " convert(char(7),tblPurch_Log.vno)+tblPurch_Log.vtype+convert(char(4),tblPurch_Log.vyear)+tblPurch_Log.vfirm in " _
                     & " (select convert(char(7),tblPurch_Log.vno)+tblPurch_Log.vtype+convert(char(4),tblPurch_Log.vyear)+tblPurch_Log.vfirm " _
                     & " from tblPurch_Log group by convert(char(7),tblPurch_Log.vno)+tblPurch_Log.vtype+convert(char(4),tblPurch_Log.vyear)+tblPurch_Log.vfirm " _
                     & " having count(convert(char(7),tblPurch_Log.vno)+tblPurch_Log.vtype+convert(char(4),tblPurch_Log.vyear)+tblPurch_Log.vfirm) > 1) "
                gCn.Execute "Insert into tmpGenTbl2(Nar13,VNO,VTYPE,VYEAR,vfirm,VBILLNO,VDt,VSRNO,Nar1,Dt1,Nar5,PartyName,Nar3,MillName,Amt1,Amt10,Amt9,Amt4,Amt5,Amt6,Amt7,Amt8,Amt3,Amt2,Amt13,Amt12,Amt11,Amt14,Amt15,Amt16,Nar8,Dt3,Amt17,UserName) " & x1

                RecYN = GProcRstOpen(tmp1, "select * from TmpGentbl2 where Nar13 = 'Purchase Rtn' and TmpGentbl2.UserName ='" & gUserName & "' order by Nar8,vno,amt18,VSRNO  ", "R", gCn)
                If RecYN > 0 Then
                   Set sheet = workbook.Sheets.Item("Purch_Rtn")
                   If (sheet Is Nothing) Then
                       MsgBox "Error Purch_Rtn Sheet"
                   End If
                   sheet.cells(1, 1) = gCName
                   sheet.cells(2, 1) = "GSTIN : " + gCGStin
                   sheet.cells(3, 1) = "Audit Log"
                   sheet.cells(4, 1) = "Purchase Return"
                   sheet.cells(5, 1) = "From : " & dtpFrDt.Value & "  To : " & dtpToDt.Value & ""
                   a = 8
                   With tmp1
                     .MoveFirst
                     Do While Not .EOF
                        sheet.cells(a, 1) = !Nar1
                        sheet.cells(a, 2) = !dt1
                        sheet.cells(a, 3) = !Nar5
                        sheet.cells(a, 4) = "Purchase Rtn."
                        sheet.cells(a, 5) = !Vno
                        sheet.cells(a, 6) = !Vdt
                        sheet.cells(a, 7) = !VbillNo
                        sheet.cells(a, 8) = !Dt3
                        sheet.cells(a, 9) = !PartyName
                        sheet.cells(a, 10) = !Nar8        '--Broker
                        If !amt17 = 1 Then
                            sheet.cells(a, 11) = !Amt6      '--Taxable
                            sheet.cells(a, 12) = !Amt7      '--C Rt
                            sheet.cells(a, 13) = !Amt8     '-Camt
                            sheet.cells(a, 14) = !Amt3     'S Rt
                            sheet.cells(a, 15) = !Amt2     'SAmt
                            sheet.cells(a, 16) = !Amt13    ' Irt
                            sheet.cells(a, 17) = !Amt12    '- Iamt
                            sheet.cells(a, 18) = !Amt11    '-Tcs Rt
                            sheet.cells(a, 19) = !Amt14    '-TcsAmt
                            sheet.cells(a, 20) = !Amt15    '-- Roff
                            sheet.cells(a, 21) = !Amt16    '--Bill Amt
                        End If
                        sheet.cells(a, 22) = !Nar3     '--Count
                        sheet.cells(a, 23) = !MillName
                        sheet.cells(a, 24) = !Amt1     '--Bag
                        sheet.cells(a, 25) = !Amt10    '-Wt
                        sheet.cells(a, 26) = !Amt9     '--Rate
                        sheet.cells(a, 27) = !Amt4     '-Per
                        sheet.cells(a, 28) = !Amt5     '--Amt
                        
                        a = a + 1
'                        prgBar.Value = prgBar.Value + 1
                       .MoveNext
                     Loop
                   End With
                End If
             End If

             '-- For Payment Log

             If OptPayment.Value = True Or OptAll.Value = True Then
                x1 = "select 'Payment',tblVoucher_Log.Vfirm,tblVoucher_Log.vno,tblVoucher_Log.vtype,tblVoucher_Log.vyear,tblVoucher_Log.Vdt,tblVoucher_Log.LogNo,tblVoucher_Log.LogTp,tblVoucher_Log.LogDate," _
                     & " tblVoucher_Log.UserName,tblMastAccount.AcName,tblMastAccount_OAc.AcName,abs(tblVoucher_Log.Vamt),tblVoucher_Log.VRefTp,tblVoucher_Log.VRefNo,tblVoucher_Log.VRefBank,'" & gUserName & "'" _
                     & " from tblVoucher_Log,tblMastAccount,tblMastAccount tblMastAccount_OAc" _
                     & " where  tblVoucher_Log.VDrAcCode = tblMastAccount.Accode And tblVoucher_Log.VCrAcCode = tblMastAccount_OAC.AcCode and tblVoucher_Log.vtype in ('BP','CP') " _
                     & " and tblVoucher_Log.Vdt>= ('" & Format(dtpFrDt.Value, "yyyy/mm/dd") & "')  and tblVoucher_Log.Vdt<= ('" & Format(dtpToDt.Value, "yyyy/mm/dd") & "') and tblVoucher_Log.vfirm = '" & gCCode & "'  and " _
                     & " convert(char(7),tblVoucher_Log.vno)+tblVoucher_Log.vtype+convert(char(4),tblVoucher_Log.vyear)+tblVoucher_Log.vfirm in " _
                     & " (select convert(char(7),tblVoucher_Log.vno)+tblVoucher_Log.vtype+convert(char(4),tblVoucher_Log.vyear)+tblVoucher_Log.vfirm " _
                     & " from tblVoucher_Log where tblVoucher_Log.VCtrNo = 1  group by convert(char(7),tblVoucher_Log.vno)+tblVoucher_Log.vtype+convert(char(4),tblVoucher_Log.vyear)+tblVoucher_Log.vfirm " _
                     & " having count(convert(char(7),tblVoucher_Log.vno)+tblVoucher_Log.vtype+convert(char(4),tblVoucher_Log.vyear)+tblVoucher_Log.vfirm) > 1) "
                gCn.Execute "Insert into tmpGenTbl2(Nar13,Vfirm,VNO,VTYPE,VYEAR,VDt,VSRNO,Nar1,Nar7,Nar2,PartyName,MillName,Amt1,Nar8,Nar9,Nar10,UserName) " & x1
                RecYN = GProcRstOpen(tmp1, "select * from TmpGentbl2 where Nar13 = 'Payment' and TmpGentbl2.UserName ='" & gUserName & "' order by Vtype,vno,VSRNO  ", "R", gCn)
                If RecYN > 0 Then
                   Set sheet = workbook.Sheets.Item("Payment")
                   If (sheet Is Nothing) Then
                       MsgBox "Error Payment Sheet"
                   End If
                   sheet.cells(1, 1) = gCName
                   sheet.cells(2, 1) = "GSTIN : " + gCGStin
                   sheet.cells(3, 1) = "Audit Log"
                   sheet.cells(4, 1) = "Payment"
                   sheet.cells(5, 1) = "From : " & dtpFrDt.Value & "  To : " & dtpToDt.Value & ""
                   a = 8
                   With tmp1
                     .MoveFirst
                     Do While Not .EOF
                        sheet.cells(a, 1) = !Nar1
                        sheet.cells(a, 2) = !Nar7
                        sheet.cells(a, 3) = !Nar2
                        If !Vtype = "BP" Then
                           sheet.cells(a, 4) = "Bank Pay."
                        Else
                           sheet.cells(a, 4) = "Cash Pay."
                        End If
                        sheet.cells(a, 5) = !Vno
                        sheet.cells(a, 6) = !Vdt
                        sheet.cells(a, 7) = !PartyName
                        sheet.cells(a, 8) = !MillName
                        sheet.cells(a, 9) = !Amt1
                        sheet.cells(a, 10) = !Nar8
                        sheet.cells(a, 11) = !Nar9
                        sheet.cells(a, 12) = !Nar10
                        a = a + 1
'                        prgBar.Value = prgBar.Value + 1
                       .MoveNext
                     Loop
                   End With
                End If
             End If

             '-- For Receipt Log

             If OptReceipt.Value = True Or OptAll.Value = True Then
                x1 = "select 'Receipt',tblVoucher_Log.Vfirm,tblVoucher_Log.vno,tblVoucher_Log.vtype,tblVoucher_Log.vyear,tblVoucher_Log.Vdt,tblVoucher_Log.LogNo,tblVoucher_Log.LogTp,tblVoucher_Log.LogDate," _
                     & " tblVoucher_Log.UserName,tblMastAccount.AcName,tblMastAccount_OAc.AcName,abs(tblVoucher_Log.Vamt),tblVoucher_Log.VRefTp,tblVoucher_Log.VRefNo,tblVoucher_Log.VRefBank,tblVoucher_Log.VBillNo,tblVoucher_Log.VBillDate,'" & gUserName & "'" _
                     & " from tblVoucher_Log,tblMastAccount,tblMastAccount tblMastAccount_OAc" _
                     & " where  tblVoucher_Log.VCrAcCode = tblMastAccount.Accode And tblVoucher_Log.VDrAcCode = tblMastAccount_OAC.AcCode and tblVoucher_Log.vtype in ('BR','CR') and tblVoucher_Log.VCtrNo = 1" _
                     & " and tblVoucher_Log.Vdt>= ('" & Format(dtpFrDt.Value, "yyyy/mm/dd") & "') and tblVoucher_Log.Vdt<= ('" & Format(dtpToDt.Value, "yyyy/mm/dd") & "')and tblVoucher_Log.vfirm = '" & gCCode & "' and " _
                     & " convert(char(7),tblVoucher_Log.vno)+tblVoucher_Log.vtype+convert(char(4),tblVoucher_Log.vyear)+tblVoucher_Log.vfirm in " _
                     & " (select convert(char(7),tblVoucher_Log.vno)+tblVoucher_Log.vtype+convert(char(4),tblVoucher_Log.vyear)+tblVoucher_Log.vfirm " _
                     & " from tblVoucher_Log where tblVoucher_Log.VCtrNo = 1  group by convert(char(7),tblVoucher_Log.vno)+tblVoucher_Log.vtype+convert(char(4),tblVoucher_Log.vyear)+tblVoucher_Log.vfirm " _
                     & " having count(convert(char(7),tblVoucher_Log.vno)+tblVoucher_Log.vtype+convert(char(4),tblVoucher_Log.vyear)+tblVoucher_Log.vfirm) > 1) "

                gCn.Execute "Insert into tmpGenTbl2(Nar13,Vfirm,VNO,VTYPE,VYEAR,VDt,VSRNO,Nar1,Nar7,Nar2,PartyName,MillName,Amt1,Nar8,Nar9,Nar10,Nar3,Dt1,UserName) " & x1
                RecYN = GProcRstOpen(tmp1, "select * from TmpGentbl2 where Nar13 = 'Receipt' and TmpGentbl2.UserName ='" & gUserName & "' order by Vtype,vno,VSRNO  ", "R", gCn)
                If RecYN > 0 Then
                   Set sheet = workbook.Sheets.Item("Receipt")
                   If (sheet Is Nothing) Then
                       MsgBox "Error Recript Sheet"
                   End If
                   sheet.cells(1, 1) = gCName
                   sheet.cells(2, 1) = "GSTIN : " + gCGStin
                   sheet.cells(3, 1) = "Audit Log"
                   sheet.cells(4, 1) = "Receipt"
                   sheet.cells(5, 1) = "From : " & dtpFrDt.Value & "  To : " & dtpToDt.Value & ""
                   a = 8
                   With tmp1
                     .MoveFirst
                     Do While Not .EOF
                        sheet.cells(a, 1) = !Nar1
                        sheet.cells(a, 2) = !Nar7
                        sheet.cells(a, 3) = !Nar2
                        If !Vtype = "BR" Then
                           sheet.cells(a, 4) = "Bank"
                        Else
                           sheet.cells(a, 4) = "Cash"
                        End If
                        sheet.cells(a, 5) = !Vno
                        sheet.cells(a, 6) = !Vdt
                        sheet.cells(a, 7) = !PartyName
                        sheet.cells(a, 8) = !MillName
                        sheet.cells(a, 9) = !Amt1
                        sheet.cells(a, 10) = !Nar8
                        sheet.cells(a, 11) = !Nar9
                        sheet.cells(a, 12) = !Nar10
                        sheet.cells(a, 13) = !Nar3
                        sheet.cells(a, 14) = !dt1
                        a = a + 1
'                        prgBar.Value = prgBar.Value + 1
                       .MoveNext
                     Loop
                   End With
                End If
             End If

             '-- For Debit Note Sales Log

             If OptDebitNote_Sales.Value = True Or OptAll.Value = True Then
                
                x1 = "select 'DBN-Sales',tblIntSale_log.Vfirm,tblIntSale_log.vno,tblIntSale_log.vtype,tblIntSale_log.vyear,tblIntSale_log.Vdt,tblIntSale_log.LogNo,tblIntSale_log.LogTp,tblIntSale_log.LogDate," _
                     & " tblIntSale_log.UserName,tblMastAccount.AcName,tblMastAccount_OAc.AcName,tblMastAccount_Br.AcName,tblIntSale_log.Vamt,tblIntSale_log.Vnar1,tblIntSale_log.VNar2,tblIntSale_log.VBillDate,tblIntSale_log.CrDrNoteNo,tblIntSale_log.VBillNo, " _
                     & " tblIntSale_log.Interest,tblIntSale_log.CgstRt,tblIntSale_log.CgstAmt,tblIntSale_log.SgstRt,tblIntSale_log.SgstAmt,tblIntSale_log.IgstRt,tblIntSale_log.IgstAmt,    '" & gUserName & "'" _
                     & " from tblIntSale_log,tblMastAccount,tblMastAccount tblMastAccount_OAc,tblMastAccount tblMastAccount_Br " _
                     & " where  tblIntSale_log.VDrAcCode = tblMastAccount.Accode And tblIntSale_log.VCrAcCode = tblMastAccount_OAC.AcCode And tblIntSale_log.Broker = tblMastAccount_Br.AcCode  and tblIntSale_log.vtype in ('SI','SN') and tblIntSale_log.VCtrNo = 1" _
                     & " and tblIntSale_log.Vdt>= ('" & Format(dtpFrDt.Value, "yyyy/mm/dd") & "') and tblIntSale_log.Vdt<= ('" & Format(dtpToDt.Value, "yyyy/mm/dd") & "')and tblIntSale_log.vfirm = '" & gCCode & "' and " _
                     & " convert(char(7),tblIntSale_log.vno)+tblIntSale_log.vtype+convert(char(4),tblIntSale_log.vyear)+tblIntSale_log.vfirm in " _
                     & " (select convert(char(7),tblIntSale_log.vno)+tblIntSale_log.vtype+convert(char(4),tblIntSale_log.vyear)+tblIntSale_log.vfirm " _
                     & " from tblIntSale_log where tblIntSale_log.VCtrNo = 1  group by convert(char(7),tblIntSale_log.vno)+tblIntSale_log.vtype+convert(char(4),tblIntSale_log.vyear)+tblIntSale_log.vfirm " _
                     & " having count(convert(char(7),tblIntSale_log.vno)+tblIntSale_log.vtype+convert(char(4),tblIntSale_log.vyear)+tblIntSale_log.vfirm) > 1) "
                
                gCn.Execute "Insert into tmpGenTbl2(Nar13,Vfirm,VNO,VTYPE,VYEAR,VDt,VSRNO,Nar1,Nar7,Nar2,PartyName,MillName,BrokerName,Amt1,Nar8,Nar9,Dt1,Nar3,Nar4,Amt2,Amt3,Amt4,Amt5,Amt6,Amt7,Amt8,UserName) " & x1

                RecYN = GProcRstOpen(tmp1, "select * from TmpGentbl2 where Nar13 = 'DBN-Sales' and TmpGentbl2.UserName ='" & gUserName & "' order by Vtype,Vno,VSRNO  ", "R", gCn)
                If RecYN > 0 Then
                   Set sheet = workbook.Sheets.Item("DRN-Sales")
                   If (sheet Is Nothing) Then
                       MsgBox "Error DRN-Sales Sheet"
                   End If
                   sheet.cells(1, 1) = gCName
                   sheet.cells(2, 1) = "GSTIN : " + gCGStin
                   sheet.cells(3, 1) = "Audit Log"
                   sheet.cells(4, 1) = "Debit Note ( Sales )"
                   sheet.cells(5, 1) = "From : " & dtpFrDt.Value & "  To : " & dtpToDt.Value & ""
                   a = 8
                   With tmp1
                     .MoveFirst
                     Do While Not .EOF
                        sheet.cells(a, 1) = !Nar1
                        sheet.cells(a, 2) = !Nar7
                        sheet.cells(a, 3) = !Nar2
                        sheet.cells(a, 4) = !Nar3
                        sheet.cells(a, 5) = !Vdt
                        sheet.cells(a, 6) = !Nar4
                        sheet.cells(a, 7) = !dt1
                        sheet.cells(a, 8) = !PartyName
                        sheet.cells(a, 9) = !MillName
                        sheet.cells(a, 10) = !BrokerName
                        sheet.cells(a, 11) = !Amt2
                        sheet.cells(a, 12) = !Amt3
                        sheet.cells(a, 13) = !Amt4
                        sheet.cells(a, 14) = !Amt5
                        sheet.cells(a, 15) = !Amt6
                        sheet.cells(a, 16) = !Amt7
                        sheet.cells(a, 17) = !Amt8
                        sheet.cells(a, 18) = !Amt1
                        a = a + 1
'                        prgBar.Value = prgBar.Value + 1
                       .MoveNext
                     Loop
                   End With
                End If
             End If

             '-- For Credit Note Sales Log

             If OptCreditNote_Sales.Value = True Or OptAll.Value = True Then
                
                x1 = "select 'CRN-Sales',tblIntSale_log.Vfirm,tblIntSale_log.vno,tblIntSale_log.vtype,tblIntSale_log.vyear,tblIntSale_log.Vdt,tblIntSale_log.LogNo,tblIntSale_log.LogTp,tblIntSale_log.LogDate," _
                     & " tblIntSale_log.UserName,tblMastAccount.AcName,tblMastAccount_OAc.AcName,tblMastAccount_Br.AcName,tblIntSale_log.Vamt,tblIntSale_log.Vnar1,tblIntSale_log.VNar2,tblIntSale_log.VBillDate,tblIntSale_log.CrDrNoteNo,tblIntSale_log.VBillNo, " _
                     & " tblIntSale_log.Interest,tblIntSale_log.CgstRt,tblIntSale_log.CgstAmt,tblIntSale_log.SgstRt,tblIntSale_log.SgstAmt,tblIntSale_log.IgstRt,tblIntSale_log.IgstAmt,    '" & gUserName & "'" _
                     & " from tblIntSale_log,tblMastAccount,tblMastAccount tblMastAccount_OAc,tblMastAccount tblMastAccount_Br " _
                     & " where  tblIntSale_log.VCrAcCode = tblMastAccount.Accode And tblIntSale_log.VDrAcCode = tblMastAccount_OAC.AcCode And tblIntSale_log.Broker = tblMastAccount_Br.AcCode  and tblIntSale_log.vtype in ('PX','PN') and tblIntSale_log.VCtrNo = 1" _
                     & " and tblIntSale_log.Vdt>= ('" & Format(dtpFrDt.Value, "yyyy/mm/dd") & "') and tblIntSale_log.Vdt<= ('" & Format(dtpToDt.Value, "yyyy/mm/dd") & "')and tblIntSale_log.vfirm = '" & gCCode & "' and " _
                     & " convert(char(7),tblIntSale_log.vno)+tblIntSale_log.vtype+convert(char(4),tblIntSale_log.vyear)+tblIntSale_log.vfirm in " _
                     & " (select convert(char(7),tblIntSale_log.vno)+tblIntSale_log.vtype+convert(char(4),tblIntSale_log.vyear)+tblIntSale_log.vfirm " _
                     & " from tblIntSale_log where tblIntSale_log.VCtrNo = 1  group by convert(char(7),tblIntSale_log.vno)+tblIntSale_log.vtype+convert(char(4),tblIntSale_log.vyear)+tblIntSale_log.vfirm " _
                     & " having count(convert(char(7),tblIntSale_log.vno)+tblIntSale_log.vtype+convert(char(4),tblIntSale_log.vyear)+tblIntSale_log.vfirm) > 1) "
                
                gCn.Execute "Insert into tmpGenTbl2(Nar13,Vfirm,VNO,VTYPE,VYEAR,VDt,VSRNO,Nar1,Nar7,Nar2,PartyName,MillName,BrokerName,Amt1,Nar8,Nar9,Dt1,Nar3,Nar4,Amt2,Amt3,Amt4,Amt5,Amt6,Amt7,Amt8,UserName) " & x1
                
                RecYN = GProcRstOpen(tmp1, "select * from TmpGentbl2 where Nar13 = 'CRN-Sales' and TmpGentbl2.UserName ='" & gUserName & "' order by Vtype,Vno,VSRNO  ", "R", gCn)
                If RecYN > 0 Then
                   Set sheet = workbook.Sheets.Item("CRN-Sales")
                   If (sheet Is Nothing) Then
                       MsgBox "Error CRN-Sales Sheet"
                   End If
                   sheet.cells(1, 1) = gCName
                   sheet.cells(2, 1) = "GSTIN : " + gCGStin
                   sheet.cells(3, 1) = "Audit Log"
                   sheet.cells(4, 1) = "Credit Note ( Sales )"
                   sheet.cells(5, 1) = "From : " & dtpFrDt.Value & "  To : " & dtpToDt.Value & ""
                   a = 8
                   With tmp1
                     .MoveFirst
                     Do While Not .EOF
                        sheet.cells(a, 1) = !Nar1
                        sheet.cells(a, 2) = !Nar7
                        sheet.cells(a, 3) = !Nar2
                        sheet.cells(a, 4) = !Vno
                        sheet.cells(a, 5) = !Vdt
                        sheet.cells(a, 6) = !Nar4
                        sheet.cells(a, 7) = !dt1
                        sheet.cells(a, 8) = !PartyName
                        sheet.cells(a, 9) = !MillName
                        sheet.cells(a, 10) = !BrokerName
                        sheet.cells(a, 11) = !Amt2
                        sheet.cells(a, 12) = !Amt3
                        sheet.cells(a, 13) = !Amt4
                        sheet.cells(a, 14) = !Amt5
                        sheet.cells(a, 15) = !Amt6
                        sheet.cells(a, 16) = !Amt7
                        sheet.cells(a, 17) = !Amt8
                        sheet.cells(a, 18) = !Amt1
                        a = a + 1
'                        prgBar.Value = prgBar.Value + 1
                       .MoveNext
                     Loop
                   End With
                End If
             End If

             '-- For Debit Note Purchase Log

             If OptDebitNote_Purchase.Value = True Or OptAll.Value = True Then
                
                x1 = "select 'DBN-Purch',tblIntSale_log.Vfirm,tblIntSale_log.vno,tblIntSale_log.vtype,tblIntSale_log.vyear,tblIntSale_log.Vdt,tblIntSale_log.LogNo,tblIntSale_log.LogTp,tblIntSale_log.LogDate," _
                     & " tblIntSale_log.UserName,tblMastAccount.AcName,tblMastAccount_OAc.AcName,tblMastAccount_Br.AcName,tblIntSale_log.Vamt,tblIntSale_log.Vnar1,tblIntSale_log.VNar2,tblIntSale_log.VBillDate,tblIntSale_log.CrDrNoteNo,tblIntSale_log.VBillNo, " _
                     & " tblIntSale_log.Interest,tblIntSale_log.CgstRt,tblIntSale_log.CgstAmt,tblIntSale_log.SgstRt,tblIntSale_log.SgstAmt,tblIntSale_log.IgstRt,tblIntSale_log.IgstAmt,    '" & gUserName & "'" _
                     & " from tblIntSale_log,tblMastAccount,tblMastAccount tblMastAccount_OAc,tblMastAccount tblMastAccount_Br " _
                     & " where  tblIntSale_log.VDrAcCode = tblMastAccount.Accode And tblIntSale_log.VCrAcCode = tblMastAccount_OAC.AcCode And tblIntSale_log.Broker = tblMastAccount_Br.AcCode  and tblIntSale_log.vtype in ('SV','SQ') and tblIntSale_log.VCtrNo = 1" _
                     & " and tblIntSale_log.Vdt>= ('" & Format(dtpFrDt.Value, "yyyy/mm/dd") & "') and tblIntSale_log.Vdt<= ('" & Format(dtpToDt.Value, "yyyy/mm/dd") & "')and tblIntSale_log.vfirm = '" & gCCode & "' and " _
                     & " convert(char(7),tblIntSale_log.vno)+tblIntSale_log.vtype+convert(char(4),tblIntSale_log.vyear)+tblIntSale_log.vfirm in " _
                     & " (select convert(char(7),tblIntSale_log.vno)+tblIntSale_log.vtype+convert(char(4),tblIntSale_log.vyear)+tblIntSale_log.vfirm " _
                     & " from tblIntSale_log where tblIntSale_log.VCtrNo = 1  group by convert(char(7),tblIntSale_log.vno)+tblIntSale_log.vtype+convert(char(4),tblIntSale_log.vyear)+tblIntSale_log.vfirm " _
                     & " having count(convert(char(7),tblIntSale_log.vno)+tblIntSale_log.vtype+convert(char(4),tblIntSale_log.vyear)+tblIntSale_log.vfirm) > 1) "
                
                gCn.Execute "Insert into tmpGenTbl2(Nar13,Vfirm,VNO,VTYPE,VYEAR,VDt,VSRNO,Nar1,Nar7,Nar2,PartyName,MillName,BrokerName,Amt1,Nar8,Nar9,Dt1,Nar3,Nar4,Amt2,Amt3,Amt4,Amt5,Amt6,Amt7,Amt8,UserName) " & x1
                RecYN = GProcRstOpen(tmp1, "select * from TmpGentbl2 where Nar13 = 'DBN-Purch' and TmpGentbl2.UserName ='" & gUserName & "' order by Vtype,Vno,VSRNO  ", "R", gCn)
                If RecYN > 0 Then
                   Set sheet = workbook.Sheets.Item("DRN-Purch")
                   If (sheet Is Nothing) Then
                       MsgBox "Error DRN-Purch Sheet"
                   End If
                   sheet.cells(1, 1) = gCName
                   sheet.cells(2, 1) = "GSTIN : " + gCGStin
                   sheet.cells(3, 1) = "Audit Log"
                   sheet.cells(4, 1) = "Debit Note ( Purchase )"
                   sheet.cells(5, 1) = "From : " & dtpFrDt.Value & "  To : " & dtpToDt.Value & ""
                   a = 8
                   With tmp1
                     .MoveFirst
                     Do While Not .EOF
                        sheet.cells(a, 1) = !Nar1
                        sheet.cells(a, 2) = !Nar7
                        sheet.cells(a, 3) = !Nar2
                        sheet.cells(a, 4) = !Vno
                        sheet.cells(a, 5) = !Vdt
                        sheet.cells(a, 6) = !Nar4
                        sheet.cells(a, 7) = !dt1
                        sheet.cells(a, 8) = !PartyName
                        sheet.cells(a, 9) = !MillName
                        sheet.cells(a, 10) = !BrokerName
                        sheet.cells(a, 11) = !Amt2
                        sheet.cells(a, 12) = !Amt3
                        sheet.cells(a, 13) = !Amt4
                        sheet.cells(a, 14) = !Amt5
                        sheet.cells(a, 15) = !Amt6
                        sheet.cells(a, 16) = !Amt7
                        sheet.cells(a, 17) = !Amt8
                        sheet.cells(a, 18) = !Amt1
                        a = a + 1
'                        prgBar.Value = prgBar.Value + 1
                       .MoveNext
                     Loop
                   End With
                End If
             End If

             '-- For Credit Note Purchase Log

             If OptCreditNote_Purchase.Value = True Or OptAll.Value = True Then
                
                x1 = "select 'CRN-Purch',tblIntSale_log.Vfirm,tblIntSale_log.vno,tblIntSale_log.vtype,tblIntSale_log.vyear,tblIntSale_log.Vdt,tblIntSale_log.LogNo,tblIntSale_log.LogTp,tblIntSale_log.LogDate," _
                     & " tblIntSale_log.UserName,tblMastAccount.AcName,tblMastAccount_OAc.AcName,tblMastAccount_Br.AcName,tblIntSale_log.Vamt,tblIntSale_log.Vnar1,tblIntSale_log.VNar2,tblIntSale_log.VBillDate,tblIntSale_log.CrDrNoteNo,tblIntSale_log.VBillNo, " _
                     & " tblIntSale_log.Interest,tblIntSale_log.CgstRt,tblIntSale_log.CgstAmt,tblIntSale_log.SgstRt,tblIntSale_log.SgstAmt,tblIntSale_log.IgstRt,tblIntSale_log.IgstAmt,    '" & gUserName & "'" _
                     & " from tblIntSale_log,tblMastAccount,tblMastAccount tblMastAccount_OAc,tblMastAccount tblMastAccount_Br " _
                     & " where  tblIntSale_log.VCrAcCode = tblMastAccount.Accode And tblIntSale_log.VDrAcCode = tblMastAccount_OAC.AcCode And tblIntSale_log.Broker = tblMastAccount_Br.AcCode  and tblIntSale_log.vtype = 'PV' and tblIntSale_log.VCtrNo = 1" _
                     & " and tblIntSale_log.Vdt>= ('" & Format(dtpFrDt.Value, "yyyy/mm/dd") & "') and tblIntSale_log.Vdt<= ('" & Format(dtpToDt.Value, "yyyy/mm/dd") & "')and tblIntSale_log.vfirm = '" & gCCode & "' and " _
                     & " convert(char(7),tblIntSale_log.vno)+tblIntSale_log.vtype+convert(char(4),tblIntSale_log.vyear)+tblIntSale_log.vfirm in " _
                     & " (select convert(char(7),tblIntSale_log.vno)+tblIntSale_log.vtype+convert(char(4),tblIntSale_log.vyear)+tblIntSale_log.vfirm " _
                     & " from tblIntSale_log where tblIntSale_log.VCtrNo = 1  group by convert(char(7),tblIntSale_log.vno)+tblIntSale_log.vtype+convert(char(4),tblIntSale_log.vyear)+tblIntSale_log.vfirm " _
                     & " having count(convert(char(7),tblIntSale_log.vno)+tblIntSale_log.vtype+convert(char(4),tblIntSale_log.vyear)+tblIntSale_log.vfirm) > 1) "
                
                gCn.Execute "Insert into tmpGenTbl2(Nar13,Vfirm,VNO,VTYPE,VYEAR,VDt,VSRNO,Nar1,Nar7,Nar2,PartyName,MillName,BrokerName,Amt1,Nar8,Nar9,Dt1,Nar3,Nar4,Amt2,Amt3,Amt4,Amt5,Amt6,Amt7,Amt8,UserName) " & x1
                RecYN = GProcRstOpen(tmp1, "select * from TmpGentbl2 where Nar13 = 'CRN-Purch' and TmpGentbl2.UserName ='" & gUserName & "' order by Vtype,Vno,VSRNO  ", "R", gCn)
                If RecYN > 0 Then
                   Set sheet = workbook.Sheets.Item("CRN-Purch")
                   If (sheet Is Nothing) Then
                       MsgBox "Error CRN-Purch Sheet"
                   End If
                   sheet.cells(1, 1) = gCName
                   sheet.cells(2, 1) = "GSTIN : " + gCGStin
                   sheet.cells(3, 1) = "Audit Log"
                   sheet.cells(4, 1) = "Credit Note ( Purchase )"
                   sheet.cells(5, 1) = "From : " & dtpFrDt.Value & "  To : " & dtpToDt.Value & ""
                   a = 8
                   With tmp1
                     .MoveFirst
                     Do While Not .EOF
                        sheet.cells(a, 1) = !Nar1
                        sheet.cells(a, 2) = !Nar7
                        sheet.cells(a, 3) = !Nar2
                        sheet.cells(a, 4) = !Vno
                        sheet.cells(a, 5) = !Vdt
                        sheet.cells(a, 6) = !Nar4
                        sheet.cells(a, 7) = !dt1
                        sheet.cells(a, 8) = !PartyName
                        sheet.cells(a, 9) = !MillName
                        sheet.cells(a, 10) = !BrokerName
                        sheet.cells(a, 11) = !Amt2
                        sheet.cells(a, 12) = !Amt3
                        sheet.cells(a, 13) = !Amt4
                        sheet.cells(a, 14) = !Amt5
                        sheet.cells(a, 15) = !Amt6
                        sheet.cells(a, 16) = !Amt7
                        sheet.cells(a, 17) = !Amt8
                        sheet.cells(a, 18) = !Amt1
'                        prgBar.Value = prgBar.Value + 1
                        a = a + 1
                       .MoveNext
                     Loop
                   End With
                End If
             End If


             '-- For JV Log

             If OptJv.Value = True Or OptAll.Value = True Then
                
                x1 = "select 'JV',tblVoucher_Log.Vfirm,tblVoucher_Log.vno,tblVoucher_Log.vtype,tblVoucher_Log.vyear,tblVoucher_Log.Vdt,tblVoucher_Log.LogNo,tblVoucher_Log.LogTp,tblVoucher_Log.LogDate," _
                     & " tblVoucher_Log.UserName,tblMastAccount.AcName,tblMastAccount_OAc.AcName,tblVoucher_Log.Vamt,tblVoucher_Log.Vnar1,tblVoucher_Log.VNar2,tblVoucher_Log.VBillNo,tblVoucher_Log.VBillDate,tblVoucher_Log.TDSOn,tblVoucher_Log.LessTds,'" & gUserName & "'" _
                     & " from tblVoucher_Log,tblMastAccount,tblMastAccount tblMastAccount_OAc" _
                     & " where  tblVoucher_Log.VCrAcCode = tblMastAccount.Accode And tblVoucher_Log.VDrAcCode = tblMastAccount_OAC.AcCode and tblVoucher_Log.vtype = 'JV' and tblVoucher_Log.VCtrNo = 1" _
                     & " and tblVoucher_Log.Vdt>= ('" & Format(dtpFrDt.Value, "yyyy/mm/dd") & "') and tblVoucher_Log.Vdt<= ('" & Format(dtpToDt.Value, "yyyy/mm/dd") & "')and tblVoucher_Log.vfirm = '" & gCCode & "' and " _
                     & " convert(char(7),tblVoucher_Log.vno)+tblVoucher_Log.vtype+convert(char(4),tblVoucher_Log.vyear)+tblVoucher_Log.vfirm in " _
                     & " (select convert(char(7),tblVoucher_Log.vno)+tblVoucher_Log.vtype+convert(char(4),tblVoucher_Log.vyear)+tblVoucher_Log.vfirm " _
                     & " from tblVoucher_Log where tblVoucher_Log.VCtrNo = 1  group by convert(char(7),tblVoucher_Log.vno)+tblVoucher_Log.vtype+convert(char(4),tblVoucher_Log.vyear)+tblVoucher_Log.vfirm " _
                     & " having count(convert(char(7),tblVoucher_Log.vno)+tblVoucher_Log.vtype+convert(char(4),tblVoucher_Log.vyear)+tblVoucher_Log.vfirm) > 1) "

                gCn.Execute "Insert into tmpGenTbl2(Nar13,Vfirm,VNO,VTYPE,VYEAR,VDt,VSRNO,Nar1,Nar7,Nar2,PartyName,MillName,Amt1,Nar8,Nar9,Nar10,Dt1,Amt2,Amt3,UserName) " & x1
                
                RecYN = GProcRstOpen(tmp1, "select * from TmpGentbl2 where Nar13 = 'JV' and TmpGentbl2.UserName ='" & gUserName & "' order by Vtype,Vno,VSRNO  ", "R", gCn)
                If RecYN > 0 Then
                   Set sheet = workbook.Sheets.Item("JV")
                   If (sheet Is Nothing) Then
                       MsgBox "Error JV Sheet"
                   End If
                   sheet.cells(1, 1) = gCName
                   sheet.cells(2, 1) = "GSTIN : " + gCGStin
                   sheet.cells(3, 1) = "Audit Log"
                   sheet.cells(4, 1) = "JV"
                   sheet.cells(5, 1) = "From : " & dtpFrDt.Value & "  To : " & dtpToDt.Value & ""
                   a = 8
                   With tmp1
                     .MoveFirst
                     Do While Not .EOF
                        sheet.cells(a, 1) = !Nar1
                        sheet.cells(a, 2) = !Nar7
                        sheet.cells(a, 3) = !Nar2
                        sheet.cells(a, 4) = !Vno
                        sheet.cells(a, 5) = !Vdt
                        sheet.cells(a, 6) = !MillName
                        sheet.cells(a, 7) = !PartyName
                        sheet.cells(a, 8) = !Amt1
                        sheet.cells(a, 9) = !Nar8
                        sheet.cells(a, 10) = !Nar9
                        sheet.cells(a, 11) = !Nar10
                        sheet.cells(a, 12) = !dt1
                        sheet.cells(a, 13) = !Amt2
                        sheet.cells(a, 14) = !Amt3
                        a = a + 1
'                        prgBar.Value = prgBar.Value + 1
                       .MoveNext
                     Loop
                   End With
                End If
             End If
             workbook.Close SaveChanges:=True
             MsgBox "Compileted"
             Set xlApp = CreateObject("Excel.Application")
             xlApp.Visible = True
             Set xlWB = xlApp.Workbooks.Open("" & flnm & "")
             excel_app.Quit
             Set excel_app = Nothing
             Exit Sub
        End Select
    .Action = 1
End With
End Sub





