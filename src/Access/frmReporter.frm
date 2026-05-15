VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmReporter 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Form1"
   ClientHeight    =   2775
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5985
   Icon            =   "frmReporter.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2775
   ScaleWidth      =   5985
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmd1 
      Caption         =   "Exit"
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
      Left            =   4200
      TabIndex        =   10
      Top             =   2040
      Width           =   1455
   End
   Begin VB.CommandButton cmd1 
      Caption         =   "Screen"
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
      Index           =   1
      Left            =   2280
      TabIndex        =   9
      Top             =   2040
      Width           =   1455
   End
   Begin VB.TextBox Text2 
      Height          =   375
      Left            =   1680
      TabIndex        =   5
      Top             =   1320
      Visible         =   0   'False
      Width           =   3975
   End
   Begin VB.TextBox Text1 
      Enabled         =   0   'False
      Height          =   375
      Left            =   1680
      TabIndex        =   3
      Top             =   120
      Width           =   3975
   End
   Begin VB.CommandButton cmd1 
      Caption         =   "Print"
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
      Index           =   0
      Left            =   360
      TabIndex        =   0
      Top             =   2040
      Width           =   1455
   End
   Begin MSComCtl2.DTPicker DTPicker1 
      Height          =   375
      Left            =   1680
      TabIndex        =   1
      Top             =   840
      Visible         =   0   'False
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   661
      _Version        =   393216
      CustomFormat    =   "dd/MM/yyyy"
      Format          =   55115779
      CurrentDate     =   38668
   End
   Begin MSComCtl2.DTPicker DTPicker2 
      Height          =   375
      Left            =   4080
      TabIndex        =   2
      Top             =   840
      Visible         =   0   'False
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   661
      _Version        =   393216
      Format          =   55115777
      CurrentDate     =   38668
   End
   Begin VB.Label Label1 
      Caption         =   "Period From :"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   2
      Left            =   120
      TabIndex        =   8
      Top             =   840
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.Label Label1 
      Caption         =   "To"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   1
      Left            =   3600
      TabIndex        =   7
      Top             =   840
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Label Label2 
      Caption         =   "Single A/c  :"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   6
      Top             =   1320
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "Report Name :"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   4
      Top             =   120
      Width           =   1575
   End
End
Attribute VB_Name = "frmReporter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim m_Type As String
Private Sub cmd1_Click(Index As Integer)
gfrmTypeStr = "START"
Select Case Index
Case 0
   Call GenRepo(False, Text1)
Case 1
   Call GenRepo(True, Text1)
Case 2
   Unload Me
End Select
If (gfrmTypeStr = "") And (Index < 2) Then
    MsgBox "No Report Available", vbInformation + vbOKOnly, "No Report"
End If
End Sub
Private Sub Form_Activate()
If Text1.Text <> "" Then
   If GProcGetColumnValue("udfreports", "Reponame", Text1, "S", "RDTRANGE", "S", "", gSetupCn) = 1 Then
      DTPicker1.Visible = True
      DTPicker2.Visible = True
      Label1(1).Visible = True
      Label1(2).Visible = True
   End If
   If GProcGetColumnValue("udfreports", "Reponame", Text1, "S", "Rsingac", "S", "", gSetupCn) = 1 Then
      Text2.Visible = True
      Label2.Visible = True
   End If
End If
m_Type = GProcGetColumnValue("udfreports", "Reponame", Text1, "S", "Rsinglefrom", "S", "", gSetupCn) & vbNullString
End Sub
Public Sub GenRepo(mR_Screen As Boolean, mR_Rpnm As String)
Dim LnCnt As Integer, PgCnt As Integer, wHdPrn As Boolean, wDotLine As String, RpHdr As String, wSubLine As String
Dim i As Double, Wt(30) As Double, wMonT(30) As Double, wTSub(30) As Double, wFldLen(30) As Integer, wGrp As String, RetVal, wMaxFld As Integer
Dim c_Field As ADODB.Field, wGTP As Boolean, X1 As String, mPc1 As String, RptRst As New Recordset, mPc As String
Dim wInvNo As String, myCode As String, myMonth As Integer, rst1 As Recordset, wDoubleLine As String
wInvNo = ""
If GProcRstOpen(RptRst, "Select * from UDFREPORTS where REPONAME = '" & mR_Rpnm & "' ", "R", gSetupCn) = 0 Then
   gfrmTypeStr = ""
   Exit Sub
End If
RptRst.MoveFirst
For i = 0 To 30
    Wt(i) = 0
    wTSub(i) = 0
    wMonT(i) = 0
    wFldLen(i) = 0
Next i
If (RptRst!RCALLREPO <> 0) Then
    Call CollectData(RptRst!RCALLREPO)
End If
mPc1 = Chr(27) + Chr(18)
mPc = ""
If mR_Screen = False Then
    Select Case RptRst!RCPI
    Case 10:
        mPc = Chr(27) + Chr(18)
    Case 12:
        mPc = Chr(27) + Chr(15)
    Case 17:
        mPc = Chr(27) + Chr(33) + Chr(1)
        mPc1 = Chr(27) + Chr(33) + Chr(0)
    End Select
End If
If gBackEndDB = gBackEndAccess Then
    X1 = "Select " + RptRst!RSQL + " From " + RptRst!RTBLS
Else '-- Oracle
    X1 = "Select " + RptRst!RSQL_ORA + " From " + RptRst!RTBLS
End If
If RptRst!RCND & vbNullString <> "" Then
   X1 = X1 + " Where " + RptRst!RCND
End If
If RptRst!RDTRANGE = 1 Then
    If InStr(UCase(X1), "WHERE") > 0 Then
        X1 = X1 + " and "
    Else
        X1 = X1 + " where "
    End If
    If gBackEndDB = gBackEndAccess Then
       X1 = X1 + " format$(" & RptRst!RDTFIELD & ",'yyyy/MM/dd') >= '" & Format$(DTPicker1, "yyyy/MM/dd") & "' and format$(" & RptRst!RDTFIELD & ",'yyyy/MM/dd') <= '" & Format$(DTPicker2, "yyyy/MM/dd") & "' "
    Else
       X1 = X1 + RptRst!RDTFIELD & " >= to_date('" & Format$(DTPicker1, "dd/MM/yyyy") & "','dd/MM/yyyy') and " & RptRst!RDTFIELD & " <= to_date('" & Format$(DTPicker2, "dd/MM/yyyy") & "','dd/MM/yyyy') "
    End If
End If
If (RptRst!RSINGAC = 1) And (Text2.Text <> "") Then
    If (RptRst!RSINGLEFROM = "A") Then
       myCode = GProcGetColumnValue("tblmastaccount", "acname", Text2, "S", "Accode", "N")
    Else
       myCode = GProcGetColumnValue("tblmastgroup", "agname", Text2, "S", "Agcode", "N")
    End If
    If InStr(UCase(X1), "WHERE") > 0 Then
        X1 = X1 + " and "
    Else
        X1 = X1 + " where "
    End If
    X1 = X1 + RptRst!RACFLD & " = " & myCode & " "
End If
If RptRst!RORDBY <> "" Then X1 = X1 + " Order by " + RptRst!RORDBY
If GProcRstOpen(rst1, X1, "R", gCn) = 0 Then
   gfrmTypeStr = ""
   Exit Sub
End If
RpHdr = ""
wGrp = "dsfhfd"
With rst1
    i = 0
    '--------------------------------------Getting Report Header
    wMaxFld = .Fields.Count
    For i = 0 To wMaxFld - 1
        wFldLen(i) = Val(Mid(RptRst!RFldLen, (i * 2) + 1, 2))
        If Not (i = 0 And RptRst!RType = "G") Then
            Select Case VarType(.Fields(i))
            Case 2, 3, 5:
               RpHdr = RpHdr + Right(Space(wFldLen(i)) + .Fields(i).Name, wFldLen(i)) + "  "
            Case Else
               RpHdr = RpHdr + Left(.Fields(i).Name + Space(wFldLen(i)), wFldLen(i)) + "  "
            End Select
        End If
    Next i
'    wDotLine = Space(RptRst!RLeftMrg) + String(Len(RpHdr), Chr(196))
'    wDoubleLine = Space(RptRst!RLeftMrg) + String(Len(RpHdr), Chr(205))
    wDotLine = Space(RptRst!RLeftMrg) + String(Len(RpHdr), "-")
    wDoubleLine = Space(RptRst!RLeftMrg) + String(Len(RpHdr), "-")
    Open "c:\rpt1.txt" For Output As #1
    wHdPrn = True
    LnCnt = 0
    PgCnt = 1
    .MoveFirst
    Do While Not .EOF
        For i = 0 To 30
            wMonT(i) = 0
        Next i
       If RptRst!RMONTHTOT = 1 Then myMonth = Month(.Fields(RptRst!RDTFIELD))
       Do While Not .EOF
         If RptRst!RMONTHTOT = 1 Then
            If myMonth <> Month(.Fields(RptRst!RDTFIELD)) Then Exit Do
        End If
         If wHdPrn Then
            wHdPrn = False
            LnCnt = 2
            If ((Len(RpHdr) / 2) - (Len(Trim(X1)) / 2) > 0) Then
               i = (Len(RpHdr) / 2) - (Len(Trim(X1)) / 2)
            Else
               i = 0
            End If
            If mR_Screen = True Then
                Print #1, Space(RptRst!RLeftMrg) + Space(Len(RpHdr) / 2 - Len(Trim(gCName)) / 2) + gCName
                X1 = RptRst!RTitle1
                If RptRst!RDTRANGE = 1 Then X1 = X1 + " From " & Format(DTPicker1, "dd/MM/yyyy") & " To " & Format(DTPicker2, "dd/MM/yyyy")
                Print #1, Space(RptRst!RLeftMrg) + Space(i) + X1 & Space(15) + "Page No. " + Str(PgCnt)
                If RptRst!RTitle2 <> "" Then
                    Print #1, Space(RptRst!RLeftMrg) + Space(Len(RpHdr) / 2 - Len(Trim(RptRst!RTitle2)) / 2) + RptRst!RTitle2
                    LnCnt = LnCnt + 1
                End If
            Else
                Print #1, Space(RptRst!RLeftMrg) + Space(Len(RpHdr) / 2 - Len(Trim(gCName))) + Chr(18) + Chr(27) + Chr(14) + gCName & Chr(27) + Chr(18)
                X1 = Chr(27) + Chr(69) & RptRst!RTitle1 & Chr(27) + Chr(70)
                If RptRst!RDTRANGE = 1 Then X1 = X1 + " From " & Format(DTPicker1, "dd/MM/yyyy") & " To " & Format(DTPicker2, "dd/MM/yyyy")
                Print #1, Space(RptRst!RLeftMrg) + Space(i) + X1 & Space(15) + "Page No. " + Str(PgCnt)
                If RptRst!RTitle2 <> "" Then
                    Print #1, Space(RptRst!RLeftMrg) + Space(Len(RpHdr) / 2 - Len(Trim(RptRst!RTitle2)) / 2) + Chr(27) + Chr(69) & RptRst!RTitle2 & Chr(27) + Chr(70)
                    LnCnt = LnCnt + 1
                End If
            End If
            PgCnt = PgCnt + 1
            Print #1, mPc & wDotLine
            Print #1, Space(RptRst!RLeftMrg) + RpHdr
            Print #1, wDoubleLine
            LnCnt = LnCnt + 5
        End If
        If RptRst!RType = "G" Then
            If wGrp <> .Fields(0) Then
                wGrp = .Fields(0)
                If mR_Screen = True Then
                    Print #1, Space(RptRst!RLeftMrg) + .Fields(0)
                Else
                    Print #1, Space(RptRst!RLeftMrg) + Chr(27) + Chr(69) & .Fields(0) & Chr(27) + Chr(70)
                End If
                LnCnt = LnCnt + 1
                For i = 0 To 30
                    wTSub(i) = 0
                Next i
            End If
        End If
        '--------------------------------------Data Print
        X1 = Space(RptRst!RLeftMrg)
        i = 0
        Dim wSuppFld As Boolean
        wSuppFld = True
        If RptRst!RMULTISIG = "M" Then
            If wInvNo <> .Fields(0) Then
                wInvNo = .Fields(0)
                wSuppFld = False
            End If
        End If
        For i = 0 To wMaxFld - 1
            If Not (i = 0 And RptRst!RType = "G") Then
                Select Case VarType(.Fields(i))
                Case 0, 1, 8: 'vbEmpty, vbNull, vbString
                    If (RptRst!RMULTISIG = "M") And (wSuppFld = True) And (InStr(RptRst!RFLDSUPP, GProcNoToStr(i, 2, 0)) = 0) Then
                       X1 = X1 + Space(wFldLen(i)) + "  "
                    Else
                        X1 = X1 + Left(.Fields(i) + Space(wFldLen(i)), wFldLen(i)) + "  "
                    End If
                Case 2, 3: 'vbInteger, vbLong
                    If (RptRst!RMULTISIG = "M") And (wSuppFld = True) And (InStr(RptRst!RFLDSUPP, GProcNoToStr(i, 2, 0)) = 0) Then
                        X1 = X1 + Space(wFldLen(i) + 2)
                    Else
                        X1 = X1 + GProcNoToStr(.Fields(i), wFldLen(i), 0) + "  "
                    End If
                    '--------------------------------Calculating Totals
                    If (InStr(RptRst!RTotals, GProcNoToStr(i, 2, 0)) > 0) Then
                       Wt(i) = Wt(i) + .Fields(i)
                       wMonT(i) = wMonT(i) + .Fields(i)
                       If RptRst!RType = "G" Then wTSub(i) = wTSub(i) + .Fields(i)
                    End If
                Case 5:  ' vbDouble
                    If (RptRst!RMULTISIG = "M") And (wSuppFld = True) And (InStr(RptRst!RFLDSUPP, GProcNoToStr(i, 2, 0)) = 0) Then
                        X1 = X1 + Space(wFldLen(i) + 2)
                    Else
                        X1 = X1 + GProcNoToStr(.Fields(i), wFldLen(i), 2) + "  "
                    End If
                    
                    If (InStr(RptRst!RTotals, GProcNoToStr(i, 2, 0)) > 0) Then
                       Wt(i) = Wt(i) + .Fields(i)
                       If RptRst!RType = "G" Then wTSub(i) = wTSub(i) + .Fields(i)
                    End If
                Case 7: 'vbDate
                    If (RptRst!RMULTISIG = "M") And (wSuppFld = True) And (InStr(RptRst!RFLDSUPP, GProcNoToStr(i, 2, 0)) = 0) Then
                        X1 = X1 + Space(10)
                    Else
                        X1 = X1 + Format(.Fields(i), "dd/MM/yy") + "  "
                    End If
                End Select
            End If
        Next i
        Print #1, X1
        LnCnt = LnCnt + 1
        If (LnCnt > (RptRst!RPgLen - 5)) Then
            Print #1, wDotLine + mPc1 + Chr(12)
            wHdPrn = True
            LnCnt = 0
        End If
        .MoveNext
        wGTP = False
        If RptRst!RType = "G" Then
            If .EOF Then
                wGTP = True
            Else
                If wGrp <> .Fields(0) Then wGTP = True
            End If
            If wGTP Then
                '--------------------------------------SubGroup Total Print
                X1 = Space(RptRst!RLeftMrg)
                wSubLine = Space(RptRst!RLeftMrg)
                For i = 1 To wMaxFld - 1
                    If wTSub(i) > 0 Then
                        If wFldLen(i) < 5 Then
                           X1 = X1 + GProcNoToStr(wTSub(i), wFldLen(i), 0) + "  "
                        Else
                           X1 = X1 + GProcNoToStr(wTSub(i), wFldLen(i), 2) + "  "
                        End If
                        wSubLine = wSubLine + String(12, "-") + "  "
                    Else
                        X1 = X1 + Space(wFldLen(i)) + "  "
                        wSubLine = wSubLine + Space(wFldLen(i)) + "  "
                    End If
                    wTSub(i) = 0
                Next i
                Print #1, wSubLine
                Print #1, X1
                LnCnt = LnCnt + 2
                DoEvents
            End If
        End If
    Loop
            If (RptRst!RMONTHTOT = 1) Then
                '--------------------------------------Monthly Total Print
                X1 = Space(RptRst!RLeftMrg)
                wSubLine = Space(RptRst!RLeftMrg)
                For i = 1 To wMaxFld - 1
                    If wMonT(i) > 0 Then
                        If wFldLen(i) < 5 Then
                           X1 = X1 + GProcNoToStr(wMonT(i), wFldLen(i), 0) + "  "
                        Else
                           X1 = X1 + GProcNoToStr(wMonT(i), wFldLen(i), 2) + "  "
                        End If
                        wSubLine = wSubLine + String(12, "-") + "  "
                    Else
                        X1 = X1 + Space(wFldLen(i)) + "  "
                        wSubLine = wSubLine + Space(wFldLen(i)) + "  "
                    End If
                    wTSub(i) = 0
                Next i
                Print #1, wSubLine
                Print #1, X1
                LnCnt = LnCnt + 2
                DoEvents
            End If
    Loop
    '--------------------------------------Total Print
    Print #1, wDotLine
    X1 = ""
    For i = 0 To wMaxFld - 1
        If Not (i = 0 And RptRst!RType = "G") Then
            If Wt(i) > 0 Then
               If wFldLen(i) >= 5 Then
                   X1 = X1 + GProcNoToStr(Wt(i), wFldLen(i), 2) + "  "
               Else
                   X1 = X1 + GProcNoToStr(Wt(i), wFldLen(i), 0) + "  "
               End If
            Else
                X1 = X1 + Space(wFldLen(i)) + "  "
            End If
        End If
    Next i
    Print #1, Space(RptRst!RLeftMrg) + "GRAND TOTAL" + Mid(X1, 12)
    .Close
End With
Print #1, wDoubleLine + mPc1 + Chr(12)
Close #1
RptRst.Close
'Unload Reporter
    If mR_Screen = False Then
       RetVal = Shell(App.Path & "\txtprint.exe", vbMaximizedFocus)
    ElseIf mR_Screen = True Then
       RetVal = Shell(App.Path & "\WORDPAD" & " c:\rpt1.txt", vbMaximizedFocus)
    End If
End Sub
Private Sub Text1_KeyPress(KeyAscii As Integer)
gClsSearch.SearchMultiField "UdfReports", "RepoName", Array("Report Name"), Array(Text1.Width), "", Chr(KeyAscii), "RepoName", Text1.Left + Me.Left, Text1.Top + Me.Top + 650
KeyAscii = 0
If gClsSearch.SearchMultiRetCol(0) <> "" Then Text1.Text = gClsSearch.SearchMultiRetCol(0)
End Sub
Private Sub Text2_KeyPress(KeyAscii As Integer)
If m_Type = "A" Then
   gClsSearch.SearchMultiField "tblmastAccount", "AcName", Array("Account Name"), Array(Text2.Width), "", Chr(KeyAscii), "AcName", Text2.Left + Me.Left, Text2.Top + Me.Top + 650
Else
   gClsSearch.SearchMultiField "tblmastGroup", "AgName", Array("Group Name"), Array(Text2.Width), "", Chr(KeyAscii), "AgName", Text2.Left + Me.Left, Text2.Top + Me.Top + 650
End If
KeyAscii = 0
If gClsSearch.SearchMultiRetCol(0) <> "" Then Text2.Text = gClsSearch.SearchMultiRetCol(0)
End Sub
