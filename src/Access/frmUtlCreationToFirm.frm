VERSION 5.00
Begin VB.Form frmUtlCreationToFirm 
   BackColor       =   &H00C0E0FF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Creation To Firm"
   ClientHeight    =   7680
   ClientLeft      =   30
   ClientTop       =   330
   ClientWidth     =   11415
   Icon            =   "frmUtlCreationToFirm.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7680
   ScaleWidth      =   11415
   Begin VB.CommandButton cmdSelect 
      Caption         =   "&Select All"
      Height          =   400
      Left            =   1050
      TabIndex        =   8
      Top             =   6525
      Width           =   1425
   End
   Begin VB.CommandButton cmdDeselect 
      Caption         =   "&Deselect"
      Height          =   400
      Left            =   2625
      TabIndex        =   9
      Top             =   6525
      Width           =   1425
   End
   Begin VB.TextBox txtTCompany 
      Height          =   375
      Left            =   5505
      TabIndex        =   3
      Text            =   "txtTCompany"
      Top             =   570
      Width           =   5640
   End
   Begin VB.TextBox txtSCompany 
      Height          =   375
      Left            =   210
      TabIndex        =   1
      Text            =   "txtSCompany"
      Top             =   570
      Width           =   5145
   End
   Begin VB.CommandButton cmdClose 
      Cancel          =   -1  'True
      Caption         =   "&Close"
      Height          =   400
      Left            =   7350
      TabIndex        =   11
      Top             =   6525
      Width           =   1425
   End
   Begin VB.CommandButton cmdOk 
      Caption         =   "&Ok"
      Default         =   -1  'True
      Height          =   400
      Left            =   5775
      TabIndex        =   10
      Top             =   6540
      Width           =   1425
   End
   Begin VB.Timer Timer1 
      Interval        =   100
      Left            =   3000
      Top             =   2385
   End
   Begin VB.ListBox lisTarget 
      ForeColor       =   &H00000000&
      Height          =   4335
      Left            =   5505
      Sorted          =   -1  'True
      Style           =   1  'Checkbox
      TabIndex        =   7
      Top             =   1560
      Width           =   5640
   End
   Begin VB.ListBox lisSource 
      ForeColor       =   &H00000000&
      Height          =   4335
      Left            =   210
      Sorted          =   -1  'True
      Style           =   1  'Checkbox
      TabIndex        =   5
      Top             =   1560
      Width           =   5145
   End
   Begin VB.Shape Shape1 
      FillColor       =   &H00BDDFD1&
      FillStyle       =   0  'Solid
      Height          =   825
      Left            =   795
      Shape           =   4  'Rounded Rectangle
      Top             =   6285
      Width           =   8295
   End
   Begin VB.Shape Shape3 
      BorderColor     =   &H00000080&
      BorderWidth     =   5
      FillColor       =   &H000000C0&
      Height          =   7725
      Left            =   -30
      Top             =   -75
      Width           =   11445
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   240
      Picture         =   "frmUtlCreationToFirm.frx":058A
      Top             =   7170
      Width           =   480
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "&To Company  :"
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
      Height          =   255
      Left            =   5505
      TabIndex        =   2
      Top             =   255
      Width           =   5640
   End
   Begin VB.Label Label1 
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "&From Company  :"
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
      Height          =   315
      Left            =   210
      TabIndex        =   0
      Top             =   255
      Width           =   5145
   End
   Begin VB.Label lblT 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   195
      Left            =   960
      TabIndex        =   13
      Top             =   4065
      Width           =   5235
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Selected &Item  :"
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
      Height          =   255
      Left            =   5505
      TabIndex        =   6
      Top             =   1170
      Width           =   5640
   End
   Begin VB.Label Label2 
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "Select From Item &List  :"
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
      Height          =   255
      Left            =   210
      TabIndex        =   4
      Top             =   1170
      Width           =   5145
   End
   Begin VB.Label lblS 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00008000&
      Height          =   195
      Left            =   300
      TabIndex        =   12
      Top             =   4065
      Width           =   4335
   End
   Begin VB.Shape Shape2 
      BackColor       =   &H00808080&
      BorderColor     =   &H004D177D&
      FillColor       =   &H00EDFDFE&
      FillStyle       =   0  'Solid
      Height          =   6015
      Left            =   75
      Top             =   120
      Width           =   11205
   End
End
Attribute VB_Name = "frmUtlCreationToFirm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim tmpSCn As Connection
Dim tmpTCn As Connection
Dim i As Long
Dim frmTypeStr As String
Private Sub Form_Load()
Timer1.Enabled = False
txtTCompany = ""
txtSCompany = gCName
frmTypeStr = gfrmTypeStr
FillItem
End Sub
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case 13:
            If Me.ActiveControl.Name <> "txtGrid" Then SendKeys "{TAB}"
    End Select
End Sub
Private Sub cmdOK_Click()
Dim Ans As Long
Dim RsS As Recordset
Dim RsT As Recordset
Dim j As Long
Dim TDefaultAcCode As Long
'====== For Account
Dim AcName As String '-- Account name
Dim SAgName As String '-- Source A/C Group
Dim SBrkAcName As String  '-- Source Broker A/C
Dim SCgName As String  '-- Source Party Group
Dim TAgCode As Long '-- To A/c -> Group Code
Dim TBrkAcCode As Long '-- To A/c -> Broker Ac Code
Dim TCgCode As Long '-- To A/c -> Party Group Code
'====== For Narration
Dim Narration As String '-- Narration
'====== For Item
Dim ItCode As Long '-- Item Code
Dim ItName As String '-- Item name
Dim SMillAcName As String  '-- Source Mill A/C
Dim TMillAcCode As Long '-- To A/c -> Mill Ac Code

If ValidateData = False Then Exit Sub
Ans = MsgBox("Selected Items will be created for " & txtTCompany & " company. Are you sure?", vbYesNo, Me.Caption)
If Ans = vbNo Then
    Exit Sub
End If
Image1.Left = 300
Timer1.Enabled = True
Timer1_Timer
'-- Set Database Connection
SetSCompanyDatabase
SetTCompanyDatabase
'=================================  Creation ============================
Select Case frmTypeStr
    Case "Account" '-Account
        For i = 0 To lisTarget.ListCount - 1
            AcName = lisTarget.List(i)
            GProcRstOpen RsS, "Select * from TblMastAccount where AcName='" & AcName & "'", "R", tmpSCn
            GProcRstOpen RsT, "Select * from TblMastAccount where AcName='" & AcName & "'", "O", tmpTCn
            With RsS
                If .EOF And .BOF Then
                Else
                    '-- Get A/c Group,Broker,Party Group from Source
                    SAgName = GProcGetColumnValue("tblMastAccount,tblMastGroup", "AcName", AcName, "S", "AgName", "S", "tblMastAccount.AGCODE=tblMastGroup.AGCODE", tmpSCn)
                    SBrkAcName = GProcGetColumnValue("tblMastAccount,tblMastaccount BrkAccount", "tblMastAccount.AcName", AcName, "S", "BrkAccount.AcName", "S", "tblMastAccount.ACBRKCODE=BrkAccount.AcCode", tmpSCn)
                    SCgName = GProcGetColumnValue("tblMastAccount,tblmastcompgroup", "AcName", AcName, "S", "CGNAME", "S", "ACCMPCODE=CGCODE", tmpSCn)
                    '-- Get Code for target
                    '-- AgCode
                    TAgCode = GProcGetColumnValue("tblMastGroup", "AGName", SAgName, "S", "AgCode", "N", "", tmpTCn)
                    If TAgCode = 0 Then TAgCode = 90001
                    '-- Brk AcCode
                    TBrkAcCode = GProcGetColumnValue("tblMastAccount", "ACName", SBrkAcName, "S", "AcCode", "N", "", tmpTCn)
                    If TBrkAcCode = 0 Then TBrkAcCode = GProcGetColumnValue("tblMastAccount", "ACName", "-", "S", "AcCode", "N", "", tmpTCn)
                    '-- comp code
                    TCgCode = GProcGetColumnValue("tblmastcompgroup", "CGNAME", SCgName, "S", "CgCode", "N", "", tmpTCn)
                    If TCgCode = 0 Then TCgCode = GProcGetColumnValue("tblmastcompgroup", "CGNAME", "-", "S", "CgCode", "N", "", tmpTCn)
                    If RsT.EOF And RsT.BOF Then
                        RsT.AddNew
                        For j = 0 To .Fields.Count - 1
                            Select Case j
                                Case 0 '-- Aid
                                    RsT.Fields(j) = GProcGenerateId("tblmastaccount", "AcCode", "", tmpTCn)
                                Case 2, 3   '-- Gid
                                    RsT.Fields(j) = TAgCode
                                Case 16  '-- Party Group Code
                                    RsT.Fields(j) = TCgCode
                                Case 17  '-- Brk Ac Code
                                    RsT.Fields(j) = TBrkAcCode
                                Case Else
                                    RsT.Fields(j) = .Fields(j)
                            End Select
                        Next j
                       ' MsgBox i
                        RsT.Update
                    End If
                End If
            End With
        Next i
    Case "Item" '-Count
        For i = 0 To lisTarget.ListCount - 1
            j = InStr(1, lisTarget.List(i), "Count : ")
            ItName = Mid(lisTarget.List(i), j + 8)
            SMillAcName = Trim(Left(lisTarget.List(i), j - 1))
            TMillAcCode = GProcGetColumnValue("tblMastAccount", "AcName", SMillAcName, "S", "AcCode", "N", "", tmpTCn)
            '-- Get Mill Accode In To-Company
            If TMillAcCode = 0 Then
                MsgBox "Mill Account " & SMillAcName & " for " & txtTCompany & " company not Exist.Related Count will not get Create.", vbOKOnly + vbInformation, "Count Creation"
            Else
                ItCode = GProcGetColumnValue("tblMastItem,TblMastAccount", "ItName", ItName, "S", "ItCode", "N", "ITMILLCODE=AcCode and AcName='" & SMillAcName & "'", tmpSCn)
                '-- If Mill Item Not Exist then Create
                If GProcGetColumnValue("tblMastItem,TblMastAccount", "ItName", ItName, "S", "ItCode", "N", "ITMILLCODE=AcCode and AcName='" & SMillAcName & "'", tmpTCn) = 0 Then
                    GProcRstOpen RsS, "Select * from TblMastItem where Itcode=" & ItCode, "R", tmpSCn
                    GProcRstOpen RsT, "Select * from TblMastItem where Itcode=0", "O", tmpTCn
                    With RsS
                        RsT.AddNew
                        For j = 0 To .Fields.Count - 1
                            Select Case j
                                Case 0 '-- ItCode
                                    RsT.Fields(j) = GProcGenerateId("tblmastItem", "ItCode", "", tmpTCn)
                                Case 9   '-- Mill AcCode
                                    RsT.Fields(j) = TMillAcCode
                                Case Else
                                    RsT.Fields(j) = .Fields(j)
                            End Select
                        Next j
                        RsT.Update
                     End With
                End If
             End If
          Next i
    Case "Godown", "Narration", "Tax", "MillExp", "Transport", "Sizer" '-- Narration
        For i = 0 To lisTarget.ListCount - 1
            Narration = lisTarget.List(i)
            GProcRstOpen RsS, "Select * from TblMastNarration where Narration='" & Narration & "'", "R", tmpSCn
            GProcRstOpen RsT, "Select * from TblMastNarration where Narration='" & Narration & "'", "O", tmpTCn
            With RsS
                If .EOF And .BOF Then
                Else
                    '-- Pu,SA Tax AcCode
                    TDefaultAcCode = GProcGetColumnValue("tblMastAccount", "ACName", "-", "S", "AcCode", "N", "", tmpTCn)
                    If RsT.EOF And RsT.BOF Then
                        RsT.AddNew
                        For j = 0 To .Fields.Count - 1
                            Select Case j
                                Case 1 '-- Narr Code
                                    RsT.Fields(j) = GProcGenerateId("TblMastNarration", "NarrCode", "", tmpTCn)
                                Case 8, 9   '-- TaxAcCode
                                    RsT.Fields(j) = TDefaultAcCode
                                Case Else
                                    RsT.Fields(j) = .Fields(j)
                            End Select
                        Next j
                       ' MsgBox i
                        RsT.Update
                    End If
                End If
            End With
        Next i
End Select
Do While True
If Image1.Left > (9490 - 300) Then
    Timer1.Enabled = False
    DoEvents
    Exit Do
Else
    Timer1.Enabled = True
    Timer1_Timer
End If
Loop
MsgBox "Selected Items created for " & txtTCompany & " company successfully.", vbOKOnly
If RsS Is Nothing Then
Else
    If RsS.State = adStateOpen Then RsS.Close
End If
If RsT Is Nothing Then
Else
    If RsT.State = adStateOpen Then RsT.Close
End If
End Sub
'-- Set Source Company Db
Private Sub SetSCompanyDatabase()
   Dim DbStr As String
   DbStr = GProcGetColumnValue("TblmastCompany,tblCompanyYear", "CName", txtSCompany, "S", "Ccode + CompYear", "T", "Compcode=CCode and  CompFdt=cdate('" & gCYSDate & "')", gSetupCn)
   If DbStr <> "" Then
    If gBackEndDB = gBackEndAccess Then '-access
        GProcConnectionOpen tmpSCn, DbStr
    Else
        GProcConnectionOpenORA tmpSCn, DbStr
    End If
   End If
End Sub
'-- Set Target Company Db
Private Sub SetTCompanyDatabase()
   Dim DbStr As String
   DbStr = GProcGetColumnValue("TblmastCompany,tblCompanyYear", "CName", txtTCompany, "S", "Ccode + CompYear", "T", "Compcode=CCode and  CompFdt=cdate('" & gCYSDate & "')", gSetupCn)
   If DbStr <> "" Then
    If gBackEndDB = gBackEndAccess Then '-access
        GProcConnectionOpen tmpTCn, DbStr
    Else
        GProcConnectionOpenORA tmpTCn, DbStr
    End If
   End If
End Sub
Private Sub cmdDeselect_Click()
Dim i As Long
Dim j As Long
j = lisTarget.ListCount - 1
For i = 0 To j
    DSelectFromSource lisTarget.List(0)
    lisTarget.RemoveItem 0
Next i
End Sub
Private Sub cmdSelect_Click()
Dim i As Long
lisTarget.Clear
For i = 0 To lisSource.ListCount - 1
    lisSource.Selected(i) = True
Next i
End Sub
Private Sub CmdClose_Click()
    Unload Me
End Sub
Private Function ValidateData() As Boolean
'-- Company Name
If GProcGetColumnValue("TblmastCompany", "CName", txtSCompany, "S", "Ccode", "N", "", gSetupCn) = 0 Then
    MsgBox "Check From Company Name.", vbInformation + vbOKOnly, Me.Caption
    txtSCompany.SetFocus
    Exit Function
End If
If GProcGetColumnValue("TblmastCompany", "CName", txtTCompany, "S", "Ccode", "N", "", gSetupCn) = 0 Then
    MsgBox "Check To Company Name.", vbInformation + vbOKOnly, Me.Caption
    txtTCompany.SetFocus
    Exit Function
End If
If txtSCompany = txtTCompany Then
    MsgBox "Check To Company Name.", vbInformation + vbOKOnly, Me.Caption
    txtTCompany.SetFocus
    Exit Function
End If
'-- Item
If lisTarget.ListCount = 0 Then
    MsgBox "Select Item", vbInformation + vbOKOnly, Me.Caption
    Exit Function
End If
ValidateData = True
End Function

Private Sub txtSCompany_KeyPress(KeyAscii As Integer)
Dim LeftPos As Long
Dim TopPos As Long
If KeyAscii = 13 Then Exit Sub
TopPos = txtSCompany.Top + Me.Top + 360
LeftPos = txtSCompany.Left + Me.Left
'-- Company Name
gClsSearch.SearchMultiField "tblMastcompany", "CName", Array("Company Name"), Array(txtSCompany.Width), "", Chr(KeyAscii), "CName", LeftPos, TopPos, False, gSetupCn
KeyAscii = 0
If gClsSearch.SearchMultiRetCol(0) <> "" Then txtSCompany = gClsSearch.SearchMultiRetCol(0)
FillItem
End Sub
Private Sub FillItem()
Dim Rs  As Recordset
Dim X As String
SetSCompanyDatabase
Me.Caption = frmTypeStr + " " + Me.Caption
Select Case frmTypeStr
    Case "Account" '-Account
        X = "select AcName from tblMastAccount order by Acname"
    Case "Item" '-Count
        If gBackEndDB = gBackEndAccess Then
            X = "select ACName + '   Count : ' + ItName from tblMastItem,tblMastAccount where ITMILLCODE =AcCode order by AcName,ItName"
        Else 'Oracle
            X = "select ACName || '   Count : ' || ItName from tblMastItem,tblMastAccount where ITMILLCODE =AcCode order by AcName,ItName"
        End If
    Case "Godown" '-Godown
        X = "Select Narration from tblmastnarration where NarrType='G' order by Narration"
    Case "Narration" '-Narration
        X = "Select Narration from tblmastnarration where NarrType='N' order by Narration"
    Case "Tax" '-Tax
        X = "Select Narration from tblmastnarration where NarrType='T' order by Narration"
    Case "MillExp" '-MillExp
        X = "Select Narration from tblmastnarration where NarrType='M' order by Narration"
    Case "Transport" '-Transport
        X = "Select Narration from tblmastnarration where NarrType='R' order by Narration"
    Case "Sizer" '-Sizer
        X = "Select Narration from tblmastnarration where NarrType='S' order by Narration"
End Select
If txtSCompany <> "" Then
    If GProcRstOpen(Rs, X, "R", tmpSCn) > 0 Then
        With Rs
            .MoveFirst
            lisSource.Clear
            lisTarget.Clear
            Do While Not .EOF
                lisSource.AddItem .Fields(0)
                .MoveNext
            Loop
            .Close
        End With
    End If
End If
End Sub
Private Sub txtTCompany_KeyPress(KeyAscii As Integer)
Dim LeftPos As Long
Dim TopPos As Long
If KeyAscii = 13 Then Exit Sub
TopPos = txtTCompany.Top + Me.Top + 360
LeftPos = txtTCompany.Left + Me.Left
'-- Company Name
gClsSearch.SearchMultiField "tblMastcompany", "CName", Array("Company Name"), Array(txtTCompany.Width), "Cname<>'" & txtSCompany & "'", Chr(KeyAscii), "CName", LeftPos, TopPos, False, gSetupCn
KeyAscii = 0
If gClsSearch.SearchMultiRetCol(0) <> "" Then txtTCompany = gClsSearch.SearchMultiRetCol(0)
End Sub
Private Sub txtSCompany_GotFocus()
GProcSelectBox Me.ActiveControl
End Sub
Private Sub txtTCompany_GotFocus()
    GProcSelectBox Me.ActiveControl
End Sub
Private Function IsInTarget(Item As String) As Boolean
Dim i As Long
IsInTarget = False
For i = 0 To lisTarget.ListCount - 1
    If lisTarget.List(i) = Item Then
       IsInTarget = True
       Exit Function
    End If
Next i
End Function
Private Sub lisSource_Click()
Dim i As Long
Dim Sstr As String
Dim j As Long
Dim isFoundFlag As Boolean
Sstr = lisSource.Text
'lisTarget.Clear
For i = 0 To lisSource.ListCount - 1
        If IsInTarget(lisSource.List(i)) = False Then
            If lisSource.Selected(i) = True Then
                lisTarget.AddItem lisSource.List(i)
            End If
        End If
Next i
End Sub
Private Sub lisTarget_Click()
Dim rsTR As String
Dim i As Long
rsTR = lisTarget.Text
DSelectFromSource rsTR
For i = 0 To lisTarget.ListCount - 1
    If lisTarget.List(i) = rsTR Then
        lisTarget.RemoveItem i
    End If
Next i
End Sub
Private Sub Timer1_Timer()
If Image1.Left < 9640 Then
    Image1.Left = Image1.Left + 50
Else
    Image1.Left = 300
End If
End Sub
Private Sub DSelectFromSource(Item As String)
Dim i As Long
For i = 0 To lisSource.ListCount - 1
    If Item = lisSource.List(i) Then
        lisSource.Selected(i) = False
    End If
Next i
End Sub

