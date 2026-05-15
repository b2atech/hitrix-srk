Attribute VB_Name = "AccountRptProc"
Option Explicit
Public Sub GProcPrepareLedger(ReportRecStrSub As String, FromDt As Date, ToDt As Date)
Dim Balance As String
Dim RsT As Recordset
Dim x As String
Dim AcName As String
Dim OldAcName As String
gCn.Execute "delete from tmpAccLedger where UserName='" & gUserName & "'"
Select Case ReportRecStrSub
    Case "All" '-- All
        gCn.Execute "insert into tmpAccLedger select *,'" & gUserName & "' as username from QryAccLedger " _
        & " where Vtype<>'OP' and Vdt>=Cdate('" & FromDt & "') and vdt<=cdate('" & ToDt & "')"
    Case "Selected Account" '-- Selected Account
        gCn.Execute "insert into tmpAccLedger select *,'" & gUserName & "' as username from QryAccLedger " _
         & " where Vtype<>'OP' and Vdt>=Cdate('" & FromDt & "') and vdt<=cdate('" & ToDt & "') and AcName in ('" & gSearchStr & "')"
    Case "Selected Group" '-- Selected Group
        gCn.Execute "insert into tmpAccLedger select *,'" & gUserName & "' as username from QryAccLedger " _
        & " where Vtype<>'OP' and Vdt>=Cdate('" & FromDt & "') and vdt<=cdate('" & ToDt & "') and AgName in ('" & gSearchStr & "')"
End Select
x = "Select * from tmpAccLedger where UserName='" & gUserName & "' and vtype<>'OP' order by AcName,Vdt,orderno,Vtype,Vno,VSrno"
GProcRstOpen RsT, x, "O"
With RsT
    If .EOF And .BOF Then
        Exit Sub
    End If
    .MoveFirst
    OldAcName = ""
    AcName = ""
    Do While Not .EOF
        'If OldAcName = "" Then OldAcName = AcName
        AcName = .Fields("acName")
        If OldAcName <> AcName Then '--Check Distinct AName
            Balance = 0
            OldAcName = AcName
            If FromDt = gCYSDate Then
                Balance = GProcGetAccountBalance(.Fields("AcName"), FromDt, True)
            Else
                Balance = GProcGetAccountBalance(.Fields("AcName"), FromDt - 1)
            End If
            If Balance <> 0 Then
                gCn.Execute "insert into tmpAccLedger(Vtype,Vdt,Balance,AcName,AcOName, " _
                & " Balance,AgName,UserName) values ('OP'," & FromDt & "," & Balance & ",'" & .Fields("acName") & "'," _
                & "'Opening Balance'," & Balance & ",'" & .Fields("AgName") & "," & gUserName & "')"
            End If
         End If
         DoEvents
         Balance = Balance + .Fields("VAmt")
        .Fields("balance") = Balance
        .Update
        .MoveNext
    Loop
End With
End Sub


