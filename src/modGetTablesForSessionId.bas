Attribute VB_Name = "modGetTablesForSessionId"
Option Explicit

Sub GetTablesForSessionId()
    Dim Qry As QueryTable
    Dim wbkReport As Workbook
    Dim shtNew As Worksheet
    Dim arrSqlQueries() As Variant
    Dim intSqlQueryIndex As Integer
    Dim strSqlQuery As String
    Dim strSessionId As String
    Dim msgBoxResult As VbMsgBoxResult
    
    msgBoxResult = vbYes
    Do While msgBoxResult = vbYes
        If Not isLANCableconnected Then
            msgBoxResult = MsgBox( _
                "You must be connected via a LAN cable from the IBM Office." & vbNewLine & _
                "Would you like to retry?", vbYesNo, "Session ID Analysis" _
            )
            If msgBoxResult = vbNo Then Exit Sub
        Else
            Exit Do
        End If
    Loop
    
    strSessionId = Application.InputBox("Type the Session ID to extract Redshift data: ", "Session ID Analysis", "00000149470725811461", Type:=2)
    If strSessionId = "False" Or Trim(strSessionId) = "" Then
        Exit Sub
    End If
    
    Application.ScreenUpdating = False
    
    arrSqlQueries = Array( _
        "select * from ing.sessions_info where session_id = '" & strSessionId & "' limit 1000;", _
        "select * from ing.policy_results where session_id = '" & strSessionId & "' limit 1000;", _
        "select * from ing.policy_invocation_stats where session_id = '" & strSessionId & "' limit 1000;" _
    )
    
    Set wbkReport = Workbooks.Add(xlWBATWorksheet)
    For intSqlQueryIndex = 0 To UBound(arrSqlQueries)
        strSqlQuery = arrSqlQueries(intSqlQueryIndex)
        If intSqlQueryIndex > 0 Then
            Set shtNew = wbkReport.Worksheets.Add(count:=1, After:=Worksheets(Worksheets.count))
            Select Case intSqlQueryIndex
            Case 1:
                shtNew.Name = "policy_results"
            Case 2:
                shtNew.Name = "policy_invocation_stats"
            End Select
        Else
            Set shtNew = wbkReport.ActiveSheet
            shtNew.Name = "sessions_info"
        End If
        Set Qry = CreateQueryTable(shtNew)
        With Qry
            .CommandText = strSqlQuery
            .AdjustColumnWidth = True
            .Refresh BackgroundQuery:=False
        End With
        Call FormatDateColumns(shtNew)
    Next intSqlQueryIndex
    Erase arrSqlQueries
    Application.ScreenUpdating = True
End Sub

