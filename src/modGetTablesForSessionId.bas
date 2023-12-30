Attribute VB_Name = "modGetTablesForSessionId"
Option Explicit

Sub GetTablesForSessionId()
Attribute GetTablesForSessionId.VB_ProcData.VB_Invoke_Func = "S\n14"
    Dim Qry As QueryTable
    Dim wbkReport As Workbook
    Dim shtNew As Worksheet
    Dim arrSqlQueries() As Variant
    Dim intSqlQueryIndex As Integer
    Dim strSqlQuery As String
    Dim strSessionId As String
    Dim strOrigin As String
    Dim strDSN As String
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
    
'    strDSN = Application.InputBox("Type the DSN to extract Redshift data: ", "Session ID Analysis", "Redshift_US", Type:=2)
    strDSN = "Redshift_EU"
    If strDSN = "False" Or Trim(strDSN) = "" Then
        Exit Sub
    End If
    
'    strOrigin = Application.InputBox("Type the Origin to extract Redshift data: ", "Session ID Analysis", "anz", Type:=2)
    strOrigin = "lgt"
    If strOrigin = "False" Or Trim(strOrigin) = "" Then
        Exit Sub
    End If
    
    strSessionId = Application.InputBox("Type the Session ID to extract Redshift data: ", "Session ID Analysis", "00000149470725811461", Type:=2)
    If strSessionId = "False" Or Trim(strSessionId) = "" Then
        Exit Sub
    End If
    
    Application.ScreenUpdating = False
    
'    arrSqlQueries = Array( _
        "select * from " & strOrigin & ".policy_invocation_stats where session_id = '" & strSessionId & "' limit 10000;" _
    )
    
    arrSqlQueries = Array( _
        "select * from " & strOrigin & ".sessions_info where session_id = '" & strSessionId & "' limit 10000;", _
        "select * from " & strOrigin & ".policy_results where session_id = '" & strSessionId & "' limit 10000;", _
        "select * from " & strOrigin & ".devices where session_id = '" & strSessionId & "' limit 10000;", _
        "select * from " & strOrigin & ".policy_invocation_stats where session_id = '" & strSessionId & "' limit 10000;" _
    )

'    arrSqlQueries = Array( _
        "select * from " & strOrigin & ".sessions_info where session_id = '" & strSessionId & "' limit 10000;", _
        "select * from " & strOrigin & ".policy_results where session_id = '" & strSessionId & "' limit 10000;", _
        "select * from " & strOrigin & ".policy_invocation_stats where session_id = '" & strSessionId & "' limit 10000;" _
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
        Set Qry = CreateQueryTable(shtNew, strDSN)
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

