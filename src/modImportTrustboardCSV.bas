Attribute VB_Name = "modImportTrustboardCSV"
Option Explicit

Function getNewestFileFullName(ByRef LatestFile As String) As String
    Dim MyPath As String
    Dim MyFile As String
    'Dim LatestFile As String
    Dim LatestDate As Date
    Dim LMD As Date
    
    MyPath = GetDownloadsPath()
    If Right(MyPath, 1) <> Application.PathSeparator Then MyPath = MyPath & Application.PathSeparator
        MyFile = Dir(MyPath & "*.csv", vbNormal)
        If Len(MyFile) = 0 Then
            MsgBox "No files were found...", vbExclamation
        Exit Function
    End If
    
    Do While Len(MyFile) > 0
        On Error Resume Next
        LMD = FileDateTime(MyPath & MyFile)
        If Err.Number <> 0 Then
            Err.Clear
            GoTo nextFile
        End If
        If LMD > LatestDate Then
            LatestFile = MyFile
            LatestDate = LMD
        End If
nextFile:
        MyFile = Dir
    Loop
    
    getNewestFileFullName = MyPath & LatestFile
End Function

Sub ImportTrustboardCSV()
Attribute ImportTrustboardCSV.VB_ProcData.VB_Invoke_Func = "I\n14"
    Dim wbk As Workbook
    Dim Sht As Worksheet
    Dim strNewestFileName As String
    Dim strNewestFileFullName As String
    Dim strPowerQueryFormula As String
    
    strNewestFileFullName = getNewestFileFullName(strNewestFileName)
    'If Left(strNewestFileName, 4) = "guid" Then 'commented out since TB changed PUID export file name around December 14 2023
    If True Then
        strPowerQueryFormula = _
            "let Source = Csv.Document(File.Contents(""" & strNewestFileFullName & """),[Delimiter="","", Encoding=65001, QuoteStyle=QuoteStyle.None]), #""Promoted Headers"" = Table.PromoteHeaders(Source, [PromoteAllScalars=true]), " & _
            "#""Changed Type"" = Table.TransformColumnTypes(#""Promoted Headers"",{{""Account Id"", type text}, {""Application"", type text}, {""Browser"", type text}, {""Browser version"", type text}, {""Classification"", type text}, {""Client Language"", type text}, {""Line Carrier"", type text}, {""Country code"", type text}, {""Date & time"", type datetime}, {""Customer session IDs"", type text}, " & _
            "{""Device ID"", type text}, {""Encrypted user ID"", type text}, {""City"", type text}, {""Country"", type text}, {""ISP"", type text}, {""IP"", type text}, {""Name"", type text}, {""Machine ID"", type text}, {""Malware Name"", type text}, {""Infected App"", type text}, {""Infected Package"", type text}, {""OS"", type text}, {""Pinpoint session ID"", type text}, " & _
            "{""Platform"", type text}, {""PUID"", type text}, {""Assessment Details"", type text}, {""Recommendation"", type text}, {""Partial result reason"", type text}, {""Reason ID"", Int64.Type}, {""Reason"", type text}, {""Risk score"", Int64.Type}, {""Classified By"", type text}, {""Status"", type text}, {""Classified At"", type text}, {""New Device"", type logical}, " & _
            "{""Activity"", type text}, {""Closed By"", type text}, {""Closed At"", type text}, {""User Agent"", type text}, {""Assigned To"", type text}, {""Phishing Url"", type text}, {""Detected At"", type text}, {""SDK Configuration"", type text}, {""SDK Version"", type text}, {""MRST App Count"", type text}, {""Call In Progress"", type text}, {""User Behavioral Score"", type text}, " & _
            "{""Risky Device"", type text}, {""Risky Connection"", type text}, {""Battery Charging"", type text}, {""Behavioral Anomaly"", type text}, {""First Seen In Account"", type text}, {""First Seen In Region"", type text}, {""Fraud MO"", type text}, {""Agent Key"", type text}, {""Marketing Name"", type text}, {""Channel"", type text}, {""Transaction Amount"", type text}, " & _
            "{""Is Alerted"", type logical}}) in #""Changed Type"""
    ElseIf Left(strNewestFileName, 2) = "20" Then
    'ElseIf True Then
        strPowerQueryFormula = _
            "let Source = Csv.Document(File.Contents(""" & strNewestFileFullName & """),[Delimiter="","", QuoteStyle=QuoteStyle.None]), #""Promoted Headers"" = Table.PromoteHeaders(Source, [PromoteAllScalars=true]), " & _
            "#""Changed Type"" = Table.TransformColumnTypes(#""Promoted Headers"",{{""Account Id"", type text}, {""Application"", type text}, {""Browser"", type text}, {""Browser version"", type text}, {""Classification"", type text}, " & _
            "{""Client Language"", type text}, {""Line Carrier"", type text}, {""Country code"", type text}, {""Date & time"", type datetime}, {""Customer session IDs"", type text}, {""Device ID"", type text}, {""Encrypted user ID"", type text}, {""City"", type text}, " & _
            "{""Country"", type text}, {""ISP"", type text}, {""IP"", type text}, {""Name"", type text}, {""Machine ID"", type text}, {""Malware Name"", type text}, {""Infected App"", type text}, {""Infected Package"", type text}, {""OS"", type text}, " & _
            "{""Pinpoint session ID"", type text}, {""Platform"", type text}, {""PUID"", type text}, {""Assessment Details"", type text}, {""Recommendation"", type text}, {""Partial result reason"", type text}, {""Reason ID"", Int64.Type}, {""Reason"", type text}, " & _
            "{""Risk score"", Int64.Type}, {""Classified By"", type text}, {""Status"", type text}, {""Classified At"", type text}, {""New Device"", type logical}, {""Activity"", type text}, {""Closed By"", type text}, {""Closed At"", type text}, " & _
            "{""User Agent"", type text}, {""Assigned To"", type text}, {""Phishing Url"", type text}, {""Detected At"", type text}, {""SDK Configuration"", Int64.Type}, {""SDK Version"", type text}, {""MRST App Count"", Int64.Type}, {""Call In Progress"", type text}, " & _
            "{""User Behavioral Score"", type text}, {""Risky Device"", type logical}, {""Risky Connection"", type logical}, {""Battery Charging"", type logical}, {""Behavioral Anomaly"", type text}, {""First Seen In Account"", type datetime}, " & _
            "{""First Seen In Region"", type datetime}, {""Fraud MO"", type text}, {""Agent Key"", type text}, {""Marketing Name"", type text}, {""Channel"", type text}, {""Transaction Amount"", type text}}) in #""Changed Type"""
    End If
    Set wbk = Workbooks.Add(xlWBATWorksheet)
'    ActiveWorkbook.Worksheets.Add
    wbk.Queries.Add Name:= _
        "TrustboardCSV" _
        , Formula:=strPowerQueryFormula
    Set Sht = wbk.ActiveSheet
    With Sht.ListObjects.Add(SourceType:=0, Source:="OLEDB;Provider=Microsoft.Mashup.OleDb.1;Data Source=$Workbook$;Location=TrustboardCSV;Extended Properties=""""", Destination:=Range("$A$1")).QueryTable
        .CommandType = xlCmdSql
        .CommandText = "SELECT * FROM TrustboardCSV"
        .Refresh BackgroundQuery:=False
    End With
    Application.CommandBars("Queries and Connections").Visible = False
End Sub


