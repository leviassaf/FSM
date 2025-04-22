Attribute VB_Name = "modRiskReasonsPrecisionHandling"
Option Explicit

'  ********* Columns *********
Private Const TB_COLUMN_EVENT_DATE As String = "Date & time"
Private Const TB_COLUMN_ACTIVITY_CONFIRMATION As String = "Classification"
Private Const TB_COLUMN_RISK_REASON_ID As String = "Reason ID"
Private Const TB_COLUMN_RISK_REASON As String = "Reason"
Private Const TB_COLUMN_RISK_SCORE As String = "Risk score"
Private Const TB_COLUMN_APPLICATION As String = "Application"
Private Const TB_COLUMN_PUID As String = "PUID"
Private Const TB_COLUMN_SESSION As String = "Pinpoint session ID"
Private Const TB_COLUMN_ACTIVITY As String = "Activity"

Private Const TB_CLASSIFICATION_CONFIRMED_FRAUD As String = "confirmed_fraud"
Private Const TB_CLASSIFICATION_CONFIRMED_LEGITIMATE As String = "confirmed_legitimate"
Private Const TB_CLASSIFICATION_UNDETERMINED As String = "undetermined"
Private Const TB_CLASSIFICATION_PENDING As String = "pending_confirmation"
'  ********* End of TB Values *********

Private StrColumnEventDate As String
Private StrColumnActivityConfirmation As String
Private StrColumnRiskReason As String
Private StrColumnRiskReasonId As String
Private StrColumnRiskScore As String
Private StrColumnApplication As String
Private StrColumnPuid As String
Private StrColumnSession As String
Private StrColumnActivity As String

Private strClassificationConfirmedFraud As String
Private strClassificationConfirmedLegitimate As String
Private strClassificationUndetermined As String
Private strClassificationPending As String

Public Sub RiskReasonsPrecision()
Attribute RiskReasonsPrecision.VB_ProcData.VB_Invoke_Func = "R\n14"
    Const REPORT_NAME As String = "Pivot Table"
    Dim shtRawData As Worksheet
    Dim Pvt As PivotTable
    
    Dim Wbk As Workbook
    
    Set Wbk = Workbooks.Add(xlWBATWorksheet)
    Set shtRawData = Wbk.ActiveSheet
    Call importData(Wbk)

    Call prepareData(shtRawData)
    Call DeleteIrrelevantRecords(shtRawData, "Reason ID", "=")

        Set Pvt = CreatePivotReport(shtRawData, REPORT_NAME)

    Call createChannelReports(Pvt)
    AppActivate Application.Caption
    
    Application.ScreenUpdating = True
    Set Wbk = Nothing
End Sub

Private Sub importData(Wbk As Workbook)
    Dim shtRawData As Worksheet
    Dim strBoxPath As String
    Dim strDetectionRateFolderPath As String
    Dim intNumberOfSourceFiles As Integer
    Dim strQueryString As String
    
    strDetectionRateFolderPath = "C:\Users\919561756\Box\Trusteer\Reporting\VBA Projects\FP Monitoring\LCL\February 2025\login"
    'strDetectionRateFolderPath = "C:\Users\919561756\Box\Trusteer\Reporting\VBA Projects\FP Monitoring\absa\Add Payee"

    If strDetectionRateFolderPath = "False" Then Exit Sub
    Application.ScreenUpdating = False
    intNumberOfSourceFiles = CountFilesInFolder(strDetectionRateFolderPath)
    
    If intNumberOfSourceFiles = 1 Then
        strQueryString = "let" & Chr(13) & "" & Chr(10) & "    Source = Csv.Document(File.Contents(""C:\Users\919561756\Box\Trusteer\Reporting\VBA Projects\FP Monitoring\Cagricole\June 2024\2024-07-10T19-14-44-risks.csv""),[Delimiter="","", Columns=65, Encoding=65001, QuoteStyle=QuoteStyle.None])," & Chr(13) & "" & Chr(10) & "    #""Promoted Headers"" = Table.PromoteHeaders(Source, [PromoteAllScalars=true])," & Chr(13) & "" & Chr(10) & "    #""Changed Type"" = Table.Transfo" & _
        "rmColumnTypes(#""Promoted Headers"",{{""Account Id"", type text}, {""Application"", type text}, {""Browser"", type text}, {""Browser version"", type text}, {""Classification"", type text}, {""Client Language"", type text}, {""Line Carrier"", type text}, {""Country code"", type text}, {""Date & time"", type datetimezone}, {""Customer session IDs"", type text}, {""Dev" & _
        "ice ID"", type text}, {""Encrypted user ID"", type text}, {""City"", type text}, {""Country"", type text}, {""ISP"", type text}, {""IP address"", type text}, {""Name"", type text}, {""Machine ID"", type text}, {""Malware Name"", type text}, {""Infected App"", type text}, {""Infected Package"", type text}, {""OS"", type text}, {""Pinpoint session ID"", type text}, {" & _
        """Platform"", type text}, {""PUID"", type text}, {""Assessment Details"", type text}, {""Recommendation"", type text}, {""Partial result reason"", type text}, {""Reason ID"", Int64.Type}, {""Detailed reason"", type text}, {""Risk score"", Int64.Type}, {""Classified By"", type text}, {""Status"", type text}, {""Classified At"", type datetimezone}, {""New Device"", ty" & _
        "pe logical}, {""Activity"", type text}, {""Closed By"", type text}, {""Closed At"", type datetimezone}, {""User Agent"", type text}, {""Assigned To"", type text}, {""Phishing URL"", type text}, {""Detected At"", type text}, {""SDK Configuration"", Int64.Type}, {""SDK Version"", type text}, {""MRST App Count"", Int64.Type}, {""Call In Progress"", type text}, {""User " & _
        "Behavioral Score"", type text}, {""Risky Device"", type logical}, {""Risky Connection"", type logical}, {""Battery Charging"", type logical}, {""Behavioral Anomaly"", type logical}, {""Device First Seen In Account"", type datetimezone}, {""Device First Seen In Region"", type datetimezone}, {""Fraud MO"", type text}, {""Agent Key"", type text}, {""Marketing Name"", t" & _
        "ype text}, {""Channel"", type text}, {""Transaction Amount"", type text}, {""GDID PUID Count Until Session"", Int64.Type}, {""Credentials submitted"", type logical}, {""Reason"", type text}, {""Device language"", type text}, {""Known risky payee"", type text}, {""New location"", type logical}, {""Transaction type"", type text}})" & Chr(13) & "" & Chr(10) & "in" & Chr(13) & "" & Chr(10) & "    #""Changed Type"""
    
        ActiveWorkbook.Worksheets.Add
        With ActiveSheet.ListObjects.Add(SourceType:=0, Source:= _
            "OLEDB;Provider=Microsoft.Mashup.OleDb.1;Data Source=$Workbook$;Location=2024-07-10T19-14-44-risks;Extended Properties=""""" _
            , Destination:=Range("$A$1")).queryTable
            .CommandType = xlCmdSql
            .CommandText = Array("SELECT * FROM [2024-07-10T19-14-44-risks]")
            .RowNumbers = False
            .FillAdjacentFormulas = False
            .PreserveFormatting = True
            .RefreshOnFileOpen = False
            .BackgroundQuery = True
            .RefreshStyle = xlInsertDeleteCells
            .SavePassword = False
            .SaveData = True
            .AdjustColumnWidth = True
            .RefreshPeriod = 0
            .PreserveColumnInfo = True
            .listObject.displayName = "foo"
            .Refresh BackgroundQuery:=False
        End With
    
    ElseIf intNumberOfSourceFiles > 1 Then
        strQueryString = "let" & Chr(13) & "" & Chr(10) & "    Source = Folder.Files(""" & strDetectionRateFolderPath & """)," & Chr(13) & "" & Chr(10) & "    #""Filtered Hidden Files1"" = Table.SelectRows(Source, each [Attributes]?[Hidden]? <> true)," & Chr(13) & "" & Chr(10) & "    #""Invoke Custom Function1"" = Table.AddColumn(#""Filtered Hidden Files1"", ""Transform File"", each #""Transform File""([Content]))," & Chr(13) & "" & Chr(10) & "    #""Renamed Colum" & _
            "ns1"" = Table.RenameColumns(#""Invoke Custom Function1"", {""Name"", ""Source.Name""})," & Chr(13) & "" & Chr(10) & "    #""Removed Other Columns1"" = Table.SelectColumns(#""Renamed Columns1"", {""Source.Name"", ""Transform File""})," & Chr(13) & "" & Chr(10) & "    #""Expanded Table Column1"" = Table.ExpandTableColumn(#""Removed Other Columns1"", ""Transform File"", Table.ColumnNames(#""Transform File""(#""Sample File""" & _
            ")))," & Chr(13) & "" & Chr(10) & "    #""Changed Type"" = Table.TransformColumnTypes(#""Expanded Table Column1"",{{""Source.Name"", type text}, {""Account Id"", type text}, {""Application"", type text}, {""Browser"", type text}, {""Browser version"", type text}, {""Classification"", type text}, {""Client Language"", type text}, {""Line Carrier"", type text}, {""Country code"", type text}, {""D" & _
            "ate & time"", type datetimezone}, {""Customer session IDs"", type text}, {""Device ID"", type text}, {""Encrypted user ID"", type text}, {""City"", type text}, {""Country"", type text}, {""ISP"", type text}, {""IP address"", type text}, {""Name"", type text}, {""Machine ID"", type text}, {""Malware Name"", type text}, {""Infected App"", type text}, {""Infected Packa" & _
            "ge"", type text}, {""OS"", type text}, {""Pinpoint session ID"", type text}, {""Platform"", type text}, {""PUID"", type text}, {""Assessment Details"", type text}, {""Recommendation"", type text}, {""Partial result reason"", type any}, {""Reason ID"", Int64.Type}, {""Detailed reason"", type text}, {""Risk score"", Int64.Type}, {""Classified By"", type text}, {""Stat" & _
            "us"", type text}, {""Classified At"", type datetime}, {""New Device"", type logical}, {""Activity"", type text}, {""Closed By"", type any}, {""Closed At"", type any}, {""User Agent"", type text}, {""Assigned To"", type text}, {""Phishing URL"", type any}, {""Detected At"", type text}, {""SDK Configuration"", Int64.Type}, {""SDK Version"", type text}, {""MRST App Cou" & _
            "nt"", Int64.Type}, {""Call In Progress"", type text}, {""User Behavioral Score"", type any}, {""Risky Device"", type logical}, {""Risky Connection"", type logical}, {""Battery Charging"", type logical}, {""Behavioral Anomaly"", type logical}, {""Device First Seen In Account"", type datetimezone}, {""Device First Seen In Region"", type datetimezone}, {""Fraud MO"", t" & _
            "ype text}, {""Agent Key"", type text}, {""Marketing Name"", type text}, {""Channel"", type text}, {""Transaction Amount"", type text}, {""GDID PUID Count Until Session"", Int64.Type}, {""Credentials submitted"", type logical}, {""Reason"", type text}, {""Device language"", type text}, {""Known risky payee"", type any}, {""New location"", type logical}, {""Transactio" & _
            "n type"", type text}})" & Chr(13) & "" & Chr(10) & "in" & Chr(13) & "" & Chr(10) & "    #""Changed Type"""
        
        With Wbk
            With .Queries
                If intNumberOfSourceFiles > 1 Then 'if more than 1 source file was found
                    .Add name:="foo report name", _
                        formula:=strQueryString
                    .Add name:="Sample File", formula:= _
                        "let Source = Folder.Files(""" & strDetectionRateFolderPath & """), Navigation1 = Source{0}[Content] in Navigation1"
                    .Add name:="Parameter1", formula:= _
                        "#""Sample File"" meta [IsParameterQuery=true, BinaryIdentifier=#""Sample File"", Type=""Binary"", IsParameterQueryRequired=true]"
                    .Add name:="Transform Sample File", formula:= _
                        "let Source = Csv.Document(Parameter1,[Delimiter="","", QuoteStyle=QuoteStyle.None]), #""Promoted Headers"" = Table.PromoteHeaders(Source, [PromoteAllScalars=true]) in #""Promoted Headers"""
                    .Add name:="Transform File", formula:= _
                        "let Source = (Parameter1) => let Source = Csv.Document(Parameter1,[Delimiter="","", QuoteStyle=QuoteStyle.None]), #""Promoted Headers"" = Table.PromoteHeaders(Source, [PromoteAllScalars=true]) in #""Promoted Headers"" in Source"
                Else
                    MsgBox "Adjust VBA code to handle importing a single source file"
                End If
            End With
            Set shtRawData = .ActiveSheet
            With shtRawData
                With .ListObjects.Add(SourceType:=0, Source:= _
                "OLEDB;Provider=Microsoft.Mashup.OleDb.1;Data Source=$Workbook$;Location=""foo report name"";Extended Properties=""""" _
                , Destination:=Range("$A$1")).queryTable
                .CommandType = xlCmdSql
                .CommandText = Array("SELECT * FROM [foo report name]")
                .Refresh BackgroundQuery:=False
                End With
                .name = "Raw Data"
            End With
        End With
    End If
End Sub

Private Sub createChannelReports(Pvt As PivotTable)
    Dim shtReport As Worksheet
    Dim pvtField As PivotField
    Dim strChannelValue As String
    Dim pvtItem As PivotItem
    
    With Pvt
        Set shtReport = ConvertPivotToTable(Pvt)
        
        Call AddWorksheetEventCode(ActiveWorkbook, shtReport, "FollowHyperlink")
      
        Set pvtField = .PivotFields("Channel")
        For Each pvtItem In pvtField.PivotItems
            .PivotFields("Channel").CurrentPage = pvtItem.name
            Set shtReport = ConvertPivotToTable(Pvt)
            
            Call AddWorksheetEventCode(ActiveWorkbook, shtReport, "FollowHyperlink")
        Next pvtItem
      
        Set pvtField = .PivotFields("Channel")
        With pvtField
            .Orientation = xlRowField
            .Position = 1
        End With
    End With
End Sub

Private Sub createHyperlinks(Sht As Worksheet, HyperlinksRange As Range)
    Dim Cel As Range
    
    For Each Cel In HyperlinksRange
        Sht.Hyperlinks.Add Anchor:=Cel, Address:="", SubAddress:="'" & Sht.name & "'!" & Cel.Address(False, False, xlA1)
    Next Cel
End Sub

Private Function ConvertPivotToTable(Pvt As PivotTable) As Worksheet
    Dim shtReport As Worksheet
    Dim rngSort As Range
    Dim rngSortKey As Range
    Dim lngLastRowIndex As Long
    Dim rngDataBody As Range
    
    Dim rngTargets As Range
    Dim rngMonths As Range
    Dim rngCell As Range
    
    Set shtReport = Worksheets.Add
    
    With shtReport
        'Convert values from pivot table to the new worksheet
        Call CopyValues(Pvt.TableRange1, Destination:=.Cells(1))
        .Rows(1).Delete
        
        Range("L1:N1").Value2 = "Grand Total"
        Range("O1:Q1").Value2 = "Precision %"
        Range("R1:T1").Value2 = "Handling %"
        Range("U1:W1").Value2 = "Fraud Distribution %"
        Range("X1:Z1").Value2 = "Alert Distribution %"
        
        Set rngMonths = Range("D2:E2")
        Set rngTargets = Range("L2:M2, O2:P2, R2:S2, U2:V2, X2:Y2")
        
        For Each rngCell In rngTargets.Areas
            rngCell.Value2 = rngMonths.Value2
        Next rngCell

        Range("N2, Q2, T2, W2, Z2").Value2 = "Evol"
        
        Application.Intersect(.UsedRange, Range("L:M")).SpecialCells(xlCellTypeBlanks).FormulaR1C1 = "=RC[-8]+RC[-6]+RC[-4]+RC[-2]"
        Application.Intersect(.UsedRange, Range("O:P")).SpecialCells(xlCellTypeBlanks).FormulaR1C1 = "=IFERROR(ROUND(RC[-11]/(RC[-11]+RC[-9]),2),0)"
        Application.Intersect(.UsedRange, Range("R:S")).SpecialCells(xlCellTypeBlanks).FormulaR1C1 = "=IFERROR(ROUND((RC[-14]+RC[-12])/RC[-6],2),0)"
        lngLastRowIndex = .Cells.SpecialCells(xlCellTypeLastCell).Row
        Application.Intersect(.UsedRange, Range("U:V")).SpecialCells(xlCellTypeBlanks).FormulaR1C1 = "=RC[-17]/R" & lngLastRowIndex & "C[-17]"
        Application.Intersect(.UsedRange, Range("X:Y")).SpecialCells(xlCellTypeBlanks).FormulaR1C1 = "=IFERROR(ROUND(RC[-12]/R" & lngLastRowIndex & "C[-12],2),0)"
        Application.Intersect(.UsedRange, Range("N:N, Q:Q, T:T, W:W, Z:Z")).SpecialCells(xlCellTypeBlanks).FormulaR1C1 = "=IF(AND(RC[-2]=0,RC[-1]=0),0,IF(RC[-2]=0,1,(RC[-1]-RC[-2])/RC[-2]))"
        
        .Range(Range("N1"), .Cells.SpecialCells(xlCellTypeLastCell)).SpecialCells(xlCellTypeFormulas).NumberFormat = "0%"
        
        With Application.Intersect(.UsedRange, Range("1:2"))
            .Font.Bold = True
            With .Interior
                .ThemeColor = xlThemeColorAccent1
                .TintAndShade = 0.8
            End With
        End With
    
        With Application.Intersect(.UsedRange, .Cells.SpecialCells(xlCellTypeLastCell).EntireRow)
            .Font.Bold = True
            With .Interior
                .ThemeColor = xlThemeColorAccent1
                .TintAndShade = 0.8
            End With
        End With
        With .UsedRange
            .Borders(xlEdgeTop).LineStyle = xlContinuous
            .Borders(xlEdgeBottom).LineStyle = xlContinuous
            .Borders(xlEdgeRight).LineStyle = xlContinuous
            .Borders(xlInsideVertical).LineStyle = xlContinuous
            .Borders(xlInsideHorizontal).LineStyle = xlContinuous
        End With
        
        'Merge cells containing same value
        Call MergeSameCells(Application.Intersect(.UsedRange, .Rows(1)))
        
        .name = Pvt.PivotFields("Channel").CurrentPage
        
        Call ApplyConditionFormat(.Range("N3:N" & CStr(lngLastRowIndex)))
        Call ApplyConditionFormat(.Range("Q3:Q" & CStr(lngLastRowIndex)))
        Call ApplyConditionFormat(.Range("T3:T" & CStr(lngLastRowIndex)))
        Call ApplyConditionFormat(.Range("W3:W" & CStr(lngLastRowIndex)))
        Call ApplyConditionFormat(.Range("Z3:Z" & CStr(lngLastRowIndex)))
        Call FormatThousandsSeparator(.Range("D:M"))
        Application.Intersect(.UsedRange, .Rows(2)).AutoFilter
        Set rngSort = .Range(.Range("A2"), .Cells.SpecialCells(xlCellTypeLastCell).Offset(-1))
        Set rngSortKey = Application.Intersect(rngSort, .Range("P:P")).Offset(1)
        Set rngSortKey = rngSortKey.Resize(rngSortKey.Rows.count - 1)
        With .Sort
            .SetRange rngSort
            .SortFields.Add2 Key:=rngSortKey, SortOn:=xlSortOnValues, Order:=xlDescending, DataOption:=xlSortNormal
            .Header = xlYes
            .Orientation = xlTopToBottom
            .Apply
        End With

        .UsedRange.EntireColumn.AutoFit
        
        Set rngDataBody = Application.Intersect(.Range("D:K"), .UsedRange)
        Set rngDataBody = rngDataBody.Resize(rngDataBody.Rows.count - 1).SpecialCells(xlCellTypeConstants, 1)
        
        Call createHyperlinks(shtReport, rngDataBody)
        Call FreezePanes
    End With
    Set ConvertPivotToTable = shtReport
Set shtReport = Nothing
End Function

Private Sub FreezePanes()
    With ActiveWindow
        .SplitColumn = 3
        .SplitRow = 2
        .FreezePanes = True
    End With
End Sub

Private Sub FormatThousandsSeparator(RngToFormat As Range)
    With RngToFormat
        Union(.SpecialCells(xlCellTypeConstants, 1), .SpecialCells(xlCellTypeFormulas, 1)).NumberFormat = "#,##0"
    End With
End Sub

Private Sub ApplyConditionFormat(RngToFormat As Range)
    With RngToFormat
        .FormatConditions.AddIconSetCondition
        With .FormatConditions(1)
            .IconSet = ActiveWorkbook.IconSets(xl3Arrows)
        
            With .IconCriteria(2)
                .Type = xlConditionValueNumber
                .Value = 0
                .Operator = 7
            End With
            With .IconCriteria(3)
                .Type = xlConditionValueNumber
                .Value = 0.1
                .Operator = 7
            End With
        End With
    End With
End Sub

Private Sub prepareData(shtRawData As Worksheet)
    Dim arrColumnsWithExceptions() As String
    
    arrColumnsWithExceptions = SetDataSourceType(shtRawData)
    If Len(Join(arrColumnsWithExceptions)) = 0 Then
        Call RemoveDuplicates(shtRawData)
    Else
        MsgBox "The following columns are required for the report:" & vbNewLine & Join(arrColumnsWithExceptions, vbNewLine)
        End
    End If

End Sub

Private Sub filterOutIrrelevantRecords(pvtField As PivotField, arrFilterOutValues As Variant)
    Dim intFilterOutValue As Integer
    
    With pvtField
        For intFilterOutValue = 0 To UBound(arrFilterOutValues)
            On Error Resume Next
            .PivotItems(CStr(arrFilterOutValues(intFilterOutValue))).Visible = False
            Err.Clear
            On Error GoTo 0
        Next intFilterOutValue
    End With
End Sub

Private Sub RemoveDuplicates(shtRawData As Worksheet)
    Dim intArray As Variant, i As Integer
    Dim rng As Range
    
    Set rng = shtRawData.UsedRange.Rows
    With rng
        ReDim intArray(0 To .Columns.count - 1)
        For i = 0 To UBound(intArray)
            intArray(i) = i + 1
        Next i
        .RemoveDuplicates Columns:=(intArray), Header:=xlYes
    End With
End Sub

Private Function CreatePivotReport(shtRawData As Worksheet, ReportName As String)
    Dim Pvt As PivotTable
    Dim shtCustomReport As Worksheet
    Dim strColumnSetFormula As String
    Dim pvtField As PivotField
    Dim shtNational As Worksheet
    Dim chartObjAlertEvolution As ChartObject
    Dim rngWeeklyRRSessionCounts As Range
    Dim rngWeeklyRR_TP_RATE_SESSION As Range
    Dim lngRowOffset As Long
    Dim lngColOffset As Long
    Dim Sht As Worksheet
    Dim rngWithCalculatedItem As Range
    Dim arrVarValuesRiskReasons() As Variant
    
    Set Pvt = GetPivotTable(shtRawData, ReportName) 'Create pivot table
    Set shtCustomReport = ActiveWorkbook.ActiveSheet
    shtCustomReport.name = ReportName

    With Pvt
        .ClearTable
        .ColumnGrand = True
        .RowGrand = False
        
        With .PivotFields(TB_COLUMN_EVENT_DATE)
            .Orientation = xlRowField
            .Position = 1
        End With
        Pvt.RowFields(1).dataRange.Range("A1").Group Start:=True, End:=True, Periods:=Array(False, False, False, False, True, False, False)
        With .PivotFields("Months (Date & time)")
            .Orientation = xlColumnField
            .Position = 1
        End With
        
        With .PivotFields("Classification")
            .Orientation = xlColumnField
            .Position = 1
        End With
        
        Set pvtField = .PivotFields("Reason ID")
        With pvtField
            .Orientation = xlRowField
            .Position = 1
        End With
        arrVarValuesRiskReasons = Array(REASON_ID__1, REASON_ID__2, REASON_ID_BLANK)
        Call filterOutIrrelevantRecords(pvtField, arrVarValuesRiskReasons)

        With .PivotFields("Reason")
            .Orientation = xlRowField
            .Position = 2
        End With
        With .AddDataField(.PivotFields(TB_COLUMN_PUID), "PUIDs", xlCount)
            .NumberFormat = "#,##0"
        End With
        
        With .PivotFields("Risk score")
            .Orientation = xlRowField
            .Position = 3
        End With
        
        Call RemovePivotTableSubtotals(Pvt)
    
        .PivotFields("Channel").Orientation = xlPageField
    
        .DisplayNullString = True
        .NullString = ""
    End With
    Set CreatePivotReport = Pvt
    
    Set Pvt = Nothing
    Set shtCustomReport = Nothing
End Function

Private Function GetPivotTable(Sht As Worksheet, Optional ReportName As String = "Pivot") As PivotTable
    Dim rngRawData As Range
    Dim shtPivot As Worksheet
    Dim Pvt As PivotTable
    Dim pvtCache As PivotCache
    Dim pvtField As PivotField
    
    Set rngRawData = Sht.Range("A1").CurrentRegion
    With ActiveWorkbook
        Set shtPivot = Worksheets.Add
        
        Set pvtCache = .PivotCaches.Create( _
            SourceType:=xlDatabase, _
            SourceData:=rngRawData, _
            Version:=6)
        Set Pvt = pvtCache.CreatePivotTable(TableDestination:=shtPivot.Range("A3"), tableName:=ReportName, DefaultVersion:=6)
    End With
    
    'Assign a VB codename to the Pivot Table Worksheet for future References
    Call RenameCodeName(shtPivot, "shtPivot")
    
    With Pvt
        .ColumnGrand = False
        .RowGrand = False
        .InGridDropZones = True
        .RowAxisLayout xlTabularRow
        .RepeatAllLabels xlRepeatLabels
        .NullString = "0"
        .RowAxisLayout xlTabularRow
        For Each pvtField In .PivotFields
            pvtField.Subtotals = Array(False, False, False, False, False, False, False, False, False, False, False, False)
        Next pvtField
    End With
    
    Set GetPivotTable = Pvt

    Set rngRawData = Nothing
    Set shtPivot = Nothing
    Set Pvt = Nothing
    Set pvtCache = Nothing
End Function

'change sub name to getArrColumnsWithExceptions
Private Function SetDataSourceType(Sht As Worksheet) As String()
    Dim arrColumnsWithExceptionsTB() As String
    Dim arrColumnsTB() As Variant
    
    arrColumnsTB = Array(TB_COLUMN_EVENT_DATE, TB_COLUMN_ACTIVITY_CONFIRMATION, TB_COLUMN_RISK_REASON_ID, TB_COLUMN_RISK_REASON, TB_COLUMN_RISK_SCORE, TB_COLUMN_APPLICATION, TB_COLUMN_PUID, TB_COLUMN_SESSION, TB_COLUMN_ACTIVITY)
    arrColumnsWithExceptionsTB = GetArrayOfMissingColumns(Sht, arrColumnsTB)
    If Len(Join(arrColumnsWithExceptionsTB)) = 0 Then
    
        StrColumnEventDate = TB_COLUMN_EVENT_DATE
        StrColumnActivityConfirmation = TB_COLUMN_ACTIVITY_CONFIRMATION
        StrColumnRiskReasonId = TB_COLUMN_RISK_REASON_ID
        StrColumnRiskReason = TB_COLUMN_RISK_REASON
        StrColumnRiskScore = TB_COLUMN_RISK_SCORE
        StrColumnApplication = TB_COLUMN_APPLICATION
        StrColumnPuid = TB_COLUMN_PUID
        StrColumnSession = TB_COLUMN_SESSION
        StrColumnActivity = TB_COLUMN_ACTIVITY
    
        strClassificationConfirmedFraud = TB_CLASSIFICATION_CONFIRMED_FRAUD
        strClassificationConfirmedLegitimate = TB_CLASSIFICATION_CONFIRMED_LEGITIMATE
        strClassificationUndetermined = TB_CLASSIFICATION_UNDETERMINED
        strClassificationPending = TB_CLASSIFICATION_PENDING
    Else
        SetDataSourceType = arrColumnsWithExceptionsTB
    End If
End Function

Private Function GetArrayOfMissingColumns(Sht As Worksheet, arrColumns() As Variant) As String()
    Dim intColumnName As Integer
    Dim arrColumnsWithExceptions() As String
    Dim intExceptionCounter As Integer
    
    For intColumnName = 0 To UBound(arrColumns)
        If GetSheetColumnIndexByTitle(CStr(arrColumns(intColumnName)), Sht, Sht.Range("A1")) = 0 Then
            ReDim Preserve arrColumnsWithExceptions(intExceptionCounter)
            arrColumnsWithExceptions(intExceptionCounter) = CStr(arrColumns(intColumnName))
            intExceptionCounter = intExceptionCounter + 1
        End If
    Next intColumnName
    GetArrayOfMissingColumns = arrColumnsWithExceptions
    Erase arrColumns
End Function

Private Sub MergeSameCells(WorkRange As Range)
    Dim cell As Range
    'turn off display alerts while merging
    Application.DisplayAlerts = False
    
    'merge all same cells in range
MergeSame:
    If WorkRange.Rows.count = 1 Then
        For Each cell In WorkRange
            If cell.Value = cell.Offset(0, 1).Value And Not IsEmpty(cell) Then
                Range(cell, cell.Offset(0, 1)).Merge
                cell.HorizontalAlignment = xlCenter
                GoTo MergeSame
            End If
        Next
    ElseIf WorkRange.Columns.count = 1 Then
        For Each cell In WorkRange
            If cell.Value = cell.Offset(1, 0).Value And Not IsEmpty(cell) Then
                Range(cell, cell.Offset(1, 0)).Merge
                cell.VerticalAlignment = xlVAlignCenter
                GoTo MergeSame
            End If
        Next
    End If
    
    'turn display alerts back on
    Application.DisplayAlerts = True
End Sub

Private Function CountFilesInFolder(folderPath As String, Optional FileExtension As String = "csv") As Integer
    Dim fileName As String
    Dim intFileCount As Integer
    
    ' Check if the folder path ends with a backslash, if not, add it
    If Right(folderPath, 1) <> "\" Then
        folderPath = folderPath & "\"
    End If
    
    ' Set the initial file count to 0
    intFileCount = 0
    
    ' Get the first file in the folder
    fileName = Dir(folderPath & "*." & FileExtension)
    
    ' Loop through all files in the folder
    Do While fileName <> ""
        ' Increment the file count
        intFileCount = intFileCount + 1
        
        ' Get the next file in the folder
        fileName = Dir()
    Loop
    
    CountFilesInFolder = intFileCount
End Function

Private Sub RenameCodeName(Sht As Worksheet, NewName As String)
    Dim VBProj As VBIDE.VBProject
    Dim vbComps As VBIDE.VBComponents
    Dim VBComp As VBIDE.VBComponent
    Dim vbProps As VBIDE.Properties
    Dim CodeNameProp As VBIDE.Property
    
    Set VBProj = Sht.Parent.VBProject
    Set vbComps = VBProj.VBComponents
    Set VBComp = vbComps(Sht.CodeName)
    Set vbProps = VBComp.Properties
    Set CodeNameProp = vbProps("_Codename")
    CodeNameProp.Value = NewName
    
    Set CodeNameProp = Nothing
    Set vbProps = Nothing
    Set VBComp = Nothing
    Set vbComps = Nothing
    Set VBProj = Nothing
End Sub

Private Sub DeleteIrrelevantRecords(Sht As Worksheet, FieldName As String, Criteria As Variant)
    Dim lngFieldColumnIndex As Long
    Dim lngEria As Long
    Dim erias As Range
    
    lngFieldColumnIndex = GetSheetColumnIndexByTitle(FieldName, Sht, Sht.Range("A1"))
    
    If lngFieldColumnIndex > 0 Then
        With Sht.Range("A1")
            .CurrentRegion.AutoFilter Field:=lngFieldColumnIndex, Criteria1:=Criteria, Operator:=xlFilterValues
            If AutoFilterRecordsFound(Sht) Then
                Set erias = Range(.Offset(1), Sht.Cells(.SpecialCells(xlCellTypeLastCell).Row, 1)).SpecialCells(xlCellTypeVisible)
                For lngEria = erias.Areas.count To 1 Step -1
                    erias.Areas(lngEria).EntireRow.Delete
                    'Range(.Offset(1), Sht.Cells(.SpecialCells(xlCellTypeLastCell).Row, 1)).SpecialCells(xlCellTypeVisible).EntireRow.Delete
                Next lngEria
                'Range(.Offset(1), Sht.Cells(.SpecialCells(xlCellTypeLastCell).Row, 1)).SpecialCells(xlCellTypeVisible).EntireRow.Delete
            End If
            
            .Parent.ShowAllData
        End With
    End If
End Sub

Private Function AutoFilterRecordsFound(Sht As Worksheet) As Boolean
    Dim lngAreasFound As Long
    
    With Sht.AutoFilter.Range
        lngAreasFound = .SpecialCells(xlCellTypeVisible).Areas.count
        AutoFilterRecordsFound = lngAreasFound > 1 Or .SpecialCells(xlCellTypeVisible).Rows.count > 1
    End With
End Function

Private Sub RemovePivotTableSubtotals(pt As PivotTable)
    Dim pvtField As PivotField
    
    On Error Resume Next
    For Each pvtField In pt.PivotFields
        If pvtField.Orientation = xlColumnField Or pvtField.Orientation = xlRowField Then
            With pvtField
                .Subtotals = Array(False, False, False, False, False, False, False, False, False, False, False, False)
            End With
        End If
        Err.Clear
    Next pvtField
    On Error GoTo 0
End Sub

Public Sub AddWorksheetEventCode(Wbk As Workbook, Sht As Worksheet, EventName As String)
'Tools > references > Microsoft Visual Basic for Applications Extensibility 5.3
'Trust access to VBA model

    Dim VBProj As VBIDE.VBProject
    Dim VBComp As VBIDE.VBComponent
    Dim CodeMod As VBIDE.CodeModule
    Dim lngLine As Long
    
    Set VBProj = Wbk.VBProject
    
    Set VBComp = VBProj.VBComponents(Sht.CodeName)
    Set CodeMod = VBComp.CodeModule
    
    With CodeMod
        lngLine = lngLine + 4
        .InsertLines lngLine, "Private Sub DrillThrough(Cel As Range)"
        lngLine = lngLine + 1
        .InsertLines lngLine, "Dim rngClassification As Range"
        lngLine = lngLine + 1
        .InsertLines lngLine, "Dim lngReasonID As Long"
        
        
        lngLine = lngLine + 1
        .InsertLines lngLine, "Dim lngRiskScore As Long"
        
        lngLine = lngLine + 1
        .InsertLines lngLine, "Dim lngMonthIndex As Long"
        
        lngLine = lngLine + 1
        .InsertLines lngLine, "Dim strChannel As String"
        
        lngLine = lngLine + 1
        .InsertLines lngLine, "Dim rngMonth As Range"
        
        lngLine = lngLine + 1
        .InsertLines lngLine, "Dim rngTable As Range"
        
        lngLine = lngLine + 1
        .InsertLines lngLine, "Dim rngHeaderFirstLine As Range"
        
        lngLine = lngLine + 1
        .InsertLines lngLine, "Dim rngHeaderSecondLine As Range"
        
        lngLine = lngLine + 1
        .InsertLines lngLine, "Set rngTable = Cel.CurrentRegion"
        
        lngLine = lngLine + 1
        .InsertLines lngLine, "Set rngHeaderFirstLine = rngTable.Resize(1)"
        
        lngLine = lngLine + 1
        .InsertLines lngLine, "Set rngHeaderSecondLine = rngHeaderFirstLine.Offset(1)"
        
        lngLine = lngLine + 1
        .InsertLines lngLine, "Set rngClassification = Application.Intersect(Cel.EntireColumn, rngHeaderFirstLine)"
        
        lngLine = lngLine + 1
        .InsertLines lngLine, "If rngClassification.Value = """" Then"
        
        lngLine = lngLine + 1
        .InsertLines lngLine, "Set rngClassification = Application.Intersect(Cel.EntireColumn, rngHeaderFirstLine).Offset(0, -1)"
        
        lngLine = lngLine + 1
        .InsertLines lngLine, "End If"
        
        lngLine = lngLine + 1
        .InsertLines lngLine, "lngReasonID = CLng(Application.Intersect(Cel.EntireRow, Range(""A:A"")).Value)"
        
        lngLine = lngLine + 1
        .InsertLines lngLine, "lngRiskScore = CLng(Application.Intersect(Cel.EntireRow, Range(""C:C"")).Value)"
        
        lngLine = lngLine + 1
        .InsertLines lngLine, "Set rngMonth = Application.Intersect(Cel.EntireColumn, rngHeaderSecondLine)"
        
        lngLine = lngLine + 1
        .InsertLines lngLine, "lngMonthIndex = Month(DateValue(""01"" & "" "" & Format(Left(rngMonth.Value2, 3), ""mmm"") & "" "" & Format(Now(), ""yyyy"")))"
        
        lngLine = lngLine + 1
        .InsertLines lngLine, "strChannel = Cel.Parent.Name"
        
        lngLine = lngLine + 1
        .InsertLines lngLine, "GetPivotDataRange(rngClassification, lngReasonID, lngRiskScore, lngMonthIndex, strChannel).ShowDetail = True"
        
        lngLine = lngLine + 1
        .InsertLines lngLine, "End Sub"
        
        lngLine = lngLine + 1
        lngLine = lngLine + 1
        .InsertLines lngLine, "Function GetPivotDataRange(ClassificationValue As Range, ReasonID As Long, RiskScore As Long, MonthIndex As Long, Channel As String) As Range"
        
        lngLine = lngLine + 1
        .InsertLines lngLine, "Dim Pvt As PivotTable"
        
        lngLine = lngLine + 1
        lngLine = lngLine + 1
        .InsertLines lngLine, "Dim rngTableItem As Range"
        
        lngLine = lngLine + 1
        .InsertLines lngLine, "Dim shtPivot As Worksheet"
        
        lngLine = lngLine + 1
        .InsertLines lngLine, "Set shtPivot = ActiveWorkbook.Worksheets(""Pivot Table"")"
        
        lngLine = lngLine + 1
        .InsertLines lngLine, "Set Pvt = shtPivot.PivotTables(1)"
        
        lngLine = lngLine + 1
        .InsertLines lngLine, "Select Case Channel"
        
        lngLine = lngLine + 1
        .InsertLines lngLine, "Case ""online"", ""mobile"":"
        
        lngLine = lngLine + 1
        .InsertLines lngLine, "With Pvt.PivotFields(""Channel"")"
        
        lngLine = lngLine + 1
        .InsertLines lngLine, ".Orientation = xlRowField"
        
        lngLine = lngLine + 1
        .InsertLines lngLine, ".Position = 1"
        
        lngLine = lngLine + 1
        .InsertLines lngLine, "End With"
        
        lngLine = lngLine + 1
        .InsertLines lngLine, "shtPivot.Select"
        
        lngLine = lngLine + 1
        .InsertLines lngLine, "Set rngTableItem = ActiveSheet.PivotTables(1).GetPivotData(""Pinpoint session ID"", ""Classification"", ClassificationValue, ""Reason ID"", ReasonID, ""Risk score"", RiskScore, ""Months (Date & time)"", MonthIndex, ""Channel"", Channel)"
        
        lngLine = lngLine + 1
        .InsertLines lngLine, "Case ""(All)"":"
        
        lngLine = lngLine + 1
        .InsertLines lngLine, "If Pvt.PivotFields(""Channel"").Orientation <> xlHidden Then"
        
        lngLine = lngLine + 1
        .InsertLines lngLine, "Pvt.PivotFields(""Channel"").Orientation = xlHidden"
        
        lngLine = lngLine + 1
        .InsertLines lngLine, "End If"
        
        lngLine = lngLine + 1
        .InsertLines lngLine, "shtPivot.Select"
        
        lngLine = lngLine + 1
        .InsertLines lngLine, "Set rngTableItem = ActiveSheet.PivotTables(1).GetPivotData(""Pinpoint session ID"", ""Classification"", ClassificationValue, ""Reason ID"", ReasonID, ""Risk score"", RiskScore, ""Months (Date & time)"", MonthIndex)"
        
        lngLine = lngLine + 1
        .InsertLines lngLine, "End Select"
                
        lngLine = lngLine + 1
        .InsertLines lngLine, "Set GetPivotDataRange = rngTableItem"
        
        lngLine = lngLine + 1
        .InsertLines lngLine, "End Function"
        
        lngLine = lngLine + 1
        lngLine = .CreateEventProc(EventName, "Worksheet")
        lngLine = lngLine + 1
        .InsertLines lngLine, "Call DrillThrough(Target.Range)"
    End With
    Set VBProj = Nothing
    Set VBComp = Nothing
    Set CodeMod = Nothing
End Sub



