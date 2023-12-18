Attribute VB_Name = "modUtils"
'Option Explicit

Public Enum enumOperatingSystem
    operatingsystemWindows = 0
    operatingsystemMac = 1
    operatingsystemUnknown = 2
End Enum

Public Enum enumDataSource
    datasourceTrustboard = 0
    datasourceAutobot = 1
    datasourceSisenseReports = 2
    datasourceDbReports = 3
End Enum

Public Enum enumReport
    reportActiveReasonsPerActivity = 0
    reportClassificationsByRiskReason = 1
    reportReasonsForForeignCountries = 2
End Enum

Public Const MAX_SHEET_NAME_LENGTH As Integer = 31

Public Const COLUMN_AB_POLICY_CATEGORY As String = "policy_category"
Public Const COLUMN_AB_ACTIVITY As String = "activity"
Public Const COLUMN_AB_RULE_NAME As String = "rule_name"
Public Const COLUMN_AB_SCORE As String = "score"
Public Const COLUMN_AB_REASON_ID As String = "reason_id"
Public Const COLUMN_AB_RECOMMENDATION As String = "recommendation"
Public Const COLUMN_AB_RULE_ID As String = "rule_id"

Public Const REASON_ID__0 As String = "0"
Public Const REASON_ID__2 As String = "-2"
Public Const REASON_ID__1 As String = "-1"
Public Const REASON_ID_19 As String = "19"
Public Const REASON_ID_BLANK As String = "="
Public Const COLUMN_TB_EVENT_DATE As String = "Date & time"
Public Const COLUMN_TB_ACTIVITY_CONFIRMATION As String = "Classification"
Public Const COLUMN_TB_RISK_REASON As String = "Reason"
Public Const COLUMN_TB_RISK_SCORE As String = "Risk score"
Public Const COLUMN_TB_APPLICATION As String = "Application"
Public Const COLUMN_TB_PUID As String = "PUID"
Public Const COLUMN_TB_SESSION As String = "Pinpoint session ID"
Public Const COLUMN_TB_ACTIVITY As String = "Activity"
Public Const COLUMN_TB_RISK_REASON_ID As String = "Reason ID"
Public Const DATETIME_FORMAT As String = "dd/mm/yyyy hh:mm:ss"

Public Const REPORT_ACTIVE_REASONS_PER_ACTIVITY_FROM_AUTOBOT As String = "Active Reasons Per Activity"
Public Const REPORT_CLASSIFICATIONS_BY_RISK_REASON As String = "Classifications By Risk Reason"
Public Const REPORT_REASONS_FOR_FOREIGN_COUNTRIES As String = "Reasons For Foreign Countries"


Public Sub SetPivotTableLayout(Pvt As PivotTable, ReportID As enumReport)
    Dim cubfldDistinctsessions As CubeField
    
    Select Case ReportID
    Case enumReport.reportClassificationsByRiskReason
        With Pvt
            With .CubeFields("[Range].[Reason ID]")
                .Orientation = xlRowField
                .PivotFields(1).Subtotals = Array(False, False, False, False, False, False, False, False, False, False, False, False)
            End With
            With .CubeFields("[Range].[Reason]")
                .Orientation = xlRowField
            End With
            With .CubeFields("[Range].[Classification]")
                .Orientation = xlColumnField
            End With
            Set cubfldDistinctsessions = .CubeFields.GetMeasure("[Range].[Pinpoint session ID]", xlDistinctCount, "Distinct sessions")
            .AddDataField cubfldDistinctsessions
        End With
    Case enumReport.reportReasonsForForeignCountries
        With Pvt
            With .CubeFields("[Range].[Reason ID]")
                .Orientation = xlRowField
                .PivotFields(1).Subtotals = Array(False, False, False, False, False, False, False, False, False, False, False, False)
            End With
            With .CubeFields("[Range].[Reason]")
                .Orientation = xlRowField
                .PivotFields(1).Subtotals = Array(False, False, False, False, False, False, False, False, False, False, False, False)
            End With
            With .CubeFields("[Range].[Risk score]")
                .Orientation = xlRowField
            End With
            Set cubfldDistinctsessions = .CubeFields.GetMeasure("[Range].[PUID]", xlDistinctCount, "Distinct PUIDs")
            .AddDataField cubfldDistinctsessions
        
            With .CubeFields("[Range].[Classification]")
                .Orientation = xlPageField
                .Position = 1
            End With
            With .CubeFields("[Range].[Country]")
                .Orientation = xlPageField
                .Position = 2
                .EnableMultiplePageItems = True
            End With
        End With
    End Select
End Sub

Public Function GetMissingColumns(Sht As Worksheet, DataSource As enumDataSource) As String()
    Dim arrColumnsWithExceptions() As String
    Dim arrColumns() As Variant
    
    Select Case DataSource
    Case enumDataSource.datasourceTrustboard
        arrColumns = Array(COLUMN_TB_EVENT_DATE, COLUMN_TB_ACTIVITY_CONFIRMATION, COLUMN_TB_RISK_REASON, COLUMN_TB_RISK_SCORE, COLUMN_TB_APPLICATION, COLUMN_TB_PUID, COLUMN_TB_SESSION, COLUMN_TB_ACTIVITY)
    Case enumDataSource.datasourceAutobot
        arrColumns = Array(COLUMN_AB_POLICY_CATEGORY, COLUMN_AB_ACTIVITY, COLUMN_AB_RULE_NAME, COLUMN_AB_SCORE, COLUMN_AB_REASON_ID, COLUMN_AB_RECOMMENDATION, COLUMN_AB_RULE_ID)
    End Select
    arrColumnsWithExceptions = GetArrayOfMissingColumns(Sht, arrColumns)
    If Len(Join(arrColumnsWithExceptions)) <> 0 Then
        GetMissingColumns = arrColumnsWithExceptions
    End If
End Function

Public Sub DeleteIrrelevantRecords(Sht As Worksheet, FieldName As String, Criteria As Variant)
    Dim lngFieldColumnIndex As Long
    
    lngFieldColumnIndex = GetSheetColumnIndexByTitle(FieldName, Sht, Sht.Range("A1"))
    
    If lngFieldColumnIndex > 0 Then
        With Sht.Range("A1")
            .CurrentRegion.AutoFilter Field:=lngFieldColumnIndex, Criteria1:=Criteria, Operator:=xlFilterValues
            If AutoFilterRecordsFound(Sht) Then
                Range(.Offset(1), Cells(.SpecialCells(xlCellTypeLastCell).Row, 1)).SpecialCells(xlCellTypeVisible).EntireRow.Delete
            End If
            
            .Parent.AutoFilterMode = False
        End With
    End If
End Sub

Public Sub ConvertDateTimeToDate(Sht As Worksheet)
    Dim rngDataRange As Range
    Dim rngCell As Range
    Dim lngEventDateColIndex As Long
    Dim arrVarDataRange() As Variant
    Dim lngIterator As Long
    Dim strText As String
    
    Set rngDataRange = GetDataRangeForColumn(Sht, Sht.Range("A1").CurrentRegion, COLUMN_TB_EVENT_DATE)
    Call ReplaceValuesInRange(rngDataRange, " UTC", vbNullString, False)
    Call ReplaceValuesInRange(rngDataRange, "T", " ", False)
    Call ReplaceValuesInRange(rngDataRange, "Z", vbNullString, False)
    
    rngDataRange.NumberFormat = DATETIME_FORMAT

    Set rngDataRange = Nothing
    Set rngCell = Nothing
    Erase arrVarDataRange
End Sub

Public Sub RenameRiskReasons(Sht As Worksheet)
    Dim rngRiskReasons As Range
    
    Set rngRiskReasons = GetDataRangeForColumn(Sht, Sht.Range("A1").CurrentRegion, COLUMN_TB_RISK_REASON)
    ReplaceValuesInRange rngRiskReasons, "Unusual activity from a new device using a new browser language", "Unusual activity using a new browser language", True
    ReplaceValuesInRange rngRiskReasons, "Unusual activity using a new browser language", "Unusual activity using a new browser language", True
    ReplaceValuesInRange rngRiskReasons, "Two subsequent logins from different geographical locations within a short timeframe", "Two subsequent logins from different geographical locations within a short timeframe", True
    ReplaceValuesInRange rngRiskReasons, "(GBR", "Two subsequent logins from different geographical locations within a short timeframe", True
    ReplaceValuesInRange rngRiskReasons, "User logged in to a phishing suspected website", "User logged in to a phishing suspected website", True
    ReplaceValuesInRange rngRiskReasons, "Unusual activity from a new device using a hosting service", "Unusual activity using a known risky hosting service", True
    ReplaceValuesInRange rngRiskReasons, "Unusual activity from a new device in a new foreign country", "Unusual activity from a new foreign country", True
    ReplaceValuesInRange rngRiskReasons, "Suspicious anomalous pattern of accesses from a new device", "Suspicious anomalous pattern of accesses", True
    ReplaceValuesInRange rngRiskReasons, "Unusual access from a new device using suspicious attributes", "Unusual access using suspicious device attributes", True
    ReplaceValuesInRange rngRiskReasons, "Unusual transaction made from a new device with a foreign currency", "Unusual transaction made with foreign currency", True
    ReplaceValuesInRange rngRiskReasons, "Unusual transaction made from a new device to a foreign country", "Unusual transaction made to a new foreign country", True
    
    With rngRiskReasons
        .Value2 = Evaluate(Replace("If(@="""","""",Trim(@))", "@", .Address))
    End With

    Set rngRiskReasons = Nothing
End Sub

Public Function GetPivotTable(Sht As Worksheet, ByVal ReportName As String) As PivotTable
    Dim rngRawData As Range
    Dim shtPivot As Worksheet
    Dim Pvt As PivotTable
    Dim pvtCache As PivotCache
    Dim wbkConn As WorkbookConnection

    Set rngRawData = Sht.Range("A1").CurrentRegion
    With ActiveWorkbook
        Set wbkConn = .Connections.Add2( _
            Name:=ReportName, _
            Description:=vbNullString, _
            ConnectionString:="WORKSHEET;" & .Path & "\[" & .Name & "]" & Sht.Name, _
            CommandText:=Sht.Name & "!" & rngRawData.Address, _
            lCmdtype:=xlCmdExcel, _
            CreateModelConnection:=True, _
            ImportRelationships:=False)
        Set shtPivot = Worksheets.Add
        
        Set pvtCache = .PivotCaches.Create( _
            SourceType:=xlExternal, _
            SourceData:=wbkConn, _
            Version:=6)
        Set Pvt = pvtCache.CreatePivotTable(TableDestination:=shtPivot.Range("A3"), _
                TableName:=ReportName, DefaultVersion:=6)
    End With
    
    With Pvt
        .ColumnGrand = True
        .RowGrand = True
        .InGridDropZones = False
        .RowAxisLayout xlTabularRow
        .RepeatAllLabels xlRepeatLabels
        .NullString = "0"
    End With
    
    Set GetPivotTable = Pvt

    Set rngRawData = Nothing
    Set shtPivot = Nothing
    Set Pvt = Nothing
    Set pvtCache = Nothing
    Set wbkConn = Nothing
End Function

Public Function GetArrayOfMissingColumns(Sht As Worksheet, arrColumns() As Variant) As String()
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

Public Function GetSheetColumnIndexByTitle(ColumnTitle As String, Optional Sht As Worksheet, Optional ListHeaderStartCell As Range) As Long
    Dim rngHeader As Range
    Dim lngColIndex As Long

    If Sht Is Nothing And ListHeaderStartCell Is Nothing Then
        Set Sht = ActiveSheet
        Set ListHeaderStartCell = Sht.Range("A1")
        Set rngHeader = ListHeaderStartCell.EntireRow
    ElseIf ListHeaderStartCell Is Nothing Then
        Set ListHeaderStartCell = Sht.Range("A1")
        Set rngHeader = ListHeaderStartCell.EntireRow
    ElseIf Sht Is Nothing Then
        Set Sht = ActiveSheet
        Set rngHeader = Application.Intersect(ListHeaderStartCell.EntireRow, ListHeaderStartCell.CurrentRegion.EntireColumn)
    Else
        Set rngHeader = Application.Intersect(ListHeaderStartCell.EntireRow, Sht.UsedRange.EntireColumn)

    End If

    Err.Clear
    On Error Resume Next
    lngColIndex = Application.WorksheetFunction.Match(ColumnTitle, rngHeader, False)
    If Err.Number <> 0 Then
        Err.Clear
        On Error GoTo 0
        Exit Function
    End If
    On Error GoTo 0

    GetSheetColumnIndexByTitle = ListHeaderStartCell.Column + lngColIndex - 1

    Set rngHeader = Nothing
End Function

Public Function AutoFilterRecordsFound(Sht As Worksheet) As Boolean
    Dim lngAreasFound As Long
    
    With Sht.AutoFilter.Range
        lngAreasFound = .SpecialCells(xlCellTypeVisible).Areas.count
        AutoFilterRecordsFound = lngAreasFound > 1 Or .SpecialCells(xlCellTypeVisible).Rows.count > 1
    End With
End Function

Public Function GetDataRangeForColumn(Sht As Worksheet, DataRange As Range, ColumnName As String) As Range
    Dim lngColIndex As Long
    Dim rngDataNoHeaders As Range
    
    lngColIndex = GetSheetColumnIndexByTitle(ColumnName, Sht, DataRange.Range("A1"))
    Set rngDataNoHeaders = Sht.Range(Sht.Cells(2, lngColIndex), Sht.Cells(DataRange.Rows.count, lngColIndex))
    Set GetDataRangeForColumn = rngDataNoHeaders

    Set rngDataNoHeaders = Nothing
End Function

Public Sub ReplaceValuesInRange(ReplaceRange As Range, strWhat As String, strReplacement As String, ReplaceEntireCell As Boolean)
    If ReplaceEntireCell Then
        strWhat = "*" & strWhat & "*"
    End If
    ReplaceRange.Replace What:=strWhat, Replacement:=strReplacement, LookAt:=xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, ReplaceFormat:=False, FormulaVersion:=xlReplaceFormula2
End Sub

Public Function GetReportName(Report As enumReport) As String
    Dim strReportName As String
    
    Select Case Report
    Case enumReport.reportActiveReasonsPerActivity
        strReportName = REPORT_ACTIVE_REASONS_PER_ACTIVITY_FROM_AUTOBOT
    Case enumReport.reportClassificationsByRiskReason
        strReportName = REPORT_CLASSIFICATIONS_BY_RISK_REASON
    Case enumReport.reportReasonsForForeignCountries
        strReportName = REPORT_REASONS_FOR_FOREIGN_COUNTRIES
    
    End Select
    GetReportName = strReportName
End Function

Public Sub DataCleanupTrustboard(Sht As Worksheet)
    Dim arrVarValuesRiskReasons() As Variant
    
    arrVarValuesRiskReasons = Array(REASON_ID__1, REASON_ID__2, REASON_ID_BLANK, REASON_ID__0)
    Call DeleteIrrelevantRecords(Sht, COLUMN_TB_RISK_REASON_ID, arrVarValuesRiskReasons)
    
    Call ConvertDateTimeToDate(Sht)
    Call RenameRiskReasons(Sht)
    
    Erase arrVarValuesRiskReasons
End Sub

Public Sub RunReportTrustboard(Report As enumReport)
    Dim Pvt As PivotTable
    Dim shtRawData As Worksheet
    Dim arrColumnsWithExceptions As Variant
    
    Set shtRawData = ActiveWorkbook.ActiveSheet
    arrColumnsWithExceptions = GetMissingColumns(shtRawData, enumDataSource.datasourceTrustboard)
    If Len(Join(arrColumnsWithExceptions)) = 0 Then
        Call DataCleanupTrustboard(shtRawData)
        
        Select Case Report
        Case enumReport.reportClassificationsByRiskReason
            Set Pvt = GetPivotTable(shtRawData, GetReportName(reportClassificationsByRiskReason))
            Call SetPivotTableLayout(Pvt, enumReport.reportClassificationsByRiskReason)
        Case enumReport.reportReasonsForForeignCountries
            Set Pvt = GetPivotTable(shtRawData, GetReportName(reportReasonsForForeignCountries))
            Call SetPivotTableLayout(Pvt, enumReport.reportReasonsForForeignCountries)
        End Select

    Else
        MsgBox "The following columns are required for the report:" & vbNewLine & Join(arrColumnsWithExceptions, vbNewLine)
    End If

    Set shtRawData = Nothing
End Sub

Public Function getOperatingSystem() As enumOperatingSystem
    Dim os As String
    os = Application.OperatingSystem
    
    If InStr(1, os, "Windows", vbTextCompare) > 0 Then
        getOperatingSystem = enumOperatingSystem.operatingsystemWindows
    ElseIf InStr(1, os, "Mac", vbTextCompare) > 0 Then
        getOperatingSystem = enumOperatingSystem.operatingsystemMac
    Else
        getOperatingSystem = enumOperatingSystem.operatingsystemUnknown
    End If
End Function

Public Function RangeToArray(rngData As Range) As Variant()
    Dim arrVarRange() As Variant
    
    arrVarRange = rngData.Value2
    RangeToArray = arrVarRange

    Erase arrVarRange
End Function

Public Sub ArrayToRange(VariantArray() As Variant, StartCell As Range)
    Dim lngRow As Long
    Dim lngCol As Long
    Dim rngTarget As Range
    
    lngRow = UBound(VariantArray, 1)
    lngCol = UBound(VariantArray, 2)
    Set rngTarget = StartCell.Resize(lngRow, lngCol)
    rngTarget = VariantArray

    Erase VariantArray
    Set rngTarget = Nothing
End Sub

Public Sub FormatDateColumns(Optional Sht As Worksheet)
    Dim rngDataTable As Range
    Dim rngHeader As Range
    Dim intColumnIndex As Integer
    
    If Sht Is Nothing Then
        Set Sht = ActiveSheet
    End If
    Set rngDataTable = Sht.UsedRange
    Set rngHeader = Application.Intersect(rngDataTable, Sht.Rows(1))
    For intColumnIndex = 1 To rngHeader.Columns.count Step 1
        If InStr(LCase(rngHeader(intColumnIndex).Value), "date") <> 0 Then
            Application.Intersect(rngDataTable, Sht.Columns(intColumnIndex)).NumberFormat = "dd/mm/yyyy"
        ElseIf InStr(LCase(rngHeader(intColumnIndex).Value), "day") <> 0 Then
            Application.Intersect(rngDataTable, Sht.Columns(intColumnIndex)).NumberFormat = "dd/mm/yyyy"
        ElseIf InStr(LCase(rngHeader(intColumnIndex).Value), "created_at") <> 0 Then
            Application.Intersect(rngDataTable, Sht.Columns(intColumnIndex)).NumberFormat = "dd/mm/yyyy hh:mm:ss"
        ElseIf InStr(LCase(rngHeader(intColumnIndex).Value), "processed_at") <> 0 Then
            Application.Intersect(rngDataTable, Sht.Columns(intColumnIndex)).NumberFormat = "dd/mm/yyyy hh:mm:ss"
        End If
    Next intColumnIndex
End Sub

Public Function isLANCableconnected() As Boolean
    Dim adapter
    
    For Each adapter In GetObject("winmgmts:root\CIMV2").ExecQuery("Select * from Win32_NetworkAdapter")
        If InStr(adapter.NetConnectionID, "Ethernet") <> 0 Then
            If adapter.NetEnabled Then
                isLANCableconnected = True
                Exit Function
            End If
        End If
    Next
End Function

Public Function CreateQueryTable(Sht As Worksheet) As QueryTable
    Dim Qry As QueryTable
    
    'Create the QueryTable object
    Select Case Application.Version
    Case "12.0", "14.0", "15.0", "16.0":
        Set Qry = Sht.ListObjects.Add( _
            SourceType:=xlSrcExternal, _
            source:="ODBC;DSN=Redshift_EU;", _
            Destination:=Sht.Range("A1") _
        ).QueryTable
    Case Else
        MsgBox "Please configure QueryTable for Excel version " & Application.Version
    End Select
    Set CreateQueryTable = Qry
    Set Qry = Nothing
End Function

Function getCurrentMacroName() As String
    Dim callerType As String
    
    Select Case TypeName(Application.Caller)
    Case "Range"
        callerType = Application.Caller.Address
    Case "String"
        callerType = Application.Caller
    Case "Error"
        callerType = "Error"
    Case Else
        callerType = "unknown"
    End Select
    getCurrentMacroName = callerType
End Function

'Public Sub LogMacroRun(macroName As String, startTime As Date, endTime As Date)
'    Dim lastRow As Long
'    lastRow = logSheet.Cells(Log.Rows.count, "A").End(xlUp).Row + 1
'
'    With logSheet
'        .Cells(lastRow, 1).Value = macroName
'        .Cells(lastRow, 2).Value = startTime
'        .Cells(lastRow, 3).Value = endTime
'        .Cells(lastRow, 4).Value = endTime - startTime
'        .Cells(lastRow, 5).Value = Application.UserName
'    End With
'End Sub
'

Function IsSQLStatementValid(sqlStatement As String, ByRef ErrorDescription As String) As Boolean
    Dim conn As Object
    Set conn = CreateObject("ADODB.Connection")
    
    conn.ConnectionString = "Redshift_EU"
    
    On Error Resume Next
    conn.Open
    On Error GoTo 0
    
    If conn.State = 1 Then
        Dim cmd As Object
        Set cmd = CreateObject("ADODB.Command")
        Set cmd.ActiveConnection = conn
        
        On Error Resume Next
        cmd.CommandText = sqlStatement
        cmd.Execute , , adExecuteNoRecords ' adExecuteNoRecords prevents returning any records, which is aimed for performance
        If Err.Number <> 0 Then
            IsSQLStatementValid = False
            ErrorDescription = Err.Description
            Err.Clear
        Else
            IsSQLStatementValid = True
        End If
        On Error GoTo 0
        
        conn.Close
    Else
        IsSQLStatementValid = False
        ErrorDescription = "Failed to connect to the database."
    End If
End Function

Public Function SplitWorksheetsByColumnValues(FieldName As String, Sht As Worksheet) As Boolean
    Dim rRange As Range, rCell As Range
    Dim strText As String
    Dim varColumnIndexForSplit As Variant
    Dim intColumnIndexForSplit As Integer
    Dim booSuccess As Boolean
    Dim lngMatchColIndex As Long

    booSuccess = False
    Sht.AutoFilterMode = False

    'Store the Column Index For Split
    varColumnIndexForSplit = Application.Match(FieldName, Application.Intersect(Sht.Rows(1), Sht.Range("A1").CurrentRegion), False)
    If IsError(varColumnIndexForSplit) Then
        SplitWorksheetsByColumnValues = booSuccess
        Exit Function
    Else
        intColumnIndexForSplit = CInt(varColumnIndexForSplit)
    End If

    'Set a range variable to the correct item column
    Set rRange = Application.Intersect(Sht.Columns(varColumnIndexForSplit), Sht.Range("A1").CurrentRegion)

    'Add a sheet called "UniqueList"
    Worksheets.Add().Name = "UniqueList"

    'Filter the Set range so only a unique list is created
     With Worksheets("UniqueList")
         rRange.AdvancedFilter xlFilterCopy, , Worksheets("UniqueList").Range("A1"), True

          'Set a range variable to the unique list, less the heading.
          Set rRange = .UsedRange.Offset(1).Resize(.UsedRange.Rows.count - 1)
     End With

     lngMatchColIndex = GetSheetColumnIndexByTitle(FieldName, Sht, Sht.Range("A1"))

     'On Error Resume Next
     With Sht
         For Each rCell In rRange
           strText = rCell.Value
          .Range("A1").AutoFilter lngMatchColIndex, strText
             'Add a sheet named as content of rCell
             Worksheets.Add().Name = Left(strText, MAX_SHEET_NAME_LENGTH)
             'Copy the visible filtered range (default of Copy Method) and leave hidden rows
             .Range("A1").CurrentRegion.SpecialCells(xlCellTypeVisible).Copy Destination:=ActiveSheet.Range("A1")
             ActiveSheet.Cells.Columns.AutoFit
         Next rCell
     End With

     Application.DisplayAlerts = False
     Worksheets("UniqueList").Delete
     Sht.Delete
     Application.DisplayAlerts = True
     booSuccess = True
     SplitWorksheetsByColumnValues = booSuccess
End Function

Function GetFieldUniqueValues(FieldName As String, Sht As Worksheet) As Variant()
    Dim rngData As Range
    Dim d As Object
    Set d = CreateObject("Scripting.Dictionary")
    Dim myArray() As Variant
    Dim rngRecommendation As Range
    
    Set rngData = Sht.Range("A1").CurrentRegion
    Set rngRecommendation = GetDataRangeForColumn(Sht, rngData, FieldName)
    myArray = RangeToArray(rngRecommendation)
    
    Dim i As Long
    For i = LBound(myArray) To UBound(myArray)
        d(myArray(i, 1)) = 1
    Next i
    
    GetFieldUniqueValues = d.keys
    
    Set d = Nothing
End Function

Function GetDownloadsPath() As String
    GetDownloadsPath = Environ$("USERPROFILE") & "\Downloads"
    'GetDownloadsPath = Environ$("USERPROFILE") & "\Documents"
End Function

Function getSelectedFolder(Optional OpenAt As Variant) As Variant
    Dim ShellApp As Object
     
    Set ShellApp = CreateObject("Shell.Application").BrowseForFolder(0, "Select the folder containing Trustboard .csv export files", 0, OpenAt)
     
    On Error Resume Next
    getSelectedFolder = ShellApp.self.Path
    On Error GoTo 0
     
    Set ShellApp = Nothing
     
     'Check for invalid or non-entries and send to the Invalid error
     'handler if found
     'Valid selections can begin L: (where L is a letter) or
     '\\ (as in \\servername\sharename.  All others are invalid
    Select Case Mid(getSelectedFolder, 2, 1)
    Case Is = ":"
        If Left(getSelectedFolder, 1) = ":" Then GoTo Invalid
    Case Is = "\"
        If Not Left(getSelectedFolder, 1) = "\" Then GoTo Invalid
    Case Else
        GoTo Invalid
    End Select
    
    Set ShellApp = Nothing
    Exit Function
     
Invalid:
     'If it was determined that the selection was invalid, set to False
    getSelectedFolder = False
End Function

Function CountFilesInFolder(FolderPath As String, Optional FileExtension As String = "csv") As Integer
    Dim filename As String
    Dim intFileCount As Integer
    
    ' Check if the folder path ends with a backslash, if not, add it
    If Right(FolderPath, 1) <> "\" Then
        FolderPath = FolderPath & "\"
    End If
    
    ' Set the initial file count to 0
    intFileCount = 0
    
    ' Get the first file in the folder
    filename = Dir(FolderPath & "*." & FileExtension)
    
    ' Loop through all files in the folder
    Do While filename <> ""
        ' Increment the file count
        intFileCount = intFileCount + 1
        
        ' Get the next file in the folder
        filename = Dir()
    Loop
    
    CountFilesInFolder = intFileCount
End Function

Sub ReplaceCsvExtensionWithNone()
    Range(Range("A2"), Range("A" & Range("A1").CurrentRegion.Rows.count)).Replace What:=".csv", Replacement:="", LookAt:=xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, ReplaceFormat:=False, FormulaVersion:=xlReplaceFormula2
End Sub

