Attribute VB_Name = "modActiveReasonsPerActivity"
Option Explicit

Sub ActiveReasonsPerActivity()
    Dim Pvt As PivotTable
    Dim shtPivot As Worksheet
    
    Call ActiveReasonsPerActivityInternal
    Set shtPivot = ActiveSheet
    Set Pvt = shtPivot.PivotTables(1)
    Call CreateCustomerFacingActiveReasons(Pvt, shtPivot)
End Sub

Sub ActiveReasonsPerActivityInternal()
    Dim shtRawData As Worksheet
    Dim Lst As ListObject
    Dim Pvt As PivotTable
    Dim shtPivot As Worksheet
    Dim rngFilteredRanges As Range
    Dim lngContiguousRangeIndex As Long
    Dim rngPinpointReasonReference As Range
    Dim arrColumnsWithExceptions As Variant
    Dim rngReasonDataNoHeader As Range
    Dim rngReasonIdDataNoHeader As Range
    Dim arrVarColumnReason() As Variant
    Dim pvtItem As PivotItem
    
    ActiveWorkbook.ApplyTheme ( _
        "C:\Program Files\Microsoft Office\root\Document Themes 16\Office 2013 - 2022 Theme.thmx")

    Set shtRawData = ActiveSheet
    ThisWorkbook.Sheets("Pinpoint Reason Reference").Copy Before:=ActiveWorkbook.Sheets(1)
        arrColumnsWithExceptions = GetMissingColumns(shtRawData, enumDataSource.datasourceAutobot)
        If Len(Join(arrColumnsWithExceptions)) = 0 Then
        
            With shtRawData
                If .ListObjects.count = 0 Then
                    Set Lst = .ListObjects.Add(xlSrcRange, .UsedRange, , xlYes)
                Else
                    Set Lst = .ListObjects(1)
                End If
            
                .Columns(6).Insert Shift:=xlToRight
                .Cells(1, 6).Value = "reason"
                Set rngPinpointReasonReference = shtPinpointReasonReference.Range("A1").CurrentRegion
                Set rngReasonDataNoHeader = GetDataRangeForColumn(shtRawData, .Range("A1").CurrentRegion, "reason")
                Set rngReasonIdDataNoHeader = GetDataRangeForColumn(shtRawData, .Range("A1").CurrentRegion, "reason_id")
                Call CleanReasonIDColumn(rngReasonIdDataNoHeader)
                rngReasonDataNoHeader.FormulaR1C1 = "=VLOOKUP([@[reason_id]],'" & shtPinpointReasonReference.Name & "'!" & rngPinpointReasonReference.Address(RowAbsolute:=True, ColumnAbsolute:=True, ReferenceStyle:=xlR1C1) & ",2,0)"
            End With
        
            arrVarColumnReason = RangeToArray(rngReasonDataNoHeader)
            Call ArrayToRange(arrVarColumnReason, rngReasonDataNoHeader.Cells(1))
            
            Set shtPivot = Sheets.Add
            shtPivot.Name = "Pivot Table"
            
            ActiveWorkbook.PivotCaches.Create(SourceType:=xlDatabase, SourceData:=shtRawData.Range("A1").CurrentRegion, Version:=6). _
                CreatePivotTable TableDestination:=shtPivot.Cells(3, 1), DefaultVersion:=6
            Set Pvt = ActiveSheet.PivotTables(1)
            
            With Pvt
                .InGridDropZones = True
                .RowAxisLayout xlTabularRow
                .ColumnGrand = False
                .RowGrand = False
                .HasAutoFormat = False
                
                With .PivotFields("reason_id")
                    .Orientation = xlRowField
                    .Position = 1
                    .Subtotals = Array(False, False, False, False, False, False, False, False, False, False, False, False)
                End With
                
                With .PivotFields("reason")
                    .Orientation = xlRowField
                    .Position = 2
                End With
                
                With .PivotFields("activity")
                    .Orientation = xlColumnField
                    .Position = 1
                End With
                .AddDataField .PivotFields("rule_name"), "Count of rule_name", xlCount
                
                ' Hide from report test policies and specific Reason IDs that are less of interest to customers
                .PivotFields("policy_category").PivotFilters.Add2 Type:=xlCaptionDoesNotContain, Value1:="test"
                With .PivotFields("reason_id")
                    On Error Resume Next 'ignore errors if the RR doesn't exist and cannot be become invisible
                    .PivotItems("-1").Visible = False
                    .PivotItems("0").Visible = False
                    .PivotItems("55").Visible = False
'                    .PivotItems("88").Visible = False
'                    .PivotItems("89").Visible = False
                    Err.Clear
                    On Error GoTo 0
                End With
                        
                With .PivotFields("policy_category")
                    .Orientation = xlPageField
                    .EnableMultiplePageItems = True
                End With
                
                For Each pvtItem In .PivotFields("policy_category").PivotItems
                    If InStr(pvtItem.Name, "test") <> 0 Then
                        pvtItem.Visible = False
                    End If
                Next pvtItem
                
                .TableRange1.EntireColumn.AutoFit
            End With
        Else
            MsgBox "The following columns are required for the report:" & vbNewLine & Join(arrColumnsWithExceptions, vbNewLine)
        End If
End Sub

Private Function GetDataRangeForColumn(Sht As Worksheet, DataRange As Range, ColumnName As String) As Range
    Dim lngColIndex As Long
    Dim rngDataNoHeaders As Range
    
    lngColIndex = GetSheetColumnIndexByTitle(ColumnName, Sht, DataRange.Range("A1"))
    Set rngDataNoHeaders = Sht.Range(Sht.Cells(2, lngColIndex), Sht.Cells(DataRange.Rows.count, lngColIndex))
    Set GetDataRangeForColumn = rngDataNoHeaders

    Set rngDataNoHeaders = Nothing
End Function

Private Function GetSheetColumnIndexByTitle(ColumnTitle As String, Optional Sht As Worksheet, Optional ListHeaderStartCell As Range) As Long
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

Sub CreateCustomerFacingActiveReasons(Pvt As PivotTable, shtSource As Worksheet)
    Dim rngSource As Range
    Dim rngDestination As Range
    Dim shtDestination As Worksheet
    Dim lngSourceRowsCount As Long
    Dim lngSourceColumnsCount As Long
    Dim ListObj As ListObject
    Dim strDestDataBodyRangeAddress As String
    
    Set shtDestination = Worksheets.Add
    Set Pvt = shtSource.PivotTables(1)
    Set rngSource = Pvt.TableRange1
    strDestDataBodyRangeAddress = Pvt.DataBodyRange.Offset(-2 - Pvt.PageFields.count).Address
    lngSourceRowsCount = rngSource.Rows.count - 1
    lngSourceColumnsCount = rngSource.Columns.count
    Set rngSource = rngSource.Offset(1).Resize(lngSourceRowsCount, lngSourceColumnsCount)
    Set rngDestination = shtDestination.Range("A1").Resize(lngSourceRowsCount, lngSourceColumnsCount)
    rngDestination.Value = rngSource.Value

    With shtDestination
        .Name = "Active Reasons by Activity"
        Set ListObj = .ListObjects.Add(xlSrcRange, .UsedRange, , xlYes)
        With ListObj
            .TableStyle = "TableStyleLight13"
            With .Range
                .Borders(xlEdgeTop).LineStyle = xlContinuous
                .Borders(xlEdgeBottom).LineStyle = xlContinuous
                .Borders(xlEdgeRight).LineStyle = xlContinuous
                .Borders(xlInsideVertical).LineStyle = xlContinuous
                .Borders(xlInsideHorizontal).LineStyle = xlContinuous
                .EntireColumn.AutoFit
            End With
        End With
        With .Range(strDestDataBodyRangeAddress)
            .SpecialCells(xlCellTypeConstants, 1).Value2 = "V"
            .EntireColumn.HorizontalAlignment = xlCenter
        End With
    End With
    Call PrepareForPrint(shtDestination)
End Sub

Sub PrepareForPrint(Sht As Worksheet)
    Dim rngPrintArea As Range
    
    Set rngPrintArea = Sht.UsedRange
    
    With Sht.PageSetup
        .PrintArea = rngPrintArea.Address
        .Orientation = xlLandscape
        .Zoom = False
        .FitToPagesWide = 1
        .FitToPagesTall = False
        .CenterHeader = "&""-,Bold""&22&A"
        .CenterHorizontally = True
    End With
End Sub

Private Sub CleanReasonIDColumn(RangeToClean As Range)
    Dim Cel As Range
    Dim lngCleansedValue As Long
    
    For Each Cel In RangeToClean
        Err.Clear
        On Error Resume Next
        lngCleansedValue = Trim(Left(Cel.Value2, InStr(Cel.Value2, ":") - 1))
        If Err.Number = 0 Then
            Cel.Value2 = lngCleansedValue
        Else
            Cel.Value2 = 0
        End If
    Next Cel
    RangeToClean.NumberFormat = "General"
End Sub
