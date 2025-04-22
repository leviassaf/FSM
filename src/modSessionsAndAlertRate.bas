Attribute VB_Name = "modSessionsAndAlertRate"
Option Explicit

' Main procedure to import CSV data with user file selection
Sub ImportSessionsWithAlertRate()
    Dim filePath As String
    Dim queryName As String
    Dim tableName As String
    
    ' Get file path from user
    filePath = GetUserSelectedFile("CSV Files (*.csv), *.csv")
    If filePath = "" Then
        MsgBox "Import canceled by user.", vbInformation
        Exit Sub
    End If
    
    ' Extract file name without extension for query and table naming
    queryName = GetFileNameWithoutExtension(filePath)
    tableName = CleanNameForTable(queryName)
    
    ' Create the query
    CreateQuery filePath, queryName
    
    ' Add a new worksheet and create the table
    AddWorksheetAndCreateTable queryName, tableName
End Sub

Sub CreateDualAxisChart()
    ' Create a chart that displays Total Sessions and Alert Rate with dual axes
    
    ' Clear any existing charts
    ClearExistingCharts
    
    ' Get data range dynamically
    Dim dataRange As Range
    Set dataRange = GetDataRange()
    
    ' Create and format the chart
    Dim chartObj As ChartObject
    Set chartObj = CreateChartObject()
    
    ' Add the data series to the chart
    AddChartSeries chartObj, dataRange
    
    ' Format the chart elements
    FormatChartElements chartObj, dataRange
    
    ' Add data labels to the series
    AddDataLabels chartObj
End Sub

' Function to get a file selected by the user
Function GetUserSelectedFile(fileFilter As String) As String
    Dim fd As Office.FileDialog
    
    Set fd = Application.FileDialog(msoFileDialogFilePicker)
    With fd
        .Title = fileFilter
        .AllowMultiSelect = False
        .Filters.Clear
        .Filters.Add "CSV Files", "*.csv"
        
        If .Show = True Then
            GetUserSelectedFile = .SelectedItems(1)
        Else
            GetUserSelectedFile = ""
        End If
    End With
End Function

' Function to extract file name without extension
Function GetFileNameWithoutExtension(filePath As String) As String
    Dim fileName As String
    fileName = Mid(filePath, InStrRev(filePath, "\") + 1)
    GetFileNameWithoutExtension = Left(fileName, InStrRev(fileName, ".") - 1)
End Function

' Function to clean a name for use as a table name
Function CleanNameForTable(name As String) As String
    Dim cleanName As String
    cleanName = Replace(name, " ", "_")
    ' Remove any other invalid characters if needed
    CleanNameForTable = cleanName
End Function

' Procedure to create a query from the CSV file
Sub CreateQuery(filePath As String, queryName As String)
    Dim formulaText As String
    
    formulaText = "let" & vbCrLf & _
        "    Source = Csv.Document(File.Contents(""" & filePath & """),[Delimiter="","", Columns=5, QuoteStyle=QuoteStyle.None])," & vbCrLf & _
        "    #""Promoted Headers"" = Table.PromoteHeaders(Source, [PromoteAllScalars=true])," & vbCrLf & _
        "    #""Changed Type"" = Table.TransformColumnTypes(#""Promoted Headers""," & _
        "{{""Group"", type text}, {""Date"", type date}, {""Alerted session"", Int64.Type}, {""Total sessions"", Int64.Type}, {""Alert rate"", Percentage.Type}})," & vbCrLf & _
        "    #""Removed Columns"" = Table.RemoveColumns(#""Changed Type"",{""Group"", ""Alerted session""})" & vbCrLf & _
        "in" & vbCrLf & _
        "    #""Removed Columns"""
    
    ActiveWorkbook.Queries.Add name:=queryName, formula:=formulaText
End Sub

' Procedure to add a worksheet and create a table linked to the query
Sub AddWorksheetAndCreateTable(queryName As String, tableName As String)
    Dim ws As Worksheet
    
    ' Add a new worksheet
    Set ws = ActiveWorkbook.Worksheets.Add
    
    ' Create a table linked to the query
    With ws.ListObjects.Add(SourceType:=0, Source:=Array( _
        "OLEDB;Provider=Microsoft.Mashup.OleDb.1;Data Source=$Workbook$;Location=""" & queryName & """;Extended Pr", _
        "operties="""""), Destination:=ws.Range("$A$1")).queryTable
        
        .CommandType = xlCmdSql
        .CommandText = Array("SELECT * FROM [" & queryName & "]")
        .BackgroundQuery = True
        .SaveData = True
        .AdjustColumnWidth = True
'        .listObject.displayName = tableName
        .Refresh BackgroundQuery:=False
    End With
End Sub

Function GetDataRange() As Range
    ' Dynamically determine the data range
    Dim ws As Worksheet
    Set ws = ActiveSheet
    
    ' Find the last row with data
    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.count, "A").End(xlUp).Row
    
    ' Set the range including headers
    Set GetDataRange = ws.Range("A1:C" & lastRow)
End Function

Sub ClearExistingCharts()
    ' Remove any existing charts
    Dim chartObj As ChartObject
    For Each chartObj In ActiveSheet.ChartObjects
        chartObj.Delete
    Next chartObj
End Sub

Function CreateChartObject() As ChartObject
    ' Create a new chart object
    Dim chartObj As ChartObject
    Set chartObj = ActiveSheet.ChartObjects.Add(Left:=100, Width:=800, Top:=300, Height:=400)
    
    ' Set chart type to Line
    chartObj.Chart.ChartType = xlLine
    
    Set CreateChartObject = chartObj
End Function

Sub AddChartSeries(chartObj As ChartObject, dataRange As Range)
    ' Add the two data series to the chart
    With chartObj.Chart
        ' Clear any existing series
        Do While .SeriesCollection.count > 0
            .SeriesCollection(1).Delete
        Loop
        
        ' Add Total Sessions series
        Dim seriesTotal As series
        Set seriesTotal = .SeriesCollection.NewSeries
        With seriesTotal
            .Values = dataRange.Offset(1, 1).Resize(dataRange.Rows.count - 1, 1)
            .XValues = dataRange.Offset(1, 0).Resize(dataRange.Rows.count - 1, 1)
            .Format.Line.ForeColor.RGB = RGB(255, 180, 0) ' Yellow color
            .Format.Line.Weight = 2
        End With
        
        ' Add Alert Rate series
        Dim seriesAlert As series
        Set seriesAlert = .SeriesCollection.NewSeries
        With seriesAlert
            .Values = dataRange.Offset(1, 2).Resize(dataRange.Rows.count - 1, 1)
            .XValues = dataRange.Offset(1, 0).Resize(dataRange.Rows.count - 1, 1)
            .Format.Line.ForeColor.RGB = RGB(0, 102, 153) ' Blue color
            .Format.Line.Weight = 2
            
            ' Set Alert Rate series to use the secondary axis
            .AxisGroup = xlSecondary
        End With
        seriesAlert.name = "Alert rate"
        seriesTotal.name = "Total Sessions"
    End With
End Sub

Sub FormatChartElements(chartObj As ChartObject, dataRange As Range)
    ' Format all chart elements
    With chartObj.Chart
        ' Set chart title
        .HasTitle = False
        
        ' Format primary axis (Total Sessions)
        With .Axes(xlValue, xlPrimary)
            .HasTitle = False
            .MinimumScale = 0
            .HasMajorGridlines = True
            .MajorGridlines.Format.Line.ForeColor.RGB = RGB(220, 220, 220)
            .MajorGridlines.Format.Line.Weight = 0.75
            
            ' Format numbers to show with K suffix for thousands
            .TickLabels.NumberFormat = "0""M"""
            
            ' Calculate a nice maximum based on data
            Dim maxSessions As Double
            maxSessions = WorksheetFunction.Max(dataRange.Offset(1, 1).Resize(dataRange.Rows.count - 1, 1))
            .MaximumScale = WorksheetFunction.Ceiling(maxSessions, 2000000)
        End With
        
        ' Format secondary axis (Alert Rate)
        With .Axes(xlValue, xlSecondary)
            .HasTitle = False
            .MinimumScale = 0
            .TickLabels.NumberFormat = "0.000%"
            
            ' Calculate a nice maximum based on data
            Dim maxRate As Double
            maxRate = WorksheetFunction.Max(dataRange.Offset(1, 2).Resize(dataRange.Rows.count - 1, 1))
            .MaximumScale = WorksheetFunction.Ceiling(maxRate, 0.001)
        End With
        
        ' Format category axis (Date)
        With .Axes(xlCategory)
            .HasTitle = False
            .TickLabelPosition = xlTickLabelPositionLow
            
            ' Format dates on X-axis
            .TickLabels.NumberFormat = "mm/yyyy"
        End With
        
        ' Add legend
        .HasLegend = True
        With .Legend
            .Position = xlLegendPositionBottom
            .Format.TextFrame2.TextRange.Font.Size = 10
        End With
        
        ' Set plot area
        With .PlotArea
            .Format.Fill.Visible = False
            .Border.LineStyle = xlNone
        End With
        
        ' Remove chart border
        .ChartArea.Format.Line.Visible = msoFalse
    End With
End Sub

Sub AddDataLabels(chartObj As ChartObject)
    ' Add data labels to both series
    Dim series As series
    
    ' Loop through each series to add data labels
    For Each series In chartObj.Chart.SeriesCollection
        With series
            .HasDataLabels = True
            
            With .DataLabels
                .ShowValue = True
                .Position = xlLabelPositionAbove
                .Format.TextFrame2.TextRange.Font.Size = 9
                
                ' Format the data labels based on series type
                If series.name = "Total Sessions" Then
                    .NumberFormat = "###0.0""K"""
                ElseIf series.name = "Alert rate" Then
                    .NumberFormat = "0.000%"
                End If
            End With
        End With
    Next series
End Sub

' Helper function to format dates in custom way if needed
Function FormatDateForAxis(inputDate As Date) As String
    FormatDateForAxis = Format(inputDate, "mm/yyyy")
End Function
