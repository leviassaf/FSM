Attribute VB_Name = "modGetTBSessionsAndAlerts"
Option Explicit

Sub getTrustboardSessionsAndAlerts()
Attribute getTrustboardSessionsAndAlerts.VB_ProcData.VB_Invoke_Func = "T\n14"
    Dim sessionsFilePath As String
    Dim alertsFilePath As String
    
    ' Create a new workbook with exactly one worksheet
    Call Workbooks.Add(xlWBATWorksheet)
    
    ' Get file paths from user
    sessionsFilePath = GetFilePath("Select the Sessions CSV file")
    If sessionsFilePath = "" Then Exit Sub
    
    alertsFilePath = GetFilePath("Select the Alerts CSV file")
    If alertsFilePath = "" Then Exit Sub
    
    Call ClearExistingQueries
    ' Create queries
    CreateSessionsQuery sessionsFilePath
    CreateAlertsQuery alertsFilePath
    CreateJoinedQuery
    
    ' Create tables in worksheets
    Call CreateQueryTable("sessions_csv")
    Call CreateQueryTable("alerts_csv")
    Call CreateQueryTable("sessions_and_alerts")
    Call DeleteAllSheetsExceptActive
    Call CreateChart
    Call CreateSecurityDashboardPPT
End Sub

' Function to get file path using dialog
Function GetFilePath(promptText As String) As String
    Dim fd As Office.FileDialog
    
    Set fd = Application.FileDialog(msoFileDialogFilePicker)
    With fd
        .Title = promptText
        .AllowMultiSelect = False
        .Filters.Clear
        .Filters.Add "CSV Files", "*.csv"
        
        If .Show = True Then
            GetFilePath = .SelectedItems(1)
        Else
            GetFilePath = ""
        End If
    End With
End Function

' Create sessions query
Sub CreateSessionsQuery(filePath As String)
    Dim queryFormula As String
    
    queryFormula = _
        "let" & vbCrLf & _
        "    Source = Csv.Document(File.Contents(""" & filePath & """)," & _
        "[Delimiter="","", Columns=3, QuoteStyle=QuoteStyle.None])," & vbCrLf & _
        "    #""Promoted Headers"" = Table.PromoteHeaders(Source, [PromoteAllScalars=true])," & vbCrLf & _
        "    #""Changed Type"" = Table.TransformColumnTypes(#""Promoted Headers""," & _
        "{{""Group"", type text}, {""Date"", type date}, " & _
        "{""Number of sessions"", Int64.Type}})," & vbCrLf & _
        "    #""Pivoted Column"" = Table.Pivot(#""Changed Type"", " & _
        "List.Distinct(#""Changed Type""[Group]), ""Group"", " & _
        """Number of sessions"", List.Sum)," & vbCrLf & _
        "    #""Renamed Columns"" = Table.RenameColumns(#""Pivoted Column""," & _
        "{{""Mobile"", ""Mobile Sessions""}, {""Online"", ""Online Sessions""}})," & vbCrLf & _
        "    #""Inserted Addition"" = Table.AddColumn(#""Renamed Columns"", ""Addition"", " & _
        "each [Mobile Sessions] + [Online Sessions], Int64.Type)," & vbCrLf & _
        "    #""Renamed Columns1"" = Table.RenameColumns(#""Inserted Addition""," & _
        "{{""Addition"", ""Total Sessions""}})" & vbCrLf & _
        "in" & vbCrLf & _
        "    #""Renamed Columns1"""
    
    ActiveWorkbook.Queries.Add name:="sessions_csv", formula:=queryFormula
End Sub

' Create alerts query
Sub CreateAlertsQuery(filePath As String)
    Dim queryFormula As String
    
    queryFormula = _
        "let" & vbCrLf & _
        "    Source = Csv.Document(File.Contents(""" & filePath & """)," & _
        "[Delimiter="","", Columns=3, QuoteStyle=QuoteStyle.None])," & vbCrLf & _
        "    #""Promoted Headers"" = Table.PromoteHeaders(Source, [PromoteAllScalars=true])," & vbCrLf & _
        "    #""Changed Type"" = Table.TransformColumnTypes(#""Promoted Headers""," & _
        "{{""Group"", type text}, {""Date"", type date}, " & _
        "{""Number of alerts"", Int64.Type}})," & vbCrLf & _
        "    #""Removed Columns"" = Table.RemoveColumns(#""Changed Type"",{""Group""})" & vbCrLf & _
        "in" & vbCrLf & _
        "    #""Removed Columns"""
    
    ActiveWorkbook.Queries.Add name:="alerts_csv", formula:=queryFormula
End Sub

' Create joined query
Sub CreateJoinedQuery()
    Dim queryFormula As String
    
    queryFormula = _
        "let" & vbCrLf & _
        "    Source = Table.NestedJoin(sessions_csv, {""Date""}, alerts_csv, {""Date""}, " & _
        """alerts_csv"", JoinKind.LeftOuter)," & vbCrLf & _
        "    #""Expanded alerts_csv"" = Table.ExpandTableColumn(Source, ""alerts_csv"", " & _
        "{""Number of alerts""}, {""alerts_csv.Number of alerts""})," & vbCrLf & _
        "    #""Renamed Columns"" = Table.RenameColumns(#""Expanded alerts_csv""," & _
        "{{""alerts_csv.Number of alerts"", ""Total Alerts""}})" & vbCrLf & _
        "in" & vbCrLf & _
        "    #""Renamed Columns"""
    
    ActiveWorkbook.Queries.Add name:="sessions_and_alerts", formula:=queryFormula
End Sub

' Create table in worksheet from query
Private Sub CreateQueryTable(queryName As String)
    ActiveWorkbook.Worksheets.Add
    
    With ActiveSheet.ListObjects.Add(SourceType:=0, Source:= _
        "OLEDB;Provider=Microsoft.Mashup.OleDb.1;Data Source=$Workbook$;" & _
        "Location=" & queryName & ";Extended Properties=""""", _
        Destination:=Range("$A$1")).queryTable
        
        .CommandType = xlCmdSql
        .CommandText = Array("SELECT * FROM [" & queryName & "]")
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
        .listObject.displayName = queryName
        .Refresh BackgroundQuery:=False
    End With
End Sub

Private Sub CreateChart()
    Dim shtActive As Worksheet
    Dim cht As Chart
    Dim rngData As Range
    Dim rngDate As Range
    Dim rngMobile As Range
    Dim rngOnline As Range
    Dim rngTotalSessions As Range
    Dim rngTotalAlerts As Range
    Dim FullSeriesCollectionIndex As Integer
    Dim Serie
    Dim lngDataRangeRowCount As Long
    
'    Application.ScreenUpdating = False
    
    ' Set the worksheet
    Set shtActive = ActiveSheet
    
    ' Define the data ranges
    With shtActive
        Set rngData = .Range("A1").CurrentRegion
        lngDataRangeRowCount = rngData.Rows.count
        Set rngDate = Intersect(rngData, .Columns(1)).Offset(1).Resize(lngDataRangeRowCount - 1)
        Set rngMobile = Intersect(rngData, .Columns(2)).Offset(1).Resize(lngDataRangeRowCount - 1)
        Set rngOnline = Intersect(rngData, .Columns(3)).Offset(1).Resize(lngDataRangeRowCount - 1)
        Set rngTotalSessions = Intersect(rngData, .Columns(4)).Offset(1).Resize(lngDataRangeRowCount - 1)
        Set rngTotalAlerts = Intersect(rngData, .Columns(5)).Offset(1).Resize(lngDataRangeRowCount - 1)
        rngData.CreateNames Top:=True, Left:=False, Bottom:=False, Right:=False
    End With
    
    ' Create a new chart
    Set cht = ActiveSheet.Shapes.AddChart2(227, xlLine).Chart
    With cht
    For Each Serie In .SeriesCollection
        Serie.Delete
    Next Serie
'        .ChartType = xlLine
        
        ' Add data series
        .SeriesCollection.NewSeries
        With .SeriesCollection(1)
            .name = "Total sessions"
            .Values = rngTotalSessions
            .XValues = rngDate
            .Format.Line.ForeColor.RGB = RGB(0, 112, 192)  ' Blue line
            .Format.Line.Weight = 2.5
            .MarkerStyle = xlMarkerStyleNone
        End With
        
        .SeriesCollection.NewSeries
        With .SeriesCollection(2)
            .name = "Mobile"
            .Values = rngMobile
            .XValues = rngDate
            .Format.Line.ForeColor.RGB = RGB(192, 128, 255)  ' Light purple line
            .Format.Line.Weight = 2.5
            .MarkerStyle = xlMarkerStyleNone
        End With
        
        .SeriesCollection.NewSeries
        With .SeriesCollection(3)
            .name = "Online"
            .Values = rngOnline
            .XValues = rngDate
            .Format.Line.ForeColor.RGB = RGB(112, 48, 160)  ' Dark purple line
            .Format.Line.Weight = 2.5
            .MarkerStyle = xlMarkerStyleNone
        End With
        
        .SeriesCollection.NewSeries
        With .SeriesCollection(4)
            .name = "Total Alerts"
            .Values = rngTotalAlerts
            .XValues = rngDate
            .Format.Line.ForeColor.RGB = RGB(255, 192, 0)  ' Gold/yellow line
            .Format.Line.Weight = 2.5
            .MarkerStyle = xlMarkerStyleNone
        End With
        
        ' Format the plot area
        With .PlotArea
            .Format.Fill.Visible = msoFalse
            .Border.LineStyle = xlLineStyleNone
        End With
        
        ' Format the legend
        .HasLegend = True
        With .Legend
            .Position = xlLegendPositionBottom
            .Format.Fill.Visible = msoFalse
            .Border.LineStyle = xlLineStyleNone
            .Font.Size = 14
            .Font.name = "Segoe UI"
            .Font.Color = RGB(96, 96, 96)
        End With
        
        ' Format the axes
        With .Axes(xlCategory)
            .HasMajorGridlines = False
            .HasMinorGridlines = False
            .TickLabelPosition = xlTickLabelPositionLow
            .TickLabels.Font.Size = 14
            .TickLabels.Font.name = "Segoe UI"
            .TickLabels.Font.Color = RGB(96, 96, 96)
            .TickLabels.Orientation = xlTickLabelOrientationUpward
            .Format.Line.ForeColor.RGB = RGB(216, 216, 216)  ' Light gray
        End With
        
        With .Axes(xlValue)
            .HasMajorGridlines = True
            .HasMinorGridlines = False
            .MajorGridlines.Format.Line.ForeColor.RGB = RGB(216, 216, 216)  ' Light gray
            .TickLabels.Font.Size = 14
            .TickLabels.Font.name = "Segoe UI"
            .TickLabels.Font.Color = RGB(96, 96, 96)
            .TickLabels.NumberFormat = "0,,\M"  ' Format with M suffix
            .Format.Line.Visible = msoFalse
            .MinimumScale = 0
        End With
        
        ' Format the chart title (remove it as it's not in the screenshot)
        .HasTitle = False
    
        For FullSeriesCollectionIndex = 1 To 4
            .FullSeriesCollection(FullSeriesCollectionIndex).ApplyDataLabels
            Select Case FullSeriesCollectionIndex
            Case 1, 3:
                With .FullSeriesCollection(FullSeriesCollectionIndex).DataLabels
                    .NumberFormat = "#.0,,""M"""
                    .Position = xlLabelPositionAbove
                    .Font.Size = 14
                End With
            Case 2:
                With .FullSeriesCollection(FullSeriesCollectionIndex).DataLabels
                    .NumberFormat = "#.0,,""M"""
                    .Position = xlLabelPositionBelow
                    .Font.Size = 14
                End With
            Case 4:
                With .FullSeriesCollection(FullSeriesCollectionIndex).DataLabels
                    .NumberFormat = "#.0,""K"""
                    .Position = xlLabelPositionAbove
                    .Font.Size = 14
                End With
            End Select
        Next FullSeriesCollectionIndex
        
        With .Axes(xlCategory)
            .TickLabels.NumberFormat = "mm/yyyy"
            .TickLabels.Orientation = 45
        End With

    End With
    Application.ScreenUpdating = True
End Sub

Private Sub CreateSecurityDashboardPPT()
    ' Declare variables
    Dim pptApp As PowerPoint.Application
    Dim pptPres As PowerPoint.Presentation
    Dim pptSlide As PowerPoint.Slide
    Dim totalSessions As Double
    Dim totalAlerts As Double
    Dim alertRate As Double
    Dim sessionData As Range
    Dim alertData As Range
    Dim shpExcelChart As PowerPoint.Shape
    
    ' Initialize PowerPoint
    On Error Resume Next
    Set pptApp = GetObject(, "PowerPoint.Application")
    If pptApp Is Nothing Then
        Set pptApp = CreateObject("PowerPoint.Application")
    End If
    On Error GoTo 0
    
    pptApp.Visible = True
    
    ' Create a new presentation
    Set pptPres = pptApp.Presentations.Add
    
    ' Add a slide
    Set pptSlide = pptPres.Slides.Add(1, 11) ' 11 = ppLayoutTitleOnly
    
    ' Set the slide title
    With pptSlide.Shapes.Title
        .Left = (pptPres.PageSetup.slideWidth - .Width) / 2
        .Height = 30
        With .TextFrame.TextRange
            .Text = "Volume of sessions and alert rate"
            .ParagraphFormat.Alignment = ppAlignCenter
        End With
    End With
    
    ' Assume chart is in active sheet
    Set sessionData = ActiveSheet.Range("Total_Sessions") ' Total Sessions column
    Set alertData = ActiveSheet.Range("Total_Alerts") ' Total Alerts column
    
    ' Calculate metrics
    totalSessions = Application.WorksheetFunction.Sum(sessionData)
    totalAlerts = Application.WorksheetFunction.Sum(alertData)
    
    If totalSessions > 0 Then
        alertRate = totalAlerts / totalSessions
    Else
        alertRate = 0
    End If
    
        ActiveSheet.ChartObjects(1).Copy
        pptSlide.Shapes.Paste
    
    Set shpExcelChart = pptSlide.Shapes(pptSlide.Shapes.count)
    
    ' Add metrics at bottom of slide
    ' Format numbers for display
    Dim formattedTotal As String
    Dim formattedAlerts As String
    Dim formattedRate As String
    
    ' Format the numbers according to the scale
    If totalSessions >= 1000000 Then
        formattedTotal = Format(totalSessions / 1000000, "0.0") & "M"
    ElseIf totalSessions >= 1000 Then
        formattedTotal = Format(totalSessions / 1000, "0.0") & "K"
    Else
        formattedTotal = Format(totalSessions, "0")
    End If
    
    If totalAlerts >= 1000000 Then
        formattedAlerts = Format(totalAlerts / 1000000, "0.0") & "M"
    ElseIf totalAlerts >= 1000 Then
        formattedAlerts = Format(totalAlerts / 1000, "0.0") & "K"
    Else
        formattedAlerts = Format(totalAlerts, "0")
    End If
    
    ' Format the alert rate as percentage
    formattedRate = Format(alertRate, "0.00%")
    
    ' Add the formatted metrics to the slide
    ' First metric: Total Sessions
    With pptSlide.Shapes.AddTextbox(msoTextOrientationHorizontal, 130, 450, 150, 100)
        With .TextFrame.TextRange
            .Text = formattedTotal
            .Font.Size = 40
            .Font.Bold = True
            .Font.Color.RGB = RGB(0, 120, 200) ' Blue
            .ParagraphFormat.Alignment = ppAlignCenter
        End With
    End With
    
    With pptSlide.Shapes.AddTextbox(msoTextOrientationHorizontal, 130, 500, 150, 50)
        With .TextFrame.TextRange
            .Text = "Analyzed"
            .Font.Size = 24
            .Font.Bold = False
            .Font.Color.RGB = RGB(0, 120, 200) ' Blue
            .ParagraphFormat.Alignment = ppAlignCenter
        End With
    End With
    
    ' Second metric: Total Alerts
    With pptSlide.Shapes.AddTextbox(msoTextOrientationHorizontal, 400, 450, 150, 100)
        With .TextFrame.TextRange
            .Text = formattedAlerts
            .Font.Size = 40
            .Font.Bold = True
            .Font.Color.RGB = RGB(255, 200, 0) ' Yellow
            .ParagraphFormat.Alignment = ppAlignCenter
        End With
    End With
    
    With pptSlide.Shapes.AddTextbox(msoTextOrientationHorizontal, 400, 500, 150, 50)
        With .TextFrame.TextRange
            .Text = "Alerted"
            .Font.Size = 24
            .Font.Bold = False
            .Font.Color.RGB = RGB(255, 200, 0) ' Yellow
            .ParagraphFormat.Alignment = ppAlignCenter
        End With
    End With
    
    ' Third metric: Alert Rate
    With pptSlide.Shapes.AddTextbox(msoTextOrientationHorizontal, 650, 450, 150, 100)
        With .TextFrame.TextRange
            .Text = formattedRate
            .Font.Size = 40
            .Font.Bold = True
            .Font.Color.RGB = RGB(0, 180, 80) ' Green
            .ParagraphFormat.Alignment = ppAlignCenter
        End With
    End With
    
    With pptSlide.Shapes.AddTextbox(msoTextOrientationHorizontal, 650, 500, 150, 50)
        With .TextFrame.TextRange
            .Text = "Alert Rate"
            .Font.Size = 24
            .Font.Bold = False
            .Font.Color.RGB = RGB(0, 180, 80) ' Green
            .ParagraphFormat.Alignment = ppAlignCenter
        End With
    End With
    
    Call ResizeChartFromCenter(pptPres, pptSlide)
    pptApp.Activate
    
    ' Clean up
    Set pptSlide = Nothing
    Set pptPres = Nothing
    Set pptApp = Nothing
    
End Sub

Private Function FormatNumber(Value As Double) As String
    ' Format numbers with appropriate suffixes (K, M, B)
    If Value >= 1000000000 Then
        FormatNumber = Format(Value / 1000000000, "0.0") & "B"
    ElseIf Value >= 1000000 Then
        FormatNumber = Format(Value / 1000000, "0.0") & "M"
    ElseIf Value >= 1000 Then
        FormatNumber = Format(Value / 1000, "0.0") & "K"
    Else
        FormatNumber = Format(Value, "0")
    End If
End Function

Private Sub ResizeChartFromCenter(pptPres As PowerPoint.Presentation, ppdSlide As PowerPoint.Slide)
    Dim oChartObj As PowerPoint.Shape
    Dim oTextbox As Object
    Dim oTextboxes As New Collection
    Dim maxTop As Single
    Dim minBottom As Single
    Dim maxLeft As Single
    Dim minRight As Single
    Dim newWidth As Single
    Dim newHeight As Single
    Dim centerX As Single
    Dim centerY As Single
    Dim slideHeight As Single
    Dim slideWidth As Single
    
    ' Get slide dimensions manually
    slideHeight = pptPres.PageSetup.slideHeight
    slideWidth = pptPres.PageSetup.slideWidth
    Set oChartObj = ppdSlide.Shapes(2)
    
    ' Find boundaries defined by textboxes
    ' Set initial values to slide dimensions
    maxTop = 0
    minBottom = slideHeight
    maxLeft = 0
    minRight = slideWidth
    
    ' Check each textbox for boundaries
    Dim shTop As Single
    Dim shBottom As Single
    Dim shLeft As Single
    Dim shRight As Single
    
    For Each oTextbox In oTextboxes
        shTop = oTextbox.Top
        shBottom = oTextbox.Top + oTextbox.Height
        shLeft = oTextbox.Left
        shRight = oTextbox.Left + oTextbox.Width
        
        ' Update top boundary (max Y value for top objects)
        If shBottom > maxTop Then
            maxTop = shBottom
        End If
        
        ' Update bottom boundary (min Y value for bottom objects)
        If shTop < minBottom Then
            minBottom = shTop
        End If
        
        ' Update left boundary (max X value for left objects)
        If shRight > maxLeft Then
            maxLeft = shRight
        End If
        
        ' Update right boundary (min X value for right objects)
        If shLeft < minRight Then
            minRight = shLeft
        End If
    Next oTextbox
    
    ' Add some padding (in points)
    Const padding As Single = 70
    maxTop = maxTop + padding
    minBottom = minBottom - padding
    maxLeft = maxLeft + padding
    minRight = minRight - padding
    
    ' Calculate available space
    newWidth = minRight - maxLeft
    newHeight = minBottom - maxTop
    
    ' Calculate current center of the chart
    centerX = oChartObj.Left + (oChartObj.Width / 2)
    centerY = oChartObj.Top + (oChartObj.Height / 2)
    
    ' Calculate new dimensions while maintaining aspect ratio
    Dim aspectRatio As Single
    aspectRatio = oChartObj.Width / oChartObj.Height
    
    If newWidth / newHeight > aspectRatio Then
        ' Height is the limiting factor
        newWidth = newHeight * aspectRatio
    Else
        ' Width is the limiting factor
        newHeight = newWidth / aspectRatio
    End If
    
    ' Resize chart from center
    With oChartObj
        .Left = centerX - (newWidth / 2)
        .Top = centerY - (newHeight / 2)
        .Width = newWidth
        .Height = newHeight
    End With
End Sub

