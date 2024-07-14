Attribute VB_Name = "modFreeRedshiftQuery"
Option Explicit

Sub FreeRedshiftQuery()
Attribute FreeRedshiftQuery.VB_ProcData.VB_Invoke_Func = "R\n14"
    Dim Qry As QueryTable
    Dim wbkReport As Workbook
    Dim shtNew As Worksheet
    Dim strSqlQuery As String
'    Dim strSessionId As String
    Dim msgBoxResult As VbMsgBoxResult
'    Dim strErrorDescription As String
    Dim Sht As Worksheet
    Dim strQueryName As String
    Dim strDSN As String
    
    msgBoxResult = vbYes
    Do While msgBoxResult = vbYes
        If Not isLANCableconnected Then
        'If Not True Then
            msgBoxResult = MsgBox( _
                "You must be connected via a LAN cable from the IBM Office." & vbNewLine & _
                "Would you like to retry?", vbYesNo, "Session ID Analysis" _
            )
            If msgBoxResult = vbNo Then Exit Sub
        Else
            Exit Do
        End If
    Loop
    
    strDSN = "Redshift_EU"
    strQueryName = ActiveCell.Offset(0, -1).Value2
    strSqlQuery = ActiveCell.Value

'    If IsSQLStatementValid(strSqlQuery, strErrorDescription) Then
'ExecuteQuery:
        Application.ScreenUpdating = False

        Set wbkReport = Workbooks.Add(xlWBATWorksheet)
        Set shtNew = wbkReport.ActiveSheet
        Set Qry = CreateQueryTable(shtNew, strDSN)
        With Qry
            .CommandText = strSqlQuery
            .AdjustColumnWidth = True
            .Refresh BackgroundQuery:=False
        End With
        Call FormatDateColumns(shtNew)
        shtNew.Name = Left(strQueryName, 31)
        
        Call SplitWorksheetsByColumnValues("business", ActiveSheet)

        For Each Sht In Worksheets
            Call CreateChart(Sht)
        Next Sht

        Application.ScreenUpdating = True
'    Else
'        If InStr(strErrorDescription, "cancelled on user's request") <> 0 Then
'            GoTo ExecuteQuery
'        Else
'            MsgBox strErrorDescription
'        End If
'    End If
End Sub

Function ReadTextFile(filePath As String) As String
    Dim fileNum As Integer
    Dim fileContent As String
    Dim fileLine As String
    
    ' Open the text file for reading
    fileNum = FreeFile
    Open filePath For Input As fileNum
    
    ' Read the content of the file
    Do Until EOF(fileNum)
        Line Input #fileNum, fileLine
        fileContent = fileContent & fileLine & vbCrLf
    Loop
    
    ' Close the file
    Close fileNum
    
    ' Remove the trailing newline (if any)
    If Right(fileContent, 2) = vbCrLf Then
        fileContent = Left(fileContent, Len(fileContent) - 2)
    End If
    
    ' Return the file content as a string
    ReadTextFile = fileContent
End Function

Public Sub CreateChart(Sht As Worksheet)
    Dim lngColIndexSessions As Long
    Dim rngDataNoHeaders As Range
    Dim Pnt As Point
    Dim lngPoint As Long
    Dim chartObj As ChartObject
    Dim MyShape As Shape
    Dim lngColumnTitleColIndex As Long
    Dim lngDateColIndex As Long
    Dim lngBaselineColIndex As Long
    Dim rngSourceData As Range
    Dim Serie As Series
    Dim dblNumberSessions As Double
    
    Set MyShape = Sht.Shapes.AddChart2(201, xlColumnClustered)
    
    Set rngSourceData = Application.Intersect(Sht.Columns("E:H"), Sht.UsedRange)

    With MyShape
        .LockAspectRatio = msoTrue
        .ScaleWidth 1.3, msoFalse
    End With
    
    With MyShape.Chart
        '.PlotVisibleOnly = False
        Set chartObj = .Parent
        .SetSourceData Source:=rngSourceData
        .PlotBy = xlColumns
        
        .ChartArea.Font.Size = 12
        
        For Each Serie In .FullSeriesCollection
            .ApplyDataLabels
        Next Serie
        
        With .FullSeriesCollection(1)
            .ChartType = xlColumnStacked
            .AxisGroup = 1
        End With
        With .FullSeriesCollection(2)
            .ChartType = xlColumnStacked
            .AxisGroup = 1
        End With
        With .FullSeriesCollection(3)
            .ChartType = xlLine
            .AxisGroup = 2
        End With
        .SetElement (msoElementChartTitleNone)
        
        With .Axes(xlValue, xlSecondary)
            .MinimumScale = 0
            .MaximumScale = 1.2
            .TickLabels.NumberFormat = "0%"
        End With
    
        With .FullSeriesCollection(3).DataLabels
            .NumberFormat = "0%"
        End With
        
        With .FullSeriesCollection(2).DataLabels.Format.TextFrame2.TextRange.Font
            With .Fill
                .Visible = msoTrue
                .ForeColor.ObjectThemeColor = msoThemeColorBackground1
                .ForeColor.TintAndShade = 0
                .ForeColor.Brightness = 0
                .Transparency = 0
                .Solid
            End With
            .Bold = msoTrue
            .Size = 12
        End With
    
        With .FullSeriesCollection(1).DataLabels.Format.TextFrame2.TextRange.Font
            With .Fill
                .Visible = msoTrue
                .ForeColor.ObjectThemeColor = msoThemeColorBackground1
                .ForeColor.TintAndShade = 0
                .ForeColor.Brightness = 0
                .Transparency = 0
                .Solid
            End With
            .Bold = msoTrue
            .Size = 12
        End With
    
        With .FullSeriesCollection(3).DataLabels.Format.TextFrame2.TextRange.Font
            .Bold = msoTrue
            .Size = 12
        End With
        
        With .Legend.Format.TextFrame2.TextRange.Font
            .Size = 12
        End With
    
        .ChartGroups(1).GapWidth = 100
    
        With .Axes(xlCategory)
            .TickLabels.NumberFormat = "[$-fr-FR]mmm-yy;@"
            .BaseUnit = xlMonths
'            .Format.TextFrame2.TextRange.Font.Size = 12
        End With
    End With
    
    With chartObj
        .Left = 10 + (Sht.Shapes.count - 1) * (Sht.Shapes(1).Width + 10)
        .Top = .Top + 100
    End With
End Sub

