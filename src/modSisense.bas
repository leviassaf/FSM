Attribute VB_Name = "modSisense"
Option Explicit

Sub DetectionRate()
    'Request by Michael Goldberg to transform Sisense export into a customer facing formatted report
    Dim rngData As Range
    
    Set rngData = Range("A1").CurrentRegion
    rngData.EntireColumn.AutoFit
    
    Columns("A:A").HorizontalAlignment = xlLeft
    
    Columns("B:D").HorizontalAlignment = xlCenter
'    Columns("D:D").style = "Percent"
    With Columns("D:D")
        .Value2 = .Value2
        .style = "Percent"
    End With

    Call format_data_range(rngData, 4)
    
End Sub

Public Function format_data_range(data_range As Range, ColumnIndexToFormat As Long)
    Dim rngMoDistributionData As Range
    Dim rngHeaderCells As Range
    
    With data_range
        .Font.ThemeColor = xlThemeColorDark1
        .Interior.ThemeColor = xlThemeColorLight1
    
        .Borders(xlEdgeLeft).ThemeColor = 1
        .Borders(xlEdgeTop).ThemeColor = 1
        .Borders(xlEdgeBottom).ThemeColor = 1
        .Borders(xlEdgeRight).ThemeColor = 1
        .Borders(xlInsideVertical).ThemeColor = 1
        .Borders(xlInsideHorizontal).ThemeColor = 1
    
        Set rngMoDistributionData = Application.Intersect(data_range, Columns(ColumnIndexToFormat))
        Set rngMoDistributionData = rngMoDistributionData.Offset(1).Resize(rngMoDistributionData.Rows.count - 1, 1)
        With rngMoDistributionData
            .FormatConditions.AddDatabar
            With .FormatConditions(1)
                .MinPoint.Modify newtype:=xlConditionValueAutomaticMin
                .MaxPoint.Modify newtype:=xlConditionValueAutomaticMax
            End With
            .FormatConditions(1).BarFillType = xlDataBarFillSolid
        End With
    End With

    Set rngHeaderCells = Application.Intersect(data_range, Rows(1))
    With rngHeaderCells
        With .Font
            .Color = vbRed
            .Size = .Size + 4
        End With
    End With
    data_range.EntireColumn.AutoFit
End Function

Sub RiskReasonDistribution()
    Dim rngData As Range
    Dim rngMoDistributionData As Range
    Dim rngHeaderCells As Range
    
    Columns("B:B").style = "Percent"
    Columns("B:B").HorizontalAlignment = xlCenter
    Set rngData = Range("A1").CurrentRegion
    
    Call format_data_range(rngData, 2)
End Sub
