Attribute VB_Name = "modSethAutobot"
Option Explicit

Sub LivePolicyListing()
    Dim rngData As Range
    Dim shtRawData As Worksheet
    
    Set shtRawData = ActiveSheet
    With ActiveWindow
        .SplitColumn = 0
        .SplitRow = 1
        .FreezePanes = True
    End With
    
    With shtRawData
        Set rngData = .UsedRange
        With rngData
            .AutoFilter
            .EntireColumn.AutoFit
            .AutoFilter Field:=5, Criteria1:="<>0", Operator:=xlAnd, Criteria2:="<>-1"
        End With
        
        With .AutoFilter.Sort
            .SortFields.Clear
            .SortFields.Add2 Key:=Application.Intersect(shtRawData.Columns(5), rngData), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
            .SortFields.Add2 Key:=Application.Intersect(shtRawData.Columns(4), rngData), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
            .Header = xlYes
            .MatchCase = False
            .Orientation = xlTopToBottom
            .SortMethod = xlPinYin
            .Apply
        End With
        rngData.AutoFilter Field:=1, Criteria1:="=L1_policies", Operator:=xlOr, Criteria2:="=L2_policies"
    End With
End Sub



